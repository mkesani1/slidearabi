"""
SlideArabi v2 — Classification-First RTL Transformer

Replaces the v1 SlideContentTransformer (22+ sequential numbered fixes)
with a deterministic 4-phase pipeline:

    Phase 1: Classify — classify every shape on the slide ONCE
    Phase 2: Panel swap — if split-panel detected, swap panel zones as units
    Phase 3: Per-shape dispatch — each ShapeRole maps to ONE handler
    Phase 4: Post-processing — text-only fixups, NO position changes

No numbered fixes. Every transform is driven by a pre-computed ShapeRole.
"""

from __future__ import annotations

import logging
from copy import deepcopy
from typing import Dict, List, Optional, Set

from lxml import etree

from slidearabi.utils import (
    A_NS, P_NS, R_NS,
    mirror_x,
    bounds_check_emu,
    clamp_emu,
    has_arabic,
    compute_script_ratio,
    ensure_pPr,
    get_placeholder_info,
    get_placeholder_info_from_xml,
)

from slidearabi_v2.shape_classifier import (
    ShapeRole,
    ShapeClassifier,
    ShapeClassification,
    SlideClassificationResult,
    SplitPanelInfo,
    _ROLE_ACTIONS,
    _BLEED_THRESHOLD_EMU,
    _MAX_BLEED_EMU,
    _POSITION_TOLERANCE_EMU,
    _FOOTER_PH_TYPES,
    _DIRECTIONAL_PRESETS,
)
from slidearabi_v2.engine_router import EngineRouter
from slidearabi_v2.v1_compat_dispatcher import V1CompatDispatcher

logger = logging.getLogger(__name__)


# ═══════════════════════════════════════════════════════════════════════════════
# SlideArabi namespace for idempotency markers
SLIDEARABI_NS = 'https://slidearabi.ai/ns/transform'

# Constants
# ═══════════════════════════════════════════════════════════════════════════════

# Directional preset swap map (type-swap or '_flipH' sentinel)
_DIRECTIONAL_SWAP: Dict[str, str] = {
    'rightArrow': 'leftArrow',
    'leftArrow': 'rightArrow',
    'rightArrowCallout': 'leftArrowCallout',
    'leftArrowCallout': 'rightArrowCallout',
    'curvedRightArrow': 'curvedLeftArrow',
    'curvedLeftArrow': 'curvedRightArrow',
    'leftRightArrow': 'leftRightArrow',   # symmetric — no change
    'upDownArrow': 'upDownArrow',          # symmetric — no change
    'chevron': '_flipH',
    'homePlate': '_flipH',
    'notchedRightArrow': '_flipH',
    'stripedRightArrow': '_flipH',
    'bentArrow': '_flipH',
    'bentUpArrow': '_flipH',
    'circularArrow': '_flipH',
    'pentagon': '_flipH',
}


# ═══════════════════════════════════════════════════════════════════════════════
# TransformReport (same as v1 for pipeline compatibility)
# ═══════════════════════════════════════════════════════════════════════════════

class TransformReport:
    """Accumulates transform statistics and diagnostics."""

    def __init__(self, phase: str = 'slide'):
        self.phase = phase
        self.counts: Dict[str, int] = {}
        self.warnings: List[str] = []
        self.errors: List[str] = []

    def add(self, key: str, count: int = 1) -> None:
        self.counts[key] = self.counts.get(key, 0) + count

    def warn(self, msg: str) -> None:
        self.warnings.append(msg)

    def error(self, msg: str) -> None:
        self.errors.append(msg)

    @property
    def total_changes(self) -> int:
        """V1-compatible property: sum of all counters."""
        return sum(self.counts.values())

    @property
    def changes_by_type(self) -> Dict[str, int]:
        """V1-compatible property: alias for counts."""
        return self.counts

    def merge(self, other: 'TransformReport') -> None:
        for k, v in other.counts.items():
            self.counts[k] = self.counts.get(k, 0) + v
        self.warnings.extend(other.warnings)
        self.errors.extend(other.errors)


# ═══════════════════════════════════════════════════════════════════════════════
# SlideContentTransformerV2
# ═══════════════════════════════════════════════════════════════════════════════

class SlideContentTransformerV2:
    """
    v2 RTL transformer using intent-based dispatch.

    Each ShapeRole maps to exactly ONE transform handler via the
    _ROLE_ACTIONS table in shape_classifier.py. No numbered fixes —
    every transform is driven by a pre-computed classification.

    Lifecycle:
        1. Pipeline instantiates with presentation + classifier + translations
        2. transform_all_slides() iterates slides
        3. Per slide: classify → panel swap → per-shape dispatch → post-process
    """

    def __init__(
        self,
        presentation,
        template_registry=None,
        layout_classifications: Optional[Dict] = None,
        translations: Optional[Dict[str, str]] = None,
        v1_transformer=None,
        engine_router: Optional[EngineRouter] = None,
    ):
        self.prs = presentation
        self.template_registry = template_registry
        self.layout_classifications = layout_classifications or {}
        self.translations = translations or {}
        self._slide_width = int(presentation.slide_width)
        self._slide_height = int(presentation.slide_height)

        # Shape classifier instance
        self.classifier = ShapeClassifier(
            slide_width=self._slide_width,
            slide_height=self._slide_height,
        )

        # Engine router and v1 fallback dispatcher
        self._router = engine_router or EngineRouter()
        self._v1_dispatcher: Optional[V1CompatDispatcher] = None
        if v1_transformer is not None:
            self._v1_dispatcher = V1CompatDispatcher(v1_transformer)

        # Pre-build lowercase translation index for O(1) case-insensitive lookups
        # Fix: keep longer Arabic value on duplicate keys (was: first-writer-wins)
        self._translations_lower: Dict[str, str] = {}
        for key, val in self.translations.items():
            lower_key = key.strip().lower()
            existing = self._translations_lower.get(lower_key)
            if existing is None or len(val) > len(existing):
                self._translations_lower[lower_key] = val

        # Normalized whitespace index — collapses internal whitespace for robust matching
        self._translations_normalized: Dict[str, str] = {}
        for key, val in self.translations.items():
            norm_key = ' '.join(key.split()).strip().lower()
            existing = self._translations_normalized.get(norm_key)
            if existing is None or len(val) > len(existing):
                self._translations_normalized[norm_key] = val

        # Handler dispatch table: position_action → handler method
        self._position_handlers = {
            'mirror': self._handle_mirror,
            'keep': self._handle_keep,
            'swap': self._handle_swap,
            'inherit': self._handle_inherit,
            'reposition': self._handle_reposition,
        }

        # Direction action handlers
        self._direction_handlers = {
            'remove_flip': self._handle_remove_flip,
            'toggle_flipH': self._handle_toggle_flipH,
            'swap_preset': self._handle_swap_preset,
            'none': self._handle_direction_noop,
        }

        # Text action handlers
        self._text_handlers = {
            'translate_rtl': self._handle_translate_rtl,
            'rtl_only': self._handle_rtl_only,
            'none': self._handle_text_noop,
        }

    # ─────────────────────────────────────────────────────────────────────
    # Public entry point
    # ─────────────────────────────────────────────────────────────────────

    def transform_all_slides(self) -> TransformReport:
        """Transform all content slides. Returns a TransformReport."""
        report = TransformReport(phase='slide_v2')
        for slide_idx, slide in enumerate(self.prs.slides):
            slide_number = slide_idx + 1
            try:
                count = self._transform_slide(slide, slide_number)
                report.add('slide_transformed', count)
            except Exception as exc:
                report.error(f'slide[{slide_number}]: {exc}')
                logger.exception('v2 transform failed on slide %d', slide_number)
        return report

    # ─────────────────────────────────────────────────────────────────────
    # Phase orchestration
    # ─────────────────────────────────────────────────────────────────────

    def _transform_slide(self, slide, slide_number: int) -> int:
        """
        Transform a single slide using the 4-phase pipeline:

        Phase 1: Classify all shapes (one-shot, immutable)
        Phase 2: Panel swap (if split-panel detected)
        Phase 3: Per-shape dispatch (position → direction → text)
        Phase 4: Post-processing (text-only, no position changes)
        """
        changes = 0

        # Resolve layout type
        layout = slide.slide_layout
        layout_type = layout._element.get('type', 'cust')
        if slide_number in self.layout_classifications:
            layout_type = self.layout_classifications[slide_number]

        # ── Phase 1: Classify ──────────────────────────────────────────
        result = self.classifier.classify_slide(
            slide, slide_number, layout_type=layout_type,
        )

        # Collect shapes (groups as single units) for transform dispatch
        all_shapes = self._collect_all_shapes(slide.shapes)
        logger.debug(
            'Slide %d: classified %d shapes (%s)',
            slide_number, len(result.classifications), layout_type,
        )

        # ── Phase 2: Panel swap ────────────────────────────────────────
        swapped_ids: Set[int] = set()
        if result.has_split_panel:
            changes += self._execute_panel_swap(
                all_shapes, result, swapped_ids
            )
            # DISABLED: collision resolver creates worse regressions than the
            # collisions it fixes. v1 production runs without collision
            # resolution and produces acceptable results.  All 4 frontier
            # models agree: disable for now, re-enable with confidence gating
            # once deck-level regression tests pass.
            #
            # changes += self._resolve_panel_swap_collisions(
            #     all_shapes, result, swapped_ids
            # )
            # changes += self._resolve_text_text_overlaps(
            #     all_shapes, result, swapped_ids
            # )

        # ── Phase 3: Per-shape dispatch ────────────────────────────────
        for shape in all_shapes:
            try:
                cls = result.get(shape)

                # ── Engine routing: v1 fallback for roles not yet in v2 ──
                if not self._router.use_v2(cls.role) and self._v1_dispatcher is not None:
                    changes += self._v1_dispatcher.dispatch(
                        shape, cls.role, layout_type, slide_number, all_shapes,
                    )
                    continue

                # Panel-swapped shapes skip position handling
                if shape.shape_id in swapped_ids:
                    # Direction transforms (remove_flip etc.) are always safe
                    changes += self._dispatch_direction(shape, cls)

                    # Text transforms: translate first, then set RTL alignment.
                    # _set_rtl_alignment already guards rtl='1' with has_arabic()
                    # per-paragraph, so it's safe to call unconditionally.
                    # (Round 12 fix: not calling it left untranslated shapes
                    # with no RTL alignment at all in panel-swap zones)
                    if getattr(shape, 'has_text_frame', False):
                        changes += self._apply_translation(shape)
                        changes += self._set_rtl_alignment(shape, cls)
                    # Handle group children text
                    if cls.role == ShapeRole.GROUP:
                        changes += self._process_group_children(shape, cls)
                    elif cls.role == ShapeRole.COMPLEX_GRAPHIC:
                        changes += self._process_complex_graphic_children(shape, cls)
                    if cls.role == ShapeRole.TABLE:
                        changes += self._handle_table(shape)
                    continue

                # Position dispatch
                changes += self._dispatch_position(shape, cls, layout)

                # Direction dispatch
                changes += self._dispatch_direction(shape, cls)

                # Text dispatch
                changes += self._dispatch_text(shape, cls)

                # Role-specific extras
                if cls.role == ShapeRole.TABLE:
                    changes += self._handle_table(shape)
                elif cls.role == ShapeRole.CHART:
                    changes += self._handle_chart(shape)
                elif cls.role == ShapeRole.GROUP:
                    changes += self._process_group_children(shape, cls)
                elif cls.role == ShapeRole.COMPLEX_GRAPHIC:
                    # P1-2: Complex graphics — translate children text ONLY.
                    # No direction reversal on children (preserve infographic layout).
                    changes += self._process_complex_graphic_children(shape, cls)
                elif cls.role == ShapeRole.CONNECTOR:
                    changes += self._handle_connector_arrowheads(shape)

            except Exception as exc:
                logger.warning(
                    'Slide %d shape "%s": %s',
                    slide_number, getattr(shape, 'name', '?'), exc,
                )

        # ── Phase 3.5: Table overlap resolution ─────────────────────
        changes += self._resolve_table_overlaps(all_shapes, result)

        # ── Phase 3.6: Table overlay icon correction ───────────────
        # Icons/shapes overlaid on tables need table-local mirroring
        # (not slide-center mirroring) to stay aligned with reversed columns
        changes += self._fix_table_overlay_icons(all_shapes, result)

        # ── Phase 4: Post-processing (text-only) ──────────────────────
        changes += self._post_process(all_shapes, result, slide_number)

        # v1 post-processing for roles handled by v1 dispatcher
        if self._v1_dispatcher is not None:
            changes += self._v1_dispatcher.dispatch_post_processing(
                all_shapes, slide_number,
            )

        # ── Phase 5: Master/layout decorative shape mirroring ────────
        # Thin decorative lines and bars from master/layout slides
        # are not in the slide's spTree so they don't get mirrored.
        # We detect them and add mirrored copies + occluders on the slide.
        changes += self._mirror_master_decorative_shapes(slide, slide_number)

        # Telemetry
        self._log_translation_coverage(all_shapes, slide_number)

        return changes

    # ─────────────────────────────────────────────────────────────────────
    # Shape collection
    # ─────────────────────────────────────────────────────────────────────

    @staticmethod
    def _collect_all_shapes(shapes) -> List:
        """Collect top-level shapes. Groups included as single units."""
        return list(shapes)

    @staticmethod
    def _collect_text_shapes_from_group(group_shape) -> List:
        """Recursively collect text-bearing shapes inside a group."""
        result = []
        try:
            for child in group_shape.shapes:
                if getattr(child, 'has_text_frame', False) and child.has_text_frame:
                    result.append(child)
                if hasattr(child, 'shapes'):
                    result.extend(
                        SlideContentTransformerV2._collect_text_shapes_from_group(child)
                    )
        except Exception:
            pass
        return result

    # ─────────────────────────────────────────────────────────────────────
    # Phase 2: Panel swap
    # ─────────────────────────────────────────────────────────────────────

    def _execute_panel_swap(
        self,
        shapes: List,
        result: SlideClassificationResult,
        swapped_ids: Set[int],
    ) -> int:
        """
        Swap left and right panel zones using per-shape mirror.

        Each panel shape is independently mirrored:
            new_left = slide_width - left - width

        This produces the same result as v1's _pre_mirror_split_panel_swap()
        and correctly reverses the internal order within each panel.
        """
        panel = result.context.split_panel
        if panel is None:
            return 0

        changes = 0

        for shape in shapes:
            sid = shape.shape_id
            cls = result.get(shape)

            if cls.role in (ShapeRole.PANEL_LEFT, ShapeRole.PANEL_RIGHT):
                try:
                    # Skip placeholders with no local xfrm: they inherit
                    # position from the layout, which MasterLayoutTransformer
                    # has already RTL-mirrored.  Swapping here would double-
                    # mirror and move the shape back to its LTR position.
                    sp_el = shape._element
                    _sp_pr = sp_el.find(f'{{{P_NS}}}spPr')
                    if _sp_pr is None:
                        _sp_pr = sp_el.find(f'{{{A_NS}}}spPr')
                    _has_local_xfrm = False
                    if _sp_pr is not None:
                        _has_local_xfrm = _sp_pr.find(f'{{{A_NS}}}xfrm') is not None

                    _is_ph = False
                    try:
                        _is_ph = shape.is_placeholder
                    except (ValueError, AttributeError):
                        pass

                    if not _has_local_xfrm and _is_ph:
                        # Inherit from mirrored layout — no swap needed
                        swapped_ids.add(sid)  # Mark as handled
                        continue

                    old_left = int(shape.left)
                    w = int(shape.width)
                    new_left = self._slide_width - old_left - w
                    new_left = max(0, min(new_left, self._slide_width - w))
                    shape.left = new_left
                    swapped_ids.add(sid)
                    changes += 1
                except Exception as exc:
                    logger.debug('Panel swap mirror failed: %s', exc)

        if changes:
            logger.debug(
                'Panel swap: mirrored %d shapes', changes,
            )
        return changes

    def _resolve_panel_swap_collisions(
        self,
        shapes: List,
        result: 'SlideClassificationResult',
        swapped_ids: Set[int],
    ) -> int:
        """
        Post-swap collision detection: check if any text shape now overlaps
        with an image shape after panel swap. If so, push the text shape
        to the clear side of the image.

        Only operates on shapes that were actually swapped (in swapped_ids).
        Deterministic: pushes text to the nearest clear edge.
        """
        if not swapped_ids:
            return 0

        changes = 0
        MIN_CLEARANCE_EMU = 50_000  # ~0.055" minimum gap

        # Collect post-swap geometry for swapped shapes
        image_rects = []   # (left, top, right, bottom, shape_id)
        text_shapes = []   # [shape, left, top, right, bottom]  (mutable for tracking)

        for shape in shapes:
            sid = shape.shape_id
            if sid not in swapped_ids:
                continue
            try:
                left = int(shape.left or 0)
                top = int(shape.top or 0)
                width = int(shape.width or 0)
                height = int(shape.height or 0)
            except (TypeError, ValueError):
                continue

            right_edge = left + width
            bottom_edge = top + height

            sp_el = shape._element
            tag = sp_el.tag.split('}')[-1] if '}' in sp_el.tag else sp_el.tag
            is_image = (
                tag == 'pic'
                or sp_el.find(f'.//{{{A_NS}}}blipFill') is not None
            )

            if is_image:
                image_rects.append((left, top, right_edge, bottom_edge, sid))
            elif getattr(shape, 'has_text_frame', False):
                text_shapes.append([shape, left, top, right_edge, bottom_edge])

        for ts in text_shapes:
            shape, t_left, t_top, t_right, t_bottom = ts
            text_width = t_right - t_left

            for i_left, i_top, i_right, i_bottom, i_sid in image_rects:
                h_overlap = t_left < i_right and t_right > i_left
                v_overlap = t_top < i_bottom and t_bottom > i_top
                if not (h_overlap and v_overlap):
                    continue

                # Collision detected — push text to clear side of image
                space_right = self._slide_width - i_right
                space_left = i_left

                if space_right >= text_width + MIN_CLEARANCE_EMU:
                    new_left = i_right + MIN_CLEARANCE_EMU
                elif space_left >= text_width + MIN_CLEARANCE_EMU:
                    new_left = i_left - text_width - MIN_CLEARANCE_EMU
                elif space_right >= space_left:
                    new_left = i_right + MIN_CLEARANCE_EMU
                else:
                    new_left = max(0, i_left - text_width - MIN_CLEARANCE_EMU)

                new_left = max(0, min(int(new_left), self._slide_width - text_width))

                if new_left != t_left:
                    shape.left = new_left
                    ts[1] = new_left
                    ts[3] = new_left + text_width
                    changes += 1

                break  # One resolution per text shape

        return changes

    def _resolve_text_text_overlaps(
        self,
        shapes: List,
        result: 'SlideClassificationResult',
        swapped_ids: Set[int],
    ) -> int:
        """
        After panel swap, detect and resolve overlaps between text shapes
        that ended up sharing the same zone.

        For each pair of overlapping text shapes (both in swapped_ids):
        - Keep the larger/upper shape in place
        - Push the smaller/lower shape downward by at least 50,000 EMU
        - If pushing down would exceed 95% of slide height, push sideways
        """
        if not swapped_ids:
            return 0

        changes = 0
        MIN_TEXT_CLEARANCE_EMU = 50_000

        text_rects = []  # (shape, left, top, width, height, area)
        for shape in shapes:
            if shape.shape_id not in swapped_ids:
                continue
            if not getattr(shape, 'has_text_frame', False):
                continue
            try:
                left = int(shape.left or 0)
                top = int(shape.top or 0)
                width = int(shape.width or 0)
                height = int(shape.height or 0)
            except (TypeError, ValueError):
                continue
            text_rects.append([shape, left, top, width, height, width * height])

        if len(text_rects) < 2:
            return 0

        # Sort by area descending, then top ascending (larger/upper first)
        text_rects.sort(key=lambda r: (-r[5], r[2]))

        for i in range(len(text_rects)):
            s_a, l_a, t_a, w_a, h_a, _ = text_rects[i]
            r_a = l_a + w_a
            b_a = t_a + h_a

            for j in range(i + 1, len(text_rects)):
                s_b, l_b, t_b, w_b, h_b, _ = text_rects[j]
                r_b = l_b + w_b
                b_b = t_b + h_b

                h_overlap = l_a < r_b and r_a > l_b
                v_overlap = t_a < b_b and b_a > t_b
                if not (h_overlap and v_overlap):
                    continue

                # Push shape B down to clear shape A
                new_top_b = b_a + MIN_TEXT_CLEARANCE_EMU
                max_bottom = int(self._slide_height * 0.95)

                if new_top_b + h_b > max_bottom:
                    # Can't push down — try sideways
                    new_left_b = r_a + MIN_TEXT_CLEARANCE_EMU
                    if new_left_b + w_b <= self._slide_width:
                        s_b.left = int(new_left_b)
                        text_rects[j][1] = int(new_left_b)
                        changes += 1
                    else:
                        new_left_b = l_a - w_b - MIN_TEXT_CLEARANCE_EMU
                        if new_left_b >= 0:
                            s_b.left = int(new_left_b)
                            text_rects[j][1] = int(new_left_b)
                            changes += 1
                    continue

                s_b.top = int(new_top_b)
                text_rects[j][2] = int(new_top_b)
                changes += 1

        return changes

    # ─────────────────────────────────────────────────────────────────────
    # Phase 3: Position dispatch
    # ─────────────────────────────────────────────────────────────────────

    def _dispatch_position(
        self, shape, cls: ShapeClassification, layout
    ) -> int:
        # ── Hard rule: Pie/doughnut-only charts keep position ────────
        # Rotationally symmetric charts don't need position mirroring;
        # only translate labels within the chart XML.
        if cls.role == ShapeRole.CHART and cls.position_action == 'mirror':
            if self._is_chart_pie_doughnut_only(shape):
                logger.debug('Pie/doughnut chart: keeping position')
                return 0

        # ── Smart table mirroring ────────────────────────────────────
        # Round 4 fix: Mirror ALL tables regardless of width.
        # Previous logic skipped tables >50% slide width, assuming
        # rtl='1' handles internal column reversal. But for empty-style
        # tables, rtl='1' doesn't reorder columns, and we now do
        # physical reversal. Tables like Slide 18 (51.4% width) were
        # NOT mirrored while adjacent text shapes WERE, causing overlap.
        # Now: tables >90% width keep position (truly full-width),
        # all others get mirrored.
        if cls.role == ShapeRole.TABLE and cls.position_action == 'mirror':
            try:
                table_w = int(shape.width or 0)
                if table_w > self._slide_width * 0.90:
                    logger.debug(
                        'Full-width table (%d > 90%% of %d): keeping position',
                        table_w, self._slide_width,
                    )
                    return 0  # Keep position for truly full-width tables
            except (TypeError, ValueError):
                pass

        handler = self._position_handlers.get(cls.position_action)
        if handler is None:
            return 0
        result = handler(shape, cls, layout)

        # ── Table bounds clamping after mirroring ────────────────────
        if cls.role == ShapeRole.TABLE and result > 0:
            self._clamp_table_bounds(shape)

        return result

    def _handle_mirror(self, shape, cls: ShapeClassification, layout) -> int:
        """Standard mirror: new_x = slide_width - (x + w)."""
        try:
            left = shape.left
            width = shape.width
            if left is None or width is None:
                return 0

            new_left = mirror_x(int(left), int(width), self._slide_width)

            # Bleed clamping: allow intentional bleeds up to MAX_BLEED_EMU
            new_left = max(new_left, -_MAX_BLEED_EMU)
            if width is not None:
                max_right = self._slide_width + _MAX_BLEED_EMU
                if new_left + int(width) > max_right:
                    new_left = max_right - int(width)

            if not bounds_check_emu(new_left, self._slide_width):
                return 0

            # P1-3: Bypass tolerance for LOGO — logos must always mirror
            # even when near-centered (e.g., header/footer logos).
            if cls.role != ShapeRole.LOGO:
                if abs(new_left - int(left)) < _POSITION_TOLERANCE_EMU:
                    return 0  # Negligible change — likely centred

            shape.left = new_left
            return 1
        except Exception as exc:
            logger.debug('_handle_mirror: %s', exc)
            return 0

    def _handle_keep(self, shape, cls: ShapeClassification, layout) -> int:
        """Do not move — backgrounds, decoratives, overlays."""
        return 0

    def _is_chart_pie_doughnut_only(self, shape) -> bool:
        """Check if shape is a pie/doughnut-only chart (no axis charts)."""
        try:
            if not (getattr(shape, 'has_chart', False) and shape.has_chart):
                return False
            chart_elem = shape.chart._part._element
            return self._is_pie_or_doughnut_only(chart_elem, self._C_NS)
        except Exception:
            return False

    def _clamp_table_bounds(self, shape) -> None:
        """Ensure table shape stays within slide boundaries after mirroring."""
        try:
            left = int(shape.left or 0)
            width = int(shape.width or 0)
            right_edge = left + width

            # Clamp right edge
            if right_edge > self._slide_width:
                new_left = self._slide_width - width
                if new_left < 0:
                    new_left = 0
                shape.left = new_left
                logger.debug(
                    'Table bounds clamp: left %d → %d (was overflowing by %d)',
                    left, new_left, right_edge - self._slide_width,
                )

            # Clamp left edge
            if int(shape.left or 0) < 0:
                shape.left = 0
        except Exception as exc:
            logger.debug('_clamp_table_bounds: %s', exc)

    def _resolve_table_overlaps(
        self, shapes: List, result: 'SlideClassificationResult'
    ) -> int:
        """
        Post-mirror check: if two tables overlap after mirroring,
        nudge them apart to prevent visual collision.
        """
        table_shapes = []
        for shape in shapes:
            cls = result.get(shape)
            if cls.role == ShapeRole.TABLE:
                try:
                    left = int(shape.left or 0)
                    width = int(shape.width or 0)
                    top = int(shape.top or 0)
                    height = int(shape.height or 0)
                    table_shapes.append((shape, left, left + width, top, top + height))
                except (TypeError, ValueError):
                    continue

        if len(table_shapes) < 2:
            return 0

        changes = 0
        # Sort by left position
        table_shapes.sort(key=lambda t: t[1])
        MIN_GAP = 50_000  # ~0.055"

        for i in range(len(table_shapes) - 1):
            s_a, l_a, r_a, t_a, b_a = table_shapes[i]
            s_b, l_b, r_b, t_b, b_b = table_shapes[i + 1]

            # Check vertical overlap first
            v_overlap = t_a < b_b and t_b < b_a
            if not v_overlap:
                continue

            # Check horizontal overlap
            if r_a > l_b:  # Overlap detected
                overlap = r_a - l_b + MIN_GAP
                w_b = r_b - l_b
                new_left_b = l_b + overlap
                if new_left_b + w_b > self._slide_width:
                    # Can't push right — push A leftward
                    w_a = r_a - l_a
                    new_left_a = l_a - overlap
                    if new_left_a >= 0:
                        s_a.left = new_left_a
                        table_shapes[i] = (s_a, new_left_a, new_left_a + w_a, t_a, b_a)
                        changes += 1
                else:
                    s_b.left = int(new_left_b)
                    table_shapes[i + 1] = (s_b, int(new_left_b), int(new_left_b) + w_b, t_b, b_b)
                    changes += 1

        if changes:
            logger.debug('Resolved %d table overlaps', changes)
        return changes

    def _fix_table_overlay_icons(
        self, shapes: List, result: 'SlideClassificationResult'
    ) -> int:
        """
        Phase 3.6: Fix shapes overlaid on PHYSICALLY-REVERSED tables.

        After table columns are physically reversed in XML, overlay icons
        that were mirrored around slide center are now misaligned with their
        target columns. Re-mirror them around the TABLE's center instead.

        Round 3 fix (5-model consensus): Only applies to tables that were
        physically reversed (have physRtlCols='1' marker). For logical-RTL-only
        tables (tblPr rtl='1' without physical reversal), the XML column
        positions haven't changed, so overlay icons stay where they are.

        Detection: small shape whose center falls inside a table's bbox.
        Transform: new_left = table_left + (table_width - (old_left - table_left) - shape_width)
        """
        # Collect ONLY physically-reversed table bounding boxes
        marker_attr = f'{{{SLIDEARABI_NS}}}physRtlCols'
        tables = []
        for shape in shapes:
            cls = result.get(shape)
            if cls.role == ShapeRole.TABLE:
                try:
                    # Only target tables that were physically reversed
                    table = shape.table
                    tbl_elem = table._tbl
                    tbl_pr = tbl_elem.find(f'{{{A_NS}}}tblPr')
                    if tbl_pr is None or tbl_pr.get(marker_attr) != '1':
                        continue  # Logical-RTL only — no icon remap needed

                    t_left = int(shape.left or 0)
                    t_top = int(shape.top or 0)
                    t_width = int(shape.width or 0)
                    t_height = int(shape.height or 0)
                    tables.append((shape, t_left, t_top, t_width, t_height))
                except (TypeError, ValueError, AttributeError):
                    continue

        if not tables:
            return 0

        changes = 0
        # Max overlay icon area as fraction of table area
        MAX_OVERLAY_AREA_FRACTION = 0.05  # 5% of table area

        for shape in shapes:
            cls = result.get(shape)
            # Only consider non-table, non-placeholder shapes
            if cls.role in (ShapeRole.TABLE, ShapeRole.PLACEHOLDER, ShapeRole.BACKGROUND,
                            ShapeRole.DECORATIVE, ShapeRole.BLEED):
                continue

            try:
                s_left = int(shape.left or 0)
                s_top = int(shape.top or 0)
                s_width = int(shape.width or 0)
                s_height = int(shape.height or 0)
            except (TypeError, ValueError):
                continue

            s_area = s_width * s_height
            s_center_x = s_left + s_width // 2
            s_center_y = s_top + s_height // 2

            for tbl_shape, t_left, t_top, t_width, t_height in tables:
                t_area = t_width * t_height
                if t_area <= 0:
                    continue

                # Check: shape center inside table bbox, and shape is small
                margin = int(self._slide_width * 0.02)  # 2% slide width tolerance
                if (t_left - margin <= s_center_x <= t_left + t_width + margin
                        and t_top - margin <= s_center_y <= t_top + t_height + margin
                        and s_area <= t_area * MAX_OVERLAY_AREA_FRACTION):

                    # This shape is overlaid on a physically-reversed table —
                    # re-mirror around table center
                    new_left = t_left + (t_width - (s_left - t_left) - s_width)

                    # Bounds check
                    if new_left < t_left - margin:
                        new_left = t_left
                    if new_left + s_width > t_left + t_width + margin:
                        new_left = t_left + t_width - s_width

                    if new_left != s_left:
                        shape.left = int(new_left)
                        changes += 1
                        logger.debug(
                            'Table overlay icon "%s": left %d → %d (physical-table remap)',
                            getattr(shape, 'name', '?'), s_left, new_left,
                        )
                    break  # Only match one table per shape

        if changes:
            logger.debug('Fixed %d table overlay icons', changes)
        return changes

    def _handle_swap(self, shape, cls: ShapeClassification, layout) -> int:
        """Panel swap — handled in Phase 2. This is a fallback no-op."""
        # Panel swap is already done in _execute_panel_swap.
        # If a shape is tagged SWAP but wasn't handled, that's a classification
        # edge case — log it but don't crash.
        logger.debug(
            'SWAP action on shape "%s" not handled by Phase 2',
            getattr(shape, 'name', '?'),
        )
        return 0

    def _handle_inherit(self, shape, cls: ShapeClassification, layout) -> int:
        """
        Placeholder: inherit position (offset) from layout while preserving
        the slide-author's width/height overrides.

        Previous approach removed the entire xfrm, which discarded the slide
        author's intentional size customization. Now we replace only the
        offset (x, y) from the layout placeholder, keeping extents (cx, cy).

        If the shape's size differs from layout by >30%, fall back to
        explicit mirroring (the layout position is unreliable).
        """
        try:
            sp_el = shape._element

            ph_info = get_placeholder_info_from_xml(sp_el)
            if ph_info is None:
                return 0
            _, ph_idx = ph_info

            # Find matching layout placeholder
            layout_ph = None
            for lph in layout.placeholders:
                if lph.placeholder_format.idx == ph_idx:
                    layout_ph = lph
                    break
            if layout_ph is None:
                return 0

            # Find local xfrm
            sp_pr = sp_el.find(f'{{{P_NS}}}spPr')
            if sp_pr is None:
                sp_pr = sp_el.find(f'{{{A_NS}}}spPr')
            if sp_pr is None:
                return 0

            xfrm = sp_pr.find(f'{{{A_NS}}}xfrm')
            if xfrm is None:
                # No local xfrm — shape inherits position from the layout.
                # After MasterLayoutTransformer, the layout is already RTL-mirrored,
                # so the inherited position is correct. Do nothing.
                return 0

            # Check for size divergence: if the slide shape's size differs
            # from the layout placeholder's size by >30%, the shape has been
            # deliberately resized and removing xfrm would snap it to wrong size.
            try:
                shape_w = int(shape.width or 0)
                shape_h = int(shape.height or 0)
                layout_w = int(layout_ph.width or 0)
                layout_h = int(layout_ph.height or 0)
                if layout_w > 0 and layout_h > 0:
                    w_ratio = abs(shape_w - layout_w) / layout_w
                    h_ratio = abs(shape_h - layout_h) / layout_h
                    if w_ratio > 0.30 or h_ratio > 0.30:
                        # Size-divergent: mirror explicitly instead of inheriting
                        return self._handle_mirror(shape, cls, layout)
            except (TypeError, ValueError, ZeroDivisionError):
                pass

            # Get layout placeholder's offset for position inheritance
            layout_xfrm = layout_ph._element.find(f'.//{{{A_NS}}}xfrm')
            if layout_xfrm is None:
                # No layout xfrm to inherit from — fall back to mirror
                return self._handle_mirror(shape, cls, layout)

            layout_off = layout_xfrm.find(f'{{{A_NS}}}off')
            if layout_off is None:
                return self._handle_mirror(shape, cls, layout)

            # Check if layout offset is essentially the same as current offset.
            # If so, the layout hasn't been RTL-transformed and inheriting
            # would keep the shape in its LTR position. Fall back to mirror.
            off = xfrm.find(f'{{{A_NS}}}off')
            if off is not None:
                try:
                    current_x = int(off.get('x', '0'))
                    layout_x = int(layout_off.get('x', '0'))
                    # If positions differ by less than 1% of slide width,
                    # layout isn't providing useful RTL repositioning
                    threshold = self._slide_width * 0.01
                    if abs(current_x - layout_x) < threshold:
                        return self._handle_mirror(shape, cls, layout)
                except (ValueError, TypeError):
                    pass

            # Replace offset only — inherit position from layout
            # but preserve the slide author's width/height overrides
            if off is not None:
                off.set('x', layout_off.get('x', '0'))
                off.set('y', layout_off.get('y', '0'))
            else:
                # No offset element — create one from layout
                new_off = etree.SubElement(xfrm, f'{{{A_NS}}}off')
                new_off.set('x', layout_off.get('x', '0'))
                new_off.set('y', layout_off.get('y', '0'))

            # ext (cx, cy) is left untouched — preserving slide-level size
            return 1

        except Exception as exc:
            logger.debug('_handle_inherit: %s', exc)
            return 0

    def _handle_reposition(
        self, shape, cls: ShapeClassification, layout
    ) -> int:
        """
        Badge reposition: move from right zone to equivalent left zone.
        Mirror the X coordinate using standard formula.
        """
        try:
            left = shape.left
            width = shape.width
            if left is None or width is None:
                return 0

            new_left = mirror_x(int(left), int(width), self._slide_width)
            if not bounds_check_emu(new_left, self._slide_width):
                return 0
            if abs(new_left - int(left)) < _POSITION_TOLERANCE_EMU:
                return 0

            shape.left = new_left
            return 1
        except Exception as exc:
            logger.debug('_handle_reposition: %s', exc)
            return 0

    # ─────────────────────────────────────────────────────────────────────
    # Phase 3: Direction dispatch
    # ─────────────────────────────────────────────────────────────────────

    def _dispatch_direction(self, shape, cls: ShapeClassification) -> int:
        handler = self._direction_handlers.get(cls.direction_action)
        if handler is None:
            return 0
        return handler(shape, cls)

    # Shapes that encode DATA (not direction) — never flip inside groups
    _DATA_SHAPE_PRESETS = frozenset({
        'arc', 'pie', 'chord', 'donut', 'blockArc',
        'wedgeRoundRectCallout', 'wedgeEllipseCallout',
    })

    def _handle_remove_flip(self, shape, cls: ShapeClassification) -> int:
        """Remove flipH/flipV from shape xfrm — prevents content mirroring.

        For GROUP shapes, only target the group-level grpSpPr/xfrm,
        NOT children's xfrms. Children may have intentional flips that
        encode data (e.g., arc fill direction in infographic charts).
        """
        changes = 0
        try:
            sp_el = shape._element

            if cls.role == ShapeRole.GROUP:
                # Only remove flip on the group-level xfrm
                xfrm = sp_el.find(f'{{{P_NS}}}grpSpPr/{{{A_NS}}}xfrm')
                if xfrm is None:
                    xfrm = sp_el.find(f'{{{A_NS}}}grpSpPr/{{{A_NS}}}xfrm')
                if xfrm is not None:
                    if xfrm.get('flipH'):
                        del xfrm.attrib['flipH']
                        changes += 1
                    if xfrm.get('flipV'):
                        del xfrm.attrib['flipV']
                        changes += 1
                return min(changes, 1)

            # Non-group: iterate all xfrms (original behavior)
            for xfrm in sp_el.iter(f'{{{A_NS}}}xfrm'):
                if xfrm.get('flipH'):
                    del xfrm.attrib['flipH']
                    changes += 1
                if xfrm.get('flipV'):
                    del xfrm.attrib['flipV']
                    changes += 1
        except Exception:
            pass
        return min(changes, 1)  # Count as single change

    def _handle_toggle_flipH(self, shape, cls: ShapeClassification) -> int:
        """Toggle flipH on connectors — mirroring reverses the line slope."""
        try:
            sp_el = shape._element
            for xfrm in sp_el.iter(f'{{{A_NS}}}xfrm'):
                current = (xfrm.get('flipH', '0') or '0').lower()
                if current in ('1', 'true'):
                    if 'flipH' in xfrm.attrib:
                        del xfrm.attrib['flipH']
                else:
                    xfrm.set('flipH', '1')
                return 1
        except Exception:
            pass
        return 0

    def _handle_swap_preset(self, shape, cls: ShapeClassification) -> int:
        """Reverse directional shapes: swap preset type or toggle flipH."""
        try:
            sp_el = shape._element
            prst_geom = sp_el.find(f'.//{{{A_NS}}}prstGeom')
            if prst_geom is None:
                return 0

            prst = prst_geom.get('prst', '')
            action = _DIRECTIONAL_SWAP.get(prst)
            if action is None:
                return 0

            if action == '_flipH':
                xfrm = sp_el.find(f'.//{{{A_NS}}}xfrm')
                if xfrm is not None:
                    current = xfrm.get('flipH', '0')
                    xfrm.set('flipH', '0' if current == '1' else '1')
                    return 1
            elif action == prst:
                return 0  # Symmetric — no change
            else:
                prst_geom.set('prst', action)
                return 1
        except Exception:
            pass
        return 0

    def _handle_direction_noop(self, shape, cls: ShapeClassification) -> int:
        return 0

    # ─────────────────────────────────────────────────────────────────────
    # Phase 3: Text dispatch
    # ─────────────────────────────────────────────────────────────────────

    def _dispatch_text(self, shape, cls: ShapeClassification) -> int:
        handler = self._text_handlers.get(cls.text_action)
        if handler is None:
            return 0
        return handler(shape, cls)

    def _handle_translate_rtl(self, shape, cls: ShapeClassification) -> int:
        """Full treatment: translate text + set RTL alignment.
        
        Always applies alignment mirroring (right-align for RTL layout).
        The rtl='1' BiDi flag is set conditionally inside _set_rtl_alignment:
        only when paragraph text actually contains Arabic, to avoid
        punctuation reordering on English text.
        """
        changes = 0
        if not getattr(shape, 'has_text_frame', False):
            return 0
        changes += self._apply_translation(shape)
        changes += self._set_rtl_alignment(shape, cls)
        return changes

    def _handle_rtl_only(self, shape, cls: ShapeClassification) -> int:
        """Set RTL properties on existing text without translation."""
        if not getattr(shape, 'has_text_frame', False):
            return 0
        return self._set_rtl_alignment(shape, cls)

    def _handle_text_noop(self, shape, cls: ShapeClassification) -> int:
        return 0

    # ─────────────────────────────────────────────────────────────────────
    # Text translation engine (ported from v1)
    # ─────────────────────────────────────────────────────────────────────

    def _fuzzy_lookup_translation(self, text: str) -> Optional[str]:
        """
        Flexible translation lookup with fallbacks:
        1. Exact match on full/stripped text
        2. Case-insensitive exact match
        3. Normalized whitespace match (collapses internal whitespace)
        4. Longest prefix match (>80% of text length, for text >40 chars)
        """
        if not text or not text.strip():
            return None

        stripped = text.strip()

        # 1. Exact match
        result = self.translations.get(text) or self.translations.get(stripped)
        if result:
            return result

        # 2. Case-insensitive match
        result = self._translations_lower.get(stripped.lower())
        if result:
            return result

        # 3. Normalized whitespace match (collapses tabs, double-spaces, etc.)
        normalized = ' '.join(stripped.split()).lower()
        result = self._translations_normalized.get(normalized)
        if result:
            return result

        # 4. Longest prefix match for long text
        if len(stripped) > 40:
            prefix_40 = stripped[:40]
            best_match = None
            best_match_len = 0
            for key, val in self.translations.items():
                if len(key) < 40 or key[:40] != prefix_40:
                    continue
                common_len = min(len(key), len(stripped))
                match_len = 40
                for i in range(40, common_len):
                    if key[i] == stripped[i]:
                        match_len += 1
                    else:
                        break
                min_len = min(len(key), len(stripped))
                if match_len > min_len * 0.80 and match_len > best_match_len:
                    best_match_len = match_len
                    best_match = val
            if best_match:
                return best_match

        return None

    def _apply_translation(self, shape) -> int:
        """
        Replace English text with Arabic translation, preserving run-level
        formatting. Sets RTL direction and lang=ar-SA on translated runs.
        """
        if not self.translations:
            return 0

        changes = 0
        try:
            tf = shape.text_frame
            for para in tf.paragraphs:
                para_text = para.text
                if not para_text or not para_text.strip():
                    continue

                arabic_text = self._fuzzy_lookup_translation(para_text)
                if not arabic_text:
                    # Log untranslated paragraphs for debugging (skip short/trivial text)
                    if len(para_text.strip()) > 10 and not has_arabic(para_text):
                        logger.warning(
                            '_apply_translation: no match for "%s" (%d chars)',
                            para_text.strip()[:80], len(para_text.strip()),
                        )
                    continue

                runs = para.runs
                if runs:
                    runs[0].text = arabic_text
                    for run in runs[1:]:
                        run.text = ''
                else:
                    # Round 4 fix: Remove <a:fld> (field) elements before
                    # appending new run. Fields like slidenum render their
                    # value; adding a run without removing the field causes
                    # doubling (e.g., "15" + "15" = "1515").
                    p_elem = para._p
                    for fld in p_elem.findall(f'{{{A_NS}}}fld'):
                        p_elem.remove(fld)
                    r_elem = etree.SubElement(p_elem, f'{{{A_NS}}}r')
                    t_elem = etree.SubElement(r_elem, f'{{{A_NS}}}t')
                    t_elem.text = arabic_text

                # Set Arabic language on runs
                for r_elem in para._p.findall(f'{{{A_NS}}}r'):
                    rPr = r_elem.find(f'{{{A_NS}}}rPr')
                    if rPr is not None:
                        rPr.set('lang', 'ar-SA')

                changes += 1

        except Exception as exc:
            logger.warning(
                '_apply_translation on "%s": %s',
                getattr(shape, 'name', '?'), exc,
            )
        return changes

    # Footer placeholder types that keep LTR direction
    _FOOTER_RTL_EXEMPT = frozenset({
        'ftr', 'sldNum', 'dt', 'footer', 'slideNumber', 'date_time',
    })

    def _set_rtl_alignment(self, shape, cls: ShapeClassification) -> int:
        """
        Set paragraph-level alignment and conditional RTL direction.

        Two decoupled concerns:
        1. Alignment mirroring (pPr algn) — ALWAYS applied for RTL layout.
           This right-aligns bullets, titles, body text etc.
        2. BiDi direction flag (pPr rtl='1') — ONLY set when the paragraph
           text actually contains Arabic characters. English text with
           rtl='1' causes punctuation reordering (e.g. ".Hello" → "Hello.").
           In production translations produce Arabic → flag is set.
           In test mode with no translations → flag is skipped, but
           alignment is still mirrored.
        """
        changes = 0
        ph_type = cls.placeholder_type

        # Determine if this is a footer-type placeholder (keep LTR)
        is_footer_ph = False
        if ph_type:
            ph_lower = ph_type.lower()
            is_footer_ph = ph_lower in self._FOOTER_RTL_EXEMPT or any(
                f in ph_lower for f in ('ftr', 'footer', 'sldnum')
            )

        try:
            tf = shape.text_frame
            for para in tf.paragraphs:
                text = para.text or ''
                if not text.strip():
                    continue

                pPr = ensure_pPr(para._p)

                # Read original alignment BEFORE any modification
                original_algn = pPr.get('algn')

                # BiDi direction flag: conditional on text content
                if is_footer_ph:
                    pPr.set('rtl', '0')
                elif has_arabic(text):
                    # Text contains Arabic → enable BiDi rendering
                    pPr.set('rtl', '1')
                    # Swap bullet indentation for RTL
                    changes += self._swap_bullet_indentation(para._p)
                else:
                    # English text — explicitly set rtl='0' to prevent
                    # inherited RTL from master/layout causing punctuation
                    # reordering (e.g. "?DO" instead of "DO?")
                    pPr.set('rtl', '0')

                # Alignment mirroring: ALWAYS apply for RTL layout
                alignment = self._compute_paragraph_alignment(
                    text, ph_type, original_algn,
                )
                pPr.set('algn', alignment)
                changes += 1

            # Set rtlCol on bodyPr only when shape contains Arabic text
            # rtlCol='1' causes punctuation reordering on English text
            shape_text = tf.text if tf else ''
            if not is_footer_ph and has_arabic(shape_text):
                body_pr = tf._txBody.find(f'{{{A_NS}}}bodyPr')
                if body_pr is not None:
                    body_pr.set('rtlCol', '1')

        except Exception as exc:
            logger.debug(
                '_set_rtl_alignment on "%s": %s',
                getattr(shape, 'name', '?'), exc,
            )
        return changes

    def _compute_paragraph_alignment(
        self, text: str, ph_type: Optional[str],
        original_algn: Optional[str] = None,
    ) -> str:
        """
        Compute OOXML algn value for a paragraph in RTL context.

        Priority:
        1. Footer placeholders → 'l'
        2. Title placeholders (not subtitle) → 'ctr'
        2.5. Original alignment is 'ctr' AND not Arabic-dominant → 'ctr'
        3. Arabic-dominant → 'r'
        4. All other text in RTL layout → 'r'
        """
        _footer_substrings = (
            'ftr', 'footer', 'sldnum', 'slide_number', 'date_time', 'date',
        )
        if ph_type in _FOOTER_PH_TYPES or (
            ph_type and any(f in ph_type for f in _footer_substrings)
        ):
            return 'l'

        if ph_type and 'subtitle' not in ph_type.lower() and 'title' in ph_type.lower():
            # Mirror the original alignment: l→r, r→l, ctr→ctr
            if original_algn == 'l':
                return 'r'
            elif original_algn == 'r':
                return 'l'
            elif original_algn == 'ctr':
                return 'ctr'
            # No explicit alignment → default to right for RTL
            return 'r'

        # Preserve explicit center alignment for non-Arabic text
        # (e.g., "THANK YOU" centered by design)
        ratios = compute_script_ratio(text)
        if original_algn == 'ctr' and ratios['arabic'] <= 0.30:
            return 'ctr'

        if ratios['arabic'] > 0.70:
            return 'r'
        # In RTL layout, all body text should be right-aligned
        return 'r'

    # ─────────────────────────────────────────────────────────────────────
    # Bullet indentation for RTL
    # ─────────────────────────────────────────────────────────────────────

    def _swap_bullet_indentation(self, para_p) -> int:
        """
        NO-OP: Do NOT swap marL → marR for bullet paragraphs.

        PowerPoint treats marL as a LOGICAL start-side margin when
        pPr rtl='1' is set (undocumented but empirically confirmed by
        all 5 frontier models, Round 11).  marL already acts as the
        right-side indent in RTL context.  Swapping marL→marR causes a
        double-reversal that collapses hanging indents and creates a
        staircase effect toward center.

        The correct approach: keep marL and indent as-is, set rtl='1'
        on the paragraph, and let PowerPoint handle the visual flip.
        """
        return 0

    # ─────────────────────────────────────────────────────────────────────
    # Role-specific handlers
    # ─────────────────────────────────────────────────────────────────────

    @staticmethod
    def _table_needs_physical_reversal(tbl_pr) -> bool:
        """
        Determine if a table requires physical column reversal.

        Round 4 fix — 5-model UNANIMOUS consensus:
        PowerPoint's `tblPr rtl='1'` only visually reorders columns
        when a valid `tableStyleId` GUID is bound. For tables with
        empty or absent `tableStyleId` (common in Google Slides exports),
        `rtl='1'` affects only text direction WITHIN cells, NOT column
        layout. Physical reversal is required for these tables.

        Returns True when `tableStyleId` is absent or empty (needs
        physical reversal because `rtl='1'` won't mirror columns).
        Returns False when `tableStyleId` is a non-empty GUID
        (PowerPoint's style pipeline will honor `rtl='1'`).
        """
        if tbl_pr is None:
            return True  # No tblPr at all → needs physical reversal
        style_id = tbl_pr.get('tableStyleId', '').strip()
        # Empty or absent tableStyleId → physical reversal needed
        # Non-empty GUID → PowerPoint style pipeline handles RTL
        return not bool(style_id)

    def _handle_table(self, shape) -> int:
        """
        Transform table for RTL.

        Round 4 strategy (5-model UNANIMOUS consensus):
        1) For tables WITH empty/absent tableStyleId (Google Slides exports):
           - Physical column reversal (swap gridCol + tc elements)
           - Set tblPr rtl='0' to prevent PowerPoint reinterpretation
           - Per-cell RTL text properties
        2) For tables WITH a valid tableStyleId GUID:
           - Set tblPr rtl='1' (PowerPoint style pipeline handles mirroring)
           - NO physical reversal (would cause double-reversal)

        Round 3/2 failures explained:
        - Round 2: physical reversal + rtl='1' = double-reversal
        - Round 3: rtl='1' alone on empty-style tables = no visual effect
        - Round 4: physical reversal + rtl='0' = deterministic, correct
        """
        changes = 0
        try:
            table = shape.table
            num_cols = len(table.columns)
            if num_cols <= 1:
                return 0

            tbl_elem = table._tbl

            # Ensure table-level tblPr exists
            tbl_pr = tbl_elem.find(f'{{{A_NS}}}tblPr')
            if tbl_pr is None:
                tbl_pr = etree.Element(f'{{{A_NS}}}tblPr')
                tbl_grid = tbl_elem.find(f'{{{A_NS}}}tblGrid')
                if tbl_grid is not None:
                    idx = list(tbl_elem).index(tbl_grid)
                    tbl_elem.insert(idx, tbl_pr)
                else:
                    tbl_elem.insert(0, tbl_pr)

            marker_attr = f'{{{SLIDEARABI_NS}}}physRtlCols'
            needs_physical = self._table_needs_physical_reversal(tbl_pr)

            if needs_physical:
                # Empty/absent tableStyleId → physical reversal + rtl='0'
                if tbl_pr.get(marker_attr) != '1':
                    changes += self._physically_reverse_table_columns(tbl_elem)
                    tbl_pr.set(marker_attr, '1')
                    # Swap firstCol/lastCol toggles when physical order changed
                    first_col = tbl_pr.get('firstCol')
                    last_col = tbl_pr.get('lastCol')
                    if first_col or last_col:
                        tbl_pr.set('firstCol', last_col or '0')
                        tbl_pr.set('lastCol', first_col or '0')
                    changes += 1
                    logger.debug(
                        'Table "%s": physical reversal applied (empty tableStyleId)',
                        getattr(shape, 'name', '?'),
                    )
                # CRITICAL: Set rtl='0' to prevent double-reversal.
                # Physical reversal is the source of truth for column order.
                tbl_pr.set('rtl', '0')
            else:
                # Non-empty tableStyleId GUID → rtl='1' only (no physical)
                if tbl_pr.get('rtl') != '1':
                    tbl_pr.set('rtl', '1')
                    changes += 1
                # Clear stale physical-reversal marker from prior buggy runs
                if tbl_pr.get(marker_attr):
                    try:
                        del tbl_pr.attrib[marker_attr]
                    except (KeyError, ValueError):
                        pass
                logger.debug(
                    'Table "%s": logical RTL only (has tableStyleId GUID)',
                    getattr(shape, 'name', '?'),
                )

            # Translate cell text AFTER physical reversal to avoid
            # DOM invalidation issues (Round 4, Gemini insight)
            if self.translations:
                for row in table.rows:
                    for cell in row.cells:
                        if cell.text_frame:
                            changes += self._translate_cell_text(cell.text_frame)

            # Set RTL properties on cell text (includes alignment fix)
            for row in table.rows:
                for col_idx, cell in enumerate(row.cells):
                    if cell.text_frame:
                        self._set_cell_rtl_properties(
                            cell.text_frame, col_idx, num_cols,
                        )
                        changes += 1

        except Exception as exc:
            logger.warning(
                '_handle_table on "%s": %s',
                getattr(shape, 'name', '?'), exc,
            )
        return changes

    def _physically_reverse_table_columns(self, tbl_elem) -> int:
        """
        Reverse gridCol order and cell order in every row.
        Preserves each cell node (including merge/format attrs) by
        moving XML nodes, not recreating them.
        """
        changes = 0

        # 1) Reverse <a:gridCol> order in <a:tblGrid>
        tbl_grid = tbl_elem.find(f'{{{A_NS}}}tblGrid')
        if tbl_grid is not None:
            grid_cols = list(tbl_grid.findall(f'{{{A_NS}}}gridCol'))
            if len(grid_cols) > 1:
                for gc in grid_cols:
                    tbl_grid.remove(gc)
                for gc in reversed(grid_cols):
                    tbl_grid.append(gc)
                changes += 1

        # 2) Reverse <a:tc> order in each <a:tr>
        for tr in tbl_elem.findall(f'{{{A_NS}}}tr'):
            cells = list(tr.findall(f'{{{A_NS}}}tc'))
            if len(cells) <= 1:
                continue
            for tc in cells:
                tr.remove(tc)
            for tc in reversed(cells):
                tr.append(tc)
            changes += 1

        return changes

    def _translate_cell_text(self, text_frame) -> int:
        """Translate text within a table cell's text frame."""
        changes = 0
        try:
            for para in text_frame.paragraphs:
                para_text = para.text
                if not para_text or not para_text.strip():
                    continue
                arabic_text = self._fuzzy_lookup_translation(para_text)
                if not arabic_text:
                    continue
                runs = para.runs
                if runs:
                    runs[0].text = arabic_text
                    for run in runs[1:]:
                        run.text = ''
                changes += 1
        except Exception:
            pass
        return changes

    def _set_cell_rtl_properties(
        self, text_frame, col_idx: int, num_cols: int
    ) -> None:
        """Set RTL paragraph properties on cell text.

        Round 3 fix (5-model consensus):
        - Arabic paragraphs: set rtl='1' + mirror algn='l' → 'r'
        - Non-Arabic paragraphs: set rtl='0', keep original alignment
        - Preserve center ('ctr'), right ('r'), justified ('just') alignment

        The Round 11 rule ("never override algn") was correct when physical
        column reversal was the mechanism. With logical RTL (tblPr rtl='1'
        only), cell-internal alignment still refers to the cell's own edges,
        so Arabic body text at algn='l' anchors to the wrong edge.
        """
        try:
            has_any_arabic = False

            for para in text_frame.paragraphs:
                text = para.text or ''
                if not text.strip():
                    continue
                pPr = ensure_pPr(para._p)
                para_has_arabic = has_arabic(text)

                if para_has_arabic:
                    has_any_arabic = True
                    pPr.set('rtl', '1')
                    # Conservative alignment normalization:
                    # only mirror explicit left → right for Arabic text.
                    # Keep ctr, just, dist, r unchanged.
                    algn = pPr.get('algn')
                    if algn == 'l' or (algn is None):
                        pPr.set('algn', 'r')
                else:
                    pPr.set('rtl', '0')
                    # Non-Arabic: keep original alignment

            # Set rtlCol only when cell includes Arabic content
            body_pr = text_frame._txBody.find(f'{{{A_NS}}}bodyPr')
            if body_pr is not None:
                if has_any_arabic:
                    body_pr.set('rtlCol', '1')
                else:
                    # Remove stale rtlCol from prior runs
                    if 'rtlCol' in body_pr.attrib:
                        del body_pr.attrib['rtlCol']
        except Exception:
            pass

    # Chart namespace (distinct from A_NS which is drawingml)
    _C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart'

    # Known Google Translate month name errors
    _MONTH_CORRECTIONS = {
        'يمشي': 'مارس',       # "March" (verb walk → month)
        'يمكن': 'مايو',       # "May" (verb can → month)
        'يُولي و': 'يوليو',   # Garbled "July"
        'يمك': 'مايو',        # Truncated "May"
    }

    def _handle_chart(self, shape) -> int:
        """
        Transform a chart for RTL.  Ported from v1's _transform_chart_rtl.

        Operations:
        1. Detect chart types; skip axis manipulation for pie/doughnut-only.
        2. Reverse catAx AND dateAx orientation for RTL reading order.
        3. Mirror valAx position (l↔r only, NOT t↔b).
        4. Reverse valAx orientation ONLY for horizontal bar / scatter axes.
        5. Mirror serAx position (3D charts).
        6. Mirror legend position; remove manual legend layout.
        7. Translate chart category labels and series names.
        8. Set RTL on chart title.
        """
        changes = 0
        try:
            if not (getattr(shape, 'has_chart', False) and shape.has_chart):
                return 0

            chart = shape.chart
            chart_part = chart._part
            chart_elem = chart_part._element
            c_ns = self._C_NS

            # ── Pie/doughnut early return ─────────────────────────────
            # Pie/doughnut charts are rotationally symmetric — no axis
            # reversal, no legend mirroring, no layout removal.
            # Only translate labels and set RTL on chart title.
            # Legend layout removal causes PowerPoint to reflow data
            # labels from their pinned positions (Round 11, 5/5 consensus).
            if self._is_pie_or_doughnut_only(chart_elem, c_ns):
                logger.debug('Pie/doughnut-only chart: translate-only path')
                if self.translations:
                    changes += self._translate_chart_labels(chart_elem, c_ns)
                try:
                    if chart.has_title and chart.chart_title.has_text_frame:
                        for para in chart.chart_title.text_frame.paragraphs:
                            para_text = (para.text or '').strip()
                            if para_text:
                                pPr = ensure_pPr(para._p)
                                if has_arabic(para_text):
                                    pPr.set('rtl', '1')
                                    pPr.set('algn', 'r')
                                else:
                                    pPr.set('rtl', '0')
                                    pPr.set('algn', 'l')
                                changes += 1
                except Exception:
                    pass
                return changes

            # ── Detect chart types present ─────────────────────────────
            axis_chart_types = set()
            for tag in ('barChart', 'bar3DChart', 'lineChart', 'line3DChart',
                        'areaChart', 'area3DChart', 'scatterChart', 'radarChart',
                        'stockChart', 'surfaceChart', 'surface3DChart',
                        'bubbleChart'):
                if chart_elem.find(f'.//{{{c_ns}}}{tag}') is not None:
                    axis_chart_types.add(tag)

            has_axis_charts = bool(axis_chart_types)

            # ── Collect horiz-bar valAx IDs ────────────────────────────
            horiz_bar_val_ax_ids: set = set()
            for bar_tag in ('barChart', 'bar3DChart'):
                for bar_chart in chart_elem.iter(f'{{{c_ns}}}{bar_tag}'):
                    bar_dir = bar_chart.find(f'{{{c_ns}}}barDir')
                    if bar_dir is not None and bar_dir.get('val') == 'bar':
                        for ax_id in bar_chart.iter(f'{{{c_ns}}}axId'):
                            horiz_bar_val_ax_ids.add(ax_id.get('val'))

            # ── Collect scatter horizontal valAx IDs ───────────────────
            scatter_horiz_val_ax_ids: set = set()
            for scatter in chart_elem.iter(f'{{{c_ns}}}scatterChart'):
                scatter_ax_ids = [
                    ax_id.get('val')
                    for ax_id in scatter.iter(f'{{{c_ns}}}axId')
                ]
                for val_ax in chart_elem.iter(f'{{{c_ns}}}valAx'):
                    ax_id_elem = val_ax.find(f'{{{c_ns}}}axId')
                    if ax_id_elem is None:
                        continue
                    ax_id_val = ax_id_elem.get('val')
                    if ax_id_val not in scatter_ax_ids:
                        continue
                    ax_pos = val_ax.find(f'{{{c_ns}}}axPos')
                    if ax_pos is not None and ax_pos.get('val') in ('b', 't'):
                        scatter_horiz_val_ax_ids.add(ax_id_val)

            # ── Step 1: Reverse catAx AND dateAx (skip for pie/donut) ─
            if has_axis_charts:
                # Collect catAx IDs that belong to horizontal bar charts
                horiz_bar_cat_ax_ids: set = set()
                for bar_tag in ('barChart', 'bar3DChart'):
                    for bar_chart in chart_elem.iter(f'{{{c_ns}}}{bar_tag}'):
                        bar_dir = bar_chart.find(f'{{{c_ns}}}barDir')
                        if bar_dir is not None and bar_dir.get('val') == 'bar':
                            for ax_id in bar_chart.iter(f'{{{c_ns}}}axId'):
                                horiz_bar_cat_ax_ids.add(ax_id.get('val'))

                for ax_tag in ('catAx', 'dateAx'):
                    for cat_ax in chart_elem.iter(f'{{{c_ns}}}{ax_tag}'):
                        # Identify whether this catAx belongs to a horizontal bar chart
                        ax_id_elem = cat_ax.find(f'{{{c_ns}}}axId')
                        ax_id_val = ax_id_elem.get('val') if ax_id_elem is not None else None
                        is_horiz_bar_cat = ax_id_val in horiz_bar_cat_ax_ids

                        # Remove crossesAt (mutually exclusive with crosses)
                        crosses_at = cat_ax.find(f'{{{c_ns}}}crossesAt')
                        if crosses_at is not None:
                            cat_ax.remove(crosses_at)

                        crosses = cat_ax.find(f'{{{c_ns}}}crosses')

                        if is_horiz_bar_cat:
                            # Horizontal bar chart: cross at max puts the catAx
                            # on the right side (correct for RTL)
                            if crosses is None:
                                crosses = etree.SubElement(
                                    cat_ax, f'{{{c_ns}}}crosses',
                                )
                            crosses.set('val', 'max')
                        else:
                            # Line/area/column: remove crosses override so
                            # PowerPoint uses default (autoZero = axis at bottom)
                            # Setting crosses='max' here would push the X-axis
                            # to the TOP of the chart — a confirmed bug.
                            if crosses is not None:
                                cat_ax.remove(crosses)

                        scaling = cat_ax.find(f'{{{c_ns}}}scaling')
                        if scaling is None:
                            scaling = etree.SubElement(
                                cat_ax, f'{{{c_ns}}}scaling',
                            )
                            cat_ax.insert(0, scaling)
                        orientation = scaling.find(f'{{{c_ns}}}orientation')
                        if orientation is None:
                            orientation = etree.SubElement(
                                scaling, f'{{{c_ns}}}orientation',
                            )
                        orientation.set('val', 'maxMin')
                        changes += 1

            # ── Step 2: Handle value axes ──────────────────────────────
            if has_axis_charts:
                for val_ax in chart_elem.iter(f'{{{c_ns}}}valAx'):
                    ax_id_elem = val_ax.find(f'{{{c_ns}}}axId')
                    ax_id_val = (
                        ax_id_elem.get('val') if ax_id_elem is not None
                        else None
                    )

                    # 2a: axPos mirroring — ONLY l↔r, NOT t↔b
                    ax_pos = val_ax.find(f'{{{c_ns}}}axPos')
                    if ax_pos is not None:
                        pos = ax_pos.get('val', 'l')
                        if pos == 'l':
                            ax_pos.set('val', 'r')
                            changes += 1
                        elif pos == 'r':
                            ax_pos.set('val', 'l')
                            changes += 1

                    # 2b: orientation reversal — ONLY horiz bar/scatter
                    is_horiz_bar = ax_id_val in horiz_bar_val_ax_ids
                    is_scatter_horiz = ax_id_val in scatter_horiz_val_ax_ids

                    if is_horiz_bar or is_scatter_horiz:
                        scaling = val_ax.find(f'{{{c_ns}}}scaling')
                        if scaling is None:
                            scaling = etree.SubElement(
                                val_ax, f'{{{c_ns}}}scaling',
                            )
                            val_ax.insert(0, scaling)
                        orientation = scaling.find(f'{{{c_ns}}}orientation')
                        if orientation is None:
                            orientation = etree.SubElement(
                                scaling, f'{{{c_ns}}}orientation',
                            )
                        orientation.set('val', 'maxMin')

                        crosses_at = val_ax.find(f'{{{c_ns}}}crossesAt')
                        if crosses_at is not None:
                            val_ax.remove(crosses_at)
                        changes += 1

            # ── Step 3: Mirror serAx position (3D charts) ─────────────
            for ser_ax in chart_elem.iter(f'{{{c_ns}}}serAx'):
                ax_pos = ser_ax.find(f'{{{c_ns}}}axPos')
                if ax_pos is not None:
                    pos = ax_pos.get('val', 'b')
                    if pos == 'l':
                        ax_pos.set('val', 'r')
                        changes += 1
                    elif pos == 'r':
                        ax_pos.set('val', 'l')
                        changes += 1

            # ── Step 4: Mirror legend position ─────────────────────────
            legend = chart_elem.find(f'.//{{{c_ns}}}legend')
            if legend is not None:
                leg_pos = legend.find(f'{{{c_ns}}}legendPos')
                if leg_pos is None:
                    leg_pos = etree.SubElement(
                        legend, f'{{{c_ns}}}legendPos',
                    )
                    leg_pos.set('val', 'l')
                    changes += 1
                else:
                    pos_val = leg_pos.get('val', 'r')
                    mirror_map = {
                        'r': 'l', 'l': 'r', 'tr': 'tl', 'tl': 'tr',
                    }
                    new_pos = mirror_map.get(pos_val, pos_val)
                    leg_pos.set('val', new_pos)
                    if new_pos != pos_val:
                        changes += 1

                # Remove manual legend layout for PowerPoint reflow
                legend_layout = legend.find(f'{{{c_ns}}}layout')
                if legend_layout is not None:
                    legend.remove(legend_layout)
                    changes += 1

            # ── Step 5: Translate chart labels ─────────────────────────
            if self.translations:
                changes += self._translate_chart_labels(chart_elem, c_ns)

            # ── Step 6: Set RTL on chart title ─────────────────────────
            try:
                if chart.has_title and chart.chart_title.has_text_frame:
                    for para in chart.chart_title.text_frame.paragraphs:
                        para_text = (para.text or '').strip()
                        if para_text:
                            pPr = ensure_pPr(para._p)
                            if has_arabic(para_text):
                                pPr.set('rtl', '1')
                                pPr.set('algn', 'r')
                            else:
                                pPr.set('rtl', '0')
                                pPr.set('algn', 'l')
                            changes += 1
            except Exception:
                pass

        except Exception as exc:
            logger.warning(
                '_handle_chart on "%s": %s',
                getattr(shape, 'name', '?'), exc,
            )
        return changes

    def _is_pie_or_doughnut_only(self, chart_elem, c_ns: str) -> bool:
        """Check if chart contains ONLY pie/doughnut types (no axis charts)."""
        pie_types = (
            'pieChart', 'pie3DChart', 'doughnutChart', 'ofPieChart',
        )
        axis_types = (
            'barChart', 'bar3DChart', 'lineChart', 'line3DChart',
            'areaChart', 'area3DChart', 'scatterChart', 'radarChart',
            'stockChart', 'surfaceChart', 'surface3DChart', 'bubbleChart',
        )
        has_pie = any(
            chart_elem.find(f'.//{{{c_ns}}}{t}') is not None
            for t in pie_types
        )
        has_axis = any(
            chart_elem.find(f'.//{{{c_ns}}}{t}') is not None
            for t in axis_types
        )
        return has_pie and not has_axis

    def _pie_has_inside_labels(self, chart_elem, c_ns: str) -> bool:
        """Detect if pie/doughnut chart uses in-segment data labels."""
        inside_positions = {'ctr', 'inEnd', 'inBase', 'bestFit'}

        for dLbls in chart_elem.iter(f'{{{c_ns}}}dLbls'):
            pos = dLbls.find(f'{{{c_ns}}}dLblPos')
            if pos is not None and pos.get('val') in inside_positions:
                return True
            show_cat = dLbls.find(f'{{{c_ns}}}showCatName')
            if show_cat is not None and show_cat.get('val') == '1':
                if pos is None or pos.get('val') != 'outEnd':
                    return True

        for dLbl in chart_elem.iter(f'{{{c_ns}}}dLbl'):
            pos = dLbl.find(f'{{{c_ns}}}dLblPos')
            if pos is not None and pos.get('val') in inside_positions:
                return True

        return False

    def _translate_chart_labels(self, chart_elem, c_ns: str) -> int:
        """
        Translate chart category labels and series names.

        For pie/doughnut with in-segment labels, SKIP translating category
        cache values (Arabic is 20-40% wider and overflows constrained
        pie segments).  Legend and series names are still translated.
        """
        changes = 0

        is_pie = self._is_pie_or_doughnut_only(chart_elem, c_ns)
        skip_cat = is_pie and self._pie_has_inside_labels(chart_elem, c_ns)

        if skip_cat:
            logger.debug(
                'Chart label translation: skipping categories for '
                'pie/doughnut with in-segment labels',
            )

        # Translate category axis labels
        if not skip_cat:
            for cat in chart_elem.iter(f'{{{c_ns}}}cat'):
                for pt in cat.iter(f'{{{c_ns}}}pt'):
                    v = pt.find(f'{{{c_ns}}}v')
                    if v is not None and v.text:
                        original = v.text.strip()
                        arabic = (
                            self.translations.get(original)
                            or self.translations.get(original.title())
                            or self.translations.get(original.upper())
                            or self.translations.get(original.lower())
                        )
                        if arabic:
                            corrected = self._MONTH_CORRECTIONS.get(
                                arabic.strip(), arabic,
                            )
                            v.text = corrected
                            changes += 1

        # Translate series names (always — appear in legend, not segments)
        for ser in chart_elem.iter(f'{{{c_ns}}}ser'):
            for tx in ser.findall(f'{{{c_ns}}}tx'):
                for v in tx.iter(f'{{{c_ns}}}v'):
                    if v.text:
                        original = v.text.strip()
                        arabic = self.translations.get(original)
                        if arabic:
                            v.text = arabic
                            changes += 1

        return changes

    def _handle_connector_arrowheads(self, shape) -> int:
        """Swap head/tail arrowhead markers on connector shapes for RTL."""
        try:
            sp_el = shape._element
            ln = sp_el.find(f'.//{{{A_NS}}}ln')
            if ln is None:
                return 0

            head = ln.find(f'{{{A_NS}}}headEnd')
            tail = ln.find(f'{{{A_NS}}}tailEnd')
            if head is None and tail is None:
                return 0

            head_copy = deepcopy(head) if head is not None else None
            tail_copy = deepcopy(tail) if tail is not None else None

            if head is not None:
                ln.remove(head)
            if tail is not None:
                ln.remove(tail)

            if tail_copy is not None:
                tail_copy.tag = f'{{{A_NS}}}headEnd'
                ln.append(tail_copy)
            if head_copy is not None:
                head_copy.tag = f'{{{A_NS}}}tailEnd'
                ln.append(head_copy)

            return 1
        except Exception as exc:
            logger.debug('_handle_connector_arrowheads: %s', exc)
            return 0

    def _process_group_children(
        self, group_shape, cls: ShapeClassification
    ) -> int:
        """
        Process children of a group shape:
        - Text: translate + RTL on all text-bearing children
        - Direction: reverse directional children (arrows, connectors)
        - No position changes on children (positions are group-relative)
        """
        changes = 0

        # Text treatment on all text-bearing children
        for child in self._collect_text_shapes_from_group(group_shape):
            changes += self._apply_translation(child)
            child_cls = ShapeClassification(
                role=ShapeRole.CONTENT_TEXT,
                position_action='keep',
                text_action='translate_rtl',
                direction_action='none',
                rule_name='group_child',
            )
            changes += self._set_rtl_alignment(child, child_cls)

            # Tables/charts inside groups
            if getattr(child, 'has_table', False) and child.has_table:
                changes += self._handle_table(child)
            if getattr(child, 'has_chart', False) and child.has_chart:
                changes += self._handle_chart(child)

        # Directional reversal on group children
        try:
            for child in group_shape.shapes:
                changes += self._reverse_child_direction(child)
        except Exception:
            pass

        return changes

    def _process_complex_graphic_children(
        self, group_shape, cls: ShapeClassification
    ) -> int:
        """
        P1-2: Complex graphic handler — translate text ONLY.
        Unlike _process_group_children, this does NOT reverse directions
        or change any child positions. Hard rule: complex shapes and
        graphics should only be translated, not mirrored.
        """
        changes = 0
        for child in self._collect_text_shapes_from_group(group_shape):
            changes += self._apply_translation(child)
            child_cls = ShapeClassification(
                role=ShapeRole.CONTENT_TEXT,
                position_action='keep',
                text_action='translate_rtl',
                direction_action='none',
                rule_name='complex_graphic_child',
            )
            changes += self._set_rtl_alignment(child, child_cls)

            # Tables/charts inside complex graphics still get RTL treatment
            if getattr(child, 'has_table', False) and child.has_table:
                changes += self._handle_table(child)
            if getattr(child, 'has_chart', False) and child.has_chart:
                changes += self._handle_chart(child)

        return changes

    def _reverse_child_direction(self, child) -> int:
        """Reverse directional shapes and connectors inside a group."""
        changes = 0

        # Directional preset swap
        try:
            sp_el = child._element
            prst_geom = sp_el.find(f'.//{{{A_NS}}}prstGeom')
            if prst_geom is not None:
                prst = prst_geom.get('prst', '')

                # Skip data-encoding shapes (arcs, pies) — they encode
                # percentage fill, not direction.  Flipping them changes
                # the visual meaning of the infographic.
                if prst in self._DATA_SHAPE_PRESETS:
                    return 0

                action = _DIRECTIONAL_SWAP.get(prst)
                if action is not None:
                    if action == '_flipH':
                        xfrm = sp_el.find(f'.//{{{A_NS}}}xfrm')
                        if xfrm is not None:
                            current = xfrm.get('flipH', '0')
                            xfrm.set('flipH', '0' if current == '1' else '1')
                            changes += 1
                    elif action != prst:
                        prst_geom.set('prst', action)
                        changes += 1
        except Exception:
            pass

        # Connector arrowhead swap
        try:
            sp_el = child._element
            tag = sp_el.tag.split('}')[-1] if '}' in sp_el.tag else sp_el.tag
            if tag == 'cxnSp':
                changes += self._handle_connector_arrowheads(child)
                # Toggle flipH on connector
                for xfrm in sp_el.iter(f'{{{A_NS}}}xfrm'):
                    current = (xfrm.get('flipH', '0') or '0').lower()
                    if current in ('1', 'true'):
                        if 'flipH' in xfrm.attrib:
                            del xfrm.attrib['flipH']
                    else:
                        xfrm.set('flipH', '1')
                    changes += 1
                    break
        except Exception:
            pass

        return changes

    # ─────────────────────────────────────────────────────────────────────
    # Logo deduplication (v1 Fix 24 equivalent)
    # ─────────────────────────────────────────────────────────────────────

    def _dedup_mirrored_logos(
        self,
        shapes: List,
        result: 'SlideClassificationResult',
        slide_number: int,
    ) -> int:
        """
        Remove duplicate logo images created by mirroring.

        Detection criteria (ALL must be true):
          1. Both shapes are images (pic or blipFill)
          2. Both reference the same rEmbed relationship ID
          3. Their widths are within 5% of each other
          4. Their heights are within 5% of each other
          5. They sit on opposite sides of the slide midpoint

        Removes the shape on the RIGHT side (the original pre-mirror position;
        mirroring copies right→left, so the left copy is the intended one).
        """
        changes = 0
        R_NS_URI = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        half_width = self._slide_width // 2

        # Collect image shapes with their rEmbed and geometry
        image_info = []  # (shape, rEmbed, left, top, width, height, center_x)

        for shape in shapes:
            try:
                sp_el = shape._element
                tag = sp_el.tag.split('}')[-1] if '}' in sp_el.tag else sp_el.tag

                is_pic = tag == 'pic'
                blip_fill = sp_el.find(f'.//{{{A_NS}}}blipFill')
                if not is_pic and blip_fill is None:
                    continue

                if blip_fill is None:
                    continue
                blip = blip_fill.find(f'{{{A_NS}}}blip')
                if blip is None:
                    continue
                r_embed = blip.get(f'{{{R_NS_URI}}}embed')
                if not r_embed:
                    continue

                left = int(shape.left or 0)
                top = int(shape.top or 0)
                width = int(shape.width or 0)
                height = int(shape.height or 0)
                center_x = left + width // 2

                image_info.append((shape, r_embed, left, top, width, height, center_x))
            except Exception:
                continue

        if len(image_info) < 2:
            return 0

        # Check all pairs for duplicates
        removed_ids: Set[int] = set()
        for i in range(len(image_info)):
            if image_info[i][0].shape_id in removed_ids:
                continue
            for j in range(i + 1, len(image_info)):
                if image_info[j][0].shape_id in removed_ids:
                    continue

                s_a, embed_a, l_a, t_a, w_a, h_a, cx_a = image_info[i]
                s_b, embed_b, l_b, t_b, w_b, h_b, cx_b = image_info[j]

                # Criterion 1: same rEmbed
                if embed_a != embed_b:
                    continue

                # Criterion 2: sizes within 5%
                if w_a == 0 or h_a == 0 or w_b == 0 or h_b == 0:
                    continue
                w_ratio = abs(w_a - w_b) / max(w_a, w_b)
                h_ratio = abs(h_a - h_b) / max(h_a, h_b)
                if w_ratio > 0.05 or h_ratio > 0.05:
                    continue

                # Criterion 3: opposite sides of slide
                a_is_left = cx_a < half_width
                b_is_left = cx_b < half_width
                if a_is_left == b_is_left:
                    continue

                # Remove the shape on the RIGHT (original pre-mirror position)
                to_remove = s_b if a_is_left else s_a

                try:
                    parent = to_remove._element.getparent()
                    if parent is not None:
                        parent.remove(to_remove._element)
                        removed_ids.add(to_remove.shape_id)
                        changes += 1
                        logger.debug(
                            'Logo dedup slide %d: removed duplicate image '
                            '(rEmbed=%s, shape_id=%d)',
                            slide_number, embed_a, to_remove.shape_id,
                        )
                except Exception:
                    pass

        return changes

    # ─────────────────────────────────────────────────────────────────────
    # Phase 5: Master/layout decorative shape mirroring
    # ─────────────────────────────────────────────────────────────────────

    def _mirror_master_decorative_shapes(self, slide, slide_number: int) -> int:
        """
        Detect thin decorative shapes (lines, bars) inherited from the master
        or layout slide and add mirrored copies on the slide layer.

        These shapes live in slideLayout/slideMaster XML and are not part of
        the slide's spTree, so Phase 3 never sees them.

        Strategy (per Codex 5.3 / Sonnet 4.6):
        - Scan the slide's layout and master for non-placeholder shapes
        - Identify thin vertical/horizontal lines or bars (width or height < 3% slide dim)
        - Clone them onto the slide spTree at the mirrored X position
        - Add a same-background occluder at the original position to hide the inherited one
        """
        changes = 0
        try:
            slide_layout = slide.slide_layout
            sources = [slide_layout]
            if hasattr(slide_layout, 'slide_master'):
                sources.append(slide_layout.slide_master)

            thin_threshold_w = int(self._slide_width * 0.03)  # 3% of slide width
            thin_threshold_h = int(self._slide_height * 0.03)

            spTree = slide._element.find(f'.//{{{P_NS}}}spTree')
            if spTree is None:
                return 0

            for source in sources:
                if source is None:
                    continue
                try:
                    source_shapes = source.shapes
                except Exception:
                    continue

                for shape in source_shapes:
                    # Skip placeholders (they have their own handling)
                    try:
                        if getattr(shape, 'placeholder_format', None) is not None:
                            continue
                    except Exception:
                        pass

                    try:
                        s_left = int(shape.left or 0)
                        s_top = int(shape.top or 0)
                        s_width = int(shape.width or 0)
                        s_height = int(shape.height or 0)
                    except (TypeError, ValueError):
                        continue

                    # Only target thin decorative shapes (lines, thin bars)
                    is_thin = (s_width < thin_threshold_w or s_height < thin_threshold_h)
                    if not is_thin:
                        continue

                    # Check if off-center (not already centered)
                    center = s_left + s_width // 2
                    slide_center = self._slide_width // 2
                    if abs(center - slide_center) < _POSITION_TOLERANCE_EMU:
                        continue  # Already centered, no need to mirror

                    # Mirror the position
                    new_left = mirror_x(s_left, s_width, self._slide_width)
                    if abs(new_left - s_left) < _POSITION_TOLERANCE_EMU:
                        continue

                    # Clone the shape element to the slide spTree at mirrored position
                    cloned = deepcopy(shape._element)
                    xfrm = cloned.find(f'.//{{{A_NS}}}xfrm')
                    if xfrm is not None:
                        off = xfrm.find(f'{{{A_NS}}}off')
                        if off is not None:
                            off.set('x', str(new_left))

                    spTree.append(cloned)
                    changes += 1

                    # Add occluder rectangle at original position to hide inherited shape
                    # (simple white/background-colored rectangle)
                    occluder = etree.SubElement(spTree, f'{{{P_NS}}}sp')
                    nvSpPr = etree.SubElement(occluder, f'{{{P_NS}}}nvSpPr')
                    cNvPr = etree.SubElement(nvSpPr, f'{{{P_NS}}}cNvPr')
                    cNvPr.set('id', '0')
                    cNvPr.set('name', 'SlideArabi Occluder')
                    cNvSpPr = etree.SubElement(nvSpPr, f'{{{P_NS}}}cNvSpPr')
                    nvPr = etree.SubElement(nvSpPr, f'{{{P_NS}}}nvPr')
                    sp_pr = etree.SubElement(occluder, f'{{{P_NS}}}spPr')
                    occ_xfrm = etree.SubElement(sp_pr, f'{{{A_NS}}}xfrm')
                    occ_off = etree.SubElement(occ_xfrm, f'{{{A_NS}}}off')
                    occ_off.set('x', str(s_left))
                    occ_off.set('y', str(s_top))
                    occ_ext = etree.SubElement(occ_xfrm, f'{{{A_NS}}}ext')
                    occ_ext.set('cx', str(s_width))
                    occ_ext.set('cy', str(s_height))
                    prstGeom = etree.SubElement(sp_pr, f'{{{A_NS}}}prstGeom')
                    prstGeom.set('prst', 'rect')
                    # Match background (white fill)
                    solidFill = etree.SubElement(sp_pr, f'{{{A_NS}}}solidFill')
                    srgbClr = etree.SubElement(solidFill, f'{{{A_NS}}}srgbClr')
                    srgbClr.set('val', 'FFFFFF')
                    # No outline
                    ln = etree.SubElement(sp_pr, f'{{{A_NS}}}ln')
                    noFill = etree.SubElement(ln, f'{{{A_NS}}}noFill')

                    logger.debug(
                        'Slide %d: mirrored master line "%s" left %d → %d + occluder',
                        slide_number, getattr(shape, 'name', '?'), s_left, new_left,
                    )

        except Exception as exc:
            logger.warning('_mirror_master_decorative_shapes slide %d: %s', slide_number, exc)

        return changes

    # ─────────────────────────────────────────────────────────────────────
    # Phase 4: Post-processing (text-only, no position changes)
    # ─────────────────────────────────────────────────────────────────────

    def _post_process(
        self,
        shapes: List,
        result: SlideClassificationResult,
        slide_number: int,
    ) -> int:
        """
        Post-processing fixups that operate on text only.
        No position changes allowed in this phase.

        Preserved from v1:
        - Fix 11: wrap=none expansion for Arabic text
        - Fix 12: title-body vertical overlap resolution
        - Fix 19: bidi base direction on mixed Arabic/English titles
        - Fix 22: Arabic autofit (normAutofit + font reduction)
        """
        changes = 0

        # Fix 11: wrap=none expansion
        for shape in shapes:
            try:
                if getattr(shape, 'has_text_frame', False):
                    changes += self._fix_wrap_none(shape)
            except Exception:
                pass

        # Fix 12: title-body overlap
        changes += self._fix_title_body_overlap(shapes, slide_number)

        # Fix 19: bidi base direction
        changes += self._fix_bidi_base_direction(shapes, slide_number)

        # Fix 22: Arabic autofit
        changes += self._apply_arabic_autofit(shapes, result, slide_number)

        # Logo deduplication (v1 Fix 24 equivalent)
        changes += self._dedup_mirrored_logos(shapes, result, slide_number)

        return changes

    def _fix_wrap_none(self, shape) -> int:
        """
        Fix 11: Expand text boxes with wrap='none' that now contain Arabic text.

        Arabic text is wider than English, so nowrap boxes may clip.
        1. Change wrap to 'square' for wrapping support
        2. Add normAutofit (80% shrink) so text scales down if needed
        3. RTL-aware width expansion: grow left to preserve right anchor
        """
        changes = 0
        try:
            tf = shape.text_frame
            body_pr = tf._txBody.find(f'{{{A_NS}}}bodyPr')
            if body_pr is None:
                return 0

            wrap_val = body_pr.get('wrap', 'square')
            if wrap_val != 'none':
                return 0

            shape_text = tf.text or ''
            if not has_arabic(shape_text):
                return 0

            # Get current geometry for width expansion
            sp_el = shape._element
            xfrm = sp_el.find(f'.//{{{A_NS}}}xfrm')
            if xfrm is None:
                return 0
            off = xfrm.find(f'{{{A_NS}}}off')
            ext = xfrm.find(f'{{{A_NS}}}ext')
            if off is None or ext is None:
                return 0

            try:
                x = int(off.get('x', 0))
                cx = int(ext.get('cx', 0))
            except (ValueError, TypeError):
                return 0

            # 1. Change wrap to 'square'
            body_pr.set('wrap', 'square')
            changes += 1

            # 2. Add normAutofit if not present
            for child_tag in ('spAutoFit', 'noAutofit'):
                for child in body_pr.findall(f'{{{A_NS}}}{child_tag}'):
                    body_pr.remove(child)
            if body_pr.find(f'{{{A_NS}}}normAutofit') is None:
                autofit = etree.SubElement(body_pr, f'{{{A_NS}}}normAutofit')
                autofit.set('fontScale', '80000')
                changes += 1

            # 3. RTL-aware width expansion (preserve right anchor)
            ARABIC_EXPANSION_FACTOR = 1.35
            MAX_EXPANSION_FACTOR = 2.0

            estimated_cx = int(cx * ARABIC_EXPANSION_FACTOR)
            max_cx = int(cx * MAX_EXPANSION_FACTOR)
            max_slide_cap = int(self._slide_width * 0.85)
            cap_cx = min(max_cx, max_slide_cap)
            new_cx = min(estimated_cx, cap_cx)

            if new_cx > cx:
                delta = new_cx - cx
                new_x = x - delta  # Shift LEFT to preserve right anchor
                if new_x < 0:
                    delta = x
                    new_x = 0
                    new_cx = cx + delta

                if new_cx > cx:
                    ext.set('cx', str(new_cx))
                    off.set('x', str(new_x))
                    changes += 1
                    logger.debug(
                        'Fix 11: RTL-aware expand textbox "%s" '
                        'cx %d->%d, x %d->%d (right edge preserved at %d)',
                        getattr(shape, 'name', '?'), cx, new_cx, x, new_x,
                        new_x + new_cx,
                    )

        except Exception as exc:
            logger.debug('_fix_wrap_none: %s', exc)

        return changes

    def _fix_title_body_overlap(
        self, shapes: List, slide_number: int
    ) -> int:
        """
        Fix 12: Detect and resolve vertical overlap between title and body
        placeholder text frames after RTL repositioning.

        If a title and body overlap vertically, nudge the body down to
        clear the title bottom edge.
        """
        title_shape = None
        body_shape = None

        for shape in shapes:
            if not getattr(shape, 'is_placeholder', False):
                continue
            ph_info = get_placeholder_info(shape)
            if ph_info is None:
                continue
            ph_type = ph_info[0]
            if ph_type and 'title' in str(ph_type).lower():
                title_shape = shape
            elif ph_type and ('body' in str(ph_type).lower() or 'subtitle' in str(ph_type).lower()):
                body_shape = shape

        if title_shape is None or body_shape is None:
            return 0

        try:
            title_bottom = int(title_shape.top or 0) + int(title_shape.height or 0)
            body_top = int(body_shape.top or 0)

            if body_top < title_bottom:
                overlap = title_bottom - body_top
                body_shape.top = title_bottom
                logger.debug(
                    'Fix 12 slide %d: nudged body down %d EMU to clear title',
                    slide_number, overlap,
                )
                return 1
        except Exception:
            pass
        return 0

    def _fix_bidi_base_direction(
        self, shapes: List, slide_number: int
    ) -> int:
        """
        Fix 19: Ensure bidi base direction is set on paragraphs with mixed
        Arabic/English content in title-sized shapes (>=16pt).

        For mixed Arabic/English text, set bodyPr rtlCol='1', paragraph
        rtl='1', and right-align to ensure the Unicode bidi algorithm
        renders correctly.
        """
        changes = 0
        for shape in shapes:
            if not getattr(shape, 'has_text_frame', False):
                continue
            try:
                tf = shape.text_frame
                text = (tf.text or '').strip()
                if not text:
                    continue

                # Must have both Arabic and Latin characters
                if not has_arabic(text):
                    continue
                has_latin = any(c.isascii() and c.isalpha() for c in text)
                if not has_latin:
                    continue

                # Font-size guard: only apply to title-sized text (>=16pt = 1600)
                max_font = 0
                sp_el = shape._element
                for rPr in sp_el.iter(f'{{{A_NS}}}rPr'):
                    sz_str = rPr.get('sz')
                    if sz_str:
                        try:
                            max_font = max(max_font, int(sz_str))
                        except ValueError:
                            pass
                if max_font > 0 and max_font < 1600:
                    continue

                # Body-level RTL
                body_pr = tf._txBody.find(f'{{{A_NS}}}bodyPr')
                if body_pr is not None:
                    if body_pr.get('rtlCol') != '1':
                        body_pr.set('rtlCol', '1')
                        changes += 1

                # Paragraph-level RTL + alignment
                for para in tf.paragraphs:
                    pPr = ensure_pPr(para._p)
                    if pPr.get('rtl') != '1':
                        pPr.set('rtl', '1')
                        changes += 1
                    current_algn = pPr.get('algn', '')
                    if current_algn not in ('r', 'ctr'):
                        pPr.set('algn', 'r')
                        changes += 1

                if changes:
                    logger.debug(
                        'Fix 19 slide %d: bidi direction set on mixed text "%s"',
                        slide_number, text[:40],
                    )
            except Exception:
                pass
        return changes

    def _apply_arabic_autofit(
        self,
        shapes: List,
        result: SlideClassificationResult,
        slide_number: int,
    ) -> int:
        """
        Fix 22: Apply normAutofit to text frames containing Arabic text.

        Two-pronged: normAutofit on bodyPr (PowerPoint) + direct font
        reduction on runs (LibreOffice compatibility).
        """
        changes = 0
        for shape in shapes:
            try:
                if not getattr(shape, 'has_text_frame', False):
                    continue
                tf = shape.text_frame
                shape_text = tf.text or ''
                if not has_arabic(shape_text):
                    continue

                body_pr = tf._txBody.find(f'{{{A_NS}}}bodyPr')
                if body_pr is None:
                    continue

                # Skip if already has normAutofit
                if body_pr.find(f'{{{A_NS}}}normAutofit') is not None:
                    continue

                # Estimate overflow
                box_width_emu = int(shape.width or 0) if shape.width else None
                box_height_emu = int(shape.height or 0) if shape.height else None

                max_font_hundredths = 0
                for rPr in shape._element.iter(f'{{{A_NS}}}rPr'):
                    sz_str = rPr.get('sz')
                    if sz_str:
                        try:
                            max_font_hundredths = max(max_font_hundredths, int(sz_str))
                        except ValueError:
                            pass
                for def_rPr in shape._element.iter(f'{{{A_NS}}}defRPr'):
                    sz_str = def_rPr.get('sz')
                    if sz_str:
                        try:
                            max_font_hundredths = max(max_font_hundredths, int(sz_str))
                        except ValueError:
                            pass

                char_count = len(shape_text.replace('\n', ''))

                # P0-6: Skip Fix 22 entirely for already-small fonts (≤12pt)
                # Arabic at 12pt or below has no headroom for reduction
                if max_font_hundredths > 0 and max_font_hundredths <= 1200:
                    continue

                if max_font_hundredths > 0 and box_width_emu and box_width_emu > 0:
                    font_pt = max_font_hundredths / 100
                    # P1-1: Fixed overflow estimation constants
                    # Arabic glyphs avg ~0.45x em (was 0.6x), removed 1.3x padding
                    estimated_char_width_emu = 0.45 * font_pt * 12700
                    lines = shape_text.split('\n')
                    max_line_chars = max(
                        (len(l.strip()) for l in lines), default=char_count,
                    )
                    estimated_line_width = (
                        max_line_chars * estimated_char_width_emu
                    )
                    overflow_ratio = estimated_line_width / box_width_emu

                    height_overflow = 0.0
                    if box_height_emu and box_height_emu > 0:
                        line_count = len(lines)
                        # Realistic OOXML default line spacing (was 1.5x)
                        estimated_line_height = font_pt * 12700 * 1.15
                        estimated_total_height = line_count * estimated_line_height
                        height_overflow = estimated_total_height / box_height_emu

                    # Require genuine overflow before triggering (was 0.9 = 90% fill)
                    threshold = 1.1
                    if (
                        max_font_hundredths >= 2400
                        and box_width_emu > self._slide_width * 0.40
                    ):
                        threshold = 1.3
                    if overflow_ratio < threshold and height_overflow < threshold:
                        continue

                    max_overflow = max(overflow_ratio, height_overflow)
                    # P0-5: Raised scale floor from 50% to 70% (prevents extreme shrink)
                    font_scale = max(int(100000 / max_overflow), 70000)
                    font_scale = min(font_scale, 100000)
                else:
                    # P0-4: Cannot measure overflow without explicit font size.
                    # Was: font_scale = 90000 (blind 10% reduction on ALL theme-inherited fonts)
                    # Now: skip — let normAutofit handle rendering without destructive reduction
                    continue

                # Remove existing autofit variants
                for child_tag in ('spAutoFit', 'noAutofit'):
                    for child in body_pr.findall(f'{{{A_NS}}}{child_tag}'):
                        body_pr.remove(child)

                # Add normAutofit
                autofit_el = etree.SubElement(body_pr, f'{{{A_NS}}}normAutofit')
                autofit_el.set('fontScale', str(font_scale))

                # Direct font size reduction for LibreOffice
                font_scale_ratio = font_scale / 100000
                for para in tf.paragraphs:
                    for run in para.runs:
                        if run.font.size is not None:
                            original = int(run.font.size)
                            new_size = max(
                                int(original * font_scale_ratio), 12 * 12700,  # 12pt min for Arabic readability
                            )
                            run.font.size = new_size
                        else:
                            rPr = run._r.find(f'{{{A_NS}}}rPr')
                            if rPr is None:
                                rPr = etree.SubElement(run._r, f'{{{A_NS}}}rPr')
                                run._r.insert(0, rPr)
                            default_sz = None
                            for src in [
                                para._p.find(f'.//{{{A_NS}}}defRPr'),
                                para._p.find(f'{{{A_NS}}}endParaRPr'),
                            ]:
                                if src is not None and src.get('sz'):
                                    try:
                                        default_sz = int(src.get('sz'))
                                    except ValueError:
                                        pass
                                    break
                            if default_sz:
                                new_sz = max(
                                    int(default_sz * font_scale_ratio), 1200,  # 1200 hundredths = 12pt min
                                )
                                rPr.set('sz', str(new_sz))

                changes += 1
                logger.debug(
                    'Fix 22 slide %d: autofit on "%s" fontScale=%d%%',
                    slide_number,
                    getattr(shape, 'name', '?'),
                    font_scale // 1000,
                )

            except Exception as exc:
                logger.debug('Fix 22 slide %d: %s', slide_number, exc)

        return changes

    # ─────────────────────────────────────────────────────────────────────
    # Telemetry
    # ─────────────────────────────────────────────────────────────────────

    def _log_translation_coverage(
        self, shapes: List, slide_number: int
    ) -> None:
        """Log a warning if zero text shapes were translated on a slide."""
        translated = 0
        total_text = 0
        for s in shapes:
            if getattr(s, 'has_text_frame', False) and any(
                p.text.strip() for p in s.text_frame.paragraphs if p.text
            ):
                total_text += 1
                if any(
                    has_arabic(p.text)
                    for p in s.text_frame.paragraphs
                    if p.text
                ):
                    translated += 1
        if total_text > 0 and translated == 0:
            logger.error(
                'Slide %d: 0/%d text shapes translated — possible extraction miss',
                slide_number, total_text,
            )
