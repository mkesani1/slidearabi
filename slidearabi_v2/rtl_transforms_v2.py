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
        self._translations_lower: Dict[str, str] = {}
        for key, val in self.translations.items():
            lower_key = key.strip().lower()
            if lower_key not in self._translations_lower:
                self._translations_lower[lower_key] = val

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
                    # Still apply direction + text
                    changes += self._dispatch_direction(shape, cls)
                    changes += self._dispatch_text(shape, cls)
                    # Handle group children text
                    if cls.role == ShapeRole.GROUP:
                        changes += self._process_group_children(shape, cls)
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
                elif cls.role == ShapeRole.CONNECTOR:
                    changes += self._handle_connector_arrowheads(shape)

            except Exception as exc:
                logger.warning(
                    'Slide %d shape "%s": %s',
                    slide_number, getattr(shape, 'name', '?'), exc,
                )

        # ── Phase 4: Post-processing (text-only) ──────────────────────
        changes += self._post_process(all_shapes, result, slide_number)

        # v1 post-processing for roles handled by v1 dispatcher
        if self._v1_dispatcher is not None:
            changes += self._v1_dispatcher.dispatch_post_processing(
                all_shapes, slide_number,
            )

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
        Swap left and right panel zones as units.

        All shapes classified as PANEL_LEFT shift right by shift_delta,
        and all PANEL_RIGHT shapes shift left by the same amount.
        This preserves internal spatial relationships within each panel.
        """
        panel = result.context.split_panel
        if panel is None:
            return 0

        changes = 0
        delta = panel.shift_delta

        for shape in shapes:
            sid = shape.shape_id
            cls = result.get(shape)

            if cls.role == ShapeRole.PANEL_LEFT:
                try:
                    shape.left = int(shape.left) + delta
                    swapped_ids.add(sid)
                    changes += 1
                except Exception as exc:
                    logger.debug('Panel swap left→right failed: %s', exc)

            elif cls.role == ShapeRole.PANEL_RIGHT:
                try:
                    shape.left = int(shape.left) - delta
                    swapped_ids.add(sid)
                    changes += 1
                except Exception as exc:
                    logger.debug('Panel swap right→left failed: %s', exc)

        if changes:
            logger.debug(
                'Panel swap: %d shapes moved, delta=%d EMU', changes, delta,
            )
        return changes

    # ─────────────────────────────────────────────────────────────────────
    # Phase 3: Position dispatch
    # ─────────────────────────────────────────────────────────────────────

    def _dispatch_position(
        self, shape, cls: ShapeClassification, layout
    ) -> int:
        handler = self._position_handlers.get(cls.position_action)
        if handler is None:
            return 0
        return handler(shape, cls, layout)

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
        Placeholder: remove local xfrm to inherit from (RTL-transformed) layout.

        If the inherited position would cause overlap with large freeform shapes,
        fall back to explicit mirroring.
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

            # Remove local xfrm so position inherits from layout
            sp_pr.remove(xfrm)
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

    def _handle_remove_flip(self, shape, cls: ShapeClassification) -> int:
        """Remove flipH/flipV from shape xfrm — prevents content mirroring."""
        changes = 0
        try:
            sp_el = shape._element
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
        """Full treatment: translate text + set RTL alignment."""
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
        3. Longest prefix match (>80% of text length, for text >40 chars)
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

        # 3. Longest prefix match for long text
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
                    continue

                runs = para.runs
                if runs:
                    runs[0].text = arabic_text
                    for run in runs[1:]:
                        run.text = ''
                else:
                    p_elem = para._p
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

    def _set_rtl_alignment(self, shape, cls: ShapeClassification) -> int:
        """
        Set paragraph-level RTL direction and alignment on every paragraph.

        CRITICAL: rtl='1' is ONLY set on paragraphs containing Arabic script.
        Setting rtl='1' on English text causes OOXML bidi reordering bugs.
        """
        changes = 0
        ph_type = cls.placeholder_type

        try:
            tf = shape.text_frame
            for para in tf.paragraphs:
                text = para.text or ''
                if not text.strip():
                    continue

                pPr = ensure_pPr(para._p)

                if has_arabic(text):
                    pPr.set('rtl', '1')
                else:
                    pPr.set('rtl', '0')

                alignment = self._compute_paragraph_alignment(text, ph_type)
                pPr.set('algn', alignment)
                changes += 1

            # Set rtlCol on bodyPr for shapes with Arabic content
            shape_text = tf.text or ''
            if has_arabic(shape_text):
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
        self, text: str, ph_type: Optional[str]
    ) -> str:
        """
        Compute OOXML algn value for a paragraph in RTL context.

        Priority:
        1. Footer placeholders → 'l'
        2. Title placeholders (not subtitle) → 'ctr'
        3. Arabic-dominant → 'r'
        4. Latin-dominant, no Arabic → 'l'
        5. Mixed → 'r'
        """
        _footer_substrings = (
            'ftr', 'footer', 'sldnum', 'slide_number', 'date_time', 'date',
        )
        if ph_type in _FOOTER_PH_TYPES or (
            ph_type and any(f in ph_type for f in _footer_substrings)
        ):
            return 'l'

        if ph_type and 'subtitle' not in ph_type.lower() and 'title' in ph_type.lower():
            return 'ctr'

        ratios = compute_script_ratio(text)
        if ratios['arabic'] > 0.70:
            return 'r'
        elif ratios['latin'] > 0.70 and not has_arabic(text):
            return 'l'
        else:
            return 'r'

    # ─────────────────────────────────────────────────────────────────────
    # Role-specific handlers
    # ─────────────────────────────────────────────────────────────────────

    def _handle_table(self, shape) -> int:
        """
        Transform a table for RTL: reverse columns, set rtl='1' on tblPr,
        translate cell text, set RTL properties on cells.
        """
        changes = 0
        try:
            table = shape.table
            num_cols = len(table.columns)
            if num_cols <= 1:
                return 0

            tbl_elem = table._tbl

            # Translate cell text before reversing
            if self.translations:
                for row in table.rows:
                    for cell in row.cells:
                        if cell.text_frame:
                            changes += self._translate_cell_text(cell.text_frame)

            # Reverse column widths
            tbl_grid = tbl_elem.find(f'{{{A_NS}}}tblGrid')
            if tbl_grid is not None:
                grid_cols = tbl_grid.findall(f'{{{A_NS}}}gridCol')
                if len(grid_cols) == num_cols:
                    widths = [gc.get('w', '0') for gc in grid_cols]
                    for i, gc in enumerate(grid_cols):
                        gc.set('w', widths[num_cols - 1 - i])
                    changes += 1

            # Reverse cell order in each row
            for row in table.rows:
                tr_elem = row._tr
                tc_elems = tr_elem.findall(f'{{{A_NS}}}tc')
                cell_copies = [deepcopy(tc) for tc in tc_elems]
                for tc in tc_elems:
                    tr_elem.remove(tc)
                for tc in reversed(cell_copies):
                    tr_elem.append(tc)
                changes += 1

            # Set table-level RTL
            tbl_pr = tbl_elem.find(f'{{{A_NS}}}tblPr')
            if tbl_pr is None:
                tbl_pr = etree.SubElement(tbl_elem, f'{{{A_NS}}}tblPr')
                tbl_elem.insert(0, tbl_pr)
            tbl_pr.set('rtl', '1')
            changes += 1

            # RTL properties on cell text
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
        """Set RTL paragraph properties on cell text."""
        try:
            for para in text_frame.paragraphs:
                text = para.text or ''
                if not text.strip():
                    continue
                pPr = ensure_pPr(para._p)
                if has_arabic(text):
                    pPr.set('rtl', '1')
                else:
                    pPr.set('rtl', '0')
                # Right-align first column, left-align last, default right
                if col_idx == 0:
                    pPr.set('algn', 'r')
                elif col_idx == num_cols - 1:
                    pPr.set('algn', 'l')
                else:
                    pPr.set('algn', 'r')

            body_pr = text_frame._txBody.find(f'{{{A_NS}}}bodyPr')
            if body_pr is not None:
                body_pr.set('rtlCol', '1')
        except Exception:
            pass

    def _handle_chart(self, shape) -> int:
        """
        Transform a chart for RTL: reverse category axis, mirror legend,
        translate labels.

        Placeholder — chart transformation requires chart part access
        and is complex. This skeleton defers to v1 implementation when
        integrated into the full pipeline.
        """
        changes = 0
        try:
            if not (getattr(shape, 'has_chart', False) and shape.has_chart):
                return 0

            chart = shape.chart
            chart_xml = chart._chartSpace

            # Reverse category axis direction
            for cat_ax in chart_xml.iter(f'{{{A_NS}}}catAx'):
                scaling = cat_ax.find(f'{{{A_NS}}}scaling')
                if scaling is not None:
                    orientation = scaling.find(f'{{{A_NS}}}orientation')
                    if orientation is not None:
                        current = orientation.get('val', 'minMax')
                        orientation.set(
                            'val',
                            'maxMin' if current == 'minMax' else 'minMax',
                        )
                        changes += 1

            # Mirror legend position
            legend = chart_xml.find(f'.//{{{A_NS}}}legend')
            if legend is not None:
                legend_pos = legend.find(f'{{{A_NS}}}legendPos')
                if legend_pos is not None:
                    pos = legend_pos.get('val', '')
                    if pos == 'r':
                        legend_pos.set('val', 'l')
                        changes += 1
                    elif pos == 'l':
                        legend_pos.set('val', 'r')
                        changes += 1

        except Exception as exc:
            logger.debug('_handle_chart: %s', exc)
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

    def _reverse_child_direction(self, child) -> int:
        """Reverse directional shapes and connectors inside a group."""
        changes = 0

        # Directional preset swap
        try:
            sp_el = child._element
            prst_geom = sp_el.find(f'.//{{{A_NS}}}prstGeom')
            if prst_geom is not None:
                prst = prst_geom.get('prst', '')
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
                if max_font_hundredths > 0 and box_width_emu and box_width_emu > 0:
                    font_pt = max_font_hundredths / 100
                    estimated_char_width_emu = 0.6 * font_pt * 12700
                    lines = shape_text.split('\n')
                    max_line_chars = max(
                        (len(l.strip()) for l in lines), default=char_count,
                    )
                    estimated_line_width = (
                        max_line_chars * estimated_char_width_emu * 1.3
                    )
                    overflow_ratio = estimated_line_width / box_width_emu

                    height_overflow = 0.0
                    if box_height_emu and box_height_emu > 0:
                        line_count = len(lines)
                        estimated_line_height = font_pt * 12700 * 1.5
                        estimated_total_height = line_count * estimated_line_height
                        height_overflow = estimated_total_height / box_height_emu

                    threshold = 0.9
                    if (
                        max_font_hundredths >= 2400
                        and box_width_emu > self._slide_width * 0.40
                    ):
                        threshold = 1.2
                    if overflow_ratio < threshold and height_overflow < threshold:
                        continue

                    max_overflow = max(overflow_ratio, height_overflow)
                    font_scale = max(int(100000 / max_overflow), 50000)
                    font_scale = min(font_scale, 100000)
                else:
                    font_scale = 90000

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
                                int(original * font_scale_ratio), 8 * 12700,
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
                                    int(default_sz * font_scale_ratio), 800,
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
