"""
SlideArabi v2 — Shape Intent Classification Layer

Classifies every shape on a slide ONCE before any RTL transforms run.
Each shape receives exactly one ShapeRole that determines its transform strategy.
No later pass may re-classify or override a shape's assigned role.

Architecture:
    1. Pre-scan: Collect raw geometry + detect slide-level patterns (panels, maps)
    2. Per-shape: Priority-ordered rule evaluation assigns one ShapeRole
    3. Validation: Assert every shape is classified, no duplicates

The classifier operates on ORIGINAL geometry only — never on mutated positions.
"""

from __future__ import annotations

import logging
from enum import Enum, auto
from dataclasses import dataclass, field
from typing import Dict, FrozenSet, List, Optional, Set, Tuple

from lxml import etree

# Import OOXML namespace constants and utilities from the existing codebase
from slidearabi.utils import (
    A_NS, P_NS, R_NS,
    mirror_x,
    bounds_check_emu,
    get_placeholder_info_from_xml,
)

logger = logging.getLogger(__name__)


# ═══════════════════════════════════════════════════════════════════════════════
# Constants — thresholds extracted from v1, consolidated here
# ═══════════════════════════════════════════════════════════════════════════════

# Background detection
_BG_WIDTH_FRACTION = 0.90       # width >= 90% of slide → background candidate
_BG_AREA_FRACTION = 0.85        # area >= 85% of slide → background (Gemini rule)
_BG_FULL_COVER_FRACTION = 0.95  # width AND height >= 95% → definite background

# Panel detection
_PANEL_MIN_WIDTH_FRACTION = 0.35   # anchor must be >= 35% slide width
_PANEL_MIN_HEIGHT_FRACTION = 0.50  # anchor must be >= 50% slide height
_FULL_WIDTH_FRACTION = 0.85        # shapes this wide span both panels

# Logo detection
_LOGO_MAX_WIDTH_FRACTION = 0.20    # image < 20% slide width = logo

# Bleed detection
_BLEED_THRESHOLD_EMU = 100_000     # ~0.11" — minimum bleed to classify as BLEED
_MAX_BLEED_EMU = 1_500_000         # ~1.64" — maximum allowed intentional bleed

# Overlay/map detection
_MAP_AREA_FRACTION = 0.40          # base image must cover >= 40% of slide area
_MAP_WIDTH_FRACTION = 0.60         # base image must be >= 60% of slide width
_OVERLAY_MAX_WIDTH_FRACTION = 0.30 # overlay shapes must be < 30% slide width
_OVERLAY_MIN_COUNT = 3             # need >= 3 small shapes to qualify as overlay

# Decorative (title/secHead layouts)
_DECORATIVE_MAX_WIDTH_FRACTION = 0.20   # small shape on title layout
_DECORATIVE_MAX_HEIGHT_FRACTION = 0.15  # short shape on title layout
_CONTENT_MIN_WIDTH_FRACTION = 0.20      # content-sized on title/secHead layout
_CONTENT_MIN_HEIGHT_FRACTION = 0.30     # content-sized on title/secHead layout

# Footer zone
_FOOTER_ZONE_FRACTION = 0.88      # shapes below 88% of slide height

# Badge detection
_BADGE_MAX_DIM_FRACTION = 0.08    # width AND height < 8% slide dim
_BADGE_MAX_TOP_FRACTION = 0.15    # badge must be in top 15%

# Position tolerance
_POSITION_TOLERANCE_EMU = 50_000  # ~0.055" — negligible position change

# Connector preset geometry names
_LINE_PRESETS = frozenset({
    'line', 'straightConnector1',
    'bentConnector2', 'bentConnector3', 'bentConnector4', 'bentConnector5',
    'curvedConnector2', 'curvedConnector3', 'curvedConnector4', 'curvedConnector5',
})

# Directional shape presets (chevrons, arrows)
_DIRECTIONAL_PRESETS = frozenset({
    'rightArrow', 'leftArrow', 'upArrow', 'downArrow',
    'leftRightArrow', 'upDownArrow', 'notchedRightArrow',
    'chevron', 'homePlate', 'pentagon',
    'stripedRightArrow', 'bentArrow', 'uturnArrow',
    'circularArrow', 'curvedRightArrow', 'curvedLeftArrow',
})

# Placeholder types
_TITLE_PH_TYPES = frozenset({'title', 'ctrTitle', 'center_title'})
_FOOTER_PH_TYPES = frozenset({'ftr', 'sldNum', 'dt', 'footer', 'slideNumber', 'date_time'})

# Title/secHead layouts where freeform shapes are treated differently
_CENTERED_LAYOUTS = frozenset({'title', 'secHead'})


# ═══════════════════════════════════════════════════════════════════════════════
# ShapeRole Enum
# ═══════════════════════════════════════════════════════════════════════════════

class ShapeRole(Enum):
    """
    Mutually exclusive shape classification for RTL transforms.
    Each role maps to exactly one transform strategy.
    Evaluated in PRIORITY ORDER — the first matching rule wins.
    """
    # ── Structural roles ──
    PLACEHOLDER      = auto()  # Layout-managed placeholder (title, body, subtitle)
    TABLE            = auto()  # Table shape — reverse columns, set RTL
    CHART            = auto()  # Chart shape — reverse axes, mirror legend
    CONNECTOR        = auto()  # Lines, arrows, connectors (cxnSp / line presets)
    DIRECTIONAL      = auto()  # Chevrons, arrow preset geometry shapes
    BACKGROUND       = auto()  # Full-slide background (image or solid fill)
    BLEED            = auto()  # Shape extending past slide edges intentionally
    FOOTER           = auto()  # Footer-zone shape (date, slide number, footer)
    BADGE            = auto()  # Small positional element (slide number badge)
    LOGO             = auto()  # Small brand image (< 20% slide width)
    OVERLAY          = auto()  # Shape on a large map/image (geographically anchored)
    PANEL_LEFT       = auto()  # Left half of split-panel layout
    PANEL_RIGHT      = auto()  # Right half of split-panel layout
    DECORATIVE       = auto()  # Accent bars, gradient panels, brand elements
    GROUP            = auto()  # Group shape containing mixed content
    CONTENT_IMAGE    = auto()  # Photo, illustration, screenshot (non-bg)
    CONTENT_TEXT     = auto()  # Freeform text box with content
    UNKNOWN          = auto()  # Fallback — apply conservative default (mirror)


# ═══════════════════════════════════════════════════════════════════════════════
# SlideRole Enum — slide-level layout classification
# ═══════════════════════════════════════════════════════════════════════════════

class SlideRole(Enum):
    """
    Slide-level layout classification for dispatch strategy.

    Detected during _pre_scan() based on shape geometry patterns.
    Influences per-shape handling (e.g., SPLIT_PANEL slides use panel swap).
    """
    STANDARD     = auto()  # Default; per-shape dispatch only
    SPLIT_PANEL  = auto()  # Two-zone layout: image half + content half
    PHOTO_COVER  = auto()  # Full-bleed background + title overlay
    MAP_OVERLAY  = auto()  # Large geo image + pin/label clusters
    TIMELINE     = auto()  # Center-axis alternating labels
    LOGO_ROW     = auto()  # Horizontal strip of partner logos


# ═══════════════════════════════════════════════════════════════════════════════
# Role → Transform Action mapping
# ═══════════════════════════════════════════════════════════════════════════════

# position_action: what to do with the shape's (x, y) coordinates
# text_action: what to do with text content
# direction_action: what to do with flipH/flipV and directional presets
_ROLE_ACTIONS: Dict[ShapeRole, Dict[str, str]] = {
    ShapeRole.BACKGROUND:    {'position': 'keep',       'text': 'rtl_only',      'direction': 'remove_flip'},
    ShapeRole.PANEL_LEFT:    {'position': 'swap',       'text': 'translate_rtl',  'direction': 'remove_flip'},
    ShapeRole.PANEL_RIGHT:   {'position': 'swap',       'text': 'translate_rtl',  'direction': 'remove_flip'},
    ShapeRole.BLEED:         {'position': 'mirror',     'text': 'translate_rtl',  'direction': 'remove_flip'},
    ShapeRole.PLACEHOLDER:   {'position': 'inherit',    'text': 'translate_rtl',  'direction': 'none'},
    ShapeRole.CONTENT_IMAGE: {'position': 'mirror',     'text': 'none',           'direction': 'remove_flip'},
    ShapeRole.CONTENT_TEXT:  {'position': 'mirror',     'text': 'translate_rtl',  'direction': 'remove_flip'},
    ShapeRole.TABLE:         {'position': 'mirror',     'text': 'translate_rtl',  'direction': 'none'},
    ShapeRole.CHART:         {'position': 'mirror',     'text': 'translate_rtl',  'direction': 'none'},
    ShapeRole.GROUP:         {'position': 'mirror',     'text': 'translate_rtl',  'direction': 'remove_flip'},
    ShapeRole.DECORATIVE:    {'position': 'keep',       'text': 'rtl_only',       'direction': 'remove_flip'},
    ShapeRole.LOGO:          {'position': 'mirror',     'text': 'none',           'direction': 'none'},
    ShapeRole.CONNECTOR:     {'position': 'mirror',     'text': 'none',           'direction': 'toggle_flipH'},
    ShapeRole.DIRECTIONAL:   {'position': 'mirror',     'text': 'translate_rtl',  'direction': 'swap_preset'},
    ShapeRole.OVERLAY:       {'position': 'keep',       'text': 'translate_rtl',  'direction': 'remove_flip'},
    ShapeRole.BADGE:         {'position': 'reposition', 'text': 'translate_rtl',  'direction': 'none'},
    ShapeRole.FOOTER:        {'position': 'mirror',     'text': 'translate_rtl',  'direction': 'none'},
    ShapeRole.UNKNOWN:       {'position': 'mirror',     'text': 'translate_rtl',  'direction': 'remove_flip'},
}


# ═══════════════════════════════════════════════════════════════════════════════
# Data Structures
# ═══════════════════════════════════════════════════════════════════════════════

@dataclass
class SplitPanelInfo:
    """Split-panel detection result for a slide."""
    left_shape_ids: FrozenSet[int]       # shape_id of shapes in left panel
    right_shape_ids: FrozenSet[int]      # shape_id of shapes in right panel
    left_anchor_id: int                  # shape_id of the left anchor shape
    right_anchor_id: int                 # shape_id of the right anchor shape
    left_bbox: Tuple[int, int]           # (min_x, max_x) of left panel
    right_bbox: Tuple[int, int]          # (min_x, max_x) of right panel
    shift_delta: int                     # EMU to shift for panel swap


@dataclass
class SlideContext:
    """Pre-scan results for slide-level patterns."""
    slide_width: int
    slide_height: int
    layout_type: str = 'cust'
    map_background_id: Optional[int] = None         # shape_id of map base image
    overlay_shape_ids: FrozenSet[int] = frozenset()  # shape_ids of overlay shapes
    split_panel: Optional[SplitPanelInfo] = None
    shape_count: int = 0
    slide_role: 'SlideRole' = None  # type: ignore[assignment]  # set in _pre_scan


@dataclass
class ShapeClassification:
    """
    Immutable classification result for a single shape.

    Once assigned, no transform phase may change the role.
    The position_action, text_action, and direction_action fields
    are derived from the role and pre-computed for dispatch.
    """
    role: ShapeRole
    position_action: str    # 'mirror', 'keep', 'swap', 'inherit', 'reposition'
    text_action: str        # 'translate_rtl', 'rtl_only', 'none'
    direction_action: str   # 'remove_flip', 'toggle_flipH', 'swap_preset', 'none'
    rule_name: str = ''     # Which detection rule matched (for debugging)
    confidence: float = 1.0

    # Panel-specific metadata
    panel_side: Optional[str] = None          # 'left' or 'right' for PANEL_* roles
    panel_shift_delta: int = 0                # EMU shift for panel swap

    # Bleed metadata
    bleed_left: int = 0      # how far the shape extends past left edge (positive = bleed)
    bleed_right: int = 0     # how far the shape extends past right edge

    # Placeholder metadata
    placeholder_type: Optional[str] = None    # 'title', 'body', 'ctrTitle', etc.
    placeholder_idx: Optional[int] = None

    @property
    def should_mirror(self) -> bool:
        return self.position_action == 'mirror'

    @property
    def should_swap(self) -> bool:
        return self.position_action == 'swap'

    @property
    def should_keep(self) -> bool:
        return self.position_action == 'keep'

    @property
    def should_translate(self) -> bool:
        return self.text_action in ('translate_rtl', 'rtl_only')


@dataclass
class SlideClassificationResult:
    """Complete classification for one slide."""
    slide_number: int
    layout_type: str
    classifications: Dict[int, ShapeClassification]  # shape.shape_id → classification
    context: SlideContext

    def get(self, shape) -> ShapeClassification:
        """Get classification for a shape. Returns UNKNOWN if not found."""
        return self.classifications.get(shape.shape_id, _DEFAULT_CLASSIFICATION)

    def get_by_id(self, shape_id: int) -> ShapeClassification:
        return self.classifications.get(shape_id, _DEFAULT_CLASSIFICATION)

    @property
    def has_split_panel(self) -> bool:
        return self.context.split_panel is not None

    @property
    def has_map_overlay(self) -> bool:
        return self.context.map_background_id is not None

    def shapes_with_role(self, role: ShapeRole) -> List[int]:
        """Return shape ids with the given role."""
        return [sid for sid, cls in self.classifications.items() if cls.role == role]


# Sentinel for unclassified shapes — conservative default (mirror + translate)
_DEFAULT_CLASSIFICATION = ShapeClassification(
    role=ShapeRole.UNKNOWN,
    position_action='mirror',
    text_action='translate_rtl',
    direction_action='remove_flip',
    rule_name='fallback',
    confidence=0.0,
)


# ═══════════════════════════════════════════════════════════════════════════════
# Internal shape data (read-only snapshot of original geometry)
# ═══════════════════════════════════════════════════════════════════════════════

@dataclass
class _ShapeData:
    """Read-only snapshot of shape geometry and properties at scan time."""
    shape: object           # python-pptx shape reference
    shape_id: int           # shape.shape_id (XML cNvPr id attribute)
    left: int = 0
    top: int = 0
    width: int = 0
    height: int = 0
    center_x: int = 0
    center_y: int = 0
    area: int = 0
    right: int = 0          # left + width
    bottom: int = 0         # top + height

    is_placeholder: bool = False
    placeholder_type: Optional[str] = None
    placeholder_idx: Optional[int] = None

    is_picture: bool = False      # <p:pic> element
    has_blip_fill: bool = False   # has <a:blipFill> with rEmbed
    has_text: bool = False
    text_length: int = 0

    is_group: bool = False
    is_connector: bool = False    # cxnSp tag
    is_directional: bool = False  # preset in _DIRECTIONAL_PRESETS
    preset_geometry: str = ''     # prst attribute value

    xml_tag: str = ''             # element tag (sp, pic, cxnSp, grpSp)

    # Fill type (for distinguishing solid panels vs images)
    has_solid_fill: bool = False
    has_gradient_fill: bool = False


def _make_classification(
    role: ShapeRole,
    rule_name: str = '',
    confidence: float = 1.0,
    **kwargs,
) -> ShapeClassification:
    """Create a ShapeClassification with role-derived actions."""
    actions = _ROLE_ACTIONS.get(role, _ROLE_ACTIONS[ShapeRole.UNKNOWN])
    return ShapeClassification(
        role=role,
        position_action=actions['position'],
        text_action=actions['text'],
        direction_action=actions['direction'],
        rule_name=rule_name,
        confidence=confidence,
        **kwargs,
    )


# ═══════════════════════════════════════════════════════════════════════════════
# ShapeClassifier — the main classification engine
# ═══════════════════════════════════════════════════════════════════════════════

class ShapeClassifier:
    """
    Classify all shapes on a slide ONCE, BEFORE any transforms.
    Operates on ORIGINAL geometry only — never on mutated positions.

    Usage:
        classifier = ShapeClassifier(slide_width, slide_height)
        result = classifier.classify_slide(slide, slide_number, layout_type)
        for shape in slide.shapes:
            cls = result.get(shape)
            print(f"{shape.name}: {cls.role.name} → {cls.position_action}")
    """

    def __init__(
        self,
        slide_width: int,
        slide_height: int,
        template_registry=None,
    ):
        self.slide_width = slide_width
        self.slide_height = slide_height
        self.slide_area = slide_width * slide_height
        self.half_width = slide_width // 2
        self.registry = template_registry

    # ───────────────────────────────────────────────────────────────────
    # Public API
    # ───────────────────────────────────────────────────────────────────

    def classify_slide(
        self,
        slide,
        slide_number: int,
        layout_type: str = 'cust',
    ) -> SlideClassificationResult:
        """
        Classify all shapes on a slide.

        Args:
            slide: python-pptx Slide object
            slide_number: 1-based slide index (for logging)
            layout_type: ST_SlideLayoutType string (from LayoutAnalyzer)

        Returns:
            SlideClassificationResult with per-shape classifications
        """
        # Phase 1: Collect raw shape data (read-only scan)
        all_shapes = self._collect_all_shapes(slide.shapes)
        shape_data_list = [self._extract_shape_data(s) for s in all_shapes]

        # Phase 2: Pre-scan for slide-level patterns
        context = self._pre_scan(shape_data_list, layout_type)

        # Phase 3: Classify each shape using priority-ordered rules
        classifications: Dict[int, ShapeClassification] = {}
        classified_ids: Set[int] = set()

        # 3a. Classify panel shapes (from pre-scan) first
        if context.split_panel is not None:
            for sid in context.split_panel.left_shape_ids:
                if sid not in classified_ids:
                    classifications[sid] = _make_classification(
                        ShapeRole.PANEL_LEFT,
                        rule_name='split_panel_left',
                        panel_side='left',
                        panel_shift_delta=context.split_panel.shift_delta,
                    )
                    classified_ids.add(sid)
            for sid in context.split_panel.right_shape_ids:
                if sid not in classified_ids:
                    classifications[sid] = _make_classification(
                        ShapeRole.PANEL_RIGHT,
                        rule_name='split_panel_right',
                        panel_side='right',
                        panel_shift_delta=context.split_panel.shift_delta,
                    )
                    classified_ids.add(sid)

        # 3b. Classify overlay shapes (from pre-scan)
        for sid in context.overlay_shape_ids:
            if sid not in classified_ids:
                classifications[sid] = _make_classification(
                    ShapeRole.OVERLAY,
                    rule_name='map_overlay',
                )
                classified_ids.add(sid)

        # 3c. Per-shape classification (priority-ordered rules)
        for sd in shape_data_list:
            if sd.shape_id in classified_ids:
                continue
            cls = self._classify_single_shape(sd, context)
            classifications[sd.shape_id] = cls
            classified_ids.add(sd.shape_id)

        # Phase 4: Validation
        self._validate(classifications, shape_data_list, slide_number)

        return SlideClassificationResult(
            slide_number=slide_number,
            layout_type=layout_type,
            classifications=classifications,
            context=context,
        )

    # ───────────────────────────────────────────────────────────────────
    # Phase 1: Shape data collection (read-only)
    # ───────────────────────────────────────────────────────────────────

    def _collect_all_shapes(self, shapes) -> list:
        """Flatten shape tree into a list (groups included as top-level)."""
        result = []
        for shape in shapes:
            result.append(shape)
        return result

    def _extract_shape_data(self, shape) -> _ShapeData:
        """Extract a read-only snapshot of shape properties."""
        sd = _ShapeData(shape=shape, shape_id=shape.shape_id)

        try:
            sd.left = int(getattr(shape, 'left', 0) or 0)
            sd.top = int(getattr(shape, 'top', 0) or 0)
            sd.width = int(getattr(shape, 'width', 0) or 0)
            sd.height = int(getattr(shape, 'height', 0) or 0)
        except (TypeError, ValueError):
            pass

        sd.right = sd.left + sd.width
        sd.bottom = sd.top + sd.height
        sd.center_x = sd.left + sd.width // 2
        sd.center_y = sd.top + sd.height // 2
        sd.area = sd.width * sd.height

        # Placeholder detection
        sd.is_placeholder = bool(getattr(shape, 'is_placeholder', False))
        if sd.is_placeholder:
            try:
                ph = shape.placeholder_format
                sd.placeholder_type = str(ph.type) if ph else None
                sd.placeholder_idx = ph.idx if ph else None
            except Exception:
                pass
            # Also try XML-based detection for more detail
            try:
                ph_info = get_placeholder_info_from_xml(shape._element)
                if ph_info:
                    sd.placeholder_type = ph_info[0] or sd.placeholder_type
                    sd.placeholder_idx = ph_info[1] if ph_info[1] is not None else sd.placeholder_idx
            except Exception:
                pass

        # XML tag analysis
        try:
            sp_el = shape._element
            raw_tag = sp_el.tag
            sd.xml_tag = raw_tag.split('}')[-1] if '}' in raw_tag else raw_tag

            # Picture detection
            sd.is_picture = sd.xml_tag == 'pic'

            # BlipFill detection (image content)
            blip_fill = sp_el.find(f'.//{{{A_NS}}}blipFill')
            if blip_fill is not None:
                blip = blip_fill.find(f'{{{A_NS}}}blip')
                if blip is not None and blip.get(f'{{{R_NS}}}embed'):
                    sd.has_blip_fill = True

            # Connector detection
            sd.is_connector = sd.xml_tag == 'cxnSp'

            # Group detection
            sd.is_group = sd.xml_tag == 'grpSp' or hasattr(shape, 'shapes')

            # Preset geometry detection
            prst_geom = sp_el.find(f'.//{{{A_NS}}}prstGeom')
            if prst_geom is not None:
                sd.preset_geometry = prst_geom.get('prst', '')

            # Line preset → connector
            if not sd.is_connector and sd.preset_geometry in _LINE_PRESETS:
                sd.is_connector = True
            if 'Connector' in sd.preset_geometry:
                sd.is_connector = True

            # Directional preset
            if sd.preset_geometry in _DIRECTIONAL_PRESETS:
                sd.is_directional = True

            # Zero-dimension shapes (perfectly straight lines)
            if not sd.is_connector and (sd.width == 0 or sd.height == 0):
                has_text = getattr(shape, 'has_text_frame', False)
                if not has_text:
                    sd.is_connector = True

            # High aspect ratio line detection (no text, aspect > 20, minor dim < 50000)
            if not sd.is_connector and sd.width > 0 and sd.height > 0:
                has_text = getattr(shape, 'has_text_frame', False)
                if not has_text:
                    major = max(sd.width, sd.height)
                    minor = min(sd.width, sd.height)
                    if minor > 0 and major / minor > 20 and minor < 50000:
                        sd.is_connector = True

            # Fill type detection
            sp_pr = sp_el.find(f'{{{P_NS}}}spPr')
            if sp_pr is None:
                sp_pr = sp_el.find(f'{{{A_NS}}}spPr')
            if sp_pr is not None:
                if sp_pr.find(f'{{{A_NS}}}solidFill') is not None:
                    sd.has_solid_fill = True
                if sp_pr.find(f'{{{A_NS}}}gradFill') is not None:
                    sd.has_gradient_fill = True

        except Exception as exc:
            logger.debug('_extract_shape_data: %s', exc)

        # Text detection
        try:
            if getattr(shape, 'has_text_frame', False) and shape.has_text_frame:
                text = shape.text_frame.text or ''
                sd.has_text = bool(text.strip())
                sd.text_length = len(text.strip())
        except Exception:
            pass

        return sd

    # ───────────────────────────────────────────────────────────────────
    # Phase 2: Pre-scan (slide-level pattern detection)
    # ───────────────────────────────────────────────────────────────────

    def _pre_scan(self, shapes: List[_ShapeData], layout_type: str) -> SlideContext:
        """Detect slide-level patterns before per-shape classification."""
        ctx = SlideContext(
            slide_width=self.slide_width,
            slide_height=self.slide_height,
            layout_type=layout_type,
            shape_count=len(shapes),
        )

        # Detect backgrounds first (needed to exclude from panel/overlay detection)
        background_ids = self._detect_backgrounds(shapes)

        # Detect map overlay (large image + small overlay shapes)
        map_bg_id, overlay_ids = self._detect_map_overlay(shapes, background_ids)
        ctx.map_background_id = map_bg_id
        ctx.overlay_shape_ids = frozenset(overlay_ids)

        # Detect split-panel layout (two large shapes on opposite halves)
        # Skip if map overlay was detected (map slides are not split-panel)
        if map_bg_id is None:
            panel_info = self._detect_split_panel(shapes, background_ids)
            ctx.split_panel = panel_info

        # ── Slide-role detection ──
        ctx.slide_role = self._detect_slide_role(shapes, ctx, background_ids)

        return ctx

    def _detect_slide_role(
        self,
        shapes: List['_ShapeData'],
        ctx: SlideContext,
        background_ids: Set[int],
    ) -> SlideRole:
        """Classify the slide into a SlideRole based on detected patterns."""

        # SPLIT_PANEL: already detected split panel
        if ctx.split_panel is not None:
            return SlideRole.SPLIT_PANEL

        # MAP_OVERLAY: already detected map overlay
        if ctx.map_background_id is not None:
            return SlideRole.MAP_OVERLAY

        non_bg = [s for s in shapes if s.shape_id not in background_ids]
        bg_shapes = [s for s in shapes if s.shape_id in background_ids]

        # PHOTO_COVER: background image + at most 2 text placeholders
        if bg_shapes:
            has_bg_image = any(
                s.has_blip_fill or s.is_picture for s in bg_shapes
            )
            text_phs = [
                s for s in non_bg
                if s.placeholder_type is not None and s.text_length > 0
            ]
            non_text_non_bg = [
                s for s in non_bg
                if s.placeholder_type is None or s.text_length == 0
            ]
            if has_bg_image and len(text_phs) <= 2 and len(non_text_non_bg) == 0:
                return SlideRole.PHOTO_COVER

        # LOGO_ROW: multiple small images in horizontal strip
        # (<20% slide height each, within 10% Y of each other)
        small_images = [
            s for s in non_bg
            if (s.has_blip_fill or s.is_picture)
            and s.height < self.slide_height * 0.20
            and s.width < self.slide_width * 0.30
        ]
        if len(small_images) >= 3:
            y_positions = [s.top for s in small_images]
            y_range = max(y_positions) - min(y_positions)
            if y_range < self.slide_height * 0.10:
                return SlideRole.LOGO_ROW

        # TIMELINE: center-axis vertical line + alternating text labels
        # Look for a thin vertical shape near the center with text shapes
        # alternating on left/right
        center_x = self.slide_width // 2
        center_tolerance = self.slide_width * 0.10
        center_lines = [
            s for s in non_bg
            if s.width < self.slide_width * 0.05  # thin
            and s.height > self.slide_height * 0.40  # tall
            and abs((s.left + s.width // 2) - center_x) < center_tolerance
        ]
        if center_lines:
            text_shapes = [
                s for s in non_bg
                if s.text_length > 0 and s.shape_id not in {c.shape_id for c in center_lines}
            ]
            if len(text_shapes) >= 4:
                left_count = sum(
                    1 for s in text_shapes
                    if s.left + s.width // 2 < center_x
                )
                right_count = len(text_shapes) - left_count
                if left_count >= 2 and right_count >= 2:
                    return SlideRole.TIMELINE

        return SlideRole.STANDARD

    def _detect_backgrounds(self, shapes: List[_ShapeData]) -> Set[int]:
        """
        Detect full-slide background shapes.

        Rules:
        - width >= 95% slide width AND height >= 95% slide height (definite)
        - OR: width >= 90% slide width (v1 compatibility)
        - OR: area >= 85% slide area AND center near slide center

        Returns set of shape_ids that are backgrounds.
        """
        bg_ids: Set[int] = set()
        for sd in shapes:
            # Rule 1: Full coverage (width AND height >= 95%)
            if (sd.width >= self.slide_width * _BG_FULL_COVER_FRACTION
                    and sd.height >= self.slide_height * _BG_FULL_COVER_FRACTION):
                bg_ids.add(sd.shape_id)
                continue

            # Rule 2: Width-dominant background (>= 90% width) — v1 compat
            if sd.width >= self.slide_width * _BG_WIDTH_FRACTION:
                bg_ids.add(sd.shape_id)
                continue

            # Rule 3: Area-based (>= 85% area, center near slide center)
            if sd.area >= self.slide_area * _BG_AREA_FRACTION:
                slide_cx = self.slide_width // 2
                slide_cy = self.slide_height // 2
                # Center must be within 5% of slide center
                if (abs(sd.center_x - slide_cx) < self.slide_width * 0.05
                        and abs(sd.center_y - slide_cy) < self.slide_height * 0.05):
                    bg_ids.add(sd.shape_id)

        return bg_ids

    def _detect_map_overlay(
        self,
        shapes: List[_ShapeData],
        background_ids: Set[int],
    ) -> Tuple[Optional[int], Set[int]]:
        """
        Detect map/geographic overlay pattern.

        Algorithm:
        1. Find the largest image where area > 40% and width > 60%
        2. Count small shapes (< 30% width) whose center is inside the image
        3. If >= 3 such shapes, classify them as overlays

        Returns (map_bg_shape_id, overlay_shape_ids)
        """
        # Find candidate base image
        base_image: Optional[_ShapeData] = None
        best_area = 0

        for sd in shapes:
            if sd.shape_id in background_ids:
                continue
            if not (sd.is_picture or sd.has_blip_fill):
                continue
            if sd.area < self.slide_area * _MAP_AREA_FRACTION:
                continue
            if sd.width < self.slide_width * _MAP_WIDTH_FRACTION:
                continue
            if sd.area > best_area:
                best_area = sd.area
                base_image = sd

        if base_image is None:
            return None, set()

        # Find small shapes inside the base image
        candidates: List[int] = []
        for sd in shapes:
            if sd.shape_id == base_image.shape_id:
                continue
            if sd.shape_id in background_ids:
                continue
            if sd.is_placeholder:
                continue
            if sd.width >= self.slide_width * _OVERLAY_MAX_WIDTH_FRACTION:
                continue
            # Center must be inside base image bounding box
            if (base_image.left <= sd.center_x <= base_image.right
                    and base_image.top <= sd.center_y <= base_image.bottom):
                candidates.append(sd.shape_id)

        if len(candidates) >= _OVERLAY_MIN_COUNT:
            return base_image.shape_id, set(candidates)

        return None, set()

    def _detect_split_panel(
        self,
        shapes: List[_ShapeData],
        background_ids: Set[int],
    ) -> Optional[SplitPanelInfo]:
        """
        Detect split-panel layout (image on one side, content on other).

        Algorithm:
        1. Find anchor shapes: width >= 35% AND height >= 50%
        2. Pick best left anchor (center_x < midpoint) and right anchor
        3. Verify asymmetry: one is image, the other is not
           - ALSO allow same-type pairs (two groups) per Gemini rule
        4. Classify all non-bg, non-full-width shapes into panels by center_x

        Returns SplitPanelInfo or None.
        """
        # Step 1: Find anchor candidates (large shapes, excluding backgrounds)
        left_anchor: Optional[_ShapeData] = None
        right_anchor: Optional[_ShapeData] = None

        for sd in shapes:
            if sd.shape_id in background_ids:
                continue
            # Include placeholders in anchor detection (v2 fix for v1 bug)
            if sd.width < self.slide_width * _PANEL_MIN_WIDTH_FRACTION:
                continue
            if sd.height < self.slide_height * _PANEL_MIN_HEIGHT_FRACTION:
                continue

            # Use visible center for shapes with negative positions (bleed fix)
            visible_left = max(0, sd.left)
            visible_right = min(self.slide_width, sd.right)
            visible_center_x = (visible_left + visible_right) // 2

            if visible_center_x < self.half_width:
                if left_anchor is None or sd.width > left_anchor.width:
                    left_anchor = sd
            else:
                if right_anchor is None or sd.width > right_anchor.width:
                    right_anchor = sd

        # Need anchors on both sides
        if left_anchor is None or right_anchor is None:
            return None

        # Step 2: Verify asymmetry (image vs non-image)
        # Allow both: (image vs non-image) AND (same-type like two groups)
        left_is_img = left_anchor.is_picture or left_anchor.has_blip_fill
        right_is_img = right_anchor.is_picture or right_anchor.has_blip_fill

        # Both images with no other distinguishing features → not a panel
        # Both non-images are OK (e.g., two groups in Lukas slide 8)
        if left_is_img and right_is_img:
            # Both are images — only proceed if they're clearly different
            # (one much larger than other, or different shape types)
            if abs(left_anchor.area - right_anchor.area) < self.slide_area * 0.10:
                return None  # Similar-sized images — not a split panel

        # Step 3: Classify all shapes into left/right panels
        left_ids: Set[int] = set()
        right_ids: Set[int] = set()

        for sd in shapes:
            if sd.shape_id in background_ids:
                continue
            # Skip full-width shapes (backgrounds/separators)
            if sd.width >= self.slide_width * _FULL_WIDTH_FRACTION:
                continue

            visible_left = max(0, sd.left)
            visible_right = min(self.slide_width, sd.right)
            visible_center_x = (visible_left + visible_right) // 2

            if visible_center_x < self.half_width:
                left_ids.add(sd.shape_id)
            else:
                right_ids.add(sd.shape_id)

        if not left_ids or not right_ids:
            return None

        # Step 4: Compute panel bounding boxes and shift delta
        left_min_x = min(
            sd.left for sd in shapes if sd.shape_id in left_ids
        )
        left_max_x = max(
            sd.right for sd in shapes if sd.shape_id in left_ids
        )
        right_min_x = min(
            sd.left for sd in shapes if sd.shape_id in right_ids
        )
        right_max_x = max(
            sd.right for sd in shapes if sd.shape_id in right_ids
        )

        shift_delta = right_min_x - left_min_x

        return SplitPanelInfo(
            left_shape_ids=frozenset(left_ids),
            right_shape_ids=frozenset(right_ids),
            left_anchor_id=left_anchor.shape_id,
            right_anchor_id=right_anchor.shape_id,
            left_bbox=(left_min_x, left_max_x),
            right_bbox=(right_min_x, right_max_x),
            shift_delta=shift_delta,
        )

    # ───────────────────────────────────────────────────────────────────
    # Phase 3: Per-shape classification (priority-ordered rules)
    # ───────────────────────────────────────────────────────────────────

    def _classify_single_shape(
        self,
        sd: _ShapeData,
        ctx: SlideContext,
    ) -> ShapeClassification:
        """
        Classify a single shape using priority-ordered rules.
        The FIRST matching rule wins — no further rules are evaluated.

        Priority order:
         1. PLACEHOLDER  (structural — XML <p:ph> element)
         2. TABLE        (structural — python-pptx API)
         3. CHART        (structural — python-pptx API)
         4. CONNECTOR    (structural — XML tag / preset)
         5. DIRECTIONAL  (structural — preset geometry)
         6. BACKGROUND   (geometric — covers full slide)
         7. BLEED        (geometric — extends past edges)
         8. FOOTER       (positional — bottom zone)
         9. BADGE        (positional — small, top, numeric)
        10. LOGO         (structural — small image)
        11. OVERLAY      (contextual — handled in pre-scan, but catch stragglers)
        12. DECORATIVE   (layout-dependent — title/secHead decoratives)
        13. GROUP        (structural — GroupShape)
        14. CONTENT_IMAGE (structural — pic/blipFill)
        15. CONTENT_TEXT  (has text content)
        16. UNKNOWN       (fallback)
        """

        # ── Priority 1: PLACEHOLDER ──
        if sd.is_placeholder:
            return _make_classification(
                ShapeRole.PLACEHOLDER,
                rule_name='placeholder_ph_element',
                placeholder_type=sd.placeholder_type,
                placeholder_idx=sd.placeholder_idx,
            )

        # ── Priority 2: TABLE ──
        if getattr(sd.shape, 'has_table', False) and sd.shape.has_table:
            return _make_classification(ShapeRole.TABLE, rule_name='has_table')

        # ── Priority 3: CHART ──
        if getattr(sd.shape, 'has_chart', False) and sd.shape.has_chart:
            return _make_classification(ShapeRole.CHART, rule_name='has_chart')

        # ── Priority 4: CONNECTOR ──
        if sd.is_connector:
            return _make_classification(
                ShapeRole.CONNECTOR,
                rule_name=f'connector_{sd.xml_tag}_{sd.preset_geometry}',
            )

        # ── Priority 5: DIRECTIONAL ──
        if sd.is_directional:
            return _make_classification(
                ShapeRole.DIRECTIONAL,
                rule_name=f'directional_{sd.preset_geometry}',
            )

        # ── Priority 6: BACKGROUND ──
        # Full-width (>= 90%) OR full coverage (>= 95% both dims)
        if sd.width >= self.slide_width * _BG_WIDTH_FRACTION:
            return _make_classification(
                ShapeRole.BACKGROUND,
                rule_name='bg_full_width',
            )
        if sd.area >= self.slide_area * _BG_AREA_FRACTION:
            slide_cx = self.slide_width // 2
            slide_cy = self.slide_height // 2
            if (abs(sd.center_x - slide_cx) < self.slide_width * 0.05
                    and abs(sd.center_y - slide_cy) < self.slide_height * 0.05):
                return _make_classification(
                    ShapeRole.BACKGROUND,
                    rule_name='bg_area_centered',
                )

        # ── Priority 7: BLEED ──
        bleed_left = max(0, -sd.left)
        bleed_right = max(0, sd.right - self.slide_width)
        if bleed_left > _BLEED_THRESHOLD_EMU or bleed_right > _BLEED_THRESHOLD_EMU:
            # Verify intentionality: visible area >= 10% of total area
            visible_left = max(0, sd.left)
            visible_right = min(self.slide_width, sd.right)
            visible_top = max(0, sd.top)
            visible_bottom = min(self.slide_height, sd.bottom)
            visible_area = max(0, visible_right - visible_left) * max(0, visible_bottom - visible_top)
            if sd.area > 0 and visible_area >= sd.area * 0.10:
                return _make_classification(
                    ShapeRole.BLEED,
                    rule_name='bleed_extends_past_edge',
                    bleed_left=bleed_left,
                    bleed_right=bleed_right,
                )

        # ── Priority 8: FOOTER ──
        if sd.center_y > self.slide_height * _FOOTER_ZONE_FRACTION:
            # Footer zone — small shapes in the bottom strip
            if sd.height < self.slide_height * 0.08:
                return _make_classification(
                    ShapeRole.FOOTER,
                    rule_name='footer_zone_small',
                )

        # ── Priority 9: BADGE ──
        if (sd.width < self.slide_width * _BADGE_MAX_DIM_FRACTION
                and sd.height < self.slide_height * _BADGE_MAX_DIM_FRACTION
                and sd.top < self.slide_height * _BADGE_MAX_TOP_FRACTION):
            # Check if it contains short numeric text (1-4 digits)
            if sd.has_text and sd.text_length <= 4:
                try:
                    text = sd.shape.text_frame.text.strip()
                    if text.isdigit():
                        return _make_classification(
                            ShapeRole.BADGE,
                            rule_name='badge_numeric_small',
                        )
                except Exception:
                    pass

        # ── Priority 10: LOGO ──
        if (sd.is_picture or sd.has_blip_fill) and not sd.has_text:
            if sd.width < self.slide_width * _LOGO_MAX_WIDTH_FRACTION:
                return _make_classification(
                    ShapeRole.LOGO,
                    rule_name='logo_small_image',
                )

        # ── Priority 11: DECORATIVE (layout-dependent) ──
        if ctx.layout_type in _CENTERED_LAYOUTS:
            # On title/secHead layouts, small shapes are decorative
            if (sd.width < self.slide_width * _DECORATIVE_MAX_WIDTH_FRACTION
                    or sd.height < self.slide_height * _DECORATIVE_MAX_HEIGHT_FRACTION):
                if not sd.has_blip_fill and not sd.is_picture:
                    return _make_classification(
                        ShapeRole.DECORATIVE,
                        rule_name='decorative_title_layout_small',
                    )
            # Content-sized shapes on title layouts get mirrored
            is_content_sized = (
                sd.width > self.slide_width * _CONTENT_MIN_WIDTH_FRACTION
                and sd.height > self.slide_height * _CONTENT_MIN_HEIGHT_FRACTION
            )
            if not is_content_sized and not sd.has_blip_fill and not sd.is_picture:
                # Small non-image shape on title layout → decorative
                return _make_classification(
                    ShapeRole.DECORATIVE,
                    rule_name='decorative_title_layout_noncontent',
                    confidence=0.7,
                )

        # ── Priority 12: GROUP ──
        if sd.is_group:
            return _make_classification(
                ShapeRole.GROUP,
                rule_name='group_shape',
            )

        # ── Priority 13: CONTENT_IMAGE ──
        if sd.is_picture or sd.has_blip_fill:
            return _make_classification(
                ShapeRole.CONTENT_IMAGE,
                rule_name='content_image',
            )

        # ── Priority 14: CONTENT_TEXT ──
        if sd.has_text:
            return _make_classification(
                ShapeRole.CONTENT_TEXT,
                rule_name='content_text',
            )

        # ── Priority 15: UNKNOWN ──
        return _make_classification(
            ShapeRole.UNKNOWN,
            rule_name='no_rule_matched',
            confidence=0.0,
        )

    # ───────────────────────────────────────────────────────────────────
    # Phase 4: Validation
    # ───────────────────────────────────────────────────────────────────

    def _validate(
        self,
        classifications: Dict[int, ShapeClassification],
        shapes: List[_ShapeData],
        slide_number: int,
    ) -> None:
        """Assert every shape is classified and log diagnostics."""
        for sd in shapes:
            if sd.shape_id not in classifications:
                logger.warning(
                    'Slide %d: shape "%s" (id=%d) not classified — assigning UNKNOWN',
                    slide_number,
                    getattr(sd.shape, 'name', '?'),
                    sd.shape_id,
                )
                classifications[sd.shape_id] = _DEFAULT_CLASSIFICATION

        # Log classification summary
        role_counts: Dict[str, int] = {}
        for cls in classifications.values():
            name = cls.role.name
            role_counts[name] = role_counts.get(name, 0) + 1

        logger.debug(
            'Slide %d: classified %d shapes — %s',
            slide_number,
            len(classifications),
            ', '.join(f'{k}={v}' for k, v in sorted(role_counts.items())),
        )


# ═══════════════════════════════════════════════════════════════════════════════
# Module-level convenience function
# ═══════════════════════════════════════════════════════════════════════════════

def classify_slide(
    slide,
    slide_number: int,
    slide_width: int,
    slide_height: int,
    layout_type: str = 'cust',
    template_registry=None,
) -> SlideClassificationResult:
    """
    Convenience function: classify all shapes on a slide.

    Args:
        slide: python-pptx Slide object
        slide_number: 1-based slide index
        slide_width: Slide width in EMU
        slide_height: Slide height in EMU
        layout_type: ST_SlideLayoutType string
        template_registry: Optional TemplateRegistry

    Returns:
        SlideClassificationResult
    """
    classifier = ShapeClassifier(slide_width, slide_height, template_registry)
    return classifier.classify_slide(slide, slide_number, layout_type)
