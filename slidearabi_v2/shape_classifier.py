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
_CENTERLINE_EXCLUSION_FRACTION = 0.05  # shapes within 5% of midpoint → not panel-assigned

# Asymmetric panel detection (one image anchor + text cluster on opposite side)
_ASYM_PANEL_MIN_WIDTH = 0.30       # relaxed anchor width for asymmetric detection
_ASYM_PANEL_IMAGE_AREA = 0.20      # image area >= 20% of slide → panel image
_ASYM_PANEL_IMAGE_WIDTH = 0.40     # image width >= 40% of slide → panel image (landscape)

# Strategy 2b: Dominant image + single text shape on opposite side
_DOMINANT_IMAGE_MIN_WIDTH = 0.40     # image must be >= 40% slide width
_DOMINANT_IMAGE_MIN_HEIGHT = 0.80    # image must be >= 80% slide height
_OPPOSITE_TEXT_MIN_WIDTH = 0.25      # text shape must be >= 25% slide width

# Cluster-based panel detection (no single anchor, but clear spatial grouping)
_CLUSTER_MIN_WIDTH_FRACTION = 0.25   # cluster bbox width >= 25% of slide
_CLUSTER_MIN_HEIGHT_FRACTION = 0.30  # cluster bbox height >= 30% of slide
_CLUSTER_MIN_GAP_FRACTION = 0.05     # horizontal gap >= 5% of slide between clusters

# Logo detection
_LOGO_MAX_WIDTH_FRACTION = 0.20    # image < 20% slide width = logo

# Bleed detection
_BLEED_THRESHOLD_EMU = 100_000     # ~0.11" — minimum bleed to classify as BLEED
_MAX_BLEED_EMU = 1_500_000         # ~1.64" — maximum allowed intentional bleed

# Overlay/map detection
_MAP_AREA_FRACTION = 0.35          # base image must cover >= 35% of slide area
_MAP_WIDTH_FRACTION = 0.55         # base image must be >= 55% of slide width
_OVERLAY_MAX_WIDTH_FRACTION = 0.30 # overlay shapes must be < 30% slide width
_OVERLAY_MAX_WIDTH_GROUP = 0.40    # GROUP shapes allowed up to 40% (larger bboxes)
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
    COMPLEX_GRAPHIC  = auto()  # Multi-child infographic — translate only, preserve position
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
    ShapeRole.COMPLEX_GRAPHIC: {'position': 'keep',      'text': 'translate_rtl',  'direction': 'none'},
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
    dominant_image: bool = False          # True if detected via Strategy 2b (single text + dominant image)


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

        # 3a. Classify overlay shapes FIRST (geographically anchored —
        #     must not be re-assigned as panel shapes)
        for sid in context.overlay_shape_ids:
            if sid not in classified_ids:
                classifications[sid] = _make_classification(
                    ShapeRole.OVERLAY,
                    rule_name='overlay_prescan',
                )
                classified_ids.add(sid)

        # 3b. Classify panel shapes (from pre-scan)
        # Footer/title placeholders must NOT become panel shapes — they are
        # structural elements that stay in their original positions.
        _NON_PANEL_PH_KEYWORDS = frozenset({
            'ftr', 'sldnum', 'dt', 'footer', 'slidenumber', 'date_time',
            'title', 'ctrtitle',
        })

        def _is_non_panel_placeholder(shape_id: int) -> bool:
            """Check if shape is a footer/title placeholder that must not become a panel."""
            sd = _sd_by_id.get(shape_id)
            if sd is None or not sd.is_placeholder:
                return False
            ph_type = (sd.placeholder_type or '').lower()
            return any(k in ph_type for k in _NON_PANEL_PH_KEYWORDS)

        _sd_by_id = {sd.shape_id: sd for sd in shape_data_list}

        if context.split_panel is not None:
            # For dominant-image panels (Strategy 2b), the single text shape
            # IS the panel content — don't exclude it even if it's a title placeholder.
            skip_ph_exclusion = getattr(context.split_panel, 'dominant_image', False)
            for sid in context.split_panel.left_shape_ids:
                if sid in classified_ids:
                    continue
                if not skip_ph_exclusion and _is_non_panel_placeholder(sid):
                    continue
                classifications[sid] = _make_classification(
                    ShapeRole.PANEL_LEFT,
                    rule_name='split_panel_left',
                    panel_side='left',
                    panel_shift_delta=context.split_panel.shift_delta,
                )
                classified_ids.add(sid)
            for sid in context.split_panel.right_shape_ids:
                if sid in classified_ids:
                    continue
                if not skip_ph_exclusion and _is_non_panel_placeholder(sid):
                    continue
                classifications[sid] = _make_classification(
                    ShapeRole.PANEL_RIGHT,
                    rule_name='split_panel_right',
                    panel_side='right',
                    panel_shift_delta=context.split_panel.shift_delta,
                )
                classified_ids.add(sid)

        # 3c. Per-shape classification (priority-ordered rules)
        for sd in shape_data_list:
            if sd.shape_id in classified_ids:
                continue
            cls = self._classify_single_shape(sd, context)

            # MAP_OVERLAY slide-level override: translate text ONLY.
            # Geographic annotations must not move, flip, or change direction.
            # Only structural placeholders (title/footer) get normal treatment.
            if (context.slide_role == SlideRole.MAP_OVERLAY
                    and cls.role not in (
                        ShapeRole.PLACEHOLDER, ShapeRole.FOOTER,
                    )):
                cls = ShapeClassification(
                    role=cls.role,
                    position_action='keep',        # Don't move
                    text_action=cls.text_action,   # Still translate
                    direction_action='none',        # Don't flip/rotate
                    rule_name=cls.rule_name + '_map_translate_only',
                    confidence=cls.confidence,
                    placeholder_type=cls.placeholder_type,
                    placeholder_idx=cls.placeholder_idx,
                )

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

        # Detect pie/donut label overlays (text shapes over pie/donut charts)
        pie_label_ids = self._detect_pie_donut_labels(shapes, background_ids)

        # Merge both overlay sets
        ctx.overlay_shape_ids = frozenset(overlay_ids | pie_label_ids)

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
        # NOTE: Do NOT skip background_ids — map images often cover >90% width
        # and get classified as backgrounds. They can still be the map base.
        base_image: Optional[_ShapeData] = None
        best_area = 0

        for sd in shapes:
            # Allow background-classified images as map base
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

        # Find small shapes inside the base image (with vertical extension
        # for icon/label rows placed just below the map image)
        bottom_extension = int(self.slide_height * 0.10)
        candidates: List[int] = []
        for sd in shapes:
            if sd.shape_id == base_image.shape_id:
                continue
            if sd.shape_id in background_ids:
                continue
            if sd.is_placeholder:
                continue
            # GROUP shapes get a wider threshold (their bboxes are larger)
            max_w = _OVERLAY_MAX_WIDTH_GROUP if sd.is_group else _OVERLAY_MAX_WIDTH_FRACTION
            if sd.width >= self.slide_width * max_w:
                continue
            # Center must be inside base image bounding box
            # (with 10% vertical extension below for legend rows)
            if (base_image.left <= sd.center_x <= base_image.right
                    and base_image.top <= sd.center_y <= base_image.bottom + bottom_extension):
                candidates.append(sd.shape_id)

        if len(candidates) >= _OVERLAY_MIN_COUNT:
            return base_image.shape_id, set(candidates)

        # Fallback: for wide images (aspect ratio >= 1.5, likely maps),
        # accept with 2 candidates instead of 3
        if base_image.height > 0:
            aspect = base_image.width / base_image.height
            if aspect >= 1.5 and len(candidates) >= 2:
                return base_image.shape_id, set(candidates)

        return None, set()

    # Pie/donut chart types (no axis, rotationally symmetric)
    _PIE_DONUT_TYPES = frozenset({
        'pieChart', 'pie3DChart', 'doughnutChart', 'ofPieChart',
    })

    # Axis-based chart types (used to exclude combo charts)
    _AXIS_CHART_TYPES = frozenset({
        'barChart', 'bar3DChart', 'lineChart', 'line3DChart',
        'areaChart', 'area3DChart', 'scatterChart', 'radarChart',
        'stockChart', 'surfaceChart', 'surface3DChart', 'bubbleChart',
    })

    def _detect_pie_donut_labels(
        self,
        shapes: List[_ShapeData],
        background_ids: Set[int],
    ) -> Set[int]:
        """
        Detect text shapes overlaid on pie/donut charts.

        Pie/donut charts are rotationally symmetric — their graphic stays
        unchanged during RTL transform.  External text labels (%, category)
        that are positioned over the chart must NOT be position-mirrored,
        or they will point to the wrong segments.

        Returns set of shape_ids for text shapes that are label overlays.
        """
        c_ns = 'http://schemas.openxmlformats.org/drawingml/2006/chart'

        # Find pie/donut-only chart shapes and their bounding boxes
        pie_chart_rects: List[Tuple[int, int, int, int]] = []  # (left, top, right, bottom)
        for sd in shapes:
            if not (getattr(sd.shape, 'has_chart', False) and sd.shape.has_chart):
                continue
            try:
                chart_elem = sd.shape.chart._part._element
                has_pie = any(
                    chart_elem.find(f'.//{{{c_ns}}}{t}') is not None
                    for t in self._PIE_DONUT_TYPES
                )
                has_axis = any(
                    chart_elem.find(f'.//{{{c_ns}}}{t}') is not None
                    for t in self._AXIS_CHART_TYPES
                )
                if has_pie and not has_axis:
                    pie_chart_rects.append((sd.left, sd.top, sd.right, sd.bottom))
            except Exception:
                continue

        if not pie_chart_rects:
            return set()

        # Find text/small shapes whose center falls inside any pie/donut chart
        # (with expanded bounding box to catch callout-style external labels).
        # Pie/donut labels often sit WELL outside the chart frame with
        # connector lines pointing inward — 5% margin is not enough.
        # Use 25% horizontal margin to catch these (Round 11 fix).
        _structural_ph_types = frozenset({
            'ftr', 'sldnum', 'dt', 'footer', 'slidenumber',
            'date_time', 'title', 'ctrtitle',
        })
        label_ids: Set[int] = set()
        margin_x = int(self.slide_width * 0.25)
        margin_y = int(self.slide_height * 0.15)
        for sd in shapes:
            if sd.shape_id in background_ids:
                continue
            # Skip structural placeholders (footer/title)
            if sd.is_placeholder:
                ph_type = (sd.placeholder_type or '').lower()
                if any(t in ph_type for t in _structural_ph_types):
                    continue
            # Include text shapes AND small non-text shapes (callouts)
            if not sd.has_text and sd.width > self.slide_width * 0.15:
                continue
            # Skip wide shapes (titles, subtitles) — pie labels are small
            if sd.width > self.slide_width * 0.20:
                continue
            for (c_left, c_top, c_right, c_bottom) in pie_chart_rects:
                if (c_left - margin_x <= sd.center_x <= c_right + margin_x
                        and c_top - margin_y <= sd.center_y <= c_bottom + margin_y):
                    label_ids.add(sd.shape_id)
                    break

        if label_ids:
            logger.debug(
                'Pie/donut overlay detection: found %d label shapes over %d pie charts',
                len(label_ids), len(pie_chart_rects),
            )
        return label_ids

    def _detect_split_panel(
        self,
        shapes: List[_ShapeData],
        background_ids: Set[int],
    ) -> Optional[SplitPanelInfo]:
        """
        Detect split-panel layout (image on one side, content on other).

        Detection strategy (cascading, ported from v1):
        1. Anchor-based: large shape (>=35% W, >=50% H) on each side,
           with image/non-image asymmetry.
        2. Asymmetric: ONE side has a large image anchor, the other has
           >=2 non-image shapes (no anchor needed on text side).
        2b. Dominant image: ONE side has a near-full-height image
            (>=40% W, >=80% H), the other has exactly 1 non-image shape
            (>=25% W). Handles slides with a single text block opposite
            a dominant photo (e.g., Lukas S1).
        3. Cluster-based: shapes form two spatially distinct groups on
           opposite sides with a clear gap and image asymmetry.

        Returns SplitPanelInfo or None.
        """
        # ── Partition shapes into left/right by center_x ──────────────
        left_shapes: List[_ShapeData] = []
        right_shapes: List[_ShapeData] = []

        for sd in shapes:
            if sd.shape_id in background_ids:
                continue
            if sd.width >= self.slide_width * _FULL_WIDTH_FRACTION:
                continue

            visible_left = max(0, sd.left)
            visible_right = min(self.slide_width, sd.right)
            visible_center_x = (visible_left + visible_right) // 2

            if visible_center_x < self.half_width:
                left_shapes.append(sd)
            else:
                right_shapes.append(sd)

        if not left_shapes or not right_shapes:
            return None

        def _is_image(sd: _ShapeData) -> bool:
            return sd.is_picture or sd.has_blip_fill

        def _is_panel_image(sd: _ShapeData) -> bool:
            """Size-weighted: only large images qualify as panel anchors."""
            if not _is_image(sd):
                return False
            if self.slide_width == 0 or self.slide_height == 0:
                return False
            wr = sd.width / self.slide_width
            hr = sd.height / self.slide_height
            ar = sd.area / self.slide_area if self.slide_area > 0 else 0
            return (
                (wr > 0.25 and hr > 0.50)
                or ar > _ASYM_PANEL_IMAGE_AREA
                or wr > _ASYM_PANEL_IMAGE_WIDTH
            )

        # ── Find anchor candidates (large shapes) ────────────────────
        left_anchor: Optional[_ShapeData] = None
        right_anchor: Optional[_ShapeData] = None

        for sd in shapes:
            if sd.shape_id in background_ids:
                continue
            if sd.width < self.slide_width * _PANEL_MIN_WIDTH_FRACTION:
                continue
            if sd.height < self.slide_height * _PANEL_MIN_HEIGHT_FRACTION:
                continue

            visible_left = max(0, sd.left)
            visible_right = min(self.slide_width, sd.right)
            visible_center_x = (visible_left + visible_right) // 2

            if visible_center_x < self.half_width:
                if left_anchor is None or sd.width > left_anchor.width:
                    left_anchor = sd
            else:
                if right_anchor is None or sd.width > right_anchor.width:
                    right_anchor = sd

        # ── Strategy 1: Anchor-based (both sides have large shapes) ───
        anchor_detected = False
        if left_anchor is not None and right_anchor is not None:
            left_is_img = _is_image(left_anchor)
            right_is_img = _is_image(right_anchor)
            if left_is_img != right_is_img:
                anchor_detected = True
            elif not left_is_img and not right_is_img:
                # Both non-images OK (e.g., two groups in Lukas slide 8)
                anchor_detected = True
            elif left_is_img and right_is_img:
                # Both images — only if clearly different sizes
                if abs(left_anchor.area - right_anchor.area) >= self.slide_area * 0.10:
                    anchor_detected = True

        # ── Strategy 2: Asymmetric (one image anchor + text cluster) ──
        asymmetric_detected = False
        if not anchor_detected:
            # Exactly one anchor exists and it is an image
            if (left_anchor is not None) != (right_anchor is not None):
                anchor = left_anchor if left_anchor is not None else right_anchor
                if _is_image(anchor):
                    if anchor is left_anchor:
                        other_side = right_shapes
                        other_has_panel_img = any(
                            _is_panel_image(s) for s in right_shapes
                        )
                    else:
                        other_side = left_shapes
                        other_has_panel_img = any(
                            _is_panel_image(s) for s in left_shapes
                        )
                    if len(other_side) >= 2 and not other_has_panel_img:
                        asymmetric_detected = True
                        logger.debug(
                            'Split panel: asymmetric — image anchor on %s, '
                            '%d shapes on %s',
                            'left' if anchor is left_anchor else 'right',
                            len(other_side),
                            'right' if anchor is left_anchor else 'left',
                        )
            # Also try: no anchors at all, but one side has a panel-sized
            # image (>40% width or >20% area) and the other doesn't
            if not asymmetric_detected and left_anchor is None and right_anchor is None:
                left_has_panel = any(_is_panel_image(s) for s in left_shapes)
                right_has_panel = any(_is_panel_image(s) for s in right_shapes)
                if left_has_panel != right_has_panel:
                    img_side = left_shapes if left_has_panel else right_shapes
                    txt_side = right_shapes if left_has_panel else left_shapes
                    if len(txt_side) >= 2:
                        asymmetric_detected = True
                        logger.debug(
                            'Split panel: asymmetric (no anchors) — '
                            'panel image on %s, %d shapes on %s',
                            'left' if left_has_panel else 'right',
                            len(txt_side),
                            'right' if left_has_panel else 'left',
                        )

        # ── Strategy 2b: Dominant image + single text on opposite side ─
        # Handles slides like Lukas S1: one large photo (>=40% W, >=80% H)
        # on one side, exactly 1 non-image text shape (>=25% W) on the other.
        # Strategy 2 requires >=2 shapes on the text side; this relaxes to 1
        # when the image is clearly dominant (near full-height).
        dominant_detected = False
        if not anchor_detected and not asymmetric_detected:
            if (left_anchor is not None) != (right_anchor is not None):
                anchor = left_anchor if left_anchor is not None else right_anchor
                if _is_image(anchor):
                    wr = anchor.width / self.slide_width if self.slide_width else 0
                    hr = anchor.height / self.slide_height if self.slide_height else 0
                    if wr >= _DOMINANT_IMAGE_MIN_WIDTH and hr >= _DOMINANT_IMAGE_MIN_HEIGHT:
                        other_side = (
                            right_shapes if anchor is left_anchor else left_shapes
                        )
                        # Exactly 1 non-image shape, wide enough to be content
                        non_img = [
                            s for s in other_side if not _is_image(s)
                        ]
                        if (len(non_img) == 1
                                and non_img[0].width
                                >= self.slide_width * _OPPOSITE_TEXT_MIN_WIDTH):
                            dominant_detected = True
                            logger.debug(
                                'Split panel: dominant image (%.0f%%W × %.0f%%H) '
                                'on %s, 1 text shape on %s',
                                wr * 100, hr * 100,
                                'left' if anchor is left_anchor else 'right',
                                'right' if anchor is left_anchor else 'left',
                            )

        # ── Strategy 3: Cluster-based fallback ────────────────────────
        cluster_detected = False
        if not anchor_detected and not asymmetric_detected and not dominant_detected:
            def _cluster_bbox(
                shape_list: List[_ShapeData],
            ) -> Tuple[int, int, int, int]:
                if not shape_list:
                    return 0, 0, 0, 0
                min_x = min(s.left for s in shape_list)
                max_x = max(s.right for s in shape_list)
                min_y = min(s.top for s in shape_list)
                max_y = max(s.bottom for s in shape_list)
                return min_x, max_x, min_y, max_y

            l_x1, l_x2, l_y1, l_y2 = _cluster_bbox(left_shapes)
            r_x1, r_x2, r_y1, r_y2 = _cluster_bbox(right_shapes)

            l_w = l_x2 - l_x1
            l_h = l_y2 - l_y1
            r_w = r_x2 - r_x1
            r_h = r_y2 - r_y1

            min_w = self.slide_width * _CLUSTER_MIN_WIDTH_FRACTION
            min_h = self.slide_height * _CLUSTER_MIN_HEIGHT_FRACTION

            left_qualifies = l_w >= min_w and l_h >= min_h
            right_qualifies = r_w >= min_w and r_h >= min_h

            gap = r_x1 - l_x2
            min_gap = self.slide_width * _CLUSTER_MIN_GAP_FRACTION

            if ((left_qualifies or right_qualifies)
                    and gap >= min_gap
                    and (len(left_shapes) >= 2 or len(right_shapes) >= 2)):
                left_has_panel = any(_is_panel_image(s) for s in left_shapes)
                right_has_panel = any(_is_panel_image(s) for s in right_shapes)
                if left_has_panel != right_has_panel:
                    cluster_detected = True

        if not (anchor_detected or asymmetric_detected
                or dominant_detected or cluster_detected):
            return None

        # ── Classify shapes into left/right panels ────────────────────
        # For anchor-based, use the found anchors.
        # For asymmetric/cluster, synthesize anchor IDs from the largest
        # shape on each side.
        if not anchor_detected:
            # Pick the largest shape on each side as synthetic anchor
            left_anchor = max(left_shapes, key=lambda s: s.area)
            right_anchor = max(right_shapes, key=lambda s: s.area)

        left_ids: Set[int] = set()
        right_ids: Set[int] = set()
        centerline_margin = int(
            self.slide_width * _CENTERLINE_EXCLUSION_FRACTION,
        )

        for sd in shapes:
            if sd.shape_id in background_ids:
                continue
            if sd.width >= self.slide_width * _FULL_WIDTH_FRACTION:
                continue

            visible_left = max(0, sd.left)
            visible_right = min(self.slide_width, sd.right)
            visible_center_x = (visible_left + visible_right) // 2

            # Anchors always assigned to their panel
            if sd.shape_id == left_anchor.shape_id:
                left_ids.add(sd.shape_id)
                continue
            if sd.shape_id == right_anchor.shape_id:
                right_ids.add(sd.shape_id)
                continue

            # Exclude small shapes near the center line — classify individually.
            # Large shapes (>30% width) or placeholders that straddle the center
            # are content elements that belong to a panel — assign by center_x.
            distance_from_midpoint = abs(visible_center_x - self.half_width)
            is_large = sd.width >= self.slide_width * 0.30
            if distance_from_midpoint < centerline_margin and not is_large:
                continue

            if visible_center_x < self.half_width:
                left_ids.add(sd.shape_id)
            else:
                right_ids.add(sd.shape_id)

        if not left_ids or not right_ids:
            return None

        # ── Compute panel bounding boxes and shift delta ──────────────
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
            dominant_image=dominant_detected,
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
        # FIX: text-bearing wide shapes must NOT be classified as BACKGROUND
        # (was silently suppressing translation on full-width text banners)
        if sd.width >= self.slide_width * _BG_WIDTH_FRACTION and not sd.has_text:
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
            # FIX: text-bearing shapes must NOT be classified as DECORATIVE
            # (was silently suppressing translation on column headers, labels, etc.)
            if (sd.width < self.slide_width * _DECORATIVE_MAX_WIDTH_FRACTION
                    or sd.height < self.slide_height * _DECORATIVE_MAX_HEIGHT_FRACTION):
                if not sd.has_blip_fill and not sd.is_picture and not sd.has_text:
                    return _make_classification(
                        ShapeRole.DECORATIVE,
                        rule_name='decorative_title_layout_small',
                    )
            # Content-sized shapes on title layouts get mirrored
            is_content_sized = (
                sd.width > self.slide_width * _CONTENT_MIN_WIDTH_FRACTION
                and sd.height > self.slide_height * _CONTENT_MIN_HEIGHT_FRACTION
            )
            if not is_content_sized and not sd.has_blip_fill and not sd.is_picture and not sd.has_text:
                # Small non-image, non-text shape on title layout → decorative
                return _make_classification(
                    ShapeRole.DECORATIVE,
                    rule_name='decorative_title_layout_noncontent',
                    confidence=0.7,
                )

        # ── Priority 11.5: COMPLEX_GRAPHIC ──
        # Groups with many children or mixed content represent infographics
        # (SWOT, timelines, org charts, process flows) that must NOT be mirrored.
        # User hard rule: "complex shapes and graphics should only be translated not mirrored"
        if sd.is_group:
            try:
                children = list(sd.shape.shapes)
                child_count = len(children)
                text_children = sum(1 for c in children if getattr(c, 'has_text_frame', False))
                non_text_children = child_count - text_children

                # Condition 1: Many children (almost certainly an infographic)
                # Condition 2: Large group with mixed content (text + non-text)
                has_mixed = text_children >= 2 and non_text_children >= 2
                is_large = (
                    sd.width > self.slide_width * 0.35
                    or sd.height > self.slide_height * 0.35
                )

                if child_count >= 6 or (has_mixed and is_large and child_count >= 4):
                    return _make_classification(
                        ShapeRole.COMPLEX_GRAPHIC,
                        rule_name=f'complex_graphic_{child_count}children',
                    )
            except Exception:
                pass

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
