"""
SlideArabi — Data Models

Frozen dataclasses representing the fully-resolved OOXML presentation structure.
Every effective property has a concrete value — no None for properties that should
always be resolvable (font size, font name, bold, italic, alignment, RTL, level).

Design principles:
1. Immutability: All resolved models are frozen dataclasses with tuple collections.
2. No None for effective values: The PropertyResolver guarantees concrete values
   by walking the 7-level OOXML inheritance chain.
3. Provenance: source_level / source_font_size_level tracks where each value
   was resolved from, enabling debugging and auditing.
4. Separation: Resolved models (read-only snapshots) are separate from
   TransformPlan/TransformAction (mutable planning models for Phase 2/3).

OOXML Constants:
- Font sizes in hundredths of a point (e.g., 1800 = 18pt)
- EMU (English Metric Units): 1 inch = 914400 EMU, 1 pt = 12700 EMU
- Placeholder types from ST_PlaceholderType
- Layout types from ST_SlideLayoutType
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# OOXML NAMESPACE CONSTANTS
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

NSMAP = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
}

A_NS = NSMAP['a']
P_NS = NSMAP['p']
R_NS = NSMAP['r']


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# DEFAULT FALLBACK VALUES (PowerPoint built-in defaults)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

DEFAULT_FONT_SIZE_PT = 18.0    # PowerPoint's ultimate default
DEFAULT_FONT_NAME = 'Calibri'  # Office default since 2007
DEFAULT_BOLD = False
DEFAULT_ITALIC = False
DEFAULT_UNDERLINE = False
DEFAULT_ALIGNMENT = 'l'        # left
DEFAULT_RTL = False
DEFAULT_LEVEL = 0
DEFAULT_ROTATION = 0.0


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# VALID ENUMERATIONS
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

VALID_ALIGNMENTS = frozenset({'l', 'r', 'ctr', 'just', 'dist'})

VALID_SHAPE_TYPES = frozenset({
    'placeholder', 'textbox', 'picture', 'chart', 'table',
    'group', 'connector', 'freeform', 'ole', 'smartart', 'media',
})

VALID_PLACEHOLDER_TYPES = frozenset({
    'title', 'body', 'ctrTitle', 'subTitle', 'dt', 'ftr', 'sldNum',
    'pic', 'chart', 'tbl', 'dgm', 'media', 'clipArt', 'obj',
})

VALID_LAYOUT_TYPES = frozenset({
    'title', 'tx', 'twoColTx', 'obj', 'secHead', 'blank', 'tbl',
    'chart', 'txAndChart', 'picTx', 'cust', 'titleOnly', 'twoObj',
    'objTx', 'txAndObj', 'dgm', 'txOverObj', 'objOverTx',
    'fourObj', 'objAndTx', 'vertTx', 'vertTitleAndTx', 'clipArtAndTx',
    'txAndClipArt', 'mediaAndTx', 'txAndMedia', 'objAndTwoObj',
    'twoObjAndObj', 'twoObjOverTx', 'txAndTwoObj', 'twoTxTwoObj',
    'txOverObj2', 'objOverTx2',
})

VALID_SOURCE_LEVELS = frozenset({'master', 'layout', 'slide'})

# Labels for the 7-level inheritance chain (used in source_font_size_level)
INHERITANCE_LEVELS = (
    'run',           # Level 1: a:rPr on the run
    'paragraph',     # Level 2: a:pPr/a:defRPr
    'textframe',     # Level 3: a:lstStyle on shape's text body
    'shape',         # Level 4: shape-level inline or lstStyle
    'layout',        # Level 5: layout placeholder
    'master',        # Level 6: master placeholder
    'txstyles',      # Level 7: master p:txStyles (title/body/other)
    'default',       # Fallback: hardcoded PowerPoint default
)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# RESOLVED MODELS — Immutable Phase 0 snapshot
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

@dataclass(frozen=True)
class ResolvedRun:
    """A single text run with all properties fully resolved through the 7-level
    OOXML inheritance chain.

    No effective property is None (except effective_color, which is None only
    when the text is genuinely transparent/no-fill).

    Attributes:
        text: The raw text content of the run.
        effective_font_size_pt: Font size in points, resolved from the chain.
            NEVER None — guaranteed by PropertyResolver.
        effective_font_name: Resolved font family name. Theme references like
            +mj-lt and +mn-lt are resolved to the actual theme font.
        effective_bold: Whether bold is applied.
        effective_italic: Whether italic is applied.
        effective_color: Hex color string (e.g., 'FF0000') or None if no fill.
        effective_underline: Whether underline is applied.
        source_font_size_level: Which inheritance level provided the font size.
            One of: 'run', 'paragraph', 'textframe', 'shape', 'layout',
            'master', 'txstyles', 'default'.
    """
    text: str
    effective_font_size_pt: float
    effective_font_name: str
    effective_bold: bool
    effective_italic: bool
    effective_color: Optional[str]
    effective_underline: bool
    source_font_size_level: str


@dataclass(frozen=True)
class ResolvedParagraph:
    """A paragraph with all properties resolved, containing resolved runs.

    Attributes:
        runs: Immutable tuple of ResolvedRun objects.
        effective_alignment: Paragraph alignment — one of 'l', 'r', 'ctr',
            'just', 'dist'.
        effective_rtl: Whether RTL direction is set.
        effective_level: Paragraph indent level (0–8).
        effective_bullet_type: Bullet type string or None if no bullet.
        effective_line_spacing: Line spacing multiplier (e.g., 1.5) or None.
        effective_space_before: Space before paragraph in points, or None.
        effective_space_after: Space after paragraph in points, or None.
    """
    runs: Tuple[ResolvedRun, ...]
    effective_alignment: str
    effective_rtl: bool
    effective_level: int
    effective_bullet_type: Optional[str]
    effective_line_spacing: Optional[float]
    effective_space_before: Optional[float]
    effective_space_after: Optional[float]


@dataclass(frozen=True)
class ResolvedShape:
    """A shape with all text/formatting properties resolved from the inheritance
    chain, plus positional information in EMU.

    Attributes:
        shape_id: The numeric ID of the shape within the presentation.
        shape_name: Human-readable shape name (e.g., 'Title 1').
        shape_type: Classification — one of VALID_SHAPE_TYPES.
        placeholder_type: Placeholder type from ST_PlaceholderType, or None
            for non-placeholder shapes.
        placeholder_idx: Placeholder index, or None.
        x_emu: Left position in EMU.
        y_emu: Top position in EMU.
        width_emu: Width in EMU.
        height_emu: Height in EMU.
        rotation_degrees: Rotation angle in degrees.
        paragraphs: Tuple of ResolvedParagraph objects.
        is_master_inherited: True if this shape originates from master/layout
            (not directly on the slide).
        source_level: Which level defines this shape — 'master', 'layout',
            or 'slide'.
        has_local_position_override: True if the slide-level shape overrides
            layout/master position.
        has_text: True if the shape contains any non-empty text.
        original_xml_element: Reference to the lxml element for write-back.
            This field is excluded from hash/eq since it's mutable.
    """
    shape_id: int
    shape_name: str
    shape_type: str
    placeholder_type: Optional[str]
    placeholder_idx: Optional[int]
    x_emu: int
    y_emu: int
    width_emu: int
    height_emu: int
    rotation_degrees: float
    paragraphs: Tuple[ResolvedParagraph, ...]
    is_master_inherited: bool
    source_level: str
    has_local_position_override: bool
    has_text: bool
    original_xml_element: Any = field(compare=False, hash=False, default=None)

    @property
    def full_text(self) -> str:
        """Concatenate all run text across all paragraphs."""
        parts = []
        for para in self.paragraphs:
            para_text = ''.join(r.text for r in para.runs)
            parts.append(para_text)
        return '\n'.join(parts)

    @property
    def is_placeholder(self) -> bool:
        """True if this shape is a placeholder."""
        return self.shape_type == 'placeholder'


@dataclass(frozen=True)
class ResolvedLayout:
    """A fully resolved slide layout with its placeholders and shapes.

    Attributes:
        layout_name: Display name of the layout.
        layout_type: ST_SlideLayoutType value (e.g., 'title', 'tx',
            'twoColTx', 'obj', 'cust').
        master_index: Index of the parent master in the presentation.
        placeholders: Resolved placeholder shapes on this layout.
        freeform_shapes: Non-placeholder shapes on this layout.
    """
    layout_name: str
    layout_type: str
    master_index: int
    placeholders: Tuple[ResolvedShape, ...]
    freeform_shapes: Tuple[ResolvedShape, ...]


@dataclass(frozen=True)
class ResolvedMaster:
    """A fully resolved slide master.

    Attributes:
        master_name: Display name of the master.
        master_index: Index position in the presentation's master list.
        placeholders: Resolved placeholder shapes on this master.
        freeform_shapes: Non-placeholder shapes on this master.
        tx_styles: Dictionary with keys 'titleStyle', 'bodyStyle', 'otherStyle',
            each mapping to level-specific property dicts from p:txStyles.
    """
    master_name: str
    master_index: int
    placeholders: Tuple[ResolvedShape, ...]
    freeform_shapes: Tuple[ResolvedShape, ...]
    tx_styles: Dict[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class ResolvedSlide:
    """A fully resolved slide with all shapes property-resolved.

    Attributes:
        slide_number: 1-based slide number.
        layout_name: Name of the slide's layout.
        layout_type: ST_SlideLayoutType value for the slide's layout.
        layout_index: Index of the layout within its master.
        master_index: Index of the master this slide's layout belongs to.
        shapes: All resolved shapes on this slide.
    """
    slide_number: int
    layout_name: str
    layout_type: str
    layout_index: int
    master_index: int
    shapes: Tuple[ResolvedShape, ...]


@dataclass(frozen=True)
class ResolvedPresentation:
    """The complete resolved presentation — an immutable snapshot produced by
    PropertyResolver in Phase 0.

    All shapes on every slide, layout, and master have fully resolved text
    properties. This is the single source of truth for all subsequent phases.

    Attributes:
        slide_width_emu: Presentation slide width in EMU.
        slide_height_emu: Presentation slide height in EMU.
        masters: Resolved slide masters.
        layouts: Resolved slide layouts.
        slides: Resolved slides.
    """
    slide_width_emu: int
    slide_height_emu: int
    masters: Tuple[ResolvedMaster, ...]
    layouts: Tuple[ResolvedLayout, ...]
    slides: Tuple[ResolvedSlide, ...]

    @property
    def total_shapes(self) -> int:
        """Total number of shapes across all slides."""
        return sum(len(s.shapes) for s in self.slides)

    @property
    def total_slides(self) -> int:
        """Number of slides."""
        return len(self.slides)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# TRANSFORM PLANNING MODELS — Mutable, for Phase 2/3
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

VALID_ACTION_TYPES = frozenset({
    'mirror',         # Mirror X position: new_x = slide_width - (x + width)
    'swap',           # Swap positions of two shapes
    'keep',           # Keep position unchanged
    'right_align',    # Set text alignment to right
    'center_align',   # Set text alignment to center
    'set_rtl',        # Set paragraph RTL direction
    'set_font',       # Change font family
    'resize_font',    # Adjust font size
    'reverse_columns', # Reverse table columns
    'reverse_axes',   # Reverse chart category axis
    'set_language',   # Set language attribute
    'remove_position', # Remove local position override (re-inherit from layout)
})


@dataclass
class TransformAction:
    """A single atomic transformation to apply to a shape.

    Attributes:
        shape_id: Target shape ID.
        action_type: Type of transformation — one of VALID_ACTION_TYPES.
        params: Action-specific parameters dictionary. Examples:
            - mirror: {'slide_width_emu': int}
            - swap: {'swap_with_shape_id': int}
            - set_rtl: {'value': True}
            - set_font: {'font_name': str}
            - resize_font: {'new_size_pt': float}
    """
    shape_id: int
    action_type: str
    params: Dict[str, Any] = field(default_factory=dict)

    def __post_init__(self):
        if self.action_type not in VALID_ACTION_TYPES:
            raise ValueError(
                f"Invalid action_type '{self.action_type}'. "
                f"Must be one of: {sorted(VALID_ACTION_TYPES)}"
            )


@dataclass
class TransformPlan:
    """A mutable plan of transformations for an entire presentation.

    Built during Phase 2 (master/layout transforms) and Phase 3 (slide transforms),
    then executed atomically.

    Attributes:
        slide_actions: Mapping from slide number to list of TransformActions.
        master_actions: Mapping from master index to list of TransformActions.
        layout_actions: Mapping from (master_index, layout_index) to actions.
        metadata: Optional metadata about the plan (e.g., rule source).
    """
    slide_actions: Dict[int, List[TransformAction]] = field(default_factory=dict)
    master_actions: Dict[int, List[TransformAction]] = field(default_factory=dict)
    layout_actions: Dict[Tuple[int, int], List[TransformAction]] = field(
        default_factory=dict
    )
    metadata: Dict[str, Any] = field(default_factory=dict)

    def add_slide_action(self, slide_number: int, action: TransformAction) -> None:
        """Add a transform action for a specific slide."""
        if slide_number not in self.slide_actions:
            self.slide_actions[slide_number] = []
        self.slide_actions[slide_number].append(action)

    def add_master_action(self, master_index: int, action: TransformAction) -> None:
        """Add a transform action for a specific master."""
        if master_index not in self.master_actions:
            self.master_actions[master_index] = []
        self.master_actions[master_index].append(action)

    def add_layout_action(
        self, master_index: int, layout_index: int, action: TransformAction
    ) -> None:
        """Add a transform action for a specific layout."""
        key = (master_index, layout_index)
        if key not in self.layout_actions:
            self.layout_actions[key] = []
        self.layout_actions[key].append(action)

    @property
    def total_actions(self) -> int:
        """Total number of actions across all targets."""
        count = sum(len(v) for v in self.slide_actions.values())
        count += sum(len(v) for v in self.master_actions.values())
        count += sum(len(v) for v in self.layout_actions.values())
        return count


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# VALIDATION MODELS — Phase 5 output
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

VALID_SEVERITIES = frozenset({'error', 'warning', 'info'})


@dataclass(frozen=True)
class ValidationIssue:
    """A single validation issue found during Phase 5 structural validation.

    Attributes:
        severity: 'error', 'warning', or 'info'.
        slide_number: Slide where the issue was found.
        shape_id: Shape ID if applicable, None for slide-level issues.
        issue_type: Machine-readable issue category (e.g., 'rtl_missing',
            'overlap_detected', 'font_below_minimum', 'out_of_bounds').
        message: Human-readable description.
        expected_value: What the validator expected.
        actual_value: What was found.
    """
    severity: str
    slide_number: int
    shape_id: Optional[int]
    issue_type: str
    message: str
    expected_value: Any = None
    actual_value: Any = None


@dataclass(frozen=True)
class ValidationReport:
    """Complete validation report from Phase 5.

    Attributes:
        issues: Tuple of all validation issues found.
        total_shapes_checked: Number of shapes examined.
        total_slides_checked: Number of slides examined.
    """
    issues: Tuple[ValidationIssue, ...]
    total_shapes_checked: int = 0
    total_slides_checked: int = 0

    @property
    def error_count(self) -> int:
        """Number of error-severity issues."""
        return sum(1 for i in self.issues if i.severity == 'error')

    @property
    def warning_count(self) -> int:
        """Number of warning-severity issues."""
        return sum(1 for i in self.issues if i.severity == 'warning')

    @property
    def info_count(self) -> int:
        """Number of info-severity issues."""
        return sum(1 for i in self.issues if i.severity == 'info')

    @property
    def has_errors(self) -> bool:
        """True if any error-severity issues exist."""
        return self.error_count > 0

    @property
    def passed(self) -> bool:
        """True if no errors were found (warnings are acceptable)."""
        return not self.has_errors
