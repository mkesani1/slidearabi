"""
template_registry.py — RTL transformation rules per layout type.

SlideArabi: Template-First Deterministic RTL Transformation Engine.

Maps each of the 36 standard OOXML ST_SlideLayoutType values to a
prescriptive set of RTL transformation rules.  The TemplateRegistry is the
"brain" of the deterministic pipeline — given a layout type it returns
exactly what to do with every placeholder, freeform shape, table, and chart.

Design principles:
1. Every layout type has explicit rules — no ambiguity.
2. Rules are composable: placeholder-level, shape-level, and slide-level.
3. Custom rules can be registered at runtime to handle enterprise layouts.
4. Unknown / custom layouts fall back to the 'cust' ruleset with conservative
   defaults (mirror freeform, set RTL, flag for review).
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from typing import Dict, Optional

logger = logging.getLogger(__name__)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# ARABIC FONT MAP — deterministic mapping for Arabic typography
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

ARABIC_FONT_MAP: Dict[str, str] = {
    'Calibri': 'Calibri',
    'Arial': 'Arial',
    'Times New Roman': 'Times New Roman',
    'Calibri Light': 'Calibri Light',
    'Tahoma': 'Tahoma',
    'Segoe UI': 'Segoe UI',
    'Cambria': 'Sakkal Majalla',
    'Georgia': 'Sakkal Majalla',
    'Verdana': 'Tahoma',
    'Trebuchet MS': 'Tahoma',
    'Garamond': 'Traditional Arabic',
    'Palatino': 'Traditional Arabic',
    'Palatino Linotype': 'Traditional Arabic',
    'Century Gothic': 'Dubai',
    'Futura': 'Dubai',
    'Helvetica': 'Arial',
    'Helvetica Neue': 'Arial',
    'Impact': 'Arial Black',
    'Comic Sans MS': 'Tahoma',
    'Courier New': 'Courier New',
    'Consolas': 'Courier New',
}


def get_arabic_font(latin_font: str) -> str:
    """Look up the recommended Arabic font for a given Latin font.

    Falls back to 'Calibri' if the font is not in the mapping.

    Args:
        latin_font: Name of the Latin/Western font.

    Returns:
        Recommended Arabic-capable font name.
    """
    return ARABIC_FONT_MAP.get(latin_font, 'Calibri')


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Rule data classes
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

@dataclass
class PlaceholderAction:
    """Prescriptive RTL action for a single placeholder.

    Attributes:
        action: The transform verb — one of:
            'keep_centered'     — preserve horizontal center, just set RTL
            'right_align'       — set text alignment to right
            'mirror'            — mirror X position across slide center
            'swap_with_partner' — swap position with a paired placeholder
            'keep_position'     — leave position unchanged, apply text RTL
            'reverse_columns'   — reverse table column order
            'reverse_axes'      — reverse chart category axis direction
        set_rtl: Whether to set RTL direction on paragraphs.
        set_alignment: Target alignment value ('r', 'l', 'ctr', None).
            None means "do not override alignment".
        swap_partner_idx: Placeholder index of the swap partner, or None.
        mirror_x: Whether to mirror the X coordinate.
        notes: Human-readable explanation for debugging.
    """
    action: str
    set_rtl: bool = True
    set_alignment: Optional[str] = None
    swap_partner_idx: Optional[int] = None
    mirror_x: bool = False
    notes: str = ''


@dataclass
class LayoutTransformRules:
    """Complete RTL transformation ruleset for a layout type.

    Attributes:
        layout_type: Canonical ST_SlideLayoutType string.
        description: Human-readable layout description.
        placeholder_rules: Map of placeholder_type (or 'idx_N') → PlaceholderAction.
        freeform_action: Default action for non-placeholder shapes:
            'mirror' — mirror X position (most common)
            'keep'   — leave position unchanged
            'analyze' — needs per-shape analysis
        mirror_master_elements: Whether to mirror inherited master shapes.
        swap_columns: Whether this layout has columns that should be swapped.
        table_action: How to handle tables:
            'reverse_columns' — reverse column order and set RTL
            'keep'           — leave table as-is
            'rtl_only'       — only set RTL on cells, don't reorder
        chart_action: How to handle charts:
            'reverse_axes'   — reverse category axis direction
            'keep'           — leave chart as-is
            'mirror_legend'  — mirror legend position only
    """
    layout_type: str
    description: str
    placeholder_rules: Dict[str, PlaceholderAction] = field(default_factory=dict)
    freeform_action: str = 'mirror'
    mirror_master_elements: bool = True
    swap_columns: bool = False
    table_action: str = 'reverse_columns'
    chart_action: str = 'reverse_axes'


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# TemplateRegistry
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

class TemplateRegistry:
    """Registry of RTL transformation rules indexed by ST_SlideLayoutType.

    Provides deterministic lookup of how each placeholder, freeform shape,
    table, and chart should be transformed when converting a presentation
    from LTR to RTL.

    Usage::

        registry = TemplateRegistry(slide_width_emu, slide_height_emu)
        rules = registry.get_rules('twoColTx')
        action = registry.get_placeholder_action('twoColTx', 'body', idx=1)
    """

    def __init__(self, slide_width_emu: int, slide_height_emu: int):
        """Initialise with slide dimensions.

        Args:
            slide_width_emu: Slide width in EMU.
            slide_height_emu: Slide height in EMU.
        """
        self._slide_width = slide_width_emu
        self._slide_height = slide_height_emu
        self._rules: Dict[str, LayoutTransformRules] = self._build_default_rules()
        self._custom_rules: Dict[str, LayoutTransformRules] = {}

    # ── public API ──────────────────────────────────────────────────────────

    def get_rules(self, layout_type: str) -> LayoutTransformRules:
        """Get transformation rules for a layout type.

        Checks custom rules first, then built-in defaults. Falls back to
        'cust' rules if the layout type is unknown.

        Args:
            layout_type: ST_SlideLayoutType string.

        Returns:
            ``LayoutTransformRules`` for the given type.
        """
        # Custom rules take precedence
        if layout_type in self._custom_rules:
            return self._custom_rules[layout_type]

        if layout_type in self._rules:
            return self._rules[layout_type]

        logger.debug(
            'Unknown layout type %r — falling back to cust rules', layout_type,
        )
        return self._rules['cust']

    def get_placeholder_action(
        self,
        layout_type: str,
        placeholder_type: str,
        placeholder_idx: int = 0,
    ) -> PlaceholderAction:
        """Get the specific action for a placeholder.

        Lookup order:
        1. 'idx_N' key (most specific — by placeholder index).
        2. placeholder_type key (by semantic type).
        3. Default: keep_position with RTL.

        Args:
            layout_type: ST_SlideLayoutType string.
            placeholder_type: Normalised placeholder type string.
            placeholder_idx: Placeholder index from OOXML.

        Returns:
            ``PlaceholderAction`` for the placeholder.
        """
        rules = self.get_rules(layout_type)

        # Try idx-specific key first
        idx_key = f'idx_{placeholder_idx}'
        if idx_key in rules.placeholder_rules:
            return rules.placeholder_rules[idx_key]

        # Try type key
        if placeholder_type in rules.placeholder_rules:
            return rules.placeholder_rules[placeholder_type]

        # Default action: keep position, set RTL, right-align text
        return PlaceholderAction(
            action='keep_position',
            set_rtl=True,
            set_alignment='r',
            notes=f'Default action for {placeholder_type} in {layout_type}',
        )

    def get_freeform_action(self, layout_type: str) -> str:
        """Get the default action for non-placeholder shapes.

        Args:
            layout_type: ST_SlideLayoutType string.

        Returns:
            Action string: 'mirror', 'keep', or 'analyze'.
        """
        rules = self.get_rules(layout_type)
        return rules.freeform_action

    def register_custom_rule(
        self, layout_name: str, rules: LayoutTransformRules
    ) -> None:
        """Register a custom transformation rule.

        Custom rules override built-in defaults for the given layout name.
        Use this for enterprise-specific layouts that don't map to standard
        OOXML types.

        Args:
            layout_name: Identifier for the custom layout.
            rules: ``LayoutTransformRules`` to associate with it.
        """
        self._custom_rules[layout_name] = rules
        logger.info('Registered custom rule for layout: %s', layout_name)

    def list_layout_types(self) -> list:
        """Return a sorted list of all registered layout type keys.

        Includes both built-in and custom rules.
        """
        keys = set(self._rules.keys()) | set(self._custom_rules.keys())
        return sorted(keys)

    # ── rule builder ────────────────────────────────────────────────────────

    def _build_default_rules(self) -> Dict[str, LayoutTransformRules]:
        """Build comprehensive RTL rules for all 36 standard layout types.

        Rules are grouped by transformation pattern:
        - Centered:       title, secHead
        - Single content: tx, obj, titleOnly, objOnly, blank
        - Two-column:     twoColTx, txAndChart, chartAndTx, txAndObj, objAndTx,
                          txAndClipArt, clipArtAndTx, txAndMedia, mediaAndTx, picTx
        - Multi-object:   twoObj, fourObj, txAndTwoObj, twoObjAndTx, twoObjOverTx,
                          twoTxTwoObj, twoObjAndObj, objAndTwoObj
        - Content-heavy:  objTx, txObj, objOverTx, txOverObj
        - Vertical:       vertTx, vertTitleAndTx, vertTitleAndTxOverChart
        - Specialised:    tbl, chart, dgm
        - Fallback:       cust

        Returns:
            Dict mapping layout type string → LayoutTransformRules.
        """
        rules: Dict[str, LayoutTransformRules] = {}

        # ──────────────────────────────────────────────────────────────
        # CENTERED LAYOUTS
        # ──────────────────────────────────────────────────────────────

        rules['title'] = LayoutTransformRules(
            layout_type='title',
            description='Title Slide — centered title and subtitle',
            placeholder_rules={
                'ctrTitle': PlaceholderAction(
                    action='keep_centered',
                    set_rtl=True,
                    set_alignment='ctr',
                    notes='Center title stays centered; set RTL',
                ),
                'subTitle': PlaceholderAction(
                    action='right_align',
                    set_rtl=True,
                    set_alignment='r',
                    notes='Subtitle right-aligns for RTL',
                ),
                'title': PlaceholderAction(
                    action='keep_centered',
                    set_rtl=True,
                    set_alignment='ctr',
                    notes='Fallback: title stays centered',
                ),
            },
            freeform_action='mirror',
            mirror_master_elements=False,
            swap_columns=False,
            table_action='reverse_columns',
            chart_action='reverse_axes',
        )

        rules['secHead'] = LayoutTransformRules(
            layout_type='secHead',
            description='Section Header — title + body both right-aligned',
            placeholder_rules={
                'title': PlaceholderAction(
                    action='right_align',
                    set_rtl=True,
                    set_alignment='r',
                    notes='Section header title right-aligns',
                ),
                'body': PlaceholderAction(
                    action='right_align',
                    set_rtl=True,
                    set_alignment='r',
                    notes='Section header body right-aligns',
                ),
            },
            freeform_action='keep',
            mirror_master_elements=False,
            swap_columns=False,
            table_action='reverse_columns',
            chart_action='reverse_axes',
        )

        # ──────────────────────────────────────────────────────────────
        # SINGLE CONTENT LAYOUTS
        # ──────────────────────────────────────────────────────────────

        rules['tx'] = LayoutTransformRules(
            layout_type='tx',
            description='Title + Text body',
            placeholder_rules={
                'title': PlaceholderAction(
                    action='keep_position',
                    set_rtl=True,
                    set_alignment='r',
                    notes='Title right-aligns',
                ),
                'body': PlaceholderAction(
                    action='keep_position',
                    set_rtl=True,
                    set_alignment='r',
                    notes='Body keeps position; RTL + right-align',
                ),
            },
            freeform_action='mirror',
            mirror_master_elements=True,
            swap_columns=False,
            table_action='reverse_columns',
            chart_action='reverse_axes',
        )

        rules['titleOnly'] = LayoutTransformRules(
            layout_type='titleOnly',
            description='Title Only — no body',
            placeholder_rules={
                'title': PlaceholderAction(
                    action='keep_position',
                    set_rtl=True,
                    set_alignment='r',
                    notes='Title right-aligns; no body to transform',
                ),
            },
            freeform_action='mirror',
            mirror_master_elements=True,
            swap_columns=False,
            table_action='reverse_columns',
            chart_action='reverse_axes',
        )

        rules['obj'] = LayoutTransformRules(
            layout_type='obj',
            description='Title + Object',
            placeholder_rules={
                'title': PlaceholderAction(
                    action='keep_position',
                    set_rtl=True,
                    set_alignment='r',
                    notes='Title right-aligns',
                ),
                'obj': PlaceholderAction(
                    action='keep_position',
                    set_rtl=False,
                    notes='Object placeholder keeps position',
                ),
            },
            freeform_action='mirror',
            mirror_master_elements=True,
            swap_columns=False,
            table_action='reverse_columns',
            chart_action='reverse_axes',
        )

        rules['objOnly'] = LayoutTransformRules(
            layout_type='objOnly',
            description='Object Only — no title',
            placeholder_rules={
                'obj': PlaceholderAction(
                    action='keep_position',
                    set_rtl=False,
                    notes='Object keeps position',
                ),
            },
            freeform_action='mirror',
            mirror_master_elements=True,
            swap_columns=False,
            table_action='reverse_columns',
            chart_action='reverse_axes',
        )

        rules['blank'] = LayoutTransformRules(
            layout_type='blank',
            description='Blank — no placeholders',
            placeholder_rules={},
            freeform_action='mirror',
            mirror_master_elements=True,
            swap_columns=False,
            table_action='reverse_columns',
            chart_action='reverse_axes',
        )

        # ──────────────────────────────────────────────────────────────
        # TWO-COLUMN / SIDE-BY-SIDE LAYOUTS
        # ──────────────────────────────────────────────────────────────

        _two_col_title = PlaceholderAction(
            action='keep_position',
            set_rtl=True,
            set_alignment='r',
            notes='Title right-aligns across full width',
        )

        rules['twoColTx'] = LayoutTransformRules(
            layout_type='twoColTx',
            description='Two Column Text — swap L/R columns',
            placeholder_rules={
                'title': _two_col_title,
                'body': PlaceholderAction(
                    action='swap_with_partner',
                    set_rtl=True,
                    set_alignment='r',
                    notes='Body columns swap positions',
                ),
            },
            freeform_action='mirror',
            mirror_master_elements=True,
            swap_columns=True,
            table_action='reverse_columns',
            chart_action='reverse_axes',
        )

        rules['txAndChart'] = LayoutTransformRules(
            layout_type='txAndChart',
            description='Text + Chart — text on left, chart on right; swap',
            placeholder_rules={
                'title': _two_col_title,
                'body': PlaceholderAction(
                    action='swap_with_partner',
                    set_rtl=True,
                    set_alignment='r',
                    notes='Text body swaps with chart',
                ),
                'chart': PlaceholderAction(
                    action='swap_with_partner',
                    set_rtl=False,
                    notes='Chart swaps with text body',
                ),
            },
            freeform_action='mirror',
            mirror_master_elements=True,
            swap_columns=True,
            table_action='reverse_columns',
            chart_action='reverse_axes',
        )

        rules['chartAndTx'] = LayoutTransformRules(
            layout_type='chartAndTx',
            description='Chart + Text — chart on left, text on right; swap',
            placeholder_rules={
                'title': _two_col_title,
                'chart': PlaceholderAction(
                    action='swap_with_partner',
                    set_rtl=False,
                    notes='Chart swaps with text body',
                ),
                'body': PlaceholderAction(
                    action='swap_with_partner',
                    set_rtl=True,
                    set_alignment='r',
                    notes='Text body swaps with chart',
                ),
            },
            freeform_action='mirror',
            mirror_master_elements=True,
            swap_columns=True,
            table_action='reverse_columns',
            chart_action='reverse_axes',
        )

        rules['txAndObj'] = LayoutTransformRules(
            layout_type='txAndObj',
            description='Text + Object — swap L/R',
            placeholder_rules={
                'title': _two_col_title,
                'body': PlaceholderAction(
                    action='swap_with_partner',
                    set_rtl=True,
                    set_alignment='r',
                    notes='Text swaps with object',
                ),
                'obj': PlaceholderAction(
                    action='swap_with_partner',
                    set_rtl=False,
                    notes='Object swaps with text',
                ),
            },
            freeform_action='mirror',
            mirror_master_elements=True,
            swap_columns=True,
            table_action='reverse_columns',
            chart_action='reverse_axes',
        )

        rules['objAndTx'] = LayoutTransformRules(
            layout_type='objAndTx',
            description='Object + Text — swap L/R',
            placeholder_rules={
                'title': _two_col_title,
                'obj': PlaceholderAction(
                    action='swap_with_partner',
                    set_rtl=False,
                    notes='Object swaps with text',
                ),
                'body': PlaceholderAction(
                    action='swap_with_partner',
                    set_rtl=True,
                    set_alignment='r',
                    notes='Text swaps with object',
                ),
            },
            freeform_action='mirror',
            mirror_master_elements=True,
            swap_columns=True,
            table_action='reverse_columns',
            chart_action='reverse_axes',
        )

        rules['txAndClipArt'] = LayoutTransformRules(
            layout_type='txAndClipArt',
            description='Text + Clip Art — swap L/R',
            placeholder_rules={
                'title': _two_col_title,
                'body': PlaceholderAction(
                    action='swap_with_partner',
                    set_rtl=True,
                    set_alignment='r',
                    notes='Text swaps with clip art',
                ),
                'clipArt': PlaceholderAction(
                    action='swap_with_partner',
                    set_rtl=False,
                    notes='Clip art swaps with text',
                ),
            },
            freeform_action='mirror',
            mirror_master_elements=True,
            swap_columns=True,
            table_action='reverse_columns',
            chart_action='reverse_axes',
        )

        rules['clipArtAndTx'] = LayoutTransformRules(
            layout_type='clipArtAndTx',
            description='Clip Art + Text — swap L/R',
            placeholder_rules={
                'title': _two_col_title,
                'clipArt': PlaceholderAction(
                    action='swap_with_partner',
                    set_rtl=False,
                    notes='Clip art swaps with text',
                ),
                'body': PlaceholderAction(
                    action='swap_with_partner',
                    set_rtl=True,
                    set_alignment='r',
                    notes='Text swaps with clip art',
                ),
            },
            freeform_action='mirror',
            mirror_master_elements=True,
            swap_columns=True,
            table_action='reverse_columns',
            chart_action='reverse_axes',
        )

        rules['txAndMedia'] = LayoutTransformRules(
            layout_type='txAndMedia',
            description='Text + Media — swap L/R',
            placeholder_rules={
                'title': _two_col_title,
                'body': PlaceholderAction(
                    action='swap_with_partner',
                    set_rtl=True,
                    set_alignment='r',
                    notes='Text swaps with media',
                ),
                'media': PlaceholderAction(
                    action='swap_with_partner',
                    set_rtl=False,
                    notes='Media swaps with text',
                ),
            },
            freeform_action='mirror',
            mirror_master_elements=True,
            swap_columns=True,
            table_action='reverse_columns',
            chart_action='reverse_axes',
        )

        rules['mediaAndTx'] = LayoutTransformRules(
            layout_type='mediaAndTx',
            description='Media + Text — swap L/R',
            placeholder_rules={
                'title': _two_col_title,
                'media': PlaceholderAction(
                    action='swap_with_partner',
                    set_rtl=False,
                    notes='Media swaps with text',
                ),
                'body': PlaceholderAction(
                    action='swap_with_partner',
                    set_rtl=True,
                    set_alignment='r',
                    notes='Text swaps with media',
                ),
            },
            freeform_action='mirror',
            mirror_master_elements=True,
            swap_columns=True,
            table_action='reverse_columns',
            chart_action='reverse_axes',
        )

        rules['picTx'] = LayoutTransformRules(
            layout_type='picTx',
            description='Picture + Caption — swap L/R',
            placeholder_rules={
                'title': _two_col_title,
                'pic': PlaceholderAction(
                    action='swap_with_partner',
                    set_rtl=False,
                    notes='Picture swaps side with body text',
                ),
                'body': PlaceholderAction(
                    action='swap_with_partner',
                    set_rtl=True,
                    set_alignment='r',
                    notes='Caption body swaps with picture',
                ),
            },
            freeform_action='mirror',
            mirror_master_elements=True,
            swap_columns=True,
            table_action='reverse_columns',
            chart_action='reverse_axes',
        )

        # ──────────────────────────────────────────────────────────────
        # CONTENT-STACKED LAYOUTS (obj + text vertically arranged)
        # ──────────────────────────────────────────────────────────────

        rules['objTx'] = LayoutTransformRules(
            layout_type='objTx',
            description='Object over Text (stacked)',
            placeholder_rules={
                'title': PlaceholderAction(
                    action='keep_position',
                    set_rtl=True,
                    set_alignment='r',
                    notes='Title right-aligns',
                ),
                'obj': PlaceholderAction(
                    action='keep_position',
                    set_rtl=False,
                    notes='Object stays in place (stacked above text)',
                ),
                'body': PlaceholderAction(
                    action='keep_position',
                    set_rtl=True,
                    set_alignment='r',
                    notes='Text body right-aligns below object',
                ),
            },
            freeform_action='mirror',
            mirror_master_elements=True,
            swap_columns=False,
            table_action='reverse_columns',
            chart_action='reverse_axes',
        )

        rules['txObj'] = LayoutTransformRules(
            layout_type='txObj',
            description='Text over Object (stacked)',
            placeholder_rules={
                'title': PlaceholderAction(
                    action='keep_position',
                    set_rtl=True,
                    set_alignment='r',
                    notes='Title right-aligns',
                ),
                'body': PlaceholderAction(
                    action='keep_position',
                    set_rtl=True,
                    set_alignment='r',
                    notes='Text body right-aligns above object',
                ),
                'obj': PlaceholderAction(
                    action='keep_position',
                    set_rtl=False,
                    notes='Object stays in place (stacked below text)',
                ),
            },
            freeform_action='mirror',
            mirror_master_elements=True,
            swap_columns=False,
            table_action='reverse_columns',
            chart_action='reverse_axes',
        )

        rules['objOverTx'] = LayoutTransformRules(
            layout_type='objOverTx',
            description='Object Over Text',
            placeholder_rules={
                'title': PlaceholderAction(
                    action='keep_position',
                    set_rtl=True,
                    set_alignment='r',
                    notes='Title right-aligns',
                ),
                'obj': PlaceholderAction(
                    action='keep_position',
                    set_rtl=False,
                    notes='Object keeps position above text',
                ),
                'body': PlaceholderAction(
                    action='keep_position',
                    set_rtl=True,
                    set_alignment='r',
                    notes='Text right-aligns below object',
                ),
            },
            freeform_action='mirror',
            mirror_master_elements=True,
            swap_columns=False,
            table_action='reverse_columns',
            chart_action='reverse_axes',
        )

        rules['txOverObj'] = LayoutTransformRules(
            layout_type='txOverObj',
            description='Text Over Object',
            placeholder_rules={
                'title': PlaceholderAction(
                    action='keep_position',
                    set_rtl=True,
                    set_alignment='r',
                    notes='Title right-aligns',
                ),
                'body': PlaceholderAction(
                    action='keep_position',
                    set_rtl=True,
                    set_alignment='r',
                    notes='Text right-aligns above object',
                ),
                'obj': PlaceholderAction(
                    action='keep_position',
                    set_rtl=False,
                    notes='Object keeps position below text',
                ),
            },
            freeform_action='mirror',
            mirror_master_elements=True,
            swap_columns=False,
            table_action='reverse_columns',
            chart_action='reverse_axes',
        )

        # ──────────────────────────────────────────────────────────────
        # MULTI-OBJECT LAYOUTS
        # ──────────────────────────────────────────────────────────────

        rules['twoObj'] = LayoutTransformRules(
            layout_type='twoObj',
            description='Two Objects — swap L/R',
            placeholder_rules={
                'title': _two_col_title,
                'obj': PlaceholderAction(
                    action='swap_with_partner',
                    set_rtl=False,
                    notes='Object placeholders swap L/R positions',
                ),
            },
            freeform_action='mirror',
            mirror_master_elements=True,
            swap_columns=True,
            table_action='reverse_columns',
            chart_action='reverse_axes',
        )

        rules['fourObj'] = LayoutTransformRules(
            layout_type='fourObj',
            description='Four Objects — mirror grid horizontally',
            placeholder_rules={
                'title': PlaceholderAction(
                    action='keep_position',
                    set_rtl=True,
                    set_alignment='r',
                    notes='Title right-aligns',
                ),
                'obj': PlaceholderAction(
                    action='mirror',
                    set_rtl=False,
                    mirror_x=True,
                    notes='Each object mirrors X position',
                ),
            },
            freeform_action='mirror',
            mirror_master_elements=True,
            swap_columns=True,
            table_action='reverse_columns',
            chart_action='reverse_axes',
        )

        rules['txAndTwoObj'] = LayoutTransformRules(
            layout_type='txAndTwoObj',
            description='Text + Two Objects — text keeps, objects swap',
            placeholder_rules={
                'title': _two_col_title,
                'body': PlaceholderAction(
                    action='keep_position',
                    set_rtl=True,
                    set_alignment='r',
                    notes='Text body right-aligns in place',
                ),
                'obj': PlaceholderAction(
                    action='swap_with_partner',
                    set_rtl=False,
                    notes='Object pair swaps L/R positions',
                ),
            },
            freeform_action='mirror',
            mirror_master_elements=True,
            swap_columns=True,
            table_action='reverse_columns',
            chart_action='reverse_axes',
        )

        rules['twoObjAndTx'] = LayoutTransformRules(
            layout_type='twoObjAndTx',
            description='Two Objects + Text — objects swap, text keeps',
            placeholder_rules={
                'title': _two_col_title,
                'obj': PlaceholderAction(
                    action='swap_with_partner',
                    set_rtl=False,
                    notes='Object pair swaps L/R positions',
                ),
                'body': PlaceholderAction(
                    action='keep_position',
                    set_rtl=True,
                    set_alignment='r',
                    notes='Text body right-aligns in place',
                ),
            },
            freeform_action='mirror',
            mirror_master_elements=True,
            swap_columns=True,
            table_action='reverse_columns',
            chart_action='reverse_axes',
        )

        rules['twoObjOverTx'] = LayoutTransformRules(
            layout_type='twoObjOverTx',
            description='Two Objects Over Text — objects swap, text under',
            placeholder_rules={
                'title': _two_col_title,
                'obj': PlaceholderAction(
                    action='swap_with_partner',
                    set_rtl=False,
                    notes='Object pair swaps L/R over text',
                ),
                'body': PlaceholderAction(
                    action='keep_position',
                    set_rtl=True,
                    set_alignment='r',
                    notes='Text body right-aligns below objects',
                ),
            },
            freeform_action='mirror',
            mirror_master_elements=True,
            swap_columns=True,
            table_action='reverse_columns',
            chart_action='reverse_axes',
        )

        rules['twoTxTwoObj'] = LayoutTransformRules(
            layout_type='twoTxTwoObj',
            description='Two Text + Two Objects — all swap L/R',
            placeholder_rules={
                'title': _two_col_title,
                'body': PlaceholderAction(
                    action='swap_with_partner',
                    set_rtl=True,
                    set_alignment='r',
                    notes='Text columns swap L/R',
                ),
                'obj': PlaceholderAction(
                    action='swap_with_partner',
                    set_rtl=False,
                    notes='Object columns swap L/R',
                ),
            },
            freeform_action='mirror',
            mirror_master_elements=True,
            swap_columns=True,
            table_action='reverse_columns',
            chart_action='reverse_axes',
        )

        rules['twoObjAndObj'] = LayoutTransformRules(
            layout_type='twoObjAndObj',
            description='Two Objects + Object — mirror grid',
            placeholder_rules={
                'title': _two_col_title,
                'obj': PlaceholderAction(
                    action='mirror',
                    set_rtl=False,
                    mirror_x=True,
                    notes='All object placeholders mirror X',
                ),
            },
            freeform_action='mirror',
            mirror_master_elements=True,
            swap_columns=True,
            table_action='reverse_columns',
            chart_action='reverse_axes',
        )

        rules['objAndTwoObj'] = LayoutTransformRules(
            layout_type='objAndTwoObj',
            description='Object + Two Objects — mirror grid',
            placeholder_rules={
                'title': _two_col_title,
                'obj': PlaceholderAction(
                    action='mirror',
                    set_rtl=False,
                    mirror_x=True,
                    notes='All object placeholders mirror X',
                ),
            },
            freeform_action='mirror',
            mirror_master_elements=True,
            swap_columns=True,
            table_action='reverse_columns',
            chart_action='reverse_axes',
        )

        # ──────────────────────────────────────────────────────────────
        # SPECIALISED CONTENT LAYOUTS
        # ──────────────────────────────────────────────────────────────

        rules['tbl'] = LayoutTransformRules(
            layout_type='tbl',
            description='Title + Table',
            placeholder_rules={
                'title': PlaceholderAction(
                    action='keep_position',
                    set_rtl=True,
                    set_alignment='r',
                    notes='Title right-aligns',
                ),
                'tbl': PlaceholderAction(
                    action='reverse_columns',
                    set_rtl=True,
                    notes='Table columns reverse for RTL reading',
                ),
            },
            freeform_action='mirror',
            mirror_master_elements=True,
            swap_columns=False,
            table_action='reverse_columns',
            chart_action='reverse_axes',
        )

        rules['chart'] = LayoutTransformRules(
            layout_type='chart',
            description='Title + Chart',
            placeholder_rules={
                'title': PlaceholderAction(
                    action='keep_position',
                    set_rtl=True,
                    set_alignment='r',
                    notes='Title right-aligns',
                ),
                'chart': PlaceholderAction(
                    action='reverse_axes',
                    set_rtl=False,
                    notes='Chart reverses category axis',
                ),
            },
            freeform_action='mirror',
            mirror_master_elements=True,
            swap_columns=False,
            table_action='reverse_columns',
            chart_action='reverse_axes',
        )

        rules['dgm'] = LayoutTransformRules(
            layout_type='dgm',
            description='Title + Diagram (SmartArt)',
            placeholder_rules={
                'title': PlaceholderAction(
                    action='keep_position',
                    set_rtl=True,
                    set_alignment='r',
                    notes='Title right-aligns',
                ),
                'dgm': PlaceholderAction(
                    action='keep_position',
                    set_rtl=True,
                    notes='Diagram keeps position; RTL on text nodes',
                ),
            },
            freeform_action='mirror',
            mirror_master_elements=True,
            swap_columns=False,
            table_action='reverse_columns',
            chart_action='reverse_axes',
        )

        # ──────────────────────────────────────────────────────────────
        # VERTICAL TEXT LAYOUTS
        # ──────────────────────────────────────────────────────────────

        rules['vertTx'] = LayoutTransformRules(
            layout_type='vertTx',
            description='Vertical Text',
            placeholder_rules={
                'title': PlaceholderAction(
                    action='keep_position',
                    set_rtl=True,
                    set_alignment='r',
                    notes='Title right-aligns',
                ),
                'body': PlaceholderAction(
                    action='keep_position',
                    set_rtl=True,
                    set_alignment='r',
                    notes='Vertical body: set RTL direction',
                ),
            },
            freeform_action='mirror',
            mirror_master_elements=True,
            swap_columns=False,
            table_action='reverse_columns',
            chart_action='reverse_axes',
        )

        rules['vertTitleAndTx'] = LayoutTransformRules(
            layout_type='vertTitleAndTx',
            description='Vertical Title + Vertical Text',
            placeholder_rules={
                'title': PlaceholderAction(
                    action='mirror',
                    set_rtl=True,
                    set_alignment='r',
                    mirror_x=True,
                    notes='Vertical title mirrors to opposite side',
                ),
                'body': PlaceholderAction(
                    action='mirror',
                    set_rtl=True,
                    set_alignment='r',
                    mirror_x=True,
                    notes='Vertical body mirrors to opposite side',
                ),
            },
            freeform_action='mirror',
            mirror_master_elements=True,
            swap_columns=True,
            table_action='reverse_columns',
            chart_action='reverse_axes',
        )

        rules['vertTitleAndTxOverChart'] = LayoutTransformRules(
            layout_type='vertTitleAndTxOverChart',
            description='Vertical Title + Text Over Chart',
            placeholder_rules={
                'title': PlaceholderAction(
                    action='mirror',
                    set_rtl=True,
                    set_alignment='r',
                    mirror_x=True,
                    notes='Vertical title mirrors',
                ),
                'body': PlaceholderAction(
                    action='keep_position',
                    set_rtl=True,
                    set_alignment='r',
                    notes='Body text right-aligns',
                ),
                'chart': PlaceholderAction(
                    action='reverse_axes',
                    set_rtl=False,
                    notes='Chart reverses axes',
                ),
            },
            freeform_action='mirror',
            mirror_master_elements=True,
            swap_columns=False,
            table_action='reverse_columns',
            chart_action='reverse_axes',
        )

        # ──────────────────────────────────────────────────────────────
        # FALLBACK / CUSTOM
        # ──────────────────────────────────────────────────────────────

        rules['cust'] = LayoutTransformRules(
            layout_type='cust',
            description='Custom / Unknown layout — conservative defaults',
            placeholder_rules={
                'title': PlaceholderAction(
                    action='keep_position',
                    set_rtl=True,
                    set_alignment='r',
                    notes='Default: title right-aligns',
                ),
                'ctrTitle': PlaceholderAction(
                    action='keep_centered',
                    set_rtl=True,
                    set_alignment='ctr',
                    notes='Default: center title stays centered',
                ),
                'subTitle': PlaceholderAction(
                    action='right_align',
                    set_rtl=True,
                    set_alignment='r',
                    notes='Default: subtitle right-aligns',
                ),
                'body': PlaceholderAction(
                    action='keep_position',
                    set_rtl=True,
                    set_alignment='r',
                    notes='Default: body right-aligns',
                ),
                'obj': PlaceholderAction(
                    action='keep_position',
                    set_rtl=False,
                    notes='Default: object keeps position',
                ),
            },
            freeform_action='mirror',
            mirror_master_elements=True,
            swap_columns=False,
            table_action='reverse_columns',
            chart_action='reverse_axes',
        )

        return rules
