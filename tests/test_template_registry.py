"""
test_template_registry.py — Comprehensive unit tests for template_registry.py.

Tests cover:
- ARABIC_FONT_MAP correctness and coverage
- get_arabic_font fallback behaviour
- PlaceholderAction and LayoutTransformRules data classes
- TemplateRegistry default rule coverage (all 36 layout types)
- TemplateRegistry.get_rules for known and unknown types
- TemplateRegistry.get_placeholder_action lookup precedence
- TemplateRegistry.get_freeform_action
- TemplateRegistry.register_custom_rule
- TemplateRegistry.list_layout_types
- Specific rule correctness for critical layout types
"""

from __future__ import annotations

import unittest

from slidearabi.template_registry import (
    ARABIC_FONT_MAP,
    LayoutTransformRules,
    PlaceholderAction,
    TemplateRegistry,
    get_arabic_font,
)


# ── Standard slide dimensions (widescreen 16:9) ────────────────────────────
SLIDE_WIDTH = 12192000   # 13.333 inches
SLIDE_HEIGHT = 6858000   # 7.5 inches


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Tests: ARABIC_FONT_MAP
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


class TestArabicFontMap(unittest.TestCase):
    """Tests for the ARABIC_FONT_MAP dictionary."""

    def test_is_dict(self):
        self.assertIsInstance(ARABIC_FONT_MAP, dict)

    def test_has_minimum_entries(self):
        """Map should have at least 20 font mappings."""
        self.assertGreaterEqual(len(ARABIC_FONT_MAP), 20)

    def test_calibri_maps_to_itself(self):
        """Calibri already supports Arabic — maps to itself."""
        self.assertEqual(ARABIC_FONT_MAP['Calibri'], 'Calibri')

    def test_arial_maps_to_itself(self):
        self.assertEqual(ARABIC_FONT_MAP['Arial'], 'Arial')

    def test_times_new_roman_maps_to_itself(self):
        self.assertEqual(ARABIC_FONT_MAP['Times New Roman'], 'Times New Roman')

    def test_cambria_to_sakkal_majalla(self):
        self.assertEqual(ARABIC_FONT_MAP['Cambria'], 'Sakkal Majalla')

    def test_georgia_to_sakkal_majalla(self):
        self.assertEqual(ARABIC_FONT_MAP['Georgia'], 'Sakkal Majalla')

    def test_garamond_to_traditional_arabic(self):
        self.assertEqual(ARABIC_FONT_MAP['Garamond'], 'Traditional Arabic')

    def test_century_gothic_to_dubai(self):
        self.assertEqual(ARABIC_FONT_MAP['Century Gothic'], 'Dubai')

    def test_helvetica_to_arial(self):
        self.assertEqual(ARABIC_FONT_MAP['Helvetica'], 'Arial')

    def test_impact_to_arial_black(self):
        self.assertEqual(ARABIC_FONT_MAP['Impact'], 'Arial Black')

    def test_comic_sans_to_tahoma(self):
        self.assertEqual(ARABIC_FONT_MAP['Comic Sans MS'], 'Tahoma')

    def test_verdana_to_tahoma(self):
        self.assertEqual(ARABIC_FONT_MAP['Verdana'], 'Tahoma')

    def test_consolas_to_courier_new(self):
        self.assertEqual(ARABIC_FONT_MAP['Consolas'], 'Courier New')

    def test_all_values_are_strings(self):
        """Every key and value in the map should be a non-empty string."""
        for k, v in ARABIC_FONT_MAP.items():
            self.assertIsInstance(k, str)
            self.assertIsInstance(v, str)
            self.assertTrue(len(k) > 0)
            self.assertTrue(len(v) > 0)


class TestGetArabicFont(unittest.TestCase):
    """Tests for the get_arabic_font convenience function."""

    def test_known_font(self):
        self.assertEqual(get_arabic_font('Cambria'), 'Sakkal Majalla')

    def test_unknown_font_fallback(self):
        """Unknown fonts fall back to 'Calibri'."""
        self.assertEqual(get_arabic_font('MyWeirdFont'), 'Calibri')

    def test_empty_string_fallback(self):
        self.assertEqual(get_arabic_font(''), 'Calibri')


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Tests: PlaceholderAction
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


class TestPlaceholderAction(unittest.TestCase):
    """Tests for PlaceholderAction data class."""

    def test_default_values(self):
        action = PlaceholderAction(action='keep_position')
        self.assertEqual(action.action, 'keep_position')
        self.assertTrue(action.set_rtl)
        self.assertIsNone(action.set_alignment)
        self.assertIsNone(action.swap_partner_idx)
        self.assertFalse(action.mirror_x)
        self.assertEqual(action.notes, '')

    def test_full_construction(self):
        action = PlaceholderAction(
            action='swap_with_partner',
            set_rtl=True,
            set_alignment='r',
            swap_partner_idx=2,
            mirror_x=False,
            notes='Test note',
        )
        self.assertEqual(action.action, 'swap_with_partner')
        self.assertEqual(action.set_alignment, 'r')
        self.assertEqual(action.swap_partner_idx, 2)
        self.assertEqual(action.notes, 'Test note')

    def test_mirror_action(self):
        action = PlaceholderAction(action='mirror', mirror_x=True)
        self.assertTrue(action.mirror_x)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Tests: LayoutTransformRules
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


class TestLayoutTransformRules(unittest.TestCase):
    """Tests for LayoutTransformRules data class."""

    def test_default_values(self):
        rules = LayoutTransformRules(
            layout_type='test',
            description='Test layout',
        )
        self.assertEqual(rules.layout_type, 'test')
        self.assertEqual(rules.placeholder_rules, {})
        self.assertEqual(rules.freeform_action, 'mirror')
        self.assertTrue(rules.mirror_master_elements)
        self.assertFalse(rules.swap_columns)
        self.assertEqual(rules.table_action, 'reverse_columns')
        self.assertEqual(rules.chart_action, 'reverse_axes')

    def test_custom_values(self):
        rules = LayoutTransformRules(
            layout_type='twoColTx',
            description='Two Column',
            freeform_action='keep',
            swap_columns=True,
            table_action='rtl_only',
            chart_action='keep',
        )
        self.assertTrue(rules.swap_columns)
        self.assertEqual(rules.freeform_action, 'keep')
        self.assertEqual(rules.table_action, 'rtl_only')
        self.assertEqual(rules.chart_action, 'keep')


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Tests: TemplateRegistry — default rules coverage
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


class TestDefaultRulesCoverage(unittest.TestCase):
    """Ensure all 36 OOXML layout types have rules."""

    def setUp(self):
        self.registry = TemplateRegistry(SLIDE_WIDTH, SLIDE_HEIGHT)

    def test_all_36_standard_types_have_rules(self):
        """Every ST_SlideLayoutType value is covered by the registry."""
        all_types = [
            'title', 'tx', 'twoColTx', 'tbl', 'txAndChart', 'chartAndTx',
            'dgm', 'chart', 'txAndClipArt', 'clipArtAndTx', 'titleOnly',
            'blank', 'txAndObj', 'objAndTx', 'objOnly', 'obj', 'txAndMedia',
            'mediaAndTx', 'objTx', 'txObj', 'objOverTx', 'txOverObj',
            'txAndTwoObj', 'twoObjAndTx', 'twoObjOverTx', 'fourObj',
            'twoTxTwoObj', 'twoObjAndObj', 'secHead', 'twoObj',
            'objAndTwoObj', 'picTx', 'vertTx', 'vertTitleAndTx',
            'vertTitleAndTxOverChart', 'cust',
        ]
        for lt in all_types:
            rules = self.registry.get_rules(lt)
            self.assertIsInstance(rules, LayoutTransformRules, f'No rules for {lt}')
            self.assertEqual(rules.layout_type, lt, f'Rule type mismatch for {lt}')

    def test_rules_count_is_36(self):
        """Internal rules dict has at least 36 entries."""
        self.assertGreaterEqual(len(self.registry._rules), 36)

    def test_list_layout_types(self):
        """list_layout_types returns sorted list."""
        types = self.registry.list_layout_types()
        self.assertIsInstance(types, list)
        self.assertGreaterEqual(len(types), 36)
        self.assertEqual(types, sorted(types))


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Tests: get_rules
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


class TestGetRules(unittest.TestCase):
    """Tests for TemplateRegistry.get_rules."""

    def setUp(self):
        self.registry = TemplateRegistry(SLIDE_WIDTH, SLIDE_HEIGHT)

    def test_known_type_returns_correct_rules(self):
        rules = self.registry.get_rules('title')
        self.assertEqual(rules.layout_type, 'title')
        self.assertIn('ctrTitle', rules.placeholder_rules)

    def test_unknown_type_falls_back_to_cust(self):
        """Unknown type gets the 'cust' (custom) fallback rules."""
        rules = self.registry.get_rules('totallyMadeUp')
        self.assertEqual(rules.layout_type, 'cust')

    def test_blank_has_no_placeholder_rules(self):
        rules = self.registry.get_rules('blank')
        self.assertEqual(rules.placeholder_rules, {})

    def test_custom_rules_take_precedence(self):
        """Custom rules override built-in defaults."""
        custom = LayoutTransformRules(
            layout_type='title',
            description='Custom title',
            freeform_action='keep',
        )
        self.registry.register_custom_rule('title', custom)
        rules = self.registry.get_rules('title')
        self.assertEqual(rules.description, 'Custom title')
        self.assertEqual(rules.freeform_action, 'keep')


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Tests: get_placeholder_action
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


class TestGetPlaceholderAction(unittest.TestCase):
    """Tests for TemplateRegistry.get_placeholder_action."""

    def setUp(self):
        self.registry = TemplateRegistry(SLIDE_WIDTH, SLIDE_HEIGHT)

    def test_title_slide_ctrTitle(self):
        action = self.registry.get_placeholder_action('title', 'ctrTitle', 0)
        self.assertEqual(action.action, 'keep_centered')
        self.assertTrue(action.set_rtl)
        self.assertEqual(action.set_alignment, 'ctr')

    def test_title_slide_subTitle(self):
        action = self.registry.get_placeholder_action('title', 'subTitle', 1)
        self.assertEqual(action.action, 'right_align')
        self.assertTrue(action.set_rtl)
        self.assertEqual(action.set_alignment, 'r')

    def test_tx_title(self):
        action = self.registry.get_placeholder_action('tx', 'title', 0)
        self.assertEqual(action.action, 'keep_position')
        self.assertTrue(action.set_rtl)
        self.assertEqual(action.set_alignment, 'r')

    def test_tx_body(self):
        action = self.registry.get_placeholder_action('tx', 'body', 1)
        self.assertEqual(action.action, 'keep_position')
        self.assertTrue(action.set_rtl)
        self.assertEqual(action.set_alignment, 'r')

    def test_twoColTx_body(self):
        action = self.registry.get_placeholder_action('twoColTx', 'body', 1)
        self.assertEqual(action.action, 'swap_with_partner')
        self.assertTrue(action.set_rtl)

    def test_chart_layout_chart_placeholder(self):
        action = self.registry.get_placeholder_action('chart', 'chart', 1)
        self.assertEqual(action.action, 'reverse_axes')
        self.assertFalse(action.set_rtl)

    def test_tbl_layout_tbl_placeholder(self):
        action = self.registry.get_placeholder_action('tbl', 'tbl', 1)
        self.assertEqual(action.action, 'reverse_columns')
        self.assertTrue(action.set_rtl)

    def test_unknown_placeholder_returns_default(self):
        """Unknown placeholder type gets default keep_position + RTL."""
        action = self.registry.get_placeholder_action('tx', 'weirdType', 99)
        self.assertEqual(action.action, 'keep_position')
        self.assertTrue(action.set_rtl)
        self.assertEqual(action.set_alignment, 'r')

    def test_blank_layout_unknown_placeholder(self):
        """Blank layout + any placeholder → default action."""
        action = self.registry.get_placeholder_action('blank', 'body', 1)
        self.assertEqual(action.action, 'keep_position')
        self.assertTrue(action.set_rtl)

    def test_idx_key_takes_precedence(self):
        """idx_N key in placeholder_rules takes precedence over type key."""
        # Register a custom rule with idx-specific action
        custom = LayoutTransformRules(
            layout_type='test_idx',
            description='Test idx precedence',
            placeholder_rules={
                'body': PlaceholderAction(
                    action='keep_position',
                    notes='by type',
                ),
                'idx_5': PlaceholderAction(
                    action='mirror',
                    mirror_x=True,
                    notes='by idx',
                ),
            },
        )
        self.registry.register_custom_rule('test_idx', custom)

        # idx_5 should win over body
        action = self.registry.get_placeholder_action('test_idx', 'body', 5)
        self.assertEqual(action.action, 'mirror')
        self.assertEqual(action.notes, 'by idx')

        # idx_1 should fall back to body
        action = self.registry.get_placeholder_action('test_idx', 'body', 1)
        self.assertEqual(action.action, 'keep_position')
        self.assertEqual(action.notes, 'by type')


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Tests: get_freeform_action
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


class TestGetFreeformAction(unittest.TestCase):
    """Tests for TemplateRegistry.get_freeform_action."""

    def setUp(self):
        self.registry = TemplateRegistry(SLIDE_WIDTH, SLIDE_HEIGHT)

    def test_title_freeform_mirror(self):
        self.assertEqual(self.registry.get_freeform_action('title'), 'mirror')

    def test_secHead_freeform_keep(self):
        self.assertEqual(self.registry.get_freeform_action('secHead'), 'keep')

    def test_tx_freeform_mirror(self):
        self.assertEqual(self.registry.get_freeform_action('tx'), 'mirror')

    def test_blank_freeform_mirror(self):
        self.assertEqual(self.registry.get_freeform_action('blank'), 'mirror')

    def test_unknown_fallback_to_cust(self):
        """Unknown layout freeform action comes from 'cust' rules."""
        self.assertEqual(self.registry.get_freeform_action('unknownXYZ'), 'mirror')


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Tests: register_custom_rule
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


class TestRegisterCustomRule(unittest.TestCase):
    """Tests for TemplateRegistry.register_custom_rule."""

    def setUp(self):
        self.registry = TemplateRegistry(SLIDE_WIDTH, SLIDE_HEIGHT)

    def test_register_new_layout(self):
        """Custom rules can be added for new layout names."""
        custom = LayoutTransformRules(
            layout_type='myEnterprise',
            description='Custom enterprise layout',
            freeform_action='keep',
        )
        self.registry.register_custom_rule('myEnterprise', custom)
        rules = self.registry.get_rules('myEnterprise')
        self.assertEqual(rules.layout_type, 'myEnterprise')
        self.assertEqual(rules.freeform_action, 'keep')

    def test_custom_rule_in_list(self):
        """Custom rules appear in list_layout_types."""
        custom = LayoutTransformRules(
            layout_type='acme_layout',
            description='ACME Corp special',
        )
        self.registry.register_custom_rule('acme_layout', custom)
        types = self.registry.list_layout_types()
        self.assertIn('acme_layout', types)

    def test_override_existing(self):
        """Custom rules can override built-in rules."""
        original = self.registry.get_rules('blank')
        self.assertEqual(original.freeform_action, 'mirror')

        custom = LayoutTransformRules(
            layout_type='blank',
            description='Override blank',
            freeform_action='keep',
        )
        self.registry.register_custom_rule('blank', custom)

        overridden = self.registry.get_rules('blank')
        self.assertEqual(overridden.freeform_action, 'keep')


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Tests: Specific layout type rules correctness
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


class TestTitleSlideRules(unittest.TestCase):
    """Detailed tests for 'title' layout rules."""

    def setUp(self):
        self.registry = TemplateRegistry(SLIDE_WIDTH, SLIDE_HEIGHT)
        self.rules = self.registry.get_rules('title')

    def test_description(self):
        self.assertIn('Title', self.rules.description)

    def test_not_swap_columns(self):
        self.assertFalse(self.rules.swap_columns)

    def test_mirror_master_false(self):
        """Title slides typically don't mirror master elements."""
        self.assertFalse(self.rules.mirror_master_elements)

    def test_ctrTitle_centered(self):
        action = self.rules.placeholder_rules['ctrTitle']
        self.assertEqual(action.action, 'keep_centered')
        self.assertEqual(action.set_alignment, 'ctr')

    def test_subTitle_right(self):
        action = self.rules.placeholder_rules['subTitle']
        self.assertEqual(action.action, 'right_align')
        self.assertEqual(action.set_alignment, 'r')


class TestTwoColTxRules(unittest.TestCase):
    """Detailed tests for 'twoColTx' layout rules."""

    def setUp(self):
        self.registry = TemplateRegistry(SLIDE_WIDTH, SLIDE_HEIGHT)
        self.rules = self.registry.get_rules('twoColTx')

    def test_swap_columns_true(self):
        self.assertTrue(self.rules.swap_columns)

    def test_body_swap(self):
        action = self.rules.placeholder_rules['body']
        self.assertEqual(action.action, 'swap_with_partner')
        self.assertTrue(action.set_rtl)
        self.assertEqual(action.set_alignment, 'r')

    def test_title_right_align(self):
        action = self.rules.placeholder_rules['title']
        self.assertTrue(action.set_rtl)
        self.assertEqual(action.set_alignment, 'r')


class TestTxAndChartRules(unittest.TestCase):
    """Detailed tests for 'txAndChart' layout rules."""

    def setUp(self):
        self.registry = TemplateRegistry(SLIDE_WIDTH, SLIDE_HEIGHT)
        self.rules = self.registry.get_rules('txAndChart')

    def test_swap_columns(self):
        self.assertTrue(self.rules.swap_columns)

    def test_body_swap(self):
        action = self.rules.placeholder_rules['body']
        self.assertEqual(action.action, 'swap_with_partner')
        self.assertTrue(action.set_rtl)

    def test_chart_swap(self):
        action = self.rules.placeholder_rules['chart']
        self.assertEqual(action.action, 'swap_with_partner')
        self.assertFalse(action.set_rtl)

    def test_chart_action_reverse_axes(self):
        self.assertEqual(self.rules.chart_action, 'reverse_axes')


class TestObjAndTxRules(unittest.TestCase):
    """Detailed tests for 'objAndTx' layout rules."""

    def setUp(self):
        self.registry = TemplateRegistry(SLIDE_WIDTH, SLIDE_HEIGHT)
        self.rules = self.registry.get_rules('objAndTx')

    def test_obj_swap(self):
        action = self.rules.placeholder_rules['obj']
        self.assertEqual(action.action, 'swap_with_partner')

    def test_body_swap(self):
        action = self.rules.placeholder_rules['body']
        self.assertEqual(action.action, 'swap_with_partner')
        self.assertTrue(action.set_rtl)


class TestPicTxRules(unittest.TestCase):
    """Detailed tests for 'picTx' layout rules."""

    def setUp(self):
        self.registry = TemplateRegistry(SLIDE_WIDTH, SLIDE_HEIGHT)
        self.rules = self.registry.get_rules('picTx')

    def test_pic_swap(self):
        action = self.rules.placeholder_rules['pic']
        self.assertEqual(action.action, 'swap_with_partner')

    def test_body_swap(self):
        action = self.rules.placeholder_rules['body']
        self.assertEqual(action.action, 'swap_with_partner')
        self.assertTrue(action.set_rtl)


class TestSecHeadRules(unittest.TestCase):
    """Detailed tests for 'secHead' layout rules."""

    def setUp(self):
        self.registry = TemplateRegistry(SLIDE_WIDTH, SLIDE_HEIGHT)
        self.rules = self.registry.get_rules('secHead')

    def test_freeform_keep(self):
        """Section headers keep freeform elements."""
        self.assertEqual(self.rules.freeform_action, 'keep')

    def test_title_right(self):
        action = self.rules.placeholder_rules['title']
        self.assertEqual(action.action, 'right_align')

    def test_body_right(self):
        action = self.rules.placeholder_rules['body']
        self.assertEqual(action.action, 'right_align')


class TestCustRules(unittest.TestCase):
    """Tests for 'cust' (custom/fallback) layout rules."""

    def setUp(self):
        self.registry = TemplateRegistry(SLIDE_WIDTH, SLIDE_HEIGHT)
        self.rules = self.registry.get_rules('cust')

    def test_has_default_placeholder_rules(self):
        """Cust rules should have reasonable defaults for common types."""
        self.assertIn('title', self.rules.placeholder_rules)
        self.assertIn('body', self.rules.placeholder_rules)
        self.assertIn('ctrTitle', self.rules.placeholder_rules)

    def test_freeform_mirror(self):
        self.assertEqual(self.rules.freeform_action, 'mirror')

    def test_table_reverse(self):
        self.assertEqual(self.rules.table_action, 'reverse_columns')

    def test_chart_reverse(self):
        self.assertEqual(self.rules.chart_action, 'reverse_axes')

    def test_title_right_align(self):
        action = self.rules.placeholder_rules['title']
        self.assertEqual(action.set_alignment, 'r')
        self.assertTrue(action.set_rtl)


class TestChartLayoutRules(unittest.TestCase):
    """Tests for 'chart' layout rules."""

    def setUp(self):
        self.registry = TemplateRegistry(SLIDE_WIDTH, SLIDE_HEIGHT)
        self.rules = self.registry.get_rules('chart')

    def test_chart_placeholder_reverses_axes(self):
        action = self.rules.placeholder_rules['chart']
        self.assertEqual(action.action, 'reverse_axes')

    def test_not_swap_columns(self):
        self.assertFalse(self.rules.swap_columns)


class TestDgmLayoutRules(unittest.TestCase):
    """Tests for 'dgm' (diagram) layout rules."""

    def setUp(self):
        self.registry = TemplateRegistry(SLIDE_WIDTH, SLIDE_HEIGHT)
        self.rules = self.registry.get_rules('dgm')

    def test_dgm_placeholder_keep_position(self):
        action = self.rules.placeholder_rules['dgm']
        self.assertEqual(action.action, 'keep_position')
        self.assertTrue(action.set_rtl)


class TestVertTxRules(unittest.TestCase):
    """Tests for vertical text layout rules."""

    def setUp(self):
        self.registry = TemplateRegistry(SLIDE_WIDTH, SLIDE_HEIGHT)

    def test_vertTx_body(self):
        rules = self.registry.get_rules('vertTx')
        action = rules.placeholder_rules['body']
        self.assertEqual(action.action, 'keep_position')
        self.assertTrue(action.set_rtl)

    def test_vertTitleAndTx_mirror(self):
        rules = self.registry.get_rules('vertTitleAndTx')
        action = rules.placeholder_rules['title']
        self.assertEqual(action.action, 'mirror')
        self.assertTrue(action.mirror_x)


class TestStackedLayoutRules(unittest.TestCase):
    """Tests for stacked (obj over text, text over obj) layout rules."""

    def setUp(self):
        self.registry = TemplateRegistry(SLIDE_WIDTH, SLIDE_HEIGHT)

    def test_objTx_body_keeps_position(self):
        rules = self.registry.get_rules('objTx')
        action = rules.placeholder_rules['body']
        self.assertEqual(action.action, 'keep_position')
        self.assertTrue(action.set_rtl)

    def test_objTx_obj_keeps_position(self):
        rules = self.registry.get_rules('objTx')
        action = rules.placeholder_rules['obj']
        self.assertEqual(action.action, 'keep_position')
        self.assertFalse(action.set_rtl)

    def test_txObj_body_keeps_position(self):
        rules = self.registry.get_rules('txObj')
        action = rules.placeholder_rules['body']
        self.assertEqual(action.action, 'keep_position')

    def test_objOverTx_not_swap(self):
        rules = self.registry.get_rules('objOverTx')
        self.assertFalse(rules.swap_columns)

    def test_txOverObj_not_swap(self):
        rules = self.registry.get_rules('txOverObj')
        self.assertFalse(rules.swap_columns)


class TestMultiObjectLayoutRules(unittest.TestCase):
    """Tests for multi-object layout rules."""

    def setUp(self):
        self.registry = TemplateRegistry(SLIDE_WIDTH, SLIDE_HEIGHT)

    def test_fourObj_mirror(self):
        rules = self.registry.get_rules('fourObj')
        action = rules.placeholder_rules['obj']
        self.assertEqual(action.action, 'mirror')
        self.assertTrue(action.mirror_x)

    def test_twoObj_swap(self):
        rules = self.registry.get_rules('twoObj')
        action = rules.placeholder_rules['obj']
        self.assertEqual(action.action, 'swap_with_partner')

    def test_txAndTwoObj_body_keep(self):
        rules = self.registry.get_rules('txAndTwoObj')
        body_action = rules.placeholder_rules['body']
        self.assertEqual(body_action.action, 'keep_position')

    def test_txAndTwoObj_obj_swap(self):
        rules = self.registry.get_rules('txAndTwoObj')
        obj_action = rules.placeholder_rules['obj']
        self.assertEqual(obj_action.action, 'swap_with_partner')

    def test_twoObjAndObj_mirror(self):
        rules = self.registry.get_rules('twoObjAndObj')
        action = rules.placeholder_rules['obj']
        self.assertEqual(action.action, 'mirror')

    def test_objAndTwoObj_mirror(self):
        rules = self.registry.get_rules('objAndTwoObj')
        action = rules.placeholder_rules['obj']
        self.assertEqual(action.action, 'mirror')


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Tests: All two-column layouts have swap_columns=True
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


class TestSwapColumnsConsistency(unittest.TestCase):
    """Verify that all L/R side-by-side layouts have swap_columns=True."""

    def setUp(self):
        self.registry = TemplateRegistry(SLIDE_WIDTH, SLIDE_HEIGHT)

    def test_two_column_layouts_swap(self):
        expected_swap = [
            'twoColTx', 'txAndChart', 'chartAndTx', 'txAndObj', 'objAndTx',
            'txAndClipArt', 'clipArtAndTx', 'txAndMedia', 'mediaAndTx',
            'picTx', 'twoObj', 'txAndTwoObj', 'twoObjAndTx', 'twoObjOverTx',
            'twoTxTwoObj', 'twoObjAndObj', 'objAndTwoObj',
        ]
        for lt in expected_swap:
            rules = self.registry.get_rules(lt)
            self.assertTrue(
                rules.swap_columns,
                f'{lt} should have swap_columns=True',
            )

    def test_non_column_layouts_dont_swap(self):
        no_swap = ['title', 'secHead', 'tx', 'titleOnly', 'blank', 'obj',
                    'objOnly', 'tbl', 'chart', 'dgm',
                    'objTx', 'txObj', 'objOverTx', 'txOverObj', 'vertTx']
        for lt in no_swap:
            rules = self.registry.get_rules(lt)
            self.assertFalse(
                rules.swap_columns,
                f'{lt} should have swap_columns=False',
            )


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Tests: All title placeholders set RTL
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


class TestAllTitlesSetRTL(unittest.TestCase):
    """Verify that every layout's title placeholder has set_rtl=True."""

    def setUp(self):
        self.registry = TemplateRegistry(SLIDE_WIDTH, SLIDE_HEIGHT)

    def test_title_rtl(self):
        for lt in self.registry.list_layout_types():
            rules = self.registry.get_rules(lt)
            for ph_type in ('title', 'ctrTitle'):
                if ph_type in rules.placeholder_rules:
                    action = rules.placeholder_rules[ph_type]
                    self.assertTrue(
                        action.set_rtl,
                        f'{lt} / {ph_type} should have set_rtl=True',
                    )


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Tests: All text body placeholders set RTL
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


class TestAllBodiesSetRTL(unittest.TestCase):
    """Body placeholders in all layouts should have set_rtl=True."""

    def setUp(self):
        self.registry = TemplateRegistry(SLIDE_WIDTH, SLIDE_HEIGHT)

    def test_body_rtl(self):
        for lt in self.registry.list_layout_types():
            rules = self.registry.get_rules(lt)
            if 'body' in rules.placeholder_rules:
                action = rules.placeholder_rules['body']
                self.assertTrue(
                    action.set_rtl,
                    f'{lt} / body should have set_rtl=True',
                )


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Tests: Table and chart actions
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


class TestTableChartActions(unittest.TestCase):
    """All layouts should have reasonable table/chart defaults."""

    def setUp(self):
        self.registry = TemplateRegistry(SLIDE_WIDTH, SLIDE_HEIGHT)

    def test_all_have_table_action(self):
        for lt in self.registry.list_layout_types():
            rules = self.registry.get_rules(lt)
            self.assertIn(
                rules.table_action,
                ('reverse_columns', 'keep', 'rtl_only'),
                f'{lt} has invalid table_action: {rules.table_action}',
            )

    def test_all_have_chart_action(self):
        for lt in self.registry.list_layout_types():
            rules = self.registry.get_rules(lt)
            self.assertIn(
                rules.chart_action,
                ('reverse_axes', 'keep', 'mirror_legend'),
                f'{lt} has invalid chart_action: {rules.chart_action}',
            )

    def test_all_have_valid_freeform_action(self):
        for lt in self.registry.list_layout_types():
            rules = self.registry.get_rules(lt)
            self.assertIn(
                rules.freeform_action,
                ('mirror', 'keep', 'analyze'),
                f'{lt} has invalid freeform_action: {rules.freeform_action}',
            )


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Tests: Slide dimensions stored
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


class TestSlideDimensions(unittest.TestCase):
    """Verify slide dimensions are stored."""

    def test_dimensions_stored(self):
        reg = TemplateRegistry(9144000, 6858000)
        self.assertEqual(reg._slide_width, 9144000)
        self.assertEqual(reg._slide_height, 6858000)

    def test_different_dimensions(self):
        reg = TemplateRegistry(12192000, 6858000)
        self.assertEqual(reg._slide_width, 12192000)


if __name__ == '__main__':
    unittest.main()
