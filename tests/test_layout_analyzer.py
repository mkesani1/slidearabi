"""
test_layout_analyzer.py — Comprehensive unit tests for layout_analyzer.py.

Tests cover:
- SlideLayoutType enum completeness
- LayoutClassification dataclass
- LayoutAnalyzer.classify_layout with explicit XML types
- LayoutAnalyzer._infer_type_from_placeholders heuristic rules
- LayoutAnalyzer._get_placeholder_summary counting
- LayoutAnalyzer._detect_two_column_spatial geometry detection
- LayoutAnalyzer.analyze_all batch classification
- Edge cases: empty presentations, missing attributes, unknown types
"""

from __future__ import annotations

import unittest
from dataclasses import dataclass
from typing import Any, Dict, List, Optional
from unittest.mock import MagicMock, PropertyMock, patch

from slidearabi.layout_analyzer import (
    LayoutAnalyzer,
    LayoutClassification,
    SlideLayoutType,
    _AI_CONFIDENCE_THRESHOLD,
    _DECORATIVE_PH_TYPES,
    _VALID_LAYOUT_TYPES,
    _normalise_ph_type,
)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Mock helpers
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


class MockPlaceholderFormat:
    """Mock for python-pptx PlaceholderFormat."""

    def __init__(self, ph_type: str, idx: int = 0):
        self.type = ph_type
        self.idx = idx


class MockPlaceholder:
    """Mock for a python-pptx placeholder shape."""

    def __init__(self, ph_type: str, idx: int = 0,
                 left: int = 0, width: int = 1000000,
                 top: int = 0, height: int = 500000,
                 name: str = ''):
        self.placeholder_format = MockPlaceholderFormat(ph_type, idx)
        self.is_placeholder = True
        self.left = left
        self.width = width
        self.top = top
        self.height = height
        self.name = name or f'{ph_type}_{idx}'


class MockElement:
    """Mock for lxml Element with get() support."""

    def __init__(self, attrs: Optional[Dict[str, str]] = None):
        self._attrs = attrs or {}

    def get(self, key: str, default=None):
        return self._attrs.get(key, default)


class MockLayout:
    """Mock for python-pptx SlideLayout."""

    def __init__(
        self,
        name: str = 'Test Layout',
        layout_type: Optional[str] = None,
        placeholders: Optional[List[MockPlaceholder]] = None,
    ):
        self.name = name
        self._placeholders = placeholders or []
        # Mock the _element for explicit type reading
        attrs = {}
        if layout_type is not None:
            attrs['type'] = layout_type
        self._element = MockElement(attrs)

    @property
    def placeholders(self):
        return self._placeholders


class MockSlide:
    """Mock for python-pptx Slide."""

    def __init__(
        self,
        slide_layout: MockLayout,
        placeholders: Optional[List[MockPlaceholder]] = None,
    ):
        self.slide_layout = slide_layout
        self._placeholders = placeholders or slide_layout._placeholders

    @property
    def placeholders(self):
        return self._placeholders


class MockPresentation:
    """Mock for python-pptx Presentation."""

    def __init__(
        self,
        slides: Optional[List[MockSlide]] = None,
        slide_width: int = 9144000,  # 10 inches
        slide_height: int = 6858000,  # 7.5 inches
        slide_masters: Optional[list] = None,
    ):
        self._slides = slides or []
        self.slide_width = slide_width
        self.slide_height = slide_height
        self.slide_masters = slide_masters or []

    @property
    def slides(self):
        return self._slides


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Tests: SlideLayoutType enum
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


class TestSlideLayoutType(unittest.TestCase):
    """Tests for the SlideLayoutType enumeration."""

    def test_has_36_members(self):
        """All 36 standard OOXML layout types are defined."""
        self.assertEqual(len(SlideLayoutType), 36)

    def test_values_are_strings(self):
        """Each enum member's value is a non-empty string."""
        for lt in SlideLayoutType:
            self.assertIsInstance(lt.value, str)
            self.assertTrue(len(lt.value) > 0)

    def test_known_types_present(self):
        """Key layout types exist in the enum."""
        expected = [
            'title', 'tx', 'twoColTx', 'tbl', 'chart', 'blank',
            'secHead', 'picTx', 'cust', 'titleOnly', 'obj',
        ]
        for name in expected:
            self.assertIn(name, _VALID_LAYOUT_TYPES)

    def test_str_enum_equality(self):
        """SlideLayoutType members are directly comparable to strings."""
        self.assertEqual(SlideLayoutType.TITLE, 'title')
        self.assertEqual(SlideLayoutType.BLANK, 'blank')
        self.assertEqual(SlideLayoutType.CUST, 'cust')


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Tests: _normalise_ph_type
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


class TestNormalisePlaceholderType(unittest.TestCase):
    """Tests for _normalise_ph_type helper."""

    def test_none_defaults_to_body(self):
        self.assertEqual(_normalise_ph_type(None), 'body')

    def test_title_enum_style(self):
        """Handles PP_PLACEHOLDER.TITLE-style string."""
        self.assertEqual(_normalise_ph_type('TITLE'), 'title')

    def test_center_title(self):
        self.assertEqual(_normalise_ph_type('CENTER_TITLE'), 'ctrTitle')

    def test_subtitle(self):
        self.assertEqual(_normalise_ph_type('SUBTITLE'), 'subTitle')

    def test_body(self):
        self.assertEqual(_normalise_ph_type('BODY'), 'body')

    def test_object(self):
        self.assertEqual(_normalise_ph_type('OBJECT'), 'obj')

    def test_chart(self):
        self.assertEqual(_normalise_ph_type('CHART'), 'chart')

    def test_table(self):
        self.assertEqual(_normalise_ph_type('TABLE'), 'tbl')

    def test_slide_number(self):
        self.assertEqual(_normalise_ph_type('SLIDE_NUMBER'), 'sldNum')

    def test_date(self):
        self.assertEqual(_normalise_ph_type('DATE'), 'dt')

    def test_footer(self):
        self.assertEqual(_normalise_ph_type('FOOTER'), 'ftr')

    def test_picture(self):
        self.assertEqual(_normalise_ph_type('PICTURE'), 'pic')

    def test_pptx_enum_with_prefix(self):
        """Handles 'PP_PLACEHOLDER.TITLE' style."""
        self.assertEqual(_normalise_ph_type('PP_PLACEHOLDER.TITLE'), 'title')

    def test_unknown_passthrough(self):
        """Unknown types are lowered and returned as-is."""
        self.assertEqual(_normalise_ph_type('WEIRD_TYPE'), 'weird_type')


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Tests: LayoutClassification dataclass
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


class TestLayoutClassification(unittest.TestCase):
    """Tests for LayoutClassification data model."""

    def test_construction(self):
        lc = LayoutClassification(
            slide_number=1,
            layout_name='Title Slide',
            explicit_type='title',
            resolved_type='title',
            confidence=1.0,
            placeholder_summary={'ctrTitle': 1, 'subTitle': 1},
            requires_ai_classification=False,
        )
        self.assertEqual(lc.slide_number, 1)
        self.assertEqual(lc.resolved_type, 'title')
        self.assertEqual(lc.confidence, 1.0)
        self.assertFalse(lc.requires_ai_classification)

    def test_low_confidence_flags_ai(self):
        """Classifications with low confidence should flag AI requirement."""
        lc = LayoutClassification(
            slide_number=2,
            layout_name='Custom',
            explicit_type=None,
            resolved_type='cust',
            confidence=0.4,
            placeholder_summary={},
            requires_ai_classification=True,
        )
        self.assertTrue(lc.requires_ai_classification)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Tests: LayoutAnalyzer explicit type reading
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


class TestExplicitTypeReading(unittest.TestCase):
    """Tests for _get_explicit_type."""

    def setUp(self):
        self.prs = MockPresentation()
        self.analyzer = LayoutAnalyzer(self.prs)

    def test_explicit_title(self):
        layout = MockLayout(layout_type='title')
        result = self.analyzer._get_explicit_type(layout)
        self.assertEqual(result, 'title')

    def test_explicit_twoColTx(self):
        layout = MockLayout(layout_type='twoColTx')
        result = self.analyzer._get_explicit_type(layout)
        self.assertEqual(result, 'twoColTx')

    def test_explicit_blank(self):
        layout = MockLayout(layout_type='blank')
        result = self.analyzer._get_explicit_type(layout)
        self.assertEqual(result, 'blank')

    def test_no_type_returns_none(self):
        layout = MockLayout()  # no layout_type
        result = self.analyzer._get_explicit_type(layout)
        self.assertIsNone(result)

    def test_cust_type(self):
        """'cust' in XML is still a valid type string."""
        layout = MockLayout(layout_type='cust')
        result = self.analyzer._get_explicit_type(layout)
        self.assertEqual(result, 'cust')


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Tests: LayoutAnalyzer.classify_layout (explicit types)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


class TestClassifyLayoutExplicit(unittest.TestCase):
    """Tests for classify_layout with explicit XML type attributes."""

    def setUp(self):
        self.prs = MockPresentation()
        self.analyzer = LayoutAnalyzer(self.prs)

    def test_explicit_type_returns_1_0_confidence(self):
        layout = MockLayout(layout_type='tx')
        lt, conf = self.analyzer.classify_layout(layout)
        self.assertEqual(lt, 'tx')
        self.assertEqual(conf, 1.0)

    def test_explicit_secHead(self):
        layout = MockLayout(layout_type='secHead')
        lt, conf = self.analyzer.classify_layout(layout)
        self.assertEqual(lt, 'secHead')
        self.assertEqual(conf, 1.0)

    def test_explicit_picTx(self):
        layout = MockLayout(layout_type='picTx')
        lt, conf = self.analyzer.classify_layout(layout)
        self.assertEqual(lt, 'picTx')
        self.assertEqual(conf, 1.0)

    def test_cust_falls_through_to_inference(self):
        """If type='cust', we try to infer a more specific type."""
        layout = MockLayout(
            layout_type='cust',
            placeholders=[
                MockPlaceholder('CENTER_TITLE', idx=0),
                MockPlaceholder('SUBTITLE', idx=1),
            ],
        )
        lt, conf = self.analyzer.classify_layout(layout)
        # Should infer 'title' from ctrTitle + subTitle
        self.assertEqual(lt, 'title')
        self.assertEqual(conf, 0.95)

    def test_caching_same_layout(self):
        """Classifying the same layout object twice uses cache."""
        layout = MockLayout(layout_type='tbl')
        lt1, c1 = self.analyzer.classify_layout(layout)
        lt2, c2 = self.analyzer.classify_layout(layout)
        self.assertEqual(lt1, lt2)
        self.assertEqual(c1, c2)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Tests: Heuristic inference rules
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


class TestInferTypeFromPlaceholders(unittest.TestCase):
    """Tests for _infer_type_from_placeholders heuristic rules."""

    def setUp(self):
        self.prs = MockPresentation()
        self.analyzer = LayoutAnalyzer(self.prs)

    def _make_layout(self, *ph_specs):
        """Create a MockLayout with given placeholder specs.
        Each spec is (type_string, idx) or just type_string.
        """
        phs = []
        for i, spec in enumerate(ph_specs):
            if isinstance(spec, tuple):
                ph_type, idx = spec
            else:
                ph_type = spec
                idx = i
            phs.append(MockPlaceholder(ph_type, idx=idx))
        return MockLayout(placeholders=phs)

    def test_rule_blank(self):
        """0 placeholders → blank (0.95)."""
        layout = self._make_layout()
        lt, conf = self.analyzer._infer_type_from_placeholders(layout)
        self.assertEqual(lt, 'blank')
        self.assertEqual(conf, 0.95)

    def test_rule_blank_with_decorative_only(self):
        """Only decorative placeholders → blank."""
        layout = self._make_layout('DATE', 'FOOTER', 'SLIDE_NUMBER')
        lt, conf = self.analyzer._infer_type_from_placeholders(layout)
        self.assertEqual(lt, 'blank')
        self.assertEqual(conf, 0.95)

    def test_rule_title_slide(self):
        """ctrTitle + subTitle → title (0.95)."""
        layout = self._make_layout('CENTER_TITLE', 'SUBTITLE')
        lt, conf = self.analyzer._infer_type_from_placeholders(layout)
        self.assertEqual(lt, 'title')
        self.assertEqual(conf, 0.95)

    def test_rule_title_only(self):
        """title + 0 body/objects → titleOnly (0.9)."""
        layout = self._make_layout('TITLE')
        lt, conf = self.analyzer._infer_type_from_placeholders(layout)
        self.assertEqual(lt, 'titleOnly')
        self.assertEqual(conf, 0.9)

    def test_rule_title_only_with_decorative(self):
        """title + decorative → still titleOnly."""
        layout = self._make_layout('TITLE', 'DATE', 'FOOTER', 'SLIDE_NUMBER')
        lt, conf = self.analyzer._infer_type_from_placeholders(layout)
        self.assertEqual(lt, 'titleOnly')
        self.assertEqual(conf, 0.9)

    def test_rule_title_and_body(self):
        """title + 1 body → tx (0.9)."""
        layout = self._make_layout('TITLE', 'BODY')
        lt, conf = self.analyzer._infer_type_from_placeholders(layout)
        self.assertEqual(lt, 'tx')
        self.assertEqual(conf, 0.9)

    def test_rule_title_and_table(self):
        """title + table → tbl (0.9)."""
        layout = self._make_layout('TITLE', 'TABLE')
        lt, conf = self.analyzer._infer_type_from_placeholders(layout)
        self.assertEqual(lt, 'tbl')
        self.assertEqual(conf, 0.9)

    def test_rule_title_and_chart(self):
        """title + chart → chart (0.9)."""
        layout = self._make_layout('TITLE', 'CHART')
        lt, conf = self.analyzer._infer_type_from_placeholders(layout)
        self.assertEqual(lt, 'chart')
        self.assertEqual(conf, 0.9)

    def test_rule_title_and_diagram(self):
        """title + diagram → dgm (0.85)."""
        layout = self._make_layout('TITLE', 'ORG_CHART')
        lt, conf = self.analyzer._infer_type_from_placeholders(layout)
        self.assertEqual(lt, 'dgm')
        self.assertEqual(conf, 0.85)

    def test_rule_title_body_chart(self):
        """title + body + chart → txAndChart (0.85)."""
        layout = self._make_layout('TITLE', 'BODY', 'CHART')
        lt, conf = self.analyzer._infer_type_from_placeholders(layout)
        self.assertEqual(lt, 'txAndChart')
        self.assertEqual(conf, 0.85)

    def test_rule_title_and_picture(self):
        """title + picture → picTx (0.85)."""
        layout = self._make_layout('TITLE', 'PICTURE')
        lt, conf = self.analyzer._infer_type_from_placeholders(layout)
        self.assertEqual(lt, 'picTx')
        self.assertEqual(conf, 0.85)

    def test_rule_title_two_objects(self):
        """title + 2 objects → twoObj (0.85)."""
        layout = self._make_layout(
            ('TITLE', 0), ('OBJECT', 1), ('OBJECT', 2)
        )
        lt, conf = self.analyzer._infer_type_from_placeholders(layout)
        self.assertEqual(lt, 'twoObj')
        self.assertEqual(conf, 0.85)

    def test_rule_title_body_two_objects(self):
        """title + body + 2 objects → txAndTwoObj (0.8)."""
        layout = self._make_layout(
            ('TITLE', 0), ('BODY', 1), ('OBJECT', 2), ('OBJECT', 3)
        )
        lt, conf = self.analyzer._infer_type_from_placeholders(layout)
        self.assertEqual(lt, 'txAndTwoObj')
        self.assertEqual(conf, 0.8)

    def test_rule_title_body_object(self):
        """title + body + object → txAndObj (0.85)."""
        layout = self._make_layout('TITLE', 'BODY', 'OBJECT')
        lt, conf = self.analyzer._infer_type_from_placeholders(layout)
        self.assertEqual(lt, 'txAndObj')
        self.assertEqual(conf, 0.85)

    def test_rule_title_single_object(self):
        """title + single object → obj (0.85)."""
        layout = self._make_layout('TITLE', 'OBJECT')
        lt, conf = self.analyzer._infer_type_from_placeholders(layout)
        self.assertEqual(lt, 'obj')
        self.assertEqual(conf, 0.85)

    def test_rule_four_objects(self):
        """4 objects → fourObj (0.85)."""
        layout = self._make_layout(
            ('OBJECT', 0), ('OBJECT', 1), ('OBJECT', 2), ('OBJECT', 3)
        )
        lt, conf = self.analyzer._infer_type_from_placeholders(layout)
        self.assertEqual(lt, 'fourObj')
        self.assertEqual(conf, 0.85)

    def test_rule_objects_only_no_title(self):
        """Object(s) without title → objOnly (0.8)."""
        layout = self._make_layout(('OBJECT', 0))
        lt, conf = self.analyzer._infer_type_from_placeholders(layout)
        self.assertEqual(lt, 'objOnly')
        self.assertEqual(conf, 0.8)

    def test_rule_fallback_cust(self):
        """Unrecognisable combination → cust (0.4)."""
        # Create a scenario that doesn't match any rule
        # Media alone with no title — reaches body-only check, but it's media
        layout = MockLayout(
            placeholders=[
                MockPlaceholder('MEDIA_CLIP', idx=0),
                MockPlaceholder('MEDIA_CLIP', idx=1),
            ]
        )
        lt, conf = self.analyzer._infer_type_from_placeholders(layout)
        # Should fall through to cust
        self.assertEqual(lt, 'cust')
        self.assertLess(conf, 0.7)

    def test_title_body_media(self):
        """title + body + media → txAndMedia (0.8)."""
        layout = self._make_layout('TITLE', 'BODY', 'MEDIA_CLIP')
        lt, conf = self.analyzer._infer_type_from_placeholders(layout)
        self.assertEqual(lt, 'txAndMedia')
        self.assertEqual(conf, 0.8)

    def test_title_body_clipart(self):
        """title + body + clip art → txAndClipArt (0.8)."""
        layout = self._make_layout('TITLE', 'BODY', 'CLIP_ART')
        lt, conf = self.analyzer._infer_type_from_placeholders(layout)
        self.assertEqual(lt, 'txAndClipArt')
        self.assertEqual(conf, 0.8)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Tests: Two-column spatial detection
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


class TestTwoColumnSpatial(unittest.TestCase):
    """Tests for _detect_two_column_spatial."""

    def setUp(self):
        self.prs = MockPresentation(slide_width=9144000)
        self.analyzer = LayoutAnalyzer(self.prs)
        self.slide_width = 9144000
        self.half = self.slide_width // 2

    def test_two_bodies_left_right(self):
        """Two body placeholders, one left and one right → two-column."""
        phs = [
            MockPlaceholder('BODY', idx=1, left=100000, width=4000000),  # left half
            MockPlaceholder('BODY', idx=2, left=4800000, width=4000000),  # right half
        ]
        self.assertTrue(
            self.analyzer._detect_two_column_spatial(phs, self.slide_width)
        )

    def test_two_bodies_both_left(self):
        """Two body placeholders both in left half → not two-column."""
        phs = [
            MockPlaceholder('BODY', idx=1, left=100000, width=2000000),
            MockPlaceholder('BODY', idx=2, left=2200000, width=2000000),
        ]
        self.assertFalse(
            self.analyzer._detect_two_column_spatial(phs, self.slide_width)
        )

    def test_single_body(self):
        """Single body placeholder → not two-column."""
        phs = [MockPlaceholder('BODY', idx=1, left=100000, width=8000000)]
        self.assertFalse(
            self.analyzer._detect_two_column_spatial(phs, self.slide_width)
        )

    def test_empty_list(self):
        """No placeholders → not two-column."""
        self.assertFalse(
            self.analyzer._detect_two_column_spatial([], self.slide_width)
        )

    def test_two_column_with_decorative(self):
        """Two-column works even if extra (non-body) shapes are ignored upstream."""
        # The method only receives body placeholders
        phs = [
            MockPlaceholder('BODY', idx=1, left=100000, width=4000000),
            MockPlaceholder('BODY', idx=2, left=5000000, width=4000000),
        ]
        self.assertTrue(
            self.analyzer._detect_two_column_spatial(phs, self.slide_width)
        )


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Tests: Two-column inference with spatial
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


class TestTwoColumnInference(unittest.TestCase):
    """Tests for two-column inference combining heuristic + spatial."""

    def setUp(self):
        self.slide_width = 9144000
        self.prs = MockPresentation(slide_width=self.slide_width)
        self.analyzer = LayoutAnalyzer(self.prs)

    def test_two_body_with_spatial_match(self):
        """title + 2 body with L/R arrangement → twoColTx (0.85)."""
        layout = MockLayout(
            placeholders=[
                MockPlaceholder('TITLE', idx=0),
                MockPlaceholder('BODY', idx=1, left=100000, width=4000000),
                MockPlaceholder('BODY', idx=2, left=5000000, width=4000000),
            ],
        )
        lt, conf = self.analyzer._infer_type_from_placeholders(layout)
        self.assertEqual(lt, 'twoColTx')
        self.assertEqual(conf, 0.85)

    def test_two_body_without_spatial_match(self):
        """title + 2 body without clear L/R → twoColTx (0.75)."""
        layout = MockLayout(
            placeholders=[
                MockPlaceholder('TITLE', idx=0),
                MockPlaceholder('BODY', idx=1, left=100000, width=2000000),
                MockPlaceholder('BODY', idx=2, left=2500000, width=2000000),
            ],
        )
        lt, conf = self.analyzer._infer_type_from_placeholders(layout)
        self.assertEqual(lt, 'twoColTx')
        self.assertEqual(conf, 0.75)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Tests: Placeholder summary
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


class TestPlaceholderSummary(unittest.TestCase):
    """Tests for _get_placeholder_summary."""

    def setUp(self):
        self.prs = MockPresentation()
        self.analyzer = LayoutAnalyzer(self.prs)

    def test_empty_layout(self):
        layout = MockLayout(placeholders=[])
        summary = self.analyzer._get_placeholder_summary(layout)
        self.assertEqual(summary, {})

    def test_title_body(self):
        layout = MockLayout(
            placeholders=[
                MockPlaceholder('TITLE', idx=0),
                MockPlaceholder('BODY', idx=1),
            ]
        )
        summary = self.analyzer._get_placeholder_summary(layout)
        self.assertEqual(summary.get('title'), 1)
        self.assertEqual(summary.get('body'), 1)

    def test_decorative_counted(self):
        """Decorative placeholders are counted (but excluded from rules by caller)."""
        layout = MockLayout(
            placeholders=[
                MockPlaceholder('TITLE', idx=0),
                MockPlaceholder('DATE', idx=10),
                MockPlaceholder('FOOTER', idx=11),
                MockPlaceholder('SLIDE_NUMBER', idx=12),
            ]
        )
        summary = self.analyzer._get_placeholder_summary(layout)
        self.assertEqual(summary.get('title'), 1)
        self.assertEqual(summary.get('dt'), 1)
        self.assertEqual(summary.get('ftr'), 1)
        self.assertEqual(summary.get('sldNum'), 1)

    def test_multiple_bodies(self):
        layout = MockLayout(
            placeholders=[
                MockPlaceholder('BODY', idx=1),
                MockPlaceholder('BODY', idx=2),
                MockPlaceholder('BODY', idx=3),
            ]
        )
        summary = self.analyzer._get_placeholder_summary(layout)
        self.assertEqual(summary.get('body'), 3)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Tests: LayoutAnalyzer.analyze_all
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


class TestAnalyzeAll(unittest.TestCase):
    """Tests for analyze_all batch classification."""

    def test_empty_presentation(self):
        prs = MockPresentation(slides=[])
        analyzer = LayoutAnalyzer(prs)
        results = analyzer.analyze_all()
        self.assertEqual(results, {})

    def test_single_slide_explicit_type(self):
        layout = MockLayout(
            name='Title Layout',
            layout_type='title',
            placeholders=[
                MockPlaceholder('CENTER_TITLE', idx=0),
                MockPlaceholder('SUBTITLE', idx=1),
            ],
        )
        slide = MockSlide(slide_layout=layout)
        prs = MockPresentation(slides=[slide])
        analyzer = LayoutAnalyzer(prs)
        results = analyzer.analyze_all()

        self.assertIn(1, results)
        cls = results[1]
        self.assertEqual(cls.slide_number, 1)
        self.assertEqual(cls.layout_name, 'Title Layout')
        self.assertEqual(cls.explicit_type, 'title')
        self.assertEqual(cls.resolved_type, 'title')
        self.assertEqual(cls.confidence, 1.0)
        self.assertFalse(cls.requires_ai_classification)

    def test_multiple_slides(self):
        layout_title = MockLayout(name='Title', layout_type='title')
        layout_tx = MockLayout(
            name='Content',
            layout_type='tx',
            placeholders=[
                MockPlaceholder('TITLE', idx=0),
                MockPlaceholder('BODY', idx=1),
            ],
        )
        layout_blank = MockLayout(name='Blank', layout_type='blank')

        slides = [
            MockSlide(slide_layout=layout_title),
            MockSlide(slide_layout=layout_tx),
            MockSlide(slide_layout=layout_blank),
        ]
        prs = MockPresentation(slides=slides)
        analyzer = LayoutAnalyzer(prs)
        results = analyzer.analyze_all()

        self.assertEqual(len(results), 3)
        self.assertEqual(results[1].resolved_type, 'title')
        self.assertEqual(results[2].resolved_type, 'tx')
        self.assertEqual(results[3].resolved_type, 'blank')

    def test_slide_requires_ai_when_low_confidence(self):
        """A custom layout with low confidence flags requires_ai."""
        layout = MockLayout(
            name='Weird Layout',
            placeholders=[
                MockPlaceholder('MEDIA_CLIP', idx=0),
                MockPlaceholder('MEDIA_CLIP', idx=1),
            ],
        )
        slide = MockSlide(slide_layout=layout)
        prs = MockPresentation(slides=[slide])
        analyzer = LayoutAnalyzer(prs)
        results = analyzer.analyze_all()

        cls = results[1]
        self.assertTrue(cls.requires_ai_classification)
        self.assertLess(cls.confidence, _AI_CONFIDENCE_THRESHOLD)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Tests: classify_slide
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


class TestClassifySlide(unittest.TestCase):
    """Tests for classify_slide method."""

    def test_returns_layout_classification(self):
        layout = MockLayout(
            name='Title and Content',
            layout_type='tx',
            placeholders=[
                MockPlaceholder('TITLE', idx=0),
                MockPlaceholder('BODY', idx=1),
            ],
        )
        slide = MockSlide(
            slide_layout=layout,
            placeholders=[
                MockPlaceholder('TITLE', idx=0),
                MockPlaceholder('BODY', idx=1),
            ],
        )
        prs = MockPresentation(slides=[slide])
        analyzer = LayoutAnalyzer(prs)

        cls = analyzer.classify_slide(slide, 1)
        self.assertIsInstance(cls, LayoutClassification)
        self.assertEqual(cls.resolved_type, 'tx')
        self.assertEqual(cls.layout_name, 'Title and Content')
        self.assertIn('title', cls.placeholder_summary)

    def test_slide_placeholder_summary_from_slide_not_layout(self):
        """Placeholder summary comes from the slide, not the layout."""
        layout = MockLayout(
            name='Content',
            layout_type='tx',
            placeholders=[
                MockPlaceholder('TITLE', idx=0),
                MockPlaceholder('BODY', idx=1),
            ],
        )
        # Slide has an extra placeholder not on layout
        slide = MockSlide(
            slide_layout=layout,
            placeholders=[
                MockPlaceholder('TITLE', idx=0),
                MockPlaceholder('BODY', idx=1),
                MockPlaceholder('CHART', idx=2),
            ],
        )
        prs = MockPresentation(slides=[slide])
        analyzer = LayoutAnalyzer(prs)

        cls = analyzer.classify_slide(slide, 1)
        self.assertEqual(cls.placeholder_summary.get('chart'), 1)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Tests: Edge cases
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


class TestEdgeCases(unittest.TestCase):
    """Edge case tests for LayoutAnalyzer."""

    def test_layout_with_broken_placeholder(self):
        """Handles placeholders that raise exceptions gracefully."""

        class BrokenPlaceholder:
            @property
            def placeholder_format(self):
                raise AttributeError("broken")

        layout = MockLayout()
        layout._placeholders = [BrokenPlaceholder()]
        prs = MockPresentation()
        analyzer = LayoutAnalyzer(prs)

        summary = analyzer._get_placeholder_summary(layout)
        # Should not crash; returns empty
        self.assertEqual(summary, {})

    def test_layout_without_element(self):
        """Handles layout with broken _element gracefully."""

        class BrokenLayout:
            name = 'Broken'
            _element = None
            @property
            def placeholders(self):
                return []

        prs = MockPresentation()
        analyzer = LayoutAnalyzer(prs)

        result = analyzer._get_explicit_type(BrokenLayout())
        self.assertIsNone(result)

    def test_get_layout_type_for_slide_valid(self):
        """get_layout_type_for_slide returns resolved type."""
        layout = MockLayout(layout_type='blank')
        slide = MockSlide(slide_layout=layout)
        prs = MockPresentation(slides=[slide])
        analyzer = LayoutAnalyzer(prs)

        result = analyzer.get_layout_type_for_slide(1)
        self.assertEqual(result, 'blank')

    def test_get_layout_type_for_slide_invalid_number(self):
        """Returns None for invalid slide number."""
        prs = MockPresentation(slides=[])
        analyzer = LayoutAnalyzer(prs)
        self.assertIsNone(analyzer.get_layout_type_for_slide(99))

    def test_get_layout_type_for_slide_zero(self):
        """Returns None for slide number 0."""
        layout = MockLayout(layout_type='tx')
        slide = MockSlide(slide_layout=layout)
        prs = MockPresentation(slides=[slide])
        analyzer = LayoutAnalyzer(prs)
        self.assertIsNone(analyzer.get_layout_type_for_slide(0))

    def test_get_all_layout_types(self):
        """get_all_layout_types returns layout name → type mapping."""

        class MockMaster:
            slide_layouts = []

        class MockMasterWithLayouts:
            def __init__(self):
                self.slide_layouts = [
                    MockLayout(name='Title Slide', layout_type='title'),
                    MockLayout(name='Content', layout_type='tx'),
                ]

        prs = MockPresentation(slide_masters=[MockMasterWithLayouts()])
        analyzer = LayoutAnalyzer(prs)
        result = analyzer.get_all_layout_types()

        self.assertEqual(result.get('Title Slide'), 'title')
        self.assertEqual(result.get('Content'), 'tx')


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Tests: Decorative placeholder filtering
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


class TestDecorativePlaceholders(unittest.TestCase):
    """Ensure decorative placeholders don't affect layout classification."""

    def test_decorative_types_are_correct(self):
        self.assertIn('dt', _DECORATIVE_PH_TYPES)
        self.assertIn('ftr', _DECORATIVE_PH_TYPES)
        self.assertIn('sldNum', _DECORATIVE_PH_TYPES)

    def test_title_with_decorative_still_title_only(self):
        """title + date + footer + slideNum → titleOnly, not tx."""
        layout = MockLayout(
            placeholders=[
                MockPlaceholder('TITLE', idx=0),
                MockPlaceholder('DATE', idx=10),
                MockPlaceholder('FOOTER', idx=11),
                MockPlaceholder('SLIDE_NUMBER', idx=12),
            ]
        )
        prs = MockPresentation()
        analyzer = LayoutAnalyzer(prs)
        lt, conf = analyzer._infer_type_from_placeholders(layout)
        self.assertEqual(lt, 'titleOnly')


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Tests: AI confidence threshold
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


class TestAIConfidenceThreshold(unittest.TestCase):
    """Test that the AI classification threshold works correctly."""

    def test_threshold_is_reasonable(self):
        self.assertGreater(_AI_CONFIDENCE_THRESHOLD, 0.5)
        self.assertLessEqual(_AI_CONFIDENCE_THRESHOLD, 1.0)

    def test_explicit_type_never_needs_ai(self):
        layout = MockLayout(layout_type='tx')
        slide = MockSlide(slide_layout=layout)
        prs = MockPresentation(slides=[slide])
        analyzer = LayoutAnalyzer(prs)
        cls = analyzer.classify_slide(slide, 1)
        self.assertFalse(cls.requires_ai_classification)

    def test_confident_inference_no_ai(self):
        layout = MockLayout(
            placeholders=[
                MockPlaceholder('CENTER_TITLE', idx=0),
                MockPlaceholder('SUBTITLE', idx=1),
            ]
        )
        slide = MockSlide(slide_layout=layout)
        prs = MockPresentation(slides=[slide])
        analyzer = LayoutAnalyzer(prs)
        cls = analyzer.classify_slide(slide, 1)
        self.assertFalse(cls.requires_ai_classification)
        self.assertGreaterEqual(cls.confidence, _AI_CONFIDENCE_THRESHOLD)


if __name__ == '__main__':
    unittest.main()
