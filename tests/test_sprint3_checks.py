"""
tests/test_sprint3_checks.py — Sprint 3 Unit Tests

Tests for V3XMLChecker checks #7 (circular centering), #8 (master mirror),
#12 (directional orientation) and corresponding fixes #7, #12.
Also tests: vision prompt builder, selective vision, API contract.
"""

from __future__ import annotations

import os
import sys

import pytest
from lxml import etree

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from v3_checks import (
    V3AutoFixer,
    V3XMLChecker,
    A_NS,
    P_NS,
    SLIDEARABI_NS,
)
from vqa_types import Severity, V3Defect


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

SLIDE_W = 9144000
SLIDE_H = 6858000


def _make_shape(x: int, y: int, cx: int, cy: int, text: str = '',
                name: str = 'Shape', prst: str = None,
                algn: str = None, anchor: str = None,
                flipH: str = None) -> etree._Element:
    """Build a minimal <p:sp> with optional preset geometry and text."""
    sp = etree.Element(f'{{{P_NS}}}sp')
    nvSpPr = etree.SubElement(sp, f'{{{P_NS}}}nvSpPr')
    cNvPr = etree.SubElement(nvSpPr, f'{{{P_NS}}}cNvPr')
    cNvPr.set('id', '1')
    cNvPr.set('name', name)

    spPr = etree.SubElement(sp, f'{{{P_NS}}}spPr')
    xfrm = etree.SubElement(spPr, f'{{{A_NS}}}xfrm')
    if flipH:
        xfrm.set('flipH', flipH)
    off = etree.SubElement(xfrm, f'{{{A_NS}}}off')
    off.set('x', str(x))
    off.set('y', str(y))
    ext = etree.SubElement(xfrm, f'{{{A_NS}}}ext')
    ext.set('cx', str(cx))
    ext.set('cy', str(cy))

    if prst:
        prstGeom = etree.SubElement(spPr, f'{{{A_NS}}}prstGeom')
        prstGeom.set('prst', prst)

    if text:
        txBody = etree.SubElement(sp, f'{{{P_NS}}}txBody')
        if anchor:
            bodyPr = etree.SubElement(txBody, f'{{{A_NS}}}bodyPr')
            bodyPr.set('anchor', anchor)
        else:
            etree.SubElement(txBody, f'{{{A_NS}}}bodyPr')
        p = etree.SubElement(txBody, f'{{{A_NS}}}p')
        if algn:
            pPr = etree.SubElement(p, f'{{{A_NS}}}pPr')
            pPr.set('algn', algn)
        r = etree.SubElement(p, f'{{{A_NS}}}r')
        t = etree.SubElement(r, f'{{{A_NS}}}t')
        t.text = text

    return sp


def _wrap_in_slide(*elements) -> etree._Element:
    """Wrap elements in a minimal slide XML structure."""
    sld = etree.Element(f'{{{P_NS}}}sld')
    cSld = etree.SubElement(sld, f'{{{P_NS}}}cSld')
    spTree = etree.SubElement(cSld, f'{{{P_NS}}}spTree')
    for elem in elements:
        spTree.append(elem)
    return sld


# ─────────────────────────────────────────────────────────────────────────────
# CHECK #7: TEXT_NOT_CENTERED_IN_SHAPE
# ─────────────────────────────────────────────────────────────────────────────

class TestCheckCircularTextCentering:

    def test_detects_uncentered_ellipse(self):
        """Ellipse with left-aligned text should be flagged."""
        sp = _make_shape(100000, 100000, 1500000, 1500000,
                        text='Test', prst='ellipse', algn='l')
        slide = _wrap_in_slide(sp)

        checker = V3XMLChecker()
        defects = checker._check_circular_text_centering(1, slide)
        assert len(defects) >= 1
        assert defects[0].code == "TEXT_NOT_CENTERED_IN_SHAPE"

    def test_passes_centered_ellipse(self):
        """Properly centered ellipse should pass."""
        sp = _make_shape(100000, 100000, 1500000, 1500000,
                        text='Test', prst='ellipse', algn='ctr', anchor='ctr')
        slide = _wrap_in_slide(sp)

        checker = V3XMLChecker()
        defects = checker._check_circular_text_centering(1, slide)
        assert len(defects) == 0

    def test_detects_near_square_without_preset(self):
        """Near-square shape (ratio <= 1.2) without preset should be checked."""
        sp = _make_shape(100000, 100000, 1000000, 1100000,
                        text='Test', algn='l')
        slide = _wrap_in_slide(sp)

        checker = V3XMLChecker()
        defects = checker._check_circular_text_centering(1, slide)
        assert len(defects) >= 1

    def test_skips_rectangular_shapes(self):
        """Rectangular shapes (ratio > 1.2) should be skipped."""
        sp = _make_shape(100000, 100000, 3000000, 1000000,
                        text='Test', algn='l')
        slide = _wrap_in_slide(sp)

        checker = V3XMLChecker()
        defects = checker._check_circular_text_centering(1, slide)
        assert len(defects) == 0

    def test_skips_empty_circular(self):
        """Ellipse with no text should not be flagged."""
        sp = _make_shape(100000, 100000, 1500000, 1500000, prst='ellipse')
        slide = _wrap_in_slide(sp)

        checker = V3XMLChecker()
        defects = checker._check_circular_text_centering(1, slide)
        assert len(defects) == 0

    def test_detects_missing_anchor(self):
        """Ellipse with centered algn but missing anchor should be flagged."""
        sp = _make_shape(100000, 100000, 1500000, 1500000,
                        text='Test', prst='ellipse', algn='ctr')
        # anchor defaults to None in _make_shape → no anchor='ctr'
        slide = _wrap_in_slide(sp)

        checker = V3XMLChecker()
        defects = checker._check_circular_text_centering(1, slide)
        assert len(defects) >= 1
        assert 'anchor not ctr' in defects[0].evidence.get('issues', [])


# ─────────────────────────────────────────────────────────────────────────────
# CHECK #8: MASTER_ELEMENT_NOT_MIRRORED
# ─────────────────────────────────────────────────────────────────────────────

class TestCheckMasterElementMirroring:

    def test_detects_unmirrored_line(self):
        """Line named 'Line 1' at same x in both orig and conv should be flagged."""
        orig_sp = _make_shape(500000, 100000, 2000000, 50000, name='Line 1')
        conv_sp = _make_shape(500000, 100000, 2000000, 50000, name='Line 1')

        orig_slide = _wrap_in_slide(orig_sp)
        conv_slide = _wrap_in_slide(conv_sp)

        checker = V3XMLChecker()
        defects = checker._check_master_element_mirroring(1, orig_slide, conv_slide)
        assert len(defects) >= 1
        assert defects[0].code == "MASTER_ELEMENT_NOT_MIRRORED"

    def test_passes_mirrored_line(self):
        """Line at correct mirrored position should pass."""
        orig_x = 500000
        cx = 2000000
        expected_x = SLIDE_W - orig_x - cx

        orig_sp = _make_shape(orig_x, 100000, cx, 50000, name='Line 1')
        conv_sp = _make_shape(expected_x, 100000, cx, 50000, name='Line 1')

        orig_slide = _wrap_in_slide(orig_sp)
        conv_slide = _wrap_in_slide(conv_sp)

        checker = V3XMLChecker()
        defects = checker._check_master_element_mirroring(1, orig_slide, conv_slide)
        assert len(defects) == 0

    def test_skips_non_master_shapes(self):
        """Shape not matching master patterns should be skipped."""
        orig_sp = _make_shape(500000, 100000, 2000000, 50000, name='Content 1')
        conv_sp = _make_shape(500000, 100000, 2000000, 50000, name='Content 1')

        orig_slide = _wrap_in_slide(orig_sp)
        conv_slide = _wrap_in_slide(conv_sp)

        checker = V3XMLChecker()
        defects = checker._check_master_element_mirroring(1, orig_slide, conv_slide)
        assert len(defects) == 0

    def test_skips_centered_master(self):
        """Centered master element should be skipped."""
        cx = 2000000
        center_x = (SLIDE_W - cx) // 2
        orig_sp = _make_shape(center_x, 100000, cx, 50000, name='Border 1')
        conv_sp = _make_shape(center_x, 100000, cx, 50000, name='Border 1')

        orig_slide = _wrap_in_slide(orig_sp)
        conv_slide = _wrap_in_slide(conv_sp)

        checker = V3XMLChecker()
        defects = checker._check_master_element_mirroring(1, orig_slide, conv_slide)
        assert len(defects) == 0

    def test_skips_fullwidth_master(self):
        """Full-width master bar should be skipped."""
        cx = int(SLIDE_W * 0.9)
        orig_sp = _make_shape(100000, 0, cx, 200000, name='Background Bar')
        conv_sp = _make_shape(100000, 0, cx, 200000, name='Background Bar')

        orig_slide = _wrap_in_slide(orig_sp)
        conv_slide = _wrap_in_slide(conv_sp)

        checker = V3XMLChecker()
        defects = checker._check_master_element_mirroring(1, orig_slide, conv_slide)
        assert len(defects) == 0


# ─────────────────────────────────────────────────────────────────────────────
# CHECK #12: DIRECTIONAL_SHAPE_NOT_FLIPPED
# ─────────────────────────────────────────────────────────────────────────────

class TestCheckDirectionalSymbolOrientation:

    def test_detects_unflipped_arrow(self):
        """Right arrow without flipH should be flagged."""
        sp = _make_shape(100000, 100000, 2000000, 500000,
                        text='Next', prst='rightArrow')
        slide = _wrap_in_slide(sp)

        checker = V3XMLChecker()
        defects = checker._check_directional_symbol_orientation(1, slide)
        assert len(defects) >= 1
        assert defects[0].code == "DIRECTIONAL_SHAPE_NOT_FLIPPED"

    def test_passes_flipped_arrow(self):
        """Right arrow with flipH='1' should pass."""
        sp = _make_shape(100000, 100000, 2000000, 500000,
                        text='Next', prst='rightArrow', flipH='1')
        slide = _wrap_in_slide(sp)

        checker = V3XMLChecker()
        defects = checker._check_directional_symbol_orientation(1, slide)
        assert len(defects) == 0

    def test_detects_unflipped_chevron(self):
        """Chevron without flipH should be flagged."""
        sp = _make_shape(100000, 100000, 2000000, 500000,
                        text='Step', prst='chevron')
        slide = _wrap_in_slide(sp)

        checker = V3XMLChecker()
        defects = checker._check_directional_symbol_orientation(1, slide)
        assert len(defects) >= 1

    def test_skips_non_directional(self):
        """Rectangle (non-directional) should be skipped."""
        sp = _make_shape(100000, 100000, 2000000, 500000,
                        text='Box', prst='rect')
        slide = _wrap_in_slide(sp)

        checker = V3XMLChecker()
        defects = checker._check_directional_symbol_orientation(1, slide)
        assert len(defects) == 0

    def test_skips_shapes_without_preset(self):
        """Shapes without prstGeom should be skipped."""
        sp = _make_shape(100000, 100000, 2000000, 500000, text='Arrow')
        slide = _wrap_in_slide(sp)

        checker = V3XMLChecker()
        defects = checker._check_directional_symbol_orientation(1, slide)
        assert len(defects) == 0

    def test_directional_presets_coverage(self):
        """All expected directional presets are in DIRECTIONAL_PRESETS."""
        expected = {'rightArrow', 'chevron', 'homePlate', 'notchedRightArrow'}
        assert expected.issubset(V3XMLChecker.DIRECTIONAL_PRESETS)


# ─────────────────────────────────────────────────────────────────────────────
# FIX #7: center_text_circular
# ─────────────────────────────────────────────────────────────────────────────

class TestFixCenterTextCircular:

    def test_centers_text_in_ellipse(self):
        """Fix should set algn='ctr' and anchor='ctr'."""
        sp = _make_shape(100000, 100000, 1500000, 1500000,
                        text='Test', prst='ellipse', algn='l')
        slide = _wrap_in_slide(sp)

        defect = V3Defect(
            code="TEXT_NOT_CENTERED_IN_SHAPE",
            category="alignment",
            severity=Severity.MEDIUM,
            slide_idx=1,
            object_id="0",
            evidence={'preset': 'ellipse', 'issues': ['algn not ctr', 'anchor not ctr']},
            fixable=True,
            autofix_action="center_text_circular",
        )

        fixer = V3AutoFixer()
        assert fixer.apply_fix(slide, defect) is True

        # Verify centering
        bodyPr = slide.find(f'.//{{{A_NS}}}bodyPr')
        assert bodyPr.get('anchor') == 'ctr'

        pPr = slide.find(f'.//{{{A_NS}}}pPr')
        assert pPr.get('algn') == 'ctr'


# ─────────────────────────────────────────────────────────────────────────────
# FIX #12: flip_directional_shape
# ─────────────────────────────────────────────────────────────────────────────

class TestFixFlipDirectionalShape:

    def test_sets_flipH(self):
        """Fix should set flipH='1' on xfrm."""
        sp = _make_shape(100000, 100000, 2000000, 500000,
                        text='Next', prst='rightArrow')
        slide = _wrap_in_slide(sp)

        defect = V3Defect(
            code="DIRECTIONAL_SHAPE_NOT_FLIPPED",
            category="mirroring",
            severity=Severity.HIGH,
            slide_idx=1,
            object_id="0",
            evidence={'preset': 'rightArrow'},
            fixable=True,
            autofix_action="flip_directional_shape",
        )

        fixer = V3AutoFixer()
        assert fixer.apply_fix(slide, defect) is True

        xfrm = slide.find(f'.//{{{A_NS}}}xfrm')
        assert xfrm.get('flipH') == '1'

    def test_idempotent(self):
        """Already-flipped shape should return False."""
        sp = _make_shape(100000, 100000, 2000000, 500000,
                        text='Next', prst='rightArrow', flipH='1')
        slide = _wrap_in_slide(sp)

        defect = V3Defect(
            code="DIRECTIONAL_SHAPE_NOT_FLIPPED",
            category="mirroring",
            severity=Severity.HIGH,
            slide_idx=1,
            object_id="0",
            evidence={'preset': 'rightArrow'},
            fixable=True,
            autofix_action="flip_directional_shape",
        )

        fixer = V3AutoFixer()
        assert fixer.apply_fix(slide, defect) is False  # No-op


# ─────────────────────────────────────────────────────────────────────────────
# ROUND-TRIP REGRESSION
# ─────────────────────────────────────────────────────────────────────────────

class TestSprint3RoundTrip:

    def test_fix7_then_recheck_yields_zero(self):
        """Apply circular centering fix, recheck → 0 defects."""
        sp = _make_shape(100000, 100000, 1500000, 1500000,
                        text='Test', prst='ellipse', algn='l')
        slide = _wrap_in_slide(sp)

        checker = V3XMLChecker()
        fixer = V3AutoFixer()

        defects = checker._check_circular_text_centering(1, slide)
        assert len(defects) >= 1

        for d in defects:
            fixer.apply_fix(slide, d)

        defects_after = checker._check_circular_text_centering(1, slide)
        assert len(defects_after) == 0

    def test_fix12_then_recheck_yields_zero(self):
        """Apply directional flip fix, recheck → 0 defects."""
        sp = _make_shape(100000, 100000, 2000000, 500000,
                        text='Go', prst='rightArrow')
        slide = _wrap_in_slide(sp)

        checker = V3XMLChecker()
        fixer = V3AutoFixer()

        defects = checker._check_directional_symbol_orientation(1, slide)
        assert len(defects) >= 1

        for d in defects:
            fixer.apply_fix(slide, d)

        defects_after = checker._check_directional_symbol_orientation(1, slide)
        assert len(defects_after) == 0


# ─────────────────────────────────────────────────────────────────────────────
# VISION PROMPT BUILDER
# ─────────────────────────────────────────────────────────────────────────────

class TestVisionPromptBuilder:

    def test_enhanced_prompts_included_when_flagged(self):
        """V3 prompt additions should appear when flag is on."""
        import v3_config
        orig_val = v3_config.ENABLE_ENHANCED_PROMPTS
        try:
            v3_config.ENABLE_ENHANCED_PROMPTS = True
            from v3_vision_prompts import build_enhanced_system_prompt
            result = build_enhanced_system_prompt("Base prompt")
            assert "TABLE COLUMN ORDER" in result
            assert "DIRECTIONAL SHAPES" in result
        finally:
            v3_config.ENABLE_ENHANCED_PROMPTS = orig_val

    def test_enhanced_prompts_excluded_when_not_flagged(self):
        """V3 prompt additions should not appear when flag is off."""
        import v3_config
        orig_val = v3_config.ENABLE_ENHANCED_PROMPTS
        try:
            v3_config.ENABLE_ENHANCED_PROMPTS = False
            from v3_vision_prompts import build_enhanced_system_prompt
            result = build_enhanced_system_prompt("Base prompt")
            assert result == "Base prompt"
        finally:
            v3_config.ENABLE_ENHANCED_PROMPTS = orig_val

    def test_xml_context_in_user_prompt(self):
        """XML defects should be injected into user prompt when flag is on."""
        import v3_config
        orig_prompts = v3_config.ENABLE_ENHANCED_PROMPTS
        orig_ctx = v3_config.ENABLE_VISION_XML_CTX
        try:
            v3_config.ENABLE_ENHANCED_PROMPTS = True
            v3_config.ENABLE_VISION_XML_CTX = True
            from v3_vision_prompts import build_enhanced_user_prompt
            defects = [{'code': 'TABLE_COLUMNS_NOT_REVERSED', 'slide': 1}]
            result = build_enhanced_user_prompt("Base", 1, defects)
            assert "XML STRUCTURAL FINDINGS" in result
            assert "TABLE_COLUMNS_NOT_REVERSED" in result
        finally:
            v3_config.ENABLE_ENHANCED_PROMPTS = orig_prompts
            v3_config.ENABLE_VISION_XML_CTX = orig_ctx


# ─────────────────────────────────────────────────────────────────────────────
# SELECTIVE VISION
# ─────────────────────────────────────────────────────────────────────────────

class TestSelectiveVision:

    def test_non_selective_returns_all(self):
        """When selective vision is off, returns all slides up to cap."""
        import v3_config
        orig_val = v3_config.ENABLE_SELECTIVE_VISION
        try:
            v3_config.ENABLE_SELECTIVE_VISION = False
            from v3_vision_prompts import select_slides_for_vision
            result = select_slides_for_vision([], 10, max_vision_slides=20)
            assert result == list(range(1, 11))
        finally:
            v3_config.ENABLE_SELECTIVE_VISION = orig_val

    def test_selective_prioritizes_critical(self):
        """Slides with CRITICAL defects should be prioritized."""
        import v3_config
        orig_val = v3_config.ENABLE_SELECTIVE_VISION
        try:
            v3_config.ENABLE_SELECTIVE_VISION = True
            from v3_vision_prompts import select_slides_for_vision
            from vqa_types import V3Defect, Severity, DefectStatus

            defects = [
                V3Defect(code="TABLE_COLUMNS_NOT_REVERSED", category="table",
                         severity=Severity.CRITICAL, slide_idx=5, object_id="0",
                         fixable=True, autofix_action="reverse_table_columns"),
                V3Defect(code="PARAGRAPH_RTL_MISSING", category="alignment",
                         severity=Severity.MEDIUM, slide_idx=2, object_id="0",
                         fixable=True, autofix_action="set_paragraph_rtl"),
            ]

            result = select_slides_for_vision(defects, 10, max_vision_slides=3)
            assert 5 in result  # Critical slide must be included
            assert len(result) <= 3
        finally:
            v3_config.ENABLE_SELECTIVE_VISION = orig_val


# ─────────────────────────────────────────────────────────────────────────────
# API CONTRACT
# ─────────────────────────────────────────────────────────────────────────────

class TestAPIContract:

    def test_completed_status(self):
        """Completed job should have proper fields."""
        from v3_api_contract import build_status_response
        resp = build_status_response(
            job_id='test-123', phase='done', progress=1.0,
            download_url='/download/test-123.pptx',
        )
        assert resp['status'] == 'completed'
        assert resp['download_available'] is True
        assert 'vqa' not in resp  # No gate result → no VQA field

    def test_completed_with_warnings(self):
        """Job with warnings should include VQA summary."""
        from v3_api_contract import build_status_response
        from vqa_types import VQAGateResult

        gate = VQAGateResult()
        gate.status = 'completed_with_warnings'
        gate.critical_remaining = 0
        gate.high_remaining = 3
        gate.warning_issues = [{'code': 'test'}]

        resp = build_status_response(
            job_id='test-123', phase='done', progress=1.0,
            gate_result=gate, download_url='/download/test-123.pptx',
        )
        assert resp['status'] == 'completed_with_warnings'
        assert resp['vqa']['high_remaining'] == 3
        assert resp['download_available'] is True

    def test_failed_qa(self):
        """Failed QA should still allow download."""
        from v3_api_contract import build_status_response
        from vqa_types import VQAGateResult

        gate = VQAGateResult()
        gate.status = 'failed_qa'
        gate.critical_remaining = 2
        gate.blocking_issues = [{'code': 'TABLE_COLUMNS_NOT_REVERSED'}]

        resp = build_status_response(
            job_id='test-123', phase='done', progress=1.0,
            gate_result=gate, download_url='/download/test-123.pptx',
        )
        assert resp['status'] == 'failed_qa'
        assert resp['download_available'] is True
        assert len(resp['vqa']['blocking_issues']) == 1

    def test_processing_status(self):
        """In-progress job should show processing."""
        from v3_api_contract import build_status_response
        resp = build_status_response(
            job_id='test-123', phase='translating', progress=0.5,
        )
        assert resp['status'] == 'processing'
        assert resp['download_available'] is False

    def test_backward_compat(self):
        """V2 clients should see standard fields regardless of V3 additions."""
        from v3_api_contract import build_status_response
        resp = build_status_response(
            job_id='test-123', phase='done', progress=1.0,
            download_url='/download/test-123.pptx',
        )
        # V2 expected fields
        assert 'job_id' in resp
        assert 'status' in resp
        assert 'progress' in resp


# ─────────────────────────────────────────────────────────────────────────────
# INTEGRATION: check_slide now includes Sprint 3 checks
# ─────────────────────────────────────────────────────────────────────────────

class TestCheckSlideIntegrationSprint3:

    def test_includes_circular_and_directional(self):
        """check_slide should detect both circular centering and directional issues."""
        ellipse = _make_shape(100000, 100000, 1500000, 1500000,
                             text='Venn', prst='ellipse', algn='l')
        arrow = _make_shape(3000000, 100000, 2000000, 500000,
                           text='Next', prst='rightArrow')
        slide = _wrap_in_slide(ellipse, arrow)

        checker = V3XMLChecker()
        defects = checker.check_slide(1, slide)

        codes = [d.code for d in defects]
        assert 'TEXT_NOT_CENTERED_IN_SHAPE' in codes
        assert 'DIRECTIONAL_SHAPE_NOT_FLIPPED' in codes
