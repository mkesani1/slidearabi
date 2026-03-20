"""
tests/test_sprint2_checks.py — Sprint 2 Unit Tests

Tests for V3XMLChecker checks #3 (icon), #4 (page number), #5 (shape position),
#6 (paragraph RTL) and corresponding fixes #3, #4, #5, #6.
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
                rtl: str = None, algn: str = None) -> etree._Element:
    """Build a minimal <p:sp> with position, size, and optional text."""
    sp = etree.Element(f'{{{P_NS}}}sp')
    nvSpPr = etree.SubElement(sp, f'{{{P_NS}}}nvSpPr')
    cNvPr = etree.SubElement(nvSpPr, f'{{{P_NS}}}cNvPr')
    cNvPr.set('id', '1')
    cNvPr.set('name', 'Shape')

    spPr = etree.SubElement(sp, f'{{{P_NS}}}spPr')
    xfrm = etree.SubElement(spPr, f'{{{A_NS}}}xfrm')
    off = etree.SubElement(xfrm, f'{{{A_NS}}}off')
    off.set('x', str(x))
    off.set('y', str(y))
    ext = etree.SubElement(xfrm, f'{{{A_NS}}}ext')
    ext.set('cx', str(cx))
    ext.set('cy', str(cy))

    if text:
        txBody = etree.SubElement(sp, f'{{{P_NS}}}txBody')
        p = etree.SubElement(txBody, f'{{{A_NS}}}p')
        if rtl or algn:
            pPr = etree.SubElement(p, f'{{{A_NS}}}pPr')
            if rtl:
                pPr.set('rtl', rtl)
            if algn:
                pPr.set('algn', algn)
        r = etree.SubElement(p, f'{{{A_NS}}}r')
        t = etree.SubElement(r, f'{{{A_NS}}}t')
        t.text = text

    return sp


def _make_pic(x: int, y: int, cx: int, cy: int) -> etree._Element:
    """Build a minimal <p:pic> with position and size."""
    pic = etree.Element(f'{{{P_NS}}}pic')
    nvPicPr = etree.SubElement(pic, f'{{{P_NS}}}nvPicPr')
    spPr = etree.SubElement(pic, f'{{{P_NS}}}spPr')
    xfrm = etree.SubElement(spPr, f'{{{A_NS}}}xfrm')
    off = etree.SubElement(xfrm, f'{{{A_NS}}}off')
    off.set('x', str(x))
    off.set('y', str(y))
    ext = etree.SubElement(xfrm, f'{{{A_NS}}}ext')
    ext.set('cx', str(cx))
    ext.set('cy', str(cy))
    blipFill = etree.SubElement(pic, f'{{{A_NS}}}blipFill')
    return pic


def _make_page_number_shape(x: int, y: int, num_fields: int = 1,
                             doubled_text: str = None) -> etree._Element:
    """Build a shape containing page number field(s)."""
    sp = _make_shape(x, y, 500000, 200000)
    txBody = etree.SubElement(sp, f'{{{P_NS}}}txBody')
    p = etree.SubElement(txBody, f'{{{A_NS}}}p')

    for _ in range(num_fields):
        fld = etree.SubElement(p, f'{{{A_NS}}}fld')
        fld.set('type', 'slidenum')
        fld.set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', 'rId1')
        t = etree.SubElement(fld, f'{{{A_NS}}}t')
        t.text = '1'

    if doubled_text:
        # Replace field text with doubled pattern
        for fld in p.iter(f'{{{A_NS}}}fld'):
            p.remove(fld)
        r = etree.SubElement(p, f'{{{A_NS}}}r')
        t = etree.SubElement(r, f'{{{A_NS}}}t')
        t.text = doubled_text

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
# CHECK #3: ICON_IN_WRONG_TABLE_CELL
# ─────────────────────────────────────────────────────────────────────────────

class TestCheckIconTableCorrespondence:

    def test_detects_unmirrored_icon(self):
        """Icon at original x=1000000 should be at slide_width-1000000-cx."""
        orig_pic = _make_pic(1000000, 500000, 200000, 200000)
        conv_pic = _make_pic(1000000, 500000, 200000, 200000)  # NOT mirrored

        orig_slide = _wrap_in_slide(orig_pic)
        conv_slide = _wrap_in_slide(conv_pic)

        checker = V3XMLChecker()
        defects = checker._check_icon_table_correspondence(1, orig_slide, conv_slide)
        assert len(defects) >= 1
        assert defects[0].code == "ICON_IN_WRONG_TABLE_CELL"

    def test_passes_mirrored_icon(self):
        """Icon at correct mirrored position should pass."""
        orig_x = 1000000
        cx = 200000
        expected_x = SLIDE_W - orig_x - cx  # 7944000

        orig_pic = _make_pic(orig_x, 500000, cx, 200000)
        conv_pic = _make_pic(expected_x, 500000, cx, 200000)

        orig_slide = _wrap_in_slide(orig_pic)
        conv_slide = _wrap_in_slide(conv_pic)

        checker = V3XMLChecker()
        defects = checker._check_icon_table_correspondence(1, orig_slide, conv_slide)
        assert len(defects) == 0

    def test_skips_when_no_pics(self):
        """No pics → no defects."""
        orig_slide = _wrap_in_slide(_make_shape(0, 0, 100000, 100000))
        conv_slide = _wrap_in_slide(_make_shape(0, 0, 100000, 100000))

        checker = V3XMLChecker()
        defects = checker._check_icon_table_correspondence(1, orig_slide, conv_slide)
        assert len(defects) == 0


# ─────────────────────────────────────────────────────────────────────────────
# CHECK #4: PAGE_NUMBER_DUPLICATED / PAGE_NUMBER_DOUBLED_STRING
# ─────────────────────────────────────────────────────────────────────────────

class TestCheckPageNumberDuplication:

    def test_detects_duplicate_fields(self):
        """Two slidenum fields in same paragraph = defect."""
        sp = _make_page_number_shape(100000, 6000000, num_fields=2)
        slide = _wrap_in_slide(sp)

        checker = V3XMLChecker()
        defects = checker._check_page_number_duplication(1, slide)
        assert len(defects) >= 1
        assert defects[0].code == "PAGE_NUMBER_DUPLICATED"

    def test_passes_single_field(self):
        """One slidenum field = OK."""
        sp = _make_page_number_shape(100000, 6000000, num_fields=1)
        slide = _wrap_in_slide(sp)

        checker = V3XMLChecker()
        defects = checker._check_page_number_duplication(1, slide)
        assert len(defects) == 0

    def test_detects_doubled_string(self):
        """Text like '1515' detected as doubled page number."""
        sp = _make_page_number_shape(100000, 6000000, doubled_text='1515')
        slide = _wrap_in_slide(sp)

        checker = V3XMLChecker()
        defects = checker._check_page_number_duplication(1, slide)
        doubled = [d for d in defects if d.code == 'PAGE_NUMBER_DOUBLED_STRING']
        assert len(doubled) >= 1

    def test_ignores_normal_text(self):
        """Normal text should not trigger doubled-string check."""
        sp = _make_shape(100000, 6000000, 500000, 200000, text='Slide 5')
        slide = _wrap_in_slide(sp)

        checker = V3XMLChecker()
        defects = checker._check_page_number_duplication(1, slide)
        assert len(defects) == 0


# ─────────────────────────────────────────────────────────────────────────────
# CHECK #5: SHAPE_NOT_MIRRORED_POSITION
# ─────────────────────────────────────────────────────────────────────────────

class TestCheckShapePositionMirroring:

    def test_detects_unmirrored_shape(self):
        """Shape not at mirrored x position should be flagged."""
        orig_x = 500000
        cx = 2000000
        orig_sp = _make_shape(orig_x, 1000000, cx, 1500000, text='Hello')
        conv_sp = _make_shape(orig_x, 1000000, cx, 1500000, text='مرحبا')  # Not mirrored

        orig_slide = _wrap_in_slide(orig_sp)
        conv_slide = _wrap_in_slide(conv_sp)

        checker = V3XMLChecker()
        defects = checker._check_shape_position_mirroring(1, orig_slide, conv_slide)
        assert len(defects) >= 1
        assert defects[0].code == "SHAPE_NOT_MIRRORED_POSITION"

    def test_passes_mirrored_shape(self):
        """Correctly mirrored shape should pass."""
        orig_x = 500000
        cx = 2000000
        expected_x = SLIDE_W - orig_x - cx
        orig_sp = _make_shape(orig_x, 1000000, cx, 1500000, text='Hello')
        conv_sp = _make_shape(expected_x, 1000000, cx, 1500000, text='مرحبا')

        orig_slide = _wrap_in_slide(orig_sp)
        conv_slide = _wrap_in_slide(conv_sp)

        checker = V3XMLChecker()
        defects = checker._check_shape_position_mirroring(1, orig_slide, conv_slide)
        assert len(defects) == 0

    def test_skips_centered_shapes(self):
        """Centered shapes should be skipped."""
        cx = 2000000
        center_x = (SLIDE_W - cx) // 2
        orig_sp = _make_shape(center_x, 1000000, cx, 1500000)
        conv_sp = _make_shape(center_x, 1000000, cx, 1500000)

        orig_slide = _wrap_in_slide(orig_sp)
        conv_slide = _wrap_in_slide(conv_sp)

        checker = V3XMLChecker()
        defects = checker._check_shape_position_mirroring(1, orig_slide, conv_slide)
        assert len(defects) == 0

    def test_skips_fullwidth_shapes(self):
        """Full-width background shapes should be skipped."""
        cx = int(SLIDE_W * 0.9)  # 90% of slide width
        orig_sp = _make_shape(100000, 0, cx, SLIDE_H)
        conv_sp = _make_shape(100000, 0, cx, SLIDE_H)

        orig_slide = _wrap_in_slide(orig_sp)
        conv_slide = _wrap_in_slide(conv_sp)

        checker = V3XMLChecker()
        defects = checker._check_shape_position_mirroring(1, orig_slide, conv_slide)
        assert len(defects) == 0

    def test_skips_tiny_shapes(self):
        """Very small decorative shapes should be skipped."""
        orig_sp = _make_shape(100000, 100000, 40000, 40000)
        conv_sp = _make_shape(100000, 100000, 40000, 40000)

        orig_slide = _wrap_in_slide(orig_sp)
        conv_slide = _wrap_in_slide(conv_sp)

        checker = V3XMLChecker()
        defects = checker._check_shape_position_mirroring(1, orig_slide, conv_slide)
        assert len(defects) == 0


# ─────────────────────────────────────────────────────────────────────────────
# CHECK #6: PARAGRAPH_RTL_MISSING
# ─────────────────────────────────────────────────────────────────────────────

class TestCheckParagraphRtl:

    def test_detects_missing_rtl_on_arabic_shape(self):
        """Arabic text in shape without rtl='1' should be flagged."""
        sp = _make_shape(100000, 100000, 3000000, 1000000, text='مرحبا بالعالم')
        slide = _wrap_in_slide(sp)

        checker = V3XMLChecker()
        defects = checker._check_paragraph_rtl(1, slide)
        assert len(defects) >= 1
        assert defects[0].code == "PARAGRAPH_RTL_MISSING"

    def test_passes_with_rtl_set(self):
        """Arabic text with rtl='1' should pass."""
        sp = _make_shape(100000, 100000, 3000000, 1000000,
                        text='مرحبا بالعالم', rtl='1')
        slide = _wrap_in_slide(sp)

        checker = V3XMLChecker()
        defects = checker._check_paragraph_rtl(1, slide)
        assert len(defects) == 0

    def test_ignores_english_text(self):
        """English text should not be flagged."""
        sp = _make_shape(100000, 100000, 3000000, 1000000, text='Hello World')
        slide = _wrap_in_slide(sp)

        checker = V3XMLChecker()
        defects = checker._check_paragraph_rtl(1, slide)
        assert len(defects) == 0

    def test_one_defect_per_shape(self):
        """Multiple Arabic paragraphs in same shape = 1 defect only."""
        sp = _make_shape(100000, 100000, 3000000, 1000000, text='مرحبا')
        # Add second paragraph
        txBody = sp.find(f'.//{{{P_NS}}}txBody')
        if txBody is None:
            txBody = sp.find(f'.//{{{A_NS}}}txBody')
        p2 = etree.SubElement(txBody, f'{{{A_NS}}}p')
        r2 = etree.SubElement(p2, f'{{{A_NS}}}r')
        t2 = etree.SubElement(r2, f'{{{A_NS}}}t')
        t2.text = 'نص آخر'

        slide = _wrap_in_slide(sp)

        checker = V3XMLChecker()
        defects = checker._check_paragraph_rtl(1, slide)
        assert len(defects) == 1  # Not 2


# ─────────────────────────────────────────────────────────────────────────────
# FIX #3: reposition_icon
# ─────────────────────────────────────────────────────────────────────────────

class TestFixRepositionIcon:

    def test_repositions_icon(self):
        """Icon should be moved to expected_x."""
        pic = _make_pic(1000000, 500000, 200000, 200000)
        slide = _wrap_in_slide(pic)

        defect = V3Defect(
            code="ICON_IN_WRONG_TABLE_CELL",
            category="icon",
            severity=Severity.HIGH,
            slide_idx=1,
            object_id="0",
            evidence={'expected_x': 7944000},
            fixable=True,
            autofix_action="reposition_icon",
        )

        fixer = V3AutoFixer()
        assert fixer.apply_fix(slide, defect) is True

        off = slide.find(f'.//{{{A_NS}}}off')
        assert off.get('x') == '7944000'


# ─────────────────────────────────────────────────────────────────────────────
# FIX #4: dedup_page_number
# ─────────────────────────────────────────────────────────────────────────────

class TestFixDedupPageNumber:

    def test_removes_duplicate_fields(self):
        """Should keep first slidenum field, remove second."""
        sp = _make_page_number_shape(100000, 6000000, num_fields=2)
        slide = _wrap_in_slide(sp)

        defect = V3Defect(
            code="PAGE_NUMBER_DUPLICATED",
            category="numbering",
            severity=Severity.HIGH,
            slide_idx=1,
            object_id="0",
            evidence={'field_count': 2},
            fixable=True,
            autofix_action="dedup_page_number",
        )

        fixer = V3AutoFixer()
        assert fixer.apply_fix(slide, defect) is True

        # Verify only one field remains
        remaining_fields = [f for f in slide.iter(f'{{{A_NS}}}fld')
                           if 'slidenum' in (f.get('type', '') or '').lower()]
        assert len(remaining_fields) == 1


# ─────────────────────────────────────────────────────────────────────────────
# FIX #5: mirror_shape_position
# ─────────────────────────────────────────────────────────────────────────────

class TestFixMirrorShapePosition:

    def test_moves_shape_to_expected_x(self):
        """Shape should be repositioned to expected_x."""
        sp = _make_shape(500000, 1000000, 2000000, 1500000)
        slide = _wrap_in_slide(sp)

        expected_x = SLIDE_W - 500000 - 2000000  # 6644000

        defect = V3Defect(
            code="SHAPE_NOT_MIRRORED_POSITION",
            category="mirroring",
            severity=Severity.HIGH,
            slide_idx=1,
            object_id="0",
            evidence={'expected_x': expected_x},
            fixable=True,
            autofix_action="mirror_shape_position",
        )

        fixer = V3AutoFixer()
        assert fixer.apply_fix(slide, defect) is True

        off = slide.find(f'.//{{{A_NS}}}off')
        assert off.get('x') == str(expected_x)


# ─────────────────────────────────────────────────────────────────────────────
# FIX #6: set_paragraph_rtl
# ─────────────────────────────────────────────────────────────────────────────

class TestFixSetParagraphRtl:

    def test_sets_rtl_on_arabic_paragraph(self):
        """Arabic paragraph should get rtl='1' and algn='r'."""
        sp = _make_shape(100000, 100000, 3000000, 1000000, text='مرحبا')
        slide = _wrap_in_slide(sp)

        defect = V3Defect(
            code="PARAGRAPH_RTL_MISSING",
            category="alignment",
            severity=Severity.MEDIUM,
            slide_idx=1,
            object_id="0",
            fixable=True,
            autofix_action="set_paragraph_rtl",
        )

        fixer = V3AutoFixer()
        assert fixer.apply_fix(slide, defect) is True

        # Verify rtl='1' is set
        pPr = slide.find(f'.//{{{A_NS}}}pPr')
        assert pPr is not None
        assert pPr.get('rtl') == '1'

    def test_preserves_centered_alignment(self):
        """Centered text should keep algn='ctr' even after fix."""
        sp = _make_shape(100000, 100000, 3000000, 1000000,
                        text='مرحبا', algn='ctr')
        slide = _wrap_in_slide(sp)

        defect = V3Defect(
            code="PARAGRAPH_RTL_MISSING",
            category="alignment",
            severity=Severity.MEDIUM,
            slide_idx=1,
            object_id="0",
            fixable=True,
            autofix_action="set_paragraph_rtl",
        )

        fixer = V3AutoFixer()
        fixer.apply_fix(slide, defect)

        pPr = slide.find(f'.//{{{A_NS}}}pPr')
        assert pPr.get('algn') == 'ctr'  # Preserved

    def test_no_change_on_english(self):
        """English-only shape should not be modified."""
        sp = _make_shape(100000, 100000, 3000000, 1000000, text='Hello')
        slide = _wrap_in_slide(sp)

        defect = V3Defect(
            code="PARAGRAPH_RTL_MISSING",
            category="alignment",
            severity=Severity.MEDIUM,
            slide_idx=1,
            object_id="0",
            fixable=True,
            autofix_action="set_paragraph_rtl",
        )

        fixer = V3AutoFixer()
        assert fixer.apply_fix(slide, defect) is False


# ─────────────────────────────────────────────────────────────────────────────
# ROUND-TRIP REGRESSION
# ─────────────────────────────────────────────────────────────────────────────

class TestSprint2RoundTrip:

    def test_fix6_then_recheck_yields_zero(self):
        """Apply paragraph RTL fix, recheck → 0 defects."""
        sp = _make_shape(100000, 100000, 3000000, 1000000, text='مرحبا بالعالم')
        slide = _wrap_in_slide(sp)

        checker = V3XMLChecker()
        fixer = V3AutoFixer()

        defects = checker._check_paragraph_rtl(1, slide)
        assert len(defects) >= 1

        for d in defects:
            fixer.apply_fix(slide, d)

        defects_after = checker._check_paragraph_rtl(1, slide)
        assert len(defects_after) == 0

    def test_fix5_then_recheck_yields_zero(self):
        """Apply shape mirror fix, recheck → 0 defects."""
        orig_x = 500000
        cx = 2000000
        expected_x = SLIDE_W - orig_x - cx

        orig_sp = _make_shape(orig_x, 1000000, cx, 1500000, text='Hello')
        conv_sp = _make_shape(orig_x, 1000000, cx, 1500000, text='مرحبا')

        orig_slide = _wrap_in_slide(orig_sp)
        conv_slide = _wrap_in_slide(conv_sp)

        checker = V3XMLChecker()
        fixer = V3AutoFixer()

        defects = checker._check_shape_position_mirroring(1, orig_slide, conv_slide)
        assert len(defects) >= 1

        for d in defects:
            fixer.apply_fix(conv_slide, d)

        defects_after = checker._check_shape_position_mirroring(1, orig_slide, conv_slide)
        assert len(defects_after) == 0
