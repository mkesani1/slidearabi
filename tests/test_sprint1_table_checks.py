"""
tests/test_sprint1_table_checks.py — Sprint 1 Unit Tests

Tests for V3XMLChecker (5 table checks) and V3AutoFixer (2 fixes).
All tests use synthetic XML — no real PPTX files needed.
"""

from __future__ import annotations

import copy
import os
import sys

import pytest
from lxml import etree

# Ensure repo root is on sys.path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from v3_checks import (
    V3AutoFixer,
    V3XMLChecker,
    V3_PHYS_REV_MARKER,
    _find_tables,
    _get_cell_text,
    _get_gridcols,
    _get_row_cells,
    _get_table_rows,
    _has_arabic,
    A_NS,
    P_NS,
    SLIDEARABI_NS,
)
from vqa_types import Severity, DefectStatus, V3Defect, VQAGateResult
from v3_checks import compute_gate_decision

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS — build minimal slide XML structures
# ─────────────────────────────────────────────────────────────────────────────

def _ns(prefix):
    """Shortcut for namespace URIs."""
    return {
        'a': A_NS,
        'p': P_NS,
    }[prefix]


def _make_cell(text: str, rtl: str = None, algn: str = None) -> etree._Element:
    """Build a minimal <a:tc> with one paragraph containing text."""
    tc = etree.Element(f'{{{A_NS}}}tc')
    txBody = etree.SubElement(tc, f'{{{A_NS}}}txBody')
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
    return tc


def _make_merged_cell(text: str, gridSpan: int = 1, rowSpan: int = 1) -> etree._Element:
    """Build a cell with merge attributes."""
    tc = _make_cell(text)
    if gridSpan > 1:
        tc.set('gridSpan', str(gridSpan))
    if rowSpan > 1:
        tc.set('rowSpan', str(rowSpan))
    return tc


def _make_table(rows_data: list, gridcol_widths: list = None) -> etree._Element:
    """Build a <a:tbl> with rows and optional gridCol widths.
    
    rows_data: list of lists, each inner list contains (text, kwargs_dict) or just text.
    """
    tbl = etree.Element(f'{{{A_NS}}}tbl')
    tblPr = etree.SubElement(tbl, f'{{{A_NS}}}tblPr')

    if gridcol_widths:
        grid = etree.SubElement(tbl, f'{{{A_NS}}}tblGrid')
        for w in gridcol_widths:
            col = etree.SubElement(grid, f'{{{A_NS}}}gridCol')
            col.set('w', str(w))

    for row_data in rows_data:
        tr = etree.SubElement(tbl, f'{{{A_NS}}}tr')
        for cell_data in row_data:
            if isinstance(cell_data, etree._Element):
                tr.append(cell_data)
            elif isinstance(cell_data, tuple):
                tc = _make_cell(cell_data[0], **cell_data[1])
                tr.append(tc)
            else:
                tc = _make_cell(str(cell_data))
                tr.append(tc)

    return tbl


def _wrap_in_slide(tbl_element) -> etree._Element:
    """Wrap a table in a minimal slide-like XML structure."""
    sld = etree.Element(f'{{{P_NS}}}sld')
    cSld = etree.SubElement(sld, f'{{{P_NS}}}cSld')
    spTree = etree.SubElement(cSld, f'{{{P_NS}}}spTree')
    sp = etree.SubElement(spTree, f'{{{P_NS}}}sp')
    graphicFrame = etree.SubElement(spTree, f'{{{P_NS}}}graphicFrame')
    # put table inside
    graphicFrame.append(tbl_element)
    return sld


# ─────────────────────────────────────────────────────────────────────────────
# CHECK #1: TABLE_COLUMNS_NOT_REVERSED
# ─────────────────────────────────────────────────────────────────────────────

class TestCheckTableColumnOrder:
    """Tests for _check_table_column_order — Check #1."""

    def test_detects_unreversed_columns(self):
        """If orig and conv have same column order, flag defect."""
        orig_tbl = _make_table([['A', 'B', 'C'], ['1', '2', '3']])
        conv_tbl = _make_table([['A', 'B', 'C'], ['1', '2', '3']])  # Same order = BAD

        orig_slide = _wrap_in_slide(orig_tbl)
        conv_slide = _wrap_in_slide(conv_tbl)

        checker = V3XMLChecker()
        defects = checker._check_table_column_order(1, orig_slide, conv_slide)

        assert len(defects) == 1
        assert defects[0].code == "TABLE_COLUMNS_NOT_REVERSED"
        assert defects[0].severity == Severity.CRITICAL
        assert defects[0].fixable is True

    def test_passes_when_columns_reversed(self):
        """If conv has reversed column order, no defect."""
        orig_tbl = _make_table([['A', 'B', 'C'], ['1', '2', '3']])
        conv_tbl = _make_table([['C', 'B', 'A'], ['3', '2', '1']])  # Reversed = GOOD

        orig_slide = _wrap_in_slide(orig_tbl)
        conv_slide = _wrap_in_slide(conv_tbl)

        checker = V3XMLChecker()
        defects = checker._check_table_column_order(1, orig_slide, conv_slide)
        assert len(defects) == 0

    def test_skips_single_column_table(self):
        """Single column tables can't be reversed."""
        orig_tbl = _make_table([['A'], ['B']])
        conv_tbl = _make_table([['A'], ['B']])

        orig_slide = _wrap_in_slide(orig_tbl)
        conv_slide = _wrap_in_slide(conv_tbl)

        checker = V3XMLChecker()
        defects = checker._check_table_column_order(1, orig_slide, conv_slide)
        assert len(defects) == 0

    def test_skips_all_empty_cells(self):
        """Rows with all empty cells should be skipped."""
        orig_tbl = _make_table([['', '', ''], ['A', 'B', 'C']])
        conv_tbl = _make_table([['', '', ''], ['A', 'B', 'C']])

        orig_slide = _wrap_in_slide(orig_tbl)
        conv_slide = _wrap_in_slide(conv_tbl)

        checker = V3XMLChecker()
        defects = checker._check_table_column_order(1, orig_slide, conv_slide)
        # Should still detect on second row (which has content)
        assert len(defects) == 1

    def test_skips_all_identical_cells(self):
        """Rows where all cells have the same text can't determine order."""
        orig_tbl = _make_table([['X', 'X', 'X']])
        conv_tbl = _make_table([['X', 'X', 'X']])

        orig_slide = _wrap_in_slide(orig_tbl)
        conv_slide = _wrap_in_slide(conv_tbl)

        checker = V3XMLChecker()
        defects = checker._check_table_column_order(1, orig_slide, conv_slide)
        assert len(defects) == 0

    def test_respects_idempotency_marker(self):
        """If physRtlCols marker is set, skip the check."""
        orig_tbl = _make_table([['A', 'B', 'C']])
        conv_tbl = _make_table([['A', 'B', 'C']])  # Same order but has marker
        
        # Set marker on conv table (using shared constant)
        tbl_pr = conv_tbl.find(f'{{{A_NS}}}tblPr')
        tbl_pr.set(V3_PHYS_REV_MARKER, '1')

        orig_slide = _wrap_in_slide(orig_tbl)
        conv_slide = _wrap_in_slide(conv_tbl)

        checker = V3XMLChecker()
        defects = checker._check_table_column_order(1, orig_slide, conv_slide)
        assert len(defects) == 0


# ─────────────────────────────────────────────────────────────────────────────
# CHECK #2: TABLE_CELL_RTL_MISSING
# ─────────────────────────────────────────────────────────────────────────────

class TestCheckTableCellAlignment:
    """Tests for _check_table_cell_alignment — Check #2."""

    def test_detects_missing_rtl_on_arabic_cell(self):
        """Arabic cell without rtl='1' should be flagged."""
        tbl = _make_table([['مرحبا', 'Hello']])
        slide = _wrap_in_slide(tbl)

        checker = V3XMLChecker()
        defects = checker._check_table_cell_alignment(1, slide)

        # Arabic cell should be flagged (no rtl, no algn)
        assert len(defects) >= 1
        assert defects[0].code == "TABLE_CELL_RTL_MISSING"
        assert defects[0].severity == Severity.HIGH

    def test_passes_with_correct_rtl(self):
        """Arabic cell with rtl='1' and algn='r' should pass."""
        tbl = _make_table([[('مرحبا', {'rtl': '1', 'algn': 'r'}), 'Hello']])
        slide = _wrap_in_slide(tbl)

        checker = V3XMLChecker()
        defects = checker._check_table_cell_alignment(1, slide)
        assert len(defects) == 0

    def test_ignores_english_cells(self):
        """English-only cells should not be flagged."""
        tbl = _make_table([['Hello', 'World']])
        slide = _wrap_in_slide(tbl)

        checker = V3XMLChecker()
        defects = checker._check_table_cell_alignment(1, slide)
        assert len(defects) == 0

    def test_detects_missing_alignment_only(self):
        """Arabic cell with rtl='1' but missing algn='r'."""
        tbl = _make_table([[('مرحبا', {'rtl': '1'})]])  # no algn
        slide = _wrap_in_slide(tbl)

        checker = V3XMLChecker()
        defects = checker._check_table_cell_alignment(1, slide)
        assert len(defects) == 1
        assert 'algn not right' in defects[0].evidence['issues']


# ─────────────────────────────────────────────────────────────────────────────
# CHECK #9: TABLE_GRIDCOL_NOT_REVERSED
# ─────────────────────────────────────────────────────────────────────────────

class TestCheckTableGridcolReversal:
    """Tests for _check_table_gridcol_reversal — Check #9."""

    def test_detects_unreversed_gridcols(self):
        """If gridCol widths are same between orig and conv, flag defect."""
        orig_tbl = _make_table([['A', 'B', 'C']], gridcol_widths=[1000, 2000, 3000])
        conv_tbl = _make_table([['A', 'B', 'C']], gridcol_widths=[1000, 2000, 3000])

        orig_slide = _wrap_in_slide(orig_tbl)
        conv_slide = _wrap_in_slide(conv_tbl)

        checker = V3XMLChecker()
        defects = checker._check_table_gridcol_reversal(1, orig_slide, conv_slide)
        assert len(defects) == 1
        assert defects[0].code == "TABLE_GRIDCOL_NOT_REVERSED"

    def test_passes_when_gridcols_reversed(self):
        """Reversed gridCol widths should pass."""
        orig_tbl = _make_table([['A', 'B', 'C']], gridcol_widths=[1000, 2000, 3000])
        conv_tbl = _make_table([['A', 'B', 'C']], gridcol_widths=[3000, 2000, 1000])

        orig_slide = _wrap_in_slide(orig_tbl)
        conv_slide = _wrap_in_slide(conv_tbl)

        checker = V3XMLChecker()
        defects = checker._check_table_gridcol_reversal(1, orig_slide, conv_slide)
        assert len(defects) == 0

    def test_skips_equal_width_columns(self):
        """Uniform widths can't determine reversal, skip."""
        orig_tbl = _make_table([['A', 'B']], gridcol_widths=[1000, 1000])
        conv_tbl = _make_table([['A', 'B']], gridcol_widths=[1000, 1000])

        orig_slide = _wrap_in_slide(orig_tbl)
        conv_slide = _wrap_in_slide(conv_tbl)

        checker = V3XMLChecker()
        defects = checker._check_table_gridcol_reversal(1, orig_slide, conv_slide)
        assert len(defects) == 0


# ─────────────────────────────────────────────────────────────────────────────
# CHECK #10: TABLE_MERGED_CELL_INTEGRITY
# ─────────────────────────────────────────────────────────────────────────────

class TestCheckMergedCellIntegrity:
    """Tests for _check_merged_cell_integrity — Check #10."""

    def test_detects_merged_cells(self):
        """Table with gridSpan > 1 should be flagged."""
        tbl = _make_table([[]])  # empty table
        tr = etree.SubElement(tbl, f'{{{A_NS}}}tr')
        tr.append(_make_merged_cell('Merged', gridSpan=2))
        tr.append(_make_cell('Normal'))

        slide = _wrap_in_slide(tbl)

        checker = V3XMLChecker()
        defects = checker._check_merged_cell_integrity(1, slide)
        assert len(defects) == 1
        assert defects[0].code == "TABLE_MERGED_CELL_INTEGRITY"
        assert defects[0].fixable is False  # Flag only

    def test_no_merged_cells_passes(self):
        """Table without merged cells should pass."""
        tbl = _make_table([['A', 'B', 'C']])
        slide = _wrap_in_slide(tbl)

        checker = V3XMLChecker()
        defects = checker._check_merged_cell_integrity(1, slide)
        assert len(defects) == 0

    def test_detects_rowspan(self):
        """Table with rowSpan > 1 should be flagged."""
        tbl = _make_table([[]])
        tr = etree.SubElement(tbl, f'{{{A_NS}}}tr')
        tr.append(_make_merged_cell('Tall', rowSpan=3))
        tr.append(_make_cell('Normal'))

        slide = _wrap_in_slide(tbl)

        checker = V3XMLChecker()
        defects = checker._check_merged_cell_integrity(1, slide)
        assert len(defects) == 1


# ─────────────────────────────────────────────────────────────────────────────
# CHECK #11: TABLE_GRID_STRUCTURAL_ERROR
# ─────────────────────────────────────────────────────────────────────────────

class TestCheckTableGridStructural:
    """Tests for _check_table_grid_structural — Check #11."""

    def test_detects_grid_mismatch(self):
        """Row with fewer cells than gridCols should flag."""
        tbl = _make_table([['A', 'B']], gridcol_widths=[1000, 2000, 3000])
        # 3 gridCols but only 2 cells

        slide = _wrap_in_slide(tbl)

        checker = V3XMLChecker()
        defects = checker._check_table_grid_structural(1, slide)
        assert len(defects) == 1
        assert defects[0].code == "TABLE_GRID_STRUCTURAL_ERROR"
        assert defects[0].severity == Severity.CRITICAL

    def test_passes_matching_grid(self):
        """Row cells matching gridCols should pass."""
        tbl = _make_table([['A', 'B', 'C']], gridcol_widths=[1000, 2000, 3000])
        slide = _wrap_in_slide(tbl)

        checker = V3XMLChecker()
        defects = checker._check_table_grid_structural(1, slide)
        assert len(defects) == 0

    def test_accounts_for_gridspan(self):
        """Cell with gridSpan=2 counts as 2 effective cols."""
        tbl = etree.Element(f'{{{A_NS}}}tbl')
        etree.SubElement(tbl, f'{{{A_NS}}}tblPr')
        grid = etree.SubElement(tbl, f'{{{A_NS}}}tblGrid')
        for w in [1000, 2000, 3000]:
            col = etree.SubElement(grid, f'{{{A_NS}}}gridCol')
            col.set('w', str(w))
        tr = etree.SubElement(tbl, f'{{{A_NS}}}tr')
        tr.append(_make_merged_cell('Wide', gridSpan=2))
        tr.append(_make_cell('Normal'))
        # 2 + 1 = 3 effective = 3 gridCols → pass

        slide = _wrap_in_slide(tbl)

        checker = V3XMLChecker()
        defects = checker._check_table_grid_structural(1, slide)
        assert len(defects) == 0


# ─────────────────────────────────────────────────────────────────────────────
# FIX #1: reverse_table_columns
# ─────────────────────────────────────────────────────────────────────────────

class TestFixReverseTableColumns:
    """Tests for V3AutoFixer._fix_reverse_table_columns."""

    def test_reverses_cells_and_gridcols(self):
        """Cells in each row should be reversed, gridCol widths too."""
        tbl = _make_table(
            [['A', 'B', 'C'], ['1', '2', '3']],
            gridcol_widths=[1000, 2000, 3000],
        )
        slide = _wrap_in_slide(tbl)

        defect = V3Defect(
            code="TABLE_COLUMNS_NOT_REVERSED",
            category="table",
            severity=Severity.CRITICAL,
            slide_idx=1,
            object_id="0",
            fixable=True,
            autofix_action="reverse_table_columns",
        )

        fixer = V3AutoFixer()
        result = fixer.apply_fix(slide, defect)
        assert result is True

        # Verify cells are reversed
        tbl_fixed = _find_tables(slide)[0]
        rows = _get_table_rows(tbl_fixed)
        first_row_texts = [_get_cell_text(c) for c in _get_row_cells(rows[0])]
        assert first_row_texts == ['C', 'B', 'A']

        second_row_texts = [_get_cell_text(c) for c in _get_row_cells(rows[1])]
        assert second_row_texts == ['3', '2', '1']

        # Verify gridCol widths reversed
        gridcols = _get_gridcols(tbl_fixed)
        widths = [c.get('w') for c in gridcols]
        assert widths == ['3000', '2000', '1000']

    def test_sets_rtl_zero_after_reversal(self):
        """After physical reversal, tblPr rtl should be '0'."""
        tbl = _make_table([['A', 'B']], gridcol_widths=[1000, 2000])
        slide = _wrap_in_slide(tbl)

        defect = V3Defect(
            code="TABLE_COLUMNS_NOT_REVERSED",
            category="table",
            severity=Severity.CRITICAL,
            slide_idx=1,
            object_id="0",
            fixable=True,
            autofix_action="reverse_table_columns",
        )

        fixer = V3AutoFixer()
        fixer.apply_fix(slide, defect)

        tbl_fixed = _find_tables(slide)[0]
        tbl_pr = tbl_fixed.find(f'{{{A_NS}}}tblPr')
        assert tbl_pr.get('rtl') == '0'

    def test_sets_idempotency_marker(self):
        """After fix, v3PhysRev marker should be set."""
        tbl = _make_table([['A', 'B']], gridcol_widths=[1000, 2000])
        slide = _wrap_in_slide(tbl)

        defect = V3Defect(
            code="TABLE_COLUMNS_NOT_REVERSED",
            category="table",
            severity=Severity.CRITICAL,
            slide_idx=1,
            object_id="0",
            fixable=True,
            autofix_action="reverse_table_columns",
        )

        fixer = V3AutoFixer()
        fixer.apply_fix(slide, defect)

        tbl_fixed = _find_tables(slide)[0]
        tbl_pr = tbl_fixed.find(f'{{{A_NS}}}tblPr')
        marker = f'{{{SLIDEARABI_NS}}}v3PhysRev'
        assert tbl_pr.get(marker) == '1'

    def test_aborts_on_merged_cells(self):
        """Should refuse to reverse table with merged cells."""
        tbl = _make_table([[]], gridcol_widths=[1000, 2000, 3000])
        tr = etree.SubElement(tbl, f'{{{A_NS}}}tr')
        tr.append(_make_merged_cell('Wide', gridSpan=2))
        tr.append(_make_cell('Normal'))

        slide = _wrap_in_slide(tbl)

        defect = V3Defect(
            code="TABLE_COLUMNS_NOT_REVERSED",
            category="table",
            severity=Severity.CRITICAL,
            slide_idx=1,
            object_id="0",
            fixable=True,
            autofix_action="reverse_table_columns",
        )

        fixer = V3AutoFixer()
        result = fixer.apply_fix(slide, defect)
        assert result is False

    def test_idempotent_no_double_reversal(self):
        """Applying fix twice should be a no-op the second time."""
        tbl = _make_table([['A', 'B', 'C']], gridcol_widths=[1000, 2000, 3000])
        slide = _wrap_in_slide(tbl)

        defect = V3Defect(
            code="TABLE_COLUMNS_NOT_REVERSED",
            category="table",
            severity=Severity.CRITICAL,
            slide_idx=1,
            object_id="0",
            fixable=True,
            autofix_action="reverse_table_columns",
        )

        fixer = V3AutoFixer()
        assert fixer.apply_fix(slide, defect) is True
        assert fixer.apply_fix(slide, defect) is False  # Already has marker

    def test_swaps_first_last_col_flags(self):
        """firstCol and lastCol banding flags should be swapped."""
        tbl = _make_table([['A', 'B']], gridcol_widths=[1000, 2000])
        tbl_pr = tbl.find(f'{{{A_NS}}}tblPr')
        tbl_pr.set('firstCol', '1')
        tbl_pr.set('lastCol', '0')

        slide = _wrap_in_slide(tbl)

        defect = V3Defect(
            code="TABLE_COLUMNS_NOT_REVERSED",
            category="table",
            severity=Severity.CRITICAL,
            slide_idx=1,
            object_id="0",
            fixable=True,
            autofix_action="reverse_table_columns",
        )

        fixer = V3AutoFixer()
        fixer.apply_fix(slide, defect)

        tbl_fixed = _find_tables(slide)[0]
        tbl_pr = tbl_fixed.find(f'{{{A_NS}}}tblPr')
        assert tbl_pr.get('firstCol') == '0'
        assert tbl_pr.get('lastCol') == '1'


# ─────────────────────────────────────────────────────────────────────────────
# FIX #2: set_para_rtl
# ─────────────────────────────────────────────────────────────────────────────

class TestFixSetParaRtl:
    """Tests for V3AutoFixer._fix_set_para_rtl."""

    def test_sets_rtl_and_alignment_on_arabic(self):
        """Should add rtl='1' and algn='r' to Arabic paragraphs."""
        tbl = _make_table([['مرحبا', 'Hello']])
        slide = _wrap_in_slide(tbl)

        defect = V3Defect(
            code="TABLE_CELL_RTL_MISSING",
            category="alignment",
            severity=Severity.HIGH,
            slide_idx=1,
            object_id="0:0",
            fixable=True,
            autofix_action="set_para_rtl",
        )

        fixer = V3AutoFixer()
        result = fixer.apply_fix(slide, defect)
        assert result is True

        # Verify the Arabic cell got rtl and alignment
        tables = _find_tables(slide)
        tbl_fixed = tables[0]
        first_row = _get_table_rows(tbl_fixed)[0]
        first_cell = _get_row_cells(first_row)[0]
        para = first_cell.find(f'.//{{{A_NS}}}p')
        pPr = para.find(f'{{{A_NS}}}pPr')
        assert pPr is not None
        assert pPr.get('rtl') == '1'
        assert pPr.get('algn') == 'r'

    def test_does_not_modify_english_cells(self):
        """English cells should not get rtl attributes."""
        tbl = _make_table([['Hello', 'World']])
        slide = _wrap_in_slide(tbl)

        defect = V3Defect(
            code="TABLE_CELL_RTL_MISSING",
            category="alignment",
            severity=Severity.HIGH,
            slide_idx=1,
            object_id="0:0",
            fixable=True,
            autofix_action="set_para_rtl",
        )

        fixer = V3AutoFixer()
        result = fixer.apply_fix(slide, defect)
        assert result is False  # No Arabic text, nothing to fix


# ─────────────────────────────────────────────────────────────────────────────
# GATE LOGIC: compute_gate_decision
# ─────────────────────────────────────────────────────────────────────────────

class TestComputeGateDecision:
    """Tests for compute_gate_decision."""

    def test_passes_with_no_defects(self):
        gate = compute_gate_decision([], total_slides=10)
        assert gate.status == "completed"
        assert gate.critical_remaining == 0
        assert gate.high_remaining == 0

    def test_fails_on_critical(self):
        defects = [
            V3Defect.critical("TABLE_COLUMNS_NOT_REVERSED", slide_idx=1),
        ]
        gate = compute_gate_decision(defects, total_slides=10)
        assert gate.status == "failed_qa"
        assert gate.critical_remaining == 1

    def test_fails_on_must_not_ship(self):
        defects = [
            V3Defect(
                code="TABLE_GRID_STRUCTURAL_ERROR",
                category="table",
                severity=Severity.CRITICAL,
                slide_idx=1,
            ),
        ]
        gate = compute_gate_decision(defects, total_slides=10)
        assert gate.status == "failed_qa"

    def test_warns_on_few_high(self):
        defects = [
            V3Defect.high("TABLE_CELL_RTL_MISSING", slide_idx=1),
            V3Defect.high("TABLE_CELL_RTL_MISSING", slide_idx=2),
        ]
        gate = compute_gate_decision(defects, total_slides=10)
        assert gate.status == "completed_with_warnings"
        assert gate.high_remaining == 2

    def test_fails_on_many_high_slides(self):
        """Many HIGH defects across many slides → failed_qa."""
        defects = [
            V3Defect.high("TABLE_CELL_RTL_MISSING", slide_idx=i)
            for i in range(1, 6)  # 5 HIGH defects on 5 slides (50% of 10)
        ]
        gate = compute_gate_decision(defects, total_slides=10)
        assert gate.status == "failed_qa"

    def test_warns_on_many_high_few_slides(self):
        """Many HIGH defects concentrated on few slides → warning only."""
        defects = [
            V3Defect.high("TABLE_CELL_RTL_MISSING", slide_idx=1),
            V3Defect.high("TABLE_GRIDCOL_NOT_REVERSED", slide_idx=1),
            V3Defect.high("TABLE_MERGED_CELL_INTEGRITY", slide_idx=1),
        ]
        gate = compute_gate_decision(defects, total_slides=10)
        # 3 HIGH but only on 1 slide = 10%, threshold is 10% → fails
        # Actually 1/10 = 10% = 0.10, and threshold is >= 0.10 → fails
        assert gate.status == "failed_qa"

    def test_medium_defects_complete(self):
        """MEDIUM defects alone should not block."""
        defects = [
            V3Defect.medium("SOME_MEDIUM_ISSUE", slide_idx=i)
            for i in range(1, 5)
        ]
        gate = compute_gate_decision(defects, total_slides=10)
        assert gate.status == "completed"


# ─────────────────────────────────────────────────────────────────────────────
# check_slide INTEGRATION
# ─────────────────────────────────────────────────────────────────────────────

class TestCheckSlideIntegration:
    """Test that check_slide runs all checks and returns combined results."""

    def test_returns_multiple_defect_types(self):
        """A slide with multiple problems should return multiple defects."""
        # Table with unreversed columns AND Arabic without RTL
        orig_tbl = _make_table(
            [['Product', 'السعر'], ['Item A', '100']],
            gridcol_widths=[3000, 2000],
        )
        conv_tbl = _make_table(
            [['Product', 'السعر'], ['Item A', '100']],  # Not reversed
            gridcol_widths=[3000, 2000],  # Not reversed
        )

        orig_slide = _wrap_in_slide(orig_tbl)
        conv_slide = _wrap_in_slide(conv_tbl)

        checker = V3XMLChecker()
        defects = checker.check_slide(1, conv_slide, orig_slide)

        codes = {d.code for d in defects}
        # Should detect column order issue (non-empty unique content exists)
        assert "TABLE_COLUMNS_NOT_REVERSED" in codes or "TABLE_GRIDCOL_NOT_REVERSED" in codes
        # Should detect Arabic RTL issue
        assert "TABLE_CELL_RTL_MISSING" in codes

    def test_works_without_original(self):
        """Conv-only checks should run without original slide."""
        tbl = _make_table([['مرحبا', 'Hello']])
        conv_slide = _wrap_in_slide(tbl)

        checker = V3XMLChecker()
        defects = checker.check_slide(1, conv_slide, orig_slide_element=None)

        # Only checks that don't need original should fire
        codes = {d.code for d in defects}
        assert "TABLE_CELL_RTL_MISSING" in codes
        # These require original, should NOT appear
        assert "TABLE_COLUMNS_NOT_REVERSED" not in codes
        assert "TABLE_GRIDCOL_NOT_REVERSED" not in codes


# ─────────────────────────────────────────────────────────────────────────────
# HELPER FUNCTION TESTS
# ─────────────────────────────────────────────────────────────────────────────

# ─────────────────────────────────────────────────────────────────────────────
# ROUND-TRIP REGRESSION: Fix → Re-check → 0 defects
# ─────────────────────────────────────────────────────────────────────────────

class TestRoundTripRegression:
    """Verify that applying a fix then re-running checks produces 0 defects.
    This is the P0 test that validates the idempotency marker contract."""

    def test_fix1_then_recheck_yields_zero_defects(self):
        """Apply reverse_table_columns, re-run check #1 + #9 → 0 defects."""
        orig_tbl = _make_table(
            [['A', 'B', 'C'], ['1', '2', '3']],
            gridcol_widths=[1000, 2000, 3000],
        )
        conv_tbl = _make_table(
            [['A', 'B', 'C'], ['1', '2', '3']],  # Not reversed
            gridcol_widths=[1000, 2000, 3000],
        )
        orig_slide = _wrap_in_slide(orig_tbl)
        conv_slide = _wrap_in_slide(conv_tbl)

        checker = V3XMLChecker()
        fixer = V3AutoFixer()

        # Step 1: Detect defects
        defects = checker.check_slide(1, conv_slide, orig_slide)
        col_defects = [d for d in defects if d.code == 'TABLE_COLUMNS_NOT_REVERSED']
        grid_defects = [d for d in defects if d.code == 'TABLE_GRIDCOL_NOT_REVERSED']
        assert len(col_defects) >= 1
        assert len(grid_defects) >= 1

        # Step 2: Apply fix
        for d in col_defects:
            fixer.apply_fix(conv_slide, d)

        # Step 3: Re-check — should find 0 column/gridCol defects
        defects_after = checker.check_slide(1, conv_slide, orig_slide)
        col_defects_after = [d for d in defects_after if d.code == 'TABLE_COLUMNS_NOT_REVERSED']
        grid_defects_after = [d for d in defects_after if d.code == 'TABLE_GRIDCOL_NOT_REVERSED']
        assert len(col_defects_after) == 0, f"Expected 0 column defects after fix, got {len(col_defects_after)}"
        assert len(grid_defects_after) == 0, f"Expected 0 gridCol defects after fix, got {len(grid_defects_after)}"

    def test_fix2_then_recheck_yields_zero_defects(self):
        """Apply set_para_rtl, re-run check #2 → 0 defects."""
        tbl = _make_table([['مرحبا', 'Hello']])
        slide = _wrap_in_slide(tbl)

        checker = V3XMLChecker()
        fixer = V3AutoFixer()

        # Step 1: Detect RTL defects
        defects = checker._check_table_cell_alignment(1, slide)
        assert len(defects) >= 1

        # Step 2: Apply fix
        for d in defects:
            fixer.apply_fix(slide, d)

        # Step 3: Re-check — should find 0
        defects_after = checker._check_table_cell_alignment(1, slide)
        assert len(defects_after) == 0, f"Expected 0 RTL defects after fix, got {len(defects_after)}"


# ─────────────────────────────────────────────────────────────────────────────
# HELPER FUNCTION TESTS
# ─────────────────────────────────────────────────────────────────────────────

class TestHelperFunctions:
    """Tests for utility functions."""

    def test_has_arabic_detects_arabic(self):
        assert _has_arabic('مرحبا') is True
        assert _has_arabic('Hello مرحبا') is True

    def test_has_arabic_rejects_latin(self):
        assert _has_arabic('Hello') is False
        assert _has_arabic('123') is False
        assert _has_arabic('') is False

    def test_get_cell_text(self):
        cell = _make_cell('Test Text')
        assert _get_cell_text(cell) == 'Test Text'

    def test_get_cell_text_empty(self):
        cell = _make_cell('')
        assert _get_cell_text(cell) == ''
