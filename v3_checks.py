"""
v3_checks.py — V3 XML Structural Checks & Auto-Fixes for SlideArabi

Sprint 1: Table checks (#1, #2, #9, #10, #11) + Fixes (#1, #2)
Sprint 2: Adds icon, page number, shape position, paragraph RTL, overlap
Sprint 3: Adds circular centering, master mirror, directional

All checks produce V3Defect objects. Gated by v3_config flags.
"""

from __future__ import annotations

import logging
import shutil
from copy import deepcopy
from dataclasses import field
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from lxml import etree

logger = logging.getLogger(__name__)

# XML namespaces
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
SLIDEARABI_NS = 'https://slidearabi.ai/ns/transform'

# Arabic character detection
import re
_ARABIC_RE = re.compile(r'[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]')


def _has_arabic(text: str) -> bool:
    return bool(_ARABIC_RE.search(text))


def _get_cell_text(tc_elem) -> str:
    """Extract all text from a table cell element."""
    texts = []
    for t_elem in tc_elem.iter(f'{{{A_NS}}}t'):
        if t_elem.text:
            texts.append(t_elem.text)
    return ' '.join(texts).strip()


def _find_tables(slide_element) -> list:
    """Find all <a:tbl> elements in a slide."""
    return slide_element.findall(f'.//{{{A_NS}}}tbl')


def _get_table_rows(tbl) -> list:
    """Get all <a:tr> elements from a table."""
    return tbl.findall(f'{{{A_NS}}}tr')


def _get_row_cells(tr) -> list:
    """Get all <a:tc> elements from a table row."""
    return tr.findall(f'{{{A_NS}}}tc')


def _get_gridcols(tbl) -> list:
    """Get gridCol elements from table grid."""
    grid = tbl.find(f'{{{A_NS}}}tblGrid')
    if grid is None:
        return []
    return grid.findall(f'{{{A_NS}}}gridCol')


def _get_tbl_pr(tbl):
    """Get tblPr element from table."""
    return tbl.find(f'{{{A_NS}}}tblPr')


# ─────────────────────────────────────────────────────────────────────────────
# V3 XML CHECKS
# ─────────────────────────────────────────────────────────────────────────────

class V3XMLChecker:
    """V3 XML structural checks — produces V3Defect instances."""

    def __init__(self, slide_width: int = 9144000, slide_height: int = 6858000):
        self.slide_width = slide_width
        self.slide_height = slide_height

    def check_slide(
        self,
        slide_num: int,
        conv_slide_element,
        orig_slide_element=None,
    ) -> list:
        """Run all V3 checks on a single slide. Returns list of V3Defect."""
        from vqa_types import V3Defect
        defects: List[V3Defect] = []

        # Table checks (require both original and converted)
        if orig_slide_element is not None:
            defects.extend(self._check_table_column_order(
                slide_num, orig_slide_element, conv_slide_element))
            defects.extend(self._check_table_gridcol_reversal(
                slide_num, orig_slide_element, conv_slide_element))

        # Table checks (converted only)
        defects.extend(self._check_table_cell_alignment(slide_num, conv_slide_element))
        defects.extend(self._check_merged_cell_integrity(slide_num, conv_slide_element))
        defects.extend(self._check_table_grid_structural(slide_num, conv_slide_element))

        return defects

    # ── Check #1: Table Column Order Reversal (CRITICAL) ──

    def _check_table_column_order(
        self, slide_num: int, orig_element, conv_element
    ) -> list:
        """Compare original vs converted table to detect unreversed columns."""
        from vqa_types import V3Defect, Severity, DefectStatus

        defects = []
        orig_tables = _find_tables(orig_element)
        conv_tables = _find_tables(conv_element)

        for tbl_idx, (orig_tbl, conv_tbl) in enumerate(
            zip(orig_tables, conv_tables)
        ):
            orig_rows = _get_table_rows(orig_tbl)
            conv_rows = _get_table_rows(conv_tbl)

            if not orig_rows or not conv_rows:
                continue

            # Use first row with meaningful text content as anchor
            anchor_orig = None
            anchor_conv = None
            for orig_row, conv_row in zip(orig_rows, conv_rows):
                orig_cells = _get_row_cells(orig_row)
                conv_cells = _get_row_cells(conv_row)
                if len(orig_cells) < 2 or len(conv_cells) < 2:
                    continue

                orig_texts = [_get_cell_text(c) for c in orig_cells]
                conv_texts = [_get_cell_text(c) for c in conv_cells]

                # Skip rows where all cells are empty or identical
                non_empty = [t for t in orig_texts if t.strip()]
                if len(non_empty) < 2:
                    continue

                # Check if texts have enough unique content to compare
                unique_orig = set(orig_texts)
                if len(unique_orig) < 2:
                    continue  # All same text — can't determine order

                anchor_orig = orig_texts
                anchor_conv = conv_texts
                break

            if anchor_orig is None or anchor_conv is None:
                continue

            # Check: if original order matches converted (not reversed), flag it
            # After proper RTL conversion, column order should be reversed
            if anchor_orig == anchor_conv:
                # Columns NOT reversed — this is a defect
                # Check idempotency marker
                conv_tbl_pr = _get_tbl_pr(conv_tbl)
                marker = f'{{{SLIDEARABI_NS}}}physRtlCols'
                if conv_tbl_pr is not None and conv_tbl_pr.get(marker) == '1':
                    continue  # Already processed

                defects.append(V3Defect(
                    code="TABLE_COLUMNS_NOT_REVERSED",
                    category="table",
                    severity=Severity.CRITICAL,
                    slide_idx=slide_num,
                    object_id=str(tbl_idx),
                    evidence={
                        'original_order': anchor_orig,
                        'converted_order': anchor_conv,
                        'row_count': len(orig_rows),
                        'col_count': len(anchor_orig),
                    },
                    fixable=True,
                    autofix_action="reverse_table_columns",
                    description=f"Table {tbl_idx+1} columns not reversed for RTL "
                                f"(slide {slide_num})",
                ))

        return defects

    # ── Check #2: Table Cell Alignment + RTL (HIGH) ──

    def _check_table_cell_alignment(self, slide_num: int, conv_element) -> list:
        """Check that Arabic table cells have rtl='1' and algn='r'."""
        from vqa_types import V3Defect, Severity

        defects = []
        tables = _find_tables(conv_element)

        for tbl_idx, tbl in enumerate(tables):
            for row in _get_table_rows(tbl):
                for cell_idx, cell in enumerate(_get_row_cells(row)):
                    cell_text = _get_cell_text(cell)
                    if not _has_arabic(cell_text):
                        continue

                    # Check each paragraph in the cell
                    for para in cell.iter(f'{{{A_NS}}}p'):
                        pPr = para.find(f'{{{A_NS}}}pPr')
                        
                        rtl_ok = False
                        algn_ok = False
                        
                        if pPr is not None:
                            rtl_ok = pPr.get('rtl') == '1'
                            algn_ok = pPr.get('algn') == 'r'

                        if not rtl_ok or not algn_ok:
                            issues = []
                            if not rtl_ok:
                                issues.append('rtl missing')
                            if not algn_ok:
                                issues.append('algn not right')

                            defects.append(V3Defect(
                                code="TABLE_CELL_RTL_MISSING",
                                category="alignment",
                                severity=Severity.HIGH,
                                slide_idx=slide_num,
                                object_id=f"{tbl_idx}:{cell_idx}",
                                evidence={
                                    'cell_text': cell_text[:50],
                                    'issues': issues,
                                },
                                fixable=True,
                                autofix_action="set_para_rtl",
                                description=f"Arabic cell missing RTL/alignment "
                                            f"(slide {slide_num}, table {tbl_idx+1})",
                            ))
                            break  # One defect per cell is enough

        return defects

    # ── Check #9: Table gridCol Width Reversal (HIGH) ──

    def _check_table_gridcol_reversal(
        self, slide_num: int, orig_element, conv_element
    ) -> list:
        """Check if gridCol widths are reversed between original and converted."""
        from vqa_types import V3Defect, Severity

        defects = []
        orig_tables = _find_tables(orig_element)
        conv_tables = _find_tables(conv_element)

        for tbl_idx, (orig_tbl, conv_tbl) in enumerate(
            zip(orig_tables, conv_tables)
        ):
            orig_cols = _get_gridcols(orig_tbl)
            conv_cols = _get_gridcols(conv_tbl)

            if len(orig_cols) < 2 or len(conv_cols) < 2:
                continue
            if len(orig_cols) != len(conv_cols):
                continue

            orig_widths = [c.get('w', '0') for c in orig_cols]
            conv_widths = [c.get('w', '0') for c in conv_cols]

            # If all widths are equal, no need to check reversal
            if len(set(orig_widths)) <= 1:
                continue

            # After RTL conversion, gridCol widths should be reversed
            if orig_widths == conv_widths:
                # Check idempotency
                conv_tbl_pr = _get_tbl_pr(conv_tbl)
                marker = f'{{{SLIDEARABI_NS}}}physRtlCols'
                if conv_tbl_pr is not None and conv_tbl_pr.get(marker) == '1':
                    continue

                defects.append(V3Defect(
                    code="TABLE_GRIDCOL_NOT_REVERSED",
                    category="table",
                    severity=Severity.HIGH,
                    slide_idx=slide_num,
                    object_id=str(tbl_idx),
                    evidence={
                        'original_widths': orig_widths,
                        'converted_widths': conv_widths,
                    },
                    fixable=True,
                    autofix_action="reverse_table_columns",
                    description=f"Table {tbl_idx+1} gridCol widths not reversed "
                                f"(slide {slide_num})",
                ))

        return defects

    # ── Check #10: Merged Cell Integrity (HIGH, flag only) ──

    def _check_merged_cell_integrity(self, slide_num: int, conv_element) -> list:
        """Detect merged cells — flag only, blocks auto-reversal."""
        from vqa_types import V3Defect, Severity

        defects = []
        tables = _find_tables(conv_element)

        for tbl_idx, tbl in enumerate(tables):
            has_merged = False
            for row in _get_table_rows(tbl):
                for cell in _get_row_cells(row):
                    grid_span = cell.get('gridSpan', '1')
                    row_span = cell.get('rowSpan', '1')
                    try:
                        if int(grid_span) > 1 or int(row_span) > 1:
                            has_merged = True
                            break
                    except ValueError:
                        pass
                if has_merged:
                    break

            if has_merged:
                defects.append(V3Defect(
                    code="TABLE_MERGED_CELL_INTEGRITY",
                    category="table",
                    severity=Severity.HIGH,
                    slide_idx=slide_num,
                    object_id=str(tbl_idx),
                    evidence={'has_merged_cells': True},
                    fixable=False,  # Flag only — blocks reversal
                    description=f"Table {tbl_idx+1} has merged cells, "
                                f"auto-reversal blocked (slide {slide_num})",
                ))

        return defects

    # ── Check #11: Table Grid Structural Integrity (CRITICAL, flag only) ──

    def _check_table_grid_structural(self, slide_num: int, conv_element) -> list:
        """Verify gridCol count matches row cell count."""
        from vqa_types import V3Defect, Severity

        defects = []
        tables = _find_tables(conv_element)

        for tbl_idx, tbl in enumerate(tables):
            gridcols = _get_gridcols(tbl)
            num_grid_cols = len(gridcols)

            if num_grid_cols == 0:
                continue

            for row_idx, row in enumerate(_get_table_rows(tbl)):
                cells = _get_row_cells(row)
                # Account for gridSpan in cell count
                effective_cols = 0
                for cell in cells:
                    span = int(cell.get('gridSpan', '1'))
                    effective_cols += span

                if effective_cols != num_grid_cols:
                    defects.append(V3Defect(
                        code="TABLE_GRID_STRUCTURAL_ERROR",
                        category="table",
                        severity=Severity.CRITICAL,
                        slide_idx=slide_num,
                        object_id=f"{tbl_idx}:row{row_idx}",
                        evidence={
                            'grid_cols': num_grid_cols,
                            'effective_row_cols': effective_cols,
                            'row_index': row_idx,
                        },
                        fixable=False,  # Can't auto-fix structural corruption
                        description=f"Table {tbl_idx+1} row {row_idx} has "
                                    f"{effective_cols} effective cols but grid "
                                    f"has {num_grid_cols} (slide {slide_num})",
                    ))

        return defects


# ─────────────────────────────────────────────────────────────────────────────
# V3 AUTO-FIXES
# ─────────────────────────────────────────────────────────────────────────────

class V3AutoFixer:
    """V3 auto-fix engine — deterministic fixes with rollback safety."""

    MARKER_NS = SLIDEARABI_NS

    def __init__(self, slide_width: int = 9144000, slide_height: int = 6858000):
        self.slide_width = slide_width
        self.slide_height = slide_height
        self._dispatch = {
            'reverse_table_columns': self._fix_reverse_table_columns,
            'set_para_rtl': self._fix_set_para_rtl,
        }

    def apply_fix(self, slide_element, defect) -> bool:
        """Apply a single fix. Returns True if successful."""
        action = defect.autofix_action
        if action not in self._dispatch:
            logger.warning(f"Unknown fix action: {action}")
            return False
        try:
            return self._dispatch[action](slide_element, defect)
        except Exception as e:
            logger.error(f"Fix {action} failed on slide {defect.slide_idx}: {e}")
            return False

    # ── Fix #1: Reverse Table Columns ──

    def _fix_reverse_table_columns(self, slide_element, defect) -> bool:
        """Reverse <a:tc> order in each row + gridCol widths.
        
        SAFETY: Aborts if any cell has gridSpan > 1 or rowSpan > 1.
        """
        tables = _find_tables(slide_element)
        tbl_idx = int(defect.object_id.split(':')[0]) if ':' in str(defect.object_id) else int(defect.object_id)

        if tbl_idx >= len(tables):
            return False

        tbl = tables[tbl_idx]

        # Safety: check for merged cells
        for row in _get_table_rows(tbl):
            for cell in _get_row_cells(row):
                try:
                    if int(cell.get('gridSpan', '1')) > 1:
                        logger.warning(f"Merged cells found, aborting reversal on slide {defect.slide_idx}")
                        return False
                    if int(cell.get('rowSpan', '1')) > 1:
                        logger.warning(f"Row-span found, aborting reversal on slide {defect.slide_idx}")
                        return False
                except ValueError:
                    pass

        # Check idempotency marker
        tbl_pr = _get_tbl_pr(tbl)
        marker = f'{{{self.MARKER_NS}}}v3PhysRev'
        if tbl_pr is not None and tbl_pr.get(marker) == '1':
            logger.debug(f"Table already reversed (v3 marker), skipping")
            return False

        # Reverse gridCol widths
        grid = tbl.find(f'{{{A_NS}}}tblGrid')
        if grid is not None:
            cols = grid.findall(f'{{{A_NS}}}gridCol')
            if len(cols) > 1:
                widths = [c.get('w') for c in cols]
                widths.reverse()
                for col, w in zip(cols, widths):
                    if w is not None:
                        col.set('w', w)

        # Reverse cell order in each row
        for row in _get_table_rows(tbl):
            cells = _get_row_cells(row)
            if len(cells) <= 1:
                continue
            # Remove all cells, re-add in reverse
            for cell in cells:
                row.remove(cell)
            for cell in reversed(cells):
                row.append(cell)

        # Set rtl='0' to prevent double-reversal
        if tbl_pr is None:
            tbl_pr = etree.SubElement(tbl, f'{{{A_NS}}}tblPr')
        tbl_pr.set('rtl', '0')

        # Set idempotency marker
        tbl_pr.set(marker, '1')

        # Swap firstCol/lastCol toggles
        first_col = tbl_pr.get('firstCol')
        last_col = tbl_pr.get('lastCol')
        if first_col or last_col:
            tbl_pr.set('firstCol', last_col or '0')
            tbl_pr.set('lastCol', first_col or '0')

        logger.info(f"V3 Fix: reversed table columns on slide {defect.slide_idx}, "
                     f"table {tbl_idx}")
        return True

    # ── Fix #2: Set Paragraph RTL + Alignment ──

    def _fix_set_para_rtl(self, slide_element, defect) -> bool:
        """Set rtl='1' and algn='r' on Arabic paragraph in table cell."""
        tables = _find_tables(slide_element)

        # Parse object_id "tbl_idx:cell_idx"
        parts = str(defect.object_id).split(':')
        if len(parts) < 2:
            return False
        tbl_idx = int(parts[0])
        # cell_idx is informational — we fix all Arabic cells in the table

        if tbl_idx >= len(tables):
            return False

        tbl = tables[tbl_idx]
        fixed = False

        for row in _get_table_rows(tbl):
            for cell in _get_row_cells(row):
                cell_text = _get_cell_text(cell)
                if not _has_arabic(cell_text):
                    continue

                for para in cell.iter(f'{{{A_NS}}}p'):
                    pPr = para.find(f'{{{A_NS}}}pPr')
                    if pPr is None:
                        pPr = etree.SubElement(para, f'{{{A_NS}}}pPr')
                        # Move to beginning of paragraph
                        para.remove(pPr)
                        para.insert(0, pPr)

                    changed = False
                    if pPr.get('rtl') != '1':
                        pPr.set('rtl', '1')
                        changed = True
                    if pPr.get('algn') != 'r':
                        pPr.set('algn', 'r')
                        changed = True

                    if changed:
                        fixed = True

        if fixed:
            logger.info(f"V3 Fix: set RTL/alignment on slide {defect.slide_idx}, "
                         f"table {tbl_idx}")
        return fixed


# ─────────────────────────────────────────────────────────────────────────────
# V3 PIPELINE ENTRY POINTS
# ─────────────────────────────────────────────────────────────────────────────

def run_v3_xml_checks(
    orig_pptx_path: str,
    conv_pptx_path: str,
    slide_width_emu: int = 9144000,
    slide_height_emu: int = 6858000,
) -> Tuple[list, Dict[str, Any]]:
    """
    Run all V3 XML structural checks on all slides.
    
    Returns:
        (defects, metadata) where defects is List[V3Defect]
    """
    from pptx import Presentation as PptxPresentation

    checker = V3XMLChecker(slide_width_emu, slide_height_emu)
    all_defects = []

    conv_prs = PptxPresentation(conv_pptx_path)
    orig_prs = PptxPresentation(orig_pptx_path)

    for slide_idx in range(len(conv_prs.slides)):
        slide_num = slide_idx + 1
        conv_slide = conv_prs.slides[slide_idx]
        
        orig_slide_elem = None
        if slide_idx < len(orig_prs.slides):
            orig_slide_elem = orig_prs.slides[slide_idx]._element

        slide_defects = checker.check_slide(
            slide_num,
            conv_slide._element,
            orig_slide_elem,
        )
        all_defects.extend(slide_defects)

    from vqa_types import Severity
    metadata = {
        'slides_checked': len(conv_prs.slides),
        'total_defects': len(all_defects),
        'critical_count': sum(1 for d in all_defects if d.severity == Severity.CRITICAL),
        'high_count': sum(1 for d in all_defects if d.severity == Severity.HIGH),
        'fixable_count': sum(1 for d in all_defects if d.fixable),
    }

    logger.info(f"V3 XML checks: {metadata['total_defects']} defects "
                f"({metadata['critical_count']} critical, "
                f"{metadata['fixable_count']} fixable) "
                f"across {metadata['slides_checked']} slides")

    return all_defects, metadata


def safe_apply_fixes(
    pptx_path: str,
    defects: list,
    slide_width: int = 9144000,
    slide_height: int = 6858000,
) -> Tuple[list, list]:
    """
    Apply V3 fixes with per-fix validation and rollback.
    
    Returns:
        (applied, failed) — lists of V3Defect objects
    """
    from pptx import Presentation as PptxPresentation
    from vqa_types import DefectStatus
    import v3_config

    if not v3_config.ENABLE_V3_XML_AUTOFIX:
        logger.info("V3 XML autofix disabled, skipping")
        return [], list(defects)

    fixer = V3AutoFixer(slide_width, slide_height)
    applied = []
    failed = []

    # Group fixable defects by slide
    fixable = [d for d in defects if d.fixable and d.status.value == 'open']
    if not fixable:
        return [], []

    # Create backup
    backup_path = str(pptx_path) + '.v3_backup'
    shutil.copy2(pptx_path, backup_path)

    try:
        prs = PptxPresentation(pptx_path)

        # Group by slide
        by_slide: Dict[int, list] = {}
        for d in fixable:
            by_slide.setdefault(d.slide_idx, []).append(d)

        for slide_idx_0, slide in enumerate(prs.slides):
            slide_num = slide_idx_0 + 1
            slide_defects = by_slide.get(slide_num, [])
            
            for defect in slide_defects:
                # Check table autofix flag for table fixes
                if defect.autofix_action == 'reverse_table_columns':
                    if not v3_config.ENABLE_V3_TABLE_FIX:
                        logger.info(f"Table autofix disabled, skipping {defect.code}")
                        defect.status = DefectStatus.UNRESOLVED
                        failed.append(defect)
                        continue

                success = fixer.apply_fix(slide._element, defect)
                if success:
                    defect.status = DefectStatus.FIXED
                    applied.append(defect)
                else:
                    defect.status = DefectStatus.UNRESOLVED
                    failed.append(defect)

        # Save and validate
        prs.save(pptx_path)

        # Validation: try opening the saved file
        try:
            PptxPresentation(pptx_path)
            logger.info(f"V3 fixes applied: {len(applied)} succeeded, "
                        f"{len(failed)} failed")
        except Exception as e:
            logger.error(f"PPTX validation failed after fixes, rolling back: {e}")
            shutil.copy2(backup_path, pptx_path)
            # Mark all as failed
            for d in applied:
                d.status = DefectStatus.UNRESOLVED
            failed.extend(applied)
            applied = []

    except Exception as e:
        logger.error(f"Fix application failed, rolling back: {e}")
        shutil.copy2(backup_path, pptx_path)
        for d in fixable:
            d.status = DefectStatus.UNRESOLVED
        failed = list(fixable)
        applied = []

    return applied, failed


def compute_gate_decision(
    unresolved_defects: list,
    total_slides: int,
) -> 'VQAGateResult':
    """Compute quality gate decision from unresolved defects."""
    from vqa_types import (
        VQAGateResult, Severity, MUST_NOT_SHIP_CODES,
        HIGH_SEVERITY_THRESHOLD, HIGH_SLIDE_RATIO_THRESHOLD,
    )

    gate = VQAGateResult()

    critical = [d for d in unresolved_defects if d.severity == Severity.CRITICAL]
    high = [d for d in unresolved_defects if d.severity == Severity.HIGH]
    must_not_ship = [d for d in unresolved_defects if d.code in MUST_NOT_SHIP_CODES]

    gate.critical_remaining = len(critical)
    gate.high_remaining = len(high)
    gate.blocking_issues = [d.to_dict() for d in critical[:5]]
    gate.warning_issues = [d.to_dict() for d in high[:10]]

    # Gate logic (Sonnet's nuanced thresholds)
    if critical or must_not_ship:
        gate.status = "failed_qa"
    elif len(high) >= HIGH_SEVERITY_THRESHOLD:
        high_slides = len(set(d.slide_idx for d in high))
        if total_slides > 0 and high_slides / total_slides >= HIGH_SLIDE_RATIO_THRESHOLD:
            gate.status = "failed_qa"
        else:
            gate.status = "completed_with_warnings"
    elif high:
        gate.status = "completed_with_warnings"
    else:
        gate.status = "completed"

    return gate
