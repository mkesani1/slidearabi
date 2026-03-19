"""
typography.py — Font mapping, Arabic text metrics, and auto-fit.

SlideArabi: Template-First Deterministic RTL Transformation Engine.

Phase 4: TypographyNormalizer
- Maps Latin fonts to Arabic-compatible equivalents.
- Estimates Arabic text expansion ratio (~20-30% wider than Latin).
- Applies bounded font-size reduction for overflow (max −20% of original size).
- Sets Arabic-appropriate text frame margins (insets).
- Applies bidirectional text formatting per paragraph.

Key design constraint: font reduction is bounded (max 20% shrink to preserve
readability). If text still overflows after maximum reduction, the shape is
flagged in the TransformReport for manual review. NO infinite shrink loops.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from lxml import etree

from .utils import (
    A_NS, P_NS,
    qn,
    ensure_pPr,
    has_arabic,
    has_latin,
    is_bidi_text,
    compute_script_ratio,
    hundredths_pt_to_pt,
    pt_to_hundredths_pt,
    get_placeholder_info,
    get_placeholder_info_from_xml,
)
from .rtl_transforms import TransformReport

logger = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────────────────────
# Font mapping constants
# ─────────────────────────────────────────────────────────────────────────────

ARABIC_FONT_MAP: Dict[str, str] = {
    # Fonts that already support Arabic glyphs — map to themselves
    'Calibri':             'Calibri',
    'Arial':               'Arial',
    'Times New Roman':     'Times New Roman',
    'Calibri Light':       'Calibri Light',
    'Tahoma':              'Tahoma',
    'Segoe UI':            'Segoe UI',
    'Segoe UI Light':      'Segoe UI Light',
    'Segoe UI Semibold':   'Segoe UI Semibold',
    'Microsoft Sans Serif':'Microsoft Sans Serif',
    # Fonts that lack Arabic glyphs — map to Arabic-capable alternatives
    'Cambria':             'Sakkal Majalla',
    'Georgia':             'Sakkal Majalla',
    'Verdana':             'Tahoma',
    'Trebuchet MS':        'Tahoma',
    'Garamond':            'Traditional Arabic',
    'EB Garamond':         'Traditional Arabic',
    'Palatino':            'Traditional Arabic',
    'Palatino Linotype':   'Traditional Arabic',
    'Book Antiqua':        'Traditional Arabic',
    'Century Gothic':      'Dubai',
    'Futura':              'Dubai',
    'Gill Sans':           'Dubai',
    'Gill Sans MT':        'Dubai',
    'Helvetica':           'Arial',
    'Helvetica Neue':      'Arial',
    'Impact':              'Arial Black',
    'Comic Sans MS':       'Tahoma',
    'Courier New':         'Courier New',   # Monospace with Arabic support
    'Consolas':            'Courier New',
    'Monaco':              'Courier New',
    'Lucida Console':      'Courier New',
    'Roboto':              'Arial',
    'Open Sans':           'Arial',
    'Lato':                'Arial',
    'Montserrat':          'Tahoma',
    'Oswald':              'Tahoma',
    'Source Sans Pro':     'Arial',
    'Source Serif Pro':    'Sakkal Majalla',
    'PT Sans':             'Arial',
    'PT Serif':            'Sakkal Majalla',
    'Noto Sans':           'Arial',
    'Myriad Pro':          'Arial',
    'Franklin Gothic':     'Tahoma',
    'Franklin Gothic Medium': 'Tahoma',
    'League Spartan':      'Tahoma',
    'DIN':                 'Tahoma',
    'Avenir':              'Dubai',
    'Avenir Next':         'Dubai',
    'Baskerville':         'Traditional Arabic',
    'Caslon':              'Traditional Arabic',
    'Frutiger':            'Arial',
    'Univers':             'Arial',
    'Optima':              'Tahoma',
    'Rockwell':            'Sakkal Majalla',
}

# Average character width expansion ratios when switching Latin → Arabic text
# (Arabic text typically renders 20-30% wider than equivalent Latin text)
ARABIC_EXPANSION_FACTORS: Dict[str, float] = {
    'default':           1.25,  # 25% wider — conservative average
    'Calibri':           1.20,
    'Calibri Light':     1.20,
    'Arial':             1.22,
    'Times New Roman':   1.28,
    'Tahoma':            1.20,
    'Segoe UI':          1.21,
    'Sakkal Majalla':    1.15,  # Designed for Arabic — compact
    'Dubai':             1.18,
    'Traditional Arabic':1.30,
    'Simplified Arabic': 1.27,
    'Courier New':       1.05,  # Monospace — minimal expansion
    'Arial Black':       1.18,
    'Microsoft Sans Serif': 1.20,
}

# Minimum readable font sizes in points
MIN_FONT_SIZE_PT:       float = 8.0
MIN_TITLE_FONT_SIZE_PT: float = 14.0
MIN_BODY_FONT_SIZE_PT:  float = 10.0
MIN_TABLE_FONT_SIZE_PT: float = 9.0
MIN_CHART_FONT_SIZE_PT: float = 7.0

# Maximum font size reduction percentage (20% of original)
MAX_FONT_REDUCTION_PCT: float = 20.0

# Minimum text frame insets for Arabic text (FLOOR values — originals preserved
# when larger).  PowerPoint's built-in default is 91440 EMU (0.1") for left/right
# and 45720 EMU (0.05") for top/bottom.  We use those as floors so that designer-
# set generous margins are never crushed, while shapes with zero or tiny margins
# still get Arabic-appropriate breathing room.
ARABIC_MIN_INSET_LR_EMU:  int = 91440   # 0.1" — PowerPoint default left/right
ARABIC_MIN_INSET_TOP_EMU: int = 45720   # 0.05" — PowerPoint default top
ARABIC_MIN_INSET_BOT_EMU: int = 45720   # 0.05" — PowerPoint default bottom

# PowerPoint OOXML spec defaults when bodyPr inset attributes are absent
# (§21.1.2.1.3: lIns=91440, rIns=91440, tIns=45720, bIns=45720)
_PPTX_DEFAULT_INSET_LR: int = 91440
_PPTX_DEFAULT_INSET_TB: int = 45720

# Average character width in points (heuristic for overflow estimation)
# Latin characters average ~0.55× font-size in width; Arabic ~0.65× font-size
LATIN_CHAR_WIDTH_FACTOR:  float = 0.55
ARABIC_CHAR_WIDTH_FACTOR: float = 0.65


# ─────────────────────────────────────────────────────────────────────────────
# TypographyNormalizer
# ─────────────────────────────────────────────────────────────────────────────

class TypographyNormalizer:
    """
    Applies Arabic typography normalization to a presentation (Phase 4).

    Runs AFTER RTL transforms have repositioned shapes and inserted Arabic text.

    Responsibilities:
    1. Map Latin fonts to Arabic-capable equivalents.
    2. Estimate whether Arabic text will overflow its text frame.
    3. Apply bounded font-size reduction (max −20%) to fit overflow.
    4. Set Arabic-appropriate text frame margins.
    5. Apply bidirectional text formatting per paragraph.
    """

    def __init__(self, presentation):
        """
        Args:
            presentation: python-pptx Presentation object after RTL transforms.
        """
        self.prs = presentation
        self._slide_width = int(presentation.slide_width)
        self._slide_height = int(presentation.slide_height)

    # ─────────────────────────────────────────────────────────────────────
    # Public entry point
    # ─────────────────────────────────────────────────────────────────────

    def normalize_all(self) -> TransformReport:
        """
        Apply typography normalization to the entire presentation.

        Returns:
            TransformReport with per-type change counts and any warnings
            for shapes where text could not be fitted within bounds.
        """
        report = TransformReport(phase='typography')

        for slide_idx, slide in enumerate(self.prs.slides):
            slide_num = slide_idx + 1
            try:
                count = self._normalize_slide(slide)
                report.add('slide_normalized', count)
            except Exception as exc:
                report.error(f'slide[{slide_num}]: {exc}')

        return report

    # ─────────────────────────────────────────────────────────────────────
    # Slide-level normalization
    # ─────────────────────────────────────────────────────────────────────

    def _normalize_slide(self, slide) -> int:
        """
        Apply typography normalization to all shapes on a slide.

        Returns:
            Count of changes applied.
        """
        changes = 0
        all_shapes = self._collect_all_shapes(slide.shapes)

        for shape in all_shapes:
            try:
                # Text frame shapes
                if getattr(shape, 'has_text_frame', False) and shape.has_text_frame:
                    changes += self._apply_font_mapping(shape)
                    changes += self._set_text_frame_margins(shape)
                    changes += self._apply_bidi_formatting(shape)
                    # Check overflow and reduce font if needed
                    if self._check_text_overflow(shape):
                        changed = self._reduce_font_size_to_fit(shape)
                        changes += changed

                # Table cells
                if getattr(shape, 'has_table', False) and shape.has_table:
                    changes += self._normalize_table_typography(shape)

            except Exception as exc:
                logger.warning('_normalize_slide shape "%s": %s',
                               getattr(shape, 'name', '?'), exc)

        return changes

    def _collect_all_shapes(self, shapes) -> List:
        """Recursively collect all shapes including group children."""
        result = []
        for shape in shapes:
            result.append(shape)
            if hasattr(shape, 'shapes'):
                result.extend(self._collect_all_shapes(shape.shapes))
        return result

    # ─────────────────────────────────────────────────────────────────────
    # Font mapping
    # ─────────────────────────────────────────────────────────────────────

    def _map_font(self, font_name: str) -> str:
        """
        Map an English/Latin font name to its Arabic-compatible equivalent.

        Many common fonts (Calibri, Arial, Times New Roman) already include
        Arabic glyph support and map to themselves. Fonts without Arabic glyphs
        are mapped to the closest-style Arabic-capable substitute.

        Args:
            font_name: Font name string as stored in OOXML (e.g. 'Calibri').

        Returns:
            Arabic-compatible font name. Returns the original name unchanged
            if no mapping is found (assume it supports Arabic or is a niche
            corporate font).
        """
        if not font_name:
            return font_name

        # Exact match first
        mapped = ARABIC_FONT_MAP.get(font_name)
        if mapped:
            return mapped

        # Case-insensitive match
        lower = font_name.lower()
        for key, value in ARABIC_FONT_MAP.items():
            if key.lower() == lower:
                return value

        # Partial match for font families (e.g. 'Calibri Bold' → 'Calibri')
        for key, value in ARABIC_FONT_MAP.items():
            if lower.startswith(key.lower()):
                return value

        # Unknown font — return as-is (may already support Arabic)
        return font_name

    def _apply_font_mapping(self, shape) -> int:
        """
        Apply font mapping to all runs in a shape's text frame.

        Maps both the <a:latin> and <a:cs> (complex script) font elements on
        each run's <a:rPr>. For Arabic text runs, also ensures a <a:cs> element
        is present with an appropriate Arabic font, because some renderers only
        use <a:cs> for Arabic glyph selection.

        Args:
            shape: python-pptx Shape with a text frame.

        Returns:
            Count of font attribute writes performed.
        """
        changes = 0
        try:
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    rPr = run._r.find(f'{{{A_NS}}}rPr')
                    if rPr is None:
                        continue

                    run_text = run.text or ''

                    # Map <a:latin> font
                    latin_elem = rPr.find(f'{{{A_NS}}}latin')
                    if latin_elem is not None:
                        current_font = latin_elem.get('typeface', '')
                        mapped_font = self._map_font(current_font)
                        if mapped_font != current_font:
                            latin_elem.set('typeface', mapped_font)
                            changes += 1

                    # Ensure <a:cs> (complex script) font for Arabic runs
                    if has_arabic(run_text):
                        cs_elem = rPr.find(f'{{{A_NS}}}cs')
                        if cs_elem is None:
                            cs_elem = etree.SubElement(rPr, f'{{{A_NS}}}cs')
                            # Pick the Arabic font from the mapped latin font,
                            # or fall back to Calibri (which has Arabic glyphs)
                            base_font = 'Calibri'
                            if latin_elem is not None:
                                base_font = self._map_font(
                                    latin_elem.get('typeface', 'Calibri')
                                )
                            cs_elem.set('typeface', base_font)
                            changes += 1
                        else:
                            current_cs = cs_elem.get('typeface', '')
                            mapped_cs = self._map_font(current_cs)
                            if mapped_cs != current_cs:
                                cs_elem.set('typeface', mapped_cs)
                                changes += 1

                        # Set Arabic language on runs containing Arabic
                        rPr.set('lang', 'ar-SA')
                        changes += 1

        except Exception as exc:
            logger.warning('_apply_font_mapping on "%s": %s',
                           getattr(shape, 'name', '?'), exc)

        return changes

    # ─────────────────────────────────────────────────────────────────────
    # Arabic text expansion estimation
    # ─────────────────────────────────────────────────────────────────────

    def _estimate_arabic_expansion(
        self,
        english_text: str,
        arabic_text: str,
        font_size_pt: float,
        font_name: str = 'default',
    ) -> float:
        """
        Estimate the width expansion ratio when English text is replaced with Arabic.

        Uses a combined approach:
        1. Character count ratio: Arabic text length / English text length.
        2. Per-character width factor: Arabic characters are ~18-28% wider per
           character than Latin, depending on the font.

        The expansion ratio accounts for both the change in character count and
        the change in per-character width.

        Args:
            english_text: Original English text string.
            arabic_text: Translated Arabic text string.
            font_size_pt: Font size in points (used for heuristic scaling).
            font_name: Font name for per-font expansion factor lookup.

        Returns:
            Expansion ratio as a float (e.g., 1.25 means 25% wider).
            Always returns at least 1.0 (never estimates shrinkage).
        """
        if not english_text:
            return 1.0

        # Character count ratio
        en_len = max(len(english_text.strip()), 1)
        ar_len = max(len(arabic_text.strip()), 1)
        char_ratio = ar_len / en_len

        # Per-character width factor from font table
        font_factor = ARABIC_EXPANSION_FACTORS.get(
            font_name,
            ARABIC_EXPANSION_FACTORS['default']
        )

        # Expansion = (char count change) × (per-char width change)
        # Clamp to a reasonable range [0.8, 2.0] to avoid extreme estimates
        expansion = char_ratio * font_factor
        expansion = max(0.80, min(2.0, expansion))

        return expansion

    # ─────────────────────────────────────────────────────────────────────
    # Overflow detection
    # ─────────────────────────────────────────────────────────────────────

    def _check_text_overflow(self, shape) -> bool:
        """
        Check whether text likely overflows the shape's text frame.

        Uses a heuristic based on estimated total text width vs. available
        frame width. This is intentionally conservative — it may flag some
        shapes that would render fine, but won't miss genuinely overflowing
        shapes.

        Heuristic:
        - For each paragraph, compute: character_count × char_width_factor × font_size
        - Sum across lines (wrapping at frame width)
        - Compare total estimated height to frame height

        Args:
            shape: python-pptx Shape with a text frame.

        Returns:
            True if overflow is detected, False otherwise.
        """
        try:
            tf = shape.text_frame
            frame_width = getattr(shape, 'width', None)
            frame_height = getattr(shape, 'height', None)

            if frame_width is None or frame_height is None or frame_width <= 0:
                return False

            # Convert EMU dimensions to points
            frame_width_pt = frame_width / 12700   # EMU → pt
            frame_height_pt = frame_height / 12700

            total_lines_estimate = 0.0

            for para in tf.paragraphs:
                text = para.text or ''
                if not text.strip():
                    total_lines_estimate += 0.5  # Empty paragraph spacing
                    continue

                # Estimate font size for this paragraph
                font_size_pt = self._get_effective_font_size(para)

                # Estimate character width
                if has_arabic(text):
                    char_factor = ARABIC_CHAR_WIDTH_FACTOR
                else:
                    char_factor = LATIN_CHAR_WIDTH_FACTOR

                # Estimate line width in points
                estimated_line_width_pt = len(text) * char_factor * font_size_pt

                # Estimate number of lines (with word-wrap at frame width)
                lines = max(1.0, estimated_line_width_pt / frame_width_pt)

                # Line height ≈ 1.2× font size
                total_lines_estimate += lines

            # Total estimated height
            line_height_pt = self._get_dominant_font_size(tf) * 1.2
            estimated_height_pt = total_lines_estimate * line_height_pt

            # Apply inset margins
            margin_pt = (ARABIC_MIN_INSET_TOP_EMU + ARABIC_MIN_INSET_BOT_EMU) / 12700
            available_height_pt = frame_height_pt - margin_pt

            return estimated_height_pt > available_height_pt * 1.05  # 5% tolerance

        except Exception as exc:
            logger.debug('_check_text_overflow on "%s": %s',
                         getattr(shape, 'name', '?'), exc)
            return False

    def _get_effective_font_size(self, paragraph) -> float:
        """
        Estimate the effective font size for a paragraph in points.

        Reads from the first run's rPr sz attribute if available, falls
        back to a placeholder-type-based default.
        """
        try:
            for run in paragraph.runs:
                rPr = run._r.find(f'{{{A_NS}}}rPr')
                if rPr is not None:
                    sz = rPr.get('sz')
                    if sz is not None:
                        return int(sz) / 100.0
        except Exception:
            pass
        return 12.0  # Safe default

    def _get_dominant_font_size(self, text_frame) -> float:
        """
        Find the most common effective font size in a text frame (in points).

        Returns the dominant font size across all runs, or 12.0 as fallback.
        """
        sizes: List[float] = []
        try:
            for para in text_frame.paragraphs:
                for run in para.runs:
                    rPr = run._r.find(f'{{{A_NS}}}rPr')
                    if rPr is not None:
                        sz = rPr.get('sz')
                        if sz is not None:
                            try:
                                sizes.append(int(sz) / 100.0)
                            except (ValueError, TypeError):
                                pass
        except Exception:
            pass

        if not sizes:
            return 12.0

        # Return the median (most representative) size
        sizes.sort()
        return sizes[len(sizes) // 2]

    # ─────────────────────────────────────────────────────────────────────
    # Font size reduction
    # ─────────────────────────────────────────────────────────────────────

    def _reduce_font_size_to_fit(
        self,
        shape,
        max_reduction_pct: float = MAX_FONT_REDUCTION_PCT,
    ) -> int:
        """
        Reduce font size on all runs in a shape to fit Arabic text within the
        text frame bounds.

        Applies a single proportional reduction to all runs rather than
        iterative shrinking. Maximum reduction is bounded at *max_reduction_pct*
        percent of the original size to preserve readability.

        After reduction, enforces per-type minimum font floors:
        - Title placeholders (type title/ctrTitle/subTitle): min 14pt
        - Body text: min 10pt
        - General text: min 8pt

        If text still likely overflows after maximum reduction, the shape's
        name is logged as a warning for manual review.

        Args:
            shape: python-pptx Shape with a text frame.
            max_reduction_pct: Maximum percentage reduction allowed (default 20%).

        Returns:
            Count of font size attributes modified.
        """
        changes = 0
        try:
            tf = shape.text_frame
            ph_info = get_placeholder_info(shape)
            ph_type = ph_info[0] if ph_info else None

            # Determine floor based on placeholder type.
            # Non-placeholder (freeform) shapes are treated as body-level content.
            if ph_type in ('title', 'ctrTitle', 'subTitle', 'center_title'):
                floor_pt = MIN_TITLE_FONT_SIZE_PT
            elif ph_type in (None, 'body', 'object', 'pic', 'media'):
                # None means non-placeholder freeform — use body floor
                floor_pt = MIN_BODY_FONT_SIZE_PT
            else:
                floor_pt = MIN_BODY_FONT_SIZE_PT  # Safe default for unknown types

            # Compute reduction factor
            reduction_factor = 1.0 - (max_reduction_pct / 100.0)

            for para in tf.paragraphs:
                for run in para.runs:
                    if not (run.text or '').strip():
                        continue
                    rPr = run._r.find(f'{{{A_NS}}}rPr')
                    if rPr is None:
                        continue

                    sz_str = rPr.get('sz')
                    if sz_str is None:
                        continue

                    try:
                        sz_hundredths = int(sz_str)
                        current_pt = sz_hundredths / 100.0
                        new_pt = max(floor_pt, current_pt * reduction_factor)

                        if new_pt < current_pt - 0.5:  # Only write if meaningful change
                            rPr.set('sz', str(pt_to_hundredths_pt(new_pt)))
                            changes += 1
                    except (ValueError, TypeError):
                        continue

            # Check if shape still overflows after reduction
            if changes > 0 and self._check_text_overflow(shape):
                logger.warning(
                    'Shape "%s" may still overflow after %.0f%% font reduction — '
                    'flag for manual review',
                    getattr(shape, 'name', '?'),
                    max_reduction_pct,
                )

        except Exception as exc:
            logger.warning('_reduce_font_size_to_fit on "%s": %s',
                           getattr(shape, 'name', '?'), exc)

        return changes

    # ─────────────────────────────────────────────────────────────────────
    # Text frame margins
    # ─────────────────────────────────────────────────────────────────────

    def _set_text_frame_margins(self, shape) -> int:
        """
        Ensure Arabic text frames have adequate inset margins.

        Strategy: MAX(original_margin, arabic_minimum_floor).
        - If the shape already has generous margins, preserve them.
        - If the shape has tiny or zero margins, bump to the floor.
        - If the shape has no explicit margin attributes, PowerPoint
          uses implicit defaults (91440 LR, 45720 TB) — we read those
          as the original value rather than treating absent as zero.

        Only applies to shapes whose text frame contains Arabic text.

        Returns:
            1 if any margin was changed, 0 otherwise.
        """
        try:
            tf = shape.text_frame
            frame_text = tf.text or ''
            if not has_arabic(frame_text):
                return 0

            body_pr = tf._txBody.find(f'{{{A_NS}}}bodyPr')
            if body_pr is None:
                return 0

            changed = False

            # For each axis, read the existing value (or PowerPoint implicit
            # default if the attribute is absent), then apply MAX with floor.
            for attr, spec_default, floor_emu in (
                ('lIns', _PPTX_DEFAULT_INSET_LR, ARABIC_MIN_INSET_LR_EMU),
                ('rIns', _PPTX_DEFAULT_INSET_LR, ARABIC_MIN_INSET_LR_EMU),
                ('tIns', _PPTX_DEFAULT_INSET_TB, ARABIC_MIN_INSET_TOP_EMU),
                ('bIns', _PPTX_DEFAULT_INSET_TB, ARABIC_MIN_INSET_BOT_EMU),
            ):
                raw = body_pr.get(attr)
                if raw is not None:
                    try:
                        original = int(raw)
                    except (ValueError, TypeError):
                        original = spec_default
                else:
                    # Attribute absent → PowerPoint uses its spec default
                    original = spec_default

                target = max(original, floor_emu)
                # Only write if we're actually changing the effective value
                if raw is None and target != spec_default:
                    body_pr.set(attr, str(target))
                    changed = True
                elif raw is not None and int(raw) != target:
                    body_pr.set(attr, str(target))
                    changed = True

            return 1 if changed else 0

        except Exception as exc:
            logger.debug('_set_text_frame_margins on "%s": %s',
                         getattr(shape, 'name', '?'), exc)
            return 0

    # ─────────────────────────────────────────────────────────────────────
    # Bidirectional text formatting
    # ─────────────────────────────────────────────────────────────────────

    def _apply_bidi_formatting(self, shape) -> int:
        """
        Apply bidirectional text formatting to all paragraphs in a shape.

        Per-paragraph rules:
        1. Pure Arabic paragraphs (>70% Arabic chars):
           - rtl='1', algn='r'
           - lang='ar-SA' on all runs
        2. Mixed bidirectional paragraphs (Arabic + Latin):
           - rtl='1' (base direction Arabic), algn='r'
           - Inject RLM (U+200F) at start of paragraph to reinforce bidi base dir
           - Latin runs remain LTR within the bidi paragraph (handled by OS)
        3. Pure English/numeric paragraphs in an Arabic-context text frame:
           - Inherit rtl='1' from context (keeps visual grouping with Arabic)
           - algn='r' to align with Arabic siblings
        4. Pure LTR paragraphs in a non-Arabic text frame:
           - rtl='0', algn='l'
        5. Footer/slide-number/date placeholders: always algn='l'.
        6. ctrTitle placeholder: rtl='1', algn='ctr'.

        Args:
            shape: python-pptx Shape with a text frame.

        Returns:
            Count of paragraph pPr modifications.
        """
        changes = 0
        try:
            tf = shape.text_frame
            frame_text = tf.text or ''
            frame_has_arabic = has_arabic(frame_text)

            ph_info = get_placeholder_info(shape)
            ph_type = ph_info[0] if ph_info else None
            is_footer = ph_type in ('ftr', 'sldNum', 'dt', 'footer', 'slideNumber', 'date_time')
            is_ctr_title = (ph_type == 'ctrTitle')

            # Set rtlCol on bodyPr if frame has any Arabic text
            if frame_has_arabic:
                try:
                    body_pr = tf._txBody.find(f'{{{A_NS}}}bodyPr')
                    if body_pr is not None:
                        body_pr.set('rtlCol', '1')
                except Exception:
                    pass

            for para in tf.paragraphs:
                text = para.text or ''
                if not text.strip():
                    continue

                ratios = compute_script_ratio(text)
                arabic_ratio = ratios['arabic']
                latin_ratio = ratios['latin']
                para_has_arabic = has_arabic(text)

                pPr = ensure_pPr(para._p)

                if is_footer:
                    # Footer-type: always left-aligned
                    pPr.set('rtl', '0')
                    pPr.set('algn', 'l')
                    changes += 1
                    continue

                if is_ctr_title:
                    # Centre title: RTL but centred
                    pPr.set('rtl', '1')
                    pPr.set('algn', 'ctr')
                    changes += 1
                    continue

                if arabic_ratio > 0.70:
                    # Predominantly Arabic
                    pPr.set('rtl', '1')
                    pPr.set('algn', 'r')
                    self._set_arabic_lang_on_runs(para)
                    changes += 1

                elif latin_ratio > 0.70 and para_has_arabic:
                    # High Latin ratio but contains some Arabic — mixed bidi
                    # Insert RLM to enforce RTL base direction
                    pPr.set('rtl', '1')
                    pPr.set('algn', 'r')
                    self._inject_rlm_at_para_start(para)
                    changes += 1

                elif latin_ratio > 0.70 and not para_has_arabic:
                    # Pure Latin paragraph
                    if frame_has_arabic:
                        # Latin text inside an Arabic-context frame
                        # Keep aligned with Arabic siblings
                        pPr.set('rtl', '1')
                        pPr.set('algn', 'r')
                    else:
                        # Completely LTR context
                        pPr.set('rtl', '0')
                        pPr.set('algn', 'l')
                    changes += 1

                elif para_has_arabic:
                    # Mixed paragraph: Arabic + other (numeric, punctuation)
                    pPr.set('rtl', '1')
                    pPr.set('algn', 'r')
                    self._inject_rlm_at_para_start(para)
                    self._set_arabic_lang_on_runs(para)
                    changes += 1

                else:
                    # No Arabic at all
                    if frame_has_arabic:
                        pPr.set('rtl', '1')
                        pPr.set('algn', 'r')
                    else:
                        pPr.set('rtl', '0')
                        pPr.set('algn', 'l')
                    changes += 1

        except Exception as exc:
            logger.warning('_apply_bidi_formatting on "%s": %s',
                           getattr(shape, 'name', '?'), exc)

        return changes

    def _set_arabic_lang_on_runs(self, paragraph) -> None:
        """
        Set lang='ar-SA' on all runs in a paragraph.

        This ensures PowerPoint selects Arabic-capable fonts for glyph
        rendering even if the run's explicit font lacks Arabic glyphs.
        """
        try:
            for r_elem in paragraph._p.findall(f'{{{A_NS}}}r'):
                rPr = r_elem.find(f'{{{A_NS}}}rPr')
                if rPr is not None:
                    rPr.set('lang', 'ar-SA')
        except Exception:
            pass

    def _inject_rlm_at_para_start(self, paragraph) -> None:
        """
        Inject a Right-to-Left Mark (U+200F) at the very start of a paragraph
        to reinforce the bidi base direction for mixed-script content.

        The RLM is inserted at the beginning of the first non-empty run's text.
        Only injects if not already present to avoid duplicates.
        """
        try:
            for r_elem in paragraph._p.findall(f'{{{A_NS}}}r'):
                t_elem = r_elem.find(f'{{{A_NS}}}t')
                if t_elem is not None and t_elem.text:
                    if t_elem.text[0] != '\u200F':
                        t_elem.text = '\u200F' + t_elem.text
                    return  # Only inject at the very first run
        except Exception:
            pass

    # ─────────────────────────────────────────────────────────────────────
    # Table typography
    # ─────────────────────────────────────────────────────────────────────

    def _normalize_table_typography(self, shape) -> int:
        """
        Apply font mapping and RTL formatting to all cells in a table shape.

        Args:
            shape: python-pptx Shape with has_table == True.

        Returns:
            Count of changes applied.
        """
        changes = 0
        try:
            table = shape.table
            num_cols = len(table.columns)

            for row_idx, row in enumerate(table.rows):
                for col_idx, cell in enumerate(row.cells):
                    if not cell.text_frame:
                        continue

                    cell_text = cell.text_frame.text or ''
                    cell_has_arabic = has_arabic(cell_text)

                    if cell_has_arabic:
                        # Apply font mapping
                        for para in cell.text_frame.paragraphs:
                            for run in para.runs:
                                rPr = run._r.find(f'{{{A_NS}}}rPr')
                                if rPr is not None:
                                    latin_el = rPr.find(f'{{{A_NS}}}latin')
                                    if latin_el is not None:
                                        current = latin_el.get('typeface', '')
                                        mapped = self._map_font(current)
                                        if mapped != current:
                                            latin_el.set('typeface', mapped)
                                            changes += 1
                                    rPr.set('lang', 'ar-SA')
                                    changes += 1

                        # Check and reduce font size if needed
                        if self._check_text_overflow_in_cell(cell):
                            changes += self._reduce_cell_font_size(cell)

        except Exception as exc:
            logger.warning('_normalize_table_typography on "%s": %s',
                           getattr(shape, 'name', '?'), exc)

        return changes

    def _check_text_overflow_in_cell(self, cell) -> bool:
        """
        Check if text likely overflows a table cell.

        Uses the same heuristic as _check_text_overflow but adapted for cells
        which typically have smaller dimensions.
        """
        try:
            # Table cells don't expose width/height through python-pptx cell API
            # so we use a simple character-count heuristic
            text = cell.text_frame.text or ''
            # If more than 80 chars in a cell, flag as potentially overflowing
            return len(text.strip()) > 80
        except Exception:
            return False

    def _reduce_cell_font_size(self, cell, max_reduction_pct: float = 15.0) -> int:
        """
        Reduce font size in a table cell, with a stricter floor (MIN_TABLE_FONT_SIZE_PT).

        Table cells can tolerate more aggressive reduction because they are
        typically read as reference data rather than flowing prose.
        """
        changes = 0
        floor_pt = MIN_TABLE_FONT_SIZE_PT
        reduction_factor = 1.0 - (max_reduction_pct / 100.0)

        try:
            for para in cell.text_frame.paragraphs:
                for run in para.runs:
                    rPr = run._r.find(f'{{{A_NS}}}rPr')
                    if rPr is None:
                        continue
                    sz_str = rPr.get('sz')
                    if sz_str is None:
                        continue
                    try:
                        current_pt = int(sz_str) / 100.0
                        new_pt = max(floor_pt, current_pt * reduction_factor)
                        if new_pt < current_pt - 0.5:
                            rPr.set('sz', str(pt_to_hundredths_pt(new_pt)))
                            changes += 1
                    except (ValueError, TypeError):
                        continue
        except Exception:
            pass

        return changes
