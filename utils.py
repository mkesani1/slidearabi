"""
utils.py — OOXML namespace helpers, coordinate math, text direction utilities.

SlideArabi: Template-First Deterministic RTL Transformation Engine.
"""

from __future__ import annotations

import re
from typing import Dict, Optional, Tuple
from lxml import etree


# ─────────────────────────────────────────────────────────────────────────────
# OOXML Namespace Map
# ─────────────────────────────────────────────────────────────────────────────

NSMAP: Dict[str, str] = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
}

# Shorthand namespace strings (for f-string use)
A_NS = NSMAP['a']
R_NS = NSMAP['r']
P_NS = NSMAP['p']
C_NS = NSMAP['c']

# EMU conversion constants
EMU_PER_INCH: int = 914400
EMU_PER_PT: int = 12700

# Hundredths-of-a-point factor (OOXML stores font sizes as sz * 100)
HUNDREDTHS_PER_PT: int = 100

# Arabic Unicode ranges for character detection
_ARABIC_RANGES = [
    (0x0600, 0x06FF),   # Arabic
    (0x0750, 0x077F),   # Arabic Supplement
    (0xFB50, 0xFDFF),   # Arabic Presentation Forms-A
    (0xFE70, 0xFEFF),   # Arabic Presentation Forms-B
]


# ─────────────────────────────────────────────────────────────────────────────
# Unit Conversion Functions
# ─────────────────────────────────────────────────────────────────────────────

def emu_to_inches(emu: int) -> float:
    """Convert EMU (English Metric Units) to inches."""
    return emu / EMU_PER_INCH


def emu_to_pt(emu: int) -> float:
    """Convert EMU to points."""
    return emu / EMU_PER_PT


def pt_to_emu(pt: float) -> int:
    """Convert points to EMU."""
    return int(pt * EMU_PER_PT)


def inches_to_emu(inches: float) -> int:
    """Convert inches to EMU."""
    return int(inches * EMU_PER_INCH)


def hundredths_pt_to_pt(val: int) -> float:
    """
    Convert OOXML font size (hundredths of a point) to points.
    OOXML stores font sizes as sz in hundredths of a point, e.g. sz="1800" = 18pt.
    """
    return val / HUNDREDTHS_PER_PT


def pt_to_hundredths_pt(pt: float) -> int:
    """
    Convert points to OOXML font size format (hundredths of a point).
    e.g. 18.0 pt → 1800.
    """
    return int(round(pt * HUNDREDTHS_PER_PT))


# ─────────────────────────────────────────────────────────────────────────────
# Coordinate Math
# ─────────────────────────────────────────────────────────────────────────────

def mirror_x(x_emu: int, width_emu: int, slide_width_emu: int) -> int:
    """
    Compute the mirrored X position for RTL layout.

    For a shape at position x with width w on a slide of total width S:
        new_x = S - (x + w)

    This places the shape's right edge where its left edge was, mirroring
    it horizontally about the centre of the slide.

    Args:
        x_emu: Current left position of the shape in EMU.
        width_emu: Width of the shape in EMU.
        slide_width_emu: Total slide width in EMU.

    Returns:
        New left position (x) in EMU after mirroring.
    """
    return slide_width_emu - (x_emu + width_emu)


def swap_positions(
    shape1_x: int,
    shape1_w: int,
    shape2_x: int,
    shape2_w: int,
    slide_width: int,
) -> Tuple[int, int]:
    """
    Compute swapped X positions for two shapes (used for column swapping).

    Preserves each shape's width but swaps horizontal positions so that
    shape1 takes the mirrored position of shape2 and vice versa.

    This is the correct RTL operation for two-column layouts: the left
    column moves to where the right column was (mirrored), and vice versa.

    Args:
        shape1_x: Left position of shape 1 in EMU.
        shape1_w: Width of shape 1 in EMU.
        shape2_x: Left position of shape 2 in EMU.
        shape2_w: Width of shape 2 in EMU.
        slide_width: Total slide width in EMU.

    Returns:
        Tuple of (new_x_for_shape1, new_x_for_shape2) in EMU.
    """
    new_x1 = mirror_x(shape2_x, shape2_w, slide_width)
    new_x2 = mirror_x(shape1_x, shape1_w, slide_width)
    return new_x1, new_x2


# ─────────────────────────────────────────────────────────────────────────────
# Text Script Detection
# ─────────────────────────────────────────────────────────────────────────────

def _is_arabic_char(ch: str) -> bool:
    """Return True if the character is in an Arabic Unicode range."""
    cp = ord(ch)
    return any(lo <= cp <= hi for lo, hi in _ARABIC_RANGES)


def _is_latin_char(ch: str) -> bool:
    """Return True if the character is a Latin letter (A-Z, a-z)."""
    return ('A' <= ch <= 'Z') or ('a' <= ch <= 'z')


def has_arabic(text: str) -> bool:
    """
    Check if *text* contains any Arabic characters.

    Covers Unicode ranges:
    - U+0600–U+06FF  Arabic block
    - U+0750–U+077F  Arabic Supplement
    - U+FB50–U+FDFF  Arabic Presentation Forms-A
    - U+FE70–U+FEFF  Arabic Presentation Forms-B

    Args:
        text: String to test.

    Returns:
        True if any character falls in an Arabic range.
    """
    return any(_is_arabic_char(ch) for ch in text)


def has_latin(text: str) -> bool:
    """
    Check if *text* contains any Latin characters (A–Z or a–z).

    Args:
        text: String to test.

    Returns:
        True if any character is a Latin letter.
    """
    return any(_is_latin_char(ch) for ch in text)


def is_bidi_text(text: str) -> bool:
    """
    Check if *text* is bidirectional (contains both Arabic and Latin characters).

    Args:
        text: String to test.

    Returns:
        True if text contains at least one Arabic character AND at least one
        Latin character, indicating bidirectional content.
    """
    return has_arabic(text) and has_latin(text)


def compute_script_ratio(text: str) -> Dict[str, float]:
    """
    Compute the ratio of Arabic, Latin, numeric, and other characters in *text*.

    Ignores whitespace when computing total character count so that
    sparse strings don't skew ratios.

    Args:
        text: Input string.

    Returns:
        Dict with keys 'arabic', 'latin', 'numeric', 'other', each
        a float in [0.0, 1.0] summing to 1.0.  Returns all-zero dict
        for empty/whitespace-only strings.
    """
    counts = {'arabic': 0, 'latin': 0, 'numeric': 0, 'other': 0}
    total = 0

    for ch in text:
        if ch.isspace():
            continue
        total += 1
        if _is_arabic_char(ch):
            counts['arabic'] += 1
        elif _is_latin_char(ch):
            counts['latin'] += 1
        elif ch.isdigit():
            counts['numeric'] += 1
        else:
            counts['other'] += 1

    if total == 0:
        return {k: 0.0 for k in counts}

    return {k: v / total for k, v in counts.items()}


# ─────────────────────────────────────────────────────────────────────────────
# OOXML XML Helpers
# ─────────────────────────────────────────────────────────────────────────────

def qn(tag: str) -> str:
    """
    Build a Clark-notation qualified name from a prefixed tag.

    Supports 'a:', 'p:', 'r:', 'c:' prefixes defined in NSMAP.

    Args:
        tag: Prefixed element name e.g. 'a:pPr', 'p:sp'.

    Returns:
        Clark notation string e.g. '{http://...}pPr'.

    Raises:
        KeyError: If the prefix is not in NSMAP.
    """
    prefix, local = tag.split(':', 1)
    return f'{{{NSMAP[prefix]}}}{local}'


def ensure_pPr(paragraph_element) -> etree._Element:
    """
    Ensure that an ``<a:pPr>`` element exists as the first child of *paragraph_element*.

    If the element already exists it is returned unchanged.  If it is missing it
    is created and inserted at position 0 so it precedes any ``<a:r>`` run children,
    as required by the DrawingML schema.

    Args:
        paragraph_element: An ``<a:p>`` lxml Element (or a python-pptx
            ``_Paragraph._p`` attribute).

    Returns:
        The existing or newly created ``<a:pPr>`` element.
    """
    pPr_tag = qn('a:pPr')
    pPr = paragraph_element.find(pPr_tag)
    if pPr is None:
        pPr = etree.Element(pPr_tag)
        paragraph_element.insert(0, pPr)
    return pPr


def set_rtl_on_paragraph(paragraph_element) -> None:
    """
    Set ``rtl='1'`` on the ``<a:pPr>`` of *paragraph_element*.

    Creates ``<a:pPr>`` if missing.

    Args:
        paragraph_element: An ``<a:p>`` lxml Element.
    """
    pPr = ensure_pPr(paragraph_element)
    pPr.set('rtl', '1')


def set_alignment_on_paragraph(paragraph_element, alignment: str) -> None:
    """
    Set the ``algn`` attribute on the ``<a:pPr>`` of *paragraph_element*.

    Creates ``<a:pPr>`` if missing.

    Args:
        paragraph_element: An ``<a:p>`` lxml Element.
        alignment: OOXML alignment value — one of ``'l'``, ``'r'``, ``'ctr'``,
            ``'just'``, ``'dist'``, ``'thaiDist'``.
    """
    pPr = ensure_pPr(paragraph_element)
    pPr.set('algn', alignment)


def get_placeholder_info(shape) -> Optional[Tuple[str, int]]:
    """
    Extract placeholder type and index from a python-pptx shape.

    Args:
        shape: A python-pptx Shape object.

    Returns:
        ``(placeholder_type, idx)`` tuple if the shape is a placeholder,
        e.g. ``('title', 0)`` or ``('body', 1)``.
        Returns ``None`` if the shape is not a placeholder or cannot be
        parsed.

    Notes:
        - ``placeholder_type`` defaults to ``'body'`` when the ``type``
          attribute is absent (OOXML convention for body placeholders).
        - ``idx`` defaults to ``0`` when the ``idx`` attribute is absent.
    """
    try:
        if not getattr(shape, 'is_placeholder', False):
            return None
        ph_fmt = shape.placeholder_format
        if ph_fmt is None:
            return None
        # PP_PLACEHOLDER enum → string name
        ph_type = str(ph_fmt.type).split('.')[-1].lower()
        ph_idx = ph_fmt.idx if ph_fmt.idx is not None else 0
        return ph_type, ph_idx
    except Exception:
        return None


def get_placeholder_info_from_xml(shape_element) -> Optional[Tuple[str, int]]:
    """
    Extract placeholder type and index directly from raw lxml shape element XML.

    Useful when iterating over XML elements without a python-pptx wrapper.

    Args:
        shape_element: An lxml Element for a ``<p:sp>`` shape.

    Returns:
        ``(placeholder_type, idx)`` or ``None``.
    """
    try:
        nv_sp_pr = shape_element.find(qn('p:nvSpPr'))
        if nv_sp_pr is None:
            # Also try picture element nvPicPr
            nv_sp_pr = shape_element.find(qn('p:nvPicPr'))
        if nv_sp_pr is None:
            return None

        # Search for ph element at any depth in nvSpPr
        ph = nv_sp_pr.find(f'.//{qn("p:ph")}')
        if ph is None:
            return None

        ph_type = ph.get('type', 'body')
        try:
            ph_idx = int(ph.get('idx', '0'))
        except (ValueError, TypeError):
            ph_idx = 0

        return ph_type, ph_idx
    except Exception:
        return None


def set_body_pr_rtl_col(txBody_element) -> None:
    """
    Set ``rtlCol='1'`` on the ``<a:bodyPr>`` of a text body element.

    This controls column direction in text frames with multiple columns
    and is required for proper Arabic text rendering.

    Args:
        txBody_element: An ``<a:txBody>`` lxml Element.
    """
    body_pr = txBody_element.find(qn('a:bodyPr'))
    if body_pr is not None:
        body_pr.set('rtlCol', '1')


def set_defRPr_lang(txBody_element, lang: str = 'ar-SA') -> None:
    """
    Set the ``lang`` attribute on all ``<a:defRPr>`` elements within *txBody_element*.

    Setting ``lang='ar-SA'`` ensures PowerPoint selects an Arabic-capable
    font for default runs, enabling correct shaping of Arabic glyphs.

    Args:
        txBody_element: An ``<a:txBody>`` lxml Element.
        lang: BCP-47 language tag, default ``'ar-SA'``.
    """
    for defRPr in txBody_element.iter(qn('a:defRPr')):
        defRPr.set('lang', lang)


def iter_paragraphs(txBody_element):
    """
    Yield all ``<a:p>`` paragraph elements within a text body.

    Args:
        txBody_element: An ``<a:txBody>`` lxml Element.

    Yields:
        ``<a:p>`` lxml elements in document order.
    """
    yield from txBody_element.findall(qn('a:p'))


def iter_runs(paragraph_element):
    """
    Yield all ``<a:r>`` run elements within a paragraph.

    Args:
        paragraph_element: An ``<a:p>`` lxml Element.

    Yields:
        ``<a:r>`` lxml elements in document order.
    """
    yield from paragraph_element.findall(qn('a:r'))


def get_run_text(run_element) -> str:
    """
    Extract text string from an ``<a:r>`` run element.

    Args:
        run_element: An ``<a:r>`` lxml Element.

    Returns:
        The text content of the run's ``<a:t>`` child, or ``''`` if absent.
    """
    t_elem = run_element.find(qn('a:t'))
    if t_elem is None:
        return ''
    return t_elem.text or ''


def set_run_text(run_element, text: str) -> None:
    """
    Set the text of an ``<a:r>`` run element.

    Creates the ``<a:t>`` child if it does not exist.

    Args:
        run_element: An ``<a:r>`` lxml Element.
        text: New text content.
    """
    t_elem = run_element.find(qn('a:t'))
    if t_elem is None:
        t_elem = etree.SubElement(run_element, qn('a:t'))
    t_elem.text = text


def get_or_create_rPr(run_element) -> etree._Element:
    """
    Get or create the ``<a:rPr>`` run properties element within an ``<a:r>``.

    The ``<a:rPr>`` element must appear before ``<a:t>`` per schema, so it
    is inserted at position 0 when created.

    Args:
        run_element: An ``<a:r>`` lxml Element.

    Returns:
        The existing or new ``<a:rPr>`` element.
    """
    rPr = run_element.find(qn('a:rPr'))
    if rPr is None:
        rPr = etree.Element(qn('a:rPr'))
        run_element.insert(0, rPr)
    return rPr


def set_run_language(run_element, lang: str = 'ar-SA') -> None:
    """
    Set the ``lang`` attribute on the ``<a:rPr>`` of a run.

    Args:
        run_element: An ``<a:r>`` lxml Element.
        lang: BCP-47 language tag, default ``'ar-SA'``.
    """
    rPr = get_or_create_rPr(run_element)
    rPr.set('lang', lang)


def bounds_check_emu(value: int, slide_dimension: int, label: str = '') -> bool:
    """
    Check whether a position value is within reasonable slide bounds.

    Allows for a generous negative margin (−1 500 000 EMU ≈ −1.64 in) for shapes
    that deliberately bleed off the slide edge (common in professional designs),
    and an equivalent positive margin beyond the slide dimension.

    Args:
        value: Position value in EMU to validate.
        slide_dimension: Slide width or height in EMU.
        label: Human-readable label for the check (used in warnings).

    Returns:
        True if the value is within acceptable bounds, False otherwise.
    """
    lower = -1_500_000
    upper = slide_dimension + 1_500_000
    return lower <= value <= upper


def clamp_emu(value: int, slide_dimension: int) -> int:
    """
    Clamp a position or size value to the range [-1_500_000, slide_dimension + 1_500_000].

    Generous bounds accommodate intentional design bleeds (shapes extending
    past slide edges) which are common in professional presentations.

    Args:
        value: EMU value to clamp.
        slide_dimension: Slide width or height in EMU.

    Returns:
        Clamped EMU value.
    """
    lower = -1_500_000
    upper = slide_dimension + 1_500_000
    return max(lower, min(upper, value))
