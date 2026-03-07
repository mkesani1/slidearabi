"""
smartart_translator.py — Translate text in SmartArt diagrams.

SmartArt (diagram) text lives in separate XML parts:
  - ppt/diagrams/dataX.xml  → source data (for re-editing)
  - ppt/diagrams/drawingX.xml → rendered shapes (for display)

python-pptx does NOT expose SmartArt text through its shape API.
This module operates on the saved PPTX (as a ZIP file) to translate
text in both data and drawing XML parts.

Usage:
    translate_smartart_in_pptx(pptx_path, translations)
"""

import logging
import re
import zipfile
import shutil
import tempfile
from pathlib import Path
from typing import Dict, Optional
from lxml import etree

logger = logging.getLogger(__name__)

A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
DGM_NS = 'http://schemas.openxmlformats.org/drawingml/2006/diagram'


def _has_arabic(text: str) -> bool:
    """Check if text contains Arabic characters."""
    return bool(re.search(r'[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]', text))


def _fuzzy_lookup(text: str, translations: Dict[str, str],
                  translations_lower: Dict[str, str]) -> Optional[str]:
    """
    Look up translation with fallbacks:
    1. Exact match
    2. Stripped match
    3. Case-insensitive match
    """
    if not text or not text.strip():
        return None
    stripped = text.strip()
    result = translations.get(text) or translations.get(stripped)
    if result:
        return result
    lower = stripped.lower()
    return translations_lower.get(lower)


def translate_smartart_in_pptx(
    pptx_path: str,
    translations: Dict[str, str],
) -> int:
    """
    Translate SmartArt text in a saved PPTX file.

    Opens the PPTX as a ZIP, finds all diagram data and drawing XML files,
    translates text, sets RTL properties, and saves back.

    Args:
        pptx_path: Path to the saved PPTX file.
        translations: Dict mapping English text → Arabic text.

    Returns:
        Count of text items translated.
    """
    if not translations:
        return 0

    # Pre-build lowercase index
    translations_lower: Dict[str, str] = {}
    for k, v in translations.items():
        lk = k.strip().lower()
        if lk not in translations_lower:
            translations_lower[lk] = v

    # Find diagram XML parts
    total_translated = 0

    # Work with a temporary file to avoid corruption
    tmp_path = pptx_path + '.smartart_tmp'
    shutil.copy2(pptx_path, tmp_path)

    try:
        # Read all entries, modify diagram XMLs, write back
        with zipfile.ZipFile(tmp_path, 'r') as zin:
            names = zin.namelist()
            diagram_parts = [
                n for n in names
                if n.startswith('ppt/diagrams/')
                and (n.endswith('.xml'))
                and ('/data' in n or '/drawing' in n)
            ]

            if not diagram_parts:
                logger.debug('No SmartArt diagrams found in %s', pptx_path)
                return 0

            logger.info('Found %d SmartArt XML parts in %s', len(diagram_parts), pptx_path)

            # Process each diagram XML
            modified_parts: Dict[str, bytes] = {}
            for part_name in diagram_parts:
                xml_bytes = zin.read(part_name)
                try:
                    root = etree.fromstring(xml_bytes)
                except etree.XMLSyntaxError:
                    continue

                count = _translate_xml_element(
                    root, translations, translations_lower
                )
                if count > 0:
                    modified_parts[part_name] = etree.tostring(
                        root, xml_declaration=True, encoding='UTF-8', standalone=True
                    )
                    total_translated += count
                    logger.info('Translated %d items in %s', count, part_name)

            if not modified_parts:
                return 0

            # Rewrite the ZIP with modified parts
            with zipfile.ZipFile(pptx_path, 'w', zipfile.ZIP_STORED) as zout:
                for name in names:
                    if name in modified_parts:
                        zout.writestr(name, modified_parts[name])
                    else:
                        zout.writestr(name, zin.read(name))

    finally:
        # Clean up temp file
        try:
            Path(tmp_path).unlink(missing_ok=True)
        except Exception:
            pass

    logger.info('SmartArt translation complete: %d items translated in %s',
                total_translated, pptx_path)
    return total_translated


def _translate_xml_element(
    root: etree._Element,
    translations: Dict[str, str],
    translations_lower: Dict[str, str],
) -> int:
    """
    Translate text in a SmartArt XML element tree.

    For each paragraph (<a:p>):
    1. Gather all run text to form the full paragraph text.
    2. Look up translation.
    3. If found, replace text in first run, clear subsequent runs.
    4. Set RTL properties on the paragraph.

    Also handles multi-run text that was split by PowerPoint
    (e.g., "Weaponised D" + "rones" → "Weaponised Drones").

    Returns:
        Count of paragraphs translated.
    """
    count = 0

    for p in root.iter(f'{{{A_NS}}}p'):
        # Gather paragraph text from runs
        runs = p.findall(f'{{{A_NS}}}r')
        if not runs:
            continue

        para_text = ''
        for r in runs:
            t = r.find(f'{{{A_NS}}}t')
            if t is not None and t.text:
                para_text += t.text

        if not para_text.strip():
            continue

        # Already Arabic — skip
        if _has_arabic(para_text):
            continue

        # Look up translation
        arabic = _fuzzy_lookup(para_text, translations, translations_lower)
        if not arabic:
            continue

        # Replace: put full Arabic text in first run, clear rest
        first_t = runs[0].find(f'{{{A_NS}}}t')
        if first_t is not None:
            first_t.text = arabic
        for r in runs[1:]:
            t = r.find(f'{{{A_NS}}}t')
            if t is not None:
                t.text = ''

        # Set Arabic language on run properties
        for r in runs:
            rPr = r.find(f'{{{A_NS}}}rPr')
            if rPr is not None:
                rPr.set('lang', 'ar-SA')

        # Set RTL on paragraph properties
        pPr = p.find(f'{{{A_NS}}}pPr')
        if pPr is None:
            pPr = etree.SubElement(p, f'{{{A_NS}}}pPr')
            p.insert(0, pPr)
        pPr.set('rtl', '1')

        count += 1

    return count
