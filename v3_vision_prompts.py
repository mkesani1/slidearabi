"""
v3_vision_prompts.py — V3 Enhanced Vision Prompt Templates

Sprint 3: Enhanced prompts for table/icon/alignment/page number/directional checks.
Includes structured JSON output schema and XML context injection.
All gated behind v3_config flags.
"""

from __future__ import annotations

import json
import logging
from typing import Any, Dict, List, Optional

import v3_config

logger = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────────────────────
# ENHANCED PROMPT ADDITIONS
# ─────────────────────────────────────────────────────────────────────────────

V3_PROMPT_ADDITIONS = """
## TABLE STRUCTURE VALIDATION (CRITICAL)

T1. TABLE COLUMN ORDER: Compare header rows between original and converted.
    In Arabic RTL, columns must appear in REVERSED left-to-right order vs. English.
    If column order appears UNCHANGED: FAIL, code=TABLE_COLUMNS_NOT_REVERSED, severity=0.9.

T2. TABLE CELL CONTENT: Verify check/cross symbols are in the correct mirrored cells.
    If a checkmark appears in the same unmirrored column: FAIL, code=ICON_CELL_MISMATCH.

T3. TABLE VISUAL COHERENCE: Flag any table that appears garbled, jumbled,
    with overlapping text or mismatched row heights: FAIL, code=TABLE_GARBLED.

T4. CELL TEXT ALIGNMENT: Arabic text and numbers in table cells must be
    RIGHT-ALIGNED. If visibly left-aligned: FAIL, code=TABLE_CELL_ALIGNMENT.

## PAGE NUMBER / SLIDE NUMBER
P1. If slide number appears doubled (e.g., "1515" instead of "15"):
    FAIL, code=PAGE_NUMBER_DUPLICATED.

## DECORATIVE ELEMENTS
S1. Template lines/shapes/borders should mirror to opposite side in RTL.
    If template elements remain on same side as English original:
    FAIL, code=MASTER_ELEMENT_NOT_MIRRORED.

## CENTERED SHAPES
S2. Text in Venn diagrams, circles, ellipses should be visually centered.
    If pushed to one edge: WARN, code=TEXT_NOT_CENTERED_IN_SHAPE.

## DIRECTIONAL SHAPES
D1. Arrows, chevrons, and other directional shapes must point LEFT in Arabic RTL.
    If a rightward arrow still points right: FAIL, code=DIRECTIONAL_SHAPE_NOT_FLIPPED.
"""

V3_OUTPUT_SCHEMA = """\
Respond with JSON only. Schema:
{
  "slide_index": <int>,
  "verdict": "PASS|FAIL|WARN",
  "checks": {
    "table_columns_not_reversed": "PASS|FAIL|NA",
    "table_icon_position_error": "PASS|FAIL|NA",
    "table_cell_alignment_error": "PASS|FAIL|NA",
    "table_garbled_layout": "PASS|FAIL|NA",
    "page_number_duplicated": "PASS|FAIL|NA",
    "master_not_mirrored": "PASS|FAIL|NA",
    "text_not_centered": "PASS|FAIL|NA",
    "directional_not_flipped": "PASS|FAIL|NA"
  },
  "confidence": <float 0.0-1.0>,
  "issues": [{"category": "...", "severity": <float>, "description": "...", "region": "..."}]
}
"""


def build_enhanced_system_prompt(base_prompt: str) -> str:
    """Append V3 table/icon/alignment sections to the base VQA system prompt.
    
    Gated by v3_config.ENABLE_ENHANCED_PROMPTS.
    """
    if not v3_config.ENABLE_ENHANCED_PROMPTS:
        return base_prompt

    return base_prompt + V3_PROMPT_ADDITIONS


def build_enhanced_user_prompt(
    base_prompt: str,
    slide_idx: int,
    xml_defects: Optional[List[Dict[str, Any]]] = None,
) -> str:
    """Optionally append XML findings and structured output schema to user prompt.
    
    - V3_ENHANCED_PROMPTS: adds structured JSON output requirement
    - V3_VISION_XML_CONTEXT: adds XML defect context for vision confirmation
    """
    prompt = base_prompt

    # Add structured output schema
    if v3_config.ENABLE_ENHANCED_PROMPTS:
        prompt += f"\n\n## OUTPUT FORMAT\n{V3_OUTPUT_SCHEMA}"

    # Add XML findings as context for vision to confirm/reject
    if v3_config.ENABLE_VISION_XML_CTX and xml_defects:
        xml_context = (
            "\n\n## XML STRUCTURAL FINDINGS (pre-computed, deterministic)\n"
            "The following issues were detected by XML analysis on this slide:\n"
            f"{json.dumps(xml_defects, indent=2)}\n"
            "Please visually confirm or reject each finding.\n"
        )
        prompt += xml_context

    return prompt


# ─────────────────────────────────────────────────────────────────────────────
# SELECTIVE VISION — DECIDE WHICH SLIDES NEED VISION QA
# ─────────────────────────────────────────────────────────────────────────────

def select_slides_for_vision(
    all_defects: list,
    total_slides: int,
    max_vision_slides: Optional[int] = None,
) -> List[int]:
    """Determine which slides need vision QA based on XML findings.
    
    Strategy:
    - Slides with CRITICAL/HIGH unresolved defects → always vision
    - Slides with no defects at all → skip vision (XML says clean)
    - Slides with MEDIUM defects → vision if within budget
    - Respect max_vision_slides cost cap
    
    Returns: list of 1-based slide indices to scan with vision.
    """
    if not v3_config.ENABLE_SELECTIVE_VISION:
        # Non-selective: scan all slides up to budget
        max_slides = max_vision_slides or v3_config.MAX_VISION_SLIDES
        return list(range(1, min(total_slides + 1, max_slides + 1)))

    from vqa_types import Severity, DefectStatus

    max_slides = max_vision_slides or v3_config.MAX_VISION_SLIDES

    # Group unresolved defects by slide
    unresolved = [d for d in all_defects
                  if d.status.value == 'open' or d.status.value == 'unresolved']

    by_slide: Dict[int, list] = {}
    for d in unresolved:
        by_slide.setdefault(d.slide_idx, []).append(d)

    # Priority tiers
    tier1_slides = set()  # CRITICAL/HIGH → must scan
    tier2_slides = set()  # MEDIUM → scan if budget
    tier3_slides = set()  # No defects → scan if budget

    for slide_num in range(1, total_slides + 1):
        slide_defects = by_slide.get(slide_num, [])
        if not slide_defects:
            tier3_slides.add(slide_num)
        elif any(d.severity in (Severity.CRITICAL, Severity.HIGH) for d in slide_defects):
            tier1_slides.add(slide_num)
        else:
            tier2_slides.add(slide_num)

    # Build final list respecting budget
    selected = sorted(tier1_slides)

    if len(selected) < max_slides:
        remaining = max_slides - len(selected)
        selected.extend(sorted(tier2_slides)[:remaining])

    if len(selected) < max_slides:
        remaining = max_slides - len(selected)
        selected.extend(sorted(tier3_slides)[:remaining])

    logger.info(f"Selective vision: {len(selected)} slides selected "
                f"(tier1={len(tier1_slides)}, tier2={len(tier2_slides)}, "
                f"tier3={len(tier3_slides)}, cap={max_slides})")

    return sorted(set(selected))[:max_slides]
