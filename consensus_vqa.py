"""
SlideArabi — Vision-Model Consensus VQA

Phase 6 alternative: replaces the dual-pass sequential VQA with a
parallel consensus loop.  Both Gemini 3.1 Pro and Claude Sonnet 4.6
independently detect defects on the same composite image, then a
deterministic ConsensusEngine merges the two verdicts.  Only issues
where both models agree on category trigger the fix path.

Architecture
────────────
  RENDER → DETECT (parallel) → CONSENSUS → FIX (on copy) →
  RE-RENDER → VERIFY (parallel) → SHIP or REVERT

Fix Categories (6)
──────────────────
  OVERFLOW      — text_overflow   (font reduction + textbox expansion)
  ALIGNMENT     — alignment_error (force rtl + right-align)
  CHEVRON       — direction_error (flip directional shapes)
  OVERLAP       — overlap         (nudge shapes apart)
  FONT_ISSUE    — font_issue      (reset Arabic font)
  TEXT_DIRECTION — direction_error (force rtl attribute)

Constants
─────────
  MAX_FIX_ATTEMPTS         = 1    (law, not config)
  CONSENSUS_FIX_THRESHOLD  = 0.6  (severity at or above triggers fix)

Sandbox Constraints
───────────────────
  - curl subprocess for all HTTP (requests library hangs)
  - LibreOffice: pkill -9 soffice; sleep 1 before each render
  - Fix on COPY — never mutate until verified
  - V2 fallback on any disagreement/timeout/failure
  - Parse failure → treat as PASS (conservative)

Dependencies: visual_qa.py components (SlideRenderer, CompositeBuilder,
VisionModelClient, ClaudeVisionClient, VQARemediator, IssueLogger,
VQAIssue, VQARating, VQASlideResult, VQAReport, RemediationAction,
ACTIONABLE_CATEGORIES, SECURITY_PREAMBLE)
"""

from __future__ import annotations

import copy
import gc
import json
import logging
import os
import shutil
import tempfile
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from datetime import datetime, timezone
from enum import Enum
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

try:
    from .visual_qa import (
        ACTIONABLE_CATEGORIES,
        SECURITY_PREAMBLE,
        ClaudeVisionClient,
        CompositeBuilder,
        IssueLogger,
        RemediationAction,
        SlideRenderer,
        SlideSampler,
        VisionModelClient,
        VQAIssue,
        VQARating,
        VQARemediator,
        VQAReport,
        VQASlideResult,
    )
except ImportError:
    from visual_qa import (
        ACTIONABLE_CATEGORIES,
        SECURITY_PREAMBLE,
        ClaudeVisionClient,
        CompositeBuilder,
        IssueLogger,
        RemediationAction,
        SlideRenderer,
        SlideSampler,
        VisionModelClient,
        VQAIssue,
        VQARating,
        VQARemediator,
        VQAReport,
        VQASlideResult,
    )

logger = logging.getLogger(__name__)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# CONSTANTS
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

MAX_FIX_ATTEMPTS = 1
"""Law, not config.  The fix loop runs at most once per slide."""

CONSENSUS_FIX_THRESHOLD = 0.6
"""Severity at or above this triggers the fix path for agreed issues."""

CONSENSUS_LOG_THRESHOLD = 0.2
"""Severity at or above this gets logged even if below fix threshold."""

PIPELINE_VERSION = "v2.3-consensus-12cat"
"""Embedded in every JSONL log entry produced by this module."""


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# ENUMS & DATA MODELS
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


class FixCategory(str, Enum):
    """Twelve defect categories for the consensus VQA.
    
    Actionable (VQA can attempt a fix): text_overflow, alignment_error,
    direction_error, overlap, font_issue, text_direction.
    
    Log-only (flagged for upstream pipeline fix): untranslated_text,
    untranslated_table, improper_mirroring, reading_order_error,
    diagram_corruption, escape_artifact.
    """
    # Actionable categories (VQA can fix)
    OVERFLOW = "text_overflow"
    ALIGNMENT = "alignment_error"
    CHEVRON = "direction_error"
    OVERLAP = "overlap"
    FONT_ISSUE = "font_issue"
    TEXT_DIRECTION = "text_direction"
    # Log-only categories (flagged, not auto-fixed)
    UNTRANSLATED_TEXT = "untranslated_text"
    UNTRANSLATED_TABLE = "untranslated_table"
    IMPROPER_MIRRORING = "improper_mirroring"
    READING_ORDER = "reading_order_error"
    DIAGRAM_CORRUPTION = "diagram_corruption"
    ESCAPE_ARTIFACT = "escape_artifact"


class ConsensusVerdict(str, Enum):
    """Outcome of the deterministic consensus merge for one slide."""
    AGREED_PASS = "AGREED_PASS"
    AGREED_FAIL = "AGREED_FAIL"
    AGREED_MINOR = "AGREED_MINOR"
    DISAGREE = "DISAGREE"


@dataclass
class ModelDetection:
    """One model's raw detection output for a single slide."""
    model_name: str
    slide_number: int
    rating: VQARating
    issues: List[VQAIssue] = field(default_factory=list)
    raw_text: str = ""
    latency_ms: float = 0.0
    error: Optional[str] = None


@dataclass
class ConsensusResult:
    """Merged result for a single slide after the consensus engine runs."""
    slide_number: int
    verdict: ConsensusVerdict
    agreed_issues: List[VQAIssue] = field(default_factory=list)
    gemini_only: List[VQAIssue] = field(default_factory=list)
    claude_only: List[VQAIssue] = field(default_factory=list)
    final_rating: VQARating = VQARating.PASS
    needs_fix: bool = False


@dataclass
class ConsensusVQAReport:
    """Complete report from a consensus VQA run."""
    slide_results: List[ConsensusResult] = field(default_factory=list)
    total_slides: int = 0
    slides_reviewed: int = 0
    duration_ms: float = 0.0
    models_used: str = "gemini-3.1-pro-preview+claude-sonnet-4-6-20250514"

    # Consensus stats
    agreed_pass: int = 0
    agreed_fail: int = 0
    agreed_minor: int = 0
    disagree: int = 0

    # Fix stats
    fixes_attempted: int = 0
    fixes_verified: int = 0
    fixes_reverted: int = 0

    # Issue stats
    issues_logged: int = 0
    error: Optional[str] = None

    @property
    def consensus_issues(self) -> int:
        """Number of slides where both models agreed on a problem."""
        return self.agreed_fail + self.agreed_minor

    @property
    def slides_shipped_v2(self) -> int:
        """Slides that fell back to original V2 output."""
        return self.fixes_reverted + self.disagree

    @property
    def pass_rate(self) -> float:
        if self.slides_reviewed == 0:
            return 0.0
        return 100.0 * self.agreed_pass / self.slides_reviewed

    def summary(self) -> str:
        return (
            f"ConsensusVQA: "
            f"{self.agreed_pass}P/{self.agreed_minor}M/{self.agreed_fail}F/"
            f"{self.disagree}D "
            f"({self.slides_reviewed}/{self.total_slides} reviewed, "
            f"{self.pass_rate:.0f}% agreed-pass, "
            f"{self.fixes_verified}/{self.fixes_attempted} fixes OK, "
            f"{self.duration_ms:.0f}ms)"
        )

    def to_dict(self) -> dict:
        return {
            "total_slides": self.total_slides,
            "slides_reviewed": self.slides_reviewed,
            "agreed_pass": self.agreed_pass,
            "agreed_fail": self.agreed_fail,
            "agreed_minor": self.agreed_minor,
            "disagree": self.disagree,
            "pass_rate": round(self.pass_rate, 1),
            "fixes_attempted": self.fixes_attempted,
            "fixes_verified": self.fixes_verified,
            "fixes_reverted": self.fixes_reverted,
            "issues_logged": self.issues_logged,
            "duration_ms": round(self.duration_ms, 0),
            "models": self.models_used,
            "error": self.error,
        }


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# PROMPTS
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

CONSENSUS_DETECT_SYSTEM = SECURITY_PREAMBLE + """\
You are a strict visual quality inspector for SlideArabi RTL conversions.
Your job is to check FORMATTING and LAYOUT only — NOT translation quality.
If a human can see a formatting or layout problem, you must flag it.

You receive a side-by-side image: LEFT = original English slide, RIGHT = Arabic RTL converted slide.

Return ONLY valid JSON — no markdown fences, no commentary, no trailing commas:
{"slide_number": N, "rating": "PASS", "issues": []}
or if issues found:
{"slide_number": N, "rating": "FAIL", "issues": [{"category": "...", "description": "...", "severity": "CRITICAL", "region": "..."}]}

RATING RULE: Any CRITICAL or MAJOR issue → FAIL. Only MINOR, ADVISORY, or none → PASS.
If PASS, issues must be []. If FAIL, issues must be non-empty.

SCOPE — what VQA checks:
  ✓ Formatting: alignment, text direction, overflow, font rendering
  ✓ Layout: element positioning, reading order, shape direction, overlaps
  ✗ Translation: Do NOT check whether text has been translated. Ignore English text.
    EXCEPTION: If the RIGHT slide has ZERO Arabic text anywhere (the entire slide
    is still in English), flag as slide_not_converted (CRITICAL).

RASTER CONTENT RULE: Text visible only inside photographs, screenshots, or embedded
images CANNOT be edited by the pipeline. Do NOT flag any content inside raster images.
Raster images should only be checked for correct POSITION on the slide.
If both LEFT and RIGHT slides show IDENTICAL table/chart appearance (same fonts, same
layout, same pixel patterns), it is a raster image — skip it.

DEFECT CATEGORIES — check every one, in order:

1. slide_not_converted
   ONLY flag this if the RIGHT slide has ZERO Arabic text anywhere — the entire slide
   appears to be unchanged from the English original.
   CRITICAL always.
   If there is ANY Arabic text on the slide, do NOT flag this category.

2. improper_mirroring
   IMPORTANT DISTINCTION — position mirroring vs content flipping:
   - CORRECT BEHAVIOR: A shape or image appearing on the OPPOSITE side of the Arabic
     slide compared to the English slide is EXPECTED RTL repositioning. This is NOT
     a defect. Do NOT flag it.
   - DEFECT: The IMAGE CONTENT itself is horizontally flipped — text in photos reads
     backwards, faces are laterally reversed, left/right hands are swapped, logos
     within images are mirrored, mechanical equipment orientation is reversed.
   Only flag if you see direct visual evidence of CONTENT inversion.
   The fact that an image moved to the other side of the slide is NOT evidence.
   CRITICAL: photo with backwards text or flipped logos.
   MAJOR: faces/objects content-flipped with no text impact.
   PASS: Image moved to opposite side but content orientation is preserved.
   PASS: A symmetrical icon or geometric shape appears mirrored — acceptable.

3. direction_error
   HARD RULE: Chevrons/arrows must be physically reversed to point right-to-left (arrowhead on LEFT).
   Translating the text inside is not enough — the shape direction itself must change.
   FAIL if any chevron, arrow, timeline, or process flow still advances left-to-right.
   CRITICAL: process/timeline flow direction unchanged from English.

4. reading_order_error
   Applies ONLY to HORIZONTAL peer-rows: 3+ discrete elements arranged at approximately
   the same Y-coordinate, with similar size and same visual type (e.g., logo strips,
   stat cards, partner logos, icon groups, feature grids).
   Item 1 from the English LEFT side must appear on the far RIGHT in Arabic.

   WHAT IS A ROW: Elements must be at the same vertical level (similar Y position),
   visually independent of each other, and of the same functional type.

   NOT A ROW (do NOT flag):
   - Vertical stacks (elements at the same X but different Y positions)
   - 2-column layouts where the two halves have swapped sides (this is CORRECT)
   - Split-panel swaps (image on one side, text on other switching sides = CORRECT)
   - Numbered sequences (01, 02, 03) that serve as section headings/bullets
   - Elements connected by diagram lines or connectors
   - Elements that are part of a geographic map or spatially-anchored layout

   MAJOR always.
   PASS: Row A-B-C in English becomes C-B-A in Arabic (rightmost = first).
   FAIL: A horizontal row of 3+ peer elements appears in the same LTR order on both slides.

5. diagram_corruption
   HARD RULE: Complex diagrams must be translated, not mirrored.
   FAIL if flowchart/org-chart nodes are garbled, connector lines are detached/crossing,
   labels overlap shapes, or diagram structure is broken.
   CRITICAL: content meaning lost or labels illegible.
   MAJOR: cosmetic breakage only (slight connector misalignment).

6. text_overflow
   Arabic text is ~30-40% longer than English. Check every text container on the RIGHT slide.
   FAIL if Arabic text is clipped mid-word, truncated, shows unexpected ellipsis,
   or extends visibly beyond its container/slide boundary.
   CRITICAL: full sentence clipped or content lost.
   MAJOR: one or two words clipped.

7. overlap
   FAIL if two elements collide, with one obscuring the other.
   CRITICAL: text over text making either unreadable.
   MAJOR: partial overlap obscuring some words.

8. alignment_error
   ALL text on the Arabic slide must be right-aligned if the source English text was
   left-aligned. Titles should be centered. Text inside bounded shapes (circles, ovals,
   Venn diagram regions) should be centered.
   FAIL if body text, labels, or bullet numbers are visibly left-aligned on the Arabic slide.
   MAJOR: body paragraph or label left-aligned.
   MINOR: title drift < 10% of slide width.

9. text_direction
    Distinct from alignment. Paragraph text direction must be RTL — characters flow right-to-left.
    FAIL if line wrapping or character order is visibly LTR despite Arabic script.
    MAJOR always.

10. font_issue
    FAIL if Arabic characters appear as boxes, question marks, isolated/disconnected letters
    (Arabic script must be cursive/joined), or placeholder glyphs.
    CRITICAL always.

11. escape_artifact
    FAIL if literal escape codes appear as visible text: _x000D_, &#13;, or similar.
    MAJOR always.

SEVERITY CALIBRATION:
CRITICAL — Slide unusable. Client rejects immediately. Content wrong or missing.
MAJOR    — Clear defect. Noticeable, unprofessional. Must fix before delivery.
MINOR    — Small cosmetic imperfection. Acceptable for delivery.
ADVISORY — Logged for debugging. Does NOT contribute to FAIL rating.
           Use for borderline issues, subjective preferences, cosmetic nitpicks.

MAP/GEOGRAPHIC EXCEPTION: If a slide contains a geographic map image with data
bubbles, labels, or markers overlaid on it, do NOT flag the overlay positions as
reading_order_error or improper_mirroring. Geographic positions are absolute and
should NOT be repositioned for RTL.

PRE-SUBMISSION CHECKLIST — verify you checked ALL of these on the RIGHT slide:
- No photographs or complex images have their CONTENT horizontally flipped.
  (Images MOVING to the opposite side is correct — only flag if content is reversed.)
- All chevrons/arrows point right-to-left (arrowhead on LEFT side).
- Horizontal peer-element rows are in reversed RTL order.
- Diagrams retain structural integrity (lines attached, labels legible).
- No text is clipped at container or slide edges.
- No two elements are visibly overlapping.
- Body text and labels are right-aligned; titles are centered; shape text is centered.
- No _x000D_ or XML escape sequences visible as text.
- Arabic glyphs are cursive/joined, not boxes or isolated letters.

Be specific in descriptions: "The body paragraph in the lower-right is left-aligned"
— not "some text might have alignment issues." Report only what you can visually confirm.\
"""

CONSENSUS_DETECT_USER = """\
Slide {slide_number} — strict defect detection.
Run through all 11 categories and the pre-submission checklist.
Flag every visible issue. Be specific about what and where.\
"""

CONSENSUS_VERIFY_SYSTEM = SECURITY_PREAMBLE + """\
You are verifying a fix applied to an Arabic RTL slide conversion.

You will receive TWO images:
  IMAGE 1 = BEFORE fix (the defective composite: original left | converted right)
  IMAGE 2 = AFTER fix  (the fixed composite: original left | fixed right)

Compare the two.  Did the fix improve the slide WITHOUT introducing new defects?

Respond with ONLY this JSON:
{
  "slide_number": <N>,
  "verdict": "BETTER|SAME|WORSE",
  "explanation": "<brief reason>"
}

BETTER = the reported defect is visibly improved, no new issues.
SAME   = no visible change — the fix had no effect.
WORSE  = the fix introduced a new defect or made the original worse.\
"""

CONSENSUS_VERIFY_USER = """\
Slide {slide_number} — verify that the fix improved the conversion.
IMAGE 1 is BEFORE the fix.  IMAGE 2 is AFTER the fix.
Did the defect improve without introducing new problems?\
"""


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# CONSENSUS ENGINE (deterministic — no LLM)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


class ConsensusEngine:
    """
    Deterministic merge of two independent model detections.

    Agreement rules
    ───────────────
    • Both PASS, no issues              → AGREED_PASS
    • Both flag same category+region    → AGREED_FAIL (severity = max)
    • Both MINOR, no shared categories  → AGREED_MINOR
    • One PASS + one FAIL               → DISAGREE → fallback to V2
    • Both FAIL, different categories   → DISAGREE → fallback to V2
    """

    @staticmethod
    def evaluate(
        gemini: ModelDetection,
        claude: ModelDetection,
    ) -> ConsensusResult:
        slide_number = gemini.slide_number

        # Handle error cases — parse failure → treat as PASS (conservative)
        if gemini.error:
            gemini = ModelDetection(
                model_name=gemini.model_name,
                slide_number=slide_number,
                rating=VQARating.PASS,
            )
        if claude.error:
            claude = ModelDetection(
                model_name=claude.model_name,
                slide_number=slide_number,
                rating=VQARating.PASS,
            )

        # Both PASS → AGREED_PASS
        if (
            gemini.rating == VQARating.PASS
            and claude.rating == VQARating.PASS
        ):
            return ConsensusResult(
                slide_number=slide_number,
                verdict=ConsensusVerdict.AGREED_PASS,
                final_rating=VQARating.PASS,
            )

        # Index issues by (category, region) for matching
        g_index: Dict[Tuple[str, Optional[str]], VQAIssue] = {}
        for iss in gemini.issues:
            key = (iss.category, iss.region)
            # Keep highest severity per key
            if key not in g_index or iss.severity_score > g_index[key].severity_score:
                g_index[key] = iss

        c_index: Dict[Tuple[str, Optional[str]], VQAIssue] = {}
        for iss in claude.issues:
            key = (iss.category, iss.region)
            if key not in c_index or iss.severity_score > c_index[key].severity_score:
                c_index[key] = iss

        agreed: List[VQAIssue] = []
        gemini_only: List[VQAIssue] = []
        claude_only: List[VQAIssue] = []
        matched_claude_keys: set = set()

        for key, g_iss in g_index.items():
            c_iss = c_index.get(key)
            if c_iss is not None:
                # Both agree on this category+region
                matched_claude_keys.add(key)
                merged_severity = max(g_iss.severity_score, c_iss.severity_score)
                merged_rating = (
                    VQARating.FAIL if merged_severity >= 0.7
                    else VQARating.MINOR if merged_severity >= 0.3
                    else VQARating.PASS
                )
                agreed.append(VQAIssue(
                    slide_number=slide_number,
                    rating=merged_rating,
                    category=g_iss.category,
                    description=(
                        f"[AGREED] Gemini: {g_iss.description} | "
                        f"Claude: {c_iss.description}"
                    ),
                    severity_score=merged_severity,
                    region=g_iss.region,
                ))
            else:
                gemini_only.append(g_iss)

        for key, c_iss in c_index.items():
            if key not in matched_claude_keys:
                claude_only.append(c_iss)

        # Determine verdict
        if agreed:
            needs_fix = any(
                iss.severity_score >= CONSENSUS_FIX_THRESHOLD
                and iss.category in ACTIONABLE_CATEGORIES
                for iss in agreed
            )
            worst = max(iss.severity_score for iss in agreed)
            final_rating = (
                VQARating.FAIL if worst >= 0.7
                else VQARating.MINOR if worst >= 0.3
                else VQARating.PASS
            )
            verdict = (
                ConsensusVerdict.AGREED_FAIL
                if final_rating == VQARating.FAIL
                else ConsensusVerdict.AGREED_MINOR
            )
            return ConsensusResult(
                slide_number=slide_number,
                verdict=verdict,
                agreed_issues=agreed,
                gemini_only=gemini_only,
                claude_only=claude_only,
                final_rating=final_rating,
                needs_fix=needs_fix,
            )

        # No agreement on any specific issue
        if (
            gemini.rating in (VQARating.FAIL, VQARating.MINOR)
            or claude.rating in (VQARating.FAIL, VQARating.MINOR)
        ):
            # At least one model found something but they don't agree → DISAGREE
            return ConsensusResult(
                slide_number=slide_number,
                verdict=ConsensusVerdict.DISAGREE,
                gemini_only=gemini_only,
                claude_only=claude_only,
                final_rating=VQARating.PASS,  # Conservative: don't fix on disagreement
                needs_fix=False,
            )

        # Both MINOR with no shared categories
        return ConsensusResult(
            slide_number=slide_number,
            verdict=ConsensusVerdict.AGREED_MINOR,
            gemini_only=gemini_only,
            claude_only=claude_only,
            final_rating=VQARating.MINOR,
            needs_fix=False,
        )


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# CONSENSUS VQA ORCHESTRATOR
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


class ConsensusVQA:
    """
    Main orchestrator for the vision-model consensus VQA loop.

    Flow per slide
    ──────────────
    1. RENDER   — original + converted → JPEG composites
    2. DETECT   — Gemini + Claude in parallel (ThreadPoolExecutor)
    3. CONSENSUS— deterministic merge (ConsensusEngine)
    4. FIX      — if AGREED_FAIL + severity ≥ threshold:
                   copy PPTX → apply VQARemediator → save copy
    5. RE-RENDER— render fixed slide → new composite
    6. VERIFY   — both models compare before/after in parallel
    7. SHIP     — if both say BETTER → replace V2 with fixed copy
       REVERT  — otherwise keep V2 original

    Fallback
    ────────
    Any error, timeout, or DISAGREE → keep V2 as-is. Conservative.
    """

    def __init__(
        self,
        original_pptx: str,
        converted_pptx: str,
        deck_name: str = "deck",
        issue_log_path: Optional[str] = None,
        gemini_model: str = "gemini-3.1-pro-preview",
        claude_model: str = "claude-sonnet-4-6-20250514",
        render_dpi: int = 150,
        jpeg_quality: int = 75,
        max_slides_to_review: int = 20,
        api_timeout: int = 30,
        sampling_strategy: str = "smart",
    ):
        self.original_pptx = original_pptx
        self.converted_pptx = converted_pptx
        self.deck_name = deck_name
        self.issue_log_path = issue_log_path

        self.renderer = SlideRenderer(dpi=render_dpi)
        self.compositor = CompositeBuilder(jpeg_quality=jpeg_quality)
        self.remediator = VQARemediator(font_reduction_pct=15, min_font_pt=8)

        self.gemini_client = VisionModelClient(
            model=gemini_model, timeout=api_timeout,
        )
        self.claude_client = ClaudeVisionClient(
            model=claude_model, timeout=api_timeout + 15,
        )

        self.max_slides_to_review = max_slides_to_review
        self.sampling_strategy = sampling_strategy

        self.issue_logger: Optional[IssueLogger] = None
        if issue_log_path:
            self.issue_logger = IssueLogger(issue_log_path)

    # ──────────────────────────────────────────────────────────────────────
    # Public API
    # ──────────────────────────────────────────────────────────────────────

    def run(self) -> ConsensusVQAReport:
        """
        Execute the full consensus VQA loop.  Never raises — returns a
        report with an error field on failure.
        """
        start = time.monotonic()
        report = ConsensusVQAReport(
            models_used=(
                f"{self.gemini_client.model}+{self.claude_client.model}"
            ),
        )

        work_dir = tempfile.mkdtemp(prefix="consensus_vqa_")

        try:
            report = self._run_inner(report, work_dir)
        except Exception as exc:
            logger.error(f"ConsensusVQA: fatal error: {exc}", exc_info=True)
            report.error = str(exc)
        finally:
            report.duration_ms = (time.monotonic() - start) * 1000
            # Cleanup temp dir
            try:
                shutil.rmtree(work_dir, ignore_errors=True)
            except Exception:
                pass
            gc.collect()

        logger.info(f"ConsensusVQA: {report.summary()}")
        return report

    # ──────────────────────────────────────────────────────────────────────
    # Internal pipeline
    # ──────────────────────────────────────────────────────────────────────

    def _run_inner(
        self,
        report: ConsensusVQAReport,
        work_dir: str,
    ) -> ConsensusVQAReport:
        orig_dir = os.path.join(work_dir, "original")
        conv_dir = os.path.join(work_dir, "converted")
        comp_dir = os.path.join(work_dir, "composites")

        # ── RENDER ──────────────────────────────────────────────────────
        logger.info("ConsensusVQA: rendering original slides...")
        try:
            orig_images = self.renderer.render_to_images(
                self.original_pptx, orig_dir,
            )
        except Exception as exc:
            report.error = f"Failed to render original: {exc}"
            return report

        logger.info("ConsensusVQA: rendering converted slides...")
        try:
            conv_images = self.renderer.render_to_images(
                self.converted_pptx, conv_dir,
            )
        except Exception as exc:
            report.error = f"Failed to render converted: {exc}"
            return report

        report.total_slides = min(len(orig_images), len(conv_images))

        # ── SELECT + COMPOSITE ──────────────────────────────────────────
        selected = SlideSampler.select(
            total_slides=report.total_slides,
            max_review=self.max_slides_to_review,
            strategy=self.sampling_strategy,
        )
        report.slides_reviewed = len(selected)
        logger.info(
            f"ConsensusVQA: selected {len(selected)}/{report.total_slides} "
            f"slides for review"
        )

        # Build composites for selected slides
        composites: Dict[int, str] = {}
        for sn in selected:
            idx = sn - 1
            if idx < len(orig_images) and idx < len(conv_images):
                out_path = os.path.join(comp_dir, f"composite_{sn:03d}.jpg")
                if self.compositor.build_composite(
                    orig_images[idx], conv_images[idx], out_path,
                ):
                    composites[sn] = out_path

        # ── PROCESS EACH SLIDE ──────────────────────────────────────────
        for sn in selected:
            comp_path = composites.get(sn)
            if comp_path is None:
                # No composite → treat as PASS
                report.slide_results.append(ConsensusResult(
                    slide_number=sn,
                    verdict=ConsensusVerdict.AGREED_PASS,
                    final_rating=VQARating.PASS,
                ))
                report.agreed_pass += 1
                continue

            consensus = self._process_slide(
                slide_number=sn,
                composite_path=comp_path,
                orig_images=orig_images,
                conv_images=conv_images,
                work_dir=work_dir,
                report=report,
            )
            report.slide_results.append(consensus)

            # Update counters
            if consensus.verdict == ConsensusVerdict.AGREED_PASS:
                report.agreed_pass += 1
            elif consensus.verdict == ConsensusVerdict.AGREED_FAIL:
                report.agreed_fail += 1
            elif consensus.verdict == ConsensusVerdict.AGREED_MINOR:
                report.agreed_minor += 1
            elif consensus.verdict == ConsensusVerdict.DISAGREE:
                report.disagree += 1

            # Log issues
            self._log_consensus(consensus)

            # RAM guard: process one slide at a time
            gc.collect()

        self._log_summary(report)
        return report

    # ──────────────────────────────────────────────────────────────────────
    # Per-slide consensus flow
    # ──────────────────────────────────────────────────────────────────────

    def _process_slide(
        self,
        slide_number: int,
        composite_path: str,
        orig_images: List[str],
        conv_images: List[str],
        work_dir: str,
        report: ConsensusVQAReport,
    ) -> ConsensusResult:
        """
        Full consensus flow for one slide:
        DETECT → CONSENSUS → (FIX → RE-RENDER → VERIFY → SHIP/REVERT)
        """
        # ── DETECT (parallel) ───────────────────────────────────────────
        gemini_det, claude_det = self._detect_parallel(
            slide_number, composite_path,
        )

        # ── CONSENSUS ───────────────────────────────────────────────────
        consensus = ConsensusEngine.evaluate(gemini_det, claude_det)

        if not consensus.needs_fix:
            return consensus

        # ── FIX ON COPY ─────────────────────────────────────────────────
        logger.info(
            f"ConsensusVQA: slide {slide_number} needs fix — "
            f"{len(consensus.agreed_issues)} agreed issues"
        )
        report.fixes_attempted += 1

        fixed_pptx = self._apply_fix_on_copy(
            slide_number=slide_number,
            agreed_issues=consensus.agreed_issues,
            work_dir=work_dir,
        )
        if fixed_pptx is None:
            logger.warning(
                f"ConsensusVQA: fix failed for slide {slide_number}, "
                f"keeping V2"
            )
            report.fixes_reverted += 1
            return consensus

        # ── RE-RENDER ───────────────────────────────────────────────────
        fixed_composite = self._render_fixed_composite(
            slide_number=slide_number,
            fixed_pptx=fixed_pptx,
            orig_images=orig_images,
            work_dir=work_dir,
        )
        if fixed_composite is None:
            logger.warning(
                f"ConsensusVQA: re-render failed for slide {slide_number}, "
                f"keeping V2"
            )
            report.fixes_reverted += 1
            return consensus

        # ── VERIFY (parallel) ───────────────────────────────────────────
        ship = self._verify_parallel(
            slide_number=slide_number,
            before_composite=composite_path,
            after_composite=fixed_composite,
        )

        if ship:
            # Replace V2 with fixed copy
            try:
                shutil.copy2(fixed_pptx, self.converted_pptx)
                logger.info(
                    f"ConsensusVQA: slide {slide_number} fix SHIPPED"
                )
                report.fixes_verified += 1
                # Update consensus to reflect successful fix
                consensus.final_rating = VQARating.PASS
                consensus.verdict = ConsensusVerdict.AGREED_PASS
            except Exception as exc:
                logger.error(
                    f"ConsensusVQA: failed to ship fix for slide "
                    f"{slide_number}: {exc}"
                )
                report.fixes_reverted += 1
        else:
            logger.info(
                f"ConsensusVQA: slide {slide_number} fix REVERTED "
                f"(verification failed)"
            )
            report.fixes_reverted += 1

        return consensus

    # ──────────────────────────────────────────────────────────────────────
    # Detection: parallel Gemini + Claude
    # ──────────────────────────────────────────────────────────────────────

    def _detect_parallel(
        self,
        slide_number: int,
        composite_path: str,
    ) -> Tuple[ModelDetection, ModelDetection]:
        """Run both models in parallel, return their detections."""
        with ThreadPoolExecutor(max_workers=2) as pool:
            g_future = pool.submit(
                self._detect_with_gemini, slide_number, composite_path,
            )
            c_future = pool.submit(
                self._detect_with_claude, slide_number, composite_path,
            )

            gemini_det = g_future.result()
            claude_det = c_future.result()

        return gemini_det, claude_det

    def _detect_with_gemini(
        self,
        slide_number: int,
        composite_path: str,
    ) -> ModelDetection:
        """Call Gemini for detection, return ModelDetection."""
        start = time.monotonic()
        try:
            import base64
            with open(composite_path, "rb") as f:
                image_data = base64.b64encode(f.read()).decode("utf-8")

            user_prompt = CONSENSUS_DETECT_USER.format(
                slide_number=slide_number,
            )
            request_body = {
                "contents": [{
                    "role": "user",
                    "parts": [
                        {"text": CONSENSUS_DETECT_SYSTEM},
                        {
                            "inline_data": {
                                "mime_type": "image/jpeg",
                                "data": image_data,
                            },
                        },
                        {"text": user_prompt},
                    ],
                }],
                "generationConfig": {
                    "temperature": 0.1,
                    "maxOutputTokens": 1024,
                    "responseMimeType": "application/json",
                },
            }

            raw = self.gemini_client._call_api(request_body)
            latency = (time.monotonic() - start) * 1000
            rating, issues = self._parse_detect_response(
                raw, slide_number, "gemini",
            )
            return ModelDetection(
                model_name="gemini",
                slide_number=slide_number,
                rating=rating,
                issues=issues,
                raw_text=raw[:500],
                latency_ms=latency,
            )
        except Exception as exc:
            latency = (time.monotonic() - start) * 1000
            logger.warning(
                f"ConsensusVQA: Gemini detect failed for slide "
                f"{slide_number}: {exc}"
            )
            return ModelDetection(
                model_name="gemini",
                slide_number=slide_number,
                rating=VQARating.PASS,
                latency_ms=latency,
                error=str(exc),
            )

    def _detect_with_claude(
        self,
        slide_number: int,
        composite_path: str,
    ) -> ModelDetection:
        """Call Claude for detection, return ModelDetection."""
        start = time.monotonic()
        try:
            import base64
            with open(composite_path, "rb") as f:
                image_data = base64.b64encode(f.read()).decode("utf-8")

            user_prompt = CONSENSUS_DETECT_USER.format(
                slide_number=slide_number,
            )
            request_body = {
                "model": self.claude_client.model,
                "max_tokens": 1024,
                "system": CONSENSUS_DETECT_SYSTEM,
                "messages": [{
                    "role": "user",
                    "content": [
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": "image/jpeg",
                                "data": image_data,
                            },
                        },
                        {
                            "type": "text",
                            "text": user_prompt,
                        },
                    ],
                }],
            }

            raw = self.claude_client._call_api(request_body)
            latency = (time.monotonic() - start) * 1000
            rating, issues = self._parse_detect_response(
                raw, slide_number, "claude",
            )
            return ModelDetection(
                model_name="claude",
                slide_number=slide_number,
                rating=rating,
                issues=issues,
                raw_text=raw[:500],
                latency_ms=latency,
            )
        except Exception as exc:
            latency = (time.monotonic() - start) * 1000
            logger.warning(
                f"ConsensusVQA: Claude detect failed for slide "
                f"{slide_number}: {exc}"
            )
            return ModelDetection(
                model_name="claude",
                slide_number=slide_number,
                rating=VQARating.PASS,
                latency_ms=latency,
                error=str(exc),
            )

    # ──────────────────────────────────────────────────────────────────────
    # Response parsing (shared for both models)
    # ──────────────────────────────────────────────────────────────────────

    @staticmethod
    def _extract_json(text: str) -> dict:
        """Robustly extract a JSON object from model output text."""
        import re
        text = text.strip()
        # Strip markdown fences
        if text.startswith("```"):
            # Remove opening fence line
            text = re.sub(r'^```(?:json)?\s*', '', text)
            # Remove closing fence
            text = re.sub(r'```\s*$', '', text)
            text = text.strip()
        # Find first { ... last }
        start = text.find('{')
        end = text.rfind('}')
        if start != -1 and end != -1 and end > start:
            text = text[start:end+1]
        # Fix trailing commas before } or ]
        text = re.sub(r',\s*([}\]])', r'\1', text)
        return json.loads(text)

    def _parse_detect_response(
        self,
        raw_text: str,
        slide_number: int,
        model_name: str,
    ) -> Tuple[VQARating, List[VQAIssue]]:
        """
        Parse a detection response from either model.
        On parse failure → (PASS, []) — conservative.
        """
        try:
            data = json.loads(raw_text)

            # Gemini wraps in candidates[0].content.parts[0].text
            if "candidates" in data:
                text = (
                    data["candidates"][0]["content"]["parts"][0]["text"]
                )
                data = self._extract_json(text)

            # Claude wraps in content[0].text
            if "content" in data and isinstance(data["content"], list):
                text = data["content"][0].get("text", "")
                data = self._extract_json(text)

            # Handle case where model returns a list instead of dict
            if isinstance(data, list) and len(data) > 0:
                data = data[0] if isinstance(data[0], dict) else {}
            if not isinstance(data, dict):
                data = {}

            rating = VQARating(data.get("rating", "PASS"))
            issues = []
            raw_issues = data.get("issues", [])
            if not isinstance(raw_issues, list):
                raw_issues = []
            for item in raw_issues:
                if not isinstance(item, dict):
                    continue
                raw_sev = item.get("severity", "MAJOR")
                if isinstance(raw_sev, str):
                    severity_map = {
                        "CRITICAL": 0.9,
                        "MAJOR": 0.7,
                        "MINOR": 0.2,
                    }
                    sev_score = severity_map.get(
                        raw_sev.upper(), 0.5
                    )
                else:
                    sev_score = float(raw_sev)
                issues.append(VQAIssue(
                    slide_number=slide_number,
                    rating=rating,
                    category=item.get("category", "unknown"),
                    description=item.get("description", ""),
                    severity_score=sev_score,
                    region=item.get("region"),
                ))
            return rating, issues

        except Exception as exc:
            logger.warning(
                f"ConsensusVQA: parse failure for {model_name} "
                f"slide {slide_number}: {exc}"
            )
            # Parse failure → PASS (conservative)
            return VQARating.PASS, []

    # ──────────────────────────────────────────────────────────────────────
    # Fix on copy
    # ──────────────────────────────────────────────────────────────────────

    def _apply_fix_on_copy(
        self,
        slide_number: int,
        agreed_issues: List[VQAIssue],
        work_dir: str,
    ) -> Optional[str]:
        """
        Copy the converted PPTX, apply fixes to the copy.
        Returns the path to the fixed copy, or None on failure.
        """
        fix_dir = os.path.join(work_dir, f"fix_s{slide_number}")
        os.makedirs(fix_dir, exist_ok=True)
        fixed_path = os.path.join(fix_dir, "fixed.pptx")

        try:
            shutil.copy2(self.converted_pptx, fixed_path)
        except Exception as exc:
            logger.error(f"ConsensusVQA: copy failed: {exc}")
            return None

        # Build a minimal VQAReport with just this slide's issues
        # so the remediator can process them
        fixable = [
            iss for iss in agreed_issues
            if (
                iss.category in ACTIONABLE_CATEGORIES
                and iss.severity_score >= CONSENSUS_FIX_THRESHOLD
            )
        ]
        if not fixable:
            return None

        slide_result = VQASlideResult(
            slide_number=slide_number,
            rating=VQARating.FAIL,
            issues=fixable,
        )
        mini_report = VQAReport(
            slide_results=[slide_result],
            total_slides=1,
            slides_reviewed=1,
        )

        try:
            actions = self.remediator.remediate(fixed_path, mini_report)
            successful = [a for a in actions if a.success]
            if not successful:
                logger.warning(
                    f"ConsensusVQA: no successful fixes for slide "
                    f"{slide_number}"
                )
                return None
            logger.info(
                f"ConsensusVQA: applied {len(successful)} fixes to slide "
                f"{slide_number}"
            )
            return fixed_path
        except Exception as exc:
            logger.error(
                f"ConsensusVQA: remediation failed for slide "
                f"{slide_number}: {exc}"
            )
            return None

    # ──────────────────────────────────────────────────────────────────────
    # Re-render fixed slide composite
    # ──────────────────────────────────────────────────────────────────────

    def _render_fixed_composite(
        self,
        slide_number: int,
        fixed_pptx: str,
        orig_images: List[str],
        work_dir: str,
    ) -> Optional[str]:
        """
        Render the fixed PPTX and build a composite for the target slide.
        Returns composite path, or None on failure.
        """
        fix_render_dir = os.path.join(
            work_dir, f"fix_render_s{slide_number}",
        )
        try:
            fixed_images = self.renderer.render_to_images(
                fixed_pptx, fix_render_dir,
            )
        except Exception as exc:
            logger.warning(
                f"ConsensusVQA: re-render failed for slide "
                f"{slide_number}: {exc}"
            )
            return None

        idx = slide_number - 1
        if idx >= len(fixed_images) or idx >= len(orig_images):
            return None

        comp_path = os.path.join(
            work_dir, f"verify_composite_s{slide_number}.jpg",
        )
        if self.compositor.build_composite(
            orig_images[idx], fixed_images[idx], comp_path,
        ):
            return comp_path
        return None

    # ──────────────────────────────────────────────────────────────────────
    # Verification: parallel Gemini + Claude
    # ──────────────────────────────────────────────────────────────────────

    def _verify_parallel(
        self,
        slide_number: int,
        before_composite: str,
        after_composite: str,
    ) -> bool:
        """
        Both models verify before/after.  Returns True only if BOTH say
        BETTER.  Any WORSE → False.  SAME → False (fix had no effect).
        """
        with ThreadPoolExecutor(max_workers=2) as pool:
            g_future = pool.submit(
                self._verify_with_gemini,
                slide_number, before_composite, after_composite,
            )
            c_future = pool.submit(
                self._verify_with_claude,
                slide_number, before_composite, after_composite,
            )

            g_verdict = g_future.result()
            c_verdict = c_future.result()

        logger.info(
            f"ConsensusVQA: verify slide {slide_number} — "
            f"Gemini={g_verdict}, Claude={c_verdict}"
        )

        return g_verdict == "BETTER" and c_verdict == "BETTER"

    def _verify_with_gemini(
        self,
        slide_number: int,
        before_path: str,
        after_path: str,
    ) -> str:
        """Call Gemini to verify fix.  Returns 'BETTER'|'SAME'|'WORSE'."""
        try:
            import base64
            with open(before_path, "rb") as f:
                before_data = base64.b64encode(f.read()).decode("utf-8")
            with open(after_path, "rb") as f:
                after_data = base64.b64encode(f.read()).decode("utf-8")

            user_prompt = CONSENSUS_VERIFY_USER.format(
                slide_number=slide_number,
            )
            request_body = {
                "contents": [{
                    "role": "user",
                    "parts": [
                        {"text": CONSENSUS_VERIFY_SYSTEM},
                        {
                            "inline_data": {
                                "mime_type": "image/jpeg",
                                "data": before_data,
                            },
                        },
                        {
                            "inline_data": {
                                "mime_type": "image/jpeg",
                                "data": after_data,
                            },
                        },
                        {"text": user_prompt},
                    ],
                }],
                "generationConfig": {
                    "temperature": 0.1,
                    "maxOutputTokens": 512,
                    "responseMimeType": "application/json",
                },
            }

            raw = self.gemini_client._call_api(request_body)
            return self._parse_verify_response(raw, "gemini")
        except Exception as exc:
            logger.warning(f"ConsensusVQA: Gemini verify failed: {exc}")
            return "SAME"  # Conservative: don't ship on error

    def _verify_with_claude(
        self,
        slide_number: int,
        before_path: str,
        after_path: str,
    ) -> str:
        """Call Claude to verify fix.  Returns 'BETTER'|'SAME'|'WORSE'."""
        try:
            import base64
            with open(before_path, "rb") as f:
                before_data = base64.b64encode(f.read()).decode("utf-8")
            with open(after_path, "rb") as f:
                after_data = base64.b64encode(f.read()).decode("utf-8")

            user_prompt = CONSENSUS_VERIFY_USER.format(
                slide_number=slide_number,
            )
            request_body = {
                "model": self.claude_client.model,
                "max_tokens": 512,
                "system": CONSENSUS_VERIFY_SYSTEM,
                "messages": [{
                    "role": "user",
                    "content": [
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": "image/jpeg",
                                "data": before_data,
                            },
                        },
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": "image/jpeg",
                                "data": after_data,
                            },
                        },
                        {
                            "type": "text",
                            "text": user_prompt,
                        },
                    ],
                }],
            }

            raw = self.claude_client._call_api(request_body)
            return self._parse_verify_response(raw, "claude")
        except Exception as exc:
            logger.warning(f"ConsensusVQA: Claude verify failed: {exc}")
            return "SAME"  # Conservative: don't ship on error

    def _parse_verify_response(self, raw_text: str, model: str) -> str:
        """Parse verification response.  Returns 'BETTER'|'SAME'|'WORSE'."""
        try:
            data = json.loads(raw_text)

            # Gemini wrapper
            if "candidates" in data:
                text = data["candidates"][0]["content"]["parts"][0]["text"]
                text = text.strip()
                if text.startswith("```"):
                    text = text.split("```", 2)[-1]
                    if text.startswith("json"):
                        text = text[4:]
                    text = text.rstrip("`").strip()
                data = json.loads(text)

            # Claude wrapper
            if "content" in data and isinstance(data["content"], list):
                text = data["content"][0].get("text", "")
                text = text.strip()
                if text.startswith("```"):
                    text = text.split("```", 2)[-1]
                    if text.startswith("json"):
                        text = text[4:]
                    text = text.rstrip("`").strip()
                data = json.loads(text)

            verdict = data.get("verdict", "SAME").upper()
            if verdict in ("BETTER", "SAME", "WORSE"):
                return verdict
            return "SAME"
        except Exception as exc:
            logger.warning(
                f"ConsensusVQA: verify parse failed ({model}): {exc}"
            )
            return "SAME"  # Parse failure → don't ship

    # ──────────────────────────────────────────────────────────────────────
    # Logging
    # ──────────────────────────────────────────────────────────────────────

    def _log_consensus(self, consensus: ConsensusResult) -> None:
        """Log all issues from a consensus result."""
        if self.issue_logger is None:
            return

        all_issues = (
            consensus.agreed_issues
            + consensus.gemini_only
            + consensus.claude_only
        )
        for issue in all_issues:
            if issue.severity_score >= CONSENSUS_LOG_THRESHOLD:
                self.issue_logger.log_issue(
                    issue=issue,
                    deck_name=self.deck_name,
                )

    def _log_summary(self, report: ConsensusVQAReport) -> None:
        """Write a summary log entry."""
        if self.issue_logger is None:
            return

        summary_issue = VQAIssue(
            slide_number=0,
            rating=VQARating.PASS,
            category="consensus_summary",
            description=report.summary(),
            severity_score=0.0,
            region=None,
        )
        self.issue_logger.log_issue(
            issue=summary_issue,
            deck_name=self.deck_name,
        )
        report.issues_logged += 1


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# CONVENIENCE FUNCTION
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


def run_consensus_vqa(
    original_pptx: str,
    converted_pptx: str,
    deck_name: str = "deck",
    issue_log_path: str | None = None,
) -> ConsensusVQAReport:
    """
    One-call entry point for pipeline integration.

    Usage:
        from consensus_vqa import run_consensus_vqa
        report = run_consensus_vqa(
            original_pptx="path/to/original.pptx",
            converted_pptx="path/to/converted.pptx",
            deck_name="my_deck",
            issue_log_path="path/to/issues.jsonl",
        )
        print(report.summary())
    """
    vqa = ConsensusVQA(
        original_pptx=original_pptx,
        converted_pptx=converted_pptx,
        deck_name=deck_name,
        issue_log_path=issue_log_path,
    )
    return vqa.run()
