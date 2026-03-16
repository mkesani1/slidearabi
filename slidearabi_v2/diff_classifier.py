"""
diff_classifier.py — Classify shape diffs as REGRESSION, IMPROVEMENT, or NEUTRAL.

Takes the raw ShapeDiff / SlideDiff output from structural_differ and applies
heuristic rules to label each difference.  Optional golden-slide baselines
allow comparison against known-good positions.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from enum import Enum
from typing import Dict, List, Optional, Tuple

from .structural_differ import (
    POSITION_TOLERANCE_EMU,
    EMU_PER_INCH,
    ShapeDiff,
    SlideDiff,
)

logger = logging.getLogger(__name__)

# Bounds threshold — matches utils.py clamp range
BOUNDS_MARGIN_EMU = 1_500_000


# ─────────────────────────────────────────────────────────────────────────────
# Enums & data classes
# ─────────────────────────────────────────────────────────────────────────────

class DiffLabel(Enum):
    REGRESSION = 'REGRESSION'
    IMPROVEMENT = 'IMPROVEMENT'
    NEUTRAL = 'NEUTRAL'
    UNKNOWN = 'UNKNOWN'


@dataclass
class ClassifiedDiff:
    """A ShapeDiff with an attached label and explanation."""
    shape_diff: ShapeDiff
    label: DiffLabel
    reason: str
    severity: str = 'info'  # 'critical' | 'major' | 'minor' | 'info'


@dataclass
class ClassifiedSlide:
    """All classified diffs for one slide, plus aggregate counts."""
    slide_number: int
    layout_type: str
    diffs: List[ClassifiedDiff] = field(default_factory=list)

    @property
    def regressions(self) -> int:
        return sum(1 for d in self.diffs if d.label == DiffLabel.REGRESSION)

    @property
    def improvements(self) -> int:
        return sum(1 for d in self.diffs if d.label == DiffLabel.IMPROVEMENT)

    @property
    def neutrals(self) -> int:
        return sum(1 for d in self.diffs if d.label == DiffLabel.NEUTRAL)


# ─────────────────────────────────────────────────────────────────────────────
# Golden slide baseline (optional)
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class GoldenShape:
    """Expected position for a shape in a golden slide."""
    shape_id: int
    expected_left: int
    expected_width: int
    tolerance: int = POSITION_TOLERANCE_EMU


@dataclass
class GoldenSlide:
    """Expected positions for all shapes on a golden slide."""
    slide_number: int
    shapes: Dict[int, GoldenShape] = field(default_factory=dict)


# ─────────────────────────────────────────────────────────────────────────────
# Background / full-width detection
# ─────────────────────────────────────────────────────────────────────────────

def _is_background_shape(diff: ShapeDiff, slide_width: int) -> bool:
    """A shape is background-like if its v1 width covers ≥ 95% of the slide."""
    v1_w = diff.v1_width if not diff.v2_only else diff.v2_width
    return v1_w >= slide_width * 0.95


def _is_out_of_bounds(diff: ShapeDiff, slide_width: int) -> bool:
    """Check if v2 position falls outside the allowed bounds range."""
    if diff.v2_only or diff.v1_only:
        return False
    left = diff.v2_left
    right = diff.v2_left + diff.v2_width
    return left < -BOUNDS_MARGIN_EMU or right > slide_width + BOUNDS_MARGIN_EMU


def _v2_mirrors_but_v1_didnt(diff: ShapeDiff) -> bool:
    """v2 moved the shape significantly but v1 left it alone (relative to original)."""
    return abs(diff.left_delta) > POSITION_TOLERANCE_EMU


def _shapes_overlap(a: ShapeDiff, b: ShapeDiff) -> bool:
    """Check if two shapes' v2 positions overlap horizontally."""
    a_left, a_right = a.v2_left, a.v2_left + a.v2_width
    b_left, b_right = b.v2_left, b.v2_left + b.v2_width
    a_top, a_bottom = a.v2_top, a.v2_top + a.v2_height
    b_top, b_bottom = b.v2_top, b.v2_top + b.v2_height
    h_overlap = a_left < b_right and b_left < a_right
    v_overlap = a_top < b_bottom and b_top < a_bottom
    return h_overlap and v_overlap


# ─────────────────────────────────────────────────────────────────────────────
# DiffClassifier
# ─────────────────────────────────────────────────────────────────────────────

class DiffClassifier:
    """Classify each ShapeDiff as REGRESSION / IMPROVEMENT / NEUTRAL / UNKNOWN.

    Rules (applied in priority order):
    1. Position delta < tolerance → NEUTRAL
    2. Shape only in v1 or v2 → UNKNOWN (shape set mismatch)
    3. v2 moves a background / full-width shape → REGRESSION (critical)
    4. v2 position outside bounds → REGRESSION (critical)
    5. v2 mirrors a content shape that v1 didn't → likely IMPROVEMENT
    6. Golden baseline match → IMPROVEMENT or REGRESSION
    7. Everything else → UNKNOWN
    """

    def __init__(
        self,
        slide_width: int = 0,
        slide_height: int = 0,
        golden_slides: Optional[Dict[int, GoldenSlide]] = None,
    ):
        self.slide_width = slide_width
        self.slide_height = slide_height
        self.golden_slides = golden_slides or {}

    def classify(self, diff: ShapeDiff) -> ClassifiedDiff:
        """Classify a single ShapeDiff."""
        sw = self.slide_width

        # Rule: shape only in one version
        if diff.v1_only:
            return ClassifiedDiff(diff, DiffLabel.UNKNOWN,
                                  'Shape only in v1 output', 'major')
        if diff.v2_only:
            return ClassifiedDiff(diff, DiffLabel.UNKNOWN,
                                  'Shape only in v2 output', 'major')

        # Rule: no meaningful position change
        if (abs(diff.left_delta) <= POSITION_TOLERANCE_EMU
                and abs(diff.top_delta) <= POSITION_TOLERANCE_EMU
                and abs(diff.width_delta) <= POSITION_TOLERANCE_EMU
                and abs(diff.height_delta) <= POSITION_TOLERANCE_EMU
                and not diff.text_changed
                and not diff.flipH_changed):
            return ClassifiedDiff(diff, DiffLabel.NEUTRAL,
                                  'No meaningful change', 'info')

        # Rule: v2 moved a background shape
        if _is_background_shape(diff, sw) and abs(diff.left_delta) > POSITION_TOLERANCE_EMU:
            return ClassifiedDiff(diff, DiffLabel.REGRESSION,
                                  'v2 moved a full-width/background shape', 'critical')

        # Rule: v2 position out of bounds
        if _is_out_of_bounds(diff, sw):
            return ClassifiedDiff(diff, DiffLabel.REGRESSION,
                                  f'v2 position outside bounds '
                                  f'(left={diff.v2_left}, right={diff.v2_right})',
                                  'critical')

        # Rule: golden baseline comparison
        golden = self.golden_slides.get(diff.slide_number)
        if golden:
            gs = golden.shapes.get(diff.shape_id)
            if gs:
                v2_matches = abs(diff.v2_left - gs.expected_left) <= gs.tolerance
                v1_matches = abs(diff.v1_left - gs.expected_left) <= gs.tolerance
                if v2_matches and not v1_matches:
                    return ClassifiedDiff(diff, DiffLabel.IMPROVEMENT,
                                          'v2 matches golden baseline; v1 does not',
                                          'minor')
                if v1_matches and not v2_matches:
                    return ClassifiedDiff(diff, DiffLabel.REGRESSION,
                                          'v1 matches golden baseline; v2 diverged',
                                          'major')
                if v2_matches and v1_matches:
                    return ClassifiedDiff(diff, DiffLabel.NEUTRAL,
                                          'Both match golden baseline', 'info')

        # Rule: position changed — classify direction
        if abs(diff.left_delta) > POSITION_TOLERANCE_EMU:
            # Content-sized shapes that v2 mirrors are likely improvements
            if not _is_background_shape(diff, sw):
                return ClassifiedDiff(diff, DiffLabel.UNKNOWN,
                                      f'v2 repositioned shape '
                                      f'(Δleft={diff.left_delta} EMU)',
                                      'minor')

        # Rule: only text changed
        if diff.text_changed and not diff.position_changed:
            return ClassifiedDiff(diff, DiffLabel.NEUTRAL,
                                  'Text changed (translation expected)', 'info')

        # Rule: only flipH changed
        if diff.flipH_changed and not diff.position_changed:
            return ClassifiedDiff(diff, DiffLabel.UNKNOWN,
                                  'flipH toggled without position change', 'minor')

        return ClassifiedDiff(diff, DiffLabel.UNKNOWN,
                              'Unclassified change', 'info')

    def classify_slide(self, slide_diff: SlideDiff) -> ClassifiedSlide:
        """Classify all diffs on a slide, including overlap checks."""
        self.slide_width = slide_diff.slide_width or self.slide_width
        self.slide_height = slide_diff.slide_height or self.slide_height

        classified = ClassifiedSlide(
            slide_number=slide_diff.slide_number,
            layout_type=slide_diff.layout_type,
        )

        per_shape: List[ClassifiedDiff] = []
        for sd in slide_diff.shape_diffs:
            per_shape.append(self.classify(sd))

        # Post-pass: check for new overlaps introduced by v2
        moved_diffs = [c for c in per_shape
                       if c.shape_diff.position_changed
                       and not c.shape_diff.v1_only
                       and not c.shape_diff.v2_only]
        for i, a in enumerate(moved_diffs):
            for b in moved_diffs[i + 1:]:
                if _shapes_overlap(a.shape_diff, b.shape_diff):
                    # Upgrade to regression if not already
                    if a.label != DiffLabel.REGRESSION:
                        a.label = DiffLabel.REGRESSION
                        a.reason += '; overlaps with another moved shape'
                        a.severity = 'major'

        classified.diffs = per_shape
        return classified


# ─────────────────────────────────────────────────────────────────────────────
# Convenience
# ─────────────────────────────────────────────────────────────────────────────

def classify_all(
    slide_diffs: List[SlideDiff],
    golden_slides: Optional[Dict[int, GoldenSlide]] = None,
) -> List[ClassifiedSlide]:
    """Classify all slide diffs at once."""
    if not slide_diffs:
        return []
    sw = slide_diffs[0].slide_width
    sh = slide_diffs[0].slide_height
    clf = DiffClassifier(sw, sh, golden_slides)
    return [clf.classify_slide(sd) for sd in slide_diffs]
