"""
SlideArabi — Visual Regression Gate

Deterministic, pixel-level quality gate for the RTL conversion pipeline.
Renders original (EN) and converted (AR) slides to images, computes
structural metrics, and gates each slide as PASS / WARN / FAIL.

This module does NOT use LLM reasoning for the gate decision.
It does NOT auto-fix anything. If a slide fails, it is flagged
for investigation. The pipeline ships V2 output unchanged for
any slide that cannot prove visual integrity.

Architecture (5-model council consensus, 2026-03-20):
─────────────────────────────────────────────────────
1. Render EN slide → image
2. Render AR slide → image
3. Mirror EN image horizontally (RTL reference baseline)
4. Compute 5 structural metrics per slide:
   a. Ink Density Ratio (content density)
   b. Edge Density Ratio (structural complexity)
   c. Tile Occupancy Similarity (spatial distribution)
   d. Tile Anomaly Count (localized element deletion/overflow)
   e. Border Overflow Score (content clipping)
5. Apply deterministic thresholds → PASS / WARN / FAIL
6. Emit structured JSON report per slide

Dependencies: Pillow, numpy (both already in requirements.txt)
No scikit-image needed — saves ~185MB on Railway 512MB tier.
"""

from __future__ import annotations

import gc
import json
import logging
import os
import time
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
from PIL import Image, ImageFilter, ImageOps

logger = logging.getLogger(__name__)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# CONSTANTS
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

# Analysis resolution — downsample to this max dimension for metrics
# Keeps memory < 5MB per slide pair on Railway 512MB
ANALYSIS_MAX_DIM = 800

# Background detection threshold (grayscale 0-255)
# Pixels >= this are considered background (near-white)
BG_THRESHOLD = 240

# Tile grid dimensions for spatial analysis
TILE_ROWS = 4
TILE_COLS = 6

# Edge detection threshold (gradient magnitude)
EDGE_THRESHOLD = 30

# Border band width in pixels (at analysis resolution) for overflow detection
BORDER_BAND_PX = 8

# Minimum ink ratio for a slide to be analyzable
# (below this, slide is essentially blank — skip metrics)
MIN_INK_FOR_ANALYSIS = 0.005


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# THRESHOLDS (calibrated from 5-model council consensus)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

# Ink Density Ratio: AR ink / EN ink
# Arabic text is typically 10-20% denser than English.
# A large DROP means elements vanished. A large SPIKE means overlap.
INK_RATIO_FAIL_LOW = 0.55       # <55% of EN density → content vanished
INK_RATIO_FAIL_HIGH = 1.60      # >160% → severe overlap/clumping
INK_RATIO_WARN_LOW = 0.70       # <70% → possible missing content
INK_RATIO_WARN_HIGH = 1.40      # >140% → possible overlap

# Edge Density Ratio: AR edges / EN edges
# Missing shapes/images = missing edges. Arabic text has more edges than Latin.
EDGE_RATIO_FAIL_LOW = 0.50      # <50% edge density = major structural loss
EDGE_RATIO_WARN_LOW = 0.65      # <65% = warning

# Tile Occupancy Similarity (Jaccard of occupied tiles, EN mirrored)
# Perfect RTL conversion ~0.4-0.7 (never 1.0 due to text shape differences
# and asymmetric content). Calibrated against V2/V3 481-slide corpus.
TILE_JACCARD_FAIL = 0.15        # <15% overlap = layout destroyed
TILE_JACCARD_WARN = 0.25        # <25% = concerning

# Tile Anomaly: tiles where EN has content but AR mirror is empty
# (catches localized element deletion like Venn diagrams vanishing)
# NOTE: RTL mirroring means asymmetric content produces some anomalies
# naturally. Threshold set at 4+ to avoid false positives on slides
# with asymmetric layouts. Calibrated against 481-slide V2/V3 corpus.
TILE_ANOMALY_FAIL = 4           # 4+ tiles with content→empty = element deletion
TILE_ANOMALY_WARN = 3           # 3 tiles = warning

# Border Overflow: elevated ink near slide edges in AR but not in EN
BORDER_OVERFLOW_FAIL = 0.15     # >15% of border band has new foreground
BORDER_OVERFLOW_WARN = 0.08     # >8% = warning


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# DATA MODELS
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


class GateStatus(str, Enum):
    PASS = "PASS"
    WARN = "WARN"
    FAIL = "FAIL"
    RENDER_ERROR = "RENDER_ERROR"
    SKIP = "SKIP"  # Slide too sparse to analyze


@dataclass
class SlideMetrics:
    """Raw computed metrics for one slide."""
    slide_number: int
    ink_ratio_en: float = 0.0       # EN ink density (0-1)
    ink_ratio_ar: float = 0.0       # AR ink density (0-1)
    ink_density_ratio: float = 0.0  # AR/EN ratio
    edge_ratio_en: float = 0.0     # EN edge density
    edge_ratio_ar: float = 0.0     # AR edge density
    edge_density_ratio: float = 0.0 # AR/EN edge ratio
    tile_jaccard: float = 0.0       # Jaccard similarity of occupied tiles
    tile_anomaly_count: int = 0     # Tiles with content→empty
    tile_overflow_count: int = 0    # Tiles with empty→content (overflow)
    border_overflow_score: float = 0.0  # Border band new-foreground ratio
    analysis_resolution: Tuple[int, int] = (0, 0)

    def to_dict(self) -> dict:
        return {
            "slide_number": self.slide_number,
            "ink_ratio_en": round(self.ink_ratio_en, 4),
            "ink_ratio_ar": round(self.ink_ratio_ar, 4),
            "ink_density_ratio": round(self.ink_density_ratio, 4),
            "edge_ratio_en": round(self.edge_ratio_en, 4),
            "edge_ratio_ar": round(self.edge_ratio_ar, 4),
            "edge_density_ratio": round(self.edge_density_ratio, 4),
            "tile_jaccard": round(self.tile_jaccard, 4),
            "tile_anomaly_count": self.tile_anomaly_count,
            "tile_overflow_count": self.tile_overflow_count,
            "border_overflow_score": round(self.border_overflow_score, 4),
            "analysis_resolution": list(self.analysis_resolution),
        }


@dataclass
class GateRule:
    """A single gate rule that was triggered."""
    rule_id: str       # e.g. "INK_DENSITY_LOW"
    metric: str        # e.g. "ink_density_ratio"
    value: float       # the actual measured value
    threshold: float   # the threshold it violated
    severity: str      # "FAIL" or "WARN"
    description: str   # human-readable explanation

    def to_dict(self) -> dict:
        return {
            "rule_id": self.rule_id,
            "metric": self.metric,
            "value": round(self.value, 4),
            "threshold": round(self.threshold, 4),
            "severity": self.severity,
            "description": self.description,
        }


@dataclass
class SlideGateResult:
    """Gate decision for one slide."""
    slide_number: int
    status: GateStatus = GateStatus.PASS
    metrics: Optional[SlideMetrics] = None
    triggered_rules: List[GateRule] = field(default_factory=list)
    error: Optional[str] = None

    @property
    def fail_count(self) -> int:
        return sum(1 for r in self.triggered_rules if r.severity == "FAIL")

    @property
    def warn_count(self) -> int:
        return sum(1 for r in self.triggered_rules if r.severity == "WARN")

    def to_dict(self) -> dict:
        return {
            "slide_number": self.slide_number,
            "status": self.status.value,
            "fail_count": self.fail_count,
            "warn_count": self.warn_count,
            "triggered_rules": [r.to_dict() for r in self.triggered_rules],
            "metrics": self.metrics.to_dict() if self.metrics else None,
            "error": self.error,
        }


@dataclass
class GateReport:
    """Deck-level gate report."""
    deck_name: str = ""
    total_slides: int = 0
    slides_analyzed: int = 0
    pass_count: int = 0
    warn_count: int = 0
    fail_count: int = 0
    skip_count: int = 0
    error_count: int = 0
    slide_results: List[SlideGateResult] = field(default_factory=list)
    duration_ms: float = 0.0

    @property
    def pass_rate(self) -> float:
        if self.slides_analyzed == 0:
            return 0.0
        return (self.pass_count / self.slides_analyzed) * 100

    @property
    def fail_rate(self) -> float:
        if self.slides_analyzed == 0:
            return 0.0
        return (self.fail_count / self.slides_analyzed) * 100

    def summary(self) -> str:
        return (
            f"Visual Gate: {self.slides_analyzed} slides analyzed | "
            f"PASS={self.pass_count} WARN={self.warn_count} "
            f"FAIL={self.fail_count} SKIP={self.skip_count} | "
            f"Pass rate: {self.pass_rate:.1f}% | "
            f"Duration: {self.duration_ms:.0f}ms"
        )

    def to_dict(self) -> dict:
        return {
            "deck_name": self.deck_name,
            "total_slides": self.total_slides,
            "slides_analyzed": self.slides_analyzed,
            "pass_count": self.pass_count,
            "warn_count": self.warn_count,
            "fail_count": self.fail_count,
            "skip_count": self.skip_count,
            "error_count": self.error_count,
            "pass_rate": round(self.pass_rate, 1),
            "fail_rate": round(self.fail_rate, 1),
            "duration_ms": round(self.duration_ms, 1),
            "slide_results": [r.to_dict() for r in self.slide_results],
        }


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# CORE METRIC COMPUTATION
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


def _load_and_prepare(image_path: str) -> np.ndarray:
    """Load image, convert to grayscale, resize to analysis resolution.
    
    Returns numpy array (H, W) of uint8 grayscale values.
    Memory: ~1MB for 800x600 image.
    """
    img = Image.open(image_path)
    
    # Resize to analysis resolution (preserve aspect ratio)
    w, h = img.size
    scale = min(ANALYSIS_MAX_DIM / w, ANALYSIS_MAX_DIM / h, 1.0)
    if scale < 1.0:
        new_w = max(1, int(w * scale))
        new_h = max(1, int(h * scale))
        img = img.resize((new_w, new_h), Image.LANCZOS)
    
    # Convert to grayscale
    gray = img.convert("L")
    arr = np.array(gray, dtype=np.uint8)
    
    # Explicit cleanup
    img.close()
    del img, gray
    
    return arr


def _compute_ink_ratio(gray: np.ndarray) -> float:
    """Compute ink density: fraction of non-background pixels."""
    total = gray.size
    if total == 0:
        return 0.0
    ink_pixels = np.sum(gray < BG_THRESHOLD)
    return float(ink_pixels / total)


def _compute_edge_density(gray: np.ndarray) -> float:
    """Compute edge density using simple gradient magnitude.
    
    Uses numpy gradient (Sobel-lite) — no scipy needed.
    """
    # Compute x and y gradients
    gy = np.abs(np.diff(gray.astype(np.int16), axis=0))
    gx = np.abs(np.diff(gray.astype(np.int16), axis=1))
    
    # Pad to original size (gradient reduces dim by 1)
    gy_padded = np.zeros_like(gray, dtype=np.int16)
    gy_padded[:-1, :] = gy
    gx_padded = np.zeros_like(gray, dtype=np.int16)
    gx_padded[:, :-1] = gx
    
    # Gradient magnitude
    magnitude = np.sqrt(gy_padded.astype(np.float32) ** 2 + 
                        gx_padded.astype(np.float32) ** 2)
    
    total = magnitude.size
    if total == 0:
        return 0.0
    edge_pixels = np.sum(magnitude > EDGE_THRESHOLD)
    
    # Cleanup
    del gy, gx, gy_padded, gx_padded, magnitude
    
    return float(edge_pixels / total)


def _compute_tile_metrics(
    en_gray: np.ndarray,
    ar_gray: np.ndarray,
) -> Tuple[float, int, int]:
    """Compute tile-based spatial metrics.
    
    Returns:
        (jaccard, anomaly_count, overflow_count)
        
    jaccard: Jaccard similarity of occupied tiles between
             horizontally-mirrored EN and AR.
    anomaly_count: Number of tiles where EN has content but
                   AR mirror is empty (element deletion).
    overflow_count: Number of tiles where EN is empty but
                    AR mirror has content (text overflow).
    """
    h_en, w_en = en_gray.shape
    h_ar, w_ar = ar_gray.shape
    
    tile_h_en = max(1, h_en // TILE_ROWS)
    tile_w_en = max(1, w_en // TILE_COLS)
    tile_h_ar = max(1, h_ar // TILE_ROWS)
    tile_w_ar = max(1, w_ar // TILE_COLS)
    
    # Minimum ink fraction for a tile to be "occupied"
    TILE_OCCUPIED_THRESHOLD = 0.02
    # Threshold for "content-rich" tile (used in anomaly detection)
    TILE_CONTENT_RICH = 0.08
    # Threshold for "nearly empty" tile
    TILE_NEARLY_EMPTY = 0.02
    
    en_occupied = set()
    ar_occupied = set()
    en_content_rich = {}  # (row, col) -> ink_ratio
    ar_tile_ink = {}      # (row, col) -> ink_ratio
    
    for r in range(TILE_ROWS):
        for c in range(TILE_COLS):
            # EN tile
            en_r_start = r * tile_h_en
            en_r_end = min((r + 1) * tile_h_en, h_en)
            en_c_start = c * tile_w_en
            en_c_end = min((c + 1) * tile_w_en, w_en)
            en_tile = en_gray[en_r_start:en_r_end, en_c_start:en_c_end]
            
            en_ink = float(np.sum(en_tile < BG_THRESHOLD)) / max(en_tile.size, 1)
            if en_ink > TILE_OCCUPIED_THRESHOLD:
                # RTL mirror: column (TILE_COLS - 1 - c)
                en_occupied.add((r, TILE_COLS - 1 - c))
            if en_ink > TILE_CONTENT_RICH:
                en_content_rich[(r, TILE_COLS - 1 - c)] = en_ink
            
            # AR tile
            ar_r_start = r * tile_h_ar
            ar_r_end = min((r + 1) * tile_h_ar, h_ar)
            ar_c_start = c * tile_w_ar
            ar_c_end = min((c + 1) * tile_w_ar, w_ar)
            ar_tile = ar_gray[ar_r_start:ar_r_end, ar_c_start:ar_c_end]
            
            ar_ink = float(np.sum(ar_tile < BG_THRESHOLD)) / max(ar_tile.size, 1)
            if ar_ink > TILE_OCCUPIED_THRESHOLD:
                ar_occupied.add((r, c))
            ar_tile_ink[(r, c)] = ar_ink
    
    # Jaccard similarity
    intersection = len(en_occupied & ar_occupied)
    union = len(en_occupied | ar_occupied)
    jaccard = intersection / max(union, 1)
    
    # Anomaly count: EN content-rich tile → AR mirror is nearly empty
    anomaly_count = 0
    for (r, c_mirrored), en_ink in en_content_rich.items():
        ar_ink = ar_tile_ink.get((r, c_mirrored), 0.0)
        if ar_ink < TILE_NEARLY_EMPTY and en_ink > TILE_CONTENT_RICH:
            anomaly_count += 1
    
    # Overflow count: EN tile empty → AR tile has unexpected content
    overflow_count = 0
    en_empty_mirrored = set()
    for r in range(TILE_ROWS):
        for c in range(TILE_COLS):
            mirrored_c = TILE_COLS - 1 - c
            # Check if EN original tile at (r, c) is empty
            en_r_start = r * tile_h_en
            en_r_end = min((r + 1) * tile_h_en, h_en)
            en_c_start = c * tile_w_en
            en_c_end = min((c + 1) * tile_w_en, w_en)
            en_tile = en_gray[en_r_start:en_r_end, en_c_start:en_c_end]
            en_ink = float(np.sum(en_tile < BG_THRESHOLD)) / max(en_tile.size, 1)
            if en_ink < TILE_NEARLY_EMPTY:
                en_empty_mirrored.add((r, mirrored_c))
    
    for (r, c_mirrored) in en_empty_mirrored:
        ar_ink = ar_tile_ink.get((r, c_mirrored), 0.0)
        if ar_ink > TILE_CONTENT_RICH:
            overflow_count += 1
    
    return jaccard, anomaly_count, overflow_count


def _compute_border_overflow(
    en_gray: np.ndarray,
    ar_gray: np.ndarray,
) -> float:
    """Compute border overflow score.
    
    Measures new foreground content near slide edges in AR
    that wasn't present in EN. Catches text overflow / clipping.
    """
    h, w = ar_gray.shape
    band = min(BORDER_BAND_PX, h // 10, w // 10)
    if band < 2:
        return 0.0
    
    # Build border masks
    en_fg = (en_gray < BG_THRESHOLD)
    ar_fg = (ar_gray < BG_THRESHOLD)
    
    # Resize EN to AR dimensions if different
    if en_fg.shape != ar_fg.shape:
        en_img = Image.fromarray(en_gray)
        en_img = en_img.resize((w, h), Image.LANCZOS)
        en_resized = np.array(en_img, dtype=np.uint8)
        en_fg = (en_resized < BG_THRESHOLD)
        # Also mirror EN for RTL comparison
        en_fg = np.fliplr(en_fg)
        del en_img, en_resized
    else:
        en_fg = np.fliplr(en_fg)
    
    # Border regions
    border_mask = np.zeros_like(ar_fg, dtype=bool)
    border_mask[:band, :] = True       # top
    border_mask[-band:, :] = True      # bottom
    border_mask[:, :band] = True       # left
    border_mask[:, -band:] = True      # right
    
    # New foreground in AR border that wasn't in EN border
    new_border_fg = ar_fg & border_mask & ~(en_fg & border_mask)
    
    border_total = np.sum(border_mask)
    if border_total == 0:
        return 0.0
    
    score = float(np.sum(new_border_fg)) / border_total
    
    del en_fg, ar_fg, border_mask, new_border_fg
    return score


def compute_slide_metrics(
    en_image_path: str,
    ar_image_path: str,
    slide_number: int,
) -> SlideMetrics:
    """Compute all visual regression metrics for one slide pair.
    
    Memory budget: ~10MB peak per call. Processes one slide at a time.
    """
    metrics = SlideMetrics(slide_number=slide_number)
    
    # Load and prepare images
    en_gray = _load_and_prepare(en_image_path)
    ar_gray = _load_and_prepare(ar_image_path)
    
    metrics.analysis_resolution = (ar_gray.shape[1], ar_gray.shape[0])
    
    # 1. Ink density
    metrics.ink_ratio_en = _compute_ink_ratio(en_gray)
    metrics.ink_ratio_ar = _compute_ink_ratio(ar_gray)
    
    if metrics.ink_ratio_en > MIN_INK_FOR_ANALYSIS:
        metrics.ink_density_ratio = metrics.ink_ratio_ar / metrics.ink_ratio_en
    else:
        # EN slide is essentially blank — ratio undefined
        metrics.ink_density_ratio = 1.0 if metrics.ink_ratio_ar < MIN_INK_FOR_ANALYSIS else 2.0
    
    # 2. Edge density
    metrics.edge_ratio_en = _compute_edge_density(en_gray)
    metrics.edge_ratio_ar = _compute_edge_density(ar_gray)
    
    if metrics.edge_ratio_en > 0.001:
        metrics.edge_density_ratio = metrics.edge_ratio_ar / metrics.edge_ratio_en
    else:
        metrics.edge_density_ratio = 1.0
    
    # 3+4. Tile metrics
    jaccard, anomaly, overflow = _compute_tile_metrics(en_gray, ar_gray)
    metrics.tile_jaccard = jaccard
    metrics.tile_anomaly_count = anomaly
    metrics.tile_overflow_count = overflow
    
    # 5. Border overflow
    metrics.border_overflow_score = _compute_border_overflow(en_gray, ar_gray)
    
    # Cleanup
    del en_gray, ar_gray
    gc.collect()
    
    return metrics


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# GATE DECISION ENGINE
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


def apply_gate(metrics: SlideMetrics) -> SlideGateResult:
    """Apply deterministic gate thresholds to computed metrics.
    
    Rules are evaluated independently. Any FAIL rule = slide FAIL.
    No FAIL but any WARN rule = slide WARN. Else PASS.
    """
    result = SlideGateResult(
        slide_number=metrics.slide_number,
        metrics=metrics,
    )
    
    # Skip gate for blank/near-blank slides
    if metrics.ink_ratio_en < MIN_INK_FOR_ANALYSIS:
        result.status = GateStatus.SKIP
        return result
    
    rules: List[GateRule] = []
    
    # ── Rule 1: Ink Density Ratio (low = content vanished) ──
    if metrics.ink_density_ratio < INK_RATIO_FAIL_LOW:
        rules.append(GateRule(
            rule_id="INK_DENSITY_LOW",
            metric="ink_density_ratio",
            value=metrics.ink_density_ratio,
            threshold=INK_RATIO_FAIL_LOW,
            severity="FAIL",
            description=(
                f"Content density dropped to {metrics.ink_density_ratio:.1%} of original "
                f"(threshold: {INK_RATIO_FAIL_LOW:.0%}). "
                "Elements likely vanished during conversion."
            ),
        ))
    elif metrics.ink_density_ratio < INK_RATIO_WARN_LOW:
        rules.append(GateRule(
            rule_id="INK_DENSITY_LOW_WARN",
            metric="ink_density_ratio",
            value=metrics.ink_density_ratio,
            threshold=INK_RATIO_WARN_LOW,
            severity="WARN",
            description=(
                f"Content density at {metrics.ink_density_ratio:.1%} of original "
                f"(warning below {INK_RATIO_WARN_LOW:.0%})."
            ),
        ))
    
    # ── Rule 2: Ink Density Ratio (high = overlap/clumping) ──
    if metrics.ink_density_ratio > INK_RATIO_FAIL_HIGH:
        rules.append(GateRule(
            rule_id="INK_DENSITY_HIGH",
            metric="ink_density_ratio",
            value=metrics.ink_density_ratio,
            threshold=INK_RATIO_FAIL_HIGH,
            severity="FAIL",
            description=(
                f"Content density at {metrics.ink_density_ratio:.1%} of original "
                f"(threshold: {INK_RATIO_FAIL_HIGH:.0%}). "
                "Elements likely overlapping or clumping."
            ),
        ))
    elif metrics.ink_density_ratio > INK_RATIO_WARN_HIGH:
        rules.append(GateRule(
            rule_id="INK_DENSITY_HIGH_WARN",
            metric="ink_density_ratio",
            value=metrics.ink_density_ratio,
            threshold=INK_RATIO_WARN_HIGH,
            severity="WARN",
            description=(
                f"Content density elevated at {metrics.ink_density_ratio:.1%} of original "
                f"(warning above {INK_RATIO_WARN_HIGH:.0%})."
            ),
        ))
    
    # ── Rule 3: Edge Density Ratio (low = structural loss) ──
    if metrics.edge_density_ratio < EDGE_RATIO_FAIL_LOW:
        rules.append(GateRule(
            rule_id="EDGE_DENSITY_LOW",
            metric="edge_density_ratio",
            value=metrics.edge_density_ratio,
            threshold=EDGE_RATIO_FAIL_LOW,
            severity="FAIL",
            description=(
                f"Edge density at {metrics.edge_density_ratio:.1%} of original "
                f"(threshold: {EDGE_RATIO_FAIL_LOW:.0%}). "
                "Shapes, images, or graphical elements likely missing."
            ),
        ))
    elif metrics.edge_density_ratio < EDGE_RATIO_WARN_LOW:
        rules.append(GateRule(
            rule_id="EDGE_DENSITY_LOW_WARN",
            metric="edge_density_ratio",
            value=metrics.edge_density_ratio,
            threshold=EDGE_RATIO_WARN_LOW,
            severity="WARN",
            description=(
                f"Edge density at {metrics.edge_density_ratio:.1%} of original "
                f"(warning below {EDGE_RATIO_WARN_LOW:.0%})."
            ),
        ))
    
    # ── Rule 4: Tile Occupancy Jaccard (low = layout destroyed) ──
    if metrics.tile_jaccard < TILE_JACCARD_FAIL:
        rules.append(GateRule(
            rule_id="TILE_LAYOUT_DESTROYED",
            metric="tile_jaccard",
            value=metrics.tile_jaccard,
            threshold=TILE_JACCARD_FAIL,
            severity="FAIL",
            description=(
                f"Spatial layout similarity is {metrics.tile_jaccard:.2f} "
                f"(threshold: {TILE_JACCARD_FAIL:.2f}). "
                "Layout appears fundamentally broken."
            ),
        ))
    elif metrics.tile_jaccard < TILE_JACCARD_WARN:
        rules.append(GateRule(
            rule_id="TILE_LAYOUT_WARN",
            metric="tile_jaccard",
            value=metrics.tile_jaccard,
            threshold=TILE_JACCARD_WARN,
            severity="WARN",
            description=(
                f"Spatial layout similarity is {metrics.tile_jaccard:.2f} "
                f"(warning below {TILE_JACCARD_WARN:.2f})."
            ),
        ))
    
    # ── Rule 5: Tile Anomaly (content→empty = element deletion) ──
    if metrics.tile_anomaly_count >= TILE_ANOMALY_FAIL:
        rules.append(GateRule(
            rule_id="TILE_ELEMENT_DELETION",
            metric="tile_anomaly_count",
            value=metrics.tile_anomaly_count,
            threshold=TILE_ANOMALY_FAIL,
            severity="FAIL",
            description=(
                f"{metrics.tile_anomaly_count} tile regions have content in original "
                f"but are empty in converted (threshold: {TILE_ANOMALY_FAIL}). "
                "Elements were likely deleted."
            ),
        ))
    elif metrics.tile_anomaly_count >= TILE_ANOMALY_WARN:
        rules.append(GateRule(
            rule_id="TILE_ELEMENT_DELETION_WARN",
            metric="tile_anomaly_count",
            value=metrics.tile_anomaly_count,
            threshold=TILE_ANOMALY_WARN,
            severity="WARN",
            description=(
                f"{metrics.tile_anomaly_count} tile regions have content→empty "
                f"(warning at {TILE_ANOMALY_WARN})."
            ),
        ))
    
    # ── Rule 6: Border Overflow (content clipping) ──
    if metrics.border_overflow_score > BORDER_OVERFLOW_FAIL:
        rules.append(GateRule(
            rule_id="BORDER_OVERFLOW",
            metric="border_overflow_score",
            value=metrics.border_overflow_score,
            threshold=BORDER_OVERFLOW_FAIL,
            severity="FAIL",
            description=(
                f"Border overflow score is {metrics.border_overflow_score:.1%} "
                f"(threshold: {BORDER_OVERFLOW_FAIL:.0%}). "
                "Content is clipping at slide edges."
            ),
        ))
    elif metrics.border_overflow_score > BORDER_OVERFLOW_WARN:
        rules.append(GateRule(
            rule_id="BORDER_OVERFLOW_WARN",
            metric="border_overflow_score",
            value=metrics.border_overflow_score,
            threshold=BORDER_OVERFLOW_WARN,
            severity="WARN",
            description=(
                f"Border overflow score is {metrics.border_overflow_score:.1%} "
                f"(warning above {BORDER_OVERFLOW_WARN:.0%})."
            ),
        ))
    
    # ── Final Decision ──
    result.triggered_rules = rules
    
    has_fail = any(r.severity == "FAIL" for r in rules)
    has_warn = any(r.severity == "WARN" for r in rules)
    
    if has_fail:
        result.status = GateStatus.FAIL
    elif has_warn:
        result.status = GateStatus.WARN
    else:
        result.status = GateStatus.PASS
    
    return result


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# ORCHESTRATOR
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


class VisualRegressionGate:
    """Orchestrates the visual regression gate across all slides.
    
    Processes slides one at a time to stay within 512MB RAM.
    Uses the existing SlideRenderer from visual_qa.py.
    """
    
    def __init__(self, render_dpi: int = 150):
        self.render_dpi = render_dpi
    
    def run(
        self,
        en_images: List[str],
        ar_images: List[str],
        deck_name: str = "",
        issue_log_path: Optional[str] = None,
    ) -> GateReport:
        """Run visual regression gate on pre-rendered slide images.
        
        Args:
            en_images: List of paths to rendered EN slide images (sorted)
            ar_images: List of paths to rendered AR slide images (sorted)
            deck_name: Name of the deck for logging
            issue_log_path: Path to write JSONL issue log
            
        Returns:
            GateReport with per-slide results
        """
        start_time = time.monotonic()
        
        total_slides = min(len(en_images), len(ar_images))
        report = GateReport(
            deck_name=deck_name,
            total_slides=total_slides,
        )
        
        logger.info(
            "Visual Gate: Starting analysis of %d slides for '%s'",
            total_slides, deck_name,
        )
        
        for slide_idx in range(total_slides):
            slide_num = slide_idx + 1
            
            try:
                # Compute metrics
                metrics = compute_slide_metrics(
                    en_image_path=en_images[slide_idx],
                    ar_image_path=ar_images[slide_idx],
                    slide_number=slide_num,
                )
                
                # Apply gate
                gate_result = apply_gate(metrics)
                
            except Exception as e:
                logger.warning(
                    "Visual Gate: Error analyzing slide %d: %s",
                    slide_num, e,
                )
                gate_result = SlideGateResult(
                    slide_number=slide_num,
                    status=GateStatus.RENDER_ERROR,
                    error=str(e),
                )
            
            report.slide_results.append(gate_result)
            
            # Update counters
            if gate_result.status == GateStatus.PASS:
                report.pass_count += 1
            elif gate_result.status == GateStatus.WARN:
                report.warn_count += 1
            elif gate_result.status == GateStatus.FAIL:
                report.fail_count += 1
            elif gate_result.status == GateStatus.SKIP:
                report.skip_count += 1
            else:
                report.error_count += 1
            
            report.slides_analyzed += 1
            
            # Log individual failures
            if gate_result.status == GateStatus.FAIL:
                rule_ids = [r.rule_id for r in gate_result.triggered_rules 
                           if r.severity == "FAIL"]
                logger.warning(
                    "Visual Gate: Slide %d FAIL — rules: %s",
                    slide_num, ", ".join(rule_ids),
                )
        
        report.duration_ms = (time.monotonic() - start_time) * 1000
        
        # Write issue log
        if issue_log_path:
            self._write_issue_log(report, issue_log_path)
        
        logger.info(report.summary())
        return report
    
    def _write_issue_log(self, report: GateReport, log_path: str) -> None:
        """Write structured JSONL issue log for all non-PASS slides."""
        try:
            Path(log_path).parent.mkdir(parents=True, exist_ok=True)
            with open(log_path, "w", encoding="utf-8") as f:
                for result in report.slide_results:
                    if result.status in (GateStatus.FAIL, GateStatus.WARN):
                        entry = {
                            "timestamp": time.strftime("%Y-%m-%dT%H:%M:%SZ"),
                            "deck_name": report.deck_name,
                            "slide_number": result.slide_number,
                            "gate_status": result.status.value,
                            "fail_count": result.fail_count,
                            "warn_count": result.warn_count,
                            "triggered_rules": [
                                r.to_dict() for r in result.triggered_rules
                            ],
                            "metrics": result.metrics.to_dict() 
                                if result.metrics else None,
                        }
                        f.write(json.dumps(entry, ensure_ascii=False) + "\n")
        except Exception as e:
            logger.warning("Visual Gate: Failed to write issue log: %s", e)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# CONVENIENCE FUNCTION
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


def run_visual_gate(
    original_pptx: str,
    converted_pptx: str,
    deck_name: str = "",
    issue_log_path: Optional[str] = None,
    render_dpi: int = 150,
) -> GateReport:
    """Run the visual regression gate end-to-end.
    
    Renders both PPTX files and computes visual regression metrics.
    This is the main entry point for pipeline integration.
    
    Args:
        original_pptx: Path to original (English) PPTX
        converted_pptx: Path to converted (Arabic) PPTX
        deck_name: Name for logging
        issue_log_path: Path for JSONL issue log
        render_dpi: DPI for slide rendering
        
    Returns:
        GateReport with per-slide PASS/WARN/FAIL decisions
    """
    import tempfile
    import shutil
    
    # Import the existing renderer
    try:
        from slidearabi.visual_qa import SlideRenderer
    except ImportError:
        from visual_qa import SlideRenderer
    
    renderer = SlideRenderer(dpi=render_dpi)
    gate = VisualRegressionGate(render_dpi=render_dpi)
    
    work_dir = tempfile.mkdtemp(prefix="visual_gate_")
    en_dir = os.path.join(work_dir, "en_renders")
    ar_dir = os.path.join(work_dir, "ar_renders")
    
    try:
        # Render both presentations
        logger.info("Visual Gate: Rendering original (EN) slides...")
        en_images = renderer.render_to_images(original_pptx, en_dir)
        
        logger.info("Visual Gate: Rendering converted (AR) slides...")
        ar_images = renderer.render_to_images(converted_pptx, ar_dir)
        
        # Run gate
        report = gate.run(
            en_images=en_images,
            ar_images=ar_images,
            deck_name=deck_name or Path(converted_pptx).stem,
            issue_log_path=issue_log_path,
        )
        
        return report
        
    finally:
        # Cleanup temp files
        shutil.rmtree(work_dir, ignore_errors=True)
        gc.collect()
