"""
audit_logger.py — Structured transformation audit logging for the SlideArabi RTL engine.

Records every transformation applied to every shape, producing a full audit trail
for debugging regressions and root cause analysis.

Usage example
-------------
    from slidearabi.audit_logger import AuditLogger

    audit = AuditLogger()
    audit.deck_name = "my_deck.pptx"
    audit.slide_count = 12
    audit.shape_count = 88

    # Inside a transform:
    audit.log_transform(
        slide_idx=1, shape_id=5, shape_name="Title 1", shape_type="text_box",
        transform_type="mirror_x",
        before_state={"x": 457200, "y": 274638, "cx": 8229600, "cy": 1143000},
        after_state ={"x": 304200, "y": 274638, "cx": 8229600, "cy": 1143000},
    )

    audit.to_json("/tmp/audit.json")
    audit.to_markdown("/tmp/audit.md")
    audit.print_summary()
"""

from __future__ import annotations

import json
import time
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional


# ─────────────────────────────────────────────────────────────────────────────
# Data model
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class TransformEntry:
    """Single transformation applied to (or exemption/skip recorded for) a shape."""

    slide_idx: int
    """1-based slide number."""

    shape_id: int
    """python-pptx shape_id or a synthetic sequential int for shapes without one."""

    shape_name: str
    """Human-readable shape name (shape.name)."""

    shape_type: str
    """Canonical shape category: 'text_box', 'image', 'chart', 'table', 'group',
    'connector', 'placeholder', 'other'."""

    transform_type: str
    """Kind of operation: 'mirror_x', 'panel_swap', 'text_rtl', 'axis_reversal',
    'alignment', 'table_rtl', 'directional_flip', 'exempt', 'skip',
    'position_inherit', 'cover_anchor', 'timeline_swap', 'logo_row_reverse',
    'bidi_direction', 'wrap_fix', 'slide_num_badge', 'autofit', etc."""

    before_state: Dict[str, Any]
    """Key shape properties captured immediately *before* the transform."""

    after_state: Dict[str, Any]
    """Key shape properties captured immediately *after* the transform."""

    notes: str = ""
    """Optional free-form annotation (e.g. 'map exemption', 'cover detection',
    'size-divergence guard triggered')."""

    timestamp: float = field(default_factory=time.time)
    """Unix timestamp of this entry (set automatically)."""


# ─────────────────────────────────────────────────────────────────────────────
# AuditLogger
# ─────────────────────────────────────────────────────────────────────────────

class AuditLogger:
    """
    Collects TransformEntry records for every RTL transformation and produces
    JSON and Markdown audit reports.

    Lifecycle
    ---------
    1. Instantiate once per pipeline run.
    2. Pass `audit` into (or construct inside) each transformer phase.
    3. Call log_transform / log_exemption / log_skip at every decision point.
    4. At the end of the run, call to_json() and/or to_markdown() to persist.
    5. Call print_summary() for a quick console overview.
    """

    def __init__(self) -> None:
        self.entries: List[TransformEntry] = []
        self.start_time: float = time.time()
        self.deck_name: str = ""
        self.slide_count: int = 0
        self.shape_count: int = 0

    # ─────────────────────────────────────────────────────────────────────────
    # Logging helpers
    # ─────────────────────────────────────────────────────────────────────────

    def log_transform(
        self,
        slide_idx: int,
        shape_id: int,
        shape_name: str,
        shape_type: str,
        transform_type: str,
        before_state: Dict[str, Any],
        after_state: Dict[str, Any],
        notes: str = "",
    ) -> None:
        """Record a single transformation applied to a shape.

        Args:
            slide_idx:      1-based slide number.
            shape_id:       Shape identifier (shape.shape_id or synthetic int).
            shape_name:     shape.name value.
            shape_type:     Canonical category: 'text_box', 'image', 'chart', …
            transform_type: Operation label: 'mirror_x', 'text_rtl', …
            before_state:   Dict of key properties *before* the transform.
            after_state:    Dict of key properties *after* the transform.
            notes:          Optional annotation string.
        """
        self.entries.append(TransformEntry(
            slide_idx=slide_idx,
            shape_id=shape_id,
            shape_name=shape_name,
            shape_type=shape_type,
            transform_type=transform_type,
            before_state=before_state,
            after_state=after_state,
            notes=notes,
        ))

    def log_exemption(
        self,
        slide_idx: int,
        shape_id: int,
        shape_name: str,
        shape_type: str,
        reason: str,
    ) -> None:
        """Record that a shape was *exempted* from a position transform.

        Args:
            slide_idx:  1-based slide number.
            shape_id:   Shape identifier.
            shape_name: shape.name value.
            shape_type: Canonical category.
            reason:     Why the shape was exempted (e.g. 'map overlay', 'logo').
        """
        self.entries.append(TransformEntry(
            slide_idx=slide_idx,
            shape_id=shape_id,
            shape_name=shape_name,
            shape_type=shape_type,
            transform_type="exempt",
            before_state={},
            after_state={},
            notes=reason,
        ))

    def log_skip(
        self,
        slide_idx: int,
        shape_id: int,
        shape_name: str,
        reason: str,
    ) -> None:
        """Record that a shape was *skipped entirely* (not examined at all).

        Args:
            slide_idx:  1-based slide number.
            shape_id:   Shape identifier.
            shape_name: shape.name value.
            reason:     Why the shape was skipped (e.g. 'no position data',
                        'full-width background', 'placeholder no layout match').
        """
        self.entries.append(TransformEntry(
            slide_idx=slide_idx,
            shape_id=shape_id,
            shape_name=shape_name,
            shape_type="unknown",
            transform_type="skip",
            before_state={},
            after_state={},
            notes=reason,
        ))

    # ─────────────────────────────────────────────────────────────────────────
    # Statistics
    # ─────────────────────────────────────────────────────────────────────────

    def summary(self) -> Dict[str, Any]:
        """Return summary statistics for the full audit log.

        Returns a dict with:
        - total_transforms  : int — all entries including exempt/skip
        - by_type           : {transform_type: count}
        - by_slide          : {slide_idx: count}
        - exempt_count      : int
        - skip_count        : int
        - shape_type_counts : {shape_type: count}
        - elapsed_seconds   : float
        - deck_name         : str
        - slide_count       : int
        - shape_count       : int
        """
        by_type: Dict[str, int] = defaultdict(int)
        by_slide: Dict[int, int] = defaultdict(int)
        shape_type_counts: Dict[str, int] = defaultdict(int)

        for entry in self.entries:
            by_type[entry.transform_type] += 1
            by_slide[entry.slide_idx] += 1
            shape_type_counts[entry.shape_type] += 1

        return {
            "total_transforms": len(self.entries),
            "by_type": dict(sorted(by_type.items(), key=lambda kv: -kv[1])),
            "by_slide": dict(sorted(by_slide.items())),
            "exempt_count": by_type.get("exempt", 0),
            "skip_count": by_type.get("skip", 0),
            "shape_type_counts": dict(sorted(shape_type_counts.items(), key=lambda kv: -kv[1])),
            "elapsed_seconds": round(time.time() - self.start_time, 3),
            "deck_name": self.deck_name,
            "slide_count": self.slide_count,
            "shape_count": self.shape_count,
        }

    # ─────────────────────────────────────────────────────────────────────────
    # Output: JSON
    # ─────────────────────────────────────────────────────────────────────────

    def to_json(self, path: str) -> None:
        """Write full audit log as JSON to *path*.

        The JSON document has two top-level keys:
        - "summary"  : the dict returned by self.summary()
        - "entries"  : list of all TransformEntry records as dicts
        """
        timestamp_iso = datetime.fromtimestamp(self.start_time, tz=timezone.utc).isoformat()

        payload = {
            "metadata": {
                "deck_name": self.deck_name,
                "timestamp": timestamp_iso,
                "slide_count": self.slide_count,
                "shape_count": self.shape_count,
                "total_entries": len(self.entries),
            },
            "summary": self.summary(),
            "entries": [
                {
                    "slide_idx": e.slide_idx,
                    "shape_id": e.shape_id,
                    "shape_name": e.shape_name,
                    "shape_type": e.shape_type,
                    "transform_type": e.transform_type,
                    "before_state": e.before_state,
                    "after_state": e.after_state,
                    "notes": e.notes,
                    "timestamp": e.timestamp,
                }
                for e in self.entries
            ],
        }

        with open(path, "w", encoding="utf-8") as fh:
            json.dump(payload, fh, indent=2, ensure_ascii=False)

    # ─────────────────────────────────────────────────────────────────────────
    # Output: Markdown
    # ─────────────────────────────────────────────────────────────────────────

    def to_markdown(self, path: str) -> None:
        """Write a human-readable audit report as Markdown to *path*.

        Structure
        ---------
        - Header with deck name, timestamp, counts
        - Summary table: transform type → count
        - Per-slide sections, each with a table: Shape | Type | Transform | Before → After | Notes
        """
        timestamp_iso = datetime.fromtimestamp(self.start_time, tz=timezone.utc).strftime(
            "%Y-%m-%dT%H:%M:%SZ"
        )
        stats = self.summary()

        # Pre-group entries by slide
        by_slide: Dict[int, List[TransformEntry]] = defaultdict(list)
        for entry in self.entries:
            by_slide[entry.slide_idx].append(entry)

        lines: List[str] = []

        # ── Header ────────────────────────────────────────────────────────────
        lines.append("# SlideArabi — Transformation Audit Report")
        lines.append("")
        lines.append(f"**Deck:** {self.deck_name or '(unknown)'}  ")
        lines.append(f"**Timestamp:** {timestamp_iso}  ")
        lines.append(
            f"**Total slides:** {self.slide_count} | "
            f"**Total shapes:** {self.shape_count} | "
            f"**Transforms:** {stats['total_transforms']}"
        )
        lines.append("")

        # ── Summary table ─────────────────────────────────────────────────────
        lines.append("## Summary")
        lines.append("")
        lines.append("| Transform Type | Count |")
        lines.append("|---|---|")
        for t_type, count in stats["by_type"].items():
            lines.append(f"| {t_type} | {count} |")
        lines.append("")
        lines.append(f"**Elapsed:** {stats['elapsed_seconds']} s")
        lines.append("")

        # ── Per-slide sections ─────────────────────────────────────────────────
        for slide_idx in sorted(by_slide.keys()):
            slide_entries = by_slide[slide_idx]
            slide_total = len(slide_entries)
            lines.append(f"## Slide {slide_idx}")
            lines.append("")
            lines.append(
                f"*{slide_total} transform{'s' if slide_total != 1 else ''} on this slide.*"
            )
            lines.append("")
            lines.append("| Shape | Type | Transform | Before → After | Notes |")
            lines.append("|---|---|---|---|---|")

            for e in slide_entries:
                # Compact before/after representation
                if e.before_state or e.after_state:
                    before_str = _compact_state(e.before_state)
                    after_str  = _compact_state(e.after_state)
                    delta = f"`{before_str}` → `{after_str}`" if before_str or after_str else "—"
                else:
                    delta = "—"

                notes_cell = e.notes if e.notes else "—"
                # Escape pipe characters inside cells
                name   = e.shape_name.replace("|", "\\|")
                s_type = e.shape_type.replace("|", "\\|")
                t_type = e.transform_type.replace("|", "\\|")
                notes_cell = notes_cell.replace("|", "\\|")

                lines.append(f"| {name} | {s_type} | {t_type} | {delta} | {notes_cell} |")

            lines.append("")

        with open(path, "w", encoding="utf-8") as fh:
            fh.write("\n".join(lines))
            fh.write("\n")

    # ─────────────────────────────────────────────────────────────────────────
    # Output: Console summary
    # ─────────────────────────────────────────────────────────────────────────

    def print_summary(self) -> None:
        """Print a concise summary of the audit log to stdout."""
        stats = self.summary()
        print("─" * 60)
        print(f"SlideArabi Audit Summary — {self.deck_name or '(deck not set)'}")
        print(f"  Slides : {self.slide_count}  |  Shapes : {self.shape_count}")
        print(f"  Total transforms logged : {stats['total_transforms']}")
        print(f"  Elapsed : {stats['elapsed_seconds']} s")
        print()
        print("  By transform type:")
        for t_type, count in stats["by_type"].items():
            bar = "█" * min(count, 40)
            print(f"    {t_type:<25} {count:>5}  {bar}")
        print()
        print("  Slides with most transforms:")
        slide_counts = sorted(stats["by_slide"].items(), key=lambda kv: -kv[1])
        for slide_idx, count in slide_counts[:10]:
            print(f"    Slide {slide_idx:<4} : {count} transforms")
        if len(slide_counts) > 10:
            print(f"    … and {len(slide_counts) - 10} more slides")
        print("─" * 60)


# ─────────────────────────────────────────────────────────────────────────────
# Internal helpers
# ─────────────────────────────────────────────────────────────────────────────

def _compact_state(state: Dict[str, Any]) -> str:
    """Produce a short key=value string for a state dict, truncated at 80 chars."""
    if not state:
        return ""
    parts = []
    for k, v in state.items():
        if isinstance(v, float):
            parts.append(f"{k}={v:.2f}")
        else:
            parts.append(f"{k}={v}")
    raw = " ".join(parts)
    if len(raw) > 80:
        raw = raw[:77] + "…"
    return raw


# ─────────────────────────────────────────────────────────────────────────────
# Convenience: shape-type classifier
# ─────────────────────────────────────────────────────────────────────────────

def classify_shape_type(shape) -> str:
    """
    Infer a canonical shape_type string from a python-pptx Shape object.

    Returns one of: 'placeholder', 'text_box', 'image', 'chart', 'table',
    'group', 'connector', 'other'.

    This helper is provided for callers that don't want to replicate the
    classification logic at every call site.
    """
    try:
        # Group shapes have a .shapes attribute
        if hasattr(shape, "shapes"):
            return "group"

        if getattr(shape, "is_placeholder", False):
            return "placeholder"

        if getattr(shape, "has_chart", False) and shape.has_chart:
            return "chart"

        if getattr(shape, "has_table", False) and shape.has_table:
            return "table"

        sp_el = getattr(shape, "_element", None)
        if sp_el is not None:
            tag = sp_el.tag
            if tag.endswith("}pic"):
                return "image"
            if tag.endswith("}cxnSp"):
                return "connector"

        if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
            return "text_box"

    except Exception:
        pass

    return "other"
