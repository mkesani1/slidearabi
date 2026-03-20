"""
vqa_types.py — Canonical V3 VQA types shared across pipeline components.

These types are the shared contract between vqa_engine.py, visual_qa.py,
server.py, and the frontend API.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from enum import Enum
from typing import Any, Dict, List, Literal, Optional


class Severity(str, Enum):
    CRITICAL = "CRITICAL"
    HIGH = "HIGH"
    MEDIUM = "MEDIUM"
    LOW = "LOW"


class DefectStatus(str, Enum):
    OPEN = "open"
    FIXED = "fixed"
    UNRESOLVED = "unresolved"
    WAIVED = "waived"


@dataclass
class V3Defect:
    """V3 defect detected by XML checks or vision models.

    Coexists with the existing vqa_engine.Defect class — V3 checks produce
    V3Defect instances, which can be converted to legacy Defect via to_legacy().
    """

    code: str  # e.g., TABLE_COLUMNS_NOT_REVERSED
    category: str  # table, alignment, numbering, mirroring, direction
    severity: Severity
    slide_idx: int  # 1-based slide number
    object_id: Optional[str] = None  # shape id / table id
    evidence: Dict[str, Any] = field(default_factory=dict)  # measured values, xml paths, bbox deltas
    fixable: bool = False
    autofix_action: Optional[str] = None  # action key for remediator dispatch
    status: DefectStatus = DefectStatus.OPEN
    source: Literal["xml", "vision"] = "xml"
    description: str = ""

    def to_dict(self) -> Dict[str, Any]:
        d = asdict(self)
        d["severity"] = self.severity.value
        d["status"] = self.status.value
        return d

    def to_legacy(self, defect_id: str = "") -> Dict[str, Any]:
        """Convert to dict compatible with existing vqa_engine.Defect format."""
        return {
            "id": defect_id or f"V3-{self.code[:8]}-S{self.slide_idx}",
            "layer": self.source,
            "check": self.code.lower(),
            "severity": self.severity.value,
            "defect_type": self.category,
            "slide": self.slide_idx,
            "shape_id": int(self.object_id) if self.object_id and self.object_id.isdigit() else None,
            "shape_name": "",
            "description": self.description,
            "coordinates": self.evidence.get("coordinates"),
            "affected_text": self.evidence.get("affected_text", ""),
            "remediation": {"action": self.autofix_action} if self.autofix_action else None,
            "auto_fixable": self.fixable,
        }

    @classmethod
    def critical(cls, code: str, slide_idx: int, **kwargs) -> "V3Defect":
        return cls(
            code=code,
            category=_category_from_code(code),
            severity=Severity.CRITICAL,
            slide_idx=slide_idx,
            **kwargs,
        )

    @classmethod
    def high(cls, code: str, slide_idx: int, **kwargs) -> "V3Defect":
        return cls(
            code=code,
            category=_category_from_code(code),
            severity=Severity.HIGH,
            slide_idx=slide_idx,
            **kwargs,
        )

    @classmethod
    def medium(cls, code: str, slide_idx: int, **kwargs) -> "V3Defect":
        return cls(
            code=code,
            category=_category_from_code(code),
            severity=Severity.MEDIUM,
            slide_idx=slide_idx,
            **kwargs,
        )


@dataclass
class VQAGateResult:
    """Final quality gate decision for a deck."""

    status: Literal["completed", "completed_with_warnings", "failed_qa", "vqa_error"] = "completed"
    slides_checked_xml: int = 0
    slides_checked_vision: int = 0
    slides_auto_fixed: int = 0
    defects_found: int = 0
    defects_fixed: int = 0
    critical_remaining: int = 0
    high_remaining: int = 0
    blocking_issues: List[Dict[str, Any]] = field(default_factory=list)
    warning_issues: List[Dict[str, Any]] = field(default_factory=list)
    cost_usd: float = 0.0

    def to_api_dict(self) -> Dict[str, Any]:
        """Serialized for /status/{job_id} response — additive 'vqa' field."""
        return {
            "gate": self.status,
            "slides_checked_xml": self.slides_checked_xml,
            "slides_checked_vision": self.slides_checked_vision,
            "slides_auto_fixed": self.slides_auto_fixed,
            "defects_found": self.defects_found,
            "defects_fixed": self.defects_fixed,
            "critical_remaining": self.critical_remaining,
            "high_remaining": self.high_remaining,
            "blocking_issues": self.blocking_issues[:5],
            "warning_issues": self.warning_issues[:10],
            "cost_usd": round(self.cost_usd, 4),
        }


# ── Defect code → category mapping ──
_CODE_TO_CATEGORY = {
    "TABLE_COLUMNS_NOT_REVERSED": "table",
    "TABLE_COLUMN_ORDER_AMBIGUOUS": "table",
    "TABLE_CELL_RTL_MISSING": "alignment",
    "TABLE_CELL_ALIGN_NOT_RIGHT": "alignment",
    "TABLE_GARBLED_LAYOUT": "table",
    "TABLE_GRID_STRUCTURAL_ERROR": "table",
    "TABLE_GRIDCOL_NOT_REVERSED": "table",
    "TABLE_MERGED_CELL_INTEGRITY": "table",
    "ICON_IN_WRONG_TABLE_CELL": "icon",
    "ICON_MISSING_IN_CONVERTED": "icon",
    "PAGE_NUMBER_DUPLICATED": "numbering",
    "PAGE_NUMBER_DOUBLED_STRING": "numbering",
    "SHAPE_NOT_MIRRORED_POSITION": "mirroring",
    "MASTER_ELEMENT_NOT_MIRRORED": "mirroring",
    "PARAGRAPH_RTL_MISSING": "alignment",
    "TEXT_NOT_CENTERED_IN_SHAPE": "alignment",
    "DIRECTIONAL_SHAPE_NOT_FLIPPED": "direction",
}


def _category_from_code(code: str) -> str:
    return _CODE_TO_CATEGORY.get(code, "unknown")


# ── Must-not-ship defect codes (gate blocks on ANY of these) ──
MUST_NOT_SHIP_CODES = frozenset(
    {
        "TABLE_COLUMNS_NOT_REVERSED",
        "TABLE_GRID_STRUCTURAL_ERROR",
        "TABLE_GARBLED_LAYOUT",
    }
)

# Gate thresholds
HIGH_SEVERITY_THRESHOLD = 3  # >= 3 HIGH defects → completed_with_warnings
HIGH_SLIDE_RATIO_THRESHOLD = 0.10  # >= 10% of slides with HIGH → failed_qa
