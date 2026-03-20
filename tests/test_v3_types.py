import importlib

from vqa_types import DefectStatus, Severity, V3Defect, VQAGateResult


def test_v3defect_factory_methods_set_expected_defaults():
    d_critical = V3Defect.critical("TABLE_COLUMNS_NOT_REVERSED", 1, object_id="42")
    d_high = V3Defect.high("PAGE_NUMBER_DUPLICATED", 2)
    d_medium = V3Defect.medium("PARAGRAPH_RTL_MISSING", 3)

    assert d_critical.severity is Severity.CRITICAL
    assert d_critical.category == "table"
    assert d_critical.slide_idx == 1
    assert d_critical.object_id == "42"
    assert d_critical.status is DefectStatus.OPEN

    assert d_high.severity is Severity.HIGH
    assert d_high.category == "numbering"

    assert d_medium.severity is Severity.MEDIUM
    assert d_medium.category == "alignment"


def test_v3defect_to_dict_serialization():
    defect = V3Defect.high(
        "ICON_MISSING_IN_CONVERTED",
        4,
        object_id="7",
        evidence={"coordinates": {"x": 10}, "affected_text": "icon missing"},
        fixable=True,
        autofix_action="reinsert_icon",
        source="vision",
        description="Icon is missing in converted slide",
    )

    data = defect.to_dict()

    assert data["severity"] == "HIGH"
    assert data["status"] == "open"
    assert data["source"] == "vision"
    assert data["fixable"] is True
    assert data["evidence"]["affected_text"] == "icon missing"


def test_v3defect_to_legacy_compatibility_shape_and_fields():
    defect = V3Defect.critical(
        "TABLE_GRID_STRUCTURAL_ERROR",
        6,
        object_id="123",
        evidence={"coordinates": {"x": 1, "y": 2}, "affected_text": "bad layout"},
        fixable=True,
        autofix_action="repair_table",
        description="Table layout appears garbled",
    )

    legacy = defect.to_legacy("S6-D001")

    assert legacy["id"] == "S6-D001"
    assert legacy["layer"] == "xml"
    assert legacy["check"] == "table_grid_structural_error"
    assert legacy["severity"] == "CRITICAL"
    assert legacy["defect_type"] == "table"
    assert legacy["slide"] == 6
    assert legacy["shape_id"] == 123
    assert legacy["shape_name"] == ""
    assert legacy["description"] == "Table layout appears garbled"
    assert legacy["coordinates"] == {"x": 1, "y": 2}
    assert legacy["affected_text"] == "bad layout"
    assert legacy["remediation"] == {"action": "repair_table"}
    assert legacy["auto_fixable"] is True


def test_v3defect_to_legacy_handles_non_digit_shape_id_and_default_id():
    defect = V3Defect.medium("UNKNOWN_CODE", 8, object_id="tbl-01")
    legacy = defect.to_legacy()

    assert legacy["id"].startswith("V3-UNKNOWN_-S8")
    assert legacy["shape_id"] is None
    assert legacy["defect_type"] == "unknown"


def test_vqagate_result_to_api_dict_truncates_and_rounds_cost():
    gate = VQAGateResult(
        status="completed_with_warnings",
        slides_checked_xml=12,
        slides_checked_vision=4,
        slides_auto_fixed=3,
        defects_found=11,
        defects_fixed=7,
        critical_remaining=1,
        high_remaining=2,
        blocking_issues=[{"i": i} for i in range(8)],
        warning_issues=[{"i": i} for i in range(15)],
        cost_usd=0.123456,
    )

    api = gate.to_api_dict()

    assert api["gate"] == "completed_with_warnings"
    assert api["slides_checked_xml"] == 12
    assert api["slides_checked_vision"] == 4
    assert api["slides_auto_fixed"] == 3
    assert api["defects_found"] == 11
    assert api["defects_fixed"] == 7
    assert api["critical_remaining"] == 1
    assert api["high_remaining"] == 2
    assert len(api["blocking_issues"]) == 5
    assert len(api["warning_issues"]) == 10
    assert api["cost_usd"] == 0.1235


def test_v3_config_defaults_are_safe_off(monkeypatch):
    keys = [
        "V3_XML_CHECKS",
        "V3_XML_AUTOFIX",
        "V3_TABLE_AUTOFIX",
        "V3_GATE_MODE",
        "V3_ENHANCED_PROMPTS",
        "V3_VISION_XML_CONTEXT",
        "V3_SELECTIVE_VISION",
        "V3_MASTER_MIRROR_FIX",
        "V3_MAX_VISION_SLIDES",
        "V3_MAX_REMEDIATION_PASSES",
        "V3_VQA_MAX_COST_USD",
        "V3_CANARY_PCT",
    ]
    for key in keys:
        monkeypatch.delenv(key, raising=False)

    import v3_config

    cfg = importlib.reload(v3_config)

    assert cfg.ENABLE_V3_VQA is False
    assert cfg.ENABLE_V3_XML_AUTOFIX is False
    assert cfg.ENABLE_V3_TABLE_FIX is False
    assert cfg.V3_GATE_MODE == "shadow"
    assert cfg.ENABLE_ENHANCED_PROMPTS is False
    assert cfg.ENABLE_VISION_XML_CTX is False
    assert cfg.ENABLE_SELECTIVE_VISION is False
    assert cfg.ENABLE_MASTER_MIRROR_FIX is False

    assert cfg.MAX_VISION_SLIDES == 20
    assert cfg.MAX_REMEDIATION_PASSES == 2
    assert cfg.VQA_TARGET_MAX_COST_USD == 0.45
    assert cfg.CANARY_PERCENTAGE == 0


def test_v3_config_default_helpers_report_disabled(monkeypatch):
    monkeypatch.delenv("V3_XML_CHECKS", raising=False)
    monkeypatch.delenv("V3_GATE_MODE", raising=False)

    import v3_config

    cfg = importlib.reload(v3_config)

    assert cfg.is_v3_enabled() is False
    assert cfg.is_gate_active() is False
    assert cfg.is_shadow_mode() is True
