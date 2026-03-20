"""
v3_config.py — V3 Feature Flags (env-var driven, all default OFF)

Controls V3 VQA quality gate features. All flags are read from environment
variables and default to safe (V2-compatible) values. Toggle via Railway env vars.
"""

import os


def _env_bool(key: str, default: bool = False) -> bool:
    val = os.environ.get(key, "").strip().lower()
    if val in ("1", "true", "yes", "on"):
        return True
    if val in ("0", "false", "no", "off", ""):
        return default
    return default


def _env_int(key: str, default: int = 0) -> int:
    try:
        return int(os.environ.get(key, str(default)))
    except (ValueError, TypeError):
        return default


def _env_float(key: str, default: float = 0.0) -> float:
    try:
        return float(os.environ.get(key, str(default)))
    except (ValueError, TypeError):
        return default


def _env_str(key: str, default: str = "") -> str:
    return os.environ.get(key, default).strip()


# ── V3 Feature Flags ──

# Master switch: enable V3 XML-first VQA pipeline
ENABLE_V3_VQA = _env_bool("V3_XML_CHECKS", False)

# Enable XML auto-fixes (deterministic remediation)
ENABLE_V3_XML_AUTOFIX = _env_bool("V3_XML_AUTOFIX", False)

# Enable table column reversal auto-fix specifically (highest risk fix)
ENABLE_V3_TABLE_FIX = _env_bool("V3_TABLE_AUTOFIX", False)

# Gate mode: "shadow" (log but don't block) or "active" (enforce terminal statuses)
# Default: "active" — pipeline.py now wires gate_result to PipelineResult,
# so the gate decision flows to build_status_response() for API consumers.
# Override to "shadow" via env var if you want logging without enforcement.
V3_GATE_MODE = _env_str("V3_GATE_MODE", "active")

# Enable enhanced vision prompts (table/icon/alignment categories)
ENABLE_ENHANCED_PROMPTS = _env_bool("V3_ENHANCED_PROMPTS", False)

# Enable passing XML defects as context to vision models
ENABLE_VISION_XML_CTX = _env_bool("V3_VISION_XML_CONTEXT", False)

# Enable selective vision (only uncertain slides get vision QA)
ENABLE_SELECTIVE_VISION = _env_bool("V3_SELECTIVE_VISION", False)

# Enable master/layout mirroring fix
ENABLE_MASTER_MIRROR_FIX = _env_bool("V3_MASTER_MIRROR_FIX", False)

# ── Tuning parameters ──

# Max vision slides per deck (cost control)
MAX_VISION_SLIDES = _env_int("V3_MAX_VISION_SLIDES", 20)

# Max remediation passes
MAX_REMEDIATION_PASSES = _env_int("V3_MAX_REMEDIATION_PASSES", 2)

# Cost target per deck (USD)
VQA_TARGET_MAX_COST_USD = _env_float("V3_VQA_MAX_COST_USD", 0.45)

# Canary rollout percentage (0-100, 0 = disabled)
CANARY_PERCENTAGE = _env_int("V3_CANARY_PCT", 0)


def is_v3_enabled() -> bool:
    """Check if V3 VQA pipeline should be used for this request."""
    return ENABLE_V3_VQA


def is_gate_active() -> bool:
    """Check if gate decisions should be enforced (not just shadow)."""
    return V3_GATE_MODE == "active"


def is_shadow_mode() -> bool:
    """Check if gate is in shadow mode (log decisions but don't block)."""
    return V3_GATE_MODE == "shadow"


def should_process_job(job_id: str) -> bool:
    """Canary check: should this job use V3 based on canary percentage."""
    if CANARY_PERCENTAGE <= 0:
        return ENABLE_V3_VQA
    if CANARY_PERCENTAGE >= 100:
        return True

    # Deterministic canary based on job_id hash
    import hashlib

    hash_val = int(hashlib.md5(job_id.encode()).hexdigest(), 16) % 100
    return hash_val < CANARY_PERCENTAGE
