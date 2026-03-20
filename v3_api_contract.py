"""
v3_api_contract.py — V3 API Contract Types for /status/{job_id}

Sprint 3: Defines the enhanced status response format with VQA gate decisions,
4 terminal statuses, and backward compatibility with V2 clients.
"""

from __future__ import annotations

from typing import Any, Dict, List, Optional


def build_status_response(
    job_id: str,
    phase: str,
    progress: float,
    gate_result: Optional[Any] = None,
    download_url: Optional[str] = None,
    error: Optional[str] = None,
) -> Dict[str, Any]:
    """Build the /status/{job_id} response with V3 VQA fields.
    
    Backward compatible: V2 clients see standard fields. V3 clients get
    additional 'vqa' and 'download_available' fields.
    
    Terminal statuses:
        - completed: all checks pass
        - completed_with_warnings: minor residual issues, download available
        - failed_qa: critical defects remain, download still available
        - vqa_error: VQA infrastructure failure, download still available
        - processing: not yet finished
        - error: non-VQA error (upload, translation, etc.)
    """
    response: Dict[str, Any] = {
        'job_id': job_id,
        'status': _resolve_status(phase, gate_result, error),
        'phase': phase,
        'progress': round(progress, 2),
    }

    # Download is available once conversion is done, regardless of QA outcome
    response['download_available'] = download_url is not None
    if download_url:
        response['download_url'] = download_url

    # V3 VQA summary (additive — V2 clients ignore unknown fields)
    if gate_result is not None:
        response['vqa'] = _format_vqa_summary(gate_result)

    if error:
        response['error'] = error

    return response


def _resolve_status(
    phase: str,
    gate_result: Optional[Any],
    error: Optional[str],
) -> str:
    """Map pipeline phase + gate result to terminal status string."""
    if error:
        return 'error'

    if phase == 'done' or phase == 'completed':
        if gate_result is None:
            return 'completed'  # V2 path — no gate
        
        gate_status = getattr(gate_result, 'status', 'completed')
        if gate_status in ('completed', 'completed_with_warnings',
                          'failed_qa', 'vqa_error'):
            return gate_status
        return 'completed'

    # Still processing
    return 'processing'


def _format_vqa_summary(gate_result: Any) -> Dict[str, Any]:
    """Format VQAGateResult into API-friendly dict."""
    summary: Dict[str, Any] = {
        'status': getattr(gate_result, 'status', 'unknown'),
        'critical_remaining': getattr(gate_result, 'critical_remaining', 0),
        'high_remaining': getattr(gate_result, 'high_remaining', 0),
    }

    blocking = getattr(gate_result, 'blocking_issues', [])
    if blocking:
        summary['blocking_issues'] = blocking[:5]  # Cap at 5

    warnings = getattr(gate_result, 'warning_issues', [])
    if warnings:
        summary['warning_issues'] = warnings[:10]  # Cap at 10

    return summary


def is_download_safe(gate_result: Optional[Any]) -> bool:
    """Check if the file should be available for download.
    
    Policy: always allow download, even for failed_qa.
    The user paid for it and may want the imperfect result.
    The status field communicates the quality level.
    """
    return True
