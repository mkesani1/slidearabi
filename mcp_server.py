"""
SlideArabi — MCP Server (Model Context Protocol)
==================================================

Provides AI agents (Claude Desktop, Cursor, Perplexity, etc.) with tools to
convert English PowerPoint presentations to Arabic RTL format.

Transport: Streamable HTTP mounted at /mcp on the FastAPI app.

Mount into the main app::

    from slidearabi.mcp_server import create_mcp_app
    app.mount("/mcp", create_mcp_app())

Client configuration (Claude Desktop)::

    {
      "mcpServers": {
        "slidearabi": {
          "type": "streamable-http",
          "url": "https://api.slidearabi.com/mcp",
          "headers": {
            "Authorization": "Bearer sa_live_xxxxxxxxxxxx"
          }
        }
      }
    }
"""

from __future__ import annotations

import base64
import hashlib
import logging
import os
import secrets
from datetime import datetime, timedelta, timezone
from io import BytesIO
from typing import Any, Optional

import httpx
from fastmcp import FastMCP
from fastmcp.exceptions import ToolError
from fastmcp.server.dependencies import CurrentHeaders
from supabase import Client as SupabaseClient, create_client

logger = logging.getLogger("slidearabi.mcp_server")

# ─── Configuration ───────────────────────────────────────────────────────────

SUPABASE_URL: str = os.environ.get("SUPABASE_URL", "")
SUPABASE_SERVICE_KEY: str = os.environ.get("SUPABASE_SERVICE_KEY", "")
INTERNAL_BACKEND_URL: str = os.environ.get(
    "INTERNAL_BACKEND_URL",
    "https://slidearabi-production.up.railway.app",
)
DASHBOARD_URL: str = "https://www.slidearabi.com/dashboard"

# ─── FastMCP Instance ────────────────────────────────────────────────────────

mcp = FastMCP(
    "SlideArabi",
    version="1.0.0",
    instructions=(
        "AI-powered English to Arabic PowerPoint conversion. "
        "Convert .pptx files with full RTL layout transformation, "
        "chart mirroring, and professional Arabic typography.\n\n"
        "Workflow: upload_presentation → convert_presentation → "
        "poll get_conversion_status every 20-30s → download_result."
    ),
)

# ─── Helpers ─────────────────────────────────────────────────────────────────


def _get_supabase() -> SupabaseClient:
    """Get a Supabase client with service role key (bypasses RLS)."""
    if not SUPABASE_URL or not SUPABASE_SERVICE_KEY:
        raise ToolError(
            "Server configuration error: Supabase credentials not set. "
            "Contact support."
        )
    return create_client(SUPABASE_URL, SUPABASE_SERVICE_KEY)


def _generate_upload_id() -> str:
    """Generate a unique upload identifier."""
    return f"upl_{secrets.token_hex(4)}"


def _generate_job_id() -> str:
    """Generate a unique job identifier."""
    return f"job_{secrets.token_hex(4)}"


async def _get_auth_context(headers: dict[str, str]) -> dict[str, Any]:
    """Extract and validate the API key from the MCP request's HTTP headers.

    Looks up the Bearer token's SHA-256 hash in the Supabase ``api_keys``
    table and resolves the associated account.

    Args:
        headers: HTTP headers dict injected via ``CurrentHeaders()``.

    Returns:
        ``{"account_id": str, "api_key_id": str, "plan_tier": str, "is_test": bool}``

    Raises:
        ToolError: If the API key is missing, invalid, or revoked.
    """
    auth_header = headers.get("authorization", "")

    if not auth_header:
        raise ToolError(
            "Missing Authorization header. Configure your MCP client with: "
            'Authorization: Bearer sa_live_xxx  — '
            "Create a key at https://www.slidearabi.com/dashboard/api-keys"
        )

    parts = auth_header.split(" ", 1)
    if len(parts) != 2 or parts[0].lower() != "bearer":
        raise ToolError("Authorization header must be: Bearer <api_key>")

    api_key = parts[1].strip()
    key_hash = hashlib.sha256(api_key.encode()).hexdigest()

    sb = _get_supabase()
    result = (
        sb.table("api_keys")
        .select("id, account_id, is_test, is_active")
        .eq("key_hash", key_hash)
        .execute()
    )

    if not result.data:
        raise ToolError(
            "Invalid API key. Create one at "
            "https://www.slidearabi.com/dashboard/api-keys"
        )

    key_row = result.data[0]
    if not key_row["is_active"]:
        raise ToolError("This API key has been revoked.")

    # Update last_used_at (fire-and-forget, non-blocking)
    try:
        sb.table("api_keys").update(
            {"last_used_at": datetime.now(timezone.utc).isoformat()}
        ).eq("id", key_row["id"]).execute()
    except Exception:
        pass

    # Get plan tier
    account = (
        sb.table("api_accounts")
        .select("plan")
        .eq("id", key_row["account_id"])
        .single()
        .execute()
    )
    plan = (
        account.data.get("plan", "pay_as_you_go") if account.data else "pay_as_you_go"
    )

    return {
        "account_id": key_row["account_id"],
        "api_key_id": key_row["id"],
        "plan_tier": plan,
        "is_test": key_row.get("is_test", False),
    }


async def _get_credit_balance(account_id: str) -> dict[str, int]:
    """Fetch the credit balance for an account.

    Returns:
        ``{"available": int, "reserved": int}``
    """
    sb = _get_supabase()
    account = (
        sb.table("api_accounts")
        .select("credits_available, credits_reserved")
        .eq("id", account_id)
        .single()
        .execute()
    )
    if not account.data:
        return {"available": 0, "reserved": 0}
    return {
        "available": account.data["credits_available"],
        "reserved": account.data["credits_reserved"],
    }


async def _call_internal_api(
    method: str, path: str, **kwargs: Any
) -> httpx.Response:
    """Call the internal SlideArabi conversion backend.

    Args:
        method: HTTP method (GET, POST, etc.).
        path: URL path to append to the backend base URL.
        **kwargs: Forwarded to ``httpx.AsyncClient.request``.
    """
    url = f"{INTERNAL_BACKEND_URL}{path}"
    async with httpx.AsyncClient(timeout=60.0) as client:
        return await client.request(method, url, **kwargs)


# ─── MCP Tool 1: Upload Presentation ────────────────────────────────────────


@mcp.tool()
async def upload_presentation(
    filename: str,
    file_base64: Optional[str] = None,
    file_url: Optional[str] = None,
    headers: dict[str, str] = CurrentHeaders(),
) -> dict[str, Any]:
    """Upload a PowerPoint (.pptx) file for English-to-Arabic RTL conversion.

    Provide the file as either base64-encoded content OR a publicly accessible
    URL. Returns an upload_id to use with convert_presentation, along with the
    slide count and credit cost estimate.

    Args:
        filename: Original filename (must end in .pptx).
        file_base64: Base64-encoded .pptx file content. Use this OR file_url.
        file_url: Publicly accessible URL to a .pptx file. Use this OR file_base64.

    Returns:
        upload_id, filename, slide_count, size_bytes, credits_required,
        credits_available, expires_at. If credits are insufficient, includes
        a warning and top_up_url.
    """
    auth = await _get_auth_context(headers)

    if not filename.lower().endswith(".pptx"):
        raise ToolError("Filename must end in .pptx")

    if not file_base64 and not file_url:
        raise ToolError("Provide either file_base64 or file_url.")

    # Get file bytes
    if file_base64:
        try:
            file_bytes = base64.b64decode(file_base64)
        except Exception:
            raise ToolError("Invalid base64 encoding.")
    else:
        try:
            async with httpx.AsyncClient(timeout=60.0) as client:
                resp = await client.get(file_url)
                resp.raise_for_status()
                file_bytes = resp.content
        except httpx.HTTPStatusError as e:
            raise ToolError(
                f"Could not download file from URL (HTTP {e.response.status_code})."
            )
        except Exception as e:
            raise ToolError(f"Could not download file from URL: {e}")

    # Size check (50 MB max for API uploads)
    if len(file_bytes) > 50 * 1024 * 1024:
        raise ToolError("File exceeds 50 MB limit.")

    # Count slides using python-pptx
    try:
        from pptx import Presentation

        prs = Presentation(BytesIO(file_bytes))
        slide_count = len(prs.slides)
    except Exception:
        raise ToolError(
            "Could not parse .pptx file. Ensure it is a valid PowerPoint presentation."
        )

    if slide_count < 1:
        raise ToolError("Presentation has no slides.")

    credits_required = max(5, slide_count)
    balance = await _get_credit_balance(auth["account_id"])

    # Stage file in Supabase Storage
    sb = _get_supabase()
    upload_id = _generate_upload_id()
    storage_path = f"api-uploads/{auth['account_id']}/{upload_id}/{filename}"

    try:
        sb.storage.from_("api-files").upload(storage_path, file_bytes)
    except Exception:
        try:
            sb.storage.create_bucket("api-files", options={"public": False})
            sb.storage.from_("api-files").upload(storage_path, file_bytes)
        except Exception as e:
            raise ToolError(f"File storage unavailable: {e}")

    expires_at = (datetime.now(timezone.utc) + timedelta(hours=24)).isoformat()

    sb.table("api_uploads").insert(
        {
            "id": upload_id,
            "account_id": auth["account_id"],
            "filename": filename,
            "slide_count": slide_count,
            "size_bytes": len(file_bytes),
            "storage_path": storage_path,
            "expires_at": expires_at,
        }
    ).execute()

    result: dict[str, Any] = {
        "upload_id": upload_id,
        "filename": filename,
        "slide_count": slide_count,
        "size_bytes": len(file_bytes),
        "credits_required": credits_required,
        "credits_available": balance["available"],
        "expires_at": expires_at,
    }

    if balance["available"] < credits_required:
        result["warning"] = (
            f"Insufficient credits: you have {balance['available']} but need "
            f"{credits_required}. Purchase credits before converting."
        )
        result["top_up_url"] = f"{DASHBOARD_URL}/credits?prefill={credits_required}"

    return result


# ─── MCP Tool 2: Convert Presentation ───────────────────────────────────────


@mcp.tool()
async def convert_presentation(
    upload_id: str,
    options: Optional[dict[str, Any]] = None,
    headers: dict[str, str] = CurrentHeaders(),
) -> dict[str, Any]:
    """Start English-to-Arabic RTL conversion of a previously uploaded PowerPoint file.

    Conversion takes 2-8 minutes depending on slide count. After calling this,
    poll get_conversion_status every 20-30 seconds until status is 'done'.

    Args:
        upload_id: The upload_id returned by upload_presentation.
        options: Optional conversion settings. Supported keys:
            - dialect: "msa" (default), "gulf", "levantine", "egyptian"
            - formality: "formal" (default), "standard", "informal"
            - preserve_fonts: true (default) keeps Arabic-compatible font variants

    Returns:
        job_id, status ("queued"), slide_count, credits_reserved,
        estimated_minutes, poll_interval_seconds.
    """
    auth = await _get_auth_context(headers)
    sb = _get_supabase()

    # Validate upload exists and belongs to this account
    upload = (
        sb.table("api_uploads")
        .select("*")
        .eq("id", upload_id)
        .eq("account_id", auth["account_id"])
        .execute()
    )
    if not upload.data:
        raise ToolError(
            f"Upload {upload_id} not found. "
            "Did you call upload_presentation first?"
        )

    upload_row = upload.data[0]
    if (
        datetime.fromisoformat(upload_row["expires_at"].replace("Z", "+00:00"))
        < datetime.now(timezone.utc)
    ):
        raise ToolError(
            "This upload has expired (24-hour limit). Please upload again."
        )

    slide_count = upload_row["slide_count"]
    credits_needed = max(5, slide_count)
    job_id = _generate_job_id()

    # Reserve credits (skip for test keys)
    if not auth.get("is_test"):
        reserve_result = sb.rpc(
            "reserve_credits",
            {
                "p_account_id": auth["account_id"],
                "p_amount": credits_needed,
                "p_job_id": job_id,
            },
        ).execute()

        if reserve_result.data is False:
            balance = await _get_credit_balance(auth["account_id"])
            return {
                "error": "INSUFFICIENT_CREDITS",
                "message": (
                    f"You have {balance['available']} credits but need "
                    f"{credits_needed}."
                ),
                "credits_available": balance["available"],
                "credits_required": credits_needed,
                "top_up_url": f"{DASHBOARD_URL}/credits?prefill={credits_needed}",
            }

    # Create job record in database
    sb.table("api_jobs").insert(
        {
            "id": job_id,
            "account_id": auth["account_id"],
            "upload_id": upload_id,
            "source": "mcp",
            "status": "queued",
            "slide_count": slide_count,
            "credits_reserved": credits_needed,
            "options": options or {},
        }
    ).execute()

    # Trigger conversion on the internal backend
    try:
        file_bytes = sb.storage.from_("api-files").download(
            upload_row["storage_path"]
        )
        backend_resp = await _call_internal_api(
            "POST",
            "/convert",
            files={
                "file": (
                    upload_row["filename"],
                    file_bytes,
                    "application/vnd.openxmlformats-officedocument"
                    ".presentationml.presentation",
                )
            },
        )
        if backend_resp.status_code == 200:
            backend_data = backend_resp.json()
            # Store the backend's internal job_id for status polling
            internal_job_id = backend_data.get("job_id")
            update_fields: dict[str, Any] = {
                "status": "processing",
                "started_at": datetime.now(timezone.utc).isoformat(),
            }
            if internal_job_id:
                update_fields["options"] = {
                    **(options or {}),
                    "_internal_job_id": internal_job_id,
                }
            sb.table("api_jobs").update(update_fields).eq(
                "id", job_id
            ).execute()
    except Exception as e:
        logger.error(
            "Failed to trigger backend for MCP job %s: %s", job_id, e
        )

    estimated_minutes = max(2, round(slide_count * 0.25))

    return {
        "job_id": job_id,
        "status": "queued",
        "slide_count": slide_count,
        "credits_reserved": credits_needed,
        "estimated_minutes": estimated_minutes,
        "poll_interval_seconds": 20,
    }


# ─── MCP Tool 3: Get Conversion Status ──────────────────────────────────────


@mcp.tool()
async def get_conversion_status(
    job_id: str,
    headers: dict[str, str] = CurrentHeaders(),
) -> dict[str, Any]:
    """Check the status and progress of a conversion job.

    Poll this every 20-30 seconds after calling convert_presentation.
    Do not call more frequently. Typical conversions take 3-6 minutes.

    Args:
        job_id: The job_id returned by convert_presentation.

    Returns:
        job_id, status, progress_percent, current_phase,
        estimated_remaining_seconds.
        When status is 'done': includes download_url and credits_charged.
        When status is 'failed': includes error_code, error_message,
        credits_refunded.
    """
    auth = await _get_auth_context(headers)
    sb = _get_supabase()

    job = (
        sb.table("api_jobs")
        .select("*")
        .eq("id", job_id)
        .eq("account_id", auth["account_id"])
        .execute()
    )
    if not job.data:
        raise ToolError(f"Job {job_id} not found.")

    row = job.data[0]

    # Resolve internal backend job_id (may differ from our API job_id)
    internal_job_id = (row.get("options") or {}).get("_internal_job_id", job_id)

    # Sync with internal backend status for in-flight jobs
    if row["status"] in ("queued", "processing"):
        try:
            backend_resp = await _call_internal_api(
                "GET", f"/status/{internal_job_id}"
            )
            if backend_resp.status_code == 200:
                backend_data = backend_resp.json()
                new_status = backend_data.get("status", row["status"])
                new_progress = backend_data.get(
                    "progress_pct", row["progress_percent"]
                )
                new_phase = backend_data.get(
                    "current_phase", row["current_phase"]
                )

                if (
                    new_status != row["status"]
                    or new_progress != row["progress_percent"]
                ):
                    update: dict[str, Any] = {
                        "progress_percent": new_progress,
                        "current_phase": new_phase,
                    }
                    if (
                        new_status == "completed"
                        and row["status"] != "done"
                    ):
                        update["status"] = "done"
                        update["completed_at"] = datetime.now(
                            timezone.utc
                        ).isoformat()
                        update["credits_charged"] = row[
                            "credits_reserved"
                        ]
                        # Settle credits
                        sb.rpc(
                            "settle_credits",
                            {
                                "p_account_id": auth["account_id"],
                                "p_amount": row["credits_reserved"],
                                "p_job_id": job_id,
                            },
                        ).execute()
                    elif (
                        new_status == "failed"
                        and row["status"] != "failed"
                    ):
                        update["status"] = "failed"
                        update["error_code"] = "PROCESSING_FAILED"
                        update["error_message"] = backend_data.get(
                            "error", "Conversion failed."
                        )
                        # Release reserved credits
                        sb.rpc(
                            "release_credits",
                            {
                                "p_account_id": auth["account_id"],
                                "p_amount": row["credits_reserved"],
                                "p_job_id": job_id,
                            },
                        ).execute()
                    else:
                        update["status"] = new_status

                    sb.table("api_jobs").update(update).eq(
                        "id", job_id
                    ).execute()
                    row.update(update)
        except Exception as e:
            logger.debug(
                "Could not sync with backend for job %s: %s", job_id, e
            )

    result: dict[str, Any] = {
        "job_id": row["id"],
        "status": row["status"],
        "progress_percent": row["progress_percent"] or 0,
        "current_phase": row["current_phase"],
        "slide_count": row["slide_count"],
    }

    if row["status"] == "done":
        result["download_url"] = (
            f"https://api.slidearabi.com/v1/conversions/{job_id}/download"
        )
        result["credits_charged"] = row.get(
            "credits_charged", row["credits_reserved"]
        )
        result["completed_at"] = row.get("completed_at")
    elif row["status"] == "failed":
        result["error_code"] = row.get("error_code", "PROCESSING_FAILED")
        result["error_message"] = row.get(
            "error_message", "Conversion failed."
        )
        result["credits_refunded"] = row["credits_reserved"]
    elif row["status"] in ("queued", "processing"):
        elapsed = 0
        if row.get("started_at"):
            started = datetime.fromisoformat(
                row["started_at"].replace("Z", "+00:00")
            )
            elapsed = (
                datetime.now(timezone.utc) - started
            ).total_seconds()
        estimated_total = row["slide_count"] * 15  # ~15s per slide
        remaining = max(0, estimated_total - elapsed)
        result["estimated_remaining_seconds"] = int(remaining)

    return result


# ─── MCP Tool 4: Download Result ────────────────────────────────────────────


@mcp.tool()
async def download_result(
    job_id: str,
    format: str = "url",
    headers: dict[str, str] = CurrentHeaders(),
) -> dict[str, Any]:
    """Download the completed Arabic RTL PowerPoint file.

    Only available after get_conversion_status returns status 'done'.
    Use format='url' (default) to get a pre-signed download link valid for
    1 hour. Use format='base64' only for small files — returns the entire
    file inline (warning: large responses).

    Args:
        job_id: The job_id of a completed conversion.
        format: 'url' (recommended) or 'base64'. Default: 'url'.

    Returns:
        download_url or file_base64, plus job_id, filename, size_bytes,
        and expires_at.
    """
    if format not in ("url", "base64"):
        raise ToolError("format must be 'url' or 'base64'.")

    auth = await _get_auth_context(headers)
    sb = _get_supabase()

    job = (
        sb.table("api_jobs")
        .select("status, result_path, slide_count, upload_id, account_id")
        .eq("id", job_id)
        .eq("account_id", auth["account_id"])
        .execute()
    )

    if not job.data:
        raise ToolError(f"Job {job_id} not found.")

    row = job.data[0]
    if row["status"] != "done":
        raise ToolError(
            f"Job is not complete (current status: {row['status']}). "
            "Call get_conversion_status to check progress."
        )

    # Get original filename from upload record
    upload = (
        sb.table("api_uploads")
        .select("filename")
        .eq("id", row["upload_id"])
        .execute()
    )
    original_name = (
        upload.data[0]["filename"] if upload.data else "presentation.pptx"
    )
    arabic_name = original_name.replace(".pptx", "_arabic.pptx")

    # Resolve internal backend job_id for download
    internal_job_id = job_id
    job_full = (
        sb.table("api_jobs")
        .select("options")
        .eq("id", job_id)
        .execute()
    )
    if job_full.data:
        internal_job_id = (job_full.data[0].get("options") or {}).get(
            "_internal_job_id", job_id
        )

    if format == "base64":
        try:
            if row.get("result_path"):
                file_bytes = sb.storage.from_("api-files").download(
                    row["result_path"]
                )
            else:
                backend_resp = await _call_internal_api(
                    "GET", f"/download/{internal_job_id}"
                )
                backend_resp.raise_for_status()
                file_bytes = backend_resp.content
            encoded = base64.b64encode(file_bytes).decode()
            return {
                "job_id": job_id,
                "filename": arabic_name,
                "file_base64": encoded,
                "size_bytes": len(file_bytes),
            }
        except ToolError:
            raise
        except Exception as e:
            raise ToolError(f"Could not retrieve file: {e}")
    else:
        # Return a pre-signed URL
        if row.get("result_path"):
            try:
                signed = sb.storage.from_("api-files").create_signed_url(
                    row["result_path"], 3600
                )
                return {
                    "job_id": job_id,
                    "filename": arabic_name,
                    "download_url": signed["signedURL"],
                    "size_bytes": None,
                    "expires_at": (
                        datetime.now(timezone.utc) + timedelta(hours=1)
                    ).isoformat(),
                }
            except Exception:
                pass

        # Fallback: direct backend download URL
        return {
            "job_id": job_id,
            "filename": arabic_name,
            "download_url": (
                f"https://api.slidearabi.com/v1/conversions/"
                f"{job_id}/download"
            ),
            "size_bytes": None,
            "expires_at": (
                datetime.now(timezone.utc) + timedelta(days=7)
            ).isoformat(),
        }


# ─── MCP Tool 5: Get Account Info ───────────────────────────────────────────


@mcp.tool()
async def get_account_info(
    headers: dict[str, str] = CurrentHeaders(),
) -> dict[str, Any]:
    """Check your credit balance, usage statistics, and account limits.

    No inputs required — uses your API key for authentication.
    If credits are low, use the top_up_url to purchase more.

    Returns:
        credits_available, credits_reserved, credits_used_this_month,
        plan, rate_limits, top_up_url.
    """
    auth = await _get_auth_context(headers)
    sb = _get_supabase()

    account = (
        sb.table("api_accounts")
        .select("*")
        .eq("id", auth["account_id"])
        .single()
        .execute()
    )
    if not account.data:
        raise ToolError("Account not found.")

    row = account.data

    # Calculate credits used this calendar month
    month_start = datetime.now(timezone.utc).replace(
        day=1, hour=0, minute=0, second=0, microsecond=0
    )
    monthly_usage = (
        sb.table("credit_transactions")
        .select("amount")
        .eq("account_id", auth["account_id"])
        .eq("type", "settle")
        .gte("created_at", month_start.isoformat())
        .execute()
    )
    credits_used_month = (
        sum(abs(t["amount"]) for t in monthly_usage.data)
        if monthly_usage.data
        else 0
    )

    plan = row["plan"]
    rate_limits = {
        "requests_per_minute": {
            "free_trial": 10,
            "pay_as_you_go": 60,
            "pro": 120,
        }.get(plan, 60),
        "concurrent_conversions": (
            1 if plan in ("free_trial", "pay_as_you_go") else 2
        ),
    }

    return {
        "credits_available": row["credits_available"],
        "credits_reserved": row["credits_reserved"],
        "credits_used_this_month": credits_used_month,
        "plan": plan,
        "rate_limits": rate_limits,
        "top_up_url": f"{DASHBOARD_URL}/credits",
    }


# ─── Mount Point ─────────────────────────────────────────────────────────────


def create_mcp_app():
    """Create the MCP ASGI app ready to mount into FastAPI.

    Uses Streamable HTTP transport for cloud-hosted MCP access.

    Usage in server.py::

        from slidearabi.mcp_server import create_mcp_app
        app.mount("/mcp", create_mcp_app())
    """
    return mcp.http_app(path="/mcp", transport="streamable-http")
