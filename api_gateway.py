"""
SlideArabi — REST API Gateway (v1)
====================================

Provides the /v1/* REST API endpoints for programmatic access to SlideArabi.
Includes API key authentication, credit management, rate limiting, and
full conversion workflow.

Mount into the main FastAPI app:
    from api_gateway import api_router
    app.include_router(api_router, prefix="/v1")
"""

import hashlib
import logging
import os
import secrets
import time
from collections import defaultdict
from datetime import datetime, timedelta, timezone
from io import BytesIO
from typing import Any, Optional

import httpx
from fastapi import (
    APIRouter,
    Depends,
    File,
    Form,
    Header,
    HTTPException,
    Query,
    Request,
    Response,
    UploadFile,
)
from fastapi.responses import JSONResponse, RedirectResponse
from pydantic import BaseModel, Field
from supabase import Client as SupabaseClient, create_client

logger = logging.getLogger("slidearabi.api_gateway")

# ─── Configuration ───────────────────────────────────────────────────────────

SUPABASE_URL: str = os.environ.get("SUPABASE_URL", "")
SUPABASE_SERVICE_KEY: str = os.environ.get("SUPABASE_SERVICE_KEY", "")
INTERNAL_BACKEND_URL: str = os.environ.get(
    "INTERNAL_BACKEND_URL",
    "https://slidearabi-production.up.railway.app",
)

MAX_UPLOAD_SIZE: int = 50 * 1024 * 1024  # 50 MB for API uploads

# Rate limit tiers: requests per minute
RATE_LIMITS: dict[str, int] = {
    "free_trial": 10,
    "pay_as_you_go": 60,
    "pro": 120,
    "enterprise": 300,
}

# Concurrent conversion limits by tier
CONCURRENT_LIMITS: dict[str, int] = {
    "free_trial": 1,
    "pay_as_you_go": 1,
    "pro": 2,
    "enterprise": 5,
}


# ─── Supabase Client ────────────────────────────────────────────────────────


def _get_supabase() -> SupabaseClient:
    """Get Supabase client with service role key (bypasses RLS)."""
    if not SUPABASE_URL or not SUPABASE_SERVICE_KEY:
        raise HTTPException(
            status_code=503,
            detail=_error_response(
                "SERVICE_UNAVAILABLE",
                "Database configuration missing. Contact support.",
            ),
        )
    return create_client(SUPABASE_URL, SUPABASE_SERVICE_KEY)


# ─── ID Generators ───────────────────────────────────────────────────────────


def generate_api_key() -> tuple[str, str, str]:
    """
    Generate a new production API key.

    Returns:
        (full_key, key_hash, key_prefix)
        full_key:   "sa_live_a1b2c3d4e5f6g7h8i9j0k1l2m3n4o5p6"
        key_hash:   SHA-256 hex digest of the full key
        key_prefix: "sa_live_a1b2c3d4" (first 8 chars of random part, for display)
    """
    random_part = secrets.token_hex(16)  # 32 hex chars
    full_key = f"sa_live_{random_part}"
    key_hash = hashlib.sha256(full_key.encode()).hexdigest()
    key_prefix = f"sa_live_{random_part[:8]}"
    return full_key, key_hash, key_prefix


def generate_test_key() -> tuple[str, str, str]:
    """Generate a sandbox API key (sa_test_ prefix).

    Returns:
        (full_key, key_hash, key_prefix)
    """
    random_part = secrets.token_hex(16)
    full_key = f"sa_test_{random_part}"
    key_hash = hashlib.sha256(full_key.encode()).hexdigest()
    key_prefix = f"sa_test_{random_part[:8]}"
    return full_key, key_hash, key_prefix


def generate_job_id() -> str:
    """Generate a unique job ID: job_xxxxxxxx."""
    return f"job_{secrets.token_hex(4)}"


def generate_upload_id() -> str:
    """Generate a unique upload ID: upl_xxxxxxxx."""
    return f"upl_{secrets.token_hex(4)}"


def generate_request_id() -> str:
    """Generate a unique request ID for error tracing."""
    return f"req_{secrets.token_hex(6)}"


# ─── Error Response Helper ───────────────────────────────────────────────────


def _error_response(
    code: str,
    message: str,
    details: Optional[dict[str, Any]] = None,
) -> dict[str, Any]:
    """Build a standardized error response.

    Format: {"error": {"code": "...", "message": "...", "details": {...}, "request_id": "..."}}
    """
    err: dict[str, Any] = {
        "error": {
            "code": code,
            "message": message,
            "request_id": generate_request_id(),
        }
    }
    if details:
        err["error"]["details"] = details
    return err


# ─── Rate Limiter (in-memory sliding window) ─────────────────────────────────


class RateLimiter:
    """In-memory sliding window rate limiter.

    Tracks timestamps of requests per key within a configurable window.
    No external dependencies (Redis-free for Phase 1).
    """

    def __init__(self) -> None:
        self._windows: dict[str, list[float]] = defaultdict(list)

    def check(
        self, key: str, limit: int, window_seconds: int = 60
    ) -> tuple[bool, int, int]:
        """Check if a request is within rate limits.

        Args:
            key: Unique identifier for the rate limit bucket (e.g., account_id).
            limit: Maximum requests allowed in the window.
            window_seconds: Sliding window duration in seconds.

        Returns:
            (allowed, remaining, reset_timestamp)
        """
        now = time.time()
        window_start = now - window_seconds

        # Prune entries older than the window
        self._windows[key] = [t for t in self._windows[key] if t > window_start]

        if len(self._windows[key]) >= limit:
            reset_at = int(self._windows[key][0] + window_seconds)
            return False, 0, reset_at

        self._windows[key].append(now)
        remaining = limit - len(self._windows[key])
        reset_at = int(now + window_seconds)
        return True, remaining, reset_at


rate_limiter = RateLimiter()


# ─── Credit System ───────────────────────────────────────────────────────────


async def reserve_credits(account_id: str, amount: int, job_id: str) -> bool:
    """Atomically reserve credits for a conversion job.

    Checks that credits_available >= amount, then moves the specified amount
    from available to reserved. Writes a 'reserve' entry to the
    credit_transactions ledger.

    Uses the PostgreSQL ``reserve_credits`` RPC function for atomicity.

    Args:
        account_id: UUID of the api_accounts row.
        amount: Number of credits to reserve.
        job_id: Associated job ID for the ledger entry.

    Returns:
        True if reservation succeeded, False if insufficient credits.
    """
    sb = _get_supabase()
    result = sb.rpc(
        "reserve_credits",
        {
            "p_account_id": account_id,
            "p_amount": amount,
            "p_job_id": job_id,
        },
    ).execute()

    success = result.data is not False and result.data is not None
    if success:
        logger.info(
            "Reserved %d credits for job %s (account %s)",
            amount, job_id, account_id,
        )
    else:
        logger.warning(
            "Credit reservation failed for job %s: need %d (account %s)",
            job_id, amount, account_id,
        )
    return success


async def settle_credits(account_id: str, job_id: str, amount: int) -> None:
    """Settle (finalize) reserved credits after a successful conversion.

    Deducts the specified amount from reserved credits and writes a 'charge'
    entry to the credit_transactions ledger.

    Uses the PostgreSQL ``settle_credits`` RPC function for atomicity.

    Args:
        account_id: UUID of the api_accounts row.
        job_id: The completed job ID.
        amount: Number of credits to settle (deduct from reserved).
    """
    sb = _get_supabase()
    try:
        sb.rpc(
            "settle_credits",
            {
                "p_account_id": account_id,
                "p_amount": amount,
                "p_job_id": job_id,
            },
        ).execute()
        logger.info(
            "Settled %d credits for job %s (account %s)",
            amount, job_id, account_id,
        )
    except Exception as exc:
        logger.error("Failed to settle credits for job %s: %s", job_id, exc)
        raise


async def release_credits(account_id: str, amount: int, job_id: str) -> None:
    """Release reserved credits back to available on failure or cancellation.

    Moves the specified amount from reserved back to available and writes a
    'release' entry to the credit_transactions ledger.

    Uses the PostgreSQL ``release_credits`` RPC function for atomicity.

    Args:
        account_id: UUID of the api_accounts row.
        amount: Number of credits to release.
        job_id: The failed/cancelled job ID.
    """
    sb = _get_supabase()
    try:
        sb.rpc(
            "release_credits",
            {
                "p_account_id": account_id,
                "p_amount": amount,
                "p_job_id": job_id,
            },
        ).execute()
        logger.info(
            "Released %d credits for job %s (account %s)",
            amount, job_id, account_id,
        )
    except Exception as exc:
        logger.error("Failed to release credits for job %s: %s", job_id, exc)
        raise


# ─── Authentication Dependency ───────────────────────────────────────────────


class AuthContext(BaseModel):
    """Authenticated API request context."""

    account_id: str
    api_key_id: str
    plan_tier: str
    is_test: bool = False


async def require_api_key(
    authorization: Optional[str] = Header(None),
) -> AuthContext:
    """FastAPI dependency that validates the API key from the Authorization header.

    Validates ``Authorization: Bearer sa_live_xxxxxxxxxxxx`` (or ``sa_test_``).
    Looks up the key by SHA-256 hash in the Supabase ``api_keys`` table.
    Rejects with 401 if invalid or revoked. Updates ``last_used_at`` on success.

    Usage::

        @router.get("/endpoint")
        async def my_endpoint(auth: AuthContext = Depends(require_api_key)):
            ...
    """
    if not authorization:
        raise HTTPException(
            status_code=401,
            detail=_error_response(
                "INVALID_API_KEY",
                "Missing Authorization header. Expected: Bearer sa_live_xxx",
            ),
        )

    parts = authorization.split(" ", 1)
    if len(parts) != 2 or parts[0].lower() != "bearer":
        raise HTTPException(
            status_code=401,
            detail=_error_response(
                "INVALID_API_KEY",
                "Authorization header must be: Bearer <api_key>",
            ),
        )

    api_key = parts[1].strip()
    if not (api_key.startswith("sa_live_") or api_key.startswith("sa_test_")):
        raise HTTPException(
            status_code=401,
            detail=_error_response(
                "INVALID_API_KEY",
                "API key must start with sa_live_ or sa_test_",
            ),
        )

    key_hash = hashlib.sha256(api_key.encode()).hexdigest()
    sb = _get_supabase()

    result = (
        sb.table("api_keys")
        .select("id, account_id, is_test, is_active")
        .eq("key_hash", key_hash)
        .execute()
    )

    if not result.data:
        raise HTTPException(
            status_code=401,
            detail=_error_response("INVALID_API_KEY", "API key not found."),
        )

    key_row = result.data[0]
    if not key_row["is_active"]:
        raise HTTPException(
            status_code=401,
            detail=_error_response("KEY_REVOKED", "This API key has been revoked."),
        )

    # Update last_used_at (fire-and-forget, non-critical)
    try:
        sb.table("api_keys").update(
            {"last_used_at": datetime.now(timezone.utc).isoformat()}
        ).eq("id", key_row["id"]).execute()
    except Exception:
        pass

    # Fetch account plan tier
    account = (
        sb.table("api_accounts")
        .select("plan")
        .eq("id", key_row["account_id"])
        .single()
        .execute()
    )
    plan = account.data.get("plan", "pay_as_you_go") if account.data else "pay_as_you_go"

    return AuthContext(
        account_id=key_row["account_id"],
        api_key_id=key_row["id"],
        plan_tier=plan,
        is_test=key_row.get("is_test", False),
    )


# ─── Rate Limit Dependency ───────────────────────────────────────────────────


async def apply_rate_limit(auth: AuthContext, response: Response) -> None:
    """Apply sliding-window rate limiting based on the account's plan tier.

    Sets X-RateLimit-* headers on every response. Raises 429 if exceeded.
    """
    limit = RATE_LIMITS.get(auth.plan_tier, 60)
    allowed, remaining, reset_at = rate_limiter.check(auth.account_id, limit)

    response.headers["X-RateLimit-Limit"] = str(limit)
    response.headers["X-RateLimit-Remaining"] = str(remaining)
    response.headers["X-RateLimit-Reset"] = str(reset_at)

    if not allowed:
        raise HTTPException(
            status_code=429,
            detail=_error_response(
                "RATE_LIMIT_EXCEEDED",
                f"Rate limit exceeded. Limit: {limit}/min.",
            ),
            headers={"Retry-After": str(max(1, reset_at - int(time.time())))},
        )


# ─── Internal Backend Client ─────────────────────────────────────────────────


async def _call_backend(method: str, path: str, **kwargs: Any) -> httpx.Response:
    """Call the internal SlideArabi conversion backend.

    Args:
        method: HTTP method (GET, POST, DELETE, etc.).
        path: URL path (e.g., ``/convert``, ``/status/{job_id}``).
        **kwargs: Passed through to ``httpx.AsyncClient.request``.

    Returns:
        The httpx.Response from the backend.
    """
    url = f"{INTERNAL_BACKEND_URL}{path}"
    async with httpx.AsyncClient(timeout=300.0) as client:
        return await client.request(method, url, **kwargs)


# ─── Webhook Delivery ────────────────────────────────────────────────────────


async def _deliver_webhook(webhook_url: str, payload: dict[str, Any]) -> None:
    """Deliver a webhook notification to the caller's URL.

    Best-effort with one retry. Logs errors but does not raise.

    Args:
        webhook_url: The URL to POST the payload to.
        payload: JSON-serializable dict with job status information.
    """
    for attempt in range(2):
        try:
            async with httpx.AsyncClient(timeout=10.0) as client:
                resp = await client.post(webhook_url, json=payload)
                if resp.status_code < 400:
                    logger.info(
                        "Webhook delivered to %s (status %d)",
                        webhook_url, resp.status_code,
                    )
                    return
                logger.warning(
                    "Webhook to %s returned %d (attempt %d)",
                    webhook_url, resp.status_code, attempt + 1,
                )
        except Exception as exc:
            logger.warning(
                "Webhook delivery failed to %s (attempt %d): %s",
                webhook_url, attempt + 1, exc,
            )


# ─── Request/Response Models ─────────────────────────────────────────────────


class ConversionRequest(BaseModel):
    """Request body for POST /v1/conversions."""

    upload_id: str
    options: Optional[dict[str, Any]] = None
    webhook_url: Optional[str] = None
    idempotency_key: Optional[str] = None


class CreateApiKeyRequest(BaseModel):
    """Request body for POST /v1/api-keys."""

    name: str = "default"
    is_test: bool = False


# ─── Router ──────────────────────────────────────────────────────────────────

api_router = APIRouter(tags=["v1"])


# ─── Upload Endpoints ────────────────────────────────────────────────────────


@api_router.post("/uploads", status_code=201)
async def create_upload(
    response: Response,
    auth: AuthContext = Depends(require_api_key),
    file: Optional[UploadFile] = File(None),
    file_url: Optional[str] = Form(None),
) -> dict[str, Any]:
    """Upload a .pptx file for conversion.

    Accepts either:
    - Multipart file upload (Content-Type: multipart/form-data)
    - Form field ``file_url`` pointing to a downloadable .pptx

    Returns upload_id, slide count, and credits required.
    """
    await apply_rate_limit(auth, response)

    # Resolve file bytes
    if file:
        if not file.filename or not file.filename.lower().endswith(".pptx"):
            raise HTTPException(
                400, _error_response("INVALID_FILE", "File must be a .pptx PowerPoint file.")
            )
        file_bytes = await file.read()
        filename = file.filename
    elif file_url:
        try:
            async with httpx.AsyncClient(timeout=60.0) as client:
                dl = await client.get(file_url)
                dl.raise_for_status()
                file_bytes = dl.content
                filename = file_url.split("/")[-1].split("?")[0] or "presentation.pptx"
        except Exception as exc:
            raise HTTPException(
                400, _error_response("INVALID_FILE", f"Could not download file from URL: {exc}")
            )
    else:
        raise HTTPException(
            400, _error_response("INVALID_FILE", "Provide either a file upload or file_url.")
        )

    if len(file_bytes) > MAX_UPLOAD_SIZE:
        raise HTTPException(
            413, _error_response("FILE_TOO_LARGE", "File exceeds 50 MB limit.")
        )

    # Count slides
    try:
        from pptx import Presentation
        prs = Presentation(BytesIO(file_bytes))
        slide_count = len(prs.slides)
    except Exception:
        raise HTTPException(
            400, _error_response("INVALID_FILE", "Could not parse .pptx file. Ensure it is a valid PowerPoint.")
        )

    if slide_count < 1:
        raise HTTPException(
            400, _error_response("INVALID_FILE", "Presentation has no slides.")
        )

    credits_required = max(5, slide_count)

    # Check current credit balance for informational response
    sb = _get_supabase()
    account = (
        sb.table("api_accounts")
        .select("credits_available")
        .eq("id", auth.account_id)
        .single()
        .execute()
    )
    credits_available = account.data.get("credits_available", 0) if account.data else 0

    # Stage file in Supabase Storage
    upload_id = generate_upload_id()
    storage_path = f"api-uploads/{auth.account_id}/{upload_id}/{filename}"

    try:
        sb.storage.from_("api-files").upload(storage_path, file_bytes)
    except Exception as exc:
        logger.error("Failed to upload file to storage: %s", exc)
        try:
            sb.storage.create_bucket("api-files", options={"public": False})
            sb.storage.from_("api-files").upload(storage_path, file_bytes)
        except Exception:
            raise HTTPException(
                500, _error_response("PROCESSING_FAILED", "File storage unavailable.")
            )

    expires_at = (datetime.now(timezone.utc) + timedelta(hours=24)).isoformat()

    sb.table("api_uploads").insert({
        "id": upload_id,
        "account_id": auth.account_id,
        "filename": filename,
        "slide_count": slide_count,
        "size_bytes": len(file_bytes),
        "storage_path": storage_path,
        "expires_at": expires_at,
    }).execute()

    return {
        "upload_id": upload_id,
        "filename": filename,
        "slide_count": slide_count,
        "size_bytes": len(file_bytes),
        "credits_required": credits_required,
        "credits_available": credits_available,
        "expires_at": expires_at,
    }


# ─── Conversion Endpoints ────────────────────────────────────────────────────


@api_router.post("/conversions", status_code=202)
async def create_conversion(
    body: ConversionRequest,
    response: Response,
    auth: AuthContext = Depends(require_api_key),
) -> dict[str, Any]:
    """Start an English-to-Arabic conversion.

    Reserves credits atomically, triggers the internal conversion pipeline,
    and returns the job_id immediately (202 Accepted).
    """
    await apply_rate_limit(auth, response)

    sb = _get_supabase()

    # Idempotency check
    if body.idempotency_key:
        existing = (
            sb.table("api_jobs")
            .select("id, status, slide_count, credits_reserved")
            .eq("idempotency_key", body.idempotency_key)
            .execute()
        )
        if existing.data:
            job = existing.data[0]
            return {
                "job_id": job["id"],
                "status": job["status"],
                "slide_count": job["slide_count"],
                "credits_reserved": job["credits_reserved"],
                "message": "Existing job returned (idempotency key match).",
            }

    # Validate upload exists and belongs to this account
    upload = (
        sb.table("api_uploads")
        .select("*")
        .eq("id", body.upload_id)
        .eq("account_id", auth.account_id)
        .execute()
    )
    if not upload.data:
        raise HTTPException(
            404, _error_response("UPLOAD_NOT_FOUND", f"Upload {body.upload_id} not found.")
        )

    upload_row = upload.data[0]
    expires_at_str = upload_row["expires_at"].replace("Z", "+00:00")
    if datetime.fromisoformat(expires_at_str) < datetime.now(timezone.utc):
        raise HTTPException(
            410, _error_response("UPLOAD_EXPIRED", "This upload has expired. Please upload again.")
        )

    slide_count: int = upload_row["slide_count"]
    credits_needed = max(5, slide_count)
    job_id = generate_job_id()

    # Reserve credits (skip for test/sandbox keys)
    if not auth.is_test:
        reserved = await reserve_credits(auth.account_id, credits_needed, job_id)
        if not reserved:
            account = (
                sb.table("api_accounts")
                .select("credits_available")
                .eq("id", auth.account_id)
                .single()
                .execute()
            )
            avail = account.data.get("credits_available", 0) if account.data else 0
            raise HTTPException(
                402,
                _error_response(
                    "INSUFFICIENT_CREDITS",
                    f"You have {avail} credits but this conversion requires {credits_needed}.",
                    {
                        "credits_available": avail,
                        "credits_required": credits_needed,
                        "top_up_url": "https://www.slidearabi.com/dashboard/credits",
                    },
                ),
            )

    # Create job record in Supabase
    sb.table("api_jobs").insert({
        "id": job_id,
        "account_id": auth.account_id,
        "upload_id": body.upload_id,
        "source": "api",
        "status": "queued",
        "slide_count": slide_count,
        "credits_reserved": credits_needed,
        "options": body.options or {},
        "webhook_url": body.webhook_url,
        "idempotency_key": body.idempotency_key,
    }).execute()

    # Trigger conversion on the internal backend
    internal_job_id: Optional[str] = None
    try:
        file_bytes = sb.storage.from_("api-files").download(upload_row["storage_path"])
        backend_resp = await _call_backend(
            "POST", "/convert",
            files={
                "file": (
                    upload_row["filename"],
                    file_bytes,
                    "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                )
            },
        )
        if backend_resp.status_code == 200:
            backend_data = backend_resp.json()
            internal_job_id = backend_data.get("job_id")
            sb.table("api_jobs").update({
                "status": "processing",
                "internal_job_id": internal_job_id,
                "started_at": datetime.now(timezone.utc).isoformat(),
            }).eq("id", job_id).execute()
        else:
            logger.warning(
                "Backend returned %d for job %s: %s",
                backend_resp.status_code, job_id, backend_resp.text[:200],
            )
    except Exception as exc:
        logger.error("Failed to trigger backend conversion for job %s: %s", job_id, exc)

    # Fetch updated credit balance for response
    account = (
        sb.table("api_accounts")
        .select("credits_available, credits_reserved")
        .eq("id", auth.account_id)
        .single()
        .execute()
    )

    return {
        "job_id": job_id,
        "status": "queued",
        "slide_count": slide_count,
        "credits_reserved": credits_needed,
        "credits_remaining": account.data.get("credits_available", 0) if account.data else 0,
        "estimated_seconds": slide_count * 15,  # ~15s per slide
        "poll_url": f"/v1/conversions/{job_id}",
        "created_at": datetime.now(timezone.utc).isoformat(),
    }


@api_router.get("/conversions/{job_id}")
async def get_conversion_status(
    job_id: str,
    response: Response,
    auth: AuthContext = Depends(require_api_key),
) -> dict[str, Any]:
    """Get the current status and progress of a conversion job.

    Proxies to the internal /status endpoint for live progress and enriches
    the response with credit information.
    """
    await apply_rate_limit(auth, response)

    sb = _get_supabase()
    job = (
        sb.table("api_jobs")
        .select("*")
        .eq("id", job_id)
        .eq("account_id", auth.account_id)
        .execute()
    )

    if not job.data:
        raise HTTPException(
            404, _error_response("JOB_NOT_FOUND", f"Job {job_id} not found.")
        )

    row = job.data[0]

    # Proxy to internal backend for live progress if job is active
    internal_status: Optional[dict[str, Any]] = None
    internal_job_id = row.get("internal_job_id")
    if internal_job_id and row["status"] in ("queued", "processing"):
        try:
            backend_resp = await _call_backend("GET", f"/status/{internal_job_id}")
            if backend_resp.status_code == 200:
                internal_status = backend_resp.json()
                backend_status = internal_status.get("status")

                # Sync terminal state: completed
                if backend_status == "completed" and row["status"] != "done":
                    sb.table("api_jobs").update({
                        "status": "done",
                        "progress_percent": 100,
                        "current_phase": "done",
                        "completed_at": datetime.now(timezone.utc).isoformat(),
                    }).eq("id", job_id).execute()
                    row["status"] = "done"
                    row["progress_percent"] = 100
                    row["current_phase"] = "done"

                    # Settle credits on completion
                    if row.get("credits_reserved", 0) > 0:
                        try:
                            await settle_credits(
                                auth.account_id, job_id, row["credits_reserved"],
                            )
                            sb.table("api_jobs").update(
                                {"credits_charged": row["credits_reserved"]}
                            ).eq("id", job_id).execute()
                        except Exception as exc:
                            logger.error(
                                "Credit settlement failed for %s: %s", job_id, exc,
                            )

                    # Deliver webhook
                    if row.get("webhook_url"):
                        await _deliver_webhook(row["webhook_url"], {
                            "event": "conversion.completed",
                            "job_id": job_id,
                            "status": "done",
                            "download_url": f"/v1/conversions/{job_id}/download",
                        })

                # Sync terminal state: failed
                elif backend_status == "failed" and row["status"] != "failed":
                    error_msg = internal_status.get("error", "Conversion failed")
                    sb.table("api_jobs").update({
                        "status": "failed",
                        "error_message": error_msg,
                        "error_code": "CONVERSION_FAILED",
                    }).eq("id", job_id).execute()
                    row["status"] = "failed"

                    # Release credits on failure
                    if row.get("credits_reserved", 0) > 0:
                        try:
                            await release_credits(
                                auth.account_id, row["credits_reserved"], job_id,
                            )
                        except Exception as exc:
                            logger.error(
                                "Credit release failed for %s: %s", job_id, exc,
                            )

                    # Deliver failure webhook
                    if row.get("webhook_url"):
                        await _deliver_webhook(row["webhook_url"], {
                            "event": "conversion.failed",
                            "job_id": job_id,
                            "status": "failed",
                            "error": error_msg,
                        })

        except Exception as exc:
            logger.warning("Could not proxy status for job %s: %s", job_id, exc)

    result: dict[str, Any] = {
        "job_id": row["id"],
        "status": row["status"],
        "progress_percent": (
            internal_status.get("progress_pct", row.get("progress_percent", 0))
            if internal_status
            else row.get("progress_percent", 0)
        ),
        "current_phase": (
            internal_status.get("current_phase", row.get("current_phase"))
            if internal_status
            else row.get("current_phase")
        ),
        "slide_count": row["slide_count"],
        "credits_reserved": row.get("credits_reserved", 0),
        "credits_charged": row.get("credits_charged", 0),
        "created_at": row.get("created_at"),
        "started_at": row.get("started_at"),
        "completed_at": row.get("completed_at"),
    }

    if row["status"] == "done":
        result["download_url"] = f"/v1/conversions/{job_id}/download"
    elif row["status"] == "failed":
        result["error_code"] = row.get("error_code")
        result["error_message"] = row.get("error_message")

    # Credit balance in response header
    try:
        acct = (
            sb.table("api_accounts")
            .select("credits_available")
            .eq("id", auth.account_id)
            .single()
            .execute()
        )
        response.headers["X-Credits-Remaining"] = str(
            acct.data.get("credits_available", 0) if acct.data else 0
        )
    except Exception:
        pass

    return result


@api_router.get("/conversions/{job_id}/download")
async def download_conversion(
    job_id: str,
    response: Response,
    auth: AuthContext = Depends(require_api_key),
) -> Any:
    """Download the completed Arabic .pptx file.

    Returns the file directly or redirects to a signed Supabase Storage URL.
    """
    await apply_rate_limit(auth, response)

    sb = _get_supabase()
    job = (
        sb.table("api_jobs")
        .select("status, result_path, internal_job_id, account_id")
        .eq("id", job_id)
        .eq("account_id", auth.account_id)
        .execute()
    )

    if not job.data:
        raise HTTPException(
            404, _error_response("JOB_NOT_FOUND", f"Job {job_id} not found.")
        )

    row = job.data[0]
    if row["status"] != "done":
        raise HTTPException(
            409, _error_response("JOB_NOT_READY", f"Job is not complete. Current status: {row['status']}")
        )

    # Prefer signed URL from Supabase Storage
    if row.get("result_path"):
        signed = sb.storage.from_("api-files").create_signed_url(row["result_path"], 3600)
        return RedirectResponse(signed["signedURL"], status_code=302)

    # Fall back to proxying from the internal backend
    internal_job_id = row.get("internal_job_id")
    if internal_job_id:
        backend_resp = await _call_backend("GET", f"/download/{internal_job_id}")
        if backend_resp.status_code == 200:
            return Response(
                content=backend_resp.content,
                media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                headers={"Content-Disposition": f'attachment; filename="{job_id}_arabic.pptx"'},
            )

    raise HTTPException(
        500, _error_response("PROCESSING_FAILED", "Download not available.")
    )


@api_router.get("/conversions")
async def list_conversions(
    response: Response,
    auth: AuthContext = Depends(require_api_key),
    status: Optional[str] = Query(None),
    limit: int = Query(20, ge=1, le=100),
    offset: int = Query(0, ge=0),
) -> dict[str, Any]:
    """List past conversion jobs for the authenticated account (paginated)."""
    await apply_rate_limit(auth, response)

    sb = _get_supabase()
    query = (
        sb.table("api_jobs")
        .select("id, status, slide_count, credits_charged, created_at, completed_at")
        .eq("account_id", auth.account_id)
    )

    if status:
        query = query.eq("status", status)

    result = query.order("created_at", desc=True).range(offset, offset + limit - 1).execute()

    return {
        "conversions": result.data,
        "count": len(result.data),
        "offset": offset,
        "limit": limit,
    }


@api_router.delete("/conversions/{job_id}")
async def cancel_conversion(
    job_id: str,
    response: Response,
    auth: AuthContext = Depends(require_api_key),
) -> dict[str, Any]:
    """Cancel a queued conversion and release reserved credits."""
    await apply_rate_limit(auth, response)

    sb = _get_supabase()
    job = (
        sb.table("api_jobs")
        .select("status, credits_reserved, account_id")
        .eq("id", job_id)
        .eq("account_id", auth.account_id)
        .execute()
    )

    if not job.data:
        raise HTTPException(
            404, _error_response("JOB_NOT_FOUND", f"Job {job_id} not found.")
        )

    row = job.data[0]
    if row["status"] != "queued":
        raise HTTPException(
            409, _error_response(
                "INVALID_STATE",
                f"Can only cancel queued jobs. Current status: {row['status']}",
            )
        )

    # Release reserved credits
    credits_refunded = row.get("credits_reserved", 0)
    if credits_refunded > 0:
        await release_credits(auth.account_id, credits_refunded, job_id)

    sb.table("api_jobs").update({"status": "cancelled"}).eq("id", job_id).execute()

    return {
        "job_id": job_id,
        "status": "cancelled",
        "credits_refunded": credits_refunded,
    }


# ─── Account Endpoints ───────────────────────────────────────────────────────


@api_router.get("/account")
async def get_account(
    response: Response,
    auth: AuthContext = Depends(require_api_key),
) -> dict[str, Any]:
    """Get account info, plan, and rate limits."""
    await apply_rate_limit(auth, response)

    sb = _get_supabase()
    account = (
        sb.table("api_accounts")
        .select("*")
        .eq("id", auth.account_id)
        .single()
        .execute()
    )

    if not account.data:
        raise HTTPException(
            404, _error_response("ACCOUNT_NOT_FOUND", "Account not found.")
        )

    row = account.data
    rpm = RATE_LIMITS.get(row["plan"], 60)
    concurrent = CONCURRENT_LIMITS.get(row["plan"], 1)

    return {
        "account_id": row["id"],
        "plan": row["plan"],
        "credits_available": row["credits_available"],
        "credits_reserved": row["credits_reserved"],
        "rate_limits": {
            "requests_per_minute": rpm,
            "concurrent_conversions": concurrent,
        },
        "auto_topup": {
            "enabled": row.get("auto_topup_enabled", False),
            "threshold": row.get("auto_topup_threshold", 20),
            "amount": row.get("auto_topup_amount", 100),
        },
        "created_at": row["created_at"],
    }


@api_router.get("/account/balance")
async def get_balance(
    response: Response,
    auth: AuthContext = Depends(require_api_key),
) -> dict[str, Any]:
    """Get current credit balance (available + reserved)."""
    await apply_rate_limit(auth, response)

    sb = _get_supabase()
    account = (
        sb.table("api_accounts")
        .select("credits_available, credits_reserved")
        .eq("id", auth.account_id)
        .single()
        .execute()
    )

    if not account.data:
        raise HTTPException(
            404, _error_response("ACCOUNT_NOT_FOUND", "Account not found.")
        )

    return {
        "credits_available": account.data["credits_available"],
        "credits_reserved": account.data["credits_reserved"],
        "top_up_url": "https://www.slidearabi.com/dashboard/credits",
    }


# ─── Credit Transaction History ──────────────────────────────────────────────


@api_router.get("/credits/transactions")
async def get_credit_transactions(
    response: Response,
    auth: AuthContext = Depends(require_api_key),
    limit: int = Query(20, ge=1, le=100),
    offset: int = Query(0, ge=0),
) -> dict[str, Any]:
    """Get paginated credit transaction history."""
    await apply_rate_limit(auth, response)

    sb = _get_supabase()
    result = (
        sb.table("credit_transactions")
        .select("id, type, amount, balance_after, description, job_id, created_at")
        .eq("account_id", auth.account_id)
        .order("created_at", desc=True)
        .range(offset, offset + limit - 1)
        .execute()
    )

    return {
        "transactions": result.data,
        "count": len(result.data),
        "offset": offset,
        "limit": limit,
    }


# ─── API Key Management ─────────────────────────────────────────────────────
# Note: In production, POST/DELETE on api-keys should require Supabase JWT
# auth (from the web dashboard), not API key auth. For Phase 1, we allow
# API key auth so the endpoints are functional end-to-end.


@api_router.post("/api-keys", status_code=201)
async def create_api_key_endpoint(
    body: CreateApiKeyRequest,
    response: Response,
    auth: AuthContext = Depends(require_api_key),
) -> dict[str, Any]:
    """Create a new API key. The full key is returned ONCE and never shown again.

    Maximum 5 active keys per account.
    """
    await apply_rate_limit(auth, response)

    sb = _get_supabase()

    # Enforce key limit
    existing = (
        sb.table("api_keys")
        .select("id")
        .eq("account_id", auth.account_id)
        .eq("is_active", True)
        .execute()
    )
    if len(existing.data) >= 5:
        raise HTTPException(
            400, _error_response("KEY_LIMIT_EXCEEDED", "Maximum 5 active API keys per account.")
        )

    if body.is_test:
        full_key, key_hash, key_prefix = generate_test_key()
    else:
        full_key, key_hash, key_prefix = generate_api_key()

    sb.table("api_keys").insert({
        "account_id": auth.account_id,
        "key_hash": key_hash,
        "key_prefix": key_prefix,
        "name": body.name,
        "is_test": body.is_test,
    }).execute()

    # Grant signup bonus on first key creation
    try:
        from stripe_credits import grant_signup_bonus
        await grant_signup_bonus(auth.account_id)
    except ImportError:
        logger.debug("stripe_credits module not available for signup bonus")

    return {
        "api_key": full_key,  # Only time the full key is shown
        "key_prefix": key_prefix,
        "name": body.name,
        "is_test": body.is_test,
        "message": "Save this key now — it will not be shown again.",
    }


@api_router.get("/api-keys")
async def list_api_keys(
    response: Response,
    auth: AuthContext = Depends(require_api_key),
) -> dict[str, Any]:
    """List all API keys for the account (key values are masked)."""
    await apply_rate_limit(auth, response)

    sb = _get_supabase()
    result = (
        sb.table("api_keys")
        .select("id, key_prefix, name, is_test, is_active, last_used_at, created_at")
        .eq("account_id", auth.account_id)
        .order("created_at", desc=True)
        .execute()
    )

    return {"api_keys": result.data}


@api_router.delete("/api-keys/{key_id}")
async def revoke_api_key(
    key_id: str,
    response: Response,
    auth: AuthContext = Depends(require_api_key),
) -> dict[str, Any]:
    """Revoke an API key (soft delete — sets is_active=False)."""
    await apply_rate_limit(auth, response)

    sb = _get_supabase()
    result = (
        sb.table("api_keys")
        .update({"is_active": False})
        .eq("id", key_id)
        .eq("account_id", auth.account_id)
        .execute()
    )

    if not result.data:
        raise HTTPException(
            404, _error_response("KEY_NOT_FOUND", "API key not found.")
        )

    return {"key_id": key_id, "status": "revoked"}
