"""
SlideArabi — Stripe Credit Purchase & Promo Code System
========================================================

Handles credit package purchases via Stripe Checkout, webhook processing,
promo code redemption, and signup bonus grants for the API/MCP platform.

Integration points:
  - Stripe Checkout for credit purchases
  - Supabase PostgreSQL for credit ledger (via RPC functions from 001_api_platform.sql)
  - FastAPI router for webhook endpoint

Usage:
    from slidearabi.stripe_credits import stripe_router, create_checkout_session
    app.include_router(stripe_router)
"""

from __future__ import annotations

import logging
import os
from typing import Any

import stripe
from fastapi import APIRouter, HTTPException, Request
from supabase import Client as SupabaseClient, create_client

logger = logging.getLogger("slidearabi.stripe_credits")

__all__ = [
    "CREDIT_PACKAGES",
    "PROMO_CODES",
    "create_checkout_session",
    "handle_checkout_completed",
    "handle_stripe_webhook",
    "redeem_promo_code",
    "grant_signup_bonus",
    "stripe_router",
]

# ─── Configuration ───────────────────────────────────────────────────────────

STRIPE_SECRET_KEY: str = os.environ.get("STRIPE_SECRET_KEY", "")
STRIPE_WEBHOOK_SECRET: str = os.environ.get("STRIPE_WEBHOOK_SECRET", "")
SUPABASE_URL: str = os.environ.get("SUPABASE_URL", "")
SUPABASE_SERVICE_KEY: str = os.environ.get("SUPABASE_SERVICE_KEY", "")

stripe.api_key = STRIPE_SECRET_KEY

# ─── Credit Packages ────────────────────────────────────────────────────────

CREDIT_PACKAGES: dict[str, dict[str, Any]] = {
    "starter_25": {
        "credits": 25,
        "price_cents": 2500,
        "name": "Starter",
        "description": "25 slide credits ($1.00/slide)",
    },
    "growth_100": {
        "credits": 100,
        "price_cents": 9000,
        "name": "Growth",
        "description": "100 slide credits ($0.90/slide — 10% savings)",
    },
    "studio_500": {
        "credits": 500,
        "price_cents": 37500,
        "name": "Studio",
        "description": "500 slide credits ($0.75/slide — 25% savings)",
    },
    "enterprise_2000": {
        "credits": 2000,
        "price_cents": 140000,
        "name": "Enterprise",
        "description": "2,000 slide credits ($0.70/slide — 30% savings)",
    },
}

# ─── Promo Codes ─────────────────────────────────────────────────────────────

PROMO_CODES: dict[str, int] = {
    "SLIDETEST2026": 25,
    "FOUNDER": 100,
    "DEMO": 50,
}

SIGNUP_BONUS_CREDITS: int = 10


# ─── Supabase Client ────────────────────────────────────────────────────────

def _get_supabase() -> SupabaseClient:
    """Get a Supabase client with service role key (bypasses RLS)."""
    if not SUPABASE_URL or not SUPABASE_SERVICE_KEY:
        raise RuntimeError("SUPABASE_URL and SUPABASE_SERVICE_KEY must be set")
    return create_client(SUPABASE_URL, SUPABASE_SERVICE_KEY)


# ─── Checkout Session ────────────────────────────────────────────────────────

async def create_checkout_session(
    account_id: str,
    package_id: str,
    success_url: str = "https://www.slidearabi.com/dashboard/credits?purchased=true",
    cancel_url: str = "https://www.slidearabi.com/dashboard/credits",
) -> dict[str, str]:
    """Create a Stripe Checkout Session for purchasing a credit package.

    Args:
        account_id: UUID of the api_accounts row.
        package_id: Key from CREDIT_PACKAGES (e.g., "growth_100").
        success_url: Redirect URL after successful payment.
        cancel_url: Redirect URL if user cancels.

    Returns:
        {"checkout_url": str, "session_id": str}

    Raises:
        HTTPException 400: If package_id is invalid.
    """
    package = CREDIT_PACKAGES.get(package_id)
    if not package:
        raise HTTPException(
            status_code=400,
            detail={
                "error": {
                    "code": "INVALID_PACKAGE",
                    "message": f"Unknown package: {package_id}. "
                    f"Valid packages: {', '.join(CREDIT_PACKAGES.keys())}",
                }
            },
        )

    sb = _get_supabase()

    # Look up the account and its Stripe customer ID
    result = (
        sb.table("api_accounts")
        .select("stripe_customer_id, user_id")
        .eq("id", account_id)
        .single()
        .execute()
    )
    if not result.data:
        raise HTTPException(status_code=404, detail="Account not found")

    stripe_customer_id: str | None = result.data.get("stripe_customer_id")

    # Create or retrieve Stripe Customer
    if not stripe_customer_id:
        user = sb.auth.admin.get_user_by_id(result.data["user_id"])
        customer = stripe.Customer.create(
            email=user.user.email if user.user else None,
            metadata={"slidearabi_account_id": account_id},
        )
        stripe_customer_id = customer.id
        sb.table("api_accounts").update(
            {"stripe_customer_id": stripe_customer_id}
        ).eq("id", account_id).execute()

    session = stripe.checkout.Session.create(
        customer=stripe_customer_id,
        mode="payment",
        line_items=[
            {
                "price_data": {
                    "currency": "usd",
                    "unit_amount": package["price_cents"],
                    "product_data": {
                        "name": package["name"],
                        "description": package["description"],
                    },
                },
                "quantity": 1,
            }
        ],
        metadata={
            "account_id": account_id,
            "package_id": package_id,
            "credits": str(package["credits"]),
        },
        success_url=success_url,
        cancel_url=cancel_url,
    )

    return {"checkout_url": session.url, "session_id": session.id}


# ─── Webhook Handlers ────────────────────────────────────────────────────────

async def handle_checkout_completed(event: dict[str, Any]) -> None:
    """Process a successful Stripe Checkout — grant credits to the account.

    Idempotent: checks if the payment_intent has already been credited.
    """
    session = event["data"]["object"]
    metadata = session.get("metadata", {})

    account_id: str | None = metadata.get("account_id")
    credits = int(metadata.get("credits", 0))
    package_id: str = metadata.get("package_id", "unknown")
    payment_intent: str | None = session.get("payment_intent")

    if not account_id or credits <= 0:
        logger.warning(
            "Checkout webhook missing account_id or credits: metadata=%s", metadata
        )
        return

    sb = _get_supabase()

    # Idempotency: skip if this payment_intent was already processed
    if payment_intent:
        existing = (
            sb.table("credit_transactions")
            .select("id")
            .eq("stripe_payment_intent", payment_intent)
            .limit(1)
            .execute()
        )
        if existing.data:
            logger.info(
                "Payment intent %s already processed — skipping duplicate",
                payment_intent,
            )
            return

    # Grant credits via the atomic PostgreSQL function
    sb.rpc(
        "grant_credits",
        {
            "p_account_id": account_id,
            "p_amount": credits,
            "p_description": f"Purchased {credits} credits ({package_id})",
            "p_type": "purchase",
            "p_stripe_pi": payment_intent,
        },
    ).execute()

    logger.info(
        "Granted %d credits to account %s (package=%s, PI=%s)",
        credits,
        account_id,
        package_id,
        payment_intent,
    )


# ─── Promo Code Redemption ───────────────────────────────────────────────────

async def redeem_promo_code(account_id: str, code: str) -> dict[str, Any]:
    """Redeem a promotional code for free credits.

    Each code can only be used once per account. Tracked via the
    credit_transactions description field.

    Supported codes:
        SLIDETEST2026 → 25 credits
        FOUNDER       → 100 credits
        DEMO          → 50 credits

    Args:
        account_id: UUID of the api_accounts row.
        code: Promo code string (case-insensitive).

    Returns:
        {"success": bool, "credits_granted": int, "message": str}
    """
    code_upper = code.strip().upper()
    credits = PROMO_CODES.get(code_upper)

    if credits is None:
        return {
            "success": False,
            "credits_granted": 0,
            "message": "Invalid promo code.",
        }

    sb = _get_supabase()

    # Check if this code was already redeemed by this account
    existing = (
        sb.table("credit_transactions")
        .select("id")
        .eq("account_id", account_id)
        .eq("type", "promo")
        .like("description", f"%{code_upper}%")
        .limit(1)
        .execute()
    )
    if existing.data:
        return {
            "success": False,
            "credits_granted": 0,
            "message": f"Promo code {code_upper} has already been redeemed on this account.",
        }

    # Grant the promo credits
    sb.rpc(
        "grant_credits",
        {
            "p_account_id": account_id,
            "p_amount": credits,
            "p_description": f"Promo code: {code_upper} — {credits} credits",
            "p_type": "promo",
            "p_stripe_pi": None,
        },
    ).execute()

    logger.info(
        "Redeemed promo %s for %d credits on account %s",
        code_upper,
        credits,
        account_id,
    )
    return {
        "success": True,
        "credits_granted": credits,
        "message": f"Promo code {code_upper} redeemed — {credits} credits added.",
    }


# ─── Signup Bonus ────────────────────────────────────────────────────────────

async def grant_signup_bonus(account_id: str) -> dict[str, Any]:
    """Grant free credits when a user creates their first API key.

    Idempotent: checks for an existing 'bonus' transaction before granting.
    Call this from the API key creation endpoint.

    Args:
        account_id: UUID of the api_accounts row.

    Returns:
        {"granted": bool, "credits": int}
    """
    sb = _get_supabase()

    # Check if bonus was already granted
    existing = (
        sb.table("credit_transactions")
        .select("id")
        .eq("account_id", account_id)
        .eq("type", "bonus")
        .limit(1)
        .execute()
    )
    if existing.data:
        return {"granted": False, "credits": 0}

    sb.rpc(
        "grant_credits",
        {
            "p_account_id": account_id,
            "p_amount": SIGNUP_BONUS_CREDITS,
            "p_description": f"Signup bonus — {SIGNUP_BONUS_CREDITS} free credits",
            "p_type": "bonus",
            "p_stripe_pi": None,
        },
    ).execute()

    logger.info(
        "Granted %d-credit signup bonus to account %s",
        SIGNUP_BONUS_CREDITS,
        account_id,
    )
    return {"granted": True, "credits": SIGNUP_BONUS_CREDITS}


# ─── FastAPI Webhook Endpoint ────────────────────────────────────────────────

stripe_router = APIRouter(tags=["stripe"])


async def handle_stripe_webhook(request: Request) -> dict[str, bool]:
    """Verify Stripe webhook signature and route events to handlers.

    This is the core webhook processing function. It verifies the signature
    using the STRIPE_WEBHOOK_SECRET, then dispatches to the appropriate handler.

    Args:
        request: The incoming FastAPI Request with raw body and Stripe headers.

    Returns:
        {"received": True} on success.

    Raises:
        HTTPException 400: If signature verification fails or payload is invalid.
    """
    payload = await request.body()
    sig_header = request.headers.get("stripe-signature", "")

    if not STRIPE_WEBHOOK_SECRET:
        logger.error("STRIPE_WEBHOOK_SECRET is not configured")
        raise HTTPException(status_code=500, detail="Webhook secret not configured")

    try:
        event = stripe.Webhook.construct_event(
            payload, sig_header, STRIPE_WEBHOOK_SECRET
        )
    except stripe.SignatureVerificationError:
        logger.warning("Invalid Stripe webhook signature")
        raise HTTPException(status_code=400, detail="Invalid signature")
    except ValueError:
        logger.error("Invalid Stripe webhook payload")
        raise HTTPException(status_code=400, detail="Invalid payload")

    # Route events to handlers
    event_type: str = event["type"]
    logger.info("Received Stripe event: %s (id=%s)", event_type, event.get("id"))

    if event_type == "checkout.session.completed":
        await handle_checkout_completed(event)
    else:
        logger.debug("Unhandled Stripe event type: %s", event_type)

    return {"received": True}


@stripe_router.post("/webhooks/stripe")
async def stripe_webhook_endpoint(request: Request) -> dict[str, bool]:
    """FastAPI route for Stripe webhooks.

    Mount on the main app::

        from slidearabi.stripe_credits import stripe_router
        app.include_router(stripe_router)

    Configure in Stripe Dashboard:
        Webhook URL: https://api.slidearabi.com/webhooks/stripe
        Events: checkout.session.completed
    """
    return await handle_stripe_webhook(request)
