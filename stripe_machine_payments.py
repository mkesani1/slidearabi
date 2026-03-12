"""
SlideArabi — Stripe Machine Payments (x402 + USDC)
=====================================================

Handles agent-initiated USDC payments via Stripe's machine payments API.
Works alongside the existing Stripe Checkout (card) flow for human users.

When an AI agent pays via x402:
  1. Stripe creates a PaymentIntent with payment_method_types=["crypto"]
  2. Agent sends USDC to the generated deposit address on Base
  3. Stripe settles the payment → lands in your normal Stripe balance as USD
  4. Webhook fires → we grant credits to the agent's account

This module provides:
  - PaymentIntent creation for agent purchases
  - Webhook handler for crypto payment completion
  - Credit auto-provisioning for x402 payments

Integration::

    from slidearabi.stripe_machine_payments import machine_router
    app.include_router(machine_router)

Environment::

    STRIPE_SECRET_KEY         — existing Stripe key
    STRIPE_WEBHOOK_SECRET     — existing webhook secret (handles both card + crypto)
    STRIPE_MACHINE_PAYMENTS   — "true" to enable (requires Stripe approval)
"""

from __future__ import annotations

import logging
import math
import os
from typing import Any

import stripe
from fastapi import APIRouter, HTTPException, Request
from pydantic import BaseModel
from supabase import Client as SupabaseClient, create_client

logger = logging.getLogger("slidearabi.machine_payments")

# ─── Configuration ───────────────────────────────────────────────────────────

STRIPE_SECRET_KEY: str = os.environ.get("STRIPE_SECRET_KEY", "")
SUPABASE_URL: str = os.environ.get("SUPABASE_URL", "")
SUPABASE_SERVICE_KEY: str = os.environ.get("SUPABASE_SERVICE_KEY", "")
MACHINE_PAYMENTS_ENABLED: bool = (
    os.environ.get("STRIPE_MACHINE_PAYMENTS", "false").lower() == "true"
)

stripe.api_key = STRIPE_SECRET_KEY

# Credit pricing in cents — matches existing packages
# $1.00 per credit (same as starter tier)
CENTS_PER_CREDIT: int = 100

# Minimum and maximum credit purchases via machine payments
MIN_CREDITS: int = 1
MAX_CREDITS: int = 500


# ─── Supabase ────────────────────────────────────────────────────────────────


def _get_supabase() -> SupabaseClient:
    if not SUPABASE_URL or not SUPABASE_SERVICE_KEY:
        raise RuntimeError("Supabase credentials not configured")
    return create_client(SUPABASE_URL, SUPABASE_SERVICE_KEY)


# ─── Request Models ──────────────────────────────────────────────────────────


class MachinePaymentRequest(BaseModel):
    """Request body for creating a machine payment."""

    credits: int
    """Number of credits to purchase (1-500)."""

    account_id: str | None = None
    """Optional: link to an existing SlideArabi account.
    If not provided, credits are associated with the payment
    and can be claimed later."""

    metadata: dict[str, str] | None = None
    """Optional: additional metadata for the payment."""


class MachinePaymentResponse(BaseModel):
    """Response with deposit details for the agent."""

    payment_intent_id: str
    """Stripe PaymentIntent ID for tracking."""

    deposit_address: str
    """USDC deposit address on Base network."""

    amount_usdc: str
    """Amount of USDC to send (e.g., "5.00")."""

    amount_usd_cents: int
    """Amount in USD cents."""

    credits: int
    """Number of credits that will be granted upon payment."""

    network: str
    """Blockchain network (always "base")."""

    currency: str
    """Payment currency (always "USDC")."""

    status: str
    """Payment status."""

    expires_at: str | None = None
    """When the deposit address expires (ISO 8601)."""


# ─── Machine Payment Creation ───────────────────────────────────────────────

machine_router = APIRouter(prefix="/v1/machine-payments", tags=["machine-payments"])


@machine_router.post("/create", response_model=MachinePaymentResponse)
async def create_machine_payment(req: MachinePaymentRequest):
    """Create a USDC deposit address for an AI agent to purchase credits.

    The agent sends USDC to the returned deposit address on Base.
    Once Stripe confirms the payment, credits are automatically granted.

    This endpoint does NOT require an API key — agents can purchase
    credits before having an account.  Include ``account_id`` to link
    the purchase to an existing account.

    Flow:
      1. Agent calls this endpoint with desired credit count
      2. Response includes a Base USDC deposit address
      3. Agent sends exact USDC amount to that address
      4. Stripe webhook fires → credits granted
      5. Agent can use credits via API key

    Args:
        req: Credit purchase request with amount and optional account ID.

    Returns:
        Deposit address, amount, and tracking info.
    """
    if not MACHINE_PAYMENTS_ENABLED:
        raise HTTPException(
            status_code=503,
            detail={
                "error": "MACHINE_PAYMENTS_DISABLED",
                "message": (
                    "Machine payments are not yet enabled. "
                    "Use the standard credit purchase flow at "
                    "https://www.slidearabi.com/dashboard/credits"
                ),
            },
        )

    if req.credits < MIN_CREDITS or req.credits > MAX_CREDITS:
        raise HTTPException(
            status_code=400,
            detail={
                "error": "INVALID_AMOUNT",
                "message": f"Credits must be between {MIN_CREDITS} and {MAX_CREDITS}.",
            },
        )

    amount_cents = req.credits * CENTS_PER_CREDIT

    # Build metadata
    metadata = {
        "source": "machine_payment",
        "credits": str(req.credits),
        "product": "slidearabi_credits",
    }
    if req.account_id:
        metadata["account_id"] = req.account_id
    if req.metadata:
        metadata.update(req.metadata)

    try:
        # Create a Stripe PaymentIntent with crypto payment method
        payment_intent = stripe.PaymentIntent.create(
            amount=amount_cents,
            currency="usd",
            payment_method_types=["crypto"],
            payment_method_data={"type": "crypto"},
            payment_method_options={
                "crypto": {"mode": "custom"},
            },
            metadata=metadata,
            confirm=True,
        )

        # Extract deposit address from the PaymentIntent
        next_action = payment_intent.get("next_action", {})
        deposit_details = next_action.get("crypto_collect_deposit_details", {})
        deposit_addresses = deposit_details.get("deposit_addresses", {})
        base_info = deposit_addresses.get("base", {})
        deposit_address = base_info.get("address", "")

        if not deposit_address:
            logger.error(
                "PaymentIntent %s did not return a deposit address: %s",
                payment_intent.id,
                next_action,
            )
            raise HTTPException(
                status_code=502,
                detail={
                    "error": "DEPOSIT_ADDRESS_UNAVAILABLE",
                    "message": "Could not generate a USDC deposit address. Try again.",
                },
            )

        # Calculate USDC amount (1:1 with USD for stablecoins)
        amount_usdc = f"{amount_cents / 100:.2f}"

        return MachinePaymentResponse(
            payment_intent_id=payment_intent.id,
            deposit_address=deposit_address,
            amount_usdc=amount_usdc,
            amount_usd_cents=amount_cents,
            credits=req.credits,
            network="base",
            currency="USDC",
            status=payment_intent.status,
            expires_at=None,  # Stripe manages expiry
        )

    except stripe.StripeError as e:
        logger.error("Stripe error creating machine payment: %s", e)
        raise HTTPException(
            status_code=502,
            detail={
                "error": "STRIPE_ERROR",
                "message": f"Payment creation failed: {str(e)}",
            },
        )


@machine_router.get("/status/{payment_intent_id}")
async def check_machine_payment_status(payment_intent_id: str):
    """Check the status of a machine payment.

    Agents can poll this to confirm their USDC payment was received
    and credits were granted.

    Args:
        payment_intent_id: The Stripe PaymentIntent ID from create_machine_payment.

    Returns:
        Current payment status and credit grant info.
    """
    if not MACHINE_PAYMENTS_ENABLED:
        raise HTTPException(status_code=503, detail="Machine payments disabled")

    try:
        pi = stripe.PaymentIntent.retrieve(payment_intent_id)

        result: dict[str, Any] = {
            "payment_intent_id": pi.id,
            "status": pi.status,
            "amount_usd_cents": pi.amount,
            "credits": int(pi.metadata.get("credits", 0)),
        }

        if pi.status == "succeeded":
            result["credits_granted"] = True
            result["message"] = "Payment confirmed. Credits have been granted."
        elif pi.status in ("requires_action", "processing"):
            result["credits_granted"] = False
            result["message"] = "Waiting for USDC payment confirmation."
        else:
            result["credits_granted"] = False
            result["message"] = f"Payment status: {pi.status}"

        return result

    except stripe.StripeError as e:
        raise HTTPException(status_code=404, detail=str(e))


# ─── Webhook Handler ────────────────────────────────────────────────────────
# This extends the existing Stripe webhook to handle crypto payments.


async def handle_crypto_payment_succeeded(event: dict[str, Any]) -> None:
    """Handle a successful USDC machine payment.

    Called from the main Stripe webhook handler when a crypto
    PaymentIntent succeeds.  Grants credits to the associated account.

    Args:
        event: The Stripe webhook event.
    """
    pi = event["data"]["object"]
    metadata = pi.get("metadata", {})

    if metadata.get("source") != "machine_payment":
        return  # Not a machine payment — ignore

    account_id = metadata.get("account_id")
    credits = int(metadata.get("credits", 0))
    pi_id = pi.get("id")

    if not credits:
        logger.warning("Machine payment %s has no credits in metadata", pi_id)
        return

    sb = _get_supabase()

    # Idempotency check
    existing = (
        sb.table("credit_transactions")
        .select("id")
        .eq("stripe_payment_intent", pi_id)
        .limit(1)
        .execute()
    )
    if existing.data:
        logger.info("Machine payment %s already processed — skipping", pi_id)
        return

    if account_id:
        # Grant credits to the specified account
        sb.rpc(
            "grant_credits",
            {
                "p_account_id": account_id,
                "p_amount": credits,
                "p_description": f"Machine payment: {credits} credits (USDC on Base)",
                "p_type": "purchase",
                "p_stripe_pi": pi_id,
            },
        ).execute()

        logger.info(
            "Machine payment: granted %d credits to %s (PI=%s)",
            credits,
            account_id,
            pi_id,
        )
    else:
        # No account linked — store as unclaimed credits
        # The agent can claim them later by providing the PI ID
        sb.table("unclaimed_credits").insert(
            {
                "stripe_payment_intent": pi_id,
                "credits": credits,
                "claimed": False,
            }
        ).execute()

        logger.info(
            "Machine payment: stored %d unclaimed credits (PI=%s)",
            credits,
            pi_id,
        )


# ─── Claim Unclaimed Credits ────────────────────────────────────────────────


class ClaimRequest(BaseModel):
    payment_intent_id: str
    account_id: str


@machine_router.post("/claim")
async def claim_credits(req: ClaimRequest):
    """Claim credits from a machine payment that wasn't linked to an account.

    If an agent purchased credits without specifying an account_id,
    they can claim them later using the PaymentIntent ID.

    Args:
        req: The PaymentIntent ID and target account ID.

    Returns:
        {\"success\": bool, \"credits\": int}
    """
    sb = _get_supabase()

    # Find unclaimed credits
    result = (
        sb.table("unclaimed_credits")
        .select("*")
        .eq("stripe_payment_intent", req.payment_intent_id)
        .eq("claimed", False)
        .limit(1)
        .execute()
    )

    if not result.data:
        raise HTTPException(
            status_code=404,
            detail="No unclaimed credits found for this payment.",
        )

    row = result.data[0]
    credits = row["credits"]

    # Grant to account
    sb.rpc(
        "grant_credits",
        {
            "p_account_id": req.account_id,
            "p_amount": credits,
            "p_description": (
                f"Claimed machine payment: {credits} credits "
                f"(PI={req.payment_intent_id})"
            ),
            "p_type": "purchase",
            "p_stripe_pi": req.payment_intent_id,
        },
    ).execute()

    # Mark as claimed
    sb.table("unclaimed_credits").update(
        {"claimed": True, "claimed_by": req.account_id}
    ).eq("id", row["id"]).execute()

    return {"success": True, "credits": credits}
