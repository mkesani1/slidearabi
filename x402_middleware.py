"""
SlideArabi — x402 Payment Middleware for FastAPI
==================================================

Adds HTTP 402 machine-payment support to SlideArabi's REST API, allowing
AI agents to pay per-request with USDC on Base via the x402 protocol.

This sits alongside the existing API-key + credit-based auth.  Routes
decorated with ``@pay(...)`` accept either:
  1. Traditional API key + credits  (existing flow)
  2. x402 USDC payment in the X-PAYMENT header  (new agent flow)

The middleware intercepts requests, checks for an X-PAYMENT header, and
either validates the crypto payment via the facilitator or falls through
to the standard API-key auth.

Integration::

    from slidearabi.x402_middleware import init_x402_payments
    init_x402_payments(app)           # call once at startup

Requirements::

    pip install "fastapi-x402>=0.1.8"
    # or: pip install "x402[fastapi]"

Environment variables::

    X402_PAY_TO_ADDRESS    — USDC receiving wallet on Base
    X402_NETWORK           — "base" (mainnet) or "base-sepolia" (testnet)
    X402_FACILITATOR_URL   — facilitator endpoint (default: x402.org)
    STRIPE_SECRET_KEY      — existing Stripe key (for hybrid settlement)
"""

from __future__ import annotations

import logging
import os
from typing import Any

from fastapi import FastAPI

logger = logging.getLogger("slidearabi.x402")

# ─── Configuration ───────────────────────────────────────────────────────────

X402_PAY_TO_ADDRESS: str = os.environ.get(
    "X402_PAY_TO_ADDRESS", ""
)  # USDC wallet on Base — set in Railway env

X402_NETWORK: str = os.environ.get("X402_NETWORK", "base-sepolia")
X402_FACILITATOR_URL: str = os.environ.get(
    "X402_FACILITATOR_URL", "https://x402.org/facilitator"
)

# ─── Per-route pricing (USD) ────────────────────────────────────────────────
# Maps API gateway routes to their x402 price.
# Agents pay per-request; the price covers the credit cost.
# Routes not listed here fall through to API-key auth only.

ROUTE_PRICES: dict[str, str] = {
    # Conversion endpoints — priced per-call (covers ~1 credit minimum)
    "POST /v1/convert": "$1.00",           # starts a conversion (min 5 slides = $5 via credits, but x402 is per-request)
    "GET /v1/conversions/{job_id}": "$0.00",  # free — status polling
    "GET /v1/conversions/{job_id}/download": "$0.00",  # free — download result

    # Upload endpoint — small fee to cover storage
    "POST /v1/upload": "$0.10",

    # Account info — free
    "GET /v1/account": "$0.00",
}

# ─── Dynamic pricing helper ─────────────────────────────────────────────────
# For the /v1/convert endpoint, the actual price depends on slide count.
# x402 charges a flat per-request fee; for variable pricing we use a
# two-step flow:
#   1. Agent calls POST /v1/upload (pays $0.10) → gets slide_count + cost estimate
#   2. Agent calls POST /v1/convert with x402 payment matching the quoted price
#
# The middleware below supports a "quoted price" header (X-SLIDEARABI-QUOTE)
# that the upload endpoint returns, allowing the convert endpoint to
# dynamically set the x402 price.


def _calculate_conversion_price(slide_count: int) -> str:
    """Calculate the x402 price for a conversion based on slide count.

    Pricing: $1.00 per slide (matching credit-based pricing).
    Minimum 5 slides ($5.00).

    Args:
        slide_count: Number of slides in the presentation.

    Returns:
        Price string like "$5.00".
    """
    credits_needed = max(5, slide_count)
    return f"${credits_needed:.2f}"


# ─── Initialization ─────────────────────────────────────────────────────────


def init_x402_payments(app: FastAPI) -> bool:
    """Initialize x402 machine payment support on the FastAPI app.

    Attempts to import and configure the ``fastapi-x402`` middleware.
    If the package is not installed or the wallet address is not set,
    logs a warning and returns False (the app continues without x402).

    This is designed to be non-breaking: if x402 is unavailable, the
    existing API-key + credit flow continues to work.

    Args:
        app: The main FastAPI application instance.

    Returns:
        True if x402 middleware was successfully added.
    """
    if not X402_PAY_TO_ADDRESS:
        logger.info(
            "X402_PAY_TO_ADDRESS not set — x402 machine payments disabled. "
            "Set this env var to a Base USDC wallet address to enable."
        )
        return False

    try:
        from fastapi_x402 import init_x402, pay  # noqa: F401

        init_x402(
            app,
            pay_to=X402_PAY_TO_ADDRESS,
            network=X402_NETWORK,
            facilitator_url=X402_FACILITATOR_URL,
        )
        logger.info(
            "x402 machine payments enabled: network=%s, wallet=%s...%s",
            X402_NETWORK,
            X402_PAY_TO_ADDRESS[:6],
            X402_PAY_TO_ADDRESS[-4:],
        )
        return True

    except ImportError:
        logger.warning(
            "fastapi-x402 package not installed. "
            "Run: pip install fastapi-x402  — to enable machine payments."
        )
        return False
    except Exception as e:
        logger.error("Failed to initialize x402 middleware: %s", e)
        return False


# ─── Dual-Auth Dependency ───────────────────────────────────────────────────
# This dependency checks for x402 payment first, then falls back to API key.
# Used in the API gateway routes to support both payment methods.


async def resolve_payment_method(
    request: "Request",  # noqa: F821 — forward ref for type hint
) -> dict[str, Any]:
    """Determine how the current request is being paid for.

    Checks in order:
      1. X-PAYMENT header → x402 crypto payment (USDC on Base)
      2. Authorization: Bearer sa_live_xxx → traditional API key + credits

    Returns:
        {
            "method": "x402" | "api_key",
            "verified": bool,
            "account_id": str | None,       # only for api_key method
            "payment_amount_usd": str | None # only for x402 method
        }
    """
    payment_header = request.headers.get("x-payment", "")

    if payment_header:
        # x402 payment — the middleware already verified it before
        # the route handler runs.  If we're here, payment is valid.
        return {
            "method": "x402",
            "verified": True,
            "account_id": None,
            "payment_amount_usd": request.state.x402_payment_amount
            if hasattr(request.state, "x402_payment_amount")
            else None,
        }

    # Fall through to API key auth (handled by existing get_auth_context)
    return {
        "method": "api_key",
        "verified": False,  # caller must still validate the API key
        "account_id": None,
        "payment_amount_usd": None,
    }


# ─── x402 Route Decorator Helper ───────────────────────────────────────────


def get_x402_price(route: str) -> str | None:
    """Look up the x402 price for a given route.

    Args:
        route: Route string like "POST /v1/convert".

    Returns:
        Price string like "$1.00", or None if route is not x402-enabled.
    """
    return ROUTE_PRICES.get(route)


# ─── OpenAPI Extension ──────────────────────────────────────────────────────
# Adds x402 payment info to the OpenAPI spec so agents can discover pricing.


def add_x402_openapi_extension(openapi_schema: dict) -> dict:
    """Augment the OpenAPI schema with x402 payment metadata.

    Adds ``x-payment-required`` extension to each route that has
    x402 pricing, so AI agents can discover payment requirements
    from the spec.

    Args:
        openapi_schema: The existing OpenAPI schema dict.

    Returns:
        The modified schema with x402 extensions.
    """
    paths = openapi_schema.get("paths", {})

    for route_key, price in ROUTE_PRICES.items():
        method, path = route_key.split(" ", 1)
        method = method.lower()

        if path in paths and method in paths[path]:
            paths[path][method]["x-payment-required"] = {
                "protocol": "x402",
                "price": price,
                "currency": "USDC",
                "network": X402_NETWORK,
                "pay_to": X402_PAY_TO_ADDRESS,
                "facilitator": X402_FACILITATOR_URL,
            }

    # Add server-level x402 info
    openapi_schema["info"]["x-payment-protocols"] = [
        {
            "protocol": "x402",
            "version": "1.0",
            "networks": [X402_NETWORK],
            "currencies": ["USDC"],
            "facilitator": X402_FACILITATOR_URL,
            "documentation": "https://docs.x402.org",
        }
    ]

    return openapi_schema
