"""
SlideArabi — xpay MCP Monetization Configuration
==================================================

Configures the xpay.sh payment proxy for the SlideArabi MCP server,
enabling AI agents to pay per-tool-call in USDC without any code
changes to the existing MCP server.

How it works:
  1. Your MCP server runs at its normal URL (e.g., https://api.slidearabi.com/mcp)
  2. xpay provides a proxy URL (https://slidearabi.mcp.xpay.sh/mcp)
  3. Agents connect to the proxy URL instead
  4. xpay intercepts each tool call, collects USDC payment, then forwards to your server
  5. Revenue flows to your USDC wallet on Base

Setup steps:
  1. Create a USDC wallet on Base (or use your existing one)
  2. Go to https://xpay.sh and register your MCP server
  3. Configure per-tool pricing (see XPAY_TOOL_PRICING below)
  4. Share the proxy URL with agents

This module provides:
  - The pricing configuration to copy into xpay dashboard
  - A health-check endpoint for xpay to verify your server
  - Helper to generate the Claude Desktop config snippet

Usage::

    from slidearabi.xpay_mcp_config import (
        get_xpay_config,
        get_claude_desktop_config,
        PROXY_URL,
    )
"""

from __future__ import annotations

import json
import os
from typing import Any

# ─── Configuration ───────────────────────────────────────────────────────────

# Your MCP server's real URL (where xpay forwards requests to)
MCP_SERVER_URL: str = os.environ.get(
    "MCP_SERVER_URL",
    "https://api.slidearabi.com/mcp",
)

# xpay proxy URL (agents connect here instead of your real server)
XPAY_SLUG: str = os.environ.get("XPAY_SLUG", "slidearabi")
PROXY_URL: str = f"https://{XPAY_SLUG}.mcp.xpay.sh/mcp"

# USDC receiving wallet on Base
XPAY_WALLET_ADDRESS: str = os.environ.get(
    "XPAY_WALLET_ADDRESS",
    os.environ.get("X402_PAY_TO_ADDRESS", ""),
)

# ─── Per-Tool Pricing ───────────────────────────────────────────────────────
# Maps each MCP tool to its USDC price per invocation.
#
# Pricing philosophy:
#   - upload_presentation:     $0.10 (storage + parsing cost)
#   - convert_presentation:    $1.00 base + $1.00/slide (dynamic, see note)
#   - get_conversion_status:   $0.00 (free — encourage polling)
#   - download_result:         $0.00 (free — they already paid for conversion)
#   - get_account_info:        $0.00 (free — account management)
#
# NOTE: xpay charges a flat fee per tool call.  For variable pricing
# (per-slide), we use the default $1.00 for convert_presentation and
# handle the slide-based pricing in our credit system.  Agents using
# xpay effectively pay $1.00 per conversion *attempt* via USDC, plus
# credits cover the per-slide cost.  Alternatively, set a higher flat
# rate (e.g., $5.00) to cover typical 5-slide minimum.

XPAY_TOOL_PRICING: dict[str, float] = {
    "upload_presentation": 0.10,       # $0.10 per upload
    "convert_presentation": 5.00,      # $5.00 per conversion (covers min 5 slides)
    "get_conversion_status": 0.00,     # free — status polling
    "download_result": 0.00,           # free — already paid
    "get_account_info": 0.00,          # free — account info
}

# Default price for any unlisted tools
XPAY_DEFAULT_PRICE: float = 0.01


# ─── Configuration Generator ────────────────────────────────────────────────


def get_xpay_config() -> dict[str, Any]:
    """Generate the xpay.sh registration configuration.

    Copy this JSON into the xpay dashboard when registering
    your MCP server, or use it with the xpay API.

    Returns:
        Configuration dict ready for xpay registration.
    """
    return {
        "server": {
            "name": "SlideArabi",
            "description": (
                "AI-powered English to Arabic PowerPoint conversion. "
                "Full RTL layout transformation, chart mirroring, and "
                "professional Arabic typography."
            ),
            "url": MCP_SERVER_URL,
            "transport": "streamable-http",
            "version": "1.0.0",
        },
        "wallet": {
            "address": XPAY_WALLET_ADDRESS,
            "network": "base",
            "currency": "USDC",
        },
        "pricing": {
            "model": "per_tool",
            "currency": "USDC",
            "tools": XPAY_TOOL_PRICING,
            "default": XPAY_DEFAULT_PRICE,
        },
        "metadata": {
            "category": "document-processing",
            "tags": [
                "arabic",
                "rtl",
                "powerpoint",
                "pptx",
                "translation",
                "localization",
                "mena",
            ],
            "website": "https://www.slidearabi.com",
            "documentation": "https://www.slidearabi.com/developers",
            "support_email": "support@slidearabi.com",
        },
    }


def get_xpay_config_json(pretty: bool = True) -> str:
    """Return the xpay configuration as a JSON string.

    Args:
        pretty: If True, format with indentation for readability.

    Returns:
        JSON string of the xpay configuration.
    """
    config = get_xpay_config()
    return json.dumps(config, indent=2 if pretty else None)


# ─── Client Configuration Generators ────────────────────────────────────────


def get_claude_desktop_config(api_key: str = "YOUR_API_KEY") -> dict[str, Any]:
    """Generate the Claude Desktop MCP config snippet.

    Users add this to their ``claude_desktop_config.json`` to connect
    Claude Desktop to SlideArabi via the xpay payment proxy.

    Args:
        api_key: The user's xpay API key (or placeholder).

    Returns:
        Config dict for claude_desktop_config.json mcpServers section.
    """
    return {
        "mcpServers": {
            "slidearabi": {
                "url": PROXY_URL,
                "headers": {
                    "Authorization": f"Bearer {api_key}",
                },
            }
        }
    }


def get_direct_mcp_config(api_key: str = "sa_live_xxx") -> dict[str, Any]:
    """Generate the direct MCP config (no xpay, uses API key + credits).

    For users who prefer the traditional credit-based model.

    Args:
        api_key: The user's SlideArabi API key (or placeholder).

    Returns:
        Config dict for claude_desktop_config.json mcpServers section.
    """
    return {
        "mcpServers": {
            "slidearabi": {
                "type": "streamable-http",
                "url": MCP_SERVER_URL,
                "headers": {
                    "Authorization": f"Bearer {api_key}",
                },
            }
        }
    }


# ─── Agent Discovery Metadata ──────────────────────────────────────────────
# JSON-LD structured data for agent discovery platforms.


def get_agent_discovery_metadata() -> dict[str, Any]:
    """Generate JSON-LD metadata for AI agent discovery.

    This metadata helps agent registries (MCP Hub, Composio,
    Smithery, etc.) index and recommend SlideArabi.

    Returns:
        JSON-LD structured data dict.
    """
    return {
        "@context": "https://schema.org",
        "@type": "SoftwareApplication",
        "name": "SlideArabi",
        "description": (
            "AI-powered English to Arabic PowerPoint conversion with full "
            "RTL layout transformation, chart mirroring, typography fixing, "
            "and dialect support. Available via REST API and MCP."
        ),
        "applicationCategory": "DocumentProcessing",
        "operatingSystem": "Cloud",
        "url": "https://www.slidearabi.com",
        "offers": {
            "@type": "Offer",
            "price": "1.00",
            "priceCurrency": "USD",
            "description": "Per slide credit",
        },
        "potentialAction": {
            "@type": "UseAction",
            "target": {
                "@type": "EntryPoint",
                "urlTemplate": "https://api.slidearabi.com/v1/convert",
                "httpMethod": "POST",
                "contentType": "multipart/form-data",
            },
        },
        "additionalProperty": [
            {
                "@type": "PropertyValue",
                "name": "mcp_server_url",
                "value": MCP_SERVER_URL,
            },
            {
                "@type": "PropertyValue",
                "name": "xpay_proxy_url",
                "value": PROXY_URL,
            },
            {
                "@type": "PropertyValue",
                "name": "payment_protocols",
                "value": "x402, api_key+credits, xpay_mcp",
            },
            {
                "@type": "PropertyValue",
                "name": "supported_formats",
                "value": "pptx",
            },
            {
                "@type": "PropertyValue",
                "name": "supported_languages",
                "value": "en→ar (MSA, Gulf, Levantine, Egyptian)",
            },
        ],
    }


# ─── Print configs (for manual setup) ──────────────────────────────────────


def print_setup_instructions() -> None:
    """Print step-by-step xpay setup instructions to the console."""
    print("=" * 70)
    print("  SlideArabi — xpay MCP Monetization Setup")
    print("=" * 70)
    print()
    print("Step 1: Go to https://xpay.sh and create an account")
    print()
    print("Step 2: Register your MCP server with this configuration:")
    print("-" * 50)
    print(get_xpay_config_json())
    print("-" * 50)
    print()
    print(f"Step 3: Your proxy URL will be: {PROXY_URL}")
    print()
    print("Step 4: Share the Claude Desktop config with users:")
    print("-" * 50)
    print(json.dumps(get_claude_desktop_config(), indent=2))
    print("-" * 50)
    print()
    print("Step 5: Update your Developers page with both connection options:")
    print(f"  - Direct (credits):  {MCP_SERVER_URL}")
    print(f"  - xpay (USDC):       {PROXY_URL}")
    print()
    print("=" * 70)


if __name__ == "__main__":
    print_setup_instructions()
