"""
SlideArabi — Payment Rails Test Suite
=======================================

Tests all three payment methods:
  1. Traditional API key + credits (existing)
  2. x402 USDC middleware (new)
  3. xpay MCP proxy config (new)

Run:  python -m pytest slidearabi/test_payment_rails.py -v
  or: python slidearabi/test_payment_rails.py   (standalone)
"""

from __future__ import annotations

import json
import os
import sys
import unittest
from unittest.mock import AsyncMock, MagicMock, patch

# Add parent to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


class TestX402Middleware(unittest.TestCase):
    """Test the x402 payment middleware configuration."""

    def test_import(self):
        """x402_middleware module imports without errors."""
        from slidearabi.x402_middleware import (
            ROUTE_PRICES,
            X402_NETWORK,
            add_x402_openapi_extension,
            get_x402_price,
            init_x402_payments,
        )
        self.assertIsNotNone(ROUTE_PRICES)
        self.assertIsNotNone(init_x402_payments)

    def test_route_prices_defined(self):
        """All key API routes have x402 pricing."""
        from slidearabi.x402_middleware import ROUTE_PRICES

        required_routes = [
            "POST /v1/convert",
            "POST /v1/upload",
            "GET /v1/conversions/{job_id}",
            "GET /v1/conversions/{job_id}/download",
            "GET /v1/account",
        ]
        for route in required_routes:
            self.assertIn(route, ROUTE_PRICES, f"Missing x402 price for {route}")

    def test_free_routes_are_zero(self):
        """Status polling and download should be free."""
        from slidearabi.x402_middleware import ROUTE_PRICES

        self.assertEqual(ROUTE_PRICES["GET /v1/conversions/{job_id}"], "$0.00")
        self.assertEqual(ROUTE_PRICES["GET /v1/conversions/{job_id}/download"], "$0.00")
        self.assertEqual(ROUTE_PRICES["GET /v1/account"], "$0.00")

    def test_paid_routes_have_price(self):
        """Upload and convert should cost money."""
        from slidearabi.x402_middleware import ROUTE_PRICES

        self.assertEqual(ROUTE_PRICES["POST /v1/convert"], "$1.00")
        self.assertEqual(ROUTE_PRICES["POST /v1/upload"], "$0.10")

    def test_calculate_conversion_price(self):
        """Conversion price scales with slide count."""
        from slidearabi.x402_middleware import _calculate_conversion_price

        self.assertEqual(_calculate_conversion_price(1), "$5.00")   # minimum 5
        self.assertEqual(_calculate_conversion_price(5), "$5.00")   # exact minimum
        self.assertEqual(_calculate_conversion_price(10), "$10.00")
        self.assertEqual(_calculate_conversion_price(45), "$45.00")

    def test_get_x402_price(self):
        """Price lookup returns correct values."""
        from slidearabi.x402_middleware import get_x402_price

        self.assertEqual(get_x402_price("POST /v1/convert"), "$1.00")
        self.assertIsNone(get_x402_price("DELETE /v1/nonexistent"))

    def test_openapi_extension(self):
        """OpenAPI schema gets x402 payment metadata."""
        from slidearabi.x402_middleware import add_x402_openapi_extension

        schema = {
            "info": {"title": "SlideArabi API"},
            "paths": {
                "/v1/upload": {
                    "post": {"summary": "Upload presentation"},
                },
                "/v1/account": {
                    "get": {"summary": "Account info"},
                },
            },
        }
        result = add_x402_openapi_extension(schema)

        # Check x402 extension added to upload route
        upload_ext = result["paths"]["/v1/upload"]["post"].get("x-payment-required")
        self.assertIsNotNone(upload_ext)
        self.assertEqual(upload_ext["protocol"], "x402")
        self.assertEqual(upload_ext["price"], "$0.10")
        self.assertEqual(upload_ext["currency"], "USDC")

        # Check server-level metadata
        self.assertIn("x-payment-protocols", result["info"])

    def test_init_without_wallet_returns_false(self):
        """init_x402_payments returns False when no wallet is configured."""
        from slidearabi.x402_middleware import init_x402_payments

        with patch.dict(os.environ, {"X402_PAY_TO_ADDRESS": ""}):
            # Re-import to pick up empty env
            import importlib
            import slidearabi.x402_middleware as mod
            importlib.reload(mod)
            app = MagicMock()
            result = mod.init_x402_payments(app)
            self.assertFalse(result)


class TestXpayMCPConfig(unittest.TestCase):
    """Test the xpay MCP proxy configuration."""

    def test_import(self):
        """xpay_mcp_config module imports without errors."""
        from slidearabi.xpay_mcp_config import (
            PROXY_URL,
            XPAY_TOOL_PRICING,
            get_claude_desktop_config,
            get_xpay_config,
        )
        self.assertIsNotNone(get_xpay_config)
        self.assertIsNotNone(PROXY_URL)

    def test_proxy_url_format(self):
        """Proxy URL follows xpay convention."""
        from slidearabi.xpay_mcp_config import PROXY_URL

        self.assertIn("mcp.xpay.sh", PROXY_URL)
        self.assertTrue(PROXY_URL.endswith("/mcp"))

    def test_tool_pricing_complete(self):
        """All MCP tools have pricing defined."""
        from slidearabi.xpay_mcp_config import XPAY_TOOL_PRICING

        expected_tools = [
            "upload_presentation",
            "convert_presentation",
            "get_conversion_status",
            "download_result",
            "get_account_info",
        ]
        for tool in expected_tools:
            self.assertIn(tool, XPAY_TOOL_PRICING, f"Missing price for {tool}")

    def test_free_tools_are_zero(self):
        """Status, download, and account tools should be free."""
        from slidearabi.xpay_mcp_config import XPAY_TOOL_PRICING

        self.assertEqual(XPAY_TOOL_PRICING["get_conversion_status"], 0.00)
        self.assertEqual(XPAY_TOOL_PRICING["download_result"], 0.00)
        self.assertEqual(XPAY_TOOL_PRICING["get_account_info"], 0.00)

    def test_convert_covers_minimum(self):
        """Convert tool pricing covers 5-slide minimum."""
        from slidearabi.xpay_mcp_config import XPAY_TOOL_PRICING

        self.assertGreaterEqual(XPAY_TOOL_PRICING["convert_presentation"], 5.00)

    def test_xpay_config_structure(self):
        """xpay config has all required sections."""
        from slidearabi.xpay_mcp_config import get_xpay_config

        config = get_xpay_config()
        self.assertIn("server", config)
        self.assertIn("wallet", config)
        self.assertIn("pricing", config)
        self.assertIn("metadata", config)

        # Server section
        self.assertEqual(config["server"]["name"], "SlideArabi")
        self.assertIn("url", config["server"])
        self.assertEqual(config["server"]["transport"], "streamable-http")

        # Pricing section
        self.assertEqual(config["pricing"]["model"], "per_tool")
        self.assertEqual(config["pricing"]["currency"], "USDC")
        self.assertIn("tools", config["pricing"])

        # Metadata
        self.assertIn("arabic", config["metadata"]["tags"])
        self.assertIn("rtl", config["metadata"]["tags"])

    def test_xpay_config_json_valid(self):
        """xpay config serializes to valid JSON."""
        from slidearabi.xpay_mcp_config import get_xpay_config_json

        json_str = get_xpay_config_json()
        parsed = json.loads(json_str)
        self.assertIsInstance(parsed, dict)

    def test_claude_desktop_config(self):
        """Claude Desktop config has correct structure."""
        from slidearabi.xpay_mcp_config import get_claude_desktop_config, PROXY_URL

        config = get_claude_desktop_config("test_key_123")
        self.assertIn("mcpServers", config)
        self.assertIn("slidearabi", config["mcpServers"])
        self.assertEqual(config["mcpServers"]["slidearabi"]["url"], PROXY_URL)
        self.assertIn("Authorization", config["mcpServers"]["slidearabi"]["headers"])

    def test_direct_mcp_config(self):
        """Direct MCP config uses the real server URL."""
        from slidearabi.xpay_mcp_config import get_direct_mcp_config, MCP_SERVER_URL

        config = get_direct_mcp_config("sa_live_test")
        self.assertEqual(
            config["mcpServers"]["slidearabi"]["url"],
            MCP_SERVER_URL,
        )

    def test_agent_discovery_metadata(self):
        """Agent discovery metadata is valid JSON-LD."""
        from slidearabi.xpay_mcp_config import get_agent_discovery_metadata

        meta = get_agent_discovery_metadata()
        self.assertEqual(meta["@type"], "SoftwareApplication")
        self.assertEqual(meta["name"], "SlideArabi")
        self.assertIn("offers", meta)
        self.assertIn("additionalProperty", meta)


class TestStripeMachinePayments(unittest.TestCase):
    """Test the Stripe machine payments module."""

    def test_import(self):
        """stripe_machine_payments module imports without errors."""
        from slidearabi.stripe_machine_payments import (
            MachinePaymentRequest,
            MachinePaymentResponse,
            machine_router,
        )
        self.assertIsNotNone(machine_router)

    def test_request_model_validation(self):
        """MachinePaymentRequest validates credit range."""
        from slidearabi.stripe_machine_payments import MachinePaymentRequest

        # Valid request
        req = MachinePaymentRequest(credits=10)
        self.assertEqual(req.credits, 10)
        self.assertIsNone(req.account_id)

        # With account
        req2 = MachinePaymentRequest(credits=50, account_id="acc_123")
        self.assertEqual(req2.account_id, "acc_123")

    def test_response_model(self):
        """MachinePaymentResponse has all required fields."""
        from slidearabi.stripe_machine_payments import MachinePaymentResponse

        resp = MachinePaymentResponse(
            payment_intent_id="pi_test",
            deposit_address="0xabc123",
            amount_usdc="5.00",
            amount_usd_cents=500,
            credits=5,
            network="base",
            currency="USDC",
            status="requires_action",
        )
        self.assertEqual(resp.network, "base")
        self.assertEqual(resp.currency, "USDC")
        self.assertEqual(resp.credits, 5)

    def test_credit_pricing(self):
        """Credit pricing matches $1.00/credit."""
        from slidearabi.stripe_machine_payments import CENTS_PER_CREDIT

        self.assertEqual(CENTS_PER_CREDIT, 100)

    def test_credit_limits(self):
        """Credit purchase limits are reasonable."""
        from slidearabi.stripe_machine_payments import MIN_CREDITS, MAX_CREDITS

        self.assertEqual(MIN_CREDITS, 1)
        self.assertGreaterEqual(MAX_CREDITS, 100)


class TestPaymentRailIntegration(unittest.TestCase):
    """Integration tests verifying all three rails work together."""

    def test_all_modules_import(self):
        """All payment modules import without conflicts."""
        from slidearabi import x402_middleware
        from slidearabi import xpay_mcp_config
        from slidearabi import stripe_machine_payments
        from slidearabi import stripe_credits

        self.assertTrue(True)  # If we get here, no import errors

    def test_pricing_consistency(self):
        """x402 and xpay pricing are consistent."""
        from slidearabi.x402_middleware import ROUTE_PRICES
        from slidearabi.xpay_mcp_config import XPAY_TOOL_PRICING

        # Upload pricing: x402=$0.10, xpay=$0.10
        self.assertEqual(ROUTE_PRICES["POST /v1/upload"], "$0.10")
        self.assertEqual(XPAY_TOOL_PRICING["upload_presentation"], 0.10)

        # Status/download should be free in both
        self.assertEqual(ROUTE_PRICES["GET /v1/conversions/{job_id}"], "$0.00")
        self.assertEqual(XPAY_TOOL_PRICING["get_conversion_status"], 0.00)

    def test_credit_packages_exist(self):
        """Traditional credit packages are still defined."""
        from slidearabi.stripe_credits import CREDIT_PACKAGES

        self.assertIn("starter_25", CREDIT_PACKAGES)
        self.assertIn("growth_100", CREDIT_PACKAGES)
        self.assertIn("studio_500", CREDIT_PACKAGES)
        self.assertIn("enterprise_2000", CREDIT_PACKAGES)

    def test_promo_codes_exist(self):
        """Promo codes are still defined."""
        from slidearabi.stripe_credits import PROMO_CODES

        self.assertIn("SLIDETEST2026", PROMO_CODES)
        self.assertIn("FOUNDER", PROMO_CODES)
        self.assertIn("DEMO", PROMO_CODES)


# ─── Run ─────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    # Create __init__.py if missing
    init_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "__init__.py")
    if not os.path.exists(init_path):
        with open(init_path, "w") as f:
            f.write('"""SlideArabi payment modules."""\n')

    unittest.main(verbosity=2)
