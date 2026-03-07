#!/usr/bin/env python3
"""
SlideArabi — Dual-LLM Translation Backend
═══════════════════════════════════════════════

Primary:   GPT-5.2 (flagship) for EN→AR translation
Secondary: Claude Sonnet 4.6 for terminology QA / consistency check

Architecture
────────────
1. Pre-processing: Protect abbreviations, numbers, brand names, URLs
2. Batching: Group strings into batches (≤40 per batch) for token efficiency
3. GPT-5.2 Translation: System prompt with domain glossary + context hints
4. Claude Sonnet 4.6 QA: Re-check terminology consistency, flag/fix issues
5. Post-processing: Restore protected tokens, validate output
6. Caching: JSON dict {english: arabic} — same format as existing caches

Sandbox Constraints
───────────────────
- Python `requests` HANGS — use `curl` via subprocess for ALL HTTP calls
- Each curl call must have --max-time timeout
- All errors caught and logged — translation failure never crashes pipeline

Dependencies
────────────
- curl (via subprocess)
- json, re, hashlib (stdlib)
- Environment variables: OPENAI_API_KEY, ANTHROPIC_API_KEY

Output Format
─────────────
Dict[str, str] — identical to the existing translations_cache JSON format
"""

from __future__ import annotations

import hashlib
import json
import logging
import os
import re
import subprocess
import tempfile
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

logger = logging.getLogger(__name__)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Configuration
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

@dataclass
class TranslatorConfig:
    """Configuration for the dual-LLM translation backend."""

    # API keys (from environment if not provided)
    openai_api_key: str = ""
    anthropic_api_key: str = ""

    # Model settings
    gpt_model: str = "gpt-5.2"              # GPT-5.2 flagship (primary translator)
    claude_model: str = "claude-sonnet-4-6"  # Claude Sonnet 4.6 (QA reviewer)

    # Batching
    batch_size: int = 40                 # Strings per API call
    max_retries: int = 2                 # Retries per API call on failure
    retry_delay: float = 2.0            # Seconds between retries

    # Parallelism
    max_translation_workers: int = 3
    max_qa_workers: int = 2

    # Timeouts (seconds)
    gpt_timeout: int = 120
    claude_timeout: int = 90

    # QA settings
    enable_qa_pass: bool = True          # Run Claude QA after GPT translation
    qa_batch_size: int = 60              # Larger batches for QA (less output)

    # Caching
    cache_dir: str = ""                  # Directory for translation caches
    use_cache: bool = True               # Use existing cache if available

    # Cost tracking
    track_costs: bool = True

    def __post_init__(self):
        if not self.openai_api_key:
            self.openai_api_key = os.environ.get("OPENAI_API_KEY", "")
        if not self.anthropic_api_key:
            self.anthropic_api_key = os.environ.get("ANTHROPIC_API_KEY", "")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Domain Glossary — The core of translation quality
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

# Abbreviations that must NEVER be translated — pass through as-is
PROTECTED_ABBREVIATIONS: Set[str] = {
    # Financial / business
    "EBITDA", "CAPEX", "OPEX", "IPO", "ROI", "ROE", "ROIC", "ROA",
    "P&L", "B2B", "B2C", "SaaS", "PaaS", "IaaS", "KPI", "KPIs",
    "YoY", "QoQ", "MoM", "CAGR", "IRR", "NPV", "DCF", "EPS",
    "PE", "P/E", "EV", "WACC", "LTV", "CAC", "ARR", "MRR",
    "GMV", "GTV", "AUM", "NAV", "ARPU", "NPS", "TAM", "SAM", "SOM",
    "FCF", "EBIT", "GAAP", "IFRS", "SOX",

    # Quarter/year identifiers
    "Q1", "Q2", "Q3", "Q4", "H1", "H2", "FY",
    "Q1-23", "Q2-23", "Q3-23", "Q4-23",
    "Q1-24", "Q2-24", "Q3-24", "Q4-24",
    "Q1-25", "Q2-25", "Q3-25", "Q4-25",
    "Q1-26", "Q2-26", "Q3-26", "Q4-26",
    "FY23", "FY24", "FY25", "FY26",
    "FY2023", "FY2024", "FY2025", "FY2026",

    # Technology
    "AI", "ML", "NLP", "LLM", "API", "SDK", "IoT", "AR", "VR",
    "GPU", "CPU", "TPU", "HW", "SW", "OS", "UI", "UX",
    "AWS", "GCP", "SSD", "RAM", "VRAM", "FPGA", "ASIC",
    "HTTP", "HTTPS", "REST", "SQL", "NoSQL", "ETL", "ELT",

    # Regulatory / compliance
    "GDPR", "HIPAA", "SOC", "SOC2", "PCI", "DSS", "ISO",
    "CCPA", "FERPA", "FDA", "SEC", "ESG",

    # Telecom
    "MOUs", "MOU", "ARPU", "PAYG", "SIM", "LTE", "5G", "4G",
    "MVNO", "MNO", "SMS", "VOIP", "VoIP",

    # Units
    "USD", "SAR", "AED", "GBP", "EUR", "BPS", "bps",
    "MW", "GW", "TWh", "MWh", "kWh",
    "TB", "GB", "MB", "KB",

    # Misc business
    "CEO", "CFO", "COO", "CTO", "CIO", "CISO", "CMO", "CPO",
    "HR", "R&D", "M&A", "JV", "MOU", "NDA", "LOI", "RFP", "RFI",
    "PMO", "OKR", "OKRs", "SLA", "SLAs",

    # Document markers
    "CONFIDENTIAL", "DRAFT", "INTERNAL", "TBD", "TBC", "N/A", "WIP",
}

# Abbreviation pattern: matches Q1-24, FY2025, H1-23, etc.
QUARTER_YEAR_PATTERN = re.compile(
    r'\b(Q[1-4]|H[12]|FY)\s*[-–]?\s*(\d{2,4})\b',
    re.IGNORECASE
)

# Number+unit patterns: 17.4M, $500K, 2.3B, 150MW, etc.
NUMBER_UNIT_PATTERN = re.compile(
    r'(\$?\d[\d,.]*\s*(?:M|B|K|T|MW|GW|TWh|MWh|kWh|bps|BPS|bn|mn|k|m|pp|p\.p\.))\b',
    re.IGNORECASE
)

# URL pattern
URL_PATTERN = re.compile(
    r'https?://\S+|www\.\S+|\S+\.(com|org|net|io|ai|gov|edu|co)\S*',
    re.IGNORECASE
)

# Email pattern
EMAIL_PATTERN = re.compile(r'\S+@\S+\.\S+')

# Glossary: specific terms with correct Arabic translations
# These override the LLM if it gets them wrong
DOMAIN_GLOSSARY: Dict[str, str] = {
    "CONFIDENTIAL": "سري",
    "Revenue pipeline": "مسار الإيرادات",
    "revenue pipeline": "مسار الإيرادات",
    "HW": "الأجهزة",
    "hardware": "الأجهزة",
    "SW": "البرمجيات",
    "software": "البرمجيات",
    "p.p.": "نقطة مئوية",
    "percentage points": "نقاط مئوية",
    "Net Margin": "صافي الهامش",
    "Total Revenue": "إجمالي الإيرادات",
    "PAYG": "الدفع عند الاستخدام",
    "Pay-as-you-go": "الدفع عند الاستخدام",
    "000s, USD": "بالآلاف، دولار أمريكي",
}


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Pre-processing: Protect tokens before sending to LLM
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

class TokenProtector:
    """
    Replaces abbreviations, URLs, emails, and number+unit patterns with
    placeholder tokens before translation, then restores them after.

    This prevents the LLM from mistranslating "HW" as "hazardous waste"
    or "Q1-24" as "Question 1-24".
    """

    def __init__(self):
        self._token_map: Dict[str, str] = {}
        self._counter = 0

    def _next_token(self) -> str:
        self._counter += 1
        return f"⟦PROT{self._counter:04d}⟧"

    def protect(self, text: str) -> str:
        """Replace protected patterns with placeholder tokens."""
        result = text

        # 1. Protect URLs first (longest matches)
        for match in URL_PATTERN.finditer(result):
            token = self._next_token()
            self._token_map[token] = match.group(0)
            result = result.replace(match.group(0), token, 1)

        # 2. Protect email addresses
        for match in EMAIL_PATTERN.finditer(result):
            token = self._next_token()
            self._token_map[token] = match.group(0)
            result = result.replace(match.group(0), token, 1)

        # 3. Protect quarter/year identifiers (Q1-24, FY2025, H1-23)
        for match in QUARTER_YEAR_PATTERN.finditer(result):
            full = match.group(0)
            if full not in self._token_map.values():
                token = self._next_token()
                self._token_map[token] = full
                result = result.replace(full, token, 1)

        # 4. Protect number+unit patterns (17.4M, $500K)
        for match in NUMBER_UNIT_PATTERN.finditer(result):
            full = match.group(0)
            if full not in self._token_map.values():
                token = self._next_token()
                self._token_map[token] = full
                result = result.replace(full, token, 1)

        # 5. Protect standalone abbreviations (whole-word match only)
        for abbr in PROTECTED_ABBREVIATIONS:
            # Use word boundary matching
            pattern = re.compile(r'\b' + re.escape(abbr) + r'\b')
            for match in pattern.finditer(result):
                if any(t in result[max(0,match.start()-6):match.end()+6]
                       for t in self._token_map):
                    continue  # Already inside a protected token
                token = self._next_token()
                self._token_map[token] = abbr
                result = result.replace(match.group(0), token, 1)

        return result

    def restore(self, text: str) -> str:
        """Replace placeholder tokens back with original values."""
        result = text
        for token, original in self._token_map.items():
            result = result.replace(token, original)
        return result

    @property
    def token_count(self) -> int:
        return len(self._token_map)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# System prompts
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

GPT_SYSTEM_PROMPT = """You are an expert English-to-Arabic translator specializing in business, finance, and technology content for GCC/Saudi executive audiences.

## Your Task
Translate each numbered English text segment into Modern Standard Arabic (MSA). Return ONLY a JSON object mapping each number to its Arabic translation.

## Critical Rules

### 1. Placeholder Tokens (⟦PROTXXXX⟧)
These are protected tokens. Copy them EXACTLY as-is into the Arabic output in the correct position. Do NOT translate, modify, or remove them. They represent abbreviations, numbers, URLs, and brand names that must stay in English.

### 2. Brand Names & Product Names
Keep ALL brand names and product names in their original English/Latin script:
- Company names: Google, Microsoft, Samsung, McKinsey, Goldman Sachs, etc.
- Product names: Spiral AI, DoorDash, Alegion, etc.
- Keep them in Latin script even when the rest of the sentence is Arabic

### 3. Financial & Business Terminology
Use established Arabic financial terminology used in Saudi/GCC business contexts:
- "Revenue" → "الإيرادات" (NOT "العائدات")
- "Revenue pipeline" → "مسار الإيرادات" (NOT "خط أنابيب الإيرادات")
- "CONFIDENTIAL" → "سري"
- "Net Margin" → "صافي الهامش"
- "EBITDA" → keep as "EBITDA"
- "Market Cap" → "القيمة السوقية"
- "Payable" → "مستحقة الدفع"

### 4. Sentence Direction
Arabic reads right-to-left. Ensure natural MSA sentence structure. Do NOT produce literal word-for-word translations.

### 5. Numbers and Units
- Keep digits in Western Arabic numerals (0-9), not Eastern Arabic (٠-٩)
- Keep currency symbols: $, €, £, SAR, AED
- "17.4 M" should stay as "17.4 مليون" (NOT "17.4 م")
- Percentages: keep the number, translate context

### 6. Formatting
- Preserve bullet point markers (•, -, ▪, etc.)
- Preserve line breaks if present in the source
- Do NOT add diacritics (tashkeel) unless critical for disambiguation

## Output Format
Return ONLY valid JSON. No markdown, no explanation, no preamble.
Example:
{"1": "النص العربي المترجم", "2": "نص آخر مترجم"}
"""

CLAUDE_QA_SYSTEM_PROMPT = """You are a bilingual Arabic-English QA specialist reviewing translations for Saudi/GCC executive presentations.

## Your Task
Review the English→Arabic translation pairs below. Check for:

1. **Abbreviation mangling**: Any English abbreviation (HW, SW, GDPR, Q1-24, CAPEX, etc.) that was incorrectly translated instead of preserved
2. **Brand name corruption**: Company or product names that were translated into Arabic instead of kept in Latin script
3. **Number/unit errors**: "M" translated as meters instead of millions, "K" as anything other than thousands, wrong currency symbols
4. **Semantic inversions**: Translations that mean the opposite or something wildly different from the source
5. **Terminology inconsistency**: Same English term translated different ways across entries
6. **Professional register**: Language that sounds informal or colloquial rather than executive-grade MSA

## Output Format
Return ONLY valid JSON with this structure:
{
  "issues_found": [
    {
      "index": "3",
      "english": "original text",
      "current_arabic": "current translation",
      "corrected_arabic": "fixed translation",
      "issue_type": "abbreviation_mangling|brand_corruption|number_error|semantic_error|inconsistency|register",
      "explanation": "brief explanation of what was wrong"
    }
  ],
  "summary": {
    "total_reviewed": 60,
    "issues_found": 3,
    "quality_score": 9.2
  }
}

If no issues found, return: {"issues_found": [], "summary": {"total_reviewed": N, "issues_found": 0, "quality_score": 10.0}}
"""


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# HTTP Clients (curl-based, sandbox-safe)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _curl_post(url: str, headers: Dict[str, str], body: dict,
               timeout: int = 120) -> Tuple[int, str]:
    """
    Make an HTTP POST via curl subprocess (Python requests HANGS in sandbox).

    Returns (status_code, response_body_text).
    """
    body_json = json.dumps(body, ensure_ascii=False)

    with tempfile.NamedTemporaryFile(
        mode='w', suffix='.json', delete=False, encoding='utf-8'
    ) as f:
        f.write(body_json)
        body_path = f.name

    try:
        cmd = [
            "curl", "-s", "-w", "\n%{http_code}",
            "-X", "POST",
            "--max-time", str(timeout),
            "-d", f"@{body_path}",
        ]
        for k, v in headers.items():
            cmd.extend(["-H", f"{k}: {v}"])
        cmd.append(url)

        result = subprocess.run(
            cmd, capture_output=True, text=True, timeout=timeout + 30
        )
        output = result.stdout.strip()

        # Last line is the HTTP status code
        lines = output.rsplit('\n', 1)
        if len(lines) == 2:
            body_text = lines[0]
            try:
                status = int(lines[1])
            except ValueError:
                status = 0
                body_text = output
        else:
            body_text = output
            status = 0

        return status, body_text

    except subprocess.TimeoutExpired:
        logger.error("curl timed out after %ds", timeout)
        return 0, '{"error": "timeout"}'
    except Exception as e:
        logger.error("curl failed: %s", e)
        return 0, f'{{"error": "{str(e)}"}}'
    finally:
        try:
            os.unlink(body_path)
        except OSError:
            pass


class GPTClient:
    """OpenAI GPT API client via curl."""

    def __init__(self, api_key: str, model: str = "gpt-5.2", timeout: int = 120):
        self.api_key = api_key
        self.model = model
        self.timeout = timeout
        self.total_input_tokens = 0
        self.total_output_tokens = 0
        self._token_lock = threading.Lock()

    def _api_call_with_retry(
        self,
        url: str,
        headers: Dict[str, str],
        body: Dict[str, Any],
        timeout: int,
        operation_name: str,
        max_retries: int = 3,
    ) -> str:
        """
        Execute curl API call with retries for transient failures.

        Retries on HTTP 429 and 5xx with exponential backoff.
        Raises RuntimeError on permanent failure.
        """
        status = 0
        response_text = ""
        for attempt in range(max_retries + 1):
            status, response_text = _curl_post(url, headers, body, timeout)

            if status == 200:
                return response_text

            should_retry = status == 429 or 500 <= status <= 599
            if should_retry and attempt < max_retries:
                backoff = 2 ** attempt
                logger.warning(
                    "%s retry %d/%d after HTTP %d; backing off %.1fs",
                    operation_name,
                    attempt + 1,
                    max_retries,
                    status,
                    backoff,
                )
                time.sleep(backoff)
                continue

            if should_retry:
                raise RuntimeError(
                    f"{operation_name} failed after {max_retries + 1} attempts; "
                    f"last HTTP {status}: {response_text[:500]}"
                )

            raise RuntimeError(
                f"{operation_name} failed with non-retriable HTTP {status}: {response_text[:500]}"
            )

        raise RuntimeError(
            f"{operation_name} failed after retries; last HTTP {status}: {response_text[:500]}"
        )

    def translate_batch(self, numbered_texts: Dict[str, str]) -> Dict[str, str]:
        """
        Send a batch of numbered English texts for translation.

        Args:
            numbered_texts: {"1": "English text 1", "2": "English text 2", ...}

        Returns:
            {"1": "Arabic translation 1", "2": "Arabic translation 2", ...}
        """
        # Build the user message with numbered texts
        lines = []
        for idx, text in sorted(numbered_texts.items(), key=lambda x: int(x[0])):
            lines.append(f"{idx}. {text}")
        user_content = "\n".join(lines)

        body = {
            "model": self.model,
            "messages": [
                {"role": "system", "content": GPT_SYSTEM_PROMPT},
                {"role": "user", "content": user_content}
            ],
            "temperature": 0.1,  # Low temperature for consistency
            "response_format": {"type": "json_object"},
            "max_tokens": 16000,
        }

        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {self.api_key}",
        }

        response_text = self._api_call_with_retry(
            url="https://api.openai.com/v1/chat/completions",
            headers=headers,
            body=body,
            timeout=self.timeout,
            operation_name="GPT translate_batch",
            max_retries=3,
        )

        try:
            resp = json.loads(response_text)
        except json.JSONDecodeError as e:
            logger.error("GPT response not valid JSON: %s", response_text[:500])
            raise RuntimeError(f"GPT response parse error: {e}")

        # Track token usage
        usage = resp.get("usage", {})
        with self._token_lock:
            self.total_input_tokens += usage.get("prompt_tokens", 0)
            self.total_output_tokens += usage.get("completion_tokens", 0)

        # Extract the translation JSON from the response
        content = resp["choices"][0]["message"]["content"]
        try:
            translations = json.loads(content)
        except json.JSONDecodeError:
            # Try to extract JSON from markdown code blocks
            match = re.search(r'```(?:json)?\s*(\{.*?\})\s*```', content, re.DOTALL)
            if match:
                translations = json.loads(match.group(1))
            else:
                logger.error("Could not parse GPT translation output: %s", content[:500])
                raise RuntimeError("GPT output not parseable as JSON")

        return translations


class ClaudeClient:
    """Anthropic Claude API client via curl."""

    def __init__(self, api_key: str, model: str = "claude-sonnet-4-6",
                 timeout: int = 90):
        self.api_key = api_key
        self.model = model
        self.timeout = timeout
        self.total_input_tokens = 0
        self.total_output_tokens = 0
        self._token_lock = threading.Lock()

    def _api_call_with_retry(
        self,
        url: str,
        headers: Dict[str, str],
        body: Dict[str, Any],
        timeout: int,
        operation_name: str,
        max_retries: int = 3,
    ) -> str:
        """
        Execute curl API call with retries for transient failures.

        Retries on HTTP 429 and 5xx with exponential backoff.
        Raises RuntimeError on permanent failure.
        """
        status = 0
        response_text = ""
        for attempt in range(max_retries + 1):
            status, response_text = _curl_post(url, headers, body, timeout)

            if status == 200:
                return response_text

            should_retry = status == 429 or 500 <= status <= 599
            if should_retry and attempt < max_retries:
                backoff = 2 ** attempt
                logger.warning(
                    "%s retry %d/%d after HTTP %d; backing off %.1fs",
                    operation_name,
                    attempt + 1,
                    max_retries,
                    status,
                    backoff,
                )
                time.sleep(backoff)
                continue

            if should_retry:
                raise RuntimeError(
                    f"{operation_name} failed after {max_retries + 1} attempts; "
                    f"last HTTP {status}: {response_text[:500]}"
                )

            raise RuntimeError(
                f"{operation_name} failed with non-retriable HTTP {status}: {response_text[:500]}"
            )

        raise RuntimeError(
            f"{operation_name} failed after retries; last HTTP {status}: {response_text[:500]}"
        )

    def qa_batch(self, pairs: List[Dict[str, str]]) -> Dict[str, Any]:
        """
        Send translation pairs for QA review.

        Args:
            pairs: [{"index": "1", "english": "...", "arabic": "..."}, ...]

        Returns:
            Parsed QA response with issues_found and summary.
        """
        # Build the user message
        lines = []
        for p in pairs:
            lines.append(f"{p['index']}. EN: {p['english']}")
            lines.append(f"   AR: {p['arabic']}")
            lines.append("")
        user_content = "\n".join(lines)

        body = {
            "model": self.model,
            "max_tokens": 8000,
            "messages": [
                {"role": "user", "content": user_content}
            ],
            "system": CLAUDE_QA_SYSTEM_PROMPT,
            "temperature": 0.0,
        }

        headers = {
            "Content-Type": "application/json",
            "x-api-key": self.api_key,
            "anthropic-version": "2023-06-01",
        }

        response_text = self._api_call_with_retry(
            url="https://api.anthropic.com/v1/messages",
            headers=headers,
            body=body,
            timeout=self.timeout,
            operation_name="Claude qa_batch",
            max_retries=3,
        )

        try:
            resp = json.loads(response_text)
        except json.JSONDecodeError as e:
            logger.error("Claude response not valid JSON: %s", response_text[:500])
            raise RuntimeError(f"Claude response parse error: {e}")

        # Track token usage
        usage = resp.get("usage", {})
        with self._token_lock:
            self.total_input_tokens += usage.get("input_tokens", 0)
            self.total_output_tokens += usage.get("output_tokens", 0)

        # Extract content
        content = resp["content"][0]["text"]
        try:
            qa_result = json.loads(content)
        except json.JSONDecodeError:
            match = re.search(r'```(?:json)?\s*(\{.*?\})\s*```', content, re.DOTALL)
            if match:
                qa_result = json.loads(match.group(1))
            else:
                logger.warning("Could not parse Claude QA output, treating as no issues")
                qa_result = {
                    "issues_found": [],
                    "summary": {"total_reviewed": len(pairs), "issues_found": 0,
                                "quality_score": 0.0, "parse_error": True}
                }

        return qa_result


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Translation Result Tracking
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

@dataclass
class TranslationReport:
    """Summary of a translation run."""
    total_strings: int = 0
    translated: int = 0
    from_cache: int = 0
    gpt_batches: int = 0
    qa_batches: int = 0
    qa_issues_found: int = 0
    qa_issues_fixed: int = 0
    gpt_input_tokens: int = 0
    gpt_output_tokens: int = 0
    claude_input_tokens: int = 0
    claude_output_tokens: int = 0
    elapsed_seconds: float = 0.0
    errors: List[str] = field(default_factory=list)

    def estimated_cost_usd(self) -> float:
        """Rough cost estimate based on published API pricing."""
        # GPT-5.2 pricing: $1.75/1M input, $14.00/1M output
        gpt_cost = (self.gpt_input_tokens * 1.75 / 1_000_000 +
                    self.gpt_output_tokens * 14.00 / 1_000_000)
        # Claude Sonnet 4.6 pricing: $3.00/1M input, $15.00/1M output
        claude_cost = (self.claude_input_tokens * 3.00 / 1_000_000 +
                       self.claude_output_tokens * 15.00 / 1_000_000)
        return gpt_cost + claude_cost

    def to_dict(self) -> Dict[str, Any]:
        return {
            "total_strings": self.total_strings,
            "translated": self.translated,
            "from_cache": self.from_cache,
            "gpt_batches": self.gpt_batches,
            "qa_batches": self.qa_batches,
            "qa_issues_found": self.qa_issues_found,
            "qa_issues_fixed": self.qa_issues_fixed,
            "tokens": {
                "gpt_input": self.gpt_input_tokens,
                "gpt_output": self.gpt_output_tokens,
                "claude_input": self.claude_input_tokens,
                "claude_output": self.claude_output_tokens,
            },
            "elapsed_seconds": round(self.elapsed_seconds, 1),
            "estimated_cost_usd": round(self.estimated_cost_usd(), 4),
            "errors": self.errors,
        }


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Main Translator Class
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

class DualLLMTranslator:
    """
    Belt-and-braces EN→AR translator.

    Pipeline:
    1. Deduplicate input strings
    2. Check cache for existing translations
    3. Pre-process: protect abbreviations/URLs/numbers with placeholder tokens
    4. Batch and translate via GPT-5.2 (flagship)
    5. Post-process: restore protected tokens
    6. QA pass via Claude Sonnet 4.6: check terminology consistency
    7. Apply QA fixes
    8. Apply domain glossary overrides (final safety net)
    9. Update cache and return
    """

    def __init__(self, config: Optional[TranslatorConfig] = None):
        self.config = config or TranslatorConfig()
        self.gpt = GPTClient(
            api_key=self.config.openai_api_key,
            model=self.config.gpt_model,
            timeout=self.config.gpt_timeout,
        )
        self.claude = ClaudeClient(
            api_key=self.config.anthropic_api_key,
            model=self.config.claude_model,
            timeout=self.config.claude_timeout,
        )
        self._cache: Dict[str, str] = {}
        self._report = TranslationReport()

    def load_cache(self, cache_path: str) -> int:
        """Load existing translation cache from JSON file."""
        p = Path(cache_path)
        if p.exists():
            try:
                with open(p, 'r', encoding='utf-8') as f:
                    self._cache = json.load(f)
                logger.info("Loaded %d cached translations from %s",
                           len(self._cache), cache_path)
                return len(self._cache)
            except Exception as e:
                logger.warning("Failed to load cache %s: %s", cache_path, e)
        return 0

    def save_cache(self, cache_path: str) -> None:
        """Save translation cache to JSON file."""
        p = Path(cache_path)
        p.parent.mkdir(parents=True, exist_ok=True)
        with open(p, 'w', encoding='utf-8') as f:
            json.dump(self._cache, f, ensure_ascii=False, indent=2)
        logger.info("Saved %d translations to %s", len(self._cache), cache_path)

    def translate(self, texts: List[str],
                  cache_path: Optional[str] = None) -> Dict[str, str]:
        """
        Translate a list of English strings to Arabic.

        Args:
            texts: List of English text strings to translate.
            cache_path: Optional path to load/save translation cache.

        Returns:
            Dict mapping each English string to its Arabic translation.
            Same format as the existing translations_cache JSON files.
        """
        start_time = time.monotonic()
        self._report = TranslationReport()

        # Deduplicate while preserving order
        unique_texts = list(dict.fromkeys(t for t in texts if t and t.strip()))
        self._report.total_strings = len(unique_texts)
        logger.info("Translation request: %d unique strings from %d total",
                    len(unique_texts), len(texts))

        # Load cache
        if cache_path and self.config.use_cache:
            self.load_cache(cache_path)

        # Split into cached vs. needing translation
        to_translate = []
        result: Dict[str, str] = {}

        for text in unique_texts:
            if text in self._cache:
                result[text] = self._cache[text]
                self._report.from_cache += 1
            else:
                to_translate.append(text)

        logger.info("Cache hits: %d, Need translation: %d",
                    self._report.from_cache, len(to_translate))

        if to_translate:
            # Step 1: Pre-process — protect tokens
            protector = TokenProtector()
            protected_texts = {}
            original_for_index: Dict[str, str] = {}

            for i, text in enumerate(to_translate, 1):
                idx = str(i)
                protected = protector.protect(text)
                protected_texts[idx] = protected
                original_for_index[idx] = text

            logger.info("Protected %d tokens across %d strings",
                       protector.token_count, len(to_translate))

            # Step 2: Batch and translate via GPT
            gpt_translations = self._gpt_translate_batched(protected_texts)

            # Step 3: Post-process — restore tokens
            restored: Dict[str, str] = {}
            for idx, arabic in gpt_translations.items():
                restored_text = protector.restore(arabic)
                restored[idx] = restored_text

            # Step 4: QA pass via Claude
            if self.config.enable_qa_pass and self.config.anthropic_api_key:
                qa_fixes = self._claude_qa_pass(original_for_index, restored)
                # Apply QA fixes
                for idx, fixed_arabic in qa_fixes.items():
                    restored[idx] = fixed_arabic
                    self._report.qa_issues_fixed += 1

            # Step 5: Apply domain glossary overrides (final safety net)
            for idx in restored:
                restored[idx] = self._apply_glossary_overrides(
                    original_for_index[idx], restored[idx]
                )

            # Step 6: Map back to original English keys
            for idx, arabic in restored.items():
                english = original_for_index[idx]
                result[english] = arabic
                self._cache[english] = arabic
                self._report.translated += 1

        # Save updated cache
        if cache_path:
            self.save_cache(cache_path)

        # Finalize report
        self._report.elapsed_seconds = time.monotonic() - start_time
        self._report.gpt_input_tokens = self.gpt.total_input_tokens
        self._report.gpt_output_tokens = self.gpt.total_output_tokens
        self._report.claude_input_tokens = self.claude.total_input_tokens
        self._report.claude_output_tokens = self.claude.total_output_tokens

        logger.info("Translation complete: %s", json.dumps(self._report.to_dict()))
        return result

    @property
    def report(self) -> TranslationReport:
        return self._report

    # ─────────────────────────────────────────────────────────────────────
    # GPT Translation (batched)
    # ─────────────────────────────────────────────────────────────────────

    def _gpt_translate_batched(
        self, protected_texts: Dict[str, str]
    ) -> Dict[str, str]:
        """
        Translate all protected texts via GPT in batches.
        """
        all_translations: Dict[str, str] = {}
        items = list(protected_texts.items())
        batch_size = self.config.batch_size
        total_batches = (len(items) + batch_size - 1) // batch_size

        batch_payloads: List[Tuple[int, Dict[str, str], Dict[str, str], int]] = []
        for batch_start in range(0, len(items), batch_size):
            batch = dict(items[batch_start:batch_start + batch_size])
            batch_num = batch_start // batch_size + 1

            renumbered: Dict[str, str] = {}
            idx_map: Dict[str, str] = {}
            for i, (orig_idx, text) in enumerate(batch.items(), 1):
                batch_idx = str(i)
                renumbered[batch_idx] = text
                idx_map[batch_idx] = orig_idx

            batch_payloads.append((batch_num, renumbered, idx_map, len(batch)))

        def process_batch(batch_num: int, renumbered: Dict[str, str], idx_map: Dict[str, str]):
            for attempt in range(self.config.max_retries + 1):
                try:
                    batch_result = self.gpt.translate_batch(renumbered)
                    return batch_num, batch_result, idx_map, None
                except Exception as e:
                    logger.warning("GPT batch %d attempt %d failed: %s",
                                  batch_num, attempt + 1, e)
                    if attempt < self.config.max_retries:
                        time.sleep(self.config.retry_delay * (attempt + 1))
                    else:
                        error_msg = (
                            f"GPT batch {batch_num} failed after "
                            f"{self.config.max_retries + 1} attempts: {e}"
                        )
                        return batch_num, {}, idx_map, error_msg
            return batch_num, {}, idx_map, f"GPT batch {batch_num} failed unexpectedly"

        batch_results_by_num: Dict[int, Tuple[Dict[str, str], Dict[str, str], Optional[str]]] = {}

        max_workers = max(1, self.config.max_translation_workers)
        if max_workers == 1:
            for batch_num, renumbered, idx_map, batch_len in batch_payloads:
                logger.info("GPT batch %d/%d (%d strings)", batch_num, total_batches, batch_len)
                result_tuple = process_batch(batch_num, renumbered, idx_map)
                _, batch_result, result_idx_map, error_msg = result_tuple
                batch_results_by_num[batch_num] = (batch_result, result_idx_map, error_msg)
        else:
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                future_to_batch_num = {}
                for batch_num, renumbered, idx_map, batch_len in batch_payloads:
                    logger.info("GPT batch %d/%d (%d strings)", batch_num, total_batches, batch_len)
                    future = executor.submit(process_batch, batch_num, renumbered, idx_map)
                    future_to_batch_num[future] = batch_num

                for future in as_completed(future_to_batch_num):
                    batch_num = future_to_batch_num[future]
                    try:
                        _, batch_result, result_idx_map, error_msg = future.result()
                    except Exception as e:
                        error_msg = f"GPT batch {batch_num} executor failure: {e}"
                        batch_result = {}
                        result_idx_map = {}
                    batch_results_by_num[batch_num] = (batch_result, result_idx_map, error_msg)

        for batch_num in sorted(batch_results_by_num.keys()):
            batch_result, idx_map, error_msg = batch_results_by_num[batch_num]
            if error_msg:
                self._report.errors.append(error_msg)
                logger.error(error_msg)
                for _, orig_idx in idx_map.items():
                    if orig_idx not in all_translations:
                        all_translations[orig_idx] = protected_texts[orig_idx]
                continue

            self._report.gpt_batches += 1
            for batch_idx, arabic in batch_result.items():
                orig_idx = idx_map.get(str(batch_idx), str(batch_idx))
                all_translations[orig_idx] = arabic

            for _, orig_idx in idx_map.items():
                if orig_idx not in all_translations:
                    all_translations[orig_idx] = protected_texts[orig_idx]

        return all_translations

    # ─────────────────────────────────────────────────────────────────────
    # Claude QA Pass
    # ─────────────────────────────────────────────────────────────────────

    def _claude_qa_pass(
        self,
        originals: Dict[str, str],
        translations: Dict[str, str],
    ) -> Dict[str, str]:
        """
        Run Claude QA on all translation pairs.

        Returns dict of {index: corrected_arabic} for any issues found.
        """
        fixes: Dict[str, str] = {}

        # Build pairs list
        pairs = []
        for idx in sorted(originals.keys(), key=lambda x: int(x)):
            if idx in translations:
                pairs.append({
                    "index": idx,
                    "english": originals[idx],
                    "arabic": translations[idx],
                })

        if not pairs:
            return fixes

        qa_batch_size = self.config.qa_batch_size
        total_batches = (len(pairs) + qa_batch_size - 1) // qa_batch_size

        qa_batches: List[Tuple[int, List[Dict[str, str]]]] = []
        for batch_start in range(0, len(pairs), qa_batch_size):
            batch = pairs[batch_start:batch_start + qa_batch_size]
            batch_num = batch_start // qa_batch_size + 1
            qa_batches.append((batch_num, batch))

        def process_qa_batch(batch_num: int, batch: List[Dict[str, str]]):
            try:
                qa_result = self.claude.qa_batch(batch)
                return batch_num, qa_result, None
            except Exception as e:
                return batch_num, None, f"Claude QA batch {batch_num} failed: {e}"

        qa_results_by_num: Dict[int, Tuple[Optional[Dict[str, Any]], Optional[str]]] = {}

        max_workers = max(1, self.config.max_qa_workers)
        if max_workers == 1:
            for batch_num, batch in qa_batches:
                logger.info("Claude QA batch %d/%d (%d pairs)",
                           batch_num, total_batches, len(batch))
                _, qa_result, error_msg = process_qa_batch(batch_num, batch)
                qa_results_by_num[batch_num] = (qa_result, error_msg)
        else:
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                future_to_batch_num = {}
                for batch_num, batch in qa_batches:
                    logger.info("Claude QA batch %d/%d (%d pairs)",
                               batch_num, total_batches, len(batch))
                    future = executor.submit(process_qa_batch, batch_num, batch)
                    future_to_batch_num[future] = batch_num

                for future in as_completed(future_to_batch_num):
                    batch_num = future_to_batch_num[future]
                    try:
                        _, qa_result, error_msg = future.result()
                    except Exception as e:
                        qa_result, error_msg = None, f"Claude QA batch {batch_num} executor failure: {e}"
                    qa_results_by_num[batch_num] = (qa_result, error_msg)

        for batch_num in sorted(qa_results_by_num.keys()):
            qa_result, error_msg = qa_results_by_num[batch_num]
            if error_msg:
                self._report.errors.append(error_msg)
                logger.warning(error_msg)
                continue

            if qa_result is None:
                continue

            self._report.qa_batches += 1

            issues = qa_result.get("issues_found", [])
            self._report.qa_issues_found += len(issues)

            for issue in issues:
                idx = str(issue.get("index", ""))
                corrected = issue.get("corrected_arabic", "")
                if idx and corrected and idx in translations:
                    fixes[idx] = corrected
                    logger.info(
                        "QA fix [%s] %s: %s → %s (%s)",
                        issue.get("issue_type", "unknown"),
                        idx,
                        translations[idx][:50],
                        corrected[:50],
                        issue.get("explanation", ""),
                    )

            summary = qa_result.get("summary", {})
            quality = summary.get("quality_score", "N/A")
            logger.info("QA batch %d quality score: %s", batch_num, quality)

        return fixes

    # ─────────────────────────────────────────────────────────────────────
    # Domain Glossary Overrides (final safety net)
    # ─────────────────────────────────────────────────────────────────────

    def _apply_glossary_overrides(self, english: str, arabic: str) -> str:
        """
        Apply deterministic glossary overrides as a final safety net.

        These catch any remaining issues that the LLMs missed.
        """
        result = arabic

        # Check if the entire English string has a glossary entry
        if english in DOMAIN_GLOSSARY:
            return DOMAIN_GLOSSARY[english]

        # Check for known bad Arabic translations and fix them
        BAD_TRANSLATIONS = {
            # "hazardous waste" for HW
            "مخلفات خطرة": "الأجهزة",
            "المخلفات الخطرة": "الأجهزة",
            "المخلفات الصلبة": "الأجهزة",
            # "father upgrades" for HW upgrades
            "ترقيات الأب": "ترقيات الأجهزة",
            # "trustworthy person" for CONFIDENTIAL
            "مؤتمن": "سري",
            # "GDP" for GDPR
            "الناتج المحلي الإجمالي": "اللائحة العامة لحماية البيانات",
            # "mouse fees" for MOUs
            "رسوم الفأرة": "دقائق الاستخدام",
            # "question" for quarter
            "السؤال 1-24": "الربع الأول 2024",
            "السؤال 2-24": "الربع الثاني 2024",
            "السؤال 3-24": "الربع الثالث 2024",
            "السؤال 4-24": "الربع الرابع 2024",
            # "pipeline" literal
            "خط أنابيب الإيرادات": "مسار الإيرادات",
            # "17.4 meters" for 17.4 M
            "17.4 م": "17.4 مليون",
        }

        for bad, good in BAD_TRANSLATIONS.items():
            if bad in result:
                result = result.replace(bad, good)

        return result


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Convenience function for pipeline integration
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def translate_texts(
    texts: List[str],
    cache_path: Optional[str] = None,
    openai_api_key: Optional[str] = None,
    anthropic_api_key: Optional[str] = None,
    enable_qa: bool = True,
) -> Tuple[Dict[str, str], TranslationReport]:
    """
    Convenience function: translate a list of English strings to Arabic.

    Returns (translations_dict, report).
    """
    config = TranslatorConfig(
        openai_api_key=openai_api_key or "",
        anthropic_api_key=anthropic_api_key or "",
        enable_qa_pass=enable_qa,
    )

    translator = DualLLMTranslator(config)
    translations = translator.translate(texts, cache_path=cache_path)
    return translations, translator.report


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# CLI for standalone testing
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

if __name__ == "__main__":
    import sys

    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

    if len(sys.argv) < 2:
        print("Usage: python llm_translator.py <texts_file.json> [cache_path.json]")
        print("  texts_file.json: JSON array of English strings, or")
        print("                   JSON object with English keys (existing cache format)")
        sys.exit(1)

    texts_file = sys.argv[1]
    cache_path = sys.argv[2] if len(sys.argv) > 2 else None

    # Load texts
    with open(texts_file, 'r', encoding='utf-8') as f:
        data = json.load(f)

    if isinstance(data, list):
        texts = data
    elif isinstance(data, dict):
        texts = list(data.keys())
    else:
        print("ERROR: Input must be a JSON array or object")
        sys.exit(1)

    print(f"Loaded {len(texts)} texts from {texts_file}")

    # Check for API keys
    if not os.environ.get("OPENAI_API_KEY"):
        print("ERROR: Set OPENAI_API_KEY environment variable")
        sys.exit(1)

    enable_qa = bool(os.environ.get("ANTHROPIC_API_KEY"))
    if not enable_qa:
        print("WARNING: ANTHROPIC_API_KEY not set — skipping Claude QA pass")

    translations, report = translate_texts(
        texts, cache_path=cache_path, enable_qa=enable_qa
    )

    print(f"\n{'='*60}")
    print("Translation Report")
    print(f"{'='*60}")
    print(json.dumps(report.to_dict(), indent=2))
    print(f"\nOutput: {len(translations)} translations")

    if cache_path:
        print(f"Cache saved to: {cache_path}")

    # Print sample
    print(f"\nSample translations:")
    for i, (en, ar) in enumerate(list(translations.items())[:5]):
        print(f"  {i+1}. EN: {en[:60]}...")
        print(f"     AR: {ar[:60]}...")
