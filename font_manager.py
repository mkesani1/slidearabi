"""
font_manager.py — Font detection and auto-install system for PPTX files.

Detects all fonts used in a PPTX, checks whether they're installed, and
automatically downloads & installs missing ones from Google Fonts.

Usage:
    python -m slidearabi.font_manager /path/to/deck.pptx
"""

from __future__ import annotations

import json
import logging
import os
import re
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

from lxml import etree

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Namespace map for DrawingML XML
# ---------------------------------------------------------------------------
DRAWINGML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NSMAP = {
    "a": DRAWINGML_NS,
    "c": CHART_NS,
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

# ---------------------------------------------------------------------------
# Fonts that are system / Microsoft proprietary — skip download attempts
# ---------------------------------------------------------------------------
SYSTEM_FONTS: Set[str] = {
    # Microsoft core fonts
    "arial", "arial black", "arial narrow", "arial unicode ms",
    "calibri", "calibri light", "cambria", "cambria math",
    "comic sans ms", "consolas", "constantia", "corbel",
    "courier new", "franklin gothic medium",
    "georgia", "impact",
    "lucida console", "lucida sans unicode",
    "microsoft himalaya", "microsoft jhenghei", "microsoft new tai lue",
    "microsoft phagspa", "microsoft sans serif", "microsoft tai le",
    "microsoft uighur", "microsoft yi baiti",
    "mingliu", "mingliu-extb", "mingliu_hkscs", "mingliu_hkscs-extb",
    "mongolian baiti", "moolboran",
    "ms gothic", "ms pgothic", "ms ui gothic",
    "mv boli",
    "nsimsun", "palatino linotype",
    "segoe print", "segoe script", "segoe ui", "segoe ui emoji",
    "segoe ui historic", "segoe ui light", "segoe ui semibold",
    "segoe ui symbol",
    "simsun", "simsun-extb",
    "sylfaen",
    "symbol",
    "tahoma",
    "times new roman",
    "trebuchet ms",
    "tunga",
    "verdana",
    "webdings", "wingdings", "wingdings 2", "wingdings 3",
    # macOS built-in
    "helvetica", "helvetica neue",
    "san francisco", "sf pro", "sf mono", "sf compact",
    "new york",
    "gill sans", "gill sans mt",  # Apple/Monotype proprietary
    # Common Linux system fonts
    "liberation mono", "liberation sans", "liberation serif",
    "dejavu sans", "dejavu serif", "dejavu sans mono",
    "freemono", "freesans", "freeserif",
    # Common Indic/CJK glyphs (usually pre-installed with system)
    "angsana new", "cordia new", "daunh penh", "dok champa",
    "dokchampa", "daunpenh",  # Windows Southeast Asian fonts
    "estrangelo edessa", "euphemia", "gautami", "iskoola pota",
    "kalinga", "kartika", "latha", "mangal",
    "nyala", "plantagenet cherokee", "raavi", "shruti", "vrinda",
    # Generic / placeholder names
    "+mj-lt", "+mn-lt", "+mj-ea", "+mn-ea", "+mj-cs", "+mn-cs",
}

# Placeholder typeface values in DrawingML — skip these
THEME_PLACEHOLDER_TYPEFACES: Set[str] = {
    "+mj-lt", "+mn-lt", "+mj-ea", "+mn-ea", "+mj-cs", "+mn-cs",
    "+mj-bidi", "+mn-bidi",
}

# Common font weight/style words that appear as suffix in font family names.
# E.g. "Roboto Black" → base family "Roboto"
WEIGHT_STYLE_SUFFIXES: List[str] = [
    "hairline", "thin", "ultralight", "extra light", "extralight",
    "light", "regular", "normal", "book",
    "medium", "demi", "semibold", "semi bold", "demibold",
    "bold", "extrabold", "extra bold", "ultrabold", "ultra bold",
    "black", "heavy", "poster",
    "condensed", "cond", "expanded", "extended",
    "italic", "oblique", "slanted",
    "narrow", "wide",
]

# Google Fonts base URL (raw GitHub)
GF_RAW_BASE = "https://raw.githubusercontent.com/google/fonts/main"
GF_API_BASE = "https://api.github.com/repos/google/fonts/contents"

# User font directory
USER_FONT_DIR = Path.home() / ".fonts"

# Known font family renames / aliases:
# Some fonts were renamed between versions or between PPTX authoring apps
# and the Google Fonts repository.  Map the old/app name → the current installed name.
FONT_ALIASES: Dict[str, str] = {
    "source sans pro": "source sans 3",
    "source sans pro semibold": "source sans 3",
    "source serif pro": "source serif 4",
    "source code pro": "source code pro",  # unchanged
    "bebas neue pro": "bebas neue",
    "gill sans": "gill sans",  # will be skipped (system font)
}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _normalize_font_name(name: str) -> str:
    """Return a cleaned-up font family name (strip extra whitespace)."""
    return " ".join(name.strip().split())


def _font_dir_key(family_name: str) -> str:
    """Convert a font family name to the Google Fonts directory key.

    Google Fonts uses all-lowercase, no spaces or hyphens for directory names.
    E.g. "Open Sans" → "opensans", "Bebas Neue" → "bebasneue"
    """
    return re.sub(r"[^a-z0-9]", "", family_name.lower())


def _is_system_font(name: str) -> bool:
    """Return True if the font is a known system/proprietary font."""
    return name.lower() in SYSTEM_FONTS


def _strip_weight_suffix(family_name: str) -> str:
    """Strip trailing weight/style words from a font family name.

    E.g. "Roboto Black" → "Roboto", "Open Sans Light" → "Open Sans"
    """
    parts = family_name.split()
    while parts:
        lower_last = parts[-1].lower()
        if lower_last in WEIGHT_STYLE_SUFFIXES:
            parts = parts[:-1]
        else:
            break
    return " ".join(parts) if parts else family_name


def _is_cjk_font(name: str) -> bool:
    """Return True if the font name contains CJK / non-ASCII characters."""
    return any(ord(c) > 127 for c in name)


# ---------------------------------------------------------------------------
# FontManager class
# ---------------------------------------------------------------------------

class FontManager:
    """Detect, check, and install fonts used in PPTX files."""

    def __init__(self, font_dir: Optional[Path] = None, timeout: int = 45):
        """
        Args:
            font_dir: Directory where user fonts are installed.
                      Defaults to ~/.fonts/
            timeout:  Timeout in seconds for curl download commands.
        """
        self.font_dir = font_dir or USER_FONT_DIR
        self.timeout = timeout
        self.font_dir.mkdir(parents=True, exist_ok=True)
        # In-memory cache: family_name → category/dir_name
        self._gf_dir_cache: Dict[str, Optional[str]] = {}

    # ------------------------------------------------------------------
    # 1. Font Detection
    # ------------------------------------------------------------------

    def detect_fonts(self, pptx_path: str) -> List[str]:
        """Return all font family names used in the PPTX (deduplicated, sorted).

        Scans:
        - All slide, layout, master XML
        - Theme XML (majorFont / minorFont)
        - Chart XML
        - Embedded font list (ppt/fonts/)
        - a:latin, a:ea, a:cs, a:rPr typeface attributes
        """
        fonts: Set[str] = set()

        with zipfile.ZipFile(pptx_path, "r") as zf:
            all_names = zf.namelist()

            # Collect XML files to scan
            xml_files = [n for n in all_names if n.endswith(".xml")]

            for xml_name in xml_files:
                try:
                    raw = zf.read(xml_name)
                    tree = etree.fromstring(raw)
                    self._collect_fonts_from_tree(tree, fonts, xml_name)
                except Exception as exc:
                    logger.debug("Could not parse %s: %s", xml_name, exc)

            # Embedded fonts
            for name in all_names:
                if name.startswith("ppt/fonts/") and not name.endswith("/"):
                    # Embedded font file — derive family name from filename
                    fname = Path(name).stem  # e.g. "font1"
                    # Not useful for family name; log debug only
                    logger.debug("Embedded font file: %s", name)

        # Filter: remove placeholders, CJK names, and empty strings
        cleaned: Set[str] = set()
        for f in fonts:
            f = _normalize_font_name(f)
            if (
                f
                and f not in THEME_PLACEHOLDER_TYPEFACES
                and not _is_cjk_font(f)
            ):
                cleaned.add(f)

        return sorted(cleaned)

    def _collect_fonts_from_tree(
        self,
        tree: etree._Element,
        fonts: Set[str],
        source_name: str = "",
    ) -> None:
        """Walk an XML element tree and collect all typeface values."""

        # Helper to add a non-empty typeface
        def add(tf: Optional[str]) -> None:
            if tf and tf.strip():
                fonts.add(tf.strip())

        is_theme = "theme" in source_name.lower()

        for elem in tree.iter():
            # Resolve local tag name (strip namespace)
            tag = etree.QName(elem.tag).localname if "{" in elem.tag else elem.tag

            # Direct typeface attributes on a:latin, a:ea, a:cs, a:sym
            if tag in ("latin", "ea", "cs", "sym"):
                add(elem.get("typeface"))

            # Run properties — may have a direct font reference
            elif tag == "rPr":
                add(elem.get("typeface"))

            # Theme font definitions
            elif tag in ("majorFont", "minorFont"):
                # Children: a:latin, a:ea, a:cs, plus supplemental font elems
                for child in elem:
                    child_tag = etree.QName(child.tag).localname if "{" in child.tag else child.tag
                    if child_tag in ("latin", "ea", "cs"):
                        add(child.get("typeface"))
                    elif child_tag == "font":
                        # Supplemental script fonts: <a:font script="..." typeface="..."/>
                        add(child.get("typeface"))

            # Paragraph-level font size / formatting (covers chart text)
            elif tag == "defRPr":
                add(elem.get("typeface"))

            # Text body default paragraph / run properties
            elif tag == "bodyPr":
                pass  # no typeface here, but skip

            # Table style / theme elements
            elif tag == "fontRef":
                # <a:fontRef idx="minor"> — references theme font, not a name
                pass

            # Also catch any element that has a "typeface" attribute
            else:
                tf = elem.get("typeface")
                if tf:
                    add(tf)

    # ------------------------------------------------------------------
    # 2. Font Status Check
    # ------------------------------------------------------------------

    def check_missing(self, font_names: List[str]) -> List[str]:
        """Return font names that are NOT installed on the system.

        Uses fc-match and verifies the returned family name matches the
        requested one (to avoid accepting fallback fonts).
        """
        missing = []
        for name in font_names:
            if not self._is_font_installed(name):
                missing.append(name)
        return missing

    def _is_font_installed(self, family_name: str) -> bool:
        """Return True if the exact font family (or its base weight variant) is installed.

        fc-match always returns *something* (a fallback), so we parse the
        output and check that the returned family matches what we asked for.

        Also tries:
        - Stripping common weight/style suffix words (e.g. "Roboto Black" → "Roboto")
        - Known font family aliases (e.g. "Source Sans Pro" → "Source Sans 3")
        """
        # Skip system/proprietary fonts — assume they're present or a
        # close-enough substitute exists
        if _is_system_font(family_name):
            logger.debug("Skipping system font check: %s", family_name)
            return True

        # Names to try: full name, alias, and weight-stripped base name
        names_to_try = [family_name]

        # Check alias map
        alias = FONT_ALIASES.get(family_name.lower())
        if alias and alias != family_name.lower():
            names_to_try.append(alias)

        # Weight-suffix stripped name
        base_name = _strip_weight_suffix(family_name)
        if base_name != family_name:
            names_to_try.append(base_name)
            # Also check alias for base name
            base_alias = FONT_ALIASES.get(base_name.lower())
            if base_alias:
                names_to_try.append(base_alias)

        for name in names_to_try:
            if self._fc_match_installed(name):
                return True

        logger.debug("Font '%s' not installed (tried: %s)", family_name, names_to_try)
        return False

    def _fc_match_installed(self, family_name: str) -> bool:
        """Run fc-match for a single family name and check if it matches."""
        try:
            result = subprocess.run(
                ["fc-match", "--format=%{family}", family_name],
                capture_output=True,
                text=True,
                timeout=10,
            )
            if result.returncode != 0:
                return False

            # fc-match returns comma-separated list of family names
            # (can include aliases / language variants)
            returned_families = [
                f.strip().lower() for f in result.stdout.split(",")
            ]
            requested_lower = family_name.lower()

            # Exact match
            if requested_lower in returned_families:
                return True

            # Partial prefix match (e.g. "Roboto Black" → returned "Roboto")
            for returned in returned_families:
                if (
                    returned.startswith(requested_lower)
                    or requested_lower.startswith(returned)
                ):
                    # Only count as match if the prefix overlap is significant
                    # (at least 4 chars, to avoid false positives on short names)
                    overlap = min(len(returned), len(requested_lower))
                    if overlap >= 4:
                        return True

            # Word-set match: "Open Sans Light" → "Open Sans" is ok
            requested_words = set(requested_lower.split())
            for returned in returned_families:
                returned_words = set(returned.split())
                # All non-weight words in returned appear in requested (or vice versa)
                significant_returned = {
                    w for w in returned_words if w not in WEIGHT_STYLE_SUFFIXES
                }
                significant_requested = {
                    w for w in requested_words if w not in WEIGHT_STYLE_SUFFIXES
                }
                if significant_returned and (
                    significant_returned <= significant_requested
                    or significant_requested <= significant_returned
                ):
                    return True

            return False

        except (subprocess.TimeoutExpired, FileNotFoundError) as exc:
            logger.warning("fc-match check failed for '%s': %s", family_name, exc)
            return False

    # ------------------------------------------------------------------
    # 3. Auto-Install from Google Fonts
    # ------------------------------------------------------------------

    def install_fonts(self, font_names: List[str]) -> Dict[str, str]:
        """Download and install fonts from Google Fonts.

        For weight variants like "Roboto Black" or "Open Sans Light", installs
        the full base family (e.g. "Roboto") which includes all weights.

        Returns a dict mapping each font name to one of:
          'installed'  — successfully downloaded and installed
          'already_installed' — was already present, nothing to do
          'skipped'    — system/proprietary font, skipping
          'not_found'  — not on Google Fonts
          'failed'     — download or install error
        """
        results: Dict[str, str] = {}
        installed_any = False

        # Deduplicate installs: multiple weight variants of same family
        # should only trigger one download
        installed_base_families: Set[str] = set()

        for name in font_names:
            if _is_system_font(name) or _is_cjk_font(name):
                logger.info("Skipping system/proprietary font: %s", name)
                results[name] = "skipped"
                continue

            if self._is_font_installed(name):
                logger.info("Font already installed: %s", name)
                results[name] = "already_installed"
                continue

            # Check if the base family was already installed this run
            base_name = _strip_weight_suffix(name)
            if base_name in installed_base_families and base_name != name:
                logger.info(
                    "Base family '%s' already installed this run (for variant '%s')",
                    base_name, name,
                )
                results[name] = "installed"
                continue

            logger.info("Installing font: %s", name)
            status = self._download_and_install(name)
            results[name] = status
            if status == "installed":
                installed_any = True
                installed_base_families.add(base_name)
                installed_base_families.add(name)

        if installed_any:
            self._rebuild_font_cache()

        return results

    def _find_gf_directory(self, family_name: str) -> Optional[Tuple[str, str]]:
        """Find the Google Fonts GitHub directory for the given font family.

        Returns (category, dir_name) or None if not found.
        Checks ofl/, apache/ categories in order.
        Uses a simple key: lowercase letters/digits only.
        """
        key = _font_dir_key(family_name)

        if key in self._gf_dir_cache:
            cached = self._gf_dir_cache[key]
            return cached  # type: ignore[return-value]

        # Some fonts have known special mappings (family name → dir name differs)
        # We'll try the straightforward approach first, then common variants
        candidates = [key]

        # "Roboto Black" → try "roboto" too (base name after stripping weight suffix)
        base_name = _strip_weight_suffix(family_name)
        base_key = _font_dir_key(base_name)
        if base_key != key and base_key not in candidates:
            candidates.append(base_key)

        # First word of family name (for multi-word families like "Open Sans Light")
        first_word_key = _font_dir_key(family_name.split()[0])
        if first_word_key not in candidates:
            candidates.append(first_word_key)

        # "Source Sans Pro" → "sourcesans3" (known rename in Google Fonts)
        # "Lato Light" → "lato" (weight variant of Lato family)
        special_mappings = {
            "sourcesanspro": "sourcesans3",
            "sourceserifpro": "sourceserif4",
            "sourcecodepro": "sourcecodepro",
            "montserrathairline": "montserrat",
            "gilsans": None,    # Not on Google Fonts (Apple proprietary)
            "gillsans": None,   # Not on Google Fonts
            "dokchampa": None,  # Windows CJK font
        }
        for orig, mapped in special_mappings.items():
            if key == orig:
                if mapped is None:
                    # Mark as definitively not found
                    self._gf_dir_cache[key] = None
                    return None
                if mapped not in candidates:
                    candidates.append(mapped)

        for category in ("ofl", "apache"):
            for candidate in candidates:
                url = f"{GF_API_BASE}/{category}/{candidate}"
                try:
                    result = subprocess.run(
                        ["curl", "-s", "-o", "/dev/null", "-w", "%{http_code}", url],
                        capture_output=True,
                        text=True,
                        timeout=15,
                    )
                    if result.stdout.strip() == "200":
                        found = (category, candidate)
                        self._gf_dir_cache[key] = found
                        return found
                except (subprocess.TimeoutExpired, OSError) as exc:
                    logger.debug("HTTP check failed for %s: %s", url, exc)

        self._gf_dir_cache[key] = None
        return None

    def _get_font_file_urls(self, category: str, dir_name: str) -> List[str]:
        """Return download URLs for all TTF/OTF files in a Google Fonts directory."""
        url = f"{GF_API_BASE}/{category}/{dir_name}"
        try:
            result = subprocess.run(
                ["curl", "-s", url],
                capture_output=True,
                text=True,
                timeout=20,
            )
            if result.returncode != 0:
                return []
            data = json.loads(result.stdout)
            if not isinstance(data, list):
                return []
            return [
                item["download_url"]
                for item in data
                if item.get("type") == "file"
                and item.get("name", "").lower().endswith((".ttf", ".otf"))
                and item.get("download_url")
            ]
        except (subprocess.TimeoutExpired, json.JSONDecodeError, OSError) as exc:
            logger.debug("Failed to list files for %s/%s: %s", category, dir_name, exc)
            return []

    def _download_font_file(self, url: str, dest_path: Path) -> bool:
        """Download a font file using curl. Returns True on success."""
        try:
            result = subprocess.run(
                ["curl", "-sL", "-o", str(dest_path), url],
                capture_output=True,
                timeout=self.timeout,
            )
            if result.returncode == 0 and dest_path.exists() and dest_path.stat().st_size > 0:
                return True
            logger.debug(
                "curl failed for %s (rc=%d, size=%s)",
                url,
                result.returncode,
                dest_path.stat().st_size if dest_path.exists() else "N/A",
            )
            return False
        except (subprocess.TimeoutExpired, OSError) as exc:
            logger.debug("Download error for %s: %s", url, exc)
            return False

    def _download_and_install(self, family_name: str) -> str:
        """Download all weights of a font family and install to user font dir.

        Returns 'installed', 'not_found', or 'failed'.
        """
        # Find the Google Fonts directory
        found = self._find_gf_directory(family_name)
        if found is None:
            logger.warning("Font '%s' not found on Google Fonts", family_name)
            return "not_found"

        category, dir_name = found
        logger.info(
            "Found '%s' at google/fonts/%s/%s", family_name, category, dir_name
        )

        # Get list of font file URLs
        file_urls = self._get_font_file_urls(category, dir_name)
        if not file_urls:
            logger.warning("No font files found for %s/%s", category, dir_name)
            return "not_found"

        # Download the Regular weight first for quick verification;
        # then download all other weights
        # Sort so Regular comes first
        def sort_key(url: str) -> int:
            lower = url.lower()
            if "regular" in lower:
                return 0
            if "bold" in lower and "italic" not in lower:
                return 1
            return 2

        file_urls_sorted = sorted(file_urls, key=sort_key)

        downloaded_count = 0
        for url in file_urls_sorted:
            file_name = url.split("/")[-1]
            dest = self.font_dir / file_name
            if dest.exists():
                logger.debug("Already have %s, skipping re-download", file_name)
                downloaded_count += 1
                continue
            logger.debug("Downloading %s → %s", url, dest)
            ok = self._download_font_file(url, dest)
            if ok:
                downloaded_count += 1
            else:
                logger.warning("Failed to download %s", url)

        if downloaded_count == 0:
            return "failed"

        logger.info(
            "Installed %d font files for '%s'", downloaded_count, family_name
        )
        return "installed"

    def _rebuild_font_cache(self) -> None:
        """Run fc-cache to rebuild the fontconfig cache."""
        logger.info("Rebuilding font cache…")
        try:
            subprocess.run(
                ["fc-cache", "-fv"],
                capture_output=True,
                timeout=60,
            )
            logger.info("Font cache rebuilt.")
        except (subprocess.TimeoutExpired, FileNotFoundError) as exc:
            logger.warning("fc-cache failed: %s", exc)

    # ------------------------------------------------------------------
    # 4. All-in-one API
    # ------------------------------------------------------------------

    def ensure_fonts(self, pptx_path: str) -> Dict[str, str]:
        """Detect all fonts in a PPTX, check which are missing, install them.

        Returns a full status dict: {font_name: status_string}
        where status is one of:
          'installed'          — freshly installed this run
          'already_installed'  — was already present
          'skipped'            — system/proprietary font, not attempted
          'not_found'          — not available on Google Fonts
          'failed'             — download/install error
        """
        logger.info("Detecting fonts in %s", pptx_path)
        all_fonts = self.detect_fonts(pptx_path)
        logger.info("Detected %d unique fonts: %s", len(all_fonts), all_fonts)

        missing = self.check_missing(all_fonts)
        logger.info(
            "%d fonts missing: %s",
            len(missing),
            missing if missing else "none",
        )

        # Build partial status for already-installed / skipped
        results: Dict[str, str] = {}
        to_install = []
        for name in all_fonts:
            if _is_system_font(name) or _is_cjk_font(name):
                results[name] = "skipped"
            elif name not in missing:
                results[name] = "already_installed"
            else:
                to_install.append(name)

        # Install missing fonts
        if to_install:
            install_results = self.install_fonts(to_install)
            results.update(install_results)

        return results


# ---------------------------------------------------------------------------
# Integration point — call at top of processing pipeline
# ---------------------------------------------------------------------------

def prepare_fonts_for_deck(pptx_path: str) -> None:
    """Call this before any LibreOffice rendering to ensure fonts are available.

    Args:
        pptx_path: Path to the PPTX file being rendered.
    """
    mgr = FontManager()
    report = mgr.ensure_fonts(pptx_path)
    for name, status in report.items():
        if status == "failed":
            logger.warning(
                "Font '%s' could not be installed — rendering may be degraded", name
            )
        elif status == "installed":
            logger.info("Font '%s' was newly installed.", name)
        elif status == "not_found":
            logger.warning(
                "Font '%s' not found on Google Fonts — rendering may be degraded", name
            )


# ---------------------------------------------------------------------------
# CLI Interface
# ---------------------------------------------------------------------------

def _cli_main(pptx_path: str) -> None:
    """Run font detection, check, and install for a PPTX file and print report."""
    from pathlib import Path

    deck_name = Path(pptx_path).name

    print(f"Font Detection Report for {deck_name}")
    print("=" * (len("Font Detection Report for ") + len(deck_name)))

    mgr = FontManager()

    # Step 1: Detect
    all_fonts = mgr.detect_fonts(pptx_path)
    print(f"Fonts found: {', '.join(all_fonts) if all_fonts else '(none)'}")

    # Step 2: Check installed vs missing
    already_installed = []
    missing = []
    skipped = []
    for name in all_fonts:
        if _is_system_font(name) or _is_cjk_font(name):
            skipped.append(name)
        elif mgr._is_font_installed(name):
            already_installed.append(name)
        else:
            missing.append(name)

    if already_installed:
        print(f"Already installed: {', '.join(already_installed)}")
    if skipped:
        print(f"Skipped (system/proprietary): {', '.join(skipped)}")
    if missing:
        print(f"Missing: {', '.join(missing)}")
    else:
        print("Missing: (none)")

    # Step 3: Install missing fonts
    if missing:
        # Install one at a time for user-visible progress, but batch the
        # fc-cache rebuild at the very end.
        install_results: Dict[str, str] = {}
        installed_base_families: Set[str] = set()
        installed_any = False

        for name in missing:
            print(f"Installing {name}...", end=" ", flush=True)

            if _is_system_font(name) or _is_cjk_font(name):
                status = "skipped"
            elif mgr._is_font_installed(name):
                status = "already_installed"
            else:
                # Check if base family was already downloaded this session
                base_name = _strip_weight_suffix(name)
                if base_name in installed_base_families and base_name != name:
                    status = "installed"
                else:
                    # Temporarily suppress fc-cache in install_fonts by
                    # calling _download_and_install directly
                    status = mgr._download_and_install(name)
                    if status == "installed":
                        installed_any = True
                        installed_base_families.add(base_name)
                        installed_base_families.add(name)

            install_results[name] = status

            if status == "installed":
                print("OK")
            elif status == "already_installed":
                print("already installed")
            elif status == "not_found":
                print("NOT FOUND on Google Fonts")
            elif status == "skipped":
                print("skipped (system font)")
            else:
                print("FAILED")

        # Single fc-cache rebuild after all installs
        if installed_any:
            print("Rebuilding font cache...", end=" ", flush=True)
            mgr._rebuild_font_cache()
            print("done")

        # Summary
        failed = [n for n, s in install_results.items() if s in ("failed", "not_found")]
        if not failed:
            print("All fonts ready.")
        else:
            print(
                f"Warning: could not install {len(failed)} font(s): "
                + ", ".join(failed)
            )
    else:
        print("All fonts already available.")


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    if len(sys.argv) < 2:
        print(f"Usage: python -m slidearabi.font_manager <path/to/deck.pptx>")
        sys.exit(1)
    _cli_main(sys.argv[1])
