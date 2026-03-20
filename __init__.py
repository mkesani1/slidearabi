"""
slidearabi — Arabic RTL slide conversion engine.

Docker compatibility: When deployed via Docker, all modules live under
the 'slidearabi' package (e.g., slidearabi.v3_config). However, many
modules use bare imports (e.g., 'import v3_config') that work in local
development but fail inside the Docker package structure.

This __init__.py registers package submodules under their bare names
in sys.modules so that bare imports resolve correctly in both environments.
"""

import importlib
import sys

# Modules that use bare cross-imports internally.
# Register them as top-level aliases so 'import v3_config' works
# even when the actual module path is 'slidearabi.v3_config'.
_ALIASED_MODULES = (
    'vqa_types',
    'v3_config',
    'v3_checks',
    'v3_api_contract',
    'v3_vision_prompts',
)


def _register_bare_aliases():
    """Make package submodules importable by bare name."""
    for mod_name in _ALIASED_MODULES:
        if mod_name in sys.modules:
            continue  # Already registered (e.g., running locally)
        pkg_name = f'slidearabi.{mod_name}'
        if pkg_name in sys.modules:
            sys.modules[mod_name] = sys.modules[pkg_name]
        else:
            # Try to import the package version and alias it
            try:
                mod = importlib.import_module(pkg_name)
                sys.modules[mod_name] = mod
            except ImportError:
                pass  # Module doesn't exist — skip silently


_register_bare_aliases()
