"""
engine_router.py — Feature-flag routing for phased v2 rollout.

Controls which ShapeRoles are handled by the v2 engine vs. falling
back to v1 via the V1CompatDispatcher.  Configuration is driven by
environment variables so ops can toggle per-role routing without
code deploys.

Environment variables
---------------------
SLIDEARABI_ENGINE_VERSION
    'v1'   — everything goes through V1CompatDispatcher (safe default)
    'v2'   — everything goes through v2 engine
    'dual' — per-role routing via ENABLED/EXCLUDED lists (default)

SLIDEARABI_V2_ENABLED_ROLES
    Comma-separated ShapeRole names handled by v2.
    Default (Phase 1): CONNECTOR,PLACEHOLDER,BACKGROUND,BLEED,OVERLAY

SLIDEARABI_V2_EXCLUDED_ROLES
    Comma-separated ShapeRole names forced to v1 even if in ENABLED.
    Takes precedence over ENABLED.  Default: empty.
"""

from __future__ import annotations

import logging
import os
from enum import Enum
from typing import FrozenSet, Optional, Set

from slidearabi_v2.shape_classifier import ShapeRole

logger = logging.getLogger(__name__)

# ── Rollout phase presets ────────────────────────────────────────────────

PHASE_1_ROLES: FrozenSet[ShapeRole] = frozenset({
    ShapeRole.CONNECTOR,
    ShapeRole.PLACEHOLDER,
    ShapeRole.BACKGROUND,
    ShapeRole.BLEED,
    ShapeRole.OVERLAY,
})

PHASE_2_ROLES: FrozenSet[ShapeRole] = PHASE_1_ROLES | frozenset({
    ShapeRole.PANEL_LEFT,
    ShapeRole.PANEL_RIGHT,
    ShapeRole.LOGO,
    ShapeRole.FOOTER,
    ShapeRole.BADGE,
    ShapeRole.DIRECTIONAL,
})

PHASE_3_ROLES: FrozenSet[ShapeRole] = frozenset(ShapeRole)


class EngineVersion(Enum):
    V1 = 'v1'
    V2 = 'v2'
    DUAL = 'dual'


class EngineRouter:
    """
    Decides whether a shape role should be processed by v2 or fall back to v1.

    Usage::

        router = EngineRouter()          # reads env vars once
        if router.use_v2(role):
            # apply v2 transforms
        else:
            # delegate to V1CompatDispatcher
    """

    def __init__(
        self,
        *,
        version: Optional[EngineVersion] = None,
        enabled_roles: Optional[Set[ShapeRole]] = None,
        excluded_roles: Optional[Set[ShapeRole]] = None,
    ) -> None:
        # Resolve engine version
        if version is not None:
            self._version = version
        else:
            raw = os.environ.get('SLIDEARABI_ENGINE_VERSION', 'dual').strip().lower()
            try:
                self._version = EngineVersion(raw)
            except ValueError:
                logger.warning(
                    'Unknown SLIDEARABI_ENGINE_VERSION=%r, defaulting to dual', raw,
                )
                self._version = EngineVersion.DUAL

        # Resolve enabled roles
        if enabled_roles is not None:
            self._enabled: FrozenSet[ShapeRole] = frozenset(enabled_roles)
        else:
            self._enabled = self._parse_roles_env(
                'SLIDEARABI_V2_ENABLED_ROLES',
                default=PHASE_1_ROLES,
            )

        # Resolve excluded roles
        if excluded_roles is not None:
            self._excluded: FrozenSet[ShapeRole] = frozenset(excluded_roles)
        else:
            self._excluded = self._parse_roles_env(
                'SLIDEARABI_V2_EXCLUDED_ROLES',
                default=frozenset(),
            )

        # Pre-compute effective set
        self._effective: FrozenSet[ShapeRole] = self._enabled - self._excluded

        logger.info(
            'EngineRouter: version=%s enabled=%d excluded=%d effective=%d',
            self._version.value,
            len(self._enabled),
            len(self._excluded),
            len(self._effective),
        )

    # ── Public API ───────────────────────────────────────────────────────

    def use_v2(self, role: ShapeRole) -> bool:
        """Return True if this role should be handled by the v2 engine."""
        if self._version is EngineVersion.V1:
            return False
        if self._version is EngineVersion.V2:
            return True
        # DUAL mode — check effective set
        return role in self._effective

    @property
    def version(self) -> EngineVersion:
        return self._version

    @property
    def enabled_roles(self) -> FrozenSet[ShapeRole]:
        return self._enabled

    @property
    def excluded_roles(self) -> FrozenSet[ShapeRole]:
        return self._excluded

    @property
    def effective_roles(self) -> FrozenSet[ShapeRole]:
        """Roles actually handled by v2 (enabled minus excluded)."""
        return self._effective

    # ── Internals ────────────────────────────────────────────────────────

    @staticmethod
    def _parse_roles_env(
        var_name: str,
        default: FrozenSet[ShapeRole],
    ) -> FrozenSet[ShapeRole]:
        """Parse a comma-separated list of ShapeRole names from an env var."""
        raw = os.environ.get(var_name, '').strip()
        if not raw:
            return default

        roles: Set[ShapeRole] = set()
        for token in raw.split(','):
            token = token.strip().upper()
            if not token:
                continue
            try:
                roles.add(ShapeRole[token])
            except KeyError:
                logger.warning(
                    '%s: unknown role %r (skipped)', var_name, token,
                )
        return frozenset(roles) if roles else default
