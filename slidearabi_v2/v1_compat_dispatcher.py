"""
v1_compat_dispatcher.py — V1 fallback dispatcher for phased v2 rollout.

Wraps v1 SlideContentTransformer methods so the v2 EngineRouter can
fall back to v1 behaviour for any ShapeRole not yet handled by v2.

Each ShapeRole maps to a sequence of v1 method calls that reproduce
the equivalent v1 behaviour for that role. The dispatcher requires
an existing v1 SlideContentTransformer instance (already initialised
with the same presentation, translations, etc.).
"""

from __future__ import annotations

import logging
from typing import TYPE_CHECKING, Callable, Dict, List, Optional

from slidearabi_v2.shape_classifier import ShapeRole

if TYPE_CHECKING:
    from slidearabi.rtl_transforms import SlideContentTransformer

logger = logging.getLogger(__name__)


class V1CompatDispatcher:
    """
    Dispatches v1 fix methods for shapes the v2 engine doesn't handle yet.

    Usage:
        sct_v1 = SlideContentTransformer(prs, ...)
        dispatcher = V1CompatDispatcher(sct_v1)
        changes = dispatcher.dispatch(shape, role, layout_type, slide_number)
    """

    def __init__(self, v1_transformer: 'SlideContentTransformer'):
        self._v1 = v1_transformer

        # Role → handler mapping
        self._handlers: Dict[ShapeRole, Callable] = {
            ShapeRole.PLACEHOLDER:   self._handle_placeholder,
            ShapeRole.TABLE:         self._handle_table,
            ShapeRole.CHART:         self._handle_chart,
            ShapeRole.CONNECTOR:     self._handle_connector,
            ShapeRole.DIRECTIONAL:   self._handle_directional,
            ShapeRole.BACKGROUND:    self._handle_background,
            ShapeRole.BLEED:         self._handle_bleed,
            ShapeRole.FOOTER:        self._handle_footer,
            ShapeRole.BADGE:         self._handle_badge,
            ShapeRole.LOGO:          self._handle_logo,
            ShapeRole.OVERLAY:       self._handle_overlay,
            ShapeRole.PANEL_LEFT:    self._handle_panel,
            ShapeRole.PANEL_RIGHT:   self._handle_panel,
            ShapeRole.DECORATIVE:    self._handle_decorative,
            ShapeRole.GROUP:         self._handle_group,
            ShapeRole.CONTENT_IMAGE: self._handle_content_image,
            ShapeRole.CONTENT_TEXT:  self._handle_content_text,
            ShapeRole.UNKNOWN:       self._handle_unknown,
        }

    def dispatch(
        self,
        shape,
        role: ShapeRole,
        layout_type: str = 'cust',
        slide_number: int = 0,
        shapes: Optional[List] = None,
    ) -> int:
        """
        Apply v1-equivalent transforms for the given shape role.

        Args:
            shape: python-pptx Shape object.
            role: The classified ShapeRole.
            layout_type: Slide layout type string.
            slide_number: 1-indexed slide number.
            shapes: All shapes on the slide (needed for some v1 fixes).

        Returns:
            Number of changes made.
        """
        handler = self._handlers.get(role, self._handle_unknown)
        try:
            return handler(shape, layout_type, slide_number, shapes)
        except Exception as exc:
            logger.warning(
                'V1CompatDispatcher: role=%s shape="%s" error: %s',
                role.name, getattr(shape, 'name', '?'), exc,
            )
            return 0

    # ── Role handlers ─────────────────────────────────────────────────────

    def _handle_placeholder(
        self, shape, layout_type: str, slide_number: int, shapes
    ) -> int:
        """v1: translate text + set RTL alignment."""
        changes = 0
        changes += self._v1_apply_translation(shape)
        changes += self._v1_set_rtl(shape)
        return changes

    def _handle_table(
        self, shape, layout_type: str, slide_number: int, shapes
    ) -> int:
        """v1: _transform_table_rtl."""
        return self._v1._transform_table_rtl(shape)

    def _handle_chart(
        self, shape, layout_type: str, slide_number: int, shapes
    ) -> int:
        """v1: _transform_chart_rtl."""
        return self._v1._transform_chart_rtl(shape)

    def _handle_connector(
        self, shape, layout_type: str, slide_number: int, shapes
    ) -> int:
        """v1: mirror + toggle flipH on connector."""
        changes = 0
        sw = self._v1._slide_width
        if self._v1._mirror_freeform_shape(shape, sw):
            changes += 1
        return changes

    def _handle_directional(
        self, shape, layout_type: str, slide_number: int, shapes
    ) -> int:
        """v1: mirror + handle directional preset swap."""
        changes = 0
        sw = self._v1._slide_width
        if self._v1._should_mirror_shape(shape, layout_type):
            if self._v1._mirror_freeform_shape(shape, sw):
                changes += 1
        changes += self._v1_set_rtl(shape)
        return changes

    def _handle_background(
        self, shape, layout_type: str, slide_number: int, shapes
    ) -> int:
        """v1: backgrounds are kept in place; only RTL text properties."""
        return self._v1_set_rtl(shape)

    def _handle_bleed(
        self, shape, layout_type: str, slide_number: int, shapes
    ) -> int:
        """v1: mirror + translate + RTL."""
        changes = 0
        sw = self._v1._slide_width
        if self._v1._mirror_freeform_shape(shape, sw):
            changes += 1
        changes += self._v1_apply_translation(shape)
        changes += self._v1_set_rtl(shape)
        return changes

    def _handle_footer(
        self, shape, layout_type: str, slide_number: int, shapes
    ) -> int:
        """v1: mirror + translate + RTL."""
        changes = 0
        sw = self._v1._slide_width
        if self._v1._mirror_freeform_shape(shape, sw):
            changes += 1
        changes += self._v1_apply_translation(shape)
        changes += self._v1_set_rtl(shape)
        return changes

    def _handle_badge(
        self, shape, layout_type: str, slide_number: int, shapes
    ) -> int:
        """v1: mirror (reposition)."""
        changes = 0
        sw = self._v1._slide_width
        if self._v1._mirror_freeform_shape(shape, sw):
            changes += 1
        changes += self._v1_apply_translation(shape)
        changes += self._v1_set_rtl(shape)
        return changes

    def _handle_logo(
        self, shape, layout_type: str, slide_number: int, shapes
    ) -> int:
        """v1: mirror small images."""
        sw = self._v1._slide_width
        return 1 if self._v1._mirror_freeform_shape(shape, sw) else 0

    def _handle_overlay(
        self, shape, layout_type: str, slide_number: int, shapes
    ) -> int:
        """v1: overlay shapes keep position, translate + RTL."""
        changes = 0
        changes += self._v1_apply_translation(shape)
        changes += self._v1_set_rtl(shape)
        return changes

    def _handle_panel(
        self, shape, layout_type: str, slide_number: int, shapes
    ) -> int:
        """v1: panel shapes mirror + translate + RTL."""
        changes = 0
        sw = self._v1._slide_width
        if self._v1._mirror_freeform_shape(shape, sw):
            changes += 1
        changes += self._v1_apply_translation(shape)
        changes += self._v1_set_rtl(shape)
        return changes

    def _handle_decorative(
        self, shape, layout_type: str, slide_number: int, shapes
    ) -> int:
        """v1: decoratives keep position; only RTL text properties."""
        return self._v1_set_rtl(shape)

    def _handle_group(
        self, shape, layout_type: str, slide_number: int, shapes
    ) -> int:
        """v1: mirror group + translate + RTL on children."""
        changes = 0
        sw = self._v1._slide_width
        if self._v1._should_mirror_shape(shape, layout_type):
            if self._v1._mirror_freeform_shape(shape, sw):
                changes += 1
        changes += self._v1_apply_translation(shape)
        changes += self._v1_set_rtl(shape)
        return changes

    def _handle_content_image(
        self, shape, layout_type: str, slide_number: int, shapes
    ) -> int:
        """v1: mirror content images."""
        changes = 0
        sw = self._v1._slide_width
        if self._v1._should_mirror_shape(shape, layout_type):
            if self._v1._mirror_freeform_shape(shape, sw):
                changes += 1
        return changes

    def _handle_content_text(
        self, shape, layout_type: str, slide_number: int, shapes
    ) -> int:
        """v1: mirror + translate + RTL."""
        changes = 0
        sw = self._v1._slide_width
        if self._v1._should_mirror_shape(shape, layout_type):
            if self._v1._mirror_freeform_shape(shape, sw):
                changes += 1
        changes += self._v1_apply_translation(shape)
        changes += self._v1_set_rtl(shape)
        return changes

    def _handle_unknown(
        self, shape, layout_type: str, slide_number: int, shapes
    ) -> int:
        """v1: conservative — mirror + translate + RTL."""
        changes = 0
        sw = self._v1._slide_width
        if self._v1._should_mirror_shape(shape, layout_type):
            if self._v1._mirror_freeform_shape(shape, sw):
                changes += 1
        changes += self._v1_apply_translation(shape)
        changes += self._v1_set_rtl(shape)
        return changes

    # ── Helpers ────────────────────────────────────────────────────────────

    def _v1_apply_translation(self, shape) -> int:
        """Apply v1 text translation to a shape."""
        if not getattr(shape, 'has_text_frame', False):
            return 0
        try:
            return self._v1._apply_translation_to_shape(shape)
        except AttributeError:
            # Fallback: v1 may use different method name
            try:
                return self._v1._translate_shape_text(shape)
            except AttributeError:
                return 0

    def _v1_set_rtl(self, shape) -> int:
        """Apply v1 RTL alignment to a shape."""
        if not getattr(shape, 'has_text_frame', False):
            return 0
        try:
            return self._v1._set_rtl_alignment_unconditional(shape)
        except AttributeError:
            return 0

    def dispatch_post_processing(
        self,
        shapes: List,
        slide_number: int,
    ) -> int:
        """
        Apply v1 post-processing fixes (Fix 11, 19, 22).

        Call after per-shape dispatch to apply text-only fixups.
        """
        changes = 0
        # Fix 11: wrap=none expansion
        for shape in shapes:
            try:
                if getattr(shape, 'has_text_frame', False):
                    changes += self._v1._fix_wrap_none_for_arabic(shape)
            except (AttributeError, Exception):
                pass

        # Fix 19: bidi base direction
        try:
            changes += self._v1._fix_bidi_base_direction(shapes, slide_number)
        except (AttributeError, Exception):
            pass

        # Fix 22: Arabic autofit
        try:
            changes += self._v1._apply_arabic_autofit(shapes, slide_number)
        except (AttributeError, Exception):
            pass

        return changes
