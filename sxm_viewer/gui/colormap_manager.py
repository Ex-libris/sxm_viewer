"""Reactive colormap-selection state: base name + reversal flag.

``ColormapManager`` is the logic layer between colormap-picking UI (the
gallery, combos, menus) and renderers. It holds two independent axes of
state — the selected *base* colormap name (never carrying the ``_r``
suffix) and an ``is_reversed`` flag — twice over:

- ``pending`` — what the UI is currently manipulating (drives Live
  Preview via ``pending_changed``);
- ``applied`` — what the app has committed (``applied_changed``, emitted
  only by ``apply()``).

Renderers subscribe to the signals and resolve pairs through
``cmap_registry.get_colormap`` (programmatic ``Colormap.reversed()``, no
``_r`` filename lookups); widgets call the setters. The name and the
flag are set independently (``select_name`` preserves the current
direction; ``set_reversed``/``toggle_reversed`` preserve the name), so
UI components can emit the two values separately.

Deliberately viewer-agnostic (no ``SXMGridViewer`` coupling), which is
why it lives here next to ``theme.py``/``palettes.py`` rather than in
``gui/controllers/``.
"""
from __future__ import annotations

from .._shared import QtCore
from .. import cmap_registry


class ColormapManager(QtCore.QObject):
    # (base_name, is_reversed) — base name never carries the _r suffix.
    pending_changed = QtCore.pyqtSignal(str, bool)
    applied_changed = QtCore.pyqtSignal(str, bool)

    def __init__(self, name="viridis", is_reversed=False, parent=None):
        super().__init__(parent)
        base, name_rev = cmap_registry.split_cmap_name(name)
        self._pending = (base, bool(is_reversed) ^ name_rev)
        self._applied = self._pending

    # --- pending selection (Live Preview state) ---

    @property
    def selected_name(self):
        return self._pending[0]

    @property
    def is_reversed(self):
        return self._pending[1]

    def select_name(self, name):
        """Set the name axis only; the current direction is preserved.

        A legacy ``_r``-suffixed string is accepted as an explicit full
        spec (it forces the reversed direction)."""
        base, name_rev = cmap_registry.split_cmap_name(name)
        self._set_pending(base, True if name_rev else self._pending[1])

    def set_reversed(self, flag):
        """Set the direction axis only; the current name is preserved."""
        self._set_pending(self._pending[0], bool(flag))

    def toggle_reversed(self):
        self._set_pending(self._pending[0], not self._pending[1])

    def select(self, name, is_reversed):
        """Set both axes at once. An ``_r`` suffix in ``name`` folds into
        the flag with the same XOR rule as ``cmap_registry.get_colormap``."""
        base, name_rev = cmap_registry.split_cmap_name(name)
        self._set_pending(base, bool(is_reversed) ^ name_rev)

    def _set_pending(self, base, is_reversed):
        new = (str(base), bool(is_reversed))
        if new == self._pending:
            return
        self._pending = new
        self.pending_changed.emit(*new)

    # --- applied selection (committed state) ---

    @property
    def applied_name(self):
        return self._applied[0]

    @property
    def applied_reversed(self):
        return self._applied[1]

    def apply(self):
        """Commit the pending selection; emits ``applied_changed`` only
        when it actually differs from the last applied value."""
        if self._applied != self._pending:
            self._applied = self._pending
            self.applied_changed.emit(*self._applied)

    def revert(self):
        """Discard the pending selection, restoring the applied one
        (emits ``pending_changed`` so live previews roll back too)."""
        self._set_pending(*self._applied)

    # --- resolution helpers ---

    def pending_cmap_name(self):
        """Joined legacy string for the pending pair (persistence form)."""
        return cmap_registry.join_cmap_name(*self._pending)

    def applied_cmap_name(self):
        return cmap_registry.join_cmap_name(*self._applied)

    def pending_colormap(self):
        """Resolved matplotlib Colormap for the pending pair."""
        return cmap_registry.get_colormap(*self._pending)
