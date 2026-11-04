"""Shared building blocks for the app's right-click menus.

Used by the thumbnail menu (``gui/viewer/thumb_menu.py``) and the image
menu of the main preview and every pop-up (``MultiPreviewCanvas.
_show_context_menu``), so they look and behave alike:

- ``section(menu, text)`` - greyed, non-clickable section title. A QLabel
  ``QWidgetAction``, because ``QMenu.addSection`` text is not drawn by the
  native Windows style.
- ``Route`` + ``add_export_section`` - an export registry rendered as
  "Again: ..." (repeat the last export to the same destination, no dialog),
  the user's pinned routes inline, and "More export" (every route plus
  "Customise pinned...").
- ``MenuPrefs`` - pinned routes and the last export for one menu *scope*
  (``"thumbnails"``, ``"view"``), stored in the config dict under
  ``context_menus.<scope>``; with no config it keeps them in memory.

Viewer-agnostic like ``theme.py``: no imports from ``main_window``.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Callable, Iterable, Optional, Sequence

from PyQt5 import QtCore, QtWidgets


# ---------------------------------------------------------------------------
# section headers
# ---------------------------------------------------------------------------
def section(menu, text, *, first=False):
    """Append a greyed section title (preceded by a separator unless
    ``first``)."""
    if not first and not menu.isEmpty():
        menu.addSeparator()
    label = QtWidgets.QLabel(str(text))
    label.setEnabled(False)
    font = label.font()
    font.setBold(True)
    if font.pointSizeF() > 0:
        font.setPointSizeF(max(7.0, font.pointSizeF() * 0.9))
    label.setFont(font)
    label.setContentsMargins(10, 2, 10, 1)
    act = QtWidgets.QWidgetAction(menu)
    act.setDefaultWidget(label)
    act.setEnabled(False)
    menu.addAction(act)
    return act


def checkable(menu, text, checked, slot=None, *, tip=None, enabled=True):
    """Add a checkable action; ``slot(checked)`` runs on toggle."""
    act = menu.addAction(text)
    act.setCheckable(True)
    act.setChecked(bool(checked))
    act.setEnabled(bool(enabled))
    if tip:
        act.setToolTip(tip)
        act.setStatusTip(tip)
    if slot is not None:
        act.triggered.connect(lambda state: slot(bool(state)))
    return act


def choice_group(menu, parent, options, current, slot):
    """Exclusive checkable options [(label, value), ...]; ``slot(value)``."""
    group = QtWidgets.QActionGroup(parent)
    group.setExclusive(True)
    actions = []
    for label, value in options:
        act = menu.addAction(label)
        act.setCheckable(True)
        act.setChecked(value == current)
        group.addAction(act)
        act.triggered.connect(lambda _=False, v=value: slot(v))
        actions.append(act)
    return actions


# ---------------------------------------------------------------------------
# preferences (pinned routes + last export), per menu scope
# ---------------------------------------------------------------------------
class MenuPrefs:
    def __init__(self, scope, default_pinned=(), *, config=None, save=None, legacy=None):
        """``config``: the app config dict (None = in-memory only);
        ``save(config)`` persists it; ``legacy``: an old top-level config key
        read once if the scope has no stored state yet."""
        self.scope = str(scope)
        self.default_pinned = list(default_pinned)
        self._config = config if isinstance(config, dict) else None
        self._save = save
        self._memory = {}
        if self._config is not None and legacy and not self._store_raw():
            old = self._config.get(legacy)
            if isinstance(old, dict) and old:
                self._write(dict(old))

    def _store_raw(self):
        if self._config is None:
            return dict(self._memory)
        menus = self._config.get("context_menus") or {}
        return dict(menus.get(self.scope) or {}) if isinstance(menus, dict) else {}

    def _write(self, data):
        if self._config is None:
            self._memory = dict(data)
            return
        menus = dict(self._config.get("context_menus") or {})
        menus[self.scope] = dict(data)
        self._config["context_menus"] = menus
        if callable(self._save):
            try:
                self._save(self._config)
            except Exception:
                pass

    def pinned(self, valid: Optional[Iterable[str]] = None):
        pinned = self._store_raw().get("pinned")
        if not isinstance(pinned, list):
            pinned = list(self.default_pinned)
        if valid is not None:
            valid = set(valid)
            pinned = [p for p in pinned if p in valid]
        return pinned

    def set_pinned(self, ids):
        data = self._store_raw()
        data["pinned"] = list(ids)
        self._write(data)

    def last(self):
        last = self._store_raw().get("last_export")
        return dict(last) if isinstance(last, dict) else {}

    def record(self, record):
        data = self._store_raw()
        data["last_export"] = dict(record)
        self._write(data)


# ---------------------------------------------------------------------------
# export routes
# ---------------------------------------------------------------------------
@dataclass
class Route:
    """One export route.

    ``run(ctx, dest)`` performs it; ``dest`` is a previously returned record
    (for "Again") or None (ask the user). It returns a JSON-able record of
    the destination (e.g. ``{"out_dir": ...}``) when there is something to
    repeat, else None. ``again(record, pinned)`` returns the "Again" label
    text, or None when the destination is gone or the entry would only
    duplicate a pinned one. ``label`` / ``available`` may depend on ``ctx``.
    """
    id: str
    label: Any  # str or callable(ctx) -> str
    run: Callable[[Any, Optional[dict]], Optional[dict]]
    tip: Any = ""  # str or callable(ctx) -> str
    available: Optional[Callable[[Any], tuple]] = None  # ctx -> (ok, reason)
    again: Optional[Callable[[dict, Sequence[str]], Optional[str]]] = None

    def text(self, ctx):
        return self.label(ctx) if callable(self.label) else str(self.label)

    def tooltip(self, ctx):
        return self.tip(ctx) if callable(self.tip) else str(self.tip or "")

    def status(self, ctx):
        if self.available is None:
            return True, ""
        try:
            ok, reason = self.available(ctx)
            return bool(ok), str(reason or "")
        except Exception as exc:
            return False, str(exc)


def again_folder(kind):
    """``Route.again`` for routes whose record is ``{"out_dir": folder}``."""
    from pathlib import Path

    def again(record, _pinned):
        dest = record.get("out_dir")
        if not dest or not Path(dest).is_dir():
            return None
        return f"{kind} → {Path(dest).name}"
    return again


def execute(route, ctx, prefs, dest=None):
    record = route.run(ctx, dest)
    if record and prefs is not None:
        prefs.record({**record, "route": route.id})
    return record


def _route_action(menu, route, ctx, prefs):
    act = menu.addAction(route.text(ctx))
    tip = route.tooltip(ctx)
    ok, reason = route.status(ctx)
    act.setEnabled(ok)
    if not ok and reason:
        tip = reason
    if tip:
        act.setToolTip(tip)
        act.setStatusTip(tip)
    act.triggered.connect(lambda _=False, r=route: execute(r, ctx, prefs))
    return act


def add_export_section(menu, routes, prefs, ctx, *, parent=None, title="Export", first=False):
    """Export section: Again + pinned routes + "More export" submenu."""
    routes = [r for r in routes if r is not None]
    by_id = {r.id: r for r in routes}
    section(menu, title, first=first)
    pinned = prefs.pinned(by_id) if prefs is not None else []
    last = prefs.last() if prefs is not None else {}
    again_route = by_id.get(last.get("route"))
    if again_route is not None and again_route.again is not None:
        try:
            text = again_route.again(last, pinned)
        except Exception:
            text = None
        if text and again_route.status(ctx)[0]:
            act = menu.addAction("↻ Again: " + text)
            act.setToolTip("Repeat your last export to the same destination")
            act.triggered.connect(lambda _=False, r=again_route, d=dict(last): execute(r, ctx, prefs, d))
    for rid in pinned:
        _route_action(menu, by_id[rid], ctx, prefs)
    more = menu.addMenu("More export" if pinned else "Export")
    more.setToolTipsVisible(True)
    for route in routes:
        _route_action(more, route, ctx, prefs)
    if prefs is not None:
        more.addSeparator()
        custom = more.addAction("Customise pinned...")
        custom.setToolTip("Choose which exports appear directly in this menu")
        custom.triggered.connect(
            lambda _=False: customise_pinned(parent or menu.parentWidget(), routes, prefs, ctx))
    return more


def customise_pinned(parent, routes, prefs, ctx=None, *, title="Pinned export actions"):
    """Checklist dialog choosing the pinned routes; returns True if saved."""
    dlg = QtWidgets.QDialog(parent)
    dlg.setWindowTitle(title)
    layout = QtWidgets.QVBoxLayout(dlg)
    layout.addWidget(QtWidgets.QLabel(
        "Exports shown directly in this right-click menu.\n"
        "Everything else stays under \"More export\"."))
    lst = QtWidgets.QListWidget()
    pinned = set(prefs.pinned())
    for route in routes:
        try:
            text = route.text(ctx) if ctx is not None else str(route.label if not callable(route.label) else route.id)
        except Exception:
            text = route.id
        item = QtWidgets.QListWidgetItem(text.rstrip("."))
        item.setToolTip(route.tooltip(ctx) if ctx is not None else "")
        item.setData(QtCore.Qt.UserRole, route.id)
        item.setFlags(item.flags() | QtCore.Qt.ItemIsUserCheckable)
        item.setCheckState(QtCore.Qt.Checked if route.id in pinned else QtCore.Qt.Unchecked)
        lst.addItem(item)
    layout.addWidget(lst)
    buttons = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
    buttons.accepted.connect(dlg.accept)
    buttons.rejected.connect(dlg.reject)
    layout.addWidget(buttons)
    if dlg.exec_() != QtWidgets.QDialog.Accepted:
        return False
    prefs.set_pinned([lst.item(i).data(QtCore.Qt.UserRole) for i in range(lst.count())
                      if lst.item(i).checkState() == QtCore.Qt.Checked])
    return True
