"""One colormap menu for every right-click menu that picks a colormap.

Used by the thumbnail menu, the image menu of the preview / pop-ups and the
figure window, so picking a colormap looks and works the same everywhere:

- a search field at the top (type to filter all ~170 maps, ~1000 with the
  optional ``colormaps`` package; Enter or click picks),
- the current map's direction ("Reversed") and a star to make it - exactly
  as used, ``X`` or ``X_r`` - a favourite,
- *Favourites* (starred), *Recent* (last used) and *Suggested* (the curated
  shortlist of ``cmap_registry`` for the context), each with a gradient
  swatch,
- "All colormaps..." -> the Colormap Gallery (search, sort, preview).

``add_cmap_submenu`` adds it as "Colormap: <name>" with the current map's
gradient as the icon, so the colormap is visible without opening anything.

``CmapPrefs`` keeps favourites, recent picks and usage in the app config
dict (keys ``colormap_favourites``, ``colormap_recent`` - exact names - and
``colormap_usage`` - per base name, shared with the gallery's "usage"
sort). It must be handed the *live* config dict of the
viewer, never a freshly loaded copy - the viewer saves its own dict whole and
would overwrite a separate write.

Viewer-agnostic like ``menu_kit``: no imports from ``main_window``.
"""
from __future__ import annotations

import time

from PyQt5 import QtCore, QtWidgets

from .. import cmap_registry
from .menu_kit import section

_RECENT_SHOWN = 5
_RECENT_KEPT = 12
_SEARCH_LIMIT = 60


def _icon(name, width=96, height=14):
    from .thumbnail_render import _colormap_icon
    try:
        return _colormap_icon(name, width=width, height=height)
    except Exception:
        return None


class CmapPrefs:
    """Favourite colormaps + usage, in ``config`` (None = in memory)."""

    def __init__(self, config=None, save=None):
        self._config = config if isinstance(config, dict) else {}
        self._save = save if isinstance(config, dict) else None

    def _persist(self):
        if callable(self._save):
            try:
                self._save(self._config)
            except Exception:
                pass

    def favourites(self):
        favs = self._config.get("colormap_favourites")
        return [str(n) for n in favs] if isinstance(favs, list) else []

    def is_favourite(self, name):
        return str(name) in self.favourites()

    def toggle_favourite(self, name):
        favs = self.favourites()
        name = str(name)
        if name in favs:
            favs.remove(name)
        else:
            favs.append(name)
        self._config["colormap_favourites"] = favs
        self._persist()

    def usage(self):
        """``{base name: (count, last used)}`` (the gallery's usage sort)."""
        out = {}
        for base, rec in (self._config.get("colormap_usage") or {}).items():
            try:
                count, last = rec
                out[str(base)] = (int(count), float(last))
            except Exception:
                continue
        return out

    def recent(self, n=_RECENT_SHOWN):
        """Last picked maps exactly as used, newest first (falls back to the
        usage record for configs from before this list existed)."""
        names = self._config.get("colormap_recent")
        if isinstance(names, list) and names:
            return [str(x) for x in names][:n]
        return self._recent_from_usage(n)

    def _recent_from_usage(self, n=_RECENT_SHOWN):
        """Last used maps (base names: usage is counted per map in either
        direction, as the gallery's usage sort does), newest first."""
        usage = self.usage()
        return [b for b, _rec in sorted(usage.items(), key=lambda kv: -kv[1][1])][:n]

    def record(self, name):
        base, _rev = cmap_registry.split_cmap_name(str(name))
        usage = dict(self._config.get("colormap_usage") or {})
        count = 0
        try:
            count = int((usage.get(base) or (0, 0))[0])
        except Exception:
            pass
        usage[base] = [count + 1, float(time.time())]
        self._config["colormap_usage"] = usage
        recent = [str(x) for x in (self._config.get("colormap_recent") or []) if str(x) != str(name)]
        self._config["colormap_recent"] = ([str(name)] + recent)[:_RECENT_KEPT]
        self._persist()

    def sort_strategy(self):
        return self._config.get("colormap_sort_strategy")

    def set_sort_strategy(self, strategy):
        self._config["colormap_sort_strategy"] = str(strategy)
        self._persist()


def _close_menu_chain(widget):
    """Close the menu holding ``widget`` and every parent menu."""
    w = widget
    while w is not None:
        if isinstance(w, QtWidgets.QMenu):
            top = w
            while isinstance(top.parentWidget(), QtWidgets.QMenu):
                top = top.parentWidget()
            top.close()
            w.close()
            return
        w = w.parentWidget()


class _CmapRow(QtWidgets.QWidget):
    """One menu entry: check mark, a real-width gradient swatch, the name.
    (A plain QAction icon is squeezed into a 16 px square - the gradient
    came out as a hairline.) Clicking triggers its ``QWidgetAction``."""

    _SWATCH = QtCore.QSize(64, 12)

    def __init__(self, action, name, parent=None):
        super().__init__(parent)
        self._action = action
        self._name = name
        self._hover = False
        self.setMouseTracking(True)
        self.setMinimumHeight(22)
        fm = self.fontMetrics()
        self.setMinimumWidth(22 + self._SWATCH.width() + 10 + fm.horizontalAdvance(name) + 16)
        icon = _icon(name, self._SWATCH.width(), self._SWATCH.height())
        self._pix = icon.pixmap(self._SWATCH) if icon is not None else None

    def enterEvent(self, _event):
        self._hover = True
        self.update()

    def leaveEvent(self, _event):
        self._hover = False
        self.update()

    def mouseReleaseEvent(self, event):
        if event.button() == QtCore.Qt.LeftButton:
            act = self._action
            _close_menu_chain(self)
            act.trigger()

    def paintEvent(self, _event):
        from PyQt5 import QtGui
        p = QtGui.QPainter(self)
        pal = self.palette()
        if self._hover:
            p.fillRect(self.rect(), pal.color(QtGui.QPalette.Highlight))
            fg = pal.color(QtGui.QPalette.HighlightedText)
        else:
            fg = pal.color(QtGui.QPalette.WindowText)
        h = self.height()
        if self._action.isChecked():
            p.setPen(fg)
            p.drawText(QtCore.QRect(4, 0, 16, h), QtCore.Qt.AlignCenter, "✓")
        sw = QtCore.QRect(22, (h - self._SWATCH.height()) // 2, self._SWATCH.width(), self._SWATCH.height())
        if self._pix is not None:
            p.drawPixmap(sw, self._pix)
        p.setPen(QtGui.QColor(0, 0, 0, 90))
        p.drawRect(sw.adjusted(0, 0, -1, -1))
        p.setPen(fg)
        p.drawText(QtCore.QRect(sw.right() + 10, 0, self.width() - sw.right() - 12, h),
                   QtCore.Qt.AlignVCenter | QtCore.Qt.AlignLeft, self._name)


def _add_cmap_row(menu, name, checked, slot):
    act = QtWidgets.QWidgetAction(menu)
    act.setText(name)
    act.setCheckable(True)
    act.setChecked(bool(checked))
    act.triggered.connect(lambda _=False: slot(name))
    act.setDefaultWidget(_CmapRow(act, name, menu))
    menu.addAction(act)
    return act


class _SearchBox(QtWidgets.QWidget):
    """Search field + live result list, embedded at the top of the menu."""

    def __init__(self, on_pick, parent=None):
        super().__init__(parent)
        self._on_pick = on_pick
        lay = QtWidgets.QVBoxLayout(self)
        lay.setContentsMargins(8, 4, 8, 4)
        lay.setSpacing(3)
        self.edit = QtWidgets.QLineEdit(self)
        self.edit.setPlaceholderText("Search colormaps…")
        self.edit.setClearButtonEnabled(True)
        self.edit.setMinimumWidth(220)
        lay.addWidget(self.edit)
        self.results = QtWidgets.QListWidget(self)
        self.results.setIconSize(QtCore.QSize(64, 12))
        self.results.setUniformItemSizes(True)
        self.results.setVisible(False)
        lay.addWidget(self.results)
        self.edit.textChanged.connect(self._filter)
        self.edit.returnPressed.connect(self._pick_first)
        self.edit.installEventFilter(self)
        self.results.itemActivated.connect(self._pick_item)
        self.results.itemClicked.connect(self._pick_item)

    def eventFilter(self, obj, event):
        if obj is self.edit and event.type() == QtCore.QEvent.KeyPress \
                and event.key() == QtCore.Qt.Key_Down and self.results.count():
            self.results.setFocus()
            self.results.setCurrentRow(0)
            return True
        return super().eventFilter(obj, event)

    def _filter(self, text):
        text = text.strip().lower()
        self.results.clear()
        if not text:
            self.results.setVisible(False)
            self._relayout()
            return
        names = cmap_registry.all_cmap_names()
        if not text.endswith("_r"):
            names = [n for n in names if not n.endswith("_r")]
        starts = [n for n in names if n.lower().startswith(text)]
        contains = [n for n in names if text in n.lower() and n not in starts]
        hits = (starts + contains)[:_SEARCH_LIMIT]
        for n in hits:
            item = QtWidgets.QListWidgetItem(n)
            icon = _icon(n, 72, 12)
            if icon is not None:
                item.setIcon(icon)
            self.results.addItem(item)
        if not hits:
            item = QtWidgets.QListWidgetItem("No colormap matches")
            item.setFlags(QtCore.Qt.NoItemFlags)
            self.results.addItem(item)
        # As tall as the hits (up to 8 rows), no empty box under one match.
        row_h = max(16, self.results.sizeHintForRow(0))
        self.results.setFixedHeight(min(self.results.count(), 8) * row_h + 2 * self.results.frameWidth() + 2)
        self.results.setVisible(True)
        self._relayout()

    def _relayout(self):
        # Grow / shrink the menu with the result list.
        self.adjustSize()
        menu = self.parentWidget()
        if isinstance(menu, QtWidgets.QMenu):
            menu.adjustSize()

    def _pick_first(self):
        if self.results.count() and self.results.item(0).flags() & QtCore.Qt.ItemIsEnabled:
            self._pick_item(self.results.item(0))

    def _pick_item(self, item):
        if item is None or not (item.flags() & QtCore.Qt.ItemIsEnabled):
            return
        name = item.text()
        _close_menu_chain(self)
        self._on_pick(name)


def populate_cmap_menu(menu, current, on_pick, *, prefs=None, context="general", on_gallery=None,
                       extras=None, note=None):
    """Fill ``menu`` with the colormap picker.

    ``current``: the map in use (``X`` or ``X_r``) or None when mixed;
    ``on_pick(name)``: apply a joined name; ``prefs``: ``CmapPrefs`` (None =
    no favourites / recent); ``on_gallery()``: open the full gallery;
    ``extras``: ``[(text, slot)]`` appended at the end (e.g. "Use the
    global colormap"); ``note``: a greyed line at the top (what the pick
    applies to)."""
    prefs = prefs or CmapPrefs()

    def pick(name):
        try:
            prefs.record(name)
        except Exception:
            pass
        on_pick(name)

    if note:
        section(menu, note, first=True)
    search = _SearchBox(pick, menu)
    act = QtWidgets.QWidgetAction(menu)
    act.setDefaultWidget(search)
    menu.addAction(act)
    menu.aboutToShow.connect(lambda: QtCore.QTimer.singleShot(0, search.edit.setFocus))

    base, rev = cmap_registry.split_cmap_name(str(current)) if current else (None, False)
    if base:
        a = menu.addAction("Reversed")
        a.setCheckable(True)
        a.setChecked(bool(rev))
        a.setToolTip("Flip the direction of the colormap in use")
        a.triggered.connect(lambda on: pick(cmap_registry.join_cmap_name(base, bool(on))))
        # A favourite is the map exactly as used (X or X_r).
        fav = prefs.is_favourite(current)
        a = menu.addAction(("★ Remove " if fav else "☆ Add ") + f"{current} " +
                           ("from favourites" if fav else "to favourites"))
        a.triggered.connect(lambda: prefs.toggle_favourite(current))

    shown = set()

    def add_group(title, names):
        # Each map once (a favourite is not repeated under Recent).
        names = [n for n in dict.fromkeys(names) if n and n not in shown]
        if not names:
            return
        section(menu, title)
        for name in names:
            _add_cmap_row(menu, name, bool(current) and name == current, pick)
            shown.add(name)

    add_group("Favourites", prefs.favourites())
    add_group("Recent", prefs.recent())
    suggested = cmap_registry.featured_cmap_names(context)
    if current and current not in shown and current not in suggested:
        suggested = [current] + suggested
    add_group("Suggested", suggested)
    if callable(on_gallery):
        menu.addSeparator()
        menu.addAction("All colormaps… (gallery: search, sort, preview)").triggered.connect(
            lambda _=False: on_gallery())
    for text, slot in extras or ():
        menu.addSeparator()
        menu.addAction(text).triggered.connect(lambda _=False, s=slot: s())
    return menu


def add_cmap_submenu(parent_menu, current, on_pick, *, label="Colormap", **kw):
    """"<label>: <current>" submenu (gradient of the current map as icon),
    filled lazily when opened."""
    title = f"{label}: {current}" if current else f"{label}: mixed"
    sub = parent_menu.addMenu(title)
    if current:
        icon = _icon(current, 16, 16)
        if icon is not None:
            sub.menuAction().setIcon(icon)
    filled = []

    def fill():
        if not filled:
            filled.append(True)
            populate_cmap_menu(sub, current, on_pick, **kw)

    sub.aboutToShow.connect(fill)
    return sub


def pick_from_gallery(parent, current, *, prefs=None, on_preview=None):
    """Run the Colormap Gallery for one pick. ``on_preview(name)`` is called
    live for every card clicked (and with ``current`` again on cancel).
    Returns the applied joined name, or None."""
    from ..cmap_sorting import DEFAULT_STRATEGY, ColormapSorter
    from .colormap_gallery import ColormapGalleryDialog
    from .colormap_manager import ColormapManager

    prefs = prefs or CmapPrefs()
    manager = ColormapManager(str(current or "viridis"), parent=parent)
    applied = []

    def _pending(name, is_reversed):
        cmap_registry.get_colormap(name, is_reversed)
        if callable(on_preview):
            on_preview(cmap_registry.join_cmap_name(name, is_reversed))

    def _applied(name, is_reversed):
        cmap_registry.get_colormap(name, is_reversed)
        applied.append(cmap_registry.join_cmap_name(name, is_reversed))

    manager.pending_changed.connect(_pending)
    manager.applied_changed.connect(_applied)
    dlg = ColormapGalleryDialog(manager, sorter=ColormapSorter(usage_provider=prefs.usage),
                                strategy=prefs.sort_strategy() or DEFAULT_STRATEGY, parent=parent)
    dlg.strategy_changed.connect(prefs.set_sort_strategy)
    try:
        dlg.exec_()
    finally:
        manager.deleteLater()
    if applied:
        prefs.record(applied[-1])
        return applied[-1]
    if callable(on_preview) and current:
        on_preview(str(current))
    return None
