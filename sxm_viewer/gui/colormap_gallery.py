"""Colormap gallery: a lazily-rendered, sortable card grid for picking a
colormap.

``ColormapGallery`` shows one card per *base* colormap name — the ``_r``
naming convention never reaches this layer; reversal is a per-selection
flag resolved programmatically by ``cmap_registry.get_colormap``. Each
card carries a small reverse (🔄) region that acts independently of plain
card clicks.

Ordering is delegated to ``cmap_sorting.ColormapSorter``: the gallery
just asks for a list of ``(section_label, [names])`` sections for the
current strategy and renders them. A "Sort by" dropdown in the header
switches strategy (Function / Color / Recently used / Name) and
re-renders; the functional and color strategies inject full-width
section headers between the cards.

The gallery is deliberately dumb UI: it emits ``name_selected`` and
``reverse_requested`` and *renders* whatever selection it is told about
via ``set_selection`` — the actual colormap selection state lives in a
``ColormapManager`` (see ``ColormapGalleryDialog`` for the canonical
wiring, including the pending/applied Live Preview split).

Large-list efficiency: cards and headers share one flat model + one
delegate that builds each gradient pixmap lazily on first paint and
caches it, and the view uses batched layout — with the optional extras
package installed (hundreds of maps) only the visible viewport ever pays
LUT-sampling cost, whatever the sort order.
"""
from __future__ import annotations

from .._shared import (
    QtCore,
    QtWidgets,
    QImage,
    QPainter,
    QPen,
    QPixmap,
    np,
)
from .. import cmap_registry
from ..cmap_sorting import (
    ColormapSorter,
    STRATEGIES,
    STRATEGY_LABELS,
    DEFAULT_STRATEGY,
)
from .colormap_manager import ColormapManager


_ROW_CMAP = "cmap"
_ROW_HEADER = "header"


class _CmapGalleryModel(QtCore.QAbstractListModel):
    """Flat model over sorted sections. Each row is either a colormap
    card or a full-width section header; a substring filter hides
    non-matching cards (and any header left with no visible members)."""

    KIND_ROLE = QtCore.Qt.UserRole + 1

    def __init__(self, parent=None):
        super().__init__(parent)
        self._sections = []   # [(label_or_None, [base_name, ...])]
        self._filter = ""
        self._rows = []       # [(_ROW_HEADER, label) | (_ROW_CMAP, name)]

    def set_sections(self, sections):
        self.beginResetModel()
        self._sections = [(label, list(names)) for label, names in sections]
        self._rebuild_rows()
        self.endResetModel()

    def set_filter(self, text):
        self.beginResetModel()
        self._filter = str(text or "").strip().lower()
        self._rebuild_rows()
        self.endResetModel()

    def _rebuild_rows(self):
        rows = []
        for label, names in self._sections:
            if self._filter:
                names = [n for n in names if self._filter in n.lower()]
            if not names:
                continue
            if label:
                rows.append((_ROW_HEADER, label))
            rows.extend((_ROW_CMAP, n) for n in names)
        self._rows = rows

    def rowCount(self, parent=QtCore.QModelIndex()):
        return 0 if parent.isValid() else len(self._rows)

    def data(self, index, role=QtCore.Qt.DisplayRole):
        if not index.isValid() or not (0 <= index.row() < len(self._rows)):
            return None
        kind, payload = self._rows[index.row()]
        if role == self.KIND_ROLE:
            return kind
        if role in (QtCore.Qt.DisplayRole, QtCore.Qt.UserRole):
            return payload
        if role == QtCore.Qt.ToolTipRole and kind == _ROW_CMAP:
            return f"{payload} — click 🔄 to reverse"
        return None

    def flags(self, index):
        if not index.isValid():
            return QtCore.Qt.NoItemFlags
        kind, _payload = self._rows[index.row()]
        if kind == _ROW_HEADER:
            # Headers are inert dividers: no hover/selection/mouse.
            return QtCore.Qt.NoItemFlags
        return QtCore.Qt.ItemIsEnabled

    def is_header(self, index):
        return (index.isValid()
                and self._rows[index.row()][0] == _ROW_HEADER)

    def index_of(self, name):
        name = str(name or "")
        for row, (kind, payload) in enumerate(self._rows):
            if kind == _ROW_CMAP and payload == name:
                return self.index(row, 0)
        return QtCore.QModelIndex()


class _CmapCardDelegate(QtWidgets.QStyledItemDelegate):
    """Paints either a colormap card (gradient bar + 🔄 region + name) or a
    full-width section header, and turns card clicks into name/reverse
    signals.

    Gradient pixmaps are built on first paint and cached per
    (name, direction) — combined with the item view only painting the
    visible viewport, this is the lazy-thumbnail behaviour that keeps
    huge registries cheap. The currently selected card is painted in the
    selection's live direction, so toggling 🔄 visibly flips it.

    Section headers are sized to (near) the viewport width so IconMode's
    wrapping flow puts them alone on their own row — the "full-width
    injected header" between card sections.
    """

    name_clicked = QtCore.pyqtSignal(str)
    reverse_clicked = QtCore.pyqtSignal(str)

    CARD_SIZE = QtCore.QSize(164, 56)
    HEADER_HEIGHT = 28
    PAD = 6
    GRAD_HEIGHT = 20
    REVERSE_W = 22

    def __init__(self, view=None, parent=None):
        super().__init__(parent)
        self._view = view
        self._pixmap_cache = {}
        self._selected = ("", False)  # (base_name, is_reversed)

    def set_view(self, view):
        self._view = view

    def set_selection(self, name, is_reversed):
        self._selected = (str(name), bool(is_reversed))

    def _is_header(self, index):
        model = index.model()
        return bool(model) and getattr(model, "is_header", None) and model.is_header(index)

    def sizeHint(self, option, index):
        if self._is_header(index):
            width = self.CARD_SIZE.width() * 2
            if self._view is not None:
                # Near-viewport width forces a solo row (nothing fits beside).
                width = max(width, self._view.viewport().width() - self.PAD)
            return QtCore.QSize(int(width), self.HEADER_HEIGHT)
        return QtCore.QSize(self.CARD_SIZE)

    # --- geometry helpers (shared by paint and hit-testing) ---

    def _card_rect(self, option_rect):
        return option_rect.adjusted(2, 2, -2, -2)

    def _reverse_rect(self, card_rect):
        return QtCore.QRect(
            card_rect.right() - self.PAD - self.REVERSE_W,
            card_rect.top() + self.PAD,
            self.REVERSE_W, self.GRAD_HEIGHT,
        )

    def _gradient_rect(self, card_rect):
        return QtCore.QRect(
            card_rect.left() + self.PAD,
            card_rect.top() + self.PAD,
            card_rect.width() - 2 * self.PAD - self.REVERSE_W - 4,
            self.GRAD_HEIGHT,
        )

    def _gradient_pixmap(self, name, is_reversed, width, height):
        key = (str(name), bool(is_reversed), int(width), int(height))
        pixmap = self._pixmap_cache.get(key)
        if pixmap is not None:
            return pixmap
        cmap = cmap_registry.get_colormap(name, is_reversed)
        grad = np.linspace(0.0, 1.0, max(2, int(width)), dtype=np.float32)
        rgba8 = (cmap(grad) * 255).astype(np.uint8)
        rgba8 = np.ascontiguousarray(
            np.repeat(rgba8[np.newaxis, :, :], max(1, int(height)), axis=0))
        h, w = rgba8.shape[:2]
        img = QImage(rgba8.data, w, h, rgba8.strides[0], QImage.Format_RGBA8888)
        pixmap = QPixmap.fromImage(img.copy())
        self._pixmap_cache[key] = pixmap
        return pixmap

    # --- painting / interaction ---

    def paint(self, painter, option, index):
        if self._is_header(index):
            return self._paint_header(painter, option, index)
        name = str(index.data(QtCore.Qt.UserRole) or "")
        if not name:
            return super().paint(painter, option, index)
        selected = name == self._selected[0]
        is_reversed = self._selected[1] if selected else False
        palette = option.palette
        card = self._card_rect(option.rect)

        painter.save()
        painter.setRenderHint(QPainter.Antialiasing, True)
        border = palette.highlight().color() if selected else palette.mid().color()
        painter.setPen(QPen(border, 2 if selected else 1))
        painter.setBrush(palette.base())
        painter.drawRoundedRect(card, 4, 4)

        grad_rect = self._gradient_rect(card)
        painter.drawPixmap(
            grad_rect,
            self._gradient_pixmap(name, is_reversed, grad_rect.width(), grad_rect.height()))

        painter.setPen(QPen(palette.text().color()))
        painter.drawText(self._reverse_rect(card), QtCore.Qt.AlignCenter, "🔄")

        text_rect = QtCore.QRect(
            card.left() + self.PAD, grad_rect.bottom() + 2,
            card.width() - 2 * self.PAD, card.bottom() - grad_rect.bottom() - 4)
        label = name + ("  (reversed)" if selected and is_reversed else "")
        elided = option.fontMetrics.elidedText(label, QtCore.Qt.ElideRight, text_rect.width())
        painter.drawText(text_rect, QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter, elided)
        painter.restore()

    def _paint_header(self, painter, option, index):
        label = str(index.data(QtCore.Qt.UserRole) or "")
        palette = option.palette
        rect = option.rect.adjusted(4, 4, -4, -2)
        painter.save()
        color = palette.text().color()
        text_rect = painter.boundingRect(
            rect, QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter, label)
        # Text on the left, a hairline rule filling the remaining width.
        painter.setPen(QPen(color))
        font = painter.font()
        font.setBold(True)
        painter.setFont(font)
        painter.drawText(rect, QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter, label)
        line_color = palette.mid().color()
        painter.setPen(QPen(line_color, 1))
        y = rect.center().y() + 1
        x0 = text_rect.right() + 8
        if x0 < rect.right():
            painter.drawLine(x0, y, rect.right(), y)
        painter.restore()

    def editorEvent(self, event, model, option, index):
        if self._is_header(index):
            return False
        if (event.type() == QtCore.QEvent.MouseButtonRelease
                and event.button() == QtCore.Qt.LeftButton):
            name = str(index.data(QtCore.Qt.UserRole) or "")
            if name:
                card = self._card_rect(option.rect)
                if self._reverse_rect(card).contains(event.pos()):
                    self.reverse_clicked.emit(name)
                else:
                    self.name_clicked.emit(name)
                return True
        return super().editorEvent(event, model, option, index)


class ColormapGallery(QtWidgets.QWidget):
    """Sortable card grid over base colormap names, with filter + sort
    controls.

    Emits the two selection axes independently (``name_selected`` on a
    card click, ``reverse_requested`` on a 🔄 click); the highlighted
    card is whatever ``set_selection`` last announced — the widget holds
    no colormap-selection state of its own (the view runs with
    NoSelection so the manager stays the single source of truth).

    Sort strategy *is* held here (a view concern, not colormap state):
    ``strategy_changed`` fires when the user switches it, so a host can
    persist the preference.
    """

    name_selected = QtCore.pyqtSignal(str)
    reverse_requested = QtCore.pyqtSignal(str)
    strategy_changed = QtCore.pyqtSignal(str)

    def __init__(self, sorter=None, strategy=DEFAULT_STRATEGY, parent=None):
        super().__init__(parent)
        self._sorter = sorter or ColormapSorter()
        self._names = cmap_registry.base_cmap_names()
        self._strategy = strategy if strategy in STRATEGIES else DEFAULT_STRATEGY

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)

        header = QtWidgets.QHBoxLayout()
        header.setSpacing(6)
        self.filter_edit = QtWidgets.QLineEdit()
        self.filter_edit.setPlaceholderText("Filter colormaps...")
        self.filter_edit.setClearButtonEnabled(True)
        header.addWidget(self.filter_edit, 1)
        header.addWidget(QtWidgets.QLabel("Sort by:"))
        self.sort_combo = QtWidgets.QComboBox()
        for key in STRATEGIES:
            self.sort_combo.addItem(STRATEGY_LABELS[key], key)
        self.sort_combo.setCurrentIndex(max(0, self.sort_combo.findData(self._strategy)))
        self.sort_combo.setToolTip(
            "Order the gallery by colormap function, dominant color, "
            "your recent usage, or name.")
        header.addWidget(self.sort_combo)
        layout.addLayout(header)

        self.view = QtWidgets.QListView()
        self.view.setViewMode(QtWidgets.QListView.IconMode)
        self.view.setResizeMode(QtWidgets.QListView.Adjust)
        self.view.setMovement(QtWidgets.QListView.Static)
        self.view.setWrapping(True)
        # Non-uniform: headers are wider/shorter than cards.
        self.view.setUniformItemSizes(False)
        self.view.setLayoutMode(QtWidgets.QListView.Batched)
        self.view.setSelectionMode(QtWidgets.QAbstractItemView.NoSelection)
        self.view.setSpacing(4)
        self.view.setVerticalScrollMode(QtWidgets.QAbstractItemView.ScrollPerPixel)
        layout.addWidget(self.view, 1)

        self._model = _CmapGalleryModel(self)
        self._delegate = _CmapCardDelegate(self.view, self)
        self.view.setModel(self._model)
        self.view.setItemDelegate(self._delegate)

        self.filter_edit.textChanged.connect(self._model.set_filter)
        self.sort_combo.currentIndexChanged.connect(self._on_sort_changed)
        self._delegate.name_clicked.connect(self.name_selected)
        self._delegate.reverse_clicked.connect(self.reverse_requested)

        self._resort()

    def _resort(self):
        self._model.set_sections(self._sorter.sort(self._names, self._strategy))

    def _on_sort_changed(self, _idx):
        strategy = self.sort_combo.currentData() or DEFAULT_STRATEGY
        if strategy == self._strategy:
            return
        self._strategy = strategy
        self._resort()
        self.strategy_changed.emit(strategy)

    @property
    def strategy(self):
        return self._strategy

    def set_selection(self, name, is_reversed):
        """Reflect the (externally owned) selection in the card grid."""
        self._delegate.set_selection(name, is_reversed)
        self.view.viewport().update()

    def scroll_to(self, name):
        index = self._model.index_of(name)
        if index.isValid():
            self.view.scrollTo(index, QtWidgets.QAbstractItemView.PositionAtCenter)


class ColormapGalleryDialog(QtWidgets.QDialog):
    """Gallery + Apply/Cancel around a caller-owned ``ColormapManager``.

    Canonical wiring for the Live Preview split: every gallery
    interaction only mutates the manager's *pending* selection (whoever
    subscribed to ``pending_changed`` — e.g. the preview canvas fast
    recolor path — updates in real time), Apply commits it
    (``applied_changed``), and Cancel/close reverts, which rolls the live
    preview back through the same ``pending_changed`` signal.

    An optional ``sorter`` (a ``ColormapSorter``, typically carrying the
    host's usage provider) and ``strategy`` seed the gallery ordering;
    ``strategy_changed`` is re-exposed so the host can persist the choice.
    """

    strategy_changed = QtCore.pyqtSignal(str)

    def __init__(self, manager: ColormapManager, sorter=None,
                 strategy=DEFAULT_STRATEGY, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Colormap Gallery")
        self.setMinimumSize(600, 440)
        # A plain QDialog only gets a close (and context-help) button; give
        # this resizable browser real minimize/maximize controls and drop
        # the useless "?" hint.
        self.setWindowFlags(
            QtCore.Qt.Window
            | QtCore.Qt.WindowMinimizeButtonHint
            | QtCore.Qt.WindowMaximizeButtonHint
            | QtCore.Qt.WindowCloseButtonHint)
        self._manager = manager

        layout = QtWidgets.QVBoxLayout(self)
        self.gallery = ColormapGallery(sorter=sorter, strategy=strategy, parent=self)
        layout.addWidget(self.gallery, 1)

        self.selection_label = QtWidgets.QLabel()
        layout.addWidget(self.selection_label)

        buttons = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Apply | QtWidgets.QDialogButtonBox.Cancel)
        apply_btn = buttons.button(QtWidgets.QDialogButtonBox.Apply)
        apply_btn.setDefault(True)
        apply_btn.clicked.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.gallery.name_selected.connect(manager.select_name)
        self.gallery.reverse_requested.connect(self._on_reverse_requested)
        self.gallery.strategy_changed.connect(self.strategy_changed)
        manager.pending_changed.connect(self._on_pending_changed)
        self._on_pending_changed(manager.selected_name, manager.is_reversed)
        self.gallery.scroll_to(manager.selected_name)

    def _on_reverse_requested(self, name):
        """🔄 toggles the current selection, or selects a new map reversed."""
        if name == self._manager.selected_name:
            self._manager.toggle_reversed()
        else:
            self._manager.select(name, True)

    def _on_pending_changed(self, name, is_reversed):
        self.gallery.set_selection(name, is_reversed)
        self.selection_label.setText(
            f"Selected: {name}" + (" (reversed)" if is_reversed else ""))

    def accept(self):
        self._manager.apply()
        super().accept()

    def reject(self):
        self._manager.revert()
        super().reject()

    def closeEvent(self, event):
        # Closing via the title bar counts as Cancel: roll back the preview.
        self._manager.revert()
        super().closeEvent(event)
