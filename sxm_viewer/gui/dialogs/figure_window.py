"""Figure window for data-aware figure recipes (height series, selection).

One clean window: the figure on a white "paper" (exactly what Copy /
PowerPoint / PDF produce - one matplotlib renderer, ``figures.height_series
.render``), a thin strip with the outputs, and an optional **Format** panel.
Every adjustment can also be made directly on the figure:

* click a frame's Δz header        -> that frame becomes the Δz reference
* hover a header, click its ×      -> drop the frame (Ctrl+Z restores)
* click a channel label (I, Δf)    -> switch / recolour / move / remove it
* click "+ channel"                -> add any other channel the frames have
* drag a colourbar                 -> contrast (upper half = top, lower half
  = bottom, Shift = both); click it to type exact Min/Max; double-click =
  automatic. Per-image colourbars edit just that panel.
* right-click anything             -> the options for that element
  (letters, scale bar, headers, ...)
* double-click a panel             -> show that image in the main viewer

The window talks to its data source (``gui/controllers/figures.py``
``FigureController``) only through a handful of methods; it never reaches
into the main window.
"""
from __future__ import annotations

import copy
import io
import math
from pathlib import Path

from ..._shared import QtCore, QtGui, QtWidgets, log_status, np
from ...cmap_registry import get_cmap as registry_get_cmap
from ...figures import height_series as hs
from ...figures import svg_layers
from ...molecules import interaction as molecule_interaction
from ...molecules import render as molecule_render

_PAPER_MARGIN = 14       # px around the paper inside the view
_CLOSE_SIZE = 14         # px, hover × on frame headers
_LIVE_POLL_MS = 1200     # source-change check while the window is visible
_DRAG_THROTTLE_MS = 30
_CLICK_SLOP_PX = 3

_LETTER_STYLE_NAMES = {"a": "a, b, c", "A": "A, B, C", "(a)": "(a), (b)", "a)": "a), b)", "a.": "a., b."}
_CORNER_NAMES = {"tl": "Top left", "tr": "Top right", "bl": "Bottom left", "br": "Bottom right"}
_LETTER_PLACES = dict(_CORNER_NAMES, out_tl="Outside, above top left", out_tr="Outside, above top right")
_HEADER_NAMES = {"auto": "Automatic (Δz / bias / size)", "dz": "Δz to the reference",
                 "zabs": "Absolute z", "bias": "Bias", "size": "Frame size", "name": "File name",
                 "none": "None (no header)"}
_MODE_TITLES = {"height": "Height series", "bias": "Bias series", "zoom": "Overview + zooms",
                "selection": "Figure"}
_CBAR_STYLES = {"outside": "Numbers outside", "inside": "Numbers inside the bar",
                "minmax": "min / max outside", "minmax_in": "min / max inside the bar"}
_CBAR_TICKS = {0: "None", 2: "2 (ends)", 3: "3", 5: "5"}
_CBAR_TEXT = {"rotated": "Rotated text", "horizontal": "Horizontal text"}
_AUTO_COLORS = {"auto": "Automatic (contrast)", "white": "White", "black": "Black"}
_ZOOM_COLORS = {"color": "One colour per zoom", "white": "White", "black": "Black"}
_DPI_CHOICES = (300, 600, 1200)
_WHEEL_COALESCE_S = 1.0
_MODE_HINTS = {
    "height": "Click a Δz header: reference  ·  ",
    "bias": "",
    "zoom": "",
    "selection": "",
}
_SCALE_NAMES = {"all": "Every image", "one": "One image", "off": "Off"}
_SCALE_LABELS = {"first": "First image only (once per length)", "all": "Every image", "none": "None"}
_CBAR_LABEL_POS = {"inside": "Centred inside the bar", "outside": "Outside the bar"}
_ROTATE_TALL = {"off": "Keep as acquired", "cw": "Turn to landscape (clockwise)",
                "ccw": "Turn to landscape (counter-clockwise)"}


class _MoleculePainter:
    """Draws molecular models into a figure panel with the app's one
    molecule renderer (``molecules.render``, also used by the preview and
    pop-ups), at the panel's print scale. Returns the artists it added."""

    def __init__(self, molecules, settings):
        self.molecules = list(molecules or [])
        self.settings = dict(settings or {})

    def draw(self, ax, scale=1.0):
        entries = molecule_render.draw_molecules(ax, self.molecules, self.settings, scale=scale)
        return molecule_render.artists_of(entries)


def _mol_scale(ax, fig, user=1.0):
    """Marker scale for a panel (``molecules.render.panel_scale``)."""
    return molecule_render.panel_scale(ax, user)


def _mol_palette_choices():
    """Figure atom-palette options: follow the preview, or a fixed one."""
    from ...molecules.model import available_atom_palettes, palette_label
    out = {"preview": "As preview"}
    out.update((k, palette_label(k)) for k in available_atom_palettes())
    return out


_gizmo_vectors = molecule_interaction.gizmo_vectors


_GIZMO_R = 38.0   # px


def _molecule_from(d):
    from ..canvases.molecular_overlay import Molecule
    return Molecule.from_dict(d)


def _figure_classes():
    from matplotlib.backends.backend_agg import FigureCanvasAgg
    from matplotlib.figure import Figure
    return Figure, FigureCanvasAgg


def _inside(box, x, y, pad=0.0):
    return box[0] - pad <= x <= box[2] + pad and box[1] - pad <= y <= box[3] + pad


class _FigureView(QtWidgets.QWidget):
    """Paints the rendered figure centred on a white sheet; maps mouse
    positions back to figure inches for the window's hit-testing."""

    hovered = QtCore.pyqtSignal(object)            # hit dict or None
    pressed = QtCore.pyqtSignal(object, object)    # hit, QMouseEvent
    moved = QtCore.pyqtSignal(object)              # QMouseEvent (while pressed)
    released = QtCore.pyqtSignal(object)
    double_clicked = QtCore.pyqtSignal(object)
    context = QtCore.pyqtSignal(object, object)    # hit, global QPoint
    wheel = QtCore.pyqtSignal(object, object)      # hit, QWheelEvent (Ctrl held)
    resized = QtCore.pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMouseTracking(True)
        self.setMinimumSize(320, 160)
        self.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self._image = None
        self._fig_in = (1.0, 1.0)
        self._layout = None
        self._hover = None
        self._pressed = False
        self.hit_test = None       # callable(x_in, y_in) -> hit dict | None
        self.add_row_box = None    # inches box of the "+ channel" handle
        self.ruler_text = ""       # page-width ruler under the paper (screen only)
        self.sel_boxes = []        # inch boxes of the selected molecule (screen only)
        self.gizmo = None          # X/Y/Z orientation gizmo of the selected molecule (screen only)

    # geometry -----------------------------------------------------------
    def available_px(self):
        return (max(50, self.width() - 2 * _PAPER_MARGIN),
                max(50, self.height() - 2 * _PAPER_MARGIN))

    def set_image(self, image, fig_in, layout):
        self._image = image
        self._fig_in = fig_in
        self._layout = layout
        self.update()

    def _image_rect(self):
        if self._image is None:
            return QtCore.QRectF()
        dpr = self._image.devicePixelRatio() or 1.0
        w = self._image.width() / dpr
        h = self._image.height() / dpr
        return QtCore.QRectF((self.width() - w) / 2.0, (self.height() - h) / 2.0, w, h)

    def to_inches(self, pos):
        rect = self._image_rect()
        if rect.isEmpty():
            return None
        x = (pos.x() - rect.left()) / rect.width() * self._fig_in[0]
        y = (pos.y() - rect.top()) / rect.height() * self._fig_in[1]
        return x, y

    def inches_rect(self, box):
        rect = self._image_rect()
        if rect.isEmpty():
            return QtCore.QRectF()
        sx = rect.width() / self._fig_in[0]
        sy = rect.height() / self._fig_in[1]
        x0, y0, x1, y1 = box
        return QtCore.QRectF(rect.left() + x0 * sx, rect.top() + y0 * sy,
                             (x1 - x0) * sx, (y1 - y0) * sy)

    def close_rect(self, header_box):
        r = self.inches_rect(header_box)
        return QtCore.QRectF(r.right() - _CLOSE_SIZE - 2, r.top() + 3, _CLOSE_SIZE, _CLOSE_SIZE)

    def add_rect(self):
        if not self.add_row_box or self._image is None:
            return QtCore.QRectF()
        r = self.inches_rect(self.add_row_box)
        w, h = 64.0, 18.0
        return QtCore.QRectF(r.center().x() - w / 2, r.center().y() - h / 2, w, h)

    # painting -----------------------------------------------------------
    def paintEvent(self, _event):
        p = QtGui.QPainter(self)
        p.setRenderHint(QtGui.QPainter.Antialiasing, True)
        p.fillRect(self.rect(), self.palette().window())
        if self._image is None:
            p.setPen(self.palette().color(QtGui.QPalette.Mid))
            p.drawText(self.rect(), QtCore.Qt.AlignCenter,
                       "Nothing to show - add a channel (Format) or restore a frame (ⓘ)")
            return
        rect = self._image_rect()
        paper = rect.adjusted(-8, -8, 8, 8)
        p.setPen(QtCore.Qt.NoPen)
        p.setBrush(QtGui.QColor(0, 0, 0, 40))
        p.drawRoundedRect(paper.translated(0, 2), 3, 3)
        p.setBrush(QtGui.QColor("white"))
        p.drawRoundedRect(paper, 3, 3)
        p.drawImage(rect, self._image)
        accent = QtGui.QColor(30, 110, 220)
        if self.ruler_text:
            # Page width ruler: the paper is exactly the column width.
            y = paper.bottom() + 9
            pen = QtGui.QPen(self.palette().color(QtGui.QPalette.Mid), 1.0)
            p.setPen(pen)
            p.drawLine(QtCore.QPointF(rect.left(), y), QtCore.QPointF(rect.right(), y))
            for x in (rect.left(), rect.right()):
                p.drawLine(QtCore.QPointF(x, y - 4), QtCore.QPointF(x, y + 4))
            f = p.font()
            f.setPointSizeF(8.0)
            p.setFont(f)
            p.drawText(QtCore.QRectF(rect.left(), y + 2, rect.width(), 16),
                       QtCore.Qt.AlignHCenter | QtCore.Qt.AlignTop, self.ruler_text)
        hit = self._hover
        # "+ channel" handle: screen-only (never exported).
        ar = self.add_rect()
        if not ar.isEmpty():
            hot = bool(hit and hit.get("kind") == "add_row")
            p.setPen(QtGui.QPen(accent, 1.0, QtCore.Qt.SolidLine if hot else QtCore.Qt.DashLine))
            p.setBrush(QtGui.QColor(30, 110, 220, 40 if hot else 0))
            p.drawRoundedRect(ar, 6, 6)
            p.setPen(accent)
            f = p.font()
            f.setPointSizeF(8.0)
            p.setFont(f)
            p.drawText(ar, QtCore.Qt.AlignCenter, "+ channel")
        if self.sel_boxes:
            p.setPen(QtGui.QPen(accent, 1.0, QtCore.Qt.DashLine))
            p.setBrush(QtCore.Qt.NoBrush)
            for b in self.sel_boxes:
                p.drawRect(self.inches_rect(b))
        if self.gizmo:
            self._paint_gizmo(p)
        if not hit or not self._layout:
            return
        kind = hit.get("kind")
        if kind in ("frame", "close"):
            r = self.inches_rect(hit["box"])
            p.setPen(QtGui.QPen(accent, 1.4))
            p.setBrush(QtCore.Qt.NoBrush)
            if self._layout.get("transpose"):
                p.drawLine(QtCore.QPointF(r.right() - 2, r.top() + 4),
                           QtCore.QPointF(r.right() - 2, r.bottom() - 4))
            else:
                p.drawLine(QtCore.QPointF(r.left() + 6, r.bottom() - 3),
                           QtCore.QPointF(r.right() - 6, r.bottom() - 3))
            if hit.get("can_close"):
                cr = self.close_rect(hit["box"])
                p.setPen(QtCore.Qt.NoPen)
                p.setBrush(QtGui.QColor(200, 50, 50) if kind == "close" else QtGui.QColor(120, 120, 120, 200))
                p.drawEllipse(cr)
                p.setPen(QtGui.QPen(QtGui.QColor("white"), 1.6))
                m = 4.0
                p.drawLine(QtCore.QPointF(cr.left() + m, cr.top() + m),
                           QtCore.QPointF(cr.right() - m, cr.bottom() - m))
                p.drawLine(QtCore.QPointF(cr.right() - m, cr.top() + m),
                           QtCore.QPointF(cr.left() + m, cr.bottom() - m))
        elif kind == "channel":
            r = self.inches_rect(hit["box"])
            p.setPen(QtGui.QPen(accent, 1.2, QtCore.Qt.DashLine))
            p.setBrush(QtCore.Qt.NoBrush)
            p.drawRoundedRect(r.adjusted(3, 3, -3, -3), 4, 4)
        elif kind == "colorbar":
            r = self.inches_rect(hit["box"])
            p.setPen(QtGui.QPen(accent, 1.4))
            p.setBrush(QtCore.Qt.NoBrush)
            p.drawRect(r.adjusted(-2, -2, 2, 2))

    def gizmo_center(self):
        return QtCore.QPointF(_GIZMO_R + 10.0, _GIZMO_R + 10.0)

    def _paint_gizmo(self, p):
        """The preview's orientation gizmo: ring + X/Y/Z arrows (depth-
        sorted, far ones faded). Inner disc: tumble in 3-D; rim: spin."""
        c = self.gizmo_center()
        rad = _GIZMO_R
        dark = self.palette().color(QtGui.QPalette.Window).lightness() < 128
        ring = QtGui.QColor("#d7dde7" if dark else "#334155")
        p.setPen(QtGui.QPen(ring, 1.0))
        p.setBrush(QtGui.QColor(10, 15, 25, 110) if dark else QtGui.QColor(255, 255, 255, 200))
        p.drawEllipse(c, rad * 0.94, rad * 0.94)
        p.setPen(QtGui.QPen(ring, 0.8, QtCore.Qt.DotLine))
        p.setBrush(QtCore.Qt.NoBrush)
        p.drawEllipse(c, rad * 0.58, rad * 0.58)
        f = p.font()
        f.setBold(True)
        f.setPointSizeF(8.0)
        p.setFont(f)
        for _idx, depth, (vx, vy), color in self.gizmo["vectors"]:
            col = QtGui.QColor(color)
            col.setAlphaF(0.45 + 0.45 * ((depth + 1.0) * 0.5))
            end = QtCore.QPointF(c.x() + vx * rad, c.y() - vy * rad)
            p.setPen(QtGui.QPen(col, 2.2))
            p.drawLine(c, end)
            p.setBrush(col)
            p.drawEllipse(end, 2.4, 2.4)
            lab = QtCore.QPointF(c.x() + vx * rad * 1.16, c.y() - vy * rad * 1.16)
            p.drawText(QtCore.QRectF(lab.x() - 7, lab.y() - 7, 14, 14), QtCore.Qt.AlignCenter,
                       "XYZ"[_idx])
        p.setPen(ring)
        p.setBrush(ring)
        p.drawEllipse(c, 2.2, 2.2)

    def _gizmo_hit(self, pos):
        if not self.gizmo:
            return None
        c = self.gizmo_center()
        lx = (pos.x() - c.x()) / _GIZMO_R
        ly = -(pos.y() - c.y()) / _GIZMO_R
        mode = molecule_interaction.gizmo_hit_mode(lx, ly)
        if mode is None:
            return None
        return {"kind": "gizmo", "mode": mode, "local": (lx, ly)}

    # mouse --------------------------------------------------------------
    def _hit(self, pos):
        g = self._gizmo_hit(pos)
        if g is not None:
            return g
        ar = self.add_rect()
        if not ar.isEmpty() and ar.contains(QtCore.QPointF(pos)):
            return {"kind": "add_row"}
        if self.hit_test is None:
            return None
        inch = self.to_inches(pos)
        if inch is None:
            return None
        hit = self.hit_test(*inch)
        if hit and hit.get("kind") == "frame" and hit.get("can_close"):
            if self.close_rect(hit["box"]).contains(QtCore.QPointF(pos)):
                hit = dict(hit, kind="close")
        return hit

    def mouseMoveEvent(self, event):
        if self._pressed:
            self.moved.emit(event)
            return
        hit = self._hit(event.pos())
        if hit != self._hover:
            self._hover = hit
            self.update()
            self.hovered.emit(hit)

    def leaveEvent(self, _event):
        if not self._pressed and self._hover is not None:
            self._hover = None
            self.update()
            self.hovered.emit(None)

    def mousePressEvent(self, event):
        if event.button() == QtCore.Qt.RightButton:
            self.context.emit(self._hit(event.pos()), event.globalPos())
            return
        if event.button() not in (QtCore.Qt.LeftButton, QtCore.Qt.MiddleButton):
            return
        self._pressed = True
        self.pressed.emit(self._hit(event.pos()), event)

    def mouseReleaseEvent(self, event):
        if event.button() not in (QtCore.Qt.LeftButton, QtCore.Qt.MiddleButton):
            return
        self._pressed = False
        self.released.emit(event)

    def mouseDoubleClickEvent(self, event):
        if event.button() == QtCore.Qt.LeftButton:
            self.double_clicked.emit(self._hit(event.pos()))

    def wheelEvent(self, event):
        # Ctrl+wheel sizes things (text, bars, letters) like everywhere else
        # in the app; a plain wheel scrolls (actual-size view).
        if event.modifiers() & QtCore.Qt.ControlModifier:
            self.wheel.emit(self._hit(event.pos()), event)
            event.accept()
            return
        event.ignore()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.resized.emit()


def _tab_icon(kind, color):
    """Small line-art icon for a Format tab (drawn, so it follows the theme)."""
    size = 22
    pix = QtGui.QPixmap(size, size)
    pix.fill(QtCore.Qt.transparent)
    p = QtGui.QPainter(pix)
    p.setRenderHint(QtGui.QPainter.Antialiasing, True)
    pen = QtGui.QPen(color, 1.5)
    p.setPen(pen)
    p.setBrush(QtCore.Qt.NoBrush)
    if kind == "page":
        p.drawRect(QtCore.QRectF(5, 3, 12, 15))
        p.drawLine(QtCore.QPointF(3, 20), QtCore.QPointF(19, 20))
        p.drawLine(QtCore.QPointF(3, 18.5), QtCore.QPointF(3, 21.5))
        p.drawLine(QtCore.QPointF(19, 18.5), QtCore.QPointF(19, 21.5))
        for x in (7.5, 11.0, 14.5):
            p.drawRect(QtCore.QRectF(x - 1.2, 6, 2.4, 4))
    elif kind == "channels":
        for i, y in enumerate((4, 9.5, 15)):
            p.setBrush(QtGui.QColor(color.red(), color.green(), color.blue(), 70 + 60 * i))
            p.drawRoundedRect(QtCore.QRectF(3, y, 16, 4), 1.5, 1.5)
    elif kind == "titles":
        f = p.font()
        f.setBold(True)
        f.setPixelSize(15)
        p.setFont(f)
        p.drawText(QtCore.QRectF(0, 0, size, 17), QtCore.Qt.AlignCenter, "T")
        p.drawLine(QtCore.QPointF(4, 19), QtCore.QPointF(18, 19))
    elif kind == "colour":
        grad = QtGui.QLinearGradient(0, 19, 0, 3)
        for t, c in ((0.0, "#30123b"), (0.35, "#28bceb"), (0.65, "#a4fc3b"), (1.0, "#fb8022")):
            grad.setColorAt(t, QtGui.QColor(c))
        p.setBrush(QtGui.QBrush(grad))
        p.drawRoundedRect(QtCore.QRectF(7, 3, 8, 16), 2, 2)
    elif kind == "letters":
        p.drawRoundedRect(QtCore.QRectF(3, 3, 16, 16), 3, 3)
        f = p.font()
        f.setBold(True)
        f.setPixelSize(12)
        p.setFont(f)
        p.drawText(QtCore.QRectF(3, 2, 16, 16), QtCore.Qt.AlignCenter, "a")
    elif kind == "scale":
        p.setBrush(color)
        p.drawRect(QtCore.QRectF(3, 13, 16, 3))
        p.setBrush(QtCore.Qt.NoBrush)
        f = p.font()
        f.setPixelSize(8)
        p.setFont(f)
        p.drawText(QtCore.QRectF(0, 2, size, 10), QtCore.Qt.AlignCenter, "nm")
    elif kind == "molecule":
        pts = [QtCore.QPointF(5, 15), QtCore.QPointF(11, 6), QtCore.QPointF(17, 15)]
        p.drawLine(pts[0], pts[1])
        p.drawLine(pts[1], pts[2])
        p.drawLine(pts[0], pts[2])
        p.setBrush(QtGui.QColor("#5ab0ff"))
        for pt_ in pts:
            p.drawEllipse(pt_, 2.6, 2.6)
    elif kind == "zoom":
        p.drawRect(QtCore.QRectF(3, 3, 16, 16))
        p.setPen(QtGui.QPen(QtGui.QColor("#ff9f1c"), 1.6))
        p.drawRect(QtCore.QRectF(10, 5, 7, 7))
    p.end()
    return QtGui.QIcon(pix)


class _Segmented(QtWidgets.QWidget):
    """A row of exclusive buttons for a short choice (2-6 options), with
    the slice of QComboBox's API the Format cards use (addItem, findData,
    currentData, setCurrentIndex, currentIndexChanged), so ``_emit`` /
    ``sync`` treat it exactly like a drop-down. Every option stays visible
    and nothing is truncated."""

    currentIndexChanged = QtCore.pyqtSignal(int)

    _QSS = ("QToolButton { border: 1px solid palette(mid); "
            "padding: 2px 7px; background: palette(base); color: palette(text); }"
            "QToolButton:checked { background: palette(highlight); color: palette(highlighted-text); }"
            "QToolButton:disabled { color: palette(mid); }"
            "QToolButton:checked:disabled { background: palette(midlight); }")

    def __init__(self, items, tips=None, parent=None):
        super().__init__(parent)
        self._lay = QtWidgets.QHBoxLayout(self)
        self._lay.setContentsMargins(0, 0, 0, 0)
        self._lay.setSpacing(0)
        self._data, self._buttons = [], []
        self._index = -1
        self._tips = dict(tips or {})
        for key, text in items.items():
            self.addItem(text, key)
        self._lay.addStretch(1)

    def addItem(self, text, data=None):
        b = QtWidgets.QToolButton(self)
        b.setText(text)
        b.setCheckable(True)
        b.setAutoRaise(False)
        # Joined buttons: every one after the first drops its left border.
        b.setStyleSheet(self._QSS + ("" if not self._buttons else "QToolButton { border-left-width: 0px; }"))
        if data in self._tips:
            b.setToolTip(self._tips[data])
        i = len(self._buttons)
        b.clicked.connect(lambda _=False, i=i: self._clicked(i))
        self._lay.insertWidget(i, b)
        self._buttons.append(b)
        self._data.append(data)

    def _clicked(self, i):
        changed = i != self._index
        self.setCurrentIndex(i)
        if not changed:
            self._buttons[i].setChecked(True)   # exclusive: clicking the active one keeps it

    def count(self):
        return len(self._data)

    def itemText(self, i):
        return self._buttons[i].text()

    def findData(self, value):
        for i, d in enumerate(self._data):
            if d == value and type(d) is type(value):
                return i
        for i, d in enumerate(self._data):
            if d == value:
                return i
        return -1

    def currentIndex(self):
        return self._index

    def currentData(self):
        return self._data[self._index] if 0 <= self._index < len(self._data) else None

    def currentText(self):
        return self.itemText(self._index) if 0 <= self._index < len(self._data) else ""

    def setCurrentIndex(self, i):
        for j, b in enumerate(self._buttons):
            b.setChecked(j == i)
        if i != self._index:
            self._index = i
            self.currentIndexChanged.emit(i)


def _is_choice(widget):
    return isinstance(widget, (QtWidgets.QComboBox, _Segmented))


class _CardPopover(QtWidgets.QFrame):
    """One Format card, floating above its chip. A Qt popup: it closes on
    a click outside or Esc, and edits apply live underneath it."""

    closed = QtCore.pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent, QtCore.Qt.Popup)
        self.setObjectName("figureCardPopover")
        self.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.setStyleSheet("QFrame#figureCardPopover { border: 1px solid palette(mid); "
                           "border-radius: 6px; background: palette(window); }")
        lay = QtWidgets.QVBoxLayout(self)
        lay.setContentsMargins(10, 8, 10, 10)
        lay.setSpacing(6)

    def hideEvent(self, event):
        super().hideEvent(event)
        self.closed.emit()


# Chip labels: (kind, title, one-line description).
_FORMAT_CARDS = (
    ("page", "Page & layout", "Journal width, arrangement of frames and channels"),
    ("channels", "Channels", "Tick to show, drag to reorder"),
    ("titles", "Text", "Font, size, frame headers, channel names"),
    ("colour", "Colour", "Ranges and colour bars"),
    ("letters", "Panel letters", "a, b, c … on the images"),
    ("scale", "Scale bar", "Where, length and label"),
    ("zoom", "Zoom outlines", "Boxes of the zooms on the overview"),
    ("molecule", "Molecular model", "The same model on every panel"),
)

_CORNER_ARROWS = {"tl": "↖", "tr": "↗", "bl": "↙", "br": "↘"}
_CORNER_TIPS = {"tl": "Top left", "tr": "Top right", "bl": "Bottom left", "br": "Bottom right",
                "out_tl": "Outside the image, above its top-left corner",
                "out_tr": "Outside the image, above its top-right corner"}


class _FormatPanel(QtWidgets.QWidget):
    """The figure's options as one row of *chips* under the figure. Each
    chip is a category with its current setting in words ("Letters A, B ·
    ↖") and opens that category's card as a pop-over above it, so every
    option stays one click away without a strip of cramped, cropped cards
    (or tabs, which hid the state). Cards follow one widget grammar: on/off
    of a whole feature = a switch in the card header; 2-6 choices = a
    segmented row (``_Segmented``); long lists = a drop-down; numbers = a
    spin box with its unit. A dot on a chip / ↺ in its card = changed from
    the default. Emits ``changed(dict)`` with the style/state keys to update
    (same contract as the old card strip)."""

    changed = QtCore.pyqtSignal(dict)
    mol_action = QtCore.pyqtSignal(str)   # load | preview | props | remove

    CARDS = _FORMAT_CARDS

    def __init__(self, parent=None):
        super().__init__(parent)
        bar = QtWidgets.QHBoxLayout(self)
        bar.setContentsMargins(0, 2, 0, 2)
        bar.setSpacing(6)
        self._syncing = False
        self._cards = {}          # kind -> _CardPopover
        self._chips = {}          # kind -> QToolButton
        self._card_keys = {}      # kind -> [widget keys]
        self._defaults = {}
        self._style = {}
        self._mode = None
        self._last_scale = "all"
        self._closed_at = (None, 0.0)
        icon_color = self.palette().color(QtGui.QPalette.WindowText)
        self._icon_color = icon_color
        w = {}
        building = {"kind": None}

        for kind, title, blurb in self.CARDS:
            chip = QtWidgets.QToolButton(self)
            chip.setIcon(_tab_icon(kind, icon_color))
            chip.setIconSize(QtCore.QSize(18, 18))
            chip.setToolButtonStyle(QtCore.Qt.ToolButtonTextBesideIcon)
            chip.setCheckable(True)
            chip.setText(title)
            chip.setToolTip(f"<b>{title}</b><br>{blurb}")
            chip.setObjectName("figureChip")
            chip.setStyleSheet(
                "QToolButton#figureChip { border: 1px solid palette(midlight); border-radius: 12px; "
                "padding: 3px 10px 3px 6px; }"
                "QToolButton#figureChip:hover { border-color: palette(mid); }"
                "QToolButton#figureChip:checked { border-color: palette(highlight); "
                "background: palette(midlight); }"
                "QToolButton#figureChip[off=\"true\"] { color: palette(mid); }")
            chip.clicked.connect(lambda _=False, k=kind: self._chip_clicked(k))
            bar.addWidget(chip)
            self._chips[kind] = chip
        bar.addStretch(1)
        self.reset_btn = QtWidgets.QToolButton(self)
        self.reset_btn.setText("Reset all")
        self.reset_btn.setAutoRaise(True)
        self.reset_btn.setToolTip("Every format option back to its default (frames, channels and "
                                  "ranges are kept)")
        bar.addWidget(self.reset_btn)

        def card(kind, toggle=None):
            """A titled pop-over card; returns add(label, widget) placing
            label/field rows top to bottom. ``toggle``: the feature's on/off
            check box, shown as a switch in the header."""
            title, blurb = next((t, b) for k, t, b in self.CARDS if k == kind)
            building["kind"] = kind
            self._card_keys[kind] = []
            pop = _CardPopover(self)
            pop.closed.connect(lambda k=kind: self._card_closed(k))
            col = pop.layout()
            head = QtWidgets.QHBoxLayout()
            head.setSpacing(6)
            ic = QtWidgets.QLabel()
            ic.setPixmap(_tab_icon(kind, icon_color).pixmap(20, 20))
            head.addWidget(ic)
            head.addWidget(QtWidgets.QLabel(f"<b>{title}</b>"))
            head.addStretch(1)
            if toggle is not None:
                head.addWidget(toggle)
            rst = QtWidgets.QToolButton()
            rst.setText("↺")
            rst.setAutoRaise(True)
            rst.setToolTip(f"{title}: back to the defaults")
            rst.clicked.connect(lambda _=False, k=kind: self._reset_card(k))
            head.addWidget(rst)
            pop._reset_btn = rst
            col.addLayout(head)
            hint = QtWidgets.QLabel(blurb)
            hint.setStyleSheet("color: palette(mid);")
            col.addWidget(hint)
            line = QtWidgets.QFrame()
            line.setFrameShape(QtWidgets.QFrame.HLine)
            line.setStyleSheet("color: palette(midlight);")
            col.addWidget(line)
            grid = QtWidgets.QGridLayout()
            grid.setHorizontalSpacing(10)
            grid.setVerticalSpacing(5)
            grid.setColumnStretch(1, 1)
            col.addLayout(grid)
            self._cards[kind] = pop
            state = {"n": 0}

            def add(label, widget):
                r = state["n"]
                state["n"] += 1
                if label is None:
                    if isinstance(widget, QtWidgets.QLayout):
                        grid.addLayout(widget, r, 0, 1, 2)
                    else:
                        grid.addWidget(widget, r, 0, 1, 2)
                    return
                lab = QtWidgets.QLabel(label)
                grid.addWidget(lab, r, 0, QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
                if isinstance(widget, QtWidgets.QLayout):
                    grid.addLayout(widget, r, 1)
                else:
                    grid.addWidget(widget, r, 1)

            def section(text_or_widget):
                r = state["n"]
                state["n"] += 1
                if isinstance(text_or_widget, str):
                    text_or_widget = QtWidgets.QLabel(f"<b>{text_or_widget}</b>")
                grid.addWidget(text_or_widget, r, 0, 1, 2)

            add.section = section
            return add

        def reg(key, widget):
            w[key] = widget
            if building["kind"]:
                self._card_keys[building["kind"]].append(key)
            return widget

        def fit(c):
            c.setSizeAdjustPolicy(QtWidgets.QComboBox.AdjustToContents)
            c.setMaxVisibleItems(18)
            return c

        def combo(key, items):
            c = QtWidgets.QComboBox()
            for k, text in items.items():
                c.addItem(text, k)
            return reg(key, fit(c))

        def seg(key, items, tips=None):
            return reg(key, _Segmented(items, tips))

        def spin(key, lo, hi, step, suffix):
            s = QtWidgets.QDoubleSpinBox()
            s.setRange(lo, hi)
            s.setSingleStep(step)
            s.setSuffix(suffix)
            s.setMinimumWidth(86)
            return reg(key, s)

        def check(key, text, tip=None):
            b = QtWidgets.QCheckBox(text)
            if tip:
                b.setToolTip(tip)
            return reg(key, b)

        def switch(key, tip):
            b = QtWidgets.QCheckBox("On")
            b.setToolTip(tip)
            f = b.font()
            f.setBold(True)
            b.setFont(f)
            w[key] = b
            return b

        def hrow(*widgets):
            lay = QtWidgets.QHBoxLayout()
            lay.setSpacing(12)
            for x in widgets:
                lay.addWidget(x)
            lay.addStretch(1)
            return lay

        # -- Page & layout
        add = card("page")
        combo("page", hs.PAGE_NAMES).setToolTip(
            "Journal column width: sets the figure width and every text size")
        add("Page", w["page"])
        spin("page_mm", 30.0, 400.0, 1.0, " mm")
        add("Custom width", w["page_mm"])
        spin("panel_w", 0.6, 4.0, 0.1, " in").setToolTip(
            "Free size only - a page width sets it (Ctrl+Shift+wheel)")
        add("Panel width", w["panel_w"])
        add.section("Arrangement")
        seg("transpose", {False: "Frames across", True: "Frames down"},
            {False: "Frames left to right, one channel per row",
             True: "Frames top to bottom, one channel per column"})
        add("Frames", w["transpose"])
        combo("rotate_tall", _ROTATE_TALL).setToolTip(
            "Tall (portrait) images turned a quarter turn for this figure; "
            "scale bars and zoom outlines follow")
        add("Tall images", w["rotate_tall"])
        add(None, hrow(check("reverse", "Reverse order", "Reverse the frame order"),
                       check("_actual", "Actual size (100 %)",
                             "Draw the figure at its printed size on this screen")))
        self._card_keys["page"].remove("_actual")

        # -- Channels
        add = card("channels")
        lst = reg("channels", QtWidgets.QListWidget())
        lst.setDragDropMode(QtWidgets.QAbstractItemView.InternalMove)
        lst.setMinimumSize(260, 150)
        lst.setToolTip("Tick the channels to show; drag to reorder")
        add(None, lst)

        # -- Text
        add = card("titles")
        ff = reg("font_family", QtWidgets.QComboBox())
        for name in hs.available_fonts():
            ff.addItem(name, name)
        ff.setToolTip("Font of every text in the figure (journal-friendly fonts first)")
        fit(ff)
        add("Font", ff)
        spin("font", 5.0, 14.0, 0.5, " pt").setToolTip("Text size (Ctrl+wheel on the figure)")
        add("Size", w["font"])
        add.section("Frame headers")
        combo("header", _HEADER_NAMES)
        add("Header", w["header"])
        add(None, hrow(check("header_name", "+ file name", "Add the file name under the header"),
                       check("ref_bold", "Bold Δz = 0", "Bold header for the Δz = 0 reference frame")))
        add.section("Channels")
        add(None, check("channel_labels", "Channel names next to the rows",
                        "I, Δf ... left of the rows (above them when frames run down)"))

        # -- Colour
        add = card("colour")
        seg("range_mode", {"row": "Shared per channel", "image": "Each image its own"},
            {"row": "One colour range per channel, the same on every frame",
             "image": "Every image scaled to its own range"})
        add("Range", w["range_mode"])
        bars = check("colorbars", "Colour bars")
        f = bars.font()
        f.setBold(True)
        bars.setFont(f)
        add.section(bars)
        seg("cbar_scope", {"auto": "Follow the range", "row": "One per channel"},
            {"row": "One bar per channel even when each image has its own range - "
                    "its ends then read min and max"})
        add("Bars", w["cbar_scope"])
        combo("cbar_style", _CBAR_STYLES)
        add("Values", w["cbar_style"])
        seg("cbar_ticks", {0: "None", 2: "2", 3: "3", 5: "5"},
            {0: "A plain gradient, no tick marks or values", 2: "The two ends"})
        add("Ticks", w["cbar_ticks"])
        seg("cbar_text", {"rotated": "Rotated", "horizontal": "Horizontal"})
        add("Text", w["cbar_text"])
        seg("cbar_label_pos", {"inside": "Inside the bar", "outside": "Outside"})
        add("Name", hrow(check("cbar_label", "Channel and unit"), w["cbar_label_pos"]))
        spin("cbar_w", 0.5, 8.0, 0.25, " mm")
        add("Width", hrow(w["cbar_w"], check("cbar_outline", "Frame",
                                             "Thin frame around each colour bar")))

        # -- Panel letters
        add = card("letters", toggle=switch("letters", "Show panel letters"))
        self._card_keys["letters"].append("letters")
        seg("letter_style", {k: k for k in _LETTER_STYLE_NAMES}, _LETTER_STYLE_NAMES)
        add("Style", w["letter_style"])
        seg("letter_scope", {"frame": "Each frame", "panel": "Every panel"},
            {"frame": "One letter per frame (column)", "panel": "A letter on every image"})
        add("Applies to", w["letter_scope"])
        seg("letter_pos", {**_CORNER_ARROWS, "out_tl": "Above ↖", "out_tr": "Above ↗"}, _CORNER_TIPS)
        add("Position", w["letter_pos"])
        seg("letter_color", {"auto": "Auto", "white": "White", "black": "Black"},
            {"auto": "White or black per letter, from the image under it. "
                     "Right-click a letter to set just that one"})
        add("Colour", w["letter_color"])
        spin("letter_size", 5.0, 24.0, 0.5, " pt")
        add("Size", hrow(w["letter_size"], check("letter_bold", "Bold"),
                         check("letter_outline", "Outline")))

        # -- Scale bar
        add = card("scale", toggle=switch("_scale_on", "Show scale bars"))
        seg("scale_bar", {"all": "Every image", "one": "One image"})
        add("Applies to", w["scale_bar"])
        w["scale_auto"] = check("scale_auto", "Auto")
        spin("scale_len_nm", 0.01, 100000.0, 0.5, " nm").setDecimals(2)
        add("Length", hrow(w["scale_auto"], w["scale_len_nm"]))
        seg("scale_label", {"first": "Once per length", "all": "Every bar", "none": "None"},
            {"first": "Label the first bar of each length only"})
        add("Label", w["scale_label"])
        spin("scale_font", 4.5, 16.0, 0.5, " pt").setSpecialValueText("Text size")
        w["scale_font"].setToolTip("Size of the length label (Ctrl+wheel on a scale bar)")
        add("Label size", w["scale_font"])
        seg("scale_pos", _CORNER_ARROWS, _CORNER_TIPS)
        add("Position", w["scale_pos"])
        seg("scale_color", {"auto": "Auto", "white": "White", "black": "Black"},
            {"auto": "White or black per bar, from the image under it"})
        add("Colour", w["scale_color"])
        spin("scale_thick", 0.5, 5.0, 0.25, " %").setToolTip(
            "Bar thickness, % of the panel's shorter side")
        add("Thickness", hrow(w["scale_thick"], check("scale_outline", "Outline")))

        # -- Zoom outlines (zoom figures only)
        add = card("zoom")
        seg("zoom_color", {"color": "One colour per zoom", "white": "White", "black": "Black"})
        add("Colour", w["zoom_color"])
        spin("zoom_lw", 0.25, 4.0, 0.25, " pt")
        add("Line width", hrow(w["zoom_lw"], check("zoom_labels", "Letters",
                                                   "Letter next to each zoom outline")))

        # -- Molecular model (.pdb / .xyz / .mol), shared by all panels
        add = card("molecule", toggle=switch("molecules_on", "Show the molecular model(s)"))
        self._card_keys["molecule"].append("molecules_on")
        b_load = QtWidgets.QPushButton("Load file…")
        b_load.setToolTip("Overlay a .pdb / .xyz / .mol model on every panel")
        b_prev = QtWidgets.QPushButton("From preview")
        b_prev.setToolTip("Use the model(s) already placed on one of these frames in the preview")
        add(None, hrow(b_load, b_prev))
        seg("mol_scope", {"all": "Every panel", "first": "First channel row"})
        add("Applies to", w["mol_scope"])
        spin("mol_size", 0.2, 4.0, 0.1, " ×").setToolTip(
            "Atom and bond size relative to the automatic (panel-width) size")
        add("Atom size", w["mol_size"])
        combo("mol_palette", _mol_palette_choices()).setToolTip(
            "Atom colours (CPK, Jmol, PyMOL, Avogadro, ASE, VESTA);\n"
            "'As preview' follows the preview's palette")
        add("Palette", w["mol_palette"])
        b_props = QtWidgets.QPushButton("Appearance…")
        b_props.setToolTip("Style, rotation, scale, mirror (the preview's dialog)")
        b_rm = QtWidgets.QPushButton("Remove")
        add(None, hrow(b_props, b_rm))
        how = QtWidgets.QLabel("Drag: move  ·  Shift+drag: spin  ·  Ctrl+Shift or middle: 3-D\n"
                               "Gizmo (top-left)  ·  X / Y / Z: rotate  ·  arrows: nudge  ·  Del")
        how.setStyleSheet("color: palette(mid);")
        add(None, how)
        self._mol_buttons = (b_props, b_rm)
        for btn, act in ((b_load, "load"), (b_prev, "preview"), (b_props, "props"), (b_rm, "remove")):
            btn.clicked.connect(lambda _=False, a=act: self.mol_action.emit(a))
        building["kind"] = None
        self.w = w

        for key, widget in w.items():
            if _is_choice(widget):
                widget.currentIndexChanged.connect(lambda _i, k=key: self._emit(k))
            elif isinstance(widget, QtWidgets.QCheckBox):
                widget.toggled.connect(lambda _on, k=key: self._emit(k))
            elif isinstance(widget, QtWidgets.QDoubleSpinBox):
                widget.valueChanged.connect(lambda _v, k=key: self._emit(k))
        w["channels"].itemChanged.connect(lambda _i: self._emit("channels"))
        w["channels"].model().rowsMoved.connect(lambda *_a: self._emit("channels"))

    # ------------------------------------------------------------------
    # Chips and pop-over cards
    # ------------------------------------------------------------------
    def sizeHint(self):
        return QtCore.QSize(1200, 34)

    def set_mode(self, mode):
        self._mode = mode
        chip = self._chips.get("zoom")
        if chip is not None:
            chip.setVisible(mode == "zoom")

    def set_defaults(self, style):
        """The style the ↺ buttons go back to (and the chips' change dot
        compares against)."""
        self._defaults = dict(style or {})

    def _chip_clicked(self, kind):
        chip = self._chips[kind]
        last, t = self._closed_at
        # The press that closed this chip's own pop-over must not reopen it.
        if last == kind and (QtCore.QTime.currentTime().msecsSinceStartOfDay() - t) < 250:
            chip.setChecked(False)
            return
        self.open_card(kind)

    def open_card(self, kind, global_pos=None):
        """Show a card above its chip (or at ``global_pos``, e.g. next to a
        clicked letter on the figure)."""
        pop = self._cards.get(kind)
        if pop is None:
            return
        for k, c in self._chips.items():
            c.setChecked(k == kind)
        pop.adjustSize()
        size = pop.sizeHint()
        screen = QtWidgets.QApplication.screenAt(global_pos or self.mapToGlobal(QtCore.QPoint(0, 0)))
        avail = (screen or QtWidgets.QApplication.primaryScreen()).availableGeometry()
        if global_pos is None:
            chip = self._chips[kind]
            anchor = chip.mapToGlobal(QtCore.QPoint(0, 0))
            x, y = anchor.x(), anchor.y() - size.height() - 6
        else:
            x, y = global_pos.x() + 12, global_pos.y() - size.height() // 2
        x = max(avail.left() + 4, min(x, avail.right() - size.width() - 4))
        y = max(avail.top() + 4, min(y, avail.bottom() - size.height() - 4))
        pop.move(x, y)
        pop.show()
        pop.raise_()

    def _card_closed(self, kind):
        self._closed_at = (kind, QtCore.QTime.currentTime().msecsSinceStartOfDay())
        chip = self._chips.get(kind)
        if chip is not None:
            chip.setChecked(False)

    def _style_keys(self, kind):
        keys = [k for k in self._card_keys.get(kind, ()) if not k.startswith("_")
                and k not in ("channels", "scale_auto")]
        if kind == "scale":
            keys.append("scale_len_nm")
        return keys

    def _card_changed(self, kind):
        if kind == "channels":
            return False
        d = self._defaults
        if not d:
            return False
        return any(self._style.get(k) != d.get(k) for k in self._style_keys(kind) if k in d)

    def _reset_card(self, kind):
        if kind == "channels":
            self.changed.emit({"rows": []})
            return
        d = self._defaults
        out = {k: d.get(k) for k in self._style_keys(kind) if k in d}
        if out:
            self.changed.emit(out)

    def _summary(self, kind, st, spec, n_mols):
        """Chip text: the category's current state in a few words, and
        whether the feature is off."""
        def name(table, key, fallback=""):
            return str(table.get(st.get(key), fallback))

        if kind == "page":
            page = st.get("page") or "free"
            if page == "free":
                txt = "Free size"
            elif page == "custom":
                txt = f"{float(st.get('page_mm') or 0):.0f} mm"
            else:
                txt = hs.PAGE_NAMES.get(page, page).replace(" column", " col.")
            return f"{txt} · {'frames down' if st.get('transpose') else 'frames across'}", False
        if kind == "channels":
            rows = (spec or {}).get("rows") or []
            labels = [hs.channel_text(r.get("display") or r.get("label") or "").split(" (")[0]
                      for r in rows]
            if not labels:
                return "Channels", False
            short = ", ".join(labels[:2]) + ("…" if len(labels) > 2 else "")
            return (short if len(short) <= 34 else f"{len(labels)} channels"), False
        if kind == "titles":
            fam = st.get("font_family") or hs.default_font()
            return f"{fam} {float(st.get('font') or 0):g} pt", False
        if kind == "colour":
            rng = "Shared range" if st.get("range_mode") == "row" else "Own ranges"
            return f"{rng} · {'bars' if st.get('colorbars') else 'no bars'}", False
        if kind == "letters":
            if not st.get("letters"):
                return "Letters off", True
            pos = _CORNER_ARROWS.get(st.get("letter_pos"), "above")
            return f"Letters {name(_LETTER_STYLE_NAMES, 'letter_style')} · {pos}", False
        if kind == "scale":
            if st.get("scale_bar") == "off":
                return "Scale bar off", True
            length = st.get("scale_len_nm")
            ln = f"{float(length):g} nm" if length else "auto"
            return f"Scale bar {ln} · {_CORNER_ARROWS.get(st.get('scale_pos'), '')}", False
        if kind == "zoom":
            return f"Outlines {float(st.get('zoom_lw') or 1):g} pt", False
        if kind == "molecule":
            if not n_mols:
                return "Molecule: none", True
            if not st.get("molecules_on", True):
                return "Molecule hidden", True
            return f"{n_mols} model{'s' if n_mols != 1 else ''}", False
        return kind, False

    def _update_chips(self, spec, n_mols):
        for kind, chip in self._chips.items():
            text, off = self._summary(kind, self._style, spec, n_mols)
            dot = "  •" if self._card_changed(kind) else ""
            chip.setText(text + dot)
            if chip.property("off") != off:
                chip.setProperty("off", off)
                chip.style().unpolish(chip)
                chip.style().polish(chip)
            pop = self._cards.get(kind)
            if pop is not None:
                pop._reset_btn.setEnabled(kind == "channels" or self._card_changed(kind))

    # Widget value <-> style value for keys shown in friendlier units.
    _TO_WIDGET = {"scale_thick": lambda v: float(v) * 100.0, "cbar_w": lambda v: float(v) * 25.4}
    _FROM_WIDGET = {"scale_thick": lambda v: float(v) / 100.0, "cbar_w": lambda v: float(v) / 25.4}

    def _emit(self, key):
        if self._syncing:
            return
        w = self.w
        if key == "channels":
            rows = []
            for i in range(w["channels"].count()):
                it = w["channels"].item(i)
                if it.checkState() == QtCore.Qt.Checked:
                    rows.append(it.data(QtCore.Qt.UserRole))
            self.changed.emit({"rows": rows})
            return
        if key in ("scale_auto", "scale_len_nm"):
            auto = w["scale_auto"].isChecked()
            w["scale_len_nm"].setEnabled(not auto)
            self.changed.emit({"scale_len_nm": None if auto else float(w["scale_len_nm"].value())})
            return
        if key in ("_scale_on", "scale_bar"):
            on = w["_scale_on"].isChecked()
            where = w["scale_bar"].currentData() or self._last_scale or "all"
            self.changed.emit({"scale_bar": where if on else "off"})
            return
        widget = w[key]
        if _is_choice(widget):
            value = widget.currentData()
        elif isinstance(widget, QtWidgets.QCheckBox):
            value = widget.isChecked()
        else:
            value = float(widget.value())
        if key in self._FROM_WIDGET:
            value = self._FROM_WIDGET[key](value)
        if key == "scale_font" and value <= widget.minimum() + 1e-9:
            value = None   # "Text size"
        if key == "page" and value in hs.PAGE_PRESETS:
            # A journal page sets the text sizes too - no hand-tuning.
            self.changed.emit(dict(hs.page_text_style(value), page=value))
            return
        self.changed.emit({key: value})

    def sync(self, style, spec, auto_scale_nm=None, actual=False, n_molecules=0):
        self._syncing = True
        self._style = dict(style or {})
        try:
            w = self.w
            w["_actual"].setChecked(bool(actual))
            for key, widget in w.items():
                if key in ("channels", "scale_auto", "scale_len_nm", "_actual", "_scale_on",
                           "scale_bar"):
                    continue
                value = style.get(key)
                if key == "font_family" and not value:
                    value = hs.default_font()
                if key in self._TO_WIDGET and value is not None:
                    value = self._TO_WIDGET[key](value)
                if _is_choice(widget):
                    idx = widget.findData(value)
                    if idx >= 0:
                        widget.setCurrentIndex(idx)
                elif isinstance(widget, QtWidgets.QCheckBox):
                    widget.setChecked(bool(value))
                elif isinstance(widget, QtWidgets.QDoubleSpinBox):
                    if value is not None:
                        widget.setValue(float(value))
                    elif key == "scale_font":
                        widget.setValue(widget.minimum())
            where = style.get("scale_bar") or "all"
            scale_on = where != "off"
            if scale_on:
                self._last_scale = where
            w["_scale_on"].setChecked(scale_on)
            idx = w["scale_bar"].findData(self._last_scale)
            w["scale_bar"].setCurrentIndex(max(0, idx))
            auto = style.get("scale_len_nm") is None
            w["scale_auto"].setChecked(auto)
            length = style.get("scale_len_nm") or auto_scale_nm
            if length:
                w["scale_len_nm"].setValue(float(length))
            lst = w["channels"]
            lst.clear()
            shown = [r["label"] for r in (spec or {}).get("rows", [])]
            others = [c["label"] for c in (spec or {}).get("channels", []) if c["label"] not in shown]
            for label in shown + others:
                it = QtWidgets.QListWidgetItem(hs.channel_text(label))
                it.setData(QtCore.Qt.UserRole, label)
                it.setFlags(it.flags() | QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsDragEnabled)
                it.setCheckState(QtCore.Qt.Checked if label in shown else QtCore.Qt.Unchecked)
                lst.addItem(it)
            letters_on = bool(style.get("letters"))
            for key in ("letter_style", "letter_scope", "letter_pos", "letter_color",
                        "letter_size", "letter_bold", "letter_outline"):
                w[key].setEnabled(letters_on)
            for key in ("scale_bar", "scale_auto", "scale_pos", "scale_color", "scale_label",
                        "scale_font", "scale_thick", "scale_outline"):
                w[key].setEnabled(scale_on)
            w["scale_len_nm"].setEnabled(scale_on and not auto)
            bars_on = bool(style.get("colorbars"))
            for key in ("cbar_scope", "cbar_ticks", "cbar_label", "cbar_style", "cbar_text",
                        "cbar_w", "cbar_label_pos", "cbar_outline"):
                w[key].setEnabled(bars_on)
            w["page_mm"].setEnabled(style.get("page") == "custom")
            w["panel_w"].setEnabled((style.get("page") or "free") == "free")
            mol_on = bool(n_molecules) and bool(style.get("molecules_on", True))
            for key in ("mol_scope", "mol_size", "mol_palette"):
                w[key].setEnabled(mol_on)
            for b in self._mol_buttons:
                b.setEnabled(bool(n_molecules))
            self._update_chips(spec, n_molecules)
        finally:
            self._syncing = False


class FigureWindow(QtWidgets.QWidget):
    """A live, data-aware figure (see module docstring)."""

    def __init__(self, source, fid, keys, *, mode="height", excluded=None, root_key=None, parent=None):
        super().__init__(parent, QtCore.Qt.Window)
        self.setAttribute(QtCore.Qt.WA_DeleteOnClose, True)
        self.setFocusPolicy(QtCore.Qt.StrongFocus)   # X/Y/Z, arrows for a selected molecule
        self.source = source
        self.fid = fid
        self.mode = mode
        self.root_key = str(root_key) if root_key else None
        self._keys = [str(k) for k in keys]
        self._last_wheel = 0.0
        # Always open fitted to the window; "actual size" is a per-window view.
        self._actual = False
        self._excluded = list(excluded or [])
        self._aborted = []          # [{key, frac}] frames whose scan stopped early
        self.state = self._initial_state(source.get_state(fid))
        self._undo, self._redo = [], []
        self._panels = {}           # (key, idx) -> panel dict
        self._members = None
        self._data_sig = None
        self.spec = None
        self.layout_info = None
        self._drag = None
        self._fig = None
        self._pending = {}
        self._mol_sel = None          # index of the selected molecule
        self._mol_drag = None
        self._mol_axes = []           # screen-figure axes that carry molecules
        self._mol_artists = []
        self._mol_cache = (None, [])  # (state repr, [Molecule])

        self._build_ui()
        self._init_shortcuts()
        self.resize(1720, 900)
        self._render_timer = QtCore.QTimer(self, singleShot=True, interval=90)
        self._render_timer.timeout.connect(self._render_screen)
        self._drag_timer = QtCore.QTimer(self, singleShot=True, interval=_DRAG_THROTTLE_MS)
        self._drag_timer.timeout.connect(self._redraw_fast)
        self._format_timer = QtCore.QTimer(self, singleShot=True, interval=160)
        self._format_timer.timeout.connect(self._flush_format)
        self._live_timer = QtCore.QTimer(self, interval=_LIVE_POLL_MS)
        self._live_timer.timeout.connect(self._check_live)
        self._live_timer.start()
        self._rebuild(reload_data=True)

    # ------------------------------------------------------------------
    # State
    # ------------------------------------------------------------------
    def _initial_state(self, saved):
        saved_style = saved.get("style") or {}
        style = dict(self.source.default_style(self.mode))
        style.update(saved_style)
        # Figures saved before the style dict stored these at the top level.
        if "range_mode" in saved and "range_mode" not in saved_style:
            style["range_mode"] = saved["range_mode"]
        if isinstance(saved.get("letters"), bool) and "letters" not in saved_style:
            style["letters"] = saved["letters"]
        if isinstance(saved.get("scale_bar"), bool) and "scale_bar" not in saved_style:
            style["scale_bar"] = "one" if saved["scale_bar"] else "off"
        return {
            "reference_key": saved.get("reference_key") or self.source.default_reference(self._keys),
            "dropped": list(saved.get("dropped") or []),
            "forced": list(saved.get("forced") or []),
            "rows": list(saved["rows"]) if saved.get("rows") else None,
            "row_choices": dict(saved.get("row_choices") or {}),
            "row_cmaps": dict(saved.get("row_cmaps") or {}),
            "letter_colors": dict(saved.get("letter_colors") or {}),
            "molecules": [dict(m) for m in (saved.get("molecules") or []) if isinstance(m, dict)],
            "show_repeats": bool(saved.get("show_repeats", False)),
            "clims": {k: v for k, v in (saved.get("clims") or {}).items() if isinstance(v, (list, tuple))},
            "style": hs.resolve_style(style),
        }

    @property
    def style(self):
        return self.state["style"]

    def _snapshot(self):
        return copy.deepcopy(self.state)

    def _apply_edit(self, changes=None, *, style=None, push=True, reload_data=False):
        new = self._snapshot()
        new.update(changes or {})
        if style:
            new["style"] = hs.resolve_style(dict(new["style"], **style))
        if new == self.state:
            return
        if push:
            self._undo.append(self._snapshot())
            del self._undo[:-100]
            self._redo.clear()
        self.state = new
        if style:
            self.source.set_default_style(self.state["style"], self.mode)
        self._commit(reload_data=reload_data)

    def _commit(self, reload_data=False):
        self.source.set_state(self.fid, self.state)
        self._rebuild(reload_data=reload_data)

    def undo(self):
        if not self._undo:
            return
        self._redo.append(self._snapshot())
        self.state = self._undo.pop()
        self._commit()
        self._flash("Undone")

    def redo(self):
        if not self._redo:
            return
        self._undo.append(self._snapshot())
        self.state = self._redo.pop()
        self._commit()

    def reset_figure(self):
        self._apply_edit({
            "reference_key": self.source.default_reference(self._keys),
            "dropped": [], "forced": [], "rows": None, "row_choices": {},
            "row_cmaps": {}, "clims": {}, "show_repeats": False, "letter_colors": {},
        })

    def reset_format(self):
        self._undo.append(self._snapshot())
        self._redo.clear()
        self.state = dict(self._snapshot(), style=hs.resolve_style(hs.RECIPE_DEFAULTS.get(self.mode, {})))
        self.source.set_default_style({}, self.mode)
        self._commit()
        self._flash("Format reset")

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------
    def _build_ui(self):
        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(8, 6, 8, 8)
        root.setSpacing(6)
        strip = QtWidgets.QHBoxLayout()
        strip.setSpacing(6)

        def button(text, tip, slot, primary=False, checkable=False):
            b = QtWidgets.QToolButton(self)
            b.setText(text)
            b.setToolTip(tip)
            b.setToolButtonStyle(QtCore.Qt.ToolButtonTextOnly)
            b.setMinimumHeight(26)
            b.setCheckable(checkable)
            if primary:
                f = b.font()
                f.setBold(True)
                b.setFont(f)
            (b.toggled if checkable else b.clicked).connect(slot)
            strip.addWidget(b)
            return b

        def split(text, tip, slot, menu_builder, primary=False):
            """Click = the usual action, arrow = its options."""
            b = button(text, tip, slot, primary=primary)
            b.setPopupMode(QtWidgets.QToolButton.MenuButtonPopup)
            m = QtWidgets.QMenu(b)
            m.aboutToShow.connect(lambda m=m: (m.clear(), menu_builder(m)))
            b.setMenu(m)
            return b

        self.copy_btn = split("Copy", "Copy the figure (Ctrl+C) - PNG and SVG on the clipboard; "
                              "paste into PowerPoint, Word, Inkscape, Obsidian ...",
                              lambda: self.copy_to_clipboard(), self._copy_menu, primary=True)
        self.ppt_btn = split("PowerPoint", "Send the figure to the open PowerPoint presentation "
                             "(arrow: current slide / as picture)", lambda: self.send_to_powerpoint(),
                             self._ppt_menu)
        self.obsidian_btn = button("Obsidian", "Add the figure to an Obsidian note", lambda: None)
        self.obsidian_btn.setPopupMode(QtWidgets.QToolButton.InstantPopup)
        om = QtWidgets.QMenu(self.obsidian_btn)
        om.aboutToShow.connect(lambda m=om: (m.clear(), self.source.obsidian_menu(
            m, self._obsidian_render, parent=self)))
        self.obsidian_btn.setMenu(om)
        self.inkscape_btn = split("Inkscape", "Open the figure in Inkscape as a layered SVG "
                                  "(images, letters, scale bars, colour bars, labels on their own layers)",
                                  lambda: self.open_in_inkscape(), self._inkscape_menu)
        self.save_btn = split("Save…", "Save as PDF, SVG, PNG, TIFF or EPS (Ctrl+S)",
                              lambda: self.save_as(), self._save_menu)
        strip.addSpacing(12)
        self.format_btn = button("Format", "Channels, letters, scale bar, colour bars, layout "
                                 "(also on right-click)", self._toggle_format, checkable=True)
        strip.addStretch(1)

        self.note_btn = QtWidgets.QToolButton(self)
        self.note_btn.setAutoRaise(True)
        self.note_btn.setPopupMode(QtWidgets.QToolButton.InstantPopup)
        self.note_btn.setVisible(False)
        strip.addWidget(self.note_btn)

        self.status = QtWidgets.QLabel("", self)
        self.status.setStyleSheet("color: palette(mid);")
        strip.addWidget(self.status)

        self.more_btn = QtWidgets.QToolButton(self)
        self.more_btn.setText("⋯")
        self.more_btn.setAutoRaise(True)
        self.more_btn.setPopupMode(QtWidgets.QToolButton.InstantPopup)
        more = QtWidgets.QMenu(self.more_btn)
        more.addAction("Open in canvas (free layout)", self.open_in_canvas)
        more.addAction("Reset figure (frames, channels, ranges)", self.reset_figure)
        more.addAction("Reset format", self.reset_format)
        self.more_btn.setMenu(more)
        strip.addWidget(self.more_btn)
        root.addLayout(strip)

        self.scroll = QtWidgets.QScrollArea(self)
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.view = _FigureView(self)
        self.view.hit_test = self._hit_test
        self.view.hovered.connect(self._on_hover)
        self.view.pressed.connect(self._on_press)
        self.view.moved.connect(self._on_drag_move)
        self.view.released.connect(self._on_release)
        self.view.double_clicked.connect(self._on_double_click)
        self.view.context.connect(self._on_context)
        self.view.wheel.connect(self._on_wheel)
        self.view.resized.connect(lambda: self._render_timer.start() if not self._actual else None)
        self.scroll.setWidget(self.view)
        root.addWidget(self.scroll, 1)
        self.format_panel = _FormatPanel(self)
        self.format_panel.changed.connect(self._on_format_changed)
        self.format_panel.reset_btn.clicked.connect(self.reset_format)
        self.format_panel.mol_action.connect(self._on_mol_action)
        self.format_panel.set_mode(self.mode)
        self.format_panel.set_defaults(hs.resolve_style(hs.RECIPE_DEFAULTS.get(self.mode, {})))
        root.addWidget(self.format_panel)
        open_fmt = bool(self.source.export_settings().get("format_open", True))
        self.format_panel.setVisible(open_fmt)
        self.format_btn.blockSignals(True)
        self.format_btn.setChecked(open_fmt)
        self.format_btn.blockSignals(False)

        self.hint = QtWidgets.QLabel(self)
        self.hint.setStyleSheet("color: palette(mid);")
        self.hint.setWordWrap(True)
        self.hint.setText(_MODE_HINTS.get(self.mode, "") + "hover a header, ×: drop  ·  click a channel "
                          "label: switch / remove  ·  drag a colourbar: contrast, click it: exact values  ·  "
                          "right-click anything: its options  ·  Ctrl+Z undo")
        root.addWidget(self.hint)

    def _init_shortcuts(self):
        for seq, slot in (
            (QtGui.QKeySequence.Copy, self.copy_to_clipboard),
            (QtGui.QKeySequence.Save, self.save_as),
            (QtGui.QKeySequence.Undo, self.undo),
            (QtGui.QKeySequence.Redo, self.redo),
            (QtGui.QKeySequence("Ctrl+Shift+Z"), self.redo),
        ):
            sc = QtWidgets.QShortcut(seq, self)
            sc.setContext(QtCore.Qt.WindowShortcut)
            sc.activated.connect(slot)

    def _toggle_format(self, on):
        self.format_panel.setVisible(bool(on))
        self.source.set_export_settings(format_open=bool(on))
        if on:
            self._sync_format_panel()
        self._render_timer.start()

    def _sync_format_panel(self):
        # isHidden, not isVisible: the first sync runs from __init__, before
        # the window is shown, when isVisible() is False for every child -
        # the panel then stayed blank (no channels, "Free size" on a
        # Double-column figure) until the first edit.
        if self.format_panel.isHidden() or not self.spec:
            return
        auto_len = None
        widths = [p.get("width_nm") for p in self._panel_map().values() if p.get("width_nm")]
        if widths:
            auto_len = hs.nice_length(0.3 * widths[0])
        self.format_panel.sync(self.style, self.spec, auto_len, actual=self._actual,
                               n_molecules=len(self._molecules()))

    def _on_format_changed(self, change):
        # Spin boxes fire per step - coalesce into one edit.
        self._pending.update(change)
        self._format_timer.start()

    def _flush_format(self):
        change, self._pending = self._pending, {}
        if not change:
            return
        if "_actual" in change:
            self._set_actual(bool(change.pop("_actual")))
        rows = change.pop("rows", None)
        if rows is not None:
            self._apply_edit({"rows": rows or None})
        if change:
            self._apply_edit(style=change)

    # ------------------------------------------------------------------
    # Data + spec
    # ------------------------------------------------------------------
    def _all_keys(self):
        forced = [k for k in self.state["forced"] if k not in self._keys]
        return self._keys + forced

    def _rebuild(self, reload_data=False):
        keys = self._all_keys()
        if reload_data or self._members is None:
            self._members = self.source.member_inputs(keys)
            self._data_sig = self.source.data_signature(keys)
            self._panels.clear()
            self._aborted = self._find_aborted(self._members)
        st = self.style
        auto_out = [a["key"] for a in self._aborted if a["key"] not in self.state["forced"]]
        self.spec = hs.build_spec(
            self._members, mode=self.mode,
            reference_key=self.state["reference_key"],
            dropped=list(self.state["dropped"]) + auto_out,
            rows=self.state["rows"],
            root_key=self.root_key,
            row_choices=self.state["row_choices"],
            header=st["header"], header_name=st["header_name"], reverse=st["reverse"],
            show_repeats=self.state["show_repeats"],
            row_cmaps=self.state["row_cmaps"])
        self._update_title()
        self._render_screen()
        self._update_note()
        self._sync_format_panel()

    def _find_aborted(self, members):
        """Frames with under a quarter of their rows scanned are no column
        of a series figure (they'd be a white panel with a strip of data).
        A user selection keeps everything it was given."""
        if self.mode == "selection":
            return []
        out = []
        for m in members:
            picks = m.get("picks") or []
            if not picks:
                continue
            frac = self.source.valid_fraction(m["key"], picks[0][0])
            if frac < 0.25:
                out.append({"key": str(m["key"]), "frac": float(frac)})
        return out

    def _panel_map(self):
        panels = {}
        for r, row in enumerate(self.spec["rows"]):
            for key, idx in row["cells"].items():
                ck = (key, int(idx))
                if ck not in self._panels:
                    try:
                        self._panels[ck] = self.source.panel(key, int(idx))
                    except Exception as exc:
                        log_status(f"[Figure] panel {Path(key).name}#{idx} failed: {exc}")
                        self._panels[ck] = None
                if self._panels[ck] is not None:
                    panels[(r, key)] = self._panels[ck]
        return panels

    def _update_title(self):
        cols = self.spec["columns"] if self.spec else []
        what = _MODE_TITLES.get(self.mode, "Figure")
        names = [c["name"] for c in cols]
        span = f"{names[0]} … {names[-1]}" if len(names) > 1 else (names[0] if names else "")
        self.setWindowTitle(f"{what} - {span} ({len(cols)} frame{'s' if len(cols) != 1 else ''})")

    def _update_note(self):
        left_out = [e for e in self._excluded if e["key"] not in self.state["forced"]
                    and e.get("reason") != "aborted"]
        dropped = [k for k in self.state["dropped"] if k in self._all_keys()]
        aborted = [a for a in self._aborted if a["key"] not in self.state["forced"]]
        aborted += [e for e in self._excluded if e.get("reason") == "aborted"
                    and e["key"] not in self.state["forced"]
                    and e["key"] not in {a["key"] for a in aborted}]
        menu = QtWidgets.QMenu(self.note_btn)
        parts = []
        if aborted:
            parts.append(f"{len(aborted)} left out (scan stopped early)")
            menu.addSection("Scan stopped early - mostly empty frame")
            for a in aborted:
                act = menu.addAction(f"Include {Path(a['key']).stem} anyway  ({a['frac'] * 100:.0f} % scanned)")
                act.triggered.connect(lambda _=False, k=a["key"]: self._apply_edit(
                    {"forced": self.state["forced"] + [k]}))
        if left_out:
            parts.append(f"{len(left_out)} left out (Z not flat)")
            menu.addSection("Z changes within the frame - no single Δz")
            for e in left_out:
                rng = e.get("z_range_nm")
                extra = f"  (Z range {rng * 1e3:.0f} pm)" if rng else ""
                act = menu.addAction(f"Include {Path(e['key']).stem} anyway{extra}")
                act.triggered.connect(lambda _=False, k=e["key"]: self._apply_edit(
                    {"forced": self.state["forced"] + [k]}, reload_data=True))
        if dropped:
            parts.append(f"{len(dropped)} removed")
            menu.addSection("Removed frames")
            for k in dropped:
                act = menu.addAction(f"Restore {Path(k).stem}")
                act.triggered.connect(lambda _=False, key=k: self._apply_edit(
                    {"dropped": [d for d in self.state["dropped"] if d != key]}))
        own = sorted({k for _r, k in (self.layout_info or {}).get("own_scale", {})},
                     key=lambda k: self._all_keys().index(k) if k in self._all_keys() else 0)
        if own:
            parts.append(f"{len(own)} on own scale")
            menu.addSection("Different baseline (or no signal) - not on the shared range")
            for k in own:
                act = menu.addAction(f"{Path(k).stem}: show in viewer")
                act.triggered.connect(lambda _=False, key=k: self._show_key(key))
        repeats = list((self.spec or {}).get("hidden_repeats") or [])
        if repeats:
            parts.append(f"{len(repeats)} repeat{'s' if len(repeats) != 1 else ''} hidden")
            menu.addSection("Re-scans under identical conditions - the latest is shown")
            menu.addAction("Show the repeats too").triggered.connect(
                lambda: self._apply_edit({"show_repeats": True}))
        elif self.state["show_repeats"] and self.mode != "selection":
            parts.append("repeats shown")
            menu.addSection("Repeats")
            menu.addAction("Hide repeats (keep the latest)").triggered.connect(
                lambda: self._apply_edit({"show_repeats": False}))
        forced = list(self.state["forced"])
        if forced:
            menu.addSection("Included anyway")
            for k in forced:
                act = menu.addAction(f"Leave out {Path(k).stem} again")
                act.triggered.connect(lambda _=False, key=k: self._apply_edit(
                    {"forced": [f for f in self.state["forced"] if f != key]}, reload_data=True))
        self.note_btn.setMenu(menu)
        self.note_btn.setText(("ⓘ " + " · ".join(parts)) if parts else ("ⓘ" if forced else ""))
        self.note_btn.setVisible(bool(parts or forced))

    # ------------------------------------------------------------------
    # Rendering
    # ------------------------------------------------------------------
    def _new_figure(self):
        Figure, FigureCanvasAgg = _figure_classes()
        fig = Figure()
        FigureCanvasAgg(fig)
        return fig

    def _molecules(self):
        """Molecule objects for the state (rebuilt only when it changes)."""
        key = repr(self.state.get("molecules"))
        if self._mol_cache[0] != key:
            mols = []
            for d in self.state.get("molecules") or []:
                try:
                    mols.append(_molecule_from(d))
                except Exception:
                    pass
            self._mol_cache = (key, mols)
        return self._mol_cache[1]

    def _draw_into(self, fig, screen=False):
        mols = self._molecules()
        draw = None
        if mols:
            painter = _MoleculePainter(mols, self._mol_settings())
            if screen:
                self._mol_axes, self._mol_artists = [], []

            def draw(ax, r, f, w, h):
                arts = painter.draw(ax, _mol_scale(ax, fig, self.style.get("mol_size", 1.0)))
                for i, a in enumerate(arts):
                    a.set_gid(f"molecule_{r}_{f}_{i}")
                if screen:
                    self._mol_axes.append(ax)
                    self._mol_artists.extend(arts)
        elif screen:
            self._mol_axes, self._mol_artists = [], []
        return hs.render(
            fig, self.spec, self._panel_map(), style=self.style, clims=self.state["clims"],
            letter_colors=self.state.get("letter_colors"), draw_molecules=draw,
            get_cmap=lambda name: registry_get_cmap(name, "gray"))

    def _render_screen(self):
        if not self._ready():
            self.view.add_row_box = None
            self.view.set_image(None, (1, 1), None)
            self.layout_info = None
            return
        if self._fig is None:
            self._fig = self._new_figure()
        self.layout_info = self._draw_into(self._fig, screen=True)
        self.view.add_row_box = self._add_row_box()
        self.view.sel_boxes = self._mol_sel_boxes()
        mols = self._molecules()
        self.view.gizmo = ({"vectors": _gizmo_vectors(mols[self._mol_sel])}
                           if self._mol_sel is not None and self._mol_sel < len(mols) else None)
        self._blit()

    def _add_row_box(self):
        """Screen-only "+ channel" handle in an empty margin: the top-left
        corner when channels are rows, left of the channel labels otherwise."""
        lay = self.layout_info
        if not lay or not self.spec.get("channels"):
            return None
        shown = {r["label"] for r in self.spec["rows"]}
        if all(c["label"] in shown for c in self.spec["channels"]):
            return None
        if not self.style.get("channel_labels", True) or self.style.get("header") == "none":
            return None   # no free margin for it - the Channels tab adds channels
        if not lay["transpose"]:
            first = min((b for _f, b in lay["frame_labels"]), key=lambda b: (b[1], b[0]))
            return (0.0, first[1], max(first[0], 0.6), first[3])
        boxes = sorted((b for _r, b in lay["channel_labels"]), key=lambda b: b[0])
        first = boxes[0]
        return (0.0, first[1], max(first[0] - 0.04, 0.6), first[3])

    def _screen_dpi(self):
        try:
            scr = self.windowHandle().screen() if self.windowHandle() else QtWidgets.QApplication.primaryScreen()
            return float(scr.physicalDotsPerInch()) or 96.0
        except Exception:
            return 96.0

    def _set_actual(self, on):
        self._actual = bool(on)
        if not self._actual:
            self.view.setMinimumSize(320, 160)
        self._render_screen()
        self._flash("Actual size (100 %)" if self._actual else "Fit to window")

    def _blit(self):
        fig = self._fig
        fig_w, fig_h = self.layout_info["size"]
        dpr = float(self.devicePixelRatioF() or 1.0)
        if self._actual:
            # Physical size: an 85 mm figure is 85 mm on this screen.
            dpi = self._screen_dpi()
            self.view.setMinimumSize(int(fig_w * dpi) + 2 * _PAPER_MARGIN + 16,
                                     int(fig_h * dpi) + 2 * _PAPER_MARGIN + 40)
        else:
            avail_w = max(50, self.scroll.viewport().width() - 2 * _PAPER_MARGIN) - 16
            avail_h = max(50, self.scroll.viewport().height() - 2 * _PAPER_MARGIN) - 16
            if self.layout_info.get("page_in"):
                avail_h -= 22   # room for the ruler
            dpi = max(30.0, min(avail_w / fig_w, avail_h / fig_h, 400.0))
        page = self.style.get("page") or "free"
        if self.layout_info.get("page_in"):
            mm = self.layout_info["page_in"] * 25.4
            self.view.ruler_text = f"{mm:.0f} mm  ·  {hs.PAGE_NAMES.get(page, '').split(' (')[0]}" + \
                ("  ·  100 %" if self._actual else "")
        else:
            self.view.ruler_text = (f"{fig_w * 25.4:.0f} × {fig_h * 25.4:.0f} mm  ·  100 %"
                                    if self._actual else "")
        fig.set_dpi(dpi * dpr)
        fig.canvas.draw()
        self._push_buffer()

    def _push_buffer(self):
        """Agg buffer of the screen figure -> the view (no re-draw)."""
        fig = self._fig
        fig_w, fig_h = self.layout_info["size"]
        dpr = float(self.devicePixelRatioF() or 1.0)
        buf = np.asarray(fig.canvas.buffer_rgba())
        h, w = buf.shape[:2]
        img = QtGui.QImage(buf.data, w, h, 4 * w, QtGui.QImage.Format_RGBA8888).copy()
        img.setDevicePixelRatio(dpr)
        self.view.set_image(img, (fig_w, fig_h), self.layout_info)

    def _redraw_fast(self):
        if self._mol_drag is not None:
            self._mol_drag_paint()
            return
        if self._fig is not None and self.layout_info is not None:
            self._blit()

    # ------------------------------------------------------------------
    # Live link
    # ------------------------------------------------------------------
    def _check_live(self):
        if not self.isVisible() or self._drag is not None:
            return
        try:
            sig = self.source.data_signature(self._all_keys())
        except Exception:
            return
        if sig != self._data_sig:
            self._rebuild(reload_data=True)
            self._flash("Updated from source")

    def closeEvent(self, event):
        self._live_timer.stop()
        try:
            self.source.forget_window(self.fid, self)
        except Exception:
            pass
        super().closeEvent(event)

    def changeEvent(self, event):
        super().changeEvent(event)
        if event.type() == QtCore.QEvent.ActivationChange and self.isActiveWindow():
            self._check_live()

    # ------------------------------------------------------------------
    # Hit-testing
    # ------------------------------------------------------------------
    def _hit_test(self, x, y):
        lay = self.layout_info
        if not lay or not self.spec:
            return None
        n_cols = len(self.spec["columns"])
        for f, box in lay["frame_labels"]:
            if _inside(box, x, y):
                return {"kind": "frame", "frame": f, "box": box, "can_close": n_cols > 1}
        for r, box in lay["channel_labels"]:
            if _inside(box, x, y):
                return {"kind": "channel", "row": r, "box": box}
        for entry in lay["colorbars"]:
            if _inside(entry["box"], x, y, pad=0.05):
                return {"kind": "colorbar", "row": entry["row"], "frame": entry["frame"],
                        "box": entry["box"], "orient": entry["orient"]}
        for entry in lay.get("letter_boxes", []):
            if _inside(entry["box"], x, y, pad=0.02):
                return {"kind": "letter", "row": entry["row"], "frame": entry["frame"], "box": entry["box"]}
        for entry in lay.get("scale_boxes", []):
            if _inside(entry["box"], x, y, pad=0.03):
                return {"kind": "scalebar", "row": entry["row"], "frame": entry["frame"], "box": entry["box"]}
        mi = self._molecule_at(x, y)
        if mi is not None:
            return {"kind": "molecule", "index": mi[0], "row": mi[1], "frame": mi[2], "box": mi[3]}
        for (r, f), box in lay["cells"].items():
            if _inside(box, x, y):
                return {"kind": "panel", "row": r, "frame": f, "box": box}
        return None

    # ------------------------------------------------------------------
    # Molecular model: one placement (panel nm) shared by every panel
    # ------------------------------------------------------------------
    def _to_data(self, r, f, x, y):
        entry = (self.layout_info or {}).get("data_boxes", {}).get((r, f))
        if not entry:
            return None
        (x0, y0, x1, y1), w, h = entry
        if x1 <= x0 or y1 <= y0:
            return None
        return (x - x0) / (x1 - x0) * w, (y1 - y) / (y1 - y0) * h

    def _mol_bbox(self, mol, pad=0.12):
        c = mol.get_transformed_coordinates()
        if c is None or not len(c):
            return None
        return (float(c[:, 0].min()) - pad, float(c[:, 0].max()) + pad,
                float(c[:, 1].min()) - pad, float(c[:, 1].max()) + pad)

    def _mol_panels(self):
        lay = self.layout_info or {}
        if not self.style.get("molecules_on", True):
            return []
        return [rf for rf in lay.get("data_boxes", {})
                if self.style.get("mol_scope") != "first" or rf[0] == 0]

    def _molecule_at(self, x, y):
        mols = self._molecules()
        if not mols:
            return None
        for (r, f) in self._mol_panels():
            entry = self.layout_info["data_boxes"][(r, f)]
            if not _inside(entry[0], x, y):
                continue
            d = self._to_data(r, f, x, y)
            if d is None:
                continue
            for i in range(len(mols) - 1, -1, -1):
                bb = self._mol_bbox(mols[i])
                if bb and bb[0] <= d[0] <= bb[1] and bb[2] <= d[1] <= bb[3]:
                    return i, r, f, entry[0]
        return None

    def _mol_sel_boxes(self, mols=None):
        """Inch boxes of the selected molecule on every panel (screen-only
        dashed highlight)."""
        mols = mols if mols is not None else self._molecules()
        if self._mol_sel is None or self._mol_sel >= len(mols):
            return []
        bb = self._mol_bbox(mols[self._mol_sel])
        if not bb:
            return []
        out = []
        for (r, f) in self._mol_panels():
            (x0, y0, x1, y1), w, h = self.layout_info["data_boxes"][(r, f)]
            sx, sy = (x1 - x0) / max(w, 1e-9), (y1 - y0) / max(h, 1e-9)
            out.append((x0 + max(0.0, bb[0]) * sx, y1 - min(h, bb[3]) * sy,
                        x0 + min(w, bb[1]) * sx, y1 - max(0.0, bb[2]) * sy))
        return out

    def _set_molecules(self, mols, push=True):
        self._apply_edit({"molecules": [m.to_dict() for m in mols]}, push=push)

    def _on_mol_action(self, action):
        mols = [m.copy() for m in self._molecules()]
        if action == "load":
            start = str(self.source.export_settings().get("mol_dir") or "")
            path, _f = QtWidgets.QFileDialog.getOpenFileName(
                self, "Molecular model", start, "Molecules (*.pdb *.xyz *.mol);;All files (*)")
            if not path:
                return
            self.source.set_export_settings(mol_dir=str(Path(path).parent))
            mol = _molecule_from({})
            try:
                mol.load(path)
                mol.filepath = path
            except Exception as exc:
                QtWidgets.QMessageBox.warning(self, "Molecular model", f"Could not read the model:\n{exc}")
                return
            if not len(mol.coordinates):
                QtWidgets.QMessageBox.warning(self, "Molecular model", "No atoms found in that file.")
                return
            dims = next(iter((self.layout_info or {}).get("data_boxes", {}).values()), None)
            if dims:
                mol.offset = np.array([dims[1] / 2.0, dims[2] / 2.0, 0.0])
            mols.append(mol)
            self._mol_sel = len(mols) - 1
            self._set_molecules(mols)
            self._flash(f"{Path(path).name}: on every panel - drag to place, X/Y/Z to rotate")
        elif action == "preview":
            found, src = self.source.preview_molecules(self._all_keys())
            if not found:
                QtWidgets.QMessageBox.information(
                    self, "Molecular model", "None of these frames has a molecule in the preview yet.")
                return
            self._mol_sel = 0
            self._apply_edit({"molecules": found})
            self._flash(f"Model(s) from {Path(src).stem} - same spot on every panel")
        elif action == "props":
            self._mol_properties(self._mol_sel if self._mol_sel is not None else 0)
        elif action == "remove":
            if not mols:
                return
            if self._mol_sel is not None and self._mol_sel < len(mols):
                mols.pop(self._mol_sel)
            else:
                mols = []
            self._mol_sel = None
            self._set_molecules(mols)

    def _mol_settings(self):
        """The preview's molecule look, with this figure's own palette
        when one is chosen (style ``mol_palette``)."""
        settings = dict(self.source.molecule_settings())
        pal = str(self.style.get("mol_palette") or "preview").lower()
        if pal != "preview":
            settings["palette"] = pal
        return settings

    def _mol_properties(self, i):
        mols = [m.copy() for m in self._molecules()]
        if not mols or i is None or i >= len(mols):
            return
        from ..canvases.molecular_overlay import MoleculePropertiesDialog
        before = self._snapshot()
        mol = mols[i]

        def changed():
            # live: every panel follows the dialog
            self.state = dict(self.state, molecules=[m.to_dict() for m in mols])
            self._rebuild()

        dlg = MoleculePropertiesDialog(mol, parent=self, callback=changed,
                                       overlay_settings={"show_shadows_available": False,
                                                         "show_hydrogens_available": False,
                                                         "palette_available": False})
        dlg.exec_()
        if self.state != before:
            self._undo.append(before)
            self._redo.clear()
            self.source.set_state(self.fid, self.state)

    def _mol_transform(self, i, *, rot=None, move=None, mirror=None, reset=False, remove=False):
        mols = [m.copy() for m in self._molecules()]
        if i is None or i >= len(mols):
            return
        m = mols[i]
        if remove:
            mols.pop(i)
            self._mol_sel = None
        elif reset:
            try:
                mols[i] = m.reset_to_file_state(keep_offset=True)
            except Exception:
                m.angles = np.zeros(3)
        else:
            molecule_interaction.apply_transform(m, rot=rot, move=move, mirror=mirror)
        self._set_molecules(mols)

    def keyPressEvent(self, event):
        i = self._mol_sel
        if i is not None and i < len(self._molecules()):
            key = event.key()
            shift = bool(event.modifiers() & QtCore.Qt.ShiftModifier)
            step = 0.5 if shift else 0.05
            axes = {QtCore.Qt.Key_X: 0, QtCore.Qt.Key_Y: 1, QtCore.Qt.Key_Z: 2}
            if key in axes:
                self._mol_transform(i, rot=molecule_interaction.key_rotation(axes[key], reverse=shift))
                return
            moves = {QtCore.Qt.Key_Left: (-step, 0), QtCore.Qt.Key_Right: (step, 0),
                     QtCore.Qt.Key_Up: (0, step), QtCore.Qt.Key_Down: (0, -step)}
            if key in moves:
                self._mol_transform(i, move=moves[key])
                return
            if key in (QtCore.Qt.Key_Delete, QtCore.Qt.Key_Backspace):
                self._mol_transform(i, remove=True)
                return
            if key == QtCore.Qt.Key_Escape:
                self._mol_sel = None
                self._render_screen()
                return
        super().keyPressEvent(event)

    def _mol_begin_drag(self, hit, event, mode="translate", local=None):
        d = None
        if hit.get("row") is not None:
            inch = self.view.to_inches(event.pos())
            d = self._to_data(hit["row"], hit["frame"], *inch) if inch else None
            if d is None:
                return
        self._mol_sel = hit["index"]
        fig = self._fig
        # Background without the molecules, cached once; each move then
        # redraws only the molecule artists (fast with many panels).
        for a in self._mol_artists:
            a.set_visible(False)
        fig.canvas.draw()
        bg = fig.canvas.copy_from_bbox(fig.bbox)
        for a in self._mol_artists:
            try:
                a.remove()
            except Exception:
                pass
        self._mol_artists = []
        mols = [m.copy() for m in self._molecules()]
        self._mol_drag = {"index": hit["index"], "row": hit["row"], "frame": hit["frame"], "start": d,
                          "offset0": np.asarray(mols[hit["index"]].offset, dtype=float).copy(),
                          "angles0": np.asarray(mols[hit["index"]].angles, dtype=float).copy(),
                          "mode": mode, "local": local, "px0": QtCore.QPointF(event.pos()),
                          "mols": mols, "bg": bg, "arts": [], "before": self._snapshot(), "moved": False}
        self._mol_drag_paint()

    def _mol_drag_paint(self):
        d = self._mol_drag
        if d is None:
            return
        fig = self._fig
        for a in d["arts"]:
            try:
                a.remove()
            except Exception:
                pass
        painter = _MoleculePainter(d["mols"], self._mol_settings())
        arts = []
        for ax in self._mol_axes:
            arts.extend((ax, a) for a in painter.draw(ax, _mol_scale(ax, fig, self.style.get("mol_size", 1.0))))
        d["arts"] = [a for _ax, a in arts]
        fig.canvas.restore_region(d["bg"])
        for ax, a in arts:
            ax.draw_artist(a)
        self.view.sel_boxes = self._mol_sel_boxes(d["mols"])
        self.view.gizmo = {"vectors": _gizmo_vectors(d["mols"][d["index"]])}
        self._push_buffer()

    def _mol_drag_move(self, event):
        d = self._mol_drag
        mol = d["mols"][d["index"]]
        mode = d["mode"]
        pos = QtCore.QPointF(event.pos())
        dpx = pos.x() - d["px0"].x()
        dpy = -(pos.y() - d["px0"].y())          # y up, like the preview's events
        # The shared gesture math (molecules/interaction.py), as in the preview.
        p = cur_local = None
        if mode in ("translate", "rotate_z"):
            inch = self.view.to_inches(event.pos())
            p = self._to_data(d["row"], d["frame"], *inch) if inch else None
            if p is None:
                return
        elif mode == "gizmo_rotate_z":
            c = self.view.gizmo_center()
            cur_local = (pos.x() - c.x(), -(pos.y() - c.y()))
        offset, angles = molecule_interaction.drag_update(
            mode, d["offset0"], d["angles0"], start=d["start"], current=p,
            delta_px=(dpx, dpy), start_local=d["local"], current_local=cur_local)
        if mode == "translate":
            mol.offset = offset
        else:
            mol.angles = angles
        d["moved"] = True
        if not self._drag_timer.isActive():
            self._drag_timer.start()

    def _mol_drag_end(self):
        d, self._mol_drag = self._mol_drag, None
        if d["moved"]:
            self._undo.append(d["before"])
            self._redo.clear()
            self.state = dict(self._snapshot(), molecules=[m.to_dict() for m in d["mols"]])
            self.source.set_state(self.fid, self.state)
        self._render_screen()

    def _on_hover(self, hit):
        cursor = QtCore.Qt.ArrowCursor
        tip = ""
        if hit:
            kind = hit["kind"]
            if kind == "close":
                cursor = QtCore.Qt.PointingHandCursor
                tip = "Remove this frame from the figure (Ctrl+Z restores)"
            elif kind == "frame":
                col = self.spec["columns"][hit["frame"]]
                if self.spec["has_z"] and self.spec["header"] == "dz":
                    cursor = QtCore.Qt.PointingHandCursor
                    tip = (f"{col['name']}\nThis frame is the Δz reference" if col["is_ref"]
                           else f"{col['name']}\nClick: make this the Δz = 0 reference")
                else:
                    tip = col["name"]
                tip += "\nRight-click: header options"
            elif kind == "molecule":
                cursor = QtCore.Qt.SizeAllCursor
                tip = ("Molecular model - same spot on every panel\nDrag: move  ·  Shift+drag: spin  ·  "
                       "Ctrl+Shift+drag or middle button: rotate in 3-D\nX / Y / Z: rotate 5° (Shift: back)  ·  "
                       "arrows: nudge  ·  Del: remove  ·  right-click: more")
            elif kind == "gizmo":
                cursor = QtCore.Qt.ClosedHandCursor
                tip = ("Orientation gizmo - drag inside the ring: tumble in 3-D  ·  "
                       "drag on the rim: spin in plane")
            elif kind == "channel":
                cursor = QtCore.Qt.PointingHandCursor
                row = self.spec["rows"][hit["row"]]
                tip = f"Channel: {hs.channel_text(row['label'])}\nClick: switch, recolour, move or remove"
            elif kind == "add_row":
                cursor = QtCore.Qt.PointingHandCursor
                tip = "Add another channel of these frames"
            elif kind == "colorbar":
                cursor = QtCore.Qt.SizeHorCursor if hit["orient"] == "h" else QtCore.Qt.SizeVerCursor
                tip = ("Drag: contrast (one end, Shift = both)  ·  Click: type Min/Max  ·  "
                       "Double-click: automatic")
            elif kind == "panel":
                col = self.spec["columns"][hit["frame"]]
                row = self.spec["rows"][hit["row"]]
                z = col.get("z_nm")
                tip = (f"{col['name']}  ·  {row['label']}" + (f"\nz_abs {z:.3f} nm" if z is not None else "")
                       + "\nDouble-click: show in the viewer  ·  Right-click: options")
        self.view.setCursor(cursor)
        self.view.setToolTip(tip)

    # ------------------------------------------------------------------
    # Clicks and drags
    # ------------------------------------------------------------------
    def _on_press(self, hit, event):
        if not hit:
            if self._mol_sel is not None:
                self._mol_sel = None
                self._render_screen()
            return
        kind = hit["kind"]
        if kind == "gizmo" and event is not None and self._mol_sel is not None:
            self._mol_begin_drag({"index": self._mol_sel, "row": None, "frame": None}, event,
                                 mode="gizmo_" + hit["mode"], local=hit["local"])
            return
        if event is not None and event.button() == QtCore.Qt.MiddleButton and kind != "molecule":
            return
        if kind == "molecule" and event is not None:
            self.setFocus()
            mods = event.modifiers()
            # Same gestures as the preview: drag = move, Shift = spin in
            # plane, Ctrl+Shift or middle button = free 3-D rotation.
            mode = molecule_interaction.drag_mode(
                middle=event.button() == QtCore.Qt.MiddleButton,
                shift=bool(mods & QtCore.Qt.ShiftModifier),
                ctrl=bool(mods & QtCore.Qt.ControlModifier))
            self._mol_begin_drag(hit, event, mode=mode)
            return
        gpos = event.globalPos() if event is not None else QtGui.QCursor.pos()
        if kind == "close":
            col = self.spec["columns"][hit["frame"]]
            self._apply_edit({"dropped": self.state["dropped"] + [col["key"]]})
            self._flash(f"Removed {col['name']} - Ctrl+Z to restore")
        elif kind == "frame" and self.spec["has_z"] and self.spec["header"] == "dz":
            col = self.spec["columns"][hit["frame"]]
            if not col["is_ref"]:
                self._apply_edit({"reference_key": col["key"]})
        elif kind == "channel":
            self.view._pressed = False
            self._channel_menu(hit["row"]).exec_(gpos)
        elif kind == "add_row":
            self.view._pressed = False
            self._add_channel_menu().exec_(gpos)
        elif kind == "colorbar":
            self._begin_cbar_drag(hit, event)
        elif kind in ("letter", "scalebar"):
            # A click on a letter / scale bar opens its card next to it.
            self.view._pressed = False
            card = "letters" if kind == "letter" else "scale"
            QtCore.QTimer.singleShot(0, lambda c=card, p=gpos: self.open_format_card(c, p))

    def open_format_card(self, kind, global_pos=None):
        """Show one Format card (the bar is turned on if it was hidden)."""
        if self.format_panel.isHidden():
            self.format_btn.setChecked(True)
        self._sync_format_panel()
        self.format_panel.open_card(kind, global_pos)

    def _card_action(self, menu, kind, text, gpos):
        menu.addSeparator()
        menu.addAction(text).triggered.connect(lambda: self.open_format_card(kind, gpos))

    def _cbar_range(self, r, f):
        lay = self.layout_info or {}
        if f is None:
            return lay.get("ranges", {}).get(r, (0.0, 1.0))
        return lay.get("panel_ranges", {}).get((r, f), (0.0, 1.0))

    def _begin_cbar_drag(self, hit, event):
        r, f = hit["row"], hit["frame"]
        if f is None and self.style["range_mode"] == "image":
            self._flash("Each image has its own range - drag or click an image's own bar "
                        "(Colour ▸ Bars: Follow the range)")
            return
        lo, hi = self._cbar_range(r, f)
        inch = self.view.to_inches(event.pos())
        x0, y0, x1, y1 = hit["box"]
        if event.modifiers() & QtCore.Qt.ShiftModifier:
            end = "both"
        elif hit["orient"] == "h":
            end = "hi" if inch and inch[0] > 0.5 * (x0 + x1) else "lo"
        else:
            end = "hi" if inch and inch[1] < 0.5 * (y0 + y1) else "lo"
        rect = self.view.inches_rect(hit["box"])
        self._drag = {"row": r, "frame": f, "orient": hit["orient"], "start": event.pos(),
                      "lo": lo, "hi": hi, "end": end, "moved": False,
                      "px": max(20.0, rect.width() if hit["orient"] == "h" else rect.height()),
                      "before": self._snapshot(), "gpos": event.globalPos()}

    def _on_drag_move(self, event):
        if self._mol_drag is not None:
            self._mol_drag_move(event)
            return
        d = self._drag
        if not d:
            return
        if d["orient"] == "h":
            delta_px = event.pos().x() - d["start"].x()
        else:
            delta_px = -(event.pos().y() - d["start"].y())
        if abs(delta_px) >= _CLICK_SLOP_PX:
            d["moved"] = True
        if not d["moved"]:
            return
        span = d["hi"] - d["lo"]
        delta = delta_px / d["px"] * span
        lo, hi = d["lo"], d["hi"]
        tiny = 1e-6 * abs(span or 1.0)
        if d["end"] == "hi":
            hi = max(lo + tiny, hi + delta)
        elif d["end"] == "lo":
            lo = min(hi - tiny, lo + delta)
        else:
            lo, hi = lo + delta, hi + delta
        d["current"] = (lo, hi)
        if d["frame"] is None:
            hs.set_row_range(self.layout_info, d["row"], lo, hi)
        else:
            hs.set_panel_range(self.layout_info, d["row"], d["frame"], lo, hi)
        if not self._drag_timer.isActive():
            self._drag_timer.start()

    def _on_release(self, _event):
        if self._mol_drag is not None:
            self._mol_drag_end()
            return
        d, self._drag = self._drag, None
        if not d:
            return
        if not d["moved"]:
            self._range_editor(d["row"], d["frame"]).exec_(d["gpos"])
            return
        if "current" in d:
            self._set_clim(d["row"], d["frame"], d["current"], before=d["before"])

    def _clim_key(self, r, f):
        row = self.spec["rows"][r]
        return hs.clim_key(row["label"], None if f is None else self.spec["columns"][f]["key"])

    def _set_clim(self, r, f, lim, before=None):
        clims = dict(self.state["clims"])
        key = self._clim_key(r, f)
        if lim is None:
            clims.pop(key, None)
        else:
            clims[key] = [float(lim[0]), float(lim[1])]
        if before is not None:
            # The drag already drew the result - just record it.
            self._undo.append(before)
            self._redo.clear()
            self.state = dict(self._snapshot(), clims=clims)
            self.source.set_state(self.fid, self.state)
            self._redraw_fast()
        else:
            self._apply_edit({"clims": clims})

    def _on_double_click(self, hit):
        if not hit:
            return
        if hit["kind"] == "colorbar":
            self._set_clim(hit["row"], hit["frame"], None)
        elif hit["kind"] == "panel":
            col = self.spec["columns"][hit["frame"]]
            row = self.spec["rows"][hit["row"]]
            idx = row["cells"].get(col["key"])
            if idx is not None:
                self.source.show_image(col["key"], idx)

    def _show_key(self, key):
        for row in (self.spec or {}).get("rows", []):
            idx = row["cells"].get(key)
            if idx is not None:
                self.source.show_image(key, idx)
                return

    # ------------------------------------------------------------------
    # Menus (right-click on any element; same options as the Format panel)
    # ------------------------------------------------------------------
    def _on_context(self, hit, gpos):
        kind = (hit or {}).get("kind")
        if kind == "channel":
            menu = self._channel_menu(hit["row"])
        elif kind == "add_row":
            menu = self._add_channel_menu()
        elif kind == "colorbar":
            menu = self._colour_menu(hit["row"], hit["frame"])
            self._card_action(menu, "colour", "All colour options…", gpos)
        elif kind == "molecule":
            i = hit["index"]
            self._mol_sel = i
            menu = QtWidgets.QMenu(self)
            mol = self._molecules()[i]
            menu.addSection(Path(str(mol.filepath or "Molecule")).name)
            menu.addAction("Appearance, rotation, scale…").triggered.connect(lambda: self._mol_properties(i))
            for text, rot in (("Rotate +90° (in plane)", [0, 0, 90]), ("Flip over (X 180°)", [180, 0, 0])):
                menu.addAction(text).triggered.connect(lambda _=False, r_=rot: self._mol_transform(i, rot=r_))
            menu.addAction("Mirror left-right").triggered.connect(lambda: self._mol_transform(i, mirror="x"))
            menu.addAction("Mirror up-down").triggered.connect(lambda: self._mol_transform(i, mirror="y"))
            menu.addAction("Reset orientation").triggered.connect(lambda: self._mol_transform(i, reset=True))
            menu.addSeparator()
            self._choice(menu, "Shown on", "mol_scope", {"all": "Every panel", "first": "First channel row"})
            self._choice(menu.addMenu("Atom palette"), None, "mol_palette", _mol_palette_choices())
            menu.addSeparator()
            menu.addAction("Remove this model").triggered.connect(lambda: self._mol_transform(i, remove=True))
            self._card_action(menu, "molecule", "All molecule options…", gpos)
        elif kind == "letter":
            menu = QtWidgets.QMenu(self)
            self._single_letter_menu(menu, hit["row"], hit["frame"])
            self._letters_menu(menu.addMenu("All letters"))
            self._card_action(menu, "letters", "All letter options…", gpos)
        elif kind == "scalebar":
            menu = QtWidgets.QMenu(self)
            self._scale_menu(menu)
            self._card_action(menu, "scale", "All scale bar options…", gpos)
        else:
            menu = QtWidgets.QMenu(self)
            if kind in ("frame", "close"):
                col = self.spec["columns"][hit["frame"]]
                menu.addSection(col["name"])
                if self.spec["has_z"]:
                    act = menu.addAction("Make this the Δz reference")
                    act.setEnabled(not col["is_ref"])
                    act.triggered.connect(lambda: self._apply_edit(
                        {"reference_key": col["key"]}, style={"header": "dz"}))
                self._frame_actions(menu, col)
                self._header_menu(menu.addMenu("Headers"))
            elif kind == "panel":
                col = self.spec["columns"][hit["frame"]]
                row = self.spec["rows"][hit["row"]]
                menu.addSection(f"{col['name']}  ·  {row['display']}")
                self._frame_actions(menu, col)
                if self.style["range_mode"] == "image":
                    menu.addAction("Colour range of this image…").triggered.connect(
                        lambda: self._range_editor(hit["row"], hit["frame"]).exec_(gpos))
                self._letters_menu(menu.addMenu("Panel letters"))
                self._scale_menu(menu.addMenu("Scale bar"))
                self._colour_menu(hit["row"], None, menu.addMenu("Colour"))
                self._channel_submenu(menu.addMenu("Channel"), hit["row"])
            else:
                self._layout_menu(menu)
                self._letters_menu(menu.addMenu("Panel letters"))
                self._scale_menu(menu.addMenu("Scale bar"))
                self._header_menu(menu.addMenu("Headers"))
        menu.exec_(gpos)

    def _frame_actions(self, menu, col):
        menu.addAction("Show in the viewer").triggered.connect(lambda: self._show_key(col["key"]))
        if len(self.spec["columns"]) > 1:
            menu.addAction("Remove this frame").triggered.connect(
                lambda: self._apply_edit({"dropped": self.state["dropped"] + [col["key"]]}))

    def _choice(self, menu, title, key, options):
        if title:
            menu.addSection(title)
        group = QtWidgets.QActionGroup(menu)
        for value, text in options.items():
            act = menu.addAction(text)
            act.setCheckable(True)
            act.setChecked(self.style.get(key) == value)
            group.addAction(act)
            act.triggered.connect(lambda _=False, v=value: self._apply_edit(style={key: v}))

    def _toggle(self, menu, text, key):
        act = menu.addAction(text)
        act.setCheckable(True)
        act.setChecked(bool(self.style.get(key)))
        act.triggered.connect(lambda on: self._apply_edit(style={key: bool(on)}))
        return act

    def _layout_menu(self, menu):
        self._choice(menu, "Layout", "transpose", {False: "Frames across, channels down",
                                                    True: "Frames down, channels across"})
        menu.addSeparator()
        self._toggle(menu, "Reverse the frame order", "reverse")
        self._choice(menu, "Tall images", "rotate_tall", _ROTATE_TALL)
        menu.addSection("Page")
        group = QtWidgets.QActionGroup(menu)
        for key, text in hs.PAGE_NAMES.items():
            if key == "custom":
                continue
            act = menu.addAction(text)
            act.setCheckable(True)
            act.setChecked((self.style.get("page") or "free") == key)
            group.addAction(act)
            act.triggered.connect(lambda _=False, k=key: self._apply_edit(
                style=dict(hs.page_text_style(k), page=k)))
        menu.addAction("More format options…").triggered.connect(lambda: self.format_btn.setChecked(True))

    def _single_letter_menu(self, menu, r, f):
        """Colour of just this letter (overrides the all-letters choice)."""
        row = self.spec["rows"][r]
        col = self.spec["columns"][f]
        key = hs.letter_key(row["label"], col["key"])
        cur = (self.state.get("letter_colors") or {}).get(key)
        menu.addSection("This letter")
        group = QtWidgets.QActionGroup(menu)
        for value, text in ((None, "Like the others"), ("white", "White"), ("black", "Black")):
            act = menu.addAction(text)
            act.setCheckable(True)
            act.setChecked(cur == value)
            group.addAction(act)
            act.triggered.connect(lambda _=False, v=value: self._set_letter_color(key, v))
        other = menu.addAction("Other colour…")
        other.setCheckable(True)
        other.setChecked(cur not in (None, "white", "black"))
        group.addAction(other)
        other.triggered.connect(lambda: self._pick_letter_color(key, cur))
        if self.state.get("letter_colors"):
            menu.addAction("Reset every letter to automatic").triggered.connect(
                lambda: self._apply_edit({"letter_colors": {}}, style={"letter_color": "auto"}))

    def _set_letter_color(self, key, value):
        colors = dict(self.state.get("letter_colors") or {})
        if value is None:
            colors.pop(key, None)
        else:
            colors[key] = value
        self._apply_edit({"letter_colors": colors})

    def _pick_letter_color(self, key, cur):
        start = QtGui.QColor(cur) if cur and cur not in ("white", "black") else QtGui.QColor("#ffd23f")
        c = QtWidgets.QColorDialog.getColor(start, self, "Letter colour")
        if c.isValid():
            self._set_letter_color(key, c.name())

    def _letters_menu(self, menu):
        self._toggle(menu, "Show letters", "letters")
        self._choice(menu, "Style", "letter_style", _LETTER_STYLE_NAMES)
        self._choice(menu, "On", "letter_scope", {"frame": "One per frame", "panel": "Every panel"})
        self._choice(menu, "Place", "letter_pos", _LETTER_PLACES)
        self._choice(menu, "Colour (all letters)", "letter_color", _AUTO_COLORS)
        menu.addSection("Size")
        group = QtWidgets.QActionGroup(menu)
        for pt in (7, 9, 11, 14):
            act = menu.addAction(f"{pt} pt")
            act.setCheckable(True)
            act.setChecked(abs(float(self.style["letter_size"]) - pt) < 0.01)
            group.addAction(act)
            act.triggered.connect(lambda _=False, v=float(pt): self._apply_edit(style={"letter_size": v}))
        menu.addSeparator()
        self._toggle(menu, "Bold", "letter_bold")
        self._toggle(menu, "Outline", "letter_outline")

    def _scale_menu(self, menu):
        self._choice(menu, "Show on", "scale_bar", _SCALE_NAMES)
        menu.addSection("Length")
        act = menu.addAction("Automatic")
        act.setCheckable(True)
        act.setChecked(self.style.get("scale_len_nm") is None)
        act.triggered.connect(lambda: self._apply_edit(style={"scale_len_nm": None}))
        menu.addAction("Set length…").triggered.connect(self._ask_scale_length)
        self._choice(menu, "Corner", "scale_pos", _CORNER_NAMES)
        self._choice(menu, "Colour", "scale_color", _AUTO_COLORS)
        self._choice(menu, "Length label", "scale_label", _SCALE_LABELS)
        menu.addSection("Label size")
        group = QtWidgets.QActionGroup(menu)
        for size in (None, 6.0, 7.0, 8.0, 10.0):
            act = menu.addAction("Text size" if size is None else f"{size:g} pt")
            act.setCheckable(True)
            act.setChecked(self.style.get("scale_font") == size)
            group.addAction(act)
            act.triggered.connect(lambda _=False, v=size: self._apply_edit(style={"scale_font": v}))
        menu.addSeparator()
        self._toggle(menu, "Outline", "scale_outline")

    def _ask_scale_length(self):
        cur = self.style.get("scale_len_nm") or 1.0
        val, ok = QtWidgets.QInputDialog.getDouble(self, "Scale bar", "Length (nm):", float(cur),
                                                   0.01, 100000.0, 2)
        if ok:
            where = self.style["scale_bar"] if self.style["scale_bar"] != "off" else "one"
            self._apply_edit(style={"scale_len_nm": float(val), "scale_bar": where})

    def _header_menu(self, menu):
        self._choice(menu, "Show", "header", _HEADER_NAMES)
        menu.addSeparator()
        self._toggle(menu, "Add the file name", "header_name")
        self._toggle(menu, "Bold the Δz = 0 reference", "ref_bold")
        self._toggle(menu, "Channel names beside the rows", "channel_labels")

    def _colour_menu(self, r, f, menu=None):
        menu = menu or QtWidgets.QMenu(self)
        row = self.spec["rows"][r]
        menu.addSection(f"{row['display']} colour")
        shared = self.style["range_mode"] == "row"
        target = f if not shared else None
        if shared or target is not None:
            menu.addAction("Set range…").triggered.connect(
                lambda: self._range_editor(r, target).exec_(QtGui.QCursor.pos()))
            act = menu.addAction("Automatic range")
            act.setEnabled(self._clim_key(r, target) in self.state["clims"])
            act.triggered.connect(lambda: self._set_clim(r, target, None))
        self._choice(menu, "Range", "range_mode", {"row": "One range per channel",
                                                   "image": "Each image its own range"})
        menu.addSeparator()
        self._toggle(menu, "Colour bars", "colorbars")
        self._choice(menu, "Ticks", "cbar_ticks", _CBAR_TICKS)
        menu.addSeparator()
        self._toggle(menu, "Channel and unit label", "cbar_label")
        self._choice(menu, "Channel name", "cbar_label_pos", _CBAR_LABEL_POS)
        self._choice(menu, "Values", "cbar_style", _CBAR_STYLES)
        menu.addSeparator()
        self._toggle(menu, "Frame around the bar", "cbar_outline")
        self._cmap_menu(menu.addMenu("Colormap"), r)
        return menu

    def _cmap_menu(self, menu, r):
        row = self.spec["rows"][r]
        group = QtWidgets.QActionGroup(menu)
        for name in self.source.featured_cmaps():
            act = menu.addAction(name)
            act.setCheckable(True)
            act.setChecked(name == row["cmap"])
            group.addAction(act)
            act.triggered.connect(lambda _=False, n=name: self._set_row_cmap(r, n))

    def _set_row_cmap(self, r, name):
        cmaps = dict(self.state["row_cmaps"])
        cmaps[self.spec["rows"][r]["label"]] = name
        self._apply_edit({"row_cmaps": cmaps})

    def _current_rows(self):
        return [row["label"] for row in self.spec["rows"]]

    def _channel_menu(self, r):
        menu = QtWidgets.QMenu(self)
        self._channel_submenu(menu, r)
        self._colour_menu(r, None, menu.addMenu("Colour"))
        return menu

    def _channel_submenu(self, menu, r):
        row = self.spec["rows"][r]
        rows = self._current_rows()
        menu.addSection("Show channel")
        group = QtWidgets.QActionGroup(menu)
        same_kind = [c for c in self.spec["channels"] if c["kind"] == row["kind"]]
        other = [c for c in self.spec["channels"] if c["kind"] != row["kind"]]
        best = next((c for c in self.spec["channels"] if c["label"] == hs.BEST), None)
        if best is not None:
            act = menu.addAction(hs.channel_text(hs.BEST))
            act.setCheckable(True)
            act.setChecked(row["label"] == hs.BEST)
            act.setToolTip("Each frame shows its own main channel (topography for constant current, "
                           "current for constant height ...)")
            act.setEnabled(row["label"] == hs.BEST or hs.BEST not in rows)
            act.triggered.connect(lambda _=False: self._replace_row(r, hs.BEST))
            menu.addSeparator()
        same_kind = [c for c in same_kind if c["label"] != hs.BEST]
        other = [c for c in other if c["label"] != hs.BEST]
        for c in same_kind:
            act = menu.addAction(c["label"])
            act.setCheckable(True)
            act.setChecked(c["label"] == row["label"])
            act.setEnabled(c["label"] == row["label"] or c["label"] not in rows)
            group.addAction(act)
            act.triggered.connect(lambda _=False, lab=c["label"]: self._replace_row(r, lab))
        if other:
            sub = menu.addMenu("Other channels")
            for c in other:
                act = sub.addAction(c["label"])
                act.setEnabled(c["label"] not in rows)
                act.triggered.connect(lambda _=False, lab=c["label"]: self._replace_row(r, lab))
        menu.addSeparator()
        horiz = bool(self.layout_info and self.layout_info.get("transpose"))
        up = menu.addAction("Move left" if horiz else "Move up")
        up.setEnabled(r > 0)
        up.triggered.connect(lambda: self._move_row(r, -1))
        down = menu.addAction("Move right" if horiz else "Move down")
        down.setEnabled(r < len(rows) - 1)
        down.triggered.connect(lambda: self._move_row(r, +1))
        rm = menu.addAction("Remove this channel")
        rm.setEnabled(len(rows) > 1)
        rm.triggered.connect(lambda: self._apply_edit({"rows": [x for i, x in enumerate(rows) if i != r]}))
        add = self._add_channel_menu(menu.addMenu("Add channel"))
        add.setEnabled(not add.isEmpty())

    def _add_channel_menu(self, menu=None):
        menu = menu or QtWidgets.QMenu(self)
        rows = self._current_rows()
        for c in self.spec["channels"]:
            if c["label"] in rows:
                continue
            menu.addAction(hs.channel_text(c["label"])).triggered.connect(
                lambda _=False, lab=c["label"]: self._apply_edit({"rows": self._current_rows() + [lab]}))
        return menu

    def _replace_row(self, r, label):
        rows = self._current_rows()
        if label == rows[r]:
            return
        rows[r] = label
        self._apply_edit({"rows": rows})
        self._flash(f"Showing {hs.channel_text(label)}")

    def _move_row(self, r, step):
        rows = self._current_rows()
        j = r + step
        if 0 <= j < len(rows):
            rows[r], rows[j] = rows[j], rows[r]
            self._apply_edit({"rows": rows})

    def _range_editor(self, r, f):
        """Popover with exact Min/Max for one colourbar (Enter applies)."""
        menu = QtWidgets.QMenu(self)
        row = self.spec["rows"][r]
        lo, hi = self._cbar_range(r, f)
        unit = (self.layout_info or {}).get("units", {}).get(r, "")
        title = row["display"] if f is None else f"{row['display']} · {self.spec['columns'][f]['name']}"
        box = QtWidgets.QWidget(menu)
        form = QtWidgets.QFormLayout(box)
        form.setContentsMargins(10, 6, 10, 6)
        form.addRow(QtWidgets.QLabel(f"<b>{title}</b>" + (f"  ({unit})" if unit else "")))
        span = abs(hi - lo) or 1.0
        decimals = max(0, min(6, int(3 - np.floor(np.log10(span)))))
        spins = []
        for text, val in (("Max", hi), ("Min", lo)):
            sp = QtWidgets.QDoubleSpinBox(box)
            sp.setRange(-1e15, 1e15)
            sp.setDecimals(decimals)
            sp.setSingleStep(span / 20.0)
            sp.setValue(float(val))
            sp.setMinimumWidth(110)
            form.addRow(text, sp)
            spins.append(sp)
        btns = QtWidgets.QHBoxLayout()
        auto = QtWidgets.QPushButton("Automatic", box)
        ok = QtWidgets.QPushButton("Apply", box)
        ok.setDefault(True)
        btns.addWidget(auto)
        btns.addWidget(ok)
        form.addRow(btns)

        def apply():
            v_hi, v_lo = spins[0].value(), spins[1].value()
            if v_hi < v_lo:
                v_hi, v_lo = v_lo, v_hi
            menu.close()
            if v_hi > v_lo:
                self._set_clim(r, f, (v_lo, v_hi))

        def automatic():
            menu.close()
            self._set_clim(r, f, None)

        ok.clicked.connect(apply)
        auto.clicked.connect(automatic)
        for sp in spins:
            sp.lineEdit().returnPressed.connect(apply)
        act = QtWidgets.QWidgetAction(menu)
        act.setDefaultWidget(box)
        menu.addAction(act)
        menu.range_spins = spins   # (tests)
        menu.range_apply = apply
        QtCore.QTimer.singleShot(0, lambda: (spins[0].setFocus(), spins[0].selectAll()))
        return menu

    # ------------------------------------------------------------------
    # Outputs
    # ------------------------------------------------------------------
    def _export_figure(self):
        fig = self._new_figure()
        self._draw_into(fig)
        return fig

    def _default_name(self):
        cols = self.spec["columns"] if self.spec else []
        stem = cols[0]["name"] if cols else "figure"
        suffix = {"height": "height_series", "bias": "bias_series", "zoom": "zooms"}.get(self.mode, "figure")
        return f"{stem}_{suffix}"

    def _export_opts(self):
        st = self.source.export_settings()
        dpi = int(st.get("dpi") or 300)
        transparent = bool(st.get("transparent", False))
        return dpi, transparent

    def _face(self):
        return "none" if self._export_opts()[1] else "white"

    def _svg_bytes(self, fig=None, layered=False):
        import matplotlib
        fig = fig or self._export_figure()
        buf = io.BytesIO()
        with matplotlib.rc_context({"svg.fonttype": "none"}):
            fig.savefig(buf, format="svg", facecolor=self._face())
        data = buf.getvalue()
        return svg_layers.layered_svg(data) if layered else data

    def _png_bytes(self, fig=None, dpi=None):
        fig = fig or self._export_figure()
        buf = io.BytesIO()
        fig.savefig(buf, format="png", dpi=dpi or self._export_opts()[0], facecolor=self._face())
        return buf.getvalue()

    def _options_menu(self, menu):
        """Resolution + background, shared by Copy / Save / Obsidian."""
        dpi, transparent = self._export_opts()
        sub = menu.addMenu(f"Resolution: {dpi} dpi")
        group = QtWidgets.QActionGroup(sub)
        for d in _DPI_CHOICES:
            act = sub.addAction(f"{d} dpi")
            act.setCheckable(True)
            act.setChecked(d == dpi)
            group.addAction(act)
            act.triggered.connect(lambda _=False, v=d: self.source.set_export_settings(dpi=v))
        act = menu.addAction("Transparent background")
        act.setCheckable(True)
        act.setChecked(transparent)
        act.triggered.connect(lambda on: self.source.set_export_settings(transparent=bool(on)))

    def _copy_menu(self, menu):
        menu.addAction("PNG + SVG (paste anywhere)").triggered.connect(lambda: self.copy_to_clipboard("both"))
        menu.addAction("PNG only").triggered.connect(lambda: self.copy_to_clipboard("png"))
        menu.addAction("SVG only (vector)").triggered.connect(lambda: self.copy_to_clipboard("svg"))
        menu.addSeparator()
        self._options_menu(menu)

    def _ppt_menu(self, menu):
        last = self.source.export_settings().get("ppt_mode") or "new_vector"
        for mode, text in (("new_vector", "New slide - editable shapes"),
                           ("cur_vector", "Current slide - editable shapes"),
                           ("new_picture", "New slide - picture"),
                           ("cur_picture", "Current slide - picture")):
            act = menu.addAction(text)
            act.setCheckable(True)
            act.setChecked(mode == last)
            act.triggered.connect(lambda _=False, m=mode: self.send_to_powerpoint(m))
        menu.addSeparator()
        self._options_menu(menu)

    def _inkscape_menu(self, menu):
        exe = self._inkscape_exe()
        act = menu.addAction("Open in Inkscape" + ("" if exe else " (not found)"))
        act.setEnabled(bool(exe))
        act.triggered.connect(self.open_in_inkscape)
        menu.addAction("Save layered SVG…").triggered.connect(lambda: self.save_as(layered=True))

    def _save_menu(self, menu):
        for text, ext in (("PDF (vector)", ".pdf"), ("SVG (vector, layered)", ".svg"),
                          ("PNG", ".png"), ("TIFF", ".tif"), ("EPS (vector)", ".eps")):
            menu.addAction(text).triggered.connect(lambda _=False, e=ext: self.save_as(ext=e))
        menu.addSeparator()
        self._options_menu(menu)

    def _ready(self):
        return bool(self.spec and self.spec["columns"] and self.spec["rows"])

    def copy_to_clipboard(self, mode="both"):
        if not self._ready():
            return
        fig = self._export_figure()
        mime = QtCore.QMimeData()
        if mode in ("both", "png"):
            mime.setImageData(QtGui.QImage.fromData(self._png_bytes(fig), "PNG"))
        if mode in ("both", "svg"):
            try:
                svg = self._svg_bytes(fig)
                mime.setData("image/svg+xml", QtCore.QByteArray(svg))
                if mode == "svg":
                    mime.setText(svg.decode("utf-8", "replace"))
            except Exception:
                pass
        QtWidgets.QApplication.clipboard().setMimeData(mime)
        self._flash({"both": "Copied - paste anywhere", "png": "Copied as PNG",
                     "svg": "Copied as SVG"}[mode])

    def send_to_powerpoint(self, mode=None):
        if not self._ready():
            return
        mode = mode or self.source.export_settings().get("ppt_mode") or "new_vector"
        self.source.set_export_settings(ppt_mode=mode)
        try:
            from .. import ppt_bridge
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "PowerPoint", f"PowerPoint export unavailable: {exc}")
            return
        ok, reason = ppt_bridge.powerpoint_support_status()
        if not ok:
            QtWidgets.QMessageBox.information(self, "PowerPoint", reason or "PowerPoint export unavailable.")
            return
        QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.WaitCursor)
        try:
            new_slide = mode.startswith("new")
            fig = self._export_figure()
            w_pt, h_pt = (float(v) * 72.0 for v in fig.get_size_inches())
            # Real size: the box *is* the figure (1 pt = 1/72 in), so an
            # 85 mm figure lands 85 mm wide, like PDF / Inkscape.
            size_kw = dict(left=36.0, top=36.0, width=w_pt, height=h_pt,
                           image_size=(w_pt, h_pt), preserve_aspect=True)
            if mode.endswith("picture"):
                pix = QtGui.QPixmap.fromImage(QtGui.QImage.fromData(self._png_bytes(fig), "PNG"))
                slide, _name = ppt_bridge.send_pixmap_to_ppt(pix, label=None, new_slide=new_slide, **size_kw)
            else:
                slide, _name = ppt_bridge.send_svg_to_ppt(self._svg_bytes(fig), label=None,
                                                          new_slide=new_slide, **size_kw)
        except Exception as exc:
            QtWidgets.QApplication.restoreOverrideCursor()
            QtWidgets.QMessageBox.warning(
                self, "PowerPoint",
                f"Could not send to PowerPoint:\n{exc}\n\nIs a presentation open?")
            return
        QtWidgets.QApplication.restoreOverrideCursor()
        self._flash(f"Sent to PowerPoint (slide {slide}) at {w_pt / 72 * 25.4:.0f} × {h_pt / 72 * 25.4:.0f} mm")

    _SAVE_FILTERS = {".pdf": "PDF (*.pdf)", ".svg": "SVG (*.svg)", ".png": "PNG image (*.png)",
                     ".tif": "TIFF image (*.tif *.tiff)", ".eps": "EPS (*.eps)"}

    def save_as(self, ext=None, layered=False):
        if not self._ready():
            return
        st = self.source.export_settings()
        ext = ext or (".svg" if layered else st.get("save_ext") or ".pdf")
        folder = st.get("save_dir") or self.source.default_save_dir(self._keys)
        start = str(Path(folder) / (self._default_name() + ext))
        order = [ext] + [e for e in self._SAVE_FILTERS if e != ext]
        path, _flt = QtWidgets.QFileDialog.getSaveFileName(
            self, "Save figure", start, ";;".join(self._SAVE_FILTERS[e] for e in order))
        if not path:
            return
        ext = Path(path).suffix.lower()
        if ext == ".tiff":
            ext = ".tif"
        if ext not in self._SAVE_FILTERS:
            path, ext = path + ".pdf", ".pdf"
        fig = self._export_figure()
        dpi, _transparent = self._export_opts()
        try:
            if ext == ".svg":
                # Always layered: a plain SVG viewer ignores the layers,
                # Inkscape shows them.
                Path(path).write_bytes(self._svg_bytes(fig, layered=True))
            elif ext == ".tif":
                fig.savefig(path, dpi=dpi, facecolor=self._face(),
                            pil_kwargs={"compression": "tiff_lzw"})
            elif ext == ".png":
                fig.savefig(path, dpi=dpi, facecolor=self._face())
            else:
                fig.savefig(path, facecolor=self._face())
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Save figure", f"Could not save:\n{exc}")
            return
        self.source.set_export_settings(save_ext=ext, save_dir=str(Path(path).parent))
        self._flash(f"Saved {Path(path).name}")

    @staticmethod
    def _inkscape_exe():
        import shutil
        exe = shutil.which("inkscape")
        if exe:
            return exe
        for cand in (r"C:\Program Files\Inkscape\bin\inkscape.exe",
                     r"C:\Program Files (x86)\Inkscape\bin\inkscape.exe",
                     r"C:\Program Files\Inkscape\inkscape.exe",
                     "/Applications/Inkscape.app/Contents/MacOS/inkscape"):
            if Path(cand).exists():
                return cand
        return None

    def open_in_inkscape(self):
        if not self._ready():
            return
        exe = self._inkscape_exe()
        if not exe:
            QtWidgets.QMessageBox.information(
                self, "Inkscape", "Inkscape was not found. Install it from inkscape.org, or use "
                "Inkscape ▸ Save layered SVG… and open the file yourself.")
            return
        import subprocess
        import tempfile
        folder = Path(tempfile.gettempdir()) / "sxm_viewer_figures"
        folder.mkdir(parents=True, exist_ok=True)
        path = folder / (self._default_name() + ".svg")
        try:
            path.write_bytes(self._svg_bytes(layered=True))
            subprocess.Popen([exe, str(path)])
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Inkscape", f"Could not open Inkscape:\n{exc}")
            return
        self._flash(f"Opened in Inkscape ({path.name}) - save it from there")

    def _obsidian_render(self, img_dir):
        """render() callback for the Obsidian export: one PNG of the figure."""
        img_dir = Path(img_dir)
        img_dir.mkdir(parents=True, exist_ok=True)
        stem = self._default_name()
        path = img_dir / f"{stem}.png"
        n = 1
        while path.exists():
            n += 1
            path = img_dir / f"{stem}_{n}.png"
        path.write_bytes(self._png_bytes())
        cols = self.spec["columns"]
        rows = ", ".join(r["display"] for r in self.spec["rows"])
        caption = (f"*{_MODE_TITLES.get(self.mode, 'Figure')}* · {len(cols)} frames · {rows} · "
                   + " / ".join(c["header"] for c in cols[:8]) + (" …" if len(cols) > 8 else ""))
        key = cols[0]["key"] if cols else ""
        idx = next(iter(self.spec["rows"][0]["cells"].values()), 0) if self.spec["rows"] else 0
        return [{"path": path, "file_key": key, "channel_idx": idx, "caption": caption}], [], False

    # ------------------------------------------------------------------
    # Ctrl+wheel: size whatever is under the cursor
    # ------------------------------------------------------------------
    def _on_wheel(self, hit, event):
        import time
        steps = event.angleDelta().y() / 120.0
        if not steps:
            return
        up = steps > 0
        st = self.style
        kind = (hit or {}).get("kind")
        shift = bool(event.modifiers() & QtCore.Qt.ShiftModifier)
        if shift:
            if (st.get("page") or "free") != "free":
                self._flash("Panel size follows the page width - set Page to Free size to change it")
                return
            change = {"panel_w": round(min(4.0, max(0.6, st["panel_w"] * (1.06 if up else 1 / 1.06))), 3)}
            msg = f"Panels {change['panel_w'] * 25.4:.0f} mm"
        elif kind == "colorbar":
            change = {"cbar_w": round(min(0.32, max(0.02, st["cbar_w"] + (0.01 if up else -0.01))), 3)}
            msg = f"Colour bar {change['cbar_w'] * 25.4:.1f} mm"
        elif kind == "scalebar":
            cur = float(st.get("scale_font") or st["font"])
            change = {"scale_font": min(16.0, max(4.5, cur + (0.5 if up else -0.5)))}
            msg = f"Scale bar label {change['scale_font']:g} pt"
        elif kind == "letter":
            change = {"letter_size": min(24.0, max(5.0, st["letter_size"] + (0.5 if up else -0.5)))}
            msg = f"Letters {change['letter_size']:g} pt"
        else:
            # Text size everywhere, letters keep their proportion.
            font = min(14.0, max(5.0, st["font"] + (0.5 if up else -0.5)))
            ratio = st["letter_size"] / st["font"] if st["font"] else 1.3
            change = {"font": font, "letter_size": round(min(24.0, max(5.0, font * ratio)), 1)}
            msg = f"Text {font:g} pt"
        now = time.monotonic()
        push = (now - self._last_wheel) > _WHEEL_COALESCE_S
        self._last_wheel = now
        self._apply_edit(style=change, push=push)
        self._flash(msg)

    def open_in_canvas(self):
        if not self.spec:
            return
        cells = []
        for r, row in enumerate(self.spec["rows"]):
            for c, col in enumerate(self.spec["columns"]):
                idx = row["cells"].get(col["key"])
                if idx is not None:
                    cells.append((r, c, col["key"], idx, row["cmap"], row["kind"]))
        self.source.open_in_canvas(cells, len(self.spec["columns"]))

    def _flash(self, text):
        self.status.setText(text)
        QtCore.QTimer.singleShot(2500, lambda: self.status.setText("")
                                 if self.status.text() == text else None)
