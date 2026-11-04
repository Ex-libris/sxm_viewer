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

    # mouse --------------------------------------------------------------
    def _hit(self, pos):
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
        if event.button() != QtCore.Qt.LeftButton:
            return
        self._pressed = True
        self.pressed.emit(self._hit(event.pos()), event)

    def mouseReleaseEvent(self, event):
        if event.button() != QtCore.Qt.LeftButton:
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
    elif kind == "zoom":
        p.drawRect(QtCore.QRectF(3, 3, 16, 16))
        p.setPen(QtGui.QPen(QtGui.QColor("#ff9f1c"), 1.6))
        p.drawRect(QtCore.QRectF(10, 5, 7, 7))
    p.end()
    return QtGui.QIcon(pix)


class _FormatPanel(QtWidgets.QWidget):
    """Every figure option at once: a row of cards under the figure, each
    with an icon, a title and a one-line hint (mirrors the right-click
    menus). Emits ``changed(dict)`` with the style/state keys to update."""

    changed = QtCore.pyqtSignal(dict)

    CARDS = (
        ("page", "Page & layout", "Journal width, arrangement"),
        ("channels", "Channels", "Tick to show · drag to reorder"),
        ("titles", "Text", "Font, size, headers, names"),
        ("colour", "Colour", "Ranges and colour bars"),
        ("letters", "Panel letters", "a, b, c … style and place"),
        ("scale", "Scale bar", "Where, length, label"),
        ("zoom", "Zoom outlines", "Boxes of the zooms"),
    )

    def __init__(self, parent=None):
        super().__init__(parent)
        outer = QtWidgets.QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)
        self.scroll = QtWidgets.QScrollArea(self)
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.scroll.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)
        strip = QtWidgets.QWidget()
        self.row = QtWidgets.QHBoxLayout(strip)
        self.row.setContentsMargins(6, 4, 6, 4)
        self.row.setSpacing(8)
        self.scroll.setWidget(strip)
        outer.addWidget(self.scroll)
        self._syncing = False
        self._cards = {}
        icon_color = self.palette().color(QtGui.QPalette.WindowText)
        w = {}

        def card(kind, columns=1):
            """A titled card; returns add(label, widget) placing rows in
            ``columns`` label/field column pairs, top to bottom."""
            title, blurb = next((t, b) for k, t, b in self.CARDS if k == kind)
            frame = QtWidgets.QFrame()
            frame.setObjectName("figureCard")
            frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
            frame.setStyleSheet("QFrame#figureCard { border: 1px solid palette(midlight); border-radius: 6px; }")
            col = QtWidgets.QVBoxLayout(frame)
            col.setContentsMargins(8, 6, 8, 6)
            col.setSpacing(4)
            head = QtWidgets.QHBoxLayout()
            head.setSpacing(6)
            ic = QtWidgets.QLabel()
            ic.setPixmap(_tab_icon(kind, icon_color).pixmap(22, 22))
            head.addWidget(ic)
            tl = QtWidgets.QLabel(f"<b>{title}</b>")
            head.addWidget(tl)
            head.addStretch(1)
            col.addLayout(head)
            hint = QtWidgets.QLabel(blurb)
            hint.setStyleSheet("color: palette(mid); font-size: 8pt;")
            col.addWidget(hint)
            grid = QtWidgets.QGridLayout()
            grid.setHorizontalSpacing(6)
            grid.setVerticalSpacing(3)
            col.addLayout(grid)
            col.addStretch(1)
            self.row.addWidget(frame)
            self._cards[kind] = frame
            state = {"n": 0}
            per_col = None

            def add(label, widget, span=False):
                nonlocal per_col
                i = state["n"]
                state["n"] += 1
                c = (i // per_col) if per_col else 0
                r = (i % per_col) if per_col else i
                if span or label is None:
                    if isinstance(widget, QtWidgets.QLayout):
                        grid.addLayout(widget, r, 2 * c, 1, 2)
                    else:
                        grid.addWidget(widget, r, 2 * c, 1, 2)
                    return
                lab = QtWidgets.QLabel(label)
                lab.setStyleSheet("color: palette(text);")
                grid.addWidget(lab, r, 2 * c)
                if isinstance(widget, QtWidgets.QLayout):
                    grid.addLayout(widget, r, 2 * c + 1)
                else:
                    grid.addWidget(widget, r, 2 * c + 1)

            def set_rows(n):
                nonlocal per_col
                per_col = int(math.ceil(n / columns))

            add.set_rows = set_rows
            return add

        def fit(c):
            # Readable fixed width in the card; the pop-up list shows the
            # full names (no card grows to its longest entry).
            c.setSizeAdjustPolicy(QtWidgets.QComboBox.AdjustToMinimumContentsLengthWithIcon)
            c.setMinimumContentsLength(15)
            c.setMaximumWidth(190)
            view = c.view()
            view.setMinimumWidth(max(view.sizeHintForColumn(0) + 28, 190))
            return c

        def combo(items):
            c = QtWidgets.QComboBox()
            for key, text in items.items():
                c.addItem(text, key)
            return fit(c)

        def spin(lo, hi, step, suffix):
            s = QtWidgets.QDoubleSpinBox()
            s.setRange(lo, hi)
            s.setSingleStep(step)
            s.setSuffix(suffix)
            s.setMinimumWidth(78)
            return s

        def check(text, tip=None):
            b = QtWidgets.QCheckBox(text)
            if tip:
                b.setToolTip(tip)
            return b

        # -- Page & layout
        add = card("page")
        add.set_rows(7)
        w["page"] = combo(hs.PAGE_NAMES)
        w["page"].setToolTip("Journal column width: sets the figure width and every text size")
        add("Page", w["page"])
        w["page_mm"] = spin(30.0, 400.0, 1.0, " mm")
        add("Custom width", w["page_mm"])
        w["_actual"] = check("Actual size (100 %)", "Draw the figure at its printed size on this screen")
        add(None, w["_actual"])
        w["transpose"] = combo({False: "Frames → across, channels ↓ down",
                                True: "Frames ↓ down, channels → across"})
        add("Arrange", w["transpose"])
        w["reverse"] = check("Reverse order", "Reverse the frame order")
        add(None, w["reverse"])
        w["rotate_tall"] = combo(_ROTATE_TALL)
        w["rotate_tall"].setToolTip("Tall (portrait) images turned a quarter turn for this figure; "
                                    "scale bars and zoom outlines follow")
        add("Tall images", w["rotate_tall"])
        w["panel_w"] = spin(0.6, 4.0, 0.1, " in")
        w["panel_w"].setToolTip("Free size only - a page width sets it (Ctrl+Shift+wheel)")
        add("Panel width", w["panel_w"])

        # -- Channels
        add = card("channels")
        w["channels"] = QtWidgets.QListWidget()
        w["channels"].setDragDropMode(QtWidgets.QAbstractItemView.InternalMove)
        w["channels"].setMinimumWidth(180)
        w["channels"].setMaximumWidth(230)
        w["channels"].setToolTip("Tick the channels to show; drag to reorder")
        add(None, w["channels"])

        # -- Text
        add = card("titles")
        add.set_rows(6)
        w["font_family"] = QtWidgets.QComboBox()
        w["font_family"].setMaxVisibleItems(18)
        for name in hs.available_fonts():
            w["font_family"].addItem(name, name)
        w["font_family"].setToolTip("Font of every text in the figure (journal-friendly fonts first)")
        fit(w["font_family"])
        add("Font", w["font_family"])
        w["font"] = spin(5.0, 14.0, 0.5, " pt")
        w["font"].setToolTip("Text size (Ctrl+wheel on the figure)")
        add("Size", w["font"])
        w["header"] = combo(_HEADER_NAMES)
        add("Frame header", w["header"])
        w["header_name"] = check("+ file name", "Add the file name under the header")
        add(None, w["header_name"])
        w["ref_bold"] = check("Bold Δz = 0", "Bold header for the Δz = 0 reference frame")
        add(None, w["ref_bold"])
        w["channel_labels"] = check("Channel names",
                                    "I, Δf ... left of the rows (above them when frames run down)")
        add(None, w["channel_labels"])

        # -- Colour
        add = card("colour", columns=2)
        add.set_rows(10)
        w["range_mode"] = combo({"row": "Shared per channel", "image": "Each image its own"})
        add("Range", w["range_mode"])
        w["colorbars"] = check("Colour bars")
        add(None, w["colorbars"])
        w["cbar_scope"] = combo({"auto": "Follow the range", "row": "One per channel (min / max)"})
        w["cbar_scope"].setToolTip("One bar per channel even when each image has its own range - "
                                   "its ends then read min and max")
        add("Bars", w["cbar_scope"])
        w["cbar_style"] = combo(_CBAR_STYLES)
        add("Values", w["cbar_style"])
        w["cbar_ticks"] = combo(_CBAR_TICKS)
        w["cbar_ticks"].setToolTip("None = a plain gradient, no tick marks or values")
        add("Ticks", w["cbar_ticks"])
        w["cbar_label"] = check("Channel and unit")
        add(None, w["cbar_label"])
        w["cbar_label_pos"] = combo(_CBAR_LABEL_POS)
        add("Name", w["cbar_label_pos"])
        w["cbar_text"] = combo(_CBAR_TEXT)
        add("Text", w["cbar_text"])
        w["cbar_w"] = spin(0.5, 8.0, 0.25, " mm")
        add("Width", w["cbar_w"])
        w["cbar_outline"] = check("Bar frame", "Thin frame around each colour bar")
        add(None, w["cbar_outline"])

        # -- Panel letters
        add = card("letters")
        add.set_rows(8)
        w["letters"] = check("Show letters")
        add(None, w["letters"])
        w["letter_style"] = combo(_LETTER_STYLE_NAMES)
        add("Style", w["letter_style"])
        w["letter_scope"] = combo({"frame": "One per frame", "panel": "Every panel"})
        add("On", w["letter_scope"])
        w["letter_pos"] = combo(_LETTER_PLACES)
        w["letter_pos"].setToolTip("Inside a corner, or outside the image above it (dark on the page)")
        add("Place", w["letter_pos"])
        w["letter_color"] = combo({"white": "White", "black": "Black"})
        add("Colour", w["letter_color"])
        w["letter_size"] = spin(5.0, 24.0, 0.5, " pt")
        add("Size", w["letter_size"])
        w["letter_bold"] = check("Bold")
        add(None, w["letter_bold"])
        w["letter_outline"] = check("Outline")
        add(None, w["letter_outline"])

        # -- Scale bar
        add = card("scale")
        add.set_rows(8)
        w["scale_bar"] = combo(_SCALE_NAMES)
        add("Show on", w["scale_bar"])
        len_row = QtWidgets.QHBoxLayout()
        w["scale_auto"] = check("Auto")
        w["scale_len_nm"] = spin(0.01, 100000.0, 0.5, " nm")
        w["scale_len_nm"].setDecimals(2)
        len_row.addWidget(w["scale_auto"])
        len_row.addWidget(w["scale_len_nm"], 1)
        add("Length", len_row)
        w["scale_label"] = combo(_SCALE_LABELS)
        add("Label", w["scale_label"])
        w["scale_font"] = spin(4.5, 16.0, 0.5, " pt")
        w["scale_font"].setSpecialValueText("Text size")
        w["scale_font"].setToolTip("Size of the length label (Ctrl+wheel on a scale bar)")
        add("Label size", w["scale_font"])
        w["scale_pos"] = combo(_CORNER_NAMES)
        add("Corner", w["scale_pos"])
        w["scale_color"] = combo({"white": "White", "black": "Black"})
        add("Colour", w["scale_color"])
        w["scale_thick"] = spin(0.5, 5.0, 0.25, " %")
        w["scale_thick"].setToolTip("Bar thickness, % of the panel's shorter side")
        add("Thickness", w["scale_thick"])
        w["scale_outline"] = check("Outline")
        add(None, w["scale_outline"])

        # -- Zoom outlines (zoom figures only)
        add = card("zoom")
        add.set_rows(3)
        w["zoom_color"] = combo(_ZOOM_COLORS)
        add("Colour", w["zoom_color"])
        w["zoom_lw"] = spin(0.25, 4.0, 0.25, " pt")
        add("Line width", w["zoom_lw"])
        w["zoom_labels"] = check("Letters", "Letter next to each zoom outline")
        add(None, w["zoom_labels"])

        end = QtWidgets.QVBoxLayout()
        end.addStretch(1)
        self.reset_btn = QtWidgets.QPushButton("Reset format")
        self.reset_btn.setToolTip("Back to the defaults (your figures keep their data choices)")
        end.addWidget(self.reset_btn)
        self.row.addLayout(end)
        self.row.addStretch(1)
        self.w = w

        for key, widget in w.items():
            if isinstance(widget, QtWidgets.QComboBox):
                widget.currentIndexChanged.connect(lambda _i, k=key: self._emit(k))
            elif isinstance(widget, QtWidgets.QCheckBox):
                widget.toggled.connect(lambda _on, k=key: self._emit(k))
            elif isinstance(widget, QtWidgets.QDoubleSpinBox):
                widget.valueChanged.connect(lambda _v, k=key: self._emit(k))
        w["channels"].itemChanged.connect(lambda _i: self._emit("channels"))
        w["channels"].model().rowsMoved.connect(lambda *_a: self._emit("channels"))

    def sizeHint(self):
        return QtCore.QSize(1200, 250)

    def set_mode(self, mode):
        frame = self._cards.get("zoom")
        if frame is not None:
            frame.setVisible(mode == "zoom")

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
        widget = w[key]
        if isinstance(widget, QtWidgets.QComboBox):
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

    def sync(self, style, spec, auto_scale_nm=None, actual=False):
        self._syncing = True
        try:
            w = self.w
            w["_actual"].setChecked(bool(actual))
            for key, widget in w.items():
                if key in ("channels", "scale_auto", "scale_len_nm", "_actual"):
                    continue
                value = style.get(key)
                if key == "font_family" and not value:
                    value = hs.default_font()
                if key in self._TO_WIDGET and value is not None:
                    value = self._TO_WIDGET[key](value)
                if isinstance(widget, QtWidgets.QComboBox):
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
                        "letter_size", "letter_bold"):
                w[key].setEnabled(letters_on)
            scale_on = style.get("scale_bar") != "off"
            for key in ("scale_auto", "scale_pos", "scale_color", "scale_label", "scale_font"):
                w[key].setEnabled(scale_on)
            w["scale_len_nm"].setEnabled(scale_on and not auto)
            bars_on = bool(style.get("colorbars"))
            for key in ("cbar_ticks", "cbar_label", "cbar_style", "cbar_text", "cbar_w",
                        "cbar_label_pos", "cbar_outline"):
                w[key].setEnabled(bars_on)
            if bars_on and style.get("cbar_style") in ("minmax", "minmax_in"):
                for key in ():
                    w[key].setEnabled(False)
            w["page_mm"].setEnabled(style.get("page") == "custom")
            w["panel_w"].setEnabled((style.get("page") or "free") == "free")
            for key in ("scale_thick", "scale_outline"):
                w[key].setEnabled(scale_on)
        finally:
            self._syncing = False


class FigureWindow(QtWidgets.QWidget):
    """A live, data-aware figure (see module docstring)."""

    def __init__(self, source, fid, keys, *, mode="height", excluded=None, root_key=None, parent=None):
        super().__init__(parent, QtCore.Qt.Window)
        self.setAttribute(QtCore.Qt.WA_DeleteOnClose, True)
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
            "row_cmaps": {}, "clims": {}, "show_repeats": False,
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

        self.splitter = QtWidgets.QSplitter(QtCore.Qt.Vertical, self)
        self.splitter.setChildrenCollapsible(False)
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
        self.splitter.addWidget(self.scroll)
        self.format_panel = _FormatPanel(self)
        self.format_panel.changed.connect(self._on_format_changed)
        self.format_panel.reset_btn.clicked.connect(self.reset_format)
        self.format_panel.set_mode(self.mode)
        self.splitter.addWidget(self.format_panel)
        self.splitter.setStretchFactor(0, 1)
        self.splitter.setStretchFactor(1, 0)
        self.splitter.setSizes([520, 260])
        open_fmt = bool(self.source.export_settings().get("format_open", True))
        self.format_panel.setVisible(open_fmt)
        self.format_btn.blockSignals(True)
        self.format_btn.setChecked(open_fmt)
        self.format_btn.blockSignals(False)
        root.addWidget(self.splitter, 1)

        self.hint = QtWidgets.QLabel(self)
        self.hint.setStyleSheet("color: palette(mid); font-size: 8pt;")
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
        if not self.format_panel.isVisible() or not self.spec:
            return
        auto_len = None
        widths = [p.get("width_nm") for p in self._panel_map().values() if p.get("width_nm")]
        if widths:
            auto_len = hs.nice_length(0.3 * widths[0])
        self.format_panel.sync(self.style, self.spec, auto_len, actual=self._actual)

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

    def _draw_into(self, fig):
        return hs.render(
            fig, self.spec, self._panel_map(), style=self.style, clims=self.state["clims"],
            get_cmap=lambda name: registry_get_cmap(name, "gray"))

    def _render_screen(self):
        if not self._ready():
            self.view.add_row_box = None
            self.view.set_image(None, (1, 1), None)
            self.layout_info = None
            return
        if self._fig is None:
            self._fig = self._new_figure()
        self.layout_info = self._draw_into(self._fig)
        self.view.add_row_box = self._add_row_box()
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
        buf = np.asarray(fig.canvas.buffer_rgba())
        h, w = buf.shape[:2]
        img = QtGui.QImage(buf.data, w, h, 4 * w, QtGui.QImage.Format_RGBA8888).copy()
        img.setDevicePixelRatio(dpr)
        self.view.set_image(img, (fig_w, fig_h), self.layout_info)

    def _redraw_fast(self):
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
        for (r, f), box in lay["cells"].items():
            if _inside(box, x, y):
                return {"kind": "panel", "row": r, "frame": f, "box": box}
        return None

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
            return
        kind = hit["kind"]
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
        elif kind == "letter":
            menu = QtWidgets.QMenu(self)
            self._letters_menu(menu)
        elif kind == "scalebar":
            menu = QtWidgets.QMenu(self)
            self._scale_menu(menu)
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

    def _letters_menu(self, menu):
        self._toggle(menu, "Show letters", "letters")
        self._choice(menu, "Style", "letter_style", _LETTER_STYLE_NAMES)
        self._choice(menu, "On", "letter_scope", {"frame": "One per frame", "panel": "Every panel"})
        self._choice(menu, "Place", "letter_pos", _LETTER_PLACES)
        self._choice(menu, "Colour", "letter_color", {"white": "White", "black": "Black"})
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
        self._choice(menu, "Colour", "scale_color", {"white": "White", "black": "Black"})
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
            if mode.endswith("picture"):
                pix = QtGui.QPixmap.fromImage(QtGui.QImage.fromData(self._png_bytes(), "PNG"))
                slide, _name = ppt_bridge.send_pixmap_to_ppt(pix, label=None, new_slide=new_slide)
            else:
                slide, _name = ppt_bridge.send_svg_to_ppt(self._svg_bytes(), label=None, new_slide=new_slide)
        except Exception as exc:
            QtWidgets.QApplication.restoreOverrideCursor()
            QtWidgets.QMessageBox.warning(
                self, "PowerPoint",
                f"Could not send to PowerPoint:\n{exc}\n\nIs a presentation open?")
            return
        QtWidgets.QApplication.restoreOverrideCursor()
        self._flash(f"Sent to PowerPoint (slide {slide})")

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
