"""Figure window for data-aware figure recipes (height series, selection).

One clean window: the figure on a white "paper" (exactly what Copy /
PowerPoint / PDF produce - one matplotlib renderer, ``figures.height_series
.render``), a thin strip with the three outputs and two toggles, and every
adjustment made directly on the figure:

* click a column's Δz header      -> that frame becomes the Δz reference
* hover a header, click its ×     -> drop the frame (Ctrl+Z restores)
* click a row label (I, Δf)       -> cycle that row's channel (fwd/bwd, ...)
* drag on a row's colourbar       -> contrast: upper half moves the top,
  lower half the bottom, Shift moves both; double-click resets
* double-click a panel            -> show that image in the main viewer

The window talks to its data source (``gui/controllers/figures.py``
``FigureController``) only through a handful of methods; it never reaches
into the main window.
"""
from __future__ import annotations

import io
from pathlib import Path

from ..._shared import QtCore, QtGui, QtWidgets, log_status, np
from ...cmap_registry import get_cmap as registry_get_cmap
from ...figures import height_series as hs

_PAPER_MARGIN = 14       # px around the paper inside the view
_CLOSE_SIZE = 14         # px, hover × on column headers
_LIVE_POLL_MS = 1200     # source-change check while the window is visible
_DRAG_THROTTLE_MS = 30


def _figure_classes():
    from matplotlib.backends.backend_agg import FigureCanvasAgg
    from matplotlib.figure import Figure
    return Figure, FigureCanvasAgg


class _FigureView(QtWidgets.QWidget):
    """Paints the rendered figure centred on a white sheet; maps mouse
    positions back to figure inches for the window's hit-testing."""

    hovered = QtCore.pyqtSignal(object)            # hit dict or None
    pressed = QtCore.pyqtSignal(object, object)    # hit, QMouseEvent
    moved = QtCore.pyqtSignal(object)              # QMouseEvent (while pressed)
    released = QtCore.pyqtSignal(object)
    double_clicked = QtCore.pyqtSignal(object)
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
        self.hit_test = None   # callable(x_in, y_in) -> hit dict | None

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
        sx = rect.width() / self._fig_in[0]
        sy = rect.height() / self._fig_in[1]
        x0, y0, x1, y1 = box
        return QtCore.QRectF(rect.left() + x0 * sx, rect.top() + y0 * sy,
                             (x1 - x0) * sx, (y1 - y0) * sy)

    def close_rect(self, header_box):
        r = self.inches_rect(header_box)
        return QtCore.QRectF(r.right() - _CLOSE_SIZE - 2, r.top() + 3, _CLOSE_SIZE, _CLOSE_SIZE)

    # painting -----------------------------------------------------------
    def paintEvent(self, _event):
        p = QtGui.QPainter(self)
        p.setRenderHint(QtGui.QPainter.Antialiasing, True)
        p.fillRect(self.rect(), self.palette().window())
        if self._image is None:
            p.setPen(self.palette().color(QtGui.QPalette.Mid))
            p.drawText(self.rect(), QtCore.Qt.AlignCenter, "Rendering…")
            return
        rect = self._image_rect()
        paper = rect.adjusted(-8, -8, 8, 8)
        p.setPen(QtCore.Qt.NoPen)
        p.setBrush(QtGui.QColor(0, 0, 0, 40))
        p.drawRoundedRect(paper.translated(0, 2), 3, 3)
        p.setBrush(QtGui.QColor("white"))
        p.drawRoundedRect(paper, 3, 3)
        p.drawImage(rect, self._image)
        hit = self._hover
        if hit and self._layout:
            accent = QtGui.QColor(30, 110, 220)
            kind = hit.get("kind")
            if kind in ("header", "close"):
                box = self._layout["headers"][hit["col"]]
                r = self.inches_rect(box)
                p.setPen(QtGui.QPen(accent, 1.4))
                p.setBrush(QtCore.Qt.NoBrush)
                p.drawLine(QtCore.QPointF(r.left() + 6, r.bottom() - 3),
                           QtCore.QPointF(r.right() - 6, r.bottom() - 3))
                if hit.get("can_close"):
                    cr = self.close_rect(box)
                    p.setPen(QtCore.Qt.NoPen)
                    p.setBrush(QtGui.QColor(200, 50, 50) if kind == "close" else QtGui.QColor(120, 120, 120, 200))
                    p.drawEllipse(cr)
                    p.setPen(QtGui.QPen(QtGui.QColor("white"), 1.6))
                    m = 4.0
                    p.drawLine(QtCore.QPointF(cr.left() + m, cr.top() + m),
                               QtCore.QPointF(cr.right() - m, cr.bottom() - m))
                    p.drawLine(QtCore.QPointF(cr.right() - m, cr.top() + m),
                               QtCore.QPointF(cr.left() + m, cr.bottom() - m))
            elif kind == "row" and hit.get("can_cycle"):
                r = self.inches_rect(hit["box"])
                p.setPen(QtGui.QPen(accent, 1.2, QtCore.Qt.DashLine))
                p.setBrush(QtCore.Qt.NoBrush)
                p.drawRoundedRect(r.adjusted(4, r.height() * 0.35, -2, -r.height() * 0.35), 4, 4)
            elif kind == "colorbar":
                r = self.inches_rect(hit["box"])
                p.setPen(QtGui.QPen(accent, 1.4))
                p.setBrush(QtCore.Qt.NoBrush)
                p.drawRect(r.adjusted(-2, -2, 2, 2))

    # mouse --------------------------------------------------------------
    def _hit(self, pos):
        if self.hit_test is None:
            return None
        inch = self.to_inches(pos)
        if inch is None:
            return None
        hit = self.hit_test(*inch)
        if hit and hit.get("kind") == "header" and hit.get("can_close") and self._layout:
            if self.close_rect(self._layout["headers"][hit["col"]]).contains(QtCore.QPointF(pos)):
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

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.resized.emit()


class FigureWindow(QtWidgets.QWidget):
    """A live, data-aware figure (see module docstring)."""

    def __init__(self, source, fid, keys, *, mode="height", excluded=None, parent=None):
        super().__init__(parent, QtCore.Qt.Window)
        self.setAttribute(QtCore.Qt.WA_DeleteOnClose, True)
        self.source = source
        self.fid = fid
        self.mode = mode
        self._keys = [str(k) for k in keys]
        self._excluded = list(excluded or [])
        saved = source.get_state(fid)
        self.state = {
            "reference_key": saved.get("reference_key") or source.default_reference(self._keys),
            "dropped": list(saved.get("dropped") or []),
            "forced": list(saved.get("forced") or []),
            "row_choices": dict(saved.get("row_choices") or {}),
            "range_mode": saved.get("range_mode") or "row",
            "letters": bool(saved.get("letters", True)),
            "scale_bar": bool(saved.get("scale_bar", True)),
            "clims": dict(saved.get("clims") or {}),
        }
        self._undo, self._redo = [], []
        self._panels = {}           # (key, idx) -> panel dict
        self._members = None
        self._data_sig = None
        self.spec = None
        self.layout_info = None
        self._drag = None
        self._fig = None

        self._build_ui()
        self._init_shortcuts()
        self.resize(1100, 520)
        self._render_timer = QtCore.QTimer(self, singleShot=True, interval=90)
        self._render_timer.timeout.connect(self._render_screen)
        self._drag_timer = QtCore.QTimer(self, singleShot=True, interval=_DRAG_THROTTLE_MS)
        self._drag_timer.timeout.connect(self._redraw_fast)
        self._live_timer = QtCore.QTimer(self, interval=_LIVE_POLL_MS)
        self._live_timer.timeout.connect(self._check_live)
        self._live_timer.start()
        self._rebuild(reload_data=True)

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------
    def _build_ui(self):
        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(8, 6, 8, 8)
        root.setSpacing(6)
        strip = QtWidgets.QHBoxLayout()
        strip.setSpacing(6)

        def button(text, tip, slot, primary=False):
            b = QtWidgets.QToolButton(self)
            b.setText(text)
            b.setToolTip(tip)
            b.setToolButtonStyle(QtCore.Qt.ToolButtonTextOnly)
            b.setAutoRaise(False)
            b.setMinimumHeight(26)
            if primary:
                f = b.font()
                f.setBold(True)
                b.setFont(f)
            b.clicked.connect(slot)
            strip.addWidget(b)
            return b

        self.copy_btn = button("Copy", "Copy the figure (Ctrl+C) - paste into PowerPoint, Word, Inkscape...",
                               self.copy_to_clipboard, primary=True)
        self.ppt_btn = button("PowerPoint", "Send the figure to the open PowerPoint presentation as "
                              "editable vector shapes (new slide)", self.send_to_powerpoint)
        self.save_btn = button("Save PDF…", "Save as PDF, SVG or PNG (Ctrl+S)", self.save_as)
        strip.addSpacing(12)

        self.range_group = QtWidgets.QButtonGroup(self)
        self.range_group.setExclusive(True)
        for mode_key, text, tip in (
            ("row", "Shared range", "One colour range per row - differences between panels are real"),
            ("image", "Per image", "Each panel scaled on its own - shows contrast, not magnitude"),
        ):
            b = QtWidgets.QToolButton(self)
            b.setText(text)
            b.setToolTip(tip)
            b.setCheckable(True)
            b.setMinimumHeight(26)
            b.setProperty("mode_key", mode_key)
            self.range_group.addButton(b)
            strip.addWidget(b)
        self.range_group.buttonClicked.connect(
            lambda b: self._apply_edit({"range_mode": b.property("mode_key")}))

        strip.addSpacing(8)
        self.letters_chk = QtWidgets.QCheckBox("Letters", self)
        self.letters_chk.toggled.connect(lambda on: self._apply_edit({"letters": bool(on)}))
        strip.addWidget(self.letters_chk)
        self.scale_chk = QtWidgets.QCheckBox("Scale bar", self)
        self.scale_chk.toggled.connect(lambda on: self._apply_edit({"scale_bar": bool(on)}))
        strip.addWidget(self.scale_chk)
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
        more.addAction("Reset figure", self.reset_figure)
        self.more_btn.setMenu(more)
        strip.addWidget(self.more_btn)
        root.addLayout(strip)

        self.view = _FigureView(self)
        self.view.hit_test = self._hit_test
        self.view.hovered.connect(self._on_hover)
        self.view.pressed.connect(self._on_press)
        self.view.moved.connect(self._on_drag_move)
        self.view.released.connect(self._on_release)
        self.view.double_clicked.connect(self._on_double_click)
        self.view.resized.connect(lambda: self._render_timer.start())
        root.addWidget(self.view, 1)

        self.hint = QtWidgets.QLabel(self)
        self.hint.setStyleSheet("color: palette(mid); font-size: 8pt;")
        self.hint.setText("Click a Δz header to make it the reference  ·  hover a header and click × to "
                          "drop a frame  ·  click I / Δf to switch channel  ·  drag a colourbar for "
                          "contrast  ·  Ctrl+Z undo")
        root.addWidget(self.hint)
        self._sync_controls()

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

    def _sync_controls(self):
        for b in self.range_group.buttons():
            b.blockSignals(True)
            b.setChecked(b.property("mode_key") == self.state["range_mode"])
            b.blockSignals(False)
        for chk, key in ((self.letters_chk, "letters"), (self.scale_chk, "scale_bar")):
            chk.blockSignals(True)
            chk.setChecked(bool(self.state[key]))
            chk.blockSignals(False)

    # ------------------------------------------------------------------
    # State / undo
    # ------------------------------------------------------------------
    def _snapshot(self):
        return {k: (list(v) if isinstance(v, list) else dict(v) if isinstance(v, dict) else v)
                for k, v in self.state.items()}

    def _apply_edit(self, changes, *, push=True, reload_data=False):
        new = self._snapshot()
        new.update(changes)
        if new == self.state:
            return
        if push:
            self._undo.append(self._snapshot())
            del self._undo[:-100]
            self._redo.clear()
        self.state = new
        self._commit(reload_data=reload_data)

    def _commit(self, reload_data=False):
        self.source.set_state(self.fid, self.state)
        self._sync_controls()
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
            "dropped": [], "forced": [], "row_choices": {}, "range_mode": "row",
            "letters": True, "scale_bar": True, "clims": {},
        })

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
        self.spec = hs.build_spec(
            self._members, mode=self.mode,
            reference_key=self.state["reference_key"],
            dropped=self.state["dropped"],
            row_choices=self.state["row_choices"])
        self._update_title()
        self._render_screen()
        self._update_note()

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

    def _row_clims(self):
        out = {}
        if self.state["range_mode"] != "row":
            return out
        for r, row in enumerate(self.spec["rows"]):
            lim = self.state["clims"].get(row["kind"])
            if lim and len(lim) == 2:
                out[r] = (float(lim[0]), float(lim[1]))
        return out

    def _update_title(self):
        cols = self.spec["columns"] if self.spec else []
        if self.mode == "height":
            what = "Height series"
        else:
            what = "Figure"
        names = [c["name"] for c in cols]
        span = f"{names[0]} … {names[-1]}" if len(names) > 1 else (names[0] if names else "")
        self.setWindowTitle(f"{what} - {span} ({len(cols)} frame{'s' if len(cols) != 1 else ''})")

    def _update_note(self):
        left_out = [e for e in self._excluded if e["key"] not in self.state["forced"]]
        dropped = [k for k in self.state["dropped"] if k in self._all_keys()]
        menu = QtWidgets.QMenu(self.note_btn)
        parts = []
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
        forced = [k for k in self.state["forced"]]
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
            fig, self.spec, self._panel_map(),
            range_mode=self.state["range_mode"], clims=self._row_clims(),
            letters=self.state["letters"], scale_bar=self.state["scale_bar"],
            get_cmap=lambda name: registry_get_cmap(name, "gray"))

    def _render_screen(self):
        if not self.spec or not self.spec["columns"] or not self.spec["rows"]:
            self.view.set_image(None, (1, 1), None)
            self.layout_info = None
            return
        if self._fig is None:
            self._fig = self._new_figure()
        self.layout_info = self._draw_into(self._fig)
        self._blit()

    def _blit(self):
        fig = self._fig
        fig_w, fig_h = self.layout_info["size"]
        avail_w, avail_h = self.view.available_px()
        avail_w -= 16
        avail_h -= 16
        dpr = float(self.devicePixelRatioF() or 1.0)
        dpi = max(30.0, min(avail_w / fig_w, avail_h / fig_h, 400.0))
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
    # Hit-testing + in-place edits
    # ------------------------------------------------------------------
    def _hit_test(self, x, y):
        lay = self.layout_info
        if not lay or not self.spec:
            return None

        def inside(box):
            return box[0] <= x <= box[2] and box[1] <= y <= box[3]

        n_cols = len(self.spec["columns"])
        for c, box in enumerate(lay["headers"]):
            if inside(box):
                return {"kind": "header", "col": c, "can_close": n_cols > 1}
        for r, box in lay["row_labels"]:
            if inside(box):
                return {"kind": "row", "row": r, "box": box,
                        "can_cycle": len(self.spec["rows"][r]["alternatives"]) > 1}
        for r, box in lay["colorbars"]:
            x0, y0, x1, y1 = box
            if x0 - 0.05 <= x <= x1 + 0.05 and y0 <= y <= y1:
                return {"kind": "colorbar", "row": r, "box": box, "y0": y0, "y1": y1}
        for (r, c), box in lay["cells"].items():
            if inside(box):
                return {"kind": "panel", "row": r, "col": c}
        return None

    def _on_hover(self, hit):
        cursor = QtCore.Qt.ArrowCursor
        tip = ""
        if hit:
            kind = hit["kind"]
            if kind == "close":
                cursor = QtCore.Qt.PointingHandCursor
                tip = "Remove this frame from the figure (Ctrl+Z restores)"
            elif kind == "header":
                cursor = QtCore.Qt.PointingHandCursor
                col = self.spec["columns"][hit["col"]]
                tip = (f"{col['name']}\nThis frame is the Δz reference" if col["is_ref"]
                       else f"{col['name']}\nClick: make this the Δz = 0 reference"
                       if self.spec["has_z"] else col["name"])
            elif kind == "row" and hit.get("can_cycle"):
                cursor = QtCore.Qt.PointingHandCursor
                row = self.spec["rows"][hit["row"]]
                alts = row["alternatives"]
                nxt = alts[(alts.index(row["choice"]) + 1) % len(alts)] if row["choice"] in alts else alts[0]
                tip = f"Channel: {row['choice']}\nClick: switch to {nxt}"
            elif kind == "colorbar":
                cursor = QtCore.Qt.SizeVerCursor
                tip = ("Drag: upper half moves the top, lower half the bottom, Shift moves both\n"
                       "Double-click: automatic range")
            elif kind == "panel":
                col = self.spec["columns"][hit["col"]]
                z = col.get("z_nm")
                tip = col["name"] + (f"\nz_abs {z:.3f} nm" if z is not None else "") + \
                    "\nDouble-click: show in the viewer"
        self.view.setCursor(cursor)
        self.view.setToolTip(tip)

    def _on_press(self, hit, event):
        if not hit:
            return
        kind = hit["kind"]
        if kind == "close":
            col = self.spec["columns"][hit["col"]]
            self._apply_edit({"dropped": self.state["dropped"] + [col["key"]]})
            self._flash(f"Removed {col['name']} - Ctrl+Z to restore")
        elif kind == "header" and self.spec["has_z"]:
            col = self.spec["columns"][hit["col"]]
            if not col["is_ref"]:
                self._apply_edit({"reference_key": col["key"]})
        elif kind == "row" and hit.get("can_cycle"):
            row = self.spec["rows"][hit["row"]]
            alts = row["alternatives"]
            nxt = alts[(alts.index(row["choice"]) + 1) % len(alts)] if row["choice"] in alts else alts[0]
            choices = dict(self.state["row_choices"])
            choices[row["kind"]] = nxt
            clims = dict(self.state["clims"])
            clims.pop(row["kind"], None)   # a different channel needs its own range
            self._apply_edit({"row_choices": choices, "clims": clims})
            self._flash(f"{row['label']} row: {nxt}")
        elif kind == "colorbar" and self.state["range_mode"] == "row":
            r = hit["row"]
            lo, hi = self.layout_info["ranges"].get(r, (0.0, 1.0))
            inch = self.view.to_inches(event.pos())
            mid = 0.5 * (hit["y0"] + hit["y1"])
            end = "both" if event.modifiers() & QtCore.Qt.ShiftModifier else (
                "hi" if inch and inch[1] < mid else "lo")
            self._drag = {"row": r, "start_y": event.pos().y(), "lo": lo, "hi": hi, "end": end,
                          "px": max(20.0, self.view.inches_rect(hit["box"]).height()),
                          "before": self._snapshot()}

    def _on_drag_move(self, event):
        d = self._drag
        if not d:
            return
        span = d["hi"] - d["lo"]
        delta = -(event.pos().y() - d["start_y"]) / d["px"] * span
        lo, hi = d["lo"], d["hi"]
        if d["end"] == "hi":
            hi = max(lo + 1e-6 * abs(span or 1.0), hi + delta)
        elif d["end"] == "lo":
            lo = min(hi - 1e-6 * abs(span or 1.0), lo + delta)
        else:
            lo, hi = lo + delta, hi + delta
        d["current"] = (lo, hi)
        hs.set_row_range(self.layout_info, d["row"], lo, hi)
        if not self._drag_timer.isActive():
            self._drag_timer.start()

    def _on_release(self, _event):
        d, self._drag = self._drag, None
        if not d or "current" not in d:
            return
        kind = self.spec["rows"][d["row"]]["kind"]
        clims = dict(self.state["clims"])
        clims[kind] = [float(d["current"][0]), float(d["current"][1])]
        self._undo.append(d["before"])
        self._redo.clear()
        self.state = dict(self._snapshot(), clims=clims)
        self.source.set_state(self.fid, self.state)
        self._redraw_fast()

    def _show_key(self, key):
        for row in (self.spec or {}).get("rows", []):
            idx = row["cells"].get(key)
            if idx is not None:
                self.source.show_image(key, idx)
                return

    def _on_double_click(self, hit):
        if not hit:
            return
        if hit["kind"] == "colorbar":
            kind = self.spec["rows"][hit["row"]]["kind"]
            if kind in self.state["clims"]:
                clims = dict(self.state["clims"])
                clims.pop(kind, None)
                self._apply_edit({"clims": clims})
        elif hit["kind"] == "panel":
            col = self.spec["columns"][hit["col"]]
            row = self.spec["rows"][hit["row"]]
            idx = row["cells"].get(col["key"])
            if idx is not None:
                self.source.show_image(col["key"], idx)

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
        suffix = "height_series" if self.mode == "height" else "figure"
        return f"{stem}_{suffix}"

    def _svg_bytes(self, fig=None):
        import matplotlib
        fig = fig or self._export_figure()
        buf = io.BytesIO()
        with matplotlib.rc_context({"svg.fonttype": "none"}):
            fig.savefig(buf, format="svg", facecolor="white")
        return buf.getvalue()

    def copy_to_clipboard(self):
        if not self.spec or not self.spec["columns"]:
            return
        fig = self._export_figure()
        buf = io.BytesIO()
        fig.savefig(buf, format="png", dpi=300, facecolor="white")
        img = QtGui.QImage.fromData(buf.getvalue(), "PNG")
        mime = QtCore.QMimeData()
        mime.setImageData(img)
        try:
            mime.setData("image/svg+xml", QtCore.QByteArray(self._svg_bytes(fig)))
        except Exception:
            pass
        QtWidgets.QApplication.clipboard().setMimeData(mime)
        self._flash("Copied - paste anywhere")

    def send_to_powerpoint(self):
        if not self.spec or not self.spec["columns"]:
            return
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
            slide, _name = ppt_bridge.send_svg_to_ppt(self._svg_bytes(), label=None, new_slide=True)
        except Exception as exc:
            QtWidgets.QApplication.restoreOverrideCursor()
            QtWidgets.QMessageBox.warning(
                self, "PowerPoint",
                f"Could not send to PowerPoint:\n{exc}\n\nIs a presentation open?")
            return
        QtWidgets.QApplication.restoreOverrideCursor()
        self._flash(f"Sent to PowerPoint (slide {slide})")

    def save_as(self):
        if not self.spec or not self.spec["columns"]:
            return
        start = str(Path(self.source.default_save_dir(self._keys)) / (self._default_name() + ".pdf"))
        path, _flt = QtWidgets.QFileDialog.getSaveFileName(
            self, "Save figure", start, "PDF (*.pdf);;SVG (*.svg);;PNG image (*.png)")
        if not path:
            return
        fig = self._export_figure()
        ext = Path(path).suffix.lower()
        try:
            if ext == ".svg":
                Path(path).write_bytes(self._svg_bytes(fig))
            elif ext == ".png":
                fig.savefig(path, dpi=300, facecolor="white")
            else:
                if ext != ".pdf":
                    path = path + ".pdf"
                fig.savefig(path, facecolor="white")
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Save figure", f"Could not save:\n{exc}")
            return
        self._flash(f"Saved {Path(path).name}")

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
