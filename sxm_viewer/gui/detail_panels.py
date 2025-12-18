"""Detail canvases and spectroscopy dialogs."""
from __future__ import annotations

from .._shared import *
from ..config import *
from ..data.io import *
from ..data.spectroscopy import *
from .thumbnails import *


class MultiPreviewCanvas(FigureCanvas):
    def __init__(self, parent=None, figsize=(6,6)):
        self.fig = Figure(figsize=figsize)
        super().__init__(self.fig)
        if parent is not None:
            self.setParent(parent)
        self.views = []
        self._ax_view_map = {}
        self._copy_feedback_handler = None
        self._drag_candidate = None  # (view, QPoint start, QImage cache)
        self._value_callback = None
        self._value_cid = self.mpl_connect('motion_notify_event', self._on_motion_value)
        # profile (interactive line) state
        self.profile_enabled = False
        self.profile_pts = None  # (x0, y0, x1, y1) in data coords of main ax
        self._profile_line = None
        self._profile_p0 = None
        self._profile_p1 = None
        self._cids = []
        self._base_click_cid = self.mpl_connect('button_press_event', self._on_base_click)
        self._dragging = None  # 'p0' or 'p1'
        self.main_ax = None
        self.profile_callback = None  # callable(x_px, vals, length_nm)

    def set_views(self, views):
        self.views = views[:]
        self._redraw()

    def clear_views(self):
        self.views = []
        self._redraw()

    def resizeEvent(self, event):
        size = event.size()
        if size.width() <= 0 or size.height() <= 0:
            safe = QtCore.QSize(max(1, size.width()), max(1, size.height()))
            safe_event = QtGui.QResizeEvent(safe, event.oldSize())
            super().resizeEvent(safe_event)
            return
        super().resizeEvent(event)

    def _redraw(self):
        self.fig.clf()
        self._ax_view_map = {}
        n = len(self.views)
        if n == 0:
            self.draw(); return
        cols = int(math.ceil(math.sqrt(n)))
        rows = int(math.ceil(n / cols))
        for i, v in enumerate(self.views):
            ax = self.fig.add_subplot(rows, cols, i+1)
            self._ax_view_map[ax] = v
            if i == 0:
                self.main_ax = ax
            arr = np.asarray(v['arr'])
            extent = v.get('extent', None)
            cmap = v.get('cmap', 'viridis')
            if extent is None:
                im = ax.imshow(arr, origin='upper', interpolation='nearest', cmap=cmap)
            else:
                im = ax.imshow(arr, extent=extent, origin='upper', interpolation='nearest', aspect='equal', cmap=cmap)
            unit = v.get('unit', '')
            if unit:
                cbar = self.fig.colorbar(im, ax=ax, fraction=0.08, pad=0.02)
                cbar.set_label(unit)
            title = v.get('title', '')
            ax.set_title(title, fontsize=9)
            ax.tick_params(labelsize=8)
        try: self.fig.tight_layout()
        except Exception: pass
        # if profile mode is enabled, (re)create artists on main ax
        if self.profile_enabled:
            self._ensure_profile_artists()
        self.draw()

    # ---------- Interactive profile helpers ----------
    def set_profile_callback(self, cb):
        self.profile_callback = cb

    def set_copy_feedback_handler(self, handler):
        self._copy_feedback_handler = handler

    def set_value_callback(self, cb):
        self._value_callback = cb

    def enable_profile(self, enable:bool):
        if enable == self.profile_enabled:
            return
        self.profile_enabled = enable
        if enable:
            self._connect_profile_events()
            self._ensure_profile_artists()
        else:
            self._disconnect_profile_events()
            self._clear_profile_artists()
            self.profile_pts = None
        self.draw_idle()

    def _connect_profile_events(self):
        if self._cids:
            return
        self._cids = [
            self.mpl_connect('button_press_event', self._on_press),
            self.mpl_connect('button_release_event', self._on_release),
            self.mpl_connect('motion_notify_event', self._on_motion),
        ]

    def _disconnect_profile_events(self):
        for cid in self._cids:
            try: self.mpl_disconnect(cid)
            except Exception: pass
        self._cids = []

    def _ensure_profile_artists(self):
        if self.main_ax is None:
            return
        if self._profile_line is None:
            # initialize points centered if not set
            if self.profile_pts is None:
                try:
                    v0 = self.views[0]
                    arr = np.asarray(v0['arr'])
                    h, w = arr.shape
                    if v0.get('extent') is None:
                        x0 = w*0.25; y0 = h*0.5; x1 = w*0.75; y1 = h*0.5
                    else:
                        xmin, xmax, ymin, ymax = v0['extent'][0], v0['extent'][1], v0['extent'][2], v0['extent'][3]
                        # our code sets extent as [0, XRange, YRange, 0], so choose a centered horizontal line
                        x0 = xmin + 0.25*(xmax - xmin); x1 = xmin + 0.75*(xmax - xmin)
                        y0 = ymax + 0.5*(ymin - ymax); y1 = y0
                except Exception:
                    x0 = 0.25; x1 = 0.75; y0 = y1 = 0.5
                self.profile_pts = (x0, y0, x1, y1)
            x0, y0, x1, y1 = self.profile_pts
            self._profile_line, = self.main_ax.plot([x0,x1],[y0,y1], color='yellow', lw=2, alpha=0.9, zorder=9)
            self._profile_p0, = self.main_ax.plot([x0],[y0], marker='o', color='yellow', ms=7, mec='black', mew=1.0, zorder=10)
            self._profile_p1, = self.main_ax.plot([x1],[y1], marker='o', color='yellow', ms=7, mec='black', mew=1.0, zorder=10)

    def _clear_profile_artists(self):
        for art in (self._profile_line, self._profile_p0, self._profile_p1):
            try:
                if art is not None:
                    art.remove()
            except Exception:
                pass
        self._profile_line = self._profile_p0 = self._profile_p1 = None
        self.draw_idle()

    def _update_profile_artists(self):
        if self._profile_line is None or self._profile_p0 is None or self._profile_p1 is None:
            return
        x0, y0, x1, y1 = self.profile_pts
        self._profile_line.set_data([x0,x1],[y0,y1])
        self._profile_p0.set_data([x0],[y0])
        self._profile_p1.set_data([x1],[y1])
        self.draw_idle()
        self._emit_profile()

    def _pt_distance_pixels(self, x, y, xp, yp):
        try:
            p_scr = self.main_ax.transData.transform((x, y))
            q_scr = self.main_ax.transData.transform((xp, yp))
            dx = p_scr[0] - q_scr[0]; dy = p_scr[1] - q_scr[1]
            return (dx*dx + dy*dy) ** 0.5
        except Exception:
            return float('inf')

    def _on_press(self, event):
        if not self.profile_enabled or event.inaxes is None or event.inaxes is not self.main_ax:
            return
        if event.button != 1:
            return
        x, y = event.xdata, event.ydata
        if x is None or y is None:
            return
        if self.profile_pts is None:
            self.profile_pts = (x, y, x, y)
            self._ensure_profile_artists()
            self._dragging = 'p1'
            self._update_profile_artists()
            return
        x0, y0, x1, y1 = self.profile_pts
        d0 = self._pt_distance_pixels(x, y, x0, y0)
        d1 = self._pt_distance_pixels(x, y, x1, y1)
        thresh = 10.0  # pixels
        if d0 <= thresh or d0 <= d1:
            if d0 <= thresh:
                self._dragging = 'p0'
                return
        if d1 <= thresh:
            self._dragging = 'p1'
            return
        # else: start a new line from here
        self.profile_pts = (x, y, x, y)
        self._dragging = 'p1'
        self._update_profile_artists()

    def _on_motion(self, event):
        if not self.profile_enabled or self._dragging is None or event.inaxes is None or event.inaxes is not self.main_ax:
            return
        x, y = event.xdata, event.ydata
        if x is None or y is None:
            return
        x0, y0, x1, y1 = self.profile_pts
        if self._dragging == 'p0':
            self.profile_pts = (x, y, x1, y1)
        elif self._dragging == 'p1':
            self.profile_pts = (x0, y0, x, y)
        self._update_profile_artists()

    def _on_release(self, event):
        if not self.profile_enabled:
            return
        self._dragging = None

    def _emit_profile(self):
        if not callable(self.profile_callback):
            return
        try:
            v0 = self.views[0]
            arr = np.asarray(v0['arr'], dtype=float)
            h, w = arr.shape
            extent = v0.get('extent', None)
            x0, y0, x1, y1 = self.profile_pts
            # map data coords to pixel indices
            if extent is None:
                c0 = x0; r0 = y0; c1 = x1; r1 = y1
                length_nm = None
            else:
                xmin, xmax = extent[0], extent[1]
                ymin, ymax = extent[2], extent[3]
                # our extent is [0, XRange, YRange, 0] so use ranges directly
                xr = (xmax - xmin) if (xmax is not None and xmin is not None) else 1.0
                yr = (ymin - ymax) if (ymin is not None and ymax is not None) else 1.0  # note inverted
                c0 = (x0 - xmin) / (xr + 1e-12) * (w - 1)
                c1 = (x1 - xmin) / (xr + 1e-12) * (w - 1)
                # y increases downward in array index
                # since extent top=ymax (0) and bottom=ymin (YRange), map linearly
                r0 = (y0 - ymax) / (ymin - ymax + 1e-12) * (h - 1)
                r1 = (y1 - ymax) / (ymin - ymax + 1e-12) * (h - 1)
                # physical length in nm using data coords directly
                try:
                    dx_nm = (x1 - x0); dy_nm = (y1 - y0)
                    length_nm = float((dx_nm*dx_nm + dy_nm*dy_nm) ** 0.5)
                except Exception:
                    length_nm = None
            # sample along the line using bilinear interpolation
            n = int(max(2, round(((c1 - c0)**2 + (r1 - r0)**2) ** 0.5) + 1))
            t = np.linspace(0.0, 1.0, n)
            cc = c0 + (c1 - c0) * t
            rr = r0 + (r1 - r0) * t
            rr = np.clip(rr, 0, h - 1)
            cc = np.clip(cc, 0, w - 1)
            i0 = np.floor(rr).astype(int)
            j0 = np.floor(cc).astype(int)
            i1 = np.clip(i0 + 1, 0, h - 1)
            j1 = np.clip(j0 + 1, 0, w - 1)
            wy = rr - i0
            wx = cc - j0
            vals = (
                (1 - wy) * (1 - wx) * arr[i0, j0] +
                wy * (1 - wx) * arr[i1, j0] +
                (1 - wy) * wx * arr[i0, j1] +
                wy * wx * arr[i1, j1]
            )
            x_px = np.linspace(0.0, float(n - 1), n)
            unit = v0.get('unit', None)
            self.profile_callback(x_px, vals, length_nm, unit)
        except Exception:
            pass

    def _on_base_click(self, event):
        if event is None or event.inaxes is None:
            return
        ax = event.inaxes
        view = self._ax_view_map.get(ax)
        if event.button == 3:
            self._show_context_menu(event, view)
            return
        if event.button != 1:
            return
        if getattr(event, 'dblclick', False):
            if view:
                self._copy_view_to_clipboard(view)
            return
        if view and getattr(event, 'guiEvent', None) is not None:
            pos = event.guiEvent.globalPos()
            self._drag_candidate = {'view': view, 'start': QtCore.QPoint(pos), 'image': None}

    def _copy_view_to_clipboard(self, view):
        try:
            qimg = self._view_to_qimage(view)
            QtWidgets.QApplication.clipboard().setImage(qimg)
            if callable(self._copy_feedback_handler):
                self._copy_feedback_handler(view)
        except Exception:
            pass

    def _view_to_qimage(self, view):
        arr = np.asarray(view.get('arr'))
        cmap = view.get('cmap', 'viridis')
        return array_to_qimage(arr, cmap_name=cmap)

    def _show_context_menu(self, event, view):
        if view is None or getattr(event, 'guiEvent', None) is None:
            return
        menu = QtWidgets.QMenu(self)
        copy_act = menu.addAction("Copy image")
        save_act = menu.addAction("Save image as...")
        chosen = menu.exec_(event.guiEvent.globalPos())
        if chosen == copy_act:
            self._copy_view_to_clipboard(view)
        elif chosen == save_act:
            self._save_view_to_file(view)

    def _save_view_to_file(self, view):
        try:
            title = view.get('title') or 'view'
            default = f"{title}.png"
            path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save view", default, "PNG Files (*.png)")
            if not path:
                return
            qimg = self._view_to_qimage(view)
            qimg.save(path, "PNG")
        except Exception:
            QtWidgets.QMessageBox.warning(self, "Save view", "Unable to save image.")

    def _start_drag(self, view, qimg=None):
        try:
            if qimg is None:
                qimg = self._view_to_qimage(view)
            pix = QtGui.QPixmap.fromImage(qimg)
            drag = QtGui.QDrag(self)
            mime = QtCore.QMimeData()
            mime.setImageData(qimg)
            drag.setMimeData(mime)
            drag.setPixmap(pix.scaled(128, 128, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation))
            drag.exec_(QtCore.Qt.CopyAction)
        except Exception:
            pass

    def mouseMoveEvent(self, event):
        if self._drag_candidate:
            start = self._drag_candidate.get('start')
            if start is not None:
                if (event.globalPos() - start).manhattanLength() >= 10:
                    view = self._drag_candidate.get('view')
                    qimg = self._drag_candidate.get('image')
                    if qimg is None and view is not None:
                        qimg = self._view_to_qimage(view)
                        self._drag_candidate['image'] = qimg
                    if view is not None and qimg is not None:
                        self._start_drag(view, qimg)
                    self._drag_candidate = None
                    super().mouseMoveEvent(event)
                    return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        self._drag_candidate = None
        super().mouseReleaseEvent(event)

    def _on_motion_value(self, event):
        if self._value_callback is None:
            return
        if event.inaxes is None or event.inaxes not in self._ax_view_map:
            self._value_callback(None, None, None, None)
            return
        view = self._ax_view_map.get(event.inaxes)
        if view is None:
            self._value_callback(None, None, None, None)
            return
        val = sample_array_value(view.get('arr'), event.xdata, event.ydata, view.get('extent'))
        if val is None:
            self._value_callback(None, event.xdata, event.ydata, view)
        else:
            self._value_callback(val, event.xdata, event.ydata, view)

class ProfileDialog(QtWidgets.QDialog):
    """Dialog showing the sampled profile and basic stats."""
    def __init__(self, x, vals, length_nm=None, parent=None, unit=None):
        super().__init__(parent)
        self.setWindowTitle('Profile measurement')
        self.resize(600, 320)
        self._unit = unit
        v = QtWidgets.QVBoxLayout()
        # matplotlib canvas for plot
        fig = Figure(figsize=(6,3))
        self.canvas = FigureCanvas(fig)
        self.ax = fig.add_subplot(111)
        (self._line,) = self.ax.plot(x, vals)
        self.ax.set_xlabel('Distance (px)')
        self.ax.set_ylabel(f"Value ({unit})" if unit else 'Value')
        self.ax_top = self.ax.twiny()
        self.ax_top.set_xlabel('Distance (nm)')
        self._set_top_axis(length_nm, len(x))
        v.addWidget(self.canvas)
        # stats area
        self.stats = QtWidgets.QLabel(self._fmt_length(length_nm))
        v.addWidget(self.stats)
        btn = QtWidgets.QPushButton('Close')
        btn.clicked.connect(self.accept)
        v.addWidget(btn, alignment=QtCore.Qt.AlignRight)
        self.setLayout(v)

    def _fmt_length(self, length_nm):
        return f"Length: {length_nm:.3f} nm" if length_nm is not None else "Length: N/A"

    def _set_top_axis(self, length_nm, n_pts):
        try:
            self.ax_top.set_xlim(self.ax.get_xlim())
            if length_nm is None or n_pts <= 1:
                self.ax_top.set_xticks([])
                self.ax_top.set_xticklabels([])
            else:
                ticks = self.ax.get_xticks()
                scale = float(length_nm) / float(n_pts - 1)
                self.ax_top.set_xticks(ticks)
                self.ax_top.set_xticklabels([f"{t*scale:.1f}" for t in ticks])
        except Exception:
            pass

    def update_data(self, x, vals, length_nm=None):
        try:
            self._line.set_data(x, vals)
            self.ax.relim(); self.ax.autoscale_view()
            self._set_top_axis(length_nm, len(x))
            self.stats.setText(self._fmt_length(length_nm))
            self.canvas.draw_idle()
        except Exception:
            pass

class SpectroscopyPopup(QtWidgets.QDialog):
    """Popup window showing spectroscopy curves for a given file."""
    def __init__(self, spec, parent=None):
        super().__init__(parent)
        self.spec = spec
        self.setWindowTitle(f"Spectroscopy: {Path(spec['path']).name}")
        self.resize(720, 520)
        layout = QtWidgets.QVBoxLayout()
        meta_txt = f"File: {Path(spec['path']).name}\nPosition: {spec.get('x','?')}/{spec.get('y','?')} nm\nTime: {spec.get('time')}"
        self.meta_label = QtWidgets.QLabel(meta_txt)
        self.meta_label.setWordWrap(True)
        layout.addWidget(self.meta_label)

        selector_layout = QtWidgets.QHBoxLayout()
        selector_layout.addWidget(QtWidgets.QLabel("Channel:"))
        self.channel_combo = QtWidgets.QComboBox()
        selector_layout.addWidget(self.channel_combo, 1)
        self.fit_btn = QtWidgets.QPushButton("Fit parabola")
        self.copy_btn = QtWidgets.QPushButton("Copy channel")
        selector_layout.addWidget(self.fit_btn)
        selector_layout.addWidget(self.copy_btn)
        layout.addLayout(selector_layout)

        self.fig = Figure(figsize=(6,4))
        self.canvas = FigureCanvas(self.fig)
        self.ax = self.fig.add_subplot(111)
        layout.addWidget(self.canvas, 1)
        self.canvas.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.canvas.customContextMenuRequested.connect(self._on_canvas_context_menu)
        self.fit_result_label = QtWidgets.QLabel("")
        self.fit_result_label.setWordWrap(True)
        layout.addWidget(self.fit_result_label)
        self.setLayout(layout)

        self.V = np.asarray(spec.get('V', []), dtype=float)
        self.channels = {name: np.asarray(vals, dtype=float) for name, vals in (spec.get('channels', {}) or {}).items()}
        for name in self.channels.keys():
            self.channel_combo.addItem(name)
        self.channel_combo.currentTextChanged.connect(self._on_channel_changed)
        self.fit_btn.clicked.connect(self._on_fit_clicked)
        self.copy_btn.clicked.connect(self._copy_channel_to_clipboard)
        self._last_fit_result = None
        if self.channel_combo.count():
            self.channel_combo.setCurrentIndex(0)
            self._plot_selected_channel()
        else:
            self.ax.text(0.5, 0.5, "No channels", ha='center', va='center', transform=self.ax.transAxes)
            self.canvas.draw()
        self._update_fit_button()

    def _channel_label_with_unit(self, name):
        base = name or ""
        unit = self.spec.get('unit_map', {}).get(name)
        if not unit and '(' in base and base.endswith(')'):
            return base
        if unit:
            return f"{base} ({unit})"
        return base

    def _on_channel_changed(self, name):
        self._last_fit_result = None
        self.fit_result_label.setText("")
        self._plot_selected_channel()
        self._update_fit_button()

    def _plot_selected_channel(self):
        self.ax.clear()
        name = self.channel_combo.currentText()
        if not name or name not in self.channels or not self.V.size:
            self.canvas.draw_idle()
            return
        bias_mv = self.V * 1000.0
        self.ax.plot(bias_mv, self.channels[name], color='#c94cfa', lw=1.5, label='Data')
        self.ax.set_xlabel("Bias (mV)")
        self.ax.set_ylabel(self._channel_label_with_unit(name))
        self.ax.grid(True, alpha=0.2)
        if self._last_fit_result and self._last_fit_result.get('channel') == name:
            self._draw_fit_overlay(self._last_fit_result)
        else:
            handles, labels = self.ax.get_legend_handles_labels()
            if handles:
                self.ax.legend()
        self.canvas.draw_idle()

    def _on_canvas_context_menu(self, pos):
        menu = QtWidgets.QMenu(self)
        copy_act = menu.addAction("Copy channel data")
        action = menu.exec_(self.canvas.mapToGlobal(pos))
        if action == copy_act:
            self._copy_channel_to_clipboard()

    def _copy_channel_to_clipboard(self):
        name = self.channel_combo.currentText()
        if not name or name not in self.channels or not self.V.size:
            QtWidgets.QMessageBox.information(self, "Copy spectroscopy", "No spectroscopy data to copy.")
            return
        bias = self.V
        values = self.channels[name]
        spec_path = Path(self.spec.get('path', ''))
        file_name = spec_path.name or 'unknown'
        folder_name = spec_path.parent.name if spec_path.parent != spec_path else ''
        pos = (self.spec.get('x'), self.spec.get('y'))
        time_str = self.spec.get('time')
        lines = [
            f"File\t{file_name}",
            f"Channel\t{name}",
            f"Position (nm)\t{pos[0] if pos[0] is not None else '?'}\t{pos[1] if pos[1] is not None else '?'}",
            f"Folder\t{folder_name}",
            f"Acquired\t{time_str}",
            "",
            f"Bias (mV)\t{self._channel_label_with_unit(name)}"
        ]
        for v, val in zip(bias, values):
            try:
                lines.append(f"{float(v) * 1000.0:.9g}\t{float(val):.9g}")
            except Exception:
                lines.append(f"{v}\t{val}")
        QtWidgets.QApplication.clipboard().setText("\n".join(lines))
        QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), "Spectroscopy copied", self)

    def _draw_fit_overlay(self, res):
        if not self.V.size:
            return
        x_dense = np.linspace(np.nanmin(self.V), np.nanmax(self.V), 400)
        y_dense = res['func'](x_dense)
        self.ax.plot(x_dense * 1000.0, y_dense, '--', color='#ff8c00', lw=1.5, label='Fit')
        b = res['b']; c = res['c']; b_err = res.get('b_err', 0.0)
        self.ax.errorbar([b * 1000.0], [c], xerr=[b_err * 1000.0], fmt='o', color='#004c99', ecolor='#004c99', capsize=4, label='LCPD')
        self.ax.legend()
        text = (
            f"a = {res['a']:.4g} +/- {res['a_err']:.2g}\n"
            f"b (LCPD) = {res['b']:.2f} +/- {res['b_err']:.2f} mV\n"
            f"c = {res['c']:.4g} +/- {res['c_err']:.2g} Hz\n"
            f"RMSE = {res['rmse']:.4g}"
        )
        self.fit_result_label.setText(text)

    def _on_fit_clicked(self):
        name = self.channel_combo.currentText()
        if not name or name not in self.channels:
            return
        try:
            res = fit_parabola_bias(self.V, self.channels[name])
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Fit failed", str(e))
            return
        res['channel'] = name
        self._last_fit_result = res
        self._plot_selected_channel()

    def _update_fit_button(self):
        enable = bool(self.channel_combo.count() and self.V.size)
        self.fit_btn.setEnabled(enable)


class MatrixSpectroViewer(QtWidgets.QDialog):
    def __init__(self, parent, image_entry, specs):
        super().__init__(parent)
        self.image_entry = image_entry
        self.specs = list(specs)
        self.viewer = parent
        self.setWindowTitle(f"Matrix Spectroscopies - {Path(image_entry['path']).name}")
        self.resize(900, 700)
        layout = QtWidgets.QVBoxLayout()
        self.canvas = FigureCanvas(Figure(figsize=(6,6)))
        layout.addWidget(self.canvas, 1)
        self.ax = self.canvas.figure.add_subplot(111)
        self.image_value_label = QtWidgets.QLabel("Value: --")
        layout.addWidget(self.image_value_label)
        controls = QtWidgets.QHBoxLayout()
        controls.addWidget(QtWidgets.QLabel("Image channel:"))
        self.channel_combo = QtWidgets.QComboBox()
        controls.addWidget(self.channel_combo, 1)
        self.fit_matrix_btn = QtWidgets.QPushButton("Fit matrix parabolas...")
        controls.addWidget(self.fit_matrix_btn)
        layout.addLayout(controls)
        self.info_label = QtWidgets.QLabel("Click a point to open its spectroscopy")
        layout.addWidget(self.info_label)
        layout.addSpacing(6)
        self.setLayout(layout)
        self.canvas.mpl_connect("button_press_event", self._on_click)
        self.canvas.mpl_connect("motion_notify_event", self._on_canvas_hover)
        self._fit_dialogs = []
        self._current_image_arr = None
        self._current_image_extent = None
        self._current_image_unit = ''
        self._populate_channels()
        self.channel_combo.currentIndexChanged.connect(self._draw)
        self.fit_matrix_btn.clicked.connect(self._on_fit_matrix)
        self._draw()

    def _populate_channels(self):
        self.channel_combo.clear()
        path = Path(self.image_entry['path'])
        header, fds = self.viewer.headers.get(str(path), (None, None))
        if not fds:
            return
        for idx, fd in enumerate(fds):
            name = fd.get('Caption', fd.get('FileName', f"chan{idx}"))
            scale = fd.get('Scale')
            offset = fd.get('Offset')
            unit = fd.get('PhysUnit', '')
            self.channel_combo.addItem(f"{idx}: {name}", (idx, scale, offset, unit))
        if self.viewer.last_preview and self.viewer.last_preview[0] == str(path):
            self.channel_combo.setCurrentIndex(int(self.viewer.last_preview[1]))
        else:
            self.channel_combo.setCurrentIndex(0)

    def _draw(self):
        path = Path(self.image_entry['path'])
        header, fds = self.viewer.headers.get(str(path), (None, None))
        data = self.channel_combo.currentData()
        if data:
            main_idx = int(data[0])
        else:
            main_idx = 0
        try:
            if header and fds and 0 <= main_idx < len(fds):
                fd = fds[main_idx]
                arr = self.viewer._get_channel_array(str(path), main_idx, header, fd)
                arr = np.asarray(arr, dtype=float)
                self.ax.imshow(arr, cmap='gray', origin='upper')
                self._current_image_arr = arr
                self._current_image_extent = None
                self._current_image_unit = fd.get('PhysUnit', '')
            else:
                self.ax.text(0.5, 0.5, Path(path).name, ha='center', va='center', transform=self.ax.transAxes)
                self._current_image_arr = None
        except Exception:
            self.ax.text(0.5, 0.5, Path(path).name, ha='center', va='center', transform=self.ax.transAxes)
            self._current_image_arr = None
        xs = []
        ys = []
        xpix = int(header.get('xPixel', 128) if header else 128)
        ypix = int(header.get('yPixel', 128) if header else 128)
        for spec in self.specs:
            coords = self.viewer._map_spec_to_pixels(spec, header or {}, xpix, ypix)
            if coords is None:
                continue
            col, row = coords
            xs.append(col)
            ys.append(row)
        if xs and ys:
            self.ax.scatter(xs, ys, s=30, c='red', alpha=0.8)
        self.canvas.draw_idle()
        if self._current_image_arr is None:
            self.image_value_label.setText("Value: --")

    def _pick_spec_from_point(self, x, y):
        best = None
        best_dist = None
        header, _ = self.viewer.headers.get(str(self.image_entry['path']), (None, None))
        xpix = int(header.get('xPixel', 128) if header else 128)
        ypix = int(header.get('yPixel', 128) if header else 128)
        for spec in self.specs:
            coords = self.viewer._map_spec_to_pixels(spec, header or {}, xpix, ypix)
            if coords is None:
                continue
            col, row = coords
            dist = (col - x)**2 + (row - y)**2
            if best is None or dist < best_dist:
                best = spec
                best_dist = dist
        return best

    def _on_click(self, event):
        if event.inaxes != self.ax:
            return
        spec = self._pick_spec_from_point(event.xdata, event.ydata)
        if not spec:
            return
        self.viewer._open_spectroscopy_popup(spec)

    def _on_fit_matrix(self):
        dlg = MatrixFitDialog(self.viewer, self.specs, parent=self)
        dlg.setAttribute(QtCore.Qt.WA_DeleteOnClose, True)
        dlg.show()
        self._fit_dialogs.append(dlg)
        dlg.finished.connect(lambda _: self._cleanup_fit_dialog(dlg))

    def _cleanup_fit_dialog(self, dlg):
        try:
            self._fit_dialogs.remove(dlg)
        except ValueError:
            pass

    def _on_canvas_hover(self, event):
        if event.inaxes != self.ax or self._current_image_arr is None:
            self.image_value_label.setText("Value: --")
            return
        val = sample_array_value(self._current_image_arr, event.xdata, event.ydata, self._current_image_extent)
        if val is None:
            self.image_value_label.setText("Value: --")
            return
        unit = self._current_image_unit or ''
        txt = f"Value: {val:.4g}"
        if unit:
            txt += f" {unit}"
        self.image_value_label.setText(txt)

class MatrixFitWorker(QtCore.QObject):
    progress = QtCore.pyqtSignal(int, int)
    finished = QtCore.pyqtSignal(object)

    def __init__(self, specs):
        super().__init__()
        self.specs = list(specs)

    @QtCore.pyqtSlot()
    def run(self):
        specs = self.specs
        if not specs:
            self.finished.emit({
                'maps': {},
                'logs': ["No spectra to fit"],
                'channel_name': "channel",
                'x_axis': None,
                'y_axis': None,
            })
            return
        first_channels = specs[0].get('channels') or {}
        channel_name = next(iter(first_channels.keys()), 'channel')
        col_candidates = [spec.get('grid_col') for spec in specs if spec.get('grid_col') is not None]
        row_candidates = [spec.get('grid_row') for spec in specs if spec.get('grid_row') is not None]
        grid_cols = grid_rows = None
        if col_candidates and row_candidates:
            grid_cols = max(col_candidates) + 1
            grid_rows = max(row_candidates) + 1
        else:
            idx_candidates = [spec.get('matrix_index') for spec in specs if spec.get('matrix_index') is not None]
            if idx_candidates:
                max_idx = max(idx_candidates)
                side = int(round(math.sqrt(max_idx + 1)))
                if side > 0:
                    grid_cols = grid_rows = side
        if not grid_cols or not grid_rows:
            total = len(specs)
            grid_cols = int(round(math.sqrt(total))) or 1
            grid_rows = int(math.ceil(total / grid_cols)) or 1
        maps = {
            'a': np.full((grid_rows, grid_cols), np.nan),
            'b': np.full((grid_rows, grid_cols), np.nan),
            'c': np.full((grid_rows, grid_cols), np.nan),
            'a_err': np.full((grid_rows, grid_cols), np.nan),
            'b_err': np.full((grid_rows, grid_cols), np.nan),
            'c_err': np.full((grid_rows, grid_cols), np.nan),
            'rmse': np.full((grid_rows, grid_cols), np.nan),
        }
        def _axis_from_specs(coord_key, index_key, size):
            if not size:
                return np.arange(0, dtype=float)
            coords = [None] * size
            for spec in specs:
                idx = spec.get(index_key)
                val = spec.get(coord_key)
                if idx is None or val is None:
                    continue
                if idx < 0 or idx >= size:
                    continue
                try:
                    coords[idx] = float(val)
                except Exception:
                    continue
            if any(v is None for v in coords):
                return np.arange(size, dtype=float)
            arr = np.asarray(coords, dtype=float)
            arr = arr - float(np.nanmin(arr))
            return arr

        logs = []
        for idx, spec in enumerate(specs):
            row = spec.get('grid_row')
            col = spec.get('grid_col')
            if row is None or col is None:
                matrix_index = spec.get('matrix_index')
                if matrix_index is not None:
                    row = matrix_index // grid_cols
                    col = matrix_index % grid_cols
                else:
                    row = idx // grid_cols
                    col = idx % grid_cols
            try:
                if row < 0 or row >= grid_rows or col < 0 or col >= grid_cols:
                    raise IndexError(f"Index {idx}: ({row}, {col}) outside grid {grid_rows}x{grid_cols}")
                V = np.asarray(spec.get('V', []), dtype=float)
                channel_data = (spec.get('channels') or {}).get(channel_name)
                if channel_data is None:
                    raise ValueError("Channel missing")
                res = fit_parabola_bias(V, channel_data)
                maps['a'][row, col] = res['a']
                maps['b'][row, col] = res['b']
                maps['c'][row, col] = res['c']
                maps['a_err'][row, col] = res['a_err']
                maps['b_err'][row, col] = res['b_err']
                maps['c_err'][row, col] = res['c_err']
                maps['rmse'][row, col] = res['rmse']
            except Exception as exc:
                logs.append(f"Index {idx}: {exc}")
            current = idx + 1
            total = len(specs)
            self.progress.emit(current, total)
            try:
                print(f"[MatrixFit] {current}/{total} processed", flush=True)
            except Exception:
                pass
        payload = {
            'maps': maps,
            'logs': logs,
            'channel_name': channel_name,
            'x_axis': _axis_from_specs('x', 'grid_col', grid_cols),
            'y_axis': _axis_from_specs('y', 'grid_row', grid_rows),
        }
        self.finished.emit(payload)


class MatrixFitDialog(QtWidgets.QDialog):
    PARAM_INFO = {
        'a': {'label': 'a', 'unit': 'a.u.', 'cmap': 'viridis'},
        'b': {'label': 'b (LCPD)', 'unit': 'mV', 'cmap': 'bwr'},
        'c': {'label': 'c', 'unit': 'Hz', 'cmap': 'gray'},
        'a_err': {'label': 'sa', 'unit': 'a.u.', 'cmap': 'magma'},
        'b_err': {'label': 'sb', 'unit': 'mV', 'cmap': 'magma'},
        'c_err': {'label': 'sc', 'unit': 'Hz', 'cmap': 'magma'},
        'rmse': {'label': 'RMSE', 'unit': 'Hz', 'cmap': 'inferno'},
    }

    def __init__(self, viewer, specs, parent=None):
        super().__init__(parent)
        self.viewer = viewer
        self.specs = list(specs)
        self.setWindowTitle("Matrix parabola fits")
        self.resize(900, 700)
        self._worker_thread = None
        self._result_payload = None
        layout = QtWidgets.QVBoxLayout(self)
        self.info_label = QtWidgets.QLabel("Fit df(V) parabolas for every point in the matrix.")
        layout.addWidget(self.info_label)
        ctrl = QtWidgets.QHBoxLayout()
        self.run_btn = QtWidgets.QPushButton("Run fits")
        self.save_btn = QtWidgets.QPushButton("Save maps...")
        self.save_btn.setEnabled(False)
        self.export_xyz_btn = QtWidgets.QPushButton("Export WSxM XYZ...")
        self.export_xyz_btn.setEnabled(False)
        ctrl.addWidget(self.run_btn)
        ctrl.addWidget(self.save_btn)
        ctrl.addWidget(self.export_xyz_btn)
        ctrl.addStretch(1)
        layout.addLayout(ctrl)
        display_box = QtWidgets.QGroupBox("Display options")
        display_layout = QtWidgets.QHBoxLayout(display_box)
        self.scale_mode_combo = QtWidgets.QComboBox()
        self.scale_mode_combo.addItem("Full range", "full")
        self.scale_mode_combo.addItem("Clip percentiles", "clip")
        self.scale_mode_combo.addItem("Centered ?max", "center")
        display_layout.addWidget(QtWidgets.QLabel("Scale:"))
        display_layout.addWidget(self.scale_mode_combo)
        self.low_pct_spin = QtWidgets.QDoubleSpinBox()
        self.low_pct_spin.setRange(0.0, 49.0)
        self.low_pct_spin.setSingleStep(0.5)
        self.low_pct_spin.setValue(2.0)
        self.high_pct_spin = QtWidgets.QDoubleSpinBox()
        self.high_pct_spin.setRange(51.0, 100.0)
        self.high_pct_spin.setSingleStep(0.5)
        self.high_pct_spin.setValue(98.0)
        display_layout.addWidget(QtWidgets.QLabel("Low %"))
        display_layout.addWidget(self.low_pct_spin)
        display_layout.addWidget(QtWidgets.QLabel("High %"))
        display_layout.addWidget(self.high_pct_spin)
        display_layout.addStretch(1)
        layout.addWidget(display_box)
        self.progress = QtWidgets.QProgressBar()
        layout.addWidget(self.progress)
        self.fig = Figure(figsize=(6,5))
        self.canvas = FigureCanvas(self.fig)
        layout.addWidget(self.canvas, 1)
        self.map_value_label = QtWidgets.QLabel("Value: --")
        layout.addWidget(self.map_value_label)
        self.logs = QtWidgets.QTextEdit()
        self.logs.setReadOnly(True)
        self.logs.setFixedHeight(120)
        layout.addWidget(self.logs)
        btn_box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Close)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)
        self.run_btn.clicked.connect(self._start_fit)
        self.save_btn.clicked.connect(self._save_maps)
        self.export_xyz_btn.clicked.connect(self._export_xyz)
        self.scale_mode_combo.currentIndexChanged.connect(self._on_display_option_changed)
        self.low_pct_spin.valueChanged.connect(self._on_display_option_changed)
        self.high_pct_spin.valueChanged.connect(self._on_display_option_changed)
        self._update_percentile_enabled()
        self._axes_to_key = {}
        self.canvas.mpl_connect('motion_notify_event', self._on_map_hover)

    def _start_fit(self):
        if self._worker_thread is not None:
            return
        self.run_btn.setEnabled(False)
        self.save_btn.setEnabled(False)
        self.export_xyz_btn.setEnabled(False)
        self.logs.clear()
        self.progress.setValue(0)
        self._result_payload = None
        worker = MatrixFitWorker(self.specs)
        thread = QtCore.QThread(self)
        self._worker = worker
        worker.moveToThread(thread)
        worker.progress.connect(self._on_progress)
        worker.finished.connect(self._on_finished)
        worker.finished.connect(thread.quit)
        worker.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)
        thread.finished.connect(self._on_thread_finished)
        thread.started.connect(worker.run)
        self._worker_thread = thread
        thread.start()

    def _on_progress(self, current, total):
        self.progress.setMaximum(total)
        self.progress.setValue(current)

    def _on_finished(self, payload):
        self._result_payload = payload
        maps = payload.get('maps', {})
        logs = payload.get('logs', [])
        channel_name = payload.get('channel_name', 'channel')
        for line in logs:
            self.logs.append(line)
        if maps:
            self._render_maps(maps, channel_name)
            self.save_btn.setEnabled(True)
            self.export_xyz_btn.setEnabled(True)
        else:
            self.map_value_label.setText("Value: --")
        self.run_btn.setEnabled(True)
        self._worker = None

    def _on_thread_finished(self):
        self._worker_thread = None

    def _current_display_mode(self):
        return self.scale_mode_combo.currentData()

    def _current_percentiles(self):
        return float(self.low_pct_spin.value()), float(self.high_pct_spin.value())

    def _update_percentile_enabled(self):
        clip = (self._current_display_mode() == 'clip')
        self.low_pct_spin.setEnabled(clip)
        self.high_pct_spin.setEnabled(clip)

    def _on_display_option_changed(self):
        self._update_percentile_enabled()
        if self._result_payload and self._result_payload.get('maps'):
            maps = self._result_payload['maps']
            channel = self._result_payload.get('channel_name', 'channel')
            self._render_maps(maps, channel)
        else:
            self.canvas.draw_idle()

    def _compute_vlims(self, arr):
        mode = self._current_display_mode()
        data = np.asarray(arr, dtype=float)
        if mode == 'clip':
            low, high = self._current_percentiles()
            return robust_limits(data, low_pct=low, high_pct=high)
        finite = data[np.isfinite(data)]
        if finite.size == 0:
            return None, None
        if mode == 'center':
            vmax = float(np.nanmax(np.abs(finite)))
            if not np.isfinite(vmax) or vmax == 0:
                return None, None
            return -vmax, vmax
        return None, None

    def _map_extent(self, arr_shape):
        payload = self._result_payload or {}
        x_axis = payload.get('x_axis')
        y_axis = payload.get('y_axis')
        if x_axis is None or y_axis is None:
            return None
        if len(x_axis) != arr_shape[1] or len(y_axis) != arr_shape[0]:
            return None
        try:
            x0 = float(np.nanmin(x_axis))
            x1 = float(np.nanmax(x_axis))
            y0 = float(np.nanmin(y_axis))
            y1 = float(np.nanmax(y_axis))
        except Exception:
            return None
        if not np.isfinite([x0, x1, y0, y1]).all() or x0 == x1 or y0 == y1:
            return None
        return [x0, x1, y0, y1]

    def _render_maps(self, maps, channel_name):
        self.fig.clf()
        self._axes_to_key = {}
        params = ['a','b','c','a_err','b_err','c_err','rmse']
        cols = 3
        rows = math.ceil(len(params)/cols)
        for idx, key in enumerate(params, 1):
            ax = self.fig.add_subplot(rows, cols, idx)
            info = self.PARAM_INFO.get(key, {'label': key, 'unit': ''})
            ax.set_title(info['label'])
            vmin, vmax = self._compute_vlims(maps[key])
            extent = self._map_extent(maps[key].shape)
            cmap = info.get('cmap', 'viridis')
            im = ax.imshow(maps[key], origin='lower', cmap=cmap, vmin=vmin, vmax=vmax, extent=extent)
            cbar = self.fig.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
            unit = info.get('unit')
            if unit:
                cbar.set_label(unit)
            if extent:
                ax.set_xlabel("x (nm)")
                ax.set_ylabel("y (nm)")
            self._axes_to_key[ax] = key
        self.fig.suptitle(f"Parabola fits - channel {channel_name}")
        self.canvas.draw_idle()

    def _save_maps(self):
        if not self._result_payload or not self._result_payload.get('maps'):
            return
        maps = self._result_payload['maps']
        channel_name = self._result_payload.get('channel_name', 'channel')
        x_axis = self._result_payload.get('x_axis')
        y_axis = self._result_payload.get('y_axis')
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save fit maps", "matrix_fit_maps.npz", "NumPy archive (*.npz)")
        if not path:
            return
        metadata = self._collect_fit_metadata(x_axis, y_axis, maps)
        metadata_json = json.dumps(metadata)
        np.savez(path, channel=channel_name, x_axis=x_axis, y_axis=y_axis, metadata=np.array(metadata_json), **maps)
        metadata_path = Path(path).with_suffix('.json')
        try:
            metadata_path.write_text(json.dumps(metadata, indent=2, default=str))
        except Exception:
            pass

    def _export_xyz(self):
        if not self._result_payload or not self._result_payload.get('maps'):
            return
        maps = self._result_payload['maps']
        x_axis = self._result_payload.get('x_axis')
        y_axis = self._result_payload.get('y_axis')
        if x_axis is None or y_axis is None:
            QtWidgets.QMessageBox.warning(self, "Missing coordinates", "Cannot export XYZ without coordinate axes.")
            return
        folder = QtWidgets.QFileDialog.getExistingDirectory(self, "Select folder for WSxM XYZ exports")
        if not folder:
            return
        save_wsxm_xyz(folder, maps['a'], x_axis, y_axis, "a", z_unit="a.u.")
        save_wsxm_xyz(folder, maps['b'], x_axis, y_axis, "b_LCPD", z_unit="mV", z_scale=1000.0)
        save_wsxm_xyz(folder, maps['c'], x_axis, y_axis, "c", z_unit="Hz")
        save_wsxm_xyz(folder, maps['a_err'], x_axis, y_axis, "a_err", z_unit="a.u.")
        save_wsxm_xyz(folder, maps['b_err'], x_axis, y_axis, "b_err", z_unit="mV", z_scale=1000.0)
        save_wsxm_xyz(folder, maps['c_err'], x_axis, y_axis, "c_err", z_unit="Hz")
        save_wsxm_xyz(folder, maps['rmse'], x_axis, y_axis, "rmse", z_unit="Hz")
        self.logs.append(f"WSxM XYZ exports saved to {folder}")

    def get_result_maps(self):
        return self._result_payload

    def _on_map_hover(self, event):
        if self._result_payload is None or not self._result_payload.get('maps'):
            self.map_value_label.setText("Value: --")
            return
        if event.inaxes not in self._axes_to_key:
            self.map_value_label.setText("Value: --")
            return
        key = self._axes_to_key.get(event.inaxes)
        arr = self._result_payload['maps'].get(key)
        if arr is None:
            self.map_value_label.setText("Value: --")
            return
        extent = self._map_extent(arr.shape)
        val = sample_array_value(arr, event.xdata, event.ydata, extent)
        if val is None:
            self.map_value_label.setText("Value: --")
            return
        info = self.PARAM_INFO.get(key, {})
        unit = info.get('unit') or ''
        label = info.get('label', key)
        text = f"{label}: {val:.4g}"
        if unit:
            text += f" {unit}"
        self.map_value_label.setText(text)

    def _collect_fit_metadata(self, x_axis, y_axis, maps):
        specs = self.specs or []
        def _axis_stats(axis):
            if axis is None:
                return (None, None)
            arr = np.asarray(axis, dtype=float)
            if arr.size == 0:
                return (None, None)
            return (float(np.nanmin(arr)), float(np.nanmax(arr)))

        x_min, x_max = _axis_stats(x_axis)
        y_min, y_max = _axis_stats(y_axis)
        meta = {
            'channel': self._result_payload.get('channel_name') if self._result_payload else None,
            'spec_count': len(specs),
            'grid_shape': list(maps['a'].shape) if 'a' in maps else None,
            'x_axis_min': x_min,
            'x_axis_max': x_max,
            'y_axis_min': y_min,
            'y_axis_max': y_max,
        }
        if specs:
            first_path = specs[0].get('path')
            try:
                meta['source_file'] = str(Path(first_path))
            except Exception:
                meta['source_file'] = str(first_path)
        biases = [np.asarray(spec.get('V', []), dtype=float) for spec in specs if spec.get('V') is not None]
        if biases:
            all_bias = np.concatenate([b for b in biases if b.size])
            if all_bias.size:
                meta['bias_min'] = float(np.nanmin(all_bias))
                meta['bias_max'] = float(np.nanmax(all_bias))
            meta['points_per_spectrum'] = int(np.nanmedian([b.size for b in biases if b.size])) if biases else None
        xs = [spec.get('x') for spec in specs if spec.get('x') is not None]
        ys = [spec.get('y') for spec in specs if spec.get('y') is not None]
        if xs:
            meta['position_x_min'] = float(np.nanmin(xs))
            meta['position_x_max'] = float(np.nanmax(xs))
        if ys:
            meta['position_y_min'] = float(np.nanmin(ys))
            meta['position_y_max'] = float(np.nanmax(ys))
        times = [spec.get('time') for spec in specs if isinstance(spec.get('time'), datetime)]
        if times:
            times.sort()
            meta['acquisition_start'] = times[0].isoformat()
            meta['acquisition_end'] = times[-1].isoformat()
            meta['estimated_duration_seconds'] = float((times[-1] - times[0]).total_seconds())
        meta['saved_at'] = datetime.utcnow().isoformat()
        return meta

    def closeEvent(self, event):
        thread = self._worker_thread
        if thread is not None and thread.isRunning():
            thread.quit()
            thread.wait()
        super().closeEvent(event)

    def closeEvent(self, event):
        if self._worker_thread is not None and self._worker_thread.isRunning():
            self._worker_thread.quit()
            self._worker_thread.wait()
        super().closeEvent(event)
class CustomFilterDialog(QtWidgets.QDialog):
    """Dialog to assemble custom filter pipelines."""
    def __init__(self, parent=None, base_image=None, apply_step_func=None):
        super().__init__(parent)
        self.setWindowTitle("Custom filter pipeline")
        self.resize(460, 480)
        self.base_image = base_image
        self.apply_step = apply_step_func
        self._pipeline = []
        layout = QtWidgets.QVBoxLayout(self)
        form = QtWidgets.QFormLayout()
        self.filter_combo = QtWidgets.QComboBox()
        for key, info in FILTER_DEFINITIONS.items():
            self.filter_combo.addItem(info['label'], key)
        form.addRow("Filter", self.filter_combo)
        self.axis_combo = QtWidgets.QComboBox()
        self.axis_combo.addItems(["both","row","col"])
        form.addRow("Axis", self.axis_combo)
        self.sigma_spin = QtWidgets.QDoubleSpinBox()
        self.sigma_spin.setRange(0.1, 50.0); self.sigma_spin.setSingleStep(0.1); self.sigma_spin.setValue(2.0)
        form.addRow("Sigma", self.sigma_spin)
        layout.addLayout(form)
        btn_row = QtWidgets.QHBoxLayout()
        add_btn = QtWidgets.QPushButton("Add step")
        remove_btn = QtWidgets.QPushButton("Remove selected")
        btn_row.addWidget(add_btn); btn_row.addWidget(remove_btn)
        layout.addLayout(btn_row)
        self.pipeline_list = QtWidgets.QListWidget()
        layout.addWidget(self.pipeline_list, 1)
        self.preview_cb = QtWidgets.QCheckBox("Preview on current image")
        layout.addWidget(self.preview_cb)
        self.preview_label = QtWidgets.QLabel("Preview unavailable")
        self.preview_label.setFixedHeight(160)
        self.preview_label.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.preview_label.setAlignment(QtCore.Qt.AlignCenter)
        layout.addWidget(self.preview_label)
        name_row = QtWidgets.QHBoxLayout()
        name_row.addWidget(QtWidgets.QLabel("Name prefix:"))
        self.name_edit = QtWidgets.QLineEdit("Custom")
        name_row.addWidget(self.name_edit)
        layout.addLayout(name_row)
        btn_box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        layout.addWidget(btn_box)
        add_btn.clicked.connect(self._on_add_step)
        remove_btn.clicked.connect(self._on_remove_step)
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)
        self.preview_cb.toggled.connect(self._update_preview)

    def _current_step(self):
        key = self.filter_combo.currentData()
        params = {}
        if key == 'flatten':
            params['axis'] = self.axis_combo.currentText()
        if key in ('highpass','lowpass'):
            params['sigma'] = float(self.sigma_spin.value())
        return {'key': key, 'params': params}

    def _on_add_step(self):
        step = self._current_step()
        label = FILTER_DEFINITIONS.get(step['key'], {}).get('label', step['key'])
        self._pipeline.append(step)
        self.pipeline_list.addItem(f"{len(self._pipeline)}. {label}")
        self._update_preview()

    def _on_remove_step(self):
        row = self.pipeline_list.currentRow()
        if row >= 0:
            self.pipeline_list.takeItem(row)
            del self._pipeline[row]
            self.pipeline_list.clear()
            for idx, step in enumerate(self._pipeline, 1):
                label = FILTER_DEFINITIONS.get(step['key'], {}).get('label', step['key'])
                self.pipeline_list.addItem(f"{idx}. {label}")
            self._update_preview()

    def _update_preview(self):
        if not self.preview_cb.isChecked() or self.base_image is None or not self.apply_step:
            self.preview_label.setText("Preview unavailable")
            self.preview_label.setPixmap(QtGui.QPixmap())
            return
        arr = np.asarray(self.base_image, dtype=float)
        for step in self._pipeline:
            arr = self.apply_step(arr, step)
        qimg = array_to_qimage(arr)
        pix = QtGui.QPixmap.fromImage(qimg).scaled(self.preview_label.width(), self.preview_label.height(),
                                                   QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation)
        self.preview_label.setPixmap(pix)
        self.preview_label.setText("")

    def pipeline_steps(self):
        return list(self._pipeline)

    def pipeline_label(self):
        return self.name_edit.text().strip() or "Custom"


class CropPreviewLabel(QtWidgets.QLabel):
    selectionMade = QtCore.pyqtSignal(int, int, int, int)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._rubber = QtWidgets.QRubberBand(QtWidgets.QRubberBand.Rectangle, self)
        self._origin = None
        self._pixmap_rect = QtCore.QRect()
        self._array_shape = (1, 1)
        self.setMouseTracking(True)

    def set_array_shape(self, shape):
        self._array_shape = shape

    def set_display_pixmap_rect(self, rect):
        self._pixmap_rect = QtCore.QRect(rect)

    def mousePressEvent(self, event):
        if event.button() == QtCore.Qt.LeftButton and self._pixmap_rect.contains(event.pos()):
            self._origin = event.pos()
            self._rubber.setGeometry(QtCore.QRect(self._origin, QtCore.QSize()))
            self._rubber.show()
        else:
            super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._origin is not None:
            rect = QtCore.QRect(self._origin, event.pos()).normalized()
            self._rubber.setGeometry(rect.intersected(self._pixmap_rect))
        else:
            super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if self._origin is not None:
            rect = QtCore.QRect(self._origin, event.pos()).normalized().intersected(self._pixmap_rect)
            self._rubber.hide()
            self._emit_selection(rect)
            self._origin = None
        else:
            super().mouseReleaseEvent(event)

    def _emit_selection(self, rect):
        if rect.isNull() or rect.width() < 2 or rect.height() < 2:
            return
        cols = max(1, self._array_shape[1])
        rows = max(1, self._array_shape[0])
        def clamp(val, lo, hi):
            return max(lo, min(val, hi))
        left = clamp(rect.left(), self._pixmap_rect.left(), self._pixmap_rect.right())
        right = clamp(rect.right(), self._pixmap_rect.left(), self._pixmap_rect.right())
        top = clamp(rect.top(), self._pixmap_rect.top(), self._pixmap_rect.bottom())
        bottom = clamp(rect.bottom(), self._pixmap_rect.top(), self._pixmap_rect.bottom())
        if self._pixmap_rect.width() <= 0 or self._pixmap_rect.height() <= 0:
            return
        rel_x0 = (left - self._pixmap_rect.left()) / self._pixmap_rect.width()
        rel_x1 = (right - self._pixmap_rect.left()) / self._pixmap_rect.width()
        rel_y0 = (top - self._pixmap_rect.top()) / self._pixmap_rect.height()
        rel_y1 = (bottom - self._pixmap_rect.top()) / self._pixmap_rect.height()
        x0 = int(clamp(round(rel_x0 * cols), 0, cols - 1))
        x1 = int(clamp(round(rel_x1 * cols), 0, cols - 1))
        y0 = int(clamp(round(rel_y0 * rows), 0, rows - 1))
        y1 = int(clamp(round(rel_y1 * rows), 0, rows - 1))
        if x1 <= x0:
            x1 = min(cols - 1, x0 + 1)
        if y1 <= y0:
            y1 = min(rows - 1, y0 + 1)
        self.selectionMade.emit(x0, x1, y0, y1)


class ImageAdjustDialog(QtWidgets.QDialog):
    def __init__(self, parent, base_image, spec, cmap_name):
        super().__init__(parent)
        self.setWindowTitle("Image adjustments")
        self.base_image = np.asarray(base_image, dtype=float)
        self.current_spec = json.loads(json.dumps(spec))
        self.selected_cmap = cmap_name
        h, w = self.base_image.shape
        layout = QtWidgets.QVBoxLayout(self)
        form = QtWidgets.QFormLayout()
        self.x0_spin = QtWidgets.QSpinBox(); self.x0_spin.setRange(0, max(0, w-1))
        self.x1_spin = QtWidgets.QSpinBox(); self.x1_spin.setRange(1, w)
        self.y0_spin = QtWidgets.QSpinBox(); self.y0_spin.setRange(0, max(0, h-1))
        self.y1_spin = QtWidgets.QSpinBox(); self.y1_spin.setRange(1, h)
        form.addRow("Crop X start", self.x0_spin)
        form.addRow("Crop X end", self.x1_spin)
        form.addRow("Crop Y start", self.y0_spin)
        form.addRow("Crop Y end", self.y1_spin)
        self.rotate_slider = QtWidgets.QSlider(QtCore.Qt.Horizontal); self.rotate_slider.setRange(-180, 180)
        form.addRow("Rotate (deg)", self.rotate_slider)
        self.flip_h_cb = QtWidgets.QCheckBox("Flip horizontally")
        self.flip_v_cb = QtWidgets.QCheckBox("Flip vertically")
        form.addRow(self.flip_h_cb)
        form.addRow(self.flip_v_cb)
        self.low_pct_spin = QtWidgets.QDoubleSpinBox(); self.low_pct_spin.setRange(0.0, 50.0); self.low_pct_spin.setDecimals(2)
        self.high_pct_spin = QtWidgets.QDoubleSpinBox(); self.high_pct_spin.setRange(50.0, 100.0); self.high_pct_spin.setDecimals(2); self.high_pct_spin.setValue(100.0)
        form.addRow("Clip low %", self.low_pct_spin)
        form.addRow("Clip high %", self.high_pct_spin)
        self.gamma_spin = QtWidgets.QDoubleSpinBox(); self.gamma_spin.setRange(0.1, 5.0); self.gamma_spin.setSingleStep(0.1); self.gamma_spin.setValue(1.0)
        form.addRow("Gamma", self.gamma_spin)
        layout.addLayout(form)
        cmap_row = QtWidgets.QHBoxLayout()
        cmap_row.addWidget(QtWidgets.QLabel("Colormap:"))
        self.cmap_combo = QtWidgets.QComboBox()
        try:
            cmap_names = sorted(colormaps.keys())
        except Exception:
            cmap_names = ['viridis','plasma','inferno','magma','cividis','gray','hot','coolwarm','turbo']
        for name in cmap_names:
            self.cmap_combo.addItem(name)
        if cmap_name in cmap_names:
            self.cmap_combo.setCurrentText(cmap_name)
        cmap_row.addWidget(self.cmap_combo, 1)
        layout.addLayout(cmap_row)
        self.preview_label = CropPreviewLabel()
        self.preview_label.setMinimumHeight(220)
        self.preview_label.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.preview_label.setAlignment(QtCore.Qt.AlignCenter)
        layout.addWidget(self.preview_label)
        self.preview_label.selectionMade.connect(self._on_crop_selection)
        btn_box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        layout.addWidget(btn_box)
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)
        for widget in (self.x0_spin, self.x1_spin, self.y0_spin, self.y1_spin,
                       self.rotate_slider, self.flip_h_cb, self.flip_v_cb,
                       self.low_pct_spin, self.high_pct_spin, self.gamma_spin):
            if isinstance(widget, QtWidgets.QSlider):
                widget.valueChanged.connect(self._on_params_changed)
            elif isinstance(widget, QtWidgets.QAbstractButton):
                widget.toggled.connect(self._on_params_changed)
            else:
                widget.valueChanged.connect(self._on_params_changed)
        self.cmap_combo.currentIndexChanged.connect(self._on_cmap_changed)
        self._apply_spec_to_controls()
        self._update_preview()

    def _apply_spec_to_controls(self):
        crop = self.current_spec.get('crop', {})
        if not crop:
            crop = {'x0': 0, 'y0': 0, 'x1': self.base_image.shape[1], 'y1': self.base_image.shape[0]}
            self.current_spec['crop'] = crop
        self.x0_spin.setValue(int(crop.get('x0', 0)))
        self.x1_spin.setValue(int(crop.get('x1', self.base_image.shape[1])))
        self.y0_spin.setValue(int(crop.get('y0', 0)))
        self.y1_spin.setValue(int(crop.get('y1', self.base_image.shape[0])))
        self.rotate_slider.setValue(int(round(self.current_spec.get('rotate', 0.0))))
        self.flip_h_cb.setChecked(bool(self.current_spec.get('flip_h')))
        self.flip_v_cb.setChecked(bool(self.current_spec.get('flip_v')))
        clip = self.current_spec.get('clip', {})
        if clip.get('low') is not None:
            self.low_pct_spin.setValue(float(clip.get('low')))
        else:
            self.low_pct_spin.setValue(0.0)
        if clip.get('high') is not None:
            self.high_pct_spin.setValue(float(clip.get('high')))
        else:
            self.high_pct_spin.setValue(100.0)
        self.gamma_spin.setValue(float(self.current_spec.get('gamma', 1.0)))
        cmap = self.current_spec.get('cmap', self.cmap_combo.currentText())
        self.cmap_combo.setCurrentText(cmap)

    def _on_params_changed(self, value=None):
        self.current_spec['crop'] = {
            'x0': int(self.x0_spin.value()),
            'x1': int(self.x1_spin.value()),
            'y0': int(self.y0_spin.value()),
            'y1': int(self.y1_spin.value()),
        }
        self.current_spec['rotate'] = float(self.rotate_slider.value())
        self.current_spec['flip_h'] = self.flip_h_cb.isChecked()
        self.current_spec['flip_v'] = self.flip_v_cb.isChecked()
        low = float(self.low_pct_spin.value())
        high = float(self.high_pct_spin.value())
        self.current_spec['clip'] = {
            'low': low if low > 0 else None,
            'high': high if high < 100 else None,
        }
        self.current_spec['gamma'] = float(self.gamma_spin.value())
        self.current_spec['cmap'] = self.cmap_combo.currentText()
        self._update_preview()

    def _update_preview(self):
        arr, _ = apply_adjustment_spec(self.base_image, None, self.current_spec)
        cmap_name = self.cmap_combo.currentText() or 'viridis'
        qimg = array_to_qimage(arr, cmap_name=cmap_name)
        pix = QtGui.QPixmap.fromImage(qimg).scaled(
            max(1, self.preview_label.width()),
            max(1, self.preview_label.height()),
            QtCore.Qt.KeepAspectRatio,
            QtCore.Qt.SmoothTransformation)
        self.preview_label.setPixmap(pix)
        label_w = max(1, self.preview_label.width())
        label_h = max(1, self.preview_label.height())
        offset_x = (label_w - pix.width()) // 2
        offset_y = (label_h - pix.height()) // 2
        rect = QtCore.QRect(offset_x, offset_y, pix.width(), pix.height())
        self.preview_label.set_display_pixmap_rect(rect)
        self.preview_label.set_array_shape(self.base_image.shape)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        QtCore.QTimer.singleShot(0, self._update_preview)

    def _on_crop_selection(self, x0, x1, y0, y1):
        self.x0_spin.setValue(x0)
        self.x1_spin.setValue(x1)
        self.y0_spin.setValue(y0)
        self.y1_spin.setValue(y1)
        self._on_params_changed()

    def _on_cmap_changed(self):
        self.current_spec['cmap'] = self.cmap_combo.currentText()
        self._update_preview()

class BatchExportSignals(QtCore.QObject):
    progress = QtCore.pyqtSignal(int, int, str)
    finished = QtCore.pyqtSignal(list, list, bool)

class BatchExportWorker(QtCore.QRunnable):
    def __init__(self, parent, paths, config, out_dir):
        super().__init__()
        self.parent = parent
        self.paths = [str(p) for p in paths]
        self.config = config
        self.out_dir = Path(out_dir)
        self.signals = BatchExportSignals()
        self._cancelled = False

    def cancel(self):
        self._cancelled = True

    def run(self):
        saved = []
        errors = []
        total = len(self.paths)
        for idx, path in enumerate(self.paths, 1):
            if self._cancelled:
                break
            try:
                result = self.parent.render_and_save_file_using_config(Path(path), self.config, self.out_dir)
                saved.extend(result)
            except Exception as e:
                errors.append(f"{Path(path).name}: {e}")
            self.signals.progress.emit(idx, total, path)
        self.signals.finished.emit(saved, errors, self._cancelled)


class _SpectroFitWorker(QtCore.QObject):
    finished = QtCore.pyqtSignal(list, list)

    def __init__(self, specs, channel):
        super().__init__()
        self.specs = list(specs)
        self.channel = channel

    def run(self):
        results = []
        logs = []
        for spec in self.specs:
            name = Path(spec['path']).name
            V = np.asarray(spec.get('V', []), dtype=float)
            channels = spec.get('channels') or {}
            data = channels.get(self.channel)
            if data is None or not V.size:
                logs.append(f"{name}: channel '{self.channel}' unavailable")
                continue
            try:
                res = fit_parabola_bias(V, data)
                res['spec'] = spec
                results.append(res)
                logs.append(f"{name}: fit ok (RMSE {res['rmse']:.3g})")
            except Exception as e:
                logs.append(f"{name}: {e}")
        self.finished.emit(results, logs)


class SpectroscopyCompareDialog(QtWidgets.QDialog):
    """Modern comparison UI for spectroscopy overlays and fitting."""
    def __init__(self, specs, parent=None):
        super().__init__(parent)
        self.specs = list(specs)
        self._line_map = {}
        self._legend_map = {}
        self._fit_results = {}
        self._fit_thread = None
        self._fit_worker = None
        self._popup_refs = []
        self.setWindowTitle("Spectroscopy comparison")
        self.resize(1250, 640)
        self._build_ui()
        self._populate_list()
        self._populate_channels()
        self._update_plot()

    def _build_ui(self):
        splitter = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        main_layout = QtWidgets.QVBoxLayout()
        main_layout.addWidget(splitter)
        self.setLayout(main_layout)

        # Left panel: filter + list
        left = QtWidgets.QWidget()
        left_layout = QtWidgets.QVBoxLayout(left)
        left_layout.setContentsMargins(4,4,4,4)
        self.filter_edit = QtWidgets.QLineEdit()
        self.filter_edit.setPlaceholderText("Filter spectra...")
        self.filter_edit.textChanged.connect(self._apply_filter)
        self.spec_list = QtWidgets.QListWidget()
        self.spec_list.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
        self.spec_list.itemChanged.connect(self._on_item_check_changed)
        self.spec_list.itemSelectionChanged.connect(self._on_list_selection_changed)
        self.spec_list.itemDoubleClicked.connect(self._on_item_double_clicked)
        self.spec_list.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.spec_list.customContextMenuRequested.connect(self._on_list_context_menu)
        left_layout.addWidget(self.filter_edit)
        left_layout.addWidget(self.spec_list, 1)
        splitter.addWidget(left)
        splitter.setStretchFactor(0, 0)

        # Center panel: plot + status
        center = QtWidgets.QWidget()
        center_layout = QtWidgets.QVBoxLayout(center)
        center_layout.setContentsMargins(4,4,4,4)
        self.fig = Figure(figsize=(5,4))
        self.canvas = FigureCanvas(self.fig)
        self.ax = self.fig.add_subplot(111)
        self.ax.grid(True, alpha=0.2)
        center_layout.addWidget(self.canvas, 1)
        self.status_label = QtWidgets.QLabel("0 selected / 0 total")
        center_layout.addWidget(self.status_label)
        splitter.addWidget(center)
        splitter.setStretchFactor(1, 2)
        self.canvas.mpl_connect('pick_event', self._on_legend_pick)

        # Right panel: controls + results
        right = QtWidgets.QWidget()
        right_layout = QtWidgets.QVBoxLayout(right)
        right_layout.setContentsMargins(6,6,6,6)
        channel_row = QtWidgets.QHBoxLayout()
        channel_row.addWidget(QtWidgets.QLabel("Channel:"))
        self.channel_combo = QtWidgets.QComboBox()
        self.channel_combo.currentTextChanged.connect(self._on_channel_changed)
        channel_row.addWidget(self.channel_combo, 1)
        right_layout.addLayout(channel_row)

        btn_row = QtWidgets.QHBoxLayout()
        self.fit_selected_btn = QtWidgets.QPushButton("Fit selected (F)")
        self.fit_all_btn = QtWidgets.QPushButton("Fit all")
        self.export_btn = QtWidgets.QPushButton("Export CSV")
        self.copy_btn = QtWidgets.QPushButton("Copy selected")
        btn_row.addWidget(self.fit_selected_btn)
        btn_row.addWidget(self.fit_all_btn)
        btn_row.addWidget(self.export_btn)
        btn_row.addWidget(self.copy_btn)
        right_layout.addLayout(btn_row)
        self.fit_selected_btn.clicked.connect(self._fit_selected)
        self.fit_all_btn.clicked.connect(self._fit_all)
        self.export_btn.clicked.connect(self._export_csv)
        self.copy_btn.clicked.connect(self._copy_selected_to_clipboard)

        QtWidgets.QShortcut(QtGui.QKeySequence("F"), self, activated=self._fit_selected)
        QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+E"), self, activated=self._export_csv)

        self.options_toggle = QtWidgets.QToolButton()
        self.options_toggle.setText("Fit options")
        self.options_toggle.setCheckable(True)
        self.options_toggle.setChecked(False)
        self.options_toggle.setArrowType(QtCore.Qt.RightArrow)
        self.options_toggle.setToolButtonStyle(QtCore.Qt.ToolButtonTextBesideIcon)
        self.options_toggle.toggled.connect(self._on_options_toggled)
        right_layout.addWidget(self.options_toggle)

        self.options_body = QtWidgets.QWidget()
        form = QtWidgets.QFormLayout(self.options_body)
        self.degree_spin = QtWidgets.QSpinBox()
        self.degree_spin.setRange(2, 2)
        self.degree_spin.setValue(2)
        self.degree_spin.setEnabled(False)
        form.addRow("Degree", self.degree_spin)
        self.mask_min = QtWidgets.QDoubleSpinBox(); self.mask_min.setRange(-1e6, 1e6); self.mask_min.setSuffix(" V")
        self.mask_max = QtWidgets.QDoubleSpinBox(); self.mask_max.setRange(-1e6, 1e6); self.mask_max.setSuffix(" V")
        form.addRow("Mask min", self.mask_min)
        form.addRow("Mask max", self.mask_max)
        self.options_body.setVisible(False)
        right_layout.addWidget(self.options_body)

        self.results_table = QtWidgets.QTableWidget(0, 10)
        self.results_table.setHorizontalHeaderLabels(["File","X (nm)","Y (nm)","a","?a","b (mV)","?b","c (Hz)","?c","RMSE"])
        self.results_table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.results_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.results_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.results_table.itemSelectionChanged.connect(self._on_table_selection)
        self.results_table.itemDoubleClicked.connect(self._on_table_double_clicked)
        self.results_table.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.results_table.customContextMenuRequested.connect(self._on_table_context_menu)
        right_layout.addWidget(self.results_table, 1)

        self.log = QtWidgets.QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumHeight(100)
        right_layout.addWidget(self.log)

        splitter.addWidget(right)
        splitter.setStretchFactor(2, 1)

    def _populate_list(self):
        self.spec_list.blockSignals(True)
        self.spec_list.clear()
        self._item_map = {}
        for spec in self.specs:
            item = QtWidgets.QListWidgetItem(Path(spec['path']).name)
            item.setFlags(item.flags() | QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsSelectable)
            item.setCheckState(QtCore.Qt.Checked)
            item.setData(QtCore.Qt.UserRole, spec)
            self.spec_list.addItem(item)
            self._item_map[str(Path(spec['path']))] = item
        self.spec_list.blockSignals(False)

    def _populate_channels(self):
        channels = sorted({name for spec in self.specs for name in (spec.get('channels') or {}).keys()})
        self.channel_combo.blockSignals(True)
        self.channel_combo.clear()
        for name in channels:
            self.channel_combo.addItem(name)
        if channels:
            self.channel_combo.setCurrentText('df' if 'df' in channels else channels[0])
        self.channel_combo.blockSignals(False)

    def _apply_filter(self, text):
        text = text.lower()
        for i in range(self.spec_list.count()):
            item = self.spec_list.item(i)
            item.setHidden(text not in item.text().lower())
        self._update_status()

    def _checked_items(self):
        return [self.spec_list.item(i) for i in range(self.spec_list.count())
                if self.spec_list.item(i).checkState() == QtCore.Qt.Checked and not self.spec_list.item(i).isHidden()]

    def _selected_items(self):
        return [item for item in self._checked_items() if item.isSelected()]

    def _on_channel_changed(self):
        self._fit_results = {}
        self._populate_results_table()
        self._update_plot()

    def _on_item_check_changed(self):
        self._update_plot()

    def _on_list_selection_changed(self):
        self._update_plot()

    def _update_plot(self):
        channel = self.channel_combo.currentText()
        self.ax.clear()
        self.ax.grid(True, alpha=0.2)
        self._line_map.clear()
        self._legend_map.clear()
        selected_ids = {str(Path(item.data(QtCore.Qt.UserRole)['path'])) for item in self._selected_items()}
        colors = itertools.cycle(matplotlib.cm.get_cmap('tab10').colors)
        plotted = 0
        for item in self._checked_items():
            spec = item.data(QtCore.Qt.UserRole)
            spec_id = str(Path(spec['path']))
            channels = spec.get('channels') or {}
            data = channels.get(channel)
            V = np.asarray(spec.get('V', []), dtype=float)
            if data is None or not V.size:
                continue
            color = next(colors)
            highlight = spec_id in selected_ids or not selected_ids
            line, = self.ax.plot(V, data, color=color, lw=2.4 if highlight else 1.2,
                                 alpha=1.0 if highlight else 0.4, label=Path(spec['path']).name)
            self._line_map[spec_id] = line
            plotted += 1
            if spec_id in self._fit_results:
                self._draw_fit_for_spec(spec_id, color)
        if plotted == 0:
            self.ax.text(0.5,0.5,"No data for selected items", ha='center', va='center', transform=self.ax.transAxes)
        else:
            legend = self.ax.legend(loc='best', fontsize=8)
            if legend:
                legend.set_draggable(True)
                for leg_line, text in zip(legend.get_lines(), legend.get_texts()):
                    leg_line.set_picker(True)
                    spec_id = self._spec_id_by_name(text.get_text())
                    if spec_id:
                        self._legend_map[leg_line] = spec_id
        self.ax.set_xlabel("Bias (mV)")
        self.ax.set_ylabel(channel)
        self.canvas.draw_idle()
        self._update_status(plotted)

    def _draw_fit_for_spec(self, spec_id, color):
        res = self._fit_results.get(spec_id)
        if not res:
            return
        spec = res.get('spec')
        V = np.asarray(spec.get('V', []), dtype=float)
        if not V.size:
            return
        x_dense = np.linspace(np.nanmin(V), np.nanmax(V), 400)
        self.ax.plot(x_dense, res['func'](x_dense), '--', color=color, lw=1.2)
        b = res['b']; c = res['c']; be = res.get('b_err', 0.0)
        self.ax.errorbar([b], [c], xerr=[be], fmt='o', color=color, ecolor=color, capsize=3)

    def _spec_id_by_name(self, name):
        for spec in self.specs:
            if Path(spec['path']).name == name:
                return str(Path(spec['path']))
        return None

    def _on_legend_pick(self, event):
        spec_id = self._legend_map.get(event.artist)
        if not spec_id:
            return
        line = self._line_map.get(spec_id)
        if not line:
            return
        visible = not line.get_visible()
        line.set_visible(visible)
        event.artist.set_alpha(1.0 if visible else 0.2)
        self.canvas.draw_idle()

    def _update_status(self, plotted=None):
        total = sum(1 for i in range(self.spec_list.count()) if not self.spec_list.item(i).isHidden())
        checked = len(self._checked_items())
        text = f"{checked} selected / {total} total"
        if plotted is not None:
            text += f" ? showing {plotted}"
        self.status_label.setText(text)

    def _show_popup_for_spec(self, spec):
        dlg = SpectroscopyPopup(spec, parent=self)
        dlg.show()
        self._popup_refs.append(dlg)

    def _on_item_double_clicked(self, item):
        self._show_popup_for_spec(item.data(QtCore.Qt.UserRole))

    def _on_list_context_menu(self, pos):
        item = self.spec_list.itemAt(pos)
        if not item:
            return
        menu = QtWidgets.QMenu(self)
        act = menu.addAction("Open popup")
        copy_act = menu.addAction("Copy selected to clipboard")
        chosen = menu.exec_(self.spec_list.mapToGlobal(pos))
        if chosen == act:
            self._show_popup_for_spec(item.data(QtCore.Qt.UserRole))
        elif chosen == copy_act:
            self._copy_selected_to_clipboard()

    def _on_table_context_menu(self, pos):
        row = self.results_table.indexAt(pos).row()
        if row < 0:
            return
        spec_id = self.results_table.item(row,0).data(QtCore.Qt.UserRole)
        menu = QtWidgets.QMenu(self)
        act = menu.addAction("Open popup")
        copy_act = menu.addAction("Copy selected to clipboard")
        chosen = menu.exec_(self.results_table.mapToGlobal(pos))
        if chosen == act:
            spec = self._spec_by_id(spec_id)
            if spec:
                self._show_popup_for_spec(spec)
        elif chosen == copy_act:
            self._copy_selected_to_clipboard()

    def _on_table_double_clicked(self, item):
        spec_id = self.results_table.item(item.row(),0).data(QtCore.Qt.UserRole)
        spec = self._spec_by_id(spec_id)
        if spec:
            self._show_popup_for_spec(spec)

    def _on_table_selection(self):
        row = self.results_table.currentRow()
        if row < 0:
            return
        spec_id = self.results_table.item(row,0).data(QtCore.Qt.UserRole)
        item = self._item_map.get(spec_id)
        if item:
            self.spec_list.setCurrentItem(item, QtCore.QItemSelectionModel.SelectCurrent)
            self._update_plot()

    def _copy_selected_to_clipboard(self):
        channel = self.channel_combo.currentText()
        if not channel:
            return
        items = self._selected_items() or self._checked_items()
        if not items:
            return
        lines = []
        for it in items:
            spec = it.data(QtCore.Qt.UserRole)
            if not spec:
                continue
            V = np.asarray(spec.get('V', []), dtype=float)
            ch = np.asarray((spec.get('channels') or {}).get(channel, []), dtype=float)
            if V.size == 0 or ch.size == 0:
                continue
            lines.append(f"# {Path(spec.get('path','')).name}  ({spec.get('x','?')}/{spec.get('y','?')} nm)")
            lines.append(f"Bias (mV)\t{channel}")
            for v, val in zip(V * 1000.0, ch):
                try:
                    lines.append(f"{float(v):.9g}\t{float(val):.9g}")
                except Exception:
                    lines.append(f"{v}\t{val}")
            lines.append("")
        if lines:
            QtWidgets.QApplication.clipboard().setText("\n".join(lines))
            QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), "Copied spectra", self)

    def _spec_by_id(self, spec_id):
        for spec in self.specs:
            if str(Path(spec['path'])) == spec_id:
                return spec
        return None

    def _fit_selected(self):
        items = self._selected_items() or self._checked_items()
        self._start_fit([item.data(QtCore.Qt.UserRole) for item in items])

    def _fit_all(self):
        self._start_fit([item.data(QtCore.Qt.UserRole) for item in self._checked_items()])

    def _start_fit(self, specs):
        if not specs or self._fit_thread:
            if not specs:
                self._log("Nothing to fit.")
            return
        channel = self.channel_combo.currentText()
        self._set_busy(True, f"Fitting {len(specs)} spectra...")
        self._fit_worker = _SpectroFitWorker(specs, channel)
        self._fit_thread = QtCore.QThread(self)
        self._fit_worker.moveToThread(self._fit_thread)
        self._fit_thread.started.connect(self._fit_worker.run)
        self._fit_worker.finished.connect(self._on_fit_finished)
        self._fit_worker.finished.connect(self._fit_thread.quit)
        self._fit_thread.finished.connect(self._cleanup_fit_thread)
        self._fit_thread.start()

    def _cleanup_fit_thread(self):
        self._fit_thread.deleteLater()
        self._fit_thread = None
        self._fit_worker = None
        self._set_busy(False, "Fit ready.")

    def _on_fit_finished(self, results, logs):
        for msg in logs:
            self._log(msg)
        for res in results:
            spec = res.get('spec')
            if spec:
                self._fit_results[str(Path(spec['path']))] = res
        self._populate_results_table()
        self._update_plot()

    def _populate_results_table(self):
        rows = []
        for spec_id, res in self._fit_results.items():
            spec = res.get('spec')
            if not spec:
                continue
            xs = spec.get('x')
            ys = spec.get('y')
            rows.append((spec_id, Path(spec['path']).name,
                         "n/a" if xs is None else f"{xs:.1f}",
                         "n/a" if ys is None else f"{ys:.1f}",
                         f"{res['a']:.4g}", f"{res['a_err']:.2g}",
                         f"{res['b']:.2f}", f"{res['b_err']:.2f}",
                         f"{res['c']:.4g}", f"{res['c_err']:.2g}",
                         f"{res['rmse']:.4g}"))
        self.results_table.setRowCount(len(rows))
        for r, data in enumerate(rows):
            spec_id, name, xval, yval, a, ae, b, be, c, ce, rmse = data
            values = [name, xval, yval, a, ae, b, be, c, ce, rmse]
            for col, val in enumerate(values):
                item = QtWidgets.QTableWidgetItem(val)
                if col == 0:
                    item.setData(QtCore.Qt.UserRole, spec_id)
                self.results_table.setItem(r, col, item)

    def _export_csv(self):
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Export CSV", "spectroscopy_fit.csv", "CSV Files (*.csv)")
        if not path:
            return
        headers = ["File","X (nm)","Y (nm)","a","da","b (mV)","db","c (Hz)","dc","RMSE"]
        with open(path, 'w', newline='') as f:
            f.write(",".join(headers) + "\n")
            for row in range(self.results_table.rowCount()):
                vals = [self.results_table.item(row, col).text() if self.results_table.item(row, col) else ""
                        for col in range(self.results_table.columnCount())]
                f.write(",".join(vals) + "\n")
        self._log(f"Exported to {path}")

    def _set_busy(self, busy, message):
        self.fit_selected_btn.setEnabled(not busy)
        self.fit_all_btn.setEnabled(not busy)
        self.export_btn.setEnabled(not busy)
        if busy:
            self.status_label.setText(message)

    def _on_options_toggled(self, checked):
        self.options_toggle.setArrowType(QtCore.Qt.DownArrow if checked else QtCore.Qt.RightArrow)
        self.options_body.setVisible(checked)

    def _log(self, text):
        self.log.appendPlainText(text)
