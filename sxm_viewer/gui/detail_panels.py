"""Detail canvases and spectroscopy dialogs."""
from __future__ import annotations

import itertools

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
        self.setFocusPolicy(QtCore.Qt.StrongFocus)
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
        self._profile_ticks = []
        self._profile_info_text = None
        self._saved_profiles = []
        self._profile_color_cycle = itertools.cycle([
            '#03a9f4', '#8bc34a', '#e91e63', '#ff7043',
            '#673ab7', '#009688', '#cddc39', '#f06292',
            '#00acc1', '#9ccc65', '#5c6bc0', '#ec407a'
        ])
        self._line_drag_origin = None
        self._active_profile_color = '#fbc02d'
        self._highlighted_overlay = None
        self._cids = []
        self._base_click_cid = self.mpl_connect('button_press_event', self._on_base_click)
        self._dragging = None  # 'p0' or 'p1'
        self.main_ax = None
        self.profile_callback = None  # callable(active_dataset, saved_datasets)
        self._profile_highlight_cb = None
        self._profile_label_scale = 1.0
        self._view_font_scale = 1.0
        self._detail_dark = False
        self._detail_grid = False
        self._colorbars = []
        self.angle_enabled = False
        self.angle_pts = None  # (vx, vy, ax, ay, bx, by)
        self._angle_lines = []
        self._angle_markers = []
        self._angle_label = None
        self._angle_dragging = None
        self._angle_cids = []
        self._angle_drag_origin = None
        self.angle_callback = None

    def set_views(self, views):
        self.views = views[:]
        # whenever a new view set arrives, clear saved overlays so we don't mix files
        self._clear_saved_profile_artists(notify=False)
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
        self._colorbars = []
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
            flip = bool(v.get('relative_axes'))
            if flip:
                arr_plot = np.flipud(arr)
            else:
                arr_plot = arr
            extent = v.get('extent', None)
            cmap = v.get('cmap', 'viridis')
            origin = 'lower' if flip else 'upper'
            if extent is None:
                im = ax.imshow(arr_plot, origin=origin, interpolation='nearest', cmap=cmap)
            else:
                im = ax.imshow(arr_plot, extent=extent, origin=origin, interpolation='nearest', aspect='equal', cmap=cmap)
            ax.set_autoscale_on(False)
            cbar_label = v.get('colorbar_label') or v.get('unit', '')
            if cbar_label:
                cbar = self.fig.colorbar(im, ax=ax, fraction=0.08, pad=0.02)
                cbar.set_label(cbar_label)
                self._colorbars.append(cbar)
            title = v.get('title', '')
            ax.set_title(title, fontsize=9)
            ax.tick_params(labelsize=8)
        try: self.fig.tight_layout()
        except Exception: pass
        self._apply_view_theme()
        self._apply_view_font_scale()
        # if profile mode is enabled, (re)create artists on main ax
        if self.profile_enabled:
            self._ensure_profile_artists()
        self.draw()

    # ---------- Interactive profile helpers ----------
    def set_profile_callback(self, cb):
        self.profile_callback = cb

    def set_profile_highlight_callback(self, cb):
        self._profile_highlight_cb = cb

    def set_profile_label_scale(self, scale):
        try:
            scale = float(scale)
        except Exception:
            return
        scale = max(0.6, min(2.5, scale))
        if abs(scale - self._profile_label_scale) <= 1e-3:
            return
        self._profile_label_scale = scale
        self._update_profile_markers()
        for entry in self._saved_profiles:
            text = entry.get('label_artist')
            base = entry.get('label_base_size', 8.0)
            if text is not None:
                try:
                    text.set_fontsize(base * self._profile_label_scale)
                except Exception:
                    pass
        self.draw_idle()

    def set_detail_theme(self, *, dark=None, grid=None):
        changed = False
        if dark is not None and bool(dark) != self._detail_dark:
            self._detail_dark = bool(dark)
            changed = True
        if grid is not None and bool(grid) != self._detail_grid:
            self._detail_grid = bool(grid)
            changed = True
        if changed:
            self._apply_view_theme()

    def set_angle_callback(self, cb):
        self.angle_callback = cb

    def enable_angle(self, enable: bool):
        if enable == self.angle_enabled:
            return
        self.angle_enabled = enable
        if enable:
            self._connect_angle_events()
            self._ensure_angle_artists()
            self._emit_angle()
        else:
            self._disconnect_angle_events()
            self._clear_angle_artists()
            self.angle_pts = None
            self._emit_angle()
        self.draw_idle()

    def clear_angle_measurement(self):
        self.angle_pts = None
        self._clear_angle_artists()
        if self.angle_enabled:
            self._ensure_angle_artists()
            self._emit_angle()

    def _connect_angle_events(self):
        if self._angle_cids:
            return
        self._angle_cids = [
            self.mpl_connect('button_press_event', self._on_angle_press),
            self.mpl_connect('button_release_event', self._on_angle_release),
            self.mpl_connect('motion_notify_event', self._on_angle_motion),
        ]

    def _disconnect_angle_events(self):
        for cid in self._angle_cids:
            try:
                self.mpl_disconnect(cid)
            except Exception:
                pass
        self._angle_cids = []

    def _apply_view_theme(self):
        dark = bool(self._detail_dark)
        fig_face = '#111217' if dark else '#ffffff'
        ax_face = '#14161c' if dark else '#ffffff'
        text_color = '#f5f5f5' if dark else '#111111'
        grid_color = '#4f5a64' if dark else '#9a9a9a'
        try:
            self.fig.set_facecolor(fig_face)
        except Exception:
            pass
        for ax in self.fig.axes:
            try:
                is_colorbar = ax in [cbar.ax for cbar in self._colorbars]
            except Exception:
                is_colorbar = False
            try:
                ax.set_facecolor(ax_face if not is_colorbar else fig_face)
                ax.tick_params(colors=text_color, labelcolor=text_color)
                ax.xaxis.label.set_color(text_color)
                ax.yaxis.label.set_color(text_color)
                for spine in ax.spines.values():
                    spine.set_color(text_color)
                if not is_colorbar:
                    if self._detail_grid:
                        ax.grid(True, color=grid_color, alpha=0.3, linewidth=0.6)
                    else:
                        ax.grid(False)
            except Exception:
                pass
        for cbar in getattr(self, '_colorbars', []):
            try:
                cbar.ax.tick_params(colors=text_color, labelcolor=text_color)
                cbar.ax.yaxis.label.set_color(text_color)
                cbar.ax.xaxis.label.set_color(text_color)
                cbar.outline.set_edgecolor(text_color)
            except Exception:
                pass
        if self.angle_pts:
            self._update_angle_artists()
        self.draw_idle()

    def _apply_view_font_scale(self):
        scale = max(0.6, min(1.8, getattr(self, '_view_font_scale', 1.0)))
        tick_size = 8 * scale
        label_size = 10 * scale
        title_size = 9 * scale
        for ax in self.fig.axes:
            try:
                ax.tick_params(labelsize=tick_size)
                ax.xaxis.label.set_fontsize(label_size)
                ax.yaxis.label.set_fontsize(label_size)
                ax.title.set_fontsize(title_size)
            except Exception:
                pass
        for cbar in getattr(self, '_colorbars', []):
            try:
                cbar.ax.tick_params(labelsize=tick_size)
                cbar.ax.yaxis.label.set_fontsize(label_size)
                cbar.ax.xaxis.label.set_fontsize(label_size)
            except Exception:
                pass
        self.draw_idle()
        if getattr(self, '_angle_label', None) is not None:
            try:
                self._angle_label.set_fontsize(9 * getattr(self, '_profile_label_scale', 1.0))
            except Exception:
                pass

    def _emit_angle(self):
        if not callable(self.angle_callback):
            return
        info = self._compute_angle_info()
        try:
            self.angle_callback(info)
        except Exception:
            pass

    def set_copy_feedback_handler(self, handler):
        self._copy_feedback_handler = handler

    def set_value_callback(self, cb):
        self._value_callback = cb

    def wheelEvent(self, event):
        try:
            mods = event.modifiers()
        except Exception:
            mods = QtCore.Qt.NoModifier
        if mods & QtCore.Qt.ControlModifier:
            delta = event.angleDelta().y() if hasattr(event, 'angleDelta') else 0
            if delta:
                step = 0.05 * (1 if delta > 0 else -1)
                self._view_font_scale = min(1.8, max(0.6, self._view_font_scale + step))
                self._apply_view_font_scale()
            event.accept()
            return
        super().wheelEvent(event)

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
            self._clear_saved_profile_artists(notify=False)
        self.draw_idle()

    def keyPressEvent(self, event):
        if self.profile_enabled and event is not None:
            try:
                if event.modifiers() & QtCore.Qt.ControlModifier and event.key() == QtCore.Qt.Key_Z:
                    self._undo_last_profile_snapshot()
                    return
            except Exception:
                pass
        super().keyPressEvent(event)

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
                self._set_profile_pts((x0, y0, x1, y1))
            x0, y0, x1, y1 = self.profile_pts or (x0, y0, x1, y1)
            self._set_profile_pts((x0, y0, x1, y1))
            x0, y0, x1, y1 = self.profile_pts
            color = self._active_profile_color
            self._profile_line, = self.main_ax.plot([x0,x1],[y0,y1], color=color, lw=2, alpha=0.95, zorder=9)
            self._profile_p0, = self.main_ax.plot([x0],[y0], marker='o', color=color, ms=7, mec='black', mew=1.0, zorder=10)
            self._profile_p1, = self.main_ax.plot([x1],[y1], marker='o', color=color, ms=7, mec='black', mew=1.0, zorder=10)
            self._update_profile_markers()

    def _ensure_angle_artists(self):
        if self.main_ax is None:
            return
        if self.angle_pts is None:
            try:
                xmin, xmax, ymin, ymax = self._angle_bounds()
            except Exception:
                xmin = ymin = 0.0
                xmax = ymax = 1.0
            vx = (xmin + xmax) * 0.5
            vy = (ymin + ymax) * 0.5
            ax = vx + 0.2 * (xmax - xmin)
            ay = vy
            bx = vx
            by = vy + 0.2 * (ymax - ymin)
            self.angle_pts = (vx, vy, ax, ay, bx, by)
        vx, vy, ax, ay, bx, by = self.angle_pts
        if not self._angle_lines:
            color1 = '#ffb300'
            color2 = '#00acc1'
            line1, = self.main_ax.plot([vx, ax], [vy, ay], color=color1, lw=2.4, alpha=0.95, zorder=9)
            line2, = self.main_ax.plot([vx, bx], [vy, by], color=color2, lw=2.4, alpha=0.95, zorder=9)
            vertex, = self.main_ax.plot([vx], [vy], marker='o', color='#ffffff', mec='#000000', ms=7, zorder=10)
            end_a, = self.main_ax.plot([ax], [ay], marker='o', color=color1, mec='#000000', ms=6, zorder=10)
            end_b, = self.main_ax.plot([bx], [by], marker='o', color=color2, mec='#000000', ms=6, zorder=10)
            self._angle_lines = [line1, line2]
            self._angle_markers = [vertex, end_a, end_b]
        self._update_angle_artists()

    def _clear_angle_artists(self):
        for art in self._angle_lines + self._angle_markers:
            try:
                if art is not None:
                    art.remove()
            except Exception:
                pass
        self._angle_lines = []
        self._angle_markers = []
        if self._angle_label is not None:
            try:
                self._angle_label.remove()
            except Exception:
                pass
        self._angle_label = None

    def _update_angle_artists(self):
        if not self.angle_pts or not self._angle_lines:
            return
        vx, vy, ax, ay, bx, by = self.angle_pts
        self._angle_lines[0].set_data([vx, ax], [vy, ay])
        self._angle_lines[1].set_data([vx, bx], [vy, by])
        self._angle_markers[0].set_data([vx], [vy])
        self._angle_markers[1].set_data([ax], [ay])
        self._angle_markers[2].set_data([bx], [by])
        self._update_angle_label()
        self.draw_idle()

    def _update_angle_label(self):
        angle_info = self._compute_angle_info()
        if self._angle_label is not None:
            try:
                self._angle_label.remove()
            except Exception:
                pass
            self._angle_label = None
        if not angle_info:
            return
        vx, vy = self.angle_pts[0], self.angle_pts[1]
        text = f"{angle_info['angle_deg']:.1f}°"
        unit = angle_info.get('unit')
        if unit:
            text += f" | L1={angle_info['len_a']:.2f} {unit} L2={angle_info['len_b']:.2f} {unit}"
        color = '#f5f5f5' if self._detail_dark else '#111111'
        bbox_face = '#060606' if self._detail_dark else 'white'
        self._angle_label = self.main_ax.text(
            vx, vy, text,
            color=color,
            fontsize=9 * getattr(self, '_profile_label_scale', 1.0),
            ha='left', va='bottom',
            bbox={'facecolor': bbox_face, 'alpha': 0.65 if self._detail_dark else 0.7, 'edgecolor': 'none', 'pad': 2},
            zorder=12)

    def _compute_angle_info(self):
        if not self.angle_pts:
            return None
        vx, vy, ax, ay, bx, by = self.angle_pts
        vec_a = np.array([ax - vx, ay - vy], dtype=float)
        vec_b = np.array([bx - vx, by - vy], dtype=float)
        len_a = float(np.hypot(vec_a[0], vec_a[1]))
        len_b = float(np.hypot(vec_b[0], vec_b[1]))
        if len_a < 1e-9 or len_b < 1e-9:
            angle_deg = 0.0
        else:
            cosang = float(np.clip(np.dot(vec_a, vec_b) / (len_a * len_b), -1.0, 1.0))
            angle_deg = float(np.degrees(np.arccos(cosang)))
        unit = self._profile_axis_unit()
        return {'angle_deg': angle_deg, 'len_a': len_a, 'len_b': len_b, 'unit': unit, 'vertex': (vx, vy)}

    def _clear_profile_artists(self):
        for art in (self._profile_line, self._profile_p0, self._profile_p1):
            try:
                if art is not None:
                    art.remove()
            except Exception:
                pass
        self._profile_line = self._profile_p0 = self._profile_p1 = None
        self._remove_profile_markers()
        self.draw_idle()

    def _update_profile_artists(self):
        if self._profile_line is None or self._profile_p0 is None or self._profile_p1 is None:
            return
        x0, y0, x1, y1 = self.profile_pts
        self._profile_line.set_data([x0,x1],[y0,y1])
        self._profile_p0.set_data([x0],[y0])
        self._profile_p1.set_data([x1],[y1])
        self._update_profile_markers()
        self.draw_idle()
        self._emit_profile()

    def activate_saved_profile(self, index):
        """Promote a saved overlay back to the active profile line."""
        if index is None or not self._saved_profiles:
            return False
        try:
            idx = int(index)
        except Exception:
            return False
        if idx < 0 or idx >= len(self._saved_profiles):
            return False
        entry = self._saved_profiles.pop(idx)
        # remove overlay artists from canvas
        for art in entry.get('artists', []):
            try:
                if art is not None:
                    art.remove()
            except Exception:
                pass
        # promote saved path to active profile
        pts = entry.get('pts')
        if pts is None:
            return False
        self._set_profile_pts(tuple(pts))
        self._ensure_profile_artists()
        self._update_profile_artists()
        self.draw_idle()
        self._emit_profile()
        return True

    def _remove_profile_markers(self):
        for tick in getattr(self, '_profile_ticks', []):
            try:
                if tick is not None:
                    tick.remove()
            except Exception:
                pass
        self._profile_ticks = []
        if self._profile_info_text is not None:
            try:
                self._profile_info_text.remove()
            except Exception:
                pass
        self._profile_info_text = None

    def _profile_axis_unit(self):
        if not self.views:
            return 'px'
        v0 = self.views[0]
        axis_unit = v0.get('axis_unit')
        if axis_unit:
            return axis_unit
        return 'px' if v0.get('extent') is None else 'axis units'

    def _format_profile_label(self, pts):
        x0, y0, x1, y1 = pts
        dx = abs(float(x1) - float(x0))
        dy = abs(float(y1) - float(y0))
        length = float(math.hypot(dx, dy))
        unit = self._profile_axis_unit()
        return f"L={length:.3g} {unit} | dx={dx:.3g} {unit} | dy={dy:.3g} {unit}"

    def _create_ticks_and_label(self, pts, color, alpha=0.85, base_size=9):
        ticks = []
        size = base_size * getattr(self, '_profile_label_scale', 1.0)
        try:
            fractions = (0.25, 0.5, 0.75)
            for frac in fractions:
                x = pts[0] + (pts[2] - pts[0]) * frac
                y = pts[1] + (pts[3] - pts[1]) * frac
                tick, = self.main_ax.plot(
                    [x], [y], marker='s', color=color,
                    ms=max(3.0, 4.0 * self._profile_label_scale),
                    alpha=alpha, zorder=9)
                ticks.append(tick)
            label_text = self._format_profile_label(pts)
            xm = pts[0] + (pts[2] - pts[0]) * 0.5
            ym = pts[1] + (pts[3] - pts[1]) * 0.5
            text = self.main_ax.text(
                xm, ym, label_text, color=color, fontsize=size,
                ha='center', va='center',
                bbox={'facecolor': 'black', 'alpha': 0.35, 'edgecolor': 'none', 'pad': 2},
                zorder=11)
        except Exception:
            for tick in ticks:
                try:
                    tick.remove()
                except Exception:
                    pass
            return [], None
        return ticks, text

    def _update_profile_markers(self):
        if self.profile_pts is None or self.main_ax is None:
            self._remove_profile_markers()
            return
        self._remove_profile_markers()
        ticks, text = self._create_ticks_and_label(self.profile_pts, color='yellow', alpha=0.9, base_size=9)
        self._profile_ticks = ticks
        self._profile_info_text = text

    def _build_profile_data(self, pts, color=None):
        if pts is None or not self.views:
            return None
        try:
            v0 = self.views[0]
            arr = np.asarray(v0['arr'], dtype=float)
            h, w = arr.shape
            extent = v0.get('extent', None)
            axis_unit = self._profile_axis_unit()
            x0, y0, x1, y1 = pts
            if extent is None:
                c0 = x0; r0 = y0; c1 = x1; r1 = y1
                length_nm = None
            else:
                xmin, xmax = extent[0], extent[1]
                ymin, ymax = extent[2], extent[3]
                xr = (xmax - xmin) if (xmax is not None and xmin is not None) else 1.0
                yr = (ymin - ymax) if (ymin is not None and ymax is not None) else 1.0
                c0 = (x0 - xmin) / (xr + 1e-12) * (w - 1)
                c1 = (x1 - xmin) / (xr + 1e-12) * (w - 1)
                r0 = (y0 - ymax) / (ymin - ymax + 1e-12) * (h - 1)
                r1 = (y1 - ymax) / (ymin - ymax + 1e-12) * (h - 1)
                try:
                    dx_nm = (x1 - x0); dy_nm = (y1 - y0)
                    length_nm = float(math.hypot(dx_nm, dy_nm))
                except Exception:
                    length_nm = None
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
            x_phys = None
            distance_unit = 'px'
            if length_nm is not None and n > 1:
                try:
                    scale = float(length_nm) / float(n - 1)
                    x_phys = x_px * scale
                    if axis_unit:
                        distance_unit = axis_unit
                except Exception:
                    x_phys = None
            return {
                'x_px': x_px,
                'x_nm': x_phys,
                'vals': vals,
                'length_nm': length_nm,
                'unit': unit,
                'axis_unit': axis_unit,
                'distance_unit': distance_unit if x_phys is not None else 'px',
                'color': color,
                'label': self._format_profile_label(pts),
                'relative_axes': bool(self.views[0].get('relative_axes')),
            }
        except Exception:
            return None

    def _pt_distance_pixels(self, x, y, xp, yp):
        try:
            p_scr = self.main_ax.transData.transform((x, y))
            q_scr = self.main_ax.transData.transform((xp, yp))
            dx = p_scr[0] - q_scr[0]; dy = p_scr[1] - q_scr[1]
            return (dx*dx + dy*dy) ** 0.5
        except Exception:
            return float('inf')

    def _profile_bounds(self):
        try:
            v0 = self.views[0]
            extent = v0.get('extent')
            arr = np.asarray(v0['arr'])
            h, w = arr.shape
            if extent is None:
                return (0.0, float(w - 1), 0.0, float(h - 1))
            x0, x1, y1, y0 = extent
            return (min(x0, x1), max(x0, x1), min(y0, y1), max(y0, y1))
        except Exception:
            return (-1e6, 1e6, -1e6, 1e6)

    def _clamp_profile_pts(self, x0, y0, x1, y1):
        xmin, xmax, ymin, ymax = self._profile_bounds()
        return (
            max(xmin, min(xmax, x0)),
            max(ymin, min(ymax, y0)),
            max(xmin, min(xmax, x1)),
            max(ymin, min(ymax, y1)),
        )

    def _set_profile_pts(self, pts):
        if pts is None:
            self.profile_pts = None
            return
        x0, y0, x1, y1 = pts
        self.profile_pts = self._clamp_profile_pts(x0, y0, x1, y1)

    def _angle_bounds(self):
        return self._profile_bounds()

    def _set_angle_pts(self, vx, vy, ax, ay, bx, by):
        xmin, xmax, ymin, ymax = self._angle_bounds()
        def clamp(val, lo, hi):
            return max(lo, min(hi, val))
        vx = clamp(vx, xmin, xmax); vy = clamp(vy, ymin, ymax)
        ax = clamp(ax, xmin, xmax); ay = clamp(ay, ymin, ymax)
        bx = clamp(bx, xmin, xmax); by = clamp(by, ymin, ymax)
        self.angle_pts = (vx, vy, ax, ay, bx, by)

    def _shift_pressed(self, event):
        key = getattr(event, 'key', None)
        if key and 'shift' in str(key).lower():
            return True
        gui = getattr(event, 'guiEvent', None)
        try:
            if gui is not None and gui.modifiers() & QtCore.Qt.ShiftModifier:
                return True
        except Exception:
            pass
        return False

    def _snapshot_active_profile(self):
        if self.profile_pts is None or self.main_ax is None:
            return
        pts = tuple(self.profile_pts)
        color = next(self._profile_color_cycle)
        line, = self.main_ax.plot([pts[0], pts[2]], [pts[1], pts[3]], color=color, lw=1.5, alpha=0.7, zorder=6, linestyle='--')
        p0, = self.main_ax.plot([pts[0]], [pts[1]], marker='o', color=color, ms=5, mec='black', mew=0.7, alpha=0.9, zorder=7)
        p1, = self.main_ax.plot([pts[2]], [pts[3]], marker='o', color=color, ms=5, mec='black', mew=0.7, alpha=0.9, zorder=7)
        base_size = 8
        ticks, text = self._create_ticks_and_label(pts, color=color, alpha=0.7, base_size=base_size)
        artists = [line, p0, p1] + ticks + ([text] if text is not None else [])
        data = self._build_profile_data(pts, color=color)
        entry = {'artists': artists, 'pts': pts, 'color': color, 'data': data}
        if text is not None:
            entry['label_artist'] = text
            entry['label_base_size'] = base_size
        self._saved_profiles.append(entry)
        self.draw_idle()
        self._emit_profile()

    def _distance_to_segment_pixels(self, x, y, pts):
        try:
            px, py = self.main_ax.transData.transform((x, y))
            x0, y0 = self.main_ax.transData.transform((pts[0], pts[1]))
            x1, y1 = self.main_ax.transData.transform((pts[2], pts[3]))
            vx, vy = x1 - x0, y1 - y0
            if vx == 0 and vy == 0:
                return ((px - x0)**2 + (py - y0)**2) ** 0.5
            t = ((px - x0) * vx + (py - y0) * vy) / (vx * vx + vy * vy)
            t = max(0.0, min(1.0, t))
            proj_x = x0 + t * vx
            proj_y = y0 + t * vy
            return ((px - proj_x)**2 + (py - proj_y)**2) ** 0.5
        except Exception:
            return float('inf')

    def _delete_snapshot_near(self, x, y):
        if x is None or y is None or not self._saved_profiles:
            return
        target = None
        for entry in reversed(self._saved_profiles):
            pts = entry.get('pts')
            if pts is None:
                continue
            dist = self._distance_to_segment_pixels(x, y, pts)
            if dist <= 12.0:
                target = entry
                break
        if target is None:
            return
        for art in target.get('artists', []):
            try:
                if art is not None:
                    art.remove()
            except Exception:
                pass
        self._saved_profiles.remove(target)
        self.draw_idle()
        self._emit_profile()

    def _remove_saved_profile(self, idx):
        if idx < 0 or idx >= len(self._saved_profiles):
            return
        entry = self._saved_profiles.pop(idx)
        for art in entry.get('artists', []):
            try:
                if art is not None:
                    art.remove()
            except Exception:
                pass
        self.highlight_saved_profile(None)
        if self._profile_highlight_cb:
            try:
                self._profile_highlight_cb(None)
            except Exception:
                pass
        self.draw_idle()
        self._emit_profile()

    def _overlay_index_near(self, x, y, thresh=10.0):
        if x is None or y is None or not self._saved_profiles:
            return None
        screen_pt = self.main_ax.transData.transform((x, y))
        for idx in reversed(range(len(self._saved_profiles))):
            entry = self._saved_profiles[idx]
            pts = entry.get('pts')
            if pts is None:
                continue
            try:
                x0, y0, x1, y1 = pts
                p0 = self.main_ax.transData.transform((x0, y0))
                p1 = self.main_ax.transData.transform((x1, y1))
                vx, vy = p1[0]-p0[0], p1[1]-p0[1]
                if vx == 0 and vy == 0:
                    dist = ((screen_pt[0]-p0[0])**2 + (screen_pt[1]-p0[1])**2) ** 0.5
                else:
                    t = ((screen_pt[0]-p0[0])*vx + (screen_pt[1]-p0[1])*vy)/(vx*vx+vy*vy)
                    t = max(0.0, min(1.0, t))
                    proj = (p0[0]+t*vx, p0[1]+t*vy)
                    dist = ((screen_pt[0]-proj[0])**2 + (screen_pt[1]-proj[1])**2) ** 0.5
                if dist <= thresh:
                    return idx
            except Exception:
                continue
        return None

    def _undo_last_profile_snapshot(self):
        if not self._saved_profiles:
            return
        entry = self._saved_profiles.pop()
        for art in entry.get('artists', []):
            try:
                if art is not None:
                    art.remove()
            except Exception:
                pass
        self.draw_idle()
        self._emit_profile()

    def _clear_saved_profile_artists(self, notify=False):
        for entry in self._saved_profiles:
            for art in entry.get('artists', []):
                try:
                    if art is not None:
                        art.remove()
                except Exception:
                    pass
        self._saved_profiles = []
        self._highlighted_overlay = None
        self.draw_idle()
        if notify:
            self._emit_profile()

    def clear_saved_profiles(self, notify=True):
        self._clear_saved_profile_artists(notify=notify)

    def highlight_saved_profile(self, index):
        """Update overlay styling to emphasize a selected entry."""
        self._highlighted_overlay = index if index is not None else None
        for idx, entry in enumerate(self._saved_profiles):
            artists = entry.get('artists', [])
            if not artists:
                continue
            line = artists[0]
            try:
                if idx == self._highlighted_overlay:
                    line.set_linewidth(2.5)
                    line.set_alpha(1.0)
                else:
                    line.set_linewidth(1.5)
                    line.set_alpha(0.75)
            except Exception:
                pass
        self.draw_idle()

    def _on_press(self, event):
        if not self.profile_enabled or event.inaxes is None or event.inaxes is not self.main_ax:
            return
        x, y = event.xdata, event.ydata
        if x is None or y is None:
            return
        shift_pressed = self._shift_pressed(event)
        if event.button == 3:
            overlay_idx = self._overlay_index_near(x, y)
            if overlay_idx is not None:
                self._remove_saved_profile(overlay_idx)
                return
            return
        if event.button != 1:
            return
        if self.profile_pts is None:
            self._set_profile_pts((x, y, x, y))
            self._ensure_profile_artists()
            self._dragging = 'p1'
            self._update_profile_artists()
            return
        x0, y0, x1, y1 = self.profile_pts
        d0 = self._pt_distance_pixels(x, y, x0, y0)
        d1 = self._pt_distance_pixels(x, y, x1, y1)
        thresh = 18.0  # pixels
        if d0 <= thresh or d0 <= d1:
            if d0 <= thresh:
                self._dragging = 'p0'
                self._line_drag_origin = None
                return
        if d1 <= thresh:
            self._dragging = 'p1'
            self._line_drag_origin = None
            return
        if self.profile_pts is not None:
            dist_line = self._distance_to_segment_pixels(x, y, self.profile_pts)
            if dist_line <= thresh:
                self._dragging = 'line'
                self._line_drag_origin = (x, y, self.profile_pts)
                return
        overlay_idx = self._overlay_index_near(x, y)
        if overlay_idx is not None:
            activated = self.activate_saved_profile(overlay_idx)
            if activated:
                if callable(self._profile_highlight_cb):
                    try:
                        self._profile_highlight_cb(None)
                    except Exception:
                        pass
                x0, y0, x1, y1 = self.profile_pts
            else:
                self.highlight_saved_profile(overlay_idx)
                if callable(self._profile_highlight_cb):
                    try:
                        self._profile_highlight_cb(overlay_idx)
                    except Exception:
                        pass
                self._dragging = None
                self._line_drag_origin = None
                return
        # else: start a new line from here
        if shift_pressed:
            self._snapshot_active_profile()
        self._set_profile_pts((x, y, x, y))
        self._dragging = 'p1'
        self._line_drag_origin = None
        self._update_profile_artists()

    def _on_motion(self, event):
        if not self.profile_enabled or self._dragging is None or event.inaxes is None or event.inaxes is not self.main_ax:
            return
        x, y = event.xdata, event.ydata
        if x is None or y is None:
            return
        x0, y0, x1, y1 = self.profile_pts
        if self._dragging == 'p0':
            self._set_profile_pts((x, y, x1, y1))
        elif self._dragging == 'p1':
            self._set_profile_pts((x0, y0, x, y))
        elif self._dragging == 'line' and self._line_drag_origin is not None and self.profile_pts is not None:
            sx, sy, pts = self._line_drag_origin
            dx = x - sx
            dy = y - sy
            self._set_profile_pts((pts[0] + dx, pts[1] + dy, pts[2] + dx, pts[3] + dy))
        self._update_profile_artists()

    def _on_release(self, event):
        if not self.profile_enabled:
            return
        self._dragging = None
        self._line_drag_origin = None

    def _on_angle_press(self, event):
        if not self.angle_enabled or event.inaxes is None or event.inaxes is not self.main_ax:
            return
        if event.button != 1:
            return
        x, y = event.xdata, event.ydata
        if x is None or y is None:
            return
        handle = self._angle_handle_at(x, y)
        if handle is None:
            return
        self._angle_dragging = handle
        self._angle_drag_origin = (x, y, self.angle_pts)

    def _on_angle_motion(self, event):
        if not self.angle_enabled or self._angle_dragging is None:
            return
        if event.inaxes is not self.main_ax:
            return
        x, y = event.xdata, event.ydata
        if x is None or y is None or self.angle_pts is None:
            return
        vx, vy, ax, ay, bx, by = self.angle_pts
        if self._angle_dragging == 'vertex':
            dx = x - vx
            dy = y - vy
            self._set_angle_pts(x, y, ax + dx, ay + dy, bx + dx, by + dy)
        elif self._angle_dragging == 'a':
            self._set_angle_pts(vx, vy, x, y, bx, by)
        elif self._angle_dragging == 'b':
            self._set_angle_pts(vx, vy, ax, ay, x, y)
        self._update_angle_artists()
        self._emit_angle()

    def _on_angle_release(self, event):
        if not self.angle_enabled:
            return
        self._angle_dragging = None
        self._angle_drag_origin = None

    def _angle_handle_at(self, x, y):
        if not self.angle_pts:
            return None
        vx, vy, ax, ay, bx, by = self.angle_pts
        handles = {
            'vertex': (vx, vy),
            'a': (ax, ay),
            'b': (bx, by),
        }
        min_handle = None
        min_dist = float('inf')
        for name, (hx, hy) in handles.items():
            dist = self._pt_distance_pixels(x, y, hx, hy)
            if dist < min_dist:
                min_dist = dist
                min_handle = name
        if min_dist <= 12.0:
            return min_handle
        return None
    def _emit_profile(self):
        if not callable(self.profile_callback):
            return
        active = self._build_profile_data(self.profile_pts, color=self._active_profile_color)
        saved_data = []
        for entry in self._saved_profiles:
            data = entry.get('data')
            if data is None:
                data = self._build_profile_data(entry.get('pts'), color=entry.get('color'))
                entry['data'] = data
            if data:
                saved_data.append(data)
        try:
            self.profile_callback(active, saved_data)
        except Exception:
            pass

    def _on_base_click(self, event):
        if event is None or event.inaxes is None:
            return
        ax = event.inaxes
        view = self._ax_view_map.get(ax)
        if self.profile_enabled and ax is self.main_ax:
            # avoid starting thumbnail drag/other actions while measuring profiles
            if event.button == 3:
                self._show_context_menu(event, view)
            return
        if self.angle_enabled and ax is self.main_ax:
            if event.button == 3:
                self._show_context_menu(event, view)
            return
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
    def __init__(self, active_profile, saved_profiles=None, parent=None, unit=None, y_label=None,
                 activate_overlay_callback=None, highlight_overlay_callback=None,
                 label_scale_callback=None, dark_mode=False):
        super().__init__(parent)
        self.setWindowTitle('Profile measurement')
        self.resize(640, 360)
        self._unit = unit
        self._y_label = y_label
        self._dark_background = bool(dark_mode)
        self._active = None
        self._saved = []
        self._marker_lines = []
        self._marker_positions = []
        self._marker_drag_idx = None
        self._marker_domain = (0.0, 1.0)
        self._marker_axis_scale = None
        self._marker_axis_unit = 'px'
        self._marker_display_unit = 'px'
        self._marker_reference_state = (None, None, None)
        self._marker_saved_positions = None
        self._markers_enabled = True
        self._marker_arrow = None
        self._marker_label = None
        self._marker_arrow_y = None
        self._marker_arrow_drag = None
        self._marker_cids = []
        self._label_scale_cb = label_scale_callback
        self._activate_overlay_cb = activate_overlay_callback
        self._highlight_overlay_cb = highlight_overlay_callback
        v = QtWidgets.QVBoxLayout()
        fig = Figure(figsize=(6,3))
        self.canvas = FigureCanvas(fig)
        self.ax = fig.add_subplot(111)
        self.ax_top = self.ax.twiny()
        self.ax_top.set_visible(False)
        self._relative_axes = True
        self._font_scale = 1.0
        self.ax.set_xlabel(self._axis_label('px'))
        splitter = QtWidgets.QSplitter(QtCore.Qt.Vertical)
        splitter.addWidget(self.canvas)
        info_widget = QtWidgets.QWidget()
        info_layout = QtWidgets.QVBoxLayout(info_widget)
        info_layout.setContentsMargins(0, 0, 0, 0)
        self.stats = QtWidgets.QLabel("")
        self.stats.setAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignTop)
        self.stats.setWordWrap(True)
        info_layout.addWidget(self.stats)
        self.marker_info = QtWidgets.QLabel("")
        self.marker_info.setAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignTop)
        info_layout.addWidget(self.marker_info)
        toggle_layout = QtWidgets.QHBoxLayout()
        toggle_layout.setContentsMargins(0, 0, 0, 0)
        self.marker_toggle = QtWidgets.QCheckBox("Show measurement markers")
        self.marker_toggle.setChecked(True)
        self.marker_toggle.toggled.connect(self._on_marker_toggle)
        toggle_layout.addWidget(self.marker_toggle)
        toggle_layout.addStretch(1)
        info_layout.addLayout(toggle_layout)
        theme_layout = QtWidgets.QHBoxLayout()
        theme_layout.setContentsMargins(0, 0, 0, 0)
        self.dark_bg_cb = QtWidgets.QCheckBox("Dark background")
        self.dark_bg_cb.setChecked(self._dark_background)
        self.dark_bg_cb.toggled.connect(self._on_theme_toggled)
        theme_layout.addWidget(self.dark_bg_cb)
        self.grid_cb = QtWidgets.QCheckBox("Show grid")
        self.grid_cb.setChecked(False)
        self.grid_cb.toggled.connect(self._on_theme_toggled)
        theme_layout.addWidget(self.grid_cb)
        theme_layout.addStretch(1)
        info_layout.addLayout(theme_layout)
        self.profile_list = QtWidgets.QListWidget()
        self.profile_list.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.profile_list.itemDoubleClicked.connect(self._on_profile_item_activated)
        self.profile_list.currentItemChanged.connect(self._on_profile_item_selected)
        info_layout.addWidget(QtWidgets.QLabel("Profiles"))
        info_layout.addWidget(self.profile_list, 1)
        btn_layout = QtWidgets.QHBoxLayout()
        self.copy_btn = QtWidgets.QPushButton('Copy XY')
        self.copy_btn.clicked.connect(self._copy_current_profile)
        btn_layout.addWidget(self.copy_btn)
        btn_layout.addStretch(1)
        self.close_btn = QtWidgets.QPushButton('Close')
        self.close_btn.clicked.connect(self.accept)
        btn_layout.addWidget(self.close_btn)
        info_layout.addLayout(btn_layout)
        splitter.addWidget(info_widget)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 1)
        v.addWidget(splitter)
        self.setLayout(v)
        self._marker_cids = [
            self.canvas.mpl_connect('button_press_event', self._on_marker_press),
            self.canvas.mpl_connect('button_release_event', self._on_marker_release),
            self.canvas.mpl_connect('motion_notify_event', self._on_marker_move),
        ]
        self._line_handles = []
        self._marker_reference = None
        self._apply_plot_theme()
        self.update_profiles(active_profile, saved_profiles or [], activate_overlay_callback=activate_overlay_callback)
        self._apply_font_scale()
        if callable(self._label_scale_cb):
            self._label_scale_cb(self._font_scale)

    def wheelEvent(self, event):
        try:
            modifiers = event.modifiers()
        except Exception:
            modifiers = QtCore.Qt.NoModifier
        if modifiers & QtCore.Qt.ControlModifier:
            angle = event.angleDelta().y() if hasattr(event, 'angleDelta') else 0
            if angle:
                step = 0.05 * (1 if angle > 0 else -1)
                self._font_scale = min(1.8, max(0.6, self._font_scale + step))
                self._apply_font_scale()
            event.accept()
            return
        super().wheelEvent(event)

    def _apply_font_scale(self):
        scale = max(0.6, min(1.8, getattr(self, '_font_scale', 1.0)))
        label_size = 10 * scale
        tick_size = 9 * scale
        try:
            self.ax.tick_params(axis='both', labelsize=tick_size)
            self.ax_top.tick_params(axis='both', labelsize=tick_size)
            self.ax.xaxis.label.set_fontsize(label_size)
            self.ax.yaxis.label.set_fontsize(label_size)
            self.ax_top.xaxis.label.set_fontsize(label_size)
        except Exception:
            pass
        for widget in (self.stats, self.marker_info):
            if widget is not None:
                font = widget.font()
                font.setPointSizeF(max(7.0, 9.0 * scale))
                widget.setFont(font)
        if self.profile_list is not None:
            font = self.profile_list.font()
            font.setPointSizeF(max(7.0, 9.0 * scale))
            self.profile_list.setFont(font)
        self.canvas.draw_idle()
        if self._marker_positions and len(self._marker_positions) >= 2:
            delta = abs(self._marker_positions[1] - self._marker_positions[0])
            self._update_marker_annotation(delta)
        if callable(self._label_scale_cb):
            self._label_scale_cb(self._font_scale)

    def _apply_ylabel(self, dataset):
        if dataset:
            unit_candidate = dataset.get('unit')
            if unit_candidate:
                self._unit = unit_candidate
        unit = self._unit
        if self._y_label and unit:
            self.ax.set_ylabel(f"{self._y_label} ({unit})")
        elif self._y_label:
            self.ax.set_ylabel(self._y_label)
        else:
            self.ax.set_ylabel(f"Value ({unit})" if unit else 'Value')

    def _fmt_length(self, title, length_nm):
        if length_nm is None:
            return f"{title}: N/A"
        return f"{title}: {length_nm:.3f} nm"

    def _format_stats_text(self, active, saved):
        lines = []
        if active:
            lines.append(self._fmt_length("Active", active.get('length_nm')))
        for idx, data in enumerate(saved, 1):
            lines.append(self._fmt_length(f"Overlay {idx}", data.get('length_nm')))
        return "\n".join(lines) if lines else "No profile data"

    def _clear_marker_lines(self, reset_saved=True):
        for line in self._marker_lines:
            try:
                if line:
                    line.remove()
            except Exception:
                pass
        self._marker_lines = []
        self._marker_positions = []
        self._marker_axis_scale = None
        self._marker_axis_unit = 'px'
        self._marker_display_unit = 'px'
        self._marker_reference = None
        if reset_saved:
            self._marker_saved_positions = None
        if self._marker_arrow is not None:
            try: self._marker_arrow.remove()
            except Exception: pass
            self._marker_arrow = None
        if self._marker_label is not None:
            try: self._marker_label.remove()
            except Exception: pass
            self._marker_label = None
        if reset_saved:
            self._marker_arrow_y = None
        self._marker_arrow_drag = None
        if self._markers_enabled:
            self.marker_info.setText("Markers: N/A")
        else:
            self.marker_info.setText("Markers hidden")
        self.canvas.draw_idle()

    def _reset_markers(self, ref_points, ref_length, reference_dataset=None, store_state=True):
        if store_state:
            self._marker_reference_state = (ref_points, ref_length, reference_dataset)
        self._clear_marker_lines(reset_saved=store_state)
        if not self._markers_enabled:
            return
        if ref_points is None or len(ref_points) == 0:
            self.canvas.draw_idle()
            return
        xmin = float(np.nanmin(ref_points))
        xmax = float(np.nanmax(ref_points))
        if not np.isfinite(xmin) or not np.isfinite(xmax) or xmax == xmin:
            self.canvas.draw_idle()
            return
        self._marker_domain = (xmin, xmax)
        span = xmax - xmin
        if self._marker_saved_positions and len(self._marker_saved_positions) == 2:
            raw_positions = self._marker_saved_positions
        else:
            raw_positions = [xmin + 0.3 * span, xmin + 0.7 * span]
        self._marker_positions = [self._clamp_marker(pos) for pos in raw_positions]
        line_color = '#f5f5f5' if self._dark_background else '#202020'
        colors = [line_color, line_color]
        for idx, pos in enumerate(self._marker_positions):
            line = self.ax.axvline(
                pos,
                color=colors[idx % len(colors)],
                linestyle='-',
                lw=2.2,
                alpha=0.95,
                zorder=8,
            )
            self._marker_lines.append(line)
        ref_vals = None
        if reference_dataset and reference_dataset.get('vals') is not None:
            ref_vals = np.asarray(reference_dataset['vals'], dtype=float)
        self._marker_reference = {
            'x': np.asarray(ref_points, dtype=float) if ref_points is not None else None,
            'y': ref_vals,
        }
        axis_unit = (reference_dataset or {}).get('axis_unit') or (reference_dataset or {}).get('distance_unit') or ''
        has_phys_axis = bool(reference_dataset and reference_dataset.get('x_nm') is not None)
        self._marker_axis_unit = 'phys' if has_phys_axis else 'px'
        if has_phys_axis:
            self._marker_axis_scale = None
            self._marker_display_unit = axis_unit or 'nm'
        else:
            px_count = len(reference_dataset.get('x_px')) if reference_dataset and reference_dataset.get('x_px') is not None else len(ref_points or [])
            if axis_unit and ref_length is not None and px_count > 1:
                self._marker_axis_scale = float(ref_length) / float(px_count - 1)
                self._marker_display_unit = axis_unit
            else:
                self._marker_axis_scale = None
                self._marker_display_unit = 'px'
        self._marker_saved_positions = list(self._marker_positions)
        self._update_marker_info()
        self.canvas.draw_idle()

    def _on_marker_toggle(self, checked):
        self._markers_enabled = bool(checked)
        if not self._markers_enabled:
            self._clear_marker_lines(reset_saved=False)
            return
        ref_points, ref_length, ref_dataset = getattr(self, '_marker_reference_state', (None, None, None))
        if ref_points is None:
            self._clear_marker_lines(reset_saved=False)
        else:
            self._reset_markers(ref_points, ref_length, ref_dataset, store_state=False)
        self.canvas.draw_idle()
    def _update_marker_info(self):
        if not self._markers_enabled:
            self.marker_info.setText("Markers hidden")
            return
        if len(self._marker_positions) < 2:
            self.marker_info.setText("Markers: N/A")
            if self._marker_arrow:
                try: self._marker_arrow.remove()
                except Exception: pass
                self._marker_arrow = None
            if self._marker_label:
                try: self._marker_label.remove()
                except Exception: pass
                self._marker_label = None
            return
        axis_delta = abs(self._marker_positions[1] - self._marker_positions[0])
        disp_value, disp_unit = self._format_marker_delta(axis_delta)
        info = f"Markers Δ: {disp_value:.3f} {disp_unit}"
        if self._marker_axis_scale is not None:
            info += f" ({axis_delta:.1f} px)"
        if self._marker_reference and self._marker_reference.get('y') is not None:
            v0 = self._marker_value_at(self._marker_positions[0])
            v1 = self._marker_value_at(self._marker_positions[1])
            if v0 is not None and v1 is not None:
                info += f" | values: {v0:.3g} → {v1:.3g} (Δ={abs(v1-v0):.3g})"
        self.marker_info.setText(info)
        self._remember_marker_positions()
        self._update_marker_annotation(axis_delta)

    def _remember_marker_positions(self):
        if self._marker_positions:
            self._marker_saved_positions = list(self._marker_positions)

    def _format_marker_delta(self, axis_delta):
        unit = self._marker_display_unit or 'px'
        if self._marker_axis_scale is not None:
            return axis_delta * self._marker_axis_scale, unit or 'nm'
        return axis_delta, unit or 'px'

    def _update_marker_annotation(self, axis_delta, arrow_y=None):
        if not self._markers_enabled or len(self._marker_positions) < 2:
            if self._marker_arrow:
                try: self._marker_arrow.remove()
                except Exception: pass
                self._marker_arrow = None
            if self._marker_label:
                try: self._marker_label.remove()
                except Exception: pass
                self._marker_label = None
            return
        x0, x1 = self._marker_positions
        xmin, xmax = min(x0, x1), max(x0, x1)
        y_min, y_max = self.ax.get_ylim()
        if arrow_y is None:
            y_level = self._marker_arrow_y
        else:
            y_level = arrow_y
        if y_level is None:
            y_level = y_min + 0.05 * (y_max - y_min)
        y_level = max(y_min + 0.01*(y_max-y_min), min(y_max - 0.01*(y_max-y_min), y_level))
        self._marker_arrow_y = y_level
        if self._marker_arrow is not None:
            try: self._marker_arrow.remove()
            except Exception: pass
        if self._marker_label is not None:
            try: self._marker_label.remove()
            except Exception: pass
        arrow_color = "#f5f5f5" if self._dark_background else "#111111"
        arrow = self.ax.annotate(
            "",
            xy=(xmax, y_level),
            xytext=(xmin, y_level),
            arrowprops=dict(arrowstyle="<->", color=arrow_color, lw=1.8),
            annotation_clip=False,
        )
        display_value, display_unit = self._format_marker_delta(axis_delta)
        text = f"{display_value:.3f} {display_unit}"
        label_size = 9.0 * getattr(self, '_font_scale', 1.0)
        bbox_face = "#050506" if self._dark_background else "white"
        label = self.ax.text(
            (xmin + xmax) / 2.0,
            y_level + 0.02 * (y_max - y_min),
            text,
            color=arrow_color,
            ha="center",
            va="bottom",
            fontsize=label_size,
            bbox=dict(boxstyle="round,pad=0.2", facecolor=bbox_face,
                      alpha=0.7 if not self._dark_background else 0.6, edgecolor="none"),
        )
        self._marker_arrow = arrow
        self._marker_label = label
        self.canvas.draw_idle()

    def _marker_value_at(self, pos):
        if not self._marker_reference:
            return None
        x = self._marker_reference.get('x')
        y = self._marker_reference.get('y')
        if x is None or y is None or len(x) == 0:
            return None
        if pos <= x[0]:
            return float(y[0])
        if pos >= x[-1]:
            return float(y[-1])
        idx = np.searchsorted(x, pos) - 1
        idx = np.clip(idx, 0, len(x) - 2)
        x0, x1 = x[idx], x[idx + 1]
        y0, y1 = y[idx], y[idx + 1]
        if x1 == x0:
            return float(y0)
        t = (pos - x0) / (x1 - x0)
        return float(y0 + t * (y1 - y0))

    def _clamp_marker(self, val):
        lo, hi = self._marker_domain
        return min(max(val, lo), hi)

    def _select_marker_index(self, xdata):
        if not self._marker_positions:
            return None
        distances = []
        for pos in self._marker_positions:
            distances.append(abs(pos - xdata))
        idx = int(np.argmin(distances))
        domain = self._marker_domain[1] - self._marker_domain[0]
        tol = max(1e-6, 0.03 * domain)
        if distances[idx] <= tol:
            return idx
        return None

    def _event_xdata_main(self, event):
        if event is None:
            return None
        x = event.xdata
        if x is None:
            return None
        if event.inaxes is self.ax or event.inaxes is None:
            return x
        if event.inaxes is self.ax_top:
            try:
                px = event.inaxes.transData.transform((x, 0))
                x_main, _ = self.ax.transData.inverted().transform(px)
                return x_main
            except Exception:
                return x
        return None

    def _on_marker_press(self, event):
        if not self._markers_enabled:
            return
        if event.button != 1:
            return
        if event.inaxes not in (self.ax, self.ax_top):
            return
        if self._arrow_hit_test(event):
            self._start_arrow_drag(event)
            return
        x = self._event_xdata_main(event)
        if x is None:
            return
        idx = self._select_marker_index(x)
        if idx is None:
            if not self._marker_positions:
                return
            idx = int(np.argmin([abs(pos - x) for pos in self._marker_positions]))
        self._marker_drag_idx = idx
        new_pos = self._clamp_marker(x)
        self._marker_positions[idx] = new_pos
        line = self._marker_lines[idx]
        line.set_xdata([new_pos, new_pos])
        self._update_marker_info()
        self.canvas.draw_idle()

    def _on_marker_move(self, event):
        if not self._markers_enabled:
            return
        if self._marker_drag_idx is None and self._marker_arrow_drag is None:
            return
        if event.inaxes not in (self.ax, self.ax_top):
            return
        if self._marker_arrow_drag is not None:
            if event.inaxes is not self.ax:
                return
            if event.ydata is None:
                return
            y_min, y_max = self.ax.get_ylim()
            target = event.ydata + self._marker_arrow_drag
            y_level = max(y_min + 0.01*(y_max-y_min), min(y_max - 0.01*(y_max-y_min), target))
            self._marker_arrow_y = y_level
            axis_delta = abs(self._marker_positions[1] - self._marker_positions[0]) if len(self._marker_positions) >= 2 else 0.0
            self._update_marker_annotation(axis_delta, arrow_y=y_level)
            return
        x = self._event_xdata_main(event)
        if x is None:
            return
        if self._marker_drag_idx is not None:
            new_pos = self._clamp_marker(x)
            self._marker_positions[self._marker_drag_idx] = new_pos
            line = self._marker_lines[self._marker_drag_idx]
            line.set_xdata([new_pos, new_pos])
            self._update_marker_info()
            self.canvas.draw_idle()

    def _on_marker_release(self, event):
        if not self._markers_enabled:
            return
        self._marker_drag_idx = None
        self._marker_arrow_drag = None
        self._remember_marker_positions()

    def _arrow_hit_test(self, event):
        if self._marker_arrow is None or len(self._marker_positions) < 2:
            return False
        if event.inaxes is not self.ax:
            return False
        if event.xdata is None or event.ydata is None:
            return False
        x0, x1 = sorted(self._marker_positions)
        span = max(1e-12, x1 - x0)
        if not (x0 - 0.02 * span <= event.xdata <= x1 + 0.02 * span):
            return False
        y_min, y_max = self.ax.get_ylim()
        y_level = self._marker_arrow_y
        if y_level is None:
            y_level = y_min + 0.05 * (y_max - y_min)
        tol = 0.12 * (y_max - y_min)
        return abs(event.ydata - y_level) <= tol

    def _start_arrow_drag(self, event):
        if event.ydata is None:
            return
        y_min, y_max = self.ax.get_ylim()
        y_level = self._marker_arrow_y
        if y_level is None:
            y_level = y_min + 0.05 * (y_max - y_min)
        self._marker_arrow_drag = y_level - event.ydata

    def _axis_label(self, unit):
        unit = unit or 'px'
        return f"l ({unit})"

    def update_profiles(self, active_profile, saved_profiles=None, activate_overlay_callback=None,
                         highlight_overlay_callback=None):
        saved_profiles = saved_profiles or []
        self._active = active_profile
        self._saved = saved_profiles
        if activate_overlay_callback is not None:
            self._activate_overlay_cb = activate_overlay_callback
        if highlight_overlay_callback is not None:
            self._highlight_overlay_cb = highlight_overlay_callback
        reference = active_profile or (saved_profiles[0] if saved_profiles else None)
        self._relative_axes = bool((reference or {}).get('relative_axes', True))
        self._line_handles = []
        datasets = []
        if active_profile:
            datasets.append(('Active', active_profile, True))
        for idx, data in enumerate(saved_profiles, 1):
            datasets.append((f"Overlay {idx}", data, False))
        if not datasets:
            self.stats.setText("No profile data")
            self._clear_marker_lines()
            self.canvas.draw_idle()
            return
        axis_label_unit = 'px'
        if reference and reference.get('x_nm') is not None:
            axis_label_unit = reference.get('axis_unit') or reference.get('distance_unit') or 'nm'
        elif datasets:
            candidate = datasets[0][1]
            if candidate.get('x_nm') is not None:
                axis_label_unit = candidate.get('axis_unit') or 'nm'
        self.ax.clear()
        self.ax_top.clear()
        self.ax_top.set_visible(False)
        self.ax.set_xlabel(self._axis_label(axis_label_unit))
        self._apply_ylabel(reference)
        ref_points = None
        ref_length = None
        marker_dataset = active_profile if active_profile else (saved_profiles[0] if saved_profiles else None)
        for label, data, is_active in datasets:
            x = data.get('x_nm')
            if x is None:
                x = data.get('x_px')
            y = data.get('vals')
            if x is None or y is None:
                continue
            color = data.get('color') or ('#ffd54f' if is_active else '#80cbc4')
            lw = 2.0 if is_active else 1.2
            alpha = 0.95 if is_active else 0.75
            line, = self.ax.plot(x, y, color=color, lw=lw, alpha=alpha, label=label)
            self._line_handles.append(line)
            if is_active:
                ref_points = x
                ref_length = data.get('length_nm')
        if marker_dataset is not None:
            ref_points = marker_dataset.get('x_nm') if marker_dataset.get('x_nm') is not None else marker_dataset.get('x_px')
            ref_length = marker_dataset.get('length_nm')
        elif ref_points is None and datasets:
            data0 = datasets[0][1]
            ref_points = data0.get('x_nm') if data0.get('x_nm') is not None else data0.get('x_px')
            ref_length = datasets[0][1].get('length_nm')
        self.ax.relim(); self.ax.autoscale_view()
        if len(datasets) > 1:
            try:
                self.ax.legend(fontsize=8, loc='upper right')
            except Exception:
                pass
        self._apply_plot_theme()
        self.stats.setText(self._format_stats_text(active_profile, saved_profiles))
        self._populate_profile_list(active_profile, saved_profiles)
        self._reset_markers(ref_points, ref_length, reference_dataset=marker_dataset)
        self._apply_font_scale()
        self.canvas.draw_idle()

    def _populate_profile_list(self, active_profile, saved_profiles):
        self.profile_list.blockSignals(True)
        self.profile_list.clear()
        target_item = None
        if active_profile:
            item = QtWidgets.QListWidgetItem(self._fmt_length("Active", active_profile.get('length_nm')))
            item.setData(QtCore.Qt.UserRole, None)
            self.profile_list.addItem(item)
            target_item = item
        for idx, data in enumerate(saved_profiles, 1):
            text = self._fmt_length(f"Overlay {idx}", data.get('length_nm'))
            item = QtWidgets.QListWidgetItem(text)
            item.setData(QtCore.Qt.UserRole, idx - 1)
            self.profile_list.addItem(item)
            if target_item is None:
                target_item = item
        if target_item:
            self.profile_list.setCurrentItem(target_item)
        self.profile_list.blockSignals(False)
        if target_item:
            self._on_profile_item_selected(target_item)

    def set_label_scale_callback(self, cb):
        self._label_scale_cb = cb
        if callable(self._label_scale_cb):
            self._label_scale_cb(self._font_scale)

    def _on_theme_toggled(self, _checked=False):
        self._dark_background = bool(self.dark_bg_cb.isChecked())
        self._apply_plot_theme()

    def _apply_plot_theme(self):
        dark = bool(self._dark_background)
        fig_face = '#111217' if dark else '#ffffff'
        ax_face = '#14161c' if dark else '#ffffff'
        text = '#f5f5f5' if dark else '#111111'
        grid_on = bool(self.grid_cb.isChecked()) if hasattr(self, 'grid_cb') else False
        grid_color = '#4f5a64' if dark else '#b0b0b0'
        try:
            self.canvas.figure.set_facecolor(fig_face)
            self.canvas.figure.set_edgecolor(fig_face)
        except Exception:
            pass
        for axis in (self.ax, self.ax_top):
            try:
                axis.set_facecolor(ax_face)
                axis.tick_params(colors=text, labelcolor=text)
                axis.xaxis.label.set_color(text)
                axis.yaxis.label.set_color(text)
                for spine in axis.spines.values():
                    spine.set_color(text)
            except Exception:
                pass
        try:
            self.ax.grid(grid_on, color=grid_color, alpha=0.35 if grid_on else 0.0)
        except Exception:
            pass
        legend = self.ax.get_legend()
        if legend is not None:
            try:
                legend.get_frame().set_facecolor(ax_face)
                legend.get_frame().set_edgecolor(text)
                for txt in legend.get_texts():
                    txt.set_color(text)
            except Exception:
                pass
        if self._marker_positions and len(self._marker_positions) >= 2:
            self._update_marker_annotation(abs(self._marker_positions[1] - self._marker_positions[0]))
        self.canvas.draw_idle()

    def select_overlay(self, idx):
        self.profile_list.blockSignals(True)
        target = None
        for i in range(self.profile_list.count()):
            item = self.profile_list.item(i)
            if item.data(QtCore.Qt.UserRole) == idx:
                target = item
                break
        if target is None and idx is None and self.profile_list.count():
            target = self.profile_list.item(0)
        if target:
            self.profile_list.setCurrentItem(target)
            self._on_profile_item_selected(target)
        self.profile_list.blockSignals(False)

    def _on_profile_item_selected(self, current, _previous=None):
        if current is None:
            return
        idx = current.data(QtCore.Qt.UserRole)
        # adjust highlight on plotted lines
        for line in self._line_handles:
            try:
                line.set_linewidth(1.2)
            except Exception:
                pass
        if idx is None:
            if self._line_handles:
                try:
                    self._line_handles[0].set_linewidth(2.4)
                except Exception:
                    pass
        else:
            line = self._line_handle_for_overlay(idx)
            if line is not None:
                try:
                    line.set_linewidth(2.4)
                except Exception:
                    pass
        self.canvas.draw_idle()
        if self._highlight_overlay_cb:
            try:
                self._highlight_overlay_cb(idx)
            except Exception:
                pass

    def _line_handle_for_overlay(self, idx):
        if idx is None:
            return None
        offset = 1 if self._active else 0
        target = idx + offset
        if target < 0 or target >= len(self._line_handles):
            return None
        return self._line_handles[target]

    def _on_profile_item_activated(self, item):
        if item is None:
            return
        idx = item.data(QtCore.Qt.UserRole)
        if idx is None:
            return
        if callable(self._activate_overlay_cb):
            self._activate_overlay_cb(idx)

    def closeEvent(self, event):
        try:
            for cid in self._marker_cids:
                self.canvas.mpl_disconnect(cid)
        except Exception:
            pass
        if self._highlight_overlay_cb:
            try:
                self._highlight_overlay_cb(None)
            except Exception:
                pass
        super().closeEvent(event)

    def _copy_current_profile(self):
        dataset = None
        current = self.profile_list.currentItem()
        idx = current.data(QtCore.Qt.UserRole) if current is not None else None
        if idx is None:
            dataset = self._active
        else:
            if idx >= 0 and idx < len(self._saved):
                dataset = self._saved[idx]
        if not dataset:
            QtWidgets.QMessageBox.information(self, "Copy profile", "No profile data available.")
            return
        x = dataset.get('x_nm')
        unit = dataset.get('axis_unit') or dataset.get('distance_unit') or 'nm'
        if x is None:
            x = dataset.get('x_px')
            unit = 'px'
        vals = dataset.get('vals')
        if x is None or vals is None:
            QtWidgets.QMessageBox.information(self, "Copy profile", "Profile data is incomplete.")
            return
        rows = [f"l ({unit})\tValue"]
        for dist, val in zip(x, vals):
            try:
                rows.append(f"{float(dist):.9g}\t{float(val):.9g}")
            except Exception:
                rows.append(f"{dist}\t{val}")
        QtWidgets.QApplication.clipboard().setText("\n".join(rows))
        QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), "Profile copied", self)

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

    def _spec_id(self, spec):
        base = str(Path(spec.get('path', '')))
        idx = spec.get('matrix_index')
        return f"{base}#m{idx}" if idx is not None else base

    def _display_name(self, spec):
        name = Path(spec.get('path', '')).name
        idx = spec.get('matrix_index')
        return f"{name} [m{idx}]" if idx is not None else name

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
        self.copy_table_btn = QtWidgets.QPushButton("Copy table")
        self.clear_sel_btn = QtWidgets.QPushButton("Clear selected")
        self.clear_all_btn = QtWidgets.QPushButton("Clear all")
        btn_row.addWidget(self.fit_selected_btn)
        btn_row.addWidget(self.fit_all_btn)
        btn_row.addWidget(self.export_btn)
        btn_row.addWidget(self.copy_btn)
        btn_row.addWidget(self.copy_table_btn)
        btn_row.addWidget(self.clear_sel_btn)
        btn_row.addWidget(self.clear_all_btn)
        right_layout.addLayout(btn_row)
        self.fit_selected_btn.clicked.connect(self._fit_selected)
        self.fit_all_btn.clicked.connect(self._fit_all)
        self.export_btn.clicked.connect(self._export_csv)
        self.copy_btn.clicked.connect(self._copy_selected_to_clipboard)
        self.copy_table_btn.clicked.connect(self._copy_table_to_clipboard)
        self.clear_sel_btn.clicked.connect(self._clear_selected)
        self.clear_all_btn.clicked.connect(self._clear_all)

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
            item = QtWidgets.QListWidgetItem(self._display_name(spec))
            item.setFlags(item.flags() | QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsSelectable)
            item.setCheckState(QtCore.Qt.Checked)
            item.setData(QtCore.Qt.UserRole, spec)
            item.setData(QtCore.Qt.UserRole + 1, self._spec_id(spec))
            self.spec_list.addItem(item)
            self._item_map[self._spec_id(spec)] = item
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
        selected_ids = {item.data(QtCore.Qt.UserRole + 1) for item in self._selected_items()}
        colors = itertools.cycle(matplotlib.cm.get_cmap('tab10').colors)
        plotted = 0
        for item in self._checked_items():
            spec = item.data(QtCore.Qt.UserRole)
            spec_id = item.data(QtCore.Qt.UserRole + 1)
            channels = spec.get('channels') or {}
            data = channels.get(channel)
            V = np.asarray(spec.get('V', []), dtype=float)
            if data is None or not V.size:
                continue
            color = next(colors)
            highlight = spec_id in selected_ids or not selected_ids
            label_txt = self._display_name(spec)
            line, = self.ax.plot(V, data, color=color, lw=2.4 if highlight else 1.2,
                                 alpha=1.0 if highlight else 0.4, label=label_txt)
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
                    name = text.get_text()
                    for spec in self.specs:
                        if self._display_name(spec) == name:
                            self._legend_map[leg_line] = self._spec_id(spec)
                            break
        self.ax.set_xlabel("Bias (mV)")
        # include unit if available
        unit = None
        # look up unit from any spec carrying this channel
        for item in self._checked_items():
            spec = item.data(QtCore.Qt.UserRole)
            if not spec:
                continue
            unit_map = spec.get('unit_map') or {}
            if channel in unit_map and unit_map[channel]:
                unit = unit_map[channel]
                break
        if unit:
            self.ax.set_ylabel(f"{channel} ({unit})")
        else:
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
            if self._display_name(spec) == name:
                return self._spec_id(spec)
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
        copy_table_act = menu.addAction("Copy table")
        clear_sel_act = menu.addAction("Clear selected")
        clear_all_act = menu.addAction("Clear all")
        chosen = menu.exec_(self.results_table.mapToGlobal(pos))
        if chosen == act:
            spec = self._spec_by_id(spec_id)
            if spec:
                self._show_popup_for_spec(spec)
        elif chosen == copy_act:
            self._copy_selected_to_clipboard()
        elif chosen == copy_table_act:
            self._copy_table_to_clipboard()
        elif chosen == clear_sel_act:
            self._clear_selected()
        elif chosen == clear_all_act:
            self._clear_all()

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
        blocks = []
        for it in items:
            spec = it.data(QtCore.Qt.UserRole)
            if not spec:
                continue
            V = np.asarray(spec.get('V', []), dtype=float)
            ch = np.asarray((spec.get('channels') or {}).get(channel, []), dtype=float)
            if V.size == 0 or ch.size == 0:
                continue
            unit_map = spec.get('unit_map') or {}
            unit = unit_map.get(channel, "")
            header_unit = f" ({unit})" if unit else ""
            block = []
            block.append(f"# {Path(spec.get('path','')).name}  ({spec.get('x','?')}/{spec.get('y','?')} nm)")
            block.append(f"Bias (mV)\t{channel}{header_unit}")
            for v, val in zip(V * 1000.0, ch):
                try:
                    block.append(f"{float(v):.9g}\t{float(val):.9g}")
                except Exception:
                    block.append(f"{v}\t{val}")
            blocks.append("\n".join(block))
        if blocks:
            QtWidgets.QApplication.clipboard().setText("\n\n".join(blocks))
            QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), "Copied spectra", self)

    def _spec_by_id(self, spec_id):
        for spec in self.specs:
            if self._spec_id(spec) == spec_id:
                return spec
        return None

    def _copy_table_to_clipboard(self):
        rows = []
        headers = ["File","X (nm)","Y (nm)","a","da","b (mV)","db","c (Hz)","dc","RMSE"]
        rows.append("\t".join(headers))
        for r in range(self.results_table.rowCount()):
            vals = []
            for c in range(self.results_table.columnCount()):
                item = self.results_table.item(r, c)
                vals.append(item.text() if item else "")
            rows.append("\t".join(vals))
        if len(rows) > 1:
            QtWidgets.QApplication.clipboard().setText("\n".join(rows))
            QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), "Copied table", self)

    def _clear_selected(self):
        removed = False
        for item in list(self._selected_items()):
            spec_id = item.data(QtCore.Qt.UserRole + 1)
            if spec_id in self._fit_results:
                self._fit_results.pop(spec_id, None)
            row = self.spec_list.row(item)
            self.spec_list.takeItem(row)
            removed = True
        if removed:
            self._update_plot()
            self._populate_results_table()
            self._update_status()

    def _clear_all(self):
        self.spec_list.clear()
        self._item_map = {}
        self._line_map.clear()
        self._legend_map.clear()
        self._fit_results = {}
        self.ax.clear()
        self.canvas.draw_idle()
        self.results_table.setRowCount(0)
        self._update_status(0)
        # also clear selection in parent so reopening starts fresh
        try:
            parent = self.parent()
            if parent and hasattr(parent, '_clear_multi_spec_selection'):
                parent._clear_multi_spec_selection()
        except Exception:
            pass

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
                self._fit_results[self._spec_id(spec)] = res
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
            rows.append((spec_id, self._display_name(spec),
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
