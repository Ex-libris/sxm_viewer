"""Detail canvases and spectroscopy dialogs."""
from __future__ import annotations

import io
import itertools
import json
import math

import numpy as np
from matplotlib import patches
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from matplotlib.ticker import FuncFormatter
from matplotlib.widgets import RectangleSelector
from mpl_toolkits.axes_grid1.anchored_artists import AnchoredSizeBar
from mpl_toolkits.axes_grid1.inset_locator import inset_axes

from ..._shared import *
from ...config import *
from ...data.io import *
from ...data.spectroscopy import *
from ..thumbnails import *

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
        self._views_callback = None
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
        self._profile_label = None
        self._profile_endpoint_labels = []
        self._profile_hud_text = None
        self._profile_marker_positions = None
        self._profile_marker_domain = None
        self._profile_marker_artists = []
        self._profile_marker_drag_idx = None
        self._profile_marker_callback = None
        self._profile_marker_key = None
        self._profile_marker_positions_by_key = {}
        self._profile_marker_domain_by_key = {}
        self._profile_state_callback = None
        self._profile_state_syncing = False
        self._profile_state_deferred = False
        self._profile_update_timer = QtCore.QTimer(self)
        self._profile_update_timer.setSingleShot(True)
        self._profile_update_timer.setInterval(50)
        self._profile_update_timer.timeout.connect(self._flush_profile_updates)
        self._saved_profiles = []
        self._profile_color_cycle = itertools.cycle([
            '#ffb300', '#64b5f6', '#81c784', '#e57373',
            '#ba68c8', '#4db6ac', '#ffd54f', '#90caf9',
            '#a5d6a7', '#ff8a65', '#9575cd', '#4fc3f7',
            '#aed581', '#f06292', '#7986cb', '#4dd0e1',
            '#dce775', '#ffb74d', '#4db6ac', '#9575cd',
            '#26a69a', '#ff7043', '#29b6f6', '#9ccc65'
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
        self._view_layout = "grid"
        self.angle_enabled = False
        self.angle_pts = None  # (vx, vy, ax, ay, bx, by)
        self._angle_lines = []
        self._angle_markers = []
        self._angle_label = None
        self._angle_len_labels = []
        self._angle_patch = None
        self._angle_dragging = None
        self._angle_cids = []
        self._angle_drag_origin = None
        self.angle_callback = None

    def draw(self):
        try:
            super().draw()
        except np.linalg.LinAlgError:
            # Ignore transient singular transforms during layout updates.
            return

    def set_views(self, views, preserve_profiles: bool = False):
        state = None
        if preserve_profiles:
            try:
                state = self.export_profile_state()
            except Exception:
                state = None
        self.views = views[:]
        if not preserve_profiles:
            # whenever a new view set arrives, clear saved overlays so we don't mix files
            self._clear_saved_profile_artists(notify=False)
            self.profile_pts = None
        self._redraw()
        if preserve_profiles and state is not None:
            try:
                self.import_profile_state(state, emit=False)
            except Exception:
                pass
        if callable(self._views_callback):
            try:
                self._views_callback(self.views)
            except Exception:
                pass

    def set_view_layout(self, layout: str):
        layout = (layout or "").strip().lower()
        if layout not in ("grid", "stacked"):
            layout = "grid"
        if layout == self._view_layout:
            return
        self._view_layout = layout
        self._redraw()

    def clear_views(self):
        self.views = []
        self._redraw()

    def set_views_callback(self, cb):
        self._views_callback = cb

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
        if self._view_layout == "stacked":
            cols = 1
            rows = n
        else:
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

    def set_profile_marker_callback(self, cb):
        self._profile_marker_callback = cb

    def set_profile_state_callback(self, cb):
        self._profile_state_callback = cb

    def export_profile_state(self):
        saved = []
        for entry in self._saved_profiles:
            pts = entry.get('pts')
            if pts is None:
                continue
            saved.append({'pts': tuple(pts), 'color': entry.get('color')})
        state = {
            'active_pts': tuple(self.profile_pts) if self.profile_pts is not None else None,
            'saved': saved,
            'marker_key': self._profile_marker_key,
            'marker_positions_by_key': dict(self._profile_marker_positions_by_key),
            'marker_domain_by_key': dict(self._profile_marker_domain_by_key),
        }
        return state

    def export_profile_datasets(self):
        """Return active/saved profile datasets for external dialogs."""
        active = self._build_profile_data(self.profile_pts, color=self._active_profile_color)
        saved = []
        for entry in self._saved_profiles:
            data = entry.get('data')
            if data is None:
                data = self._build_profile_data(entry.get('pts'), color=entry.get('color'))
                entry['data'] = data
            if data:
                saved.append(data)
        return active, saved

    def import_profile_state(self, state, emit=True):
        if state is None:
            return
        if self._profile_state_syncing:
            return
        try:
            self._profile_state_syncing = True
            active_pts = state.get('active_pts')
            saved = state.get('saved') or []
            self._profile_marker_key = state.get('marker_key')
            self._profile_marker_positions_by_key = dict(state.get('marker_positions_by_key') or {})
            self._profile_marker_domain_by_key = dict(state.get('marker_domain_by_key') or {})
            if active_pts is not None:
                self._set_profile_pts(tuple(active_pts))
                if not self.profile_enabled:
                    self.enable_profile(True)
            self._clear_saved_profile_artists(notify=False)
            for entry in saved:
                pts = entry.get('pts')
                if pts is None:
                    continue
                self._add_saved_profile_from_pts(tuple(pts), entry.get('color'))
            self._ensure_profile_artists()
            self._update_profile_artists()
            self.set_profile_marker_key(self._profile_marker_key)
        finally:
            self._profile_state_syncing = False
        if emit:
            self._emit_profile_state()

    def set_profile_marker_key(self, key):
        self._profile_marker_key = key
        if key is None:
            self._profile_marker_positions = self._profile_marker_positions_by_key.get(None)
            self._profile_marker_domain = self._profile_marker_domain_by_key.get(None)
        else:
            try:
                idx = int(key)
            except Exception:
                idx = None
            if idx is None:
                self._profile_marker_positions = None
                self._profile_marker_domain = None
            else:
                self._profile_marker_positions = self._profile_marker_positions_by_key.get(idx)
                self._profile_marker_domain = self._profile_marker_domain_by_key.get(idx)
        self._update_profile_marker_artists()
        self._update_profile_hud()

    def set_profile_marker_positions(self, positions, domain=None, emit=True, profile_key=None):
        key = self._profile_marker_key if profile_key is None else profile_key
        if positions is None or len(positions) < 2:
            self._profile_marker_positions = None
            if domain is not None:
                self._profile_marker_domain = tuple(domain)
            if key is not None:
                self._profile_marker_positions_by_key.pop(key, None)
                if domain is not None:
                    self._profile_marker_domain_by_key[key] = tuple(domain)
            self._clear_profile_marker_artists()
            if emit and callable(self._profile_marker_callback):
                self._profile_marker_callback(None, None)
            return
        if domain is not None:
            self._profile_marker_domain = tuple(domain)
        self._profile_marker_positions = [float(p) for p in positions]
        if key is not None:
            self._profile_marker_positions_by_key[key] = list(self._profile_marker_positions)
            if domain is not None:
                self._profile_marker_domain_by_key[key] = tuple(domain)
        self._update_profile_marker_artists()
        if emit and callable(self._profile_marker_callback):
            self._profile_marker_callback(list(self._profile_marker_positions),
                                          tuple(self._profile_marker_domain) if self._profile_marker_domain else None)

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
                self._angle_label.set_fontsize(9 * scale)
            except Exception:
                pass
        for lbl in getattr(self, '_angle_len_labels', []):
            try:
                lbl.set_fontsize(8 * scale)
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

    def get_overview_pixmap(self):
        """Return a pixmap snapshot of the current canvas (with overlays)."""
        try:
            return self.grab()
        except Exception:
            return None

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
            self._profile_marker_positions = None
            self._profile_marker_domain = None
            self._clear_profile_hud()
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
            self._profile_endpoint_labels = self._create_endpoint_labels((x0, y0, x1, y1), color)
            self._profile_label = self._create_profile_id_label((x0, y0, x1, y1), "Active", color)
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
        for lbl in self._angle_len_labels:
            try:
                lbl.remove()
            except Exception:
                pass
        self._angle_len_labels = []
        if self._angle_patch is not None:
            try:
                self._angle_patch.remove()
            except Exception:
                pass
        self._angle_patch = None

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
        for lbl in self._angle_len_labels:
            try:
                lbl.remove()
            except Exception:
                pass
        self._angle_len_labels = []
        if self._angle_patch is not None:
            try:
                self._angle_patch.remove()
            except Exception:
                pass
            self._angle_patch = None
        if not angle_info:
            self.draw_idle()
            return
        vx, vy = self.angle_pts[0], self.angle_pts[1]
        text = f"{angle_info['angle_deg']:.1f}┬░"
        unit = angle_info.get('unit')
        color = '#f5f5f5' if self._detail_dark else '#111111'
        bbox_face = '#060606' if self._detail_dark else 'white'
        font_scale = getattr(self, '_view_font_scale', 1.0)
        self._angle_label = self.main_ax.text(
            vx, vy, text,
            color=color,
            fontsize=9 * font_scale,
            ha='center', va='center',
            bbox={'facecolor': bbox_face, 'alpha': 0.65 if self._detail_dark else 0.7, 'edgecolor': 'none', 'pad': 2},
            zorder=12)
        # position label along bisector
        vec_a = np.array([self.angle_pts[2] - vx, self.angle_pts[3] - vy], dtype=float)
        vec_b = np.array([self.angle_pts[4] - vx, self.angle_pts[5] - vy], dtype=float)
        len_a = angle_info['len_a'] or 1.0
        len_b = angle_info['len_b'] or 1.0
        bis = vec_a / max(len_a, 1e-9) + vec_b / max(len_b, 1e-9)
        if np.allclose(bis, 0):
            bis = np.array([-(vec_a[1]), vec_a[0]])
        bis = bis / (np.linalg.norm(bis) + 1e-9)
        offset = min(len_a, len_b) * 0.2
        bx = vx + bis[0] * offset
        by = vy + bis[1] * offset
        self._angle_label.set_position((bx, by))
        # draw wedge patch
        theta_a = math.degrees(math.atan2(vec_a[1], vec_a[0]))
        theta_b = math.degrees(math.atan2(vec_b[1], vec_b[0]))
        theta1, theta2 = theta_a, theta_b
        diff = (theta2 - theta1) % 360.0
        if diff > 180:
            theta1, theta2 = theta2, theta1
        radius = min(len_a, len_b) * 0.25
        radius = max(radius, 1e-3)
        wedge = patches.Wedge((vx, vy), radius, theta1, theta2,
                              facecolor=color, alpha=0.15, edgecolor='none', zorder=8)
        self._angle_patch = wedge
        self.main_ax.add_patch(wedge)
        # length labels along arms
        if unit:
            mid_a = (vx + (vec_a[0] * 0.6), vy + (vec_a[1] * 0.6))
            mid_b = (vx + (vec_b[0] * 0.6), vy + (vec_b[1] * 0.6))
            lbl_a = self.main_ax.text(mid_a[0], mid_a[1],
                                      f"{len_a:.2f} {unit}",
                                      color=color,
                                      fontsize=8 * font_scale,
                                      ha='center', va='bottom',
                                      bbox={'facecolor': bbox_face, 'alpha': 0.5 if self._detail_dark else 0.6,
                                            'edgecolor': 'none', 'pad': 1},
                                      zorder=12)
            lbl_b = self.main_ax.text(mid_b[0], mid_b[1],
                                      f"{len_b:.2f} {unit}",
                                      color=color,
                                      fontsize=8 * font_scale,
                                      ha='center', va='bottom',
                                      bbox={'facecolor': bbox_face, 'alpha': 0.5 if self._detail_dark else 0.6,
                                            'edgecolor': 'none', 'pad': 1},
                                      zorder=12)
            self._angle_len_labels = [lbl_a, lbl_b]

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
        self._clear_profile_marker_artists()
        if self._profile_label is not None:
            try:
                self._profile_label.remove()
            except Exception:
                pass
        self._profile_label = None
        for lbl in self._profile_endpoint_labels:
            try:
                lbl.remove()
            except Exception:
                pass
        self._profile_endpoint_labels = []
        self._clear_profile_hud()
        self.draw_idle()

    def _update_profile_artists(self):
        if self._profile_line is None or self._profile_p0 is None or self._profile_p1 is None:
            return
        x0, y0, x1, y1 = self.profile_pts
        self._profile_line.set_data([x0,x1],[y0,y1])
        self._profile_p0.set_data([x0],[y0])
        self._profile_p1.set_data([x1],[y1])
        self._update_profile_markers()
        self._update_profile_marker_artists()
        self.draw_idle()
        self._emit_profile()

    def _update_profile_artists_fast(self):
        if self._profile_line is None or self._profile_p0 is None or self._profile_p1 is None:
            return
        if self.profile_pts is None:
            return
        x0, y0, x1, y1 = self.profile_pts
        self._profile_line.set_data([x0, x1], [y0, y1])
        self._profile_p0.set_data([x0], [y0])
        self._profile_p1.set_data([x1], [y1])
        self._update_profile_labels()
        self.draw_idle()

    def _schedule_profile_update(self):
        if not self._profile_update_timer.isActive():
            self._profile_update_timer.start()

    def _flush_profile_updates(self):
        if not self.profile_enabled:
            return
        if self.profile_pts is None:
            return
        self._update_profile_markers()
        self._emit_profile()
        self.draw_idle()

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
        if self._profile_marker_positions_by_key:
            new_map = {}
            new_domain = {}
            moved_active = False
            for key, value in self._profile_marker_positions_by_key.items():
                if key is None:
                    new_map[key] = value
                    continue
                if key == idx:
                    new_map[None] = value
                    moved_active = True
                    continue
                if key > idx:
                    new_map[key - 1] = value
                else:
                    new_map[key] = value
            for key, value in self._profile_marker_domain_by_key.items():
                if key is None:
                    new_domain[key] = value
                    continue
                if key == idx:
                    new_domain[None] = value
                    continue
                if key > idx:
                    new_domain[key - 1] = value
                else:
                    new_domain[key] = value
            self._profile_marker_positions_by_key = new_map
            self._profile_marker_domain_by_key = new_domain
            if moved_active:
                self._profile_marker_key = None
                self._profile_marker_positions = new_map.get(None)
                self._profile_marker_domain = new_domain.get(None)
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

    def remove_saved_profile(self, index):
        """Remove a saved profile overlay by index."""
        if index is None:
            return False
        try:
            idx = int(index)
        except Exception:
            return False
        if idx < 0 or idx >= len(self._saved_profiles):
            return False
        self._remove_saved_profile(idx)
        return True

    def snapshot_active_profile(self):
        """Public hook: save the current active profile as an overlay."""
        self._snapshot_active_profile()

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

    def _clear_profile_hud(self):
        if self._profile_hud_text is not None:
            try:
                self._profile_hud_text.remove()
            except Exception:
                pass
        self._profile_hud_text = None

    def _create_profile_id_label(self, pts, text, color):
        if self.main_ax is None or pts is None:
            return None
        x0, y0, x1, y1 = pts
        xm = x0 + 0.5 * (x1 - x0)
        ym = y0 + 0.5 * (y1 - y0)
        try:
            return self.main_ax.text(
                xm, ym, text, color=color, fontsize=8,
                ha='center', va='bottom',
                bbox={'facecolor': 'black', 'alpha': 0.25, 'edgecolor': 'none', 'pad': 1.5},
                zorder=11)
        except Exception:
            return None

    def _create_endpoint_labels(self, pts, color):
        if self.main_ax is None or pts is None:
            return []
        x0, y0, x1, y1 = pts
        labels = []
        try:
            labels.append(self.main_ax.text(
                x0, y0, "A", color=color, fontsize=8,
                ha='right', va='bottom',
                bbox={'facecolor': 'black', 'alpha': 0.25, 'edgecolor': 'none', 'pad': 1.0},
                zorder=11))
            labels.append(self.main_ax.text(
                x1, y1, "B", color=color, fontsize=8,
                ha='left', va='bottom',
                bbox={'facecolor': 'black', 'alpha': 0.25, 'edgecolor': 'none', 'pad': 1.0},
                zorder=11))
        except Exception:
            return []
        return labels

    def _profile_axis_unit(self):
        if not self.views:
            return 'px'
        v0 = self.views[0]
        axis_unit = v0.get('axis_unit')
        if axis_unit:
            return axis_unit
        return 'px' if v0.get('extent') is None else 'nm'

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
        self._update_profile_marker_artists()
        self._update_profile_labels()
        self._update_profile_hud()

    def _update_profile_labels(self):
        if self.profile_pts is None or self.main_ax is None:
            return
        x0, y0, x1, y1 = self.profile_pts
        if self._profile_endpoint_labels:
            try:
                self._profile_endpoint_labels[0].set_position((x0, y0))
                self._profile_endpoint_labels[1].set_position((x1, y1))
            except Exception:
                pass
        if self._profile_label is not None:
            try:
                xm = x0 + 0.5 * (x1 - x0)
                ym = y0 + 0.5 * (y1 - y0)
                self._profile_label.set_position((xm, ym))
            except Exception:
                pass

    def _update_profile_hud(self):
        if self.main_ax is None or self.profile_pts is None:
            self._clear_profile_hud()
            return
        key = self._profile_marker_key
        pts = self._profile_marker_pts() or self.profile_pts
        data = self._build_profile_data(pts, color=self._active_profile_color)
        if not data:
            self._clear_profile_hud()
            return
        unit = data.get('axis_unit') or data.get('distance_unit') or 'px'
        length = data.get('length_nm')
        title = "Active" if key is None else f"Overlay {int(key) + 1}"
        marker_delta = None
        if self._profile_marker_positions and self._profile_marker_domain:
            marker_delta = abs(self._profile_marker_positions[1] - self._profile_marker_positions[0])
        parts = [title]
        if length is not None:
            parts.append(f"L={length:.3g} {unit}")
        if marker_delta is not None:
            parts.append(f"Δ={marker_delta:.3g} {unit}")
        text = " | ".join(parts)
        if self._profile_hud_text is None:
            self._profile_hud_text = self.main_ax.text(
                0.02, 0.98, text, transform=self.main_ax.transAxes,
                ha='left', va='top', fontsize=9,
                color="#f5f5f5",
                bbox={'facecolor': 'black', 'alpha': 0.35, 'edgecolor': 'none', 'pad': 2},
                zorder=20)
        else:
            try:
                self._profile_hud_text.set_text(text)
            except Exception:
                pass

    def _clear_profile_marker_artists(self):
        for art in self._profile_marker_artists:
            try:
                if art is not None:
                    art.remove()
            except Exception:
                pass
        self._profile_marker_artists = []
        self.draw_idle()

    def _profile_marker_points(self):
        pts = self._profile_marker_pts()
        if pts is None or self._profile_marker_positions is None:
            return []
        if not self._profile_marker_domain:
            return []
        x0, y0, x1, y1 = pts
        dom_min, dom_max = self._profile_marker_domain
        span = float(dom_max - dom_min) if dom_max != dom_min else 0.0
        if span == 0.0:
            return []
        points = []
        for pos in self._profile_marker_positions:
            t = (float(pos) - dom_min) / span
            t = max(0.0, min(1.0, t))
            px = x0 + (x1 - x0) * t
            py = y0 + (y1 - y0) * t
            points.append((px, py))
        return points

    def _profile_marker_pts(self):
        if self._profile_marker_key is None:
            return self.profile_pts
        try:
            idx = int(self._profile_marker_key)
        except Exception:
            return self.profile_pts
        if idx < 0 or idx >= len(self._saved_profiles):
            return self.profile_pts
        entry = self._saved_profiles[idx]
        return entry.get('pts') or self.profile_pts

    def _update_profile_marker_artists(self):
        pts = self._profile_marker_pts()
        if self.main_ax is None or pts is None:
            self._clear_profile_marker_artists()
            return
        points = self._profile_marker_points()
        if len(points) < 2:
            self._clear_profile_marker_artists()
            return
        for art in self._profile_marker_artists:
            try:
                if art is not None:
                    art.remove()
            except Exception:
                pass
        self._profile_marker_artists = []
        color = '#ff5252'
        x0, y0, x1, y1 = pts
        vx = x1 - x0
        vy = y1 - y0
        length = float(math.hypot(vx, vy)) if vx or vy else 0.0
        if length > 0:
            nx = -vy / length
            ny = vx / length
        else:
            nx, ny = 0.0, 0.0
        tick_len = 0.03 * length if length > 0 else 0.0
        for px, py in points:
            if tick_len > 0:
                tick, = self.main_ax.plot(
                    [px - nx * tick_len, px + nx * tick_len],
                    [py - ny * tick_len, py + ny * tick_len],
                    color=color, lw=2.0, alpha=0.9, zorder=12,
                )
                self._profile_marker_artists.append(tick)
            marker, = self.main_ax.plot([px], [py], marker='o', color=color,
                                        ms=5, mec='white', mew=0.7, zorder=13)
            self._profile_marker_artists.append(marker)
        self.draw_idle()
        self._update_profile_hud()

    def _update_profile_marker_artists_fast(self):
        pts = self._profile_marker_pts()
        if self.main_ax is None or pts is None:
            return
        points = self._profile_marker_points()
        if len(points) < 2:
            return
        x0, y0, x1, y1 = pts
        vx = x1 - x0
        vy = y1 - y0
        length = float(math.hypot(vx, vy)) if vx or vy else 0.0
        if length > 0:
            nx = -vy / length
            ny = vx / length
        else:
            nx, ny = 0.0, 0.0
        tick_len = 0.03 * length if length > 0 else 0.0
        expected = len(points) * (2 if tick_len > 0 else 1)
        if len(self._profile_marker_artists) != expected:
            self._update_profile_marker_artists()
            return
        idx = 0
        for px, py in points:
            if tick_len > 0:
                try:
                    self._profile_marker_artists[idx].set_data(
                        [px - nx * tick_len, px + nx * tick_len],
                        [py - ny * tick_len, py + ny * tick_len],
                    )
                except Exception:
                    pass
                idx += 1
            try:
                self._profile_marker_artists[idx].set_data([px], [py])
            except Exception:
                pass
            idx += 1
        self.draw_idle()

    def _profile_marker_hit(self, x, y):
        points = self._profile_marker_points()
        if not points:
            return None
        min_idx = None
        min_dist = float('inf')
        for idx, (px, py) in enumerate(points):
            dist = self._pt_distance_pixels(x, y, px, py)
            if dist < min_dist:
                min_dist = dist
                min_idx = idx
        if min_dist <= 12.0:
            return min_idx
        return None

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
            meta = v0.get('meta') if isinstance(v0, dict) else None
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
                'meta': meta,
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
        overlay_idx = len(self._saved_profiles) + 1
        overlay_label = self._create_profile_id_label(pts, f"Overlay {overlay_idx}", color)
        if overlay_label is not None:
            try:
                overlay_label.set_visible(False)
            except Exception:
                pass
        endpoint_labels = self._create_endpoint_labels(pts, color)
        for lbl in endpoint_labels:
            try:
                lbl.set_visible(False)
            except Exception:
                pass
        artists = [line, p0, p1] + ticks + ([text] if text is not None else [])
        if overlay_label is not None:
            artists.append(overlay_label)
        artists += endpoint_labels
        data = self._build_profile_data(pts, color=color)
        entry = {'artists': artists, 'pts': pts, 'color': color, 'data': data,
                 'overlay_label_artist': overlay_label, 'endpoint_labels': endpoint_labels}
        if text is not None:
            entry['label_artist'] = text
            entry['label_base_size'] = base_size
        self._saved_profiles.append(entry)
        self._refresh_overlay_labels()
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
        if self._profile_marker_positions_by_key:
            new_map = {}
            new_domain = {}
            for key, value in self._profile_marker_positions_by_key.items():
                if key is None:
                    new_map[key] = value
                    continue
                if key == idx:
                    continue
                if key > idx:
                    new_map[key - 1] = value
                else:
                    new_map[key] = value
            for key, value in self._profile_marker_domain_by_key.items():
                if key is None:
                    new_domain[key] = value
                    continue
                if key == idx:
                    continue
                if key > idx:
                    new_domain[key - 1] = value
                else:
                    new_domain[key] = value
            self._profile_marker_positions_by_key = new_map
            self._profile_marker_domain_by_key = new_domain
            if self._profile_marker_key is not None:
                if self._profile_marker_key == idx:
                    self._profile_marker_key = None
                elif self._profile_marker_key > idx:
                    self._profile_marker_key -= 1
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
        self._refresh_overlay_labels()
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
        self._refresh_overlay_labels()

    def _add_saved_profile_from_pts(self, pts, color):
        if pts is None or self.main_ax is None:
            return
        pts = tuple(pts)
        color = color or next(self._profile_color_cycle)
        line, = self.main_ax.plot([pts[0], pts[2]], [pts[1], pts[3]],
                                  color=color, lw=1.5, alpha=0.7, zorder=6, linestyle='--')
        p0, = self.main_ax.plot([pts[0]], [pts[1]], marker='o', color=color,
                                ms=5, mec='black', mew=0.7, alpha=0.9, zorder=7)
        p1, = self.main_ax.plot([pts[2]], [pts[3]], marker='o', color=color,
                                ms=5, mec='black', mew=0.7, alpha=0.9, zorder=7)
        base_size = 8
        ticks, text = self._create_ticks_and_label(pts, color=color, alpha=0.7, base_size=base_size)
        overlay_idx = len(self._saved_profiles) + 1
        overlay_label = self._create_profile_id_label(pts, f"Overlay {overlay_idx}", color)
        if overlay_label is not None:
            try:
                overlay_label.set_visible(False)
            except Exception:
                pass
        endpoint_labels = self._create_endpoint_labels(pts, color)
        for lbl in endpoint_labels:
            try:
                lbl.set_visible(False)
            except Exception:
                pass
        artists = [line, p0, p1] + ticks + ([text] if text is not None else [])
        if overlay_label is not None:
            artists.append(overlay_label)
        artists += endpoint_labels
        data = self._build_profile_data(pts, color=color)
        entry = {'artists': artists, 'pts': pts, 'color': color, 'data': data,
                 'overlay_label_artist': overlay_label, 'endpoint_labels': endpoint_labels}
        if text is not None:
            entry['label_artist'] = text
            entry['label_base_size'] = base_size
        self._saved_profiles.append(entry)
        self._refresh_overlay_labels()

    def _refresh_overlay_labels(self):
        for idx, entry in enumerate(self._saved_profiles, 1):
            label = entry.get('overlay_label_artist')
            if label is not None:
                try:
                    label.set_text(f"Overlay {idx}")
                except Exception:
                    pass

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
                    line.set_alpha(0.35)
            except Exception:
                pass
            for label in entry.get('endpoint_labels', []) or []:
                try:
                    label.set_visible(idx == self._highlighted_overlay)
                except Exception:
                    pass
            label_artist = entry.get('overlay_label_artist')
            if label_artist is not None:
                try:
                    label_artist.set_visible(idx == self._highlighted_overlay)
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
        marker_idx = self._profile_marker_hit(x, y)
        if marker_idx is not None:
            self._profile_marker_drag_idx = marker_idx
            self._dragging = None
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
        if not self.profile_enabled or event.inaxes is None or event.inaxes is not self.main_ax:
            return
        x, y = event.xdata, event.ydata
        if x is None or y is None:
            return
        if self._profile_marker_drag_idx is not None:
            if not self._profile_marker_domain or self._profile_marker_positions is None:
                return
            pts = self._profile_marker_pts()
            if pts is None:
                return
            x0, y0, x1, y1 = pts
            vx = x1 - x0
            vy = y1 - y0
            denom = vx * vx + vy * vy
            if denom <= 1e-12:
                return
            t = ((x - x0) * vx + (y - y0) * vy) / denom
            t = max(0.0, min(1.0, t))
            dom_min, dom_max = self._profile_marker_domain
            pos = dom_min + t * (dom_max - dom_min)
            self._profile_marker_positions[self._profile_marker_drag_idx] = pos
            self._update_profile_marker_artists_fast()
            if self._profile_marker_key is not None:
                self._profile_marker_positions_by_key[self._profile_marker_key] = list(self._profile_marker_positions)
            else:
                self._profile_marker_positions_by_key[None] = list(self._profile_marker_positions)
            if callable(self._profile_marker_callback):
                self._profile_marker_callback(list(self._profile_marker_positions), tuple(self._profile_marker_domain))
            self._schedule_profile_update()
            return
        if self._dragging is None:
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
        self._update_profile_artists_fast()
        self._schedule_profile_update()

    def _on_release(self, event):
        if not self.profile_enabled:
            return
        self._dragging = None
        self._line_drag_origin = None
        self._profile_marker_drag_idx = None
        self._flush_profile_updates()
        if self._profile_state_deferred:
            self._profile_state_deferred = False
            self._flush_profile_state()

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
            self._emit_profile_state()
            return
        active = self._build_profile_data(self.profile_pts, color=self._active_profile_color)
        if active:
            ref = active.get('x_nm') if active.get('x_nm') is not None else active.get('x_px')
            if ref is not None:
                try:
                    ref_arr = np.asarray(ref, dtype=float)
                    if ref_arr.size:
                        self._profile_marker_domain = (float(ref_arr.min()), float(ref_arr.max()))
                        if self._profile_marker_key is not None:
                            self._profile_marker_domain_by_key[self._profile_marker_key] = self._profile_marker_domain
                        else:
                            self._profile_marker_domain_by_key[None] = self._profile_marker_domain
                        self._update_profile_marker_artists()
                except Exception:
                    pass
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
        self._emit_profile_state()

    def _emit_profile_state(self):
        if self._profile_state_syncing:
            return
        if not callable(self._profile_state_callback):
            return
        if self._dragging is not None or self._profile_marker_drag_idx is not None:
            self._profile_state_deferred = True
            return
        try:
            self._profile_state_callback(self.export_profile_state())
        except Exception:
            pass

    def _flush_profile_state(self):
        if self._profile_state_syncing:
            return
        if not callable(self._profile_state_callback):
            return
        try:
            self._profile_state_callback(self.export_profile_state())
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
        copy_svg_act = menu.addAction("Copy view as SVG (vector)")
        save_act = menu.addAction("Save image as...")
        save_svg_act = menu.addAction("Save view as SVG...")
        save_pdf_act = menu.addAction("Save view as PDF...")
        chosen = menu.exec_(event.guiEvent.globalPos())
        if chosen == copy_act:
            self._copy_view_to_clipboard(view)
        elif chosen == copy_svg_act:
            self._copy_view_as_svg(view)
        elif chosen == save_act:
            self._save_view_to_file(view)
        elif chosen == save_svg_act:
            self._save_view_vector(view, "svg")
        elif chosen == save_pdf_act:
            self._save_view_vector(view, "pdf")

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

    def _copy_view_as_svg(self, view):
        try:
            fig = self._render_view_figure(view)
            buf = io.BytesIO()
            fig.savefig(buf, format="svg", bbox_inches="tight", pad_inches=0.02)
            svg_bytes = buf.getvalue()
            mime = QtCore.QMimeData()
            mime.setData("image/svg+xml", svg_bytes)
            QtWidgets.QApplication.clipboard().setMimeData(mime)
            if callable(self._copy_feedback_handler):
                self._copy_feedback_handler(view)
        except Exception:
            pass

    def _save_view_vector(self, view, fmt):
        fmt = (fmt or "").strip().lower()
        if fmt not in ("svg", "pdf"):
            return
        try:
            title = view.get('title') or 'view'
            default = f"{title}.{fmt}"
            label = "SVG Files (*.svg)" if fmt == "svg" else "PDF Files (*.pdf)"
            path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save view", default, label)
            if not path:
                return
            if not path.lower().endswith(f".{fmt}"):
                path = f"{path}.{fmt}"
            fig = self._render_view_figure(view)
            fig.savefig(path, format=fmt, bbox_inches="tight", pad_inches=0.02)
            try:
                QtGui.QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(path))
            except Exception:
                pass
        except Exception:
            QtWidgets.QMessageBox.warning(self, "Save view", "Unable to save vector image.")

    def _render_view_figure(self, view):
        fig = Figure(figsize=(6, 6))
        ax = fig.add_subplot(1, 1, 1)
        arr = np.asarray(view.get('arr'))
        flip = bool(view.get('relative_axes'))
        if flip:
            arr_plot = np.flipud(arr)
        else:
            arr_plot = arr
        extent = view.get('extent', None)
        cmap = view.get('cmap', 'viridis')
        origin = 'lower' if flip else 'upper'
        if extent is None:
            im = ax.imshow(arr_plot, origin=origin, interpolation='nearest', cmap=cmap)
        else:
            im = ax.imshow(arr_plot, extent=extent, origin=origin, interpolation='nearest', aspect='equal', cmap=cmap)
        ax.set_autoscale_on(False)
        cbar_label = view.get('colorbar_label') or view.get('unit', '')
        cbar = None
        if cbar_label:
            cbar = fig.colorbar(im, ax=ax, fraction=0.08, pad=0.02)
            cbar.set_label(cbar_label)
        title = view.get('title', '')
        if title:
            ax.set_title(title, fontsize=9)
        ax.tick_params(labelsize=8)
        self._style_export_figure(fig, ax, cbar)
        try:
            fig.tight_layout()
        except Exception:
            pass
        return fig

    def _style_export_figure(self, fig, ax, cbar):
        dark = bool(self._detail_dark)
        fig_face = '#111217' if dark else '#ffffff'
        ax_face = '#14161c' if dark else '#ffffff'
        text_color = '#f5f5f5' if dark else '#111111'
        grid_color = '#4f5a64' if dark else '#9a9a9a'
        try:
            fig.set_facecolor(fig_face)
        except Exception:
            pass
        try:
            ax.set_facecolor(ax_face)
            ax.tick_params(colors=text_color, labelcolor=text_color)
            ax.xaxis.label.set_color(text_color)
            ax.yaxis.label.set_color(text_color)
            for spine in ax.spines.values():
                spine.set_color(text_color)
            if self._detail_grid:
                ax.grid(True, color=grid_color, alpha=0.3, linewidth=0.6)
            else:
                ax.grid(False)
        except Exception:
            pass
        if cbar is not None:
            try:
                cbar.ax.tick_params(colors=text_color, labelcolor=text_color)
                cbar.ax.yaxis.label.set_color(text_color)
                cbar.ax.xaxis.label.set_color(text_color)
                cbar.outline.set_edgecolor(text_color)
            except Exception:
                pass
        scale = max(0.6, min(1.8, getattr(self, '_view_font_scale', 1.0)))
        tick_size = 8 * scale
        label_size = 10 * scale
        title_size = 9 * scale
        try:
            ax.tick_params(labelsize=tick_size)
            ax.xaxis.label.set_fontsize(label_size)
            ax.yaxis.label.set_fontsize(label_size)
            ax.title.set_fontsize(title_size)
        except Exception:
            pass
        if cbar is not None:
            try:
                cbar.ax.tick_params(labelsize=tick_size)
                cbar.ax.yaxis.label.set_fontsize(label_size)
                cbar.ax.xaxis.label.set_fontsize(label_size)
            except Exception:
                pass

    def _start_drag(self, view, qimg=None):
        try:
            if qimg is None:
                qimg = self._view_to_qimage(view)
            pix = QtGui.QPixmap.fromImage(qimg)
            drag = QtGui.QDrag(self)
            mime = QtCore.QMimeData()
            mime.setImageData(qimg)
            try:
                meta = view.get('meta') or {}
                payload = {
                    'file_path': meta.get('file_path'),
                    'channel_index': meta.get('channel_index'),
                    'cmap': view.get('cmap'),
                }
                if payload.get('file_path') is not None and payload.get('channel_index') is not None:
                    mime.setData('application/x-sxm-view', json.dumps(payload).encode('utf-8'))
            except Exception:
                pass
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

class SafeFigureCanvas(FigureCanvas):
    def draw(self):
        try:
            super().draw()
        except np.linalg.LinAlgError:
            # Ignore transient singular transforms during layout updates.
            return
