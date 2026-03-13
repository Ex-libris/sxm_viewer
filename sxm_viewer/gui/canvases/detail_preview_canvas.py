"""Detail canvases and spectroscopy dialogs."""
from __future__ import annotations

import io
import itertools
import json
import math
import time
import warnings
from typing import List
from pathlib import Path

import numpy as np
from matplotlib import patches
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from matplotlib.ticker import FuncFormatter
from matplotlib.widgets import RectangleSelector
from mpl_toolkits.axes_grid1.anchored_artists import AnchoredSizeBar
from mpl_toolkits.axes_grid1.inset_locator import inset_axes
from mpl_toolkits.axes_grid1 import make_axes_locatable
import matplotlib
from matplotlib.collections import LineCollection
import matplotlib.patheffects as PathEffects

from ..._shared import QtCore, QtGui, QtWidgets
from .molecular_overlay import (
    Molecule,
    MoleculePropertiesDialog,
    get_atom_color,
    get_atom_radius,
    available_atom_palettes,
)
from ..thumbnail_render import sample_array_value, array_to_qimage

try:
    from scipy import ndimage
    from scipy.spatial import ConvexHull
    _HAS_SCIPY = True
except Exception:  # pragma: no cover - optional dependency
    ndimage = None
    ConvexHull = None
    _HAS_SCIPY = False

try:
    from skimage import measure as sk_measure
    from skimage import filters as sk_filters
    from skimage import morphology as sk_morph
    _HAS_SKIMAGE = True
except Exception:  # pragma: no cover - optional dependency
    sk_measure = None
    sk_filters = None
    sk_morph = None
    _HAS_SKIMAGE = False

try:
    import cv2
    _HAS_CV2 = True
except Exception:  # pragma: no cover - optional dependency
    cv2 = None
    _HAS_CV2 = False

_FIXED_CROP_HISTORY_LIMIT = 96

class MultiPreviewCanvas(FigureCanvas):
    _RECENT_MOLECULES = []

    def __init__(self, parent=None, figsize=(6,6)):
        self.fig = Figure(figsize=figsize)
        super().__init__(self.fig)
        if parent is not None:
            self.setParent(parent)
        # Allow the canvas (and any parent dialog) to shrink freely
        # instead of being constrained by the initial figure size.
        try:
            self.setMinimumSize(0, 0)
            self.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        except Exception:
            pass
        self._compact_size_hints = False
        self._resize_reflow_threshold_px = 3
        self._last_resize_size = QtCore.QSize(-1, -1)
        self._resize_draft_timer = QtCore.QTimer(self)
        self._resize_draft_timer.setSingleShot(True)
        self._resize_draft_timer.setInterval(45)
        self._resize_draft_timer.timeout.connect(self._reflow_after_resize)
        self._resize_settle_timer = QtCore.QTimer(self)
        self._resize_settle_timer.setSingleShot(True)
        self._resize_settle_timer.setInterval(140)
        self._resize_settle_timer.timeout.connect(self._finalize_after_resize)
        self.setFocusPolicy(QtCore.Qt.StrongFocus)
        self.views = []
        self._ax_view_map = {}
        self._relative_axes_override = None
        self._suspend_zoom_restore = False
        self._image_meta = {}
        self._copy_feedback_handler = None
        self._views_callback = None
        self._drag_candidate = None  # (view, QPoint start, QImage cache)
        self._crop_callback = None  # callable(view_dict) -> None
        self._crop_start = None
        self._crop_rect = None
        self._crop_ax = None
        self._crop_square = False
        self._crop_last_ts = 0.0
        self._crop_move_throttle_ms = 12  # throttle mouse-move driven crop updates
        self._fixed_crop_template = None
        self._fixed_crop_template_bounds = None
        self._fixed_crop_template_pixel_bounds = None
        self._fixed_crop_template_view_key = None
        self._fixed_crop_template_visible = False
        self._fixed_crop_history_visible = False
        self._fixed_crop_quick_mode = False
        self._fixed_crop_history = []
        self._fixed_crop_sequence = 1
        self._fixed_crop_template_unit = "nm"
        self._fixed_crop_history_callback = None
        self._fixed_crop_template_manual_dims = None
        self._fixed_crop_history_highlight_seq = None
        self._fixed_crop_history_highlight_artists = {}
        self._double_click_callback = None  # callable(view_dict) -> None
        self._filter_menu_callback = None  # callable(menu, view, canvas)
        self._stp_export_callback = None
        self._arrange_windows_callback = None
        self._value_callback = None
        self._value_cid = self.mpl_connect('motion_notify_event', self._on_motion_value)
        self.mpl_connect('motion_notify_event', self._on_molecule_motion)
        self.mpl_connect('button_release_event', self._on_molecule_release)
        self._crop_release_cid = self.mpl_connect('button_release_event', self._on_crop_release)
        # profile (interactive line) state
        self.profile_enabled = False
        self.profile_pts = None  # (x0, y0, x1, y1) in data coords of main ax
        self._profile_line = None
        self._profile_p0 = None
        self._profile_p1 = None
        self._profile_ticks = None
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
            '#000000', '#e6194B', '#4363d8', '#3cb44b',
            '#911eb4', '#f58231', '#a9a9a9'
        ])
        self._line_drag_origin = None
        self._active_profile_color = '#fbc02d'
        self._active_profile_lw = 2.0
        self._highlighted_overlay = None
        self._cids = []
        self._base_click_cid = self.mpl_connect('button_press_event', self._on_base_click)
        self._dragging = None  # 'p0' or 'p1'
        self.main_ax = None
        self.profile_callback = None  # callable(active_dataset, saved_datasets)
        self._profile_highlight_cb = None
        self._profile_label_scale = 1.0
        self._profile_label_mode = "length"
        self._view_font_scale = 1.0
        self._colorbar_orientation = 'vertical'
        self._show_ticks = True
        self._show_colorbar = True
        self._show_title = True
        self._fit_to_canvas = False
        self._frame_fill_mode = False
        self._frame_fill_prev_state = None
        self._detail_dark = False
        self._detail_grid = False
        self._colorbars = []
        self._highlight_pulse_strength = 1.0
        self._view_layout = "grid"
        self._spectra_points = {}
        self._spectra_click_cb = None
        self._zoom_reset_limits = {}
        self.angle_enabled = False
        self.angle_pts = None  # (vx, vy, ax, ay, bx, by) for the active frame
        self._angle_frames = []
        self._active_angle_frame_idx = -1
        self._angle_frame_colors = [
            ('#ffb300', '#00acc1'),
            ('#43a047', '#1e88e5'),
            ('#d32f2f', '#7b1fa2'),
            ('#7c4dff', '#f4511e'),
            ('#00897b', '#fdd835'),
        ]
        self._show_hydrogens = True
        self._default_bond_color = (0.9, 0.9, 0.9)
        self._bond_color_mode = 'default'  # default | single | by_atoms
        self._recent_molecule_paths = []
        self._recent_molecule_cb = None
        self._angle_dragging = None
        self._angle_cids = []
        self._angle_background = None
        self._angle_blit_active = False
        self.angle_callback = None
        self.scale_bar_enabled = False
        self._scale_bar_pos = (0.94, 0.06)  # default lower right (axes coords)
        self._scale_bar_artists = []
        self._scale_bar_cids = []
        self._scale_bar_drag_start = None
        self._profile_echo_artists = []
        self._scale_bar_settings = {
            'text_color': None,
            'bar_color': None,
            'font_family': 'sans-serif'
        }
        # Outline extraction state
        self.outline_mode = False  # if True, Alt+drag will outline blobs
        self._outline_start = None
        self._outline_rect = None
        self._outline_ax = None
        self._outline_threshold = 0.8  # percentile (0-1)
        self._outline_default_color = "#ffffff"
        self._outline_default_lw = 1.6
        self._outline_default_ls = (0, (6, 4))
        self._outlines = {}  # key -> list[np.ndarray[N,2]]
        self._outline_order = []  # global order for undo [(key, idx)]
        # Molecular overlay state
        self.molecules = []
        self._molecule_drag_idx = None
        self._molecule_drag_start = None
        self._molecule_drag_start_px = None
        self._molecule_drag_mol_start = None
        self._molecule_drag_mol_angles = None
        self._molecule_drag_mode = None
        self._molecule_rotation_guide = None
        self._molecule_artists = []
        self._molecule_history = []
        self._molecule_drag_snapshot = False
        self._show_molecule_shadow = True
        self.show_molecules = True
        self.mpl_connect('scroll_event', self._on_scroll_zoom)
        self._profile_background = None
        self._active_profile_original_color = None
        self._profile_blit_active = False
        self._profile_animation_enabled = False
        # Molecule palette
        self.molecule_palette = "pymol"
        self._molecule_palette_cb = None
        self._zoom_reset_limits = {}
        # Pan/zoom state
        self._pan_active = False
        self._pan_ax = None
        self._pan_start = None
        self._pan_start_lim = None
        self._pan_last_ts = 0.0
        self._pan_throttle_ms = 16
        self.mpl_connect('scroll_event', self._on_scroll_zoom)

    def set_show_title(self, show: bool):
        """Toggle rendering of title/date overlays in views."""
        self._show_title = bool(show)
        self._redraw()

    def set_show_molecules(self, show: bool):
        """Toggle rendering of molecular overlays in views."""
        self.show_molecules = bool(show)
        self._redraw()

    def set_fit_to_canvas(self, enabled: bool):
        """If enabled, stretch view axes to fill the available canvas area."""
        self._fit_to_canvas = bool(enabled)
        self._redraw()

    def set_frame_fill_mode(self, enabled: bool):
        """Toggle a minimalist full-frame display mode for popups."""
        enabled = bool(enabled)
        if enabled == self._frame_fill_mode:
            return
        if enabled:
            self._frame_fill_prev_state = {
                "show_ticks": bool(self._show_ticks),
                "show_colorbar": bool(self._show_colorbar),
                "show_title": bool(self._show_title),
                "fit_to_canvas": bool(self._fit_to_canvas),
            }
            self._show_ticks = False
            self._show_colorbar = False
            self._show_title = False
            # Keep geometric fidelity: frame-fill hides decorations
            # but still preserves equal aspect (no stretching).
            self._fit_to_canvas = False
        else:
            prev = self._frame_fill_prev_state or {}
            self._show_ticks = bool(prev.get("show_ticks", True))
            self._show_colorbar = bool(prev.get("show_colorbar", True))
            self._show_title = bool(prev.get("show_title", True))
            self._fit_to_canvas = bool(prev.get("fit_to_canvas", False))
            self._frame_fill_prev_state = None
        self._frame_fill_mode = enabled
        self._redraw()

    def draw(self):
        try:
            super().draw()
        except np.linalg.LinAlgError:
            # Ignore transient singular transforms during layout updates.
            return

    def set_compact_size_hints(self, enabled: bool = True):
        self._compact_size_hints = bool(enabled)
        try:
            self.updateGeometry()
        except Exception:
            pass

    def minimumSizeHint(self):
        if getattr(self, "_compact_size_hints", False):
            return QtCore.QSize(24, 24)
        try:
            return super().minimumSizeHint()
        except Exception:
            return QtCore.QSize(160, 120)

    def sizeHint(self):
        if getattr(self, "_compact_size_hints", False):
            try:
                dpi = float(self.fig.get_dpi())
                width_px = int(max(220.0, min(900.0, float(self.fig.get_figwidth()) * dpi)))
                height_px = int(max(160.0, min(720.0, float(self.fig.get_figheight()) * dpi)))
                return QtCore.QSize(width_px, height_px)
            except Exception:
                return QtCore.QSize(420, 320)
        try:
            return super().sizeHint()
        except Exception:
            return QtCore.QSize(420, 320)

    def set_views(self, views, preserve_profiles: bool = False):
        state = None
        if preserve_profiles:
            try:
                state = self.export_profile_state()
            except Exception:
                state = None
        self.views = views[:]
        self._spectra_points = {}
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

    def set_relative_axes_override(self, value):
        """Force all views to use relative axes (True/False) or None to defer to view settings."""
        if value is None:
            new_val = None
        else:
            new_val = bool(value)
        if self._relative_axes_override == new_val:
            return
        self._relative_axes_override = new_val
        self.suspend_zoom_restore()
        self._redraw()

    def suspend_zoom_restore(self):
        self._suspend_zoom_restore = True

    def _use_relative_axes(self, view):
        if self._relative_axes_override is not None:
            return bool(self._relative_axes_override)
        return bool(view.get('relative_axes'))

    def _axis_meta_for_view(self, view):
        for ax, v in self._ax_view_map.items():
            if v is view:
                return ax, self._image_meta.get(ax, {})
        return None, {}

    def clear_views(self):
        self.views = []
        self._redraw()

    def set_views_callback(self, cb):
        self._views_callback = cb

    def set_crop_callback(self, cb):
        """Register a callback to receive cropped views created via drag-crop."""
        self._crop_callback = cb

    def set_double_click_callback(self, cb):
        """Register a callback to pop out the clicked view on double-click."""
        self._double_click_callback = cb

    def set_filter_menu_callback(self, cb):
        """Provide a callback to populate filter actions on the context menu."""
        self._filter_menu_callback = cb

    def set_stp_export_callback(self, cb):
        """Register callback for WSxM STP export requests."""
        self._stp_export_callback = cb

    def set_window_arrange_callback(self, cb):
        """Register callback invoked when the user requests window tiling."""
        self._arrange_windows_callback = cb

    def export_molecule_state(self):
        return [mol.to_dict() for mol in (self.molecules or [])]

    def import_molecule_state(self, state):
        self.molecules = []
        for entry in state or []:
            try:
                self.molecules.append(Molecule.from_dict(entry))
            except Exception:
                continue
        self._redraw()

    def export_angle_state(self):
        frames = []
        for frame in self._angle_frames or []:
            frames.append({
                "pts": list(frame.get("pts") or (0.0, 0.0, 1.0, 0.0, 0.0, 1.0)),
                "color_a": frame.get("color_a"),
                "color_b": frame.get("color_b"),
                "style": frame.get("style", "dots"),
            })
        return {
            "frames": frames,
            "active_idx": int(self._active_angle_frame_idx),
        }

    def import_angle_state(self, state):
        if not state:
            return
        try:
            self._clear_angle_artists()
        except Exception:
            self._angle_frames = []
            self._active_angle_frame_idx = -1
            self.angle_pts = None
        frames = []
        for entry in state.get("frames", []) or []:
            pts = entry.get("pts") or (0.0, 0.0, 1.0, 0.0, 0.0, 1.0)
            frame = {
                "pts": tuple(pts),
                "color_a": entry.get("color_a", "#ffb300"),
                "color_b": entry.get("color_b", "#00acc1"),
                "style": entry.get("style", "dots"),
                "lines": [],
                "markers": [],
                "arrows": [],
                "label": None,
                "len_labels": [],
                "patch": None,
            }
            frames.append(frame)
        self._angle_frames = frames
        if frames:
            self._set_active_angle_frame_index(int(state.get("active_idx", 0)))
            self._ensure_angle_frames()
            self._update_angle_artists()

    def set_view_clim(self, view, clim):
        """Update the color limits for a specific view and redraw while preserving overlays."""
        if not view or clim is None:
            return
        try:
            lo, hi = clim
        except Exception:
            return
        view['clim'] = (float(lo), float(hi))
        # redraw while preserving profiles/angles where possible
        try:
            self.set_views(self.views, preserve_profiles=True)
        except Exception:
            self._redraw()

    def set_spectra_click_callback(self, cb):
        """Register a callback for spectroscopy marker clicks (spec, event)."""
        self._spectra_click_cb = cb

    def resizeEvent(self, event):
        size = event.size()
        safe_size = QtCore.QSize(max(1, size.width()), max(1, size.height()))
        if safe_size != size:
            event = QtGui.QResizeEvent(safe_size, event.oldSize())
        try:
            super().resizeEvent(event)
        except ValueError:
            fallback = QtGui.QResizeEvent(
                QtCore.QSize(max(10, safe_size.width()), max(10, safe_size.height())),
                event.oldSize(),
            )
            try:
                super().resizeEvent(fallback)
            except ValueError:
                pass
        # Resize should stay lightweight during dragging and only
        # do a full redraw after resize settles.
        try:
            if getattr(self, "views", None):
                prev_size = self._last_resize_size
                if prev_size.width() <= 0 or prev_size.height() <= 0:
                    prev_size = event.oldSize()
                dw = abs(safe_size.width() - max(0, prev_size.width()))
                dh = abs(safe_size.height() - max(0, prev_size.height()))
                self._last_resize_size = QtCore.QSize(safe_size.width(), safe_size.height())
                if max(dw, dh) >= int(self._resize_reflow_threshold_px):
                    self._resize_draft_timer.start()
                self._resize_settle_timer.start()
        except Exception:
            pass

    def _reflow_after_resize(self):
        try:
            if getattr(self, "views", None):
                scale = max(0.6, min(2.5, getattr(self, "_view_font_scale", 1.0)))
                self._apply_tight_layout_safe(pad=max(0.25, 0.35 * scale))
                self.draw_idle()
        except Exception:
            pass

    def _finalize_after_resize(self):
        try:
            if getattr(self, "views", None):
                try:
                    if QtWidgets.QApplication.mouseButtons() != QtCore.Qt.NoButton:
                        self._resize_settle_timer.start()
                        return
                except Exception:
                    pass
                self._resize_draft_timer.stop()
                self._redraw()
        except Exception:
            pass

    def _redraw(self):
        # Preserve current zoom/limits per view before clearing
        current_limits = {}
        preserve_zoom = not getattr(self, "_suspend_zoom_restore", False)
        if preserve_zoom:
            try:
                for ax, v in list(self._ax_view_map.items()):
                    try:
                        current_limits[self._outline_key(v)] = (ax.get_xlim(), ax.get_ylim())
                    except Exception:
                        continue
            except Exception:
                current_limits = {}
        else:
            self._suspend_zoom_restore = False

        self.fig.clf()
        self._ax_view_map = {}
        self._image_meta = {}
        # reset zoom baselines for new axes
        self._zoom_reset_limits = {}
        self._scale_bar_artists = []
        self._colorbars = []
        self._molecule_artists = []
        self._spectra_points = {}
        for frame in self._angle_frames:
            frame['lines'] = []
            frame['markers'] = []
            frame['arrows'] = []
            frame['label'] = None
            frame['len_labels'] = []
            frame['patch'] = None
        # Reset profile artists as figure was cleared
        self._profile_line = None
        self._profile_p0 = None
        self._profile_p1 = None
        self._profile_echo_artists = []
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
            flip = self._use_relative_axes(v)
            if flip:
                arr_plot = np.flipud(arr)
            else:
                arr_plot = arr
            raw_extent = v.get('extent_raw')
            if raw_extent is None:
                raw_extent = v.get('extent')
            cmap = v.get('cmap', 'viridis')
            origin = 'lower' if flip else 'upper'
            display_extent = self._display_extent_for_view(v, raw_extent)
            aspect_mode = "auto" if self._fit_to_canvas else "equal"
            if display_extent is None:
                im = ax.imshow(
                    arr_plot,
                    origin=origin,
                    interpolation='nearest',
                    aspect=aspect_mode,
                    cmap=cmap,
                )
            else:
                im = ax.imshow(
                    arr_plot,
                    extent=display_extent,
                    origin=origin,
                    interpolation='nearest',
                    aspect=aspect_mode,
                    cmap=cmap,
                )
            try:
                self._image_meta[ax] = {
                    'extent': im.get_extent(),
                    'origin': origin,
                    'shape': arr_plot.shape,
                }
            except Exception:
                self._image_meta[ax] = {
                    'extent': display_extent,
                    'origin': origin,
                    'shape': arr_plot.shape,
                }
            clim = v.get('clim')
            if clim:
                try:
                    im.set_clim(*clim)
                except Exception:
                    pass
            ax.set_autoscale_on(False)
            cbar_label = v.get('colorbar_label') or v.get('unit', '')
            if cbar_label and self._show_colorbar:
                try:
                    divider = make_axes_locatable(ax)
                    if self._colorbar_orientation == 'horizontal':
                        cax = divider.append_axes("bottom", size="5%", pad=0.08)
                        cbar = self.fig.colorbar(im, cax=cax, orientation='horizontal')
                        cbar.set_label(cbar_label)
                        cbar.ax.xaxis.set_label_coords(0.5, 0.5)
                        cbar.ax.xaxis.label.set_horizontalalignment('center')
                        cbar.ax.xaxis.label.set_verticalalignment('center')
                    else:
                        cax = divider.append_axes("right", size="4%", pad=0.02)
                        cbar = self.fig.colorbar(im, cax=cax, orientation='vertical')
                        cbar.set_label(cbar_label)
                        cbar.ax.yaxis.set_label_coords(0.5, 0.5)
                        cbar.ax.yaxis.label.set_horizontalalignment('center')
                        cbar.ax.yaxis.label.set_verticalalignment('center')
                except Exception:
                    cbar = self.fig.colorbar(im, ax=ax, fraction=0.08, pad=0.02, orientation=self._colorbar_orientation)
                    cbar.set_label(cbar_label)
                if not self._show_ticks:
                    cbar.set_ticks([])
                self._colorbars.append(cbar)
            title = v.get('title', '')
            ax.set_title(title, fontsize=9)
            ax.tick_params(labelsize=8)
            if not self._show_ticks:
                ax.set_xticks([])
                ax.set_yticks([])
            # Restore previous zoom if available
            if preserve_zoom:
                prev_lim = current_limits.get(self._outline_key(v))
                if prev_lim:
                    try:
                        ax.set_xlim(prev_lim[0])
                        ax.set_ylim(prev_lim[1])
                    except Exception:
                        pass
            if self.scale_bar_enabled:
                try:
                    self._add_scale_bar(ax, v)
                except Exception:
                    pass
            if ax not in self._zoom_reset_limits:
                self._zoom_reset_limits[ax] = (ax.get_xlim(), ax.get_ylim())
            # Draw molecules on every view
            self._draw_molecules(ax)
            self._draw_spectra(ax)
            try:
                self._draw_outlines(ax, v)
            except Exception:
                pass
            try:
                self._draw_fixed_crop_history(ax, v)
            except Exception:
                pass
            try:
                self._render_template_overlay(ax, v)
            except Exception:
                pass
        self._apply_tight_layout_safe(pad=0.25)
        self._apply_view_theme()
        self._apply_view_font_scale()
        # if profile mode is enabled, (re)create artists on main ax
        if self.profile_enabled:
            self._ensure_profile_artists()
            self._emit_profile()
        if self.angle_enabled:
            self._ensure_angle_frames()
        self._update_highlight_artists()
        self.draw()

    def _draw_molecules(self, ax):
        if not self.show_molecules or not self.molecules:
            return

        for mol in self.molecules:
            coords = mol.get_transformed_coordinates()
            if len(coords) == 0:
                continue
            hide_h = not getattr(self, "_show_hydrogens", True)
            
            # Z-range for depth cueing
            z_vals = coords[:, 2]
            z_min = z_vals.min()
            z_range = z_vals.ptp()
            if z_range < 1e-6:
                z_range = 1.0

            lc = None
            sc = None
            shadow_sc = None
            atom_style = (mol.render_style or "shaded").lower()
            bond_style = (mol.bond_style or "default").lower()
            # Style params
            if atom_style == "spacefill":
                size_base, size_scale = 120, 200
                shadow_alpha = 0.25
            elif atom_style == "sticks":
                # Sticks: medium bonds, tiny atom caps
                size_base, size_scale = 18, 35
                shadow_alpha = 0.12
            elif atom_style == "skeletal":
                size_base, size_scale = 10, 20
                shadow_alpha = 0.0
            elif atom_style == "licorice":
                size_base, size_scale = 60, 100
                shadow_alpha = 0.2
            elif atom_style == "wire":
                size_base, size_scale = 20, 40
                shadow_alpha = 0.1
            else:  # shaded / flat default
                size_base, size_scale = 80, 140
                shadow_alpha = 0.25

            # Draw Bonds
            if 'Bonds' in mol.display_mode and len(mol.bonds) > 0:
                lines = []
                colors = []
                linewidths = []
                lw_scale = 1.0
                if bond_style == "thick":
                    lw_scale = 1.6
                elif bond_style == "thin":
                    lw_scale = 0.7
                display_mode_lower = (mol.display_mode or "").lower()
                force_atom_bond_colors = (atom_style == "sticks") or ("bonds only" in display_mode_lower)
                for (i, j) in mol.bonds:
                    if i >= len(coords) or j >= len(coords): continue
                    ei = (mol.elements[i] or "").strip().upper() if i < len(mol.elements) else ""
                    ej = (mol.elements[j] or "").strip().upper() if j < len(mol.elements) else ""
                    if hide_h:
                        try:
                            if ei == 'H' or ej == 'H':
                                continue
                        except Exception:
                            pass
                    p1 = coords[i]
                    p2 = coords[j]
                    lines.append([(p1[0], p1[1]), (p2[0], p2[1])])
                    
                    # Depth cueing for bonds
                    z_mid = (p1[2] + p2[2]) * 0.5
                    z_norm = (z_mid - z_min) / z_range
                    alpha = 0.4 + 0.6 * z_norm
                    # Choose bond color
                    bond_mode = getattr(mol, "bond_color_mode", None) or self._bond_color_mode
                    if force_atom_bond_colors and bond_mode == "default":
                        bond_mode = "by_atoms"
                    if bond_mode == "single" and mol.bond_color_override:
                        br, bg, bb, _ = matplotlib.colors.to_rgba(mol.bond_color_override)
                    elif bond_mode == "by_atoms":
                        c1 = mol.atom_color_map.get(ei, mol.atom_color_map.get(ei.title(), None)) if hasattr(mol, "atom_color_map") else None
                        c2 = mol.atom_color_map.get(ej, mol.atom_color_map.get(ej.title(), None)) if hasattr(mol, "atom_color_map") else None
                        if not c1: c1 = get_atom_color(ei, self.molecule_palette)
                        if not c2: c2 = get_atom_color(ej, self.molecule_palette)
                        r1, g1, b1, _ = matplotlib.colors.to_rgba(c1)
                        r2, g2, b2, _ = matplotlib.colors.to_rgba(c2)
                        br, bg, bb = (0.5*(r1+r2), 0.5*(g1+g2), 0.5*(b1+b2))
                    else:
                        br, bg, bb = self._default_bond_color
                    colors.append((br, bg, bb, alpha))
                    linewidths.append((1.0 + 2.0 * z_norm) * lw_scale)
                
                lc = LineCollection(lines, colors=colors, linewidths=linewidths, zorder=29)
                ax.add_collection(lc)
                lc.set_pickradius(5) # Help hit testing if needed later

            # Draw Atoms
            if 'Atoms' in mol.display_mode:
                # Sort atoms by Z for simple painter's algorithm
                order = np.argsort(z_vals)
                if hide_h:
                    order = [idx for idx in order if str(mol.elements[idx]).strip().upper() != 'H']
                if len(order) == 0:
                    continue
                coords_sorted = coords[order]
                elements_sorted = [mol.elements[i] for i in order]
                
                x = coords_sorted[:, 0]
                y = coords_sorted[:, 1]
                z = coords_sorted[:, 2]
                
                z_norm = (z - z_min) / z_range
                rad_mode = getattr(mol, "radius_mode", "covalent")
                rad_scale = getattr(mol, "radius_scale", 1.0)
                rad_ref = max(get_atom_radius('C', 'covalent'), 1e-3)
                radius_factors = []
                for e in elements_sorted:
                    r_el = get_atom_radius(e, rad_mode)
                    radius_factors.append(max(r_el / rad_ref, 0.05) * rad_scale)
                sizes = (size_base + size_scale * z_norm) * np.array(radius_factors)
                
                base_colors = []
                for e in elements_sorted:
                    m = getattr(mol, "atom_color_map", {}) or {}
                    override = m.get(e.upper()) or m.get(e.title())
                    if override:
                        base_colors.append(override)
                    elif mol.atom_color_override:
                        base_colors.append(mol.atom_color_override)
                    else:
                        base_colors.append(get_atom_color(e, self.molecule_palette))
                rgba_colors = [matplotlib.colors.to_rgba(c) for c in base_colors]
                final_colors = []
                for i, (r, g, b, a) in enumerate(rgba_colors):
                    depth_alpha = 0.45 + 0.55 * z_norm[i]
                    highlight = 0.25 if atom_style in ("shaded", "spacefill") else 0.0
                    r_h = min(1.0, r + (1 - r) * highlight)
                    g_h = min(1.0, g + (1 - g) * highlight)
                    b_h = min(1.0, b + (1 - b) * highlight)
                    final_colors.append((r_h, g_h, b_h, depth_alpha))
                
                if self._show_molecule_shadow and atom_style in ("shaded", "spacefill", "licorice"):
                    shadow_sc = ax.scatter(
                        x + 0.05, y - 0.05,
                        s=sizes * 1.25,
                        c=[(0, 0, 0, shadow_alpha)] * len(x),
                        edgecolors='none',
                        linewidths=0,
                        zorder=28,
                    )
                sc = ax.scatter(x, y, s=sizes, c=final_colors, edgecolors='black',
                                linewidths=0.6 if atom_style != "wire" else 0.3, zorder=30)

            self._molecule_artists.append({
                'mol': mol,
                'ax': ax,
                'scatter': sc,
                'shadow': shadow_sc,
                'lines': lc
            })

    def _draw_spectra(self, ax):
        view = self._ax_view_map.get(ax, {})
        specs = view.get('spectra') or []
        if not specs:
            self._spectra_points[ax] = []
            return
        raw_extent = view.get('extent_raw')
        rel = self._use_relative_axes(view)
        arr_vals = np.asarray(view.get('arr'))
        try:
            arr_h, arr_w = arr_vals.shape
        except Exception:
            arr_h = arr_w = 0
        meta = self._image_meta.get(ax, {})
        extent_used = meta.get('extent')
        origin = meta.get('origin', 'upper')
        shape = meta.get('shape')
        if shape and len(shape) == 2:
            arr_h, arr_w = shape
        extent = extent_used or (view.get('extent') or raw_extent)
        if raw_extent:
            x0, x1, y1, y0 = raw_extent
        else:
            x0, x1, y0, y1 = 0.0, 1.0, 0.0, 1.0
        if extent:
            vals = list(extent)
            if len(vals) == 4:
                ex0, ex1, ey0, ey1 = vals
            else:
                ex0, ex1, ey0, ey1 = x0, x1, y1, y0
        else:
            ex0, ex1 = 0.0, float(max(arr_w - 1, 1))
            ey0, ey1 = 0.0, float(max(arr_h - 1, 1))
        cols = max(arr_w - 1, 1)
        rows = max(arr_h - 1, 1)
        def _axis_from_pixel(col, row):
            if extent_used is not None and meta.get('shape'):
                xmin, xmax, ymin, ymax = extent_used
                span_x = xmax - xmin
                span_y = ymax - ymin
                if cols == 0:
                    x_axis = xmin
                else:
                    x_axis = xmin + (col / float(cols)) * span_x
                if rows == 0:
                    y_axis = ymax if str(origin).lower() == 'upper' else ymin
                else:
                    if str(origin).lower() == 'upper':
                        y_axis = ymax - (row / float(rows)) * span_y
                    else:
                        y_axis = ymin + (row / float(rows)) * span_y
                return x_axis, y_axis
            x_axis = ex0 if cols == 0 else ex0 + (col / float(cols)) * (ex1 - ex0)
            y_axis = ey0 if rows == 0 else ey0 + (row / float(rows)) * (ey1 - ey0)
            return x_axis, y_axis
        normal_xs = []
        normal_ys = []
        highlight_xs = []
        highlight_ys = []
        points = []
        missing_specs = []
        highlight_spec = view.get('highlight_spec')
        pulse = float(getattr(self, "_highlight_pulse_strength", 1.0) or 1.0)
        pixel_lookup = {id(spec): (col, row) for spec, col, row in (view.get('spec_pixels') or [])}
        for idx, s in enumerate(specs):
            coords = pixel_lookup.get(id(s))
            if coords is None:
                missing_specs.append(s)
                continue
            x, y = _axis_from_pixel(coords[0], coords[1])
            points.append((x, y, s))
            if highlight_spec is not None and s is highlight_spec:
                highlight_xs.append(x); highlight_ys.append(y)
            else:
                normal_xs.append(x); normal_ys.append(y)
        # Fallback grid placement for entries without coordinates so markers still show up
        m = len(missing_specs)
        if m:
            cols = int(math.ceil(math.sqrt(m)))
            rows = int(math.ceil(m / float(max(cols, 1))))
            dx = (x1 - x0) / float(max(cols, 1))
            dy = (y1 - y0) / float(max(rows, 1))
            for i, spec in enumerate(missing_specs):
                r = i // cols
                c = i % cols
                fx = x0 + (c + 0.5) * dx
                fy = y0 + (r + 0.5) * dy
                points.append((fx, fy, spec))
                if highlight_spec is not None and spec is highlight_spec:
                    highlight_xs.append(fx); highlight_ys.append(fy)
                else:
                    normal_xs.append(fx); normal_ys.append(fy)
        if not (normal_xs or highlight_xs):
            self._spectra_points[ax] = []
            return
        try:
            if normal_xs:
                ax.scatter(normal_xs, normal_ys, s=28, marker='o', facecolor='#ffcc00', edgecolor='#1a1a1a', linewidths=0.7, alpha=0.9, zorder=35)
            if highlight_xs:
                outer = 260 * (0.9 + 0.3 * pulse)
                core = 140 * (0.7 + 0.3 * pulse)
                ax.scatter(highlight_xs, highlight_ys, s=outer, marker='o', facecolor='#ffe8fb', edgecolor='none', alpha=0.22, zorder=36)
                ax.scatter(highlight_xs, highlight_ys, s=core, marker='o', facecolor='none', edgecolor='#ff5fb7', linewidths=2.4, alpha=0.9, zorder=37)
            self._spectra_points[ax] = points
        except Exception:
            self._spectra_points[ax] = []

    def update_highlight_pulse(self, strength: float):
        self._highlight_pulse_strength = float(strength) if strength else 1.0
        self.draw_idle()

    def _hit_spectrum_point(self, event):
        """Return the nearest spectrum under the cursor if within a small pixel radius."""
        if event is None or event.inaxes is None:
            return None
        pts = self._spectra_points.get(event.inaxes) or []
        if not pts:
            return None
        try:
            ex, ey = float(event.x), float(event.y)
        except Exception:
            return None
        best = None
        best_d2 = None
        for x, y, spec in pts:
            try:
                sx, sy = event.inaxes.transData.transform((x, y))
            except Exception:
                continue
            dx = sx - ex
            dy = sy - ey
            d2 = dx * dx + dy * dy
            if best_d2 is None or d2 < best_d2:
                best_d2 = d2
                best = spec
        if best is None:
            return None
        # 12px radius for reliable clicks
        if best_d2 is not None and best_d2 <= 144.0:
            return best
        return None

    def _update_molecule_artists(self):
        """Update positions of existing molecule artists without full redraw."""
        for entry in self._molecule_artists:
            mol = entry['mol']
            sc = entry['scatter']
            shadow_sc = entry.get('shadow')
            lc = entry['lines']
            coords = mol.get_transformed_coordinates()
            if len(coords) == 0:
                continue

            if lc:
                lines = []
                for (i, j) in mol.bonds:
                    if i >= len(coords) or j >= len(coords):
                        continue
                    p1 = coords[i]
                    p2 = coords[j]
                    lines.append([(p1[0], p1[1]), (p2[0], p2[1])])
                lc.set_segments(lines)

            if sc:
                order = np.argsort(coords[:, 2])
                coords_sorted = coords[order]
                sc.set_offsets(np.c_[coords_sorted[:, 0], coords_sorted[:, 1]])
                if shadow_sc:
                    shadow_sc.set_offsets(np.c_[coords_sorted[:, 0] + 0.05, coords_sorted[:, 1] - 0.05])

        self.draw_idle()

    def enable_scale_bar(self, enable: bool):
        if enable == self.scale_bar_enabled:
            return
        self.scale_bar_enabled = enable
        if enable:
            self._connect_scale_bar_events()
        else:
            self._disconnect_scale_bar_events()
        self._redraw()

    def _connect_scale_bar_events(self):
        if self._scale_bar_cids:
            return
        self._scale_bar_cids = [
            self.mpl_connect('button_press_event', self._on_sb_press),
            self.mpl_connect('motion_notify_event', self._on_sb_motion),
            self.mpl_connect('button_release_event', self._on_sb_release),
        ]

    def _disconnect_scale_bar_events(self):
        for cid in self._scale_bar_cids:
            self.mpl_disconnect(cid)
        self._scale_bar_cids = []

    def _calculate_best_scale_bar(self, width, unit):
        if width <= 0:
            return 1.0, unit
        # Target roughly 15-20% of the image width
        target = width * 0.18
        exponent = math.floor(math.log10(target))
        fraction = target / (10**exponent)
        
        # Candidates for "elegant" sizes: 1, 2, 3, 4, 5, 10
        candidates = [1, 2, 3, 4, 5, 10]
        best_mantissa = min(candidates, key=lambda x: abs(x - fraction))
        size = best_mantissa * (10**exponent)
        
        # Auto-format label for common units
        label = f"{size:g} {unit}"
        if unit == 'nm':
            if size < 1.0:
                label = f"{size*1000:.0f} pm"
            elif size >= 1000:
                label = f"{size/1000:.2g} µm"
            else:
                label = f"{size:g} nm"
        elif unit == 'µm':
            if size < 1.0:
                label = f"{size*1000:.0f} nm"
            else:
                label = f"{size:g} µm"
        
        return size, label

    def _add_scale_bar(self, ax, view):
        extent = view.get('extent')
        if extent is None:
            # Fallback for pixel coords
            h, w = np.shape(view['arr'])
            width = w
            unit = 'px'
        else:
            width = abs(extent[1] - extent[0])
            unit = view.get('axis_unit') or 'nm'
            
        size, label = self._calculate_best_scale_bar(width, unit)
        
        font_scale = getattr(self, '_view_font_scale', 1.0)
        sb = AnchoredSizeBar(ax.transData, size, label, 
                             loc='center',  # Anchor point on the artist itself
                             pad=0.4, borderpad=0, sep=3, 
                             frameon=False, 
                             size_vertical=width*0.004*font_scale,
                             color='white',
                             label_top=True,
                             bbox_to_anchor=self._scale_bar_pos,
                             bbox_transform=ax.transAxes)
        
        # Apply font scaling
        sb.size_bar.get_children()[0].set_linewidth(0) # remove border if any
        text = sb.txt_label.get_children()[0]
        
        text.set_fontsize(10 * font_scale)
        text.set_fontweight('bold')
        sb.set_zorder(20)
        
        ax.add_artist(sb)
        self._scale_bar_artists.append(sb)

    def _on_sb_press(self, event):
        if not self.scale_bar_enabled: return
        
        # Check if we clicked a scale bar
        target_sb = None
        for sb in self._scale_bar_artists:
            if sb.contains(event)[0]:
                target_sb = sb
                break
        
        if target_sb is None:
            return

        if event.button == 1:
            self._scale_bar_drag_start = (event.x, event.y)
        elif event.button == 3:
            self._show_sb_context_menu(event)

    def _show_sb_context_menu(self, event):
        menu = QtWidgets.QMenu(self)
        
        col_menu = menu.addMenu("Colors")
        txt_act = col_menu.addAction("Text Color")
        bar_act = col_menu.addAction("Bar Color")
        
        font_menu = menu.addMenu("Font")
        # Top common fonts in Python/World (Windows/Linux/Mac safe-ish subset)
        fonts = [
            "Arial", "DejaVu Sans", "Times New Roman", "Courier New",
            "Verdana", "Tahoma", "Georgia", "Segoe UI",
            "Trebuchet MS", "Impact", "Calibri", "Cambria"
        ]
        for font_name in fonts:
            act = font_menu.addAction(font_name)
            # Show font in its own style
            try:
                f = QtGui.QFont(font_name)
                f.setPointSize(10)
                act.setFont(f)
            except Exception:
                pass
            act.triggered.connect(lambda checked, f=font_name: self._set_sb_font(f))
            
        txt_act.triggered.connect(self._pick_sb_text_color)
        bar_act.triggered.connect(self._pick_sb_bar_color)
        
        if getattr(event, 'guiEvent', None):
            menu.exec_(event.guiEvent.globalPos())

    def _set_sb_font(self, font):
        self._scale_bar_settings['font_family'] = font
        self._redraw()

    def _pick_sb_text_color(self):
        col = QtWidgets.QColorDialog.getColor(QtCore.Qt.white, self, "Select Text Color")
        if col.isValid():
            self._scale_bar_settings['text_color'] = col.name()
            self._redraw()

    def _pick_sb_bar_color(self):
        col = QtWidgets.QColorDialog.getColor(QtCore.Qt.white, self, "Select Bar Color")
        if col.isValid():
            self._scale_bar_settings['bar_color'] = col.name()
            self._redraw()

    def _on_sb_motion(self, event):
        if self._scale_bar_drag_start is None:
            if self.scale_bar_enabled and event.inaxes:
                for sb in self._scale_bar_artists:
                    if sb.contains(event)[0]:
                        self.setCursor(QtCore.Qt.SizeAllCursor)
                        return
            self.setCursor(QtCore.Qt.ArrowCursor)
            return

        if event.inaxes is None: return
        ax = event.inaxes
        dx = (event.x - self._scale_bar_drag_start[0]) / ax.bbox.width
        dy = (event.y - self._scale_bar_drag_start[1]) / ax.bbox.height
        cur_x, cur_y = self._scale_bar_pos
        self._scale_bar_pos = (cur_x + dx, cur_y + dy)
        self._scale_bar_drag_start = (event.x, event.y)
        
        for sb in self._scale_bar_artists:
            if sb.axes:
                sb.set_bbox_to_anchor(self._scale_bar_pos, sb.axes.transAxes)
            
        self.draw_idle()

    def _on_sb_release(self, event):
        self._scale_bar_drag_start = None

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
            saved.append({'pts': tuple(pts), 'color': entry.get('color'), 'lw': entry.get('lw')})
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
                self._add_saved_profile_from_pts(tuple(pts), entry.get('color'), entry.get('lw'))
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
            self._ensure_angle_frames()
            self._emit_angle()
        else:
            self._disconnect_angle_events()
            self._clear_angle_artists()
            self.angle_pts = None
            self._emit_angle()
        self.draw_idle()

    def clear_angle_measurement(self):
        self._clear_angle_artists()
        if self.angle_enabled:
            self._ensure_angle_frames()
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
        # Update scale bar colors
        sb_settings = getattr(self, '_scale_bar_settings', {})
        sb_text_col = sb_settings.get('text_color') or text_color
        sb_bar_col = sb_settings.get('bar_color') or text_color
        
        for sb in self._scale_bar_artists:
            try:
                sb.size_bar.get_children()[0].set_color(sb_bar_col)
                sb.txt_label.get_children()[0].set_color(sb_text_col)
            except Exception:
                pass
        if self.angle_pts:
            self._update_angle_artists()
        self.draw_idle()

    def _apply_view_font_scale(self):
        scale = max(0.6, min(2.5, getattr(self, '_view_font_scale', 1.0)))
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
        # Update scale bar font size
        for sb in self._scale_bar_artists:
            try:
                sb.txt_label.get_children()[0].set_fontsize(10 * scale)
            except Exception:
                pass
        for frame in self._angle_frames:
            label = frame.get('label')
            if label is not None:
                try:
                    label.set_fontsize(9 * scale)
                except Exception:
                    pass
            for lbl in frame.get('len_labels', []):
                try:
                    lbl.set_fontsize(8 * scale)
                except Exception:
                    pass
        self._apply_tight_layout_safe(pad=max(0.25, 0.35 * scale))
        self.draw_idle()

    def set_profile_label_mode(self, mode: str):
        mode = (mode or "").strip().lower()
        if mode not in ("length", "full", "hidden"):
            mode = "length"
        if mode == self._profile_label_mode:
            return
        self._profile_label_mode = mode
        self._update_profile_markers()
        for entry in self._saved_profiles:
            text = entry.get("label_artist")
            pts = entry.get("pts")
            if text is None or pts is None:
                continue
            try:
                label_text = self._format_profile_label(pts)
                text.set_text(label_text or "")
                text.set_visible(bool(label_text))
            except Exception:
                continue
        self.draw_idle()

    def _apply_tight_layout_safe(self, pad: float = 0.25):
        """Apply tight layout without warning spam; fallback when it cannot fit."""
        try:
            with warnings.catch_warnings(record=True) as caught:
                warnings.simplefilter("always", UserWarning)
                self.fig.tight_layout(pad=float(max(0.05, pad)))
            tight_layout_failed = any(
                "Tight layout not applied" in str(w.message)
                for w in (caught or [])
            )
            if not tight_layout_failed:
                return
            scale = max(0.6, min(2.5, getattr(self, "_view_font_scale", 1.0)))
            extra = min(0.06, 0.012 * max(0.0, scale - 1.0))
            margin = min(0.18, 0.07 + extra)
            self.fig.subplots_adjust(
                left=margin,
                right=0.985,
                bottom=margin,
                top=0.965,
                wspace=0.14,
                hspace=0.18,
            )
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
                self._view_font_scale = min(2.5, max(0.6, self._view_font_scale + step))
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
            try:
                self._emit_profile()
            except Exception:
                pass
        else:
            self._disconnect_profile_events()
            self._clear_profile_artists()
            self.profile_pts = None
            self._clear_saved_profile_artists(notify=False)
            self._active_profile_original_color = None
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
        if event is not None:
            try:
                if event.modifiers() & QtCore.Qt.ControlModifier and event.key() == QtCore.Qt.Key_Z:
                    # Undo priority: outlines -> molecules
                    if self._undo_last_outline():
                        return
                    self.undo_last_molecule_change()
                    return
                if event.modifiers() == QtCore.Qt.NoModifier and event.key() == QtCore.Qt.Key_R:
                    self._reset_view_zoom()
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
            self._profile_line, = self.main_ax.plot([x0,x1],[y0,y1], color=color, lw=self._active_profile_lw, alpha=0.95, zorder=9)
            self._profile_p0, = self.main_ax.plot([x0],[y0], marker='o', color=color, ms=7, mec='black', mew=1.0, zorder=10)
            self._profile_p1, = self.main_ax.plot([x1],[y1], marker='o', color=color, ms=7, mec='black', mew=1.0, zorder=10)
            self._profile_endpoint_labels = self._create_endpoint_labels((x0, y0, x1, y1), color)
            self._profile_label = self._create_profile_id_label((x0, y0, x1, y1), "Active", color)
            self._update_profile_markers()
        
        # Clear existing echo artists to prevent duplicates
        for entry in self._profile_echo_artists:
            for art in entry.values():
                try: art.remove()
                except Exception: pass
        self._profile_echo_artists = []
        x0, y0, x1, y1 = self.profile_pts
        color = self._active_profile_color
        for ax in self._ax_view_map:
            if ax is self.main_ax:
                continue
            try:
                l, = ax.plot([x0,x1],[y0,y1], color=color, lw=self._active_profile_lw, alpha=0.95, zorder=9)
                p0, = ax.plot([x0],[y0], marker='o', color=color, ms=7, mec='black', mew=1.0, zorder=10)
                p1, = ax.plot([x1],[y1], marker='o', color=color, ms=7, mec='black', mew=1.0, zorder=10)
                self._profile_echo_artists.append({'line': l, 'p0': p0, 'p1': p1})
            except Exception:
                pass

    def _ensure_angle_frames(self):
        if self.main_ax is None:
            return
        if not self._angle_frames:
            self._add_angle_frame_at(None, None)
        for frame in self._angle_frames:
            self._ensure_frame_artists(frame)

    def _add_angle_frame_at(self, x, y):
        center = (x, y) if (x is not None and y is not None) else None
        frame = self._create_angle_frame(center=center)
        self._angle_frames.append(frame)
        self._set_active_angle_frame_index(len(self._angle_frames) - 1)
        self._ensure_frame_artists(frame)
        self._update_angle_artists()
        self._emit_angle()

    def _create_angle_frame(self, center=None):
        pts = self._default_angle_pts(center=center)
        color_idx = len(self._angle_frames) % len(self._angle_frame_colors) if self._angle_frame_colors else 0
        color_a, color_b = self._angle_frame_colors[color_idx] if self._angle_frame_colors else ('#ffb300', '#00acc1')
        return {
            'pts': pts,
            'color_a': color_a,
            'color_b': color_b,
            'style': 'dots',
            'lines': [],
            'markers': [],
            'arrows': [],
            'label': None,
            'len_labels': [],
            'patch': None,
        }

    def _default_angle_pts(self, center=None):
        vx = vy = 0.0
        ax = vx + 1.0
        ay = vy
        bx = vx
        by = vy + 1.0
        if self.main_ax is not None:
            try:
                xlim = self.main_ax.get_xlim()
                ylim = self.main_ax.get_ylim()
                base_x = 0.5 * (xlim[0] + xlim[1])
                base_y = 0.5 * (ylim[0] + ylim[1])
                span_x = max(abs(xlim[1] - xlim[0]), 1e-6)
                span_y = max(abs(ylim[1] - ylim[0]), 1e-6)
                radius = max(0.2 * min(span_x, span_y), 0.1)
                if center is not None:
                    base_x, base_y = center
                vx, vy = base_x, base_y
                ax = vx + radius
                ay = vy
                bx = vx
                by = vy + radius
            except Exception:
                if center is not None:
                    vx, vy = center
                    ax = vx + 1.0
                    ay = vy
                    bx = vx
                    by = vy + 1.0
        elif center is not None:
            vx, vy = center
            ax = vx + 1.0
            ay = vy
            bx = vx
            by = vy + 1.0
        return (vx, vy, ax, ay, bx, by)

    def _set_active_angle_frame_index(self, idx):
        if not self._angle_frames:
            self._active_angle_frame_idx = -1
            self.angle_pts = None
            return
        idx = max(0, min(idx, len(self._angle_frames) - 1))
        self._active_angle_frame_idx = idx
        frame = self._angle_frames[idx]
        self.angle_pts = frame.get('pts')

    def _get_active_angle_frame(self):
        if self._active_angle_frame_idx < 0 or self._active_angle_frame_idx >= len(self._angle_frames):
            return None
        return self._angle_frames[self._active_angle_frame_idx]

    def _ensure_frame_artists(self, frame):
        if not self.main_ax:
            return
        vx, vy, ax, ay, bx, by = frame['pts']
        if not frame.get('lines'):
            line1, = self.main_ax.plot([vx, ax], [vy, ay], color=frame['color_a'], lw=2.4, alpha=0.95, zorder=9)
            line2, = self.main_ax.plot([vx, bx], [vy, by], color=frame['color_b'], lw=2.4, alpha=0.95, zorder=9)
            frame['lines'] = [line1, line2]
        if not frame.get('markers'):
            vertex, = self.main_ax.plot([vx], [vy], marker='o', color='#ffffff', mec='#000000', ms=7, zorder=10)
            end_a, = self.main_ax.plot([ax], [ay], marker='o', color=frame['color_a'], mec='#000000', ms=6, zorder=10)
            end_b, = self.main_ax.plot([bx], [by], marker='o', color=frame['color_b'], mec='#000000', ms=6, zorder=10)
            frame['markers'] = [vertex, end_a, end_b]
        if not frame.get('arrows'):
            arrow_a = patches.FancyArrowPatch((vx, vy), (ax, ay),
                                              arrowstyle='-|>', mutation_scale=14,
                                              linewidth=2.4, color=frame['color_a'],
                                              shrinkA=0, shrinkB=0, zorder=9)
            arrow_b = patches.FancyArrowPatch((vx, vy), (bx, by),
                                              arrowstyle='-|>', mutation_scale=14,
                                              linewidth=2.4, color=frame['color_b'],
                                              shrinkA=0, shrinkB=0, zorder=9)
            self.main_ax.add_patch(arrow_a)
            self.main_ax.add_patch(arrow_b)
            frame['arrows'] = [arrow_a, arrow_b]

    def _update_angle_artists(self):
        if not self._angle_frames:
            return
        for frame in self._angle_frames:
            self._ensure_frame_artists(frame)
            vx, vy, ax, ay, bx, by = frame['pts']
            style = frame.get('style', 'dots')
            lines = frame.get('lines', [])
            if len(lines) == 2:
                lines[0].set_data([vx, ax], [vy, ay])
                lines[1].set_data([vx, bx], [vy, by])
                lines[0].set_visible(style == 'dots')
                lines[1].set_visible(style == 'dots')
            markers = frame.get('markers', [])
            if markers and len(markers) >= 3:
                markers[0].set_data([vx], [vy])
                markers[1].set_data([ax], [ay])
                markers[2].set_data([bx], [by])
            visible_markers = (style == 'dots')
            for marker in markers:
                marker.set_visible(visible_markers)
            arrows = frame.get('arrows', [])
            if arrows and len(arrows) >= 2:
                arrows[0].set_positions((vx, vy), (ax, ay))
                arrows[1].set_positions((vx, vy), (bx, by))
                arrows[0].set_visible(style == 'arrows')
                arrows[1].set_visible(style == 'arrows')
            self._update_frame_label(frame)
        if self._angle_blit_active and self._angle_background is not None:
            self._blit_angle_frames()
        else:
            self.draw_idle()

    def _update_frame_label(self, frame):
        angle_info = self._compute_angle_info(frame=frame)
        label = frame.get('label')
        if label is not None:
            try:
                label.remove()
            except Exception:
                pass
        frame['label'] = None
        for lbl in frame.get('len_labels', []):
            try:
                lbl.remove()
            except Exception:
                pass
        frame['len_labels'] = []
        patch = frame.get('patch')
        if patch is not None:
            try:
                patch.remove()
            except Exception:
                pass
        frame['patch'] = None
        if not angle_info:
            return
        vx, vy, ax, ay, bx, by = frame['pts']
        text = f"{angle_info['angle_deg']:.1f}-"
        unit = angle_info.get('unit')
        color = '#f5f5f5' if self._detail_dark else '#111111'
        bbox_face = '#060606' if self._detail_dark else 'white'
        font_scale = getattr(self, '_view_font_scale', 1.0)
        frame['label'] = self.main_ax.text(
            vx, vy, text,
            color=color,
            fontsize=9 * font_scale,
            ha='center', va='center',
            bbox={'facecolor': bbox_face, 'alpha': 0.65 if self._detail_dark else 0.7, 'edgecolor': 'none', 'pad': 2},
            zorder=12)
        vec_a = np.array([ax - vx, ay - vy], dtype=float)
        vec_b = np.array([bx - vx, by - vy], dtype=float)
        len_a = angle_info['len_a'] or 1.0
        len_b = angle_info['len_b'] or 1.0
        bis = vec_a / max(len_a, 1e-9) + vec_b / max(len_b, 1e-9)
        if np.allclose(bis, 0):
            bis = np.array([-(vec_a[1]), vec_a[0]])
        bis = bis / (np.linalg.norm(bis) + 1e-9)
        offset = min(len_a, len_b) * 0.2
        bx_label = vx + bis[0] * offset
        by_label = vy + bis[1] * offset
        frame['label'].set_position((bx_label, by_label))
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
        frame['patch'] = wedge
        self.main_ax.add_patch(wedge)
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
            frame['len_labels'] = [lbl_a, lbl_b]

    def _clear_angle_artists(self):
        for frame in list(self._angle_frames):
            for art in frame.get('lines', []) + frame.get('markers', []) + frame.get('arrows', []):
                try:
                    if art is not None:
                        art.remove()
                except Exception:
                    pass
            for lbl in frame.get('len_labels', []):
                try:
                    lbl.remove()
                except Exception:
                    pass
            if frame.get('label') is not None:
                try:
                    frame['label'].remove()
                except Exception:
                    pass
            if frame.get('patch') is not None:
                try:
                    frame['patch'].remove()
                except Exception:
                    pass
        self._angle_frames = []
        self._active_angle_frame_idx = -1
        self.angle_pts = None
        self._reset_angle_blit()

    def _set_angle_pts(self, vx, vy, ax, ay, bx, by, frame=None):
        target = frame if frame is not None else self._get_active_angle_frame()
        if target is None:
            return
        xmin, xmax, ymin, ymax = self._profile_bounds()
        def clamp(val, lo, hi):
            return max(lo, min(hi, val))
        vx = clamp(vx, xmin, xmax)
        vy = clamp(vy, ymin, ymax)
        ax = clamp(ax, xmin, xmax)
        ay = clamp(ay, ymin, ymax)
        bx = clamp(bx, xmin, xmax)
        by = clamp(by, ymin, ymax)
        target['pts'] = (vx, vy, ax, ay, bx, by)
        if target is self._get_active_angle_frame():
            self.angle_pts = target['pts']

    def _angle_handle_at(self, x, y):
        best = None
        for idx, frame in enumerate(self._angle_frames):
            if not frame.get('pts'):
                continue
            vx, vy, ax, ay, bx, by = frame['pts']
            handles = {
                'vertex': (vx, vy),
                'a': (ax, ay),
                'b': (bx, by),
            }
            for name, (hx, hy) in handles.items():
                dist = self._pt_distance_pixels(x, y, hx, hy)
                if best is None or dist < best[0]:
                    best = (dist, idx, name)
        if best and best[0] <= 12.0:
            self._set_active_angle_frame_index(best[1])
            return best[1], best[2]
        return None

    def _update_active_angle_style(self, style=None):
        frame = self._get_active_angle_frame()
        if frame is None:
            return
        if style is None:
            style = 'arrows' if frame.get('style', 'dots') == 'dots' else 'dots'
        frame['style'] = style
        self._update_angle_artists()
        self._emit_angle()

    def _compute_angle_info(self, frame=None):
        frame = frame or self._get_active_angle_frame()
        if not frame or not frame.get('pts'):
            return None
        vx, vy, ax, ay, bx, by = frame['pts']
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
        return {
            'angle_deg': angle_deg,
            'len_a': len_a,
            'len_b': len_b,
            'unit': unit,
            'vertex': (vx, vy),
            'frame_index': self._active_angle_frame_idx,
            'total_frames': len(self._angle_frames),
            'style': frame.get('style', 'dots'),
        }

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
        for entry in self._profile_echo_artists:
            for art in entry.values():
                try: art.remove()
                except Exception: pass
        self._profile_echo_artists = []
        self._clear_profile_hud()
        self.draw_idle()

    def _update_profile_artists(self):
        if self._profile_line is None or self._profile_p0 is None or self._profile_p1 is None:
            return
        x0, y0, x1, y1 = self.profile_pts
        self._profile_line.set_data([x0,x1],[y0,y1])
        self._profile_p0.set_data([x0],[y0])
        self._profile_p1.set_data([x1],[y1])
        for entry in self._profile_echo_artists:
            try:
                entry['line'].set_data([x0,x1],[y0,y1])
                entry['p0'].set_data([x0],[y0])
                entry['p1'].set_data([x1],[y1])
            except Exception: pass
        self._update_profile_markers()
        self._update_profile_marker_artists()
        if self._profile_blit_active and self._profile_background is not None:
            self._blit_profile_artists()
        else:
            self.draw_idle()
        if self._dragging is None:
            self._emit_profile()

    def _update_profile_artists_fast(self, draw=True):
        if self._profile_line is None or self._profile_p0 is None or self._profile_p1 is None:
            return
        if self.profile_pts is None:
            return
        x0, y0, x1, y1 = self.profile_pts
        self._profile_line.set_data([x0, x1], [y0, y1])
        self._profile_p0.set_data([x0], [y0])
        self._profile_p1.set_data([x1], [y1])
        self._update_profile_labels()
        if draw:
            self.draw_idle()

    def _schedule_profile_update(self):
        if not self._profile_update_timer.isActive():
            self._profile_update_timer.start()

    def _flush_profile_updates(self):
        if not self.profile_enabled:
            return
        if self.profile_pts is None:
            return
        if self._dragging is not None:
            self._schedule_profile_update()
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
        self._active_profile_original_color = entry.get('color')
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
        if getattr(self, '_profile_ticks', None) is not None:
            try:
                self._profile_ticks.remove()
            except Exception:
                pass
            self._profile_ticks = None
            
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
        mode = getattr(self, "_profile_label_mode", "length")
        if mode == "hidden":
            return ""
        if mode == "full":
            return f"L={length:.3g} {unit} | dx={dx:.3g} {unit} | dy={dy:.3g} {unit}"
        return f"L={length:.3g} {unit}"

    def _create_ticks_and_label(self, pts, color, alpha=0.85, base_size=9):
        size = base_size * getattr(self, '_profile_label_scale', 1.0)
        try:
            fractions = (0.25, 0.5, 0.75)
            tx, ty = [], []
            for frac in fractions:
                x = pts[0] + (pts[2] - pts[0]) * frac
                y = pts[1] + (pts[3] - pts[1]) * frac
                tx.append(x)
                ty.append(y)
            ticks, = self.main_ax.plot(
                tx, ty, marker='s', linestyle='None', color=color,
                ms=max(3.0, 4.0 * self._profile_label_scale),
                alpha=alpha, zorder=9)
            label_text = self._format_profile_label(pts)
            xm = pts[0] + (pts[2] - pts[0]) * 0.5
            ym = pts[1] + (pts[3] - pts[1]) * 0.5
            text = None
            if label_text:
                text = self.main_ax.text(
                    xm, ym, label_text, color=color, fontsize=size,
                    ha='center', va='center',
                    bbox={'facecolor': 'black', 'alpha': 0.28, 'edgecolor': 'none', 'pad': 1.5},
                    zorder=11)
        except Exception:
            return None, None
        return ticks, text

    def _update_profile_markers(self):
        if self.profile_pts is None or self.main_ax is None:
            self._remove_profile_markers()
            return
        
        # Reuse existing artists if possible
        if self._profile_ticks is not None and self._profile_info_text is not None:
            self._remove_profile_markers() # Fallback to recreate if complex update needed, or optimize further
            ticks, text = self._create_ticks_and_label(self.profile_pts, color='yellow', alpha=0.9, base_size=9)
        else:
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
            parts.append(f"?={marker_delta:.3g} {unit}")
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
        if self._profile_animation_enabled:
            # Ensure newly created marker artists join the blit cycle.
            self._set_profile_animated(True)
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

    def _build_profile_data(self, pts, color=None, view=None):
        if pts is None or not self.views:
            return None
        try:
            v0 = view if view is not None else self.views[0]
            arr_raw = np.asarray(v0['arr'], dtype=float)
            flip = self._use_relative_axes(v0)
            arr_src = np.flipud(arr_raw) if flip else arr_raw
            h, w = arr_src.shape
            ax, meta = self._axis_meta_for_view(v0)
            extent_meta = meta.get('extent')
            origin_meta = meta.get('origin', 'upper')
            # Use the same extent/origin that imshow used (handles overrides & relative axes)
            if extent_meta is not None:
                extent = extent_meta
            else:
                extent = self._display_extent_for_view(v0, v0.get('extent', None))
                origin_meta = 'lower' if flip else 'upper'
            axis_unit = self._profile_axis_unit()
            x0, y0, x1, y1 = pts
            if extent is None:
                c0 = x0; r0 = y0; c1 = x1; r1 = y1
                length_nm = None
            else:
                xmin, xmax = extent[0], extent[1]
                ymin, ymax = extent[2], extent[3]
                xr = (xmax - xmin) if (xmax is not None and xmin is not None) else 1.0
                yr = (ymax - ymin) if (ymax is not None and ymin is not None) else 1.0
                c0 = (x0 - xmin) / (xr + 1e-12) * max(w - 1, 1)
                c1 = (x1 - xmin) / (xr + 1e-12) * max(w - 1, 1)
                if str(origin_meta).lower() == 'upper':
                    r0 = (ymax - y0) / (yr + 1e-12) * max(h - 1, 1)
                    r1 = (ymax - y1) / (yr + 1e-12) * max(h - 1, 1)
                else:
                    r0 = (y0 - ymin) / (yr + 1e-12) * max(h - 1, 1)
                    r1 = (y1 - ymin) / (yr + 1e-12) * max(h - 1, 1)
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
                (1 - wy) * (1 - wx) * arr_src[i0, j0] +
                wy * (1 - wx) * arr_src[i1, j0] +
                (1 - wy) * wx * arr_src[i0, j1] +
                wy * wx * arr_src[i1, j1]
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
                'relative_axes': self._use_relative_axes(v0),
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
            arr = np.asarray(v0['arr'])
            h, w = arr.shape
            extent_raw = v0.get('extent', None)
            display_extent = self._display_extent_for_view(v0, extent_raw)
            extent_use = display_extent if display_extent is not None else (0.0, float(w - 1), 0.0, float(h - 1))
            if display_extent is None:
                return (0.0, float(w - 1), 0.0, float(h - 1))
            x0, x1, y1, y0 = extent_use
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
        if self._active_profile_original_color:
            color = self._active_profile_original_color
            self._active_profile_original_color = None
        else:
            color = next(self._profile_color_cycle)
        lw = self._active_profile_lw
        line, = self.main_ax.plot([pts[0], pts[2]], [pts[1], pts[3]], color=color, lw=lw, alpha=0.7, zorder=6, linestyle='--')
        # Combine endpoints into one artist
        endpoints, = self.main_ax.plot([pts[0], pts[2]], [pts[1], pts[3]], marker='o', linestyle='None', color=color, ms=5, mec='black', mew=0.7, alpha=0.9, zorder=7)
        
        artists = [line, endpoints]
        
        # Add echo artists for other views
        for ax in self._ax_view_map:
            if ax is self.main_ax:
                continue
            try:
                l, = ax.plot([pts[0], pts[2]], [pts[1], pts[3]],
                             color=color, lw=lw, alpha=0.7, zorder=6, linestyle='--')
                ep, = ax.plot([pts[0], pts[2]], [pts[1], pts[3]], 
                              marker='o', linestyle='None', color=color,
                              ms=5, mec='black', mew=0.7, alpha=0.9, zorder=7)
                artists.extend([l, ep])
            except Exception:
                pass

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
        if ticks: artists.append(ticks)
        if text: artists.append(text)
        if overlay_label is not None:
            artists.append(overlay_label)
        artists += endpoint_labels
        data = self._build_profile_data(pts, color=color)
        entry = {'artists': artists, 'pts': pts, 'color': color, 'data': data,
                 'overlay_label_artist': overlay_label, 'endpoint_labels': endpoint_labels, 'lw': lw}
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

    def _add_saved_profile_from_pts(self, pts, color, lw=1.5):
        if pts is None or self.main_ax is None:
            return
        pts = tuple(pts)
        color = color or next(self._profile_color_cycle)
        lw = float(lw or 1.5)
        line, = self.main_ax.plot([pts[0], pts[2]], [pts[1], pts[3]],
                                  color=color, lw=lw, alpha=0.7, zorder=6, linestyle='--')
        endpoints, = self.main_ax.plot([pts[0], pts[2]], [pts[1], pts[3]], marker='o', linestyle='None', color=color,
                                       ms=5, mec='black', mew=0.7, alpha=0.9, zorder=7)
        
        artists = [line, endpoints]
        
        # Add echo artists for other views
        for ax in self._ax_view_map:
            if ax is self.main_ax:
                continue
            try:
                l, = ax.plot([pts[0], pts[2]], [pts[1], pts[3]],
                             color=color, lw=lw, alpha=0.7, zorder=6, linestyle='--')
                ep, = ax.plot([pts[0], pts[2]], [pts[1], pts[3]], 
                              marker='o', linestyle='None', color=color,
                              ms=5, mec='black', mew=0.7, alpha=0.9, zorder=7)
                artists.extend([l, ep])
            except Exception:
                pass

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
        if ticks: artists.append(ticks)
        if text: artists.append(text)
        if overlay_label is not None:
            artists.append(overlay_label)
        artists += endpoint_labels
        data = self._build_profile_data(pts, color=color)
        entry = {'artists': artists, 'pts': pts, 'color': color, 'data': data,
                 'overlay_label_artist': overlay_label, 'endpoint_labels': endpoint_labels, 'lw': lw}
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
            base_lw = entry.get('lw', 1.5)
            
            # Update all line artists (main + echoes)
            for art in artists:
                # Check if it's a line (has set_linewidth) and not markers (linestyle='None')
                if hasattr(art, 'set_linewidth') and hasattr(art, 'get_linestyle'):
                    if art.get_linestyle() != 'None':
                        try:
                            if idx == self._highlighted_overlay:
                                art.set_linewidth(base_lw + 1.0)
                                art.set_alpha(1.0)
                            else:
                                art.set_linewidth(base_lw)
                                art.set_alpha(0.35)
                        except Exception:
                            pass
                    # Optional: dim markers too
                    elif art.get_linestyle() == 'None':
                        try:
                            if idx == self._highlighted_overlay:
                                art.set_alpha(0.9)
                            else:
                                art.set_alpha(0.35)
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
        
        # Right click context menu for profiles
        if event.button == 3:
            # Check overlay first (increased threshold for easier hitting)
            overlay_idx = self._overlay_index_near(x, y, thresh=15.0)
            if overlay_idx is not None:
                self._show_profile_context_menu(event, overlay_idx=overlay_idx)
                return
            # Check active profile
            if self.profile_pts is not None:
                dist_line = self._distance_to_segment_pixels(x, y, self.profile_pts)
                if dist_line <= 15.0:
                    self._show_profile_context_menu(event, active=True)
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
            self._line_drag_origin = None
            self._set_profile_animated(True)
            self._prepare_profile_blit()
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
                self._set_profile_animated(True)
                self._prepare_profile_blit()
                return
        if d1 <= thresh:
            self._dragging = 'p1'
            self._line_drag_origin = None
            self._set_profile_animated(True)
            self._prepare_profile_blit()
            return
        if self.profile_pts is not None:
            dist_line = self._distance_to_segment_pixels(x, y, self.profile_pts)
            if dist_line <= thresh:
                self._dragging = 'line'
                self._line_drag_origin = (x, y, self.profile_pts)
                self._set_profile_animated(True)
                self._prepare_profile_blit()
                return
        # Increased threshold to 15.0 to prevent accidental "misses" causing profile loss
        overlay_idx = self._overlay_index_near(x, y, thresh=15.0)
        if overlay_idx is not None:
            if self.profile_pts is not None:
                self._snapshot_active_profile()
            activated = self.activate_saved_profile(overlay_idx)
            if activated:
                if callable(self._profile_highlight_cb):
                    try:
                        self._profile_highlight_cb(None)
                    except Exception:
                        pass
                x0, y0, x1, y1 = self.profile_pts
                return
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
        if self.profile_pts is not None:
            self._snapshot_active_profile()
        self._active_profile_original_color = None
        self._set_profile_pts((x, y, x, y))
        self._dragging = 'p1'
        self._line_drag_origin = None
        
        # Prepare for blitting
        self._set_profile_animated(True)
        self._prepare_profile_blit()
        if self._profile_blit_active:
            self._draw_profile_animated()
            self._blit_profile_artists()
        self._update_profile_artists()

    def _show_profile_context_menu(self, event, overlay_idx=None, active=False):
        menu = QtWidgets.QMenu(self)
        color_act = menu.addAction("Change Color")
        thicker_act = menu.addAction("Thicker")
        thinner_act = menu.addAction("Thinner")
        menu.addSeparator()
        label_menu = menu.addMenu("Label detail")
        label_modes = [
            ("Length only", "length"),
            ("Full (L, dx, dy)", "full"),
            ("Hidden", "hidden"),
        ]
        label_actions = {}
        current_mode = getattr(self, "_profile_label_mode", "length")
        for label_txt, mode_key in label_modes:
            act = label_menu.addAction(label_txt)
            act.setCheckable(True)
            act.setChecked(current_mode == mode_key)
            label_actions[act] = mode_key
        
        if overlay_idx is not None:
            menu.addSeparator()
            delete_act = menu.addAction("Delete Profile")
        
        action = menu.exec_(event.guiEvent.globalPos())
        
        if action == color_act:
            self._change_profile_color(overlay_idx, active)
        elif action == thicker_act:
            self._change_profile_width(overlay_idx, active, 0.5)
        elif action == thinner_act:
            self._change_profile_width(overlay_idx, active, -0.5)
        elif action in label_actions:
            self.set_profile_label_mode(label_actions[action])
        elif overlay_idx is not None and action == delete_act:
            self._remove_saved_profile(overlay_idx)

    def _change_profile_color(self, overlay_idx, active):
        current_color = self._active_profile_color
        if overlay_idx is not None and 0 <= overlay_idx < len(self._saved_profiles):
            current_color = self._saved_profiles[overlay_idx].get('color', current_color)
        
        col = QtWidgets.QColorDialog.getColor(QtGui.QColor(current_color), self, "Select Profile Color")
        if not col.isValid(): return
        new_color = col.name()
        
        if active:
            self._active_profile_color = new_color
            if self._profile_line:
                self._profile_line.set_color(new_color)
            if self._profile_p0:
                self._profile_p0.set_color(new_color)
            if self._profile_p1:
                self._profile_p1.set_color(new_color)
            for entry in self._profile_echo_artists:
                if entry.get('line'): entry['line'].set_color(new_color)
                if entry.get('p0'): entry['p0'].set_color(new_color)
                if entry.get('p1'): entry['p1'].set_color(new_color)
            self.draw_idle()
            self._emit_profile()
        
        if overlay_idx is not None and 0 <= overlay_idx < len(self._saved_profiles):
            entry = self._saved_profiles[overlay_idx]
            entry['color'] = new_color
            # Update artists
            for art in entry.get('artists', []):
                try: art.set_color(new_color)
                except: pass
                try: art.set_markeredgecolor('black')
                except: pass
            if entry.get('overlay_label_artist'):
                entry['overlay_label_artist'].set_color(new_color)
            self.draw_idle()
            self._emit_profile()

    def _change_profile_width(self, overlay_idx, active, delta):
        if active:
            self._active_profile_lw = max(0.5, self._active_profile_lw + delta)
            if self._profile_line:
                self._profile_line.set_linewidth(self._active_profile_lw)
                for entry in self._profile_echo_artists:
                    if entry.get('line'): entry['line'].set_linewidth(self._active_profile_lw)
            self.draw_idle()
        
        if overlay_idx is not None and 0 <= overlay_idx < len(self._saved_profiles):
            entry = self._saved_profiles[overlay_idx]
            cur_lw = entry.get('lw', 1.5)
            new_lw = max(0.5, cur_lw + delta)
            entry['lw'] = new_lw
            # Update artists (first artist is usually the line)
            artists = entry.get('artists', [])
            if artists and isinstance(artists[0], Line2D):
                artists[0].set_linewidth(new_lw)
            self.draw_idle()

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
        # Hide echo artists during drag for performance
        for entry in self._profile_echo_artists:
            for art in entry.values():
                art.set_visible(False)
        if self._dragging == 'p0':
            self._set_profile_pts((x, y, x1, y1))
        elif self._dragging == 'p1':
            self._set_profile_pts((x0, y0, x, y))
        elif self._dragging == 'line' and self._line_drag_origin is not None and self.profile_pts is not None:
            sx, sy, pts = self._line_drag_origin
            dx = x - sx
            dy = y - sy
            self._set_profile_pts((pts[0] + dx, pts[1] + dy, pts[2] + dx, pts[3] + dy))
        
        # Use blitting for smooth drag
        if self._profile_blit_active and self._profile_background is not None:
            self._update_profile_artists_fast(draw=False)
            self._blit_profile_artists()
        else:
            self._update_profile_artists_fast()
            
        self._schedule_profile_update()

    def _on_release(self, event):
        if not self.profile_enabled:
            return
        self._dragging = None
        self._set_profile_animated(False)
        self._reset_profile_blit()
        # Restore echo artists
        for entry in self._profile_echo_artists:
            for art in entry.values():
                art.set_visible(True)
        self._line_drag_origin = None
        self._profile_marker_drag_idx = None
        self._flush_profile_updates()
        if self._profile_state_deferred:
            self._profile_state_deferred = False
            self._flush_profile_state()

    def _profile_animation_artists(self):
        artists = [
            self._profile_line,
            self._profile_p0,
            self._profile_p1,
            self._profile_label,
            self._profile_ticks,
            self._profile_info_text,
        ]
        artists.extend(self._profile_endpoint_labels)
        artists.extend(self._profile_marker_artists)
        return [art for art in artists if art is not None]

    def _set_profile_animated(self, animated):
        """Set animated state for active profile artists to enable/disable blitting."""
        self._profile_animation_enabled = bool(animated)
        for art in self._profile_animation_artists():
            try:
                art.set_animated(animated)
            except Exception:
                pass
        if not animated:
            self._reset_profile_blit()

    def _draw_profile_animated(self):
        """Draw only the active profile artists (for blitting)."""
        for art in self._profile_animation_artists():
            try:
                visible = art.get_visible()
            except Exception:
                visible = True
            if visible:
                try:
                    self.main_ax.draw_artist(art)
                except Exception:
                    pass

    def _prepare_profile_blit(self):
        if self.main_ax is None:
            self._profile_background = None
            self._profile_blit_active = False
            return
        try:
            self.draw()
        except Exception:
            pass
        try:
            self._profile_background = self.copy_from_bbox(self.main_ax.bbox)
            self._profile_blit_active = self._profile_background is not None
        except Exception:
            self._profile_background = None
            self._profile_blit_active = False

    def _reset_profile_blit(self):
        self._profile_background = None
        self._profile_blit_active = False

    def _blit_profile_artists(self):
        if not self.main_ax or self._profile_background is None:
            self.draw_idle()
            return
        try:
            self.restore_region(self._profile_background)
        except Exception:
            self.draw_idle()
            return
        self._draw_profile_animated()
        try:
            self.blit(self.main_ax.bbox)
        except Exception:
            self.draw_idle()

    def _prepare_angle_blit(self):
        if self.main_ax is None:
            self._angle_background = None
            self._angle_blit_active = False
            return
        try:
            self.draw()
        except Exception:
            pass
        try:
            self._angle_background = self.copy_from_bbox(self.main_ax.bbox)
            self._angle_blit_active = self._angle_background is not None
        except Exception:
            self._angle_background = None
            self._angle_blit_active = False

    def _reset_angle_blit(self):
        self._angle_background = None
        self._angle_blit_active = False

    def _draw_angle_frame_fast(self, frame):
        if not frame or not self.main_ax:
            return
        artists = []
        artists.extend(frame.get('lines', []))
        artists.extend(frame.get('markers', []))
        artists.extend(frame.get('arrows', []))
        patch = frame.get('patch')
        if patch is not None:
            artists.append(patch)
        label = frame.get('label')
        if label is not None:
            artists.append(label)
        artists.extend(frame.get('len_labels', []))
        for art in artists:
            if art is not None and art.get_visible():
                try:
                    self.main_ax.draw_artist(art)
                except Exception:
                    pass

    def _blit_angle_frames(self):
        if not self.main_ax or self._angle_background is None:
            self.draw_idle()
            return
        try:
            self.restore_region(self._angle_background)
            for frame in self._angle_frames:
                self._draw_angle_frame_fast(frame)
            self.blit(self.main_ax.bbox)
        except Exception:
            self.draw_idle()

    def _on_angle_press(self, event):
        if not self.angle_enabled or event.inaxes is None or event.inaxes is not self.main_ax:
            return
        if event.button != 1:
            return
        x, y = event.xdata, event.ydata
        if x is None or y is None:
            return
        gui_event = getattr(event, "guiEvent", None)
        ctrl_shift = False
        if gui_event is not None:
            mods = gui_event.modifiers()
            ctrl_shift = bool(mods & QtCore.Qt.ControlModifier) and bool(mods & QtCore.Qt.ShiftModifier)
        else:
            key = getattr(event, "key", "")
            ctrl_shift = "control" in str(key).lower() and "shift" in str(key).lower()
        if ctrl_shift:
            self._add_angle_frame_at(x, y)
            return
        hit = self._angle_handle_at(x, y)
        if not hit:
            return
        self._angle_dragging = hit
        self._prepare_angle_blit()

    def _on_angle_motion(self, event):
        if not self.angle_enabled or self._angle_dragging is None:
            return
        if event.inaxes is not self.main_ax:
            return
        x, y = event.xdata, event.ydata
        if x is None or y is None:
            return
        frame_idx, handle = self._angle_dragging
        if frame_idx < 0 or frame_idx >= len(self._angle_frames):
            return
        frame = self._angle_frames[frame_idx]
        vx, vy, ax, ay, bx, by = frame['pts']
        if handle == 'vertex':
            dx = x - vx
            dy = y - vy
            self._set_angle_pts(x, y, ax + dx, ay + dy, bx + dx, by + dy, frame)
        elif handle == 'a':
            self._set_angle_pts(vx, vy, x, y, bx, by, frame)
        elif handle == 'b':
            self._set_angle_pts(vx, vy, ax, ay, x, y, frame)
        self._update_angle_artists()
        self._emit_angle()

    def _on_angle_release(self, event):
        if not self.angle_enabled:
            return
        self._angle_dragging = None
        self._reset_angle_blit()
        self._update_angle_artists()

    def _emit_profile(self):
        if not getattr(self, "profile_enabled", False) or self.profile_pts is None:
            self._emit_profile_state()
            return
        if not callable(self.profile_callback):
            self._emit_profile_state()
            return
        active = self._build_profile_data(self.profile_pts, color=self._active_profile_color, view=self.views[0] if self.views else None)
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
            # Build extra channels if present
            if len(self.views) > 1:
                extras = []
                extra_colors = ['#ff4081', '#00e5ff', '#76ff03', '#d500f9']
                for i, v in enumerate(self.views[1:]):
                    col = extra_colors[i % len(extra_colors)]
                    p = self._build_profile_data(self.profile_pts, color=col, view=v)
                    if p:
                        name = v.get('colorbar_label') or v.get('title') or f"Ch{i+2}"
                        p['name'] = name
                        extras.append(p)
                active['extra_channels'] = extras

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
        # If clicking on a scale bar, do not trigger base canvas actions (like drag/copy)
        if self.scale_bar_enabled:
            for sb in self._scale_bar_artists:
                if sb.contains(event)[0]:
                    return
        if self.scale_bar_enabled and self._scale_bar_drag_start is not None:
            return
        ax = event.inaxes
        
        if self._check_molecule_hit(event):
            return

        view = self._ax_view_map.get(ax)
        # Double-click: pop out the clicked view if callback provided
        if getattr(event, "dblclick", False) and event.button == 1 and view is not None:
            if callable(self._double_click_callback):
                try:
                    self._double_click_callback(view)
                except Exception:
                    pass
            return
        # Crop/outline rectangle start (handle before tool guards):
        #   Shift + drag -> arbitrary rectangle (crop)
        #   Ctrl + Shift + drag -> square selection (crop)
        #   Alt + drag -> outline extraction in ROI
        # Right-click: outline context menu (style/clear/undo)
        if event.button == 3 and view is not None:
            if event.xdata is not None and event.ydata is not None and self._outlines.get(self._outline_key(view)):
                if self._outline_hit_test(view, event.xdata, event.ydata, ax=ax, event=event):
                    self._show_outline_menu(view)
                    return
        gui_mods = None
        try:
            gui_mods = getattr(getattr(event, "guiEvent", None), "modifiers", lambda: QtCore.Qt.NoModifier)()
        except Exception:
            gui_mods = None
        if gui_mods is None or gui_mods == QtCore.Qt.NoModifier:
            try:
                gui_mods = QtWidgets.QApplication.keyboardModifiers()
            except Exception:
                gui_mods = QtCore.Qt.NoModifier
        mods_qt = gui_mods
        alt_pressed = bool(mods_qt & QtCore.Qt.AltModifier) or 'alt' in str(getattr(event, "key", "")).lower() or bool(getattr(self, "outline_mode", False))
        want_square = (event.button == 1) and (mods_qt & QtCore.Qt.ControlModifier) and (mods_qt & QtCore.Qt.ShiftModifier)
        want_rect = (event.button == 1) and (mods_qt & QtCore.Qt.ShiftModifier)
        # Allow outlining via Alt+left click OR middle click as a fallback shortcut
        want_outline = ((event.button == 1 and alt_pressed) or event.button == 2) and not want_rect and not want_square
        if want_outline and view is not None:
            # Alt+click: outline dominant blob around clicked point (no drag needed)
            if event.xdata is not None and event.ydata is not None:
                # print(f"Alt detected! xdata={event.xdata}, ydata={event.ydata}")
                self._outline_from_point(view, ax, event.xdata, event.ydata)
            else:
                print("[Outline] Alt detected but coordinates are None; ignoring.")
            return
        # Pan start: only in browse-like mode, left button, zoomed view
        if event.button in (1, 2) and not want_rect and view is not None and not self.profile_enabled and not self.angle_enabled:
            if event.xdata is not None and event.ydata is not None and self._is_zoomed(ax):
                self._pan_active = True
                self._pan_ax = ax
                self._pan_start = (event.xdata, event.ydata)
                self._pan_start_lim = (ax.get_xlim(), ax.get_ylim())
                return
        if want_rect and view is not None:
            self._crop_start = (event.xdata, event.ydata)
            self._crop_ax = ax
            self._crop_square = bool(want_square)
            try:
                from matplotlib.patches import Rectangle
                rect = Rectangle((event.xdata, event.ydata), 0, 0,
                                 linewidth=2.2, edgecolor='#ff00cc', facecolor='none', alpha=0.9, linestyle='--')
                ax.add_patch(rect)
                self._crop_rect = rect
                self.draw_idle()
            except Exception:
                self._crop_rect = None
            return
        if event.button == 1:
            hit = self._hit_spectrum_point(event)
            if hit is not None and callable(self._spectra_click_cb):
                try:
                    self._spectra_click_cb(hit, event)
                except Exception:
                    pass
                return
            if (
                self._fixed_crop_quick_mode
                and view is not None
                and not want_rect
                and not want_square
                and not (mods_qt & QtCore.Qt.ShiftModifier)
                and not (mods_qt & QtCore.Qt.ControlModifier)
            ):
                try:
                    if self._apply_fixed_crop_quick(event, view, ax):
                        return
                except Exception:
                    pass
            # Avoid accidental 1 px manual crops when quick mode is enabled and Shift+click with no drag
            if (
                self._fixed_crop_quick_mode
                and (mods_qt & QtCore.Qt.ShiftModifier)
                and not want_rect
                and not want_square
            ):
                return
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

    def _load_molecule_dialog(self):
        start_dir = ""
        if self._recent_molecule_paths:
            start_dir = str(Path(self._recent_molecule_paths[0]).parent)
        elif MultiPreviewCanvas._RECENT_MOLECULES:
            start_dir = str(Path(MultiPreviewCanvas._RECENT_MOLECULES[0]).parent)
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Load Molecule", start_dir, "Molecule Files (*.xyz *.pdb *.mol);;All Files (*)"
        )
        if path:
            self.add_molecule(path)

    def add_molecule(self, path):
        try:
            # Be robust to numpy arrays or non-string inputs
            if isinstance(path, np.ndarray):
                if path.size == 0:
                    raise ValueError("Empty molecule path")
                path = str(path.flatten()[0])
            elif not isinstance(path, (str, Path)):
                path = str(path)
            self._push_molecule_snapshot()
            mol = Molecule(path)
            # Center in current view if possible
            if self.main_ax:
                xlim = self.main_ax.get_xlim()
                ylim = self.main_ax.get_ylim()
                mol.offset = np.array([(xlim[0]+xlim[1])/2, (ylim[0]+ylim[1])/2, 0.0])
            self.molecules.append(mol)
            # Track recent paths (MRU up to 8)
            try:
                norm = str(Path(path).resolve())
                for lst in (self._recent_molecule_paths, MultiPreviewCanvas._RECENT_MOLECULES):
                    if norm in lst:
                        lst.remove(norm)
                    lst.insert(0, norm)
                    if len(lst) > 8:
                        del lst[8:]
                if callable(self._recent_molecule_cb):
                    try:
                        self._recent_molecule_cb(self.get_recent_molecule_paths())
                    except Exception:
                        pass
            except Exception:
                pass
            self._redraw()
        except Exception as e:
            import traceback
            print(f"Failed to load molecule: {e}")
            traceback.print_exc()

    def get_recent_molecule_paths(self):
        """Return MRU list of molecule paths (combined local + global)."""
        recent_all = []
        for lst in (self._recent_molecule_paths, MultiPreviewCanvas._RECENT_MOLECULES):
            for p in lst:
                if p not in recent_all:
                    recent_all.append(p)
        return recent_all[:8]

    def set_recent_molecule_callback(self, cb):
        """Callback invoked when MRU list changes; cb(list[str])."""
        self._recent_molecule_cb = cb

    def _pick_color(self, initial_hex: str | None = None) -> str | None:
        """Show a QColorDialog and return a hex string or None."""
        initial = QtGui.QColor(initial_hex) if initial_hex else QtGui.QColor("#cccccc")
        color = QtWidgets.QColorDialog.getColor(initial, self, "Select color")
        if color.isValid():
            return color.name()
        return None

    def _clear_molecules(self):
        if not self.molecules:
            return
        self._push_molecule_snapshot()
        self.molecules = []
        self._molecule_artists = []
        self._redraw()

    def reset_molecules(self):
        """Public helper to clear all molecules with undo support."""
        self._clear_molecules()

    def _push_molecule_snapshot(self):
        """Save current molecule state for undo."""
        try:
            snap = [m.copy() for m in self.molecules]
            self._molecule_history.append(snap)
            if len(self._molecule_history) > 20:
                self._molecule_history.pop(0)
        except Exception:
            pass

    def undo_last_molecule_change(self):
        """Undo the latest molecule change, if any."""
        if not self._molecule_history:
            return
        try:
            last = self._molecule_history.pop()
            self.molecules = [m.copy() for m in last]
            self._redraw()
        except Exception:
            pass

    def _check_molecule_hit(self, event):
        # When measurement tools are active, avoid picking/dragging molecules to prevent accidental moves.
        if self.profile_enabled or self.angle_enabled:
            return False
        if not self.show_molecules or not self.molecules or event.inaxes is None:
            return False
        
        # Simple hit test: check distance to any atom in any molecule
        # Iterate in reverse to pick top-most
        for idx, mol in reversed(list(enumerate(self.molecules))):
            coords = mol.get_transformed_coordinates()
            if len(coords) == 0: continue
            if not getattr(self, "_show_hydrogens", True):
                mask = [str(el).strip().upper() != 'H' for el in mol.elements]
                if not any(mask):
                    continue
                coords = coords[np.array(mask)]
            
            # Map event data coords to display coords is hard without transform
            # We'll just check data distance. 
            # Assuming atoms are roughly 0.1-0.3 nm radius visually
            dx = coords[:, 0] - event.xdata
            dy = coords[:, 1] - event.ydata
            dist_sq = dx*dx + dy*dy
            min_dist = np.min(dist_sq)
            
            # Threshold: 0.5 nm radius click tolerance
            if min_dist < 0.25: 
                if event.button == 1 or event.button == 2:
                    if not self._molecule_drag_snapshot:
                        self._push_molecule_snapshot()
                        self._molecule_drag_snapshot = True
                    self._molecule_drag_idx = idx
                    self._molecule_drag_start = (event.xdata, event.ydata)
                    self._molecule_drag_start_px = (event.x, event.y)
                    self._molecule_drag_mol_start = mol.offset.copy()
                    self._molecule_drag_mol_angles = mol.angles.copy()
                    
                    key = str(event.key).lower() if event.key else ''
                    if event.button == 2:
                        # Middle button drag: full 3D rotate
                        self._molecule_drag_mode = 'rotate_3d'
                    elif 'control' in key and 'shift' in key:
                        self._molecule_drag_mode = 'rotate_3d'
                    elif 'shift' in key:
                        self._molecule_drag_mode = 'rotate_z'
                    else:
                        self._molecule_drag_mode = 'translate'
                        
                    return True
                elif event.button == 3:
                    self._show_molecule_menu(event, mol)
                    return True
        return False

    def _on_molecule_motion(self, event):
        if self._pan_active and self._pan_ax is event.inaxes:
            if event.xdata is None or event.ydata is None or self._pan_start is None or self._pan_start_lim is None:
                return
            now_ms = time.perf_counter() * 1000.0
            if (now_ms - self._pan_last_ts) < self._pan_throttle_ms:
                return
            self._pan_last_ts = now_ms
            x0, y0 = self._pan_start
            (xlim0, ylim0) = self._pan_start_lim
            dx = event.xdata - x0
            dy = event.ydata - y0
            new_xlim = (xlim0[0] - dx, xlim0[1] - dx)
            new_ylim = (ylim0[0] - dy, ylim0[1] - dy)
            base_xlim, base_ylim = self._zoom_reset_limits.get(self._pan_ax, (xlim0, ylim0))
            new_xlim = self._clamp_limits(new_xlim, base_xlim)
            new_ylim = self._clamp_limits(new_ylim, base_ylim)
            self._pan_ax.set_xlim(new_xlim)
            self._pan_ax.set_ylim(new_ylim)
            # Partial refresh: redraw only this axes if possible, else full canvas
            try:
                self._pan_ax.figure.canvas.draw_idle()
            except Exception:
                self.draw_idle()
            return
        if self._molecule_drag_idx is not None:
            if event.xdata is None or event.ydata is None:
                return
            mol = self.molecules[self._molecule_drag_idx]
            
            if self._molecule_drag_mode == 'translate':
                dx = event.xdata - self._molecule_drag_start[0]
                dy = event.ydata - self._molecule_drag_start[1]
                mol.offset = self._molecule_drag_mol_start + np.array([dx, dy, 0.0])
            
            elif self._molecule_drag_mode == 'rotate_z':
                center = self._molecule_drag_mol_start
                v_start = np.array([self._molecule_drag_start[0] - center[0], self._molecule_drag_start[1] - center[1]])
                v_curr = np.array([event.xdata - center[0], event.ydata - center[1]])
                if np.linalg.norm(v_start) > 0.01 and np.linalg.norm(v_curr) > 0.01:
                    angle_start = np.arctan2(v_start[1], v_start[0])
                    angle_curr = np.arctan2(v_curr[1], v_curr[0])
                    delta_deg = np.degrees(angle_curr - angle_start)
                    new_angles = self._molecule_drag_mol_angles.copy()
                    new_angles[2] += delta_deg
                    mol.angles = new_angles
            
            elif self._molecule_drag_mode == 'rotate_3d':
                if event.x is None or event.y is None: return
                dx_px = event.x - self._molecule_drag_start_px[0]
                dy_px = event.y - self._molecule_drag_start_px[1]
                sensitivity = 0.5 # degrees per pixel
                new_angles = self._molecule_drag_mol_angles.copy()
                new_angles[0] += dy_px * sensitivity
                new_angles[1] += dx_px * sensitivity
                mol.angles = new_angles

            # Update rotation guide (visual only, no full redraw needed)
            if self._molecule_drag_mode in ('rotate_z', 'rotate_3d'):
                if self._molecule_rotation_guide is None and self.main_ax:
                    self._molecule_rotation_guide = patches.Circle(
                        (mol.offset[0], mol.offset[1]), 
                        radius=2.0, # Fixed visual radius or dynamic based on molecule size
                        fill=False, edgecolor='yellow', linestyle='--', linewidth=1.5, alpha=0.6, zorder=40
                    )
                    self.main_ax.add_patch(self._molecule_rotation_guide)
                elif self._molecule_rotation_guide:
                    self._molecule_rotation_guide.center = (mol.offset[0], mol.offset[1])

            self._update_molecule_artists()

    def _on_molecule_release(self, event):
        if self._molecule_drag_idx is not None:
            self._molecule_drag_idx = None
            self._molecule_drag_start = None
            self._molecule_drag_start_px = None
            self._molecule_drag_mode = None
            self._molecule_drag_mol_angles = None
            
            if self._molecule_rotation_guide:
                try: self._molecule_rotation_guide.remove()
                except: pass
                self._molecule_rotation_guide = None
                self._redraw()
            self._molecule_drag_snapshot = False
        if self._pan_active:
            self._pan_active = False
            self._pan_ax = None
            self._pan_start = None
            self._pan_start_lim = None
            self._pan_last_ts = 0.0

    def set_molecule_palette(self, palette: str, notify: bool = True):
        palette = (palette or "cpk").lower()
        if palette not in available_atom_palettes():
            palette = "cpk"
        if palette == getattr(self, "molecule_palette", None):
            return
        self.molecule_palette = palette
        self._redraw()
        if notify and callable(self._molecule_palette_cb):
            try:
                self._molecule_palette_cb(palette)
            except Exception:
                pass

    def set_molecule_palette_callback(self, cb):
        self._molecule_palette_cb = cb

    def _show_molecule_menu(self, event, mol):
        style = QtWidgets.QApplication.style()
        icon = lambda std: style.standardIcon(std) if style else QtGui.QIcon()

        menu = QtWidgets.QMenu(self)

        # Properties
        props_act = menu.addAction(icon(QtWidgets.QStyle.SP_FileDialogDetailedView), "Properties (Rotate/Scale)...")
        menu.addSeparator()

        # View toggles
        toggle_shadow_act = menu.addAction(icon(QtWidgets.QStyle.SP_DialogYesButton), "Show shadows")
        toggle_shadow_act.setCheckable(True)
        toggle_shadow_act.setChecked(self._show_molecule_shadow)
        show_h_act = menu.addAction(icon(QtWidgets.QStyle.SP_TitleBarShadeButton), "Show hydrogens")
        show_h_act.setCheckable(True)
        show_h_act.setChecked(getattr(self, "_show_hydrogens", True))

        # Colors submenu
        colors_menu = menu.addMenu(icon(QtWidgets.QStyle.SP_DialogOpenButton), "Colors...")
        atom_color_act = colors_menu.addAction(icon(QtWidgets.QStyle.SP_DriveDVDIcon), "Set atom color...")
        atom_elem_color_act = colors_menu.addAction(icon(QtWidgets.QStyle.SP_FileIcon), "Set element color...")
        bond_color_act = colors_menu.addAction(icon(QtWidgets.QStyle.SP_FileDialogListView), "Set bond color...")
        reset_colors_act = colors_menu.addAction(icon(QtWidgets.QStyle.SP_BrowserReload), "Reset colors")

        menu.addSeparator()
        reset_all_act = menu.addAction(icon(QtWidgets.QStyle.SP_MessageBoxWarning), "Reset all molecules")

        # Undo
        undo_act = menu.addAction(icon(QtWidgets.QStyle.SP_ArrowBack), "Undo last change")
        undo_act.setShortcut(QtGui.QKeySequence("Ctrl+Z"))

        # Palette submenu
        pal_menu = menu.addMenu(icon(QtWidgets.QStyle.SP_DialogHelpButton), "Palette")
        current_pal = (getattr(self, "molecule_palette", "cpk") or "cpk").lower()
        palette_actions = {}
        for pal in available_atom_palettes():
            act = pal_menu.addAction(pal.title())
            act.setCheckable(True)
            act.setChecked(pal == current_pal)
            palette_actions[act] = pal

        menu.addSeparator()

        dup_act = menu.addAction(icon(QtWidgets.QStyle.SP_FileDialogNewFolder), "Duplicate")
        dup_act.setShortcut(QtGui.QKeySequence("Ctrl+D"))
        del_act = menu.addAction(icon(QtWidgets.QStyle.SP_TrashIcon), "Delete")
        del_act.setShortcut(QtGui.QKeySequence(QtCore.Qt.Key_Delete))

        action = menu.exec_(event.guiEvent.globalPos())
        if action == props_act:
            dlg = MoleculePropertiesDialog(mol, self, callback=self._redraw)
            dlg.show()
        elif action == atom_color_act:
            c = self._pick_color(mol.atom_color_override or get_atom_color('C', self.molecule_palette))
            if c:
                self._push_molecule_snapshot()
                mol.atom_color_override = c
                self._redraw()
        elif action == atom_elem_color_act:
            elem, ok = QtWidgets.QInputDialog.getText(self, "Element", "Element symbol:", text="C")
            if ok and elem.strip():
                c = self._pick_color(get_atom_color(elem.strip(), self.molecule_palette))
                if c:
                    self._push_molecule_snapshot()
                    m = getattr(mol, "atom_color_map", None)
                    if m is None:
                        mol.atom_color_map = {}
                        m = mol.atom_color_map
                    m[elem.strip().upper()] = c
                    self._redraw()
        elif action == bond_color_act:
            c = self._pick_color(mol.bond_color_override or "#e0e0e0")
            if c:
                self._push_molecule_snapshot()
                mol.bond_color_override = c
                mol.bond_color_mode = "single"
                self._redraw()
        elif action == reset_colors_act:
            self._push_molecule_snapshot()
            mol.atom_color_override = None
            mol.bond_color_override = None
            mol.bond_color_mode = 'default'
            mol.atom_color_map = {}
            self._redraw()
        elif action == reset_all_act:
            self.reset_molecules()
        elif action == undo_act:
            self.undo_last_molecule_change()
        elif action == toggle_shadow_act:
            self._show_molecule_shadow = toggle_shadow_act.isChecked()
            self._redraw()
        elif action == dup_act:
            self._push_molecule_snapshot()
            new_mol = mol.copy()
            new_mol.offset += np.array([1.0, 1.0, 0.0]) # Slight offset
            self.molecules.append(new_mol)
            self._redraw()
        elif action == del_act:
            if mol in self.molecules:
                self._push_molecule_snapshot()
                self.molecules.remove(mol)
                self._redraw()
        elif action in palette_actions:
            self._push_molecule_snapshot()
            self.set_molecule_palette(palette_actions[action])
        elif action == show_h_act:
            self._push_molecule_snapshot()
            self._show_hydrogens = show_h_act.isChecked()
            self._redraw()

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
        if view is None:
            return
        menu = QtWidgets.QMenu(self)
        # theme-aware styling
        try:
            if self._detail_dark:
                menu.setStyleSheet("QMenu { background: #1e1e24; color: #f5f5f5; } QMenu::item:selected { background: #2c2c34; }")
            else:
                menu.setStyleSheet("")
        except Exception:
            pass
        copy_act = menu.addAction("Copy image")
        copy_svg_act = menu.addAction("Copy view as SVG (vector)")
        copy_disp_png = menu.addAction("Copy displayed (PNG)")
        copy_disp_svg = menu.addAction("Copy displayed (SVG)")
        save_act = menu.addAction("Save image as...")
        save_svg_act = menu.addAction("Save view as SVG...")
        save_pdf_act = menu.addAction("Save view as PDF...")
        export_stp_act = menu.addAction("Export as WSxM STP...")
        
        menu.addSeparator()
        reset_zoom_act = menu.addAction("Reset Zoom")
        
        cbar_text = "Horizontal Colorbar" if self._colorbar_orientation == 'vertical' else "Vertical Colorbar"
        toggle_cbar_act = menu.addAction(cbar_text)

        menu.addSeparator()
        show_ticks_act = menu.addAction("Show Ticks")
        show_ticks_act.setCheckable(True)
        show_ticks_act.setChecked(self._show_ticks)
        show_cbar_act = menu.addAction("Show Colorbar")
        show_cbar_act.setCheckable(True)
        show_cbar_act.setChecked(self._show_colorbar)

        # Filters (provided by parent viewer)
        if callable(self._filter_menu_callback):
            menu.addSeparator()
            try:
                self._filter_menu_callback(menu, view, self)
            except Exception:
                pass
        
        angle_style_act = None
        if self.angle_enabled and self._angle_frames:
            menu.addSeparator()
            angle_style_act = menu.addAction("Use arrowheads for active angle")
            angle_style_act.setCheckable(True)
            active = self._get_active_angle_frame()
            angle_style_act.setChecked(active and active.get('style', 'dots') == 'arrows')

        menu.addSeparator()
        load_mol_act = menu.addAction("Load Molecule (XYZ/PDB)...")
        recent_menu = None
        recent_actions = {}
        recent_all = []
        for lst in (self._recent_molecule_paths, MultiPreviewCanvas._RECENT_MOLECULES):
            for p in lst:
                if p not in recent_all:
                    recent_all.append(p)
        if recent_all:
            recent_menu = menu.addMenu("Load Recent")
            for p in recent_all[:8]:
                act = recent_menu.addAction(Path(p).name)
                act.setToolTip(p)
                recent_actions[act] = p
        clear_mols_act = menu.addAction("Clear Molecules")

        arrange_act = None
        if callable(self._arrange_windows_callback):
            menu.addSeparator()
            arrange_act = menu.addAction("Arrange pop-outs")

        global_pos = None
        if event is not None:
            gui_event = getattr(event, "guiEvent", None)
            if gui_event is not None:
                try:
                    global_pos = gui_event.globalPos()
                except Exception:
                    global_pos = None
        if global_pos is None:
            global_pos = QtGui.QCursor.pos()

        chosen = menu.exec_(global_pos)
        if chosen == copy_act:
            self._copy_view_to_clipboard(view)
        elif chosen == copy_svg_act:
            self._copy_view_as_svg(view)
        elif chosen == copy_disp_png:
            self._copy_displayed("png")
        elif chosen == copy_disp_svg:
            self._copy_displayed("svg")
        elif chosen == save_act:
            self._save_view_to_file(view)
        elif chosen == save_svg_act:
            self._save_view_vector(view, "svg")
        elif chosen == save_pdf_act:
            self._save_view_vector(view, "pdf")
        elif chosen == export_stp_act:
            if callable(self._stp_export_callback):
                try:
                    self._stp_export_callback(view)
                except Exception:
                    pass
        elif chosen == reset_zoom_act:
            self._reset_view_zoom()
        elif chosen == toggle_cbar_act:
            self._toggle_colorbar_orientation()
        elif chosen == show_ticks_act:
            self._toggle_ticks()
        elif chosen == show_cbar_act:
            self._toggle_colorbar()
        elif chosen == load_mol_act:
            self._load_molecule_dialog()
        elif recent_menu and chosen in recent_actions:
            self.add_molecule(recent_actions[chosen])
        elif chosen == clear_mols_act:
            self._clear_molecules()
        elif arrange_act and chosen == arrange_act:
            try:
                self._arrange_windows_callback()
            except Exception:
                pass
        elif angle_style_act and chosen == angle_style_act:
            checked = angle_style_act.isChecked()
            self._update_active_angle_style('arrows' if checked else 'dots')
        elif chosen == reset_zoom_act:
            self._reset_view_zoom()

    def _save_view_to_file(self, view):
        try:
            tgt_view = view or (self.views[0] if self.views else {})
            title = tgt_view.get('title') or 'view'
            default = f"{title}.png"
            path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save view", default, "PNG Files (*.png)")
            if not path:
                return
            if len(self.views) > 1:
                fig = self._render_views_grid(self.views)
                buf = io.BytesIO()
                fig.savefig(buf, format='png', dpi=300, bbox_inches='tight')
                qimg = QtGui.QImage.fromData(buf.getvalue())
            else:
                qimg = self._view_to_qimage(tgt_view)
            qimg.save(path, "PNG")
        except Exception:
            QtWidgets.QMessageBox.warning(self, "Save view", "Unable to save image.")

    def _copy_displayed(self, fmt='png'):
        """Copy the current figure exactly as displayed (including overlays)."""
        buf = io.BytesIO()
        if fmt == 'svg':
            with matplotlib.rc_context({'svg.fonttype': 'none'}):
                self.fig.savefig(buf, format='svg', bbox_inches='tight')
            mime = QtCore.QMimeData()
            mime.setData("image/svg+xml", buf.getvalue())
            QtWidgets.QApplication.clipboard().setMimeData(mime)
        else:
            self.fig.savefig(buf, format='png', dpi=300, bbox_inches='tight')
            qimg = QtGui.QImage.fromData(buf.getvalue())
            QtWidgets.QApplication.clipboard().setImage(qimg)

    def _copy_view_as_svg(self, view):
        try:
            if len(self.views) > 1:
                fig = self._render_views_grid(self.views)
            else:
                fig = self._render_view_figure(view or (self.views[0] if self.views else {}))
            buf = io.BytesIO()
            with matplotlib.rc_context({'svg.fonttype': 'none'}):
                fig.savefig(buf, format="svg", bbox_inches="tight", pad_inches=0.02)
            svg_bytes = buf.getvalue()
            mime = QtCore.QMimeData()
            mime.setData("image/svg+xml", svg_bytes)
            QtWidgets.QApplication.clipboard().setMimeData(mime)
            if callable(self._copy_feedback_handler):
                self._copy_feedback_handler(view)
        except Exception:
            pass
        finally:
            try:
                import matplotlib.pyplot as _plt  # type: ignore
                _plt.close(fig)
            except Exception:
                pass

    def _save_view_vector(self, view, fmt):
        fmt = (fmt or "").strip().lower()
        if fmt not in ("svg", "pdf"):
            return
        try:
            tgt_view = view or (self.views[0] if self.views else {})
            title = tgt_view.get('title') or 'view'
            default = f"{title}.{fmt}"
            label = "SVG Files (*.svg)" if fmt == "svg" else "PDF Files (*.pdf)"
            path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save view", default, label)
            if not path:
                return
            if not path.lower().endswith(f".{fmt}"):
                path = f"{path}.{fmt}"
            if len(self.views) > 1:
                fig = self._render_views_grid(self.views)
            else:
                fig = self._render_view_figure(tgt_view)
            if fmt == 'svg':
                with matplotlib.rc_context({'svg.fonttype': 'none'}):
                    fig.savefig(path, format=fmt, bbox_inches="tight", pad_inches=0.02)
            else:
                fig.savefig(path, format=fmt, bbox_inches="tight", pad_inches=0.02)
            try:
                QtGui.QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(path))
            except Exception:
                pass
        except Exception:
            QtWidgets.QMessageBox.warning(self, "Save view", "Unable to save vector image.")

    def _reset_view_zoom(self):
        # Restore stored axis limits if present; otherwise no-op.
        did = False
        for ax, lims in list(self._zoom_reset_limits.items()):
            try:
                xlim, ylim = lims
                ax.set_xlim(xlim)
                ax.set_ylim(ylim)
                did = True
            except Exception:
                continue
        if did:
            self.draw_idle()

    def _toggle_colorbar_orientation(self):
        self._colorbar_orientation = 'horizontal' if self._colorbar_orientation == 'vertical' else 'vertical'
        self._redraw()

    def _toggle_ticks(self):
        self._show_ticks = not self._show_ticks
        self._redraw()

    def _toggle_colorbar(self):
        self._show_colorbar = not self._show_colorbar
        self._redraw()

    def _on_scroll_zoom(self, event):
        """Mouse wheel zoom centered at cursor."""
        ax = getattr(event, 'inaxes', None)
        if ax is None:
            return
        try:
            delta = 0
            btn = getattr(event, 'button', None)
            if btn in ('up', 'down'):
                delta = 1 if btn == 'up' else -1
            else:
                step = getattr(event, 'step', 0)
                if step:
                    delta = 1 if step > 0 else -1
            if not delta:
                return
            scale = 0.9 if delta > 0 else 1.1
        except Exception:
            scale = 0.9
        xlim = ax.get_xlim()
        ylim = ax.get_ylim()
        if ax not in self._zoom_reset_limits:
            self._zoom_reset_limits[ax] = (xlim, ylim)
        x0 = event.xdata if event.xdata is not None else (xlim[0] + xlim[1]) * 0.5
        y0 = event.ydata if event.ydata is not None else (ylim[0] + ylim[1]) * 0.5
        xr = abs(xlim[1] - xlim[0]) * scale
        yr = abs(ylim[1] - ylim[0]) * scale
        new_xlim = (x0 - xr * 0.5, x0 + xr * 0.5)
        new_ylim = (y0 - yr * 0.5, y0 + yr * 0.5)
        base_xlim, base_ylim = self._zoom_reset_limits.get(ax, (xlim, ylim))
        new_xlim = self._clamp_limits(new_xlim, base_xlim)
        new_ylim = self._clamp_limits(new_ylim, base_ylim)
        ax.set_xlim(new_xlim)
        ax.set_ylim(new_ylim)
        self.draw_idle()

    def _clamp_limits(self, new_lim, base_lim):
        """Clamp new limits to stay within base limits, preserving axis direction."""
        try:
            base_min = min(base_lim[0], base_lim[1])
            base_max = max(base_lim[0], base_lim[1])
            new_min = min(new_lim[0], new_lim[1])
            new_max = max(new_lim[0], new_lim[1])
            width = new_max - new_min
            base_width = base_max - base_min
            if width >= base_width:
                return base_lim
            start = max(base_min, min(new_min, base_max - width))
            end = start + width
            return (start, end) if base_lim[0] <= base_lim[1] else (end, start)
        except Exception:
            return new_lim

    def _is_zoomed(self, ax):
        """Return True if current limits differ from stored reset limits."""
        try:
            if ax not in self._zoom_reset_limits:
                self._zoom_reset_limits[ax] = (ax.get_xlim(), ax.get_ylim())
                return False
            base_xlim, base_ylim = self._zoom_reset_limits.get(ax, (ax.get_xlim(), ax.get_ylim()))
            cur_xlim, cur_ylim = ax.get_xlim(), ax.get_ylim()
            tol = 1e-9
            zoomed_x = abs(cur_xlim[0] - base_xlim[0]) > tol or abs(cur_xlim[1] - base_xlim[1]) > tol
            zoomed_y = abs(cur_ylim[0] - base_ylim[0]) > tol or abs(cur_ylim[1] - base_ylim[1]) > tol
            return zoomed_x or zoomed_y
        except Exception:
            return False

    def _session_signature_for_view(self, view):
        meta = view.get("meta") or {}
        file_path = meta.get("file_path") or meta.get("path") or ""
        channel = meta.get("channel_index")
        try:
            channel = int(channel) if channel is not None else None
        except Exception:
            channel = None
        return {
            "file": str(file_path),
            "channel": channel,
            "crop_sequence": view.get("crop_sequence"),
            "title": view.get("title"),
        }

    @staticmethod
    def _signature_key(signature):
        if not signature:
            return None
        return (
            signature.get("file"),
            signature.get("channel"),
            signature.get("crop_sequence"),
            signature.get("title"),
        )

    def export_zoom_states(self):
        states = []
        for ax, view in self._ax_view_map.items():
            sig = self._session_signature_for_view(view)
            key = self._signature_key(sig)
            if key is None:
                continue
            try:
                cur_xlim = tuple(ax.get_xlim())
                cur_ylim = tuple(ax.get_ylim())
            except Exception:
                continue
            base_xlim, base_ylim = self._zoom_reset_limits.get(ax, (cur_xlim, cur_ylim))
            states.append(
                {
                    "signature": sig,
                    "xlim": cur_xlim,
                    "ylim": cur_ylim,
                    "base_xlim": tuple(base_xlim),
                    "base_ylim": tuple(base_ylim),
                }
            )
        return states

    def apply_zoom_states(self, states):
        if not states:
            return
        lookup = {}
        for ax, view in self._ax_view_map.items():
            sig = self._session_signature_for_view(view)
            key = self._signature_key(sig)
            if key is not None:
                lookup[key] = ax
        updated = False
        for entry in states:
            key = self._signature_key(entry.get("signature"))
            if key is None:
                continue
            ax = lookup.get(key)
            if ax is None:
                continue
            xlim = entry.get("xlim")
            ylim = entry.get("ylim")
            try:
                if xlim:
                    ax.set_xlim(float(xlim[0]), float(xlim[1]))
                if ylim:
                    ax.set_ylim(float(ylim[0]), float(ylim[1]))
            except Exception:
                continue
            base_xlim = entry.get("base_xlim")
            base_ylim = entry.get("base_ylim")
            if base_xlim and base_ylim:
                try:
                    self._zoom_reset_limits[ax] = (tuple(base_xlim), tuple(base_ylim))
                except Exception:
                    pass
            updated = True
        if updated:
            self.draw_idle()

    def _display_extent_for_view(self, view, extent):
        """Return the extent that should be passed to matplotlib based on relative axes."""
        if extent is None:
            return None
        if not self._use_relative_axes(view):
            return self._normalize_extent(extent)
        try:
            x0, x1, y1, y0 = extent
        except Exception:
            return self._normalize_extent(extent)
        width = max(abs(x1 - x0), 1e-6)
        height = max(abs(y0 - y1), 1e-6)
        return self._normalize_extent((0.0, width, 0.0, height))

    def _normalize_extent(self, extent):
        if not extent or len(extent) != 4:
            return extent
        x0, x1, y1, y0 = extent
        tol = 1e-6
        if abs(x1 - x0) < tol:
            x1 = x0 + tol
        if abs(y1 - y0) < tol:
            y0 = y1 + tol
        return (x0, x1, y1, y0)

    def _render_view_figure(self, view):
        fig = Figure(figsize=(6, 6))
        ax = fig.add_subplot(1, 1, 1)
        arr = np.asarray(view.get('arr'))
        flip = self._use_relative_axes(view)
        if flip:
            arr_plot = np.flipud(arr)
        else:
            arr_plot = arr
        raw_extent = view.get('extent_raw')
        if raw_extent is None:
            raw_extent = view.get('extent')
        cmap = view.get('cmap', 'viridis')
        origin = 'lower' if flip else 'upper'
        display_extent = self._display_extent_for_view(view, raw_extent)
        if display_extent is None:
            im = ax.imshow(arr_plot, origin=origin, interpolation='nearest', cmap=cmap)
        else:
            im = ax.imshow(
                arr_plot,
                extent=display_extent,
                origin=origin,
                interpolation='nearest',
                aspect='equal',
                cmap=cmap,
            )
        # Ensure axes limits reflect the current extent (important when toggling relative axes)
        try:
            ext = display_extent if display_extent is not None else im.get_extent()
            if ext is not None:
                x0, x1, y1, y0 = ext
                ax.set_xlim(x0, x1)
                if flip:
                    ax.set_ylim(y0, y1)
                else:
                    ax.set_ylim(y1, y0)
        except Exception:
            pass
        ax.set_autoscale_on(False)
        cbar_label = view.get('colorbar_label') or view.get('unit', '')
        cbar = None
        if cbar_label and self._show_colorbar:
            try:
                divider = make_axes_locatable(ax)
                if self._colorbar_orientation == 'horizontal':
                    cax = divider.append_axes("bottom", size="5%", pad=0.08)
                    cbar = fig.colorbar(im, cax=cax, orientation='horizontal')
                    cbar.set_label(cbar_label)
                    cbar.ax.xaxis.set_label_coords(0.5, 0.5)
                    cbar.ax.xaxis.label.set_horizontalalignment('center')
                    cbar.ax.xaxis.label.set_verticalalignment('center')
                else:
                    cax = divider.append_axes("right", size="4%", pad=0.02)
                    cbar = fig.colorbar(im, cax=cax, orientation='vertical')
                    cbar.set_label(cbar_label)
                    cbar.ax.yaxis.set_label_coords(0.5, 0.5)
                    cbar.ax.yaxis.label.set_horizontalalignment('center')
                    cbar.ax.yaxis.label.set_verticalalignment('center')
            except Exception:
                cbar = fig.colorbar(im, ax=ax, fraction=0.08, pad=0.02, orientation=self._colorbar_orientation)
                cbar.set_label(cbar_label)
            if not self._show_ticks:
                cbar.set_ticks([])
        try:
            self._draw_outlines(ax, view)
        except Exception:
            pass
        title = view.get('title', '')
        if title and self._show_title:
            ax.set_title(title, fontsize=9)
        ax.tick_params(labelsize=8)

        if self.scale_bar_enabled:
            extent_for_scale = display_extent if display_extent is not None else raw_extent
            if extent_for_scale is None:
                h, w = np.shape(view['arr'])
                width = w
                unit = 'px'
            else:
                width = abs(extent_for_scale[1] - extent_for_scale[0])
                unit = view.get('axis_unit') or 'nm'
            
            size, label = self._calculate_best_scale_bar(width, unit)
            # Hide unit text if blank to avoid default "nm" showing up when unset
            label = label if label and label.strip() else None
            font_scale = getattr(self, '_view_font_scale', 1.0)
            
            dark = bool(self._detail_dark)
            default_color = '#f5f5f5' if dark else '#111111'
            sb_settings = getattr(self, '_scale_bar_settings', {})
            sb_text_col = sb_settings.get('text_color') or default_color
            sb_bar_col = sb_settings.get('bar_color') or default_color
            font_family = sb_settings.get('font_family', 'sans-serif')

            sb = AnchoredSizeBar(ax.transData, size, label, loc='center',
                                 pad=0.4, borderpad=0, sep=3, frameon=False,
                                 size_vertical=width*0.004*font_scale, color=sb_bar_col,
                                 label_top=True,
                                 bbox_to_anchor=self._scale_bar_pos, bbox_transform=ax.transAxes)
            sb.size_bar.get_children()[0].set_linewidth(0)
            text = sb.txt_label.get_children()[0]
            text.set_color(sb_text_col)
            text.set_fontfamily(font_family)
            text.set_fontsize(10 * font_scale)
            text.set_fontweight('bold')
            ax.add_artist(sb)

        if not self._show_ticks:
            ax.set_xticks([])
            ax.set_yticks([])

        self._draw_molecules(ax)

        self._style_export_figure(fig, ax, cbar)
        try:
            fig.tight_layout()
        except Exception:
            pass
        return fig

    def render_crop_entry_figure(self, entry):
        if not entry:
            return None
        view = entry.get("view_snapshot")
        if not view:
            return None
        try:
            return self._render_view_figure(view)
        except Exception:
            return None

    def _render_views_grid(self, views):
        """Render multiple views into a single figure grid."""
        views = views or []
        total = len(views)
        if total == 0:
            return self._render_view_figure({})
        cols = int(math.ceil(math.sqrt(total)))
        rows = int(math.ceil(total / cols))
        fig = Figure(figsize=(6 * cols, 6 * rows))
        dark = bool(self._detail_dark)
        fig_face = '#111217' if dark else '#ffffff'
        fig.set_facecolor(fig_face)
        text_color = '#f5f5f5' if dark else '#111111'
        font_scale = getattr(self, '_view_font_scale', 1.0)
        for i, view in enumerate(views, 1):
            ax = fig.add_subplot(rows, cols, i)
            arr = np.asarray(view.get('arr'))
            flip = self._use_relative_axes(view)
            arr_plot = np.flipud(arr) if flip else arr
            raw_extent = view.get('extent_raw')
            if raw_extent is None:
                raw_extent = view.get('extent')
            cmap = view.get('cmap', 'viridis')
            origin = 'lower' if flip else 'upper'
            display_extent = self._display_extent_for_view(view, raw_extent)
            if display_extent is None:
                im = ax.imshow(arr_plot, origin=origin, interpolation='nearest', cmap=cmap)
            else:
                im = ax.imshow(
                    arr_plot,
                    extent=display_extent,
                    origin=origin,
                    interpolation='nearest',
                    aspect='equal',
                    cmap=cmap,
                )
            try:
                ext = display_extent if display_extent is not None else im.get_extent()
                if ext is not None:
                    x0, x1, y1, y0 = ext
                    ax.set_xlim(x0, x1)
                    if flip:
                        ax.set_ylim(y0, y1)
                    else:
                        ax.set_ylim(y1, y0)
            except Exception:
                pass
            # record base limits for reset before any restore
            try:
                self._zoom_reset_limits[ax] = (ax.get_xlim(), ax.get_ylim())
            except Exception:
                pass
            ax.set_autoscale_on(False)
            if not self._show_ticks:
                ax.set_xticks([])
                ax.set_yticks([])
            ax.tick_params(labelsize=8 * font_scale, colors=text_color, labelcolor=text_color)
            for spine in ax.spines.values():
                spine.set_color(text_color)
            cbar_label = view.get('colorbar_label') or view.get('unit', '')
            if cbar_label and self._show_colorbar:
                try:
                    divider = make_axes_locatable(ax)
                    cax = divider.append_axes("right", size="5%", pad=0.05)
                    cbar = fig.colorbar(im, cax=cax, orientation='vertical')
                    cbar.set_label(cbar_label, size=10 * font_scale)
                    cbar.ax.yaxis.label.set_color(text_color)
                    cbar.ax.tick_params(colors=text_color, labelcolor=text_color, labelsize=8 * font_scale)
                    if not self._show_ticks:
                        cbar.set_ticks([])
                    cbar.outline.set_edgecolor(text_color)
                except Exception:
                    pass
            try:
                self._draw_outlines(ax, view)
            except Exception:
                pass
            title = view.get('title', '') or view.get('label', '')
            if title and self._show_title:
                ax.set_title(title, fontsize=9 * font_scale, color=text_color)
        fig.tight_layout()
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
        scale = max(0.6, min(2.5, getattr(self, '_view_font_scale', 1.0)))
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

    # ------------------------------------------------------------------ #
    # Outlines                                                           #
    # ------------------------------------------------------------------ #
    def set_outline_percentile(self, pct: float):
        """Set outline threshold (0-1 fraction)."""
        try:
            pct = float(pct)
        except Exception:
            return
        self._outline_threshold = max(0.05, min(0.99, pct))

    def _outline_key(self, view):
        meta = view.get("meta") or {}
        extent = view.get("extent_raw")
        if extent is None:
            extent = view.get("extent")
        return (meta.get("file_path"), tuple(extent or ()))

    def _guess_view_unit(self, view):
        unit = ""
        if view:
            unit = (view.get("unit") or view.get("colorbar_label") or "").strip()
        if not unit:
            unit = "nm"
        return unit

    def _add_outlines(self, view, contours_world):
        """Add one or more outlines for a view and track order for undo."""
        key = self._outline_key(view)
        if key is None:
            return
        entries = self._outlines.setdefault(key, [])
        for contour in contours_world:
            entry = {
                "pts": np.asarray(contour, dtype=float),
                "color": "#ffffff",
                "lw": 1.6,
                "ls": (0, (6, 4)),
            }
            entries.append(entry)
            self._outline_order.append((key, len(entries) - 1))

    def _draw_outlines(self, ax, view):
        outlines = self._outlines.get(self._outline_key(view), [])
        if not outlines:
            return
        for entry in outlines:
            pts = entry.get("pts") if isinstance(entry, dict) else entry
            if pts is None or len(pts) < 2:
                continue
            color = entry.get("color", self._outline_default_color) if isinstance(entry, dict) else self._outline_default_color
            lw = entry.get("lw", self._outline_default_lw) if isinstance(entry, dict) else self._outline_default_lw
            ls = entry.get("ls", self._outline_default_ls) if isinstance(entry, dict) else self._outline_default_ls
            try:
                line, = ax.plot(
                    pts[:, 0],
                    pts[:, 1],
                    color=color,
                    linewidth=lw,
                    alpha=0.9,
                    linestyle=ls,
                )
                line.set_path_effects([
                    PathEffects.withStroke(linewidth=lw * 1.5, foreground="black", alpha=0.4)
                ])
            except Exception:
                continue

    def _reset_outline_state(self):
        try:
            if self._outline_rect is not None:
                self._outline_rect.remove()
        except Exception:
            pass
        self._outline_start = None
        self._outline_rect = None
        self._outline_ax = None

    def clear_outlines(self, view=None):
        """Clear stored outlines for all views or a specific view."""
        if view is None:
            self._outlines.clear()
            self._outline_order.clear()
        else:
            try:
                key = self._outline_key(view)
                self._outlines.pop(key, None)
                self._outline_order = [(k, idx) for (k, idx) in self._outline_order if k != key]
            except Exception:
                pass
        self._redraw()

    def _outline_hit_test(self, view, xdata, ydata, ax=None, event=None, tol_px=10, tol_frac=0.02):
        """Return True if a click is close to any outline in the view."""
        outlines = self._outlines.get(self._outline_key(view), [])
        if not outlines:
            return False
        # Pixel-based hit test for better usability (fall back to data tolerance).
        try:
            if ax is not None:
                click_px = None
                try:
                    ex = getattr(event, "x", None)
                    ey = getattr(event, "y", None)
                    if ex is not None and ey is not None:
                        click_px = np.array([float(ex), float(ey)])
                except Exception:
                    click_px = None
                if click_px is None and xdata is not None and ydata is not None:
                    try:
                        click_px = np.array(ax.transData.transform((xdata, ydata)), dtype=float)
                    except Exception:
                        click_px = None
                if click_px is not None:
                    tol_px = max(1.0, float(tol_px))
                    for entry in outlines:
                        pts = entry.get("pts") if isinstance(entry, dict) else entry
                        if pts is None or len(pts) < 2:
                            continue
                        try:
                            pts_px = ax.transData.transform(pts)
                        except Exception:
                            pts_px = None
                        if pts_px is None:
                            continue
                        diffs = pts_px - click_px
                        d2 = np.sum(diffs * diffs, axis=1)
                        if np.min(d2) <= tol_px * tol_px:
                            return True
        except Exception:
            pass
        try:
            xs = []
            ys = []
            for entry in outlines:
                pts = entry.get("pts") if isinstance(entry, dict) else entry
                if pts is None or len(pts) == 0:
                    continue
                xs.extend(pts[:, 0])
                ys.extend(pts[:, 1])
            if not xs or not ys:
                return False
            xr = max(xs) - min(xs)
            yr = max(ys) - min(ys)
            tol = max(xr, yr) * tol_frac
            for entry in outlines:
                pts = entry.get("pts") if isinstance(entry, dict) else entry
                if pts is None or len(pts) < 2:
                    continue
                diffs = pts - np.array([xdata, ydata])
                d2 = np.sum(diffs * diffs, axis=1)
                if np.min(d2) <= tol * tol:
                    return True
        except Exception:
            return False
        return False

    def _show_outline_menu(self, view):
        """Context menu to tweak outline style and clear/undo."""
        menu = QtWidgets.QMenu(self)
        act_color = menu.addAction("Change color…")
        width_menu = menu.addMenu("Line width")
        style_menu = menu.addMenu("Line style")
        for w in (0.8, 1.2, 1.6, 2.0, 3.0):
            a = width_menu.addAction(f"{w:.1f}")
            a.setData(("width", w))
        styles = {
            "Solid": ("solid", "solid"),
            "Dashed": ("dashed", (0, (6, 4))),
            "Dotted": ("dotted", (0, (2, 4))),
            "Dense dash": ("dense", (0, (4, 2))),
        }
        for name, val in styles.items():
            a = style_menu.addAction(name)
            a.setData(("style", val))
        menu.addSeparator()
        act_undo = menu.addAction("Undo last outline")
        act_clear = menu.addAction("Clear outlines")
        chosen = menu.exec_(QtGui.QCursor.pos())
        if chosen is None:
            return
        data = chosen.data()
        if chosen == act_color:
            col = QtWidgets.QColorDialog.getColor(QtGui.QColor("white"), self, "Outline color")
            if col.isValid():
                self._set_outline_style(view, color=col.name())
        elif chosen == act_undo:
            self._undo_last_outline()
        elif chosen == act_clear:
            self.clear_outlines(view=view)
        elif data and isinstance(data, tuple):
            kind, val = data
            if kind == "width":
                self._set_outline_style(view, lw=float(val))
            elif kind == "style":
                name, ls = val
                self._set_outline_style(view, ls=ls)

    def _set_outline_style(self, view, color=None, lw=None, ls=None):
        """Apply style settings to all outlines of a view."""
        key = self._outline_key(view)
        entries = self._outlines.get(key, [])
        new_entries = []
        for e in entries:
            if isinstance(e, dict):
                entry = dict(e)
            else:
                entry = {"pts": np.asarray(e), "color": "#ffffff", "lw": 1.6, "ls": (0, (6, 4))}
            if color is not None:
                entry["color"] = color
            if lw is not None:
                entry["lw"] = lw
            if ls is not None:
                entry["ls"] = ls
            new_entries.append(entry)
        self._outlines[key] = new_entries
        self._redraw()

    def _undo_last_outline(self):
        """Undo the most recently added outline, if any. Returns True if something was undone."""
        while self._outline_order:
            key, idx = self._outline_order.pop()
            outlines = self._outlines.get(key)
            if outlines and 0 <= idx < len(outlines):
                try:
                    outlines.pop(idx)
                    if not outlines:
                        self._outlines.pop(key, None)
                    self._redraw()
                    return True
                except Exception:
                    continue
        return False

    def _mask_component_from_seed(self, roi: np.ndarray, seed: tuple | None = None):
        """Return a binary mask for the component containing the seed (or None)."""
        if roi.size == 0:
            return None
        # Blur slightly
        if _HAS_SCIPY and ndimage is not None:
            roi_blur = ndimage.gaussian_filter(roi, sigma=1.2)
        else:
            roi_blur = roi
        # Candidate thresholds: Otsu then percentiles
        thresholds = []
        if _HAS_SKIMAGE and sk_filters is not None:
            try:
                thresholds.append(float(sk_filters.threshold_otsu(roi_blur)))
            except Exception:
                pass
        thresholds.extend([
            np.percentile(roi_blur, p) for p in (90, 85, 80, 75, 70, 65, 60)
        ])
        for thresh_val in thresholds:
            try:
                mask = roi_blur >= thresh_val
                if _HAS_SKIMAGE and sk_morph is not None:
                    mask = sk_morph.remove_small_objects(mask, min_size=max(8, mask.size // 800))
                    mask = sk_morph.binary_closing(mask, sk_morph.disk(1))
                elif _HAS_SCIPY and ndimage is not None:
                    mask = ndimage.binary_opening(mask)
                    mask = ndimage.binary_closing(mask)
                if _HAS_SCIPY and ndimage is not None:
                    labels, nlab = ndimage.label(mask)
                    if nlab == 0:
                        continue
                    sr, sc = (int(seed[0]), int(seed[1])) if seed is not None else (None, None)
                    target_lbl = None
                    if sr is not None and 0 <= sr < labels.shape[0] and 0 <= sc < labels.shape[1]:
                        lbl = labels[sr, sc]
                        if lbl > 0:
                            target_lbl = lbl
                    if target_lbl is None:
                        sizes = ndimage.sum(mask, labels, index=range(1, nlab + 1))
                        target_lbl = 1 + int(np.argmax(sizes))
                    if target_lbl > 0:
                        return labels == target_lbl
                else:
                    # No labeling available; fallback: require seed to be foreground
                    sr, sc = (int(seed[0]), int(seed[1])) if seed is not None else (None, None)
                    if sr is not None and 0 <= sr < mask.shape[0] and 0 <= sc < mask.shape[1] and mask[sr, sc]:
                        return mask
            except Exception:
                continue
        # OpenCV flood-fill fallback
        if _HAS_CV2 and cv2 is not None and seed is not None:
            try:
                r8 = roi.astype(np.float32)
                r_norm = cv2.normalize(r8, None, 0, 255, cv2.NORM_MINMAX)
                flooded = r_norm.copy()
                mask_ff = np.zeros((r_norm.shape[0] + 2, r_norm.shape[1] + 2), np.uint8)
                cv2.floodFill(flooded, mask_ff, (int(seed[1]), int(seed[0])), 255, 5, 5, flags=4 | cv2.FLOODFILL_MASK_ONLY)
                ff_mask = (mask_ff[1:-1, 1:-1] > 0)
                if ff_mask.any():
                    return ff_mask
            except Exception:
                pass
        # Pure NumPy flood from seed as last resort
        if seed is not None:
            sr, sc = int(seed[0]), int(seed[1])
            if 0 <= sr < roi.shape[0] and 0 <= sc < roi.shape[1]:
                # Use the last computed mask from thresholds if available, else simple percentile
                try:
                    base_mask = roi_blur >= np.percentile(roi_blur, 75)
                except Exception:
                    base_mask = roi_blur >= roi_blur.mean()
                if base_mask[sr, sc]:
                    visited = np.zeros_like(base_mask, dtype=np.bool_)
                    stack = [(sr, sc)]
                    visited[sr, sc] = True
                    h, w = base_mask.shape
                    while stack:
                        r, c = stack.pop()
                        for dr, dc in ((1,0), (-1,0), (0,1), (0,-1)):
                            nr, nc = r + dr, c + dc
                            if 0 <= nr < h and 0 <= nc < w and not visited[nr, nc] and base_mask[nr, nc]:
                                visited[nr, nc] = True
                                stack.append((nr, nc))
                    if visited.any():
                        return visited
        return None

    def _contours_from_mask(self, mask: np.ndarray) -> List[np.ndarray]:
        """Return list of contour point arrays for a binary mask in pixel coords."""
        outlines_mask: List[np.ndarray] = []
        if mask is None or mask.size == 0:
            return outlines_mask
        # Preferred contour extraction
        if _HAS_SKIMAGE and sk_measure is not None:
            try:
                outlines_mask = [np.array(c) for c in sk_measure.find_contours(mask.astype(float), 0.5) if len(c) >= 4]
                if outlines_mask and _HAS_SCIPY and ndimage is not None:
                    try:
                        smoothed = []
                        for c in outlines_mask:
                            if c.shape[0] > 10:
                                r_s = ndimage.gaussian_filter1d(c[:, 0], sigma=1.0, mode='wrap')
                                c_s = ndimage.gaussian_filter1d(c[:, 1], sigma=1.0, mode='wrap')
                                smoothed.append(np.column_stack([r_s, c_s]))
                            else:
                                smoothed.append(c)
                        outlines_mask = smoothed
                    except Exception:
                        pass
            except Exception:
                outlines_mask = []
        # Fallback: perimeter then convex hull
        if not outlines_mask:
            try:
                if _HAS_SCIPY and ndimage is not None:
                    eroded = ndimage.binary_erosion(mask)
                    perim = mask & ~eroded
                else:
                    perim = mask.copy()
                    perim[1:-1, 1:-1] &= ~mask[:-2, 1:-1] | ~mask[2:, 1:-1] | ~mask[1:-1, :-2] | ~mask[1:-1, 2:]
                coords = np.argwhere(perim)
                if coords.shape[0] >= 4:
                    if ConvexHull is not None:
                        try:
                            hull = ConvexHull(coords)
                            coords = coords[hull.vertices]
                        except Exception:
                            pass
                    outlines_mask = [coords]
            except Exception:
                outlines_mask = []
        # Last resort: all foreground pixels
        if not outlines_mask:
            ys, xs = np.nonzero(mask)
            if len(xs) > 0:
                outlines_mask = [np.column_stack([ys, xs])]
        return outlines_mask

    def _finish_outline_drag(self, event):
        """Finalize an outline selection and store the dashed contour."""
        # Drag-based outlines are deprecated in favor of click-based extraction,
        # but keep this for Alt+drag legacy use.
        if self._outline_start is None or self._outline_ax is None:
            return
        if getattr(event, "button", None) != 1 or event.inaxes is not self._outline_ax:
            # Only respond to matching left-button release inside the same axes
            self._reset_outline_state()
            return
        view = self._ax_view_map.get(self._outline_ax)
        if view is None:
            self._reset_outline_state()
            return
        self._outline_from_point(view, self._outline_ax, *self._outline_start, drag_end=(event.xdata, event.ydata))
        self._reset_outline_state()

    def _outline_from_point(self, view, ax, xdata, ydata, drag_end=None):
        """Create an adaptive outline around the dominant blob near a clicked point."""
        if xdata is None or ydata is None or view is None or ax is None:
            return
        arr = np.asarray(view.get("arr"))
        if arr.size == 0:
            return
        flip = bool(view.get("relative_axes"))
        arr_disp = np.flipud(arr) if flip else arr
        h, w = arr_disp.shape[:2]
        xmin, xmax = ax.get_xlim()
        ymin, ymax = ax.get_ylim()
        if xmax == xmin or ymax == ymin:
            return
        # Map click to pixel indices (full-image seed; no ROI window)
        def _map_x(x):
            frac = (x - xmin) / (xmax - xmin)
            return int(np.clip(round(frac * (w - 1)), 0, w - 1))
        def _map_y(y):
            frac = (y - ymin) / (ymax - ymin)
            return int(np.clip(round(frac * (h - 1)), 0, h - 1))
        cx = _map_x(xdata)
        cy = _map_y(ydata)
        seed = (cy, cx)
        comp_mask = self._mask_component_from_seed(arr_disp, seed=seed)
        if comp_mask is None:
            # Fallback: tiny seed disk so the user sees some feedback
            comp_mask = np.zeros_like(arr_disp, dtype=bool)
            rr0, rr1 = max(0, cy - 3), min(h, cy + 4)
            cc0, cc1 = max(0, cx - 3), min(w, cx + 4)
            comp_mask[rr0:rr1, cc0:cc1] = True
        outlines_mask = self._contours_from_mask(comp_mask)
        if not outlines_mask:
            # Last resort: outline the tiny seed box
            outlines_mask = [np.array([[cy-3, cx-3],[cy-3, cx+3],[cy+3, cx+3],[cy+3, cx-3]])]
        ext = view.get("extent")
        outlines_world: List[np.ndarray] = []
        if ext is None:
            def _lerp_x_idx(idx):
                return xmin + (xmax - xmin) * (idx / max(1, w - 1))
            def _lerp_y_idx(idx):
                return ymin + (ymax - ymin) * (idx / max(1, h - 1))
        else:
            # Extent is (x0, x1, y0, y1) in data coords
            x_extent0, x_extent1, y_extent0, y_extent1 = ext
            def _lerp_x_idx(idx):
                return x_extent0 + (x_extent1 - x_extent0) * (idx / max(1, w - 1))
            def _lerp_y_idx(idx):
                return y_extent0 + (y_extent1 - y_extent0) * (idx / max(1, h - 1))
        for contour in outlines_mask:
            if contour is None or len(contour) < 2:
                continue
            pts_world = []
            for r_idx, c_idx in contour:
                pts_world.append((_lerp_x_idx(c_idx), _lerp_y_idx(r_idx)))
            if len(pts_world) >= 2:
                outlines_world.append(np.array(pts_world, dtype=float))
        if outlines_world:
            self._add_outlines(view, outlines_world)
            self._redraw()

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
        if self._outline_rect is not None and self._outline_start is not None and event.inaxes is self._outline_ax:
            try:
                x0, y0 = self._outline_start
                x1, y1 = event.xdata, event.ydata
                if x1 is not None and y1 is not None:
                    self._outline_rect.set_x(min(x0, x1))
                    self._outline_rect.set_y(min(y0, y1))
                    self._outline_rect.set_width(abs(x1 - x0))
                    self._outline_rect.set_height(abs(y1 - y0))
                    self.draw_idle()
            except Exception:
                pass
        if self._crop_rect is not None and self._crop_start is not None and event.inaxes is self._crop_ax:
            try:
                # Throttle drag updates to keep UI responsive
                now_ms = time.perf_counter() * 1000.0
                if (now_ms - getattr(self, "_crop_last_ts", 0.0)) < getattr(self, "_crop_move_throttle_ms", 12):
                    return
                self._crop_last_ts = now_ms
                x0, y0 = self._crop_start
                x1, y1 = event.xdata, event.ydata
                if x1 is not None and y1 is not None:
                    if self._crop_square:
                        dx = x1 - x0
                        dy = y1 - y0
                        side = min(abs(dx), abs(dy))
                        sx = 1 if dx >= 0 else -1
                        sy = 1 if dy >= 0 else -1
                        x1 = x0 + sx * side
                        y1 = y0 + sy * side
                    self._crop_rect.set_x(min(x0, x1))
                    self._crop_rect.set_y(min(y0, y1))
                    self._crop_rect.set_width(abs(x1 - x0))
                    self._crop_rect.set_height(abs(y1 - y0))
                    self.draw_idle()
            except Exception:
                pass
        if self._value_callback is None:
            return
        # Performance: skip value inspection while dragging profiles or molecules
        if getattr(self, '_dragging', None) is not None or getattr(self, '_molecule_drag_idx', None) is not None:
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

    def _on_crop_release(self, event):
        """Finish a crop drag and emit a cropped view copy."""
        if self._outline_start is not None and self._outline_ax is not None:
            self._finish_outline_drag(event)
            return
        if self._crop_start is None or self._crop_ax is None:
            return
        if getattr(event, "button", None) != 1 or event.inaxes is not self._crop_ax:
            # Only respond to the matching left-button release inside the same axes
            self._reset_crop_state()
            return
        x1, y1 = event.xdata, event.ydata
        x0, y0 = self._crop_start
        if x1 is not None and y1 is not None and self._crop_square:
            dx = x1 - x0
            dy = y1 - y0
            side = min(abs(dx), abs(dy))
            sx = 1 if dx >= 0 else -1
            sy = 1 if dy >= 0 else -1
            x1 = x0 + sx * side
            y1 = y0 + sy * side
        # clean up the rectangle artist
        try:
            if self._crop_rect is not None:
                self._crop_rect.remove()
                self._crop_rect = None
                self.draw_idle()
        except Exception:
            pass
        view = self._ax_view_map.get(self._crop_ax)
        if view is None or x1 is None or y1 is None:
            self._reset_crop_state()
            return
        arr = np.asarray(view.get("arr"))
        if arr.size == 0:
            self._reset_crop_state()
            return
        flip = bool(view.get("relative_axes"))
        arr_disp = np.flipud(arr) if flip else arr
        h, w = arr_disp.shape[:2]
        xlim0, xlim1 = self._crop_ax.get_xlim()
        ylim0, ylim1 = self._crop_ax.get_ylim()
        x_min, x_max = (xlim0, xlim1) if xlim0 <= xlim1 else (xlim1, xlim0)
        y_min, y_max = (ylim0, ylim1) if ylim0 <= ylim1 else (ylim1, ylim0)
        # clamp rectangle to axis limits (handles inverted axes)
        x0c = min(max(x0, x_min), x_max)
        x1c = min(max(x1, x_min), x_max)
        y0c = min(max(y0, y_min), y_max)
        y1c = min(max(y1, y_min), y_max)
        if xlim0 == xlim1 or ylim0 == ylim1:
            self._reset_crop_state()
            return
        def _map_x(x):
            frac = (x - xlim0) / (xlim1 - xlim0)
            return int(np.clip(round(frac * (w - 1)), 0, w - 1))
        def _map_y(y):
            frac = (y - ylim0) / (ylim1 - ylim0)
            return int(np.clip(round(frac * (h - 1)), 0, h - 1))
        c0, c1 = _map_x(x0c), _map_x(x1c)
        r0, r1 = _map_y(y0c), _map_y(y1c)
        if c1 < c0:
            c0, c1 = c1, c0
        if r1 < r0:
            r0, r1 = r1, r0
        cropped_disp = arr_disp[r0:r1 + 1, c0:c1 + 1]
        if cropped_disp.size == 0:
            self._reset_crop_state()
            return
        cropped_arr = np.flipud(cropped_disp) if flip else cropped_disp
        # Build new extent in data coordinates, preserving orientation
        ext = view.get("extent")
        if ext is None:
            crop_extent = None
        else:
            x_extent0, x_extent1, y_extent1, y_extent0 = ext
            def _lerp_x(idx):
                return x_extent0 + (x_extent1 - x_extent0) * (idx / (w - 1))
            def _lerp_y(idx):
                return y_extent0 + (y_extent1 - y_extent0) * (idx / (h - 1))
            crop_extent = (
                _lerp_x(c0),
                _lerp_x(c1),
                _lerp_y(r1),  # note: extent ordering is (x0, x1, y1, y0)
                _lerp_y(r0),
            )
        bounds_data = self._pixel_bounds_to_axis_bounds(self._crop_ax, w, h, c0, c1, r0, r1)
        entry = self._register_crop_entry(view, bounds_data, (c0, c1, r0, r1), self._crop_square, update_size=True)
        new_view = dict(view)
        try:
            new_view["arr"] = np.array(cropped_arr, copy=True)
        except Exception:
            new_view["arr"] = cropped_arr
        if crop_extent is not None:
            new_view["extent_raw"] = crop_extent
            display_extent = self._display_extent_for_view(new_view, crop_extent)
            if display_extent is not None:
                new_view["extent"] = display_extent
            else:
                new_view.pop("extent", None)
        else:
            new_view.pop("extent", None)
            new_view.pop("extent_raw", None)
        title = view.get("title") or "crop"
        new_view["title"] = f"{title} [crop]"
        if entry and entry.get("sequence") is not None:
            new_view["crop_sequence"] = entry["sequence"]
            entry["view_snapshot"] = dict(new_view)
        if callable(self._crop_callback):
            try:
                self._crop_callback(new_view)
            except Exception:
                pass
        ax = self._crop_ax
        if self._fixed_crop_history_visible and entry and ax:
            self._render_history_entry(ax, entry)
        if self._fixed_crop_template_visible and ax:
            self._render_template_overlay(ax, view)
        self.draw_idle()
        self._reset_crop_state()

    def _reset_crop_state(self):
        self._crop_start = None
        self._crop_rect = None
        self._crop_ax = None
        self._crop_square = False
        self._crop_last_ts = 0.0
        self._outline_start = None
        self._outline_rect = None
        self._outline_ax = None

    def enable_fixed_crop_quick_mode(self, enabled: bool):
        self._fixed_crop_quick_mode = bool(enabled)

    def show_fixed_crop_template(self, visible: bool):
        self._fixed_crop_template_visible = bool(visible)
        self._redraw()

    def show_fixed_crop_history(self, visible: bool):
        self._fixed_crop_history_visible = bool(visible)
        self._redraw()

    def is_fixed_crop_history_visible(self):
        return bool(self._fixed_crop_history_visible)

    def set_fixed_crop_history_highlight(self, seq):
        seq = int(seq) if seq is not None else None
        if self._fixed_crop_history_highlight_seq == seq:
            return
        self._fixed_crop_history_highlight_seq = seq
        self._update_highlight_artists()
        self.draw_idle()

    def _cleanup_highlight_artists(self):
        for artists in list(self._fixed_crop_history_highlight_artists.values()):
            for artist in artists:
                try:
                    artist.remove()
                except Exception:
                    pass
        self._fixed_crop_history_highlight_artists.clear()

    def _update_highlight_artists(self):
        if self._fixed_crop_history_visible:
            self._cleanup_highlight_artists()
            return
        seq = self._fixed_crop_history_highlight_seq
        if seq is None:
            self._cleanup_highlight_artists()
            return
        active_keys = set()
        for ax, view in self._ax_view_map.items():
            key = self._outline_key(view)
            if key is None:
                continue
            entry = next(
                (entry for entry in self._fixed_crop_history
                 if entry.get("key") == key and entry.get("sequence") == seq),
                None,
            )
            if entry is None:
                continue
            active_keys.add(key)
            x0, x1, y0, y1 = entry.get("data_bounds") or (0, 0, 0, 0)
            left, right = min(x0, x1), max(x0, x1)
            bottom, top = min(y0, y1), max(y0, y1)
            width = right - left
            height = top - bottom
            if width <= 0 or height <= 0:
                continue
            artists = self._fixed_crop_history_highlight_artists.get(key)
            if artists:
                fill, outline = artists
                fill.set_xy((left, bottom))
                fill.set_width(width)
                fill.set_height(height)
                outline.set_xy((left, bottom))
                outline.set_width(width)
                outline.set_height(height)
            else:
                fill = patches.Rectangle(
                    (left, bottom),
                    width,
                    height,
                    linewidth=0,
                    edgecolor='none',
                    facecolor='#ff66ff',
                    alpha=0.15,
                    zorder=17,
                )
                outline = patches.Rectangle(
                    (left, bottom),
                    width,
                    height,
                    linewidth=3.0,
                    edgecolor='#ffffff',
                    facecolor='none',
                    alpha=0.7,
                    linestyle='-',
                    zorder=19,
                )
                ax.add_patch(fill)
                ax.add_patch(outline)
                self._fixed_crop_history_highlight_artists[key] = (fill, outline)
        for key in list(self._fixed_crop_history_highlight_artists.keys()):
            if key not in active_keys:
                artists = self._fixed_crop_history_highlight_artists.pop(key, ())
                for artist in artists:
                    try:
                        artist.remove()
                    except Exception:
                        pass

    def _axis_coord_to_pixel(self, ax, coord, length, axis_key):
        if ax is None or coord is None or length <= 0:
            return 0
        limits = ax.get_xlim() if axis_key == 'x' else ax.get_ylim()
        lim0, lim1 = limits
        denom = max(1, length - 1)
        frac = (coord - lim0) / (lim1 - lim0) if lim1 != lim0 else 0.0
        idx = int(round(frac * denom))
        return int(np.clip(idx, 0, length - 1))

    def _pixel_bounds_to_axis_bounds(self, ax, w, h, c0, c1, r0, r1):
        if ax is None:
            return None
        def _interp(idx, lim, size):
            denom = max(1, size - 1)
            frac = idx / denom if denom else 0.0
            return lim[0] + (lim[1] - lim[0]) * frac
        xlim = ax.get_xlim()
        ylim = ax.get_ylim()
        left = min(c0, c1)
        right = max(c0, c1)
        bottom = min(r0, r1)
        top = max(r0, r1)
        return (
            _interp(left, xlim, w),
            _interp(right, xlim, w),
            _interp(bottom, ylim, h),
            _interp(top, ylim, h),
        )

    def _compute_crop_extent(self, view, w, h, c0, c1, r0, r1):
        ext = view.get("extent")
        if ext is None:
            return None
        x_extent0, x_extent1, y_extent1, y_extent0 = ext
        def _lerp(idx, start, end, denom):
            return start + (end - start) * (idx / denom) if denom != 0 else start
        denom_x = max(1, w - 1)
        denom_y = max(1, h - 1)
        return (
            _lerp(c0, x_extent0, x_extent1, denom_x),
            _lerp(c1, x_extent0, x_extent1, denom_x),
            _lerp(r1, y_extent0, y_extent1, denom_y),
            _lerp(r0, y_extent0, y_extent1, denom_y),
        )

    def _register_crop_entry(self, view, bounds_data, pixel_bounds, square, update_size=True):
        key = self._outline_key(view)
        if key is None:
            return None
        c0, c1, r0, r1 = pixel_bounds
        width = max(1, int(abs(c1 - c0) + 1))
        height = max(1, int(abs(r1 - r0) + 1))
        seq = self._fixed_crop_sequence
        real_unit = self._guess_view_unit(view)
        real_size = (0.0, 0.0)
        if bounds_data:
            x0, x1, y0, y1 = bounds_data
            real_size = (abs(x1 - x0), abs(y1 - y0))
        entry = {
            "key": key,
            "data_bounds": bounds_data,
            "pixel_bounds": pixel_bounds,
            "sequence": seq,
            "square": bool(square),
            "real_size": real_size,
            "unit": real_unit,
            "color": self._crop_color_for_seq(seq),
        }
        self._fixed_crop_history.append(entry)
        self._fixed_crop_sequence = seq + 1
        if update_size or self._fixed_crop_template is None:
            self._update_template_from_entry(entry)
        else:
            self._fixed_crop_template_bounds = bounds_data
            self._fixed_crop_template_pixel_bounds = pixel_bounds
            self._fixed_crop_template_view_key = key
        if len(self._fixed_crop_history) > _FIXED_CROP_HISTORY_LIMIT:
            self._fixed_crop_history.pop(0)
        self._emit_fixed_crop_history_update()
        return entry

    def _render_history_entry(self, ax, entry, show_label=True):
        if not entry or ax is None:
            return
        data_bounds = entry.get("data_bounds")
        if not data_bounds:
            return
        x0, x1, y0, y1 = data_bounds
        left, right = min(x0, x1), max(x0, x1)
        bottom, top = min(y0, y1), max(y0, y1)
        width = right - left
        height = top - bottom
        if width <= 0 or height <= 0:
            return
        seq = entry.get("sequence")
        is_active = seq is not None and seq == self._fixed_crop_history_highlight_seq
        edge_color = entry.get("color") or ('#ff66ff' if is_active else '#ffd166')
        line_width = 2.3 if is_active else 1.8
        alpha = 1.0 if is_active else 0.6
        if is_active:
            highlight_fill = patches.Rectangle(
                (left, bottom),
                width,
                height,
                linewidth=0,
                edgecolor='none',
                facecolor=edge_color,
                alpha=0.15,
                zorder=17,
            )
            ax.add_patch(highlight_fill)
        rect = patches.Rectangle(
            (left, bottom),
            width,
            height,
            linewidth=line_width,
            edgecolor=edge_color,
            facecolor='none',
            alpha=alpha,
            linestyle='-',
            zorder=18,
        )
        ax.add_patch(rect)
        if is_active:
            highlight_outline = patches.Rectangle(
                (left, bottom),
                width,
                height,
                linewidth=max(line_width + 1.2, 3.0),
                edgecolor='#ffffff',
                facecolor='none',
                alpha=0.75,
                linestyle='-',
                zorder=19,
            )
            ax.add_patch(highlight_outline)
            # handles
            handle_size = max(width, height) * 0.02
            for (cx, cy) in [(left, bottom), (left + width, bottom), (left, bottom + height), (left + width, bottom + height)]:
                h_rect = patches.Rectangle((cx - handle_size/2, cy - handle_size/2), handle_size, handle_size,
                                           facecolor=edge_color, edgecolor=edge_color, alpha=0.95, zorder=19)
                ax.add_patch(h_rect)
        if not show_label or seq is None:
            return
        label_offset_x = width * 0.02 if width > 0 else 0.0
        label_offset_y = height * 0.02 if height > 0 else 0.0
        label_x = left + label_offset_x
        label_y = top - label_offset_y
        real_size = entry.get("real_size", (0.0, 0.0))
        unit = entry.get("unit") or self._fixed_crop_template_unit or "nm"
        pixel_bounds = entry.get("pixel_bounds")
        px_label = ""
        if pixel_bounds:
            px_width = int(abs(pixel_bounds[1] - pixel_bounds[0]) + 1)
            px_height = int(abs(pixel_bounds[3] - pixel_bounds[2]) + 1)
            px_label = f"{px_width}×{px_height} px"
        real_label = ""
        if any(real_size):
            real_label = f"{real_size[0]:.2f} {unit} × {real_size[1]:.2f} {unit}"
        text_lines = [f"#{seq}"]
        if real_label:
            text_lines.append(real_label)
        if not real_label and px_label:
            text_lines.append(px_label)
        ax.text(
            label_x,
            label_y,
            "\n".join(text_lines),
            color=edge_color,
            fontsize=7,
            fontweight='bold',
            verticalalignment='top',
            horizontalalignment='left',
            bbox=dict(facecolor='#111111', alpha=0.65, pad=1, edgecolor='none'),
            zorder=19,
        )

    def _draw_fixed_crop_history(self, ax, view):
        key = self._outline_key(view)
        if key is None:
            return
        entries = [entry for entry in self._fixed_crop_history if entry.get("key") == key]
        if not entries:
            return
        if self._fixed_crop_history_visible:
            for entry in entries:
                self._render_history_entry(ax, entry)
        else:
            highlight_seq = self._fixed_crop_history_highlight_seq
            if highlight_seq is None:
                return
            entry = next((e for e in entries if e.get("sequence") == highlight_seq), None)
            if entry:
                self._render_history_entry(ax, entry, show_label=False)

    def _emit_fixed_crop_history_update(self):
        if callable(self._fixed_crop_history_callback):
            try:
                self._fixed_crop_history_callback(list(self._fixed_crop_history))
            except Exception:
                pass

    def set_fixed_crop_history_callback(self, cb):
        self._fixed_crop_history_callback = cb

    def _crop_color_for_seq(self, seq: int):
        palette = [
            "#ff66ff", "#66c2ff", "#ffa600", "#00c896",
            "#c084ff", "#ff6b6b", "#4dd0e1", "#ffd166",
        ]
        try:
            return palette[seq % len(palette)]
        except Exception:
            return palette[0]

    def _compute_template_bounds_from_pixels(self, view, ax, width_px, height_px):
        if view is None or ax is None:
            return None
        arr_obj = view.get("arr")
        if arr_obj is None:
            return None
        arr = np.asarray(arr_obj)
        if arr.size == 0:
            return None
        flip = bool(view.get("relative_axes"))
        arr_disp = np.flipud(arr) if flip else arr
        h, w = arr_disp.shape[:2]
        if w == 0 or h == 0:
            return None
        width = max(1, min(int(width_px), w))
        height = max(1, min(int(height_px), h))
        cx = w / 2.0
        cy = h / 2.0
        c0 = int(np.clip(int(round(cx - width / 2.0)), 0, max(0, w - width)))
        r0 = int(np.clip(int(round(cy - height / 2.0)), 0, max(0, h - height)))
        c1 = c0 + width - 1
        r1 = r0 + height - 1
        bounds = self._pixel_bounds_to_axis_bounds(ax, w, h, c0, c1, r0, r1)
        return bounds, (c0, c1, r0, r1)

    def _compute_pixels_for_real_size(self, view, ax, real_width, real_height):
        if view is None or ax is None:
            return None
        arr_obj = view.get("arr")
        if arr_obj is None:
            return None
        arr = np.asarray(arr_obj)
        if arr.size == 0:
            return None
        extent = view.get("extent")
        if not extent:
            return None
        x0, x1, y1, y0 = extent
        x_span = abs(x1 - x0)
        y_span = abs(y1 - y0)
        flip = bool(view.get("relative_axes"))
        arr_disp = np.flipud(arr) if flip else arr
        h, w = arr_disp.shape[:2]
        if w == 0 or h == 0:
            return None
        denom_x = max(1, w - 1)
        denom_y = max(1, h - 1)
        px_width = max(1, int(round((real_width * denom_x) / max(1e-6, x_span))))
        px_height = max(1, int(round((real_height * denom_y) / max(1e-6, y_span))))
        return px_width, px_height

    def _update_template_from_entry(self, entry):
        if not entry:
            return
        pixel_bounds = entry.get("pixel_bounds")
        if not pixel_bounds:
            return
        c0, c1, r0, r1 = pixel_bounds
        width = int(abs(c1 - c0) + 1)
        height = int(abs(r1 - r0) + 1)
        self._fixed_crop_template = {
            "width": width,
            "height": height,
            "square": bool(entry.get("square", False)),
        }
        self._fixed_crop_template_bounds = entry.get("data_bounds")
        self._fixed_crop_template_pixel_bounds = pixel_bounds
        self._fixed_crop_template_view_key = entry.get("key")
        self._fixed_crop_template_unit = entry.get("unit") or self._fixed_crop_template_unit
        self._fixed_crop_template_manual_dims = None

    def _maybe_restore_manual_template(self):
        dims = self._fixed_crop_template_manual_dims
        if not dims:
            return False
        ax = self.main_ax or next(iter(self._ax_view_map.keys()), None)
        view = self._ax_view_map.get(ax) if ax else None
        if dims.get("type") == "px":
            width = dims.get("width")
            height = dims.get("height")
            template = self._compute_template_bounds_from_pixels(view, ax, width, height)
            if template:
                bounds_data, pixel_bounds = template
                self._fixed_crop_template = {
                    "width": width,
                    "height": height,
                    "square": bool(self._fixed_crop_template.get("square", False)),
                }
                self._fixed_crop_template_bounds = bounds_data
                self._fixed_crop_template_pixel_bounds = pixel_bounds
                self._fixed_crop_template_view_key = self._outline_key(view)
                return True
        elif dims.get("type") == "real":
            width = dims.get("width")
            height = dims.get("height")
            unit = dims.get("unit") or self._fixed_crop_template_unit
            px_dims = self._compute_pixels_for_real_size(view, ax, width, height)
            if not px_dims:
                return False
            template = self._compute_template_bounds_from_pixels(view, ax, px_dims[0], px_dims[1])
            if template:
                bounds_data, pixel_bounds = template
                self._fixed_crop_template = {
                    "width": pixel_bounds[1] - pixel_bounds[0] + 1,
                    "height": pixel_bounds[3] - pixel_bounds[2] + 1,
                    "square": bool(self._fixed_crop_template.get("square", False)),
                }
                self._fixed_crop_template_bounds = bounds_data
                self._fixed_crop_template_pixel_bounds = pixel_bounds
                self._fixed_crop_template_view_key = self._outline_key(view)
                self._fixed_crop_template_unit = unit
                return True
        return False

    def set_fixed_crop_template_size(self, width, height, square=False):
        width = max(1, int(width))
        height = max(1, int(height))
        ax = self.main_ax or next(iter(self._ax_view_map.keys()), None)
        view = self._ax_view_map.get(ax) if ax else next(iter(self._ax_view_map.values()), None)
        if view is None or ax is None:
            return False
        template = self._compute_template_bounds_from_pixels(view, ax, width, height)
        if not template:
            return False
        bounds_data, pixel_bounds = template
        self._fixed_crop_template = {
            "width": width,
            "height": height,
            "square": bool(square),
        }
        self._fixed_crop_template_bounds = bounds_data
        self._fixed_crop_template_pixel_bounds = pixel_bounds
        self._fixed_crop_template_view_key = self._outline_key(view)
        self._fixed_crop_template_manual_dims = {
            "type": "px",
            "width": width,
            "height": height,
        }
        self.draw_idle()
        return True

    def get_fixed_crop_template_size(self):
        if not self._fixed_crop_template:
            return (0, 0)
        return (int(self._fixed_crop_template.get("width", 0)), int(self._fixed_crop_template.get("height", 0)))

    def get_fixed_crop_template_real_size(self):
        if not self._fixed_crop_template_bounds:
            return (0.0, 0.0, self._fixed_crop_template_unit or "nm")
        x0, x1, y0, y1 = self._fixed_crop_template_bounds
        return (abs(x1 - x0), abs(y1 - y0), self._fixed_crop_template_unit or "nm")

    def set_fixed_crop_template_real_size(self, real_width, real_height, square=False):
        real_width = float(real_width)
        real_height = float(real_height)
        ax = self.main_ax or next(iter(self._ax_view_map.keys()), None)
        view = self._ax_view_map.get(ax) if ax else next(iter(self._ax_view_map.values()), None)
        if view is None or ax is None:
            return False
        px_dims = self._compute_pixels_for_real_size(view, ax, real_width, real_height)
        if not px_dims:
            return False
        template = self._compute_template_bounds_from_pixels(view, ax, px_dims[0], px_dims[1])
        if not template:
            return False
        bounds_data, pixel_bounds = template
        px_width = int(abs(pixel_bounds[1] - pixel_bounds[0]) + 1)
        px_height = int(abs(pixel_bounds[3] - pixel_bounds[2]) + 1)
        self._fixed_crop_template = {
            "width": px_width,
            "height": px_height,
            "square": bool(square),
        }
        self._fixed_crop_template_bounds = bounds_data
        self._fixed_crop_template_pixel_bounds = pixel_bounds
        self._fixed_crop_template_view_key = self._outline_key(view)
        self._fixed_crop_template_manual_dims = {
            "type": "real",
            "width": real_width,
            "height": real_height,
            "unit": self._guess_view_unit(view),
        }
        self._fixed_crop_template_unit = self._guess_view_unit(view)
        self.draw_idle()
        return True

    def get_main_view_shape(self):
        """Return (height, width) of the current main view array, or (0,0) if unavailable."""
        ax = self.main_ax or next(iter(self._ax_view_map.keys()), None)
        view = self._ax_view_map.get(ax) if ax else None
        if not view:
            return (0, 0)
        arr_obj = view.get("arr")
        if arr_obj is None:
            return (0, 0)
        arr = np.asarray(arr_obj)
        if arr.ndim < 2:
            return (0, 0)
        return arr.shape[:2]

    def undo_fixed_crop_entry(self):
        if not self._fixed_crop_history:
            return None
        entry = self._fixed_crop_history[-1]
        return self.remove_fixed_crop_history_entry(entry.get("sequence"))

    def remove_fixed_crop_history_entry(self, seq):
        if seq is None:
            return None
        idx = None
        for i, entry in enumerate(self._fixed_crop_history):
            if entry.get("sequence") == seq:
                idx = i
                break
        if idx is None:
            return None
        entry = self._fixed_crop_history.pop(idx)
        if self._fixed_crop_history:
            self._update_template_from_entry(self._fixed_crop_history[-1])
        else:
            if not self._maybe_restore_manual_template():
                self._fixed_crop_template = None
                self._fixed_crop_template_bounds = None
                self._fixed_crop_template_pixel_bounds = None
                self._fixed_crop_template_view_key = None
                self._fixed_crop_history_highlight_seq = None
        self._emit_fixed_crop_history_update()
        self.draw_idle()
        return entry

    def clear_fixed_crop_history(self):
        if not self._fixed_crop_history:
            return
        self._fixed_crop_history.clear()
        self._fixed_crop_sequence = 1
        self._fixed_crop_history_highlight_seq = None
        self._emit_fixed_crop_history_update()
        if not self._maybe_restore_manual_template():
            self._fixed_crop_template = None
            self._fixed_crop_template_bounds = None
            self._fixed_crop_template_pixel_bounds = None
            self._fixed_crop_template_view_key = None
        self._fixed_crop_history_highlight_seq = None
        self._cleanup_highlight_artists()
        self.draw_idle()

    def _render_template_overlay(self, ax, view):
        if not self._fixed_crop_template_visible or ax is None:
            return
        if not self._fixed_crop_template_bounds or not self._fixed_crop_template:
            return
        key = self._outline_key(view)
        if key != self._fixed_crop_template_view_key:
            return
        x0, x1, y0, y1 = self._fixed_crop_template_bounds
        left, right = min(x0, x1), max(x0, x1)
        bottom, top = min(y0, y1), max(y0, y1)
        width = right - left
        height = top - bottom
        if width <= 0 or height <= 0:
            return
        rect = patches.Rectangle(
            (left, bottom),
            width,
            height,
            linewidth=1.2,
            edgecolor='#ff66ff',
            facecolor='none',
            alpha=0.85,
            linestyle='--',
            zorder=17,
        )
        ax.add_patch(rect)
        px_width = int(self._fixed_crop_template.get('width', int(width)))
        px_height = int(self._fixed_crop_template.get('height', int(height)))
        real_unit = self._fixed_crop_template_unit or "nm"
        real_label = ""
        if self._fixed_crop_template_bounds:
            bx0, bx1, by0, by1 = self._fixed_crop_template_bounds
            real_dx = abs(bx1 - bx0)
            real_dy = abs(by1 - by0)
            real_label = f"{real_dx:.3g} {real_unit} × {real_dy:.3g} {real_unit}"
        size_label = f"{real_label}\n({px_width}×{px_height} px)" if real_label else f"{px_width}×{px_height} px"
        ax.text(
            left + (width * 0.01),
            top - (height * 0.02),
            size_label,
            color='#ff66ff',
            fontsize=7,
            weight='medium',
            verticalalignment='top',
            horizontalalignment='left',
            bbox=dict(facecolor='#111111', alpha=0.7, pad=1, edgecolor='none'),
            zorder=18,
        )

    def _apply_fixed_crop_quick(self, event, view, ax):
        template = self._fixed_crop_template
        if template is None or ax is None:
            return False
        arr = np.asarray(view.get("arr"))
        if arr.size == 0:
            return False
        flip = bool(view.get("relative_axes"))
        arr_disp = np.flipud(arr) if flip else arr
        h, w = arr_disp.shape[:2]
        if w == 0 or h == 0:
            return False
        width = min(max(2, int(template.get("width", 2))), w)
        height = min(max(2, int(template.get("height", 2))), h)
        if width <= 0 or height <= 0:
            return False
        cx = self._axis_coord_to_pixel(ax, event.xdata, w, 'x')
        cy = self._axis_coord_to_pixel(ax, event.ydata, h, 'y')
        max_c0 = max(0, w - width)
        max_r0 = max(0, h - height)
        c0 = int(np.clip(cx - width // 2, 0, max_c0))
        r0 = int(np.clip(cy - height // 2, 0, max_r0))
        c1 = c0 + width - 1
        r1 = r0 + height - 1
        cropped_disp = arr_disp[r0:r1 + 1, c0:c1 + 1]
        if cropped_disp.size == 0:
            return False
        cropped_arr = np.flipud(cropped_disp) if flip else cropped_disp
        crop_extent = self._compute_crop_extent(view, w, h, c0, c1, r0, r1)
        bounds_data = self._pixel_bounds_to_axis_bounds(ax, w, h, c0, c1, r0, r1)
        entry = self._register_crop_entry(view, bounds_data, (c0, c1, r0, r1), template.get("square", False), update_size=False)
        new_view = dict(view)
        try:
            new_view["arr"] = np.array(cropped_arr, copy=True)
        except Exception:
            new_view["arr"] = cropped_arr
        if crop_extent is not None:
            new_view["extent_raw"] = crop_extent
            display_extent = self._display_extent_for_view(new_view, crop_extent)
            if display_extent is not None:
                new_view["extent"] = display_extent
            else:
                new_view.pop("extent", None)
        else:
            new_view.pop("extent", None)
            new_view.pop("extent_raw", None)
        new_view["title"] = f"{view.get('title') or 'crop'} [crop]"
        if entry and entry.get("sequence") is not None:
            new_view["crop_sequence"] = entry["sequence"]
            entry["view_snapshot"] = dict(new_view)
        if callable(self._crop_callback):
            try:
                self._crop_callback(new_view)
            except Exception:
                pass
        if self._fixed_crop_history_visible and entry:
            self._render_history_entry(ax, entry)
        if self._fixed_crop_template_visible:
            self._render_template_overlay(ax, view)
        self.draw_idle()
        return True
class SafeFigureCanvas(FigureCanvas):
    def draw(self):
        try:
            super().draw()
        except np.linalg.LinAlgError:
            # Ignore transient singular transforms during layout updates.
            return
