"""Detail canvases and spectroscopy dialogs."""
from __future__ import annotations

import copy
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
from matplotlib.transforms import Affine2D
import matplotlib
from matplotlib.collections import LineCollection
import matplotlib.patheffects as PathEffects

from ..._shared import QtCore, QtGui, QtWidgets
from ...config import load_config, save_config
from .molecular_overlay import (
    Molecule,
    MoleculePropertiesDialog,
    get_atom_color,
    get_atom_radius,
    available_atom_palettes,
    normalize_molecule_render_style,
)
from ..plot_typography import add_font_menu_action, normalize_font_family, apply_text_style
from ..palettes import DEFAULT_COLOR_CYCLE, get_color_cycle
from ..profile_links import register_profile_canvas, notify_profile_source_changed
from ..ppt_bridge import powerpoint_support_status, send_pixmap_to_ppt
from ..thumbnail_render import _interp_index, sample_array_value, array_to_qimage, _colormap_icon
from ..system_open import add_source_file_menu

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
_UNDO_HISTORY_LIMIT = 24
_MOLECULE_FILE_EXTS = {".xyz", ".pdb", ".mol"}
_DEFAULT_MOLECULE_STYLE = {
    "display_mode": "Bonds Only",
    "render_style": "licorice",
    "bond_style": "thin",
    "radius_mode": "vdw",
    "radius_scale": 1.0,
    "palette": "avogadro",
    "show_shadows": False,
    "show_hydrogens": False,
    "atom_color_override": None,
    "bond_color_override": None,
    "bond_color_mode": "default",
    "atom_color_map": {},
}

class MultiPreviewCanvas(FigureCanvas):
    _RECENT_MOLECULES = []
    _DRAG_VIEW_SNAPSHOTS = {}
    _DRAG_VIEW_SNAPSHOT_LIMIT = 32

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
            self.setAcceptDrops(True)
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
        self._render_suspended = False
        self._render_pending = False
        self.setFocusPolicy(QtCore.Qt.StrongFocus)
        self._overlay_shortcuts = []
        self.views = []
        self._ax_view_map = {}
        self._relative_axes_override = None
        self._suspend_zoom_restore = False
        self._image_meta = {}
        self._copy_feedback_handler = None
        self._views_callback = None
        self._drag_candidate = None  # (view, QPoint start, QImage cache)
        self._crop_callback = None  # callable(view_dict) -> None
        self._virtual_copy_callback = None
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
        self._fixed_crop_transform_mode = False
        self._fixed_crop_template_drag = None
        self._fixed_crop_history = []
        self._fixed_crop_sequence = 1
        self._fixed_crop_template_unit = "nm"
        self._fixed_crop_history_callback = None
        self._fixed_crop_template_manual_dims = None
        self._fixed_crop_history_highlight_seq = None
        self._fixed_crop_history_highlight_artists = {}
        self._fixed_crop_drag_last_ts = 0.0
        self._fixed_crop_drag_throttle_ms = 12.0
        self._fixed_crop_overlay_artists = {}
        self._fixed_crop_cursor_mode = None
        self._double_click_callback = None  # callable(view_dict) -> None
        self._filter_menu_callback = None  # callable(menu, view, canvas)
        self._histogram_dialog_callback = None
        self._histogram_auto_callback = None
        self._histogram_reset_callback = None
        self._display_relative_zero_menu_callback = None
        self._display_relative_zero_menu_state_callback = None
        self._display_relative_zero_menu_tooltip = ""
        self._apply_popup_style_callback = None
        self._apply_popup_style_label = "Apply this style to all pop-ups"
        self._apply_popup_style_tooltip = ""
        self._collection_menu_callback = None
        self._collection_help_callback = None
        self._compare_menu_callback = None
        self._compare_menu_state_callback = None
        self._stp_export_callback = None
        self._arrange_windows_callback = None
        self._minimize_windows_callback = None
        self._restore_windows_callback = None
        self._close_windows_callback = None
        self._value_callback = None
        self._undo_history = []
        self._undo_restore_in_progress = False
        self._undo_suspend_depth = 0
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
        self._profile_user_enabled = False
        self._profile_quick_transient = False
        self._profile_move_only = False
        self._profile_update_timer = QtCore.QTimer(self)
        self._profile_update_timer.setSingleShot(True)
        self._profile_update_timer.setInterval(50)
        self._profile_update_timer.timeout.connect(self._flush_profile_updates)
        self._saved_profiles = []
        self._profile_live_source_id = f"canvas-{id(self):x}"
        self._profile_saved_profile_seq = 1
        self._profile_palette_name = DEFAULT_COLOR_CYCLE
        self._profile_palette_colors = get_color_cycle(self._profile_palette_name)
        self._profile_color_cycle = itertools.cycle(self._profile_palette_colors)
        self._line_drag_origin = None
        self._saved_profile_drag = None
        self._active_profile_color = '#fbc02d'
        self._active_profile_lw = 2.0
        self._active_profile_line_style = "-"
        self._active_profile_marker_style = "o"
        self._active_profile_marker_size = 7.0
        self._active_profile_original_id = None
        self._highlighted_overlay = None
        self._cids = []
        self._base_click_cid = self.mpl_connect('button_press_event', self._on_base_click)
        self._dragging = None  # 'p0' or 'p1'
        self.main_ax = None
        self.profile_callback = None  # callable(active_dataset, saved_datasets)
        self._profile_highlight_cb = None
        self._profile_label_scale = 1.0
        self._profile_label_mode = "length"
        self._measurement_shortcuts_enabled = True
        self._view_font_scale = 1.0
        self._font_family = normalize_font_family(matplotlib.rcParams.get("font.family", [None])[0], "sans-serif")
        self._plot_font_bold = bool(getattr(parent, "_plot_font_bold", False))
        self._plot_font_italic = bool(getattr(parent, "_plot_font_italic", False))
        self._plot_font_underline = bool(getattr(parent, "_plot_font_underline", False))
        self._colorbar_orientation = 'vertical'
        self._show_ticks = True
        self._show_colorbar = True
        self._show_title = True
        self._show_acquisition_overlay = False
        self._show_profile_overlays = True
        self._show_angle_overlays = True
        self._show_shortcut_hint = True
        self._show_image_size_overlay = False
        self._shortcut_hint_artist = None
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
        default_style = self._load_molecule_default_style()
        self._show_hydrogens = bool(default_style.get("show_hydrogens", False))
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
            'font_family': self._font_family
        }
        self._font_change_callback = None
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
        self._active_molecule_idx = None
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
        self._show_molecule_gizmo = False
        self._molecule_gizmo_until = 0.0
        self._molecule_gizmo_axes = None
        self._molecule_gizmo_artists = {}
        self._molecule_gizmo_drag = None
        self._molecule_gizmo_timer = QtCore.QTimer(self)
        self._molecule_gizmo_timer.setSingleShot(True)
        self._molecule_gizmo_timer.timeout.connect(self._on_molecule_gizmo_timeout)
        self._show_molecule_shadow = bool(default_style.get("show_shadows", False))
        self.show_molecules = True
        self._profile_background = None
        self._active_profile_original_color = None
        self._profile_blit_active = False
        self._profile_animation_enabled = False
        # Molecule palette
        self.molecule_palette = str(default_style.get("palette", "avogadro") or "avogadro").lower()
        self._molecule_palette_cb = None
        self._zoom_reset_limits = {}
        # Pan/zoom state
        self._pan_active = False
        self._pan_ax = None
        self._pan_start = None
        self._install_overlay_shortcuts()
        self._pan_start_lim = None
        self._pan_last_ts = 0.0
        self._pan_throttle_ms = 16
        self._scroll_zoom_cid = self.mpl_connect('scroll_event', self._on_scroll_zoom)

    def set_show_title(self, show: bool):
        """Toggle rendering of title/date overlays in views."""
        show = bool(show)
        if show == self._show_title:
            return
        self.push_undo_state("show_title")
        self._show_title = show
        self._redraw()
        self._notify_views_callback()

    def set_show_acquisition_overlay(self, show: bool):
        """Toggle acquisition metadata HUD in the top-right image corner."""
        show = bool(show)
        if show == self._show_acquisition_overlay:
            return
        self.push_undo_state("acquisition_overlay")
        self._show_acquisition_overlay = show
        self._redraw()
        self._notify_views_callback()

    def set_show_molecules(self, show: bool):
        """Toggle rendering of molecular overlays in views."""
        show = bool(show)
        if show == self.show_molecules:
            return
        self.push_undo_state("show_molecules")
        self.show_molecules = show
        self._redraw()
        self._notify_views_callback()

    def set_show_molecule_gizmo(self, show: bool):
        """Toggle the molecule orientation gizmo."""
        show = bool(show)
        if show == bool(getattr(self, "_show_molecule_gizmo", False)):
            return
        self.push_undo_state("show_molecule_gizmo")
        self._show_molecule_gizmo = show
        if show:
            self._wake_molecule_gizmo(1800, redraw=False)
        else:
            self._molecule_gizmo_until = 0.0
            try:
                self._molecule_gizmo_timer.stop()
            except Exception:
                pass
        self._redraw()
        self._notify_views_callback()

    def set_profile_tool_enabled(self, enabled: bool):
        enabled = bool(enabled)
        if enabled == bool(getattr(self, "_profile_user_enabled", self.profile_enabled)):
            return
        self.push_undo_state("profile_tool")
        self._profile_user_enabled = enabled
        if enabled:
            self._profile_move_only = False
        self.enable_profile(enabled)
        if enabled:
            try:
                self._emit_profile()
            except Exception:
                pass

    def deactivate_profile_tool(self, *, clear_active: bool = True, clear_saved: bool = False):
        """Disable profile interaction while optionally preserving saved overlays."""
        self._profile_user_enabled = False
        self._profile_quick_transient = False
        self._profile_move_only = False
        if self.profile_enabled:
            try:
                self._disconnect_profile_events()
            except Exception:
                pass
        self.profile_enabled = False
        if clear_active:
            try:
                self._clear_profile_artists()
            except Exception:
                pass
            self.profile_pts = None
            self._active_profile_original_color = None
            self._profile_marker_positions = None
            self._profile_marker_domain = None
            try:
                self._clear_profile_hud()
            except Exception:
                pass
        if clear_saved:
            try:
                self._clear_saved_profile_artists(notify=False)
            except Exception:
                pass
        try:
            self._emit_profile_state()
        except Exception:
            pass
        self.draw_idle()

    def set_angle_tool_enabled(self, enabled: bool):
        self.enable_angle(bool(enabled))

    def set_measurement_shortcuts_enabled(self, enabled: bool):
        """Enable or disable Ctrl-based quick profile/angle shortcuts for this canvas."""
        self._measurement_shortcuts_enabled = bool(enabled)

    def clear_measurement_overlays(self):
        try:
            self.push_undo_state("clear_measurements")
            self._clear_angle_artists()
            self._clear_profile_artists()
            self._clear_saved_profile_artists(notify=False)
            self.profile_pts = None
            if self.angle_enabled:
                self._undo_suspend_depth += 1
                try:
                    self._ensure_angle_frames()
                finally:
                    self._undo_suspend_depth = max(0, self._undo_suspend_depth - 1)
                self._emit_angle()
            self._emit_profile_state()
            self.draw_idle()
        except Exception:
            pass

    def apply_display_preset(self, preset: str):
        preset = (preset or "").strip().lower()
        if preset not in {"focus", "analysis", "publication"}:
            return
        self.push_undo_state(f"preset:{preset}")
        if preset == "focus":
            self._show_ticks = False
            self._show_colorbar = False
            self._show_title = False
            self._show_profile_overlays = False
            self._show_angle_overlays = False
        elif preset == "analysis":
            self._show_ticks = True
            self._show_colorbar = True
            self._show_title = True
            self._show_profile_overlays = True
            self._show_angle_overlays = True
        else:
            self._show_ticks = False
            self._show_colorbar = True
            self._show_title = True
            self._show_profile_overlays = False
            self._show_angle_overlays = False
        self._apply_profile_visibility()
        self._apply_angle_visibility()
        self._redraw()
        self._notify_views_callback()

    def set_fit_to_canvas(self, enabled: bool):
        """If enabled, stretch view axes to fill the available canvas area."""
        enabled = bool(enabled)
        if enabled == self._fit_to_canvas:
            return
        self._fit_to_canvas = enabled
        self._redraw()

    def set_frame_fill_mode(self, enabled: bool):
        """Toggle a minimalist full-frame display mode for popups."""
        enabled = bool(enabled)
        if enabled == self._frame_fill_mode:
            return
        self.push_undo_state("frame_fill")
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
        self._notify_views_callback()

    def draw(self):
        if getattr(self, "_render_suspended", False):
            self._render_pending = True
            return
        try:
            super().draw()
            self._render_pending = False
        except np.linalg.LinAlgError:
            # Ignore transient singular transforms during layout updates.
            return

    def draw_idle(self):
        if getattr(self, "_render_suspended", False):
            self._render_pending = True
            return
        try:
            super().draw_idle()
        except np.linalg.LinAlgError:
            return

    def set_render_suspended(self, suspended: bool):
        suspended = bool(suspended)
        previously = bool(getattr(self, "_render_suspended", False))
        self._render_suspended = suspended
        if previously and not suspended and getattr(self, "_render_pending", False):
            self._render_pending = False
            try:
                super().draw()
            except np.linalg.LinAlgError:
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
        self.push_undo_state("view_layout")
        self._view_layout = layout
        self._redraw()
        self._notify_views_callback()

    def set_relative_axes_override(self, value):
        """Force all views to use relative axes (True/False) or None to defer to view settings."""
        if value is None:
            new_val = None
        else:
            new_val = bool(value)
        if self._relative_axes_override == new_val:
            return
        self.push_undo_state("relative_axes")
        self._relative_axes_override = new_val
        self.suspend_zoom_restore()
        self._redraw()
        self._notify_views_callback()

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

    def _notify_views_callback(self):
        if callable(self._views_callback):
            try:
                self._views_callback(self.views)
            except Exception:
                pass

    def set_crop_callback(self, cb):
        """Register a callback to receive cropped views created via drag-crop."""
        self._crop_callback = cb

    def set_virtual_copy_callback(self, cb):
        """Register a callback that can promote a view into a virtual thumbnail copy."""
        self._virtual_copy_callback = cb

    def set_double_click_callback(self, cb):
        """Register a callback to pop out the clicked view on double-click."""
        self._double_click_callback = cb

    def set_filter_menu_callback(self, cb):
        """Provide a callback to populate filter actions on the context menu."""
        self._filter_menu_callback = cb

    def set_histogram_dialog_callback(self, cb):
        """Register callback to open the histogram dialog for this canvas."""
        self._histogram_dialog_callback = cb

    def set_histogram_auto_callback(self, cb):
        """Register callback that applies automatic contrast to the active view."""
        self._histogram_auto_callback = cb

    def set_histogram_reset_callback(self, cb):
        """Register callback that resets contrast to the full data range."""
        self._histogram_reset_callback = cb

    def set_display_relative_zero_menu_callback(self, cb, state_cb=None, tooltip=None):
        """Register an optional popup-local menu action for relative-zero display."""
        self._display_relative_zero_menu_callback = cb
        self._display_relative_zero_menu_state_callback = state_cb
        self._display_relative_zero_menu_tooltip = str(tooltip or "")

    def set_apply_popup_style_callback(self, cb, label=None, tooltip=None):
        """Register an optional action that applies this popup style to peer popups."""
        self._apply_popup_style_callback = cb
        if label:
            self._apply_popup_style_label = str(label)
        self._apply_popup_style_tooltip = str(tooltip or "")

    def set_collection_menu_callback(self, cb, help_cb=None):
        """Register collection-menu handlers used to save curated views into cross-folder collections."""
        self._collection_menu_callback = cb
        self._collection_help_callback = help_cb

    def set_compare_menu_callback(self, cb, state_cb=None):
        """Register compare-menu handlers used to populate A/B view comparisons from right-clicks."""
        self._compare_menu_callback = cb
        self._compare_menu_state_callback = state_cb

    def set_stp_export_callback(self, cb):
        """Register callback for WSxM STP export requests."""
        self._stp_export_callback = cb

    def set_window_arrange_callback(self, cb):
        """Register callback invoked when the user requests window tiling."""
        self._arrange_windows_callback = cb

    def set_window_minimize_callback(self, cb):
        """Register callback invoked when the user requests minimizing pop-out windows."""
        self._minimize_windows_callback = cb

    def set_window_restore_callback(self, cb):
        """Register callback invoked when the user requests recalling pop-out windows."""
        self._restore_windows_callback = cb

    def set_window_close_callback(self, cb):
        """Register callback invoked when the user requests closing all pop-out windows."""
        self._close_windows_callback = cb

    def set_plot_font_family_callback(self, cb):
        """Register a callback used when the user picks a new plot font."""
        self._font_change_callback = cb

    def set_plot_font_family(self, family: str):
        """Apply a shared font family to the canvas and its scale bar."""
        self.set_plot_typography(family=family)

    def _plot_style_state(self):
        return {
            "bold": bool(getattr(self, "_plot_font_bold", False)),
            "italic": bool(getattr(self, "_plot_font_italic", False)),
            "underline": bool(getattr(self, "_plot_font_underline", False)),
        }

    def set_plot_typography(self, *, family=None, bold=None, italic=None, underline=None):
        """Apply shared typography settings to the preview canvas."""
        changes = {
            "bold": bold,
            "italic": italic,
            "underline": underline,
        }
        if family is not None:
            family = normalize_font_family(family, "sans-serif")
            self._font_family = family
            self._scale_bar_settings["font_family"] = family
        for key, attr in (("bold", "_plot_font_bold"), ("italic", "_plot_font_italic"), ("underline", "_plot_font_underline")):
            value = changes.get(key)
            if value is not None:
                setattr(self, attr, bool(value))
        self._redraw()

    def _apply_plot_font_family_choice(self, family: str):
        """Route a font choice through the owner first, then fall back locally."""
        family = normalize_font_family(family, "sans-serif")
        if callable(self._font_change_callback):
            try:
                self._font_change_callback(family)
                return
            except Exception:
                pass
        self.set_plot_font_family(family)

    def set_show_profile_overlays(self, show: bool):
        show = bool(show)
        if show == self._show_profile_overlays:
            return
        self.push_undo_state("show_profile_overlays")
        self._show_profile_overlays = show
        self._apply_profile_visibility()
        self.draw_idle()
        self._notify_views_callback()

    def set_show_angle_overlays(self, show: bool):
        show = bool(show)
        if show == self._show_angle_overlays:
            return
        self.push_undo_state("show_angle_overlays")
        self._show_angle_overlays = show
        self._apply_angle_visibility()
        self.draw_idle()
        self._notify_views_callback()

    def set_show_shortcut_hint(self, show: bool):
        show = bool(show)
        if show == self._show_shortcut_hint:
            return
        self.push_undo_state("show_shortcut_hint")
        self._show_shortcut_hint = show
        if not self._show_shortcut_hint:
            self._clear_shortcut_hint_artist()
        self._redraw()
        self._notify_views_callback()

    def _install_overlay_shortcuts(self):
        """Bind overlay toggles at the window level so they work from toolbars too."""
        if self._overlay_shortcuts:
            return
        try:
            shortcuts = [
                ("Ctrl+1", lambda: self.set_show_profile_overlays(not self._show_profile_overlays)),
                ("Ctrl+2", lambda: self.set_show_angle_overlays(not self._show_angle_overlays)),
                ("Ctrl+3", lambda: self.set_show_molecules(not self.show_molecules)),
                ("Ctrl+4", lambda: self.enable_scale_bar(not self.scale_bar_enabled)),
                ("Ctrl+5", lambda: self.set_show_acquisition_overlay(not self._show_acquisition_overlay)),
                ("Ctrl+H", lambda: self.set_show_shortcut_hint(not self._show_shortcut_hint)),
                ("Ctrl+E", lambda: self.enable_fixed_crop_transform_mode(not self._fixed_crop_transform_mode)),
                (QtCore.Qt.Key_Return, self._on_apply_fixed_crop_shortcut),
                (QtCore.Qt.Key_Enter, self._on_apply_fixed_crop_shortcut),
                (QtCore.Qt.Key_Escape, self._on_cancel_fixed_crop_shortcut),
            ]
            for seq, handler in shortcuts:
                shortcut = QtWidgets.QShortcut(QtGui.QKeySequence(seq), self)
                shortcut.setContext(QtCore.Qt.WidgetWithChildrenShortcut)
                shortcut.activated.connect(handler)
                self._overlay_shortcuts.append(shortcut)
        except Exception:
            self._overlay_shortcuts = []

    def _on_apply_fixed_crop_shortcut(self):
        if not self._fixed_crop_template_visible:
            return
        view, ax = self._fixed_crop_target_view()
        if view is None or ax is None:
            return
        self._apply_fixed_crop_template(view, ax)

    def _on_cancel_fixed_crop_shortcut(self):
        if not self._fixed_crop_transform_mode:
            return
        self.enable_fixed_crop_transform_mode(False)

    def _clear_shortcut_hint_artist(self):
        art = getattr(self, "_shortcut_hint_artist", None)
        if art is None:
            return
        try:
            art.remove()
        except Exception:
            pass
        self._shortcut_hint_artist = None

    def _shortcut_hint_hit(self, event):
        art = getattr(self, "_shortcut_hint_artist", None)
        if art is None or event is None:
            return False
        try:
            return bool(art.contains(event)[0])
        except Exception:
            return False

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

    def _clone_undo_value(self, value):
        if isinstance(value, np.ndarray):
            return np.array(value, copy=True)
        try:
            return copy.deepcopy(value)
        except Exception:
            return value

    def _clone_undo_view(self, view):
        if not isinstance(view, dict):
            return self._clone_undo_value(view)
        cloned = {}
        for key, value in view.items():
            cloned[key] = self._clone_undo_value(value)
        return cloned

    def export_outline_state(self):
        groups = []
        for key, entries in (self._outlines or {}).items():
            try:
                file_path, extent = key
            except Exception:
                file_path, extent = None, ()
            exported_entries = []
            for entry in entries or []:
                if isinstance(entry, dict):
                    pts = entry.get("pts")
                    color = entry.get("color", self._outline_default_color)
                    lw = entry.get("lw", self._outline_default_lw)
                    ls = entry.get("ls", self._outline_default_ls)
                else:
                    pts = entry
                    color = self._outline_default_color
                    lw = self._outline_default_lw
                    ls = self._outline_default_ls
                try:
                    pts_arr = np.asarray(pts, dtype=float)
                except Exception:
                    continue
                if pts_arr.ndim != 2 or pts_arr.shape[0] < 2:
                    continue
                exported_entries.append(
                    {
                        "pts": pts_arr.tolist(),
                        "color": color,
                        "lw": float(lw),
                        "ls": self._clone_undo_value(ls),
                    }
                )
            if exported_entries:
                groups.append(
                    {
                        "file_path": file_path,
                        "extent": list(extent or ()),
                        "entries": exported_entries,
                    }
                )
        order = []
        for key, idx in self._outline_order or []:
            try:
                file_path, extent = key
                idx_int = int(idx)
            except Exception:
                continue
            order.append(
                {
                    "file_path": file_path,
                    "extent": list(extent or ()),
                    "index": idx_int,
                }
            )
        return {
            "groups": groups,
            "order": order,
            "threshold": float(self._outline_threshold),
        }

    def import_outline_state(self, state):
        self._outlines = {}
        self._outline_order = []
        if not isinstance(state, dict):
            return
        try:
            self._outline_threshold = max(0.05, min(0.99, float(state.get("threshold", self._outline_threshold))))
        except Exception:
            pass
        for group in state.get("groups", []) or []:
            key = (
                group.get("file_path"),
                tuple(group.get("extent") or ()),
            )
            entries = []
            for entry in group.get("entries", []) or []:
                try:
                    pts_arr = np.asarray(entry.get("pts"), dtype=float)
                except Exception:
                    continue
                if pts_arr.ndim != 2 or pts_arr.shape[0] < 2:
                    continue
                entries.append(
                    {
                        "pts": pts_arr,
                        "color": entry.get("color", self._outline_default_color),
                        "lw": float(entry.get("lw", self._outline_default_lw)),
                        "ls": self._clone_undo_value(entry.get("ls", self._outline_default_ls)),
                    }
                )
            if entries:
                self._outlines[key] = entries
        order = []
        for item in state.get("order", []) or []:
            key = (
                item.get("file_path"),
                tuple(item.get("extent") or ()),
            )
            try:
                idx = int(item.get("index", -1))
            except Exception:
                continue
            outlines = self._outlines.get(key)
            if outlines and 0 <= idx < len(outlines):
                order.append((key, idx))
        if not order:
            for key, entries in self._outlines.items():
                order.extend((key, idx) for idx in range(len(entries)))
        self._outline_order = order

    def export_canvas_undo_state(self):
        return {
            "views": [self._clone_undo_view(v) for v in (self.views or [])],
            "profile_state": self._clone_undo_value(self.export_profile_state()),
            "profile_user_enabled": bool(getattr(self, "_profile_user_enabled", self.profile_enabled)),
            "active_profile_color": self._active_profile_color,
            "active_profile_lw": float(self._active_profile_lw),
            "active_profile_line_style": self._active_profile_line_style,
            "active_profile_marker_style": self._active_profile_marker_style,
            "active_profile_marker_size": float(self._active_profile_marker_size),
            "active_profile_original_color": self._active_profile_original_color,
            "profile_label_mode": self._profile_label_mode,
            "angle_state": self._clone_undo_value(self.export_angle_state()),
            "angle_enabled": bool(self.angle_enabled),
            "molecule_state": self._clone_undo_value(self.export_molecule_state()),
            "outline_state": self._clone_undo_value(self.export_outline_state()),
            "show_title": bool(self._show_title),
            "show_acquisition_overlay": bool(self._show_acquisition_overlay),
            "show_molecules": bool(self.show_molecules),
            "show_profile_overlays": bool(self._show_profile_overlays),
            "show_angle_overlays": bool(self._show_angle_overlays),
            "show_shortcut_hint": bool(self._show_shortcut_hint),
            "show_molecule_gizmo": bool(self._show_molecule_gizmo),
            "show_ticks": bool(self._show_ticks),
            "show_colorbar": bool(self._show_colorbar),
            "scale_bar_enabled": bool(self.scale_bar_enabled),
            "colorbar_orientation": self._colorbar_orientation,
            "view_layout": self._view_layout,
            "frame_fill_mode": bool(self._frame_fill_mode),
            "frame_fill_prev_state": self._clone_undo_value(self._frame_fill_prev_state),
            "fit_to_canvas": bool(self._fit_to_canvas),
            "relative_axes_override": self._clone_undo_value(self._relative_axes_override),
        }

    def push_undo_state(self, label=None):
        if self._undo_restore_in_progress or self._undo_suspend_depth > 0:
            return False
        try:
            state = self.export_canvas_undo_state()
        except Exception:
            return False
        if label is not None:
            state["_label"] = str(label)
        self._undo_history.append(state)
        if len(self._undo_history) > _UNDO_HISTORY_LIMIT:
            self._undo_history.pop(0)
        return True

    def handle_undo_request(self):
        if self.undo_last_action():
            return True
        if self._undo_last_profile_snapshot():
            return True
        if self._undo_last_outline():
            return True
        return bool(self.undo_last_molecule_change())

    def undo_last_action(self):
        if not self._undo_history:
            return False
        state = self._undo_history.pop()
        if not isinstance(state, dict):
            return False
        self._undo_restore_in_progress = True
        self._undo_suspend_depth += 1
        try:
            try:
                self.enable_profile(False)
            except Exception:
                pass
            try:
                self.enable_angle(False)
            except Exception:
                pass

            self._show_title = bool(state.get("show_title", self._show_title))
            self._show_acquisition_overlay = bool(state.get("show_acquisition_overlay", self._show_acquisition_overlay))
            self.show_molecules = bool(state.get("show_molecules", self.show_molecules))
            self._show_profile_overlays = bool(state.get("show_profile_overlays", self._show_profile_overlays))
            self._show_angle_overlays = bool(state.get("show_angle_overlays", self._show_angle_overlays))
            self._show_shortcut_hint = bool(state.get("show_shortcut_hint", self._show_shortcut_hint))
            self._show_molecule_gizmo = bool(state.get("show_molecule_gizmo", self._show_molecule_gizmo))
            if not self._show_shortcut_hint:
                self._clear_shortcut_hint_artist()
            self._show_ticks = bool(state.get("show_ticks", self._show_ticks))
            self._show_colorbar = bool(state.get("show_colorbar", self._show_colorbar))
            self._colorbar_orientation = str(state.get("colorbar_orientation", self._colorbar_orientation) or "vertical")
            self._view_layout = str(state.get("view_layout", self._view_layout) or "grid")
            self._frame_fill_mode = bool(state.get("frame_fill_mode", self._frame_fill_mode))
            self._frame_fill_prev_state = self._clone_undo_value(state.get("frame_fill_prev_state"))
            self._fit_to_canvas = bool(state.get("fit_to_canvas", self._fit_to_canvas))
            self._relative_axes_override = self._clone_undo_value(state.get("relative_axes_override", self._relative_axes_override))

            desired_scale_bar = bool(state.get("scale_bar_enabled", self.scale_bar_enabled))
            if desired_scale_bar != self.scale_bar_enabled:
                self.scale_bar_enabled = desired_scale_bar
                if desired_scale_bar:
                    self._connect_scale_bar_events()
                else:
                    self._disconnect_scale_bar_events()

            self.molecules = []
            for entry in state.get("molecule_state", []) or []:
                try:
                    self.molecules.append(Molecule.from_dict(entry))
                except Exception:
                    continue
            self.import_outline_state(state.get("outline_state"))

            views = [self._clone_undo_view(v) for v in (state.get("views") or [])]
            self.set_views(views, preserve_profiles=False)

            self._active_profile_color = state.get("active_profile_color", self._active_profile_color) or self._active_profile_color
            try:
                self._active_profile_lw = float(state.get("active_profile_lw", self._active_profile_lw))
            except Exception:
                pass
            self._active_profile_line_style = self._normalize_profile_line_style(
                state.get("active_profile_line_style", self._active_profile_line_style),
                self._active_profile_line_style,
            )
            self._active_profile_marker_style = self._normalize_profile_marker_style(
                state.get("active_profile_marker_style", self._active_profile_marker_style),
                self._active_profile_marker_style,
            )
            try:
                self._active_profile_marker_size = float(
                    state.get("active_profile_marker_size", self._active_profile_marker_size)
                )
            except Exception:
                pass
            self._active_profile_original_color = state.get("active_profile_original_color")
            mode = str(state.get("profile_label_mode", self._profile_label_mode) or "length").strip().lower()
            if mode not in ("length", "full", "hidden"):
                mode = "length"
            self._profile_label_mode = mode
            self._profile_user_enabled = bool(state.get("profile_user_enabled", getattr(self, "_profile_user_enabled", False)))
            profile_state = state.get("profile_state")
            if isinstance(profile_state, dict):
                self.import_profile_state(self._clone_undo_value(profile_state), emit=False)

            angle_state = state.get("angle_state")
            if bool(state.get("angle_enabled", False)):
                self.angle_enabled = True
                self._connect_angle_events()
                if isinstance(angle_state, dict):
                    self.import_angle_state(self._clone_undo_value(angle_state))
                else:
                    self._ensure_angle_frames()
                self._apply_angle_visibility()
            else:
                self.angle_enabled = False

            self._apply_profile_visibility()
            self.draw_idle()
            self._notify_views_callback()
            return True
        finally:
            self._undo_suspend_depth = max(0, self._undo_suspend_depth - 1)
            self._undo_restore_in_progress = False

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
        current_base_limits = {}
        preserve_zoom = not getattr(self, "_suspend_zoom_restore", False)
        profile_state = None
        try:
            profile_state = self.export_profile_state()
        except Exception:
            profile_state = None
        if preserve_zoom:
            try:
                for ax, v in list(self._ax_view_map.items()):
                    try:
                        key = self._outline_key(v)
                        current_limits[key] = (ax.get_xlim(), ax.get_ylim())
                        current_base_limits[key] = self._zoom_reset_limits.get(ax, (ax.get_xlim(), ax.get_ylim()))
                    except Exception:
                        continue
            except Exception:
                current_limits = {}
                current_base_limits = {}
        else:
            self._suspend_zoom_restore = False

        self.fig.clf()
        self._ax_view_map = {}
        self._image_meta = {}
        self._fixed_crop_overlay_artists = {}
        self._molecule_gizmo_axes = None
        self._molecule_gizmo_artists = {}
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
                try:
                    apply_text_style(cbar.ax.xaxis.label, family=self._font_family, **self._plot_style_state())
                    apply_text_style(cbar.ax.yaxis.label, family=self._font_family, **self._plot_style_state())
                    for lbl in list(cbar.ax.get_xticklabels()) + list(cbar.ax.get_yticklabels()):
                        apply_text_style(lbl, family=self._font_family, **self._plot_style_state())
                except Exception:
                    pass
                self._colorbars.append(cbar)
            title = v.get('title', '')
            if title and self._show_title:
                ax.set_title(title, fontsize=9)
                apply_text_style(ax.title, family=self._font_family, **self._plot_style_state())
            else:
                ax.set_title("")
            ax.tick_params(labelsize=8)
            for lbl in list(ax.get_xticklabels()) + list(ax.get_yticklabels()):
                apply_text_style(lbl, family=self._font_family, **self._plot_style_state())
            self._draw_acquisition_overlay(ax, v)
            if ax is self.main_ax:
                self._draw_shortcut_hint(ax)
            if not self._show_ticks:
                ax.set_xticks([])
                ax.set_yticks([])
            # Restore previous zoom if available
            if preserve_zoom:
                key = self._outline_key(v)
                prev_lim = current_limits.get(key)
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
            self._draw_image_size_overlay(ax, v)
            if ax not in self._zoom_reset_limits:
                key = self._outline_key(v)
                base_lim = current_base_limits.get(key)
                self._zoom_reset_limits[ax] = base_lim if base_lim else (ax.get_xlim(), ax.get_ylim())
            # Draw molecules on every view
            self._draw_molecules(ax)
            if ax is self.main_ax:
                self._draw_molecule_gizmo(ax)
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
        profile_restored = False
        if isinstance(profile_state, dict):
            try:
                if (
                    profile_state.get("active_pts") is not None
                    or profile_state.get("saved")
                    or profile_state.get("enabled")
                    or profile_state.get("user_enabled")
                ):
                    self.import_profile_state(profile_state, emit=False)
                    profile_restored = True
            except Exception:
                profile_restored = False
        if not profile_restored and self.profile_enabled:
            self._ensure_profile_artists()
            self._emit_profile()
        if self.angle_enabled:
            self._ensure_angle_frames()
            self._apply_angle_visibility()
        self._apply_profile_visibility()
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
            z_range = z_vals.max() - z_vals.min()
            if z_range < 1e-6:
                z_range = 1.0

            lc = None
            lc_underlay = None
            sc = None
            shadow_sc = None
            atom_underlay_sc = None
            atom_style = normalize_molecule_render_style(mol.render_style)
            bond_style = (mol.bond_style or "default").lower()
            style_profiles = {
                "shaded": {
                    "size_base": 80, "size_scale": 140, "shadow_alpha": 0.25,
                    "show_shadow": True, "atom_alpha": 1.0, "highlight": 0.25,
                    "edgecolor": "black", "edgewidth": 0.6, "bond_scale": 1.0,
                    "bond_alpha_min": 0.4, "bond_alpha_span": 0.6, "split_bonds": False,
                    "force_bond_by_atoms": False, "outline": False,
                },
                "flat": {
                    "size_base": 76, "size_scale": 126, "shadow_alpha": 0.0,
                    "show_shadow": False, "atom_alpha": 0.98, "highlight": 0.0,
                    "edgecolor": "black", "edgewidth": 0.5, "bond_scale": 1.0,
                    "bond_alpha_min": 0.55, "bond_alpha_span": 0.4, "split_bonds": False,
                    "force_bond_by_atoms": False, "outline": False,
                },
                "ballstick": {
                    "size_base": 34, "size_scale": 62, "shadow_alpha": 0.14,
                    "show_shadow": True, "atom_alpha": 1.0, "highlight": 0.18,
                    "edgecolor": "black", "edgewidth": 0.75, "bond_scale": 1.35,
                    "bond_alpha_min": 0.55, "bond_alpha_span": 0.4, "split_bonds": True,
                    "force_bond_by_atoms": True, "outline": False,
                },
                "cpk": {
                    "size_base": 150, "size_scale": 240, "shadow_alpha": 0.28,
                    "show_shadow": True, "atom_alpha": 0.96, "highlight": 0.3,
                    "edgecolor": "black", "edgewidth": 0.75, "bond_scale": 0.75,
                    "bond_alpha_min": 0.3, "bond_alpha_span": 0.35, "split_bonds": False,
                    "force_bond_by_atoms": False, "outline": False,
                },
                "licorice": {
                    "size_base": 48, "size_scale": 82, "shadow_alpha": 0.2,
                    "show_shadow": True, "atom_alpha": 0.98, "highlight": 0.12,
                    "edgecolor": "black", "edgewidth": 0.55, "bond_scale": 1.5,
                    "bond_alpha_min": 0.5, "bond_alpha_span": 0.45, "split_bonds": True,
                    "force_bond_by_atoms": True, "outline": False,
                },
                "wire": {
                    "size_base": 14, "size_scale": 24, "shadow_alpha": 0.0,
                    "show_shadow": False, "atom_alpha": 0.82, "highlight": 0.0,
                    "edgecolor": "black", "edgewidth": 0.3, "bond_scale": 0.65,
                    "bond_alpha_min": 0.45, "bond_alpha_span": 0.35, "split_bonds": False,
                    "force_bond_by_atoms": False, "outline": False,
                },
                "line": {
                    "size_base": 8, "size_scale": 14, "shadow_alpha": 0.0,
                    "show_shadow": False, "atom_alpha": 0.8, "highlight": 0.0,
                    "edgecolor": "none", "edgewidth": 0.0, "bond_scale": 0.5,
                    "bond_alpha_min": 0.55, "bond_alpha_span": 0.25, "split_bonds": True,
                    "force_bond_by_atoms": True, "outline": False,
                },
                "sticks": {
                    "size_base": 18, "size_scale": 35, "shadow_alpha": 0.12,
                    "show_shadow": True, "atom_alpha": 0.96, "highlight": 0.08,
                    "edgecolor": "black", "edgewidth": 0.35, "bond_scale": 1.6,
                    "bond_alpha_min": 0.55, "bond_alpha_span": 0.4, "split_bonds": True,
                    "force_bond_by_atoms": True, "outline": False,
                },
                "skeletal": {
                    "size_base": 10, "size_scale": 20, "shadow_alpha": 0.0,
                    "show_shadow": False, "atom_alpha": 0.88, "highlight": 0.0,
                    "edgecolor": "black", "edgewidth": 0.25, "bond_scale": 0.85,
                    "bond_alpha_min": 0.5, "bond_alpha_span": 0.35, "split_bonds": False,
                    "force_bond_by_atoms": False, "outline": False,
                },
                "outline": {
                    "size_base": 78, "size_scale": 128, "shadow_alpha": 0.0,
                    "show_shadow": False, "atom_alpha": 0.12, "highlight": 0.0,
                    "edgecolor": "white", "edgewidth": 1.0, "bond_scale": 1.1,
                    "bond_alpha_min": 0.85, "bond_alpha_span": 0.05, "split_bonds": False,
                    "force_bond_by_atoms": False, "outline": True,
                },
                "ghost": {
                    "size_base": 82, "size_scale": 132, "shadow_alpha": 0.0,
                    "show_shadow": False, "atom_alpha": 0.28, "highlight": 0.05,
                    "edgecolor": "black", "edgewidth": 0.35, "bond_scale": 0.9,
                    "bond_alpha_min": 0.22, "bond_alpha_span": 0.18, "split_bonds": False,
                    "force_bond_by_atoms": False, "outline": False,
                },
            }
            profile = style_profiles.get(atom_style, style_profiles["shaded"])
            size_base = profile["size_base"]
            size_scale = profile["size_scale"]
            shadow_alpha = profile["shadow_alpha"]

            def _atom_base_rgba(idx):
                elem = mol.elements[idx] if idx < len(mol.elements) else ""
                cmap = getattr(mol, "atom_color_map", {}) or {}
                override = cmap.get(str(elem).upper()) or cmap.get(str(elem).title())
                if override:
                    return matplotlib.colors.to_rgba(override)
                if mol.atom_color_override:
                    return matplotlib.colors.to_rgba(mol.atom_color_override)
                return matplotlib.colors.to_rgba(get_atom_color(elem, self.molecule_palette))

            # Draw Bonds
            if 'Bonds' in mol.display_mode and len(mol.bonds) > 0:
                lines = []
                underlay_lines = []
                colors = []
                underlay_colors = []
                linewidths = []
                underlay_widths = []
                lw_scale = profile["bond_scale"]
                if bond_style == "thick":
                    lw_scale *= 1.6
                elif bond_style == "thin":
                    lw_scale *= 0.7
                display_mode_lower = (mol.display_mode or "").lower()
                force_atom_bond_colors = profile["force_bond_by_atoms"] or ("bonds only" in display_mode_lower)
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
                    z_mid = (p1[2] + p2[2]) * 0.5
                    z_norm = (z_mid - z_min) / z_range
                    alpha = profile["bond_alpha_min"] + profile["bond_alpha_span"] * z_norm
                    lw = (1.0 + 2.0 * z_norm) * lw_scale
                    bond_mode = getattr(mol, "bond_color_mode", None) or self._bond_color_mode
                    if force_atom_bond_colors and bond_mode == "default":
                        bond_mode = "by_atoms"
                    rgba1 = _atom_base_rgba(i)
                    rgba2 = _atom_base_rgba(j)
                    if bond_mode == "single" and mol.bond_color_override:
                        br, bg, bb, _ = matplotlib.colors.to_rgba(mol.bond_color_override)
                        segment_color = (br, bg, bb, alpha)
                        lines.append([(p1[0], p1[1]), (p2[0], p2[1])])
                        colors.append(segment_color)
                        linewidths.append(lw)
                        if profile["outline"]:
                            underlay_lines.append([(p1[0], p1[1]), (p2[0], p2[1])])
                            underlay_colors.append((0.0, 0.0, 0.0, 0.95))
                            underlay_widths.append(lw + 2.4)
                    elif bond_mode == "by_atoms":
                        r1, g1, b1, _ = rgba1
                        r2, g2, b2, _ = rgba2
                        if profile["split_bonds"]:
                            mid = 0.5 * (p1 + p2)
                            lines.append([(p1[0], p1[1]), (mid[0], mid[1])])
                            colors.append((r1, g1, b1, alpha))
                            linewidths.append(lw)
                            lines.append([(mid[0], mid[1]), (p2[0], p2[1])])
                            colors.append((r2, g2, b2, alpha))
                            linewidths.append(lw)
                            if profile["outline"]:
                                underlay_lines.extend([
                                    [(p1[0], p1[1]), (mid[0], mid[1])],
                                    [(mid[0], mid[1]), (p2[0], p2[1])],
                                ])
                                underlay_colors.extend([(0.0, 0.0, 0.0, 0.95)] * 2)
                                underlay_widths.extend([lw + 2.4, lw + 2.4])
                        else:
                            br, bg, bb = (0.5 * (r1 + r2), 0.5 * (g1 + g2), 0.5 * (b1 + b2))
                            lines.append([(p1[0], p1[1]), (p2[0], p2[1])])
                            colors.append((br, bg, bb, alpha))
                            linewidths.append(lw)
                            if profile["outline"]:
                                underlay_lines.append([(p1[0], p1[1]), (p2[0], p2[1])])
                                underlay_colors.append((0.0, 0.0, 0.0, 0.95))
                                underlay_widths.append(lw + 2.4)
                    else:
                        br, bg, bb = self._default_bond_color
                        lines.append([(p1[0], p1[1]), (p2[0], p2[1])])
                        colors.append((br, bg, bb, alpha))
                        linewidths.append(lw)
                        if profile["outline"]:
                            underlay_lines.append([(p1[0], p1[1]), (p2[0], p2[1])])
                            underlay_colors.append((0.0, 0.0, 0.0, 0.95))
                            underlay_widths.append(lw + 2.4)

                if underlay_lines:
                    lc_underlay = LineCollection(underlay_lines, colors=underlay_colors, linewidths=underlay_widths, zorder=28.7)
                    ax.add_collection(lc_underlay)
                    try:
                        lc_underlay.set_capstyle("round")
                    except Exception:
                        pass
                lc = LineCollection(lines, colors=colors, linewidths=linewidths, zorder=29)
                ax.add_collection(lc)
                lc.set_pickradius(5)
                try:
                    if atom_style in {"ballstick", "licorice", "sticks", "line"}:
                        lc.set_capstyle("round")
                except Exception:
                    pass

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
                if atom_style == "cpk" and str(rad_mode or "").lower() == "covalent":
                    rad_mode = "vdw"
                rad_scale = getattr(mol, "radius_scale", 1.0)
                rad_ref = max(get_atom_radius('C', 'covalent'), 1e-3)
                radius_factors = []
                for e in elements_sorted:
                    r_el = get_atom_radius(e, rad_mode)
                    radius_factors.append(max(r_el / rad_ref, 0.05) * rad_scale)
                sizes = (size_base + size_scale * z_norm) * np.array(radius_factors)
                
                rgba_colors = [_atom_base_rgba(i) for i in order]
                final_colors = []
                for i, (r, g, b, a) in enumerate(rgba_colors):
                    depth_alpha = profile["atom_alpha"] * (0.45 + 0.55 * z_norm[i])
                    highlight = profile["highlight"]
                    r_h = min(1.0, r + (1 - r) * highlight)
                    g_h = min(1.0, g + (1 - g) * highlight)
                    b_h = min(1.0, b + (1 - b) * highlight)
                    final_colors.append((r_h, g_h, b_h, depth_alpha))
                
                if profile["outline"]:
                    atom_underlay_sc = ax.scatter(
                        x, y,
                        s=sizes * 1.55,
                        c=[(0.0, 0.0, 0.0, 0.92)] * len(x),
                        edgecolors='none',
                        linewidths=0,
                        zorder=28.85,
                    )
                if self._show_molecule_shadow and profile["show_shadow"]:
                    shadow_sc = ax.scatter(
                        x + 0.05, y - 0.05,
                        s=sizes * 1.25,
                        c=[(0, 0, 0, shadow_alpha)] * len(x),
                        edgecolors='none',
                        linewidths=0,
                        zorder=28,
                    )
                sc = ax.scatter(
                    x,
                    y,
                    s=sizes,
                    c=final_colors,
                    edgecolors=profile["edgecolor"],
                    linewidths=profile["edgewidth"],
                    zorder=30,
                )
                if atom_style == "outline":
                    try:
                        sc.set_path_effects([PathEffects.Stroke(linewidth=2.2, foreground='black'), PathEffects.Normal()])
                    except Exception:
                        pass

            self._molecule_artists.append({
                'mol': mol,
                'ax': ax,
                'scatter': sc,
                'atom_underlay': atom_underlay_sc,
                'shadow': shadow_sc,
                'line_underlay': lc_underlay,
                'lines': lc
            })

    def _active_molecule_for_gizmo(self):
        if not self.molecules:
            return None
        idx = getattr(self, "_active_molecule_idx", None)
        if isinstance(idx, int) and 0 <= idx < len(self.molecules):
            return self.molecules[idx]
        return self.molecules[0] if self.molecules else None

    def _active_molecule_index_for_gizmo(self):
        if not self.molecules:
            return None
        idx = getattr(self, "_active_molecule_idx", None)
        if isinstance(idx, int) and 0 <= idx < len(self.molecules):
            return idx
        return 0

    def _should_show_molecule_gizmo(self):
        if not self.show_molecules or not self.molecules:
            return False
        if bool(getattr(self, "_show_molecule_gizmo", False)):
            return True
        return (time.perf_counter() * 1000.0) < float(getattr(self, "_molecule_gizmo_until", 0.0) or 0.0)

    def _molecule_orientation_matrix(self, mol):
        if mol is None:
            return np.eye(3, dtype=float)
        rads = np.radians(np.asarray(getattr(mol, "angles", [0.0, 0.0, 0.0]), dtype=float))
        cx, cy, cz = np.cos(rads)
        sx, sy, sz = np.sin(rads)
        rx = np.array([[1.0, 0.0, 0.0], [0.0, cx, -sx], [0.0, sx, cx]], dtype=float)
        ry = np.array([[cy, 0.0, sy], [0.0, 1.0, 0.0], [-sy, 0.0, cy]], dtype=float)
        rz = np.array([[cz, -sz, 0.0], [sz, cz, 0.0], [0.0, 0.0, 1.0]], dtype=float)
        mat = rz @ ry @ rx
        if bool(getattr(mol, "mirror_x", False)):
            mat[0, :] *= -1.0
        if bool(getattr(mol, "mirror_y", False)):
            mat[1, :] *= -1.0
        return mat

    def _molecule_axis_screen_vectors(self, mol):
        mat = self._molecule_orientation_matrix(mol)
        vectors = []
        for idx, color in enumerate(("#ef4444", "#22c55e", "#3b82f6")):
            vec = np.asarray(mat[:, idx], dtype=float)
            screen = np.array([vec[0] + 0.38 * vec[2], vec[1] + 0.18 * vec[2]], dtype=float)
            vectors.append((idx, vec, screen, color))
        return vectors

    def _molecule_gizmo_hit_test(self, event):
        gizmo_ax = getattr(self, "_molecule_gizmo_axes", None)
        if gizmo_ax is None or not self._should_show_molecule_gizmo():
            return None
        if event is None or getattr(event, "x", None) is None or getattr(event, "y", None) is None:
            return None
        bbox = getattr(gizmo_ax, "bbox", None)
        if bbox is None or not bbox.contains(event.x, event.y):
            return None
        if getattr(event, "button", None) not in (1, None):
            return None
        width = float(getattr(bbox, "width", 0.0) or 0.0)
        height = float(getattr(bbox, "height", 0.0) or 0.0)
        if width <= 0.0 or height <= 0.0:
            return None
        local_x = ((float(event.x) - float(bbox.x0)) / width) * 2.0 - 1.0
        local_y = ((float(event.y) - float(bbox.y0)) / height) * 2.0 - 1.0
        radius = math.hypot(local_x, local_y)
        if radius > 1.02:
            return None
        mode = "rotate_z" if radius >= 0.58 else "rotate_xy"
        return {"mode": mode, "local": (float(local_x), float(local_y))}

    def _begin_molecule_gizmo_drag(self, hit, event):
        idx = self._active_molecule_index_for_gizmo()
        if idx is None:
            return False
        if not hit or event is None or getattr(event, "x", None) is None or getattr(event, "y", None) is None:
            return False
        try:
            self._push_molecule_snapshot()
        except Exception:
            pass
        self._active_molecule_idx = idx
        self._molecule_gizmo_drag = {
            "idx": idx,
            "mode": str(hit.get("mode") or "rotate_xy"),
            "start_px": (float(event.x), float(event.y)),
            "start_local": tuple(hit.get("local") or (0.0, 0.0)),
            "start_angles": np.array(self.molecules[idx].angles, dtype=float, copy=True),
        }
        self._wake_molecule_gizmo(2400, redraw=False)
        self._update_molecule_gizmo_overlay()
        return True

    def _wake_molecule_gizmo(self, duration_ms: int = 1800, *, redraw: bool = True):
        now_ms = time.perf_counter() * 1000.0
        self._molecule_gizmo_until = max(
            float(getattr(self, "_molecule_gizmo_until", 0.0) or 0.0),
            now_ms + float(duration_ms),
        )
        try:
            self._molecule_gizmo_timer.start(max(40, int(duration_ms)))
        except Exception:
            pass
        if not redraw and self._molecule_gizmo_axes is None:
            redraw = True
        if redraw:
            try:
                self._redraw()
            except Exception:
                try:
                    self.draw_idle()
                except Exception:
                    pass

    def _on_molecule_gizmo_timeout(self):
        if bool(getattr(self, "_show_molecule_gizmo", False)):
            return
        remaining = float(getattr(self, "_molecule_gizmo_until", 0.0) or 0.0) - (time.perf_counter() * 1000.0)
        if remaining > 40.0:
            try:
                self._molecule_gizmo_timer.start(int(remaining))
            except Exception:
                pass
            return
        self._molecule_gizmo_until = 0.0
        if self._molecule_gizmo_axes is not None:
            self._redraw()

    def _update_molecule_gizmo_overlay(self):
        gizmo_ax = getattr(self, "_molecule_gizmo_axes", None)
        artists = getattr(self, "_molecule_gizmo_artists", {}) or {}
        if gizmo_ax is None or not artists:
            return False
        mol = self._active_molecule_for_gizmo()
        if mol is None or not self._should_show_molecule_gizmo():
            return False
        vectors = self._molecule_axis_screen_vectors(mol)
        norms = [max(0.22, float(np.linalg.norm(screen))) for _, _, screen, _ in vectors]
        scale = 0.74 / max(norms or [1.0])
        for order, (idx, vec, screen, _color) in enumerate(sorted(vectors, key=lambda item: float(item[1][2]))):
            arrow = artists.get(f"arrow_{idx}")
            label = artists.get(f"label_{idx}")
            if arrow is None or label is None:
                continue
            end = screen * scale
            label_pos = screen * min(scale * 1.14, 0.92)
            alpha = 0.45 + 0.45 * ((float(vec[2]) + 1.0) * 0.5)
            try:
                arrow.set_positions((0.0, 0.0), (float(end[0]), float(end[1])))
                arrow.set_alpha(alpha)
                arrow.set_zorder(3 + order)
                label.set_position((float(label_pos[0]), float(label_pos[1])))
                label.set_alpha(max(0.68, alpha))
                label.set_zorder(6 + order)
            except Exception:
                continue
        try:
            gizmo_ax.figure.canvas.draw_idle()
        except Exception:
            pass
        return True

    def _draw_molecule_gizmo(self, ax):
        if ax is None or not self._should_show_molecule_gizmo():
            self._molecule_gizmo_axes = None
            self._molecule_gizmo_artists = {}
            return
        mol = self._active_molecule_for_gizmo()
        if mol is None:
            self._molecule_gizmo_axes = None
            self._molecule_gizmo_artists = {}
            return

        gizmo_ax = ax.inset_axes([0.02, 0.72, 0.16, 0.16], zorder=36)
        gizmo_ax.set_facecolor((0.0, 0.0, 0.0, 0.0))
        gizmo_ax.set_xticks([])
        gizmo_ax.set_yticks([])
        gizmo_ax.set_xlim(-1.0, 1.0)
        gizmo_ax.set_ylim(-1.0, 1.0)
        for spine in gizmo_ax.spines.values():
            spine.set_visible(False)

        dark = bool(getattr(self, "_detail_dark", False))
        ring_edge = "#d7dde7" if dark else "#334155"
        bg_face = (0.04, 0.06, 0.10, 0.42) if dark else (1.0, 1.0, 1.0, 0.74)
        ring = patches.Circle((0.0, 0.0), 0.94, facecolor=bg_face, edgecolor=ring_edge, linewidth=1.0, zorder=1)
        core = patches.Circle((0.0, 0.0), 0.06, facecolor=ring_edge, edgecolor="none", alpha=0.92, zorder=7)
        gizmo_ax.add_patch(ring)
        gizmo_ax.add_patch(core)

        artists = {"ring": ring, "core": core}
        for idx, (label_text, color) in enumerate((("X", "#ef4444"), ("Y", "#22c55e"), ("Z", "#3b82f6"))):
            arrow = patches.FancyArrowPatch(
                (0.0, 0.0),
                (0.0, 0.0),
                arrowstyle="-|>",
                mutation_scale=10.5,
                linewidth=2.2,
                color=color,
                alpha=0.92,
                zorder=4,
            )
            gizmo_ax.add_patch(arrow)
            label = gizmo_ax.text(
                0.0,
                0.0,
                label_text,
                color=color,
                fontsize=8.5,
                fontweight="bold",
                ha="center",
                va="center",
                zorder=8,
            )
            artists[f"arrow_{idx}"] = arrow
            artists[f"label_{idx}"] = label

        self._molecule_gizmo_axes = gizmo_ax
        self._molecule_gizmo_artists = artists
        self._update_molecule_gizmo_overlay()

    def _reset_molecule_to_file_state(self, idx=None, *, keep_offset: bool = True):
        if idx is None:
            idx = self._active_molecule_index_for_gizmo()
        if idx is None or idx < 0 or idx >= len(self.molecules):
            return False
        try:
            current = self.molecules[idx]
        except Exception:
            return False
        try:
            new_mol = current.reset_to_file_state(keep_offset=keep_offset)
        except Exception:
            return False
        self._push_molecule_snapshot()
        self.molecules[idx] = new_mol
        self._active_molecule_idx = idx
        self._wake_molecule_gizmo(2400, redraw=False)
        self._redraw()
        return True

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
        def _spec_identity(spec):
            if not spec:
                return None
            base = spec.get("path")
            try:
                base = str(Path(base))
            except Exception:
                base = str(base)
            idx = spec.get("matrix_index")
            if idx is not None:
                return f"{base}#idx{idx}"
            x = spec.get("x")
            y = spec.get("y")
            if x is not None or y is not None:
                try:
                    x_val = float(x) if x is not None else ""
                    y_val = float(y) if y is not None else ""
                    return f"{base}#pos{round(x_val, 6)}_{round(y_val, 6)}"
                except Exception:
                    return f"{base}#pos{x}_{y}"
            return base
        def _axis_from_pixel(col, row):
            row_use = float(row)
            # Stored spectro marker rows are in thumbnail/image pixel space
            # with row 0 at the top. Relative-axes preview flips the image and
            # draws it with origin='lower', so convert to the displayed row.
            if str(origin).lower() == 'lower' and rows > 0:
                row_use = float(rows) - row_use
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
                        y_axis = ymax - (row_use / float(rows)) * span_y
                    else:
                        y_axis = ymin + (row_use / float(rows)) * span_y
                return x_axis, y_axis
            x_axis = ex0 if cols == 0 else ex0 + (col / float(cols)) * (ex1 - ex0)
            y_axis = ey0 if rows == 0 else ey0 + (row_use / float(rows)) * (ey1 - ey0)
            return x_axis, y_axis
        normal_xs = []
        normal_ys = []
        highlight_xs = []
        highlight_ys = []
        points = []
        missing_specs = []
        highlight_spec = view.get('highlight_spec')
        highlight_key = _spec_identity(highlight_spec)
        pulse = float(getattr(self, "_highlight_pulse_strength", 1.0) or 1.0)
        pixel_lookup = {id(spec): (col, row) for spec, col, row in (view.get('spec_pixels') or [])}
        stack_badges = list(view.get("stack_badges") or [])
        for idx, s in enumerate(specs):
            coords = pixel_lookup.get(id(s))
            if coords is None:
                missing_specs.append(s)
                continue
            x, y = _axis_from_pixel(coords[0], coords[1])
            points.append((x, y, s))
            if highlight_key is not None and _spec_identity(s) == highlight_key:
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
                if highlight_key is not None and _spec_identity(spec) == highlight_key:
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
            for badge in stack_badges:
                try:
                    bx, by = _axis_from_pixel(float(badge.get("col")), float(badge.get("row")))
                    label = str(badge.get("label") or "").strip()
                    if not label:
                        continue
                    ax.text(
                        bx,
                        by,
                        label,
                        fontsize=7.2 * getattr(self, "_font_scale", 1.0),
                        fontweight="bold",
                        color="#ffe478",
                        ha="left",
                        va="bottom",
                        zorder=38,
                        bbox=dict(boxstyle="round,pad=0.2", facecolor="#281e12", edgecolor="#ffe0a0", linewidth=0.9, alpha=0.92),
                    )
                except Exception:
                    continue
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
        self.push_undo_state("scale_bar")
        self.scale_bar_enabled = enable
        if enable:
            self._connect_scale_bar_events()
        else:
            self._disconnect_scale_bar_events()
        self._redraw()
        self._notify_views_callback()

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

    def _scale_bar_span(self, ax, view):
        width = 0.0
        unit = view.get('axis_unit') or 'nm'
        try:
            xlim = ax.get_xlim()
            width = abs(float(xlim[1]) - float(xlim[0]))
        except Exception:
            width = 0.0
        if width > 0:
            if view.get('extent') is None and view.get('extent_raw') is None:
                unit = 'px'
            return width, unit
        extent = view.get('extent')
        if extent is None:
            extent = view.get('extent_raw')
        if extent is None:
            h, w = np.shape(view['arr'])
            return float(w), 'px'
        return abs(extent[1] - extent[0]), unit

    def _add_scale_bar(self, ax, view):
        width, unit = self._scale_bar_span(ax, view)
        size, label = self._calculate_best_scale_bar(width, unit)
        label = label if label and str(label).strip() else None
        
        font_scale = getattr(self, '_view_font_scale', 1.0)
        dark = bool(self._detail_dark)
        default_color = '#f5f5f5' if dark else '#111111'
        sb_settings = getattr(self, '_scale_bar_settings', {})
        sb_text_col = sb_settings.get('text_color') or default_color
        sb_bar_col = sb_settings.get('bar_color') or default_color
        font_family = sb_settings.get('font_family', 'sans-serif')
        sb = AnchoredSizeBar(ax.transData, size, label, 
                             loc='center',  # Anchor point on the artist itself
                             pad=0.4, borderpad=0, sep=3, 
                             frameon=False, 
                             size_vertical=width*0.004*font_scale,
                             color=sb_bar_col,
                             label_top=True,
                             bbox_to_anchor=self._scale_bar_pos,
                             bbox_transform=ax.transAxes)
        
        # Apply font scaling
        sb.size_bar.get_children()[0].set_linewidth(0) # remove border if any
        try:
            sb.size_bar.get_children()[0].set_color(sb_bar_col)
        except Exception:
            pass
        text = sb.txt_label.get_children()[0]
        text.set_color(sb_text_col)
        text.set_fontfamily(font_family)
        text.set_fontsize(10 * font_scale)
        text.set_fontweight('bold')
        try:
            apply_text_style(text, family=font_family, **self._plot_style_state())
        except Exception:
            pass
        sb.set_zorder(20)
        
        ax.add_artist(sb)
        self._scale_bar_artists.append(sb)

    def _refresh_scale_bars(self, ax=None, redraw: bool = False):
        keep = []
        for sb in list(self._scale_bar_artists):
            sb_ax = getattr(sb, "axes", None)
            if ax is not None and sb_ax is not ax:
                keep.append(sb)
                continue
            try:
                sb.remove()
            except Exception:
                if ax is not None and sb_ax is not ax:
                    keep.append(sb)
        self._scale_bar_artists = keep
        if not self.scale_bar_enabled:
            if redraw:
                self.draw_idle()
            return
        for target_ax, view in list(self._ax_view_map.items()):
            if ax is not None and target_ax is not ax:
                continue
            try:
                self._add_scale_bar(target_ax, view)
            except Exception:
                continue
        if redraw:
            self.draw_idle()

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

    def _normalize_profile_line_style(self, style, default="-"):
        style_key = str(style or default or "-").strip().lower()
        mapping = {
            "solid": "-",
            "-": "-",
            "dash": "--",
            "dashed": "--",
            "--": "--",
            "dot": ":",
            "dotted": ":",
            ":": ":",
            "dashdot": "-.",
            "-.": "-.",
            "none": "None",
            "": "None",
        }
        return mapping.get(style_key, default)

    def _normalize_profile_marker_style(self, style, default="o"):
        style_key = str(style or default or "o").strip().lower()
        mapping = {
            "none": "None",
            "": "None",
            "circle": "o",
            "o": "o",
            "square": "s",
            "s": "s",
            "triangle": "^",
            "^": "^",
            "diamond": "D",
            "d": "D",
            "plus": "+",
            "+": "+",
            "cross": "x",
            "x": "x",
            "star": "*",
            "*": "*",
        }
        return mapping.get(style_key, default)

    def _next_saved_profile_id(self):
        profile_id = f"profile-{int(getattr(self, '_profile_saved_profile_seq', 1))}"
        self._profile_saved_profile_seq = int(getattr(self, '_profile_saved_profile_seq', 1)) + 1
        return profile_id

    def _ensure_saved_profile_id(self, entry):
        if not isinstance(entry, dict):
            return ""
        profile_id = str(entry.get("profile_id") or "").strip()
        if not profile_id:
            profile_id = self._next_saved_profile_id()
            entry["profile_id"] = profile_id
        return profile_id

    def _profile_live_ref(self, profile_key=None, entry=None):
        source_id = register_profile_canvas(self)
        if not source_id:
            return None
        if entry is not None:
            profile_id = self._ensure_saved_profile_id(entry)
            if not profile_id:
                return None
            return {"source_id": source_id, "kind": "saved", "profile_id": profile_id}
        if profile_key is None:
            return {"source_id": source_id, "kind": "active"}
        try:
            idx = int(profile_key)
        except Exception:
            return None
        if idx < 0 or idx >= len(self._saved_profiles):
            return None
        profile_id = self._ensure_saved_profile_id(self._saved_profiles[idx])
        if not profile_id:
            return None
        return {"source_id": source_id, "kind": "saved", "profile_id": profile_id}

    def _saved_profile_index_from_ref(self, profile_ref):
        if not isinstance(profile_ref, dict):
            return None
        source_id = str(profile_ref.get("source_id") or "").strip()
        if source_id != str(register_profile_canvas(self) or "").strip():
            return None
        if str(profile_ref.get("kind") or "").strip().lower() == "active":
            return None
        profile_id = str(
            profile_ref.get("profile_id")
            or profile_ref.get("overlay_id")
            or ""
        ).strip()
        if not profile_id:
            return None
        for idx, entry in enumerate(self._saved_profiles):
            if str(self._ensure_saved_profile_id(entry) or "") == profile_id:
                return idx
        return None

    def _build_profile_style(self, *, color=None, lw=None, line_style=None, marker_style=None, marker_size=None, active=False):
        return {
            "color": color,
            "lw": float(lw if lw is not None else (self._active_profile_lw if active else 1.5)),
            "line_style": self._normalize_profile_line_style(line_style, "-" if active else "--"),
            "marker_style": self._normalize_profile_marker_style(marker_style, self._active_profile_marker_style if active else "o"),
            "marker_size": float(marker_size if marker_size is not None else (self._active_profile_marker_size if active else 5.0)),
        }

    def export_profile_state(self):
        saved = []
        for entry in self._saved_profiles:
            pts = entry.get('pts')
            if pts is None:
                continue
            saved.append({
                'pts': tuple(pts),
                'color': entry.get('color'),
                'lw': entry.get('lw'),
                'line_style': entry.get('line_style'),
                'marker_style': entry.get('marker_style'),
                'marker_size': entry.get('marker_size'),
            })
        state = {
            'active_pts': tuple(self.profile_pts) if self.profile_pts is not None else None,
            'saved': saved,
            'enabled': bool(self.profile_enabled),
            'user_enabled': bool(getattr(self, "_profile_user_enabled", self.profile_enabled)),
            'active_color': self._active_profile_color,
            'active_lw': float(self._active_profile_lw),
            'active_line_style': self._active_profile_line_style,
            'active_marker_style': self._active_profile_marker_style,
            'active_marker_size': float(self._active_profile_marker_size),
            'active_profile_original_id': self._active_profile_original_id,
            'marker_key': self._profile_marker_key,
            'marker_positions_by_key': dict(self._profile_marker_positions_by_key),
            'marker_domain_by_key': dict(self._profile_marker_domain_by_key),
        }
        return state

    def export_profile_datasets(self):
        """Return active/saved profile datasets for external dialogs."""
        active = self._build_profile_data(
            self.profile_pts,
            color=self._active_profile_color,
            lw=self._active_profile_lw,
            line_style=self._active_profile_line_style,
            marker_style=self._active_profile_marker_style,
            marker_size=self._active_profile_marker_size,
            live_profile_ref=self._profile_live_ref(None),
        )
        if isinstance(active, dict) and self._active_profile_original_id:
            active["profile_id"] = str(self._active_profile_original_id)
        saved = []
        for entry in self._saved_profiles:
            data = entry.get('data')
            if data is None:
                data = self._build_profile_data(
                    entry.get('pts'),
                    color=entry.get('color'),
                    lw=entry.get('lw'),
                    line_style=entry.get('line_style'),
                    marker_style=entry.get('marker_style'),
                    marker_size=entry.get('marker_size'),
                    live_profile_ref=self._profile_live_ref(entry=entry),
                )
                if isinstance(data, dict):
                    data["profile_id"] = str(self._ensure_saved_profile_id(entry))
                entry['data'] = data
            elif isinstance(data, dict):
                data['live_profile_ref'] = self._profile_live_ref(entry=entry)
                data["profile_id"] = str(self._ensure_saved_profile_id(entry))
            if data:
                saved.append(data)
        return active, saved

    @staticmethod
    def _normalize_profile_marker_key_map(mapping):
        """Restore JSON-loaded marker-key maps to native None/int keys."""
        normalized = {}
        for key, value in dict(mapping or {}).items():
            if key in (None, "null", "None", ""):
                normalized[None] = value
                continue
            try:
                normalized[int(key)] = value
            except Exception:
                normalized[key] = value
        return normalized

    def import_profile_state(self, state, emit=True):
        if state is None:
            return
        if self._profile_state_syncing:
            return
        try:
            self._profile_state_syncing = True
            active_pts = state.get('active_pts')
            saved = state.get('saved') or []
            enabled = bool(state.get('enabled', bool(active_pts is not None)))
            self._profile_user_enabled = bool(state.get('user_enabled', enabled))
            self._active_profile_color = state.get('active_color', self._active_profile_color) or self._active_profile_color
            try:
                self._active_profile_lw = float(state.get('active_lw', self._active_profile_lw))
            except Exception:
                pass
            self._active_profile_line_style = self._normalize_profile_line_style(
                state.get('active_line_style', self._active_profile_line_style),
                self._active_profile_line_style,
            )
            self._active_profile_marker_style = self._normalize_profile_marker_style(
                state.get('active_marker_style', self._active_profile_marker_style),
                self._active_profile_marker_style,
            )
            try:
                self._active_profile_marker_size = float(
                    state.get('active_marker_size', self._active_profile_marker_size)
                )
            except Exception:
                pass
            original_id = state.get('active_profile_original_id')
            self._active_profile_original_id = str(original_id).strip() if original_id else None
            marker_key = state.get('marker_key')
            if marker_key in ("null", "None", ""):
                marker_key = None
            elif marker_key is not None:
                try:
                    marker_key = int(marker_key)
                except Exception:
                    pass
            self._profile_marker_key = marker_key
            self._profile_marker_positions_by_key = self._normalize_profile_marker_key_map(
                state.get('marker_positions_by_key') or {}
            )
            self._profile_marker_domain_by_key = self._normalize_profile_marker_key_map(
                state.get('marker_domain_by_key') or {}
            )
            if not enabled and self.profile_enabled:
                self.enable_profile(False)
            if active_pts is not None:
                self._set_profile_pts(tuple(active_pts))
                if enabled and not self.profile_enabled:
                    self.enable_profile(True)
            self._clear_saved_profile_artists(notify=False)
            for entry in saved:
                pts = entry.get('pts')
                if pts is None:
                    continue
                self._add_saved_profile_from_pts(
                    tuple(pts),
                    entry.get('color'),
                    entry.get('lw'),
                    line_style=entry.get('line_style'),
                    marker_style=entry.get('marker_style'),
                    marker_size=entry.get('marker_size'),
                )
            if enabled:
                self._ensure_profile_artists()
                self._update_profile_artists()
            self._apply_profile_visibility()
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
        self.push_undo_state("angle_tool")
        self.angle_enabled = enable
        if enable:
            self._connect_angle_events()
            self._undo_suspend_depth += 1
            try:
                self._ensure_angle_frames()
            finally:
                self._undo_suspend_depth = max(0, self._undo_suspend_depth - 1)
            self._emit_angle()
        else:
            self._disconnect_angle_events()
            self._clear_angle_artists()
            self.angle_pts = None
            self._emit_angle()
        self.draw_idle()

    def clear_angle_measurement(self):
        if self.angle_enabled or self._angle_frames:
            self.push_undo_state("clear_angle")
        self._clear_angle_artists()
        if self.angle_enabled:
            self._undo_suspend_depth += 1
            try:
                self._ensure_angle_frames()
            finally:
                self._undo_suspend_depth = max(0, self._undo_suspend_depth - 1)
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
                apply_text_style(ax.xaxis.label, family=self._font_family, **self._plot_style_state())
                apply_text_style(ax.yaxis.label, family=self._font_family, **self._plot_style_state())
                apply_text_style(ax.title, family=self._font_family, **self._plot_style_state())
                for lbl in list(ax.get_xticklabels()) + list(ax.get_yticklabels()):
                    apply_text_style(lbl, family=self._font_family, **self._plot_style_state())
            except Exception:
                pass
        for cbar in getattr(self, '_colorbars', []):
            try:
                cbar.ax.tick_params(labelsize=tick_size)
                cbar.ax.yaxis.label.set_fontsize(label_size)
                cbar.ax.xaxis.label.set_fontsize(label_size)
                apply_text_style(cbar.ax.yaxis.label, family=self._font_family, **self._plot_style_state())
                apply_text_style(cbar.ax.xaxis.label, family=self._font_family, **self._plot_style_state())
                for lbl in list(cbar.ax.get_xticklabels()) + list(cbar.ax.get_yticklabels()):
                    apply_text_style(lbl, family=self._font_family, **self._plot_style_state())
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
        self.push_undo_state("profile_label_mode")
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

    def _notify_copy_feedback(self, view=None, *, fmt="png", displayed=False):
        if not callable(self._copy_feedback_handler):
            return
        payload = {
            "format": str(fmt or "png").lower(),
            "displayed": bool(displayed),
            "canvas": self,
        }
        ref_view = view
        if ref_view is None:
            ref_view = self.views[0] if self.views else {}
        try:
            self._copy_feedback_handler(ref_view, payload)
        except TypeError:
            try:
                self._copy_feedback_handler(ref_view)
            except Exception:
                pass
        except Exception:
            pass

    def get_overview_pixmap(self):
        """Return a pixmap snapshot of the current canvas (with overlays)."""
        try:
            return self.grab()
        except Exception:
            return None

    def _set_axes_titles_visible(self, visible: bool):
        changed = []
        for ax in getattr(self.fig, "axes", []) or []:
            try:
                title_artist = ax.title
            except Exception:
                continue
            if title_artist is None:
                continue
            try:
                title_text = str(title_artist.get_text() or "").strip()
                was_visible = bool(title_artist.get_visible())
            except Exception:
                continue
            if not title_text or was_visible == bool(visible):
                continue
            try:
                title_artist.set_visible(bool(visible))
                changed.append((title_artist, was_visible))
            except Exception:
                continue
        return changed

    def _render_displayed_pixmap(self, *, show_titles=True):
        buf = io.BytesIO()
        title_state = []

        def _save():
            nonlocal title_state
            if not show_titles:
                title_state = self._set_axes_titles_visible(False)
            try:
                self.fig.savefig(buf, format="png", dpi=300, bbox_inches="tight")
            finally:
                if title_state:
                    for title_artist, was_visible in title_state:
                        try:
                            title_artist.set_visible(bool(was_visible))
                        except Exception:
                            pass

        self._save_current_figure_without_shortcut_hint(_save)
        data = buf.getvalue()
        if not data:
            return None
        pixmap = QtGui.QPixmap()
        if not pixmap.loadFromData(data, "PNG"):
            return None
        if title_state:
            self.draw_idle()
        return pixmap

    def _resolve_powerpoint_label(self, view=None):
        if isinstance(view, dict):
            title = str(view.get("title") or "").strip()
            if title:
                return title
            path_text = str(view.get("path") or "").strip()
            if path_text:
                return Path(path_text).stem

        if len(self.views or []) == 1:
            only_view = self.views[0] or {}
            title = str(only_view.get("title") or "").strip()
            if title:
                return title

        try:
            window = self.window()
        except Exception:
            window = None
        if window is not None:
            try:
                window_title = str(window.windowTitle() or "").strip()
            except Exception:
                window_title = ""
            if window_title and window_title.lower() != "sxm viewer":
                return window_title
        return None

    def _show_powerpoint_success(self, slide_number, shape_name):
        _ = shape_name
        QtWidgets.QToolTip.showText(
            QtGui.QCursor.pos(),
            f"Sent to slide {slide_number}",
            self,
            self.rect(),
            2500,
        )

    def _send_displayed_to_powerpoint(self, view=None, *, new_slide=True):
        label_text = self._resolve_powerpoint_label(view)
        hide_titles = bool(label_text) and len(self.views or []) == 1
        pixmap = self._render_displayed_pixmap(show_titles=not hide_titles)
        if pixmap is None or pixmap.isNull():
            pixmap = self.get_overview_pixmap()

        try:
            slide_number, shape_name = send_pixmap_to_ppt(
                pixmap,
                label=label_text,
                new_slide=bool(new_slide),
            )
        except ConnectionError:
            QtWidgets.QMessageBox.critical(
                self,
                "PowerPoint",
                "PowerPoint is not running. Please open a presentation first.",
            )
            return
        except ValueError:
            QtWidgets.QMessageBox.warning(self, "PowerPoint", "No image to send.")
            return
        except EnvironmentError as exc:
            QtWidgets.QMessageBox.critical(self, "PowerPoint", str(exc))
            return
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, "PowerPoint", str(exc))
            return

        self._show_powerpoint_success(slide_number, shape_name)

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
            if self.profile_pts is not None:
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
        if event is not None:
            try:
                mods = event.modifiers()
                key = event.key()
            except Exception:
                mods = QtCore.Qt.NoModifier
                key = None
            if self._fixed_crop_transform_mode:
                if key in (QtCore.Qt.Key_Return, QtCore.Qt.Key_Enter):
                    self._on_apply_fixed_crop_shortcut()
                    return
                if key == QtCore.Qt.Key_Escape:
                    self._on_cancel_fixed_crop_shortcut()
                    return
            if mods & QtCore.Qt.ControlModifier:
                if key == QtCore.Qt.Key_C:
                    self._copy_displayed("png")
                    return
                if key == QtCore.Qt.Key_1:
                    self.set_show_profile_overlays(not self._show_profile_overlays)
                    return
                if key == QtCore.Qt.Key_2:
                    self.set_show_angle_overlays(not self._show_angle_overlays)
                    return
                if key == QtCore.Qt.Key_3:
                    self.set_show_molecules(not self.show_molecules)
                    return
                if key == QtCore.Qt.Key_4:
                    self.enable_scale_bar(not self.scale_bar_enabled)
                    return
                if key == QtCore.Qt.Key_5:
                    self.set_show_acquisition_overlay(not self._show_acquisition_overlay)
                    return
                if key == QtCore.Qt.Key_H:
                    self.set_show_shortcut_hint(not self._show_shortcut_hint)
                    return
            if not (mods & (QtCore.Qt.ControlModifier | QtCore.Qt.AltModifier | QtCore.Qt.MetaModifier)):
                if key == QtCore.Qt.Key_A and callable(self._histogram_auto_callback):
                    try:
                        self._histogram_auto_callback(self)
                    except Exception:
                        pass
                    return
                if key == QtCore.Qt.Key_M:
                    self._load_molecule_dialog()
                    return
            if (mods & QtCore.Qt.ControlModifier) and key == QtCore.Qt.Key_Z:
                if self.handle_undo_request():
                    return
            if not (mods & (QtCore.Qt.ControlModifier | QtCore.Qt.AltModifier | QtCore.Qt.MetaModifier)):
                if self._handle_popup_keyboard_shortcuts(key, shift=bool(mods & QtCore.Qt.ShiftModifier)):
                    return
        if event is not None:
            try:
                if event.modifiers() == QtCore.Qt.NoModifier and event.key() == QtCore.Qt.Key_R:
                    self._reset_view_zoom()
                    return
            except Exception:
                pass
        super().keyPressEvent(event)

    def _handle_popup_keyboard_shortcuts(self, key, *, shift: bool = False):
        """Handle popup shortcuts that should work without enabling tools explicitly."""
        if key is None:
            return False
        if self._rotate_selected_molecule_from_key(key, reverse=shift):
            return True
        if shift and key == QtCore.Qt.Key_R:
            return self._reset_molecule_to_file_state()
        if key in (QtCore.Qt.Key_0, QtCore.Qt.Key_Z):
            setter = getattr(self, "_popup_relative_zero_setter", None)
            if callable(setter):
                try:
                    setter(not bool(getattr(self, "_popup_relative_zero_enabled", False)))
                except Exception:
                    pass
                return True
        return False

    def _rotate_selected_molecule_from_key(self, key, *, reverse: bool = False):
        """Rotate the actively selected molecule around X/Y/Z using keyboard keys."""
        axis_map = {
            QtCore.Qt.Key_X: 0,
            QtCore.Qt.Key_Y: 1,
            QtCore.Qt.Key_Z: 2,
        }
        axis = axis_map.get(key)
        idx = getattr(self, "_active_molecule_idx", None)
        if axis is None or idx is None or idx < 0 or idx >= len(self.molecules):
            return False
        try:
            self._push_molecule_snapshot()
        except Exception:
            pass
        try:
            mol = self.molecules[idx]
            new_angles = np.array(mol.angles, dtype=float, copy=True)
            new_angles[axis] += -5.0 if reverse else 5.0
            mol.angles = new_angles
            self._update_molecule_artists()
            self._wake_molecule_gizmo(2200, redraw=False)
            self._update_molecule_gizmo_overlay()
        except Exception:
            return False
        return True

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
            if self.profile_pts is None:
                return
            x0, y0, x1, y1 = self.profile_pts
            self._set_profile_pts((x0, y0, x1, y1))
            x0, y0, x1, y1 = self.profile_pts
            color = self._active_profile_color
            self._profile_line, = self.main_ax.plot(
                [x0, x1], [y0, y1],
                color=color,
                lw=self._active_profile_lw,
                alpha=0.95,
                zorder=9,
                linestyle=self._active_profile_line_style,
            )
            self._profile_p0, = self.main_ax.plot(
                [x0], [y0],
                marker=self._active_profile_marker_style,
                linestyle='None',
                color=color,
                ms=self._active_profile_marker_size,
                mec='black',
                mew=1.0,
                zorder=10,
            )
            self._profile_p1, = self.main_ax.plot(
                [x1], [y1],
                marker=self._active_profile_marker_style,
                linestyle='None',
                color=color,
                ms=self._active_profile_marker_size,
                mec='black',
                mew=1.0,
                zorder=10,
            )
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
                l, = ax.plot(
                    [x0, x1], [y0, y1],
                    color=color,
                    lw=self._active_profile_lw,
                    alpha=0.95,
                    zorder=9,
                    linestyle=self._active_profile_line_style,
                )
                p0, = ax.plot(
                    [x0], [y0],
                    marker=self._active_profile_marker_style,
                    linestyle='None',
                    color=color,
                    ms=self._active_profile_marker_size,
                    mec='black',
                    mew=1.0,
                    zorder=10,
                )
                p1, = ax.plot(
                    [x1], [y1],
                    marker=self._active_profile_marker_style,
                    linestyle='None',
                    color=color,
                    ms=self._active_profile_marker_size,
                    mec='black',
                    mew=1.0,
                    zorder=10,
                )
                self._profile_echo_artists.append({'line': l, 'p0': p0, 'p1': p1})
            except Exception:
                pass
        self._apply_profile_visibility()

    def _ensure_angle_frames(self):
        if self.main_ax is None:
            return
        if not self._angle_frames:
            self._add_angle_frame_at(None, None)
        for frame in self._angle_frames:
            self._ensure_frame_artists(frame)

    def _add_angle_frame_at(self, x, y):
        self.push_undo_state("add_angle_frame")
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
        self._apply_angle_visibility()
        if self._angle_blit_active and self._angle_background is not None:
            self._blit_angle_frames()
        else:
            self.draw_idle()

    def _apply_angle_visibility(self):
        overlay_visible = bool(self._show_angle_overlays)
        active_idx = self._active_angle_frame_idx
        for idx, frame in enumerate(self._angle_frames or []):
            # Keep the active measurement visible while the tool is enabled.
            frame_visible = bool(self.angle_enabled and idx == active_idx) or overlay_visible
            style = frame.get("style", "dots")
            for art in frame.get("lines", []) or []:
                if art is None:
                    continue
                try:
                    art.set_visible(frame_visible and style == "dots")
                except Exception:
                    pass
            for art in frame.get("markers", []) or []:
                if art is None:
                    continue
                try:
                    art.set_visible(frame_visible and style == "dots")
                except Exception:
                    pass
            for art in frame.get("arrows", []) or []:
                if art is None:
                    continue
                try:
                    art.set_visible(frame_visible and style == "arrows")
                except Exception:
                    pass
            label = frame.get("label")
            if label is not None:
                try:
                    label.set_visible(frame_visible)
                except Exception:
                    pass
            patch = frame.get("patch")
            if patch is not None:
                try:
                    patch.set_visible(frame_visible)
                except Exception:
                    pass
            for lbl in frame.get("len_labels", []) or []:
                try:
                    lbl.set_visible(frame_visible)
                except Exception:
                    pass

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
        if style == frame.get('style', 'dots'):
            return
        self.push_undo_state("angle_style")
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
        self._apply_profile_visibility()
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
        self._apply_profile_visibility()
        if draw:
            self.draw_idle()

    def _apply_profile_visibility(self):
        active_visible = bool(self.profile_pts is not None and self._show_profile_overlays)
        overlay_visible = bool(self._show_profile_overlays)
        active_artists = [
            self._profile_line,
            self._profile_p0,
            self._profile_p1,
            self._profile_ticks,
            self._profile_info_text,
            self._profile_label,
            self._profile_hud_text,
        ]
        active_artists.extend(list(self._profile_endpoint_labels or []))
        active_artists.extend(list(self._profile_marker_artists or []))
        for echo in self._profile_echo_artists or []:
            if isinstance(echo, dict):
                active_artists.extend(echo.values())
        for art in active_artists:
            if art is None:
                continue
            try:
                art.set_visible(active_visible)
            except Exception:
                continue
        for entry in self._saved_profiles or []:
            for art in entry.get("artists", []) or []:
                if art is None:
                    continue
                try:
                    art.set_visible(overlay_visible)
                except Exception:
                    continue

    def _schedule_profile_update(self):
        if not self._profile_update_timer.isActive():
            self._profile_update_timer.start()

    def _flush_profile_updates(self):
        if not (self.profile_enabled or self._profile_move_only):
            return
        if self.profile_pts is None:
            return
        if self._dragging is not None:
            self._schedule_profile_update()
            return
        self._update_profile_markers()
        self._emit_profile()
        self.draw_idle()

    def activate_saved_profile(self, index, push_undo=True):
        """Promote a saved overlay back to the active profile line."""
        if index is None or not self._saved_profiles:
            return False
        try:
            idx = int(index)
        except Exception:
            return False
        if idx < 0 or idx >= len(self._saved_profiles):
            return False
        if push_undo:
            self.push_undo_state("activate_profile")
        entry = self._saved_profiles.pop(idx)
        self._active_profile_original_color = entry.get('color')
        self._active_profile_original_id = str(self._ensure_saved_profile_id(entry))
        entry_color = entry.get('color')
        if entry_color:
            self._active_profile_color = str(entry_color)
        try:
            self._active_profile_lw = float(entry.get('lw', self._active_profile_lw))
        except Exception:
            pass
        self._active_profile_line_style = self._normalize_profile_line_style(
            entry.get('line_style'),
            self._active_profile_line_style,
        )
        self._active_profile_marker_style = self._normalize_profile_marker_style(
            entry.get('marker_style'),
            self._active_profile_marker_style,
        )
        try:
            self._active_profile_marker_size = float(entry.get('marker_size', self._active_profile_marker_size))
        except Exception:
            pass
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
        color = self._active_profile_color or '#fbc02d'
        
        # Reuse existing artists if possible
        if self._profile_ticks is not None and self._profile_info_text is not None:
            self._remove_profile_markers() # Fallback to recreate if complex update needed, or optimize further
            ticks, text = self._create_ticks_and_label(self.profile_pts, color=color, alpha=0.9, base_size=9)
        else:
            self._remove_profile_markers()
            ticks, text = self._create_ticks_and_label(self.profile_pts, color=color, alpha=0.9, base_size=9)
            
        self._profile_ticks = ticks
        self._profile_info_text = text
        self._update_profile_marker_artists()
        self._update_profile_labels()
        self._update_profile_hud()
        self._apply_profile_visibility()

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

    def _build_profile_data(self, pts, color=None, view=None, lw=None, line_style=None, marker_style=None, marker_size=None, live_profile_ref=None):
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
            source_path = str(v0.get('path') or (meta or {}).get('path') or (meta or {}).get('file_path') or "").strip()
            source_title = str(v0.get('title') or "").strip()
            source_acq = str(self._acquisition_overlay_text(v0) or "").strip()
            source_folder = ""
            source_folder_name = ""
            source_name = ""
            source_datetime = str(v0.get("datetime") or (meta or {}).get("datetime") or "").strip()
            source_date = str(v0.get("date") or (meta or {}).get("date") or "").strip()
            source_time = str(v0.get("time") or (meta or {}).get("time") or "").strip()
            if source_path:
                try:
                    p = Path(source_path)
                    source_folder = str(p.parent)
                    source_folder_name = p.parent.name
                    source_name = p.name
                except Exception:
                    source_name = source_path
            return {
                'x_px': x_px,
                'x_nm': x_phys,
                'vals': vals,
                'length_nm': length_nm,
                'unit': unit,
                'axis_unit': axis_unit,
                'distance_unit': distance_unit if x_phys is not None else 'px',
                'color': color,
                'lw': float(lw) if lw is not None else None,
                'line_style': self._normalize_profile_line_style(line_style, '-'),
                'marker_style': self._normalize_profile_marker_style(marker_style, 'o'),
                'marker_size': float(marker_size) if marker_size is not None else None,
                'label': self._format_profile_label(pts),
                'relative_axes': self._use_relative_axes(v0),
                'meta': meta,
                'source_path': source_path,
                'source_folder': source_folder,
                'source_folder_name': source_folder_name,
                'source_file_name': source_name,
                'source_title': source_title,
                'source_acquisition_text': source_acq,
                'source_datetime': source_datetime,
                'source_date': source_date,
                'source_time': source_time,
                'live_profile_ref': copy.deepcopy(live_profile_ref) if isinstance(live_profile_ref, dict) else None,
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
        self.push_undo_state("snapshot_profile")
        pts = tuple(self.profile_pts)
        original_profile_id = str(self._active_profile_original_id or "").strip() or None
        if self._active_profile_original_color:
            color = self._active_profile_original_color
            self._active_profile_original_color = None
        else:
            color = next(self._profile_color_cycle)
        lw = self._active_profile_lw
        line_style = self._normalize_profile_line_style(self._active_profile_line_style, '--')
        marker_style = self._normalize_profile_marker_style(self._active_profile_marker_style, 'o')
        marker_size = float(self._active_profile_marker_size)
        line, = self.main_ax.plot(
            [pts[0], pts[2]], [pts[1], pts[3]],
            color=color,
            lw=lw,
            alpha=0.7,
            zorder=6,
            linestyle=line_style,
        )
        # Combine endpoints into one artist
        endpoints, = self.main_ax.plot(
            [pts[0], pts[2]], [pts[1], pts[3]],
            marker=marker_style,
            linestyle='None',
            color=color,
            ms=marker_size,
            mec='black',
            mew=0.7,
            alpha=0.9,
            zorder=7,
        )
        
        artists = [line, endpoints]
        
        # Add echo artists for other views
        for ax in self._ax_view_map:
            if ax is self.main_ax:
                continue
            try:
                l, = ax.plot([pts[0], pts[2]], [pts[1], pts[3]],
                             color=color, lw=lw, alpha=0.7, zorder=6, linestyle=line_style)
                ep, = ax.plot([pts[0], pts[2]], [pts[1], pts[3]], 
                              marker=marker_style, linestyle='None', color=color,
                              ms=marker_size, mec='black', mew=0.7, alpha=0.9, zorder=7)
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
        data = self._build_profile_data(
            pts,
            color=color,
            lw=lw,
            line_style=line_style,
            marker_style=marker_style,
            marker_size=marker_size,
        )
        profile_id = original_profile_id or self._next_saved_profile_id()
        entry = {'artists': artists, 'pts': pts, 'color': color, 'data': data,
                 'overlay_label_artist': overlay_label, 'endpoint_labels': endpoint_labels, 'lw': lw,
                 'line_style': line_style, 'marker_style': marker_style, 'marker_size': marker_size,
                 'line_artist': line, 'endpoint_artist': endpoints, 'profile_id': profile_id}
        self._active_profile_original_id = None
        if isinstance(data, dict):
            data['live_profile_ref'] = self._profile_live_ref(entry=entry)
        if text is not None:
            entry['label_artist'] = text
            entry['label_base_size'] = base_size
        self._saved_profiles.append(entry)
        self._refresh_overlay_labels()
        self._apply_profile_visibility()
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
        self.push_undo_state("delete_profile")
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
        self.push_undo_state("remove_profile")
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

    def _saved_profile_hit(self, x, y, thresh=18.0):
        if x is None or y is None or not self._saved_profiles:
            return None
        for idx in reversed(range(len(self._saved_profiles))):
            entry = self._saved_profiles[idx]
            pts = entry.get('pts')
            if pts is None:
                continue
            try:
                x0, y0, x1, y1 = pts
                d0 = self._pt_distance_pixels(x, y, x0, y0)
                d1 = self._pt_distance_pixels(x, y, x1, y1)
                if d0 <= thresh or d1 <= thresh:
                    return {"idx": idx, "mode": "p0" if d0 <= d1 else "p1"}
                dist_line = self._distance_to_segment_pixels(x, y, pts)
                if dist_line <= thresh:
                    return {"idx": idx, "mode": "line"}
            except Exception:
                continue
        return None

    def _select_saved_profile_overlay(self, idx):
        if idx is None or idx < 0 or idx >= len(self._saved_profiles):
            self.highlight_saved_profile(None)
            if callable(self._profile_highlight_cb):
                try:
                    self._profile_highlight_cb(None)
                except Exception:
                    pass
            return
        self.highlight_saved_profile(idx)
        if callable(self._profile_highlight_cb):
            try:
                self._profile_highlight_cb(idx)
            except Exception:
                pass
        self.set_profile_marker_key(idx)

    def _rebuild_saved_profile_entry(self, idx, pts, *, redraw=True):
        if idx < 0 or idx >= len(self._saved_profiles) or self.main_ax is None or pts is None:
            return False
        prev = self._saved_profiles[idx]
        for art in prev.get('artists', []):
            try:
                if art is not None:
                    art.remove()
            except Exception:
                pass
        pts = tuple(pts)
        color = prev.get('color') or '#ffffff'
        lw = float(prev.get('lw', 1.5) or 1.5)
        line_style = self._normalize_profile_line_style(prev.get('line_style'), '--')
        marker_style = self._normalize_profile_marker_style(prev.get('marker_style'), 'o')
        marker_size = float(prev.get('marker_size', 5.0) or 5.0)
        line, = self.main_ax.plot(
            [pts[0], pts[2]], [pts[1], pts[3]],
            color=color, lw=lw, alpha=0.7, zorder=6, linestyle=line_style
        )
        endpoints, = self.main_ax.plot(
            [pts[0], pts[2]], [pts[1], pts[3]],
            marker=marker_style, linestyle='None', color=color,
            ms=marker_size, mec='black', mew=0.7, alpha=0.9, zorder=7
        )
        artists = [line, endpoints]
        for ax in self._ax_view_map:
            if ax is self.main_ax:
                continue
            try:
                l, = ax.plot(
                    [pts[0], pts[2]], [pts[1], pts[3]],
                    color=color, lw=lw, alpha=0.7, zorder=6, linestyle=line_style
                )
                ep, = ax.plot(
                    [pts[0], pts[2]], [pts[1], pts[3]],
                    marker=marker_style, linestyle='None', color=color,
                    ms=marker_size, mec='black', mew=0.7, alpha=0.9, zorder=7
                )
                artists.extend([l, ep])
            except Exception:
                pass
        base_size = int(prev.get('label_base_size', 8) or 8)
        ticks, text = self._create_ticks_and_label(pts, color=color, alpha=0.7, base_size=base_size)
        overlay_label = self._create_profile_id_label(pts, f"Overlay {idx + 1}", color)
        endpoint_labels = self._create_endpoint_labels(pts, color)
        if overlay_label is not None:
            try:
                overlay_label.set_visible(False)
            except Exception:
                pass
            artists.append(overlay_label)
        for lbl in endpoint_labels:
            try:
                lbl.set_visible(False)
            except Exception:
                pass
        if ticks is not None:
            artists.append(ticks)
        if text is not None:
            artists.append(text)
        artists += endpoint_labels
        data = self._build_profile_data(
            pts,
            color=color,
            lw=lw,
            line_style=line_style,
            marker_style=marker_style,
            marker_size=marker_size,
            live_profile_ref=self._profile_live_ref(entry=prev),
        )
        if isinstance(data, dict):
            data["profile_id"] = str(self._ensure_saved_profile_id(prev))
        entry = {
            'artists': artists,
            'pts': pts,
            'color': color,
            'data': data,
            'overlay_label_artist': overlay_label,
            'endpoint_labels': endpoint_labels,
            'lw': lw,
            'line_style': line_style,
            'marker_style': marker_style,
            'marker_size': marker_size,
            'line_artist': line,
            'endpoint_artist': endpoints,
            'profile_id': self._ensure_saved_profile_id(prev),
        }
        if text is not None:
            entry['label_artist'] = text
            entry['label_base_size'] = base_size
        self._saved_profiles[idx] = entry
        self._refresh_overlay_labels()
        if self._highlighted_overlay is not None:
            self.highlight_saved_profile(self._highlighted_overlay)
        else:
            self._apply_profile_visibility()
        if self._profile_marker_key == idx:
            self._update_profile_marker_artists()
            self._update_profile_hud()
        if redraw:
            self.draw_idle()
        return True

    def _undo_last_profile_snapshot(self):
        if not self._saved_profiles:
            return False
        entry = self._saved_profiles.pop()
        for art in entry.get('artists', []):
            try:
                if art is not None:
                    art.remove()
            except Exception:
                pass
        self.draw_idle()
        self._emit_profile()
        return True

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

    def _add_saved_profile_from_pts(self, pts, color, lw=1.5, line_style='--', marker_style='o', marker_size=5.0):
        if pts is None or self.main_ax is None:
            return
        pts = tuple(pts)
        color = color or next(self._profile_color_cycle)
        lw = float(lw or 1.5)
        line_style = self._normalize_profile_line_style(line_style, '--')
        marker_style = self._normalize_profile_marker_style(marker_style, 'o')
        marker_size = float(marker_size or 5.0)
        line, = self.main_ax.plot([pts[0], pts[2]], [pts[1], pts[3]],
                                  color=color, lw=lw, alpha=0.7, zorder=6, linestyle=line_style)
        endpoints, = self.main_ax.plot([pts[0], pts[2]], [pts[1], pts[3]], marker=marker_style, linestyle='None', color=color,
                                       ms=marker_size, mec='black', mew=0.7, alpha=0.9, zorder=7)
        
        artists = [line, endpoints]
        
        # Add echo artists for other views
        for ax in self._ax_view_map:
            if ax is self.main_ax:
                continue
            try:
                l, = ax.plot([pts[0], pts[2]], [pts[1], pts[3]],
                             color=color, lw=lw, alpha=0.7, zorder=6, linestyle=line_style)
                ep, = ax.plot([pts[0], pts[2]], [pts[1], pts[3]], 
                              marker=marker_style, linestyle='None', color=color,
                              ms=marker_size, mec='black', mew=0.7, alpha=0.9, zorder=7)
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
        data = self._build_profile_data(
            pts,
            color=color,
            lw=lw,
            line_style=line_style,
            marker_style=marker_style,
            marker_size=marker_size,
        )
        entry = {'artists': artists, 'pts': pts, 'color': color, 'data': data,
                 'overlay_label_artist': overlay_label, 'endpoint_labels': endpoint_labels, 'lw': lw,
                 'line_style': line_style, 'marker_style': marker_style, 'marker_size': marker_size,
                 'line_artist': line, 'endpoint_artist': endpoints,
                 'profile_id': self._next_saved_profile_id()}
        if isinstance(data, dict):
            data['live_profile_ref'] = self._profile_live_ref(entry=entry)
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
        if self._saved_profiles:
            self.push_undo_state("clear_profiles")
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
        if (not self.profile_enabled and not self._profile_move_only) or event.inaxes is None or event.inaxes is not self.main_ax:
            return
        x, y = event.xdata, event.ydata
        if x is None or y is None:
            return
        shift_pressed = self._shift_pressed(event)
        mods_qt = self._event_qt_modifiers(event)
        ctrl_pressed = bool(mods_qt & QtCore.Qt.ControlModifier)
        
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
        saved_hit = self._saved_profile_hit(x, y)
        if saved_hit is not None:
            overlay_idx = int(saved_hit.get("idx"))
            self._select_saved_profile_overlay(overlay_idx)
            self.push_undo_state("move_saved_profile")
            self._saved_profile_drag = {
                "idx": overlay_idx,
                "mode": str(saved_hit.get("mode") or "line"),
                "origin": tuple(self._saved_profiles[overlay_idx].get("pts") or (x, y, x, y)),
                "start": (x, y),
            }
            return
        if self.profile_pts is None:
            if not ctrl_pressed:
                return
            if self._profile_move_only:
                return
            self.push_undo_state("start_profile")
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
                self.push_undo_state("move_profile")
                self._dragging = 'p0'
                self._line_drag_origin = None
                self._set_profile_animated(True)
                self._prepare_profile_blit()
                return
        if d1 <= thresh:
            self.push_undo_state("move_profile")
            self._dragging = 'p1'
            self._line_drag_origin = None
            self._set_profile_animated(True)
            self._prepare_profile_blit()
            return
        if self.profile_pts is not None:
            dist_line = self._distance_to_segment_pixels(x, y, self.profile_pts)
            if dist_line <= thresh:
                self.push_undo_state("move_profile")
                self._dragging = 'line'
                self._line_drag_origin = (x, y, self.profile_pts)
                self._set_profile_animated(True)
                self._prepare_profile_blit()
                return
        # Increased threshold to 15.0 to prevent accidental "misses" causing profile loss
        overlay_idx = self._overlay_index_near(x, y, thresh=15.0)
        if overlay_idx is not None:
            self._select_saved_profile_overlay(overlay_idx)
            self._dragging = None
            self._line_drag_origin = None
            return
        # else: start a new line from here
        if not ctrl_pressed:
            return
        if self.profile_pts is not None:
            self._snapshot_active_profile()
        else:
            self.push_undo_state("start_profile")
        if self._profile_move_only:
            return
        self._active_profile_original_color = None
        self._active_profile_original_id = None
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

    def _apply_active_profile_style(self):
        color = self._active_profile_color
        if self._profile_line is not None:
            try:
                self._profile_line.set_color(color)
                self._profile_line.set_linewidth(self._active_profile_lw)
                self._profile_line.set_linestyle(self._active_profile_line_style)
            except Exception:
                pass
        for point in (self._profile_p0, self._profile_p1):
            if point is None:
                continue
            try:
                point.set_color(color)
                point.set_marker(self._active_profile_marker_style)
                point.set_markersize(self._active_profile_marker_size)
                point.set_markeredgecolor('black')
            except Exception:
                pass
        for entry in self._profile_echo_artists or []:
            try:
                if entry.get('line') is not None:
                    entry['line'].set_color(color)
                    entry['line'].set_linewidth(self._active_profile_lw)
                    entry['line'].set_linestyle(self._active_profile_line_style)
                for key in ('p0', 'p1'):
                    artist = entry.get(key)
                    if artist is None:
                        continue
                    artist.set_color(color)
                    artist.set_marker(self._active_profile_marker_style)
                    artist.set_markersize(self._active_profile_marker_size)
                    artist.set_markeredgecolor('black')
            except Exception:
                pass
        for artist in list(self._profile_endpoint_labels or []) + [self._profile_label, self._profile_info_text]:
            if artist is None:
                continue
            try:
                artist.set_color(color)
            except Exception:
                pass
        if self._profile_ticks is not None:
            try:
                self._profile_ticks.set_color(color)
                self._profile_ticks.set_markersize(max(3.0, 4.0 * self._profile_label_scale))
            except Exception:
                pass

    def _apply_saved_profile_style(self, entry):
        if not entry:
            return
        color = entry.get('color') or '#ffffff'
        lw = float(entry.get('lw', 1.5) or 1.5)
        line_style = self._normalize_profile_line_style(entry.get('line_style'), '--')
        marker_style = self._normalize_profile_marker_style(entry.get('marker_style'), 'o')
        marker_size = float(entry.get('marker_size', 5.0) or 5.0)
        line_artists = []
        endpoint_artists = []
        for art in entry.get('artists', []) or []:
            if art is None:
                continue
            try:
                linestyle = art.get_linestyle() if hasattr(art, 'get_linestyle') else None
            except Exception:
                linestyle = None
            if linestyle not in (None, 'None', 'none', ''):
                line_artists.append(art)
            elif hasattr(art, 'set_marker'):
                endpoint_artists.append(art)
        for art in line_artists:
            try:
                art.set_color(color)
                art.set_linewidth(lw)
                art.set_linestyle(line_style)
            except Exception:
                pass
        for art in endpoint_artists:
            try:
                art.set_color(color)
                art.set_marker(marker_style)
                art.set_markersize(marker_size)
                if hasattr(art, 'set_markeredgecolor'):
                    art.set_markeredgecolor('black')
            except Exception:
                pass
        for artist in [entry.get('overlay_label_artist'), entry.get('label_artist')] + list(entry.get('endpoint_labels') or []):
            if artist is None:
                continue
            try:
                artist.set_color(color)
            except Exception:
                pass
        entry['data'] = None

    def set_profile_style(self, profile_key=None, profile_ref=None, **changes):
        """Update active or saved profile styling and refresh emitted datasets."""
        idx = None
        if profile_ref is not None:
            if str(profile_ref.get("kind") or "").strip().lower() == "active":
                if str(profile_ref.get("source_id") or "").strip() != str(register_profile_canvas(self) or "").strip():
                    return False
                active = True
            else:
                idx = self._saved_profile_index_from_ref(profile_ref)
                if idx is None:
                    return False
                active = False
        else:
            active = profile_key is None
            if not active:
                try:
                    idx = int(profile_key)
                except Exception:
                    return False
                if idx < 0 or idx >= len(self._saved_profiles):
                    return False
        if active and self.profile_pts is None:
            return False
        self.push_undo_state("profile_style")
        if active:
            color = changes.get('color')
            if color:
                self._active_profile_color = str(color)
            lw = changes.get('lw')
            if lw is not None:
                self._active_profile_lw = max(0.5, float(lw))
            line_style = changes.get('line_style')
            if line_style is not None:
                self._active_profile_line_style = self._normalize_profile_line_style(
                    line_style, self._active_profile_line_style
                )
            marker_style = changes.get('marker_style')
            if marker_style is not None:
                self._active_profile_marker_style = self._normalize_profile_marker_style(
                    marker_style, self._active_profile_marker_style
                )
            marker_size = changes.get('marker_size')
            if marker_size is not None:
                self._active_profile_marker_size = max(2.0, float(marker_size))
            self._apply_active_profile_style()
        else:
            entry = self._saved_profiles[idx]
            color = changes.get('color')
            if color:
                entry['color'] = str(color)
            lw = changes.get('lw')
            if lw is not None:
                entry['lw'] = max(0.5, float(lw))
            line_style = changes.get('line_style')
            if line_style is not None:
                entry['line_style'] = self._normalize_profile_line_style(line_style, entry.get('line_style', '--'))
            marker_style = changes.get('marker_style')
            if marker_style is not None:
                entry['marker_style'] = self._normalize_profile_marker_style(marker_style, entry.get('marker_style', 'o'))
            marker_size = changes.get('marker_size')
            if marker_size is not None:
                entry['marker_size'] = max(2.0, float(marker_size))
            self._apply_saved_profile_style(entry)
        self.draw_idle()
        self._emit_profile()
        return True

    def apply_profile_palette(self, palette_name: str):
        """Apply a named color cycle to the active profile and saved overlays."""
        colors = get_color_cycle(palette_name)
        if not colors:
            return False
        self.push_undo_state("profile_palette")
        self._profile_palette_name = palette_name or DEFAULT_COLOR_CYCLE
        self._profile_palette_colors = list(colors)
        self._profile_color_cycle = itertools.cycle(self._profile_palette_colors)
        if self.profile_pts is not None:
            self._active_profile_color = self._profile_palette_colors[0]
            self._apply_active_profile_style()
        for idx, entry in enumerate(self._saved_profiles):
            entry['color'] = self._profile_palette_colors[(idx + 1) % len(self._profile_palette_colors)]
            self._apply_saved_profile_style(entry)
        self.draw_idle()
        self._emit_profile()
        return True

    def _change_profile_color(self, overlay_idx, active):
        current_color = self._active_profile_color
        if overlay_idx is not None and 0 <= overlay_idx < len(self._saved_profiles):
            current_color = self._saved_profiles[overlay_idx].get('color', current_color)
        
        col = QtWidgets.QColorDialog.getColor(QtGui.QColor(current_color), self, "Select Profile Color")
        if not col.isValid(): return
        self.set_profile_style(None if active else overlay_idx, color=col.name())

    def _change_profile_width(self, overlay_idx, active, delta):
        has_overlay = overlay_idx is not None and 0 <= overlay_idx < len(self._saved_profiles)
        if not active and not has_overlay:
            return
        if active:
            self.set_profile_style(None, lw=max(0.5, self._active_profile_lw + delta))
        elif has_overlay:
            entry = self._saved_profiles[overlay_idx]
            self.set_profile_style(overlay_idx, lw=max(0.5, float(entry.get('lw', 1.5) or 1.5) + delta))

    def _on_motion(self, event):
        if (not self.profile_enabled and not self._profile_move_only) or event.inaxes is None or event.inaxes is not self.main_ax:
            return
        x, y = event.xdata, event.ydata
        if x is None or y is None:
            return
        if self._saved_profile_drag is not None:
            drag = self._saved_profile_drag
            idx = drag.get("idx")
            if idx is None or idx < 0 or idx >= len(self._saved_profiles):
                self._saved_profile_drag = None
                return
            pts = drag.get("origin") or self._saved_profiles[idx].get("pts")
            if pts is None:
                return
            x0, y0, x1, y1 = pts
            mode = drag.get("mode")
            if mode == "p0":
                new_pts = (x, y, x1, y1)
            elif mode == "p1":
                new_pts = (x0, y0, x, y)
            else:
                sx, sy = drag.get("start", (x, y))
                dx = x - sx
                dy = y - sy
                new_pts = (x0 + dx, y0 + dy, x1 + dx, y1 + dy)
            self._rebuild_saved_profile_entry(idx, new_pts, redraw=True)
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
        if not (self.profile_enabled or self._profile_move_only):
            return
        if self._saved_profile_drag is not None:
            self._saved_profile_drag = None
            self._emit_profile()
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
        if getattr(self, "_profile_quick_transient", False):
            self._profile_quick_transient = False
            self._profile_user_enabled = False
            self._profile_move_only = self.profile_pts is not None

    def _profile_hit_test(self, event, *, thresh: float = 18.0):
        if event is None or event.inaxes is not self.main_ax:
            return False
        x, y = event.xdata, event.ydata
        if x is None or y is None:
            return False
        if self._profile_marker_hit(x, y) is not None:
            return True
        if self.profile_pts is not None:
            x0, y0, x1, y1 = self.profile_pts
            if self._pt_distance_pixels(x, y, x0, y0) <= thresh:
                return True
            if self._pt_distance_pixels(x, y, x1, y1) <= thresh:
                return True
            if self._distance_to_segment_pixels(x, y, self.profile_pts) <= thresh:
                return True
        return self._overlay_index_near(x, y, thresh=15.0) is not None

    def _angle_hit_test(self, event, *, thresh: float = 12.0):
        if event is None or event.inaxes is not self.main_ax:
            return False
        x, y = event.xdata, event.ydata
        if x is None or y is None:
            return False
        for frame in self._angle_frames or []:
            pts = frame.get("pts")
            if not pts:
                continue
            vx, vy, ax, ay, bx, by = pts
            for hx, hy in ((vx, vy), (ax, ay), (bx, by)):
                if self._pt_distance_pixels(x, y, hx, hy) <= thresh:
                    return True
        return False

    def _molecule_overlay_hit_test(self, event, *, thresh: float = 14.0):
        if (
            event is None
            or not self.show_molecules
            or not self.molecules
            or event.inaxes is None
            or event.xdata is None
            or event.ydata is None
            or getattr(event, "x", None) is None
            or getattr(event, "y", None) is None
        ):
            return False
        try:
            ev_px = np.array([float(event.x), float(event.y)], dtype=float)
            for mol in reversed(list(self.molecules)):
                coords = mol.get_transformed_coordinates()
                if len(coords) == 0:
                    continue
                if not getattr(self, "_show_hydrogens", True):
                    mask = [str(el).strip().upper() != "H" for el in mol.elements]
                    if not any(mask):
                        continue
                    coords = coords[np.array(mask)]
                pts_px = event.inaxes.transData.transform(coords[:, :2])
                dists = np.hypot(pts_px[:, 0] - ev_px[0], pts_px[:, 1] - ev_px[1])
                if dists.size and float(np.nanmin(dists)) <= float(thresh):
                    return True
        except Exception:
            pass
        return False

    def _interactive_overlay_hit_test(self, event):
        if event is None:
            return False
        if self.scale_bar_enabled:
            for sb in self._scale_bar_artists:
                try:
                    if sb.contains(event)[0]:
                        return True
                except Exception:
                    continue
        if self._molecule_gizmo_hit_test(event) is not None:
            return True
        ax = getattr(event, "inaxes", None)
        view = self._ax_view_map.get(ax) if ax is not None else None
        if self._fixed_crop_transform_mode and view is not None:
            try:
                if self._fixed_crop_template_handle_hit(event, view, ax) is not None:
                    return True
            except Exception:
                pass
        if self._profile_hit_test(event):
            return True
        if self._angle_hit_test(event):
            return True
        return self._molecule_overlay_hit_test(event)

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
        self.push_undo_state("move_angle")
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
        if not (
            getattr(self, "profile_enabled", False)
            or getattr(self, "_profile_move_only", False)
            or bool(self._saved_profiles)
        ):
            self._emit_profile_state()
            return
        active = None
        if self.profile_pts is not None:
            active = self._build_profile_data(
                self.profile_pts,
                color=self._active_profile_color,
                view=self.views[0] if self.views else None,
                lw=self._active_profile_lw,
                line_style=self._active_profile_line_style,
                marker_style=self._active_profile_marker_style,
                marker_size=self._active_profile_marker_size,
                live_profile_ref=self._profile_live_ref(None),
            )
            if isinstance(active, dict) and self._active_profile_original_id:
                active["profile_id"] = str(self._active_profile_original_id)
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
                    p = self._build_profile_data(
                        self.profile_pts,
                        color=col,
                        view=v,
                        lw=self._active_profile_lw,
                        line_style=self._active_profile_line_style,
                        marker_style=self._active_profile_marker_style,
                        marker_size=self._active_profile_marker_size,
                        live_profile_ref=self._profile_live_ref(None),
                    )
                    if p:
                        name = v.get('colorbar_label') or v.get('title') or f"Ch{i+2}"
                        p['name'] = name
                        extras.append(p)
                active['extra_channels'] = extras

        saved_data = []
        for entry in self._saved_profiles:
            data = entry.get('data')
            if data is None:
                data = self._build_profile_data(
                    entry.get('pts'),
                    color=entry.get('color'),
                    lw=entry.get('lw'),
                    line_style=entry.get('line_style'),
                    marker_style=entry.get('marker_style'),
                    marker_size=entry.get('marker_size'),
                    live_profile_ref=self._profile_live_ref(entry=entry),
                )
                if isinstance(data, dict):
                    data["profile_id"] = str(self._ensure_saved_profile_id(entry))
                entry['data'] = data
            elif isinstance(data, dict):
                data['live_profile_ref'] = self._profile_live_ref(entry=entry)
                data["profile_id"] = str(self._ensure_saved_profile_id(entry))
            if data:
                saved_data.append(data)
        try:
            notify_profile_source_changed(self, active, saved_data)
        except Exception:
            pass
        if callable(self.profile_callback):
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

    def _event_qt_modifiers(self, event):
        mods_qt = QtCore.Qt.NoModifier
        try:
            mods_qt = getattr(getattr(event, "guiEvent", None), "modifiers", lambda: QtCore.Qt.NoModifier)()
        except Exception:
            mods_qt = QtCore.Qt.NoModifier
        if mods_qt == QtCore.Qt.NoModifier:
            try:
                mods_qt = QtWidgets.QApplication.keyboardModifiers()
            except Exception:
                mods_qt = QtCore.Qt.NoModifier
        return mods_qt

    def _on_base_click(self, event):
        if event is None or event.inaxes is None:
            return
        if self._shortcut_hint_hit(event):
            self.set_show_shortcut_hint(False)
            return
        # If clicking on a scale bar, do not trigger base canvas actions (like drag/copy)
        if self.scale_bar_enabled:
            for sb in self._scale_bar_artists:
                if sb.contains(event)[0]:
                    return
        if self.scale_bar_enabled and self._scale_bar_drag_start is not None:
            return
        gizmo_hit = self._molecule_gizmo_hit_test(event)
        if gizmo_hit is not None:
            self._begin_molecule_gizmo_drag(gizmo_hit, event)
            return
        ax = event.inaxes
        view = self._ax_view_map.get(ax)
        if ax is getattr(self, "_molecule_gizmo_axes", None):
            return
        if self._fixed_crop_transform_mode and event.button == 1 and view is not None:
            mods_qt = self._event_qt_modifiers(event)
            hit = self._fixed_crop_template_handle_hit(event, view, ax)
            if hit is not None:
                if hit.get("mode") == "move" and bool(mods_qt & QtCore.Qt.ControlModifier):
                    hit = dict(hit)
                    hit["mode"] = "rotate"
                if self._begin_fixed_crop_template_drag(hit, event, view, ax):
                    return

        if self._check_molecule_hit(event):
            return
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
        #   Shift + drag while crop-template square mode is active -> square selection
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
        template_square = bool((self._fixed_crop_template or {}).get("square", False))
        want_square = (
            (event.button == 1)
            and bool(mods_qt & QtCore.Qt.ShiftModifier)
            and (
                bool(mods_qt & QtCore.Qt.ControlModifier)
                or (self._fixed_crop_quick_mode and template_square)
            )
        )
        want_rect = (event.button == 1) and (mods_qt & QtCore.Qt.ShiftModifier)
        quick_profile = (
            event.button == 1
            and ax is self.main_ax
            and bool(getattr(self, "_measurement_shortcuts_enabled", True))
            and bool(mods_qt & QtCore.Qt.ControlModifier)
            and not bool(mods_qt & QtCore.Qt.AltModifier)
            and not bool(mods_qt & QtCore.Qt.ShiftModifier)
        )
        quick_angle = (
            event.button == 1
            and ax is self.main_ax
            and bool(getattr(self, "_measurement_shortcuts_enabled", True))
            and bool(mods_qt & QtCore.Qt.ControlModifier)
            and bool(mods_qt & QtCore.Qt.AltModifier)
            and not bool(mods_qt & QtCore.Qt.ShiftModifier)
        )
        if quick_angle:
            try:
                if self.profile_enabled:
                    self.set_profile_tool_enabled(False)
                if not self.angle_enabled:
                    self.set_angle_tool_enabled(True)
                if event.xdata is not None and event.ydata is not None:
                    self._add_angle_frame_at(event.xdata, event.ydata)
            except Exception:
                pass
            return
        if quick_profile:
            try:
                if self.angle_enabled:
                    self.set_angle_tool_enabled(False)
                was_enabled = bool(self.profile_enabled)
                if self._profile_move_only:
                    self._profile_move_only = False
                    self._profile_user_enabled = True
                    self._profile_quick_transient = True
                    was_enabled = False
                if not self.profile_enabled:
                    self.set_profile_tool_enabled(True)
                    self._profile_quick_transient = True
                if not was_enabled:
                    # Start the first profile drag immediately when the tool
                    # is activated via Ctrl+Click.
                    self._on_press(event)
            except Exception:
                pass
            return
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
        overlay_hit = self._interactive_overlay_hit_test(event)
        # Pan start: left/middle drag on the image background when zoomed.
        # Overlay tools only reserve their handles/bodies, so missed clicks still pan.
        if event.button in (1, 2) and not want_rect and view is not None and not overlay_hit:
            if event.xdata is not None and event.ydata is not None and self._is_zoomed(ax):
                self._pan_active = True
                self._pan_ax = ax
                self._pan_start = (event.xdata, event.ydata)
                self._pan_start_lim = (ax.get_xlim(), ax.get_ylim())
                return
        if self._fixed_crop_transform_mode and event.button == 1 and view is not None:
            return
        if (self.profile_enabled or self.angle_enabled) and ax is self.main_ax:
            if event.button == 3:
                self._show_context_menu(event, view)
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
        if self.profile_enabled and ax is self.main_ax and overlay_hit:
            # avoid starting thumbnail drag/other actions while measuring profiles
            if event.button == 3:
                self._show_context_menu(event, view)
            return
        if self.angle_enabled and ax is self.main_ax and overlay_hit:
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

    def _load_molecule_default_style(self):
        try:
            cfg = load_config()
            stored = cfg.get("molecule_default_style")
        except Exception:
            stored = None
        style = dict(_DEFAULT_MOLECULE_STYLE)
        if isinstance(stored, dict):
            style.update({str(k): v for k, v in stored.items()})
        style["render_style"] = normalize_molecule_render_style(style.get("render_style", "licorice"))
        style["display_mode"] = str(style.get("display_mode") or "Bonds Only")
        style["bond_style"] = str(style.get("bond_style") or "thin").lower()
        style["radius_mode"] = str(style.get("radius_mode") or "vdw").lower()
        try:
            style["radius_scale"] = float(style.get("radius_scale", 1.0))
        except Exception:
            style["radius_scale"] = 1.0
        style["palette"] = str(style.get("palette") or "avogadro").lower()
        style["show_shadows"] = bool(style.get("show_shadows", False))
        style["show_hydrogens"] = bool(style.get("show_hydrogens", False))
        return style

    def _molecule_style_from_molecule(self, mol):
        style = dict(_DEFAULT_MOLECULE_STYLE)
        if mol is not None:
            style.update(
                {
                    "display_mode": getattr(mol, "display_mode", style["display_mode"]),
                    "render_style": normalize_molecule_render_style(getattr(mol, "render_style", style["render_style"])),
                    "bond_style": str(getattr(mol, "bond_style", style["bond_style"]) or style["bond_style"]).lower(),
                    "radius_mode": str(getattr(mol, "radius_mode", style["radius_mode"]) or style["radius_mode"]).lower(),
                    "radius_scale": float(getattr(mol, "radius_scale", style["radius_scale"]) or style["radius_scale"]),
                    "atom_color_override": getattr(mol, "atom_color_override", None),
                    "bond_color_override": getattr(mol, "bond_color_override", None),
                    "bond_color_mode": getattr(mol, "bond_color_mode", "default"),
                    "atom_color_map": dict(getattr(mol, "atom_color_map", {}) or {}),
                }
            )
        style["palette"] = str(getattr(self, "molecule_palette", style["palette"]) or style["palette"]).lower()
        style["show_shadows"] = bool(getattr(self, "_show_molecule_shadow", style["show_shadows"]))
        style["show_hydrogens"] = bool(getattr(self, "_show_hydrogens", style["show_hydrogens"]))
        return style

    def _save_molecule_default_style(self, style=None):
        style = dict(style or {})
        if not style:
            return False
        normalized = dict(_DEFAULT_MOLECULE_STYLE)
        normalized.update(style)
        normalized["render_style"] = normalize_molecule_render_style(normalized.get("render_style", "licorice"))
        normalized["bond_style"] = str(normalized.get("bond_style") or "thin").lower()
        normalized["radius_mode"] = str(normalized.get("radius_mode") or "vdw").lower()
        normalized["palette"] = str(normalized.get("palette") or "avogadro").lower()
        normalized["show_shadows"] = bool(normalized.get("show_shadows", False))
        normalized["show_hydrogens"] = bool(normalized.get("show_hydrogens", False))
        try:
            normalized["radius_scale"] = float(normalized.get("radius_scale", 1.0))
        except Exception:
            normalized["radius_scale"] = 1.0
        try:
            cfg = load_config()
            cfg["molecule_default_style"] = normalized
            cfg["molecule_palette"] = normalized["palette"]
            save_config(cfg)
        except Exception:
            return False
        self._show_molecule_shadow = normalized["show_shadows"]
        self._show_hydrogens = normalized["show_hydrogens"]
        self.set_molecule_palette(normalized["palette"], notify=True)
        return True

    def _apply_molecule_default_style(self, mol):
        if mol is None:
            return
        style = self._load_molecule_default_style()
        mol.display_mode = style.get("display_mode", "Bonds Only")
        mol.render_style = normalize_molecule_render_style(style.get("render_style", "licorice"))
        mol.bond_style = str(style.get("bond_style", "thin") or "thin").lower()
        mol.radius_mode = str(style.get("radius_mode", "vdw") or "vdw").lower()
        try:
            mol.radius_scale = float(style.get("radius_scale", 1.0))
        except Exception:
            mol.radius_scale = 1.0
        mol.atom_color_override = style.get("atom_color_override")
        mol.bond_color_override = style.get("bond_color_override")
        mol.bond_color_mode = style.get("bond_color_mode", "default")
        mol.atom_color_map = dict(style.get("atom_color_map") or {})
        self._show_molecule_shadow = bool(style.get("show_shadows", False))
        self._show_hydrogens = bool(style.get("show_hydrogens", False))
        self.molecule_palette = str(style.get("palette", "avogadro") or "avogadro").lower()

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
            self._apply_molecule_default_style(mol)
            # Center in current view if possible
            if self.main_ax:
                xlim = self.main_ax.get_xlim()
                ylim = self.main_ax.get_ylim()
                mol.offset = np.array([(xlim[0]+xlim[1])/2, (ylim[0]+ylim[1])/2, 0.0])
            self.molecules.append(mol)
            self._active_molecule_idx = len(self.molecules) - 1
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
        self._active_molecule_idx = None
        self._molecule_artists = []
        self._redraw()

    def reset_molecules(self):
        """Public helper to clear all molecules with undo support."""
        self._clear_molecules()

    def _push_molecule_snapshot(self):
        """Save current molecule state for undo."""
        self.push_undo_state("molecules")
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
            return False
        try:
            last = self._molecule_history.pop()
            self.molecules = [m.copy() for m in last]
            self._redraw()
            return True
        except Exception:
            return False

    def _check_molecule_hit(self, event):
        # Angle editing has exclusive ownership of the canvas.
        if self.angle_enabled:
            return False
        if not self.show_molecules or not self.molecules or event.inaxes is None:
            return False
        # If the pointer is actually on a profile, let the profile tool win.
        if self._profile_hit_test(event):
            return False
        
        # Hit test in screen pixels so molecule editing does not steal large
        # background regions when the scan is zoomed or uses physical units.
        # Iterate in reverse to pick top-most
        for idx, mol in reversed(list(enumerate(self.molecules))):
            coords = mol.get_transformed_coordinates()
            if len(coords) == 0: continue
            if not getattr(self, "_show_hydrogens", True):
                mask = [str(el).strip().upper() != 'H' for el in mol.elements]
                if not any(mask):
                    continue
                coords = coords[np.array(mask)]

            try:
                atom_px = event.inaxes.transData.transform(coords[:, :2])
                event_px = np.array([float(event.x), float(event.y)], dtype=float)
                dist_px = np.hypot(atom_px[:, 0] - event_px[0], atom_px[:, 1] - event_px[1])
                hit = bool(dist_px.size and float(np.nanmin(dist_px)) <= 14.0)
            except Exception:
                dx = coords[:, 0] - event.xdata
                dy = coords[:, 1] - event.ydata
                hit = bool(np.min(dx * dx + dy * dy) < 0.25)

            if hit:
                self._active_molecule_idx = idx
                self._wake_molecule_gizmo(2200, redraw=False)
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
                    self._update_molecule_gizmo_overlay()
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
        gizmo_drag = getattr(self, "_molecule_gizmo_drag", None)
        if gizmo_drag is not None:
            idx = gizmo_drag.get("idx")
            if idx is None or idx < 0 or idx >= len(self.molecules):
                self._molecule_gizmo_drag = None
                return
            if getattr(event, "x", None) is None or getattr(event, "y", None) is None:
                return
            mol = self.molecules[idx]
            start_angles = gizmo_drag.get("start_angles")
            if start_angles is None:
                start_angles = mol.angles
            new_angles = np.array(start_angles, dtype=float, copy=True)
            if gizmo_drag.get("mode") == "rotate_z":
                gizmo_ax = getattr(self, "_molecule_gizmo_axes", None)
                bbox = getattr(gizmo_ax, "bbox", None)
                if bbox is None:
                    return
                center_x = float(bbox.x0 + (bbox.width * 0.5))
                center_y = float(bbox.y0 + (bbox.height * 0.5))
                start_local = gizmo_drag.get("start_local") or (0.0, 0.0)
                start_angle = math.degrees(math.atan2(float(start_local[1]), float(start_local[0])))
                current_angle = math.degrees(math.atan2(float(event.y) - center_y, float(event.x) - center_x))
                new_angles[2] += current_angle - start_angle
            else:
                dx_px = float(event.x) - float(gizmo_drag["start_px"][0])
                dy_px = float(event.y) - float(gizmo_drag["start_px"][1])
                sensitivity = 0.45
                new_angles[0] += dy_px * sensitivity
                new_angles[1] += dx_px * sensitivity
            mol.angles = new_angles
            self._wake_molecule_gizmo(2400, redraw=False)
            self._update_molecule_artists()
            self._update_molecule_gizmo_overlay()
            return
        if self._molecule_drag_idx is not None:
            if event.xdata is None or event.ydata is None:
                return
            mol = self.molecules[self._molecule_drag_idx]
            self._wake_molecule_gizmo(2200, redraw=False)
            
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
            self._update_molecule_gizmo_overlay()

    def _on_molecule_release(self, event):
        if self._molecule_gizmo_drag is not None:
            self._molecule_gizmo_drag = None
            return
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
        palette = (palette or "avogadro").lower()
        if palette not in available_atom_palettes():
            palette = "avogadro"
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

    def apply_view_colormap(self, cmap_name: str, *, target_view=None, notify: bool = True):
        """Apply a colormap to one or all current views and redraw in place."""
        cmap_name = str(cmap_name or "").strip()
        if not cmap_name:
            return False
        targets = []
        for current_view in list(self.views or []):
            if target_view is not None and current_view is not target_view:
                continue
            if str(current_view.get("cmap") or "") == cmap_name:
                continue
            targets.append(current_view)
        if not targets:
            return False
        self.push_undo_state("colormap")
        changed = False
        for current_view in targets:
            current_view["cmap"] = cmap_name
            changed = True
        self._redraw()
        if notify:
            self._notify_views_callback()
        return True

    def _show_molecule_menu(self, event, mol):
        style = QtWidgets.QApplication.style()
        icon = lambda std: style.standardIcon(std) if style else QtGui.QIcon()

        menu = QtWidgets.QMenu(self)

        # Main edit entry point
        props_act = menu.addAction(icon(QtWidgets.QStyle.SP_FileDialogDetailedView), "Edit molecule...")
        menu.addSeparator()

        # Quick view toggles
        toggle_shadow_act = menu.addAction(icon(QtWidgets.QStyle.SP_DialogYesButton), "Show shadows")
        toggle_shadow_act.setCheckable(True)
        toggle_shadow_act.setChecked(self._show_molecule_shadow)
        show_h_act = menu.addAction(icon(QtWidgets.QStyle.SP_TitleBarShadeButton), "Show hydrogens")
        show_h_act.setCheckable(True)
        show_h_act.setChecked(getattr(self, "_show_hydrogens", True))
        menu.addSeparator()

        # Reset/state
        reset_file_act = menu.addAction(icon(QtWidgets.QStyle.SP_BrowserReload), "Reset to file state")
        reset_file_act.setShortcut(QtGui.QKeySequence("Shift+R"))
        reset_file_act.setEnabled(bool(getattr(mol, "filepath", None)))
        reset_all_act = menu.addAction(icon(QtWidgets.QStyle.SP_MessageBoxWarning), "Reset all molecules")

        undo_act = menu.addAction(icon(QtWidgets.QStyle.SP_ArrowBack), "Undo last change")
        undo_act.setShortcut(QtGui.QKeySequence("Ctrl+Z"))

        # Palette submenu
        pal_menu = menu.addMenu(icon(QtWidgets.QStyle.SP_DialogHelpButton), "Atom palette")
        current_pal = (getattr(self, "molecule_palette", "avogadro") or "avogadro").lower()
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
            overlay_settings = {
                "palette": getattr(self, "molecule_palette", "avogadro"),
                "show_shadows": bool(self._show_molecule_shadow),
                "show_hydrogens": bool(getattr(self, "_show_hydrogens", True)),
                "show_shadows_available": True,
                "show_hydrogens_available": True,
                "palette_available": True,
                "save_default_callback": lambda style, c=self: c._save_molecule_default_style(style),
            }
            def _apply():
                self._show_molecule_shadow = bool(overlay_settings.get("show_shadows", True))
                self._show_hydrogens = bool(overlay_settings.get("show_hydrogens", True))
                self.set_molecule_palette(overlay_settings.get("palette", "avogadro"), notify=False)
                self._redraw()
            dlg = MoleculePropertiesDialog(mol, self, callback=_apply, overlay_settings=overlay_settings)
            dlg.show()
        elif action == reset_file_act:
            idx = None
            try:
                idx = self.molecules.index(mol)
            except Exception:
                idx = getattr(self, "_active_molecule_idx", None)
            self._reset_molecule_to_file_state(idx)
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
            self._notify_copy_feedback(view, fmt="png", displayed=False)
        except Exception:
            pass

    def _view_to_qimage(self, view):
        arr = np.asarray(view.get('arr'))
        cmap = view.get('cmap', 'viridis')
        return array_to_qimage(arr, cmap_name=cmap)

    @classmethod
    def _stash_drag_view_snapshot(cls, view):
        if not isinstance(view, dict):
            return None
        try:
            snapshot = dict(view)
        except Exception:
            return None
        arr = snapshot.get("arr")
        if arr is not None:
            try:
                snapshot["arr"] = np.array(arr, copy=True)
            except Exception:
                pass
        meta = snapshot.get("meta")
        if isinstance(meta, dict):
            try:
                snapshot["meta"] = dict(meta)
            except Exception:
                pass
        token = f"viewdrag_{time.time_ns()}"
        cls._DRAG_VIEW_SNAPSHOTS[token] = {
            "view": snapshot,
            "ts_ns": time.time_ns(),
        }
        while len(cls._DRAG_VIEW_SNAPSHOTS) > cls._DRAG_VIEW_SNAPSHOT_LIMIT:
            try:
                oldest = next(iter(cls._DRAG_VIEW_SNAPSHOTS))
            except Exception:
                break
            cls._DRAG_VIEW_SNAPSHOTS.pop(oldest, None)
        return token

    @classmethod
    def consume_drag_view_snapshot(cls, token):
        if not token:
            return None
        entry = cls._DRAG_VIEW_SNAPSHOTS.pop(str(token), None)
        if not isinstance(entry, dict):
            return None
        view = entry.get("view")
        if not isinstance(view, dict):
            return None
        result = dict(view)
        arr = result.get("arr")
        if arr is not None:
            try:
                result["arr"] = np.array(arr, copy=True)
            except Exception:
                pass
        meta = result.get("meta")
        if isinstance(meta, dict):
            try:
                result["meta"] = dict(meta)
            except Exception:
                pass
        return result

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
        # Quick tools (no top toolbar required)
        quick_menu = menu.addMenu("Quick tools")
        profile_tool_act = quick_menu.addAction("Profile tool  (Ctrl+Click)")
        profile_tool_act.setCheckable(True)
        profile_tool_act.setChecked(bool(self.profile_enabled))
        angle_tool_act = quick_menu.addAction("Angle tool  (Ctrl+Alt+Click)")
        angle_tool_act.setCheckable(True)
        angle_tool_act.setChecked(bool(self.angle_enabled))
        if not bool(getattr(self, "_measurement_shortcuts_enabled", True)):
            profile_tool_act.setEnabled(False)
            angle_tool_act.setEnabled(False)
        quick_menu.addSeparator()
        edit_crop_frame_act = quick_menu.addAction("Edit crop template")
        edit_crop_frame_act.setCheckable(True)
        edit_crop_frame_act.setChecked(bool(self._fixed_crop_transform_mode))
        apply_crop_frame_act = quick_menu.addAction("Apply crop template  (Enter)")
        apply_crop_frame_act.setEnabled(bool(self._fixed_crop_template_visible and self._fixed_crop_template))
        exit_crop_frame_act = quick_menu.addAction("Exit template editor")
        exit_crop_frame_act.setEnabled(bool(self._fixed_crop_transform_mode))
        quick_menu.addSeparator()
        clear_overlays_act = quick_menu.addAction("Clear profile/angle overlays")
        auto_hist_act = quick_menu.addAction("Auto contrast (1-99%)  (A)")
        auto_hist_act.setEnabled(callable(self._histogram_auto_callback))
        reset_hist_act = quick_menu.addAction("Reset range to data min/max")
        reset_hist_act.setEnabled(callable(self._histogram_reset_callback))
        histogram_act = quick_menu.addAction("Histogram...")
        histogram_act.setEnabled(callable(self._histogram_dialog_callback))

        display_menu = menu.addMenu("Display")
        presets_menu = display_menu.addMenu("Preset")
        preset_focus_act = presets_menu.addAction("Focus")
        preset_analysis_act = presets_menu.addAction("Analysis")
        preset_publication_act = presets_menu.addAction("Publication")
        display_menu.addSeparator()
        show_scale_act = display_menu.addAction("Show Scale bar")
        show_scale_act.setCheckable(True)
        show_scale_act.setChecked(bool(self.scale_bar_enabled))
        show_ticks_act = display_menu.addAction("Show Ticks")
        show_ticks_act.setCheckable(True)
        show_ticks_act.setChecked(bool(self._show_ticks))
        show_cbar_act = display_menu.addAction("Show Colorbar")
        show_cbar_act.setCheckable(True)
        show_cbar_act.setChecked(bool(self._show_colorbar))
        rel_zero_act = None
        if callable(self._display_relative_zero_menu_callback):
            rel_zero_act = display_menu.addAction("Values relative to zero/reference")
            rel_zero_act.setCheckable(True)
            try:
                rel_zero_act.setChecked(bool(self._display_relative_zero_menu_state_callback()))
            except Exception:
                rel_zero_act.setChecked(False)
            rel_zero_tip = self._display_relative_zero_menu_tooltip or "Display values relative to the current zero/reference"
            rel_zero_act.setToolTip(rel_zero_tip)
            rel_zero_act.setStatusTip(rel_zero_tip)
        cbar_orient_menu = display_menu.addMenu("Colorbar orientation")
        cbar_orient_group = QtWidgets.QActionGroup(self)
        cbar_orient_group.setExclusive(True)
        cbar_vert_act = cbar_orient_menu.addAction("Vertical")
        cbar_vert_act.setCheckable(True)
        cbar_vert_act.setChecked(self._colorbar_orientation == 'vertical')
        cbar_orient_group.addAction(cbar_vert_act)
        cbar_horiz_act = cbar_orient_menu.addAction("Horizontal")
        cbar_horiz_act.setCheckable(True)
        cbar_horiz_act.setChecked(self._colorbar_orientation == 'horizontal')
        cbar_orient_group.addAction(cbar_horiz_act)
        show_title_act = display_menu.addAction("Show Title")
        show_title_act.setCheckable(True)
        show_title_act.setChecked(bool(self._show_title))
        show_profiles_act = display_menu.addAction("Show Profiles")
        show_profiles_act.setCheckable(True)
        show_profiles_act.setChecked(bool(self._show_profile_overlays))
        show_crop_history_act = None
        if callable(self._apply_popup_style_callback) and (self._fixed_crop_quick_mode or self._fixed_crop_history):
            show_crop_history_act = display_menu.addAction("Show Crop Overlays")
            show_crop_history_act.setCheckable(True)
            show_crop_history_act.setChecked(bool(self._fixed_crop_history_visible))
            crop_tip = "Show or hide crop-template overlays in this pop-up only."
            show_crop_history_act.setToolTip(crop_tip)
            show_crop_history_act.setStatusTip(crop_tip)
        acq_overlay_act = display_menu.addAction("Show Acquisition HUD")
        acq_overlay_act.setCheckable(True)
        acq_overlay_act.setChecked(bool(self._show_acquisition_overlay))
        hint_act = display_menu.addAction("Show Shortcut Hint")
        hint_act.setCheckable(True)
        hint_act.setChecked(bool(self._show_shortcut_hint))
        gizmo_act = display_menu.addAction("Show Molecule Gizmo")
        gizmo_act.setCheckable(True)
        gizmo_act.setChecked(bool(getattr(self, "_show_molecule_gizmo", False)))
        frame_fill_act = display_menu.addAction("Frame fill")
        frame_fill_act.setCheckable(True)
        frame_fill_act.setChecked(bool(self._frame_fill_mode))
        rel_axes_act = display_menu.addAction("Relative axes")
        rel_axes_act.setCheckable(True)
        rel_axes_act.setChecked(bool(self._use_relative_axes(view)))
        apply_popup_style_act = None
        if callable(self._apply_popup_style_callback):
            display_menu.addSeparator()
            apply_popup_style_act = display_menu.addAction(self._apply_popup_style_label or "Apply this style to all pop-ups")
            popup_style_tip = self._apply_popup_style_tooltip or "Copy font size, typography and display layout from this popup to the other open pop-ups"
            apply_popup_style_act.setToolTip(popup_style_tip)
            apply_popup_style_act.setStatusTip(popup_style_tip)

        layout_menu = display_menu.addMenu("Layout")
        layout_grid_act = layout_menu.addAction("Grid")
        layout_grid_act.setCheckable(True)
        layout_grid_act.setChecked(self._view_layout == "grid")
        layout_stack_act = layout_menu.addAction("Stacked")
        layout_stack_act.setCheckable(True)
        layout_stack_act.setChecked(self._view_layout == "stacked")

        overlays_menu = menu.addMenu("Overlays")
        show_profile_overlay_act = overlays_menu.addAction("Show Saved Profiles  (Ctrl+1)")
        show_profile_overlay_act.setCheckable(True)
        show_profile_overlay_act.setChecked(bool(self._show_profile_overlays))
        show_angle_overlay_act = overlays_menu.addAction("Show Saved Angles  (Ctrl+2)")
        show_angle_overlay_act.setCheckable(True)
        show_angle_overlay_act.setChecked(bool(self._show_angle_overlays))
        show_molecule_overlay_act = overlays_menu.addAction("Show Molecules  (Ctrl+3)")
        show_molecule_overlay_act.setCheckable(True)
        show_molecule_overlay_act.setChecked(bool(self.show_molecules))

        analysis_menu = menu.addMenu("Analysis")
        # Filters (provided by parent viewer)
        if callable(self._filter_menu_callback):
            try:
                self._filter_menu_callback(analysis_menu, view, self)
            except Exception:
                pass
        angle_style_act = None
        if self.angle_enabled and self._angle_frames:
            angle_style_act = analysis_menu.addAction("Use arrowheads for active angle")
            angle_style_act.setCheckable(True)
            active = self._get_active_angle_frame()
            angle_style_act.setChecked(active and active.get('style', 'dots') == 'arrows')

        copy_menu = menu.addMenu("Copy")
        copy_disp_png = copy_menu.addAction("Copy displayed as PNG  (Ctrl+C)")
        copy_disp_svg = copy_menu.addAction("Copy displayed as SVG")
        copy_menu.addSeparator()
        copy_act = copy_menu.addAction("Copy data image only (PNG)")
        copy_svg_act = copy_menu.addAction("Copy data view as SVG (vector)")

        send_ppt_act = menu.addAction("Send to PowerPoint")
        send_ppt_current_act = menu.addAction("Send to Current Slide")
        ppt_supported, ppt_reason = powerpoint_support_status()
        if not ppt_supported:
            send_ppt_act.setEnabled(False)
            send_ppt_current_act.setEnabled(False)
            send_ppt_act.setToolTip(ppt_reason or "")
            send_ppt_current_act.setToolTip(ppt_reason or "")

        export_menu = menu.addMenu("Save / Export")
        save_act = export_menu.addAction("Save data image as PNG...")
        save_svg_act = export_menu.addAction("Save displayed view as SVG...")
        save_pdf_act = export_menu.addAction("Save displayed view as PDF...")
        export_menu.addSeparator()
        export_stp_act = export_menu.addAction("Export as WSxM STP...")

        virtual_copy_act = menu.addAction("Create virtual copy in thumbnails")
        virtual_copy_act.setEnabled(bool(callable(self._virtual_copy_callback) and view.get("arr") is not None))

        molecules_menu = menu.addMenu("Molecules")
        load_mol_act = molecules_menu.addAction("Load Molecule (XYZ/PDB)...")
        recent_menu = None
        recent_actions = {}
        recent_all = []
        for lst in (self._recent_molecule_paths, MultiPreviewCanvas._RECENT_MOLECULES):
            for p in lst:
                if p not in recent_all:
                    recent_all.append(p)
        if recent_all:
            recent_menu = molecules_menu.addMenu("Load Recent")
            for p in recent_all[:8]:
                act = recent_menu.addAction(Path(p).name)
                act.setToolTip(p)
                recent_actions[act] = p
        clear_mols_act = molecules_menu.addAction("Clear Molecules")

        cmap_menu = menu.addMenu("Colormap")
        popup_cmap_apply_all_act = None
        cmap_actions = {}
        cmap_group = QtWidgets.QActionGroup(self)
        cmap_group.setExclusive(True)
        common_cmaps = [
            "viridis",
            "plasma",
            "inferno",
            "magma",
            "cividis",
            "turbo",
            "gray",
            "afmhot",
            "Blues_r",
            "RdBu_r",
            "coolwarm",
        ]
        try:
            available_cmaps = sorted(str(name) for name in matplotlib.colormaps.keys())
        except Exception:
            available_cmaps = list(common_cmaps)
        seen_cmaps = []
        for cmap_name in common_cmaps + available_cmaps:
            if cmap_name not in seen_cmaps:
                seen_cmaps.append(cmap_name)
        current_cmap = str((view or {}).get("cmap") or "viridis")
        more_cmaps_menu = None
        for idx, cmap_name in enumerate(seen_cmaps):
            parent_menu = cmap_menu if idx < 12 else more_cmaps_menu
            if parent_menu is None:
                more_cmaps_menu = cmap_menu.addMenu("More...")
                parent_menu = more_cmaps_menu
            act = parent_menu.addAction(cmap_name)
            act.setCheckable(True)
            act.setChecked(cmap_name == current_cmap)
            try:
                act.setIcon(_colormap_icon(cmap_name, width=96, height=14))
            except Exception:
                pass
            cmap_group.addAction(act)
            cmap_actions[act] = cmap_name
        if callable(self._apply_popup_style_callback):
            cmap_menu.addSeparator()
            popup_cmap_apply_all_act = cmap_menu.addAction("Apply this colormap to all pop-ups")

        collection_add_act = None
        collection_remove_act = None
        collection_help_act = None
        if callable(self._collection_menu_callback) and view is not None:
            collection_menu = menu.addMenu("Collection")
            collection_add_act = collection_menu.addAction("Add This View to Collection...")
            collection_add_act.setToolTip("Save this view into a curated cross-folder collection.")
            collection_remove_act = collection_menu.addAction("Remove This View from Collection")
            collection_remove_act.setToolTip("Remove the matching item from the current collection file.")
            if callable(self._collection_help_callback):
                collection_menu.addSeparator()
                collection_help_act = collection_menu.addAction("How Collections Work")

        source_meta = dict((view or {}).get("meta") or {})
        source_path = str((view or {}).get("path") or source_meta.get("path") or source_meta.get("file_path") or "").strip()
        add_source_file_menu(menu, source_path, self)

        compare_set_a_act = None
        compare_set_b_act = None
        compare_with_a_act = None
        compare_with_b_act = None
        compare_open_act = None
        compare_swap_act = None
        compare_clear_act = None
        if callable(self._compare_menu_callback) and view is not None:
            compare_state = {}
            if callable(self._compare_menu_state_callback):
                try:
                    compare_state = dict(self._compare_menu_state_callback() or {})
                except Exception:
                    compare_state = {}
            compare_menu = menu.addMenu("Compare")
            label_a = str(compare_state.get("label_a") or "Empty")
            label_b = str(compare_state.get("label_b") or "Empty")
            compare_menu.setToolTip(f"A: {label_a}\nB: {label_b}")
            compare_set_a_act = compare_menu.addAction("Set This View as Compare A")
            compare_set_a_act.setToolTip(f"Current A: {label_a}")
            compare_set_b_act = compare_menu.addAction("Set This View as Compare B")
            compare_set_b_act.setToolTip(f"Current B: {label_b}")
            compare_menu.addSeparator()
            compare_with_a_act = compare_menu.addAction("Compare A with This")
            compare_with_a_act.setEnabled(bool(compare_state.get("has_a")))
            compare_with_b_act = compare_menu.addAction("Compare B with This")
            compare_with_b_act.setEnabled(bool(compare_state.get("has_b")))
            compare_open_act = compare_menu.addAction("Open A/B Comparison")
            compare_open_act.setEnabled(bool(compare_state.get("has_a")) and bool(compare_state.get("has_b")))
            compare_swap_act = compare_menu.addAction("Swap A and B")
            compare_swap_act.setEnabled(bool(compare_state.get("has_a")) and bool(compare_state.get("has_b")))
            compare_clear_act = compare_menu.addAction("Clear Compare Selection")
            compare_clear_act.setEnabled(bool(compare_state.get("has_a")) or bool(compare_state.get("has_b")))

        view_menu = menu.addMenu("View")
        reset_zoom_act = view_menu.addAction("Reset Zoom")
        arrange_act = None
        minimize_act = None
        restore_act = None
        close_all_act = None
        window_actions_added = False
        if callable(self._arrange_windows_callback):
            view_menu.addSeparator()
            window_actions_added = True
            arrange_act = view_menu.addAction("Arrange pop-outs")
        if callable(self._minimize_windows_callback):
            if not window_actions_added:
                view_menu.addSeparator()
                window_actions_added = True
            minimize_act = view_menu.addAction("Minimize pop-outs")
        if callable(self._restore_windows_callback):
            if not window_actions_added:
                view_menu.addSeparator()
                window_actions_added = True
            restore_act = view_menu.addAction("Bring pop-outs to front")
        if callable(self._close_windows_callback):
            if not window_actions_added:
                view_menu.addSeparator()
                window_actions_added = True
            close_all_act = view_menu.addAction("Close all pop-outs")

        add_font_menu_action(
            menu,
            self,
            self._font_family,
            self._apply_plot_font_family_choice,
            current_style=self._plot_style_state(),
            apply_style_callback=self.set_plot_typography,
        )

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
        elif chosen == send_ppt_act:
            self._send_displayed_to_powerpoint(view, new_slide=True)
        elif chosen == send_ppt_current_act:
            self._send_displayed_to_powerpoint(view, new_slide=False)
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
        elif chosen == virtual_copy_act:
            if callable(self._virtual_copy_callback):
                try:
                    self._virtual_copy_callback(view)
                except Exception:
                    pass
        elif chosen == reset_zoom_act:
            self._reset_view_zoom()
        elif chosen in (cbar_vert_act, cbar_horiz_act):
            self._set_colorbar_orientation('vertical' if chosen == cbar_vert_act else 'horizontal')
        elif chosen == profile_tool_act:
            self.set_profile_tool_enabled(profile_tool_act.isChecked())
        elif chosen == angle_tool_act:
            self.set_angle_tool_enabled(angle_tool_act.isChecked())
        elif chosen == edit_crop_frame_act:
            self.enable_fixed_crop_transform_mode(edit_crop_frame_act.isChecked())
        elif chosen == apply_crop_frame_act:
            self._on_apply_fixed_crop_shortcut()
        elif chosen == exit_crop_frame_act:
            self.enable_fixed_crop_transform_mode(False)
        elif chosen == clear_overlays_act:
            self.clear_measurement_overlays()
        elif chosen == auto_hist_act and callable(self._histogram_auto_callback):
            try:
                self._histogram_auto_callback(self)
            except Exception:
                pass
        elif chosen == reset_hist_act and callable(self._histogram_reset_callback):
            try:
                self._histogram_reset_callback(self)
            except Exception:
                pass
        elif chosen == histogram_act and callable(self._histogram_dialog_callback):
            try:
                self._histogram_dialog_callback(self)
            except Exception:
                pass
        elif chosen == preset_focus_act:
            self.apply_display_preset("focus")
        elif chosen == preset_analysis_act:
            self.apply_display_preset("analysis")
        elif chosen == preset_publication_act:
            self.apply_display_preset("publication")
        elif chosen == show_scale_act:
            self.enable_scale_bar(show_scale_act.isChecked())
            self._notify_views_callback()
        elif chosen == show_ticks_act:
            self._toggle_ticks()
        elif chosen == show_cbar_act:
            self._toggle_colorbar()
        elif rel_zero_act and chosen == rel_zero_act:
            try:
                self._display_relative_zero_menu_callback(rel_zero_act.isChecked())
            except Exception:
                pass
        elif chosen == show_title_act:
            self.set_show_title(show_title_act.isChecked())
        elif chosen == show_profiles_act:
            self.set_show_profile_overlays(show_profiles_act.isChecked())
        elif show_crop_history_act and chosen == show_crop_history_act:
            self.show_fixed_crop_history(show_crop_history_act.isChecked())
        elif chosen == acq_overlay_act:
            self.set_show_acquisition_overlay(acq_overlay_act.isChecked())
        elif chosen == hint_act:
            self.set_show_shortcut_hint(hint_act.isChecked())
        elif chosen == gizmo_act:
            self.set_show_molecule_gizmo(gizmo_act.isChecked())
        elif chosen == frame_fill_act:
            self.set_frame_fill_mode(frame_fill_act.isChecked())
            self._notify_views_callback()
        elif chosen == rel_axes_act:
            self.set_relative_axes_override(rel_axes_act.isChecked())
        elif apply_popup_style_act and chosen == apply_popup_style_act:
            try:
                self._apply_popup_style_callback()
            except Exception:
                pass
        elif chosen == layout_grid_act:
            self.set_view_layout("grid")
        elif chosen == layout_stack_act:
            self.set_view_layout("stacked")
        elif chosen == show_profile_overlay_act:
            self.set_show_profile_overlays(show_profile_overlay_act.isChecked())
        elif chosen == show_angle_overlay_act:
            self.set_show_angle_overlays(show_angle_overlay_act.isChecked())
        elif chosen == show_molecule_overlay_act:
            self.set_show_molecules(show_molecule_overlay_act.isChecked())
        elif chosen == load_mol_act:
            self._load_molecule_dialog()
        elif recent_menu and chosen in recent_actions:
            self.add_molecule(recent_actions[chosen])
        elif chosen == clear_mols_act:
            self._clear_molecules()
        elif chosen in cmap_actions:
            try:
                self.apply_view_colormap(cmap_actions[chosen], target_view=view, notify=True)
            except Exception:
                pass
        elif popup_cmap_apply_all_act and chosen == popup_cmap_apply_all_act:
            try:
                self._apply_popup_style_callback()
            except Exception:
                pass
        elif collection_add_act and chosen == collection_add_act:
            try:
                self._collection_menu_callback("collection_add", view, self)
            except Exception:
                pass
        elif collection_remove_act and chosen == collection_remove_act:
            try:
                self._collection_menu_callback("collection_remove", view, self)
            except Exception:
                pass
        elif collection_help_act and chosen == collection_help_act:
            try:
                self._collection_help_callback()
            except Exception:
                pass
        elif compare_set_a_act and chosen == compare_set_a_act:
            try:
                self._compare_menu_callback("set_a", view, self)
            except Exception:
                pass
        elif compare_set_b_act and chosen == compare_set_b_act:
            try:
                self._compare_menu_callback("set_b", view, self)
            except Exception:
                pass
        elif compare_with_a_act and chosen == compare_with_a_act:
            try:
                self._compare_menu_callback("compare_with_a", view, self)
            except Exception:
                pass
        elif compare_with_b_act and chosen == compare_with_b_act:
            try:
                self._compare_menu_callback("compare_with_b", view, self)
            except Exception:
                pass
        elif compare_open_act and chosen == compare_open_act:
            try:
                self._compare_menu_callback("open_compare", view, self)
            except Exception:
                pass
        elif compare_swap_act and chosen == compare_swap_act:
            try:
                self._compare_menu_callback("swap_compare", view, self)
            except Exception:
                pass
        elif compare_clear_act and chosen == compare_clear_act:
            try:
                self._compare_menu_callback("clear_compare", view, self)
            except Exception:
                pass
        elif arrange_act and chosen == arrange_act:
            try:
                self._arrange_windows_callback()
            except Exception:
                pass
        elif minimize_act and chosen == minimize_act:
            try:
                self._minimize_windows_callback()
            except Exception:
                pass
        elif restore_act and chosen == restore_act:
            try:
                self._restore_windows_callback()
            except Exception:
                pass
        elif close_all_act and chosen == close_all_act:
            try:
                self._close_windows_callback()
            except Exception:
                pass
        elif angle_style_act and chosen == angle_style_act:
            checked = angle_style_act.isChecked()
            self._update_active_angle_style('arrows' if checked else 'dots')

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
        def _save_without_hint():
            if fmt == 'svg':
                with matplotlib.rc_context({'svg.fonttype': 'none'}):
                    self.fig.savefig(buf, format='svg', bbox_inches='tight')
            else:
                self.fig.savefig(buf, format='png', dpi=300, bbox_inches='tight')
        if fmt == 'svg':
            self._save_current_figure_without_shortcut_hint(_save_without_hint)
            mime = QtCore.QMimeData()
            mime.setData("image/svg+xml", buf.getvalue())
            QtWidgets.QApplication.clipboard().setMimeData(mime)
            self._notify_copy_feedback(fmt="svg", displayed=True)
        else:
            self._save_current_figure_without_shortcut_hint(_save_without_hint)
            qimg = QtGui.QImage.fromData(buf.getvalue())
            QtWidgets.QApplication.clipboard().setImage(qimg)
            self._notify_copy_feedback(fmt="png", displayed=True)

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
            self._notify_copy_feedback(view, fmt="svg", displayed=False)
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
            self._refresh_scale_bars()
            self.draw_idle()

    def _toggle_colorbar_orientation(self):
        self._set_colorbar_orientation('horizontal' if self._colorbar_orientation == 'vertical' else 'vertical')

    def _set_colorbar_orientation(self, orientation):
        orientation = str(orientation or 'vertical').strip().lower()
        if orientation not in ('vertical', 'horizontal'):
            orientation = 'vertical'
        if self._colorbar_orientation == orientation:
            return
        self.push_undo_state("colorbar_orientation")
        self._colorbar_orientation = orientation
        self._redraw()
        self._notify_views_callback()

    def _toggle_ticks(self):
        self.push_undo_state("show_ticks")
        self._show_ticks = not self._show_ticks
        self._redraw()
        self._notify_views_callback()

    def _toggle_colorbar(self):
        self.push_undo_state("show_colorbar")
        self._show_colorbar = not self._show_colorbar
        self._redraw()
        self._notify_views_callback()

    def _on_scroll_zoom(self, event):
        """Mouse wheel zoom centered at cursor."""
        ax = getattr(event, 'inaxes', None)
        if ax is None:
            return
        if (
            getattr(self, "_pan_active", False)
            or getattr(self, "_dragging", None) is not None
            or getattr(self, "_saved_profile_drag", None) is not None
            or getattr(self, "_profile_marker_drag_idx", None) is not None
            or getattr(self, "_angle_dragging", None) is not None
            or getattr(self, "_molecule_drag_idx", None) is not None
            or getattr(self, "_molecule_gizmo_drag", None) is not None
            or getattr(self, "_scale_bar_drag_start", None) is not None
            or getattr(self, "_fixed_crop_template_drag", None) is not None
            or getattr(self, "_crop_start", None) is not None
            or getattr(self, "_outline_start", None) is not None
        ):
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
        new_xlim = (
            x0 - (x0 - xlim[0]) * scale,
            x0 + (xlim[1] - x0) * scale,
        )
        new_ylim = (
            y0 - (y0 - ylim[0]) * scale,
            y0 + (ylim[1] - y0) * scale,
        )
        base_xlim, base_ylim = self._zoom_reset_limits.get(ax, (xlim, ylim))
        new_xlim = self._clamp_limits(new_xlim, base_xlim)
        new_ylim = self._clamp_limits(new_ylim, base_ylim)
        ax.set_xlim(new_xlim)
        ax.set_ylim(new_ylim)
        self._refresh_scale_bars(ax=ax)
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
            self._refresh_scale_bars()
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
            apply_text_style(ax.title, family=self._font_family, **self._plot_style_state())
        self._draw_acquisition_overlay(ax, view)
        ax.tick_params(labelsize=8)
        for lbl in list(ax.get_xticklabels()) + list(ax.get_yticklabels()):
            apply_text_style(lbl, family=self._font_family, **self._plot_style_state())

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

        self._draw_image_size_overlay(ax, view)

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
                    apply_text_style(cbar.ax.yaxis.label, family=self._font_family, **self._plot_style_state())
                    for lbl in list(cbar.ax.get_xticklabels()) + list(cbar.ax.get_yticklabels()):
                        apply_text_style(lbl, family=self._font_family, **self._plot_style_state())
                except Exception:
                    pass
            try:
                self._draw_outlines(ax, view)
            except Exception:
                pass
            title = view.get('title', '') or view.get('label', '')
            if title and self._show_title:
                ax.set_title(title, fontsize=9 * font_scale, color=text_color)
                apply_text_style(ax.title, family=self._font_family, **self._plot_style_state())
            self._draw_acquisition_overlay(ax, view)
            self._draw_image_size_overlay(ax, view)
            for lbl in list(ax.get_xticklabels()) + list(ax.get_yticklabels()):
                apply_text_style(lbl, family=self._font_family, **self._plot_style_state())
        fig.tight_layout()
        return fig

    def _acquisition_overlay_text(self, view):
        if not view:
            return ""
        text = view.get("acquisition_overlay_text")
        if text:
            return str(text).strip()
        meta = view.get("meta") or {}
        text = meta.get("acquisition_overlay_text")
        if text:
            return str(text).strip()
        mode = str(meta.get("acquisition_mode") or "").strip().upper()
        if mode == "CH":
            z_nm = meta.get("acquisition_z_abs_nm")
            try:
                return f"CH  z_abs {float(z_nm):.3f} nm"
            except Exception:
                return ""
        if mode == "CC":
            parts = []
            bias = meta.get("acquisition_bias_text")
            setp = meta.get("acquisition_setpoint_text")
            if bias:
                parts.append(f"Bias {bias}")
            if setp:
                parts.append(f"Iset {setp}")
            if parts:
                return "CC  " + " | ".join(parts)
        return ""

    def _draw_acquisition_overlay(self, ax, view):
        if not self._show_acquisition_overlay or ax is None:
            return
        text = self._acquisition_overlay_text(view)
        if not text:
            return
        scale = max(0.6, min(2.5, getattr(self, "_view_font_scale", 1.0)))
        fontsize = max(7.0, 8.5 * scale)
        text_artist = ax.text(
            0.985,
            0.985,
            text,
            transform=ax.transAxes,
            ha="right",
            va="top",
            fontsize=fontsize,
            fontweight="semibold",
            color="#f5f7fb",
            bbox={
                "facecolor": "black",
                "alpha": 0.42,
                "edgecolor": "none",
                "boxstyle": "round,pad=0.22",
            },
            zorder=26,
        )
        try:
            apply_text_style(text_artist, family=self._font_family, **self._plot_style_state())
        except Exception:
            pass

    def _image_size_overlay_text(self, view):
        if not getattr(self, "_show_image_size_overlay", False) or not view:
            return ""
        extent = view.get("extent_raw")
        if extent is None:
            extent = view.get("extent")
        width = None
        height = None
        unit = str(view.get("axis_unit") or "").strip()
        if extent is not None:
            try:
                x0, x1, y1, y0 = extent
                width = abs(float(x1) - float(x0))
                height = abs(float(y0) - float(y1))
            except Exception:
                width = None
                height = None
        if width is None or height is None:
            try:
                arr = np.asarray(view.get("arr"))
                if arr.ndim >= 2:
                    height = float(arr.shape[0])
                    width = float(arr.shape[1])
                    unit = "px"
            except Exception:
                return ""
        if width is None or height is None:
            return ""
        if not unit:
            unit = "px" if view.get("extent") is None and view.get("extent_raw") is None else "nm"

        def _fmt(value):
            value = float(value)
            if unit == "px":
                return str(int(round(value)))
            if abs(value - round(value)) < 1e-6:
                return str(int(round(value)))
            if abs(value) >= 100:
                return f"{value:.0f}"
            if abs(value) >= 10:
                return f"{value:.1f}".rstrip("0").rstrip(".")
            return f"{value:.3g}"

        return f"{_fmt(width)} x {_fmt(height)} {unit}".strip()

    def _draw_image_size_overlay(self, ax, view):
        if ax is None:
            return
        text = self._image_size_overlay_text(view)
        if not text:
            return
        scale = max(0.6, min(2.5, getattr(self, "_view_font_scale", 1.0)))
        fontsize = max(7.0, 8.2 * scale)
        y_pos = 0.16 if self.scale_bar_enabled else 0.02
        text_artist = ax.text(
            0.985,
            y_pos,
            text,
            transform=ax.transAxes,
            ha="right",
            va="bottom",
            fontsize=fontsize,
            fontweight="semibold",
            color="#f5f7fb",
            bbox={
                "facecolor": "black",
                "alpha": 0.35,
                "edgecolor": "none",
                "boxstyle": "round,pad=0.2",
            },
            zorder=21,
        )
        try:
            apply_text_style(text_artist, family=self._font_family, **self._plot_style_state())
        except Exception:
            pass

    def _draw_shortcut_hint(self, ax):
        if not self._show_shortcut_hint or ax is None:
            return
        self._clear_shortcut_hint_artist()
        scale = max(0.6, min(2.5, getattr(self, "_view_font_scale", 1.0)))
        fontsize = max(6.5, 7.0 * scale)
        hint = (
            "Ctrl+Click profile | Ctrl+Alt+Click angle | A auto contrast | 0 rel-zero | "
            "click molecule then X/Y/Z rotate, Shift+X/Y/Z reverse | Shift+drag mol = Z rotate | "
            "Ctrl+Shift+drag or middle-drag mol = 3D rotate | Ctrl+1/2/3 saved overlays | click to hide"
        )
        hint_artist = ax.text(
            0.012,
            0.012,
            hint,
            transform=ax.transAxes,
            ha="left",
            va="bottom",
            fontsize=fontsize,
            color="#f3f5f9",
            bbox={
                "facecolor": "black",
                "alpha": 0.32,
                "edgecolor": "none",
                "boxstyle": "round,pad=0.18",
            },
            zorder=24,
        )
        try:
            hint_artist.set_gid("ui_shortcut_hint")
            self._shortcut_hint_artist = hint_artist
            apply_text_style(hint_artist, family=self._font_family, **self._plot_style_state())
        except Exception:
            pass

    def _set_shortcut_hint_artist_visibility(self, visible: bool):
        art = getattr(self, "_shortcut_hint_artist", None)
        if art is None:
            return False
        try:
            art.set_visible(bool(visible))
            return True
        except Exception:
            return False

    def _save_current_figure_without_shortcut_hint(self, save_fn):
        if not callable(save_fn):
            return
        changed = False
        if getattr(self, "_show_shortcut_hint", False):
            changed = self._set_shortcut_hint_artist_visibility(False)
        try:
            save_fn()
        finally:
            if changed:
                self._set_shortcut_hint_artist_visibility(True)
                self.draw_idle()

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
            apply_text_style(ax.xaxis.label, family=self._font_family, **self._plot_style_state())
            apply_text_style(ax.yaxis.label, family=self._font_family, **self._plot_style_state())
            apply_text_style(ax.title, family=self._font_family, **self._plot_style_state())
            for lbl in list(ax.get_xticklabels()) + list(ax.get_yticklabels()):
                apply_text_style(lbl, family=self._font_family, **self._plot_style_state())
        except Exception:
            pass
        if cbar is not None:    
            try:
                cbar.ax.tick_params(labelsize=tick_size)
                cbar.ax.yaxis.label.set_fontsize(label_size)
                cbar.ax.xaxis.label.set_fontsize(label_size)
                apply_text_style(cbar.ax.yaxis.label, family=self._font_family, **self._plot_style_state())
                apply_text_style(cbar.ax.xaxis.label, family=self._font_family, **self._plot_style_state())
                for lbl in list(cbar.ax.get_xticklabels()) + list(cbar.ax.get_yticklabels()):
                    apply_text_style(lbl, family=self._font_family, **self._plot_style_state())
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
        self.push_undo_state("add_outline")
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
            has_outlines = bool(self._outlines)
        else:
            try:
                has_outlines = bool(self._outlines.get(self._outline_key(view)))
            except Exception:
                has_outlines = False
        if has_outlines:
            self.push_undo_state("clear_outlines")
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
        if entries:
            self.push_undo_state("outline_style")
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
        flip = self._use_relative_axes(view)
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
                channel_idx = view.get('channel_idx', meta.get('channel_index'))
                if channel_idx is not None:
                    try:
                        channel_idx = int(channel_idx)
                    except Exception:
                        channel_idx = None
                drag_token = self._stash_drag_view_snapshot(view)
                payload = {
                    'file_path': view.get('path') or meta.get('path') or meta.get('file_path'),
                    'channel_index': channel_idx,
                    'cmap': view.get('cmap'),
                    'drag_origin': 'preview_canvas',
                    'view_drag_token': drag_token,
                }
                if payload.get('file_path') is not None and payload.get('channel_index') is not None:
                    mime.setData('application/x-sxm-view', json.dumps(payload).encode('utf-8'))
                elif drag_token:
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

    def _molecule_paths_from_mime(self, mime):
        paths = []
        if mime is None or not mime.hasUrls():
            return paths
        for url in mime.urls():
            try:
                if not url.isLocalFile():
                    continue
                path = Path(url.toLocalFile())
            except Exception:
                continue
            if path.suffix.lower() in _MOLECULE_FILE_EXTS and path.exists():
                paths.append(path)
        return paths

    def dragEnterEvent(self, event):
        try:
            if self._molecule_paths_from_mime(event.mimeData()):
                event.acceptProposedAction()
                return
        except Exception:
            pass
        super().dragEnterEvent(event)

    def dragMoveEvent(self, event):
        try:
            if self._molecule_paths_from_mime(event.mimeData()):
                event.acceptProposedAction()
                return
        except Exception:
            pass
        super().dragMoveEvent(event)

    def dropEvent(self, event):
        try:
            paths = self._molecule_paths_from_mime(event.mimeData())
            if paths:
                for path in paths:
                    self.add_molecule(str(path))
                event.acceptProposedAction()
                return
        except Exception:
            pass
        super().dropEvent(event)

    def _on_motion_value(self, event):
        if self._fixed_crop_transform_mode:
            ax = event.inaxes if event is not None else None
            view = self._ax_view_map.get(ax) if ax is not None else None
            if self._fixed_crop_template_drag is not None:
                try:
                    if self._update_fixed_crop_template_drag(event):
                        mode = (self._fixed_crop_template_drag or {}).get("mode")
                        self._set_fixed_crop_cursor(mode=mode, dragging=True)
                        now_ms = time.perf_counter() * 1000.0
                        if (now_ms - float(self._fixed_crop_drag_last_ts or 0.0)) >= float(self._fixed_crop_drag_throttle_ms or 12.0):
                            self._fixed_crop_drag_last_ts = now_ms
                            drag_ax = (self._fixed_crop_template_drag or {}).get("ax") or ax
                            drag_view = self._ax_view_map.get(drag_ax) if drag_ax is not None else view
                            self._refresh_fixed_crop_overlay_fast(drag_ax, drag_view, dragging=True)
                        return
                except Exception:
                    pass
            else:
                mode = None
                if ax is not None and view is not None and event is not None and event.xdata is not None and event.ydata is not None:
                    hit = self._fixed_crop_template_handle_hit(event, view, ax)
                    if hit is not None:
                        mode = hit.get("mode")
                        if mode == "move" and bool(self._event_qt_modifiers(event) & QtCore.Qt.ControlModifier):
                            mode = "rotate"
                self._set_fixed_crop_cursor(mode=mode, dragging=False)
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
        if self._fixed_crop_template_drag is not None:
            drag_ax = (self._fixed_crop_template_drag or {}).get("ax")
            try:
                self._update_fixed_crop_template_drag(event)
            except Exception:
                pass
            self._finish_fixed_crop_template_drag()
            self._fixed_crop_drag_last_ts = 0.0
            drag_view = self._ax_view_map.get(drag_ax) if drag_ax is not None else None
            if drag_ax is not None and drag_view is not None:
                self._refresh_fixed_crop_overlay_fast(drag_ax, drag_view, dragging=False)
            else:
                self._redraw()
            return
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
        flip = self._use_relative_axes(view)
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
        c0 = self._axis_coord_to_pixel(view, x0c, w, 'x', ax=self._crop_ax)
        c1 = self._axis_coord_to_pixel(view, x1c, w, 'x', ax=self._crop_ax)
        r0 = self._axis_coord_to_pixel(view, y0c, h, 'y', ax=self._crop_ax)
        r1 = self._axis_coord_to_pixel(view, y1c, h, 'y', ax=self._crop_ax)
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
        crop_extent = self._compute_crop_extent(view, w, h, c0, c1, r0, r1)
        bounds_data = crop_extent if crop_extent is not None else self._pixel_bounds_to_axis_bounds(view, self._crop_ax, w, h, c0, c1, r0, r1)
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

    def enable_fixed_crop_transform_mode(self, enabled: bool):
        enabled = bool(enabled)
        if enabled == self._fixed_crop_transform_mode:
            return
        self._fixed_crop_transform_mode = enabled
        if enabled:
            self._fixed_crop_template_visible = True
            self._ensure_fixed_crop_template_for_transform()
            try:
                self.setFocus(QtCore.Qt.OtherFocusReason)
            except Exception:
                pass
        else:
            self._fixed_crop_template_drag = None
            self._fixed_crop_drag_last_ts = 0.0
            self._fixed_crop_template_visible = False
            self._set_fixed_crop_cursor(mode=None, dragging=False)
        self._notify_views_callback()
        self._redraw()

    def _fixed_crop_target_view(self, prefer_view=None, prefer_ax=None):
        if prefer_view is not None and prefer_ax is not None:
            return prefer_view, prefer_ax
        key = self._fixed_crop_template_view_key
        if key is not None:
            for ax, view in self._ax_view_map.items():
                if self._outline_key(view) == key:
                    return view, ax
        ax = self.main_ax or next(iter(self._ax_view_map.keys()), None)
        if ax is None:
            return None, None
        return self._ax_view_map.get(ax), ax

    def _ensure_fixed_crop_template_for_transform(self):
        if self._fixed_crop_template and self._fixed_crop_template.get("pixel_bounds"):
            return True
        ax = self.main_ax or next(iter(self._ax_view_map.keys()), None)
        view = self._ax_view_map.get(ax) if ax is not None else None
        if view is None or ax is None:
            return False
        arr_obj = view.get("arr")
        if arr_obj is None:
            return False
        arr = np.asarray(arr_obj)
        if arr.ndim < 2 or arr.size == 0:
            return False
        h, w = arr.shape[:2]
        seed_w = max(8, int(round(w * 0.45)))
        seed_h = max(8, int(round(h * 0.45)))
        template = self._compute_template_bounds_from_pixels(view, ax, seed_w, seed_h)
        if not template:
            return False
        bounds_data, pixel_bounds = template
        self._fixed_crop_template = {
            "width": int(abs(pixel_bounds[1] - pixel_bounds[0]) + 1),
            "height": int(abs(pixel_bounds[3] - pixel_bounds[2]) + 1),
            "square": False,
            "rotate": 0.0,
            "pixel_bounds": tuple(pixel_bounds),
        }
        self._fixed_crop_template_bounds = bounds_data
        self._fixed_crop_template_pixel_bounds = tuple(pixel_bounds)
        self._fixed_crop_template_view_key = self._outline_key(view)
        return True

    def show_fixed_crop_template(self, visible: bool):
        visible = bool(visible)
        if visible == self._fixed_crop_template_visible:
            if visible:
                self._ensure_fixed_crop_template_for_transform()
            return
        self._fixed_crop_template_visible = visible
        if self._fixed_crop_template_visible:
            self._ensure_fixed_crop_template_for_transform()
        self._redraw()

    def show_fixed_crop_history(self, visible: bool):
        visible = bool(visible)
        if visible == self._fixed_crop_history_visible:
            return
        self._fixed_crop_history_visible = visible
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

    def _history_entry_geometry(self, entry):
        if not entry:
            return None
        data_bounds = entry.get("data_bounds")
        if not data_bounds:
            return None
        x0, x1, y0, y1 = data_bounds
        left, right = min(x0, x1), max(x0, x1)
        bottom, top = min(y0, y1), max(y0, y1)
        width = right - left
        height = top - bottom
        if width <= 0 or height <= 0:
            return None
        angle = float(entry.get("rotate", 0.0) or 0.0)
        cx = (left + right) * 0.5
        cy = (bottom + top) * 0.5
        corners_local = np.array(
            [
                [left, bottom],
                [right, bottom],
                [right, top],
                [left, top],
            ],
            dtype=float,
        )
        if abs(angle) > 1e-9:
            rot = Affine2D().rotate_deg_around(cx, cy, angle)
            corners = rot.transform(corners_local)
            label_anchor = rot.transform((left + (width * 0.02), top - (height * 0.02)))
        else:
            corners = corners_local
            label_anchor = np.array([left + (width * 0.02), top - (height * 0.02)], dtype=float)
        return {
            "left": float(left),
            "right": float(right),
            "bottom": float(bottom),
            "top": float(top),
            "width": float(width),
            "height": float(height),
            "angle": float(angle),
            "corners": corners,
            "label_anchor": label_anchor,
        }

    def _update_highlight_artists(self):
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
                 if entry.get("key") == key and entry.get("sequence") == seq and bool(entry.get("visible", True))),
                None,
            )
            if entry is None:
                continue
            geom = self._history_entry_geometry(entry)
            if geom is None:
                continue
            active_keys.add(key)
            corners = geom["corners"]
            artists = self._fixed_crop_history_highlight_artists.get(key)
            if artists:
                fill, outline = artists
                fill.set_xy(corners)
                outline.set_xy(corners)
            else:
                fill = patches.Polygon(
                    corners,
                    closed=True,
                    linewidth=0,
                    edgecolor='none',
                    facecolor='#ff66ff',
                    alpha=0.15,
                    zorder=17,
                )
                outline = patches.Polygon(
                    corners,
                    closed=True,
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

    def _view_extent(self, view):
        if not view:
            return None
        raw_extent = view.get("extent_raw")
        if raw_extent is None:
            raw_extent = view.get("extent")
        extent = self._display_extent_for_view(view, raw_extent)
        if extent is None:
            extent = raw_extent
        try:
            extent = tuple(extent) if extent is not None else None
        except Exception:
            extent = None
        return extent if extent and len(extent) == 4 else None

    def _axis_coord_to_pixel(self, view, coord, length, axis_key, ax=None):
        if coord is None or length <= 0:
            return 0
        extent = self._view_extent(view)
        if extent is not None:
            lim0, lim1 = (extent[0], extent[1]) if axis_key == 'x' else (extent[2], extent[3])
        elif ax is not None:
            limits = ax.get_xlim() if axis_key == 'x' else ax.get_ylim()
            lim0, lim1 = limits
        else:
            lim0, lim1 = 0.0, float(max(1, length - 1))
        idx = _interp_index(coord, lim0, lim1, length)
        if idx is None:
            return 0
        return int(np.clip(round(idx), 0, length - 1))

    def _axis_coord_to_pixel_float(self, view, coord, length, axis_key, ax=None):
        if coord is None or length <= 0:
            return 0.0
        extent = self._view_extent(view)
        if extent is not None:
            lim0, lim1 = (extent[0], extent[1]) if axis_key == 'x' else (extent[2], extent[3])
        elif ax is not None:
            limits = ax.get_xlim() if axis_key == 'x' else ax.get_ylim()
            lim0, lim1 = limits
        else:
            lim0, lim1 = 0.0, float(max(1, length - 1))
        idx = _interp_index(coord, lim0, lim1, length)
        if idx is None:
            return 0.0
        return float(np.clip(idx, 0.0, max(0.0, float(length - 1))))

    def _index_to_axis_coord(self, idx, start, end, size):
        if size <= 0:
            return start
        denom = max(1, size - 1)
        frac = idx / denom if denom else 0.0
        if end > start:
            return start + (end - start) * frac
        return end + (start - end) * frac

    def _pixel_bounds_to_axis_bounds(self, view, ax, w, h, c0, c1, r0, r1):
        extent = self._view_extent(view)
        if extent is None and ax is None:
            return None
        def _interp(idx, lim, size):
            return self._index_to_axis_coord(idx, lim[0], lim[1], size)
        if extent is not None:
            xlim = (extent[0], extent[1])
            ylim = (extent[2], extent[3])
        else:
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
        x_extent0, x_extent1, y_extent_top, y_extent_bottom = ext
        return (
            self._index_to_axis_coord(c0, x_extent0, x_extent1, w),
            self._index_to_axis_coord(c1, x_extent0, x_extent1, w),
            self._index_to_axis_coord(r0, y_extent_top, y_extent_bottom, h),
            self._index_to_axis_coord(r1, y_extent_top, y_extent_bottom, h),
        )

    def _clear_fixed_crop_overlay_artists(self, ax=None):
        if ax is None:
            targets = list(self._fixed_crop_overlay_artists.keys())
        else:
            targets = [ax]
        for target_ax in targets:
            artists = self._fixed_crop_overlay_artists.pop(target_ax, [])
            for artist in artists:
                try:
                    artist.remove()
                except Exception:
                    pass

    def _fixed_crop_rotate_handle_points(self, ax, geom):
        angle = float(geom.get("angle", 0.0) or 0.0)
        rot = Affine2D().rotate_deg_around(geom["cx"], geom["cy"], angle)
        top_mid = rot.transform((geom["cx"], geom["top"]))
        center_px = np.asarray(ax.transData.transform((geom["cx"], geom["cy"])), dtype=float)
        top_mid_px = np.asarray(ax.transData.transform(top_mid), dtype=float)
        vec = top_mid_px - center_px
        norm = float(np.hypot(vec[0], vec[1]))
        if norm <= 1e-9:
            direction = np.array([0.0, -1.0], dtype=float)
        else:
            direction = vec / norm
        rotate_px = top_mid_px + (direction * 28.0)
        rotate_pt = ax.transData.inverted().transform(rotate_px)
        return top_mid, rotate_pt

    def _set_fixed_crop_cursor(self, mode=None, dragging=False):
        target = (mode, bool(dragging))
        if self._fixed_crop_cursor_mode == target:
            return
        self._fixed_crop_cursor_mode = target
        cursor = QtCore.Qt.ArrowCursor
        if mode == "rotate":
            cursor = QtCore.Qt.CrossCursor
        elif mode == "move":
            cursor = QtCore.Qt.ClosedHandCursor if dragging else QtCore.Qt.SizeAllCursor
        elif mode in ("resize_nw", "resize_se"):
            cursor = QtCore.Qt.SizeFDiagCursor
        elif mode in ("resize_ne", "resize_sw"):
            cursor = QtCore.Qt.SizeBDiagCursor
        try:
            self.setCursor(cursor)
        except Exception:
            pass

    def _refresh_fixed_crop_overlay_fast(self, ax, view, dragging=False):
        if ax is None:
            return
        if not self._fixed_crop_template_visible or view is None:
            self._clear_fixed_crop_overlay_artists(ax=ax)
            self.draw_idle()
            return
        updated = False
        try:
            updated = self._update_template_overlay_artists(ax, view, skip_label=bool(dragging))
        except Exception:
            updated = False
        if not updated:
            self._clear_fixed_crop_overlay_artists(ax=ax)
            try:
                self._render_template_overlay(ax, view)
            except Exception:
                pass
        self.draw_idle()

    def _fixed_crop_template_geometry(self, view, ax):
        template = self._fixed_crop_template
        if template is None or not template.get("pixel_bounds") or view is None or ax is None:
            return None
        c0, c1, r0, r1 = [int(v) for v in template.get("pixel_bounds")]
        arr_obj = view.get("arr")
        if arr_obj is None:
            return None
        arr = np.asarray(arr_obj)
        if arr.ndim < 2 or arr.size == 0:
            return None
        flip = self._use_relative_axes(view)
        arr_disp = np.flipud(arr) if flip else arr
        h, w = arr_disp.shape[:2]
        if w <= 0 or h <= 0:
            return None
        bounds = self._pixel_bounds_to_axis_bounds(view, ax, w, h, c0, c1, r0, r1)
        if not bounds:
            return None
        left, right, bottom, top = bounds
        left, right = (left, right) if left <= right else (right, left)
        bottom, top = (bottom, top) if bottom <= top else (top, bottom)
        width = right - left
        height = top - bottom
        if width <= 0 or height <= 0:
            return None
        angle = float(template.get("rotate", 0.0) or 0.0)
        return {
            "left": float(left),
            "right": float(right),
            "bottom": float(bottom),
            "top": float(top),
            "width": float(width),
            "height": float(height),
            "cx": float((left + right) * 0.5),
            "cy": float((bottom + top) * 0.5),
            "angle": angle,
            "pixel_bounds": (int(c0), int(c1), int(r0), int(r1)),
            "arr_shape": (h, w),
            "flip": flip,
        }

    def _fixed_crop_template_contains_point(self, xdata, ydata, geom):
        if geom is None or xdata is None or ydata is None:
            return False
        rot = Affine2D().rotate_deg_around(geom["cx"], geom["cy"], float(geom.get("angle", 0.0) or 0.0))
        inv = rot.inverted()
        lx, ly = inv.transform((xdata, ydata))
        return geom["left"] <= lx <= geom["right"] and geom["bottom"] <= ly <= geom["top"]

    def _fixed_crop_template_handle_hit(self, event, view, ax):
        if event is None or event.xdata is None or event.ydata is None:
            return None
        geom = self._fixed_crop_template_geometry(view, ax)
        if geom is None:
            return None
        angle = float(geom.get("angle", 0.0) or 0.0)
        rot = Affine2D().rotate_deg_around(geom["cx"], geom["cy"], angle)
        inv = rot.inverted()
        ev_world = np.array([event.xdata, event.ydata], dtype=float)
        ev_px = ax.transData.transform(ev_world)

        # Rotation handle (always rendered outward from the top edge in screen space)
        _top_mid_world, handle_world = self._fixed_crop_rotate_handle_points(ax, geom)
        handle_px = ax.transData.transform(handle_world)
        if float(np.hypot(*(ev_px - handle_px))) <= 22.0:
            return {"mode": "rotate", "geom": geom}

        # Corner handles
        corners = {
            "resize_nw": (geom["left"], geom["top"]),
            "resize_ne": (geom["right"], geom["top"]),
            "resize_sw": (geom["left"], geom["bottom"]),
            "resize_se": (geom["right"], geom["bottom"]),
        }
        for mode, pt in corners.items():
            pt_px = ax.transData.transform(rot.transform(pt))
            if float(np.hypot(*(ev_px - pt_px))) <= 18.0:
                return {"mode": mode, "geom": geom}

        lx, ly = inv.transform(ev_world)
        if geom["left"] <= lx <= geom["right"] and geom["bottom"] <= ly <= geom["top"]:
            return {"mode": "move", "geom": geom}
        return None

    def _update_fixed_crop_template_from_bounds(self, view, ax, left, right, bottom, top, angle=None):
        if view is None or ax is None:
            return False
        arr_obj = view.get("arr")
        if arr_obj is None:
            return False
        arr = np.asarray(arr_obj)
        if arr.ndim < 2 or arr.size == 0:
            return False
        flip = self._use_relative_axes(view)
        arr_disp = np.flipud(arr) if flip else arr
        h, w = arr_disp.shape[:2]
        left = float(left)
        right = float(right)
        bottom = float(bottom)
        top = float(top)
        left, right = (left, right) if left <= right else (right, left)
        bottom, top = (bottom, top) if bottom <= top else (top, bottom)
        c0 = self._axis_coord_to_pixel(view, left, w, 'x', ax=ax)
        c1 = self._axis_coord_to_pixel(view, right, w, 'x', ax=ax)
        r0 = self._axis_coord_to_pixel(view, bottom, h, 'y', ax=ax)
        r1 = self._axis_coord_to_pixel(view, top, h, 'y', ax=ax)
        if c1 < c0:
            c0, c1 = c1, c0
        if r1 < r0:
            r0, r1 = r1, r0
        bounds_data = self._pixel_bounds_to_axis_bounds(view, ax, w, h, c0, c1, r0, r1)
        if not bounds_data:
            return False
        self._fixed_crop_template = {
            "width": int(abs(c1 - c0) + 1),
            "height": int(abs(r1 - r0) + 1),
            "square": bool((self._fixed_crop_template or {}).get("square", False)),
            "rotate": float(angle if angle is not None else (self._fixed_crop_template or {}).get("rotate", 0.0) or 0.0),
            "pixel_bounds": (int(c0), int(c1), int(r0), int(r1)),
        }
        self._fixed_crop_template_bounds = bounds_data
        self._fixed_crop_template_pixel_bounds = (int(c0), int(c1), int(r0), int(r1))
        self._fixed_crop_template_view_key = self._outline_key(view)
        self._fixed_crop_template_manual_dims = None
        return True

    def _begin_fixed_crop_template_drag(self, hit, event, view, ax):
        if not hit or event is None or event.xdata is None or event.ydata is None:
            return False
        geom = hit.get("geom")
        if geom is None:
            return False
        self._fixed_crop_template_drag = {
            "mode": hit.get("mode"),
            "view_key": self._outline_key(view),
            "ax": ax,
            "press": (float(event.xdata), float(event.ydata)),
            "last": (float(event.xdata), float(event.ydata)),
            "bounds_start": (geom["left"], geom["right"], geom["bottom"], geom["top"]),
            "pixel_bounds_start": tuple(int(v) for v in (geom.get("pixel_bounds") or (0, 0, 0, 0))),
            "arr_shape": tuple(geom.get("arr_shape") or (0, 0)),
            "angle_start": float(geom.get("angle", 0.0) or 0.0),
            "center_start": (geom["cx"], geom["cy"]),
        }
        if hit.get("mode") == "rotate":
            cx, cy = geom["cx"], geom["cy"]
            self._fixed_crop_template_drag["press_angle"] = math.degrees(math.atan2(event.ydata - cy, event.xdata - cx))
        return True

    def _update_fixed_crop_template_drag(self, event):
        drag = self._fixed_crop_template_drag
        if not drag or event is None or event.xdata is None or event.ydata is None:
            return False
        current = (float(event.xdata), float(event.ydata))
        last = drag.get("last")
        if last is not None:
            if abs(current[0] - float(last[0])) < 1e-9 and abs(current[1] - float(last[1])) < 1e-9:
                return False
        drag["last"] = current
        ax = drag.get("ax")
        if event.inaxes is not ax:
            return False
        view = self._ax_view_map.get(ax)
        if view is None or self._outline_key(view) != drag.get("view_key"):
            return False
        left0, right0, bottom0, top0 = drag["bounds_start"]
        mode = drag.get("mode")
        angle = float(drag.get("angle_start", 0.0) or 0.0)
        press_x, press_y = drag.get("press", (event.xdata, event.ydata))

        if mode == "move":
            h, w = tuple(drag.get("arr_shape") or (0, 0))
            c0s, c1s, r0s, r1s = tuple(drag.get("pixel_bounds_start") or (0, 0, 0, 0))
            if h > 0 and w > 0 and c1s >= c0s and r1s >= r0s:
                press_c = self._axis_coord_to_pixel_float(view, press_x, w, "x", ax=ax)
                press_r = self._axis_coord_to_pixel_float(view, press_y, h, "y", ax=ax)
                curr_c = self._axis_coord_to_pixel_float(view, event.xdata, w, "x", ax=ax)
                curr_r = self._axis_coord_to_pixel_float(view, event.ydata, h, "y", ax=ax)
                if None not in (press_c, press_r, curr_c, curr_r):
                    width_px = max(1, int(c1s - c0s + 1))
                    height_px = max(1, int(r1s - r0s + 1))
                    dc = float(curr_c - press_c)
                    dr = float(curr_r - press_r)
                    start_c = float(c0s) + dc
                    start_r = float(r0s) + dr
                    max_c0 = max(0.0, float(w - width_px))
                    max_r0 = max(0.0, float(h - height_px))
                    c0 = int(round(np.clip(start_c, 0.0, max_c0)))
                    r0 = int(round(np.clip(start_r, 0.0, max_r0)))
                    c1 = int(c0 + width_px - 1)
                    r1 = int(r0 + height_px - 1)
                    bounds = self._pixel_bounds_to_axis_bounds(view, ax, w, h, c0, c1, r0, r1)
                    if bounds:
                        return self._update_fixed_crop_template_from_bounds(
                            view, ax, bounds[0], bounds[1], bounds[2], bounds[3], angle=angle
                        )
            dx = float(event.xdata - press_x)
            dy = float(event.ydata - press_y)
            return self._update_fixed_crop_template_from_bounds(
                view, ax, left0 + dx, right0 + dx, bottom0 + dy, top0 + dy, angle=angle
            )

        if mode == "rotate":
            cx, cy = drag.get("center_start", ((left0 + right0) * 0.5, (bottom0 + top0) * 0.5))
            press_angle = float(drag.get("press_angle", 0.0))
            current = math.degrees(math.atan2(event.ydata - cy, event.xdata - cx))
            delta = current - press_angle
            return self._update_fixed_crop_template_from_bounds(
                view, ax, left0, right0, bottom0, top0, angle=angle + delta
            )

        # Corner resize in the unrotated local frame.
        cx, cy = drag.get("center_start", ((left0 + right0) * 0.5, (bottom0 + top0) * 0.5))
        inv = Affine2D().rotate_deg_around(cx, cy, angle).inverted()
        lx, ly = inv.transform((event.xdata, event.ydata))
        min_span = 1e-9
        left, right, bottom, top = left0, right0, bottom0, top0
        if mode == "resize_nw":
            left = min(lx, right0 - min_span)
            top = max(ly, bottom0 + min_span)
        elif mode == "resize_ne":
            right = max(lx, left0 + min_span)
            top = max(ly, bottom0 + min_span)
        elif mode == "resize_sw":
            left = min(lx, right0 - min_span)
            bottom = min(ly, top0 - min_span)
        elif mode == "resize_se":
            right = max(lx, left0 + min_span)
            bottom = min(ly, top0 - min_span)
        else:
            return False

        if bool((self._fixed_crop_template or {}).get("square", False)):
            side = min(max(right - left, min_span), max(top - bottom, min_span))
            if mode == "resize_nw":
                left = right - side
                top = bottom + side
            elif mode == "resize_ne":
                right = left + side
                top = bottom + side
            elif mode == "resize_sw":
                left = right - side
                bottom = top - side
            elif mode == "resize_se":
                right = left + side
                bottom = top - side

        return self._update_fixed_crop_template_from_bounds(view, ax, left, right, bottom, top, angle=angle)

    def _finish_fixed_crop_template_drag(self):
        self._fixed_crop_template_drag = None

    def _register_crop_entry(self, view, bounds_data, pixel_bounds, square, angle=0.0, update_size=True):
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
            "visible": True,
            "square": bool(square),
            "rotate": float(angle or 0.0),
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
            if self._fixed_crop_template is not None:
                self._fixed_crop_template["rotate"] = float(angle or 0.0)
        if len(self._fixed_crop_history) > _FIXED_CROP_HISTORY_LIMIT:
            self._fixed_crop_history.pop(0)
        self._emit_fixed_crop_history_update()
        return entry

    def _render_history_entry(self, ax, entry, show_label=True):
        if not entry or ax is None:
            return
        if not bool(entry.get("visible", True)):
            return
        geom = self._history_entry_geometry(entry)
        if geom is None:
            return
        left = geom["left"]
        right = geom["right"]
        bottom = geom["bottom"]
        top = geom["top"]
        width = geom["width"]
        height = geom["height"]
        corners = geom["corners"]
        seq = entry.get("sequence")
        is_active = seq is not None and seq == self._fixed_crop_history_highlight_seq
        edge_color = entry.get("color") or ('#ff66ff' if is_active else '#ffd166')
        line_width = 2.3 if is_active else 1.8
        alpha = 1.0 if is_active else 0.6
        if is_active:
            highlight_fill = patches.Polygon(
                corners,
                closed=True,
                linewidth=0,
                edgecolor='none',
                facecolor=edge_color,
                alpha=0.15,
                zorder=17,
            )
            ax.add_patch(highlight_fill)
        rect = patches.Polygon(
            corners,
            closed=True,
            linewidth=line_width,
            edgecolor=edge_color,
            facecolor='none',
            alpha=alpha,
            linestyle='-',
            zorder=18,
        )
        ax.add_patch(rect)
        if is_active:
            highlight_outline = patches.Polygon(
                corners,
                closed=True,
                linewidth=max(line_width + 1.2, 3.0),
                edgecolor='#ffffff',
                facecolor='none',
                alpha=0.75,
                linestyle='-',
                zorder=19,
            )
            ax.add_patch(highlight_outline)
            ax.scatter(
                corners[:, 0],
                corners[:, 1],
                s=24,
                marker="s",
                color=edge_color,
                edgecolors="#ffffff",
                linewidths=0.35,
                alpha=0.9,
                zorder=20,
            )
        if not show_label or seq is None:
            return
        label_x = float(geom["label_anchor"][0])
        label_y = float(geom["label_anchor"][1])
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
        entries = [
            entry for entry in self._fixed_crop_history
            if entry.get("key") == key and bool(entry.get("visible", True))
        ]
        if not entries:
            return
        for entry in entries:
            self._render_history_entry(ax, entry)

    def _emit_fixed_crop_history_update(self):
        if callable(self._fixed_crop_history_callback):
            try:
                self._fixed_crop_history_callback(list(self._fixed_crop_history))
            except Exception:
                pass

    def set_fixed_crop_history_callback(self, cb):
        self._fixed_crop_history_callback = cb

    def get_fixed_crop_history_entry(self, seq):
        seq = int(seq) if seq is not None else None
        if seq is None:
            return None
        for entry in self._fixed_crop_history:
            if entry.get("sequence") == seq:
                clone = dict(entry)
                view_snapshot = entry.get("view_snapshot")
                if isinstance(view_snapshot, dict):
                    clone["view_snapshot"] = dict(view_snapshot)
                return clone
        return None

    def import_fixed_crop_history_entry(self, entry, *, update_size=True):
        if not isinstance(entry, dict):
            return None
        bounds_data = entry.get("data_bounds")
        if not bounds_data or len(bounds_data) != 4:
            return None
        target_view, target_ax = self._fixed_crop_target_view()
        if target_view is None or target_ax is None:
            return None
        source_key = entry.get("key") or ()
        target_key = self._outline_key(target_view)
        source_path = source_key[0] if isinstance(source_key, tuple) and source_key else None
        target_path = target_key[0] if isinstance(target_key, tuple) and target_key else None
        if source_path and target_path and str(source_path) != str(target_path):
            return None
        arr_obj = target_view.get("arr")
        if arr_obj is None:
            return None
        arr = np.asarray(arr_obj)
        if arr.ndim < 2 or arr.size == 0:
            return None
        h, w = arr.shape[:2]
        left, right = min(bounds_data[0], bounds_data[1]), max(bounds_data[0], bounds_data[1])
        bottom, top = min(bounds_data[2], bounds_data[3]), max(bounds_data[2], bounds_data[3])
        c_left = self._axis_coord_to_pixel_float(target_view, left, w, "x", ax=target_ax)
        c_right = self._axis_coord_to_pixel_float(target_view, right, w, "x", ax=target_ax)
        r_bottom = self._axis_coord_to_pixel_float(target_view, bottom, h, "y", ax=target_ax)
        r_top = self._axis_coord_to_pixel_float(target_view, top, h, "y", ax=target_ax)
        if any(v is None for v in (c_left, c_right, r_bottom, r_top)):
            return None
        c0 = int(np.clip(math.floor(min(c_left, c_right)), 0, max(0, w - 1)))
        c1 = int(np.clip(math.ceil(max(c_left, c_right)), 0, max(0, w - 1)))
        r0 = int(np.clip(math.floor(min(r_bottom, r_top)), 0, max(0, h - 1)))
        r1 = int(np.clip(math.ceil(max(r_bottom, r_top)), 0, max(0, h - 1)))
        imported = self._register_crop_entry(
            target_view,
            tuple(float(v) for v in bounds_data),
            (c0, c1, r0, r1),
            bool(entry.get("square", False)),
            angle=float(entry.get("rotate", 0.0) or 0.0),
            update_size=update_size,
        )
        if imported is None:
            return None
        imported["visible"] = bool(entry.get("visible", True))
        imported["real_size"] = tuple(entry.get("real_size") or imported.get("real_size") or (0.0, 0.0))
        imported["unit"] = entry.get("unit") or imported.get("unit")
        view_snapshot = entry.get("view_snapshot")
        if isinstance(view_snapshot, dict):
            snapshot = dict(view_snapshot)
            if imported.get("sequence") is not None:
                snapshot["crop_sequence"] = imported["sequence"]
            imported["view_snapshot"] = snapshot
        self._emit_fixed_crop_history_update()
        self.draw_idle()
        return imported

    def build_view_from_fixed_crop_entry(self, base_view, entry, *, title_suffix=" [crop]"):
        """Rebuild a cropped channel view using the geometry stored in crop history."""
        if not isinstance(base_view, dict) or not isinstance(entry, dict):
            return None
        bounds_data = entry.get("data_bounds")
        pixel_bounds = entry.get("pixel_bounds")
        if not bounds_data or len(bounds_data) != 4 or not pixel_bounds or len(pixel_bounds) != 4:
            return None
        angle = float(entry.get("rotate", 0.0) or 0.0)
        width_px = max(2, int(abs(pixel_bounds[1] - pixel_bounds[0]) + 1))
        height_px = max(2, int(abs(pixel_bounds[3] - pixel_bounds[2]) + 1))
        helper_ax = self.main_ax
        if helper_ax is None:
            fig = Figure(figsize=(1, 1))
            helper_ax = fig.add_subplot(111)
            extent = self._view_extent(base_view)
            if extent is not None and len(extent) == 4:
                try:
                    helper_ax.set_xlim(float(extent[0]), float(extent[1]))
                    helper_ax.set_ylim(float(extent[2]), float(extent[3]))
                except Exception:
                    pass
        if abs(angle) <= 1e-9:
            cropped_arr = self._extract_axis_aligned_crop(base_view, pixel_bounds)
        else:
            cropped_arr = self._extract_rotated_crop(base_view, helper_ax, bounds_data, width_px, height_px, angle)
        if cropped_arr is None:
            return None
        new_view = dict(base_view)
        try:
            new_view["arr"] = np.array(cropped_arr, copy=True)
        except Exception:
            new_view["arr"] = cropped_arr
        try:
            finite = np.asarray(cropped_arr)[np.isfinite(cropped_arr)]
            if finite.size:
                lo = float(finite.min())
                hi = float(finite.max())
                if hi <= lo:
                    hi = lo + 1e-12
                new_view["clim"] = (lo, hi)
            else:
                new_view.pop("clim", None)
        except Exception:
            new_view.pop("clim", None)
        crop_extent = tuple(float(v) for v in bounds_data)
        new_view["extent_raw"] = crop_extent
        display_extent = self._display_extent_for_view(new_view, crop_extent)
        if display_extent is not None:
            new_view["extent"] = display_extent
        else:
            new_view.pop("extent", None)
        base_title = str(base_view.get("title") or "crop")
        suffix = str(title_suffix or "")
        new_view["title"] = base_title if not suffix or base_title.endswith(suffix.strip()) else f"{base_title}{suffix}"
        try:
            real_size = tuple(entry.get("real_size") or ())
            if len(real_size) == 2:
                new_view["_popup_image_size_override"] = {
                    "width": float(real_size[0]),
                    "height": float(real_size[1]),
                    "unit": str(entry.get("unit") or self._guess_view_unit(base_view) or "nm"),
                }
        except Exception:
            pass

        seq = entry.get("sequence")
        if seq is not None:
            new_view["crop_sequence"] = seq
        return new_view

    def set_fixed_crop_history_entry_visible(self, seq, visible: bool):
        seq = int(seq) if seq is not None else None
        if seq is None:
            return False
        target = None
        for entry in self._fixed_crop_history:
            if entry.get("sequence") == seq:
                target = entry
                break
        if target is None:
            return False
        visible = bool(visible)
        if bool(target.get("visible", True)) == visible:
            return False
        target["visible"] = visible
        if self._fixed_crop_history_highlight_seq == seq and not visible:
            self._fixed_crop_history_highlight_seq = None
        self._cleanup_highlight_artists()
        self._emit_fixed_crop_history_update()
        self._redraw()
        return True

    def is_fixed_crop_history_entry_visible(self, seq):
        seq = int(seq) if seq is not None else None
        if seq is None:
            return False
        for entry in self._fixed_crop_history:
            if entry.get("sequence") == seq:
                return bool(entry.get("visible", True))
        return False

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
        flip = self._use_relative_axes(view)
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
        bounds = self._pixel_bounds_to_axis_bounds(view, ax, w, h, c0, c1, r0, r1)
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
        extent = self._view_extent(view)
        if not extent:
            return None
        x0, x1, y1, y0 = extent
        x_span = abs(x1 - x0)
        y_span = abs(y1 - y0)
        flip = self._use_relative_axes(view)
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
            "rotate": float(entry.get("rotate", 0.0) or 0.0),
            "pixel_bounds": tuple(pixel_bounds),
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
                    "rotate": float(self._fixed_crop_template.get("rotate", 0.0) or 0.0),
                    "pixel_bounds": tuple(pixel_bounds),
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
                    "rotate": float(self._fixed_crop_template.get("rotate", 0.0) or 0.0),
                    "pixel_bounds": tuple(pixel_bounds),
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
            "rotate": float((self._fixed_crop_template or {}).get("rotate", 0.0) or 0.0),
            "pixel_bounds": tuple(pixel_bounds),
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
            "rotate": float((self._fixed_crop_template or {}).get("rotate", 0.0) or 0.0),
            "pixel_bounds": tuple(pixel_bounds),
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
        self._clear_fixed_crop_overlay_artists(ax=ax)
        geom = self._fixed_crop_template_geometry(view, ax)
        if geom is None:
            return
        artists = []
        left = float(geom["left"])
        right = float(geom["right"])
        bottom = float(geom["bottom"])
        top = float(geom["top"])
        width = float(geom["width"])
        height = float(geom["height"])
        cx = float(geom["cx"])
        cy = float(geom["cy"])
        angle = float(geom.get("angle", 0.0) or 0.0)
        rot = Affine2D().rotate_deg_around(cx, cy, angle)

        corners_local = np.array(
            [
                [left, bottom],
                [right, bottom],
                [right, top],
                [left, top],
            ],
            dtype=float,
        )
        corners = rot.transform(corners_local)
        frame = patches.Polygon(
            corners,
            closed=True,
            linewidth=1.25,
            edgecolor="#f46cff",
            facecolor=(1.0, 0.58, 1.0, 0.025),
            alpha=0.9,
            linestyle="--",
            zorder=17,
        )
        ax.add_patch(frame)
        artists.append(frame)

        px_width = int(self._fixed_crop_template.get("width", int(width)))
        px_height = int(self._fixed_crop_template.get("height", int(height)))
        real_unit = self._fixed_crop_template_unit or "nm"
        real_label = ""
        if self._fixed_crop_template_bounds:
            bx0, bx1, by0, by1 = self._fixed_crop_template_bounds
            real_dx = abs(bx1 - bx0)
            real_dy = abs(by1 - by0)
            real_label = f"{real_dx:.3g} {real_unit} x {real_dy:.3g} {real_unit}"
        size_label = f"{real_label}\n({px_width}x{px_height} px)" if real_label else f"{px_width}x{px_height} px"

        label_anchor = rot.transform((left + (width * 0.015), top - (height * 0.02)))
        lbl = ax.text(
            float(label_anchor[0]),
            float(label_anchor[1]),
            size_label,
            color="#ff66ff",
            fontsize=7,
            weight="medium",
            verticalalignment="top",
            horizontalalignment="left",
            bbox=dict(facecolor="#111111", alpha=0.7, pad=1, edgecolor="none"),
            zorder=18,
        )
        artists.append(lbl)

        if self._fixed_crop_transform_mode:
            corner_size = 58.0
            corner_pts = ax.scatter(
                corners[:, 0],
                corners[:, 1],
                s=corner_size,
                marker="s",
                color="#f46cff",
                edgecolors="#ffe1ff",
                linewidths=0.5,
                zorder=19,
            )
            artists.append(corner_pts)
            top_mid, rotate_pt = self._fixed_crop_rotate_handle_points(ax, geom)
            handle_line = ax.plot(
                [top_mid[0], rotate_pt[0]],
                [top_mid[1], rotate_pt[1]],
                linestyle="-",
                linewidth=1.0,
                color="#ff66ff",
                alpha=0.9,
                zorder=19,
            )[0]
            artists.append(handle_line)
            rotate_marker = ax.scatter(
                [rotate_pt[0]],
                [rotate_pt[1]],
                s=62,
                marker="o",
                color="#222222",
                edgecolors="#f46cff",
                linewidths=0.9,
                zorder=20,
            )
            artists.append(rotate_marker)
            rotate_lbl = ax.text(
                float(rotate_pt[0]),
                float(rotate_pt[1]),
                "R",
                color="#ff66ff",
                fontsize=8,
                fontweight="bold",
                ha="center",
                va="center",
                zorder=21,
            )
            artists.append(rotate_lbl)
        self._fixed_crop_overlay_artists[ax] = artists

    def _update_template_overlay_artists(self, ax, view, skip_label=False):
        if ax is None or view is None:
            return False
        artists = self._fixed_crop_overlay_artists.get(ax)
        if not artists:
            return False
        geom = self._fixed_crop_template_geometry(view, ax)
        if geom is None:
            return False

        expected_len = 6 if self._fixed_crop_transform_mode else 2
        if len(artists) != expected_len:
            return False

        left = float(geom["left"])
        right = float(geom["right"])
        bottom = float(geom["bottom"])
        top = float(geom["top"])
        width = float(geom["width"])
        height = float(geom["height"])
        cx = float(geom["cx"])
        cy = float(geom["cy"])
        angle = float(geom.get("angle", 0.0) or 0.0)
        rot = Affine2D().rotate_deg_around(cx, cy, angle)

        corners_local = np.array(
            [
                [left, bottom],
                [right, bottom],
                [right, top],
                [left, top],
            ],
            dtype=float,
        )
        corners = rot.transform(corners_local)

        frame = artists[0]
        frame.set_xy(corners)

        label_anchor = rot.transform((left + (width * 0.015), top - (height * 0.02)))
        lbl = artists[1]
        lbl.set_position((float(label_anchor[0]), float(label_anchor[1])))
        if not skip_label:
            px_width = int(self._fixed_crop_template.get("width", int(width)))
            px_height = int(self._fixed_crop_template.get("height", int(height)))
            real_unit = self._fixed_crop_template_unit or "nm"
            real_label = ""
            if self._fixed_crop_template_bounds:
                bx0, bx1, by0, by1 = self._fixed_crop_template_bounds
                real_dx = abs(bx1 - bx0)
                real_dy = abs(by1 - by0)
                real_label = f"{real_dx:.3g} {real_unit} x {real_dy:.3g} {real_unit}"
            size_label = f"{real_label}\n({px_width}x{px_height} px)" if real_label else f"{px_width}x{px_height} px"
            lbl.set_text(size_label)

        if self._fixed_crop_transform_mode:
            corner_pts = artists[2]
            corner_pts.set_offsets(corners)
            top_mid, rotate_pt = self._fixed_crop_rotate_handle_points(ax, geom)
            handle_line = artists[3]
            handle_line.set_data(
                [float(top_mid[0]), float(rotate_pt[0])],
                [float(top_mid[1]), float(rotate_pt[1])],
            )
            rotate_marker = artists[4]
            rotate_marker.set_offsets(
                np.array([[float(rotate_pt[0]), float(rotate_pt[1])]], dtype=float)
            )
            rotate_lbl = artists[5]
            rotate_lbl.set_position((float(rotate_pt[0]), float(rotate_pt[1])))
        return True

    def _extract_rotated_crop(self, view, ax, bounds_data, width_px, height_px, angle_deg):
        if view is None or ax is None or not bounds_data:
            return None
        arr_obj = view.get("arr")
        if arr_obj is None:
            return None
        arr = np.asarray(arr_obj)
        if arr.ndim < 2 or arr.size == 0:
            return None
        flip = self._use_relative_axes(view)
        arr_disp = np.flipud(arr) if flip else arr
        h, w = arr_disp.shape[:2]
        if h <= 0 or w <= 0:
            return None
        width_px = int(np.clip(int(width_px), 2, w))
        height_px = int(np.clip(int(height_px), 2, h))
        x0, x1, y0, y1 = [float(v) for v in bounds_data]
        left, right = (x0, x1) if x0 <= x1 else (x1, x0)
        bottom, top = (y0, y1) if y0 <= y1 else (y1, y0)
        if right <= left or top <= bottom:
            return None

        xs = np.linspace(left, right, width_px, dtype=np.float64)
        ys = np.linspace(bottom, top, height_px, dtype=np.float64)
        gx, gy = np.meshgrid(xs, ys)

        angle_deg = float(angle_deg or 0.0)
        if abs(angle_deg) > 1e-9:
            cx = (left + right) * 0.5
            cy = (bottom + top) * 0.5
            points = np.column_stack((gx.ravel(), gy.ravel()))
            points = Affine2D().rotate_deg_around(cx, cy, angle_deg).transform(points)
            gx = points[:, 0].reshape((height_px, width_px))
            gy = points[:, 1].reshape((height_px, width_px))

        extent = self._view_extent(view)
        if extent is not None:
            xlim0, xlim1 = float(extent[0]), float(extent[1])
            ylim0, ylim1 = float(extent[2]), float(extent[3])
        else:
            xlim = ax.get_xlim()
            ylim = ax.get_ylim()
            xlim0, xlim1 = float(xlim[0]), float(xlim[1])
            ylim0, ylim1 = float(ylim[0]), float(ylim[1])

        def _coord_grid_to_index_grid(values, start, end, size):
            if size <= 1 or abs(end - start) <= 1e-15:
                return np.zeros_like(values, dtype=np.float64)
            if end > start:
                idx = (values - start) / (end - start)
            else:
                idx = (values - end) / (start - end)
            idx = idx * float(size - 1)
            return np.clip(idx, 0.0, float(size - 1))

        cols = _coord_grid_to_index_grid(gx, xlim0, xlim1, w)
        rows = _coord_grid_to_index_grid(gy, ylim0, ylim1, h)

        if _HAS_SCIPY and ndimage is not None:
            try:
                sampled = ndimage.map_coordinates(
                    arr_disp.astype(np.float64, copy=False),
                    [rows, cols],
                    order=1,
                    mode="nearest",
                )
            except Exception:
                sampled = None
        else:
            sampled = None
        if sampled is None:
            ri = np.clip(np.rint(rows).astype(np.int64), 0, h - 1)
            ci = np.clip(np.rint(cols).astype(np.int64), 0, w - 1)
            sampled = arr_disp[ri, ci]

        if sampled.size == 0:
            return None
        cropped_disp = np.asarray(sampled).reshape((height_px, width_px))
        cropped_arr = np.flipud(cropped_disp) if flip else cropped_disp
        return np.array(cropped_arr, copy=True)

    def _extract_axis_aligned_crop(self, view, pixel_bounds):
        if view is None or not pixel_bounds:
            return None
        arr_obj = view.get("arr")
        if arr_obj is None:
            return None
        arr = np.asarray(arr_obj)
        if arr.ndim < 2 or arr.size == 0:
            return None
        flip = self._use_relative_axes(view)
        arr_disp = np.flipud(arr) if flip else arr
        h, w = arr_disp.shape[:2]
        if h <= 0 or w <= 0:
            return None
        try:
            c0, c1, r0, r1 = [int(v) for v in pixel_bounds]
        except Exception:
            return None
        left = max(0, min(c0, c1))
        right = min(w - 1, max(c0, c1))
        top = max(0, min(r0, r1))
        bottom = min(h - 1, max(r0, r1))
        if right < left or bottom < top:
            return None
        cropped_disp = arr_disp[top:bottom + 1, left:right + 1]
        if cropped_disp.size == 0:
            return None
        cropped_arr = np.flipud(cropped_disp) if flip else cropped_disp
        return np.array(cropped_arr, copy=True)

    def _build_cropped_view_from_selection(
        self,
        view,
        ax,
        cropped_arr,
        bounds_data,
        pixel_bounds,
        square=False,
        angle=0.0,
        update_size=False,
        auto_virtual_copy=False,
        prompt_virtual_copy=True,
    ):
        if view is None or ax is None:
            return False
        if cropped_arr is None:
            return False
        arr = np.asarray(cropped_arr)
        if arr.ndim < 2 or arr.size == 0:
            return False
        entry = self._register_crop_entry(
            view,
            bounds_data,
            pixel_bounds,
            square,
            angle=angle,
            update_size=update_size,
        )
        new_view = dict(view)
        try:
            new_view["arr"] = np.array(arr, copy=True)
        except Exception:
            new_view["arr"] = arr
        try:
            finite = arr[np.isfinite(arr)]
            if finite.size:
                lo = float(finite.min())
                hi = float(finite.max())
                if hi <= lo:
                    hi = lo + 1e-12
                new_view["clim"] = (lo, hi)
            else:
                new_view.pop("clim", None)
        except Exception:
            new_view.pop("clim", None)
        if not new_view.get("path"):
            meta = new_view.get("meta") or {}
            src_path = meta.get("path") or meta.get("file_path")
            if src_path:
                new_view["path"] = src_path
        if bounds_data is not None:
            crop_extent = tuple(float(v) for v in bounds_data)
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
        if auto_virtual_copy:
            new_view["_auto_virtual_copy"] = True
        elif not prompt_virtual_copy:
            new_view["_skip_virtual_copy_prompt"] = True
        if entry and entry.get("sequence") is not None:
            try:
                real_size = tuple(entry.get("real_size") or ())
                if len(real_size) == 2:
                    new_view["_popup_image_size_override"] = {
                        "width": float(real_size[0]),
                        "height": float(real_size[1]),
                        "unit": str(entry.get("unit") or self._guess_view_unit(view) or "nm"),
                    }
            except Exception:
                pass
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

    def _apply_fixed_crop_template(self, view, ax):
        template = self._fixed_crop_template
        if template is None or ax is None or view is None:
            return False
        geom = self._fixed_crop_template_geometry(view, ax)
        if geom is None:
            return False
        width = int(template.get("width", 0) or 0)
        height = int(template.get("height", 0) or 0)
        if width < 2 or height < 2:
            return False
        angle = float(template.get("rotate", 0.0) or 0.0)
        bounds_data = (geom["left"], geom["right"], geom["bottom"], geom["top"])
        pixel_bounds = tuple(int(v) for v in (template.get("pixel_bounds") or geom.get("pixel_bounds") or (0, 0, 0, 0)))
        if abs(angle) <= 1e-9:
            cropped_arr = self._extract_axis_aligned_crop(view, pixel_bounds)
        else:
            cropped_arr = self._extract_rotated_crop(view, ax, bounds_data, width, height, angle)
        if cropped_arr is None:
            return False
        ok = self._build_cropped_view_from_selection(
            view=view,
            ax=ax,
            cropped_arr=cropped_arr,
            bounds_data=bounds_data,
            pixel_bounds=pixel_bounds,
            square=bool(template.get("square", False)),
            angle=angle,
            update_size=False,
            auto_virtual_copy=True,
        )
        if ok:
            if self._fixed_crop_quick_mode:
                self._fixed_crop_template_visible = True
                self._fixed_crop_template_drag = None
                self._fixed_crop_drag_last_ts = 0.0
                self._notify_views_callback()
                self._redraw()
            else:
                self.enable_fixed_crop_transform_mode(False)
        return ok

    def _apply_fixed_crop_quick(self, event, view, ax):
        template = self._fixed_crop_template
        if template is None or ax is None or event is None:
            return False
        if event.xdata is None or event.ydata is None:
            return False
        arr_obj = view.get("arr")
        if arr_obj is None:
            return False
        arr = np.asarray(arr_obj)
        if arr.ndim < 2 or arr.size == 0:
            return False
        flip = self._use_relative_axes(view)
        arr_disp = np.flipud(arr) if flip else arr
        h, w = arr_disp.shape[:2]
        if w <= 0 or h <= 0:
            return False
        width = int(np.clip(int(template.get("width", 2) or 2), 2, w))
        height = int(np.clip(int(template.get("height", 2) or 2), 2, h))
        cx = self._axis_coord_to_pixel_float(view, event.xdata, w, "x", ax=ax)
        cy = self._axis_coord_to_pixel_float(view, event.ydata, h, "y", ax=ax)
        max_c0 = max(0, w - width)
        max_r0 = max(0, h - height)
        c0 = int(np.clip(int(round(cx - (width * 0.5))), 0, max_c0))
        r0 = int(np.clip(int(round(cy - (height * 0.5))), 0, max_r0))
        c1 = int(c0 + width - 1)
        r1 = int(r0 + height - 1)
        bounds_data = self._pixel_bounds_to_axis_bounds(view, ax, w, h, c0, c1, r0, r1)
        if not bounds_data:
            return False
        angle = float(template.get("rotate", 0.0) or 0.0)
        pixel_bounds = (c0, c1, r0, r1)
        if abs(angle) <= 1e-9:
            cropped_arr = self._extract_axis_aligned_crop(view, pixel_bounds)
        else:
            cropped_arr = self._extract_rotated_crop(view, ax, bounds_data, width, height, angle)
        if cropped_arr is None:
            return False
        return self._build_cropped_view_from_selection(
            view=view,
            ax=ax,
            cropped_arr=cropped_arr,
            bounds_data=bounds_data,
            pixel_bounds=pixel_bounds,
            square=bool(template.get("square", False)),
            angle=angle,
            update_size=False,
            auto_virtual_copy=False,
            prompt_virtual_copy=False,
        )


class SafeFigureCanvas(FigureCanvas):
    def draw(self):
        try:
            super().draw()
        except np.linalg.LinAlgError:
            # Ignore transient singular transforms during layout updates.
            return

