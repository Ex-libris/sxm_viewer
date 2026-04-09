"""Main Qt widget implementing the SXM Viewer."""
from __future__ import annotations

import math
import time
import re
import json
import os
import copy
import shutil
import tempfile
import threading
from collections import OrderedDict, defaultdict
from functools import partial
from datetime import datetime
from pathlib import Path

import numpy as np
import io
from matplotlib import colormaps
from matplotlib.figure import Figure
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from PyQt5 import QtCore, QtGui, QtWidgets
import sip
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QDialog, QVBoxLayout, QCheckBox, QPushButton, QLabel, QListWidget, QListWidgetItem

from mpl_toolkits.axes_grid1 import make_axes_locatable
from .._shared import log_status, log_emitter
from ..app_meta import APP_NAME, apply_window_icon
from ..config import (
    CONFIG_PATH,
    CH_EQUALITY_TOL_NM,
    CH_SAMPLE_POINTS,
    CHANNEL_DATA_CACHE_LIMIT,
    FILTERED_CACHE_LIMIT,
    load_config,
    save_config,
    load_header_cache,
    save_header_cache,
)
from ..data.matrix import MatrixDataset, parse_matrix_filename
from ..data.io import parse_header, read_channel_file, normalize_unit_and_data

RECENT_FOLDER_LIMIT = 30
RECENT_SESSION_LIMIT = 30
from ..data.spectroscopy import is_matrix_file_entry
from ..processing.filters import (
    flatten_remove_median,
    subtract_best_fit_plane,
    subtract_2nd_order_plane,
    gaussian_filter_image,
    highpass_filter,
    laplacian_filter_image,
    FILTER_DEFINITIONS,
    _gaussian_available,
    _filter_signature,
)
from ..processing.detection import _find_topography_channel, _sample_channel_values_for_tagging
from ..utils.units import (
    _NUMERIC_RE,
    _UNIT_DISPLAY_CHOICES,
    _SI_BASE_UNITS,
    _auto_display_unit,
    _safe_float,
)
from .thumbnail_render import _ThumbnailJob, _colormap_icon, _value_in_nm, apply_adjustment_spec
from .thumbnail_render import _ThumbnailJob, _colormap_icon, _value_in_nm, apply_adjustment_spec, convert_to_si
from .minimap import FrameMiniMap
from .detail_panels import (
    BatchExportSignals,
    BatchExportWorker,
    CustomFilterDialog,
    ImageAdjustDialog,
    ImageAdjustPreviewPanel,
    MatrixFitDialog,
    MatrixFitWorker,
    MatrixSpectroViewer,
    MultiPreviewCanvas,
    ProfileDialog,
    SafeFigureCanvas,
    SingleFilterDialog,
    SpectroscopyCompareDialog,
    SpectroscopyPopup,
    _SpectroFitWorker,
)
from .spectroscopy.summary_dialog import SpectroSummaryDialog
from .viewer import measurement as viewer_measurement
from .viewer import thumbnails as viewer_thumbnails
from .controllers.preview_popup import spawn_preview_popup
from .controllers.histogram import open_histogram_dialog
from .controllers.quick_crop import QuickCropController
from .controllers.collection import CollectionController
from .controllers.thumbnail_controller import ThumbnailController
from .controllers.spectro_compare import SpectroCompareController
from .controllers.session import SessionController
from .viewer import loader as viewer_loader
from .viewer import preview as viewer_preview
from .viewer.state import ViewerState
from .plot_typography import add_font_menu_action, normalize_font_family, set_matplotlib_font_family
from .canvases.molecular_overlay import available_atom_palettes
from .spectroscopy import controller as spectro_controller
from .spectroscopy import overlays as spectro_overlays
from .spectroscopy import popups as spectro_popups
from .viewer import thumbnail_ui as viewer_thumb_ui
from .viewer import export as viewer_export
from .canvases.canvas_window import ExperimentalCanvasWindow
from .palettes import DEFAULT_COLOR_CYCLE
from .system_open import add_source_file_menu

# Tolerance for deciding constant-height images; allow a slightly larger spread than strict equality
CH_RANGE_TOL_NM = max(CH_EQUALITY_TOL_NM, 0.02)  # ~20 pm default floor
VIRTUAL_COPY_INSERT_START = "__virtual_copy_start__"


class _CollectionTrayList(QtWidgets.QListWidget):
    """Visual tray showing the current collection and accepting quick drag-add actions."""

    def __init__(self, viewer):
        super().__init__(viewer)
        self.viewer = viewer
        self.setAcceptDrops(True)
        self.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.setViewMode(QtWidgets.QListView.ListMode)
        self.setIconSize(QtCore.QSize(72, 72))
        self.setSpacing(6)
        self.setWordWrap(True)
        self.setResizeMode(QtWidgets.QListView.Adjust)
        self.setUniformItemSizes(False)
        self.setAlternatingRowColors(False)
        self.setToolTip(
            "Current collection tray.\n"
            "Drag thumbnails here to append fresh copies.\n"
            "Drag preview views here for quick copies.\n"
            "Use popup Collection actions when you want popup-specific overlays preserved."
        )

    def _restore_workspace_after_drop(self):
        try:
            self.viewer.on_recall_popouts()
        except Exception:
            pass
        try:
            host = self.window()
            if host is not None:
                host.show()
                host.raise_()
                host.activateWindow()
        except Exception:
            pass

    def _accepts_mime(self, mime):
        return bool(
            mime
            and (
                mime.hasFormat("application/x-sxm-thumb-selection")
                or mime.hasFormat("application/x-sxm-view")
            )
        )

    def dragEnterEvent(self, event):
        if self._accepts_mime(event.mimeData()):
            event.acceptProposedAction()
            return
        super().dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if self._accepts_mime(event.mimeData()):
            event.acceptProposedAction()
            return
        super().dragMoveEvent(event)

    def dropEvent(self, event):
        mime = event.mimeData()
        if mime is None:
            super().dropEvent(event)
            return
        try:
            if mime.hasFormat("application/x-sxm-thumb-selection"):
                payload = json.loads(bytes(mime.data("application/x-sxm-thumb-selection")).decode("utf-8"))
                entries = list((payload or {}).get("entries") or [])
                self.viewer.collection_controller.add_thumbnail_entries(entries)
                QtCore.QTimer.singleShot(0, self.viewer._refresh_collection_tray)
                QtCore.QTimer.singleShot(0, self._restore_workspace_after_drop)
                event.acceptProposedAction()
                return
            if mime.hasFormat("application/x-sxm-view"):
                payload = json.loads(bytes(mime.data("application/x-sxm-view")).decode("utf-8"))
                self.viewer.collection_controller.add_from_view_drag_payload(payload)
                QtCore.QTimer.singleShot(0, self.viewer._refresh_collection_tray)
                QtCore.QTimer.singleShot(0, self._restore_workspace_after_drop)
                event.acceptProposedAction()
                return
        except Exception:
            pass
        super().dropEvent(event)


class _CollectionTrayWindow(QtWidgets.QDialog):
    """Floating collection tray window kept separate from the main viewer layout."""

    def __init__(self, viewer, group_widget):
        super().__init__(viewer)
        self.viewer = viewer
        try:
            self.setParent(None, self.windowFlags())
            self.setWindowFlag(QtCore.Qt.Window, True)
            self.setWindowIcon(viewer.windowIcon())
            self.setWindowModality(QtCore.Qt.NonModal)
            self.setWindowFlags(
                self.windowFlags()
                | QtCore.Qt.WindowCloseButtonHint
                | QtCore.Qt.WindowMinimizeButtonHint
                | QtCore.Qt.WindowSystemMenuHint
            )
        except Exception:
            pass
        self.setWindowTitle("Collection Tray")
        self.resize(430, 520)
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(0)
        group_widget.setParent(self)
        group_widget.setVisible(True)
        layout.addWidget(group_widget)

    def _notify_popup_actions(self):
        try:
            controller = getattr(self.viewer, "quick_crop_controller", None)
            if controller:
                controller.update_popup_actions()
        except Exception:
            pass

    def showEvent(self, event):
        super().showEvent(event)
        self._notify_popup_actions()

    def hideEvent(self, event):
        super().hideEvent(event)
        self._notify_popup_actions()

    def closeEvent(self, event):
        self.hide()
        event.ignore()

# Patch export module with missing dependency
viewer_export.convert_to_si = convert_to_si

from . import main_window_layout
from . import main_window_spectro
from . import main_window_toolbar
from .constants import (
    LEFT_PANEL_SPACING,
    MAIN_SPLITTER_SIZES_COLUMNS,
    MAIN_SPLITTER_SIZES_STACKED,
    MAIN_WINDOW_SIZE,
    META_FONT_FAMILY,
    META_FONT_SIZE,
    THUMB_LAYOUT_SPACING,
    UI_FONT_BOLD_SIZE,
    UI_FONT_FAMILY,
    UI_FONT_SIZE,
)

class SXMGridViewer(QtWidgets.QWidget):
    SpectroSummaryDialog = SpectroSummaryDialog
    FRAME_ZOOM_SLIDER_MIN = 0
    FRAME_ZOOM_SLIDER_MAX = 600
    FRAME_ZOOM_SLIDER_DEFAULT = 200
    MODE_BROWSE = 0
    MODE_MEASURE = 1
    MODE_SPECTRO = 2

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        log_status("Initializing SXM Viewer...")
        self._app_start_ts = time.perf_counter()
        self.setWindowTitle(APP_NAME)
        apply_window_icon(self)
        self.resize(*MAIN_WINDOW_SIZE)

        log_status("Loading configuration...")
        self.config = load_config()
        time_source = self.config.get("image_time_source", "mtime")
        if time_source not in ("mtime", "header"):
            time_source = "mtime"
        self.image_time_source = time_source
        if self.config.get("image_time_source") != time_source:
            self.config["image_time_source"] = time_source
            save_config(self.config)
        self.last_dir = Path(self.config.get("last_dir", str(Path.cwd())))
        raw_recents = self.config.get("recent_dirs", [])
        self.recent_dirs = []
        for entry in raw_recents:
            if not entry:
                continue
            try:
                self.recent_dirs.append(str(Path(entry)))
            except Exception:
                continue
        raw_session_recents = self.config.get("recent_session_paths", self.config.get("recent_session_dirs", []))
        self.recent_session_paths = []
        for entry in raw_session_recents:
            if not entry:
                continue
            try:
                self.recent_session_paths.append(str(Path(entry)))
            except Exception:
                continue
        self._normalize_recent_session_history(persist=True)
        config_changed = False
        if "session_recovery_enabled" not in self.config:
            self.config["session_recovery_enabled"] = True
            config_changed = True
        if "session_recovery_interval_min" not in self.config:
            self.config["session_recovery_interval_min"] = 5
            config_changed = True
        if config_changed:
            save_config(self.config)
        self._current_session_path = None
        self._closed_window_history = []
        self._closed_window_history_limit = 6
        self._suspend_window_history = False
        self._autosave_busy = False
        self._session_recovery_enabled = bool(self.config.get("session_recovery_enabled", True))
        try:
            self._session_recovery_interval_min = max(1, int(self.config.get("session_recovery_interval_min", 5) or 5))
        except Exception:
            self._session_recovery_interval_min = 5
        self._workspace_window_shutdown = False
        self.last_channel_index = int(self.config.get("last_channel_index", 0))
        default_cmap = "Blues_r"
        thumb_cfg = self.config.get("thumbnail_cmap")
        preview_cfg = self.config.get("preview_cmap")
        config_changed = False
        if not thumb_cfg and not preview_cfg:
            thumb_cfg = preview_cfg = default_cmap
            self.config['thumbnail_cmap'] = thumb_cfg
            self.config['preview_cmap'] = preview_cfg
            config_changed = True
        elif not thumb_cfg:
            thumb_cfg = preview_cfg or default_cmap
            self.config['thumbnail_cmap'] = thumb_cfg
            config_changed = True
        elif not preview_cfg:
            preview_cfg = thumb_cfg or default_cmap
            self.config['preview_cmap'] = preview_cfg
            config_changed = True
        self.thumb_cmap = thumb_cfg or default_cmap
        self.preview_cmap = preview_cfg or self.thumb_cmap
        if config_changed:
            save_config(self.config)
        self.spec_folder_path = Path(self.config.get("spectra_folder", str(self.last_dir)))
        self.show_spectra = bool(self.config.get("show_spectra", True))
        self.show_spectro_miniatures = bool(self.config.get("show_spectro_miniatures", False))
        self.spectro_share_overlapping_repeats = bool(self.config.get("spectro_share_overlapping_repeats", False))
        self.spectro_miniature_default_channel = str(self.config.get("spectro_miniature_default_channel", "") or "")
        self.spectro_thumb_channel_by_path = dict(self.config.get("spectro_thumb_channel_by_path", {}) or {})
        self.spectro_highlight_glow = bool(self.config.get("spectro_highlight_glow", True))
        preview_cfg = self.config.get("show_preview_spectra")
        if preview_cfg is None:
            preview_cfg = self.show_spectra
        self.show_preview_spectra = bool(preview_cfg)
        # Defaults: disable tag auto-detection and allow users to re-enable via config
        self.auto_detect_tags = bool(self.config.get("auto_detect_tags", False))
        # Allow skipping Nanonis scan conversion if cache already exists
        self.convert_nanonis_enabled = bool(self.config.get("convert_nanonis_enabled", True))
        # Enable persistent spectroscopy disk cache (per-folder) by default
        self.spectro_disk_cache_enabled = bool(self.config.get("spectro_disk_cache_enabled", True))
        self.spectro_manifest_cache_enabled = bool(self.config.get("spectro_manifest_cache_enabled", True))
        self.spectro_lazy_payload_enabled = bool(self.config.get("spectro_lazy_payload_enabled", True))
        # Lazily load spectroscopies (defer until requested) to speed up initial folder loads
        self.lazy_spectros_enabled = bool(self.config.get("lazy_spectros_enabled", True))
        self.thumb_size_px = int(self.config.get("thumb_size_px", 160))
        self.thumb_grid_columns = 1
        self.display_units_si = bool(self.config.get("display_units_si", False))
        self.display_units_relative = bool(self.config.get("display_units_relative", False))
        self.relative_axes = bool(self.config.get("relative_axes", False))
        self.preserve_profiles_on_channel_change = bool(
            self.config.get("preserve_profiles_on_channel_change", True)
        )
        self.tags = self.config.get("tags", {})  # persistent tags: {path: {"tag":"constant-height","abs_z_pm":int,...}}
        if not self.auto_detect_tags:
            self.tags = {
                str(key): value
                for key, value in dict(self.tags or {}).items()
                if not (isinstance(value, dict) and value.get("auto") and not value.get("manual"))
            }
            self.config["tags"] = self.tags
            save_config(self.config)
        self.session_controller = SessionController(self)
        self.collection_controller = CollectionController(self)
        self.frame_map_entries = []
        self.show_shortcuts_panel = bool(self.config.get("show_shortcuts_panel", False))
        self.hidden_frame_keys = set()
        self.frame_real_view = False
        self.show_matrix_markers = bool(self.config.get("show_matrix_markers", True))
        # default to showing single markers so spectroscopies are visible by default
        self.show_single_markers = bool(self.config.get("show_single_markers", True))
        self.compact_markers = bool(self.config.get("compact_markers", True))
        self.spectro_single_grid_as_matrix = bool(self.config.get("spectro_single_grid_as_matrix", False))
        self.spectro_force_single_mode = bool(self.config.get("spectro_force_single_mode", False))
        self.dark_mode = bool(self.config.get('dark_mode', False))
        self.detail_dark_view = bool(self.config.get('detail_dark_view', self.dark_mode))
        self._detail_theme_follows_dark_mode = bool(self.config.get('detail_theme_follows_dark_mode', True))
        self.detail_grid_view = bool(self.config.get('detail_grid_view', False))
        self.show_molecules = bool(self.config.get('show_molecules', True))
        self.show_molecule_gizmo = bool(self.config.get("show_molecule_gizmo", False))
        self.show_acquisition_overlay = bool(self.config.get("show_acquisition_overlay", False))
        self.profile_label_mode = str(self.config.get("profile_label_mode", "length") or "length").strip().lower()
        if self.profile_label_mode not in {"length", "full", "hidden"}:
            self.profile_label_mode = "length"
        self.canvas_display_options = dict(self.config.get("canvas_display_options", {}))
        self.molecule_palette = str(self.config.get("molecule_palette", "cpk") or "cpk").lower()
        self.recent_molecules = list(self.config.get("recent_molecules", []))
        self.quick_crop_mode = bool(self.config.get("quick_crop_mode", False))
        self.quick_crop_aspect_mode = str(self.config.get("quick_crop_aspect_mode", "free") or "free").strip().lower()
        if self.quick_crop_aspect_mode not in {"free", "keep", "square"}:
            self.quick_crop_aspect_mode = "free"
        # Keep crop template editor opt-in at startup for cleaner preview/popup canvases.
        self.show_crop_template_overlay = False
        self.show_crop_history_overlay = True
        self._collection_item_snapshots = {}
        self._collection_source = None
        self._current_collection_mode = None
        self._workspace_kind = "folder"
        self._display_defaults = {
            'show_matrix_markers': True,
            'show_single_markers': True,
            'compact_markers': True,
            'detail_dark_view': bool(self.dark_mode),
            'detail_grid_view': False,
            'show_molecules': True,
            'show_molecule_gizmo': False,
            'show_acquisition_overlay': False,
            'profile_label_mode': "length",
            'show_crop_template_overlay': False,
            'show_crop_history_overlay': True,
        }
        self._popup_canvases = []
        self._active_preview_popup = None
        self._active_preview_canvas = None
        c_single = self.config.get('spectro_marker_color_single')
        if c_single:
            self.spectro_marker_color_single = QtGui.QColor(c_single)
        else:
            self.spectro_marker_color_single = QtGui.QColor(255, 20, 147, 255)
        c_matrix = self.config.get('spectro_marker_color_matrix')
        if c_matrix:
            self.spectro_marker_color_matrix = QtGui.QColor(c_matrix)
        else:
            self.spectro_marker_color_matrix = QtGui.QColor(64, 200, 255, 200)
        self.spectro_color_cycle = self.config.get('spectro_color_cycle', DEFAULT_COLOR_CYCLE)
        self.spectro_marker_symbol = self.config.get('spectro_marker_symbol', 'circle')
        self.spectro_marker_size = float(self.config.get('spectro_marker_size', 5.0))
        self.frame_entry_pixmaps = {}
        self._frame_real_pixmap_cache = {}
        self._processed_views = {}
        self.molecule_overlays = {}
        self._temp_reveal = set()
        self.spectro_dock = None
        self._spectro_browser_entries = []
        self._highlight_phase = 0.0
        self._highlight_pulse_strength = 1.0
        self._highlight_timer = QtCore.QTimer(self)
        # Debounced marker refresh to avoid repaint storms
        self._marker_refresh_timer = QtCore.QTimer(self)
        self._marker_refresh_timer.setSingleShot(True)
        self._marker_refresh_timer.timeout.connect(self._refresh_thumbnail_markers)
        self._spectro_manifest_save_timer = QtCore.QTimer(self)
        self._spectro_manifest_save_timer.setSingleShot(True)
        self._spectro_manifest_save_timer.timeout.connect(self._flush_spectro_manifest_save)
        self._spectro_manifest_save_inflight = False
        self._spectro_manifest_save_pending = False
        # Preview docking state
        self.preview_detached = False
        self.preview_locked = bool(self.config.get("preview_locked", False))
        self._preview_dialog = None
        self._highlight_timer.setInterval(350)
        self._highlight_timer.timeout.connect(self._on_highlight_tick)
        self._highlighted_spec = None

        self.files = []
        self.headers = {}
        self.thumb_cache = {}
        self._thumb_data_cache = {}
        self._thumb_crop_cache = {}
        self._topo_stats_cache = {}
        self._channel_data_cache = OrderedDict()
        self._channel_cache_lock = threading.Lock()
        self._filtered_channel_cache = OrderedDict()
        self._filtered_cache_lock = threading.Lock()
        self._thumb_labels = {}
        self._thumb_generation = 0
        self._thumb_data_lock = threading.Lock()
        self._thumb_threadpool = QtCore.QThreadPool()
        self._thumb_meta = {}
        self._thumb_loaded = set()
        self._thumb_inflight = set()
        self._thumb_card_height = None
        try:
            self._thumb_threadpool.setMaxThreadCount(max(2, min(6, QtCore.QThreadPool.globalInstance().maxThreadCount())))
        except Exception:
            pass
        self._pending_profile_enable = False
        self._pending_angle_enable = False
        self._last_profile_payload = None

        self.per_file_channel_cmap = {}
        self.last_preview = None
        self.spectros = []
        self.matrix_spectros = []
        self.files_with_matrix = set()
        self.spectros_by_image = defaultdict(list)
        self._spectros_loaded = False
        self._spectros_loading = False
        self._spectros_pending = False
        self._spectro_cache = {}
        self._spectro_manifest_entries = {}
        self._spectro_deferred = set()
        self._spectro_miniature_cache = OrderedDict()
        self._spectro_autoload_timer = QtCore.QTimer(self)
        self._spectro_autoload_timer.setSingleShot(True)
        self._spectro_autoload_timer.timeout.connect(self._run_pending_spectro_load)
        # spectro_eager_limit: 0 means no deferral; otherwise parse at most N spectroscopy files eagerly
        limit_cfg = int(self.config.get("spectro_eager_limit", 300))
        self.spectro_eager_limit = max(0, limit_cfg)
        self.image_time_index = {}
        self._spectro_popups = []
        self._popup_refs = []
        self._deferred_popup_entries = []
        self._deferred_popup_serial = 0
        self._multi_spectro_popups = []
        self._multi_single_popup_anchor = None
        self._last_clicked_spec = None
        self._popup_counter = 0  # used to stagger dialog positions
        self._multi_spec_selection = []
        self._multi_spec_selection_keys = set()
        self._workspace_loading = False
        self.spectro_compare_controller = SpectroCompareController(self)
        from .controllers.image_compare import ImageCompareController

        self.image_compare_controller = ImageCompareController(self)
        self.thumb_multi_select = set()
        self.spectro_thumb_multi_select = set()
        self.current_spectro_thumb_files = []
        self.selected_spectro_thumb_file = None
        self._canvas_display_syncing = False
        self._last_canvas_display_options = {}
        self._profile_dialogs = []
        self._clipboard_export_dir = None
        self._clipboard_copy_worker = None
        self._clipboard_copy_total = 0
        self._toast_registry = {}
        self._batch_export_progress = None
        self._batch_export_worker = None
        self.virtual_copies = {}
        self.virtual_copy_order = []
        self.thumbnail_filters = {}
        self.image_adjustments = defaultdict(dict)
        self._last_base_array = None
        self._last_base_extent = None
        self._last_base_unit = None
        self._spectro_hist_cache = {}
        self.matrix_datasets = {}
        log_status("Loading header cache...")
        self.header_cache = load_header_cache()
        self._header_cache_dirty = False
        self.state = ViewerState.from_viewer(self)
        # Deprecated: previously stored concrete arrays for extra views
        # self.added_views kept for backward compatibility but not used for rendering
        self.added_views = []
        # New: store extra view specifications to rebuild per selected file
        # Each spec: { 'caption': str, 'index': int, 'cmap': str }
        self.extra_view_specs = []
        # Thumbnail helpers: mapping from file path -> container widget for selection styling
        self.thumb_widgets = {}
        self.selected_file_for_thumbs = None

        # Plot typography defaults are shared across preview, popups and dialogs.
        self._plot_font_family = normalize_font_family(self.config.get("plot_font_family", UI_FONT_FAMILY), UI_FONT_FAMILY)
        self._plot_font_bold = bool(self.config.get("plot_font_bold", False))
        self._plot_font_italic = bool(self.config.get("plot_font_italic", False))
        self._plot_font_underline = bool(self.config.get("plot_font_underline", False))
        set_matplotlib_font_family(self._plot_font_family)
        # fonts
        base_font = QtGui.QFont(UI_FONT_FAMILY, UI_FONT_SIZE)
        bold_font = QtGui.QFont(UI_FONT_FAMILY, UI_FONT_BOLD_SIZE, QtGui.QFont.Bold)
        meta_font = QtGui.QFont(META_FONT_FAMILY, META_FONT_SIZE)
        try:
            app = QtWidgets.QApplication.instance()
            if app is not None:
                app.setFont(base_font)
        except Exception:
            pass

        self.toolbar_open_act = None
        self.toolbar_export_png_act = None
        self.toolbar_export_xyz_act = None
        self.toolbar_load_session_act = None
        self.toolbar_load_session_btn = None
        self.toolbar_load_session_menu = None
        self.toolbar_save_session_act = None
        self.toolbar_popups_raise_act = None
        self.toolbar_popups_btn = None
        self.toolbar_popups_menu = None
        self.toolbar_adjust_act = None
        self.toolbar_dark_btn = None
        self.toolbar_display_btn = None
        self.toolbar_image_btn = None
        self.toolbar_image_menu = None
        self.toolbar_tools_btn = None
        self.toolbar_tools_menu = None
        self.toolbar_load_mol_btn = None
        self.toolbar_spectro_btn = None
        self.toolbar_spectro_menu = None
        self.toolbar_spectro_markers_act = None
        self.toolbar_spectro_preview_act = None
        self.toolbar_spectro_miniatures_act = None
        self.toolbar_spectro_matrix_markers_act = None
        self.toolbar_spectro_single_markers_act = None
        self.toolbar_spectro_compact_markers_act = None
        self.toolbar_spectro_highlight_act = None
        self.toolbar_spectro_grid_as_matrix_act = None
        self.toolbar_spectro_force_single_act = None
        self.toolbar_spectro_thumb_btn = None
        self.toolbar_spectro_preview_btn = None
        self.toolbar_spectro_miniatures_btn = None
        self.preview_spectra_toggle_btn = None
        self.browse_molecules_btn = None
        self.browse_molecules_menu = None
        self.preview_molecules_toggle_btn = None
        self.display_molecule_gizmo_act = None
        self.preview_grid_toggle_btn = None
        self.preview_adjust_btn = None
        self._canvas_window = None
        self._session_activity_strip = None
        self._session_activity_title = None
        self._session_activity_detail = None
        self._session_activity_progress = None
        self._session_activity_hide_timer = QtCore.QTimer(self)
        self._session_activity_hide_timer.setSingleShot(True)
        self._session_activity_hide_timer.timeout.connect(self._hide_session_activity)
        self._activity_log_pending = []
        self._activity_log_flush_timer = QtCore.QTimer(self)
        self._activity_log_flush_timer.setSingleShot(True)
        self._activity_log_flush_timer.timeout.connect(self._flush_activity_log_pending)

        # UI: left controls + meta + inspector; middle thumbs; right preview
        left_v = QtWidgets.QVBoxLayout(); left_v.setSpacing(LEFT_PANEL_SPACING)
        essentials_group = QtWidgets.QGroupBox("Data paths")
        essentials_layout = QtWidgets.QVBoxLayout(essentials_group)

        # Images path (label above to save horizontal space)
        img_container = QtWidgets.QWidget()
        img_v = QtWidgets.QVBoxLayout(img_container)
        img_v.setContentsMargins(0, 0, 0, 0)
        img_v.setSpacing(4)
        img_v.addWidget(QtWidgets.QLabel("Images"))
        path_h = QtWidgets.QHBoxLayout()
        self.path_le = QtWidgets.QLineEdit(str(self.last_dir))
        self.open_btn = QtWidgets.QToolButton()
        self.open_btn.setText("Open folder")
        self.open_btn.setPopupMode(QtWidgets.QToolButton.MenuButtonPopup)
        self.open_recent_menu = QtWidgets.QMenu(self.open_btn)
        self.open_btn.setMenu(self.open_recent_menu)
        path_h.addWidget(self.path_le); path_h.addWidget(self.open_btn)
        self._refresh_recent_dirs_menu()
        img_v.addLayout(path_h)
        essentials_layout.addWidget(img_container)

        # Spectra path: label above the path field
        spec_container = QtWidgets.QWidget()
        spec_v = QtWidgets.QVBoxLayout(spec_container)
        spec_v.setContentsMargins(0, 0, 0, 0)
        spec_v.setSpacing(4)
        spec_v.addWidget(QtWidgets.QLabel("Spectra"))
        spec_row = QtWidgets.QHBoxLayout()
        self.spec_folder_le = QtWidgets.QLineEdit(str(self.spec_folder_path))
        self.spec_folder_le.setPlaceholderText("Defaults to SXM folder")
        self.spec_folder_btn = QtWidgets.QPushButton("Browse")
        spec_row.addWidget(self.spec_folder_le, 1)
        spec_row.addWidget(self.spec_folder_btn)
        spec_v.addLayout(spec_row)
        essentials_layout.addWidget(spec_container)

        # Channel controls stay near the preview workspace because channel switching is a preview task.
        controls_h = QtWidgets.QHBoxLayout()
        controls_h.setContentsMargins(0, 0, 0, 0)
        controls_h.setSpacing(6)
        self.channel_label = QtWidgets.QLabel("Channel")
        self.channel_label.setFont(bold_font)
        self.channel_prev_btn = QtWidgets.QToolButton()
        self.channel_prev_btn.setArrowType(QtCore.Qt.LeftArrow)
        self.channel_prev_btn.setAutoRaise(True)
        self.channel_prev_btn.setToolTip("Previous channel")
        self.channel_dropdown = QtWidgets.QComboBox()
        self.channel_dropdown.setMinimumWidth(240)
        self.channel_dropdown.setMinimumContentsLength(16)
        self.channel_dropdown.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        try:
            self.channel_dropdown.setSizeAdjustPolicy(QtWidgets.QComboBox.AdjustToMinimumContentsLengthWithIcon)
        except Exception:
            pass
        self.channel_next_btn = QtWidgets.QToolButton()
        self.channel_next_btn.setArrowType(QtCore.Qt.RightArrow)
        self.channel_next_btn.setAutoRaise(True)
        self.channel_next_btn.setToolTip("Next channel")
        self.thumb_cmap_combo = QtWidgets.QComboBox(); self.preview_cmap_combo = QtWidgets.QComboBox()
        
        # populate colormap combos with all available matplotlib colormaps and icons
        try:
            cmap_list = sorted(colormaps.keys())
        except Exception:
            cmap_list = ['viridis','plasma','inferno','magma','cividis','gray','hot','coolwarm','turbo']
        for m in cmap_list:
            try:
                icon = _colormap_icon(m, width=96, height=14)
            except Exception:
                icon = QIcon()
            self.thumb_cmap_combo.addItem(icon, m)
            self.preview_cmap_combo.addItem(icon, m)

        self.thumb_cmap_combo.setCurrentText(self.thumb_cmap); self.preview_cmap_combo.setCurrentText(self.preview_cmap)
        controls_h.addWidget(self.channel_label)
        controls_h.addWidget(self.channel_prev_btn)
        controls_h.addWidget(self.channel_dropdown, 1)
        controls_h.addWidget(self.channel_next_btn)

        # Dark mode handled via toolbar toggle; placeholder kept for compatibility
        self.dark_mode_cb = None
        left_v.addWidget(essentials_group)

        self.collection_group = QtWidgets.QGroupBox("Current Collection")
        collection_layout = QtWidgets.QVBoxLayout(self.collection_group)
        collection_layout.setContentsMargins(8, 8, 8, 8)
        collection_layout.setSpacing(6)
        self.collection_target_label = QtWidgets.QLabel("No collection selected yet.")
        self.collection_target_label.setWordWrap(True)
        self.collection_target_label.setStyleSheet("font-weight: 600;")
        self.collection_target_label.setTextInteractionFlags(QtCore.Qt.TextSelectableByMouse)
        collection_layout.addWidget(self.collection_target_label)
        self.collection_hint_label = QtWidgets.QLabel(
            "Drag thumbnails here to append fresh copies, or use popup Collection actions to preserve popup overlays.",
            self.collection_group,
        )
        self.collection_hint_label.setWordWrap(True)
        self.collection_hint_label.setStyleSheet("color: #5a6b7d;")
        collection_layout.addWidget(self.collection_hint_label)
        collection_btn_row = QtWidgets.QHBoxLayout()
        self.collection_choose_btn = QtWidgets.QPushButton("Choose...")
        self.collection_choose_btn.clicked.connect(self.on_choose_current_collection)
        self.collection_open_btn = QtWidgets.QPushButton("Open")
        self.collection_open_btn.clicked.connect(self.on_open_collection)
        self.collection_add_selected_btn = QtWidgets.QPushButton("Add selected")
        self.collection_add_selected_btn.setToolTip("Add the currently selected thumbnail(s) to the active collection")
        self.collection_add_selected_btn.clicked.connect(self.on_add_selected_thumbnails_to_collection)
        self.collection_refresh_btn = QtWidgets.QToolButton()
        self.collection_refresh_btn.setText("Refresh")
        self.collection_refresh_btn.setToolTip("Reload the collection tray from disk")
        self.collection_refresh_btn.clicked.connect(self._refresh_collection_tray)
        collection_btn_row.addWidget(self.collection_choose_btn)
        collection_btn_row.addWidget(self.collection_open_btn)
        collection_btn_row.addWidget(self.collection_add_selected_btn)
        collection_btn_row.addWidget(self.collection_refresh_btn)
        collection_btn_row.addStretch(1)
        collection_layout.addLayout(collection_btn_row)
        self.collection_tray_list = _CollectionTrayList(self)
        self.collection_tray_list.setMinimumHeight(170)
        collection_layout.addWidget(self.collection_tray_list, 1)
        self.collection_group.setVisible(True)
        self.collection_tray_window = _CollectionTrayWindow(self, self.collection_group)
        self.collection_tray_window.hide()

        details_group = QtWidgets.QGroupBox("Details")
        details_group.setCheckable(True)
        details_group.setChecked(True)
        details_layout = QtWidgets.QVBoxLayout(details_group)
        self.meta_box = QtWidgets.QTextEdit()
        self.meta_box.setReadOnly(True)
        # Keep the background transparent so HTML metadata respects the application palette
        # when switching between light and dark modes.
        try:
            self.meta_box.setStyleSheet("QTextEdit { background-color: transparent; }")
        except Exception:
            pass
        self.meta_box.setFont(meta_font)
        self.meta_box.setMinimumWidth(380)
        try:
            self.meta_box.setLineWrapMode(QtWidgets.QTextEdit.NoWrap)
        except Exception:
            pass
        self.meta_box.setPlaceholderText("File metadata / header appears when selecting a thumbnail.")
        # Metadata font size control (user preference persisted to config)
        try:
            meta_font_h = QtWidgets.QHBoxLayout()
            meta_font_h.addStretch(1)
            meta_font_h.addWidget(QtWidgets.QLabel("Font:"))
            self.meta_font_spin = QtWidgets.QSpinBox()
            self.meta_font_spin.setRange(8, 24)
            self.meta_font_spin.setValue(int(self.config.get('meta_font_size', 10)))
            self.meta_font_spin.setToolTip("Font size for the metadata panel")
            self.meta_font_spin.valueChanged.connect(self.on_meta_font_changed)
            meta_font_h.addWidget(self.meta_font_spin)
            details_layout.addLayout(meta_font_h)
        except Exception:
            pass
        details_layout.addWidget(self.meta_box, 1)
        self.activity_group = QtWidgets.QGroupBox("Activity log")
        self.activity_group.setCheckable(True)
        self.activity_group.setChecked(True)
        activity_layout = QtWidgets.QVBoxLayout(self.activity_group)
        header = QtWidgets.QHBoxLayout()
        header.addStretch(1)
        self.activity_clear_btn = QtWidgets.QToolButton()
        self.activity_clear_btn.setText("Clear")
        self.activity_clear_btn.setAutoRaise(True)
        header.addWidget(self.activity_clear_btn)
        activity_layout.addLayout(header)
        self.activity_log_box = QtWidgets.QPlainTextEdit()
        self.activity_log_box.setReadOnly(True)
        self.activity_log_box.setMaximumHeight(140)
        self.activity_log_box.setLineWrapMode(QtWidgets.QPlainTextEdit.NoWrap)
        try:
            self.activity_log_box.document().setMaximumBlockCount(500)
        except Exception:
            pass
        activity_layout.addWidget(self.activity_log_box)
        details_layout.addWidget(self.activity_group)
        self._activity_log_entries = []
        self.activity_group.toggled.connect(self.activity_log_box.setVisible)
        self.activity_clear_btn.clicked.connect(self._on_clear_activity_log)
        self.meta_box.setVisible(True)
        details_group.toggled.connect(self.meta_box.setVisible)

        frame_group = QtWidgets.QGroupBox("Folder layout (±1 µm)")
        frame_layout = QtWidgets.QVBoxLayout(frame_group)
        self.frame_map_widget = FrameMiniMap()
        self.frame_map_widget.entryClicked.connect(self._on_frame_map_clicked)
        self.frame_map_widget.entryShiftClicked.connect(self._on_frame_map_entry_shift_clicked)
        self.frame_map_widget.zoomChanged.connect(self._on_frame_map_zoom_changed)
        self.frame_map_widget.setToolTip(
            "Frame layout:\n"
            "  Click to focus a frame\n"
            "  Shift+Click hides a frame (Show all resets)\n"
            "  Mouse wheel zooms view; drag to pan\n"
            "  Toggle Show real view for channel thumbnails"
        )
        frame_layout.addWidget(self.frame_map_widget)
        zoom_row = QtWidgets.QHBoxLayout()
        zoom_row.addWidget(QtWidgets.QLabel("Zoom:"))
        self.frame_zoom_slider = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        self.frame_zoom_slider.setRange(self.FRAME_ZOOM_SLIDER_MIN, self.FRAME_ZOOM_SLIDER_MAX)  # logarithmic: 0.01x to 1e4x
        slider_val = int(self.config.get('frame_map_zoom', self.FRAME_ZOOM_SLIDER_DEFAULT))
        slider_val = self._normalize_frame_zoom_slider_value(slider_val)
        self.frame_zoom_slider.setValue(slider_val)
        self.frame_zoom_slider.valueChanged.connect(self._on_frame_zoom_changed)
        zoom_row.addWidget(self.frame_zoom_slider, 1)
        zoom_reset_btn = QtWidgets.QPushButton("Reset")
        zoom_reset_btn.setFixedWidth(60)
        zoom_reset_btn.clicked.connect(self._reset_frame_view)
        zoom_row.addWidget(zoom_reset_btn)
        frame_layout.addLayout(zoom_row)

        # Metadata font size control has been moved next to the Details header (see below)
        # (Block removed here to change placement.)
        frame_btn_row = QtWidgets.QHBoxLayout()
        self.frame_show_all_btn = QtWidgets.QPushButton("Show all frames")
        self.frame_show_all_btn.clicked.connect(self._on_frame_show_all_clicked)
        frame_btn_row.addWidget(self.frame_show_all_btn)
        self.frame_real_view_btn = QtWidgets.QPushButton("Show real view")
        self.frame_real_view_btn.setCheckable(True)
        self.frame_real_view_btn.toggled.connect(self._on_frame_real_view_toggled)
        frame_btn_row.addWidget(self.frame_real_view_btn)
        frame_btn_row.addStretch(1)
        frame_layout.addLayout(frame_btn_row)

        # Make details (metadata) + frame layout vertically resizable by the user
        self.left_splitter = QtWidgets.QSplitter(QtCore.Qt.Vertical)
        # Make handle visibly wider and styled so it's easy to find in dark mode
        self.left_splitter.setHandleWidth(10)
        self.left_splitter.setStyleSheet("""
        QSplitter::handle:vertical {
            background: rgba(255,255,255,0.06);
            margin-left: 4px;
            margin-right: 4px;
            border-top: 1px solid rgba(0,0,0,0.2);
            border-bottom: 1px solid rgba(0,0,0,0.2);
        }
        QSplitter::handle:vertical:hover {
            background: rgba(255,255,255,0.12);
        }
        """)
        # When user resizes the left/right panes, schedule a thumbnail reflow
        try:
            self.left_splitter.splitterMoved.connect(lambda pos, idx: self._thumbs_reflow_timer.start(150))
        except Exception:
            pass
        self.left_splitter.addWidget(details_group)
        self.left_splitter.addWidget(frame_group)
        self.left_splitter.setStretchFactor(0, 1)
        self.left_splitter.setStretchFactor(1, 0)
        # restore saved sizes if present, otherwise use a sensible default
        sizes = self.config.get('left_splitter_sizes')
        if isinstance(sizes, (list, tuple)) and len(sizes) >= 2:
            try:
                self.left_splitter.setSizes(list(sizes[:2]))
            except Exception:
                pass
        else:
            try:
                # default: make details area a bit larger than the layout area
                self.left_splitter.setSizes([500, 200])
            except Exception:
                pass

        def _save_left_splitter(pos, index):
            try:
                self.config['left_splitter_sizes'] = self.left_splitter.sizes()
                save_config(self.config)
            except Exception:
                pass

        self.left_splitter.splitterMoved.connect(_save_left_splitter)
        left_v.addWidget(self.left_splitter, 1)

        # Path line-edit: tooltip + clear button for convenience.
        full_path = str(self.last_dir)
        self.path_le.setText(full_path)
        self.path_le.setToolTip(full_path)
        try:
            self.path_le.setClearButtonEnabled(True)
        except Exception:
            pass

        tag_h = QtWidgets.QHBoxLayout()
        self.tag_ch_btn = QtWidgets.QPushButton("Tag as CH")
        self.tag_cc_btn = QtWidgets.QPushButton("Tag as CC")
        self.untag_btn = QtWidgets.QPushButton("Untag")
        self.auto_tag_cb = QtWidgets.QCheckBox("Auto CH/CC")
        self.auto_tag_cb.setToolTip("Auto-detect constant-height/current from topography variance")
        self.auto_tag_cb.setChecked(self.auto_detect_tags)
        
        # Purge config button
        self.purge_config_btn = QtWidgets.QPushButton('Purge config')
        tag_h.addWidget(self.purge_config_btn)
        tag_h.addWidget(self.tag_ch_btn); tag_h.addWidget(self.tag_cc_btn); tag_h.addWidget(self.untag_btn); tag_h.addWidget(self.auto_tag_cb)
        left_v.addLayout(tag_h)

        # NOTE:
        # Removed the "File channels (selected file)" inspector (list + cmap + "Show channel" button).
        # That UI duplicated functionality already provided via the thumbnails and the "Add channel view"
        # dialog. We rely on thumbnails + Add dialog going forward, so we keep the left panel slimmer.

        left_w = QtWidgets.QWidget(); left_w.setLayout(left_v)

        # Right panel with splitter for thumbnails/preview
        title_lbl = QtWidgets.QLabel("Thumbnails"); title_lbl.setFont(bold_font)
        self.scroll = QtWidgets.QScrollArea(); self.thumb_container = QtWidgets.QWidget(); self.thumb_layout = QtWidgets.QGridLayout(); self.thumb_layout.setSpacing(THUMB_LAYOUT_SPACING)
        self.scroll.setToolTip(
            "Thumbnails:\n"
            "  Shift+Click or Ctrl+Click to multi-select\n"
            "  Ctrl+Wheel to change thumbnail size\n"
            "  Right-click a frame for filters & exports"
        )
        self.thumb_container.setLayout(self.thumb_layout); self.scroll.setWidgetResizable(True); self.scroll.setWidget(self.thumb_container)
        self._thumb_viewport = self.scroll.viewport()
        self.scroll.setAcceptDrops(True)
        self.thumb_container.setAcceptDrops(True)
        self._thumb_viewport.setAcceptDrops(True)
        self._thumb_viewport.installEventFilter(self)
        self.scroll.installEventFilter(self)
        self.thumb_container.installEventFilter(self)
        try:
            self.scroll.verticalScrollBar().valueChanged.connect(lambda _: self._request_visible_thumbs())
        except Exception:
            pass
        thumbs_panel = QtWidgets.QWidget()
        self.left_w = left_w
        thumbs_panel_layout = QtWidgets.QVBoxLayout(); thumbs_panel_layout.setContentsMargins(0,0,0,0)
        thumbs_toolbar = QtWidgets.QHBoxLayout()
        thumbs_toolbar.addWidget(QtWidgets.QLabel('Sort:'))
        self.thumb_sort_combo = QtWidgets.QComboBox()
        self.thumb_sort_combo.addItems(['Name (A-Z)', 'Date (new-old)', 'Date (old-new)', 'Tag (CH-CC-U)'])
        thumbs_toolbar.addWidget(self.thumb_sort_combo)
        thumbs_toolbar.addSpacing(8)
        thumbs_toolbar.addWidget(QtWidgets.QLabel('Filter:'))
        self.thumb_filter_combo = QtWidgets.QComboBox()
        self.thumb_filter_combo.addItems(['All', 'Constant height', 'Constant current', 'Untagged', 'Matrix datasets'])
        thumbs_toolbar.addWidget(self.thumb_filter_combo)
        self.clear_thumb_list_btn = QtWidgets.QPushButton("Clear thumbnails")
        self.clear_thumb_list_btn.setToolTip("Remove the current thumbnail session and start fresh")
        self.clear_thumb_list_btn.clicked.connect(self.clear_loaded_images)
        thumbs_toolbar.addWidget(self.clear_thumb_list_btn)
        thumbs_toolbar.addSpacing(8)
        self.matrix_summary_label = QtWidgets.QLabel("")
        self.matrix_summary_label.setObjectName("matrixSummaryLabel")
        self.matrix_summary_label.setVisible(False)
        self.matrix_summary_label.setCursor(QtCore.Qt.PointingHandCursor)
        self.matrix_summary_label.setStyleSheet(
            "#matrixSummaryLabel {"
            " padding: 2px 10px; border-radius: 12px; "
            " background-color: rgba(100, 180, 255, 0.18); color: #e6f2ff; "
            " border: 1px solid rgba(120, 200, 255, 0.65); font-weight: 600;"
            "}"
        )
        self.matrix_summary_label.mousePressEvent = lambda event: self._focus_first_matrix_dataset()
        thumbs_toolbar.addWidget(self.matrix_summary_label)
        thumbs_toolbar.addStretch(1)
        self.unit_display_cb = QtWidgets.QCheckBox("SI")
        self.unit_display_cb.setChecked(self.display_units_si)
        self.unit_display_cb.setToolTip("Show SI units in preview annotations")
        self.unit_relative_cb = QtWidgets.QCheckBox("Zero")
        self.unit_relative_cb.setChecked(self.display_units_relative)
        self.unit_relative_cb.setToolTip("Display values relative to the current zero/reference")
        self.relative_axes_cb = QtWidgets.QCheckBox("Axes")
        self.relative_axes_cb.setChecked(self.relative_axes)
        self.relative_axes_cb.setToolTip("Use relative axes in the preview")
        # Keep the thumbnails header simple so the preview workspace owns the channel workflow.
        header_h = QtWidgets.QHBoxLayout()
        header_h.setContentsMargins(0,0,0,0)
        header_h.setSpacing(8)
        header_h.addWidget(title_lbl)
        header_h.addStretch(1)
        thumbs_panel_layout.addLayout(header_h)
        thumbs_panel_layout.addWidget(self.scroll, 1)
        thumbs_panel_layout.addLayout(thumbs_toolbar)

        # restore sort/filter from config if present
        try:
            sort_label = self.config.get('thumb_sort', 'Name (A-Z)')
            if sort_label in [self.thumb_sort_combo.itemText(i) for i in range(self.thumb_sort_combo.count())]:
                self.thumb_sort_combo.setCurrentText(sort_label)
            filt_label = self.config.get('thumb_filter', 'All')
            if filt_label in [self.thumb_filter_combo.itemText(i) for i in range(self.thumb_filter_combo.count())]:
                self.thumb_filter_combo.setCurrentText(filt_label)
        except Exception:
            pass
        thumbs_panel.setLayout(thumbs_panel_layout)

        preview_panel = QtWidgets.QWidget()
        preview_panel_layout = QtWidgets.QVBoxLayout(); preview_panel_layout.setContentsMargins(0,0,0,0); preview_panel_layout.setSpacing(6)
        self.preview_workspace_frame = QtWidgets.QFrame()
        self.preview_workspace_frame.setObjectName("previewWorkspaceFrame")
        preview_workspace_layout = QtWidgets.QVBoxLayout(self.preview_workspace_frame)
        preview_workspace_layout.setContentsMargins(10, 8, 10, 8)
        preview_workspace_layout.setSpacing(6)
        preview_header = QtWidgets.QHBoxLayout()
        preview_header.setContentsMargins(0, 0, 0, 0)
        preview_header.setSpacing(8)
        self.preview_title_label = QtWidgets.QLabel("Preview")
        self.preview_title_label.setFont(bold_font)
        preview_header.addWidget(self.preview_title_label)
        self.channel_controls_widget = QtWidgets.QWidget()
        self.channel_controls_widget.setLayout(controls_h)
        self.channel_controls_widget.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        preview_header.addWidget(self.channel_controls_widget, 1)
        # Dock/lock controls
        self.preview_lock_cb = QtWidgets.QCheckBox("Lock")
        self.preview_lock_cb.setChecked(self.preview_locked)
        self.preview_lock_cb.setToolTip("Lock preview inside the main window")
        self.preview_lock_cb.toggled.connect(self.on_preview_lock_toggled)
        self.preview_detach_btn = QtWidgets.QToolButton()
        self.preview_detach_btn.clicked.connect(self.on_toggle_preview_detach)
        self._update_preview_detach_button()
        self.scale_bar_cb = QtWidgets.QCheckBox("Bar")
        self.scale_bar_cb.setChecked(bool(self.config.get("show_scale_bar", False)))
        self.scale_bar_cb.setToolTip("Show the scale bar in preview and pop-outs")
        self.preview_hist_btn = QtWidgets.QToolButton()
        self.preview_hist_btn.setText("Histogram")
        self.preview_hist_btn.setToolTip("Show histogram and adjust display range")
        self.preview_hist_btn.setPopupMode(QtWidgets.QToolButton.MenuButtonPopup)
        self.preview_hist_menu = QtWidgets.QMenu(self.preview_hist_btn)
        self.preview_hist_menu.addAction("Histogram...", lambda: self._open_histogram_dialog(self.preview_canvas))
        self.preview_hist_menu.addAction("Auto (1–99%)", lambda: self._auto_contrast(self.preview_canvas))
        self.preview_hist_menu.addAction("Reset range", lambda: self._reset_contrast(self.preview_canvas))
        self.show_preview_title = bool(self.config.get('show_preview_title', True))
        act_title = self.preview_hist_menu.addAction("Show title/date")
        act_title.setCheckable(True)
        act_title.setChecked(self.show_preview_title)
        act_title.triggered.connect(self._on_toggle_preview_title)
        self.preview_hist_btn.setMenu(self.preview_hist_menu)
        self.preview_hist_btn.clicked.connect(lambda _: self._open_histogram_dialog(self.preview_canvas))
        self.preview_adjust_btn = QtWidgets.QToolButton()
        self.preview_adjust_btn.setText("Crop/Rotate")
        self.preview_adjust_btn.setToolTip("Open crop, rotate, flip, clipping, gamma, and colormap controls")
        self.preview_adjust_btn.clicked.connect(self.on_adjust_image)
        self.preview_adjust_btn.setEnabled(False)
        self.preview_molecules_toggle_btn = QtWidgets.QToolButton()
        self.preview_molecules_toggle_btn.setText("Mol")
        self.preview_molecules_toggle_btn.setCheckable(True)
        self.preview_molecules_toggle_btn.setChecked(self.show_molecules)
        self.preview_molecules_toggle_btn.setToolTip(
            "Show or hide molecular overlays on the preview and pop-outs. "
            "Click a molecule, then use X/Y/Z to rotate; Shift+X/Y/Z rotates the opposite way."
        )
        self.preview_molecules_toggle_btn.toggled.connect(self.on_show_molecules_toggled)
        self.preview_grid_toggle_btn = QtWidgets.QToolButton()
        self.preview_grid_toggle_btn.setText("Grid")
        self.preview_grid_toggle_btn.setCheckable(True)
        self.preview_grid_toggle_btn.setChecked(self.detail_grid_view)
        self.preview_grid_toggle_btn.setToolTip("Show or hide the detail grid overlay")
        self.preview_grid_toggle_btn.toggled.connect(self.on_detail_grid_toggled)
        self.toolbar_display_btn = QtWidgets.QToolButton()
        self.toolbar_display_btn.setText("Display")
        self.toolbar_display_btn.setPopupMode(QtWidgets.QToolButton.InstantPopup)
        self.toolbar_display_btn.setToolTip("Preview and overlay display options")
        self.toolbar_display_btn.setMenu(main_window_layout._ensure_display_menu(self))
        self.toolbar_load_mol_btn = QtWidgets.QLabel()
        self.toolbar_load_mol_btn.setFixedSize(44, 28)
        self.toolbar_load_mol_btn.setAlignment(QtCore.Qt.AlignCenter)
        self.toolbar_load_mol_btn.setToolTip("Load a molecular structure overlay (XYZ, PDB, MOL)")
        self.toolbar_load_mol_btn.setCursor(QtCore.Qt.PointingHandCursor)
        self._molecule_pixmap_size = QtCore.QSize(32, 18)
        self.toolbar_load_mol_btn.mousePressEvent = lambda event: self.on_load_molecule()
        self.toolbar_dark_btn = QtWidgets.QPushButton("Dark")
        self.toolbar_dark_btn.setCheckable(True)
        self.toolbar_dark_btn.setToolTip("Toggle dark mode")
        self.toolbar_dark_btn.setMinimumWidth(64)
        self.toolbar_dark_btn.setFixedHeight(28)
        self.toolbar_dark_btn.toggled.connect(self.on_dark_mode_toggled)
        preview_workspace_layout.addLayout(preview_header)

        self.thumb_cmap_label = QtWidgets.QLabel("Thumb")
        self.thumb_cmap_label.setToolTip("Colormap used for thumbnails")
        self.preview_cmap_label = QtWidgets.QLabel("Preview")
        self.preview_cmap_label.setToolTip("Colormap used for the preview")
        self.thumb_cmap_combo.setMinimumWidth(120)
        self.preview_cmap_combo.setMinimumWidth(120)
        preview_state_row = QtWidgets.QHBoxLayout()
        preview_state_row.setContentsMargins(0, 0, 0, 0)
        preview_state_row.setSpacing(8)
        preview_state_row.addWidget(self.thumb_cmap_label)
        preview_state_row.addWidget(self.thumb_cmap_combo)
        preview_state_row.addSpacing(8)
        preview_state_row.addWidget(self.preview_cmap_label)
        preview_state_row.addWidget(self.preview_cmap_combo)
        preview_state_row.addStretch(1)
        preview_workspace_layout.addLayout(preview_state_row)
        preview_panel_layout.addWidget(self.preview_workspace_frame)

        self.quick_crop_controls = QtWidgets.QWidget()
        quick_layout = QtWidgets.QVBoxLayout(self.quick_crop_controls)
        quick_layout.setContentsMargins(0, 0, 0, 0)
        quick_layout.setSpacing(4)
        quick_toggle_row = QtWidgets.QHBoxLayout()
        quick_toggle_row.setContentsMargins(0, 0, 0, 0)
        quick_toggle_row.setSpacing(6)
        self.quick_crop_btn = QtWidgets.QPushButton("Crop template: Off")
        self.quick_crop_btn.setCheckable(True)
        self.quick_crop_btn.setToolTip(
            "Enable repeated cropping from the current template (Ctrl+Shift+C). "
            "Click the preview to apply the template. Shift+drag draws a manual crop; "
            "Ctrl+Shift+drag forces a square manual crop."
        )
        quick_toggle_row.addWidget(self.quick_crop_btn)
        self.quick_crop_edit_btn = QtWidgets.QToolButton()
        self.quick_crop_edit_btn.setText("Edit frame")
        self.quick_crop_edit_btn.setCheckable(True)
        self.quick_crop_edit_btn.setToolTip(
            "Move, resize, and rotate the current crop template on the preview (Ctrl+E)."
        )
        quick_toggle_row.addWidget(self.quick_crop_edit_btn)
        quick_toggle_row.addStretch(1)
        quick_layout.addLayout(quick_toggle_row)
        self.quick_crop_detail_widget = QtWidgets.QWidget()
        quick_detail_layout = QtWidgets.QVBoxLayout(self.quick_crop_detail_widget)
        quick_detail_layout.setContentsMargins(0, 0, 0, 0)
        quick_detail_layout.setSpacing(4)
        quick_template_row = QtWidgets.QHBoxLayout()
        quick_template_row.setContentsMargins(0, 0, 0, 0)
        quick_template_row.setSpacing(6)
        quick_template_row.addWidget(QtWidgets.QLabel("W"))
        self.quick_crop_real_width_spin = QtWidgets.QDoubleSpinBox()
        self.quick_crop_real_width_spin.setRange(0.01, 10000.0)
        self.quick_crop_real_width_spin.setDecimals(3)
        self.quick_crop_real_width_spin.setSingleStep(0.1)
        self.quick_crop_real_width_spin.setFixedWidth(72)
        self.quick_crop_real_width_spin.setToolTip("Template width in real-space units.")
        self.quick_crop_real_width_spin.setValue(5.0)
        self._quick_crop_aspect = 1.0
        self._quick_crop_last_real_size = [self.quick_crop_real_width_spin.value(), 5.0]
        quick_template_row.addWidget(self.quick_crop_real_width_spin)
        quick_template_row.addWidget(QtWidgets.QLabel("H"))
        self.quick_crop_real_height_spin = QtWidgets.QDoubleSpinBox()
        self.quick_crop_real_height_spin.setRange(0.01, 10000.0)
        self.quick_crop_real_height_spin.setDecimals(3)
        self.quick_crop_real_height_spin.setSingleStep(0.1)
        self.quick_crop_real_height_spin.setFixedWidth(72)
        self.quick_crop_real_height_spin.setToolTip("Template height in real-space units.")
        self.quick_crop_real_height_spin.setValue(5.0)
        self._quick_crop_last_real_size = [self.quick_crop_real_width_spin.value(),
                                           self.quick_crop_real_height_spin.value()]
        self._quick_crop_aspect = self._quick_crop_last_real_size[0] / max(0.001, self._quick_crop_last_real_size[1])
        quick_template_row.addWidget(self.quick_crop_real_height_spin)
        self.quick_crop_real_unit_lbl = QtWidgets.QLabel("nm")
        quick_template_row.addWidget(self.quick_crop_real_unit_lbl)
        quick_template_row.addWidget(QtWidgets.QLabel("Aspect"))
        self.quick_crop_aspect_combo = QtWidgets.QComboBox()
        self.quick_crop_aspect_combo.addItem("Free", "free")
        self.quick_crop_aspect_combo.addItem("Keep ratio", "keep")
        self.quick_crop_aspect_combo.addItem("Square", "square")
        self.quick_crop_aspect_combo.setToolTip(
            "Free: width and height change independently. "
            "Keep ratio: template size edits preserve the current ratio. "
            "Square: the template stays square, and Shift+drag manual crops stay square "
            "while crop-template mode is on."
        )
        aspect_index = max(0, self.quick_crop_aspect_combo.findData(self.quick_crop_aspect_mode))
        self.quick_crop_aspect_combo.setCurrentIndex(aspect_index)
        quick_template_row.addWidget(self.quick_crop_aspect_combo)
        self.quick_crop_real_px_info_lbl = QtWidgets.QLabel("")
        self.quick_crop_real_px_info_lbl.setMinimumWidth(0)
        self.quick_crop_real_px_info_lbl.setSizePolicy(QtWidgets.QSizePolicy.Maximum, QtWidgets.QSizePolicy.Preferred)
        quick_template_row.addWidget(self.quick_crop_real_px_info_lbl)
        self.quick_crop_controller = QuickCropController(self)
        self.thumbnail_controller = ThumbnailController(self)
        self.quick_crop_actions_btn = QtWidgets.QToolButton()
        self.quick_crop_actions_btn.setText("Actions")
        self.quick_crop_actions_btn.setPopupMode(QtWidgets.QToolButton.InstantPopup)
        self.quick_crop_actions_btn.setToolTip(
            "Crop-template history, export, and pop-out management."
        )
        self.quick_crop_actions_menu = QtWidgets.QMenu(self.quick_crop_actions_btn)
        self.quick_crop_undo_act = self.quick_crop_actions_menu.addAction("Undo latest crop")
        self.quick_crop_undo_act.setToolTip("Undo latest crop (Ctrl+Z)")
        self.quick_crop_undo_act.triggered.connect(self.quick_crop_controller.undo_last_crop)
        self.quick_crop_close_act = self.quick_crop_actions_menu.addAction("Close latest pop-out")
        self.quick_crop_close_act.setToolTip("Close the latest quick-crop pop-out (Ctrl+Shift+W)")
        self.quick_crop_close_act.triggered.connect(self.quick_crop_controller.close_latest_popup)
        self.quick_crop_clear_act = self.quick_crop_actions_menu.addAction("Clear history")
        self.quick_crop_clear_act.setToolTip("Clear crop history markers and pop-outs")
        self.quick_crop_clear_act.triggered.connect(self.quick_crop_controller.clear_history)
        self.quick_crop_export_act = self.quick_crop_actions_menu.addAction("Export selected crops")
        self.quick_crop_export_act.setToolTip("Export the selected crops (Shift+click) as images")
        self.quick_crop_export_act.triggered.connect(self.quick_crop_controller.export_selected_crops)
        self.quick_crop_actions_menu.addSeparator()
        self.quick_crop_tile_act = self.quick_crop_actions_menu.addAction("Tile pop-outs")
        self.quick_crop_tile_act.setToolTip("Arrange all open pop-out windows on screen")
        self.quick_crop_tile_act.triggered.connect(self.on_arrange_popouts)
        self.quick_crop_minimize_act = self.quick_crop_actions_menu.addAction("Minimize pop-outs")
        self.quick_crop_minimize_act.setToolTip("Minimize all open pop-out windows (Ctrl+Shift+M)")
        self.quick_crop_minimize_act.triggered.connect(self.on_minimize_popouts)
        self.quick_crop_actions_btn.setMenu(self.quick_crop_actions_menu)
        quick_template_row.addStretch(1)
        quick_template_row.addWidget(self.quick_crop_actions_btn)
        quick_detail_layout.addLayout(quick_template_row)
        self.quick_crop_hint_lbl = QtWidgets.QLabel("")
        self.quick_crop_hint_lbl.setWordWrap(True)
        self.quick_crop_hint_lbl.setMinimumWidth(0)
        self.quick_crop_hint_lbl.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
        quick_detail_layout.addWidget(self.quick_crop_hint_lbl)
        quick_layout.addWidget(self.quick_crop_detail_widget)
        preview_panel_layout.addWidget(self.quick_crop_controls)
        self.quick_crop_btn.clicked.connect(lambda: self._set_quick_crop_mode(not self.quick_crop_mode))
        self.quick_crop_edit_btn.toggled.connect(self._on_quick_crop_edit_toggled)
        self.quick_crop_aspect_combo.currentIndexChanged.connect(lambda _=None: self._on_quick_crop_aspect_mode_changed())
        self.quick_crop_real_width_spin.valueChanged.connect(lambda _=None: self.quick_crop_controller.on_real_spin_changed(self.quick_crop_real_width_spin))
        self.quick_crop_real_height_spin.valueChanged.connect(lambda _=None: self.quick_crop_controller.on_real_spin_changed(self.quick_crop_real_height_spin))
        self.quick_crop_detail_widget.setVisible(bool(self.quick_crop_mode))

        self.crop_history_panel = QtWidgets.QWidget()
        crop_hist_layout = QtWidgets.QVBoxLayout(self.crop_history_panel)
        crop_hist_layout.setContentsMargins(0, 0, 0, 0)
        crop_hist_layout.setSpacing(4)
        crop_history_header = QtWidgets.QHBoxLayout()
        crop_history_header.setContentsMargins(0, 0, 0, 0)
        crop_history_header.setSpacing(8)
        self.crop_history_label = QtWidgets.QLabel("Crop history")
        crop_history_header.addWidget(self.crop_history_label)
        crop_history_header.addStretch(1)
        crop_hist_layout.addLayout(crop_history_header)
        self.crop_history_scroll = QtWidgets.QScrollArea()
        self.crop_history_scroll.setWidgetResizable(True)
        self.crop_history_scroll.setFixedHeight(180)
        self.crop_history_content = QtWidgets.QWidget()
        self.crop_history_layout = QtWidgets.QVBoxLayout(self.crop_history_content)
        self.crop_history_layout.setContentsMargins(0, 0, 0, 0)
        self.crop_history_layout.setSpacing(6)
        self.crop_history_layout.addStretch(1)
        self.crop_history_scroll.setWidget(self.crop_history_content)
        crop_hist_layout.addWidget(self.crop_history_scroll)
        preview_panel_layout.addWidget(self.crop_history_panel)
        # Place the lower controls (modes + context actions) directly under the Preview header
        self.lower_control_frame = self._create_lower_controls()
        preview_panel_layout.addWidget(self.lower_control_frame)

        self.preview_canvas = MultiPreviewCanvas(self, figsize=(6,5))
        self.preview_canvas.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.preview_canvas.setMinimumWidth(240)
        self.preview_canvas.setToolTip(
            "Preview area:\n"
            "  Right-click for copy/save/PowerPoint options\n"
            "  Enable 'Measure profile' for line sampling\n"
            "  Ctrl+C copies the displayed preview as PNG"
        )
        try:
            self.preview_canvas._undo_suspend_depth += 1
        except Exception:
            pass
        try:
            self.preview_canvas.set_show_title(self.show_preview_title)
        except Exception:
            pass
        try:
            self.preview_canvas.set_show_molecules(self.show_molecules)
        except Exception:
            pass
        try:
            self.preview_canvas.set_show_molecule_gizmo(self.show_molecule_gizmo)
        except Exception:
            pass
        try:
            self.preview_canvas.set_show_acquisition_overlay(self.show_acquisition_overlay)
        except Exception:
            pass
        try:
            self.preview_canvas.set_profile_callback(self._on_profile_updated)
        except Exception:
            pass
        try:
            if hasattr(self.preview_canvas, "set_profile_highlight_callback"):
                self.preview_canvas.set_profile_highlight_callback(self._on_canvas_overlay_highlight)
        except Exception:
            pass
        try:
            self.preview_canvas.set_profile_label_mode(self.profile_label_mode)
        except Exception:
            pass
        self.preview_canvas.set_copy_feedback_handler(self._on_view_copied)
        try:
            self.preview_canvas.set_plot_font_family_callback(lambda fam: self.set_plot_font_family(fam))
            if hasattr(self.preview_canvas, "set_plot_typography"):
                self.preview_canvas.set_plot_typography(
                    family=self._plot_font_family,
                    bold=self._plot_font_bold,
                    italic=self._plot_font_italic,
                    underline=self._plot_font_underline,
                )
            else:
                self.preview_canvas.set_plot_font_family(self._plot_font_family)
        except Exception:
            pass
        try:
            self.preview_canvas.set_molecule_palette(self.molecule_palette, notify=False)
            self.preview_canvas.set_molecule_palette_callback(self._on_molecule_palette_changed)
        except Exception:
            pass
        preview_panel_layout.addWidget(self.preview_canvas, 1)
        self.preview_value_label = QtWidgets.QLabel("Value: --")
        preview_panel_layout.addWidget(self.preview_value_label)
        self.angle_value_label = QtWidgets.QLabel("Angle: --")
        preview_panel_layout.addWidget(self.angle_value_label)
        preview_panel.setLayout(preview_panel_layout)
        self.preview_canvas.set_value_callback(self._on_preview_value)
        self.preview_canvas.set_spectra_click_callback(self._on_preview_spec_click)
        self.preview_canvas.set_crop_callback(lambda v, c=self.preview_canvas: self._on_preview_crop(v, c))
        self.preview_canvas.set_virtual_copy_callback(self._create_virtual_copy_from_popup_view)
        self.preview_canvas.set_double_click_callback(
            lambda v=None: self._spawn_preview_popup(
                [self._copy_view_for_popup(v)] if v else [],
                title=self._friendly_view_title(v, default="Preview copy") if v else "Preview copy",
            )
        )
        self.preview_canvas.set_filter_menu_callback(
            lambda menu, view, c=self.preview_canvas: self._populate_canvas_filter_menu(menu, c, view)
        )
        self.preview_canvas.set_histogram_dialog_callback(lambda c: self._open_histogram_dialog(c))
        self.preview_canvas.set_histogram_auto_callback(lambda c: self._auto_contrast(c))
        self.preview_canvas.set_histogram_reset_callback(lambda c: self._reset_contrast(c))
        self.preview_canvas.set_compare_menu_callback(
            lambda action, view, c=self.preview_canvas: self.on_compare_menu_action(action, view, c),
            state_cb=self.compare_menu_state,
        )
        if hasattr(self.preview_canvas, "set_collection_menu_callback"):
            self.preview_canvas.set_collection_menu_callback(
                lambda action, view, c=self.preview_canvas: self.collection_controller.handle_canvas_menu_action(action, view, c),
                help_cb=self.on_collection_help,
            )
        self.preview_canvas.set_stp_export_callback(self._export_view_as_stp)
        self.preview_canvas.set_window_arrange_callback(self.on_arrange_popouts)
        self.preview_canvas.set_window_minimize_callback(self.on_minimize_popouts)
        self.preview_canvas.set_window_restore_callback(self.on_restore_popouts)
        self.preview_canvas.set_window_close_callback(self.on_close_popouts)
        self.preview_canvas.set_fixed_crop_history_callback(self._on_fixed_crop_history_updated)
        # Seed molecule recents from config and listen for updates
        try:
            if getattr(self, "recent_molecules", None):
                self.preview_canvas._recent_molecule_paths = list(self.recent_molecules)
                MultiPreviewCanvas._RECENT_MOLECULES = list(self.recent_molecules)
        except Exception:
            pass
        try:
            self.preview_canvas.set_recent_molecule_callback(self._on_recent_molecules_updated)
        except Exception:
            pass
        self.preview_canvas.enable_scale_bar(self.scale_bar_cb.isChecked())
        self.preview_canvas.enable_fixed_crop_quick_mode(self.quick_crop_mode)
        self.preview_canvas.show_fixed_crop_template(self.show_crop_template_overlay)
        self.preview_canvas.show_fixed_crop_history(True)
        try:
            self.preview_canvas._undo_suspend_depth = max(0, getattr(self.preview_canvas, "_undo_suspend_depth", 0) - 1)
        except Exception:
            pass
        self.preview_canvas.set_views_callback(
            lambda _=None, c=self.preview_canvas: self._on_preview_canvas_state_changed(c)
        )
        if self.canvas_display_options:
            self._apply_canvas_display_options(
                self.canvas_display_options,
                source_canvas=self.preview_canvas,
                persist=False,
            )
        self.quick_crop_toggle_shortcut = QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+Shift+C"), self)
        self.quick_crop_toggle_shortcut.setContext(QtCore.Qt.WidgetWithChildrenShortcut)
        self.quick_crop_toggle_shortcut.activated.connect(lambda: self._set_quick_crop_mode(not self.quick_crop_mode))
        self.quick_crop_undo_shortcut = QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+Z"), self)
        self.quick_crop_undo_shortcut.setContext(QtCore.Qt.WidgetWithChildrenShortcut)
        self.quick_crop_undo_shortcut.activated.connect(self._on_global_undo_requested)
        self.quick_crop_close_shortcut = QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+Shift+W"), self)
        self.quick_crop_close_shortcut.setContext(QtCore.Qt.WidgetWithChildrenShortcut)
        self.quick_crop_close_shortcut.activated.connect(self.quick_crop_controller.close_latest_popup)
        self.quick_crop_real_shortcut = QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+Shift+R"), self)
        self.quick_crop_real_shortcut.setContext(QtCore.Qt.WidgetWithChildrenShortcut)
        self.quick_crop_real_shortcut.activated.connect(self.quick_crop_controller.apply_template_from_controls)
        self.quick_crop_template_shortcut = QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+Shift+T"), self)
        self.quick_crop_template_shortcut.setContext(QtCore.Qt.WidgetWithChildrenShortcut)
        self.quick_crop_template_shortcut.activated.connect(lambda: self.on_show_crop_template_overlay_toggled(not self.show_crop_template_overlay))
        self.popups_recall_shortcut = QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+Shift+P"), self)
        self.popups_recall_shortcut.setContext(QtCore.Qt.ApplicationShortcut)
        self.popups_recall_shortcut.activated.connect(self.on_recall_popouts)
        self.quick_crop_minimize_shortcut = QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+Shift+M"), self)
        self.quick_crop_minimize_shortcut.setContext(QtCore.Qt.ApplicationShortcut)
        self.quick_crop_minimize_shortcut.activated.connect(self.on_minimize_popouts)
        self.save_session_shortcut = QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+S"), self)
        self.save_session_shortcut.setContext(QtCore.Qt.WidgetWithChildrenShortcut)
        self.save_session_shortcut.activated.connect(self.on_save_session)
        self._apply_detail_view_theme()
        # apply saved metadata font size
        try:
            font = self.meta_box.font()
            font.setPointSize(int(self.config.get('meta_font_size', 10)))
            self.meta_box.setFont(font)
        except Exception:
            pass
        # open_canvas handled in toolbar

        # Store for layout toggling
        self._thumbs_panel = thumbs_panel
        self._preview_panel = preview_panel

        self._right_splitter = QtWidgets.QSplitter(QtCore.Qt.Vertical)
        self._right_splitter.addWidget(self._thumbs_panel)
        self._right_splitter.addWidget(self._preview_panel)
        self._right_splitter.setStretchFactor(0, 3)
        self._right_splitter.setStretchFactor(1, 2)
        self._right_container = QtWidgets.QWidget()
        right_layout = QtWidgets.QVBoxLayout(); right_layout.setContentsMargins(0,0,0,0)
        right_layout.addWidget(self._right_splitter, 1)
        self._right_container.setLayout(right_layout)

        main_splitter = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        main_splitter.addWidget(left_w)
        main_splitter.addWidget(self._thumbs_panel)
        main_splitter.addWidget(self._preview_panel)
        main_splitter.setHandleWidth(8)
        # left = inspector, middle = thumbnails, right = preview stack
        main_splitter.setStretchFactor(0, 1)
        main_splitter.setStretchFactor(1, 2)
        main_splitter.setStretchFactor(2, 3)
        self.main_splitter = main_splitter
        self._layout_mode = "columns"
        try:
            self.preview_canvas.set_view_layout("stacked")
        except Exception:
            pass
        self._layout_sizes = {}
        self._set_quick_crop_mode(self.quick_crop_mode, save=False)
        if hasattr(self, "quick_crop_controller"):
            self.quick_crop_controller.refresh_history_panel()

        # Prevent panes from collapsing to zero width when the user drags the splitter.
        # This avoids the left inspector disappearing when the user expands the thumbnails.
        try:
            main_splitter.setCollapsible(0, False)
            main_splitter.setCollapsible(1, True)
            main_splitter.setCollapsible(2, True)
        except Exception:
            # older PyQt versions may not support setCollapsible; ignore safely
            pass

        # Ensure the left widget cannot shrink below a useful width
        try:
            left_w.setMinimumWidth(360)
        except Exception:
            pass
        try:
            thumbs_panel.setMinimumWidth(140)
        except Exception:
            pass
        try:
            preview_panel.setMinimumWidth(220)
        except Exception:
            pass
        try:
            thumbs_panel.setSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Expanding)
            preview_panel.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        except Exception:
            pass

        # Set reasonable initial sizes (left, right). Adjust these numbers to taste.
        try:
            main_splitter.setSizes(list(MAIN_SPLITTER_SIZES_COLUMNS))
        except Exception:
            pass

        # Responsive thumbnail reflow: debounce splitter moves & window resizes to avoid
        # repeated rebuilds while the user is dragging.
        self._thumbs_reflow_timer = QtCore.QTimer(self)
        self._thumbs_reflow_timer.setSingleShot(True)
        self._thumbs_reflow_timer.timeout.connect(lambda: self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex()))
        try:
            main_splitter.splitterMoved.connect(lambda pos, idx: self._thumbs_reflow_timer.start(150))
        except Exception:
            # older Qt versions may not expose splitterMoved the same way; ignore
            pass

        toolbar = self._create_toolbar()
        session_activity = self._create_session_activity_strip()
        container_layout = QtWidgets.QVBoxLayout()
        container_layout.setContentsMargins(0, 0, 0, 0)
        self.shortcuts_panel = self._create_shortcuts_panel()
        container_layout.addWidget(self.shortcuts_panel)
        if toolbar is not None:
            container_layout.addWidget(toolbar)
        container_layout.addWidget(session_activity)
        container_layout.addWidget(main_splitter)
        self.setLayout(container_layout)
        self._set_shortcuts_panel_visible(self.show_shortcuts_panel, remember=False)
        self._refresh_deferred_popup_ui()
        self._autosave_timer = QtCore.QTimer(self)
        self._autosave_timer.setSingleShot(False)
        self._autosave_timer.timeout.connect(self._on_autosave_timer)
        self._refresh_session_recovery_ui()
        self._refresh_autosave_timer()
        QtCore.QTimer.singleShot(900, self._maybe_offer_recovery_session)
        # Bootstrap the last-used image folder on startup so images and
        # spectroscopies appear together instead of rendering spectroscopy
        # miniatures alone before any image folder has been loaded.
        QtCore.QTimer.singleShot(200, self._startup_load_initial_workspace)

        # signals
        self.open_btn.clicked.connect(self.open_folder_dialog)
        self.path_le.returnPressed.connect(self.open_folder_by_path)
        self.spec_folder_btn.clicked.connect(self.on_spec_folder_browse)
        self.spec_folder_le.returnPressed.connect(self.on_spec_folder_entered)
        self.channel_prev_btn.clicked.connect(lambda: self._step_channel(-1))
        self.channel_next_btn.clicked.connect(lambda: self._step_channel(1))
        self.channel_dropdown.currentIndexChanged.connect(self.on_channel_dropdown_changed)
        self.thumb_cmap_combo.currentIndexChanged.connect(self.on_thumb_cmap_changed)
        self.preview_cmap_combo.currentIndexChanged.connect(self.on_preview_cmap_changed)
        self.thumb_sort_combo.currentIndexChanged.connect(self.on_thumb_sort_changed)
        self.thumb_filter_combo.currentIndexChanged.connect(self.on_thumb_filter_changed)
        self.unit_display_cb.toggled.connect(self.on_unit_display_toggled)
        self.unit_relative_cb.toggled.connect(self.on_unit_relative_toggled)
        self.relative_axes_cb.toggled.connect(self.on_relative_axes_toggled)
        self.scale_bar_cb.toggled.connect(self.on_scale_bar_toggled)
        # no size slider callback
        # inspector widgets removed -> no connections required here
        self.add_view_btn.clicked.connect(self.on_add_view)
        self.clear_views_btn.clicked.connect(self.on_clear_views)
        self.measure_profile_btn.clicked.connect(self._on_start_profile)
        self.measure_angle_btn.clicked.connect(self._on_start_angle)
        self.exit_profile_btn.clicked.connect(self._on_exit_profile_mode)
        self.clear_profile_btn.clicked.connect(self._on_clear_profile_measurement)
        self.show_profile_window_btn.clicked.connect(self._on_show_profile_window)
        show_spectra_cb = getattr(self, "show_spectra_cb", None)
        if show_spectra_cb is not None:
            show_spectra_cb.toggled.connect(self.on_show_preview_spectra_toggled)
        if getattr(self, "spectro_thumbnail_markers_cb", None) is not None:
            self.spectro_thumbnail_markers_cb.toggled.connect(self.on_show_spectra_toggled)
        if getattr(self, "spectro_preview_markers_cb", None) is not None:
            self.spectro_preview_markers_cb.toggled.connect(self.on_show_preview_spectra_toggled)
        if getattr(self, "spectro_miniatures_cb", None) is not None:
            self.spectro_miniatures_cb.toggled.connect(self.on_show_spectro_miniatures_toggled)
        if getattr(self, "grid_as_matrix_cb", None) is not None:
            self.grid_as_matrix_cb.toggled.connect(self.on_spectro_grid_as_matrix_toggled)
        if getattr(self, "force_single_cb", None) is not None:
            self.force_single_cb.toggled.connect(self.on_spectro_force_single_toggled)
        self.clear_spec_selection_btn.clicked.connect(self.on_clear_spec_selection)
        self.tag_ch_btn.clicked.connect(lambda: self.on_manual_tag('constant-height'))
        self.tag_cc_btn.clicked.connect(lambda: self.on_manual_tag('constant-current'))
        self.untag_btn.clicked.connect(lambda: self.on_manual_tag(None))
        self.auto_tag_cb.toggled.connect(self._on_toggle_auto_tags)

        try:
            self.purge_config_btn.clicked.connect(self._on_purge_config)
        except Exception:
            pass
        # apply initial dark mode palette
        try:
            self._apply_dark_mode(self.dark_mode)
        except Exception:
            pass
        self._update_toolbar_actions(False)
        self._sync_channel_nav_buttons()
        self._init_mode_shortcuts()
        try:
            log_emitter.message_logged.connect(self._append_activity_log)
        except Exception:
            pass

    def closeEvent(self, event):
        try:
            self._save_recovery_snapshot(reason="close")
        except Exception:
            pass
        try:
            self._close_workspace_windows(record_history=False, include_canvas=True)
        except Exception:
            pass
        try:
            window = getattr(self, "collection_tray_window", None)
            if window is not None:
                window.hide()
        except Exception:
            pass
        super().closeEvent(event)

    def _startup_load_initial_workspace(self):
        """Load the last-used image folder on startup when available."""
        if getattr(self, "files", None):
            try:
                self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())
            except Exception:
                pass
            return
        if "last_dir" not in getattr(self, "config", {}):
            return
        try:
            folder = Path(self.last_dir)
        except Exception:
            return
        if not folder.exists() or not folder.is_dir():
            return
        has_images = False
        try:
            has_images = any(
                child.is_file() and child.suffix.lower() in {".txt", ".sxm"}
                for child in folder.iterdir()
            )
        except Exception:
            has_images = False
        if has_images:
            self.load_folder(folder)

    def _apply_dark_mode(self, enabled: bool):
        app = QtWidgets.QApplication.instance()
        if app is None:
            return
        if enabled:
            app.setStyle('Fusion')
            palette = QtGui.QPalette()
            palette.setColor(QtGui.QPalette.Window, QtGui.QColor(53,53,53))
            palette.setColor(QtGui.QPalette.WindowText, QtCore.Qt.white)
            palette.setColor(QtGui.QPalette.Base, QtGui.QColor(35,35,35))
            palette.setColor(QtGui.QPalette.AlternateBase, QtGui.QColor(53,53,53))
            palette.setColor(QtGui.QPalette.ToolTipBase, QtCore.Qt.white)
            palette.setColor(QtGui.QPalette.ToolTipText, QtCore.Qt.white)
            palette.setColor(QtGui.QPalette.Text, QtCore.Qt.white)
            palette.setColor(QtGui.QPalette.Button, QtGui.QColor(53,53,53))
            palette.setColor(QtGui.QPalette.ButtonText, QtCore.Qt.white)
            palette.setColor(QtGui.QPalette.BrightText, QtCore.Qt.red)
            palette.setColor(QtGui.QPalette.Highlight, QtGui.QColor(42,130,218))
            palette.setColor(QtGui.QPalette.HighlightedText, QtCore.Qt.black)
            app.setPalette(palette)
            # apply left-panel dark style so group titles and labels match the theme
            try:
                if hasattr(self, 'left_w') and self.left_w is not None:
                    self.left_w.setStyleSheet("QGroupBox:title { color: #e6e6e6; } QLabel { color: #e6e6e6; } QPushButton { color: #f0f0f0; }")
            except Exception:
                pass
        else:
            app.setPalette(app.style().standardPalette())
            try:
                if hasattr(self, 'left_w') and self.left_w is not None:
                    # clear custom styling to return to native look
                    self.left_w.setStyleSheet("")
            except Exception:
                pass
        try:
            win = self._canvas_window_ref()
            if win is not None and hasattr(win, "set_dark_mode"):
                win.set_dark_mode(bool(enabled))
        except Exception:
            pass
        if hasattr(self, 'shortcuts_label'):
            self.shortcuts_label.setText(self._shortcuts_html())
        try:
            self._apply_lower_control_theme()
        except Exception:
            pass
        self._apply_detail_view_theme()
        try:
            self._apply_molecule_button_theme()
        except Exception:
            pass
        try:
            self._apply_preview_workspace_theme()
        except Exception:
            pass

    def _set_detail_dark_view_state(self, enabled: bool, *, follow_dark_mode=None, persist: bool = True):
        self.detail_dark_view = bool(enabled)
        if follow_dark_mode is not None:
            self._detail_theme_follows_dark_mode = bool(follow_dark_mode)
        act = getattr(self, "detail_dark_act", None)
        if act is not None:
            try:
                act.blockSignals(True)
                act.setChecked(self.detail_dark_view)
                act.blockSignals(False)
            except Exception:
                pass
        if persist:
            self.config["detail_dark_view"] = self.detail_dark_view
            self.config["detail_theme_follows_dark_mode"] = bool(
                getattr(self, "_detail_theme_follows_dark_mode", True)
            )
            save_config(self.config)

    def _apply_detail_view_theme(self):
        canvases = [getattr(self, "preview_canvas", None)] + list(getattr(self, "_popup_canvases", []) or [])
        for canvas in canvases:
            if canvas is not None and hasattr(canvas, "set_detail_theme"):
                try:
                    canvas.set_detail_theme(dark=self.detail_dark_view, grid=self.detail_grid_view)
                except Exception:
                    continue

    def _apply_preview_workspace_theme(self):
        dark = bool(getattr(self, "dark_mode", False))
        frame = getattr(self, "preview_workspace_frame", None)
        if frame is not None:
            if dark:
                border = "#4c4c4c"
                bg = "#2a2a2a"
            else:
                border = "#d8dce5"
                bg = "#f6f8fb"
            frame.setStyleSheet(
                f"""
QFrame#previewWorkspaceFrame {{
    border: 1px solid {border};
    border-radius: 8px;
    background-color: {bg};
}}
"""
            )
        combo_style = (
            "QComboBox { background-color: #1f1f1f; border: 1px solid #444444; color: #f0f0f0; padding: 4px; border-radius: 4px; }"
            if dark else
            ""
        )
        for combo in (getattr(self, "thumb_cmap_combo", None), getattr(self, "preview_cmap_combo", None)):
            if combo is not None:
                combo.setStyleSheet(combo_style)
        label_style = "color: #f0f0f0; font-weight: 600;" if dark else "color: #202020; font-weight: 600;"
        for label in (
            getattr(self, "preview_title_label", None),
            getattr(self, "channel_label", None),
            getattr(self, "thumb_cmap_label", None),
            getattr(self, "preview_cmap_label", None),
        ):
            if label is not None:
                label.setStyleSheet(label_style)
        btn = getattr(self, "toolbar_dark_btn", None)
        if btn is not None:
            if dark:
                button_style = (
                    "QPushButton { padding: 4px 10px; border: 1px solid #5a5a5a; border-radius: 6px; background-color: #343434; color: #f0f0f0; }"
                    "QPushButton:checked { background-color: #2b6cb0; border-color: #2b6cb0; color: #ffffff; font-weight: 600; }"
                )
            else:
                button_style = (
                    "QPushButton { padding: 4px 10px; border: 1px solid #c8cfdb; border-radius: 6px; background-color: #ffffff; color: #202020; }"
                    "QPushButton:checked { background-color: #2b6cb0; border-color: #2b6cb0; color: #ffffff; font-weight: 600; }"
                )
            btn.setStyleSheet(button_style)
            try:
                btn.blockSignals(True)
                btn.setChecked(self.dark_mode)
                btn.blockSignals(False)
            except Exception:
                pass

    def _apply_molecule_button_theme(self):
        btn = getattr(self, "toolbar_load_mol_btn", None)
        if btn is None:
            return
        if getattr(self, "dark_mode", False):
            base = "#343434"
            hover = "#3d3d3d"
            border = "#5a5a5a"
            color = QtGui.QColor("#ffffff")
        else:
            base = "#ffffff"
            hover = "#eef4ff"
            border = "#c8cfdb"
            color = QtGui.QColor("#1d1d1d")
        btn.setStyleSheet(
            f"""
QLabel {{
    background-color: {base};
    border: 1px solid {border};
    border-radius: 6px;
}}
QLabel:hover {{
    background-color: {hover};
}}
"""
        )
        self._update_molecule_pixmap(color)

    def _update_molecule_pixmap(self, color: QtGui.QColor | None = None):
        btn = getattr(self, "toolbar_load_mol_btn", None)
        size = getattr(self, "_molecule_pixmap_size", None)
        if btn is None or size is None:
            return
        if color is None:
            color = QtGui.QColor("#ffffff" if getattr(self, "dark_mode", False) else "#1d1d1d")
        try:
            from . import main_window_toolbar as _toolbar_mod
            pixmap = _toolbar_mod._load_molecule_pixmap(size, color)
        except Exception:
            pixmap = None
        if pixmap and not pixmap.isNull():
            btn.setText("")
            btn.setPixmap(pixmap)
        else:
            btn.setPixmap(QtGui.QPixmap())
            btn.setText("Mol")

    def _append_activity_log(self, message: str):
        box = getattr(self, "activity_log_box", None)
        if box is None:
            return
        try:
            self._activity_log_pending.append(f"[{datetime.now().strftime('%H:%M:%S')}] {message}")
            if not self._activity_log_flush_timer.isActive():
                self._activity_log_flush_timer.start(60)
        except Exception:
            entry = f"[{datetime.now().strftime('%H:%M:%S')}] {message}"
            box.appendPlainText(entry)
            box.verticalScrollBar().setValue(box.verticalScrollBar().maximum())

    def _flush_activity_log_pending(self):
        box = getattr(self, "activity_log_box", None)
        if box is None:
            self._activity_log_pending = []
            return
        pending = list(getattr(self, "_activity_log_pending", []) or [])
        if not pending:
            return
        self._activity_log_pending = []
        try:
            box.appendPlainText("\n".join(pending))
            box.verticalScrollBar().setValue(box.verticalScrollBar().maximum())
        except Exception:
            for entry in pending:
                try:
                    box.appendPlainText(entry)
                except Exception:
                    pass
        try:
            QtWidgets.QApplication.processEvents(QtCore.QEventLoop.ExcludeUserInputEvents, 5)
        except Exception:
            pass

    def _on_clear_activity_log(self):
        if hasattr(self, "activity_log_box"):
            self.activity_log_box.clear()
        self._activity_log_pending = []

    def _create_lower_controls(self):
        return main_window_layout.create_lower_controls(self)

    def _build_browse_context_page(self):
        return main_window_layout.build_browse_context_page(self)

    def _build_measure_context_page(self):
        return main_window_layout.build_measure_context_page(self)

    def _build_spectro_context_page(self):
        return main_window_layout.build_spectro_context_page(self)

    def _build_display_widget(self, parent):
        return main_window_layout.build_display_widget(self, parent)

    def _apply_lower_control_theme(self):
        return main_window_layout.apply_lower_control_theme(self)

    def _on_mode_button_clicked(self, mode):
        self._apply_mode(mode)

    def _mode_name(self, mode):
        mapping = {
            self.MODE_BROWSE: "Browse",
            self.MODE_MEASURE: "Measure",
            self.MODE_SPECTRO: "Spectroscopy",
        }
        return mapping.get(mode, "Browse")

    def _mode_from_name(self, name):
        mapping = {
            "Browse": self.MODE_BROWSE,
            "Measure": self.MODE_MEASURE,
            "Spectroscopy": self.MODE_SPECTRO,
        }
        return mapping.get(str(name), self.MODE_BROWSE)

    def _apply_mode(self, mode, remember=True):
        if not hasattr(self, 'mode_stack'):
            return
        if mode not in (self.MODE_BROWSE, self.MODE_MEASURE, self.MODE_SPECTRO):
            mode = self.MODE_BROWSE
        self.mode_stack.setCurrentIndex(mode)
        self.current_mode = mode
        btn = getattr(self, 'mode_buttons', {}).get(mode)
        if btn and not btn.isChecked():
            btn.blockSignals(True)
            btn.setChecked(True)
            btn.blockSignals(False)
        if remember:
            settings = QtCore.QSettings()
            settings.setValue("lowerPane/lastMode", self._mode_name(mode))
        try:
            if mode == self.MODE_MEASURE:
                self._on_start_profile(force_enable=True)
            else:
                self._disable_profile_mode()
        except Exception:
            pass

    def _init_mode_shortcuts(self):
        self._mode_shortcuts = []
        shortcuts = [
            (QtGui.QKeySequence("Ctrl+B"), self.MODE_BROWSE),
            (QtGui.QKeySequence("Ctrl+M"), self.MODE_MEASURE),
            (QtGui.QKeySequence("Ctrl+Alt+S"), self.MODE_SPECTRO),
        ]
        for seq, mode in shortcuts:
            shortcut = QtWidgets.QShortcut(seq, self)
            shortcut.setContext(QtCore.Qt.WidgetWithChildrenShortcut)
            shortcut.activated.connect(lambda m=mode: self._on_mode_shortcut(m))
            self._mode_shortcuts.append(shortcut)

    def _on_mode_shortcut(self, mode):
        self._apply_mode(mode)
        btn = getattr(self, 'mode_buttons', {}).get(mode)
        if btn:
            try:
                btn.setFocus(QtCore.Qt.ShortcutFocusReason)
            except Exception:
                pass

    def _reset_display_options(self):
        defaults = getattr(self, '_display_defaults', {})
        action_pairs = [
            (getattr(self, 'matrix_markers_act', None), defaults.get('show_matrix_markers', True)),
            (getattr(self, 'single_markers_act', None), defaults.get('show_single_markers', True)),
            (getattr(self, 'compact_markers_act', None), defaults.get('compact_markers', True)),
            (getattr(self, 'detail_dark_act', None), defaults.get('detail_dark_view', bool(self.dark_mode))),
            (getattr(self, 'detail_grid_act', None), defaults.get('detail_grid_view', False)),
            (getattr(self, 'molecules_act', None), defaults.get('show_molecules', True)),
            (getattr(self, 'display_molecule_gizmo_act', None), defaults.get('show_molecule_gizmo', False)),
            (getattr(self, 'acquisition_overlay_act', None), defaults.get('show_acquisition_overlay', False)),
            (getattr(self, 'crop_template_act', None), defaults.get('show_crop_template_overlay', False)),
            (getattr(self, 'crop_history_act', None), defaults.get('show_crop_history_overlay', False)),
        ]
        for action, state in action_pairs:
            if action is not None:
                action.setChecked(state)
        self.on_profile_label_mode_changed(defaults.get("profile_label_mode", "length"))

    def _update_spectro_stats_label(self, stats=None):
        return main_window_spectro.update_spectro_stats_label(self, stats=stats)

    def _create_shortcuts_panel(self):
        return main_window_layout.create_shortcuts_panel(self)

    # ---------- Spectroscopy quick-inspect helpers & dialog ----------
    def _header_extent(self, header):
        return main_window_spectro.header_extent(self, header)

    def _header_scan_angle(self, header):
        """Return the configured scan angle (degrees) for a header, defaulting to 0.0."""
        if not header:
            return 0.0
        for key in ("Angle", "ScanAngle", "scan_angle", "Scan_Angle"):
            if key not in header:
                continue
            val = header.get(key)
            if val in (None, ""):
                continue
            try:
                return float(val)
            except Exception:
                try:
                    parsed = _safe_float(val)
                    if parsed is not None:
                        return float(parsed)
                except Exception:
                    continue
        return 0.0

    def _display_extent(self, extent, header=None):
        return main_window_spectro.display_extent(self, extent, header=header)

    def _spectros_near_thumb_pos(self, file_key: str, header: dict, thumb_pos_px: QtCore.QPoint, thumb_dims):
        return main_window_spectro.spectros_near_thumb_pos(self, file_key, header, thumb_pos_px, thumb_dims)

    def on_open_spectro_browser(self, entries):
        """Hook: replace with a full spectro browser. Minimal fallback shows the summary again."""
        self.open_spectro_browser(entries)

    def _next_popup_pos(self, offset=40):
        """Return a cascading popup position within the available screen."""
        screen = QtWidgets.QApplication.primaryScreen()
        avail = screen.availableGeometry() if screen else QtCore.QRect(0, 0, 1600, 900)
        base = self._popup_spawn_origin()
        # incrementing counter avoids stacking even if dialogs close quickly
        self._popup_counter = (self._popup_counter + 1) % 12
        idx = self._popup_counter
        pos = base + QtCore.QPoint(offset * (idx % 6), offset * (idx % 6))
        # clamp to screen
        x = max(avail.left(), min(pos.x(), avail.right() - 200))
        y = max(avail.top(), min(pos.y(), avail.bottom() - 150))
        return QtCore.QPoint(x, y)

    def _popup_spawn_origin(self):
        """Choose a popup origin that stays out of the center of the preview area."""
        if not self.isVisible():
            return QtGui.QCursor.pos()
        try:
            frame = self.frameGeometry()
        except Exception:
            frame = QtCore.QRect()
        if not frame.isValid():
            return QtGui.QCursor.pos()
        # Bias toward the right/top portion of the main window so thumbnails remain unobstructed
        x = frame.left() + int(frame.width() * 0.65)
        y = frame.top() + int(frame.height() * 0.15)
        screen = QtWidgets.QApplication.primaryScreen()
        if screen:
            avail = screen.availableGeometry()
            x = min(max(avail.left(), x), avail.right() - 240)
            y = min(max(avail.top(), y), avail.bottom() - 200)
        return QtCore.QPoint(x, y)

    def reveal_points_for_file(self, file_key):
        """Temporarily reveal point markers for a given file and repaint thumbnails."""
        if not hasattr(self, '_temp_reveal'):
            self._temp_reveal = set()
        key = str(file_key)
        self._temp_reveal.add(key)
        try:
            self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())
        except Exception:
            pass
        # auto-revert after 8 seconds
        try:
            QtCore.QTimer.singleShot(8000, lambda: (self._temp_reveal.discard(key), self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())))
        except Exception:
            pass

    def _open_single_spectro_popup(self, spectro):
        return main_window_spectro.open_single_spectro_popup(self, spectro)

    def _open_spectro_summary_for_file(self, file_key, show_mode="single", quiet=False):
        if not self._spectros_loaded:
            self.ensure_spectros_loaded(refresh=False)
        return main_window_spectro.open_spectro_summary_for_file(self, file_key, show_mode=show_mode, quiet=quiet)

    def _open_matrix_explorer_for_file(self, file_key):
        if not self._spectros_loaded:
            self.ensure_spectros_loaded(refresh=False)
        image_specs = [s for s in self.spectros_by_image.get(str(file_key), []) if s.get('matrix_index') is not None]
        dataset_specs = list(image_specs)
        dataset = None
        dataset_key = image_specs[0].get('matrix_dataset') if image_specs else None
        if dataset_key:
            dataset = self.matrix_datasets.get(dataset_key)
            full = [spec for spec in self.matrix_spectros if spec.get('matrix_dataset') == dataset_key]
            if full:
                dataset_specs = full
        if not dataset_specs:
            QtWidgets.QMessageBox.information(self, "Matrix explorer", "No matrix spectroscopies available for this image.")
            return

        entry = {'path': Path(file_key)}
        try:
            entry['time'] = Path(file_key).stat().st_mtime
        except Exception:
            entry['time'] = None

        dlg = MatrixSpectroViewer(
            self,
            entry,
            dataset_specs,
            dataset=dataset,
            palette_name=getattr(self, "spectro_color_cycle", DEFAULT_COLOR_CYCLE),
        )
        dlg.setAttribute(QtCore.Qt.WA_DeleteOnClose, True)
        dlg.show()
        self._popup_refs.append(dlg)
        dlg.finished.connect(lambda _: self._popup_refs.remove(dlg) if dlg in self._popup_refs else None)
        controller = getattr(self, "quick_crop_controller", None)
        if controller:
            dlg.finished.connect(lambda _=None, c=controller: c.update_popup_actions())
            controller.update_popup_actions()

    # ---------- Preview pop-outs ----------
    def _copy_view_for_popup(self, view):
        """Deep-copy a preview view dict so pop-outs do not share array state."""
        if not view:
            return {}
        new_view = dict(view)
        arr = view.get("arr")
        if arr is not None:
            try:
                new_view["arr"] = np.asarray(arr)
            except Exception:
                new_view["arr"] = arr
        return new_view

    def _spawn_preview_popup(self, views, title=None, **kwargs):
        return spawn_preview_popup(self, views, title=title, **kwargs)

    def _set_active_preview_popup(self, dlg=None, canvas=None):
        self._active_preview_popup = dlg
        self._active_preview_canvas = canvas

    def _clear_active_preview_popup(self, dlg=None):
        if dlg is None or self._active_preview_popup is dlg:
            self._active_preview_popup = None
            self._active_preview_canvas = None

    def _capture_canvas_style_snapshot(self, canvas=None):
        canvas = canvas or getattr(self, "_active_preview_canvas", None) or getattr(self, "preview_canvas", None)
        if canvas is None:
            return {}
        try:
            family = normalize_font_family(
                getattr(canvas, "_font_family", getattr(self, "_plot_font_family", UI_FONT_FAMILY)),
                UI_FONT_FAMILY,
            )
        except Exception:
            family = getattr(self, "_plot_font_family", UI_FONT_FAMILY)
        try:
            scale = float(getattr(canvas, "_view_font_scale", 1.0) or 1.0)
        except Exception:
            scale = 1.0
        scale = max(0.6, min(2.5, scale))
        rel_axes_enabled = None
        try:
            views = list(getattr(canvas, "views", []) or [])
            if views and hasattr(canvas, "_use_relative_axes"):
                rel_axes_enabled = bool(canvas._use_relative_axes(views[0]))
            else:
                rel_override = getattr(canvas, "_relative_axes_override", None)
                rel_axes_enabled = None if rel_override is None else bool(rel_override)
        except Exception:
            rel_axes_enabled = None
        rel_zero_enabled = None
        try:
            rel_zero_enabled = bool(getattr(canvas, "_popup_relative_zero_enabled"))
        except Exception:
            try:
                views = list(getattr(canvas, "views", []) or [])
                if views:
                    rel_zero_enabled = bool((views[0] or {}).get("display_relative_zero", False))
            except Exception:
                rel_zero_enabled = None
        return {
            "plot_typography": {
                "family": family,
                "bold": bool(getattr(canvas, "_plot_font_bold", False)),
                "italic": bool(getattr(canvas, "_plot_font_italic", False)),
                "underline": bool(getattr(canvas, "_plot_font_underline", False)),
            },
            "view_font_scale": scale,
            "display_options": self._canvas_display_state_from_canvas(canvas),
            "scale_bar_pos": list(getattr(canvas, "_scale_bar_pos", (0.94, 0.06))),
            "scale_bar_settings": dict(getattr(canvas, "_scale_bar_settings", {}) or {}),
            "relative_axes_enabled": rel_axes_enabled,
            "relative_zero_enabled": rel_zero_enabled,
            "view_cmaps": [str((view or {}).get("cmap") or "") for view in list(getattr(canvas, "views", []) or [])],
        }

    def _apply_canvas_style_snapshot(self, canvas, style_snapshot, *, notify=True, redraw=True):
        if canvas is None or not style_snapshot:
            return False
        typography = dict(style_snapshot.get("plot_typography") or {})
        family = typography.get("family")
        if family is not None:
            try:
                family = normalize_font_family(family, UI_FONT_FAMILY)
                canvas._font_family = family
                settings = dict(getattr(canvas, "_scale_bar_settings", {}) or {})
                settings["font_family"] = family
                canvas._scale_bar_settings = settings
            except Exception:
                pass
        for key, attr in (
            ("bold", "_plot_font_bold"),
            ("italic", "_plot_font_italic"),
            ("underline", "_plot_font_underline"),
        ):
            if key in typography:
                try:
                    setattr(canvas, attr, bool(typography.get(key)))
                except Exception:
                    pass
        try:
            canvas._view_font_scale = max(0.6, min(2.5, float(style_snapshot.get("view_font_scale", getattr(canvas, "_view_font_scale", 1.0)))))
        except Exception:
            pass
        display = dict(style_snapshot.get("display_options") or {})
        if display:
            try:
                canvas._show_ticks = bool(display.get("show_ticks", getattr(canvas, "_show_ticks", True)))
                canvas._show_colorbar = bool(display.get("show_colorbar", getattr(canvas, "_show_colorbar", True)))
                orient = str(display.get("colorbar_orientation", getattr(canvas, "_colorbar_orientation", "vertical")) or "vertical").strip().lower()
                canvas._colorbar_orientation = orient if orient in ("vertical", "horizontal") else "vertical"
                canvas._show_title = bool(display.get("show_title", getattr(canvas, "_show_title", True)))
                canvas._show_acquisition_overlay = bool(display.get("show_acquisition_overlay", getattr(canvas, "_show_acquisition_overlay", False)))
                canvas._show_shortcut_hint = bool(display.get("show_shortcut_hint", getattr(canvas, "_show_shortcut_hint", True)))
                canvas._show_profile_overlays = bool(display.get("show_profile_overlays", getattr(canvas, "_show_profile_overlays", True)))
                canvas._show_angle_overlays = bool(display.get("show_angle_overlays", getattr(canvas, "_show_angle_overlays", True)))
                canvas.show_molecules = bool(display.get("show_molecules", getattr(canvas, "show_molecules", True)))
                canvas._show_molecule_gizmo = bool(display.get("show_molecule_gizmo", getattr(canvas, "_show_molecule_gizmo", False)))
                desired_scale_bar = bool(display.get("scale_bar_enabled", getattr(canvas, "scale_bar_enabled", False)))
                current_scale_bar = bool(getattr(canvas, "scale_bar_enabled", False))
                canvas.scale_bar_enabled = desired_scale_bar
                if desired_scale_bar != current_scale_bar:
                    if desired_scale_bar:
                        canvas._connect_scale_bar_events()
                    else:
                        canvas._disconnect_scale_bar_events()
                canvas._frame_fill_mode = bool(display.get("frame_fill_mode", getattr(canvas, "_frame_fill_mode", False)))
                layout = str(display.get("view_layout", getattr(canvas, "_view_layout", "grid")) or "grid").strip().lower()
                canvas._view_layout = layout if layout in ("grid", "stacked") else "grid"
            except Exception:
                pass
        sb_settings = style_snapshot.get("scale_bar_settings")
        if sb_settings is not None:
            try:
                canvas._scale_bar_settings = dict(sb_settings or {})
            except Exception:
                pass
        sb_pos = style_snapshot.get("scale_bar_pos")
        if sb_pos is not None:
            try:
                canvas._scale_bar_pos = tuple(sb_pos)
            except Exception:
                pass
        rel_axes_enabled = style_snapshot.get("relative_axes_enabled", None)
        if rel_axes_enabled is not None:
            try:
                canvas._relative_axes_override = bool(rel_axes_enabled)
            except Exception:
                pass
        else:
            try:
                rel_override = display.get("relative_axes_override", getattr(canvas, "_relative_axes_override", None))
                canvas._relative_axes_override = None if rel_override is None else bool(rel_override)
            except Exception:
                pass
        rel_zero_enabled = style_snapshot.get("relative_zero_enabled", None)
        if rel_zero_enabled is not None:
            try:
                rel_zero_setter = getattr(canvas, "_popup_relative_zero_setter", None)
                if callable(rel_zero_setter):
                    rel_zero_setter(bool(rel_zero_enabled))
            except Exception:
                pass
        view_cmaps = list(style_snapshot.get("view_cmaps") or [])
        if view_cmaps:
            try:
                target_views = list(getattr(canvas, "views", []) or [])
                if len(view_cmaps) == len(target_views):
                    for target_view, cmap_name in zip(target_views, view_cmaps):
                        if cmap_name:
                            target_view["cmap"] = str(cmap_name)
                elif target_views and view_cmaps[0]:
                    for target_view in target_views:
                        target_view["cmap"] = str(view_cmaps[0])
            except Exception:
                pass
        if redraw:
            try:
                canvas._redraw()
            except Exception:
                try:
                    canvas.draw_idle()
                except Exception:
                    pass
        else:
            try:
                canvas._apply_view_font_scale()
            except Exception:
                pass
        if notify:
            try:
                canvas._notify_views_callback()
            except Exception:
                pass
        return True

    def _apply_popup_style_to_all(self, source_canvas=None):
        source_canvas = source_canvas or getattr(self, "_active_preview_canvas", None)
        if source_canvas is None:
            return 0
        style_snapshot = self._capture_canvas_style_snapshot(source_canvas)
        if not style_snapshot:
            return 0
        count = 0
        for canvas in list(getattr(self, "_popup_canvases", []) or []):
            if canvas is None or canvas is source_canvas:
                continue
            try:
                if self._apply_canvas_style_snapshot(canvas, style_snapshot, notify=True):
                    count += 1
            except Exception:
                continue
        if count:
            try:
                log_status(f"Applied popup style to {count} pop-out(s)")
            except Exception:
                pass
        return count

    def _on_molecule_palette_changed(self, palette: str):
        palette = (palette or "cpk").lower()
        self.molecule_palette = palette
        try:
            self.config["molecule_palette"] = palette
            save_config(self.config)
        except Exception:
            pass
        try:
            self.preview_canvas.set_molecule_palette(palette, notify=False)
        except Exception:
            pass
        for canv in list(self._popup_canvases):
            try:
                canv.set_molecule_palette(palette, notify=False)
            except Exception:
                continue

    def _on_preview_crop(self, view, source_canvas=None):
        """Receive cropped view from preview canvas and pop it out."""
        if not view:
            return
        preview_canvas = getattr(self, "preview_canvas", None)
        seq = view.get("crop_sequence")
        if (
            bool(getattr(self, "quick_crop_mode", False))
            and source_canvas is not None
            and preview_canvas is not None
            and source_canvas is not preview_canvas
            and seq is not None
        ):
            try:
                entry = source_canvas.get_fixed_crop_history_entry(seq)
            except Exception:
                entry = None
            if entry:
                try:
                    preview_canvas.import_fixed_crop_history_entry(entry, update_size=True)
                except Exception:
                    pass
        auto_virtual_copy = bool(view.get("_auto_virtual_copy", False))
        skip_virtual_copy_prompt = bool(view.get("_skip_virtual_copy_prompt", False))
        # Offer to save crop as a virtual copy in thumbnails, unless explicitly
        # requested by the crop tool for frictionless iterative workflows.
        try:
            path = view.get("path") or (view.get("meta") or {}).get("path")
            if path:
                if auto_virtual_copy:
                    self._create_virtual_crop_view(view)
                elif not skip_virtual_copy_prompt:
                    ret = QtWidgets.QMessageBox.question(
                        self,
                        "Save crop",
                        "Add this cropped view to thumbnails as a virtual copy?",
                        QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                        QtWidgets.QMessageBox.No,
                    )
                    if ret == QtWidgets.QMessageBox.Yes:
                        self._create_virtual_crop_view(view)
        except Exception:
            pass
        title = view.get("title") or "Cropped view"
        if self.show_crop_history_overlay and seq is not None:
            title = f"{title} #{seq}"
        self._spawn_preview_popup([self._copy_view_for_popup(view)], title=title, source_canvas=source_canvas)

    # ---------- Spectro browser dock ----------
    def _ensure_spectro_dock(self):
        return main_window_spectro.ensure_spectro_dock(self)

    def open_spectro_browser(self, entries=None):
        if not self._spectros_loaded:
            self.ensure_spectros_loaded(refresh=False)
        return main_window_spectro.open_spectro_browser(self, entries=entries)

    def _filter_spectro_browser(self):
        return main_window_spectro.filter_spectro_browser(self)

    def _on_spectro_browser_selection(self, current, _prev):
        return main_window_spectro.on_spectro_browser_selection(self, current, _prev)

    def _shortcuts_html(self):
        color = "#f0f4ff" if getattr(self, 'dark_mode', False) else "#203050"
        return (
            "<ul style='margin:4px 12px;padding-left:12px;color:%s'>"
            "<li><b>Shift+Click</b> minimap frame = hide entry</li>"
            "<li><b>Show all frames</b> button resets minimap filters</li>"
            "<li><b>Ctrl+Wheel</b> over thumbnails = resize previews</li>"
            "<li><b>Shift+Click</b> spectroscopy marker = multi-select</li>"
            "<li><b>Ctrl+Drag</b> thumbnails = reorder export selection</li>"
            "<li><b>Ctrl+A</b> in thumbnails = select all visible thumbnails</li>"
            "<li><b>Shift/Ctrl+Click</b> thumbnails + <b>Ctrl+C</b> = copy selected as separate PNG files</li>"
            "<li><b>Ctrl+C</b> over preview/popup = copy displayed PNG</li>"
            "<li><b>Popup canvas</b>: A auto contrast, 0 toggles relative-zero, Ctrl+Click profile, Ctrl+Alt+Click angle, click a molecule then X/Y/Z rotate it, Shift+X/Y/Z rotates opposite, Shift+drag rotates around Z, Ctrl+Shift+drag or middle-drag rotates in 3D, Ctrl+1/2/3 saved overlays</li>"
            "<li><b>Ctrl+S</b> = save current session | <b>Ctrl+Z</b> = reopen last closed window (when no other undo applies)</li>"
            "<li><b>Ctrl+Shift+M</b> = minimize all open pop-outs</li>"
            "</ul>"
        ) % color

    def _set_shortcuts_panel_visible(self, visible, remember=True):
        if hasattr(self, 'shortcuts_panel'):
            self.shortcuts_panel.setVisible(bool(visible))
        if remember:
            self.show_shortcuts_panel = bool(visible)
            self.config['show_shortcuts_panel'] = self.show_shortcuts_panel
            save_config(self.config)

    def _on_hide_shortcuts_panel(self):
        self._set_shortcuts_panel_visible(False)

    def _on_shortcuts_never_show_clicked(self):
        self._set_shortcuts_panel_visible(False)

    def _on_show_shortcuts_requested(self):
        self._set_shortcuts_panel_visible(True)

    def _window_history_views_dir(self):
        path = Path(tempfile.gettempdir()) / "sxm_viewer_window_history"
        path.mkdir(parents=True, exist_ok=True)
        return path

    def _session_recovery_path(self):
        raw = str(self.config.get("session_recovery_path", "") or "").strip()
        if raw:
            try:
                return Path(raw)
            except Exception:
                pass
        return CONFIG_PATH.with_name(".sxm_viewer_recovery.json")

    def _workspace_has_content(self):
        return bool(
            getattr(self, "files", None)
            or getattr(self, "_processed_views", None)
            or getattr(self, "last_preview", None)
            or getattr(self, "_popup_refs", None)
            or getattr(self, "_spectro_popups", None)
            or getattr(self, "_multi_spectro_popups", None)
            or getattr(self, "_deferred_popup_entries", None)
            or getattr(self, "virtual_copy_order", None)
        )

    def _capture_window_state_payload(self, window):
        payload = {}
        if window is None:
            return payload
        try:
            geo = window.geometry()
            payload["geometry"] = [int(geo.x()), int(geo.y()), int(geo.width()), int(geo.height())]
        except Exception:
            pass
        try:
            payload["window_state"] = int(window.windowState())
        except Exception:
            pass
        try:
            payload["title"] = str(window.windowTitle() or "")
        except Exception:
            pass
        return payload

    def _apply_window_state_payload(self, window, payload):
        if window is None or not isinstance(payload, dict):
            return
        geom = payload.get("geometry")
        if geom and len(geom) == 4:
            try:
                x, y, w, h = [int(v) for v in geom]
                window.setGeometry(x, y, w, h)
            except Exception:
                pass
        if payload.get("window_state") is not None:
            try:
                window.setWindowState(QtCore.Qt.WindowStates(int(payload.get("window_state"))))
            except Exception:
                pass

    def _push_closed_window_history(self, payload):
        if self._suspend_window_history or not isinstance(payload, dict):
            return
        kind = str(payload.get("kind") or "").strip()
        if not kind:
            return
        history = list(getattr(self, "_closed_window_history", []) or [])
        history.append(payload)
        self._closed_window_history = history[-int(getattr(self, "_closed_window_history_limit", 6) or 6):]
        try:
            log_status(f"Stored closed window in reopen history: {kind}")
        except Exception:
            pass

    def _clear_closed_window_history(self):
        self._closed_window_history = []

    def _remember_closed_preview_popup(self, dlg, canvas):
        if self._suspend_window_history or dlg is None or canvas is None:
            return
        try:
            if getattr(dlg, "_window_history_captured", False):
                return
            dlg._window_history_captured = True
        except Exception:
            pass
        try:
            history_dir = self._window_history_views_dir()
            prefix = f"closed_popup_{int(time.time() * 1000)}"
            snapshot = self.session_controller._capture_canvas_snapshot(canvas, history_dir, prefix=prefix, include_arrays=True)
        except Exception:
            snapshot = None
            history_dir = None
        if not snapshot or history_dir is None:
            return
        payload = {"kind": "preview_popup", "snapshot": snapshot, "views_dir": str(history_dir)}
        payload.update(self._capture_window_state_payload(dlg))
        self._push_closed_window_history(payload)

    def _remember_closed_main_profile_dialog(self, dlg):
        if self._suspend_window_history or dlg is None:
            return
        try:
            state = viewer_measurement.export_profile_dialog_state(self)
        except Exception:
            state = None
        if not state:
            return
        payload = {"kind": "main_profile", "state": state}
        payload.update(self._capture_window_state_payload(dlg))
        self._push_closed_window_history(payload)

    def _remember_closed_popup_profile_dialog(self, controller, dlg):
        if self._suspend_window_history or controller is None or dlg is None:
            return
        try:
            state = controller.export_dialog_state()
        except Exception:
            state = None
        if not state:
            return
        payload = {"kind": "popup_profile", "controller": controller, "state": state}
        payload.update(self._capture_window_state_payload(dlg))
        self._push_closed_window_history(payload)

    def _remember_closed_spectro_dialog(self, dlg):
        if self._suspend_window_history or dlg is None:
            return
        payload = None
        try:
            if isinstance(dlg, SpectroscopyPopup):
                payload = {
                    "kind": "spectro_popup",
                    "spec": copy.deepcopy(getattr(dlg, "spec", None)),
                    "channel": str(dlg.channel_combo.currentText() or ""),
                    "axis_key": dlg.axis_combo.currentData(),
                }
            elif isinstance(dlg, SpectroscopyCompareDialog):
                payload = {
                    "kind": "spectro_compare",
                    "specs": copy.deepcopy(list(getattr(dlg, "specs", []) or [])),
                    "palette_name": str(getattr(dlg, "_palette_name", DEFAULT_COLOR_CYCLE) or DEFAULT_COLOR_CYCLE),
                    "state": dlg._snapshot_state() if hasattr(dlg, "_snapshot_state") else None,
                }
            elif isinstance(dlg, MatrixSpectroViewer):
                payload = {
                    "kind": "matrix_viewer",
                    "image_entry": copy.deepcopy(getattr(dlg, "image_entry", None)),
                    "specs": copy.deepcopy(list(getattr(dlg, "specs", []) or [])),
                    "dataset": copy.deepcopy(getattr(dlg, "dataset", None)),
                    "palette_name": str(getattr(dlg, "palette_name", DEFAULT_COLOR_CYCLE) or DEFAULT_COLOR_CYCLE),
                }
        except Exception:
            payload = None
        if not payload:
            return
        payload.update(self._capture_window_state_payload(dlg))
        self._push_closed_window_history(payload)

    def _remember_closed_canvas_window(self, win):
        if self._suspend_window_history or win is None:
            return
        try:
            state = win._capture_state()
        except Exception:
            state = None
        if not state:
            return
        payload = {"kind": "canvas_window", "state": state}
        payload.update(self._capture_window_state_payload(win))
        self._push_closed_window_history(payload)

    def _restore_closed_window_payload(self, payload):
        if not isinstance(payload, dict):
            return False
        kind = str(payload.get("kind") or "")
        if kind == "preview_popup":
            snapshot = payload.get("snapshot") or {}
            views_dir = payload.get("views_dir")
            if not snapshot or not views_dir:
                return False
            try:
                dlg = self.session_controller._restore_popup_dialog_from_snapshot(
                    snapshot,
                    Path(str(views_dir)),
                    geometry=payload.get("geometry"),
                    window_state=payload.get("window_state"),
                    title=payload.get("title") or snapshot.get("window_title") or "Preview",
                    visible=True,
                    active=True,
                )
                return dlg is not None
            except Exception:
                return False
        if kind == "main_profile":
            try:
                dlg = viewer_measurement.restore_profile_dialog_state(self, payload.get("state") or {})
                return dlg is not None
            except Exception:
                return False
        if kind == "popup_profile":
            controller = payload.get("controller")
            if controller is None:
                return False
            try:
                dlg = controller.restore_dialog_state(payload.get("state") or {})
                return dlg is not None
            except Exception:
                return False
        if kind == "spectro_popup":
            spec = payload.get("spec")
            if not spec:
                return False
            try:
                dlg = spectro_popups._open_spectroscopy_popup(self, spec)
                if dlg is None:
                    return False
                channel = str(payload.get("channel") or "")
                if channel:
                    idx = dlg.channel_combo.findText(channel)
                    if idx >= 0:
                        dlg.channel_combo.setCurrentIndex(idx)
                axis_key = payload.get("axis_key")
                if axis_key is not None:
                    idx = dlg.axis_combo.findData(axis_key)
                    if idx >= 0:
                        dlg.axis_combo.setCurrentIndex(idx)
                self._apply_window_state_payload(dlg, payload)
                return True
            except Exception:
                return False
        if kind == "spectro_compare":
            specs = list(payload.get("specs") or [])
            if len(specs) < 2:
                return False
            try:
                dlg = SpectroscopyCompareDialog(specs, parent=self, palette_name=payload.get("palette_name") or DEFAULT_COLOR_CYCLE)
                if payload.get("state") and hasattr(dlg, "_apply_state"):
                    try:
                        dlg._apply_state(payload.get("state"))
                    except Exception:
                        pass
                self._apply_window_state_payload(dlg, payload)
                dlg.show()
                self._multi_spectro_popups.append(dlg)
                dlg.finished.connect(lambda _: self._multi_spectro_popups.remove(dlg) if dlg in self._multi_spectro_popups else None)
                dlg.finished.connect(lambda _: self._remember_closed_spectro_dialog(dlg))
                controller = getattr(self, "quick_crop_controller", None)
                if controller:
                    dlg.finished.connect(lambda _=None, c=controller: c.update_popup_actions())
                    controller.update_popup_actions()
                return True
            except Exception:
                return False
        if kind == "matrix_viewer":
            try:
                dlg = MatrixSpectroViewer(
                    self,
                    payload.get("image_entry") or {},
                    payload.get("specs") or [],
                    dataset=payload.get("dataset"),
                    palette_name=payload.get("palette_name") or DEFAULT_COLOR_CYCLE,
                )
                self._apply_window_state_payload(dlg, payload)
                dlg.show()
                self._popup_refs.append(dlg)
                dlg.finished.connect(lambda _: self._popup_refs.remove(dlg) if dlg in self._popup_refs else None)
                dlg.finished.connect(lambda _: self._remember_closed_spectro_dialog(dlg))
                controller = getattr(self, "quick_crop_controller", None)
                if controller:
                    dlg.finished.connect(lambda _=None, c=controller: c.update_popup_actions())
                    controller.update_popup_actions()
                return True
            except Exception:
                return False
        if kind == "canvas_window":
            try:
                self._on_open_canvas()
                win = self._canvas_window_ref()
                if win is None:
                    return False
                if payload.get("state"):
                    try:
                        win._restore_state(payload.get("state"))
                    except Exception:
                        pass
                self._apply_window_state_payload(win, payload)
                return True
            except Exception:
                return False
        return False

    def _restore_last_closed_window(self):
        history = list(getattr(self, "_closed_window_history", []) or [])
        while history:
            payload = history.pop()
            self._closed_window_history = history
            if self._restore_closed_window_payload(payload):
                try:
                    log_status(f"Reopened last closed window ({payload.get('kind')})")
                except Exception:
                    pass
                return True
        return False

    def _iter_workspace_windows(self, *, include_canvas: bool = True):
        seen = set()
        candidates = []
        for attr in ("_profile_dialog",):
            dlg = getattr(self, attr, None)
            if dlg is not None:
                candidates.append(dlg)
        for attr in ("_profile_dialogs", "_spectro_popups", "_multi_spectro_popups", "_popup_refs"):
            candidates.extend(list(getattr(self, attr, []) or []))
        tray = getattr(self, "collection_tray_window", None)
        if tray is not None:
            candidates.append(tray)
        if include_canvas:
            win = self._canvas_window_ref()
            if win is not None:
                candidates.append(win)
        for dlg in candidates:
            if dlg is None:
                continue
            key = id(dlg)
            if key in seen:
                continue
            seen.add(key)
            yield dlg

    def _close_workspace_windows(self, *, record_history: bool = False, include_canvas: bool = True):
        previous = bool(getattr(self, "_suspend_window_history", False))
        self._suspend_window_history = (not record_history) or previous
        try:
            for dlg in list(self._iter_workspace_windows(include_canvas=include_canvas)):
                try:
                    dlg.close()
                except Exception:
                    continue
        finally:
            self._suspend_window_history = previous

    def _prepare_for_workspace_load(self, kind="workspace"):
        self._workspace_window_shutdown = True
        try:
            self._clear_closed_window_history()
            self._close_workspace_windows(record_history=False, include_canvas=True)
        finally:
            self._workspace_window_shutdown = False
        self._current_session_path = None

    def _refresh_autosave_timer(self):
        timer = getattr(self, "_autosave_timer", None)
        if timer is None:
            return
        if self._session_recovery_enabled:
            timer.start(int(max(1, self._session_recovery_interval_min) * 60 * 1000))
        else:
            timer.stop()

    def _save_recovery_snapshot(self, *, reason="autosave"):
        if self._autosave_busy or not self._session_recovery_enabled or not self._workspace_has_content():
            return False
        self._autosave_busy = True
        try:
            recovery_path = self._session_recovery_path()
            recovery_path.parent.mkdir(parents=True, exist_ok=True)
            saved = self.session_controller.save_session(
                session_path=recovery_path,
                prompt_if_missing=False,
                record_recent=False,
                set_current=False,
                autosave=True,
                quiet=True,
            )
            if saved:
                self.config["session_recovery_path"] = str(recovery_path)
                self.config["session_recovery_timestamp"] = int(time.time())
                save_config(self.config)
                try:
                    if reason == "autosave":
                        log_status(f"Auto-saved recovery workspace to {recovery_path}")
                except Exception:
                    pass
                return True
            return False
        finally:
            self._autosave_busy = False

    def _discard_recovery_snapshot(self):
        recovery_path = self._session_recovery_path()
        data_dir = recovery_path.parent / f"{recovery_path.stem}_data"
        try:
            if recovery_path.exists():
                recovery_path.unlink()
        except Exception:
            pass
        try:
            if data_dir.exists():
                shutil.rmtree(data_dir, ignore_errors=True)
        except Exception:
            pass

    def _on_autosave_timer(self):
        self._save_recovery_snapshot(reason="autosave")

    def _maybe_offer_recovery_session(self):
        if not self._session_recovery_enabled or self._workspace_has_content() or getattr(self, "_workspace_loading", False):
            return
        recovery_path = self._session_recovery_path()
        if not recovery_path.exists():
            return
        try:
            ts = datetime.fromtimestamp(recovery_path.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            ts = "recently"
        box = QtWidgets.QMessageBox(self)
        box.setWindowTitle("Recover autosaved workspace")
        box.setIcon(QtWidgets.QMessageBox.Question)
        box.setTextFormat(QtCore.Qt.RichText)
        box.setText(
            f"An autosaved workspace was found from <b>{ts}</b>.<br><br>"
            "Recover it now, ignore it for this launch, or discard it."
        )
        recover_btn = box.addButton("Recover", QtWidgets.QMessageBox.AcceptRole)
        later_btn = box.addButton("Ignore for now", QtWidgets.QMessageBox.RejectRole)
        discard_btn = box.addButton("Discard", QtWidgets.QMessageBox.DestructiveRole)
        box.exec_()
        clicked = box.clickedButton()
        if clicked is recover_btn:
            self.session_controller.load_session(
                session_path=recovery_path,
                record_recent=False,
                set_current=False,
            )
        elif clicked is discard_btn:
            self._discard_recovery_snapshot()
        else:
            return

    def on_recover_latest_autosave(self):
        recovery_path = self._session_recovery_path()
        if not recovery_path.exists():
            QtWidgets.QMessageBox.information(self, "Recovery", "No autosaved workspace is available.")
            return
        self.session_controller.load_session(
            session_path=recovery_path,
            record_recent=False,
            set_current=False,
        )

    def on_discard_recovery_snapshot(self):
        self._discard_recovery_snapshot()
        QtWidgets.QMessageBox.information(self, "Recovery", "Autosaved recovery data was discarded.")

    def on_toggle_session_recovery(self, enabled: bool):
        self._session_recovery_enabled = bool(enabled)
        self.config["session_recovery_enabled"] = self._session_recovery_enabled
        save_config(self.config)
        self._refresh_autosave_timer()
        self._refresh_session_recovery_ui()

    def on_set_session_recovery_interval(self, minutes: int):
        try:
            minutes = max(1, int(minutes))
        except Exception:
            minutes = 5
        self._session_recovery_interval_min = minutes
        self.config["session_recovery_interval_min"] = minutes
        save_config(self.config)
        self._refresh_autosave_timer()
        self._refresh_session_recovery_ui()

    def _refresh_session_recovery_ui(self):
        text = f"Autosave recovery: {'On' if self._session_recovery_enabled else 'Off'} ({self._session_recovery_interval_min} min)"
        for attr in ("session_recovery_status_act", "toolbar_session_recovery_status_act"):
            act = getattr(self, attr, None)
            if act is not None:
                try:
                    act.setText(text)
                except Exception:
                    pass
        for attr in ("session_recovery_enable_act", "toolbar_session_recovery_enable_act"):
            act = getattr(self, attr, None)
            if act is not None:
                try:
                    act.blockSignals(True)
                    act.setChecked(self._session_recovery_enabled)
                    act.blockSignals(False)
                except Exception:
                    pass
        recovery_exists = self._session_recovery_path().exists()
        for attr in ("session_recovery_open_act", "toolbar_session_recovery_open_act", "session_recovery_discard_act", "toolbar_session_recovery_discard_act"):
            act = getattr(self, attr, None)
            if act is not None:
                try:
                    act.setEnabled(recovery_exists)
                except Exception:
                    pass
        for minutes, act in dict(getattr(self, "session_recovery_interval_actions", {}) or {}).items():
            try:
                act.blockSignals(True)
                act.setChecked(int(minutes) == int(self._session_recovery_interval_min))
                act.blockSignals(False)
            except Exception:
                pass

    def on_save_session_as(self):
        self.session_controller.save_session_as()

    def _focused_canvas_with_undo(self):
        widget = QtWidgets.QApplication.focusWidget()
        while widget is not None:
            undo_fn = getattr(widget, "handle_undo_request", None)
            if callable(undo_fn):
                return widget
            undo_fn = getattr(widget, "undo_last_action", None)
            if callable(undo_fn):
                return widget
            widget = widget.parentWidget()
        return None

    def _on_global_undo_requested(self):
        focus_widget = QtWidgets.QApplication.focusWidget()
        if isinstance(
            focus_widget,
            (
                QtWidgets.QLineEdit,
                QtWidgets.QTextEdit,
                QtWidgets.QPlainTextEdit,
                QtWidgets.QAbstractSpinBox,
            ),
        ):
            try:
                focus_widget.undo()
            except Exception:
                pass
            return
        canvas = self._focused_canvas_with_undo()
        if canvas is not None:
            try:
                handle_undo = getattr(canvas, "handle_undo_request", None)
                if callable(handle_undo) and handle_undo():
                    return
                undo_last = getattr(canvas, "undo_last_action", None)
                if callable(undo_last) and undo_last():
                    return
            except Exception:
                pass
        try:
            if self.quick_crop_controller.undo_last_crop():
                return
        except Exception:
            pass
        self._restore_last_closed_window()

    def _handle_local_file_mime_drop(self, mime):
        if mime is None or not mime.hasUrls():
            return False
        dirs = []
        files = []
        for url in mime.urls():
            if url.isLocalFile():
                path = Path(url.toLocalFile())
                if path.is_dir():
                    dirs.append(path)
                elif path.exists():
                    files.append(path)
        if not dirs and not files:
            return False
        if len(dirs) == 1 and not files:
            self.load_folder(dirs[0])
            return True
        drop_image_files = []
        drop_spectro_files = []
        for folder in dirs:
            try:
                drop_image_files.extend(viewer_loader.collect_folder_image_paths(self, folder))
            except Exception:
                continue
        explicit_images, explicit_spectros = viewer_loader.classify_dropped_paths(self, files)
        drop_image_files.extend(explicit_images)
        drop_spectro_files.extend(explicit_spectros)
        if drop_image_files:
            folder_hint = None
            if len(dirs) == 1 and not files:
                folder_hint = dirs[0]
            elif len(explicit_images) and len({str(p.parent) for p in explicit_images}) == 1 and not dirs:
                folder_hint = explicit_images[0].parent
            self.load_files(drop_image_files, folder_hint=folder_hint, append=True, refresh_spectros=False)
        if drop_spectro_files:
            spectro_hint = None
            if len(drop_spectro_files) == 1:
                spectro_hint = drop_spectro_files[0].parent
            loaded_specs = self.load_spectroscopy_files(drop_spectro_files, folder_hint=spectro_hint, append=True, refresh=True)
            if len(drop_spectro_files) == 1 and not drop_image_files:
                try:
                    dropped = str(Path(drop_spectro_files[0]).resolve()).lower()
                except Exception:
                    dropped = str(drop_spectro_files[0]).lower()
                spec = None
                for item in loaded_specs or []:
                    try:
                        key = str(Path(item.get("path", "")).resolve()).lower()
                    except Exception:
                        key = str(item.get("path", "")).lower()
                    if key == dropped:
                        spec = item
                        break
                if spec is None and loaded_specs:
                    spec = loaded_specs[0]
                if spec is not None:
                    self._open_spectroscopy_popup(spec)
        return True

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                if url.isLocalFile():
                    event.acceptProposedAction()
                    return
        super().dragEnterEvent(event)

    def dropEvent(self, event):
        if self._handle_local_file_mime_drop(event.mimeData()):
            event.acceptProposedAction()
            return
        super().dropEvent(event)

    def eventFilter(self, obj, event):
        thumb_objects = (
            getattr(self, '_thumb_viewport', None),
            getattr(self, 'thumb_container', None),
            getattr(self, 'scroll', None),
        )
        preview_canvas = getattr(self, 'preview_canvas', None)

        if obj in thumb_objects and event.type() in (
            QtCore.QEvent.DragEnter,
            QtCore.QEvent.DragMove,
            QtCore.QEvent.Drop,
        ):
            if self._handle_thumbnail_drag_event(event):
                return True

        # Handle Ctrl+Wheel over the thumbnails to resize thumbnails
        if obj in thumb_objects and event.type() == QtCore.QEvent.Wheel:
            if event.modifiers() & QtCore.Qt.ControlModifier:
                delta = event.angleDelta().y() or event.pixelDelta().y()
                if delta != 0:
                    step = 16 if delta > 0 else -16
                    self._resize_thumbnail_scale(step)
                event.accept()
                return True

        # Arrow navigation when the scroll area / container has focus
        if obj in thumb_objects and event.type() == QtCore.QEvent.KeyPress:
            key = event.key()
            if (event.modifiers() & QtCore.Qt.ControlModifier) and key == QtCore.Qt.Key_A:
                if self._select_all_thumbnails():
                    event.accept()
                    return True
            if key in (
                QtCore.Qt.Key_Left,
                QtCore.Qt.Key_Right,
                QtCore.Qt.Key_Up,
                QtCore.Qt.Key_Down,
            ):
                focus_widget = QtWidgets.QApplication.focusWidget()
                if not self._focus_widget_blocks_thumb_nav(focus_widget):
                    if self._handle_thumbnail_navigation(key, event.modifiers()):
                        event.accept()
                        return True

        # Rubber band selection on thumb_container
        if obj is getattr(self, 'thumb_container', None):
            if event.type() == QtCore.QEvent.MouseButtonPress:
                if event.button() == QtCore.Qt.LeftButton:
                    self._rubber_band_origin = event.pos()
                    if not hasattr(self, '_rubber_band'):
                        self._rubber_band = QtWidgets.QRubberBand(QtWidgets.QRubberBand.Rectangle, self.thumb_container)
                    self._rubber_band.setGeometry(QtCore.QRect(self._rubber_band_origin, QtCore.QSize()))
                    self._rubber_band.show()
                    
                    if not hasattr(self, 'thumb_multi_select') or self.thumb_multi_select is None:
                        self.thumb_multi_select = set()
                    if not hasattr(self, 'spectro_thumb_multi_select') or self.spectro_thumb_multi_select is None:
                        self.spectro_thumb_multi_select = set()
                    self._selection_before_drag = set(self.thumb_multi_select)
                    self._spectro_selection_before_drag = set(self.spectro_thumb_multi_select)
                    
                    if not (event.modifiers() & (QtCore.Qt.ShiftModifier | QtCore.Qt.ControlModifier)):
                        self._selection_before_drag = set()
                        self._spectro_selection_before_drag = set()
                        self._clear_thumb_multi_selection()
                        self._clear_spectro_thumb_multi_selection()
                    return True
            elif event.type() == QtCore.QEvent.MouseMove:
                if hasattr(self, '_rubber_band') and self._rubber_band.isVisible():
                    rect = QtCore.QRect(self._rubber_band_origin, event.pos()).normalized()
                    self._rubber_band.setGeometry(rect)
                    self._update_rubber_band_selection(rect, event.modifiers())
                    return True
            elif event.type() == QtCore.QEvent.MouseButtonRelease:
                if hasattr(self, '_rubber_band') and self._rubber_band.isVisible():
                    self._rubber_band.hide()
                    if hasattr(self, '_selection_before_drag'):
                        del self._selection_before_drag
                    if hasattr(self, '_spectro_selection_before_drag'):
                        del self._spectro_selection_before_drag
                    return True

        # When the thumbnail viewport or container is resized, debounce and repopulate so
        # the thumbnail grid recomputes columns responsively.
        if obj in (getattr(self, '_thumb_viewport', None),
                   getattr(self, 'thumb_container', None),
                   getattr(self, 'scroll', None)) and event.type() == QtCore.QEvent.Resize:
            try:
                self._thumbs_reflow_timer.start(150)
            except Exception:
                pass
        if obj in thumb_objects and event.type() in (QtCore.QEvent.Scroll, QtCore.QEvent.Wheel):
            try:
                # defer slightly to let the scroll settle
                QtCore.QTimer.singleShot(0, self._request_visible_thumbs)
            except Exception:
                pass
        if obj is preview_canvas and event.type() == QtCore.QEvent.KeyPress:
            key = event.key()
            if key in (QtCore.Qt.Key_Left, QtCore.Qt.Key_Right, QtCore.Qt.Key_Up, QtCore.Qt.Key_Down):
                if getattr(self, "current_mode", self.MODE_BROWSE) == self.MODE_BROWSE and self._handle_thumbnail_navigation(key, event.modifiers()):
                    event.accept()
                    return True
        # allow normal resize processing to continue
        return super().eventFilter(obj, event)

    def keyPressEvent(self, event):
        key = event.key()
        mods = event.modifiers()
        if (mods & QtCore.Qt.ControlModifier) and key == QtCore.Qt.Key_A:
            focus_widget = QtWidgets.QApplication.focusWidget()
            if self._is_widget_in_thumbnail_area(focus_widget):
                if self._select_all_thumbnails():
                    event.accept()
                    return
        if (mods & QtCore.Qt.ControlModifier) and key == QtCore.Qt.Key_C:
            focus_widget = QtWidgets.QApplication.focusWidget()
            if self._is_widget_in_thumbnail_area(focus_widget):
                if self._copy_thumbnail_selection_to_clipboard_files():
                    event.accept()
                    return
        if (mods & QtCore.Qt.ControlModifier) and key == QtCore.Qt.Key_D:
            views = getattr(self.preview_canvas, "views", None)
            if views:
                self._spawn_preview_popup([self._copy_view_for_popup(v) for v in views], title="Preview copy")
                event.accept()
                return
        if key in (
            QtCore.Qt.Key_Left,
            QtCore.Qt.Key_Right,
            QtCore.Qt.Key_Up,
            QtCore.Qt.Key_Down,
        ):
            focus_widget = QtWidgets.QApplication.focusWidget()
            if not self._focus_widget_blocks_thumb_nav(focus_widget):
                if self._handle_thumbnail_navigation(key, mods):
                    event.accept()
                    return
        super().keyPressEvent(event)

    def _focus_widget_blocks_thumb_nav(self, widget):
        if widget is None:
            return False
        blocking_types = (
            QtWidgets.QLineEdit,
            QtWidgets.QTextEdit,
            QtWidgets.QPlainTextEdit,
            QtWidgets.QSpinBox,
            QtWidgets.QDoubleSpinBox,
            QtWidgets.QAbstractSpinBox,
            QtWidgets.QComboBox,
        )
        return isinstance(widget, blocking_types)

    def _is_widget_descendant(self, widget, ancestor):
        if widget is None or ancestor is None:
            return False
        cur = widget
        while cur is not None:
            if cur is ancestor:
                return True
            cur = cur.parentWidget()
        return False

    def _is_widget_in_thumbnail_area(self, widget):
        thumb_container = getattr(self, "thumb_container", None)
        thumb_viewport = getattr(self, "_thumb_viewport", None)
        scroll = getattr(self, "scroll", None)
        if widget is None:
            return False
        return (
            self._is_widget_descendant(widget, thumb_container)
            or self._is_widget_descendant(widget, thumb_viewport)
            or self._is_widget_descendant(widget, scroll)
        )

    def _thumbnail_virtual_drag_payload(self, mime):
        if mime is None or not mime.hasFormat("application/x-sxm-view"):
            return None
        try:
            payload = json.loads(bytes(mime.data("application/x-sxm-view")).decode("utf-8"))
        except Exception:
            return None
        if not isinstance(payload, dict):
            return None
        if str(payload.get("drag_origin") or "") != "preview_canvas":
            return None
        return payload

    def _thumbnail_drop_insert_anchor(self, global_pos):
        keys = [
            str(key)
            for key in list(getattr(self, "current_thumb_files", []) or [])
            if str(key) in getattr(self, "thumb_widgets", {})
        ]
        if not keys:
            return VIRTUAL_COPY_INSERT_START
        thumb_columns = int(getattr(self, "thumb_grid_columns", 1) or 1)
        best_idx = None
        best_rect = None
        best_dist = None
        for idx, key in enumerate(keys):
            widget = getattr(self, "thumb_widgets", {}).get(key)
            if widget is None:
                continue
            try:
                rect = QtCore.QRect(widget.mapToGlobal(QtCore.QPoint(0, 0)), widget.size())
            except Exception:
                continue
            if rect.contains(global_pos):
                before = global_pos.x() < rect.center().x() if thumb_columns > 1 else global_pos.y() < rect.center().y()
                anchor_idx = idx - 1 if before else idx
                return VIRTUAL_COPY_INSERT_START if anchor_idx < 0 else keys[anchor_idx]
            dist = abs(global_pos.x() - rect.center().x()) + abs(global_pos.y() - rect.center().y())
            if best_dist is None or dist < best_dist:
                best_idx = idx
                best_rect = rect
                best_dist = dist
        if best_idx is None or best_rect is None:
            return keys[-1]
        before = global_pos.x() < best_rect.center().x() if thumb_columns > 1 else global_pos.y() < best_rect.center().y()
        anchor_idx = best_idx - 1 if before else best_idx
        return VIRTUAL_COPY_INSERT_START if anchor_idx < 0 else keys[anchor_idx]

    def _handle_thumbnail_drag_event(self, event):
        mime = getattr(event, "mimeData", lambda: None)()
        payload = self._thumbnail_virtual_drag_payload(mime)
        if payload is None:
            if mime is None or not mime.hasUrls():
                return False
            if event.type() in (QtCore.QEvent.DragEnter, QtCore.QEvent.DragMove):
                event.acceptProposedAction()
                return True
            if event.type() != QtCore.QEvent.Drop:
                return False
            handled = self._handle_local_file_mime_drop(mime)
            if handled:
                event.acceptProposedAction()
            else:
                event.ignore()
            return handled
        event_type = event.type()
        if event_type in (QtCore.QEvent.DragEnter, QtCore.QEvent.DragMove):
            event.acceptProposedAction()
            return True
        if event_type != QtCore.QEvent.Drop:
            return False
        try:
            global_pos = event.globalPos()
        except Exception:
            global_pos = QtGui.QCursor.pos()
        anchor_key = self._thumbnail_drop_insert_anchor(global_pos)
        created = self._create_virtual_copy_from_drag_payload(payload, insert_after_key=anchor_key)
        if created:
            event.acceptProposedAction()
        else:
            event.ignore()
        return True

    def _ordered_thumbnail_selection(self):
        selected = set(getattr(self, "thumb_multi_select", set()) or [])
        ordered = []
        for fp in list(getattr(self, "current_thumb_files", []) or []):
            s = str(fp)
            if s in selected:
                ordered.append(s)
        if ordered:
            return ordered
        current = str(getattr(self, "selected_file_for_thumbs", "") or "")
        return [current] if current else []

    def _select_all_thumbnails(self):
        files = [str(fp) for fp in list(getattr(self, "current_thumb_files", []) or []) if str(fp)]
        spectro_files = [str(fp) for fp in list(getattr(self, "current_spectro_thumb_files", []) or []) if str(fp)]
        if not files and not spectro_files:
            return False
        self.thumb_multi_select = set(files)
        self.spectro_thumb_multi_select = set(spectro_files)
        if files:
            self.last_thumb_anchor = files[-1]
        if spectro_files:
            self.last_spectro_thumb_anchor = spectro_files[-1]
        self._refresh_thumb_selection_styles()
        self._refresh_spectro_thumb_selection_styles()
        return True

    def _copy_thumbnail_selection_to_clipboard_files(self):
        targets = self._ordered_thumbnail_selection()
        if not targets:
            return False
        if getattr(self, "_clipboard_copy_worker", None) is not None:
            self._show_toast("Clipboard copy already running...")
            return True
        try:
            max_items = int(self.config.get("clipboard_copy_max_images", 48))
        except Exception:
            max_items = 48
        if max_items < 1:
            max_items = 48
        if len(targets) > max_items:
            self._show_toast(
                f"Selection too large ({len(targets)}). Copying first {max_items} images.",
                duration_ms=1800,
            )
            targets = targets[:max_items]
        try:
            clip_dir = Path(tempfile.gettempdir()) / "sxm_viewer_clipboard"
            clip_dir.mkdir(parents=True, exist_ok=True)
            session_dir = clip_dir / f"multi_{int(time.time() * 1000)}"
            session_dir.mkdir(parents=True, exist_ok=True)
            self._clipboard_export_dir = session_dir
            cfg = self.get_current_detail_config()
            worker = BatchExportWorker(self, targets, cfg, session_dir)
            worker.signals.finished.connect(self._on_clipboard_copy_finished)
            self._clipboard_copy_worker = worker
            self._clipboard_copy_total = len(targets)
            self._show_toast(f"Preparing {len(targets)} image(s) for clipboard...", duration_ms=1200)
            QtCore.QThreadPool.globalInstance().start(worker)
            return True
        except Exception:
            self._clipboard_copy_worker = None
            self._clipboard_copy_total = 0
            self._show_toast("Clipboard copy failed to start", duration_ms=1700)
            return False

    def _on_clipboard_copy_finished(self, saved, errors, cancelled):
        self._clipboard_copy_worker = None
        total = int(getattr(self, "_clipboard_copy_total", 0) or 0)
        self._clipboard_copy_total = 0
        saved = list(saved or [])
        errors = list(errors or [])
        if cancelled:
            self._show_toast("Clipboard copy canceled", duration_ms=1400)
            return
        if not saved:
            if errors:
                self._show_toast("Clipboard copy failed", duration_ms=1800)
            else:
                self._show_toast("No images copied", duration_ms=1400)
            return
        try:
            mime = QtCore.QMimeData()
            urls = [QtCore.QUrl.fromLocalFile(str(Path(p))) for p in saved]
            mime.setUrls(urls)
            mime.setText("\n".join(str(Path(p)) for p in saved))
            QtWidgets.QApplication.clipboard().setMimeData(mime)
        except Exception:
            self._show_toast("Copied files, but clipboard assignment failed", duration_ms=1800)
            return
        if errors:
            self._show_toast(
                f"Copied {len(saved)}/{max(total, len(saved))} images ({len(errors)} skipped)",
                duration_ms=1900,
            )
        else:
            self._show_toast(f"Copied {len(saved)} image file(s)", duration_ms=1400)

    def _update_rubber_band_selection(self, rect, modifiers):
        in_rect = set()
        for key, widget in self.thumb_widgets.items():
            if widget.geometry().intersects(rect):
                in_rect.add(str(key))
        in_spectro_rect = set()
        for key, widget in getattr(self, "spectro_thumb_widgets", {}).items():
            if widget.geometry().intersects(rect):
                in_spectro_rect.add(str(key))

        base = set(getattr(self, '_selection_before_drag', set()) or [])
        base_spectro = set(getattr(self, '_spectro_selection_before_drag', set()) or [])
        if modifiers & QtCore.Qt.ControlModifier:
            new_selection = base.symmetric_difference(in_rect)
            new_spectro_selection = base_spectro.symmetric_difference(in_spectro_rect)
        elif modifiers & QtCore.Qt.ShiftModifier:
            new_selection = base.union(in_rect)
            new_spectro_selection = base_spectro.union(in_spectro_rect)
        else:
            new_selection = in_rect
            new_spectro_selection = in_spectro_rect

        if new_selection == set(getattr(self, "thumb_multi_select", set()) or []) and new_spectro_selection == set(getattr(self, "spectro_thumb_multi_select", set()) or []):
            return
        self.thumb_multi_select = new_selection
        self.spectro_thumb_multi_select = new_spectro_selection
        self._refresh_thumb_selection_styles()
        self._refresh_spectro_thumb_selection_styles()

    def _thumb_dimensions(self):
        return viewer_thumb_ui._thumb_dimensions(self)

    def _resize_thumbnail_scale(self, delta_px):
        return viewer_thumb_ui._resize_thumbnail_scale(self, delta_px)

    def _create_toolbar(self):
        return main_window_toolbar.create_main_toolbar(self)

    def _update_toolbar_actions(self, enabled: bool):
        return main_window_toolbar.update_toolbar_actions(self, enabled)

    def _create_session_activity_strip(self):
        strip = QtWidgets.QFrame(self)
        strip.setObjectName("sessionActivityStrip")
        strip.setVisible(False)
        strip.setStyleSheet(
            """
            QFrame#sessionActivityStrip {
                border: 1px solid #b8cbe8;
                border-radius: 8px;
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 rgba(238,244,255,235),
                    stop:1 rgba(247,250,255,235));
            }
            QLabel#sessionActivityTitle {
                color: #14345f;
                font-weight: 600;
            }
            QLabel#sessionActivityDetail {
                color: #38506f;
            }
            QProgressBar#sessionActivityProgress {
                min-height: 14px;
                border-radius: 7px;
                background: rgba(255,255,255,185);
                border: 1px solid #c4d3ea;
                text-align: center;
                color: #17395f;
                font-weight: 600;
            }
            QProgressBar#sessionActivityProgress::chunk {
                border-radius: 7px;
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #3276ff,
                    stop:1 #67b8ff);
            }
            """
        )
        layout = QtWidgets.QHBoxLayout(strip)
        layout.setContentsMargins(12, 8, 12, 8)
        layout.setSpacing(12)

        text_col = QtWidgets.QVBoxLayout()
        text_col.setContentsMargins(0, 0, 0, 0)
        text_col.setSpacing(2)
        title = QtWidgets.QLabel("Session activity", strip)
        title.setObjectName("sessionActivityTitle")
        detail = QtWidgets.QLabel("", strip)
        detail.setObjectName("sessionActivityDetail")
        detail.setWordWrap(True)
        text_col.addWidget(title)
        text_col.addWidget(detail)

        progress = QtWidgets.QProgressBar(strip)
        progress.setObjectName("sessionActivityProgress")
        progress.setTextVisible(True)
        progress.setMinimumWidth(220)
        progress.setRange(0, 100)
        progress.setValue(0)
        progress.setFormat("%p%")

        layout.addLayout(text_col, 1)
        layout.addWidget(progress, 0)

        self._session_activity_strip = strip
        self._session_activity_title = title
        self._session_activity_detail = detail
        self._session_activity_progress = progress
        return strip

    def _set_session_activity(self, message, detail="", value=None, stage="loading", visible=True, hide_delay_ms=0):
        strip = getattr(self, "_session_activity_strip", None)
        title = getattr(self, "_session_activity_title", None)
        detail_label = getattr(self, "_session_activity_detail", None)
        progress = getattr(self, "_session_activity_progress", None)
        if strip is None or title is None or detail_label is None or progress is None:
            return
        try:
            self._session_activity_hide_timer.stop()
        except Exception:
            pass
        palettes = {
            "loading": ("#3276ff", "#67b8ff", "#b8cbe8"),
            "hydrating": ("#1ea57c", "#63d3b0", "#a8dacd"),
            "popup": ("#d8891f", "#efbc52", "#e4cca5"),
            "complete": ("#2a9d5b", "#68c788", "#b5d8bf"),
            "error": ("#c94b4b", "#ef8686", "#e3bbbb"),
        }
        start_color, end_color, border_color = palettes.get(stage, palettes["loading"])
        try:
            strip.setStyleSheet(
                f"""
                QFrame#sessionActivityStrip {{
                    border: 1px solid {border_color};
                    border-radius: 8px;
                    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                        stop:0 rgba(248,250,255,238),
                        stop:1 rgba(255,255,255,230));
                }}
                QLabel#sessionActivityTitle {{
                    color: #14345f;
                    font-weight: 600;
                }}
                QLabel#sessionActivityDetail {{
                    color: #38506f;
                }}
                QProgressBar#sessionActivityProgress {{
                    min-height: 14px;
                    border-radius: 7px;
                    background: rgba(255,255,255,185);
                    border: 1px solid {border_color};
                    text-align: center;
                    color: #17395f;
                    font-weight: 600;
                }}
                QProgressBar#sessionActivityProgress::chunk {{
                    border-radius: 7px;
                    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                        stop:0 {start_color},
                        stop:1 {end_color});
                }}
                """
            )
        except Exception:
            pass
        title.setText(str(message or "Session activity"))
        detail_label.setText(str(detail or ""))
        if value is None:
            progress.setRange(0, 0)
            progress.setFormat("")
        else:
            progress.setRange(0, 100)
            progress.setValue(max(0, min(100, int(value))))
            progress.setFormat("%p%")
        strip.setVisible(bool(visible))
        if hide_delay_ms and visible:
            self._session_activity_hide_timer.start(int(max(0, hide_delay_ms)))

    def _hide_session_activity(self):
        strip = getattr(self, "_session_activity_strip", None)
        progress = getattr(self, "_session_activity_progress", None)
        if progress is not None:
            try:
                progress.setRange(0, 100)
                progress.setValue(0)
                progress.setFormat("%p%")
            except Exception:
                pass
        if strip is not None:
            strip.setVisible(False)

    def _describe_deferred_popup_entry(self, entry):
        if not isinstance(entry, dict):
            return "Deferred pop-up"
        title = str(entry.get("title") or "").strip()
        if title:
            return title
        snapshot = entry.get("snapshot") or {}
        title = str(snapshot.get("window_title") or "").strip()
        if title:
            return title
        first_view = ((snapshot.get("views") or [{}]) or [{}])[0]
        meta = first_view.get("meta") or {}
        file_name = meta.get("file_name") or Path(str(first_view.get("path") or "")).name
        channel = meta.get("channel") or ""
        if file_name and channel:
            return f"{file_name} | {channel}"
        if file_name:
            return str(file_name)
        return "Deferred pop-up"

    def _active_popup_windows(self):
        controller = getattr(self, "quick_crop_controller", None)
        if controller is not None and hasattr(controller, "tracked_popups"):
            try:
                return list(controller.tracked_popups())
            except Exception:
                pass
        active = []
        for dlg in list(self._iter_workspace_windows(include_canvas=False)):
            try:
                if dlg is not None and (dlg.isVisible() or dlg.isMinimized()):
                    active.append(dlg)
            except Exception:
                continue
        return active

    def _popup_window_label(self, dlg):
        if dlg is None:
            return "Pop-up"
        try:
            title = str(dlg.windowTitle() or "").strip()
        except Exception:
            title = ""
        if title:
            return title
        try:
            fallback = str(type(dlg).__name__ or "").strip()
        except Exception:
            fallback = ""
        return fallback or "Pop-up"

    def _popup_window_menu_label(self, dlg):
        text = self._popup_window_label(dlg)
        try:
            if dlg is not None and (dlg.windowState() & QtCore.Qt.WindowMinimized):
                text = f"{text} [minimized]"
        except Exception:
            pass
        return text

    def _focus_popup_window(self, dlg):
        controller = getattr(self, "quick_crop_controller", None)
        if controller is not None and hasattr(controller, "focus_popup"):
            try:
                return bool(controller.focus_popup(dlg))
            except Exception:
                pass
        if dlg is None:
            return False
        try:
            state = dlg.windowState()
            if state & QtCore.Qt.WindowMinimized:
                dlg.showNormal()
            else:
                dlg.show()
            dlg.raise_()
            dlg.activateWindow()
        except Exception:
            return False
        return True

    def _rebuild_popup_menu(self):
        menu = getattr(self, "toolbar_popups_menu", None)
        if menu is None:
            return
        menu.clear()
        active = list(self._active_popup_windows())
        entries = list(getattr(self, "_deferred_popup_entries", []) or [])
        if not active and not entries:
            empty_act = menu.addAction("No open or saved pop-ups")
            empty_act.setEnabled(False)
            return
        metrics = None
        try:
            metrics = self.fontMetrics()
        except Exception:
            metrics = None
        if active:
            header_act = menu.addAction(f"{len(active)} open pop-out{'s' if len(active) != 1 else ''}")
            header_act.setEnabled(False)
            menu.addSeparator()
            bring_act = menu.addAction("Bring all to front")
            bring_act.setShortcut(QtGui.QKeySequence("Ctrl+Shift+P"))
            bring_act.setToolTip("Restore minimized pop-outs and raise the open pop-out windows")
            bring_act.triggered.connect(self.on_recall_popouts)
            arrange_act = menu.addAction("Arrange all")
            arrange_act.triggered.connect(self.on_arrange_popouts)
            minimize_act = menu.addAction("Minimize all")
            minimize_act.triggered.connect(self.on_minimize_popouts)
            close_act = menu.addAction("Close all")
            close_act.triggered.connect(self.on_close_popouts)
            menu.addSeparator()
            for dlg in active:
                full_text = self._popup_window_menu_label(dlg)
                text = full_text
                if metrics is not None:
                    try:
                        text = metrics.elidedText(full_text, QtCore.Qt.ElideRight, 340)
                    except Exception:
                        pass
                act = menu.addAction(text)
                act.setToolTip(full_text)
                act.triggered.connect(partial(self._focus_popup_window, dlg))
        if entries:
            if active:
                menu.addSeparator()
            header_act = menu.addAction(f"{len(entries)} saved pop-up{'s' if len(entries) != 1 else ''}")
            header_act.setEnabled(False)
            menu.addSeparator()
            restore_all = menu.addAction(f"Restore all saved ({len(entries)})")
            restore_all.triggered.connect(lambda: self.session_controller.restore_all_deferred_popups())
            menu.addSeparator()
            for entry in entries:
                text = self._describe_deferred_popup_entry(entry)
                if metrics is not None:
                    try:
                        text = metrics.elidedText(text, QtCore.Qt.ElideRight, 340)
                    except Exception:
                        pass
                act = menu.addAction(text)
                act.triggered.connect(partial(self.session_controller.restore_deferred_popup, entry_id=entry.get("id")))

    def _refresh_popup_ui(self, *, popups=None):
        btn = getattr(self, "toolbar_popups_btn", None)
        action = getattr(self, "toolbar_popups_raise_act", None)
        if btn is None and action is None:
            return
        active = list(popups) if popups is not None else self._active_popup_windows()
        active_count = len(active)
        saved_count = len(getattr(self, "_deferred_popup_entries", []) or [])
        enabled = bool(active_count or saved_count)
        if active_count:
            label = f"Pop-ups ({active_count} open)"
        elif saved_count:
            label = f"Pop-ups ({saved_count} saved)"
        else:
            label = "Pop-ups"
        if active_count and saved_count:
            tool_tip = (
                f"Click to bring {active_count} open pop-out(s) to the front. "
                f"Use the arrow for window actions and {saved_count} saved pop-up(s)."
            )
        elif active_count:
            tool_tip = (
                f"Click to bring {active_count} open pop-out(s) to the front "
                "or use the arrow for arrange/minimize/focus actions."
            )
        elif saved_count:
            tool_tip = (
                f"Click to restore {saved_count} saved pop-up(s) "
                "or use the arrow for individual restore actions."
            )
        else:
            tool_tip = "No open or saved pop-ups are available."
        if action is not None:
            action.setText(label)
            action.setToolTip(tool_tip)
            action.setStatusTip(tool_tip)
            action.setEnabled(enabled)
        if btn is not None:
            btn.setEnabled(enabled)
            btn.setToolTip(tool_tip)
        try:
            if btn is not None:
                btn.setStyleSheet(
                    "QToolButton { font-weight: 600; color: #0f4c81; }"
                    if active_count
                    else ("QToolButton { font-weight: 600; color: #7a4d00; }" if saved_count else "")
                )
        except Exception:
            pass
        self._rebuild_popup_menu()

    def _rebuild_deferred_popup_menu(self):
        self._rebuild_popup_menu()

    def _refresh_deferred_popup_ui(self):
        self._refresh_popup_ui()

    def _step_channel(self, delta: int):
        combo = getattr(self, "channel_dropdown", None)
        if combo is None or combo.count() <= 0 or not combo.isEnabled():
            return
        current = combo.currentIndex()
        if current < 0:
            current = 0
        target = max(0, min(combo.count() - 1, current + int(delta)))
        if target != current:
            combo.setCurrentIndex(target)

    def _sync_channel_nav_buttons(self):
        combo = getattr(self, "channel_dropdown", None)
        prev_btn = getattr(self, "channel_prev_btn", None)
        next_btn = getattr(self, "channel_next_btn", None)
        if combo is None:
            return
        has_channels = bool(combo.isEnabled()) and combo.count() > 0
        current = combo.currentIndex()
        if prev_btn is not None:
            prev_btn.setEnabled(has_channels and current > 0)
        if next_btn is not None:
            next_btn.setEnabled(has_channels and 0 <= current < combo.count() - 1)

    def _on_toggle_layout_mode(self):
        target = "stacked" if self._layout_mode == "columns" else "columns"
        self._apply_layout_mode(target)

    def _apply_layout_mode(self, mode: str):
        if not hasattr(self, "main_splitter"):
            return
        if mode not in ("columns", "stacked"):
            mode = "columns"
        # preserve sizes
        if hasattr(self, "_layout_mode"):
            self._layout_sizes[self._layout_mode] = self.main_splitter.sizes()
        # detach all but left
        for idx in reversed(range(self.main_splitter.count())):
            widget = self.main_splitter.widget(idx)
            if widget is getattr(self, "left_w", None):
                continue
            widget.setParent(None)
        if mode == "columns":
            # reattach panels directly
            if self._thumbs_panel.parent() is not None and self._thumbs_panel.parent() is not self.main_splitter:
                self._thumbs_panel.setParent(None)
            if self._preview_panel.parent() is not None and self._preview_panel.parent() is not self.main_splitter:
                self._preview_panel.setParent(None)
            self.main_splitter.addWidget(self._thumbs_panel)
            self.main_splitter.addWidget(self._preview_panel)
            self.main_splitter.setStretchFactor(0, 1)
            self.main_splitter.setStretchFactor(1, 2)
            self.main_splitter.setStretchFactor(2, 3)
            try:
                self.preview_canvas.set_view_layout("stacked")
            except Exception:
                pass
        else:
            # stack thumbs + preview vertically on the right
            if self._thumbs_panel.parent() is not self._right_splitter:
                self._thumbs_panel.setParent(None)
                self._right_splitter.insertWidget(0, self._thumbs_panel)
            if self._preview_panel.parent() is not self._right_splitter:
                self._preview_panel.setParent(None)
                self._right_splitter.addWidget(self._preview_panel)
            self.main_splitter.addWidget(self._right_container)
            self.main_splitter.setStretchFactor(0, 1)
            self.main_splitter.setStretchFactor(1, 3)
            try:
                self.preview_canvas.set_view_layout("grid")
            except Exception:
                pass
        self._layout_mode = mode
        if hasattr(self, "toolbar_layout_act"):
            self.toolbar_layout_act.setText("Layout: Columns" if mode == "columns" else "Layout: Stack")
        sizes = self._layout_sizes.get(mode)
        if sizes:
            self.main_splitter.setSizes(sizes)
        else:
            if mode == "columns":
                self.main_splitter.setSizes(list(MAIN_SPLITTER_SIZES_COLUMNS))
            else:
                self.main_splitter.setSizes(list(MAIN_SPLITTER_SIZES_STACKED))

    def on_dark_mode_toggled(self, checked: bool):
        self.dark_mode = bool(checked)
        try:
            if hasattr(self, 'toolbar_dark_btn'):
                self.toolbar_dark_btn.blockSignals(True)
                self.toolbar_dark_btn.setChecked(self.dark_mode)
                self.toolbar_dark_btn.blockSignals(False)
        except Exception:
            pass
        if getattr(self, "_detail_theme_follows_dark_mode", True):
            self._set_detail_dark_view_state(self.dark_mode, follow_dark_mode=True, persist=False)
        self.config['dark_mode'] = self.dark_mode
        self.config['detail_dark_view'] = self.detail_dark_view
        self.config['detail_theme_follows_dark_mode'] = bool(
            getattr(self, "_detail_theme_follows_dark_mode", True)
        )
        save_config(self.config)
        self._apply_dark_mode(self.dark_mode)
        if self.last_preview:
            self.show_file_channel(self.last_preview[0], self.last_preview[1])

    # ---------- folder load & auto-detection ----------
    def open_folder_dialog(self):
        d = QtWidgets.QFileDialog.getExistingDirectory(self, "Select data folder", str(self.last_dir))
        if d:
            self.load_folder(Path(d))

    def open_folder_by_path(self):
        p = Path(self.path_le.text().strip())
        if p.exists() and p.is_dir():
            self.load_folder(p)

    def _refresh_recent_dirs_menu(self):
        menu = getattr(self, "open_recent_menu", None)
        if menu is None:
            return
        menu.clear()
        recents = getattr(self, "recent_dirs", [])
        if not recents:
            act = menu.addAction("No recent folders")
            act.setEnabled(False)
            return
        for path in recents:
            act = menu.addAction(path)
            act.setToolTip(path)
            act.triggered.connect(lambda checked=False, p=path: self.load_folder(Path(p)))
        menu.addSeparator()
        clear_act = menu.addAction("Clear recent folders")
        clear_act.triggered.connect(self._clear_recent_dirs)

    def _record_recent_dir(self, folder: Path):
        folder_path = Path(folder)
        folder_str = str(folder_path)
        recents = []
        for p in getattr(self, "recent_dirs", []):
            if not p:
                continue
            try:
                if Path(p).resolve() == folder_path.resolve():
                    continue
            except Exception:
                if p == folder_str:
                    continue
            recents.append(p)
        recents.insert(0, folder_str)
        self.recent_dirs = recents[:RECENT_FOLDER_LIMIT]
        self.config["recent_dirs"] = self.recent_dirs
        save_config(self.config)
        self._refresh_recent_dirs_menu()

    def _clear_recent_dirs(self):
        self.recent_dirs = []
        self.config["recent_dirs"] = []
        save_config(self.config)
        self._refresh_recent_dirs_menu()

    def _refresh_recent_session_dirs_menu(self):
        menus = [
            getattr(self, "load_session_recent_menu", None),
            getattr(self, "toolbar_load_session_menu", None),
        ]
        for menu in menus:
            if menu is None:
                continue
            menu.clear()
            recents = getattr(self, "recent_session_paths", [])
            if not recents:
                act = menu.addAction("No recent sessions")
                act.setEnabled(False)
                continue
            for path in recents:
                try:
                    session_path = Path(path)
                except Exception:
                    session_path = Path(str(path))
                act = menu.addAction(str(session_path))
                act.setToolTip(str(session_path))
                act.triggered.connect(
                    lambda checked=False, p=str(session_path): self.on_load_recent_session(Path(p))
                )
            menu.addSeparator()
            clear_act = menu.addAction("Clear recent sessions")
            clear_act.triggered.connect(self._clear_recent_session_dirs)

    def _normalize_recent_session_history(self, persist=False):
        recents = []
        has_session_file = False
        for p in getattr(self, "recent_session_paths", []):
            if not p:
                continue
            try:
                candidate = Path(p)
            except Exception:
                continue
            if str(candidate) not in recents:
                recents.append(str(candidate))
            try:
                if candidate.suffix.lower() == ".json":
                    has_session_file = True
            except Exception:
                pass
        if has_session_file:
            recents = [p for p in recents if Path(p).suffix.lower() == ".json"]
        recents = recents[:RECENT_SESSION_LIMIT]
        changed = recents != list(getattr(self, "recent_session_paths", []))
        self.recent_session_paths = recents
        if persist and changed:
            self.config["recent_session_paths"] = self.recent_session_paths
            self.config.pop("recent_session_dirs", None)
            save_config(self.config)
        return changed

    def _record_recent_session(self, session_path: Path):
        session_path = Path(session_path)
        session_str = str(session_path)
        recents = []
        for p in getattr(self, "recent_session_paths", []):
            if not p:
                continue
            try:
                if Path(p).resolve() == session_path.resolve():
                    continue
            except Exception:
                if p == session_str:
                    continue
            recents.append(p)
        recents.insert(0, session_str)
        self.recent_session_paths = recents[:RECENT_SESSION_LIMIT]
        self._normalize_recent_session_history(persist=False)
        self.config["recent_session_paths"] = self.recent_session_paths
        self.config.pop("recent_session_dirs", None)
        save_config(self.config)
        self._refresh_recent_session_dirs_menu()

    def _clear_recent_session_dirs(self):
        self.recent_session_paths = []
        self.config["recent_session_paths"] = []
        self.config.pop("recent_session_dirs", None)
        save_config(self.config)
        self._refresh_recent_session_dirs_menu()

    def _resolve_recent_session_target(self, session_path: Path):
        session_path = Path(session_path)
        if session_path.is_file():
            return session_path
        if not session_path.is_dir():
            return None
        preferred = session_path / "sxm_session.json"
        if preferred.exists():
            return preferred
        try:
            json_files = sorted(
                (p for p in session_path.glob("*.json") if p.is_file()),
                key=lambda p: p.stat().st_mtime,
                reverse=True,
            )
        except Exception:
            json_files = []
        if json_files:
            return json_files[0]
        return None

    def on_load_recent_session(self, session_path: Path):
        session_path = Path(session_path)
        resolved = self._resolve_recent_session_target(session_path)
        if resolved is not None:
            self.session_controller.load_session(session_path=resolved)
            return
        if session_path.exists() and session_path.is_dir():
            self.session_controller.load_session(start_dir=session_path)
            return
        self.session_controller.load_session(session_path=session_path)

    def _on_recent_molecules_updated(self, paths):
        """Persist recent molecule file paths to config (up to 8)."""
        try:
            recent = []
            for p in paths or []:
                if p and p not in recent:
                    recent.append(p)
                if len(recent) >= 8:
                    break
            self.recent_molecules = recent
            self.config["recent_molecules"] = recent
            save_config(self.config)
        except Exception:
            pass

    def _load_recent_molecule(self, path):
        if not self.preview_canvas or not path:
            return
        try:
            self.preview_canvas.add_molecule(path)
            self.on_show_molecules_toggled(True)
            self._on_recent_molecules_updated(self.preview_canvas.get_recent_molecule_paths())
        except Exception:
            pass

    def _clear_preview_molecules(self):
        if not self.preview_canvas:
            return
        try:
            self.preview_canvas.reset_molecules()
        except Exception:
            try:
                self.preview_canvas._clear_molecules()
            except Exception:
                pass

    def _populate_browse_molecules_menu(self):
        menu = getattr(self, "browse_molecules_menu", None)
        if menu is None:
            return
        try:
            menu.clear()
        except Exception:
            return

        show_act = menu.addAction("Show molecules")
        show_act.setCheckable(True)
        show_act.setChecked(bool(getattr(self, "show_molecules", True)))
        show_act.toggled.connect(self.on_show_molecules_toggled)

        menu.addSeparator()

        load_act = menu.addAction("Load molecule...")
        load_act.triggered.connect(self.on_load_molecule)

        recent = []
        try:
            if self.preview_canvas is not None:
                recent = list(self.preview_canvas.get_recent_molecule_paths() or [])
        except Exception:
            recent = list(getattr(self, "recent_molecules", []) or [])
        if recent:
            recent_menu = menu.addMenu("Load recent")
            for path in recent[:8]:
                act = recent_menu.addAction(Path(path).name)
                act.setToolTip(str(path))
                act.triggered.connect(lambda _checked=False, p=path: self._load_recent_molecule(p))

        clear_act = menu.addAction("Clear molecules")
        clear_act.setEnabled(bool(getattr(getattr(self, "preview_canvas", None), "molecules", []) or []))
        clear_act.triggered.connect(self._clear_preview_molecules)

        palette_menu = menu.addMenu("Palette")
        palette_group = QtWidgets.QActionGroup(palette_menu)
        current_palette = str(getattr(self, "molecule_palette", "cpk") or "cpk").lower()
        for palette in available_atom_palettes():
            label = {"cpk": "CPK", "pymol": "PyMOL", "jmol": "Jmol"}.get(palette, palette.capitalize())
            act = palette_menu.addAction(label)
            act.setCheckable(True)
            act.setChecked(current_palette == palette)
            act.triggered.connect(lambda checked=False, p=palette: checked and self._on_molecule_palette_changed(p))
            palette_group.addAction(act)

    def _on_toggle_preview_title(self, checked):
        """Toggle title/date overlay in Preview and pop-outs."""
        self.show_preview_title = bool(checked)
        self.config['show_preview_title'] = self.show_preview_title
        save_config(self.config)
        try:
            self.preview_canvas.set_show_title(self.show_preview_title)
        except Exception:
            pass

    def _classify_topography_values(self, vals, tolerance_nm: float | None = None):
        """Simple rule: CH if any full row (finite pixels) is exactly flat; otherwise CC. 1D stays CC unless fully flat."""
        try:
            arr = np.asarray(vals, dtype=float)
        except Exception:
            return None
        if arr.ndim == 2:
            if arr.shape[0] == 0:
                return None
            def _is_flat(row):
                row_fin = row[np.isfinite(row)]
                return row_fin.size > 0 and np.ptp(row_fin) == 0.0
            top_flat = _is_flat(arr[0])
            bottom_flat = _is_flat(arr[-1])
            median = float(np.nanmedian(arr))
            prange = float(np.nanmax(arr) - np.nanmin(arr))
            if top_flat and bottom_flat:
                abs_pm = int(round(median * 1000.0))
                return {'tag': 'constant-height', 'abs_pm': abs_pm, 'rng_nm': prange, 'median_nm': median}
            return {'tag': 'constant-current', 'rng_nm': prange, 'median_nm': median}
        else:
            arr = arr[np.isfinite(arr)]
            if arr.size == 0:
                return None
            if np.ptp(arr) == 0.0:
                median = float(np.nanmedian(arr))
                abs_pm = int(round(median * 1000.0))
                return {'tag': 'constant-height', 'abs_pm': abs_pm, 'rng_nm': 0.0, 'median_nm': median}
            median = float(np.nanmedian(arr))
            prange = float(np.nanmax(arr) - np.nanmin(arr))
            return {'tag': 'constant-current', 'rng_nm': prange, 'median_nm': median}

    def _auto_preview_clim(self, arr, *, relative_zero: bool = False):
        """Compute color limits ignoring a dominant flat stripe (e.g., aborted scans)."""
        try:
            data = np.asarray(arr, dtype=float)
            finite = data[np.isfinite(data)]
            if finite.size == 0:
                return None
            hist, edges = np.histogram(finite, bins=256)
            idx_max = int(np.argmax(hist))
            frac = hist[idx_max] / float(finite.size)
            if frac > 0.5:
                lo_edge, hi_edge = edges[idx_max], edges[idx_max + 1]
                finite = finite[(finite < lo_edge) | (finite > hi_edge)]
                if finite.size == 0:
                    finite = data[np.isfinite(data)]
            vmin = float(np.nanpercentile(finite, 1.0))
            vmax = float(np.nanpercentile(finite, 99.0))
            if relative_zero:
                vmin = 0.0
            if vmin == vmax:
                return None
            return (vmin, vmax)
        except Exception:
            return None

    def load_folder(self, folder:Path):
        start = time.perf_counter()
        self._workspace_loading = True
        try:
            self._prepare_for_workspace_load(kind="folder")
            result = viewer_loader.load_folder(self, folder)
        finally:
            self._workspace_loading = False
        end = time.perf_counter()
        folder_ms = (end - start) * 1000.0
        gui_ms = (end - getattr(self, "_app_start_ts", start)) * 1000.0
        log_status(f"[Perf] Load folder: {folder_ms:.0f} ms | since GUI init: {gui_ms:.0f} ms")
        if self.auto_detect_tags:
            try:
                self._auto_detect_tags_for_folder()
            except Exception:
                pass
        return result

    def load_files(self, files, folder_hint: Path | None = None, *, append: bool = False, refresh_spectros: bool = True):
        start = time.perf_counter()
        self._workspace_loading = True
        try:
            result = viewer_loader.load_files(
                self,
                files,
                folder_hint=folder_hint,
                source_label="drop" if append else "files",
                append=append,
                refresh_spectros=refresh_spectros,
            )
        finally:
            self._workspace_loading = False
        end = time.perf_counter()
        files_ms = (end - start) * 1000.0
        gui_ms = (end - getattr(self, "_app_start_ts", start)) * 1000.0
        log_status(f"[Perf] Load files: {files_ms:.0f} ms | since GUI init: {gui_ms:.0f} ms")
        if self.auto_detect_tags:
            try:
                self._auto_detect_tags_for_folder()
            except Exception:
                pass
        return result

    def load_spectroscopy_files(self, files, folder_hint: Path | None = None, *, append: bool = True, refresh: bool = True):
        start = time.perf_counter()
        result = viewer_loader.load_spectroscopy_files(
            self,
            files,
            folder_hint=folder_hint,
            append=append,
            refresh=refresh,
        )
        end = time.perf_counter()
        ms = (end - start) * 1000.0
        gui_ms = (end - getattr(self, "_app_start_ts", start)) * 1000.0
        log_status(f"[Perf] Load spectros: {ms:.0f} ms | since GUI init: {gui_ms:.0f} ms")
        return result

    def clear_loaded_images(self):
        """Clear the current image session and leave the app ready for fresh drops."""
        self.files = []
        self.headers.clear()
        self.frame_map_entries = []
        self.hidden_frame_keys.clear()
        self.selected_file_for_thumbs = None
        self.current_inspector_header = None
        self.current_inspector_channel = None
        self.last_preview = None
        self.current_thumb_files = []
        self.thumb_multi_select = set()
        self.spectro_thumb_multi_select = set()
        self.current_spectro_thumb_files = []
        self.selected_spectro_thumb_file = None
        self.thumbnail_filters = {}
        self.virtual_copies = {}
        self.virtual_copy_order = []
        self.added_views = []
        self.extra_view_specs = []
        self.molecule_overlays = {}
        self.frame_entry_pixmaps = {}
        self._frame_real_pixmap_cache = {}
        self._processed_views = {}
        self._collection_item_snapshots = {}
        self._workspace_kind = "folder"
        self._current_session_path = None
        self.matrix_datasets = {}
        self._spectro_hist_cache = {}
        self._last_base_array = None
        self._last_base_extent = None
        self._last_base_unit = None
        self.spectros = []
        self.matrix_spectros = []
        self.spectros_by_image = defaultdict(list)
        self.files_with_matrix = set()
        self._spectros_loaded = False
        self._spectros_pending = False
        self.spectro_thumb_channel_by_path = {}
        self._deferred_popup_entries = []
        self._deferred_popup_serial = 0
        try:
            self._thumb_generation += 1
        except Exception:
            pass
        self._invalidate_thumbnail_cache()
        self._invalidate_channel_cache()
        self._update_toolbar_actions(False)
        try:
            self.clear_thumbs()
        except Exception:
            pass
        try:
            self.preview_canvas.set_views([])
        except Exception:
            pass
        try:
            self.preview_value_label.setText("Value: --")
            self.angle_value_label.setText("Angle: --")
        except Exception:
            pass
        try:
            self.meta_box.clear()
        except Exception:
            pass
        try:
            self.channel_dropdown.blockSignals(True)
            self.channel_dropdown.clear()
            self.channel_dropdown.setEnabled(False)
        except Exception:
            pass
        finally:
            try:
                self.channel_dropdown.blockSignals(False)
            except Exception:
                pass
        self._sync_channel_nav_buttons()
        try:
            self.frame_map_widget.set_entries([])
            self.frame_map_widget.clear_hidden_entries()
        except Exception:
            pass
        try:
            self._update_spectro_stats_label()
        except Exception:
            pass
        try:
            self._refresh_deferred_popup_ui()
        except Exception:
            pass
        try:
            self._hide_session_activity()
        except Exception:
            pass

    def _spectroscopy_metadata_lines(self, spec):
        """Format spectroscopy metadata for the Details panel without dumping large arrays."""
        if not spec:
            return ["No spectroscopy metadata."]

        def _fmt(value):
            if value is None:
                return "None"
            if isinstance(value, (str, int, float, bool)):
                return str(value)
            if isinstance(value, Path):
                return str(value)
            if isinstance(value, dict):
                return f"dict({len(value)})"
            if isinstance(value, (list, tuple, set)):
                return f"{type(value).__name__}({len(value)})"
            if hasattr(value, "shape"):
                try:
                    arr = np.asarray(value)
                    return f"array(shape={arr.shape}, dtype={arr.dtype})"
                except Exception:
                    return type(value).__name__
            return str(value)

        lines = ["Spectroscopy details", ""]
        for key in ("path", "source", "time", "file_mtime", "image_key", "matrix_dataset", "matrix_index", "x", "y", "AxisLabel", "AxisUnit", "AltAxisLabel", "AltAxisUnit"):
            if key in spec:
                lines.append(f"{key}: {_fmt(spec.get(key))}")
        channels = spec.get("channels") or {}
        if channels:
            lines.append("")
            lines.append(f"Channels ({len(channels)}):")
            for name, values in channels.items():
                try:
                    arr = np.asarray(values)
                    shape = arr.shape
                except Exception:
                    shape = "?"
                lines.append(f"  - {name}: shape={shape}")
        axis_choices = spec.get("AxisChoices") or []
        if axis_choices:
            lines.append("")
            lines.append(f"Axis choices ({len(axis_choices)}):")
            for ax in axis_choices:
                key = ax.get("key") or ax.get("label") or "Axis"
                label = ax.get("label") or "Axis"
                unit = ax.get("unit") or ""
                lines.append(f"  - {key}: {label}" + (f" ({unit})" if unit else ""))
        lines.append("")
        lines.append("Raw fields:")
        for key in sorted(spec.keys(), key=lambda s: str(s).lower()):
            if key in {"channels", "AxisChoices"}:
                continue
            lines.append(f"  {key}: {_fmt(spec.get(key))}")
        return lines

    def show_spectroscopy_details(self, spec):
        """Show a spectroscopy entry in the left Details panel."""
        try:
            self.meta_box.setPlainText("\n".join(self._spectroscopy_metadata_lines(spec)))
        except Exception:
            pass
        try:
            if hasattr(self, "details_group"):
                self.details_group.setChecked(True)
        except Exception:
            pass

    def on_set_spectro_thumbnail_channel(self, channel_name: str, paths=None):
        """Set the rendered spectroscopy channel for one or more miniature cards."""
        channel_name = str(channel_name or "").strip()
        if not channel_name:
            return
        targets = [str(Path(p)) for p in (paths or []) if p]
        if targets:
            for key in targets:
                self.spectro_thumb_channel_by_path[key] = channel_name
            self.config["spectro_thumb_channel_by_path"] = self.spectro_thumb_channel_by_path
        else:
            self.spectro_miniature_default_channel = channel_name
            self.config["spectro_miniature_default_channel"] = channel_name
        save_config(self.config)
        self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())

    def set_plot_typography(self, *, family=None, bold=None, italic=None, underline=None, refresh: bool = True):
        """Set shared plot typography and redraw visible plot surfaces."""
        changed = False
        style_changes = {
            "bold": bold,
            "italic": italic,
            "underline": underline,
        }
        if family is not None:
            family = normalize_font_family(family, UI_FONT_FAMILY)
            if family != getattr(self, "_plot_font_family", None):
                self._plot_font_family = family
                self.config["plot_font_family"] = family
                changed = True
        for key, attr in (("bold", "_plot_font_bold"), ("italic", "_plot_font_italic"), ("underline", "_plot_font_underline")):
            val = style_changes.get(key)
            if val is not None and bool(val) != getattr(self, attr, False):
                setattr(self, attr, bool(val))
                self.config[f"plot_font_{key}"] = bool(val)
                changed = True
        if changed:
            save_config(self.config)
        set_matplotlib_font_family(self._plot_font_family)
        if not refresh:
            return {
                "family": self._plot_font_family,
                "bold": self._plot_font_bold,
                "italic": self._plot_font_italic,
                "underline": self._plot_font_underline,
            }
        canvases = [getattr(self, "preview_canvas", None)] + list(getattr(self, "_popup_canvases", []))
        for canv in canvases:
            if canv is None:
                continue
            try:
                if hasattr(canv, "set_plot_typography"):
                    canv.set_plot_typography(
                        family=self._plot_font_family,
                        bold=self._plot_font_bold,
                        italic=self._plot_font_italic,
                        underline=self._plot_font_underline,
                    )
                else:
                    canv.set_plot_font_family(self._plot_font_family)
            except Exception:
                try:
                    canv._redraw()
                except Exception:
                    pass
        profile_dialogs = []
        main_profile_dialog = getattr(self, "_profile_dialog", None)
        if main_profile_dialog is not None:
            profile_dialogs.append(main_profile_dialog)
        profile_dialogs.extend(list(getattr(self, "_profile_dialogs", []) or []))
        seen_profile_dialogs = set()
        for dlg in profile_dialogs:
            if dlg is None:
                continue
            dlg_id = id(dlg)
            if dlg_id in seen_profile_dialogs:
                continue
            seen_profile_dialogs.add(dlg_id)
            try:
                if hasattr(dlg, "set_plot_typography"):
                    dlg.set_plot_typography(
                        family=self._plot_font_family,
                        bold=self._plot_font_bold,
                        italic=self._plot_font_italic,
                        underline=self._plot_font_underline,
                    )
                else:
                    dlg.set_plot_font_family(self._plot_font_family)
            except Exception:
                pass
        for dlg in list(getattr(self, "_spectro_popups", []) or []):
            if dlg is None:
                continue
            try:
                if hasattr(dlg, "set_plot_typography"):
                    dlg.set_plot_typography(
                        family=self._plot_font_family,
                        bold=self._plot_font_bold,
                        italic=self._plot_font_italic,
                        underline=self._plot_font_underline,
                    )
                else:
                    dlg.set_plot_font_family(self._plot_font_family)
            except Exception:
                pass
        for dlg in list(getattr(self, "_multi_spectro_popups", []) or []):
            if dlg is None:
                continue
            try:
                if hasattr(dlg, "set_plot_typography"):
                    dlg.set_plot_typography(
                        family=self._plot_font_family,
                        bold=self._plot_font_bold,
                        italic=self._plot_font_italic,
                        underline=self._plot_font_underline,
                    )
                else:
                    dlg.set_plot_font_family(self._plot_font_family)
            except Exception:
                pass
        return self._plot_font_family

    def set_plot_font_family(self, family: str, *, refresh: bool = True):
        """Backward-compatible wrapper for shared font family updates."""
        return self.set_plot_typography(family=family, refresh=refresh)

    def _auto_detect_tags_for_folder(self):
        """Auto-detect CH/CC (topography variance rule) for the current folder."""
        for p in self.files:
            key = str(p)
            tag_info = self.tags.get(key, {})
            if tag_info.get('manual'):
                continue  # keep user overrides
            hdr, fds = self.headers.get(key, (None, None))
            if not fds:
                continue

            topo_idx = _find_topography_channel(fds)
            if topo_idx is None and len(fds) > 0:
                topo_idx = 0
            if topo_idx is None:
                continue

            fd = fds[topo_idx]
            try:
                raw_arr = self._get_channel_array(key, topo_idx, hdr, fd)
            except Exception:
                continue
            _, arr_nm = normalize_unit_and_data(raw_arr, fd.get('PhysUnit',''))
            tag_info = self._classify_topography_values(arr_nm)
            if not tag_info:
                continue
            info = {'tag': tag_info['tag'], 'auto': True, 'rng_nm': tag_info.get('rng_nm')}
            if tag_info['tag'] == 'constant-height':
                info['abs_z_pm'] = tag_info.get('abs_pm')
            self.tags[key] = info

        # persist tags after the initial auto pass
        self.config['tags'] = self.tags
        save_config(self.config)

    # ---------- thumbnails population with badge overlay ----------
    def clear_thumbs(self):
        return viewer_thumb_ui.clear_thumbs(self)

    def populate_thumbnails_for_channel(self, channel_idx:int):
        return viewer_thumb_ui.populate_thumbnails_for_channel(self, channel_idx)

    def _thumbnail_filter_signature(self, file_key):
        return viewer_thumbnails._thumbnail_filter_signature(self, file_key)

    def _downsample_for_thumbnail(self, arr, thumb_w, thumb_h):
        return viewer_thumbnails._downsample_for_thumbnail(self, arr, thumb_w, thumb_h)

    def _map_spec_to_pixels(self, spec, header, xpix, ypix, file_key=None):
        return viewer_preview._map_spec_to_pixels(self, spec, header, xpix, ypix, file_key=file_key)

    def _matrix_bbox_pixels(self, m_specs, header, xpix, ypix, w_scale, h_scale, file_key=None):
        return viewer_preview._matrix_bbox_pixels(self, m_specs, header, xpix, ypix, w_scale, h_scale, file_key=file_key)

    def _fallback_spec_coords(self, idx, xpix, ypix):
        return viewer_preview._fallback_spec_coords(self, idx, xpix, ypix)

    def _decorate_thumbnail_pixmap(self, pix, file_key, channel_idx, header, fds, thumb_crop=None):
        """Draw tag borders, filter badges, and spectroscopy markers."""
        marker_defs = []
        taginfo = self.tags.get(str(file_key), {})
        if taginfo:
            tag = taginfo.get('tag')
            painter = QtGui.QPainter(pix)
            pen = QtGui.QPen()
            pen.setWidth(4)
            if tag == 'constant-height':
                pen.setColor(QtGui.QColor(0, 180, 0))
                painter.setPen(pen)
                painter.drawRect(2, 2, pix.width() - 5, pix.height() - 5)
                painter.setFont(QtGui.QFont("Segoe UI", 9, QtGui.QFont.Bold))
                painter.setPen(QtGui.QColor(255, 255, 255))
                painter.drawText(6, 18, "CH")
            elif tag == 'constant-current':
                pen.setColor(QtGui.QColor(30, 100, 200))
                painter.setPen(pen)
                painter.drawRect(2, 2, pix.width() - 5, pix.height() - 5)
                painter.setFont(QtGui.QFont("Segoe UI", 9, QtGui.QFont.Bold))
                painter.setPen(QtGui.QColor(255, 255, 255))
                painter.drawText(6, 18, "CC")
            painter.end()
        if file_key in self.thumbnail_filters:
            painter = QtGui.QPainter(pix)
            painter.setBrush(QtGui.QColor(160, 16, 239, 220))
            painter.setPen(QtGui.QPen(QtGui.QColor('black')))
            painter.drawEllipse(pix.width() - 24, 6, 18, 18)
            painter.setPen(QtGui.QColor('white'))
            painter.setFont(QtGui.QFont("Segoe UI", 9, QtGui.QFont.Bold))
            painter.drawText(QtCore.QRect(pix.width() - 24, 6, 18, 18), QtCore.Qt.AlignCenter, "F")
            painter.end()
        highlight_spec = None
        if getattr(self, '_highlighted_spec', None):
            highlight_key = str(self._highlighted_spec.get('image_key') or self._highlighted_spec.get('path') or '')
            shared_keys = [str(key) for key in (self._highlighted_spec.get("shared_image_keys") or []) if key]
            if highlight_key == str(file_key) or str(file_key) in shared_keys:
                highlight_spec = self._highlighted_spec
        if not getattr(self, "spectro_highlight_glow", True):
            highlight_spec = None
        if header and fds and 0 <= channel_idx < len(fds):
            try:
                xpix = int(header.get('xPixel', 128))
                ypix = int(header.get('yPixel', xpix))
                marker_defs = self._render_spectroscopy_overlays(
                    pix,
                    header,
                    str(file_key),
                    xpix,
                    ypix,
                    selected_spec=highlight_spec,
                    thumb_crop=thumb_crop,
                )
            except Exception:
                marker_defs = []
        return marker_defs

    def _schedule_thumbnail_job(self, file_key, channel_idx, header, fd, thumb_w, thumb_h, cmap_name, generation):
        if file_key in self._thumb_inflight or file_key in self._thumb_loaded:
            return
        self._thumb_inflight.add(file_key)
        job = _ThumbnailJob(self, file_key, channel_idx, header, fd, thumb_w, thumb_h, cmap_name, generation)
        job.signals.finished.connect(self._on_thumbnail_job_finished)
        job.signals.failed.connect(self._on_thumbnail_job_failed)
        self._thumb_threadpool.start(job)

    def _on_thumbnail_job_finished(self, file_key, channel_idx, qimg, data_key, cmap_name, generation):
        if generation != self._thumb_generation:
            return
        label = self._thumb_labels.get(file_key)
        if label is None or qimg is None:
            return
        dims = label.property("thumb_dims")
        if not dims:
            dims = self._thumb_dimensions()
        thumb_w, thumb_h = dims
        base_pix = QtGui.QPixmap.fromImage(qimg).scaled(thumb_w, thumb_h, QtCore.Qt.KeepAspectRatio, QtCore.Qt.FastTransformation)
        try:
            self.thumb_cache[(data_key, cmap_name)] = base_pix
        except Exception:
            pass
        crop_info = None
        try:
            with self._thumb_data_lock:
                crop_info = self._thumb_crop_cache.get(data_key)
        except Exception:
            crop_info = None
        pix = base_pix.copy()
        header, fds = self.headers.get(str(file_key), (None, None))
        markers = self._decorate_thumbnail_pixmap(pix, file_key, channel_idx, header, fds, thumb_crop=crop_info)
        label.setPixmap(pix)
        label.setProperty("spec_markers", markers)
        try:
            label.setProperty("thumb_crop", crop_info)
        except Exception:
            pass
        self._thumb_inflight.discard(file_key)
        self._thumb_loaded.add(file_key)
        try:
            self._request_visible_thumbs()
        except Exception:
            pass

    def _on_thumbnail_job_failed(self, file_key, channel_idx, error, generation):
        if generation != self._thumb_generation:
            return
        label = self._thumb_labels.get(file_key)
        if label is None:
            return
        dims = label.property("thumb_dims")
        if not dims:
            dims = self._thumb_dimensions()
        thumb_w, thumb_h = dims
        pix = QtGui.QPixmap(thumb_w, thumb_h)
        pix.fill(QtGui.QColor('black'))
        label.setPixmap(pix)
        label.setProperty("spec_markers", [])
        self._thumb_inflight.discard(file_key)
        try:
            log_status(f"Thumbnail failed for {file_key}: {error}")
        except Exception:
            pass
        try:
            self._request_visible_thumbs()
        except Exception:
            pass

    def _request_visible_thumbs(self):
        """Schedule thumbnail rendering for currently visible rows (+margin)."""
        if not getattr(self, 'current_thumb_files', None):
            return
        vp = getattr(self, '_thumb_viewport', None)
        scroll = getattr(self, 'scroll', None)
        cols = max(1, getattr(self, 'thumb_grid_columns', 1))
        card_h = getattr(self, '_thumb_card_height', None) or (self.thumb_size_px + 48)
        try:
            y0 = scroll.verticalScrollBar().value() if scroll else 0
            vh = vp.height() if vp else card_h * 4
        except Exception:
            y0 = 0; vh = card_h * 4
        first_row = max(0, int(y0 // card_h) - 2)
        last_row = int((y0 + vh) // card_h) + 2
        start_idx = max(0, first_row * cols)
        end_idx = min(len(self.current_thumb_files), (last_row + 1) * cols)
        visible_keys = self.current_thumb_files[start_idx:end_idx]
        for key in visible_keys:
            if key in self._thumb_loaded or key in self._thumb_inflight:
                continue
            meta = self._thumb_meta.get(key)
            if not meta:
                continue
            channel_idx, header, fd, thumb_w, thumb_h, cmap_name, gen = meta
            if gen != self._thumb_generation:
                continue
            self._schedule_thumbnail_job(key, channel_idx, header, fd, thumb_w, thumb_h, cmap_name, gen)

    def _get_thumbnail_array(self, file_key, channel_idx, header, fd, thumb_w, thumb_h):
        return viewer_thumbnails._get_thumbnail_array(self, file_key, channel_idx, header, fd, thumb_w, thumb_h)

    def _thumbnail_data_key(self, file_key, channel_idx, fd, thumb_w, thumb_h):
        return viewer_thumbnails._thumbnail_data_key(self, file_key, channel_idx, fd, thumb_w, thumb_h)

    def _invalidate_thumbnail_cache(self, paths=None):
        return viewer_thumbnails._invalidate_thumbnail_cache(self, paths=paths)

    def _is_processed_key(self, key: str):
        """Return True for in-memory processed/virtual entries (drift copies, etc.)."""
        try:
            if isinstance(key, str) and key.startswith("processed_"):
                return True
            return str(key) in getattr(self, "_processed_views", {})
        except Exception:
            return False

    def _make_processed_key(self, origin_path: str, op: str = "virtual", channel_idx=None):
        """Generate a unique processed key for a virtual copy based on origin and op."""
        stem = Path(origin_path).stem
        chan = f"_ch{channel_idx}" if channel_idx is not None else ""
        ctr = getattr(self, "_virtual_counter", 0) + 1
        self._virtual_counter = ctr
        return f"processed_{stem}_{op}{chan}_{ctr}"

    def _normalize_virtual_copy_order(self):
        ordered = []
        seen = set()
        for key in list(getattr(self, "virtual_copy_order", []) or []):
            skey = str(key)
            if not skey or skey in seen or skey not in getattr(self, "_processed_views", {}):
                continue
            ordered.append(skey)
            seen.add(skey)
        self.virtual_copy_order = ordered
        return ordered

    def _processed_insert_anchor(self, processed_key):
        data = getattr(self, "_processed_views", {}).get(str(processed_key)) or {}
        anchor = data.get("insert_after")
        if anchor in (None, ""):
            anchor = data.get("source") or VIRTUAL_COPY_INSERT_START
        return str(anchor) if anchor else VIRTUAL_COPY_INSERT_START

    def _ordered_virtual_thumbnail_files(self, real_files, processed_files=None):
        real_keys = [str(p) for p in list(real_files or []) if p and not self._is_processed_key(str(p))]
        candidate_set = {str(p) for p in list(processed_files or []) if p and self._is_processed_key(str(p))}
        ordered_processed = [key for key in self._normalize_virtual_copy_order() if key in candidate_set]
        for key in candidate_set:
            if key not in ordered_processed:
                ordered_processed.append(key)
        after_map = defaultdict(list)
        for key in ordered_processed:
            after_map[self._processed_insert_anchor(key)].append(key)
        result = []
        visited = set()

        def _append_children(anchor):
            for child in after_map.get(str(anchor), []):
                if child in visited:
                    continue
                visited.add(child)
                result.append(child)
                _append_children(child)

        _append_children(VIRTUAL_COPY_INSERT_START)
        for key in real_keys:
            result.append(key)
            _append_children(key)
        for key in ordered_processed:
            if key in visited:
                continue
            result.append(key)
            _append_children(key)
        return result

    def _thumbnail_image_display_order(self):
        current = [str(key) for key in list(getattr(self, "current_thumb_files", []) or []) if str(key)]
        if current:
            return current
        real_keys = [str(p) for p in list(getattr(self, "files", []) or []) if not self._is_processed_key(str(p))]
        processed_keys = [str(p) for p in list(getattr(self, "files", []) or []) if self._is_processed_key(str(p))]
        return self._ordered_virtual_thumbnail_files(real_keys, processed_keys)

    def _set_processed_insert_after(self, processed_key, after_key=None, display_order=None):
        key = str(processed_key)
        if key not in getattr(self, "_processed_views", {}):
            return
        anchor = after_key
        if anchor in (None, "", VIRTUAL_COPY_INSERT_START):
            anchor = VIRTUAL_COPY_INSERT_START
        else:
            anchor = str(anchor)
        self._processed_views[key]["insert_after"] = anchor
        order = [item for item in self._normalize_virtual_copy_order() if item != key]
        if display_order is None:
            display_order = self._thumbnail_image_display_order()
        display_order = [str(item) for item in list(display_order or []) if str(item) and str(item) != key]
        if anchor == VIRTUAL_COPY_INSERT_START:
            slot = 0
        elif anchor in display_order:
            slot = display_order.index(anchor) + 1
        else:
            slot = len(display_order)
        processed_before = 0
        for existing in display_order[:slot]:
            if self._is_processed_key(existing) and existing in order:
                processed_before += 1
        order.insert(min(processed_before, len(order)), key)
        self.virtual_copy_order = order

    def _insert_processed_after_source(self, processed_key: str, origin_path: str, insert_after_key=None):
        """Insert processed entry immediately after its origin or a supplied display anchor."""
        processed_key = str(processed_key)
        origin_path = str(origin_path)
        anchor_key = insert_after_key
        if anchor_key in (None, "", VIRTUAL_COPY_INSERT_START):
            anchor_key = origin_path
        else:
            anchor_key = str(anchor_key)
        try:
            cur_files = [str(p) for p in self.files]
            if processed_key in cur_files:
                idx = cur_files.index(processed_key)
                self.files.pop(idx)
                cur_files.pop(idx)
            try:
                pos = cur_files.index(anchor_key)
            except ValueError:
                try:
                    pos = cur_files.index(origin_path)
                except ValueError:
                    pos = len(self.files) - 1
            self.files.insert(pos + 1, Path(processed_key))
        except Exception:
            # fallback append
            self.files.append(Path(processed_key))
        self._set_processed_insert_after(processed_key, after_key=anchor_key)

    def _remove_virtual_entries(self, paths):
        """Remove selected virtual copies from in-memory store and UI."""
        if not paths:
            return
        keys = {str(Path(p)) for p in paths if self._is_processed_key(str(p))}
        if not keys:
            return
        for k in keys:
            self._processed_views.pop(k, None)
            self.headers.pop(k, None)
            self.molecule_overlays.pop(k, None)
        self.files = [p for p in self.files if str(p) not in keys]
        self._channel_data_cache = OrderedDict()
        self._invalidate_thumbnail_cache(keys)
        self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())

    def _channel_cache_key(self, file_key, channel_idx, fd):
        fname = fd.get('FileName')
        if not fname:
            raise ValueError("Missing FileName for channel")
        if self._is_processed_key(file_key):
            return (str(file_key), int(channel_idx), 0.0)
        bin_path = Path(file_key).parent / fname
        try:
            mtime = bin_path.stat().st_mtime
        except Exception:
            mtime = 0.0
        return (str(bin_path), int(channel_idx), mtime)

    def _get_channel_array(self, file_key, channel_idx, header, fd):
        key = self._channel_cache_key(file_key, channel_idx, fd)
        cache = self._channel_data_cache
        with self._channel_cache_lock:
            arr = cache.get(key)
            if arr is not None:
                cache.move_to_end(key)
                return arr
        xpix = int(header.get('xPixel', 128))
        ypix = int(header.get('yPixel', xpix))
        if self._is_processed_key(file_key):
            data = self._processed_views.get(str(file_key))
            if data is None:
                raise FileNotFoundError(f"Processed view not found: {file_key}")
            arr = None
            arr_by_channel = data.get('arr_by_channel') or {}
            if arr_by_channel:
                arr = arr_by_channel.get(channel_idx)
                if arr is None:
                    # fallback to any available channel
                    try:
                        arr = next(iter(arr_by_channel.values()))
                    except Exception:
                        arr = None
            if arr is None and 'arr' in data:
                arr = data.get('arr')
            if arr is None:
                raise FileNotFoundError(f"Processed data missing array for {file_key}")
            arr = np.asarray(arr)
        else:
            bin_path = Path(key[0])
            arr = read_channel_file(bin_path, xpix, ypix,
                                    scale=fd.get('Scale', 1.0), offset=fd.get('Offset', 0.0))
        with self._channel_cache_lock:
            cache[key] = arr
            while len(cache) > CHANNEL_DATA_CACHE_LIMIT:
                cache.popitem(last=False)
        return arr

    def _get_filtered_channel_array(self, file_key, channel_idx, header, fd):
        file_key = str(file_key)
        channel_key = self._channel_cache_key(file_key, channel_idx, fd)
        arr = self._get_channel_array(file_key, channel_idx, header, fd)
        unit = fd.get('PhysUnit','')
        unit_final, arr_conv = normalize_unit_and_data(arr, unit)
        spec = self.thumbnail_filters.get(file_key)
        sig = _filter_signature(spec)
        cache_key = (channel_key, unit_final, sig)
        with self._filtered_cache_lock:
            cached = self._filtered_channel_cache.get(cache_key)
            if cached is not None:
                self._filtered_channel_cache.move_to_end(cache_key)
                return unit_final, cached
        result = np.asarray(arr_conv, dtype=float)
        if sig:
            result = self._apply_filter_pipeline(result, spec.get('steps', []))
        with self._filtered_cache_lock:
            self._filtered_channel_cache[cache_key] = result
            while len(self._filtered_channel_cache) > FILTERED_CACHE_LIMIT:
                self._filtered_channel_cache.popitem(last=False)
        return unit_final, result

    def _invalidate_channel_cache(self, paths=None):
        with self._channel_cache_lock:
            if not paths:
                self._channel_data_cache.clear()
                with self._filtered_cache_lock:
                    self._filtered_channel_cache.clear()
                self._frame_real_pixmap_cache.clear()
                self._processed_views.clear()
                return
            parent_dirs = {str(Path(p).parent) for p in paths if not self._is_processed_key(str(p))}
            to_remove = []
            for k in list(self._channel_data_cache.keys()):
                if self._is_processed_key(k[0]) or str(Path(k[0]).parent) in parent_dirs:
                    to_remove.append(k)
            for k in to_remove:
                self._channel_data_cache.pop(k, None)
        self._invalidate_filtered_cache(paths)

    def _invalidate_filtered_cache(self, paths=None):
        with self._filtered_cache_lock:
            if not paths:
                self._filtered_channel_cache.clear()
                self._frame_real_pixmap_cache.clear()
                return
            parent_dirs = {str(Path(p).parent) for p in paths if not self._is_processed_key(str(p))}
            to_remove = [k for k in self._filtered_channel_cache.keys()
                        if self._is_processed_key(k[0][0]) or str(Path(k[0][0]).parent) in parent_dirs]
            for k in to_remove:
                self._filtered_channel_cache.pop(k, None)
        self._frame_real_pixmap_cache.clear()

    def on_thumb_sort_changed(self, idx):
        return viewer_thumb_ui.on_thumb_sort_changed(self, idx)

    def on_thumb_filter_changed(self, idx):
        return viewer_thumb_ui.on_thumb_filter_changed(self, idx)

    def on_unit_display_toggled(self, checked: bool):
        self.display_units_si = bool(checked)
        for widget in (
            getattr(self, "unit_display_cb", None),
            getattr(self, "display_units_si_act", None),
        ):
            if widget is None:
                continue
            try:
                widget.blockSignals(True)
                widget.setChecked(self.display_units_si)
                widget.blockSignals(False)
            except Exception:
                pass
        self.config['display_units_si'] = self.display_units_si
        save_config(self.config)
        if self.last_preview:
            self.show_file_channel(self.last_preview[0], self.last_preview[1])

    def on_unit_relative_toggled(self, checked: bool):
        self.display_units_relative = bool(checked)
        for widget in (
            getattr(self, "unit_relative_cb", None),
            getattr(self, "display_units_relative_act", None),
        ):
            if widget is None:
                continue
            try:
                widget.blockSignals(True)
                widget.setChecked(self.display_units_relative)
                widget.blockSignals(False)
            except Exception:
                pass
        self.config['display_units_relative'] = self.display_units_relative
        save_config(self.config)
        if self.last_preview:
            self.show_file_channel(self.last_preview[0], self.last_preview[1])

    def on_relative_axes_toggled(self, checked: bool):
        self.relative_axes = bool(checked)
        for widget in (
            getattr(self, "relative_axes_cb", None),
            getattr(self, "relative_axes_act", None),
        ):
            if widget is None:
                continue
            try:
                widget.blockSignals(True)
                widget.setChecked(self.relative_axes)
                widget.blockSignals(False)
            except Exception:
                pass
        self.config['relative_axes'] = self.relative_axes
        save_config(self.config)
        # Prevent restoring stale profile state when switching axes mode
        self._suppress_profile_restore = True
        # Clear any active profiles that were built with the previous axis mode
        try:
            if getattr(self, 'preview_canvas', None):
                self.preview_canvas.enable_profile(False)
                if hasattr(self.preview_canvas, "_clear_saved_profile_artists"):
                    self.preview_canvas._clear_saved_profile_artists(notify=False)
                self.preview_canvas.profile_pts = None
        except Exception:
            pass
        # Clear cached zoom limits so switching relative axes always recomputes extents
        try:
            if hasattr(self.preview_canvas, "_zoom_reset_limits"):
                self.preview_canvas._zoom_reset_limits = {}
                self.preview_canvas._reset_view_zoom()
        except Exception:
            pass
        # Invalidate cached frames so extents are recalculated
        try:
            self._frame_real_pixmap_cache.clear()
            self._filtered_channel_cache.clear()
        except Exception:
            pass
        try:
            if getattr(self, 'preview_canvas', None):
                self.preview_canvas.set_relative_axes_override(self.relative_axes)
        except Exception:
            pass
        if self.last_preview:
            try:
                key, idx = self.last_preview
            except Exception:
                key = idx = None
            self.last_preview = None  # force rebuild of current view
            if key is not None and idx is not None:
                try:
                    if getattr(self, 'preview_canvas', None):
                        self.preview_canvas.suspend_zoom_restore()
                except Exception:
                    pass
                self.show_file_channel(key, idx)
                # After the view is rebuilt, reset zoom so the new extent
                # (relative vs absolute) is applied immediately.
                try:
                    if getattr(self, 'preview_canvas', None):
                        self.preview_canvas._reset_view_zoom()
                except Exception:
                    pass

    def on_scale_bar_toggled(self, checked: bool):
        for widget in (
            getattr(self, "scale_bar_cb", None),
            getattr(self, "display_scale_bar_act", None),
        ):
            if widget is None:
                continue
            try:
                widget.blockSignals(True)
                widget.setChecked(bool(checked))
                widget.blockSignals(False)
            except Exception:
                pass
        options = self._canvas_display_state_from_canvas(getattr(self, "preview_canvas", None))
        options["scale_bar_enabled"] = bool(checked)
        self._apply_canvas_display_options(options, source_canvas=getattr(self, "preview_canvas", None), persist=True)
        if self.last_preview:
            self.show_file_channel(self.last_preview[0], self.last_preview[1])

    # removed size change handler

    def _parse_header_datetime(self, header, path=None):
        return viewer_loader._parse_header_datetime(self, header, path=path)

    def _header_datetime_dt(self, header, path):
        try:
            ts = float(self._parse_header_datetime(header or {}, path=path))
            if ts <= 0:
                ts = Path(path).stat().st_mtime
            return datetime.fromtimestamp(ts)
        except Exception:
            return datetime.fromtimestamp(Path(path).stat().st_mtime)

    def _build_image_timestamp_index(self):
        self.image_time_index = {}
        self.image_meta = []
        for p in self.files:
            header, _ = self.headers.get(str(p), (None, None))
            if header is None:
                continue
            dt = self._header_datetime_dt(header, p)
            self.image_time_index[str(p)] = dt
            self.image_meta.append({'path': Path(p), 'time': dt})

    def _build_metadata_html(self, header_path:Path, header:dict, fd:dict, channel_idx:int, unit_normalized:str, unit_display:str, arr_display:np.ndarray, zero_offset:float|None) -> str:
        return viewer_preview._build_metadata_html(self, header_path, header, fd, channel_idx, unit_normalized, unit_display, arr_display, zero_offset)

    def _build_single_channel_view(self, header_path_str, channel_idx: int, *, cmap_override=None, use_local_cmap=False):
        return viewer_preview.build_single_channel_view(
            self,
            header_path_str,
            channel_idx,
            cmap_override=cmap_override,
            use_local_cmap=use_local_cmap,
        )

    def _frame_entry_from_header(self, path, header):
        if header is None:
            return None

        def as_nm(key, unit_key):
            val = _safe_float(header.get(key))
            unit = header.get(unit_key, header.get('PhysUnit', 'nm'))
            return _value_in_nm(val, unit)

        x_range_nm = as_nm('XScanRange', 'XPhysUnit')
        y_range_nm = as_nm('YScanRange', 'YPhysUnit')
        cx_nm = as_nm('xCenter', 'XPhysUnit')
        cy_nm = as_nm('yCenter', 'YPhysUnit')
        if None in (x_range_nm, y_range_nm, cx_nm, cy_nm):
            return None
        angle = _safe_float(header.get('Angle')) or 0.0
        clamp = lambda v: max(-1000.0, min(1000.0, v))
        return {
            'key': str(path),
            'cx_nm': clamp(cx_nm),
            'cy_nm': clamp(cy_nm),
            'x_range_nm': max(5.0, min(2000.0, abs(x_range_nm))),
            'y_range_nm': max(5.0, min(2000.0, abs(y_range_nm))),
            'angle_deg': float(angle),
            'tag': (self.tags.get(str(path), {}) or {}).get('tag')
        }

    def _rebuild_frame_map_entries(self):
        entries = []
        for p in self.files:
            header, _ = self.headers.get(str(p), (None, None))
            entry = self._frame_entry_from_header(p, header)
            if entry:
                entries.append(entry)
        self.frame_map_entries = entries
        if hasattr(self, 'frame_map_widget'):
            self.frame_map_widget.set_entries(entries)
            self.frame_map_widget.set_hidden_entries(self.hidden_frame_keys)
            self._refresh_frame_map_pixmaps()

    def _on_frame_map_entry_shift_clicked(self, key):
        if not key:
            return
        self.hidden_frame_keys.add(str(key))
        if getattr(self, 'selected_file_for_thumbs', None) == str(key):
            self.selected_file_for_thumbs = None
        if hasattr(self, 'frame_map_widget'):
            self.frame_map_widget.set_hidden_entries(self.hidden_frame_keys)

    def _on_frame_show_all_clicked(self):
        if not self.hidden_frame_keys:
            return
        self.hidden_frame_keys.clear()
        if hasattr(self, 'frame_map_widget'):
            self.frame_map_widget.clear_hidden_entries()

    def _on_frame_real_view_toggled(self, checked):
        self.frame_real_view = bool(checked)
        if hasattr(self, 'frame_real_view_btn'):
            self.frame_real_view_btn.setText("Hide real view" if checked else "Show real view")
        if hasattr(self, 'frame_map_widget'):
            self.frame_map_widget.set_real_view_enabled(self.frame_real_view)
        self._refresh_frame_map_pixmaps()

    def _refresh_frame_map_pixmaps(self):
        if not getattr(self, 'frame_map_widget', None):
            return
        if not self.frame_real_view:
            self.frame_entry_pixmaps = {}
            self.frame_map_widget.set_entry_pixmaps({})
            return
        channel_idx = self.channel_dropdown.currentIndex() if self.channel_dropdown.count() else 0
        cmap = self.thumb_cmap_combo.currentText() or self.thumb_cmap
        pixmaps = {}
        thumb_w, thumb_h = 96, 72
        for entry in self.frame_map_entries:
            key = entry.get('key')
            pix = self._thumbnail_pixmap_for_file(key, channel_idx, thumb_w, thumb_h, cmap)
            if pix is not None:
                pixmaps[key] = pix
        self.frame_entry_pixmaps = pixmaps
        self.frame_map_widget.set_entry_pixmaps(pixmaps)

    def _slider_value_to_zoom(self, slider_val: int) -> float:
        exp = (float(slider_val) - float(self.FRAME_ZOOM_SLIDER_DEFAULT)) / 100.0
        zoom = 10.0 ** exp
        return float(np.clip(zoom, 0.01, 10000.0))

    def _zoom_to_slider_value(self, zoom: float) -> int:
        zoom = max(0.01, min(zoom, 10000.0))
        return int(round(100.0 * math.log10(zoom) + self.FRAME_ZOOM_SLIDER_DEFAULT))

    def _normalize_frame_zoom_slider_value(self, stored: int) -> int:
        if stored < self.FRAME_ZOOM_SLIDER_MIN:
            return self.FRAME_ZOOM_SLIDER_MIN
        if stored > self.FRAME_ZOOM_SLIDER_MAX:
            # legacy linear scaling stored zoom * 100
            legacy_zoom = max(0.01, stored / 100.0)
            return self._zoom_to_slider_value(legacy_zoom)
        return stored

    def _thumbnail_pixmap_for_file(self, file_key, channel_idx, width, height, cmap_name):
        return viewer_thumb_ui._thumbnail_pixmap_for_file(self, file_key, channel_idx, width, height, cmap_name)

    def _update_frame_map_active(self, key):
        if hasattr(self, 'frame_map_widget'):
            self.frame_map_widget.set_active_key(key)

    def _on_frame_map_clicked(self, key):
        if not key:
            return
        header, _ = self.headers.get(str(key), (None, None))
        if header is None:
            return
        self.selected_file_for_thumbs = str(key)
        self._refresh_thumb_selection_styles()
        channel_idx = self.channel_dropdown.currentIndex()
        try:
            self.show_file_channel(str(key), channel_idx)
        except Exception:
            pass

    def _apply_frame_zoom_slider(self):
        if hasattr(self, 'frame_map_widget') and hasattr(self, 'frame_zoom_slider'):
            factor = self._slider_value_to_zoom(self.frame_zoom_slider.value())
            self.frame_map_widget.set_zoom_factor(factor)

    def _on_frame_map_zoom_changed(self, factor):
        if not hasattr(self, 'frame_zoom_slider'):
            return
        val = self._zoom_to_slider_value(factor)
        if self.frame_zoom_slider.value() == val:
            return
        self.frame_zoom_slider.blockSignals(True)
        self.frame_zoom_slider.setValue(val)
        self.frame_zoom_slider.blockSignals(False)
        self.config['frame_map_zoom'] = val
        save_config(self.config)

    def _reset_frame_view(self):
        if not hasattr(self, 'frame_map_widget') or not hasattr(self, 'frame_zoom_slider'):
            return
        self.frame_zoom_slider.setValue(self.FRAME_ZOOM_SLIDER_DEFAULT)
        self._apply_frame_zoom_slider()
        self.frame_map_widget.reset_pan()

    def _on_frame_zoom_changed(self, value):
        self.config['frame_map_zoom'] = value
        save_config(self.config)
        self._apply_frame_zoom_slider()

    def _refresh_thumb_selection_styles(self):
        return viewer_thumb_ui._refresh_thumb_selection_styles(self)

    def _refresh_spectro_thumb_selection_styles(self):
        return viewer_thumb_ui._refresh_spectro_thumb_selection_styles(self)

    def _clear_spectro_thumb_multi_selection(self, update_styles=True):
        return viewer_thumb_ui._clear_spectro_thumb_multi_selection(self, update_styles=update_styles)

    def _schedule_marker_refresh(self, delay_ms: int = 120):
        try:
            self._marker_refresh_timer.start(max(0, int(delay_ms)))
        except Exception:
            self._schedule_marker_refresh()

    def _refresh_thumbnail_markers(self):
        labels = getattr(self, '_thumb_labels', {}) or {}
        if not labels:
            return
        try:
            cmap_name = self.thumb_cmap_combo.currentText()
        except Exception:
            cmap_name = None
        if not cmap_name:
            cmap_name = getattr(self, 'thumb_cmap', 'viridis')
        for file_key, label in labels.items():
            if label is None:
                continue
            try:
                thumb_dims = label.property("thumb_dims") or (0, 0)
                channel_idx = int(label.property("channel_index") or 0)
            except Exception:
                continue
            if not thumb_dims or thumb_dims[0] <= 0 or thumb_dims[1] <= 0:
                continue
            base_pix = viewer_thumb_ui._thumbnail_pixmap_for_file(
                self, file_key, channel_idx, thumb_dims[0], thumb_dims[1], cmap_name
            )
            if base_pix is None:
                continue
            pix = base_pix.copy()
            header, fds = self.headers.get(str(file_key), (None, None))
            crop_info = None
            try:
                if fds and 0 <= channel_idx < len(fds):
                    fd = fds[channel_idx]
                    data_key = self._thumbnail_data_key(file_key, channel_idx, fd, thumb_dims[0], thumb_dims[1])
                    with self._thumb_data_lock:
                        crop_info = self._thumb_crop_cache.get(data_key)
            except Exception:
                crop_info = None
            try:
                markers = self._decorate_thumbnail_pixmap(pix, file_key, channel_idx, header, fds, thumb_crop=crop_info)
            except Exception:
                markers = []
            label.setPixmap(pix)
            label.setProperty("spec_markers", markers)
            try:
                label.setProperty("thumb_crop", crop_info)
            except Exception:
                pass

    def _make_thumb_press_handler(self, label_widget):
        return viewer_thumb_ui._make_thumb_press_handler(self, label_widget)

    def _make_thumb_release_handler(self, label_widget):
        return viewer_thumb_ui._make_thumb_release_handler(self, label_widget)

    def _make_thumb_move_handler(self, label_widget):
        return viewer_thumb_ui._make_thumb_move_handler(self, label_widget)

    def _make_thumb_double_handler(self, label_widget):
        return viewer_thumb_ui._make_thumb_double_handler(self, label_widget)

    def _canvas_window_ref(self):
        win = getattr(self, "_canvas_window", None)
        if win is None:
            return None
        try:
            if sip.isdeleted(win):
                self._canvas_window = None
                return None
        except Exception:
            self._canvas_window = None
            return None
        return win

    def _on_open_canvas(self):
        win = self._canvas_window_ref()
        if win is None or not win.isVisible():
            win = ExperimentalCanvasWindow(self, self)
            self._canvas_window = win
            win.finished.connect(lambda _=None, w=win: self._remember_closed_canvas_window(w))
        win.show()
        win.raise_()
        try:
            win.activateWindow()
        except Exception:
            pass
        # Automatically push the current thumbnail selection (or current preview) into the canvas.
        try:
            targets = list(getattr(self, "thumb_multi_select", []) or [])
            if not targets:
                if getattr(self, "selected_file_for_thumbs", None):
                    targets = [self.selected_file_for_thumbs]
                elif self.last_preview:
                    targets = [self.last_preview[0]]
            payloads = []
            for fp in sorted(set(targets)):
                # When dropping into the canvas we want all selected files, not just the drag origin.
                payloads.append({"file_path": fp, "channel_index": None, "cmap": None})
            if payloads:
                try:
                    win.handle_drop(payloads, [])
                except Exception:
                    pass
        except Exception:
            pass

    def _ensure_canvas_for_drag(self):
        """Open the canvas window as a drop target during thumbnail drags."""
        win = self._canvas_window_ref()
        if win is None or not win.isVisible():
            win = ExperimentalCanvasWindow(self, self)
            self._canvas_window = win
            win.finished.connect(lambda _=None, w=win: self._remember_closed_canvas_window(w))
        win.show()
        win.raise_()
        try:
            win.activateWindow()
        except Exception:
            pass

    def on_thumbnail_clicked(self, header_path_str, channel_idx):
        controller = getattr(self, "thumbnail_controller", None)
        if controller:
            return controller.handle_thumbnail_clicked(header_path_str, channel_idx)

    def on_thumbnail_double_clicked(self, header_path_str, channel_idx):
        controller = getattr(self, "thumbnail_controller", None)
        if controller:
            return controller.handle_thumbnail_double_clicked(header_path_str, channel_idx)

    # NOTE: removed on_file_channel_selected and on_file_channel_show_clicked
    # These functions supported the removed per-file inspector UI. The same "show channel"
    # functionality is available via the thumbnail UI and the "Add channel view" dialog.

    # ---------- preview + metadata ---------- 
    def show_file_channel(self, header_path_str, channel_idx:int, use_local_cmap=False):
        highlight = getattr(self, '_highlighted_spec', None)
        if highlight:
            try:
                highlight_path = str(highlight.get('image_key') or highlight.get('path') or '')
            except Exception:
                highlight_path = ''
            if highlight_path and highlight_path != str(header_path_str):
                self._highlight_spectrum_entry(None)
        try:
            prev_key = str(self.last_preview[0]) if self.last_preview else None
        except Exception:
            prev_key = None
        try:
            new_key = str(header_path_str) if header_path_str is not None else None
        except Exception:
            new_key = None
        if new_key and new_key != prev_key:
            self._store_molecule_overlay(prev_key)
            self._load_molecule_overlay(new_key)
        result = viewer_preview.show_file_channel(self, header_path_str, channel_idx, use_local_cmap=use_local_cmap)
        try:
            if new_key:
                self.collection_controller.apply_snapshot_for_file(new_key)
        except Exception:
            pass
        return result

    def _store_molecule_overlay(self, file_key=None):
        """Persist current molecule overlays for a specific file key."""
        if not file_key:
            return
        canvas = getattr(self, "preview_canvas", None)
        if canvas is None:
            return
        try:
            state = canvas.export_molecule_state()
        except Exception:
            return
        if state is None:
            return
        self.molecule_overlays[str(file_key)] = state

    def _load_molecule_overlay(self, file_key=None):
        """Load molecule overlays for a specific file key (defaults to empty)."""
        if not file_key:
            return
        canvas = getattr(self, "preview_canvas", None)
        if canvas is None:
            return
        key = str(file_key)
        state = self.molecule_overlays.get(key)
        if state is None:
            state = []
        try:
            canvas.import_molecule_state(state)
        except Exception:
            pass

    def _clear_molecules_for_paths(self, paths):
        if not paths:
            return
        keys = {str(Path(p)) for p in paths}
        for key in keys:
            self.molecule_overlays[key] = []
        try:
            if self.last_preview and str(self.last_preview[0]) in keys:
                self._load_molecule_overlay(self.last_preview[0])
        except Exception:
            pass

    def _copy_molecules_from_source(self, paths):
        if not paths:
            return
        try:
            if self.last_preview:
                self._store_molecule_overlay(self.last_preview[0])
        except Exception:
            pass
        keys = {str(Path(p)) for p in paths if self._is_processed_key(str(p))}
        if not keys:
            return
        for key in keys:
            try:
                src = self._processed_views.get(str(key), {}).get("source")
                if not src:
                    continue
                src_key = str(src)
                state = self.molecule_overlays.get(src_key)
                if state is None:
                    continue
                self.molecule_overlays[str(key)] = state
            except Exception:
                continue
        try:
            if self.last_preview and str(self.last_preview[0]) in keys:
                self._load_molecule_overlay(self.last_preview[0])
        except Exception:
            pass

    def get_current_detail_config(self):
        """Return JSON-friendly configuration describing current detail view state."""
        cfg = {'channels': [], 'cmaps': {}, 'vmin_vmax': {}, 'figure_size': list(self.preview_canvas.fig.get_size_inches())}
        main_desc = None
        if self.last_preview:
            file_key = str(self.last_preview[0])
            header, fds = self.headers.get(file_key, (None, None))
            if header and fds:
                idx = int(self.last_preview[1])
                if 0 <= idx < len(fds):
                    cap = fds[idx].get('Caption', fds[idx].get('FileName', f"chan{idx}"))
                    key = f"idx_{idx}_{cap}"
                    main_desc = {'type': 'index', 'index': idx, 'caption': cap, 'key': key}
                    cfg['channels'].append(main_desc)
                    cmap = self.per_file_channel_cmap.get((file_key, idx), self.preview_cmap_combo.currentText() or self.preview_cmap)
                    cfg['cmaps'][key] = cmap
                    cfg['vmin_vmax'][key] = None
        # include extra views
        for spec in getattr(self, 'extra_view_specs', []):
            key = f"spec_{spec.get('caption','')}#{spec.get('index',-1)}"
            desc = {'type': 'spec', 'spec': spec.copy(), 'key': key}
            cfg['channels'].append(desc)
            cfg['cmaps'][key] = spec.get('cmap', self.preview_cmap_combo.currentText() or self.preview_cmap)
            cfg['vmin_vmax'][key] = None
        return cfg

    def _apply_filters_to_array(self, file_path, arr):
        spec = self.thumbnail_filters.get(str(file_path))
        if not spec:
            return arr
        return self._apply_filter_pipeline(arr, spec.get('steps', []))

    def on_save_session(self):
        """Legacy hook delegating to SessionController for compatibility."""
        self.session_controller.save_session()

    def on_load_session(self):
        """Legacy hook delegating to SessionController for compatibility."""
        self.session_controller.load_session()

    def on_open_collection(self):
        """Open a curated cross-folder collection workspace."""
        self.collection_controller.load_collection()
        self.show_collection_tray(activate=False)

    def on_collection_help(self):
        """Explain linked vs portable collections and how they are intended to be used."""
        self.collection_controller.show_help()

    def on_choose_current_collection(self):
        self.collection_controller.choose_current_collection()
        self.show_collection_tray(activate=False)

    def on_clear_current_collection(self):
        self.collection_controller.clear_current_collection()
        try:
            window = getattr(self, "collection_tray_window", None)
            if window is not None:
                window.hide()
        except Exception:
            pass

    def on_add_current_preview_to_collection(self):
        self.collection_controller.add_current_preview()
        self.show_collection_tray(activate=False)

    def on_add_active_popup_to_collection(self):
        self.collection_controller.add_active_popup()
        self.show_collection_tray(activate=False)

    def on_add_all_popups_to_collection(self):
        self.collection_controller.add_all_popups()
        self.show_collection_tray(activate=False)

    def on_add_selected_crops_to_collection(self):
        self.collection_controller.add_selected_crop_history()
        self.show_collection_tray(activate=False)

    def on_show_collection_tray(self):
        self.show_collection_tray(activate=True)

    def on_add_selected_thumbnails_to_collection(self):
        targets = list(self._ordered_thumbnail_selection() or [])
        if not targets:
            QtWidgets.QMessageBox.information(
                self,
                "Collections",
                "Select one or more thumbnails first, then add or drag them into the collection tray.",
            )
            return
        channel_idx = int(self.channel_dropdown.currentIndex() or 0)
        entries = [{"file_path": str(path), "channel_index": channel_idx} for path in targets if path]
        self.collection_controller.add_thumbnail_entries(entries)
        self._refresh_collection_tray()
        self.show_collection_tray(activate=False)

    def show_collection_tray(self, activate=True):
        window = getattr(self, "collection_tray_window", None)
        if window is None:
            return
        try:
            window.show()
            if activate:
                window.raise_()
                window.activateWindow()
        except Exception:
            pass

    def _refresh_collection_ui(self):
        current = str(getattr(self, "_collection_source", "") or "").strip()
        short = current if current else "none"
        if len(short) > 96:
            short = "..." + short[-93:]
        text = f"Current collection: {short}"
        tooltip = current or "No current collection selected. Add actions will ask for a collection file."
        button_text = "Collections"
        if current:
            try:
                button_text = f"Collection: {Path(current).stem}"
            except Exception:
                button_text = "Collection: active"
        for attr in ("collection_current_path_act", "toolbar_collection_current_path_act"):
            act = getattr(self, attr, None)
            if act is None:
                continue
            try:
                act.setText(text)
                act.setToolTip(tooltip)
            except Exception:
                pass
        for attr in ("collection_clear_target_act", "toolbar_collection_clear_target_act"):
            act = getattr(self, attr, None)
            if act is None:
                continue
            try:
                act.setEnabled(bool(current))
            except Exception:
                pass
        btn = getattr(self, "toolbar_collection_btn", None)
        if btn is not None:
            try:
                btn.setText(button_text)
                btn.setToolTip(tooltip)
            except Exception:
                pass
        try:
            self._refresh_collection_tray()
        except Exception:
            pass

    def _refresh_collection_tray(self):
        current = str(getattr(self, "_collection_source", "") or "").strip()
        group = getattr(self, "collection_group", None)
        tray = getattr(self, "collection_tray_list", None)
        window = getattr(self, "collection_tray_window", None)
        if group is None or tray is None:
            return
        tray.clear()
        if window is not None:
            try:
                window.setWindowTitle("Collection Tray" if not current else f"Collection Tray - {Path(current).stem}")
            except Exception:
                pass
        if not current:
            group.setTitle("Current Collection")
            target_label = getattr(self, "collection_target_label", None)
            if target_label is not None:
                target_label.setText("No collection selected yet.")
            item = QListWidgetItem("Choose or open a collection, then drag thumbnails or preview views here.")
            item.setFlags(QtCore.Qt.NoItemFlags)
            tray.addItem(item)
            return
        current_path = Path(current)
        target_label = getattr(self, "collection_target_label", None)
        if target_label is not None:
            if current_path.exists():
                target_label.setText(f"{current_path.name}\n{current_path}")
            else:
                target_label.setText(f"{current_path.name}\n{current_path}\n(New collection target)")
        entries = self.collection_controller.tray_entries_for_current_collection(icon_size=72)
        group.setTitle(f"Current Collection ({len(entries)})")
        if not entries:
            item = QListWidgetItem("Collection is ready.\nDrag thumbnails or preview views here to start adding items.")
            item.setFlags(QtCore.Qt.NoItemFlags)
            tray.addItem(item)
            return
        for entry in entries:
            item = QListWidgetItem(entry.get("icon") or QtGui.QIcon(), entry.get("text") or entry.get("label") or "Collection item")
            item.setToolTip(entry.get("tool_tip") or "")
            item.setData(QtCore.Qt.UserRole, entry)
            item.setSizeHint(QtCore.QSize(220, 82))
            tray.addItem(item)

    def on_recall_popouts(self):
        """Bring tracked pop-out dialogs back to the foreground, or restore saved ones if none are open."""
        controller = getattr(self, "quick_crop_controller", None)
        if controller and hasattr(controller, "raise_popups"):
            try:
                if controller.raise_popups():
                    controller.update_popup_actions()
                    return
            except Exception:
                pass
        if getattr(self, "_deferred_popup_entries", None):
            try:
                self.session_controller.restore_all_deferred_popups()
            except Exception:
                pass

    def on_arrange_popouts(self):
        """Tile all visible pop-out dialogs (preview, spectroscopy, profiles, etc.)."""
        controller = getattr(self, "quick_crop_controller", None)
        if controller:
            controller.arrange_popups()

    def on_minimize_popouts(self):
        """Minimize all visible pop-out dialogs (preview, spectroscopy, profiles, etc.)."""
        controller = getattr(self, "quick_crop_controller", None)
        if controller:
            controller.minimize_popups()

    def on_restore_popouts(self):
        """Restore minimized pop-out dialogs and bring them back to the foreground."""
        self.on_recall_popouts()

    def on_close_popouts(self):
        """Close all tracked pop-out dialogs without clearing crop history."""
        self._close_workspace_windows(record_history=False, include_canvas=False)
        controller = getattr(self, "quick_crop_controller", None)
        if controller:
            controller.update_popup_actions()

    def compare_menu_state(self):
        """Expose transient A/B compare-slot state for preview context menus."""
        controller = getattr(self, "image_compare_controller", None)
        if controller:
            return controller.menu_state()
        return {}

    def on_compare_menu_action(self, action, view, canvas=None):
        """Route preview compare-menu actions through the shared image compare controller."""
        controller = getattr(self, "image_compare_controller", None)
        if controller:
            controller.handle_menu_action(action, view, canvas=canvas)

    def _view_source_path(self, view):
        if not view:
            return None
        path = view.get("path")
        meta = view.get("meta") or {}
        if not path:
            path = meta.get("path") or meta.get("file_path")
        return path

    def _is_crop_view(self, view):
        if not view:
            return False
        title = str(view.get("title") or "").lower()
        label = str(view.get("label") or "").lower()
        if "[crop]" in title or label == "[crop]":
            return True
        return False

    def _filter_action_label(self, filter_key):
        base_label = FILTER_DEFINITIONS.get(filter_key, {}).get("label", str(filter_key or "").title())
        return f"{base_label}..."

    def _clone_filter_source_views(self, canvas, views):
        clone_view = getattr(canvas, "_clone_undo_view", None)
        cloned = []
        for view in list(views or []):
            try:
                if callable(clone_view):
                    cloned.append(clone_view(view))
                else:
                    cloned.append(copy.deepcopy(view))
            except Exception:
                cloned.append(view)
        return cloned

    def _normalize_preview_filter_steps(self, steps):
        if steps is None:
            return []
        if isinstance(steps, dict):
            return [steps]
        return [step for step in list(steps or []) if isinstance(step, dict)]

    def _set_filter_pipeline_on_canvas(self, canvas, steps, label=None, source_views=None, push_undo=False):
        if canvas is None:
            return
        base_views = source_views if source_views is not None else getattr(canvas, "views", None)
        if not base_views:
            return
        steps = self._normalize_preview_filter_steps(steps)
        if push_undo:
            try:
                canvas.push_undo_state("filter")
            except Exception:
                pass
        new_views = []
        for view in self._clone_filter_source_views(canvas, base_views):
            nv = dict(view)
            base = nv.get("_filter_base_arr")
            if base is None:
                try:
                    base = np.array(nv.get("arr"), copy=True)
                except Exception:
                    base = nv.get("arr")
                nv["_filter_base_arr"] = base
            if not steps:
                nv["arr"] = np.array(base, copy=True) if base is not None else nv.get("arr")
                nv.pop("filter_steps", None)
                nv.pop("filter_label", None)
            else:
                nv["arr"] = self._apply_filter_pipeline(base, steps) if base is not None else nv.get("arr")
                nv["filter_steps"] = copy.deepcopy(steps)
                nv["filter_label"] = label
                try:
                    clim = self._auto_preview_clim(
                        nv["arr"],
                        relative_zero=bool(nv.get("display_relative_zero", False)),
                    )
                except Exception:
                    clim = None
                if clim is not None:
                    nv["clim"] = clim
                else:
                    nv.pop("clim", None)
            new_views.append(nv)
        canvas.set_views(new_views, preserve_profiles=True)

    def _build_canvas_filter_preview_callback(self, canvas, source_views):
        def _preview(steps, label=None):
            if steps is None:
                self._restore_filter_views_on_canvas(canvas, source_views)
                return
            self._set_filter_pipeline_on_canvas(
                canvas,
                steps,
                label=label,
                source_views=source_views,
                push_undo=False,
            )
        return _preview

    def _restore_filter_views_on_canvas(self, canvas, source_views):
        if canvas is None or not source_views:
            return
        canvas.set_views(self._clone_filter_source_views(canvas, source_views), preserve_profiles=True)

    def _base_filter_image_from_views(self, views):
        try:
            if views:
                return views[0].get("_filter_base_arr") or views[0].get("arr")
        except Exception:
            return None
        return None

    def _load_filter_base_array_for_path(self, focus_path):
        base_arr = None
        if not focus_path:
            return None
        try:
            focus_key = str(focus_path)
            header, fds = self.headers.get(focus_key, (None, None))
            if header and fds:
                idx = None
                if self.last_preview and str(self.last_preview[0]) == focus_key:
                    idx = int(self.last_preview[1])
                if idx is None:
                    idx = 0
                if 0 <= idx < len(fds):
                    fd = fds[idx]
                    arr = self._get_channel_array(focus_key, idx, header, fd)
                    base_arr = normalize_unit_and_data(arr, fd.get("PhysUnit", ""))[1]
        except Exception:
            base_arr = None
        return base_arr

    def _normalize_filter_preview_clim(self, clim):
        try:
            if clim is None:
                return None
            lo, hi = clim
            lo = float(lo)
            hi = float(hi)
            if not np.isfinite(lo) or not np.isfinite(hi) or hi <= lo:
                return None
            return (lo, hi)
        except Exception:
            return None

    def _filter_preview_render_state(self, view=None):
        cmap_name = None
        clim = None
        if isinstance(view, dict):
            cmap_name = str(view.get("cmap") or "").strip() or None
            clim = self._normalize_filter_preview_clim(view.get("clim"))
        if not cmap_name:
            try:
                cmap_name = str(self.preview_cmap_combo.currentText() or "").strip() or None
            except Exception:
                cmap_name = None
        if not cmap_name:
            cmap_name = str(getattr(self, "preview_cmap", "viridis") or "viridis")
        return cmap_name, clim

    def _filter_preview_context_for_path(self, focus_path):
        preview_target = "selected image"
        preview_callback = None
        original_views = None
        preview_cmap_name = None
        preview_clim = None
        base_arr = self._load_filter_base_array_for_path(focus_path)
        try:
            preview_target = Path(str(focus_path)).name if focus_path else preview_target
        except Exception:
            preview_target = str(focus_path or preview_target)
        canvas = getattr(self, "preview_canvas", None)
        if (
            focus_path
            and canvas is not None
            and getattr(canvas, "views", None)
            and self.last_preview
            and str(self.last_preview[0]) == str(focus_path)
        ):
            original_views = self._clone_filter_source_views(canvas, canvas.views)
            base_arr = self._base_filter_image_from_views(original_views)
            preview_callback = self._build_canvas_filter_preview_callback(canvas, original_views)
            preview_target = self._friendly_view_title(original_views[0] if original_views else None, preview_target)
            preview_cmap_name, preview_clim = self._filter_preview_render_state(original_views[0] if original_views else None)
        if preview_cmap_name is None:
            preview_cmap_name, preview_clim = self._filter_preview_render_state(None)
        return base_arr, preview_callback, original_views, preview_target, preview_cmap_name, preview_clim

    def _single_filter_step_spec(
        self,
        filter_key,
        parent=None,
        base_image=None,
        preview_callback=None,
        preview_target_text="current image",
        preview_cmap_name="viridis",
        preview_clim=None,
        show_preview_thumbnail=True,
    ):
        if not filter_key:
            return None, None
        defaults = FILTER_DEFINITIONS.get(filter_key, {})
        if filter_key in ("highpass", "lowpass"):
            initial_params = self.config.get(f"{filter_key}_filter_params", {})
        elif filter_key == "laplacian":
            initial_params = self.config.get("laplacian_filter_params", {})
        else:
            initial_params = defaults
        dlg = SingleFilterDialog(
            parent=parent or self,
            filter_key=filter_key,
            base_image=base_image,
            apply_step_func=self._run_filter_step,
            preview_callback=preview_callback,
            initial_params=initial_params,
            preview_target_text=preview_target_text,
            preview_cmap_name=preview_cmap_name,
            preview_clim=preview_clim,
            show_preview_thumbnail=show_preview_thumbnail,
        )
        if dlg.exec_() != QtWidgets.QDialog.Accepted:
            return None, None
        step = dlg.current_step()
        label = dlg.current_step_label()
        params = dict(step.get("params") or {})
        if filter_key in ("highpass", "lowpass"):
            self.config[f"{filter_key}_filter_params"] = params
            save_config(self.config)
        elif filter_key == "laplacian":
            self.config["laplacian_filter_params"] = params
            save_config(self.config)
        return step, label

    def _populate_canvas_filter_menu(self, menu, canvas, view=None):
        """Populate a context menu with quick filter actions for a preview canvas."""
        if menu is None or canvas is None:
            return
        filt_menu = menu.addMenu("Filters")
        for key, info in FILTER_DEFINITIONS.items():
            act = QtWidgets.QAction(self._filter_action_label(key), filt_menu)
            if info.get("needs_gaussian") and not _gaussian_available():
                act.setEnabled(False)
                act.setToolTip("Requires scipy or OpenCV.")
            act.triggered.connect(lambda _, k=key: self._apply_filter_to_canvas(canvas, filter_key=k))
            filt_menu.addAction(act)
        filt_menu.addSeparator()
        filt_menu.addAction("Custom pipeline...", lambda: self._open_custom_filter_for_canvas(canvas))
        filt_menu.addAction("Clear filter", lambda: self._apply_filter_to_canvas(canvas, pipeline=[]))

    def _apply_filter_to_canvas(self, canvas, filter_key=None, pipeline=None, label=None):
        """Apply a filter pipeline to the views of a popup/preview canvas."""
        if not canvas or not getattr(canvas, "views", None):
            return
        steps = pipeline
        if steps is None and filter_key:
            original_views = self._clone_filter_source_views(canvas, canvas.views)
            base_arr = self._base_filter_image_from_views(original_views)
            preview_callback = self._build_canvas_filter_preview_callback(canvas, original_views)
            preview_cmap_name, preview_clim = self._filter_preview_render_state(original_views[0] if original_views else None)
            step, step_label = self._single_filter_step_spec(
                filter_key,
                parent=canvas,
                base_image=base_arr,
                preview_callback=preview_callback,
                preview_target_text=self._friendly_view_title(original_views[0] if original_views else None, "current image"),
                preview_cmap_name=preview_cmap_name,
                preview_clim=preview_clim,
                show_preview_thumbnail=False,
            )
            self._restore_filter_views_on_canvas(canvas, original_views)
            if step is None:
                return
            steps = [step]
            label = label or step_label
        self._set_filter_pipeline_on_canvas(canvas, steps, label=label, push_undo=True)

    def _open_custom_filter_for_canvas(self, canvas):
        if not canvas or not getattr(canvas, "views", None):
            return
        original_views = self._clone_filter_source_views(canvas, canvas.views)
        base_arr = self._base_filter_image_from_views(original_views)
        preview_cmap_name, preview_clim = self._filter_preview_render_state(original_views[0] if original_views else None)
        dlg = CustomFilterDialog(
            self,
            base_arr,
            self._run_filter_step,
            preview_callback=self._build_canvas_filter_preview_callback(canvas, original_views),
            preview_target_text=self._friendly_view_title(original_views[0] if original_views else None, "current image"),
            preview_cmap_name=preview_cmap_name,
            preview_clim=preview_clim,
            show_preview_thumbnail=False,
        )
        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            steps = dlg.pipeline_steps()
            label = dlg.pipeline_label()
            self._restore_filter_views_on_canvas(canvas, original_views)
            self._apply_filter_to_canvas(canvas, pipeline=steps, label=label)
            return
        self._restore_filter_views_on_canvas(canvas, original_views)

    def _apply_filter_pipeline(self, arr, steps):
        result = np.asarray(arr, dtype=float)
        for step in steps:
            result = self._run_filter_step(result, step)
        return result

    def _friendly_view_title(self, view, default="Preview"):
        if not view:
            return default
        title = view.get("title")
        if title:
            return title
        path = view.get("path") or ""
        label = view.get("label") or ""
        fname = ""
        try:
            fname = Path(str(path)).name if path else ""
        except Exception:
            fname = str(path)
        if fname and label:
            return f"{fname} - {label}"
        if fname:
            return fname
        if label:
            return label
        return default

    def _run_filter_step(self, arr, step):
        key = step.get('key')
        params = step.get('params', {})
        try:
            if key == 'flatten':
                axis = params.get('axis', 'both')
                return flatten_remove_median(arr, axis=axis)
            if key == 'tilt':
                return subtract_best_fit_plane(arr)
            if key == 'plane2':
                return subtract_2nd_order_plane(arr)
            if key == 'lowpass':
                sigma = params.get('sigma', 2.0)
                return gaussian_filter_image(arr, sigma)
            if key == 'highpass':
                sigma = params.get('sigma', 2.0)
                return highpass_filter(arr, sigma)
            if key == 'laplacian':
                sigma = params.get('sigma', FILTER_DEFINITIONS.get('laplacian', {}).get('default_sigma', 0.6))
                neighbors = params.get('neighbors', FILTER_DEFINITIONS.get('laplacian', {}).get('default_neighbors', 8))
                absolute = params.get('absolute', FILTER_DEFINITIONS.get('laplacian', {}).get('default_absolute', True))
                return laplacian_filter_image(arr, sigma=sigma, neighbors=neighbors, absolute=absolute)
        except Exception:
            pass
        return arr

    # ---------- dz helpers ----------
    def _dz_vs_previous_ch(self, header_path:Path):
        """Return dz pm and previous CH filename (most recent earlier file that is CH)."""
        key = str(header_path)
        info = self.tags.get(key, {})
        cur_abs = info.get('abs_z_pm', None)
        if cur_abs is None: return None, None
        try: idx = self.files.index(header_path)
        except ValueError:
            idx = None
            for i,p in enumerate(self.files):
                if str(p) == str(header_path): idx = i; break
        if idx is None: return None, None
        for j in range(idx-1, -1, -1):
            keyj = str(self.files[j]); infoj = self.tags.get(keyj, {})
            if infoj.get('tag') == 'constant-height' and infoj.get('abs_z_pm') is not None:
                return (cur_abs - infoj.get('abs_z_pm')), Path(keyj).name
        return None, None

    def _dz_vs_last_before_ch(self, header_path:Path):
        """Return dz pm vs last previous file that is not CH (e.g., last topo or CC before starting CH)."""
        key = str(header_path)
        info = self.tags.get(key, {})
        cur_abs = info.get('abs_z_pm', None)
        if cur_abs is None: return None, None
        try: idx = self.files.index(header_path)
        except ValueError:
            idx = None
            for i,p in enumerate(self.files):
                if str(p) == str(header_path): idx = i; break
        if idx is None: return None, None
        # search backwards for first previous file that is NOT CH
        for j in range(idx-1, -1, -1):
            keyj = str(self.files[j]); infoj = self.tags.get(keyj, {})
            if infoj.get('tag') != 'constant-height' and infoj.get('abs_z_pm') is not None:
                return (cur_abs - infoj.get('abs_z_pm')), Path(keyj).name
        return None, None

    # ---------- Add / Clear extra views ----------
    def on_add_view(self):
        if not hasattr(self, 'current_inspector_header') or self.current_inspector_header is None:
            QtWidgets.QMessageBox.information(self, "No file selected", "Please select a thumbnail first.")
            return
        hdr_path = Path(self.current_inspector_header); header, fds = self.headers.get(str(hdr_path), (None, None))
        if header is None: return
        dlg = QtWidgets.QDialog(self); dlg.setWindowTitle("Add channel view")
        v = QtWidgets.QVBoxLayout()
        listw = QtWidgets.QListWidget()
        for idx, fd in enumerate(fds):
            cap = fd.get('Caption', fd.get('FileName', f"chan{idx}"))
            it = QtWidgets.QListWidgetItem(f"{idx}: {cap}"); it.setData(QtCore.Qt.UserRole, idx); listw.addItem(it)
        v.addWidget(listw)
        hm = QtWidgets.QHBoxLayout()
        hm.addWidget(QtWidgets.QLabel("Cmap:"))
        cmapcombo = QtWidgets.QComboBox()
        # Populate cmap list with icons, falling back to a fixed list if colormaps is unavailable
        try:
            cmap_names = sorted(colormaps.keys())
        except Exception:
            cmap_names = ['viridis','plasma','inferno','magma','cividis','gray','hot','coolwarm','turbo']
        for name in cmap_names:
            try:
                icon = _colormap_icon(name, width=96, height=14)
            except Exception:
                icon = QIcon()
            cmapcombo.addItem(icon, name)
        if 'viridis' in cmap_names:
            try:
                cmapcombo.setCurrentText('viridis')
            except Exception:
                pass
        hm.addWidget(cmapcombo)
        v.addLayout(hm)
        btn_h = QtWidgets.QHBoxLayout(); add_btn = QtWidgets.QPushButton("Add"); cancel_btn = QtWidgets.QPushButton("Cancel")
        btn_h.addWidget(add_btn); btn_h.addWidget(cancel_btn); v.addLayout(btn_h)
        dlg.setLayout(v)
        add_btn.clicked.connect(dlg.accept); cancel_btn.clicked.connect(dlg.reject)
        if dlg.exec_() != QtWidgets.QDialog.Accepted: return
        sel = listw.currentItem()
        if not sel: QtWidgets.QMessageBox.information(self, "Choose channel", "Please select a channel to add."); return
        idx = sel.data(QtCore.Qt.UserRole); cmap = cmapcombo.currentText()
        # Record spec by caption and index; rebuild dynamically for selected file
        fd = fds[idx]
        cap = fd.get('Caption', fd.get('FileName', f"chan{idx}"))
        key = str(hdr_path)
        spec = self._ensure_extra_spec_entry(cap, idx, cmap)
        self._set_extra_spec_override(spec, key, cmap)
        if self.last_preview:
            self.show_file_channel(self.last_preview[0], self.last_preview[1])

    def _get_cached_header(self, path):
        """Return cached (header, fds) tuple if file is unchanged."""
        entry = self.header_cache.get(str(path))
        if not entry:
            return None
        try:
            mtime = Path(path).stat().st_mtime
        except Exception:
            return None
        if abs(entry.get('mtime', 0.0) - mtime) > 1e-6:
            return None
        header = entry.get('header')
        fds = entry.get('fds')
        if header is None or fds is None:
            return None
        return header, fds

    def _store_header_cache(self, path, header, fds):
        """Store parsed header info for future sessions."""
        try:
            mtime = Path(path).stat().st_mtime
        except Exception:
            return
        self.header_cache[str(path)] = {
            'mtime': mtime,
            'header': header,
            'fds': fds,
        }
        self._header_cache_dirty = True

    def _save_header_cache(self):
        if getattr(self, '_header_cache_dirty', False):
            save_header_cache(self.header_cache)
            self._header_cache_dirty = False

    def on_clear_views(self):
        self.added_views = []
        self.extra_view_specs = []
        if self.last_preview: self.show_file_channel(self.last_preview[0], self.last_preview[1])

    # ---------- helpers for extra view mapping ----------
    def _find_existing_extra_spec(self, caption, idx):
        """Return the stored spec entry for a given caption/index combo if it exists."""
        cap_norm = (caption or '').strip().lower()
        try:
            idx = int(idx)
        except Exception:
            idx = -1
        for spec in getattr(self, 'extra_view_specs', []):
            spec_cap = (spec.get('caption') or '').strip().lower()
            try:
                spec_idx = int(spec.get('index', -1))
            except Exception:
                spec_idx = -1
            if cap_norm and spec_cap and cap_norm == spec_cap:
                return spec
            if (not cap_norm) and idx != -1 and idx == spec_idx:
                return spec
        return None

    def _ensure_extra_spec_entry(self, caption, idx, cmap):
        """Fetch an existing spec entry or create a new one."""
        spec = self._find_existing_extra_spec(caption, idx)
        if spec is None:
            spec = {'caption': caption, 'index': int(idx), 'cmap': str(cmap), 'cmap_overrides': {}}
            self.extra_view_specs.append(spec)
        else:
            spec.setdefault('cmap_overrides', {})
            if 'cmap' not in spec or not spec['cmap']:
                spec['cmap'] = str(cmap)
        return spec

    def _resolve_extra_spec_cmap(self, spec, file_key):
        """Choose the best cmap for a spec, honoring per-file overrides when available."""
        if not spec:
            return self.preview_cmap_combo.currentText() or self.preview_cmap
        overrides = spec.get('cmap_overrides') or {}
        if file_key in overrides:
            return overrides[file_key]
        return spec.get('cmap', self.preview_cmap_combo.currentText() or self.preview_cmap)

    def _set_extra_spec_override(self, spec, file_key, cmap):
        """Store the cmap override for a spec/file pair."""
        if spec is None:
            return
        od = spec.setdefault('cmap_overrides', {})
        od[file_key] = str(cmap)

    def _find_channel_index_for_spec(self, fds, spec):
        """Given the list of file descriptors for a file and a spec dict
        {'caption': str, 'index': int, ...}, return the best matching channel index.
        Prefers exact caption match (case-insensitive), then substring match, then stored index.
        Returns None if no suitable channel is found.
        """
        if not fds:
            return None
        target_cap = (spec.get('caption') or '').strip().lower()
        if target_cap:
            # exact caption match
            for i, fd in enumerate(fds):
                cap_i = (fd.get('Caption','') or '').strip().lower()
                if cap_i == target_cap:
                    return i
            # substring caption match
            for i, fd in enumerate(fds):
                cap_i = (fd.get('Caption','') or '').strip().lower()
                if target_cap in cap_i and cap_i:
                    return i
            # try FileName match if caption didn't work
            for i, fd in enumerate(fds):
                fn_i = (fd.get('FileName','') or '').strip().lower()
                if fn_i == target_cap or (target_cap and target_cap in fn_i):
                    return i
        # fallback to stored index
        try:
            idx = int(spec.get('index', -1))
        except Exception:
            idx = -1
        if 0 <= idx < len(fds):
            return idx
        return None

    # ---------- Export PNGs ----------
    def _sanitize_filename_component(self, s: str) -> str:
        try:
            s = str(s)
        except Exception:
            s = ""
        # Replace invalid Windows filename chars and compress spaces
        s = re.sub(r'[<>:"/\\|?*]+', '_', s)
        s = s.strip().replace(' ', '_')
        s = re.sub(r'_+', '_', s)
        return s or "unnamed"

    def _get_adjust_spec(self, file_key, channel_idx):
        return (self.image_adjustments.get(str(file_key)) or {}).get(int(channel_idx))

    def _set_adjust_spec(self, file_key, channel_idx, spec):
        file_key = str(file_key)
        channel_idx = int(channel_idx)
        if spec:
            self.image_adjustments.setdefault(file_key, {})[channel_idx] = spec
        else:
            mapping = self.image_adjustments.get(file_key)
            if mapping and channel_idx in mapping:
                del mapping[channel_idx]
            if mapping and not mapping:
                self.image_adjustments.pop(file_key, None)

    def _apply_adjustments_for_channel(self, file_key, channel_idx, arr, extent):
        spec = self._get_adjust_spec(file_key, channel_idx)
        if not spec:
            return np.array(arr, dtype=float, copy=True), extent
        return apply_adjustment_spec(arr, extent, spec)

    def _scale_unit_for_display(self, unit, arr):
        arr_np = np.asarray(arr, dtype=float)
        unit_label = unit or ""
        factor = 1.0
        range_probe = arr_np
        if getattr(self, 'display_units_relative', False):
            finite = arr_np[np.isfinite(arr_np)]
            if finite.size:
                range_probe = arr_np - float(np.nanmin(finite))
        if unit_label:
            if getattr(self, 'display_units_si', False):
                target = _SI_BASE_UNITS.get(unit_label, (unit_label, 1.0))
                unit_label, factor = target
            else:
                unit_label, factor = _auto_display_unit(unit_label, range_probe)
        arr_scaled = arr_np * float(factor)
        zero_offset = None
        if getattr(self, 'display_units_relative', False):
            finite = arr_scaled[np.isfinite(arr_scaled)]
            if finite.size:
                zero_offset = float(np.nanmin(finite))
                arr_scaled = arr_scaled - zero_offset
        return unit_label or unit, arr_scaled, zero_offset

    def _collect_channel_exports(self, header_path_str, main_channel_idx=None):
        return viewer_export._collect_channel_exports(self, header_path_str, main_channel_idx)

    def _axes_from_extent(self, header, arr_shape, extent):
        h, w = arr_shape
        if extent:
            x_vals = np.linspace(extent[0], extent[1], w)
            y_vals = np.linspace(extent[2], extent[3], h)
        else:
            x_vals = np.arange(w, dtype=float)
            y_vals = np.arange(h, dtype=float)
        x_unit = (header.get('XPhysUnit') or header.get('PhysUnit') or 'px') if header else 'px'
        y_unit = (header.get('YPhysUnit') or header.get('PhysUnit') or 'px') if header else 'px'
        return x_vals, y_vals, x_unit, y_unit

    def _xyz_filename(self, header_path, caption):
        base = f"{header_path.stem} {caption}".strip()
        safe = re.sub(r'[<>:\"/\\|?*]+', '_', base)
        return f"{safe}.xyz"

    def _write_xyz_file(self, path, x_vals, y_vals, z_vals, x_unit, y_unit, z_unit, metadata_lines):
        log_status(f"Writing XYZ: {path}")
        with open(path, 'w', encoding='utf-8') as f:
            f.write("WSxM file copyright UAM\n")
            f.write("WSxM ASCII XYZ file\n")
            f.write(f"X[{x_unit}]\t\tY[{y_unit}]\t\tZ[{z_unit}]\n\n")
            for iy, y in enumerate(y_vals):
                for ix, x in enumerate(x_vals):
                    f.write(f"{x:.9g}\t{y:.9g}\t{z_vals[iy, ix]:.9g}\n")

    def on_export_pngs(self):
        return viewer_export.on_export_pngs(self)

    def on_export_xyz_files(self):
        return viewer_export.on_export_xyz_files(self)

    def on_adjust_image(self):
        if not self.last_preview or not hasattr(self, '_last_base_array'):
            QtWidgets.QMessageBox.information(self, "Adjust image", "Select an image first.")
            return
        file_key, channel_idx = self.last_preview
        base_arr = getattr(self, '_last_base_array', None)
        if base_arr is None:
            QtWidgets.QMessageBox.information(self, "Adjust image", "Image data not available.")
            return
        current_cmap = self.per_file_channel_cmap.get((file_key, int(channel_idx)), self.preview_cmap_combo.currentText() or self.preview_cmap)
        spec = self._get_adjust_spec(file_key, channel_idx) or {
            'crop': {'x0': 0, 'y0': 0, 'x1': base_arr.shape[1], 'y1': base_arr.shape[0]},
            'rotate': 0.0,
            'flip_h': False,
            'flip_v': False,
            'clip': {'low': None, 'high': None},
            'gamma': 1.0,
            'cmap': current_cmap,
        }
        spec.setdefault('cmap', current_cmap)
        base_extent = getattr(self, '_last_base_extent', None)
        axis_unit = getattr(self, '_last_axis_unit', 'px')
        display_extent = getattr(self, '_last_display_extent', None)
        colorbar_label = getattr(self, '_last_colorbar_label', None)
        dlg = ImageAdjustDialog(self, base_arr, spec, spec.get('cmap', current_cmap),
                                base_extent=base_extent, display_extent=display_extent,
                                axis_unit=axis_unit, colorbar_label=colorbar_label,
                                base_unit=getattr(self, '_last_base_unit', None),
                                relative_axes=bool(getattr(self, 'relative_axes', False)))
        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            new_spec = dlg.current_spec
            self._set_adjust_spec(file_key, channel_idx, new_spec)
            new_cmap = dlg.cmap_combo.currentText()
            if new_cmap:
                self.per_file_channel_cmap[(str(file_key), int(channel_idx))] = new_cmap
            self.show_file_channel(file_key, channel_idx)

    def _prepare_render_items(self, header_path, config):
        header_path = Path(header_path)
        header, fds = self.headers.get(str(header_path), (None, None))
        if header is None or fds is None:
            header, fds = parse_header(header_path)
            self.headers[str(header_path)] = (header, fds)
        try:
            xpix = int(header.get('xPixel', 128))
            ypix = int(header.get('yPixel', xpix))
        except Exception:
            xpix = 128; ypix = 128
        base_extent = self._header_extent(header)
        extent = self._display_extent(base_extent, header)
        render_items = []
        for desc in config.get('channels', []):
            key = desc.get('key') or f"idx_{desc.get('index')}"
            idx = None
            if desc.get('type') == 'index':
                idx = int(desc.get('index', -1))
            elif desc.get('type') == 'spec':
                idx = self._find_channel_index_for_spec(fds, desc.get('spec'))
            if idx is None or idx < 0 or idx >= len(fds):
                continue
            fd = fds[idx]
            fname = fd.get('FileName')
            try:
                unit_final, arr_conv = self._get_filtered_channel_array(str(header_path), idx, header, fd)
            except Exception:
                continue
            label = fd.get('Caption', fd.get('FileName', f"chan{idx}"))
            cmap = config.get('cmaps', {}).get(key, self.preview_cmap_combo.currentText() or self.preview_cmap)
            unit_display, arr_display, _ = self._scale_unit_for_display(unit_final, arr_conv)
            v_range = config.get('vmin_vmax', {}).get(key)
            vmin = vmax = None
            if isinstance(v_range, (list, tuple)) and len(v_range) == 2:
                vmin, vmax = v_range
            colorbar_label = label
            if unit_display:
                colorbar_label = f"{label} [{unit_display}]"
            title_text = f"{header_path.name} - {label}"
            render_items.append({'arr': arr_display, 'extent': extent, 'unit': unit_display, 'label': label,
                                 'cmap': cmap, 'vmin': vmin, 'vmax': vmax, 'relative_axes': bool(self.relative_axes),
                                 'colorbar_label': colorbar_label, 'title': title_text})
        return render_items

    def render_and_save_file_using_config(self, header_path, config, out_dir):
        """
        Render the given file using the supplied config (as returned by get_current_detail_config)
        and save a multi-panel PNG. Returns a list with the saved file path.
        """
        header_path = Path(header_path)
        out_dir = Path(out_dir)
        out_dir.mkdir(parents=True, exist_ok=True)
        render_items = self._prepare_render_items(header_path, config)
        if not render_items:
            raise ValueError("No matching channels for export.")
        fig_size = config.get('figure_size', (6, 5))
        if not isinstance(fig_size, (list, tuple)) or len(fig_size) != 2:
            fig_size = (6, 5)
        fig_w, fig_h = fig_size
        fig = Figure(figsize=(fig_w, fig_h), dpi=300)
        total = len(render_items)
        cols = int(math.ceil(math.sqrt(total)))
        rows = int(math.ceil(total / cols))
        for i, item in enumerate(render_items, 1):
            ax = fig.add_subplot(rows, cols, i)
            arr_plot = item['arr']
            flip = bool(item.get('relative_axes'))
            origin = 'lower' if flip else 'upper'
            if flip:
                arr_plot = np.flipud(arr_plot)
            im = ax.imshow(arr_plot, extent=item['extent'], origin=origin, interpolation='nearest',
                           aspect='equal' if item['extent'] else 'auto', cmap=item['cmap'],
                           vmin=item['vmin'], vmax=item['vmax'])
            if item.get('relative_axes') and item.get('extent') is not None:
                pass
            ax.set_title(item.get('title', item['label']), fontsize=9)
            ax.tick_params(labelsize=8)
            if item.get('colorbar_label') or item.get('unit'):
                cbar = fig.colorbar(im, ax=ax, fraction=0.08, pad=0.02)
                cbar.set_label(item.get('colorbar_label') or item.get('unit'))
        try:
            fig.tight_layout()
        except Exception:
            pass
        base = self._sanitize_filename_component(header_path.stem)
        chlist = "_".join([self._sanitize_filename_component(it['label']) for it in render_items])
        fname = f"{base}__channels_{chlist}.png"
        out_path = out_dir / fname
        counter = 1
        while out_path.exists():
            out_path = out_dir / f"{base}__channels_{chlist}_{counter}.png"
            counter += 1
        fig.savefig(out_path, dpi=300, bbox_inches='tight')
        return [str(out_path)]

    def copy_selected_as_svg(self, paths):
        """Render selected files to a single SVG and copy to clipboard."""
        import io
        import matplotlib
        from mpl_toolkits.axes_grid1.anchored_artists import AnchoredSizeBar

        if not paths:
            return
        
        config = self.get_current_detail_config()
        all_items = []
        for p in paths:
            try:
                items = self._prepare_render_items(p, config)
                if items:
                    all_items.extend(items)
            except Exception:
                pass
        
        if not all_items:
            QtWidgets.QMessageBox.warning(self, "Copy SVG", "No valid data found in selection.")
            return

        # Layout: simple grid
        total = len(all_items)
        cols = int(math.ceil(math.sqrt(total)))
        rows = int(math.ceil(total / cols))
        
        # Base size on config but scale up for grid
        base_w, base_h = config.get('figure_size', (6, 5))
        fig = Figure(figsize=(base_w * cols, base_h * rows))
        
        # Apply theme to figure background
        dark = bool(self.detail_dark_view)
        fig_face = '#111217' if dark else '#ffffff'
        fig.set_facecolor(fig_face)
        
        # Text color for axes titles etc
        text_color = '#f5f5f5' if dark else '#111111'
        
        sb_enabled = self.scale_bar_cb.isChecked()
        sb_pos = getattr(self.preview_canvas, '_scale_bar_pos', (0.94, 0.06))
        
        # Scale bar settings
        sb_settings = getattr(self.preview_canvas, '_scale_bar_settings', {})
        sb_font = sb_settings.get('font_family', 'sans-serif')
        sb_text_col = sb_settings.get('text_color') or text_color
        sb_bar_col = sb_settings.get('bar_color') or text_color
        font_scale = getattr(self.preview_canvas, '_view_font_scale', 1.0)
        show_ticks = getattr(self.preview_canvas, '_show_ticks', True)
        show_cbar = getattr(self.preview_canvas, '_show_colorbar', True)

        for i, item in enumerate(all_items, 1):
            ax = fig.add_subplot(rows, cols, i)
            arr_plot = item['arr']
            flip = bool(item.get('relative_axes'))
            origin = 'lower' if flip else 'upper'
            if flip:
                arr_plot = np.flipud(arr_plot)
            
            im = ax.imshow(arr_plot, extent=item['extent'], origin=origin, interpolation='nearest',
                           aspect='equal' if item['extent'] else 'auto', cmap=item['cmap'],
                           vmin=item['vmin'], vmax=item['vmax'])
            
            ax.set_title(item.get('title', item['label']), fontsize=9 * font_scale, color=text_color)
            ax.tick_params(labelsize=8 * font_scale, colors=text_color, labelcolor=text_color)
            for spine in ax.spines.values():
                spine.set_color(text_color)
            
            if not show_ticks:
                ax.set_xticks([])
                ax.set_yticks([])
            
            cbar_label = item.get('colorbar_label') or item.get('unit')
            if cbar_label and show_cbar:
                try:
                    divider = make_axes_locatable(ax)
                    cax = divider.append_axes("right", size="5%", pad=0.05)
                    cbar = fig.colorbar(im, cax=cax, orientation='vertical')
                    cbar.set_label(cbar_label, size=10 * font_scale)
                    cbar.ax.yaxis.label.set_color(text_color)
                    cbar.ax.tick_params(colors=text_color, labelcolor=text_color, labelsize=8 * font_scale)
                    if not show_ticks:
                        cbar.set_ticks([])
                    cbar.outline.set_edgecolor(text_color)
                    cbar.ax.yaxis.set_label_coords(0.5, 0.5)
                    cbar.ax.yaxis.label.set_horizontalalignment('center')
                    cbar.ax.yaxis.label.set_verticalalignment('center')
                except Exception:
                    pass
            
            if sb_enabled and self.preview_canvas:
                # Reuse logic from canvas to calculate size
                width = abs(item['extent'][1] - item['extent'][0]) if item['extent'] else arr_plot.shape[1]
                unit = 'nm' if item['extent'] else 'px' # simplified assumption based on prepare_render_items
                size, label = self.preview_canvas._calculate_best_scale_bar(width, unit)
                sb = AnchoredSizeBar(ax.transData, size, label, loc='center',
                                     pad=0.4, borderpad=0, sep=3, frameon=False,
                                     size_vertical=width*0.004*font_scale, color=sb_bar_col,
                                     label_top=True,
                                     bbox_to_anchor=sb_pos, bbox_transform=ax.transAxes)
                sb.size_bar.get_children()[0].set_linewidth(0)
                text = sb.txt_label.get_children()[0]
                text.set_color(sb_text_col)
                text.set_fontsize(10 * font_scale)
                text.set_fontweight('bold')
                ax.add_artist(sb)

        buf = io.BytesIO()
        with matplotlib.rc_context({'svg.fonttype': 'none'}):
            fig.savefig(buf, format="svg", bbox_inches="tight")
        mime = QtCore.QMimeData()
        mime.setData("image/svg+xml", buf.getvalue())
        QtWidgets.QApplication.clipboard().setMimeData(mime)

    # ---------- Profile measurement (interactive line) ----------
    def _on_start_profile(self, force_enable=False):
        return viewer_measurement._on_start_profile(self, force_enable=force_enable)

    def _on_start_angle(self, force_enable=False):
        return viewer_measurement._on_start_angle(self, force_enable=force_enable)

    def _disable_profile_mode(self):
        return viewer_measurement._disable_profile_mode(self)

    def _disable_angle_mode(self, reset_button=True):
        return viewer_measurement._disable_angle_mode(self, reset_button=reset_button)

    def _on_exit_profile_mode(self):
        return viewer_measurement._on_exit_profile_mode(self)

    def _on_clear_profile_measurement(self):
        return viewer_measurement._on_clear_profile_measurement(self)

    def _on_profile_updated(self, active_profile, saved_profiles):
        return viewer_measurement._on_profile_updated(self, active_profile, saved_profiles)

    def _on_angle_updated(self, info):
        return viewer_measurement._on_angle_updated(self, info)

    def _on_show_profile_window(self):
        return viewer_measurement._on_show_profile_window(self)

    def _on_canvas_overlay_highlight(self, idx):
        return viewer_measurement._on_canvas_overlay_highlight(self, idx)

    def _resolve_toast_host(self, target=None):
        host = None
        if isinstance(target, QtWidgets.QWidget):
            host = target.window() if target.window() is not None else target
        if host is None:
            host = self
        return host

    def _show_toast(self, message, *, duration_ms=1400, target=None, variant="default"):
        text = str(message or "").strip()
        if not text:
            return
        host = self._resolve_toast_host(target)
        if host is None:
            return
        try:
            if sip.isdeleted(host):
                host = self
        except Exception:
            pass
        if not isinstance(host, QtWidgets.QWidget):
            host = self
        key = int(id(host))
        entry = self._toast_registry.get(key)
        toast = None
        timer = None
        if entry:
            toast, timer = entry
            try:
                if sip.isdeleted(toast):
                    toast = None
            except Exception:
                toast = None
        if toast is None:
            toast = QtWidgets.QLabel(host)
            toast.setObjectName("appToast")
            toast.setAlignment(QtCore.Qt.AlignCenter)
            toast.setWordWrap(True)
            toast.setTextFormat(QtCore.Qt.PlainText)
            toast.setAttribute(QtCore.Qt.WA_TransparentForMouseEvents, True)
            timer = QtCore.QTimer(toast)
            timer.setSingleShot(True)
            timer.timeout.connect(toast.hide)
            self._toast_registry[key] = (toast, timer)
        styles = {
            "default": (
                "QLabel#appToast {"
                "background-color: rgba(18, 24, 34, 212);"
                "color: #f5f7fb;"
                "border: 1px solid rgba(255,255,255,55);"
                "border-radius: 10px;"
                "padding: 6px 12px;"
                "font-weight: 600;"
                "}"
            ),
            "success": (
                "QLabel#appToast {"
                "background-color: rgba(14, 92, 54, 232);"
                "color: #f5fff8;"
                "border: 1px solid rgba(170, 248, 204, 200);"
                "border-radius: 12px;"
                "padding: 8px 14px;"
                "font-weight: 700;"
                "}"
            ),
        }
        toast.setStyleSheet(styles.get(str(variant or "default"), styles["default"]))
        toast.setText(text)
        rect = host.rect()
        margin = 14
        max_width = max(260, min(760, int(rect.width() - (margin * 2))))
        metrics = QtGui.QFontMetrics(toast.font())
        flags = int(QtCore.Qt.AlignCenter | QtCore.Qt.TextWordWrap | QtCore.Qt.TextWrapAnywhere)
        text_rect = metrics.boundingRect(QtCore.QRect(0, 0, max_width - 28, 2000), flags, text)
        toast_width = min(max_width, max(180, int(text_rect.width() + 28)))
        toast_height = max(36, int(text_rect.height() + 18))
        toast.resize(toast_width, toast_height)
        x = max(margin, int((rect.width() - toast.width()) * 0.5))
        y = max(margin, int(rect.height() - toast.height() - margin))
        toast.move(x, y)
        toast.show()
        toast.raise_()
        if timer is not None:
            timer.start(max(900, int(duration_ms)))

    def _show_saved_path_toast(self, title, path, *, detail=None, duration_ms=4200, target=None):
        if not path:
            return
        try:
            path_obj = Path(path)
            file_name = path_obj.name
            full_path = str(path_obj)
        except Exception:
            file_name = ""
            full_path = str(path)
        lines = [str(title or "Saved").strip()]
        if detail:
            lines[0] = f"{lines[0]} | {str(detail).strip()}"
        if file_name:
            lines.append(file_name)
        if full_path and full_path != file_name:
            lines.append(full_path)
        host = target
        if host is None:
            try:
                host = QtWidgets.QApplication.activeWindow()
            except Exception:
                host = None
        self._show_toast("\n".join(line for line in lines if line), duration_ms=duration_ms, target=host, variant="success")

    def _on_view_copied(self, view=None, info=None, target=None):
        if not isinstance(view, dict):
            view = {}
        if not isinstance(info, dict):
            info = {}
        fmt = str(info.get("format") or "png").upper()
        displayed = bool(info.get("displayed", False))
        canvas = info.get("canvas")
        host = target if isinstance(target, QtWidgets.QWidget) else canvas
        if displayed:
            msg = f"Copied displayed image ({fmt})"
        else:
            title = view.get('title') or 'Image'
            msg = f"Copied '{title}' ({fmt})"
        self._show_toast(msg, duration_ms=1400, target=host)

    def _on_preview_value(self, value, x, y, view):
        return viewer_preview._on_preview_value(self, value, x, y, view)

    def _view_finite_values(self, view):
        if not view:
            return None, None, None
        try:
            arr = np.asarray(view.get("arr"))
        except Exception:
            return None, None, None
        finite = arr[np.isfinite(arr)]
        if finite.size == 0:
            return None, None, None
        return float(finite.min()), float(finite.max()), finite

    def _apply_clim_to_view(self, canvas, view, lo, hi):
        if canvas is None or view is None:
            return
        try:
            canvas.set_view_clim(view, (float(lo), float(hi)))
        except Exception:
            pass

    def _auto_contrast(self, canvas, pct_low=1.0, pct_high=99.0):
        view = canvas.views[0] if canvas and getattr(canvas, "views", None) else None
        vmin, vmax, finite = self._view_finite_values(view)
        if finite is None:
            return
        try:
            canvas.push_undo_state("auto_contrast")
        except Exception:
            pass
        try:
            lo, hi = np.percentile(finite, [pct_low, pct_high])
        except Exception:
            lo, hi = vmin, vmax
        if view and bool(view.get("display_relative_zero", False)):
            lo = 0.0
            hi = max(float(hi), float(vmax or 0.0), 0.0)
        self._apply_clim_to_view(canvas, view, lo, hi)

    def _reset_contrast(self, canvas):
        view = canvas.views[0] if canvas and getattr(canvas, "views", None) else None
        vmin, vmax, _ = self._view_finite_values(view)
        if vmin is None:
            return
        try:
            canvas.push_undo_state("reset_contrast")
        except Exception:
            pass
        if view and bool(view.get("display_relative_zero", False)):
            vmin = 0.0
        self._apply_clim_to_view(canvas, view, vmin, vmax)

    def _open_histogram_dialog(self, canvas):
        open_histogram_dialog(self, canvas)

    def _is_matrix_spec(self, spec) -> bool:
        try:
            if not spec:
                return False
            if spec.get('matrix_dataset'):
                return True
            if spec.get('matrix_index') is None:
                return False
            return is_matrix_file_entry(spec)
        except Exception:
            return False

    def _on_preview_spec_click(self, spec, event=None):
        controller = getattr(self, "spectro_compare_controller", None)
        if controller:
            controller.handle_preview_click(spec, event=event)

    def on_manual_tag(self, tag):
        if self.last_preview is None:
            QtWidgets.QMessageBox.information(self, "No file selected", "Please select a thumbnail first."); return
        header_path_str, ch_idx = self.last_preview; header_path = Path(header_path_str); key = str(header_path)
        if tag is None:
            if key in self.tags:
                del self.tags[key]
        else:
            info = {'tag': tag, 'manual': True}
            if tag == 'constant-height':
                try:
                    hdr, fds = self.headers.get(key)
                    topo_idx = _find_topography_channel(fds)
                    if topo_idx is None:
                        topo_idx = ch_idx
                    fd = fds[topo_idx]
                    arr = self._get_channel_array(key, topo_idx, hdr, fd)
                    _, arr_nm = normalize_unit_and_data(arr, fd.get('PhysUnit',''))
                    tag_info = self._classify_topography_values(arr_nm)
                    info['abs_z_pm'] = tag_info.get('abs_pm') if tag_info else None
                    info['rng_nm'] = tag_info.get('rng_nm') if tag_info else None
                except Exception:
                    info['abs_z_pm'] = None
                    info['rng_nm'] = None
            self.tags[key] = info
        self.config['tags'] = self.tags; save_config(self.config)
        # refresh thumbnails & preview (so badges/metadata update)
        self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())
        if self.last_preview: self.show_file_channel(self.last_preview[0], self.last_preview[1])

    def _clear_auto_tags(self, *, persist: bool = True) -> bool:
        """Remove auto-detected CH/CC tags while preserving manual overrides."""
        changed = False
        kept = {}
        for key, value in dict(self.tags or {}).items():
            if isinstance(value, dict) and value.get("auto") and not value.get("manual"):
                changed = True
                continue
            kept[str(key)] = value
        if changed:
            self.tags = kept
            self.config["tags"] = self.tags
            if persist:
                save_config(self.config)
        return changed

    def _on_toggle_auto_tags(self, checked: bool):
        self.auto_detect_tags = bool(checked)
        self.config['auto_detect_tags'] = self.auto_detect_tags
        save_config(self.config)
        if checked:
            try:
                self._auto_detect_tags_for_folder()
                # refresh thumbnails to show badges
                self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())
                if self.last_preview:
                    self.show_file_channel(self.last_preview[0], self.last_preview[1])
            except Exception:
                pass
        else:
            changed = self._clear_auto_tags()
            if changed:
                self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())
                if self.last_preview:
                    self.show_file_channel(self.last_preview[0], self.last_preview[1])

    # ---------- Spectroscopy helpers ----------
    def on_load_molecule(self):
        if not self.preview_canvas:
            return
        # Offer recent molecules first for quick access
        recent = []
        try:
            recent = self.preview_canvas.get_recent_molecule_paths()
        except Exception:
            recent = []
        if recent:
            menu = QtWidgets.QMenu(self)
            actions = {}
            for p in recent:
                act = menu.addAction(Path(p).name)
                act.setToolTip(p)
                actions[act] = p
            browse_act = menu.addAction("Browse...")
            chosen = menu.exec_(QtGui.QCursor.pos())
            if chosen and chosen in actions:
                self.preview_canvas.add_molecule(actions[chosen])
                self.on_show_molecules_toggled(True)
                self._on_recent_molecules_updated(self.preview_canvas.get_recent_molecule_paths())
                return
        # Fallback: open file dialog
        self.preview_canvas._load_molecule_dialog()
        self.on_show_molecules_toggled(True)

    def on_spec_folder_browse(self):
        folder = QtWidgets.QFileDialog.getExistingDirectory(self, "Select spectroscopy folder", str(self.spec_folder_path))
        if folder:
            self.spec_folder_le.setText(folder)
            self._set_spec_folder(Path(folder))

    def on_spec_folder_entered(self):
        text = self.spec_folder_le.text().strip()
        if not text:
            return
        self._set_spec_folder(Path(text))

    def _set_spec_folder(self, path:Path, *, force_reload: bool = False):
        current = getattr(self, "spec_folder_path", None)
        same_folder = False
        try:
            if current is not None:
                same_folder = Path(current).resolve() == Path(path).resolve()
        except Exception:
            same_folder = str(current or "") == str(path or "")
        try:
            self.spec_folder_path = Path(path)
            self.config['spectra_folder'] = str(self.spec_folder_path)
            save_config(self.config)
        except Exception:
            pass
        if same_folder and not force_reload:
            return
        # Changing the spectroscopy path should refresh spectroscopy immediately.
        self._reload_spectros(refresh=True)

    def ensure_spectros_loaded(self, refresh: bool = True):
        """Load spectroscopies on-demand if they were deferred."""
        if self._spectros_loaded:
            return True
        if getattr(self, "_spectros_loading", False):
            return False
        try:
            self._spectro_autoload_timer.stop()
        except Exception:
            pass
        self._spectros_loading = True
        self._spectros_pending = False
        try:
            log_status("[Lazy] Loading spectroscopy references...")
            self._reload_spectros(refresh=refresh)
            if not refresh and self._spectros_loaded:
                self._schedule_marker_refresh()
                if self.last_preview:
                    try:
                        self.show_file_channel(self.last_preview[0], self.last_preview[1])
                    except Exception:
                        pass
        finally:
            self._spectros_loading = False
        return True

    def _schedule_pending_spectro_load(self, delay_ms: int = 1200):
        if self._spectros_loaded or not getattr(self, "_spectros_pending", False):
            return
        if getattr(self, "_spectros_loading", False):
            return
        try:
            self._spectro_autoload_timer.start(max(0, int(delay_ms)))
        except Exception:
            pass

    def _run_pending_spectro_load(self):
        if self._spectros_loaded or not getattr(self, "_spectros_pending", False):
            return
        if getattr(self, "_spectros_loading", False):
            return
        self.ensure_spectros_loaded(refresh=True)

    def _reload_spectros(self, refresh=True):
        # unless we complete a successful reload, consider spectra cache stale
        self._spectros_loaded = False
        self._spectros_pending = False
        self._spectro_miniature_cache.clear()
        t_scan_start = time.perf_counter()
        try:
            folder = getattr(self, 'spec_folder_path', None) or self.last_dir
            folder = Path(folder)
        except Exception:
            folder = self.last_dir
        log_status(f"Scanning spectroscopy files in: {folder}")
        self._spectro_deferred = set()
        self.spectros, spec_stats = self._scan_spectros(folder)
        t_scan_end = time.perf_counter()
        if spec_stats:
            total_entries = spec_stats.get('total_specs', len(self.spectros))
            single_files = spec_stats.get('single_dat_files', 0)
            single_entries = spec_stats.get('single_entries', single_files)
            matrix_files = spec_stats.get('matrix_dat_files', 0)
            matrix_entries = spec_stats.get('matrix_specs', 0)
            # keep stats for UI but avoid duplicate terminal spam (loader already logged)
        else:
            log_status(f"Loaded {len(self.spectros)} spectroscopy entries")
        t_assign_start = time.perf_counter()
        self._assign_spectros_to_images()
        t_assign_end = time.perf_counter()
        self.matrix_spectros = [spec for spec in self.spectros if spec.get('matrix_index') is not None]
        self._clear_multi_spec_selection()
        self._update_spectro_stats_label(spec_stats)
        self._spectros_loaded = True
        self._update_matrix_summary_banner()
        if refresh:
            t_thumb_start = time.perf_counter()
            self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())
            if self.last_preview:
                self.show_file_channel(self.last_preview[0], self.last_preview[1])
            t_thumb_end = time.perf_counter()
        else:
            t_thumb_start = t_assign_end
            t_thumb_end = t_assign_end
        scan_ms = (t_scan_end - t_scan_start) * 1000.0
        assign_ms = (t_assign_end - t_assign_start) * 1000.0
        thumb_ms = (t_thumb_end - t_thumb_start) * 1000.0
        log_status(f"[Perf] Spectros: scan {scan_ms:.0f} ms | assign {assign_ms:.0f} ms | thumbs {thumb_ms:.0f} ms")

    def _scan_spectros(self, folder:Path):
        return viewer_loader._scan_spectros(self, folder)

    def hydrate_spectro_entry(self, spec):
        return viewer_loader.hydrate_spectro_file(self, spec)

    def hydrate_spectro_entries(self, specs):
        return viewer_loader.hydrate_spectro_entries(self, specs)

    def refresh_spectro_manifest(self):
        return viewer_loader.refresh_spectro_manifest_from_viewer(self)

    def _schedule_spectro_manifest_save(self):
        self._spectro_manifest_save_pending = True
        try:
            self._spectro_manifest_save_timer.start(400)
        except Exception:
            self._flush_spectro_manifest_save()

    def _flush_spectro_manifest_save(self):
        if self._spectro_manifest_save_inflight:
            self._spectro_manifest_save_pending = True
            return
        folder = getattr(self, "spec_folder_path", None) or getattr(self, "last_dir", None)
        manifest_entries = dict(getattr(self, "_spectro_manifest_entries", {}) or {})
        if not folder or not manifest_entries:
            self._spectro_manifest_save_pending = False
            return
        self._spectro_manifest_save_pending = False
        self._spectro_manifest_save_inflight = True

        def _persist(snapshot_folder, snapshot_manifest):
            try:
                viewer_loader.save_spectro_manifest_snapshot(snapshot_folder, snapshot_manifest)
            finally:
                QtCore.QTimer.singleShot(0, self._on_spectro_manifest_save_finished)

        threading.Thread(
            target=_persist,
            args=(folder, manifest_entries),
            name="spectro-manifest-save",
            daemon=True,
        ).start()

    def _on_spectro_manifest_save_finished(self):
        self._spectro_manifest_save_inflight = False
        if self._spectro_manifest_save_pending:
            self._schedule_spectro_manifest_save()

    def _assign_spectros_to_images(self):
        spectro_controller._assign_spectros_to_images(self)
        try:
            self.files_with_matrix = {
                key for key, entries in (self.spectros_by_image or {}).items()
                if any(spec.get('matrix_index') is not None for spec in entries)
            }
        except Exception:
            self.files_with_matrix = set()
        try:
            QtCore.QTimer.singleShot(0, self.refresh_spectro_manifest)
        except Exception:
            pass
        self._update_matrix_summary_banner()

    def _choose_image_for_spec(self, spec, images, image_extents):
        return spectro_controller._choose_image_for_spec(self, spec, images, image_extents)

    def _extent_center(self, extent):
        return spectro_controller._extent_center(self, extent)

    def _spec_within_extent(self, sx, sy, extent, margin_frac=0.05):
        return spectro_controller._spec_within_extent(self, sx, sy, extent, margin_frac=margin_frac)

    def _match_spec_to_image_by_hint(self, spec, images, with_score=False):
        return spectro_controller._match_spec_to_image_by_hint(self, spec, images, with_score=with_score)

    def _map_spec_to_pixels(self, spec, header, xpix, ypix, file_key=None, thumb_crop=None):
        try:
            x = float(spec.get('x'))
            y = float(spec.get('y'))
        except Exception:
            x = y = None
        if x is None or y is None:
            # fallback placement using a stable order index if present
            try:
                idx = int(spec.get('order_idx', 1))
            except Exception:
                idx = 1
            return self._fallback_spec_coords(idx, xpix, ypix)
        try:
            extent = self._header_extent(header) if header is not None else [0.0, 1.0, 1.0, 0.0]
        except Exception:
            extent = [0.0, 1.0, 1.0, 0.0]
        x0, x1, y1, y0 = extent
        xspan = x1 - x0
        yspan = y1 - y0
        if xspan <= 0 or yspan <= 0:
            # try to map using spectroscopy cloud extents if available
            fallback = self._map_spec_by_spec_extent(file_key, spec, xpix, ypix)
            if fallback is not None:
                col, row = fallback
                return self._apply_thumb_crop_to_coords(col, row, xpix, ypix, thumb_crop)
            grid_fallback = self._map_spec_by_grid(spec, xpix, ypix)
            if grid_fallback is not None:
                col, row = grid_fallback
                return self._apply_thumb_crop_to_coords(col, row, xpix, ypix, thumb_crop)
            return None
        cx = 0.5 * (x0 + x1)
        cy = 0.5 * (y0 + y1)
        dx_norm = (x - cx) / xspan
        dy_norm = (y - cy) / yspan
        angle_deg = self._header_scan_angle(header)
        if angle_deg:
            theta = math.radians(angle_deg)
            cos_t = math.cos(theta)
            sin_t = math.sin(theta)
            u = dx_norm * cos_t + dy_norm * sin_t
            v = -dx_norm * sin_t + dy_norm * cos_t
        else:
            u = dx_norm
            v = dy_norm
        frac_x = (u + 0.5)
        frac_y = (0.5 - v)
        if not (0.0 <= frac_x <= 1.0 and 0.0 <= frac_y <= 1.0):
            # try spectroscopy cloud extent before clamping/grid
            fallback = self._map_spec_by_spec_extent(file_key, spec, xpix, ypix)
            if fallback is not None:
                col, row = fallback
                return self._apply_thumb_crop_to_coords(col, row, xpix, ypix, thumb_crop)
            grid_pt = self._map_spec_by_grid(spec, xpix, ypix)
            if grid_pt is not None:
                col, row = grid_pt
                return self._apply_thumb_crop_to_coords(col, row, xpix, ypix, thumb_crop)
            frac_x = min(max(frac_x, 0.0), 1.0)
            frac_y = min(max(frac_y, 0.0), 1.0)
        cols = max(1, int(xpix) - 1)
        rows = max(1, int(ypix) - 1)
        col = frac_x * cols
        row = frac_y * rows
        return self._apply_thumb_crop_to_coords(col, row, xpix, ypix, thumb_crop)

    def _apply_thumb_crop_to_coords(self, col, row, xpix, ypix, thumb_crop):
        if thumb_crop is None:
            return col, row
        try:
            r0 = int(thumb_crop.get("r0"))
            r1 = int(thumb_crop.get("r1"))
        except Exception:
            return col, row
        if r1 <= r0:
            return col, row
        crop_rows = r1 - r0 + 1
        try:
            row = float(row) - float(r0)
        except Exception:
            return col, row
        row = min(max(row, 0.0), max(0.0, crop_rows - 1))
        return col, row

    def _map_spec_by_spec_extent(self, file_key, spec, xpix, ypix):
        """Fallback mapping using the min/max of all specs for this image to keep real-space layout."""
        if not file_key:
            file_key = spec.get('image_key')
        if not file_key:
            return None
        entries = self.spectros_by_image.get(str(file_key), [])
        xs = [s.get('x') for s in entries if s.get('x') is not None]
        ys = [s.get('y') for s in entries if s.get('y') is not None]
        if not xs or not ys:
            return None
        try:
            xmin, xmax = float(min(xs)), float(max(xs))
            ymin, ymax = float(min(ys)), float(max(ys))
        except Exception:
            return None
        # pad spans to avoid zero division
        span_x = xmax - xmin
        span_y = ymax - ymin
        if span_x == 0 or span_y == 0:
            span_x = span_x or 1.0
            span_y = span_y or 1.0
        try:
            x = float(spec.get('x')); y = float(spec.get('y'))
        except Exception:
            return None
        frac_x = (x - xmin) / span_x
        frac_y = (ymax - y) / span_y
        frac_x = min(max(frac_x, 0.0), 1.0)
        frac_y = min(max(frac_y, 0.0), 1.0)
        col = frac_x * max(1, xpix - 1)
        row = frac_y * max(1, ypix - 1)
        return col, row

    def _map_spec_by_grid(self, spec, xpix, ypix):
        grid_cols = spec.get('grid_cols')
        grid_rows = spec.get('grid_rows')
        if not grid_cols or not grid_rows:
            return None
        try:
            col_idx = int(spec.get('grid_col', 0))
            row_idx = int(spec.get('grid_row', 0))
        except Exception:
            return None
        cols = max(1, int(grid_cols) - 1)
        rows = max(1, int(grid_rows) - 1)
        if grid_cols <= 0 or grid_rows <= 0:
            return None
        col_frac = col_idx / cols if cols > 0 else 0.0
        row_frac = row_idx / rows if rows > 0 else 0.0
        col = col_frac * max(1, xpix - 1)
        row = row_frac * max(1, ypix - 1)
        return col, row

    def _fallback_spec_coords(self, idx, xpix, ypix):
        """Fallback placement for specs lacking coordinates: spread markers on a 3x3 grid."""
        slots = [
            (0.15, 0.15), (0.50, 0.15), (0.85, 0.15),
            (0.15, 0.50), (0.50, 0.50), (0.85, 0.50),
            (0.15, 0.85), (0.50, 0.85), (0.85, 0.85),
        ]
        frac_x, frac_y = slots[(idx - 1) % len(slots)]
        col = frac_x * max(1, xpix - 1)
        row = frac_y * max(1, ypix - 1)
        return col, row

    def _render_spectroscopy_overlays(
        self,
        pixmap,
        header,
        file_key,
        xpix,
        ypix,
        reveal_points_override=None,
        selected_spec=None,
        entries_override=None,
        matrix_as_points=False,
        thumb_crop=None,
    ):
        """Render spectroscopy markers directly on the thumbnail pixmap."""
        if not self.show_spectra and not reveal_points_override:
            return []
        if not self._spectros_loaded:
            if getattr(self, "lazy_spectros_enabled", False) and getattr(self, "_spectros_pending", False):
                try:
                    self._schedule_pending_spectro_load(delay_ms=250)
                except Exception:
                    pass
            return []
        return spectro_overlays._render_spectroscopy_overlays(
            self,
            pixmap,
            header,
            file_key,
            xpix,
            ypix,
            reveal_points_override=reveal_points_override,
            selected_spec=selected_spec,
            entries_override=entries_override,
            matrix_as_points=matrix_as_points,
            thumb_crop=thumb_crop,
        )

    def _matrix_bbox_pixels(self, m_specs, header, xpix, ypix, w_scale, h_scale, file_key=None, thumb_crop=None):
        xs = []
        ys = []
        for idx, spec in enumerate(m_specs, 1):
            c = self._map_spec_to_pixels(spec, header, xpix, ypix, file_key, thumb_crop=thumb_crop)
            if c is None:
                c = self._fallback_spec_coords(idx, xpix, ypix)
            col, row = c
            xs.append(col * w_scale)
            ys.append(row * h_scale)
        if not xs or not ys:
            return None
        xmin = min(xs); xmax = max(xs)
        ymin = min(ys); ymax = max(ys)
        width = max(xmax - xmin, 0.0)
        height = max(ymax - ymin, 0.0)
        max_w = max(1.0, (max(xpix - 1, 1)) * w_scale)
        crop_rows = None
        if thumb_crop:
            try:
                r0 = int(thumb_crop.get("r0"))
                r1 = int(thumb_crop.get("r1"))
                if r1 > r0:
                    crop_rows = r1 - r0 + 1
            except Exception:
                crop_rows = None
        y_denom = max(1, (crop_rows - 1)) if crop_rows else max(1, ypix - 1)
        max_h = max(1.0, y_denom * h_scale)
        if width == 0 and height == 0:
            base = min(max_w, max_h) * 0.2
            base = max(base, 18.0)
            return QtCore.QRectF(xmin - base / 2.0, ymin - base / 2.0, base, base)
        min_span = min(max_w, max_h) * 0.12
        width = max(width, min_span)
        height = max(height, min_span)
        pad = max(4.0, min(14.0, min(max_w, max_h) * 0.05))
        cx = (xmax + xmin) / 2.0
        cy = (ymax + ymin) / 2.0
        rect = QtCore.QRectF(
            cx - width / 2.0 - pad,
            cy - height / 2.0 - pad,
            width + 2 * pad,
            height + 2 * pad,
        )
        scene_rect = QtCore.QRectF(0.0, 0.0, max_w, max_h)
        rect = rect.intersected(scene_rect)
        return rect

    def _label_pos_to_pix_coords(self, label_widget, pos):
        pix = label_widget.pixmap()
        if pix is None:
            return None
        offset_x = (label_widget.width() - pix.width()) / 2.0
        offset_y = (label_widget.height() - pix.height()) / 2.0
        x = pos.x() - offset_x
        y = pos.y() - offset_y
        if x < 0 or y < 0 or x > pix.width() or y > pix.height():
            return None
        return x, y

    def _scroll_to_thumbnail(self, file_key):
        controller = getattr(self, "thumbnail_controller", None)
        if controller:
            return controller.scroll_to_thumbnail(file_key)

    def _handle_thumbnail_navigation(self, key, modifiers=QtCore.Qt.NoModifier):
        controller = getattr(self, "thumbnail_controller", None)
        if controller:
            return controller.handle_navigation(key, modifiers=modifiers)
        return False

    def _activate_thumbnail_by_index(self, index):
        controller = getattr(self, "thumbnail_controller", None)
        if controller:
            return controller.activate_thumbnail_by_index(index)
        return False

    def _focus_first_matrix_dataset(self):
        controller = getattr(self, "thumbnail_controller", None)
        if controller:
            return controller.focus_first_matrix_dataset()

    def _update_matrix_summary_banner(self):
        controller = getattr(self, "thumbnail_controller", None)
        if controller:
            return controller.update_matrix_summary_banner()

    def _handle_spec_marker_click(self, label_widget, event):
        if getattr(event, 'button', None) and event.button() != QtCore.Qt.LeftButton:
            return False
        if not self.show_spectra:
            return False
        markers = label_widget.property("spec_markers") or []
        if not markers:
            return False
        coords = self._label_pos_to_pix_coords(label_widget, event.pos())
        if coords is None:
            return False
        x, y = coords
        file_key = str(label_widget.property("file_path"))
        hit_info = None
        fallback = None
        best_d2 = None
        tol_px = 14.0
        tol2 = tol_px * tol_px
        for info in markers:
            rect = info.get('rect')
            if rect is None:
                continue
            if rect.contains(x, y):
                hit_info = info
                break
            center = rect.center()
            try:
                dx = float(x - center.x())
                dy = float(y - center.y())
            except Exception:
                continue
            d2 = dx * dx + dy * dy
            if best_d2 is None or d2 < best_d2:
                best_d2 = d2
                fallback = info
        if hit_info is None and fallback is not None and best_d2 is not None and best_d2 <= tol2:
            hit_info = fallback
        if hit_info is None:
            return False
        if hit_info.get('label') == 'badge':
            self._open_spectro_summary_for_file(file_key)
            return True
        mods = QtCore.Qt.NoModifier
        if event is not None:
            try:
                mods = event.modifiers()
            except Exception:
                mods = QtCore.Qt.NoModifier
        spec = hit_info.get('spec')
        controller = getattr(self, "spectro_compare_controller", None)
        if controller:
            is_matrix = hit_info.get('kind') == 'matrix'
            if controller.handle_marker_click(spec, file_key, is_matrix, mods):
                return True
        return False

    def _handle_spec_hover(self, label_widget, event):
        if not self.show_spectra:
            QtWidgets.QToolTip.hideText()
            return False
        markers = label_widget.property("spec_markers") or []
        if not markers:
            QtWidgets.QToolTip.hideText()
            return False
        coords = self._label_pos_to_pix_coords(label_widget, event.pos())
        if coords is None:
            QtWidgets.QToolTip.hideText()
            return False
        x, y = coords
        for info in markers:
            rect = info.get('rect')
            if rect and rect.contains(x, y):
                if info.get('label') == 'badge':
                    QtWidgets.QToolTip.showText(label_widget.mapToGlobal(event.pos()), "Spectroscopy summary")
                    return True
                spec = info.get('spec') or {}
                tooltip = info.get('tooltip')
                if not tooltip:
                    tooltip = Path(spec.get('path', '')).name
                    idx = spec.get('matrix_index')
                    if idx is not None:
                        tooltip = f"{tooltip} [{idx}]"
                    xs = spec.get('x'); ys = spec.get('y')
                    if xs is not None and ys is not None:
                        tooltip = f"{tooltip}\n({xs:.1f}, {ys:.1f}) nm"
                    stack_summary = str(spec.get("xy_stack_summary") or "").strip()
                    if stack_summary:
                        tooltip = f"{tooltip}\n{stack_summary}"
                QtWidgets.QToolTip.showText(label_widget.mapToGlobal(event.pos()), tooltip)
                return True
        QtWidgets.QToolTip.hideText()
        return False

    def _open_spectroscopy_popup(self, spec):
        controller = getattr(self, "spectro_compare_controller", None)
        if controller:
            return controller.open_single_popup(spec)
        if not self._spectros_loaded:
            self.ensure_spectros_loaded(refresh=False)
        return spectro_popups._open_spectroscopy_popup(self, spec)

    def _ensure_single_spectro_popup(self, spec):
        controller = getattr(self, "spectro_compare_controller", None)
        if controller:
            return controller.ensure_single_popup(spec)
        if not spec:
            return None
        key = self._spec_identity_key(spec)
        if key and getattr(self, "_spectro_popups", None):
            for dlg in list(self._spectro_popups):
                dlg_spec = getattr(dlg, "spec", None)
                if dlg_spec and self._spec_identity_key(dlg_spec) == key:
                    try:
                        dlg.raise_()
                        dlg.activateWindow()
                    except Exception:
                        pass
                    return dlg
        return self._open_spectroscopy_popup(spec)

    def _append_spec_to_single_popup(self, spec):
        controller = getattr(self, "spectro_compare_controller", None)
        if controller:
            return controller.append_spec_to_single_popup(spec)

    def _prime_multi_selection_anchor(self, current_spec):
        controller = getattr(self, "spectro_compare_controller", None)
        if controller:
            return controller.prime_multi_selection_anchor(current_spec)

    def _highlight_spectrum_entry(self, spec):
        if not getattr(self, "spectro_highlight_glow", True):
            # ensure timer stopped and preview restored
            if self._highlight_timer.isActive():
                self._highlight_timer.stop()
            self._highlighted_spec = None
            self._highlight_phase = 0.0
            self._highlight_pulse_strength = 1.0
            self._schedule_marker_refresh()
            if hasattr(self, 'preview_canvas') and self.preview_canvas:
                try:
                    self.preview_canvas.update_highlight_pulse(1.0)
                except Exception:
                    pass
            return
        previous_spec = getattr(self, '_highlighted_spec', None)
        self._highlighted_spec = spec
        if spec:
            self._highlight_phase = 0.0
            if not self._highlight_timer.isActive():
                self._highlight_timer.start()
            try:
                shared_keys = [str(key) for key in (spec.get("shared_image_keys") or []) if key]
                current_preview_key = str(self.last_preview[0]) if self.last_preview else ""
                if current_preview_key and current_preview_key in shared_keys:
                    target_key = current_preview_key
                else:
                    target_key = str(spec.get('image_key') or spec.get('path') or '')
            except Exception:
                target_key = ''
            if self.last_preview and target_key and str(self.last_preview[0]) == target_key:
                self.show_file_channel(self.last_preview[0], self.last_preview[1])
            self._on_highlight_tick(force=True)
        else:
            if self._highlight_timer.isActive():
                self._highlight_timer.stop()
            self._highlight_phase = 0.0
            self._highlight_pulse_strength = 1.0
            self._schedule_marker_refresh()
            if hasattr(self, 'preview_canvas') and self.preview_canvas:
                try:
                    self.preview_canvas.update_highlight_pulse(1.0)
                except Exception:
                    pass
            prev_key = None
            if previous_spec:
                try:
                    prev_key = str(previous_spec.get('image_key') or previous_spec.get('path') or '')
                except Exception:
                    prev_key = None
            if prev_key and self.last_preview and str(self.last_preview[0]) == prev_key:
                self.show_file_channel(self.last_preview[0], self.last_preview[1])

    def _on_highlight_tick(self, force=False):
        if not self._highlighted_spec or not getattr(self, "spectro_highlight_glow", True):
            if self._highlight_timer.isActive():
                self._highlight_timer.stop()
            return
        if not force:
            self._highlight_phase = (self._highlight_phase + 0.35) % (2 * math.pi)
        pulse = 0.9 + 0.4 * (0.5 * (1.0 + math.sin(self._highlight_phase)))
        self._highlight_pulse_strength = pulse
        self._schedule_marker_refresh()
        try:
            if hasattr(self, 'preview_canvas') and self.preview_canvas:
                self.preview_canvas.update_highlight_pulse(pulse)
        except Exception:
            pass
    def _on_thumb_context_menu(self, label_widget, pos):
        fp = str(label_widget.property("file_path"))
        # If user has a multi-selection, operate on all of them (plus the clicked one)
        targets = sorted(set(self.thumb_multi_select or []) | {fp})
        menu = QtWidgets.QMenu(self)
        sub = menu.addMenu("Apply filter")
        for key, info in FILTER_DEFINITIONS.items():
            act = QtWidgets.QAction(self._filter_action_label(key), menu)
            if info.get('needs_gaussian') and not _gaussian_available():
                act.setEnabled(False)
                act.setToolTip("Requires scipy or OpenCV.")
            act.triggered.connect(lambda _, k=key, paths=list(targets), focus=fp: self._apply_filter_to_paths(paths, k, focus_path=focus))
            sub.addAction(act)
        custom_act = QtWidgets.QAction("Custom pipeline...", menu)
        custom_act.triggered.connect(lambda _, paths=list(targets), focus=fp: self._open_custom_filter_dialog(paths, focus))
        sub.addAction(custom_act)
        clear_one = QtWidgets.QAction("Clear filter", menu)
        clear_one.triggered.connect(lambda _, paths=[fp]: self._clear_filter_for_paths(paths))
        menu.addAction(clear_one)
        if len(targets) > 1:
            clear_sel = QtWidgets.QAction("Clear filter (selected)", menu)
            clear_sel.triggered.connect(lambda _, paths=list(targets): self._clear_filter_for_paths(paths))
            menu.addAction(clear_sel)
        menu.addSeparator()
        add_source_file_menu(menu, fp, self)

        menu.addSeparator()
        copy_svg_act = QtWidgets.QAction("Copy selected as SVG (current view)", menu)
        copy_svg_act.triggered.connect(lambda: self.copy_selected_as_svg(targets))
        menu.addAction(copy_svg_act)

        export_png_act = QtWidgets.QAction("Export PNGs...", menu)
        export_png_act.triggered.connect(self.on_export_pngs)
        menu.addAction(export_png_act)
        export_xyz_act = QtWidgets.QAction("Export XYZ...", menu)
        export_xyz_act.triggered.connect(self.on_export_xyz_files)
        menu.addAction(export_xyz_act)
        export_stp_act = QtWidgets.QAction("Export WSxM STP...", menu)
        export_stp_act.triggered.connect(self.on_export_stp_files)
        menu.addAction(export_stp_act)
        adjust_act = QtWidgets.QAction("Adjust image...", menu)
        adjust_act.triggered.connect(self.on_adjust_image)
        menu.addAction(adjust_act)

        # Virtual copies submenu
        virt_menu = menu.addMenu("Virtual copy")
        virt_cur = QtWidgets.QAction("Current channel", virt_menu)
        virt_cur.triggered.connect(lambda _, paths=list(targets): self._create_virtual_channel_copies(paths, self.channel_dropdown.currentIndex()))
        virt_menu.addAction(virt_cur)
        virt_other = QtWidgets.QAction("Choose channel...", virt_menu)
        virt_other.triggered.connect(lambda _, paths=list(targets): self._create_virtual_channel_copies(paths, None))
        virt_menu.addAction(virt_other)
        virt_menu.addSeparator()
        virt_remove = QtWidgets.QAction("Remove virtual copies (selected)", virt_menu)
        virt_remove.triggered.connect(lambda _, paths=list(targets): self._remove_virtual_entries(paths))
        virt_menu.addAction(virt_remove)

        mol_menu = menu.addMenu("Molecules")
        mol_clear = QtWidgets.QAction("Clear molecules (selected)", mol_menu)
        mol_clear.triggered.connect(lambda _, paths=list(targets): self._clear_molecules_for_paths(paths))
        mol_menu.addAction(mol_clear)
        mol_copy = QtWidgets.QAction("Copy molecules from source", mol_menu)
        mol_copy.setEnabled(any(self._is_processed_key(str(p)) for p in targets))
        mol_copy.triggered.connect(lambda _, paths=list(targets): self._copy_molecules_from_source(paths))
        mol_menu.addAction(mol_copy)

        if hasattr(self, '_clear_multi_spec_selection'):
            menu.addSeparator()
            clear_specs_act = QtWidgets.QAction("Clear spectroscopy selections", menu)
            clear_specs_act.triggered.connect(self._clear_multi_spec_selection)
            menu.addAction(clear_specs_act)

        menu.addSeparator()
        drift_act = QtWidgets.QAction("Drift-correct and export...", menu)
        drift_act.triggered.connect(lambda _, paths=list(targets): self._on_drift_correct(paths))
        menu.addAction(drift_act)
        anim_act = QtWidgets.QAction("Create animation from selection...", menu)
        anim_act.triggered.connect(lambda _, paths=list(targets): self._on_create_animation(paths))
        menu.addAction(anim_act)

        menu.exec_(label_widget.mapToGlobal(pos))

    def _on_spectro_thumb_context_menu(self, label_widget, pos):
        spec = label_widget.property("spectro_entry")
        if not spec:
            return
        key = str(spec.get("path", "") or "")
        selected = sorted(set(getattr(self, "spectro_thumb_multi_select", set()) or []))
        targets = selected if selected else [key]
        menu = QtWidgets.QMenu(self)
        open_act = QtWidgets.QAction("Open spectroscopy", menu)
        open_act.triggered.connect(lambda: self._open_spectroscopy_popup(spec))
        menu.addAction(open_act)
        details_act = QtWidgets.QAction("Show metadata in Details", menu)
        details_act.triggered.connect(lambda: self.show_spectroscopy_details(spec))
        menu.addAction(details_act)
        add_source_file_menu(menu, spec.get("path"), self)

        channels = list((spec.get("channels") or {}).keys())
        common = set(channels)
        if selected:
            for path in selected:
                entry = None
                for s in getattr(self, "spectros", []) or []:
                    if str(s.get("path", "") or "") == path:
                        entry = s
                        break
                if entry is None:
                    continue
                common &= set((entry.get("channels") or {}).keys())
        channel_menu = menu.addMenu("Miniature channel")
        channel_list = sorted(common) if common else channels
        if not channel_list:
            channel_list = channels
        if channel_list:
            current = self.spectro_thumb_channel_by_path.get(key) or self.spectro_miniature_default_channel or channel_list[0]
            for ch_name in channel_list:
                act = QtWidgets.QAction(ch_name, channel_menu)
                act.setCheckable(True)
                act.setChecked(ch_name == current)
                act.triggered.connect(lambda _checked, ch=ch_name, paths=list(targets): self.on_set_spectro_thumbnail_channel(ch, paths))
                channel_menu.addAction(act)

        menu.exec_(label_widget.mapToGlobal(pos))

    def _apply_filter_to_paths(self, paths, filter_key=None, pipeline=None, label=None, focus_path=None):
        if not paths:
            return
        if len(paths) > 12:
            ret = QtWidgets.QMessageBox.question(self, "Filters", f"Apply filter to {len(paths)} images? This may use significant memory.",
                                                 QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No, QtWidgets.QMessageBox.No)
            if ret != QtWidgets.QMessageBox.Yes:
                return
        if filter_key and FILTER_DEFINITIONS.get(filter_key, {}).get('needs_gaussian') and not _gaussian_available():
            QtWidgets.QMessageBox.warning(self, "Filters", "Gaussian filters require scipy or OpenCV.")
            return
        if pipeline is None:
            base_arr, preview_callback, original_views, preview_target, preview_cmap_name, preview_clim = self._filter_preview_context_for_path(focus_path)
            step, spec_label = self._single_filter_step_spec(
                filter_key,
                parent=self,
                base_image=base_arr,
                preview_callback=preview_callback,
                preview_target_text=preview_target,
                preview_cmap_name=preview_cmap_name,
                preview_clim=preview_clim,
            )
            if original_views is not None:
                self._restore_filter_views_on_canvas(self.preview_canvas, original_views)
            if step is None:
                return
            spec_steps = [step]
        else:
            spec_steps = pipeline
            spec_label = label or 'Custom'
        path_keys = {str(Path(p)) for p in paths}
        for key in path_keys:
            steps_copy = [dict(step) for step in spec_steps]
            self.thumbnail_filters[key] = {'steps': steps_copy, 'label': spec_label}
        self._invalidate_thumbnail_cache(path_keys)
        self._invalidate_filtered_cache(path_keys)
        self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())
        if self.last_preview and str(self.last_preview[0]) in path_keys:
            self.show_file_channel(self.last_preview[0], self.last_preview[1])

    def _clear_filter_for_paths(self, paths):
        changed = False
        path_keys = {str(Path(p)) for p in paths}
        for key in path_keys:
            if self.thumbnail_filters.pop(key, None) is not None:
                changed = True
        if changed:
            self._invalidate_thumbnail_cache(path_keys)
            self._invalidate_filtered_cache(path_keys)
            self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())
            if self.last_preview and str(self.last_preview[0]) in path_keys:
                self.show_file_channel(self.last_preview[0], self.last_preview[1])

    def _open_custom_filter_dialog(self, paths, focus_path):
        base_arr, preview_callback, original_views, preview_target, preview_cmap_name, preview_clim = self._filter_preview_context_for_path(focus_path)
        dlg = CustomFilterDialog(
            self,
            base_arr,
            self._run_filter_step,
            preview_callback=preview_callback,
            preview_target_text=preview_target,
            preview_cmap_name=preview_cmap_name,
            preview_clim=preview_clim,
        )
        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            pipeline = dlg.pipeline_steps()
            if pipeline:
                if original_views is not None:
                    self._restore_filter_views_on_canvas(self.preview_canvas, original_views)
                self._apply_filter_to_paths(paths, pipeline=pipeline, label=dlg.pipeline_label())
                return
        if original_views is not None:
            self._restore_filter_views_on_canvas(self.preview_canvas, original_views)

    def _toggle_thumb_multi_selection(self, file_path):
        return viewer_thumb_ui._toggle_thumb_multi_selection(self, file_path)

    def _clear_thumb_multi_selection(self, update_styles=True):
        return viewer_thumb_ui._clear_thumb_multi_selection(self, update_styles=update_styles)

    def _spec_identity_key(self, spec):
        controller = getattr(self, "spectro_compare_controller", None)
        if controller:
            return controller.spec_identity_key(spec)
        if not spec:
            return None
        base = spec.get('path')
        try:
            base = str(Path(base))
        except Exception:
            base = str(base)
        idx = spec.get('matrix_index')
        if idx is not None:
            return f"{base}#idx{idx}"
        x = spec.get('x')
        y = spec.get('y')
        if x is not None or y is not None:
            try:
                x_val = float(x) if x is not None else ''
                y_val = float(y) if y is not None else ''
                return f"{base}#pos{round(x_val,6)}_{round(y_val,6)}"
            except Exception:
                return f"{base}#pos{x}_{y}"
        order_idx = spec.get('order_idx')
        if order_idx is not None:
            return f"{base}#order{order_idx}"
        return base

    def _toggle_multi_spec_selection(self, spec):
        controller = getattr(self, "spectro_compare_controller", None)
        if controller:
            return controller.toggle_multi_spec_selection(spec)

    def _update_spec_selection_label(self):
        controller = getattr(self, "spectro_compare_controller", None)
        if controller:
            return controller.update_spec_selection_label()

    def _clear_multi_spec_selection(self):
        controller = getattr(self, "spectro_compare_controller", None)
        if controller:
            return controller.clear_multi_spec_selection()

    def _on_drift_correct(self, paths):
        if not paths:
            return
        try:
            from scipy import ndimage  # type: ignore
        except Exception:
            QtWidgets.QMessageBox.warning(self, "Drift correction", "scipy is required for alignment interpolation.")
            return
        # Prefer skimage phase_cross_correlation; fall back to OpenCV ECC; else zeros
        try:
            from skimage.registration import phase_cross_correlation  # type: ignore
        except Exception:
            phase_cross_correlation = None  # type: ignore
        try:
            import cv2  # type: ignore
            has_cv = True
        except Exception:
            has_cv = False
        channel_idx = self.channel_dropdown.currentIndex()
        images = []
        names_full = []
        names_display = []
        missing = 0
        # Preserve user selection order while dropping duplicates
        seen = set()
        ordered_paths = []
        for p in paths:
            ps = str(Path(p))
            if ps in seen:
                continue
            seen.add(ps)
            ordered_paths.append(ps)
        for p in ordered_paths:
            try:
                header, fds = self.headers.get(p, (None, None))
                if header is None or fds is None:
                    header, fds = parse_header(Path(p))
                if not fds:
                    continue
                # Prefer current channel, but fall back to any available channel with data
                indices = [channel_idx] + [i for i in range(len(fds)) if i != channel_idx]
                arr = None
                for idx in indices:
                    if idx < 0 or idx >= len(fds):
                        continue
                    try:
                        arr = self._get_channel_array(p, idx, header, fds[idx])
                    except Exception:
                        arr = None
                    if arr is not None:
                        break
                if arr is None:
                    missing += 1
                    continue
                names_full.append(str(p))
                names_display.append(Path(p).stem)
                images.append(np.array(arr, dtype=float))
            except Exception:
                missing += 1
                continue
        if len(images) < 2:
            QtWidgets.QMessageBox.information(
                self,
                "Drift correction",
                f"Need at least two images to align.\nLoaded: {len(images)} / Selected: {len(set(paths))}\n"
                f"Skipped/missing: {missing}",
            )
            return
        # Align relative to the first frame
        ref_idx = 0
        reference = images[ref_idx]
        shifts = np.zeros((len(images), 2), dtype=float)
        ref_gray = reference.astype(np.float32)
        ref_gray = (ref_gray - ref_gray.min()) / max(ref_gray.max() - ref_gray.min(), 1e-6)
        # apply a Hann window to reduce edge effects
        try:
            win_y = np.hanning(ref_gray.shape[0])
            win_x = np.hanning(ref_gray.shape[1])
            window = np.sqrt(np.outer(win_y, win_x))
            ref_gray *= window
        except Exception:
            pass
        for i, img in enumerate(images):
            if i == ref_idx:
                continue
            target = img.astype(np.float32)
            target = (target - target.min()) / max(target.max() - target.min(), 1e-6)
            try:
                target *= window
            except Exception:
                pass
            try:
                if phase_cross_correlation is not None:
                    shift, _, _ = phase_cross_correlation(ref_gray, target, upsample_factor=20, normalization="phase")
                    shifts[i] = [float(shift[0]), float(shift[1])]  # dy, dx mapping target -> ref
                elif has_cv:
                    import cv2  # type: ignore
                    criteria = (cv2.TERM_CRITERIA_EPS | cv2.TERM_CRITERIA_COUNT, 100, 1e-6)
                    warp_matrix = np.eye(2, 3, dtype=np.float32)
                    _, warp_matrix = cv2.findTransformECC(ref_gray, target, warp_matrix, cv2.MOTION_TRANSLATION, criteria)
                    shifts[i] = [warp_matrix[1, 2], warp_matrix[0, 2]]  # dy, dx that map target -> ref
                else:
                    shifts[i] = [0.0, 0.0]
            except Exception:
                shifts[i] = [0.0, 0.0]
        
        H, W = images[0].shape[:2]
        
        # Calculate the intersection crop region after alignment
        # After shifting image i by (dy, dx), its valid region is constrained
        top = int(np.ceil(max(0, np.max(shifts[:, 0]))))
        bottom = int(np.floor(min(H, H + np.min(shifts[:, 0]))))
        left = int(np.ceil(max(0, np.max(shifts[:, 1]))))
        right = int(np.floor(min(W, W + np.min(shifts[:, 1]))))
        
        # Ensure valid bounds
        top = max(0, min(top, H - 1))
        left = max(0, min(left, W - 1))
        bottom = max(top + 1, min(bottom, H))
        right = max(left + 1, min(right, W))
        
        # REMOVED: Square enforcement logic that was causing severe overcropping
        # The intersection crop is sufficient - no need to force square dimensions
        
        aligned = []
        for img, shift in zip(images, shifts):
            dy, dx = shift
            try:
                # FIXED: Apply shift directly (removed negation)
                warped = ndimage.shift(img, [dy, dx], order=3, mode="reflect", cval=0.0)
            except Exception:
                warped = img
            aligned.append(warped[top:bottom, left:right])
        
        self._show_alignment_preview(names_display, aligned, shifts, channel_idx, names_full, crop_bounds=(top, bottom, left, right))

    def _on_create_animation(self, paths):
        if not paths:
            return
        try:
            import imageio.v3 as iio  # type: ignore
        except Exception:
            QtWidgets.QMessageBox.warning(self, "Animation", "imageio is required to create GIF/MP4 animations.")
            return

        def _resize_frame(arr: np.ndarray, target_w: int, target_h: int) -> np.ndarray:
            if arr.shape[0] == target_h and arr.shape[1] == target_w:
                return arr
            # Prefer Pillow if available
            try:
                from PIL import Image
                mode = "L" if arr.ndim == 2 else "RGB"
                im = Image.fromarray(arr, mode=mode if mode else None)
                im = im.resize((target_w, target_h), Image.BILINEAR)
                return np.array(im)
            except Exception:
                pass
            # Fallback to scipy if present
            try:
                from scipy import ndimage as _ndi  # type: ignore
                zoom = (target_h / arr.shape[0], target_w / arr.shape[1]) + (() if arr.ndim == 2 else (1,))
                return _ndi.zoom(arr, zoom, order=1)
            except Exception:
                pass
            # Last resort: nearest-neighbor using numpy repeat
            y_idx = np.linspace(0, arr.shape[0] - 1, target_h).astype(int)
            x_idx = np.linspace(0, arr.shape[1] - 1, target_w).astype(int)
            if arr.ndim == 2:
                return arr[np.ix_(y_idx, x_idx)]
            else:
                return arr[np.ix_(y_idx, x_idx, np.arange(arr.shape[2]))]

        # Build a rich export dialog with preview and options
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("Export animation")
        dlg.resize(800, 640)
        vbox = QtWidgets.QVBoxLayout(dlg); vbox.setContentsMargins(10, 10, 10, 10); vbox.setSpacing(8)

        # Gather frames
        channel_idx = self.channel_dropdown.currentIndex()
        frames = []
        missing = 0
        names = []
        for p in sorted({str(Path(p)) for p in paths}):
            try:
                header, fds = self.headers.get(p, (None, None))
                if header is None or fds is None:
                    header, fds = parse_header(Path(p))
                if not fds:
                    continue
                indices = [channel_idx] + [i for i in range(len(fds)) if i != channel_idx]
                arr = None
                for idx in indices:
                    if idx < 0 or idx >= len(fds):
                        continue
                    try:
                        arr = self._get_channel_array(p, idx, header, fds[idx])
                    except Exception:
                        arr = None
                    if arr is not None:
                        break
                if arr is None:
                    missing += 1
                    continue
                frames.append(np.array(arr, dtype=float))
                names.append(Path(p).name)
            except Exception:
                missing += 1
                continue
        if not frames:
            QtWidgets.QMessageBox.information(
                self,
                "Animation",
                f"No frames could be loaded. Selected: {len(set(paths))}, skipped: {missing}",
            )
            return

        # Controls row
        controls = QtWidgets.QHBoxLayout(); controls.setSpacing(12)
        controls.addWidget(QtWidgets.QLabel("Format:"))
        fmt_combo = QtWidgets.QComboBox(); fmt_combo.addItems(["gif", "mp4", "png-seq"]); controls.addWidget(fmt_combo)
        controls.addWidget(QtWidgets.QLabel("FPS:"))
        fps_spin = QtWidgets.QSpinBox(); fps_spin.setRange(1, 60); fps_spin.setValue(6); controls.addWidget(fps_spin)
        controls.addWidget(QtWidgets.QLabel("Duration (s):"))
        dur_spin = QtWidgets.QDoubleSpinBox(); dur_spin.setRange(0.1, 120.0); dur_spin.setDecimals(1); dur_spin.setSingleStep(0.5); controls.addWidget(dur_spin)
        dur_spin.setValue(max(0.1, len(frames) / fps_spin.value()))
        def _update_duration():
            dur_spin.setValue(max(0.1, len(frames) / max(1, fps_spin.value())))
        fps_spin.valueChanged.connect(_update_duration)
        controls.addStretch(1)
        vbox.addLayout(controls)

        # Overlay toggles
        overlay_row = QtWidgets.QHBoxLayout(); overlay_row.setSpacing(12)
        scale_cb = QtWidgets.QCheckBox("Include scale bar"); scale_cb.setChecked(True)
        markers_cb = QtWidgets.QCheckBox("Include markers/overlays"); markers_cb.setChecked(True)
        mol_cb = QtWidgets.QCheckBox("Include molecules"); mol_cb.setChecked(True)
        overlay_row.addWidget(scale_cb); overlay_row.addWidget(markers_cb); overlay_row.addWidget(mol_cb); overlay_row.addStretch(1)
        vbox.addLayout(overlay_row)

        # Resolution
        res_row = QtWidgets.QHBoxLayout(); res_row.setSpacing(12)
        res_row.addWidget(QtWidgets.QLabel("Resolution:"))
        res_combo = QtWidgets.QComboBox(); res_combo.addItems(["Auto", "720p", "1080p", "Custom"]); res_row.addWidget(res_combo)
        w_spin = QtWidgets.QSpinBox(); w_spin.setRange(256, 4096); w_spin.setValue(frames[0].shape[1]); res_row.addWidget(QtWidgets.QLabel("W")); res_row.addWidget(w_spin)
        h_spin = QtWidgets.QSpinBox(); h_spin.setRange(256, 4096); h_spin.setValue(frames[0].shape[0]); res_row.addWidget(QtWidgets.QLabel("H")); res_row.addWidget(h_spin)
        def _on_res_change(text):
            presets = {"720p": (1280, 720), "1080p": (1920, 1080)}
            if text in presets:
                w_spin.setValue(presets[text][0]); h_spin.setValue(presets[text][1])
            elif text == "Auto":
                w_spin.setValue(frames[0].shape[1]); h_spin.setValue(frames[0].shape[0])
        res_combo.currentTextChanged.connect(_on_res_change)
        vbox.addLayout(res_row)

        # Preview canvas
        prev_label = QtWidgets.QLabel(); prev_label.setAlignment(QtCore.Qt.AlignCenter)
        prev_label.setMinimumHeight(260)
        vbox.addWidget(prev_label, 1)

        # Buttons
        btn_row = QtWidgets.QHBoxLayout(); btn_row.addStretch(1)
        save_btn = QtWidgets.QPushButton("Save…"); cancel_btn = QtWidgets.QPushButton("Cancel")
        btn_row.addWidget(save_btn); btn_row.addWidget(cancel_btn)
        vbox.addLayout(btn_row)

        def _render_frame(arr):
            # simple normalization for preview
            a = np.asarray(arr, dtype=float)
            rng = a.max() - a.min()
            if rng <= 0:
                norm = np.zeros_like(a, dtype=np.uint8)
            else:
                norm = ((a - a.min()) / rng * 255.0).clip(0, 255).astype(np.uint8)
            h, w = norm.shape[:2]
            if norm.ndim == 2:
                rgb = np.stack([norm]*3, axis=-1)
            else:
                rgb = norm
            qimg = QtGui.QImage(rgb.data, w, h, 3*w, QtGui.QImage.Format_RGB888)
            pm = QtGui.QPixmap.fromImage(qimg.copy())
            return pm

        def _fit_resize_with_pad(rgb: np.ndarray, tw: int, th: int, fill=255):
            rgb = np.asarray(rgb)
            if rgb.ndim == 2:
                rgb = np.stack([rgb]*3, axis=-1)
            h, w = rgb.shape[:2]
            if h == 0 or w == 0:
                return np.zeros((th, tw, 3), dtype=np.uint8)
            scale = min(tw / w, th / h)
            new_w = max(1, int(round(w * scale)))
            new_h = max(1, int(round(h * scale)))
            resized = _resize_frame(rgb, new_w, new_h)
            canvas = np.full((th, tw, 3), fill, dtype=np.uint8)
            off_x = (tw - new_w) // 2
            off_y = (th - new_h) // 2
            canvas[off_y:off_y+new_h, off_x:off_x+new_w, :] = resized if resized.ndim == 3 else np.stack([resized]*3, axis=-1)
            return canvas

        def _update_preview(idx=0):
            pm = _render_frame(frames[idx % len(frames)])
            if not pm.isNull():
                pm = pm.scaled(prev_label.width(), prev_label.height(), QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation)
                prev_label.setPixmap(pm)
        _update_preview(0)

        # Auto-play preview timer
        timer = QtCore.QTimer(dlg); timer.setInterval(400)
        idx_ref = {"i": 0}
        def _tick():
            idx_ref["i"] = (idx_ref["i"] + 1) % len(frames)
            _update_preview(idx_ref["i"])
        timer.timeout.connect(_tick); timer.start()

        def _save():
            fmt = fmt_combo.currentText()
            default_name = f"animation.{ 'gif' if fmt=='gif' else ('mp4' if fmt=='mp4' else 'png') }"
            filter_str = "GIF (*.gif);;MP4 (*.mp4);;PNG sequence (*.png)"
            out_path, _ = QtWidgets.QFileDialog.getSaveFileName(dlg, "Save animation", default_name, filter_str)
            if not out_path:
                return
            # render each frame via the preview canvas so overlays (scale bar) are honored
            target_w, target_h = w_spin.value(), h_spin.value()
            canvas = getattr(self, "preview_canvas", None)
            orig_last = getattr(self, "last_preview", None)
            orig_views = list(getattr(canvas, "views", [])) if canvas else []
            norm_frames = []

            def _render_path(path_str: str):
                if not canvas:
                    return None
                try:
                    # toggle scale bar; drop molecules/markers for clean export
                    prev_scale = canvas.scale_bar_enabled
                    prev_ticks = canvas._show_ticks
                    prev_mols = list(getattr(canvas, "molecules", []))
                    canvas.scale_bar_enabled = scale_cb.isChecked()
                    canvas._show_ticks = prev_ticks  # keep ticks as-is
                    canvas.molecules = []  # always exclude molecules for animation per requirement
                    # show file/channel and draw a fresh figure (with scale bar)
                    self.show_file_channel(path_str, channel_idx)
                    QtWidgets.QApplication.processEvents(QtCore.QEventLoop.AllEvents, 50)
                    if not getattr(canvas, "views", None):
                        return None
                    view = canvas.views[0]
                    fig = canvas._render_view_figure(view)
                    fig.set_dpi(100)
                    fig.set_size_inches(target_w / 100.0, target_h / 100.0)
                    fig.canvas.draw()
                    buf = fig.canvas.buffer_rgba()
                    if buf is None:
                        return None
                    arr = np.asarray(buf)
                    rgb = arr[:, :, :3].copy()
                    rgb = _fit_resize_with_pad(rgb, target_w, target_h, fill=255)
                    try:
                        import matplotlib.pyplot as _plt  # type: ignore
                        _plt.close(fig)
                    except Exception:
                        pass
                    # restore
                    canvas.scale_bar_enabled = prev_scale
                    canvas._show_ticks = prev_ticks
                    canvas.molecules = prev_mols
                    return rgb
                except Exception:
                    return None

            for p in sorted({str(Path(p)) for p in paths}):
                rendered = _render_path(p)
                if rendered is not None:
                    norm_frames.append(rendered)

            # fallback to raw frames if rendering failed
            if not norm_frames:
                for arr in frames:
                    a = np.asarray(arr, dtype=float)
                    rng = a.max() - a.min()
                    if rng <= 0:
                        norm = np.zeros_like(a, dtype=np.uint8)
                    else:
                        norm = ((a - a.min()) / rng * 255.0).clip(0, 255).astype(np.uint8)
                    if norm.shape[0] != target_h or norm.shape[1] != target_w:
                        norm = _resize_frame(norm, target_w, target_h)
                    norm_frames.append(norm)

            # restore original view
            try:
                if orig_last:
                    self.show_file_channel(orig_last[0], orig_last[1])
                elif canvas and orig_views:
                    canvas.views = orig_views
                    canvas._redraw()
            except Exception:
                pass
            try:
                if fmt == "gif":
                    iio.imwrite(out_path, norm_frames, plugin="pillow", loop=0, duration=1000.0 / max(1, fps_spin.value()))
                elif fmt == "mp4":
                    iio.imwrite(out_path, norm_frames, plugin="ffmpeg", fps=max(1, fps_spin.value()))
                else:
                    stem = Path(out_path).with_suffix("")
                    for i, fr in enumerate(norm_frames):
                        iio.imwrite(f"{stem}_{i:03d}.png", fr)
                QtWidgets.QMessageBox.information(dlg, "Animation", f"Saved animation to {out_path}")
                dlg.accept()
            except Exception as exc:
                QtWidgets.QMessageBox.warning(dlg, "Animation", f"Failed to save animation: {exc}")

        save_btn.clicked.connect(_save)
        cancel_btn.clicked.connect(dlg.reject)
        dlg.exec_()

    def _show_alignment_preview(self, names, aligned, shifts, channel_idx, source_paths=None, crop_bounds=None):
        """Preview aligned/cropped images and optionally save outputs/animation."""
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("Drift correction preview")
        dlg.resize(900, 720)
        layout = QtWidgets.QVBoxLayout(dlg)

        info = QtWidgets.QPlainTextEdit()
        info.setReadOnly(True)
        info.setMaximumHeight(140)
        text_lines = []
        max_shift = 0.0
        for name, shift in zip(names, shifts):
            mag = float(np.hypot(shift[0], shift[1]))
            max_shift = max(max_shift, mag)
            text_lines.append(f"{name}: dy={shift[0]:.3f} px, dx={shift[1]:.3f} px | |d|={mag:.3f} px")
        if crop_bounds:
            top, bottom, left, right = crop_bounds
            crop_h = max(0, bottom - top)
            crop_w = max(0, right - left)
            text_lines.append(f"\nCrop: top={top}, bottom={bottom}, left={left}, right={right}  -> size={crop_w}x{crop_h}px")
        text_lines.append(f"Max shift magnitude: {max_shift:.3f} px")
        info.setPlainText("\n".join(text_lines))
        layout.addWidget(info)

        # Controls row: cmap + speed slider
        controls = QtWidgets.QHBoxLayout()
        controls.addWidget(QtWidgets.QLabel("Colormap:"))
        cmap_combo = QtWidgets.QComboBox()
        try:
            cmap_combo.addItems(sorted(colormaps.keys()))
        except Exception:
            cmap_combo.addItems(["gray", "viridis", "plasma", "magma", "cividis"])
        if hasattr(self, "thumb_cmap"):
            idx = cmap_combo.findText(self.thumb_cmap)
            if idx >= 0:
                cmap_combo.setCurrentIndex(idx)
        controls.addWidget(cmap_combo)
        controls.addSpacing(12)
        controls.addWidget(QtWidgets.QLabel("Speed (fps):"))
        speed_slider = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        speed_slider.setRange(1, 30)
        speed_slider.setValue(6)
        controls.addWidget(speed_slider)
        controls.addStretch(1)
        layout.addLayout(controls)

        preview_label = QtWidgets.QLabel("")
        preview_label.setAlignment(QtCore.Qt.AlignCenter)
        preview_label.setMinimumHeight(320)
        layout.addWidget(preview_label, 1)

        btn_row = QtWidgets.QHBoxLayout()
        save_imgs_btn = QtWidgets.QPushButton("Save aligned PNGs...")
        save_gif_btn = QtWidgets.QPushButton("Save animation...")
        save_virtual_btn = QtWidgets.QPushButton("Save corrected copies to thumbnails")
        btn_row.addWidget(save_imgs_btn)
        btn_row.addWidget(save_gif_btn)
        btn_row.addWidget(save_virtual_btn)
        btn_row.addStretch(1)
        layout.addLayout(btn_row)

        # Build simple QTimer-based preview using RGB frames to avoid GIF issues
        preview_timer = QtCore.QTimer(dlg)
        preview_timer.setSingleShot(False)
        frames_rgb = []

        def _build_frames(cmap_name):
            nonlocal frames_rgb
            frames_rgb = []
            try:
                import matplotlib.cm as mcm
            except Exception:
                mcm = None
            cmap_lookup = getattr(mcm, "cmap_d", None)
            cmap = None
            if mcm:
                try:
                    if (cmap_lookup and cmap_name in cmap_lookup) or hasattr(mcm, "get_cmap"):
                        cmap = mcm.get_cmap(cmap_name)
                except Exception:
                    cmap = None
            for arr in aligned:
                arr = np.asarray(arr, dtype=float)
                rng = arr.max() - arr.min()
                if rng <= 0:
                    base = np.zeros_like(arr, dtype=float)
                else:
                    base = (arr - arr.min()) / rng
                if cmap is not None:
                    rgb = (cmap(base)[:, :, :3] * 255.0).astype(np.uint8)
                else:
                    rgb = np.repeat((base * 255.0).astype(np.uint8)[..., None], 3, axis=2)
                frames_rgb.append(rgb)

        def _update_preview():
            if not frames_rgb:
                preview_label.setText("Preview unavailable")
                return
            idx = (preview_timer.property("frame_idx") or 0) % len(frames_rgb)
            frame = frames_rgb[idx]
            h, w, _ = frame.shape
            qimg = QtGui.QImage(frame.data, w, h, 3 * w, QtGui.QImage.Format_RGB888)
            preview_label.setPixmap(QtGui.QPixmap.fromImage(qimg))
            preview_timer.setProperty("frame_idx", (idx + 1) % len(frames_rgb))

        def _render_preview(cmap_name, fps):
            _build_frames(cmap_name)
            interval = max(30, int(1000 / max(1, fps)))
            preview_timer.setInterval(interval)
            preview_timer.setProperty("frame_idx", 0)
            preview_timer.start()
            _update_preview()

        _render_preview(cmap_combo.currentText(), speed_slider.value())
        cmap_combo.currentTextChanged.connect(lambda name: _render_preview(name, speed_slider.value()))
        speed_slider.valueChanged.connect(lambda val: _render_preview(cmap_combo.currentText(), val))

        def _save_imgs():
            out_dir = QtWidgets.QFileDialog.getExistingDirectory(self, "Select output folder")
            if not out_dir:
                return
            out_dir = Path(out_dir)
            out_dir.mkdir(parents=True, exist_ok=True)
            for name, arr in zip(names, aligned):
                out_path = out_dir / f"{name}_aligned.png"
                try:
                    import imageio.v3 as iio  # type: ignore
                    iio.imwrite(out_path, arr.astype(np.float32))
                except Exception:
                    try:
                        from matplotlib import pyplot as plt  # type: ignore
                        plt.imsave(out_path, arr, cmap=cmap_combo.currentText())
                    except Exception:
                        np.savetxt(out_path.with_suffix(".txt"), arr)
            QtWidgets.QMessageBox.information(self, "Drift correction", f"Saved aligned images to {out_dir}")

        def _save_anim():
            try:
                import imageio.v3 as iio  # type: ignore
            except Exception:
                QtWidgets.QMessageBox.warning(dlg, "Animation", "imageio is required to save animations.")
                return
            out_path, _ = QtWidgets.QFileDialog.getSaveFileName(dlg, "Save animation", "aligned.gif", "GIF (*.gif);;MP4 (*.mp4)")
            if not out_path:
                return
            try:
                import matplotlib.cm as mcm
            except Exception:
                mcm = None
            cmap_lookup = getattr(mcm, "cmap_d", None)
            frames_out = []
            for arr in aligned:
                arr = np.asarray(arr, dtype=float)
                rng = arr.max() - arr.min()
                if rng <= 0:
                    base = np.zeros_like(arr, dtype=float)
                else:
                    base = (arr - arr.min()) / rng
                if mcm and ((cmap_lookup and cmap_combo.currentText() in cmap_lookup) or hasattr(mcm, "get_cmap")):
                    cmap = mcm.get_cmap(cmap_combo.currentText())
                    frames_out.append((cmap(base)[:, :, :3] * 255.0).astype(np.uint8))
                else:
                    frames_out.append((base * 255.0).astype(np.uint8))
            suffix = Path(out_path).suffix.lower()
            fps = max(1, speed_slider.value())
            try:
                if suffix == ".mp4":
                    iio.imwrite(out_path, frames_out, plugin="ffmpeg", fps=fps)
                else:
                    iio.imwrite(out_path, frames_out, plugin="pillow", loop=0, duration=max(20, int(1000 / fps)))
            except Exception as exc:
                QtWidgets.QMessageBox.warning(dlg, "Animation", f"Failed to save animation: {exc}")
                return
            QtWidgets.QMessageBox.information(dlg, "Animation", f"Saved animation to {out_path}")

        def _save_virtual():
            added = 0
            existing = set(str(p) for p in self.files)
            top, bottom, left, right = crop_bounds if crop_bounds else (None, None, None, None)
            for idx, arr in enumerate(aligned):
                orig = str(source_paths[idx] if source_paths and idx < len(source_paths) else names[idx])
                header_fds = self.headers.get(orig)
                if not header_fds:
                    continue
                header, fds = header_fds
                if not fds:
                    continue
                fd_src = fds[channel_idx if 0 <= channel_idx < len(fds) else 0]
                header_new = dict(header)
                header_new['xPixel'] = arr.shape[1]
                header_new['yPixel'] = arr.shape[0]
                arr_by_channel = {}
                try:
                    from scipy import ndimage as _ndi  # type: ignore
                except Exception:
                    _ndi = None
                dy, dx = shifts[idx]
                for ch_idx, fd_ch in enumerate(fds):
                    try:
                        if ch_idx == channel_idx:
                            raw_arr = arr  # already aligned/cropped for the primary channel
                        else:
                            raw_arr = self._get_channel_array(orig, ch_idx, header, fd_ch)
                    except Exception:
                        continue
                    try:
                        if ch_idx == channel_idx:
                            shifted = raw_arr
                        elif _ndi is not None:
                            shifted = _ndi.shift(raw_arr, [-dy, -dx], order=1, mode="reflect", cval=0.0)
                        else:
                            shifted = raw_arr
                        if all(v is not None for v in (top, bottom, left, right)):
                            shifted = shifted[top:bottom, left:right]
                        arr_by_channel[ch_idx] = np.array(shifted, copy=True)
                    except Exception:
                        continue
                if not arr_by_channel:
                    continue
                # adjust header dims to cropped size
                sample_arr = next(iter(arr_by_channel.values()))
                header_new['xPixel'] = sample_arr.shape[1]
                header_new['yPixel'] = sample_arr.shape[0]
                fds_new = [dict(fd) for fd in fds]
                caption_base = fds[channel_idx].get('Caption') or Path(orig).name if 0 <= channel_idx < len(fds) else Path(orig).name
                for i, fd_new in enumerate(fds_new):
                    fd_new['FileName'] = f"{Path(orig).name}_drift_ch{i}"
                    fd_new['Caption'] = f"{caption_base} [drift]"
                processed_key = f"processed_{Path(orig).stem}_drift"
                self._processed_views[processed_key] = {
                    'arr_by_channel': arr_by_channel,
                    'header': header_new,
                    'fds': fds_new,
                    'channel_idx': channel_idx,
                    'source': orig,
                }
                self.headers[processed_key] = (header_new, fds_new)
                if processed_key not in existing:
                    self.files.append(Path(processed_key))
                    existing.add(processed_key)
                added += 1
            if added:
                try:
                    # insert new processed entries right after the last selected source in current ordering
                    cur_files = [str(p) for p in self.files]
                    inserted = []
                    for idx, src in enumerate(source_paths or []):
                        src_str = str(src)
                        pk = f"processed_{Path(src_str).stem}_drift"
                        try:
                            pos = cur_files.index(src_str)
                        except ValueError:
                            pos = len(self.files)
                        if pk not in cur_files:
                            self.files.insert(pos + 1, Path(pk))
                            cur_files.insert(pos + 1, pk)
                            inserted.append(pk)
                    if not inserted:
                        for pk in list(self._processed_views.keys()):
                            if pk not in cur_files:
                                self.files.append(Path(pk))
                    self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())
                except Exception:
                    pass
                QtWidgets.QMessageBox.information(dlg, "Drift correction", f"Added {added} drift-corrected copy(ies) to thumbnails.\nLook for entries tagged [drift].")
            else:
                QtWidgets.QMessageBox.information(dlg, "Drift correction", "No corrected copies were created (missing headers or channels).")

        save_imgs_btn.clicked.connect(_save_imgs)
        save_gif_btn.clicked.connect(_save_anim)
        save_virtual_btn.clicked.connect(_save_virtual)
        dlg.exec_()

    # ---------- Virtual copies (channels, crops, drift) ----------

    def _virtual_copy_source_anchor(self, view):
        if not view:
            return VIRTUAL_COPY_INSERT_START
        path = view.get("path") or (view.get("meta") or {}).get("path") or (view.get("meta") or {}).get("file_path")
        return str(path) if path else VIRTUAL_COPY_INSERT_START

    def _create_virtual_copy_from_popup_view(self, view):
        return self._create_virtual_view_copy(view, insert_after_key=self._virtual_copy_source_anchor(view))

    def _create_virtual_copy_from_drag_payload(self, payload, insert_after_key=None):
        if not isinstance(payload, dict):
            return None
        drag_token = payload.get("view_drag_token")
        if drag_token:
            view = MultiPreviewCanvas.consume_drag_view_snapshot(drag_token)
            if view:
                return self._create_virtual_view_copy(view, insert_after_key=insert_after_key)
        file_path = payload.get("file_path")
        channel_idx = payload.get("channel_index")
        if not file_path or channel_idx is None:
            return None
        try:
            channel_idx = int(channel_idx)
        except Exception:
            return None
        created = self._create_virtual_channel_copies(
            [str(file_path)],
            channel_idx=channel_idx,
            insert_after_key=insert_after_key,
        )
        return created

    def _create_virtual_channel_copies(self, paths, channel_idx=None, insert_after_key=None):
        """Create virtual copies of selected images for a specific channel."""
        if not paths:
            return 0
        targets = [str(Path(p)) for p in paths]
        # If channel not provided, ask the user using first file's channels
        if channel_idx is None:
            first = targets[0]
            header, fds = self.headers.get(first, (None, None))
            if header is None or fds is None:
                header, fds = parse_header(Path(first))
            if not fds:
                return
            dlg = QtWidgets.QInputDialog(self)
            dlg.setInputMode(QtWidgets.QInputDialog.IntInput)
            dlg.setIntRange(0, len(fds) - 1)
            dlg.setIntValue(self.channel_dropdown.currentIndex())
            dlg.setLabelText(f"Channel index (0-{len(fds)-1}) for virtual copy")
            if dlg.exec_() != QtWidgets.QDialog.Accepted:
                return 0
            channel_idx = dlg.intValue()
        added = 0
        anchor_key = insert_after_key
        for p in targets:
            try:
                header, fds = self.headers.get(p, (None, None))
                if header is None or fds is None:
                    header, fds = parse_header(Path(p))
                if not fds or channel_idx < 0 or channel_idx >= len(fds):
                    continue
                # Build arrays for all channels so switching works
                arr_by_channel = {}
                for ch_idx, fd in enumerate(fds):
                    try:
                        arr_by_channel[ch_idx] = np.array(self._get_channel_array(p, ch_idx, header, fd), copy=True)
                    except Exception:
                        continue
                if not arr_by_channel:
                    continue
                fds_new = [dict(fd) for fd in fds]
                for i, fd_new in enumerate(fds_new):
                    fd_new['FileName'] = f"{Path(p).name}_virt_ch{i}"
                    fd_new['Caption'] = f"{fd_new.get('Caption') or Path(p).name} [ch{i}]"
                key = self._make_processed_key(p, op="ch", channel_idx=channel_idx)
                self._processed_views[key] = {
                    'arr_by_channel': arr_by_channel,
                    'header': dict(header),
                    'fds': fds_new,
                    'channel_idx': channel_idx,
                    'source': p,
                    'label': f"[ch{channel_idx}]",
                    'op': 'channel',
                }
                self.headers[key] = (dict(header), fds_new)
                self._insert_processed_after_source(key, p, insert_after_key=anchor_key)
                if insert_after_key not in (None, "", VIRTUAL_COPY_INSERT_START):
                    anchor_key = key
                added += 1
            except Exception:
                continue
        if added:
            self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())
        return added

    def _create_virtual_view_copy(self, view, insert_after_key=None, tag=None, op=None):
        """Create a virtual thumbnail copy from the current popup/preview view snapshot."""
        if not view:
            return None
        path = view.get("path") or (view.get("meta") or {}).get("path") or (view.get("meta") or {}).get("file_path")
        arr = view.get("arr")
        ch_idx = view.get("channel_idx")
        if ch_idx is None:
            ch_idx = (view.get("meta") or {}).get("channel_index")
        if path is None or arr is None:
            return None
        try:
            arr = np.asarray(arr)
        except Exception:
            return None
        if arr.ndim < 2 or arr.size == 0:
            return None
        path = str(path)
        try:
            header, fds = self.headers.get(path, (None, None))
            if header is None or fds is None:
                header, fds = parse_header(Path(path))
            if not fds:
                return None
            ch_idx = int(ch_idx) if ch_idx is not None else 0
            title = str(view.get("title") or "")
            inferred_crop = bool(view.get("crop_sequence") is not None or "[crop]" in title.lower())
            tag = str(tag or ("[crop]" if inferred_crop else "[copy]"))
            op_name = str(op or ("crop" if inferred_crop else "copy"))
            arr_by_channel = {ch_idx: np.array(arr, copy=True)}
            fds_new = [dict(fd) for fd in fds]
            for i, fd_new in enumerate(fds_new):
                base_caption = fd_new.get("Caption") or Path(path).name
                fd_new["FileName"] = f"{Path(path).name}_{op_name}_ch{i}"
                fd_new["Caption"] = f"{base_caption} {tag}"
            header_new = dict(header)
            header_new["xPixel"] = int(arr.shape[1])
            header_new["yPixel"] = int(arr.shape[0])
            stored_extent = None
            view_extent = view.get("extent_raw")
            if view_extent is None:
                view_extent = view.get("extent")
            if view_extent is not None and len(view_extent) == 4:
                try:
                    x0, x1, y_a, y_b = [float(v) for v in view_extent]
                    stored_extent = (x0, x1, y_a, y_b)
                    xmin, xmax = sorted((x0, x1))
                    ymin, ymax = sorted((y_a, y_b))
                    x_range = max(xmax - xmin, 1e-12)
                    y_range = max(ymax - ymin, 1e-12)
                    x_center = 0.5 * (xmin + xmax)
                    y_center = 0.5 * (ymin + ymax)
                    header_new["XScanRange"] = x_range
                    header_new["YScanRange"] = y_range
                    header_new["XRange"] = x_range
                    header_new["YRange"] = y_range
                    header_new["xCenter"] = x_center
                    header_new["yCenter"] = y_center
                    header_new["XCenter"] = x_center
                    header_new["YCenter"] = y_center
                except Exception:
                    pass
            key = self._make_processed_key(path, op=op_name, channel_idx=ch_idx)
            self._processed_views[key] = {
                "arr_by_channel": arr_by_channel,
                "header": header_new,
                "fds": fds_new,
                "channel_idx": ch_idx,
                "source": path,
                "extent_raw": stored_extent,
                "label": tag,
                "op": op_name,
            }
            self.headers[key] = (header_new, fds_new)
            self._insert_processed_after_source(key, path, insert_after_key=insert_after_key)
            self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())
            return key
        except Exception:
            return None

    def _create_virtual_crop_view(self, view, insert_after_key=None):
        """Create a virtual copy from a cropped preview view (single channel)."""
        return self._create_virtual_view_copy(view, insert_after_key=insert_after_key, tag="[crop]", op="crop")

    def _create_virtual_copy_from_history(self, seq):
        if seq is None:
            return
        entry = None
        if hasattr(self, "quick_crop_controller"):
            entry = self.quick_crop_controller.get_history_entry(seq)
        if not entry:
            return
        view_snapshot = entry.get("view_snapshot")
        if not view_snapshot:
            QtWidgets.QMessageBox.information(self, "Virtual copy", "This crop does not have a stored snapshot.")
            return
        self._create_virtual_crop_view(dict(view_snapshot))

    def on_clear_spec_selection(self):
        self._clear_multi_spec_selection()

    def _open_multi_spectroscopy_popup(self):
        controller = getattr(self, "spectro_compare_controller", None)
        if controller:
            return controller.open_multi_popup()
        if not self._spectros_loaded:
            self.ensure_spectros_loaded(refresh=False)
        return spectro_popups._open_multi_spectroscopy_popup(self)

    def on_show_matrix_spectro_viewer(self):
        if not self._spectros_loaded:
            self.ensure_spectros_loaded(refresh=False)
        return spectro_popups.on_show_matrix_spectro_viewer(self)

    def on_spec_coord_mode_changed(self, idx):
        try:
            self.spec_coord_mode = self.spec_coord_combo.currentText()
        except Exception:
            self.spec_coord_mode = 'Auto'
        self.config['spec_coord_mode'] = self.spec_coord_mode; save_config(self.config)
        self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())
        if self.last_preview:
            self.show_file_channel(self.last_preview[0], self.last_preview[1])

    def on_spec_invert_changed(self, checked: bool):
        self.spec_invert_y = bool(checked)
        self.config['spectro_invert_y'] = self.spec_invert_y; save_config(self.config)
        self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())
        if self.last_preview:
            self.show_file_channel(self.last_preview[0], self.last_preview[1])

    def on_pick_spectro_single_color(self):
        col = QtWidgets.QColorDialog.getColor(self.spectro_marker_color_single, self, "Select Single Marker Color", QtWidgets.QColorDialog.ShowAlphaChannel)
        if col.isValid():
            self.spectro_marker_color_single = col
            self.config['spectro_marker_color_single'] = col.name(QtGui.QColor.HexArgb)
            save_config(self.config)
            self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())
            if self.last_preview:
                self.show_file_channel(self.last_preview[0], self.last_preview[1])
            self._schedule_marker_refresh()

    def on_pick_spectro_matrix_color(self):
        col = QtWidgets.QColorDialog.getColor(self.spectro_marker_color_matrix, self, "Select Matrix Marker Color", QtWidgets.QColorDialog.ShowAlphaChannel)
        if col.isValid():
            self.spectro_marker_color_matrix = col
            self.config['spectro_marker_color_matrix'] = col.name(QtGui.QColor.HexArgb)
            save_config(self.config)
            self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())
            if self.last_preview:
                self.show_file_channel(self.last_preview[0], self.last_preview[1])
            self._schedule_marker_refresh()
            self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())

    def set_spectro_color_cycle(self, name: str):
        cycle = name or DEFAULT_COLOR_CYCLE
        if cycle == self.spectro_color_cycle:
            return
        self.spectro_color_cycle = cycle
        self.config['spectro_color_cycle'] = cycle
        save_config(self.config)
        for dlg in getattr(self, '_multi_spectro_popups', []):
            try:
                if dlg and dlg.isVisible() and hasattr(dlg, 'set_palette_name'):
                    dlg.set_palette_name(cycle)
            except Exception:
                continue

    def on_set_spectro_symbol(self, symbol):
        self.spectro_marker_symbol = symbol
        self.config['spectro_marker_symbol'] = symbol
        save_config(self.config)
        self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())
        if self.last_preview:
            self.show_file_channel(self.last_preview[0], self.last_preview[1])
        self._schedule_marker_refresh()

    def on_set_spectro_size(self, size):
        self.spectro_marker_size = float(size)
        self.config['spectro_marker_size'] = self.spectro_marker_size
        save_config(self.config)
        self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())
        if self.last_preview:
            self.show_file_channel(self.last_preview[0], self.last_preview[1])
        self._schedule_marker_refresh()

    def _populate_marker_style_menu(self, menu):
        col_single = menu.addAction("Single marker color...")
        col_single.triggered.connect(self.on_pick_spectro_single_color)
        col_matrix = menu.addAction("Matrix marker color...")
        col_matrix.triggered.connect(self.on_pick_spectro_matrix_color)
        menu.addSeparator()
        sym_grp = QtWidgets.QActionGroup(menu)
        current_symbol = getattr(self, 'spectro_marker_symbol', 'circle')
        for sym in ['circle', 'square', 'triangle', 'diamond']:
            act = QtWidgets.QAction(sym.capitalize(), menu)
            act.setCheckable(True)
            act.setChecked(current_symbol == sym)
            act.triggered.connect(lambda checked, s=sym: self.on_set_spectro_symbol(s))
            sym_grp.addAction(act)
        menu.addSeparator()
        size_menu = menu.addMenu("Marker Size")
        size_grp = QtWidgets.QActionGroup(menu)
        current_size = getattr(self, 'spectro_marker_size', 5.0)
        for label, val in [("Tiny", 2.0), ("Small", 3.5), ("Medium", 5.0), ("Large", 7.0), ("Huge", 10.0)]:
            act = size_menu.addAction(label)
            act.setCheckable(True)
            act.setChecked(abs(current_size - val) < 0.1)
            act.triggered.connect(lambda checked, v=val: self.on_set_spectro_size(v))
            size_grp.addAction(act)
        return menu

    def on_meta_font_changed(self, val:int):
        try:
            font = self.meta_box.font()
            font.setPointSize(int(val))
            self.meta_box.setFont(font)
            self.config['meta_font_size'] = int(val); save_config(self.config)
            # Re-render current metadata HTML so inline styles reflect the new font size
            try:
                if getattr(self, 'last_preview', None):
                    self.show_file_channel(self.last_preview[0], self.last_preview[1])
            except Exception:
                pass
        except Exception:
            pass

    def on_preview_lock_toggled(self, checked: bool):
        self.preview_locked = bool(checked)
        for widget in (
            getattr(self, "preview_lock_cb", None),
            getattr(self, "tools_preview_lock_act", None),
        ):
            if widget is None:
                continue
            try:
                widget.blockSignals(True)
                widget.setChecked(self.preview_locked)
                widget.blockSignals(False)
            except Exception:
                pass
        self.config["preview_locked"] = self.preview_locked; save_config(self.config)
        self._update_preview_detach_button()
        if self.preview_locked and getattr(self, "preview_detached", False):
            self._attach_preview()

    def _update_preview_detach_button(self):
        btn = getattr(self, "preview_detach_btn", None)
        act = getattr(self, "tools_preview_detach_act", None)
        detached = bool(getattr(self, "preview_detached", False))
        locked = bool(getattr(self, "preview_locked", False))
        label = "Dock preview" if detached else "Float preview"
        tooltip = (
            "Dock the floating preview back into the main window"
            if detached
            else "Detach the preview pane into its own floating window"
        )
        for widget in (btn, act):
            if widget is None:
                continue
            try:
                widget.setText(label)
                widget.setToolTip(tooltip)
                widget.setEnabled(not locked)
            except Exception:
                pass

    def on_toggle_preview_detach(self):
        if self.preview_locked:
            return
        if getattr(self, "preview_detached", False):
            self._attach_preview()
        else:
            self._detach_preview()

    def _detach_preview(self):
        if self.preview_locked or getattr(self, "preview_detached", False):
            try:
                if self._preview_dialog:
                    self._preview_dialog.show(); self._preview_dialog.raise_(); self._preview_dialog.activateWindow()
            except Exception:
                pass
            return
        try:
            # Remove from splitter
            try:
                w = self.main_splitter.widget(2)
                if w is self._preview_panel:
                    self._preview_panel.setParent(None)
            except Exception:
                pass
            if self._preview_dialog is None:
                dlg = QtWidgets.QDialog(self)
                dlg.setWindowTitle("Preview")
                dlg.setAttribute(QtCore.Qt.WA_DeleteOnClose, False)
                layout = QtWidgets.QVBoxLayout()
                layout.setContentsMargins(0, 0, 0, 0)
                dlg.setLayout(layout)
                self._preview_dialog = dlg
            lay = self._preview_dialog.layout()
            if lay is not None and self._preview_panel not in lay.children():
                try:
                    lay.addWidget(self._preview_panel)
                except Exception:
                    pass
            self.preview_detached = True
            self._update_preview_detach_button()
            try:
                self._preview_dialog.resize(self._preview_panel.size())
            except Exception:
                pass
            if self._preview_dialog:
                # Ensure re-attach when the dialog is closed
                def _on_close(ev):
                    self._attach_preview()
                    ev.accept()
                try:
                    self._preview_dialog.closeEvent = _on_close
                except Exception:
                    pass
                try:
                    self._preview_dialog.show(); self._preview_dialog.raise_(); self._preview_dialog.activateWindow()
                except Exception:
                    pass
        except Exception:
            pass

    def _attach_preview(self):
        if not getattr(self, "preview_detached", False):
            return
        try:
            if self._preview_dialog and self._preview_panel:
                try:
                    lay = self._preview_dialog.layout()
                    if lay is not None:
                        lay.removeWidget(self._preview_panel)
                except Exception:
                    pass
                try:
                    self._preview_dialog.hide()
                except Exception:
                    pass
            # Insert back into the main splitter at index 2 (rightmost pane)
            try:
                self.main_splitter.insertWidget(2, self._preview_panel)
                self.main_splitter.setStretchFactor(0, 1)
                self.main_splitter.setStretchFactor(1, 2)
                self.main_splitter.setStretchFactor(2, 3)
            except Exception:
                pass
            self.preview_detached = False
            self._update_preview_detach_button()
        except Exception:
            pass

    def on_dark_mode_toggled(self, checked: bool):
        self.dark_mode = bool(checked)
        try:
            if hasattr(self, 'toolbar_dark_btn'):
                self.toolbar_dark_btn.blockSignals(True)
                self.toolbar_dark_btn.setChecked(self.dark_mode)
                self.toolbar_dark_btn.blockSignals(False)
        except Exception:
            pass
        if getattr(self, "_detail_theme_follows_dark_mode", True):
            self._set_detail_dark_view_state(self.dark_mode, follow_dark_mode=True, persist=False)
        self.config['dark_mode'] = self.dark_mode
        self.config['detail_dark_view'] = self.detail_dark_view
        self.config['detail_theme_follows_dark_mode'] = bool(
            getattr(self, "_detail_theme_follows_dark_mode", True)
        )
        save_config(self.config)
        self._apply_dark_mode(self.dark_mode)
        if self.last_preview:
            self.show_file_channel(self.last_preview[0], self.last_preview[1])

    # ---------- control callbacks ----------
    def on_channel_dropdown_changed(self, idx):
        self.last_channel_index = int(idx); self.config['last_channel_index'] = self.last_channel_index; save_config(self.config)
        self._sync_channel_nav_buttons()
        self.populate_thumbnails_for_channel(idx)
        if getattr(self, 'frame_real_view', False):
            self._refresh_frame_map_pixmaps()
        # Refresh preview to the current selection on the newly chosen channel
        try:
            target_file = None
            if getattr(self, 'selected_file_for_thumbs', None):
                target_file = self.selected_file_for_thumbs
            elif getattr(self, 'current_thumb_files', None):
                target_file = self.current_thumb_files[0] if self.current_thumb_files else None
            if target_file:
                self.show_file_channel(target_file, idx)
        except Exception:
            pass

    def on_thumb_cmap_changed(self, idx):
        return viewer_thumb_ui.on_thumb_cmap_changed(self, idx)

    def on_preview_cmap_changed(self, idx):
        return viewer_preview.on_preview_cmap_changed(self, idx)

    def on_show_spectra_toggled(self, checked):
        self.show_spectra = bool(checked)
        self.config['show_spectra'] = self.show_spectra; save_config(self.config)
        # Keep UI toggles in sync
        try:
            for attr in (
                "spectro_overlay_act",
                "preview_spectra_toggle_btn",
                "spectro_thumbnail_markers_cb",
                "toolbar_spectro_markers_act",
            ):
                widget = getattr(self, attr, None)
                if widget is None:
                    continue
                widget.blockSignals(True)
                widget.setChecked(self.show_spectra)
                widget.blockSignals(False)
        except Exception:
            pass
        if self.show_spectra:
            if not self._spectros_loaded:
                self.ensure_spectros_loaded(refresh=False)
            else:
                # already loaded for this session; just update counts
                self._update_spectro_stats_label()
        else:
            self._clear_multi_spec_selection()
            self._update_spectro_stats_label()
        self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())
        if self.last_preview:
            self.show_file_channel(self.last_preview[0], self.last_preview[1])
        self._schedule_marker_refresh()

    def on_show_spectro_miniatures_toggled(self, checked: bool):
        self.show_spectro_miniatures = bool(checked)
        self.config["show_spectro_miniatures"] = self.show_spectro_miniatures
        save_config(self.config)
        try:
            for attr in (
                "spectro_miniatures_act",
                "spectro_miniatures_cb",
                "toolbar_spectro_miniatures_act",
            ):
                widget = getattr(self, attr, None)
                if widget is None:
                    continue
                widget.blockSignals(True)
                widget.setChecked(self.show_spectro_miniatures)
                widget.blockSignals(False)
        except Exception:
            pass
        if self.show_spectro_miniatures and not self._spectros_loaded:
            self.ensure_spectros_loaded(refresh=False)
        self._update_spectro_stats_label()
        # Miniatures are a thumbnail presentation choice, so a full repopulate is enough.
        self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())

    def on_show_preview_spectra_toggled(self, checked: bool):
        self.show_preview_spectra = bool(checked)
        self.config['show_preview_spectra'] = self.show_preview_spectra; save_config(self.config)
        try:
            for attr in (
                "show_spectra_cb",
                "spectro_preview_markers_cb",
                "toolbar_spectro_preview_act",
            ):
                widget = getattr(self, attr, None)
                if widget is None:
                    continue
                widget.blockSignals(True)
                widget.setChecked(self.show_preview_spectra)
                widget.blockSignals(False)
        except Exception:
            pass
        if self.show_preview_spectra and not self._spectros_loaded:
            self.ensure_spectros_loaded(refresh=False)
        self._update_spectro_stats_label()
        if self.last_preview:
            self.show_file_channel(self.last_preview[0], self.last_preview[1])

    def on_spectro_share_overlapping_repeats_toggled(self, checked: bool):
        self.spectro_share_overlapping_repeats = bool(checked)
        self.config["spectro_share_overlapping_repeats"] = self.spectro_share_overlapping_repeats
        save_config(self.config)
        act = getattr(self, "toolbar_spectro_repeat_share_act", None)
        if act is not None:
            try:
                act.blockSignals(True)
                act.setChecked(self.spectro_share_overlapping_repeats)
                act.blockSignals(False)
            except Exception:
                pass
        if self.spectros:
            self._assign_spectros_to_images()
        self._update_spectro_stats_label()
        self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())
        if self.last_preview:
            self.show_file_channel(self.last_preview[0], self.last_preview[1])
        self._schedule_marker_refresh()

    def on_toggle_highlight_glow(self, checked: bool):
        self.spectro_highlight_glow = bool(checked)
        self.config['spectro_highlight_glow'] = self.spectro_highlight_glow; save_config(self.config)
        for attr in ("highlight_glow_act", "toolbar_spectro_highlight_act"):
            act = getattr(self, attr, None)
            if act is not None:
                try:
                    act.blockSignals(True)
                    act.setChecked(self.spectro_highlight_glow)
                    act.blockSignals(False)
                except Exception:
                    pass
        if not self.spectro_highlight_glow:
            self._highlight_spectrum_entry(None)
        else:
            self._schedule_marker_refresh()
            if self.last_preview:
                self.show_file_channel(self.last_preview[0], self.last_preview[1])

    def on_spectro_grid_as_matrix_toggled(self, checked: bool):
        self.spectro_single_grid_as_matrix = bool(checked)
        self.config["spectro_single_grid_as_matrix"] = self.spectro_single_grid_as_matrix
        save_config(self.config)
        act = getattr(self, "toolbar_spectro_grid_as_matrix_act", None)
        if act is not None:
            try:
                act.blockSignals(True)
                act.setChecked(self.spectro_single_grid_as_matrix)
                act.blockSignals(False)
            except Exception:
                pass
        self._reload_spectros(refresh=True)

    def on_spectro_force_single_toggled(self, checked: bool):
        self.spectro_force_single_mode = bool(checked)
        self.config["spectro_force_single_mode"] = self.spectro_force_single_mode
        save_config(self.config)
        act = getattr(self, "toolbar_spectro_force_single_act", None)
        if act is not None:
            try:
                act.blockSignals(True)
                act.setChecked(self.spectro_force_single_mode)
                act.blockSignals(False)
            except Exception:
                pass
        self._reload_spectros(refresh=True)

    def on_show_matrix_markers_toggled(self, checked: bool):
        self.show_matrix_markers = bool(checked)
        self.config['show_matrix_markers'] = self.show_matrix_markers; save_config(self.config)
        self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())
        if self.last_preview:
            self.show_file_channel(self.last_preview[0], self.last_preview[1])
        self._schedule_marker_refresh()
        for attr in ('matrix_markers_act', 'toolbar_spectro_matrix_markers_act'):
            act = getattr(self, attr, None)
            if act is not None:
                act.blockSignals(True)
                act.setChecked(self.show_matrix_markers)
                act.blockSignals(False)

    def on_show_single_markers_toggled(self, checked: bool):
        self.show_single_markers = bool(checked)
        self.config['show_single_markers'] = self.show_single_markers; save_config(self.config)
        self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())
        if self.last_preview:
            self.show_file_channel(self.last_preview[0], self.last_preview[1])
        self._schedule_marker_refresh()
        for attr in ('single_markers_act', 'toolbar_spectro_single_markers_act'):
            act = getattr(self, attr, None)
            if act is not None:
                act.blockSignals(True)
                act.setChecked(self.show_single_markers)
                act.blockSignals(False)

    def on_compact_markers_toggled(self, checked: bool):
        self.compact_markers = bool(checked)
        self.config['compact_markers'] = self.compact_markers; save_config(self.config)
        self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())
        if self.last_preview:
            self.show_file_channel(self.last_preview[0], self.last_preview[1])
        self._schedule_marker_refresh()
        for attr in ('compact_markers_act', 'toolbar_spectro_compact_markers_act'):
            act = getattr(self, attr, None)
            if act is not None:
                act.blockSignals(True)
                act.setChecked(self.compact_markers)
                act.blockSignals(False)

    def on_detail_dark_toggled(self, checked: bool):
        self._set_detail_dark_view_state(
            checked,
            follow_dark_mode=(bool(checked) == bool(getattr(self, "dark_mode", False))),
            persist=True,
        )
        self._apply_detail_view_theme()

    def on_detail_grid_toggled(self, checked: bool):
        self.detail_grid_view = bool(checked)
        self.config['detail_grid_view'] = self.detail_grid_view; save_config(self.config)
        try:
            for attr in ("detail_grid_act", "preview_grid_toggle_btn"):
                widget = getattr(self, attr, None)
                if widget is None:
                    continue
                widget.blockSignals(True)
                widget.setChecked(self.detail_grid_view)
                widget.blockSignals(False)
        except Exception:
            pass
        self._apply_detail_view_theme()

    def _canvas_display_state_from_canvas(self, canvas):
        if canvas is None:
            return {}
        try:
            layout = str(getattr(canvas, "_view_layout", "grid") or "grid").strip().lower()
            if layout not in ("grid", "stacked"):
                layout = "grid"
        except Exception:
            layout = "grid"
        relative_axes = getattr(canvas, "_relative_axes_override", None)
        if relative_axes is not None:
            relative_axes = bool(relative_axes)
        return {
            "show_ticks": bool(getattr(canvas, "_show_ticks", True)),
            "show_colorbar": bool(getattr(canvas, "_show_colorbar", True)),
            "colorbar_orientation": str(getattr(canvas, "_colorbar_orientation", "vertical") or "vertical").strip().lower(),
            "show_title": bool(getattr(canvas, "_show_title", True)),
            "show_acquisition_overlay": bool(getattr(canvas, "_show_acquisition_overlay", False)),
            "show_shortcut_hint": bool(getattr(canvas, "_show_shortcut_hint", True)),
            "show_profile_overlays": bool(getattr(canvas, "_show_profile_overlays", True)),
            "show_angle_overlays": bool(getattr(canvas, "_show_angle_overlays", True)),
            "show_molecules": bool(getattr(canvas, "show_molecules", True)),
            "show_molecule_gizmo": bool(getattr(canvas, "_show_molecule_gizmo", False)),
            "scale_bar_enabled": bool(getattr(canvas, "scale_bar_enabled", False)),
            "frame_fill_mode": bool(getattr(canvas, "_frame_fill_mode", False)),
            "relative_axes_override": relative_axes,
            "view_layout": layout,
        }

    def _on_canvas_display_options_changed(self, canvas):
        if self._canvas_display_syncing:
            return
        options = self._canvas_display_state_from_canvas(canvas)
        if not options:
            return
        if options == getattr(self, "_last_canvas_display_options", {}):
            return
        self._apply_canvas_display_options(options, source_canvas=canvas, persist=True)

    def _apply_canvas_display_options(self, options, source_canvas=None, persist=True):
        if not isinstance(options, dict) or not options:
            return
        self._canvas_display_syncing = True
        try:
            normalized = {
                "show_ticks": bool(options.get("show_ticks", True)),
                "show_colorbar": bool(options.get("show_colorbar", True)),
                "colorbar_orientation": str(options.get("colorbar_orientation", "vertical") or "vertical").strip().lower(),
                "show_title": bool(options.get("show_title", True)),
                "show_acquisition_overlay": bool(options.get("show_acquisition_overlay", False)),
                "show_shortcut_hint": bool(options.get("show_shortcut_hint", True)),
                "show_profile_overlays": bool(options.get("show_profile_overlays", True)),
                "show_angle_overlays": bool(options.get("show_angle_overlays", True)),
                "show_molecules": bool(options.get("show_molecules", True)),
                "show_molecule_gizmo": bool(options.get("show_molecule_gizmo", False)),
                "scale_bar_enabled": bool(options.get("scale_bar_enabled", False)),
                "frame_fill_mode": bool(options.get("frame_fill_mode", False)),
                "relative_axes_override": options.get("relative_axes_override", None),
                "view_layout": str(options.get("view_layout", "grid") or "grid").strip().lower(),
            }
            if normalized["view_layout"] not in ("grid", "stacked"):
                normalized["view_layout"] = "grid"
            if normalized["colorbar_orientation"] not in ("vertical", "horizontal"):
                normalized["colorbar_orientation"] = "vertical"
            rel = normalized["relative_axes_override"]
            if rel is not None:
                normalized["relative_axes_override"] = bool(rel)

            self.show_molecules = normalized["show_molecules"]
            self.show_molecule_gizmo = normalized["show_molecule_gizmo"]
            self.show_acquisition_overlay = normalized["show_acquisition_overlay"]
            try:
                if hasattr(self, "scale_bar_cb") and self.scale_bar_cb is not None:
                    self.scale_bar_cb.blockSignals(True)
                    self.scale_bar_cb.setChecked(normalized["scale_bar_enabled"])
                    self.scale_bar_cb.blockSignals(False)
                if hasattr(self, "display_scale_bar_act") and self.display_scale_bar_act is not None:
                    self.display_scale_bar_act.blockSignals(True)
                    self.display_scale_bar_act.setChecked(normalized["scale_bar_enabled"])
                    self.display_scale_bar_act.blockSignals(False)
            except Exception:
                pass
            for widget_name, key in (
                ("molecules_act", "show_molecules"),
                ("browse_molecules_btn", "show_molecules"),
                ("preview_molecules_toggle_btn", "show_molecules"),
                ("display_molecule_gizmo_act", "show_molecule_gizmo"),
                ("acquisition_overlay_act", "show_acquisition_overlay"),
            ):
                act = getattr(self, widget_name, None)
                if act is not None:
                    try:
                        act.blockSignals(True)
                        act.setChecked(bool(normalized[key]))
                        act.blockSignals(False)
                    except Exception:
                        pass

            canvases = [getattr(self, "preview_canvas", None)] + list(getattr(self, "_popup_canvases", []))
            for canv in canvases:
                if canv is None:
                    continue
                try:
                    canv._show_ticks = normalized["show_ticks"]
                    canv._show_colorbar = normalized["show_colorbar"]
                    canv._colorbar_orientation = normalized["colorbar_orientation"]
                except Exception:
                    pass
                try:
                    canv.set_show_title(normalized["show_title"])
                except Exception:
                    pass
                try:
                    canv.set_show_acquisition_overlay(normalized["show_acquisition_overlay"])
                except Exception:
                    pass
                try:
                    canv.set_show_shortcut_hint(normalized["show_shortcut_hint"])
                except Exception:
                    pass
                try:
                    canv.set_show_profile_overlays(normalized["show_profile_overlays"])
                except Exception:
                    pass
                try:
                    canv.set_show_angle_overlays(normalized["show_angle_overlays"])
                except Exception:
                    pass
                try:
                    canv.set_show_molecules(normalized["show_molecules"])
                except Exception:
                    pass
                try:
                    canv.set_show_molecule_gizmo(normalized["show_molecule_gizmo"])
                except Exception:
                    pass
                try:
                    canv.enable_scale_bar(normalized["scale_bar_enabled"])
                except Exception:
                    pass
                try:
                    canv.set_frame_fill_mode(normalized["frame_fill_mode"])
                except Exception:
                    pass
                try:
                    canv.set_relative_axes_override(normalized["relative_axes_override"])
                except Exception:
                    pass
                try:
                    canv.set_view_layout(normalized["view_layout"])
                except Exception:
                    pass
                try:
                    canv._redraw()
                except Exception:
                    pass

            self._last_canvas_display_options = dict(normalized)
            self.canvas_display_options = dict(normalized)
            if persist:
                self.config["canvas_display_options"] = dict(normalized)
                self.config["show_molecules"] = self.show_molecules
                self.config["show_molecule_gizmo"] = self.show_molecule_gizmo
                self.config["show_acquisition_overlay"] = self.show_acquisition_overlay
                self.config["show_scale_bar"] = normalized["scale_bar_enabled"]
                save_config(self.config)
        finally:
            self._canvas_display_syncing = False

    def on_show_molecules_toggled(self, checked: bool):
        self.show_molecules = bool(checked)
        options = self._canvas_display_state_from_canvas(getattr(self, "preview_canvas", None))
        options["show_molecules"] = self.show_molecules
        self._apply_canvas_display_options(options, source_canvas=getattr(self, "preview_canvas", None), persist=True)

    def on_show_acquisition_overlay_toggled(self, checked: bool):
        self.show_acquisition_overlay = bool(checked)
        options = self._canvas_display_state_from_canvas(getattr(self, "preview_canvas", None))
        options["show_acquisition_overlay"] = self.show_acquisition_overlay
        self._apply_canvas_display_options(options, source_canvas=getattr(self, "preview_canvas", None), persist=True)

    def on_profile_label_mode_changed(self, mode: str):
        mode = str(mode or "length").strip().lower()
        if mode not in {"length", "full", "hidden"}:
            mode = "length"
        self.profile_label_mode = mode
        self.config["profile_label_mode"] = mode
        save_config(self.config)
        canvases = [getattr(self, "preview_canvas", None)] + list(getattr(self, "_popup_canvases", []))
        for canv in canvases:
            if canv is None:
                continue
            try:
                canv.set_profile_label_mode(mode)
            except Exception:
                continue
        for key, action in (getattr(self, "profile_label_actions", {}) or {}).items():
            if action is None:
                continue
            action.blockSignals(True)
            action.setChecked(key == mode)
            action.blockSignals(False)

    def on_fixed_crop_quick_toggled(self, checked: bool):
        self._set_quick_crop_mode(checked)

    def on_show_crop_template_overlay_toggled(self, checked: bool):
        self.show_crop_template_overlay = bool(checked)
        self.config['show_crop_template_overlay'] = self.show_crop_template_overlay; save_config(self.config)
        canvases = [getattr(self, 'preview_canvas', None)] + list(getattr(self, '_popup_canvases', []))
        for canv in canvases:
            if canv:
                try:
                    canv.show_fixed_crop_template(self.show_crop_template_overlay)
                except Exception:
                    continue
        act = getattr(self, 'crop_template_act', None)
        if act is not None:
            act.blockSignals(True)
            act.setChecked(self.show_crop_template_overlay)
            act.blockSignals(False)

    def on_show_crop_history_overlay_toggled(self, checked: bool):
        self.show_crop_history_overlay = True
        self.config['show_crop_history_overlay'] = True; save_config(self.config)
        act = getattr(self, 'crop_history_act', None)
        if act is not None:
            act.blockSignals(True)
            act.setChecked(True)
            act.blockSignals(False)
        cb = getattr(self, 'crop_history_overlay_cb', None)
        if cb is not None:
            cb.blockSignals(True)
            cb.setChecked(True)
            cb.blockSignals(False)
        if hasattr(self, "quick_crop_controller"):
            self.quick_crop_controller.update_hint()
            self.quick_crop_controller.refresh_history_panel()
            self.quick_crop_controller.update_active_sequence_from_stack()

    def _set_quick_crop_mode(self, enabled: bool, save: bool = True):
        controller = getattr(self, "quick_crop_controller", None)
        if controller is None:
            self.quick_crop_mode = bool(enabled)
            if save:
                self.config['quick_crop_mode'] = self.quick_crop_mode
                save_config(self.config)
            return
        controller.set_mode(enabled, save=save)

    def _update_quick_crop_hint(self):
        controller = getattr(self, "quick_crop_controller", None)
        if controller:
            controller.update_hint()

    def _on_preview_canvas_state_changed(self, canvas):
        self._on_canvas_display_options_changed(canvas)
        self._update_quick_crop_hint()

    def _sync_quick_crop_template_controls(self):
        controller = getattr(self, "quick_crop_controller", None)
        if controller:
            controller.sync_template_controls()

    def _on_quick_crop_edit_toggled(self, checked: bool):
        controller = getattr(self, "quick_crop_controller", None)
        if controller:
            controller.set_edit_mode(bool(checked))

    def _on_quick_crop_aspect_mode_changed(self):
        combo = getattr(self, "quick_crop_aspect_combo", None)
        if combo is None:
            return
        mode = str(combo.currentData() or combo.currentText() or "free").strip().lower()
        if mode not in {"free", "keep", "square"}:
            mode = "free"
        self.quick_crop_aspect_mode = mode
        self.config["quick_crop_aspect_mode"] = mode
        save_config(self.config)
        controller = getattr(self, "quick_crop_controller", None)
        if controller:
            controller.on_aspect_mode_changed()

    def _on_quick_crop_real_spin_changed(self, _=None):
        try:
            sender = self.sender()
        except Exception:
            sender = None
        controller = getattr(self, "quick_crop_controller", None)
        if controller:
            controller.on_real_spin_changed(sender)

    def _apply_quick_crop_template_from_controls(self):
        controller = getattr(self, "quick_crop_controller", None)
        if controller:
            controller.apply_template_from_controls()

    def _on_fixed_crop_history_updated(self, entries):
        controller = getattr(self, "quick_crop_controller", None)
        if controller:
            controller.on_history_updated(entries)

    def on_export_selected_same_view(self):
        return viewer_export.on_export_selected_same_view(self)
    def on_export_stp_files(self):
        return viewer_export.on_export_wsxm_stp_files(self)

    def _export_view_as_stp(self, view):
        return viewer_export.export_view_as_stp(self, view)

    def _on_batch_export_progress(self, current, total, path):
        return viewer_export._on_batch_export_progress(self, current, total, path)

    def _on_batch_export_finished(self, saved_paths, errors, cancelled):
        return viewer_export._on_batch_export_finished(self, saved_paths, errors, cancelled)

    def _on_purge_config(self):
        """Purge stored configuration data (tags, last_dir, cmaps) and clear runtime caches."""
        try:
            # backup current config
            try:
                if CONFIG_PATH.exists():
                    CONFIG_PATH.with_suffix('.bak').write_text(CONFIG_PATH.read_text())
            except Exception:
                pass
            # clear in-memory
            self.tags = {}
            self._invalidate_thumbnail_cache()
            self._invalidate_channel_cache()
            self.per_file_channel_cmap.clear()
            # clear config file on disk
            try:
                if CONFIG_PATH.exists():
                    CONFIG_PATH.unlink()
            except Exception:
                pass
            # reset defaults
            self.config = {}
            self.last_dir = Path.cwd()
            self.last_channel_index = 0
            QtWidgets.QMessageBox.information(self, 'Purge config', 'Configuration and tags purged. Please reopen your folder.')
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, 'Purge failed', str(e))
