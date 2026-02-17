"""Main Qt widget implementing the SXM Viewer."""
from __future__ import annotations

import math
import time
import re
import json
import os
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
from ..data.spectroscopy import is_matrix_file_entry
from ..processing.filters import (
    flatten_remove_median,
    subtract_best_fit_plane,
    subtract_2nd_order_plane,
    gaussian_filter_image,
    highpass_filter,
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
from .viewer import loader as viewer_loader
from .viewer import preview as viewer_preview
from .viewer.state import ViewerState
from .spectroscopy import controller as spectro_controller
from .spectroscopy import overlays as spectro_overlays
from .spectroscopy import popups as spectro_popups
from .viewer import thumbnail_ui as viewer_thumb_ui
from .viewer import export as viewer_export
from .canvases.canvas_window import ExperimentalCanvasWindow
from .palettes import DEFAULT_COLOR_CYCLE

# Tolerance for deciding constant-height images; allow a slightly larger spread than strict equality
CH_RANGE_TOL_NM = max(CH_EQUALITY_TOL_NM, 0.02)  # ~20 pm default floor

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
        self.setWindowTitle("SXM Viewer")
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
        self.detail_grid_view = bool(self.config.get('detail_grid_view', False))
        self.show_molecules = bool(self.config.get('show_molecules', True))
        self.molecule_palette = str(self.config.get("molecule_palette", "cpk") or "cpk").lower()
        self.recent_molecules = list(self.config.get("recent_molecules", []))
        self.quick_crop_mode = bool(self.config.get("quick_crop_mode", False))
        self.show_crop_template_overlay = bool(self.config.get("show_crop_template_overlay", False))
        self.show_crop_history_overlay = bool(self.config.get("show_crop_history_overlay", False))
        self._display_defaults = {
            'show_matrix_markers': True,
            'show_single_markers': True,
            'compact_markers': True,
            'detail_dark_view': bool(self.dark_mode),
            'detail_grid_view': False,
            'show_molecules': True,
            'show_crop_template_overlay': False,
            'show_crop_history_overlay': False,
        }
        self._popup_canvases = []
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
        self._spectro_deferred = set()
        # spectro_eager_limit: 0 means no deferral; otherwise parse at most N spectroscopy files eagerly
        limit_cfg = int(self.config.get("spectro_eager_limit", 300))
        self.spectro_eager_limit = max(0, limit_cfg)
        self.image_time_index = {}
        self._spectro_popups = []
        self._popup_refs = []
        self._multi_spectro_popups = []
        self._multi_single_popup_anchor = None
        self._last_clicked_spec = None
        self._popup_counter = 0  # used to stagger dialog positions
        self._multi_spec_selection = []
        self._multi_spec_selection_keys = set()
        self.thumb_multi_select = set()
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
        self.toolbar_adjust_act = None
        self._canvas_window = None

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

        # Channel controls (moved later into the Selected channel area)
        controls_h = QtWidgets.QHBoxLayout()
        self.channel_label = QtWidgets.QLabel("Channel:")
        self.channel_label.setFont(bold_font)
        self.channel_dropdown = QtWidgets.QComboBox(); self.channel_dropdown.setMinimumWidth(160)
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
        # Note: don't add these to the essentials panel here; we'll insert the layout into the Selected channel area below.
        controls_h.addWidget(self.channel_label); controls_h.addWidget(self.channel_dropdown)

        # Colormap combos will be shown in the main toolbar next to the dark-mode toggle (see main_window_toolbar)
        # Dark mode handled via toolbar toggle; placeholder kept for compatibility
        self.dark_mode_cb = None
        left_v.addWidget(essentials_group)

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
        title_lbl = QtWidgets.QLabel(""); title_lbl.setFont(bold_font)
        self.scroll = QtWidgets.QScrollArea(); self.thumb_container = QtWidgets.QWidget(); self.thumb_layout = QtWidgets.QGridLayout(); self.thumb_layout.setSpacing(THUMB_LAYOUT_SPACING)
        self.scroll.setToolTip(
            "Thumbnails:\n"
            "  Shift+Click or Ctrl+Click to multi-select\n"
            "  Ctrl+Wheel to change thumbnail size\n"
            "  Right-click a frame for filters & exports"
        )
        self.thumb_container.setLayout(self.thumb_layout); self.scroll.setWidgetResizable(True); self.scroll.setWidget(self.thumb_container)
        self._thumb_viewport = self.scroll.viewport()
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
        self.unit_display_cb = QtWidgets.QCheckBox("Show SI units")
        self.unit_display_cb.setChecked(self.display_units_si)
        self.unit_relative_cb = QtWidgets.QCheckBox("Relative zero")
        self.unit_relative_cb.setChecked(self.display_units_relative)
        self.relative_axes_cb = QtWidgets.QCheckBox("Relative axes")
        self.relative_axes_cb.setChecked(self.relative_axes)
        # Create a compact header: title on the left, channel controls on the right
        header_h = QtWidgets.QHBoxLayout()
        header_h.setContentsMargins(0,0,0,0)
        header_h.setSpacing(8)
        header_h.addWidget(title_lbl)
        header_h.addStretch(1)
        # create a compact container for the controls so they do not span the full width
        self.channel_controls_widget = QtWidgets.QWidget()
        self.channel_controls_widget.setLayout(controls_h)
        self.channel_controls_widget.setSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        header_h.addWidget(self.channel_controls_widget)
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
        preview_panel_layout = QtWidgets.QVBoxLayout(); preview_panel_layout.setContentsMargins(0,0,0,0)
        preview_header = QtWidgets.QHBoxLayout()
        preview_header.addWidget(QtWidgets.QLabel("Preview"))
        preview_header.addStretch(1)
        # Dock/lock controls
        self.preview_lock_cb = QtWidgets.QCheckBox("Lock")
        self.preview_lock_cb.setChecked(self.preview_locked)
        self.preview_lock_cb.setToolTip("Lock preview inside the main window")
        self.preview_lock_cb.toggled.connect(self.on_preview_lock_toggled)
        self.preview_detach_btn = QtWidgets.QToolButton()
        self.preview_detach_btn.setText("Detach")
        self.preview_detach_btn.setToolTip("Pop out the preview pane into its own window")
        self.preview_detach_btn.clicked.connect(self.on_toggle_preview_detach)
        self.preview_detach_btn.setEnabled(not self.preview_locked)
        display_strip = QtWidgets.QWidget()
        display_layout = QtWidgets.QHBoxLayout(display_strip)
        display_layout.setContentsMargins(0, 0, 0, 0)
        display_layout.setSpacing(8)
        display_layout.addWidget(self.unit_display_cb)
        display_layout.addWidget(self.unit_relative_cb)
        display_layout.addWidget(self.relative_axes_cb)
        self.scale_bar_cb = QtWidgets.QCheckBox("Scale bar")
        self.scale_bar_cb.setChecked(bool(self.config.get("show_scale_bar", False)))
        display_layout.addWidget(self.scale_bar_cb)
        self.preview_hist_btn = QtWidgets.QToolButton()
        self.preview_hist_btn.setText("Histogram")
        self.preview_hist_btn.setToolTip("Show histogram and adjust display range")
        self.preview_hist_btn.setPopupMode(QtWidgets.QToolButton.MenuButtonPopup)
        self.preview_hist_menu = QtWidgets.QMenu(self.preview_hist_btn)
        self.preview_hist_menu.addAction("Adjust…", lambda: self._open_histogram_dialog(self.preview_canvas))
        self.preview_hist_menu.addAction("Auto (1–99%)", lambda: self._auto_contrast(self.preview_canvas))
        self.preview_hist_menu.addAction("Reset range", lambda: self._reset_contrast(self.preview_canvas))
        self.show_preview_title = bool(self.config.get('show_preview_title', True))
        act_title = self.preview_hist_menu.addAction("Show title/date")
        act_title.setCheckable(True)
        act_title.setChecked(self.show_preview_title)
        act_title.triggered.connect(self._on_toggle_preview_title)
        self.preview_hist_btn.setMenu(self.preview_hist_menu)
        self.preview_hist_btn.clicked.connect(lambda _: self._open_histogram_dialog(self.preview_canvas))
        preview_header.addWidget(self.preview_hist_btn)
        preview_header.addWidget(self.preview_detach_btn)
        preview_header.addWidget(self.preview_lock_cb)
        preview_header.addWidget(display_strip)
        # Canvas launch button moved to the main toolbar for prominence.
        preview_panel_layout.addLayout(preview_header)

        self.quick_crop_controls = QtWidgets.QWidget()
        quick_layout = QtWidgets.QHBoxLayout(self.quick_crop_controls)
        quick_layout.setContentsMargins(0, 0, 0, 0)
        quick_layout.setSpacing(6)
        self.quick_crop_btn = QtWidgets.QPushButton("Quick crop: Off")
        self.quick_crop_btn.setCheckable(True)
        self.quick_crop_btn.setToolTip("Toggle quick crop mode (Ctrl+Shift+C). Shift+drag to size, click to spawn crops.")
        quick_layout.addWidget(self.quick_crop_btn)
        quick_layout.addWidget(QtWidgets.QLabel("Real W"))
        self.quick_crop_real_width_spin = QtWidgets.QDoubleSpinBox()
        self.quick_crop_real_width_spin.setRange(0.01, 10000.0)
        self.quick_crop_real_width_spin.setDecimals(3)
        self.quick_crop_real_width_spin.setSingleStep(0.1)
        self.quick_crop_real_width_spin.setFixedWidth(80)
        self.quick_crop_real_width_spin.setToolTip("Real-space width (current preview unit)")
        self.quick_crop_real_width_spin.setValue(5.0)
        self._quick_crop_aspect = 1.0
        self._quick_crop_last_real_size = [self.quick_crop_real_width_spin.value(), 5.0]
        quick_layout.addWidget(self.quick_crop_real_width_spin)
        quick_layout.addWidget(QtWidgets.QLabel("Real H"))
        self.quick_crop_real_height_spin = QtWidgets.QDoubleSpinBox()
        self.quick_crop_real_height_spin.setRange(0.01, 10000.0)
        self.quick_crop_real_height_spin.setDecimals(3)
        self.quick_crop_real_height_spin.setSingleStep(0.1)
        self.quick_crop_real_height_spin.setFixedWidth(80)
        self.quick_crop_real_height_spin.setToolTip("Real-space height (current preview unit)")
        self.quick_crop_real_height_spin.setValue(5.0)
        self._quick_crop_last_real_size = [self.quick_crop_real_width_spin.value(),
                                           self.quick_crop_real_height_spin.value()]
        self._quick_crop_aspect = self._quick_crop_last_real_size[0] / max(0.001, self._quick_crop_last_real_size[1])
        quick_layout.addWidget(self.quick_crop_real_height_spin)
        self.quick_crop_lock_aspect_cb = QtWidgets.QCheckBox("Lock aspect")
        self.quick_crop_lock_aspect_cb.setToolTip("Keep width/height ratio when editing one dimension")
        quick_layout.addWidget(self.quick_crop_lock_aspect_cb)
        self.quick_crop_real_unit_lbl = QtWidgets.QLabel("nm")
        quick_layout.addWidget(self.quick_crop_real_unit_lbl)
        self.quick_crop_square_cb = QtWidgets.QCheckBox("Square")
        self.quick_crop_square_cb.setToolTip("Force quick crops to remain square")
        quick_layout.addWidget(self.quick_crop_square_cb)
        self.quick_crop_real_px_info_lbl = QtWidgets.QLabel("")
        self.quick_crop_real_px_info_lbl.setFixedWidth(140)
        quick_layout.addWidget(self.quick_crop_real_px_info_lbl)
        self.quick_crop_undo_btn = QtWidgets.QToolButton()
        self.quick_crop_undo_btn.setText("Undo")
        self.quick_crop_undo_btn.setToolTip("Undo latest crop (Ctrl+Z)")
        quick_layout.addWidget(self.quick_crop_undo_btn)
        self.quick_crop_close_btn = QtWidgets.QToolButton()
        self.quick_crop_close_btn.setText("Close pop-out")
        self.quick_crop_close_btn.setToolTip("Close the latest quick-crop pop-out (Ctrl+Shift+W)")
        quick_layout.addWidget(self.quick_crop_close_btn)
        self.quick_crop_clear_btn = QtWidgets.QToolButton()
        self.quick_crop_clear_btn.setText("Clear history")
        self.quick_crop_clear_btn.setToolTip("Clear crop history markers and pop-outs")
        quick_layout.addWidget(self.quick_crop_clear_btn)
        self.quick_crop_export_btn = QtWidgets.QToolButton()
        self.quick_crop_export_btn.setText("Export selection")
        self.quick_crop_export_btn.setToolTip("Export the selected crops (Shift+click) as images")
        self.quick_crop_export_btn.setEnabled(False)
        self.quick_crop_controller = QuickCropController(self)
        self.quick_crop_export_btn.clicked.connect(self.quick_crop_controller.export_selected_crops)
        quick_layout.addWidget(self.quick_crop_export_btn)
        self.quick_crop_tile_btn = QtWidgets.QToolButton()
        self.quick_crop_tile_btn.setText("Tile pop-outs")
        self.quick_crop_tile_btn.setToolTip("Arrange all open pop-out windows on screen")
        self.quick_crop_tile_btn.setEnabled(False)
        self.quick_crop_tile_btn.clicked.connect(self.quick_crop_controller.arrange_popups)
        quick_layout.addWidget(self.quick_crop_tile_btn)
        quick_layout.addStretch(1)
        self.quick_crop_hint_lbl = QtWidgets.QLabel("")
        quick_layout.addWidget(self.quick_crop_hint_lbl)
        preview_panel_layout.addWidget(self.quick_crop_controls)
        self.quick_crop_btn.clicked.connect(lambda: self._set_quick_crop_mode(not self.quick_crop_mode))
        self.quick_crop_undo_btn.clicked.connect(self.quick_crop_controller.undo_last_crop)
        self.quick_crop_close_btn.clicked.connect(self.quick_crop_controller.close_latest_popup)
        self.quick_crop_clear_btn.clicked.connect(self.quick_crop_controller.clear_history)
        self.quick_crop_square_cb.toggled.connect(lambda _: self.quick_crop_controller.apply_template_from_controls())
        self.quick_crop_real_width_spin.valueChanged.connect(lambda _=None: self.quick_crop_controller.on_real_spin_changed(self.quick_crop_real_width_spin))
        self.quick_crop_real_height_spin.valueChanged.connect(lambda _=None: self.quick_crop_controller.on_real_spin_changed(self.quick_crop_real_height_spin))
        self.quick_crop_lock_aspect_cb.toggled.connect(lambda _: self.quick_crop_controller.on_real_spin_changed())

        self.crop_history_panel = QtWidgets.QWidget()
        crop_hist_layout = QtWidgets.QVBoxLayout(self.crop_history_panel)
        crop_hist_layout.setContentsMargins(0, 0, 0, 0)
        crop_hist_layout.setSpacing(4)
        self.crop_history_label = QtWidgets.QLabel("Crop history")
        crop_hist_layout.addWidget(self.crop_history_label)
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
            "  Right-click for copy/save options\n"
            "  Enable 'Measure profile' for line sampling\n"
            "  Ctrl+C copies the focused image to clipboard"
        )
        try:
            self.preview_canvas.set_show_title(self.show_preview_title)
        except Exception:
            pass
        try:
            self.preview_canvas.set_show_molecules(self.show_molecules)
        except Exception:
            pass
        self.preview_canvas.set_copy_feedback_handler(self._on_view_copied)
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
        self.preview_canvas.set_crop_callback(self._on_preview_crop)
        self.preview_canvas.set_double_click_callback(
            lambda v=None: self._spawn_preview_popup(
                [self._copy_view_for_popup(v)] if v else [],
                title=self._friendly_view_title(v, default="Preview copy") if v else "Preview copy",
            )
        )
        self.preview_canvas.set_filter_menu_callback(
            lambda menu, view, c=self.preview_canvas: self._populate_canvas_filter_menu(menu, c, view)
        )
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
        self.preview_canvas.show_fixed_crop_history(self.show_crop_history_overlay)
        self.quick_crop_toggle_shortcut = QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+Shift+C"), self)
        self.quick_crop_toggle_shortcut.setContext(QtCore.Qt.WidgetWithChildrenShortcut)
        self.quick_crop_toggle_shortcut.activated.connect(lambda: self._set_quick_crop_mode(not self.quick_crop_mode))
        self.quick_crop_undo_shortcut = QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+Z"), self)
        self.quick_crop_undo_shortcut.setContext(QtCore.Qt.WidgetWithChildrenShortcut)
        self.quick_crop_undo_shortcut.activated.connect(self.quick_crop_controller.undo_last_crop)
        self.quick_crop_close_shortcut = QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+Shift+W"), self)
        self.quick_crop_close_shortcut.setContext(QtCore.Qt.WidgetWithChildrenShortcut)
        self.quick_crop_close_shortcut.activated.connect(self.quick_crop_controller.close_latest_popup)
        self.quick_crop_real_shortcut = QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+Shift+R"), self)
        self.quick_crop_real_shortcut.setContext(QtCore.Qt.WidgetWithChildrenShortcut)
        self.quick_crop_real_shortcut.activated.connect(self.quick_crop_controller.apply_template_from_controls)
        self.quick_crop_history_shortcut = QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+Shift+H"), self)
        self.quick_crop_history_shortcut.setContext(QtCore.Qt.WidgetWithChildrenShortcut)
        self.quick_crop_history_shortcut.activated.connect(lambda: self.on_show_crop_history_overlay_toggled(not self.show_crop_history_overlay))
        self.quick_crop_template_shortcut = QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+Shift+T"), self)
        self.quick_crop_template_shortcut.setContext(QtCore.Qt.WidgetWithChildrenShortcut)
        self.quick_crop_template_shortcut.activated.connect(lambda: self.on_show_crop_template_overlay_toggled(not self.show_crop_template_overlay))
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
        container_layout = QtWidgets.QVBoxLayout()
        container_layout.setContentsMargins(0, 0, 0, 0)
        self.shortcuts_panel = self._create_shortcuts_panel()
        container_layout.addWidget(self.shortcuts_panel)
        if toolbar is not None:
            container_layout.addWidget(toolbar)
        container_layout.addWidget(main_splitter)
        self.setLayout(container_layout)
        self._set_shortcuts_panel_visible(self.show_shortcuts_panel, remember=False)
        # Ensure optimal initial thumbnail layout
        QtCore.QTimer.singleShot(200, lambda: self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex()))

        # signals
        self.open_btn.clicked.connect(self.open_folder_dialog)
        self.path_le.returnPressed.connect(self.open_folder_by_path)
        self.spec_folder_btn.clicked.connect(self.on_spec_folder_browse)
        self.spec_folder_le.returnPressed.connect(self.on_spec_folder_entered)
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
        self.show_spectra_cb.toggled.connect(self.on_show_preview_spectra_toggled)
        if hasattr(self, "grid_as_matrix_cb"):
            self.grid_as_matrix_cb.toggled.connect(self.on_spectro_grid_as_matrix_toggled)
        if hasattr(self, "force_single_cb"):
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
        self._init_mode_shortcuts()
        try:
            log_emitter.message_logged.connect(self._append_activity_log)
        except Exception:
            pass

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

    def _apply_detail_view_theme(self):
        canvas = getattr(self, 'preview_canvas', None)
        if canvas is not None and hasattr(canvas, 'set_detail_theme'):
            canvas.set_detail_theme(dark=self.detail_dark_view, grid=self.detail_grid_view)

    def _apply_molecule_button_theme(self):
        btn = getattr(self, "toolbar_load_mol_btn", None)
        if btn is None:
            return
        if getattr(self, "dark_mode", False):
            base = "#2d2d2d"
            hover = "#3a3a3a"
            color = QtGui.QColor("#ffffff")
        else:
            base = "#f0f3ff"
            hover = base
            color = QtGui.QColor("#1d1d1d")
        btn.setStyleSheet(
            f"""
QLabel {{
    background-color: {base};
    border: none;
    border-radius: 3px;
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
            btn.setPixmap(pixmap)

    def _append_activity_log(self, message: str):
        box = getattr(self, "activity_log_box", None)
        if box is None:
            return
        entry = f"[{datetime.now().strftime('%H:%M:%S')}] {message}"
        box.appendPlainText(entry)
        box.verticalScrollBar().setValue(box.verticalScrollBar().maximum())
        try:
            QtWidgets.QApplication.processEvents(QtCore.QEventLoop.AllEvents, 5)
        except Exception:
            pass

    def _on_clear_activity_log(self):
        if hasattr(self, "activity_log_box"):
            self.activity_log_box.clear()

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
            (QtGui.QKeySequence("Ctrl+S"), self.MODE_SPECTRO),
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
            (getattr(self, 'crop_template_act', None), defaults.get('show_crop_template_overlay', False)),
            (getattr(self, 'crop_history_act', None), defaults.get('show_crop_history_overlay', False)),
        ]
        for action, state in action_pairs:
            if action is not None:
                action.setChecked(state)

    def _update_spectro_stats_label(self, stats=None):
        return main_window_spectro.update_spectro_stats_label(self, stats=stats)

    def _create_shortcuts_panel(self):
        return main_window_layout.create_shortcuts_panel(self)

    # ---------- Spectroscopy quick-inspect helpers & dialog ----------
    def _header_extent(self, header):
        return main_window_spectro.header_extent(self, header)

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

    def _spawn_preview_popup(self, views, title=None):
        return spawn_preview_popup(self, views, title=title)

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

    def _on_preview_crop(self, view):
        """Receive cropped view from preview canvas and pop it out."""
        if not view:
            return
        # Offer to save crop as a virtual copy in thumbnails
        try:
            path = view.get("path") or (view.get("meta") or {}).get("path")
            if path:
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
        seq = view.get("crop_sequence")
        if self.show_crop_history_overlay and seq is not None:
            title = f"{title} #{seq}"
        self._spawn_preview_popup([self._copy_view_for_popup(view)], title=title)

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
            "<li><b>Ctrl+C</b> over preview = copy current image</li>"
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

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                if url.isLocalFile():
                    path = Path(url.toLocalFile())
                    if path.is_dir():
                        event.acceptProposedAction()
                        return
        super().dragEnterEvent(event)

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                if url.isLocalFile():
                    path = Path(url.toLocalFile())
                    if path.is_dir():
                        self.load_folder(path)
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
                    self._selection_before_drag = set(self.thumb_multi_select)
                    
                    if not (event.modifiers() & (QtCore.Qt.ShiftModifier | QtCore.Qt.ControlModifier)):
                        self._selection_before_drag = set()
                        self._clear_thumb_multi_selection()
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

    def _update_rubber_band_selection(self, rect, modifiers):
        in_rect = set()
        for key, widget in self.thumb_widgets.items():
            if widget.geometry().intersects(rect):
                in_rect.add(str(key))
        
        base = getattr(self, '_selection_before_drag', set())
        if modifiers & QtCore.Qt.ControlModifier:
            new_selection = base.symmetric_difference(in_rect)
        elif modifiers & QtCore.Qt.ShiftModifier:
            new_selection = base.union(in_rect)
        else:
            new_selection = in_rect
            
        self.thumb_multi_select = new_selection
        self._refresh_thumb_selection_styles()

    def _thumb_dimensions(self):
        return viewer_thumb_ui._thumb_dimensions(self)

    def _resize_thumbnail_scale(self, delta_px):
        return viewer_thumb_ui._resize_thumbnail_scale(self, delta_px)

    def _create_toolbar(self):
        return main_window_toolbar.create_main_toolbar(self)

    def _update_toolbar_actions(self, enabled: bool):
        return main_window_toolbar.update_toolbar_actions(self, enabled)

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
        # keep toolbar toggle in sync and show ON/OFF text
        try:
            if hasattr(self, 'toolbar_dark_btn'):
                self.toolbar_dark_btn.setChecked(self.dark_mode)
                self.toolbar_dark_btn.setText('dark mode: ON' if self.dark_mode else 'dark mode: OFF')
        except Exception:
            pass
        self.config['dark_mode'] = self.dark_mode; save_config(self.config)
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
        self.recent_dirs = recents[:8]
        self.config["recent_dirs"] = self.recent_dirs
        save_config(self.config)
        self._refresh_recent_dirs_menu()

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

    def _auto_preview_clim(self, arr):
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
            if vmin == vmax:
                return None
            return (vmin, vmax)
        except Exception:
            return None

    def load_folder(self, folder:Path):
        start = time.perf_counter()
        result = viewer_loader.load_folder(self, folder)
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
            if highlight_key == str(file_key):
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

    def _insert_processed_after_source(self, processed_key: str, origin_path: str):
        """Insert processed entry immediately after its origin in self.files."""
        try:
            cur_files = [str(p) for p in self.files]
            if processed_key in cur_files:
                return
            try:
                pos = cur_files.index(str(origin_path))
            except ValueError:
                pos = len(self.files) - 1
            self.files.insert(pos + 1, Path(processed_key))
        except Exception:
            # fallback append
            self.files.append(Path(processed_key))

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
        self.config['display_units_si'] = self.display_units_si
        save_config(self.config)
        if self.last_preview:
            self.show_file_channel(self.last_preview[0], self.last_preview[1])

    def on_unit_relative_toggled(self, checked: bool):
        self.display_units_relative = bool(checked)
        self.config['display_units_relative'] = self.display_units_relative
        save_config(self.config)
        if self.last_preview:
            self.show_file_channel(self.last_preview[0], self.last_preview[1])

    def on_relative_axes_toggled(self, checked: bool):
        self.relative_axes = bool(checked)
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
        if self.last_preview:
            try:
                key, idx = self.last_preview
            except Exception:
                key = idx = None
            self.last_preview = None  # force rebuild of current view
            if key is not None and idx is not None:
                self.show_file_channel(key, idx)
                # After the view is rebuilt, reset zoom so the new extent
                # (relative vs absolute) is applied immediately.
                try:
                    if getattr(self, 'preview_canvas', None):
                        self.preview_canvas._reset_view_zoom()
                except Exception:
                    pass

    def on_scale_bar_toggled(self, checked: bool):
        self.config['show_scale_bar'] = bool(checked)
        save_config(self.config)
        if self.preview_canvas:
            self.preview_canvas.enable_scale_bar(bool(checked))
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
        win.show()
        win.raise_()
        try:
            win.activateWindow()
        except Exception:
            pass

    def on_thumbnail_clicked(self, header_path_str, channel_idx):
        """
        Thumbnail clicked -> preview.
        We no longer populate a persistent per-file inspector list (UI removed).
        Instead we:
          - show the clicked channel in the main preview,
          - update the thumb selection highlight,
          - record the current file header path and channel index so dialogs like
            "Add channel view" can reuse them via current_inspector_* attributes.
        """
        # show preview as before
        self.show_file_channel(header_path_str, channel_idx)
        # highlight selection in thumbnails
        try:
            self.selected_file_for_thumbs = str(header_path_str)
            self._refresh_thumb_selection_styles()
        except Exception:
            pass
        # record the header and channel idx for dialogs that expect them
        key = str(header_path_str)
        self.current_inspector_header = key
        self.current_inspector_channel = int(channel_idx)

    def on_thumbnail_double_clicked(self, header_path_str, channel_idx):
        """
        Double-click a thumbnail: show it in preview and open a popup window (like preview zoom).
        """
        try:
            self.on_thumbnail_clicked(header_path_str, channel_idx)
        except Exception:
            pass
        try:
            views = getattr(self.preview_canvas, "views", None)
            if not views:
                return
            copied = [self._copy_view_for_popup(v) for v in views]
            title = self._friendly_view_title(views[0], default=Path(header_path_str).name if header_path_str else "Preview")
            self._spawn_preview_popup(copied, title=title)
        except Exception:
            pass

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
        return viewer_preview.show_file_channel(self, header_path_str, channel_idx, use_local_cmap=use_local_cmap)

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

    def _jsonify(self, obj):
        if isinstance(obj, dict):
            return {str(k): self._jsonify(v) for k, v in obj.items()}
        if isinstance(obj, (list, tuple)):
            return [self._jsonify(v) for v in obj]
        if isinstance(obj, set):
            return [self._jsonify(v) for v in sorted(obj)]
        if isinstance(obj, Path):
            return str(obj)
        try:
            if isinstance(obj, np.ndarray):
                return obj.tolist()
        except Exception:
            pass
        try:
            if isinstance(obj, np.generic):
                return obj.item()
        except Exception:
            pass
        return obj

    def _collect_session_state(self, session_path: Path):
        data_dir = session_path.parent / f"{session_path.stem}_data"
        processed_dir = data_dir / "processed"
        os.makedirs(processed_dir, exist_ok=True)
        processed = {}
        try:
            if self.last_preview:
                self._store_molecule_overlay(self.last_preview[0])
        except Exception:
            pass
        for key, data in (self._processed_views or {}).items():
            arr_files = {}
            arr_by_channel = data.get("arr_by_channel") or {}
            for ch_idx, arr in arr_by_channel.items():
                fname = f"{key}_ch{ch_idx}.npy"
                np.save(processed_dir / fname, np.asarray(arr))
                arr_files[str(ch_idx)] = str(Path("processed") / fname)
            processed[key] = {
                "source": data.get("source"),
                "op": data.get("op"),
                "label": data.get("label"),
                "channel_idx": data.get("channel_idx"),
                "header": data.get("header"),
                "fds": data.get("fds"),
                "arr_files": arr_files,
            }
        preview_state = {}
        if getattr(self, "preview_canvas", None):
            try:
                preview_state["profile"] = self.preview_canvas.export_profile_state()
            except Exception:
                preview_state["profile"] = None
            try:
                preview_state["angle"] = self.preview_canvas.export_angle_state()
            except Exception:
                preview_state["angle"] = None
            try:
                preview_state["molecules"] = self.preview_canvas.export_molecule_state()
            except Exception:
                preview_state["molecules"] = None
            try:
                preview_state["scale_bar_pos"] = list(getattr(self.preview_canvas, "_scale_bar_pos", (0.94, 0.06)))
                preview_state["scale_bar_settings"] = dict(getattr(self.preview_canvas, "_scale_bar_settings", {}) or {})
            except Exception:
                pass
            preview_state["view_layout"] = getattr(self.preview_canvas, "_view_layout", "grid")
        canvas_state = None
        win = self._canvas_window_ref()
        if win is not None and win.isVisible():
            try:
                canvas_state = win._capture_state()
            except Exception:
                canvas_state = None
        ui_state = {
            "thumb_size_px": getattr(self, "thumb_size_px", 160),
            "thumb_sort": self.thumb_sort_combo.currentText() if hasattr(self, "thumb_sort_combo") else None,
            "thumb_filter": self.thumb_filter_combo.currentText() if hasattr(self, "thumb_filter_combo") else None,
            "thumb_cmap": self.thumb_cmap_combo.currentText() if hasattr(self, "thumb_cmap_combo") else None,
            "preview_cmap": self.preview_cmap_combo.currentText() if hasattr(self, "preview_cmap_combo") else None,
            "channel_index": int(self.channel_dropdown.currentIndex()) if hasattr(self, "channel_dropdown") else 0,
            "mode": int(getattr(self, "current_mode", self.MODE_BROWSE)),
            "show_preview_title": bool(getattr(self, "show_preview_title", True)),
            "show_spectra": bool(getattr(self, "show_spectra", True)),
            "show_preview_spectra": bool(getattr(self, "show_preview_spectra", True)),
            "show_matrix_markers": bool(getattr(self, "show_matrix_markers", True)),
            "show_single_markers": bool(getattr(self, "show_single_markers", True)),
            "compact_markers": bool(getattr(self, "compact_markers", True)),
            "detail_dark_view": bool(getattr(self, "detail_dark_view", False)),
            "detail_grid_view": bool(getattr(self, "detail_grid_view", False)),
            "show_molecules": bool(getattr(self, "show_molecules", True)),
            "relative_axes": bool(getattr(self, "relative_axes", False)),
            "display_units_relative": bool(getattr(self, "display_units_relative", False)),
            "display_units_si": bool(getattr(self, "display_units_si", False)),
            "scale_bar": bool(self.scale_bar_cb.isChecked()) if hasattr(self, "scale_bar_cb") else False,
            "preview_locked": bool(getattr(self, "preview_locked", False)),
        }
        payload = {
            "version": 1,
            "image_folder": str(getattr(self, "last_dir", "") or ""),
            "spectra_folder": str(getattr(self, "spec_folder_path", "") or ""),
            "files": [str(p) for p in (self.files or [])],
            "processed": processed,
            "image_adjustments": self.image_adjustments,
            "thumbnail_filters": self.thumbnail_filters,
            "per_file_channel_cmap": self.per_file_channel_cmap,
            "extra_view_specs": self.extra_view_specs,
            "tags": self.tags,
            "molecule_overlays": self.molecule_overlays,
            "thumb_multi_select": list(getattr(self, "thumb_multi_select", set()) or []),
            "selected_file_for_thumbs": getattr(self, "selected_file_for_thumbs", None),
            "last_preview": self.last_preview,
            "ui": ui_state,
            "preview_state": preview_state,
            "canvas_state": canvas_state,
            "data_dir": data_dir.name,
        }
        return self._jsonify(payload)

    def on_save_session(self):
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Save session",
            str(Path(getattr(self, "last_dir", ".")).joinpath("sxm_session.json")),
            "SXM Session (*.json)",
        )
        if not path:
            return
        session_path = Path(path)
        if session_path.suffix.lower() != ".json":
            session_path = session_path.with_suffix(".json")
        try:
            payload = self._collect_session_state(session_path)
            session_path.parent.mkdir(parents=True, exist_ok=True)
            with open(session_path, "w", encoding="utf-8") as f:
                json.dump(payload, f, indent=2)
            log_status(f"Saved session to {session_path}")
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Save session", f"Unable to save session: {exc}")

    def _apply_session_state(self, payload: dict, session_path: Path):
        if not isinstance(payload, dict):
            return False
        image_folder = payload.get("image_folder") or ""
        if image_folder:
            try:
                self.load_folder(Path(image_folder))
            except Exception:
                pass
        spectra_folder = payload.get("spectra_folder") or ""
        if spectra_folder:
            try:
                self._set_spec_folder(Path(spectra_folder))
            except Exception:
                pass
        data_dir = payload.get("data_dir") or ""
        data_dir = session_path.parent / data_dir if data_dir else session_path.parent
        processed_dir = data_dir / "processed"
        self._processed_views = {}
        for key, entry in (payload.get("processed") or {}).items():
            try:
                arr_by_channel = {}
                for ch_idx, rel_path in (entry.get("arr_files") or {}).items():
                    arr_path = processed_dir / Path(rel_path).name
                    if arr_path.exists():
                        arr_by_channel[int(ch_idx)] = np.load(arr_path, allow_pickle=False)
                header = entry.get("header") or {}
                fds = entry.get("fds") or []
                self._processed_views[str(key)] = {
                    "arr_by_channel": arr_by_channel,
                    "header": header,
                    "fds": fds,
                    "channel_idx": entry.get("channel_idx"),
                    "source": entry.get("source"),
                    "label": entry.get("label"),
                    "op": entry.get("op"),
                }
                self.headers[str(key)] = (header, fds)
            except Exception:
                continue
        session_files = []
        for fp in payload.get("files", []) or []:
            if self._is_processed_key(str(fp)) and str(fp) in self._processed_views:
                session_files.append(Path(str(fp)))
            elif Path(str(fp)).exists():
                session_files.append(Path(str(fp)))
        if session_files:
            self.files = session_files
        self.image_adjustments = payload.get("image_adjustments") or {}
        self.thumbnail_filters = payload.get("thumbnail_filters") or {}
        self.per_file_channel_cmap = payload.get("per_file_channel_cmap") or {}
        self.extra_view_specs = payload.get("extra_view_specs") or []
        self.tags = payload.get("tags") or {}
        self.molecule_overlays = payload.get("molecule_overlays") or {}
        self.thumb_multi_select = set(payload.get("thumb_multi_select") or [])
        self.selected_file_for_thumbs = payload.get("selected_file_for_thumbs")
        pending_preview = payload.get("last_preview")
        self.last_preview = None
        ui = payload.get("ui") or {}
        if hasattr(self, "thumb_sort_combo") and ui.get("thumb_sort"):
            self.thumb_sort_combo.setCurrentText(ui.get("thumb_sort"))
        if hasattr(self, "thumb_filter_combo") and ui.get("thumb_filter"):
            self.thumb_filter_combo.setCurrentText(ui.get("thumb_filter"))
        if hasattr(self, "thumb_cmap_combo") and ui.get("thumb_cmap"):
            self.thumb_cmap_combo.setCurrentText(ui.get("thumb_cmap"))
            try:
                self.on_thumb_cmap_changed(self.thumb_cmap_combo.currentIndex())
            except Exception:
                pass
        if hasattr(self, "preview_cmap_combo") and ui.get("preview_cmap"):
            self.preview_cmap_combo.setCurrentText(ui.get("preview_cmap"))
            try:
                self.on_preview_cmap_changed(self.preview_cmap_combo.currentIndex())
            except Exception:
                pass
        if hasattr(self, "channel_dropdown"):
            idx = int(ui.get("channel_index", self.channel_dropdown.currentIndex()))
            self.channel_dropdown.setCurrentIndex(max(0, idx))
        try:
            self._apply_mode(int(ui.get("mode", self.MODE_BROWSE)), remember=False)
        except Exception:
            pass
        try:
            self.on_show_spectra_toggled(bool(ui.get("show_spectra", True)))
            self.on_show_preview_spectra_toggled(bool(ui.get("show_preview_spectra", True)))
            self.on_show_matrix_markers_toggled(bool(ui.get("show_matrix_markers", True)))
            self.on_show_single_markers_toggled(bool(ui.get("show_single_markers", True)))
            self.on_compact_markers_toggled(bool(ui.get("compact_markers", True)))
            self.on_detail_dark_toggled(bool(ui.get("detail_dark_view", False)))
            self.on_detail_grid_toggled(bool(ui.get("detail_grid_view", False)))
            self.on_show_molecules_toggled(bool(ui.get("show_molecules", True)))
            self.on_relative_axes_toggled(bool(ui.get("relative_axes", False)))
            self.on_unit_relative_toggled(bool(ui.get("display_units_relative", False)))
            self.on_unit_display_toggled(bool(ui.get("display_units_si", False)))
            self.on_scale_bar_toggled(bool(ui.get("scale_bar", False)))
            self._on_toggle_preview_title(bool(ui.get("show_preview_title", True)))
        except Exception:
            pass
        try:
            self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())
        except Exception:
            pass
        if pending_preview:
            try:
                self.show_file_channel(pending_preview[0], pending_preview[1])
            except Exception:
                pass
        preview_state = payload.get("preview_state") or {}
        if getattr(self, "preview_canvas", None):
            try:
                if preview_state.get("view_layout"):
                    self.preview_canvas.set_view_layout(preview_state.get("view_layout"))
            except Exception:
                pass
            try:
                if preview_state.get("profile"):
                    self.preview_canvas.import_profile_state(preview_state.get("profile"), emit=False)
            except Exception:
                pass
            try:
                if preview_state.get("angle"):
                    self.preview_canvas.import_angle_state(preview_state.get("angle"))
            except Exception:
                pass
            try:
                if preview_state.get("molecules") and not self.molecule_overlays:
                    self.preview_canvas.import_molecule_state(preview_state.get("molecules"))
            except Exception:
                pass
            try:
                if preview_state.get("scale_bar_pos"):
                    self.preview_canvas._scale_bar_pos = tuple(preview_state.get("scale_bar_pos"))
                if preview_state.get("scale_bar_settings"):
                    self.preview_canvas._scale_bar_settings = dict(preview_state.get("scale_bar_settings") or {})
            except Exception:
                pass
        try:
            if preview_state.get("molecules") and self.last_preview:
                key = str(self.last_preview[0])
                if key not in self.molecule_overlays:
                    self.molecule_overlays[key] = preview_state.get("molecules")
        except Exception:
            pass
        canvas_state = payload.get("canvas_state")
        if canvas_state:
            try:
                self._on_open_canvas()
                win = self._canvas_window_ref()
                if win:
                    win._restore_state(canvas_state)
            except Exception:
                pass
        return True

    def on_load_session(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Load session",
            str(Path(getattr(self, "last_dir", "."))),
            "SXM Session (*.json)",
        )
        if not path:
            return
        session_path = Path(path)
        try:
            with open(session_path, "r", encoding="utf-8") as f:
                payload = json.load(f)
            if self._apply_session_state(payload, session_path):
                log_status(f"Loaded session from {session_path}")
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Load session", f"Unable to load session: {exc}")

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
    def _populate_canvas_filter_menu(self, menu, canvas, view=None):
        """Populate a context menu with quick filter actions for a preview canvas."""
        if menu is None or canvas is None:
            return
        filt_menu = menu.addMenu("Filters")
        for key, info in FILTER_DEFINITIONS.items():
            act = QtWidgets.QAction(info.get("label", key.title()), filt_menu)
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
            params = {}
            if filter_key in ('highpass', 'lowpass'):
                params['sigma'] = FILTER_DEFINITIONS.get(filter_key, {}).get('default_sigma', 2.0)
            steps = [{'key': filter_key, 'params': params}]
            label = label or FILTER_DEFINITIONS.get(filter_key, {}).get('label', filter_key)
        new_views = []
        for v in canvas.views:
            nv = dict(v)
            base = nv.get('_filter_base_arr')
            if base is None:
                try:
                    base = np.array(nv.get('arr'), copy=True)
                except Exception:
                    base = nv.get('arr')
                nv['_filter_base_arr'] = base
            if not steps:
                nv['arr'] = np.array(base, copy=True) if base is not None else nv.get('arr')
                nv.pop('filter_steps', None)
                nv.pop('filter_label', None)
            else:
                nv['arr'] = self._apply_filter_pipeline(base, steps) if base is not None else nv.get('arr')
                nv['filter_steps'] = steps
                nv['filter_label'] = label
            new_views.append(nv)
        canvas.set_views(new_views, preserve_profiles=True)

    def _open_custom_filter_for_canvas(self, canvas):
        if not canvas or not getattr(canvas, "views", None):
            return
        base_arr = None
        try:
            if canvas.views:
                base_arr = canvas.views[0].get('_filter_base_arr') or canvas.views[0].get('arr')
        except Exception:
            base_arr = None
        dlg = CustomFilterDialog(self, base_arr, self._run_filter_step)
        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            steps = dlg.pipeline_steps()
            label = dlg.pipeline_label()
            self._apply_filter_to_canvas(canvas, pipeline=steps, label=label)

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

    def _on_view_copied(self, view):
        title = view.get('title') or 'View'
        msg = f"Copied '{title}' to clipboard"
        try:
            QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), msg)
        except Exception:
            pass

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
            lo, hi = np.percentile(finite, [pct_low, pct_high])
        except Exception:
            lo, hi = vmin, vmax
        self._apply_clim_to_view(canvas, view, lo, hi)

    def _reset_contrast(self, canvas):
        view = canvas.views[0] if canvas and getattr(canvas, "views", None) else None
        vmin, vmax, _ = self._view_finite_values(view)
        if vmin is None:
            return
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
        if not spec or not self.show_spectra:
            return
        mods = QtCore.Qt.NoModifier
        try:
            if event is not None:
                if hasattr(event, "modifiers"):
                    mods = event.modifiers()
                elif getattr(event, "guiEvent", None) is not None and hasattr(event.guiEvent, "modifiers"):
                    mods = event.guiEvent.modifiers()
        except Exception:
            mods = QtCore.Qt.NoModifier
        file_key = str(spec.get('image_key') or spec.get('path') or '')
        if mods & QtCore.Qt.ShiftModifier:
            self._prime_multi_selection_anchor(spec)
            key = self._spec_identity_key(spec) if spec else None
            already_selected = bool(key and key in getattr(self, "_multi_spec_selection_keys", set()))
            self._toggle_multi_spec_selection(spec)
            added = bool(
                spec and key and not already_selected and key in getattr(self, "_multi_spec_selection_keys", set())
            )
            if added:
                self._append_spec_to_single_popup(spec)
                try:
                    self._highlight_spectrum_entry(spec)
                except Exception:
                    pass
            return
        self._clear_multi_spec_selection()
        self._last_clicked_spec = spec
        force_matrix = bool(mods & QtCore.Qt.ControlModifier)
        is_matrix = self._is_matrix_spec(spec) or (force_matrix and spec.get('matrix_index') is not None)
        if is_matrix and file_key:
            self._open_matrix_explorer_for_file(file_key)
        else:
            self._open_spectroscopy_popup(spec)

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

    def _on_toggle_auto_tags(self, checked: bool):
        self.auto_detect_tags = bool(checked)
        self.config['auto_detect_tags'] = self.auto_detect_tags
        save_config(self.config)
        if checked:
            try:
                self._auto_detect_tags_for_folder()
                # refresh thumbnails to show badges
                self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())
            except Exception:
                pass

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
                self._on_recent_molecules_updated(self.preview_canvas.get_recent_molecule_paths())
                return
        # Fallback: open file dialog
        self.preview_canvas._load_molecule_dialog()

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

    def _set_spec_folder(self, path:Path):
        try:
            self.spec_folder_path = Path(path)
            self.config['spectra_folder'] = str(self.spec_folder_path)
            save_config(self.config)
        except Exception:
            pass
        # if lazy loading is enabled, mark spectros pending; otherwise reload immediately
        if getattr(self, "lazy_spectros_enabled", False):
            self._spectros_pending = True
            self._spectros_loaded = False
            self._update_spectro_stats_label()
        else:
            self._reload_spectros(refresh=True)

    def ensure_spectros_loaded(self, refresh: bool = True):
        """Load spectroscopies on-demand if they were deferred."""
        if self._spectros_loaded:
            return True
        if not self.show_spectra:
            return False
        if getattr(self, "_spectros_loading", False):
            return False
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

    def _reload_spectros(self, refresh=True):
        # unless we complete a successful reload, consider spectra cache stale
        self._spectros_loaded = False
        self._spectros_pending = False
        t_scan_start = time.perf_counter()
        try:
            folder = getattr(self, 'spec_folder_path', None) or self.last_dir
            folder = Path(folder)
        except Exception:
            folder = self.last_dir
        log_status(f"Scanning spectroscopy files in: {folder}")
        if not self.show_spectra:
            self.spectros = []
            self.spectros_by_image = defaultdict(list)
            self._spectro_deferred = set()
            self._clear_multi_spec_selection()
            self._update_spectro_stats_label()
            return
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

    def _assign_spectros_to_images(self):
        spectro_controller._assign_spectros_to_images(self)
        try:
            self.files_with_matrix = {
                key for key, entries in (self.spectros_by_image or {}).items()
                if any(spec.get('matrix_index') is not None for spec in entries)
            }
        except Exception:
            self.files_with_matrix = set()
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
        frac_x = (x - x0) / xspan
        frac_y = (y1 - y) / yspan  # invert so larger y appears lower on the pixmap
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
                    QtCore.QTimer.singleShot(0, lambda: self.ensure_spectros_loaded(refresh=True))
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
        if not file_key:
            return
        widget = self.thumb_widgets.get(str(file_key))
        if widget is None:
            return
        try:
            self.scroll.ensureWidgetVisible(widget)
        except Exception:
            try:
                bar = self.scroll.verticalScrollBar()
                if bar is not None:
                    bar.setValue(widget.y())
            except Exception:
                pass

    def _handle_thumbnail_navigation(self, key, modifiers=QtCore.Qt.NoModifier):
        files = list(getattr(self, 'current_thumb_files', []) or [])
        if not files:
            return False
        cols = max(1, int(getattr(self, 'thumb_grid_columns', 1) or 1))
        selected = str(getattr(self, 'selected_file_for_thumbs', '') or '')
        try:
            current_idx = files.index(selected)
        except ValueError:
            current_idx = -1

        if current_idx < 0:
            new_idx = 0
        else:
            new_idx = current_idx
            if key == QtCore.Qt.Key_Left:
                if new_idx > 0:
                    new_idx -= 1
            elif key == QtCore.Qt.Key_Right:
                if new_idx < len(files) - 1:
                    new_idx += 1
            elif key == QtCore.Qt.Key_Up:
                if new_idx - cols >= 0:
                    new_idx -= cols
                else:
                    new_idx = 0
            elif key == QtCore.Qt.Key_Down:
                if new_idx + cols < len(files):
                    new_idx += cols
                else:
                    new_idx = len(files) - 1
            else:
                return False
        if new_idx == current_idx:
            # If nothing selected yet, ensure first entry is highlighted.
            if current_idx == -1:
                return self._activate_thumbnail_by_index(0)
            return False
        return self._activate_thumbnail_by_index(new_idx)

    def _activate_thumbnail_by_index(self, index):
        files = list(getattr(self, 'current_thumb_files', []) or [])
        if not files or not (0 <= index < len(files)):
            return False
        file_key = files[index]
        current_highlight = getattr(self, '_highlighted_spec', None)
        if current_highlight:
            try:
                highlight_path = str(current_highlight.get('image_key') or current_highlight.get('path') or '')
            except Exception:
                highlight_path = ''
            if highlight_path and highlight_path != file_key:
                self._highlight_spectrum_entry(None)
        self._clear_thumb_multi_selection(update_styles=False)
        label = getattr(self, '_thumb_labels', {}).get(file_key)
        if label is not None:
            try:
                channel_idx = int(label.property("channel_index") or 0)
            except Exception:
                channel_idx = self.channel_dropdown.currentIndex()
        else:
            channel_idx = self.channel_dropdown.currentIndex()
        self.on_thumbnail_clicked(file_key, channel_idx)
        self.last_thumb_anchor = str(file_key)
        self._scroll_to_thumbnail(file_key)
        return True

    def _focus_first_matrix_dataset(self):
        matrix_files = list(getattr(self, 'files_with_matrix', set()) or [])
        if not matrix_files:
            return
        target = None
        for path in getattr(self, 'current_thumb_files', []):
            if path in matrix_files:
                target = path
                break
        if target is None:
            target = matrix_files[0]
        self._scroll_to_thumbnail(target)
        self.selected_file_for_thumbs = target
        self._refresh_thumb_selection_styles()

    def _update_matrix_summary_banner(self):
        label = getattr(self, 'matrix_summary_label', None)
        if label is None:
            return
        matrix_count = len(getattr(self, 'matrix_datasets', {}) or {})
        if matrix_count <= 0:
            label.hide()
            return
        noun = "Matrix dataset" if matrix_count == 1 else "Matrix datasets"
        label.setText(f"{noun}: {matrix_count} · click to focus")
        label.setToolTip("Click to jump to the first thumbnail containing a matrix spectroscopy grid.")
        label.show()

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
        if mods & QtCore.Qt.ShiftModifier:
            self._prime_multi_selection_anchor(spec)
            key = self._spec_identity_key(spec) if spec else None
            already_selected = bool(key and key in getattr(self, "_multi_spec_selection_keys", set()))
            self._toggle_multi_spec_selection(spec)
            added = bool(
                spec and key and not already_selected and key in getattr(self, "_multi_spec_selection_keys", set())
            )
            if added:
                self._append_spec_to_single_popup(spec)
                try:
                    self._highlight_spectrum_entry(spec)
                except Exception:
                    pass
            return True
        self._clear_multi_spec_selection()
        self._last_clicked_spec = spec
        force_matrix = bool(mods & QtCore.Qt.ControlModifier) and spec and spec.get('matrix_index') is not None
        is_matrix = hit_info.get('kind') == 'matrix' or self._is_matrix_spec(spec)
        if (is_matrix or force_matrix) and file_key:
            self._open_matrix_explorer_for_file(file_key)
        else:
            self._open_spectroscopy_popup(spec)
        try:
            self._highlight_spectrum_entry(spec)
        except Exception:
            pass
        return True

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
                QtWidgets.QToolTip.showText(label_widget.mapToGlobal(event.pos()), tooltip)
                return True
        QtWidgets.QToolTip.hideText()
        return False

    def _open_spectroscopy_popup(self, spec):
        if not self._spectros_loaded:
            self.ensure_spectros_loaded(refresh=False)
        return spectro_popups._open_spectroscopy_popup(self, spec)

    def _ensure_single_spectro_popup(self, spec):
        """Raise an existing single spectroscopy popup or open a new one."""
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

    def _single_popup_for_key(self, key):
        if not key or not getattr(self, "_spectro_popups", None):
            return None
        for dlg in list(self._spectro_popups):
            dlg_spec = getattr(dlg, "spec", None)
            if dlg_spec and self._spec_identity_key(dlg_spec) == key:
                if getattr(dlg, "isVisible", None):
                    try:
                        if dlg.isVisible():
                            return dlg
                    except Exception:
                        continue
                else:
                    return dlg
        return None

    def _active_single_popup(self):
        anchor = getattr(self, "_multi_single_popup_anchor", None)
        if not anchor:
            return None
        dlg = self._single_popup_for_key(anchor)
        if dlg is None:
            self._multi_single_popup_anchor = None
        return dlg

    def _append_spec_to_single_popup(self, spec):
        if not spec:
            return
        key = self._spec_identity_key(spec)
        if not key:
            return
        dlg = self._active_single_popup()
        if dlg is None:
            dlg = self._ensure_single_spectro_popup(spec)
            if dlg:
                self._multi_single_popup_anchor = key
            return
        if key == self._multi_single_popup_anchor:
            return
        if hasattr(dlg, "add_external_spectrum"):
            try:
                dlg.add_external_spectrum(spec)
            except Exception:
                pass

    def _prime_multi_selection_anchor(self, current_spec):
        """If user previously clicked a spec without Shift, use it as the first multi selection."""
        if self._multi_spec_selection or not getattr(self, "_last_clicked_spec", None):
            return
        candidate = self._last_clicked_spec
        if candidate is None:
            return
        if current_spec and self._spec_identity_key(candidate) == self._spec_identity_key(current_spec):
            return
        key = self._spec_identity_key(candidate)
        if not key or key in getattr(self, "_multi_spec_selection_keys", set()):
            self._last_clicked_spec = None
            return
        self._multi_spec_selection.append(candidate)
        self._multi_spec_selection_keys.add(key)
        self._update_spec_selection_label()
        self._last_clicked_spec = None

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
            act = QtWidgets.QAction(info['label'], menu)
            if info.get('needs_gaussian') and not _gaussian_available():
                act.setEnabled(False)
                act.setToolTip("Requires scipy or OpenCV.")
            act.triggered.connect(lambda _, k=key, paths=list(targets): self._apply_filter_to_paths(paths, k))
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
        copy_svg_act = QtWidgets.QAction("Copy selected as SVG (current view)", menu)
        copy_svg_act.triggered.connect(lambda: self.copy_selected_as_svg(targets))
        menu.addAction(copy_svg_act)

        export_png_act = QtWidgets.QAction("Export PNGs...", menu)
        export_png_act.triggered.connect(self.on_export_pngs)
        menu.addAction(export_png_act)
        export_xyz_act = QtWidgets.QAction("Export XYZ...", menu)
        export_xyz_act.triggered.connect(self.on_export_xyz_files)
        menu.addAction(export_xyz_act)
        adjust_act = QtWidgets.QAction("Adjust image...", menu)
        adjust_act.triggered.connect(self.on_adjust_image)
        menu.addAction(adjust_act)

        menu.addSeparator()
        overlay_act = QtWidgets.QAction("Show spectroscopy overlays", menu)
        overlay_act.setCheckable(True)
        overlay_act.setChecked(self.show_spectra)
        overlay_act.triggered.connect(self.on_show_spectra_toggled)
        menu.addAction(overlay_act)
        glow_act = QtWidgets.QAction("Spectro highlight glow", menu)
        glow_act.setCheckable(True)
        glow_act.setChecked(getattr(self, "spectro_highlight_glow", True))
        glow_act.triggered.connect(self.on_toggle_highlight_glow)
        menu.addAction(glow_act)
        marker_menu = menu.addMenu("Marker style")
        self._populate_marker_style_menu(marker_menu)

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

    def _apply_filter_to_paths(self, paths, filter_key=None, pipeline=None, label=None):
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
            params = {}
            if filter_key in ('highpass', 'lowpass'):
                params['sigma'] = FILTER_DEFINITIONS.get(filter_key, {}).get('default_sigma', 2.0)
            step = {'key': filter_key, 'params': params}
            spec_steps = [step]
            spec_label = FILTER_DEFINITIONS.get(filter_key, {}).get('label', filter_key)
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
        base_arr = None
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
                    base_arr = normalize_unit_and_data(arr, fd.get('PhysUnit',''))[1]
        except Exception:
            base_arr = None
        dlg = CustomFilterDialog(self, base_arr, self._run_filter_step)
        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            pipeline = dlg.pipeline_steps()
            if pipeline:
                self._apply_filter_to_paths(paths, pipeline=pipeline, label=dlg.pipeline_label())

    def _toggle_thumb_multi_selection(self, file_path):
        return viewer_thumb_ui._toggle_thumb_multi_selection(self, file_path)

    def _clear_thumb_multi_selection(self, update_styles=True):
        return viewer_thumb_ui._clear_thumb_multi_selection(self, update_styles=update_styles)

    def _spec_identity_key(self, spec):
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
        if not spec:
            return
        key = self._spec_identity_key(spec) or str(Path(spec.get('path')))
        if key in self._multi_spec_selection_keys:
            self._multi_spec_selection = [s for s in self._multi_spec_selection if self._spec_identity_key(s) != key]
            self._multi_spec_selection_keys.remove(key)
        else:
            self._multi_spec_selection.append(spec)
            self._multi_spec_selection_keys.add(key)
        self._update_spec_selection_label()
        if not self._multi_spec_selection:
            self._multi_single_popup_anchor = None
        if len(self._multi_spec_selection) >= 2:
            self._open_multi_spectroscopy_popup()

    def _update_spec_selection_label(self):
        count = len(self._multi_spec_selection)
        if hasattr(self, 'spec_selection_label'):
            self.spec_selection_label.setText(f"Spectra selected: {count}")

    def _clear_multi_spec_selection(self):
        self._multi_spec_selection = []
        self._multi_spec_selection_keys = set()
        self._multi_single_popup_anchor = None
        self._last_clicked_spec = None
        for dlg in list(self._multi_spectro_popups):
            try:
                dlg.close()
            except Exception:
                pass
        self._multi_spectro_popups = []
        self._update_spec_selection_label()

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
        ref_gray = (ref_gray - ref_gray.min()) / max(ref_gray.ptp(), 1e-6)
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
            target = (target - target.min()) / max(target.ptp(), 1e-6)
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

    def _create_virtual_channel_copies(self, paths, channel_idx=None):
        """Create virtual copies of selected images for a specific channel."""
        if not paths:
            return
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
                return
            channel_idx = dlg.intValue()
        added = 0
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
                self._insert_processed_after_source(key, p)
                added += 1
            except Exception:
                continue
        if added:
            self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())

    def _create_virtual_crop_view(self, view):
        """Create a virtual copy from a cropped preview view (single channel)."""
        if not view:
            return
        path = view.get("path") or (view.get("meta") or {}).get("path")
        arr = view.get("arr")
        ch_idx = view.get("channel_idx") or (view.get("meta") or {}).get("channel_idx")
        if path is None or arr is None:
            return
        try:
            header, fds = self.headers.get(str(path), (None, None))
            if header is None or fds is None:
                header, fds = parse_header(Path(path))
            if not fds:
                return
            ch_idx = int(ch_idx) if ch_idx is not None else 0
            arr_by_channel = {ch_idx: np.array(arr, copy=True)}
            fds_new = [dict(fd) for fd in fds]
            for i, fd_new in enumerate(fds_new):
                fd_new['FileName'] = f"{Path(path).name}_crop_ch{i}"
                fd_new['Caption'] = f"{fd_new.get('Caption') or Path(path).name} [crop]"
            header_new = dict(header)
            header_new['xPixel'] = arr.shape[1]
            header_new['yPixel'] = arr.shape[0]
            key = self._make_processed_key(str(path), op="crop", channel_idx=ch_idx)
            self._processed_views[key] = {
                'arr_by_channel': arr_by_channel,
                'header': header_new,
                'fds': fds_new,
                'channel_idx': ch_idx,
                'source': str(path),
                'label': "[crop]",
                'op': 'crop',
            }
            self.headers[key] = (header_new, fds_new)
            self._insert_processed_after_source(key, str(path))
            self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())
        except Exception:
            return

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
        self.config["preview_locked"] = self.preview_locked; save_config(self.config)
        if hasattr(self, "preview_detach_btn"):
            try:
                self.preview_detach_btn.setEnabled(not self.preview_locked)
            except Exception:
                pass
        if self.preview_locked and getattr(self, "preview_detached", False):
            self._attach_preview()

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
            if hasattr(self, "preview_detach_btn"):
                try:
                    self.preview_detach_btn.setText("Attach")
                except Exception:
                    pass
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
            if hasattr(self, "preview_detach_btn"):
                try:
                    self.preview_detach_btn.setText("Detach")
                except Exception:
                    pass
        except Exception:
            pass

    def on_dark_mode_toggled(self, checked: bool):
        self.dark_mode = bool(checked)
        # keep toolbar toggle in sync and show ON/OFF text
        try:
            if hasattr(self, 'toolbar_dark_btn'):
                self.toolbar_dark_btn.setChecked(self.dark_mode)
                self.toolbar_dark_btn.setText('dark mode: ON' if self.dark_mode else 'dark mode: OFF')
        except Exception:
            pass
        # update toolbar combobox and label styles to match dark/light theme
        try:
            combo_style = "QComboBox { background-color: #1f1f1f; border: 1px solid #444444; color: #f0f0f0; padding: 4px; }" if self.dark_mode else ""
            label_style = "padding-left:8px; padding-right:4px; color: #e6e6e6;" if self.dark_mode else "padding-left:8px; padding-right:4px; color: #202020;"
            if hasattr(self, 'thumb_cmap_combo'):
                self.thumb_cmap_combo.setStyleSheet(combo_style)
            if hasattr(self, 'preview_cmap_combo'):
                self.preview_cmap_combo.setStyleSheet(combo_style)
            if hasattr(self, 'thumb_cmap_label'):
                self.thumb_cmap_label.setStyleSheet(label_style)
            if hasattr(self, 'preview_cmap_label'):
                self.preview_cmap_label.setStyleSheet(label_style)
        except Exception:
            pass
        self.config['dark_mode'] = self.dark_mode; save_config(self.config)
        self._apply_dark_mode(self.dark_mode)
        if self.last_preview:
            self.show_file_channel(self.last_preview[0], self.last_preview[1])

    # ---------- control callbacks ----------
    def on_channel_dropdown_changed(self, idx):
        self.last_channel_index = int(idx); self.config['last_channel_index'] = self.last_channel_index; save_config(self.config)
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
            if hasattr(self, "spectro_overlay_act"):
                self.spectro_overlay_act.blockSignals(True)
                self.spectro_overlay_act.setChecked(self.show_spectra)
                self.spectro_overlay_act.blockSignals(False)
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

    def on_show_preview_spectra_toggled(self, checked: bool):
        self.show_preview_spectra = bool(checked)
        self.config['show_preview_spectra'] = self.show_preview_spectra; save_config(self.config)
        try:
            if hasattr(self, "show_spectra_cb"):
                self.show_spectra_cb.blockSignals(True)
                self.show_spectra_cb.setChecked(self.show_preview_spectra)
                self.show_spectra_cb.blockSignals(False)
        except Exception:
            pass
        if self.last_preview:
            self.show_file_channel(self.last_preview[0], self.last_preview[1])

    def on_toggle_highlight_glow(self, checked: bool):
        self.spectro_highlight_glow = bool(checked)
        self.config['spectro_highlight_glow'] = self.spectro_highlight_glow; save_config(self.config)
        act = getattr(self, 'highlight_glow_act', None)
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
        self._reload_spectros(refresh=True)

    def on_spectro_force_single_toggled(self, checked: bool):
        self.spectro_force_single_mode = bool(checked)
        self.config["spectro_force_single_mode"] = self.spectro_force_single_mode
        save_config(self.config)
        self._reload_spectros(refresh=True)

    def on_show_matrix_markers_toggled(self, checked: bool):
        self.show_matrix_markers = bool(checked)
        self.config['show_matrix_markers'] = self.show_matrix_markers; save_config(self.config)
        self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())
        if self.last_preview:
            self.show_file_channel(self.last_preview[0], self.last_preview[1])
        self._schedule_marker_refresh()
        act = getattr(self, 'matrix_markers_act', None)
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
        act = getattr(self, 'single_markers_act', None)
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
        act = getattr(self, 'compact_markers_act', None)
        if act is not None:
            act.blockSignals(True)
            act.setChecked(self.compact_markers)
            act.blockSignals(False)

    def on_detail_dark_toggled(self, checked: bool):
        self.detail_dark_view = bool(checked)
        self.config['detail_dark_view'] = self.detail_dark_view; save_config(self.config)
        self._apply_detail_view_theme()

    def on_detail_grid_toggled(self, checked: bool):
        self.detail_grid_view = bool(checked)
        self.config['detail_grid_view'] = self.detail_grid_view; save_config(self.config)
        self._apply_detail_view_theme()

    def on_show_molecules_toggled(self, checked: bool):
        self.show_molecules = bool(checked)
        self.config['show_molecules'] = self.show_molecules; save_config(self.config)
        try:
            if getattr(self, "preview_canvas", None):
                self.preview_canvas.set_show_molecules(self.show_molecules)
        except Exception:
            pass
        for canv in getattr(self, '_popup_canvases', []):
            try:
                canv.set_show_molecules(self.show_molecules)
            except Exception:
                continue
        act = getattr(self, 'molecules_act', None)
        if act is not None:
            act.blockSignals(True)
            act.setChecked(self.show_molecules)
            act.blockSignals(False)

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
        self.show_crop_history_overlay = bool(checked)
        self.config['show_crop_history_overlay'] = self.show_crop_history_overlay; save_config(self.config)
        canvases = [getattr(self, 'preview_canvas', None)] + list(getattr(self, '_popup_canvases', []))
        for canv in canvases:
            if canv:
                try:
                    canv.show_fixed_crop_history(self.show_crop_history_overlay)
                except Exception:
                    continue
        act = getattr(self, 'crop_history_act', None)
        if act is not None:
            act.blockSignals(True)
            act.setChecked(self.show_crop_history_overlay)
            act.blockSignals(False)
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

    def _sync_quick_crop_template_controls(self):
        controller = getattr(self, "quick_crop_controller", None)
        if controller:
            controller.sync_template_controls()

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
