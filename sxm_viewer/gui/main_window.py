"""Main Qt widget implementing the SXM Viewer."""
from __future__ import annotations

import math
import re
import threading
from collections import OrderedDict, defaultdict
from datetime import datetime
from pathlib import Path

import numpy as np
from matplotlib import colormaps
from matplotlib.figure import Figure
from PyQt5 import QtCore, QtGui, QtWidgets
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
        self.setWindowTitle("SXM Viewer")
        self.resize(*MAIN_WINDOW_SIZE)

        log_status("Loading configuration...")
        self.config = load_config()
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
        self.thumb_size_px = int(self.config.get("thumb_size_px", 160))
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
        self._display_defaults = {
            'show_matrix_markers': True,
            'show_single_markers': True,
            'compact_markers': True,
            'detail_dark_view': bool(self.dark_mode),
            'detail_grid_view': False,
        }
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
        self._temp_reveal = set()
        self.spectro_dock = None
        self._spectro_browser_entries = []

        self.files = []
        self.headers = {}
        self.thumb_cache = {}
        self._thumb_data_cache = {}
        self._topo_stats_cache = {}
        self._channel_data_cache = OrderedDict()
        self._channel_cache_lock = threading.Lock()
        self._filtered_channel_cache = OrderedDict()
        self._filtered_cache_lock = threading.Lock()
        self._thumb_labels = {}
        self._thumb_generation = 0
        self._thumb_data_lock = threading.Lock()
        self._thumb_threadpool = QtCore.QThreadPool()
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
        self._spectro_cache = {}
        self._spectro_deferred = set()
        # spectro_eager_limit: 0 means no deferral; otherwise minimum of 5000 to avoid accidental truncation
        limit_cfg = int(self.config.get("spectro_eager_limit", 0))
        self.spectro_eager_limit = 0 if limit_cfg <= 0 else max(5000, limit_cfg)
        self.image_time_index = {}
        self._spectro_popups = []
        self._popup_refs = []
        self._multi_spectro_popups = []
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
        
        # Purge config button
        self.purge_config_btn = QtWidgets.QPushButton('Purge config')
        tag_h.addWidget(self.purge_config_btn)
        tag_h.addWidget(self.tag_ch_btn); tag_h.addWidget(self.tag_cc_btn); tag_h.addWidget(self.untag_btn)
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
        preview_header.addWidget(display_strip)
        # Canvas launch button moved to the main toolbar for prominence.
        preview_panel_layout.addLayout(preview_header)
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
        self.preview_canvas.set_copy_feedback_handler(self._on_view_copied)
        preview_panel_layout.addWidget(self.preview_canvas, 1)
        self.preview_value_label = QtWidgets.QLabel("Value: --")
        preview_panel_layout.addWidget(self.preview_value_label)
        self.angle_value_label = QtWidgets.QLabel("Angle: --")
        preview_panel_layout.addWidget(self.angle_value_label)
        preview_panel.setLayout(preview_panel_layout)
        self.preview_canvas.set_value_callback(self._on_preview_value)
        self.preview_canvas.set_spectra_click_callback(self._on_preview_spec_click)
        self.preview_canvas.enable_scale_bar(self.scale_bar_cb.isChecked())
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
        self.show_spectra_cb.toggled.connect(self.on_show_spectra_toggled)
        if hasattr(self, "grid_as_matrix_cb"):
            self.grid_as_matrix_cb.toggled.connect(self.on_spectro_grid_as_matrix_toggled)
        if hasattr(self, "force_single_cb"):
            self.force_single_cb.toggled.connect(self.on_spectro_force_single_toggled)
        self.show_matrix_spectra_btn.clicked.connect(self.on_show_matrix_spectro_viewer)
        self.clear_spec_selection_btn.clicked.connect(self.on_clear_spec_selection)
        self.export_selected_btn.clicked.connect(self.on_export_selected_same_view)
        self.tag_ch_btn.clicked.connect(lambda: self.on_manual_tag('constant-height'))
        self.tag_cc_btn.clicked.connect(lambda: self.on_manual_tag('constant-current'))
        self.untag_btn.clicked.connect(lambda: self.on_manual_tag(None))

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
        base = self.frameGeometry().topLeft() if self.isVisible() else QtGui.QCursor.pos()
        # incrementing counter avoids stacking even if dialogs close quickly
        self._popup_counter = (self._popup_counter + 1) % 12
        idx = self._popup_counter
        pos = base + QtCore.QPoint(offset * (idx % 6), offset * (idx % 6))
        # clamp to screen
        x = max(avail.left(), min(pos.x(), avail.right() - 200))
        y = max(avail.top(), min(pos.y(), avail.bottom() - 150))
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

    def _open_spectro_summary_for_file(self, file_key, show_mode="single"):
        return main_window_spectro.open_spectro_summary_for_file(self, file_key, show_mode=show_mode)

    def _open_matrix_explorer_for_file(self, file_key):
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

    # ---------- Spectro browser dock ----------
    def _ensure_spectro_dock(self):
        return main_window_spectro.ensure_spectro_dock(self)

    def open_spectro_browser(self, entries=None):
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
        # Handle Ctrl+Wheel over the thumbnails to resize thumbnails
        if obj in (getattr(self, '_thumb_viewport', None),
                   getattr(self, 'thumb_container', None),
                   getattr(self, 'scroll', None)) and event.type() == QtCore.QEvent.Wheel:
            if event.modifiers() & QtCore.Qt.ControlModifier:
                delta = event.angleDelta().y() or event.pixelDelta().y()
                if delta != 0:
                    step = 16 if delta > 0 else -16
                    self._resize_thumbnail_scale(step)
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
            # allow normal resize processing to continue
            return False
        return super().eventFilter(obj, event)

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

    def load_folder(self, folder:Path):
        return viewer_loader.load_folder(self, folder)

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
            vals = None
            samples = _sample_channel_values_for_tagging(key, hdr, fd, CH_SAMPLE_POINTS)
            if samples is not None and samples.size:
                arr_input = samples if samples.ndim > 1 else samples.reshape(1, -1)
                _, arr_nm = normalize_unit_and_data(arr_input, fd.get('PhysUnit',''))
                vals = np.asarray(arr_nm, dtype=float).ravel()
            else:
                try:
                    raw_arr = self._get_channel_array(key, topo_idx, hdr, fd)
                except Exception:
                    continue
                _, arr_nm = normalize_unit_and_data(raw_arr, fd.get('PhysUnit',''))
                vals = np.asarray(arr_nm, dtype=float).ravel()
            vals = vals[np.isfinite(vals)]
            if vals.size == 0:
                continue
            sample_count = min(CH_SAMPLE_POINTS, vals.size)
            if vals.size <= sample_count:
                samples = vals
            else:
                idx = np.linspace(0, vals.size - 1, sample_count, dtype=int)
                samples = vals[idx]

            sample_range = float(np.nanmax(samples) - np.nanmin(samples)) if samples.size else float('inf')
            if sample_range <= CH_EQUALITY_TOL_NM:
                median_nm = float(np.nanmedian(samples)) if samples.size else None
                abs_pm = int(round(median_nm * 1000.0)) if median_nm is not None else None
                self.tags[key] = {'tag': 'constant-height', 'abs_z_pm': abs_pm}
            else:
                self.tags[key] = {'tag': 'constant-current'}

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

    def _decorate_thumbnail_pixmap(self, pix, file_key, channel_idx, header, fds):
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
        if header and fds and 0 <= channel_idx < len(fds):
            try:
                xpix = int(header.get('xPixel', 128))
                ypix = int(header.get('yPixel', xpix))
                
                marker_defs = self._render_spectroscopy_overlays(pix, header, str(file_key), xpix, ypix)
            except Exception:
                marker_defs = []
        return marker_defs

    def _schedule_thumbnail_job(self, file_key, channel_idx, header, fd, thumb_w, thumb_h, cmap_name, generation):
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
        pix = base_pix.copy()
        header, fds = self.headers.get(str(file_key), (None, None))
        markers = self._decorate_thumbnail_pixmap(pix, file_key, channel_idx, header, fds)
        label.setPixmap(pix)
        label.setProperty("spec_markers", markers)

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
        try:
            log_status(f"Thumbnail failed for {file_key}: {error}")
        except Exception:
            pass

    def _get_thumbnail_array(self, file_key, channel_idx, header, fd, thumb_w, thumb_h):
        return viewer_thumbnails._get_thumbnail_array(self, file_key, channel_idx, header, fd, thumb_w, thumb_h)

    def _thumbnail_data_key(self, file_key, channel_idx, fd, thumb_w, thumb_h):
        return viewer_thumbnails._thumbnail_data_key(self, file_key, channel_idx, fd, thumb_w, thumb_h)

    def _invalidate_thumbnail_cache(self, paths=None):
        return viewer_thumbnails._invalidate_thumbnail_cache(self, paths=paths)

    def _channel_cache_key(self, file_key, channel_idx, fd):
        fname = fd.get('FileName')
        if not fname:
            raise ValueError("Missing FileName for channel")
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
                return
            parent_dirs = {str(Path(p).parent) for p in paths}
            to_remove = [k for k in self._channel_data_cache.keys() if str(Path(k[0]).parent) in parent_dirs]
            for k in to_remove:
                self._channel_data_cache.pop(k, None)
        self._invalidate_filtered_cache(paths)

    def _invalidate_filtered_cache(self, paths=None):
        with self._filtered_cache_lock:
            if not paths:
                self._filtered_channel_cache.clear()
                self._frame_real_pixmap_cache.clear()
                return
            parent_dirs = {str(Path(p).parent) for p in paths}
            to_remove = [k for k in self._filtered_channel_cache.keys()
                        if str(Path(k[0][0]).parent) in parent_dirs]
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
        if self.last_preview:
            self.show_file_channel(self.last_preview[0], self.last_preview[1])

    def on_scale_bar_toggled(self, checked: bool):
        self.config['show_scale_bar'] = bool(checked)
        save_config(self.config)
        if self.preview_canvas:
            self.preview_canvas.enable_scale_bar(bool(checked))
        if self.last_preview:
            self.show_file_channel(self.last_preview[0], self.last_preview[1])

    # removed size change handler

    def _parse_header_datetime(self, header):
        return viewer_loader._parse_header_datetime(self, header)

    def _header_datetime_dt(self, header, path):
        try:
            ts = float(self._parse_header_datetime(header or {}))
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
            try:
                markers = self._decorate_thumbnail_pixmap(pix, file_key, channel_idx, header, fds)
            except Exception:
                markers = []
            label.setPixmap(pix)
            label.setProperty("spec_markers", markers)

    def _make_thumb_press_handler(self, label_widget):
        return viewer_thumb_ui._make_thumb_press_handler(self, label_widget)

    def _make_thumb_release_handler(self, label_widget):
        return viewer_thumb_ui._make_thumb_release_handler(self, label_widget)

    def _make_thumb_move_handler(self, label_widget):
        return viewer_thumb_ui._make_thumb_move_handler(self, label_widget)

    def _on_open_canvas(self):
        if self._canvas_window is None or not self._canvas_window.isVisible():
            self._canvas_window = ExperimentalCanvasWindow(self, self)
        self._canvas_window.show()
        self._canvas_window.raise_()
        self._canvas_window.activateWindow()

    def _ensure_canvas_for_drag(self):
        """Open the canvas window as a drop target during thumbnail drags."""
        if self._canvas_window is None or not self._canvas_window.isVisible():
            self._canvas_window = ExperimentalCanvasWindow(self, self)
        self._canvas_window.show()
        self._canvas_window.raise_()

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

    # NOTE: removed on_file_channel_selected and on_file_channel_show_clicked
    # These functions supported the removed per-file inspector UI. The same "show channel"
    # functionality is available via the thumbnail UI and the "Add channel view" dialog.

    # ---------- preview + metadata ---------- 
    def show_file_channel(self, header_path_str, channel_idx:int, use_local_cmap=False):
        return viewer_preview.show_file_channel(self, header_path_str, channel_idx, use_local_cmap=use_local_cmap)

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

    def _apply_filter_pipeline(self, arr, steps):
        result = np.asarray(arr, dtype=float)
        for step in steps:
            result = self._run_filter_step(result, step)
        return result

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
            self._toggle_multi_spec_selection(spec)
            return
        self._clear_multi_spec_selection()
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
                    phys = (fd.get('PhysUnit','') or '').lower()
                    arr_nm = arr
                    hist, edges = np.histogram(arr_nm.ravel(), bins=200)
                    imax = int(np.argmax(hist))
                    mode_val = 0.5*(edges[imax] + edges[imax+1])
                    abs_pm = int(round(mode_val * 1000.0))
                    info['abs_z_pm'] = abs_pm
                except Exception:
                    info['abs_z_pm'] = None
            self.tags[key] = info
        self.config['tags'] = self.tags; save_config(self.config)
        # refresh thumbnails & preview (so badges/metadata update)
        self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())
        if self.last_preview: self.show_file_channel(self.last_preview[0], self.last_preview[1])

    # ---------- Spectroscopy helpers ----------
    def on_load_molecule(self):
        if self.preview_canvas:
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
        self._reload_spectros(refresh=True)

    def _reload_spectros(self, refresh=True):
        # unless we complete a successful reload, consider spectra cache stale
        self._spectros_loaded = False
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
        if spec_stats:
            total_entries = spec_stats.get('total_specs', len(self.spectros))
            single_files = spec_stats.get('single_dat_files', 0)
            single_entries = spec_stats.get('single_entries', single_files)
            matrix_files = spec_stats.get('matrix_dat_files', 0)
            matrix_entries = spec_stats.get('matrix_specs', 0)
            # keep stats for UI but avoid duplicate terminal spam (loader already logged)
        else:
            log_status(f"Loaded {len(self.spectros)} spectroscopy entries")
        self._assign_spectros_to_images()
        self.matrix_spectros = [spec for spec in self.spectros if spec.get('matrix_index') is not None]
        self._clear_multi_spec_selection()
        self._update_spectro_stats_label(spec_stats)
        self._spectros_loaded = True
        self._update_matrix_summary_banner()
        if refresh:
            self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())
            if self.last_preview:
                self.show_file_channel(self.last_preview[0], self.last_preview[1])

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

    def _map_spec_to_pixels(self, spec, header, xpix, ypix, file_key=None):
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
                return fallback
            return self._map_spec_by_grid(spec, xpix, ypix)
        frac_x = (x - x0) / xspan
        frac_y = (y1 - y) / yspan  # invert so larger y appears lower on the pixmap
        if not (0.0 <= frac_x <= 1.0 and 0.0 <= frac_y <= 1.0):
            # try spectroscopy cloud extent before clamping/grid
            fallback = self._map_spec_by_spec_extent(file_key, spec, xpix, ypix)
            if fallback is not None:
                return fallback
            grid_pt = self._map_spec_by_grid(spec, xpix, ypix)
            if grid_pt is not None:
                return grid_pt
            frac_x = min(max(frac_x, 0.0), 1.0)
            frac_y = min(max(frac_y, 0.0), 1.0)
        cols = max(1, int(xpix) - 1)
        rows = max(1, int(ypix) - 1)
        col = frac_x * cols
        row = frac_y * rows
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

    def _render_spectroscopy_overlays(self, pixmap, header, file_key, xpix, ypix, reveal_points_override=None, selected_spec=None, entries_override=None, matrix_as_points=False):
        """Render spectroscopy markers directly on the thumbnail pixmap."""
        if not self.show_spectra and not reveal_points_override:
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
        )

    def _matrix_bbox_pixels(self, m_specs, header, xpix, ypix, w_scale, h_scale, file_key=None):
        xs = []
        ys = []
        for idx, spec in enumerate(m_specs, 1):
            c = self._map_spec_to_pixels(spec, header, xpix, ypix, file_key)
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
        max_h = max(1.0, (max(ypix - 1, 1)) * h_scale)
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
        for info in markers:
            rect = info.get('rect')
            if rect and rect.contains(x, y):
                if info.get('label') == 'badge':
                    self._open_spectro_summary_for_file(file_key)
                    return True
                mods = event.modifiers() if event is not None else QtCore.Qt.NoModifier
                if mods & QtCore.Qt.ShiftModifier:
                    self._toggle_multi_spec_selection(info.get('spec'))
                else:
                    self._clear_multi_spec_selection()
                    spec = info.get('spec')
                    force_matrix = bool(mods & QtCore.Qt.ControlModifier) and spec and spec.get('matrix_index') is not None
                    is_matrix = info.get('kind') == 'matrix' or self._is_matrix_spec(spec)
                    if (is_matrix or force_matrix) and file_key:
                        self._open_matrix_explorer_for_file(file_key)
                    else:
                        self._open_spectroscopy_popup(spec)
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
                QtWidgets.QToolTip.showText(label_widget.mapToGlobal(event.pos()), tooltip)
                return True
        QtWidgets.QToolTip.hideText()
        return False

    def _open_spectroscopy_popup(self, spec):
        return spectro_popups._open_spectroscopy_popup(self, spec)

    def _on_thumb_context_menu(self, label_widget, pos):
        fp = str(label_widget.property("file_path"))
        targets = list(self.thumb_multi_select) if self.thumb_multi_select and fp in self.thumb_multi_select else [fp]
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

        if hasattr(self, '_clear_multi_spec_selection'):
            menu.addSeparator()
            clear_specs_act = QtWidgets.QAction("Clear spectroscopy selections", menu)
            clear_specs_act.triggered.connect(self._clear_multi_spec_selection)
            menu.addAction(clear_specs_act)

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
        if len(self._multi_spec_selection) >= 2:
            self._open_multi_spectroscopy_popup()

    def _update_spec_selection_label(self):
        count = len(self._multi_spec_selection)
        if hasattr(self, 'spec_selection_label'):
            self.spec_selection_label.setText(f"Spectra selected: {count}")

    def _clear_multi_spec_selection(self):
        self._multi_spec_selection = []
        self._multi_spec_selection_keys = set()
        for dlg in list(self._multi_spectro_popups):
            try:
                dlg.close()
            except Exception:
                pass
        self._multi_spectro_popups = []
        self._update_spec_selection_label()

    def on_clear_spec_selection(self):
        self._clear_multi_spec_selection()

    def _open_multi_spectroscopy_popup(self):
        return spectro_popups._open_multi_spectroscopy_popup(self)

    def on_show_matrix_spectro_viewer(self):
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
            self._refresh_thumbnail_markers()

    def on_pick_spectro_matrix_color(self):
        col = QtWidgets.QColorDialog.getColor(self.spectro_marker_color_matrix, self, "Select Matrix Marker Color", QtWidgets.QColorDialog.ShowAlphaChannel)
        if col.isValid():
            self.spectro_marker_color_matrix = col
            self.config['spectro_marker_color_matrix'] = col.name(QtGui.QColor.HexArgb)
            save_config(self.config)
            self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())
            if self.last_preview:
                self.show_file_channel(self.last_preview[0], self.last_preview[1])
            self._refresh_thumbnail_markers()
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
        self._refresh_thumbnail_markers()

    def on_set_spectro_size(self, size):
        self.spectro_marker_size = float(size)
        self.config['spectro_marker_size'] = self.spectro_marker_size
        save_config(self.config)
        self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())
        if self.last_preview:
            self.show_file_channel(self.last_preview[0], self.last_preview[1])
        self._refresh_thumbnail_markers()

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

    def on_thumb_cmap_changed(self, idx):
        return viewer_thumb_ui.on_thumb_cmap_changed(self, idx)

    def on_preview_cmap_changed(self, idx):
        return viewer_preview.on_preview_cmap_changed(self, idx)

    def on_show_spectra_toggled(self, checked):
        self.show_spectra = bool(checked)
        self.config['show_spectra'] = self.show_spectra; save_config(self.config)
        # Keep UI toggles in sync
        try:
            if hasattr(self, "show_spectra_cb"):
                self.show_spectra_cb.blockSignals(True)
                self.show_spectra_cb.setChecked(self.show_spectra)
                self.show_spectra_cb.blockSignals(False)
            if hasattr(self, "spectro_overlay_act"):
                self.spectro_overlay_act.blockSignals(True)
                self.spectro_overlay_act.setChecked(self.show_spectra)
                self.spectro_overlay_act.blockSignals(False)
        except Exception:
            pass
        if self.show_spectra:
            if not self._spectros_loaded:
                self._reload_spectros(refresh=False)
            else:
                # already loaded for this session; just update counts
                self._update_spectro_stats_label()
        else:
            self._clear_multi_spec_selection()
            self._update_spectro_stats_label()
        self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())
        if self.last_preview:
            self.show_file_channel(self.last_preview[0], self.last_preview[1])
        self._refresh_thumbnail_markers()

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
        self._refresh_thumbnail_markers()
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
        self._refresh_thumbnail_markers()
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
        self._refresh_thumbnail_markers()
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
