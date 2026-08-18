"""Initial state for :class:`SXMGridViewer` - config, defaults, caches.

`__init__` was 1588 lines, ~15% of the whole class, and the place where
all ~800 attributes are born. That made "what state does this object have,
and where does each field come from?" nearly unanswerable - the single
worst problem for anyone new to the code.

This module owns the **first phase**: everything read from config or set to
a default, before any widget exists. `main_window.__init__` now reads as
three clear steps - `init_state(self)`, build widgets, wire them together.

Rules for this module:

* Only plain state. No widget construction, no layout, no signal wiring -
  those belong to the widget phase, which runs after this returns.
* `QTimer`s are allowed (they only need the viewer to exist as a parent)
  and are kept here because their *state* - interval policy, pending
  payloads - is initialization, not layout.
* Attributes assigned here must not be read by the widget phase as local
  variables; the extraction was verified against that (see
  `scripts/analysis/_init_locals.py` in the branch history).
"""
from __future__ import annotations

import threading
import time
from collections import OrderedDict, defaultdict
from pathlib import Path

from .._shared import QtCore, QtGui, QtWidgets, log_status, np
from .. import cmap_registry
from ..config import load_config, save_config
from ..config_io import load_header_cache
from . import theme as ui_theme
from .constants import UI_FONT_FAMILY, UI_FONT_SIZE
from .controllers.collection import CollectionController
from .controllers.filter_controller import FilterController
from .controllers.recent_files_controller import RecentFilesController
from .controllers.report import ReportController
from .controllers.session import SessionController
from .controllers.spectro_compare import SpectroCompareController
from .debounce import ACCUMULATE, LATEST, Debouncer
from .palettes import DEFAULT_COLOR_CYCLE
from .plot_typography import normalize_font_family, set_matplotlib_font_family


def init_state(self):
    """Populate every non-widget attribute on a fresh SXMGridViewer.

    Takes ``self`` (not ``viewer``) so the moved code reads identically to
    when it lived in ``__init__`` - this is a lift, not a rewrite.
    """
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
    self.recent_files_controller = RecentFilesController(self)
    self._normalize_recent_session_history(persist=True)
    last_collection_dir = self.config.get("last_collection_dir")
    try:
        self._last_collection_dir = Path(last_collection_dir) if last_collection_dir else Path(self.last_dir)
    except Exception:
        self._last_collection_dir = Path(self.last_dir)
    self.recent_collections = list(self.config.get("recent_collections", []) or [])
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
    # "Full amber imagery": while the Amber theme is active, render all
    # data imagery with the amber phosphor colormap (display-time only —
    # per-file colormap choices are preserved and restored on toggle-off).
    self.amber_full_imagery = bool(self.config.get('amber_full_imagery', False))
    # Explicit favourites (saved via the star buttons) win over whatever
    # was merely last-used, so every session starts from the user's
    # chosen defaults.
    _fav_cmaps = self.config.get('favorite_cmaps') or {}
    if _fav_cmaps.get('thumbnails'):
        self.thumb_cmap = str(_fav_cmaps['thumbnails'])
    if _fav_cmaps.get('preview'):
        self.preview_cmap = str(_fav_cmaps['preview'])
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
    self.starred = {str(k) for k in (self.config.get("starred") or [])}  # persistent favourites: file paths
    if not self.auto_detect_tags:
        self.tags = {
            str(key): value
            for key, value in dict(self.tags or {}).items()
            if not (isinstance(value, dict) and value.get("auto") and not value.get("manual"))
        }
        self.config["tags"] = self.tags
        save_config(self.config)
    self.filter_controller = FilterController(self)
    self.session_controller = SessionController(self)
    self.collection_controller = CollectionController(self)
    self.report_controller = ReportController(self)
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
    # Named UI theme (light/dark/amber). `dark_mode` stays as the derived
    # legacy flag every existing dark-branch keys off (amber counts as dark).
    self.ui_theme = ui_theme.resolve_theme_name(self.config)
    self.dark_mode = ui_theme.is_dark_theme(self.ui_theme)
    # Arm the display-time colormap override before anything renders so
    # a session persisted with Amber + full-amber-imagery starts amber.
    self._sync_forced_cmap()
    # Global UI font scale in percent (monitor-relative, user-adjustable).
    try:
        self.ui_font_scale = max(60, min(200, int(self.config.get('ui_font_scale', 100))))
    except Exception:
        self.ui_font_scale = 100
    self.detail_dark_view = bool(self.config.get('detail_dark_view', self.dark_mode))
    self._detail_theme_follows_dark_mode = bool(self.config.get('detail_theme_follows_dark_mode', True))
    self.detail_grid_view = bool(self.config.get('detail_grid_view', False))
    self.show_molecules = bool(self.config.get('show_molecules', True))
    self.show_molecule_gizmo = bool(self.config.get("show_molecule_gizmo", False))
    self.show_acquisition_overlay = bool(self.config.get("show_acquisition_overlay", False))
    self.show_filter_summary = bool(self.config.get("show_filter_summary", True))
    self.profile_label_mode = str(self.config.get("profile_label_mode", "length") or "length").strip().lower()
    if self.profile_label_mode not in {"length", "full", "hidden"}:
        self.profile_label_mode = "length"
    self.canvas_display_options = dict(self.config.get("canvas_display_options", {}))
    molecule_style = self.config.get("molecule_default_style") if isinstance(self.config.get("molecule_default_style"), dict) else {}
    self.molecule_palette = str(
        self.config.get("molecule_palette", molecule_style.get("palette", "avogadro")) or "avogadro"
    ).lower()
    self.recent_molecules = list(self.config.get("recent_molecules", []))
    self.recent_svg_molecules = list(self.config.get("recent_svg_molecules", []))
    self.last_svg_molecule_dir = str(self.config.get("last_svg_molecule_dir", "") or "")
    self.svg_molecule_style_defaults = dict(self.config.get("svg_molecule_style_defaults", {}) or {})
    self.quick_crop_mode = bool(self.config.get("quick_crop_mode", False))
    self.quick_crop_aspect_mode = str(self.config.get("quick_crop_aspect_mode", "free") or "free").strip().lower()
    if self.quick_crop_aspect_mode not in {"free", "keep", "square"}:
        self.quick_crop_aspect_mode = "free"
    # Keep crop template editor opt-in at startup for cleaner preview/popup canvases.
    self.show_crop_template_overlay = False
    self.show_crop_history_overlay = True
    self._collection_item_snapshots = {}
    restored_collection_path = str(self.config.get("current_collection_path", "") or "").strip()
    if restored_collection_path and Path(restored_collection_path).exists():
        self._collection_source = restored_collection_path
        self._current_collection_mode = str(self.config.get("current_collection_mode", "") or "") or "linked"
    else:
        self._collection_source = None
        self._current_collection_mode = None
    self._display_defaults = {
        'show_matrix_markers': True,
        'show_single_markers': True,
        'compact_markers': True,
        'detail_dark_view': bool(self.dark_mode),
        'detail_grid_view': False,
        'show_molecules': True,
        'show_molecule_gizmo': False,
        'show_acquisition_overlay': False,
        'show_filter_summary': True,
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
    c_stack = self.config.get('spectro_marker_color_stack')
    if c_stack:
        self.spectro_marker_color_stack = QtGui.QColor(c_stack)
    else:
        self.spectro_marker_color_stack = QtGui.QColor(165, 141, 242, 235)
    self.spectro_color_cycle = (self.config.get('favorite_color_cycle')
                                or self.config.get('spectro_color_cycle', DEFAULT_COLOR_CYCLE))
    self.spectro_marker_symbol = self.config.get('spectro_marker_symbol', 'circle')
    self.spectro_marker_size = float(self.config.get('spectro_marker_size', 5.0))
    self.frame_entry_pixmaps = {}
    self._frame_real_pixmap_cache = {}
    self._processed_views = {}
    self.molecule_overlays = {}
    self.svg_molecule_overlays = {}
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
    # Debounced refreshes (see gui/debounce.py). Each of these used to
    # be a hand-rolled pending/timer/flush attribute triple.
    self._thumbnail_render_debounce = Debouncer(
        self._flush_thumbnail_render_state_refresh,
        interval_ms=120, mode=ACCUMULATE, parent=self)
    self._preview_request_debounce = Debouncer(
        self._flush_preview_request,
        interval_ms=0, mode=LATEST, parent=self)
    self._preview_render_in_progress = False
    self._compact_histogram_apply_timer = QtCore.QTimer(self)
    self._compact_histogram_apply_timer.setSingleShot(True)
    self._compact_histogram_apply_timer.timeout.connect(self._flush_compact_histogram_clim)
    self._pending_compact_histogram_clim = None
    self._pending_compact_histogram_final = False
    self._suppress_compact_histogram_refresh = False
    self._compact_histogram_gesture_active = False
    self._spectro_manifest_save_timer = QtCore.QTimer(self)
    self._spectro_manifest_save_timer.setSingleShot(True)
    self._spectro_manifest_save_timer.timeout.connect(self._flush_spectro_manifest_save)
    self._spectro_manifest_save_inflight = False
    self._spectro_manifest_save_pending = False
    self._left_sidebar_min_width = 300
    self._left_sidebar_target_width = 340
    self._left_sidebar_soft_max_width = 380
    self._left_sidebar_rebalance_timer = QtCore.QTimer(self)
    self._left_sidebar_rebalance_timer.setSingleShot(True)
    self._left_sidebar_rebalance_timer.timeout.connect(self._rebalance_main_splitter)
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
    self.per_file_channel_clim = {}
    self.last_preview = None
    self.spectros = []
    self.matrix_spectros = []
    self.files_with_matrix = set()
    self.files_with_spectra = set()
    self.spectros_by_image = defaultdict(list)
    self.spectro_sites_by_image = defaultdict(list)
    self.spectro_site_index = {}
    self.spectro_groups_by_image = defaultdict(list)
    self.spectro_group_index = {}
    self._spec_extent_cache = {}
    self._spectros_loaded = False
    self._spectros_loading = False
    self._spectros_pending = False
    self._spectro_cache = {}
    self._spectro_manifest_entries = {}
    self._spectro_miniature_cache = OrderedDict()
    self._spectro_autoload_timer = QtCore.QTimer(self)
    self._spectro_autoload_timer.setSingleShot(True)
    self._spectro_autoload_timer.timeout.connect(self._run_pending_spectro_load_async)
    self._spectro_scan_thread = None
    self._spectro_scan_worker = None
    self._spectro_manifest_pending_save = False
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
    self.virtual_copy_order = []
    self.thumbnail_filters = {}
    self.image_adjustments = defaultdict(dict)
    # Global-undo entries for Crop/Rotate actions: ("tone", key, ch,
    # old_spec) restores a display spec, ("copy", key) removes an
    # [edit] virtual copy. Consumed by _undo_last_adjustment.
    self._adjustment_undo_stack = []
    self._last_base_array = None
    self._last_base_extent = None
    self._last_base_unit = None
    self._spectro_hist_cache = {}
    self.matrix_datasets = {}
    log_status("Loading header cache...")
    _hc_t0 = time.perf_counter()
    # Join the background load started at the top of __init__ instead of
    # loading synchronously here - by this point the load has usually
    # already run concurrently with widget construction above, so this
    # join is typically near-instant rather than paying the full cost.
    self._header_cache_thread.join()
    self.header_cache = self._header_cache_bg_result.get("cache", {})
    log_status(f"[Perf] Header cache loaded: {(time.perf_counter() - _hc_t0) * 1000:.0f} ms (background-overlapped) | {len(self.header_cache)} entries")
    self._header_cache_dirty = False
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
    # Activity-log batching state lives in ActivityLog (gui/activity_log.py),
    # constructed once the text box exists in _create_lower_controls.
    self.activity_log = None
