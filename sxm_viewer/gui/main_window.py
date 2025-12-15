"""Main Qt widget implementing the SXM grid viewer."""
from __future__ import annotations

import re

from .._shared import *
from ..config import *
from ..data.io import *
from ..data.spectroscopy import *
from ..processing.filters import *
from ..processing.detection import *
from .thumbnails import *
from .minimap import FrameMiniMap
from .detail_panels import *


class SXMGridViewer(QtWidgets.QWidget):
    FRAME_ZOOM_SLIDER_MIN = 0
    FRAME_ZOOM_SLIDER_MAX = 600
    FRAME_ZOOM_SLIDER_DEFAULT = 200

    def __init__(self):
        super().__init__()
        log_status("Initializing SXM Grid Viewer...")
        self.setWindowTitle("SXM Grid Viewer")
        self.resize(1250, 840)

        log_status("Loading configuration...")
        self.config = load_config()
        self.last_dir = Path(self.config.get("last_dir", str(Path.cwd())))
        self.last_channel_index = int(self.config.get("last_channel_index", 0))
        self.thumb_cmap = self.config.get("thumbnail_cmap", "viridis")
        self.preview_cmap = self.config.get("preview_cmap", "viridis")
        self.spec_folder_path = Path(self.config.get("spectra_folder", str(self.last_dir)))
        self.show_spectra = bool(self.config.get("show_spectra", True))
        self.thumb_size_px = int(self.config.get("thumb_size_px", 160))
        self.tags = self.config.get("tags", {})  # persistent tags: {path: {"tag":"constant-height","abs_z_pm":int,...}}
        self.frame_map_entries = []
        self.show_shortcuts_panel = bool(self.config.get("show_shortcuts_panel", False))
        self.hidden_frame_keys = set()
        self.frame_real_view = False
        self.frame_entry_pixmaps = {}
        self._frame_real_pixmap_cache = {}

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

        self.per_file_channel_cmap = {}
        self.last_preview = None
        self.spectros = []
        self.matrix_spectros = []
        self.spectros_by_image = defaultdict(list)
        self._spectro_cache = {}
        self.image_time_index = {}
        self._spectro_popups = []
        self._popup_refs = []
        self._multi_spectro_popups = []
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
        log_status("Loading header cache...")
        self.header_cache = load_header_cache()
        self._header_cache_dirty = False
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
        base_font = QtGui.QFont("Segoe UI", 11)
        bold_font = QtGui.QFont("Segoe UI", 11, QtGui.QFont.Bold)
        meta_font = QtGui.QFont("Consolas", 16)
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

        # UI: left controls + meta + inspector; right thumbs + preview
        left_v = QtWidgets.QVBoxLayout(); left_v.setSpacing(8)
        path_h = QtWidgets.QHBoxLayout()
        self.path_le = QtWidgets.QLineEdit(str(self.last_dir))
        self.open_btn = QtWidgets.QPushButton("Open folder")
        path_h.addWidget(self.path_le); path_h.addWidget(self.open_btn); left_v.addLayout(path_h)

        spec_path_h = QtWidgets.QHBoxLayout()
        spec_path_h.addWidget(QtWidgets.QLabel("Spectra folder:"))
        self.spec_folder_le = QtWidgets.QLineEdit(str(self.spec_folder_path))
        self.spec_folder_le.setPlaceholderText("Defaults to SXM folder")
        self.spec_folder_btn = QtWidgets.QPushButton("Browse")
        spec_path_h.addWidget(self.spec_folder_le, 1)
        spec_path_h.addWidget(self.spec_folder_btn)
        left_v.addLayout(spec_path_h)

        controls_h = QtWidgets.QHBoxLayout()
        self.channel_label = QtWidgets.QLabel("Channel:")
        self.channel_label.setFont(bold_font)
        self.channel_dropdown = QtWidgets.QComboBox(); self.channel_dropdown.setMinimumWidth(380)
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
        controls_h.addWidget(self.channel_label); controls_h.addWidget(self.channel_dropdown)
        controls_h.addWidget(QtWidgets.QLabel("Thumb cmap:")); controls_h.addWidget(self.thumb_cmap_combo)
        controls_h.addWidget(QtWidgets.QLabel("Preview cmap:")); controls_h.addWidget(self.preview_cmap_combo)
        # Dark mode toggle
        self.dark_mode = bool(self.config.get('dark_mode', False))
        self.dark_mode_cb = QtWidgets.QCheckBox('Dark mode')
        self.dark_mode_cb.setChecked(self.dark_mode)
        controls_h.addWidget(self.dark_mode_cb)
        left_v.addLayout(controls_h)

        # Metadata / inspector box: preserve formatting and readability.
        self.meta_box = QtWidgets.QTextEdit()
        self.meta_box.setReadOnly(True)
        self.meta_box.setFont(meta_font)
        self.meta_box.setMinimumWidth(380)
        try:
            self.meta_box.setLineWrapMode(QtWidgets.QTextEdit.NoWrap)
        except Exception:
            pass
        self.meta_box.setPlaceholderText("File metadata / header appears when selecting a thumbnail.")
        left_v.addWidget(self.meta_box, 1)

        frame_group = QtWidgets.QGroupBox("Folder layout (±1 µm)")
        frame_layout = QtWidgets.QVBoxLayout(frame_group)
        self.frame_map_widget = FrameMiniMap()
        self.frame_map_widget.entryClicked.connect(self._on_frame_map_clicked)
        self.frame_map_widget.entryShiftClicked.connect(self._on_frame_map_entry_shift_clicked)
        self.frame_map_widget.zoomChanged.connect(self._on_frame_map_zoom_changed)
        self.frame_map_widget.setToolTip(
            "Frame layout:\n"
            " • Click to focus a frame\n"
            " • Shift+Click hides a frame (Show all resets)\n"
            " • Mouse wheel zooms view; drag to pan\n"
            " • Toggle “Show real view” for channel thumbnails"
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
        left_v.addWidget(frame_group)

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
        title_lbl = QtWidgets.QLabel("Selected channel"); title_lbl.setFont(bold_font)
        self.scroll = QtWidgets.QScrollArea(); self.thumb_container = QtWidgets.QWidget(); self.thumb_layout = QtWidgets.QGridLayout(); self.thumb_layout.setSpacing(14)
        self.scroll.setToolTip(
            "Thumbnails:\n"
            " • Shift+Click or Ctrl+Click to multi-select\n"
            " • Ctrl+Wheel to change thumbnail size\n"
            " • Right-click a frame for filters & exports"
        )
        self.thumb_container.setLayout(self.thumb_layout); self.scroll.setWidgetResizable(True); self.scroll.setWidget(self.thumb_container)
        self._thumb_viewport = self.scroll.viewport()
        self._thumb_viewport.installEventFilter(self)
        self.scroll.installEventFilter(self)
        self.thumb_container.installEventFilter(self)
        thumbs_panel = QtWidgets.QWidget()
        thumbs_panel_layout = QtWidgets.QVBoxLayout(); thumbs_panel_layout.setContentsMargins(0,0,0,0)
        thumbs_panel_layout.addWidget(title_lbl)
        thumbs_panel_layout.addWidget(self.scroll, 1)
        thumbs_toolbar = QtWidgets.QHBoxLayout()
        thumbs_toolbar.addWidget(QtWidgets.QLabel('Sort:'))
        self.thumb_sort_combo = QtWidgets.QComboBox()
        self.thumb_sort_combo.addItems(['Name (A-Z)', 'Date (new-old)', 'Date (old-new)', 'Tag (CH-CC-U)'])
        thumbs_toolbar.addWidget(self.thumb_sort_combo)
        thumbs_toolbar.addSpacing(8)
        thumbs_toolbar.addWidget(QtWidgets.QLabel('Filter:'))
        self.thumb_filter_combo = QtWidgets.QComboBox()
        self.thumb_filter_combo.addItems(['Name (A-Z)', 'Date (new-old)', 'Date (old-new)', 'Tag (CH-CC-U)'])
        thumbs_toolbar.addWidget(self.thumb_filter_combo)
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
        self.preview_canvas = MultiPreviewCanvas(self, figsize=(6,5))
        self.preview_canvas.setToolTip(
            "Preview area:\n"
            " • Right-click for copy/save options\n"
            " • Enable 'Measure profile' for line sampling\n"
            " • Ctrl+C copies the focused image to clipboard"
        )
        self.preview_canvas.set_copy_feedback_handler(self._on_view_copied)
        preview_panel_layout.addWidget(self.preview_canvas, 1)
        self.preview_value_label = QtWidgets.QLabel("Value: --")
        preview_panel_layout.addWidget(self.preview_value_label)
        extra_h = QtWidgets.QHBoxLayout()
        self.add_view_btn = QtWidgets.QPushButton("Add channel view"); self.clear_views_btn = QtWidgets.QPushButton("Clear extra views")
        self.export_pngs_btn = QtWidgets.QPushButton("Export PNGs")
        self.export_xyz_btn = QtWidgets.QPushButton("Export XYZ")
        self.adjust_image_btn = QtWidgets.QPushButton("Adjust image")
        self.adjust_image_btn.setEnabled(False)
        extra_h.addWidget(self.add_view_btn); extra_h.addWidget(self.clear_views_btn); extra_h.addWidget(self.export_pngs_btn); extra_h.addWidget(self.export_xyz_btn); extra_h.addWidget(self.adjust_image_btn)
        self.measure_profile_btn = QtWidgets.QPushButton('Measure profile')
        extra_h.addWidget(self.measure_profile_btn)
        self.show_spectra_cb = QtWidgets.QCheckBox("Show spectroscopies")
        self.show_spectra_cb.setChecked(self.show_spectra)
        extra_h.addWidget(self.show_spectra_cb)
        self.show_matrix_spectra_btn = QtWidgets.QPushButton("Show Matrix spectros")
        extra_h.addWidget(self.show_matrix_spectra_btn)
        self.clear_spec_selection_btn = QtWidgets.QPushButton("Clear spec selection")
        extra_h.addWidget(self.clear_spec_selection_btn)
        self.spec_selection_label = QtWidgets.QLabel("Spectra selected: 0")
        font_small = QtGui.QFont("Segoe UI", 9)
        self.spec_selection_label.setFont(font_small)
        extra_h.addWidget(self.spec_selection_label)
        self.export_selected_btn = QtWidgets.QPushButton("Export selected (same view)")
        extra_h.addWidget(self.export_selected_btn)
        preview_panel_layout.addLayout(extra_h)
        preview_panel.setLayout(preview_panel_layout)
        self.preview_canvas.set_value_callback(self._on_preview_value)

        right_splitter = QtWidgets.QSplitter(QtCore.Qt.Vertical)
        right_splitter.addWidget(thumbs_panel)
        right_splitter.addWidget(preview_panel)
        right_splitter.setStretchFactor(0, 3)
        right_splitter.setStretchFactor(1, 2)
        right_w = QtWidgets.QWidget()
        right_layout = QtWidgets.QVBoxLayout(); right_layout.setContentsMargins(0,0,0,0)
        right_layout.addWidget(right_splitter, 1)
        right_w.setLayout(right_layout)

        main_splitter = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        main_splitter.addWidget(left_w)
        main_splitter.addWidget(right_w)
        # prefer to give the right pane more space by default
        main_splitter.setStretchFactor(0, 1)
        main_splitter.setStretchFactor(1, 3)

        # Prevent panes from collapsing to zero width when the user drags the splitter.
        # This avoids the left inspector disappearing when the user expands the thumbnails.
        try:
            main_splitter.setCollapsible(0, False)
            main_splitter.setCollapsible(1, False)
        except Exception:
            # older PyQt versions may not support setCollapsible; ignore safely
            pass

        # Ensure the left widget cannot shrink below a useful width
        try:
            left_w.setMinimumWidth(360)
        except Exception:
            pass

        # Set reasonable initial sizes (left, right). Adjust these numbers to taste.
        try:
            main_splitter.setSizes([360, 1080])
        except Exception:
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
        # no size slider callback
        # inspector widgets removed -> no connections required here
        self.add_view_btn.clicked.connect(self.on_add_view)
        self.clear_views_btn.clicked.connect(self.on_clear_views)
        self.export_pngs_btn.clicked.connect(self.on_export_pngs)
        self.adjust_image_btn.clicked.connect(self.on_adjust_image)
        self.export_xyz_btn.clicked.connect(self.on_export_xyz_files)
        self.measure_profile_btn.clicked.connect(self._on_start_profile)
        self.show_spectra_cb.toggled.connect(self.on_show_spectra_toggled)
        self.show_matrix_spectra_btn.clicked.connect(self.on_show_matrix_spectro_viewer)
        self.clear_spec_selection_btn.clicked.connect(self.on_clear_spec_selection)
        self.export_selected_btn.clicked.connect(self.on_export_selected_same_view)
        self.tag_ch_btn.clicked.connect(lambda: self.on_manual_tag('constant-height'))
        self.tag_cc_btn.clicked.connect(lambda: self.on_manual_tag('constant-current'))
        self.untag_btn.clicked.connect(lambda: self.on_manual_tag(None))
        self.dark_mode_cb.toggled.connect(self.on_dark_mode_toggled)

        # autoload
        if self.last_dir.exists():
            QtCore.QTimer.singleShot(50, lambda: self.load_folder(self.last_dir))
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
        else:
            app.setPalette(app.style().standardPalette())
        if hasattr(self, 'shortcuts_label'):
            self.shortcuts_label.setText(self._shortcuts_html())

    def _create_shortcuts_panel(self):
        frame = QtWidgets.QFrame()
        frame.setObjectName("shortcutsPanel")
        frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        frame.setStyleSheet("""
        QFrame#shortcutsPanel {
            background: rgba(64, 96, 160, 25%);
            border: 1px solid rgba(64, 96, 160, 60%);
            border-radius: 8px;
            padding: 6px;
        }
        """)
        layout = QtWidgets.QVBoxLayout(frame)
        layout.setContentsMargins(10, 8, 8, 8)
        header = QtWidgets.QHBoxLayout()
        title = QtWidgets.QLabel("Shortcuts & gestures")
        title.setFont(QtGui.QFont("Segoe UI", 10, QtGui.QFont.Bold))
        header.addWidget(title)
        header.addStretch(1)
        never_btn = QtWidgets.QPushButton("Don't show again")
        never_btn.setFlat(True)
        never_btn.setCursor(QtCore.Qt.PointingHandCursor)
        never_btn.clicked.connect(self._on_shortcuts_never_show_clicked)
        header.addWidget(never_btn)
        close_btn = QtWidgets.QToolButton()
        close_btn.setText("✕")
        close_btn.setAutoRaise(True)
        close_btn.setCursor(QtCore.Qt.PointingHandCursor)
        close_btn.clicked.connect(self._on_hide_shortcuts_panel)
        header.addWidget(close_btn)
        layout.addLayout(header)
        self.shortcuts_label = QtWidgets.QLabel(self._shortcuts_html())
        self.shortcuts_label.setWordWrap(True)
        self.shortcuts_label.setTextFormat(QtCore.Qt.RichText)
        layout.addWidget(self.shortcuts_label)
        return frame

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

    def eventFilter(self, obj, event):
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
        return super().eventFilter(obj, event)

    def _thumb_dimensions(self):
        """Return (width, height) for thumbnails preserving 4:3 aspect ratio."""
        w = int(max(64, min(360, getattr(self, 'thumb_size_px', 160))))
        h = int(max(48, round(w * 0.75)))
        return w, h

    def _resize_thumbnail_scale(self, delta_px):
        new_w = int(max(64, min(360, self.thumb_size_px + delta_px)))
        if new_w == self.thumb_size_px:
            return
        self.thumb_size_px = new_w
        self.config['thumb_size_px'] = new_w
        save_config(self.config)
        self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())

    def _create_toolbar(self):
        try:
            toolbar = QtWidgets.QToolBar("Main toolbar", self)
        except Exception:
            return None
        toolbar.setIconSize(QtCore.QSize(20, 20))

        def _icon(name):
            icon = QIcon.fromTheme(name)
            return icon if icon and not icon.isNull() else QIcon()

        self.toolbar_open_act = toolbar.addAction(_icon("folder-open"), "Open folder")
        self.toolbar_open_act.triggered.connect(self.open_folder_dialog)
        toolbar.addSeparator()

        self.toolbar_export_png_act = toolbar.addAction(_icon("image-x-generic"), "Export PNGs")
        self.toolbar_export_png_act.triggered.connect(self.on_export_pngs)

        self.toolbar_export_xyz_act = toolbar.addAction(_icon("document-save"), "Export XYZ")
        self.toolbar_export_xyz_act.triggered.connect(self.on_export_xyz_files)

        toolbar.addSeparator()
        self.toolbar_adjust_act = toolbar.addAction(_icon("transform-crop"), "Adjust image")
        self.toolbar_adjust_act.triggered.connect(self.on_adjust_image)
        self.toolbar_shortcuts_act = toolbar.addAction(_icon("help-about"), "Shortcuts")
        self.toolbar_shortcuts_act.triggered.connect(self._on_show_shortcuts_requested)

        self._update_toolbar_actions(False)
        return toolbar

    def _update_toolbar_actions(self, enabled: bool):
        for act in (self.toolbar_export_png_act, self.toolbar_export_xyz_act, self.toolbar_adjust_act):
            if act is not None:
                act.setEnabled(bool(enabled))

    def on_dark_mode_toggled(self, checked: bool):
        self.dark_mode = bool(checked)
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

    def load_folder(self, folder:Path):
        folder = Path(folder)
        log_status(f"Loading folder: {folder}")
        self._update_toolbar_actions(False)
        self.last_dir = folder
        self.path_le.setText(str(folder))
        # persist last dir early
        self.config['last_dir'] = str(folder)
        save_config(self.config)

        txts = sorted(folder.glob("*.txt"))
        log_status(f"Found {len(txts)} .txt files")
        self.files = txts
        self.headers.clear()
        self._invalidate_thumbnail_cache()
        self._invalidate_channel_cache()
        self.thumb_multi_select = set()
        cache_hits = 0
        cache_miss = 0
        for t in txts:
            cached = self._get_cached_header(t)
            if cached:
                hdr, fds = cached
                cache_hits += 1
            else:
                try:
                    hdr, fds = parse_header(t)
                    cache_miss += 1
                    self._store_header_cache(t, hdr, fds)
                except Exception:
                    continue
            self.headers[str(t)] = (hdr, fds)
        if cache_miss:
            self._save_header_cache()
        log_status(f"Headers loaded (hits={cache_hits}, miss={cache_miss})")
        if not self.headers:
            self.meta_box.setPlainText("No valid .txt headers found")
            self.clear_thumbs(); return
        self._build_image_timestamp_index()
        self._rebuild_frame_map_entries()

        # build channel dropdown from first header
        first_key = next(iter(self.headers))
        _, first_fds = self.headers[first_key]
        labels = []
        for idx, fd in enumerate(first_fds):
            cap = fd.get('Caption', fd.get('FileName', f"chan{idx}"))
            labels.append(f"{idx}: {cap}")
        max_channels = max(len(v[1]) for v in self.headers.values())
        if max_channels > len(labels):
            for idx in range(len(labels), max_channels):
                labels.append(f"{idx}: chan{idx}")

        self.channel_dropdown.blockSignals(True)
        self.channel_dropdown.clear()
        for lab in labels:
            self.channel_dropdown.addItem(lab)
            self.channel_dropdown.setItemData(self.channel_dropdown.count()-1, lab, QtCore.Qt.ToolTipRole)
        self.channel_dropdown.setMinimumWidth(380)
        if 0 <= self.last_channel_index < self.channel_dropdown.count():
            self.channel_dropdown.setCurrentIndex(self.last_channel_index)
        else:
            self.last_channel_index = 0; self.channel_dropdown.setCurrentIndex(0)
        self.channel_dropdown.blockSignals(False)

        # set cmaps
        try: self.thumb_cmap_combo.setCurrentText(self.thumb_cmap)
        except: pass
        try: self.preview_cmap_combo.setCurrentText(self.preview_cmap)
        except: pass
        # set icon sizes for cmap combos
        try:
            self.thumb_cmap_combo.setIconSize(QtCore.QSize(96, 14))
            self.preview_cmap_combo.setIconSize(QtCore.QSize(96, 14))
        except Exception:
            pass

        # auto-detect tags for files not already tagged
        log_status("Auto-detecting tags...")
        self._auto_detect_tags_for_folder()

        # load spectroscopy markers referencing this folder
        log_status("Loading spectroscopy references...")
        self._reload_spectros(refresh=False)

        QtCore.QTimer.singleShot(0, lambda: self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex()))
        log_status("Folder load complete.")

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
        while self.thumb_layout.count():
            item = self.thumb_layout.takeAt(0); w = item.widget()
            if w: w.setParent(None)
        self.thumb_widgets = {}
        self._thumb_labels = {}

    def populate_thumbnails_for_channel(self, channel_idx:int):
        self.clear_thumbs()
        max_cols = 4; row = 0; col = 0
        thumb_w, thumb_h = self._thumb_dimensions()
        cmap_name = self.thumb_cmap_combo.currentText() or self.thumb_cmap
        self._thumb_generation += 1
        generation = self._thumb_generation
        self.meta_box.setPlainText(f"Building thumbnails for channel {channel_idx} ...")
        files_iter = list(self.files)

        filt = (self.thumb_filter_combo.currentText() if hasattr(self, 'thumb_filter_combo') else 'All')
        if filt != 'All':
            def include(path_str):
                tag = (self.tags.get(path_str, {}) or {}).get('tag', None)
                if filt == 'CH only':
                    return tag == 'constant-height'
                if filt == 'CC only':
                    return tag == 'constant-current'
                if filt == 'Untagged':
                    return tag is None
                return True
            files_iter = [t for t in files_iter if include(str(t))]

        sort_mode = (self.thumb_sort_combo.currentText() if hasattr(self, 'thumb_sort_combo') else 'Name (A?Z)')
        if sort_mode.startswith('Name'):
            files_iter.sort(key=lambda p: Path(p).name.lower())
        elif 'Date (new' in sort_mode or 'Date (old' in sort_mode:
            rev = ('new' in sort_mode)
            def sort_key_date(p):
                hdr = self.headers.get(str(p), (None, None))[0]
                return self._parse_header_datetime(hdr)
            files_iter.sort(key=sort_key_date, reverse=rev)
        elif sort_mode.startswith('Tag'):
            order = {'constant-height': 0, 'constant-current': 1, None: 2}
            files_iter.sort(key=lambda p: (order.get((self.tags.get(str(p), {}) or {}).get('tag', None), 2), Path(p).name.lower()))

        for i, t in enumerate(files_iter):
            key = str(t)
            if key not in self.headers:
                continue
            header, fds = self.headers[key]
            lbl = QtWidgets.QLabel()
            lbl.setAlignment(QtCore.Qt.AlignCenter)
            lbl.setProperty("file_path", key)
            lbl.setProperty("channel_index", int(channel_idx))
            lbl.setProperty("spec_markers", [])
            lbl.setProperty("thumb_dims", (thumb_w, thumb_h))
            placeholder = QtGui.QPixmap(thumb_w, thumb_h)
            placeholder.fill(QtGui.QColor('#0b0b12'))
            lbl.setPixmap(placeholder)
            lbl.setMouseTracking(True)
            lbl.mousePressEvent = self._make_thumb_click_handler(lbl)
            lbl.mouseMoveEvent = self._make_thumb_move_handler(lbl)
            lbl.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
            lbl.customContextMenuRequested.connect(lambda pos, lb=lbl: self._on_thumb_context_menu(lb, pos))
            vbox = QtWidgets.QVBoxLayout(); vbox.setContentsMargins(0,0,0,0); vbox.setSpacing(2)
            card = QtWidgets.QFrame(); card.setFrameShape(QtWidgets.QFrame.StyledPanel); card.setLineWidth(0)
            card_layout = QtWidgets.QVBoxLayout(card); card_layout.setContentsMargins(4,4,4,4); card_layout.setSpacing(4)
            vbox.addWidget(lbl)
            cap = QtWidgets.QLabel(Path(t).name); cap.setAlignment(QtCore.Qt.AlignCenter); cap.setMaximumHeight(18)
            cap.setFont(QtGui.QFont("Segoe UI", 9)); vbox.addWidget(cap)
            card_layout.addLayout(vbox)
            self.thumb_layout.addWidget(card, row, col)
            self.thumb_widgets[key] = card
            self._thumb_labels[key] = lbl
            try:
                if key in getattr(self, 'thumb_multi_select', set()):
                    card.setStyleSheet("QFrame { border: 2px solid #a36bff; border-radius: 10px; background-color: rgba(163,107,255,40); }")
                elif key == str(getattr(self, 'selected_file_for_thumbs', None)):
                    card.setStyleSheet("QFrame { border: 2px solid #5f8dd3; border-radius: 10px; background-color: rgba(95,141,211,40); }")
                else:
                    card.setStyleSheet("QFrame { border: 1px solid rgba(255,255,255,30); border-radius: 10px; background-color: transparent; }")
            except Exception:
                pass

            if fds and 0 <= channel_idx < len(fds):
                fd = fds[channel_idx]
                base_pix = None
                data_key = None
                try:
                    data_key = self._thumbnail_data_key(key, channel_idx, fd, thumb_w, thumb_h)
                except Exception:
                    data_key = None
                if data_key:
                    base_pix = self.thumb_cache.get((data_key, cmap_name))
                if base_pix is not None:
                    pix = base_pix.copy()
                    markers = self._decorate_thumbnail_pixmap(pix, key, channel_idx, header, fds)
                    lbl.setPixmap(pix)
                    lbl.setProperty("spec_markers", markers)
                else:
                    lbl.setProperty("spec_markers", [])
                    self._schedule_thumbnail_job(key, channel_idx, header, fd, thumb_w, thumb_h, cmap_name, generation)
            else:
                blank = QtGui.QPixmap(thumb_w, thumb_h)
                blank.fill(QtGui.QColor('black'))
                lbl.setPixmap(blank)
                lbl.setProperty("spec_markers", [])

            col += 1
            if col >= max_cols:
                col = 0; row += 1
        self.meta_box.setPlainText(f"Thumbnails built for channel {channel_idx}  (thumb cmap: {cmap_name})")
        self._refresh_frame_map_pixmaps()
    def _thumbnail_filter_signature(self, file_key):
        spec = self.thumbnail_filters.get(str(file_key))
        return _filter_signature(spec)

    def _downsample_for_thumbnail(self, arr, thumb_w, thumb_h):
        arr = np.asarray(arr, dtype=float)
        if arr.size == 0:
            return arr
        h, w = arr.shape
        if h > thumb_h or w > thumb_w:
            ys = np.linspace(0, h - 1, thumb_h).astype(int)
            xs = np.linspace(0, w - 1, thumb_w).astype(int)
            return arr[np.ix_(ys, xs)]
        return arr

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
                marker_defs = self._draw_spectro_markers_on_pixmap(pix, header, str(file_key), xpix, ypix)
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
        filter_sig = self._thumbnail_filter_signature(file_key)
        fname = fd.get("FileName")
        if not fname:
            raise ValueError("Missing FileName for channel")
        bin_path = Path(file_key).parent / fname
        try:
            bin_mtime = bin_path.stat().st_mtime
        except Exception:
            bin_mtime = 0.0
        data_key = (file_key, channel_idx, bin_mtime, filter_sig, thumb_w, thumb_h)
        with self._thumb_data_lock:
            cached = self._thumb_data_cache.get(data_key)
        if cached is not None:
            return data_key, cached
        _, arr_conv = self._get_filtered_channel_array(file_key, channel_idx, header, fd)
        thumb_arr = self._downsample_for_thumbnail(arr_conv, thumb_w, thumb_h)
        with self._thumb_data_lock:
            self._thumb_data_cache[data_key] = thumb_arr
        return data_key, thumb_arr

    def _thumbnail_data_key(self, file_key, channel_idx, fd, thumb_w, thumb_h):
        filter_sig = self._thumbnail_filter_signature(file_key)
        fname = fd.get("FileName")
        if not fname:
            raise ValueError("Missing FileName for channel")
        bin_path = Path(file_key).parent / fname
        try:
            bin_mtime = bin_path.stat().st_mtime
        except Exception:
            bin_mtime = 0.0
        return (file_key, channel_idx, bin_mtime, filter_sig, thumb_w, thumb_h)

    def _invalidate_thumbnail_cache(self, paths=None):
        if not paths:
            with self._thumb_data_lock:
                self._thumb_data_cache.clear()
            self.thumb_cache.clear()
            self._frame_real_pixmap_cache.clear()
            return
        path_set = {str(Path(p)) for p in paths}
        with self._thumb_data_lock:
            data_keys = [k for k in self._thumb_data_cache.keys() if k[0] in path_set]
            for k in data_keys:
                self._thumb_data_cache.pop(k, None)
        pix_keys = [k for k in self.thumb_cache.keys() if k[0][0] in path_set]
        for k in pix_keys:
            self.thumb_cache.pop(k, None)
        self._frame_real_pixmap_cache.clear()

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
        try:
            self.config['thumb_sort'] = self.thumb_sort_combo.currentText(); save_config(self.config)
        except Exception:
            pass
        self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())

    def on_thumb_filter_changed(self, idx):
        try:
            self.config['thumb_filter'] = self.thumb_filter_combo.currentText(); save_config(self.config)
        except Exception:
            pass
        self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())

    # removed size change handler

    def _parse_header_datetime(self, header):
        """Return a sortable key (float timestamp) parsed from header Date/Time if possible; otherwise 0.0.
        Accepts common formats, falls back to 0.0 on failure."""
        try:
            date = str(header.get('Date', '') or '').strip()
            time = str(header.get('Time', '') or '').strip()
            if not date and not time:
                return 0.0
            candidates = []
            if date and time:
                candidates.append(f"{date} {time}")
            if date:
                candidates.append(date)
            fmts = [
                '%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M', '%Y/%m/%d %H:%M:%S', '%d/%m/%Y %H:%M:%S',
                '%d-%m-%Y %H:%M:%S', '%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y'
            ]
            for s in candidates:
                for fmt in fmts:
                    try:
                        dt = datetime.strptime(s, fmt)
                        return dt.timestamp()
                    except Exception:
                        continue
            return 0.0
        except Exception:
            return 0.0

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

    def _build_metadata_html(self, header_path:Path, header:dict, fd:dict, channel_idx:int, unit_final:str, arr_conv:np.ndarray) -> str:
        """Return HTML for the metadata pane with clearer styling and sections."""
        def esc(s):
            try:
                return str(s).replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
            except Exception:
                return ''
        dark = bool(getattr(self, 'dark_mode', False))
        text_color = '#e0e0e0' if dark else '#222'
        label_color = '#a0a0a0' if dark else '#555'
        accent_border = '#6fa8ff' if dark else '#4a7edb'
        accent_bg = 'rgba(111,168,255,0.16)' if dark else 'rgba(74,126,219,0.10)'
        filename = header_path.name
        date = header.get('Date', '')
        time = header.get('Time', '')
        bias = header.get('Bias', None); bias_unit = header.get('BiasPhysUnit', '')
        setp = header.get('SetPoint', None); setp_unit = header.get('SetPointPhysUnit', '')
        user = header.get('UserName', '')
        cap = fd.get('Caption','')
        phys_orig = fd.get('PhysUnit','')
        scale = fd.get('Scale','')
        offset = fd.get('Offset','')
        def fmt_number(val, precision=3):
            try:
                num = float(val)
                return f"{num:.{precision}f}".rstrip('0').rstrip('.')
            except Exception:
                if val is None:
                    return ''
                return esc(val)

        # stats
        try:
            flat = np.asarray(arr_conv).ravel()
            vmin = np.nanmin(flat); vmax = np.nanmax(flat); vmed = np.nanmedian(flat)
            stats = f"min={vmin:.6g} | max={vmax:.6g} | median={vmed:.6g}"
        except Exception:
            stats = "min/max/median: N/A"
        # tags
        taginfo = self.tags.get(str(header_path), {})
        tag_label = taginfo.get('tag', None)
        tag_chip = ''
        if tag_label == 'constant-height':
            chip_color = '#2e7d32'; chip_text = 'CH'
        elif tag_label == 'constant-current':
            chip_color = '#1565c0'; chip_text = 'CC'
        else:
            chip_color = None; chip_text = ''
        if chip_color:
            tag_chip = f"<span style='background:{chip_color};color:#fff;border-radius:10px;padding:2px 8px;font-weight:600'>" \
                       f"{chip_text}</span> <span style='color:#555'>({esc(tag_label)})</span>"
        # abs z + dzs
        ch_lines = ''
        abs_nm = None
        if tag_label == 'constant-height':
            abs_pm = taginfo.get('abs_z_pm', None)
            if abs_pm is not None:
                abs_nm = abs_pm/1000.0
                ch_lines += f"<div>Const-height (abs z): <b>{abs_nm:.3f} nm</b></div>"
            dz_prev_nonch, prevname = self._dz_vs_last_before_ch(header_path)
            if dz_prev_nonch is not None:
                ch_lines += f"<div>dz vs prev non-CH (<i>{esc(prevname)}</i>): <b>{dz_prev_nonch:+.0f} pm</b> ({dz_prev_nonch/1000.0:+.3f} nm)</div>"
            dz_prev_ch, prevch_name = self._dz_vs_previous_ch(header_path)
            if dz_prev_ch is not None:
                ch_lines += f"<div>dz vs prev CH (<i>{esc(prevch_name)}</i>): <b>{dz_prev_ch:+.0f} pm</b> ({dz_prev_ch/1000.0:+.3f} nm)</div>"

        # control params
        params = {}
        def collect_params(d):
            for k,v in (d or {}).items():
                kl = str(k).lower()
                if any(tok in kl for tok in ('ki','kp','pll','ampl','amplitude','amp','setpoint','natural','natfreq','freq','f0','kpl','kipl','lockin')):
                    try:
                        params[k] = float(v)
                    except Exception:
                        params[k] = v
        collect_params(header); collect_params(fd)
        params_rows = ''.join([f"<tr><td>{esc(k)}</td><td style='text-align:right'>{esc(v)}</td></tr>" for k,v in params.items()])

        spec_section = ''
        spec_entries = self.spectros_by_image.get(str(header_path), [])
        if self.show_spectra and spec_entries:
            rows = []
            for idx, spec in enumerate(spec_entries[:6], 1):
                name = Path(spec['path']).name
                matrix_idx = spec.get('matrix_index')
                if matrix_idx is not None:
                    name = f"{name} [{matrix_idx}]"
                xs = spec.get('x')
                ys = spec.get('y')
                pos_txt = f"{xs:.1f}/{ys:.1f} nm" if xs is not None and ys is not None else "n/a"
                rows.append(f"<tr><td>S{idx}</td><td>{esc(name)}</td><td style='text-align:right'>{esc(pos_txt)}</td></tr>")
            if len(spec_entries) > 6:
                rows.append(f"<tr><td colspan='3' style='text-align:center;color:{label_color}'>+ {len(spec_entries)-6} more�</td></tr>")
            spec_section = f"""
            <div style='height:6px'></div>
            <div style='font-weight:600; color:{label_color}; margin-bottom:2px'>Spectroscopies ({len(spec_entries)})</div>
            <table style='width:100%; border-collapse:collapse' cellspacing='0' cellpadding='2'>
              {''.join(rows)}
            </table>
            """

        scan_entries = [
            ('XScanRange', 'X scan', header.get('XScanRange'), header.get('XPhysUnit', header.get('PhysUnit',''))),
            ('YScanRange', 'Y scan', header.get('YScanRange'), header.get('YPhysUnit', header.get('PhysUnit',''))),
            ('Speed', 'Speed', header.get('Speed'), ''),
            ('LineRate', 'Line rate', header.get('LineRate'), ''),
            ('Angle', 'Angle', header.get('Angle'), 'deg'),
            ('xPixel', 'x pixels', header.get('xPixel'), ''),
            ('yPixel', 'y pixels', header.get('yPixel'), ''),
            ('xCenter', 'x center', header.get('xCenter'), header.get('XPhysUnit', '')),
            ('yCenter', 'y center', header.get('yCenter'), header.get('YPhysUnit', '')),
            ('dzdx', 'dz/dx', header.get('dzdx') or header.get('dz/dx'), ''),
            ('dzdy', 'dz/dy', header.get('dzdy') or header.get('dz/dy'), ''),
            ('overscan[%]', 'Overscan (%)', header.get('overscan[%]'), '%'),
        ]
        scan_rows = []
        for key, label, val, extra_unit in scan_entries:
            if val is None or val == '':
                continue
            if isinstance(val, float):
                val_txt = f"{val:.3f}"
            else:
                val_txt = esc(val)
            unit_txt = extra_unit or ''
            scan_rows.append(f"<tr><td>{esc(label)}</td><td style='text-align:right'>{val_txt} {esc(unit_txt)}</td></tr>")
        scan_section = ""
        if scan_rows:
            scan_section = f"""
            <div style='height:6px'></div>
            <div style='font-weight:600; color:{label_color}; margin-bottom:2px'>Scan metadata</div>
            <table style='width:100%; border-collapse:collapse' cellspacing='0' cellpadding='2'>
              {''.join(scan_rows)}
            </table>
            """

        # key metadata highlight
        x_range = header.get('XScanRange'); y_range = header.get('YScanRange')
        x_unit = header.get('XPhysUnit', header.get('PhysUnit','nm'))
        y_unit = header.get('YPhysUnit', header.get('PhysUnit','nm'))
        xpix = header.get('xPixel') or header.get('XPixel')
        ypix = header.get('yPixel') or header.get('YPixel')
        x_center = header.get('xCenter'); y_center = header.get('yCenter')
        piezo_txt = f"{abs_nm:.3f} nm" if abs_nm is not None else "—"
        date_display = " ".join(t for t in (date, time) if t).strip() or "—"
        size_txt = "—"
        if x_range is not None and y_range is not None:
            size_txt = f"{fmt_number(x_range)} {esc(x_unit)} × {fmt_number(y_range)} {esc(y_unit)}"
        pixel_txt = "—"
        if xpix is not None and ypix is not None:
            pixel_txt = f"{fmt_number(xpix,0)} × {fmt_number(ypix,0)}"
        center_txt = "—"
        if x_center is not None and y_center is not None:
            center_txt = f"{fmt_number(x_center)} / {fmt_number(y_center)} {esc(x_unit)}"
        bias_txt = f"{fmt_number(bias)} {esc(bias_unit)}" if bias is not None else "—"
        setp_txt = f"{fmt_number(setp)} {esc(setp_unit)}" if setp is not None else "—"
        key_rows = [
            ("Acquired", date_display),
            ("Bias", bias_txt),
            ("Setpoint", setp_txt),
            ("Image size", size_txt),
            ("Pixels", pixel_txt),
            ("X/Y center", center_txt),
            ("Piezo Z", piezo_txt),
        ]
        key_section_rows = "".join(
            f"<tr><td style='padding:2px 6px;color:{label_color};font-weight:600'>{esc(lbl)}</td>"
            f"<td style='padding:2px 6px;text-align:right;font-size:14px'><span style='color:{text_color};font-weight:600'>{val or '—'}</span></td></tr>"
            for lbl, val in key_rows if val
        )
        key_section = f"""
        <div style='border:1px solid {accent_border}; border-radius:12px; background:{accent_bg}; padding:8px; margin-bottom:8px;'>
          <table style='width:100%; border-collapse:collapse'>{key_section_rows}</table>
        </div>
        """

        html = f"""
        <div style='font-family:Segoe UI, Arial; font-size:14px; color:{text_color}'>
          <div style='font-weight:600; font-size:16px; margin-bottom:4px'>{esc(filename)} {tag_chip}</div>
          {key_section}
          <table style='width:100%; border-collapse:collapse' cellspacing='0' cellpadding='2'>
            <tr><td style='color:{label_color}'>Date</td><td style='text-align:right'>{esc(date) or '&nbsp;'}</td></tr>
            <tr><td style='color:{label_color}'>Time</td><td style='text-align:right'>{esc(time) or '&nbsp;'}</td></tr>
            <tr><td style='color:{label_color}'>Bias</td><td style='text-align:right'>{'' if bias is None else esc(bias)} {esc(bias_unit)}</td></tr>
            <tr><td style='color:{label_color}'>SetPoint</td><td style='text-align:right'>{'' if setp is None else esc(setp)} {esc(setp_unit)}</td></tr>
            <tr><td style='color:{label_color}'>User</td><td style='text-align:right'>{esc(user)}</td></tr>
          </table>
          <div style='height:6px'></div>
          {spec_section}
          <div style='height:6px'></div>
          <div style='font-weight:600; color:%s; margin-bottom:2px'>Channel</div>
          <table style='width:100%; border-collapse:collapse' cellspacing='0' cellpadding='2'>
            <tr><td style='color:{label_color}'>Index</td><td style='text-align:right'>{channel_idx}</td></tr>
            <tr><td style='color:{label_color}'>Caption</td><td style='text-align:right'>{esc(cap)}</td></tr>
            <tr><td style='color:{label_color}'>Unit (orig)</td><td style='text-align:right'>{esc(phys_orig)}</td></tr>
            <tr><td style='color:{label_color}'>Shown unit</td><td style='text-align:right'><b>{esc(unit_final)}</b></td></tr>
            <tr><td style='color:{label_color}'>Scale</td><td style='text-align:right'>{esc(scale)}</td></tr>
            <tr><td style='color:{label_color}'>Offset</td><td style='text-align:right'>{esc(offset)}</td></tr>
            <tr><td style='color:{label_color}'>Stats</td><td style='text-align:right'>{esc(stats)}</td></tr>
          </table>
          <div style='height:6px'></div>
          {ch_lines}
          {('<div style=\'height:6px\'></div><div style=\'font-weight:600; color:#333; margin-bottom:2px\'>Control params</div>' if params_rows else '')}
          {('<table style=\'width:100%; border-collapse:collapse\' cellspacing=\'0\' cellpadding=\'2\'>' + params_rows + '</table>') if params_rows else ''}
          {scan_section}
        </div>
        """
        return html

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
        if not file_key:
            return None
        header, fds = self.headers.get(str(file_key), (None, None))
        if not header or not fds:
            return None
        if channel_idx < 0 or channel_idx >= len(fds):
            if not fds:
                return None
            channel_idx = min(max(channel_idx, 0), len(fds) - 1)
        fd = fds[channel_idx]
        try:
            data_key, arr = self._get_thumbnail_array(str(file_key), channel_idx, header, fd, width, height)
        except Exception:
            return None
        cache_key = ('frame', data_key, cmap_name)
        pix = self._frame_real_pixmap_cache.get(cache_key)
        if pix is None:
            try:
                qimg = array_to_qimage(arr, cmap_name=cmap_name)
                pix = QtGui.QPixmap.fromImage(qimg)
                self._frame_real_pixmap_cache[cache_key] = pix
            except Exception:
                pix = None
        return pix

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
        sel = str(getattr(self, 'selected_file_for_thumbs', '') or '')
        multi = getattr(self, 'thumb_multi_select', set())
        for fp, w in list(getattr(self, 'thumb_widgets', {}).items()):
            try:
                if str(fp) in multi:
                    w.setStyleSheet("QFrame { border: 2px solid #a36bff; border-radius: 10px; background-color: rgba(163,107,255,40); }")
                elif str(fp) == sel and sel:
                    w.setStyleSheet("QFrame { border: 2px solid #5f8dd3; border-radius: 10px; background-color: rgba(95,141,211,40); }")
                else:
                    w.setStyleSheet("QFrame { border: 1px solid rgba(255,255,255,30); border-radius: 10px; background-color: transparent; }")
            except Exception:
                continue

    def _make_thumb_click_handler(self, label_widget):
        def handler(event):
            if event.button() != QtCore.Qt.LeftButton:
                return
            if self._handle_spec_marker_click(label_widget, event):
                return
            fp = label_widget.property("file_path")
            ch_idx = int(label_widget.property("channel_index"))
            mods = event.modifiers() if event is not None else QtCore.Qt.NoModifier
            if mods & QtCore.Qt.ShiftModifier:
                self._toggle_thumb_multi_selection(fp)
                return
            if mods & QtCore.Qt.ControlModifier:
                self._toggle_thumb_multi_selection(fp)
                return
            self._clear_thumb_multi_selection(update_styles=False)
            self.on_thumbnail_clicked(fp, ch_idx)
        return handler

    def _make_thumb_move_handler(self, label_widget):
        def handler(event):
            if not self._handle_spec_hover(label_widget, event):
                QtWidgets.QLabel.mouseMoveEvent(label_widget, event)
        return handler

    # ---------- thumbnail clicked -> preview + inspector populate ----------
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
        self.last_preview = (str(header_path_str), int(channel_idx))
        if hasattr(self, 'adjust_image_btn'):
            self.adjust_image_btn.setEnabled(True)
        self._update_toolbar_actions(True)
        header_path = Path(header_path_str)
        # track selected file for thumbnail highlighting
        try:
            self.selected_file_for_thumbs = str(header_path)
            self._refresh_thumb_selection_styles()
        except Exception:
            pass
        self._update_frame_map_active(str(header_path))
        file_key = str(header_path)
        header, fds = self.headers.get(file_key, (None,None))
        if header is None or channel_idx < 0 or channel_idx >= len(fds): return
        fd = fds[channel_idx]; fname = fd.get("FileName")
        try:
            xpix = int(header.get('xPixel', 128)); ypix = int(header.get('yPixel', xpix))
            extent = self._header_extent(header)
            unit_final, arr_conv = self._get_filtered_channel_array(file_key, channel_idx, header, fd)
            self._last_base_array = np.asarray(arr_conv)
            self._last_base_extent = extent
            self._last_base_unit = unit_final
            arr_conv, extent = self._apply_adjustments_for_channel(file_key, channel_idx, self._last_base_array, extent)
        except Exception as e:
            self.meta_box.setPlainText("Error reading channel: %s" % str(e)); return

        cmap_to_use = self.preview_cmap_combo.currentText() or self.preview_cmap
        if use_local_cmap:
            cmap_to_use = self.per_file_channel_cmap.get((file_key, channel_idx), cmap_to_use)

        # build views (main + dynamic extras based on current file)
        views = []
        main = {'arr': arr_conv, 'extent': extent, 'cmap': cmap_to_use, 'unit': unit_final, 'title': f"{Path(header_path).name} {fd.get('Caption','')}"}
        views.append(main)

        # Rebuild extra views for the currently selected file using stored specifications
        for spec in getattr(self, 'extra_view_specs', []):
            try:
                # Find matching channel in this file (by caption first, then by index)
                idx2 = self._find_channel_index_for_spec(fds, spec)
                if idx2 is None:
                    continue
                fd2 = fds[idx2]
                unit2_final, arr2_conv = self._get_filtered_channel_array(file_key, idx2, header, fd2)
                cmap2 = self._resolve_extra_spec_cmap(spec, file_key)
                title2 = f"{Path(header_path).name} {fd2.get('Caption','')}"
                views.append({'arr': arr2_conv, 'extent': extent, 'cmap': cmap2, 'unit': unit2_final, 'title': title2})
            except Exception:
                # Skip extra view if anything fails for this file
                continue

        self.preview_canvas.set_views(views)

        # Styled HTML metadata
        try:
            html = self._build_metadata_html(header_path, header, fd, channel_idx, unit_final, arr_conv)
            self.meta_box.setHtml(html)
        except Exception:
            self.meta_box.setPlainText(f"File: {header_path.name}")

    def _header_extent(self, header):
        if not header:
            return None
        try:
            x_range = float(header.get('XScanRange', header.get('ScanRange', 0.0)) or 0.0)
        except Exception:
            x_range = 0.0
        try:
            y_range = float(header.get('YScanRange', header.get('ScanRange', 0.0)) or 0.0)
        except Exception:
            y_range = 0.0
        if x_range > 0 and y_range > 0:
            return [0.0, x_range, y_range, 0.0]
        return None

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

    def _collect_channel_exports(self, header_path_str, main_channel_idx=None):
        header_path = Path(header_path_str)
        file_key = str(header_path)
        header, fds = self.headers.get(file_key, (None, None))
        if header is None or not fds:
            return header, []
        extent = self._header_extent(header)
        exports = []
        channel_idx = main_channel_idx
        if channel_idx is None:
            channel_idx = 0
        if channel_idx < 0 or channel_idx >= len(fds):
            channel_idx = 0
        def _append(idx, cmap=None):
            if idx is None or idx < 0 or idx >= len(fds):
                return
            fd = fds[idx]
            try:
                unit_final, arr_conv = self._get_filtered_channel_array(file_key, idx, header, fd)
            except Exception:
                return
            cap = fd.get('Caption', fd.get('FileName', f"chan{idx}"))
            adj_arr, adj_extent = self._apply_adjustments_for_channel(file_key, idx, arr_conv, extent)
            exports.append({
                'arr': adj_arr,
                'extent': adj_extent,
                'unit': unit_final,
                'caption': cap,
                'idx': idx,
                'cmap': cmap,
                'fd': fd,
            })
        cmap_main = self.per_file_channel_cmap.get((file_key, channel_idx), self.preview_cmap_combo.currentText() or self.preview_cmap)
        _append(channel_idx, cmap_main)
        for spec in getattr(self, 'extra_view_specs', []):
            try:
                idx2 = self._find_channel_index_for_spec(fds, spec)
            except Exception:
                idx2 = None
            if idx2 is None:
                continue
            cmap2 = self._resolve_extra_spec_cmap(spec, file_key)
            _append(idx2, cmap2)
        return header, exports

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
        # Export high-quality PNGs for the currently selected file's visible channels (main + extras)
        if not self.last_preview:
            QtWidgets.QMessageBox.information(self, "No selection", "Select a file/channel first.")
            return
        header_path_str, channel_idx = self.last_preview
        header_path = Path(header_path_str)
        header, exports = self._collect_channel_exports(header_path_str, channel_idx)
        if header is None or not exports:
            QtWidgets.QMessageBox.information(self, "Export", "No channels to export.")
            return

        default_dir = str(getattr(self, 'last_dir', header_path.parent))
        out_dir = QtWidgets.QFileDialog.getExistingDirectory(self, "Select export folder", default_dir)
        if not out_dir:
            return

        # Metadata for naming
        date = self._sanitize_filename_component(header.get('Date', ''))
        time = self._sanitize_filename_component(header.get('Time', ''))
        file_base = self._sanitize_filename_component(Path(header_path_str).stem)

        # Save each channel as a separate high-DPI PNG
        from matplotlib.figure import Figure
        for item in exports:
            try:
                fig = Figure(figsize=(6, 5), dpi=300)
                ax = fig.add_subplot(1,1,1)
                arr = np.asarray(item['arr'])
                cmapname = item.get('cmap', 'viridis')
                extent = item.get('extent')
                if extent is None:
                    im = ax.imshow(arr, origin='upper', interpolation='nearest', cmap=cmapname)
                else:
                    im = ax.imshow(arr, extent=extent, origin='upper', interpolation='nearest', aspect='equal', cmap=cmapname)
                unit = item.get('unit') or ''
                if unit:
                    cbar = fig.colorbar(im, ax=ax, fraction=0.08, pad=0.02)
                    cbar.set_label(unit)
                ax.set_title(item.get('caption') or '')
                try:
                    fig.tight_layout()
                except Exception:
                    pass

                chan_name = self._sanitize_filename_component(item.get('caption') or f"chan{item.get('idx',0)}")
                parts = [p for p in (chan_name, file_base, date, time) if p]
                fname = "__".join(parts) + ".png"
                out_path = str(Path(out_dir) / fname)
                fig.savefig(out_path, dpi=300, bbox_inches='tight')
            except Exception as e:
                # keep going for other channels
                print('Export failed for a channel:', e)

        QtWidgets.QMessageBox.information(self, "Export", f"Exported {len(exports)} PNG(s) to\n{out_dir}")

    def on_export_xyz_files(self):
        targets = list(getattr(self, 'thumb_multi_select', set()))
        if not targets:
            if getattr(self, 'selected_file_for_thumbs', None):
                targets = [self.selected_file_for_thumbs]
            elif self.last_preview:
                targets = [self.last_preview[0]]
        if not targets:
            QtWidgets.QMessageBox.information(self, "Export", "No thumbnails selected.")
            return
        out_dir = QtWidgets.QFileDialog.getExistingDirectory(self, "Select folder for XYZ export", str(self.last_dir))
        if not out_dir:
            return
        out_dir = Path(out_dir)
        try:
            out_dir.mkdir(parents=True, exist_ok=True)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Export", f"Cannot create folder: {exc}")
            return
        exported = []
        channel_idx = self.channel_dropdown.currentIndex()
        for file_key in targets:
            header, exports = self._collect_channel_exports(file_key, channel_idx)
            if header is None or not exports:
                log_status(f"[XYZ Export] No channels for {file_key}")
                continue
            header_path = Path(file_key)
            for item in exports:
                arr_si, z_unit = convert_to_si(item['arr'], item.get('unit'))
                if not z_unit:
                    z_unit = item.get('unit') or 'arb.'
                extent = item.get('extent')
                x_vals, y_vals, x_unit, y_unit = self._axes_from_extent(header, arr_si.shape, extent)
                date_token = self._sanitize_filename_component(header.get('Date', ''))
                time_token = self._sanitize_filename_component(header.get('Time', ''))
                base_name = self._sanitize_filename_component(header_path.stem)
                chan_token = self._sanitize_filename_component(item.get('caption') or f"chan{item.get('idx')}")
                parts = [p for p in (chan_token, base_name, date_token, time_token) if p]
                fname = "__".join(parts) + ".xyz"
                full_path = out_dir / fname
                meta_lines = [
                    f"Source file: {header_path.name}",
                    f"Channel: {item.get('caption') or ''} (index {item.get('idx')})",
                    f"Date: {header.get('Date', '')} Time: {header.get('Time', '')}",
                    f"Bias: {header.get('Bias', '')} {header.get('BiasPhysUnit', '')}",
                    f"Dimensions: {header.get('xPixel','?')} x {header.get('yPixel','?')} pixels",
                    f"X range: {header.get('XScanRange', header.get('ScanRange','?'))} {header.get('XPhysUnit','')}",
                    f"Y range: {header.get('YScanRange', header.get('ScanRange','?'))} {header.get('YPhysUnit','')}",
                ]
                try:
                    self._write_xyz_file(full_path, x_vals, y_vals, arr_si, x_unit, y_unit, z_unit, meta_lines)
                    exported.append(str(full_path))
                except Exception as exc:
                    QtWidgets.QMessageBox.warning(self, "Export", f"Failed to export {fname}: {exc}")
                    log_status(f"[XYZ Export] Failed {full_path}: {exc}")
        if not exported:
            QtWidgets.QMessageBox.information(self, "Export", "No XYZ files were created.")
        else:
            preview = "\n".join(exported[:5])
            if len(exported) > 5:
                preview += "\n..."
            QtWidgets.QMessageBox.information(self, "Export", f"Exported {len(exported)} XYZ file(s) to {out_dir}:\n{preview}")

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
        dlg = ImageAdjustDialog(self, base_arr, spec, spec.get('cmap', current_cmap))
        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            new_spec = dlg.current_spec
            self._set_adjust_spec(file_key, channel_idx, new_spec)
            new_cmap = dlg.cmap_combo.currentText()
            if new_cmap:
                self.per_file_channel_cmap[(str(file_key), int(channel_idx))] = new_cmap
            self.show_file_channel(file_key, channel_idx)

    def render_and_save_file_using_config(self, header_path, config, out_dir):
        """
        Render the given file using the supplied config (as returned by get_current_detail_config)
        and save a multi-panel PNG. Returns a list with the saved file path.
        """
        header_path = Path(header_path)
        out_dir = Path(out_dir)
        out_dir.mkdir(parents=True, exist_ok=True)
        header, fds = self.headers.get(str(header_path), (None, None))
        if header is None or fds is None:
            header, fds = parse_header(header_path)
            self.headers[str(header_path)] = (header, fds)
        try:
            xpix = int(header.get('xPixel', 128))
            ypix = int(header.get('yPixel', xpix))
        except Exception:
            xpix = 128; ypix = 128
        XScanRange = float(header.get('XScanRange', 0.0)) if header.get('XScanRange') else None
        YScanRange = float(header.get('YScanRange', 0.0)) if header.get('YScanRange') else None
        extent = [0.0, float(XScanRange), float(YScanRange), 0.0] if (XScanRange and YScanRange) else None
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
            v_range = config.get('vmin_vmax', {}).get(key)
            vmin = vmax = None
            if isinstance(v_range, (list, tuple)) and len(v_range) == 2:
                vmin, vmax = v_range
            render_items.append({'arr': arr_conv, 'extent': extent, 'unit': unit_final, 'label': label, 'cmap': cmap, 'vmin': vmin, 'vmax': vmax})
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
            im = ax.imshow(item['arr'], extent=item['extent'], origin='upper', interpolation='nearest',
                           aspect='equal' if item['extent'] else 'auto', cmap=item['cmap'],
                           vmin=item['vmin'], vmax=item['vmax'])
            ax.set_title(item['label'], fontsize=9)
            ax.tick_params(labelsize=8)
            if item['unit']:
                cbar = fig.colorbar(im, ax=ax, fraction=0.08, pad=0.02)
                cbar.set_label(item['unit'])
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

    # ---------- Profile measurement (interactive line) ----------
    def _on_start_profile(self):
        # toggle interactive line profile mode
        views = getattr(self.preview_canvas, 'views', [])
        if not views:
            QtWidgets.QMessageBox.information(self, "Measure profile", "No image to measure. Load a channel first.")
            return
        active = getattr(self.preview_canvas, 'profile_enabled', False)
        if not active:
            # enter profile mode
            self.preview_canvas.set_profile_callback(self._on_profile_updated)
            self.preview_canvas.enable_profile(True)
            try: self.measure_profile_btn.setText('Exit profile')
            except Exception: pass
            self.meta_box.setPlainText("Profile mode: drag the yellow endpoints on the main image. Close to exit.")
            self._profile_dialog = None
        else:
            # exit profile mode
            self.preview_canvas.enable_profile(False)
            try: self.measure_profile_btn.setText('Measure profile')
            except Exception: pass
            try:
                if hasattr(self, '_profile_dialog') and self._profile_dialog is not None:
                    self._profile_dialog.close()
            except Exception:
                pass

    def _on_profile_updated(self, x_px, vals, length_nm, unit):
        # create or update a persistent profile dialog
        try:
            if not hasattr(self, '_profile_dialog') or self._profile_dialog is None:
                self._profile_dialog = ProfileDialog(x_px, vals, length_nm=length_nm, parent=self, unit=unit)
                self._profile_dialog.show()
            else:
                self._profile_dialog.update_data(x_px, vals, length_nm=length_nm)
        except Exception:
            pass

    def _on_view_copied(self, view):
        title = view.get('title') or 'View'
        msg = f"Copied '{title}' to clipboard"
        try:
            QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), msg)
        except Exception:
            pass

    def _on_preview_value(self, value, x, y, view):
        if value is None or view is None:
            self.preview_value_label.setText("Value: --")
            return
        unit = view.get('unit') or ''
        title = view.get('title') or ''
        text = f"{title}: {value:.4g}"
        if unit:
            text += f" {unit}"
        self.preview_value_label.setText(text)

    # ---------- manual tagging (still available) ----------
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
        try:
            folder = getattr(self, 'spec_folder_path', None) or self.last_dir
            folder = Path(folder)
        except Exception:
            folder = self.last_dir
        log_status(f"Scanning spectroscopy files in: {folder}")
        if not self.show_spectra:
            self.spectros = []
            self.spectros_by_image = defaultdict(list)
            self._clear_multi_spec_selection()
            return
        self.spectros = self._scan_spectros(folder)
        log_status(f"Loaded {len(self.spectros)} spectroscopy entries")
        self._assign_spectros_to_images()
        self.matrix_spectros = [spec for spec in self.spectros if spec.get('matrix_index') is not None]
        self._clear_multi_spec_selection()
        if refresh:
            self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())
            if self.last_preview:
                self.show_file_channel(self.last_preview[0], self.last_preview[1])

    def _scan_spectros(self, folder:Path):
        specs = []
        if not folder or not Path(folder).exists():
            return specs
        patterns = ("*.dat","*.DAT","*.txt","*.TXT")
        cache = self._spectro_cache
        seen_keys = set()
        files = []
        for pat in patterns:
            files.extend(sorted(folder.glob(pat)))
        total = len(files)
        if total:
            log_status(f"Scanning {total} spectroscopy file(s)...")
        progress_step = max(1, total // 20) if total else 1
        for idx, f in enumerate(files, 1):
            p = Path(f)
            if p.is_dir():
                continue
            key = str(p)
            if key in seen_keys:
                continue
            seen_keys.add(key)
            try:
                mtime = p.stat().st_mtime
            except Exception:
                mtime = 0.0
            cached = cache.get(key)
            if cached and abs(cached.get('mtime', 0.0) - mtime) <= 1e-6:
                spec_list = cached.get('data') or []
            else:
                try:
                    spec_list = parse_spectroscopy_file(p)
                except Exception:
                    continue
                cache[key] = {'mtime': mtime, 'data': spec_list}
            specs.extend(spec_list or [])
            if total and (idx % progress_step == 0 or idx == total):
                pct = idx / total * 100.0
                log_status(f"  - spectroscopy load {idx}/{total} ({pct:4.0f}%)")
        stale = [k for k in list(cache.keys()) if k not in seen_keys]
        for k in stale:
            cache.pop(k, None)
        specs.sort(key=lambda s: s.get('time') or datetime.min)
        return specs

    def _assign_spectros_to_images(self):
        """Assign spectroscopy entries to images, using a single pass over time-sorted lists when possible."""
        self.spectros_by_image = defaultdict(list)
        images = list(getattr(self, 'image_meta', []) or [])
        specs = list(self.spectros or [])
        if not images or not specs:
            return
        try:
            images.sort(key=lambda img: img.get('time') or datetime.min)
        except Exception:
            pass
        try:
            specs.sort(key=lambda s: s.get('time') or datetime.min)
        except Exception:
            pass
        img_idx = 0
        n_img = len(images)
        for spec in specs:
            st = spec.get('time')
            if st is None:
                match = self._match_spec_to_image_by_hint(spec, images)
            else:
                while img_idx + 1 < n_img and (images[img_idx + 1].get('time') or datetime.max) <= st:
                    img_idx += 1
                match = images[img_idx] if 0 <= img_idx < n_img else None
                if match is None:
                    match = self._match_spec_to_image_by_hint(spec, images)
            if not match:
                continue
            image_key = str(match['path'])
            spec['image_key'] = image_key
            self.spectros_by_image[image_key].append(spec)
        for k in list(self.spectros_by_image.keys()):
            self.spectros_by_image[k].sort(key=lambda s: s.get('time') or datetime.min)

    def _match_spec_to_image_by_hint(self, spec, images):
        def normalize(stem):
            stem = stem.lower().strip()
            stem = re.sub(r'(?:_matrix|-matrix).*$', '', stem)
            stem = stem.replace('-', '_')
            return stem
        spec_stem = normalize(Path(spec.get('path', '')).stem)
        if not spec_stem:
            return None
        spec_tokens = [tok for tok in spec_stem.split('_') if tok]
        best = None
        best_score = -1
        for img in images:
            img_stem = normalize(Path(img['path']).stem)
            img_tokens = [tok for tok in img_stem.split('_') if tok]
            score = 0
            for a, b in zip(spec_tokens, img_tokens):
                if a == b:
                    score += 10
                else:
                    break
            common_prefix = 0
            for a, b in zip(spec_stem, img_stem):
                if a == b:
                    common_prefix += 1
                else:
                    break
            score += common_prefix
            if spec_stem in img_stem or img_stem in spec_stem:
                score += 50
            if score > best_score:
                best_score = score
                best = img
        return best

    def _map_spec_to_pixels(self, spec, header, xpix, ypix):
        try:
            x = float(spec.get('x'))
            y = float(spec.get('y'))
        except Exception:
            return None
        if x is None or y is None:
            return None
        try:
            x_center = float(header.get('xCenter', 0.0))
            y_center = float(header.get('yCenter', 0.0))
            x_range = float(header.get('XScanRange', header.get('ScanRange', 0.0)) or 0.0)
            y_range = float(header.get('YScanRange', header.get('ScanRange', 0.0)) or 0.0)
        except Exception:
            return None
        if x_range <= 0 or y_range <= 0:
            return None
        xmin = x_center - x_range/2.0
        ymin = y_center - y_range/2.0
        xmax = x_center + x_range/2.0
        ymax = y_center + y_range/2.0
        if xmax == xmin or ymax == ymin:
            return None
        frac_x = (x - xmin) / (xmax - xmin)
        frac_y = (ymax - y) / (ymax - ymin)
        if not (0.0 <= frac_x <= 1.0 and 0.0 <= frac_y <= 1.0):
            return self._map_spec_by_grid(spec, xpix, ypix)
        cols = max(1, int(xpix) - 1)
        rows = max(1, int(ypix) - 1)
        col = frac_x * cols
        row = frac_y * rows
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

    def _draw_spectro_markers_on_pixmap(self, pixmap, header, file_key, xpix, ypix):
        if not self.show_spectra:
            return []
        specs = self.spectros_by_image.get(file_key, [])
        if not specs:
            return []
        markers = []
        painter = QtGui.QPainter(pixmap)
        painter.setRenderHint(QtGui.QPainter.Antialiasing, True)
        w_scale = pixmap.width() / max(1, xpix - 1)
        h_scale = pixmap.height() / max(1, ypix - 1)
        for idx, spec in enumerate(specs, 1):
            coords = self._map_spec_to_pixels(spec, header, xpix, ypix)
            if coords is None:
                continue
            col, row = coords
            x = col * w_scale
            y = row * h_scale
            radius = 9
            center = QtCore.QPointF(x, y)
            painter.setBrush(QtGui.QColor(70, 150, 220, 220))
            pen = QtGui.QPen(QtGui.QColor(20, 20, 20))
            pen.setWidth(1)
            painter.setPen(pen)
            painter.drawEllipse(center, radius, radius)
            painter.setPen(QtGui.QPen(QtGui.QColor(255, 255, 255)))
            painter.setFont(QtGui.QFont("Segoe UI", 8, QtGui.QFont.Bold))
            painter.drawText(QtCore.QRectF(x-radius, y-radius, radius*2, radius*2), QtCore.Qt.AlignCenter, str(idx))
            markers.append({'rect': QtCore.QRectF(x-radius, y-radius, radius*2, radius*2), 'spec': spec, 'label': idx})
        painter.end()
        return markers

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
        for info in markers:
            rect = info.get('rect')
            if rect and rect.contains(x, y):
                mods = event.modifiers() if event is not None else QtCore.Qt.NoModifier
                if mods & QtCore.Qt.ShiftModifier:
                    self._toggle_multi_spec_selection(info.get('spec'))
                else:
                    self._clear_multi_spec_selection()
                    self._open_spectroscopy_popup(info.get('spec'))
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
                spec = info.get('spec') or {}
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
        if not spec:
            return
        try:
            dlg = SpectroscopyPopup(spec, parent=self)
            dlg.show()
            self._spectro_popups.append(dlg)
            dlg.finished.connect(lambda _: self._spectro_popups.remove(dlg) if dlg in self._spectro_popups else None)
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Spectroscopy", str(e))

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
        path = str(file_path)
        if not hasattr(self, 'thumb_multi_select'):
            self.thumb_multi_select = set()
        if path in self.thumb_multi_select:
            self.thumb_multi_select.remove(path)
        else:
            self.thumb_multi_select.add(path)
        self._refresh_thumb_selection_styles()

    def _clear_thumb_multi_selection(self, update_styles=True):
        self.thumb_multi_select = set()
        if update_styles:
            self._refresh_thumb_selection_styles()

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
        specs = list(self._multi_spec_selection)
        if len(specs) < 2:
            return
        # close previous multi popups
        for dlg in list(self._multi_spectro_popups):
            try:
                dlg.close()
            except Exception:
                pass
        self._multi_spectro_popups = []
        dlg = SpectroscopyCompareDialog(specs, parent=self)
        dlg.show()
        self._multi_spectro_popups.append(dlg)
        dlg.finished.connect(lambda _: self._multi_spectro_popups.remove(dlg) if dlg in self._multi_spectro_popups else None)

    def on_show_matrix_spectro_viewer(self):
        matrix_files = defaultdict(list)
        for spec in self.matrix_spectros:
            matrix_files[str(spec.get('path'))].append(spec)
        if not matrix_files:
            QtWidgets.QMessageBox.information(self, "Matrix spectra", "No matrix spectroscopy files detected for this folder.")
            return
        choices = sorted(matrix_files.items(), key=lambda item: Path(item[0]).name.lower())
        names = [Path(dat_path).name for dat_path, _ in choices]
        item, ok = QtWidgets.QInputDialog.getItem(self, "Matrix spectroscopies", "Select matrix file:", names, 0, False)
        if not ok or not item:
            return
        dat_key = None; target_specs = None
        for k, specs in choices:
            if Path(k).name == item:
                dat_key = k
                target_specs = specs
                break
        if not dat_key or not target_specs:
            return
        images = getattr(self, 'image_meta', [])
        if not images:
            QtWidgets.QMessageBox.information(self, "Matrix spectra", "No SXM images available to anchor matrix data.")
            return
        images = getattr(self, 'image_meta', [])
        if not images:
            QtWidgets.QMessageBox.information(self, "Matrix spectra", "No SXM images available to anchor matrix data.")
            return
        try:
            matrix_time = datetime.fromtimestamp(Path(dat_key).stat().st_mtime)
        except Exception:
            matrix_time = None
        if matrix_time is None:
            for spec in target_specs:
                if spec.get('time'):
                    matrix_time = spec.get('time')
                    break
        base_name = _matrix_base_name(Path(dat_key).stem).lower()
        candidates = [img for img in images if _matrix_base_name(Path(img['path']).stem).lower() == base_name]
        match = None
        if candidates:
            earlier = [img for img in candidates if img.get('time') and matrix_time and img['time'] <= matrix_time]
            if earlier:
                earlier.sort(key=lambda img: img['time'], reverse=True)
                match = earlier[0]
            else:
                candidates.sort(key=lambda img: abs((img.get('time') or datetime.min) - (matrix_time or datetime.min)))
                match = candidates[0]
        if not match:
            match = find_last_image_for_spec(matrix_time, images)
        if not match:
            QtWidgets.QMessageBox.warning(self, "Matrix spectra", "Could not find a preceding SXM image for this matrix file.")
            return
        entry = {'path': Path(match['path']), 'time': match.get('time')}
        dlg = MatrixSpectroViewer(self, entry, target_specs)
        dlg.show()
        self._popup_refs.append(dlg)
        dlg.finished.connect(lambda _: self._popup_refs.remove(dlg) if dlg in self._popup_refs else None)

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

    def on_dark_mode_toggled(self, checked: bool):
        self.dark_mode = bool(checked)
        self.config['dark_mode'] = self.dark_mode; save_config(self.config)
        self._apply_dark_mode(self.dark_mode)
        if self.last_preview:
            self.show_file_channel(self.last_preview[0], self.last_preview[1])

    # ---------- control callbacks ----------
    def on_channel_dropdown_changed(self, idx):
        self.last_channel_index = int(idx); self.config['last_channel_index'] = self.last_channel_index; save_config(self.config)
        self.populate_thumbnails_for_channel(idx)

    def on_thumb_cmap_changed(self, idx):
        self.thumb_cmap = self.thumb_cmap_combo.currentText(); self.config['thumbnail_cmap'] = self.thumb_cmap; save_config(self.config)
        self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())

    def on_preview_cmap_changed(self, idx):
        self.preview_cmap = self.preview_cmap_combo.currentText(); self.config['preview_cmap'] = self.preview_cmap; save_config(self.config)
        if self.last_preview: self.show_file_channel(self.last_preview[0], self.last_preview[1])

    def on_show_spectra_toggled(self, checked):
        self.show_spectra = bool(checked)
        self.config['show_spectra'] = self.show_spectra; save_config(self.config)
        if self.show_spectra:
            self._reload_spectros(refresh=False)
        else:
            self.spectros = []
            self.spectros_by_image = defaultdict(list)
            self._clear_multi_spec_selection()
        self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())
        if self.last_preview:
            self.show_file_channel(self.last_preview[0], self.last_preview[1])

    def on_export_selected_same_view(self):
        targets = list(getattr(self, 'thumb_multi_select', set()))
        if not targets:
            if getattr(self, 'selected_file_for_thumbs', None):
                targets = [self.selected_file_for_thumbs]
            elif self.last_preview:
                targets = [self.last_preview[0]]
        if not targets:
            QtWidgets.QMessageBox.information(self, "Export", "No thumbnails selected.")
            return
        config = self.get_current_detail_config()
        if not config.get('channels'):
            QtWidgets.QMessageBox.information(self, "Export", "No channels configured to export.")
            return
        out_dir = QtWidgets.QFileDialog.getExistingDirectory(self, "Select export folder", str(self.last_dir))
        if not out_dir:
            return
        worker = BatchExportWorker(self, targets, config, out_dir)
        worker.signals.progress.connect(self._on_batch_export_progress)
        worker.signals.finished.connect(self._on_batch_export_finished)
        self._batch_export_worker = worker
        progress = QtWidgets.QProgressDialog("Exporting...", "Cancel", 0, len(targets), self)
        progress.setWindowTitle("Batch export")
        progress.setWindowModality(QtCore.Qt.WindowModal)
        progress.canceled.connect(worker.cancel)
        progress.show()
        self._batch_export_progress = progress
        QtCore.QThreadPool.globalInstance().start(worker)

    def _on_batch_export_progress(self, current, total, path):
        dlg = getattr(self, '_batch_export_progress', None)
        if dlg is None:
            return
        dlg.setMaximum(total)
        dlg.setValue(current)
        dlg.setLabelText(f"Exporting {Path(path).name} ({current}/{total})")

    def _on_batch_export_finished(self, saved_paths, errors, cancelled):
        dlg = getattr(self, '_batch_export_progress', None)
        if dlg is not None:
            dlg.close()
            self._batch_export_progress = None
        self._batch_export_worker = None
        msg_lines = [f"Saved {len(saved_paths)} file(s)."]
        if saved_paths:
            preview_paths = "\n".join(saved_paths[:5])
            msg_lines.append(preview_paths + ("\n..." if len(saved_paths) > 5 else ""))
        if cancelled:
            msg_lines.append("Operation cancelled.")
        if errors:
            msg_lines.append("Errors:\n" + "\n".join(errors[:10]))
        QtWidgets.QMessageBox.information(self, "Batch export", "\n".join(msg_lines))

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

