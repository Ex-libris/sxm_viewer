"""Detail canvases and spectroscopy dialogs."""
from __future__ import annotations

import copy
import itertools
import json
import math

import numpy as np
from matplotlib import patches
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from matplotlib.ticker import AutoMinorLocator, FuncFormatter, MaxNLocator
from matplotlib.widgets import RectangleSelector
from mpl_toolkits.axes_grid1.anchored_artists import AnchoredSizeBar
from mpl_toolkits.axes_grid1.inset_locator import inset_axes

from ..._shared import (
    QtCore,
    QtGui,
    QtWidgets,
    QIcon,
    QPixmap,
    QImage,
    QPainter,
    QPen,
    QBrush,
    FigureCanvas,
    Figure,
    Line2D,
    colormaps,
    np,
    Path,
    defaultdict,
    OrderedDict,
    datetime,
    hashlib,
    itertools,
    io,
    json,
    math,
    os,
    sys,
    threading,
    _scipy_ndimage,
    log_status,
    matplotlib,
)
from ..plot_typography import add_font_menu_action, normalize_font_family, apply_text_style, apply_qfont_style
from ..figure_layout_presets import (
    iter_figure_layout_presets,
    get_figure_layout_preset,
    apply_figure_layout,
    preset_pixel_size,
    apply_canvas_widget_preset,
    copy_figure_to_clipboard,
    save_figure_with_dialog,
)
from ...config import (
    CONFIG_PATH,
    HEADER_CACHE_PATH,
    HEADER_CACHE_VERSION,
    CH_EQUALITY_TOL_NM,
    CH_SAMPLE_POINTS,
    CHANNEL_DATA_CACHE_LIMIT,
    FILTERED_CACHE_LIMIT,
    THUMB_DISK_CACHE_DIR,
    load_config,
    save_config,
    load_header_cache,
    save_header_cache,
)
from ...data.io import (
    parse_header,
    read_channel_file,
    normalize_unit_and_data,
    _split_key_value,
    _coerce_value,
    _canonical_header_key,
    _parse_inline_channels,
    _trailing_digits,
    _load_ascii_grid,
    _load_binary_grid,
    _load_tokenized_grid,
    _load_binary_with_inference,
    _binary_dtype_candidates,
)
from ...data.spectroscopy import (
    parse_spectroscopy_file,
    fit_parabola_bias,
    find_last_image_for_spec,
    _matrix_base_name,
    _rows_to_spec,
    _channel_labels,
    _clean_channel_label,
    _normalize_bias_axis,
    _extract_meta,
    _guess_index_from_name,
    _extract_section_value,
    _parse_section_metadata,
    _split_key_value,
    _split_tokens,
    _split_header_columns,
    _row_is_numeric,
    _normalize_meta_key,
    _coerce_value,
    _maybe_float,
    _maybe_int,
    _parse_datetime,
    _parse_date_and_time,
    _mtime,
    _read_text,
)
from ..thumbnail_render import (
    array_to_qimage,
    _ThumbnailJobSignals,
    _ThumbnailJob,
    _colormap_icon,
    convert_to_si,
    _unit_to_nm_factor,
    _value_in_nm,
    robust_limits,
    _interp_index,
    sample_array_value,
    apply_adjustment_spec,
    _rotate_extent_box,
    _trim_nan_border,
    save_wsxm_xyz,
)
from ..canvases.detail_preview import MultiPreviewCanvas, SafeFigureCanvas
from ..palettes import DEFAULT_COLOR_CYCLE, get_color_cycle, list_color_cycles
from ..profile_links import (
    register_profile_dialog,
    unregister_profile_dialog,
    apply_live_profile_style,
    profile_ref_key,
)
from .profile_data import axis_label, format_marker_delta, format_stats_text, fmt_length

_PROFILE_COMPOSITE_MIME = "application/x-sxm-profile-composite"


class _ProfileCompositeDragButton(QtWidgets.QToolButton):
    """Drag button used to compose profile dialogs without interfering with plot gestures."""

    def __init__(self, owner):
        super().__init__(owner)
        self._owner_dialog = owner
        self._drag_start_pos = None
        self._drag_started = False
        self.setObjectName("profileComposeButton")
        self.setText("Combine")
        self.setCursor(QtCore.Qt.OpenHandCursor)
        self.setAutoRaise(False)
        self.setToolButtonStyle(QtCore.Qt.ToolButtonTextOnly)
        self.setToolTip(
            "Drag this onto another profile window to create a new composite window."
        )

    def mousePressEvent(self, event):
        if event.button() == QtCore.Qt.LeftButton:
            self._drag_start_pos = event.pos()
            self._drag_started = False
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if (
            self._drag_start_pos is not None
            and event.buttons() & QtCore.Qt.LeftButton
            and (event.pos() - self._drag_start_pos).manhattanLength() >= QtWidgets.QApplication.startDragDistance()
        ):
            self._drag_start_pos = None
            self._drag_started = True
            try:
                self._owner_dialog.start_profile_composite_drag()
            except Exception:
                pass
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if (
            event.button() == QtCore.Qt.LeftButton
            and not self._drag_started
            and hasattr(self._owner_dialog, "_show_compose_help")
        ):
            try:
                self._owner_dialog._show_compose_help(self.mapToGlobal(event.pos()))
            except Exception:
                pass
        self._drag_start_pos = None
        self._drag_started = False
        super().mouseReleaseEvent(event)


class ProfileDialog(QtWidgets.QDialog):
    """Dialog showing the sampled profile and basic stats."""
    def __init__(self, active_profile, saved_profiles=None, parent=None, unit=None, y_label=None,
                 activate_overlay_callback=None, highlight_overlay_callback=None,
                 label_scale_callback=None, delete_overlay_callback=None,
                 marker_update_callback=None, marker_select_callback=None,
                 add_overlay_callback=None, style_update_callback=None,
                 palette_callback=None, dark_mode=False):
        super().__init__(parent)
        self.setWindowTitle('Profile measurement')
        self.setAcceptDrops(True)
        self.setWindowFlags(
            self.windowFlags()
            | QtCore.Qt.WindowMinimizeButtonHint
            | QtCore.Qt.WindowSystemMenuHint
        )
        self.resize(900, 600)
        self.setMinimumSize(700, 450)
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
        self._delete_overlay_cb = delete_overlay_callback
        self._marker_update_cb = marker_update_callback
        self._marker_key_cb = marker_select_callback
        self._add_overlay_cb = add_overlay_callback
        self._style_update_cb = style_update_callback
        self._palette_cb = palette_callback
        self._profile_palette_name = DEFAULT_COLOR_CYCLE
        self._marker_syncing = False
        self._marker_positions_by_key = {}
        self._marker_domain_by_key = {}
        self._current_marker_key = None
        self._last_saved_count = 0
        self._line_handles_by_key = {}
        self._ordered_profile_entries_cache = []
        self._toggle_buttons = []
        self._advanced_controls_visible = False
        self._legend_visible = True
        self._legend_loc = 'upper right'
        self._legend_fontsize = 8.0
        self._legend_frame_fill_visible = True
        self._legend_outline_visible = True
        self._legend_outline_width = 1.0
        self._legend_custom_anchor = None
        self._legend_artist = None
        self._legend_drag = None
        self._figure_preset_key = "interactive"
        self._metadata_visible = False
        self._metadata_show_filename = True
        self._metadata_show_acquisition = True
        self._metadata_show_time = False
        self._metadata_show_folder_name = False
        self._metadata_show_folder = False
        self._metadata_artist = None
        self._owner = parent
        self._workspace_registered = False
        self._composite_mode = False
        self._composite_origin_id = hex(id(self))
        self._canvas_drag_start_pos = None
        self._canvas_drag_started = False
        owner = self._owner
        self._plot_font_family = normalize_font_family(getattr(owner, "_plot_font_family", None), "sans-serif")
        self._plot_font_bold = bool(getattr(owner, "_plot_font_bold", False))
        self._plot_font_italic = bool(getattr(owner, "_plot_font_italic", False))
        self._plot_font_underline = bool(getattr(owner, "_plot_font_underline", False))
        v = QtWidgets.QVBoxLayout()
        fig = Figure(figsize=(6,3))
        self.canvas = SafeFigureCanvas(fig)
        self.canvas.installEventFilter(self)
        self.canvas.setToolTip(
            "Drag the plot margin onto another profile window to create a composite."
        )
        self.ax = fig.add_subplot(111)
        self.canvas.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.canvas.customContextMenuRequested.connect(self._on_context_menu)
        self.ax_top = self.ax.twiny()
        self.ax_top.set_visible(False)
        self.ax_right = self.ax.twinx()
        self.ax_right.set_visible(False)
        self._relative_axes = True
        self._font_scale = 1.0
        self.ax.set_xlabel(axis_label('px'))
        splitter = QtWidgets.QSplitter(QtCore.Qt.Vertical)
        self._splitter = splitter
        plot_splitter = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        self._plot_splitter = plot_splitter
        plot_splitter.addWidget(self.canvas)
        # --- Preview panel disabled (commented out) ---
        # The original implementation created a full Matplotlib-based preview inside this dialog
        # (a `MultiPreviewCanvas`) which duplicated rendering work from the main preview. In some
        # environments this doubles the CPU/GPU and memory load (matplotlib figures, colorbars and
        # event callbacks), making the profile dialog expensive to open and use. The block below is
        # intentionally commented out to save resources. To re-enable the preview, uncomment the
        # block and remove the lightweight placeholder that follows.
        #
        # context_widget = QtWidgets.QWidget()
        # self._context_widget = context_widget
        # context_layout = QtWidgets.QVBoxLayout(context_widget)
        # context_layout.setContentsMargins(4, 4, 4, 4)
        # context_title = QtWidgets.QLabel("Preview")
        # context_title.setAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
        # context_layout.addWidget(context_title)
        # self.context_canvas = MultiPreviewCanvas(parent=context_widget)
        # self.context_canvas.setMinimumWidth(320)
        # context_layout.addWidget(self.context_canvas, 1)
        # plot_splitter.addWidget(context_widget)
        # plot_splitter.setStretchFactor(0, 3)
        # plot_splitter.setStretchFactor(1, 2)
        # plot_splitter.setSizes([700, 500])
        # Instead of the heavy preview we previously added a lightweight placeholder widget to indicate
        # the preview is disabled. To avoid reserving any dialog space for this optional preview, the
        # placeholder is intentionally commented out and not added to the layout below. To restore the
        # debug placeholder (or re-enable a lightweight preview) in the future, uncomment the lines
        # below and the placeholder will appear.
        # placeholder = QtWidgets.QLabel("Preview disabled to reduce resource usage", alignment=QtCore.Qt.AlignCenter)
        # placeholder.setMinimumWidth(320)
        # placeholder.setStyleSheet("color: #999;")
        # keep attributes present but empty so other code won't fail if it references them
        self._context_widget = None
        self.context_canvas = None
        # (placeholder not added to layout to avoid occupying space)

        splitter.addWidget(plot_splitter)
        info_widget = QtWidgets.QWidget()
        info_layout = QtWidgets.QVBoxLayout(info_widget)
        info_layout.setContentsMargins(0, 0, 0, 0)
        self.stats = QtWidgets.QLabel("")
        self.stats.setAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignTop)
        self.stats.setWordWrap(True)
        self.stats.setVisible(False)
        self.marker_info = QtWidgets.QLabel("")
        self.marker_info.setAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignTop)
        self.marker_info.setVisible(False)
        self.controls_hint = QtWidgets.QLabel(
            "Shortcuts: V markers, G grid, L lines, P points, Del remove overlay, Ctrl+Wheel font size"
        )
        self.controls_hint.setObjectName("profileControlsHint")
        self.controls_hint.setWordWrap(True)
        self.controls_hint.setVisible(False)

        controls_panel = QtWidgets.QWidget()
        controls_layout = QtWidgets.QVBoxLayout(controls_panel)
        controls_layout.setContentsMargins(0, 0, 0, 0)
        controls_layout.setSpacing(6)

        primary_row = QtWidgets.QHBoxLayout()
        primary_row.setContentsMargins(0, 0, 0, 0)
        primary_row.setSpacing(6)

        self.marker_toggle = self._make_toggle_button(
            "Markers", checked=True, tooltip="Show/hide draggable measurement markers"
        )
        self.marker_toggle.toggled.connect(self._on_marker_toggle)
        primary_row.addWidget(self.marker_toggle)

        self.show_lines_cb = self._make_toggle_button(
            "Lines", checked=True, tooltip="Show connecting profile line"
        )
        self.show_lines_cb.toggled.connect(self._on_plot_option_changed)
        primary_row.addWidget(self.show_lines_cb)

        self.show_points_cb = self._make_toggle_button(
            "Points", checked=False, tooltip="Show sampled data points"
        )
        self.show_points_cb.toggled.connect(self._on_plot_option_changed)
        primary_row.addWidget(self.show_points_cb)

        self.grid_cb = self._make_toggle_button(
            "Grid", checked=False, tooltip="Toggle grid on profile axis"
        )
        self.grid_cb.toggled.connect(self._on_theme_toggled)
        primary_row.addWidget(self.grid_cb)

        self.dark_bg_cb = self._make_toggle_button(
            "Dark", checked=self._dark_background, tooltip="Toggle dark plotting background"
        )
        self.dark_bg_cb.toggled.connect(self._on_theme_toggled)
        primary_row.addWidget(self.dark_bg_cb)

        primary_row.addStretch(1)
        self.advanced_toggle_btn = self._make_toggle_button(
            "Advanced \u25bc", checked=False, tooltip="Show/hide advanced profile controls"
        )
        self.advanced_toggle_btn.toggled.connect(self._set_advanced_options_visible)
        primary_row.addWidget(self.advanced_toggle_btn)
        controls_layout.addLayout(primary_row)

        self._advanced_controls_widget = QtWidgets.QWidget()
        advanced_row = QtWidgets.QHBoxLayout(self._advanced_controls_widget)
        advanced_row.setContentsMargins(0, 0, 0, 0)
        advanced_row.setSpacing(6)

        self.extra_ticks_cb = self._make_toggle_button(
            "Extra ticks", checked=False, tooltip="Enable additional minor tick marks"
        )
        self.extra_ticks_cb.toggled.connect(self._on_plot_option_changed)
        advanced_row.addWidget(self.extra_ticks_cb)

        self.precision_cb = self._make_toggle_button(
            "Precision", checked=False, tooltip="Higher tick density for fine inspection"
        )
        self.precision_cb.toggled.connect(self._on_plot_option_changed)
        advanced_row.addWidget(self.precision_cb)

        self.multi_channel_cb = self._make_toggle_button(
            "Multi-channel", checked=False, tooltip="Plot extra channel profiles when available"
        )
        self.multi_channel_cb.toggled.connect(self._on_plot_option_changed)
        advanced_row.addWidget(self.multi_channel_cb)

        # Preview control disabled because the dialog preview is commented out to save resources.
        # self.preview_toggle_cb = QtWidgets.QCheckBox("Show preview")
        # self.preview_toggle_cb.setChecked(True)
        # self.preview_toggle_cb.toggled.connect(self._on_preview_toggle)
        # plot_layout.addWidget(self.preview_toggle_cb)
        # (If you re-enable the preview panel above, uncomment these lines to restore the toggle.)
        self.preserve_profiles_cb = self._make_toggle_button(
            "Preserve profiles", checked=True, tooltip="Keep overlays when changing channel"
        )
        self.preserve_profiles_cb.toggled.connect(self._on_preserve_toggle)
        advanced_row.addWidget(self.preserve_profiles_cb)
        advanced_row.addStretch(1)
        controls_layout.addWidget(self._advanced_controls_widget)
        self._set_advanced_options_visible(False)

        info_layout.addWidget(controls_panel)
        self.profile_list = QtWidgets.QListWidget()
        self.profile_list.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.profile_list.setAlternatingRowColors(True)
        self.profile_list.setUniformItemSizes(True)
        self.profile_list.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.profile_list.itemDoubleClicked.connect(self._on_profile_item_activated)
        self.profile_list.currentItemChanged.connect(self._on_profile_item_selected)
        self.profile_list.customContextMenuRequested.connect(self._on_profile_list_context_menu)
        profiles_header = QtWidgets.QHBoxLayout()
        profiles_header.setContentsMargins(0, 0, 0, 0)
        profiles_header.setSpacing(6)
        profiles_header.addWidget(QtWidgets.QLabel("Profiles"))
        profiles_header.addStretch(1)
        self.compose_drag_btn = _ProfileCompositeDragButton(self)
        profiles_header.addWidget(self.compose_drag_btn)
        info_layout.addLayout(profiles_header)
        info_layout.addWidget(self.profile_list, 1)
        btn_layout = QtWidgets.QHBoxLayout()
        self.copy_btn = QtWidgets.QPushButton('Copy XY')
        self.copy_btn.clicked.connect(self._copy_current_profile)
        btn_layout.addWidget(self.copy_btn)
        self.add_btn = QtWidgets.QPushButton('Add overlay')
        self.add_btn.clicked.connect(self._add_overlay_from_active)
        btn_layout.addWidget(self.add_btn)
        self.delete_btn = QtWidgets.QPushButton('Delete')
        self.delete_btn.clicked.connect(self._delete_selected_profile)
        btn_layout.addWidget(self.delete_btn)
        btn_layout.addStretch(1)
        self.close_btn = QtWidgets.QPushButton('Close')
        self.close_btn.clicked.connect(self.accept)
        btn_layout.addWidget(self.close_btn)
        info_layout.addLayout(btn_layout)
        splitter.addWidget(info_widget)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([300, 140])
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
        self._delete_shortcut = QtWidgets.QShortcut(QtGui.QKeySequence("Delete"), self)
        self._delete_shortcut.activated.connect(self._delete_selected_profile)
        self._delete_shortcut_back = QtWidgets.QShortcut(QtGui.QKeySequence("Backspace"), self)
        self._delete_shortcut_back.activated.connect(self._delete_selected_profile)
        self._delete_list_shortcut = QtWidgets.QShortcut(QtGui.QKeySequence("Delete"), self.profile_list)
        self._delete_list_shortcut.setContext(QtCore.Qt.WidgetShortcut)
        self._delete_list_shortcut.activated.connect(self._delete_selected_profile)
        self._delete_list_shortcut_back = QtWidgets.QShortcut(QtGui.QKeySequence("Backspace"), self.profile_list)
        self._delete_list_shortcut_back.setContext(QtCore.Qt.WidgetShortcut)
        self._delete_list_shortcut_back.activated.connect(self._delete_selected_profile)
        self._context_source = None
        self._context_syncing = False
        self._preserve_cb = None
        self._refresh_action_button_states()
        register_profile_dialog(self)

    def detach_as_workspace_window(self):
        """Make the dialog an independent top-level window so it does not drag the main viewer to front."""
        owner = getattr(self, "_owner", None)
        try:
            self.setParent(None, self.windowFlags())
            self.setWindowFlag(QtCore.Qt.Window, True)
            self.setWindowModality(QtCore.Qt.NonModal)
            if owner is not None and hasattr(owner, "windowIcon"):
                self.setWindowIcon(owner.windowIcon())
        except Exception:
            pass

    def _on_context_menu(self, pos):
        menu = QtWidgets.QMenu(self)
        preset_menu = menu.addMenu("Figure preset")
        preset_actions = {}
        for preset in iter_figure_layout_presets():
            act = preset_menu.addAction(preset.label)
            act.setCheckable(True)
            act.setChecked(self._figure_preset_key == preset.key)
            preset_actions[act] = preset.key
        menu.addSeparator()
        metadata_menu = menu.addMenu("Metadata")
        metadata_show_act = metadata_menu.addAction("Show metadata on plot")
        metadata_show_act.setCheckable(True)
        metadata_show_act.setChecked(bool(self._metadata_visible))
        metadata_file_act = metadata_menu.addAction("File name")
        metadata_file_act.setCheckable(True)
        metadata_file_act.setChecked(bool(self._metadata_show_filename))
        metadata_acq_act = metadata_menu.addAction("Acquisition title")
        metadata_acq_act.setCheckable(True)
        metadata_acq_act.setChecked(bool(self._metadata_show_acquisition))
        metadata_time_act = metadata_menu.addAction("Acquisition time")
        metadata_time_act.setCheckable(True)
        metadata_time_act.setChecked(bool(self._metadata_show_time))
        metadata_folder_name_act = metadata_menu.addAction("Folder name")
        metadata_folder_name_act.setCheckable(True)
        metadata_folder_name_act.setChecked(bool(self._metadata_show_folder_name))
        metadata_folder_act = metadata_menu.addAction("Folder")
        metadata_folder_act.setCheckable(True)
        metadata_folder_act.setChecked(bool(self._metadata_show_folder))
        menu.addSeparator()
        copy_png = menu.addAction("Copy plot (PNG 300 dpi)")
        copy_png_600 = menu.addAction("Copy plot (PNG 600 dpi)")
        copy_svg = menu.addAction("Copy plot (SVG)")
        save_menu = menu.addMenu("Save plot")
        save_png_300 = save_menu.addAction("PNG 300 dpi...")
        save_png_600 = save_menu.addAction("PNG 600 dpi...")
        save_svg = save_menu.addAction("SVG (vector)...")
        save_pdf = save_menu.addAction("PDF (vector)...")
        menu.addSeparator()
        legend = self.ax.get_legend()
        legend_menu = menu.addMenu("Legend")
        show_legend_act = legend_menu.addAction("Show legend")
        show_legend_act.setCheckable(True)
        show_legend_act.setChecked(bool(self._legend_visible))
        font_menu = legend_menu.addMenu("Font size")
        font_actions = {}
        for size in (7.0, 8.0, 9.0, 10.0, 12.0, 14.0):
            act = font_menu.addAction(f"{size:.0f} pt")
            act.setCheckable(True)
            act.setChecked(abs(float(self._legend_fontsize) - size) < 1e-6)
            font_actions[act] = size
        pos_menu = legend_menu.addMenu("Position")
        pos_actions = {}
        for label, value in (
            ("Upper right", "upper right"),
            ("Upper left", "upper left"),
            ("Lower right", "lower right"),
            ("Lower left", "lower left"),
            ("Upper center", "upper center"),
            ("Lower center", "lower center"),
            ("Center right", "center right"),
            ("Center left", "center left"),
            ("Center", "center"),
        ):
            act = pos_menu.addAction(label)
            act.setCheckable(True)
            act.setChecked(self._legend_custom_anchor is None and self._legend_loc == value)
            pos_actions[act] = value
        legend_menu.addSeparator()
        frame_fill_act = legend_menu.addAction("Fill background")
        frame_fill_act.setCheckable(True)
        frame_fill_act.setChecked(bool(self._legend_frame_fill_visible))
        outline_act = legend_menu.addAction("Show outline")
        outline_act.setCheckable(True)
        outline_act.setChecked(bool(self._legend_outline_visible))
        outline_width_menu = legend_menu.addMenu("Outline thickness")
        outline_width_actions = {}
        for width in (0.5, 1.0, 1.5, 2.0, 3.0):
            act = outline_width_menu.addAction(f"{width:.1f} pt")
            act.setCheckable(True)
            act.setChecked(abs(float(self._legend_outline_width) - width) < 1e-6)
            outline_width_actions[act] = width
        reset_drag_act = legend_menu.addAction("Reset dragged position")
        reset_drag_act.setEnabled(self._legend_custom_anchor is not None)
        if legend is not None:
            legend_menu.addSeparator()
            drag_hint_act = legend_menu.addAction("Drag legend with left mouse")
            drag_hint_act.setEnabled(False)
        target_idx = self._selected_overlay_index()
        target_active = target_idx is None
        style_target = "Active profile" if target_active else f"Overlay {int(target_idx) + 1}"
        style_menu = menu.addMenu(f"Style {style_target}")
        pick_color_act = style_menu.addAction("Pick color...")
        width_menu = style_menu.addMenu("Line thickness")
        width_actions = {}
        current_lw = self._profile_value(target_idx, "lw", 1.5 if not target_active else 2.0)
        for width in (0.5, 1.0, 1.5, 2.0, 3.0, 4.0, 6.0):
            act = width_menu.addAction(f"{width:.1f} pt")
            act.setCheckable(True)
            act.setChecked(abs(float(current_lw) - width) < 1e-6)
            width_actions[act] = width
        line_menu = style_menu.addMenu("Line style")
        line_actions = {}
        current_line_style = self._profile_value(target_idx, "line_style", "-" if target_active else "--")
        for label, value in (("Solid", "-"), ("Dashed", "--"), ("Dotted", ":"), ("Dash-dot", "-."), ("None", "None")):
            act = line_menu.addAction(label)
            act.setCheckable(True)
            act.setChecked(str(current_line_style) == value)
            line_actions[act] = value
        marker_menu = style_menu.addMenu("Marker shape")
        marker_actions = {}
        current_marker = self._profile_value(target_idx, "marker_style", "o")
        for label, value in (("Circle", "o"), ("Square", "s"), ("Triangle", "^"), ("Diamond", "D"), ("Plus", "+"), ("Cross", "x"), ("Star", "*"), ("None", "None")):
            act = marker_menu.addAction(label)
            act.setCheckable(True)
            act.setChecked(str(current_marker) == value)
            marker_actions[act] = value
        marker_size_menu = style_menu.addMenu("Marker size")
        marker_size_actions = {}
        current_marker_size = self._profile_value(target_idx, "marker_size", 7.0 if target_active else 5.0)
        for size in (3.0, 5.0, 7.0, 9.0, 12.0):
            act = marker_size_menu.addAction(f"{size:.0f} pt")
            act.setCheckable(True)
            act.setChecked(abs(float(current_marker_size) - size) < 1e-6)
            marker_size_actions[act] = size
        palette_menu = menu.addMenu("Apply Color Palette")
        palette_actions = {}
        for name in list_color_cycles():
            act = palette_menu.addAction(name)
            act.setCheckable(True)
            act.setChecked(name == self._profile_palette_name)
            palette_actions[act] = name
        add_font_menu_action(
            menu,
            self,
            self._plot_font_family,
            self.set_plot_font_family,
            current_style=self._font_style_state(),
            apply_style_callback=self.set_plot_typography,
        )
        action = menu.exec_(self.canvas.mapToGlobal(pos))
        if action in preset_actions:
            self._apply_figure_preset(preset_actions[action])
        elif action == metadata_show_act:
            self._metadata_visible = bool(metadata_show_act.isChecked())
            self.update_profiles(
                self._active,
                self._saved,
                activate_overlay_callback=self._activate_overlay_cb,
                highlight_overlay_callback=self._highlight_overlay_cb,
            )
        elif action == metadata_file_act:
            self._metadata_show_filename = bool(metadata_file_act.isChecked())
            self.update_profiles(
                self._active,
                self._saved,
                activate_overlay_callback=self._activate_overlay_cb,
                highlight_overlay_callback=self._highlight_overlay_cb,
            )
        elif action == metadata_acq_act:
            self._metadata_show_acquisition = bool(metadata_acq_act.isChecked())
            self.update_profiles(
                self._active,
                self._saved,
                activate_overlay_callback=self._activate_overlay_cb,
                highlight_overlay_callback=self._highlight_overlay_cb,
            )
        elif action == metadata_time_act:
            self._metadata_show_time = bool(metadata_time_act.isChecked())
            self.update_profiles(
                self._active,
                self._saved,
                activate_overlay_callback=self._activate_overlay_cb,
                highlight_overlay_callback=self._highlight_overlay_cb,
            )
        elif action == metadata_folder_name_act:
            self._metadata_show_folder_name = bool(metadata_folder_name_act.isChecked())
            self.update_profiles(
                self._active,
                self._saved,
                activate_overlay_callback=self._activate_overlay_cb,
                highlight_overlay_callback=self._highlight_overlay_cb,
            )
        elif action == metadata_folder_act:
            self._metadata_show_folder = bool(metadata_folder_act.isChecked())
            self.update_profiles(
                self._active,
                self._saved,
                activate_overlay_callback=self._activate_overlay_cb,
                highlight_overlay_callback=self._highlight_overlay_cb,
            )
        elif action == copy_png:
            self._copy_plot("png", dpi=300)
        elif action == copy_png_600:
            self._copy_plot("png", dpi=600)
        elif action == copy_svg:
            self._copy_plot("svg")
        elif action == save_png_300:
            self._save_plot("png", dpi=300)
        elif action == save_png_600:
            self._save_plot("png", dpi=600)
        elif action == save_svg:
            self._save_plot("svg")
        elif action == save_pdf:
            self._save_plot("pdf")
        elif action == show_legend_act:
            self._legend_visible = bool(show_legend_act.isChecked())
            self.update_profiles(
                self._active,
                self._saved,
                activate_overlay_callback=self._activate_overlay_cb,
                highlight_overlay_callback=self._highlight_overlay_cb,
            )
        elif action in font_actions:
            self._legend_fontsize = float(font_actions[action])
            self.update_profiles(
                self._active,
                self._saved,
                activate_overlay_callback=self._activate_overlay_cb,
                highlight_overlay_callback=self._highlight_overlay_cb,
            )
        elif action in pos_actions:
            self._legend_loc = pos_actions[action]
            self._legend_custom_anchor = None
            self.update_profiles(
                self._active,
                self._saved,
                activate_overlay_callback=self._activate_overlay_cb,
                highlight_overlay_callback=self._highlight_overlay_cb,
            )
        elif action == frame_fill_act:
            self._legend_frame_fill_visible = bool(frame_fill_act.isChecked())
            self._apply_plot_theme()
        elif action == outline_act:
            self._legend_outline_visible = bool(outline_act.isChecked())
            self._apply_plot_theme()
        elif action in outline_width_actions:
            self._legend_outline_width = float(outline_width_actions[action])
            self._apply_plot_theme()
        elif action == reset_drag_act:
            self._legend_custom_anchor = None
            self.update_profiles(
                self._active,
                self._saved,
                activate_overlay_callback=self._activate_overlay_cb,
                highlight_overlay_callback=self._highlight_overlay_cb,
            )
        elif action == pick_color_act:
            current = QtGui.QColor(self._profile_value(target_idx, "color", "#fbc02d"))
            picked = QtWidgets.QColorDialog.getColor(current, self, f"Select color for {style_target}")
            if picked.isValid():
                self._apply_profile_style_change(target_idx, color=picked.name())
        elif action in width_actions:
            self._apply_profile_style_change(target_idx, lw=width_actions[action])
        elif action in line_actions:
            self._apply_profile_style_change(target_idx, line_style=line_actions[action])
        elif action in marker_actions:
            self._apply_profile_style_change(target_idx, marker_style=marker_actions[action])
        elif action in marker_size_actions:
            self._apply_profile_style_change(target_idx, marker_size=marker_size_actions[action])
        elif action in palette_actions:
            self._apply_palette_change(palette_actions[action])

    def _on_profile_list_context_menu(self, pos):
        self._on_context_menu(self.canvas.mapFromGlobal(self.profile_list.viewport().mapToGlobal(pos)))

    def _profile_value(self, profile_key, field, default=None):
        if profile_key is None:
            dataset = self._active or {}
        elif 0 <= int(profile_key) < len(self._saved):
            dataset = self._saved[int(profile_key)] or {}
        else:
            dataset = {}
        return dataset.get(field, default)

    def _dataset_for_profile_key(self, profile_key):
        if profile_key is None:
            return self._active
        try:
            idx = int(profile_key)
        except Exception:
            return None
        if 0 <= idx < len(self._saved):
            return self._saved[idx]
        return None

    @staticmethod
    def _profile_id_sort_value(dataset):
        if not isinstance(dataset, dict):
            return None
        profile_id = str(
            dataset.get("profile_id")
            or ((dataset.get("live_profile_ref") or {}).get("profile_id"))
            or ""
        ).strip()
        if not profile_id:
            return None
        digits = []
        for ch in reversed(profile_id):
            if ch.isdigit():
                digits.append(ch)
            elif digits:
                break
        if digits:
            try:
                return int("".join(reversed(digits)))
            except Exception:
                return profile_id
        return profile_id

    def _ordered_profile_entries(self, active_profile, saved_profiles):
        entries = []
        if active_profile:
            entries.append({
                "key": None,
                "label": "Active",
                "data": active_profile,
                "is_active": True,
                "sort_value": self._profile_id_sort_value(active_profile),
                "fallback": -1,
            })
        for idx, data in enumerate(saved_profiles or []):
            entries.append({
                "key": idx,
                "label": f"Overlay {idx + 1}",
                "data": data,
                "is_active": False,
                "sort_value": self._profile_id_sort_value(data),
                "fallback": idx,
            })
        def _sort_key(entry):
            sort_value = entry.get("sort_value")
            if sort_value is None:
                return (0 if entry.get("is_active") else 1, entry.get("fallback", 0))
            return (1, sort_value)
        return sorted(entries, key=_sort_key)

    def _live_profile_ref(self, profile_key):
        dataset = self._dataset_for_profile_key(profile_key)
        ref = dataset.get("live_profile_ref") if isinstance(dataset, dict) else None
        return ref if profile_ref_key(ref) is not None else None

    def _apply_profile_style_change(self, profile_key, **changes):
        updated = False
        live_ref = self._live_profile_ref(profile_key)
        if callable(self._style_update_cb):
            try:
                updated = bool(self._style_update_cb(profile_key, **changes))
            except Exception:
                updated = False
        if not updated and live_ref is not None:
            updated = bool(apply_live_profile_style(live_ref, **changes))
        target = self._dataset_for_profile_key(profile_key)
        if target is not None and (updated or live_ref is None):
            target.update(changes)
            updated = True
        if updated:
            current = profile_key
            self.update_profiles(
                self._active,
                self._saved,
                activate_overlay_callback=self._activate_overlay_cb,
                highlight_overlay_callback=self._highlight_overlay_cb,
            )
            self.select_overlay(current)

    def _apply_palette_change(self, palette_name):
        colors = get_color_cycle(palette_name)
        if not colors:
            return
        if callable(self._palette_cb) and not self._composite_mode:
            try:
                self._palette_cb(palette_name)
            except Exception:
                pass
            self._profile_palette_name = palette_name
            color_iter = iter(colors)
            if self._active is not None:
                self._active["color"] = next(color_iter, colors[0])
            for idx, entry in enumerate(self._saved):
                entry["color"] = colors[(idx + 1) % len(colors)]
            self.update_profiles(
                self._active,
                self._saved,
                activate_overlay_callback=self._activate_overlay_cb,
                highlight_overlay_callback=self._highlight_overlay_cb,
            )
            return
        updated = False
        if self._active is not None:
            live_ref = self._live_profile_ref(None)
            color = colors[0]
            if live_ref is not None:
                updated = bool(apply_live_profile_style(live_ref, color=color)) or updated
            self._active["color"] = color
            updated = True
        for idx, entry in enumerate(self._saved):
            color = colors[(idx + 1) % len(colors)]
            live_ref = self._live_profile_ref(idx)
            if live_ref is not None:
                updated = bool(apply_live_profile_style(live_ref, color=color)) or updated
            entry["color"] = color
            updated = True
        if updated:
            self._profile_palette_name = palette_name
            self.update_profiles(
                self._active,
                self._saved,
                activate_overlay_callback=self._activate_overlay_cb,
                highlight_overlay_callback=self._highlight_overlay_cb,
            )

    def _font_style_state(self):
        return {
            "bold": bool(getattr(self, "_plot_font_bold", False)),
            "italic": bool(getattr(self, "_plot_font_italic", False)),
            "underline": bool(getattr(self, "_plot_font_underline", False)),
        }

    def set_plot_typography(self, **changes):
        """Update shared plot typography and redraw with the new style."""
        family = changes.get("family", None)
        owner = getattr(self, "_owner", None)
        style_changes = {
            "bold": changes.get("bold", None),
            "italic": changes.get("italic", None),
            "underline": changes.get("underline", None),
        }
        if family is not None:
            family = normalize_font_family(family, "sans-serif")
            self._plot_font_family = family
        if owner is not None and hasattr(owner, "set_plot_typography"):
            target = {
                "family": family if family is not None else self._plot_font_family,
                "bold": bool(style_changes["bold"] if style_changes["bold"] is not None else self._plot_font_bold),
                "italic": bool(style_changes["italic"] if style_changes["italic"] is not None else self._plot_font_italic),
                "underline": bool(style_changes["underline"] if style_changes["underline"] is not None else self._plot_font_underline),
            }
            if any(getattr(owner, f"_plot_font_{k}", None) != v for k, v in target.items()):
                try:
                    owner.set_plot_typography(**target)
                    return
                except Exception:
                    pass
        for key, attr in (("bold", "_plot_font_bold"), ("italic", "_plot_font_italic"), ("underline", "_plot_font_underline")):
            if style_changes[key] is not None:
                setattr(self, attr, bool(style_changes[key]))
        self.update_profiles(self._active, self._saved)

    def set_plot_font_family(self, family: str):
        """Rebuild the profile plot with a new shared font family."""
        self.set_plot_typography(family=family)

    def _apply_figure_preset(self, preset_key):
        """Apply a publication/slide sizing preset to the profile plot."""
        preset = get_figure_layout_preset(preset_key)
        self._figure_preset_key = preset.key
        apply_figure_layout(self.canvas.figure, preset)
        plot_w_px, plot_h_px = preset_pixel_size(self, preset, max_fraction=0.62)
        apply_canvas_widget_preset(self.canvas, preset, plot_w_px, plot_h_px)
        self._plot_font_family = normalize_font_family(preset.font_family, "sans-serif")
        self._plot_font_bold = False
        self._plot_font_italic = False
        self._plot_font_underline = False
        self._font_scale = float(preset.font_scale)
        self._legend_fontsize = float(preset.legend_font_pt)
        if callable(self._label_scale_cb):
            try:
                self._label_scale_cb(self._font_scale)
            except Exception:
                pass
        self.update_profiles(
            self._active,
            self._saved,
            activate_overlay_callback=self._activate_overlay_cb,
            highlight_overlay_callback=self._highlight_overlay_cb,
        )
        try:
            total_w = max(720, int(plot_w_px + 140))
            total_h = max(520, int(plot_h_px + 260))
            self.resize(total_w, total_h)
            if hasattr(self, "_splitter") and self._splitter is not None:
                self._splitter.setSizes([plot_h_px + 40, 220])
        except Exception:
            pass
        QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), f"Applied preset: {preset.label}", self)

    def _metadata_text(self, dataset):
        """Build the optional source-metadata block shown above the profile plot."""
        if not self._metadata_visible or not dataset:
            return ""
        lines = []
        if self._metadata_show_filename:
            name = str(dataset.get("source_file_name") or "").strip()
            if not name:
                path_text = str(dataset.get("source_path") or "").strip()
                if path_text:
                    try:
                        name = Path(path_text).name
                    except Exception:
                        name = path_text
            if name:
                lines.append(f"File: {name}")
        if self._metadata_show_acquisition:
            acq = str(dataset.get("source_acquisition_text") or dataset.get("source_title") or "").strip()
            if acq:
                lines.append(f"Acq: {acq}")
        if self._metadata_show_time:
            when = str(dataset.get("source_datetime") or "").strip()
            if not when:
                source_date = str(dataset.get("source_date") or "").strip()
                source_time = str(dataset.get("source_time") or "").strip()
                if source_date and source_time:
                    when = f"{source_date} {source_time}"
                else:
                    when = source_date or source_time
            if when:
                lines.append(f"Time: {when}")
        if self._metadata_show_folder_name:
            folder_name = str(dataset.get("source_folder_name") or "").strip()
            if not folder_name:
                path_text = str(dataset.get("source_path") or "").strip()
                if path_text:
                    try:
                        folder_name = Path(path_text).parent.name
                    except Exception:
                        folder_name = ""
            if folder_name:
                lines.append(f"Folder name: {folder_name}")
        if self._metadata_show_folder:
            folder = str(dataset.get("source_folder") or "").strip()
            if not folder:
                path_text = str(dataset.get("source_path") or "").strip()
                if path_text:
                    try:
                        folder = str(Path(path_text).parent)
                    except Exception:
                        folder = ""
            if folder:
                lines.append(f"Folder: {folder}")
        return "\n".join(lines)

    def _apply_metadata_overlay(self, dataset):
        """Render the optional metadata block onto the profile figure."""
        if self._metadata_artist is not None:
            try:
                self._metadata_artist.remove()
            except Exception:
                pass
            self._metadata_artist = None
        text = self._metadata_text(dataset)
        if not text:
            return
        dark = bool(self._dark_background)
        box_face = "#111111" if dark else "#ffffff"
        text_color = "#f5f5f5" if dark else "#111111"
        try:
            self._metadata_artist = self.canvas.figure.text(
                0.02,
                0.985,
                text,
                ha="left",
                va="top",
                fontsize=max(5.0, 6.0 * getattr(self, "_font_scale", 1.0)),
                color=text_color,
                bbox={"facecolor": box_face, "alpha": 0.72, "edgecolor": "none", "pad": 2.0},
            )
            apply_text_style(self._metadata_artist, family=self._plot_font_family, **self._font_style_state())
            self.canvas.figure.subplots_adjust(top=0.84 if "\n" in text else 0.88)
        except Exception:
            self._metadata_artist = None

    def _copy_plot(self, fmt, *, dpi=300):
        copy_figure_to_clipboard(self, self.canvas.figure, fmt, dpi=dpi)

    def _save_plot(self, fmt, *, dpi=300):
        save_figure_with_dialog(self, self.canvas.figure, default_stem="profile_measurement", fmt=fmt, dpi=dpi)

    def _make_toggle_button(self, text, *, checked=False, tooltip=None):
        btn = QtWidgets.QToolButton(self)
        btn.setObjectName("profileToggleButton")
        btn.setText(text)
        btn.setCheckable(True)
        btn.setChecked(bool(checked))
        btn.setAutoRaise(False)
        btn.setCursor(QtCore.Qt.PointingHandCursor)
        btn.setToolButtonStyle(QtCore.Qt.ToolButtonTextOnly)
        btn.setSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        if tooltip:
            btn.setToolTip(tooltip)
        self._toggle_buttons.append(btn)
        return btn

    def _set_advanced_options_visible(self, visible):
        visible = bool(visible)
        self._advanced_controls_visible = visible
        if hasattr(self, "_advanced_controls_widget") and self._advanced_controls_widget is not None:
            self._advanced_controls_widget.setVisible(visible)
        if hasattr(self, "advanced_toggle_btn") and self.advanced_toggle_btn is not None:
            self.advanced_toggle_btn.blockSignals(True)
            self.advanced_toggle_btn.setChecked(visible)
            self.advanced_toggle_btn.setText("Advanced \u25b2" if visible else "Advanced \u25bc")
            self.advanced_toggle_btn.blockSignals(False)

    def _apply_toggle_button_styles(self):
        dark = bool(self._dark_background)
        if dark:
            inactive_bg = "#1e2430"
            inactive_border = "#46556e"
            inactive_text = "#d4deee"
            active_bg = "#2f6fcb"
            active_border = "#79a9f2"
            active_text = "#ffffff"
            hint_color = "#b9c6d8"
            compose_button_bg = "#202633"
            compose_button_border = "#5a6880"
            compose_button_text = "#e4ebf7"
            compose_button_hover_border = "#82aef1"
            compose_drop_bg = "#26477a"
            compose_drop_border = "#a8ceff"
        else:
            inactive_bg = "#f3f5f9"
            inactive_border = "#aeb7c5"
            inactive_text = "#1f2a3d"
            active_bg = "#1f6fd7"
            active_border = "#5b97e8"
            active_text = "#ffffff"
            hint_color = "#4b5b73"
            compose_button_bg = "#f5f7fa"
            compose_button_border = "#b8c2cf"
            compose_button_text = "#314056"
            compose_button_hover_border = "#5b97e8"
            compose_drop_bg = "#e6f0ff"
            compose_drop_border = "#4f8fe3"
        style = (
            "QToolButton#profileToggleButton {"
            f"background-color: {inactive_bg};"
            f"color: {inactive_text};"
            f"border: 1px solid {inactive_border};"
            "border-radius: 12px;"
            "padding: 4px 12px;"
            "font-weight: 600;"
            "}"
            "QToolButton#profileToggleButton:checked {"
            f"background-color: {active_bg};"
            f"color: {active_text};"
            f"border: 1px solid {active_border};"
            "}"
            "QToolButton#profileToggleButton:hover {"
            f"border: 1px solid {active_border};"
            "}"
        )
        for btn in self._toggle_buttons:
            try:
                btn.setStyleSheet(style)
            except Exception:
                pass
        hint = self.findChild(QtWidgets.QLabel, "profileControlsHint")
        if hint is not None:
            hint.setStyleSheet(f"color: {hint_color};")
        compose_btn = getattr(self, "compose_drag_btn", None)
        if compose_btn is not None:
            drop_active = bool(compose_btn.property("dropActive"))
            button_bg = compose_drop_bg if drop_active else compose_button_bg
            button_border = compose_drop_border if drop_active else compose_button_border
            button_style = (
                "QToolButton#profileComposeButton {"
                f"background-color: {button_bg};"
                f"color: {compose_button_text};"
                f"border: 1px solid {button_border};"
                "border-radius: 10px;"
                "padding: 4px 10px;"
                "font-weight: 600;"
                "}"
                "QToolButton#profileComposeButton:hover {"
                f"border: 1px solid {compose_button_hover_border};"
                "}"
                "QToolButton#profileComposeButton:disabled {"
                f"background-color: {inactive_bg};"
                f"color: {hint_color};"
                f"border: 1px solid {inactive_border};"
                "}"
            )
            try:
                compose_btn.setStyleSheet(button_style)
            except Exception:
                pass

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

    def keyPressEvent(self, event):
        key = event.key()
        if key in (QtCore.Qt.Key_Delete, QtCore.Qt.Key_Backspace):
            self._delete_selected_profile()
            event.accept()
            return
        try:
            mods = event.modifiers()
        except Exception:
            mods = QtCore.Qt.NoModifier
        if mods == QtCore.Qt.NoModifier:
            if key == QtCore.Qt.Key_V and hasattr(self, "marker_toggle"):
                self.marker_toggle.toggle()
                event.accept()
                return
            if key == QtCore.Qt.Key_G and hasattr(self, "grid_cb"):
                self.grid_cb.toggle()
                event.accept()
                return
            if key == QtCore.Qt.Key_L and hasattr(self, "show_lines_cb"):
                self.show_lines_cb.toggle()
                event.accept()
                return
            if key == QtCore.Qt.Key_P and hasattr(self, "show_points_cb"):
                self.show_points_cb.toggle()
                event.accept()
                return
            if key == QtCore.Qt.Key_M and hasattr(self, "multi_channel_cb"):
                self.multi_channel_cb.toggle()
                event.accept()
                return
            if key == QtCore.Qt.Key_T and hasattr(self, "extra_ticks_cb"):
                self.extra_ticks_cb.toggle()
                event.accept()
                return
            if key == QtCore.Qt.Key_R and hasattr(self, "precision_cb"):
                self.precision_cb.toggle()
                event.accept()
                return
            if key == QtCore.Qt.Key_A and hasattr(self, "advanced_toggle_btn"):
                self.advanced_toggle_btn.toggle()
                event.accept()
                return
        super().keyPressEvent(event)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self.context_canvas is not None:
            self.context_canvas.draw_idle()

    def _apply_font_scale(self):
        scale = max(0.6, min(1.8, getattr(self, '_font_scale', 1.0)))
        label_size = 10 * scale
        tick_size = 9 * scale
        try:
            self.ax.tick_params(axis='both', labelsize=tick_size)
            self.ax_top.tick_params(axis='both', labelsize=tick_size)
            self.ax_right.tick_params(axis='both', labelsize=tick_size)
            self.ax.xaxis.label.set_fontsize(label_size)
            self.ax.yaxis.label.set_fontsize(label_size)
            self.ax_top.xaxis.label.set_fontsize(label_size)
            self.ax_right.yaxis.label.set_fontsize(label_size)
            style = self._font_style_state()
            for text in (self.ax.xaxis.label, self.ax.yaxis.label, self.ax_top.xaxis.label, self.ax_right.yaxis.label):
                apply_text_style(text, family=self._plot_font_family, **style)
            for text in list(self.ax.get_xticklabels()) + list(self.ax.get_yticklabels()) + list(self.ax_top.get_xticklabels()) + list(self.ax_right.get_yticklabels()):
                apply_text_style(text, family=self._plot_font_family, **style)
            title = self.ax.get_title()
            if title:
                apply_text_style(self.ax.title, family=self._plot_font_family, **style)
            self._apply_legend_style()
        except Exception:
            pass
        for widget in (self.stats, self.marker_info):
            if widget is not None:
                font = widget.font()
                font.setPointSizeF(max(7.0, 9.0 * scale))
                font = apply_qfont_style(
                    font,
                    family=self._plot_font_family,
                    bold=self._plot_font_bold,
                    italic=self._plot_font_italic,
                    underline=self._plot_font_underline,
                )
                widget.setFont(font)
        if self.profile_list is not None:
            font = self.profile_list.font()
            font.setPointSizeF(max(7.0, 9.0 * scale))
            font = apply_qfont_style(
                font,
                family=self._plot_font_family,
                bold=self._plot_font_bold,
                italic=self._plot_font_italic,
                underline=self._plot_font_underline,
            )
            self.profile_list.setFont(font)
        for btn in getattr(self, "_toggle_buttons", []):
            try:
                font = btn.font()
                font.setPointSizeF(max(7.0, 8.8 * scale))
                font = apply_qfont_style(
                    font,
                    family=self._plot_font_family,
                    bold=self._plot_font_bold,
                    italic=self._plot_font_italic,
                    underline=self._plot_font_underline,
                )
                btn.setFont(font)
            except Exception:
                pass
        for btn in (getattr(self, "copy_btn", None), getattr(self, "add_btn", None), getattr(self, "delete_btn", None), getattr(self, "close_btn", None)):
            if btn is None:
                continue
            try:
                font = btn.font()
                font.setPointSizeF(max(7.0, 9.0 * scale))
                font = apply_qfont_style(
                    font,
                    family=self._plot_font_family,
                    bold=self._plot_font_bold,
                    italic=self._plot_font_italic,
                    underline=self._plot_font_underline,
                )
                btn.setFont(font)
            except Exception:
                pass
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
        return fmt_length(title, length_nm)

    def _format_stats_text(self, active, saved):
        return format_stats_text(active, saved)

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
        self._notify_marker_positions()
        self.canvas.draw_idle()

    def _ensure_marker_lines(self):
        if not self._markers_enabled:
            return False
        if not self._marker_positions:
            return False
        if len(self._marker_lines) == len(self._marker_positions):
            return True
        for line in self._marker_lines:
            try:
                if line:
                    line.remove()
            except Exception:
                pass
        self._marker_lines = []
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
        self.canvas.draw_idle()
        return len(self._marker_lines) == len(self._marker_positions)

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
        x_phys = (reference_dataset or {}).get('x_nm')
        x_px = (reference_dataset or {}).get('x_px')
        has_phys_axis = bool(reference_dataset and x_phys is not None)
        self._marker_axis_unit = 'phys' if has_phys_axis else 'px'
        if has_phys_axis:
            self._marker_display_unit = axis_unit or 'nm'
            self._marker_axis_scale = None
            try:
                if x_px is not None and ref_points is not None:
                    ref_arr = np.asarray(ref_points, dtype=float)
                    px_arr = np.asarray(x_px, dtype=float)
                    if ref_arr.size == px_arr.size and np.allclose(ref_arr, px_arr, rtol=0.0, atol=1e-6):
                        span_px = float(px_arr[-1] - px_arr[0])
                        span_phys = float(np.asarray(x_phys, dtype=float)[-1] - np.asarray(x_phys, dtype=float)[0])
                        if span_px != 0.0 and np.isfinite(span_px) and np.isfinite(span_phys):
                            self._marker_axis_scale = span_phys / span_px
            except Exception:
                self._marker_axis_scale = None
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
        self._notify_marker_positions()
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
        self._notify_marker_positions()
    def _update_marker_info(self):
        if not self._markers_enabled:
            self.marker_info.setText("Markers hidden")
            self._notify_marker_positions()
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
            self._notify_marker_positions()
            return
        axis_delta = abs(self._marker_positions[1] - self._marker_positions[0])
        disp_value, disp_unit = self._format_marker_delta(axis_delta)
        info = f"Markers +: {disp_value:.3f} {disp_unit}"
        if self._marker_axis_scale is not None:
            info += f" ({axis_delta:.1f} px)"
        if self._marker_reference and self._marker_reference.get('y') is not None:
            v0 = self._marker_value_at(self._marker_positions[0])
            v1 = self._marker_value_at(self._marker_positions[1])
            if v0 is not None and v1 is not None:
                info += f" | values: {v0:.3g} G {v1:.3g} (+={abs(v1-v0):.3g})"
        self.marker_info.setText(info)
        self._remember_marker_positions()
        self._update_marker_annotation(axis_delta)
        self._notify_marker_positions()

    def _remember_marker_positions(self):
        if self._marker_positions:
            self._marker_saved_positions = list(self._marker_positions)

    def _format_marker_delta(self, axis_delta):
        return format_marker_delta(axis_delta, self._marker_axis_scale, self._marker_display_unit)

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
        arrow_color = "#f5f5f5" if self._dark_background else "#111111"
        display_value, display_unit = self._format_marker_delta(axis_delta)
        text = f"{display_value:.3f} {display_unit}"
        label_size = 9.0 * getattr(self, '_font_scale', 1.0)
        bbox_face = "#050506" if self._dark_background else "white"
        bbox_alpha = 0.7 if not self._dark_background else 0.6

        if self._marker_arrow is not None:
            try:
                self._marker_arrow.xy = (xmax, y_level)
                self._marker_arrow.set_position((xmin, y_level))
                if hasattr(self._marker_arrow, 'arrow_patch'):
                    self._marker_arrow.arrow_patch.set_edgecolor(arrow_color)
                    self._marker_arrow.arrow_patch.set_facecolor(arrow_color)
            except Exception:
                try: self._marker_arrow.remove()
                except: pass
                self._marker_arrow = None

        if self._marker_arrow is None:
            self._marker_arrow = self.ax.annotate(
                "",
                xy=(xmax, y_level),
                xytext=(xmin, y_level),
                arrowprops=dict(arrowstyle="<->", color=arrow_color, lw=1.8),
                annotation_clip=False,
            )

        label_x = (xmin + xmax) / 2.0
        label_y = y_level + 0.02 * (y_max - y_min)

        if self._marker_label is not None:
            try:
                self._marker_label.set_text(text)
                self._marker_label.set_position((label_x, label_y))
                self._marker_label.set_color(arrow_color)
                self._marker_label.set_fontsize(label_size)
                self._marker_label.set_fontfamily(self._plot_font_family)
                self._marker_label.set_fontweight("bold" if self._plot_font_bold else "normal")
                self._marker_label.set_fontstyle("italic" if self._plot_font_italic else "normal")
                self._marker_label.set_underline(bool(self._plot_font_underline))
            except Exception:
                try: self._marker_label.remove()
                except: pass
                self._marker_label = None

        if self._marker_label is None:
            self._marker_label = self.ax.text(
                label_x,
                label_y,
                text,
                color=arrow_color,
                ha="center",
                va="bottom",
                fontsize=label_size,
                bbox=dict(boxstyle="round,pad=0.2", facecolor=bbox_face,
                          alpha=bbox_alpha, edgecolor="none"),
            )

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
        legend = self.ax.get_legend()
        if (
            event is not None
            and event.button == 1
            and event.inaxes is self.ax
            and legend is not None
            and legend.get_visible()
        ):
            try:
                renderer = self.canvas.figure.canvas.get_renderer()
                bbox = legend.get_window_extent(renderer=renderer)
                if bbox.contains(event.x, event.y):
                    self._legend_drag = (
                        float(event.x - bbox.x0),
                        float(event.y - bbox.y1),
                    )
                    return
            except Exception:
                pass
        if not self._markers_enabled:
            return
        if event.button != 1:
            return
        if event.inaxes not in (self.ax, self.ax_top):
            return
        if not self._ensure_marker_lines():
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
        if idx >= len(self._marker_lines):
            if not self._ensure_marker_lines():
                return
        self._marker_drag_idx = idx
        new_pos = self._clamp_marker(x)
        self._marker_positions[idx] = new_pos
        line = self._marker_lines[idx]
        line.set_xdata([new_pos, new_pos])
        self._update_marker_info()
        self.canvas.draw_idle()

    def _on_marker_move(self, event):
        if self._legend_drag is not None:
            if event.inaxes is not self.ax or event.x is None or event.y is None:
                return
            try:
                dx, dy = self._legend_drag
                top_left_display = (float(event.x - dx), float(event.y - dy))
                anchor = self.ax.transAxes.inverted().transform(top_left_display)
                self._legend_custom_anchor = (float(anchor[0]), float(anchor[1]))
                self._legend_loc = 'upper left'
                self._apply_legend_style()
                self.canvas.draw_idle()
            except Exception:
                pass
            return
        if not self._markers_enabled:
            return
        if self._marker_drag_idx is None and self._marker_arrow_drag is None:
            return
        if event.inaxes not in (self.ax, self.ax_top):
            return
        if self._marker_drag_idx is not None and not self._ensure_marker_lines():
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
            if self._marker_drag_idx >= len(self._marker_lines):
                if not self._ensure_marker_lines():
                    return
            new_pos = self._clamp_marker(x)
            self._marker_positions[self._marker_drag_idx] = new_pos
            line = self._marker_lines[self._marker_drag_idx]
            line.set_xdata([new_pos, new_pos])
            self._update_marker_info()
            self.canvas.draw_idle()

    def _on_marker_release(self, event):
        self._legend_drag = None
        if not self._markers_enabled:
            return
        self._marker_drag_idx = None
        self._marker_arrow_drag = None
        self._remember_marker_positions()
        self._notify_marker_positions()
        self.canvas.draw_idle()

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
        return axis_label(unit)

    def _on_plot_option_changed(self, _checked=False):
        self.update_profiles(self._active, self._saved)

    def _on_preview_toggle(self, checked):
        # Preview toggling is a no-op because the preview panel has been disabled to conserve resources.
        # This safe no-op prevents any remaining UI hooks from raising exceptions.
        return

    def _on_preserve_toggle(self, checked):
        if callable(self._preserve_cb):
            try:
                self._preserve_cb(bool(checked))
            except Exception:
                pass

    def set_preserve_profiles_callback(self, cb, *, enabled=None):
        self._preserve_cb = cb
        if enabled is not None and hasattr(self, 'preserve_profiles_cb'):
            try:
                self.preserve_profiles_cb.setChecked(bool(enabled))
            except Exception:
                pass

    def set_context_source(self, source_canvas, *, dark=None, grid=None):
        # Context preview syncing is disabled to reduce resource consumption. The original method
        # mirrored views, theme, layout and profile callbacks from the main canvas into the dialog
        # preview which required creating additional Matplotlib canvases and event hooks.
        # That behavior is intentionally turned off. If you want to re-enable the dialog preview
        # later, restore the original implementation (look at the commented block above where the
        # MultiPreviewCanvas creation was removed).
        self._context_source = None
        return

    @staticmethod
    def _json_ready_profile_value(value):
        if isinstance(value, np.ndarray):
            return {"__profile_array__": value.tolist()}
        if isinstance(value, np.generic):
            return value.item()
        if isinstance(value, dict):
            return {str(k): ProfileDialog._json_ready_profile_value(v) for k, v in value.items()}
        if isinstance(value, (list, tuple)):
            return [ProfileDialog._json_ready_profile_value(v) for v in value]
        return value

    @staticmethod
    def _profile_value_from_json(value):
        if isinstance(value, dict):
            if "__profile_array__" in value:
                try:
                    return np.asarray(value.get("__profile_array__"))
                except Exception:
                    return np.asarray([])
            return {k: ProfileDialog._profile_value_from_json(v) for k, v in value.items()}
        if isinstance(value, list):
            return [ProfileDialog._profile_value_from_json(v) for v in value]
        return value

    @staticmethod
    def _profile_dataset_signature(dataset):
        if not isinstance(dataset, dict):
            return ""
        digest = hashlib.sha1()
        meta = dataset.get("meta") or {}
        for key in (
            "source_path",
            "source_file_name",
            "source_title",
            "source_acquisition_text",
            "label",
            "unit",
            "axis_unit",
            "distance_unit",
        ):
            digest.update(str(dataset.get(key) or "").encode("utf-8", errors="ignore"))
            digest.update(b"\0")
        for key in ("channel", "file_name", "datetime", "date", "time"):
            digest.update(str(meta.get(key) or "").encode("utf-8", errors="ignore"))
            digest.update(b"\0")
        for key in ("x_px", "x_nm", "vals"):
            try:
                arr = np.asarray(dataset.get(key) if dataset.get(key) is not None else [], dtype=float)
            except Exception:
                arr = np.asarray([], dtype=float)
            digest.update(str(arr.shape).encode("utf-8", errors="ignore"))
            digest.update(arr.tobytes())
        return digest.hexdigest()

    def _dataset_display_name(self, dataset, fallback_label):
        if not isinstance(dataset, dict):
            return str(fallback_label or "Profile").strip()
        existing = str(dataset.get("display_name") or "").strip()
        if existing:
            return existing
        meta = dict(dataset.get("meta") or {})
        source_name = str(
            dataset.get("source_file_name")
            or meta.get("file_name")
            or dataset.get("source_title")
            or ""
        ).strip()
        channel = str(meta.get("channel") or "").strip()
        profile_label = str(dataset.get("label") or fallback_label or "").strip()
        parts = []
        if source_name:
            parts.append(source_name)
        if channel:
            parts.append(channel)
        if profile_label:
            parts.append(profile_label)
        if not parts:
            parts.append(str(fallback_label or "Profile").strip() or "Profile")
        return " | ".join(parts[:3])

    def _clone_dataset_for_composite(self, dataset, fallback_label):
        if not isinstance(dataset, dict):
            return None
        cloned = copy.deepcopy(dataset)
        cloned["display_name"] = self._dataset_display_name(cloned, fallback_label)
        return cloned

    def _current_profile_entries(self):
        entries = []
        if self._active:
            dataset = self._clone_dataset_for_composite(self._active, "Active")
            if dataset:
                entries.append({"dataset": dataset, "signature": self._profile_dataset_signature(dataset)})
        for idx, data in enumerate(self._saved, 1):
            dataset = self._clone_dataset_for_composite(data, f"Overlay {idx}")
            if dataset:
                entries.append({"dataset": dataset, "signature": self._profile_dataset_signature(dataset)})
        return entries

    def _composite_payload(self):
        entries = self._current_profile_entries()
        if not entries:
            return None
        return {
            "origin_dialog_id": self._composite_origin_id,
            "dialog_title": str(self.windowTitle() or "Profile measurement"),
            "entries": [
                {
                    "signature": entry.get("signature") or "",
                    "dataset": self._json_ready_profile_value(entry.get("dataset") or {}),
                }
                for entry in entries
            ],
        }

    def start_profile_composite_drag(self):
        payload = self._composite_payload()
        if not payload:
            QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), "No profile data to drag", self)
            return
        drag = QtGui.QDrag(self)
        mime = QtCore.QMimeData()
        mime.setData(_PROFILE_COMPOSITE_MIME, json.dumps(payload).encode("utf-8"))
        drag.setMimeData(mime)
        pixmap = QtGui.QPixmap(120, 28)
        pixmap.fill(QtGui.QColor("#1d3557"))
        painter = QtGui.QPainter(pixmap)
        painter.setPen(QtGui.QColor("#f1faee"))
        painter.drawText(pixmap.rect().adjusted(8, 0, -8, 0), QtCore.Qt.AlignVCenter | QtCore.Qt.AlignLeft, "Composite profile")
        painter.end()
        drag.setPixmap(pixmap)
        drag.setHotSpot(QtCore.QPoint(16, 14))
        drag.exec_(QtCore.Qt.CopyAction)

    def _entries_from_profile_payload(self, payload):
        entries = []
        for raw in list((payload or {}).get("entries") or []):
            if not isinstance(raw, dict):
                continue
            dataset = self._profile_value_from_json(raw.get("dataset"))
            if not isinstance(dataset, dict):
                continue
            signature = str(raw.get("signature") or self._profile_dataset_signature(dataset))
            dataset["display_name"] = self._dataset_display_name(dataset, dataset.get("display_name") or "Profile")
            entries.append({"dataset": dataset, "signature": signature})
        return entries

    def _merge_profile_entries(self, incoming_entries):
        merged = []
        seen = set()
        for group in (self._current_profile_entries(), incoming_entries):
            for entry in group:
                signature = str(entry.get("signature") or "")
                dataset = entry.get("dataset")
                if not isinstance(dataset, dict):
                    continue
                if signature and signature in seen:
                    continue
                if signature:
                    seen.add(signature)
                merged.append({"dataset": copy.deepcopy(dataset), "signature": signature})
        return merged

    def _register_workspace_dialog(self):
        if self._workspace_registered:
            return
        owner = getattr(self, "_owner", None)
        dialogs = getattr(owner, "_profile_dialogs", None)
        if dialogs is not None and self not in dialogs:
            dialogs.append(self)
        refs = getattr(owner, "_popup_refs", None)
        if refs is not None and self not in refs:
            refs.append(self)
        controller = getattr(owner, "quick_crop_controller", None)
        if controller:
            try:
                controller.update_popup_actions()
            except Exception:
                pass
        self._workspace_registered = True

    def _deregister_workspace_dialog(self):
        if not self._workspace_registered:
            return
        owner = getattr(self, "_owner", None)
        dialogs = getattr(owner, "_profile_dialogs", None)
        if dialogs is not None and self in dialogs:
            dialogs.remove(self)
        refs = getattr(owner, "_popup_refs", None)
        if refs is not None and self in refs:
            refs.remove(self)
        controller = getattr(owner, "quick_crop_controller", None)
        if controller:
            try:
                controller.update_popup_actions()
            except Exception:
                pass
        self._workspace_registered = False

    def _spawn_composite_dialog(self, merged_entries):
        datasets = [copy.deepcopy(entry.get("dataset") or {}) for entry in merged_entries if isinstance(entry.get("dataset"), dict)]
        if not datasets:
            return None
        active = datasets[0]
        saved = datasets[1:]
        owner = getattr(self, "_owner", None)
        unit = active.get("unit") or self._unit
        meta = dict(active.get("meta") or {})
        y_label = str(meta.get("channel") or self._y_label or "Profile value").strip()
        dlg = ProfileDialog(
            active,
            saved,
            parent=owner,
            unit=unit,
            y_label=y_label,
            dark_mode=bool(self._dark_background),
        )
        dlg._composite_mode = True
        dlg.setWindowTitle(f"Profile composite ({len(datasets)})")
        if hasattr(dlg, "detach_as_workspace_window"):
            dlg.detach_as_workspace_window()
        try:
            dlg.set_plot_typography(
                family=self._plot_font_family,
                bold=self._plot_font_bold,
                italic=self._plot_font_italic,
                underline=self._plot_font_underline,
            )
        except Exception:
            pass
        try:
            dlg._font_scale = float(getattr(self, "_font_scale", 1.0) or 1.0)
        except Exception:
            dlg._font_scale = 1.0
        try:
            dlg.show_lines_cb.setChecked(bool(self.show_lines_cb.isChecked()))
            dlg.show_points_cb.setChecked(bool(self.show_points_cb.isChecked()))
            dlg.grid_cb.setChecked(bool(self.grid_cb.isChecked()))
            dlg.dark_bg_cb.setChecked(bool(self.dark_bg_cb.isChecked()))
            dlg.extra_ticks_cb.setChecked(bool(self.extra_ticks_cb.isChecked()))
            dlg.precision_cb.setChecked(bool(self.precision_cb.isChecked()))
            dlg.multi_channel_cb.setChecked(bool(self.multi_channel_cb.isChecked()))
            dlg._set_advanced_options_visible(bool(self._advanced_controls_visible))
        except Exception:
            pass
        dlg.update_profiles(active, saved)
        dlg._apply_font_scale()
        try:
            base_geo = self.frameGeometry()
            dlg.move(base_geo.topLeft() + QtCore.QPoint(36, 36))
        except Exception:
            pass
        dlg.setAttribute(QtCore.Qt.WA_DeleteOnClose, True)
        dlg._register_workspace_dialog()
        dlg.finished.connect(lambda _=None, ref=dlg: ref._deregister_workspace_dialog())
        dlg.show()
        try:
            dlg.raise_()
            dlg.activateWindow()
        except Exception:
            pass
        return dlg

    def _create_composite_from_drop(self, payload):
        if not isinstance(payload, dict):
            return None
        if str(payload.get("origin_dialog_id") or "") == self._composite_origin_id:
            return None
        incoming_entries = self._entries_from_profile_payload(payload)
        if not incoming_entries:
            return None
        merged = self._merge_profile_entries(incoming_entries)
        if len(merged) <= 1:
            return None
        return self._spawn_composite_dialog(merged)

    def _refresh_action_button_states(self):
        try:
            self.add_btn.setEnabled(callable(getattr(self, "_add_overlay_cb", None)))
        except Exception:
            pass
        try:
            self.delete_btn.setEnabled(bool(self._active or self._saved))
        except Exception:
            pass
        try:
            self.compose_drag_btn.setEnabled(bool(self._active or self._saved))
        except Exception:
            pass

    def _set_compose_drop_active(self, active):
        btn = getattr(self, "compose_drag_btn", None)
        if btn is None:
            return
        try:
            btn.setProperty("dropActive", bool(active))
        except Exception:
            return
        self._apply_toggle_button_styles()

    def _show_compose_help(self, global_pos=None):
        message = (
            "Drag Combine onto another profile window.\n"
            "You can also drag from the plot margin to create a composite."
        )
        try:
            QtWidgets.QToolTip.showText(global_pos or QtGui.QCursor.pos(), message, self)
        except Exception:
            pass

    def _qt_pos_in_main_axes(self, pos):
        if pos is None:
            return False
        try:
            bbox = self.ax.get_window_extent()
        except Exception:
            return False
        if bbox is None:
            return False
        try:
            dpr = float(self.canvas.devicePixelRatioF())
        except Exception:
            dpr = 1.0
        try:
            height = self.canvas.height() * dpr
        except Exception:
            height = self.canvas.height()
        x = float(pos.x()) * dpr
        y = float(height - (pos.y() * dpr))
        try:
            return bool(bbox.contains(x, y))
        except Exception:
            return False

    def eventFilter(self, source, event):
        if source == self.canvas:
            etype = event.type()
            if etype == QtCore.QEvent.MouseButtonPress and event.button() == QtCore.Qt.LeftButton:
                if self._qt_pos_in_main_axes(event.pos()):
                    self._canvas_drag_start_pos = None
                    self._canvas_drag_started = False
                else:
                    self._canvas_drag_start_pos = event.pos()
                    self._canvas_drag_started = False
            elif (
                etype == QtCore.QEvent.MouseMove
                and self._canvas_drag_start_pos is not None
                and event.buttons() & QtCore.Qt.LeftButton
            ):
                if (event.pos() - self._canvas_drag_start_pos).manhattanLength() >= QtWidgets.QApplication.startDragDistance():
                    self._canvas_drag_start_pos = None
                    self._canvas_drag_started = True
                    try:
                        self.start_profile_composite_drag()
                    except Exception:
                        pass
                    return True
            elif etype == QtCore.QEvent.MouseButtonRelease and event.button() == QtCore.Qt.LeftButton:
                self._canvas_drag_start_pos = None
                self._canvas_drag_started = False
        return super().eventFilter(source, event)

    def dragEnterEvent(self, event):
        if event.mimeData().hasFormat(_PROFILE_COMPOSITE_MIME):
            self._set_compose_drop_active(True)
            event.acceptProposedAction()
            return
        self._set_compose_drop_active(False)
        event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasFormat(_PROFILE_COMPOSITE_MIME):
            self._set_compose_drop_active(True)
            event.acceptProposedAction()
            return
        self._set_compose_drop_active(False)
        event.ignore()

    def dragLeaveEvent(self, event):
        self._set_compose_drop_active(False)
        super().dragLeaveEvent(event)

    def dropEvent(self, event):
        if not event.mimeData().hasFormat(_PROFILE_COMPOSITE_MIME):
            self._set_compose_drop_active(False)
            event.ignore()
            return
        data = event.mimeData().data(_PROFILE_COMPOSITE_MIME)
        try:
            payload = json.loads(bytes(data).decode("utf-8"))
        except Exception:
            self._set_compose_drop_active(False)
            event.ignore()
            return
        dlg = self._create_composite_from_drop(payload)
        self._set_compose_drop_active(False)
        if dlg is None:
            QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), "No new composite created", self)
            event.ignore()
            return
        QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), "Composite profile created", dlg)
        event.acceptProposedAction()

    def refresh_linked_profiles(self, datasets_by_ref, source_id=None):
        if not self._composite_mode or not isinstance(datasets_by_ref, dict):
            return False
        changed = False

        def _refresh_dataset(dataset):
            if not isinstance(dataset, dict):
                return dataset, False
            key = profile_ref_key(dataset.get("live_profile_ref"))
            if key is None:
                return dataset, False
            if source_id and str(key[0]) != str(source_id):
                return dataset, False
            updated = datasets_by_ref.get(key)
            if not isinstance(updated, dict):
                return dataset, False
            refreshed = copy.deepcopy(updated)
            refreshed["display_name"] = self._dataset_display_name(
                refreshed,
                dataset.get("display_name") or refreshed.get("display_name") or "Profile",
            )
            return refreshed, True

        current_key = self._selected_overlay_index()
        self._active, active_changed = _refresh_dataset(self._active)
        changed = changed or active_changed
        for idx, dataset in enumerate(list(self._saved)):
            refreshed, item_changed = _refresh_dataset(dataset)
            if item_changed:
                self._saved[idx] = refreshed
                changed = True
        if not changed:
            return False
        self.update_profiles(
            self._active,
            self._saved,
            activate_overlay_callback=self._activate_overlay_cb,
            highlight_overlay_callback=self._highlight_overlay_cb,
        )
        self.select_overlay(current_key)
        return True

    def update_profiles(self, active_profile, saved_profiles=None, activate_overlay_callback=None,
                         highlight_overlay_callback=None):
        saved_profiles = saved_profiles or []
        self._active = active_profile
        self._saved = saved_profiles
        entries = self._ordered_profile_entries(active_profile, saved_profiles)
        self._ordered_profile_entries_cache = entries
        if len(saved_profiles) != self._last_saved_count:
            # Overlay indices may shift; drop overlay-specific marker positions.
            keep = self._marker_positions_by_key.get(None)
            keep_domain = self._marker_domain_by_key.get(None)
            self._marker_positions_by_key = {None: keep} if keep is not None else {}
            self._marker_domain_by_key = {None: keep_domain} if keep_domain is not None else {}
        self._last_saved_count = len(saved_profiles)
        if activate_overlay_callback is not None:
            self._activate_overlay_cb = activate_overlay_callback
        if highlight_overlay_callback is not None:
            self._highlight_overlay_cb = highlight_overlay_callback
        reference = active_profile or (saved_profiles[0] if saved_profiles else None)
        self._relative_axes = bool((reference or {}).get('relative_axes', True))
        self._line_handles = []
        self._line_handles_by_key = {}
        datasets = []
        for entry in entries:
            datasets.append((entry["label"], entry["data"], entry["is_active"], entry["key"]))
        if not datasets:
            self.stats.setText("No profile data")
            self.profile_list.blockSignals(True)
            self.profile_list.clear()
            self.profile_list.blockSignals(False)
            self._clear_marker_lines()
            self._refresh_action_button_states()
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
        self.ax_right.clear()
        self.ax_right.set_visible(False)
        self.ax.set_xlabel(self._axis_label(axis_label_unit))
        self._apply_ylabel(reference)
        show_points = bool(self.show_points_cb.isChecked()) if hasattr(self, 'show_points_cb') else False
        show_lines = bool(self.show_lines_cb.isChecked()) if hasattr(self, 'show_lines_cb') else True
        precision_mode = bool(self.precision_cb.isChecked()) if hasattr(self, 'precision_cb') else False
        marker_alpha = 0.55 if show_points else 1.0
        line_alpha_active = 0.9 if show_lines else marker_alpha
        line_alpha_overlay = 0.65 if show_lines else marker_alpha
        ref_points = None
        ref_length = None
        marker_dataset = active_profile if active_profile else (saved_profiles[0] if saved_profiles else None)
        
        # Add extra channels if requested
        if self.multi_channel_cb.isChecked() and active_profile and active_profile.get('extra_channels'):
            for extra in active_profile.get('extra_channels', []):
                datasets.append((extra.get('name', 'Extra'), extra, True, None))

        for label, data, is_active, profile_key in datasets:
            x = data.get('x_nm')
            if x is None:
                x = data.get('x_px')
            y = data.get('vals')
            if x is None or y is None:
                continue
            color = data.get('color') or ('#ffd54f' if is_active else '#80cbc4')
            lw = float(data.get('lw') or (1.5 if is_active else 1.0))
            alpha = line_alpha_active if is_active else line_alpha_overlay
            style_line = data.get('line_style') or ('-' if is_active else '--')
            if not show_lines or str(style_line).lower() == 'none':
                style_line = 'None'
            style_marker = data.get('marker_style') or 'o'
            if not show_points or str(style_marker).lower() == 'none':
                style_marker = None
            marker_size = float(data.get('marker_size') or (3.5 if is_active else 3.0)) if style_marker else None
            
            # Determine axis
            target_ax = self.ax
            data_unit = data.get('unit') or ''
            ref_unit = (reference.get('unit') if reference else '') or ''
            plot_label = self._dataset_display_name(data, label)
            if data_unit != ref_unit:
                target_ax = self.ax_right
                self.ax_right.set_visible(True)
                self.ax_right.set_ylabel(f"{plot_label} ({data_unit})")
                self.ax_right.yaxis.set_label_position("right")
                self.ax_right.yaxis.label.set_color(color)
                self.ax_right.tick_params(axis='y', colors=color)
                self.ax_right.spines['right'].set_color(color)
                self.ax_right.spines['right'].set_visible(True)

            line, = target_ax.plot(
                x, y, color=color, lw=lw, label=plot_label,
                linestyle=style_line, marker=style_marker, markersize=marker_size,
                markeredgewidth=0.9 if style_marker else 0.0,
                markerfacecolor='none' if style_marker else color,
                markeredgecolor=color if style_marker else 'none',
                markevery=1,
                alpha=alpha,
            )
            self._line_handles.append(line)
            if profile_key not in self._line_handles_by_key:
                self._line_handles_by_key[profile_key] = line
            if is_active and label == 'Active':
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
        if hasattr(self, 'extra_ticks_cb') and self.extra_ticks_cb.isChecked():
            try:
                self.ax.xaxis.set_minor_locator(AutoMinorLocator(4))
                self.ax.yaxis.set_minor_locator(AutoMinorLocator(4))
                self.ax.tick_params(which='minor', length=2.5, width=0.6, color='#7d7d7d')
            except Exception:
                pass
        if precision_mode:
            try:
                self.ax.xaxis.set_minor_locator(AutoMinorLocator(5))
                self.ax.yaxis.set_minor_locator(AutoMinorLocator(5))
                self.ax.tick_params(which='minor', length=2.0, width=0.5, color='#6b6b6b')
            except Exception:
                pass
        if len(datasets) > 1 and self._legend_visible:
            try:
                legend_kwargs = {
                    "fontsize": self._legend_fontsize * getattr(self, "_font_scale", 1.0),
                    "loc": self._legend_loc,
                }
                if self._legend_custom_anchor is not None:
                    legend_kwargs["loc"] = "upper left"
                    legend_kwargs["bbox_to_anchor"] = self._legend_custom_anchor
                    legend_kwargs["bbox_transform"] = self.ax.transAxes
                self.ax.legend(**legend_kwargs)
            except Exception:
                pass
        else:
            try:
                existing = self.ax.get_legend()
                if existing is not None:
                    existing.remove()
            except Exception:
                pass
        try:
            self.canvas.figure.subplots_adjust(top=0.92)
        except Exception:
            pass
        self._apply_metadata_overlay(reference)
        self._apply_plot_theme()
        self.stats.setText(self._format_stats_text(active_profile, saved_profiles))
        self._populate_profile_list(active_profile, saved_profiles)
        valid_keys = {entry["key"] for entry in entries}
        if self._current_marker_key not in valid_keys:
            if active_profile is not None:
                self._current_marker_key = None
            else:
                self._current_marker_key = entries[0]["key"] if entries else None
        self.select_overlay(self._current_marker_key)
        self._reset_markers(ref_points, ref_length, reference_dataset=marker_dataset)
        if self._current_marker_key in self._marker_positions_by_key:
            positions = self._marker_positions_by_key.get(self._current_marker_key)
            domain = self._marker_domain_by_key.get(self._current_marker_key)
            if positions:
                self.set_marker_positions(positions, domain=domain)
        if callable(self._marker_key_cb):
            self._marker_key_cb(self._current_marker_key)
        self._apply_font_scale()
        self._refresh_action_button_states()
        try:
            if hasattr(self, "_splitter") and self._splitter is not None:
                total = max(1, self.height())
                self._splitter.setSizes([int(total * 0.7), int(total * 0.3)])
        except Exception:
            pass
        
        # Re-apply right axis styling if it was used, as _apply_plot_theme may have reset colors
        if self.ax_right.get_visible():
            reference = active_profile or (saved_profiles[0] if saved_profiles else None)
            ref_unit = (reference.get('unit') if reference else '') or ''
            for label, data, is_active in datasets:
                data_unit = data.get('unit') or ''
                if data_unit != ref_unit:
                    color = data.get('color') or ('#ffd54f' if is_active else '#80cbc4')
                    self.ax_right.yaxis.label.set_color(color)
                    self.ax_right.tick_params(axis='y', colors=color)
                    self.ax_right.spines['right'].set_color(color)
                    break

        self.canvas.draw_idle()

    def _populate_profile_list(self, active_profile, saved_profiles):
        self.profile_list.blockSignals(True)
        self.profile_list.clear()
        target_item = None
        for entry in self._ordered_profile_entries(active_profile, saved_profiles):
            data = entry["data"] or {}
            key = entry["key"]
            label = entry["label"]
            text = self._fmt_length(self._dataset_display_name(data, label), data.get('length_nm'))
            item = QtWidgets.QListWidgetItem(text)
            self._apply_item_color(item, data.get('color'))
            item.setData(QtCore.Qt.UserRole, key)
            self.profile_list.addItem(item)
            if target_item is None and key == self._current_marker_key:
                target_item = item
            elif target_item is None:
                target_item = item
        if target_item:
            self.profile_list.setCurrentItem(target_item)
        self.profile_list.blockSignals(False)
        if target_item:
            self._on_profile_item_selected(target_item)

    def _apply_item_color(self, item, color):
        if item is None or not color:
            return
        try:
            pix = QtGui.QPixmap(12, 12)
            pix.fill(QtGui.QColor(color))
            item.setIcon(QtGui.QIcon(pix))
        except Exception:
            pass

    def set_label_scale_callback(self, cb):
        self._label_scale_cb = cb
        if callable(self._label_scale_cb):
            self._label_scale_cb(self._font_scale)

    def set_marker_update_callback(self, cb):
        self._marker_update_cb = cb

    def set_marker_select_callback(self, cb):
        self._marker_key_cb = cb

    def set_add_overlay_callback(self, cb):
        self._add_overlay_cb = cb
        self._refresh_action_button_states()

    def set_delete_overlay_callback(self, cb):
        self._delete_overlay_cb = cb
        self._refresh_action_button_states()

    def set_style_update_callback(self, cb):
        self._style_update_cb = cb

    def set_palette_callback(self, cb):
        self._palette_cb = cb

    def set_marker_positions(self, positions, domain=None):
        if self._marker_syncing:
            return
        try:
            self._marker_syncing = True
            if positions is None or len(positions) < 2:
                self._marker_saved_positions = None
                if self._current_marker_key in self._marker_positions_by_key:
                    self._marker_positions_by_key.pop(self._current_marker_key, None)
                    self._marker_domain_by_key.pop(self._current_marker_key, None)
                self._clear_marker_lines(reset_saved=False)
                return
            if domain is not None:
                self._marker_domain = tuple(domain)
            self._marker_positions = [self._clamp_marker(p) for p in positions]
            if self._current_marker_key is not None:
                self._marker_positions_by_key[self._current_marker_key] = list(self._marker_positions)
                self._marker_domain_by_key[self._current_marker_key] = tuple(self._marker_domain)
            else:
                self._marker_positions_by_key[None] = list(self._marker_positions)
                self._marker_domain_by_key[None] = tuple(self._marker_domain)
            if not self._ensure_marker_lines():
                self._clear_marker_lines(reset_saved=False)
                return
            for idx, pos in enumerate(self._marker_positions):
                try:
                    self._marker_lines[idx].set_xdata([pos, pos])
                except Exception:
                    pass
            self._update_marker_info()
        finally:
            self._marker_syncing = False

    def _notify_marker_positions(self):
        if self._marker_syncing:
            return
        if not callable(self._marker_update_cb):
            return
        if not self._markers_enabled or len(self._marker_positions) < 2:
            self._marker_update_cb(None, None)
            return
        if self._current_marker_key is not None:
            self._marker_positions_by_key[self._current_marker_key] = list(self._marker_positions)
            self._marker_domain_by_key[self._current_marker_key] = tuple(self._marker_domain)
        else:
            self._marker_positions_by_key[None] = list(self._marker_positions)
            self._marker_domain_by_key[None] = tuple(self._marker_domain)
        self._marker_update_cb(list(self._marker_positions), tuple(self._marker_domain))

    def _on_theme_toggled(self, _checked=False):
        self._dark_background = bool(self.dark_bg_cb.isChecked())
        self._apply_plot_theme()

    def _apply_legend_style(self):
        legend = self.ax.get_legend()
        self._legend_artist = legend
        if legend is None:
            return
        dark = bool(self._dark_background)
        ax_face = '#14161c' if dark else '#ffffff'
        text = '#f5f5f5' if dark else '#111111'
        try:
            legend.set_visible(bool(self._legend_visible))
            if self._legend_custom_anchor is not None:
                legend.set_bbox_to_anchor(self._legend_custom_anchor, transform=self.ax.transAxes)
                legend._loc = 2  # upper left
            else:
                legend.set_bbox_to_anchor(None)
            frame = legend.get_frame()
            frame.set_facecolor(ax_face if self._legend_frame_fill_visible else ax_face)
            frame.set_alpha(0.88 if self._legend_frame_fill_visible else 0.0)
            frame.set_linewidth(float(self._legend_outline_width))
            frame.set_edgecolor(text if self._legend_outline_visible else 'none')
            for txt in legend.get_texts():
                txt.set_color(text)
                txt.set_fontsize(self._legend_fontsize * getattr(self, '_font_scale', 1.0))
                apply_text_style(txt, family=self._plot_font_family, **self._font_style_state())
        except Exception:
            pass

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
        for axis in (self.ax, self.ax_top, self.ax_right):
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
            if grid_on:
                self.ax.grid(True, color=grid_color, alpha=0.35)
            else:
                self.ax.grid(False)
        except Exception:
            pass
        if self._metadata_artist is not None:
            try:
                self._metadata_artist.set_color(text)
                self._metadata_artist.set_fontsize(max(5.0, 6.0 * getattr(self, '_font_scale', 1.0)))
                apply_text_style(self._metadata_artist, family=self._plot_font_family, **self._font_style_state())
                patch = self._metadata_artist.get_bbox_patch()
                if patch is not None:
                    patch.set_facecolor(ax_face)
                    patch.set_alpha(0.78)
                    patch.set_edgecolor('none')
            except Exception:
                pass
        self._apply_legend_style()
        self._apply_toggle_button_styles()
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
        self._current_marker_key = idx
        if callable(self._marker_key_cb):
            try:
                self._marker_key_cb(self._current_marker_key)
            except Exception:
                pass
        # adjust highlight on plotted lines
        for profile_key, line in list(self._line_handles_by_key.items()):
            try:
                dataset = self._dataset_for_profile_key(profile_key) or {}
                base_lw = float(dataset.get('lw') or (1.5 if profile_key is None else 1.0))
                line.set_linewidth(base_lw)
            except Exception:
                pass
        if idx is None:
            line = self._line_handles_by_key.get(None)
            if line is not None:
                try:
                    dataset = self._active or {}
                    base_lw = float(dataset.get('lw') or 1.5)
                    line.set_linewidth(base_lw + 0.4)
                except Exception:
                    pass
        else:
            line = self._line_handle_for_overlay(idx)
            if line is not None:
                try:
                    dataset = self._saved[idx] if 0 <= idx < len(self._saved) else {}
                    base_lw = float(dataset.get('lw') or 1.0)
                    line.set_linewidth(base_lw + 0.4)
                except Exception:
                    pass
        self.canvas.draw_idle()
        if self._highlight_overlay_cb:
            try:
                self._highlight_overlay_cb(idx)
            except Exception:
                pass
        dataset = None
        if idx is None:
            dataset = self._active
        elif idx >= 0 and idx < len(self._saved):
            dataset = self._saved[idx]
        if dataset:
            ref_points = dataset.get('x_nm') if dataset.get('x_nm') is not None else dataset.get('x_px')
            ref_length = dataset.get('length_nm')
            self._reset_markers(ref_points, ref_length, reference_dataset=dataset, store_state=False)
            if idx in self._marker_positions_by_key:
                positions = self._marker_positions_by_key.get(idx)
                domain = self._marker_domain_by_key.get(idx)
                if positions:
                    self.set_marker_positions(positions, domain=domain)
            else:
                self._marker_positions_by_key[idx] = list(self._marker_positions)
                self._marker_domain_by_key[idx] = tuple(self._marker_domain)

    def _line_handle_for_overlay(self, idx):
        if idx is None:
            return None
        return self._line_handles_by_key.get(idx)

    def _on_profile_item_activated(self, item):
        if item is None:
            return
        # Avoid destructive double-click behavior; selection is enough.
        return

    def _add_overlay_from_active(self):
        if callable(self._add_overlay_cb):
            self._add_overlay_cb()

    def _selected_overlay_index(self):
        """Return the currently selected saved-profile index, if any."""
        candidates = []
        current = self.profile_list.currentItem()
        if current is not None:
            candidates.append(current)
        for item in self.profile_list.selectedItems():
            if item not in candidates:
                candidates.append(item)
        for item in candidates:
            try:
                idx = item.data(QtCore.Qt.UserRole)
            except Exception:
                idx = None
            if idx is None:
                continue
            try:
                return int(idx)
            except Exception:
                continue
        current_key = getattr(self, "_current_marker_key", None)
        if current_key is not None:
            try:
                return int(current_key)
            except Exception:
                pass
        return None

    def _delete_selected_profile(self):
        current = self.profile_list.currentItem()
        if current is None:
            QtWidgets.QMessageBox.information(self, "Delete profile", "Select a profile to delete.")
            return
        idx = current.data(QtCore.Qt.UserRole)
        if callable(self._delete_overlay_cb) and idx is not None:
            removed = bool(self._delete_overlay_cb(idx))
            if removed and 0 <= idx < len(self._saved):
                saved = list(self._saved)
                saved.pop(idx)
                self.update_profiles(
                    self._active,
                    saved,
                    activate_overlay_callback=self._activate_overlay_cb,
                    highlight_overlay_callback=self._highlight_overlay_cb,
                )
            return
        if callable(self._delete_overlay_cb) and idx is None:
            QtWidgets.QMessageBox.information(self, "Delete profile", "The active live profile cannot be deleted here.")
            return
        active = copy.deepcopy(self._active) if self._active is not None else None
        saved = copy.deepcopy(list(self._saved or []))
        if idx is None:
            active = saved.pop(0) if saved else None
        else:
            try:
                saved.pop(int(idx))
            except Exception:
                QtWidgets.QMessageBox.information(self, "Delete profile", "Select a valid profile to delete.")
                return
        self.update_profiles(
            active,
            saved,
            activate_overlay_callback=self._activate_overlay_cb,
            highlight_overlay_callback=self._highlight_overlay_cb,
        )

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
        self._deregister_workspace_dialog()
        unregister_profile_dialog(self)
        super().closeEvent(event)

    def _copy_current_profile(self):
        datasets = []
        if self._active:
            datasets.append(("Active", self._active))
        for idx, data in enumerate(self._saved, 1):
            datasets.append((f"Overlay {idx}", data))
        if not datasets:
            QtWidgets.QMessageBox.information(self, "Copy profile", "No profile data available.")
            return
        meta = (self._active or (self._saved[0] if self._saved else {})).get('meta') or {}
        channel_label = self._y_label or meta.get('channel') or "Value"
        channel_unit = (self._active or (self._saved[0] if self._saved else {})).get('unit') or ""
        header = [f"Channel: {channel_label}{f' [{channel_unit}]' if channel_unit else ''}"]
        if meta.get('file_name'):
            header.append(f"Image: {meta.get('file_name')}")
        if meta.get('date') or meta.get('time'):
            header.append(f"Date: {meta.get('date','')} Time: {meta.get('time','')}".strip())
        if meta.get('datetime'):
            header.append(f"Timestamp: {meta.get('datetime')}")
        blocks = ["\n".join(header)]
        columns = []
        max_len = 0
        for name, dataset in datasets:
            x = dataset.get('x_nm')
            unit = dataset.get('axis_unit') or dataset.get('distance_unit') or 'nm'
            if x is None:
                x = dataset.get('x_px')
                unit = 'px'
            vals = dataset.get('vals')
            if x is None or vals is None:
                continue
            x = list(x)
            vals = list(vals)
            max_len = max(max_len, len(x), len(vals))
            columns.append((name, unit, x, vals))
        if not columns:
            QtWidgets.QMessageBox.information(self, "Copy profile", "Profile data is incomplete.")
            return
        header_row = []
        for name, unit, _x, _vals in columns:
            header_row.append(f"{name} d ({unit})")
            header_row.append(f"{name} {channel_label} ({channel_unit})".rstrip())
        rows = ["\t".join(header_row)]
        for i in range(max_len):
            row = []
            for _name, _unit, x, vals in columns:
                try:
                    dist = x[i]
                    row.append(f"{float(dist):.9g}")
                except Exception:
                    row.append("")
                try:
                    val = vals[i]
                    row.append(f"{float(val):.9g}")
                except Exception:
                    row.append("")
            rows.append("\t".join(row))
        blocks.append("\n".join(rows))
        QtWidgets.QApplication.clipboard().setText("\n\n".join(blocks))
        QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), "Profiles copied", self)
