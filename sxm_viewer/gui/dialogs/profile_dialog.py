"""Detail canvases and spectroscopy dialogs."""
from __future__ import annotations

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
from ..plot_typography import add_font_menu_action, normalize_font_family
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
from .profile_data import axis_label, format_marker_delta, format_stats_text, fmt_length

class ProfileDialog(QtWidgets.QDialog):
    """Dialog showing the sampled profile and basic stats."""
    def __init__(self, active_profile, saved_profiles=None, parent=None, unit=None, y_label=None,
                 activate_overlay_callback=None, highlight_overlay_callback=None,
                 label_scale_callback=None, delete_overlay_callback=None,
                 marker_update_callback=None, marker_select_callback=None,
                 add_overlay_callback=None, dark_mode=False):
        super().__init__(parent)
        self.setWindowTitle('Profile measurement')
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
        self._marker_syncing = False
        self._marker_positions_by_key = {}
        self._marker_domain_by_key = {}
        self._current_marker_key = None
        self._last_saved_count = 0
        self._toggle_buttons = []
        self._advanced_controls_visible = False
        owner = self.parent()
        self._plot_font_family = normalize_font_family(getattr(owner, "_plot_font_family", None), "sans-serif")
        v = QtWidgets.QVBoxLayout()
        fig = Figure(figsize=(6,3))
        self.canvas = SafeFigureCanvas(fig)
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
        info_layout.addWidget(self.stats)
        self.marker_info = QtWidgets.QLabel("")
        self.marker_info.setAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignTop)
        info_layout.addWidget(self.marker_info)
        controls_hint = QtWidgets.QLabel(
            "Shortcuts: V markers, G grid, L lines, P points, Del remove overlay, Ctrl+Wheel font size"
        )
        controls_hint.setObjectName("profileControlsHint")
        controls_hint.setWordWrap(True)
        info_layout.addWidget(controls_hint)

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
        self.profile_list.itemDoubleClicked.connect(self._on_profile_item_activated)
        self.profile_list.currentItemChanged.connect(self._on_profile_item_selected)
        info_layout.addWidget(QtWidgets.QLabel("Profiles"))
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
        self._context_source = None
        self._context_syncing = False
        self._preserve_cb = None

    def _on_context_menu(self, pos):
        menu = QtWidgets.QMenu(self)
        copy_png = menu.addAction("Copy plot (PNG)")
        copy_svg = menu.addAction("Copy plot (SVG)")
        add_font_menu_action(menu, self, self._plot_font_family, self.set_plot_font_family)
        action = menu.exec_(self.canvas.mapToGlobal(pos))
        if action == copy_png:
            self._copy_plot("png")
        elif action == copy_svg:
            self._copy_plot("svg")

    def set_plot_font_family(self, family: str):
        """Rebuild the profile plot with a new shared font family."""
        owner = self.parent()
        if owner is not None and hasattr(owner, "set_plot_font_family") and getattr(owner, "_plot_font_family", None) != family:
            try:
                owner.set_plot_font_family(family)
                return
            except Exception:
                pass
        self._plot_font_family = normalize_font_family(family, "sans-serif")
        self.update_profiles(self._active, self._saved)

    def _copy_plot(self, fmt):
        buf = io.BytesIO()
        if fmt == "svg":
            with matplotlib.rc_context({'svg.fonttype': 'none'}):
                self.canvas.figure.savefig(buf, format="svg", bbox_inches='tight')
            data = buf.getvalue()
            mime = QtCore.QMimeData()
            mime.setData("image/svg+xml", data)
            QtWidgets.QApplication.clipboard().setMimeData(mime)
        else:
            self.canvas.figure.savefig(buf, format="png", dpi=300, bbox_inches='tight')
            qimg = QtGui.QImage.fromData(buf.getvalue())
            QtWidgets.QApplication.clipboard().setImage(qimg)

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
        if not self._toggle_buttons:
            return
        dark = bool(self._dark_background)
        if dark:
            inactive_bg = "#1e2430"
            inactive_border = "#46556e"
            inactive_text = "#d4deee"
            active_bg = "#2f6fcb"
            active_border = "#79a9f2"
            active_text = "#ffffff"
            hint_color = "#b9c6d8"
        else:
            inactive_bg = "#f3f5f9"
            inactive_border = "#aeb7c5"
            inactive_text = "#1f2a3d"
            active_bg = "#1f6fd7"
            active_border = "#5b97e8"
            active_text = "#ffffff"
            hint_color = "#4b5b73"
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
        for btn in getattr(self, "_toggle_buttons", []):
            try:
                font = btn.font()
                font.setPointSizeF(max(7.0, 8.8 * scale))
                btn.setFont(font)
            except Exception:
                pass
        for btn in (getattr(self, "copy_btn", None), getattr(self, "add_btn", None), getattr(self, "delete_btn", None), getattr(self, "close_btn", None)):
            if btn is None:
                continue
            try:
                font = btn.font()
                font.setPointSizeF(max(7.0, 9.0 * scale))
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

    def update_profiles(self, active_profile, saved_profiles=None, activate_overlay_callback=None,
                         highlight_overlay_callback=None):
        saved_profiles = saved_profiles or []
        self._active = active_profile
        self._saved = saved_profiles
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
        self.ax_right.clear()
        self.ax_right.set_visible(False)
        self.ax.set_xlabel(self._axis_label(axis_label_unit))
        self._apply_ylabel(reference)
        show_points = bool(self.show_points_cb.isChecked()) if hasattr(self, 'show_points_cb') else False
        show_lines = bool(self.show_lines_cb.isChecked()) if hasattr(self, 'show_lines_cb') else True
        precision_mode = bool(self.precision_cb.isChecked()) if hasattr(self, 'precision_cb') else False
        marker_style = 'o' if show_points else None
        line_style = '-' if show_lines else 'None'
        marker_size = 3.0 if show_points else None
        marker_alpha = 0.55 if show_points else 1.0
        line_alpha_active = 0.9 if show_lines else marker_alpha
        line_alpha_overlay = 0.65 if show_lines else marker_alpha
        ref_points = None
        ref_length = None
        marker_dataset = active_profile if active_profile else (saved_profiles[0] if saved_profiles else None)
        
        # Add extra channels if requested
        if self.multi_channel_cb.isChecked() and active_profile and active_profile.get('extra_channels'):
            for extra in active_profile.get('extra_channels', []):
                datasets.append((extra.get('name', 'Extra'), extra, True))

        for label, data, is_active in datasets:
            x = data.get('x_nm')
            if x is None:
                x = data.get('x_px')
            y = data.get('vals')
            if x is None or y is None:
                continue
            color = data.get('color') or ('#ffd54f' if is_active else '#80cbc4')
            lw = 1.5 if is_active else 1.0
            alpha = line_alpha_active if is_active else line_alpha_overlay
            
            # Determine axis
            target_ax = self.ax
            data_unit = data.get('unit') or ''
            ref_unit = (reference.get('unit') if reference else '') or ''
            if data_unit != ref_unit:
                target_ax = self.ax_right
                self.ax_right.set_visible(True)
                self.ax_right.set_ylabel(f"{label} ({data_unit})")
                self.ax_right.yaxis.set_label_position("right")
                self.ax_right.yaxis.label.set_color(color)
                self.ax_right.tick_params(axis='y', colors=color)
                self.ax_right.spines['right'].set_color(color)
                self.ax_right.spines['right'].set_visible(True)

            line, = target_ax.plot(
                x, y, color=color, lw=lw, label=label,
                linestyle=line_style, marker=marker_style, markersize=marker_size,
                markeredgewidth=0.9 if show_points else 0.0,
                markerfacecolor='none' if show_points else color,
                markeredgecolor=color if show_points else 'none',
                markevery=1,
                alpha=alpha,
            )
            self._line_handles.append(line)
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
        if len(datasets) > 1:
            try:
                self.ax.legend(fontsize=8, loc='upper right')
            except Exception:
                pass
        self._apply_plot_theme()
        self.stats.setText(self._format_stats_text(active_profile, saved_profiles))
        self._populate_profile_list(active_profile, saved_profiles)
        if active_profile is not None:
            self._current_marker_key = None
        else:
            self._current_marker_key = 0 if saved_profiles else None
        self._reset_markers(ref_points, ref_length, reference_dataset=marker_dataset)
        if self._current_marker_key in self._marker_positions_by_key:
            positions = self._marker_positions_by_key.get(self._current_marker_key)
            domain = self._marker_domain_by_key.get(self._current_marker_key)
            if positions:
                self.set_marker_positions(positions, domain=domain)
        if callable(self._marker_key_cb):
            self._marker_key_cb(self._current_marker_key)
        self._apply_font_scale()
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
        if active_profile:
            item = QtWidgets.QListWidgetItem(self._fmt_length("Active", active_profile.get('length_nm')))
            self._apply_item_color(item, active_profile.get('color'))
            item.setData(QtCore.Qt.UserRole, None)
            self.profile_list.addItem(item)
            target_item = item
        for idx, data in enumerate(saved_profiles, 1):
            text = self._fmt_length(f"Overlay {idx}", data.get('length_nm'))
            item = QtWidgets.QListWidgetItem(text)
            self._apply_item_color(item, data.get('color'))
            item.setData(QtCore.Qt.UserRole, idx - 1)
            self.profile_list.addItem(item)
            if target_item is None:
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
        legend = self.ax.get_legend()
        if legend is not None:
            try:
                legend.get_frame().set_facecolor(ax_face)
                legend.get_frame().set_edgecolor(text)
                for txt in legend.get_texts():
                    txt.set_color(text)
            except Exception:
                pass
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
        offset = 1 if self._active else 0
        target = idx + offset
        if target < 0 or target >= len(self._line_handles):
            return None
        return self._line_handles[target]

    def _on_profile_item_activated(self, item):
        if item is None:
            return
        # Avoid destructive double-click behavior; selection is enough.
        return

    def _add_overlay_from_active(self):
        if callable(self._add_overlay_cb):
            self._add_overlay_cb()

    def _delete_selected_profile(self):
        current = self.profile_list.currentItem()
        idx = current.data(QtCore.Qt.UserRole) if current is not None else None
        if idx is None:
            QtWidgets.QMessageBox.information(self, "Delete profile", "Select an overlay to delete.")
            return
        if callable(self._delete_overlay_cb):
            self._delete_overlay_cb(idx)

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
