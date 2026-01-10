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
        v = QtWidgets.QVBoxLayout()
        fig = Figure(figsize=(6,3))
        self.canvas = SafeFigureCanvas(fig)
        self.ax = fig.add_subplot(111)
        self.ax_top = self.ax.twiny()
        self.ax_top.set_visible(False)
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
        toggle_layout = QtWidgets.QHBoxLayout()
        toggle_layout.setContentsMargins(0, 0, 0, 0)
        self.marker_toggle = QtWidgets.QCheckBox("Show measurement markers")
        self.marker_toggle.setChecked(True)
        self.marker_toggle.toggled.connect(self._on_marker_toggle)
        toggle_layout.addWidget(self.marker_toggle)
        toggle_layout.addStretch(1)
        info_layout.addLayout(toggle_layout)
        theme_layout = QtWidgets.QHBoxLayout()
        theme_layout.setContentsMargins(0, 0, 0, 0)
        self.dark_bg_cb = QtWidgets.QCheckBox("Dark background")
        self.dark_bg_cb.setChecked(self._dark_background)
        self.dark_bg_cb.toggled.connect(self._on_theme_toggled)
        theme_layout.addWidget(self.dark_bg_cb)
        self.grid_cb = QtWidgets.QCheckBox("Show grid")
        self.grid_cb.setChecked(False)
        self.grid_cb.toggled.connect(self._on_theme_toggled)
        theme_layout.addWidget(self.grid_cb)
        theme_layout.addStretch(1)
        info_layout.addLayout(theme_layout)
        plot_layout = QtWidgets.QHBoxLayout()
        plot_layout.setContentsMargins(0, 0, 0, 0)
        self.show_points_cb = QtWidgets.QCheckBox("Show data points")
        self.show_points_cb.setChecked(False)
        self.show_points_cb.toggled.connect(self._on_plot_option_changed)
        plot_layout.addWidget(self.show_points_cb)
        self.show_lines_cb = QtWidgets.QCheckBox("Show connecting lines")
        self.show_lines_cb.setChecked(True)
        self.show_lines_cb.toggled.connect(self._on_plot_option_changed)
        plot_layout.addWidget(self.show_lines_cb)
        self.extra_ticks_cb = QtWidgets.QCheckBox("Extra ticks")
        self.extra_ticks_cb.setChecked(False)
        self.extra_ticks_cb.toggled.connect(self._on_plot_option_changed)
        plot_layout.addWidget(self.extra_ticks_cb)
        self.precision_cb = QtWidgets.QCheckBox("Precision mode")
        self.precision_cb.setChecked(False)
        self.precision_cb.toggled.connect(self._on_plot_option_changed)
        plot_layout.addWidget(self.precision_cb)
        # Preview control disabled because the dialog preview is commented out to save resources.
        # self.preview_toggle_cb = QtWidgets.QCheckBox("Show preview")
        # self.preview_toggle_cb.setChecked(True)
        # self.preview_toggle_cb.toggled.connect(self._on_preview_toggle)
        # plot_layout.addWidget(self.preview_toggle_cb)
        # (If you re-enable the preview panel above, uncomment these lines to restore the toggle.)
        self.preserve_profiles_cb = QtWidgets.QCheckBox("Keep profiles on channel switch")
        self.preserve_profiles_cb.setChecked(True)
        self.preserve_profiles_cb.toggled.connect(self._on_preserve_toggle)
        plot_layout.addWidget(self.preserve_profiles_cb)
        plot_layout.addStretch(1)
        info_layout.addLayout(plot_layout)
        self.profile_list = QtWidgets.QListWidget()
        self.profile_list.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
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
            self.ax.xaxis.label.set_fontsize(label_size)
            self.ax.yaxis.label.set_fontsize(label_size)
            self.ax_top.xaxis.label.set_fontsize(label_size)
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
        if self._marker_arrow is not None:
            try: self._marker_arrow.remove()
            except Exception: pass
        if self._marker_label is not None:
            try: self._marker_label.remove()
            except Exception: pass
        arrow_color = "#f5f5f5" if self._dark_background else "#111111"
        arrow = self.ax.annotate(
            "",
            xy=(xmax, y_level),
            xytext=(xmin, y_level),
            arrowprops=dict(arrowstyle="<->", color=arrow_color, lw=1.8),
            annotation_clip=False,
        )
        display_value, display_unit = self._format_marker_delta(axis_delta)
        text = f"{display_value:.3f} {display_unit}"
        label_size = 9.0 * getattr(self, '_font_scale', 1.0)
        bbox_face = "#050506" if self._dark_background else "white"
        label = self.ax.text(
            (xmin + xmax) / 2.0,
            y_level + 0.02 * (y_max - y_min),
            text,
            color=arrow_color,
            ha="center",
            va="bottom",
            fontsize=label_size,
            bbox=dict(boxstyle="round,pad=0.2", facecolor=bbox_face,
                      alpha=0.7 if not self._dark_background else 0.6, edgecolor="none"),
        )
        self._marker_arrow = arrow
        self._marker_label = label
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
            line, = self.ax.plot(
                x, y, color=color, lw=lw, label=label,
                linestyle=line_style, marker=marker_style, markersize=marker_size,
                markeredgewidth=0.9 if show_points else 0.0,
                markerfacecolor='none' if show_points else color,
                markeredgecolor=color if show_points else 'none',
                markevery=1,
                alpha=alpha,
            )
            self._line_handles.append(line)
            if is_active:
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
        for axis in (self.ax, self.ax_top):
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




