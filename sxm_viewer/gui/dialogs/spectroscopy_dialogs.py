"""Detail canvases and spectroscopy dialogs."""
from __future__ import annotations

import functools
import itertools
import json
import math

import numpy as np
from matplotlib import patches
from matplotlib.backend_bases import MouseButton
from matplotlib import colors as mcolors
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from matplotlib.ticker import FuncFormatter
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
from ...data.channel_units import guess_channel_unit
from ..palettes import list_color_cycles, get_color_cycle, DEFAULT_COLOR_CYCLE
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
from .matrix_fit import MatrixFitDialog

class SpectroscopyPopup(QtWidgets.QDialog):
    """Popup window showing spectroscopy curves for a given file."""
    SCIENCE_PALETTE = [
        "#1f77b4", "#aec7e8", "#ff7f0e", "#ffbb78", "#2ca02c", "#98df8a",
        "#d62728", "#ff9896", "#9467bd", "#c5b0d5", "#8c564b", "#c49c94",
        "#e377c2", "#f7b6d2", "#7f7f7f", "#c7c7c7", "#bcbd22", "#dbdb8d",
        "#17becf", "#9edae5", "#393b79", "#5254a3", "#6b6ecf", "#9c9ede",
        "#e41a1c", "#377eb8", "#4daf4a", "#984ea3", "#ff7f00", "#ffff33",
        "#a65628", "#f781bf", "#999999", "#66c2a5", "#fc8d62", "#8da0cb",
        "#e78ac3", "#a6d854", "#ffd92f", "#e5c494", "#b3b3b3", "#8b9dc3",
        "#f96855", "#56a3a6", "#9f5f9d", "#2d5d82", "#73c2ff", "#ffaec9"
    ]
    def __init__(self, spec, parent=None):
        super().__init__(parent)
        self.spec = spec
        self.setWindowTitle(f"Spectroscopy: {Path(spec['path']).name}")
        self.resize(720, 520)
        layout = QtWidgets.QVBoxLayout()
        meta_txt = f"File: {Path(spec['path']).name}\nPosition: {spec.get('x','?')}/{spec.get('y','?')} nm\nTime: {spec.get('time')}"
        self.meta_label = QtWidgets.QLabel(meta_txt)
        self.meta_label.setWordWrap(True)
        layout.addWidget(self.meta_label)

        selector_layout = QtWidgets.QHBoxLayout()
        selector_layout.addWidget(QtWidgets.QLabel("Channel:"))
        self.channel_combo = QtWidgets.QComboBox()
        selector_layout.addWidget(self.channel_combo, 1)
        self.fit_btn = QtWidgets.QPushButton("Fit parabola")
        self.copy_btn = QtWidgets.QPushButton("Copy channel")
        selector_layout.addWidget(self.fit_btn)
        selector_layout.addWidget(self.copy_btn)
        layout.addLayout(selector_layout)

        self.fig = Figure(figsize=(6,4))
        self.canvas = FigureCanvas(self.fig)
        self.ax = self.fig.add_subplot(111)
        self._active_line_color = self.SCIENCE_PALETTE[0]
        self._swatch_buttons = []
        self._curve_entries = []
        self._selected_curve_index = 0
        self._drag_start_pos = None
        self._font_scale = 1.0
        self.setAcceptDrops(True)
        self.canvas.installEventFilter(self)
        self._palette_swatches = self._create_palette_swatch_widget()
        layout.addWidget(self.canvas, 1)
        layout.addWidget(self._palette_swatches)
        self.curve_list = QtWidgets.QListWidget()
        self.curve_list.setFixedHeight(72)
        self.curve_list.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.curve_list.currentRowChanged.connect(self._on_curve_selection_changed)
        layout.addWidget(self.curve_list)
        self.canvas.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.canvas.customContextMenuRequested.connect(self._on_canvas_context_menu)
        self.fit_result_label = QtWidgets.QLabel("")
        self.fit_result_label.setWordWrap(True)
        layout.addWidget(self.fit_result_label)
        self.setLayout(layout)

        self.axes = self._collect_axes(spec)
        self.V = np.asarray(self.axes[0]["values"], dtype=float) if self.axes else np.asarray([], dtype=float)
        self.axis_label = self.axes[0].get("label") if self.axes else "Axis"
        self.axis_unit = self.axes[0].get("unit") if self.axes else ""
        self.channels = {name: np.asarray(vals, dtype=float) for name, vals in (spec.get('channels', {}) or {}).items()}
        # Axis selector
        selector_layout2 = QtWidgets.QHBoxLayout()
        selector_layout2.addWidget(QtWidgets.QLabel("Axis:"))
        self.axis_combo = QtWidgets.QComboBox()
        for ax in self.axes:
            self.axis_combo.addItem(self._axis_display_name(ax), ax.get("key"))
        self.axis_combo.currentIndexChanged.connect(self._on_axis_changed)
        selector_layout2.addWidget(self.axis_combo, 1)
        layout.addLayout(selector_layout2)
        for name in self.channels.keys():
            self.channel_combo.addItem(name)
        if self.channel_combo.count():
            self.channel_combo.setCurrentIndex(0)
        self.channel_combo.currentTextChanged.connect(self._on_channel_changed)
        self.fit_btn.clicked.connect(self._on_fit_clicked)
        self.copy_btn.clicked.connect(self._copy_channel_to_clipboard)
        self._last_fit_result = None
        self._initialize_curve_entries()
        if self.channel_combo.count():
            self._plot_selected_channel()
        else:
            self.ax.text(0.5, 0.5, "No channels", ha='center', va='center', transform=self.ax.transAxes)
            self.canvas.draw()
        self._update_fit_button()

    def _channel_label_with_unit(self, name):
        base = name or ""
        unit = self.spec.get('unit_map', {}).get(name)
        if not unit:
            unit = guess_channel_unit(name)
        if not unit and '(' in base and base.endswith(')'):
            return base
        if unit:
            return f"{base} ({unit})"
        return base

    def _axis_display_name(self, axis):
        label = axis.get("label") or "Axis"
        unit = axis.get("unit") or ""
        if unit:
            if unit.lower() == "v":
                return f"{label} (mV)" if "mV" not in label else label
            return f"{label} ({unit})" if unit not in label else label
        return label

    def _collect_axes(self, spec):
        axes = []
        for ax in (spec.get("AxisChoices") or []):
            vals = np.asarray(ax.get("values", []), dtype=float)
            label = ax.get("label") or "Axis"
            unit = ax.get("unit") or ""
            key = ax.get("key") or label
            axes.append({"key": key, "label": label, "unit": unit, "values": vals})
        # Deduplicate identical axes (same values) to avoid duplicate Bias entries
        if axes:
            deduped = []
            seen_vals = []
            for ax in axes:
                vals = np.asarray(ax.get("values", []), dtype=float)
                if any(np.array_equal(vals, sv) for sv in seen_vals):
                    continue
                seen_vals.append(vals)
                deduped.append(ax)
            if deduped:
                return deduped
        primary = {
            "key": "primary",
            "label": spec.get('AxisLabel') or "Bias",
            "unit": spec.get('AxisUnit') or ("V" if "bias" in str(spec.get('AxisLabel') or "").lower() else ""),
            "values": np.asarray(spec.get('V', []), dtype=float),
        }
        axes.append(primary)
        alt_vals = spec.get('AltAxis')
        if alt_vals is not None:
            axes.append(
                {
                    "key": "alt",
                    "label": spec.get('AltAxisLabel') or "Z rel",
                    "unit": spec.get('AltAxisUnit') or "nm",
                    "values": np.asarray(alt_vals, dtype=float),
                }
            )
        return axes

    def _on_axis_changed(self, idx):
        key = self.axis_combo.currentData()
        selected = None
        for ax in self.axes:
            if ax.get("key") == key:
                selected = ax
                break
        if selected is None and self.axes:
            selected = self.axes[0]
        if selected is None:
            self.V = np.asarray([])
            self.axis_label = "Axis"
            self.axis_unit = ""
            return
        self._last_fit_result = None
        self.fit_result_label.setText("")
        self.V = np.asarray(selected.get("values", []), dtype=float)
        self.axis_label = selected.get("label") or "Axis"
        unit = (selected.get("unit") or "").strip()
        if unit.lower() == "v":
            try:
                max_abs = float(np.nanmax(np.abs(self.V))) if np.isfinite(self.V).any() else 0.0
                if max_abs > 5.0:
                    unit = "mV"  # mislabeled mV data; keep values as-is but relabel to avoid extra scaling
            except Exception:
                pass
        self.axis_unit = unit or ""
        self._update_primary_axis(axis_vals=self.V, axis_label=self.axis_label, axis_unit=self.axis_unit)
        self._plot_selected_channel()
        self._update_fit_button()

    def _on_channel_changed(self, name):
        self._last_fit_result = None
        self.fit_result_label.setText("")
        if self._curve_entries:
            channel = name or ""
            values = np.asarray(self.channels.get(channel), dtype=float) if channel else np.asarray([], dtype=float)
            entry = self._curve_entries[0]
            entry["values"] = values
            entry["channel"] = channel
            entry["label"] = f"{Path(self.spec.get('path','')).name} ({channel})" if channel else Path(self.spec.get('path','')).name
            self._update_curve_list()
        self._plot_selected_channel()
        self._update_fit_button()

    def _update_primary_axis(self, axis_vals, axis_label, axis_unit):
        if not self._curve_entries:
            return
        entry = self._curve_entries[0]
        entry["axis_vals"] = np.asarray(axis_vals, dtype=float)
        entry["axis_label"] = axis_label
        entry["axis_unit"] = axis_unit

    def _plot_selected_channel(self):
        self.ax.clear()
        if not self._curve_entries:
            self.canvas.draw_idle()
            return
        axis_label = self.axis_label or "Axis"
        axis_unit = (self.axis_unit or "").strip()
        axis_plot_scale = 1.0
        axis_plot_unit = axis_unit
        if axis_unit.lower() == "v" and self.V.size:
            axis_plot_scale = 1000.0
            axis_plot_unit = "mV"
        if axis_unit and axis_unit not in axis_label:
            axis_label = f"{axis_label} ({axis_unit})"
        plotted = False
        for entry in self._curve_entries:
            axis_vals = np.asarray(entry.get("axis_vals", []), dtype=float)
            values = np.asarray(entry.get("values", []), dtype=float)
            if axis_vals.size == 0 or values.size == 0:
                continue
            scaled_axis = axis_vals * axis_plot_scale
            self.ax.plot(scaled_axis, values, color=entry.get("color", '#c94cfa'),
                         lw=1.5, label=entry.get("label", "Data"))
            plotted = True
        self._axis_plot_scale = axis_plot_scale
        self._axis_plot_unit = axis_plot_unit
        self.ax.set_xlabel(axis_label)
        name = self.channel_combo.currentText()
        self.ax.set_ylabel(self._channel_label_with_unit(name))
        self.ax.grid(True, alpha=0.2)
        if plotted:
            self.ax.legend()
        if self._last_fit_result and self._last_fit_result.get('channel') == name:
            self._draw_fit_overlay(self._last_fit_result)
        self._apply_font_scale()
        self.canvas.draw_idle()

    def _on_canvas_context_menu(self, pos):
        menu = QtWidgets.QMenu(self)
        copy_act = menu.addAction("Copy channel data")
        action = menu.exec_(self.canvas.mapToGlobal(pos))
        if action == copy_act:
            self._copy_channel_to_clipboard()

    def _copy_channel_to_clipboard(self):
        name = self.channel_combo.currentText()
        if not name or name not in self.channels or not self.V.size:
            QtWidgets.QMessageBox.information(self, "Copy spectroscopy", "No spectroscopy data to copy.")
            return
        bias = self.V
        scale = getattr(self, "_axis_plot_scale", 1.0) or 1.0
        unit = getattr(self, "_axis_plot_unit", self.axis_unit) or ""
        bias_vals = bias * scale
        values = self.channels[name]
        spec_path = Path(self.spec.get('path', ''))
        file_name = spec_path.name or 'unknown'
        folder_name = spec_path.parent.name if spec_path.parent != spec_path else ''
        pos = (self.spec.get('x'), self.spec.get('y'))
        time_str = self.spec.get('time')
        lines = [
            f"File\t{file_name}",
            f"Channel\t{name}",
            f"Position (nm)\t{pos[0] if pos[0] is not None else '?'}\t{pos[1] if pos[1] is not None else '?'}",
            f"Folder\t{folder_name}",
            f"Acquired\t{time_str}",
            "",
            f"Bias ({unit or 'arb'})\t{self._channel_label_with_unit(name)}"
        ]
        for v, val in zip(bias_vals, values):
            try:
                lines.append(f"{float(v):.9g}\t{float(val):.9g}")
            except Exception:
                lines.append(f"{v}\t{val}")
        QtWidgets.QApplication.clipboard().setText("\n".join(lines))
        QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), "Spectroscopy copied", self)

    def _initialize_curve_entries(self):
        channel = self.channel_combo.currentText()
        axis_vals = np.asarray(self.axes[0]["values"], dtype=float) if self.axes else np.asarray([], dtype=float)
        values = np.asarray(self.channels.get(channel), dtype=float) if channel else np.asarray([], dtype=float)
        entry = {
            "label": f"{Path(self.spec.get('path','')).name} ({channel})" if channel else Path(self.spec.get('path','')).name,
            "axis_vals": axis_vals,
            "values": values,
            "color": self._active_line_color,
            "spec_path": str(Path(self.spec.get('path',''))),
            "channel": channel,
            "axis_label": self.axis_label,
            "axis_unit": self.axis_unit,
        }
        self._curve_entries = [entry]
        self._selected_curve_index = 0
        self._update_curve_list()

    def _update_curve_list(self):
        if not hasattr(self, "curve_list"):
            return
        current = max(0, min(self._selected_curve_index, len(self._curve_entries) - 1 if self._curve_entries else 0))
        self.curve_list.blockSignals(True)
        self.curve_list.clear()
        for entry in self._curve_entries:
            self.curve_list.addItem(entry.get("label", ""))
        self.curve_list.blockSignals(False)
        if self.curve_list.count():
            self.curve_list.setCurrentRow(current)
        self._selected_curve_index = current

    def _on_curve_selection_changed(self, row):
        if row < 0:
            return
        self._selected_curve_index = row

    def _current_entry(self):
        if not self._curve_entries:
            return None
        idx = self._selected_curve_index
        if idx < 0 or idx >= len(self._curve_entries):
            idx = 0
        return self._curve_entries[idx]

    def _apply_font_scale(self):
        scale = getattr(self, "_font_scale", 1.0)
        self.ax.tick_params(labelsize=8 * scale)
        self.ax.xaxis.label.set_fontsize(10 * scale)
        self.ax.yaxis.label.set_fontsize(10 * scale)
        legend = self.ax.get_legend()
        if legend:
            for text in legend.get_texts():
                text.set_fontsize(8 * scale)
        for widget, base in ((self.meta_label, 9.0), (self.fit_result_label, 8.5)):
            font = widget.font()
            font.setPointSizeF(base * scale)
            widget.setFont(font)

    def eventFilter(self, source, event):
        if source == self.canvas:
            if event.type() == QtCore.QEvent.MouseButtonPress and event.button() == QtCore.Qt.LeftButton:
                self._drag_start_pos = event.pos()
            elif event.type() == QtCore.QEvent.MouseMove and self._drag_start_pos is not None:
                if (event.pos() - self._drag_start_pos).manhattanLength() >= QtWidgets.QApplication.startDragDistance():
                    self._start_drag()
                    self._drag_start_pos = None
            elif event.type() == QtCore.QEvent.MouseButtonRelease:
                self._drag_start_pos = None
        return super().eventFilter(source, event)

    def dragEnterEvent(self, event):
        if event.mimeData().hasFormat("application/x-sxm-spectroscopy"):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        data = event.mimeData().data("application/x-sxm-spectroscopy")
        try:
            payload = json.loads(bytes(data).decode("utf-8"))
        except Exception:
            event.ignore()
            return
        self._add_entry_from_drop(payload)
        event.acceptProposedAction()

    def _add_entry_from_drop(self, payload):
        axis_vals = np.asarray(payload.get("axis_vals") or [], dtype=float)
        values = np.asarray(payload.get("values") or [], dtype=float)
        color = payload.get("color") or self.SCIENCE_PALETTE[len(self._curve_entries) % len(self.SCIENCE_PALETTE)]
        label = payload.get("label") or Path(payload.get("spec_path", "")).name
        entry = {
            "label": label,
            "axis_vals": axis_vals,
            "values": values,
            "color": color,
            "spec_path": payload.get("spec_path", ""),
            "channel": payload.get("channel"),
            "axis_label": payload.get("axis_label", self.axis_label),
            "axis_unit": payload.get("axis_unit", self.axis_unit),
        }
        self._curve_entries.append(entry)
        self._selected_curve_index = len(self._curve_entries) - 1
        self._update_curve_list()
        self._plot_selected_channel()

    def _start_drag(self):
        entry = self._current_entry()
        if not entry:
            return
        drag = QtGui.QDrag(self)
        mime = QtCore.QMimeData()
        payload = {
            "label": entry.get("label"),
            "spec_path": entry.get("spec_path"),
            "axis_vals": entry.get("axis_vals").tolist() if isinstance(entry.get("axis_vals"), np.ndarray) else list(entry.get("axis_vals") or []),
            "values": entry.get("values").tolist() if isinstance(entry.get("values"), np.ndarray) else list(entry.get("values") or []),
            "color": entry.get("color"),
            "channel": entry.get("channel"),
            "axis_label": entry.get("axis_label"),
            "axis_unit": entry.get("axis_unit"),
        }
        mime.setData("application/x-sxm-spectroscopy", json.dumps(payload).encode("utf-8"))
        drag.setMimeData(mime)
        pixmap = QtGui.QPixmap(32, 32)
        pixmap.fill(QtGui.QColor(entry.get("color", "#000000")))
        drag.setPixmap(pixmap)
        drag.setHotSpot(QtCore.QPoint(16, 16))
        drag.exec_(QtCore.Qt.CopyAction)

    def _create_palette_swatch_widget(self):
        swatch_widget = QtWidgets.QWidget()
        outer_layout = QtWidgets.QHBoxLayout(swatch_widget)
        outer_layout.setContentsMargins(0, 8, 0, 0)
        outer_layout.setSpacing(6)
        label = QtWidgets.QLabel("Color strip:")
        label.setFixedWidth(90)
        outer_layout.addWidget(label, alignment=QtCore.Qt.AlignTop)
        grid_widget = QtWidgets.QWidget()
        grid_layout = QtWidgets.QGridLayout(grid_widget)
        grid_layout.setSpacing(3)
        grid_layout.setContentsMargins(0, 0, 0, 0)
        rows = 2
        swatches_per_row = (len(self.SCIENCE_PALETTE) + rows - 1) // rows
        for idx, color in enumerate(self.SCIENCE_PALETTE):
            row = idx // swatches_per_row
            col = idx % swatches_per_row
            button = QtWidgets.QPushButton()
            button.setFixedSize(24, 24)
            button.setFlat(True)
            base_style = (
                f"background-color:{color}; border:1px solid #aaa; border-radius:3px;"
            )
            button.setProperty("baseStyle", base_style)
            button.setStyleSheet(base_style)
            button.clicked.connect(functools.partial(self._on_swatch_clicked, color, button))
            button.setAccessibleDescription(f"Select color {idx+1}")
            grid_layout.addWidget(button, row, col)
            self._swatch_buttons.append(button)
        outer_layout.addWidget(grid_widget, 1)
        swatch_widget.setAccessibleName("Color cycle swatches")
        swatch_widget.setAccessibleDescription("Displays available colors for the single spectrum plot")
        if self._swatch_buttons:
            self._set_active_swatch(self._swatch_buttons[0])
        return swatch_widget

    def _set_active_swatch(self, button):
        for btn in self._swatch_buttons:
            base = btn.property("baseStyle") or ""
            btn.setStyleSheet(base)
        if button:
            base = button.property("baseStyle") or ""
            button.setStyleSheet(f"{base} border:2px solid #333;")

    def _on_swatch_clicked(self, color, button):
        self._active_line_color = color
        self._set_active_swatch(button)
        entry = self._current_entry()
        if entry:
            entry["color"] = color
        self._plot_selected_channel()

    def _draw_fit_overlay(self, res):
        if not self.V.size:
            return
        scale = getattr(self, "_axis_plot_scale", 1.0) or 1.0
        x_dense = np.linspace(np.nanmin(self.V), np.nanmax(self.V), 400)
        y_dense = res['func'](x_dense)
        self.ax.plot(x_dense * scale, y_dense, '--', color='#ff8c00', lw=1.5, label='Fit')
        v0 = res.get('v0')
        v0_err = res.get('v0_err')
        if v0 is not None and np.isfinite(v0):
            y0 = res['func'](v0)
            x_plot = v0 * scale
            xerr = v0_err * scale if v0_err is not None else None
            self.ax.errorbar([x_plot], [y0], xerr=[xerr] if xerr is not None else None,
                             fmt='o', color='#004c99', ecolor='#004c99', capsize=4, label='LCPD')
        self.ax.legend()
        axis_unit = getattr(self, "_axis_plot_unit", self.axis_unit) or ""
        v0_txt = ""
        if v0 is not None and np.isfinite(v0):
            v_disp = v0 * scale
            v_err_disp = (res.get('v0_err') or 0.0) * scale
            unit_txt = axis_unit or ("mV" if scale == 1000.0 else "V")
            v0_txt = f"LCPD = {v_disp:.3g} {unit_txt}"
            if v_err_disp:
                v0_txt += f" +/- {v_err_disp:.3g}"
        text = (
            f"a = {res['a']:.4g} +/- {res['a_err']:.2g}\n"
            f"b = {res['b']:.4g} +/- {res['b_err']:.2g}\n"
            f"c = {res['c']:.4g} +/- {res['c_err']:.2g} Hz\n"
            f"{v0_txt}\n"
            f"RMSE = {res['rmse']:.4g}"
        )
        self.fit_result_label.setText(text)

    def _on_fit_clicked(self):
        name = self.channel_combo.currentText()
        if not name or name not in self.channels:
            return
        try:
            res = fit_parabola_bias(self.V, self.channels[name])
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Fit failed", str(e))
            return
        res['channel'] = name
        a = res.get('a'); b = res.get('b')
        v0 = None; v0_err = None
        try:
            if a is not None and b is not None and np.isfinite(a) and np.isfinite(b) and a != 0:
                v0 = -b / (2.0 * a)
                da = res.get('a_err', 0.0)
                db = res.get('b_err', 0.0)
                term1 = (db / (2.0 * a)) ** 2 if a != 0 else 0.0
                term2 = ((b * da) / (2.0 * (a ** 2))) ** 2 if a != 0 else 0.0
                v0_err = math.sqrt(max(term1 + term2, 0.0))
        except Exception:
            v0 = None; v0_err = None
        res['v0'] = v0
        res['v0_err'] = v0_err
        self._last_fit_result = res
        self._plot_selected_channel()

    def _update_fit_button(self):
        enable = bool(self.channel_combo.count() and self.V.size)
        self.fit_btn.setEnabled(enable)

class MatrixSpectroViewer(QtWidgets.QDialog):
    MARKER_STYLE_OPTIONS = [
        ("Circle", "o"),
        ("Square", "s"),
        ("Diamond", "D"),
        ("Triangle", "^"),
        ("Cross", "X"),
    ]
    MARKER_SIZE_PRESETS = [16, 28, 42]
    def __init__(self, parent, image_entry, specs, dataset=None, palette_name=None):
        super().__init__(parent)
        self.image_entry = image_entry
        self.specs = list(specs)
        self.viewer = parent
        self.dataset = dataset
        self.anchor_path = str(image_entry.get('path') or "")
        if self.anchor_path:
            try:
                self.image_entry['path'] = Path(self.anchor_path)
            except Exception:
                pass
        self._resolve_anchor_path()
        base_name = self._matrix_file_name()
        self.setWindowTitle(f"Matrix Explorer - {base_name}")
        self.resize(1100, 720)
        root = QtWidgets.QVBoxLayout(self)

        splitter = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        root.addWidget(splitter, 1)

        left_panel = QtWidgets.QWidget()
        left_layout = QtWidgets.QVBoxLayout(left_panel)

        self.canvas = FigureCanvas(Figure(figsize=(6,6)))
        self.ax = self.canvas.figure.add_subplot(111)
        left_layout.addWidget(self.canvas, 1)

        self.image_value_label = QtWidgets.QLabel("Value: --")
        left_layout.addWidget(self.image_value_label)

        controls = QtWidgets.QHBoxLayout()
        controls.addWidget(QtWidgets.QLabel("Channel map:"))
        self.channel_combo = QtWidgets.QComboBox()
        controls.addWidget(self.channel_combo, 1)
        self.map_mode_combo = QtWidgets.QComboBox()
        self.map_mode_combo.addItems(["Max amplitude", "Peak position", "Integral"])
        controls.addWidget(self.map_mode_combo)
        left_layout.addLayout(controls)

        ref_controls = QtWidgets.QHBoxLayout()
        ref_controls.addWidget(QtWidgets.QLabel("Reference image:"))
        self.image_channel_combo = QtWidgets.QComboBox()
        ref_controls.addWidget(self.image_channel_combo, 1)
        left_layout.addLayout(ref_controls)

        palette_controls = QtWidgets.QHBoxLayout()
        palette_controls.addWidget(QtWidgets.QLabel("Color cycle:"))
        self.palette_combo = QtWidgets.QComboBox()
        for name in list_color_cycles():
            self.palette_combo.addItem(name)
        palette_controls.addWidget(self.palette_combo, 1)
        left_layout.addLayout(palette_controls)

        self.show_positions_cb = QtWidgets.QCheckBox("Show all spectroscopy positions")
        self.show_positions_cb.setChecked(True)
        self.show_positions_cb.toggled.connect(self._draw_image_layer)
        left_layout.addWidget(self.show_positions_cb)

        self.fit_matrix_btn = QtWidgets.QPushButton("Fit matrix parabolas...")
        left_layout.addWidget(self.fit_matrix_btn)
        self.reset_view_btn = QtWidgets.QPushButton("Reset view")
        left_layout.addWidget(self.reset_view_btn)
        self.matrix_info_label = QtWidgets.QLabel("")
        self.matrix_info_label.setWordWrap(True)
        left_layout.addWidget(self.matrix_info_label)

        splitter.addWidget(left_panel)

        right_panel = QtWidgets.QWidget()
        right_layout = QtWidgets.QVBoxLayout(right_panel)

        self.curve_canvas = FigureCanvas(Figure(figsize=(4,4)))
        self.curve_ax = self.curve_canvas.figure.add_subplot(111)
        right_layout.addWidget(self.curve_canvas, 3)

        self.selection_table = QtWidgets.QTableWidget(0, 3)
        self.selection_table.setHorizontalHeaderLabels(["Channel", "X (nm)", "Y (nm)"])
        self.selection_table.horizontalHeader().setStretchLastSection(True)
        right_layout.addWidget(self.selection_table, 2)

        export_row = QtWidgets.QHBoxLayout()
        self.export_csv_btn = QtWidgets.QPushButton("Export selection to CSV")
        export_row.addWidget(self.export_csv_btn)
        self.clear_selection_btn = QtWidgets.QPushButton("Clear selection")
        export_row.addWidget(self.clear_selection_btn)
        export_row.addStretch(1)
        right_layout.addLayout(export_row)

        splitter.addWidget(right_panel)
        splitter.setStretchFactor(0, 2)
        splitter.setStretchFactor(1, 1)

        self._channel_specs = self._group_specs_by_channel()

        self.canvas.mpl_connect("button_press_event", self._on_click)
        self.canvas.mpl_connect("motion_notify_event", self._on_canvas_hover)
        self.canvas.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.canvas.customContextMenuRequested.connect(self._on_canvas_context_menu)
        self._fit_dialogs = []
        self._current_image_arr = None
        self._current_image_extent = None
        self._current_image_unit = ''
        self._selection = []
        self._selection_keys = set()
        self._selection_artists = []
        self._position_marker_config = {
            "marker": "o",
            "size": 28,
            "facecolor": "#ffffff",
            "edgecolor": "#101010",
            "linewidth": 0.4,
            "alpha": 0.85,
        }
        self._aggregate_mode = False
        self._focused_key = None
        self.palette_name = palette_name or getattr(self.viewer, "spectro_color_cycle", DEFAULT_COLOR_CYCLE)
        self._color_palette = get_color_cycle(self.palette_name)
        if not self._color_palette:
            self._color_palette = ["#4c78a8"]
        self._color_index = 0

        self._populate_channels()
        self._populate_image_channels()
        self.channel_combo.currentIndexChanged.connect(self._on_channel_combo_changed)
        self.map_mode_combo.currentIndexChanged.connect(self._draw_image_layer)
        self.image_channel_combo.currentIndexChanged.connect(self._draw_image_layer)
        self.fit_matrix_btn.clicked.connect(self._on_fit_matrix)
        self.reset_view_btn.clicked.connect(self._reset_matrix_view)
        self.export_csv_btn.clicked.connect(self._on_export_selection)
        self.clear_selection_btn.clicked.connect(self._clear_selection)
        idx = self.palette_combo.findText(self.palette_name)
        self.palette_combo.blockSignals(True)
        if idx >= 0:
            self.palette_combo.setCurrentIndex(idx)
        else:
            self.palette_combo.setCurrentIndex(0)
            self.palette_name = self.palette_combo.currentText()
            self._color_palette = get_color_cycle(self.palette_name)
        self.palette_combo.blockSignals(False)
        self.selection_table.itemSelectionChanged.connect(self._update_curve_from_selection)
        self.palette_combo.currentTextChanged.connect(self._on_palette_changed)
        self._draw_image_layer()
        self._update_matrix_info_label()

    def _group_specs_by_channel(self):
        mapping = defaultdict(list)
        self._channel_labels_map = {}
        for spec in self.specs:
            path = spec.get('path')
            if not path:
                continue
            key = self._normalize_path(path)
            mapping[key].append(spec)
            if key not in self._channel_labels_map:
                label = spec.get('channel_name') or spec.get('channel_code')
                if not label:
                    chs = spec.get('channels') or {}
                    if len(chs) == 1:
                        label = next(iter(chs.keys()))
                self._channel_labels_map[key] = label or Path(key).name
        return mapping

    def _reset_color_cycle(self):
        self._color_index = 0

    def _next_color(self):
        if not self._color_palette:
            self._color_palette = ["#4c78a8"]
        color = self._color_palette[self._color_index % len(self._color_palette)]
        self._color_index += 1
        return color

    def _selection_key(self, spec):
        return (
            str(spec.get('path')),
            spec.get('matrix_index'),
            spec.get('channel_name') or spec.get('channel_code'),
        )

    def _variant_color(self, base_color, factor=0.35):
        rgb = np.array(mcolors.to_rgb(base_color))
        factor = min(max(factor, 0.0), 1.0)
        adjusted = rgb + (1.0 - rgb) * factor
        return mcolors.to_hex(np.clip(adjusted, 0.0, 1.0))

    def _event_modifiers(self, event):
        qevent = getattr(event, "guiEvent", None)
        if qevent is None:
            return QtCore.Qt.NoModifier
        try:
            return qevent.modifiers()
        except Exception:
            return QtCore.Qt.NoModifier

    def _channel_unit_for_spec(self, spec, channel_label):
        unit_map = spec.get('unit_map') or {}
        if channel_label and channel_label in unit_map and unit_map[channel_label]:
            return unit_map[channel_label]
        if unit_map:
            for key, val in unit_map.items():
                if val:
                    return val
        return guess_channel_unit(channel_label)

    def _extract_channel_data(self, spec, channel_label):
        channels = spec.get('channels') or {}
        ys = None
        label = channel_label
        if label in channels:
            ys = np.asarray(channels[label], dtype=float)
        elif channels:
            label, values = next(iter(channels.items()))
            ys = np.asarray(values, dtype=float)
        elif spec.get('data'):
            data = spec.get('data')
            try:
                xs = np.asarray(data[0], dtype=float)
                ys = np.asarray(data[1], dtype=float)
                unit = self._channel_unit_for_spec(spec, label)
                x_unit = spec.get("AxisUnit") or ""
                return xs, ys, unit, label, x_unit
            except Exception:
                return None, None, None, label, ""
        xs = np.asarray(spec.get('V', []), dtype=float)
        if xs.size == 0 or ys is None or ys.size == 0:
            data = spec.get('data')
            if data:
                try:
                    xs = np.asarray(data[0], dtype=float)
                    ys = np.asarray(data[1], dtype=float)
                except Exception:
                    return None, None, None, label, ""
        if xs.size == 0 or ys is None or ys.size == 0:
            return None, None, None, label, ""
        unit = self._channel_unit_for_spec(spec, label)
        x_unit = spec.get("AxisUnit") or ""
        return xs, ys, unit, label, x_unit

    def _remove_selection_entry(self, key):
        self._selection = [entry for entry in self._selection if entry.get("key") != key]
        self._selection_keys.discard(key)
        if self._selection:
            self._focused_key = self._selection[-1].get("key")
        else:
            self._focused_key = None
            self._aggregate_mode = False

    def _update_selection_markers(self, redraw=True):
        for artist in getattr(self, "_selection_artists", []):
            try:
                artist.remove()
            except Exception:
                pass
        self._selection_artists = []
        if not self._selection:
            if redraw:
                self.canvas.draw_idle()
            return
        for entry in self._selection:
            coords = entry.get("coords")
            if not coords:
                continue
            size = 110 if entry.get("key") == self._focused_key else 70
            face = entry.get("color", "#4c78a8")
            edge = "#101010"
            artist = self.ax.scatter(
                [coords[0]],
                [coords[1]],
                s=size,
                facecolors=face,
                edgecolors=edge,
                linewidths=1.0,
                alpha=0.95,
                zorder=5,
            )
            self._selection_artists.append(artist)
        if redraw:
            self.canvas.draw_idle()

    def _populate_channels(self):
        self.channel_combo.clear()
        added = set()
        if self.dataset and self.dataset.channels:
            for ch in self.dataset.channels:
                path = self._normalize_path(ch.get('path', ch.get('filename')))
                if path not in self._channel_specs or path in added:
                    continue
                label = ch.get('label') or self._channel_labels_map.get(path) or Path(path).name
                self.channel_combo.addItem(label, path)
                self._channel_labels_map[path] = label
                added.add(path)
        for path in sorted(self._channel_specs.keys()):
            if path in added:
                continue
            label = self._channel_labels_map.get(path, Path(path).name)
            self.channel_combo.addItem(label, path)
            self._channel_labels_map[path] = label
            added.add(path)
        if self.channel_combo.count():
            self.channel_combo.setCurrentIndex(0)

    def _populate_image_channels(self):
        anchor = self.anchor_path or self.image_entry.get('path')
        path = Path(anchor) if anchor else None
        header, fds = self.viewer.headers.get(str(path), (None, None)) if path else (None, None)
        self.image_channel_combo.blockSignals(True)
        self.image_channel_combo.clear()
        if not fds:
            self.image_channel_combo.addItem("No image", -1)
            self.image_channel_combo.setEnabled(False)
        else:
            self.image_channel_combo.setEnabled(True)
            for idx, fd in enumerate(fds):
                label = fd.get('Caption', fd.get('FileName', f"Channel {idx}"))
                self.image_channel_combo.addItem(label, idx)
            default_idx = 0
            if self.viewer.last_preview and self.viewer.last_preview[0] == str(path):
                try:
                    prev_idx = int(self.viewer.last_preview[1])
                except Exception:
                    prev_idx = 0
                if 0 <= prev_idx < len(fds):
                    default_idx = prev_idx
            self.image_channel_combo.setCurrentIndex(default_idx)
        self.image_channel_combo.blockSignals(False)

    def _matrix_file_name(self):
        if self.dataset and getattr(self.dataset, "channels", None):
            first = next((ch for ch in self.dataset.channels if ch.get('filename') or ch.get('path')), None)
            if first:
                name = first.get('filename') or first.get('path')
                if name:
                    return Path(name).name
        if self.specs:
            name = self.specs[0].get('path')
            if name:
                return Path(name).name
        return "matrix"

    def _resolve_anchor_path(self):
        headers = getattr(self.viewer, 'headers', {})
        if self.anchor_path and str(self.anchor_path) in headers:
            return
        anchor = next((spec.get('image_key') for spec in self.specs if spec.get('image_key')), None)
        if anchor:
            self.anchor_path = str(anchor)
            try:
                self.image_entry['path'] = Path(self.anchor_path)
            except Exception:
                pass

    def _update_matrix_info_label(self):
        matrix_name = self._matrix_file_name()
        total = len(self.specs)
        rows = max((spec.get('grid_rows') or 0) for spec in self.specs) if self.specs else 0
        cols = max((spec.get('grid_cols') or 0) for spec in self.specs) if self.specs else 0
        xs = [float(spec.get('x')) for spec in self.specs if spec.get('x') is not None]
        ys = [float(spec.get('y')) for spec in self.specs if spec.get('y') is not None]
        x_txt = "n/a"
        y_txt = "n/a"
        if xs:
            xmin, xmax = min(xs), max(xs)
            x_txt = f"{xmin:.2f}→{xmax:.2f} nm (Δ {xmax - xmin:.2f} nm)"
        if ys:
            ymin, ymax = min(ys), max(ys)
            y_txt = f"{ymin:.2f}→{ymax:.2f} nm (Δ {ymax - ymin:.2f} nm)"
        times = [spec.get('time') for spec in self.specs if isinstance(spec.get('time'), datetime)]
        time_txt = "n/a"
        if times:
            start = min(times)
            end = max(times)
            if start and end:
                time_txt = f"{start:%Y-%m-%d %H:%M:%S}"
                if end != start:
                    try:
                        seconds = abs((end - start).total_seconds())
                    except Exception:
                        seconds = 0.0
                    time_txt += f" → {end:%H:%M:%S} (Δ {seconds:.1f}s)"
        info = (
            f"<b>{matrix_name}</b><br>"
            f"Points: {total} ({rows}×{cols})<br>"
            f"X range: {x_txt}<br>"
            f"Y range: {y_txt}<br>"
            f"Acquired: {time_txt}"
        )
        self.matrix_info_label.setText(info)

    def _on_palette_changed(self):
        self.palette_name = self.palette_combo.currentText() or DEFAULT_COLOR_CYCLE
        self._color_palette = get_color_cycle(self.palette_name)
        self._reset_color_cycle()
        if hasattr(self.viewer, "set_spectro_color_cycle"):
            self.viewer.set_spectro_color_cycle(self.palette_name)
        if self._selection:
            self._apply_palette_to_selection()
            self._refresh_selection_table()
        else:
            self._update_selection_markers()

    def _apply_palette_to_selection(self):
        self._reset_color_cycle()
        for entry in self._selection:
            entry["color"] = self._next_color()

    def _reset_matrix_view(self):
        self._clear_selection()
        if self.channel_combo.count():
            self.channel_combo.blockSignals(True)
            self.channel_combo.setCurrentIndex(0)
            self.channel_combo.blockSignals(False)
        if self.map_mode_combo.count():
            self.map_mode_combo.setCurrentIndex(0)
        if self.image_channel_combo.count():
            self.image_channel_combo.setCurrentIndex(0)
        target = getattr(self.viewer, "spectro_color_cycle", DEFAULT_COLOR_CYCLE)
        self.palette_combo.blockSignals(True)
        idx = self.palette_combo.findText(target)
        if idx < 0:
            idx = 0
        self.palette_combo.setCurrentIndex(idx)
        self.palette_combo.blockSignals(False)
        self._on_palette_changed()
        self._draw_image_layer()

    def _on_channel_combo_changed(self):
        self._clear_selection()
        self._draw_image_layer()

    def _on_canvas_context_menu(self, pos):
        menu = QtWidgets.QMenu(self)

        style_menu = menu.addMenu("Marker style")
        style_group = QtWidgets.QActionGroup(menu)
        current_marker = self._position_marker_config.get("marker", "o")
        for label, marker in self.MARKER_STYLE_OPTIONS:
            act = style_menu.addAction(label)
            act.setCheckable(True)
            act.setChecked(current_marker == marker)
            act.triggered.connect(functools.partial(self._set_position_marker_style, marker))
            style_group.addAction(act)

        size_menu = menu.addMenu("Marker size")
        size_group = QtWidgets.QActionGroup(menu)
        current_size = self._position_marker_config.get("size", 28)
        for size in self.MARKER_SIZE_PRESETS:
            act = size_menu.addAction(f"{size} pt")
            act.setCheckable(True)
            act.setChecked(current_size == size)
            act.triggered.connect(functools.partial(self._set_position_marker_size, size))
            size_group.addAction(act)
        custom_size = size_menu.addAction("Custom...")
        custom_size.triggered.connect(self._choose_custom_marker_size)

        fill_act = menu.addAction("Marker fill color...")
        fill_act.triggered.connect(functools.partial(self._choose_position_marker_color, "facecolor"))
        edge_act = menu.addAction("Marker edge color...")
        edge_act.triggered.connect(functools.partial(self._choose_position_marker_color, "edgecolor"))

        menu.addSeparator()
        clear_act = menu.addAction("Clear selections")
        reset_act = menu.addAction("Reset view")
        action = menu.exec_(self.canvas.mapToGlobal(pos))
        if action == clear_act:
            self._clear_selection()
        elif action == reset_act:
            self._reset_matrix_view()

    def _set_position_marker_style(self, marker):
        if not marker:
            return
        if self._position_marker_config.get("marker") == marker:
            return
        self._position_marker_config["marker"] = marker
        self._draw_image_layer()

    def _set_position_marker_size(self, size):
        if size <= 0:
            return
        if self._position_marker_config.get("size") == size:
            return
        self._position_marker_config["size"] = size
        self._draw_image_layer()

    def _choose_custom_marker_size(self):
        current = int(self._position_marker_config.get("size", 28))
        size, ok = QtWidgets.QInputDialog.getInt(
            self, "Marker size", "Marker size (pts):", current, 6, 200, 1
        )
        if ok:
            self._set_position_marker_size(size)

    def _choose_position_marker_color(self, role):
        current = self._position_marker_config.get(role, "#ffffff")
        color = QtWidgets.QColorDialog.getColor(QtGui.QColor(current), self, "Select marker color")
        if not color.isValid():
            return
        self._position_marker_config[role] = color.name()
        self._draw_image_layer()

    def _draw_image_layer(self):
        anchor = self.anchor_path or self.image_entry.get('path')
        if not anchor:
            return
        path = Path(anchor)
        header, fds = self.viewer.headers.get(str(path), (None, None))
        header_map = header or {}
        channel_specs = self._current_channel_specs()
        self.ax.clear()
        self._selection_artists = []
        agg_mode = self.map_mode_combo.currentText()
        metric = None
        file_key = str(path)
        if agg_mode == "Max amplitude":
            metric = self._build_stat_metric(np.nanmax, channel_specs, header_map, file_key)
        elif agg_mode == "Integral":
            metric = self._build_integral_metric(channel_specs, header_map, file_key)
        elif agg_mode == "Peak position":
            metric = self._build_peak_metric(channel_specs, header_map, file_key)
        metric_valid = metric is not None and np.isfinite(metric).any()
        if metric_valid:
            self.ax.imshow(metric, cmap='inferno', origin='upper')
            self._current_image_arr = metric
            self._current_image_unit = ''
        elif header and fds:
            try:
                idx = self.image_channel_combo.currentData()
                if idx is None or idx < 0 or idx >= len(fds):
                    idx = 0
                fd = fds[idx]
                arr = self.viewer._get_channel_array(str(path), idx, header, fd)
                self.ax.imshow(arr, cmap='gray', origin='upper')
                self._current_image_arr = np.asarray(arr, dtype=float)
                self._current_image_unit = fd.get('PhysUnit', '')
            except Exception:
                self.ax.text(0.5, 0.5, Path(path).name, ha='center', va='center', transform=self.ax.transAxes)
                self._current_image_arr = None
        else:
            self.ax.text(0.5, 0.5, Path(path).name, ha='center', va='center', transform=self.ax.transAxes)
            self._current_image_arr = None
        xpix = int(header_map.get('xPixel', 128))
        ypix = int(header_map.get('yPixel', 128))
        xs = []
        ys = []
        if getattr(self.show_positions_cb, "isChecked", lambda: True)():
            overlay_specs = self.specs
        else:
            overlay_specs = channel_specs
        if overlay_specs:
            for spec in overlay_specs:
                coords = self.viewer._map_spec_to_pixels(spec, header_map, xpix, ypix, file_key=file_key)
                if coords:
                    xs.append(coords[0])
                    ys.append(coords[1])
            if xs and ys:
                cfg = self._position_marker_config
                self.ax.scatter(
                    xs,
                    ys,
                    s=cfg.get("size", 28),
                    marker=cfg.get("marker", "o"),
                    facecolors=cfg.get("facecolor", "#ffffff"),
                    edgecolors=cfg.get("edgecolor", "#101010"),
                    linewidths=cfg.get("linewidth", 0.4),
                    alpha=cfg.get("alpha", 0.85),
                    zorder=2,
                )
        self._update_selection_markers(redraw=False)
        self.canvas.draw_idle()
        if self._current_image_arr is None:
            self.image_value_label.setText("Value: --")

    def _current_channel_specs(self):
        path = self.channel_combo.currentData()
        if not path:
            return []
        return self._channel_specs.get(self._normalize_path(path), [])

    def _normalize_path(self, path):
        try:
            return str(Path(path))
        except Exception:
            return str(path)

    def _channel_label_for_path(self, path):
        key = self._normalize_path(path)
        label = self._channel_labels_map.get(key)
        if label:
            return label
        specs = self._channel_specs.get(key)
        if specs:
            sample = specs[0]
            label = sample.get('channel_name') or sample.get('channel_code')
            if not label:
                channels = sample.get('channels') or {}
                if len(channels) == 1:
                    label = next(iter(channels.keys()))
        return label or Path(key).name

    def _build_stat_metric(self, fn, channel_specs, header, file_key):
        if not channel_specs:
            return None
        xpix = int(header.get('xPixel', 128) if header else 128)
        ypix = int(header.get('yPixel', 128) if header else 128)
        grid = np.full((ypix, xpix), np.nan, dtype=float)
        for spec in channel_specs:
            data = spec.get('data')
            coords = self.viewer._map_spec_to_pixels(spec, header or {}, xpix, ypix, file_key=file_key)
            if data is None or coords is None:
                continue
            try:
                values = np.asarray(data[1], dtype=float)
                grid[coords[1], coords[0]] = fn(values)
            except Exception:
                continue
        return grid

    def _build_integral_metric(self, channel_specs, header, file_key):
        if not channel_specs:
            return None
        xpix = int(header.get('xPixel', 128) if header else 128)
        ypix = int(header.get('yPixel', 128) if header else 128)
        grid = np.full((ypix, xpix), np.nan, dtype=float)
        for spec in channel_specs:
            data = spec.get('data')
            coords = self.viewer._map_spec_to_pixels(spec, header or {}, xpix, ypix, file_key=file_key)
            if data is None or coords is None:
                continue
            try:
                xs = np.asarray(data[0], dtype=float)
                ys = np.asarray(data[1], dtype=float)
                grid[coords[1], coords[0]] = np.trapz(ys, xs)
            except Exception:
                continue
        return grid

    def _build_peak_metric(self, channel_specs, header, file_key):
        if not channel_specs:
            return None
        xpix = int(header.get('xPixel', 128) if header else 128)
        ypix = int(header.get('yPixel', 128) if header else 128)
        grid = np.full((ypix, xpix), np.nan, dtype=float)
        for spec in channel_specs:
            data = spec.get('data')
            coords = self.viewer._map_spec_to_pixels(spec, header or {}, xpix, ypix, file_key=file_key)
            if data is None or coords is None:
                continue
            try:
                ys = np.asarray(data[1], dtype=float)
                idx = int(np.nanargmax(ys))
                xs = np.asarray(data[0], dtype=float)
                grid[coords[1], coords[0]] = xs[idx]
            except Exception:
                continue
        return grid

    def _pick_spec_from_point(self, x, y, channel_specs, file_key):
        best = None
        best_dist = None
        header, _ = self.viewer.headers.get(str(self.image_entry['path']), (None, None))
        xpix = int(header.get('xPixel', 128) if header else 128)
        ypix = int(header.get('yPixel', 128) if header else 128)
        for spec in channel_specs:
            coords = self.viewer._map_spec_to_pixels(spec, header or {}, xpix, ypix, file_key=file_key)
            if coords is None:
                continue
            col, row = coords
            dist = (col - x)**2 + (row - y)**2
            if best is None or dist < best_dist:
                best = spec
                best_dist = dist
        return best

    def _on_click(self, event):
        if event.inaxes != self.ax or event.button != MouseButton.LEFT:
            return
        channel_specs = self._current_channel_specs()
        spec = self._pick_spec_from_point(event.xdata, event.ydata, channel_specs, str(self.image_entry['path']))
        if not spec:
            return
        header, _ = self.viewer.headers.get(str(self.image_entry['path']), (None, None))
        xpix = int(header.get('xPixel', 128) if header else 128)
        ypix = int(header.get('yPixel', 128) if header else 128)
        coords = self.viewer._map_spec_to_pixels(spec, header or {}, xpix, ypix, file_key=str(self.image_entry['path']))
        key = self._selection_key(spec)
        mods = self._event_modifiers(event)
        shift = bool(mods & QtCore.Qt.ShiftModifier)
        if shift and key in self._selection_keys:
            self._remove_selection_entry(key)
            if hasattr(self.viewer, '_toggle_multi_spec_selection'):
                self.viewer._toggle_multi_spec_selection(spec)
            self._refresh_selection_table()
            return
        if not shift:
            self._selection = []
            self._selection_keys = set()
            self._aggregate_mode = False
            self._focused_key = None
            self._reset_color_cycle()
            if hasattr(self.viewer, '_clear_multi_spec_selection'):
                self.viewer._clear_multi_spec_selection()
        primary_label = self._channel_label_for_path(self.channel_combo.currentData())
        multi = self._gather_multi_channel_specs(spec.get('matrix_index')) or [(primary_label, spec)]
        if primary_label:
            multi.sort(key=lambda item: 0 if item[0] == primary_label else 1)
        color = self._next_color()
        nm_coords = (spec.get('x'), spec.get('y'))
        entry = {
            "spec": spec,
            "coords": coords,
            "nm_coords": nm_coords,
            "multi": multi,
            "label": primary_label,
            "color": color,
            "key": key,
            "unit": self._channel_unit_for_spec(spec, primary_label),
        }
        self._selection.append(entry)
        self._selection_keys.add(key)
        self._focused_key = key
        if shift:
            self._aggregate_mode = True
            if hasattr(self.viewer, '_toggle_multi_spec_selection'):
                self.viewer._toggle_multi_spec_selection(spec)
        else:
            self.viewer._open_spectroscopy_popup(spec)
        max_sel = 24
        if len(self._selection) > max_sel:
            overflow = len(self._selection) - max_sel
            for stale in self._selection[:overflow]:
                self._selection_keys.discard(stale.get("key"))
            self._selection = self._selection[-max_sel:]
        self._refresh_selection_table()

    def _refresh_selection_table(self):
        self.selection_table.setRowCount(len(self._selection))
        for row, entry in enumerate(self._selection):
            label = entry.get("label", "Channel")
            color = QtGui.QColor(entry.get("color", "#4c78a8"))
            swatch = color.lighter(140)
            item = QtWidgets.QTableWidgetItem(label or "Channel")
            item.setData(QtCore.Qt.UserRole, entry.get("key"))
            item.setBackground(swatch)
            self.selection_table.setItem(row, 0, item)
            nm = entry.get("nm_coords") or (None, None)
            x_nm, y_nm = nm
            x_item = QtWidgets.QTableWidgetItem(f"{x_nm:.2f}" if x_nm is not None else "--")
            y_item = QtWidgets.QTableWidgetItem(f"{y_nm:.2f}" if y_nm is not None else "--")
            x_item.setBackground(swatch)
            y_item.setBackground(swatch)
            self.selection_table.setItem(row, 1, x_item)
            self.selection_table.setItem(row, 2, y_item)
        self.selection_table.scrollToBottom()
        if self._aggregate_mode:
            self._update_curve_plot()
        else:
            self._update_curve_plot(self._selection[-1] if self._selection else None)
        self._update_selection_markers()

    def _update_curve_from_selection(self):
        rows = self.selection_table.selectionModel().selectedRows()
        if not rows:
            return
        idx = rows[0].row()
        if 0 <= idx < len(self._selection):
            entry = self._selection[idx]
            self._focused_key = entry.get("key")
            if self._aggregate_mode:
                self._update_curve_plot()
            else:
                self._update_curve_plot(entry)
            self._update_selection_markers()

    def _update_curve_plot(self, entry=None):
        self.curve_ax.clear()
        entries = []
        if self._aggregate_mode:
            entries = list(self._selection)
        else:
            entry = entry or (self._selection[-1] if self._selection else None)
            if entry:
                entries = [entry]
                self._focused_key = entry.get("key")
        if not entries:
            self.curve_canvas.draw_idle()
            return
        legend_handles = []
        labels_seen = set()
        units_seen = []
        xlabel = "Bias"
        for sel in entries:
            base_color = sel.get("color", "#4c78a8")
            is_focus = sel.get("key") == self._focused_key
            multi = sel.get("multi") or [(sel.get("label"), sel.get("spec"))]
            for idx, (label, spec) in enumerate(multi):
                xs, ys, unit, resolved_label, x_unit = self._extract_channel_data(spec, label)
                if xs is None or ys is None:
                    continue
                labels_seen.add(resolved_label or label)
                if unit:
                    units_seen.append(unit)
                bias_vals = xs
                if x_unit:
                    xlabel = f"Bias ({x_unit})" if "bias" in xlabel.lower() else f"{xlabel} ({x_unit})"
                color = base_color if idx == 0 else self._variant_color(base_color, 0.35 + idx * 0.15)
                style = '-' if idx == 0 else '--'
                lw = 2.4 if is_focus and idx == 0 else 1.4
                alpha = 1.0 if is_focus else 0.75
                legend_label = resolved_label or label or "channel"
                if self._aggregate_mode:
                    nm = sel.get("nm_coords") or (None, None)
                    if nm[0] is not None and nm[1] is not None:
                        legend_label = f"{legend_label} @ ({nm[0]:.1f}, {nm[1]:.1f} nm)"
                line, = self.curve_ax.plot(bias_vals, ys, style, color=color, lw=lw, alpha=alpha, label=legend_label)
                legend_handles.append(line)
        self.curve_ax.set_xlabel(xlabel)
        axis_label = "Signal"
        if not self._aggregate_mode:
            active = entries[0]
            unit = active.get("unit") or (units_seen[0] if units_seen else None)
            base_label = active.get("label") or next(iter(labels_seen), "Signal")
            if unit:
                axis_label = f"{base_label} ({unit})"
            else:
                axis_label = base_label
        elif units_seen:
            distinct = {u for u in units_seen if u}
            if len(distinct) == 1:
                axis_label = f"Signal ({distinct.pop()})"
        self.curve_ax.set_ylabel(axis_label)
        self.curve_ax.grid(True, alpha=0.3)
        if legend_handles:
            self.curve_ax.legend(loc='upper right', fontsize=8)
        self.curve_canvas.draw_idle()

    def _gather_multi_channel_specs(self, matrix_index):
        if matrix_index is None:
            return []
        entries = []
        selected = self._normalize_path(self.channel_combo.currentData())
        for path, specs in self._channel_specs.items():
            if selected and path != selected:
                continue
            for spec in specs:
                if spec.get('matrix_index') == matrix_index:
                    entries.append((self._channel_label_for_path(path), spec))
                    break
        return entries

    def _clear_selection(self):
        self._selection = []
        self._selection_keys = set()
        self._aggregate_mode = False
        self._focused_key = None
        self._reset_color_cycle()
        if hasattr(self.viewer, '_clear_multi_spec_selection'):
            self.viewer._clear_multi_spec_selection()
        self.selection_table.clearContents()
        self.selection_table.setRowCount(0)
        self.curve_ax.clear()
        self._update_selection_markers()
        self.curve_canvas.draw_idle()

    def _on_fit_matrix(self):
        dlg = MatrixFitDialog(self.viewer, self.specs, parent=self)
        dlg.setAttribute(QtCore.Qt.WA_DeleteOnClose, True)
        dlg.show()
        self._fit_dialogs.append(dlg)
        dlg.finished.connect(lambda _: self._cleanup_fit_dialog(dlg))

    def _cleanup_fit_dialog(self, dlg):
        try:
            self._fit_dialogs.remove(dlg)
        except ValueError:
            pass

    def _on_export_selection(self):
        if not self._selection:
            QtWidgets.QMessageBox.information(self, "Export", "Select at least one spectrum.")
            return
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Export selection to CSV", "matrix_selection.csv", "CSV Files (*.csv)")
        if not path:
            return
        with open(path, "w", encoding="utf-8") as fh:
            fh.write("channel,index,x_nm,y_nm,bias,bias_unit,value\n")
            for entry in self._selection:
                spec = entry.get("spec")
                nm_coords = entry.get("nm_coords")
                if not spec:
                    continue
                bias_vals, _, bias_unit = self._axis_for_spec(spec)
                channels = spec.get('channels') or {}
                label = entry.get("label", Path(spec.get('path','')).name)
                ys = channels.get(label) or (spec.get('data')[1] if spec.get('data') else None)
                if bias_vals is None or ys is None:
                    continue
                x_nm = nm_coords[0] if nm_coords and nm_coords[0] is not None else float('nan')
                y_nm = nm_coords[1] if nm_coords and nm_coords[1] is not None else float('nan')
                idx = spec.get('matrix_index')
                for xv, yv in zip(bias_vals, ys):
                    fh.write(f"{label},{idx},{x_nm},{y_nm},{xv},{bias_unit},{yv}\n")
        QtWidgets.QMessageBox.information(self, "Export", f"Exported {len(self._selection)} selections to {Path(path).name}")

    def _on_canvas_hover(self, event):
        if event.inaxes != self.ax or self._current_image_arr is None:
            self.image_value_label.setText("Value: --")
            return
        val = sample_array_value(self._current_image_arr, event.xdata, event.ydata, self._current_image_extent)
        if val is None:
            self.image_value_label.setText("Value: --")
            return
        unit = self._current_image_unit or ''
        txt = f"Value: {val:.4g}"
        if unit:
            txt += f" {unit}"
        self.image_value_label.setText(txt)

class _SpectroFitWorker(QtCore.QObject):
    finished = QtCore.pyqtSignal(list, list)
    progress = QtCore.pyqtSignal(int, str)

    def __init__(self, specs, channel, axis_key):
        super().__init__()
        self.specs = list(specs)
        self.channel = channel
        self.axis_key = axis_key or "primary"

    @staticmethod
    def _axis_for_spec_with_key(spec, key):
        for ax in spec.get("AxisChoices") or []:
            if ax.get("key") == key:
                vals = np.asarray(ax.get("values", []), dtype=float)
                return vals, ax.get("label") or "Axis", ax.get("unit") or ""
        if key == "alt":
            alt_vals = spec.get("AltAxis")
            if alt_vals is not None:
                vals = np.asarray(alt_vals, dtype=float)
                return vals, spec.get("AltAxisLabel") or "Z rel", spec.get("AltAxisUnit") or ""
        vals = np.asarray(spec.get("V", []), dtype=float)
        return vals, spec.get("AxisLabel") or "Axis", spec.get("AxisUnit") or ""

    def run(self):
        results = []
        logs = []
        total_specs = len(self.specs)
        for i, spec in enumerate(self.specs):
            name = Path(spec['path']).name
            progress_msg = f"Fitting {name} ({i+1}/{total_specs})"
            self.progress.emit(int((i / total_specs) * 100), progress_msg)

            V, axis_label, axis_unit = self._axis_for_spec_with_key(spec, self.axis_key)
            channels = spec.get('channels') or {}
            data = channels.get(self.channel)
            if data is None or not V.size:
                logs.append(f"{name}: channel '{self.channel}' unavailable for axis '{axis_label}'")
                continue
            try:
                res = fit_parabola_bias(V, data)
                res['spec'] = spec
                res['axis_key'] = self.axis_key
                res['axis_label'] = axis_label
                res['axis_unit'] = axis_unit
                a = res.get('a'); b = res.get('b')
                v0 = None; v0_err = None
                if a is not None and b is not None and np.isfinite(a) and np.isfinite(b) and a != 0:
                    v0 = -b / (2.0 * a)
                    da = res.get('a_err', 0.0)
                    db = res.get('b_err', 0.0)
                    term1 = (db / (2.0 * a)) ** 2 if a != 0 else 0.0
                    term2 = ((b * da) / (2.0 * (a ** 2))) ** 2 if a != 0 else 0.0
                    v0_err = math.sqrt(max(term1 + term2, 0.0))
                res['v0'] = v0
                res['v0_err'] = v0_err
                results.append(res)
                logs.append(f"{name}: fit ok (RMSE {res['rmse']:.3g})")
            except Exception as e:
                logs.append(f"{name}: {e}")
        self.progress.emit(100, "Fit complete")
        self.finished.emit(results, logs)

class SpectroscopyCompareDialog(QtWidgets.QDialog):
    """Modern comparison UI for spectroscopy overlays and fitting."""
    def __init__(self, specs, parent=None, palette_name=None):
        super().__init__(parent)
        self.specs = list(specs)
        self._palette_name = palette_name or DEFAULT_COLOR_CYCLE
        self._color_cycle = get_color_cycle(self._palette_name)
        if not self._color_cycle:
            self._color_cycle = get_color_cycle(DEFAULT_COLOR_CYCLE)
        self._line_map = {}
        self._legend_map = {}
        self._fit_results = {}
        self._fit_thread = None
        self._fit_worker = None
        self._popup_refs = []
        self._background_spec_id = None
        self._relative_zero_enabled = False
        self._font_scale = 1.0
        self._lcpd_line_info = {}
        self._delta_selection = []
        self._delta_annotation_artists = []
        self._delta_hint_text = (
            "Hint: Shift+click two LCPD lines to show ΔLCPD annotations; "
            "toggle Points and Lines to change what is visible."
        )
        self._undo_stack = []
        self._suppress_undo_push = True
        self._lcpd_line_info = {}
        self._delta_selection = []
        self._delta_annotation_artists = []
        self.setWindowTitle("Spectroscopy comparison")
        self.resize(1400, 700)  # Increased size for better layout
        self._build_ui()
        self._populate_list()
        self._populate_channels()
        self._populate_axes()
        self._update_plot()
        self._suppress_undo_push = False

    def _get_icon(self, name):
        """Get a themed icon, falling back to empty icon if not available."""
        icon = QIcon.fromTheme(name)
        return icon if icon and not icon.isNull() else QIcon()

    def _display_name(self, spec):
        name = Path(spec.get('path', '')).name
        idx = spec.get('matrix_index')
        return f"{name} [m{idx}]" if idx is not None else name

    def _spec_id(self, spec):
        base = str(Path(spec.get('path', '')))
        idx = spec.get('matrix_index')
        return f"{base}#m{idx}" if idx is not None else base

    def _get_icon(self, name):
        """Get a themed icon, falling back to empty icon if not available."""
        icon = QIcon.fromTheme(name)
        return icon if icon and not icon.isNull() else QIcon()

    def _build_ui(self):
        splitter = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        main_layout = QtWidgets.QVBoxLayout()
        main_layout.addWidget(splitter)
        self.setLayout(main_layout)

        # Left panel: filter + list
        left = QtWidgets.QWidget()
        left_layout = QtWidgets.QVBoxLayout(left)
        left_layout.setContentsMargins(4,4,4,4)
        self.filter_edit = QtWidgets.QLineEdit()
        self.filter_edit.setPlaceholderText("Filter spectra...")
        self.filter_edit.setToolTip("Filter spectra by filename, type, position, or channels")
        self.filter_edit.textChanged.connect(self._apply_filter)
        self.filter_edit.setAccessibleName("Spectrum filter")
        self.filter_edit.setAccessibleDescription("Enter text to filter the list of spectra")
        self.spec_list = QtWidgets.QTreeWidget()
        self.spec_list.setHeaderLabels(["File", "Type", "Pos (nm)", "Time", "Chans"])
        self.spec_list.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
        self.spec_list.setAlternatingRowColors(True)
        self.spec_list.setRootIsDecorated(False)
        self.spec_list.setSortingEnabled(True)
        self.spec_list.itemChanged.connect(self._on_item_check_changed)
        self.spec_list.itemSelectionChanged.connect(self._on_list_selection_changed)
        self.spec_list.itemDoubleClicked.connect(self._on_item_double_clicked)
        self.spec_list.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.spec_list.customContextMenuRequested.connect(self._on_list_context_menu)
        self.spec_list.setAccessibleName("Spectra list")
        self.spec_list.setAccessibleDescription("List of available spectra. Check boxes to include in plot, select for additional operations")
        left_layout.addWidget(self.filter_edit)
        left_layout.addWidget(self.spec_list, 1)
        splitter.addWidget(left)
        splitter.setStretchFactor(0, 0)

        # Center panel: plot + status
        center = QtWidgets.QWidget()
        center_layout = QtWidgets.QVBoxLayout(center)
        center_layout.setContentsMargins(4,4,4,4)
        self.fig = Figure(figsize=(5,4))
        self.canvas = FigureCanvas(self.fig)
        self.ax = self.fig.add_subplot(111)
        self.ax.grid(True, alpha=0.2)
        center_layout.addWidget(self.canvas, 1)
        self.canvas.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.canvas.customContextMenuRequested.connect(self._on_compare_canvas_menu)
        self.canvas.mpl_connect("button_press_event", self._on_compare_canvas_click)
        self.canvas.mpl_connect("motion_notify_event", self._on_compare_canvas_motion)
        self.canvas.mpl_connect("key_press_event", self._on_compare_canvas_keypress)
        self.canvas.setAccessibleName("Spectroscopy comparison plot")
        self.canvas.setAccessibleDescription("Interactive plot showing selected spectra")

        # Progress bar for fitting operations
        self.progress_bar = QtWidgets.QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setAccessibleName("Fitting progress")
        self.progress_bar.setAccessibleDescription("Shows progress of spectrum fitting operations")
        center_layout.addWidget(self.progress_bar)

        self.status_label = QtWidgets.QLabel("0 selected / 0 total")
        self.status_label.setAccessibleName("Status information")
        self.status_label.setAccessibleDescription("Shows current selection and plot status")
        center_layout.addWidget(self.status_label)

        self.hint_label = QtWidgets.QLabel(self._delta_hint_text)
        self.hint_label.setWordWrap(True)
        self.hint_label.setAccessibleName("Plot interaction hint")
        self.hint_label.setAccessibleDescription("Tips about interacting with the comparison plot")
        center_layout.addWidget(self.hint_label)

        # Visualization controls (Waterfall)
        vis_group = QtWidgets.QGroupBox("Visualization")
        vis_layout = QtWidgets.QVBoxLayout(vis_group)
        vis_row = QtWidgets.QHBoxLayout()
        self.waterfall_cb = QtWidgets.QCheckBox("Waterfall")
        self.waterfall_cb.setToolTip("Stack spectra vertically with offset for better visibility")
        self.waterfall_cb.toggled.connect(self._on_visual_toggle)
        self.waterfall_cb.setAccessibleName("Waterfall display")
        self.waterfall_cb.setAccessibleDescription("Enable waterfall stacking of spectra")
        vis_row.addWidget(self.waterfall_cb)

        self.show_points_cb = QtWidgets.QCheckBox("Points")
        self.show_points_cb.setToolTip("Show data points")
        self.show_points_cb.toggled.connect(self._on_visual_toggle)
        vis_row.addWidget(self.show_points_cb)

        self.lines_cb = QtWidgets.QCheckBox("Lines")
        self.lines_cb.setToolTip("Show lines connecting the spectroscopy curves")
        self.lines_cb.setAccessibleName("Lines toggle")
        self.lines_cb.setAccessibleDescription("Show/hide the curves connecting the spectroscopy data")
        self.lines_cb.setChecked(True)
        self.lines_cb.toggled.connect(self._on_visual_toggle)
        vis_row.addWidget(self.lines_cb)

        self.offset_spin = QtWidgets.QDoubleSpinBox()
        self.offset_spin.setRange(-1e9, 1e9)
        self.offset_spin.setDecimals(14) # High precision for small currents
        self.offset_spin.setSingleStep(0.1)
        self.offset_spin.setToolTip("Vertical offset between waterfall spectra")
        self.offset_spin.valueChanged.connect(self._on_offset_changed)
        self.offset_spin.setAccessibleName("Waterfall offset")
        self.offset_spin.setAccessibleDescription("Set the vertical spacing between stacked spectra")
        vis_row.addWidget(QtWidgets.QLabel("Offset:"))
        vis_row.addWidget(self.offset_spin)
        vis_row.addStretch(1)
        vis_layout.addLayout(vis_row)
        undo_row = QtWidgets.QHBoxLayout()
        self.undo_btn = QtWidgets.QPushButton("Undo")
        self.undo_btn.setToolTip("Revert the most recent change to the comparison (Ctrl+Z)")
        self.undo_btn.setEnabled(False)
        self.undo_btn.clicked.connect(self._undo_last_action)
        undo_row.addWidget(self.undo_btn)
        undo_row.addStretch(1)
        vis_layout.addLayout(undo_row)
        center_layout.addWidget(vis_group)
        splitter.addWidget(center)
        splitter.setStretchFactor(1, 2)
        self.canvas.mpl_connect('pick_event', self._on_legend_pick)

        # Right panel: controls + results
        right = QtWidgets.QWidget()
        right_layout = QtWidgets.QVBoxLayout(right)
        right_layout.setContentsMargins(6,6,6,6)

        # Data Selection Group
        data_group = QtWidgets.QGroupBox("Data Selection")
        data_layout = QtWidgets.QVBoxLayout(data_group)

        channel_row = QtWidgets.QHBoxLayout()
        channel_row.addWidget(QtWidgets.QLabel("Channel:"))
        self.channel_combo = QtWidgets.QComboBox()
        self.channel_combo.setToolTip("Select which channel to plot and analyze")
        self.channel_combo.currentTextChanged.connect(self._on_channel_changed)
        self.channel_combo.setAccessibleName("Channel selection")
        self.channel_combo.setAccessibleDescription("Choose which data channel to display")
        channel_row.addWidget(self.channel_combo, 1)
        data_layout.addLayout(channel_row)

        axis_row = QtWidgets.QHBoxLayout()
        axis_row.addWidget(QtWidgets.QLabel("Axis:"))
        self.axis_combo = QtWidgets.QComboBox()
        self.axis_combo.setToolTip("Select X-axis for plotting (bias voltage or Z position)")
        self.axis_combo.currentIndexChanged.connect(self._on_axis_changed)
        self.axis_combo.setAccessibleName("Axis selection")
        self.axis_combo.setAccessibleDescription("Choose the X-axis variable for the plot")
        axis_row.addWidget(self.axis_combo, 1)
        self.relative_cb = QtWidgets.QCheckBox("Relative Z (zero at min)")
        self.relative_cb.setToolTip("Shift Z-axis to start from zero at minimum value")
        self.relative_cb.toggled.connect(self._on_relative_toggled)
        self.relative_cb.setAccessibleName("Relative Z mode")
        self.relative_cb.setAccessibleDescription("Enable relative Z-axis scaling")
        axis_row.addWidget(self.relative_cb)
        data_layout.addLayout(axis_row)

        right_layout.addWidget(data_group)

        # Visualization Group
        viz_group = QtWidgets.QGroupBox("Appearance")
        viz_layout = QtWidgets.QVBoxLayout(viz_group)

        palette_row = QtWidgets.QHBoxLayout()
        palette_row.addWidget(QtWidgets.QLabel("Color cycle:"))
        self.palette_combo = QtWidgets.QComboBox()
        for name in list_color_cycles():
            self.palette_combo.addItem(name)
        self.palette_combo.setToolTip("Select color palette for spectrum lines")
        self.palette_combo.currentTextChanged.connect(self._on_palette_changed_compare)
        self.palette_combo.setAccessibleName("Color palette")
        self.palette_combo.setAccessibleDescription("Choose color scheme for plotting multiple spectra")
        self.palette_combo.blockSignals(True)
        default_idx = max(0, self.palette_combo.findText(self._palette_name))
        self.palette_combo.setCurrentIndex(default_idx)
        self.palette_combo.blockSignals(False)
        palette_row.addWidget(self.palette_combo, 1)
        viz_layout.addLayout(palette_row)

        self.palette_swatches = QtWidgets.QWidget()
        swatch_layout = QtWidgets.QHBoxLayout(self.palette_swatches)
        swatch_layout.setSpacing(3)
        swatch_layout.setContentsMargins(0, 4, 0, 4)
        swatch_layout.setAlignment(QtCore.Qt.AlignLeft)
        self.palette_swatches.setAccessibleName("Color cycle swatches")
        self.palette_swatches.setAccessibleDescription("Shows the colors currently available in the selected color cycle")
        viz_layout.addWidget(self.palette_swatches)

        right_layout.addWidget(viz_group)

        # Analysis Group
        analysis_group = QtWidgets.QGroupBox("Analysis")
        analysis_layout = QtWidgets.QVBoxLayout(analysis_group)

        # KPFM subsection
        kpfm_group = QtWidgets.QGroupBox("KPFM")
        kpfm_layout = QtWidgets.QVBoxLayout(kpfm_group)

        fit_row = QtWidgets.QHBoxLayout()
        self.fit_selected_btn = QtWidgets.QPushButton(self._get_icon("system-run"), "Fit selected (F)")
        self.fit_selected_btn.setToolTip("Fit parabola to selected spectra")
        self.fit_all_btn = QtWidgets.QPushButton(self._get_icon("edit-select-all"), "Fit all")
        self.fit_all_btn.setToolTip("Fit parabola to all checked spectra")
        fit_row.addWidget(self.fit_selected_btn)
        fit_row.addWidget(self.fit_all_btn)
        kpfm_layout.addLayout(fit_row)

        export_row = QtWidgets.QHBoxLayout()
        self.export_btn = QtWidgets.QPushButton(self._get_icon("document-save"), "Export CSV")
        self.export_btn.setToolTip("Export fit results to CSV file")
        export_row.addWidget(self.export_btn)
        export_row.addStretch(1)
        kpfm_layout.addLayout(export_row)

        analysis_layout.addWidget(kpfm_group)

        # Forces/Background subsection
        forces_group = QtWidgets.QGroupBox("Forces/Background")
        forces_layout = QtWidgets.QVBoxLayout(forces_group)

        bg_row = QtWidgets.QHBoxLayout()
        self.bg_set_btn = QtWidgets.QPushButton(self._get_icon("list-add"), "Set background")
        self.bg_set_btn.setToolTip("Set selected spectrum as background for subtraction")
        self.bg_clear_btn = QtWidgets.QPushButton(self._get_icon("list-remove"), "Clear background")
        self.bg_clear_btn.setToolTip("Remove background subtraction")
        bg_row.addWidget(self.bg_set_btn)
        bg_row.addWidget(self.bg_clear_btn)
        forces_layout.addLayout(bg_row)

        force_row = QtWidgets.QHBoxLayout()
        self.force_btn = QtWidgets.QPushButton(self._get_icon("transform-scale"), "Convert to force")
        self.force_btn.setToolTip("Convert spectra to force curves (experimental)")
        force_row.addWidget(self.force_btn)
        force_row.addStretch(1)
        forces_layout.addLayout(force_row)

        analysis_layout.addWidget(forces_group)

        right_layout.addWidget(analysis_group)

        # Actions Group
        actions_group = QtWidgets.QGroupBox("Actions")
        actions_layout = QtWidgets.QVBoxLayout(actions_group)

        copy_row = QtWidgets.QHBoxLayout()
        self.copy_btn = QtWidgets.QPushButton(self._get_icon("edit-copy"), "Copy selected")
        self.copy_btn.setToolTip("Copy selected spectra data to clipboard")
        self.copy_table_btn = QtWidgets.QPushButton(self._get_icon("edit-copy"), "Copy table")
        self.copy_table_btn.setToolTip("Copy fit results table to clipboard")
        copy_row.addWidget(self.copy_btn)
        copy_row.addWidget(self.copy_table_btn)
        actions_layout.addLayout(copy_row)

        clear_row = QtWidgets.QHBoxLayout()
        self.clear_sel_btn = QtWidgets.QPushButton(self._get_icon("edit-clear"), "Clear selected")
        self.clear_sel_btn.setToolTip("Remove selected spectra from list")
        self.clear_all_btn = QtWidgets.QPushButton(self._get_icon("edit-clear-all"), "Clear all")
        self.clear_all_btn.setToolTip("Clear all spectra from list")
        clear_row.addWidget(self.clear_sel_btn)
        clear_row.addWidget(self.clear_all_btn)
        actions_layout.addLayout(clear_row)

        help_row = QtWidgets.QHBoxLayout()
        self.help_btn = QtWidgets.QPushButton(self._get_icon("help-about"), "Help")
        self.help_btn.setToolTip("Show help for spectroscopy comparison")
        self.help_btn.setAccessibleName("Help")
        self.help_btn.setAccessibleDescription("Open help documentation for spectroscopy comparison features")
        self.help_btn.clicked.connect(self._show_help)
        help_row.addWidget(self.help_btn)
        help_row.addStretch(1)
        actions_layout.addLayout(help_row)

        right_layout.addWidget(actions_group)

        # Connect button signals
        self.fit_selected_btn.clicked.connect(self._fit_selected)
        self.fit_all_btn.clicked.connect(self._fit_all)
        self.export_btn.clicked.connect(self._export_csv)
        self.bg_set_btn.clicked.connect(self._on_set_background)
        self.bg_clear_btn.clicked.connect(self._on_clear_background)
        self.force_btn.clicked.connect(self._on_convert_force)
        self.copy_btn.clicked.connect(self._copy_selected_to_clipboard)
        self.copy_table_btn.clicked.connect(self._copy_table_to_clipboard)
        self.clear_sel_btn.clicked.connect(self._clear_selected)
        self.clear_all_btn.clicked.connect(self._clear_all)

        # Set accessibility for buttons
        self.fit_selected_btn.setAccessibleName("Fit selected spectra")
        self.fit_selected_btn.setAccessibleDescription("Perform parabolic fit on selected spectra")
        self.fit_all_btn.setAccessibleName("Fit all spectra")
        self.fit_all_btn.setAccessibleDescription("Perform parabolic fit on all checked spectra")
        self.export_btn.setAccessibleName("Export results")
        self.export_btn.setAccessibleDescription("Save fit results to CSV file")
        self.bg_set_btn.setAccessibleName("Set background")
        self.bg_set_btn.setAccessibleDescription("Use selected spectrum as background for subtraction")
        self.bg_clear_btn.setAccessibleName("Clear background")
        self.bg_clear_btn.setAccessibleDescription("Remove background subtraction")
        self.force_btn.setAccessibleName("Convert to force")
        self.force_btn.setAccessibleDescription("Convert spectra to force curves")
        self.copy_btn.setAccessibleName("Copy spectra")
        self.copy_btn.setAccessibleDescription("Copy selected spectra data to clipboard")
        self.copy_table_btn.setAccessibleName("Copy table")
        self.copy_table_btn.setAccessibleDescription("Copy fit results table to clipboard")
        self.clear_sel_btn.setAccessibleName("Clear selected")
        self.clear_sel_btn.setAccessibleDescription("Remove selected spectra from the list")
        self.clear_all_btn.setAccessibleName("Clear all")
        self.clear_all_btn.setAccessibleDescription("Remove all spectra from the list")

        # Keyboard shortcuts
        QtWidgets.QShortcut(QtGui.QKeySequence("F"), self, activated=self._fit_selected)
        QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+E"), self, activated=self._export_csv)
        QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+A"), self, activated=self._select_all_visible)
        QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+Shift+A"), self, activated=self._invert_selection)
        QtWidgets.QShortcut(QtGui.QKeySequence("Delete"), self, activated=self._clear_selected)
        QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+Delete"), self, activated=self._clear_all)

        # Fit options collapsible section
        self.options_toggle = QtWidgets.QToolButton()
        self.options_toggle.setText("Fit options")
        self.options_toggle.setToolTip("Show/hide advanced fitting options")
        self.options_toggle.setCheckable(True)
        self.options_toggle.setChecked(False)
        self.options_toggle.setArrowType(QtCore.Qt.RightArrow)
        self.options_toggle.setToolButtonStyle(QtCore.Qt.ToolButtonTextBesideIcon)
        self.options_toggle.toggled.connect(self._on_options_toggled)
        self.options_toggle.setAccessibleName("Fit options toggle")
        self.options_toggle.setAccessibleDescription("Expand to show advanced fitting parameters")
        right_layout.addWidget(self.options_toggle)

        self.options_body = QtWidgets.QWidget()
        form = QtWidgets.QFormLayout(self.options_body)
        self.degree_spin = QtWidgets.QSpinBox()
        self.degree_spin.setRange(2, 2)
        self.degree_spin.setValue(2)
        self.degree_spin.setEnabled(False)
        self.degree_spin.setToolTip("Polynomial degree for fitting (fixed at 2)")
        form.addRow("Degree", self.degree_spin)
        self.mask_min = QtWidgets.QDoubleSpinBox()
        self.mask_min.setRange(-1e6, 1e6)
        self.mask_min.setSuffix(" V")
        self.mask_min.setToolTip("Minimum bias voltage to include in fit")
        self.mask_min.setAccessibleName("Fit mask minimum")
        self.mask_min.setAccessibleDescription("Exclude data below this bias voltage from fitting")
        self.mask_max = QtWidgets.QDoubleSpinBox()
        self.mask_max.setRange(-1e6, 1e6)
        self.mask_max.setSuffix(" V")
        self.mask_max.setToolTip("Maximum bias voltage to include in fit")
        self.mask_max.setAccessibleName("Fit mask maximum")
        self.mask_max.setAccessibleDescription("Exclude data above this bias voltage from fitting")
        form.addRow("Mask min", self.mask_min)
        form.addRow("Mask max", self.mask_max)
        self.options_body.setVisible(False)
        right_layout.addWidget(self.options_body)

        # Separator
        separator = QtWidgets.QFrame()
        separator.setFrameShape(QtWidgets.QFrame.HLine)
        separator.setFrameShadow(QtWidgets.QFrame.Sunken)
        right_layout.addWidget(separator)

        # Results table
        table_label = QtWidgets.QLabel("Fit Results")
        table_label.setStyleSheet("font-weight: bold;")
        right_layout.addWidget(table_label)

        self.results_table = QtWidgets.QTableWidget(0, 10)
        self.results_table.setHorizontalHeaderLabels(["File","X (nm)","Y (nm)","a","δa","LCPD","δLCPD","c (Hz)","δc","RMSE"])
        self.results_table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.results_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.results_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.results_table.setSortingEnabled(True)  # Enable sorting
        self.results_table.itemSelectionChanged.connect(self._on_table_selection)
        self.results_table.itemDoubleClicked.connect(self._on_table_double_clicked)
        self.results_table.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.results_table.customContextMenuRequested.connect(self._on_table_context_menu)
        self.results_table.setAccessibleName("Fit results table")
        self.results_table.setAccessibleDescription("Table showing results of parabolic fits to spectra")
        right_layout.addWidget(self.results_table, 1)

        # Log
        log_label = QtWidgets.QLabel("Log")
        log_label.setStyleSheet("font-weight: bold;")
        right_layout.addWidget(log_label)

        self.log = QtWidgets.QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumHeight(100)
        self.log.setAccessibleName("Operation log")
        self.log.setAccessibleDescription("Shows messages from fitting and other operations")
        right_layout.addWidget(self.log)

        splitter.addWidget(right)
        splitter.setStretchFactor(2, 1)

    def _populate_list(self):
        self.spec_list.blockSignals(True)
        self.spec_list.clear()
        self._item_map = {}
        for spec in self.specs:
            path = Path(spec.get('path', ''))
            name = path.name
            
            # Type/Index
            midx = spec.get('matrix_index')
            type_str = f"Matrix [{midx}]" if midx is not None else "Single"
            
            # Pos
            x, y = spec.get('x'), spec.get('y')
            pos_str = f"{x:.1f}, {y:.1f}" if x is not None and y is not None else "-"
            
            # Time
            t = spec.get('time')
            time_str = ""
            if isinstance(t, datetime):
                time_str = t.strftime("%H:%M:%S")
            else:
                time_str = str(t)

            # Channels
            chans = list((spec.get('channels') or {}).keys())
            chans_str = ", ".join(chans)

            item = QtWidgets.QTreeWidgetItem([name, type_str, pos_str, time_str, chans_str])
            item.setFlags(item.flags() | QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsSelectable)
            item.setCheckState(0, QtCore.Qt.Checked)
            item.setData(0, QtCore.Qt.UserRole, spec)
            item.setData(0, QtCore.Qt.UserRole + 1, self._spec_id(spec))
            self.spec_list.addTopLevelItem(item)
            self._item_map[self._spec_id(spec)] = item
        
        for i in range(5):
            self.spec_list.resizeColumnToContents(i)
        self.spec_list.blockSignals(False)

    def set_specs(self, specs):
        """Update the dialog with a new list of spectra without reopening."""
        self.specs = list(specs)
        self._fit_results = {}
        self._item_map = {}
        prev_channel = self.channel_combo.currentText()
        filter_text = self.filter_edit.text()
        self.spec_list.blockSignals(True)
        self.spec_list.clear()
        self.spec_list.blockSignals(False)
        self._populate_list()
        self._populate_channels()
        if prev_channel:
            idx = self.channel_combo.findText(prev_channel)
            if idx >= 0:
                self.channel_combo.setCurrentIndex(idx)
        if filter_text:
            self.filter_edit.setText(filter_text)
            self._apply_filter(filter_text)
        self._populate_results_table()
        self._update_plot()

    def set_palette_name(self, name):
        cycle = name or DEFAULT_COLOR_CYCLE
        if cycle == self._palette_name:
            return
        self._palette_name = cycle
        self._color_cycle = get_color_cycle(self._palette_name)
        if not self._color_cycle:
            self._color_cycle = get_color_cycle(DEFAULT_COLOR_CYCLE)
        idx = self.palette_combo.findText(self._palette_name)
        self.palette_combo.blockSignals(True)
        if idx >= 0:
            self.palette_combo.setCurrentIndex(idx)
        else:
            self.palette_combo.setCurrentIndex(0)
            self._palette_name = self.palette_combo.currentText()
            self._color_cycle = get_color_cycle(self._palette_name)
        self.palette_combo.blockSignals(False)
        self._update_plot()
        self._update_color_swatches()

    def _populate_channels(self):
        channels = sorted({name for spec in self.specs for name in (spec.get('channels') or {}).keys()})
        self.channel_combo.blockSignals(True)
        self.channel_combo.clear()
        for name in channels:
            self.channel_combo.addItem(name)
        if channels:
            self.channel_combo.setCurrentText('df' if 'df' in channels else channels[0])
        self.channel_combo.blockSignals(False)

    def _populate_axes(self):
        axes = []
        for spec in self.specs:
            if spec.get("AxisChoices"):
                for ax in spec.get("AxisChoices"):
                    axes.append((ax.get("key"), ax.get("label") or "Axis", ax.get("unit") or "", np.asarray(ax.get("values", []), dtype=float)))
            else:
                primary_lbl = spec.get("AxisLabel") or "Axis"
                primary_unit = spec.get("AxisUnit") or ""
                axes.append(("primary", primary_lbl, primary_unit, np.asarray(spec.get("V", []), dtype=float)))
                if spec.get("AltAxis") is not None:
                    axes.append(("alt", spec.get("AltAxisLabel") or "Z rel", spec.get("AltAxisUnit") or "", np.asarray(spec.get("AltAxis"), dtype=float)))
        # dedupe by key+values to avoid duplicate bias axes
        seen = []
        options = []
        for key, lbl, unit, vals in axes:
            duplicate = False
            for s_key, s_vals in seen:
                if key == s_key and np.array_equal(vals, s_vals):
                    duplicate = True
                    break
            if duplicate:
                continue
            seen.append((key, vals))
            disp = lbl if not unit else (f"{lbl} ({unit})" if unit not in lbl else lbl)
            options.append((disp, key))
        self.axis_combo.blockSignals(True)
        self.axis_combo.clear()
        for disp, key in options:
            self.axis_combo.addItem(disp, key)
        # default to primary if available
        idx = max(0, self.axis_combo.findData("primary"))
        self.axis_combo.setCurrentIndex(idx)
        self.axis_combo.blockSignals(False)
        self._update_color_swatches()

    def _apply_filter(self, text):
        text = text.lower()
        root = self.spec_list.invisibleRootItem()
        for i in range(root.childCount()):
            item = root.child(i)
            match = False
            for c in range(item.columnCount()):
                if text in item.text(c).lower():
                    match = True
                    break
            item.setHidden(not match)
        self._update_status()

    def _checked_items(self):
        items = []
        root = self.spec_list.invisibleRootItem()
        for i in range(root.childCount()):
            item = root.child(i)
            if item.checkState(0) == QtCore.Qt.Checked and not item.isHidden():
                items.append(item)
        return items

    def _selected_items(self):
        return self.spec_list.selectedItems()

    def _axis_for_spec(self, spec):
        """Return (values, label, unit) for the currently selected axis choice."""
        axis_choice = getattr(self, "axis_combo", None)
        choice_key = axis_choice.currentData() if axis_choice is not None else "primary"
        return self._axis_for_spec_with_key(spec, choice_key)

    def _axis_for_spec_with_key(self, spec, choice_key):
        for ax in spec.get("AxisChoices") or []:
            if ax.get("key") == choice_key:
                vals = np.asarray(ax.get("values", []), dtype=float)
                return vals, ax.get("label") or "Axis", ax.get("unit") or ""
        if choice_key == "alt":
            alt_vals = spec.get("AltAxis")
            if alt_vals is not None:
                vals = np.asarray(alt_vals, dtype=float)
                return vals, spec.get("AltAxisLabel") or "Z rel", spec.get("AltAxisUnit") or ""
        vals = np.asarray(spec.get("V", []), dtype=float)
        return vals, spec.get("AxisLabel") or "Axis", spec.get("AxisUnit") or ""

    def _on_set_background(self):
        self._record_user_action("Set background")
        items = self._selected_items() or self._checked_items()
        if not items:
            QtWidgets.QMessageBox.information(self, "Background", "Select a spectrum to set as background.")
            return
        spec = items[0].data(0, QtCore.Qt.UserRole)
        self._background_spec_id = self._spec_id(spec) if spec else None
        self._log(f"Background set: {Path(spec.get('path','')).name if spec else ''}")
        self._update_plot()

    def _on_clear_background(self):
        self._record_user_action("Clear background")
        self._background_spec_id = None
        self._update_plot()

    def _background_for(self, spec):
        if not self._background_spec_id:
            return None
        for s in self.specs:
            if self._spec_id(s) == self._background_spec_id:
                return s
        return None

    def _subtract_background(self, x_vals, y_vals, bg_spec):
        if bg_spec is None:
            return y_vals
        bg_x, _, _ = self._axis_for_spec(bg_spec)
        bg_channels = bg_spec.get("channels") or {}
        channel = self.channel_combo.currentText()
        bg_y = np.asarray(bg_channels.get(channel), dtype=float)
        if bg_y.size == 0 or bg_x.size == 0:
            return y_vals
        try:
            bg_interp = np.interp(x_vals, bg_x, bg_y)
            return y_vals - bg_interp
        except Exception:
            return y_vals

    def _on_convert_force(self):
        # Prompt for parameters
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("Convert to force")
        form = QtWidgets.QFormLayout(dlg)
        f0_edit = QtWidgets.QDoubleSpinBox(); f0_edit.setRange(0, 1e9); f0_edit.setValue(0.0)
        a_edit = QtWidgets.QDoubleSpinBox(); a_edit.setRange(0, 1e6); a_edit.setDecimals(9); a_edit.setValue(0.0)
        q_edit = QtWidgets.QDoubleSpinBox(); q_edit.setRange(0, 1e6); q_edit.setValue(0.0)
        method_combo = QtWidgets.QComboBox(); method_combo.addItems(["saderF", "matrixF"])
        form.addRow("f0 (Hz)", f0_edit)
        form.addRow("Amplitude A (m)", a_edit)
        form.addRow("Q", q_edit)
        form.addRow("Method", method_combo)
        btns = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        form.addRow(btns)
        btns.accepted.connect(dlg.accept); btns.rejected.connect(dlg.reject)
        if dlg.exec_() != QtWidgets.QDialog.Accepted:
            return
        f0 = f0_edit.value(); A = a_edit.value(); Q = q_edit.value(); method = method_combo.currentText()
        items = self._selected_items() or self._checked_items()
        if not items:
            QtWidgets.QMessageBox.information(self, "Force conversion", "Select spectra to convert.")
            return
        new_specs = []
        for item in items:
            spec = item.data(0, QtCore.Qt.UserRole)
            if not spec:
                continue
            channels = spec.get("channels") or {}
            converted = {}
            for name, arr in channels.items():
                try:
                    converted[name] = np.asarray(arr, dtype=float).copy()
                except Exception:
                    continue
            new_spec = dict(spec)
            new_spec["channels"] = converted
            new_spec["ForceMethod"] = method
            new_spec["ForceParams"] = {"f0": f0, "A": A, "Q": Q}
            new_specs.append(new_spec)
        if not new_specs:
            return
        # Open a twin dialog with converted data
        twin = SpectroscopyCompareDialog(new_specs, parent=self.parent(), palette_name=self._palette_name)
        twin.setWindowTitle("Spectroscopy comparison (force)")
        twin.show()
        self._popup_refs.append(twin)

    def _channel_unit_for_spec(self, spec, channel_label):
        unit_map = spec.get('unit_map') or {}
        if channel_label and channel_label in unit_map and unit_map[channel_label]:
            return unit_map[channel_label]
        if unit_map:
            for key, val in unit_map.items():
                if val:
                    return val
        return guess_channel_unit(channel_label)

    def _on_channel_changed(self):
        self._record_user_action(f"Channel → {self.channel_combo.currentText()}")
        self._fit_results = {}
        self._populate_results_table()
        self._update_plot()

    def _on_axis_changed(self):
        self._record_user_action(f"Axis → {self.axis_combo.currentText()}")
        self._fit_results = {}
        self.results_table.setRowCount(0)
        self._update_plot()

    def _on_relative_toggled(self, checked):
        self._record_user_action(f"Relative Z → {'on' if checked else 'off'}")
        self._relative_zero_enabled = bool(checked)
        self._update_plot()

    def _on_item_check_changed(self, item, column):
        self._record_user_action("Traffic: checked item changed")
        self._update_plot()

    def _on_list_selection_changed(self):
        self._record_user_action("Selection changed")
        self._update_plot()

    def _update_plot(self):
        channel = self.channel_combo.currentText()
        self.ax.clear()
        self.ax.grid(True, alpha=0.2)
        self._lcpd_line_info.clear()
        self._clear_delta_selection(redraw=False)
        self._line_map.clear()
        self._legend_map.clear()
        
        waterfall = self.waterfall_cb.isChecked()
        show_points = self.show_points_cb.isChecked()
        show_lines = self.lines_cb.isChecked()
        offset_val = self.offset_spin.value()
        scale = self._estimate_channel_scale(channel)
        self._configure_offset_spin(scale)
        relative_nm = bool(self._relative_zero_enabled)

        selected_ids = {item.data(0, QtCore.Qt.UserRole + 1) for item in self._selected_items()}
        colors = self._iter_color_cycle()
        plotted = 0

        # Precompute relative zero if needed
        rel_zero = 0.0
        if relative_nm:
            mins = []
            for item in self._selected_items() or self._checked_items():
                spec = item.data(0, QtCore.Qt.UserRole)
                if not spec:
                    continue
                axis_vals, _, unit = self._axis_for_spec(spec)
                if axis_vals.size and unit == "nm":
                    mins.append(np.nanmin(axis_vals))
            if mins:
                rel_zero = min(mins)

        # Plot both checked items AND selected items (even if unchecked) for quick preview
        root = self.spec_list.invisibleRootItem()
        for i in range(root.childCount()):
            item = root.child(i)
            if item.isHidden(): continue
            if item.checkState(0) != QtCore.Qt.Checked and not item.isSelected():
                continue

            spec = item.data(0, QtCore.Qt.UserRole)
            spec_id = item.data(0, QtCore.Qt.UserRole + 1)
            channels = spec.get('channels') or {}
            data = channels.get(channel)
            axis_vals, axis_label, axis_unit = self._axis_for_spec(spec)
            if data is None or not axis_vals.size:
                continue

            # Apply background subtraction if requested
            bg_spec = self._background_for(spec)
            y_base = self._subtract_background(axis_vals, data, bg_spec)
            # Apply waterfall offset
            y_data = y_base + (plotted * offset_val) if waterfall else y_base
            x_vals = axis_vals
            if relative_nm and axis_unit == "nm":
                x_vals = x_vals - rel_zero
            axis_plot_scale = 1.0
            axis_unit_plot = axis_unit
            if axis_unit.lower() == "v" and np.isfinite(x_vals).any():
                axis_plot_scale = 1000.0
                axis_unit_plot = "mV"
                x_vals = x_vals * axis_plot_scale
            color = next(colors)
            highlight = spec_id in selected_ids or not selected_ids
            label_txt = self._display_name(spec)
            line_kwargs = {
                "color": color,
                "lw": 2.4 if highlight else 1.2,
                "alpha": 1.0 if highlight else 0.4,
                "label": label_txt,
            }
            line_kwargs["linestyle"] = "-" if show_lines else "None"
            if show_points:
                line_kwargs.update({
                    "marker": "o",
                    "markersize": 2.6,
                    "markerfacecolor": color,
                    "markeredgecolor": color,
                    "markeredgewidth": 0.6,
                })
            line, = self.ax.plot(x_vals, y_data, **line_kwargs)
            self._line_map[spec_id] = line
            plotted += 1
            if spec_id in self._fit_results:
                self._draw_fit_for_spec(spec_id, color, offset=(plotted - 1) * offset_val if waterfall else 0.0)
        if plotted == 0:
            self.ax.text(0.5,0.5,"No data for selected items", ha='center', va='center', transform=self.ax.transAxes)
        else:
            legend = self.ax.legend(loc='best', fontsize=8)
            if legend:
                legend.set_draggable(True)
                for leg_line, text in zip(legend.get_lines(), legend.get_texts()):
                    leg_line.set_picker(True)
                    name = text.get_text()
                    for spec in self.specs:
                        if self._display_name(spec) == name:
                            self._legend_map[leg_line] = self._spec_id(spec)
                            break
        xlabel = "Axis"
        if relative_nm:
            xlabel = "Z (nm, relative)"
        else:
            # derive from any available spec axis label
            for item in self._selected_items() or self._checked_items():
                spec = item.data(0, QtCore.Qt.UserRole)
                if spec:
                    _, lbl, unit = self._axis_for_spec(spec)
                    if lbl:
                        xlabel = lbl
                    if unit and unit.lower() == "v":
                        xlabel = f"{lbl} (mV)"
                    elif unit and unit not in xlabel:
                        xlabel = f"{lbl} ({unit})"
                    break
        self.ax.set_xlabel(xlabel)
        unit = None
        # Find a representative unit for the y-axis label
        for item in self._checked_items() or self._selected_items():
            spec = item.data(0, QtCore.Qt.UserRole)
            if spec:
                unit = self._channel_unit_for_spec(spec, channel)
            if unit:
                break
        self.ax.set_ylabel(f"{channel} ({unit})" if unit else channel)
        self._apply_font_scale()
        # canvas.draw_idle() is called in _apply_font_scale
        self._update_status(plotted)

    def _iter_color_cycle(self):
        palette = self._color_cycle or get_color_cycle(DEFAULT_COLOR_CYCLE)
        if not palette:
            palette = get_color_cycle(DEFAULT_COLOR_CYCLE)
        return itertools.cycle(palette)

    def _update_color_swatches(self):
        container = getattr(self, "palette_swatches", None)
        if container is None or container.layout() is None:
            return
        layout = container.layout()
        while layout.count():
            item = layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()
        if not self._color_cycle:
            return
        for color in self._color_cycle:
            try:
                color_hex = mcolors.to_hex(color)
            except Exception:
                color_hex = str(color)
            swatch = QtWidgets.QLabel()
            swatch.setFixedSize(20, 20)
            swatch.setStyleSheet(f"background-color: {color_hex}; border: 1px solid #888;")
            layout.addWidget(swatch)
        layout.addStretch(1)

    def _estimate_channel_scale(self, channel):
        spreads = []
        root = self.spec_list.invisibleRootItem()
        for i in range(root.childCount()):
            item = root.child(i)
            if item.isHidden():
                continue
            spec = item.data(0, QtCore.Qt.UserRole)
            if not spec:
                continue
            arr = (spec.get('channels') or {}).get(channel)
            if arr is None:
                continue
            vec = np.asarray(arr, dtype=float)
            if vec.size == 0:
                continue
            try:
                rng = float(np.nanmax(vec) - np.nanmin(vec))
            except Exception:
                continue
            if np.isfinite(rng) and rng > 0:
                spreads.append(rng)
        if not spreads:
            return 1.0
        val = float(np.nanmedian(spreads))
        if not np.isfinite(val) or val <= 0:
            val = 1.0
        return val

    def _configure_offset_spin(self, scale):
        if not np.isfinite(scale) or scale <= 0:
            scale = 1.0
        rng = max(scale * 20.0, 1e-6)
        step = max(scale / 10.0, rng / 200.0)
        value = self.offset_spin.value()
        self.offset_spin.blockSignals(True)
        self.offset_spin.setRange(-rng, rng)
        self.offset_spin.setSingleStep(step)
        if value > rng or value < -rng:
            self.offset_spin.setValue(0.0)
        self.offset_spin.blockSignals(False)

    def _on_palette_changed_compare(self, name):
        self._palette_name = name or DEFAULT_COLOR_CYCLE
        self._color_cycle = get_color_cycle(self._palette_name)
        if not self._color_cycle:
            self._color_cycle = get_color_cycle(DEFAULT_COLOR_CYCLE)
        parent = self.parent()
        if parent and hasattr(parent, "set_spectro_color_cycle"):
            parent.set_spectro_color_cycle(self._palette_name)
        self._update_color_swatches()
        self._update_plot()

    def _draw_fit_for_spec(self, spec_id, color, offset=0.0):
        res = self._fit_results.get(spec_id)
        if not res:
            return
        spec = res.get('spec')
        axis_key = res.get('axis_key', "primary")
        axis_vals, _, axis_unit = self._axis_for_spec(spec) if axis_key is None else self._axis_for_spec_with_key(spec, axis_key)
        V = np.asarray(axis_vals, dtype=float)
        if not V.size:
            return
        scale = 1000.0 if (axis_unit or "").lower() == "v" else 1.0
        x_dense = np.linspace(np.nanmin(V), np.nanmax(V), 400)
        self.ax.plot(x_dense * scale, res['func'](x_dense) + offset, '--', color=color, lw=1.2)
        v0 = res.get('v0'); v0_err = res.get('v0_err')
        if v0 is not None and np.isfinite(v0):
            x_plot = v0 * scale
            y_plot = res['func'](v0) + offset
            xerr = v0_err * scale if v0_err is not None else None
            self.ax.axvline(x_plot, color=color, linestyle='--', alpha=0.85, lw=1.0, dashes=(4, 3))
            self.ax.errorbar([x_plot], [y_plot], xerr=[xerr] if xerr is not None else None,
                             fmt='o', color=color, ecolor=color, capsize=3,
                             markeredgecolor='black', markeredgewidth=0.8, markersize=5, markerfacecolor=color)
            axis_unit_clean = axis_unit or ""
            display_unit = axis_unit_clean
            if scale == 1000.0 and axis_unit_clean.lower() == "v":
                display_unit = "mV"
            elif not display_unit:
                display_unit = "arb"
            self._lcpd_line_info[spec_id] = {
                "x": x_plot,
                "display_unit": display_unit,
                "axis_unit": axis_unit_clean,
                "color": color,
                "spec_id": spec_id,
                "display_name": self._display_name(spec),
            }

    def _spec_id_by_name(self, name):
        for spec in self.specs:
            if self._display_name(spec) == name:
                return self._spec_id(spec)
        return None

    def _on_legend_pick(self, event):
        spec_id = self._legend_map.get(event.artist)
        if not spec_id:
            return
        line = self._line_map.get(spec_id)
        if not line:
            return
        visible = not line.get_visible()
        line.set_visible(visible)
        event.artist.set_alpha(1.0 if visible else 0.2)
        self.canvas.draw_idle()

    def _update_status(self, plotted=None):
        root = self.spec_list.invisibleRootItem()
        total = 0
        for i in range(root.childCount()):
            if not root.child(i).isHidden():
                total += 1
        checked = len(self._checked_items())
        text = f"{checked} selected / {total} total"
        if plotted is not None:
            text += f" | showing {plotted}"
        bg_txt = "BG set" if self._background_spec_id else "No BG"
        mode_txt = "Relative" if self._relative_zero_enabled else "Absolute"
        text += f" | {bg_txt} | {mode_txt}"
        self.status_label.setText(text)

    def _show_popup_for_spec(self, spec):
        dlg = SpectroscopyPopup(spec, parent=self)
        dlg.show()
        self._popup_refs.append(dlg)

    def _on_item_double_clicked(self, item):
        self._show_popup_for_spec(item.data(0, QtCore.Qt.UserRole))

    def _on_list_context_menu(self, pos):
        item = self.spec_list.itemAt(pos)
        if not item:
            return
        menu = QtWidgets.QMenu(self)
        act = menu.addAction("Open popup")
        copy_act = menu.addAction("Copy selected to clipboard")
        chosen = menu.exec_(self.spec_list.mapToGlobal(pos))
        if chosen == act:
            self._show_popup_for_spec(item.data(0, QtCore.Qt.UserRole))
        elif chosen == copy_act:
            self._copy_selected_to_clipboard()

    def _on_table_context_menu(self, pos):
        row = self.results_table.indexAt(pos).row()
        if row < 0:
            return
        spec_id = self.results_table.item(row,0).data(QtCore.Qt.UserRole)
        menu = QtWidgets.QMenu(self)
        act = menu.addAction("Open popup")
        copy_act = menu.addAction("Copy selected to clipboard")
        copy_table_act = menu.addAction("Copy table")
        clear_sel_act = menu.addAction("Clear selected")
        clear_all_act = menu.addAction("Clear all")
        chosen = menu.exec_(self.results_table.mapToGlobal(pos))
        if chosen == act:
            spec = self._spec_by_id(spec_id)
            if spec:
                self._show_popup_for_spec(spec)
        elif chosen == copy_act:
            self._copy_selected_to_clipboard()
        elif chosen == copy_table_act:
            self._copy_table_to_clipboard()
        elif chosen == clear_sel_act:
            self._clear_selected()
        elif chosen == clear_all_act:
            self._clear_all()

    def _on_table_double_clicked(self, item):
        spec_id = self.results_table.item(item.row(),0).data(QtCore.Qt.UserRole)
        spec = self._spec_by_id(spec_id)
        if spec:
            self._show_popup_for_spec(spec)

    def _on_table_selection(self):
        row = self.results_table.currentRow()
        if row < 0:
            return
        spec_id = self.results_table.item(row,0).data(QtCore.Qt.UserRole)
        item = self._item_map.get(spec_id)
        if item:
            self.spec_list.setCurrentItem(item)
            self._update_plot()

    def _copy_selected_to_clipboard(self):
        channel = self.channel_combo.currentText()
        if not channel:
            return
        items = self._selected_items() or self._checked_items()
        if not items:
            return
        blocks = []
        for it in items:
            spec = it.data(0, QtCore.Qt.UserRole)
            if not spec:
                continue
            axis_vals, _, axis_unit = self._axis_for_spec(spec)
            ch = np.asarray((spec.get('channels') or {}).get(channel, []), dtype=float)
            if axis_vals.size == 0 or ch.size == 0:
                continue
            unit_map = spec.get('unit_map') or {}
            unit = unit_map.get(channel, "")
            header_unit = f" ({unit})" if unit else ""
            axis_label = axis_unit or "arb"
            block = []
            block.append(f"# {Path(spec.get('path','')).name}  ({spec.get('x','?')}/{spec.get('y','?')} nm)")
            block.append(f"Bias ({axis_label})\t{channel}{header_unit}")
            for v, val in zip(axis_vals, ch):
                try:
                    block.append(f"{float(v):.9g}\t{float(val):.9g}")
                except Exception:
                    block.append(f"{v}\t{val}")
            blocks.append("\n".join(block))
        if blocks:
            QtWidgets.QApplication.clipboard().setText("\n\n".join(blocks))
            QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), "Copied spectra", self)

    def _spec_by_id(self, spec_id):
        for spec in self.specs:
            if self._spec_id(spec) == spec_id:
                return spec
        return None

    def _copy_table_to_clipboard(self):
        rows = []
        headers = ["File","X (nm)","Y (nm)","a","da","LCPD","dLCPD","c (Hz)","dc","RMSE"]
        rows.append("\t".join(headers))
        for r in range(self.results_table.rowCount()):
            vals = []
            for c in range(self.results_table.columnCount()):
                item = self.results_table.item(r, c)
                vals.append(item.text() if item else "")
            rows.append("\t".join(vals))
        if len(rows) > 1:
            QtWidgets.QApplication.clipboard().setText("\n".join(rows))
            QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), "Copied table", self)

    def _on_compare_canvas_menu(self, pos):
        menu = QtWidgets.QMenu(self)
        copy_png = menu.addAction("Copy plot as PNG")
        copy_svg = menu.addAction("Copy plot as SVG")
        menu.addSeparator()
        save_png = menu.addAction("Save PNG...")
        save_svg = menu.addAction("Save SVG...")
        action = menu.exec_(self.canvas.mapToGlobal(pos))
        if action == copy_png:
            self._copy_canvas_to_clipboard("png")
        elif action == copy_svg:
            self._copy_canvas_to_clipboard("svg")
        elif action == save_png:
            self._save_canvas("png")
        elif action == save_svg:
            self._save_canvas("svg")

    def _on_compare_canvas_keypress(self, event):
        if not event or not hasattr(event, "key"):
            return
        key = (event.key or "").lower()
        if key in ("ctrl+z", "control+z"):
            self._undo_last_action()
            gui_event = getattr(event, "guiEvent", None)
            if gui_event:
                gui_event.accept()
    def _copy_canvas_to_clipboard(self, fmt):
        buf = io.BytesIO()
        if fmt == "svg":
            with matplotlib.rc_context({'svg.fonttype': 'none'}):
                self.fig.savefig(buf, format="svg", bbox_inches="tight")
            mime = QtCore.QMimeData()
            mime.setData("image/svg+xml", buf.getvalue())
            QtWidgets.QApplication.clipboard().setMimeData(mime)
            QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), "Plot copied as SVG", self)
        else:
            self.fig.savefig(buf, format="png", dpi=300, bbox_inches="tight")
            image = QtGui.QImage.fromData(buf.getvalue(), "PNG")
            QtWidgets.QApplication.clipboard().setImage(image)
            QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), "Plot copied as PNG", self)

    def _save_canvas(self, fmt):
        if fmt == "png":
            filt = "PNG Files (*.png)"
            default = "spectra.png"
        else:
            filt = "SVG Files (*.svg)"
            default = "spectra.svg"
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save plot", default, filt)
        if not path:
            return
        try:
            self.fig.savefig(path, format=fmt, dpi=300, bbox_inches="tight")
            QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), f"Saved {Path(path).name}", self)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Save plot", str(exc))

    def _set_hint_text(self, text=None):
        label = getattr(self, "hint_label", None)
        if label:
            label.setText(text or self._delta_hint_text)

    def _on_compare_canvas_click(self, event):
        if not event or event.button != MouseButton.LEFT or event.inaxes != self.ax:
            return
        shift_pressed = False
        gui_event = getattr(event, "guiEvent", None)
        if gui_event is not None and hasattr(gui_event, "modifiers"):
            shift_pressed = bool(gui_event.modifiers() & QtCore.Qt.ShiftModifier)
        else:
            key = getattr(event, "key", "")
            if key and "shift" in str(key).lower():
                shift_pressed = True
        if not shift_pressed or event.xdata is None:
            return
        candidate = self._find_nearest_lcpd_line(event.xdata)
        if not candidate:
            return
        spec_id, info = candidate
        if not self._delta_selection:
            self._delta_selection = [info]
            self._set_hint_text("Shift+click a second LCPD line to annotate ΔLCPD.")
            return
        first = self._delta_selection[0]
        if info["spec_id"] == first["spec_id"]:
            self._delta_selection = [info]
            self._set_hint_text("Pick a different LCPD line and Shift+click to measure ΔLCPD.")
            return
        self._create_delta_annotation(first, info)
        self._delta_selection = []

    def _on_compare_canvas_motion(self, event):
        hovered = None
        if event and event.inaxes == self.ax and event.xdata is not None:
            hovered = self._find_nearest_lcpd_line(event.xdata)
        if self._delta_selection:
            if hovered:
                info = hovered[1]
                self._set_hint_text(
                    f"Shift+click {info.get('display_name', 'the line')} to finish ΔLCPD."
                )
                self.canvas.setCursor(QtCore.Qt.PointingHandCursor)
            else:
                self._set_hint_text("Shift+click a second LCPD line to annotate ΔLCPD.")
                self.canvas.setCursor(QtCore.Qt.ArrowCursor)
            return
        if hovered:
            info = hovered[1]
            self._set_hint_text(
                f"Shift+click {info.get('display_name', 'this LCPD')} to tag it for ΔLCPD."
            )
            self.canvas.setCursor(QtCore.Qt.PointingHandCursor)
        else:
            self._set_hint_text()
            self.canvas.setCursor(QtCore.Qt.ArrowCursor)

    def _find_nearest_lcpd_line(self, x_val):
        if not self._lcpd_line_info:
            return None
        xlim = self.ax.get_xlim()
        if not all(np.isfinite(val) for val in xlim):
            return None
        span = abs(xlim[1] - xlim[0])
        tol = max(span * 0.02, 1e-6)
        best = None
        for spec_id, info in self._lcpd_line_info.items():
            dist = abs(info["x"] - x_val)
            if dist <= tol and (best is None or dist < best[0]):
                best = (dist, spec_id, info)
        return (best[1], best[2]) if best else None

    def _clear_delta_annotation(self, redraw=True):
        for art in getattr(self, "_delta_annotation_artists", []):
            try:
                art.remove()
            except Exception:
                pass
        self._delta_annotation_artists = []
        if redraw:
            self.canvas.draw_idle()

    def _clear_delta_selection(self, redraw=True):
        self._delta_selection = []
        self._clear_delta_annotation(redraw=redraw)
        self._set_hint_text()

    def _create_delta_annotation(self, first, second):
        self._clear_delta_annotation(redraw=False)
        x1 = first["x"]
        x2 = second["x"]
        y_lower, y_upper = sorted(self.ax.get_ylim())
        span = y_upper - y_lower
        gap = max(0.04 * span if span else 1.0, 0.05)
        height = y_upper - gap
        min_height = y_lower + (0.02 * (span or 1.0))
        if height < min_height:
            height = min_height
        text_offset = max(0.02 * (span or 1.0), 0.1)
        text_y = height + text_offset
        unit1 = (first.get("display_unit") or "").strip()
        unit2 = (second.get("display_unit") or "").strip()
        if unit1 == unit2:
            unit_label = unit1 or "arb"
        else:
            unit_label = unit1 or unit2 or "arb"
        delta = abs(x2 - x1)
        delta_text = f"ΔLCPD = {delta:.3g} {unit_label}"
        arrowprops = dict(arrowstyle="<->", color="black", linewidth=1.0, shrinkA=0, shrinkB=0)
        annotation_line = self.ax.annotate(
            "",
            xy=(max(x1, x2), height),
            xytext=(min(x1, x2), height),
            arrowprops=arrowprops,
            clip_on=False,
            zorder=10,
        )
        text_artist = self.ax.text(
            0.5 * (x1 + x2),
            text_y,
            delta_text,
            ha="center",
            va="bottom",
            fontsize=8 * getattr(self, "_font_scale", 1.0),
            bbox=dict(facecolor="white", edgecolor="black", linewidth=0.6, boxstyle="round,pad=0.3", alpha=0.9),
            clip_on=False,
            zorder=10,
        )
        self._delta_annotation_artists = [annotation_line, text_artist]
        self.canvas.draw_idle()
        self._set_hint_text()

    def _on_visual_toggle(self, checked):
        sender = self.sender()
        label = sender.text() if sender else "Visual toggle"
        self._record_user_action(f"{label} → {'on' if checked else 'off'}")
        self._update_plot()

    def _on_offset_changed(self, value):
        self._record_user_action(f"Waterfall offset → {value:.3g}")
        self._update_plot()

    def _undo_last_action(self):
        if not self._undo_stack:
            return
        desc, state = self._undo_stack.pop()
        self._apply_state(state)
        self._set_hint_text(f"Reverted: {desc}")
        if hasattr(self, "undo_btn"):
            self.undo_btn.setEnabled(bool(self._undo_stack))

    def _record_user_action(self, desc):
        if self._suppress_undo_push:
            return
        state = self._snapshot_state()
        if not state:
            return
        if self._undo_stack and self._undo_stack[-1][1] == state:
            return
        self._undo_stack.append((desc, state))
        if len(self._undo_stack) > 30:
            self._undo_stack.pop(0)
        if hasattr(self, "undo_btn"):
            self.undo_btn.setEnabled(True)

    def _snapshot_state(self):
        checked = []
        selected = []
        root = self.spec_list.invisibleRootItem()
        for i in range(root.childCount()):
            item = root.child(i)
            spec_id = item.data(0, QtCore.Qt.UserRole + 1)
            if spec_id:
                if item.checkState(0) == QtCore.Qt.Checked:
                    checked.append(spec_id)
                if item.isSelected():
                    selected.append(spec_id)
        return {
            "channel": self.channel_combo.currentText(),
            "axis_key": self.axis_combo.currentData(),
            "waterfall": self.waterfall_cb.isChecked(),
            "show_points": self.show_points_cb.isChecked(),
            "show_lines": self.lines_cb.isChecked(),
            "offset": float(self.offset_spin.value()),
            "relative": self._relative_zero_enabled,
            "background": self._background_spec_id,
            "checked": checked,
            "selected": selected,
        }

    def _apply_state(self, state):
        if not state:
            return
        self._suppress_undo_push = True
        try:
            self.channel_combo.blockSignals(True)
            target = state.get("channel") or ""
            idx = self.channel_combo.findText(target)
            if idx >= 0:
                self.channel_combo.setCurrentIndex(idx)
            self.channel_combo.blockSignals(False)

            axis_target = state.get("axis_key")
            idx = self.axis_combo.findData(axis_target)
            if idx >= 0:
                self.axis_combo.blockSignals(True)
                self.axis_combo.setCurrentIndex(idx)
                self.axis_combo.blockSignals(False)

            for checkbox, key in (
                (self.waterfall_cb, "waterfall"),
                (self.show_points_cb, "show_points"),
                (self.lines_cb, "show_lines"),
                (self.relative_cb, "relative"),
            ):
                checkbox.blockSignals(True)
                checkbox.setChecked(bool(state.get(key)))
                checkbox.blockSignals(False)

            self._relative_zero_enabled = bool(state.get("relative"))

            self.offset_spin.blockSignals(True)
            self.offset_spin.setValue(state.get("offset", 0.0))
            self.offset_spin.blockSignals(False)

            self._background_spec_id = state.get("background")

            self._set_selection_state(state.get("checked", []), state.get("selected", []))

            self._update_plot()
        finally:
            self._suppress_undo_push = False

    def _set_selection_state(self, checked_ids, selected_ids):
        root = self.spec_list.invisibleRootItem()
        self.spec_list.blockSignals(True)
        try:
            checked_set = set(checked_ids)
            selected_set = set(selected_ids)
            for i in range(root.childCount()):
                item = root.child(i)
                spec_id = item.data(0, QtCore.Qt.UserRole + 1)
                if spec_id:
                    item.setCheckState(0, QtCore.Qt.Checked if spec_id in checked_set else QtCore.Qt.Unchecked)
                    item.setSelected(spec_id in selected_set)
        finally:
            self.spec_list.blockSignals(False)

    def _clear_selected(self):
        self._record_user_action("Clear selected spectra")
        removed = False
        for item in list(self._selected_items()):
            spec_id = item.data(0, QtCore.Qt.UserRole + 1)
            if spec_id in self._fit_results:
                self._fit_results.pop(spec_id, None)
            row = self.spec_list.indexOfTopLevelItem(item)
            self.spec_list.takeItem(row)
            removed = True
        if removed:
            self._update_plot()
            self._populate_results_table()
            self._update_status()

    def _clear_all(self):
        self._record_user_action("Clear all spectra")
        self.spec_list.clear()
        self._item_map = {}
        self._line_map.clear()
        self._legend_map.clear()
        self._fit_results = {}
        self.ax.clear()
        self.canvas.draw_idle()
        self.results_table.setRowCount(0)
        self._update_status(0)
        # also clear selection in parent so reopening starts fresh
        try:
            parent = self.parent()
            if parent and hasattr(parent, '_clear_multi_spec_selection'):
                parent._clear_multi_spec_selection()
        except Exception:
            pass

    def _select_all_visible(self):
        """Select all visible (non-filtered) spectra."""
        self._record_user_action("Select all visible spectra")
        root = self.spec_list.invisibleRootItem()
        for i in range(root.childCount()):
            item = root.child(i)
            if not item.isHidden():
                item.setCheckState(0, QtCore.Qt.Checked)
        self._update_plot()

    def _invert_selection(self):
        """Invert the checked state of all visible spectra."""
        self._record_user_action("Invert selection")
        root = self.spec_list.invisibleRootItem()
        for i in range(root.childCount()):
            item = root.child(i)
            if not item.isHidden():
                current_state = item.checkState(0)
                new_state = QtCore.Qt.Unchecked if current_state == QtCore.Qt.Checked else QtCore.Qt.Checked
                item.setCheckState(0, new_state)
        self._update_plot()

    def _fit_selected(self):
        items = self._selected_items() or self._checked_items()
        self._start_fit([item.data(0, QtCore.Qt.UserRole) for item in items])

    def _fit_all(self):
        self._start_fit([item.data(0, QtCore.Qt.UserRole) for item in self._checked_items()])

    def _start_fit(self, specs):
        if not specs or self._fit_thread:
            if not specs:
                self._log("Nothing to fit.")
            return
        channel = self.channel_combo.currentText()
        self._set_busy(True, f"Fitting {len(specs)} spectra...")
        axis_key = self.axis_combo.currentData()
        self._fit_worker = _SpectroFitWorker(specs, channel, axis_key)
        self._fit_thread = QtCore.QThread(self)
        self._fit_worker.moveToThread(self._fit_thread)
        self._fit_thread.started.connect(self._fit_worker.run)
        self._fit_worker.finished.connect(self._on_fit_finished)
        self._fit_worker.progress.connect(self._on_fit_progress)
        self._fit_worker.finished.connect(self._fit_thread.quit)
        self._fit_thread.finished.connect(self._cleanup_fit_thread)
        self._fit_thread.start()

    def _cleanup_fit_thread(self):
        self._fit_thread.deleteLater()
        self._fit_thread = None
        self._fit_worker = None
        self._set_busy(False, "Fit ready.")

    def _on_fit_finished(self, results, logs):
        for msg in logs:
            self._log(msg)
        for res in results:
            spec = res.get('spec')
            if spec:
                self._fit_results[self._spec_id(spec)] = res
        self._populate_results_table()
        self._update_plot()

    def _on_fit_progress(self, percentage, message):
        self.progress_bar.setValue(percentage)
        self.status_label.setText(message)

    def _populate_results_table(self):
        rows = []
        for spec_id, res in self._fit_results.items():
            spec = res.get('spec')
            if not spec:
                continue
            xs = spec.get('x')
            ys = spec.get('y')
            axis_unit = res.get('axis_unit') or ''
            scale = 1000.0 if axis_unit.lower() == "v" else 1.0
            v0 = res.get('v0')
            v0_err = res.get('v0_err')
            v0_disp = "n/a"
            v0_err_disp = "n/a"
            if v0 is not None and np.isfinite(v0):
                v0_disp = f"{v0 * scale:.4g}"
                if v0_err is not None and np.isfinite(v0_err):
                    v0_err_disp = f"{v0_err * scale:.3g}"
            rows.append((spec_id, self._display_name(spec),
                         "n/a" if xs is None else f"{xs:.1f}",
                         "n/a" if ys is None else f"{ys:.1f}",
                         f"{res['a']:.4g}", f"{res['a_err']:.2g}",
                         v0_disp, v0_err_disp,
                         f"{res['c']:.4g}", f"{res['c_err']:.2g}",
                         f"{res['rmse']:.4g}"))
        self.results_table.setRowCount(len(rows))
        for r, data in enumerate(rows):
            spec_id, name, xval, yval, a, ae, b, be, c, ce, rmse = data
            values = [name, xval, yval, a, ae, b, be, c, ce, rmse]
            for col, val in enumerate(values):
                item = QtWidgets.QTableWidgetItem(val)
                if col == 0:
                    item.setData(QtCore.Qt.UserRole, spec_id)
                self.results_table.setItem(r, col, item)

    def _export_csv(self):
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Export CSV", "spectroscopy_fit.csv", "CSV Files (*.csv)")
        if not path:
            return
        headers = ["File","X (nm)","Y (nm)","a","da","b (mV)","db","c (Hz)","dc","RMSE"]
        with open(path, 'w', newline='') as f:
            f.write(",".join(headers) + "\n")
            for row in range(self.results_table.rowCount()):
                vals = [self.results_table.item(row, col).text() if self.results_table.item(row, col) else ""
                        for col in range(self.results_table.columnCount())]
                f.write(",".join(vals) + "\n")
        self._log(f"Exported to {path}")

    def _set_busy(self, busy, message):
        self.fit_selected_btn.setEnabled(not busy)
        self.fit_all_btn.setEnabled(not busy)
        self.export_btn.setEnabled(not busy)
        self.progress_bar.setVisible(busy)
        if busy:
            self.status_label.setText(message)
            self.progress_bar.setValue(0)
        else:
            self.progress_bar.setVisible(False)

    def _on_options_toggled(self, checked):
        self.options_toggle.setArrowType(QtCore.Qt.DownArrow if checked else QtCore.Qt.RightArrow)
        self.options_body.setVisible(checked)

    def _show_help(self):
        """Show help dialog for spectroscopy comparison features."""
        help_text = """
        <h2>Spectroscopy Comparison Help</h2>
        
        <h3>Getting Started</h3>
        <p>Use the spectrum list on the left to select which spectra to compare. Check the boxes to include spectra in the plot, or select items for additional operations.</p>
        
        <h3>Data Selection</h3>
        <ul>
        <li><b>Channel:</b> Choose which data channel to plot and analyze</li>
        <li><b>Axis:</b> Select the X-axis (bias voltage or Z position)</li>
        <li><b>Relative Z:</b> Shift Z-axis to start from zero at minimum value</li>
        </ul>
        
        <h3>Visualization</h3>
        <ul>
        <li><b>Color Cycle:</b> Select color palette for multiple spectra</li>
        <li><b>Waterfall:</b> Stack spectra vertically with offset</li>
        <li><b>Offset:</b> Adjust vertical spacing in waterfall mode</li>
        <li><b>Lines/Points:</b> Use the Lines toggle to hide the smooth curves and Points to show the raw markers.</li>
        </ul>

        <h3>Interactions</h3>
        <ul>
        <li><b>Shift+Click:</b> Click two LCPD guide lines while holding Shift to draw a ΔLCPD annotation between them.</li>
        </ul>
        
        <h3>Analysis</h3>
        <h4>KPFM</h4>
        <ul>
        <li><b>Fit Selected/All:</b> Perform parabolic fits on spectra</li>
        <li><b>Export CSV:</b> Save fit results to CSV file</li>
        </ul>
        <h4>Forces/Background</h4>
        <ul>
        <li><b>Set/Clear Background:</b> Subtract background spectrum</li>
        <li><b>Convert to Force:</b> Experimental force curve conversion</li>
        </ul>
        
        <h3>Actions</h3>
        <ul>
        <li><b>Copy:</b> Copy data to clipboard</li>
        <li><b>Export:</b> Save results to CSV</li>
        <li><b>Clear:</b> Remove spectra from list</li>
        </ul>
        
        <h3>Keyboard Shortcuts</h3>
        <ul>
        <li><b>F:</b> Fit selected spectra</li>
        <li><b>Ctrl+E:</b> Export to CSV</li>
        <li><b>Ctrl+A:</b> Select all visible spectra</li>
        <li><b>Ctrl+Shift+A:</b> Invert selection</li>
        <li><b>Delete:</b> Clear selected spectra</li>
        <li><b>Ctrl+Delete:</b> Clear all spectra</li>
        </ul>
        """
        
        dialog = QtWidgets.QDialog(self)
        dialog.setWindowTitle("Spectroscopy Comparison Help")
        dialog.resize(600, 500)
        
        layout = QtWidgets.QVBoxLayout(dialog)
        text_edit = QtWidgets.QTextEdit()
        text_edit.setHtml(help_text)
        text_edit.setReadOnly(True)
        layout.addWidget(text_edit)
        
        buttons = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok)
        buttons.accepted.connect(dialog.accept)
        layout.addWidget(buttons)
        
        dialog.exec_()

    def wheelEvent(self, event):
        try:
            modifiers = event.modifiers()
        except Exception:
            modifiers = QtCore.Qt.NoModifier
        if modifiers & QtCore.Qt.ControlModifier:
            angle = event.angleDelta().y() if hasattr(event, 'angleDelta') else 0
            if angle:
                step = 0.05 * (1 if angle > 0 else -1)
                self._font_scale = min(2.5, max(0.6, self._font_scale + step))
                self._apply_font_scale()
            event.accept()
            return
        super().wheelEvent(event)

    def _apply_font_scale(self):
        scale = getattr(self, '_font_scale', 1.0)
        self.ax.tick_params(labelsize=8 * scale)
        self.ax.xaxis.label.set_fontsize(10 * scale)
        self.ax.yaxis.label.set_fontsize(10 * scale)
        if self.ax.get_legend():
            plt_legend = self.ax.get_legend()
            for text in plt_legend.get_texts():
                text.set_fontsize(8 * scale)
        self.canvas.draw_idle()

    def _log(self, text):
        self.log.appendPlainText(text)
