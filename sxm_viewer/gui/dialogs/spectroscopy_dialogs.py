"""Detail canvases and spectroscopy dialogs."""
from __future__ import annotations

import itertools
import json
import math

import numpy as np
from matplotlib import patches
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
from .matrix_fit import MatrixFitDialog

class SpectroscopyPopup(QtWidgets.QDialog):
    """Popup window showing spectroscopy curves for a given file."""
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
        layout.addWidget(self.canvas, 1)
        self.canvas.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.canvas.customContextMenuRequested.connect(self._on_canvas_context_menu)
        self.fit_result_label = QtWidgets.QLabel("")
        self.fit_result_label.setWordWrap(True)
        layout.addWidget(self.fit_result_label)
        self.setLayout(layout)

        self.V = np.asarray(spec.get('V', []), dtype=float)
        self.channels = {name: np.asarray(vals, dtype=float) for name, vals in (spec.get('channels', {}) or {}).items()}
        for name in self.channels.keys():
            self.channel_combo.addItem(name)
        self.channel_combo.currentTextChanged.connect(self._on_channel_changed)
        self.fit_btn.clicked.connect(self._on_fit_clicked)
        self.copy_btn.clicked.connect(self._copy_channel_to_clipboard)
        self._last_fit_result = None
        if self.channel_combo.count():
            self.channel_combo.setCurrentIndex(0)
            self._plot_selected_channel()
        else:
            self.ax.text(0.5, 0.5, "No channels", ha='center', va='center', transform=self.ax.transAxes)
            self.canvas.draw()
        self._update_fit_button()

    def _channel_label_with_unit(self, name):
        base = name or ""
        unit = self.spec.get('unit_map', {}).get(name)
        if not unit and '(' in base and base.endswith(')'):
            return base
        if unit:
            return f"{base} ({unit})"
        return base

    def _on_channel_changed(self, name):
        self._last_fit_result = None
        self.fit_result_label.setText("")
        self._plot_selected_channel()
        self._update_fit_button()

    def _plot_selected_channel(self):
        self.ax.clear()
        name = self.channel_combo.currentText()
        if not name or name not in self.channels or not self.V.size:
            self.canvas.draw_idle()
            return
        bias_mv = self.V * 1000.0
        self.ax.plot(bias_mv, self.channels[name], color='#c94cfa', lw=1.5, label='Data')
        self.ax.set_xlabel("Bias (mV)")
        self.ax.set_ylabel(self._channel_label_with_unit(name))
        self.ax.grid(True, alpha=0.2)
        if self._last_fit_result and self._last_fit_result.get('channel') == name:
            self._draw_fit_overlay(self._last_fit_result)
        else:
            handles, labels = self.ax.get_legend_handles_labels()
            if handles:
                self.ax.legend()
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
            f"Bias (mV)\t{self._channel_label_with_unit(name)}"
        ]
        for v, val in zip(bias, values):
            try:
                lines.append(f"{float(v) * 1000.0:.9g}\t{float(val):.9g}")
            except Exception:
                lines.append(f"{v}\t{val}")
        QtWidgets.QApplication.clipboard().setText("\n".join(lines))
        QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), "Spectroscopy copied", self)

    def _draw_fit_overlay(self, res):
        if not self.V.size:
            return
        x_dense = np.linspace(np.nanmin(self.V), np.nanmax(self.V), 400)
        y_dense = res['func'](x_dense)
        self.ax.plot(x_dense * 1000.0, y_dense, '--', color='#ff8c00', lw=1.5, label='Fit')
        b = res['b']; c = res['c']; b_err = res.get('b_err', 0.0)
        self.ax.errorbar([b * 1000.0], [c], xerr=[b_err * 1000.0], fmt='o', color='#004c99', ecolor='#004c99', capsize=4, label='LCPD')
        self.ax.legend()
        text = (
            f"a = {res['a']:.4g} +/- {res['a_err']:.2g}\n"
            f"b (LCPD) = {res['b']:.2f} +/- {res['b_err']:.2f} mV\n"
            f"c = {res['c']:.4g} +/- {res['c_err']:.2g} Hz\n"
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
        self._last_fit_result = res
        self._plot_selected_channel()

    def _update_fit_button(self):
        enable = bool(self.channel_combo.count() and self.V.size)
        self.fit_btn.setEnabled(enable)

class MatrixSpectroViewer(QtWidgets.QDialog):
    def __init__(self, parent, image_entry, specs):
        super().__init__(parent)
        self.image_entry = image_entry
        self.specs = list(specs)
        self.viewer = parent
        self.setWindowTitle(f"Matrix Spectroscopies - {Path(image_entry['path']).name}")
        self.resize(900, 700)
        layout = QtWidgets.QVBoxLayout()
        self.canvas = FigureCanvas(Figure(figsize=(6,6)))
        layout.addWidget(self.canvas, 1)
        self.ax = self.canvas.figure.add_subplot(111)
        self.image_value_label = QtWidgets.QLabel("Value: --")
        layout.addWidget(self.image_value_label)
        controls = QtWidgets.QHBoxLayout()
        controls.addWidget(QtWidgets.QLabel("Image channel:"))
        self.channel_combo = QtWidgets.QComboBox()
        controls.addWidget(self.channel_combo, 1)
        self.fit_matrix_btn = QtWidgets.QPushButton("Fit matrix parabolas...")
        controls.addWidget(self.fit_matrix_btn)
        layout.addLayout(controls)
        self.info_label = QtWidgets.QLabel("Click a point to open its spectroscopy")
        layout.addWidget(self.info_label)
        layout.addSpacing(6)
        self.setLayout(layout)
        self.canvas.mpl_connect("button_press_event", self._on_click)
        self.canvas.mpl_connect("motion_notify_event", self._on_canvas_hover)
        self._fit_dialogs = []
        self._current_image_arr = None
        self._current_image_extent = None
        self._current_image_unit = ''
        self._populate_channels()
        self.channel_combo.currentIndexChanged.connect(self._draw)
        self.fit_matrix_btn.clicked.connect(self._on_fit_matrix)
        self._draw()

    def _populate_channels(self):
        self.channel_combo.clear()
        path = Path(self.image_entry['path'])
        header, fds = self.viewer.headers.get(str(path), (None, None))
        if not fds:
            return
        for idx, fd in enumerate(fds):
            name = fd.get('Caption', fd.get('FileName', f"chan{idx}"))
            scale = fd.get('Scale')
            offset = fd.get('Offset')
            unit = fd.get('PhysUnit', '')
            self.channel_combo.addItem(f"{idx}: {name}", (idx, scale, offset, unit))
        if self.viewer.last_preview and self.viewer.last_preview[0] == str(path):
            self.channel_combo.setCurrentIndex(int(self.viewer.last_preview[1]))
        else:
            self.channel_combo.setCurrentIndex(0)

    def _draw(self):
        path = Path(self.image_entry['path'])
        header, fds = self.viewer.headers.get(str(path), (None, None))
        data = self.channel_combo.currentData()
        if data:
            main_idx = int(data[0])
        else:
            main_idx = 0
        try:
            if header and fds and 0 <= main_idx < len(fds):
                fd = fds[main_idx]
                arr = self.viewer._get_channel_array(str(path), main_idx, header, fd)
                arr = np.asarray(arr, dtype=float)
                self.ax.imshow(arr, cmap='gray', origin='upper')
                self._current_image_arr = arr
                self._current_image_extent = None
                self._current_image_unit = fd.get('PhysUnit', '')
            else:
                self.ax.text(0.5, 0.5, Path(path).name, ha='center', va='center', transform=self.ax.transAxes)
                self._current_image_arr = None
        except Exception:
            self.ax.text(0.5, 0.5, Path(path).name, ha='center', va='center', transform=self.ax.transAxes)
            self._current_image_arr = None
        xs = []
        ys = []
        xpix = int(header.get('xPixel', 128) if header else 128)
        ypix = int(header.get('yPixel', 128) if header else 128)
        for spec in self.specs:
            coords = self.viewer._map_spec_to_pixels(spec, header or {}, xpix, ypix)
            if coords is None:
                continue
            col, row = coords
            xs.append(col)
            ys.append(row)
        if xs and ys:
            self.ax.scatter(xs, ys, s=30, c='red', alpha=0.8)
        self.canvas.draw_idle()
        if self._current_image_arr is None:
            self.image_value_label.setText("Value: --")

    def _pick_spec_from_point(self, x, y):
        best = None
        best_dist = None
        header, _ = self.viewer.headers.get(str(self.image_entry['path']), (None, None))
        xpix = int(header.get('xPixel', 128) if header else 128)
        ypix = int(header.get('yPixel', 128) if header else 128)
        for spec in self.specs:
            coords = self.viewer._map_spec_to_pixels(spec, header or {}, xpix, ypix)
            if coords is None:
                continue
            col, row = coords
            dist = (col - x)**2 + (row - y)**2
            if best is None or dist < best_dist:
                best = spec
                best_dist = dist
        return best

    def _on_click(self, event):
        if event.inaxes != self.ax:
            return
        spec = self._pick_spec_from_point(event.xdata, event.ydata)
        if not spec:
            return
        self.viewer._open_spectroscopy_popup(spec)

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

    def __init__(self, specs, channel):
        super().__init__()
        self.specs = list(specs)
        self.channel = channel

    def run(self):
        results = []
        logs = []
        for spec in self.specs:
            name = Path(spec['path']).name
            V = np.asarray(spec.get('V', []), dtype=float)
            channels = spec.get('channels') or {}
            data = channels.get(self.channel)
            if data is None or not V.size:
                logs.append(f"{name}: channel '{self.channel}' unavailable")
                continue
            try:
                res = fit_parabola_bias(V, data)
                res['spec'] = spec
                results.append(res)
                logs.append(f"{name}: fit ok (RMSE {res['rmse']:.3g})")
            except Exception as e:
                logs.append(f"{name}: {e}")
        self.finished.emit(results, logs)

class SpectroscopyCompareDialog(QtWidgets.QDialog):
    """Modern comparison UI for spectroscopy overlays and fitting."""
    def __init__(self, specs, parent=None):
        super().__init__(parent)
        self.specs = list(specs)
        self._line_map = {}
        self._legend_map = {}
        self._fit_results = {}
        self._fit_thread = None
        self._fit_worker = None
        self._popup_refs = []
        self.setWindowTitle("Spectroscopy comparison")
        self.resize(1250, 640)
        self._build_ui()
        self._populate_list()
        self._populate_channels()
        self._update_plot()

    def _spec_id(self, spec):
        base = str(Path(spec.get('path', '')))
        idx = spec.get('matrix_index')
        return f"{base}#m{idx}" if idx is not None else base

    def _display_name(self, spec):
        name = Path(spec.get('path', '')).name
        idx = spec.get('matrix_index')
        return f"{name} [m{idx}]" if idx is not None else name

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
        self.filter_edit.textChanged.connect(self._apply_filter)
        self.spec_list = QtWidgets.QListWidget()
        self.spec_list.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
        self.spec_list.itemChanged.connect(self._on_item_check_changed)
        self.spec_list.itemSelectionChanged.connect(self._on_list_selection_changed)
        self.spec_list.itemDoubleClicked.connect(self._on_item_double_clicked)
        self.spec_list.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.spec_list.customContextMenuRequested.connect(self._on_list_context_menu)
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
        self.status_label = QtWidgets.QLabel("0 selected / 0 total")
        center_layout.addWidget(self.status_label)
        splitter.addWidget(center)
        splitter.setStretchFactor(1, 2)
        self.canvas.mpl_connect('pick_event', self._on_legend_pick)

        # Right panel: controls + results
        right = QtWidgets.QWidget()
        right_layout = QtWidgets.QVBoxLayout(right)
        right_layout.setContentsMargins(6,6,6,6)
        channel_row = QtWidgets.QHBoxLayout()
        channel_row.addWidget(QtWidgets.QLabel("Channel:"))
        self.channel_combo = QtWidgets.QComboBox()
        self.channel_combo.currentTextChanged.connect(self._on_channel_changed)
        channel_row.addWidget(self.channel_combo, 1)
        right_layout.addLayout(channel_row)

        btn_row = QtWidgets.QHBoxLayout()
        self.fit_selected_btn = QtWidgets.QPushButton("Fit selected (F)")
        self.fit_all_btn = QtWidgets.QPushButton("Fit all")
        self.export_btn = QtWidgets.QPushButton("Export CSV")
        self.copy_btn = QtWidgets.QPushButton("Copy selected")
        self.copy_table_btn = QtWidgets.QPushButton("Copy table")
        self.clear_sel_btn = QtWidgets.QPushButton("Clear selected")
        self.clear_all_btn = QtWidgets.QPushButton("Clear all")
        btn_row.addWidget(self.fit_selected_btn)
        btn_row.addWidget(self.fit_all_btn)
        btn_row.addWidget(self.export_btn)
        btn_row.addWidget(self.copy_btn)
        btn_row.addWidget(self.copy_table_btn)
        btn_row.addWidget(self.clear_sel_btn)
        btn_row.addWidget(self.clear_all_btn)
        right_layout.addLayout(btn_row)
        self.fit_selected_btn.clicked.connect(self._fit_selected)
        self.fit_all_btn.clicked.connect(self._fit_all)
        self.export_btn.clicked.connect(self._export_csv)
        self.copy_btn.clicked.connect(self._copy_selected_to_clipboard)
        self.copy_table_btn.clicked.connect(self._copy_table_to_clipboard)
        self.clear_sel_btn.clicked.connect(self._clear_selected)
        self.clear_all_btn.clicked.connect(self._clear_all)

        QtWidgets.QShortcut(QtGui.QKeySequence("F"), self, activated=self._fit_selected)
        QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+E"), self, activated=self._export_csv)

        self.options_toggle = QtWidgets.QToolButton()
        self.options_toggle.setText("Fit options")
        self.options_toggle.setCheckable(True)
        self.options_toggle.setChecked(False)
        self.options_toggle.setArrowType(QtCore.Qt.RightArrow)
        self.options_toggle.setToolButtonStyle(QtCore.Qt.ToolButtonTextBesideIcon)
        self.options_toggle.toggled.connect(self._on_options_toggled)
        right_layout.addWidget(self.options_toggle)

        self.options_body = QtWidgets.QWidget()
        form = QtWidgets.QFormLayout(self.options_body)
        self.degree_spin = QtWidgets.QSpinBox()
        self.degree_spin.setRange(2, 2)
        self.degree_spin.setValue(2)
        self.degree_spin.setEnabled(False)
        form.addRow("Degree", self.degree_spin)
        self.mask_min = QtWidgets.QDoubleSpinBox(); self.mask_min.setRange(-1e6, 1e6); self.mask_min.setSuffix(" V")
        self.mask_max = QtWidgets.QDoubleSpinBox(); self.mask_max.setRange(-1e6, 1e6); self.mask_max.setSuffix(" V")
        form.addRow("Mask min", self.mask_min)
        form.addRow("Mask max", self.mask_max)
        self.options_body.setVisible(False)
        right_layout.addWidget(self.options_body)

        self.results_table = QtWidgets.QTableWidget(0, 10)
        self.results_table.setHorizontalHeaderLabels(["File","X (nm)","Y (nm)","a","?a","b (mV)","?b","c (Hz)","?c","RMSE"])
        self.results_table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.results_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.results_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.results_table.itemSelectionChanged.connect(self._on_table_selection)
        self.results_table.itemDoubleClicked.connect(self._on_table_double_clicked)
        self.results_table.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.results_table.customContextMenuRequested.connect(self._on_table_context_menu)
        right_layout.addWidget(self.results_table, 1)

        self.log = QtWidgets.QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumHeight(100)
        right_layout.addWidget(self.log)

        splitter.addWidget(right)
        splitter.setStretchFactor(2, 1)

    def _populate_list(self):
        self.spec_list.blockSignals(True)
        self.spec_list.clear()
        self._item_map = {}
        for spec in self.specs:
            item = QtWidgets.QListWidgetItem(self._display_name(spec))
            item.setFlags(item.flags() | QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsSelectable)
            item.setCheckState(QtCore.Qt.Checked)
            item.setData(QtCore.Qt.UserRole, spec)
            item.setData(QtCore.Qt.UserRole + 1, self._spec_id(spec))
            self.spec_list.addItem(item)
            self._item_map[self._spec_id(spec)] = item
        self.spec_list.blockSignals(False)

    def _populate_channels(self):
        channels = sorted({name for spec in self.specs for name in (spec.get('channels') or {}).keys()})
        self.channel_combo.blockSignals(True)
        self.channel_combo.clear()
        for name in channels:
            self.channel_combo.addItem(name)
        if channels:
            self.channel_combo.setCurrentText('df' if 'df' in channels else channels[0])
        self.channel_combo.blockSignals(False)

    def _apply_filter(self, text):
        text = text.lower()
        for i in range(self.spec_list.count()):
            item = self.spec_list.item(i)
            item.setHidden(text not in item.text().lower())
        self._update_status()

    def _checked_items(self):
        return [self.spec_list.item(i) for i in range(self.spec_list.count())
                if self.spec_list.item(i).checkState() == QtCore.Qt.Checked and not self.spec_list.item(i).isHidden()]

    def _selected_items(self):
        return [item for item in self._checked_items() if item.isSelected()]

    def _on_channel_changed(self):
        self._fit_results = {}
        self._populate_results_table()
        self._update_plot()

    def _on_item_check_changed(self):
        self._update_plot()

    def _on_list_selection_changed(self):
        self._update_plot()

    def _update_plot(self):
        channel = self.channel_combo.currentText()
        self.ax.clear()
        self.ax.grid(True, alpha=0.2)
        self._line_map.clear()
        self._legend_map.clear()
        selected_ids = {item.data(QtCore.Qt.UserRole + 1) for item in self._selected_items()}
        colors = itertools.cycle(matplotlib.cm.get_cmap('tab10').colors)
        plotted = 0
        for item in self._checked_items():
            spec = item.data(QtCore.Qt.UserRole)
            spec_id = item.data(QtCore.Qt.UserRole + 1)
            channels = spec.get('channels') or {}
            data = channels.get(channel)
            V = np.asarray(spec.get('V', []), dtype=float)
            if data is None or not V.size:
                continue
            color = next(colors)
            highlight = spec_id in selected_ids or not selected_ids
            label_txt = self._display_name(spec)
            line, = self.ax.plot(V, data, color=color, lw=2.4 if highlight else 1.2,
                                 alpha=1.0 if highlight else 0.4, label=label_txt)
            self._line_map[spec_id] = line
            plotted += 1
            if spec_id in self._fit_results:
                self._draw_fit_for_spec(spec_id, color)
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
        self.ax.set_xlabel("Bias (mV)")
        # include unit if available
        unit = None
        # look up unit from any spec carrying this channel
        for item in self._checked_items():
            spec = item.data(QtCore.Qt.UserRole)
            if not spec:
                continue
            unit_map = spec.get('unit_map') or {}
            if channel in unit_map and unit_map[channel]:
                unit = unit_map[channel]
                break
        if unit:
            self.ax.set_ylabel(f"{channel} ({unit})")
        else:
            self.ax.set_ylabel(channel)
        self.canvas.draw_idle()
        self._update_status(plotted)

    def _draw_fit_for_spec(self, spec_id, color):
        res = self._fit_results.get(spec_id)
        if not res:
            return
        spec = res.get('spec')
        V = np.asarray(spec.get('V', []), dtype=float)
        if not V.size:
            return
        x_dense = np.linspace(np.nanmin(V), np.nanmax(V), 400)
        self.ax.plot(x_dense, res['func'](x_dense), '--', color=color, lw=1.2)
        b = res['b']; c = res['c']; be = res.get('b_err', 0.0)
        self.ax.errorbar([b], [c], xerr=[be], fmt='o', color=color, ecolor=color, capsize=3)

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
        total = sum(1 for i in range(self.spec_list.count()) if not self.spec_list.item(i).isHidden())
        checked = len(self._checked_items())
        text = f"{checked} selected / {total} total"
        if plotted is not None:
            text += f" ? showing {plotted}"
        self.status_label.setText(text)

    def _show_popup_for_spec(self, spec):
        dlg = SpectroscopyPopup(spec, parent=self)
        dlg.show()
        self._popup_refs.append(dlg)

    def _on_item_double_clicked(self, item):
        self._show_popup_for_spec(item.data(QtCore.Qt.UserRole))

    def _on_list_context_menu(self, pos):
        item = self.spec_list.itemAt(pos)
        if not item:
            return
        menu = QtWidgets.QMenu(self)
        act = menu.addAction("Open popup")
        copy_act = menu.addAction("Copy selected to clipboard")
        chosen = menu.exec_(self.spec_list.mapToGlobal(pos))
        if chosen == act:
            self._show_popup_for_spec(item.data(QtCore.Qt.UserRole))
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
            self.spec_list.setCurrentItem(item, QtCore.QItemSelectionModel.SelectCurrent)
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
            spec = it.data(QtCore.Qt.UserRole)
            if not spec:
                continue
            V = np.asarray(spec.get('V', []), dtype=float)
            ch = np.asarray((spec.get('channels') or {}).get(channel, []), dtype=float)
            if V.size == 0 or ch.size == 0:
                continue
            unit_map = spec.get('unit_map') or {}
            unit = unit_map.get(channel, "")
            header_unit = f" ({unit})" if unit else ""
            block = []
            block.append(f"# {Path(spec.get('path','')).name}  ({spec.get('x','?')}/{spec.get('y','?')} nm)")
            block.append(f"Bias (mV)\t{channel}{header_unit}")
            for v, val in zip(V * 1000.0, ch):
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
        headers = ["File","X (nm)","Y (nm)","a","da","b (mV)","db","c (Hz)","dc","RMSE"]
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

    def _clear_selected(self):
        removed = False
        for item in list(self._selected_items()):
            spec_id = item.data(QtCore.Qt.UserRole + 1)
            if spec_id in self._fit_results:
                self._fit_results.pop(spec_id, None)
            row = self.spec_list.row(item)
            self.spec_list.takeItem(row)
            removed = True
        if removed:
            self._update_plot()
            self._populate_results_table()
            self._update_status()

    def _clear_all(self):
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

    def _fit_selected(self):
        items = self._selected_items() or self._checked_items()
        self._start_fit([item.data(QtCore.Qt.UserRole) for item in items])

    def _fit_all(self):
        self._start_fit([item.data(QtCore.Qt.UserRole) for item in self._checked_items()])

    def _start_fit(self, specs):
        if not specs or self._fit_thread:
            if not specs:
                self._log("Nothing to fit.")
            return
        channel = self.channel_combo.currentText()
        self._set_busy(True, f"Fitting {len(specs)} spectra...")
        self._fit_worker = _SpectroFitWorker(specs, channel)
        self._fit_thread = QtCore.QThread(self)
        self._fit_worker.moveToThread(self._fit_thread)
        self._fit_thread.started.connect(self._fit_worker.run)
        self._fit_worker.finished.connect(self._on_fit_finished)
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

    def _populate_results_table(self):
        rows = []
        for spec_id, res in self._fit_results.items():
            spec = res.get('spec')
            if not spec:
                continue
            xs = spec.get('x')
            ys = spec.get('y')
            rows.append((spec_id, self._display_name(spec),
                         "n/a" if xs is None else f"{xs:.1f}",
                         "n/a" if ys is None else f"{ys:.1f}",
                         f"{res['a']:.4g}", f"{res['a_err']:.2g}",
                         f"{res['b']:.2f}", f"{res['b_err']:.2f}",
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
        if busy:
            self.status_label.setText(message)

    def _on_options_toggled(self, checked):
        self.options_toggle.setArrowType(QtCore.Qt.DownArrow if checked else QtCore.Qt.RightArrow)
        self.options_body.setVisible(checked)

    def _log(self, text):
        self.log.appendPlainText(text)




