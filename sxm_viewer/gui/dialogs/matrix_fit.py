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

class MatrixFitWorker(QtCore.QObject):
    progress = QtCore.pyqtSignal(int, int)
    finished = QtCore.pyqtSignal(object)

    def __init__(self, specs):
        super().__init__()
        self.specs = list(specs)

    @QtCore.pyqtSlot()
    def run(self):
        specs = self.specs
        if not specs:
            self.finished.emit({
                'maps': {},
                'logs': ["No spectra to fit"],
                'channel_name': "channel",
                'x_axis': None,
                'y_axis': None,
            })
            return
        def _pick_df_channel(spec_list):
            candidates = []
            for spec in spec_list:
                chans = (spec.get('channels') or {}).keys()
                for ch in chans:
                    candidates.append(str(ch))
            if not candidates:
                return None
            def _score(name):
                low = name.lower()
                score = 0
                if "df" in low or "deltaf" in low or "delf" in low:
                    score += 10
                if "kpfm" in low:
                    score += 6
                if "freq" in low or "frequency" in low:
                    score += 2
                return score
            scored = sorted([( -_score(n), n) for n in set(candidates)])
            best = scored[0][1] if scored else None
            return best

        channel_name = _pick_df_channel(specs) or 'channel'
        axis_unit = specs[0].get('AxisUnit') or "V"
        col_candidates = [spec.get('grid_col') for spec in specs if spec.get('grid_col') is not None]
        row_candidates = [spec.get('grid_row') for spec in specs if spec.get('grid_row') is not None]
        matrix_indices = [spec.get('matrix_index') for spec in specs if spec.get('matrix_index') is not None]
        grid_cols = grid_rows = None
        if col_candidates and row_candidates:
            grid_cols = max(col_candidates) + 1
            grid_rows = max(row_candidates) + 1
        else:
            if matrix_indices:
                min_idx = min(matrix_indices)
                max_idx = max(matrix_indices)
                # detect 1-based indexing and normalize for grid sizing
                if min_idx >= 1:
                    max_idx = max_idx - 1
                side = int(round(math.sqrt(max_idx + 1)))
                if side > 0:
                    grid_cols = grid_rows = side
        if not grid_cols or not grid_rows:
            total = len(specs)
            grid_cols = int(round(math.sqrt(total))) or 1
            grid_rows = int(math.ceil(total / grid_cols)) or 1
        zero_based_indices = True
        if matrix_indices:
            min_idx = min(matrix_indices)
            max_idx = max(matrix_indices)
            if min_idx >= 1 and max_idx == grid_cols * grid_rows:
                zero_based_indices = False
        maps = {
            'a': np.full((grid_rows, grid_cols), np.nan),
            'b': np.full((grid_rows, grid_cols), np.nan),
            'c': np.full((grid_rows, grid_cols), np.nan),
            'a_err': np.full((grid_rows, grid_cols), np.nan),
            'b_err': np.full((grid_rows, grid_cols), np.nan),
            'c_err': np.full((grid_rows, grid_cols), np.nan),
            'rmse': np.full((grid_rows, grid_cols), np.nan),
        }
        def _axis_from_specs(coord_key, index_key, size):
            if not size:
                return np.arange(0, dtype=float)
            coords = [None] * size
            for spec in specs:
                idx = spec.get(index_key)
                val = spec.get(coord_key)
                if idx is None or val is None:
                    continue
                if idx < 0 or idx >= size:
                    continue
                try:
                    coords[idx] = float(val)
                except Exception:
                    continue
            if any(v is None for v in coords):
                return np.arange(size, dtype=float)
            arr = np.asarray(coords, dtype=float)
            arr = arr - float(np.nanmin(arr))
            return arr

        logs = []
        for idx, spec in enumerate(specs):
            row = spec.get('grid_row')
            col = spec.get('grid_col')
            if row is None or col is None:
                matrix_index = spec.get('matrix_index')
                if matrix_index is not None:
                    idx_val = int(matrix_index)
                    if not zero_based_indices:
                        idx_val -= 1
                    row = idx_val // grid_cols
                    col = idx_val % grid_cols
                else:
                    row = idx // grid_cols
                    col = idx % grid_cols
            try:
                if row < 0 or row >= grid_rows or col < 0 or col >= grid_cols:
                    raise IndexError(f"Index {idx}: ({row}, {col}) outside grid {grid_rows}x{grid_cols}")
                V = np.asarray(spec.get('V', []), dtype=float)
                channels = spec.get('channels') or {}
                if channel_name not in channels:
                    raise ValueError(f"Channel '{channel_name}' missing; available: {', '.join(channels.keys()) or 'none'}")
                channel_data = channels.get(channel_name)
                if channel_data is None:
                    raise ValueError("Channel missing")
                res = fit_parabola_bias(V, channel_data)
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
                maps['a'][row, col] = res['a']
                maps['b'][row, col] = v0 if v0 is not None else np.nan
                maps['c'][row, col] = res['c']
                maps['a_err'][row, col] = res['a_err']
                maps['b_err'][row, col] = v0_err if v0_err is not None else np.nan
                maps['c_err'][row, col] = res['c_err']
                maps['rmse'][row, col] = res['rmse']
            except Exception as exc:
                logs.append(f"Index {idx}: {exc}")
            current = idx + 1
            total = len(specs)
            self.progress.emit(current, total)
            try:
                print(f"[MatrixFit] {current}/{total} processed", flush=True)
            except Exception:
                pass
        payload = {
            'maps': maps,
            'logs': logs,
            'channel_name': channel_name,
            'x_axis': _axis_from_specs('x', 'grid_col', grid_cols),
            'y_axis': _axis_from_specs('y', 'grid_row', grid_rows),
            'axis_unit': axis_unit,
        }
        self.finished.emit(payload)

class MatrixFitDialog(QtWidgets.QDialog):
    PARAM_INFO = {
        'a': {'label': 'a', 'unit': 'a.u.', 'cmap': 'viridis'},
        'b': {'label': 'LCPD', 'unit': 'mV', 'cmap': 'bwr'},
        'c': {'label': 'c', 'unit': 'Hz', 'cmap': 'gray'},
        'a_err': {'label': 'sa', 'unit': 'a.u.', 'cmap': 'magma'},
        'b_err': {'label': 'LCPD err', 'unit': 'mV', 'cmap': 'magma'},
        'c_err': {'label': 'sc', 'unit': 'Hz', 'cmap': 'magma'},
        'rmse': {'label': 'RMSE', 'unit': 'Hz', 'cmap': 'inferno'},
    }

    def __init__(self, viewer, specs, parent=None):
        super().__init__(parent)
        self.viewer = viewer
        self.specs = list(specs)
        self.setWindowTitle("Matrix parabola fits")
        self.resize(900, 700)
        self._worker_thread = None
        self._result_payload = None
        layout = QtWidgets.QVBoxLayout(self)
        self.info_label = QtWidgets.QLabel("Fit KPFM df(V) parabolas for every point in the matrix (other channels are skipped).")
        layout.addWidget(self.info_label)
        ctrl = QtWidgets.QHBoxLayout()
        self.run_btn = QtWidgets.QPushButton("Run fits")
        self.save_btn = QtWidgets.QPushButton("Save maps...")
        self.save_btn.setEnabled(False)
        self.export_xyz_btn = QtWidgets.QPushButton("Export WSxM XYZ...")
        self.export_xyz_btn.setEnabled(False)
        ctrl.addWidget(self.run_btn)
        ctrl.addWidget(self.save_btn)
        ctrl.addWidget(self.export_xyz_btn)
        ctrl.addStretch(1)
        layout.addLayout(ctrl)
        display_box = QtWidgets.QGroupBox("Display options")
        display_layout = QtWidgets.QHBoxLayout(display_box)
        self.scale_mode_combo = QtWidgets.QComboBox()
        self.scale_mode_combo.addItem("Full range", "full")
        self.scale_mode_combo.addItem("Clip percentiles", "clip")
        self.scale_mode_combo.addItem("Centered ?max", "center")
        display_layout.addWidget(QtWidgets.QLabel("Scale:"))
        display_layout.addWidget(self.scale_mode_combo)
        self.low_pct_spin = QtWidgets.QDoubleSpinBox()
        self.low_pct_spin.setRange(0.0, 49.0)
        self.low_pct_spin.setSingleStep(0.5)
        self.low_pct_spin.setValue(2.0)
        self.high_pct_spin = QtWidgets.QDoubleSpinBox()
        self.high_pct_spin.setRange(51.0, 100.0)
        self.high_pct_spin.setSingleStep(0.5)
        self.high_pct_spin.setValue(98.0)
        display_layout.addWidget(QtWidgets.QLabel("Low %"))
        display_layout.addWidget(self.low_pct_spin)
        display_layout.addWidget(QtWidgets.QLabel("High %"))
        display_layout.addWidget(self.high_pct_spin)
        display_layout.addStretch(1)
        layout.addWidget(display_box)
        self.progress = QtWidgets.QProgressBar()
        layout.addWidget(self.progress)
        self.fig = Figure(figsize=(6,5))
        self.canvas = FigureCanvas(self.fig)
        layout.addWidget(self.canvas, 1)
        self.map_value_label = QtWidgets.QLabel("Value: --")
        layout.addWidget(self.map_value_label)
        self.logs = QtWidgets.QTextEdit()
        self.logs.setReadOnly(True)
        self.logs.setFixedHeight(120)
        layout.addWidget(self.logs)
        btn_box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Close)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)
        self.run_btn.clicked.connect(self._start_fit)
        self.save_btn.clicked.connect(self._save_maps)
        self.export_xyz_btn.clicked.connect(self._export_xyz)
        self.scale_mode_combo.currentIndexChanged.connect(self._on_display_option_changed)
        self.low_pct_spin.valueChanged.connect(self._on_display_option_changed)
        self.high_pct_spin.valueChanged.connect(self._on_display_option_changed)
        self._update_percentile_enabled()
        self._axes_to_key = {}
        self.canvas.mpl_connect('motion_notify_event', self._on_map_hover)

    def _start_fit(self):
        if self._worker_thread is not None:
            return
        self.run_btn.setEnabled(False)
        self.save_btn.setEnabled(False)
        self.export_xyz_btn.setEnabled(False)
        self.logs.clear()
        self.progress.setValue(0)
        self._result_payload = None
        worker = MatrixFitWorker(self.specs)
        thread = QtCore.QThread(self)
        self._worker = worker
        worker.moveToThread(thread)
        worker.progress.connect(self._on_progress)
        worker.finished.connect(self._on_finished)
        worker.finished.connect(thread.quit)
        worker.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)
        thread.finished.connect(self._on_thread_finished)
        thread.started.connect(worker.run)
        self._worker_thread = thread
        thread.start()

    def _on_progress(self, current, total):
        self.progress.setMaximum(total)
        self.progress.setValue(current)

    def _on_finished(self, payload):
        self._result_payload = payload
        maps = payload.get('maps', {})
        logs = payload.get('logs', [])
        channel_name = payload.get('channel_name', 'channel')
        for line in logs:
            self.logs.append(line)
        if maps:
            self._render_maps(maps, channel_name)
            self.save_btn.setEnabled(True)
            self.export_xyz_btn.setEnabled(True)
        else:
            self.map_value_label.setText("Value: --")
        self.run_btn.setEnabled(True)
        self._worker = None

    def _on_thread_finished(self):
        self._worker_thread = None

    def _current_display_mode(self):
        return self.scale_mode_combo.currentData()

    def _current_percentiles(self):
        return float(self.low_pct_spin.value()), float(self.high_pct_spin.value())

    def _update_percentile_enabled(self):
        clip = (self._current_display_mode() == 'clip')
        self.low_pct_spin.setEnabled(clip)
        self.high_pct_spin.setEnabled(clip)

    def _on_display_option_changed(self):
        self._update_percentile_enabled()
        if self._result_payload and self._result_payload.get('maps'):
            maps = self._result_payload['maps']
            channel = self._result_payload.get('channel_name', 'channel')
            self._render_maps(maps, channel)
        else:
            self.canvas.draw_idle()

    def _compute_vlims(self, arr):
        mode = self._current_display_mode()
        data = np.asarray(arr, dtype=float)
        if mode == 'clip':
            low, high = self._current_percentiles()
            return robust_limits(data, low_pct=low, high_pct=high)
        finite = data[np.isfinite(data)]
        if finite.size == 0:
            return None, None
        if mode == 'center':
            vmax = float(np.nanmax(np.abs(finite)))
            if not np.isfinite(vmax) or vmax == 0:
                return None, None
            return -vmax, vmax
        return None, None

    def _map_extent(self, arr_shape):
        payload = self._result_payload or {}
        x_axis = payload.get('x_axis')
        y_axis = payload.get('y_axis')
        if x_axis is None or y_axis is None:
            return None
        if len(x_axis) != arr_shape[1] or len(y_axis) != arr_shape[0]:
            return None
        try:
            x0 = float(np.nanmin(x_axis))
            x1 = float(np.nanmax(x_axis))
            y0 = float(np.nanmin(y_axis))
            y1 = float(np.nanmax(y_axis))
        except Exception:
            return None
        if not np.isfinite([x0, x1, y0, y1]).all() or x0 == x1 or y0 == y1:
            return None
        return [x0, x1, y0, y1]

    def _render_maps(self, maps, channel_name):
        self.fig.clf()
        self._axes_to_key = {}
        params = ['a','b','c','a_err','b_err','c_err','rmse']
        cols = 3
        rows = math.ceil(len(params)/cols)
        axis_unit = (self._result_payload or {}).get('axis_unit') or self.PARAM_INFO.get('b', {}).get('unit') or ''
        for idx, key in enumerate(params, 1):
            ax = self.fig.add_subplot(rows, cols, idx)
            info = self.PARAM_INFO.get(key, {'label': key, 'unit': ''})
            ax.set_title(info['label'])
            vmin, vmax = self._compute_vlims(maps[key])
            extent = self._map_extent(maps[key].shape)
            cmap = info.get('cmap', 'viridis')
            im = ax.imshow(maps[key], origin='lower', cmap=cmap, vmin=vmin, vmax=vmax, extent=extent)
            cbar = self.fig.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
            unit = info.get('unit')
            if key in ('b', 'b_err') and axis_unit:
                unit = axis_unit
            if unit:
                cbar.set_label(unit)
            if extent:
                ax.set_xlabel("x (nm)")
                ax.set_ylabel("y (nm)")
            self._axes_to_key[ax] = key
        self.fig.suptitle(f"Parabola fits - channel {channel_name}")
        self.canvas.draw_idle()

    def _save_maps(self):
        if not self._result_payload or not self._result_payload.get('maps'):
            return
        maps = self._result_payload['maps']
        channel_name = self._result_payload.get('channel_name', 'channel')
        x_axis = self._result_payload.get('x_axis')
        y_axis = self._result_payload.get('y_axis')
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save fit maps", "matrix_fit_maps.npz", "NumPy archive (*.npz)")
        if not path:
            return
        metadata = self._collect_fit_metadata(x_axis, y_axis, maps)
        metadata_json = json.dumps(metadata)
        np.savez(path, channel=channel_name, x_axis=x_axis, y_axis=y_axis, metadata=np.array(metadata_json), **maps)
        metadata_path = Path(path).with_suffix('.json')
        try:
            metadata_path.write_text(json.dumps(metadata, indent=2, default=str))
        except Exception:
            pass

    def _export_xyz(self):
        if not self._result_payload or not self._result_payload.get('maps'):
            return
        maps = self._result_payload['maps']
        x_axis = self._result_payload.get('x_axis')
        y_axis = self._result_payload.get('y_axis')
        if x_axis is None or y_axis is None:
            QtWidgets.QMessageBox.warning(self, "Missing coordinates", "Cannot export XYZ without coordinate axes.")
            return
        folder = QtWidgets.QFileDialog.getExistingDirectory(self, "Select folder for WSxM XYZ exports")
        if not folder:
            return
        save_wsxm_xyz(folder, maps['a'], x_axis, y_axis, "a", z_unit="a.u.")
        save_wsxm_xyz(folder, maps['b'], x_axis, y_axis, "b_LCPD", z_unit="mV", z_scale=1000.0)
        save_wsxm_xyz(folder, maps['c'], x_axis, y_axis, "c", z_unit="Hz")
        save_wsxm_xyz(folder, maps['a_err'], x_axis, y_axis, "a_err", z_unit="a.u.")
        save_wsxm_xyz(folder, maps['b_err'], x_axis, y_axis, "b_err", z_unit="mV", z_scale=1000.0)
        save_wsxm_xyz(folder, maps['c_err'], x_axis, y_axis, "c_err", z_unit="Hz")
        save_wsxm_xyz(folder, maps['rmse'], x_axis, y_axis, "rmse", z_unit="Hz")
        self.logs.append(f"WSxM XYZ exports saved to {folder}")

    def get_result_maps(self):
        return self._result_payload

    def _on_map_hover(self, event):
        if self._result_payload is None or not self._result_payload.get('maps'):
            self.map_value_label.setText("Value: --")
            return
        if event.inaxes not in self._axes_to_key:
            self.map_value_label.setText("Value: --")
            return
        key = self._axes_to_key.get(event.inaxes)
        arr = self._result_payload['maps'].get(key)
        if arr is None:
            self.map_value_label.setText("Value: --")
            return
        extent = self._map_extent(arr.shape)
        val = sample_array_value(arr, event.xdata, event.ydata, extent)
        if val is None:
            self.map_value_label.setText("Value: --")
            return
        info = self.PARAM_INFO.get(key, {})
        unit = info.get('unit') or ''
        label = info.get('label', key)
        text = f"{label}: {val:.4g}"
        if unit:
            text += f" {unit}"
        self.map_value_label.setText(text)

    def _collect_fit_metadata(self, x_axis, y_axis, maps):
        specs = self.specs or []
        def _axis_stats(axis):
            if axis is None:
                return (None, None)
            arr = np.asarray(axis, dtype=float)
            if arr.size == 0:
                return (None, None)
            return (float(np.nanmin(arr)), float(np.nanmax(arr)))

        x_min, x_max = _axis_stats(x_axis)
        y_min, y_max = _axis_stats(y_axis)
        meta = {
            'channel': self._result_payload.get('channel_name') if self._result_payload else None,
            'spec_count': len(specs),
            'grid_shape': list(maps['a'].shape) if 'a' in maps else None,
            'x_axis_min': x_min,
            'x_axis_max': x_max,
            'y_axis_min': y_min,
            'y_axis_max': y_max,
        }
        if specs:
            first_path = specs[0].get('path')
            try:
                meta['source_file'] = str(Path(first_path))
            except Exception:
                meta['source_file'] = str(first_path)
        biases = [np.asarray(spec.get('V', []), dtype=float) for spec in specs if spec.get('V') is not None]
        if biases:
            all_bias = np.concatenate([b for b in biases if b.size])
            if all_bias.size:
                meta['bias_min'] = float(np.nanmin(all_bias))
                meta['bias_max'] = float(np.nanmax(all_bias))
            meta['points_per_spectrum'] = int(np.nanmedian([b.size for b in biases if b.size])) if biases else None
        xs = [spec.get('x') for spec in specs if spec.get('x') is not None]
        ys = [spec.get('y') for spec in specs if spec.get('y') is not None]
        if xs:
            meta['position_x_min'] = float(np.nanmin(xs))
            meta['position_x_max'] = float(np.nanmax(xs))
        if ys:
            meta['position_y_min'] = float(np.nanmin(ys))
            meta['position_y_max'] = float(np.nanmax(ys))
        times = [spec.get('time') for spec in specs if isinstance(spec.get('time'), datetime)]
        if times:
            times.sort()
            meta['acquisition_start'] = times[0].isoformat()
            meta['acquisition_end'] = times[-1].isoformat()
            meta['estimated_duration_seconds'] = float((times[-1] - times[0]).total_seconds())
        meta['saved_at'] = datetime.utcnow().isoformat()
        return meta

    def closeEvent(self, event):
        thread = self._worker_thread
        if thread is not None and thread.isRunning():
            thread.quit()
            thread.wait()
        super().closeEvent(event)

    def closeEvent(self, event):
        if self._worker_thread is not None and self._worker_thread.isRunning():
            self._worker_thread.quit()
            self._worker_thread.wait()
        super().closeEvent(event)





