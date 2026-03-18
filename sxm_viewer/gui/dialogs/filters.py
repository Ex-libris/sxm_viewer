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
from ...processing.filters import FILTER_DEFINITIONS
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

class CustomFilterDialog(QtWidgets.QDialog):
    """Dialog to assemble custom filter pipelines."""
    def __init__(self, parent=None, base_image=None, apply_step_func=None):
        super().__init__(parent)
        self.setWindowTitle("Custom filter pipeline")
        self.resize(460, 480)
        self.base_image = base_image
        self.apply_step = apply_step_func
        self._pipeline = []
        layout = QtWidgets.QVBoxLayout(self)
        form = QtWidgets.QFormLayout()
        self.filter_combo = QtWidgets.QComboBox()
        for key, info in FILTER_DEFINITIONS.items():
            self.filter_combo.addItem(info['label'], key)
        form.addRow("Filter", self.filter_combo)
        self.axis_combo = QtWidgets.QComboBox()
        self.axis_combo.addItems(["both","row","col"])
        self.axis_label = QtWidgets.QLabel("Axis")
        form.addRow(self.axis_label, self.axis_combo)
        self.sigma_spin = QtWidgets.QDoubleSpinBox()
        self.sigma_spin.setRange(0.1, 50.0); self.sigma_spin.setSingleStep(0.1); self.sigma_spin.setValue(2.0)
        self.sigma_label = QtWidgets.QLabel("Sigma")
        form.addRow(self.sigma_label, self.sigma_spin)
        self.lap_sigma_spin = QtWidgets.QDoubleSpinBox()
        self.lap_sigma_spin.setRange(0.0, 20.0)
        self.lap_sigma_spin.setSingleStep(0.1)
        self.lap_sigma_spin.setValue(float(FILTER_DEFINITIONS.get("laplacian", {}).get("default_sigma", 0.6)))
        self.lap_sigma_label = QtWidgets.QLabel("Laplace sigma")
        form.addRow(self.lap_sigma_label, self.lap_sigma_spin)
        self.lap_neighbors_combo = QtWidgets.QComboBox()
        self.lap_neighbors_combo.addItem("4-neighbor", 4)
        self.lap_neighbors_combo.addItem("8-neighbor", 8)
        self.lap_neighbors_label = QtWidgets.QLabel("Laplace stencil")
        self.lap_neighbors_combo.setCurrentIndex(1)
        form.addRow(self.lap_neighbors_label, self.lap_neighbors_combo)
        self.lap_abs_cb = QtWidgets.QCheckBox("Absolute response")
        self.lap_abs_cb.setChecked(bool(FILTER_DEFINITIONS.get("laplacian", {}).get("default_absolute", True)))
        self.lap_abs_label = QtWidgets.QLabel("Laplace output")
        form.addRow(self.lap_abs_label, self.lap_abs_cb)
        layout.addLayout(form)
        btn_row = QtWidgets.QHBoxLayout()
        add_btn = QtWidgets.QPushButton("Add step")
        remove_btn = QtWidgets.QPushButton("Remove selected")
        btn_row.addWidget(add_btn); btn_row.addWidget(remove_btn)
        layout.addLayout(btn_row)
        self.pipeline_list = QtWidgets.QListWidget()
        layout.addWidget(self.pipeline_list, 1)
        self.preview_cb = QtWidgets.QCheckBox("Preview on current image")
        layout.addWidget(self.preview_cb)
        self.preview_label = QtWidgets.QLabel("Preview unavailable")
        self.preview_label.setFixedHeight(160)
        self.preview_label.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.preview_label.setAlignment(QtCore.Qt.AlignCenter)
        layout.addWidget(self.preview_label)
        name_row = QtWidgets.QHBoxLayout()
        name_row.addWidget(QtWidgets.QLabel("Name prefix:"))
        self.name_edit = QtWidgets.QLineEdit("Custom")
        name_row.addWidget(self.name_edit)
        layout.addLayout(name_row)
        btn_box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        layout.addWidget(btn_box)
        add_btn.clicked.connect(self._on_add_step)
        remove_btn.clicked.connect(self._on_remove_step)
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)
        self.preview_cb.toggled.connect(self._update_preview)
        self.filter_combo.currentIndexChanged.connect(self._on_filter_selection_changed)
        self._on_filter_selection_changed()

    def _set_param_row_visible(self, label_widget, field_widget, visible):
        label_widget.setVisible(bool(visible))
        field_widget.setVisible(bool(visible))

    def _on_filter_selection_changed(self, _idx=None):
        key = self.filter_combo.currentData()
        self._set_param_row_visible(self.axis_label, self.axis_combo, key == "flatten")
        self._set_param_row_visible(self.sigma_label, self.sigma_spin, key in ("highpass", "lowpass"))
        show_lap = key == "laplacian"
        self._set_param_row_visible(self.lap_sigma_label, self.lap_sigma_spin, show_lap)
        self._set_param_row_visible(self.lap_neighbors_label, self.lap_neighbors_combo, show_lap)
        self._set_param_row_visible(self.lap_abs_label, self.lap_abs_cb, show_lap)

    def _current_step(self):
        key = self.filter_combo.currentData()
        params = {}
        if key == 'flatten':
            params['axis'] = self.axis_combo.currentText()
        if key in ('highpass','lowpass'):
            params['sigma'] = float(self.sigma_spin.value())
        if key == 'laplacian':
            params['sigma'] = float(self.lap_sigma_spin.value())
            params['neighbors'] = int(self.lap_neighbors_combo.currentData() or 8)
            params['absolute'] = bool(self.lap_abs_cb.isChecked())
        return {'key': key, 'params': params}

    def _on_add_step(self):
        step = self._current_step()
        label = FILTER_DEFINITIONS.get(step['key'], {}).get('label', step['key'])
        self._pipeline.append(step)
        self.pipeline_list.addItem(f"{len(self._pipeline)}. {label}")
        self._update_preview()

    def _on_remove_step(self):
        row = self.pipeline_list.currentRow()
        if row >= 0:
            self.pipeline_list.takeItem(row)
            del self._pipeline[row]
            self.pipeline_list.clear()
            for idx, step in enumerate(self._pipeline, 1):
                label = FILTER_DEFINITIONS.get(step['key'], {}).get('label', step['key'])
                self.pipeline_list.addItem(f"{idx}. {label}")
            self._update_preview()

    def _update_preview(self):
        if not self.preview_cb.isChecked() or self.base_image is None or not self.apply_step:
            self.preview_label.setText("Preview unavailable")
            self.preview_label.setPixmap(QtGui.QPixmap())
            return
        arr = np.asarray(self.base_image, dtype=float)
        for step in self._pipeline:
            arr = self.apply_step(arr, step)
        qimg = array_to_qimage(arr)
        pix = QtGui.QPixmap.fromImage(qimg).scaled(self.preview_label.width(), self.preview_label.height(),
                                                   QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation)
        self.preview_label.setPixmap(pix)
        self.preview_label.setText("")

    def pipeline_steps(self):
        return list(self._pipeline)

    def pipeline_label(self):
        return self.name_edit.text().strip() or "Custom"



# === BEGIN: Image adjustment classes (drop-in replacement) ===
# These classes are intended to replace the existing ImageAdjustPreviewPanel and ImageAdjustDialog
# in detail_panels.py without adding new third-party dependencies.





