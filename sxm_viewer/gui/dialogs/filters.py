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
from ..ppt_mixin import PPTContextMenuMixin
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

class SingleFilterDialog(QtWidgets.QDialog):
    """Single-step filter dialog with live preview support."""

    def __init__(
        self,
        parent=None,
        filter_key=None,
        base_image=None,
        apply_step_func=None,
        preview_callback=None,
        initial_params=None,
        preview_target_text="current image",
    ):
        super().__init__(parent)
        self.filter_key = str(filter_key or "").strip().lower()
        self.base_image = base_image
        self.apply_step = apply_step_func
        self._preview_callback = preview_callback
        self._initial_params = dict(initial_params or {})
        self._preview_target_text = str(preview_target_text or "current image").strip()
        self._preview_timer = QtCore.QTimer(self)
        self._preview_timer.setSingleShot(True)
        self._preview_timer.setInterval(40)
        self._preview_timer.timeout.connect(self._update_preview)

        filter_label = FILTER_DEFINITIONS.get(self.filter_key, {}).get("label", self.filter_key or "Filter")
        self.setWindowTitle(f"{filter_label} preview")
        self.resize(460, 420)

        layout = QtWidgets.QVBoxLayout(self)
        intro = QtWidgets.QLabel(
            f"Preview updates live on the {self._preview_target_text}. "
            "Press OK to apply or Cancel to restore the original image."
        )
        intro.setWordWrap(True)
        layout.addWidget(intro)

        form = QtWidgets.QFormLayout()
        self.axis_combo = QtWidgets.QComboBox()
        self.axis_combo.addItems(["both", "row", "col"])
        self.axis_label = QtWidgets.QLabel("Axis")
        form.addRow(self.axis_label, self.axis_combo)

        self.sigma_spin = QtWidgets.QDoubleSpinBox()
        self.sigma_spin.setDecimals(2)
        self.sigma_spin.setRange(0.05, 50.0)
        self.sigma_spin.setSingleStep(0.1)
        self.sigma_spin.setValue(float(self._initial_params.get("sigma", FILTER_DEFINITIONS.get(self.filter_key, {}).get("default_sigma", 2.0))))
        self.sigma_label = QtWidgets.QLabel("Sigma (px)")
        form.addRow(self.sigma_label, self.sigma_spin)

        self.lap_sigma_spin = QtWidgets.QDoubleSpinBox()
        self.lap_sigma_spin.setDecimals(2)
        self.lap_sigma_spin.setRange(0.0, 20.0)
        self.lap_sigma_spin.setSingleStep(0.1)
        self.lap_sigma_spin.setValue(float(self._initial_params.get("sigma", FILTER_DEFINITIONS.get("laplacian", {}).get("default_sigma", 0.6))))
        self.lap_sigma_label = QtWidgets.QLabel("Laplace sigma")
        form.addRow(self.lap_sigma_label, self.lap_sigma_spin)

        self.lap_neighbors_combo = QtWidgets.QComboBox()
        self.lap_neighbors_combo.addItem("4-neighbor", 4)
        self.lap_neighbors_combo.addItem("8-neighbor", 8)
        neigh_default = int(self._initial_params.get("neighbors", FILTER_DEFINITIONS.get("laplacian", {}).get("default_neighbors", 8)))
        self.lap_neighbors_combo.setCurrentIndex(1 if neigh_default == 8 else 0)
        self.lap_neighbors_label = QtWidgets.QLabel("Laplace stencil")
        form.addRow(self.lap_neighbors_label, self.lap_neighbors_combo)

        self.lap_abs_cb = QtWidgets.QCheckBox("Absolute response")
        self.lap_abs_cb.setChecked(bool(self._initial_params.get("absolute", FILTER_DEFINITIONS.get("laplacian", {}).get("default_absolute", True))))
        self.lap_abs_label = QtWidgets.QLabel("Laplace output")
        form.addRow(self.lap_abs_label, self.lap_abs_cb)
        layout.addLayout(form)

        self.preview_label = QtWidgets.QLabel("Preview unavailable")
        self.preview_label.setFixedHeight(180)
        self.preview_label.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.preview_label.setAlignment(QtCore.Qt.AlignCenter)
        PPTContextMenuMixin.install(self.preview_label, label_text="Filter preview")
        layout.addWidget(self.preview_label)

        btn_box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        layout.addWidget(btn_box)
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)

        self.axis_combo.currentTextChanged.connect(self._schedule_preview_update)
        self.sigma_spin.valueChanged.connect(self._schedule_preview_update)
        self.lap_sigma_spin.valueChanged.connect(self._schedule_preview_update)
        self.lap_neighbors_combo.currentIndexChanged.connect(self._schedule_preview_update)
        self.lap_abs_cb.toggled.connect(self._schedule_preview_update)

        self._on_filter_selection_changed()
        self._schedule_preview_update()

    def _set_param_row_visible(self, label_widget, field_widget, visible):
        label_widget.setVisible(bool(visible))
        field_widget.setVisible(bool(visible))

    def _on_filter_selection_changed(self):
        key = self.filter_key
        self._set_param_row_visible(self.axis_label, self.axis_combo, key == "flatten")
        self._set_param_row_visible(self.sigma_label, self.sigma_spin, key in ("highpass", "lowpass"))
        show_lap = key == "laplacian"
        self._set_param_row_visible(self.lap_sigma_label, self.lap_sigma_spin, show_lap)
        self._set_param_row_visible(self.lap_neighbors_label, self.lap_neighbors_combo, show_lap)
        self._set_param_row_visible(self.lap_abs_label, self.lap_abs_cb, show_lap)

    def _schedule_preview_update(self, *_args):
        self._preview_timer.start()

    def current_step(self):
        params = {}
        if self.filter_key == "flatten":
            params["axis"] = self.axis_combo.currentText()
        elif self.filter_key in ("highpass", "lowpass"):
            params["sigma"] = float(self.sigma_spin.value())
        elif self.filter_key == "laplacian":
            params["sigma"] = float(self.lap_sigma_spin.value())
            params["neighbors"] = int(self.lap_neighbors_combo.currentData() or 8)
            params["absolute"] = bool(self.lap_abs_cb.isChecked())
        return {"key": self.filter_key, "params": params}

    def current_step_label(self):
        step = self.current_step()
        label = FILTER_DEFINITIONS.get(step["key"], {}).get("label", step["key"])
        params = step.get("params") or {}
        if step["key"] in ("highpass", "lowpass") and params.get("sigma") is not None:
            return f"{label} (sigma={float(params['sigma']):.2f} px)"
        if step["key"] == "laplacian":
            return (
                f"{label} (sigma={float(params.get('sigma', 0.0)):.2f} px, "
                f"{int(params.get('neighbors', 8))}-nbr)"
            )
        if step["key"] == "flatten" and params.get("axis"):
            return f"{label} ({params['axis']})"
        return label

    def _update_preview(self):
        step = self.current_step()
        if self.base_image is None or not self.apply_step:
            self.preview_label.setText("Preview unavailable")
            self.preview_label.setPixmap(QtGui.QPixmap())
        else:
            arr = np.asarray(self.base_image, dtype=float)
            try:
                arr = self.apply_step(arr, step)
            except Exception:
                pass
            qimg = array_to_qimage(arr)
            pix = QtGui.QPixmap.fromImage(qimg).scaled(
                self.preview_label.width(),
                self.preview_label.height(),
                QtCore.Qt.KeepAspectRatio,
                QtCore.Qt.SmoothTransformation,
            )
            self.preview_label.setPixmap(pix)
            self.preview_label.setText("")
        if callable(self._preview_callback):
            try:
                self._preview_callback(step, self.current_step_label())
            except Exception:
                pass


class CustomFilterDialog(QtWidgets.QDialog):
    """Dialog to assemble custom filter pipelines."""
    def __init__(self, parent=None, base_image=None, apply_step_func=None, preview_callback=None, preview_target_text="current image"):
        super().__init__(parent)
        self.setWindowTitle("Custom filter pipeline")
        self.resize(460, 480)
        self.base_image = base_image
        self.apply_step = apply_step_func
        self._preview_callback = preview_callback
        self._preview_target_text = str(preview_target_text or "current image").strip()
        self._pipeline = []
        self._preview_timer = QtCore.QTimer(self)
        self._preview_timer.setSingleShot(True)
        self._preview_timer.setInterval(40)
        self._preview_timer.timeout.connect(self._update_preview)
        layout = QtWidgets.QVBoxLayout(self)
        intro = QtWidgets.QLabel(
            f"Preview updates live on the {self._preview_target_text}. "
            "Press OK to apply or Cancel to restore the original image."
        )
        intro.setWordWrap(True)
        layout.addWidget(intro)
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
        self.preview_cb.setChecked(True)
        layout.addWidget(self.preview_cb)
        self.preview_label = QtWidgets.QLabel("Preview unavailable")
        self.preview_label.setFixedHeight(160)
        self.preview_label.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.preview_label.setAlignment(QtCore.Qt.AlignCenter)
        PPTContextMenuMixin.install(self.preview_label, label_text="Filter preview")
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
        self.axis_combo.currentTextChanged.connect(self._schedule_preview_update)
        self.sigma_spin.valueChanged.connect(self._schedule_preview_update)
        self.lap_sigma_spin.valueChanged.connect(self._schedule_preview_update)
        self.lap_neighbors_combo.currentIndexChanged.connect(self._schedule_preview_update)
        self.lap_abs_cb.toggled.connect(self._schedule_preview_update)
        self._on_filter_selection_changed()
        self._schedule_preview_update()

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
        self._schedule_preview_update()

    def _schedule_preview_update(self, *_args):
        self._preview_timer.start()

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
        self._schedule_preview_update()

    def _on_remove_step(self):
        row = self.pipeline_list.currentRow()
        if row >= 0:
            self.pipeline_list.takeItem(row)
            del self._pipeline[row]
            self.pipeline_list.clear()
            for idx, step in enumerate(self._pipeline, 1):
                label = FILTER_DEFINITIONS.get(step['key'], {}).get('label', step['key'])
                self.pipeline_list.addItem(f"{idx}. {label}")
            self._schedule_preview_update()

    def _preview_steps(self):
        if self._pipeline:
            return list(self._pipeline)
        current = self._current_step()
        return [current] if current else []

    def _update_preview(self):
        steps = self._preview_steps()
        if not self.preview_cb.isChecked() or not steps or self.base_image is None or not self.apply_step:
            self.preview_label.setText("Preview unavailable")
            self.preview_label.setPixmap(QtGui.QPixmap())
        else:
            arr = np.asarray(self.base_image, dtype=float)
            for step in steps:
                arr = self.apply_step(arr, step)
            qimg = array_to_qimage(arr)
            pix = QtGui.QPixmap.fromImage(qimg).scaled(self.preview_label.width(), self.preview_label.height(),
                                                       QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation)
            self.preview_label.setPixmap(pix)
            self.preview_label.setText("")
        if callable(self._preview_callback):
            try:
                self._preview_callback(steps if self.preview_cb.isChecked() else None, self.pipeline_label())
            except Exception:
                pass

    def pipeline_steps(self):
        if self._pipeline:
            return list(self._pipeline)
        current = self._current_step()
        return [current] if current else []

    def pipeline_label(self):
        custom = self.name_edit.text().strip()
        if custom and custom != "Custom":
            return custom
        if self._pipeline:
            return custom or "Custom"
        current = self._current_step()
        return FILTER_DEFINITIONS.get((current or {}).get("key"), {}).get("label", custom or "Custom")



# === BEGIN: Image adjustment classes (drop-in replacement) ===
# These classes are intended to replace the existing ImageAdjustPreviewPanel and ImageAdjustDialog
# in detail_panels.py without adding new third-party dependencies.





