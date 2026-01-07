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

from ..._shared import *
from ...config import *
from ...data.io import *
from ...data.spectroscopy import *
from ..thumbnails import *

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
        form.addRow("Axis", self.axis_combo)
        self.sigma_spin = QtWidgets.QDoubleSpinBox()
        self.sigma_spin.setRange(0.1, 50.0); self.sigma_spin.setSingleStep(0.1); self.sigma_spin.setValue(2.0)
        form.addRow("Sigma", self.sigma_spin)
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

    def _current_step(self):
        key = self.filter_combo.currentData()
        params = {}
        if key == 'flatten':
            params['axis'] = self.axis_combo.currentText()
        if key in ('highpass','lowpass'):
            params['sigma'] = float(self.sigma_spin.value())
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
# in detail_panels.py without adding new third‑party dependencies.
