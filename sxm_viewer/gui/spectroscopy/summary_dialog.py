"""Spectroscopy summary dialog."""
from __future__ import annotations

from pathlib import Path

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QDialog, QVBoxLayout, QCheckBox, QPushButton, QLabel, QListWidget, QListWidgetItem

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
from ..ppt_mixin import PPTContextMenuMixin
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

class SpectroSummaryDialog(QtWidgets.QDialog):
    """Compact modal that lists spectra for a given file and offers quick actions."""
    def __init__(self, parent, file_key, header, fds, entries, nearest=None, show_mode="single"):
        super().__init__(parent)
        self.viewer = parent
        self._file_key = str(file_key)
        self._header = header or {}
        self._fds = fds or []
        self._entries = list(entries)
        self._show_mode = show_mode  # "single" or "matrix"
        self._single_entries = [s for s in self._entries if s.get('matrix_index') is None]
        self._matrix_entries = [s for s in self._entries if s.get('matrix_index') is not None]
        self._active_entries = self._single_entries if self._show_mode != "matrix" else self._matrix_entries
        # ensure marker colors exist on viewer
        if not hasattr(self.viewer, 'spectro_marker_color_single'):
            self.viewer.spectro_marker_color_single = QtGui.QColor(255, 160, 0, 200)
        if not hasattr(self.viewer, 'spectro_marker_color_matrix'):
            self.viewer.spectro_marker_color_matrix = QtGui.QColor(64, 200, 255, 200)
        self._dialog_cmap = getattr(self.viewer, 'preview_cmap', 'viridis')
        self._spec_to_item = {}
        self.setWindowTitle(f"Spectros: {Path(file_key).name}")
        self.setMinimumWidth(520)
        layout = QVBoxLayout(self)
        n_total = len(self._entries)
        n_matrix = len(self._matrix_entries)
        n_single = len(self._single_entries)
        label = QLabel(f"<b>{n_total}</b> spectroscopies  Single: {n_single}  Matrix: {n_matrix}")
        layout.addWidget(label)

        # Preview + channel selector + show-points toggle
        top_row = QtWidgets.QHBoxLayout()
        self.preview_lbl = QLabel("Preview")
        self.preview_lbl.setAlignment(QtCore.Qt.AlignCenter)
        self.preview_lbl.setMinimumSize(220, 200)
        self.preview_lbl.setStyleSheet("QLabel { border: 1px solid #555; background: #111; }")
        PPTContextMenuMixin.install(self.preview_lbl, label_text=f"{Path(self._file_key).name} preview")
        top_row.addWidget(self.preview_lbl, 1)
        side_v = QtWidgets.QVBoxLayout()
        self.channel_combo = QtWidgets.QComboBox()
        for idx, fd in enumerate(self._fds):
            cap = fd.get('Caption', fd.get('FileName', f"chan{idx}"))
            self.channel_combo.addItem(f"{idx}: {cap}", userData=idx)
        self.channel_combo.currentIndexChanged.connect(self._render_preview)
        side_v.addWidget(self.channel_combo)
        self.show_points_cb = QCheckBox("Show points on preview")
        self.show_points_cb.setChecked(True)
        self.show_points_cb.toggled.connect(self._render_preview)
        side_v.addWidget(self.show_points_cb)
        color_btn = QPushButton("Marker color")
        color_btn.clicked.connect(self._pick_marker_color)
        side_v.addWidget(color_btn)
        # Matrix file filter (only in matrix mode)
        self.matrix_filter_combo = None
        self.matrix_cmap_combo = None
        if self._show_mode == "matrix":
            self._matrix_groups = {}
            for s in self._matrix_entries:
                path_key = str(s.get('path') or "")
                self._matrix_groups.setdefault(path_key, []).append(s)
            if self._matrix_groups:
                self.matrix_filter_combo = QtWidgets.QComboBox()
                self.matrix_filter_combo.addItem("All matrix files", userData=None)
                for path_key, specs in sorted(self._matrix_groups.items()):
                    self.matrix_filter_combo.addItem(Path(path_key).name, userData=path_key)
                self.matrix_filter_combo.currentIndexChanged.connect(self._on_matrix_filter_changed)
                side_v.addWidget(self.matrix_filter_combo)
            # Colormap selector for matrix preview
            try:
                cmap_list = sorted(colormaps.keys())
            except Exception:
                cmap_list = ['viridis','plasma','inferno','magma','cividis','gray','hot','coolwarm','turbo']
            self.matrix_cmap_combo = QtWidgets.QComboBox()
            for name in cmap_list:
                self.matrix_cmap_combo.addItem(name)
            self.matrix_cmap_combo.setCurrentText(self._dialog_cmap)
            self.matrix_cmap_combo.currentTextChanged.connect(self._on_matrix_cmap_changed)
            side_v.addWidget(self.matrix_cmap_combo)
        side_v.addStretch(1)
        top_row.addLayout(side_v)
        layout.addLayout(top_row)

        self.list_w = QListWidget()
        self._rebuild_list()
        layout.addWidget(self.list_w, 1)
        btn_row = QtWidgets.QHBoxLayout()
        open_btn = QPushButton("Open spectro browser")
        open_btn.clicked.connect(self._on_open_browser)
        btn_row.addWidget(open_btn)
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        btn_row.addWidget(close_btn)
        layout.addLayout(btn_row)
        self.list_w.itemDoubleClicked.connect(lambda it: self._open_single(it.data(QtCore.Qt.UserRole)))
        self.list_w.currentItemChanged.connect(self._on_list_selection_changed)
        self.list_w.itemClicked.connect(self._on_item_clicked)
        self.preview_lbl.mousePressEvent = self._on_preview_click

        self._render_preview()

    def _render_preview(self):
        try:
            if not self._fds:
                self.preview_lbl.setText("No channels")
                return
            idx = self.channel_combo.currentData() if self.channel_combo.count() else 0
            fd = self._fds[int(idx)] if idx is not None and 0 <= int(idx) < len(self._fds) else self._fds[0]
            channel_label = fd.get('Caption', fd.get('FileName', f"chan{idx}"))
            self.preview_lbl.ppt_image_label = f"{Path(self._file_key).name} - {channel_label}"
            unit_final, arr = self.viewer._get_filtered_channel_array(self._file_key, self._fds.index(fd), self._header, fd)
            unit_disp, arr_disp, _ = self.viewer._scale_unit_for_display(unit_final, arr)
            arr_disp = self.viewer._downsample_for_thumbnail(arr_disp, 240, 200)
            qimg = array_to_qimage(arr_disp, cmap_name=self._dialog_cmap)
            pix = QtGui.QPixmap.fromImage(qimg.scaled(240, 200, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation))
            self._preview_markers = []
            if self.show_points_cb.isChecked():
                try:
                    xpix = int(self._header.get('xPixel', arr.shape[1] if arr.ndim == 2 else 0))
                    ypix = int(self._header.get('yPixel', arr.shape[0] if arr.ndim == 2 else 0))
                    # force point rendering, no density, smaller markers for dialog preview
                    selected = None
                    cur = self.list_w.currentItem()
                    if cur:
                        selected = cur.data(QtCore.Qt.UserRole)
                    # render only the subset requested (single or matrix)
                    entries = self._active_entries
                    saved_single = self.viewer.show_single_markers
                    saved_matrix = self.viewer.show_matrix_markers
                    try:
                        # force singles on so matrix entries drawn as points too
                        self.viewer.show_single_markers = True
                        self.viewer.show_matrix_markers = (self._show_mode == "matrix")
                        self._preview_markers = self.viewer._render_spectroscopy_overlays(
                            pix, self._header, self._file_key, xpix, ypix,
                            reveal_points_override=True, selected_spec=selected, entries_override=entries,
                            matrix_as_points=(self._show_mode == "matrix"))
                    finally:
                        self.viewer.show_single_markers = saved_single
                        self.viewer.show_matrix_markers = saved_matrix
                except Exception:
                    pass
            self.preview_lbl.setPixmap(pix)
        except Exception:
            self.preview_lbl.setText("Preview unavailable")

    def _on_matrix_filter_changed(self, _idx):
        if self.matrix_filter_combo is None:
            return
        path_key = self.matrix_filter_combo.currentData()
        if path_key:
            self._active_entries = self._matrix_groups.get(path_key, [])
        else:
            self._active_entries = self._matrix_entries
        self._rebuild_list()
        self._render_preview()

    def _on_matrix_cmap_changed(self, name):
        self._dialog_cmap = name or self._dialog_cmap
        self._render_preview()

    def _pick_marker_color(self):
        try:
            current = self.viewer.spectro_marker_color_matrix if self._show_mode == "matrix" else self.viewer.spectro_marker_color_single
            color = QtWidgets.QColorDialog.getColor(current, self, "Select marker color")
            if color.isValid():
                if self._show_mode == "matrix":
                    self.viewer.spectro_marker_color_matrix = color
                else:
                    self.viewer.spectro_marker_color_single = color
                # refresh preview and thumbnails
                self._render_preview()
                try:
                    self.viewer.populate_thumbnails_for_channel(self.viewer.channel_dropdown.currentIndex())
                except Exception:
                    pass
        except Exception:
            pass

    def _on_open_browser(self):
        try:
            self.viewer.on_open_spectro_browser(self._entries)
        except Exception:
            pass
        self.accept()

    def _open_single(self, spectro):
        try:
            self.viewer._open_single_spectro_popup(spectro)
        except Exception:
            pass
        # keep dialog open for quick browsing

    def _on_list_selection_changed(self, current, _prev):
        try:
            self._render_preview()
        except Exception:
            pass

    def _on_item_clicked(self, item):
        try:
            spec = item.data(QtCore.Qt.UserRole)
            mods = QtWidgets.QApplication.keyboardModifiers()
            if mods & QtCore.Qt.ShiftModifier:
                self.viewer._toggle_multi_spec_selection(spec)
                return
            if spec:
                self.viewer._open_single_spectro_popup(spec)
        except Exception:
            pass

    def _on_preview_click(self, event):
        if not hasattr(self, '_preview_markers'):
            return
        pos = event.pos()
        pix = self.preview_lbl.pixmap()
        if pix is None:
            return
        offset_x = (self.preview_lbl.width() - pix.width()) / 2.0
        offset_y = (self.preview_lbl.height() - pix.height()) / 2.0
        px = pos.x() - offset_x
        py = pos.y() - offset_y
        if px < 0 or py < 0 or px > pix.width() or py > pix.height():
            return
        for info in self._preview_markers:
            rect = info.get('rect')
            spec = info.get('spec')
            if rect and spec and rect.contains(px, py):
                key = self.viewer._spec_identity_key(spec) or str(spec.get('path'))
                item = self._spec_to_item.get(key)
                if item:
                    self.list_w.setCurrentItem(item)
                try:
                    mods = event.modifiers()
                    if mods & QtCore.Qt.ShiftModifier:
                        self.viewer._toggle_multi_spec_selection(spec)
                    else:
                        self.viewer._open_single_spectro_popup(spec)
                except Exception:
                    pass
                break

    def _rebuild_list(self):
        self.list_w.clear()
        self._spec_to_item.clear()
        list_source = self._active_entries
        for idx, s in enumerate(list_source, 1):
            sx = s.get('x'); sy = s.get('y'); mid = s.get('matrix_index')
            if sx is not None and sy is not None:
                txt = f"{idx}. {sx:.1f}/{sy:.1f} nm"
            else:
                txt = f"{idx}. <no-pos>"
            if mid is not None:
                txt += f"  [matrix {mid}]"
            it = QListWidgetItem(txt)
            it.setData(QtCore.Qt.UserRole, s)
            self.list_w.addItem(it)
            self._spec_to_item[self.viewer._spec_identity_key(s) or str(idx)] = it

# ---------- Viewer stubs / hooks for integration ----------
__all__ = [
    "SpectroSummaryDialog",
]





