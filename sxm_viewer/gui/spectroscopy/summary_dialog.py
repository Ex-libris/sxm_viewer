"""Spectroscopy summary dialog."""
from __future__ import annotations

from pathlib import Path

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QDialog, QVBoxLayout, QCheckBox, QPushButton, QLabel, QTreeWidget, QTreeWidgetItem

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
from ... import cmap_registry
from ..ppt_mixin import PPTContextMenuMixin
from ..system_open import add_source_file_menu
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
        self._site_to_item = {}
        self.setWindowTitle(f"Spectroscopy at {Path(file_key).name} ({len(self._entries)} spectra)")
        self.setMinimumWidth(520)
        layout = QVBoxLayout(self)
        n_total = len(self._entries)
        n_matrix = len(self._matrix_entries)
        n_single = len(self._single_entries)
        site_count = len({str(s.get("site_key") or self.viewer._spec_identity_key(s) or id(s)) for s in self._active_entries})
        self.summary_label = QLabel(
            f"<b>{n_total}</b> spectra  Positions: {site_count}  Single: {n_single}  Grid map: {n_matrix}"
        )
        layout.addWidget(self.summary_label)
        self.info_label = QLabel("Single-click to inspect. Double-click a position to compare its spectra, or a single spectrum to open it.")
        self.info_label.setWordWrap(True)
        self.info_label.setStyleSheet("QLabel { color: #666; }")
        layout.addWidget(self.info_label)

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
            _featured = cmap_registry.featured_cmap_names("general")
            cmap_list = _featured + [n for n in cmap_registry.all_cmap_names()
                                     if n not in set(_featured)]
            self.matrix_cmap_combo = QtWidgets.QComboBox()
            for name in cmap_list:
                self.matrix_cmap_combo.addItem(name)
            self.matrix_cmap_combo.setCurrentText(self._dialog_cmap)
            self.matrix_cmap_combo.currentTextChanged.connect(self._on_matrix_cmap_changed)
            side_v.addWidget(self.matrix_cmap_combo)
        side_v.addStretch(1)
        top_row.addLayout(side_v)
        layout.addLayout(top_row)

        self.list_w = QTreeWidget()
        self.list_w.setHeaderHidden(True)
        self.list_w.setRootIsDecorated(True)
        self.list_w.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
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
        self.list_w.itemDoubleClicked.connect(self._on_item_double_clicked)
        self.list_w.currentItemChanged.connect(self._on_list_selection_changed)
        self.list_w.customContextMenuRequested.connect(self._on_list_context_menu)
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
                    visible_entries = list(self._active_entries)
                    if cur:
                        payload = cur.data(0, QtCore.Qt.UserRole)
                        if isinstance(payload, dict):
                            kind = str(payload.get("kind") or "")
                            if kind in {"site", "spec"}:
                                selected = payload.get("spec")
                            if kind == "site":
                                visible_entries = list(payload.get("specs") or visible_entries)
                    # render either the whole file or the selected site subset
                    entries = visible_entries
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

    def _open_site(self, spec):
        try:
            controller = getattr(self.viewer, "spectro_compare_controller", None)
            if controller and spec:
                controller.open_site_popup(spec, file_key=self._file_key)
        except Exception:
            pass

    def _on_list_selection_changed(self, current, _prev):
        try:
            payload = current.data(0, QtCore.Qt.UserRole) if current is not None else None
            if isinstance(payload, dict):
                kind = str(payload.get("kind") or "")
                if kind == "site":
                    spec = payload.get("spec")
                    site_summary = str((spec or {}).get("site_summary") or (spec or {}).get("site_display") or "Position").strip()
                    count = len(list(payload.get("specs") or []))
                    self.info_label.setText(f"{site_summary}\n{count} spectra at this position.")
                elif kind == "spec":
                    spec = payload.get("spec")
                    if spec:
                        txt = Path(spec.get("path", "")).name
                        summary = str(spec.get("site_summary") or "").strip()
                        if summary:
                            txt = f"{txt}\n{summary}"
                        self.info_label.setText(txt)
                elif kind == "image":
                    self.info_label.setText(
                        f"{Path(self._file_key).name}\n{len(self._active_entries)} spectra in this image summary."
                    )
            self._highlight_from_item(current)
            self._render_preview()
        except Exception:
            pass

    def _on_item_double_clicked(self, item, _column=0):
        try:
            payload = item.data(0, QtCore.Qt.UserRole)
            if not isinstance(payload, dict):
                return
            kind = str(payload.get("kind") or "")
            if kind == "site":
                self._open_site(payload.get("spec"))
                return
            if kind == "spec":
                spec = payload.get("spec")
                mods = QtWidgets.QApplication.keyboardModifiers()
                if mods & QtCore.Qt.ShiftModifier:
                    self.viewer._toggle_multi_spec_selection(spec)
                    return
                if spec:
                    self.viewer._open_single_spectro_popup(spec)
                return
            if kind == "image":
                self.viewer.on_open_spectro_browser(self._entries)
                return
        except Exception:
            pass

    def _current_selected_spec(self):
        current = self.list_w.currentItem()
        if current is None:
            return None
        payload = current.data(0, QtCore.Qt.UserRole)
        if not isinstance(payload, dict):
            return None
        kind = str(payload.get("kind") or "")
        if kind in {"site", "spec"}:
            return payload.get("spec")
        return None

    def _current_visible_entries(self):
        current = self.list_w.currentItem()
        if current is None:
            return list(self._active_entries)
        payload = current.data(0, QtCore.Qt.UserRole)
        if not isinstance(payload, dict):
            return list(self._active_entries)
        kind = str(payload.get("kind") or "")
        if kind == "site":
            return list(payload.get("specs") or [])
        if kind == "spec":
            spec = payload.get("spec")
            return [spec] if spec is not None else list(self._active_entries)
        return list(self._active_entries)

    def _payload_specs(self, payload):
        if not isinstance(payload, dict):
            return []
        kind = str(payload.get("kind") or "")
        if kind == "spec":
            spec = payload.get("spec")
            return [spec] if spec is not None else []
        if kind == "site":
            specs = list(payload.get("specs") or [])
            if specs:
                return specs
            spec = payload.get("spec")
            return [spec] if spec is not None else []
        return []

    def _site_label(self, specs):
        first = specs[0]
        display = str(first.get("site_display") or "Position").strip()
        trace_count = len(specs)
        channel_count = int(first.get("site_channel_count") or 0)
        low_conf_count = sum(
            1 for spec in list(specs or [])
            if str(spec.get("assignment_confidence") or "").strip().lower() == "low"
        )
        extras = [f"{trace_count} spectrum" if trace_count == 1 else f"{trace_count} spectra"]
        if channel_count:
            extras.append(f"{channel_count} ch")
        if low_conf_count:
            extras.append(f"low conf {low_conf_count}")
        if first.get("site_has_z_stack"):
            extras.append("Z series")
        elif int(first.get("xy_stack_count") or 0) > 1:
            extras.append("repeats")
        if first.get("site_has_matrix"):
            extras.append("grid map")
        return f"{display}  [{' | '.join(extras)}]"

    def _spec_label(self, spec, index):
        sx = spec.get('x')
        sy = spec.get('y')
        parts = [f"{index}."]
        channel = str(spec.get("channel_name") or "").strip()
        if channel:
            parts.append(channel)
        if sx is not None and sy is not None:
            try:
                parts.append(f"{float(sx):.1f}/{float(sy):.1f} nm")
            except Exception:
                parts.append(f"{sx}/{sy}")
        else:
            parts.append("<no-pos>")
        if spec.get('matrix_index') is not None:
            parts.append(f"[matrix {spec.get('matrix_index')}]")
        parts.append(Path(spec.get("path", "")).name)
        return "  ".join(str(p) for p in parts if str(p))

    def _site_groups(self):
        groups = OrderedDict()
        for spec in list(self._active_entries):
            site_key = str(spec.get("site_key") or self.viewer._spec_identity_key(spec) or id(spec))
            groups.setdefault(site_key, []).append(spec)
        return groups

    def _highlight_from_item(self, item):
        if item is None:
            try:
                self.viewer._highlight_spectrum_entry(None)
            except Exception:
                pass
            return
        payload = item.data(0, QtCore.Qt.UserRole)
        if not isinstance(payload, dict):
            return
        kind = str(payload.get("kind") or "")
        if kind in {"site", "spec"}:
            spec = payload.get("spec")
            if spec is not None:
                try:
                    self.viewer._highlight_spectrum_entry(spec)
                except Exception:
                    pass
                return
        try:
            self.viewer._highlight_spectrum_entry(None)
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
                try:
                    label = str(info.get("label") or "")
                    if label == "stack-badge":
                        site_key = str(spec.get("site_key") or "")
                        item = self._site_to_item.get(site_key)
                        if item:
                            self.list_w.setCurrentItem(item)
                        self._open_site(spec)
                        break
                    key = self.viewer._spec_identity_key(spec) or str(spec.get('path'))
                    item = self._spec_to_item.get(key)
                    if item:
                        self.list_w.setCurrentItem(item)
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
        self._site_to_item.clear()
        groups = self._site_groups()
        site_count = len(groups)
        self.summary_label.setText(
            f"<b>{len(self._entries)}</b> spectroscopies  Sites: {site_count}  Single: {len(self._single_entries)}  Matrix: {len(self._matrix_entries)}"
        )
        image_item = QTreeWidgetItem([f"{Path(self._file_key).name}  [{site_count} site" + ("" if site_count == 1 else "s") + f" | {len(self._active_entries)} spectra]"])
        image_item.setData(0, QtCore.Qt.UserRole, {"kind": "image", "specs": list(self._active_entries)})
        self.list_w.addTopLevelItem(image_item)
        global_index = 0
        for site_key, specs in groups.items():
            site_item = QTreeWidgetItem([self._site_label(specs)])
            site_item.setData(0, QtCore.Qt.UserRole, {
                "kind": "site",
                "site_key": site_key,
                "spec": specs[0],
                "specs": list(specs),
            })
            site_summary = str(specs[0].get("site_summary") or specs[0].get("site_display") or "").strip()
            if site_summary:
                site_item.setToolTip(0, site_summary)
            image_item.addChild(site_item)
            self._site_to_item[site_key] = site_item
            for spec in specs:
                global_index += 1
                it = QTreeWidgetItem([self._spec_label(spec, global_index)])
                it.setData(0, QtCore.Qt.UserRole, {"kind": "spec", "spec": spec})
                self._spec_to_item[self.viewer._spec_identity_key(spec) or str(global_index)] = it
                site_item.addChild(it)
            site_item.setExpanded(True)
        image_item.setExpanded(True)

    def _on_list_context_menu(self, pos):
        item = self.list_w.itemAt(pos)
        if item is None:
            return
        payload = item.data(0, QtCore.Qt.UserRole)
        if not isinstance(payload, dict):
            return
        kind = str(payload.get("kind") or "")
        menu = QtWidgets.QMenu(self.list_w)

        open_action = QtWidgets.QAction("Open", menu)
        open_action.triggered.connect(lambda _=False, item=item: self._on_item_double_clicked(item, 0))
        menu.addAction(open_action)

        specs = self._payload_specs(payload)
        source_path = specs[0].get("path") if specs else (self._file_key if kind == "image" else None)
        if source_path:
            menu.addSeparator()
            add_source_file_menu(menu, source_path, self)
        if kind in {"site", "spec"} and specs:
            menu.addSeparator()
            current_image_key = ""
            if hasattr(self.viewer, "_current_spectro_assignment_target_image_key"):
                current_image_key = str(self.viewer._current_spectro_assignment_target_image_key() or "")
            current_image_name = Path(current_image_key).name if current_image_key else ""
            assign_label = "Assign to current image"
            if current_image_name:
                assign_label = f"Assign to current image ({current_image_name})"
            assign_action = QtWidgets.QAction(assign_label, menu)
            assign_action.setEnabled(bool(current_image_key))
            if current_image_key and hasattr(self.viewer, "_apply_spectro_assignment_override"):
                assign_action.triggered.connect(
                    lambda _=False, specs=list(specs), image_key=current_image_key: self.viewer._apply_spectro_assignment_override(specs, image_key)
                )
            menu.addAction(assign_action)

            clear_action = QtWidgets.QAction("Clear manual assignment", menu)
            clear_action.setEnabled(any(str(spec.get("assignment_override_image_key") or "").strip() for spec in specs))
            if hasattr(self.viewer, "_clear_spectro_assignment_override"):
                clear_action.triggered.connect(
                    lambda _=False, specs=list(specs): self.viewer._clear_spectro_assignment_override(specs)
                )
            menu.addAction(clear_action)

        menu.exec_(self.list_w.viewport().mapToGlobal(pos))

# ---------- Viewer stubs / hooks for integration ----------
__all__ = [
    "SpectroSummaryDialog",
]





