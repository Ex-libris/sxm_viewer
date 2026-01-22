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

class ImageAdjustPreviewPanel(QtWidgets.QWidget):
    """
    Two-panel Matplotlib view:
      - workspace: where user edits crop (and optionally rotation by ctrl+right-drag)
      - preview: final result preview (with colorbar + optional scalebar)

    IMPORTANT: workspace pan/zoom is view-only and does not affect export.
    Crop coordinates are pixel indices with end-exclusive convention (x1/y1 are slice end).
    """
    selectionMade = QtCore.pyqtSignal(int, int, int, int)
    rotationChanged = QtCore.pyqtSignal(float)

    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self.workspace_label = QtWidgets.QLabel(
            "Workspace: drag to crop. Scroll to zoom. Middle-drag to pan. Double click to reset view."
        )
        layout.addWidget(self.workspace_label)

        self.workspace_fig = Figure(figsize=(4, 4))
        self.workspace_canvas = FigureCanvas(self.workspace_fig)
        self.workspace_ax = self.workspace_fig.add_subplot(111)
        layout.addWidget(self.workspace_canvas, 2)

        self.preview_label = QtWidgets.QLabel("Result preview")
        layout.addWidget(self.preview_label)

        self.preview_fig = Figure(figsize=(4, 4))
        self.preview_canvas = FigureCanvas(self.preview_fig)
        self.preview_ax = self.preview_fig.add_subplot(111)
        layout.addWidget(self.preview_canvas, 3)

        self.workspace_selector = RectangleSelector(
            self.workspace_ax,
            self._on_workspace_select,
            useblit=True,
            button=[1],
            minspanx=2,
            minspany=2,
            interactive=True,
            props=dict(edgecolor='#ffca28', facecolor='none', linewidth=1.5),
        )

        self._base_shape = (1, 1)
        self._axis_unit = 'px'
        self._crop_spec = {'x0': 0, 'y0': 0, 'x1': 1, 'y1': 1}

        self._current_rotation = 0.0
        self._rotation_drag = None
        self._pan_drag = None
        self._workspace_xlim = None
        self._workspace_ylim = None

        self._result_cbar = None
        self._result_cbar_ax = None
        self._result_scalebar = None
        self._scalebar_enabled = True
        self._colorbar_label = ''

        self.workspace_ax.set_facecolor('#0b0b0b')
        self.preview_ax.set_facecolor('#0b0b0b')
        self.workspace_ax.xaxis.set_major_formatter(FuncFormatter(self._workspace_tick_format_x))
        self.workspace_ax.yaxis.set_major_formatter(FuncFormatter(self._workspace_tick_format_y))

        self.workspace_canvas.mpl_connect('button_press_event', self._on_press)
        self.workspace_canvas.mpl_connect('motion_notify_event', self._on_motion)
        self.workspace_canvas.mpl_connect('button_release_event', self._on_release)
        self.workspace_canvas.mpl_connect('scroll_event', self._on_scroll)

    def set_colorbar_label(self, text):
        self._colorbar_label = text or ''

    def set_rotation_angle(self, angle):
        self._current_rotation = float(angle)

    def set_base_geometry(self, shape, axis_unit=None):
        self._base_shape = tuple(shape)
        if axis_unit:
            self._axis_unit = axis_unit
        self.reset_workspace_view()

    def reset_workspace_view(self):
        self._workspace_xlim = None
        self._workspace_ylim = None

    def update_workspace(self, arr, axis_unit, crop_spec, cmap_name):
        self._axis_unit = axis_unit or self._axis_unit
        self._crop_spec = crop_spec or self._crop_spec

        self.workspace_ax.clear()
        a = np.asarray(arr)
        a = np.flipud(a)
        h, w = a.shape[:2]
        self.workspace_ax.imshow(a, extent=(0, w, 0, h), origin='lower', aspect='equal', cmap=cmap_name)

        if self._workspace_xlim is None or self._workspace_ylim is None:
            self.workspace_ax.set_xlim(0, w)
            self.workspace_ax.set_ylim(0, h)
        else:
            self.workspace_ax.set_xlim(*self._workspace_xlim)
            self.workspace_ax.set_ylim(*self._workspace_ylim)

        self.workspace_ax.set_title("Crop workspace", fontsize=10)
        self.workspace_ax.set_xlabel(f"x [{self._axis_unit}]")
        self.workspace_ax.set_ylabel(f"y [{self._axis_unit}]")

        self._apply_crop_rectangle()
        try:
            self.workspace_fig.tight_layout()
        except Exception:
            pass
        self.workspace_canvas.draw_idle()

    def update_result(self, arr, extent, axis_unit, cmap_name, scalebar_enabled):
        self._scalebar_enabled = bool(scalebar_enabled)

        self.preview_ax.clear()
        a = np.asarray(arr)
        a = np.flipud(a)
        a = np.ma.masked_invalid(a)

        if extent is None:
            h, w = a.shape[:2]
            im = self.preview_ax.imshow(a, extent=(0, w, 0, h), origin='lower', aspect='equal', cmap=cmap_name)
        else:
            im = self.preview_ax.imshow(a, extent=extent, origin='lower', aspect='equal', cmap=cmap_name)

        self.preview_ax.set_title("Result preview", fontsize=10)
        self.preview_ax.set_xlabel(f"x [{axis_unit or 'px'}]")
        self.preview_ax.set_ylabel(f"y [{axis_unit or 'px'}]")

        # stable inset colorbar
        if self._result_cbar is not None:
            try:
                self._result_cbar.remove()
            except Exception:
                pass
            self._result_cbar = None
        if self._result_cbar_ax is not None:
            try:
                self._result_cbar_ax.remove()
            except Exception:
                pass
            self._result_cbar_ax = None

        try:
            self._result_cbar_ax = inset_axes(
                self.preview_ax, width="3%", height="85%",
                loc='center left',
                bbox_to_anchor=(1.02, 0.08, 1, 1),
                bbox_transform=self.preview_ax.transAxes,
                borderpad=0
            )
            self._result_cbar = self.preview_fig.colorbar(im, cax=self._result_cbar_ax)
        except Exception:
            self._result_cbar = self.preview_fig.colorbar(im, ax=self.preview_ax, fraction=0.046, pad=0.02)

        if self._result_cbar is not None and self._colorbar_label:
            try:
                self._result_cbar.set_label(self._colorbar_label)
            except Exception:
                pass

        # scalebar
        if self._result_scalebar is not None:
            try:
                self._result_scalebar.remove()
            except Exception:
                pass
            self._result_scalebar = None

        if self._scalebar_enabled and extent is not None:
            length = self._nice_length(abs(float(extent[1]) - float(extent[0])))
            if length > 0:
                bar = AnchoredSizeBar(
                    self.preview_ax.transData, length,
                    f"{self._format_value(length)} {axis_unit or ''}".strip(),
                    'lower right', pad=0.35, color='white',
                    frameon=True, size_vertical=0.4
                )
                self.preview_ax.add_artist(bar)
                self._result_scalebar = bar

        try:
            self.preview_fig.tight_layout()
        except Exception:
            pass
        self.preview_canvas.draw_idle()

    # ---------- interaction ----------
    def _on_workspace_select(self, eclick, erelease):
        if eclick.xdata is None or erelease.xdata is None:
            return
        h, w = self._base_shape
        x0f = float(min(eclick.xdata, erelease.xdata))
        x1f = float(max(eclick.xdata, erelease.xdata))
        y0f = float(min(eclick.ydata, erelease.ydata))
        y1f = float(max(eclick.ydata, erelease.ydata))

        x0 = int(np.clip(np.floor(x0f), 0, max(0, w - 1)))
        y0 = int(np.clip(np.floor(y0f), 0, max(0, h - 1)))
        x1 = int(np.clip(np.ceil(x1f), x0 + 1, w))
        y1 = int(np.clip(np.ceil(y1f), y0 + 1, h))
        self.selectionMade.emit(x0, x1, y0, y1)

    def _on_press(self, event):
        if event.inaxes is not self.workspace_ax:
            return
        if getattr(event, 'dblclick', False):
            self.reset_workspace_view()
            self.workspace_canvas.draw_idle()
            return
        if event.button == 2 and event.xdata is not None and event.ydata is not None:
            self._pan_drag = (float(event.xdata), float(event.ydata),
                              tuple(self.workspace_ax.get_xlim()), tuple(self.workspace_ax.get_ylim()))
            return
        if event.button == 3 and event.x is not None:
            key = (event.key or '').lower()
            if 'control' in key:
                self._rotation_drag = (event.x, self._current_rotation)

    def _on_motion(self, event):
        if event.inaxes is not self.workspace_ax:
            return
        if self._pan_drag and event.xdata is not None and event.ydata is not None:
            sx, sy, xlim0, ylim0 = self._pan_drag
            dx = sx - float(event.xdata)
            dy = sy - float(event.ydata)
            x0, x1 = xlim0[0] + dx, xlim0[1] + dx
            y0, y1 = ylim0[0] + dy, ylim0[1] + dy
            h, w = self._base_shape
            x0, x1 = self._clamp_span(x0, x1, 0.0, float(w))
            y0, y1 = self._clamp_span(y0, y1, 0.0, float(h))
            self._workspace_xlim = (x0, x1)
            self._workspace_ylim = (y0, y1)
            self.workspace_ax.set_xlim(x0, x1)
            self.workspace_ax.set_ylim(y0, y1)
            self.workspace_canvas.draw_idle()
            return
        if self._rotation_drag and event.x is not None:
            start_x, start_angle = self._rotation_drag
            delta = event.x - start_x
            new_angle = float(np.clip(start_angle + delta * 0.4, -180.0, 180.0))
            self.rotationChanged.emit(new_angle)

    def _on_release(self, event):
        self._rotation_drag = None
        self._pan_drag = None

    def _on_scroll(self, event):
        if event.inaxes is not self.workspace_ax:
            return
        if event.xdata is None or event.ydata is None:
            return
        base = 1.15
        if getattr(event, 'button', None) == 'up':
            scale = 1.0 / base
        elif getattr(event, 'button', None) == 'down':
            scale = base
        else:
            return
        x = float(event.xdata)
        y = float(event.ydata)
        x0, x1 = self.workspace_ax.get_xlim()
        y0, y1 = self.workspace_ax.get_ylim()
        nx0 = x - (x - x0) * scale
        nx1 = x + (x1 - x) * scale
        ny0 = y - (y - y0) * scale
        ny1 = y + (y1 - y) * scale
        h, w = self._base_shape
        nx0, nx1 = self._clamp_span(nx0, nx1, 0.0, float(w))
        ny0, ny1 = self._clamp_span(ny0, ny1, 0.0, float(h))
        self._workspace_xlim = (nx0, nx1)
        self._workspace_ylim = (ny0, ny1)
        self.workspace_ax.set_xlim(nx0, nx1)
        self.workspace_ax.set_ylim(ny0, ny1)
        self.workspace_canvas.draw_idle()

    # ---------- helpers ----------
    def _apply_crop_rectangle(self):
        if not self._crop_spec:
            return
        h, w = self._base_shape
        x0 = float(self._crop_spec.get('x0', 0))
        x1 = float(self._crop_spec.get('x1', w))
        y0 = float(self._crop_spec.get('y0', 0))
        y1 = float(self._crop_spec.get('y1', h))
        x0 = max(0.0, min(x0, float(max(0, w - 1))))
        y0 = max(0.0, min(y0, float(max(0, h - 1))))
        x1 = max(x0 + 1.0, min(x1, float(w)))
        y1 = max(y0 + 1.0, min(y1, float(h)))
        self.workspace_selector.set_active(False)
        self.workspace_selector.extents = (x0, x1, y0, y1)
        self.workspace_selector.set_active(True)

    def _workspace_tick_format_x(self, value, pos=None):
        return f"{value:g}"

    def _workspace_tick_format_y(self, value, pos=None):
        return f"{value:g}"

    def _nice_length(self, length):
        if length <= 0:
            return 0.0
        # 1-2-5 style
        exp = math.floor(math.log10(length))
        base = 10 ** exp
        scaled = length / base
        if scaled < 2:
            return 1 * base
        if scaled < 5:
            return 2 * base
        return 5 * base

    def _format_value(self, v):
        if abs(v) >= 10:
            return f"{v:.0f}"
        if abs(v) >= 1:
            return f"{v:.1f}"
        return f"{v:.2f}"

    def _clamp_span(self, a, b, lo, hi):
        span = b - a
        if span <= 0:
            return lo, hi
        if a < lo:
            a = lo
            b = lo + span
        if b > hi:
            b = hi
            a = hi - span
        a = max(lo, a)
        b = min(hi, b)
        return a, b

class ImageAdjustDialog(QtWidgets.QDialog):
    """
    Key change vs the old behavior:
      - Crop can be applied either on the ORIGINAL image, or AFTER rotation (rotated coordinate space).
        This fixes the "wonky" feeling where the user rotates then tries to crop, but crop still refers to the original.
    """
    def __init__(self, parent, base_image, spec, cmap_name, base_extent=None, display_extent=None,
                 axis_unit=None, colorbar_label=None, base_unit=None, relative_axes=False):
        super().__init__(parent)
        self.viewer = parent
        self.setWindowTitle("Image adjustments")

        self.base_image = np.asarray(base_image)
        self.base_extent = base_extent
        self.axis_unit = axis_unit or 'px'
        self.colorbar_label = colorbar_label or ''

        # spec fields:
        #   crop: always stored, interpreted depending on crop_mode
        #   crop_mode: 'pre' or 'post'
        self.current_spec = json.loads(json.dumps(spec or {}))
        self.current_spec.setdefault('crop', {'x0': 0, 'y0': 0, 'x1': self.base_image.shape[1], 'y1': self.base_image.shape[0]})
        self.current_spec.setdefault('rotate', 0.0)
        self.current_spec.setdefault('flip_h', False)
        self.current_spec.setdefault('flip_v', False)
        self.current_spec.setdefault('clip', {'low': None, 'high': None})
        self.current_spec.setdefault('gamma', 1.0)
        self.current_spec.setdefault('cmap', cmap_name or 'viridis')
        self.current_spec.setdefault('crop_mode', 'pre')  # 'pre' = crop before rotate, 'post' = crop after rotate
        self.current_spec.setdefault('lock_square', False)
        self.current_spec.setdefault('auto_trim', True)

        self._undo_stack = []
        self._redo_stack = []
        self._updating_controls = False
        self._live_prev_spec = None

        self._preview_timer = QtCore.QTimer(self)
        self._preview_timer.setSingleShot(True)
        self._preview_timer.timeout.connect(self._update_preview)

        # ---- layout ----
        main_layout = QtWidgets.QHBoxLayout(self)

        controls = QtWidgets.QWidget()
        controls_layout = QtWidgets.QVBoxLayout(controls)
        controls_layout.setContentsMargins(10, 10, 10, 10)
        controls_layout.setSpacing(10)
        main_layout.addWidget(controls, 0)

        # Crop group
        crop_group = QtWidgets.QGroupBox("Crop (pixels)")
        crop_form = QtWidgets.QFormLayout(crop_group)
        self.x0_spin = QtWidgets.QSpinBox(); self.x0_spin.setRange(0, self.base_image.shape[1]-1)
        self.x1_spin = QtWidgets.QSpinBox(); self.x1_spin.setRange(1, self.base_image.shape[1])
        self.y0_spin = QtWidgets.QSpinBox(); self.y0_spin.setRange(0, self.base_image.shape[0]-1)
        self.y1_spin = QtWidgets.QSpinBox(); self.y1_spin.setRange(1, self.base_image.shape[0])
        crop_form.addRow("X start", self.x0_spin)
        crop_form.addRow("X end", self.x1_spin)
        crop_form.addRow("Y start", self.y0_spin)
        crop_form.addRow("Y end", self.y1_spin)
        controls_layout.addWidget(crop_group)

        # Geometry group
        geom_group = QtWidgets.QGroupBox("Geometry")
        geom_layout = QtWidgets.QVBoxLayout(geom_group)

        rot_row = QtWidgets.QHBoxLayout()
        rot_row.addWidget(QtWidgets.QLabel("Rotate (deg)"))
        self.rotate_slider = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        self.rotate_slider.setRange(-180, 180)
        self.rotate_value_label = QtWidgets.QLabel("0 deg")
        rot_row.addWidget(self.rotate_slider, 1)
        rot_row.addWidget(self.rotate_value_label)
        geom_layout.addLayout(rot_row)

        self.flip_h_cb = QtWidgets.QCheckBox("Flip horizontally")
        self.flip_v_cb = QtWidgets.QCheckBox("Flip vertically")
        geom_layout.addWidget(self.flip_h_cb)
        geom_layout.addWidget(self.flip_v_cb)
        controls_layout.addWidget(geom_group)

        # Tone mapping group
        tone_group = QtWidgets.QGroupBox("Tone mapping")
        tone_form = QtWidgets.QFormLayout(tone_group)
        self.low_pct_spin = QtWidgets.QDoubleSpinBox(); self.low_pct_spin.setRange(0.0, 100.0); self.low_pct_spin.setDecimals(2)
        self.high_pct_spin = QtWidgets.QDoubleSpinBox(); self.high_pct_spin.setRange(0.0, 100.0); self.high_pct_spin.setDecimals(2)
        self.gamma_spin = QtWidgets.QDoubleSpinBox(); self.gamma_spin.setRange(0.05, 10.0); self.gamma_spin.setDecimals(2); self.gamma_spin.setSingleStep(0.05)
        tone_form.addRow("Clip low %", self.low_pct_spin)
        tone_form.addRow("Clip high %", self.high_pct_spin)
        tone_form.addRow("Gamma", self.gamma_spin)
        controls_layout.addWidget(tone_group)

        # buttons
        btn_row = QtWidgets.QHBoxLayout()
        self.undo_btn = QtWidgets.QPushButton("Undo")
        self.redo_btn = QtWidgets.QPushButton("Redo")
        self.reset_btn = QtWidgets.QPushButton("Reset")
        btn_row.addWidget(self.undo_btn)
        btn_row.addWidget(self.redo_btn)
        btn_row.addWidget(self.reset_btn)
        controls_layout.addLayout(btn_row)

        # Colormap
        cmap_row = QtWidgets.QHBoxLayout()
        cmap_row.addWidget(QtWidgets.QLabel("Colormap:"))
        self.cmap_combo = QtWidgets.QComboBox()
        # You already likely populate this elsewhere; keep a safe default here:
        for name in sorted(set(['viridis', 'plasma', 'inferno', 'magma', 'cividis', str(self.current_spec.get('cmap', 'viridis'))])):
            self.cmap_combo.addItem(name)
        cmap_row.addWidget(self.cmap_combo, 1)
        controls_layout.addLayout(cmap_row)

        # Options
        opt_group = QtWidgets.QGroupBox("Options")
        opt_layout = QtWidgets.QVBoxLayout(opt_group)

        self.scalebar_cb = QtWidgets.QCheckBox("Show scalebar")
        self.crop_mode_combo = QtWidgets.QComboBox()
        self.crop_mode_combo.addItems(["Crop on original", "Crop after rotation"])
        self.lock_square_cb = QtWidgets.QCheckBox("Square crop")
        self.auto_trim_cb = QtWidgets.QCheckBox("Auto-crop rotation border")

        opt_layout.addWidget(self.scalebar_cb)

        cm_row = QtWidgets.QHBoxLayout()
        cm_row.addWidget(QtWidgets.QLabel("Crop mode"))
        cm_row.addWidget(self.crop_mode_combo, 1)
        opt_layout.addLayout(cm_row)

        opt_layout.addWidget(self.lock_square_cb)
        opt_layout.addWidget(self.auto_trim_cb)

        controls_layout.addWidget(opt_group)
        controls_layout.addStretch(1)

        # preview
        preview_widget = QtWidgets.QWidget()
        preview_layout = QtWidgets.QVBoxLayout(preview_widget)
        preview_layout.setContentsMargins(6, 10, 10, 10)
        self.preview_panel = ImageAdjustPreviewPanel()
        self.preview_panel.set_colorbar_label(self.colorbar_label)
        preview_layout.addWidget(self.preview_panel)
        main_layout.addWidget(preview_widget, 1)

        btn_box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        preview_layout.addWidget(btn_box)
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)

        # connections
        for spin in (self.x0_spin, self.x1_spin, self.y0_spin, self.y1_spin,
                     self.low_pct_spin, self.high_pct_spin, self.gamma_spin):
            spin.valueChanged.connect(self._on_params_changed_live)
            if hasattr(spin, 'editingFinished'):
                spin.editingFinished.connect(self._commit_live_change)

        self.rotate_slider.sliderPressed.connect(self._begin_live_change)
        self.rotate_slider.valueChanged.connect(self._on_params_changed_live)
        self.rotate_slider.sliderReleased.connect(self._commit_live_change)

        for w in (self.flip_h_cb, self.flip_v_cb, self.scalebar_cb, self.lock_square_cb, self.auto_trim_cb):
            w.toggled.connect(self._on_discrete_change)
        self.cmap_combo.currentIndexChanged.connect(self._on_discrete_change)
        self.crop_mode_combo.currentIndexChanged.connect(self._on_crop_mode_changed)

        self.undo_btn.clicked.connect(self._on_undo)
        self.redo_btn.clicked.connect(self._on_redo)
        self.reset_btn.clicked.connect(self._on_reset)

        self.preview_panel.selectionMade.connect(self._on_crop_selection)
        self.preview_panel.rotationChanged.connect(self._on_workspace_rotation_drag)

        self._apply_spec_to_controls()
        self._update_preview()

    # ---------- state / history ----------
    def _push_history(self, prev_spec):
        if prev_spec is None or prev_spec == self.current_spec:
            return
        self._undo_stack.append(json.loads(json.dumps(prev_spec)))
        if len(self._undo_stack) > 50:
            self._undo_stack.pop(0)
        self._redo_stack.clear()

    def _begin_live_change(self):
        if self._updating_controls:
            return
        if self._live_prev_spec is None:
            self._live_prev_spec = json.loads(json.dumps(self.current_spec))

    def _commit_live_change(self):
        if self._updating_controls:
            return
        if self._live_prev_spec is None:
            return
        self._push_history(self._live_prev_spec)
        self._live_prev_spec = None

    def _schedule_preview_update(self):
        if self._preview_timer.isActive():
            self._preview_timer.stop()
        self._preview_timer.start(40)

    # ---------- UI <-> spec ----------
    def _apply_spec_to_controls(self):
        self._updating_controls = True
        crop = self.current_spec.get('crop', {})
        self.x0_spin.setValue(int(crop.get('x0', 0)))
        self.x1_spin.setValue(int(crop.get('x1', self.base_image.shape[1])))
        self.y0_spin.setValue(int(crop.get('y0', 0)))
        self.y1_spin.setValue(int(crop.get('y1', self.base_image.shape[0])))

        self.rotate_slider.setValue(int(round(float(self.current_spec.get('rotate', 0.0) or 0.0))))
        self.rotate_value_label.setText(f"{self.rotate_slider.value()} deg")

        self.flip_h_cb.setChecked(bool(self.current_spec.get('flip_h', False)))
        self.flip_v_cb.setChecked(bool(self.current_spec.get('flip_v', False)))

        clip = self.current_spec.get('clip', {}) or {}
        self.low_pct_spin.setValue(float(clip.get('low', 0.0) or 0.0))
        self.high_pct_spin.setValue(float(clip.get('high', 100.0) or 100.0))
        self.gamma_spin.setValue(float(self.current_spec.get('gamma', 1.0) or 1.0))

        self.scalebar_cb.setChecked(True)
        self.lock_square_cb.setChecked(bool(self.current_spec.get('lock_square', False)))
        self.auto_trim_cb.setChecked(bool(self.current_spec.get('auto_trim', True)))

        mode = self.current_spec.get('crop_mode', 'pre')
        self.crop_mode_combo.setCurrentIndex(0 if mode == 'pre' else 1)

        cmap = self.current_spec.get('cmap', self.cmap_combo.currentText())
        # If cmap exists in combo, select it; else add.
        if self.cmap_combo.findText(cmap) < 0:
            self.cmap_combo.addItem(cmap)
        self.cmap_combo.setCurrentText(cmap)

        self._sanitize_crop_controls()
        self._updating_controls = False

    def _sanitize_crop_controls(self):
        h, w = self.base_image.shape[:2]
        x0 = int(self.x0_spin.value())
        x1 = int(self.x1_spin.value())
        y0 = int(self.y0_spin.value())
        y1 = int(self.y1_spin.value())

        x0 = max(0, min(x0, max(0, w - 1)))
        y0 = max(0, min(y0, max(0, h - 1)))
        x1 = max(x0 + 1, min(x1, w))
        y1 = max(y0 + 1, min(y1, h))

        self.x0_spin.blockSignals(True); self.x1_spin.blockSignals(True)
        self.y0_spin.blockSignals(True); self.y1_spin.blockSignals(True)
        self.x0_spin.setValue(x0); self.x1_spin.setValue(x1)
        self.y0_spin.setValue(y0); self.y1_spin.setValue(y1)
        self.x0_spin.blockSignals(False); self.x1_spin.blockSignals(False)
        self.y0_spin.blockSignals(False); self.y1_spin.blockSignals(False)

        self.x0_spin.setMaximum(max(0, x1 - 1))
        self.x1_spin.setMinimum(min(w, x0 + 1))
        self.y0_spin.setMaximum(max(0, y1 - 1))
        self.y1_spin.setMinimum(min(h, y0 + 1))

    def _collect_spec_from_controls(self):
        self._sanitize_crop_controls()
        low = float(self.low_pct_spin.value())
        high = float(self.high_pct_spin.value())
        if high < low:
            high = low
            self.high_pct_spin.setValue(high)

        mode = 'pre' if self.crop_mode_combo.currentIndex() == 0 else 'post'
        spec = {
            'crop': {
                'x0': int(self.x0_spin.value()),
                'x1': int(self.x1_spin.value()),
                'y0': int(self.y0_spin.value()),
                'y1': int(self.y1_spin.value()),
            },
            'crop_mode': mode,
            'rotate': float(self.rotate_slider.value()),
            'flip_h': self.flip_h_cb.isChecked(),
            'flip_v': self.flip_v_cb.isChecked(),
            'clip': {
                'low': low if low > 0 else None,
                'high': high if high < 100 else None,
            },
            'gamma': float(self.gamma_spin.value()),
            'cmap': self.cmap_combo.currentText(),
            'lock_square': self.lock_square_cb.isChecked(),
            'auto_trim': self.auto_trim_cb.isChecked(),
        }
        return spec

    # ---------- callbacks ----------
    def _on_params_changed_live(self, value=None):
        if self._updating_controls:
            return
        self.current_spec = self._collect_spec_from_controls()
        self.rotate_value_label.setText(f"{int(round(self.rotate_slider.value()))} deg")
        self.preview_panel.set_rotation_angle(float(self.current_spec.get('rotate', 0.0)))
        self._schedule_preview_update()

    def _on_discrete_change(self, value=None):
        if self._updating_controls:
            return
        prev = json.loads(json.dumps(self.current_spec))
        self.current_spec = self._collect_spec_from_controls()
        self.rotate_value_label.setText(f"{int(round(self.rotate_slider.value()))} deg")
        self.preview_panel.set_rotation_angle(float(self.current_spec.get('rotate', 0.0)))
        self._push_history(prev)
        self._schedule_preview_update()

    def _on_crop_mode_changed(self, idx):
        if self._updating_controls:
            return
        # Crop coordinates refer to a different image. Do not try to map.
        # Reset to full frame and let user reselect cleanly.
        prev = json.loads(json.dumps(self.current_spec))
        if idx == 0:
            self.current_spec['crop_mode'] = 'pre'
            h, w = self.base_image.shape[:2]
        else:
            self.current_spec['crop_mode'] = 'post'
            # rotated workspace shape depends on rotation; we set full crop later in _update_preview
            h, w = self.base_image.shape[:2]
        self.current_spec['crop'] = {'x0': 0, 'y0': 0, 'x1': int(w), 'y1': int(h)}
        self._push_history(prev)
        self._apply_spec_to_controls()
        self._schedule_preview_update()

    def _on_crop_selection(self, x0, x1, y0, y1):
        if self._updating_controls:
            return
        prev = json.loads(json.dumps(self.current_spec))
        if self.lock_square_cb.isChecked():
            # shrink to square anchored at (x0,y0)
            size = max(1, min(int(x1) - int(x0), int(y1) - int(y0)))
            x1 = int(x0) + size
            y1 = int(y0) + size
        self._updating_controls = True
        self.x0_spin.setValue(int(x0))
        self.x1_spin.setValue(int(x1))
        self.y0_spin.setValue(int(y0))
        self.y1_spin.setValue(int(y1))
        self._updating_controls = False
        self.current_spec = self._collect_spec_from_controls()
        self._push_history(prev)
        self._schedule_preview_update()

    def _on_workspace_rotation_drag(self, angle):
        angle = int(np.clip(round(angle), -180, 180))
        if self.rotate_slider.value() == angle:
            return
        self._begin_live_change()
        self.rotate_slider.setValue(angle)

    def _on_undo(self):
        if not self._undo_stack:
            return
        self._redo_stack.append(json.loads(json.dumps(self.current_spec)))
        self.current_spec = self._undo_stack.pop()
        self._apply_spec_to_controls()
        self._update_preview()

    def _on_redo(self):
        if not self._redo_stack:
            return
        self._undo_stack.append(json.loads(json.dumps(self.current_spec)))
        self.current_spec = self._redo_stack.pop()
        self._apply_spec_to_controls()
        self._update_preview()

    def _on_reset(self):
        prev = json.loads(json.dumps(self.current_spec))
        h, w = self.base_image.shape[:2]
        self.current_spec = {
            'crop': {'x0': 0, 'y0': 0, 'x1': int(w), 'y1': int(h)},
            'crop_mode': 'pre',
            'rotate': 0.0,
            'flip_h': False,
            'flip_v': False,
            'clip': {'low': None, 'high': None},
            'gamma': 1.0,
            'cmap': self.current_spec.get('cmap', 'viridis'),
            'lock_square': False,
            'auto_trim': True,
        }
        self._push_history(prev)
        self._apply_spec_to_controls()
        self._update_preview()

    # ---------- processing ----------
    def _update_preview(self):
        cmap = self.current_spec.get('cmap', 'viridis') or 'viridis'
        mode = self.current_spec.get('crop_mode', 'pre')

        # 1) Build a "geometry-only" spec to generate the rotated workspace when needed.
        geom_spec = json.loads(json.dumps(self.current_spec))
        geom_spec['clip'] = {'low': None, 'high': None}
        geom_spec['gamma'] = 1.0

        # For geometry-only pass, do not crop (full frame).
        h0, w0 = self.base_image.shape[:2]
        geom_spec['crop'] = {'x0': 0, 'y0': 0, 'x1': int(w0), 'y1': int(h0)}

        # Use your existing engine (thumbnails.py) so export stays consistent.
        geom_arr, geom_extent = apply_adjustment_spec(self.base_image, self.base_extent, geom_spec)

        if mode == 'post':
            # Workspace is the rotated image.
            self.preview_panel.set_base_geometry(geom_arr.shape[:2], axis_unit=self.axis_unit)
            ws_crop = self.current_spec.get('crop', {'x0': 0, 'y0': 0, 'x1': geom_arr.shape[1], 'y1': geom_arr.shape[0]})
            self.preview_panel.update_workspace(geom_arr, self.axis_unit, ws_crop, cmap)
        else:
            # Workspace is the original image.
            self.preview_panel.set_base_geometry(self.base_image.shape[:2], axis_unit=self.axis_unit)
            ws_crop = self.current_spec.get('crop', {'x0': 0, 'y0': 0, 'x1': w0, 'y1': h0})
            self.preview_panel.update_workspace(self.base_image, self.axis_unit, ws_crop, cmap)

        # 2) Compute final result.
        if mode == 'pre':
            arr_result, extent_result = apply_adjustment_spec(self.base_image, self.base_extent, self.current_spec)
        else:
            # Apply geometry first, then crop in rotated coordinates, then tone-map.
            crop = self.current_spec.get('crop', {})
            x0 = int(crop.get('x0', 0)); x1 = int(crop.get('x1', geom_arr.shape[1]))
            y0 = int(crop.get('y0', 0)); y1 = int(crop.get('y1', geom_arr.shape[0]))
            x0 = max(0, min(x0, geom_arr.shape[1]-1))
            y0 = max(0, min(y0, geom_arr.shape[0]-1))
            x1 = max(x0+1, min(x1, geom_arr.shape[1]))
            y1 = max(y0+1, min(y1, geom_arr.shape[0]))
            cropped = geom_arr[y0:y1, x0:x1]
            extent_result = self._crop_extent(geom_extent, geom_arr.shape, x0, x1, y0, y1)
            # Now apply clip/gamma only (no extra geometry).
            tone_spec = json.loads(json.dumps(self.current_spec))
            tone_spec['crop'] = {'x0': 0, 'y0': 0, 'x1': cropped.shape[1], 'y1': cropped.shape[0]}
            tone_spec['rotate'] = 0.0
            tone_spec['flip_h'] = False
            tone_spec['flip_v'] = False
            arr_result, extent_result = apply_adjustment_spec(cropped, extent_result, tone_spec)

        # Optional trim to rectangle after rotation.
        if self.auto_trim_cb.isChecked() and abs(float(self.current_spec.get('rotate', 0.0) or 0.0)) > 1e-6:
            arr_result, extent_result = self._trim_finite_border(arr_result, extent_result)

        self.preview_panel.update_result(arr_result, extent_result, self.axis_unit, cmap, self.scalebar_cb.isChecked())

    def _crop_extent(self, extent, full_shape, x0, x1, y0, y1):
        if extent is None:
            return None
        try:
            xmin, xmax, ymin, ymax = extent
            h, w = full_shape[:2]
            dx = (float(xmax) - float(xmin)) / max(1, w)
            dy = (float(ymax) - float(ymin)) / max(1, h)
            return [float(xmin) + dx * x0, float(xmin) + dx * x1,
                    float(ymin) + dy * y0, float(ymin) + dy * y1]
        except Exception:
            return extent

    def _trim_finite_border(self, arr, extent):
        a = np.asarray(arr, dtype=float)
        if a.size == 0:
            return arr, extent
        mask = np.isfinite(a)
        if not mask.any():
            return arr, extent
        ys = np.where(mask.any(axis=1))[0]
        xs = np.where(mask.any(axis=0))[0]
        y0 = int(ys[0]); y1 = int(ys[-1]) + 1
        x0 = int(xs[0]); x1 = int(xs[-1]) + 1
        trimmed = a[y0:y1, x0:x1]
        new_extent = self._crop_extent(extent, a.shape, x0, x1, y0, y1)
        return trimmed, new_extent
# === END: Image adjustment classes (drop-in replacement) ===





