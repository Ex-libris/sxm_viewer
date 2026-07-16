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
from ... import cmap_registry
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
    resample_geometry,
    largest_inscribed_rect,
    geometry_is_identity,
    _rotate_extent_box,
    _trim_nan_border,
    save_wsxm_xyz,
)

class ImageAdjustPreviewPanel(QtWidgets.QWidget):
    """
    Two-panel Matplotlib view:
      - workspace: always shows the *rotated* image (flips applied); the
        crop rectangle is drawn and edited in that rotated frame, clamped
        by the dialog to the largest inscribed rectangle so the result can
        never contain out-of-source (NaN) pixels. Rotation via
        ctrl+right-drag or the slider.
      - preview: final result preview (with colorbar + optional scalebar),
        rendered by the exact same engine that produces the virtual copy.

    IMPORTANT: workspace pan/zoom is view-only and does not affect export.
    Crop coordinates are (left, right, bottom, top) in rotated-frame
    source-pixel units (y up), matching thumbnail_render.resample_geometry.
    """
    selectionMade = QtCore.pyqtSignal(float, float, float, float)
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

        self._ws_bounds = (0.0, 1.0, 0.0, 1.0)  # (x0, x1, y0, y1) workspace extent
        self._axis_unit = 'px'
        self._crop_rect = (0.0, 1.0, 0.0, 1.0)  # (left, right, bottom, top)
        self._inscribed_rect = None
        self._inscribed_patch = None

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

    def set_workspace_bounds(self, bounds, axis_unit=None):
        """bounds = (x0, x1, y0, y1) of the rotated workspace frame."""
        self._ws_bounds = tuple(float(v) for v in bounds)
        if axis_unit:
            self._axis_unit = axis_unit
        self.reset_workspace_view()

    def reset_workspace_view(self):
        self._workspace_xlim = None
        self._workspace_ylim = None

    def update_workspace(self, arr, ws_bounds, axis_unit, crop_rect, cmap_name,
                         inscribed_rect=None):
        """Show the rotated image (already resampled by the dialog) whose
        display extent is ws_bounds, with the crop selector at crop_rect
        and a dashed guide at the largest inscribed rectangle."""
        self._axis_unit = axis_unit or self._axis_unit
        if ws_bounds is not None and tuple(ws_bounds) != self._ws_bounds:
            self.set_workspace_bounds(ws_bounds, axis_unit)
        if crop_rect is not None:
            self._crop_rect = tuple(float(v) for v in crop_rect)
        self._inscribed_rect = tuple(float(v) for v in inscribed_rect) if inscribed_rect else None

        self.workspace_ax.clear()
        self._inscribed_patch = None
        a = np.ma.masked_invalid(np.flipud(np.asarray(arr)))
        x0, x1, y0, y1 = self._ws_bounds
        self.workspace_ax.imshow(a, extent=(x0, x1, y0, y1), origin='lower',
                                 aspect='equal', cmap=cmap_name)

        if self._workspace_xlim is None or self._workspace_ylim is None:
            self.workspace_ax.set_xlim(x0, x1)
            self.workspace_ax.set_ylim(y0, y1)
        else:
            self.workspace_ax.set_xlim(*self._workspace_xlim)
            self.workspace_ax.set_ylim(*self._workspace_ylim)

        self.workspace_ax.set_title("Crop workspace (rotated frame)", fontsize=10)
        self.workspace_ax.set_xlabel(f"x [{self._axis_unit}]")
        self.workspace_ax.set_ylabel(f"y [{self._axis_unit}]")

        if self._inscribed_rect is not None:
            il, ir, ib, it = self._inscribed_rect
            self._inscribed_patch = patches.Rectangle(
                (il, ib), ir - il, it - ib, fill=False,
                edgecolor='#4dd0e1', linestyle='--', linewidth=1.0, alpha=0.8)
            self.workspace_ax.add_patch(self._inscribed_patch)

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
        left = float(min(eclick.xdata, erelease.xdata))
        right = float(max(eclick.xdata, erelease.xdata))
        bottom = float(min(eclick.ydata, erelease.ydata))
        top = float(max(eclick.ydata, erelease.ydata))
        if right - left < 1.0 or top - bottom < 1.0:
            return
        # The dialog clamps to the inscribed rectangle and square lock.
        self.selectionMade.emit(left, right, bottom, top)

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
            wx0, wx1, wy0, wy1 = self._ws_bounds
            x0, x1 = self._clamp_span(x0, x1, wx0, wx1)
            y0, y1 = self._clamp_span(y0, y1, wy0, wy1)
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
        wx0, wx1, wy0, wy1 = self._ws_bounds
        nx0, nx1 = self._clamp_span(nx0, nx1, wx0, wx1)
        ny0, ny1 = self._clamp_span(ny0, ny1, wy0, wy1)
        self._workspace_xlim = (nx0, nx1)
        self._workspace_ylim = (ny0, ny1)
        self.workspace_ax.set_xlim(nx0, nx1)
        self.workspace_ax.set_ylim(ny0, ny1)
        self.workspace_canvas.draw_idle()

    # ---------- helpers ----------
    def _apply_crop_rectangle(self):
        if not self._crop_rect:
            return
        left, right, bottom, top = self._crop_rect
        wx0, wx1, wy0, wy1 = self._ws_bounds
        left = max(wx0, min(left, wx1 - 1.0))
        bottom = max(wy0, min(bottom, wy1 - 1.0))
        right = max(left + 1.0, min(right, wx1))
        top = max(bottom + 1.0, min(top, wy1))
        self.workspace_selector.set_active(False)
        self.workspace_selector.extents = (left, right, bottom, top)
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
    """Crop/Rotate and display adjustments (redesigned semantics).

    - **Geometry** (crop / rotate / flips) never alters the original:
      on Apply the caller turns it into an adjacent virtual-copy
      thumbnail (see ``MainWindow.on_adjust_image``). The workspace
      always shows the rotated image; the crop rectangle lives in that
      rotated frame and is clamped to the largest inscribed rectangle,
      so the result never contains NaN padding (the old
      small-image-in-an-empty-frame problem is impossible by
      construction).
    - **Display** (clip / gamma / colormap) stays a live, reversible
      per-file adjustment on the original image.

    The result preview is rendered by ``resample_geometry`` — the same
    function that produces the final copy, so preview == result.
    Callers read ``geometry_spec()`` / ``tone_spec()`` / ``selected_cmap()``
    after Accept.
    """
    def __init__(self, parent, base_image, spec, cmap_name, base_extent=None, display_extent=None,
                 axis_unit=None, colorbar_label=None, base_unit=None, relative_axes=False):
        super().__init__(parent)
        self.viewer = parent
        self.setWindowTitle("Crop / Rotate & display adjustments")

        self.base_image = np.asarray(base_image)
        self.base_extent = base_extent
        self.axis_unit = axis_unit or 'px'
        self.colorbar_label = colorbar_label or ''

        self.current_spec = self._seed_from_legacy(spec, cmap_name)
        self._last_angle = float(self.current_spec.get('rotate', 0.0) or 0.0)

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

        # Geometry group — everything here lands in a virtual copy on Apply.
        geom_group = QtWidgets.QGroupBox("Geometry — creates a virtual copy")
        geom_layout = QtWidgets.QVBoxLayout(geom_group)
        geom_hint = QtWidgets.QLabel(
            "Apply adds a new [edit] thumbnail next to the source; the "
            "original image is never altered.")
        geom_hint.setWordWrap(True)
        geom_hint.setStyleSheet("font-size: 10px;")
        geom_layout.addWidget(geom_hint)

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

        crop_form = QtWidgets.QFormLayout()
        # Crop rect in rotated-frame pixels (can be negative once rotated —
        # the frame's bounding box extends past the source edges).
        self.left_spin = QtWidgets.QSpinBox(); self.left_spin.setRange(-100000, 100000)
        self.right_spin = QtWidgets.QSpinBox(); self.right_spin.setRange(-100000, 100000)
        self.bottom_spin = QtWidgets.QSpinBox(); self.bottom_spin.setRange(-100000, 100000)
        self.top_spin = QtWidgets.QSpinBox(); self.top_spin.setRange(-100000, 100000)
        crop_form.addRow("Crop left (px)", self.left_spin)
        crop_form.addRow("Crop right (px)", self.right_spin)
        crop_form.addRow("Crop bottom (px)", self.bottom_spin)
        crop_form.addRow("Crop top (px)", self.top_spin)
        geom_layout.addLayout(crop_form)

        self.lock_square_cb = QtWidgets.QCheckBox("Square crop")
        geom_layout.addWidget(self.lock_square_cb)
        controls_layout.addWidget(geom_group)

        # Display group — live, reversible adjustments on the original.
        tone_group = QtWidgets.QGroupBox("Display — live on this image")
        tone_form = QtWidgets.QFormLayout(tone_group)
        self.low_pct_spin = QtWidgets.QDoubleSpinBox(); self.low_pct_spin.setRange(0.0, 100.0); self.low_pct_spin.setDecimals(2)
        self.high_pct_spin = QtWidgets.QDoubleSpinBox(); self.high_pct_spin.setRange(0.0, 100.0); self.high_pct_spin.setDecimals(2)
        self.gamma_spin = QtWidgets.QDoubleSpinBox(); self.gamma_spin.setRange(0.05, 10.0); self.gamma_spin.setDecimals(2); self.gamma_spin.setSingleStep(0.05)
        tone_form.addRow("Clip low %", self.low_pct_spin)
        tone_form.addRow("Clip high %", self.high_pct_spin)
        tone_form.addRow("Gamma", self.gamma_spin)
        self.cmap_combo = QtWidgets.QComboBox()
        _current_cmap = str(self.current_spec.get('cmap', 'viridis') or 'viridis')
        _names = cmap_registry.featured_cmap_names("general")
        if _current_cmap not in _names:
            _names = [_current_cmap] + _names
        for name in _names:
            try:
                icon = _colormap_icon(name, width=96, height=14)
            except Exception:
                icon = QIcon()
            self.cmap_combo.addItem(icon, name)
        tone_form.addRow("Colormap", self.cmap_combo)
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

        # Options
        opt_group = QtWidgets.QGroupBox("Options")
        opt_layout = QtWidgets.QVBoxLayout(opt_group)
        self.scalebar_cb = QtWidgets.QCheckBox("Show scalebar")
        opt_layout.addWidget(self.scalebar_cb)
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
        ok_btn = btn_box.button(QtWidgets.QDialogButtonBox.Ok)
        if ok_btn is not None:
            ok_btn.setText("Apply")
        preview_layout.addWidget(btn_box)
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)

        # connections
        for spin in (self.left_spin, self.right_spin, self.bottom_spin, self.top_spin,
                     self.low_pct_spin, self.high_pct_spin, self.gamma_spin):
            spin.valueChanged.connect(self._on_params_changed_live)
            if hasattr(spin, 'editingFinished'):
                spin.editingFinished.connect(self._commit_live_change)

        self.rotate_slider.sliderPressed.connect(self._begin_live_change)
        self.rotate_slider.valueChanged.connect(self._on_params_changed_live)
        self.rotate_slider.sliderReleased.connect(self._commit_live_change)

        for w in (self.flip_h_cb, self.flip_v_cb, self.scalebar_cb, self.lock_square_cb):
            w.toggled.connect(self._on_discrete_change)
        self.cmap_combo.currentIndexChanged.connect(self._on_discrete_change)

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

    # ---------- geometry helpers ----------
    def _current_bbox(self, angle_deg):
        """Bounding box (left, right, bottom, top) of the source frame
        rotated by angle_deg about its center, in rotated-frame px."""
        h, w = self.base_image.shape[:2]
        rad = math.radians(float(angle_deg or 0.0))
        cos_a = abs(math.cos(rad))
        sin_a = abs(math.sin(rad))
        bw = w * cos_a + h * sin_a
        bh = w * sin_a + h * cos_a
        return (0.5 * w - 0.5 * bw, 0.5 * w + 0.5 * bw,
                0.5 * h - 0.5 * bh, 0.5 * h + 0.5 * bh)

    def _inscribed_rect_for(self, angle_deg):
        """Largest centered axis-aligned rect fully inside the rotated
        source, shrunk half a pixel per side so edge interpolation can
        never sample outside."""
        h, w = self.base_image.shape[:2]
        rw, rh = largest_inscribed_rect(w, h, angle_deg)
        if abs(float(angle_deg or 0.0)) > 1e-9:
            rw = max(2.0, rw - 1.0)
            rh = max(2.0, rh - 1.0)
        return (0.5 * w - 0.5 * rw, 0.5 * w + 0.5 * rw,
                0.5 * h - 0.5 * rh, 0.5 * h + 0.5 * rh)

    def _clamp_crop_rect(self, rect, angle_deg, square=False):
        il, ir, ib, it = self._inscribed_rect_for(angle_deg)
        left, right, bottom, top = [float(v) for v in rect]
        left = max(il, min(left, ir - 2.0))
        bottom = max(ib, min(bottom, it - 2.0))
        right = max(left + 2.0, min(right, ir))
        top = max(bottom + 2.0, min(top, it))
        if square:
            size = min(right - left, top - bottom)
            right = left + size
            top = bottom + size
        return (left, right, bottom, top)

    def _seed_from_legacy(self, spec, cmap_name):
        """Build the dialog's spec from a stored one. New (tone-only)
        specs and old geometry specs both land here; legacy geometry is
        re-interpreted in the rotated-frame model (a one-time,
        approximate migration — the live preview shows the exact result
        before anything is applied)."""
        src = json.loads(json.dumps(spec or {}))
        h, w = self.base_image.shape[:2]
        rotate = float(src.get('rotate', 0.0) or 0.0)
        clip = src.get('clip') or {}
        out = {
            'rotate': rotate,
            'flip_h': bool(src.get('flip_h', False)),
            'flip_v': bool(src.get('flip_v', False)),
            'clip': {'low': clip.get('low'), 'high': clip.get('high')},
            'gamma': float(src.get('gamma', 1.0) or 1.0),
            'cmap': str(src.get('cmap') or cmap_name or 'viridis'),
            'lock_square': bool(src.get('lock_square', False)),
        }
        crop_rect = None
        legacy_crop = src.get('crop') or {}
        if legacy_crop and abs(rotate) <= 1e-3:
            # Unrotated legacy crop maps exactly: stored rows are raw
            # top-origin slices, the rotated-frame rect is bottom-origin.
            try:
                x0 = float(legacy_crop.get('x0', 0))
                x1 = float(legacy_crop.get('x1', w))
                y0 = float(legacy_crop.get('y0', 0))
                y1 = float(legacy_crop.get('y1', h))
                if (x0, x1, y0, y1) != (0.0, float(w), 0.0, float(h)):
                    crop_rect = (x0, x1, h - y1, h - y0)
            except Exception:
                crop_rect = None
        if crop_rect is None:
            crop_rect = self._inscribed_rect_for(rotate)
        left, right, bottom, top = self._clamp_crop_rect(crop_rect, rotate)
        out['crop'] = {'x0': int(round(left)), 'x1': int(round(right)),
                       'y0': int(round(bottom)), 'y1': int(round(top))}
        return out

    # ---------- UI <-> spec ----------
    def _apply_spec_to_controls(self):
        self._updating_controls = True
        crop = self.current_spec.get('crop', {})
        insc = self._inscribed_rect_for(self.current_spec.get('rotate', 0.0))
        self.left_spin.setValue(int(crop.get('x0', round(insc[0]))))
        self.right_spin.setValue(int(crop.get('x1', round(insc[1]))))
        self.bottom_spin.setValue(int(crop.get('y0', round(insc[2]))))
        self.top_spin.setValue(int(crop.get('y1', round(insc[3]))))

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

        cmap = self.current_spec.get('cmap', self.cmap_combo.currentText())
        if self.cmap_combo.findText(cmap) < 0:
            try:
                icon = _colormap_icon(cmap, width=96, height=14)
            except Exception:
                icon = QIcon()
            self.cmap_combo.addItem(icon, cmap)
        self.cmap_combo.setCurrentText(cmap)

        self._last_angle = float(self.current_spec.get('rotate', 0.0) or 0.0)
        self._updating_controls = False

    def _collect_spec_from_controls(self):
        low = float(self.low_pct_spin.value())
        high = float(self.high_pct_spin.value())
        if high < low:
            high = low
            self.high_pct_spin.blockSignals(True)
            self.high_pct_spin.setValue(high)
            self.high_pct_spin.blockSignals(False)

        angle = float(self.rotate_slider.value())
        rect = (float(self.left_spin.value()), float(self.right_spin.value()),
                float(self.bottom_spin.value()), float(self.top_spin.value()))
        left, right, bottom, top = self._clamp_crop_rect(
            rect, angle, square=self.lock_square_cb.isChecked())
        crop = {'x0': int(round(left)), 'x1': int(round(right)),
                'y0': int(round(bottom)), 'y1': int(round(top))}
        for spin, val in ((self.left_spin, crop['x0']), (self.right_spin, crop['x1']),
                          (self.bottom_spin, crop['y0']), (self.top_spin, crop['y1'])):
            if int(spin.value()) != val:
                spin.blockSignals(True)
                spin.setValue(val)
                spin.blockSignals(False)

        return {
            'crop': crop,
            'rotate': angle,
            'flip_h': self.flip_h_cb.isChecked(),
            'flip_v': self.flip_v_cb.isChecked(),
            'clip': {
                'low': low if low > 0 else None,
                'high': high if high < 100 else None,
            },
            'gamma': float(self.gamma_spin.value()),
            'cmap': self.cmap_combo.currentText(),
            'lock_square': self.lock_square_cb.isChecked(),
        }

    # ---------- result API (read by MainWindow.on_adjust_image) ----------
    def geometry_spec(self):
        """Geometry for resample_geometry, or None when it is an identity
        (no copy should be created)."""
        c = self.current_spec
        crop = c.get('crop') or {}
        geom = {
            'flip_h': bool(c.get('flip_h')),
            'flip_v': bool(c.get('flip_v')),
            'rotate': float(c.get('rotate', 0.0) or 0.0),
            'crop_rect': (float(crop.get('x0', 0)), float(crop.get('x1', 0)),
                          float(crop.get('y0', 0)), float(crop.get('y1', 0))),
        }
        if geometry_is_identity(geom, self.base_image.shape):
            return None
        return geom

    def tone_spec(self):
        """Clip/gamma display adjustments, or None when identity."""
        c = self.current_spec
        clip = c.get('clip') or {}
        low = clip.get('low')
        high = clip.get('high')
        gamma = float(c.get('gamma', 1.0) or 1.0)
        if low is None and high is None and abs(gamma - 1.0) <= 1e-3:
            return None
        return {'clip': {'low': low, 'high': high}, 'gamma': gamma}

    def selected_cmap(self):
        return self.cmap_combo.currentText()

    # ---------- callbacks ----------
    def _on_params_changed_live(self, value=None):
        if self._updating_controls:
            return
        angle = float(self.rotate_slider.value())
        if abs(angle - self._last_angle) > 1e-9:
            # Rotation changed: the old rect lives in a different frame —
            # reset the crop to the largest inscribed rectangle.
            self._last_angle = angle
            insc = self._inscribed_rect_for(angle)
            self._updating_controls = True
            self.left_spin.setValue(int(round(insc[0])))
            self.right_spin.setValue(int(round(insc[1])))
            self.bottom_spin.setValue(int(round(insc[2])))
            self.top_spin.setValue(int(round(insc[3])))
            self._updating_controls = False
        self.current_spec = self._collect_spec_from_controls()
        self.rotate_value_label.setText(f"{int(round(self.rotate_slider.value()))} deg")
        self.preview_panel.set_rotation_angle(angle)
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

    def _on_crop_selection(self, left, right, bottom, top):
        if self._updating_controls:
            return
        prev = json.loads(json.dumps(self.current_spec))
        angle = float(self.rotate_slider.value())
        left, right, bottom, top = self._clamp_crop_rect(
            (left, right, bottom, top), angle,
            square=self.lock_square_cb.isChecked())
        self._updating_controls = True
        self.left_spin.setValue(int(round(left)))
        self.right_spin.setValue(int(round(right)))
        self.bottom_spin.setValue(int(round(bottom)))
        self.top_spin.setValue(int(round(top)))
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
            'rotate': 0.0,
            'flip_h': False,
            'flip_v': False,
            'clip': {'low': None, 'high': None},
            'gamma': 1.0,
            'cmap': self.current_spec.get('cmap', 'viridis'),
            'lock_square': False,
        }
        self._push_history(prev)
        self._apply_spec_to_controls()
        self._update_preview()

    # ---------- processing ----------
    def _update_preview(self):
        spec = self.current_spec
        cmap = spec.get('cmap', 'viridis') or 'viridis'
        # Honor the full-amber display override so the dialog matches the
        # app's rendering while the mode is active.
        cmap_disp = cmap_registry.effective_cmap_name(cmap)
        angle = float(spec.get('rotate', 0.0) or 0.0)
        flips = {'flip_h': bool(spec.get('flip_h')), 'flip_v': bool(spec.get('flip_v'))}
        crop = spec.get('crop') or {}
        crop_rect = (float(crop.get('x0', 0)), float(crop.get('x1', 0)),
                     float(crop.get('y0', 0)), float(crop.get('y1', 0)))

        # Workspace: the whole rotated frame (NaN outside the source is
        # shown masked), crop selector + inscribed guide on top.
        bbox = self._current_bbox(angle)
        ws_geom = dict(flips, rotate=angle, crop_rect=bbox)
        ws_arr, _ = resample_geometry(self.base_image, None, ws_geom)
        self.preview_panel.update_workspace(
            ws_arr, bbox, 'px', crop_rect, cmap_disp,
            inscribed_rect=self._inscribed_rect_for(angle))

        # Result: the exact same engine that will build the virtual copy,
        # plus the live clip/gamma on top.
        geom = dict(flips, rotate=angle, crop_rect=crop_rect)
        arr_result, extent_result = resample_geometry(self.base_image, self.base_extent, geom)
        tone = self.tone_spec()
        if tone is not None:
            arr_result, extent_result = apply_adjustment_spec(arr_result, extent_result, tone)

        self.preview_panel.update_result(arr_result, extent_result, self.axis_unit,
                                         cmap_disp, self.scalebar_cb.isChecked())
# === END: Image adjustment classes ===





