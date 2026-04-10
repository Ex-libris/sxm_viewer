"""Detail canvases and spectroscopy dialogs."""
from __future__ import annotations

import functools
import itertools
import json
import math
import time
import re

import numpy as np
from matplotlib import patches
from matplotlib.backend_bases import MouseButton
from matplotlib import colors as mcolors
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from matplotlib.ticker import FuncFormatter, AutoMinorLocator, MultipleLocator, MaxNLocator
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
from ..plot_typography import add_font_menu_action, normalize_font_family, apply_text_style, apply_qfont_style
from ..system_open import add_source_file_menu
from ..figure_layout_presets import (
    iter_figure_layout_presets,
    get_figure_layout_preset,
    apply_figure_layout,
    preset_pixel_size,
    apply_canvas_widget_preset,
    copy_figure_to_clipboard,
    save_figure_with_dialog,
)
from ..mpl_compat import InsetPosition
try:
    from scipy import signal as _scipy_signal
except Exception:  # pragma: no cover
    _scipy_signal = None
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
from ..spectroscopy import overlays as spectro_overlays

def _normalize_topo_axis(values: np.ndarray, unit_hint: str | None) -> tuple[np.ndarray, str]:
    arr = np.asarray(values, dtype=float)
    unit = (unit_hint or "").strip()
    unit_lower = unit.lower()
    if unit_lower in ("m", "meter", "meters"):
        return arr * 1e9, "nm"
    if unit_lower in ("um", "micron", "microns"):
        return arr * 1e3, "nm"
    if unit_lower in ("pm", "picometer", "picometers"):
        return arr * 1e-3, "nm"
    # Heuristic: if values are tiny, assume meters and convert to nm.
    if unit_lower in ("nm", "nanometer", "nanometers", ""):
        try:
            max_abs = float(np.nanmax(np.abs(arr))) if np.isfinite(arr).any() else 0.0
        except Exception:
            max_abs = 0.0
        if max_abs and max_abs < 1e-3:
            return arr * 1e9, "nm"
        return arr, (unit if unit else "nm")
    return arr, unit


def _topo_axis_from_spec(spec: dict | None) -> dict | None:
    if not spec:
        return None
    channels = spec.get("channels") or {}
    if not isinstance(channels, dict):
        return None
    unit_map = spec.get("unit_map") or {}
    for name, vals in channels.items():
        low = str(name).strip().lower()
        if not low:
            continue
        if (
            "topo" in low
            or "topography" in low
            or "piezo" in low
            or low in ("z_abs", "zabs", "absz", "abs_z", "z-abs")
        ):
            unit_hint = unit_map.get(name) or guess_channel_unit(name) or ""
            arr, unit = _normalize_topo_axis(np.asarray(vals, dtype=float), unit_hint)
            return {"key": "topo", "label": name or "Topo", "unit": unit, "values": arr}
    return None


def _available_channel_names(spec: dict | None) -> list[str]:
    if not spec:
        return []
    channels = spec.get("channels") or {}
    names = []
    if isinstance(channels, dict):
        names.extend(str(name) for name in channels.keys() if str(name).strip())
    if not names:
        names.extend(str(name) for name in (spec.get("available_channels") or []) if str(name).strip())
    deduped = []
    seen = set()
    for name in names:
        if name in seen:
            continue
        seen.add(name)
        deduped.append(name)
    return deduped


_Z_UNIT_FACTORS_TO_NM = {
    "": None,
    "m": 1e9,
    "meter": 1e9,
    "meters": 1e9,
    "nm": 1.0,
    "nanometer": 1.0,
    "nanometers": 1.0,
    "pm": 1e-3,
    "picometer": 1e-3,
    "picometers": 1e-3,
    "um": 1e3,
    "micrometer": 1e3,
    "micrometers": 1e3,
    "mm": 1e6,
    "a": 0.1,
    "angstrom": 0.1,
    "angstroms": 0.1,
    "å": 0.1,
}


def _z_like_name(text: str | None) -> bool:
    low = str(text or "").strip().lower()
    if not low:
        return False
    return any(
        token in low
        for token in (
            "z piezo",
            "z absolute",
            "absolute z",
            "z_abs",
            "z-abs",
            "abs z",
            "z-controller",
            "controller>z",
            "topo",
            "topography",
            "piezo",
        )
    ) or low in {"z", "z abs", "absz"}


def _scalar_to_nm(value, unit_hint: str | None = None) -> float | None:
    if value is None:
        return None
    parsed = None
    parsed_unit = ""
    if isinstance(value, (int, float, np.floating, np.integer)):
        try:
            parsed = float(value)
        except Exception:
            parsed = None
    else:
        text = str(value).strip()
        if not text:
            return None
        match = re.match(r"^([+-]?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)(.*)$", text)
        if match:
            try:
                parsed = float(match.group(1))
            except Exception:
                parsed = None
            parsed_unit = match.group(2).strip()
        else:
            try:
                parsed = float(text)
            except Exception:
                parsed = None
    if parsed is None or not np.isfinite(parsed):
        return None
    unit_key = str(parsed_unit or unit_hint or "").strip().lower().replace("µ", "u")
    unit_key = unit_key.strip("[]() ")
    factor = _Z_UNIT_FACTORS_TO_NM.get(unit_key)
    if factor is not None:
        return float(parsed) * factor
    if not unit_key:
        return float(parsed) * 1e9 if abs(float(parsed)) < 1e-3 else float(parsed)
    return None


def _constant_axis_value_nm(values, unit_hint: str | None = None, tol_nm: float = 1e-3) -> float | None:
    try:
        arr = np.asarray(values, dtype=float).ravel()
    except Exception:
        return None
    if arr.size == 0:
        return None
    finite = arr[np.isfinite(arr)]
    if finite.size == 0:
        return None
    arr_nm, _unit = _normalize_topo_axis(finite, unit_hint)
    try:
        span = float(np.nanmax(arr_nm) - np.nanmin(arr_nm))
        center = float(np.nanmedian(arr_nm))
    except Exception:
        return None
    limit = max(float(tol_nm), abs(center) * 1e-6)
    if span <= limit:
        return center
    return None


def _metadata_z_from_spec(spec: dict | None) -> tuple[float | None, str]:
    if not spec:
        return None, ""
    best: tuple[int, float, str] | None = None
    for key, value in list((spec or {}).items()):
        label = str(key or "").strip()
        if not label or not _z_like_name(label):
            continue
        label_low = label.lower()
        unit_hint = ""
        if "(m)" in label_low:
            unit_hint = "m"
        elif "(nm)" in label_low:
            unit_hint = "nm"
        elif "(pm)" in label_low:
            unit_hint = "pm"
        elif "(um)" in label_low:
            unit_hint = "um"
        level = _scalar_to_nm(value, unit_hint=unit_hint)
        if level is None:
            continue
        score = 0
        preferred_label = None
        if "z-controller" in label_low and ">z" in label_low.replace("_", ">"):
            score = 120
            preferred_label = "Z piezo absolute"
        elif "z-controller" in label_low:
            score = 110
            preferred_label = "Z piezo absolute"
        elif "absolute z" in label_low or "z absolute" in label_low or "z_abs" in label_low or "abs z" in label_low:
            score = 100
            preferred_label = "Z piezo absolute"
        elif label_low.endswith("z_(m)") or label_low.endswith("z (m)") or label_low in {"z", "z_nm"}:
            score = 80
            preferred_label = "Z"
        elif "piezo" in label_low:
            score = 70
            preferred_label = "Z piezo"
        elif "topo" in label_low or "topography" in label_low:
            score = 20
            preferred_label = "Topo"
        clean_label = preferred_label or re.sub(r"\s*\(.*?\)", "", label).replace("_", " ").strip() or "Z"
        candidate = (score, float(level), clean_label)
        if best is None or candidate[0] > best[0]:
            best = candidate
    if best is None:
        return None, ""
    return best[1], best[2]


def _style_kwargs(style_state: dict | None = None) -> dict:
    style_state = style_state or {}
    return {
        "bold": bool(style_state.get("bold", False)),
        "italic": bool(style_state.get("italic", False)),
        "underline": bool(style_state.get("underline", False)),
    }

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
        "#f96855", "#56a3a6", "#9f5f9d", "#2d5d82", "#73c2ff", "#ffaec9",
        "#000000", "#202020", "#404040", "#808080", "#c0c0c0", "#ffffff"
    ]
    def __init__(self, spec, parent=None):
        super().__init__(parent)
        self.spec = spec
        self.viewer = parent
        self.setWindowTitle(f"Spectroscopy: {Path(spec['path']).name}")
        self.resize(860, 640)
        self._toggle_buttons = []
        self._advanced_controls_visible = False
        self._dark_background = False

        root_layout = QtWidgets.QVBoxLayout(self)
        root_layout.setContentsMargins(8, 8, 8, 8)
        root_layout.setSpacing(6)

        self.fig = Figure(figsize=(6, 4))
        self.canvas = FigureCanvas(self.fig)
        self.ax = self.fig.add_subplot(111)
        self.channel_combo = QtWidgets.QComboBox()
        self.axis_combo = QtWidgets.QComboBox()
        self.fit_btn = QtWidgets.QPushButton("Fit parabola")
        self.copy_btn = QtWidgets.QPushButton("Copy channel")
        self._active_line_color = self.SCIENCE_PALETTE[0]
        self._swatch_buttons = []
        self._curve_entries = []
        self._selected_curve_index = 0
        self._drag_start_pos = None
        self._font_scale = 1.0
        self._grid_enabled = True
        self._legend_enabled = True
        self._show_markers = False
        self._show_line = True
        self._x_log = False
        self._y_log = False
        self._line_width = 1.5
        self._legend_loc = "best"
        self._legend_font = 8
        self._legend_bg = True
        self._legend_border = True
        self._figure_preset_key = "interactive"
        self._show_position_inset = True
        self._position_inset_ax = None
        self._inset_bbox = None
        self._inset_drag_cids = []
        self._inset_dragging = False
        self._inset_drag_offset = (0.0, 0.0)
        self._suppress_drag_until_release = False
        # Resolve the real viewer so thumbnail/header lookups work even when
        # this popup is spawned from a comparison dialog.
        self.viewer = None
        if isinstance(parent, QtWidgets.QWidget):
            if hasattr(parent, "viewer") and getattr(parent, "viewer", None) is not None:
                self.viewer = parent.viewer
            else:
                self.viewer = parent
        self._plot_font_family = normalize_font_family(getattr(self.viewer, "_plot_font_family", None), "sans-serif")
        self._plot_font_bold = bool(getattr(self.viewer, "_plot_font_bold", False))
        self._plot_font_italic = bool(getattr(self.viewer, "_plot_font_italic", False))
        self._plot_font_underline = bool(getattr(self.viewer, "_plot_font_underline", False))
        self._filter_cfg = {
            "gaussian": {"enabled": False, "sigma": 1.0},
            "savgol": {"enabled": False, "window": 11, "poly": 3},
            "median": {"enabled": False, "size": 3},
            "fft": {"enabled": False, "cutoff": 0.15},
            "notch": {"enabled": False, "freq": 50.0, "width": 5.0},
            "derive": {"enabled": False, "window": 11, "poly": 3},
        }
        self.setAcceptDrops(True)
        self.canvas.installEventFilter(self)
        self.canvas.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.canvas.customContextMenuRequested.connect(self._on_canvas_context_menu)
        self._splitter = QtWidgets.QSplitter(QtCore.Qt.Vertical)
        self._splitter.addWidget(self.canvas)

        info_widget = QtWidgets.QWidget()
        info_layout = QtWidgets.QVBoxLayout(info_widget)
        info_layout.setContentsMargins(0, 0, 0, 0)
        info_layout.setSpacing(6)

        meta_txt = (
            f"File: {Path(spec['path']).name}\n"
            f"Position: {spec.get('x','?')}/{spec.get('y','?')} nm\n"
            f"Time: {spec.get('time')}"
        )
        self.meta_label = QtWidgets.QLabel(meta_txt)
        self.meta_label.setWordWrap(True)
        self.meta_label.setObjectName("spectroMetaLabel")

        selector_row = QtWidgets.QHBoxLayout()
        selector_row.setContentsMargins(0, 0, 0, 0)
        selector_row.setSpacing(6)
        selector_row.addWidget(QtWidgets.QLabel("Channel:"))
        selector_row.addWidget(self.channel_combo, 1)
        selector_row.addWidget(QtWidgets.QLabel("Axis:"))
        selector_row.addWidget(self.axis_combo, 1)
        selector_row.addWidget(self.fit_btn)
        selector_row.addWidget(self.copy_btn)
        info_layout.addLayout(selector_row)

        controls_panel = QtWidgets.QWidget()
        controls_layout = QtWidgets.QVBoxLayout(controls_panel)
        controls_layout.setContentsMargins(0, 0, 0, 0)
        controls_layout.setSpacing(6)

        primary_row = QtWidgets.QHBoxLayout()
        primary_row.setContentsMargins(0, 0, 0, 0)
        primary_row.setSpacing(6)
        self.markers_toggle = self._make_toggle_button("Markers", checked=self._show_markers, tooltip="Show sampled data points")
        self.markers_toggle.toggled.connect(lambda checked: self._set_plot_option("markers", checked))
        primary_row.addWidget(self.markers_toggle)
        self.lines_toggle = self._make_toggle_button("Lines", checked=self._show_line, tooltip="Show line between spectroscopy samples")
        self.lines_toggle.toggled.connect(lambda checked: self._set_plot_option("lines", checked))
        primary_row.addWidget(self.lines_toggle)
        self.grid_toggle = self._make_toggle_button("Grid", checked=self._grid_enabled, tooltip="Toggle plot grid")
        self.grid_toggle.toggled.connect(lambda checked: self._set_plot_option("grid", checked))
        primary_row.addWidget(self.grid_toggle)
        self.dark_bg_toggle = self._make_toggle_button("Dark", checked=self._dark_background, tooltip="Toggle dark spectroscopy plot background")
        self.dark_bg_toggle.toggled.connect(lambda checked: self._set_plot_option("dark", checked))
        primary_row.addWidget(self.dark_bg_toggle)
        primary_row.addStretch(1)
        self.advanced_toggle_btn = self._make_toggle_button("Advanced ▼", checked=False, tooltip="Show/hide advanced spectroscopy controls")
        self.advanced_toggle_btn.toggled.connect(self._set_advanced_options_visible)
        primary_row.addWidget(self.advanced_toggle_btn)
        controls_layout.addLayout(primary_row)

        self._advanced_controls_widget = QtWidgets.QWidget()
        advanced_layout = QtWidgets.QVBoxLayout(self._advanced_controls_widget)
        advanced_layout.setContentsMargins(0, 0, 0, 0)
        advanced_layout.setSpacing(6)
        advanced_row = QtWidgets.QHBoxLayout()
        advanced_row.setContentsMargins(0, 0, 0, 0)
        advanced_row.setSpacing(6)
        self.legend_toggle = self._make_toggle_button("Legend", checked=self._legend_enabled, tooltip="Show or hide the legend")
        self.legend_toggle.toggled.connect(lambda checked: self._set_plot_option("legend", checked))
        advanced_row.addWidget(self.legend_toggle)
        self.inset_toggle = self._make_toggle_button("Position inset", checked=self._show_position_inset, tooltip="Show miniature of acquisition image with spectrum position")
        self.inset_toggle.toggled.connect(lambda checked: self._set_plot_option("inset", checked))
        advanced_row.addWidget(self.inset_toggle)
        self.logx_toggle = self._make_toggle_button("Log X", checked=self._x_log, tooltip="Use logarithmic X axis when data permits")
        self.logx_toggle.toggled.connect(lambda checked: self._set_axis_log("x", checked))
        advanced_row.addWidget(self.logx_toggle)
        self.logy_toggle = self._make_toggle_button("Log Y", checked=self._y_log, tooltip="Use logarithmic Y axis when data permits")
        self.logy_toggle.toggled.connect(lambda checked: self._set_axis_log("y", checked))
        advanced_row.addWidget(self.logy_toggle)
        advanced_row.addStretch(1)
        advanced_layout.addLayout(advanced_row)
        advanced_layout.addWidget(self.meta_label)
        self._palette_swatches = self._create_palette_swatch_widget()
        advanced_layout.addWidget(self._palette_swatches)
        tools_row = QtWidgets.QHBoxLayout()
        tools_row.setContentsMargins(0, 0, 0, 0)
        tools_row.setSpacing(6)
        self.trace_style_btn = QtWidgets.QToolButton(self)
        self.trace_style_btn.setText("Traces")
        self.trace_style_btn.setPopupMode(QtWidgets.QToolButton.InstantPopup)
        self.trace_style_btn.setToolTip("Change line thickness, style, and colors for individual traces")
        self.trace_style_menu = QtWidgets.QMenu(self.trace_style_btn)
        self.trace_style_menu.aboutToShow.connect(self._populate_trace_style_menu)
        self.trace_style_btn.setMenu(self.trace_style_menu)
        tools_row.addWidget(self.trace_style_btn)
        self.legend_menu_btn = QtWidgets.QToolButton(self)
        self.legend_menu_btn.setText("Legend")
        self.legend_menu_btn.setPopupMode(QtWidgets.QToolButton.InstantPopup)
        self.legend_menu_btn.setToolTip("Adjust legend position, font size, background, and border")
        self.legend_menu = QtWidgets.QMenu(self.legend_menu_btn)
        self.legend_menu.aboutToShow.connect(self._populate_legend_menu)
        self.legend_menu_btn.setMenu(self.legend_menu)
        tools_row.addWidget(self.legend_menu_btn)
        self.filters_btn = QtWidgets.QToolButton(self)
        self.filters_btn.setText("Filters")
        self.filters_btn.setPopupMode(QtWidgets.QToolButton.InstantPopup)
        self.filters_btn.setToolTip("Apply smoothing or derivative filters to the plotted traces")
        self.filters_menu = QtWidgets.QMenu(self.filters_btn)
        self.filters_menu.aboutToShow.connect(self._populate_filter_menu)
        self.filters_btn.setMenu(self.filters_menu)
        tools_row.addWidget(self.filters_btn)
        tools_row.addStretch(1)
        advanced_layout.addLayout(tools_row)
        controls_layout.addWidget(self._advanced_controls_widget)
        self._set_advanced_options_visible(False)
        info_layout.addWidget(controls_panel)

        self.fit_result_label = QtWidgets.QLabel("")
        self.fit_result_label.setWordWrap(True)
        self.fit_result_label.setVisible(False)
        info_layout.addWidget(self.fit_result_label)

        traces_header = QtWidgets.QHBoxLayout()
        traces_header.setContentsMargins(0, 0, 0, 0)
        traces_header.setSpacing(6)
        traces_header.addWidget(QtWidgets.QLabel("Traces"))
        traces_header.addStretch(1)
        info_layout.addLayout(traces_header)

        self.curve_list = QtWidgets.QListWidget()
        self.curve_list.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.curve_list.setAlternatingRowColors(True)
        self.curve_list.setUniformItemSizes(True)
        self.curve_list.currentRowChanged.connect(self._on_curve_selection_changed)
        self.curve_list.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.curve_list.customContextMenuRequested.connect(self._on_curve_list_context_menu)
        info_layout.addWidget(self.curve_list, 1)

        button_row = QtWidgets.QHBoxLayout()
        button_row.setContentsMargins(0, 0, 0, 0)
        button_row.setSpacing(6)
        self.copy_all_btn = QtWidgets.QPushButton("Copy all")
        self.copy_all_btn.clicked.connect(self._copy_all_traces_to_clipboard)
        button_row.addWidget(self.copy_all_btn)
        self.remove_btn = QtWidgets.QPushButton("Remove")
        self.remove_btn.clicked.connect(self._remove_selected_curve)
        button_row.addWidget(self.remove_btn)
        button_row.addStretch(1)
        self.close_btn = QtWidgets.QPushButton("Close")
        self.close_btn.clicked.connect(self.accept)
        button_row.addWidget(self.close_btn)
        info_layout.addLayout(button_row)

        self._splitter.addWidget(info_widget)
        self._splitter.setStretchFactor(0, 3)
        self._splitter.setStretchFactor(1, 1)
        self._splitter.setSizes([340, 190])
        root_layout.addWidget(self._splitter)
        self._apply_toggle_button_styles()

        self.axes = self._collect_axes(spec)
        self.V = np.asarray(self.axes[0]["values"], dtype=float) if self.axes else np.asarray([], dtype=float)
        self.axis_label = self.axes[0].get("label") if self.axes else "Axis"
        self.axis_unit = self.axes[0].get("unit") if self.axes else ""
        self.channels = {name: np.asarray(vals, dtype=float) for name, vals in (spec.get('channels', {}) or {}).items()}
        for ax in self.axes:
            self.axis_combo.addItem(self._axis_display_name(ax), ax.get("key"))
        self.axis_combo.currentIndexChanged.connect(self._on_axis_changed)
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
        self._sync_toggle_states()
        self._update_fit_button()
        self._refresh_action_button_states()

    def _make_toggle_button(self, text, *, checked=False, tooltip=None):
        btn = QtWidgets.QToolButton(self)
        btn.setObjectName("spectroToggleButton")
        btn.setText(text)
        btn.setCheckable(True)
        btn.setChecked(bool(checked))
        btn.setAutoRaise(False)
        btn.setCursor(QtCore.Qt.PointingHandCursor)
        btn.setToolButtonStyle(QtCore.Qt.ToolButtonTextOnly)
        btn.setSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        if tooltip:
            btn.setToolTip(tooltip)
        self._toggle_buttons.append(btn)
        return btn

    def _set_advanced_options_visible(self, visible):
        visible = bool(visible)
        self._advanced_controls_visible = visible
        if hasattr(self, "_advanced_controls_widget") and self._advanced_controls_widget is not None:
            self._advanced_controls_widget.setVisible(visible)
        if hasattr(self, "advanced_toggle_btn") and self.advanced_toggle_btn is not None:
            self.advanced_toggle_btn.blockSignals(True)
            self.advanced_toggle_btn.setChecked(visible)
            self.advanced_toggle_btn.setText("Advanced ▲" if visible else "Advanced ▼")
            self.advanced_toggle_btn.blockSignals(False)

    def _apply_toggle_button_styles(self):
        dark = bool(self._dark_background)
        if dark:
            inactive_bg = "#1e2430"
            inactive_border = "#46556e"
            inactive_text = "#d4deee"
            active_bg = "#2f6fcb"
            active_border = "#79a9f2"
            active_text = "#ffffff"
            panel_text = "#dce5f3"
            list_bg = "#10151f"
            list_alt = "#171d29"
            list_border = "#445167"
        else:
            inactive_bg = "#f3f5f9"
            inactive_border = "#aeb7c5"
            inactive_text = "#1f2a3d"
            active_bg = "#1f6fd7"
            active_border = "#5b97e8"
            active_text = "#ffffff"
            panel_text = "#314056"
            list_bg = "#ffffff"
            list_alt = "#f7f9fc"
            list_border = "#c2cad6"
        style = (
            "QToolButton#spectroToggleButton {"
            f"background-color: {inactive_bg};"
            f"color: {inactive_text};"
            f"border: 1px solid {inactive_border};"
            "border-radius: 12px;"
            "padding: 4px 12px;"
            "font-weight: 600;"
            "}"
            "QToolButton#spectroToggleButton:checked {"
            f"background-color: {active_bg};"
            f"color: {active_text};"
            f"border: 1px solid {active_border};"
            "}"
            "QToolButton#spectroToggleButton:hover {"
            f"border: 1px solid {active_border};"
            "}"
        )
        for btn in self._toggle_buttons:
            try:
                btn.setStyleSheet(style)
            except Exception:
                pass
        list_style = (
            "QListWidget {"
            f"background: {list_bg};"
            f"alternate-background-color: {list_alt};"
            f"border: 1px solid {list_border};"
            "border-radius: 6px;"
            "}"
        )
        try:
            self.curve_list.setStyleSheet(list_style)
        except Exception:
            pass
        try:
            self.meta_label.setStyleSheet(f"color: {panel_text};")
        except Exception:
            pass
        try:
            self.fit_result_label.setStyleSheet(f"color: {panel_text};")
        except Exception:
            pass

    def _entry_name(self, entry):
        return str((entry or {}).get("label") or "Trace")

    def _normalize_curve_style(self, entry):
        if not entry:
            return
        entry.setdefault("color", self._active_line_color)
        entry.setdefault("lw", None)
        entry.setdefault("ls", "-")

    def _set_curve_style(self, index=None, **changes):
        if not self._curve_entries:
            return
        idx = self._selected_curve_index if index is None else int(index)
        if idx < 0 or idx >= len(self._curve_entries):
            return
        entry = self._curve_entries[idx]
        self._normalize_curve_style(entry)
        if "color" in changes and changes["color"]:
            entry["color"] = str(changes["color"])
            if idx == self._selected_curve_index:
                self._active_line_color = entry["color"]
        if "lw" in changes and changes["lw"] is not None:
            entry["lw"] = max(0.4, min(5.0, float(changes["lw"])))
        if "ls" in changes and changes["ls"] is not None:
            entry["ls"] = str(changes["ls"])
        self._update_curve_list()
        self._plot_selected_channel()

    def _pick_curve_color(self, index=None, button=None):
        entry = self._current_entry() if index is None else (self._curve_entries[index] if 0 <= int(index) < len(self._curve_entries) else None)
        current = QtGui.QColor((entry or {}).get("color") or self._active_line_color or "#000000")
        color = QtWidgets.QColorDialog.getColor(current, self, "Select trace color")
        if not color.isValid():
            return
        hex_color = color.name()
        self._set_curve_style(index=index, color=hex_color)
        if button is not None:
            try:
                button.setStyleSheet(f"background:{hex_color};")
            except Exception:
                pass

    def _apply_global_line_style(self, lw, ls):
        for entry in self._curve_entries:
            self._normalize_curve_style(entry)
            entry["lw"] = max(0.4, min(5.0, float(lw)))
            entry["ls"] = str(ls)
        self._update_curve_list()
        self._plot_selected_channel()

    def _reset_curve_colors_to_palette(self):
        for idx, entry in enumerate(self._curve_entries):
            entry["color"] = self.SCIENCE_PALETTE[idx % len(self.SCIENCE_PALETTE)]
            if idx == self._selected_curve_index:
                self._active_line_color = entry["color"]
        self._update_curve_list()
        self._plot_selected_channel()

    def _populate_trace_style_menu(self, menu=None):
        menu = menu or getattr(self, "trace_style_menu", None)
        if menu is None:
            return
        menu.clear()
        ls_labels = {"Solid": "-", "Dashed": "--", "Dotted": ":", "Dash-dot": "-."}
        for idx, entry in enumerate(self._curve_entries):
            self._normalize_curve_style(entry)
            row = QtWidgets.QWidget()
            h = QtWidgets.QHBoxLayout(row)
            h.setContentsMargins(6, 2, 6, 2)
            h.setSpacing(6)
            name_lbl = QtWidgets.QLabel(self._entry_name(entry))
            name_lbl.setMinimumWidth(110)
            lw_spin = QtWidgets.QDoubleSpinBox()
            lw_spin.setRange(0.4, 5.0)
            lw_spin.setSingleStep(0.2)
            lw_spin.setValue(float(entry.get("lw") or self._line_width))
            ls_combo = QtWidgets.QComboBox()
            for label, value in ls_labels.items():
                ls_combo.addItem(label, value)
            ls_combo.setCurrentIndex(max(0, ls_combo.findData(entry.get("ls") or "-")))
            col_btn = QtWidgets.QPushButton()
            col_btn.setFixedWidth(36)
            color = entry.get("color") or self._active_line_color or "#000000"
            col_btn.setStyleSheet(f"background:{color};")
            col_btn.clicked.connect(lambda _=None, i=idx, btn=col_btn: self._pick_curve_color(i, btn))
            lw_spin.valueChanged.connect(lambda val, i=idx: self._set_curve_style(i, lw=float(val)))
            ls_combo.currentIndexChanged.connect(lambda _i, i=idx, cb=ls_combo: self._set_curve_style(i, ls=cb.currentData()))
            h.addWidget(name_lbl, 1)
            h.addWidget(QtWidgets.QLabel("Thick"))
            h.addWidget(lw_spin)
            h.addWidget(QtWidgets.QLabel("Style"))
            h.addWidget(ls_combo)
            h.addWidget(col_btn)
            act = QtWidgets.QWidgetAction(menu)
            act.setDefaultWidget(row)
            menu.addAction(act)
        if self._curve_entries:
            menu.addSeparator()
        global_row = QtWidgets.QWidget()
        gh = QtWidgets.QHBoxLayout(global_row)
        gh.setContentsMargins(6, 4, 6, 4)
        gh.setSpacing(6)
        gh.addWidget(QtWidgets.QLabel("All:"))
        all_lw = QtWidgets.QDoubleSpinBox()
        all_lw.setRange(0.4, 5.0)
        all_lw.setSingleStep(0.2)
        all_lw.setValue(self._line_width)
        all_ls = QtWidgets.QComboBox()
        for label, value in ls_labels.items():
            all_ls.addItem(label, value)
        all_ls.setCurrentIndex(0)
        apply_all_btn = QtWidgets.QPushButton("Apply")
        apply_all_btn.clicked.connect(lambda _=None: self._apply_global_line_style(all_lw.value(), all_ls.currentData()))
        reset_cycle_btn = QtWidgets.QPushButton("Reset colors")
        reset_cycle_btn.clicked.connect(lambda _=None: self._reset_curve_colors_to_palette())
        gh.addWidget(QtWidgets.QLabel("Thickness"))
        gh.addWidget(all_lw)
        gh.addWidget(QtWidgets.QLabel("Style"))
        gh.addWidget(all_ls)
        gh.addWidget(apply_all_btn)
        gh.addWidget(reset_cycle_btn)
        act = QtWidgets.QWidgetAction(menu)
        act.setDefaultWidget(global_row)
        menu.addAction(act)

    def _populate_legend_menu(self, menu=None):
        menu = menu or getattr(self, "legend_menu", None)
        if menu is None:
            return
        menu.clear()
        legend_show_act = QtWidgets.QAction("Show legend", menu, checkable=True, checked=self._legend_enabled)
        legend_show_act.toggled.connect(lambda checked: self._set_plot_option("legend", checked))
        menu.addAction(legend_show_act)
        pos_combo = QtWidgets.QComboBox()
        pos_combo.addItems(["Best", "Top-left", "Top-right", "Bottom-left", "Bottom-right"])
        pos_map = {
            "Best": "best",
            "Top-left": "upper left",
            "Top-right": "upper right",
            "Bottom-left": "lower left",
            "Bottom-right": "lower right",
        }
        current_label = {
            "best": "Best",
            "upper left": "Top-left",
            "upper right": "Top-right",
            "lower left": "Bottom-left",
            "lower right": "Bottom-right",
        }.get(self._legend_loc, "Best")
        pos_combo.setCurrentIndex(max(0, pos_combo.findText(current_label)))
        pos_widget = QtWidgets.QWidget()
        pos_h = QtWidgets.QHBoxLayout(pos_widget)
        pos_h.setContentsMargins(6, 2, 6, 2)
        pos_h.addWidget(QtWidgets.QLabel("Position"))
        pos_h.addWidget(pos_combo, 1)
        pos_act = QtWidgets.QWidgetAction(menu)
        pos_act.setDefaultWidget(pos_widget)
        menu.addAction(pos_act)
        font_widget = QtWidgets.QWidget()
        fw_h = QtWidgets.QHBoxLayout(font_widget)
        fw_h.setContentsMargins(6, 2, 6, 2)
        font_spin = QtWidgets.QSpinBox()
        font_spin.setRange(6, 18)
        font_spin.setValue(int(self._legend_font))
        fw_h.addWidget(QtWidgets.QLabel("Font size"))
        fw_h.addWidget(font_spin)
        font_act = QtWidgets.QWidgetAction(menu)
        font_act.setDefaultWidget(font_widget)
        menu.addAction(font_act)
        bg_act = QtWidgets.QAction("Background", menu, checkable=True, checked=self._legend_bg)
        border_act = QtWidgets.QAction("Border", menu, checkable=True, checked=self._legend_border)
        bg_act.toggled.connect(self._set_legend_bg)
        border_act.toggled.connect(self._set_legend_border)
        menu.addAction(bg_act)
        menu.addAction(border_act)
        pos_combo.currentTextChanged.connect(lambda txt: self._set_legend_position(pos_map.get(txt, "best")))
        font_spin.valueChanged.connect(self._set_legend_font)

    def _set_legend_position(self, loc):
        self._legend_loc = str(loc or "best")
        self._plot_selected_channel()

    def _set_legend_font(self, size):
        self._legend_font = float(size)
        self._plot_selected_channel()

    def _set_legend_bg(self, enabled):
        self._legend_bg = bool(enabled)
        self._plot_selected_channel()

    def _set_legend_border(self, enabled):
        self._legend_border = bool(enabled)
        self._plot_selected_channel()

    def _apply_data_filters(self, x_vals, y_vals, y_unit, x_unit):
        data = np.asarray(y_vals, dtype=float)
        x_arr = np.asarray(x_vals, dtype=float) if np.size(x_vals) == data.size else np.linspace(0, data.size - 1, data.size)
        if data.size == 0:
            return data, y_unit
        result = data.copy()
        dx = float(np.nanmean(np.diff(x_arr))) if x_arr.size > 1 else 1.0
        if not math.isfinite(dx) or dx == 0:
            dx = 1.0

        def _odd(value, minimum):
            v = max(minimum, int(value) or minimum)
            return v + 1 if v % 2 == 0 else v

        gauss = self._filter_cfg.get("gaussian", {})
        if gauss.get("enabled"):
            sigma = max(0.1, float(gauss.get("sigma", 1.0)))
            if _scipy_ndimage is not None:
                result = _scipy_ndimage.gaussian_filter1d(result, sigma=max(0.05, sigma), mode="nearest")
            else:
                radius = max(1, int(3 * sigma))
                xs = np.arange(-radius, radius + 1)
                kernel = np.exp(-(xs ** 2) / (2.0 * sigma ** 2))
                kernel /= kernel.sum() or 1.0
                result = np.convolve(result, kernel, mode="same")

        median = self._filter_cfg.get("median", {})
        if median.get("enabled"):
            size = _odd(median.get("size", 3), 3)
            if _scipy_ndimage is not None:
                result = _scipy_ndimage.median_filter(result, size=size, mode="nearest")
            else:
                pad = size // 2
                padded = np.pad(result, pad, mode="edge")
                out = np.empty_like(result)
                for i in range(result.size):
                    out[i] = np.median(padded[i:i + size])
                result = out

        sav_cfg = self._filter_cfg.get("savgol", {})
        if sav_cfg.get("enabled"):
            window = _odd(sav_cfg.get("window", 11), 5)
            window = min(window, result.size - 1 if result.size % 2 == 0 else result.size)
            window = max(5, window if window % 2 == 1 else window - 1)
            poly = max(2, min(int(sav_cfg.get("poly", 3)), window - 1))
            if window >= 3 and window <= result.size:
                if _scipy_signal is not None:
                    result = _scipy_signal.savgol_filter(result, window, poly, mode="interp")
                else:
                    result = np.convolve(result, np.ones(window) / float(window), mode="same")

        fft_cfg = self._filter_cfg.get("fft", {})
        if fft_cfg.get("enabled") and result.size >= 8:
            cutoff = min(max(float(fft_cfg.get("cutoff", 0.15)), 0.0), 0.5)
            if cutoff > 0.0:
                centered = result - np.nanmean(result)
                freq = np.fft.rfftfreq(result.size, d=dx)
                spectrum = np.fft.rfft(centered)
                nyquist = 0.5 / dx
                spectrum *= (np.abs(freq) <= cutoff * nyquist)
                result = np.fft.irfft(spectrum, n=result.size) + np.nanmean(result)

        notch = self._filter_cfg.get("notch", {})
        if notch.get("enabled") and result.size >= 8:
            freq = abs(float(notch.get("freq", 50.0)))
            width = max(0.0001, abs(float(notch.get("width", 5.0))))
            if freq > 0.0:
                centered = result - np.nanmean(result)
                spectrum = np.fft.rfft(centered)
                freqs = np.fft.rfftfreq(result.size, d=dx)
                spectrum *= ~(np.abs(freqs - freq) < width)
                result = np.fft.irfft(spectrum, n=result.size) + np.nanmean(result)

        deriv = self._filter_cfg.get("derive", {})
        unit = y_unit
        if deriv.get("enabled"):
            window = _odd(deriv.get("window", 11), 5)
            window = min(window, result.size - 1 if result.size % 2 == 0 else result.size)
            window = max(5, window if window % 2 == 1 else window - 1)
            poly = max(2, min(int(deriv.get("poly", 3)), window - 1))
            if _scipy_signal is not None and window >= 5 and window <= result.size:
                result = _scipy_signal.savgol_filter(result, window, poly, deriv=1, delta=dx, mode="interp")
            else:
                result = np.gradient(result, x_arr)
            unit = f"d({unit or 'arb'})/d({x_unit or 'x'})"
        return result, unit

    def _set_filter_enabled(self, section, enabled):
        self._filter_cfg.setdefault(section, {})["enabled"] = bool(enabled)
        self._plot_selected_channel()

    def _set_filter_value(self, section, key, value):
        self._filter_cfg.setdefault(section, {})[key] = value
        self._plot_selected_channel()

    def _reset_filters(self):
        for section in self._filter_cfg.values():
            section["enabled"] = False
        self._plot_selected_channel()

    def _populate_filter_menu(self, menu=None):
        menu = menu or getattr(self, "filters_menu", None)
        if menu is None:
            return
        menu.clear()
        cfg = self._filter_cfg

        def widget_action(widget):
            act = QtWidgets.QWidgetAction(menu)
            act.setDefaultWidget(widget)
            menu.addAction(act)

        g_row = QtWidgets.QWidget()
        g_layout = QtWidgets.QHBoxLayout(g_row)
        g_layout.setContentsMargins(6, 2, 6, 2)
        g_layout.setSpacing(6)
        g_cb = QtWidgets.QCheckBox("Gaussian σ")
        g_cb.setChecked(cfg.get("gaussian", {}).get("enabled", False))
        g_spin = QtWidgets.QDoubleSpinBox()
        g_spin.setRange(0.1, 10.0)
        g_spin.setSingleStep(0.1)
        g_spin.setValue(float(cfg.get("gaussian", {}).get("sigma", 1.0)))
        g_cb.toggled.connect(lambda chk: self._set_filter_enabled("gaussian", chk))
        g_spin.valueChanged.connect(lambda val: self._set_filter_value("gaussian", "sigma", float(val)))
        g_layout.addWidget(g_cb)
        g_layout.addWidget(g_spin)
        widget_action(g_row)

        sg_row = QtWidgets.QWidget()
        sg_layout = QtWidgets.QHBoxLayout(sg_row)
        sg_layout.setContentsMargins(6, 2, 6, 2)
        sg_layout.setSpacing(6)
        sg_cb = QtWidgets.QCheckBox("Savitzky-Golay")
        sg_cb.setChecked(cfg.get("savgol", {}).get("enabled", False))
        sg_win = QtWidgets.QSpinBox()
        sg_win.setRange(5, 201)
        sg_win.setSingleStep(2)
        sg_win.setValue(int(cfg.get("savgol", {}).get("window", 11)))
        sg_poly = QtWidgets.QSpinBox()
        sg_poly.setRange(2, 10)
        sg_poly.setValue(int(cfg.get("savgol", {}).get("poly", 3)))
        sg_cb.toggled.connect(lambda chk: self._set_filter_enabled("savgol", chk))
        sg_win.valueChanged.connect(lambda val: self._set_filter_value("savgol", "window", int(val)))
        sg_poly.valueChanged.connect(lambda val: self._set_filter_value("savgol", "poly", int(val)))
        sg_layout.addWidget(sg_cb)
        sg_layout.addWidget(QtWidgets.QLabel("Window"))
        sg_layout.addWidget(sg_win)
        sg_layout.addWidget(QtWidgets.QLabel("Poly"))
        sg_layout.addWidget(sg_poly)
        widget_action(sg_row)

        med_row = QtWidgets.QWidget()
        med_layout = QtWidgets.QHBoxLayout(med_row)
        med_layout.setContentsMargins(6, 2, 6, 2)
        med_layout.setSpacing(6)
        med_cb = QtWidgets.QCheckBox("Median")
        med_cb.setChecked(cfg.get("median", {}).get("enabled", False))
        med_spin = QtWidgets.QSpinBox()
        med_spin.setRange(3, 51)
        med_spin.setSingleStep(2)
        med_spin.setValue(int(cfg.get("median", {}).get("size", 3)))
        med_cb.toggled.connect(lambda chk: self._set_filter_enabled("median", chk))
        med_spin.valueChanged.connect(lambda val: self._set_filter_value("median", "size", int(val)))
        med_layout.addWidget(med_cb)
        med_layout.addWidget(QtWidgets.QLabel("Size"))
        med_layout.addWidget(med_spin)
        widget_action(med_row)

        fft_row = QtWidgets.QWidget()
        fft_layout = QtWidgets.QHBoxLayout(fft_row)
        fft_layout.setContentsMargins(6, 2, 6, 2)
        fft_layout.setSpacing(6)
        fft_cb = QtWidgets.QCheckBox("FFT low-pass")
        fft_cb.setChecked(cfg.get("fft", {}).get("enabled", False))
        fft_cut = QtWidgets.QDoubleSpinBox()
        fft_cut.setRange(0.01, 0.5)
        fft_cut.setSingleStep(0.01)
        fft_cut.setDecimals(3)
        fft_cut.setValue(float(cfg.get("fft", {}).get("cutoff", 0.15)))
        fft_cb.toggled.connect(lambda chk: self._set_filter_enabled("fft", chk))
        fft_cut.valueChanged.connect(lambda val: self._set_filter_value("fft", "cutoff", float(val)))
        fft_layout.addWidget(fft_cb)
        fft_layout.addWidget(QtWidgets.QLabel("Cutoff"))
        fft_layout.addWidget(fft_cut)
        widget_action(fft_row)

        notch_row = QtWidgets.QWidget()
        notch_layout = QtWidgets.QHBoxLayout(notch_row)
        notch_layout.setContentsMargins(6, 2, 6, 2)
        notch_layout.setSpacing(6)
        notch_cb = QtWidgets.QCheckBox("Notch")
        notch_cb.setChecked(cfg.get("notch", {}).get("enabled", False))
        notch_freq = QtWidgets.QDoubleSpinBox()
        notch_freq.setRange(0.1, 5000.0)
        notch_freq.setSingleStep(1.0)
        notch_freq.setDecimals(3)
        notch_freq.setValue(float(cfg.get("notch", {}).get("freq", 50.0)))
        notch_width = QtWidgets.QDoubleSpinBox()
        notch_width.setRange(0.001, 500.0)
        notch_width.setSingleStep(0.5)
        notch_width.setDecimals(3)
        notch_width.setValue(float(cfg.get("notch", {}).get("width", 5.0)))
        notch_cb.toggled.connect(lambda chk: self._set_filter_enabled("notch", chk))
        notch_freq.valueChanged.connect(lambda val: self._set_filter_value("notch", "freq", float(val)))
        notch_width.valueChanged.connect(lambda val: self._set_filter_value("notch", "width", float(val)))
        notch_layout.addWidget(notch_cb)
        notch_layout.addWidget(QtWidgets.QLabel("Freq"))
        notch_layout.addWidget(notch_freq)
        notch_layout.addWidget(QtWidgets.QLabel("Width"))
        notch_layout.addWidget(notch_width)
        widget_action(notch_row)

        deriv_row = QtWidgets.QWidget()
        deriv_layout = QtWidgets.QHBoxLayout(deriv_row)
        deriv_layout.setContentsMargins(6, 2, 6, 2)
        deriv_layout.setSpacing(6)
        deriv_cb = QtWidgets.QCheckBox("dY/dX")
        deriv_cb.setChecked(cfg.get("derive", {}).get("enabled", False))
        deriv_win = QtWidgets.QSpinBox()
        deriv_win.setRange(5, 201)
        deriv_win.setSingleStep(2)
        deriv_win.setValue(int(cfg.get("derive", {}).get("window", 11)))
        deriv_poly = QtWidgets.QSpinBox()
        deriv_poly.setRange(2, 10)
        deriv_poly.setValue(int(cfg.get("derive", {}).get("poly", 3)))
        deriv_cb.toggled.connect(lambda chk: self._set_filter_enabled("derive", chk))
        deriv_win.valueChanged.connect(lambda val: self._set_filter_value("derive", "window", int(val)))
        deriv_poly.valueChanged.connect(lambda val: self._set_filter_value("derive", "poly", int(val)))
        deriv_layout.addWidget(deriv_cb)
        deriv_layout.addWidget(QtWidgets.QLabel("Window"))
        deriv_layout.addWidget(deriv_win)
        deriv_layout.addWidget(QtWidgets.QLabel("Poly"))
        deriv_layout.addWidget(deriv_poly)
        widget_action(deriv_row)

        reset_btn = QtWidgets.QPushButton("Disable all filters")
        reset_btn.clicked.connect(lambda _=None: self._reset_filters())
        widget_action(reset_btn)

    def _set_plot_option(self, option, checked):
        checked = bool(checked)
        if option == "markers":
            self._show_markers = checked
        elif option == "lines":
            self._show_line = checked
        elif option == "grid":
            self._grid_enabled = checked
        elif option == "dark":
            self._dark_background = checked
            self._apply_toggle_button_styles()
        elif option == "legend":
            self._legend_enabled = checked
        elif option == "inset":
            self._show_position_inset = checked
        self._sync_toggle_states()
        self._plot_selected_channel()

    def _sync_toggle_states(self):
        mapping = (
            (getattr(self, "markers_toggle", None), bool(self._show_markers)),
            (getattr(self, "lines_toggle", None), bool(self._show_line)),
            (getattr(self, "grid_toggle", None), bool(self._grid_enabled)),
            (getattr(self, "dark_bg_toggle", None), bool(self._dark_background)),
            (getattr(self, "legend_toggle", None), bool(self._legend_enabled)),
            (getattr(self, "inset_toggle", None), bool(self._show_position_inset)),
            (getattr(self, "logx_toggle", None), bool(self._x_log)),
            (getattr(self, "logy_toggle", None), bool(self._y_log)),
        )
        for btn, state in mapping:
            if btn is None:
                continue
            try:
                btn.blockSignals(True)
                btn.setChecked(state)
            finally:
                try:
                    btn.blockSignals(False)
                except Exception:
                    pass

    def _refresh_action_button_states(self):
        has_entries = bool(self._curve_entries)
        try:
            self.copy_btn.setEnabled(bool(has_entries and self.channel_combo.count()))
        except Exception:
            pass
        try:
            self.copy_all_btn.setEnabled(bool(len(self._curve_entries) > 1))
        except Exception:
            pass
        try:
            self.remove_btn.setEnabled(bool(len(self._curve_entries) > 1 and self._selected_curve_index >= 0))
        except Exception:
            pass

    def _remove_selected_curve(self):
        if len(self._curve_entries) <= 1:
            return
        idx = int(max(0, min(self._selected_curve_index, len(self._curve_entries) - 1)))
        try:
            self._curve_entries.pop(idx)
        except Exception:
            return
        self._selected_curve_index = max(0, min(idx, len(self._curve_entries) - 1))
        self._update_curve_list()
        self._plot_selected_channel()
        self._refresh_action_button_states()

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

    def _channel_unit_for_spec(self, spec, channel_label):
        unit_map = (spec or {}).get('unit_map') or {}
        if channel_label and channel_label in unit_map and unit_map[channel_label]:
            return unit_map[channel_label]
        if unit_map:
            for _key, val in unit_map.items():
                if val:
                    return val
        return guess_channel_unit(channel_label)

    def _channel_unit_for_channel(self, name):
        """Return the best-known unit string for a given spectroscopy channel."""
        if not name:
            return ""
        spec = None
        for entry in self._curve_entries or []:
            if entry.get("channel") == name and entry.get("spec"):
                spec = entry.get("spec")
                break
        if spec is None:
            spec = getattr(self, "spec", None)
        unit = ""
        if spec:
            unit = self._channel_unit_for_spec(spec, name) or ""
        if not unit:
            unit = (self.spec.get('unit_map', {}) or {}).get(name, "") if getattr(self, "spec", None) else ""
        if not unit:
            unit = guess_channel_unit(name) or ""
        if not unit:
            match = re.search(r"\(([^)]+)\)", str(name))
            if match:
                unit = match.group(1).strip()
        return unit

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
        extra_topo = _topo_axis_from_spec(spec)
        if extra_topo and not any((a.get("key") == "topo") or ("topo" in str(a.get("label", "")).lower()) for a in axes):
            axes.append(extra_topo)
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

    def _axis_values_for_spec(self, spec, axis_key):
        """Return axis values/label/unit for a given spec and axis key."""
        if spec is None:
            return np.asarray([]), "Axis", ""
        axis_key = axis_key or "primary"
        choices = spec.get("AxisChoices") or []
        for choice in choices:
            key = choice.get("key") or choice.get("label")
            if key == axis_key:
                vals = np.asarray(choice.get("values", []), dtype=float)
                return vals, choice.get("label") or "Axis", choice.get("unit") or ""
        if axis_key == "topo":
            extra_topo = _topo_axis_from_spec(spec)
            if extra_topo is not None:
                return (
                    np.asarray(extra_topo.get("values", []), dtype=float),
                    extra_topo.get("label") or "Topo",
                    extra_topo.get("unit") or "nm",
                )
        if axis_key == "alt":
            alt_vals = spec.get("AltAxis")
            if alt_vals is not None:
                vals = np.asarray(alt_vals, dtype=float)
                return vals, spec.get("AltAxisLabel") or "Z rel", spec.get("AltAxisUnit") or "nm"
        vals = np.asarray(spec.get("V", []), dtype=float)
        label = spec.get("AxisLabel") or "Axis"
        unit = spec.get("AxisUnit") or ""
        return vals, label, unit

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
        self.fit_result_label.setVisible(False)
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
        self.fit_result_label.setVisible(False)
        self._apply_channel_to_entries(name or "")
        self._plot_selected_channel()
        self._update_fit_button()

    def _apply_channel_to_entries(self, channel):
        """Force every curve entry to use the specified channel if available."""
        changed = False
        missing = []
        target = (channel or "").strip()
        for entry in self._curve_entries or []:
            spec = entry.get("spec")
            if spec is None:
                spec = self._resolve_spec_from_viewer(entry)
                if spec:
                    entry["spec"] = spec
            if not spec:
                continue
            channels = spec.get("channels") or {}
            if target and target in channels:
                try:
                    values = np.asarray(channels[target], dtype=float)
                except Exception:
                    continue
                entry["values"] = values
                entry["channel"] = target
                entry["label"] = f"{Path(spec.get('path','')).name} ({target})"
                changed = True
            else:
                if target:
                    missing.append(Path(spec.get('path','')).name)
        if changed:
            self._update_curve_list()
        if missing:
            QtWidgets.QToolTip.showText(
                QtGui.QCursor.pos(),
                f"Channel '{target}' missing for: {', '.join(missing[:4])}"
                + ("…" if len(missing) > 4 else ""),
                self,
            )
        return changed

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
            self._apply_plot_theme()
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
        active_marker = 'o' if self._show_markers else None
        if not self._show_line and active_marker is None:
            active_marker = 'o'
        filtered_units = []
        for entry in self._curve_entries:
            self._normalize_curve_style(entry)
            axis_vals = np.asarray(entry.get("axis_vals", []), dtype=float)
            values = np.asarray(entry.get("values", []), dtype=float)
            if axis_vals.size == 0 or values.size == 0:
                continue
            scaled_axis = axis_vals * axis_plot_scale
            spec = entry.get("spec")
            y_unit = self._channel_unit_for_spec(spec, entry.get("channel")) if spec else self._channel_unit_for_channel(entry.get("channel"))
            values_plot, y_unit = self._apply_data_filters(scaled_axis, values, y_unit, axis_plot_unit or axis_unit)
            filtered_units.append(y_unit)
            self.ax.plot(
                scaled_axis,
                values_plot,
                color=entry.get("color", '#c94cfa'),
                lw=float(entry.get("lw") or self._line_width) if self._show_line else 0.0,
                linestyle=(entry.get("ls") or '-') if self._show_line else 'None',
                marker=active_marker,
                markersize=4 if active_marker else None,
                label=entry.get("label", "Data"),
            )
            plotted = True
        self._axis_plot_scale = axis_plot_scale
        self._axis_plot_unit = axis_plot_unit
        self._apply_axis_scaling()
        self.ax.set_xlabel(axis_label)
        name = self.channel_combo.currentText()
        ylabel = self._channel_label_with_unit(name)
        unit = next((u for u in filtered_units if u), "")
        if unit:
            ylabel = f"{name or 'Signal'} ({unit})"
        self.ax.set_ylabel(ylabel)
        if self._grid_enabled:
            self.ax.grid(True, alpha=0.25)
        else:
            self.ax.grid(False)
        if plotted and self._legend_enabled:
            legend = self.ax.legend(loc=self._legend_loc or 'best', fontsize=self._legend_font)
            if legend:
                legend.set_draggable(True)
        if self._last_fit_result and self._last_fit_result.get('channel') == name:
            self._draw_fit_overlay(self._last_fit_result)
        self._update_position_inset()
        self._apply_plot_theme()
        self._apply_font_scale()
        self.canvas.draw_idle()

    def _on_canvas_context_menu(self, pos):
        menu = QtWidgets.QMenu(self)
        preset_menu = menu.addMenu("Figure preset")
        preset_actions = {}
        for preset in iter_figure_layout_presets():
            act = preset_menu.addAction(preset.label)
            act.setCheckable(True)
            act.setChecked(self._figure_preset_key == preset.key)
            preset_actions[act] = preset.key
        menu.addSeparator()
        copy_data_act = menu.addAction("Copy channel data")
        copy_all_act = menu.addAction("Copy all traces (table)")
        copy_png_act = menu.addAction("Copy plot as PNG (300 dpi)")
        copy_png_600_act = menu.addAction("Copy plot as PNG (600 dpi)")
        copy_svg_act = menu.addAction("Copy plot as SVG")
        save_menu = menu.addMenu("Save plot")
        save_png_300_act = save_menu.addAction("PNG 300 dpi...")
        save_png_600_act = save_menu.addAction("PNG 600 dpi...")
        save_svg_act = save_menu.addAction("SVG (vector)...")
        save_pdf_act = save_menu.addAction("PDF (vector)...")
        add_source_file_menu(menu, self.spec.get("path"), self)
        add_font_menu_action(
            menu,
            self,
            self._plot_font_family,
            self.set_plot_font_family,
            current_style=self._font_style_state(),
            apply_style_callback=self.set_plot_typography,
        )
        traces_menu = menu.addMenu("Traces")
        self._populate_trace_style_menu(traces_menu)
        filters_menu = menu.addMenu("Filters")
        self._populate_filter_menu(filters_menu)
        legend_menu = menu.addMenu("Legend")
        self._populate_legend_menu(legend_menu)
        style_menu = menu.addMenu("Plot style")
        grid_act = style_menu.addAction("Show grid")
        grid_act.setCheckable(True)
        grid_act.setChecked(self._grid_enabled)
        legend_act = style_menu.addAction("Show legend")
        legend_act.setCheckable(True)
        legend_act.setChecked(self._legend_enabled)
        marker_act = style_menu.addAction("Show markers")
        marker_act.setCheckable(True)
        marker_act.setChecked(self._show_markers)
        lines_act = style_menu.addAction("Show lines")
        lines_act.setCheckable(True)
        lines_act.setChecked(self._show_line)
        style_menu.addSeparator()
        xlog_act = style_menu.addAction("Log X axis")
        xlog_act.setCheckable(True)
        xlog_act.setChecked(self._x_log)
        ylog_act = style_menu.addAction("Log Y axis")
        ylog_act.setCheckable(True)
        ylog_act.setChecked(self._y_log)
        style_menu.addSeparator()
        width_menu = style_menu.addMenu("Line width")
        width_actions = []
        width_presets = [
            ("Ultra thin (0.6 px)", 0.6),
            ("Thin (1.0 px)", 1.0),
            ("Medium (1.6 px)", 1.6),
            ("Bold (2.4 px)", 2.4),
            ("Heavy (3.5 px)", 3.5),
        ]
        for label, value in width_presets:
            act = width_menu.addAction(label)
            act.setCheckable(True)
            act.setChecked(abs(self._line_width - value) < 0.25)
            act.setData(("line_width", value))
            width_actions.append(act)
        width_menu.addSeparator()
        width_inc_act = width_menu.addAction("Increase")
        width_dec_act = width_menu.addAction("Decrease")
        style_menu.addSeparator()
        position_act = style_menu.addAction("Show position inset")
        position_act.setCheckable(True)
        position_act.setChecked(self._show_position_inset)
        reset_act = style_menu.addAction("Reset style")
        action = menu.exec_(self.canvas.mapToGlobal(pos))
        if action in preset_actions:
            self._apply_figure_preset(preset_actions[action])
        elif action == copy_data_act:
            self._copy_channel_to_clipboard()
        elif action == copy_all_act:
            self._copy_all_traces_to_clipboard()
        elif action == copy_png_act:
            self._copy_plot_as_png(dpi=300)
        elif action == copy_png_600_act:
            self._copy_plot_as_png(dpi=600)
        elif action == copy_svg_act:
            self._copy_plot_as_svg()
        elif action == save_png_300_act:
            self._save_plot_export("png", dpi=300)
        elif action == save_png_600_act:
            self._save_plot_export("png", dpi=600)
        elif action == save_svg_act:
            self._save_plot_export("svg")
        elif action == save_pdf_act:
            self._save_plot_export("pdf")
        elif action == grid_act:
            self._grid_enabled = grid_act.isChecked()
            self._plot_selected_channel()
        elif action == legend_act:
            self._legend_enabled = legend_act.isChecked()
            self._plot_selected_channel()
        elif action == marker_act:
            self._show_markers = marker_act.isChecked()
            self._plot_selected_channel()
        elif action == lines_act:
            self._show_line = lines_act.isChecked()
            self._plot_selected_channel()
        elif action == xlog_act:
            self._set_axis_log("x", xlog_act.isChecked())
        elif action == ylog_act:
            self._set_axis_log("y", ylog_act.isChecked())
        elif action in width_actions:
            data = action.data()
            if isinstance(data, tuple):
                self._line_width = float(data[1])
                entry = self._current_entry()
                if entry is not None:
                    entry["lw"] = self._line_width
                if self._line_width <= 0 and not self._show_markers:
                    self._show_markers = True
            self._plot_selected_channel()
        elif action == width_inc_act:
            self._adjust_line_width(+0.4)
        elif action == width_dec_act:
            self._adjust_line_width(-0.4)
        elif action == position_act:
            self._show_position_inset = position_act.isChecked()
            self._plot_selected_channel()
        elif action == reset_act:
            self._reset_plot_style()

    def _on_curve_list_context_menu(self, pos):
        row = self.curve_list.indexAt(pos).row()
        if row >= 0:
            self.curve_list.setCurrentRow(row)
        menu = QtWidgets.QMenu(self)
        self._populate_trace_style_menu(menu)
        if len(self._curve_entries) > 1:
            menu.addSeparator()
            remove_act = menu.addAction("Remove selected trace")
        else:
            remove_act = None
        chosen = menu.exec_(self.curve_list.mapToGlobal(pos))
        if remove_act is not None and chosen == remove_act:
            self._remove_selected_curve()

    def _font_style_state(self):
        return {
            "bold": bool(getattr(self, "_plot_font_bold", False)),
            "italic": bool(getattr(self, "_plot_font_italic", False)),
            "underline": bool(getattr(self, "_plot_font_underline", False)),
        }

    def set_plot_typography(self, **changes):
        """Update the plot typography state and redraw the active channel."""
        family = changes.get("family", None)
        viewer = getattr(self, "viewer", None)
        style_changes = {
            "bold": changes.get("bold", None),
            "italic": changes.get("italic", None),
            "underline": changes.get("underline", None),
        }
        if family is not None:
            family = normalize_font_family(family, "sans-serif")
            self._plot_font_family = family
        if viewer is not None and hasattr(viewer, "set_plot_typography"):
            target = {
                "family": family if family is not None else self._plot_font_family,
                "bold": bool(style_changes["bold"] if style_changes["bold"] is not None else self._plot_font_bold),
                "italic": bool(style_changes["italic"] if style_changes["italic"] is not None else self._plot_font_italic),
                "underline": bool(style_changes["underline"] if style_changes["underline"] is not None else self._plot_font_underline),
            }
            if any(getattr(viewer, f"_plot_font_{k}", None) != v for k, v in target.items()):
                try:
                    viewer.set_plot_typography(**target)
                    return
                except Exception:
                    pass
        for key, attr in (("bold", "_plot_font_bold"), ("italic", "_plot_font_italic"), ("underline", "_plot_font_underline")):
            if style_changes[key] is not None:
                setattr(self, attr, bool(style_changes[key]))
        self._plot_selected_channel()

    def set_plot_font_family(self, family: str):
        """Refresh the spectroscopy plot with a new shared font family."""
        self.set_plot_typography(family=family)

    def _apply_figure_preset(self, preset_key):
        """Apply a shared journal/slide layout preset to this spectroscopy plot."""
        preset = get_figure_layout_preset(preset_key)
        self._figure_preset_key = preset.key
        apply_figure_layout(self.fig, preset)
        plot_w_px, plot_h_px = preset_pixel_size(self, preset, max_fraction=0.62)
        apply_canvas_widget_preset(self.canvas, preset, plot_w_px, plot_h_px)
        self._plot_font_family = normalize_font_family(preset.font_family, "sans-serif")
        self._plot_font_bold = False
        self._plot_font_italic = False
        self._plot_font_underline = False
        self._font_scale = float(preset.font_scale)
        self._line_width = float(preset.line_width)
        try:
            total_w = max(760, int(plot_w_px + 120))
            total_h = max(560, int(plot_h_px + 240))
            self.resize(total_w, total_h)
        except Exception:
            pass
        self._plot_selected_channel()
        QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), f"Applied preset: {preset.label}", self)

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

    def _copy_all_traces_to_clipboard(self):
        if not self._curve_entries:
            QtWidgets.QMessageBox.information(self, "Copy spectroscopy", "No traces to copy.")
            return
        rows = []
        axis_scale = getattr(self, "_axis_plot_scale", 1.0) or 1.0
        axis_unit = getattr(self, "_axis_plot_unit", self.axis_unit) or ""
        traces = []
        for idx, entry in enumerate(self._curve_entries):
            axis_vals = np.asarray(entry.get("axis_vals", []), dtype=float)
            values = np.asarray(entry.get("values", []), dtype=float)
            if axis_vals.size == 0 or values.size == 0:
                continue
            traces.append({
                "label": entry.get("label") or f"Trace {idx+1}",
                "path": str(Path(entry.get("spec_path") or self.spec.get("path", ""))),
                "matrix_index": entry.get("matrix_index"),
                "channel": entry.get("channel") or "",
                "x_unit": axis_unit,
                "y_unit": self._channel_unit_for_channel(entry.get("channel")),
                "x_vals": axis_vals * axis_scale,
                "y_vals": values,
                "time": self.spec.get("time"),
                "pos_x": self.spec.get("x"),
                "pos_y": self.spec.get("y"),
            })
        if not traces:
            QtWidgets.QMessageBox.information(self, "Copy spectroscopy", "No traces to copy.")
            return
        name_row = []
        pos_row = []
        unit_row = []
        max_len = 0
        for trace in traces:
            label = trace.get("label") or "trace"
            acq_raw = trace.get("time")
            acq = "" if acq_raw is None else str(acq_raw)
            name_row += [label, acq]
            px = trace.get("pos_x")
            py = trace.get("pos_y")
            pos_row += [
                "" if px is None else f"{float(px):.4g} nm",
                "" if py is None else f"{float(py):.4g} nm",
            ]
            unit_row += [trace.get("x_unit") or "", trace.get("y_unit") or ""]
            x_vals = trace.get("x_vals") if trace.get("x_vals") is not None else []
            y_vals = trace.get("y_vals") if trace.get("y_vals") is not None else []
            max_len = max(max_len, len(x_vals), len(y_vals))
            trace["_x_arr"] = x_vals
            trace["_y_arr"] = y_vals
        rows.append("\t".join(name_row))
        rows.append("\t".join(pos_row))
        rows.append("\t".join(unit_row))
        for i in range(max_len):
            line_parts = []
            for trace in traces:
                x_vals = trace.get("_x_arr", [])
                y_vals = trace.get("_y_arr", [])
                x_val = "" if i >= len(x_vals) else x_vals[i]
                y_val = "" if i >= len(y_vals) else y_vals[i]
                try:
                    line_parts.append("" if x_val == "" else f"{float(x_val):.9g}")
                except Exception:
                    line_parts.append(str(x_val))
                try:
                    line_parts.append("" if y_val == "" else f"{float(y_val):.9g}")
                except Exception:
                    line_parts.append(str(y_val))
            rows.append("\t".join(line_parts))
        QtWidgets.QApplication.clipboard().setText("\n".join(rows))
        QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), "Copied all traces", self)

    def _copy_plot_as_png(self, *, dpi=300):
        try:
            copy_figure_to_clipboard(self, self.fig, "png", dpi=dpi)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Copy plot", f"Unable to copy PNG: {exc}")

    def _copy_plot_as_svg(self):
        try:
            copy_figure_to_clipboard(self, self.fig, "svg")
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Copy plot", f"Unable to copy SVG: {exc}")

    def _save_plot_export(self, fmt, *, dpi=300):
        save_figure_with_dialog(self, self.fig, default_stem="spectroscopy_plot", fmt=fmt, dpi=dpi)

    def _attach_spec_metadata(self, entry):
        spec = entry.get("spec")
        if spec is None:
            spec = self._resolve_spec_from_viewer(entry)
            if spec:
                entry["spec"] = spec
        if spec:
            entry.setdefault("matrix_index", spec.get("matrix_index"))
            entry.setdefault("image_key", str(spec.get("image_key") or ""))
            entry.setdefault("nm_coords", (spec.get("x"), spec.get("y")))
            coords = self._spec_thumbnail_coords(spec=spec, file_key=entry.get("image_key"))
            if coords:
                entry["coords"] = coords
        return entry

    def _resolve_spec_from_viewer(self, entry):
        viewer = getattr(self, "viewer", None)
        if not viewer:
            return None
        path = entry.get("spec_path")
        if not path:
            return None
        matrix_index = entry.get("matrix_index")
        image_key = entry.get("image_key")
        candidates = [spec for spec in getattr(viewer, "spectros", []) if str(spec.get("path")) == str(path)]
        if image_key:
            candidates = [spec for spec in candidates if str(spec.get("image_key")) == str(image_key)]
        if matrix_index is not None:
            for spec in candidates:
                if spec.get("matrix_index") == matrix_index:
                    return spec
        return candidates[0] if candidates else None

    def _initialize_curve_entries(self):
        channel = self.channel_combo.currentText()
        axis_vals = np.asarray(self.axes[0]["values"], dtype=float) if self.axes else np.asarray([], dtype=float)
        values = np.asarray(self.channels.get(channel), dtype=float) if channel else np.asarray([], dtype=float)
        entry = {
            "label": f"{Path(self.spec.get('path','')).name} ({channel})" if channel else Path(self.spec.get('path','')).name,
            "axis_vals": axis_vals,
            "values": values,
            "color": self._active_line_color,
            "lw": self._line_width,
            "ls": "-",
            "spec_path": str(Path(self.spec.get('path',''))),
            "channel": channel,
            "axis_label": self.axis_label,
            "axis_unit": self.axis_unit,
            "spec": self.spec,
            "matrix_index": self.spec.get("matrix_index"),
            "image_key": str(self.spec.get("image_key") or ""),
            "coords": None,
            "nm_coords": (self.spec.get("x"), self.spec.get("y")),
        }
        self._attach_spec_metadata(entry)
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
            item = QtWidgets.QListWidgetItem(entry.get("label", ""))
            color = entry.get("color")
            if color:
                try:
                    pix = QtGui.QPixmap(12, 12)
                    pix.fill(QtGui.QColor(color))
                    item.setIcon(QtGui.QIcon(pix))
                except Exception:
                    pass
            self.curve_list.addItem(item)
        self.curve_list.blockSignals(False)
        if self.curve_list.count():
            self.curve_list.setCurrentRow(current)
        self._selected_curve_index = current
        self._refresh_action_button_states()

    def _on_curve_selection_changed(self, row):
        if row < 0:
            return
        self._selected_curve_index = row
        entry = self._current_entry()
        if entry:
            self._active_line_color = entry.get("color") or self._active_line_color
            try:
                self._line_width = float(entry.get("lw") or self._line_width)
            except Exception:
                pass
            matched = None
            for btn in self._swatch_buttons:
                base = str(btn.property("baseStyle") or "")
                if self._active_line_color and self._active_line_color.lower() in base.lower():
                    matched = btn
                    break
            self._set_active_swatch(matched)
        self._refresh_action_button_states()

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
        style = _style_kwargs(self._font_style_state())
        apply_text_style(self.ax.xaxis.label, family=self._plot_font_family, **style)
        apply_text_style(self.ax.yaxis.label, family=self._plot_font_family, **style)
        for text in list(self.ax.get_xticklabels()) + list(self.ax.get_yticklabels()):
            apply_text_style(text, family=self._plot_font_family, **style)
        legend = self.ax.get_legend()
        if legend:
            for text in legend.get_texts():
                text.set_fontsize(8 * scale)
                apply_text_style(text, family=self._plot_font_family, **style)
            try:
                for text in legend.get_title().texts if hasattr(legend.get_title(), "texts") else []:
                    apply_text_style(text, family=self._plot_font_family, **style)
            except Exception:
                pass
        for widget, base in ((self.meta_label, 9.0), (self.fit_result_label, 8.5)):
            font = widget.font()
            font.setPointSizeF(base * scale)
            font = apply_qfont_style(font, family=self._plot_font_family, **style)
            widget.setFont(font)
        for widget, base in (
            (self.channel_combo, 9.0),
            (self.fit_btn, 9.0),
            (self.copy_btn, 9.0),
            (self.axis_combo, 9.0),
            (self.copy_all_btn, 9.0),
            (self.remove_btn, 9.0),
            (self.close_btn, 9.0),
            (self.curve_list, 9.0),
        ):
            if widget is None:
                continue
            try:
                font = widget.font()
                font.setPointSizeF(base * scale)
                font = apply_qfont_style(font, family=self._plot_font_family, **style)
                widget.setFont(font)
            except Exception:
                pass

    def _apply_plot_theme(self):
        dark = bool(self._dark_background)
        fig_face = "#0f1720" if dark else "#ffffff"
        ax_face = "#111827" if dark else "#ffffff"
        text = "#e5edf8" if dark else "#1f2937"
        grid = "#41526a" if dark else "#d7deea"
        legend_face = "#0f1720" if dark else "#ffffff"
        legend_edge = "#6f86a5" if dark else "#7d8ea8"
        self.fig.patch.set_facecolor(fig_face)
        self.ax.set_facecolor(ax_face)
        try:
            self.ax.tick_params(axis='both', colors=text)
        except Exception:
            pass
        for axis in (self.ax.xaxis, self.ax.yaxis):
            try:
                axis.label.set_color(text)
            except Exception:
                pass
        for spine in self.ax.spines.values():
            try:
                spine.set_color(text)
            except Exception:
                pass
        if self._grid_enabled:
            try:
                self.ax.grid(True, color=grid, alpha=0.35 if dark else 0.45)
            except Exception:
                pass
        legend = self.ax.get_legend()
        if legend:
            try:
                frame = legend.get_frame()
                frame.set_facecolor(legend_face if self._legend_bg else (0, 0, 0, 0))
                frame.set_edgecolor(legend_edge if self._legend_border else (0, 0, 0, 0))
                frame.set_alpha((0.88 if dark else 0.95) if self._legend_bg else 0.0)
                frame.set_linewidth(0.8 if self._legend_border else 0.0)
            except Exception:
                pass
            for text_artist in legend.get_texts():
                try:
                    text_artist.set_color(text)
                except Exception:
                    pass
        if self._position_inset_ax is not None:
            try:
                self._position_inset_ax.set_facecolor(ax_face)
                self._position_inset_ax.title.set_color(text)
                for spine in self._position_inset_ax.spines.values():
                    spine.set_color(text)
            except Exception:
                pass

    def _entries_support_log_axis(self, axis: str) -> bool:
        sequences = []
        key = "axis_vals" if axis == "x" else "values"
        for entry in self._curve_entries:
            raw = entry.get(key)
            if raw is None:
                continue
            data = np.asarray(raw, dtype=float)
            if data.size:
                sequences.append(data)
        if not sequences:
            return False
        return all(np.all(seq > 0) for seq in sequences)

    def _set_axis_log(self, axis: str, checked: bool):
        if checked:
            if not self._entries_support_log_axis(axis):
                self._show_plot_warning(
                    "Cannot enable log {} axis: data contains non-positive values.".format(axis.upper())
                )
                checked = False
        if axis == "x":
            self._x_log = checked
        else:
            self._y_log = checked
        self._sync_toggle_states()
        self._plot_selected_channel()

    def _adjust_line_width(self, delta: float):
        self._line_width = max(0.4, min(5.0, self._line_width + delta))
        entry = self._current_entry()
        if entry is not None:
            entry["lw"] = self._line_width
        self._plot_selected_channel()

    def _reset_plot_style(self):
        self._grid_enabled = True
        self._legend_enabled = True
        self._show_markers = False
        self._show_line = True
        self._x_log = False
        self._y_log = False
        self._line_width = 1.5
        self._legend_loc = "best"
        self._legend_font = 8
        self._legend_bg = True
        self._legend_border = True
        self._show_position_inset = True
        for section in self._filter_cfg.values():
            section["enabled"] = False
        for idx, entry in enumerate(self._curve_entries):
            entry["lw"] = self._line_width
            entry["ls"] = "-"
            entry["color"] = self.SCIENCE_PALETTE[idx % len(self.SCIENCE_PALETTE)]
            if idx == self._selected_curve_index:
                self._active_line_color = entry["color"]
        self._sync_toggle_states()
        self._update_curve_list()
        self._plot_selected_channel()

    def _show_plot_warning(self, message: str):
        center = self.canvas.mapToGlobal(self.canvas.rect().center())
        QtWidgets.QToolTip.showText(center, message, self.canvas)

    def _apply_axis_scaling(self):
        if self._x_log and not self._entries_support_log_axis("x"):
            self._x_log = False
        if self._y_log and not self._entries_support_log_axis("y"):
            self._y_log = False
        self.ax.set_xscale("log" if self._x_log else "linear")
        self.ax.set_yscale("log" if self._y_log else "linear")

    def _load_thumbnail_array(self, file_key=None):
        viewer = getattr(self, "viewer", None)
        file_key = file_key or str(self.spec.get("image_key") or "")
        if not viewer or not file_key:
            return None
        thumb = None
        label = getattr(viewer, "_thumb_labels", {}).get(file_key) if hasattr(viewer, "_thumb_labels") else None
        if label is not None and label.pixmap():
            thumb = label.pixmap()
        if thumb is None:
            try:
                width = int(getattr(viewer, "thumb_size_px", 160))
                height = max(48, int(round(width * 0.75)))
                cmap = viewer.thumb_cmap_combo.currentText() if hasattr(viewer, "thumb_cmap_combo") else None
                cmap = cmap or getattr(viewer, "thumb_cmap", "viridis")
                channel_idx = viewer.channel_dropdown.currentIndex() if hasattr(viewer, "channel_dropdown") else 0
                thumb = viewer._thumbnail_pixmap_for_file(file_key, channel_idx, width, height, cmap)
            except Exception:
                return None
        if thumb is None:
            return None
        qimg = thumb.toImage().convertToFormat(QtGui.QImage.Format_RGBA8888)
        ptr = qimg.bits()
        ptr.setsize(qimg.byteCount())
        arr = np.frombuffer(ptr, dtype=np.uint8).reshape((qimg.height(), qimg.width(), 4))
        arr = arr[..., :3] / 255.0
        gray = np.clip(arr[..., 0] * 0.299 + arr[..., 1] * 0.587 + arr[..., 2] * 0.114, 0.0, 1.0)
        tinted = np.stack([gray, gray, gray], axis=-1)
        return tinted

    def _spec_thumbnail_coords(self, spec=None, file_key=None, dims=None):
        viewer = getattr(self, "viewer", None)
        spec = spec or self.spec
        file_key = file_key or str(spec.get("image_key") or "")
        if not viewer or not file_key or spec is None:
            return None
        header, _ = viewer.headers.get(file_key, (None, None))
        if header is None:
            return None
        if dims and len(dims) == 2:
            width = max(2, int(dims[0]))
            height = max(2, int(dims[1]))
        else:
            width = int(getattr(viewer, "thumb_size_px", 160))
            height = max(48, int(round(width * 0.75)))
        try:
            coords = viewer._map_spec_to_pixels(spec, header, width, height, file_key=file_key)
        except Exception:
            coords = None
        return coords

    def _update_position_inset(self):
        if self._position_inset_ax is not None:
            try:
                self._position_inset_ax.remove()
            except Exception:
                pass
            self._position_inset_ax = None
        if not self._show_position_inset:
            self._remove_inset_drag_handlers()
            return
        base_entry = self._curve_entries[0] if self._curve_entries else None
        base_key = base_entry.get("image_key") if base_entry else str(self.spec.get("image_key") or "")
        image = self._load_thumbnail_array(base_key)
        image_dims = None
        if image is not None:
            try:
                image_dims = (int(image.shape[1]), int(image.shape[0]))
            except Exception:
                image_dims = None
        markers = self._collect_inset_markers(base_key, image_dims=image_dims)
        if image is None or not markers:
            self._remove_inset_drag_handlers()
            return
        if self._inset_bbox is None:
            self._inset_bbox = [0.04, 0.04, 0.32, 0.32]
        self._position_inset_ax = inset_axes(self.ax, width="28%", height="28%", loc="lower left", borderpad=0.8)
        self._position_inset_ax.set_axes_locator(InsetPosition(self.ax, self._inset_bbox))
        self._position_inset_ax.imshow(image, origin="upper")
        self._position_inset_ax.set_xticks([])
        self._position_inset_ax.set_yticks([])
        self._position_inset_ax.set_title("Position", fontsize=7.5 * self._font_scale)
        for color, coords in markers:
            try:
                self._position_inset_ax.scatter(
                    coords[0],
                    coords[1],
                    s=52,
                    facecolors="none",
                    edgecolors=color,
                    linewidths=1.7,
                )
            except Exception:
                continue
        self._install_inset_drag_handlers()

    def _install_inset_drag_handlers(self):
        """Install matplotlib callbacks so the inset can be dragged."""
        self._remove_inset_drag_handlers()
        if not self.canvas:
            return

        def on_press(event):
            if event.button != MouseButton.LEFT:
                return
            if self._position_inset_ax is None or not self._show_position_inset:
                return
            bbox = self._position_inset_ax.bbox
            if bbox is None:
                return
            if bbox.contains(event.x, event.y):
                self._inset_dragging = True
                self._inset_drag_offset = (event.x - bbox.x0, event.y - bbox.y0)

        def on_motion(event):
            if not self._inset_dragging or self._position_inset_ax is None:
                return
            if event.x is None or event.y is None:
                return
            bbox = self._position_inset_ax.bbox
            if bbox is None:
                return
            new_x = event.x - self._inset_drag_offset[0]
            new_y = event.y - self._inset_drag_offset[1]
            try:
                inv = self.ax.transAxes.inverted()
            except Exception:
                return
            ax_coords = inv.transform((new_x, new_y))
            width = self._inset_bbox[2] if self._inset_bbox is not None else 0.28
            height = self._inset_bbox[3] if self._inset_bbox is not None else 0.28
            x0 = min(max(ax_coords[0], 0.0), 1.0 - width)
            y0 = min(max(ax_coords[1], 0.0), 1.0 - height)
            if self._inset_bbox is None:
                self._inset_bbox = [x0, y0, width, height]
            else:
                self._inset_bbox[0] = x0
                self._inset_bbox[1] = y0
            self._position_inset_ax.set_axes_locator(InsetPosition(self.ax, self._inset_bbox))
            self.canvas.draw_idle()

        def on_release(event):
            if event.button != MouseButton.LEFT:
                return
            self._inset_dragging = False

        self._inset_drag_cids = [
            self.canvas.mpl_connect("button_press_event", on_press),
            self.canvas.mpl_connect("motion_notify_event", on_motion),
            self.canvas.mpl_connect("button_release_event", on_release),
        ]

    def _remove_inset_drag_handlers(self):
        cids = getattr(self, "_inset_drag_cids", None) or []
        if self.canvas:
            for cid in cids:
                try:
                    self.canvas.mpl_disconnect(cid)
                except Exception:
                    pass
        self._inset_drag_cids = []
        self._inset_dragging = False
        self._inset_drag_offset = (0.0, 0.0)
        self._suppress_drag_until_release = False

    def _collect_inset_markers(self, image_key, image_dims=None):
        markers = []
        if not self._curve_entries:
            coords = self._spec_thumbnail_coords(dims=image_dims)
            if coords is not None:
                markers.append(("#ff3b6a", coords))
            return markers
        for entry in self._curve_entries:
            if image_key and entry.get("image_key") and entry.get("image_key") != image_key:
                continue
            color = entry.get("color", "#ff3b6a")
            coords = entry.get("coords")
            if coords is None:
                spec = entry.get("spec")
                if spec is None:
                    spec = self._resolve_spec_from_viewer(entry)
                    if spec:
                        entry["spec"] = spec
                if spec:
                    coords = self._spec_thumbnail_coords(spec=spec, file_key=entry.get("image_key"), dims=image_dims)
                    if coords:
                        entry["coords"] = coords
            if coords is not None:
                markers.append((color, coords))
        if not markers:
            coords = self._spec_thumbnail_coords(dims=image_dims)
            if coords is not None:
                markers.append(("#ff3b6a", coords))
        return markers

    def _qt_pos_hits_inset(self, pos):
        if self._position_inset_ax is None or pos is None:
            return False
        try:
            bbox = self._position_inset_ax.get_window_extent()
        except Exception:
            return False
        if bbox is None:
            return False
        try:
            dpr = float(self.canvas.devicePixelRatioF())
        except Exception:
            dpr = 1.0
        try:
            height = self.canvas.height() * dpr
        except Exception:
            height = self.canvas.height()
        x = float(pos.x()) * dpr
        y = float(height - (pos.y() * dpr))
        try:
            return bool(bbox.contains(x, y))
        except Exception:
            return False

    def eventFilter(self, source, event):
        if source == self.canvas:
            if event.type() == QtCore.QEvent.MouseButtonPress and event.button() == QtCore.Qt.LeftButton:
                if self._qt_pos_hits_inset(event.pos()):
                    self._drag_start_pos = None
                    self._suppress_drag_until_release = True
                else:
                    self._drag_start_pos = event.pos()
                    self._suppress_drag_until_release = False
            elif (
                event.type() == QtCore.QEvent.MouseMove
                and self._drag_start_pos is not None
                and not self._suppress_drag_until_release
            ):
                if (event.pos() - self._drag_start_pos).manhattanLength() >= QtWidgets.QApplication.startDragDistance():
                    self._start_drag()
                    self._drag_start_pos = None
            elif event.type() == QtCore.QEvent.MouseButtonRelease:
                self._drag_start_pos = None
                self._suppress_drag_until_release = False
            elif event.type() == QtCore.QEvent.Wheel and event.modifiers() & QtCore.Qt.ControlModifier:
                delta = event.angleDelta().y()
                if delta:
                    step = 0.12 if delta > 0 else -0.12
                    new_scale = max(0.6, min(2.4, self._font_scale + step))
                    if not math.isclose(new_scale, self._font_scale, rel_tol=1e-3, abs_tol=1e-3):
                        self._font_scale = new_scale
                        self._apply_font_scale()
                        self.canvas.draw_idle()
                event.accept()
                return True
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
        spec_path = payload.get("spec_path", "")
        channel = payload.get("channel")
        # Avoid duplicating the curve when a drag/drop occurs onto the same popup.
        for entry in self._curve_entries:
            if spec_path and spec_path == entry.get("spec_path") and (channel or "") == (entry.get("channel") or ""):
                QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), "Curve already present", self)
                return
        entry = {
            "label": label,
            "axis_vals": axis_vals,
            "values": values,
            "color": color,
            "lw": self._line_width,
            "ls": "-",
            "spec_path": spec_path,
            "channel": channel,
            "axis_label": payload.get("axis_label", self.axis_label),
            "axis_unit": payload.get("axis_unit", self.axis_unit),
            "matrix_index": payload.get("matrix_index"),
            "image_key": payload.get("image_key"),
            "coords": tuple(payload.get("coords")) if payload.get("coords") else None,
            "nm_coords": tuple(payload.get("nm_coords")) if payload.get("nm_coords") else None,
            "spec": payload.get("spec"),
        }
        self._attach_spec_metadata(entry)
        self._curve_entries.append(entry)
        self._selected_curve_index = len(self._curve_entries) - 1
        self._update_curve_list()
        self._plot_selected_channel()

    def add_external_spectrum(self, spec, channel=None, axis_key=None):
        """Append another spectroscopy entry into this popup."""
        if spec is None:
            return False
        channels = spec.get("channels") or {}
        channel_name = channel or self.channel_combo.currentText()
        data = np.asarray(channels.get(channel_name) or [], dtype=float)
        if data.size == 0:
            for name, values in channels.items():
                arr = np.asarray(values, dtype=float)
                if arr.size:
                    channel_name = name
                    data = arr
                    break
        if data.size == 0:
            return False
        axis_key = axis_key or self.axis_combo.currentData()
        axis_vals, axis_label, axis_unit = self._axis_values_for_spec(spec, axis_key)
        axis_vals = np.asarray(axis_vals, dtype=float)
        if axis_vals.size == 0:
            return False
        n = min(len(axis_vals), len(data))
        if n <= 0:
            return False
        if len(axis_vals) != len(data):
            axis_vals = axis_vals[:n]
            data = data[:n]
        payload = {
            "label": f"{Path(spec.get('path', '')).name} ({channel_name})",
            "axis_vals": axis_vals,
            "values": data,
            "color": self.SCIENCE_PALETTE[len(self._curve_entries) % len(self.SCIENCE_PALETTE)],
            "spec_path": str(Path(spec.get("path", ""))),
            "channel": channel_name,
            "axis_label": axis_label,
            "axis_unit": axis_unit,
            "matrix_index": spec.get("matrix_index"),
            "image_key": str(spec.get("image_key") or ""),
            "coords": None,
            "nm_coords": (spec.get("x"), spec.get("y")),
            "spec": spec,
        }
        self._add_entry_from_drop(payload)
        return True

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
            "matrix_index": entry.get("matrix_index"),
            "image_key": entry.get("image_key"),
            "coords": entry.get("coords"),
            "nm_coords": entry.get("nm_coords"),
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
        custom_btn = QtWidgets.QPushButton("Custom...")
        custom_btn.setToolTip("Pick any color with the system color dialog")
        custom_btn.clicked.connect(self._choose_custom_swatch_color)
        outer_layout.addWidget(custom_btn, 0, QtCore.Qt.AlignTop)
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
        self._update_curve_list()
        self._plot_selected_channel()

    def _choose_custom_swatch_color(self):
        current = QtGui.QColor(self._active_line_color or "#000000")
        color = QtWidgets.QColorDialog.getColor(current, self, "Select spectroscopy color")
        if not color.isValid():
            return
        hex_color = color.name()
        self._active_line_color = hex_color
        entry = self._current_entry()
        if entry:
            entry["color"] = hex_color
        self._set_active_swatch(None)
        self._update_curve_list()
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
        if self._legend_enabled:
            legend = self.ax.legend(loc=self._legend_loc or 'best', fontsize=self._legend_font)
            if legend:
                legend.set_draggable(True)
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
        self.fit_result_label.setVisible(True)

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
        if not enable:
            self.fit_result_label.clear()
            self.fit_result_label.setVisible(False)

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
        t0 = time.perf_counter()
        super().__init__(parent)
        self.image_entry = image_entry
        self.specs = list(specs)
        self.viewer = parent
        self._plot_font_family = normalize_font_family(getattr(self.viewer, "_plot_font_family", None), "sans-serif")
        self._plot_font_bold = bool(getattr(self.viewer, "_plot_font_bold", False))
        self._plot_font_italic = bool(getattr(self.viewer, "_plot_font_italic", False))
        self._plot_font_underline = bool(getattr(self.viewer, "_plot_font_underline", False))
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
        try:
            log_status(f"[Perf] Matrix explorer init: {(time.perf_counter()-t0)*1000:.0f} ms | specs={len(self.specs)} markers={len(self.specs)}")
        except Exception:
            pass
        # Debounced redraw timer to avoid heavy repaints during window resize/move
        self._resize_timer = QtCore.QTimer(self)
        self._resize_timer.setSingleShot(True)
        self._resize_timer.timeout.connect(self._draw_image_layer)
        # Suppress canvas updates while the window is being moved; re-enable shortly after movement stops
        self._move_timer = QtCore.QTimer(self)
        self._move_timer.setSingleShot(True)
        self._move_timer.timeout.connect(self._end_move_updates)
        self._movement_active = False

        # Initialize state and wiring
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
        # Guard against palette_name not being provided by callers
        palette_choice = palette_name or getattr(self.viewer, "spectro_color_cycle", DEFAULT_COLOR_CYCLE)
        self.palette_name = palette_choice
        self._color_palette = get_color_cycle(palette_choice)
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

    def resizeEvent(self, event):
        super().resizeEvent(event)
        # Delay redraw to after resize to reduce jank
        try:
            self._resize_timer.start(120)
        except Exception:
            self._draw_image_layer()

    def moveEvent(self, event):
        super().moveEvent(event)
        self._begin_move_updates()
        try:
            self._move_timer.start(150)
        except Exception:
            self._end_move_updates()

    def _begin_move_updates(self):
        if self._movement_active:
            return
        self._movement_active = True
        try:
            self.canvas.setUpdatesEnabled(False)
            self.curve_canvas.setUpdatesEnabled(False)
        except Exception:
            pass

    def _end_move_updates(self):
        if not self._movement_active:
            return
        self._movement_active = False
        try:
            self.canvas.setUpdatesEnabled(True)
            self.curve_canvas.setUpdatesEnabled(True)
        except Exception:
            pass
        try:
            self._draw_image_layer()
            self._update_plot()
        except Exception:
            pass

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

        add_font_menu_action(
            menu,
            self,
            self._plot_font_family,
            self.set_plot_font_family,
            current_style=self._font_style_state(),
            apply_style_callback=self.set_plot_typography,
        )
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

    def _font_style_state(self):
        return {
            "bold": bool(getattr(self, "_plot_font_bold", False)),
            "italic": bool(getattr(self, "_plot_font_italic", False)),
            "underline": bool(getattr(self, "_plot_font_underline", False)),
        }

    def set_plot_typography(self, **changes):
        """Refresh the matrix explorer typography."""
        family = changes.get("family", None)
        viewer = getattr(self, "viewer", None)
        if family is not None:
            family = normalize_font_family(family, "sans-serif")
            self._plot_font_family = family
        if viewer is not None and hasattr(viewer, "set_plot_typography"):
            target = {
                "family": family if family is not None else self._plot_font_family,
                "bold": bool(changes.get("bold", self._plot_font_bold)),
                "italic": bool(changes.get("italic", self._plot_font_italic)),
                "underline": bool(changes.get("underline", self._plot_font_underline)),
            }
            if any(getattr(viewer, f"_plot_font_{k}", None) != v for k, v in target.items()):
                try:
                    viewer.set_plot_typography(**target)
                    return
                except Exception:
                    pass
        for key, attr in (("bold", "_plot_font_bold"), ("italic", "_plot_font_italic"), ("underline", "_plot_font_underline")):
            if key in changes:
                setattr(self, attr, bool(changes[key]))
        self._draw_image_layer()

    def set_plot_font_family(self, family: str):
        """Refresh the matrix explorer plot with a new shared font family."""
        family = normalize_font_family(family, "sans-serif")
        viewer = getattr(self, "viewer", None)
        if viewer is not None and hasattr(viewer, "set_plot_font_family") and getattr(viewer, "_plot_font_family", None) != family:
            try:
                viewer.set_plot_font_family(family)
                return
            except Exception:
                pass
        self.set_plot_typography(family=family)

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
        style = _style_kwargs(self._font_style_state())
        for text in list(self.ax.get_xticklabels()) + list(self.ax.get_yticklabels()):
            apply_text_style(text, family=self._plot_font_family, **style)
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


class KPFMFitTrendDialog(QtWidgets.QDialog):
    METRIC_OPTIONS = [
        ("LCPD", "lcpd"),
        ("a", "a"),
        ("c", "c"),
        ("RMSE", "rmse"),
    ]

    def __init__(self, rows, parent=None):
        super().__init__(parent)
        self._rows = list(rows or [])
        self.setWindowTitle("KPFM fits vs Z")
        self.resize(760, 500)
        try:
            self.setWindowFlag(QtCore.Qt.MSWindowsFixedSizeDialogHint, False)
        except Exception:
            pass
        try:
            self.setSizeGripEnabled(True)
        except Exception:
            pass
        try:
            self.setMinimumSize(520, 360)
        except Exception:
            pass
        layout = QtWidgets.QVBoxLayout(self)

        controls = QtWidgets.QHBoxLayout()
        controls.addWidget(QtWidgets.QLabel("Metric:"))
        self.metric_combo = QtWidgets.QComboBox()
        for label, key in self.METRIC_OPTIONS:
            self.metric_combo.addItem(label, key)
        self.metric_combo.currentIndexChanged.connect(self._update_plot)
        controls.addWidget(self.metric_combo)

        self.error_cb = QtWidgets.QCheckBox("Show errors")
        self.error_cb.setChecked(True)
        self.error_cb.toggled.connect(self._update_plot)
        controls.addWidget(self.error_cb)

        self.relative_z_cb = QtWidgets.QCheckBox("Relative Z")
        self.relative_z_cb.setToolTip("Plot Z relative to the minimum fitted Z value")
        self.relative_z_cb.toggled.connect(self._update_plot)
        controls.addWidget(self.relative_z_cb)

        self.sort_z_cb = QtWidgets.QCheckBox("Sort by Z")
        self.sort_z_cb.setChecked(True)
        self.sort_z_cb.toggled.connect(self._update_plot)
        controls.addWidget(self.sort_z_cb)
        controls.addStretch(1)
        layout.addLayout(controls)

        self.fig = Figure(figsize=(6, 4))
        self.canvas = FigureCanvas(self.fig)
        self.ax = self.fig.add_subplot(111)
        layout.addWidget(self.canvas, 1)

        self.status_label = QtWidgets.QLabel("")
        layout.addWidget(self.status_label)
        self._update_plot()

    def set_rows(self, rows):
        self._rows = list(rows or [])
        self._update_plot()

    def set_relative_z(self, checked: bool):
        self.relative_z_cb.blockSignals(True)
        self.relative_z_cb.setChecked(bool(checked))
        self.relative_z_cb.blockSignals(False)
        self._update_plot()

    def _metric_meta(self, key):
        if key == "lcpd":
            units = [str(row.get("lcpd_unit") or "").strip() for row in self._rows if row.get("lcpd_unit")]
            unit = units[0] if units else ""
            return "LCPD", unit
        if key == "a":
            return "a", ""
        if key == "c":
            return "c", "Hz"
        if key == "rmse":
            return "RMSE", ""
        return key, ""

    def _update_plot(self):
        self.ax.clear()
        metric_key = self.metric_combo.currentData() or "lcpd"
        rows = []
        err_key = f"{metric_key}_err"
        for row in self._rows:
            try:
                z_val = float(row.get("z_nm"))
                y_val = row.get(metric_key)
                if y_val is None or not np.isfinite(float(y_val)):
                    continue
                y_val = float(y_val)
            except Exception:
                continue
            y_err = row.get(err_key)
            try:
                y_err = float(y_err) if y_err is not None and np.isfinite(float(y_err)) else None
            except Exception:
                y_err = None
            rows.append((z_val, y_val, y_err, row))
        if self.sort_z_cb.isChecked():
            rows.sort(key=lambda item: item[0])
        if not rows:
            self.ax.text(0.5, 0.5, "No fitted spectra with Z metadata", ha="center", va="center", transform=self.ax.transAxes)
            self.ax.set_axis_off()
            self.status_label.setText("No fit results with usable Z metadata.")
            self.canvas.draw_idle()
            return

        self.ax.set_axis_on()
        z_vals = np.asarray([item[0] for item in rows], dtype=float)
        if self.relative_z_cb.isChecked():
            z_plot = z_vals - float(np.nanmin(z_vals))
            x_label = "Z relative (nm)"
        else:
            z_plot = z_vals
            x_label = "Z (nm)"
        y_vals = np.asarray([item[1] for item in rows], dtype=float)
        y_errs = [item[2] for item in rows]
        use_errors = self.error_cb.isChecked() and any(err is not None for err in y_errs)
        plotted_yerr = None
        if use_errors:
            plotted_yerr = np.asarray([0.0 if err is None else float(err) for err in y_errs], dtype=float)
        self.ax.errorbar(
            z_plot,
            y_vals,
            yerr=plotted_yerr,
            fmt="o-",
            color="#1f77b4",
            ecolor="#4c78a8",
            capsize=3,
            lw=1.3,
            markersize=4.5,
        )
        label, unit = self._metric_meta(metric_key)
        self.ax.set_xlabel(x_label)
        self.ax.set_ylabel(f"{label} ({unit})" if unit else label)
        self.ax.grid(True, alpha=0.25)
        self.ax.xaxis.set_major_locator(MaxNLocator(nbins=6))
        self.ax.yaxis.set_major_locator(MaxNLocator(nbins=6))
        self.status_label.setText(f"{len(rows)} fitted spectra shown.")
        try:
            self.fig.tight_layout()
        except Exception:
            pass
        self.canvas.draw_idle()


class SpectroscopyCompareDialog(QtWidgets.QDialog):
    """Modern comparison UI for spectroscopy overlays and fitting."""
    def __init__(self, specs, parent=None, palette_name=None):
        t0 = time.perf_counter()
        super().__init__(parent)
        self.specs = list(specs)
        self.viewer = parent if hasattr(parent, "headers") else getattr(parent, "viewer", None)
        self.headers = getattr(self.viewer, "headers", {}) if self.viewer is not None else {}
        self._palette_name = palette_name or DEFAULT_COLOR_CYCLE
        self._color_cycle = get_color_cycle(self._palette_name)
        if not self._color_cycle:
            self._color_cycle = get_color_cycle(DEFAULT_COLOR_CYCLE)
        self._line_map = {}
        self._legend_map = {}
        self._fit_results = {}
        self._fit_thread = None
        self._fit_worker = None
        self._fit_trend_dialog = None
        self._popup_refs = []
        self._compare_inset_image_cache = OrderedDict()
        self._curve_data_cache = OrderedDict()
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
        try:
            self.setWindowFlag(QtCore.Qt.MSWindowsFixedSizeDialogHint, False)
        except Exception:
            pass
        try:
            self.setSizeGripEnabled(True)
        except Exception:
            pass
        try:
            self.setMinimumSize(760, 420)
        except Exception:
            pass
        self._plot_grid_enabled = True
        self._plot_legend_enabled = True
        self._plot_x_log = False
        self._plot_y_log = False
        self._plot_line_width = 1.6
        self._figure_preset_key = "interactive"
        self._show_position_inset = True
        self._position_inset_ax = None
        self._inset_bbox = None
        self._minima_artists = []
        self._inset_dragging = False
        self._inset_drag_offset = (0.0, 0.0)
        self._minima_meta = []
        self._dragging_minima = None
        self._point_labels = []
        self._point_label_drag = None
        self._last_mouse_xy = None
        self._curve_styles = {}  # spec_id -> {color, lw, ls}
        self._plotted_spec_ids = []
        self._legend_loc = "best"
        self._legend_font = 8
        self._legend_bg = True
        self._legend_border = True
        self._legend_filename_only = False
        self._plot_font_family = normalize_font_family(getattr(self.viewer, "_plot_font_family", None), "sans-serif")
        self._plot_font_bold = bool(getattr(self.viewer, "_plot_font_bold", False))
        self._plot_font_italic = bool(getattr(self.viewer, "_plot_font_italic", False))
        self._plot_font_underline = bool(getattr(self.viewer, "_plot_font_underline", False))
        self._grid_major = True
        self._grid_minor = False
        self._grid_alpha = 0.25
        self._grid_lw = 0.8
        self._grid_ls = "--"
        self._tick_cfg = {
            "x": {"direction": "out", "major": None, "minor_count": 0, "length": 6},
            "y": {"direction": "out", "major": None, "minor_count": 0, "length": 6},
        }
        self._filter_controls = {}
        self._filter_cfg = {
            "gaussian": {"enabled": False, "sigma": 1.0},
            "savgol": {"enabled": False, "window": 11, "poly": 3},
            "median": {"enabled": False, "size": 3},
            "fft": {"enabled": False, "cutoff": 0.15},
            "notch": {"enabled": False, "freq": 50.0, "width": 5.0},
            "derive": {"enabled": False, "window": 11, "poly": 3},
        }
        self._build_ui()
        self._populate_list()
        self._populate_channels()
        self._populate_axes()
        self._update_plot()
        self._suppress_undo_push = False
        try:
            log_status(f"[Perf] Spectroscopy comparison init: {(time.perf_counter()-t0)*1000:.0f} ms | spectra={len(self.specs)}")
        except Exception:
            pass
        # Debounced redraw timer to smooth window resize/move
        self._resize_timer = QtCore.QTimer(self)
        self._resize_timer.setSingleShot(True)
        self._resize_timer.timeout.connect(self._update_plot)
        self._plot_update_timer = QtCore.QTimer(self)
        self._plot_update_timer.setSingleShot(True)
        self._plot_update_timer.timeout.connect(self._flush_requested_plot_update)
        # Suppress canvas updates while the window is being moved; re-enable shortly after movement stops
        self._move_timer = QtCore.QTimer(self)
        self._move_timer.setSingleShot(True)
        self._move_timer.timeout.connect(self._end_move_updates)
        self._movement_active = False
        self._plot_update_pending = False
        self._last_hint_text = None
        self._last_mouse_text = None
        self._last_canvas_cursor = None
        self._last_hover_update_ts = 0.0

    def resizeEvent(self, event):
        super().resizeEvent(event)
        try:
            self._resize_timer.start(120)
        except Exception:
            self._update_plot()

    def moveEvent(self, event):
        super().moveEvent(event)
        self._begin_move_updates()
        try:
            self._move_timer.start(150)
        except Exception:
            self._end_move_updates()

    def _begin_move_updates(self):
        if self._movement_active:
            return
        self._movement_active = True
        try:
            self.canvas.setUpdatesEnabled(False)
        except Exception:
            pass

    def _end_move_updates(self):
        if not self._movement_active:
            return
        self._movement_active = False
        try:
            self.canvas.setUpdatesEnabled(True)
        except Exception:
            pass
        try:
            self._update_plot()
        except Exception:
            pass

    def _request_plot_update(self, *, delay_ms=35):
        self._plot_update_pending = True
        try:
            self._plot_update_timer.start(max(0, int(delay_ms)))
        except Exception:
            self._flush_requested_plot_update()

    def _flush_requested_plot_update(self):
        if not getattr(self, "_plot_update_pending", False):
            return
        if getattr(self, "_movement_active", False):
            try:
                self._plot_update_timer.start(80)
            except Exception:
                pass
            return
        self._plot_update_pending = False
        self._update_plot()

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

    def _filter_signature(self):
        cfg = self._filter_cfg or {}
        return (
            bool(cfg.get("gaussian", {}).get("enabled")), float(cfg.get("gaussian", {}).get("sigma", 1.0)),
            bool(cfg.get("savgol", {}).get("enabled")), int(cfg.get("savgol", {}).get("window", 11)), int(cfg.get("savgol", {}).get("poly", 3)),
            bool(cfg.get("median", {}).get("enabled")), int(cfg.get("median", {}).get("size", 3)),
            bool(cfg.get("fft", {}).get("enabled")), float(cfg.get("fft", {}).get("cutoff", 0.15)),
            bool(cfg.get("notch", {}).get("enabled")), float(cfg.get("notch", {}).get("freq", 50.0)), float(cfg.get("notch", {}).get("width", 5.0)),
            bool(cfg.get("derive", {}).get("enabled")), int(cfg.get("derive", {}).get("window", 11)), int(cfg.get("derive", {}).get("poly", 3)),
        )

    def _clear_curve_data_cache(self):
        self._curve_data_cache.clear()

    def _decimate_curve_for_display(self, x_vals, y_vals, plotted_count=1):
        try:
            x_arr = np.asarray(x_vals, dtype=float)
            y_arr = np.asarray(y_vals, dtype=float)
        except Exception:
            return x_vals, y_vals
        n = int(min(x_arr.size, y_arr.size))
        if n <= 0:
            return x_vals, y_vals
        try:
            canvas_width = max(320, int(self.canvas.width()))
        except Exception:
            canvas_width = 800
        if plotted_count >= 24:
            max_points = canvas_width
        elif plotted_count >= 12:
            max_points = int(canvas_width * 1.35)
        else:
            max_points = int(canvas_width * 2.0)
        max_points = max(256, max_points)
        if n <= max_points:
            return x_arr, y_arr
        step = max(1, int(math.ceil(n / float(max_points))))
        idx = np.arange(0, n, step, dtype=int)
        if idx.size == 0 or idx[-1] != (n - 1):
            idx = np.append(idx, n - 1)
        return x_arr[idx], y_arr[idx]

    def _curve_payload_for_plot(self, spec, spec_id, channel, axis_choice, bg_id, filter_sig):
        cache_key = (spec_id, channel, axis_choice, bg_id, filter_sig)
        cached_curve = self._curve_data_cache.get(cache_key)
        if cached_curve is not None:
            self._curve_data_cache.move_to_end(cache_key)
            return {
                "x": np.array(cached_curve["x"], copy=True),
                "y": np.array(cached_curve["y"], copy=True),
                "axis_label": cached_curve["axis_label"],
                "axis_unit": cached_curve["axis_unit"],
                "axis_unit_plot": cached_curve["axis_unit_plot"],
                "y_unit_raw": cached_curve["y_unit_raw"],
                "y_unit_final": cached_curve["y_unit_final"],
            }
        channels = spec.get("channels") or {}
        data = channels.get(channel)
        axis_vals, axis_label, axis_unit = self._axis_for_spec(spec)
        if data is None or not axis_vals.size:
            return None
        bg_spec = self._background_for(spec)
        y_base = self._subtract_background(axis_vals, data, bg_spec)
        x_vals = axis_vals
        axis_unit_plot = axis_unit
        if axis_unit.lower() == "v" and np.isfinite(x_vals).any():
            axis_unit_plot = "mV"
            x_vals = x_vals * 1000.0
        y_unit_raw = self._channel_unit_for_spec(spec, channel)
        y_filtered, y_unit_final = self._apply_data_filters(x_vals, y_base, y_unit_raw, axis_unit_plot)
        payload = {
            "x": np.array(x_vals, copy=True),
            "y": np.array(y_filtered, copy=True),
            "axis_label": axis_label,
            "axis_unit": axis_unit,
            "axis_unit_plot": axis_unit_plot,
            "y_unit_raw": y_unit_raw,
            "y_unit_final": y_unit_final,
        }
        self._curve_data_cache[cache_key] = {
            "x": np.array(payload["x"], copy=True),
            "y": np.array(payload["y"], copy=True),
            "axis_label": axis_label,
            "axis_unit": axis_unit,
            "axis_unit_plot": axis_unit_plot,
            "y_unit_raw": y_unit_raw,
            "y_unit_final": y_unit_final,
        }
        while len(self._curve_data_cache) > 512:
            self._curve_data_cache.popitem(last=False)
        return payload

    def _clear_empty_plot_text(self):
        txt = getattr(self, "_empty_plot_text", None)
        if txt is not None:
            try:
                txt.remove()
            except Exception:
                pass
        self._empty_plot_text = None

    def _set_empty_plot_text(self):
        self._clear_empty_plot_text()
        try:
            self._empty_plot_text = self.ax.text(
                0.5,
                0.5,
                "No data for selected items",
                ha="center",
                va="center",
                transform=self.ax.transAxes,
            )
        except Exception:
            self._empty_plot_text = None

    def _rebuild_compare_legend(self, plotted_spec_ids):
        self._legend_map.clear()
        existing = self.ax.get_legend()
        if existing is not None:
            try:
                existing.remove()
            except Exception:
                pass
        if not self._plot_legend_enabled or not plotted_spec_ids:
            return
        legend_loc = self._legend_loc or ("upper right" if len(plotted_spec_ids) >= 12 else "best")
        legend = self.ax.legend(loc=legend_loc, fontsize=self._legend_font)
        if not legend:
            return
        legend.set_draggable(len(plotted_spec_ids) < 16)
        try:
            frame = legend.get_frame()
            frame.set_alpha(0.9 if self._legend_bg else 0.0)
            frame.set_facecolor("white" if self._legend_bg else (0, 0, 0, 0))
            frame.set_edgecolor("black" if self._legend_border else (0, 0, 0, 0))
            frame.set_linewidth(0.8 if self._legend_border else 0.0)
        except Exception:
            pass
        for leg_line, spec_id in zip(legend.get_lines(), plotted_spec_ids):
            try:
                leg_line.set_picker(True)
            except Exception:
                pass
            self._legend_map[leg_line] = spec_id

    def _can_incremental_plot_update(self):
        if not self._line_map:
            return False
        if getattr(self, "_fit_results", None):
            return False
        if getattr(self, "_minima_meta", None):
            return False
        if getattr(self, "_point_labels", None):
            return False
        if getattr(self, "_delta_annotation_artists", None):
            return False
        return True

    def _update_plot_incremental(self):
        plot_items = self._visible_plot_items()
        if not plot_items:
            return False
        old_plotted_ids = list(getattr(self, "_plotted_spec_ids", []))
        channel = self.channel_combo.currentText()
        waterfall = self.waterfall_cb.isChecked()
        show_points = self.show_points_cb.isChecked()
        show_lines = self.lines_cb.isChecked()
        offset_val = self.offset_spin.value()
        relative_nm = bool(self._relative_zero_enabled)
        axis_choice = self.axis_combo.currentData() if getattr(self, "axis_combo", None) is not None else "primary"
        bg_id = self._background_spec_id or ""
        filter_sig = self._filter_signature()
        selected_ids = {item.data(0, QtCore.Qt.UserRole + 1) for item in self._selected_items()}
        colors = self._iter_color_cycle()
        visible_count = max(1, len(plot_items))
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
        y_units_after_filters = []
        plotted = 0
        plotted_spec_ids = []
        active_ids = set()
        xlabel = "Axis"
        for item in plot_items:
            spec = item.data(0, QtCore.Qt.UserRole)
            spec_id = item.data(0, QtCore.Qt.UserRole + 1)
            payload = self._curve_payload_for_plot(spec, spec_id, channel, axis_choice, bg_id, filter_sig)
            if payload is None:
                line = self._line_map.pop(spec_id, None)
                if line is not None:
                    try:
                        line.remove()
                    except Exception:
                        pass
                continue
            x_vals = payload["x"]
            y_filtered = payload["y"]
            axis_label = payload["axis_label"]
            axis_unit = payload["axis_unit"]
            y_unit_raw = payload["y_unit_raw"]
            y_unit_final = payload["y_unit_final"]
            y_data = y_filtered + (plotted * offset_val) if waterfall else y_filtered
            x_plot = x_vals - rel_zero if (relative_nm and axis_unit == "nm") else x_vals
            x_plot, y_plot = self._decimate_curve_for_display(x_plot, y_data, visible_count)
            if not x_plot.size:
                continue
            y_units_after_filters.append(y_unit_final or y_unit_raw)
            color = next(colors)
            highlight = spec_id in selected_ids or not selected_ids
            label_txt = self._display_name(spec)
            if self._legend_filename_only:
                try:
                    label_txt = Path(spec.get("path") or "").name or label_txt
                except Exception:
                    pass
            line = self._line_map.get(spec_id)
            if line is None:
                line, = self.ax.plot([], [])
                self._line_map[spec_id] = line
            style = self._curve_styles.get(spec_id) or {}
            line.set_data(x_plot, y_plot)
            line.set_label(label_txt)
            line.set_visible(True)
            line.set_color(style.get("color") or color)
            width = float(style.get("lw", self._plot_line_width))
            line.set_linewidth(width * (1.35 if highlight else 0.85))
            line.set_alpha(1.0 if highlight else 0.4)
            line.set_linestyle(style.get("ls") or ("-" if show_lines else "None"))
            line.set_marker("o" if show_points else "None")
            if show_points:
                try:
                    line.set_markersize(2.6)
                    line.set_markerfacecolor(style.get("color") or color)
                    line.set_markeredgecolor(style.get("color") or color)
                    line.set_markeredgewidth(0.6)
                except Exception:
                    pass
            active_ids.add(spec_id)
            plotted_spec_ids.append(spec_id)
            plotted += 1
            if relative_nm:
                xlabel = "Z (nm, relative)"
            elif axis_label:
                if axis_unit and axis_unit.lower() == "v":
                    xlabel = f"{axis_label} (mV)"
                elif axis_unit and axis_unit not in str(axis_label):
                    xlabel = f"{axis_label} ({axis_unit})"
                else:
                    xlabel = axis_label
        for spec_id in list(self._line_map.keys()):
            if spec_id not in active_ids:
                line = self._line_map.pop(spec_id, None)
                if line is not None:
                    try:
                        line.remove()
                    except Exception:
                        pass
        if not plotted_spec_ids:
            return False
        self._clear_empty_plot_text()
        self.ax.set_xlabel(xlabel)
        unit = next((val for val in y_units_after_filters if val), None)
        self.ax.set_ylabel(f"{channel} ({unit})" if unit else channel)
        self.ax.set_xscale("log" if self._plot_x_log else "linear")
        self.ax.set_yscale("log" if self._plot_y_log else "linear")
        self._apply_grid_and_ticks()
        self.ax.relim()
        self.ax.autoscale_view()
        if plotted_spec_ids != old_plotted_ids or self.ax.get_legend() is None:
            self._rebuild_compare_legend(plotted_spec_ids)
            self._update_position_inset_compare()
        else:
            self._apply_legend_settings()
        self._apply_font_scale()
        self._plotted_spec_ids = plotted_spec_ids
        self._update_status(plotted)
        return True

    def _visible_plot_items(self):
        items = []
        root = self.spec_list.invisibleRootItem()
        for i in range(root.childCount()):
            item = root.child(i)
            if item.isHidden():
                continue
            if item.checkState(0) == QtCore.Qt.Checked or item.isSelected():
                items.append(item)
        return items

    def _restyle_existing_lines(self):
        if not self._line_map or not self._plotted_spec_ids:
            return False
        selected_ids = {item.data(0, QtCore.Qt.UserRole + 1) for item in self._selected_items()}
        base_width = self._plot_line_width
        show_points = self.show_points_cb.isChecked()
        show_lines = self.lines_cb.isChecked()
        for spec_id in self._plotted_spec_ids:
            line = self._line_map.get(spec_id)
            if line is None:
                continue
            style = self._curve_styles.get(spec_id) or {}
            width = float(style.get("lw", base_width))
            highlight = spec_id in selected_ids or not selected_ids
            line.set_linewidth(width * (1.35 if highlight else 0.85))
            line.set_alpha(1.0 if highlight else 0.4)
            line.set_linestyle(style.get("ls") or ("-" if show_lines else "None"))
            marker = "o" if show_points else "None"
            line.set_marker(marker)
            if show_points:
                color = style.get("color") or line.get_color()
                try:
                    line.set_markersize(2.6)
                    line.set_markerfacecolor(color)
                    line.set_markeredgecolor(color)
                    line.set_markeredgewidth(0.6)
                except Exception:
                    pass
        self.canvas.draw_idle()
        return True

    def _apply_palette_to_existing_lines(self):
        if not self._line_map or not self._plotted_spec_ids:
            return False
        cycle = list(self._color_cycle or [])
        if not cycle:
            cycle = get_color_cycle(DEFAULT_COLOR_CYCLE)
        if not cycle:
            return False
        for idx, spec_id in enumerate(self._plotted_spec_ids):
            line = self._line_map.get(spec_id)
            if line is None:
                continue
            style = self._curve_styles.get(spec_id) or {}
            if style.get("color"):
                color = style.get("color")
            else:
                color = cycle[idx % len(cycle)]
            try:
                line.set_color(color)
                line.set_markerfacecolor(color)
                line.set_markeredgecolor(color)
            except Exception:
                pass
        legend = self.ax.get_legend()
        if legend:
            for idx, leg_line in enumerate(legend.get_lines()):
                if idx >= len(self._plotted_spec_ids):
                    break
                spec_id = self._plotted_spec_ids[idx]
                style = self._curve_styles.get(spec_id) or {}
                color = style.get("color") or cycle[idx % len(cycle)]
                try:
                    leg_line.set_color(color)
                    leg_line.set_markerfacecolor(color)
                    leg_line.set_markeredgecolor(color)
                except Exception:
                    pass
        self.canvas.draw_idle()
        return True

    def _fit_result_headers(self):
        return [
            "File",
            "X (nm)",
            "Y (nm)",
            "Z rel (nm)" if self._fit_z_display_relative_enabled() else "Z (nm)",
            "a",
            "da",
            "LCPD",
            "dLCPD",
            "c (Hz)",
            "dc",
            "RMSE",
        ]

    def _fit_z_display_relative_enabled(self):
        checkbox = getattr(self, "fit_relative_z_cb", None)
        return bool(checkbox and checkbox.isChecked())

    def _sequence_z_offset_nm(self):
        values = []
        for spec in list(getattr(self, "specs", []) or []):
            z_val, _label = self._resolve_spec_z_value(spec)
            if z_val is not None:
                values.append(float(z_val))
        if not values:
            return 0.0
        return float(min(values))

    def _refresh_fit_result_headers(self):
        headers = self._fit_result_headers()
        table = getattr(self, "results_table", None)
        if table is None or table.columnCount() != len(headers):
            return
        for idx, text in enumerate(headers):
            item = table.horizontalHeaderItem(idx)
            if item is None:
                item = QtWidgets.QTableWidgetItem(text)
                table.setHorizontalHeaderItem(idx, item)
            else:
                item.setText(text)

    def _resolve_spec_z_value(self, spec):
        if not spec:
            return None, ""
        for key, label in (
            ("z_level_nm", spec.get("z_level_label") or "Z"),
            ("xy_stack_z_level_nm", spec.get("xy_stack_z_label") or spec.get("z_level_label") or "Z"),
            ("z_abs_nm", "Z abs"),
            ("z_nm", "Z"),
        ):
            value = spec.get(key)
            try:
                if value is not None and np.isfinite(float(value)):
                    z_val = float(value)
                    if key == "xy_stack_z_level_nm" and spec.get("z_level_nm") is None:
                        spec["z_level_nm"] = z_val
                        spec["z_level_label"] = str(label)
                        spec["z_level_unit"] = "nm"
                    return z_val, str(label or "Z")
            except Exception:
                continue

        level, label = _metadata_z_from_spec(spec)
        if level is not None:
            spec["z_level_nm"] = float(level)
            spec["z_level_label"] = str(label or "Z")
            spec["z_level_unit"] = "nm"
            return float(level), str(label or "Z")

        for axis in list(spec.get("AxisChoices") or []):
            label = str(axis.get("label") or axis.get("key") or "Z")
            if not _z_like_name(label) and str(axis.get("key") or "").strip().lower() not in {"z", "topo"}:
                continue
            level = _constant_axis_value_nm(axis.get("values"), axis.get("unit") or "")
            if level is not None:
                spec["z_level_nm"] = float(level)
                spec["z_level_label"] = label
                spec["z_level_unit"] = "nm"
                return float(level), label

        extra_topo = _topo_axis_from_spec(spec)
        if extra_topo is not None:
            level = _constant_axis_value_nm(extra_topo.get("values"), extra_topo.get("unit") or "nm")
            if level is not None:
                label = str(extra_topo.get("label") or "Topo")
                spec["z_level_nm"] = float(level)
                spec["z_level_label"] = label
                spec["z_level_unit"] = "nm"
                return float(level), label

        value = spec.get("topo_nm")
        try:
            if value is not None and np.isfinite(float(value)):
                return float(value), "Topo"
        except Exception:
            pass

        return None, ""

    def _format_fit_result_z(self, spec):
        if not spec:
            return "n/a", ""
        z_num, label = self._resolve_spec_z_value(spec)
        if z_num is None:
            return "n/a", ""
        label_text = str(label or "Z").strip() or "Z"
        if self._fit_z_display_relative_enabled():
            offset = self._sequence_z_offset_nm()
            z_rel = z_num - offset
            return f"{z_rel:.6g}", f"{label_text}: absolute {z_num:.6g} nm | relative {z_rel:.6g} nm"
        return f"{z_num:.6g}", f"{label_text}: absolute {z_num:.6g} nm"

    def _fit_trend_rows(self):
        rows = []
        for spec_id, res in self._fit_results.items():
            spec = res.get("spec")
            if not spec:
                continue
            z_nm, z_label = self._resolve_spec_z_value(spec)
            if z_nm is None:
                continue
            axis_unit = str(res.get("axis_unit") or "").strip()
            lcpd_unit = "mV" if axis_unit.lower() == "v" else (axis_unit or "")
            lcpd_scale = 1000.0 if axis_unit.lower() == "v" else 1.0
            v0 = res.get("v0")
            v0_err = res.get("v0_err")
            try:
                lcpd = float(v0) * lcpd_scale if v0 is not None and np.isfinite(float(v0)) else None
            except Exception:
                lcpd = None
            try:
                lcpd_err = float(v0_err) * lcpd_scale if v0_err is not None and np.isfinite(float(v0_err)) else None
            except Exception:
                lcpd_err = None
            rows.append({
                "spec_id": spec_id,
                "name": self._display_name(spec),
                "z_nm": float(z_nm),
                "z_label": str(z_label or "Z"),
                "lcpd": lcpd,
                "lcpd_err": lcpd_err,
                "lcpd_unit": lcpd_unit,
                "a": float(res["a"]) if np.isfinite(float(res["a"])) else None,
                "a_err": float(res["a_err"]) if np.isfinite(float(res["a_err"])) else None,
                "c": float(res["c"]) if np.isfinite(float(res["c"])) else None,
                "c_err": float(res["c_err"]) if np.isfinite(float(res["c_err"])) else None,
                "rmse": float(res["rmse"]) if np.isfinite(float(res["rmse"])) else None,
                "rmse_err": None,
            })
        return rows

    def _update_fit_trend_state(self):
        rows = self._fit_trend_rows()
        if hasattr(self, "fit_vs_z_btn"):
            self.fit_vs_z_btn.setEnabled(bool(rows))
        dlg = getattr(self, "_fit_trend_dialog", None)
        if dlg is not None:
            try:
                dlg.set_rows(rows)
            except Exception:
                pass

    def _show_fit_vs_z_dialog(self):
        rows = self._fit_trend_rows()
        if not rows:
            QtWidgets.QMessageBox.information(self, "KPFM fits vs Z", "No fitted spectra with usable Z metadata are available.")
            return
        dlg = getattr(self, "_fit_trend_dialog", None)
        if dlg is None or not dlg.isVisible():
            dlg = KPFMFitTrendDialog(rows, parent=self)
            try:
                dlg.setWindowModality(QtCore.Qt.NonModal)
                dlg.setAttribute(QtCore.Qt.WA_DeleteOnClose, True)
            except Exception:
                pass
            try:
                dlg.set_relative_z(self._fit_z_display_relative_enabled())
            except Exception:
                pass
            dlg.finished.connect(lambda _=None: setattr(self, "_fit_trend_dialog", None))
            self._fit_trend_dialog = dlg
            dlg.show()
        else:
            dlg.set_rows(rows)
            try:
                dlg.set_relative_z(self._fit_z_display_relative_enabled())
            except Exception:
                pass
            try:
                dlg.raise_()
                dlg.activateWindow()
            except Exception:
                pass

    def _get_icon(self, name):
        """Get a themed icon, falling back to empty icon if not available."""
        icon = QIcon.fromTheme(name)
        return icon if icon and not icon.isNull() else QIcon()

    def _build_ui(self):
        splitter = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        splitter.setChildrenCollapsible(False)
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
        left.setMinimumWidth(180)
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
        self.canvas.mpl_connect("button_press_event", self._on_inset_press)
        self.canvas.mpl_connect("motion_notify_event", self._on_inset_motion)
        self.canvas.mpl_connect("button_release_event", self._on_inset_release)
        self.canvas.mpl_connect("button_press_event", self._on_minima_press)
        self.canvas.mpl_connect("motion_notify_event", self._on_minima_motion)
        self.canvas.mpl_connect("button_release_event", self._on_minima_release)
        self.canvas.mpl_connect("motion_notify_event", self._on_mouse_move)
        self.canvas.mpl_connect("button_press_event", self._on_point_label_press)
        self.canvas.mpl_connect("motion_notify_event", self._on_point_label_motion)
        self.canvas.mpl_connect("button_release_event", self._on_point_label_release)
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
        self.mouse_label = QtWidgets.QLabel("x: —   y: —")
        self.mouse_label.setAccessibleName("Mouse position")
        self.mouse_label.setAccessibleDescription("Displays current mouse position on the plot")
        center_layout.addWidget(self.mouse_label)

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

        self.position_inset_cb = QtWidgets.QCheckBox("Position inset")
        self.position_inset_cb.setToolTip("Show miniature of the acquisition image with spectrum locations")
        self.position_inset_cb.setChecked(True)
        self.position_inset_cb.toggled.connect(self._on_visual_toggle)
        vis_row.addWidget(self.position_inset_cb)

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
        center.setMinimumWidth(420)
        splitter.addWidget(center)
        splitter.setStretchFactor(1, 2)
        self.canvas.mpl_connect('pick_event', self._on_legend_pick)

        # Right panel: controls + results
        right = QtWidgets.QScrollArea()
        right.setWidgetResizable(True)
        right.setFrameShape(QtWidgets.QFrame.NoFrame)
        right.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        right_content = QtWidgets.QWidget()
        right_layout = QtWidgets.QVBoxLayout(right_content)
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
        self.axis_combo.setToolTip("Select X-axis for plotting (bias voltage, Z, or Topo position)")
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
        filters_panel = self._build_filter_panel()
        if filters_panel:
            right_layout.addWidget(filters_panel)

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

        trend_row = QtWidgets.QHBoxLayout()
        self.fit_vs_z_btn = QtWidgets.QPushButton(self._get_icon("office-chart-line"), "Plot fits vs Z")
        self.fit_vs_z_btn.setToolTip("Open an optional plot of fitted KPFM quantities against Z with error bars")
        self.fit_vs_z_btn.setEnabled(False)
        trend_row.addWidget(self.fit_vs_z_btn)
        trend_row.addStretch(1)
        kpfm_layout.addLayout(trend_row)

        z_mode_row = QtWidgets.QHBoxLayout()
        self.fit_relative_z_cb = QtWidgets.QCheckBox("Relative Z values")
        self.fit_relative_z_cb.setToolTip("Display fitted spectroscopy Z values relative to the minimum Z of the sequence instead of absolute piezo extension")
        self.fit_relative_z_cb.toggled.connect(self._on_fit_z_display_toggled)
        z_mode_row.addWidget(self.fit_relative_z_cb)
        z_mode_row.addStretch(1)
        kpfm_layout.addLayout(z_mode_row)

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
        self.fit_vs_z_btn.clicked.connect(self._show_fit_vs_z_dialog)
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
        self.fit_vs_z_btn.setAccessibleName("Plot fits versus Z")
        self.fit_vs_z_btn.setAccessibleDescription("Open a plot of fitted KPFM values against Z with error bars")
        self.fit_relative_z_cb.setAccessibleName("Relative Z values")
        self.fit_relative_z_cb.setAccessibleDescription("Toggle fitted spectroscopy Z display between absolute piezo extension and values relative to the minimum Z in the sequence")
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

        fit_headers = self._fit_result_headers()
        self.results_table = QtWidgets.QTableWidget(0, len(fit_headers))
        self.results_table.setHorizontalHeaderLabels(fit_headers)
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

        right_layout.addStretch(0)
        right_content.setMinimumWidth(320)
        right.setWidget(right_content)
        right.setMinimumWidth(340)
        splitter.addWidget(right)
        splitter.setStretchFactor(2, 1)
        try:
            splitter.setSizes([220, 820, 380])
        except Exception:
            pass

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
            stack_label = str(spec.get("xy_stack_display") or "").strip()
            if stack_label:
                type_str = f"{type_str} {stack_label}"
            
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
            chans = _available_channel_names(spec)
            chans_str = ", ".join(chans)

            item = QtWidgets.QTreeWidgetItem([name, type_str, pos_str, time_str, chans_str])
            item.setFlags(item.flags() | QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsSelectable)
            item.setCheckState(0, QtCore.Qt.Checked)
            item.setData(0, QtCore.Qt.UserRole, spec)
            item.setData(0, QtCore.Qt.UserRole + 1, self._spec_id(spec))
            stack_summary = str(spec.get("xy_stack_summary") or "").strip()
            if stack_summary:
                item.setToolTip(0, stack_summary)
                item.setToolTip(1, stack_summary)
                item.setToolTip(2, stack_summary)
            self.spec_list.addTopLevelItem(item)
            self._item_map[self._spec_id(spec)] = item
        
        for i in range(5):
            self.spec_list.resizeColumnToContents(i)
        self.spec_list.blockSignals(False)

    def set_specs(self, specs):
        """Update the dialog with a new list of spectra without reopening."""
        self.specs = list(specs)
        self._fit_results = {}
        self._update_fit_trend_state()
        self._item_map = {}
        self._clear_curve_data_cache()
        self._plotted_spec_ids = []
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

    def _on_fit_z_display_toggled(self, checked):
        self._record_user_action(f"KPFM Z display → {'relative' if checked else 'absolute'}")
        self._refresh_fit_result_headers()
        self._populate_results_table()
        dlg = getattr(self, "_fit_trend_dialog", None)
        if dlg is not None:
            try:
                dlg.set_relative_z(bool(checked))
            except Exception:
                pass

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
        channels = sorted({name for spec in self.specs for name in _available_channel_names(spec)})
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
            extra_topo = _topo_axis_from_spec(spec)
            if extra_topo and not any(a[0] == "topo" for a in axes):
                axes.append(("topo", extra_topo.get("label") or "Topo", extra_topo.get("unit") or "nm", np.asarray(extra_topo.get("values", []), dtype=float)))
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
        self._request_plot_update(delay_ms=90)

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
        if choice_key == "topo":
            extra_topo = _topo_axis_from_spec(spec)
            if extra_topo is not None:
                vals = np.asarray(extra_topo.get("values", []), dtype=float)
                return vals, extra_topo.get("label") or "Topo", extra_topo.get("unit") or "nm"
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
        self._request_plot_update()

    def _on_clear_background(self):
        self._record_user_action("Clear background")
        self._background_spec_id = None
        self._request_plot_update()

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

    @staticmethod
    def _axis_to_meters(axis_vals: np.ndarray, axis_unit: str) -> Tuple[Optional[np.ndarray], Optional[np.ndarray]]:
        """Return (axis_m, axis_nm) or (None, None) if unit unsupported."""
        if axis_vals is None:
            return None, None
        unit = (axis_unit or "").strip().lower()
        scale = None
        if unit in ("m", "meter", "meters"):
            scale = 1.0
        elif unit in ("nm",):
            scale = 1e-9
        elif unit in ("pm",):
            scale = 1e-12
        elif unit in ("um",):
            scale = 1e-6
        if scale is None:
            return None, None
        axis_m = np.asarray(axis_vals, dtype=float) * scale
        axis_nm = axis_m * 1e9
        return axis_m, axis_nm

    @staticmethod
    def _force_from_freq_shift(z_m: np.ndarray, df_hz: np.ndarray, f0: float, k: float, amp_m: float) -> Optional[np.ndarray]:
        """Sader-Jarvis style force reconstruction (numeric trapz)."""
        if f0 <= 0 or k <= 0 or amp_m <= 0:
            return None
        z = np.asarray(z_m, dtype=float)
        df = np.asarray(df_hz, dtype=float)
        if z.size < 2 or df.size != z.size:
            return None
        order = np.argsort(z)
        z_sorted = z[order]
        df_sorted = df[order]
        F_sorted = np.zeros_like(df_sorted)
        n = len(z_sorted)
        for i in range(n - 1):
            u = z_sorted[i + 1 :] - z_sorted[i]
            if u.size == 0:
                continue
            integrand = (df_sorted[i + 1 :] / f0) * np.sin(u / amp_m)
            F_sorted[i] = 2.0 * k * amp_m * np.trapz(integrand, u)
        if n > 1:
            F_sorted[-1] = F_sorted[-2]
        F = np.empty_like(F_sorted)
        F[order] = F_sorted
        return F

    def _on_convert_force(self):
        # Prompt for parameters
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("Convert to force")
        form = QtWidgets.QFormLayout(dlg)
        f0_edit = QtWidgets.QDoubleSpinBox(); f0_edit.setRange(0, 1e9); f0_edit.setDecimals(3); f0_edit.setValue(0.0)
        k_edit = QtWidgets.QDoubleSpinBox(); k_edit.setRange(0, 1e6); k_edit.setDecimals(3); k_edit.setValue(0.0)
        a_edit = QtWidgets.QDoubleSpinBox(); a_edit.setRange(0, 1e-3); a_edit.setDecimals(11); a_edit.setSingleStep(1e-11); a_edit.setValue(50e-12)
        q_edit = QtWidgets.QDoubleSpinBox(); q_edit.setRange(0, 1e6); q_edit.setDecimals(2); q_edit.setValue(0.0)
        method_combo = QtWidgets.QComboBox(); method_combo.addItems(["saderF", "matrixF"])
        form.addRow("f0 (Hz)", f0_edit)
        form.addRow("Spring constant k (N/m)", k_edit)
        form.addRow("Amplitude A (m)", a_edit)
        form.addRow("Q", q_edit)
        form.addRow("Method", method_combo)
        # load persisted params if available
        cfg = load_config()
        last_force = cfg.get("force_params", {})
        try:
            f0_edit.setValue(float(last_force.get("f0", f0_edit.value())))
        except Exception:
            pass
        try:
            k_edit.setValue(float(last_force.get("k", k_edit.value())))
        except Exception:
            pass
        try:
            a_edit.setValue(float(last_force.get("A", a_edit.value())))
        except Exception:
            pass
        try:
            q_edit.setValue(float(last_force.get("Q", q_edit.value())))
        except Exception:
            pass
        mth = str(last_force.get("method") or "")
        if mth in [method_combo.itemText(i) for i in range(method_combo.count())]:
            method_combo.setCurrentText(mth)
        buttons_row = QtWidgets.QHBoxLayout()
        remember_btn = QtWidgets.QPushButton("Remember")
        clear_btn = QtWidgets.QPushButton("Clear saved")
        buttons_row.addWidget(remember_btn)
        buttons_row.addWidget(clear_btn)
        buttons_row.addStretch(1)
        form.addRow(buttons_row)
        btns = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        form.addRow(btns)
        btns.accepted.connect(dlg.accept); btns.rejected.connect(dlg.reject)
        def _remember():
            params = {
                "f0": f0_edit.value(),
                "k": k_edit.value(),
                "A": a_edit.value(),
                "Q": q_edit.value(),
                "method": method_combo.currentText(),
            }
            cfg = load_config()
            cfg["force_params"] = params
            save_config(cfg)
        def _clear():
            cfg = load_config()
            if "force_params" in cfg:
                cfg.pop("force_params", None)
                save_config(cfg)
        remember_btn.clicked.connect(_remember)
        clear_btn.clicked.connect(_clear)
        if dlg.exec_() != QtWidgets.QDialog.Accepted:
            return
        f0 = f0_edit.value(); k = k_edit.value(); A = a_edit.value(); Q = q_edit.value(); method = method_combo.currentText()
        if f0 <= 0 or k <= 0 or A <= 0:
            QtWidgets.QMessageBox.information(self, "Force conversion", "Enter positive values for f0, k, and A.")
            return
        items = self._selected_items() or self._checked_items()
        if not items:
            QtWidgets.QMessageBox.information(self, "Force conversion", "Select spectra to convert.")
            return
        new_specs = []
        failed = 0
        channel = self.channel_combo.currentText()
        for item in items:
            spec = item.data(0, QtCore.Qt.UserRole)
            if not spec:
                continue
            axis_vals, axis_label, axis_unit = self._axis_for_spec(spec)
            axis_m, axis_nm = self._axis_to_meters(axis_vals, axis_unit)
            if axis_m is None or axis_nm is None:
                failed += 1
                continue
            channels = spec.get("channels") or {}
            if channel not in channels:
                failed += 1
                continue
            y_vals = np.asarray(channels.get(channel), dtype=float)
            bg = self._background_for(spec)
            y_vals = self._subtract_background(axis_vals, y_vals, bg)
            force_curve = self._force_from_freq_shift(axis_m, y_vals, f0=f0, k=k, amp_m=A)
            if force_curve is None:
                failed += 1
                continue
            force_label = f"Force_{channel}"
            new_spec = dict(spec)
            new_spec["channels"] = {force_label: force_curve}
            new_spec["unit_map"] = {force_label: "N"}
            new_spec["AxisLabel"] = "Z"
            new_spec["AxisUnit"] = "nm"
            new_spec["V"] = axis_nm
            new_spec["AxisChoices"] = [{"key": "primary", "label": "Z", "unit": "nm", "values": axis_nm}]
            new_spec["ForceMethod"] = method
            new_spec["ForceParams"] = {"f0": f0, "A": A, "Q": Q, "k": k}
            new_specs.append(new_spec)
        if not new_specs:
            QtWidgets.QMessageBox.information(self, "Force conversion", "Could not convert the selected spectra (check that the axis is Z/distance and parameters are valid).")
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
        self._update_fit_trend_state()
        self._populate_results_table()
        self._validate_log_axes()
        self._request_plot_update(delay_ms=25)

    def _on_axis_changed(self):
        self._record_user_action(f"Axis → {self.axis_combo.currentText()}")
        self._fit_results = {}
        self._update_fit_trend_state()
        self.results_table.setRowCount(0)
        self._validate_log_axes()
        self._request_plot_update(delay_ms=25)

    def _on_relative_toggled(self, checked):
        self._record_user_action(f"Relative Z → {'on' if checked else 'off'}")
        self._relative_zero_enabled = bool(checked)
        self._request_plot_update(delay_ms=25)

    def _on_item_check_changed(self, item, column):
        self._record_user_action("Traffic: checked item changed")
        self._request_plot_update(delay_ms=20)

    def _on_list_selection_changed(self):
        self._record_user_action("Selection changed")
        target_ids = [item.data(0, QtCore.Qt.UserRole + 1) for item in self._visible_plot_items()]
        if target_ids == list(getattr(self, "_plotted_spec_ids", [])) and self._restyle_existing_lines():
            return
        self._request_plot_update(delay_ms=20)

    def _update_plot(self):
        if self._can_incremental_plot_update() and self._update_plot_incremental():
            return
        channel = self.channel_combo.currentText()
        self.ax.clear()
        self._empty_plot_text = None
        self._grid_major = bool(self._plot_grid_enabled)
        # Base grid handled in _apply_grid_and_ticks
        self.ax.grid(False)
        self._lcpd_line_info.clear()
        self._clear_delta_selection(redraw=False)
        self._line_map.clear()
        self._legend_map.clear()
        self._clear_minima_annotations()
        self._clear_point_labels(redraw=False)
        
        waterfall = self.waterfall_cb.isChecked()
        show_points = self.show_points_cb.isChecked()
        show_lines = self.lines_cb.isChecked()
        offset_val = self.offset_spin.value()
        scale = self._estimate_channel_scale(channel)
        self._configure_offset_spin(scale)
        relative_nm = bool(self._relative_zero_enabled)

        selected_ids = {item.data(0, QtCore.Qt.UserRole + 1) for item in self._selected_items()}
        colors = self._iter_color_cycle()
        plot_items = self._visible_plot_items()
        visible_count = max(1, len(plot_items))
        plotted = 0
        plotted_spec_ids = []
        axis_choice = self.axis_combo.currentData() if getattr(self, "axis_combo", None) is not None else "primary"
        bg_id = self._background_spec_id or ""
        filter_sig = self._filter_signature()

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
        y_units_after_filters = []
        for item in plot_items:
            spec = item.data(0, QtCore.Qt.UserRole)
            spec_id = item.data(0, QtCore.Qt.UserRole + 1)
            payload = self._curve_payload_for_plot(spec, spec_id, channel, axis_choice, bg_id, filter_sig)
            if payload is None:
                continue
            x_vals = payload["x"]
            y_filtered = payload["y"]
            axis_label = payload["axis_label"]
            axis_unit = payload["axis_unit"]
            y_unit_raw = payload["y_unit_raw"]
            y_unit_final = payload["y_unit_final"]
            # Apply waterfall offset
            y_data = y_filtered + (plotted * offset_val) if waterfall else y_filtered
            x_plot = x_vals - rel_zero if (relative_nm and axis_unit == "nm") else x_vals
            x_plot, y_plot = self._decimate_curve_for_display(x_plot, y_data, visible_count)
            y_units_after_filters.append(y_unit_final or y_unit_raw)
            color = next(colors)
            highlight = spec_id in selected_ids or not selected_ids
            label_txt = self._display_name(spec)
            if self._legend_filename_only:
                try:
                    label_txt = Path(spec.get("path") or "").name or label_txt
                except Exception:
                    pass
            base_width = self._plot_line_width
            line_kwargs = {
                "color": color,
                "lw": base_width * (1.35 if highlight else 0.85),
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
            line, = self.ax.plot(x_plot, y_plot, **line_kwargs)
            style = self._curve_styles.get(spec_id)
            if style:
                try:
                    if style.get("color"):
                        line.set_color(style.get("color"))
                    if style.get("lw"):
                        line.set_linewidth(style.get("lw"))
                    if style.get("ls"):
                        line.set_linestyle(style.get("ls"))
                except Exception:
                    pass
            self._line_map[spec_id] = line
            plotted_spec_ids.append(spec_id)
            plotted += 1
            if spec_id in self._fit_results:
                self._draw_fit_for_spec(spec_id, color, offset=(plotted - 1) * offset_val if waterfall else 0.0)
        if plotted == 0:
            self._set_empty_plot_text()
        elif self._plot_legend_enabled:
            legend_loc = self._legend_loc or ("upper right" if plotted >= 12 else "best")
            legend = self.ax.legend(loc=legend_loc, fontsize=self._legend_font)
            if legend:
                legend.set_draggable(plotted < 16)
                try:
                    frame = legend.get_frame()
                    frame.set_alpha(0.9 if self._legend_bg else 0.0)
                    frame.set_facecolor("white" if self._legend_bg else (0, 0, 0, 0))
                    frame.set_edgecolor("black" if self._legend_border else (0, 0, 0, 0))
                    frame.set_linewidth(0.8 if self._legend_border else 0.0)
                except Exception:
                    pass
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
        for val in y_units_after_filters:
            if val:
                unit = val
                break
        self.ax.set_ylabel(f"{channel} ({unit})" if unit else channel)
        self.ax.set_xscale("log" if self._plot_x_log else "linear")
        self.ax.set_yscale("log" if self._plot_y_log else "linear")
        self._apply_grid_and_ticks()
        self._update_position_inset_compare()
        self._apply_font_scale()
        # canvas.draw_idle() is called in _apply_font_scale
        self._plotted_spec_ids = plotted_spec_ids
        self._update_status(plotted)

    def _load_thumbnail_array_for_inset(self, file_key):
        viewer = getattr(self, "viewer", None)
        if not viewer or not file_key:
            return None
        cache_key = None
        try:
            width = int(getattr(viewer, "thumb_size_px", 160))
            height = max(48, int(round(width * 0.75)))
            cmap = viewer.thumb_cmap_combo.currentText() if hasattr(viewer, "thumb_cmap_combo") else None
            cmap = cmap or getattr(viewer, "thumb_cmap", "viridis")
            channel_idx = viewer.channel_dropdown.currentIndex() if hasattr(viewer, "channel_dropdown") else 0
            cache_key = (str(file_key), width, height, str(cmap), int(channel_idx))
            cached = self._compare_inset_image_cache.get(cache_key)
            if cached is not None:
                self._compare_inset_image_cache.move_to_end(cache_key)
                return np.array(cached, copy=True)
        except Exception:
            cache_key = None
        thumb = None
        label = getattr(viewer, "_thumb_labels", {}).get(file_key) if hasattr(viewer, "_thumb_labels") else None
        if label is not None and label.pixmap():
            thumb = label.pixmap()
        if thumb is None:
            try:
                thumb = viewer._thumbnail_pixmap_for_file(file_key, channel_idx, width, height, cmap)
            except Exception:
                return None
        if thumb is None:
            return None
        qimg = thumb.toImage().convertToFormat(QtGui.QImage.Format_RGBA8888)
        ptr = qimg.bits()
        ptr.setsize(qimg.byteCount())
        arr = np.frombuffer(ptr, dtype=np.uint8).reshape((qimg.height(), qimg.width(), 4))
        arr = arr[..., :3] / 255.0
        gray = np.clip(arr[..., 0] * 0.299 + arr[..., 1] * 0.587 + arr[..., 2] * 0.114, 0.0, 1.0)
        tinted = np.stack([gray, gray, gray], axis=-1)
        if cache_key is not None:
            self._compare_inset_image_cache[cache_key] = np.array(tinted, copy=True)
            while len(self._compare_inset_image_cache) > 6:
                self._compare_inset_image_cache.popitem(last=False)
        return tinted

    def _spec_thumbnail_coords_for_compare(self, spec=None, file_key=None, dims=None):
        viewer = getattr(self, "viewer", None)
        spec = spec or None
        if spec is None:
            items = self._checked_items() or self._selected_items()
            if items:
                spec = items[0].data(0, QtCore.Qt.UserRole)
        file_key = file_key or (str(spec.get("image_key") or "") if spec else "")
        if not viewer or not file_key or spec is None:
            return None
        header, _ = viewer.headers.get(file_key, (None, None))
        if header is None:
            return None
        if dims and len(dims) == 2:
            width = max(2, int(dims[0]))
            height = max(2, int(dims[1]))
        else:
            width = int(getattr(viewer, "thumb_size_px", 160))
            height = max(48, int(round(width * 0.75)))
        try:
            coords = viewer._map_spec_to_pixels(spec, header, width, height, file_key=file_key)
        except Exception:
            coords = None
        return coords

    def _collect_inset_markers_compare(self, base_key, image_dims=None):
        viewer = getattr(self, "viewer", None)
        if not viewer or not base_key:
            return []
        markers = []
        width = int(image_dims[0]) if image_dims else int(getattr(viewer, "thumb_size_px", 160))
        height = int(image_dims[1]) if image_dims else max(48, int(round(width * 0.75)))
        root = self.spec_list.invisibleRootItem()
        for i in range(root.childCount()):
            item = root.child(i)
            if item.isHidden():
                continue
            if item.checkState(0) != QtCore.Qt.Checked and not item.isSelected():
                continue
            spec = item.data(0, QtCore.Qt.UserRole)
            if not spec:
                continue
            key = str(spec.get("image_key") or "")
            if key != base_key:
                continue
            try:
                header, _ = viewer.headers.get(key, (None, None))
            except Exception:
                header = None
            try:
                coords = viewer._map_spec_to_pixels(spec, header, width, height, file_key=key)
            except Exception:
                coords = None
            if coords is None:
                continue
            spec_id = self._spec_id(spec)
            line = self._line_map.get(spec_id)
            color = line.get_color() if line else "#d65f5f"
            markers.append({"color": color, "coords": coords, "spec": spec})
        return markers

    def _update_position_inset_compare(self):
        if self._position_inset_ax is not None:
            try:
                self._position_inset_ax.remove()
            except Exception:
                pass
            self._position_inset_ax = None
        self._show_position_inset = bool(getattr(self, "position_inset_cb", None) and self.position_inset_cb.isChecked())
        if not self._show_position_inset:
            return
        items = self._checked_items() or self._selected_items()
        if not items:
            return
        base_spec = items[0].data(0, QtCore.Qt.UserRole)
        base_key = str(base_spec.get("image_key") or "") if base_spec else ""
        image = self._load_thumbnail_array_for_inset(base_key)
        image_dims = None
        if image is not None:
            try:
                image_dims = (int(image.shape[1]), int(image.shape[0]))
            except Exception:
                image_dims = None
        markers = self._collect_inset_markers_compare(base_key, image_dims=image_dims)
        if image is None or not markers:
            return
        if self._inset_bbox is None:
            self._inset_bbox = [0.04, 0.04, 0.28, 0.28]
        self._position_inset_ax = inset_axes(self.ax, width="26%", height="26%", loc="lower left", borderpad=0.8)
        self._position_inset_ax.set_axes_locator(InsetPosition(self.ax, self._inset_bbox))
        self._position_inset_ax.imshow(image, origin="upper")
        self._position_inset_ax.set_xticks([])
        self._position_inset_ax.set_yticks([])
        self._position_inset_ax.set_title("Position", fontsize=7.5 * getattr(self, "_font_scale", 1.0))
        raw_coords = []
        for marker in markers:
            color = marker.get("color")
            coords = marker.get("coords")
            spec = marker.get("spec")
            try:
                raw_coords.append((spec, float(coords[0]), float(coords[1])))
                self._position_inset_ax.scatter(
                    coords[0],
                    coords[1],
                    s=50,
                    facecolors="none",
                    edgecolors=color,
                    linewidths=1.5,
                )
            except Exception:
                continue
        for badge in spectro_overlays._stack_badges_from_coords(raw_coords):
            try:
                self._position_inset_ax.text(
                    float(badge.get("col")) + 6.0,
                    float(badge.get("row")) - 6.0,
                    str(badge.get("label") or ""),
                    fontsize=6.6 * getattr(self, "_font_scale", 1.0),
                    fontweight="bold",
                    color="#ffe478",
                    ha="left",
                    va="bottom",
                    bbox=dict(boxstyle="round,pad=0.18", facecolor="#281e12", edgecolor="#ffe0a0", linewidth=0.8, alpha=0.92),
                )
            except Exception:
                continue

    def _on_inset_press(self, event):
        if event is None or event.button != MouseButton.LEFT:
            return
        if self._position_inset_ax is None or not self._show_position_inset:
            return
        bbox = self._position_inset_ax.bbox
        if bbox is None:
            return
        if bbox.contains(event.x, event.y):
            self._inset_dragging = True
            self._inset_drag_offset = (event.x - bbox.x0, event.y - bbox.y0)

    def _on_inset_motion(self, event):
        if not self._inset_dragging or self._position_inset_ax is None:
            return
        if event.x is None or event.y is None:
            return
        bbox = self._position_inset_ax.bbox
        if bbox is None:
            return
        try:
            inv = self.ax.transAxes.inverted()
        except Exception:
            return
        ax_coords = inv.transform((event.x - self._inset_drag_offset[0], event.y - self._inset_drag_offset[1]))
        width = self._inset_bbox[2] if self._inset_bbox is not None else 0.26
        height = self._inset_bbox[3] if self._inset_bbox is not None else 0.26
        x0 = min(max(ax_coords[0], 0.0), 1.0 - width)
        y0 = min(max(ax_coords[1], 0.0), 1.0 - height)
        if self._inset_bbox is None:
            self._inset_bbox = [x0, y0, width, height]
        else:
            self._inset_bbox[0] = x0
            self._inset_bbox[1] = y0
        try:
            self._position_inset_ax.set_axes_locator(InsetPosition(self.ax, self._inset_bbox))
        except Exception:
            pass
        self.canvas.draw_idle()

    def _on_inset_release(self, event):
        if event is None or event.button != MouseButton.LEFT:
            return
        self._inset_dragging = False

    def _validate_log_axes(self):
        if not getattr(self, "_plot_x_log", False) and not getattr(self, "_plot_y_log", False):
            return
        invalid = False
        if self._plot_x_log and not self._entries_support_log_axis("x"):
            self._plot_x_log = False
            invalid = True
        if self._plot_y_log and not self._entries_support_log_axis("y"):
            self._plot_y_log = False
            invalid = True
        if invalid:
            QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), "Log axis disabled (non-positive values)", self)

    def _entries_support_log_axis(self, axis: str) -> bool:
        axis = (axis or "x").lower()
        items = self._checked_items() or self._selected_items()
        if not items:
            return False
        channel = self.channel_combo.currentText()
        for item in items:
            spec = item.data(0, QtCore.Qt.UserRole)
            if not spec:
                continue
            x_vals, _, _ = self._axis_for_spec(spec)
            channels = spec.get("channels") or {}
            y_vals = channels.get(channel)
            arr = x_vals if axis == "x" else y_vals
            if arr is None:
                return False
            vec = np.asarray(arr, dtype=float)
            vec = vec[np.isfinite(vec)]
            if not vec.size:
                return False
            if np.nanmin(vec) <= 0.0:
                return False
        return True

    def _set_plot_axis_log(self, axis: str, enabled: bool):
        axis = (axis or "x").lower()
        attr = "_plot_x_log" if axis == "x" else "_plot_y_log"
        if enabled:
            if not self._entries_support_log_axis(axis):
                QtWidgets.QMessageBox.information(
                    self,
                    "Log axis unavailable",
                    f"Cannot enable log scale on the {axis.upper()} axis because the plotted data contains zero or negative values.",
                )
                return
        setattr(self, attr, bool(enabled))
        self._request_plot_update(delay_ms=25)

    def _reset_plot_style(self):
        self._plot_grid_enabled = True
        self._plot_legend_enabled = True
        self._plot_x_log = False
        self._plot_y_log = False
        self._plot_line_width = 1.6
        self._set_visual_checkbox(self.show_points_cb, False)
        self._set_visual_checkbox(self.lines_cb, True)
        self._request_plot_update(delay_ms=25)

    def _set_visual_checkbox(self, checkbox, state):
        checkbox.blockSignals(True)
        checkbox.setChecked(bool(state))
        checkbox.blockSignals(False)
        self._on_visual_toggle(bool(state))

    def _bump_line_width(self, delta):
        self._plot_line_width = float(min(4.5, max(0.4, self._plot_line_width + delta)))
        self._request_plot_update(delay_ms=20)

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
        if self._apply_palette_to_existing_lines():
            return
        self._request_plot_update(delay_ms=25)

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
        stack_keys = set()
        for i in range(root.childCount()):
            if not root.child(i).isHidden():
                total += 1
                spec = root.child(i).data(0, QtCore.Qt.UserRole) or {}
                stack_key = spec.get("xy_stack_key")
                if stack_key and int(spec.get("xy_stack_count") or 0) > 1:
                    stack_keys.add(str(stack_key))
        checked = len(self._checked_items())
        text = f"{checked} selected / {total} total"
        if plotted is not None:
            text += f" | showing {plotted}"
        if stack_keys:
            text += f" | XY stacks {len(stack_keys)}"
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
        headers = [
            self.results_table.horizontalHeaderItem(c).text() if self.results_table.horizontalHeaderItem(c) else ""
            for c in range(self.results_table.columnCount())
        ]
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

    def _resolve_minima_overlaps(self):
        """Simple re-offset of minima labels to reduce overlap."""
        if not self._minima_meta:
            return
        ylim = self.ax.get_ylim()
        y_span = abs(ylim[1] - ylim[0]) if ylim and len(ylim) == 2 else 1.0
        for idx, meta in enumerate(self._minima_meta):
            txt = meta.get("text")
            if txt is None:
                continue
            sign = 1 if (idx % 2) == 0 else -1
            step = 1 + (idx // 2) * 0.6
            y_offset = sign * step * 0.04 * y_span
            pos = txt.get_position()
            txt.set_position((meta.get("x", pos[0]), pos[1] + y_offset))
        self.canvas.draw_idle()

    def _add_point_label_at_cursor(self):
        """Add a user point label at the last mouse position (optionally snapped to nearest curve)."""
        if self._last_mouse_xy is None:
            return
        x, y = self._last_mouse_xy
        snap_x, snap_y = self._snap_to_nearest_curve(x, y)
        x_use = snap_x if snap_x is not None else x
        y_use = snap_y if snap_y is not None else y
        marker = self.ax.scatter([x_use], [y_use], color="#444", s=24, zorder=7)
        vline = self.ax.axvline(x_use, color="#777", linestyle="--", linewidth=0.9, alpha=0.75, zorder=6)
        txt = self.ax.text(
            x_use, y_use, f"{x_use:.4g}, {y_use:.4g}",
            fontsize=7 * getattr(self, "_font_scale", 1.0),
            ha="left",
            va="bottom",
            bbox=dict(facecolor="white", edgecolor="#444", alpha=0.8, linewidth=0.6, boxstyle="round,pad=0.2"),
            zorder=7,
        )
        self._point_labels.append({"marker": marker, "text": txt, "vline": vline, "x": x_use, "y": y_use})
        self.canvas.draw_idle()

    def _clear_point_labels(self, redraw=True):
        for pl in getattr(self, "_point_labels", []):
            for art in pl.values():
                try:
                    art.remove()
                except Exception:
                    pass
        self._point_labels = []
        if redraw:
            self.canvas.draw_idle()

    def _on_point_label_press(self, event):
        if not event or event.inaxes != self.ax or event.button != MouseButton.LEFT:
            return
        if event.xdata is None or event.ydata is None:
            return
        for pl in getattr(self, "_point_labels", []):
            txt = pl.get("text")
            if txt is None:
                continue
            contains, _ = txt.contains(event)
            if contains:
                tx, ty = txt.get_position()
                self._point_label_drag = {"pl": pl, "offset": (tx - event.xdata, ty - event.ydata)}
                break

    def _on_point_label_motion(self, event):
        if not self._point_label_drag or event.inaxes != self.ax:
            return
        if event.xdata is None or event.ydata is None:
            return
        pl = self._point_label_drag.get("pl")
        txt = pl.get("text") if pl else None
        if txt is None:
            self._point_label_drag = None
            return
        dx, dy = self._point_label_drag.get("offset", (0.0, 0.0))
        txt.set_position((event.xdata + dx, event.ydata + dy))
        self.canvas.draw_idle()

    def _on_point_label_release(self, event):
        if self._point_label_drag and event and event.inaxes == self.ax:
            self.canvas.draw_idle()
        self._point_label_drag = None

    def _snap_to_nearest_curve(self, x, y):
        """Find nearest data point among plotted lines (in data space)."""
        nearest = None
        best_d2 = None
        for line in self._line_map.values():
            try:
                xs = line.get_xdata(orig=False)
                ys = line.get_ydata(orig=False)
                if xs is None or ys is None:
                    continue
                xs = np.asarray(xs, dtype=float)
                ys = np.asarray(ys, dtype=float)
                if xs.size == 0 or ys.size == 0:
                    continue
                d2 = (xs - x) ** 2 + (ys - y) ** 2
                idx = int(np.nanargmin(d2))
                val = d2[idx]
                if best_d2 is None or val < best_d2:
                    best_d2 = val
                    nearest = (float(xs[idx]), float(ys[idx]))
            except Exception:
                continue
        return nearest if nearest else (None, None)

    def _apply_data_filters(self, x_vals, y_vals, y_unit, x_unit):
        """Apply any enabled data filters/derivatives."""
        data = np.asarray(y_vals, dtype=float)
        x_arr = np.asarray(x_vals, dtype=float) if np.size(x_vals) == data.size else np.linspace(0, data.size - 1, data.size)
        if data.size == 0:
            return data, y_unit
        cfg = getattr(self, "_filter_cfg", {})
        result = data.copy()
        dx = float(np.nanmean(np.diff(x_arr))) if x_arr.size > 1 else 1.0
        if not math.isfinite(dx) or dx == 0:
            dx = 1.0

        def _odd(value, minimum):
            v = max(minimum, int(value) or minimum)
            return v + 1 if v % 2 == 0 else v

        # Gaussian smoothing
        gauss = cfg.get("gaussian", {})
        if gauss.get("enabled"):
            sigma = max(0.1, float(gauss.get("sigma", 1.0)))
            if _scipy_ndimage is not None:
                result = _scipy_ndimage.gaussian_filter1d(result, sigma=max(0.05, sigma), mode="nearest")
            else:
                radius = max(1, int(3 * sigma))
                xs = np.arange(-radius, radius + 1)
                kernel = np.exp(-(xs ** 2) / (2.0 * sigma ** 2))
                kernel /= kernel.sum() or 1.0
                result = np.convolve(result, kernel, mode="same")

        # Median filtering
        median = cfg.get("median", {})
        if median.get("enabled"):
            size = _odd(median.get("size", 3), 3)
            if _scipy_ndimage is not None:
                result = _scipy_ndimage.median_filter(result, size=size, mode="nearest")
            else:
                pad = size // 2
                padded = np.pad(result, pad, mode="edge")
                out = np.empty_like(result)
                for i in range(result.size):
                    window = padded[i:i + size]
                    out[i] = np.median(window)
                result = out

        # Savitzky-Golay smoothing
        sav_cfg = cfg.get("savgol", {})
        if sav_cfg.get("enabled"):
            window = _odd(sav_cfg.get("window", 11), 5)
            window = min(window, result.size - 1 if result.size % 2 == 0 else result.size)
            window = max(5, window if window % 2 == 1 else window - 1)
            poly = max(2, min(int(sav_cfg.get("poly", 3)), window - 1))
            if window >= 3 and window <= result.size:
                if _scipy_signal is not None:
                    result = _scipy_signal.savgol_filter(result, window, poly, mode="interp")
                else:
                    kernel = np.ones(window) / float(window)
                    result = np.convolve(result, kernel, mode="same")

        # FFT low-pass
        fft_cfg = cfg.get("fft", {})
        if fft_cfg.get("enabled") and result.size >= 8:
            cutoff = float(fft_cfg.get("cutoff", 0.15))
            cutoff = min(max(cutoff, 0.0), 0.5)
            if cutoff > 0.0:
                centered = result - np.nanmean(result)
                freq = np.fft.rfftfreq(result.size, d=dx)
                spectrum = np.fft.rfft(centered)
                nyquist = 0.5 / dx
                thresh = cutoff * nyquist
                mask = np.abs(freq) <= thresh
                spectrum *= mask
                recovered = np.fft.irfft(spectrum, n=result.size)
                result = recovered + np.nanmean(result)

        # Notch filter
        notch = cfg.get("notch", {})
        if notch.get("enabled") and result.size >= 8:
            freq = abs(float(notch.get("freq", 50.0)))
            width = max(0.0001, abs(float(notch.get("width", 5.0))))
            if freq > 0.0:
                centered = result - np.nanmean(result)
                spectrum = np.fft.rfft(centered)
                freqs = np.fft.rfftfreq(result.size, d=dx)
                mask = np.ones_like(freqs, dtype=bool)
                notch_region = np.abs(freqs - freq) < width
                mask[notch_region] = False
                spectrum *= mask
                recovered = np.fft.irfft(spectrum, n=result.size)
                result = recovered + np.nanmean(result)

        # Derivative (dY/dX)
        deriv = cfg.get("derive", {})
        unit = y_unit
        if deriv.get("enabled"):
            window = _odd(deriv.get("window", 11), 5)
            window = min(window, result.size - 1 if result.size % 2 == 0 else result.size)
            window = max(5, window if window % 2 == 1 else window - 1)
            poly = max(2, min(int(deriv.get("poly", 3)), window - 1))
            if _scipy_signal is not None and window >= 5 and window <= result.size:
                result = _scipy_signal.savgol_filter(result, window, poly, deriv=1, delta=dx, mode="interp")
            else:
                result = np.gradient(result, x_arr)
            denom = x_unit or "x"
            unit = f"d({unit or 'arb'})/d({denom})"

        return result, unit

    def _register_filter_control(self, section, key, widget):
        self._filter_controls.setdefault(section, {})[key] = widget

    def _sync_filter_controls(self, section):
        refs = self._filter_controls.get(section)
        if not refs:
            return
        cfg = self._filter_cfg.get(section, {})
        for key, widget in refs.items():
            if widget is None:
                continue
            blocker = getattr(widget, "blockSignals", None)
            if callable(blocker):
                blocker(True)
            try:
                if key == "enabled" and isinstance(widget, QtWidgets.QAbstractButton):
                    widget.setChecked(bool(cfg.get("enabled")))
                elif hasattr(widget, "setValue"):
                    value = cfg.get(key)
                    if value is not None:
                        widget.setValue(value)
            finally:
                if callable(blocker):
                    blocker(False)

    def _set_filter_enabled(self, section, enabled):
        self._filter_cfg.setdefault(section, {})["enabled"] = bool(enabled)
        self._sync_filter_controls(section)
        self._request_plot_update(delay_ms=50)

    def _set_filter_value(self, section, key, value):
        cfg = self._filter_cfg.setdefault(section, {})
        cfg[key] = value
        self._sync_filter_controls(section)
        self._request_plot_update(delay_ms=60)

    def _build_filter_menu(self, menu):
        cfg = self._filter_cfg
        def widget_action(widget):
            act = QtWidgets.QWidgetAction(menu)
            act.setDefaultWidget(widget)
            menu.addAction(act)
        # Gaussian
        g_row = QtWidgets.QWidget()
        g_layout = QtWidgets.QHBoxLayout(g_row); g_layout.setContentsMargins(6,2,6,2); g_layout.setSpacing(6)
        g_cb = QtWidgets.QCheckBox("Gaussian σ")
        g_cb.setChecked(cfg.get("gaussian", {}).get("enabled", False))
        g_spin = QtWidgets.QDoubleSpinBox(); g_spin.setRange(0.1, 10.0); g_spin.setSingleStep(0.1)
        g_spin.setValue(float(cfg.get("gaussian", {}).get("sigma", 1.0)))
        g_cb.toggled.connect(lambda chk: self._set_filter_enabled("gaussian", chk))
        g_spin.valueChanged.connect(lambda val: self._set_filter_value("gaussian", "sigma", float(val)))
        g_layout.addWidget(g_cb); g_layout.addWidget(g_spin)
        widget_action(g_row)
        # Savitzky-Golay smoothing
        sg_row = QtWidgets.QWidget(); sg_layout = QtWidgets.QHBoxLayout(sg_row); sg_layout.setContentsMargins(6,2,6,2); sg_layout.setSpacing(6)
        sg_cb = QtWidgets.QCheckBox("Savitzky-Golay")
        sg_cb.setChecked(cfg.get("savgol", {}).get("enabled", False))
        sg_win = QtWidgets.QSpinBox(); sg_win.setRange(5, 201); sg_win.setSingleStep(2); sg_win.setValue(int(cfg.get("savgol", {}).get("window", 11)))
        sg_poly = QtWidgets.QSpinBox(); sg_poly.setRange(2, 10); sg_poly.setValue(int(cfg.get("savgol", {}).get("poly", 3)))
        sg_cb.toggled.connect(lambda chk: self._set_filter_enabled("savgol", chk))
        sg_win.valueChanged.connect(lambda val: self._set_filter_value("savgol", "window", int(val)))
        sg_poly.valueChanged.connect(lambda val: self._set_filter_value("savgol", "poly", int(val)))
        sg_layout.addWidget(sg_cb); sg_layout.addWidget(QtWidgets.QLabel("Window")); sg_layout.addWidget(sg_win)
        sg_layout.addWidget(QtWidgets.QLabel("Poly")); sg_layout.addWidget(sg_poly)
        widget_action(sg_row)
        # Median
        med_row = QtWidgets.QWidget(); med_layout = QtWidgets.QHBoxLayout(med_row); med_layout.setContentsMargins(6,2,6,2); med_layout.setSpacing(6)
        med_cb = QtWidgets.QCheckBox("Median")
        med_cb.setChecked(cfg.get("median", {}).get("enabled", False))
        med_spin = QtWidgets.QSpinBox(); med_spin.setRange(3, 51); med_spin.setSingleStep(2); med_spin.setValue(int(cfg.get("median", {}).get("size", 3)))
        med_cb.toggled.connect(lambda chk: self._set_filter_enabled("median", chk))
        med_spin.valueChanged.connect(lambda val: self._set_filter_value("median", "size", int(val)))
        med_layout.addWidget(med_cb); med_layout.addWidget(QtWidgets.QLabel("Size")); med_layout.addWidget(med_spin)
        widget_action(med_row)
        # FFT low-pass
        fft_row = QtWidgets.QWidget(); fft_layout = QtWidgets.QHBoxLayout(fft_row); fft_layout.setContentsMargins(6,2,6,2); fft_layout.setSpacing(6)
        fft_cb = QtWidgets.QCheckBox("FFT low-pass")
        fft_cb.setChecked(cfg.get("fft", {}).get("enabled", False))
        fft_cut = QtWidgets.QDoubleSpinBox(); fft_cut.setRange(0.01, 0.5); fft_cut.setSingleStep(0.01); fft_cut.setDecimals(3)
        fft_cut.setValue(float(cfg.get("fft", {}).get("cutoff", 0.15)))
        fft_cb.toggled.connect(lambda chk: self._set_filter_enabled("fft", chk))
        fft_cut.valueChanged.connect(lambda val: self._set_filter_value("fft", "cutoff", float(val)))
        fft_layout.addWidget(fft_cb); fft_layout.addWidget(QtWidgets.QLabel("Cutoff (Nyquist frac)")); fft_layout.addWidget(fft_cut)
        widget_action(fft_row)
        # Notch
        notch_row = QtWidgets.QWidget(); notch_layout = QtWidgets.QHBoxLayout(notch_row); notch_layout.setContentsMargins(6,2,6,2); notch_layout.setSpacing(6)
        notch_cb = QtWidgets.QCheckBox("Notch")
        notch_cb.setChecked(cfg.get("notch", {}).get("enabled", False))
        notch_freq = QtWidgets.QDoubleSpinBox(); notch_freq.setRange(0.1, 5000.0); notch_freq.setSingleStep(1.0); notch_freq.setDecimals(3)
        notch_freq.setValue(float(cfg.get("notch", {}).get("freq", 50.0)))
        notch_width = QtWidgets.QDoubleSpinBox(); notch_width.setRange(0.001, 500.0); notch_width.setSingleStep(0.5); notch_width.setDecimals(3)
        notch_width.setValue(float(cfg.get("notch", {}).get("width", 5.0)))
        notch_cb.toggled.connect(lambda chk: self._set_filter_enabled("notch", chk))
        notch_freq.valueChanged.connect(lambda val: self._set_filter_value("notch", "freq", float(val)))
        notch_width.valueChanged.connect(lambda val: self._set_filter_value("notch", "width", float(val)))
        notch_layout.addWidget(notch_cb); notch_layout.addWidget(QtWidgets.QLabel("Freq")); notch_layout.addWidget(notch_freq)
        notch_layout.addWidget(QtWidgets.QLabel("Width")); notch_layout.addWidget(notch_width)
        widget_action(notch_row)
        # Derivative
        deriv_row = QtWidgets.QWidget(); deriv_layout = QtWidgets.QHBoxLayout(deriv_row); deriv_layout.setContentsMargins(6,2,6,2); deriv_layout.setSpacing(6)
        deriv_cb = QtWidgets.QCheckBox("dY/dX (SG)")
        deriv_cb.setChecked(cfg.get("derive", {}).get("enabled", False))
        deriv_win = QtWidgets.QSpinBox(); deriv_win.setRange(5, 201); deriv_win.setSingleStep(2); deriv_win.setValue(int(cfg.get("derive", {}).get("window", 11)))
        deriv_poly = QtWidgets.QSpinBox(); deriv_poly.setRange(2, 10); deriv_poly.setValue(int(cfg.get("derive", {}).get("poly", 3)))
        deriv_cb.toggled.connect(lambda chk: self._set_filter_enabled("derive", chk))
        deriv_win.valueChanged.connect(lambda val: self._set_filter_value("derive", "window", int(val)))
        deriv_poly.valueChanged.connect(lambda val: self._set_filter_value("derive", "poly", int(val)))
        deriv_layout.addWidget(deriv_cb); deriv_layout.addWidget(QtWidgets.QLabel("Window")); deriv_layout.addWidget(deriv_win)
        deriv_layout.addWidget(QtWidgets.QLabel("Poly")); deriv_layout.addWidget(deriv_poly)
        widget_action(deriv_row)
        reset_btn = QtWidgets.QPushButton("Disable all filters")
        reset_btn.clicked.connect(lambda _=None: self._reset_filters())
        widget_action(reset_btn)

    def _reset_filters(self):
        for name, section in self._filter_cfg.items():
            section["enabled"] = False
            self._sync_filter_controls(name)
        self._request_plot_update(delay_ms=60)

    def _build_filter_panel(self):
        group = QtWidgets.QGroupBox("Filters")
        layout = QtWidgets.QVBoxLayout(group)
        cfg = self._filter_cfg

        def add_row(widget):
            layout.addWidget(widget)

        def make_checkbox(label, section, tooltip):
            cb = QtWidgets.QCheckBox(label)
            cb.setToolTip(tooltip)
            cb.setChecked(cfg.get(section, {}).get("enabled", False))
            cb.toggled.connect(lambda chk, sec=section: self._set_filter_enabled(sec, chk))
            self._register_filter_control(section, "enabled", cb)
            return cb

        # Gaussian
        g_widget = QtWidgets.QWidget()
        g_layout = QtWidgets.QHBoxLayout(g_widget); g_layout.setContentsMargins(0,0,0,0); g_layout.setSpacing(6)
        g_cb = make_checkbox("Gaussian", "gaussian", "Apply Gaussian smoothing (σ controls blur)")
        g_spin = QtWidgets.QDoubleSpinBox()
        g_spin.setRange(0.1, 10.0); g_spin.setSingleStep(0.1); g_spin.setDecimals(2)
        g_spin.setToolTip("Gaussian σ (points)")
        g_spin.setValue(float(cfg.get("gaussian", {}).get("sigma", 1.0)))
        g_spin.valueChanged.connect(lambda val: self._set_filter_value("gaussian", "sigma", float(val)))
        self._register_filter_control("gaussian", "sigma", g_spin)
        g_layout.addWidget(g_cb)
        g_layout.addWidget(QtWidgets.QLabel("σ:"))
        g_layout.addWidget(g_spin, 1)
        add_row(g_widget)

        # Savitzky-Golay
        sg_widget = QtWidgets.QWidget()
        sg_layout = QtWidgets.QHBoxLayout(sg_widget); sg_layout.setContentsMargins(0,0,0,0); sg_layout.setSpacing(6)
        sg_cb = make_checkbox("Savitzky-Golay", "savgol", "Polynomial smoothing filter")
        sg_win = QtWidgets.QSpinBox(); sg_win.setRange(5, 201); sg_win.setSingleStep(2)
        sg_win.setValue(int(cfg.get("savgol", {}).get("window", 11)))
        sg_win.setToolTip("Window length (odd)")
        sg_poly = QtWidgets.QSpinBox(); sg_poly.setRange(2, 10); sg_poly.setValue(int(cfg.get("savgol", {}).get("poly", 3)))
        sg_poly.setToolTip("Polynomial order")
        sg_win.valueChanged.connect(lambda val: self._set_filter_value("savgol", "window", int(val)))
        sg_poly.valueChanged.connect(lambda val: self._set_filter_value("savgol", "poly", int(val)))
        self._register_filter_control("savgol", "window", sg_win)
        self._register_filter_control("savgol", "poly", sg_poly)
        sg_layout.addWidget(sg_cb)
        sg_layout.addWidget(QtWidgets.QLabel("Window"))
        sg_layout.addWidget(sg_win)
        sg_layout.addWidget(QtWidgets.QLabel("Poly"))
        sg_layout.addWidget(sg_poly)
        add_row(sg_widget)

        # Median
        med_widget = QtWidgets.QWidget()
        med_layout = QtWidgets.QHBoxLayout(med_widget); med_layout.setContentsMargins(0,0,0,0); med_layout.setSpacing(6)
        med_cb = make_checkbox("Median", "median", "Median filter (spike removal)")
        med_spin = QtWidgets.QSpinBox(); med_spin.setRange(3, 51); med_spin.setSingleStep(2)
        med_spin.setValue(int(cfg.get("median", {}).get("size", 3)))
        med_spin.setToolTip("Window size (odd)")
        med_spin.valueChanged.connect(lambda val: self._set_filter_value("median", "size", int(val)))
        self._register_filter_control("median", "size", med_spin)
        med_layout.addWidget(med_cb)
        med_layout.addWidget(QtWidgets.QLabel("Size"))
        med_layout.addWidget(med_spin)
        add_row(med_widget)

        # FFT low-pass
        fft_widget = QtWidgets.QWidget()
        fft_layout = QtWidgets.QHBoxLayout(fft_widget); fft_layout.setContentsMargins(0,0,0,0); fft_layout.setSpacing(6)
        fft_cb = make_checkbox("FFT low-pass", "fft", "Low-pass frequency filtering (fraction of Nyquist)")
        fft_spin = QtWidgets.QDoubleSpinBox(); fft_spin.setRange(0.01, 0.5); fft_spin.setSingleStep(0.01); fft_spin.setDecimals(3)
        fft_spin.setValue(float(cfg.get("fft", {}).get("cutoff", 0.15)))
        fft_spin.setToolTip("Cutoff (0-0.5 of Nyquist)")
        fft_spin.valueChanged.connect(lambda val: self._set_filter_value("fft", "cutoff", float(val)))
        self._register_filter_control("fft", "cutoff", fft_spin)
        fft_layout.addWidget(fft_cb)
        fft_layout.addWidget(QtWidgets.QLabel("Cutoff"))
        fft_layout.addWidget(fft_spin)
        add_row(fft_widget)

        # Notch
        notch_widget = QtWidgets.QWidget()
        notch_layout = QtWidgets.QHBoxLayout(notch_widget); notch_layout.setContentsMargins(0,0,0,0); notch_layout.setSpacing(6)
        notch_cb = make_checkbox("Notch", "notch", "Remove a narrow frequency band (e.g., mains noise)")
        notch_freq = QtWidgets.QDoubleSpinBox(); notch_freq.setRange(0.1, 5000.0); notch_freq.setDecimals(2); notch_freq.setSingleStep(1.0)
        notch_freq.setValue(float(cfg.get("notch", {}).get("freq", 50.0)))
        notch_freq.setToolTip("Notch frequency (Hz or axis units)")
        notch_width = QtWidgets.QDoubleSpinBox(); notch_width.setRange(0.001, 500.0); notch_width.setDecimals(3); notch_width.setSingleStep(0.5)
        notch_width.setValue(float(cfg.get("notch", {}).get("width", 5.0)))
        notch_width.setToolTip("Notch width")
        notch_freq.valueChanged.connect(lambda val: self._set_filter_value("notch", "freq", float(val)))
        notch_width.valueChanged.connect(lambda val: self._set_filter_value("notch", "width", float(val)))
        self._register_filter_control("notch", "freq", notch_freq)
        self._register_filter_control("notch", "width", notch_width)
        notch_layout.addWidget(notch_cb)
        notch_layout.addWidget(QtWidgets.QLabel("Freq"))
        notch_layout.addWidget(notch_freq)
        notch_layout.addWidget(QtWidgets.QLabel("Width"))
        notch_layout.addWidget(notch_width)
        add_row(notch_widget)

        # Derivative
        deriv_widget = QtWidgets.QWidget()
        deriv_layout = QtWidgets.QHBoxLayout(deriv_widget); deriv_layout.setContentsMargins(0,0,0,0); deriv_layout.setSpacing(6)
        deriv_cb = make_checkbox("Derivative (dY/dX)", "derive", "Numerical derivative using Savitzky-Golay")
        deriv_win = QtWidgets.QSpinBox(); deriv_win.setRange(5, 201); deriv_win.setSingleStep(2)
        deriv_win.setValue(int(cfg.get("derive", {}).get("window", 11)))
        deriv_win.setToolTip("Window length (odd)")
        deriv_poly = QtWidgets.QSpinBox(); deriv_poly.setRange(2, 10); deriv_poly.setValue(int(cfg.get("derive", {}).get("poly", 3)))
        deriv_poly.setToolTip("Polynomial order")
        deriv_win.valueChanged.connect(lambda val: self._set_filter_value("derive", "window", int(val)))
        deriv_poly.valueChanged.connect(lambda val: self._set_filter_value("derive", "poly", int(val)))
        self._register_filter_control("derive", "window", deriv_win)
        self._register_filter_control("derive", "poly", deriv_poly)
        deriv_layout.addWidget(deriv_cb)
        deriv_layout.addWidget(QtWidgets.QLabel("Window"))
        deriv_layout.addWidget(deriv_win)
        deriv_layout.addWidget(QtWidgets.QLabel("Poly"))
        deriv_layout.addWidget(deriv_poly)
        add_row(deriv_widget)

        reset_btn = QtWidgets.QPushButton("Disable all filters")
        reset_btn.clicked.connect(self._reset_filters)
        layout.addWidget(reset_btn)
        return group

    def _toggle_position_inset_from_menu(self, checked):
        checked = bool(checked)
        if hasattr(self, "position_inset_cb") and self.position_inset_cb:
            block = self.position_inset_cb.blockSignals
            try:
                block(True)
                self.position_inset_cb.setChecked(checked)
            finally:
                block(False)
        self._show_position_inset = checked
        self._update_position_inset_compare()
    def _gather_plotted_traces(self):
        traces = []
        channel = self.channel_combo.currentText()
        relative_nm = self.relative_cb.isChecked()
        waterfall = self.waterfall_cb.isChecked()
        offset_val = self.offset_spin.value()
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
        plotted = 0
        root = self.spec_list.invisibleRootItem()
        for i in range(root.childCount()):
            item = root.child(i)
            if item.isHidden():
                continue
            if item.checkState(0) != QtCore.Qt.Checked and not item.isSelected():
                continue
            spec = item.data(0, QtCore.Qt.UserRole)
            if not spec:
                continue
            spec_id = item.data(0, QtCore.Qt.UserRole + 1)
            channels = spec.get("channels") or {}
            data = channels.get(channel)
            axis_vals, _, axis_unit = self._axis_for_spec(spec)
            if data is None or not axis_vals.size:
                continue
            bg_spec = self._background_for(spec)
            y_base = self._subtract_background(axis_vals, data, bg_spec)
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
            y_unit = self._channel_unit_for_spec(spec, channel)
            y_filtered, y_unit_final = self._apply_data_filters(x_vals, y_data, y_unit, axis_unit_plot)
            traces.append({
                "label": self._display_name(spec),
                "path": spec.get("path"),
                "matrix_index": spec.get("matrix_index"),
                "channel": channel,
                "x_unit": axis_unit_plot,
                "y_unit": y_unit_final,
                "x_vals": np.asarray(x_vals, dtype=float),
                "y_vals": np.asarray(y_filtered, dtype=float),
                "spec_id": spec_id,
                "time": spec.get("time"),
                "pos_x": spec.get("x"),
                "pos_y": spec.get("y"),
            })
            plotted += 1
        return traces

    def _copy_all_traces_to_clipboard(self):
        traces = self._gather_plotted_traces()
        if not traces:
            QtWidgets.QMessageBox.information(self, "Copy spectra", "No spectra to copy.")
            return

        # Build human-friendly table: per-trace paired columns
        name_row = []
        pos_row = []
        unit_row = []
        max_len = 0
        for trace in traces:
            label = trace.get("label") or Path(trace.get("path") or "").name or "trace"
            acq_raw = trace.get("time")
            acq = "" if acq_raw is None else str(acq_raw)
            name_row += [label, acq]
            px = trace.get("pos_x")
            py = trace.get("pos_y")
            pos_row += [
                "" if px is None else f"{float(px):.4g} nm",
                "" if py is None else f"{float(py):.4g} nm",
            ]
            unit_row += [trace.get("x_unit") or "", trace.get("y_unit") or ""]
            x_raw = trace.get("x_vals")
            y_raw = trace.get("y_vals")
            x_vals = x_raw if x_raw is not None else []
            y_vals = y_raw if y_raw is not None else []
            max_len = max(max_len, len(x_vals), len(y_vals))
            trace["_x_arr"] = x_vals
            trace["_y_arr"] = y_vals

        rows = []
        rows.append("\t".join(name_row))
        rows.append("\t".join(pos_row))
        rows.append("\t".join(unit_row))
        for i in range(max_len):
            line_parts = []
            for trace in traces:
                x_vals = trace.get("_x_arr", [])
                y_vals = trace.get("_y_arr", [])
                x_val = "" if i >= len(x_vals) else x_vals[i]
                y_val = "" if i >= len(y_vals) else y_vals[i]
                try:
                    line_parts.append("" if x_val == "" else f"{float(x_val):.9g}")
                except Exception:
                    line_parts.append(str(x_val))
                try:
                    line_parts.append("" if y_val == "" else f"{float(y_val):.9g}")
                except Exception:
                    line_parts.append(str(y_val))
            rows.append("\t".join(line_parts))

        for trace in traces:
            file_name = Path(trace.get("path") or "").name
            point = trace.get("matrix_index")
            point_txt = "" if point is None else str(point)
            channel = trace.get("channel") or ""
            x_unit = trace.get("x_unit") or ""
            y_unit = trace.get("y_unit") or ""
            label = trace.get("label") or file_name or "trace"
        QtWidgets.QApplication.clipboard().setText("\n".join(rows))
        QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), "Copied all traces", self)

    def _clear_minima_annotations(self):
        """Remove previously drawn minima markers/labels."""
        if not getattr(self, "_minima_artists", None):
            self._minima_artists = []
            return
        for art in self._minima_artists:
            try:
                art.remove()
            except Exception:
                pass
        self._minima_artists = []
        self._minima_meta = []

    def _annotate_minima(self):
        """Find and mark the x-position of the minimum for each plotted trace."""
        traces = self._gather_plotted_traces()
        self._clear_minima_annotations()
        if not traces:
            QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), "No spectra plotted", self)
            self.canvas.draw_idle()
            return
        artists = []
        ylim = self.ax.get_ylim()
        y_span = abs(ylim[1] - ylim[0]) if ylim and len(ylim) == 2 else 1.0
        for t_idx, trace in enumerate(traces):
            x_raw = trace.get("x_vals")
            y_raw = trace.get("y_vals")
            x_vals = np.asarray([] if x_raw is None else x_raw, dtype=float)
            y_vals = np.asarray([] if y_raw is None else y_raw, dtype=float)
            if x_vals.size == 0 or y_vals.size == 0:
                continue
            idx = np.nanargmin(y_vals)
            x_min = x_vals[idx]
            y_min = y_vals[idx]
            spec_id = trace.get("spec_id")
            line = self._line_map.get(spec_id)
            color = line.get_color() if line else "#d65f5f"
            lbl = trace.get("label") or "trace"
            vline = self.ax.axvline(x_min, color=color, linestyle="--", linewidth=1.2, alpha=0.85)
            marker = self.ax.scatter([x_min], [y_min], color=color, s=26, zorder=6)
            artists.extend([vline, marker])
            # Vertical offset to reduce overlap; alternate above/below and increase with index
            sign = 1 if (t_idx % 2) == 0 else -1
            step = 1 + (t_idx // 2) * 0.6
            y_offset = sign * step * 0.04 * y_span
            try:
                txt = self.ax.text(
                    x_min, y_min + y_offset, f"{lbl}\n{float(x_min):.4g} {trace.get('x_unit','')}",
                    fontsize=7 * getattr(self, "_font_scale", 1.0),
                    color=color,
                    ha="center",
                    va="bottom",
                    bbox=dict(facecolor="white", edgecolor=color, alpha=0.85, linewidth=0.6, boxstyle="round,pad=0.2"),
                    picker=True,
                )
                artists.append(txt)
                self._minima_meta.append({"vline": vline, "marker": marker, "text": txt, "x": x_min, "color": color})
            except Exception:
                pass
        self._minima_artists = artists
        self.canvas.draw_idle()

    def _on_compare_canvas_menu(self, pos):
        menu = QtWidgets.QMenu(self)
        preset_menu = menu.addMenu("Figure preset")
        preset_actions = {}
        for preset in iter_figure_layout_presets():
            act = preset_menu.addAction(preset.label)
            act.setCheckable(True)
            act.setChecked(self._figure_preset_key == preset.key)
            preset_actions[act] = preset.key
        menu.addSeparator()
        copy_png = menu.addAction("Copy plot as PNG (300 dpi)")
        copy_png_600 = menu.addAction("Copy plot as PNG (600 dpi)")
        copy_svg = menu.addAction("Copy plot as SVG")
        copy_all = menu.addAction("Copy all traces (table)")
        save_menu = menu.addMenu("Save plot")
        save_png_300 = save_menu.addAction("PNG 300 dpi...")
        save_png_600 = save_menu.addAction("PNG 600 dpi...")
        save_svg = save_menu.addAction("SVG (vector)...")
        save_pdf = save_menu.addAction("PDF (vector)...")
        menu.addSeparator()
        add_font_menu_action(
            menu,
            self,
            self._plot_font_family,
            self.set_plot_font_family,
            current_style=self._font_style_state(),
            apply_style_callback=self.set_plot_typography,
        )
        menu.addSeparator()
        minima_act = menu.addAction("Find minima (x-position)")
        resolve_act = menu.addAction("Resolve minima overlaps")
        menu.addSeparator()
        add_point_act = menu.addAction("Add point label here")
        clear_points_act = menu.addAction("Clear point labels")
        inset_act = menu.addAction("Show position inset")
        inset_act.setCheckable(True)
        inset_act.setChecked(self._show_position_inset)
        inset_act.toggled.connect(self._toggle_position_inset_from_menu)
        menu.addSeparator()
        # Lines submenu (per-curve controls)
        lines_menu = menu.addMenu("Lines")
        ls_labels = {"Solid": "-", "Dashed": "--", "Dotted": ":", "Dash-dot": "-."}
        for spec_id, line in self._line_map.items():
            row = QtWidgets.QWidget()
            h = QtWidgets.QHBoxLayout(row); h.setContentsMargins(6, 2, 6, 2); h.setSpacing(6)
            name_lbl = QtWidgets.QLabel(self._curve_name(spec_id)); name_lbl.setMinimumWidth(90)
            lw_spin = QtWidgets.QDoubleSpinBox(); lw_spin.setRange(0.5, 5.0); lw_spin.setSingleStep(0.5)
            lw_spin.setValue(float(line.get_linewidth() or 1.0))
            ls_combo = QtWidgets.QComboBox(); [ls_combo.addItem(k, v) for k, v in ls_labels.items()]
            current_ls = line.get_linestyle() or "-"
            idx = max(0, ls_combo.findData(current_ls))
            ls_combo.setCurrentIndex(idx)
            col_btn = QtWidgets.QPushButton(); col_btn.setFixedWidth(36)
            def _set_btn_color(btn, c):
                btn.setStyleSheet(f"background:{c};")
            color = line.get_color() or "#000"
            _set_btn_color(col_btn, color)
            col_btn.clicked.connect(lambda _=None, sid=spec_id, btn=col_btn: self._pick_curve_color(sid, btn))
            lw_spin.valueChanged.connect(lambda val, sid=spec_id: self._set_curve_style(sid, lw=float(val)))
            ls_combo.currentIndexChanged.connect(lambda _i, sid=spec_id, cb=ls_combo: self._set_curve_style(sid, ls=cb.currentData()))
            h.addWidget(name_lbl, 1)
            h.addWidget(QtWidgets.QLabel("Thick"), 0)
            h.addWidget(lw_spin, 0)
            h.addWidget(QtWidgets.QLabel("Style"), 0)
            h.addWidget(ls_combo, 0)
            h.addWidget(col_btn, 0)
            act = QtWidgets.QWidgetAction(lines_menu); act.setDefaultWidget(row)
            lines_menu.addAction(act)
        lines_menu.addSeparator()
        # Global apply
        global_row = QtWidgets.QWidget()
        gh = QtWidgets.QHBoxLayout(global_row); gh.setContentsMargins(6, 4, 6, 4); gh.setSpacing(6)
        gh.addWidget(QtWidgets.QLabel("All:"), 0)
        all_lw = QtWidgets.QDoubleSpinBox(); all_lw.setRange(0.5, 5.0); all_lw.setSingleStep(0.5); all_lw.setValue(self._plot_line_width)
        all_ls = QtWidgets.QComboBox(); [all_ls.addItem(k, v) for k, v in ls_labels.items()]
        all_ls.setCurrentIndex(0)
        apply_all_btn = QtWidgets.QPushButton("Apply")
        apply_all_btn.clicked.connect(lambda _=None: self._apply_global_line_style(all_lw.value(), all_ls.currentData()))
        reset_cycle_act = QtWidgets.QPushButton("Reset colors to cycle")
        reset_cycle_act.clicked.connect(lambda _=None: self._reset_colors_to_cycle())
        gh.addWidget(QtWidgets.QLabel("Thickness"))
        gh.addWidget(all_lw)
        gh.addWidget(QtWidgets.QLabel("Style"))
        gh.addWidget(all_ls)
        gh.addWidget(apply_all_btn)
        gh.addWidget(reset_cycle_act)
        g_act = QtWidgets.QWidgetAction(lines_menu); g_act.setDefaultWidget(global_row)
        lines_menu.addAction(g_act)
        menu.addSeparator()
        filters_menu = menu.addMenu("Filters")
        self._build_filter_menu(filters_menu)
        menu.addSeparator()
        # Legend submenu
        legend_menu = menu.addMenu("Legend")
        legend_show_act = QtWidgets.QAction("Show legend", legend_menu, checkable=True, checked=self._plot_legend_enabled)
        legend_menu.addAction(legend_show_act)
        pos_combo = QtWidgets.QComboBox(); pos_combo.addItems(["Best", "Top-left", "Top-right", "Bottom-left", "Bottom-right"])
        pos_map = {"Best": "best", "Top-left": "upper left", "Top-right": "upper right", "Bottom-left": "lower left", "Bottom-right": "lower right"}
        pos_combo.setCurrentIndex(max(0, pos_combo.findText({
            "best": "Best",
            "upper left": "Top-left",
            "upper right": "Top-right",
            "lower left": "Bottom-left",
            "lower right": "Bottom-right",
        }.get(self._legend_loc, "Best"))))
        pos_widget = QtWidgets.QWidget(); pos_h = QtWidgets.QHBoxLayout(pos_widget); pos_h.setContentsMargins(6,2,6,2); pos_h.addWidget(QtWidgets.QLabel("Position")); pos_h.addWidget(pos_combo,1)
        pos_act = QtWidgets.QWidgetAction(legend_menu); pos_act.setDefaultWidget(pos_widget); legend_menu.addAction(pos_act)
        font_widget = QtWidgets.QWidget(); fw_h = QtWidgets.QHBoxLayout(font_widget); fw_h.setContentsMargins(6,2,6,2)
        font_spin = QtWidgets.QSpinBox(); font_spin.setRange(6, 18); font_spin.setValue(int(self._legend_font))
        fw_h.addWidget(QtWidgets.QLabel("Font size")); fw_h.addWidget(font_spin)
        font_act = QtWidgets.QWidgetAction(legend_menu); font_act.setDefaultWidget(font_widget); legend_menu.addAction(font_act)
        bg_act = QtWidgets.QAction("Background", legend_menu, checkable=True, checked=self._legend_bg)
        border_act = QtWidgets.QAction("Border", legend_menu, checkable=True, checked=self._legend_border)
        fname_act = QtWidgets.QAction("Use filename only", legend_menu, checkable=True, checked=self._legend_filename_only)
        legend_menu.addActions([bg_act, border_act, fname_act])
        menu.addSeparator()
        # Grid & ticks submenu
        grid_menu = menu.addMenu("Grid / ticks")
        grid_major_cb = QtWidgets.QCheckBox("Show major grid"); grid_major_cb.setChecked(self._grid_major)
        grid_minor_cb = QtWidgets.QCheckBox("Show minor grid"); grid_minor_cb.setChecked(self._grid_minor)
        alpha_spin = QtWidgets.QDoubleSpinBox(); alpha_spin.setRange(0.0, 1.0); alpha_spin.setSingleStep(0.05); alpha_spin.setValue(self._grid_alpha)
        lw_spin = QtWidgets.QDoubleSpinBox(); lw_spin.setRange(0.2, 2.0); lw_spin.setSingleStep(0.1); lw_spin.setValue(self._grid_lw)
        ls_combo = QtWidgets.QComboBox(); [ls_combo.addItem(lbl, val) for lbl, val in [("Solid", "-"), ("Dashed", "--"), ("Dotted", ":"), ("Dash-dot", "-.")]]
        ls_combo.setCurrentIndex(max(0, ls_combo.findData(self._grid_ls)))
        gm_row = QtWidgets.QWidget(); gm_h = QtWidgets.QHBoxLayout(gm_row); gm_h.setContentsMargins(6,2,6,2); gm_h.addWidget(grid_major_cb); gm_h.addWidget(grid_minor_cb)
        gm_act = QtWidgets.QWidgetAction(grid_menu); gm_act.setDefaultWidget(gm_row); grid_menu.addAction(gm_act)
        g2_row = QtWidgets.QWidget(); g2_h = QtWidgets.QHBoxLayout(g2_row); g2_h.setContentsMargins(6,2,6,2)
        g2_h.addWidget(QtWidgets.QLabel("Alpha")); g2_h.addWidget(alpha_spin)
        g2_h.addWidget(QtWidgets.QLabel("Width")); g2_h.addWidget(lw_spin)
        g2_h.addWidget(QtWidgets.QLabel("Style")); g2_h.addWidget(ls_combo)
        g2_act = QtWidgets.QWidgetAction(grid_menu); g2_act.setDefaultWidget(g2_row); grid_menu.addAction(g2_act)
        # Tick controls
        pos_options = ["Outside", "Inside", "Both", "None"]
        dir_map = {"Outside": "out", "Inside": "in", "Both": "inout", "None": "out"}
        def tick_section(axis_key):
            cfg = self._tick_cfg.get(axis_key, {})
            container = QtWidgets.QWidget(); layout = QtWidgets.QFormLayout(container); layout.setContentsMargins(6,2,6,2)
            pos_combo = QtWidgets.QComboBox(); pos_combo.addItems(pos_options)
            current_dir = cfg.get("direction", "out")
            reverse_map = {"out": "Outside", "in": "Inside", "inout": "Both"}
            pos_combo.setCurrentText(reverse_map.get(current_dir, "Outside"))
            maj_spin = QtWidgets.QDoubleSpinBox(); maj_spin.setRange(0.0, 1e6); maj_spin.setDecimals(6); maj_spin.setSingleStep(0.1)
            if cfg.get("major") is not None:
                maj_spin.setValue(cfg.get("major"))
            minor_spin = QtWidgets.QSpinBox(); minor_spin.setRange(0, 10); minor_spin.setValue(int(cfg.get("minor_count") or 0))
            len_spin = QtWidgets.QSpinBox(); len_spin.setRange(2, 20); len_spin.setValue(int(cfg.get("length") or 6))
            layout.addRow(f"{axis_key.upper()} position", pos_combo)
            layout.addRow("Major spacing", maj_spin)
            layout.addRow("Minor count", minor_spin)
            layout.addRow("Tick length (px)", len_spin)
            return container, pos_combo, maj_spin, minor_spin, len_spin
        x_widget, x_pos_combo, x_maj_spin, x_min_spin, x_len_spin = tick_section("x")
        y_widget, y_pos_combo, y_maj_spin, y_min_spin, y_len_spin = tick_section("y")
        x_act = QtWidgets.QWidgetAction(grid_menu); x_act.setDefaultWidget(x_widget); grid_menu.addAction(x_act)
        y_act = QtWidgets.QWidgetAction(grid_menu); y_act.setDefaultWidget(y_widget); grid_menu.addAction(y_act)
        both_cb = QtWidgets.QCheckBox("Apply X settings to Y"); both_cb.setChecked(False)
        both_act = QtWidgets.QWidgetAction(grid_menu); both_act.setDefaultWidget(both_cb); grid_menu.addAction(both_act)
        # Live legend handlers
        pos_combo.currentTextChanged.connect(lambda txt: self._set_legend_position(pos_map.get(txt, "best")))
        font_spin.valueChanged.connect(lambda val: self._set_legend_font(val))
        bg_act.toggled.connect(lambda checked: self._set_legend_bg(checked))
        border_act.toggled.connect(lambda checked: self._set_legend_border(checked))
        fname_act.toggled.connect(lambda checked: self._set_legend_filename_only(checked))
        # Live grid/tick handlers
        def apply_ticks():
            def _apply_axis(ax_key, pos_c, maj_c, min_c, len_c):
                cfg = self._tick_cfg.get(ax_key, {})
                cfg["direction"] = dir_map.get(pos_c.currentText(), "out")
                cfg["length"] = int(len_c.value())
                val = maj_c.value()
                cfg["major"] = float(val) if val > 0 else None
                cfg["minor_count"] = int(min_c.value())
                self._tick_cfg[ax_key] = cfg
            _apply_axis("x", x_pos_combo, x_maj_spin, x_min_spin, x_len_spin)
            if both_cb.isChecked():
                y_pos_combo.setCurrentText(x_pos_combo.currentText())
                y_maj_spin.setValue(x_maj_spin.value())
                y_min_spin.setValue(x_min_spin.value())
                y_len_spin.setValue(x_len_spin.value())
            _apply_axis("y", y_pos_combo, y_maj_spin, y_min_spin, y_len_spin)
            self._request_plot_update(delay_ms=25)
        grid_major_cb.toggled.connect(lambda chk: setattr(self, "_grid_major", bool(chk)) or self._request_plot_update(delay_ms=20))
        grid_minor_cb.toggled.connect(lambda chk: setattr(self, "_grid_minor", bool(chk)) or self._request_plot_update(delay_ms=20))
        alpha_spin.valueChanged.connect(lambda v: setattr(self, "_grid_alpha", float(v)) or self._request_plot_update(delay_ms=25))
        lw_spin.valueChanged.connect(lambda v: setattr(self, "_grid_lw", float(v)) or self._request_plot_update(delay_ms=25))
        ls_combo.currentIndexChanged.connect(lambda _i: setattr(self, "_grid_ls", ls_combo.currentData()) or self._request_plot_update(delay_ms=20))
        x_pos_combo.currentIndexChanged.connect(lambda _i: apply_ticks())
        y_pos_combo.currentIndexChanged.connect(lambda _i: apply_ticks())
        x_maj_spin.valueChanged.connect(lambda _v: apply_ticks())
        y_maj_spin.valueChanged.connect(lambda _v: apply_ticks())
        x_min_spin.valueChanged.connect(lambda _v: apply_ticks())
        y_min_spin.valueChanged.connect(lambda _v: apply_ticks())
        x_len_spin.valueChanged.connect(lambda _v: apply_ticks())
        y_len_spin.valueChanged.connect(lambda _v: apply_ticks())
        style_menu = menu.addMenu("Plot style")
        grid_act = style_menu.addAction("Show grid")
        grid_act.setCheckable(True)
        grid_act.setChecked(self._plot_grid_enabled)
        legend_act = style_menu.addAction("Show legend")
        legend_act.setCheckable(True)
        legend_act.setChecked(self._plot_legend_enabled)
        points_act = style_menu.addAction("Show points")
        points_act.setCheckable(True)
        points_act.setChecked(self.show_points_cb.isChecked())
        lines_act = style_menu.addAction("Show lines")
        lines_act.setCheckable(True)
        lines_act.setChecked(self.lines_cb.isChecked())
        style_menu.addSeparator()
        xlog_act = style_menu.addAction("Log X axis")
        xlog_act.setCheckable(True)
        xlog_act.setChecked(self._plot_x_log)
        ylog_act = style_menu.addAction("Log Y axis")
        ylog_act.setCheckable(True)
        ylog_act.setChecked(self._plot_y_log)
        style_menu.addSeparator()
        width_menu = style_menu.addMenu("Line width")
        width_actions = []
        width_presets = [
            ("Ultra thin (0.6 px)", 0.6),
            ("Thin (1.0 px)", 1.0),
            ("Medium (1.6 px)", 1.6),
            ("Bold (2.4 px)", 2.4),
            ("Heavy (3.5 px)", 3.5),
        ]
        for label, value in width_presets:
            act = width_menu.addAction(label)
            act.setCheckable(True)
            act.setChecked(abs(self._plot_line_width - value) < 0.21)
            act.setData(value)
            width_actions.append(act)
        width_menu.addSeparator()
        width_inc_act = width_menu.addAction("Increase")
        width_dec_act = width_menu.addAction("Decrease")
        reset_act = style_menu.addAction("Reset style")
        action = menu.exec_(self.canvas.mapToGlobal(pos))
        if action in preset_actions:
            self._apply_figure_preset(preset_actions[action])
        elif action == copy_png:
            self._copy_canvas_to_clipboard("png", dpi=300)
        elif action == copy_png_600:
            self._copy_canvas_to_clipboard("png", dpi=600)
        elif action == copy_svg:
            self._copy_canvas_to_clipboard("svg")
        elif action == copy_all:
            self._copy_all_traces_to_clipboard()
        elif action == minima_act:
            self._annotate_minima()
        elif action == resolve_act:
            self._resolve_minima_overlaps()
        elif action == add_point_act:
            self._add_point_label_at_cursor()
        elif action == clear_points_act:
            self._clear_point_labels()
        elif action == legend_show_act:
            self._plot_legend_enabled = legend_show_act.isChecked()
            self._update_plot()
        elif action == grid_act:
            self._plot_grid_enabled = grid_act.isChecked()
            self._grid_major = self._plot_grid_enabled
            self._update_plot()
        elif action == save_png_300:
            self._save_canvas("png", dpi=300)
        elif action == save_png_600:
            self._save_canvas("png", dpi=600)
        elif action == save_svg:
            self._save_canvas("svg")
        elif action == save_pdf:
            self._save_canvas("pdf")
        elif action == legend_act:
            self._plot_legend_enabled = legend_act.isChecked()
            self._update_plot()
        elif action == points_act:
            self._set_visual_checkbox(self.show_points_cb, points_act.isChecked())
        elif action == lines_act:
            self._set_visual_checkbox(self.lines_cb, lines_act.isChecked())
        elif action == xlog_act:
            self._set_plot_axis_log("x", xlog_act.isChecked())
        elif action == ylog_act:
            self._set_plot_axis_log("y", ylog_act.isChecked())
        elif action in width_actions:
            val = action.data()
            try:
                self._plot_line_width = float(val)
                self._update_plot()
            except Exception:
                pass
        elif action == width_inc_act:
            self._bump_line_width(+0.4)
        elif action == width_dec_act:
            self._bump_line_width(-0.4)
        elif action == reset_act:
            self._reset_plot_style()

    def set_plot_font_family(self, family: str):
        """Refresh the comparison plot with a new shared font family."""
        family = normalize_font_family(family, "sans-serif")
        self.set_plot_typography(family=family)

    def _apply_figure_preset(self, preset_key):
        """Apply a shared journal/slide layout preset to the comparison plot."""
        preset = get_figure_layout_preset(preset_key)
        self._figure_preset_key = preset.key
        apply_figure_layout(self.fig, preset)
        plot_w_px, plot_h_px = preset_pixel_size(self, preset, max_fraction=0.58)
        apply_canvas_widget_preset(self.canvas, preset, plot_w_px, plot_h_px)
        self._plot_font_family = normalize_font_family(preset.font_family, "sans-serif")
        self._plot_font_bold = False
        self._plot_font_italic = False
        self._plot_font_underline = False
        self._font_scale = float(preset.font_scale)
        self._legend_font = float(preset.legend_font_pt)
        self._plot_line_width = float(preset.line_width)
        try:
            total_w = max(1100, int(plot_w_px + 540))
            total_h = max(720, int(plot_h_px + 220))
            self.resize(total_w, total_h)
        except Exception:
            pass
        self._update_plot()
        QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), f"Applied preset: {preset.label}", self)

    def _font_style_state(self):
        return {
            "bold": bool(getattr(self, "_plot_font_bold", False)),
            "italic": bool(getattr(self, "_plot_font_italic", False)),
            "underline": bool(getattr(self, "_plot_font_underline", False)),
        }

    def set_plot_typography(self, **changes):
        """Refresh the comparison plot typography."""
        family = changes.get("family", None)
        viewer = getattr(self, "viewer", None)
        if family is not None:
            family = normalize_font_family(family, "sans-serif")
            self._plot_font_family = family
        if viewer is not None and hasattr(viewer, "set_plot_typography"):
            target = {
                "family": family if family is not None else self._plot_font_family,
                "bold": bool(changes.get("bold", self._plot_font_bold)),
                "italic": bool(changes.get("italic", self._plot_font_italic)),
                "underline": bool(changes.get("underline", self._plot_font_underline)),
            }
            if any(getattr(viewer, f"_plot_font_{k}", None) != v for k, v in target.items()):
                try:
                    viewer.set_plot_typography(**target)
                    return
                except Exception:
                    pass
        for key, attr in (("bold", "_plot_font_bold"), ("italic", "_plot_font_italic"), ("underline", "_plot_font_underline")):
            if key in changes:
                setattr(self, attr, bool(changes[key]))
        self._update_plot()

    def _on_compare_canvas_keypress(self, event):
        if not event or not hasattr(event, "key"):
            return
        key = (event.key or "").lower()
        if key in ("ctrl+z", "control+z"):
            self._undo_last_action()
            gui_event = getattr(event, "guiEvent", None)
            if gui_event:
                gui_event.accept()
    def _copy_canvas_to_clipboard(self, fmt, *, dpi=300):
        copy_figure_to_clipboard(self, self.fig, fmt, dpi=dpi)

    def _save_canvas(self, fmt, *, dpi=300):
        save_figure_with_dialog(self, self.fig, default_stem="spectroscopy_compare", fmt=fmt, dpi=dpi)

    def _set_hint_text(self, text=None):
        label = getattr(self, "hint_label", None)
        if label:
            value = text or self._delta_hint_text
            if value != getattr(self, "_last_hint_text", None):
                label.setText(value)
                self._last_hint_text = value

    def _set_canvas_cursor(self, shape):
        if shape == getattr(self, "_last_canvas_cursor", None):
            return
        try:
            self.canvas.setCursor(shape)
            self._last_canvas_cursor = shape
        except Exception:
            pass

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
        now = time.perf_counter()
        dragging_minima = bool(getattr(self, "_dragging_minima", None))
        if not dragging_minima and (now - getattr(self, "_last_hover_update_ts", 0.0)) < 0.04:
            return
        self._last_hover_update_ts = now
        # Mouse readout
        try:
            x = event.xdata
            y = event.ydata
        except Exception:
            x = y = None
        if x is not None and y is not None:
            self._last_mouse_xy = (float(x), float(y))
        lbl = getattr(self, "mouse_label", None)
        if lbl is not None:
            if x is None or y is None:
                text = "x: —   y: —"
            else:
                try:
                    text = f"x: {float(x):.4g}   y: {float(y):.4g}"
                except Exception:
                    text = f"x: {x}   y: {y}"
            if text != getattr(self, "_last_mouse_text", None):
                lbl.setText(text)
                self._last_mouse_text = text

        hovered = None
        if event and event.inaxes == self.ax and event.xdata is not None:
            hovered = self._find_nearest_lcpd_line(event.xdata)
        if self._delta_selection:
            if hovered:
                info = hovered[1]
                self._set_hint_text(
                    f"Shift+click {info.get('display_name', 'the line')} to finish ΔLCPD."
                )
                self._set_canvas_cursor(QtCore.Qt.PointingHandCursor)
            else:
                self._set_hint_text("Shift+click a second LCPD line to annotate ΔLCPD.")
                self._set_canvas_cursor(QtCore.Qt.ArrowCursor)
            return
        if hovered:
            info = hovered[1]
            self._set_hint_text(
                f"Shift+click {info.get('display_name', 'this LCPD')} to tag it for ΔLCPD."
            )
            self._set_canvas_cursor(QtCore.Qt.PointingHandCursor)
        else:
            self._set_hint_text()
            self._set_canvas_cursor(QtCore.Qt.ArrowCursor)
        if dragging_minima and event and event.inaxes == self.ax and event.ydata is not None:
            meta = self._dragging_minima
            txt = meta.get("text")
            if txt:
                txt.set_position((meta.get("x", event.xdata), float(event.ydata)))
                self.canvas.draw_idle()

    # Alias to satisfy mpl connections that refer to _on_mouse_move
    def _on_mouse_move(self, event):
        self._on_compare_canvas_motion(event)

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

    def _on_minima_press(self, event):
        if event is None or event.inaxes != self.ax or event.button != MouseButton.LEFT:
            return
        for meta in self._minima_meta:
            txt = meta.get("text")
            if txt is None:
                continue
            contains, _ = txt.contains(event)
            if contains:
                self._dragging_minima = meta
                break

    def _on_minima_motion(self, event):
        if not self._dragging_minima:
            return
        if event is None or event.inaxes != self.ax or event.ydata is None:
            return
        meta = self._dragging_minima
        txt = meta.get("text")
        if txt is None:
            return
        x_fixed = meta.get("x", event.xdata)
        txt.set_position((x_fixed, float(event.ydata)))
        self.canvas.draw_idle()

    def _on_minima_release(self, event):
        if event is None or event.button != MouseButton.LEFT:
            return
        self._dragging_minima = None

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
        label = "Visual toggle"
        if sender:
            text_attr = getattr(sender, "text", None)
            if callable(text_attr):
                try:
                    label = text_attr() or label
                except Exception:
                    pass
        self._record_user_action(f"{label} → {'on' if checked else 'off'}")
        self._request_plot_update(delay_ms=20)

    def _on_offset_changed(self, value):
        self._record_user_action(f"Waterfall offset → {value:.3g}")
        self._request_plot_update(delay_ms=20)

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
            if row >= 0:
                self.spec_list.takeTopLevelItem(row)
            removed = True
        if removed:
            self._request_plot_update(delay_ms=20)
            self._populate_results_table()
            self._update_fit_trend_state()
            self._update_status()

    def _clear_all(self):
        self._record_user_action("Clear all spectra")
        self.spec_list.clear()
        self._item_map = {}
        self._line_map.clear()
        self._legend_map.clear()
        self._fit_results = {}
        self._plotted_spec_ids = []
        self._clear_curve_data_cache()
        self._update_fit_trend_state()
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
        self._request_plot_update(delay_ms=20)

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
        self._request_plot_update(delay_ms=20)

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
        self._update_fit_trend_state()
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
            z_disp, z_tooltip = self._format_fit_result_z(spec)
            rows.append((spec_id, self._display_name(spec),
                         "n/a" if xs is None else f"{xs:.1f}",
                         "n/a" if ys is None else f"{ys:.1f}",
                         z_disp, z_tooltip,
                         f"{res['a']:.4g}", f"{res['a_err']:.2g}",
                         v0_disp, v0_err_disp,
                         f"{res['c']:.4g}", f"{res['c_err']:.2g}",
                         f"{res['rmse']:.4g}"))
        self.results_table.setRowCount(len(rows))
        for r, data in enumerate(rows):
            spec_id, name, xval, yval, zval, z_tooltip, a, ae, b, be, c, ce, rmse = data
            values = [name, xval, yval, zval, a, ae, b, be, c, ce, rmse]
            for col, val in enumerate(values):
                item = QtWidgets.QTableWidgetItem(val)
                if col == 0:
                    item.setData(QtCore.Qt.UserRole, spec_id)
                if col == 3 and z_tooltip:
                    item.setToolTip(z_tooltip)
                self.results_table.setItem(r, col, item)

    def _export_csv(self):
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Export CSV", "spectroscopy_fit.csv", "CSV Files (*.csv)")
        if not path:
            return
        headers = [
            self.results_table.horizontalHeaderItem(c).text() if self.results_table.horizontalHeaderItem(c) else ""
            for c in range(self.results_table.columnCount())
        ]
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
        <li><b>Axis:</b> Select the X-axis (bias voltage, Z, or Topo position)</li>
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
        style = _style_kwargs(self._font_style_state())
        apply_text_style(self.ax.xaxis.label, family=self._plot_font_family, **style)
        apply_text_style(self.ax.yaxis.label, family=self._plot_font_family, **style)
        for text in list(self.ax.get_xticklabels()) + list(self.ax.get_yticklabels()):
            apply_text_style(text, family=self._plot_font_family, **style)
        if self.ax.get_title():
            apply_text_style(self.ax.title, family=self._plot_font_family, **style)
        if self.ax.get_legend():
            plt_legend = self.ax.get_legend()
            for text in plt_legend.get_texts():
                text.set_fontsize(8 * scale)
                apply_text_style(text, family=self._plot_font_family, **style)
        for meta in getattr(self, "_minima_meta", []):
            txt = meta.get("text")
            if txt:
                txt.set_fontsize(7 * scale)
                apply_text_style(txt, family=self._plot_font_family, **style)
        for pl in getattr(self, "_point_labels", []):
            txt = pl.get("text")
            if txt:
                txt.set_fontsize(7 * scale)
                apply_text_style(txt, family=self._plot_font_family, **style)
        self.canvas.draw_idle()

    def _curve_name(self, spec_id):
        try:
            for spec in self.specs:
                if self._spec_id(spec) == spec_id:
                    return self._display_name(spec)
        except Exception:
            pass
        return str(spec_id)

    def _set_curve_style(self, spec_id, *, lw=None, ls=None, color=None):
        try:
            st = self._curve_styles.get(spec_id, {})
            if lw is not None:
                st["lw"] = lw
            if ls is not None:
                st["ls"] = ls
            if color is not None:
                st["color"] = color
            self._curve_styles[spec_id] = st
            line = self._line_map.get(spec_id)
            if line:
                if lw is not None:
                    line.set_linewidth(lw)
                if ls is not None:
                    line.set_linestyle(ls)
                if color is not None:
                    line.set_color(color)
            self.canvas.draw_idle()
        except Exception:
            pass

    def _reset_colors_to_cycle(self):
        if not self._line_map:
            return
        cycle = list(self._color_cycle) if getattr(self, "_color_cycle", None) else []
        if not cycle:
            cycle = get_color_cycle(self._palette_name or DEFAULT_COLOR_CYCLE)
        if not cycle:
            return
        for idx, (spec_id, line) in enumerate(self._line_map.items()):
            color = cycle[idx % len(cycle)]
            self._set_curve_style(spec_id, color=color)

    def _apply_global_line_style(self, lw, ls):
        try:
            lw = float(lw)
        except Exception:
            lw = None
        for spec_id in list(self._line_map.keys()):
            self._set_curve_style(spec_id, lw=lw, ls=ls)
        try:
            if lw is not None:
                self._plot_line_width = lw
        except Exception:
            pass
        self.canvas.draw_idle()

    def _pick_curve_color(self, spec_id, btn):
        current = None
        try:
            line = self._line_map.get(spec_id)
            current = line.get_color() if line else None
        except Exception:
            current = None
        color = QtWidgets.QColorDialog.getColor(QtGui.QColor(current or "#000000"), self, "Select line color")
        if not color.isValid():
            return
        hex_col = color.name()
        btn.setStyleSheet(f"background:{hex_col};")
        self._set_curve_style(spec_id, color=hex_col)

    def _apply_legend_settings(self):
        legend = self.ax.get_legend()
        if not legend:
            return
        try:
            legend.set_visible(self._plot_legend_enabled)
            frame = legend.get_frame()
            frame.set_alpha(0.9 if self._legend_bg else 0.0)
            frame.set_facecolor("white" if self._legend_bg else (0, 0, 0, 0))
            frame.set_edgecolor("black" if self._legend_border else (0, 0, 0, 0))
            frame.set_linewidth(0.8 if self._legend_border else 0.0)
            for text in legend.get_texts():
                text.set_fontsize(self._legend_font)
            if self._plot_legend_enabled:
                legend.set_draggable(len(getattr(self, "_plotted_spec_ids", [])) < 16)
            self.canvas.draw_idle()
        except Exception:
            pass

    def _apply_grid_and_ticks(self):
        ax = self.ax
        try:
            if self._grid_major:
                ax.grid(True, which="major", alpha=self._grid_alpha, linewidth=self._grid_lw, linestyle=self._grid_ls)
            else:
                ax.grid(False, which="major")
            if self._grid_minor:
                ax.grid(True, which="minor", alpha=self._grid_alpha * 0.8, linewidth=max(0.2, self._grid_lw * 0.7), linestyle=self._grid_ls)
            else:
                ax.grid(False, which="minor")
        except Exception:
            pass
        # Ticks
        for axis_key, axis in (("x", ax.xaxis), ("y", ax.yaxis)):
            cfg = self._tick_cfg.get(axis_key, {})
            direction = cfg.get("direction", "out")
            length = int(cfg.get("length", 6))
            try:
                ax.tick_params(axis=axis_key, direction=direction, length=length)
            except Exception:
                pass
            major = cfg.get("major")
            if major and major > 0:
                try:
                    if abs(major) < 1e-12:
                        major = None
                    else:
                        data_lim = axis.get_view_interval()
                        span = abs(data_lim[1] - data_lim[0])
                        if np.isfinite(span) and span / major > 2000:
                            raise ValueError("Too many major ticks requested.")
                        axis.set_major_locator(MultipleLocator(major))
                except Exception:
                    try:
                        axis.set_major_locator(MaxNLocator(nbins='auto'))
                    except Exception:
                        pass
            else:
                try:
                    axis.set_major_locator(MaxNLocator(nbins='auto'))
                except Exception:
                    pass
            minor_count = int(cfg.get("minor_count", 0))
            if minor_count > 0:
                try:
                    minor_count = max(1, min(10, minor_count))
                    axis.set_minor_locator(AutoMinorLocator(minor_count + 1))
                except Exception:
                    pass
            else:
                try:
                    axis.set_minor_locator(AutoMinorLocator())
                except Exception:
                    pass

    def _set_legend_position(self, loc):
        self._legend_loc = loc or "best"
        self._update_plot()

    def _set_legend_font(self, size):
        try:
            self._legend_font = float(size)
        except Exception:
            return
        self._apply_legend_settings()

    def _set_legend_bg(self, enabled):
        self._legend_bg = bool(enabled)
        self._apply_legend_settings()

    def _set_legend_border(self, enabled):
        self._legend_border = bool(enabled)
        self._apply_legend_settings()

    def _set_legend_filename_only(self, enabled):
        self._legend_filename_only = bool(enabled)
        self._update_plot()

    def _log(self, text):
        self.log.appendPlainText(text)
