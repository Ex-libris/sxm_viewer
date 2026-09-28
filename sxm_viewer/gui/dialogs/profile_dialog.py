"""Profile measurement dialog.

matplotlib-based line-profile viewer - see the project's profile-dialog
redesign brief and ``profile_dialog_v2.py`` (a standalone mockup) for the
full plan this implements (units, measuring, discoverability, exports,
visual weight, the profile-list table).

Deliberately matplotlib, not pyqtgraph: an earlier pass of this redesign
used pyqtgraph (its ``LinearRegionItem``/``enableAutoSIPrefix`` are a
genuinely good fit for the band and the unit formatting), but some of the
PCs this app runs on sit next to the microscope, offline, with a
locked-down Python install that only has what the app already ships with -
`pip install`ing a new dependency there isn't an option. matplotlib is
already a hard dependency of the whole app, so everything here is built on
it: ``matplotlib.widgets.SpanSelector`` (interactive + drag-from-anywhere,
available since matplotlib 3.5, well within this project's
``matplotlib>=3.8.2`` pin) for the measurement band, and
``matplotlib.ticker.EngFormatter`` plus a small hand-rolled SI-prefix
formatter (``profile_units.si_format``, numpy-only) for unit display.
Reverting to matplotlib also means regaining this codebase's existing
matplotlib infrastructure (``plot_typography.py``, ``figure_layout_
presets.py``) that a pyqtgraph version would have had to duplicate.

Units (the reason this dialog was touched at all)
---------------------------------------------------
Two real bugs motivated this: a y axis labelled ``(Hz)`` with a bare
``1e-15`` multiplier floating above it, and a y axis labelled
``Current [A] (A)`` (the unit printed twice) showing non-physical current
values. Both were symptoms of treating units as display strings instead of
data:

* ``profile_units.to_si_base`` converts the already-normalized values this
  dialog receives (``nm``/``A``/``V``/``Hz`` - see ``data.io.
  normalize_unit_and_data``) to true SI base units (``m``/``A``/``V``/
  ``Hz``) exactly once, at ingestion. ``EngFormatter``/``si_format`` then
  derive ``pA``/``nm``/``mHz`` display on their own - a bare multiplier is
  structurally impossible once the axis is fed true base units.
* ``profile_units.compose_axis_label`` builds a channel's axis label from
  its name + unit exactly once, stripping any ``[unit]``/``(unit)`` the
  channel name already embeds, so ``Current [A] (A)`` can't be constructed.

The distance axis is converted from the app-wide ``nm`` convention to
metres *only inside this dialog*, at the boundary where a profile dataset
dict enters it - the shared extraction path (``detail_preview_canvas.
_build_profile_data``) and every other nm-based consumer (session state,
on-image overlays, other dialogs) are untouched. Positions handed back to
the source canvas (``marker_update_callback`` / ``set_marker_positions``)
are converted back to nm/px at that same boundary, so the live link with
the drawn profile line on the image keeps working in the units it always
used.
"""
from __future__ import annotations

import copy
import hashlib
import json

import numpy as np
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from matplotlib.ticker import AutoMinorLocator, EngFormatter, ScalarFormatter
from matplotlib.widgets import SpanSelector

from ..._shared import QtCore, QtGui, QtWidgets, Path, matplotlib
from ..canvases.detail_preview import SafeFigureCanvas
from ..plot_typography import add_font_menu_action, normalize_font_family, apply_text_style, apply_qfont_style
from ..figure_layout_presets import (
    FigureLayoutPreset,
    apply_figure_layout,
    preset_pixel_size,
    apply_canvas_widget_preset,
)
from ..palettes import DEFAULT_COLOR_CYCLE, get_color_cycle, list_color_cycles
from ..profile_links import (
    register_profile_dialog,
    unregister_profile_dialog,
    apply_live_profile_style,
    profile_ref_key,
)
from .. import theme as ui_theme
from .profile_units import to_si_base, compose_axis_label, si_format
from .profile_measure import fmt_tol, find_peaks, band_statistics, peaks_in_range, snap_edges

_PROFILE_COMPOSITE_MIME = "application/x-sxm-profile-composite"

# Curve colors when a dataset carries none of its own: desaturated amber -
# a saturated hairline over high-frequency spikes shimmers, which is what
# the previous dialog's brighter default did.
_FALLBACK_ACTIVE_COLOR_LIGHT = "#D9822B"
_FALLBACK_ACTIVE_COLOR_DARK = "#E0A458"
_FALLBACK_OVERLAY_COLOR_LIGHT = "#2E7D8F"
_FALLBACK_OVERLAY_COLOR_DARK = "#5BC0BE"
_DEFAULT_CURVE_WIDTH = 1.3

# Figure presets local to this dialog, not the shared figure_layout_presets.py
# tuple: that module's presets are square (image-panel dialogs), while a
# line profile is inherently a wide trace, so a shared square aspect would
# be wrong here. Reuses the same FigureLayoutPreset shape (and the shared
# module's generic, engine-agnostic preset_pixel_size/apply_canvas_widget_preset
# helpers) purely for structural consistency.
_PROFILE_FIGURE_PRESETS: tuple = (
    FigureLayoutPreset("interactive", "Interactive", 152.4, 101.6, "sans-serif", 1.0, 8.0, 1.6),
    FigureLayoutPreset("single_column_85", "Single column (85 mm)", 85.0, 55.0, "Arial", 0.78, 6.0, 1.0),
    FigureLayoutPreset("double_column_180", "Double column (180 mm)", 180.0, 95.0, "Arial", 0.92, 7.0, 1.3),
    FigureLayoutPreset("slide_254", "Slide (254 mm)", 254.0, 143.0, "Arial", 1.1, 9.0, 1.8),
)


def _get_profile_figure_preset(key: "str | None") -> FigureLayoutPreset:
    wanted = str(key or "").strip()
    for preset in _PROFILE_FIGURE_PRESETS:
        if preset.key == wanted:
            return preset
    return _PROFILE_FIGURE_PRESETS[0]


class _ProfileCompositeDragButton(QtWidgets.QToolButton):
    """Drag button used to compose profile dialogs without interfering with plot gestures."""

    def __init__(self, owner):
        super().__init__(owner)
        self._owner_dialog = owner
        self._drag_start_pos = None
        self._drag_started = False
        self.setObjectName("profileComposeButton")
        self.setText("Combine")
        self.setCursor(QtCore.Qt.OpenHandCursor)
        self.setAutoRaise(False)
        self.setToolButtonStyle(QtCore.Qt.ToolButtonTextOnly)
        self.setToolTip(
            "Drag this onto another profile window to create a new composite window."
        )

    def mousePressEvent(self, event):
        if event.button() == QtCore.Qt.LeftButton:
            self._drag_start_pos = event.pos()
            self._drag_started = False
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if (
            self._drag_start_pos is not None
            and event.buttons() & QtCore.Qt.LeftButton
            and (event.pos() - self._drag_start_pos).manhattanLength() >= QtWidgets.QApplication.startDragDistance()
        ):
            self._drag_start_pos = None
            self._drag_started = True
            try:
                self._owner_dialog.start_profile_composite_drag()
            except Exception:
                pass
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if (
            event.button() == QtCore.Qt.LeftButton
            and not self._drag_started
            and hasattr(self._owner_dialog, "_show_compose_help")
        ):
            try:
                self._owner_dialog._show_compose_help(self.mapToGlobal(event.pos()))
            except Exception:
                pass
        self._drag_start_pos = None
        self._drag_started = False
        super().mouseReleaseEvent(event)


class _Readout(QtWidgets.QFrame):
    """Permanent measurement strip: two rows of five fields below the plot.

    Replaces the old in-plot distance arrow - an annotation drawn over the
    data collides with the trace and ends up baked into exported figures.
    Monospaced, selectable values so a number can be copied straight out.
    """

    ROW1 = [("cursor", "Cursor"), ("left", "Left edge"), ("right", "Right edge"),
            ("dx", "Width Δd"), ("dy", "Height Δy")]
    ROW2 = [("mean", "Mean"), ("rms", "RMS"), ("slope", "Slope"),
            ("peaks", "Peaks in band"), ("px", "Pixel size")]

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("profileReadout")
        self.value_labels = {}
        grid = QtWidgets.QGridLayout(self)
        grid.setContentsMargins(10, 6, 10, 6)
        grid.setHorizontalSpacing(18)
        grid.setVerticalSpacing(1)
        mono = QtGui.QFontDatabase.systemFont(QtGui.QFontDatabase.FixedFont)
        for row, fields in ((0, self.ROW1), (2, self.ROW2)):
            for col, (key, title) in enumerate(fields):
                cap = QtWidgets.QLabel(title.upper())
                cap.setObjectName("profileReadoutCaption")
                val = QtWidgets.QLabel("-")
                val.setFont(mono)
                val.setObjectName("profileReadoutValue")
                val.setTextInteractionFlags(QtCore.Qt.TextSelectableByMouse)
                grid.addWidget(cap, row, col)
                grid.addWidget(val, row + 1, col)
                self.value_labels[key] = val
        for c in range(5):
            grid.setColumnStretch(c, 1)

    def set(self, key, text):
        self.value_labels[key].setText(text)

    def clear_all(self):
        for label in self.value_labels.values():
            label.setText("-")

    def apply_theme(self, *, text_color, muted_color, panel_color, border_color):
        self.setStyleSheet(
            "QFrame#profileReadout {"
            f"background: {panel_color}; border: 1px solid {border_color};"
            "border-radius: 6px; }"
            "QLabel#profileReadoutCaption {"
            f"color: {muted_color}; font-size: 7.5pt; letter-spacing: 0.5px; }}"
            "QLabel#profileReadoutValue {"
            f"color: {text_color}; font-size: 9.5pt; }}"
        )


class _Toast(QtWidgets.QLabel):
    """Transient confirmation naming what was copied/saved and at what size.

    Silent clipboard writes make people click twice to be sure.
    """

    def __init__(self, parent):
        super().__init__(parent)
        self.setObjectName("profileToast")
        self.setAlignment(QtCore.Qt.AlignCenter)
        self.hide()
        self._timer = QtCore.QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.timeout.connect(self.hide)

    def show_message(self, text, msec=2600):
        self.setText(text)
        self.adjustSize()
        parent = self.parent()
        if parent is not None:
            self.move(max(8, (parent.width() - self.width()) // 2), 10)
        self.raise_()
        self.show()
        self._timer.start(msec)

    def apply_theme(self, *, text_color, bg_color):
        self.setStyleSheet(
            "QLabel#profileToast {"
            f"background: {bg_color}; color: {text_color};"
            "border-radius: 6px; padding: 6px 14px; font-size: 9pt; font-weight: 600;"
            "}"
        )


class _ColorChip(QtWidgets.QFrame):
    """A profile's color swatch as a real cell widget, not an item background.

    An item background gets washed out by a selected row's highlight; a
    small widget painted on top of the row never does. Double-click opens
    a color picker.
    """

    def __init__(self, color, on_double_click=None, parent=None):
        super().__init__(parent)
        self._on_double_click = on_double_click
        self.setFixedSize(30, 14)
        self.setCursor(QtCore.Qt.PointingHandCursor)
        self.setToolTip("Double-click to change color")
        self.set_color(color)

    def set_color(self, color):
        color = color or "#888888"
        self.setStyleSheet(
            f"background-color: {color}; border-radius: 3px; border: 1px solid rgba(0, 0, 0, 70);"
        )

    def mouseDoubleClickEvent(self, event):
        if event.button() == QtCore.Qt.LeftButton and callable(self._on_double_click):
            self._on_double_click()
        super().mouseDoubleClickEvent(event)


class ProfileDialog(QtWidgets.QDialog):
    """Dialog showing the sampled profile and basic stats."""

    def __init__(self, active_profile, saved_profiles=None, parent=None, unit=None, y_label=None,
                 activate_overlay_callback=None, highlight_overlay_callback=None,
                 label_scale_callback=None, delete_overlay_callback=None,
                 marker_update_callback=None, marker_select_callback=None,
                 add_overlay_callback=None, style_update_callback=None,
                 palette_callback=None, profile_display_callback=None, dark_mode=False):
        super().__init__(parent)
        self.setWindowTitle('Profile measurement')
        self.setAcceptDrops(True)
        self.setWindowFlags(
            self.windowFlags()
            | QtCore.Qt.WindowMinimizeButtonHint
            | QtCore.Qt.WindowSystemMenuHint
        )
        self.resize(900, 600)
        self.setMinimumSize(700, 450)
        self._unit = unit
        self._y_label = y_label
        self._owner_dark_mode_hint = bool(dark_mode)
        self._active = None
        self._saved = []

        # -- measurement band state -----------------------------------------
        # The band's external-facing position/domain (what marker_update_cb
        # / set_marker_positions exchange with the source canvas) stays in
        # the SAME units the canvas always used (nm, or px when a profile
        # has no physical extent) - the live link needs no unit awareness
        # of its own. Everything else (the SpanSelector itself, the readout
        # strip's numbers, snap-to-peak, keyboard nudge) operates directly
        # in plot-space SI-base units (metres/A/V/Hz), since that is what's
        # actually plotted after Priority 1's unit fix.
        # _external_to_plot_x/_plot_to_external_x are the sole boundary.
        # Unlike a signal-driven widget, SpanSelector.extents = (...) never
        # re-triggers onselect/onmove_callback (confirmed empirically) -
        # every programmatic move (snap, nudge, the live link) drives the
        # readout/notify update explicitly, so no recursion guard is needed.
        settings = QtCore.QSettings()
        self._span = None
        self._vline = None
        self._hline = None
        self._marker_positions = []
        self._marker_domain = (0.0, 1.0)
        self._marker_reference_state = (None, None, None)
        self._marker_saved_positions = None
        self._markers_enabled = bool(settings.value("profileDialog/measureEnabled", True, type=bool))
        self._snap_enabled = bool(settings.value("profileDialog/snapToPeaks", True, type=bool))
        self._crosshair_enabled = bool(settings.value("profileDialog/crosshairEnabled", True, type=bool))
        self._include_band_in_export = bool(settings.value("profileDialog/includeBandInExport", False, type=bool))
        self._marker_syncing = False
        self._marker_positions_by_key = {}
        self._marker_domain_by_key = {}
        self._current_marker_key = None
        self._overlay_visible_by_key = {}
        self._x_si_factor = 1.0
        self._x_unit = ''
        self._x_auto_prefix = False
        self._ref_x_plot = None
        self._ref_y_plot = None
        self._ref_y_unit_si = ''
        self._ref_peak_idx = np.array([], dtype=int)
        self._px_size_plot = 0.0
        self._readout_precision = 4
        self._hint_dismissed = bool(settings.value("profileDialog/hintDismissed", False, type=bool))
        self._plot_font_family = normalize_font_family(getattr(parent, "_plot_font_family", None), "sans-serif")
        self._plot_font_bold = bool(getattr(parent, "_plot_font_bold", False))
        self._plot_font_italic = bool(getattr(parent, "_plot_font_italic", False))
        self._plot_font_underline = bool(getattr(parent, "_plot_font_underline", False))

        self._last_saved_count = 0
        self._line_handles_by_key = {}
        self._ordered_profile_entries_cache = []
        self._toggle_buttons = []
        self._legend_visible = True
        self._legend_fontsize = 8.0
        self._legend_item = None
        self._figure_preset_key = "interactive"
        self._metadata_visible = False
        self._metadata_show_filename = True
        self._metadata_show_acquisition = True
        self._metadata_show_time = False
        self._metadata_show_folder_name = False
        self._metadata_show_folder = False
        self._metadata_item = None
        self._owner = parent
        self._theme_mode = str(settings.value("profileDialog/themeMode", "system"))
        if self._theme_mode not in ("light", "dark", "system"):
            self._theme_mode = "system"
        self._dark_background = self._resolve_dark_background()
        self._workspace_registered = False
        self._composite_mode = False
        self._composite_origin_id = hex(id(self))
        self._canvas_drag_start_pos = None
        self._canvas_drag_started = False
        self._font_scale = 1.0
        self._right_curves = []
        self._label_scale_cb = label_scale_callback
        self._activate_overlay_cb = activate_overlay_callback
        self._highlight_overlay_cb = highlight_overlay_callback
        self._delete_overlay_cb = delete_overlay_callback
        self._marker_update_cb = marker_update_callback
        self._marker_key_cb = marker_select_callback
        self._add_overlay_cb = add_overlay_callback
        self._style_update_cb = style_update_callback
        self._palette_cb = palette_callback
        self._profile_display_cb = profile_display_callback
        self._profile_palette_name = DEFAULT_COLOR_CYCLE
        self._preserve_cb = None
        self._context_source = None

        self._build_ui()
        self._build_shortcuts()
        self._apply_plot_theme()
        self.update_profiles(active_profile, saved_profiles or [], activate_overlay_callback=activate_overlay_callback)
        self._apply_font_scale()
        if callable(self._label_scale_cb):
            self._label_scale_cb(self._font_scale)
        self._refresh_action_button_states()
        register_profile_dialog(self)

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(12, 10, 12, 10)
        root.setSpacing(8)

        # -- header: identity on the left, export promoted to a visible
        # split button on the right. Export/style used to live only behind
        # an unadvertised right-click; no other dialog in this app has a
        # shared top-strip convention to fold into (spectroscopy_dialogs.py
        # exposes the same actions the same way - context-menu-only, plus
        # unrelated QPushButton rows for data-copy actions), so this header
        # is new to this window rather than reused from elsewhere.
        header = QtWidgets.QHBoxLayout()
        header.setContentsMargins(0, 0, 0, 0)
        header.setSpacing(8)
        self.header_title = QtWidgets.QLabel("Profile measurement")
        self.header_title.setObjectName("profileHeaderTitle")
        header.addWidget(self.header_title)
        header.addStretch(1)
        self.export_btn = QtWidgets.QToolButton()
        self.export_btn.setObjectName("profileExportButton")
        self.export_btn.setText("  Copy plot  ")
        self.export_btn.setToolTip("Copy a 300 dpi PNG to the clipboard (Ctrl+C). Click the arrow for more options.")
        self.export_btn.setPopupMode(QtWidgets.QToolButton.MenuButtonPopup)
        self.export_btn.setToolButtonStyle(QtCore.Qt.ToolButtonTextOnly)
        self.export_btn.clicked.connect(lambda: self._copy_plot("png", dpi=300))
        self._export_menu = QtWidgets.QMenu(self)
        self._export_menu.aboutToShow.connect(self._populate_export_menu)
        self.export_btn.setMenu(self._export_menu)
        header.addWidget(self.export_btn)

        # View menu replaces the old per-window "Dark" toggle button - SPM
        # labs are dim and this app already has its own app-wide theme
        # system (gui/theme.py), so a lone per-dialog toggle fighting that
        # is the wrong control; Light/Dark/Follow-system here, persisted.
        self.view_btn = QtWidgets.QToolButton()
        self.view_btn.setObjectName("profileViewButton")
        self.view_btn.setText("  View ▾  ")
        self.view_btn.setPopupMode(QtWidgets.QToolButton.InstantPopup)
        self.view_btn.setToolButtonStyle(QtCore.Qt.ToolButtonTextOnly)
        self._view_menu = QtWidgets.QMenu(self)
        self._view_menu.aboutToShow.connect(self._populate_view_menu)
        self.view_btn.setMenu(self._view_menu)
        header.addWidget(self.view_btn)
        root.addLayout(header)

        # -- one-time hint bar: states the three non-obvious interactions
        # once, dismissible, persisted so it doesn't reappear every open.
        self.hint_bar = QtWidgets.QFrame()
        self.hint_bar.setObjectName("profileHintBar")
        hint_layout = QtWidgets.QHBoxLayout(self.hint_bar)
        hint_layout.setContentsMargins(10, 5, 6, 5)
        hint_label = QtWidgets.QLabel(
            "Drag the shaded band or its edges to measure · "
            "←/→ nudge the left edge, Shift+←/→ the right · "
            "right-click the plot for export/style options"
        )
        hint_label.setObjectName("profileHintLabel")
        hint_layout.addWidget(hint_label)
        hint_layout.addStretch(1)
        hint_close = QtWidgets.QToolButton()
        hint_close.setText("×")
        hint_close.setObjectName("profileHintClose")
        hint_close.setAutoRaise(True)
        hint_close.clicked.connect(self._dismiss_hint_bar)
        hint_layout.addWidget(hint_close)
        root.addWidget(self.hint_bar)
        self.hint_bar.setVisible(not self._hint_dismissed)

        splitter = QtWidgets.QSplitter(QtCore.Qt.Vertical)
        self._splitter = splitter

        fig = Figure(figsize=(6, 3))
        self.canvas = SafeFigureCanvas(fig)
        self.canvas.installEventFilter(self)
        self.canvas.setToolTip(
            "Drag the plot margin onto another profile window to create a composite."
        )
        self.ax = fig.add_subplot(111)
        self.canvas.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.canvas.customContextMenuRequested.connect(self._on_context_menu)
        self.ax_right = self.ax.twinx()
        self.ax_right.set_visible(False)
        self.toast = _Toast(self.canvas)

        # Measurement band (both edges + whole band draggable) - replaces
        # the old pair of thin, hard-to-grab vertical lines. SpanSelector's
        # onselect fires on mouse release (-> _on_span_select, matching the
        # old pyqtgraph sigRegionChangeFinished); onmove_callback fires
        # continuously during a drag (-> _on_span_move, matching
        # sigRegionChanged).
        self._span = SpanSelector(
            self.ax, self._on_span_select, "horizontal",
            interactive=True, drag_from_anywhere=True, useblit=False,
            onmove_callback=self._on_span_move,
            props=dict(alpha=0.25),
        )
        self._span.set_active(self._markers_enabled)
        self._span.set_visible(self._markers_enabled)

        # Crosshair: hidden until the pointer is actually over the plot.
        self._vline = self.ax.axvline(0, visible=False, zorder=5, lw=1, ls="--")
        self._hline = self.ax.axhline(0, visible=False, zorder=5, lw=1, ls="--")
        self._motion_cid = self.canvas.mpl_connect("motion_notify_event", self._on_mouse_moved)

        self.readout = _Readout()
        plot_column = QtWidgets.QWidget()
        plot_column_layout = QtWidgets.QVBoxLayout(plot_column)
        plot_column_layout.setContentsMargins(0, 0, 0, 0)
        plot_column_layout.setSpacing(6)
        plot_column_layout.addWidget(self.canvas, 1)
        plot_column_layout.addWidget(self.readout)
        splitter.addWidget(plot_column)

        info_widget = QtWidgets.QWidget()
        info_layout = QtWidgets.QVBoxLayout(info_widget)
        info_layout.setContentsMargins(0, 0, 0, 0)
        self.stats = QtWidgets.QLabel("")
        self.stats.setAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignTop)
        self.stats.setWordWrap(True)
        self.stats.setVisible(False)
        self.marker_info = QtWidgets.QLabel("")
        self.marker_info.setAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignTop)
        # Initial visibility mirrors _on_measure_toggle's steady state, but
        # that handler is wired up *after* the toggle button's initial
        # setChecked() (see _make_toggle_button) so it never fires during
        # construction - set both explicitly here instead.
        self.readout.setVisible(self._markers_enabled)
        self.marker_info.setVisible(not self._markers_enabled)
        self.controls_hint = QtWidgets.QLabel(
            "Shortcuts: M measure, S snap, C crosshair, G grid, L lines, P points, R reset zoom, "
            "←/→ nudge left edge (Shift=right edge, Ctrl=×10), "
            "Del remove overlay, Ctrl+Wheel font size"
        )
        self.controls_hint.setObjectName("profileControlsHint")
        self.controls_hint.setWordWrap(True)
        self.controls_hint.setVisible(False)

        controls_panel = QtWidgets.QWidget()
        controls_layout = QtWidgets.QVBoxLayout(controls_panel)
        controls_layout.setContentsMargins(0, 0, 0, 0)
        controls_layout.setSpacing(6)

        # One row, grouped with thin separators, instead of a loud primary
        # row plus a hidden-behind-"Advanced" second row: the toggles are
        # cheap view options and shouldn't out-compete the plot for
        # attention every time the eye crosses the dialog.
        toggle_row = QtWidgets.QHBoxLayout()
        toggle_row.setContentsMargins(0, 0, 0, 0)
        toggle_row.setSpacing(4)

        self.measure_toggle = self._make_toggle_button(
            "Measure", checked=self._markers_enabled,
            tooltip="Show/hide the draggable measurement band and readout strip (M)",
        )
        self.measure_toggle.toggled.connect(self._on_measure_toggle)
        toggle_row.addWidget(self.measure_toggle)

        self.show_lines_cb = self._make_toggle_button(
            "Lines", checked=True, tooltip="Show connecting profile line (L)"
        )
        self.show_lines_cb.toggled.connect(self._on_plot_option_changed)
        toggle_row.addWidget(self.show_lines_cb)

        self.show_points_cb = self._make_toggle_button(
            "Points", checked=False,
            tooltip="Show individual sampled points - reveals the pixel grid (P)",
        )
        self.show_points_cb.toggled.connect(self._on_plot_option_changed)
        toggle_row.addWidget(self.show_points_cb)

        self.grid_cb = self._make_toggle_button(
            "Grid", checked=False, tooltip="Toggle background grid (G)"
        )
        self.grid_cb.toggled.connect(self._on_grid_toggled)
        toggle_row.addWidget(self.grid_cb)

        self._add_toggle_row_separator(toggle_row)

        self.snap_toggle = self._make_toggle_button(
            "Snap to peaks", checked=self._snap_enabled,
            tooltip="Band edges jump to the nearest detected peak (S)",
        )
        self.snap_toggle.toggled.connect(self._on_snap_toggle)
        toggle_row.addWidget(self.snap_toggle)

        self.crosshair_toggle = self._make_toggle_button(
            "Crosshair", checked=self._crosshair_enabled,
            tooltip="Live cursor readout while hovering the plot (C)",
        )
        self.crosshair_toggle.toggled.connect(self._on_crosshair_toggle)
        toggle_row.addWidget(self.crosshair_toggle)

        self._add_toggle_row_separator(toggle_row)

        self.extra_ticks_cb = self._make_toggle_button(
            "Ticks", checked=False, tooltip="Enable additional minor tick marks (T)"
        )
        self.extra_ticks_cb.toggled.connect(self._on_plot_option_changed)
        toggle_row.addWidget(self.extra_ticks_cb)

        self.precision_cb = self._make_toggle_button(
            "Precision", checked=False, tooltip="Higher tick density for fine inspection"
        )
        self.precision_cb.toggled.connect(self._on_plot_option_changed)
        toggle_row.addWidget(self.precision_cb)

        self.multi_channel_cb = self._make_toggle_button(
            "Multi-ch", checked=False,
            tooltip="Plot extra channel profiles when the source image has more than one",
        )
        self.multi_channel_cb.toggled.connect(self._on_plot_option_changed)
        toggle_row.addWidget(self.multi_channel_cb)

        self._add_toggle_row_separator(toggle_row)

        self.preserve_profiles_cb = self._make_toggle_button(
            "Preserve profiles", checked=True, tooltip="Keep overlays when changing channel"
        )
        self.preserve_profiles_cb.toggled.connect(self._on_preserve_toggle)
        toggle_row.addWidget(self.preserve_profiles_cb)

        toggle_row.addStretch(1)
        controls_layout.addLayout(toggle_row)

        info_layout.addWidget(controls_panel)

        # Real table, not a pipe-delimited string - overlays only earn
        # their keep if they're comparable at a glance. Colour is a cell
        # widget (see _ColorChip) so row selection never washes it out.
        self.profile_table = QtWidgets.QTableWidget(0, 7)
        self.profile_table.setObjectName("profileTable")
        self.profile_table.setHorizontalHeaderLabels(["", "Color", "File", "Channel", "Dir", "Length", "Points"])
        self.profile_table.verticalHeader().setVisible(False)
        self.profile_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.profile_table.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.profile_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.profile_table.setShowGrid(False)
        self.profile_table.setAlternatingRowColors(True)
        header_view = self.profile_table.horizontalHeader()
        header_view.setSectionResizeMode(2, QtWidgets.QHeaderView.Stretch)
        for col in (0, 1, 3, 4, 5, 6):
            header_view.setSectionResizeMode(col, QtWidgets.QHeaderView.ResizeToContents)
        self.profile_table.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.profile_table.itemSelectionChanged.connect(self._on_profile_row_selected)
        self.profile_table.itemChanged.connect(self._on_profile_visibility_changed)
        self.profile_table.customContextMenuRequested.connect(self._on_profile_list_context_menu)

        profiles_header = QtWidgets.QHBoxLayout()
        profiles_header.setContentsMargins(0, 0, 0, 0)
        profiles_header.setSpacing(6)
        profiles_header.addWidget(QtWidgets.QLabel("Profiles"))
        profiles_header.addStretch(1)
        self.compose_drag_btn = _ProfileCompositeDragButton(self)
        profiles_header.addWidget(self.compose_drag_btn)
        info_layout.addLayout(profiles_header)
        info_layout.addWidget(self.profile_table)

        btn_layout = QtWidgets.QHBoxLayout()
        self.copy_btn = QtWidgets.QPushButton('Copy XY')
        self.copy_btn.clicked.connect(self._copy_current_profile)
        btn_layout.addWidget(self.copy_btn)
        self.add_btn = QtWidgets.QPushButton('Add overlay')
        self.add_btn.clicked.connect(self._add_overlay_from_active)
        btn_layout.addWidget(self.add_btn)
        btn_layout.addStretch(1)
        # The destructive action lives apart from the harmless ones, named
        # and confirmed at the point of use (see _delete_selected_profile).
        self.delete_btn = QtWidgets.QPushButton('Delete')
        self.delete_btn.setObjectName("profileDeleteButton")
        self.delete_btn.clicked.connect(self._delete_selected_profile)
        btn_layout.addWidget(self.delete_btn)
        btn_layout.addSpacing(16)
        self.close_btn = QtWidgets.QPushButton('Close')
        self.close_btn.clicked.connect(self.accept)
        btn_layout.addWidget(self.close_btn)
        info_layout.addLayout(btn_layout)
        splitter.addWidget(info_widget)
        # Plot gets ~65% of the vertical space, not about half - it's the
        # reason the window exists.
        splitter.setStretchFactor(0, 65)
        splitter.setStretchFactor(1, 35)
        splitter.setSizes([340, 160])
        root.addWidget(splitter)

    def _build_shortcuts(self):
        self._delete_shortcut = QtWidgets.QShortcut(QtGui.QKeySequence("Delete"), self)
        self._delete_shortcut.activated.connect(self._delete_selected_profile)
        self._delete_shortcut_back = QtWidgets.QShortcut(QtGui.QKeySequence("Backspace"), self)
        self._delete_shortcut_back.activated.connect(self._delete_selected_profile)
        self._delete_list_shortcut = QtWidgets.QShortcut(QtGui.QKeySequence("Delete"), self.profile_table)
        self._delete_list_shortcut.setContext(QtCore.Qt.WidgetShortcut)
        self._delete_list_shortcut.activated.connect(self._delete_selected_profile)
        self._delete_list_shortcut_back = QtWidgets.QShortcut(QtGui.QKeySequence("Backspace"), self.profile_table)
        self._delete_list_shortcut_back.setContext(QtCore.Qt.WidgetShortcut)
        self._delete_list_shortcut_back.activated.connect(self._delete_selected_profile)

        self._copy_plot_shortcut = QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+C"), self)
        self._copy_plot_shortcut.activated.connect(lambda: self._copy_plot("png", dpi=300))
        self._copy_xy_shortcut = QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+Shift+C"), self)
        self._copy_xy_shortcut.activated.connect(self._copy_current_profile)
        self._save_plot_shortcut = QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+S"), self)
        self._save_plot_shortcut.activated.connect(lambda: self._save_plot("png", dpi=300))

    def _dismiss_hint_bar(self):
        self._hint_dismissed = True
        QtCore.QSettings().setValue("profileDialog/hintDismissed", True)
        self.hint_bar.setVisible(False)

    def detach_as_workspace_window(self):
        """Make the dialog an independent top-level window so it does not drag the main viewer to front."""
        owner = getattr(self, "_owner", None)
        try:
            self.setParent(None, self.windowFlags())
            self.setWindowFlag(QtCore.Qt.Window, True)
            self.setWindowModality(QtCore.Qt.NonModal)
            if owner is not None and hasattr(owner, "windowIcon"):
                self.setWindowIcon(owner.windowIcon())
        except Exception:
            pass

    def set_context_source(self, source_canvas, *, dark=None, grid=None):
        # The in-dialog preview panel was removed upstream (duplicated
        # rendering work from the main preview canvas); this is kept as a
        # no-op so existing callers (controllers/profile.py, viewer/
        # measurement.py) don't need hasattr guards.
        self._context_source = None
        return

    # -------------------------------------------------------------- units
    def _distance_axis_params(self, reference, datasets):
        """Resolve the shared distance-axis unit/SI factor for this render.

        Mirrors the old dialog's axis_label_unit derivation: prefer the
        reference dataset's declared axis unit, else the first dataset
        with a physical (x_nm) axis, else fall back to plain pixels.
        """
        axis_unit = None
        if reference and reference.get('x_nm') is not None:
            axis_unit = reference.get('axis_unit') or reference.get('distance_unit') or 'nm'
        else:
            for _label, data, _is_active, _key in datasets:
                if data.get('x_nm') is not None:
                    axis_unit = data.get('axis_unit') or 'nm'
                    break
        if axis_unit is None:
            return 1.0, '', False
        _dummy, si_unit, auto_prefix = to_si_base(np.asarray([1.0]), axis_unit)
        factor = 1e-9 if si_unit == 'm' and axis_unit.strip().lower() == 'nm' else 1.0
        return factor, si_unit, auto_prefix

    def _external_to_plot_x(self, value):
        return float(value) * self._x_si_factor

    def _plot_to_external_x(self, value):
        if not self._x_si_factor:
            return float(value)
        return float(value) / self._x_si_factor

    # -------------------------------------------------------------- data
    def update_profiles(self, active_profile, saved_profiles=None, activate_overlay_callback=None,
                         highlight_overlay_callback=None):
        saved_profiles = saved_profiles or []
        self._active = active_profile
        self._saved = saved_profiles
        entries = self._ordered_profile_entries(active_profile, saved_profiles)
        self._ordered_profile_entries_cache = entries
        if len(saved_profiles) != self._last_saved_count:
            keep = self._marker_positions_by_key.get(None)
            keep_domain = self._marker_domain_by_key.get(None)
            self._marker_positions_by_key = {None: keep} if keep is not None else {}
            self._marker_domain_by_key = {None: keep_domain} if keep_domain is not None else {}
        self._last_saved_count = len(saved_profiles)
        if activate_overlay_callback is not None:
            self._activate_overlay_cb = activate_overlay_callback
        if highlight_overlay_callback is not None:
            self._highlight_overlay_cb = highlight_overlay_callback
        reference = active_profile or (saved_profiles[0] if saved_profiles else None)

        datasets = [(e["label"], e["data"], e["is_active"], e["key"]) for e in entries]
        if self.multi_channel_cb.isChecked() and active_profile and active_profile.get('extra_channels'):
            for extra in active_profile.get('extra_channels', []):
                datasets.append((extra.get('name', 'Extra'), extra, True, None))

        # Remove only the previously plotted curves/legend, not the whole
        # axes: ax.clear() orphans the persistent SpanSelector's artists
        # (confirmed empirically - they drop out of ax.get_children()),
        # breaking the band's rendering on the next redraw.
        for line in list(self._line_handles_by_key.values()):
            try:
                line.remove()
            except Exception:
                pass
        self._line_handles_by_key = {}
        for line in self._right_curves:
            try:
                line.remove()
            except Exception:
                pass
        self._right_curves = []
        if self._legend_item is not None:
            try:
                self._legend_item.remove()
            except Exception:
                pass
            self._legend_item = None
        self.ax_right.set_visible(False)

        if not datasets:
            self.stats.setText("No profile data")
            if hasattr(self, "header_title"):
                self.header_title.setText("Profile measurement — no data")
            self.profile_table.blockSignals(True)
            self.profile_table.setRowCount(0)
            self.profile_table.blockSignals(False)
            self._resize_profile_table()
            self._marker_reference_state = (None, None, None)
            self._reset_region(None, None)
            self._refresh_action_button_states()
            self.canvas.draw_idle()
            return

        self._x_si_factor, self._x_unit, self._x_auto_prefix = self._distance_axis_params(reference, datasets)
        x_label, x_unit_clean = compose_axis_label("Distance", self._x_unit or ('px' if self._x_si_factor == 1.0 and not self._x_unit else self._x_unit))
        if self._x_unit and self._x_auto_prefix:
            self.ax.set_xlabel(x_label)
            self.ax.xaxis.set_major_formatter(EngFormatter(unit=x_unit_clean))
        else:
            self.ax.set_xlabel(f"{x_label} ({x_unit_clean})" if x_unit_clean else (x_label or "Distance (px)"))
            self.ax.xaxis.set_major_formatter(ScalarFormatter())
        self._apply_ylabel(reference)

        show_points = bool(self.show_points_cb.isChecked())
        show_lines = bool(self.show_lines_cb.isChecked())
        ref_unit_si = self._reference_si_unit(reference)

        legend_entries = []
        marker_dataset = active_profile if active_profile else (saved_profiles[0] if saved_profiles else None)
        for label, data, is_active, profile_key in datasets:
            if not self._overlay_visible_by_key.get(profile_key, True):
                continue
            x_raw = data.get('x_nm')
            if x_raw is None:
                x_raw = data.get('x_px')
            y_raw = data.get('vals')
            if x_raw is None or y_raw is None:
                continue
            x_plot = np.asarray(x_raw, dtype=float) * self._x_si_factor
            data_unit = data.get('unit') or ''
            y_plot, y_unit_si, _auto = to_si_base(y_raw, data_unit)
            color = data.get('color') or self._fallback_curve_color(is_active)
            lw = float(data.get('lw') or _DEFAULT_CURVE_WIDTH)
            marker_style = data.get('marker_style') or 'o'
            marker_size = float(data.get('marker_size') or (5.0 if is_active else 4.0))
            plot_label = self._dataset_display_name(data, label)
            on_right = bool(y_unit_si) and y_unit_si != ref_unit_si
            target_ax = self.ax_right if on_right else self.ax
            line, = target_ax.plot(
                x_plot, y_plot, color=color,
                linewidth=lw if show_lines else 0.0,
                linestyle='-' if show_lines else 'None',
                marker=marker_style if show_points else None,
                markersize=marker_size, markerfacecolor=color, markeredgecolor=color,
                label=plot_label,
            )
            if on_right:
                self.ax_right.set_visible(True)
                self.ax_right.set_ylabel(plot_label, color=color)
                self.ax_right.tick_params(axis='y', colors=color)
                self.ax_right.yaxis.set_major_formatter(
                    EngFormatter(unit=y_unit_si) if y_unit_si else ScalarFormatter()
                )
                self._right_curves.append(line)
            self._line_handles_by_key[profile_key] = line
            legend_entries.append((line, plot_label))
        if marker_dataset is not None:
            ref_points = marker_dataset.get('x_nm') if marker_dataset.get('x_nm') is not None else marker_dataset.get('x_px')
            ref_length = marker_dataset.get('length_nm')
        elif datasets:
            data0 = datasets[0][1]
            ref_points = data0.get('x_nm') if data0.get('x_nm') is not None else data0.get('x_px')
            ref_length = data0.get('length_nm')
        else:
            ref_points = None
            ref_length = None

        self.ax.relim()
        self.ax.autoscale_view()
        if self.ax_right.get_visible():
            self.ax_right.relim()
            self.ax_right.autoscale_view()
        if self.extra_ticks_cb.isChecked():
            self.ax.xaxis.set_minor_locator(AutoMinorLocator(4))
            self.ax.yaxis.set_minor_locator(AutoMinorLocator(4))
        if self.precision_cb.isChecked():
            self.ax.xaxis.set_minor_locator(AutoMinorLocator(5))
            self.ax.yaxis.set_minor_locator(AutoMinorLocator(5))

        if len(legend_entries) > 1 and self._legend_visible:
            handles = [h for h, _n in legend_entries]
            labels = [n for _h, n in legend_entries]
            self._legend_item = self.ax.legend(handles, labels, loc='best')
            self._apply_legend_style()

        self._apply_metadata_overlay(reference)
        self._apply_plot_theme()
        self.stats.setText(self._format_stats_text(active_profile, saved_profiles))
        if hasattr(self, "header_title"):
            self.header_title.setText(self._dataset_display_name(reference, "Profile measurement") if reference else "Profile measurement")
        self._populate_profile_list(active_profile, saved_profiles)
        valid_keys = {entry["key"] for entry in entries}
        if self._current_marker_key not in valid_keys:
            self._current_marker_key = None if active_profile is not None else (entries[0]["key"] if entries else None)
        self.select_overlay(self._current_marker_key)
        self._reset_region(ref_points, ref_length, reference_dataset=marker_dataset)
        if self._current_marker_key in self._marker_positions_by_key:
            positions = self._marker_positions_by_key.get(self._current_marker_key)
            domain = self._marker_domain_by_key.get(self._current_marker_key)
            if positions:
                self.set_marker_positions(positions, domain=domain)
        if callable(self._marker_key_cb):
            self._marker_key_cb(self._current_marker_key)
        self._apply_font_scale()
        self._refresh_action_button_states()
        self.canvas.draw_idle()

    def _fallback_curve_color(self, is_active):
        dark = bool(self._dark_background)
        if is_active:
            return _FALLBACK_ACTIVE_COLOR_DARK if dark else _FALLBACK_ACTIVE_COLOR_LIGHT
        return _FALLBACK_OVERLAY_COLOR_DARK if dark else _FALLBACK_OVERLAY_COLOR_LIGHT

    def _reference_si_unit(self, reference):
        unit = (reference.get('unit') if reference else '') or ''
        _values, si_unit, _auto = to_si_base(np.asarray([1.0]), unit)
        return si_unit

    def _apply_ylabel(self, dataset):
        if dataset:
            unit_candidate = dataset.get('unit')
            if unit_candidate:
                self._unit = unit_candidate
        _values, si_unit, auto_prefix = to_si_base(np.asarray([1.0]), self._unit)
        label, unit_clean = compose_axis_label(self._y_label or 'Value', si_unit)
        if unit_clean and auto_prefix:
            self.ax.set_ylabel(label or 'Value')
            self.ax.yaxis.set_major_formatter(EngFormatter(unit=unit_clean))
        else:
            self.ax.set_ylabel(f"{label} ({unit_clean})" if unit_clean else (label or 'Value'))
            self.ax.yaxis.set_major_formatter(ScalarFormatter())

    def _format_stats_text(self, active, saved):
        lines = []
        if active:
            lines.append(self._fmt_length("Active", active.get("length_nm")))
        for idx, data in enumerate(saved, 1):
            lines.append(self._fmt_length(f"Overlay {idx}", data.get("length_nm")))
        return "\n".join(lines) if lines else "No profile data"

    @staticmethod
    def _fmt_length(title, length_nm):
        if length_nm is None:
            return f"{title}: N/A"
        return f"{title}: {length_nm:.3f} nm"

    # ---------------------------------------------------------- ordering
    def _profile_value(self, profile_key, field, default=None):
        dataset = self._dataset_for_profile_key(profile_key) or {}
        # `or default`, not `.get(field, default)`: real datasets from
        # _build_profile_data store an explicit `None` for style fields
        # (lw/marker_size/...) when no override was supplied, so a plain
        # dict.get default (only used when the key is *absent*) would
        # never fire and callers would try to float(None).
        value = dataset.get(field)
        return value if value is not None else default

    def _dataset_for_profile_key(self, profile_key):
        if profile_key is None:
            return self._active
        try:
            idx = int(profile_key)
        except Exception:
            return None
        if 0 <= idx < len(self._saved):
            return self._saved[idx]
        return None

    @staticmethod
    def _profile_id_sort_value(dataset):
        try:
            return str(dataset.get('profile_id') or '')
        except Exception:
            return ''

    def _ordered_profile_entries(self, active_profile, saved_profiles):
        entries = []
        if active_profile:
            entries.append({"label": "Active", "data": active_profile, "is_active": True, "key": None})
        for idx, data in enumerate(saved_profiles or []):
            entries.append({"label": f"Overlay {idx + 1}", "data": data, "is_active": False, "key": idx})
        return entries

    def _live_profile_ref(self, profile_key):
        dataset = self._dataset_for_profile_key(profile_key)
        if not isinstance(dataset, dict):
            return None
        return dataset.get('live_profile_ref')

    # -------------------------------------------------------------- style
    def _apply_profile_style_change(self, profile_key, **changes):
        dataset = self._dataset_for_profile_key(profile_key)
        if not isinstance(dataset, dict):
            return
        applied = False
        if callable(self._style_update_cb):
            try:
                applied = bool(self._style_update_cb(profile_key, **changes))
            except Exception:
                applied = False
        if not applied:
            ref = self._live_profile_ref(profile_key)
            if ref is not None:
                try:
                    applied = bool(apply_live_profile_style(ref, **changes))
                except Exception:
                    applied = False
        dataset.update(changes)
        self.update_profiles(
            self._active, self._saved,
            activate_overlay_callback=self._activate_overlay_cb,
            highlight_overlay_callback=self._highlight_overlay_cb,
        )

    def _apply_palette_change(self, palette_name):
        self._profile_palette_name = palette_name
        applied = False
        if callable(self._palette_cb):
            try:
                applied = bool(self._palette_cb(palette_name))
            except Exception:
                applied = False
        if not applied:
            colors = get_color_cycle(palette_name)
            entries = self._ordered_profile_entries(self._active, self._saved)
            for i, entry in enumerate(entries):
                data = entry["data"]
                if isinstance(data, dict):
                    data["color"] = colors[i % len(colors)]
        self.update_profiles(
            self._active, self._saved,
            activate_overlay_callback=self._activate_overlay_cb,
            highlight_overlay_callback=self._highlight_overlay_cb,
        )

    # ------------------------------------------------------------ toggles
    def _make_toggle_button(self, text, *, checked=False, tooltip=None):
        btn = QtWidgets.QToolButton(self)
        btn.setObjectName("profileToggleButton")
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

    def _add_toggle_row_separator(self, layout):
        line = QtWidgets.QFrame()
        line.setFrameShape(QtWidgets.QFrame.VLine)
        line.setObjectName("profileToggleSeparator")
        layout.addSpacing(4)
        layout.addWidget(line)
        layout.addSpacing(4)

    def _apply_toggle_button_styles(self):
        # Flat, ~26 px pills with a quiet checked state (tinted background,
        # accent text), not saturated fills - nine identical loud pills
        # used to out-compete the plot for attention on every glance.
        p = ui_theme.toggle_pill_palette(bool(self._dark_background))
        style = (
            "QToolButton#profileToggleButton {"
            f"background-color: {p['inactive_bg']};"
            f"color: {p['inactive_text']};"
            f"border: 1px solid {p['inactive_border']};"
            "border-radius: 5px;"
            "padding: 3px 9px;"
            "font-weight: 500;"
            "min-height: 18px;"
            "}"
            "QToolButton#profileToggleButton:checked {"
            f"background-color: {p['active_bg']};"
            f"color: {p['active_text']};"
            f"border: 1px solid {p['active_border']};"
            "font-weight: 600;"
            "}"
            "QToolButton#profileToggleButton:hover {"
            f"border: 1px solid {p['active_border']};"
            "}"
        )
        for btn in self._toggle_buttons:
            try:
                btn.setStyleSheet(style)
            except Exception:
                pass
        sep_style = f"QFrame#profileToggleSeparator {{ color: {p['inactive_border']}; }}"
        for sep in self.findChildren(QtWidgets.QFrame, "profileToggleSeparator"):
            sep.setStyleSheet(sep_style)
        hint = self.findChild(QtWidgets.QLabel, "profileControlsHint")
        if hint is not None:
            hint.setStyleSheet(f"color: {p['hint_color']};")
        compose_btn = getattr(self, "compose_drag_btn", None)
        if compose_btn is not None:
            drop_active = bool(compose_btn.property("dropActive"))
            button_bg = p['compose_drop_bg'] if drop_active else p['compose_button_bg']
            button_border = p['compose_drop_border'] if drop_active else p['compose_button_border']
            button_style = (
                "QToolButton#profileComposeButton {"
                f"background-color: {button_bg};"
                f"color: {p['compose_button_text']};"
                f"border: 1px solid {button_border};"
                "border-radius: 10px;"
                "padding: 4px 10px;"
                "font-weight: 600;"
                "}"
                "QToolButton#profileComposeButton:hover {"
                f"border: 1px solid {p['compose_button_hover_border']};"
                "}"
                "QToolButton#profileComposeButton:disabled {"
                f"background-color: {p['inactive_bg']};"
                f"color: {p['hint_color']};"
                f"border: 1px solid {p['inactive_border']};"
                "}"
            )
            try:
                compose_btn.setStyleSheet(button_style)
            except Exception:
                pass
        # The destructive action is visually distinct (a quiet red outline,
        # not a filled danger button that would out-shout the rest of the
        # bottom bar) and named/confirmed at the point of use, not just
        # colored differently.
        danger = "#e05252" if self._dark_background else "#b3261e"
        delete_btn = getattr(self, "delete_btn", None)
        if delete_btn is not None:
            delete_btn.setStyleSheet(
                "QPushButton#profileDeleteButton {"
                f"color: {danger}; border: 1px solid {danger};"
                "border-radius: 4px; padding: 4px 10px; background: transparent;"
                "}"
                "QPushButton#profileDeleteButton:hover {"
                f"background: {danger}; color: #ffffff;"
                "}"
            )
        table = getattr(self, "profile_table", None)
        if table is not None:
            table.setStyleSheet(
                "QTableWidget#profileTable {"
                f"background: {p.get('list_bg', p['inactive_bg'])}; color: {p['panel_text']};"
                f"border: 1px solid {p['inactive_border']}; border-radius: 6px;"
                "}"
                "QHeaderView::section {"
                f"background: {p['inactive_bg']}; color: {p['hint_color']};"
                "border: none; padding: 3px 6px; font-size: 7.5pt;"
                "}"
                "QTableWidget#profileTable::item:selected {"
                f"background: {p['active_bg']}; color: {p['active_text']};"
                "}"
            )

    # -------------------------------------------------------------- events
    def wheelEvent(self, event):
        try:
            modifiers = event.modifiers()
        except Exception:
            modifiers = QtCore.Qt.NoModifier
        if modifiers & QtCore.Qt.ControlModifier:
            angle = event.angleDelta().y() if hasattr(event, 'angleDelta') else 0
            if angle:
                step = 0.05 * (1 if angle > 0 else -1)
                self._font_scale = min(1.8, max(0.6, self._font_scale + step))
                self._apply_font_scale()
            event.accept()
            return
        super().wheelEvent(event)

    def keyPressEvent(self, event):
        key = event.key()
        if key in (QtCore.Qt.Key_Delete, QtCore.Qt.Key_Backspace):
            self._delete_selected_profile()
            event.accept()
            return
        try:
            mods = event.modifiers()
        except Exception:
            mods = QtCore.Qt.NoModifier
        if key in (QtCore.Qt.Key_Left, QtCore.Qt.Key_Right) and self._markers_enabled:
            self._nudge_region(key, mods)
            event.accept()
            return
        if mods == QtCore.Qt.NoModifier:
            if key == QtCore.Qt.Key_R:
                self.ax.relim()
                self.ax.autoscale_view()
                if self.ax_right.get_visible():
                    self.ax_right.relim()
                    self.ax_right.autoscale_view()
                self.canvas.draw_idle()
                event.accept()
                return
            key_map = {
                QtCore.Qt.Key_M: "measure_toggle",
                QtCore.Qt.Key_S: "snap_toggle",
                QtCore.Qt.Key_C: "crosshair_toggle",
                QtCore.Qt.Key_G: "grid_cb",
                QtCore.Qt.Key_L: "show_lines_cb",
                QtCore.Qt.Key_P: "show_points_cb",
                QtCore.Qt.Key_T: "extra_ticks_cb",
            }
            attr = key_map.get(key)
            if attr and hasattr(self, attr):
                getattr(self, attr).toggle()
                event.accept()
                return
        super().keyPressEvent(event)

    def _nudge_region(self, key, mods):
        if self._span is None or self._px_size_plot <= 0:
            return
        step = self._px_size_plot * (10 if mods & QtCore.Qt.ControlModifier else 1)
        step = -step if key == QtCore.Qt.Key_Left else step
        lo, hi = self._span.extents
        if mods & QtCore.Qt.ShiftModifier:
            hi += step
        else:
            lo += step
        if lo > hi:
            lo, hi = hi, lo
        self._span.extents = (lo, hi)
        self._sync_marker_positions_from_region()
        self._remember_marker_positions()
        self._update_readout()
        self._notify_marker_positions()
        self.canvas.draw_idle()

    def resizeEvent(self, event):
        super().resizeEvent(event)

    def _font_style_state(self):
        return {
            "bold": bool(getattr(self, "_plot_font_bold", False)),
            "italic": bool(getattr(self, "_plot_font_italic", False)),
            "underline": bool(getattr(self, "_plot_font_underline", False)),
        }

    def _apply_font_scale(self):
        # Axis/tick text is deliberately the loudest type in the window -
        # the previous dialog had it the other way round, with chrome
        # (toggle buttons) typographically louder than the data it framed.
        scale = max(0.6, min(1.8, getattr(self, '_font_scale', 1.0)))
        label_size = 11.5 * scale
        tick_size = 10.0 * scale
        try:
            self.ax.tick_params(axis='both', labelsize=tick_size)
            self.ax_right.tick_params(axis='both', labelsize=tick_size)
            self.ax.xaxis.label.set_fontsize(label_size)
            self.ax.yaxis.label.set_fontsize(label_size)
            self.ax_right.yaxis.label.set_fontsize(label_size)
            style = self._font_style_state()
            for text in (self.ax.xaxis.label, self.ax.yaxis.label, self.ax_right.yaxis.label):
                apply_text_style(text, family=self._plot_font_family, **style)
            for text in list(self.ax.get_xticklabels()) + list(self.ax.get_yticklabels()) + list(self.ax_right.get_yticklabels()):
                apply_text_style(text, family=self._plot_font_family, **style)
            self._apply_legend_style()
        except Exception:
            pass
        for widget in (self.stats, self.marker_info):
            font = widget.font()
            font.setPointSizeF(max(6.5, 8.3 * scale))
            font = apply_qfont_style(font, family=self._plot_font_family, **self._font_style_state())
            widget.setFont(font)
        font = self.profile_table.font()
        font.setPointSizeF(max(7.0, 9.0 * scale))
        font = apply_qfont_style(font, family=self._plot_font_family, **self._font_style_state())
        self.profile_table.setFont(font)
        self._resize_profile_table()
        for btn in self._toggle_buttons:
            font = btn.font()
            font.setPointSizeF(max(6.5, 8.0 * scale))
            font = apply_qfont_style(font, family=self._plot_font_family, **self._font_style_state())
            btn.setFont(font)
        for label in self.readout.value_labels.values():
            font = label.font()
            font.setPointSizeF(max(7.0, 9.0 * scale))
            font = apply_qfont_style(font, family=self._plot_font_family, **self._font_style_state())
            label.setFont(font)
        self.canvas.draw_idle()
        self._update_readout()
        if callable(self._label_scale_cb):
            self._label_scale_cb(self._font_scale)

    def eventFilter(self, source, event):
        if source == self.canvas:
            etype = event.type()
            if etype == QtCore.QEvent.MouseButtonPress and event.button() == QtCore.Qt.LeftButton:
                if self._qt_pos_in_main_axes(event.pos()):
                    self._canvas_drag_start_pos = None
                    self._canvas_drag_started = False
                else:
                    self._canvas_drag_start_pos = event.pos()
                    self._canvas_drag_started = False
            elif (
                etype == QtCore.QEvent.MouseMove
                and self._canvas_drag_start_pos is not None
                and event.buttons() & QtCore.Qt.LeftButton
            ):
                if (event.pos() - self._canvas_drag_start_pos).manhattanLength() >= QtWidgets.QApplication.startDragDistance():
                    self._canvas_drag_start_pos = None
                    self._canvas_drag_started = True
                    try:
                        self.start_profile_composite_drag()
                    except Exception:
                        pass
                    return True
            elif etype == QtCore.QEvent.MouseButtonRelease and event.button() == QtCore.Qt.LeftButton:
                self._canvas_drag_start_pos = None
                self._canvas_drag_started = False
        return super().eventFilter(source, event)

    def _qt_pos_in_main_axes(self, pos):
        if pos is None:
            return False
        try:
            bbox = self.ax.get_window_extent()
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

    # ----------------------------------------------------- measurement band
    def _clamp_external(self, val):
        lo, hi = self._marker_domain
        return min(max(val, lo), hi)

    def _sync_marker_positions_from_region(self):
        if self._span is None:
            self._marker_positions = []
            return
        lo_plot, hi_plot = self._span.extents
        self._marker_positions = [self._plot_to_external_x(lo_plot), self._plot_to_external_x(hi_plot)]

    def _remember_marker_positions(self):
        if self._marker_positions:
            self._marker_saved_positions = list(self._marker_positions)

    def _clear_region(self, reset_saved=True):
        if self._span is not None:
            self._span.set_visible(False)
        self._marker_positions = []
        if reset_saved:
            self._marker_saved_positions = None
        self._ref_x_plot = None
        self._ref_y_plot = None
        self._ref_y_unit_si = ''
        self._ref_peak_idx = np.array([], dtype=int)
        self._px_size_plot = 0.0
        self.readout.clear_all()
        self.marker_info.setText("Measure: N/A" if self._markers_enabled else "Measure off")
        self._notify_marker_positions()
        self.canvas.draw_idle()

    def _reset_region(self, ref_points, ref_length, reference_dataset=None, store_state=True):
        if store_state:
            self._marker_reference_state = (ref_points, ref_length, reference_dataset)
        if not self._markers_enabled or ref_points is None or len(ref_points) == 0:
            self._clear_region(reset_saved=store_state)
            return
        xmin = float(np.nanmin(ref_points))
        xmax = float(np.nanmax(ref_points))
        if not np.isfinite(xmin) or not np.isfinite(xmax) or xmax == xmin:
            self._clear_region(reset_saved=store_state)
            return
        self._marker_domain = (xmin, xmax)
        span = xmax - xmin
        if self._marker_saved_positions and len(self._marker_saved_positions) == 2:
            raw_positions = [self._clamp_external(p) for p in self._marker_saved_positions]
        else:
            raw_positions = [xmin + 0.3 * span, xmin + 0.7 * span]

        # Reference arrays/peaks/pixel-size in PLOT-space (SI base) - the
        # readout, snap-to-peak and keyboard-nudge step size all operate
        # here, not in the external nm/px space used for the canvas link.
        ref_x_plot = np.asarray(ref_points, dtype=float) * self._x_si_factor
        self._ref_x_plot = ref_x_plot
        ref_vals = (reference_dataset or {}).get('vals')
        if ref_vals is not None:
            y_plot, y_unit_si, _auto = to_si_base(ref_vals, (reference_dataset or {}).get('unit'))
            self._ref_y_plot = y_plot
            self._ref_y_unit_si = y_unit_si
        else:
            self._ref_y_plot = None
            self._ref_y_unit_si = ''
        self._px_size_plot = float(ref_x_plot[1] - ref_x_plot[0]) if ref_x_plot.size >= 2 else 0.0
        self._ref_peak_idx = (
            find_peaks(self._ref_y_plot) if self._ref_y_plot is not None and self._ref_y_plot.size >= 3
            else np.array([], dtype=int)
        )

        lo_plot = self._external_to_plot_x(min(raw_positions))
        hi_plot = self._external_to_plot_x(max(raw_positions))
        self._span.extents = (lo_plot, hi_plot)
        self._span.set_visible(True)
        self._span.set_active(True)
        self._snap_region()
        self._sync_marker_positions_from_region()
        self._remember_marker_positions()
        self._update_readout()
        self._notify_marker_positions()
        self.canvas.draw_idle()

    def _on_measure_toggle(self, checked):
        self._markers_enabled = bool(checked)
        QtCore.QSettings().setValue("profileDialog/measureEnabled", self._markers_enabled)
        self.readout.setVisible(self._markers_enabled)
        self.marker_info.setVisible(not self._markers_enabled)
        if self._span is not None:
            self._span.set_active(self._markers_enabled)
        if not self._markers_enabled:
            self._clear_region(reset_saved=False)
            return
        ref_points, ref_length, ref_dataset = self._marker_reference_state
        if ref_points is None:
            self._clear_region(reset_saved=False)
        else:
            self._reset_region(ref_points, ref_length, ref_dataset, store_state=False)

    def _on_snap_toggle(self, checked):
        self._snap_enabled = bool(checked)
        QtCore.QSettings().setValue("profileDialog/snapToPeaks", self._snap_enabled)
        if self._snap_enabled:
            self._snap_region()
            self._sync_marker_positions_from_region()
            self._update_readout()
            self._notify_marker_positions()
            self.canvas.draw_idle()

    def _on_crosshair_toggle(self, checked):
        self._crosshair_enabled = bool(checked)
        QtCore.QSettings().setValue("profileDialog/crosshairEnabled", self._crosshair_enabled)
        if not self._crosshair_enabled and self._vline is not None:
            self._vline.set_visible(False)
            self._hline.set_visible(False)
            self.readout.set("cursor", "-")
            self.canvas.draw_idle()

    def _on_span_move(self, _vmin, _vmax):
        """Live update while dragging - mirrors the old sigRegionChanged path."""
        self._update_readout()
        self._notify_marker_positions()

    def _on_span_select(self, vmin, vmax):
        """Drag finished (mouse release) - mirrors sigRegionChangeFinished.

        Unlike the old LinearRegionItem, SpanSelector has no built-in
        drag-range clamp, so the domain bound is applied here rather than
        via a setBounds() equivalent.
        """
        lo_bound = self._external_to_plot_x(self._marker_domain[0])
        hi_bound = self._external_to_plot_x(self._marker_domain[1])
        lo, hi = (vmin, vmax) if vmin <= vmax else (vmax, vmin)
        lo = max(lo_bound, min(lo, hi_bound))
        hi = max(lo_bound, min(hi, hi_bound))
        if (lo, hi) != (min(vmin, vmax), max(vmin, vmax)):
            self._span.extents = (lo, hi)
        self._sync_marker_positions_from_region()
        self._remember_marker_positions()
        self._snap_region()
        self._update_readout()
        self._notify_marker_positions()
        self.canvas.draw_idle()

    def _snap_region(self):
        """Snap both band edges to the nearest peak within a few pixels.

        ``SpanSelector.extents = ...`` (unlike a signal-driven widget)
        never re-triggers onselect/onmove_callback on its own (confirmed
        empirically), so - unlike the pyqtgraph version this replaced -
        there is no recursion risk in calling this from ``_on_span_select``.
        """
        if not self._snap_enabled or self._span is None:
            return
        if self._ref_x_plot is None or self._ref_peak_idx.size == 0 or self._px_size_plot <= 0:
            return
        lo, hi = self._span.extents
        peak_x = self._ref_x_plot[self._ref_peak_idx]
        new_lo, new_hi = snap_edges((lo, hi), peak_x, tol_px=6, px_size=self._px_size_plot)
        if abs(new_lo - lo) < 1e-15 and abs(new_hi - hi) < 1e-15:
            return
        self._span.extents = (new_lo, new_hi)

    def _fmt_x(self, value):
        d = self._readout_precision
        if self._x_unit:
            return si_format(value, unit=self._x_unit, precision=d)
        return f"{value:.{d}g} px"

    def _fmt_y(self, value, unit):
        d = self._readout_precision
        if unit:
            return si_format(value, unit=unit, precision=d)
        return f"{value:.{d}g}"

    def _update_readout(self):
        self._sync_marker_positions_from_region()
        if not self._markers_enabled or self._span is None or len(self._marker_positions) < 2:
            self.readout.clear_all()
            return
        lo_ext, hi_ext = self._marker_positions
        lo_plot = self._external_to_plot_x(lo_ext)
        hi_plot = self._external_to_plot_x(hi_ext)
        y_unit = self._ref_y_unit_si or ''
        have_ref = self._ref_x_plot is not None and self._ref_y_plot is not None
        stats = band_statistics(self._ref_x_plot, self._ref_y_plot, lo_plot, hi_plot) if have_ref else {"n": 0}

        self.readout.set("left", f"{self._fmt_x(lo_plot)}, {self._fmt_y(stats.get('y_left', float('nan')), y_unit)}" if have_ref else self._fmt_x(lo_plot))
        self.readout.set("right", f"{self._fmt_x(hi_plot)}, {self._fmt_y(stats.get('y_right', float('nan')), y_unit)}" if have_ref else self._fmt_x(hi_plot))
        # Precision tied to sampling: the tolerance is half a pixel, so Δd
        # can never claim more precision than the data actually supports
        # (the old dialog's flat ".3f" could show "79.328 nm" on ~10 nm
        # sampling - three digits the measurement never earned).
        half_px = self._px_size_plot / 2.0
        if self._x_unit:
            self.readout.set("dx", fmt_tol(hi_plot - lo_plot, half_px, self._x_unit, digits=self._readout_precision))
        else:
            self.readout.set("dx", f"{hi_plot - lo_plot:.{self._readout_precision}g} ± {half_px:.2g} px")
        if have_ref and 'y_left' in stats and 'y_right' in stats:
            self.readout.set("dy", self._fmt_y(stats['y_right'] - stats['y_left'], y_unit))
        else:
            self.readout.set("dy", "-")

        if have_ref and 'mean' in stats:
            self.readout.set("mean", self._fmt_y(stats['mean'], y_unit))
            self.readout.set("rms", self._fmt_y(stats['rms'], y_unit))
            # Per-nanometre reads far better than the per-metre (e.g. A/m)
            # that plain SI scaling of a metres-based x-axis would give.
            if self._x_unit == 'm':
                slope_per_nm = stats['slope'] * 1e-9
                self.readout.set("slope", self._fmt_y(slope_per_nm, y_unit) + " / nm")
            else:
                self.readout.set("slope", f"{stats['slope']:.3g} {y_unit}/px".strip())
            peaks_n = peaks_in_range(self._ref_x_plot, self._ref_peak_idx, lo_plot, hi_plot)
            self.readout.set("peaks", str(peaks_n))
        else:
            for key in ("mean", "rms", "slope", "peaks"):
                self.readout.set(key, "-")
        n = stats.get("n", 0)
        px_text = self._fmt_x(self._px_size_plot) if self._px_size_plot > 0 else "-"
        self.readout.set("px", f"{px_text}  (n = {n} in band)")

    def _on_mouse_moved(self, event):
        inside = (
            self._crosshair_enabled and event.inaxes is self.ax
            and event.xdata is not None and event.ydata is not None
        )
        if self._vline is not None:
            self._vline.set_visible(inside)
            self._hline.set_visible(inside)
        if not inside:
            self.readout.set("cursor", "-")
            self.canvas.draw_idle()
            return
        self._vline.set_xdata([event.xdata, event.xdata])
        self._hline.set_ydata([event.ydata, event.ydata])
        y_unit = self._ref_y_unit_si or self._reference_si_unit(self._active or (self._saved[0] if self._saved else None))
        self.readout.set("cursor", f"{self._fmt_x(event.xdata)}, {self._fmt_y(event.ydata, y_unit)}")
        self.canvas.draw_idle()

    def set_marker_positions(self, positions, domain=None):
        if self._marker_syncing:
            return
        try:
            self._marker_syncing = True
            if positions is None or len(positions) < 2:
                self._marker_saved_positions = None
                if self._current_marker_key in self._marker_positions_by_key:
                    self._marker_positions_by_key.pop(self._current_marker_key, None)
                    self._marker_domain_by_key.pop(self._current_marker_key, None)
                self._clear_region(reset_saved=False)
                return
            if domain is not None:
                self._marker_domain = tuple(domain)
            clamped = [self._clamp_external(p) for p in positions]
            if self._current_marker_key is not None:
                self._marker_positions_by_key[self._current_marker_key] = list(clamped)
                self._marker_domain_by_key[self._current_marker_key] = tuple(self._marker_domain)
            else:
                self._marker_positions_by_key[None] = list(clamped)
                self._marker_domain_by_key[None] = tuple(self._marker_domain)
            if self._span is None:
                return
            self._span.extents = (self._external_to_plot_x(min(clamped)), self._external_to_plot_x(max(clamped)))
            self._span.set_visible(self._markers_enabled)
            self._update_readout()
            self.canvas.draw_idle()
        finally:
            self._marker_syncing = False

    def _notify_marker_positions(self):
        if self._marker_syncing:
            return
        self._sync_marker_positions_from_region()
        if not callable(self._marker_update_cb):
            return
        if not self._markers_enabled or len(self._marker_positions) < 2:
            self._marker_update_cb(None, None)
            return
        key = self._current_marker_key
        self._marker_positions_by_key[key] = list(self._marker_positions)
        self._marker_domain_by_key[key] = tuple(self._marker_domain)
        self._marker_update_cb(list(self._marker_positions), tuple(self._marker_domain))

    def set_marker_update_callback(self, cb):
        self._marker_update_cb = cb

    def set_marker_select_callback(self, cb):
        self._marker_key_cb = cb

    def set_label_scale_callback(self, cb):
        self._label_scale_cb = cb
        if callable(self._label_scale_cb):
            self._label_scale_cb(self._font_scale)

    def set_add_overlay_callback(self, cb):
        self._add_overlay_cb = cb
        self._refresh_action_button_states()

    def set_delete_overlay_callback(self, cb):
        self._delete_overlay_cb = cb
        self._refresh_action_button_states()

    def set_style_update_callback(self, cb):
        self._style_update_cb = cb

    def set_palette_callback(self, cb):
        self._palette_cb = cb

    def set_preserve_profiles_callback(self, cb, *, enabled=None):
        self._preserve_cb = cb
        if enabled is not None:
            try:
                self.preserve_profiles_cb.setChecked(bool(enabled))
            except Exception:
                pass

    def _on_plot_option_changed(self, _checked=False):
        self.update_profiles(self._active, self._saved)

    def _on_preserve_toggle(self, checked):
        if callable(self._preserve_cb):
            try:
                self._preserve_cb(bool(checked))
            except Exception:
                pass

    # ------------------------------------------------------------- theme
    def _on_grid_toggled(self, _checked=False):
        self._apply_plot_theme()

    def _resolve_dark_background(self):
        """Dark/light for the current _theme_mode.

        "system" here means "the owning app's current theme," not the OS -
        this app already has its own Light/Dark/Amber theme system the
        user controls from the main window (see gui/theme.py), so that is
        what a per-dialog "Follow system" choice should track. Falls back
        to the dark_mode hint passed at construction (or the OS palette)
        only when there is no owning viewer to read from - e.g. a
        standalone/test dialog.
        """
        if self._theme_mode == "light":
            return False
        if self._theme_mode == "dark":
            return True
        owner = getattr(self, "_owner", None)
        if owner is not None and hasattr(owner, "dark_mode"):
            try:
                return bool(owner.dark_mode)
            except Exception:
                pass
        if self._owner_dark_mode_hint is not None:
            return bool(self._owner_dark_mode_hint)
        try:
            pal = QtWidgets.QApplication.palette()
            return pal.color(QtGui.QPalette.Window).lightness() < 128
        except Exception:
            return False

    def _set_theme_mode(self, mode):
        mode = mode if mode in ("light", "dark", "system") else "system"
        self._theme_mode = mode
        QtCore.QSettings().setValue("profileDialog/themeMode", mode)
        self._dark_background = self._resolve_dark_background()
        self._apply_plot_theme()

    def _populate_view_menu(self):
        menu = self._view_menu
        menu.clear()
        group = QtWidgets.QActionGroup(self)
        group.setExclusive(True)
        for label, mode in (("Light", "light"), ("Dark", "dark"), ("Follow system", "system")):
            act = menu.addAction(label)
            act.setCheckable(True)
            act.setChecked(self._theme_mode == mode)
            act.triggered.connect(lambda _=False, m=mode: self._set_theme_mode(m))
            group.addAction(act)

    def _plot_chrome(self):
        """Plot chrome colors for the current dark flag + app theme.

        Chrome only (backgrounds, ticks, labels, markers, legend frame) -
        the profile trace colors themselves are user/palette-driven and
        never touched here. ``accent`` is for measurement markers/labels.
        """
        dark = bool(self._dark_background)
        if dark and ui_theme.current_theme() == ui_theme.THEME_AMBER:
            c = ui_theme.mpl_chrome_colors(ui_theme.THEME_AMBER)
            return {
                "fig_face": c["fig_face"],
                "ax_face": c["ax_face"],
                "text": c["text"],
                "grid": c["grid"],
                "accent": ui_theme.AMBER["amber_bright"],
                "box_face": ui_theme.AMBER["panel_bg"],
            }
        return {
            "fig_face": '#111217' if dark else '#ffffff',
            "ax_face": '#14161c' if dark else '#ffffff',
            "text": '#f5f5f5' if dark else '#111111',
            "grid": '#4f5a64' if dark else '#b0b0b0',
            "accent": '#f5f5f5' if dark else '#202020',
            "box_face": '#111111' if dark else '#ffffff',
        }

    def _apply_plot_theme(self):
        chrome = self._plot_chrome()
        self.canvas.figure.set_facecolor(chrome["fig_face"])
        grid_on = bool(self.grid_cb.isChecked()) if hasattr(self, 'grid_cb') else False
        try:
            self.ax.set_facecolor(chrome["ax_face"])
            self.ax.tick_params(colors=chrome["text"], labelcolor=chrome["text"])
            self.ax.xaxis.label.set_color(chrome["text"])
            self.ax.yaxis.label.set_color(chrome["text"])
            for spine in self.ax.spines.values():
                spine.set_color(chrome["text"])
            # The right axis (dual-unit overlay) carries its curve's own
            # accent color, set in update_profiles right after each curve
            # is plotted - only its background/spine are generic chrome,
            # never its label/tick color, or every theme refresh would
            # stomp the color that ties it back to its curve.
            self.ax_right.set_facecolor(chrome["ax_face"])
            if grid_on:
                self.ax.grid(True, color=chrome["grid"], alpha=0.35)
            else:
                self.ax.grid(False)
        except Exception:
            pass
        if self._metadata_item is not None:
            try:
                self._metadata_item.set_color(chrome["text"])
                patch = self._metadata_item.get_bbox_patch()
                if patch is not None:
                    patch.set_facecolor(chrome["box_face"])
            except Exception:
                pass
        self._apply_legend_style()
        self._apply_toggle_button_styles()
        if self._span is not None:
            try:
                band = QtGui.QColor(chrome["accent"])
                self._span.set_props(facecolor=(band.redF(), band.greenF(), band.blueF(), 0.22))
                self._span.set_handle_props(color=chrome["accent"])
            except Exception:
                pass
        if self._vline is not None:
            for line in (self._vline, self._hline):
                line.set_color(chrome["text"])
        if hasattr(self, "readout"):
            self.readout.apply_theme(
                text_color=chrome["text"], muted_color=chrome.get("grid", chrome["text"]),
                panel_color=chrome["ax_face"], border_color=chrome["grid"],
            )
        if hasattr(self, "toast"):
            self.toast.apply_theme(text_color="#ffffff", bg_color="rgba(30,32,36,235)")
        if hasattr(self, "hint_bar"):
            self.hint_bar.setStyleSheet(
                "QFrame#profileHintBar {"
                f"background: {chrome['ax_face']}; border: 1px solid {chrome['grid']};"
                "border-radius: 6px; }"
                f"QLabel#profileHintLabel {{ color: {chrome['text']}; font-size: 8.5pt; }}"
                f"QToolButton#profileHintClose {{ color: {chrome['text']}; font-size: 12pt; border: none; }}"
            )
        if hasattr(self, "header_title"):
            self.header_title.setStyleSheet(f"color: {chrome['text']}; font-size: 10.5pt; font-weight: 600;")
        if hasattr(self, "export_btn"):
            accent = QtGui.QColor(chrome["accent"])
            accent_text = "#111111" if accent.lightnessF() > 0.55 else "#ffffff"
            self.export_btn.setStyleSheet(
                "QToolButton#profileExportButton {"
                f"background-color: {chrome['accent']}; color: {accent_text};"
                "border: none; border-radius: 5px; padding: 5px 12px; font-weight: 600;"
                "}"
                "QToolButton#profileExportButton::menu-button {"
                "width: 18px; border-left: 1px solid rgba(255,255,255,90);"
                "}"
                "QToolButton#profileExportButton:hover { background-color: " + chrome["accent"] + "; }"
            )
        self.canvas.draw_idle()

    def _apply_legend_style(self):
        if self._legend_item is None:
            return
        chrome = self._plot_chrome()
        try:
            self._legend_item.set_visible(bool(self._legend_visible))
            frame = self._legend_item.get_frame()
            frame.set_facecolor(chrome["box_face"])
            frame.set_alpha(0.88)
            frame.set_edgecolor(chrome["text"])
            for txt in self._legend_item.get_texts():
                txt.set_color(chrome["text"])
                txt.set_fontsize(self._legend_fontsize * getattr(self, '_font_scale', 1.0))
                apply_text_style(txt, family=self._plot_font_family, **self._font_style_state())
        except Exception:
            pass

    # --------------------------------------------------------- metadata
    def _metadata_text(self, dataset):
        if not self._metadata_visible or not dataset:
            return ""
        lines = []
        if self._metadata_show_filename:
            name = str(dataset.get("source_file_name") or "").strip()
            if not name:
                path_text = str(dataset.get("source_path") or "").strip()
                if path_text:
                    try:
                        name = Path(path_text).name
                    except Exception:
                        name = path_text
            if name:
                lines.append(f"File: {name}")
        if self._metadata_show_acquisition:
            acq = str(dataset.get("source_acquisition_text") or dataset.get("source_title") or "").strip()
            if acq:
                lines.append(f"Acq: {acq}")
        if self._metadata_show_time:
            when = str(dataset.get("source_datetime") or "").strip()
            if not when:
                source_date = str(dataset.get("source_date") or "").strip()
                source_time = str(dataset.get("source_time") or "").strip()
                when = f"{source_date} {source_time}".strip() if (source_date and source_time) else (source_date or source_time)
            if when:
                lines.append(f"Time: {when}")
        if self._metadata_show_folder_name:
            folder_name = str(dataset.get("source_folder_name") or "").strip()
            if not folder_name:
                path_text = str(dataset.get("source_path") or "").strip()
                if path_text:
                    try:
                        folder_name = Path(path_text).parent.name
                    except Exception:
                        folder_name = ""
            if folder_name:
                lines.append(f"Folder name: {folder_name}")
        if self._metadata_show_folder:
            folder = str(dataset.get("source_folder") or "").strip()
            if not folder:
                path_text = str(dataset.get("source_path") or "").strip()
                if path_text:
                    try:
                        folder = str(Path(path_text).parent)
                    except Exception:
                        folder = ""
            if folder:
                lines.append(f"Folder: {folder}")
        return "\n".join(lines)

    def _apply_metadata_overlay(self, dataset):
        if self._metadata_item is not None:
            try:
                self._metadata_item.remove()
            except Exception:
                pass
            self._metadata_item = None
        text = self._metadata_text(dataset)
        if not text:
            return
        chrome = self._plot_chrome()
        try:
            # Figure-fraction coordinates (not axes-data coordinates): the
            # corner position stays fixed regardless of pan/zoom, with no
            # need to reposition on every view-range change.
            self._metadata_item = self.canvas.figure.text(
                0.02, 0.985, text, ha="left", va="top",
                fontsize=max(5.0, 6.0 * getattr(self, "_font_scale", 1.0)),
                color=chrome["text"],
                bbox={"facecolor": chrome["box_face"], "alpha": 0.72, "edgecolor": "none", "pad": 2.0},
            )
            apply_text_style(self._metadata_item, family=self._plot_font_family, **self._font_style_state())
        except Exception:
            self._metadata_item = None

    # ----------------------------------------------------------- export
    class _LightForExport:
        """Exports are always light, whatever the UI theme is set to.

        Also hides the measurement band/crosshair unless the user opted in
        via "Include measurement band in exports" - a measuring instrument
        is not data and should not appear in a figure unless asked for.
        Both concerns share one context manager so every export path
        (copy PNG/SVG, save PNG/SVG/PDF) gets both for free.
        """

        def __init__(self, dlg):
            self.dlg = dlg
            self.prev_dark = None
            self.prev_band_visible = None
            self.prev_crosshair_visible = None

        def __enter__(self):
            dlg = self.dlg
            self.prev_dark = dlg._dark_background
            if self.prev_dark:
                dlg._dark_background = False
                dlg._apply_plot_theme()
            if not dlg._include_band_in_export:
                if dlg._span is not None:
                    self.prev_band_visible = dlg._span.get_visible()
                    dlg._span.set_visible(False)
                if dlg._vline is not None:
                    self.prev_crosshair_visible = (dlg._vline.get_visible(), dlg._hline.get_visible())
                    dlg._vline.set_visible(False)
                    dlg._hline.set_visible(False)
            if self.prev_dark:
                QtWidgets.QApplication.processEvents()

        def __exit__(self, *_a):
            dlg = self.dlg
            if self.prev_band_visible is not None and dlg._span is not None:
                dlg._span.set_visible(self.prev_band_visible)
            if self.prev_crosshair_visible is not None and dlg._vline is not None:
                dlg._vline.set_visible(self.prev_crosshair_visible[0])
                dlg._hline.set_visible(self.prev_crosshair_visible[1])
            if self.prev_dark:
                dlg._dark_background = True
                dlg._apply_plot_theme()

    def _copy_plot(self, fmt, *, dpi=300):
        # Same savefig-based logic as the shared figure_layout_presets.py
        # helpers (copy_figure_to_clipboard/save_figure_with_dialog), just
        # inlined so the confirmation is this dialog's own styled Toast
        # rather than their plain QToolTip - kept local rather than adding
        # a "silent" mode to a module other dialogs also use unchanged.
        import io
        with self._LightForExport(self):
            buf = io.BytesIO()
            if fmt == "svg":
                with matplotlib.rc_context({"svg.fonttype": "none"}):
                    self.canvas.figure.savefig(buf, format="svg", bbox_inches="tight", pad_inches=0.02)
                mime = QtCore.QMimeData()
                data = buf.getvalue()
                mime.setData("image/svg+xml", data)
                try:
                    mime.setText(data.decode("utf-8"))
                except Exception:
                    pass
                QtWidgets.QApplication.clipboard().setMimeData(mime)
                self.toast.show_message("Copied SVG markup to the clipboard")
                return
            self.canvas.figure.savefig(buf, format="png", dpi=int(dpi), bbox_inches="tight", pad_inches=0.02)
            image = QtGui.QImage.fromData(buf.getvalue(), "PNG")
        QtWidgets.QApplication.clipboard().setImage(image)
        self.toast.show_message(f"Copied PNG  {image.width()} × {image.height()} px  —  {int(dpi)} dpi")

    def _save_plot(self, fmt, *, dpi=300):
        default_name = f"profile_measurement.{fmt}"
        filt = {"svg": "SVG Files (*.svg)", "pdf": "PDF Files (*.pdf)"}.get(fmt, "PNG Files (*.png)")
        path, _sel = QtWidgets.QFileDialog.getSaveFileName(self, "Save plot", default_name, filt)
        if not path:
            return
        with self._LightForExport(self):
            try:
                if fmt == "svg":
                    with matplotlib.rc_context({"svg.fonttype": "none"}):
                        self.canvas.figure.savefig(path, format="svg", bbox_inches="tight", pad_inches=0.02)
                elif fmt == "pdf":
                    with matplotlib.rc_context({"pdf.fonttype": 42, "ps.fonttype": 42}):
                        self.canvas.figure.savefig(path, format="pdf", bbox_inches="tight", pad_inches=0.02)
                else:
                    self.canvas.figure.savefig(path, format="png", dpi=int(dpi), bbox_inches="tight", pad_inches=0.02)
            except Exception as exc:
                QtWidgets.QMessageBox.warning(self, "Save plot", str(exc))
                return
        self.toast.show_message(f"Saved {Path(path).name}")

    def _export_px(self, dpi):
        preset = _get_profile_figure_preset(self._figure_preset_key)
        width_mm = preset.width_mm if preset.key != "interactive" else 152.4
        return int(round(width_mm / 25.4 * dpi))

    def _provenance_header_lines(self, dataset):
        """File/channel/direction/timestamp/units/pixel-size/sampling/measurement.

        Shared by Copy XY and Save CSV so the two can't drift apart - a
        profile without its provenance is unusable a year later, and that
        provenance should mean the same thing wherever it's written.
        """
        dataset = dataset or {}
        meta = dict(dataset.get('meta') or {})
        channel_unit = dataset.get('unit') or ""
        channel_label, channel_unit_clean = compose_axis_label(
            self._y_label or meta.get('channel') or "Value", channel_unit
        )
        lines = [f"Channel: {channel_label}{f' [{channel_unit_clean}]' if channel_unit_clean else ''}"]
        file_name = dataset.get('source_file_name') or meta.get('file_name')
        if file_name:
            lines.append(f"File: {file_name}")
        direction = meta.get('direction') or dataset.get('direction')
        if direction:
            lines.append(f"Direction: {direction}")
        timestamp = dataset.get('source_datetime') or meta.get('datetime')
        if not timestamp:
            date_part = dataset.get('source_date') or meta.get('date')
            time_part = dataset.get('source_time') or meta.get('time')
            if date_part or time_part:
                timestamp = f"{date_part or ''} {time_part or ''}".strip()
        if timestamp:
            lines.append(f"Timestamp: {timestamp}")
        # _build_profile_data (detail_preview_canvas.py) always samples via
        # bilinear interpolation along the drawn line - stated explicitly
        # rather than read from a field, since the dataset dict carries no
        # "sampling mode" key of its own.
        lines.append("Sampling: bilinear interpolation along the drawn line")
        x = dataset.get('x_nm')
        unit = dataset.get('axis_unit') or dataset.get('distance_unit') or 'nm'
        if x is None:
            x = dataset.get('x_px')
            unit = 'px'
        if x is not None and len(x) >= 2:
            px_size = float(x[1]) - float(x[0])
            lines.append(f"Pixel size: {px_size:.6g} {unit}")
        if self._markers_enabled and len(self._marker_positions) >= 2 and self._px_size_plot > 0:
            lo_ext, hi_ext = self._marker_positions
            dx_plot = abs(self._external_to_plot_x(hi_ext) - self._external_to_plot_x(lo_ext))
            half_px = self._px_size_plot / 2.0
            measurement = (
                fmt_tol(dx_plot, half_px, self._x_unit, digits=self._readout_precision)
                if self._x_unit else f"{dx_plot:.{self._readout_precision}g} px"
            )
            lines.append(f"Measurement: Δd = {measurement}")
        return lines

    def _copy_current_profile(self):
        datasets = []
        if self._active:
            datasets.append(("Active", self._active))
        for idx, data in enumerate(self._saved, 1):
            datasets.append((f"Overlay {idx}", data))
        if not datasets:
            QtWidgets.QMessageBox.information(self, "Copy profile", "No profile data available.")
            return
        reference_dataset = self._active or (self._saved[0] if self._saved else {})
        meta = (reference_dataset or {}).get('meta') or {}
        channel_unit = (reference_dataset or {}).get('unit') or ""
        channel_label, channel_unit_clean = compose_axis_label(self._y_label or meta.get('channel') or "Value", channel_unit)
        blocks = ["\n".join(self._provenance_header_lines(reference_dataset))]
        columns = []
        max_len = 0
        for name, dataset in datasets:
            x = dataset.get('x_nm')
            unit = dataset.get('axis_unit') or dataset.get('distance_unit') or 'nm'
            if x is None:
                x = dataset.get('x_px')
                unit = 'px'
            vals = dataset.get('vals')
            if x is None or vals is None:
                continue
            x = list(x)
            vals = list(vals)
            max_len = max(max_len, len(x), len(vals))
            columns.append((name, unit, x, vals))
        if not columns:
            QtWidgets.QMessageBox.information(self, "Copy profile", "Profile data is incomplete.")
            return
        header_row = []
        for name, unit, _x, _vals in columns:
            header_row.append(f"{name} d ({unit})")
            header_row.append(f"{name} {channel_label} ({channel_unit_clean})".rstrip())
        rows = ["\t".join(header_row)]
        for i in range(max_len):
            row = []
            for _name, _unit, x, vals in columns:
                try:
                    row.append(f"{float(x[i]):.9g}")
                except Exception:
                    row.append("")
                try:
                    row.append(f"{float(vals[i]):.9g}")
                except Exception:
                    row.append("")
            rows.append("\t".join(row))
        blocks.append("\n".join(rows))
        QtWidgets.QApplication.clipboard().setText("\n\n".join(blocks))
        self.toast.show_message(f"Copied {max_len} XY {'pairs' if len(columns) == 1 else f'pairs x {len(columns)} profiles'} with a provenance header")

    # ------------------------------------------------------------ context menu
    def _on_context_menu(self, pos):
        """Right-click accelerator for the same menu the header button shows.

        Anchored to the plot's top-right corner rather than the click point:
        opening exactly at the cursor put the menu over the data it was
        styling, hiding the very marker being measured.
        """
        self._export_menu.exec_(self._context_menu_anchor_pos())

    def _context_menu_anchor_pos(self):
        corner = self.canvas.rect().topRight() + QtCore.QPoint(-4, 4)
        return self.canvas.mapToGlobal(corner)

    def _on_profile_list_context_menu(self, pos):
        self._export_menu.exec_(self.profile_table.viewport().mapToGlobal(pos))

    def _refresh_plot(self):
        self.update_profiles(
            self._active, self._saved,
            activate_overlay_callback=self._activate_overlay_cb,
            highlight_overlay_callback=self._highlight_overlay_cb,
        )

    def _populate_export_menu(self):
        """(Re)build self._export_menu's contents from current state.

        Called from ``aboutToShow`` before every display (header button
        dropdown or right-click accelerator), so checked-states and the
        "Style <target>" submenu's target always reflect what's selected
        right now - not whatever was selected the last time the menu
        happened to be built.
        """
        menu = self._export_menu
        menu.clear()

        preset_menu = menu.addMenu("Figure preset")
        for preset in _PROFILE_FIGURE_PRESETS:
            act = preset_menu.addAction(preset.label)
            act.setCheckable(True)
            act.setChecked(self._figure_preset_key == preset.key)
            act.triggered.connect(lambda _=False, key=preset.key: self._apply_figure_preset(key))

        add_font_menu_action(
            menu, self, self._plot_font_family, self.set_plot_font_family,
            current_style=self._font_style_state(), apply_style_callback=self.set_plot_typography,
        )

        menu.addSeparator()
        metadata_menu = menu.addMenu("Metadata")
        self._add_toggle_action(metadata_menu, "Show metadata on plot", self._metadata_visible,
                                 lambda v: self._set_flag_and_refresh("_metadata_visible", v))
        self._add_toggle_action(metadata_menu, "File name", self._metadata_show_filename,
                                 lambda v: self._set_flag_and_refresh("_metadata_show_filename", v))
        self._add_toggle_action(metadata_menu, "Acquisition title", self._metadata_show_acquisition,
                                 lambda v: self._set_flag_and_refresh("_metadata_show_acquisition", v))
        self._add_toggle_action(metadata_menu, "Acquisition time", self._metadata_show_time,
                                 lambda v: self._set_flag_and_refresh("_metadata_show_time", v))
        self._add_toggle_action(metadata_menu, "Folder name", self._metadata_show_folder_name,
                                 lambda v: self._set_flag_and_refresh("_metadata_show_folder_name", v))
        self._add_toggle_action(metadata_menu, "Folder", self._metadata_show_folder,
                                 lambda v: self._set_flag_and_refresh("_metadata_show_folder", v))

        menu.addSeparator()
        act = menu.addAction("Copy plot — PNG 300 dpi\tCtrl+C")
        act.triggered.connect(lambda: self._copy_plot("png", dpi=300))
        act = menu.addAction("Copy plot — PNG 600 dpi")
        act.triggered.connect(lambda: self._copy_plot("png", dpi=600))
        act = menu.addAction("Copy plot — SVG")
        act.triggered.connect(lambda: self._copy_plot("svg"))
        act = menu.addAction("Copy data — XY\tCtrl+Shift+C")
        act.triggered.connect(self._copy_current_profile)
        save_menu = menu.addMenu("Save plot")
        act = save_menu.addAction("PNG 300 dpi...\tCtrl+S")
        act.triggered.connect(lambda: self._save_plot("png", dpi=300))
        act = save_menu.addAction("PNG 600 dpi...")
        act.triggered.connect(lambda: self._save_plot("png", dpi=600))
        act = save_menu.addAction("SVG (vector)...")
        act.triggered.connect(lambda: self._save_plot("svg"))
        act = save_menu.addAction("PDF (vector)...")
        act.triggered.connect(lambda: self._save_plot("pdf"))
        act = menu.addAction("Save data as CSV...")
        act.triggered.connect(self._save_csv)
        self._add_toggle_action(menu, "Include measurement band in exports", self._include_band_in_export,
                                 self._set_include_band_in_export)

        menu.addSeparator()
        legend_menu = menu.addMenu("Legend")
        self._add_toggle_action(legend_menu, "Show legend", self._legend_visible,
                                 lambda v: self._set_flag_and_refresh("_legend_visible", v))
        size_menu = legend_menu.addMenu("Font size")
        for size in (7.0, 8.0, 9.0, 10.0, 12.0, 14.0):
            act = size_menu.addAction(f"{size:.0f} pt")
            act.setCheckable(True)
            act.setChecked(abs(float(self._legend_fontsize) - size) < 1e-6)
            act.triggered.connect(lambda _=False, s=size: self._set_legend_fontsize(s))

        target_idx = self._selected_overlay_index()
        target_active = target_idx is None
        style_target = "Active profile" if target_active else f"Overlay {int(target_idx) + 1}"
        style_menu = menu.addMenu(f"Style {style_target}")
        pick_color_act = style_menu.addAction("Pick color...")
        pick_color_act.triggered.connect(lambda: self._pick_style_color(target_idx))
        width_menu = style_menu.addMenu("Line thickness")
        current_lw = self._profile_value(target_idx, "lw", 1.5 if not target_active else 2.0)
        for width in (0.5, 1.0, 1.5, 2.0, 3.0, 4.0, 6.0):
            act = width_menu.addAction(f"{width:.1f} pt")
            act.setCheckable(True)
            act.setChecked(abs(float(current_lw) - width) < 1e-6)
            act.triggered.connect(lambda _=False, w=width: self._apply_profile_style_change(target_idx, lw=w))
        marker_size_menu = style_menu.addMenu("Marker size")
        current_marker_size = self._profile_value(target_idx, "marker_size", 7.0 if target_active else 5.0)
        for size in (3.0, 5.0, 7.0, 9.0, 12.0):
            act = marker_size_menu.addAction(f"{size:.0f} pt")
            act.setCheckable(True)
            act.setChecked(abs(float(current_marker_size) - size) < 1e-6)
            act.triggered.connect(lambda _=False, s=size: self._apply_profile_style_change(target_idx, marker_size=s))

        palette_menu = menu.addMenu("Apply Color Palette")
        for name in list_color_cycles():
            act = palette_menu.addAction(name)
            act.setCheckable(True)
            act.setChecked(name == self._profile_palette_name)
            act.triggered.connect(lambda _=False, n=name: self._apply_palette_change(n))

        menu.addSeparator()
        reset_act = menu.addAction("Reset style")
        reset_act.triggered.connect(self._reset_style)

    def _add_toggle_action(self, menu, text, checked, on_toggled):
        act = menu.addAction(text)
        act.setCheckable(True)
        act.setChecked(bool(checked))
        act.triggered.connect(lambda checked_state: on_toggled(bool(checked_state)))
        return act

    def _set_flag_and_refresh(self, attr, value):
        setattr(self, attr, bool(value))
        self._refresh_plot()

    def _set_include_band_in_export(self, value):
        self._include_band_in_export = bool(value)
        QtCore.QSettings().setValue("profileDialog/includeBandInExport", self._include_band_in_export)

    def _set_legend_fontsize(self, size):
        self._legend_fontsize = float(size)
        self._apply_legend_style()

    def _pick_style_color(self, target_idx):
        current = QtGui.QColor(self._profile_value(target_idx, "color", "#fbc02d"))
        picked = QtWidgets.QColorDialog.getColor(current, self, "Select color")
        if picked.isValid():
            self._apply_profile_style_change(target_idx, color=picked.name())

    def _reset_style(self):
        self._profile_palette_name = DEFAULT_COLOR_CYCLE
        self._legend_fontsize = 8.0
        self._legend_visible = True
        self._apply_palette_change(DEFAULT_COLOR_CYCLE)
        self._apply_figure_preset("interactive")
        self.toast.show_message("Style reset to defaults")

    def _apply_figure_preset(self, preset_key):
        preset = _get_profile_figure_preset(preset_key)
        self._figure_preset_key = preset.key
        self._font_scale = float(preset.font_scale)
        self._legend_fontsize = float(preset.legend_font_pt)
        if callable(self._label_scale_cb):
            try:
                self._label_scale_cb(self._font_scale)
            except Exception:
                pass
        # Drive the on-screen plot size from the same width/height the
        # preset will export at (at a comfortable screen scale), so what's
        # visible while composing the figure resembles what comes out -
        # not just matching font sizes while the plot itself stays
        # whatever size the window happens to be.
        apply_figure_layout(self.canvas.figure, preset)
        w_px, h_px = preset_pixel_size(self, preset, max_fraction=0.6)
        apply_canvas_widget_preset(self.canvas, preset, w_px, h_px)
        self._apply_font_scale()
        if preset.key != "interactive":
            try:
                total_w = max(720, int(w_px + 140))
                total_h = max(520, int(h_px + 320))
                self.resize(total_w, total_h)
                if hasattr(self, "_splitter") and self._splitter is not None:
                    self._splitter.setSizes([h_px + 80, 220])
            except Exception:
                pass
        self.toast.show_message(f"Preset: {preset.label} — exports {self._export_px(300)} px at 300 dpi")

    def _save_csv(self):
        active = self._active
        saved = self._saved
        if not active and not saved:
            QtWidgets.QMessageBox.information(self, "Save data", "No profile data available.")
            return
        path, _sel = QtWidgets.QFileDialog.getSaveFileName(self, "Save data", "profile.csv", "CSV Files (*.csv)")
        if not path:
            return
        dataset = active or saved[0]
        meta = dataset.get('meta') or {}
        channel_unit = dataset.get('unit') or ""
        channel_label, channel_unit_clean = compose_axis_label(self._y_label or meta.get('channel') or "Value", channel_unit)
        x = dataset.get('x_nm')
        unit = dataset.get('axis_unit') or dataset.get('distance_unit') or 'nm'
        if x is None:
            x = dataset.get('x_px')
            unit = 'px'
        vals = dataset.get('vals')
        try:
            with open(path, "w", encoding="utf-8", newline="") as fh:
                for line in self._provenance_header_lines(dataset):
                    fh.write(f"# {line}\n")
                fh.write(f"d_{unit},{channel_label}_{channel_unit_clean or 'value'}\n")
                x_list = list(x) if x is not None else []
                vals_list = list(vals) if vals is not None else []
                for xi, yi in zip(x_list, vals_list):
                    fh.write(f"{float(xi):.9g},{float(yi):.9g}\n")
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Save data", str(exc))
            return
        self.toast.show_message(f"Saved {Path(path).name} with a provenance header")

    # -------------------------------------------------------------- typography
    def set_plot_typography(self, **changes):
        """Update shared plot typography (family/bold/italic/underline)."""
        family = changes.get("family", None)
        owner = getattr(self, "_owner", None)
        style_changes = {
            "bold": changes.get("bold", None),
            "italic": changes.get("italic", None),
            "underline": changes.get("underline", None),
        }
        if family is not None:
            family = normalize_font_family(family, "sans-serif")
            self._plot_font_family = family
        if owner is not None and hasattr(owner, "set_plot_typography"):
            target = {
                "family": family if family is not None else self._plot_font_family,
                "bold": bool(style_changes["bold"] if style_changes["bold"] is not None else self._plot_font_bold),
                "italic": bool(style_changes["italic"] if style_changes["italic"] is not None else self._plot_font_italic),
                "underline": bool(style_changes["underline"] if style_changes["underline"] is not None else self._plot_font_underline),
            }
            if any(getattr(owner, f"_plot_font_{k}", None) != v for k, v in target.items()):
                try:
                    owner.set_plot_typography(**target)
                    return
                except Exception:
                    pass
        for key, attr in (("bold", "_plot_font_bold"), ("italic", "_plot_font_italic"), ("underline", "_plot_font_underline")):
            if style_changes[key] is not None:
                setattr(self, attr, bool(style_changes[key]))
        self._apply_font_scale()

    def set_plot_font_family(self, family: str):
        self.set_plot_typography(family=family)

    # --------------------------------------------------------- profile list
    def _dataset_point_count(self, dataset):
        x = dataset.get('x_nm')
        if x is None:
            x = dataset.get('x_px')
        try:
            return len(x) if x is not None else 0
        except Exception:
            return 0

    def _format_length_cell(self, dataset):
        """One clear (total length, pixel size) pair from the SAME array.

        The old dialog showed two independently-computed lengths side by
        side ("L=198 nm: 198.320 nm") - one baked into a label at
        extraction time, one recomputed here - which read as either a
        nominal/actual pair (it wasn't) or noise (it was). Deriving both
        numbers from one x array can't disagree with itself.
        """
        x = dataset.get('x_nm')
        unit = dataset.get('axis_unit') or dataset.get('distance_unit') or 'nm'
        if x is None:
            x = dataset.get('x_px')
            unit = 'px'
        if x is None or len(x) == 0:
            return "N/A"
        x_arr = np.asarray(x, dtype=float)
        total_ext = float(x_arr[-1] - x_arr[0])
        px_ext = float(x_arr[1] - x_arr[0]) if x_arr.size >= 2 else 0.0
        total_plot, si_unit, _auto = to_si_base(np.asarray([total_ext]), unit)
        px_plot, _u2, _a2 = to_si_base(np.asarray([px_ext]), unit)
        if si_unit:
            total_txt = si_format(float(total_plot[0]), unit=si_unit, precision=4)
            px_txt = si_format(float(px_plot[0]), unit=si_unit, precision=3)
        else:
            total_txt = f"{total_ext:.4g} {unit}"
            px_txt = f"{px_ext:.3g} {unit}"
        return f"{total_txt}  ({px_txt}/px)"

    def _populate_profile_list(self, active_profile, saved_profiles):
        table = self.profile_table
        table.blockSignals(True)
        entries = self._ordered_profile_entries(active_profile, saved_profiles)
        table.setRowCount(len(entries))
        target_row = None
        for row, entry in enumerate(entries):
            data = entry["data"] or {}
            key = entry["key"]
            label = entry["label"]

            vis_item = QtWidgets.QTableWidgetItem()
            vis_item.setFlags(QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            visible = self._overlay_visible_by_key.get(key, True)
            vis_item.setCheckState(QtCore.Qt.Checked if visible else QtCore.Qt.Unchecked)
            vis_item.setToolTip("Show / hide this profile's curve")
            vis_item.setData(QtCore.Qt.UserRole, key)
            table.setItem(row, 0, vis_item)

            color = data.get('color') or self._fallback_curve_color(entry["is_active"])
            chip = _ColorChip(color, on_double_click=lambda k=key: self._pick_row_color(k))
            table.setCellWidget(row, 1, chip)

            meta = data.get('meta') or {}
            file_name = data.get('source_file_name') or meta.get('file_name') or ''
            channel_name = meta.get('channel') or self._y_label or ''
            direction = meta.get('direction') or data.get('direction') or ''
            values = {
                2: file_name,
                3: channel_name,
                4: direction or '—',
                5: self._format_length_cell(data),
                6: str(self._dataset_point_count(data)),
            }
            for col, text in values.items():
                cell = QtWidgets.QTableWidgetItem(text)
                cell.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
                cell.setData(QtCore.Qt.UserRole, key)
                if col == 2:
                    cell.setToolTip(self._dataset_display_name(data, label))
                table.setItem(row, col, cell)

            if target_row is None:
                target_row = row
        table.blockSignals(False)
        self._resize_profile_table()
        if target_row is not None:
            table.selectRow(target_row)
            self._on_profile_row_selected()

    def _resize_profile_table(self):
        table = self.profile_table
        header_h = table.horizontalHeader().height()
        total = header_h + table.frameWidth() * 2 + 6
        for row in range(min(table.rowCount(), 4)):
            total += table.rowHeight(row)
        if table.rowCount() == 0:
            total += 24
        table.setFixedHeight(max(total, 70))

    def select_overlay(self, idx):
        table = self.profile_table
        table.blockSignals(True)
        target_row = None
        for row in range(table.rowCount()):
            item = table.item(row, 0)
            if item is not None and item.data(QtCore.Qt.UserRole) == idx:
                target_row = row
                break
        if target_row is None and idx is None and table.rowCount():
            target_row = 0
        if target_row is not None:
            table.selectRow(target_row)
            self._on_profile_row_selected()
        table.blockSignals(False)

    def _on_profile_visibility_changed(self, item):
        if item.column() != 0:
            return
        key = item.data(QtCore.Qt.UserRole)
        self._overlay_visible_by_key[key] = item.checkState() == QtCore.Qt.Checked
        self._refresh_plot()

    def _pick_row_color(self, key):
        self._pick_style_color(key)

    def _on_profile_row_selected(self):
        row = self.profile_table.currentRow()
        if row < 0:
            return
        id_item = self.profile_table.item(row, 0)
        idx = id_item.data(QtCore.Qt.UserRole) if id_item is not None else None
        self._current_marker_key = idx
        if callable(self._marker_key_cb):
            try:
                self._marker_key_cb(self._current_marker_key)
            except Exception:
                pass
        for profile_key, item in list(self._line_handles_by_key.items()):
            try:
                dataset = self._dataset_for_profile_key(profile_key) or {}
                base_lw = float(dataset.get('lw') or _DEFAULT_CURVE_WIDTH)
                item.set_linewidth(base_lw + (0.4 if profile_key == idx else 0.0))
            except Exception:
                pass
        self.canvas.draw_idle()
        if self._highlight_overlay_cb:
            try:
                self._highlight_overlay_cb(idx)
            except Exception:
                pass
        dataset = self._dataset_for_profile_key(idx)
        if dataset:
            ref_points = dataset.get('x_nm') if dataset.get('x_nm') is not None else dataset.get('x_px')
            ref_length = dataset.get('length_nm')
            self._reset_region(ref_points, ref_length, reference_dataset=dataset, store_state=False)
            if idx in self._marker_positions_by_key:
                positions = self._marker_positions_by_key.get(idx)
                domain = self._marker_domain_by_key.get(idx)
                if positions:
                    self.set_marker_positions(positions, domain=domain)
            else:
                self._marker_positions_by_key[idx] = list(self._marker_positions)
                self._marker_domain_by_key[idx] = tuple(self._marker_domain)

    def _add_overlay_from_active(self):
        if callable(self._add_overlay_cb):
            self._add_overlay_cb()

    def _selected_overlay_index(self):
        row = self.profile_table.currentRow()
        if row >= 0:
            item = self.profile_table.item(row, 0)
            if item is not None:
                idx = item.data(QtCore.Qt.UserRole)
                if idx is not None:
                    try:
                        return int(idx)
                    except Exception:
                        pass
        current_key = self._current_marker_key
        if current_key is not None:
            try:
                return int(current_key)
            except Exception:
                pass
        return None

    def _refresh_action_button_states(self):
        try:
            self.add_btn.setEnabled(callable(self._add_overlay_cb))
        except Exception:
            pass
        try:
            self.delete_btn.setEnabled(bool(self._active or self._saved))
        except Exception:
            pass
        try:
            self.compose_drag_btn.setEnabled(bool(self._active or self._saved))
        except Exception:
            pass

    def _delete_selected_profile(self):
        row = self.profile_table.currentRow()
        if row < 0:
            QtWidgets.QMessageBox.information(self, "Delete profile", "Select a profile to delete.")
            return
        id_item = self.profile_table.item(row, 0)
        idx = id_item.data(QtCore.Qt.UserRole) if id_item is not None else None
        target_dataset = self._dataset_for_profile_key(idx) or {}
        # Name the actual source file, not a generic display label (which
        # can be just a channel/style name) - "Delete <this file>?" is what
        # lets someone confirm they're removing the right thing.
        target_name = (
            target_dataset.get('source_file_name')
            or (target_dataset.get('meta') or {}).get('file_name')
            or self._dataset_display_name(target_dataset, "this profile")
        )
        confirmed = QtWidgets.QMessageBox.question(
            self, "Delete profile", f'Delete "{target_name}"?',
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No, QtWidgets.QMessageBox.No,
        )
        if confirmed != QtWidgets.QMessageBox.Yes:
            return
        if callable(self._delete_overlay_cb) and idx is not None:
            removed = bool(self._delete_overlay_cb(idx))
            if removed and 0 <= idx < len(self._saved):
                saved = list(self._saved)
                saved.pop(idx)
                self.update_profiles(
                    self._active, saved,
                    activate_overlay_callback=self._activate_overlay_cb,
                    highlight_overlay_callback=self._highlight_overlay_cb,
                )
            return
        if callable(self._delete_overlay_cb) and idx is None:
            QtWidgets.QMessageBox.information(self, "Delete profile", "The active live profile cannot be deleted here.")
            return
        active = copy.deepcopy(self._active) if self._active is not None else None
        saved = copy.deepcopy(list(self._saved or []))
        if idx is None:
            active = saved.pop(0) if saved else None
        else:
            try:
                saved.pop(int(idx))
            except Exception:
                QtWidgets.QMessageBox.information(self, "Delete profile", "Select a valid profile to delete.")
                return
        self.update_profiles(
            active, saved,
            activate_overlay_callback=self._activate_overlay_cb,
            highlight_overlay_callback=self._highlight_overlay_cb,
        )

    # --------------------------------------------------------- composite drag
    @staticmethod
    def _json_ready_profile_value(value):
        if isinstance(value, np.ndarray):
            return {"__profile_array__": value.tolist()}
        if isinstance(value, np.generic):
            return value.item()
        if isinstance(value, dict):
            return {str(k): ProfileDialog._json_ready_profile_value(v) for k, v in value.items()}
        if isinstance(value, (list, tuple)):
            return [ProfileDialog._json_ready_profile_value(v) for v in value]
        return value

    @staticmethod
    def _profile_value_from_json(value):
        if isinstance(value, dict):
            if "__profile_array__" in value:
                try:
                    return np.asarray(value.get("__profile_array__"))
                except Exception:
                    return np.asarray([])
            return {k: ProfileDialog._profile_value_from_json(v) for k, v in value.items()}
        if isinstance(value, list):
            return [ProfileDialog._profile_value_from_json(v) for v in value]
        return value

    @staticmethod
    def _profile_dataset_signature(dataset):
        if not isinstance(dataset, dict):
            return ""
        digest = hashlib.sha1()
        meta = dataset.get("meta") or {}
        for key in ("source_path", "source_file_name", "source_title", "source_acquisition_text",
                    "label", "unit", "axis_unit", "distance_unit"):
            digest.update(str(dataset.get(key) or "").encode("utf-8", errors="ignore"))
            digest.update(b"\0")
        for key in ("channel", "file_name", "datetime", "date", "time"):
            digest.update(str(meta.get(key) or "").encode("utf-8", errors="ignore"))
            digest.update(b"\0")
        for key in ("x_px", "x_nm", "vals"):
            try:
                arr = np.asarray(dataset.get(key) if dataset.get(key) is not None else [], dtype=float)
            except Exception:
                arr = np.asarray([], dtype=float)
            digest.update(str(arr.shape).encode("utf-8", errors="ignore"))
            digest.update(arr.tobytes())
        return digest.hexdigest()

    def _dataset_display_name(self, dataset, fallback_label):
        if not isinstance(dataset, dict):
            return str(fallback_label or "Profile").strip()
        existing = str(dataset.get("display_name") or "").strip()
        if existing:
            return existing
        meta = dict(dataset.get("meta") or {})
        source_name = str(
            dataset.get("source_file_name") or meta.get("file_name") or dataset.get("source_title") or ""
        ).strip()
        channel = str(meta.get("channel") or "").strip()
        profile_label = str(dataset.get("label") or fallback_label or "").strip()
        parts = [p for p in (source_name, channel, profile_label) if p]
        if not parts:
            parts.append(str(fallback_label or "Profile").strip() or "Profile")
        return " | ".join(parts[:3])

    def _clone_dataset_for_composite(self, dataset, fallback_label):
        if not isinstance(dataset, dict):
            return None
        cloned = copy.deepcopy(dataset)
        cloned["display_name"] = self._dataset_display_name(cloned, fallback_label)
        return cloned

    def _current_profile_entries(self):
        entries = []
        if self._active:
            dataset = self._clone_dataset_for_composite(self._active, "Active")
            if dataset:
                entries.append({"dataset": dataset, "signature": self._profile_dataset_signature(dataset)})
        for idx, data in enumerate(self._saved, 1):
            dataset = self._clone_dataset_for_composite(data, f"Overlay {idx}")
            if dataset:
                entries.append({"dataset": dataset, "signature": self._profile_dataset_signature(dataset)})
        return entries

    def _composite_payload(self):
        entries = self._current_profile_entries()
        if not entries:
            return None
        return {
            "origin_dialog_id": self._composite_origin_id,
            "dialog_title": str(self.windowTitle() or "Profile measurement"),
            "entries": [
                {"signature": e.get("signature") or "", "dataset": self._json_ready_profile_value(e.get("dataset") or {})}
                for e in entries
            ],
        }

    def start_profile_composite_drag(self):
        payload = self._composite_payload()
        if not payload:
            QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), "No profile data to drag", self)
            return
        drag = QtGui.QDrag(self)
        mime = QtCore.QMimeData()
        mime.setData(_PROFILE_COMPOSITE_MIME, json.dumps(payload).encode("utf-8"))
        drag.setMimeData(mime)
        pixmap = QtGui.QPixmap(120, 28)
        pixmap.fill(QtGui.QColor("#1d3557"))
        painter = QtGui.QPainter(pixmap)
        painter.setPen(QtGui.QColor("#f1faee"))
        painter.drawText(pixmap.rect().adjusted(8, 0, -8, 0), QtCore.Qt.AlignVCenter | QtCore.Qt.AlignLeft, "Composite profile")
        painter.end()
        drag.setPixmap(pixmap)
        drag.setHotSpot(QtCore.QPoint(16, 14))
        drag.exec_(QtCore.Qt.CopyAction)

    def _entries_from_profile_payload(self, payload):
        entries = []
        for raw in list((payload or {}).get("entries") or []):
            if not isinstance(raw, dict):
                continue
            dataset = self._profile_value_from_json(raw.get("dataset"))
            if not isinstance(dataset, dict):
                continue
            signature = str(raw.get("signature") or self._profile_dataset_signature(dataset))
            dataset["display_name"] = self._dataset_display_name(dataset, dataset.get("display_name") or "Profile")
            entries.append({"dataset": dataset, "signature": signature})
        return entries

    def _merge_profile_entries(self, incoming_entries):
        merged = []
        seen = set()
        for group in (self._current_profile_entries(), incoming_entries):
            for entry in group:
                signature = str(entry.get("signature") or "")
                dataset = entry.get("dataset")
                if not isinstance(dataset, dict):
                    continue
                if signature and signature in seen:
                    continue
                if signature:
                    seen.add(signature)
                merged.append({"dataset": copy.deepcopy(dataset), "signature": signature})
        return merged

    def _register_workspace_dialog(self):
        if self._workspace_registered:
            return
        owner = getattr(self, "_owner", None)
        dialogs = getattr(owner, "_profile_dialogs", None)
        if dialogs is not None and self not in dialogs:
            dialogs.append(self)
        refs = getattr(owner, "_popup_refs", None)
        if refs is not None and self not in refs:
            refs.append(self)
        controller = getattr(owner, "quick_crop_controller", None)
        if controller:
            try:
                controller.update_popup_actions()
            except Exception:
                pass
        self._workspace_registered = True

    def _deregister_workspace_dialog(self):
        if not self._workspace_registered:
            return
        owner = getattr(self, "_owner", None)
        dialogs = getattr(owner, "_profile_dialogs", None)
        if dialogs is not None and self in dialogs:
            dialogs.remove(self)
        refs = getattr(owner, "_popup_refs", None)
        if refs is not None and self in refs:
            refs.remove(self)
        controller = getattr(owner, "quick_crop_controller", None)
        if controller:
            try:
                controller.update_popup_actions()
            except Exception:
                pass
        self._workspace_registered = False

    def _spawn_composite_dialog(self, merged_entries):
        datasets = [copy.deepcopy(e.get("dataset") or {}) for e in merged_entries if isinstance(e.get("dataset"), dict)]
        if not datasets:
            return None
        active = datasets[0]
        saved = datasets[1:]
        owner = getattr(self, "_owner", None)
        unit = active.get("unit") or self._unit
        meta = dict(active.get("meta") or {})
        y_label = str(meta.get("channel") or self._y_label or "Profile value").strip()
        dlg = ProfileDialog(active, saved, parent=owner, unit=unit, y_label=y_label, dark_mode=bool(self._dark_background))
        dlg._composite_mode = True
        dlg.setWindowTitle(f"Profile composite ({len(datasets)})")
        if hasattr(dlg, "detach_as_workspace_window"):
            dlg.detach_as_workspace_window()
        for attr in ("show_lines_cb", "show_points_cb", "grid_cb", "extra_ticks_cb",
                     "precision_cb", "multi_channel_cb", "measure_toggle", "snap_toggle", "crosshair_toggle"):
            src = getattr(self, attr, None)
            dst = getattr(dlg, attr, None)
            if src is not None and dst is not None:
                try:
                    dst.setChecked(bool(src.isChecked()))
                except Exception:
                    pass
        try:
            dlg._theme_mode = self._theme_mode
            dlg._dark_background = self._dark_background
            dlg._apply_plot_theme()
        except Exception:
            pass
        dlg.update_profiles(active, saved)
        try:
            base_geo = self.frameGeometry()
            dlg.move(base_geo.topLeft() + QtCore.QPoint(36, 36))
        except Exception:
            pass
        dlg.setAttribute(QtCore.Qt.WA_DeleteOnClose, True)
        dlg._register_workspace_dialog()
        dlg.finished.connect(lambda _=None, ref=dlg: ref._deregister_workspace_dialog())
        dlg.show()
        try:
            dlg.raise_()
            dlg.activateWindow()
        except Exception:
            pass
        return dlg

    def _create_composite_from_drop(self, payload):
        if not isinstance(payload, dict):
            return None
        if str(payload.get("origin_dialog_id") or "") == self._composite_origin_id:
            return None
        incoming_entries = self._entries_from_profile_payload(payload)
        if not incoming_entries:
            return None
        merged = self._merge_profile_entries(incoming_entries)
        if len(merged) <= 1:
            return None
        return self._spawn_composite_dialog(merged)

    def _set_compose_drop_active(self, active):
        btn = getattr(self, "compose_drag_btn", None)
        if btn is None:
            return
        try:
            btn.setProperty("dropActive", bool(active))
        except Exception:
            return
        self._apply_toggle_button_styles()

    def _show_compose_help(self, global_pos=None):
        message = (
            "Drag Combine onto another profile window.\n"
            "You can also drag from the plot margin to create a composite."
        )
        try:
            QtWidgets.QToolTip.showText(global_pos or QtGui.QCursor.pos(), message, self)
        except Exception:
            pass

    def dragEnterEvent(self, event):
        if event.mimeData().hasFormat(_PROFILE_COMPOSITE_MIME):
            self._set_compose_drop_active(True)
            event.acceptProposedAction()
            return
        self._set_compose_drop_active(False)
        event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasFormat(_PROFILE_COMPOSITE_MIME):
            self._set_compose_drop_active(True)
            event.acceptProposedAction()
            return
        self._set_compose_drop_active(False)
        event.ignore()

    def dragLeaveEvent(self, event):
        self._set_compose_drop_active(False)
        super().dragLeaveEvent(event)

    def dropEvent(self, event):
        if not event.mimeData().hasFormat(_PROFILE_COMPOSITE_MIME):
            self._set_compose_drop_active(False)
            event.ignore()
            return
        data = event.mimeData().data(_PROFILE_COMPOSITE_MIME)
        try:
            payload = json.loads(bytes(data).decode("utf-8"))
        except Exception:
            self._set_compose_drop_active(False)
            event.ignore()
            return
        dlg = self._create_composite_from_drop(payload)
        self._set_compose_drop_active(False)
        if dlg is None:
            QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), "No new composite created", self)
            event.ignore()
            return
        QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), "Composite profile created", dlg)
        event.acceptProposedAction()

    def refresh_linked_profiles(self, datasets_by_ref, source_id=None):
        if not self._composite_mode or not isinstance(datasets_by_ref, dict):
            return False
        changed = False

        def _refresh_dataset(dataset):
            if not isinstance(dataset, dict):
                return dataset, False
            key = profile_ref_key(dataset.get("live_profile_ref"))
            if key is None:
                return dataset, False
            if source_id and str(key[0]) != str(source_id):
                return dataset, False
            updated = datasets_by_ref.get(key)
            if not isinstance(updated, dict):
                return dataset, False
            refreshed = copy.deepcopy(updated)
            refreshed["display_name"] = self._dataset_display_name(
                refreshed, dataset.get("display_name") or refreshed.get("display_name") or "Profile",
            )
            return refreshed, True

        current_key = self._selected_overlay_index()
        self._active, active_changed = _refresh_dataset(self._active)
        changed = changed or active_changed
        for idx, dataset in enumerate(list(self._saved)):
            refreshed, item_changed = _refresh_dataset(dataset)
            if item_changed:
                self._saved[idx] = refreshed
                changed = True
        if not changed:
            return False
        self.update_profiles(
            self._active, self._saved,
            activate_overlay_callback=self._activate_overlay_cb,
            highlight_overlay_callback=self._highlight_overlay_cb,
        )
        self.select_overlay(current_key)
        return True

    # --------------------------------------------------------------- close
    def closeEvent(self, event):
        if self._highlight_overlay_cb:
            try:
                self._highlight_overlay_cb(None)
            except Exception:
                pass
        self._deregister_workspace_dialog()
        unregister_profile_dialog(self)
        super().closeEvent(event)
