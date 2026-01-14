"""Layout builders for the main SXM Viewer window."""
from __future__ import annotations

from PyQt5 import QtCore, QtGui, QtWidgets

from .constants import UI_FONT_FAMILY
from .styles import MAIN_SHORTCUTS_PANEL_STYLE, lower_control_frame_style, mode_selector_style


def create_lower_controls(viewer):
    frame = QtWidgets.QFrame()
    frame.setObjectName("lowerControlFrame")
    frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
    layout = QtWidgets.QHBoxLayout(frame)
    layout.setContentsMargins(8, 4, 8, 4)
    layout.setSpacing(12)

    mode_widget = QtWidgets.QWidget(frame)
    mode_widget.setObjectName("modeSelector")
    viewer.mode_selector_widget = mode_widget
    mode_widget.setSizePolicy(QtWidgets.QSizePolicy.Maximum, QtWidgets.QSizePolicy.Preferred)
    mode_layout = QtWidgets.QHBoxLayout(mode_widget)
    mode_layout.setContentsMargins(0, 0, 0, 0)
    mode_layout.setSpacing(0)
    viewer.mode_button_group = QtWidgets.QButtonGroup(mode_widget)
    viewer.mode_button_group.setExclusive(True)
    viewer.mode_buttons = {}
    mode_definitions = [
        (viewer.MODE_BROWSE, "Browse", "Ctrl+B"),
        (viewer.MODE_MEASURE, "Measure", "Ctrl+M"),
        (viewer.MODE_SPECTRO, "Spectroscopy", "Ctrl+S"),
    ]
    for mode, label, shortcut in mode_definitions:
        btn = QtWidgets.QToolButton(mode_widget)
        btn.setText(label)
        btn.setCheckable(True)
        btn.setAutoRaise(True)
        btn.setFocusPolicy(QtCore.Qt.StrongFocus)
        btn.setToolButtonStyle(QtCore.Qt.ToolButtonTextOnly)
        btn.setToolTip(f"{label} mode ({shortcut})")
        btn.clicked.connect(lambda checked, m=mode: viewer._on_mode_button_clicked(m))
        viewer.mode_button_group.addButton(btn, mode)
        viewer.mode_buttons[mode] = btn
        mode_layout.addWidget(btn)
    layout.addWidget(mode_widget)

    viewer.mode_stack = QtWidgets.QStackedWidget(frame)
    viewer.mode_stack.addWidget(build_browse_context_page(viewer))
    viewer.mode_stack.addWidget(build_measure_context_page(viewer))
    viewer.mode_stack.addWidget(build_spectro_context_page(viewer))
    layout.addWidget(viewer.mode_stack, 1)

    display_widget = build_display_widget(viewer, frame)
    layout.addWidget(display_widget)

    layout.setStretch(0, 0)
    layout.setStretch(1, 1)
    layout.setStretch(2, 0)

    settings = QtCore.QSettings()
    saved_mode = str(settings.value("lowerPane/lastMode", "Browse"))
    mode = viewer._mode_from_name(saved_mode)
    # Always start in Browse so no measurement overlays appear by default.
    if mode == viewer.MODE_MEASURE:
        mode = viewer.MODE_BROWSE
    viewer._apply_mode(mode, remember=False)
    viewer._apply_lower_control_theme()
    return frame


def build_browse_context_page(viewer):
    page = QtWidgets.QWidget()
    layout = QtWidgets.QHBoxLayout(page)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.setSpacing(6)
    viewer.add_view_btn = QtWidgets.QPushButton("+ Channel")
    viewer.add_view_btn.setToolTip("Add the current channel as an extra preview")
    viewer.clear_views_btn = QtWidgets.QPushButton("Clear views")
    viewer.clear_views_btn.setToolTip("Remove extra previews and keep only the main view")
    for btn in (
        viewer.add_view_btn,
        viewer.clear_views_btn,
    ):
        layout.addWidget(btn)
    layout.addStretch(1)
    return page


def build_measure_context_page(viewer):
    page = QtWidgets.QWidget()
    layout = QtWidgets.QHBoxLayout(page)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.setSpacing(6)
    viewer.measure_profile_btn = QtWidgets.QPushButton("Profile")
    viewer.measure_profile_btn.setToolTip("Start or stop interactive profile measurement")
    viewer.measure_angle_btn = QtWidgets.QPushButton("Angle")
    viewer.measure_angle_btn.setToolTip("Start or stop angle measurement tool")
    viewer.exit_profile_btn = QtWidgets.QPushButton("Exit")
    viewer.exit_profile_btn.setToolTip("Exit the profile measurement mode")
    viewer.clear_profile_btn = QtWidgets.QPushButton("Clear")
    viewer.clear_profile_btn.setToolTip("Clear the current profile line and start fresh")
    viewer.show_profile_window_btn = QtWidgets.QPushButton("Show")
    viewer.show_profile_window_btn.setToolTip("Reopen the profile dialog with current measurements")
    layout.addWidget(viewer.measure_profile_btn)
    layout.addWidget(viewer.measure_angle_btn)
    layout.addWidget(viewer.exit_profile_btn)
    layout.addWidget(viewer.clear_profile_btn)
    layout.addWidget(viewer.show_profile_window_btn)
    layout.addStretch(1)
    return page


def build_spectro_context_page(viewer):
    page = QtWidgets.QWidget()
    layout = QtWidgets.QHBoxLayout(page)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.setSpacing(6)
    viewer.show_spectra_cb = QtWidgets.QCheckBox("Show spectroscopies")
    viewer.show_spectra_cb.setChecked(viewer.show_spectra)
    viewer.show_spectra_cb.setToolTip("Toggle spectroscopy overlays in thumbnails")
    viewer.show_matrix_spectra_btn = QtWidgets.QPushButton("Show Matrix spectros")
    viewer.show_matrix_spectra_btn.setToolTip("Open a matrix spectroscopy viewer for the folder")
    viewer.clear_spec_selection_btn = QtWidgets.QPushButton("Clear spec selection")
    viewer.clear_spec_selection_btn.setToolTip("Clear the multi-selection of spectroscopy points")
    viewer.export_selected_btn = QtWidgets.QPushButton("Export selected (same view)")
    viewer.export_selected_btn.setToolTip("Export selected thumbnails using the same view layout")
    viewer.spec_selection_label = QtWidgets.QLabel("Spectra selected: 0")
    font_small = QtGui.QFont(UI_FONT_FAMILY, 9)
    viewer.spec_selection_label.setFont(font_small)
    layout.addWidget(viewer.show_spectra_cb)
    layout.addWidget(viewer.show_matrix_spectra_btn)
    layout.addWidget(viewer.clear_spec_selection_btn)
    layout.addWidget(viewer.export_selected_btn)
    layout.addWidget(viewer.spec_selection_label)
    layout.addStretch(1)
    return page


def _ensure_display_menu(viewer):
    if getattr(viewer, "display_menu", None):
        return viewer.display_menu
    viewer.display_menu = QtWidgets.QMenu(viewer)
    viewer.matrix_markers_act = viewer.display_menu.addAction("Matrix markers")
    viewer.matrix_markers_act.setCheckable(True)
    viewer.matrix_markers_act.setChecked(viewer.show_matrix_markers)
    viewer.matrix_markers_act.setToolTip("Toggle matrix spectroscopy markers")
    viewer.matrix_markers_act.toggled.connect(viewer.on_show_matrix_markers_toggled)
    viewer.single_markers_act = viewer.display_menu.addAction("Single markers")
    viewer.single_markers_act.setCheckable(True)
    viewer.single_markers_act.setChecked(viewer.show_single_markers)
    viewer.single_markers_act.setToolTip("Toggle single spectroscopy markers")
    viewer.single_markers_act.toggled.connect(viewer.on_show_single_markers_toggled)
    viewer.compact_markers_act = viewer.display_menu.addAction("Compact markers")
    viewer.compact_markers_act.setCheckable(True)
    viewer.compact_markers_act.setChecked(viewer.compact_markers)
    viewer.compact_markers_act.setToolTip("Use compact marker rendering")
    viewer.compact_markers_act.toggled.connect(viewer.on_compact_markers_toggled)
    viewer.density_markers_act = viewer.display_menu.addAction("Density overlay")
    viewer.density_markers_act.setCheckable(True)
    viewer.density_markers_act.setChecked(viewer.use_density_markers)
    viewer.density_markers_act.setToolTip("Show density overlay for spectroscopy clusters")
    viewer.density_markers_act.toggled.connect(viewer.on_density_markers_toggled)
    viewer.display_menu.addSeparator()
    
    markers_menu = viewer.display_menu.addMenu("Marker Style")
    col_single = markers_menu.addAction("Single marker color...")
    col_single.triggered.connect(viewer.on_pick_spectro_single_color)
    col_matrix = markers_menu.addAction("Matrix marker color...")
    col_matrix.triggered.connect(viewer.on_pick_spectro_matrix_color)
    markers_menu.addSeparator()
    sym_grp = QtWidgets.QActionGroup(viewer)
    for sym in ['circle', 'square', 'triangle', 'diamond']:
        act = markers_menu.addAction(sym.capitalize())
        act.setCheckable(True)
        act.setChecked(getattr(viewer, 'spectro_marker_symbol', 'circle') == sym)
        act.triggered.connect(lambda checked, s=sym: viewer.on_set_spectro_symbol(s))
        sym_grp.addAction(act)
    
    markers_menu.addSeparator()
    size_menu = markers_menu.addMenu("Marker Size")
    size_grp = QtWidgets.QActionGroup(viewer)
    current_size = getattr(viewer, 'spectro_marker_size', 5.0)
    for label, val in [("Tiny", 2.0), ("Small", 3.5), ("Medium", 5.0), ("Large", 7.0), ("Huge", 10.0)]:
        act = size_menu.addAction(label)
        act.setCheckable(True)
        act.setChecked(abs(current_size - val) < 0.1)
        act.triggered.connect(lambda checked, v=val: viewer.on_set_spectro_size(v))
        size_grp.addAction(act)

    viewer.display_menu.addSeparator()
    viewer.detail_dark_act = viewer.display_menu.addAction("Detail dark background")
    viewer.detail_dark_act.setCheckable(True)
    viewer.detail_dark_act.setChecked(viewer.detail_dark_view)
    viewer.detail_dark_act.setToolTip("Toggle dark background for the detailed preview view")
    viewer.detail_dark_act.toggled.connect(viewer.on_detail_dark_toggled)
    viewer.detail_grid_act = viewer.display_menu.addAction("Detail grid")
    viewer.detail_grid_act.setCheckable(True)
    viewer.detail_grid_act.setChecked(viewer.detail_grid_view)
    viewer.detail_grid_act.setToolTip("Toggle grid overlay on the detailed preview")
    viewer.detail_grid_act.toggled.connect(viewer.on_detail_grid_toggled)
    viewer.display_menu.addSeparator()
    reset_act = viewer.display_menu.addAction("Reset view")
    reset_act.setToolTip("Reset all display toggles to defaults")
    reset_act.triggered.connect(viewer._reset_display_options)
    return viewer.display_menu


def build_display_widget(viewer, parent):
    container = QtWidgets.QWidget(parent)
    layout = QtWidgets.QHBoxLayout(container)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.setSpacing(6)
    _ensure_display_menu(viewer)

    viewer.spectro_browser_btn = QtWidgets.QPushButton("Spectro Browser", container)
    viewer.spectro_browser_btn.setToolTip("Open the spectroscopy browser")
    viewer.spectro_browser_btn.clicked.connect(lambda: viewer.open_spectro_browser())
    layout.addWidget(viewer.spectro_browser_btn)

    layout.addStretch(1)
    viewer.spectro_stats_label = QtWidgets.QLabel(
        "Spectros: -- (Single: --, Matrix datasets: --)", container
    )
    stats_font = QtGui.QFont(UI_FONT_FAMILY, 9)
    viewer.spectro_stats_label.setFont(stats_font)
    viewer.spectro_stats_label.setToolTip("Summary of spectroscopy content for the loaded folder")
    layout.addWidget(viewer.spectro_stats_label, 0, QtCore.Qt.AlignRight)
    return container


def apply_lower_control_theme(viewer):
    frame = getattr(viewer, "lower_control_frame", None)
    mode_widget = getattr(viewer, "mode_selector_widget", None)
    if frame is None:
        return
    dark = bool(getattr(viewer, "dark_mode", False))
    if dark:
        border = "#4c4c4c"
        bg = "#2d2d2d"
        mode_border = "#5a5a5a"
        mode_text = "#f0f0f0"
        mode_checked = "#2b6cb0"
    else:
        border = "#c8c8c8"
        bg = "#f5f5f5"
        mode_border = "#b7b7b7"
        mode_text = "#202020"
        mode_checked = "#3d7dd8"
    frame.setStyleSheet(lower_control_frame_style(border, bg))
    if mode_widget is not None:
        mode_widget.setStyleSheet(mode_selector_style(mode_border, mode_text, mode_checked))


def create_shortcuts_panel(viewer):
    frame = QtWidgets.QFrame()
    frame.setObjectName("shortcutsPanel")
    frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
    frame.setStyleSheet(MAIN_SHORTCUTS_PANEL_STYLE)
    layout = QtWidgets.QVBoxLayout(frame)
    layout.setContentsMargins(10, 8, 8, 8)
    header = QtWidgets.QHBoxLayout()
    title = QtWidgets.QLabel("Shortcuts & gestures")
    title.setFont(QtGui.QFont(UI_FONT_FAMILY, 10, QtGui.QFont.Bold))
    header.addWidget(title)
    header.addStretch(1)
    never_btn = QtWidgets.QPushButton("Don't show again")
    never_btn.setFlat(True)
    never_btn.setCursor(QtCore.Qt.PointingHandCursor)
    never_btn.clicked.connect(viewer._on_shortcuts_never_show_clicked)
    header.addWidget(never_btn)
    close_btn = QtWidgets.QToolButton()
    close_btn.setText("?")
    close_btn.setAutoRaise(True)
    close_btn.setCursor(QtCore.Qt.PointingHandCursor)
    close_btn.clicked.connect(viewer._on_hide_shortcuts_panel)
    header.addWidget(close_btn)
    layout.addLayout(header)
    viewer.shortcuts_label = QtWidgets.QLabel(viewer._shortcuts_html())
    viewer.shortcuts_label.setWordWrap(True)
    viewer.shortcuts_label.setTextFormat(QtCore.Qt.RichText)
    layout.addWidget(viewer.shortcuts_label)
    return frame
