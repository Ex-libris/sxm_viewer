"""Layout builders for the main SXM Viewer window."""
from __future__ import annotations

from PyQt5 import QtCore, QtGui, QtWidgets

from .constants import UI_FONT_FAMILY
from .styles import MAIN_SHORTCUTS_PANEL_STYLE, lower_control_frame_style, mode_selector_style


def _configure_compact_control(widget):
    try:
        widget.setSizePolicy(QtWidgets.QSizePolicy.Maximum, QtWidgets.QSizePolicy.Fixed)
    except Exception:
        pass
    return widget


def _add_menu_widget(menu, widget):
    action = QtWidgets.QWidgetAction(menu)
    action.setDefaultWidget(widget)
    menu.addAction(action)
    return action


def create_lower_controls(viewer):
    frame = QtWidgets.QFrame()
    frame.setObjectName("lowerControlFrame")
    frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
    layout = QtWidgets.QVBoxLayout(frame)
    layout.setContentsMargins(8, 4, 8, 6)
    layout.setSpacing(6)

    top_row = QtWidgets.QHBoxLayout()
    top_row.setContentsMargins(0, 0, 0, 0)
    top_row.setSpacing(8)

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
        (viewer.MODE_SPECTRO, "Spectro", "Ctrl+S"),
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
    top_row.addWidget(mode_widget)
    top_row.addStretch(1)
    layout.addLayout(top_row)

    viewer.mode_stack = QtWidgets.QStackedWidget(frame)
    viewer.mode_stack.addWidget(build_browse_context_page(viewer))
    viewer.mode_stack.addWidget(build_measure_context_page(viewer))
    viewer.mode_stack.addWidget(build_spectro_context_page(viewer))
    layout.addWidget(viewer.mode_stack)

    display_widget = build_display_widget(viewer, frame)
    layout.addWidget(display_widget)

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
    viewer.add_view_btn = _configure_compact_control(QtWidgets.QPushButton("+ View"))
    viewer.add_view_btn.setToolTip("Add the current channel as an extra preview")
    viewer.clear_views_btn = _configure_compact_control(QtWidgets.QPushButton("Clear views"))
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
    viewer.measure_profile_btn = _configure_compact_control(QtWidgets.QPushButton("Profile"))
    viewer.measure_profile_btn.setToolTip("Start or stop interactive profile measurement")
    viewer.measure_angle_btn = _configure_compact_control(QtWidgets.QPushButton("Angle"))
    viewer.measure_angle_btn.setToolTip("Start or stop angle measurement tool")
    viewer.clear_profile_btn = _configure_compact_control(QtWidgets.QPushButton("Clear"))
    viewer.clear_profile_btn.setToolTip("Clear the current profile line and start fresh")
    viewer.show_profile_window_btn = _configure_compact_control(QtWidgets.QPushButton("Profiles"))
    viewer.show_profile_window_btn.setToolTip("Reopen the profile dialog with current measurements")
    viewer.exit_profile_btn = _configure_compact_control(QtWidgets.QPushButton("Done"))
    viewer.exit_profile_btn.setToolTip("Exit the profile measurement mode")
    layout.addWidget(viewer.measure_profile_btn)
    layout.addWidget(viewer.measure_angle_btn)
    layout.addWidget(viewer.clear_profile_btn)
    layout.addWidget(viewer.show_profile_window_btn)
    layout.addWidget(viewer.exit_profile_btn)
    layout.addStretch(1)
    return page


def build_spectro_context_page(viewer):
    page = QtWidgets.QWidget()
    layout = QtWidgets.QHBoxLayout(page)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.setSpacing(6)
    viewer.show_spectra_cb = _configure_compact_control(QtWidgets.QCheckBox("Preview markers"))
    viewer.show_spectra_cb.setChecked(getattr(viewer, "show_preview_spectra", True))
    viewer.show_spectra_cb.setToolTip("Toggle spectroscopy overlays in the preview panel")
    viewer.clear_spec_selection_btn = _configure_compact_control(QtWidgets.QPushButton("Clear selection"))
    viewer.clear_spec_selection_btn.setToolTip("Clear the multi-selection of spectroscopy points")
    viewer.grid_as_matrix_cb = QtWidgets.QCheckBox("NxN singles as matrix")
    viewer.grid_as_matrix_cb.setChecked(getattr(viewer, "spectro_single_grid_as_matrix", False))
    viewer.grid_as_matrix_cb.setToolTip("Interpret square grids of single .dat spectra as matrix datasets")
    viewer.force_single_cb = QtWidgets.QCheckBox("Force single mode")
    viewer.force_single_cb.setChecked(getattr(viewer, "spectro_force_single_mode", False))
    viewer.force_single_cb.setToolTip("Ignore matrix hints and treat all .dat as single spectra")
    viewer.spectro_more_btn = QtWidgets.QToolButton(page)
    viewer.spectro_more_btn.setText("More")
    viewer.spectro_more_btn.setToolButtonStyle(QtCore.Qt.ToolButtonTextOnly)
    viewer.spectro_more_btn.setPopupMode(QtWidgets.QToolButton.InstantPopup)
    viewer.spectro_more_btn.setToolTip("Show less-frequent spectroscopy options")
    _configure_compact_control(viewer.spectro_more_btn)
    viewer.spectro_more_menu = QtWidgets.QMenu(viewer.spectro_more_btn)
    _add_menu_widget(viewer.spectro_more_menu, viewer.grid_as_matrix_cb)
    _add_menu_widget(viewer.spectro_more_menu, viewer.force_single_cb)
    viewer.spectro_more_btn.setMenu(viewer.spectro_more_menu)
    viewer.spec_selection_label = QtWidgets.QLabel("Selected: 0")
    font_small = QtGui.QFont(UI_FONT_FAMILY, 9)
    viewer.spec_selection_label.setFont(font_small)
    viewer.spec_selection_label.setMinimumWidth(0)
    viewer.spec_selection_label.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
    layout.addWidget(viewer.show_spectra_cb)
    layout.addWidget(viewer.clear_spec_selection_btn)
    layout.addWidget(viewer.spectro_more_btn)
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
    viewer.spectro_overlay_act = viewer.display_menu.addAction("Show spectroscopy overlays")
    viewer.spectro_overlay_act.setCheckable(True)
    viewer.spectro_overlay_act.setChecked(viewer.show_spectra)
    viewer.spectro_overlay_act.setToolTip("Toggle spectroscopy overlays in thumbnails and preview")
    viewer.spectro_overlay_act.toggled.connect(viewer.on_show_spectra_toggled)
    viewer.molecules_act = viewer.display_menu.addAction("Show molecules")
    viewer.molecules_act.setCheckable(True)
    viewer.molecules_act.setChecked(getattr(viewer, "show_molecules", True))
    viewer.molecules_act.setToolTip("Toggle molecular overlays in the preview")
    viewer.molecules_act.toggled.connect(viewer.on_show_molecules_toggled)
    viewer.acquisition_overlay_act = viewer.display_menu.addAction("Show acquisition overlay")
    viewer.acquisition_overlay_act.setCheckable(True)
    viewer.acquisition_overlay_act.setChecked(getattr(viewer, "show_acquisition_overlay", False))
    viewer.acquisition_overlay_act.setToolTip("Show CC/CH acquisition parameters in the top-right of preview and pop-ups")
    viewer.acquisition_overlay_act.toggled.connect(viewer.on_show_acquisition_overlay_toggled)
    viewer.fixed_crop_quick_act = viewer.display_menu.addAction("Quick crop mode")
    viewer.fixed_crop_quick_act.setCheckable(True)
    viewer.fixed_crop_quick_act.setChecked(getattr(viewer, "quick_crop_mode", False))
    viewer.fixed_crop_quick_act.setToolTip("Enable clicking to spawn fixed-size crops")
    viewer.fixed_crop_quick_act.toggled.connect(viewer.on_fixed_crop_quick_toggled)
    viewer.crop_template_act = viewer.display_menu.addAction("Show crop template")
    viewer.crop_template_act.setCheckable(True)
    viewer.crop_template_act.setChecked(getattr(viewer, "show_crop_template_overlay", False))
    viewer.crop_template_act.setToolTip("Overlay the reusable crop frame in the preview")
    viewer.crop_template_act.toggled.connect(viewer.on_show_crop_template_overlay_toggled)
    viewer.crop_history_act = viewer.display_menu.addAction("Show crop history")
    viewer.crop_history_act.setCheckable(True)
    viewer.crop_history_act.setChecked(True)
    viewer.crop_history_act.setVisible(False)
    viewer.display_menu.addSeparator()
    viewer.profile_label_menu = viewer.display_menu.addMenu("Profile labels")
    viewer.profile_label_group = QtWidgets.QActionGroup(viewer.profile_label_menu)
    viewer.profile_label_group.setExclusive(True)
    viewer.profile_label_actions = {}
    label_modes = [
        ("Length only", "length"),
        ("Full (L, dx, dy)", "full"),
        ("Hidden", "hidden"),
    ]
    current_mode = str(getattr(viewer, "profile_label_mode", "length") or "length").lower()
    for label_text, mode_key in label_modes:
        act = viewer.profile_label_menu.addAction(label_text)
        act.setCheckable(True)
        act.setChecked(current_mode == mode_key)
        act.triggered.connect(lambda checked, m=mode_key: checked and viewer.on_profile_label_mode_changed(m))
        viewer.profile_label_group.addAction(act)
        viewer.profile_label_actions[mode_key] = act
    viewer.display_menu.addSeparator()
    
    markers_menu = viewer.display_menu.addMenu("Marker Style")
    if hasattr(viewer, "_populate_marker_style_menu"):
        viewer._populate_marker_style_menu(markers_menu)
    viewer.highlight_glow_act = viewer.display_menu.addAction("Spectro highlight glow")
    viewer.highlight_glow_act.setCheckable(True)
    viewer.highlight_glow_act.setChecked(getattr(viewer, "spectro_highlight_glow", True))
    viewer.highlight_glow_act.setToolTip("Pulse the selected spectroscopy marker in thumbnails and preview")
    viewer.highlight_glow_act.toggled.connect(viewer.on_toggle_highlight_glow)

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
    viewer.display_menu.addSeparator()
    save_session_act = viewer.display_menu.addAction("Save session...")
    save_session_act.setToolTip("Save the full viewer state (virtual copies, overlays, profiles, etc.)")
    save_session_act.triggered.connect(viewer.on_save_session)
    load_session_act = viewer.display_menu.addAction("Load session...")
    load_session_act.setToolTip("Restore a previously saved viewer session")
    load_session_act.triggered.connect(viewer.on_load_session)
    viewer.load_session_recent_menu = viewer.display_menu.addMenu("Recent sessions")
    viewer._refresh_recent_session_dirs_menu()
    viewer.display_menu.addSeparator()
    collections_menu = viewer.display_menu.addMenu("Collections")
    open_collection_act = collections_menu.addAction("Open collection...")
    open_collection_act.setToolTip("Open a curated cross-folder collection of selected analysis views")
    open_collection_act.triggered.connect(viewer.on_open_collection)
    add_preview_collection_act = collections_menu.addAction("Add current preview...")
    add_preview_collection_act.setToolTip("Store the current preview into a collection file")
    add_preview_collection_act.triggered.connect(viewer.on_add_current_preview_to_collection)
    add_popup_collection_act = collections_menu.addAction("Add active pop-up...")
    add_popup_collection_act.setToolTip("Store the active preview pop-up into a collection file")
    add_popup_collection_act.triggered.connect(viewer.on_add_active_popup_to_collection)
    add_all_popups_collection_act = collections_menu.addAction("Add all open pop-ups...")
    add_all_popups_collection_act.setToolTip("Append all current preview pop-outs into the same collection")
    add_all_popups_collection_act.triggered.connect(viewer.on_add_all_popups_to_collection)
    add_crop_collection_act = collections_menu.addAction("Add selected crop history...")
    add_crop_collection_act.setToolTip("Append the selected crop-history snapshots into a collection")
    add_crop_collection_act.triggered.connect(viewer.on_add_selected_crops_to_collection)
    collections_menu.addSeparator()
    collection_help_act = collections_menu.addAction("What is a collection?")
    collection_help_act.setToolTip("Explain linked vs portable collections")
    collection_help_act.triggered.connect(viewer.on_collection_help)
    arrange_act = viewer.display_menu.addAction("Arrange pop-outs")
    arrange_act.setToolTip("Tile and align all open preview/spectroscopy/profile windows")
    arrange_act.triggered.connect(viewer.on_arrange_popouts)
    minimize_act = viewer.display_menu.addAction("Minimize pop-outs")
    minimize_act.setToolTip("Minimize all open preview/spectroscopy/profile windows (Ctrl+Shift+M)")
    minimize_act.setShortcut(QtGui.QKeySequence("Ctrl+Shift+M"))
    minimize_act.triggered.connect(viewer.on_minimize_popouts)
    restore_act = viewer.display_menu.addAction("Restore pop-outs")
    restore_act.setToolTip("Restore any minimized pop-out windows to their previous size")
    restore_act.triggered.connect(viewer.on_restore_popouts)
    close_all_act = viewer.display_menu.addAction("Close all pop-outs")
    close_all_act.setToolTip("Close all open preview/spectroscopy/profile windows")
    close_all_act.triggered.connect(viewer.on_close_popouts)
    return viewer.display_menu


def build_display_widget(viewer, parent):
    container = QtWidgets.QWidget(parent)
    layout = QtWidgets.QHBoxLayout(container)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.setSpacing(8)
    _ensure_display_menu(viewer)

    viewer.spectro_browser_btn = _configure_compact_control(QtWidgets.QPushButton("Spectro Browser", container))
    viewer.spectro_browser_btn.setToolTip("Open the spectroscopy browser")
    viewer.spectro_browser_btn.clicked.connect(lambda: viewer.open_spectro_browser())
    layout.addWidget(viewer.spectro_browser_btn)

    viewer.spectro_stats_label = QtWidgets.QLabel(
        "Spectra -- | Single -- | Matrix --", container
    )
    stats_font = QtGui.QFont(UI_FONT_FAMILY, 9)
    viewer.spectro_stats_label.setFont(stats_font)
    viewer.spectro_stats_label.setToolTip("Summary of spectroscopy content for the loaded folder")
    viewer.spectro_stats_label.setWordWrap(True)
    viewer.spectro_stats_label.setMinimumWidth(0)
    viewer.spectro_stats_label.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
    layout.addWidget(viewer.spectro_stats_label, 1)
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
