"""Layout builders for the main SXM Viewer window."""
from __future__ import annotations

from PyQt5 import QtCore, QtGui, QtWidgets

from .compact_histogram import CompactHistogramWidget
from .constants import UI_FONT_FAMILY
from .styles import MAIN_SHORTCUTS_PANEL_STYLE, lower_control_frame_style, mode_selector_style
from . import theme as ui_theme


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
        (viewer.MODE_SPECTRO, "Spectro", "Ctrl+Alt+S"),
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
    viewer.browse_molecules_btn = _configure_compact_control(QtWidgets.QToolButton(frame))
    viewer.browse_molecules_btn.setObjectName("modeAccessoryButton")
    viewer.browse_molecules_btn.setText("Molecules")
    viewer.browse_molecules_btn.setCheckable(True)
    viewer.browse_molecules_btn.setChecked(bool(getattr(viewer, "show_molecules", True)))
    viewer.browse_molecules_btn.setToolButtonStyle(QtCore.Qt.ToolButtonTextOnly)
    viewer.browse_molecules_btn.setPopupMode(QtWidgets.QToolButton.MenuButtonPopup)
    viewer.browse_molecules_btn.setToolTip(
        "Toggle molecule overlays. Use the arrow for load, recent, clear, and palette options. "
        "Click a molecule, then use X/Y/Z to rotate; Shift+X/Y/Z rotates the opposite way."
    )
    viewer.browse_molecules_btn.toggled.connect(viewer.on_show_molecules_toggled)
    viewer.browse_molecules_menu = QtWidgets.QMenu(viewer.browse_molecules_btn)
    viewer.browse_molecules_menu.aboutToShow.connect(viewer._populate_browse_molecules_menu)
    viewer.browse_molecules_btn.setMenu(viewer.browse_molecules_menu)
    top_row.addWidget(viewer.browse_molecules_btn)

    viewer.compact_histogram = CompactHistogramWidget(frame)
    viewer.compact_histogram.climChanged.connect(viewer._on_compact_histogram_clim_changed)
    viewer.compact_histogram.autoRequested.connect(viewer._on_compact_histogram_auto_requested)
    viewer.compact_histogram.resetRequested.connect(viewer._on_compact_histogram_reset_requested)
    top_row.addWidget(viewer.compact_histogram)

    viewer.let_the_robot_btn = _configure_compact_control(QtWidgets.QToolButton(frame))
    viewer.let_the_robot_btn.setObjectName("letTheRobotButton")
    viewer.let_the_robot_btn.setText("\U0001F916 Let the robot")
    viewer.let_the_robot_btn.setToolButtonStyle(QtCore.Qt.ToolButtonTextOnly)
    viewer.let_the_robot_btn.setToolTip(
        "Diagnose the current image (tilt, incomplete scan, glitched lines, spikes, "
        "noise, contrast) and apply whichever fixes apply, as visible/editable steps "
        "in the normal filter pipeline."
    )
    viewer.let_the_robot_btn.clicked.connect(lambda: viewer._on_let_the_robot_clicked(viewer.preview_canvas))
    top_row.addWidget(viewer.let_the_robot_btn)

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
    try:
        # Keep the lower pane sized from its live sizeHint so wrapped
        # spectroscopy text can reflow without pinning the preview to a stale
        # geometry.
        frame.setSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Maximum)
    except Exception:
        pass
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


def build_spec_selection_tray(viewer):
    """Small pill-shaped bar under the thumbnails that appears while spectroscopy
    markers are Shift+click-selected: "N spectra selected - Compare | Clear".
    Replaces the old behavior of auto-opening the comparison window as soon as a
    second spectrum was selected."""
    tray = QtWidgets.QFrame()
    tray.setObjectName("specSelectionTray")
    tray.setStyleSheet(
        "#specSelectionTray {"
        " border-radius: 12px;"
        " background-color: rgba(255, 170, 60, 0.16);"
        " border: 1px solid rgba(255, 190, 90, 0.65);"
        "}"
    )
    row = QtWidgets.QHBoxLayout(tray)
    row.setContentsMargins(10, 3, 6, 3)
    row.setSpacing(8)
    viewer.spec_selection_tray_label = QtWidgets.QLabel("")
    viewer.spec_selection_tray_label.setStyleSheet("font-weight: 600; background: transparent; border: none;")
    row.addWidget(viewer.spec_selection_tray_label)
    row.addStretch(1)
    viewer.spec_selection_tray_compare_btn = _configure_compact_control(QtWidgets.QPushButton("Compare"))
    viewer.spec_selection_tray_compare_btn.setToolTip("Open the selected spectra together in a comparison window")
    viewer.spec_selection_tray_compare_btn.setFocusPolicy(QtCore.Qt.NoFocus)
    viewer.spec_selection_tray_compare_btn.clicked.connect(
        lambda: viewer.spectro_compare_controller.open_multi_popup()
    )
    row.addWidget(viewer.spec_selection_tray_compare_btn)
    viewer.spec_selection_tray_clear_btn = _configure_compact_control(QtWidgets.QPushButton("Clear"))
    viewer.spec_selection_tray_clear_btn.setToolTip("Clear the spectroscopy selection")
    viewer.spec_selection_tray_clear_btn.setFocusPolicy(QtCore.Qt.NoFocus)
    viewer.spec_selection_tray_clear_btn.clicked.connect(viewer.on_clear_spec_selection)
    row.addWidget(viewer.spec_selection_tray_clear_btn)
    tray.setVisible(False)
    viewer.spec_selection_tray = tray
    return tray


def build_spectro_context_page(viewer):
    page = QtWidgets.QWidget()
    layout = QtWidgets.QHBoxLayout(page)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.setSpacing(6)
    viewer.show_spectra_cb = None
    viewer.open_current_site_btn = _configure_compact_control(QtWidgets.QPushButton("Current position"))
    viewer.open_current_site_btn.setToolTip("Open the position summary for the currently highlighted or selected spectroscopy point")
    viewer.review_low_conf_btn = _configure_compact_control(QtWidgets.QPushButton("Review low conf"))
    viewer.review_low_conf_btn.setToolTip("Open the spectroscopy browser focused on low-confidence image assignments")
    viewer.clear_spec_selection_btn = _configure_compact_control(QtWidgets.QPushButton("Clear selection"))
    viewer.clear_spec_selection_btn.setToolTip("Clear the multi-selection of spectroscopy points")
    viewer.grid_as_matrix_cb = None
    viewer.force_single_cb = None
    viewer.spectro_more_btn = None
    viewer.spectro_more_menu = None
    viewer.spec_selection_label = QtWidgets.QLabel("Selected: 0")
    font_small = QtGui.QFont(UI_FONT_FAMILY, 9)
    viewer.spec_selection_label.setFont(font_small)
    viewer.spec_selection_label.setMinimumWidth(0)
    viewer.spec_selection_label.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
    viewer.spectro_mode_hint_label = QtWidgets.QLabel("Toolbar `Spectroscopy` controls markers and miniatures. Use `Current position` and `Review low conf` to inspect linked spectra.")
    viewer.spectro_mode_hint_label.setFont(font_small)
    layout.addWidget(viewer.spectro_mode_hint_label)
    layout.addWidget(viewer.open_current_site_btn)
    layout.addWidget(viewer.review_low_conf_btn)
    layout.addWidget(viewer.clear_spec_selection_btn)
    layout.addWidget(viewer.spec_selection_label)
    layout.addStretch(1)
    return page


def _ensure_display_menu(viewer):
    if getattr(viewer, "display_menu", None):
        return viewer.display_menu
    viewer.display_menu = QtWidgets.QMenu(viewer)
    viewer.display_menu.setToolTipsVisible(True)
    viewer.display_filter_menu = viewer.display_menu.addMenu("Show only")
    viewer.display_filter_menu.setToolTipsVisible(True)
    viewer.display_filter_group = QtWidgets.QActionGroup(viewer.display_filter_menu)
    viewer.display_filter_group.setExclusive(True)
    viewer.display_filter_actions = {}
    for text, filt_label, shortcut, tip in (
        ("All images", "All", None, "Show every image in the folder"),
        ("Starred ★", "Starred", "Ctrl+Alt+F", "Show only starred favourites (press again to show all)"),
        ("Constant height (CH)", "Constant height", "Ctrl+Alt+H", "Show only constant-height images (press again to show all)"),
        ("Constant current (CC)", "Constant current", "Ctrl+Alt+C", "Show only constant-current images (press again to show all)"),
        ("With spectroscopy", "With spectroscopy", "Ctrl+Alt+P", "Show only images with linked spectroscopy points (press again to show all)"),
    ):
        act = viewer.display_filter_menu.addAction(text)
        act.setCheckable(True)
        act.setToolTip(tip)
        if shortcut:
            act.setShortcut(QtGui.QKeySequence(shortcut))
            act.setShortcutContext(QtCore.Qt.ApplicationShortcut)
        act.triggered.connect(lambda checked, lab=filt_label: viewer.set_thumb_filter_mode(lab, toggle=True))
        viewer.display_filter_group.addAction(act)
        viewer.display_filter_actions[filt_label] = act
        # Register on the window so the shortcut fires while the menu is closed.
        viewer.addAction(act)

    def _show_view_filter_help():
        QtWidgets.QMessageBox.information(
            viewer,
            "Favourites & view filters",
            "<h3>Starring favourites</h3>"
            "<p>Select one or more thumbnails and press <b>S</b>, or right-click and choose "
            "<b>★ Star</b>. A gold star badge marks starred images and a short star "
            "animation plays when you star one.</p>"
            "<p><b>Removing a star works the same way</b>: press <b>S</b> again on starred "
            "images (S always toggles), or right-click → <b>Remove star</b>.</p>"
            "<p>Stars are remembered across sessions — reopen the folder later and your "
            "favourites are still marked.</p>"
            "<h3>Showing only certain images</h3>"
            "<p><b>Ctrl+Alt+F</b> shows only starred images, <b>Ctrl+Alt+H</b> only "
            "constant-height (CH) images, <b>Ctrl+Alt+C</b> only constant-current (CC) "
            "images, and <b>Ctrl+Alt+P</b> only images with linked spectroscopy points. "
            "Pressing the same shortcut a second time shows all images again.</p>"
            "<p>The CH/CC views use the CH/CC tags — set them manually with the "
            "<b>Tag as CH / Tag as CC</b> buttons, or enable <b>Auto CH/CC</b> to detect "
            "them automatically.</p>"
            "<p>This <i>Show only</i> menu, the <b>Filter</b> dropdown above the thumbnails, "
            "and the shortcuts all control the same filter, so you can use whichever is "
            "most comfortable. The active filter is restored the next time you open the app.</p>"
            "<h3>Virtual copies</h3>"
            "<p>Virtual copies follow the image they came from: a copy of a CH/CC-tagged "
            "image counts as CH/CC too (and shows the same border badge), and a copy of a "
            "starred image starts out starred. So copies created while a filter is active "
            "stay visible in that view, and you can still star/unstar each copy "
            "independently afterwards.</p>",
        )

    viewer.display_filter_menu.addSeparator()
    viewer.display_filter_help_act = viewer.display_filter_menu.addAction("How favourites && filters work...")
    viewer.display_filter_help_act.setToolTip("Explains starring favourites, removing stars, and the view-filter shortcuts")
    viewer.display_filter_help_act.triggered.connect(_show_view_filter_help)

    def _sync_display_filter_actions():
        combo = getattr(viewer, "thumb_filter_combo", None)
        current = combo.currentText() if combo is not None else "All"
        for lab, act in viewer.display_filter_actions.items():
            act.blockSignals(True)
            act.setChecked(lab == current)
            act.blockSignals(False)

    viewer._sync_display_filter_actions = _sync_display_filter_actions
    viewer.display_menu.aboutToShow.connect(_sync_display_filter_actions)
    _sync_display_filter_actions()
    viewer.display_menu.addSeparator()
    viewer.display_units_si_act = viewer.display_menu.addAction("Show SI units")
    viewer.display_units_si_act.setCheckable(True)
    viewer.display_units_si_act.setChecked(bool(getattr(viewer, "display_units_si", False)))
    viewer.display_units_si_act.setToolTip("Show SI units in preview annotations")
    viewer.display_units_si_act.toggled.connect(viewer.on_unit_display_toggled)
    viewer.display_units_relative_act = viewer.display_menu.addAction("Values relative to zero")
    viewer.display_units_relative_act.setCheckable(True)
    viewer.display_units_relative_act.setChecked(bool(getattr(viewer, "display_units_relative", False)))
    viewer.display_units_relative_act.setToolTip("Display values relative to the current zero/reference")
    viewer.display_units_relative_act.toggled.connect(viewer.on_unit_relative_toggled)
    viewer.relative_axes_act = viewer.display_menu.addAction("Relative axes")
    viewer.relative_axes_act.setCheckable(True)
    viewer.relative_axes_act.setChecked(bool(getattr(viewer, "relative_axes", False)))
    viewer.relative_axes_act.setToolTip("Use relative axes in the preview")
    viewer.relative_axes_act.toggled.connect(viewer.on_relative_axes_toggled)
    viewer.display_scale_bar_act = viewer.display_menu.addAction("Scale bar")
    viewer.display_scale_bar_act.setCheckable(True)
    viewer.display_scale_bar_act.setChecked(bool(getattr(viewer, "config", {}).get("show_scale_bar", False)))
    viewer.display_scale_bar_act.setToolTip("Show the scale bar in preview and pop-outs")
    viewer.display_scale_bar_act.toggled.connect(viewer.on_scale_bar_toggled)
    viewer.reset_cmap_favorites_act = viewer.display_menu.addAction("Reset colormap defaults")
    viewer.reset_cmap_favorites_act.setToolTip(
        "Forget the favourite colormaps/color cycle saved with the ★ buttons "
        "(preview, thumbnails, grid map, reports) and go back to the built-in defaults")
    viewer.reset_cmap_favorites_act.triggered.connect(viewer.clear_colormap_favorites)
    viewer.extra_cmaps_status_act = viewer.display_menu.addAction("Extra colormaps...")
    viewer.extra_cmaps_status_act.setToolTip(
        "Status of the optional 'colormaps' package (pratiman-91) — when "
        "installed, its colormaps appear in every colormap picker")
    viewer.extra_cmaps_status_act.triggered.connect(viewer.on_extra_colormaps_status)
    viewer.display_menu.addSeparator()
    viewer.molecules_act = viewer.display_menu.addAction("Show molecules")
    viewer.molecules_act.setCheckable(True)
    viewer.molecules_act.setChecked(getattr(viewer, "show_molecules", True))
    viewer.molecules_act.setToolTip("Toggle molecular overlays in the preview")
    viewer.molecules_act.toggled.connect(viewer.on_show_molecules_toggled)
    viewer.display_molecule_gizmo_act = viewer.display_menu.addAction("Molecule gizmo")
    viewer.display_molecule_gizmo_act.setCheckable(True)
    viewer.display_molecule_gizmo_act.setChecked(bool(getattr(viewer, "show_molecule_gizmo", False)))
    viewer.display_molecule_gizmo_act.setToolTip("Show a small orientation gizmo for the active molecule")
    viewer.display_molecule_gizmo_act.toggled.connect(
        lambda checked: viewer._apply_canvas_display_options(
            {
                **viewer._canvas_display_state_from_canvas(getattr(viewer, "preview_canvas", None)),
                "show_molecule_gizmo": bool(checked),
            },
            source_canvas=getattr(viewer, "preview_canvas", None),
            persist=True,
        )
    )
    viewer.acquisition_overlay_act = viewer.display_menu.addAction("Show acquisition overlay")
    viewer.acquisition_overlay_act.setCheckable(True)
    viewer.acquisition_overlay_act.setChecked(getattr(viewer, "show_acquisition_overlay", False))
    viewer.acquisition_overlay_act.setToolTip("Show CC/CH acquisition parameters in the top-right of preview and pop-ups")
    viewer.acquisition_overlay_act.toggled.connect(viewer.on_show_acquisition_overlay_toggled)
    viewer.filter_summary_act = viewer.display_menu.addAction("Show filter label")
    viewer.filter_summary_act.setCheckable(True)
    viewer.filter_summary_act.setChecked(getattr(viewer, "show_filter_summary", True))
    viewer.filter_summary_act.setToolTip("Show the applied-filter label in the top-left of preview and pop-ups")
    viewer.filter_summary_act.toggled.connect(viewer.on_show_filter_summary_toggled)
    viewer.fixed_crop_quick_act = viewer.display_menu.addAction("Crop template mode")
    viewer.fixed_crop_quick_act.setCheckable(True)
    viewer.fixed_crop_quick_act.setChecked(getattr(viewer, "quick_crop_mode", False))
    viewer.fixed_crop_quick_act.setToolTip("Enable clicking to spawn repeated crops from the current template")
    viewer.fixed_crop_quick_act.toggled.connect(viewer.on_fixed_crop_quick_toggled)
    viewer.crop_template_act = viewer.display_menu.addAction("Show crop template")
    viewer.crop_template_act.setCheckable(True)
    viewer.crop_template_act.setChecked(getattr(viewer, "show_crop_template_overlay", False))
    viewer.crop_template_act.setToolTip("Overlay the reusable crop template in the preview")
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


def _ensure_tools_menu(viewer):
    if getattr(viewer, "tools_menu", None):
        return viewer.tools_menu
    viewer.tools_menu = QtWidgets.QMenu(viewer)

    viewer.tools_load_molecule_act = viewer.tools_menu.addAction("Load molecule...")
    viewer.tools_load_molecule_act.setToolTip("Load a molecular structure overlay onto the preview canvas")
    viewer.tools_load_molecule_act.triggered.connect(viewer.on_load_molecule)
    viewer.tools_menu.addSeparator()

    viewer.tools_preview_detach_act = viewer.tools_menu.addAction("Float preview")
    viewer.tools_preview_detach_act.setToolTip("Detach the preview pane into its own floating window")
    viewer.tools_preview_detach_act.triggered.connect(viewer.on_toggle_preview_detach)
    viewer.tools_preview_lock_act = viewer.tools_menu.addAction("Lock preview")
    viewer.tools_preview_lock_act.setCheckable(True)
    viewer.tools_preview_lock_act.setChecked(bool(getattr(viewer, "preview_locked", False)))
    viewer.tools_preview_lock_act.setToolTip("Lock the preview inside the main window")
    viewer.tools_preview_lock_act.toggled.connect(viewer.on_preview_lock_toggled)
    viewer.tools_menu.addSeparator()

    viewer.tools_reopen_window_act = viewer.tools_menu.addAction("Reopen closed window")
    viewer.tools_reopen_window_act.setToolTip("Restore the most recently closed popup/tool window (up to 6 levels)")
    viewer.tools_reopen_window_act.setShortcut(QtGui.QKeySequence("Ctrl+Z"))
    viewer.tools_reopen_window_act.triggered.connect(viewer._restore_last_closed_window)

    recovery_menu = viewer.tools_menu.addMenu("Recovery")
    viewer.session_recovery_status_act = recovery_menu.addAction("Autosave recovery: --")
    viewer.session_recovery_status_act.setEnabled(False)
    viewer.session_recovery_enable_act = recovery_menu.addAction("Enable autosave recovery")
    viewer.session_recovery_enable_act.setCheckable(True)
    viewer.session_recovery_enable_act.toggled.connect(viewer.on_toggle_session_recovery)
    interval_menu = recovery_menu.addMenu("Autosave interval")
    interval_group = QtWidgets.QActionGroup(interval_menu)
    viewer.session_recovery_interval_actions = {}
    for minutes in (2, 5, 10, 15, 30):
        act = interval_menu.addAction(f"{minutes} min")
        act.setCheckable(True)
        act.triggered.connect(lambda checked=False, m=minutes: viewer.on_set_session_recovery_interval(m))
        interval_group.addAction(act)
        viewer.session_recovery_interval_actions[int(minutes)] = act
        if minutes == getattr(viewer, "_session_recovery_interval_min", 5):
            act.setChecked(True)
    recovery_menu.addSeparator()
    viewer.session_recovery_open_act = recovery_menu.addAction("Recover latest autosave now")
    viewer.session_recovery_open_act.triggered.connect(viewer.on_recover_latest_autosave)
    viewer.session_recovery_discard_act = recovery_menu.addAction("Discard autosaved recovery")
    viewer.session_recovery_discard_act.triggered.connect(viewer.on_discard_recovery_snapshot)
    return viewer.tools_menu


def build_display_widget(viewer, parent):
    container = QtWidgets.QWidget(parent)
    layout = QtWidgets.QVBoxLayout(container)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.setSpacing(4)
    _ensure_display_menu(viewer)

    viewer.spectro_section_title = QtWidgets.QLabel("Spectroscopy", container)
    viewer.spectro_section_title.setFont(QtGui.QFont(UI_FONT_FAMILY, 10, QtGui.QFont.Bold))
    viewer.spectro_section_title.setToolTip("Use the top toolbar `Spectroscopy` button for spectroscopy display controls and browser access")
    layout.addWidget(viewer.spectro_section_title)
    viewer.spectro_thumbnail_markers_cb = None
    viewer.spectro_preview_markers_cb = None
    viewer.spectro_miniatures_cb = None
    viewer.spectro_browser_btn = None

    viewer.spectro_hint_label = QtWidgets.QLabel(
        "Use the top toolbar `Spectroscopy` button to open the browser and control markers and miniatures.",
        container,
    )
    hint_font = QtGui.QFont(UI_FONT_FAMILY, 9)
    viewer.spectro_hint_label.setFont(hint_font)
    viewer.spectro_hint_label.setWordWrap(True)
    viewer.spectro_hint_label.setToolTip("Thumbnail markers and preview markers show point positions. Miniatures show spectroscopy traces as their own cards.")
    layout.addWidget(viewer.spectro_hint_label)

    viewer.spectro_stats_label = QtWidgets.QLabel(
        "Spectroscopy pending load", container
    )
    stats_font = QtGui.QFont(UI_FONT_FAMILY, 9)
    viewer.spectro_stats_label.setFont(stats_font)
    viewer.spectro_stats_label.setToolTip("Summary of spectroscopy content for the loaded folder")
    viewer.spectro_stats_label.setWordWrap(True)
    viewer.spectro_stats_label.setMinimumWidth(0)
    viewer.spectro_stats_label.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
    layout.addWidget(viewer.spectro_stats_label)
    return container


def apply_lower_control_theme(viewer):
    frame = getattr(viewer, "lower_control_frame", None)
    mode_widget = getattr(viewer, "mode_selector_widget", None)
    molecules_btn = getattr(viewer, "browse_molecules_btn", None)
    if frame is None:
        return
    dark = bool(getattr(viewer, "dark_mode", False))
    amber = ui_theme.is_amber(viewer)
    mode_checked_text = "#ffffff"
    if amber:
        t = ui_theme.AMBER
        border = t["border"]
        bg = t["panel_bg"]
        mode_border = t["border"]
        mode_text = t["text_primary"]
        mode_checked = t["selection_bg"]
        mode_checked_text = t["amber_bright"]
    elif dark:
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
        mode_widget.setStyleSheet(
            mode_selector_style(mode_border, mode_text, mode_checked, mode_checked_text)
        )
    if molecules_btn is not None:
        molecules_btn.setStyleSheet(
            "QToolButton#modeAccessoryButton {"
            f" border: 1px solid {mode_border};"
            " padding: 6px 12px;"
            " background: transparent;"
            f" color: {mode_text};"
            "}"
            "QToolButton#modeAccessoryButton:checked {"
            f" background: {mode_checked};"
            f" color: {mode_checked_text};"
            "}"
        )


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
