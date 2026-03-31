"""UI construction helpers for the canvas window."""
from __future__ import annotations

from ..._shared import QtCore, QtGui, QtWidgets, colormaps
from ..styles import (
    CANVAS_DIALOG_STYLE,
    CANVAS_DIALOG_STYLE_LIGHT,
    CANVAS_HEADER_STYLE,
    CANVAS_LABEL_STYLE,
    CANVAS_LABEL_VALUE_STYLE,
    CANVAS_RANGE_LABEL_STYLE,
    CANVAS_RANGE_TO_LABEL_STYLE,
    CANVAS_REMOVE_BUTTON_STYLE,
    CANVAS_SECTION_STYLE,
    CANVAS_STATUS_LABEL_STYLE,
    CANVAS_STATS_LABEL_STYLE,
    CANVAS_TOOLBAR_GROUP_LABEL_STYLE,
    CANVAS_TOOLBAR_GROUP_STYLE,
    CANVAS_TOOLBAR_SEPARATOR_STYLE,
    CANVAS_TOOLBAR_WIDGET_STYLE,
)
from ..thumbnail_render import _colormap_icon
from ..text import (
    CANVAS_BUTTON_TEXT,
    CANVAS_CHECKBOX_TEXT,
    CANVAS_LABEL_TEXT,
    CANVAS_PLACEHOLDER_TEXT,
    CANVAS_TOOLBAR_GROUP_TITLES,
    CANVAS_TOOLTIPS,
)


def create_toolbar_group(title):
    """Create a visually grouped section in toolbar."""
    container = QtWidgets.QWidget()
    container.setProperty("toolbarGroup", True)
    container_layout = QtWidgets.QVBoxLayout(container)
    container_layout.setContentsMargins(0, 0, 0, 0)
    container_layout.setSpacing(2)

    label = QtWidgets.QLabel(title)
    label.setProperty("toolbarLabel", True)
    # Style via global stylesheet; avoid hardcoded palette.
    label.setStyleSheet("")
    label.setAlignment(QtCore.Qt.AlignLeft)
    container_layout.addWidget(label)

    group = QtWidgets.QWidget()
    # Styling comes from the theme; avoid fixed colors.
    group.setStyleSheet("")
    group_layout = QtWidgets.QHBoxLayout(group)
    group_layout.setContentsMargins(8, 4, 8, 4)
    group_layout.setSpacing(4)
    container_layout.addWidget(group)
    return container, group_layout


def create_separator():
    """Create a vertical separator line."""
    separator = QtWidgets.QFrame()
    separator.setFrameShape(QtWidgets.QFrame.VLine)
    separator.setFrameShadow(QtWidgets.QFrame.Sunken)
    separator.setStyleSheet(CANVAS_TOOLBAR_SEPARATOR_STYLE)
    return separator


def create_toolbar_section(title: str, widgets: list) -> QtWidgets.QWidget:
    """Create a visually grouped toolbar section."""
    section = QtWidgets.QWidget()
    section.setProperty("canvasQuickSection", True)
    section.setStyleSheet("")
    layout = QtWidgets.QHBoxLayout(section)
    layout.setContentsMargins(5, 3, 5, 3)
    layout.setSpacing(3)

    if title:
        label = QtWidgets.QLabel(title)
        label.setProperty("canvasQuickLabel", True)
        label.setStyleSheet("")
        layout.addWidget(label)

    for widget in widgets:
        layout.addWidget(widget)

    return section


def build_toolbar(window):
    """Build a compact top strip for the primary canvas actions."""
    toolbar_widget = QtWidgets.QWidget()
    toolbar_widget.setObjectName("canvasToolbar")
    toolbar_widget.setStyleSheet("")
    toolbar_widget.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)

    main_layout = QtWidgets.QHBoxLayout(toolbar_widget)
    main_layout.setContentsMargins(6, 5, 6, 5)
    main_layout.setSpacing(6)

    window.save_btn = QtWidgets.QPushButton(CANVAS_BUTTON_TEXT["save"])
    window.save_btn.setToolTip(CANVAS_TOOLTIPS["save"])
    window.load_btn = QtWidgets.QPushButton(CANVAS_BUTTON_TEXT["load"])
    window.load_btn.setToolTip(CANVAS_TOOLTIPS["load"])
    window.export_btn = QtWidgets.QPushButton(CANVAS_BUTTON_TEXT["export"])
    window.export_btn.setToolTip(CANVAS_TOOLTIPS["export"])
    main_layout.addWidget(create_toolbar_section("File", [window.save_btn, window.load_btn, window.export_btn]))

    window.layout_2x2_btn = QtWidgets.QPushButton(CANVAS_BUTTON_TEXT["layout_2x2"])
    window.layout_2x2_btn.setToolTip(CANVAS_TOOLTIPS["layout_2x2"])
    window.layout_1x3_btn = QtWidgets.QPushButton(CANVAS_BUTTON_TEXT["layout_1x3"])
    window.layout_1x3_btn.setToolTip(CANVAS_TOOLTIPS["layout_1x3"])
    window.layout_3x1_btn = QtWidgets.QPushButton(CANVAS_BUTTON_TEXT["layout_3x1"])
    window.layout_3x1_btn.setToolTip(CANVAS_TOOLTIPS["layout_3x1"])
    main_layout.addWidget(create_toolbar_section("Layout", [window.layout_2x2_btn, window.layout_1x3_btn, window.layout_3x1_btn]))

    window.show_grid_check = QtWidgets.QCheckBox(CANVAS_CHECKBOX_TEXT["grid"])
    window.show_grid_check.setToolTip(CANVAS_TOOLTIPS["grid"])
    window.snap_grid_check = QtWidgets.QCheckBox(CANVAS_CHECKBOX_TEXT["snap"])
    window.snap_grid_check.setToolTip(CANVAS_TOOLTIPS["snap"])
    window.canvas_color_btn = QtWidgets.QPushButton(CANVAS_BUTTON_TEXT["canvas_color"])
    window.canvas_color_btn.setToolTip(CANVAS_TOOLTIPS["canvas_color"])
    window.canvas_color_btn.setMaximumWidth(60)

    window.display_preset_combo = QtWidgets.QComboBox()
    for label in ("Custom", "Clean", "Analysis", "Publication"):
        window.display_preset_combo.addItem(label)
    window.display_preset_combo.setMaximumWidth(110)
    window.apply_preset_btn = QtWidgets.QPushButton("Apply")
    main_layout.addWidget(create_toolbar_section("Preset", [window.display_preset_combo, window.apply_preset_btn]))

    window.show_colorbar_check = QtWidgets.QCheckBox(CANVAS_CHECKBOX_TEXT["colorbar"])
    window.show_colorbar_check.setToolTip(CANVAS_TOOLTIPS["colorbar"])
    window.show_colorbar_check.setChecked(window._global_show_colorbar)
    window.sync_cbar_check = QtWidgets.QCheckBox(CANVAS_CHECKBOX_TEXT["ranges"])
    window.sync_cbar_check.setToolTip(CANVAS_TOOLTIPS["ranges"])
    window.sync_by_channel_check = QtWidgets.QCheckBox(CANVAS_CHECKBOX_TEXT["colors_by_channel"])
    window.sync_by_channel_check.setChecked(window._sync_by_channel)
    window.sync_by_channel_check.setToolTip(CANVAS_TOOLTIPS["colors_by_channel"])
    window.overlay_info_check = QtWidgets.QCheckBox(CANVAS_CHECKBOX_TEXT["overlay_info"])
    window.overlay_info_check.setChecked(window._show_overlay_info)
    window.overlay_info_check.setToolTip(CANVAS_TOOLTIPS["overlay_info"])
    window.overlay_file_check = QtWidgets.QCheckBox(CANVAS_CHECKBOX_TEXT["overlay_file"])
    window.overlay_file_check.setChecked(window._show_overlay_file)
    window.overlay_file_check.setToolTip(CANVAS_TOOLTIPS["overlay_file"])
    window.toolbar_show_title_check = QtWidgets.QCheckBox("Title")
    window.toolbar_show_title_check.setChecked(window._global_show_title)
    window.toolbar_metadata_bar_check = QtWidgets.QCheckBox("Meta")
    window.toolbar_metadata_bar_check.setChecked(window._metadata_bar_default)
    window.toolbar_unit_badge_check = QtWidgets.QCheckBox("Unit")
    window.toolbar_unit_badge_check.setChecked(window._metadata_unit_default)
    window.toolbar_scale_bar_check = QtWidgets.QCheckBox("Scale")
    window.toolbar_scale_bar_check.setChecked(window._global_show_scale_bar)

    window.colorbar_ticks_check = QtWidgets.QCheckBox(CANVAS_CHECKBOX_TEXT["colorbar_ticks"])
    window.colorbar_ticks_check.setToolTip(CANVAS_TOOLTIPS["colorbar_ticks"])
    window.colorbar_ticks_check.setChecked(window._global_show_colorbar_ticks)
    window.colorbar_mode_combo = QtWidgets.QComboBox()
    for label in ("Bottom", "Top", "Right", "Left", "Inset", "Hidden"):
        window.colorbar_mode_combo.addItem(label)

    window.align_btn = QtWidgets.QPushButton(CANVAS_BUTTON_TEXT["align_selected"])
    window.align_btn.setToolTip(CANVAS_TOOLTIPS["align_selected"])
    window.align_channels_btn = QtWidgets.QPushButton(CANVAS_BUTTON_TEXT["align_by_channel"])
    window.align_channels_btn.setToolTip(CANVAS_TOOLTIPS["align_by_channel"])
    window.polish_btn = QtWidgets.QPushButton(CANVAS_BUTTON_TEXT["polish_layout"])
    window.polish_btn.setToolTip(CANVAS_TOOLTIPS["polish_layout"])
    window.reset_alignment_btn = QtWidgets.QPushButton(CANVAS_BUTTON_TEXT["reset_alignment"])
    window.reset_alignment_btn.setToolTip(CANVAS_TOOLTIPS["reset_alignment"])
    main_layout.addWidget(create_toolbar_section("Arrange", [window.align_btn, window.align_channels_btn, window.polish_btn, window.reset_alignment_btn]))
    main_layout.addStretch(1)

    window.save_btn.clicked.connect(window._on_save_canvas)
    window.load_btn.clicked.connect(window._on_load_canvas)
    window.export_btn.clicked.connect(window._on_export_image)
    window.layout_2x2_btn.clicked.connect(lambda: window._apply_layout("2x2"))
    window.layout_1x3_btn.clicked.connect(lambda: window._apply_layout("1x3"))
    window.layout_3x1_btn.clicked.connect(lambda: window._apply_layout("3x1"))
    window.show_grid_check.toggled.connect(window.view.set_show_grid)
    window.snap_grid_check.toggled.connect(window.view.set_snap_to_grid)
    window.sync_cbar_check.toggled.connect(window._on_sync_colorbars_toggled)
    window.sync_by_channel_check.toggled.connect(window._on_sync_by_channel_toggled)
    window.overlay_info_check.toggled.connect(window._on_overlay_info_toggled)
    window.overlay_file_check.toggled.connect(window._on_overlay_file_toggled)
    window.toolbar_show_title_check.toggled.connect(window._on_global_show_title_toggled)
    window.toolbar_metadata_bar_check.toggled.connect(window._on_metadata_bar_toggled)
    window.toolbar_unit_badge_check.toggled.connect(window._on_metadata_unit_toggled)
    window.toolbar_scale_bar_check.toggled.connect(window._on_scale_bar_toggled)
    window.colorbar_ticks_check.toggled.connect(window._on_global_show_colorbar_ticks_toggled)
    window.colorbar_mode_combo.currentTextChanged.connect(window._on_colorbar_position_changed)
    window.canvas_color_btn.clicked.connect(window._on_canvas_color_clicked)
    window.apply_preset_btn.clicked.connect(window._on_apply_display_preset_clicked)
    window.align_btn.clicked.connect(window._on_align_selected)
    window.align_channels_btn.clicked.connect(window._on_align_by_channels)
    window.polish_btn.clicked.connect(window._on_polish_layout)
    window.reset_alignment_btn.clicked.connect(window._reset_locked_alignment)

    window._on_overlay_info_toggled(window.overlay_info_check.isChecked())
    window._on_overlay_file_toggled(window.overlay_file_check.isChecked())
    window.colorbar_mode_combo.setCurrentText(window._colorbar_mode.capitalize())

    return toolbar_widget


def build_left_controls(window):
    """Build a compact left rail with the frequently toggled canvas options."""
    panel = QtWidgets.QScrollArea()
    panel.setWidgetResizable(True)
    panel.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
    panel.setObjectName("canvasLeftRail")
    panel.setMinimumWidth(160)
    panel.setMaximumWidth(190)

    body = QtWidgets.QWidget()
    layout = QtWidgets.QVBoxLayout(body)
    layout.setContentsMargins(10, 10, 10, 10)
    layout.setSpacing(10)

    quick_header = QtWidgets.QLabel("Quick Controls")
    quick_header.setProperty("canvasHeader", True)
    quick_header.setStyleSheet("")
    layout.addWidget(quick_header)

    canvas_group = QtWidgets.QGroupBox("Canvas")
    canvas_layout = QtWidgets.QVBoxLayout(canvas_group)
    canvas_layout.setSpacing(6)
    canvas_layout.addWidget(window.show_grid_check)
    canvas_layout.addWidget(window.snap_grid_check)
    canvas_layout.addWidget(window.canvas_color_btn)
    layout.addWidget(canvas_group)

    display_group = QtWidgets.QGroupBox("Display")
    display_layout = QtWidgets.QVBoxLayout(display_group)
    display_layout.setSpacing(6)
    for widget in (
        window.toolbar_show_title_check,
        window.toolbar_metadata_bar_check,
        window.toolbar_unit_badge_check,
        window.toolbar_scale_bar_check,
        window.overlay_info_check,
        window.overlay_file_check,
        window.show_colorbar_check,
    ):
        display_layout.addWidget(widget)
    layout.addWidget(display_group)

    scale_group = QtWidgets.QGroupBox("Scale Bar")
    scale_layout = QtWidgets.QVBoxLayout(scale_group)
    scale_layout.setSpacing(6)
    window.left_scale_bar_combo = QtWidgets.QComboBox()
    window.left_scale_bar_combo.addItem("Auto")
    for label in ("0.5 nm", "1 nm", "2 nm", "3 nm", "5 nm", "10 nm", "20 nm", "50 nm", "100 nm"):
        window.left_scale_bar_combo.addItem(label)
    scale_layout.addWidget(window.left_scale_bar_combo)
    layout.addWidget(scale_group)

    molecules_group = QtWidgets.QGroupBox("Molecules")
    molecules_layout = QtWidgets.QVBoxLayout(molecules_group)
    molecules_layout.setSpacing(6)
    window.canvas_molecules_check = QtWidgets.QCheckBox("Show")
    window.canvas_molecules_check.setChecked(bool(getattr(window, "_global_show_molecules", False)))
    window.canvas_molecule_load_btn = QtWidgets.QPushButton("Load...")
    window.canvas_molecule_clear_btn = QtWidgets.QPushButton("Clear")
    molecules_layout.addWidget(window.canvas_molecules_check)
    molecules_layout.addWidget(window.canvas_molecule_load_btn)
    molecules_layout.addWidget(window.canvas_molecule_clear_btn)
    layout.addWidget(molecules_group)

    colorbar_group = QtWidgets.QGroupBox("Colorbar")
    colorbar_layout = QtWidgets.QVBoxLayout(colorbar_group)
    colorbar_layout.setSpacing(6)
    colorbar_layout.addWidget(window.colorbar_ticks_check)
    colorbar_pos_label = QtWidgets.QLabel(CANVAS_LABEL_TEXT["position"])
    colorbar_pos_label.setStyleSheet("")
    colorbar_layout.addWidget(colorbar_pos_label)
    colorbar_layout.addWidget(window.colorbar_mode_combo)
    layout.addWidget(colorbar_group)

    sync_group = QtWidgets.QGroupBox("Sync")
    sync_layout = QtWidgets.QVBoxLayout(sync_group)
    sync_layout.setSpacing(6)
    sync_layout.addWidget(window.sync_cbar_check)
    sync_layout.addWidget(window.sync_by_channel_check)
    layout.addWidget(sync_group)

    layout.addStretch(1)
    panel.setWidget(body)
    window.canvas_molecules_check.toggled.connect(window._on_canvas_show_molecules_toggled)
    window.canvas_molecule_load_btn.clicked.connect(window._on_canvas_load_molecule)
    window.canvas_molecule_clear_btn.clicked.connect(window._on_canvas_clear_molecules)
    window.left_scale_bar_combo.currentTextChanged.connect(window._on_scale_bar_size_changed)
    return panel


def build_inspector(window):
    """Build inspector panel with better visual hierarchy."""
    scroll = QtWidgets.QScrollArea()
    scroll.setWidgetResizable(True)
    scroll.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)

    panel = QtWidgets.QWidget()
    panel.setObjectName("inspectorPanel")
    layout = QtWidgets.QVBoxLayout(panel)
    layout.setContentsMargins(12, 12, 12, 12)
    layout.setSpacing(12)

    # Header with higher contrast
    header = QtWidgets.QLabel(CANVAS_LABEL_TEXT["selected_item"])
    header.setProperty("canvasHeader", True)
    header.setStyleSheet("")
    layout.addWidget(header)

    window.selection_hint = QtWidgets.QLabel("Select a tile to edit its labels, scale bar, colormap and export settings.")
    window.selection_hint.setWordWrap(True)
    window.selection_hint.setStyleSheet("")
    layout.addWidget(window.selection_hint)

    tabs = QtWidgets.QTabWidget()
    tabs.setDocumentMode(True)
    tabs.setTabPosition(QtWidgets.QTabWidget.North)
    tile_tab = QtWidgets.QWidget()
    tile_layout = QtWidgets.QVBoxLayout(tile_tab)
    tile_layout.setContentsMargins(0, 8, 0, 0)
    tile_layout.setSpacing(12)
    data_tab = QtWidgets.QWidget()
    data_layout = QtWidgets.QVBoxLayout(data_tab)
    data_layout.setContentsMargins(0, 8, 0, 0)
    data_layout.setSpacing(12)

    # Item Info Group
    info_group = QtWidgets.QGroupBox(CANVAS_LABEL_TEXT["item_info"])
    info_layout = QtWidgets.QFormLayout()
    info_layout.setLabelAlignment(QtCore.Qt.AlignRight)
    info_layout.setVerticalSpacing(10)
    info_layout.setHorizontalSpacing(12)

    file_label_text = QtWidgets.QLabel(CANVAS_LABEL_TEXT["file"])
    file_label_text.setStyleSheet("")
    window.file_label = QtWidgets.QLabel("-")
    window.file_label.setWordWrap(True)
    window.file_label.setStyleSheet("")
    info_layout.addRow(file_label_text, window.file_label)

    channel_label_text = QtWidgets.QLabel(CANVAS_LABEL_TEXT["channel"])
    channel_label_text.setStyleSheet("")
    window.channel_label = QtWidgets.QLabel("-")
    window.channel_label.setStyleSheet("")
    info_layout.addRow(channel_label_text, window.channel_label)

    info_group.setLayout(info_layout)
    tile_layout.addWidget(info_group)

    # Appearance Group
    appearance_group = QtWidgets.QGroupBox(CANVAS_LABEL_TEXT["appearance"])
    appearance_layout = QtWidgets.QFormLayout()
    appearance_layout.setLabelAlignment(QtCore.Qt.AlignRight)
    appearance_layout.setVerticalSpacing(10)
    appearance_layout.setHorizontalSpacing(12)

    colorbar_label = QtWidgets.QLabel(CANVAS_LABEL_TEXT["label"])
    colorbar_label.setStyleSheet("")
    window.colorbar_edit = QtWidgets.QLineEdit()
    window.colorbar_edit.setPlaceholderText(CANVAS_PLACEHOLDER_TEXT["colorbar_label"])
    appearance_layout.addRow(colorbar_label, window.colorbar_edit)

    font_color_label = QtWidgets.QLabel(CANVAS_LABEL_TEXT["font_color"])
    font_color_label.setStyleSheet("")
    font_color_row = QtWidgets.QWidget()
    font_color_layout = QtWidgets.QHBoxLayout(font_color_row)
    font_color_layout.setContentsMargins(0, 0, 0, 0)
    font_color_layout.setSpacing(6)
    window.font_color_auto_check = QtWidgets.QCheckBox(CANVAS_CHECKBOX_TEXT["font_color_auto"])
    window.font_color_auto_check.setChecked(True)
    window.font_color_btn = QtWidgets.QPushButton("Pick")
    window.font_color_btn.setMaximumWidth(80)
    font_color_layout.addWidget(window.font_color_auto_check)
    font_color_layout.addWidget(window.font_color_btn)
    appearance_layout.addRow(font_color_label, font_color_row)

    text_scale_label = QtWidgets.QLabel(CANVAS_LABEL_TEXT["text_scale"])
    text_scale_label.setStyleSheet("")
    window.text_scale_slider = QtWidgets.QSlider(QtCore.Qt.Horizontal)
    window.text_scale_slider.setMinimum(1)
    window.text_scale_slider.setMaximum(240)
    window.text_scale_slider.setValue(int(round(window._global_text_scale * 100)))
    window.text_scale_slider.setEnabled(True)
    appearance_layout.addRow(text_scale_label, window.text_scale_slider)

    scale_bar_label = QtWidgets.QLabel(CANVAS_LABEL_TEXT["scale_bar"])
    scale_bar_label.setStyleSheet("")
    scale_bar_row = QtWidgets.QWidget()
    scale_bar_layout = QtWidgets.QHBoxLayout(scale_bar_row)
    scale_bar_layout.setContentsMargins(0, 0, 0, 0)
    scale_bar_layout.setSpacing(6)
    window.scale_bar_check = QtWidgets.QCheckBox(CANVAS_CHECKBOX_TEXT["scale_bar"])
    window.scale_bar_check.setChecked(window._global_show_scale_bar)
    window.scale_bar_combo = QtWidgets.QComboBox()
    window.scale_bar_combo.addItem("Auto")
    for label in ("0.5 nm", "1 nm", "2 nm", "3 nm", "5 nm", "10 nm", "20 nm", "50 nm", "100 nm"):
        window.scale_bar_combo.addItem(label)
    scale_bar_layout.addWidget(window.scale_bar_check)
    scale_bar_layout.addWidget(window.scale_bar_combo)
    appearance_layout.addRow(scale_bar_label, scale_bar_row)

    appearance_group.setLayout(appearance_layout)
    tile_layout.addWidget(appearance_group)

    # Colormap Group
    colormap_group = QtWidgets.QGroupBox(CANVAS_LABEL_TEXT["colormap"])
    colormap_layout = QtWidgets.QVBoxLayout()
    colormap_layout.setSpacing(10)

    window.cmap_combo = QtWidgets.QComboBox()
    try:
        cmap_list = sorted(colormaps.keys())
    except Exception:
        cmap_list = ["viridis", "plasma", "inferno", "magma", "cividis"]
    for name in cmap_list:
        try:
            icon = _colormap_icon(name, width=96, height=14)
        except Exception:
            icon = QtGui.QIcon()
        window.cmap_combo.addItem(icon, name)
    colormap_layout.addWidget(window.cmap_combo)

    range_label = QtWidgets.QLabel(CANVAS_LABEL_TEXT["range"])
    range_label.setStyleSheet("")
    colormap_layout.addWidget(range_label)

    range_container = QtWidgets.QWidget()
    range_layout = QtWidgets.QHBoxLayout(range_container)
    range_layout.setContentsMargins(0, 0, 0, 0)
    range_layout.setSpacing(8)

    window.vmin_edit = QtWidgets.QLineEdit()
    window.vmin_edit.setPlaceholderText(CANVAS_PLACEHOLDER_TEXT["range_min"])
    range_layout.addWidget(window.vmin_edit)

    to_label = QtWidgets.QLabel(CANVAS_LABEL_TEXT["range_to"])
    to_label.setStyleSheet("")
    range_layout.addWidget(to_label)

    window.vmax_edit = QtWidgets.QLineEdit()
    window.vmax_edit.setPlaceholderText(CANVAS_PLACEHOLDER_TEXT["range_max"])
    range_layout.addWidget(window.vmax_edit)

    colormap_layout.addWidget(range_container)

    window.auto_range_btn = QtWidgets.QPushButton(CANVAS_BUTTON_TEXT["auto_range"])
    window.auto_range_btn.setToolTip(CANVAS_TOOLTIPS["auto_range"])
    window.copy_range_btn = QtWidgets.QPushButton(CANVAS_BUTTON_TEXT["copy_range"])
    window.copy_range_btn.setToolTip(CANVAS_TOOLTIPS["copy_range"])

    colormap_layout.addWidget(window.auto_range_btn)
    colormap_layout.addWidget(window.copy_range_btn)
    colormap_group.setLayout(colormap_layout)
    tile_layout.addWidget(colormap_group)

    # Statistics Group
    stats_group = QtWidgets.QGroupBox(CANVAS_LABEL_TEXT["statistics"])
    stats_layout = QtWidgets.QVBoxLayout()
    window.stats_label = QtWidgets.QLabel("-")
    window.stats_label.setWordWrap(True)
    window.stats_label.setStyleSheet("")
    stats_layout.addWidget(window.stats_label)
    stats_group.setLayout(stats_layout)
    data_layout.addWidget(stats_group)

    # Actions Group
    actions_group = QtWidgets.QGroupBox(CANVAS_LABEL_TEXT["actions"])
    actions_layout = QtWidgets.QVBoxLayout()
    actions_layout.setSpacing(8)

    window.duplicate_btn = QtWidgets.QPushButton(CANVAS_BUTTON_TEXT["duplicate"])
    window.duplicate_btn.setToolTip(CANVAS_TOOLTIPS["duplicate"])

    window.remove_btn = QtWidgets.QPushButton(CANVAS_BUTTON_TEXT["remove_item"])
    window.remove_btn.setToolTip(CANVAS_TOOLTIPS["remove_item"])
    window.remove_btn.setProperty("destructive", True)
    window.remove_btn.setStyleSheet("")

    actions_layout.addWidget(window.duplicate_btn)
    actions_layout.addWidget(window.remove_btn)
    actions_group.setLayout(actions_layout)
    data_layout.addWidget(actions_group)
    data_layout.addStretch(1)

    tabs.addTab(tile_tab, "Tile")
    tabs.addTab(data_tab, "Stats / Actions")
    layout.addWidget(tabs)

    layout.addStretch(1)

    # Connect signals
    window.colorbar_edit.editingFinished.connect(window._on_colorbar_changed)
    window.text_scale_slider.valueChanged.connect(window._on_text_scale_changed)
    window.font_color_auto_check.toggled.connect(window._on_font_color_auto_toggled)
    window.font_color_btn.clicked.connect(window._on_font_color_pick)
    window.scale_bar_check.toggled.connect(window._on_scale_bar_toggled)
    window.scale_bar_combo.currentTextChanged.connect(window._on_scale_bar_size_changed)
    window.cmap_combo.currentTextChanged.connect(window._on_cmap_changed)
    window.vmin_edit.editingFinished.connect(window._on_range_changed)
    window.vmax_edit.editingFinished.connect(window._on_range_changed)
    window.auto_range_btn.clicked.connect(window._on_auto_range)
    window.copy_range_btn.clicked.connect(window._on_copy_range)
    window.show_colorbar_check.toggled.connect(window._on_global_show_colorbar_toggled)
    window.duplicate_btn.clicked.connect(window._on_duplicate_item)
    window.remove_btn.clicked.connect(window._on_remove_item)

    scroll.setWidget(panel)
    window._set_inspector_enabled(False)
    return scroll


def apply_styles(window, dark: bool = False):
    """Apply a compact canvas stylesheet that follows the main viewer theme."""
    if dark:
        base = CANVAS_DIALOG_STYLE
        bg = "#1d2228"
        fg = "#e2e7ee"
        panel = "#242b33"
        border = "#3a434f"
        toolbar_bg = "#222931"
        group_bg = "#2a313a"
        label_fg = "#aeb8c4"
        btn_bg = "#2f3741"
        btn_fg = "#edf2f7"
        btn_hover = "#39424d"
        btn_press = "#454f5b"
        input_bg = "#1b2026"
        input_fg = "#edf2f7"
        input_border = "#4a5563"
        accent = "#334155"
        slider = "#4b5563"
        handle = "#cbd5e1"
        hint = "#94a3b8"
        view_bg = "#202224"
    else:
        base = CANVAS_DIALOG_STYLE_LIGHT
        bg = "#f6f3ee"
        fg = "#1f2428"
        panel = "#fffdf9"
        border = "#d7d0c7"
        toolbar_bg = "#ebe4da"
        group_bg = "#f8f4ee"
        label_fg = "#695f55"
        btn_bg = "#f8f4ee"
        btn_fg = "#1f2428"
        btn_hover = "#efe7dc"
        btn_press = "#e5dbcf"
        input_bg = "#fffdf9"
        input_fg = "#1f2428"
        input_border = "#cfc5ba"
        accent = "#efe7dc"
        slider = "#ddd4c8"
        handle = "#7a6f63"
        hint = "#7a7a7a"
        view_bg = "#f3efe8"
    window.setStyleSheet(
        base
        + f"""
QDialog {{
    background-color: {bg};
    color: {fg};
}}
QGraphicsView {{
    background: {view_bg};
    border: none;
}}
QWidget#inspectorPanel {{
    background-color: {panel};
    border-left: 1px solid {border};
}}
QScrollArea#canvasLeftRail {{
    background-color: {bg};
    border-right: 1px solid {border};
}}
QGroupBox {{
    border: 1px solid {border};
    border-radius: 8px;
    margin-top: 8px;
    background: {panel};
}}
QGroupBox::title {{
    subcontrol-origin: margin;
    left: 8px;
    padding: 0 4px 0 4px;
    background: transparent;
    color: {label_fg};
    font-weight: 600;
}}
QWidget#canvasToolbar {{
    background-color: {toolbar_bg};
    border: none;
    border-bottom: 1px solid {border};
}}
QWidget[toolbarGroup="true"] {{
    background-color: {group_bg};
    border: 1px solid {border};
    border-radius: 6px;
    padding: 4px;
}}
QLabel[toolbarLabel="true"] {{
    color: {label_fg};
    font-weight: 600;
    font-size: 10px;
}}
QWidget[canvasQuickSection="true"] {{
    background-color: {group_bg};
    border: 1px solid {border};
    border-radius: 8px;
}}
QLabel[canvasQuickLabel="true"] {{
    color: {label_fg};
    font-size: 9px;
    font-weight: 700;
    text-transform: uppercase;
}}
QLabel[canvasHeader="true"] {{
    color: {fg};
    font-weight: 700;
    font-size: 12px;
    padding: 6px 8px;
    background-color: {accent};
    border: 1px solid {border};
    border-radius: 7px;
}}
QPushButton {{
    background-color: {btn_bg};
    color: {btn_fg};
    border: 1px solid {border};
    border-radius: 6px;
    padding: 3px 7px;
    min-height: 20px;
}}
QPushButton:hover {{
    background-color: {btn_hover};
}}
QPushButton:pressed {{
    background-color: {btn_press};
}}
QPushButton[destructive="true"] {{
    background-color: #c95146;
    color: #ffffff;
    border-color: #b33f36;
}}
QPushButton[destructive="true"]:hover {{
    background-color: #d35d52;
}}
QCheckBox {{
    color: {fg};
}}
QTabWidget::pane {{
    border: none;
    top: -1px;
}}
QTabBar::tab {{
    background: transparent;
    border: 1px solid transparent;
    padding: 4px 8px;
    margin-right: 4px;
    color: {label_fg};
}}
QTabBar::tab:selected {{
    background: #ffffff;
    border-color: {border};
    border-bottom-color: #ffffff;
    border-top-left-radius: 6px;
    border-top-right-radius: 6px;
    color: {fg};
}}
QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox {{
    background-color: {input_bg};
    color: {input_fg};
    border: 1px solid {input_border};
    border-radius: 6px;
    padding: 2px 4px;
}}
QSlider::groove:horizontal {{
    background: {slider};
    height: 6px;
    border-radius: 3px;
}}
QSlider::handle:horizontal {{
    background: {handle};
    border: 1px solid {border};
    width: 12px;
    margin: -4px 0;
    border-radius: 6px;
}}
"""
    )
    if hasattr(window, "selection_hint"):
        window.selection_hint.setStyleSheet(
            f"QLabel {{ color: {hint}; padding: 2px 4px 8px 4px; }}"
        )


def apply_status_style(window):
    if not hasattr(window, "status_label") or window.status_label is None:
        return
    dark = bool(getattr(window, "_dark", False))
    bg = "#26303a" if dark else "#efe7dc"
    fg = "#c7d2de" if dark else "#5d554d"
    border = "#3a434f" if dark else "#d7d0c7"
    window.status_label.setStyleSheet(
        "QLabel {"
        " padding: 5px 10px;"
        f" background-color: {bg};"
        f" color: {fg};"
        f" border-top: 1px solid {border};"
        "}"
    )



