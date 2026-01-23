"""UI construction helpers for the canvas window."""
from __future__ import annotations

from ..._shared import QtCore, QtGui, QtWidgets, colormaps
from ..styles import (
    CANVAS_DIALOG_STYLE,
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
    container_layout = QtWidgets.QVBoxLayout(container)
    container_layout.setContentsMargins(0, 0, 0, 0)
    container_layout.setSpacing(2)

    label = QtWidgets.QLabel(title)
    label.setStyleSheet(CANVAS_TOOLBAR_GROUP_LABEL_STYLE)
    label.setAlignment(QtCore.Qt.AlignLeft)
    container_layout.addWidget(label)

    group = QtWidgets.QWidget()
    group.setStyleSheet(CANVAS_TOOLBAR_GROUP_STYLE)
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
    section.setStyleSheet(CANVAS_SECTION_STYLE)
    layout = QtWidgets.QHBoxLayout(section)
    layout.setContentsMargins(8, 4, 8, 4)
    layout.setSpacing(4)

    if title:
        label = QtWidgets.QLabel(title)
        label.setStyleSheet(CANVAS_TOOLBAR_GROUP_LABEL_STYLE)
        layout.addWidget(label)

    for widget in widgets:
        layout.addWidget(widget)

    return section


def build_toolbar(window):
    """Build organized toolbar with clear visual grouping."""
    toolbar_widget = QtWidgets.QWidget()
    toolbar_widget.setStyleSheet(CANVAS_TOOLBAR_WIDGET_STYLE)
    toolbar_widget.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)

    main_layout = QtWidgets.QVBoxLayout(toolbar_widget)
    main_layout.setContentsMargins(8, 6, 8, 6)
    main_layout.setSpacing(6)

    row1 = QtWidgets.QHBoxLayout()
    row1.setSpacing(8)

    file_group, file_layout = create_toolbar_group(CANVAS_TOOLBAR_GROUP_TITLES["file"])
    window.save_btn = QtWidgets.QPushButton(CANVAS_BUTTON_TEXT["save"])
    window.save_btn.setToolTip(CANVAS_TOOLTIPS["save"])
    window.load_btn = QtWidgets.QPushButton(CANVAS_BUTTON_TEXT["load"])
    window.load_btn.setToolTip(CANVAS_TOOLTIPS["load"])
    window.export_btn = QtWidgets.QPushButton(CANVAS_BUTTON_TEXT["export"])
    window.export_btn.setToolTip(CANVAS_TOOLTIPS["export"])
    file_layout.addWidget(window.save_btn)
    file_layout.addWidget(window.load_btn)
    file_layout.addWidget(window.export_btn)
    row1.addWidget(file_group)

    row1.addWidget(create_separator())

    layout_group, layout_layout = create_toolbar_group(CANVAS_TOOLBAR_GROUP_TITLES["layout"])
    window.layout_2x2_btn = QtWidgets.QPushButton(CANVAS_BUTTON_TEXT["layout_2x2"])
    window.layout_2x2_btn.setToolTip(CANVAS_TOOLTIPS["layout_2x2"])
    window.layout_1x3_btn = QtWidgets.QPushButton(CANVAS_BUTTON_TEXT["layout_1x3"])
    window.layout_1x3_btn.setToolTip(CANVAS_TOOLTIPS["layout_1x3"])
    window.layout_3x1_btn = QtWidgets.QPushButton(CANVAS_BUTTON_TEXT["layout_3x1"])
    window.layout_3x1_btn.setToolTip(CANVAS_TOOLTIPS["layout_3x1"])
    layout_layout.addWidget(window.layout_2x2_btn)
    layout_layout.addWidget(window.layout_1x3_btn)
    layout_layout.addWidget(window.layout_3x1_btn)
    row1.addWidget(layout_group)

    row1.addWidget(create_separator())

    view_group, view_layout = create_toolbar_group(CANVAS_TOOLBAR_GROUP_TITLES["view"])
    window.show_grid_check = QtWidgets.QCheckBox(CANVAS_CHECKBOX_TEXT["grid"])
    window.show_grid_check.setToolTip(CANVAS_TOOLTIPS["grid"])
    window.snap_grid_check = QtWidgets.QCheckBox(CANVAS_CHECKBOX_TEXT["snap"])
    window.snap_grid_check.setToolTip(CANVAS_TOOLTIPS["snap"])
    window.canvas_color_btn = QtWidgets.QPushButton(CANVAS_BUTTON_TEXT["canvas_color"])
    window.canvas_color_btn.setToolTip(CANVAS_TOOLTIPS["canvas_color"])
    window.canvas_color_btn.setMaximumWidth(60)
    view_layout.addWidget(window.show_grid_check)
    view_layout.addWidget(window.snap_grid_check)
    view_layout.addWidget(window.canvas_color_btn)
    row1.addWidget(view_group)

    row1.addStretch(1)
    main_layout.addLayout(row1)

    row2 = QtWidgets.QHBoxLayout()
    row2.setSpacing(8)

    annotation_group, annotation_layout = create_toolbar_group(CANVAS_TOOLBAR_GROUP_TITLES["annotate"])
    window.show_colorbar_check = QtWidgets.QCheckBox(CANVAS_CHECKBOX_TEXT["colorbar"])
    window.show_colorbar_check.setToolTip(CANVAS_TOOLTIPS["colorbar"])
    window.show_colorbar_check.setChecked(window._global_show_colorbar)
    annotation_layout.addWidget(window.show_colorbar_check)
    row2.addWidget(annotation_group)

    row2.addWidget(create_separator())

    sync_group, sync_layout = create_toolbar_group(CANVAS_TOOLBAR_GROUP_TITLES["sync"])
    window.sync_cbar_check = QtWidgets.QCheckBox(CANVAS_CHECKBOX_TEXT["ranges"])
    window.sync_cbar_check.setToolTip(CANVAS_TOOLTIPS["ranges"])
    window.sync_by_channel_check = QtWidgets.QCheckBox(CANVAS_CHECKBOX_TEXT["colors_by_channel"])
    window.sync_by_channel_check.setChecked(window._sync_by_channel)
    window.sync_by_channel_check.setToolTip(CANVAS_TOOLTIPS["colors_by_channel"])
    sync_layout.addWidget(window.sync_cbar_check)
    sync_layout.addWidget(window.sync_by_channel_check)
    row2.addWidget(sync_group)

    row2.addWidget(create_separator())

    overlay_group, overlay_layout = create_toolbar_group(CANVAS_TOOLBAR_GROUP_TITLES["overlay"])
    window.overlay_info_check = QtWidgets.QCheckBox(CANVAS_CHECKBOX_TEXT["overlay_info"])
    window.overlay_info_check.setChecked(window._show_overlay_info)
    window.overlay_info_check.setToolTip(CANVAS_TOOLTIPS["overlay_info"])
    window.overlay_file_check = QtWidgets.QCheckBox(CANVAS_CHECKBOX_TEXT["overlay_file"])
    window.overlay_file_check.setChecked(window._show_overlay_file)
    window.overlay_file_check.setToolTip(CANVAS_TOOLTIPS["overlay_file"])
    overlay_layout.addWidget(window.overlay_info_check)
    overlay_layout.addWidget(window.overlay_file_check)
    row2.addWidget(overlay_group)

    row2.addWidget(create_separator())

    colorbar_group, colorbar_layout = create_toolbar_group(CANVAS_TOOLBAR_GROUP_TITLES["colorbar"])
    window.colorbar_ticks_check = QtWidgets.QCheckBox(CANVAS_CHECKBOX_TEXT["colorbar_ticks"])
    window.colorbar_ticks_check.setToolTip(CANVAS_TOOLTIPS["colorbar_ticks"])
    window.colorbar_ticks_check.setChecked(window._global_show_colorbar_ticks)
    window.colorbar_mode_combo = QtWidgets.QComboBox()
    for label in ("Bottom", "Top", "Right", "Left", "Inset", "Hidden"):
        window.colorbar_mode_combo.addItem(label)
    colorbar_layout.addWidget(window.colorbar_ticks_check)
    colorbar_layout.addWidget(QtWidgets.QLabel(CANVAS_LABEL_TEXT["position"]))
    colorbar_layout.addWidget(window.colorbar_mode_combo)
    row2.addWidget(colorbar_group)
    row2.addWidget(create_separator())

    align_group, align_layout = create_toolbar_group(CANVAS_TOOLBAR_GROUP_TITLES["align"])
    window.align_btn = QtWidgets.QPushButton(CANVAS_BUTTON_TEXT["align_selected"])
    window.align_btn.setToolTip(CANVAS_TOOLTIPS["align_selected"])
    window.align_channels_btn = QtWidgets.QPushButton(CANVAS_BUTTON_TEXT["align_by_channel"])
    window.align_channels_btn.setToolTip(CANVAS_TOOLTIPS["align_by_channel"])
    window.polish_btn = QtWidgets.QPushButton(CANVAS_BUTTON_TEXT["polish_layout"])
    window.polish_btn.setToolTip(CANVAS_TOOLTIPS["polish_layout"])
    window.reset_alignment_btn = QtWidgets.QPushButton(CANVAS_BUTTON_TEXT["reset_alignment"])
    window.reset_alignment_btn.setToolTip(CANVAS_TOOLTIPS["reset_alignment"])
    align_layout.addWidget(window.align_btn)
    align_layout.addWidget(window.align_channels_btn)
    align_layout.addWidget(window.polish_btn)
    align_layout.addWidget(window.reset_alignment_btn)
    row2.addWidget(align_group)

    row2.addStretch(1)
    main_layout.addLayout(row2)

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
    window.colorbar_ticks_check.toggled.connect(window._on_global_show_colorbar_ticks_toggled)
    window.colorbar_mode_combo.currentTextChanged.connect(window._on_colorbar_position_changed)
    window.canvas_color_btn.clicked.connect(window._on_canvas_color_clicked)
    window.align_btn.clicked.connect(window._on_align_selected)
    window.align_channels_btn.clicked.connect(window._on_align_by_channels)
    window.polish_btn.clicked.connect(window._on_polish_layout)
    window.reset_alignment_btn.clicked.connect(window._reset_locked_alignment)

    window._on_overlay_info_toggled(window.overlay_info_check.isChecked())
    window._on_overlay_file_toggled(window.overlay_file_check.isChecked())
    window.colorbar_mode_combo.setCurrentText(window._colorbar_mode.capitalize())

    return toolbar_widget


def build_inspector(window):
    """Build inspector panel with better visual hierarchy."""
    scroll = QtWidgets.QScrollArea()
    scroll.setWidgetResizable(True)
    scroll.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)

    panel = QtWidgets.QWidget()
    layout = QtWidgets.QVBoxLayout(panel)
    layout.setContentsMargins(12, 12, 12, 12)
    layout.setSpacing(12)

    # Header with higher contrast
    header = QtWidgets.QLabel(CANVAS_LABEL_TEXT["selected_item"])
    header.setStyleSheet(CANVAS_HEADER_STYLE)
    layout.addWidget(header)

    # Item Info Group
    info_group = QtWidgets.QGroupBox(CANVAS_LABEL_TEXT["item_info"])
    info_layout = QtWidgets.QFormLayout()
    info_layout.setLabelAlignment(QtCore.Qt.AlignRight)
    info_layout.setVerticalSpacing(10)
    info_layout.setHorizontalSpacing(12)

    label_style = CANVAS_LABEL_STYLE

    file_label_text = QtWidgets.QLabel(CANVAS_LABEL_TEXT["file"])
    file_label_text.setStyleSheet(label_style)
    window.file_label = QtWidgets.QLabel("-")
    window.file_label.setWordWrap(True)
    window.file_label.setStyleSheet(CANVAS_LABEL_VALUE_STYLE)
    info_layout.addRow(file_label_text, window.file_label)

    channel_label_text = QtWidgets.QLabel(CANVAS_LABEL_TEXT["channel"])
    channel_label_text.setStyleSheet(label_style)
    window.channel_label = QtWidgets.QLabel("-")
    window.channel_label.setStyleSheet(CANVAS_LABEL_VALUE_STYLE)
    info_layout.addRow(channel_label_text, window.channel_label)

    info_group.setLayout(info_layout)
    layout.addWidget(info_group)

    # Appearance Group
    appearance_group = QtWidgets.QGroupBox(CANVAS_LABEL_TEXT["appearance"])
    appearance_layout = QtWidgets.QFormLayout()
    appearance_layout.setLabelAlignment(QtCore.Qt.AlignRight)
    appearance_layout.setVerticalSpacing(10)
    appearance_layout.setHorizontalSpacing(12)

    colorbar_label = QtWidgets.QLabel(CANVAS_LABEL_TEXT["label"])
    colorbar_label.setStyleSheet(label_style)
    window.colorbar_edit = QtWidgets.QLineEdit()
    window.colorbar_edit.setPlaceholderText(CANVAS_PLACEHOLDER_TEXT["colorbar_label"])
    appearance_layout.addRow(colorbar_label, window.colorbar_edit)

    font_color_label = QtWidgets.QLabel(CANVAS_LABEL_TEXT["font_color"])
    font_color_label.setStyleSheet(label_style)
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
    text_scale_label.setStyleSheet(label_style)
    window.text_scale_slider = QtWidgets.QSlider(QtCore.Qt.Horizontal)
    window.text_scale_slider.setMinimum(1)
    window.text_scale_slider.setMaximum(240)
    window.text_scale_slider.setValue(int(round(window._global_text_scale * 100)))
    window.text_scale_slider.setEnabled(True)
    appearance_layout.addRow(text_scale_label, window.text_scale_slider)

    scale_bar_label = QtWidgets.QLabel(CANVAS_LABEL_TEXT["scale_bar"])
    scale_bar_label.setStyleSheet(label_style)
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
    layout.addWidget(appearance_group)

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
    range_label.setStyleSheet(CANVAS_RANGE_LABEL_STYLE)
    colormap_layout.addWidget(range_label)

    range_container = QtWidgets.QWidget()
    range_layout = QtWidgets.QHBoxLayout(range_container)
    range_layout.setContentsMargins(0, 0, 0, 0)
    range_layout.setSpacing(8)

    window.vmin_edit = QtWidgets.QLineEdit()
    window.vmin_edit.setPlaceholderText(CANVAS_PLACEHOLDER_TEXT["range_min"])
    range_layout.addWidget(window.vmin_edit)

    to_label = QtWidgets.QLabel(CANVAS_LABEL_TEXT["range_to"])
    to_label.setStyleSheet(CANVAS_RANGE_TO_LABEL_STYLE)
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
    layout.addWidget(colormap_group)

    # Statistics Group
    stats_group = QtWidgets.QGroupBox(CANVAS_LABEL_TEXT["statistics"])
    stats_layout = QtWidgets.QVBoxLayout()
    window.stats_label = QtWidgets.QLabel("-")
    window.stats_label.setWordWrap(True)
    window.stats_label.setStyleSheet(CANVAS_STATS_LABEL_STYLE)
    stats_layout.addWidget(window.stats_label)
    stats_group.setLayout(stats_layout)
    layout.addWidget(stats_group)

    # Actions Group
    actions_group = QtWidgets.QGroupBox(CANVAS_LABEL_TEXT["actions"])
    actions_layout = QtWidgets.QVBoxLayout()
    actions_layout.setSpacing(8)

    window.duplicate_btn = QtWidgets.QPushButton(CANVAS_BUTTON_TEXT["duplicate"])
    window.duplicate_btn.setToolTip(CANVAS_TOOLTIPS["duplicate"])

    window.remove_btn = QtWidgets.QPushButton(CANVAS_BUTTON_TEXT["remove_item"])
    window.remove_btn.setToolTip(CANVAS_TOOLTIPS["remove_item"])
    window.remove_btn.setStyleSheet(CANVAS_REMOVE_BUTTON_STYLE)

    actions_layout.addWidget(window.duplicate_btn)
    actions_layout.addWidget(window.remove_btn)
    actions_group.setLayout(actions_layout)
    layout.addWidget(actions_group)

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
    """Apply scientific GUI styling - high contrast, clear organization."""
    if dark:
        bg = "#1b1d23"
        fg = "#f0f0f0"
        panel = "#21242b"
        border = "#2e333d"
    else:
        bg = "#f7f7f7"
        fg = "#1b1b1b"
        panel = "#ffffff"
        border = "#c4c4c4"
    window.setStyleSheet(
        CANVAS_DIALOG_STYLE
        + f"""
QDialog {{
    background-color: {bg};
    color: {fg};
}}
QWidget#inspectorPanel {{
    background-color: {panel};
    border-left: 1px solid {border};
}}
QGroupBox {{
    border: 1px solid {border};
    margin-top: 6px;
}}
QGroupBox::title {{
    subcontrol-origin: margin;
    left: 8px;
    padding: 0 4px 0 4px;
    background: transparent;
}}
"""
    )


def apply_status_style(window):
    window.status_label.setStyleSheet(CANVAS_STATUS_LABEL_STYLE)



