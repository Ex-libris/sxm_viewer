"""Toolbar helpers for the main SXM Viewer."""
from __future__ import annotations

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QToolBar, QPushButton

from .constants import TOOLBAR_CANVAS_MIN_HEIGHT, TOOLBAR_CANVAS_MIN_WIDTH, UI_FONT_FAMILY
from .styles import MAIN_TOOLBAR_CANVAS_BUTTON_STYLE


def create_main_toolbar(viewer):
    try:
        toolbar = QToolBar("Main toolbar", viewer)
    except Exception:
        return None
    toolbar.setIconSize(QtCore.QSize(20, 20))

    def _icon(name):
        icon = QIcon.fromTheme(name)
        return icon if icon and not icon.isNull() else QIcon()

    viewer.toolbar_open_act = toolbar.addAction(_icon("folder-open"), "Open folder")
    viewer.toolbar_open_act.triggered.connect(viewer.open_folder_dialog)
    toolbar.addSeparator()

    viewer.toolbar_canvas_btn = QPushButton("Open Canvas")
    viewer.toolbar_canvas_btn.setToolTip("Open the publication canvas for layout/export")
    viewer.toolbar_canvas_btn.setMinimumHeight(TOOLBAR_CANVAS_MIN_HEIGHT)
    viewer.toolbar_canvas_btn.setMinimumWidth(TOOLBAR_CANVAS_MIN_WIDTH)
    viewer.toolbar_canvas_btn.setStyleSheet(MAIN_TOOLBAR_CANVAS_BUTTON_STYLE)
    viewer.toolbar_canvas_btn.clicked.connect(viewer._on_open_canvas)
    toolbar.addWidget(viewer.toolbar_canvas_btn)
    toolbar.addSeparator()

    viewer.toolbar_export_png_act = toolbar.addAction(_icon("image-x-generic"), "Export PNGs")
    viewer.toolbar_export_png_act.triggered.connect(viewer.on_export_pngs)

    viewer.toolbar_export_xyz_act = toolbar.addAction(_icon("document-save"), "Export XYZ")
    viewer.toolbar_export_xyz_act.triggered.connect(viewer.on_export_xyz_files)

    toolbar.addSeparator()
    viewer.toolbar_adjust_act = toolbar.addAction(_icon("transform-crop"), "Adjust image")
    viewer.toolbar_adjust_act.triggered.connect(viewer.on_adjust_image)
    viewer.toolbar_spectro_browser_act = toolbar.addAction(_icon("view-list"), "Spectro browser")
    viewer.toolbar_spectro_browser_act.triggered.connect(lambda: viewer.open_spectro_browser())
    viewer.toolbar_shortcuts_act = toolbar.addAction(_icon("help-about"), "Shortcuts")
    viewer.toolbar_shortcuts_act.triggered.connect(viewer._on_show_shortcuts_requested)

    toolbar.addSeparator()
    viewer.toolbar_layout_act = toolbar.addAction("Layout: Columns")
    viewer.toolbar_layout_act.setToolTip("Toggle between Columns and Stack layouts")
    viewer.toolbar_layout_act.triggered.connect(viewer._on_toggle_layout_mode)
    # Add a visible Dark Mode toggle button next to layout
    # Visible dark-mode toggle with ON/OFF text
    viewer.toolbar_dark_btn = QPushButton("dark mode: ON" if viewer.dark_mode else "dark mode: OFF")
    viewer.toolbar_dark_btn.setCheckable(True)
    viewer.toolbar_dark_btn.setChecked(viewer.dark_mode)
    viewer.toolbar_dark_btn.setToolTip("Toggle dark mode")
    viewer.toolbar_dark_btn.setMinimumWidth(100)
    try:
        viewer.toolbar_dark_btn.setStyleSheet(MAIN_TOOLBAR_CANVAS_BUTTON_STYLE)
    except Exception:
        pass
    # Use toggled so the checked state is passed through correctly
    viewer.toolbar_dark_btn.toggled.connect(viewer.on_dark_mode_toggled)
    toolbar.addWidget(viewer.toolbar_dark_btn)

    # Add small colormap controls next to dark-mode toggle for quick access
    try:
        # store labels on viewer so we can update their palette on dark/light toggles
        viewer.thumb_cmap_label = QtWidgets.QLabel("Thumb cmap:")
        viewer.thumb_cmap_label.setFont(QtGui.QFont(UI_FONT_FAMILY, 9))
        viewer.thumb_cmap_label.setStyleSheet("padding-left:8px; padding-right:4px;" + (" color: #e6e6e6;" if getattr(viewer, 'dark_mode', False) else ""))
        toolbar.addWidget(viewer.thumb_cmap_label)
        # Use the same combobox instance so selections stay in sync
        # apply dark styling to combos if viewer currently uses dark mode
        try:
            if getattr(viewer, 'dark_mode', False):
                combo_style = "QComboBox { background-color: #1f1f1f; border: 1px solid #444444; color: #f0f0f0; padding: 4px; }"
            else:
                combo_style = ""
            viewer.thumb_cmap_combo.setStyleSheet(combo_style)
            viewer.thumb_cmap_combo.setMinimumWidth(120)
            toolbar.addWidget(viewer.thumb_cmap_combo)

            viewer.preview_cmap_label = QtWidgets.QLabel("Preview cmap:")
            viewer.preview_cmap_label.setFont(QtGui.QFont(UI_FONT_FAMILY, 9))
            viewer.preview_cmap_label.setStyleSheet("padding-left:8px; padding-right:4px;" + (" color: #e6e6e6;" if getattr(viewer, 'dark_mode', False) else ""))
            toolbar.addWidget(viewer.preview_cmap_label)
            viewer.preview_cmap_combo.setStyleSheet(combo_style)
            viewer.preview_cmap_combo.setMinimumWidth(120)
            toolbar.addWidget(viewer.preview_cmap_combo)
        except Exception:
            # If combos aren't available at toolbar creation time, ignore and they'll appear where created
            pass
    except Exception:
        # outer try failed (toolbar creation at import time); ignore - combos will be laid out later
        pass

    update_toolbar_actions(viewer, False)
    return toolbar


def update_toolbar_actions(viewer, enabled: bool):
    for act in (viewer.toolbar_export_png_act, viewer.toolbar_export_xyz_act, viewer.toolbar_adjust_act):
        if act is not None:
            act.setEnabled(bool(enabled))



