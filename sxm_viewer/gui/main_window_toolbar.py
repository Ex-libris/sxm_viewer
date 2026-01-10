"""Toolbar helpers for the main SXM grid viewer."""
from __future__ import annotations

from PyQt5 import QtCore
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QToolBar, QPushButton

from .constants import TOOLBAR_CANVAS_MIN_HEIGHT, TOOLBAR_CANVAS_MIN_WIDTH
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

    update_toolbar_actions(viewer, False)
    return toolbar


def update_toolbar_actions(viewer, enabled: bool):
    for act in (viewer.toolbar_export_png_act, viewer.toolbar_export_xyz_act, viewer.toolbar_adjust_act):
        if act is not None:
            act.setEnabled(bool(enabled))



