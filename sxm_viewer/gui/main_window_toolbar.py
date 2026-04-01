"""Toolbar helpers for the main SXM Viewer."""
from __future__ import annotations

from pathlib import Path

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtGui import QIcon
from PyQt5.QtSvg import QSvgRenderer
from PyQt5.QtWidgets import QToolBar, QPushButton

from .constants import TOOLBAR_CANVAS_MIN_HEIGHT, TOOLBAR_CANVAS_MIN_WIDTH
from .styles import MAIN_TOOLBAR_CANVAS_BUTTON_STYLE

_MOLECULE_ICON_PATH = Path(__file__).resolve().parent.parent / "Pentacene_acsv.svg"


def _load_molecule_pixmap(size: QtCore.QSize, color: QtGui.QColor | None = None):
    """Render molecule SVG directly to a pixmap at the requested size, recoloring if requested."""
    try:
        if not _MOLECULE_ICON_PATH.exists():
            print(f"Molecule icon path does not exist: {_MOLECULE_ICON_PATH}")
            return QtGui.QPixmap()
        
        renderer = QSvgRenderer(str(_MOLECULE_ICON_PATH))
        if not renderer.isValid():
            return QtGui.QPixmap()

        image = QtGui.QImage(size, QtGui.QImage.Format_ARGB32)
        image.fill(QtCore.Qt.transparent)
        
        painter = QtGui.QPainter(image)
        painter.setRenderHint(QtGui.QPainter.Antialiasing, True)
        painter.setRenderHint(QtGui.QPainter.SmoothPixmapTransform, True)

        padding = 4
        avail_w = max(0, size.width() - 2 * padding)
        avail_h = max(0, size.height() - 2 * padding)
        default_size = renderer.defaultSize()
        base_w = default_size.width() or 1
        base_h = default_size.height() or 1
        scale = min(avail_w / base_w, avail_h / base_h)
        draw_w = base_w * scale
        draw_h = base_h * scale
        offset_x = padding + (avail_w - draw_w) / 2
        offset_y = padding + (avail_h - draw_h) / 2
        target_rect = QtCore.QRectF(offset_x, offset_y, draw_w, draw_h)
        renderer.render(painter, target_rect)
        painter.end()
        if color is not None:
            tint = QtGui.QColor(color)
            for y in range(image.height()):
                for x in range(image.width()):
                    alpha = QtGui.qAlpha(image.pixel(x, y))
                    if alpha:
                        tint.setAlpha(alpha)
                        image.setPixelColor(x, y, tint)
        return QtGui.QPixmap.fromImage(image)
            
    except Exception as e:
        return QtGui.QPixmap()


def create_main_toolbar(viewer):
    try:
        toolbar = QToolBar("Main toolbar", viewer)
    except Exception:
        return None
    toolbar.setMovable(False)
    toolbar.setFloatable(False)
    toolbar.setIconSize(QtCore.QSize(20, 20))

    def _icon(name):
        icon = QIcon.fromTheme(name)
        return icon if icon and not icon.isNull() else QIcon()

    viewer.toolbar_open_act = toolbar.addAction(_icon("folder-open"), "Open folder")
    viewer.toolbar_open_act.triggered.connect(viewer.open_folder_dialog)
    viewer.toolbar_load_session_act = QtWidgets.QAction(_icon("document-open"), "Load Session", viewer)
    viewer.toolbar_load_session_act.setToolTip("Restore a saved SXM viewer session")
    viewer.toolbar_load_session_act.triggered.connect(viewer.on_load_session)
    viewer.toolbar_load_session_btn = QtWidgets.QToolButton(toolbar)
    viewer.toolbar_load_session_btn.setDefaultAction(viewer.toolbar_load_session_act)
    viewer.toolbar_load_session_btn.setPopupMode(QtWidgets.QToolButton.MenuButtonPopup)
    viewer.toolbar_load_session_menu = QtWidgets.QMenu(viewer.toolbar_load_session_btn)
    viewer.toolbar_load_session_btn.setMenu(viewer.toolbar_load_session_menu)
    toolbar.addWidget(viewer.toolbar_load_session_btn)
    try:
        viewer._refresh_recent_session_dirs_menu()
    except Exception:
        pass
    viewer.toolbar_save_session_act = toolbar.addAction(_icon("document-save"), "Save Session")
    viewer.toolbar_save_session_act.setToolTip("Save the current SXM viewer session")
    viewer.toolbar_save_session_act.triggered.connect(viewer.on_save_session)
    viewer.toolbar_collection_btn = QtWidgets.QToolButton(toolbar)
    viewer.toolbar_collection_btn.setText("Collections")
    viewer.toolbar_collection_btn.setToolTip("Save selected preview/pop-up/crop results into curated cross-folder collections")
    viewer.toolbar_collection_btn.setPopupMode(QtWidgets.QToolButton.InstantPopup)
    viewer.toolbar_collection_menu = QtWidgets.QMenu(viewer.toolbar_collection_btn)
    viewer.toolbar_collection_menu.addAction("Open Collection...", viewer.on_open_collection)
    viewer.toolbar_collection_menu.addAction("Add Current Preview...", viewer.on_add_current_preview_to_collection)
    viewer.toolbar_collection_menu.addAction("Add Active Pop-up...", viewer.on_add_active_popup_to_collection)
    viewer.toolbar_collection_menu.addAction("Add All Open Pop-ups...", viewer.on_add_all_popups_to_collection)
    viewer.toolbar_collection_menu.addAction("Add Selected Crop History...", viewer.on_add_selected_crops_to_collection)
    viewer.toolbar_collection_menu.addSeparator()
    viewer.toolbar_collection_menu.addAction("What Is a Collection?", viewer.on_collection_help)
    viewer.toolbar_collection_btn.setMenu(viewer.toolbar_collection_menu)
    toolbar.addWidget(viewer.toolbar_collection_btn)
    viewer.toolbar_popups_btn = QtWidgets.QToolButton(toolbar)
    viewer.toolbar_popups_btn.setText("Pop-ups")
    viewer.toolbar_popups_btn.setToolTip("Restore deferred session pop-outs on demand")
    viewer.toolbar_popups_btn.setPopupMode(QtWidgets.QToolButton.InstantPopup)
    viewer.toolbar_popups_menu = QtWidgets.QMenu(viewer.toolbar_popups_btn)
    viewer.toolbar_popups_menu.aboutToShow.connect(viewer._rebuild_deferred_popup_menu)
    viewer.toolbar_popups_btn.setMenu(viewer.toolbar_popups_menu)
    viewer.toolbar_popups_btn.setEnabled(False)
    toolbar.addWidget(viewer.toolbar_popups_btn)
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
    for act in (viewer.toolbar_export_png_act, viewer.toolbar_export_xyz_act):
        if act is not None:
            act.setEnabled(bool(enabled))
    btn = getattr(viewer, "preview_adjust_btn", None)
    if btn is not None:
        btn.setEnabled(bool(enabled))
