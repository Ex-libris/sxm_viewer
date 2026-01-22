"""Toolbar helpers for the main SXM Viewer."""
from __future__ import annotations

from pathlib import Path

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtGui import QIcon, QImage
from PyQt5.QtSvg import QSvgRenderer
from PyQt5.QtWidgets import QToolBar, QPushButton, QToolButton

from .constants import TOOLBAR_CANVAS_MIN_HEIGHT, TOOLBAR_CANVAS_MIN_WIDTH, UI_FONT_FAMILY
from .styles import MAIN_TOOLBAR_CANVAS_BUTTON_STYLE
from .main_window_layout import _ensure_display_menu

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
    viewer.toolbar_display_btn = QToolButton()
    viewer.toolbar_display_btn.setText("Display Options")
    viewer.toolbar_display_btn.setPopupMode(QToolButton.InstantPopup)
    viewer.toolbar_display_btn.setToolTip("Display options and marker overlays")
    viewer.toolbar_display_btn.setMenu(_ensure_display_menu(viewer))
    toolbar.addWidget(viewer.toolbar_display_btn)
    
    # Create a clickable label with the molecule image instead of a button with icon
    viewer.toolbar_load_mol_btn = QtWidgets.QLabel()
    viewer.toolbar_load_mol_btn.setMinimumHeight(TOOLBAR_CANVAS_MIN_HEIGHT)
    viewer.toolbar_load_mol_btn.setMinimumWidth(TOOLBAR_CANVAS_MIN_WIDTH)
    viewer.toolbar_load_mol_btn.setAlignment(QtCore.Qt.AlignCenter)
    viewer.toolbar_load_mol_btn.setToolTip("Overlay a molecular structure (XYZ, PDB, MOL)")
    viewer.toolbar_load_mol_btn.setCursor(QtCore.Qt.PointingHandCursor)
    
    # Load and render the molecule directly to a pixmap at button size
    icon_size = QtCore.QSize(TOOLBAR_CANVAS_MIN_WIDTH - 10, TOOLBAR_CANVAS_MIN_HEIGHT - 6)
    viewer._molecule_pixmap_size = icon_size
    pixmap = _load_molecule_pixmap(icon_size, QtGui.QColor("#1d1d1d"))
    if not pixmap.isNull():
        viewer.toolbar_load_mol_btn.setPixmap(pixmap)
    else:
        viewer.toolbar_load_mol_btn.setText("Load Molecule")

    if hasattr(viewer, "_apply_molecule_button_theme"):
        viewer._apply_molecule_button_theme()

    # Make it clickable
    viewer.toolbar_load_mol_btn.mousePressEvent = lambda event: viewer.on_load_molecule()
    toolbar.addWidget(viewer.toolbar_load_mol_btn)

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
