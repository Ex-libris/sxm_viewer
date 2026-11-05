"""Toolbar helpers for the main SXM Viewer."""
from __future__ import annotations

from pathlib import Path

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtGui import QIcon
from PyQt5.QtSvg import QSvgRenderer
from PyQt5.QtWidgets import QToolBar

from .._shared import log_status
from .constants import TOOLBAR_CANVAS_MIN_HEIGHT, TOOLBAR_CANVAS_MIN_WIDTH
from .styles import MAIN_TOOLBAR_CANVAS_BUTTON_STYLE

_MOLECULE_ICON_PATH = Path(__file__).resolve().parent.parent / "Pentacene_acsv.svg"


def _load_molecule_pixmap(size: QtCore.QSize, color: QtGui.QColor | None = None):
    """Render molecule SVG directly to a pixmap at the requested size, recoloring if requested."""
    try:
        if not _MOLECULE_ICON_PATH.exists():
            log_status(f"Molecule icon not found: {_MOLECULE_ICON_PATH}")
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
            
    except Exception:
        return QtGui.QPixmap()


def figure_button_qss(push_button_qss: str) -> str:
    """The accent push-button style, for the Figure split tool button."""
    return (push_button_qss.replace("QPushButton", "QToolButton")
            + "QToolButton { padding-right: 22px; }"
            "QToolButton::menu-button { border: none; width: 18px; }")


def new_figure_from_selection(viewer):
    """A new custom figure of the thumbnail selection (else the previewed
    image). A selection that is exactly one detected series opens that."""
    fc = getattr(viewer, "figure_controller", None)
    if fc is None:
        return None
    keys = fc.selection_keys()
    if not keys:
        QtWidgets.QMessageBox.information(
            viewer, "New figure", "Select the images for the figure in the thumbnail grid "
            "(or preview one), then click New Figure.")
        return None
    return fc.open_selection(keys) if len(keys) >= 2 else fc.open_custom(keys)


def populate_figure_menu(viewer, menu):
    fc = getattr(viewer, "figure_controller", None)
    if fc is None:
        return
    keys = fc.selection_keys()
    act = menu.addAction(f"New figure from the selection ({len(keys)})" if keys else "New figure")
    act.setEnabled(bool(keys))
    act.triggered.connect(lambda _=False: new_figure_from_selection(viewer))
    recent = fc.recent_custom_figures()
    if recent:
        menu.addSection("Recent figures")
        for item in recent:
            n = len(item["keys"])
            text = f"{item['name']}  ({n} image{'s' if n != 1 else ''})"
            if item["missing"]:
                text += "  - folder not loaded"
            a = menu.addAction(text)
            a.setToolTip("\n".join(Path(k).name for k in item["keys"][:12]))
            a.triggered.connect(lambda _=False, fid=item["fid"]: fc.open_saved(fid))
    menu.addSeparator()
    menu.addAction("Open figure file…").triggered.connect(lambda _=False: open_figure_file(viewer))
    menu.addSeparator()
    legacy = menu.addAction("Canvas (legacy free layout)…")
    legacy.setToolTip("The old free-layout canvas; it will be retired - figures do grids, and "
                      "Inkscape (figure > Inkscape) takes any free layout")
    legacy.triggered.connect(lambda _=False: viewer._on_open_canvas())


def open_figure_file(viewer):
    from .controllers.figures import FIGURE_FILE_SUFFIX
    fc = getattr(viewer, "figure_controller", None)
    if fc is None:
        return None
    path, _ = QtWidgets.QFileDialog.getOpenFileName(
        viewer, "Open figure file", str(getattr(viewer, "last_dir", "") or ""),
        f"SXM figure (*{FIGURE_FILE_SUFFIX});;All files (*)")
    return fc.open_figure_file(path) if path else None


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
    viewer.toolbar_save_session_act.setToolTip("Save the current SXM viewer session (Ctrl+S)")
    viewer.toolbar_save_session_act.setShortcut(QtGui.QKeySequence("Ctrl+S"))
    viewer.toolbar_save_session_act.triggered.connect(viewer.on_save_session)
    viewer.toolbar_collection_btn = QtWidgets.QToolButton(toolbar)
    viewer.toolbar_collection_btn.setText("Collections")
    viewer.toolbar_collection_btn.setToolTip("Save selected preview/pop-up/crop results into curated cross-folder collections")
    viewer.toolbar_collection_btn.setPopupMode(QtWidgets.QToolButton.InstantPopup)
    viewer.toolbar_collection_menu = QtWidgets.QMenu(viewer.toolbar_collection_btn)
    viewer.toolbar_collection_current_path_act = viewer.toolbar_collection_menu.addAction("Current collection: none")
    viewer.toolbar_collection_current_path_act.setEnabled(False)
    viewer.toolbar_collection_menu.addAction("Create a Collection...", viewer.on_create_collection)
    viewer.toolbar_collection_menu.addAction("Open a Collection...", viewer.on_browse_collections)
    viewer.toolbar_collection_clear_target_act = viewer.toolbar_collection_menu.addAction("Clear Current Collection Target", viewer.on_clear_current_collection)
    viewer.toolbar_collection_menu.addSeparator()
    viewer.toolbar_recent_collections_menu = viewer.toolbar_collection_menu.addMenu("Recent Collections")
    viewer.toolbar_collection_menu.addAction("Show Collection Tray", viewer.on_show_collection_tray)
    viewer.toolbar_collection_add_selected_act = viewer.toolbar_collection_menu.addAction(
        "Add Selected Thumbnails...", viewer.on_add_selected_thumbnails_to_collection
    )
    viewer.toolbar_collection_add_selected_to_act = viewer.toolbar_collection_menu.addAction(
        "Add Selected Thumbnails to...", viewer.on_add_selected_thumbnails_to_collection_picker
    )
    viewer.toolbar_collection_menu.addAction("Add Current Preview...", viewer.on_add_current_preview_to_collection)
    viewer.toolbar_collection_menu.addAction("Add Active Pop-up...", viewer.on_add_active_popup_to_collection)
    viewer.toolbar_collection_menu.addAction("Add All Open Pop-ups...", viewer.on_add_all_popups_to_collection)
    viewer.toolbar_collection_menu.addAction("Add Selected Crop History...", viewer.on_add_selected_crops_to_collection)
    viewer.toolbar_collection_menu.addSeparator()
    viewer.toolbar_collection_menu.addAction("What Is a Collection?", viewer.on_collection_help)
    viewer.toolbar_collection_btn.setMenu(viewer.toolbar_collection_menu)
    viewer.toolbar_collection_menu.aboutToShow.connect(viewer._refresh_collection_toolbar_menu_labels)
    toolbar.addWidget(viewer.toolbar_collection_btn)
    try:
        viewer._refresh_collection_ui()
    except Exception:
        pass
    viewer.toolbar_popups_raise_act = QtWidgets.QAction("Pop-ups", viewer)
    viewer.toolbar_popups_raise_act.triggered.connect(viewer.on_recall_popouts)
    viewer.toolbar_popups_btn = QtWidgets.QToolButton(toolbar)
    viewer.toolbar_popups_btn.setDefaultAction(viewer.toolbar_popups_raise_act)
    viewer.toolbar_popups_btn.setPopupMode(QtWidgets.QToolButton.MenuButtonPopup)
    viewer.toolbar_popups_menu = QtWidgets.QMenu(viewer.toolbar_popups_btn)
    viewer.toolbar_popups_menu.aboutToShow.connect(viewer._rebuild_popup_menu)
    viewer.toolbar_popups_btn.setMenu(viewer.toolbar_popups_menu)
    viewer.toolbar_popups_btn.setEnabled(False)
    toolbar.addWidget(viewer.toolbar_popups_btn)
    toolbar.addSeparator()

    # Figure: the selected images (else the previewed one) as a publication
    # figure; the arrow lists recent figures, figure files and the legacy
    # free-layout canvas.
    viewer.toolbar_figure_btn = QtWidgets.QToolButton(toolbar)
    viewer.toolbar_figure_btn.setText("New Figure")
    viewer.toolbar_figure_btn.setToolTip(
        "Make a publication figure of the selected images (else the previewed image) - "
        "one row per channel; drag headers to reorder, right-click an image for its own "
        "filters / Zero / range. Arrow: recent figures, open a figure file")
    viewer.toolbar_figure_btn.setMinimumHeight(TOOLBAR_CANVAS_MIN_HEIGHT)
    viewer.toolbar_figure_btn.setMinimumWidth(TOOLBAR_CANVAS_MIN_WIDTH)
    viewer.toolbar_figure_btn.setPopupMode(QtWidgets.QToolButton.MenuButtonPopup)
    viewer.toolbar_figure_btn.setStyleSheet(figure_button_qss(MAIN_TOOLBAR_CANVAS_BUTTON_STYLE))
    viewer.toolbar_figure_btn.clicked.connect(lambda _=False: new_figure_from_selection(viewer))
    figure_menu = QtWidgets.QMenu(viewer.toolbar_figure_btn)
    figure_menu.aboutToShow.connect(lambda m=figure_menu: (m.clear(), populate_figure_menu(viewer, m)))
    viewer.toolbar_figure_btn.setMenu(figure_menu)
    toolbar.addWidget(viewer.toolbar_figure_btn)
    toolbar.addSeparator()

    viewer.toolbar_export_png_act = toolbar.addAction(_icon("image-x-generic"), "Export PNGs")
    viewer.toolbar_export_png_act.setToolTip(
        "Export the selected thumbnails (or the previewed image) as PNGs, "
        "rendered as displayed - current display preset included")
    viewer.toolbar_export_png_act.triggered.connect(lambda _=False: viewer.on_export_pngs())

    viewer.toolbar_obsidian_btn = QtWidgets.QToolButton(toolbar)
    viewer.toolbar_obsidian_btn.setText("Obsidian")
    viewer.toolbar_obsidian_btn.setToolTip(
        "Append the selected images (as displayed) to an Obsidian note - "
        "the active note of the open vault, or one you choose")
    viewer.toolbar_obsidian_btn.setPopupMode(QtWidgets.QToolButton.InstantPopup)
    viewer.toolbar_obsidian_menu = QtWidgets.QMenu(viewer.toolbar_obsidian_btn)
    viewer.toolbar_obsidian_menu.setToolTipsVisible(True)
    viewer.toolbar_obsidian_menu.aboutToShow.connect(
        lambda: viewer.populate_obsidian_menu(viewer.toolbar_obsidian_menu))
    viewer.toolbar_obsidian_btn.setMenu(viewer.toolbar_obsidian_menu)
    toolbar.addWidget(viewer.toolbar_obsidian_btn)

    viewer.toolbar_export_xyz_act = toolbar.addAction(_icon("document-save"), "Export XYZ")
    viewer.toolbar_export_xyz_act.triggered.connect(lambda _=False: viewer.on_export_xyz_files())

    viewer.toolbar_report_act = toolbar.addAction(_icon("x-office-document"), "Report")
    viewer.toolbar_report_act.setToolTip(
        "Generate a PDF report of the loaded folder: session overview map, images "
        "with spectroscopy positions and curves, grid maps with curve clustering, "
        "and a needs-attention appendix")
    viewer.toolbar_report_act.triggered.connect(
        lambda: viewer.report_controller.generate_folder_report())

    toolbar.addSeparator()
    viewer.toolbar_image_btn = QtWidgets.QToolButton(toolbar)
    viewer.toolbar_image_btn.setText("Image")
    viewer.toolbar_image_btn.setToolTip("Histogram, contrast, and crop/rotate actions for the current preview")
    viewer.toolbar_image_btn.setPopupMode(QtWidgets.QToolButton.InstantPopup)
    viewer.toolbar_image_menu = QtWidgets.QMenu(viewer.toolbar_image_btn)
    viewer.toolbar_histogram_act = viewer.toolbar_image_menu.addAction("Histogram...")
    viewer.toolbar_histogram_act.setToolTip("Show histogram and adjust display range")
    viewer.toolbar_histogram_act.triggered.connect(lambda: viewer._open_histogram_dialog(viewer.preview_canvas))
    viewer.toolbar_histogram_auto_act = viewer.toolbar_image_menu.addAction("Auto contrast (1-99%)")
    viewer.toolbar_histogram_auto_act.triggered.connect(lambda: viewer._auto_contrast(viewer.preview_canvas))
    viewer.toolbar_histogram_reset_act = viewer.toolbar_image_menu.addAction("Reset range")
    viewer.toolbar_histogram_reset_act.triggered.connect(lambda: viewer._reset_contrast(viewer.preview_canvas))
    viewer.toolbar_image_menu.addSeparator()
    viewer.toolbar_crop_rotate_act = viewer.toolbar_image_menu.addAction("Crop/Rotate...")
    viewer.toolbar_crop_rotate_act.setToolTip(
        "Crop/rotate/flip into a new virtual copy (the original is never "
        "altered); clip, gamma, and colormap stay live on this image")
    viewer.toolbar_crop_rotate_act.triggered.connect(viewer.on_adjust_image)
    viewer.toolbar_reset_adjust_act = viewer.toolbar_image_menu.addAction("Reset display adjustments")
    viewer.toolbar_reset_adjust_act.setToolTip(
        "Clear the previewed image's clip/gamma display adjustments (Ctrl+Z restores)")
    viewer.toolbar_reset_adjust_act.triggered.connect(viewer.on_reset_image_adjustments)
    viewer.toolbar_image_menu.addSeparator()
    viewer.toolbar_rotate_cw_act = viewer.toolbar_image_menu.addAction("Rotate 90° clockwise  (R)")
    viewer.toolbar_rotate_cw_act.triggered.connect(lambda: viewer.rotate_images(1))
    viewer.toolbar_rotate_ccw_act = viewer.toolbar_image_menu.addAction("Rotate 90° anticlockwise  (Shift+R)")
    viewer.toolbar_rotate_ccw_act.triggered.connect(lambda: viewer.rotate_images(-1))
    viewer.toolbar_reset_orient_act = viewer.toolbar_image_menu.addAction("Reset orientation")
    viewer.toolbar_reset_orient_act.setToolTip(
        "Show the selected/previewed image in its acquired orientation again (Ctrl+Z restores)")
    viewer.toolbar_reset_orient_act.triggered.connect(lambda: viewer.reset_image_orientation())

    def _sync_image_menu():
        act = getattr(viewer, "toolbar_reset_adjust_act", None)
        if act is None:
            return
        has_spec = False
        try:
            if viewer.last_preview:
                has_spec = bool(viewer._get_adjust_spec(*viewer.last_preview))
        except Exception:
            has_spec = False
        act.setEnabled(has_spec)
        try:
            viewer.toolbar_reset_orient_act.setEnabled(
                any(viewer._get_image_orientation(p) for p in viewer._orientation_targets()))
        except Exception:
            pass

    viewer.toolbar_image_menu.aboutToShow.connect(_sync_image_menu)
    viewer.toolbar_image_btn.setMenu(viewer.toolbar_image_menu)
    toolbar.addWidget(viewer.toolbar_image_btn)

    try:
        from . import main_window_layout

        viewer.toolbar_display_btn.setText("Display")
        viewer.toolbar_display_btn.setToolTip("Preview and overlay display options")
        viewer.toolbar_display_btn.setMenu(main_window_layout._ensure_display_menu(viewer))
        toolbar.addWidget(viewer.toolbar_display_btn)

        viewer.toolbar_tools_btn = QtWidgets.QToolButton(toolbar)
        viewer.toolbar_tools_btn.setText("Tools")
        viewer.toolbar_tools_btn.setToolTip("Preview tools, docking, and recovery options")
        viewer.toolbar_tools_btn.setPopupMode(QtWidgets.QToolButton.InstantPopup)
        viewer.toolbar_tools_btn.setMenu(main_window_layout._ensure_tools_menu(viewer))

        def _sync_tools_menu():
            detached = bool(getattr(viewer, "preview_detached", False))
            locked = bool(getattr(viewer, "preview_locked", False))
            detach_act = getattr(viewer, "tools_preview_detach_act", None)
            lock_act = getattr(viewer, "tools_preview_lock_act", None)
            if detach_act is not None:
                detach_act.setText("Dock preview" if detached else "Float preview")
                detach_act.setToolTip(
                    "Dock the floating preview back into the main window"
                    if detached
                    else "Detach the preview pane into its own floating window"
                )
                detach_act.setEnabled(not locked)
            if lock_act is not None:
                lock_act.blockSignals(True)
                lock_act.setChecked(locked)
                lock_act.blockSignals(False)

        viewer.toolbar_tools_menu = viewer.toolbar_tools_btn.menu()
        viewer.toolbar_tools_menu.aboutToShow.connect(_sync_tools_menu)
        _sync_tools_menu()
        toolbar.addWidget(viewer.toolbar_tools_btn)

        if getattr(viewer, "toolbar_dark_btn", None) is not None:
            toolbar.addWidget(viewer.toolbar_dark_btn)
    except Exception:
        pass

    viewer.toolbar_spectro_browser_act = QtWidgets.QAction(_icon("view-list"), "Spectroscopy", viewer)
    viewer.toolbar_spectro_browser_act.setToolTip("Open the spectroscopy browser. Use the dropdown for spectroscopy display controls.")
    viewer.toolbar_spectro_browser_act.triggered.connect(lambda: viewer.open_spectro_browser())
    viewer.toolbar_spectro_btn = QtWidgets.QToolButton(toolbar)
    viewer.toolbar_spectro_btn.setDefaultAction(viewer.toolbar_spectro_browser_act)
    viewer.toolbar_spectro_btn.setPopupMode(QtWidgets.QToolButton.MenuButtonPopup)
    viewer.toolbar_spectro_menu = QtWidgets.QMenu(viewer.toolbar_spectro_btn)
    viewer.toolbar_spectro_menu.setToolTipsVisible(True)
    viewer.toolbar_spectro_menu.addAction("Open browser", lambda: viewer.open_spectro_browser())
    viewer.toolbar_spectro_menu.addAction("Open current position", viewer.on_open_current_spectro_site)
    viewer.toolbar_spectro_review_low_conf_act = viewer.toolbar_spectro_menu.addAction("Review low confidence")
    viewer.toolbar_spectro_review_low_conf_act.setToolTip("Open the spectroscopy browser filtered to low-confidence image assignments")
    viewer.toolbar_spectro_review_low_conf_act.triggered.connect(viewer.on_review_low_conf_spectros)
    viewer.toolbar_spectro_review_current_image_act = viewer.toolbar_spectro_menu.addAction("Review low confidence (current image)")
    viewer.toolbar_spectro_review_current_image_act.setToolTip("Review low-confidence spectroscopy assignments that affect the current preview image")
    viewer.toolbar_spectro_review_current_image_act.triggered.connect(lambda: viewer.on_review_low_conf_spectros(current_image_only=True))
    viewer.toolbar_spectro_review_off_frame_act = viewer.toolbar_spectro_menu.addAction("Review off-frame spectroscopies")
    viewer.toolbar_spectro_review_off_frame_act.setToolTip("Open the spectroscopy browser filtered to spectra whose real position falls outside every image")
    viewer.toolbar_spectro_review_off_frame_act.triggered.connect(viewer.on_review_off_frame_spectros)
    viewer.toolbar_spectro_menu.addSeparator()
    viewer.toolbar_spectro_repeat_share_act = viewer.toolbar_spectro_menu.addAction("Share across repeat scans")
    viewer.toolbar_spectro_repeat_share_act.setCheckable(True)
    viewer.toolbar_spectro_repeat_share_act.setChecked(getattr(viewer, "spectro_share_overlapping_repeats", False))
    repeat_tip = (
        "When enabled, spectros whose coordinates fall inside several near-identical overlapping image scans "
        "are shown on all of those repeat scans instead of being split to only the nearest image in time."
    )
    viewer.toolbar_spectro_repeat_share_act.setToolTip(repeat_tip)
    viewer.toolbar_spectro_repeat_share_act.setStatusTip(repeat_tip)
    viewer.toolbar_spectro_repeat_share_act.setWhatsThis(repeat_tip)
    viewer.toolbar_spectro_repeat_share_act.toggled.connect(viewer.on_spectro_share_overlapping_repeats_toggled)
    viewer.toolbar_spectro_markers_act = viewer.toolbar_spectro_menu.addAction("Thumbnail markers")
    viewer.toolbar_spectro_markers_act.setCheckable(True)
    viewer.toolbar_spectro_markers_act.setChecked(getattr(viewer, "show_spectra", True))
    viewer.toolbar_spectro_markers_act.setToolTip("Show clickable spectroscopy point markers on image thumbnails")
    viewer.toolbar_spectro_markers_act.toggled.connect(viewer.on_show_spectra_toggled)
    viewer.toolbar_spectro_preview_act = viewer.toolbar_spectro_menu.addAction("Preview markers")
    viewer.toolbar_spectro_preview_act.setCheckable(True)
    viewer.toolbar_spectro_preview_act.setChecked(getattr(viewer, "show_preview_spectra", getattr(viewer, "show_spectra", True)))
    viewer.toolbar_spectro_preview_act.setToolTip("Show spectroscopy point markers on the main preview")
    viewer.toolbar_spectro_preview_act.toggled.connect(viewer.on_show_preview_spectra_toggled)
    viewer.toolbar_spectro_miniatures_act = viewer.toolbar_spectro_menu.addAction("Thumbnail miniatures")
    viewer.toolbar_spectro_miniatures_act.setCheckable(True)
    viewer.toolbar_spectro_miniatures_act.setChecked(getattr(viewer, "show_spectro_miniatures", False))
    viewer.toolbar_spectro_miniatures_act.setToolTip("Show spectroscopy miniatures as separate thumbnail cards in the main grid")
    viewer.toolbar_spectro_miniatures_act.toggled.connect(viewer.on_show_spectro_miniatures_toggled)
    viewer.toolbar_spectro_menu.addSeparator()
    viewer.toolbar_spectro_matrix_markers_act = viewer.toolbar_spectro_menu.addAction("Matrix markers")
    viewer.toolbar_spectro_matrix_markers_act.setCheckable(True)
    viewer.toolbar_spectro_matrix_markers_act.setChecked(getattr(viewer, "show_matrix_markers", True))
    viewer.toolbar_spectro_matrix_markers_act.toggled.connect(viewer.on_show_matrix_markers_toggled)
    viewer.toolbar_spectro_single_markers_act = viewer.toolbar_spectro_menu.addAction("Single markers")
    viewer.toolbar_spectro_single_markers_act.setCheckable(True)
    viewer.toolbar_spectro_single_markers_act.setChecked(getattr(viewer, "show_single_markers", True))
    viewer.toolbar_spectro_single_markers_act.toggled.connect(viewer.on_show_single_markers_toggled)
    viewer.toolbar_spectro_compact_markers_act = viewer.toolbar_spectro_menu.addAction("Compact markers")
    viewer.toolbar_spectro_compact_markers_act.setCheckable(True)
    viewer.toolbar_spectro_compact_markers_act.setChecked(getattr(viewer, "compact_markers", True))
    viewer.toolbar_spectro_compact_markers_act.toggled.connect(viewer.on_compact_markers_toggled)
    viewer.toolbar_spectro_menu.addSeparator()
    viewer.toolbar_spectro_menu.addAction("Clear selection", viewer.on_clear_spec_selection)
    viewer.toolbar_spectro_highlight_act = viewer.toolbar_spectro_menu.addAction("Spectro highlight glow")
    viewer.toolbar_spectro_highlight_act.setCheckable(True)
    viewer.toolbar_spectro_highlight_act.setChecked(getattr(viewer, "spectro_highlight_glow", True))
    viewer.toolbar_spectro_highlight_act.toggled.connect(viewer.on_toggle_highlight_glow)
    marker_menu = viewer.toolbar_spectro_menu.addMenu("Marker style")
    if hasattr(viewer, "_populate_marker_style_menu"):
        viewer._populate_marker_style_menu(marker_menu)
    viewer.toolbar_spectro_marker_legend_act = viewer.toolbar_spectro_menu.addAction("What do these markers mean?")
    viewer.toolbar_spectro_marker_legend_act.setToolTip("Show a legend of the spectroscopy marker symbols, rings, and badges")
    viewer.toolbar_spectro_marker_legend_act.triggered.connect(viewer._show_spectro_marker_legend)
    viewer.toolbar_spectro_menu.addSeparator()
    viewer.toolbar_spectro_grid_as_matrix_act = viewer.toolbar_spectro_menu.addAction("NxN singles as matrix")
    viewer.toolbar_spectro_grid_as_matrix_act.setCheckable(True)
    viewer.toolbar_spectro_grid_as_matrix_act.setChecked(getattr(viewer, "spectro_single_grid_as_matrix", False))
    viewer.toolbar_spectro_grid_as_matrix_act.toggled.connect(viewer.on_spectro_grid_as_matrix_toggled)
    viewer.toolbar_spectro_force_single_act = viewer.toolbar_spectro_menu.addAction("Force single mode")
    viewer.toolbar_spectro_force_single_act.setCheckable(True)
    viewer.toolbar_spectro_force_single_act.setChecked(getattr(viewer, "spectro_force_single_mode", False))
    viewer.toolbar_spectro_force_single_act.toggled.connect(viewer.on_spectro_force_single_toggled)
    viewer.toolbar_spectro_btn.setMenu(viewer.toolbar_spectro_menu)
    toolbar.addWidget(viewer.toolbar_spectro_btn)
    toolbar.addSeparator()
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
    # The folder report describes the whole loaded folder, not the previewed
    # image, so it must be clickable as soon as a folder is loaded - gate it
    # on whether any files are present rather than on a preview/thumbnail
    # selection (the `enabled` flag, which only goes True once an image is
    # previewed).
    report_act = getattr(viewer, "toolbar_report_act", None)
    if report_act is not None:
        report_act.setEnabled(bool(getattr(viewer, "files", None)))
    btn = getattr(viewer, "preview_adjust_btn", None)
    if btn is not None:
        btn.setEnabled(bool(enabled))
    for widget in (
        getattr(viewer, "toolbar_image_btn", None),
    ):
        if widget is not None:
            widget.setEnabled(bool(enabled))
