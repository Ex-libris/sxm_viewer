"""Thumbnail right-click menu.

Module-function convention (see ``gui/viewer/*.py``): ``SXMGridViewer.
_on_thumb_context_menu`` is a one-line shim into ``show_thumb_context_menu``.
Built from the shared ``gui/menu_kit.py`` pieces, like the preview/pop-up
image menu, so both look and behave alike:

- greyed header (the file name, or "N images selected") + Star;
- **Display**: Colormap / Orientation / Zero / Filters / CH height label,
  Adjust image, Reset display adjustments (only when there are any);
- **Export**: menu_kit export section over ``export_routes.thumbnail_routes``
  (Again / pinned / More export), scope ``"thumbnails"``;
- **Analyse & create**: compare spectra, virtual copy, drift correction,
  animation, molecules (only when something applies);
- **Organise**: collections, source file, clear spectroscopy selection
  (only when there is one).

The folder report is not here - it covers the whole folder and lives on
the toolbar.
"""
from __future__ import annotations

from pathlib import Path

from ..._shared import QtWidgets, log_status
from ... import cmap_registry
from ...processing.filters import FILTER_DEFINITIONS, _gaussian_available
from .. import drift_animation
from ..menu_kit import add_export_section, section
from ..system_open import add_source_file_menu
from ..thumbnail_render import _colormap_icon
from . import ch_height as viewer_ch_height
from . import export_routes


def _ordered(viewer, paths):
    order = {str(fp): i for i, fp in enumerate(getattr(viewer, "current_thumb_files", []) or [])}
    uniq = list(dict.fromkeys(str(p) for p in paths if p))
    return sorted(uniq, key=lambda k: (order.get(k, len(order)), k))


def _add_filters_menu(viewer, menu, fp, targets):
    steps = viewer._thumbnail_filter_steps(fp)
    summary = viewer._filter_pipeline_label_from_steps(steps)
    sub = menu.addMenu("Filters" + (f"  ({summary})" if summary else ""))
    for key, info in FILTER_DEFINITIONS.items():
        if info.get('requires_dialog'):
            # Needs a per-image review dialog (e.g. periodic-noise removal) -
            # not meaningful as a blind batch action.
            continue
        prefix = "Add step: " if steps else ""
        act = sub.addAction(f"{prefix}{viewer._filter_action_label(key)}")
        if info.get('needs_gaussian') and not _gaussian_available():
            act.setEnabled(False)
            act.setToolTip("Requires scipy or OpenCV.")
        act.triggered.connect(lambda _, k=key, paths=list(targets), focus=fp:
                              viewer._apply_filter_to_paths(paths, k, focus_path=focus))
    custom = sub.addAction("Edit custom pipeline..." if steps else "Custom pipeline...")
    custom.triggered.connect(lambda _, paths=list(targets), focus=fp: viewer._open_custom_filter_dialog(paths, focus))
    sub.addSeparator()
    any_filtered = any(viewer._thumbnail_filter_steps(p) for p in targets)
    if len(targets) > 1:
        clear_sel = sub.addAction(f"Clear filters ({len(targets)} selected)")
        clear_sel.setEnabled(any_filtered)
        clear_sel.triggered.connect(lambda _, paths=list(targets): viewer._clear_filter_for_paths(paths))
    clear_one = sub.addAction("Clear filter (this image)" if len(targets) > 1 else "Clear filter")
    clear_one.setEnabled(bool(steps))
    clear_one.triggered.connect(lambda _, paths=[fp]: viewer._clear_filter_for_paths(paths))


def _add_cmap_menu(viewer, menu, label_widget, cmap_targets):
    cmap_menu = menu.addMenu("Colormap")
    try:
        ch_for_cmap = int(label_widget.property("channel_index"))
    except Exception:
        ch_for_cmap = None
    current_cmaps = {(viewer.per_file_channel_cmap or {}).get((p, ch_for_cmap)) for p in cmap_targets}
    current_cmap = next(iter(current_cmaps)) if len(current_cmaps) == 1 else None
    cmap_names = cmap_registry.all_cmap_names()
    featured = [n for n in ["viridis", "cividis", "Blues_r", "gray", "inferno", "magma", "plasma",
                            "coolwarm", "turbo"] if n in cmap_names]
    if current_cmap and current_cmap not in featured:
        featured.insert(0, current_cmap)
    for cmap_name in featured:
        act = cmap_menu.addAction(cmap_name)
        try:
            act.setIcon(_colormap_icon(cmap_name, width=96, height=14))
        except Exception:
            pass
        act.setCheckable(True)
        act.setChecked(cmap_name == current_cmap)
        act.triggered.connect(lambda _, paths=list(cmap_targets), name=cmap_name:
                              viewer._set_thumbnail_entry_cmap(paths, name))
    cmap_menu.addSeparator()
    gallery = cmap_menu.addAction("\U0001F3A8 Colormap gallery...")
    gallery.triggered.connect(lambda _, paths=list(cmap_targets): viewer._pick_thumbnail_cmap_from_gallery(paths))
    reset = cmap_menu.addAction("Use global thumbnail/preview cmap")
    reset.setEnabled(any(c for c in current_cmaps))
    reset.triggered.connect(lambda _, paths=list(cmap_targets): viewer._set_thumbnail_entry_cmap(paths, None))


def _add_orientation_menu(viewer, menu, targets):
    orient_menu = menu.addMenu("Orientation")
    for text, turns in (("Rotate 90° clockwise  (R)", 1),
                        ("Rotate 90° anticlockwise  (Shift+R)", -1),
                        ("Rotate 180°", 2)):
        act = orient_menu.addAction(text)
        act.triggered.connect(lambda _, paths=list(targets), k=turns: viewer.rotate_images(k, paths))
    orient_menu.addSeparator()
    reset = orient_menu.addAction("Reset to acquired orientation")
    reset.setEnabled(any(viewer._get_image_orientation(p) for p in targets))
    reset.triggered.connect(lambda _, paths=list(targets): viewer.reset_image_orientation(paths))


def _add_reset_adjustments(viewer, menu, targets):
    try:
        ch = int(viewer.channel_dropdown.currentIndex())
        has_adjust = any(bool(viewer._get_adjust_spec(p, ch)) for p in targets)
    except Exception:
        has_adjust = False
    if not has_adjust:
        return
    act = menu.addAction("Reset display adjustments")
    act.setToolTip("Clear the clip/gamma display adjustments from Image > Crop/Rotate (Ctrl+Z restores)")

    def _reset(_=False, paths=list(targets)):
        ch = int(viewer.channel_dropdown.currentIndex())
        for p in paths:
            old = viewer._get_adjust_spec(p, ch)
            if not old:
                continue
            viewer._adjustment_undo_stack.append(("tone", str(p), int(ch), dict(old)))
            viewer._set_adjust_spec(p, ch, None)
            viewer._refresh_adjusted_channel(p, ch)

    act.triggered.connect(_reset)


def _add_molecules_menu(viewer, menu, targets):
    overlays = getattr(viewer, "molecule_overlays", {}) or {}
    has_mol = any(overlays.get(str(Path(p))) for p in targets)
    has_copy_src = any(viewer._is_processed_key(str(p)) for p in targets)
    if not (has_mol or has_copy_src):
        return
    mol_menu = menu.addMenu("Molecules")
    clear = mol_menu.addAction("Clear molecules")
    clear.setEnabled(has_mol)
    clear.triggered.connect(lambda _, paths=list(targets): viewer._clear_molecules_for_paths(paths))
    copy = mol_menu.addAction("Copy molecules from source")
    copy.setEnabled(has_copy_src)
    copy.triggered.connect(lambda _, paths=list(targets): viewer._copy_molecules_from_source(paths))


def _add_collections_menu(viewer, menu, targets):
    collection_menu = menu.addMenu("Collections")
    show_tray = collection_menu.addAction("Show Collection Tray")
    show_tray.triggered.connect(viewer.on_show_collection_tray)
    n = len([p for p in targets if p])
    noun = "Thumbnails" if n != 1 else "Thumbnail"
    add_act = collection_menu.addAction(f"Add Selected {noun}...")
    add_act.setEnabled(n > 0)
    add_act.triggered.connect(viewer.on_add_selected_thumbnails_to_collection)
    if n and not getattr(viewer, "_collection_source", None):
        add_act.setToolTip("Choose or open a collection first, or use this action to create one when prompted.")
    add_to = collection_menu.addAction(f"Add Selected {noun} to...")
    add_to.setEnabled(n > 0)
    add_to.setToolTip("Route the selected thumbnail(s) to a collection you pick now - does not change your "
                      "current collection target.")
    add_to.triggered.connect(viewer.on_add_selected_thumbnails_to_collection_picker)
    remove = collection_menu.addAction(f"Remove Selected {noun} from Collection")
    remove.triggered.connect(
        lambda: viewer.collection_controller.remove_thumbnail_entries(
            [{"file_path": str(path), "channel_index": int(viewer.channel_dropdown.currentIndex() or 0)}
             for path in targets if path]))
    remove.setEnabled(bool(getattr(viewer, "_collection_source", None)) and n > 0)
    collection_menu.addSeparator()
    collection_menu.addAction("Create a Collection...").triggered.connect(viewer.on_create_collection)
    collection_menu.addAction("Open a Collection...").triggered.connect(viewer.on_browse_collections)
    collection_menu.addAction("Clear Current Collection Target").triggered.connect(viewer.on_clear_current_collection)


def _has_spectro_selection(viewer):
    return bool(getattr(viewer, "_multi_spec_selection", None)
                or getattr(viewer, "_multi_spectro_popups", None))


def build_thumb_context_menu(viewer, label_widget):
    fp = str(label_widget.property("file_path"))
    # A multi-selection plus the clicked thumbnail, in grid order.
    targets = _ordered(viewer, set(viewer.thumb_multi_select or []) | {fp})
    known = [p for p in targets if p in getattr(viewer, "thumb_widgets", {})]
    menu = QtWidgets.QMenu(viewer)
    menu.setToolTipsVisible(True)

    section(menu, Path(fp).name if len(targets) == 1 else f"{len(targets)} images selected", first=True)
    all_starred = all(str(p) in viewer.starred for p in targets)
    star = menu.addAction(("Remove star" if all_starred else "★ Star") + "  (S)")
    star.setToolTip("Mark as favourite - view only starred images with Ctrl+Alt+F or "
                    "Display → Show only → Starred")
    star.triggered.connect(lambda _, paths=list(targets): viewer.toggle_star_for_paths(paths))

    # ---- Display
    section(menu, "Display")
    if known:
        _add_cmap_menu(viewer, menu, label_widget, known)
        _add_orientation_menu(viewer, menu, known)
        viewer._populate_zero_menu(menu.addMenu("Zero"), known)
    _add_filters_menu(viewer, menu, fp, targets)
    if len(known) == 1 and viewer_ch_height.ch_abs_z_nm(viewer, known[0]) is not None:
        viewer_ch_height.populate_menu(viewer, menu.addMenu("CH height label"), known[0])
    menu.addAction("Adjust image...").triggered.connect(viewer.on_adjust_image)
    _add_reset_adjustments(viewer, menu, targets)

    # ---- Export
    try:
        add_export_section(menu, export_routes.thumbnail_routes(viewer),
                           export_routes.menu_prefs(viewer, "thumbnails"), list(targets), parent=viewer)
    except Exception as exc:
        log_status(f"[Thumb menu] export section failed: {exc}")

    # ---- Analyse & create
    section(menu, "Analyse & create")
    try:
        image_specs = viewer.spectro_compare_controller.all_specs_for_image(fp)
    except Exception:
        image_specs = []
    if len(image_specs) >= 2:
        act = menu.addAction(f"Compare all spectra on this image ({len(image_specs)})...")
        act.setToolTip("Open every point spectrum assigned to this image together in a comparison window")
        act.triggered.connect(lambda _, key=fp: viewer.spectro_compare_controller.open_all_specs_popup(key))
    virt = menu.addMenu("Virtual copy")
    virt.addAction("Current channel").triggered.connect(
        lambda _, paths=list(targets): viewer._create_virtual_channel_copies(paths, viewer.channel_dropdown.currentIndex()))
    virt.addAction("Choose channel...").triggered.connect(
        lambda _, paths=list(targets): viewer._create_virtual_channel_copies(paths, None))
    virt.addSeparator()
    remove_virt = virt.addAction("Remove virtual copies")
    remove_virt.setEnabled(any(viewer._is_processed_key(str(p)) for p in targets))
    remove_virt.triggered.connect(lambda _, paths=list(targets): viewer._remove_virtual_entries(paths))
    menu.addAction("Drift-correct and export...").triggered.connect(
        lambda _, paths=list(targets): drift_animation.on_drift_correct(viewer, paths))
    menu.addAction("Create animation...").triggered.connect(
        lambda _, paths=list(targets): drift_animation.on_create_animation(viewer, paths))
    _add_molecules_menu(viewer, menu, targets)

    # ---- Organise
    section(menu, "Organise")
    _add_collections_menu(viewer, menu, targets)
    add_source_file_menu(menu, fp, viewer)
    if _has_spectro_selection(viewer):
        menu.addAction("Clear spectroscopy selection").triggered.connect(viewer._clear_multi_spec_selection)
    return menu


def show_thumb_context_menu(viewer, label_widget, pos):
    menu = build_thumb_context_menu(viewer, label_widget)
    menu.exec_(label_widget.mapToGlobal(pos))
