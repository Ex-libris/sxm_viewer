"""Spectroscopy browser helpers for SXMGridViewer."""
from __future__ import annotations

from ..._shared import (
    QtCore,
    QtGui,
    QtWidgets,
    QIcon,
    QPixmap,
    QImage,
    QPainter,
    QPen,
    QBrush,
    FigureCanvas,
    Figure,
    Line2D,
    colormaps,
    np,
    Path,
    defaultdict,
    OrderedDict,
    datetime,
    hashlib,
    itertools,
    io,
    json,
    math,
    os,
    sys,
    threading,
    _scipy_ndimage,
    log_status,
    matplotlib,
)
from PyQt5.QtWidgets import QLabel, QTreeWidget, QTreeWidgetItem


def _spec_position_text(spec):
    try:
        if spec.get("x") is not None and spec.get("y") is not None:
            return f"{float(spec.get('x')):.1f}/{float(spec.get('y')):.1f}"
    except Exception:
        pass
    return ""


def _spec_leaf_label(spec, index=None):
    bits = []
    if index is not None:
        bits.append(f"{index}.")
    bits.append(Path(spec.get("path", "")).name)
    channel = str(spec.get("channel_name") or "").strip()
    if channel:
        bits.append(f"({channel})")
    pos = _spec_position_text(spec)
    if pos:
        bits.append(pos)
    stack = str(spec.get("xy_stack_display") or "").strip()
    if stack:
        bits.append(f"[{stack}]")
    return " ".join(bit for bit in bits if bit)


def _site_tree_label(site_specs):
    first = site_specs[0]
    display = str(first.get("site_display") or "Site").strip()
    trace_count = int(first.get("site_trace_count") or len(site_specs) or 0)
    channel_count = int(first.get("site_channel_count") or 0)
    low_conf_count = sum(
        1 for spec in list(site_specs or [])
        if str(spec.get("assignment_confidence") or "").strip().lower() == "low"
    )
    extras = [f"{trace_count} trace" + ("" if trace_count == 1 else "s")]
    if channel_count:
        extras.append(f"{channel_count} ch")
    if low_conf_count:
        extras.append(f"low conf {low_conf_count}")
    if first.get("site_has_z_stack"):
        extras.append("Z-stack")
    elif int(first.get("xy_stack_count") or 0) > 1:
        extras.append("same-XY")
    if first.get("site_has_matrix"):
        extras.append("matrix")
    return f"{display}  [{' | '.join(extras)}]"


def _image_tree_label(image_key, specs, site_count):
    name = Path(str(image_key or "")).name if image_key else "Unassigned spectra"
    return f"{name}  [{site_count} site" + ("" if site_count == 1 else "s") + f" | {len(specs)} spectra]"


def _spec_search_blob(spec):
    values = [
        Path(spec.get("path", "")).name,
        str(spec.get("channel_name") or ""),
        str(spec.get("site_display") or ""),
        str(spec.get("site_summary") or ""),
        str(spec.get("xy_stack_display") or ""),
        str(spec.get("xy_stack_summary") or ""),
        str(spec.get("assignment_summary") or ""),
        str(spec.get("assignment_reason_label") or spec.get("assignment_reason") or ""),
        _spec_position_text(spec),
    ]
    return " ".join(v for v in values if v).lower()


def _browser_filter_flags(viewer):
    return {
        "current_image_only": bool(getattr(getattr(viewer, "spectro_filter_current_image_cb", None), "isChecked", lambda: False)()),
        "z_stacks_only": bool(getattr(getattr(viewer, "spectro_filter_z_stack_cb", None), "isChecked", lambda: False)()),
        "matrix_only": bool(getattr(getattr(viewer, "spectro_filter_matrix_cb", None), "isChecked", lambda: False)()),
        "low_conf_only": bool(getattr(getattr(viewer, "spectro_filter_low_conf_cb", None), "isChecked", lambda: False)()),
    }


def _spec_passes_browser_filters(viewer, spec, flags):
    if flags.get("matrix_only") and not bool(spec.get("site_has_matrix") or spec.get("matrix_index") is not None):
        return False
    if flags.get("z_stacks_only") and not bool(spec.get("site_has_z_stack") or spec.get("xy_stack_z_varies")):
        return False
    if flags.get("low_conf_only") and str(spec.get("assignment_confidence") or "").strip().lower() != "low":
        return False
    if flags.get("current_image_only"):
        try:
            current_preview = str(viewer.last_preview[0]) if getattr(viewer, "last_preview", None) else ""
        except Exception:
            current_preview = ""
        image_key = str(spec.get("image_key") or spec.get("primary_image_key") or "")
        shared_keys = [str(key) for key in (spec.get("shared_image_keys") or []) if key]
        if current_preview:
            if image_key != current_preview and current_preview not in shared_keys:
                return False
        else:
            return False
    return True


def _set_browser_preview_to_image(viewer, image_key, text):
    try:
        viewer.spectro_preview_lbl.setText(text)
    except Exception:
        pass
    try:
        if image_key and image_key in viewer._thumb_labels:
            viewer.selected_file_for_thumbs = image_key
            viewer._refresh_thumb_selection_styles()
    except Exception:
        pass
    try:
        if hasattr(viewer, "_highlight_spectrum_entry"):
            viewer._highlight_spectrum_entry(None)
    except Exception:
        pass


def _open_browser_item(viewer, item):
    if not item:
        return
    payload = item.data(0, QtCore.Qt.UserRole)
    if not isinstance(payload, dict):
        return
    kind = str(payload.get("kind") or "")
    if kind == "spec":
        spec = payload.get("spec")
        if spec is not None and hasattr(viewer, "_show_spectro_popup"):
            viewer._show_spectro_popup(spec)
        return
    if kind == "site":
        spec = payload.get("spec")
        image_key = str(payload.get("image_key") or "")
        if spec is not None and hasattr(viewer, "_open_spectro_summary_for_site"):
            viewer._open_spectro_summary_for_site(spec, file_key=image_key, quiet=True)
        return
    if kind == "image":
        image_key = str(payload.get("image_key") or "")
        if image_key:
            viewer._open_spectro_summary_for_file(image_key, quiet=True)


def _browser_iter_items(root_item):
    if root_item is None:
        return
    yield root_item
    for idx in range(root_item.childCount()):
        child = root_item.child(idx)
        yield from _browser_iter_items(child)


def _select_first_browser_match(viewer, predicate=None):
    tree = getattr(viewer, "spectro_list", None)
    if tree is None:
        return False
    for top_idx in range(tree.topLevelItemCount()):
        item = tree.topLevelItem(top_idx)
        for candidate in _browser_iter_items(item):
            payload = candidate.data(0, QtCore.Qt.UserRole)
            if not isinstance(payload, dict):
                continue
            if predicate is not None and not predicate(payload):
                continue
            tree.setCurrentItem(candidate)
            tree.scrollToItem(candidate)
            return True
    return False


def _browser_payload_specs(payload):
    if not isinstance(payload, dict):
        return []
    kind = str(payload.get("kind") or "")
    if kind == "spec":
        spec = payload.get("spec")
        return [spec] if spec is not None else []
    if kind == "site":
        specs = list(payload.get("specs") or [])
        if specs:
            return specs
        spec = payload.get("spec")
        return [spec] if spec is not None else []
    return []


def _update_spectro_stats_label(viewer, stats=None):
    if not hasattr(viewer, "spectro_stats_label"):
        return
    thumb_markers = bool(getattr(viewer, "show_spectra", True))
    preview_markers = bool(getattr(viewer, "show_preview_spectra", thumb_markers))
    miniatures = bool(getattr(viewer, "show_spectro_miniatures", False))
    shared_repeats = bool(getattr(viewer, "spectro_share_overlapping_repeats", False))
    mode_text = (
        f"Thumbnail markers {'On' if thumb_markers else 'Off'} | "
        f"Preview {'On' if preview_markers else 'Off'} | "
        f"Miniatures {'On' if miniatures else 'Off'} | "
        f"Assignment {'Shared repeats' if shared_repeats else 'Single image'}"
    )
    assignment_tip = (
        "Assignment mode: shared repeats. Each spectrum first picks one primary image using spatial containment, acquisition order, and filename hints; spectra inside overlapping repeat scans are then shown on those repeats too."
        if shared_repeats
        else "Assignment mode: single image. Each spectrum is attached to one primary image using spatial containment first, then acquisition order, then weaker filename hints."
    )
    if getattr(viewer, "_spectros_loading", False):
        viewer.spectro_stats_label.setText(f"Spectroscopy loading...\n{mode_text}")
        viewer.spectro_stats_label.setToolTip(
            "Spectroscopy files are being scanned now. "
            "Thumbnail markers draw clickable points on image thumbnails. "
            "Preview markers draw the same points in the preview panel. "
            "Miniatures add separate spectroscopy cards into the thumbnail stream. "
            + assignment_tip
        )
        return
    if getattr(viewer, "_spectros_pending", False) and not getattr(viewer, "_spectros_loaded", False):
        viewer.spectro_stats_label.setText(f"Spectroscopy pending load\n{mode_text}")
        viewer.spectro_stats_label.setToolTip(
            "Spectroscopy scanning is deferred until a browser or visible spectroscopy mode needs it. "
            "Thumbnail markers draw clickable points on image thumbnails. "
            "Preview markers draw the same points in the preview panel. "
            "Miniatures add separate spectroscopy cards into the thumbnail stream. "
            + assignment_tip
        )
        return
    total = len(getattr(viewer, "spectros", []) or [])
    single_count = sum(1 for s in getattr(viewer, "spectros", []) if s.get("matrix_index") is None)
    low_conf_count = sum(
        1 for s in (getattr(viewer, "spectros", []) or [])
        if str(s.get("assignment_confidence") or "").strip().lower() == "low"
    )
    xy_stack_count = len({
        str(s.get("xy_stack_key"))
        for s in (getattr(viewer, "spectros", []) or [])
        if s.get("xy_stack_key") and int(s.get("xy_stack_count") or 0) > 1
    })
    site_count = sum(len(entries or []) for entries in (getattr(viewer, "spectro_sites_by_image", {}) or {}).values())
    if stats:
        total = stats.get("total_specs", total)
        single_count = stats.get("single_entries", single_count)
    matrix_datasets = getattr(viewer, "matrix_datasets", {}) or {}
    matrix_count = len(matrix_datasets)
    sample_ds = next(iter(matrix_datasets.values()), None)
    matrix_desc = ""
    if sample_ds:
        matrix_desc = f" ({sample_ds.cols}x{sample_ds.rows})"
    elif matrix_count == 0:
        matrix_desc = ""
    viewer.spectro_stats_label.setText(
        f"Spectra {total} | Sites {site_count} | Single {single_count} | Low conf {low_conf_count} | XY stacks {xy_stack_count} | Matrix {matrix_count}{matrix_desc}\n{mode_text}"
    )
    viewer.spectro_stats_label.setToolTip(
        f"Loaded spectroscopy entries: {total}. Single traces: {single_count}. "
        f"Image-relative sites: {site_count}. "
        f"Low-confidence assignments: {low_conf_count}. "
        f"Same-XY stacks: {xy_stack_count}. "
        f"Matrix datasets: {matrix_count}{matrix_desc}. "
        "Thumbnail markers draw clickable points on image thumbnails. "
        "Preview markers draw the same points in the preview panel. "
        "Miniatures add separate spectroscopy cards into the thumbnail stream. "
        + assignment_tip
    )


def _ensure_spectro_dock(viewer):
    if viewer.spectro_dock:
        return
    dock = QtWidgets.QDockWidget("Spectro Browser", viewer)
    dock.setFloating(True)
    container = QtWidgets.QWidget(dock)
    v = QtWidgets.QVBoxLayout(container)
    v.setContentsMargins(6, 6, 6, 6)
    v.setSpacing(6)
    viewer.spectro_search = QtWidgets.QLineEdit()
    viewer.spectro_search.setPlaceholderText("Search image, site, file, channel, or position")
    v.addWidget(viewer.spectro_search)
    filter_row = QtWidgets.QHBoxLayout()
    filter_row.setContentsMargins(0, 0, 0, 0)
    filter_row.setSpacing(6)
    viewer.spectro_filter_current_image_cb = QtWidgets.QCheckBox("Current image")
    viewer.spectro_filter_z_stack_cb = QtWidgets.QCheckBox("Z-stacks")
    viewer.spectro_filter_matrix_cb = QtWidgets.QCheckBox("Matrix")
    viewer.spectro_filter_low_conf_cb = QtWidgets.QCheckBox("Low confidence")
    filter_row.addWidget(viewer.spectro_filter_current_image_cb)
    filter_row.addWidget(viewer.spectro_filter_z_stack_cb)
    filter_row.addWidget(viewer.spectro_filter_matrix_cb)
    filter_row.addWidget(viewer.spectro_filter_low_conf_cb)
    filter_row.addStretch(1)
    v.addLayout(filter_row)
    viewer.spectro_list = QTreeWidget()
    viewer.spectro_list.setHeaderHidden(True)
    viewer.spectro_list.setRootIsDecorated(True)
    viewer.spectro_list.setUniformRowHeights(False)
    viewer.spectro_list.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
    v.addWidget(viewer.spectro_list, 1)
    viewer.spectro_preview_lbl = QLabel("Select a spectroscopy site or trace")
    viewer.spectro_preview_lbl.setAlignment(QtCore.Qt.AlignCenter)
    viewer.spectro_preview_lbl.setMinimumHeight(120)
    viewer.spectro_preview_lbl.setStyleSheet("QLabel { color: #999; }")
    v.addWidget(viewer.spectro_preview_lbl)
    container.setLayout(v)
    dock.setWidget(container)
    viewer.spectro_dock = dock
    viewer.spectro_search.textChanged.connect(viewer._filter_spectro_browser)
    viewer.spectro_filter_current_image_cb.toggled.connect(viewer._filter_spectro_browser)
    viewer.spectro_filter_z_stack_cb.toggled.connect(viewer._filter_spectro_browser)
    viewer.spectro_filter_matrix_cb.toggled.connect(viewer._filter_spectro_browser)
    viewer.spectro_filter_low_conf_cb.toggled.connect(viewer._filter_spectro_browser)
    viewer.spectro_list.currentItemChanged.connect(viewer._on_spectro_browser_selection)
    viewer.spectro_list.itemDoubleClicked.connect(viewer._on_spectro_browser_activate)
    if hasattr(viewer, "_on_spectro_browser_context_menu"):
        viewer.spectro_list.customContextMenuRequested.connect(viewer._on_spectro_browser_context_menu)


def open_spectro_browser(viewer, entries=None):
    viewer._ensure_spectro_dock()
    if entries is None:
        entries = list(viewer.spectros or [])
    viewer._spectro_browser_entries = list(entries)
    viewer._filter_spectro_browser()
    viewer.spectro_dock.show()
    viewer.spectro_dock.raise_()


def _filter_spectro_browser(viewer):
    if not hasattr(viewer, "spectro_list"):
        return
    txt = viewer.spectro_search.text().strip().lower() if hasattr(viewer, "spectro_search") else ""
    flags = _browser_filter_flags(viewer)
    viewer.spectro_list.clear()

    grouped = OrderedDict()
    for spec in list(getattr(viewer, "_spectro_browser_entries", []) or []):
        if not _spec_passes_browser_filters(viewer, spec, flags):
            continue
        image_key = str(spec.get("image_key") or spec.get("primary_image_key") or "")
        grouped.setdefault(image_key, OrderedDict())
        site_key = str(spec.get("site_key") or viewer._spec_identity_key(spec) or id(spec))
        grouped[image_key].setdefault(site_key, []).append(spec)

    global_index = 0
    for image_key, sites in grouped.items():
        image_name = Path(image_key).name if image_key else "Unassigned spectra"
        image_blob = f"{image_name} {image_key}".lower()
        image_specs = []
        visible_sites = []
        for site_key, site_specs in sites.items():
            first = site_specs[0]
            site_blob = " ".join((
                str(first.get("site_display") or ""),
                str(first.get("site_summary") or ""),
                str(first.get("xy_stack_summary") or ""),
                " ".join(str(ch) for ch in (first.get("site_channels") or [])),
            )).lower()
            site_matches = bool(txt and (txt in image_blob or txt in site_blob))
            visible_specs = []
            for spec in site_specs:
                if not txt or site_matches or txt in _spec_search_blob(spec):
                    visible_specs.append(spec)
            if visible_specs:
                visible_sites.append((site_key, visible_specs))
                image_specs.extend(visible_specs)
        if not visible_sites:
            continue

        image_item = QTreeWidgetItem([_image_tree_label(image_key, image_specs, len(visible_sites))])
        image_item.setToolTip(0, image_key or image_name)
        image_item.setData(0, QtCore.Qt.UserRole, {
            "kind": "image",
            "image_key": image_key,
            "specs": image_specs,
        })
        viewer.spectro_list.addTopLevelItem(image_item)

        for site_key, site_specs in visible_sites:
            first = site_specs[0]
            site_item = QTreeWidgetItem([_site_tree_label(site_specs)])
            site_summary = str(first.get("site_summary") or first.get("site_display") or "").strip()
            if site_summary:
                site_item.setToolTip(0, site_summary)
            site_item.setData(0, QtCore.Qt.UserRole, {
                "kind": "site",
                "image_key": image_key,
                "site_key": site_key,
                "spec": first,
                "specs": site_specs,
            })
            image_item.addChild(site_item)

            for spec in site_specs:
                global_index += 1
                leaf = QTreeWidgetItem([_spec_leaf_label(spec, index=global_index)])
                assignment_summary = str(spec.get("assignment_summary") or "").strip()
                assignment_conf = str(spec.get("assignment_confidence") or "").strip()
                tip_lines = [Path(spec.get("path", "")).name]
                pos = _spec_position_text(spec)
                if pos:
                    tip_lines.append(f"Position: {pos} nm")
                stack = str(spec.get("xy_stack_summary") or "").strip()
                if stack:
                    tip_lines.append(stack)
                if assignment_summary:
                    if assignment_conf:
                        tip_lines.append(f"Assignment: {assignment_summary} ({assignment_conf} confidence)")
                    else:
                        tip_lines.append(f"Assignment: {assignment_summary}")
                leaf.setToolTip(0, "\n".join(line for line in tip_lines if line))
                leaf.setData(0, QtCore.Qt.UserRole, {
                    "kind": "spec",
                    "image_key": image_key,
                    "site_key": site_key,
                    "spec": spec,
                })
                site_item.addChild(leaf)
            site_item.setExpanded(True)
        image_item.setExpanded(True)


def _on_spectro_browser_selection(viewer, current, _prev):
    if not current:
        if hasattr(viewer, "_highlight_spectrum_entry"):
            viewer._highlight_spectrum_entry(None)
        return
    payload = current.data(0, QtCore.Qt.UserRole)
    if not isinstance(payload, dict):
        if hasattr(viewer, "_highlight_spectrum_entry"):
            viewer._highlight_spectrum_entry(None)
        return

    kind = str(payload.get("kind") or "")
    if kind == "image":
        image_key = str(payload.get("image_key") or "")
        specs = list(payload.get("specs") or [])
        image_name = Path(image_key).name if image_key else "Unassigned spectra"
        site_count = len(list((getattr(viewer, "spectro_sites_by_image", {}) or {}).get(image_key, []) or []))
        text = f"{image_name}\n{len(specs)} spectra | {site_count} site" + ("" if site_count == 1 else "s")
        _set_browser_preview_to_image(viewer, image_key, text)
        return

    if kind == "site":
        specs = list(payload.get("specs") or [])
        first = payload.get("spec") or (specs[0] if specs else None)
        image_key = str(payload.get("image_key") or (first.get("image_key") if first else "") or "")
        site_summary = ""
        if first:
            site_summary = str(first.get("site_summary") or first.get("site_display") or "").strip()
        lines = [site_summary or "Site"]
        if specs:
            lines.append(f"{len(specs)} spectra")
        _set_browser_preview_to_image(viewer, image_key, "\n".join(line for line in lines if line))
        try:
            if first is not None and hasattr(viewer, "_highlight_spectrum_entry"):
                viewer._highlight_spectrum_entry(first)
        except Exception:
            pass
        return

    spec = payload.get("spec")
    if spec is None:
        if hasattr(viewer, "_highlight_spectrum_entry"):
            viewer._highlight_spectrum_entry(None)
        return

    try:
        x = spec.get("x")
        y = spec.get("y")
        lines = [Path(spec.get("path", "")).name, f"({x},{y})"]
        summary = str(spec.get("xy_stack_summary") or "").strip()
        if summary:
            lines.append(summary)
        site_summary = str(spec.get("site_summary") or spec.get("site_display") or "").strip()
        if site_summary:
            lines.append(site_summary)
        assignment_summary = str(spec.get("assignment_summary") or "").strip()
        assignment_conf = str(spec.get("assignment_confidence") or "").strip()
        if assignment_summary:
            if assignment_conf:
                lines.append(f"Assignment: {assignment_conf} confidence")
            lines.append(assignment_summary)
        viewer.spectro_preview_lbl.setText("\n".join(lines))
    except Exception:
        viewer.spectro_preview_lbl.setText(Path(spec.get("path", "")).name)

    try:
        image_key = spec.get("image_key")
        shared_keys = [str(key) for key in (spec.get("shared_image_keys") or []) if key]
        current_preview = str(viewer.last_preview[0]) if getattr(viewer, "last_preview", None) else ""
        if current_preview and current_preview in shared_keys:
            image_key = current_preview
        elif not image_key and shared_keys:
            image_key = shared_keys[0]
        if image_key and image_key in viewer._thumb_labels:
            viewer.selected_file_for_thumbs = image_key
            viewer._refresh_thumb_selection_styles()
    except Exception:
        pass
    try:
        if hasattr(viewer, "_highlight_spectrum_entry"):
            viewer._highlight_spectrum_entry(spec)
    except Exception:
        pass


def _on_spectro_browser_activate(viewer, item, _column=0):
    _open_browser_item(viewer, item)


def _on_spectro_browser_context_menu(viewer, pos):
    tree = getattr(viewer, "spectro_list", None)
    if tree is None:
        return
    item = tree.itemAt(pos)
    if item is None:
        return
    payload = item.data(0, QtCore.Qt.UserRole)
    if not isinstance(payload, dict):
        return
    kind = str(payload.get("kind") or "")
    menu = QtWidgets.QMenu(tree)

    open_action = QtWidgets.QAction("Open", menu)
    open_action.triggered.connect(lambda _=False, item=item: _open_browser_item(viewer, item))
    menu.addAction(open_action)

    specs = _browser_payload_specs(payload)
    if kind in {"site", "spec"} and specs:
        menu.addSeparator()
        current_image_key = ""
        if hasattr(viewer, "_current_spectro_assignment_target_image_key"):
            current_image_key = str(viewer._current_spectro_assignment_target_image_key() or "")
        current_image_name = Path(current_image_key).name if current_image_key else ""
        assign_label = "Assign to current image"
        if current_image_name:
            assign_label = f"Assign to current image ({current_image_name})"
        assign_action = QtWidgets.QAction(assign_label, menu)
        assign_action.setEnabled(bool(current_image_key))
        if current_image_key and hasattr(viewer, "_apply_spectro_assignment_override"):
            assign_action.triggered.connect(
                lambda _=False, specs=list(specs), image_key=current_image_key: viewer._apply_spectro_assignment_override(specs, image_key)
            )
        menu.addAction(assign_action)

        clear_action = QtWidgets.QAction("Clear manual assignment", menu)
        clear_action.setEnabled(any(str(spec.get("assignment_override_image_key") or "").strip() for spec in specs))
        if hasattr(viewer, "_clear_spectro_assignment_override"):
            clear_action.triggered.connect(
                lambda _=False, specs=list(specs): viewer._clear_spectro_assignment_override(specs)
            )
        menu.addAction(clear_action)

    menu.exec_(tree.viewport().mapToGlobal(pos))


__all__ = [
    "_update_spectro_stats_label",
    "_ensure_spectro_dock",
    "open_spectro_browser",
    "_filter_spectro_browser",
    "_on_spectro_browser_selection",
    "_on_spectro_browser_activate",
    "_on_spectro_browser_context_menu",
    "_select_first_browser_match",
]
