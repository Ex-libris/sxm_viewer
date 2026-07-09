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
from . import popups as spectro_popups


def _spec_position_text(spec):
    try:
        if spec.get("x") is not None and spec.get("y") is not None:
            return f"{float(spec.get('x')):.1f}/{float(spec.get('y')):.1f}"
    except Exception:
        pass
    return ""


def _spec_detail_rows(spec):
    """Compact "Position / Channels / Assignment" detail lines for a solo
    (single-spectrum) entry's expandable children - the information that
    used to live in the now-removed "Position X/Y nm [...]" wrapper row,
    relocated under the file-name-first row it described."""
    rows = []
    pos = _spec_position_text(spec)
    if pos:
        rows.append(f"Position: {pos} nm")
    channel_count = spec.get("site_channel_count") or 0
    channels = list(spec.get("site_channels") or spec.get("available_channels") or [])
    if channel_count or channels:
        n = int(channel_count or len(channels))
        label = f"{n} channel" + ("" if n == 1 else "s")
        if channels:
            label = f"{label}: {', '.join(str(c) for c in channels[:6])}" + (", ..." if len(channels) > 6 else "")
        rows.append(label)
    assignment_summary = str(spec.get("assignment_summary") or "").strip()
    assignment_conf = str(spec.get("assignment_confidence") or "").strip()
    if assignment_summary:
        rows.append(f"Assignment ({assignment_conf or 'n/a'} confidence): {assignment_summary}")
    return rows


def _add_spec_detail_children(item, spec, image_key, site_key):
    for row_text in _spec_detail_rows(spec):
        child = QTreeWidgetItem([row_text])
        child.setData(0, QtCore.Qt.UserRole, {
            "kind": "spec",
            "image_key": image_key,
            "site_key": site_key,
            "spec": spec,
        })
        item.addChild(child)


def _spec_leaf_label(spec, index=None):
    # Off-frame/low-confidence status is now shown as an icon
    # (_browser_status_icon) rather than bracketed text - keeps leaf rows
    # scannable, especially for large flat lists (a 512-point grid).
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


def _site_common_name(site_specs):
    """Longest common filename prefix across a multi-member site's traces
    (stripped of a trailing run-number/index), e.g. "Z-Spectroscopy" for
    Z-Spectroscopy00001.dat..00016.dat - used so a Z-stack/grid cluster's
    group row still leads with something name-like, matching how a user
    would recognize the series in a folder explorer, rather than only a
    position value."""
    import re as _re
    stems = [Path(s.get("path", "")).stem for s in site_specs if s.get("path")]
    stems = [_re.sub(r"[\d_\-]+$", "", stem) for stem in stems]
    stems = [s for s in stems if s]
    if not stems:
        return ""
    common = stems[0]
    for stem in stems[1:]:
        n = 0
        while n < len(common) and n < len(stem) and common[n] == stem[n]:
            n += 1
        common = common[:n]
    return common.strip(" _-")


def _site_tree_label(site_specs):
    first = site_specs[0]
    display = str(first.get("site_display") or "Position").strip()
    trace_count = int(first.get("site_trace_count") or len(site_specs) or 0)
    channel_count = int(first.get("site_channel_count") or 0)
    low_conf_count = sum(
        1 for spec in list(site_specs or [])
        if str(spec.get("assignment_confidence") or "").strip().lower() == "low"
    )
    off_frame_count = sum(1 for spec in list(site_specs or []) if _spec_is_off_frame(spec))
    extras = [f"{trace_count} spectrum" if trace_count == 1 else f"{trace_count} spectra"]
    if channel_count:
        extras.append(f"{channel_count} ch")
    if low_conf_count:
        extras.append(f"low conf {low_conf_count}")
    if off_frame_count:
        extras.append(f"off-frame {off_frame_count}")
    if first.get("site_has_z_stack"):
        extras.append("Z series")
    elif int(first.get("xy_stack_count") or 0) > 1:
        extras.append("repeats")
    if first.get("site_has_matrix"):
        extras.append("grid map")
    name = _site_common_name(site_specs)
    title = f"{name} series" if name else display
    return f"{title}  [{display} | {' | '.join(extras)}]"


def _image_tree_label(image_key, specs, site_count, dataset_count=0):
    name = Path(str(image_key or "")).name if image_key else "Unassigned spectra"
    bits = [f"{site_count} position" + ("" if site_count == 1 else "s")]
    if dataset_count:
        bits.append(f"{dataset_count} grid" + ("" if dataset_count == 1 else "s"))
    bits.append(f"{len(specs)} spectra")
    return f"{name}  [{' | '.join(bits)}]"


def _matrix_dataset_label(viewer, dataset_key, ds_specs):
    """Grid dims/point count/channel count for a .3ds dataset's browser
    row, sourced from the already-parsed MatrixDataset (viewer.matrix_datasets)
    instead of implying one row per point."""
    ds = (getattr(viewer, "matrix_datasets", {}) or {}).get(dataset_key)
    first = ds_specs[0] if ds_specs else {}
    name = Path(str(first.get("path") or dataset_key or "Grid")).name
    detail = []
    if ds is not None:
        detail.append(f"{ds.cols}×{ds.rows} grid ({ds.cols * ds.rows} pts)")
    else:
        detail.append(f"{len(ds_specs)} pts")
    # ds.channels counts entries in the MatrixDataset's own channel-file
    # list, which for a Nanonis .3ds is always 1 (the whole grid is one
    # file with several measured channels inside) - not the actual number
    # of measured channels. site_channel_count/available_channels (already
    # used for the solo-item detail rows) gives the real per-point count
    # for both Nanonis and Omicron matrix formats.
    ch_count = int(first.get("site_channel_count") or 0) or len(first.get("available_channels") or ())
    if not ch_count and ds is not None:
        ch_count = len(ds.channels)
    if ch_count:
        detail.append(f"{ch_count} ch")
    pos = _spec_position_text(first)
    if pos:
        detail.append(f"centered {pos} nm")
    acquired = first.get("time") or first.get("display_time")
    if acquired:
        try:
            detail.append(acquired.strftime("%Y-%m-%d %H:%M"))
        except Exception:
            pass
    return f"{name}  [{' | '.join(detail)}]"


def _matrix_dataset_tooltip(viewer, dataset_key, ds_specs):
    ds = (getattr(viewer, "matrix_datasets", {}) or {}).get(dataset_key)
    first = ds_specs[0] if ds_specs else {}
    lines = [Path(str(first.get("path") or dataset_key or "")).name]
    if ds is not None:
        try:
            lines.append(ds.summary())
        except Exception:
            pass
    assignment_summary = str(first.get("assignment_summary") or "").strip()
    assignment_conf = str(first.get("assignment_confidence") or "").strip()
    if assignment_summary:
        if assignment_conf:
            lines.append(f"Assignment: {assignment_summary} ({assignment_conf} confidence)")
        else:
            lines.append(f"Assignment: {assignment_summary}")
    lines.append("Double-click or right-click → Open to launch the Grid Map Explorer.")
    return "\n".join(lines)


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


def _spec_is_off_frame(spec):
    return bool(spec.get("off_frame_direction"))


_BROWSER_ICON_SIZE = 18
_browser_icon_cache = {}


def _browser_type_icon(kind):
    """Small colored glyph summarizing a site's kind (single point, matrix/
    grid, Z-stack) - reuses the same accent colors as the on-thumbnail
    marker overlays (see gui/spectroscopy/overlays.py) so the browser's
    visual language matches what's drawn on the images themselves, instead
    of relying purely on bracketed text tags."""
    cached = _browser_icon_cache.get(kind)
    if cached is not None:
        return cached
    size = _BROWSER_ICON_SIZE
    pix = QPixmap(size, size)
    pix.fill(QtCore.Qt.transparent)
    painter = QPainter(pix)
    painter.setRenderHint(QPainter.Antialiasing, True)
    cx = cy = size / 2.0
    if kind == "matrix":
        color = QtGui.QColor(64, 200, 255)
        painter.setPen(QtCore.Qt.NoPen)
        painter.setBrush(QtGui.QBrush(color))
        step = 4.2
        start = cx - step
        for r in range(3):
            for c in range(3):
                painter.drawEllipse(QtCore.QPointF(start + c * step, start + r * step), 1.3, 1.3)
    elif kind == "zstack":
        color = QtGui.QColor(165, 141, 242)
        painter.setPen(QtCore.Qt.NoPen)
        bar_w, bar_h, gap = 11.0, 2.6, 1.6
        total_h = bar_h * 3 + gap * 2
        top = cy - total_h / 2.0
        for i, w in enumerate((1.0, 0.75, 0.55)):
            c = QtGui.QColor(color)
            c.setAlpha(max(120, 235 - i * 45))
            painter.setBrush(QtGui.QBrush(c))
            bw = bar_w * w
            painter.drawRoundedRect(QtCore.QRectF(cx - bw / 2.0, top + i * (bar_h + gap), bw, bar_h), bar_h / 2.0, bar_h / 2.0)
    else:  # "single"
        color = QtGui.QColor(255, 160, 0)
        painter.setPen(QtGui.QPen(QtGui.QColor(20, 12, 0), 1.0))
        painter.setBrush(QtGui.QBrush(color))
        painter.drawEllipse(QtCore.QPointF(cx, cy), 5.0, 5.0)
    painter.end()
    icon = QIcon(pix)
    _browser_icon_cache[kind] = icon
    return icon


def _browser_status_icon(*, low_conf=False, off_frame=False):
    """Small overlay glyph for assignment-status flags on a leaf row - a
    dashed amber ring for a low-confidence link, or the same orange flag
    used for off-frame markers on thumbnails (see _draw_off_frame_chevron in
    overlays.py). Returns None when neither flag applies, so callers can
    skip setting an icon entirely for the common (unflagged) case."""
    if not low_conf and not off_frame:
        return None
    key = ("status", bool(low_conf), bool(off_frame))
    cached = _browser_icon_cache.get(key)
    if cached is not None:
        return cached
    size = _BROWSER_ICON_SIZE
    pix = QPixmap(size, size)
    pix.fill(QtCore.Qt.transparent)
    painter = QPainter(pix)
    painter.setRenderHint(QPainter.Antialiasing, True)
    cx = cy = size / 2.0
    if off_frame:
        color = QtGui.QColor(255, 140, 0)
        outline = QtGui.QColor(20, 12, 0)
        painter.setPen(QtGui.QPen(outline, 1.4))
        painter.drawLine(QtCore.QPointF(cx - 4.5, cy - 5.5), QtCore.QPointF(cx - 4.5, cy + 5.5))
        path = QtGui.QPainterPath()
        path.moveTo(cx - 4.5, cy - 5.5)
        path.lineTo(cx + 5.5, cy - 2.2)
        path.lineTo(cx - 4.5, cy + 1.0)
        path.closeSubpath()
        painter.setPen(QtGui.QPen(outline, 1.0))
        painter.setBrush(QtGui.QBrush(color))
        painter.drawPath(path)
    elif low_conf:
        color = QtGui.QColor(255, 210, 96)
        pen = QtGui.QPen(color, 1.8)
        pen.setStyle(QtCore.Qt.DashLine)
        painter.setPen(pen)
        painter.setBrush(QtCore.Qt.NoBrush)
        painter.drawEllipse(QtCore.QPointF(cx, cy), 6.5, 6.5)
    painter.end()
    icon = QIcon(pix)
    _browser_icon_cache[key] = icon
    return icon


def _site_kind(site_specs):
    first = site_specs[0] if site_specs else {}
    if first.get("site_has_matrix"):
        return "matrix"
    if first.get("site_has_z_stack") or int(first.get("xy_stack_count") or 0) > 1:
        return "zstack"
    return "single"


def _browser_filter_flags(viewer):
    return {
        "current_image_only": bool(getattr(getattr(viewer, "spectro_filter_current_image_cb", None), "isChecked", lambda: False)()),
        "z_stacks_only": bool(getattr(getattr(viewer, "spectro_filter_z_stack_cb", None), "isChecked", lambda: False)()),
        "matrix_only": bool(getattr(getattr(viewer, "spectro_filter_matrix_cb", None), "isChecked", lambda: False)()),
        "low_conf_only": bool(getattr(getattr(viewer, "spectro_filter_low_conf_cb", None), "isChecked", lambda: False)()),
        "off_frame_only": bool(getattr(getattr(viewer, "spectro_filter_off_frame_cb", None), "isChecked", lambda: False)()),
    }


def _spec_passes_browser_filters(viewer, spec, flags):
    if flags.get("matrix_only") and not bool(spec.get("site_has_matrix") or spec.get("matrix_index") is not None):
        return False
    if flags.get("z_stacks_only") and not bool(spec.get("site_has_z_stack") or spec.get("xy_stack_z_varies")):
        return False
    if flags.get("low_conf_only") and str(spec.get("assignment_confidence") or "").strip().lower() != "low":
        return False
    if flags.get("off_frame_only") and not _spec_is_off_frame(spec):
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


def _update_browser_preview_image(viewer, image_key):
    """Render a real colored thumbnail of the assigned image - with all of
    its usual spectroscopy overlays (markers, footprints, off-frame flags,
    and the pulsing glow on whatever _highlight_spectrum_entry last set) -
    into the browser's preview panel, instead of leaving it as bare text.
    Call this AFTER _highlight_spectrum_entry so the glow reflects the
    current selection."""
    lbl = getattr(viewer, "spectro_preview_img_lbl", None)
    if lbl is None:
        return
    if not image_key or image_key not in getattr(viewer, "headers", {}):
        lbl.clear()
        lbl.setVisible(False)
        return
    try:
        header, fds = viewer.headers.get(str(image_key), (None, None))
        if not header or not fds:
            lbl.clear()
            lbl.setVisible(False)
            return
        channel_idx = viewer.channel_dropdown.currentIndex() if hasattr(viewer, "channel_dropdown") else 0
        if channel_idx < 0 or channel_idx >= len(fds):
            channel_idx = 0
        cmap = viewer.thumb_cmap_combo.currentText() if hasattr(viewer, "thumb_cmap_combo") else None
        cmap = cmap or getattr(viewer, "thumb_cmap", "viridis")
        size = max(200, min(lbl.width() or 200, lbl.height() or 200))
        pix = viewer._thumbnail_pixmap_for_file(str(image_key), channel_idx, size, size, cmap)
        if pix is None:
            lbl.clear()
            lbl.setVisible(False)
            return
        pix = pix.copy()
        try:
            viewer._decorate_thumbnail_pixmap(pix, str(image_key), channel_idx, header, fds)
        except Exception:
            pass
        lbl.setPixmap(pix.scaled(lbl.width(), lbl.height(), QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation))
        lbl.setVisible(True)
    except Exception:
        lbl.clear()
        lbl.setVisible(False)


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
    _update_browser_preview_image(viewer, image_key)


def _open_browser_item(viewer, item):
    if not item:
        return
    payload = item.data(0, QtCore.Qt.UserRole)
    if not isinstance(payload, dict):
        return
    kind = str(payload.get("kind") or "")
    if kind == "spec":
        # Previously called viewer._show_spectro_popup, a method that is
        # never defined anywhere in the codebase - every hasattr() guard on
        # it was always False, so "Open"/double-click on an individual
        # spectrum silently did nothing. _open_spectroscopy_popup is the
        # real function (used by every other click-to-open path in the app)
        # that actually opens the "Spectrum" trace window.
        spec = payload.get("spec")
        if spec is not None:
            spectro_popups._open_spectroscopy_popup(viewer, spec)
        return
    if kind == "site":
        # A genuine multi-member group (Z-stack/grid cluster at one
        # position - single-spectrum sites no longer create a "site" node,
        # see _filter_spectro_browser). Open all of them together as a
        # comparison window (or the single "Spectrum" popup if it turns out
        # to be just one) instead of the older "Spectroscopy at position..."
        # summary dialog, which read as a disconnected, out-of-place UI
        # relative to the rest of the click-to-open experience.
        specs = list(payload.get("specs") or [])
        first = payload.get("spec") or (specs[0] if specs else None)
        if not specs and first is not None:
            specs = [first]
        if specs:
            title = None
            site_display = str((first or {}).get("site_display") or "").strip()
            if site_display:
                title = f"Spectrum comparison - {site_display}"
            spectro_popups._open_spectroscopy_compare_popup(viewer, specs, title=title)
        return
    if kind == "matrix_dataset":
        # A grid can be 1000+ points - there is no per-point "Spectrum"
        # popup that makes sense here. Open the existing Grid Map Explorer
        # for exactly this .3ds dataset instead (dataset_key disambiguates
        # which grid when an image hosts several).
        image_key = str(payload.get("image_key") or "")
        dataset_key = str(payload.get("dataset_key") or "")
        if image_key and hasattr(viewer, "_open_matrix_explorer_for_file"):
            viewer._open_matrix_explorer_for_file(image_key, dataset_key=dataset_key or None)
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
            # Sites/images can be collapsed by default now (see
            # _BROWSER_AUTO_EXPAND_SITES/_SITE_SPECS) - a match inside a
            # collapsed branch must be expanded explicitly, or
            # setCurrentItem leaves it selected but invisible.
            ancestor = candidate.parent()
            while ancestor is not None:
                ancestor.setExpanded(True)
                ancestor = ancestor.parent()
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
    off_frame_count = sum(1 for s in (getattr(viewer, "spectros", []) or []) if _spec_is_off_frame(s))
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
        f"Spectra {total} | Sites {site_count} | Single {single_count} | Low conf {low_conf_count} | "
        f"Off-frame {off_frame_count} | XY stacks {xy_stack_count} | Matrix {matrix_count}{matrix_desc}\n{mode_text}"
    )
    viewer.spectro_stats_label.setToolTip(
        f"Loaded spectroscopy entries: {total}. Single traces: {single_count}. "
        f"Image-relative sites: {site_count}. "
        f"Low-confidence assignments: {low_conf_count}. "
        f"Off-frame (real position outside every image): {off_frame_count}. "
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
    dock.resize(460, 640)
    container = QtWidgets.QWidget(dock)
    v = QtWidgets.QVBoxLayout(container)
    v.setContentsMargins(8, 8, 8, 8)
    v.setSpacing(6)

    search_row = QtWidgets.QHBoxLayout()
    search_row.setContentsMargins(0, 0, 0, 0)
    search_row.setSpacing(6)
    viewer.spectro_search = QtWidgets.QLineEdit()
    viewer.spectro_search.setPlaceholderText("Search image, site, file, channel, or position")
    viewer.spectro_search.setClearButtonEnabled(True)
    search_row.addWidget(viewer.spectro_search, 1)
    v.addLayout(search_row)

    viewer.spectro_results_lbl = QLabel("")
    viewer.spectro_results_lbl.setObjectName("spectroResultsLabel")
    v.addWidget(viewer.spectro_results_lbl)

    filter_row = QtWidgets.QHBoxLayout()
    filter_row.setContentsMargins(0, 0, 0, 0)
    filter_row.setSpacing(6)
    viewer.spectro_filter_current_image_cb = QtWidgets.QCheckBox("Current image")
    viewer.spectro_filter_z_stack_cb = QtWidgets.QCheckBox("Z-stacks")
    viewer.spectro_filter_matrix_cb = QtWidgets.QCheckBox("Matrix")
    viewer.spectro_filter_low_conf_cb = QtWidgets.QCheckBox("Low confidence")
    viewer.spectro_filter_off_frame_cb = QtWidgets.QCheckBox("Off-frame")
    viewer.spectro_filter_off_frame_cb.setToolTip(
        "Show only spectra whose real position falls outside every image - "
        "assigned to the nearest plausible image as a fallback (often reference "
        "points acquired deliberately off to the side)."
    )
    filter_row.addWidget(viewer.spectro_filter_current_image_cb)
    filter_row.addWidget(viewer.spectro_filter_z_stack_cb)
    filter_row.addWidget(viewer.spectro_filter_matrix_cb)
    filter_row.addWidget(viewer.spectro_filter_low_conf_cb)
    filter_row.addWidget(viewer.spectro_filter_off_frame_cb)
    filter_row.addStretch(1)
    v.addLayout(filter_row)

    viewer.spectro_list = QTreeWidget()
    viewer.spectro_list.setHeaderHidden(True)
    viewer.spectro_list.setRootIsDecorated(True)
    viewer.spectro_list.setUniformRowHeights(False)
    viewer.spectro_list.setIndentation(16)
    viewer.spectro_list.setIconSize(QtCore.QSize(_BROWSER_ICON_SIZE, _BROWSER_ICON_SIZE))
    viewer.spectro_list.setAlternatingRowColors(True)
    viewer.spectro_list.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)

    preview_frame = QtWidgets.QFrame()
    preview_frame.setObjectName("spectroPreviewFrame")
    preview_frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
    preview_layout = QtWidgets.QHBoxLayout(preview_frame)
    preview_layout.setContentsMargins(6, 6, 6, 6)
    preview_layout.setSpacing(8)
    viewer.spectro_preview_img_lbl = QLabel()
    viewer.spectro_preview_img_lbl.setAlignment(QtCore.Qt.AlignCenter)
    # Fixed-size (120x120) made the reference-image miniature too small to
    # read clearly, and gave the user no way to see it bigger. Use a taller
    # minimum instead of a fixed size, and put the tree/preview in a
    # splitter (below) so the user can drag the divider to grow the preview
    # well past its default - same "let the user resize it" spirit as the
    # Position inset's corner-drag resize handle.
    viewer.spectro_preview_img_lbl.setMinimumSize(200, 200)
    viewer.spectro_preview_img_lbl.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
    viewer.spectro_preview_img_lbl.setScaledContents(False)
    preview_layout.addWidget(viewer.spectro_preview_img_lbl)
    viewer.spectro_preview_lbl = QLabel("Select a position or spectrum")
    viewer.spectro_preview_lbl.setAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignTop)
    viewer.spectro_preview_lbl.setWordWrap(True)
    preview_layout.addWidget(viewer.spectro_preview_lbl, 1)

    splitter = QtWidgets.QSplitter(QtCore.Qt.Vertical)
    splitter.addWidget(viewer.spectro_list)
    splitter.addWidget(preview_frame)
    splitter.setStretchFactor(0, 3)
    splitter.setStretchFactor(1, 2)
    splitter.setSizes([420, 260])
    v.addWidget(splitter, 1)

    container.setLayout(v)
    dock.setWidget(container)
    viewer.spectro_dock = dock
    viewer.spectro_search.textChanged.connect(viewer._filter_spectro_browser)
    viewer.spectro_filter_current_image_cb.toggled.connect(viewer._filter_spectro_browser)
    viewer.spectro_filter_z_stack_cb.toggled.connect(viewer._filter_spectro_browser)
    viewer.spectro_filter_matrix_cb.toggled.connect(viewer._filter_spectro_browser)
    viewer.spectro_filter_low_conf_cb.toggled.connect(viewer._filter_spectro_browser)
    viewer.spectro_filter_off_frame_cb.toggled.connect(viewer._filter_spectro_browser)
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


_BROWSER_AUTO_EXPAND_SITES = 8
_BROWSER_AUTO_EXPAND_SITE_SPECS = 20


def _update_browser_filter_counts(viewer):
    """Live counts on each filter checkbox (e.g. "Off-frame (2)") so the
    user knows before checking one whether there's anything to find,
    instead of clicking a checkbox that silently empties the tree."""
    entries = list(getattr(viewer, "_spectro_browser_entries", []) or [])
    z_stack_n = sum(1 for s in entries if bool(s.get("site_has_z_stack") or s.get("xy_stack_z_varies")))
    matrix_n = sum(1 for s in entries if bool(s.get("site_has_matrix") or s.get("matrix_index") is not None))
    low_conf_n = sum(1 for s in entries if str(s.get("assignment_confidence") or "").strip().lower() == "low")
    off_frame_n = sum(1 for s in entries if _spec_is_off_frame(s))
    for attr, label, n in (
        ("spectro_filter_z_stack_cb", "Z-stacks", z_stack_n),
        ("spectro_filter_matrix_cb", "Matrix", matrix_n),
        ("spectro_filter_low_conf_cb", "Low confidence", low_conf_n),
        ("spectro_filter_off_frame_cb", "Off-frame", off_frame_n),
    ):
        widget = getattr(viewer, attr, None)
        if widget is not None:
            widget.setText(f"{label} ({n})")


def _filter_spectro_browser(viewer):
    if not hasattr(viewer, "spectro_list"):
        return
    txt = viewer.spectro_search.text().strip().lower() if hasattr(viewer, "spectro_search") else ""
    flags = _browser_filter_flags(viewer)
    viewer.spectro_list.clear()
    _update_browser_filter_counts(viewer)

    grouped = OrderedDict()
    for spec in list(getattr(viewer, "_spectro_browser_entries", []) or []):
        if not _spec_passes_browser_filters(viewer, spec, flags):
            continue
        image_key = str(spec.get("image_key") or spec.get("primary_image_key") or "")
        bucket = grouped.setdefault(image_key, {"sites": OrderedDict(), "datasets": OrderedDict()})
        dataset_key = spec.get("matrix_dataset")
        if spec.get("matrix_index") is not None and dataset_key:
            # A grid can be 1024+ individual points - grouping those by
            # site (like single spectra) meant even one grid flooded the
            # tree with a row per point. Group by the .3ds file itself
            # instead: one row per grid, opened via the existing Grid Map
            # Explorer rather than drilled into point-by-point.
            bucket["datasets"].setdefault(str(dataset_key), []).append(spec)
            continue
        site_key = str(spec.get("site_key") or viewer._spec_identity_key(spec) or id(spec))
        bucket["sites"].setdefault(site_key, []).append(spec)

    single_image = len(grouped) <= 1
    global_index = 0
    shown_count = 0
    for image_key, bucket in grouped.items():
        sites = bucket["sites"]
        datasets = bucket["datasets"]
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

        visible_datasets = []
        for dataset_key, ds_specs in datasets.items():
            first = ds_specs[0]
            ds_blob = f"{Path(str(first.get('path') or '')).name} {dataset_key}".lower()
            ds_matches = bool(txt and (txt in image_blob or txt in ds_blob))
            visible_specs = [s for s in ds_specs if not txt or ds_matches or txt in _spec_search_blob(s)]
            if visible_specs:
                visible_datasets.append((dataset_key, visible_specs))
                image_specs.extend(visible_specs)

        if not visible_sites and not visible_datasets:
            continue
        shown_count += len(image_specs)

        image_item = QTreeWidgetItem([_image_tree_label(image_key, image_specs, len(visible_sites), len(visible_datasets))])
        image_item.setToolTip(0, image_key or image_name)
        bold_font = image_item.font(0)
        bold_font.setBold(True)
        image_item.setFont(0, bold_font)
        image_item.setData(0, QtCore.Qt.UserRole, {
            "kind": "image",
            "image_key": image_key,
            "specs": image_specs,
        })
        viewer.spectro_list.addTopLevelItem(image_item)

        for dataset_key, ds_specs in visible_datasets:
            ds_item = QTreeWidgetItem([_matrix_dataset_label(viewer, dataset_key, ds_specs)])
            ds_item.setIcon(0, _browser_type_icon("matrix"))
            ds_item.setToolTip(0, _matrix_dataset_tooltip(viewer, dataset_key, ds_specs))
            ds_item.setData(0, QtCore.Qt.UserRole, {
                "kind": "matrix_dataset",
                "image_key": image_key,
                "dataset_key": dataset_key,
                "specs": ds_specs,
            })
            image_item.addChild(ds_item)

        for site_key, site_specs in visible_sites:
            first = site_specs[0]
            if len(site_specs) == 1:
                # The overwhelming majority of positions are one file, one
                # trace. Wrapping each of those in a "Position X/Y nm
                # [1 spectrum]" parent + a near-duplicate "name.dat ..."
                # child was pure redundancy - two rows and a click to reveal
                # one file - and buried the filename (the thing users
                # actually recognize from Explorer) inside brackets on the
                # child. Collapse to one name-first row per file; position/
                # channel/assignment detail becomes its expandable children.
                global_index += 1
                spec = first
                solo_item = QTreeWidgetItem([_spec_leaf_label(spec, index=global_index)])
                solo_item.setIcon(0, _browser_type_icon(_site_kind(site_specs)))
                low_conf = str(spec.get("assignment_confidence") or "").strip().lower() == "low"
                off_frame = _spec_is_off_frame(spec)
                status_icon = _browser_status_icon(low_conf=low_conf, off_frame=off_frame)
                if status_icon is not None:
                    solo_item.setIcon(0, status_icon)
                site_summary = str(first.get("site_summary") or first.get("site_display") or "").strip()
                tip_lines = [Path(spec.get("path", "")).name] + _spec_detail_rows(spec)
                if site_summary:
                    tip_lines.append(site_summary)
                solo_item.setToolTip(0, "\n".join(line for line in tip_lines if line))
                solo_item.setData(0, QtCore.Qt.UserRole, {
                    "kind": "spec",
                    "image_key": image_key,
                    "site_key": site_key,
                    "spec": spec,
                })
                _add_spec_detail_children(solo_item, spec, image_key, site_key)
                image_item.addChild(solo_item)
                solo_item.setExpanded(False)
                continue

            site_item = QTreeWidgetItem([_site_tree_label(site_specs)])
            site_item.setIcon(0, _browser_type_icon(_site_kind(site_specs)))
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
                low_conf = str(spec.get("assignment_confidence") or "").strip().lower() == "low"
                off_frame = _spec_is_off_frame(spec)
                status_icon = _browser_status_icon(low_conf=low_conf, off_frame=off_frame)
                if status_icon is not None:
                    leaf.setIcon(0, status_icon)
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
            # Large single-image grids/stacks (hundreds of positions) used to
            # auto-expand every level unconditionally, producing a wall of
            # text with no way to get an overview first. Only expand a site
            # by default when its trace count is small enough to be useful
            # at a glance; searching or opening a filter still surfaces
            # matches directly via _select_first_browser_match.
            site_item.setExpanded(len(site_specs) <= _BROWSER_AUTO_EXPAND_SITE_SPECS)
        image_item.setExpanded(single_image or len(visible_sites) <= _BROWSER_AUTO_EXPAND_SITES)

    total_count = len(list(getattr(viewer, "_spectro_browser_entries", []) or []))
    if hasattr(viewer, "spectro_results_lbl"):
        if txt or any(flags.values()):
            viewer.spectro_results_lbl.setText(f"Showing {shown_count} of {total_count} spectra")
        else:
            viewer.spectro_results_lbl.setText(f"{total_count} spectra")


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

    if kind == "matrix_dataset":
        specs = list(payload.get("specs") or [])
        dataset_key = str(payload.get("dataset_key") or "")
        image_key = str(payload.get("image_key") or "")
        lines = [_matrix_dataset_label(viewer, dataset_key, specs), "Double-click to open in the Grid Map Explorer"]
        viewer.spectro_preview_lbl.setText("\n".join(line for line in lines if line))
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
        _update_browser_preview_image(viewer, image_key)
        return

    if kind == "site":
        specs = list(payload.get("specs") or [])
        first = payload.get("spec") or (specs[0] if specs else None)
        image_key = str(payload.get("image_key") or (first.get("image_key") if first else "") or "")
        site_summary = ""
        if first:
            site_summary = str(first.get("site_summary") or first.get("site_display") or "").strip()
        lines = [site_summary or "Position"]
        if specs:
            lines.append(f"{len(specs)} spectra")
        viewer.spectro_preview_lbl.setText("\n".join(line for line in lines if line))
        try:
            if image_key and image_key in viewer._thumb_labels:
                viewer.selected_file_for_thumbs = image_key
                viewer._refresh_thumb_selection_styles()
        except Exception:
            pass
        try:
            if first is not None and hasattr(viewer, "_highlight_spectrum_entry"):
                viewer._highlight_spectrum_entry(first)
        except Exception:
            pass
        # Highlight must be set BEFORE rendering the preview pixmap so its
        # pulsing glow (drawn by _decorate_thumbnail_pixmap from
        # self._highlighted_spec) actually shows up in this render.
        _update_browser_preview_image(viewer, image_key)
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

    image_key = spec.get("image_key")
    try:
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
    _update_browser_preview_image(viewer, image_key)


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
