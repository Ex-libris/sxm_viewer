"""Virtual copies: non-destructive derived images.

Every crop/rotate/channel-extract in this app produces a **virtual copy**
rather than touching the source file - a synthetic entry in
`viewer._processed_views` that behaves like a real image in the thumbnail
grid but owns its own array. This module creates them, from all the
routes a user can trigger: a popup view, a drag payload, a channel pick,
a crop, or a replayed history step.

Extracted from `main_window.py` (230 lines) - see
`docs/refactor/GOD_CLASS_PLAN.md`. Module functions taking `viewer` first,
per the `gui/viewer/*.py` convention.

Key invariant: the source file's data is never modified. A copy carries a
tag (`[crop]`, `[edit]`, ...) so the grid can show its provenance, and
`_normalize_virtual_copy_order` keeps it adjacent to its source.
"""
from __future__ import annotations

from pathlib import Path

from .._shared import QtCore, QtWidgets, np


def virtual_copy_source_anchor(viewer, view):
    if not view:
        return VIRTUAL_COPY_INSERT_START
    path = view.get("path") or (view.get("meta") or {}).get("path") or (view.get("meta") or {}).get("file_path")
    return str(path) if path else VIRTUAL_COPY_INSERT_START

def create_virtual_copy_from_popup_view(viewer, view):
    return create_virtual_view_copy(viewer, view, insert_after_key=virtual_copy_source_anchor(viewer, view))

def create_virtual_copy_from_drag_payload(viewer, payload, insert_after_key=None):
    if not isinstance(payload, dict):
        return None
    drag_token = payload.get("view_drag_token")
    if drag_token:
        view = MultiPreviewCanvas.consume_drag_view_snapshot(drag_token)
        if view:
            return create_virtual_view_copy(viewer, view, insert_after_key=insert_after_key)
    file_path = payload.get("file_path")
    channel_idx = payload.get("channel_index")
    if not file_path or channel_idx is None:
        return None
    try:
        channel_idx = int(channel_idx)
    except Exception:
        return None
    created = create_virtual_channel_copies(viewer, 
        [str(file_path)],
        channel_idx=channel_idx,
        insert_after_key=insert_after_key,
    )
    return created

def create_virtual_channel_copies(viewer, paths, channel_idx=None, insert_after_key=None):
    """Create virtual copies of selected images for a specific channel."""
    if not paths:
        return 0
    targets = [str(Path(p)) for p in paths]
    # If channel not provided, ask the user using first file's channels
    if channel_idx is None:
        first = targets[0]
        header, fds = viewer.headers.get(first, (None, None))
        if header is None or fds is None:
            header, fds = parse_header(Path(first))
        if not fds:
            return
        channel_idx = viewer._choose_channel_index_for_virtual_copy(
            fds,
            current_idx=viewer.channel_dropdown.currentIndex(),
        )
        if channel_idx is None:
            return 0
    added = 0
    anchor_key = insert_after_key
    for p in targets:
        try:
            header, fds = viewer.headers.get(p, (None, None))
            if header is None or fds is None:
                header, fds = parse_header(Path(p))
            if not fds or channel_idx < 0 or channel_idx >= len(fds):
                continue
            # Build arrays for all channels so switching works
            arr_by_channel = {}
            for ch_idx, fd in enumerate(fds):
                try:
                    arr_by_channel[ch_idx] = np.array(viewer._get_channel_array(p, ch_idx, header, fd), copy=True)
                except Exception:
                    continue
            if not arr_by_channel:
                continue
            fds_new = [dict(fd) for fd in fds]
            for i, fd_new in enumerate(fds_new):
                fd_new['FileName'] = f"{Path(p).name}_virt_ch{i}"
                fd_new['Caption'] = f"{fd_new.get('Caption') or Path(p).name} [ch{i}]"
            key = viewer._make_processed_key(p, op="ch", channel_idx=channel_idx)
            viewer._processed_views[key] = {
                'arr_by_channel': arr_by_channel,
                'header': dict(header),
                'fds': fds_new,
                'channel_idx': channel_idx,
                'source': p,
                'label': f"[ch{channel_idx}]",
                'op': 'channel',
            }
            viewer.headers[key] = (dict(header), fds_new)
            viewer._inherit_star_for_virtual_copy(key, p)
            viewer._insert_processed_after_source(key, p, insert_after_key=anchor_key)
            if insert_after_key not in (None, "", VIRTUAL_COPY_INSERT_START):
                anchor_key = key
            added += 1
        except Exception:
            continue
    if added:
        viewer.populate_thumbnails_for_channel(viewer.channel_dropdown.currentIndex())
    return added

def create_virtual_view_copy(viewer, view, insert_after_key=None, tag=None, op=None):
    """Create a virtual thumbnail copy from the current popup/preview view snapshot."""
    if not view:
        return None
    path = view.get("path") or (view.get("meta") or {}).get("path") or (view.get("meta") or {}).get("file_path")
    arr = view.get("arr")
    ch_idx = view.get("channel_idx")
    if ch_idx is None:
        ch_idx = (view.get("meta") or {}).get("channel_index")
    if path is None or arr is None:
        return None
    try:
        arr = np.asarray(arr)
    except Exception:
        return None
    if arr.ndim < 2 or arr.size == 0:
        return None
    path = str(path)
    try:
        header, fds = viewer.headers.get(path, (None, None))
        if header is None or fds is None:
            header, fds = parse_header(Path(path))
        if not fds:
            return None
        ch_idx = int(ch_idx) if ch_idx is not None else 0
        title = str(view.get("title") or "")
        inferred_crop = bool(view.get("crop_sequence") is not None or "[crop]" in title.lower())
        tag = str(tag or ("[crop]" if inferred_crop else "[copy]"))
        op_name = str(op or ("crop" if inferred_crop else "copy"))
        arr_by_channel = {ch_idx: np.array(arr, copy=True)}
        fds_new = [dict(fd) for fd in fds]
        for i, fd_new in enumerate(fds_new):
            base_caption = fd_new.get("Caption") or Path(path).name
            fd_new["FileName"] = f"{Path(path).name}_{op_name}_ch{i}"
            fd_new["Caption"] = f"{base_caption} {tag}"
        # A copy's array doesn't always represent the same physical
        # quantity as the anchor's own channel ch_idx (e.g. a matrix-fit
        # parameter map borrows a real scan image purely as a header/units
        # template - its values are LCPD/RMSE/etc, not whatever channel
        # ch_idx originally measured). Without this, _get_filtered_
        # channel_array's normalize_unit_and_data(arr, fd['PhysUnit'])
        # would silently label/scale the copy using the anchor's
        # original unit. fd_overrides lets the caller correct
        # PhysUnit/Scale/Offset/Caption/etc. post-clone for exactly
        # this case; grid-slice/crop copies (which do share the anchor's
        # own units) simply don't pass it.
        fd_overrides = view.get("fd_overrides")
        if fd_overrides and 0 <= ch_idx < len(fds_new):
            fds_new[ch_idx].update(fd_overrides)
        header_new = dict(header)
        header_new["xPixel"] = int(arr.shape[1])
        header_new["yPixel"] = int(arr.shape[0])
        stored_extent = None
        view_extent = view.get("extent_raw")
        if view_extent is None:
            view_extent = view.get("extent")
        if view_extent is not None and len(view_extent) == 4:
            try:
                x0, x1, y_a, y_b = [float(v) for v in view_extent]
                # Every regular header-driven image in this app relies on
                # one fixed invariant: array row 0 corresponds to the
                # SMALLER real-world y (the Nanonis adapter's own
                # Direction-based row flip exists specifically to
                # guarantee this at conversion time - see providers/
                # nanonis/adapter.py). y_a/y_b here follow the (x0, x1,
                # bottom, top) imshow-extent convention the caller drew
                # with, where y_b ("top") is where the array's row 0 was
                # actually placed on screen - normally that's already
                # the smaller value (y_a > y_b), but a source view whose
                # row 0 happens to be the *larger* value (confirmed for
                # MatrixSpectroViewer's grid-slice virtual copies, whose
                # own row order has no relationship to this convention)
                # would otherwise silently invert once XScanRange/
                # yCenter below collapse the extent to a symmetric
                # range+center and lose that distinction - producing a
                # vertical mirror the moment this copy is redrawn
                # through the standard header_extent-based pipeline.
                if y_b > y_a:
                    arr = np.flipud(arr)
                    arr_by_channel[ch_idx] = arr
                    y_a, y_b = y_b, y_a
                stored_extent = (x0, x1, y_a, y_b)
                xmin, xmax = sorted((x0, x1))
                ymin, ymax = sorted((y_a, y_b))
                x_range = max(xmax - xmin, 1e-12)
                y_range = max(ymax - ymin, 1e-12)
                x_center = 0.5 * (xmin + xmax)
                y_center = 0.5 * (ymin + ymax)
                header_new["XScanRange"] = x_range
                header_new["YScanRange"] = y_range
                header_new["XRange"] = x_range
                header_new["YRange"] = y_range
                header_new["xCenter"] = x_center
                header_new["yCenter"] = y_center
                header_new["XCenter"] = x_center
                header_new["YCenter"] = y_center
            except Exception:
                pass
        key = viewer._make_processed_key(path, op=op_name, channel_idx=ch_idx)
        viewer._processed_views[key] = {
            "arr_by_channel": arr_by_channel,
            "header": header_new,
            "fds": fds_new,
            "channel_idx": ch_idx,
            "source": path,
            "extent_raw": stored_extent,
            "label": tag,
            "op": op_name,
        }
        viewer.headers[key] = (header_new, fds_new)
        viewer._inherit_star_for_virtual_copy(key, path)
        viewer._insert_processed_after_source(key, path, insert_after_key=insert_after_key)
        viewer.populate_thumbnails_for_channel(viewer.channel_dropdown.currentIndex())
        return key
    except Exception:
        return None

def create_virtual_crop_view(viewer, view, insert_after_key=None):
    """Create a virtual copy from a cropped preview view (single channel)."""
    return create_virtual_view_copy(viewer, view, insert_after_key=insert_after_key, tag="[crop]", op="crop")

def create_virtual_copy_from_history(viewer, seq):
    if seq is None:
        return
    entry = None
    if hasattr(viewer, "quick_crop_controller"):
        entry = viewer.quick_crop_controller.get_history_entry(seq)
    if not entry:
        return
    view_snapshot = entry.get("view_snapshot")
    if not view_snapshot:
        QtWidgets.QMessageBox.information(viewer, "Virtual copy", "This crop does not have a stored snapshot.")
        return
    create_virtual_crop_view(viewer, dict(view_snapshot))

