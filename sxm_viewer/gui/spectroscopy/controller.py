"""Spectroscopy assignment helpers for SXMGridViewer."""
from __future__ import annotations

import re
from functools import lru_cache
from pathlib import Path
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
from ...data.spectroscopy import is_matrix_file_entry


def _shared_repeat_spec_targets(viewer, spec, primary_key, repeat_groups, image_extents):
    primary_key = str(primary_key or "")
    group_keys = list(repeat_groups.get(primary_key) or [primary_key])
    if len(group_keys) <= 1:
        return [primary_key] if primary_key else []
    sx = spec.get("x")
    sy = spec.get("y")
    shared = []
    if sx is not None and sy is not None:
        for key in group_keys:
            ext = image_extents.get(str(key))
            if ext and viewer._spec_within_extent(sx, sy, ext, margin_frac=0.04):
                shared.append(str(key))
    if not shared:
        shared = [str(key) for key in group_keys]
    if primary_key and primary_key not in shared:
        shared.insert(0, primary_key)
    return list(OrderedDict((str(key), None) for key in shared).keys())


def _images_form_repeat_group(viewer, image_a, image_b, image_extents, *, overlap_threshold=0.82, size_tol_frac=0.15, angle_tol_deg=5.0):
    ext_a = image_extents.get(str(image_a.get("path")))
    ext_b = image_extents.get(str(image_b.get("path")))
    if not ext_a or not ext_b:
        return False
    try:
        ax0, ax1, ay1, ay0 = [float(v) for v in ext_a]
        bx0, bx1, by1, by0 = [float(v) for v in ext_b]
    except Exception:
        return False
    ax_min, ax_max = sorted((ax0, ax1))
    ay_min, ay_max = sorted((ay0, ay1))
    bx_min, bx_max = sorted((bx0, bx1))
    by_min, by_max = sorted((by0, by1))
    aw = ax_max - ax_min
    ah = ay_max - ay_min
    bw = bx_max - bx_min
    bh = by_max - by_min
    if min(aw, ah, bw, bh) <= 0:
        return False
    if abs(aw - bw) / max(aw, bw) > size_tol_frac:
        return False
    if abs(ah - bh) / max(ah, bh) > size_tol_frac:
        return False
    inter_w = min(ax_max, bx_max) - max(ax_min, bx_min)
    inter_h = min(ay_max, by_max) - max(ay_min, by_min)
    if inter_w <= 0 or inter_h <= 0:
        return False
    overlap_x = inter_w / min(aw, bw)
    overlap_y = inter_h / min(ah, bh)
    if overlap_x < overlap_threshold or overlap_y < overlap_threshold:
        return False
    try:
        header_a, _ = viewer.headers.get(str(image_a.get("path")), (None, None))
        header_b, _ = viewer.headers.get(str(image_b.get("path")), (None, None))
        angle_a = float(viewer._header_scan_angle(header_a)) if header_a is not None else 0.0
        angle_b = float(viewer._header_scan_angle(header_b)) if header_b is not None else 0.0
        if abs(angle_a - angle_b) > angle_tol_deg:
            return False
    except Exception:
        pass
    return True


def _repeat_image_groups(viewer, images, image_extents):
    if not images:
        return {}
    parents = list(range(len(images)))

    def find(idx):
        while parents[idx] != idx:
            parents[idx] = parents[parents[idx]]
            idx = parents[idx]
        return idx

    def union(a, b):
        ra = find(a)
        rb = find(b)
        if ra != rb:
            parents[rb] = ra

    for idx, image_a in enumerate(images):
        for jdx in range(idx + 1, len(images)):
            image_b = images[jdx]
            if _images_form_repeat_group(viewer, image_a, image_b, image_extents):
                union(idx, jdx)

    grouped = OrderedDict()
    for idx, image in enumerate(images):
        grouped.setdefault(find(idx), []).append(image)

    image_to_group = {}
    for group in grouped.values():
        group.sort(key=lambda img: img.get("time") or datetime.min)
        keys = [str(img.get("path")) for img in group if img.get("path")]
        for key in keys:
            image_to_group[key] = list(keys)
    return image_to_group


def _assign_spec_to_image_bucket(viewer, spec, image_key, image_time_index, *, shared_keys=None):
    image_key = str(image_key)
    specs_for_image = viewer.spectros_by_image[image_key]
    spec["image_key"] = image_key
    spec["shared_image_keys"] = list(shared_keys or [image_key])
    spec["shared_repeat_assignment"] = len(spec["shared_image_keys"]) > 1
    spec["order_idx"] = len(specs_for_image) + 1
    spec["display_time"] = image_time_index.get(image_key) or _spec_time_for_assignment(spec)
    specs_for_image.append(spec)


def _assignment_payload(reason, confidence, summary, *, image_key=None):
    reason_labels = {
        "manual_override": "Manual override",
        "preset_image_key": "Preset image reference",
        "xy_inside_extent_causal_time": "Inside image bounds + acquisition order",
        "xy_inside_extent": "Inside image bounds",
        "causal_time": "Acquisition order",
        "name_hint": "Filename/session hint",
        "xy_nearest_extent": "Nearest plausible image footprint",
        "closest_time_fallback": "Closest-time fallback",
        "first_image_fallback": "First-image fallback",
        "unknown": "Unknown",
    }
    return {
        "assignment_reason": str(reason or ""),
        "assignment_reason_label": str(reason_labels.get(str(reason or ""), str(reason or "").replace("_", " ").strip())),
        "assignment_confidence": str(confidence or ""),
        "assignment_summary": str(summary or ""),
        "assignment_image_key": str(image_key or ""),
    }


def _assign_spectros_to_images(viewer):
    """Assign spectroscopy entries to images using time and spatial sanity (prefer in-extent matches)."""
    viewer.spectros_by_image = defaultdict(list)
    images = list(getattr(viewer, 'image_meta', []) or [])
    specs = list(viewer.spectros or [])
    if not images or not specs:
        return
    image_time_index = {str(img.get("path")): img.get("time") for img in images}
    # precompute extents for images
    image_extents = {}
    for img in images:
        try:
            header, _fds = viewer.headers.get(str(img['path']), (None, None))
            extent = viewer._header_extent(header or {}) if header is not None else None
        except Exception:
            extent = None
        image_extents[str(img['path'])] = extent
    try:
        images.sort(key=lambda img: img.get('time') or datetime.min)
    except Exception:
        pass
    try:
        specs.sort(key=lambda s: _spec_time_for_assignment(s) or datetime.min)
    except Exception:
        pass

    share_repeat_scans = bool(getattr(viewer, "spectro_share_overlapping_repeats", False))
    repeat_groups = _repeat_image_groups(viewer, images, image_extents) if share_repeat_scans else {}
    image_paths = {str(img.get('path')) for img in images}
    image_paths_lower = {p.lower(): p for p in image_paths}
    image_by_key = {str(img.get("path")): img for img in images if img.get("path")}

    debug_nanonis = {"total": 0, "assigned": 0, "missing": 0}

    for spec in specs:
        primary_match = None
        assignment_meta = None
        override_key = spec.get("assignment_override_image_key")
        if override_key:
            mapped = image_paths_lower.get(str(override_key).lower())
            if mapped:
                primary_match = image_by_key.get(mapped)
                assignment_meta = _assignment_payload(
                    "manual_override",
                    "high",
                    "Assigned to an image selected manually by the user.",
                    image_key=mapped,
                )
        preset_key = spec.get('image_key')
        # Trust a cached image_key only for matrix (.3ds/grid) points, whose
        # single shared anchor is computed once by the dedicated matrix-
        # anchoring pass at scan time (_assign_matrix_reference in loader.py) -
        # re-running full per-point image matching wouldn't even apply the same
        # logic to them. Plain single spectra always re-match below instead of
        # trusting a preset: with nested overview/zoom scans (e.g. a 400 nm
        # overview and dozens of few-nm close-ups), the overview's extent
        # legitimately contains every zoomed-in position too, so "is this point
        # inside the preset image" can't tell a stale/coarse assignment from a
        # correct one - only the causal-time-aware matching below (which
        # prefers the most specific, chronologically-appropriate image among
        # ALL containing candidates) can, and it is cheap enough to redo on
        # every load. A stale preset on a single spectrum would otherwise
        # self-perpetuate forever, since every rescan re-persists whatever
        # image_key a spec dict currently carries into the manifest/disk cache.
        is_matrix_point = spec.get('matrix_index') is not None or is_matrix_file_entry(spec)
        if primary_match is None and preset_key and is_matrix_point:
            mapped = image_paths_lower.get(str(preset_key).lower())
            if mapped:
                primary_match = image_by_key.get(mapped)
                assignment_meta = _assignment_payload(
                    "preset_image_key",
                    "high",
                    "Assigned from an existing image reference already stored on this spectrum.",
                    image_key=mapped,
                )
            else:
                if spec.get('source') == 'nanonis_3ds':
                    debug_nanonis["total"] += 1
                    debug_nanonis["missing"] += 1
        match = primary_match
        if match is None:
            match, assignment_meta = viewer._choose_image_for_spec(spec, images, image_extents, with_details=True)
        if not match and images:
            # Fallback: pick closest by time, otherwise first image to avoid dropping markers.
            st = _spec_time_for_assignment(spec)
            if st is not None:
                try:
                    match = min(images, key=lambda img: abs((img.get('time') or datetime.min) - st))
                    assignment_meta = _assignment_payload(
                        "closest_time_fallback",
                        "low",
                        "Fallback assignment to the closest image in time because stronger spatial or filename cues were unavailable.",
                        image_key=str(match.get("path") or ""),
                    )
                except Exception:
                    match = images[0]
                    assignment_meta = _assignment_payload(
                        "first_image_fallback",
                        "low",
                        "Fallback assignment to the first available image because the time-based fallback failed.",
                        image_key=str(match.get("path") or ""),
                    )
            else:
                match = images[0]
                assignment_meta = _assignment_payload(
                    "first_image_fallback",
                    "low",
                    "Fallback assignment to the first available image because this spectrum had no usable assignment time.",
                    image_key=str(match.get("path") or ""),
                )
        if not match and images:
            match = images[0]
            assignment_meta = _assignment_payload(
                "first_image_fallback",
                "low",
                "Fallback assignment to the first available image because no stronger match was found.",
                image_key=str(match.get("path") or ""),
            )
        if not match:
            continue
        image_key = str(match['path'])
        assignment_meta = assignment_meta or _assignment_payload(
            "unknown",
            "low",
            "Assigned without recorded provenance.",
            image_key=image_key,
        )
        spec.update(assignment_meta)
        target_keys = [image_key]
        if share_repeat_scans:
            target_keys = _shared_repeat_spec_targets(viewer, spec, image_key, repeat_groups, image_extents)
        spec["primary_image_key"] = image_key
        spec["shared_image_keys"] = list(target_keys)
        spec["shared_repeat_assignment"] = len(target_keys) > 1
        if spec["shared_repeat_assignment"]:
            spec["assignment_summary"] = (
                f"{spec.get('assignment_summary') or 'Assigned to one primary image.'} "
                "Also shown on overlapping repeat scans."
            ).strip()
        _assign_spec_to_image_bucket(viewer, spec, image_key, image_time_index, shared_keys=target_keys)
        for shared_key in target_keys:
            if str(shared_key) == image_key:
                continue
            clone = dict(spec)
            _assign_spec_to_image_bucket(viewer, clone, shared_key, image_time_index, shared_keys=target_keys)
        if spec.get('source') == 'nanonis_3ds':
            debug_nanonis["total"] += 1
            debug_nanonis["assigned"] += 1
    # If nothing got assigned (e.g., all matches failed), place all specs on the first image to ensure visibility.
    if not viewer.spectros_by_image and images and specs:
        primary = images[0]
        image_key = str(primary['path'])
        for idx, spec in enumerate(specs, 1):
            target_keys = [image_key]
            if share_repeat_scans:
                target_keys = _shared_repeat_spec_targets(viewer, spec, image_key, repeat_groups, image_extents)
            spec.update(_assignment_payload(
                "first_image_fallback",
                "low",
                "Fallback assignment to the first available image because no spectra could be matched normally.",
                image_key=image_key,
            ))
            spec["primary_image_key"] = image_key
            spec["shared_image_keys"] = list(target_keys)
            spec["shared_repeat_assignment"] = len(target_keys) > 1
            _assign_spec_to_image_bucket(viewer, spec, image_key, image_time_index, shared_keys=target_keys)
            for shared_key in target_keys:
                if str(shared_key) == image_key:
                    continue
                clone = dict(spec)
                _assign_spec_to_image_bucket(viewer, clone, shared_key, image_time_index, shared_keys=target_keys)
    for k in list(viewer.spectros_by_image.keys()):
        viewer.spectros_by_image[k].sort(key=lambda s: (
            s.get('display_time') or s.get('time') or datetime.min,
            s.get('order_idx') or 0,
        ))
    _reassign_split_measurement_series(viewer)
    _annotate_xy_stacks(viewer)

    # Debug log for nanonis 3ds assignments
    # Suppress debug summary in normal runs


def _is_dat_spec(spec):
    try:
        path = spec.get('path') or ''
        return Path(path).suffix.lower() == '.dat'
    except Exception:
        return False


def _spec_time_for_assignment(spec):
    if _is_dat_spec(spec):
        return spec.get('file_mtime') or spec.get('time')
    return spec.get('time')


@lru_cache(maxsize=4096)
def _normalized_path_str(base):
    return str(Path(base))


def _spec_identity_key(spec):
    if not spec:
        return None
    base = spec.get("path")
    # Cached: a folder's specs share a small number of distinct file paths
    # (e.g. ~1440 spectra from one matrix file all have the same "path"),
    # so normalizing it fresh through Path() on every single spec is
    # thousands of redundant reconstructions of the same few strings.
    try:
        base = _normalized_path_str(base)
    except Exception:
        base = str(base)
    idx = spec.get("matrix_index")
    if idx is not None:
        return f"{base}#idx{idx}"
    x = spec.get("x")
    y = spec.get("y")
    if x is not None or y is not None:
        try:
            x_val = float(x) if x is not None else ""
            y_val = float(y) if y is not None else ""
            return f"{base}#pos{round(x_val, 6)}_{round(y_val, 6)}"
        except Exception:
            return f"{base}#pos{x}_{y}"
    return base


def _value_to_nm(value, unit_hint="nm"):
    try:
        val = float(value)
    except Exception:
        return None
    unit = str(unit_hint or "").strip().lower()
    if unit in ("m", "meter", "meters"):
        return val * 1e9
    if unit in ("um", "micron", "microns", "µm"):
        return val * 1e3
    if unit in ("pm", "picometer", "picometers"):
        return val * 1e-3
    return val


def _constant_axis_value_nm(values, unit_hint="nm", tol_nm=1e-3):
    try:
        arr = np.asarray(values, dtype=float).ravel()
    except Exception:
        return None
    if arr.size == 0:
        return None
    finite = arr[np.isfinite(arr)]
    if finite.size == 0:
        return None
    arr_nm = np.asarray([_value_to_nm(v, unit_hint) for v in finite], dtype=float)
    arr_nm = arr_nm[np.isfinite(arr_nm)]
    if arr_nm.size == 0:
        return None
    try:
        span = float(np.nanmax(arr_nm) - np.nanmin(arr_nm))
        center = float(np.nanmedian(arr_nm))
    except Exception:
        return None
    limit = max(float(tol_nm), abs(center) * 1e-6)
    if span <= limit:
        return center
    return None


_Z_LEVEL_TOKENS = ("z-controller", "absolute z", "z absolute", "z_abs", "abs z", "topo", "topography", "piezo")
_Z_LEVEL_TOKEN_RE = re.compile("|".join(re.escape(token) for token in _Z_LEVEL_TOKENS))
_Z_LEVEL_EXACT_LABELS = frozenset({"z", "z (m)", "z_nm"})


def _metadata_z_from_spec(spec):
    if not spec:
        return None, None, None
    best = None
    for key, value in (spec or {}).items():
        label = str(key or "").strip()
        if not label:
            continue
        label_low = label.lower()
        # This runs across every key of every spec (most of which aren't
        # z-channel related at all), so the token check needs to be cheap
        # for the common non-matching case: a single compiled-once regex
        # search instead of looping over a freshly-built tuple of substrings
        # with `any(...)` on every call.
        if not _Z_LEVEL_TOKEN_RE.search(label_low) and label_low not in _Z_LEVEL_EXACT_LABELS:
            continue
        unit_hint = ""
        if "(m)" in label_low:
            unit_hint = "m"
        elif "(nm)" in label_low:
            unit_hint = "nm"
        elif "(pm)" in label_low:
            unit_hint = "pm"
        elif "(um)" in label_low:
            unit_hint = "um"
        level = _value_to_nm(value, unit_hint=unit_hint)
        if level is None:
            continue
        score = 0
        preferred_label = None
        if "z-controller" in label_low and ">z" in label_low.replace("_", ">"):
            score = 120
            preferred_label = "Z piezo absolute"
        elif "z-controller" in label_low:
            score = 110
            preferred_label = "Z piezo absolute"
        elif "absolute z" in label_low or "z absolute" in label_low or "z_abs" in label_low or "abs z" in label_low:
            score = 100
            preferred_label = "Z piezo absolute"
        elif label_low.endswith("z_(m)") or label_low.endswith("z (m)") or label_low in {"z", "z_nm"}:
            score = 80
            preferred_label = "Z"
        elif "piezo" in label_low:
            score = 70
            preferred_label = "Z piezo"
        elif "topo" in label_low or "topography" in label_low:
            score = 20
            preferred_label = "Topo"
        clean_label = preferred_label or re.sub(r"\s*\(.*?\)", "", label).replace("_", " ").strip() or "Z"
        candidate = (score, float(level), clean_label, "nm")
        if best is None or candidate[0] > best[0]:
            best = candidate
    if best is None:
        return None, None, None
    return best[1], best[2], best[3]


def _extract_spec_z_level(spec):
    direct = spec.get("z_level_nm")
    if direct is not None:
        try:
            return float(direct), str(spec.get("z_level_label") or "Z"), str(spec.get("z_level_unit") or "nm")
        except Exception:
            pass
    for key, label in (
        ("z_abs_nm", "Z abs"),
        ("z_nm", "Z"),
    ):
        value = spec.get(key)
        if value is not None:
            try:
                return float(value), label, "nm"
            except Exception:
                continue
    level, label, unit = _metadata_z_from_spec(spec)
    if level is not None:
        spec["z_level_nm"] = float(level)
        spec["z_level_label"] = str(label or "Z")
        spec["z_level_unit"] = str(unit or "nm")
        return float(level), str(label or "Z"), str(unit or "nm")
    for axis in list(spec.get("AxisChoices") or []):
        try:
            key = str(axis.get("key") or "").strip().lower()
            label = str(axis.get("label") or key or "Z")
            unit = str(axis.get("unit") or "nm")
            low = label.lower()
            vals = axis.get("values")
        except Exception:
            continue
        if key not in {"topo", "z"} and not any(token in low for token in ("topo", "piezo", "z abs", "z_abs", "absolute z")):
            continue
        level = _constant_axis_value_nm(vals, unit_hint=unit)
        if level is not None:
            return level, label, "nm"
    channels = spec.get("channels") or {}
    unit_map = spec.get("unit_map") or {}
    if isinstance(channels, dict):
        for name, vals in channels.items():
            low = str(name or "").strip().lower()
            if not low or not any(token in low for token in ("topo", "piezo", "z_abs", "abs_z", "absolute z")):
                continue
            level = _constant_axis_value_nm(vals, unit_hint=unit_map.get(name) or "")
            if level is not None:
                return level, str(name), "nm"
    alt_vals = spec.get("AltAxis")
    alt_label = str(spec.get("AltAxisLabel") or "Z")
    if alt_vals is not None and any(token in alt_label.lower() for token in ("z", "topo", "piezo")):
        level = _constant_axis_value_nm(alt_vals, unit_hint=spec.get("AltAxisUnit") or "nm")
        if level is not None:
            return level, alt_label, "nm"
    topo_value = spec.get("topo_nm")
    if topo_value is not None:
        try:
            return float(topo_value), "Topo", "nm"
        except Exception:
            pass
    return None, None, None


def _clear_site_annotations(spec):
    for key in (
        "site_key",
        "site_image_key",
        "site_display",
        "site_summary",
        "site_member_count",
        "site_trace_count",
        "site_channel_count",
        "site_channels",
        "site_has_matrix",
        "site_has_z_stack",
        "site_z_label",
        "site_z_min_nm",
        "site_z_max_nm",
        "site_x_nm",
        "site_y_nm",
    ):
        spec.pop(key, None)


def _site_sort_tuple(site):
    try:
        return (
            float(site.get("y_nm")) if site.get("y_nm") is not None else float("inf"),
            float(site.get("x_nm")) if site.get("x_nm") is not None else float("inf"),
            str(site.get("display") or ""),
        )
    except Exception:
        return (float("inf"), float("inf"), str(site.get("display") or ""))


def _site_key_for_spec(spec, *, image_key, xy_tol_nm=0.05):
    image_key = str(image_key or "")
    try:
        sx = float(spec.get("x"))
        sy = float(spec.get("y"))
        qx = int(round(sx / xy_tol_nm))
        qy = int(round(sy / xy_tol_nm))
        return f"{image_key}::xy:{qx}:{qy}", sx, sy
    except Exception:
        pass
    if spec.get('matrix_index') is not None or is_matrix_file_entry(spec):
        row = spec.get("grid_row")
        col = spec.get("grid_col")
        if row is not None and col is not None:
            return f"{image_key}::grid:{row}:{col}", None, None
    identity = _spec_identity_key(spec) or f"{image_key}::spec:{id(spec)}"
    return f"{image_key}::spec:{identity}", None, None


def _format_position_nm(x_nm, y_nm):
    """User-facing name for an XY measurement position, e.g. 'Position 12.3/45.6 nm'."""
    try:
        if x_nm is None or y_nm is None:
            return ""
        return f"Position {float(x_nm):.1f}/{float(y_nm):.1f} nm"
    except Exception:
        return ""


def _derive_site_payload(site_key, image_key, members, x_nm, y_nm):
    member_count = len(members)
    channels = []
    channel_seen = set()
    has_matrix = False
    has_z_stack = False
    z_values = []
    z_label = None
    for spec in members:
        is_matrix_member = spec.get("matrix_index") is not None or is_matrix_file_entry(spec)
        has_matrix = has_matrix or bool(is_matrix_member)
        has_z_stack = has_z_stack or bool(spec.get("xy_stack_z_varies"))
        for name in list((spec.get("channels") or {}).keys()) + list(spec.get("available_channels") or []):
            name = str(name or "").strip()
            if not name or name in channel_seen:
                continue
            channel_seen.add(name)
            channels.append(name)
        if is_matrix_member:
            # xy_stack_z_varies (checked above) is only ever set on non-matrix
            # specs (_annotate_xy_stacks explicitly skips matrix entries), so
            # a matrix member can never make has_z_stack true and its Z level
            # is never read below - skip the (expensive, metadata-key-scanning)
            # extraction entirely. On a grid-heavy folder this is thousands of
            # wasted regex scans per site-building pass.
            continue
        level, label, _unit = _extract_spec_z_level(spec)
        if level is not None:
            z_values.append(float(level))
            if z_label is None and label:
                z_label = str(label)
    channel_count = len(channels)
    if x_nm is not None and y_nm is not None:
        display = _format_position_nm(x_nm, y_nm) or f"Position {site_key.split('::')[-1]}"
    elif has_matrix and member_count == 1:
        spec0 = members[0]
        display = f"Grid point {spec0.get('grid_row', '?')}/{spec0.get('grid_col', '?')}"
    else:
        display = f"Position {site_key.split('::')[-1]}"
    parts = [f"{member_count} " + ("spectrum" if member_count == 1 else "spectra")]
    if channel_count:
        parts.append(f"{channel_count} channel" + ("" if channel_count == 1 else "s"))
    if has_matrix:
        parts.append("grid map")
    if has_z_stack and z_values:
        z_min = min(z_values)
        z_max = max(z_values)
        parts.append(f"{z_label or 'Z'} series {z_min:.3f}-{z_max:.3f} nm")
    elif has_z_stack:
        parts.append("Z series")
    summary = f"{display}\n" + " | ".join(parts)
    payload = {
        "site_key": site_key,
        "site_image_key": str(image_key or ""),
        "site_display": display,
        "site_summary": summary,
        "site_member_count": member_count,
        "site_trace_count": member_count,
        "site_channel_count": channel_count,
        "site_channels": list(channels),
        "site_has_matrix": bool(has_matrix),
        "site_has_z_stack": bool(has_z_stack),
        "site_z_label": str(z_label or ""),
        "site_z_min_nm": min(z_values) if z_values else None,
        "site_z_max_nm": max(z_values) if z_values else None,
        "site_x_nm": x_nm,
        "site_y_nm": y_nm,
    }
    return payload


def _build_spectro_sites(viewer):
    originals = list(getattr(viewer, "spectros", []) or [])
    entries_by_image = getattr(viewer, "spectros_by_image", {}) or {}
    if not originals and not entries_by_image:
        viewer.spectro_sites_by_image = defaultdict(list)
        viewer.spectro_site_index = {}
        return
    for spec in originals:
        _clear_site_annotations(spec)
    groups_by_image = OrderedDict()
    primary_site_payloads = {}
    site_index = {}
    for image_key, entries in list(entries_by_image.items()):
        image_key = str(image_key or "")
        groups = OrderedDict()
        for spec in list(entries or []):
            _clear_site_annotations(spec)
            site_key, sx, sy = _site_key_for_spec(spec, image_key=image_key)
            group = groups.setdefault(site_key, {"members": [], "x_nm": sx, "y_nm": sy})
            if group.get("x_nm") is None and sx is not None:
                group["x_nm"] = sx
            if group.get("y_nm") is None and sy is not None:
                group["y_nm"] = sy
            group["members"].append(spec)
        sites = []
        for site_key, group in groups.items():
            payload = _derive_site_payload(site_key, image_key, group["members"], group.get("x_nm"), group.get("y_nm"))
            site = {
                "key": site_key,
                "image_key": image_key,
                "display": payload["site_display"],
                "summary": payload["site_summary"],
                "members": list(group["members"]),
                "x_nm": payload["site_x_nm"],
                "y_nm": payload["site_y_nm"],
                "channels": list(payload["site_channels"]),
                "channel_count": payload["site_channel_count"],
                "trace_count": payload["site_trace_count"],
                "has_matrix": payload["site_has_matrix"],
                "has_z_stack": payload["site_has_z_stack"],
                "z_label": payload["site_z_label"],
                "z_min_nm": payload["site_z_min_nm"],
                "z_max_nm": payload["site_z_max_nm"],
            }
            sites.append(site)
            site_index[site_key] = site
            for spec in group["members"]:
                spec.update(payload)
                if str(spec.get("primary_image_key") or spec.get("image_key") or "") == image_key:
                    identity = _spec_identity_key(spec)
                    if identity:
                        primary_site_payloads[identity] = dict(payload)
        sites.sort(key=_site_sort_tuple)
        groups_by_image[image_key] = sites
    for spec in originals:
        identity = _spec_identity_key(spec)
        payload = primary_site_payloads.get(identity)
        if payload:
            spec.update(payload)
    viewer.spectro_sites_by_image = defaultdict(list, groups_by_image)
    viewer.spectro_site_index = site_index


# Cluster detection tuning: kept as module constants rather than viewer config
# since these are geometry heuristics, not user preferences.
_GROUP_CLUSTER_MIN_MEMBERS = 4  # below this a "shape" isn't visually distinct from plain markers
_GROUP_CLUSTER_MAX_SITES = 400  # perf guard against the O(n^2) nearest-neighbor pass
_GROUP_CLUSTER_NN_FACTOR = 3.0  # sites within (factor x median nearest-neighbor distance) join a cluster


def _cluster_sites_by_proximity(sites):
    """Group sites into spatial clusters via single-linkage on nearest-neighbor
    distance, auto-scaled to this image's own spectrum spacing - a fixed nm
    threshold can't generalize across wildly different scan ranges. Sites that
    are each far from every other site stay singleton clusters, so a handful
    of occasional points scattered across an otherwise empty image are left
    alone rather than being forced into a group."""
    n = len(sites)
    if n < 2:
        return [[i] for i in range(n)]
    xs = [float(s["x_nm"]) for s in sites]
    ys = [float(s["y_nm"]) for s in sites]
    nn_dists = []
    for i in range(n):
        best = None
        for j in range(n):
            if i == j:
                continue
            d2 = (xs[i] - xs[j]) ** 2 + (ys[i] - ys[j]) ** 2
            if best is None or d2 < best:
                best = d2
        if best is not None:
            nn_dists.append(math.sqrt(best))
    if not nn_dists:
        return [[i] for i in range(n)]
    nn_dists.sort()
    median_nn = nn_dists[len(nn_dists) // 2]
    threshold2 = max(median_nn * _GROUP_CLUSTER_NN_FACTOR, 1e-6) ** 2

    parent = list(range(n))

    def _find(a):
        while parent[a] != a:
            parent[a] = parent[parent[a]]
            a = parent[a]
        return a

    def _union(a, b):
        ra, rb = _find(a), _find(b)
        if ra != rb:
            parent[ra] = rb

    for i in range(n):
        for j in range(i + 1, n):
            d2 = (xs[i] - xs[j]) ** 2 + (ys[i] - ys[j]) ** 2
            if d2 <= threshold2:
                _union(i, j)
    clusters = defaultdict(list)
    for i in range(n):
        clusters[_find(i)].append(i)
    return list(clusters.values())


def _axis_bin_count(values):
    """Number of distinct positions along one axis, after merging values that
    are within a data-driven tolerance of each other (raster grid coordinates
    are rarely bit-exact, but cluster tightly around a handful of steps)."""
    uniq = sorted(set(values))
    if len(uniq) <= 1:
        return len(uniq)
    gaps = sorted(uniq[i + 1] - uniq[i] for i in range(len(uniq) - 1))
    median_gap = gaps[len(gaps) // 2]
    tol = max(median_gap * 0.4, 1e-9)
    bins = 1
    for i in range(1, len(uniq)):
        if uniq[i] - uniq[i - 1] > tol:
            bins += 1
    return bins


def _classify_cluster_shape(member_sites):
    """Classify a spatial cluster of sites as a raster grid, a line, or an
    irregular cloud, purely from XY geometry (acquisition order is not used).
    Grid detection: quantize x/y independently into columns/rows and check the
    point count against columns x rows (allowing ~40% missing for aborted
    grids). Line vs. cloud: a highly elongated bounding box reads as a line;
    everything else is a cloud. Diagonal lines can still read as a cloud in
    this v1 heuristic - an acceptable simplification, not a correctness bug."""
    xs = [float(s["x_nm"]) for s in member_sites]
    ys = [float(s["y_nm"]) for s in member_sites]
    xmin, xmax = min(xs), max(xs)
    ymin, ymax = min(ys), max(ys)
    dx = max(xmax - xmin, 1e-9)
    dy = max(ymax - ymin, 1e-9)
    n_cols = _axis_bin_count(xs)
    n_rows = _axis_bin_count(ys)
    total = len(member_sites)
    if n_cols >= 2 and n_rows >= 2 and total >= 0.6 * n_cols * n_rows:
        return "grid", n_rows, n_cols
    if min(dx, dy) <= 0.2 * max(dx, dy):
        return "line", None, None
    return "cloud", None, None


def _detect_spectro_groups(viewer):
    """Cluster each image's non-grid-map (i.e. loose .dat) spectroscopy sites
    into spatial groups - a manual NxM grid of separate files, a line of
    positions, or a cloud concentrated in one area of the image - so the UI
    can show at a glance WHERE spectra were acquired instead of an undifferentiated
    scatter of markers. Isolated points with no nearby neighbors are left
    ungrouped and keep rendering as plain single markers. Must run after
    _build_spectro_sites (consumes viewer.spectro_sites_by_image)."""
    sites_by_image = getattr(viewer, "spectro_sites_by_image", {}) or {}
    groups_by_image = defaultdict(list)
    group_index = {}
    for spec in list(getattr(viewer, "spectros", []) or []):
        spec.pop("spectro_group_key", None)
        spec.pop("spectro_group_kind", None)
        spec.pop("spectro_group_display", None)
    for image_key, sites in sites_by_image.items():
        usable = [
            s for s in sites
            if not s.get("has_matrix") and s.get("x_nm") is not None and s.get("y_nm") is not None
        ]
        if len(usable) < _GROUP_CLUSTER_MIN_MEMBERS or len(usable) > _GROUP_CLUSTER_MAX_SITES:
            continue
        clusters = _cluster_sites_by_proximity(usable)
        group_idx = 0
        for idx_list in clusters:
            member_sites = [usable[i] for i in idx_list]
            if len(member_sites) < _GROUP_CLUSTER_MIN_MEMBERS:
                continue
            kind, grid_rows, grid_cols = _classify_cluster_shape(member_sites)
            xs = [float(s["x_nm"]) for s in member_sites]
            ys = [float(s["y_nm"]) for s in member_sites]
            members = [spec for s in member_sites for spec in s["members"]]
            group_key = f"{image_key}::sgroup:{group_idx}"
            group_idx += 1
            point_count = len(member_sites)
            if kind == "grid":
                display = f"{grid_cols}x{grid_rows} grid ({point_count} points)"
            elif kind == "line":
                display = f"Line of {point_count} points"
            else:
                display = f"Cloud of {point_count} points"
            channels = []
            seen_ch = set()
            for spec in members:
                for name in list((spec.get("channels") or {}).keys()) + list(spec.get("available_channels") or []):
                    name = str(name or "").strip()
                    if name and name not in seen_ch:
                        seen_ch.add(name)
                        channels.append(name)
            group = {
                "group_key": group_key,
                "image_key": image_key,
                "kind": kind,
                "site_keys": [s["key"] for s in member_sites],
                "members": members,
                "x_nm": sum(xs) / len(xs),
                "y_nm": sum(ys) / len(ys),
                "bbox_nm": (min(xs), max(xs), min(ys), max(ys)),
                "grid_rows": grid_rows,
                "grid_cols": grid_cols,
                "point_count": point_count,
                "trace_count": len(members),
                "channels": channels,
                "display": display,
                "summary": (
                    f"{display}\n{len(members)} spectra | "
                    f"{len(channels)} channel" + ("" if len(channels) == 1 else "s")
                ),
            }
            groups_by_image[image_key].append(group)
            group_index[group_key] = group
            for spec in members:
                spec["spectro_group_key"] = group_key
                spec["spectro_group_kind"] = kind
                spec["spectro_group_display"] = display
    viewer.spectro_groups_by_image = groups_by_image
    viewer.spectro_group_index = group_index


def _reassign_split_measurement_series(viewer):
    """Fix up repeated-measurement (XY-stack/Z-stack) series that got split
    across multiple images by the per-spec causal-time assignment above.

    That assignment operates per-file: each spectrum is attached to "the
    most recent image acquired before it". A series of repeats at one
    physical site is normally taken in quick succession (seconds apart),
    but if a new reference image happens to get captured partway through
    the series, the repeats before that image and the repeats after it
    resolve to two different images - even though they're all one
    measurement. Confirmed on real data: a 41-repeat Z-stack split into
    "1 spec on image A" + "40 specs on image B" because one new image was
    captured 8 seconds after the first repeat and 14 seconds before the
    second.

    Fix: after the causal-time pass, regroup by physical site (matching
    _annotate_xy_stacks's own site definition: rounded XY position, no
    per-image scoping) and majority-vote - reassign every member of a
    split site onto whichever image already holds the most of its repeats,
    so _annotate_xy_stacks (right after this) sees one consistent owner
    and produces one combined "Zx41" label instead of two fragments.
    Matrix/grid entries are excluded (matches _annotate_xy_stacks's own
    exclusion) - those are anchored as a whole grid by a separate,
    dedicated pass, not per-point, so they can't be split this way.
    """
    specs = list(getattr(viewer, "spectros", []) or [])
    if not specs:
        return
    xy_tol_nm = 0.05  # matches _annotate_xy_stacks's own site tolerance
    groups = OrderedDict()
    for spec in specs:
        # is_matrix_file_entry alone only catches Omicron/Anfatec matrix
        # .dat files (by filename); Nanonis .3ds grid points are flagged via
        # matrix_index instead - matching the is_matrix_point pattern used
        # above in _assign_spectros_to_images (both checks are needed to
        # exclude every matrix/grid spec). Confirmed the hard way: without
        # the matrix_index check, grid files that happen to share raster-
        # aligned XY coordinates (common - multiple grids scanning the same
        # sample region with similar pixel pitch) got treated as "split
        # sites" and shuffled between images by the thousands.
        if spec.get('matrix_index') is not None or is_matrix_file_entry(spec):
            continue
        try:
            sx = float(spec.get("x"))
            sy = float(spec.get("y"))
        except Exception:
            continue
        qx = int(round(sx / xy_tol_nm))
        qy = int(round(sy / xy_tol_nm))
        groups.setdefault((qx, qy), []).append(spec)

    for members in groups.values():
        if len(members) <= 1:
            continue
        counts = {}
        for spec in members:
            key = str(spec.get("image_key") or "")
            counts[key] = counts.get(key, 0) + 1
        if len(counts) <= 1:
            continue  # not split - every member already agrees
        max_count = max(counts.values())
        candidates = [key for key, cnt in counts.items() if cnt == max_count]
        if len(candidates) == 1:
            majority_key = candidates[0]
        else:
            # Deterministic tie-break: prefer whichever candidate image
            # holds the earliest-timestamped member of this site, rather
            # than relying on dict iteration order.
            first_time_by_key = {}
            for spec in members:
                key = str(spec.get("image_key") or "")
                if key not in candidates:
                    continue
                t = spec.get("time") or spec.get("display_time")
                if t is not None and (key not in first_time_by_key or t < first_time_by_key[key]):
                    first_time_by_key[key] = t
            majority_key = min(candidates, key=lambda k: first_time_by_key.get(k) or datetime.max)
        if not majority_key:
            continue

        for spec in members:
            old_key = str(spec.get("image_key") or "")
            if old_key == majority_key:
                continue
            bucket = viewer.spectros_by_image.get(old_key)
            if bucket is not None:
                try:
                    bucket.remove(spec)
                except ValueError:
                    pass
            spec["image_key"] = majority_key
            spec["primary_image_key"] = majority_key
            if spec.get("shared_image_keys") == [old_key]:
                spec["shared_image_keys"] = [majority_key]
            spec["assignment_reason"] = "zstack_series_majority"
            spec["assignment_confidence"] = "high"
            spec["assignment_summary"] = (
                "Consolidated onto the image most of this repeated-measurement "
                "series belongs to (the series was split across images because "
                "a new reference image was captured mid-series)."
            )
            viewer.spectros_by_image.setdefault(majority_key, []).append(spec)

    for k in list(viewer.spectros_by_image.keys()):
        viewer.spectros_by_image[k].sort(key=lambda s: (
            s.get('display_time') or s.get('time') or datetime.min,
            s.get('order_idx') or 0,
        ))


def _annotate_xy_stacks(viewer):
    originals = list(getattr(viewer, "spectros", []) or [])
    if not originals:
        return
    fields = (
        "xy_stack_key",
        "xy_stack_count",
        "xy_stack_display",
        "xy_stack_summary",
        "xy_stack_z_varies",
        "xy_stack_z_level_nm",
        "xy_stack_z_label",
        "xy_stack_z_min_nm",
        "xy_stack_z_max_nm",
    )
    for spec in originals:
        for key in fields:
            spec.pop(key, None)

    xy_tol_nm = 0.05
    z_tol_nm = 1e-3
    groups = OrderedDict()
    for spec in originals:
        # Both checks needed - is_matrix_file_entry alone only catches
        # Omicron/Anfatec matrix .dat files, not Nanonis .3ds grid points
        # (flagged via matrix_index instead). See
        # _reassign_split_measurement_series's comment for how this gap
        # manifests if left unfixed: grid files sharing raster-aligned XY
        # coordinates and the same image owner would get merged into one
        # bogus "Zx" xy-stack.
        if spec.get('matrix_index') is not None or is_matrix_file_entry(spec):
            continue
        try:
            sx = float(spec.get("x"))
            sy = float(spec.get("y"))
        except Exception:
            continue
        owners = [str(key) for key in (spec.get("shared_image_keys") or [spec.get("primary_image_key") or spec.get("image_key")]) if key]
        owners = sorted(OrderedDict((key, None) for key in owners).keys())
        owner_key = "|".join(owners) if owners else ""
        qx = int(round(sx / xy_tol_nm))
        qy = int(round(sy / xy_tol_nm))
        group_key = f"{owner_key}::{qx}:{qy}"
        groups.setdefault(group_key, []).append(spec)

    annotations = {}
    for group_key, members in groups.items():
        if len(members) <= 1:
            continue
        z_values = []
        z_label = None
        for spec in members:
            level, label, _unit = _extract_spec_z_level(spec)
            if level is not None:
                spec["xy_stack_z_level_nm"] = float(level)
                spec["xy_stack_z_label"] = str(label or "Z")
                z_values.append(float(level))
                if z_label is None and label:
                    z_label = str(label)
        unique_z = {round(val / z_tol_nm) for val in z_values} if z_values else set()
        z_varies = len(unique_z) > 1
        z_min = min(z_values) if z_values else None
        z_max = max(z_values) if z_values else None
        if z_varies and z_min is not None and z_max is not None:
            summary = f"Z series: {len(members)} spectra at one position\n{z_label or 'Z'} {z_min:.3f}-{z_max:.3f} nm"
            display = f"Zx{len(members)}"
        else:
            summary = f"Repeated spectra: {len(members)} at one position"
            display = f"x{len(members)}"
        for spec in members:
            identity = _spec_identity_key(spec)
            if not identity:
                continue
            annotations[identity] = {
                "xy_stack_key": group_key,
                "xy_stack_count": len(members),
                "xy_stack_display": display,
                "xy_stack_summary": summary,
                "xy_stack_z_varies": bool(z_varies),
                "xy_stack_z_level_nm": spec.get("xy_stack_z_level_nm"),
                "xy_stack_z_label": spec.get("xy_stack_z_label") or z_label,
                "xy_stack_z_min_nm": z_min,
                "xy_stack_z_max_nm": z_max,
            }
            spec.update(annotations[identity])

    for entries in list((getattr(viewer, "spectros_by_image", {}) or {}).values()):
        for spec in entries:
            identity = _spec_identity_key(spec)
            if not identity:
                continue
            ann = annotations.get(identity)
            if ann:
                spec.update(ann)
    _build_spectro_sites(viewer)
    _detect_spectro_groups(viewer)


def _choose_image_for_spec(viewer, spec, images, image_extents, *, with_details=False):
    """Pick the best image for a spectroscopy based on extent containment first, then time/hint."""
    def _result(match, reason, confidence, summary):
        image_key = str(match.get("path") or "") if match else ""
        details = _assignment_payload(reason, confidence, summary, image_key=image_key)
        if with_details:
            return match, details
        return match

    st = _spec_time_for_assignment(spec)
    sx = spec.get('x'); sy = spec.get('y')
    # Computed once and reused by both the .dat-specific branch below and the
    # generic fallback path further down - they used to run this same O(images)
    # containment scan twice for any spec that matched neither (e.g. one whose
    # position falls outside every image's extent), doubling the cost of the
    # single most expensive part of assignment for no different result.
    containment_candidates = None
    if sx is not None and sy is not None:
        containment_candidates = []
        for img in images:
            ext = image_extents.get(str(img['path']))
            if ext and viewer._spec_within_extent(sx, sy, ext, margin_frac=0.02):
                containment_candidates.append(img)
    if _is_dat_spec(spec):
        # Prefer spatial matching for .dat when coordinates are available.
        if sx is not None and sy is not None:
            candidates = containment_candidates
            if candidates:
                timed_match = _image_before_spec_time(candidates, st)
                if timed_match is not None:
                    return _result(
                        timed_match,
                        "xy_inside_extent_causal_time",
                        "high",
                        "Assigned to an image whose field of view contains the spectrum position, then resolved by acquisition order.",
                    )
                return _result(
                    candidates[0],
                    "xy_inside_extent",
                    "high",
                    "Assigned to an image whose field of view contains the spectrum position.",
                )
        # Next, prefer time ordering for .dat when coordinates don't match extents.
        time_match = _image_before_spec_time(images, st)
        if time_match is not None:
            return _result(
                time_match,
                "causal_time",
                "medium",
                "Assigned to the latest image acquired at or before the spectrum time.",
            )
        hint_match = None
        hint_score = -1
        try:
            hint_match, hint_score = viewer._match_spec_to_image_by_hint(spec, images, with_score=True)  # type: ignore[arg-type]
        except TypeError:
            # Backward compatibility if viewer overrides without new arg
            hint_match = viewer._match_spec_to_image_by_hint(spec, images)
        if hint_match is not None and hint_score is not None and hint_score >= 60:
            return _result(
                hint_match,
                "name_hint",
                "medium",
                "Assigned by a strong filename/session naming hint after spatial and time checks were inconclusive.",
            )
        if hint_match is not None:
            return _result(
                hint_match,
                "name_hint",
                "low",
                "Assigned by a weak filename/session naming hint because stronger cues were unavailable.",
            )
    # First pass: images whose extents contain the point (with a small margin)
    # - already computed above as containment_candidates; a .dat spec that
    # reaches this point already had its (empty) result, so no need to redo it.
    candidates = containment_candidates or []
    if sx is not None and sy is not None and candidates:
        timed_match = _image_before_spec_time(candidates, st)
        if timed_match is not None:
            return _result(
                timed_match,
                "xy_inside_extent_causal_time",
                "high",
                "Assigned to an image whose field of view contains the spectrum position, then resolved by acquisition order.",
            )
        return _result(
            candidates[0],
            "xy_inside_extent",
            "high",
            "Assigned to an image whose field of view contains the spectrum position.",
        )
    # Second pass: closest by space (even if slightly outside), then by time
    if sx is not None and sy is not None:
        scored = []
        for img in images:
            ext = image_extents.get(str(img['path']))
            if not ext:
                continue
            cx, cy = viewer._extent_center(ext)
            try:
                d2 = (float(sx) - cx) ** 2 + (float(sy) - cy) ** 2
            except Exception:
                continue
            scored.append((d2, img))
        if scored:
            scored.sort(key=lambda t: (t[0], abs(((t[1].get('time') or datetime.min) - st)) if st else datetime.max))
            best = scored[0][1]
            # ensure distance is not absurdly large compared to image span
            ext = image_extents.get(str(best['path']))
            if ext and viewer._spec_within_extent(sx, sy, ext, margin_frac=1.0):
                return _result(
                    best,
                    "xy_nearest_extent",
                    "medium",
                    "Assigned to the nearest plausible image footprint because the spectrum fell outside all image bounds.",
                )
    # Fallback: time-ordered + name hints
    if st:
        try:
            idx = 0
            n_img = len(images)
            while idx + 1 < n_img and (images[idx + 1].get('time') or datetime.max) <= st:
                idx += 1
            match = images[idx] if 0 <= idx < n_img else None
        except Exception:
            match = None
        if match:
            return _result(
                match,
                "causal_time",
                "medium",
                "Assigned to the latest image acquired at or before the spectrum time.",
            )
    hint_match = viewer._match_spec_to_image_by_hint(spec, images)
    if hint_match is not None:
        return _result(
            hint_match,
            "name_hint",
            "low",
            "Assigned by filename/session naming similarity because spatial and time cues were unavailable.",
        )
    if with_details:
        return None, None
    return None


def _image_before_spec_time(images, spec_time):
    if not images or spec_time is None:
        return None
    last_before = None
    for img in images:
        img_time = img.get('time') or datetime.min
        if img_time <= spec_time:
            last_before = img
        else:
            break
    if last_before is not None:
        return last_before
    try:
        return min(images, key=lambda img: abs((img.get('time') or datetime.min) - spec_time))
    except Exception:
        return images[0] if images else None


def _extent_center(viewer, extent):
    try:
        x0, x1, y1, y0 = extent
        cx = 0.5 * (x0 + x1)
        cy = 0.5 * (y0 + y1)
        return float(cx), float(cy)
    except Exception:
        return 0.0, 0.0


def _spec_within_extent(viewer, sx, sy, extent, margin_frac=0.05):
    try:
        x0, x1, y1, y0 = extent
        xmin, xmax = sorted((x0, x1))
        ymin, ymax = sorted((y0, y1))
        mx = (xmax - xmin) * margin_frac
        my = (ymax - ymin) * margin_frac
        xmin -= mx; xmax += mx; ymin -= my; ymax += my
        return xmin <= float(sx) <= xmax and ymin <= float(sy) <= ymax
    except Exception:
        return False


def _match_spec_to_image_by_hint(viewer, spec, images, *, with_score=False):
    def normalize(stem):
        stem = stem.lower().strip()
        stem = re.sub(r'(?:_matrix|-matrix).*$', '', stem)
        stem = stem.replace('-', '_')
        return stem
    spec_stem = normalize(Path(spec.get('path', '')).stem)
    if not spec_stem:
        return (None, -1) if with_score else None
    spec_tokens = [tok for tok in spec_stem.split('_') if tok]
    best = None
    best_score = -1
    for img in images:
        img_stem = normalize(Path(img['path']).stem)
        img_tokens = [tok for tok in img_stem.split('_') if tok]
        score = 0
        for a, b in zip(spec_tokens, img_tokens):
            if a == b:
                score += 10
            else:
                break
        common_prefix = 0
        for a, b in zip(spec_stem, img_stem):
            if a == b:
                common_prefix += 1
            else:
                break
        score += common_prefix
        if spec_stem in img_stem or img_stem in spec_stem:
            score += 50
        if score > best_score:
            best_score = score
            best = img
    if with_score:
        return best, best_score
    return best
__all__ = [
    "_assign_spectros_to_images",
    "_choose_image_for_spec",
    "_extent_center",
    "_spec_within_extent",
    "_match_spec_to_image_by_hint",
    "_image_before_spec_time",
    "_is_dat_spec",
    "_spec_time_for_assignment",
]



