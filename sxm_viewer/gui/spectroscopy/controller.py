"""Spectroscopy assignment helpers for SXMGridViewer."""
from __future__ import annotations

import re
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
        preset_key = spec.get('image_key')
        if preset_key:
            mapped = image_paths_lower.get(str(preset_key).lower())
            if mapped:
                primary_match = image_by_key.get(mapped)
            else:
                if spec.get('source') == 'nanonis_3ds':
                    debug_nanonis["total"] += 1
                    debug_nanonis["missing"] += 1
        match = primary_match or viewer._choose_image_for_spec(spec, images, image_extents)
        if not match and images:
            # Fallback: pick closest by time, otherwise first image to avoid dropping markers.
            st = _spec_time_for_assignment(spec)
            if st is not None:
                try:
                    match = min(images, key=lambda img: abs((img.get('time') or datetime.min) - st))
                except Exception:
                    match = images[0]
            else:
                match = images[0]
        if not match and images:
            match = images[0]
        if not match:
            continue
        image_key = str(match['path'])
        target_keys = [image_key]
        if share_repeat_scans:
            target_keys = _shared_repeat_spec_targets(viewer, spec, image_key, repeat_groups, image_extents)
        spec["primary_image_key"] = image_key
        spec["shared_image_keys"] = list(target_keys)
        spec["shared_repeat_assignment"] = len(target_keys) > 1
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


def _spec_identity_key(spec):
    if not spec:
        return None
    base = spec.get("path")
    try:
        base = str(Path(base))
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


def _metadata_z_from_spec(spec):
    if not spec:
        return None, None, None
    best = None
    for key, value in list((spec or {}).items()):
        label = str(key or "").strip()
        if not label:
            continue
        label_low = label.lower()
        if not any(token in label_low for token in ("z-controller", "absolute z", "z absolute", "z_abs", "abs z", "topo", "topography", "piezo")) and label_low not in {"z", "z (m)", "z_nm"}:
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
        if is_matrix_file_entry(spec):
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
            summary = f"Z-stack: {len(members)} spectra at one XY\n{z_label or 'Z'} {z_min:.3f} to {z_max:.3f} nm"
            display = f"Zx{len(members)}"
        else:
            summary = f"Coincident spectra: {len(members)} at one XY"
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


def _choose_image_for_spec(viewer, spec, images, image_extents):
    """Pick the best image for a spectroscopy based on extent containment first, then time/hint."""
    st = _spec_time_for_assignment(spec)
    sx = spec.get('x'); sy = spec.get('y')
    if _is_dat_spec(spec):
        # Prefer spatial matching for .dat when coordinates are available.
        if sx is not None and sy is not None:
            candidates = []
            for img in images:
                ext = image_extents.get(str(img['path']))
                if ext and viewer._spec_within_extent(sx, sy, ext, margin_frac=0.02):
                    candidates.append(img)
            if candidates:
                if st:
                    candidates.sort(key=lambda img: abs((img.get('time') or datetime.min) - st))
                return candidates[0]
        # Next, prefer time ordering for .dat when coordinates don't match extents.
        time_match = _image_before_spec_time(images, st)
        if time_match is not None:
            return time_match
        hint_match = None
        hint_score = -1
        try:
            hint_match, hint_score = viewer._match_spec_to_image_by_hint(spec, images, with_score=True)  # type: ignore[arg-type]
        except TypeError:
            # Backward compatibility if viewer overrides without new arg
            hint_match = viewer._match_spec_to_image_by_hint(spec, images)
        if hint_match is not None and hint_score is not None and hint_score >= 60:
            return hint_match
        if hint_match is not None:
            return hint_match
    candidates = []
    # First pass: images whose extents contain the point (with a small margin)
    if sx is not None and sy is not None:
        for img in images:
            ext = image_extents.get(str(img['path']))
            if ext and viewer._spec_within_extent(sx, sy, ext, margin_frac=0.02):
                candidates.append(img)
        if candidates:
            if st:
                candidates.sort(key=lambda img: abs((img.get('time') or datetime.min) - st))
            return candidates[0]
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
                return best
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
            return match
    return viewer._match_spec_to_image_by_hint(spec, images)


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



