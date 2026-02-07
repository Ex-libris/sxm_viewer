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
def _assign_spectros_to_images(viewer):
    """Assign spectroscopy entries to images using time and spatial sanity (prefer in-extent matches)."""
    viewer.spectros_by_image = defaultdict(list)
    images = list(getattr(viewer, 'image_meta', []) or [])
    specs = list(viewer.spectros or [])
    if not images or not specs:
        return
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

    image_paths = {str(img.get('path')) for img in images}
    image_paths_lower = {p.lower(): p for p in image_paths}

    debug_nanonis = {"total": 0, "assigned": 0, "missing": 0}

    for spec in specs:
        preset_key = spec.get('image_key')
        if preset_key:
            mapped = image_paths_lower.get(str(preset_key).lower())
            if mapped:
                image_key = mapped
                specs_for_image = viewer.spectros_by_image[image_key]
                spec['order_idx'] = len(specs_for_image) + 1
                specs_for_image.append(spec)
                if spec.get('source') == 'nanonis_3ds':
                    debug_nanonis["total"] += 1
                    debug_nanonis["assigned"] += 1
                continue
            else:
                if spec.get('source') == 'nanonis_3ds':
                    debug_nanonis["total"] += 1
                    debug_nanonis["missing"] += 1
        match = viewer._choose_image_for_spec(spec, images, image_extents)
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
        spec['image_key'] = image_key
        specs_for_image = viewer.spectros_by_image[image_key]
        spec['order_idx'] = len(specs_for_image) + 1  # stable order for fallback placement
        specs_for_image.append(spec)
        if spec.get('source') == 'nanonis_3ds':
            debug_nanonis["total"] += 1
            debug_nanonis["assigned"] += 1
    # If nothing got assigned (e.g., all matches failed), place all specs on the first image to ensure visibility.
    if not viewer.spectros_by_image and images and specs:
        primary = images[0]
        image_key = str(primary['path'])
        for idx, spec in enumerate(specs, 1):
            spec['image_key'] = image_key
            spec['order_idx'] = idx
            viewer.spectros_by_image[image_key].append(spec)
    for k in list(viewer.spectros_by_image.keys()):
        viewer.spectros_by_image[k].sort(key=lambda s: s.get('time') or datetime.min)

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



