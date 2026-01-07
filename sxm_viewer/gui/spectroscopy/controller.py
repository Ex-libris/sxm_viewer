"""Spectroscopy assignment helpers for SXMGridViewer."""
from __future__ import annotations

import re
from pathlib import Path
from ..._shared import *

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
        specs.sort(key=lambda s: s.get('time') or datetime.min)
    except Exception:
        pass

    for spec in specs:
        match = viewer._choose_image_for_spec(spec, images, image_extents)
        if not match:
            continue
        image_key = str(match['path'])
        spec['image_key'] = image_key
        viewer.spectros_by_image[image_key].append(spec)
    for k in list(viewer.spectros_by_image.keys()):
        viewer.spectros_by_image[k].sort(key=lambda s: s.get('time') or datetime.min)


def _choose_image_for_spec(viewer, spec, images, image_extents):
    """Pick the best image for a spectroscopy based on extent containment first, then time/hint."""
    st = spec.get('time')
    sx = spec.get('x'); sy = spec.get('y')
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


def _match_spec_to_image_by_hint(viewer, spec, images):
    def normalize(stem):
        stem = stem.lower().strip()
        stem = re.sub(r'(?:_matrix|-matrix).*$', '', stem)
        stem = stem.replace('-', '_')
        return stem
    spec_stem = normalize(Path(spec.get('path', '')).stem)
    if not spec_stem:
        return None
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
    return best
__all__ = [
    "_assign_spectros_to_images",
    "_choose_image_for_spec",
    "_extent_center",
    "_spec_within_extent",
    "_match_spec_to_image_by_hint",
]
