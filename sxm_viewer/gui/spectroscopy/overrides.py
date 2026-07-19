"""Manual spectroscopy -> image assignment overrides.

When the automatic assignment (`controller._assign_spectros_to_images`)
picks the wrong image for a spectrum, the user can pin it to a chosen
image. This module owns that workflow: identifying which *original* spec
dicts a user selection refers to, applying/clearing the override, and
re-running assignment afterwards.

Follows the `gui/spectroscopy/*.py` module-function convention - every
public function takes the `SXMGridViewer` as its first argument
(`viewer`). Extracted from `main_window.py` as part of the god-class
reduction (see `docs/refactor/GOD_CLASS_PLAN.md`).

**The subtle part is `resolve_override_targets`.** A selection can carry
*copies* of spec dicts (from a dialog, a browser row, a popup) rather
than the originals in `viewer.spectros`. Writing the override onto a copy
silently does nothing - the next re-assignment reads the originals. So
selections are matched back to originals by identity first, then by a
content signature.
"""
from __future__ import annotations

import os
from collections import defaultdict
from pathlib import Path

from ..._shared import QtWidgets, log_status


def override_signature(spec):
    """Stable identity for a spec, used to match copies back to originals.

    Deliberately *not* just the file path: one `.dat` file can hold many
    spectra, and a grid file thousands, so path alone collides. Includes
    matrix index, rounded coordinates, channel and source. Paths are
    case-folded on Windows because the same spec can arrive with
    differently-cased paths from different code paths.
    """
    if not spec:
        return None
    path = str(spec.get("path") or "")
    if path:
        try:
            path = str(Path(path))
        except (TypeError, ValueError):
            pass
        if os.name == "nt":
            path = path.lower()
    try:
        x_val = round(float(spec.get("x")), 6) if spec.get("x") is not None else None
    except (TypeError, ValueError):
        x_val = spec.get("x")
    try:
        y_val = round(float(spec.get("y")), 6) if spec.get("y") is not None else None
    except (TypeError, ValueError):
        y_val = spec.get("y")
    return (
        path,
        spec.get("matrix_index"),
        x_val,
        y_val,
        str(spec.get("channel_name") or "").strip(),
        str(spec.get("source") or "").strip(),
        spec.get("order_idx"),
    )


def resolve_override_targets(viewer, specs):
    """Map a user selection back to the original spec dicts.

    Returns the entries from ``viewer.spectros`` that ``specs`` refers to,
    de-duplicated and order-preserving. Identity match first (the common
    case), then signature match for copies.
    """
    originals = list(getattr(viewer, "spectros", []) or [])
    if not originals:
        return []
    original_ids = {id(spec) for spec in originals}
    by_signature = defaultdict(list)
    for original in originals:
        signature = override_signature(original)
        if signature is not None:
            by_signature[signature].append(original)

    resolved, seen_ids = [], set()
    for spec in list(specs or []):
        if id(spec) in original_ids:
            candidates = [spec]
        else:
            candidates = list(by_signature.get(override_signature(spec)) or [])
        for candidate in candidates:
            if id(candidate) in seen_ids:
                continue
            seen_ids.add(id(candidate))
            resolved.append(candidate)
    return resolved


def current_target_image_key(viewer):
    """Which image a "assign to current image" action should target.

    Prefers the previewed image, then the thumbnail selection; returns ""
    when neither corresponds to a loaded image.
    """
    candidates = []
    try:
        if viewer.last_preview:
            candidates.append(str(viewer.last_preview[0]))
    except (AttributeError, IndexError, TypeError):
        pass
    selected = str(getattr(viewer, "selected_file_for_thumbs", "") or "").strip()
    if selected:
        candidates.append(selected)
    image_keys = {
        str(img.get("path"))
        for img in list(getattr(viewer, "image_meta", []) or [])
        if img and img.get("path")
    }
    for key in candidates:
        if key and key in image_keys:
            return key
    return ""


def current_focus_spec(viewer):
    """The spec a context action should act on, by decreasing priority:
    highlighted, last clicked, browser selection, multi-selection."""
    for candidate in (getattr(viewer, "_highlighted_spec", None),
                      getattr(viewer, "_last_clicked_spec", None)):
        if candidate is not None:
            return candidate
    spectro_list = getattr(viewer, "spectro_list", None)
    current_item = None
    if spectro_list is not None:
        try:
            current_item = spectro_list.currentItem()
        except (AttributeError, RuntimeError):
            current_item = None
    if current_item is not None:
        from ..._shared import QtCore
        try:
            payload = current_item.data(0, QtCore.Qt.UserRole)
        except (AttributeError, RuntimeError):
            payload = None
        if isinstance(payload, dict) and str(payload.get("kind") or "") in {"site", "spec"}:
            spec = payload.get("spec")
            if spec is not None:
                return spec
    selected = list(getattr(viewer, "_multi_spec_selection", []) or [])
    return selected[0] if selected else None


def refresh_after_override(viewer, focus_specs=None):
    """Re-run assignment and refresh every view that shows assignments."""
    viewer._assign_spectros_to_images()
    viewer.matrix_spectros = [spec for spec in viewer.spectros
                              if spec.get("matrix_index") is not None]
    viewer._update_spectro_stats_label()
    try:
        viewer.populate_thumbnails_for_channel(viewer.channel_dropdown.currentIndex())
    except Exception:
        pass
    try:
        if viewer.last_preview:
            viewer.show_file_channel(viewer.last_preview[0], viewer.last_preview[1])
    except Exception:
        pass
    if getattr(viewer, "spectro_dock", None):
        try:
            viewer._filter_spectro_browser()
        except Exception:
            pass
    focus_entries = resolve_override_targets(viewer, focus_specs or [])
    if focus_entries:
        try:
            viewer._highlight_spectrum_entry(focus_entries[0])
        except Exception:
            pass


def apply_override(viewer, specs, image_key=""):
    """Pin ``specs`` to ``image_key`` (default: the current image)."""
    targets = resolve_override_targets(viewer, specs)
    if not targets:
        return
    image_key = str(image_key or current_target_image_key(viewer) or "").strip()
    if not image_key:
        QtWidgets.QMessageBox.information(
            viewer, "Spectroscopy",
            "Open or select an image first, then assign the spectroscopy to it.")
        return
    changed = 0
    for spec in targets:
        if str(spec.get("assignment_override_image_key") or "") == image_key:
            continue
        spec["assignment_override_image_key"] = image_key
        changed += 1
    if changed <= 0:
        return
    log_status(f"[Spectro] Manual assignment override: {changed} "
               f"spectrum/s -> {Path(image_key).name}")
    refresh_after_override(viewer, focus_specs=targets)


def clear_override(viewer, specs):
    """Drop the manual override, restoring automatic assignment."""
    targets = resolve_override_targets(viewer, specs)
    if not targets:
        return
    changed = 0
    for spec in targets:
        if spec.pop("assignment_override_image_key", None) not in (None, ""):
            changed += 1
    if changed <= 0:
        return
    log_status(f"[Spectro] Cleared manual assignment override on {changed} spectrum/s")
    refresh_after_override(viewer, focus_specs=targets)


__all__ = [
    "override_signature",
    "resolve_override_targets",
    "current_target_image_key",
    "current_focus_spec",
    "refresh_after_override",
    "apply_override",
    "clear_override",
]
