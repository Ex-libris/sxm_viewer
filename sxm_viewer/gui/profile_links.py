"""Shared helpers for live-linked profile dialogs and source canvases."""
from __future__ import annotations

import copy

_PROFILE_DIALOGS = {}
_PROFILE_CANVASES = {}


def _profile_source_id(canvas):
    source_id = str(getattr(canvas, "_profile_live_source_id", "") or "").strip()
    if not source_id:
        source_id = f"canvas-{id(canvas):x}"
        try:
            canvas._profile_live_source_id = source_id
        except Exception:
            pass
    _PROFILE_CANVASES[source_id] = canvas
    return source_id


def register_profile_canvas(canvas):
    if canvas is None:
        return ""
    try:
        return _profile_source_id(canvas)
    except Exception:
        return ""


def unregister_profile_canvas(canvas):
    if canvas is None:
        return
    source_id = str(getattr(canvas, "_profile_live_source_id", "") or "").strip()
    if source_id and _PROFILE_CANVASES.get(source_id) is canvas:
        _PROFILE_CANVASES.pop(source_id, None)


def register_profile_dialog(dialog):
    if dialog is None:
        return
    _PROFILE_DIALOGS[id(dialog)] = dialog


def unregister_profile_dialog(dialog):
    if dialog is None:
        return
    _PROFILE_DIALOGS.pop(id(dialog), None)


def profile_ref_key(profile_ref):
    if not isinstance(profile_ref, dict):
        return None
    source_id = str(profile_ref.get("source_id") or "").strip()
    kind = str(profile_ref.get("kind") or "").strip().lower()
    if not source_id or kind not in {"active", "saved"}:
        return None
    if kind == "active":
        return (source_id, "active", "")
    profile_id = str(
        profile_ref.get("profile_id")
        or profile_ref.get("overlay_id")
        or ""
    ).strip()
    if not profile_id:
        return None
    return (source_id, "saved", profile_id)


def _datasets_by_ref(active_profile, saved_profiles):
    mapping = {}
    for dataset in [active_profile] + list(saved_profiles or []):
        if not isinstance(dataset, dict):
            continue
        key = profile_ref_key(dataset.get("live_profile_ref"))
        if key is None:
            continue
        mapping[key] = copy.deepcopy(dataset)
    return mapping


def notify_profile_source_changed(canvas, active_profile, saved_profiles):
    source_id = register_profile_canvas(canvas)
    if not source_id:
        return
    mapping = _datasets_by_ref(active_profile, saved_profiles)
    if not mapping:
        return
    stale = []
    for key, dialog in list(_PROFILE_DIALOGS.items()):
        if dialog is None:
            stale.append(key)
            continue
        try:
            if hasattr(dialog, "refresh_linked_profiles"):
                dialog.refresh_linked_profiles(mapping, source_id=source_id)
        except RuntimeError:
            stale.append(key)
        except Exception:
            continue
    for key in stale:
        _PROFILE_DIALOGS.pop(key, None)


def apply_live_profile_style(profile_ref, **changes):
    key = profile_ref_key(profile_ref)
    if key is None:
        return False
    canvas = _PROFILE_CANVASES.get(key[0])
    if canvas is None or not hasattr(canvas, "set_profile_style"):
        return False
    try:
        return bool(canvas.set_profile_style(profile_ref=profile_ref, **changes))
    except RuntimeError:
        if _PROFILE_CANVASES.get(key[0]) is canvas:
            _PROFILE_CANVASES.pop(key[0], None)
        return False
    except Exception:
        return False
