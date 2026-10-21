"""Constant-height (CH) height label: absolute z_abs or Δz to a reference.

Module-function convention (see ``gui/viewer/*.py``): every function takes
the ``SXMGridViewer`` as ``viewer``. State lives on the viewer:

- ``viewer.ch_z_label_mode`` - ``"abs"`` (``CH  z_abs 1.234 nm``) or
  ``"rel"`` (``CH  Δz = +20 pm`` relative to the reference image).
- ``viewer.ch_z_reference`` - ``{"file_key", "z_abs_nm"}`` or None. The
  height is stored with the key so the label keeps working when the
  reference file is not loaded.
- ``viewer.ch_label_in_publication`` - also draw the CH line (only that,
  no CC parameters) in the Publication display preset.

Canvases never read these directly: they get callbacks
(``MultiPreviewCanvas.set_ch_height_label_callbacks``) and work out the
label text at draw time, so changing the reference only needs a redraw.
"""
from __future__ import annotations

from pathlib import Path

from ..._shared import QtWidgets, log_status
from ...config import save_config
from ...utils.units import format_ch_height_label
from .preview import _build_acquisition_overlay_info

LABEL_MODES = ("abs", "rel")


def init_state(viewer):
    config = viewer.config
    mode = str(config.get("ch_z_label_mode", "abs") or "abs")
    viewer.ch_z_label_mode = mode if mode in LABEL_MODES else "abs"
    viewer.ch_z_reference = _normalize_reference(config.get("ch_z_reference"))
    viewer.ch_label_in_publication = bool(config.get("ch_label_in_publication", False))
    style = config.get("ch_label_style")
    viewer.ch_label_style = dict(style) if isinstance(style, dict) else {}


def _normalize_reference(ref):
    if not isinstance(ref, dict) or not ref.get("file_key"):
        return None
    try:
        z_nm = float(ref.get("z_abs_nm"))
    except (TypeError, ValueError):
        return None
    return {"file_key": str(ref["file_key"]), "z_abs_nm": z_nm}


def ch_abs_z_nm(viewer, file_key):
    """Raw CH tip height (nm) of a loaded image, or None if not CH / unknown."""
    if not file_key:
        return None
    header, fds = (getattr(viewer, "headers", {}) or {}).get(str(file_key), (None, None))
    if not header:
        return None
    try:
        info = _build_acquisition_overlay_info(viewer, Path(str(file_key)), header, fds or [], 0)
    except Exception:
        return None
    if info.get("mode") != "CH":
        return None
    return info.get("z_abs_nm")


def reference_nm(viewer, file_key=None):
    """Height Δz labels are relative to, or None when labels show z_abs."""
    if getattr(viewer, "ch_z_label_mode", "abs") != "rel":
        return None
    ref = getattr(viewer, "ch_z_reference", None)
    if not ref:
        return None
    # Prefer the live height of a loaded reference, so a corrected CH tag
    # also corrects the reference; the stored value covers unloaded files.
    live = ch_abs_z_nm(viewer, ref["file_key"])
    if live is not None and live != ref.get("z_abs_nm"):
        ref["z_abs_nm"] = float(live)
        _persist(viewer)
    return ref.get("z_abs_nm")


def is_reference(viewer, file_key):
    ref = getattr(viewer, "ch_z_reference", None)
    return bool(ref and file_key and str(file_key) == ref.get("file_key"))


def label_text(viewer, file_key, z_abs_nm):
    """CH label for one image under the current mode/reference."""
    ref_nm = reference_nm(viewer, file_key)
    return format_ch_height_label(
        z_abs_nm, ref_nm, is_reference=ref_nm is not None and is_reference(viewer, file_key))


def _persist(viewer):
    viewer.config["ch_z_label_mode"] = viewer.ch_z_label_mode
    viewer.config["ch_z_reference"] = dict(viewer.ch_z_reference) if viewer.ch_z_reference else None
    viewer.config["ch_label_in_publication"] = bool(viewer.ch_label_in_publication)
    viewer.config["ch_label_style"] = dict(getattr(viewer, "ch_label_style", {}) or {})
    save_config(viewer.config)


def state_snapshot(viewer):
    ref = getattr(viewer, "ch_z_reference", None)
    return (getattr(viewer, "ch_z_label_mode", "abs"), dict(ref) if ref else None)


def restore_state(viewer, mode, reference, persist=True, refresh=True):
    viewer.ch_z_label_mode = mode if mode in LABEL_MODES else "abs"
    viewer.ch_z_reference = _normalize_reference(reference)
    if viewer.ch_z_reference is None:
        viewer.ch_z_label_mode = "abs"
    if persist:
        _persist(viewer)
    if refresh:
        refresh_labels(viewer)


def _push_undo(viewer):
    stack = getattr(viewer, "_adjustment_undo_stack", None)
    if stack is not None:
        stack.append(("chref",) + state_snapshot(viewer))


def set_reference(viewer, file_key):
    """Make ``file_key`` the Δz = 0 image and switch labels to Δz."""
    z_nm = ch_abs_z_nm(viewer, file_key)
    if z_nm is None:
        log_status(f"[CH height] {Path(str(file_key)).name}: not a constant-height image with a known z_abs")
        return False
    _push_undo(viewer)
    restore_state(viewer, "rel", {"file_key": str(file_key), "z_abs_nm": float(z_nm)})
    msg = f"[CH height] Δz reference: {Path(str(file_key)).name} (z_abs {z_nm:.3f} nm)"
    if not _any_label_visible(viewer):
        msg += " - turn on Display > Show acquisition overlay to see the labels"
    log_status(msg + " (Ctrl+Z undoes)")
    return True


def clear_reference(viewer):
    if not getattr(viewer, "ch_z_reference", None):
        return False
    _push_undo(viewer)
    restore_state(viewer, "abs", None)
    log_status("[CH height] Δz reference cleared - labels show z_abs (Ctrl+Z undoes)")
    return True


def set_label_mode(viewer, mode):
    mode = mode if mode in LABEL_MODES else "abs"
    if mode == "rel" and not getattr(viewer, "ch_z_reference", None):
        log_status("[CH height] Set a Δz reference image first (right-click a CH image)")
        return False
    if mode == getattr(viewer, "ch_z_label_mode", "abs"):
        return False
    _push_undo(viewer)
    restore_state(viewer, mode, viewer.ch_z_reference)
    return True


def set_label_in_publication(viewer, on):
    viewer.ch_label_in_publication = bool(on)
    _persist(viewer)
    refresh_labels(viewer)


def _canvases(viewer):
    canvases = [getattr(viewer, "preview_canvas", None)] + list(getattr(viewer, "_popup_canvases", []) or [])
    return [c for c in canvases if c is not None]


def _any_label_visible(viewer):
    for canvas in _canvases(viewer):
        if getattr(canvas, "_show_acquisition_overlay", False):
            return True
        if getattr(canvas, "_publication_mode", False) and viewer.ch_label_in_publication:
            return True
    return False


def refresh_labels(viewer):
    """Redraw canvases showing a CH label (text is resolved at draw time)."""
    for canvas in _canvases(viewer):
        if not getattr(canvas, "views", None):
            continue
        try:
            canvas._redraw()
        except Exception:
            pass


def install_canvas_callbacks(viewer, canvas):
    """Wire a preview/pop-up canvas to the shared CH-label state."""
    if not hasattr(canvas, "set_ch_height_label_callbacks"):
        return
    canvas.set_ch_height_label_callbacks(
        text_cb=lambda file_key, z_nm: label_text(viewer, file_key, z_nm),
        publication_cb=lambda: bool(getattr(viewer, "ch_label_in_publication", False)),
        menu_cb=lambda menu, file_key: populate_menu(viewer, menu, file_key),
        style_cb=lambda: getattr(viewer, "ch_label_style", None) or {},
    )


def open_style_dialog(viewer):
    """Font / colour / outline / box / corner editor, applied live."""
    from ..dialogs.ch_label_style import ChLabelStyleDialog

    dlg = ChLabelStyleDialog(getattr(viewer, "ch_label_style", None) or {}, viewer)

    def _apply(style):
        viewer.ch_label_style = dict(style)
        refresh_labels(viewer)

    dlg.style_changed.connect(_apply)
    if not _any_label_visible(viewer):
        log_status("[CH height] No CH label is visible - turn on Display > Show acquisition overlay to preview")
    dlg.exec_()
    _persist(viewer)


def populate_menu(viewer, menu, file_key=None):
    """Shared "CH height label" submenu: Display menu, thumbnail and
    preview right-click."""
    ref = getattr(viewer, "ch_z_reference", None)
    if file_key:
        name = Path(str(file_key)).name
        z_nm = ch_abs_z_nm(viewer, file_key)
        set_act = menu.addAction(f"Set as Δz reference (Δz = 0): {name}")
        set_act.setCheckable(True)
        set_act.setChecked(is_reference(viewer, file_key))
        set_act.setEnabled(z_nm is not None)
        if z_nm is None:
            set_act.setToolTip("Only constant-height images with a known z_abs can be the reference")
        set_act.triggered.connect(lambda _=False, k=str(file_key): set_reference(viewer, k))
        menu.addSeparator()
    group = QtWidgets.QActionGroup(menu)
    abs_act = menu.addAction("Show absolute height (z_abs)")
    rel_text = "Show Δz relative to reference"
    if ref:
        rel_text += f" ({Path(ref['file_key']).name}, z_abs {ref['z_abs_nm']:.3f} nm)"
    rel_act = menu.addAction(rel_text)
    mode = getattr(viewer, "ch_z_label_mode", "abs")
    for act, act_mode in ((abs_act, "abs"), (rel_act, "rel")):
        act.setCheckable(True)
        act.setChecked(mode == act_mode)
        group.addAction(act)
        act.triggered.connect(lambda _=False, m=act_mode: set_label_mode(viewer, m))
    rel_act.setEnabled(bool(ref))
    rel_act.setToolTip("Δz > 0: tip further from the surface than at the reference image")
    clear_act = menu.addAction("Clear Δz reference")
    clear_act.setEnabled(bool(ref))
    clear_act.triggered.connect(lambda _=False: clear_reference(viewer))
    menu.addSeparator()
    pub_act = menu.addAction("Show CH height label in Publication preset")
    pub_act.setCheckable(True)
    pub_act.setChecked(bool(getattr(viewer, "ch_label_in_publication", False)))
    pub_act.setToolTip("Publication otherwise hides the acquisition overlay; this keeps just the CH height line")
    pub_act.triggered.connect(lambda checked: set_label_in_publication(viewer, checked))
    style_act = menu.addAction("Label style (font, colours, position)...")
    style_act.triggered.connect(lambda _=False: open_style_dialog(viewer))
