"""Export routes for the right-click menus (see ``gui/menu_kit.py``).

Two scopes, each with its own pinned list / "Again" memory in config
``context_menus.<scope>``:

- ``"thumbnails"`` - ctx is the list of selected file keys
  (``thumbnail_routes``);
- ``"view"`` - ctx is ``{"canvas": MultiPreviewCanvas, "view": dict}`` for
  the main preview and every pop-up. The canvas defines its own built-in
  routes (copy/save/PowerPoint/Inkscape/STP); ``install_canvas_menu_support``
  hands it the viewer-backed prefs and the routes that need the viewer
  (Obsidian).
"""
from __future__ import annotations

import itertools
from pathlib import Path

from ...config import save_config
from ..menu_kit import MenuPrefs, Route, again_folder as _again_folder
from ..ppt_bridge import powerpoint_support_status

THUMB_DEFAULT_PINNED = ("png", "ppt_new")
VIEW_DEFAULT_PINNED = ("copy_png", "ppt_new")


def menu_prefs(viewer, scope):
    defaults = THUMB_DEFAULT_PINNED if scope == "thumbnails" else VIEW_DEFAULT_PINNED
    config = getattr(viewer, "config", None)
    return MenuPrefs(scope, defaults, config=config, save=save_config,
                     legacy="thumb_menu" if scope == "thumbnails" else None)


# ---------------------------------------------------------------------------
# shared helpers
# ---------------------------------------------------------------------------
def _ppt_available(_ctx=None):
    ok, reason = powerpoint_support_status()
    return ok, reason or "PowerPoint is unavailable"


def active_obsidian_note():
    """(note, vault) for the first open vault's active note, or (None, None)."""
    try:
        from ...integrations import obsidian
        vaults = obsidian.list_vaults()
        for vault in [v for v in vaults if v.is_open] or vaults[:1]:
            note = obsidian.active_note(vault.path)
            if note is not None:
                return note, vault.path
    except Exception:
        pass
    return None, None


def _obsidian_active_label(_ctx=None):
    note, _ = active_obsidian_note()
    return f"Obsidian → {note.stem}" if note is not None else "Obsidian: active note"


def _obsidian_active_tip(_ctx=None):
    note, _ = active_obsidian_note()
    return f"Append to the note open in Obsidian:\n{note}" if note is not None else "No note is open in Obsidian"


def _obsidian_active_available(_ctx=None):
    return (active_obsidian_note()[0] is not None), "No note is open in Obsidian"


def _again_obsidian(record, pinned):
    note, vault = record.get("note"), record.get("vault")
    if not note or not vault or not Path(vault).is_dir():
        return None
    if "obsidian_active" in pinned and active_obsidian_note()[0] == Path(note):
        return None  # the pinned "active note" entry already does exactly this
    return f"Obsidian → {Path(note).stem}"


def _obsidian_dest(dest):
    if dest and dest.get("note"):
        return Path(dest["note"]), Path(dest["vault"]) if dest.get("vault") else None
    return None, None


def _obsidian_routes(run_with):
    """``run_with(ctx, note, vault) -> record|None``; note None = dialog."""
    def run_active(ctx, dest):
        note, vault = _obsidian_dest(dest)
        if note is None:
            note, vault = active_obsidian_note()
        return run_with(ctx, note, vault)

    def run_choose(ctx, dest):
        note, vault = _obsidian_dest(dest)
        return run_with(ctx, note, vault)

    return [
        Route("obsidian_active", _obsidian_active_label, run_active, tip=_obsidian_active_tip,
              available=_obsidian_active_available, again=_again_obsidian),
        Route("obsidian", "Obsidian: choose note...", run_choose,
              tip="Pick a vault and note (or create one)", again=_again_obsidian),
    ]


# ---------------------------------------------------------------------------
# thumbnails (ctx = list of file keys)
# ---------------------------------------------------------------------------
def thumbnail_routes(viewer):
    def folder_route(rid, label, tip, fn, kind):
        def run(paths, dest):
            out = fn(list(paths), (dest or {}).get("out_dir"))
            return {"out_dir": out} if out else None
        return Route(rid, label, run, tip=tip, again=_again_folder(kind))

    def obsidian_run(paths, note, vault):
        return viewer.export_to_obsidian(note, vault, list(paths))

    return [
        folder_route("png", "Export PNGs (as displayed)...",
                     "One PNG per image, rendered as the preview shows it (display preset, overlays, colormap)",
                     viewer.on_export_pngs, "PNGs"),
        Route("svg", "Copy as SVG", lambda paths, _d: viewer.copy_selected_as_svg(list(paths)),
              tip="All selected images in one SVG on the clipboard"),
        Route("ppt_new", "Send to PowerPoint",
              lambda paths, _d: viewer._send_thumbnail_targets_to_powerpoint(list(paths), new_slide=True),
              tip="One new slide per image", available=_ppt_available),
        Route("ppt_current", "Send to current PowerPoint slide",
              lambda paths, _d: viewer._send_thumbnail_targets_to_powerpoint(list(paths), new_slide=False),
              tip="All images onto the slide open in PowerPoint", available=_ppt_available),
        *_obsidian_routes(obsidian_run),
        folder_route("xyz", "Export XYZ...", "Plain-text x y z files, one per channel",
                     viewer.on_export_xyz_files, "XYZ"),
        folder_route("stp", "Export WSxM STP...", "WSxM .stp files, one per channel",
                     viewer.on_export_stp_files, "WSxM STP"),
    ]


# ---------------------------------------------------------------------------
# single views (ctx = {"canvas", "view"}): viewer-backed extras
# ---------------------------------------------------------------------------
def _unique(path):
    path = Path(path)
    if not path.exists():
        return path
    for n in itertools.count(2):
        cand = path.with_name(f"{path.stem}_{n}{path.suffix}")
        if not cand.exists():
            return cand


def view_extra_routes(viewer):
    def obsidian_run(ctx, note, vault):
        canvas, view = ctx["canvas"], ctx["view"]
        meta = dict((view or {}).get("meta") or {})
        file_key = str((view or {}).get("path") or meta.get("file_path") or "")
        channel_idx = meta.get("channel_index")

        def render(img_dir):
            out = _unique(Path(img_dir) / canvas.default_export_filename(view, "png"))
            Path(img_dir).mkdir(parents=True, exist_ok=True)
            canvas.save_displayed_png(out)
            return [{"path": out, "file_key": file_key, "channel_idx": channel_idx}], [], False

        return viewer.export_to_obsidian(note, vault, render=render, parent=canvas.window())

    return _obsidian_routes(obsidian_run)


def install_canvas_menu_support(viewer, canvas):
    """Give a preview/pop-up canvas the shared "view" menu prefs and the
    viewer-backed export routes."""
    if not hasattr(canvas, "set_context_menu_support"):
        return
    from ..cmap_menu import CmapPrefs
    canvas.set_context_menu_support(
        prefs_factory=lambda: menu_prefs(viewer, "view"),
        extra_routes=view_extra_routes(viewer),
        cmap_prefs_factory=lambda: CmapPrefs(getattr(viewer, "config", None), save_config),
    )
