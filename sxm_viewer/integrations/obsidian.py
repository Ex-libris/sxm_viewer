"""Qt-free helpers for exporting images into Obsidian notes.

Obsidian keeps everything we need in plain files, so no plugin or running
app connection is required:

- the vault registry: ``obsidian.json`` in Obsidian's per-user config folder
  (``%APPDATA%/obsidian`` on Windows) - ``{"vaults": {id: {path, ts, open}}}``;
- the open note: ``<vault>/.obsidian/workspace.json`` - the ``active`` leaf id
  and each leaf's ``state.state.file`` (vault-relative), plus
  ``lastOpenFiles`` (most recent first);
- link/attachment preferences: ``<vault>/.obsidian/app.json`` -
  ``attachmentFolderPath`` ("/" or "" = vault root, "./x" = relative to the
  note) and ``useMarkdownLinks`` (default False = ``![[wikilinks]]``).

Obsidian watches its vault, so appending to a note that is open in the
editor updates it live.
"""
from __future__ import annotations

import json
import os
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Optional
from urllib.parse import quote


@dataclass(frozen=True)
class Vault:
    path: Path
    name: str
    is_open: bool = False
    ts: int = 0


def obsidian_config_dir() -> Path:
    if sys.platform.startswith("win"):
        base = os.environ.get("APPDATA") or str(Path.home() / "AppData" / "Roaming")
        return Path(base) / "obsidian"
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / "obsidian"
    base = os.environ.get("XDG_CONFIG_HOME") or str(Path.home() / ".config")
    return Path(base) / "obsidian"


def _read_json(path: Path) -> dict:
    try:
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def list_vaults(config_dir: Optional[Path] = None) -> List[Vault]:
    """Known vaults that still exist on disk: open ones first, then most
    recently used."""
    data = _read_json(Path(config_dir or obsidian_config_dir()) / "obsidian.json")
    vaults = []
    for entry in (data.get("vaults") or {}).values():
        if not isinstance(entry, dict) or not entry.get("path"):
            continue
        path = Path(str(entry["path"]))
        if not path.is_dir():
            continue
        try:
            ts = int(entry.get("ts") or 0)
        except Exception:
            ts = 0
        vaults.append(Vault(path=path, name=path.name, is_open=bool(entry.get("open")), ts=ts))
    vaults.sort(key=lambda v: (not v.is_open, -v.ts))
    return vaults


def _workspace(vault: Path) -> dict:
    return _read_json(Path(vault) / ".obsidian" / "workspace.json")


def _iter_leaves(node):
    if isinstance(node, dict):
        if node.get("type") == "leaf":
            yield node
        for child in node.get("children") or []:
            yield from _iter_leaves(child)


def _leaf_markdown_file(leaf) -> Optional[str]:
    state = leaf.get("state") or {}
    if state.get("type") != "markdown":
        return None
    rel = (state.get("state") or {}).get("file")
    return str(rel) if rel and str(rel).lower().endswith(".md") else None


def active_note(vault: Path) -> Optional[Path]:
    """Absolute path of the note in the vault's active editor tab, else the
    most recently opened note; None if nothing suitable exists."""
    vault = Path(vault)
    ws = _workspace(vault)
    active_id = ws.get("active")
    for leaf in _iter_leaves(ws.get("main")):
        if leaf.get("id") == active_id:
            rel = _leaf_markdown_file(leaf)
            if rel and (vault / rel).is_file():
                return vault / rel
    for rel in ws.get("lastOpenFiles") or []:
        if str(rel).lower().endswith(".md") and (vault / rel).is_file():
            return vault / rel
    return None


def recent_notes(vault: Path, limit: int = 25) -> List[Path]:
    """Notes open in tabs (active first) followed by recently opened ones."""
    vault = Path(vault)
    ws = _workspace(vault)
    rels: List[str] = []
    first = active_note(vault)
    if first is not None:
        rels.append(first.relative_to(vault).as_posix())
    for leaf in _iter_leaves(ws.get("main")):
        rel = _leaf_markdown_file(leaf)
        if rel:
            rels.append(rel)
    rels.extend(str(r) for r in ws.get("lastOpenFiles") or [] if str(r).lower().endswith(".md"))
    out, seen = [], set()
    for rel in rels:
        if rel in seen:
            continue
        seen.add(rel)
        if (vault / rel).is_file():
            out.append(vault / rel)
        if len(out) >= limit:
            break
    return out


def vault_for_path(path: Path, vaults: Optional[Iterable[Vault]] = None) -> Optional[Path]:
    """The vault containing ``path`` (a ``.obsidian`` folder above it)."""
    path = Path(path).resolve()
    for vault in vaults or []:
        try:
            path.relative_to(Path(vault.path).resolve())
            return Path(vault.path)
        except Exception:
            continue
    for parent in [path] + list(path.parents):
        if (parent / ".obsidian").is_dir():
            return parent
    return None


def app_settings(vault: Path) -> dict:
    return _read_json(Path(vault) / ".obsidian" / "app.json")


def default_attachment_setting(vault: Path) -> str:
    return str(app_settings(vault).get("attachmentFolderPath") or "/")


def attachment_dir(vault: Path, note: Path, setting: Optional[str] = None) -> Path:
    """Folder new images go into, following Obsidian's own rules for
    ``attachmentFolderPath`` (``setting`` overrides the vault's value)."""
    vault, note = Path(vault), Path(note)
    raw = (setting if setting is not None and str(setting).strip() else default_attachment_setting(vault))
    raw = str(raw).strip().replace("\\", "/")
    if raw in ("", "/"):
        return vault
    if raw == "." or raw.startswith("./"):
        sub = raw[2:] if raw.startswith("./") else ""
        return note.parent / sub if sub else note.parent
    return vault / raw.strip("/")


def embed_line(vault: Path, note: Path, image: Path, *, width: int = 0,
               markdown_links: Optional[bool] = None) -> str:
    """An image embed in the vault's preferred link style."""
    vault, note, image = Path(vault), Path(note), Path(image)
    if markdown_links is None:
        markdown_links = bool(app_settings(vault).get("useMarkdownLinks"))
    if markdown_links:
        rel = os.path.relpath(image, note.parent).replace("\\", "/")
        alt = f"|{int(width)}" if width and width > 0 else ""
        return f"![{image.stem}{alt}]({quote(rel)})"
    rel = image.relative_to(vault).as_posix()
    suffix = f"|{int(width)}" if width and width > 0 else ""
    return f"![[{rel}{suffix}]]"


def resolve_note_path(vault: Path, text: str) -> Path:
    """Vault-relative or absolute note path from user text; adds ``.md``."""
    text = str(text or "").strip().replace("\\", "/")
    if not text:
        raise ValueError("No note given")
    path = Path(text)
    if not path.is_absolute():
        path = Path(vault) / text
    if path.suffix.lower() != ".md":
        path = path.with_name(path.name + ".md")
    return path


def append_to_note(note: Path, text: str) -> None:
    """Append ``text`` as a new block at the end of ``note`` (created, with
    its folder, if missing), keeping the file's existing newline style."""
    note = Path(note)
    note.parent.mkdir(parents=True, exist_ok=True)
    existing = ""
    if note.exists():
        with open(note, "r", encoding="utf-8", newline="") as fh:
            existing = fh.read()
    nl = "\r\n" if "\r\n" in existing else "\n"
    block = text.strip("\n").replace("\r\n", "\n").replace("\n", nl)
    if existing and not existing.endswith(("\n", "\r")):
        prefix = nl + nl
    elif existing and not existing.endswith(nl + nl):
        prefix = nl
    else:
        prefix = ""
    with open(note, "a", encoding="utf-8", newline="") as fh:
        fh.write(prefix + block + nl)
