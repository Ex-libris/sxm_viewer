"""Recent folders & sessions history and menus for the main window.

Extracted from SXMGridViewer. ``recent_dirs`` and ``recent_session_paths``
remain attributes of the viewer (other modules read them directly); this
controller manages those lists, their menus, and config persistence,
reaching window state via ``self.viewer``.
"""
from __future__ import annotations

from pathlib import Path

from ...config import save_config

RECENT_FOLDER_LIMIT = 30
RECENT_SESSION_LIMIT = 30


class RecentFilesController:
    """Manages recent-folder and recent-session history and menus."""

    def __init__(self, viewer):
        self.viewer = viewer

    def _refresh_recent_dirs_menu(self):
        menu = getattr(self.viewer, "open_recent_menu", None)
        if menu is None:
            return
        menu.clear()
        recents = getattr(self.viewer, "recent_dirs", [])
        if not recents:
            act = menu.addAction("No recent folders")
            act.setEnabled(False)
            return
        for path in recents:
            act = menu.addAction(path)
            act.setToolTip(path)
            act.triggered.connect(lambda checked=False, p=path: self.viewer.load_folder(Path(p)))
        menu.addSeparator()
        clear_act = menu.addAction("Clear recent folders")
        clear_act.triggered.connect(self._clear_recent_dirs)

    def _record_recent_dir(self, folder: Path):
        folder_path = Path(folder)
        folder_str = str(folder_path)
        recents = []
        for p in getattr(self.viewer, "recent_dirs", []):
            if not p:
                continue
            try:
                if Path(p).resolve() == folder_path.resolve():
                    continue
            except Exception:
                if p == folder_str:
                    continue
            recents.append(p)
        recents.insert(0, folder_str)
        self.viewer.recent_dirs = recents[:RECENT_FOLDER_LIMIT]
        self.viewer.config["recent_dirs"] = self.viewer.recent_dirs
        save_config(self.viewer.config)
        self._refresh_recent_dirs_menu()

    def _clear_recent_dirs(self):
        self.viewer.recent_dirs = []
        self.viewer.config["recent_dirs"] = []
        save_config(self.viewer.config)
        self._refresh_recent_dirs_menu()

    def _refresh_recent_session_dirs_menu(self):
        menus = [
            getattr(self.viewer, "load_session_recent_menu", None),
            getattr(self.viewer, "toolbar_load_session_menu", None),
        ]
        for menu in menus:
            if menu is None:
                continue
            menu.clear()
            recents = getattr(self.viewer, "recent_session_paths", [])
            if not recents:
                act = menu.addAction("No recent sessions")
                act.setEnabled(False)
                continue
            for path in recents:
                try:
                    session_path = Path(path)
                except Exception:
                    session_path = Path(str(path))
                act = menu.addAction(str(session_path))
                act.setToolTip(str(session_path))
                act.triggered.connect(
                    lambda checked=False, p=str(session_path): self.on_load_recent_session(Path(p))
                )
            menu.addSeparator()
            clear_act = menu.addAction("Clear recent sessions")
            clear_act.triggered.connect(self._clear_recent_session_dirs)

    def _normalize_recent_session_history(self, persist=False):
        recents = []
        has_session_file = False
        for p in getattr(self.viewer, "recent_session_paths", []):
            if not p:
                continue
            try:
                candidate = Path(p)
            except Exception:
                continue
            if str(candidate) not in recents:
                recents.append(str(candidate))
            try:
                if candidate.suffix.lower() == ".json":
                    has_session_file = True
            except Exception:
                pass
        if has_session_file:
            recents = [p for p in recents if Path(p).suffix.lower() == ".json"]
        recents = recents[:RECENT_SESSION_LIMIT]
        changed = recents != list(getattr(self.viewer, "recent_session_paths", []))
        self.viewer.recent_session_paths = recents
        if persist and changed:
            self.viewer.config["recent_session_paths"] = self.viewer.recent_session_paths
            self.viewer.config.pop("recent_session_dirs", None)
            save_config(self.viewer.config)
        return changed

    def _record_recent_session(self, session_path: Path):
        session_path = Path(session_path)
        session_str = str(session_path)
        recents = []
        for p in getattr(self.viewer, "recent_session_paths", []):
            if not p:
                continue
            try:
                if Path(p).resolve() == session_path.resolve():
                    continue
            except Exception:
                if p == session_str:
                    continue
            recents.append(p)
        recents.insert(0, session_str)
        self.viewer.recent_session_paths = recents[:RECENT_SESSION_LIMIT]
        self._normalize_recent_session_history(persist=False)
        self.viewer.config["recent_session_paths"] = self.viewer.recent_session_paths
        self.viewer.config.pop("recent_session_dirs", None)
        save_config(self.viewer.config)
        self._refresh_recent_session_dirs_menu()

    def _clear_recent_session_dirs(self):
        self.viewer.recent_session_paths = []
        self.viewer.config["recent_session_paths"] = []
        self.viewer.config.pop("recent_session_dirs", None)
        save_config(self.viewer.config)
        self._refresh_recent_session_dirs_menu()

    def _resolve_recent_session_target(self, session_path: Path):
        session_path = Path(session_path)
        if session_path.is_file():
            return session_path
        if not session_path.is_dir():
            return None
        preferred = session_path / "sxm_session.json"
        if preferred.exists():
            return preferred
        try:
            json_files = sorted(
                (p for p in session_path.glob("*.json") if p.is_file()),
                key=lambda p: p.stat().st_mtime,
                reverse=True,
            )
        except Exception:
            json_files = []
        if json_files:
            return json_files[0]
        return None

    def on_load_recent_session(self, session_path: Path):
        session_path = Path(session_path)
        resolved = self._resolve_recent_session_target(session_path)
        if resolved is not None:
            self.viewer.session_controller.load_session(session_path=resolved)
            return
        if session_path.exists() and session_path.is_dir():
            self.viewer.session_controller.load_session(start_dir=session_path)
            return
        self.viewer.session_controller.load_session(session_path=session_path)
