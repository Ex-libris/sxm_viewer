from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path

from .._shared import QtCore, QtGui, QtWidgets


def _coerce_local_path(path) -> Path | None:
    text = str(path or "").strip()
    if not text:
        return None
    try:
        return Path(text).expanduser()
    except Exception:
        return Path(text)


def _show_open_error(parent, title: str, detail: str):
    try:
        QtWidgets.QMessageBox.warning(parent, title, detail)
    except Exception:
        pass


def _launch(args: list[str]) -> bool:
    try:
        subprocess.Popen(args)
        return True
    except Exception:
        return False


def reveal_in_file_manager(path, parent=None) -> bool:
    target = _coerce_local_path(path)
    if target is None:
        _show_open_error(parent, "Source file", "No source file is available for this item.")
        return False
    if not target.exists():
        _show_open_error(parent, "Source file", f"File not found:\n{target}")
        return False
    try:
        if sys.platform.startswith("win"):
            if target.is_file():
                return _launch(["explorer", f"/select,{str(target)}"])
            return _launch(["explorer", str(target)])
        if sys.platform == "darwin":
            if target.is_file():
                return _launch(["open", "-R", str(target)])
            return _launch(["open", str(target)])
        folder = target if target.is_dir() else target.parent
        if shutil.which("xdg-open") and _launch(["xdg-open", str(folder)]):
            return True
        if shutil.which("gio") and _launch(["gio", "open", str(folder)]):
            return True
        return QtGui.QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(str(folder)))
    except Exception as exc:
        _show_open_error(parent, "Source file", f"Unable to open the file manager.\n{exc}")
        return False


def open_in_text_editor(path, parent=None) -> bool:
    target = _coerce_local_path(path)
    if target is None:
        _show_open_error(parent, "Inspect file", "No source file is available for this item.")
        return False
    if not target.exists():
        _show_open_error(parent, "Inspect file", f"File not found:\n{target}")
        return False
    if target.is_dir():
        _show_open_error(parent, "Inspect file", f"Expected a file, got a folder:\n{target}")
        return False
    try:
        if sys.platform.startswith("win"):
            try:
                os.startfile(str(target), "edit")
                return True
            except Exception:
                if _launch(["notepad.exe", str(target)]):
                    return True
        elif sys.platform == "darwin":
            if _launch(["open", "-t", str(target)]):
                return True
        else:
            if shutil.which("xdg-open") and _launch(["xdg-open", str(target)]):
                return True
            if shutil.which("gio") and _launch(["gio", "open", str(target)]):
                return True
        if QtGui.QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(str(target))):
            return True
    except Exception as exc:
        _show_open_error(parent, "Inspect file", f"Unable to open the file in a text editor.\n{exc}")
        return False
    _show_open_error(parent, "Inspect file", f"Unable to open the file in a text editor.\n{target}")
    return False


def add_source_file_menu(menu, path, parent=None, *, title: str = "Source file", include_copy: bool = True):
    source_path = _coerce_local_path(path)
    submenu = menu.addMenu(title)
    reveal_act = submenu.addAction("Show in file manager")
    inspect_act = submenu.addAction("Open in text editor")
    copy_act = submenu.addAction("Copy file path") if include_copy else None
    enabled = source_path is not None
    submenu.setEnabled(enabled)
    reveal_act.setEnabled(enabled)
    inspect_act.setEnabled(enabled)
    if copy_act is not None:
        copy_act.setEnabled(enabled)
    if enabled:
        path_text = str(source_path)
        submenu.setToolTip(path_text)
        reveal_act.setToolTip(path_text)
        inspect_act.setToolTip(path_text)
        if copy_act is not None:
            copy_act.setToolTip(path_text)
        reveal_act.triggered.connect(lambda _checked=False, p=path_text: reveal_in_file_manager(p, parent))
        inspect_act.triggered.connect(lambda _checked=False, p=path_text: open_in_text_editor(p, parent))
        if copy_act is not None:
            copy_act.triggered.connect(lambda _checked=False, p=path_text: QtWidgets.QApplication.clipboard().setText(p))
    return submenu


__all__ = [
    "add_source_file_menu",
    "open_in_text_editor",
    "reveal_in_file_manager",
]
