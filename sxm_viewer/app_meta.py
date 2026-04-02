"""Application metadata and drop-in icon discovery."""
from __future__ import annotations

import os
import sys
from pathlib import Path

from PyQt5 import QtCore, QtGui, QtWidgets

APP_NAME = "SXM Viewer"
APP_DISPLAY_NAME = APP_NAME
APP_ORGANIZATION_NAME = "OmicronInfinitySXM"
APP_USER_MODEL_ID = "OmicronInfinitySXM.SXMViewer"

_ICON_BASE_NAMES = (
    "app_icon",
    "sxm_viewer_icon",
    "sxmviewer_icon",
)
_ICON_EXTENSIONS_BY_PLATFORM = {
    "win32": (".ico", ".png", ".svg", ".icns"),
    "darwin": (".icns", ".png", ".svg", ".ico"),
    "default": (".png", ".svg", ".ico", ".icns"),
}


def project_root() -> Path:
    return Path(__file__).resolve().parents[1]


def samples_dir() -> Path:
    return project_root() / "samples"


def _icon_candidate_paths(base_dir: Path | None = None):
    directory = Path(base_dir) if base_dir else samples_dir()
    suffixes = _ICON_EXTENSIONS_BY_PLATFORM.get(QtCore.QSysInfo.productType(), None)
    if not suffixes:
        suffixes = _ICON_EXTENSIONS_BY_PLATFORM.get(sys.platform, _ICON_EXTENSIONS_BY_PLATFORM["default"])
    for stem in _ICON_BASE_NAMES:
        for suffix in suffixes:
            yield directory / f"{stem}{suffix}"


def find_app_icon_path(base_dir: Path | None = None) -> Path | None:
    for path in _icon_candidate_paths(base_dir):
        if path.is_file():
            return path
    return None


def load_app_icon(base_dir: Path | None = None) -> tuple[QtGui.QIcon, Path | None]:
    for path in _icon_candidate_paths(base_dir):
        if not path.is_file():
            continue
        icon = QtGui.QIcon(str(path))
        if not icon.isNull():
            return icon, path
    return QtGui.QIcon(), None


def apply_window_icon(window) -> Path | None:
    app = QtWidgets.QApplication.instance()
    if app is None:
        return None
    icon = app.windowIcon()
    if icon.isNull():
        return None
    try:
        window.setWindowIcon(icon)
    except Exception:
        return None
    return find_app_icon_path()


def _set_windows_app_user_model_id() -> None:
    if os.name != "nt":
        return
    try:
        import ctypes

        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_USER_MODEL_ID)
    except Exception:
        return


def configure_application(app: QtWidgets.QApplication, *, icon_dir: Path | None = None) -> Path | None:
    QtCore.QCoreApplication.setApplicationName(APP_NAME)
    QtCore.QCoreApplication.setOrganizationName(APP_ORGANIZATION_NAME)
    if hasattr(app, "setApplicationDisplayName"):
        try:
            app.setApplicationDisplayName(APP_DISPLAY_NAME)
        except Exception:
            pass
    _set_windows_app_user_model_id()
    icon, icon_path = load_app_icon(icon_dir)
    if not icon.isNull():
        app.setWindowIcon(icon)
    return icon_path
