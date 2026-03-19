"""Shared helpers for plot font-family selection and persistence."""
from __future__ import annotations

from functools import lru_cache

from .._shared import QtGui, QtWidgets, matplotlib


_DEFAULT_FONT_FAMILY = "sans-serif"


@lru_cache(maxsize=1)
def available_font_families() -> tuple[str, ...]:
    db = QtGui.QFontDatabase()
    try:
        return tuple(db.families())
    except Exception:
        return ()


def normalize_font_family(family: str | None, fallback: str = _DEFAULT_FONT_FAMILY) -> str:
    fam = str(family or "").strip()
    if not fam:
        return fallback
    families = available_font_families()
    if families and fam not in families and fam.lower() != fallback.lower():
        return fallback
    return fam


def set_matplotlib_font_family(family: str | None) -> str:
    fam = normalize_font_family(family)
    # Keep the font choice global so every plot surface stays visually aligned.
    matplotlib.rcParams["font.family"] = [fam]
    return fam


def choose_font_family(parent, current_family: str | None = None, *, title: str = "Choose plot font") -> str | None:
    base = QtGui.QFont(normalize_font_family(current_family))
    try:
        base.setPointSize(10)
    except Exception:
        pass
    font, ok = QtWidgets.QFontDialog.getFont(base, parent, title)
    if not ok:
        return None
    family = str(font.family() or "").strip()
    return family or None


def add_font_menu_action(menu, parent, current_family: str | None, apply_callback):
    """Add a small typography submenu with a font picker."""
    font_menu = menu.addMenu("Typography")
    current = normalize_font_family(current_family)
    current_act = font_menu.addAction(f"Current: {current}")
    current_act.setEnabled(False)
    choose_act = font_menu.addAction("Choose font...")
    reset_act = font_menu.addAction("Use default")

    def _choose():
        family = choose_font_family(parent, current_family=current, title="Choose plot font")
        if family and callable(apply_callback):
            apply_callback(family)

    def _reset():
        if callable(apply_callback):
            apply_callback(_DEFAULT_FONT_FAMILY)

    choose_act.triggered.connect(_choose)
    reset_act.triggered.connect(_reset)
    return font_menu
