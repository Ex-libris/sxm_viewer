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


def apply_text_style(text_obj, *, family: str | None = None, bold: bool | None = None, italic: bool | None = None, underline: bool | None = None):
    """Apply a shared text style to a Matplotlib text artist."""
    if text_obj is None:
        return
    try:
        if family:
            text_obj.set_fontfamily(normalize_font_family(family))
    except Exception:
        pass
    try:
        if bold is not None:
            text_obj.set_fontweight("bold" if bold else "normal")
    except Exception:
        pass
    try:
        if italic is not None:
            text_obj.set_fontstyle("italic" if italic else "normal")
    except Exception:
        pass
    try:
        if underline is not None and hasattr(text_obj, "set_underline"):
            text_obj.set_underline(bool(underline))
    except Exception:
        pass


def apply_qfont_style(font, *, family: str | None = None, bold: bool | None = None, italic: bool | None = None, underline: bool | None = None):
    """Apply the same style model to a Qt font."""
    if font is None:
        font = QtGui.QFont()
    try:
        if family:
            font.setFamily(normalize_font_family(family))
    except Exception:
        pass
    try:
        if bold is not None:
            font.setBold(bool(bold))
    except Exception:
        pass
    try:
        if italic is not None:
            font.setItalic(bool(italic))
    except Exception:
        pass
    try:
        if underline is not None:
            font.setUnderline(bool(underline))
    except Exception:
        pass
    return font


def add_font_menu_action(menu, parent, current_family: str | None, apply_callback, *, current_style: dict | None = None, apply_style_callback=None):
    """Add typography controls with family and optional bold/italic/underline toggles."""
    font_menu = menu.addMenu("Typography")
    current = normalize_font_family(current_family)
    current_act = font_menu.addAction(f"Current: {current}")
    current_act.setEnabled(False)
    choose_act = font_menu.addAction("Choose font...")
    reset_act = font_menu.addAction("Use default")
    current_style = current_style or {}

    def _choose():
        family = choose_font_family(parent, current_family=current, title="Choose plot font")
        if family and callable(apply_callback):
            apply_callback(family)

    def _reset():
        if callable(apply_callback):
            apply_callback(_DEFAULT_FONT_FAMILY)

    choose_act.triggered.connect(_choose)
    reset_act.triggered.connect(_reset)

    if callable(apply_style_callback):
        font_menu.addSeparator()
        bold_act = font_menu.addAction("Bold")
        bold_act.setCheckable(True)
        bold_act.setChecked(bool(current_style.get("bold", False)))
        italic_act = font_menu.addAction("Italic")
        italic_act.setCheckable(True)
        italic_act.setChecked(bool(current_style.get("italic", False)))
        underline_act = font_menu.addAction("Underline")
        underline_act.setCheckable(True)
        underline_act.setChecked(bool(current_style.get("underline", False)))
        reset_style_act = font_menu.addAction("Reset typography style")

        def _push_style(**changes):
            try:
                apply_style_callback(**changes)
            except Exception:
                pass

        bold_act.toggled.connect(lambda checked: _push_style(bold=bool(checked)))
        italic_act.toggled.connect(lambda checked: _push_style(italic=bool(checked)))
        underline_act.toggled.connect(lambda checked: _push_style(underline=bool(checked)))
        reset_style_act.triggered.connect(lambda: _push_style(bold=False, italic=False, underline=False))

    return font_menu
