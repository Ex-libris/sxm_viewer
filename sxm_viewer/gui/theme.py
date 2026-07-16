"""Named UI themes for the SXM Viewer application chrome.

This module centralizes the semantic design tokens and generated Qt
palette/stylesheets for the named UI themes (``light``, ``dark``,
``amber``).

IMPORTANT — scientific data vs. chrome separation
--------------------------------------------------
Everything in this module styles *application chrome only*: window and
panel backgrounds, controls, menus, borders, labels, selection frames.
None of these colors may ever be applied to scientific imagery — data
pixmaps (thumbnails, crop/filter previews, canvas tiles) and matplotlib
image artists/colorbars always render through the user-selected
colormap.  Qt palettes and QSS cannot recolor pixmap contents or
matplotlib artists, and it must stay that way: never add a theme rule
that paints *over* an image widget (no overlays, no opacity washes, no
blend tricks).  Selection/hover treatments for image-bearing widgets are
borders and external frames only.

The single sanctioned exception is the explicit, user-opted "Full amber
imagery" mode: while the Amber theme is active and the user checks that
toggle, renderers recolor image artists through
``sxm_viewer.cmap_registry.effective_cmap`` (a display-time override —
per-file colormap choices and exports keep the true colormaps).  That
override lives in the registry, not here; this module still never
touches data imagery.
"""
from __future__ import annotations

from PyQt5 import QtCore, QtGui, QtWidgets

THEME_LIGHT = "light"
THEME_DARK = "dark"
THEME_AMBER = "amber"
THEMES = (THEME_LIGHT, THEME_DARK, THEME_AMBER)

_THEME_LABELS = {
    THEME_LIGHT: "Light",
    THEME_DARK: "Dark",
    THEME_AMBER: "Amber",
}

# Semantic design tokens for the amber phosphor theme.  The palette aims
# for a professional laboratory-instrument look: near-black warm
# backgrounds, warm amber text, restrained bright-amber accents.
# NOTE: window_bg / amber_muted / amber_primary are duplicated as the
# "gui_amber_theme" colormap stops in sxm_viewer/cmap_registry.py
# (AMBER_CMAP_STOPS) — that module must stay Qt-free and cannot import
# this one. Keep the two in sync.
AMBER = {
    "window_bg": "#100b05",
    "panel_bg": "#171006",
    "panel_bg_raised": "#201609",
    "input_bg": "#0d0904",

    "amber_primary": "#ffb000",
    "amber_bright": "#ffd166",
    "amber_secondary": "#c88700",
    "amber_muted": "#805b18",
    "amber_disabled": "#5a431d",

    "border": "#6f4b0b",
    "border_active": "#ffb000",
    "selection_bg": "#4a3106",
    "hover_bg": "#2b1d08",

    "text_primary": "#ffc24b",
    "text_secondary": "#c99432",
    "text_disabled": "#72582b",

    "error": "#ff6b4a",
    "warning": "#ffb000",
    "success": "#d6b84b",
}

# Readable monospace stack for the instrument-panel chrome.  No bundled
# fonts: every entry falls back gracefully if missing.
AMBER_FONT_STACK = 'Consolas, "Cascadia Mono", "DejaVu Sans Mono", monospace'


def normalize(name) -> str:
    """Coerce any user/config value into a known theme name."""
    name = str(name or "").strip().lower()
    return name if name in THEMES else THEME_LIGHT


# The app has exactly one active theme; the main window registers it here on
# every apply so leaf widgets (dialogs, popups) that have no viewer reference
# can still style their chrome correctly.
_CURRENT_THEME = THEME_LIGHT


def set_current_theme(name: str):
    global _CURRENT_THEME
    _CURRENT_THEME = normalize(name)


def current_theme() -> str:
    return _CURRENT_THEME


def theme_label(name: str) -> str:
    return _THEME_LABELS.get(normalize(name), "Light")


def resolve_theme_name(config) -> str:
    """Resolve the configured theme, honoring the legacy ``dark_mode`` key.

    ``ui_theme`` wins when present; otherwise a pre-existing boolean
    ``dark_mode`` maps onto the light/dark pair so old configs keep
    their look.
    """
    try:
        raw = config.get("ui_theme", None)
    except Exception:
        raw = None
    if raw is not None:
        return normalize(raw)
    try:
        legacy_dark = bool(config.get("dark_mode", False))
    except Exception:
        legacy_dark = False
    return THEME_DARK if legacy_dark else THEME_LIGHT


def is_dark_theme(name) -> bool:
    """Amber is a dark-chrome theme: existing dark-mode branches
    (matplotlib detail chrome, dialogs) should treat it as dark."""
    return normalize(name) in (THEME_DARK, THEME_AMBER)


def is_amber(viewer_or_name) -> bool:
    name = viewer_or_name
    if not isinstance(name, str):
        name = getattr(viewer_or_name, "ui_theme", THEME_LIGHT)
    return normalize(name) == THEME_AMBER


def amber_palette() -> QtGui.QPalette:
    """QPalette generated from the amber tokens (chrome only)."""
    t = AMBER
    p = QtGui.QPalette()
    c = QtGui.QColor

    p.setColor(QtGui.QPalette.Window, c(t["window_bg"]))
    p.setColor(QtGui.QPalette.WindowText, c(t["text_primary"]))
    p.setColor(QtGui.QPalette.Base, c(t["input_bg"]))
    p.setColor(QtGui.QPalette.AlternateBase, c(t["panel_bg"]))
    p.setColor(QtGui.QPalette.Text, c(t["text_primary"]))
    p.setColor(QtGui.QPalette.Button, c(t["panel_bg"]))
    p.setColor(QtGui.QPalette.ButtonText, c(t["text_primary"]))
    p.setColor(QtGui.QPalette.BrightText, c(t["amber_bright"]))
    p.setColor(QtGui.QPalette.Highlight, c(t["selection_bg"]))
    p.setColor(QtGui.QPalette.HighlightedText, c(t["amber_bright"]))
    p.setColor(QtGui.QPalette.Link, c(t["amber_bright"]))
    p.setColor(QtGui.QPalette.LinkVisited, c(t["amber_secondary"]))
    p.setColor(QtGui.QPalette.ToolTipBase, c(t["panel_bg_raised"]))
    p.setColor(QtGui.QPalette.ToolTipText, c(t["text_primary"]))
    try:
        p.setColor(QtGui.QPalette.PlaceholderText, c(t["amber_muted"]))
    except Exception:
        pass
    # Structural roles Fusion uses for frames/bevels.
    p.setColor(QtGui.QPalette.Light, c(t["panel_bg_raised"]))
    p.setColor(QtGui.QPalette.Midlight, c(t["panel_bg_raised"]))
    p.setColor(QtGui.QPalette.Mid, c(t["border"]))
    p.setColor(QtGui.QPalette.Dark, c("#060402"))
    p.setColor(QtGui.QPalette.Shadow, c("#000000"))

    for role in (
        QtGui.QPalette.WindowText,
        QtGui.QPalette.Text,
        QtGui.QPalette.ButtonText,
        QtGui.QPalette.HighlightedText,
    ):
        p.setColor(QtGui.QPalette.Disabled, role, c(t["text_disabled"]))
    p.setColor(QtGui.QPalette.Disabled, QtGui.QPalette.Highlight, c(t["panel_bg_raised"]))
    return p


def amber_app_qss(font_scale: int = 100) -> str:
    """Application-wide stylesheet for the amber theme.

    Styles standard controls only.  Deliberately does NOT paint plain
    ``QWidget`` backgrounds (the palette handles those) so custom-painted
    widgets — including matplotlib canvases and pixmap-bearing labels —
    are never touched.  ``font_scale`` is the user's UI font scale in
    percent (see the main window's font-scale slider).
    """
    t = dict(AMBER)
    t["font"] = AMBER_FONT_STACK
    try:
        scale = max(60, min(200, int(font_scale)))
    except Exception:
        scale = 100
    t["font_pt"] = f"{9.0 * scale / 100.0:.1f}"
    return """
/* ---- chrome typography (monospace instrument panel) ----
   Base 9pt: monospace at the app's default 11pt is wider than Segoe UI
   and clips fixed-width buttons; 9pt keeps every existing control label
   inside its historical hit-target.  Scaled by the UI font-scale %. */
QPushButton, QToolButton, QLabel, QComboBox, QLineEdit, QSpinBox,
QDoubleSpinBox, QCheckBox, QRadioButton, QMenu, QTabBar, QGroupBox,
QHeaderView, QListView, QTreeView, QTableView, QToolTip, QStatusBar,
QProgressBar, QPlainTextEdit, QTextEdit, QMenuBar, QToolBar {{
    font-family: {font};
    font-size: {font_pt}pt;
}}

/* ---- toolbars / panels ---- */
QToolBar {{
    background: {panel_bg};
    border-bottom: 1px solid {border};
    spacing: 4px;
    padding: 2px;
}}
QToolBar::separator {{
    background: {border};
    width: 1px;
    margin: 4px 4px;
}}
QStatusBar {{
    background: {panel_bg};
    color: {text_secondary};
    border-top: 1px solid {border};
}}
QSplitter::handle {{
    background: {panel_bg_raised};
}}

/* ---- buttons: dark surface, thin amber border, bright label ---- */
QPushButton {{
    background-color: {panel_bg};
    color: {text_primary};
    border: 1px solid {border};
    border-radius: 2px;
    padding: 4px 10px;
}}
QPushButton:hover {{
    background-color: {hover_bg};
    border-color: {amber_secondary};
}}
QPushButton:pressed, QPushButton:checked {{
    background-color: {selection_bg};
    border-color: {border_active};
    color: {amber_bright};
}}
QPushButton:focus {{
    border: 1px solid {border_active};
}}
QPushButton:disabled {{
    color: {text_disabled};
    border-color: {amber_disabled};
    background-color: {window_bg};
}}
QPushButton::menu-indicator {{
    subcontrol-origin: padding;
    subcontrol-position: center right;
    right: 4px;
}}

QToolButton {{
    background-color: transparent;
    color: {text_primary};
    border: 1px solid transparent;
    border-radius: 2px;
    padding: 3px 6px;
}}
QToolButton:hover {{
    background-color: {hover_bg};
    border-color: {border};
}}
QToolButton:pressed, QToolButton:checked {{
    background-color: {selection_bg};
    border-color: {border_active};
    color: {amber_bright};
}}
QToolButton:focus {{
    border: 1px solid {border_active};
}}
QToolButton:disabled {{
    color: {text_disabled};
}}

/* ---- inputs: near-black inset, amber caret, clear focus ---- */
QLineEdit, QSpinBox, QDoubleSpinBox, QPlainTextEdit, QTextEdit {{
    background-color: {input_bg};
    color: {text_primary};
    border: 1px solid {border};
    border-radius: 2px;
    padding: 2px 4px;
    selection-background-color: {selection_bg};
    selection-color: {amber_bright};
}}
QLineEdit:focus, QSpinBox:focus, QDoubleSpinBox:focus,
QPlainTextEdit:focus, QTextEdit:focus {{
    border: 1px solid {border_active};
}}
QLineEdit:read-only, QPlainTextEdit:read-only {{
    color: {text_secondary};
    background-color: {panel_bg};
}}
QLineEdit:disabled, QSpinBox:disabled, QDoubleSpinBox:disabled {{
    color: {text_disabled};
    border-color: {amber_disabled};
}}

QComboBox {{
    background-color: {input_bg};
    color: {text_primary};
    border: 1px solid {border};
    border-radius: 2px;
    padding: 2px 6px;
}}
QComboBox:hover {{
    border-color: {amber_secondary};
}}
QComboBox:focus {{
    border: 1px solid {border_active};
}}
QComboBox:disabled {{
    color: {text_disabled};
    border-color: {amber_disabled};
}}
QComboBox::drop-down {{
    border-left: 1px solid {border};
    width: 18px;
}}
/* Opaque popup with unmistakable selection (Windows needs this spelled out). */
QComboBox QAbstractItemView {{
    background-color: {panel_bg_raised};
    color: {text_primary};
    border: 1px solid {border};
    selection-background-color: {selection_bg};
    selection-color: {amber_bright};
    outline: none;
}}

/* ---- menus ---- */
QMenu {{
    background-color: {panel_bg_raised};
    color: {text_primary};
    border: 1px solid {border};
    padding: 2px;
}}
QMenu::item {{
    padding: 4px 22px 4px 22px;
    border-radius: 2px;
}}
QMenu::item:selected {{
    background-color: {selection_bg};
    color: {amber_bright};
}}
QMenu::item:disabled {{
    color: {text_disabled};
}}
QMenu::separator {{
    height: 1px;
    background: {border};
    margin: 3px 6px;
}}
QMenu::indicator {{
    width: 12px;
    height: 12px;
    left: 5px;
}}
QMenuBar {{
    background-color: {panel_bg};
    color: {text_primary};
}}
QMenuBar::item:selected {{
    background-color: {selection_bg};
    color: {amber_bright};
}}

/* ---- checkboxes / radios ---- */
QCheckBox, QRadioButton {{
    color: {text_primary};
    spacing: 6px;
}}
QCheckBox:disabled, QRadioButton:disabled {{
    color: {text_disabled};
}}
QCheckBox::indicator, QRadioButton::indicator {{
    width: 13px;
    height: 13px;
    border: 1px solid {border};
    background-color: {input_bg};
}}
QRadioButton::indicator {{
    border-radius: 7px;
}}
QCheckBox::indicator:checked {{
    background-color: {amber_primary};
    border-color: {amber_primary};
}}
QRadioButton::indicator:checked {{
    background-color: {amber_primary};
    border-color: {amber_primary};
}}
QCheckBox::indicator:hover, QRadioButton::indicator:hover {{
    border-color: {border_active};
}}
QCheckBox:focus, QRadioButton:focus {{
    outline: 1px solid {border_active};
}}

/* ---- tabs ---- */
QTabWidget::pane {{
    border: 1px solid {border};
    top: -1px;
}}
QTabBar::tab {{
    background: {panel_bg};
    color: {text_secondary};
    border: 1px solid {border};
    border-bottom: none;
    padding: 4px 12px;
    margin-right: 2px;
    border-top-left-radius: 2px;
    border-top-right-radius: 2px;
}}
QTabBar::tab:selected {{
    background: {panel_bg_raised};
    color: {amber_bright};
    border-top: 2px solid {amber_primary};
}}
QTabBar::tab:hover:!selected {{
    background: {hover_bg};
    color: {text_primary};
}}

/* ---- group boxes ---- */
QGroupBox {{
    border: 1px solid {border};
    border-radius: 2px;
    margin-top: 10px;
    color: {text_primary};
}}
QGroupBox::title {{
    subcontrol-origin: margin;
    left: 8px;
    padding: 0 4px;
    color: {text_secondary};
}}

/* ---- item views / headers ---- */
QListView, QTreeView, QTableView, QListWidget, QTreeWidget, QTableWidget {{
    background-color: {input_bg};
    alternate-background-color: {panel_bg};
    color: {text_primary};
    border: 1px solid {border};
    selection-background-color: {selection_bg};
    selection-color: {amber_bright};
    outline: none;
}}
QTreeView::item:hover, QListView::item:hover {{
    background-color: {hover_bg};
}}
QHeaderView::section {{
    background-color: {panel_bg_raised};
    color: {text_secondary};
    border: none;
    border-right: 1px solid {border};
    border-bottom: 1px solid {border};
    padding: 3px 6px;
}}

/* ---- scrollbars: narrow but targetable ---- */
QScrollBar:vertical {{
    background: {window_bg};
    width: 12px;
    margin: 0;
}}
QScrollBar::handle:vertical {{
    background: {amber_muted};
    border-radius: 2px;
    min-height: 24px;
    margin: 2px 2px;
}}
QScrollBar::handle:vertical:hover {{
    background: {amber_secondary};
}}
QScrollBar:horizontal {{
    background: {window_bg};
    height: 12px;
    margin: 0;
}}
QScrollBar::handle:horizontal {{
    background: {amber_muted};
    border-radius: 2px;
    min-width: 24px;
    margin: 2px 2px;
}}
QScrollBar::handle:horizontal:hover {{
    background: {amber_secondary};
}}
QScrollBar::add-line, QScrollBar::sub-line {{
    height: 0;
    width: 0;
    background: none;
    border: none;
}}
QScrollBar::add-page, QScrollBar::sub-page {{
    background: none;
}}

/* ---- sliders ---- */
QSlider::groove:horizontal {{
    background: {panel_bg_raised};
    height: 4px;
    border-radius: 2px;
}}
QSlider::handle:horizontal {{
    background: {amber_primary};
    border: 1px solid {border};
    width: 12px;
    margin: -5px 0;
    border-radius: 2px;
}}
QSlider::handle:horizontal:hover {{
    background: {amber_bright};
}}
QSlider::groove:vertical {{
    background: {panel_bg_raised};
    width: 4px;
    border-radius: 2px;
}}
QSlider::handle:vertical {{
    background: {amber_primary};
    border: 1px solid {border};
    height: 12px;
    margin: 0 -5px;
    border-radius: 2px;
}}

/* ---- progress ---- */
QProgressBar {{
    background-color: {input_bg};
    border: 1px solid {border};
    border-radius: 2px;
    color: {text_primary};
    text-align: center;
}}
QProgressBar::chunk {{
    background-color: {amber_secondary};
}}

/* ---- tooltips ---- */
QToolTip {{
    background-color: {panel_bg_raised};
    color: {text_primary};
    border: 1px solid {border};
    padding: 3px 6px;
}}

/* ---- scroll areas / frames stay quiet; palette provides fill ---- */
QScrollArea {{
    border: none;
}}
""".format(**t)


def amber_left_panel_qss() -> str:
    t = AMBER
    return (
        f"QGroupBox:title {{ color: {t['text_secondary']}; }}"
        f" QLabel {{ color: {t['text_primary']}; }}"
        f" QPushButton {{ color: {t['text_primary']}; }}"
    )


def amber_canvas_button_qss() -> str:
    """Amber replacement for the blue 'Open Canvas' toolbar button."""
    t = AMBER
    return (
        "QPushButton {"
        f" background-color: {t['selection_bg']};"
        f" color: {t['amber_bright']};"
        f" border: 1px solid {t['border_active']};"
        " font-weight: 600;"
        " padding: 4px 10px;"
        " border-radius: 2px;"
        "}"
        f"QPushButton:hover {{ background-color: {t['hover_bg']}; }}"
        f"QPushButton:pressed {{ background-color: {t['panel_bg']}; }}"
    )


def amber_canvas_dialog_qss() -> str:
    """Amber counterpart of styles.CANVAS_DIALOG_STYLE (chrome only)."""
    t = AMBER
    return f"""
QDialog {{
    background-color: {t['window_bg']};
    color: {t['text_primary']};
}}
QGroupBox {{
    border: 1px solid {t['border']};
    border-radius: 2px;
    margin-top: 10px;
}}
QGroupBox:title {{
    subcontrol-origin: margin;
    left: 10px;
    padding: 0 4px;
    color: {t['text_secondary']};
}}
QLabel {{
    color: {t['text_primary']};
}}
QLineEdit, QComboBox, QSpinBox {{
    background-color: {t['input_bg']};
    border: 1px solid {t['border']};
    padding: 4px;
    border-radius: 2px;
    color: {t['text_primary']};
}}
QPushButton {{
    background-color: {t['panel_bg']};
    border: 1px solid {t['border']};
    border-radius: 2px;
    padding: 4px 8px;
    color: {t['text_primary']};
}}
QPushButton:hover {{
    background-color: {t['hover_bg']};
}}
QPushButton:pressed {{
    background-color: {t['selection_bg']};
}}
QCheckBox {{
    color: {t['text_primary']};
}}
QScrollArea {{
    background: transparent;
}}
"""


def amber_shortcuts_panel_qss() -> str:
    t = AMBER
    return (
        "QFrame#shortcutsPanel {"
        f" background: {t['panel_bg']};"
        f" border: 1px solid {t['border']};"
        " border-radius: 2px;"
        " padding: 6px;"
        "}"
    )


def plot_chrome_mode(viewer_or_name) -> str:
    """Which matplotlib *chrome* palette a plot should use for the current
    app theme: 'amber' under the amber theme, 'dark' under dark, else
    'light'.  Chrome means figure/axes backgrounds, ticks, labels, grid,
    legend frame — never data artists, colormaps, or trace colors.

    Accepts a theme name, the viewer, or any object holding a ``viewer``
    attribute (popups/dialogs)."""
    if isinstance(viewer_or_name, str):
        return normalize(viewer_or_name)
    obj = viewer_or_name
    candidates = (obj, getattr(obj, "viewer", None))
    for cand in candidates:
        name = getattr(cand, "ui_theme", None) if cand is not None else None
        if name:
            return normalize(name)
    for cand in candidates:
        if cand is not None and getattr(cand, "dark_mode", False):
            return THEME_DARK
    # Leaf widgets without a viewer reference: the app-wide registered theme.
    return _CURRENT_THEME


def mpl_chrome_colors(mode: str) -> dict:
    """Figure/axes chrome colors per theme mode ('light'/'dark'/'amber').

    The light and dark values reproduce the palette the Spectrum popup
    already used, so existing themes render identically."""
    mode = normalize(mode)
    if mode == THEME_AMBER:
        t = AMBER
        return {
            "fig_face": t["window_bg"],
            "ax_face": "#14100a",
            "text": t["text_secondary"],
            "grid": t["selection_bg"],
            "legend_face": t["panel_bg"],
            "legend_edge": t["border"],
            "spine": t["border"],
        }
    if mode == THEME_DARK:
        return {
            "fig_face": "#0f1720",
            "ax_face": "#111827",
            "text": "#e5edf8",
            "grid": "#41526a",
            "legend_face": "#0f1720",
            "legend_edge": "#6f86a5",
            "spine": "#5a6b82",
        }
    return {
        "fig_face": "#ffffff",
        "ax_face": "#ffffff",
        "text": "#1f2937",
        "grid": "#d7deea",
        "legend_face": "#ffffff",
        "legend_edge": "#7d8ea8",
        "spine": "#4b5563",
    }


def style_axes_chrome(ax, colors: dict):
    """Restyle one axes' chrome (background, spines, ticks, labels, title,
    grid line colors).  Never touches images, lines, or collections."""
    try:
        ax.set_facecolor(colors["ax_face"])
        for spine in ax.spines.values():
            spine.set_color(colors["spine"])
        ax.tick_params(colors=colors["text"], labelcolor=colors["text"], which="both")
        ax.xaxis.label.set_color(colors["text"])
        ax.yaxis.label.set_color(colors["text"])
        ax.title.set_color(colors["text"])
        for gline in list(ax.get_xgridlines()) + list(ax.get_ygridlines()):
            gline.set_color(colors["grid"])
    except Exception:
        pass


def style_figure_chrome(fig, mode: str):
    """Apply theme chrome to a whole matplotlib figure (facecolor + every
    axes).  Safe to call after each redraw; data artists are untouched."""
    colors = mpl_chrome_colors(mode)
    try:
        fig.patch.set_facecolor(colors["fig_face"])
    except Exception:
        pass
    for ax in getattr(fig, "axes", []) or []:
        style_axes_chrome(ax, colors)
    return colors


def toggle_pill_palette(dark: bool) -> dict:
    """Palette for the Profile/Spectrum dialogs' pill toggle buttons and
    side panels.  The amber app theme wins over the per-dialog dark flag;
    otherwise the historical blue dark/light sets are returned verbatim."""
    if current_theme() == THEME_AMBER:
        t = AMBER
        return {
            "inactive_bg": t["panel_bg"],
            "inactive_border": t["border"],
            "inactive_text": t["text_primary"],
            "active_bg": t["selection_bg"],
            "active_border": t["border_active"],
            "active_text": t["amber_bright"],
            "hint_color": t["text_secondary"],
            "panel_text": t["text_primary"],
            "list_bg": t["input_bg"],
            "list_alt": t["panel_bg"],
            "list_border": t["border"],
            "compose_button_bg": t["panel_bg"],
            "compose_button_border": t["border"],
            "compose_button_text": t["text_primary"],
            "compose_button_hover_border": t["border_active"],
            "compose_drop_bg": t["selection_bg"],
            "compose_drop_border": t["amber_bright"],
            "radius": "3px",
        }
    if dark:
        return {
            "inactive_bg": "#1e2430",
            "inactive_border": "#46556e",
            "inactive_text": "#d4deee",
            "active_bg": "#2f6fcb",
            "active_border": "#79a9f2",
            "active_text": "#ffffff",
            "hint_color": "#b9c6d8",
            "panel_text": "#dce5f3",
            "list_bg": "#10151f",
            "list_alt": "#171d29",
            "list_border": "#445167",
            "compose_button_bg": "#202633",
            "compose_button_border": "#5a6880",
            "compose_button_text": "#e4ebf7",
            "compose_button_hover_border": "#82aef1",
            "compose_drop_bg": "#26477a",
            "compose_drop_border": "#a8ceff",
            "radius": "12px",
        }
    return {
        "inactive_bg": "#f3f5f9",
        "inactive_border": "#aeb7c5",
        "inactive_text": "#1f2a3d",
        "active_bg": "#1f6fd7",
        "active_border": "#5b97e8",
        "active_text": "#ffffff",
        "hint_color": "#4b5b73",
        "panel_text": "#314056",
        "list_bg": "#ffffff",
        "list_alt": "#f7f9fc",
        "list_border": "#c2cad6",
        "compose_button_bg": "#f5f7fa",
        "compose_button_border": "#b8c2cf",
        "compose_button_text": "#314056",
        "compose_button_hover_border": "#5b97e8",
        "compose_drop_bg": "#e6f0ff",
        "compose_drop_border": "#4f8fe3",
        "radius": "12px",
    }


class TitleBarThemer(QtCore.QObject):
    """App-level event filter: whenever any top-level window is shown,
    stamp the native title bar to match the current theme (dark/amber →
    dark title bar).  Windows-only effect; harmless elsewhere."""

    def __init__(self, viewer):
        super().__init__()
        self._viewer = viewer

    def eventFilter(self, obj, event):
        try:
            if (
                event.type() == QtCore.QEvent.Show
                and isinstance(obj, QtWidgets.QWidget)
                and obj.isWindow()
            ):
                apply_native_titlebar(
                    obj, is_dark_theme(getattr(self._viewer, "ui_theme", THEME_LIGHT))
                )
        except Exception:
            pass
        return False


def apply_native_titlebar(widget, dark: bool):
    """Best-effort dark native title bar on Windows (DWM immersive dark
    mode).  No-ops silently on other platforms or unsupported builds."""
    try:
        # Only on the real Windows platform plugin: offscreen/minimal
        # platforms return synthetic winIds that are not valid HWNDs.
        if QtGui.QGuiApplication.platformName() != "windows":
            return
        import ctypes
        hwnd = int(widget.winId())
        value = ctypes.c_int(1 if dark else 0)
        for attr in (20, 19):  # 20 on current Win10/11, 19 on older builds
            res = ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, attr, ctypes.byref(value), ctypes.sizeof(value)
            )
            if res == 0:
                break
    except Exception:
        pass


_THUMB_FRAME_STYLE_CACHE: dict = {}


def thumb_frame_styles(theme_name: str) -> dict:
    """Selection-frame styles for thumbnail cards, keyed by state.

    These are borders and card backgrounds around the image label only —
    never overlays on the pixmap itself (see module docstring).  The
    non-amber values reproduce the historical hardcoded styles verbatim.
    Memoized because the thumbnail populate loop asks per card.
    """
    key = normalize(theme_name)
    cached = _THUMB_FRAME_STYLE_CACHE.get(key)
    if cached is not None:
        return cached
    if key == THEME_AMBER:
        t = AMBER
        styles = {
            # multi-selected thumbnail (was purple #a36bff) — dashed keeps it
            # distinguishable from single selection without relying on hue
            "multi": (
                f"QFrame {{ border: 2px dashed {t['amber_bright']}; border-radius: 4px;"
                f" background-color: {t['selection_bg']}; }}"
            ),
            # selected/current thumbnail (was blue #5f8dd3)
            "selected": (
                f"QFrame {{ border: 2px solid {t['border_active']}; border-radius: 4px;"
                f" background-color: {t['selection_bg']}; }}"
            ),
            # idle image card (was faint white)
            "default": (
                f"QFrame {{ border: 1px solid {t['border']}; border-radius: 4px;"
                " background-color: transparent; }"
            ),
            # idle spectroscopy miniature card (was soft amber already)
            "spectro_card": (
                f"QFrame {{ border: 1px solid {t['amber_muted']}; border-radius: 4px;"
                f" background-color: {t['hover_bg']}; }}"
            ),
            # multi-selected spectroscopy card (was orange #ff9c3a)
            "spectro_multi": (
                f"QFrame {{ border: 2px dashed {t['amber_bright']}; border-radius: 4px;"
                f" background-color: {t['selection_bg']}; }}"
            ),
        }
    else:
        styles = {
            "multi": "QFrame { border: 2px solid #a36bff; border-radius: 10px; background-color: rgba(163,107,255,40); }",
            "selected": "QFrame { border: 2px solid #5f8dd3; border-radius: 10px; background-color: rgba(95,141,211,40); }",
            "default": "QFrame { border: 1px solid rgba(255,255,255,30); border-radius: 10px; background-color: transparent; }",
            "spectro_card": "QFrame { border: 1px solid rgba(255, 184, 77, 170); border-radius: 10px; background-color: rgba(255, 184, 77, 22); }",
            "spectro_multi": "QFrame { border: 2px solid #ff9c3a; border-radius: 10px; background-color: rgba(255,156,58,42); }",
        }
    _THUMB_FRAME_STYLE_CACHE[key] = styles
    return styles


__all__ = [
    "THEME_LIGHT",
    "THEME_DARK",
    "THEME_AMBER",
    "THEMES",
    "AMBER",
    "AMBER_FONT_STACK",
    "normalize",
    "theme_label",
    "resolve_theme_name",
    "is_dark_theme",
    "is_amber",
    "amber_palette",
    "amber_app_qss",
    "amber_left_panel_qss",
    "amber_canvas_button_qss",
    "amber_canvas_dialog_qss",
    "amber_shortcuts_panel_qss",
    "thumb_frame_styles",
    "plot_chrome_mode",
    "mpl_chrome_colors",
    "style_axes_chrome",
    "style_figure_chrome",
    "apply_native_titlebar",
    "TitleBarThemer",
]
