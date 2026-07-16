"""Single source of truth for colormap registration, enumeration, and
resolution.

This module is deliberately Qt-free (only matplotlib + stdlib) so the
headless packages (``reporting/``) can use it, mirroring the import
isolation rules in CLAUDE.md.

The matplotlib global registry (``matplotlib.colormaps``) remains the
underlying store: everything registered here lands in it, so code that
still calls matplotlib directly keeps resolving every name registered
here. Migration of call sites can therefore be incremental.

Public surface:

- ``ensure_registered()`` — idempotent registration of the custom amber
  colormap (always) and the optional pratiman-91 ``colormaps`` package
  (when importable). Called lazily by every other public function.
- ``all_cmap_names()`` / ``featured_cmap_names(context)`` /
  ``grouped_cmap_names()`` — enumeration for the GUI combo boxes.
- ``get_cmap(name, fallback)`` — never-raising name -> Colormap
  resolution with a shared object cache.
- ``set_forced_cmap(name)`` / ``effective_cmap_name(name)`` /
  ``effective_cmap(name, fallback)`` — the display-time override hook
  used by the "Full amber imagery" mode. The registry stays
  theme-unaware: the GUI layer decides when to force (same module-global
  pattern as ``gui.theme._CURRENT_THEME``).
"""

from collections import OrderedDict

import matplotlib
from matplotlib import colors as _mcolors


# --- Custom amber colormap -------------------------------------------------
# Stops mirror gui/theme.py's AMBER tokens (window_bg, amber_muted,
# amber_primary). Keep the two files in sync — this module must stay
# Qt-free, so it cannot import gui.theme.
AMBER_CMAP_NAME = "gui_amber_theme"
AMBER_CMAP_STOPS = ("#100b05", "#805b18", "#ffb000")

# The optional extra-colormaps package (https://pratiman-91.github.io/colormaps/).
EXTRA_PACKAGE_NAME = "colormaps"
EXTRA_PACKAGE_INSTALL_HINT = "pip install colormaps"

_REGISTERED = False
_EXTRA_STATUS = {"available": False, "count": 0, "skipped": 0, "error": None}
_EXTRA_NAMES: list = []

_CMAP_OBJECT_CACHE: dict = {}

_FORCED_CMAP = None

# Curated per-context featured lists. These are colormap knowledge, not
# widget knowledge — keeping them here (instead of scattered per-widget
# constants) is what prevents the lists from drifting apart again.
_FEATURED = {
    # Mirrors the historical `common_cmaps` shortlist from the preview
    # canvas context menu (Blues_r is the app's default thumbnail/preview
    # cmap — keep it featured).
    "general": [
        "viridis", "plasma", "inferno", "magma", "cividis",
        "turbo", "gray", "afmhot", "Blues_r", "RdBu_r", "coolwarm",
    ],
    # Historical gui/controllers/image_compare.py TOPO_CMAPS.
    "topo": [
        "viridis", "plasma", "magma", "inferno", "cividis", "afmhot", "gray",
    ],
    # Historical publication-canvas tile context-menu shortlist
    # (gui/canvases/canvas_items.py).
    "canvas_tile": [
        "viridis", "plasma", "magma", "inferno", "cividis", "afmhot",
        "gray", "Blues_r", "RdBu_r",
    ],
    # Historical gui/canvases/molecular_overlay.py curated categories,
    # following matplotlib's documented colormap classes. Diverging fits a
    # signed deviation best; sequential is for magnitude-only coloring;
    # qualitative is for categorical data (no implied ordering).
    "molecule_diverging": [
        "coolwarm", "RdBu", "RdYlBu", "RdYlGn", "PiYG", "PRGn", "BrBG",
        "PuOr", "Spectral", "seismic", "bwr",
    ],
    "molecule_sequential": [
        "viridis", "plasma", "inferno", "magma", "cividis", "YlOrRd",
        "YlGnBu", "OrRd", "PuBu", "BuGn",
    ],
    "molecule_qualitative": [
        "tab10", "Set2", "Set1", "Dark2", "Accent", "Paired", "Pastel1",
        "Pastel2", "tab20",
    ],
}


def _register_amber():
    if AMBER_CMAP_NAME in matplotlib.colormaps:
        return
    cmap = _mcolors.LinearSegmentedColormap.from_list(
        AMBER_CMAP_NAME, list(AMBER_CMAP_STOPS)
    )
    matplotlib.colormaps.register(cmap)
    reversed_name = AMBER_CMAP_NAME + "_r"
    if reversed_name not in matplotlib.colormaps:
        matplotlib.colormaps.register(cmap.reversed(reversed_name))


def _register_extra_package():
    global _EXTRA_STATUS, _EXTRA_NAMES
    # Recent package versions self-register their colormaps into
    # matplotlib's global registry as a side effect of the import, so
    # snapshot the registry first: anything that appears afterwards was
    # contributed by the package.
    before = set(matplotlib.colormaps)
    try:
        import colormaps as _pkg  # pratiman-91, PyPI "colormaps"
    except Exception as exc:
        is_missing = isinstance(exc, ImportError)
        _EXTRA_STATUS = {
            "available": False,
            "count": 0,
            "skipped": 0,
            "error": None if is_missing else f"{type(exc).__name__}: {exc}",
        }
        return

    skipped = 0
    # The package exposes its colormaps as module attributes; accessing an
    # attribute lazily builds the Colormap AND (in current versions)
    # registers it into matplotlib as a side effect. So: touch every
    # attribute, register manually only when the access didn't already do
    # it, and treat a name matplotlib had *before* the import as a genuine
    # collision (matplotlib's own colormap wins).
    for attr in dir(_pkg):
        if attr.startswith("_"):
            continue
        try:
            obj = getattr(_pkg, attr)
        except Exception:
            continue
        if not isinstance(obj, _mcolors.Colormap):
            continue
        name = str(getattr(obj, "name", attr)) or attr
        if name in before:
            skipped += 1
            continue
        if name not in matplotlib.colormaps:
            try:
                matplotlib.colormaps.register(obj, name=name)
            except Exception:
                skipped += 1
    # Everything that appeared in matplotlib's registry since the snapshot
    # came from the package (attribute side effects included).
    names = set(matplotlib.colormaps) - before
    _EXTRA_NAMES = sorted(names)
    _EXTRA_STATUS = {
        "available": True,
        "count": len(names),
        "skipped": skipped,
        "error": None,
    }


def ensure_registered():
    """Idempotently register the amber cmap and any optional extras."""
    global _REGISTERED
    if _REGISTERED:
        return
    _REGISTERED = True
    _register_amber()
    _register_extra_package()


def extra_colormaps_status():
    """Status dict for the Display-menu entry:
    {"available": bool, "count": int, "skipped": int, "error": str|None}.
    """
    ensure_registered()
    return dict(_EXTRA_STATUS)


def extra_cmap_names():
    ensure_registered()
    return list(_EXTRA_NAMES)


def all_cmap_names():
    """Every selectable colormap name, sorted (the single replacement for
    the scattered ``sorted(colormaps.keys())`` + private fallback lists)."""
    ensure_registered()
    return sorted(matplotlib.colormaps.keys())


def featured_cmap_names(context="general"):
    """Curated shortlist for a context, filtered to registered names."""
    ensure_registered()
    names = _FEATURED.get(context) or _FEATURED["general"]
    return [n for n in names if n in matplotlib.colormaps]


def grouped_cmap_names():
    """OrderedDict of {group label: [names]} for pickers that want to
    present the (potentially large) list grouped by source."""
    ensure_registered()
    featured = featured_cmap_names("general")
    custom = [n for n in (AMBER_CMAP_NAME, AMBER_CMAP_NAME + "_r")
              if n in matplotlib.colormaps]
    extras = sorted(_EXTRA_NAMES)
    shown = set(featured) | set(custom) | set(extras)
    rest = [n for n in all_cmap_names() if n not in shown]
    return OrderedDict((
        ("Featured", featured),
        ("Custom", custom),
        ("Extra (colormaps pkg)", extras),
        ("Matplotlib", rest),
    ))


def get_cmap(name, fallback="viridis"):
    """Resolve a colormap name to a Colormap object; never raises.

    Returned objects are cached and shared: matplotlib's ``get_cmap``
    rebuilds the LUT on every call (~ms each across hundreds of
    thumbnails), and no call site in this codebase mutates a returned
    Colormap (no set_bad/set_over/set_under) — keep it that way, the
    cache relies on it.
    """
    ensure_registered()
    key = str(name) if name else str(fallback)
    cmap = _CMAP_OBJECT_CACHE.get(key)
    if cmap is not None:
        return cmap
    try:
        cmap = matplotlib.colormaps.get_cmap(key)
    except Exception:
        try:
            cmap = matplotlib.colormaps.get_cmap(str(fallback))
        except Exception:
            cmap = matplotlib.colormaps.get_cmap("viridis")
    _CMAP_OBJECT_CACHE[key] = cmap
    return cmap


# --- Display-time forced override (the "Full amber imagery" hook) ----------

def set_forced_cmap(name):
    """Force every registry-routed render surface to one colormap
    (``None`` restores normal per-view colormaps). Display-time only:
    callers must never write the forced name back into per-file state."""
    global _FORCED_CMAP
    _FORCED_CMAP = str(name) if name else None


def forced_cmap_name():
    return _FORCED_CMAP


def effective_cmap_name(name):
    """The colormap name a renderer should actually draw with."""
    return _FORCED_CMAP if _FORCED_CMAP else name


def effective_cmap(name, fallback="viridis"):
    return get_cmap(effective_cmap_name(name), fallback=fallback)
