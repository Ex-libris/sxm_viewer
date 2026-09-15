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
- ``split_cmap_name(name)`` / ``join_cmap_name(base, is_reversed)`` —
  translate between the legacy matplotlib ``X_r`` string convention
  (still the persistence/interchange format: ``per_file_channel_cmap``,
  sessions, exports all store joined strings) and the
  ``(base name, is_reversed)`` pair model used by the gallery/manager.
- ``get_colormap(name, is_reversed, fallback)`` — never-raising
  resolution of a ``(base, reversed)`` pair; the flip is programmatic
  (``Colormap.reversed()``), not a registry-name lookup, so it works
  for cmaps that have no registered ``_r`` twin (it registers the
  reversed variant on first use so joined names keep resolving
  app-wide).
- ``base_cmap_names()`` / ``grouped_base_cmap_names()`` /
  ``featured_cmap_entries(context)`` — ``_r``-free enumeration for the
  colormap gallery (one entry per base map; direction is UI state).
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
# Entries are (base name, is_reversed) pairs: direction is data, never an
# ``_r`` suffix baked into the name (legacy joined strings come out of
# featured_cmap_names, which composes them on the fly).
_FEATURED = {
    # Mirrors the historical `common_cmaps` shortlist from the preview
    # canvas context menu (reversed Blues is the app's default
    # thumbnail/preview cmap — keep it featured).
    "general": [
        ("viridis", False), ("plasma", False), ("inferno", False),
        ("magma", False), ("cividis", False), ("turbo", False),
        ("gray", False), ("afmhot", False), ("Blues", True),
        ("RdBu", True), ("coolwarm", False),
    ],
    # Historical gui/controllers/image_compare.py TOPO_CMAPS.
    "topo": [
        ("viridis", False), ("plasma", False), ("magma", False),
        ("inferno", False), ("cividis", False), ("afmhot", False),
        ("gray", False),
    ],
    # Historical publication-canvas tile context-menu shortlist
    # (gui/canvases/canvas_items.py).
    "canvas_tile": [
        ("viridis", False), ("plasma", False), ("magma", False),
        ("inferno", False), ("cividis", False), ("afmhot", False),
        ("gray", False), ("Blues", True), ("RdBu", True),
    ],
    # Historical gui/canvases/molecular_overlay.py curated categories,
    # following matplotlib's documented colormap classes. Diverging fits a
    # signed deviation best; sequential is for magnitude-only coloring;
    # qualitative is for categorical data (no implied ordering).
    "molecule_diverging": [
        ("coolwarm", False), ("RdBu", False), ("RdYlBu", False),
        ("RdYlGn", False), ("PiYG", False), ("PRGn", False),
        ("BrBG", False), ("PuOr", False), ("Spectral", False),
        ("seismic", False), ("bwr", False),
    ],
    "molecule_sequential": [
        ("viridis", False), ("plasma", False), ("inferno", False),
        ("magma", False), ("cividis", False), ("YlOrRd", False),
        ("YlGnBu", False), ("OrRd", False), ("PuBu", False),
        ("BuGn", False),
    ],
    "molecule_qualitative": [
        ("tab10", False), ("Set2", False), ("Set1", False),
        ("Dark2", False), ("Accent", False), ("Paired", False),
        ("Pastel1", False), ("Pastel2", False), ("tab20", False),
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


# Qualitative (categorical) colormaps contributed by the optional
# "colormaps" package, worth surfacing as discrete color cycles for
# line/marker plots. The package groups maps only by source library
# (cartocolors/colorbrewer/tableau/...), never by "qualitative", so this
# is a curated shortlist — same "colormap knowledge lives here, not in
# widgets" rationale as _FEATURED above. Names use the package's own
# (mostly lowercase) registered spelling; qualitative_extra_cmap_names()
# filters to what the package actually contributed, so matplotlib-builtin
# shadows the package couldn't override (e.g. "prism") drop out.
_QUALITATIVE_EXTRA_CANDIDATES = (
    # cartocolors qualitative
    "antique", "bold", "pastel", "safe", "vivid",
    # colorbrewer qualitative (package lowercase variants; distinct from
    # matplotlib's capitalized Set1/Dark2/... builtins)
    "accent", "dark2", "paired", "pastel1", "pastel2",
    "set1", "set2", "set3",
    # tableau (all qualitative)
    "tableau_10", "tableau_20", "tableaulight_10", "tableaumedium_10",
    "colorblind_10", "trafficlight_9",
    "bluered_6", "bluered_12", "greenorange_6", "greenorange_12",
    "purplegray_6", "purplegray_12", "gray_5",
    # ncar_ncl categorical
    "Cat12",
)


def qualitative_extra_cmap_names():
    """Registered qualitative/categorical colormaps from the optional
    ``colormaps`` package, suitable as discrete color cycles. Empty list
    when the package isn't installed. Filtered to names the package
    actually contributed (``_EXTRA_NAMES``), so a matplotlib builtin that
    shadowed a same-named package map is not mistaken for the qualitative
    variant."""
    ensure_registered()
    extra = set(_EXTRA_NAMES)
    return [n for n in _QUALITATIVE_EXTRA_CANDIDATES if n in extra]


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


def featured_cmap_entries(context="general"):
    """Curated shortlist for a context as ``(base_name, is_reversed)``
    pairs — the primary data model; direction is a flag, never an ``_r``
    name. Filtered to registered base names."""
    ensure_registered()
    entries = _FEATURED.get(context) or _FEATURED["general"]
    return [(n, bool(r)) for (n, r) in entries if n in matplotlib.colormaps]


def featured_cmap_names(context="general"):
    """Legacy joined-string view of ``featured_cmap_entries`` for callers
    that consume matplotlib-style names (combo boxes, menus). Output is
    unchanged from the historical ``_r``-suffixed lists."""
    names = []
    for base, is_reversed in featured_cmap_entries(context):
        if is_reversed:
            # Materialize/register the reversed variant so the joined name
            # resolves everywhere a plain string is expected.
            get_colormap(base, True)
        names.append(join_cmap_name(base, is_reversed))
    return names


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


# --- Base-name / reversed-flag model (colormap gallery + manager) ----------

def split_cmap_name(name):
    """``"Blues_r"`` -> ``("Blues", True)``; anything else -> ``(name, False)``.

    The ``_r`` suffix is only treated as a reversal marker when the base
    name is itself registered, so a hypothetical map genuinely named
    ``*_r`` with no forward twin is left alone."""
    ensure_registered()
    s = str(name or "")
    if s.endswith("_r"):
        base = s[:-2]
        if base and base in matplotlib.colormaps:
            return base, True
    return s, False


def join_cmap_name(base, is_reversed):
    """Compose the legacy string form of a ``(base, is_reversed)`` pair.

    Joined strings remain the persistence/interchange format
    (``per_file_channel_cmap``, sessions, exports); ``base`` must already
    be a base name (use ``split_cmap_name`` first if unsure)."""
    return str(base) + "_r" if is_reversed else str(base)


def get_colormap(name, is_reversed=False, fallback="viridis"):
    """Resolve a ``(name, is_reversed)`` pair to a Colormap; never raises.

    The transformation is programmatic — ``Colormap.reversed()`` — not a
    ``_r``-filename lookup, so it works for any registered map (including
    extra-package maps with no registered reversed twin). An ``_r`` suffix
    already present in ``name`` is folded into the flag (XOR: asking for
    the reverse of ``"Blues_r"`` yields forward Blues). On first use of a
    reversed variant with no registered twin, the variant is registered
    under the joined name so plain-string call sites (sessions, combos,
    exports) keep resolving it. Shares ``get_cmap``'s object cache —
    same rule applies: never mutate a returned Colormap."""
    ensure_registered()
    base, name_rev = split_cmap_name(name if name else fallback)
    is_reversed = bool(is_reversed) ^ name_rev
    if base not in matplotlib.colormaps:
        return get_cmap(str(fallback))
    if not is_reversed:
        return get_cmap(base, fallback=fallback)
    joined = join_cmap_name(base, True)
    cached = _CMAP_OBJECT_CACHE.get(joined)
    if cached is not None:
        return cached
    if joined in matplotlib.colormaps:
        return get_cmap(joined, fallback=fallback)
    forward = get_cmap(base, fallback=fallback)
    try:
        cmap = forward.reversed(joined)
    except Exception:
        return forward
    try:
        matplotlib.colormaps.register(cmap, name=joined)
    except Exception:
        pass
    _CMAP_OBJECT_CACHE[joined] = cmap
    return cmap


def base_cmap_names():
    """Every selectable colormap collapsed to base names, sorted —
    ``_r`` twins are folded into their forward map (the gallery's
    enumeration; direction is per-selection UI state, not a list entry)."""
    ensure_registered()
    names = set(matplotlib.colormaps)
    return sorted(n for n in names
                  if not (n.endswith("_r") and n[:-2] in names))


def grouped_base_cmap_names():
    """``grouped_cmap_names`` counterpart for the gallery: OrderedDict of
    {group label: [base names]} with ``_r`` twins collapsed."""
    ensure_registered()
    featured = []
    for base, _rev in featured_cmap_entries("general"):
        if base not in featured:
            featured.append(base)
    custom = [AMBER_CMAP_NAME] if AMBER_CMAP_NAME in matplotlib.colormaps else []
    extras = sorted({split_cmap_name(n)[0] for n in _EXTRA_NAMES})
    shown = set(featured) | set(custom) | set(extras)
    rest = [n for n in base_cmap_names() if n not in shown]
    return OrderedDict((
        ("Featured", featured),
        ("Custom", custom),
        ("Extra (colormaps pkg)", extras),
        ("Matplotlib", rest),
    ))


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
