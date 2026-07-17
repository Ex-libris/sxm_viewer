"""Colormap metadata classification and the ``ColormapSorter`` service.

Qt-free (matplotlib + stdlib only), like ``cmap_registry`` — so the same
"colormap knowledge, not widget knowledge" rule applies: the gallery UI
asks this module for an *ordered, sectioned* list and stays out of the
classification/ordering business entirely.

Two layers:

- **Metadata enrichment.** ``classify(base_name)`` returns
  ``{"type": ..., "primary_color": ...}``. Rather than a hand-maintained
  ``metadata.json`` (which could only ever cover matplotlib's ~170 named
  maps, never the ~1.7k the optional ``colormaps`` package contributes),
  metadata is *derived by sampling each colormap's LUT*: lightness
  profile → functional type, dominant hue → primary color. matplotlib's
  own documented category tables (``_KNOWN_TYPES``) are used as
  authoritative overrides for the maps it names; everything else is
  auto-classified. ``build_metadata_mapping()`` / ``dump_metadata_json()``
  materialize the whole mapping if a static artifact is ever wanted.

- **Sort strategies.** ``ColormapSorter.sort(names, strategy)`` returns a
  list of ``(section_label_or_None, [base_names])`` sections — the sort
  *and* the section headers in one decoupled structure the grid just
  renders. Strategies: ``functional`` (grouped by intent), ``similarity``
  (grouped by base hue), ``usage`` (usage count then recency, flat), and
  ``alphabetical`` (flat). Usage data is injected via a provider callable
  so the sorter never reaches into config/Qt.
"""
from __future__ import annotations

import json

import numpy as np
from matplotlib.colors import ListedColormap, rgb_to_hsv

from . import cmap_registry


# --- Functional taxonomy ---------------------------------------------------
# matplotlib's documented colormap classes (base names, no ``_r``). Used as
# authoritative type overrides; unnamed maps fall back to LUT sampling.
_KNOWN_TYPES = {}
def _tag(names, type_):
    for n in names:
        _KNOWN_TYPES[n] = type_

_tag(["viridis", "plasma", "inferno", "magma", "cividis"], "perceptually_uniform")
_tag(["Greys", "Purples", "Blues", "Greens", "Oranges", "Reds",
      "YlOrBr", "YlOrRd", "OrRd", "PuRd", "RdPu", "BuPu", "GnBu", "PuBu",
      "YlGnBu", "PuBuGn", "BuGn", "YlGn",
      "binary", "gist_yarg", "gist_gray", "gray", "bone", "pink", "spring",
      "summer", "autumn", "winter", "cool", "Wistia", "hot", "afmhot",
      "gist_heat", "copper"], "sequential")
_tag(["PiYG", "PRGn", "BrBG", "PuOr", "RdGy", "RdBu", "RdYlBu", "RdYlGn",
      "Spectral", "coolwarm", "bwr", "seismic"], "diverging")
_tag(["twilight", "twilight_shifted", "hsv"], "cyclic")
_tag(["Pastel1", "Pastel2", "Paired", "Accent", "Dark2", "Set1", "Set2",
      "Set3", "tab10", "tab20", "tab20b", "tab20c"], "qualitative")
_tag(["flag", "prism", "ocean", "gist_earth", "terrain", "gist_stern",
      "gnuplot", "gnuplot2", "CMRmap", "cubehelix", "brg", "gist_rainbow",
      "rainbow", "jet", "turbo", "nipy_spectral", "gist_ncar"], "miscellaneous")

# Section order + display labels for functional grouping.
FUNCTIONAL_ORDER = (
    "perceptually_uniform", "sequential", "diverging",
    "cyclic", "qualitative", "miscellaneous", "other",
)
FUNCTIONAL_LABELS = {
    "perceptually_uniform": "Perceptually Uniform",
    "sequential": "Sequential",
    "diverging": "Diverging",
    "cyclic": "Cyclic",
    "qualitative": "Qualitative",
    "miscellaneous": "Miscellaneous",
    "other": "Other",
}

# Section order + display labels for visual-similarity (hue) grouping.
COLOR_ORDER = (
    "grayscale", "red", "orange", "brown", "yellow", "green",
    "cyan", "blue", "purple", "pink", "multi",
)
COLOR_LABELS = {
    "grayscale": "Monochrome / Grayscale",
    "red": "Red", "orange": "Orange", "brown": "Brown", "yellow": "Yellow",
    "green": "Green", "cyan": "Cyan", "blue": "Blue", "purple": "Purple",
    "pink": "Pink / Magenta", "multi": "Multi-hue / Rainbow",
}

_META_CACHE = {}


def _luminance(rgb):
    # Rec. 709 relative luminance; good enough proxy for perceived lightness.
    return 0.2126 * rgb[:, 0] + 0.7152 * rgb[:, 1] + 0.0722 * rgb[:, 2]


# colorcet (`cet_<class>_...`) encodes the functional class in the token
# after ``cet_`` — the biggest third-party family and 100% reliable by name.
_CET_CLASS = {
    "l": "sequential", "linear": "sequential",
    "d": "diverging", "diverging": "diverging",
    "c": "cyclic", "cyclic": "cyclic",
    "i": "sequential", "isoluminant": "sequential",
    "r": "miscellaneous", "rainbow": "miscellaneous",
    "g": "qualitative", "glasbey": "qualitative",
}

# Qualitative families that ship as ``name_<count>`` variants (ColorBrewer,
# CARTOColors, matplotlib qualitative). Matched on the base token so
# ``accent_3``/``dark2_8``/``bold_5`` all resolve.
_QUALITATIVE_STEMS = {
    "accent", "dark2", "set1", "set2", "set3", "paired", "pastel1",
    "pastel2", "tab10", "tab20", "tab20b", "tab20c", "category10",
    "category20", "bold", "vivid", "antique", "safe", "prism", "pastel",
    "glasbey", "d3", "observable10", "retro", "classic",
}


def _name_type(base):
    """Functional class inferable from the name alone (colorcet class
    letter, qualitative family stems), else ``None``."""
    low = base.lower()
    if low.startswith("cet_"):
        token = low[4:].split("_", 1)[0]
        cls = _CET_CLASS.get(token)
        if cls:
            return cls
    stem = low.rsplit("_", 1)[0] if low.rsplit("_", 1)[-1].isdigit() else low
    if stem in _QUALITATIVE_STEMS:
        return "qualitative"
    return None


def _classify_type_by_sampling(cmap, rgb):
    """Functional class from a colormap object + its sampled RGB (N×3)."""
    n = len(rgb)
    lum = _luminance(rgb)
    steps = np.linalg.norm(np.diff(rgb, axis=0), axis=1)
    dl = np.diff(lum)
    monotonic = bool(np.all(dl >= -0.02) or np.all(dl <= 0.02))
    # Qualitative: abrupt banded jumps (discrete palette) with no smooth
    # lightness ramp. A ListedColormap with few entries is a strong hint,
    # but sequential ColorBrewer maps (e.g. blues_9) are *also* discrete —
    # they stay monotonic in lightness, so the non-monotonic gate keeps
    # them out.
    big_jumps = int(np.count_nonzero(steps > 0.20))
    discrete = isinstance(cmap, ListedColormap) and int(getattr(cmap, "N", 256)) <= 24
    if not monotonic and (big_jumps >= 3 or (discrete and big_jumps >= 2)):
        return "qualitative"
    # Cyclic: endpoints meet (same color at 0 and 1) and it actually moves.
    endpoint_gap = float(np.linalg.norm(rgb[0] - rgb[-1]))
    if endpoint_gap < 0.10 and float(np.mean(steps)) > 0.01:
        return "cyclic"
    if monotonic:
        return "sequential"
    # Interior lightness extremum past both endpoints, roughly centered →
    # diverging (light-center or dark-center).
    mid = lum[n // 2]
    ends = 0.5 * (lum[0] + lum[-1])
    imax, imin = int(np.argmax(lum)), int(np.argmin(lum))
    centered = (0.25 * n) < max(imax, imin) < (0.75 * n) or (0.25 * n) < min(imax, imin) < (0.75 * n)
    if abs(mid - ends) > 0.12 and centered:
        return "diverging"
    return "miscellaneous"


def _classify_color_by_sampling(rgb):
    """Primary-color bucket from sampled RGB via circular-mean hue."""
    hsv = rgb_to_hsv(np.clip(rgb, 0.0, 1.0))
    hue, sat, val = hsv[:, 0], hsv[:, 1], hsv[:, 2]
    colored = (sat > 0.15) & (val > 0.1)
    if np.count_nonzero(colored) < max(2, 0.2 * len(rgb)):
        return "grayscale"
    w = sat[colored]
    ang = hue[colored] * 2.0 * np.pi
    cx = float(np.sum(w * np.cos(ang)))
    cy = float(np.sum(w * np.sin(ang)))
    r = float(np.hypot(cx, cy) / max(1e-9, np.sum(w)))  # 1=concentrated
    if r < 0.6:
        return "multi"  # hue spread wide → rainbow-like
    mean_hue = (np.arctan2(cy, cx) / (2.0 * np.pi)) % 1.0
    return _hue_bucket(mean_hue, float(np.mean(val[colored])))


def _hue_bucket(h, mean_val):
    if h < 0.04 or h >= 0.96:
        return "red"
    if h < 0.075:
        # Dark, low-value oranges read as brown.
        return "brown" if mean_val < 0.5 else "orange"
    if h < 0.11:
        return "orange"
    if h < 0.19:
        return "yellow"
    if h < 0.44:
        return "green"
    if h < 0.52:
        return "cyan"
    if h < 0.72:
        return "blue"
    if h < 0.83:
        return "purple"
    return "pink"


def classify(base_name):
    """Return ``{"type": ..., "primary_color": ...}`` for a base name.

    Cached. Type resolution order: matplotlib's documented tables →
    name-based rules (colorcet class letter, qualitative family stems) →
    LUT sampling (lightness profile + discreteness). Primary color is
    always sampled (auto-derived, so it covers extra-package maps a static
    table never would)."""
    base, _rev = cmap_registry.split_cmap_name(base_name)
    cached = _META_CACHE.get(base)
    if cached is not None:
        return cached
    try:
        cmap = cmap_registry.get_colormap(base, False)
        rgb = np.asarray(cmap(np.linspace(0.0, 1.0, 64)))[:, :3]
        type_ = (_KNOWN_TYPES.get(base) or _name_type(base)
                 or _classify_type_by_sampling(cmap, rgb))
        color = _classify_color_by_sampling(rgb)
    except Exception:
        type_ = _KNOWN_TYPES.get(base) or _name_type(base) or "other"
        color = "grayscale"
    meta = {"type": type_, "primary_color": color}
    _META_CACHE[base] = meta
    return meta


def build_metadata_mapping(names=None):
    """The full ``{name: {type, primary_color}}`` mapping — the
    materialized form of the auto-classifier (a ``metadata.json`` if you
    want one, but normally consumed live via :func:`classify`)."""
    if names is None:
        names = cmap_registry.base_cmap_names()
    return {n: classify(n) for n in names}


def dump_metadata_json(path, names=None):
    """Write :func:`build_metadata_mapping` to ``path`` as JSON."""
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(build_metadata_mapping(names), fh, indent=2, sort_keys=True)


# --- Sorter service --------------------------------------------------------

STRATEGIES = ("functional", "similarity", "usage", "alphabetical")
STRATEGY_LABELS = {
    "functional": "Function",
    "similarity": "Color",
    "usage": "Recently used",
    "alphabetical": "Name (A–Z)",
}
DEFAULT_STRATEGY = "functional"


class ColormapSorter:
    """Turns a flat list of base colormap names into ordered sections.

    ``usage_provider`` (optional) is a callable ``() -> {base_name:
    (count, last_used_ts)}`` — injected so the sorter stays decoupled from
    config and Qt; only the ``usage`` strategy consults it.
    """

    def __init__(self, usage_provider=None):
        self._usage_provider = usage_provider

    def sort(self, names, strategy):
        strategy = strategy if strategy in STRATEGIES else DEFAULT_STRATEGY
        names = [str(n) for n in names]
        if strategy == "alphabetical":
            return [(None, sorted(names, key=str.lower))]
        if strategy == "usage":
            return self._sort_usage(names)
        if strategy == "similarity":
            return self._group(names, key="primary_color",
                                order=COLOR_ORDER, labels=COLOR_LABELS)
        return self._group(names, key="type",
                           order=FUNCTIONAL_ORDER, labels=FUNCTIONAL_LABELS)

    def _group(self, names, key, order, labels):
        buckets = {}
        for name in names:
            bucket = classify(name).get(key) or "other"
            buckets.setdefault(bucket, []).append(name)
        rank = {b: i for i, b in enumerate(order)}
        sections = []
        # Known buckets in taxonomy order, then any stragglers alphabetically.
        for bucket in sorted(buckets, key=lambda b: (rank.get(b, len(order)), b)):
            members = sorted(buckets[bucket], key=str.lower)
            sections.append((labels.get(bucket, bucket.title()), members))
        return sections

    def _sort_usage(self, names):
        stats = {}
        if self._usage_provider is not None:
            try:
                stats = dict(self._usage_provider() or {})
            except Exception:
                stats = {}

        def usage_key(name):
            count, last = stats.get(name, (0, 0.0))
            # Descending count, then descending recency, then A–Z.
            return (-int(count or 0), -float(last or 0.0), name.lower())

        ordered = sorted(names, key=usage_key)
        used = [n for n in ordered if int(stats.get(n, (0, 0))[0] or 0) > 0]
        unused = [n for n in ordered if n not in set(used)]
        sections = []
        if used:
            sections.append(("Recently used", used))
        sections.append(("All colormaps" if used else None, unused))
        return sections
