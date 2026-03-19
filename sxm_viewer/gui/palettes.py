"""Reusable color cycle definitions for spectroscopy plots."""
from __future__ import annotations

from collections import OrderedDict
from typing import List

COLOR_CYCLES = OrderedDict({
    "Vivid": [
        "#4c78a8", "#f58518", "#e45756", "#72b7b2", "#54a24b",
        "#eeca3b", "#b279a2", "#ff9da6", "#9c755f", "#bab0ab",
    ],
    "Rainbow": [
        "#e41a1c", "#ff7f00", "#ffff33", "#4daf4a", "#377eb8",
        "#984ea3", "#f781bf", "#000000",
    ],
    "Grayscale": [
        "#000000", "#1f1f1f", "#4a4a4a", "#7a7a7a", "#a6a6a6",
        "#d0d0d0", "#f0f0f0",
    ],
    "Viridis": [
        "#440154", "#482878", "#3e4989", "#31688e", "#26838f",
        "#1f9d8a", "#6cce5a", "#b6de2b", "#fee825",
    ],
    "Plasma": [
        "#0d0887", "#46039f", "#7201a8", "#9c179e", "#bd3786",
        "#d8576b", "#ed7953", "#fb9f3a", "#fdca26", "#f0f921",
    ],
    "Cividis": [
        "#002051", "#183364", "#2f476f", "#465d75", "#5e7375",
        "#778975", "#909e71", "#abb366", "#c6c55a", "#e1d550",
    ],
    "Midnight": [
        "#003f5c", "#2f4b7c", "#665191", "#a05195", "#d45087",
        "#f95d6a", "#ff7c43", "#ffa600",
    ],
    "Tableau": [
        "#4E79A7", "#F28E2B", "#E15759", "#76B7B2", "#59A14F",
        "#EDC948", "#B07AA1", "#FF9DA7", "#9C755F", "#BAB0AC",
    ],
    "ColorBlind": [
        "#1170aa", "#fc7d0b", "#a3acb9", "#57606c", "#5fa2ce",
        "#c85200", "#7b848f", "#a1cdec", "#ffbc79", "#b79f00",
    ],
    "TolVibrant": [
        "#EE6677", "#228833", "#4477AA", "#CCBB44", "#66CCEE",
        "#AA3377", "#BBBBBB",
    ],
    "TolMuted": [
        "#332288", "#88CCEE", "#44AA99", "#117733", "#999933",
        "#DDCC77", "#CC6677", "#882255", "#AA4499",
    ],
    "TolLight": [
        "#77AADD", "#99DDFF", "#44BB99", "#BBCC33", "#AAAA00",
        "#EEDD88", "#EE8866", "#FFAABB", "#DDDDDD",
    ],
    "Pastel": [
        "#AEC6CF", "#FFB347", "#FF6961", "#77DD77", "#CFCFC4",
        "#F49AC2", "#B39EB5", "#E0BBE4", "#C7CEEA", "#FFDAC1",
        "#E2F0CB", "#B5EAD7",
    ],
    "Set2": [
        "#66c2a5", "#fc8d62", "#8da0cb", "#e78ac3", "#a6d854",
        "#ffd92f", "#e5c494", "#b3b3b3",
    ],
    "Set3": [
        "#8dd3c7", "#ffffb3", "#bebada", "#fb8072", "#80b1d3",
        "#fdb462", "#b3de69", "#fccde5", "#d9d9d9", "#bc80bd",
        "#ccebc5", "#ffed6f",
    ],
    "Warm": [
        "#6e40aa", "#bf3a96", "#de4968", "#f05f57", "#f47b4a",
        "#f9a249", "#f8c05c", "#f3d76b", "#f0f921",
    ],
    "Cool": [
        "#0d0887", "#4601a8", "#7201b8", "#9c179e", "#bd3786",
        "#d8576b", "#ed7953", "#fb9f3a", "#fdc527", "#f7e225",
    ],
    "Solar": [
        "#00204c", "#00356f", "#004b92", "#0062b5", "#1877c9",
        "#3d8dd3", "#62a4dc", "#86bae4", "#abd1ed", "#cfe7f5",
    ],
    "Black & White": [
        "#000000", "#ffffff", "#666666", "#999999", "#cccccc",
    ],
})

DEFAULT_COLOR_CYCLE = next(iter(COLOR_CYCLES))


def list_color_cycles() -> List[str]:
    return list(COLOR_CYCLES.keys())


def get_color_cycle(name: str | None) -> List[str]:
    if not name:
        return COLOR_CYCLES[DEFAULT_COLOR_CYCLE][:]
    return list(COLOR_CYCLES.get(name, COLOR_CYCLES[DEFAULT_COLOR_CYCLE]))


__all__ = ["COLOR_CYCLES", "DEFAULT_COLOR_CYCLE", "list_color_cycles", "get_color_cycle"]
