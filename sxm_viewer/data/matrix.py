"""Matrix dataset helpers."""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict

import numpy as np


class MatrixDataset:
    """Lightweight container describing a matrix dataset and its channel files."""

    def __init__(self, base, rows, cols):
        self.base = base
        self.rows = rows
        self.cols = cols
        # list of dicts: {'filename','channel_code','label','spectra_count','path','points_per_trace'}
        self.channels = []

    def add_channel(self, filename, channel_code=None, label=None, spectra_count=0, path=None, points_per_trace=None):
        self.channels.append({
            'filename': filename,
            'channel_code': channel_code,
            'label': label,
            'spectra_count': spectra_count,
            'path': str(path) if path else filename,
            'points_per_trace': points_per_trace,
        })

    def summary(self):
        return f"{self.base}: {len(self.channels)} channel(s) — {self.rows}×{self.cols} each"


@dataclass
class MatrixDataCube:
    """Structured dataset produced from a matrix .dat file."""

    path: str
    dataset_key: str
    channel: str
    bias: np.ndarray
    x: np.ndarray
    y: np.ndarray
    data: np.ndarray  # shape: (bias, rows, cols)
    metadata: Dict[str, Any] = field(default_factory=dict)


def parse_matrix_filename(fname: str):
    """
    Heuristic parser for matrix filenames.
    Returns (base, channel_code, channel_label).
    Examples:
      angii_au111_00df_Matrix.dat -> base=angii_au111, channel_code=00df
      angii_au111_00It_to_PC_Matrix.dat -> base=angii_au111, channel_code=00It_to_PC
    """
    stem = Path(fname).stem
    # strip extension and trailing "_Matrix" if present
    stem = re.sub(r'(?i)_matrix$', '', stem)
    channel_code = None
    base = stem
    # First try to split on the first underscore that introduces the numeric acquisition id.
    m = re.match(r'^(?P<base>.+?)_(?P<code>\d.*)$', stem)
    if not m:
        # Fallback: split on the final underscore
        m = re.match(r'^(?P<base>.+?)_(?P<code>[^_]+)$', stem)
    if m:
        base = m.group('base')
        channel_code = m.group('code')
    channel_label = channel_code
    return base, channel_code, channel_label


def matrix_dataset_key(base: str | None, channel_code: str | None):
    """
    Return (dataset_key, display_label) for a matrix acquisition.

    The dataset key groups the simultaneous channels captured for a given
    acquisition. Many labs encode the acquisition index as a numeric prefix
    (e.g. ``07df``, ``07Topo``). We strip that prefix for the human-facing
    label but keep it in the dataset key so that ``07df`` / ``07Topo`` are
    grouped separately from ``13df`` / ``13Topo``.
    """
    base = base or ""
    dataset_key = base or ""
    display_label = channel_code or ""
    if channel_code:
        m = re.match(r"^(\d+)(.*)$", channel_code)
        if m:
            prefix = m.group(1)
            rest = (m.group(2) or "").strip("_- ")
            if prefix:
                dataset_key = f"{base}_{prefix}" if base else prefix
            display_label = rest or channel_code
        else:
            display_label = channel_code
    dataset_key = dataset_key or base or ""
    display_label = display_label or channel_code or base or "channel"
    return dataset_key, display_label


__all__ = [
    "MatrixDataset",
    "MatrixDataCube",
    "parse_matrix_filename",
    "matrix_dataset_key",
]
