"""Matrix dataset helpers."""
from __future__ import annotations

import re
from pathlib import Path
class MatrixDataset:
    """Lightweight container describing a matrix dataset and its channel files."""
    def __init__(self, base, rows, cols):
        self.base = base
        self.rows = rows
        self.cols = cols
        self.channels = []  # list of dicts: {'filename','channel_code','label','spectra_count','path'}

    def add_channel(self, filename, channel_code=None, label=None, spectra_count=0, path=None):
        self.channels.append({
            'filename': filename,
            'channel_code': channel_code,
            'label': label,
            'spectra_count': spectra_count,
            'path': str(path) if path else filename,
        })

    def summary(self):
        return f"{self.base}: {len(self.channels)} channel(s) GÇö {self.rows}+ù{self.cols} each"


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
    # attempt to split on the last underscore chunk that contains digits/letters
    m = re.match(r'^(?P<base>.+?)_(?P<code>[0-9A-Za-z]+[^_]*)$', stem)
    if m:
        base = m.group('base')
        channel_code = m.group('code')
    channel_label = channel_code
    return base, channel_code, channel_label
__all__ = [
    "MatrixDataset",
    "parse_matrix_filename",
]



