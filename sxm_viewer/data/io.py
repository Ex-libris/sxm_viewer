"""
Low-level helpers for SXM image headers and channel data.

The original viewer relied on inline parsing logic for Omicron/Anfatec style
```.txt``` headers.  The rewritten GUI imports these helpers, so we provide a
compact reimplementation that focuses on resiliency instead of format-perfect
parsing.  The overall goal is:

* Extract numeric scan metadata (`xPixel`, `XScanRange`, etc.) from arbitrary
  `key=value` or `key: value` lines.
* Collect per-channel descriptors (`FileName`, `Caption`, `PhysUnit`, ...).
* Read the referenced data files into NumPy arrays, supporting both ASCII and
  binary float dumps.
* Convert per-channel data into consistent units when possible so downstream
  logic can compare heights/currents safely.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import numpy as np


@dataclass
class ChannelInfo:
    """Descriptor for a single SXM channel entry."""

    caption: str = ""
    file_name: str = ""
    phys_unit: str = ""
    scale: float = 1.0
    offset: float = 0.0

    def as_dict(self) -> Dict[str, object]:
        return {
            "Caption": self.caption,
            "FileName": self.file_name,
            "PhysUnit": self.phys_unit,
            "Scale": self.scale,
            "Offset": self.offset,
        }


def parse_header(path: Path | str) -> Tuple[Dict[str, object], List[Dict[str, object]]]:
    """
    Parse an Omicron/Anfatec style header file.

    The real-world files frequently look like::

        xPixel = 256
        yPixel = 256
        XScanRange[nm] = 80
        YScanRange[nm] = 80

        [Channel 0]
        FileName = frame00000.DAT
        Caption = Topography
        PhysUnit = nm
        Scale = 1.0
        Offset = 0.0

    Some acquisitions use ``key:value`` pairs, others rely on tab-delimited
    text, and the encoding is usually cp1252.  We relax the parser to accept
    whichever separator appears and only keep fields referenced by the GUI.
    """
    path = Path(path)
    try:
        raw = path.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        raw = path.read_text(encoding="cp1252", errors="ignore")
    header: Dict[str, object] = {}
    channels: List[ChannelInfo] = []
    current: Optional[ChannelInfo] = None

    def _flush_current():
        nonlocal current
        if current is None:
            return
        if current.file_name:
            channels.append(current)
        current = None

    for line in raw.splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("#"):
            continue
        lower = stripped.lower()
        if lower in {"filedescbegin", "filedescend"}:
            # ANFATEC headers wrap channel blocks with FileDescBegin/End markers.
            if lower == "filedescbegin":
                _flush_current()
                current = ChannelInfo()
            else:
                _flush_current()
            continue
        if stripped.startswith("[") and stripped.endswith("]"):
            # Section markers typically separate channels; start a new record.
            _flush_current()
            current = ChannelInfo()
            continue
        key, value = _split_key_value(stripped)
        if key is None:
            continue
        key_norm = key.lower()
        if key_norm.startswith("file") or key_norm.startswith("chan"):
            if current is None:
                current = ChannelInfo()
        target = header
        if current is not None and key_norm in {
            "filename",
            "file",
            "caption",
            "physunit",
            "scale",
            "offset",
        }:
            target = current.__dict__
        parsed_val = _coerce_value(value)
        if target is header:
            header[_canonical_header_key(key)] = parsed_val
        else:
            if key_norm.startswith("file"):
                target["file_name"] = Path(str(parsed_val)).name
            elif key_norm == "caption":
                target["caption"] = str(parsed_val)
            elif key_norm == "physunit":
                target["phys_unit"] = str(parsed_val)
            elif key_norm == "scale":
                target["scale"] = float(parsed_val)
            elif key_norm == "offset":
                target["offset"] = float(parsed_val)
    _flush_current()
    if not channels:
        # Some legacy headers describe channels inline without explicit sections.
        inline = _parse_inline_channels(header)
        if inline:
            channels.extend(inline)
    return header, [ch.as_dict() for ch in channels]


def read_channel_file(
    path: Path | str,
    xpix: int,
    ypix: int,
    *,
    scale: float = 1.0,
    offset: float = 0.0,
) -> np.ndarray:
    """
    Read a channel file as a 2D NumPy array.

    The SXM controller stores data either as ASCII grids (tab separated) or as
    binary little-endian floats.  We try a lightweight cascade:

    1. Attempt to load via ``np.loadtxt`` (handles ASCII tables quickly).
    2. Fallback to ``np.fromfile`` assuming ``float32`` then ``float64``.
    3. Parse whitespace-delimited tokens manually as a last resort.

    The viewer expects values scaled/offset in the same way as the controller,
    so we apply both parameters regardless of the data type.
    """
    target_count = int(xpix) * int(ypix)
    path = Path(path)
    arr = _load_ascii_grid(path, target_count)
    if arr is None:
        arr = _load_binary_with_inference(path, target_count)
    if arr is None:
        arr = _load_tokenized_grid(path, target_count)
    if arr is None:
        raise ValueError(f"Unable to load channel data from {path}")
    arr = np.asarray(arr, dtype=float)
    if arr.size < target_count:
        padded = np.full(target_count, np.nan, dtype=float)
        padded[: arr.size] = arr
        arr = padded
    arr = arr[:target_count]
    arr = arr.reshape(int(ypix), int(xpix))
    arr = arr * float(scale) + float(offset)
    return arr


def normalize_unit_and_data(
    arr: np.ndarray, unit: str | None
) -> Tuple[str, np.ndarray]:
    """
    Convert numerical data into consistent units when possible.

    The GUI compares channels tagged as topography by assuming nanometers, so
    we convert pm/Angstrom/um/mm/m into nm.  Current, voltage, and frequency channels
    are converted into A/V/Hz respectively.  When a unit is unknown we simply
    return it unchanged.
    """
    if arr is None:
        return unit or "", np.array([])
    data = np.asarray(arr, dtype=float)
    if unit is None:
        return "", data
    key = str(unit).strip()
    key_lower = key.lower()
    info = _UNIT_NORMALIZATION.get(key_lower)
    if info:
        target_unit, factor = info
        data = data * factor
        return target_unit, data
    # Some headers embed units like "XScanRange[nm]" -> strip the suffix.
    bracket_idx = key_lower.find("[")
    if bracket_idx >= 0 and key_lower.endswith("]"):
        inner = key_lower[bracket_idx + 1 : -1]
        info = _UNIT_NORMALIZATION.get(inner)
        if info:
            target_unit, factor = info
            data = data * factor
            return target_unit, data
    return key, data


# --------------------------------------------------------------------------- #
# Internal helpers                                                           #
# --------------------------------------------------------------------------- #

def _split_key_value(line: str) -> Tuple[Optional[str], Optional[str]]:
    for sep in ("=", ":", "\t"):
        if sep in line:
            left, right = line.split(sep, 1)
            key = left.strip()
            value = right.strip()
            if key:
                return key, value
    return None, None


def _coerce_value(value: str) -> object:
    value = value.strip().strip('"')
    if not value:
        return ""
    try:
        if "." in value or "e" in value.lower():
            return float(value)
        return int(value)
    except Exception:
        return value


def _canonical_header_key(key: str) -> str:
    """
    Normalize header keys so GUI lookups stay reliable even when the text file
    mixes ``xPixel`` / ``XPixel`` / ``X Pixel`` spellings.
    """
    key = key.strip()
    key = key.replace(" ", "")
    key = key.replace("[nm]", "")
    key = key.replace("[um]", "")
    key = key.replace("[A]", "")
    return key


def _parse_inline_channels(header: Dict[str, object]) -> List[ChannelInfo]:
    inline: List[ChannelInfo] = []
    # Common pattern: FileName0, Caption0, PhysUnit0, etc.
    indices = set()
    for key in header.keys():
        suffix = _trailing_digits(key)
        if suffix is not None:
            indices.add(suffix)
    for idx in sorted(indices):
        info = ChannelInfo()
        info.file_name = str(header.get(f"FileName{idx}", ""))
        info.caption = str(header.get(f"Caption{idx}", ""))
        info.phys_unit = str(header.get(f"PhysUnit{idx}", ""))
        try:
            info.scale = float(header.get(f"Scale{idx}", 1.0))
        except Exception:
            info.scale = 1.0
        try:
            info.offset = float(header.get(f"Offset{idx}", 0.0))
        except Exception:
            info.offset = 0.0
        if info.file_name:
            inline.append(info)
    return inline


def _trailing_digits(key: str) -> Optional[int]:
    digits = []
    for ch in reversed(key):
        if ch.isdigit():
            digits.append(ch)
        else:
            break
    if not digits:
        return None
    try:
        return int("".join(reversed(digits)))
    except Exception:
        return None


def _load_ascii_grid(path: Path, count: int) -> Optional[np.ndarray]:
    try:
        data = np.loadtxt(path, dtype=float)
    except Exception:
        return None
    data = np.asarray(data, dtype=float)
    if data.size == 0:
        return None
    return data.reshape(-1)


def _load_binary_grid(path: Path, count: int, dtype) -> Optional[np.ndarray]:
    try:
        data = np.fromfile(path, dtype=dtype, count=count)
    except Exception:
        return None
    if data.size == 0:
        return None
    return data


def _load_tokenized_grid(path: Path, count: int) -> Optional[np.ndarray]:
    try:
        text = path.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        text = path.read_text(encoding="cp1252", errors="ignore")
    tokens: List[float] = []
    for token in text.replace(",", " ").split():
        try:
            tokens.append(float(token))
        except Exception:
            continue
        if len(tokens) >= count:
            break
    if not tokens:
        return None
    return np.asarray(tokens, dtype=float)


def _load_binary_with_inference(path: Path, count: int) -> Optional[np.ndarray]:
    """
    Attempt to read ``path`` as a binary grid by inferring a suitable dtype.
    Priority:
      1. Use filesize hints (bytes per sample) to choose candidates.
      2. Prefer integer dtypes for ``.int`` style files so controller counts are decoded correctly.
      3. Fall back to float32/float64.
    """
    candidates = _binary_dtype_candidates(path, count)
    for code in candidates:
        dtype = np.dtype(code).newbyteorder("<")
        arr = _load_binary_grid(path, count, dtype=dtype)
        if arr is None or arr.size == 0:
            continue
        return arr
    # as a last resort try generic floats
    for code in ("<f4", "<f8"):
        dtype = np.dtype(code).newbyteorder("<")
        arr = _load_binary_grid(path, count, dtype=dtype)
        if arr is not None and arr.size:
            return arr
    return None


def _binary_dtype_candidates(path: Path, target_count: int) -> List[str]:
    suffix = path.suffix.lower()
    prefer_int = suffix in {".int", ".ita", ".it", ".itm"}
    try:
        size = path.stat().st_size
    except Exception:
        size = 0
    approx = 0
    if size and target_count:
        approx = int(round(size / max(1, target_count)))
    order: List[str] = []
    seen = set()

    def _extend(items):
        for item in items:
            if item not in seen:
                seen.add(item)
                order.append(item)

    size_hints = {
        1: ("<u1", "<i1"),
        2: ("<i2", "<u2"),
        4: ("<i4", "<u4", "<f4"),
        8: ("<i8", "<u8", "<f8"),
    }
    if approx in size_hints:
        _extend(size_hints[approx])
    if prefer_int:
        _extend(("<i4", "<u4", "<i2", "<u2", "<i8", "<u8"))
        _extend(("<f4", "<f8"))
    else:
        _extend(("<f4", "<f8"))
        _extend(("<i4", "<u4", "<i2", "<u2", "<i8", "<u8"))
    if not order:
        _extend(("<f4", "<f8", "<i4", "<u4"))
    valid: List[str] = []
    for code in order:
        dtype = np.dtype(code)
        itemsize = dtype.itemsize or 1
        if size and (size % itemsize) not in (0, itemsize - 1):
            # tolerate padding but skip obvious mismatches
            if (size // itemsize) < target_count:
                continue
        valid.append(code)
    return valid or ["<f4", "<f8"]


_UNIT_NORMALIZATION: Dict[str, Tuple[str, float]] = {
    "pm": ("nm", 1e-3),
    "nm": ("nm", 1.0),
    "um": ("nm", 1e3),
    "mm": ("nm", 1e6),
    "m": ("nm", 1e9),
    "ang": ("nm", 0.1),
    "angstrom": ("nm", 0.1),
    "a": ("A", 1.0),
    "fa": ("A", 1e-15),
    "pa": ("A", 1e-12),
    "na": ("A", 1e-9),
    "ma": ("A", 1e-3),
    "ua": ("A", 1e-6),
    "v": ("V", 1.0),
    "mv": ("V", 1e-3),
    "kv": ("V", 1e3),
    "hz": ("Hz", 1.0),
    "khz": ("Hz", 1e3),
    "mhz": ("Hz", 1e6),
    "ghz": ("Hz", 1e9),
}


__all__ = [
    "parse_header",
    "read_channel_file",
    "normalize_unit_and_data",
]



