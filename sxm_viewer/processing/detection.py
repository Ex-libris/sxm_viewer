"""Detection helpers for tagging SXM files."""
from __future__ import annotations

from pathlib import Path

import numpy as np

from ..config import CH_SAMPLE_POINTS, CH_EQUALITY_TOL_NM


def _detect_dtype_for_file(path, expected_pixels):
    """
    Choose a dtype candidate for a binary channel file by filesize heuristics.
    Returns a numpy dtype (e.g. np.int16) or None on failure.
    This DOES NOT read the whole file; it only inspects the filesize.
    """
    candidate = [np.int16, np.uint16, np.int32, np.uint32, np.int64, np.float32, np.float64, np.uint8]
    filesize = Path(path).stat().st_size
    for dt in candidate:
        s = np.dtype(dt).itemsize
        # accept if filesize is at least expected_pixels * itemsize (some files may have padding)
        if filesize >= expected_pixels * s and (filesize % s == 0 or (filesize // s) >= expected_pixels):
            return dt
    # fallback: try float32
    return np.float32

def _sample_channel_values_for_tagging(file_key, header, fd, sample_count=CH_SAMPLE_POINTS):
    """Read a small selection of samples from a channel file for constant-height detection."""
    fname = fd.get("FileName")
    if not fname:
        return None
    bin_path = Path(file_key).parent / fname
    if not bin_path.exists():
        return None
    try:
        xpix = max(1, int(header.get('xPixel', 128)))
    except Exception:
        xpix = 128
    try:
        ypix = max(1, int(header.get('yPixel', xpix)))
    except Exception:
        ypix = xpix
    expected = xpix * ypix
    dtype = _detect_dtype_for_file(bin_path, expected)
    if dtype is None:
        dtype = np.float32
    try:
        mem = np.memmap(bin_path, dtype=np.dtype(dtype), mode='r')
    except Exception:
        return None
    try:
        total = int(mem.size)
        if total <= 0:
            return None
        count = max(1, min(sample_count, total))
        if total <= count:
            samples = np.asarray(mem[:total], dtype=float)
        else:
            idx = np.linspace(0, total - 1, count, dtype=np.int64)
            samples = np.asarray(mem[idx], dtype=float)
    finally:
        try:
            del mem
        except Exception:
            pass
    try:
        scale = float(fd.get('Scale', 1.0))
    except Exception:
        scale = 1.0
    try:
        offset = float(fd.get('Offset', 0.0))
    except Exception:
        offset = 0.0
    return samples * scale + offset

def header_indicates_constant(header):
    """Return 'CH' or 'CC' or None based on header textual indicators."""
    if not header: return None
    # combine keys & values into lowercase string for searching
    entries = [f"{k}:{str(v)}".lower() for k,v in header.items()]
    combined = " ".join(entries)
    # look for phrases that indicate constant-height or constant-current
    if any(kw in combined for kw in ('constant-height', 'constant height', 'constantheight', 'constheight', 'mode: constant', 'scanmode: constant', 'operationmode: constant')):
        return 'CH'
    if any(kw in combined for kw in ('constant-current', 'constant current', 'constantcurrent', 'feedback: current', 'mode: current', 'scanmode: current')):
        return 'CC'
    # some vendors write "constant" but not long form; ensure words appear in same entry
    for entry in entries:
        if 'constant' in entry:
            if 'height' in entry:
                return 'CH'
            if 'current' in entry:
                return 'CC'
    return None


def _find_topography_channel(fds):
    """
    Return index of the TRUE topographic channel or None if not found.
    Priority:
      1. Caption contains 'topo' (case insensitive)
      2. FileName contains 'topo'
      3. Caption contains 'height' but avoid sensor/feedback/setpoint channels
      4. PhysUnit looks like a length (nm/pm/ï¿½m/ï¿½/etc)
      5. Otherwise return None
    """
    if not fds:
        return None
    # step 1: caption has "topo"
    for i, fd in enumerate(fds):
        cap = (fd.get("Caption","") or "").lower()
        if "topo" in cap:
            return i
    # step 2: filename has "topo"
    for i, fd in enumerate(fds):
        fn = (fd.get("FileName","") or "").lower()
        if "topo" in fn:
            return i
    # step 3: caption has "height" but avoid sensor/feedback channels
    for i, fd in enumerate(fds):
        cap = (fd.get("Caption","") or "").lower()
        if "height" in cap and "sensor" not in cap and "feedback" not in cap and "setpoint" not in cap:
            return i
    # step 4: inspect physical units for length-like values
    def _looks_like_length_unit(unit):
        u = (unit or "").strip().lower()
        if not u:
            return False
        length_tokens = (
            "nm", "nanometer", "nanometre", "pm", "picometer", "picometre",
            "um", "micrometer", "micrometre",
            "ang", "angstrom", "angstroms", "aa", "meter", "metre"
        )
        current_tokens = ("pa", "ma", "na", "ua", "amp", "ampere", "a ")
        if any(tok in u for tok in current_tokens):
            return False
        return any(tok in u for tok in length_tokens)

    for i, fd in enumerate(fds):
        if _looks_like_length_unit(fd.get("PhysUnit")):
            return i
    return None

def filedesc_indicates_current_or_topo(fd):
    """Return 'current' or 'topo' or None based on FileDesc keys/Caption/FileName."""
    s = " ".join([str(fd.get(k,'')) for k in ('Caption','FileName','PhysUnit')]).lower()
    if any(tok in s for tok in ('current','i ', 'it_', 'lia', 'amp','pa','a ')):
        return 'current'
    if any(tok in s for tok in ('topo','height','z ')):
        return 'topo'
    return None


__all__ = [
    "_detect_dtype_for_file",
    "_sample_channel_values_for_tagging",
    "header_indicates_constant",
    "_find_topography_channel",
    "filedesc_indicates_current_or_topo",
]
