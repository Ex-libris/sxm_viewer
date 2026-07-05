"""Best-effort scan-timing extraction from a parsed header dict.

Used by the periodic-noise filter to map a known temporal frequency (e.g. a
500 Hz turbo pump harmonic) to the spatial frequency it would alias to in a
raster-scanned image. Returns None when timing can't be determined, rather
than guessing - callers fall back to blind peak detection in that case.
"""
from __future__ import annotations

from typing import Optional

# Anfatec/Omicron acquisition software doesn't have one universal header key
# for scan speed/timing across vendors and versions, and no real sample
# header was available to confirm the exact key used in practice - these are
# plausible candidates tried best-effort, not a confirmed source.
_ANFATEC_TIME_PER_LINE_KEYS = (
    "ScanTime",
    "TimePerLine",
    "TimePerLine[s]",
    "LineTime",
    "LineTime[s]",
)
_ANFATEC_SPEED_KEYS = (
    "ScanSpeed",
    "ScanSpeed[nm/s]",
    "XScanSpeed",
    "Speed",
)


def _safe_float(value) -> Optional[float]:
    try:
        result = float(value)
    except (TypeError, ValueError):
        return None
    if result <= 0:
        return None
    return result


def estimate_scan_timing(header: dict) -> Optional[dict]:
    """Return {"line_time_s", "pixel_dwell_s", "source"} or None.

    line_time_s is the time to acquire one scan line (fast-axis sweep);
    pixel_dwell_s is the time per pixel within that line.
    """
    if not header:
        return None
    xpix = _safe_float(header.get("xPixel"))
    ypix = _safe_float(header.get("yPixel"))
    if not xpix or not ypix:
        return None

    scan_time_fwd = _safe_float(header.get("ScanTimeForward[s]"))
    if scan_time_fwd is not None:
        # Nanonis's SCAN_TIME header field is already the time for one line
        # (forward/backward), not the whole frame - cross-checked against a
        # real file's own acq_time: lines * scan_time matched acq_time almost
        # exactly. Dividing by yPixel here would be a ~yPixel-times error.
        line_time = scan_time_fwd
        return {
            "line_time_s": line_time,
            "pixel_dwell_s": line_time / xpix,
            "source": "nanonis",
        }

    for key in _ANFATEC_TIME_PER_LINE_KEYS:
        line_time = _safe_float(header.get(key))
        if line_time is not None:
            return {
                "line_time_s": line_time,
                "pixel_dwell_s": line_time / xpix,
                "source": "anfatec",
            }

    x_range = _safe_float(header.get("XScanRange"))
    for key in _ANFATEC_SPEED_KEYS:
        speed = _safe_float(header.get(key))
        if speed is not None and x_range:
            line_time = x_range / speed
            return {
                "line_time_s": line_time,
                "pixel_dwell_s": line_time / xpix,
                "source": "anfatec",
            }

    return None


__all__ = ["estimate_scan_timing"]
