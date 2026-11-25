"""Qt-free helpers for animation export (frame order, timing, padding).

The dialog (`gui/animation_export.py`) renders one RGB array per image;
everything that turns those into a GIF / MP4 frame stream lives here so it
can be unit-tested without Qt or a viewer.
"""
from __future__ import annotations

import math

import numpy as np

LOOP_MODES = ("loop", "pingpong", "once")


def playback_sequence(n: int, mode: str = "loop") -> list[int]:
    """Frame indices in playback order. ``pingpong`` plays forward then back
    without repeating the end frames (0 1 2 3 2 1), so a looping GIF turns
    around smoothly."""
    n = max(0, int(n))
    seq = list(range(n))
    if mode == "pingpong" and n > 2:
        seq += list(range(n - 2, 0, -1))
    return seq


def frame_durations_ms(seq: list[int], fps: float, hold_last_s: float = 0.0,
                       last_index: int | None = None) -> list[float]:
    """Per-entry display time (ms) for a variable-timing format (GIF).
    Every occurrence of ``last_index`` (default: the highest index, i.e. the
    final image / the ping-pong turnaround) gets ``hold_last_s`` extra."""
    base = 1000.0 / max(float(fps), 1e-3)
    if not seq:
        return []
    last = max(seq) if last_index is None else last_index
    hold = max(0.0, float(hold_last_s)) * 1000.0
    return [base + (hold if i == last else 0.0) for i in seq]


def constant_rate_sequence(seq: list[int], fps: float, hold_last_s: float = 0.0,
                           last_index: int | None = None) -> list[int]:
    """The same playback for a constant-rate format (MP4): the hold is made
    by repeating the held frame ``round(hold * fps)`` extra times."""
    if not seq:
        return []
    last = max(seq) if last_index is None else last_index
    extra = int(round(max(0.0, float(hold_last_s)) * max(float(fps), 1e-3)))
    out = []
    for i in seq:
        out.append(i)
        if i == last and extra:
            out.extend([i] * extra)
    return out


def total_duration_s(n: int, fps: float, hold_last_s: float = 0.0, mode: str = "loop") -> float:
    """Length of one pass of the animation in seconds."""
    seq = playback_sequence(n, mode)
    return sum(frame_durations_ms(seq, fps, hold_last_s)) / 1000.0


def _to_rgb(frame) -> np.ndarray:
    a = np.asarray(frame)
    if a.ndim == 2:
        a = np.stack([a] * 3, axis=-1)
    if a.shape[-1] == 4:
        a = a[..., :3]
    if a.dtype != np.uint8:
        a = np.clip(a, 0, 255).astype(np.uint8)
    return a


def pad_to_common(frames, fill=(255, 255, 255), multiple: int = 1) -> list[np.ndarray]:
    """Centre every frame on a canvas of the largest width / height (rounded
    up to ``multiple`` - 2 for H.264's yuv420p), so all frames share one
    size without resampling the rendered pixels."""
    rgb = [_to_rgb(f) for f in frames]
    if not rgb:
        return []
    multiple = max(1, int(multiple))
    h = max(f.shape[0] for f in rgb)
    w = max(f.shape[1] for f in rgb)
    h = int(math.ceil(h / multiple) * multiple)
    w = int(math.ceil(w / multiple) * multiple)
    fill = np.asarray(fill, dtype=np.uint8).reshape(1, 1, 3)
    out = []
    for f in rgb:
        if f.shape[0] == h and f.shape[1] == w:
            out.append(f)
            continue
        canvas = np.empty((h, w, 3), dtype=np.uint8)
        canvas[...] = fill
        y0 = (h - f.shape[0]) // 2
        x0 = (w - f.shape[1]) // 2
        canvas[y0:y0 + f.shape[0], x0:x0 + f.shape[1]] = f
        out.append(canvas)
    return out


def shared_range(ranges) -> tuple[float, float] | None:
    """One colour range for a series from each frame's own (p1, p99):
    the lowest low and highest high, so no frame saturates."""
    vals = [(float(lo), float(hi)) for lo, hi in ranges
            if lo is not None and hi is not None and math.isfinite(lo) and math.isfinite(hi)]
    if not vals:
        return None
    lo = min(v[0] for v in vals)
    hi = max(v[1] for v in vals)
    if not hi > lo:
        return None
    return lo, hi
