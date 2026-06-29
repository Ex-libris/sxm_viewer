"""Pure filter-pipeline logic extracted from the main window.

These methods take arrays / step dicts and return arrays or formatted
labels; none of them touch window or Qt state, so they can be unit-tested
in isolation by instantiating ``FilterController(None)``.

The main window keeps thin delegating stubs (``self.filter_controller.X``)
so existing call sites are unchanged.
"""
from __future__ import annotations

import numpy as np

from ...processing.filters import (
    FILTER_DEFINITIONS,
    flatten_remove_median,
    subtract_best_fit_plane,
    subtract_2nd_order_plane,
    line_flatten_image,
    gaussian_filter_image,
    highpass_filter,
    laplacian_filter_image,
    log_filter_image,
    histogram_equalize_image,
    clahe_filter_image,
)
from ..thumbnail_render import detect_valid_scan_region


class FilterController:
    """Encapsulates the pure filter-pipeline logic for the main window."""

    def __init__(self, viewer):
        self.viewer = viewer

    def _filter_action_label(self, filter_key):
        base_label = FILTER_DEFINITIONS.get(filter_key, {}).get("label", str(filter_key or "").title())
        return f"{base_label}..."

    def _normalize_preview_filter_steps(self, steps):
        if steps is None:
            return []
        if isinstance(steps, dict):
            return [steps]
        return [step for step in list(steps or []) if isinstance(step, dict)]

    def _filter_pipeline_label_from_steps(self, steps, default="Custom"):
        normalized = self._normalize_preview_filter_steps(steps)
        if not normalized:
            return ""
        labels = []
        for step in normalized:
            key = str(step.get("key") or "").strip()
            if not key:
                continue
            label = FILTER_DEFINITIONS.get(key, {}).get("label", key.replace("_", " ").title())
            labels.append(str(label).strip())
        if not labels:
            return str(default or "Custom")
        return " -> ".join(labels)

    def _filter_badge_text(self, steps):
        count = len(self._normalize_preview_filter_steps(steps))
        if count <= 1:
            return "F"
        return f"F{min(count, 9)}"

    def _filter_pipeline_tooltip(self, label, steps):
        summary = self._filter_pipeline_label_from_steps(steps)
        count = len(self._normalize_preview_filter_steps(steps))
        if not summary:
            return ""
        if label and label != summary:
            return f"Filter pipeline ({count} step{'s' if count != 1 else ''}): {label}\n{summary}"
        return f"Filter pipeline ({count} step{'s' if count != 1 else ''}): {summary}"

    def _apply_filter_pipeline(self, arr, steps):
        result = np.asarray(arr, dtype=float)
        for step in steps:
            result = self._run_filter_step_on_valid_region(result, step)
        return result

    def _run_filter_step_on_valid_region(self, arr, step):
        work = np.asarray(arr, dtype=float)
        if work.ndim != 2:
            return self._run_filter_step(work, step)
        try:
            region = detect_valid_scan_region(work)
        except Exception:
            region = None
        if not region:
            return self._run_filter_step(work, step)
        r0, r1 = region
        if r1 < r0:
            return self._run_filter_step(work, step)
        out = np.array(work, copy=True)
        try:
            filtered = self._run_filter_step(work[r0:r1 + 1, :], step)
        except Exception:
            filtered = self._run_filter_step(work, step)
            return np.asarray(filtered, dtype=float)
        try:
            out[r0:r1 + 1, :] = np.asarray(filtered, dtype=float)
        except Exception:
            return np.asarray(filtered, dtype=float)
        return out

    def _run_filter_step(self, arr, step):
        key = step.get('key')
        params = step.get('params', {})
        try:
            if key == 'flatten':
                axis = params.get('axis', 'both')
                return flatten_remove_median(arr, axis=axis)
            if key == 'tilt':
                return subtract_best_fit_plane(arr)
            if key == 'plane2':
                return subtract_2nd_order_plane(arr)
            if key == 'lowpass':
                sigma = params.get('sigma', 2.0)
                return gaussian_filter_image(arr, sigma)
            if key == 'highpass':
                sigma = params.get('sigma', 2.0)
                return highpass_filter(arr, sigma)
            if key == 'laplacian':
                sigma = params.get('sigma', FILTER_DEFINITIONS.get('laplacian', {}).get('default_sigma', 0.6))
                neighbors = params.get('neighbors', FILTER_DEFINITIONS.get('laplacian', {}).get('default_neighbors', 8))
                absolute = params.get('absolute', FILTER_DEFINITIONS.get('laplacian', {}).get('default_absolute', True))
                return laplacian_filter_image(arr, sigma=sigma, neighbors=neighbors, absolute=absolute)
            if key == 'log':
                epsilon = params.get('epsilon', FILTER_DEFINITIONS.get('log', {}).get('default_epsilon', 1e-3))
                return log_filter_image(arr, epsilon=epsilon)
            if key == 'histeq':
                return histogram_equalize_image(arr)
            if key == 'clahe':
                clip_limit = params.get('clip_limit', FILTER_DEFINITIONS.get('clahe', {}).get('default_clip_limit', 0.03))
                tile_size = params.get('tile_size', FILTER_DEFINITIONS.get('clahe', {}).get('default_tile_size', 8))
                return clahe_filter_image(arr, clip_limit=clip_limit, tile_size=tile_size)
            if key == 'line_flatten':
                axis = params.get('axis', FILTER_DEFINITIONS.get('line_flatten', {}).get('default_axis', 'row'))
                method = params.get('method', FILTER_DEFINITIONS.get('line_flatten', {}).get('default_method', 'median'))
                return line_flatten_image(arr, axis=axis, method=method)
        except Exception:
            pass
        return arr
