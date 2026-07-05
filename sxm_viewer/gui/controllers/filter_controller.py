"""Pure filter-pipeline logic extracted from the main window.

These methods take arrays / step dicts and return arrays or formatted
labels; none of them touch window or Qt state, so they can be unit-tested
in isolation by instantiating ``FilterController(None)``.

The main window keeps thin delegating stubs (``self.filter_controller.X``)
so existing call sites are unchanged.
"""
from __future__ import annotations

import copy
from pathlib import Path

import numpy as np

from ..._shared import QtWidgets
from ...config import save_config
from ...data.io import normalize_unit_and_data
from ...processing.filters import (
    FILTER_DEFINITIONS,
    _gaussian_available,
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
    repair_bad_lines,
    remove_spikes,
)
from ..dialogs.filters import CustomFilterDialog, SingleFilterDialog
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
            if key == 'line_repair':
                ratio = params.get('ratio', FILTER_DEFINITIONS.get('line_repair', {}).get('default_ratio', 25.0))
                return repair_bad_lines(arr, ratio=ratio)
            if key == 'spike_removal':
                ratio = params.get('ratio', FILTER_DEFINITIONS.get('spike_removal', {}).get('default_ratio', 25.0))
                window = params.get('window', FILTER_DEFINITIONS.get('spike_removal', {}).get('default_window', 3))
                return remove_spikes(arr, ratio=ratio, window=window)
        except Exception:
            pass
        return arr

    # ---- view / pipeline-state helpers (no Qt; read viewer state via self.viewer) ----
    def _clone_filter_source_views(self, canvas, views):
        clone_view = getattr(canvas, "_clone_undo_view", None)
        cloned = []
        for view in list(views or []):
            try:
                if callable(clone_view):
                    cloned.append(clone_view(view))
                else:
                    cloned.append(copy.deepcopy(view))
            except Exception:
                cloned.append(view)
        return cloned

    def _canvas_filter_steps(self, canvas, view=None):
        target_view = view
        if target_view is None and canvas is not None:
            try:
                target_view = getattr(canvas, "active_view_and_axes", lambda: (None, None))()[0]
            except Exception:
                target_view = None
            if target_view is None:
                target_view = next(iter(getattr(canvas, "views", []) or []), None)
        steps = []
        if isinstance(target_view, dict):
            steps = self._normalize_preview_filter_steps(target_view.get("filter_steps"))
        return [copy.deepcopy(step) for step in steps]

    def _canvas_filter_label(self, canvas, view=None):
        target_view = view
        if target_view is None and canvas is not None:
            try:
                target_view = getattr(canvas, "active_view_and_axes", lambda: (None, None))()[0]
            except Exception:
                target_view = None
            if target_view is None:
                target_view = next(iter(getattr(canvas, "views", []) or []), None)
        label = ""
        if isinstance(target_view, dict):
            label = str(target_view.get("filter_label") or "").strip()
        if label:
            return label
        return self._filter_pipeline_label_from_steps(self._canvas_filter_steps(canvas, view=target_view))

    def _thumbnail_filter_steps(self, file_key):
        spec = (getattr(self.viewer, "thumbnail_filters", {}) or {}).get(str(file_key)) or {}
        steps = self._normalize_preview_filter_steps(spec.get("steps"))
        return [copy.deepcopy(step) for step in steps]

    def _thumbnail_filter_label(self, file_key):
        spec = (getattr(self.viewer, "thumbnail_filters", {}) or {}).get(str(file_key)) or {}
        label = str(spec.get("label") or "").strip()
        if label:
            return label
        return self._filter_pipeline_label_from_steps(spec.get("steps"))

    def _base_filter_image_from_views(self, views):
        # Use explicit None checks: `array or fallback` raises on a
        # multi-element numpy array ("truth value is ambiguous"), which
        # previously got swallowed and returned None for filtered views.
        try:
            if views:
                base = views[0].get("_filter_base_arr")
                if base is None:
                    base = views[0].get("arr")
                return base
        except Exception:
            return None
        return None

    def _normalize_filter_preview_clim(self, clim):
        try:
            if clim is None:
                return None
            lo, hi = clim
            lo = float(lo)
            hi = float(hi)
            if not np.isfinite(lo) or not np.isfinite(hi) or hi <= lo:
                return None
            return (lo, hi)
        except Exception:
            return None

    # ---- canvas pipeline application (reads/writes viewer state via self.viewer) ----
    def _apply_filters_to_array(self, file_path, arr):
        spec = self.viewer.thumbnail_filters.get(str(file_path))
        if not spec:
            return arr
        return self._apply_filter_pipeline(arr, spec.get('steps', []))

    def _set_filter_pipeline_on_canvas(self, canvas, steps, label=None, source_views=None, push_undo=False):
        if canvas is None:
            return
        base_views = source_views if source_views is not None else getattr(canvas, "views", None)
        if not base_views:
            return
        steps = self._normalize_preview_filter_steps(steps)
        if steps and not str(label or "").strip():
            label = self._filter_pipeline_label_from_steps(steps, default="Custom")
        if push_undo:
            try:
                canvas.push_undo_state("filter")
            except Exception:
                pass
        new_views = []
        for view in self._clone_filter_source_views(canvas, base_views):
            nv = dict(view)
            base = nv.get("_filter_base_arr")
            if base is None:
                try:
                    base = np.array(nv.get("arr"), copy=True)
                except Exception:
                    base = nv.get("arr")
                nv["_filter_base_arr"] = base
            if not steps:
                nv["arr"] = np.array(base, copy=True) if base is not None else nv.get("arr")
                nv.pop("filter_steps", None)
                nv.pop("filter_label", None)
                nv.pop("clim", None)  # drop stale clim from filtered data
            else:
                nv["arr"] = self._apply_filter_pipeline(base, steps) if base is not None else nv.get("arr")
                nv["filter_steps"] = copy.deepcopy(steps)
                nv["filter_label"] = label
                try:
                    clim = self.viewer._auto_preview_clim(
                        nv["arr"],
                        relative_zero=bool(nv.get("display_relative_zero", False)),
                    )
                except Exception:
                    clim = None
                if clim is not None:
                    nv["clim"] = clim
                else:
                    nv.pop("clim", None)
            new_views.append(nv)
        canvas.set_views(new_views, preserve_profiles=True)

    def _build_canvas_filter_preview_callback(self, canvas, source_views):
        def _preview(steps, label=None):
            if steps is None:
                self._restore_filter_views_on_canvas(canvas, source_views)
                return
            self._set_filter_pipeline_on_canvas(
                canvas,
                steps,
                label=label,
                source_views=source_views,
                push_undo=False,
            )
        return _preview

    def _restore_filter_views_on_canvas(self, canvas, source_views):
        if canvas is None or not source_views:
            return
        canvas.set_views(self._clone_filter_source_views(canvas, source_views), preserve_profiles=True)

    def _load_filter_base_array_for_path(self, focus_path):
        base_arr = None
        if not focus_path:
            return None
        try:
            focus_key = str(focus_path)
            header, fds = self.viewer.headers.get(focus_key, (None, None))
            if header and fds:
                idx = None
                if self.viewer.last_preview and str(self.viewer.last_preview[0]) == focus_key:
                    idx = int(self.viewer.last_preview[1])
                if idx is None:
                    idx = 0
                if 0 <= idx < len(fds):
                    fd = fds[idx]
                    arr = self.viewer._get_channel_array(focus_key, idx, header, fd)
                    base_arr = normalize_unit_and_data(arr, fd.get("PhysUnit", ""))[1]
        except Exception:
            base_arr = None
        return base_arr

    def _filter_preview_render_state(self, view=None):
        cmap_name = None
        clim = None
        if isinstance(view, dict):
            cmap_name = str(view.get("cmap") or "").strip() or None
            clim = self._normalize_filter_preview_clim(view.get("clim"))
        if not cmap_name:
            try:
                cmap_name = str(self.viewer.preview_cmap_combo.currentText() or "").strip() or None
            except Exception:
                cmap_name = None
        if not cmap_name:
            cmap_name = str(getattr(self.viewer, "preview_cmap", "viridis") or "viridis")
        return cmap_name, clim

    def _filter_preview_context_for_path(self, focus_path):
        preview_target = "selected image"
        preview_callback = None
        original_views = None
        preview_cmap_name = None
        preview_clim = None
        base_arr = self._load_filter_base_array_for_path(focus_path)
        try:
            preview_target = Path(str(focus_path)).name if focus_path else preview_target
        except Exception:
            preview_target = str(focus_path or preview_target)
        canvas = getattr(self.viewer, "preview_canvas", None)
        if (
            focus_path
            and canvas is not None
            and getattr(canvas, "views", None)
            and self.viewer.last_preview
            and str(self.viewer.last_preview[0]) == str(focus_path)
        ):
            original_views = self._clone_filter_source_views(canvas, canvas.views)
            base_arr = self._base_filter_image_from_views(original_views)
            preview_callback = self._build_canvas_filter_preview_callback(canvas, original_views)
            preview_target = self.viewer._friendly_view_title(original_views[0] if original_views else None, preview_target)
            preview_cmap_name, preview_clim = self._filter_preview_render_state(original_views[0] if original_views else None)
        if preview_cmap_name is None:
            preview_cmap_name, preview_clim = self._filter_preview_render_state(None)
        return base_arr, preview_callback, original_views, preview_target, preview_cmap_name, preview_clim

    def _single_filter_step_spec(
        self,
        filter_key,
        parent=None,
        base_image=None,
        preview_callback=None,
        preview_target_text="current image",
        preview_cmap_name="viridis",
        preview_clim=None,
        show_preview_thumbnail=True,
    ):
        if not filter_key:
            return None, None
        defaults = FILTER_DEFINITIONS.get(filter_key, {})
        if filter_key in ("highpass", "lowpass"):
            initial_params = self.viewer.config.get(f"{filter_key}_filter_params", {})
        elif filter_key == "laplacian":
            initial_params = self.viewer.config.get("laplacian_filter_params", {})
        else:
            initial_params = defaults
        dlg = SingleFilterDialog(
            parent=parent or self.viewer,
            filter_key=filter_key,
            base_image=base_image,
            apply_step_func=self._run_filter_step,
            preview_callback=preview_callback,
            initial_params=initial_params,
            preview_target_text=preview_target_text,
            preview_cmap_name=preview_cmap_name,
            preview_clim=preview_clim,
            show_preview_thumbnail=show_preview_thumbnail,
        )
        if dlg.exec_() != QtWidgets.QDialog.Accepted:
            return None, None
        step = dlg.current_step()
        label = dlg.current_step_label()
        params = dict(step.get("params") or {})
        if filter_key in ("highpass", "lowpass"):
            self.viewer.config[f"{filter_key}_filter_params"] = params
            save_config(self.viewer.config)
        elif filter_key == "laplacian":
            self.viewer.config["laplacian_filter_params"] = params
            save_config(self.viewer.config)
        return step, label

    def _populate_canvas_filter_menu(self, menu, canvas, view=None):
        """Populate a context menu with quick filter actions for a preview canvas."""
        if menu is None or canvas is None:
            return
        filt_menu = menu.addMenu("Filters")
        current_steps = self._canvas_filter_steps(canvas, view=view)
        current_summary = self._filter_pipeline_label_from_steps(current_steps)
        if current_summary:
            status_act = QtWidgets.QAction(f"Current: {current_summary}", filt_menu)
            status_act.setEnabled(False)
            filt_menu.addAction(status_act)
            filt_menu.addSeparator()
        for key, info in FILTER_DEFINITIONS.items():
            prefix = "Add step: " if current_steps else ""
            act = QtWidgets.QAction(f"{prefix}{self._filter_action_label(key)}", filt_menu)
            if info.get("needs_gaussian") and not _gaussian_available():
                act.setEnabled(False)
                act.setToolTip("Requires scipy or OpenCV.")
            act.triggered.connect(lambda _, k=key: self._apply_filter_to_canvas(canvas, filter_key=k))
            filt_menu.addAction(act)
        filt_menu.addSeparator()
        custom_label = "Edit custom pipeline..." if current_steps else "Custom pipeline..."
        filt_menu.addAction(custom_label, lambda: self._open_custom_filter_for_canvas(canvas))
        filt_menu.addAction("Clear filter", lambda: self._apply_filter_to_canvas(canvas, pipeline=[]))

    def _apply_filter_to_canvas(self, canvas, filter_key=None, pipeline=None, label=None):
        """Apply a filter pipeline to the views of a popup/preview canvas."""
        if not canvas or not getattr(canvas, "views", None):
            return
        steps = pipeline
        if steps is None and filter_key:
            original_views = self._clone_filter_source_views(canvas, canvas.views)
            base_arr = self._base_filter_image_from_views(original_views)
            preview_callback = self._build_canvas_filter_preview_callback(canvas, original_views)
            preview_cmap_name, preview_clim = self._filter_preview_render_state(original_views[0] if original_views else None)
            step, step_label = self._single_filter_step_spec(
                filter_key,
                parent=canvas,
                base_image=base_arr,
                preview_callback=preview_callback,
                preview_target_text=self.viewer._friendly_view_title(original_views[0] if original_views else None, "current image"),
                preview_cmap_name=preview_cmap_name,
                preview_clim=preview_clim,
                show_preview_thumbnail=False,
            )
            self._restore_filter_views_on_canvas(canvas, original_views)
            if step is None:
                return
            existing_steps = self._canvas_filter_steps(canvas)
            steps = list(existing_steps) + [step]
            label = label or self._filter_pipeline_label_from_steps(steps, default=step_label)
        self._set_filter_pipeline_on_canvas(canvas, steps, label=label, push_undo=True)

    def _open_custom_filter_for_canvas(self, canvas):
        if not canvas or not getattr(canvas, "views", None):
            return
        original_views = self._clone_filter_source_views(canvas, canvas.views)
        base_arr = self._base_filter_image_from_views(original_views)
        preview_cmap_name, preview_clim = self._filter_preview_render_state(original_views[0] if original_views else None)
        existing_steps = self._canvas_filter_steps(canvas)
        existing_label = self._canvas_filter_label(canvas)
        dialog_parent = None
        try:
            dialog_parent = canvas.window() if canvas is not None else None
        except Exception:
            dialog_parent = None
        if dialog_parent is None:
            dialog_parent = canvas or self.viewer
        dlg = CustomFilterDialog(
            dialog_parent,
            base_arr,
            self._run_filter_step,
            preview_callback=self._build_canvas_filter_preview_callback(canvas, original_views),
            preview_target_text=self.viewer._friendly_view_title(original_views[0] if original_views else None, "current image"),
            preview_cmap_name=preview_cmap_name,
            preview_clim=preview_clim,
            show_preview_thumbnail=False,
            initial_pipeline=existing_steps,
            initial_name=existing_label or "Custom",
        )
        try:
            dlg.raise_()
            dlg.activateWindow()
        except Exception:
            pass
        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            steps = dlg.pipeline_steps()
            label = dlg.pipeline_label()
            self._restore_filter_views_on_canvas(canvas, original_views)
            self._apply_filter_to_canvas(canvas, pipeline=steps, label=label)
            return
        self._restore_filter_views_on_canvas(canvas, original_views)

    def _apply_filter_to_paths(self, paths, filter_key=None, pipeline=None, label=None, focus_path=None):
        if not paths:
            return
        if len(paths) > 12:
            ret = QtWidgets.QMessageBox.question(self.viewer, "Filters", f"Apply filter to {len(paths)} images? This may use significant memory.",
                                                 QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No, QtWidgets.QMessageBox.No)
            if ret != QtWidgets.QMessageBox.Yes:
                return
        if filter_key and FILTER_DEFINITIONS.get(filter_key, {}).get('needs_gaussian') and not _gaussian_available():
            QtWidgets.QMessageBox.warning(self.viewer, "Filters", "Gaussian filters require scipy or OpenCV.")
            return
        if pipeline is None:
            base_arr, preview_callback, original_views, preview_target, preview_cmap_name, preview_clim = self._filter_preview_context_for_path(focus_path)
            step, spec_label = self._single_filter_step_spec(
                filter_key,
                parent=self.viewer,
                base_image=base_arr,
                preview_callback=preview_callback,
                preview_target_text=preview_target,
                preview_cmap_name=preview_cmap_name,
                preview_clim=preview_clim,
            )
            if original_views is not None:
                self._restore_filter_views_on_canvas(self.viewer.preview_canvas, original_views)
            if step is None:
                return
            existing_steps = self._thumbnail_filter_steps(focus_path)
            spec_steps = list(existing_steps) + [step]
            spec_label = self._filter_pipeline_label_from_steps(spec_steps, default=spec_label)
        else:
            spec_steps = pipeline
            spec_label = label or self._filter_pipeline_label_from_steps(spec_steps, default='Custom')
        path_keys = {str(Path(p)) for p in paths}
        for key in path_keys:
            steps_copy = [dict(step) for step in spec_steps]
            self.viewer.thumbnail_filters[key] = {'steps': steps_copy, 'label': spec_label}
        self.viewer._invalidate_thumbnail_cache(path_keys)
        self.viewer._invalidate_filtered_cache(path_keys)
        self.viewer.populate_thumbnails_for_channel(self.viewer.channel_dropdown.currentIndex())
        if self.viewer.last_preview and str(self.viewer.last_preview[0]) in path_keys:
            self.viewer.show_file_channel(self.viewer.last_preview[0], self.viewer.last_preview[1])

    def _clear_filter_for_paths(self, paths):
        changed = False
        path_keys = {str(Path(p)) for p in paths}
        for key in path_keys:
            if self.viewer.thumbnail_filters.pop(key, None) is not None:
                changed = True
        if changed:
            self.viewer._invalidate_thumbnail_cache(path_keys)
            self.viewer._invalidate_filtered_cache(path_keys)
            # Clear stored clim overrides for affected files so that
            # _resolve_preview_clim falls back to _auto_preview_clim
            # instead of returning a clim that was computed for the
            # now-removed filter.
            clim_map = getattr(self.viewer, "per_file_channel_clim", None) or {}
            for clim_key in list(clim_map.keys()):
                try:
                    file_key = str(clim_key[0]) if isinstance(clim_key, tuple) else ""
                except Exception:
                    continue
                if file_key in path_keys:
                    clim_map.pop(clim_key, None)
            self.viewer.populate_thumbnails_for_channel(self.viewer.channel_dropdown.currentIndex())
            if self.viewer.last_preview and str(self.viewer.last_preview[0]) in path_keys:
                self.viewer.show_file_channel(self.viewer.last_preview[0], self.viewer.last_preview[1])

    def _open_custom_filter_dialog(self, paths, focus_path):
        base_arr, preview_callback, original_views, preview_target, preview_cmap_name, preview_clim = self._filter_preview_context_for_path(focus_path)
        existing_steps = self._thumbnail_filter_steps(focus_path)
        existing_label = self._thumbnail_filter_label(focus_path)
        dlg = CustomFilterDialog(
            self.viewer,
            base_arr,
            self._run_filter_step,
            preview_callback=preview_callback,
            preview_target_text=preview_target,
            preview_cmap_name=preview_cmap_name,
            preview_clim=preview_clim,
            initial_pipeline=existing_steps,
            initial_name=existing_label or "Custom",
        )
        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            pipeline = dlg.pipeline_steps()
            if pipeline:
                if original_views is not None:
                    self._restore_filter_views_on_canvas(self.viewer.preview_canvas, original_views)
                self._apply_filter_to_paths(paths, pipeline=pipeline, label=dlg.pipeline_label())
                return
        if original_views is not None:
            self._restore_filter_views_on_canvas(self.viewer.preview_canvas, original_views)
