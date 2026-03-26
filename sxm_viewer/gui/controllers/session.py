"""Session save/load controller for SXM Viewer."""
from __future__ import annotations

import json
import os
import time
from pathlib import Path
from typing import Any

from ..._shared import QtWidgets, log_status, np


class SessionController:
    """Handles serialising/deserialising the viewer state to JSON session files."""

    def __init__(self, viewer):
        self.viewer = viewer

    # ------------------------------------------------------------------
    def save_session(self):
        viewer = self.viewer
        default_path = Path(getattr(viewer, "last_dir", ".")).joinpath("sxm_session.json")
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            viewer,
            "Save session",
            str(default_path),
            "SXM Session (*.json)",
        )
        if not path:
            return
        session_path = Path(path)
        if session_path.suffix.lower() != ".json":
            session_path = session_path.with_suffix(".json")
        try:
            payload = self._collect_session_state(session_path)
            session_path.parent.mkdir(parents=True, exist_ok=True)
            with open(session_path, "w", encoding="utf-8") as fh:
                json.dump(payload, fh, indent=2)
            log_status(f"Saved session to {session_path}")
        except Exception as exc:
            QtWidgets.QMessageBox.warning(viewer, "Save session", f"Unable to save session: {exc}")

    # ------------------------------------------------------------------
    def load_session(self):
        viewer = self.viewer
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            viewer,
            "Load session",
            str(Path(getattr(viewer, "last_dir", "."))),
            "SXM Session (*.json)",
        )
        if not path:
            return
        session_path = Path(path)
        try:
            with open(session_path, "r", encoding="utf-8") as fh:
                payload = json.load(fh)
            if self._apply_session_state(payload, session_path):
                log_status(f"Loaded session from {session_path}")
        except Exception as exc:
            QtWidgets.QMessageBox.warning(viewer, "Load session", f"Unable to load session: {exc}")

    # ------------------------------------------------------------------
    def _collect_session_state(self, session_path: Path):
        viewer = self.viewer
        data_dir = session_path.parent / f"{session_path.stem}_data"
        processed_dir = data_dir / "processed"
        views_dir = data_dir / "views"
        os.makedirs(processed_dir, exist_ok=True)
        os.makedirs(views_dir, exist_ok=True)
        processed = {}
        try:
            if viewer.last_preview:
                viewer._store_molecule_overlay(viewer.last_preview[0])
        except Exception:
            pass
        for key, data in (getattr(viewer, "_processed_views", {}) or {}).items():
            arr_files = {}
            arr_by_channel = data.get("arr_by_channel") or {}
            for ch_idx, arr in arr_by_channel.items():
                fname = f"{key}_ch{ch_idx}.npy"
                np.save(processed_dir / fname, np.asarray(arr))
                arr_files[str(ch_idx)] = str(Path("processed") / fname)
            processed[key] = {
                "source": data.get("source"),
                "op": data.get("op"),
                "label": data.get("label"),
                "channel_idx": data.get("channel_idx"),
                "header": data.get("header"),
                "fds": data.get("fds"),
                "arr_files": arr_files,
            }
        preview_state = {}
        canvas_state = None
        preview = getattr(viewer, "preview_canvas", None)
        if preview:
            try:
                preview_state["profile"] = preview.export_profile_state()
            except Exception:
                preview_state["profile"] = None
            try:
                preview_state["angle"] = preview.export_angle_state()
            except Exception:
                preview_state["angle"] = None
            try:
                preview_state["molecules"] = preview.export_molecule_state()
            except Exception:
                preview_state["molecules"] = None
            try:
                preview_state["scale_bar_pos"] = list(getattr(preview, "_scale_bar_pos", (0.94, 0.06)))
                preview_state["scale_bar_settings"] = dict(getattr(preview, "_scale_bar_settings", {}) or {})
            except Exception:
                pass
            preview_state["view_layout"] = getattr(preview, "_view_layout", "grid")
        preview_canvas_snapshot = self._capture_canvas_snapshot(
            preview,
            views_dir,
            prefix="preview",
            include_arrays=False,
        )
        win = viewer._canvas_window_ref()
        if win is not None and win.isVisible():
            try:
                canvas_state = win._capture_state()
            except Exception:
                canvas_state = None
        ui_state = {
            "thumb_size_px": getattr(viewer, "thumb_size_px", 160),
            "thumb_sort": viewer.thumb_sort_combo.currentText() if hasattr(viewer, "thumb_sort_combo") else None,
            "thumb_filter": viewer.thumb_filter_combo.currentText() if hasattr(viewer, "thumb_filter_combo") else None,
            "thumb_cmap": viewer.thumb_cmap_combo.currentText() if hasattr(viewer, "thumb_cmap_combo") else None,
            "preview_cmap": viewer.preview_cmap_combo.currentText() if hasattr(viewer, "preview_cmap_combo") else None,
            "channel_index": int(viewer.channel_dropdown.currentIndex()) if hasattr(viewer, "channel_dropdown") else 0,
            "mode": int(getattr(viewer, "current_mode", viewer.MODE_BROWSE)),
            "show_preview_title": bool(getattr(viewer, "show_preview_title", True)),
            "show_spectra": bool(getattr(viewer, "show_spectra", True)),
            "show_preview_spectra": bool(getattr(viewer, "show_preview_spectra", True)),
            "show_matrix_markers": bool(getattr(viewer, "show_matrix_markers", True)),
            "show_single_markers": bool(getattr(viewer, "show_single_markers", True)),
            "compact_markers": bool(getattr(viewer, "compact_markers", True)),
            "detail_dark_view": bool(getattr(viewer, "detail_dark_view", False)),
            "detail_grid_view": bool(getattr(viewer, "detail_grid_view", False)),
            "show_molecules": bool(getattr(viewer, "show_molecules", True)),
            "show_acquisition_overlay": bool(getattr(viewer, "show_acquisition_overlay", False)),
            "profile_label_mode": str(getattr(viewer, "profile_label_mode", "length") or "length"),
            "relative_axes": bool(getattr(viewer, "relative_axes", False)),
            "display_units_relative": bool(getattr(viewer, "display_units_relative", False)),
            "display_units_si": bool(getattr(viewer, "display_units_si", False)),
            "scale_bar": bool(viewer.scale_bar_cb.isChecked()) if hasattr(viewer, "scale_bar_cb") else False,
            "preview_locked": bool(getattr(viewer, "preview_locked", False)),
        }
        payload = {
            "version": 1,
            "image_folder": str(getattr(viewer, "last_dir", "") or ""),
            "spectra_folder": str(getattr(viewer, "spec_folder_path", "") or ""),
            "files": [str(p) for p in (getattr(viewer, "files", []) or [])],
            "processed": processed,
            "image_adjustments": getattr(viewer, "image_adjustments", {}),
            "thumbnail_filters": getattr(viewer, "thumbnail_filters", {}),
            "per_file_channel_cmap": getattr(viewer, "per_file_channel_cmap", {}),
            "extra_view_specs": getattr(viewer, "extra_view_specs", []),
            "tags": getattr(viewer, "tags", {}),
            "molecule_overlays": getattr(viewer, "molecule_overlays", {}),
            "thumb_multi_select": list(getattr(viewer, "thumb_multi_select", set()) or []),
            "selected_file_for_thumbs": getattr(viewer, "selected_file_for_thumbs", None),
            "last_preview": getattr(viewer, "last_preview", None),
            "ui": ui_state,
            "preview_state": preview_state,
            "preview_canvas_snapshot": preview_canvas_snapshot,
            "canvas_state": canvas_state,
            "data_dir": data_dir.name,
            "popup_canvases": self._capture_popup_snapshots(views_dir),
        }
        return self._jsonify(payload)

    # ------------------------------------------------------------------
    def _apply_session_state(self, payload: dict, session_path: Path):
        viewer = self.viewer
        if not isinstance(payload, dict):
            return False
        t0 = time.perf_counter()
        phase_t = t0
        image_folder = payload.get("image_folder") or ""
        if image_folder:
            try:
                viewer.load_folder(Path(image_folder))
            except Exception:
                pass
        load_folder_dt = time.perf_counter() - phase_t
        phase_t = time.perf_counter()
        spectra_folder = payload.get("spectra_folder") or ""
        if spectra_folder:
            try:
                viewer._set_spec_folder(Path(spectra_folder))
            except Exception:
                pass
        data_dir = payload.get("data_dir") or ""
        data_dir = session_path.parent / data_dir if data_dir else session_path.parent
        processed_dir = data_dir / "processed"
        views_dir = data_dir / "views"
        viewer._processed_views = {}
        for key, entry in (payload.get("processed") or {}).items():
            try:
                arr_by_channel = {}
                for ch_idx, rel_path in (entry.get("arr_files") or {}).items():
                    arr_path = processed_dir / Path(rel_path).name
                    if arr_path.exists():
                        arr_by_channel[int(ch_idx)] = np.load(arr_path, allow_pickle=False)
                header = entry.get("header") or {}
                fds = entry.get("fds") or []
                viewer._processed_views[str(key)] = {
                    "arr_by_channel": arr_by_channel,
                    "header": header,
                    "fds": fds,
                    "channel_idx": entry.get("channel_idx"),
                    "source": entry.get("source"),
                    "label": entry.get("label"),
                    "op": entry.get("op"),
                }
                viewer.headers[str(key)] = (header, fds)
            except Exception:
                continue
        session_files = []
        for fp in payload.get("files", []) or []:
            path_str = str(fp)
            if viewer._is_processed_key(path_str) and path_str in viewer._processed_views:
                session_files.append(Path(path_str))
            elif Path(path_str).exists():
                session_files.append(Path(path_str))
        if session_files:
            viewer.files = session_files
        viewer.image_adjustments = payload.get("image_adjustments") or {}
        viewer.thumbnail_filters = payload.get("thumbnail_filters") or {}
        viewer.per_file_channel_cmap = payload.get("per_file_channel_cmap") or {}
        viewer.extra_view_specs = payload.get("extra_view_specs") or []
        viewer.tags = payload.get("tags") or {}
        viewer.molecule_overlays = payload.get("molecule_overlays") or {}
        viewer.thumb_multi_select = set(payload.get("thumb_multi_select") or [])
        viewer.selected_file_for_thumbs = payload.get("selected_file_for_thumbs")
        pending_preview = payload.get("last_preview")
        viewer.last_preview = None
        ui = payload.get("ui") or {}
        self._apply_ui_state_fast(ui)
        apply_ui_dt = time.perf_counter() - phase_t
        phase_t = time.perf_counter()
        try:
            viewer.populate_thumbnails_for_channel(viewer.channel_dropdown.currentIndex())
        except Exception:
            pass
        thumbs_dt = time.perf_counter() - phase_t
        phase_t = time.perf_counter()
        if pending_preview:
            try:
                viewer.show_file_channel(pending_preview[0], pending_preview[1])
            except Exception:
                pass
        preview_build_dt = time.perf_counter() - phase_t
        phase_t = time.perf_counter()
        preview_state = payload.get("preview_state") or {}
        preview = getattr(viewer, "preview_canvas", None)
        if preview:
            try:
                if preview_state.get("view_layout"):
                    preview.set_view_layout(preview_state.get("view_layout"))
            except Exception:
                pass
            try:
                if preview_state.get("profile"):
                    preview.import_profile_state(preview_state.get("profile"), emit=False)
            except Exception:
                pass
            try:
                if preview_state.get("angle"):
                    preview.import_angle_state(preview_state.get("angle"))
            except Exception:
                pass
            try:
                if preview_state.get("molecules") and not viewer.molecule_overlays:
                    preview.import_molecule_state(preview_state.get("molecules"))
            except Exception:
                pass
            try:
                if preview_state.get("scale_bar_pos"):
                    preview._scale_bar_pos = tuple(preview_state.get("scale_bar_pos"))
                if preview_state.get("scale_bar_settings"):
                    preview._scale_bar_settings = dict(preview_state.get("scale_bar_settings") or {})
            except Exception:
                pass
        try:
            if preview_state.get("molecules") and viewer.last_preview:
                key = str(viewer.last_preview[0])
                if key not in viewer.molecule_overlays:
                    viewer.molecule_overlays[key] = preview_state.get("molecules")
        except Exception:
            pass
        preview_snapshot = payload.get("preview_canvas_snapshot")
        if preview_snapshot:
            self._restore_canvas_snapshot(
                preview,
                preview_snapshot,
                views_dir,
                viewer=viewer,
                require_view_match=True,
            )
        preview_restore_dt = time.perf_counter() - phase_t
        phase_t = time.perf_counter()
        canvas_state = payload.get("canvas_state")
        if canvas_state:
            try:
                viewer._on_open_canvas()
                win = viewer._canvas_window_ref()
                if win:
                    win._restore_state(canvas_state)
            except Exception:
                pass
        canvas_window_dt = time.perf_counter() - phase_t
        phase_t = time.perf_counter()
        popup_defs = payload.get("popup_canvases") or []
        popup_stats = {"count": 0, "elapsed": 0.0, "arrays": 0.0, "spawn": 0.0, "state": 0.0, "show": 0.0}
        if popup_defs:
            popup_stats = self._restore_popup_canvases(popup_defs, views_dir)
        total_dt = time.perf_counter() - t0
        try:
            popup_count = int(popup_stats.get("count", 0))
            popup_elapsed = float(popup_stats.get("elapsed", time.perf_counter() - phase_t))
            popup_tail = f" | popups {popup_count} in {popup_elapsed:.2f}s"
            if popup_count:
                popup_tail += " [arrays %.2fs | spawn %.2fs | state %.2fs | show %.2fs]" % (
                    float(popup_stats.get("arrays", 0.0)),
                    float(popup_stats.get("spawn", 0.0)),
                    float(popup_stats.get("state", 0.0)),
                    float(popup_stats.get("show", 0.0)),
                )
            log_status(
                "[Session] load %.2fs | folder %.2fs | ui %.2fs | thumbs %.2fs | preview %.2fs + %.2fs | canvas %.2fs%s"
                % (
                    total_dt,
                    load_folder_dt,
                    apply_ui_dt,
                    thumbs_dt,
                    preview_build_dt,
                    preview_restore_dt,
                    canvas_window_dt,
                    popup_tail,
                )
            )
        except Exception:
            pass
        return True

    @staticmethod
    def _set_checked_silent(widget, checked):
        if widget is None:
            return
        try:
            prev = widget.blockSignals(True)
            widget.setChecked(bool(checked))
            widget.blockSignals(prev)
        except Exception:
            pass

    @staticmethod
    def _set_current_text_silent(widget, text):
        if widget is None or text in (None, ""):
            return
        try:
            prev = widget.blockSignals(True)
            widget.setCurrentText(str(text))
            widget.blockSignals(prev)
        except Exception:
            pass

    @staticmethod
    def _set_current_index_silent(widget, index):
        if widget is None or index is None:
            return
        try:
            prev = widget.blockSignals(True)
            widget.setCurrentIndex(int(index))
            widget.blockSignals(prev)
        except Exception:
            pass

    def _apply_ui_state_fast(self, ui: dict):
        viewer = self.viewer
        if not isinstance(ui, dict):
            ui = {}
        try:
            viewer.thumb_size_px = int(ui.get("thumb_size_px", getattr(viewer, "thumb_size_px", 160)) or getattr(viewer, "thumb_size_px", 160))
        except Exception:
            pass
        self._set_current_text_silent(getattr(viewer, "thumb_sort_combo", None), ui.get("thumb_sort"))
        self._set_current_text_silent(getattr(viewer, "thumb_filter_combo", None), ui.get("thumb_filter"))
        self._set_current_text_silent(getattr(viewer, "thumb_cmap_combo", None), ui.get("thumb_cmap"))
        self._set_current_text_silent(getattr(viewer, "preview_cmap_combo", None), ui.get("preview_cmap"))
        try:
            if hasattr(viewer, "thumb_cmap_combo"):
                viewer.thumb_cmap = viewer.thumb_cmap_combo.currentText() or getattr(viewer, "thumb_cmap", "")
        except Exception:
            pass
        try:
            if hasattr(viewer, "preview_cmap_combo"):
                viewer.preview_cmap = viewer.preview_cmap_combo.currentText() or getattr(viewer, "preview_cmap", "")
        except Exception:
            pass
        if hasattr(viewer, "channel_dropdown"):
            try:
                idx = int(ui.get("channel_index", viewer.channel_dropdown.currentIndex()))
            except Exception:
                idx = viewer.channel_dropdown.currentIndex()
            idx = max(0, min(idx, max(0, viewer.channel_dropdown.count() - 1)))
            self._set_current_index_silent(viewer.channel_dropdown, idx)
        try:
            viewer._apply_mode(int(ui.get("mode", viewer.MODE_BROWSE)), remember=False)
        except Exception:
            pass

        viewer.show_preview_title = bool(ui.get("show_preview_title", getattr(viewer, "show_preview_title", True)))
        viewer.show_spectra = bool(ui.get("show_spectra", getattr(viewer, "show_spectra", True)))
        viewer.show_preview_spectra = bool(ui.get("show_preview_spectra", getattr(viewer, "show_preview_spectra", True)))
        viewer.show_matrix_markers = bool(ui.get("show_matrix_markers", getattr(viewer, "show_matrix_markers", True)))
        viewer.show_single_markers = bool(ui.get("show_single_markers", getattr(viewer, "show_single_markers", True)))
        viewer.compact_markers = bool(ui.get("compact_markers", getattr(viewer, "compact_markers", True)))
        viewer.detail_dark_view = bool(ui.get("detail_dark_view", getattr(viewer, "detail_dark_view", False)))
        viewer.detail_grid_view = bool(ui.get("detail_grid_view", getattr(viewer, "detail_grid_view", False)))
        viewer.relative_axes = bool(ui.get("relative_axes", getattr(viewer, "relative_axes", False)))
        viewer.display_units_relative = bool(ui.get("display_units_relative", getattr(viewer, "display_units_relative", False)))
        viewer.display_units_si = bool(ui.get("display_units_si", getattr(viewer, "display_units_si", False)))
        viewer.preview_locked = bool(ui.get("preview_locked", getattr(viewer, "preview_locked", False)))
        viewer.show_molecules = bool(ui.get("show_molecules", getattr(viewer, "show_molecules", True)))
        viewer.show_acquisition_overlay = bool(ui.get("show_acquisition_overlay", getattr(viewer, "show_acquisition_overlay", False)))
        profile_label_mode = str(ui.get("profile_label_mode", getattr(viewer, "profile_label_mode", "length")) or "length").strip().lower()
        if profile_label_mode not in {"length", "full", "hidden"}:
            profile_label_mode = "length"
        viewer.profile_label_mode = profile_label_mode

        self._set_checked_silent(getattr(viewer, "unit_display_cb", None), viewer.display_units_si)
        self._set_checked_silent(getattr(viewer, "unit_relative_cb", None), viewer.display_units_relative)
        self._set_checked_silent(getattr(viewer, "relative_axes_cb", None), viewer.relative_axes)
        self._set_checked_silent(getattr(viewer, "show_spectra_cb", None), viewer.show_preview_spectra)
        self._set_checked_silent(getattr(viewer, "scale_bar_cb", None), bool(ui.get("scale_bar", False)))
        self._set_checked_silent(getattr(viewer, "preview_lock_cb", None), viewer.preview_locked)

        for action_name, value in (
            ("spectro_overlay_act", viewer.show_spectra),
            ("matrix_markers_act", viewer.show_matrix_markers),
            ("single_markers_act", viewer.show_single_markers),
            ("compact_markers_act", viewer.compact_markers),
            ("detail_dark_act", viewer.detail_dark_view),
            ("detail_grid_act", viewer.detail_grid_view),
            ("molecules_act", viewer.show_molecules),
            ("acquisition_overlay_act", viewer.show_acquisition_overlay),
        ):
            self._set_checked_silent(getattr(viewer, action_name, None), value)
        for key, action in (getattr(viewer, "profile_label_actions", {}) or {}).items():
            self._set_checked_silent(action, key == profile_label_mode)
        try:
            if hasattr(viewer, "preview_detach_btn"):
                viewer.preview_detach_btn.setEnabled(not viewer.preview_locked)
        except Exception:
            pass

        try:
            viewer._apply_detail_view_theme()
        except Exception:
            pass
        try:
            if viewer.show_spectra or viewer.show_preview_spectra or viewer.show_matrix_markers or viewer.show_single_markers:
                if not getattr(viewer, "_spectros_loaded", False):
                    viewer.ensure_spectros_loaded(refresh=False)
                else:
                    viewer._update_spectro_stats_label()
            else:
                viewer._clear_multi_spec_selection()
                viewer._update_spectro_stats_label()
        except Exception:
            pass
        try:
            options = viewer._canvas_display_state_from_canvas(getattr(viewer, "preview_canvas", None))
            options["show_molecules"] = viewer.show_molecules
            options["show_acquisition_overlay"] = viewer.show_acquisition_overlay
            options["scale_bar_enabled"] = bool(ui.get("scale_bar", False))
            options["relative_axes_override"] = viewer.relative_axes
            options["show_title"] = viewer.show_preview_title
            viewer._apply_canvas_display_options(options, source_canvas=getattr(viewer, "preview_canvas", None), persist=False)
        except Exception:
            pass
        try:
            preview = getattr(viewer, "preview_canvas", None)
            if preview is not None:
                preview._profile_label_mode = profile_label_mode
                preview._notify_views_callback()
        except Exception:
            pass

    # ------------------------------------------------------------------
    def _capture_popup_snapshots(self, views_dir: Path):
        viewer = self.viewer
        canvases = list(getattr(viewer, "_popup_canvases", []) or [])
        snapshots = []
        for idx, canvas in enumerate(canvases):
            snap = self._capture_canvas_snapshot(canvas, views_dir, prefix=f"popup{idx}", include_arrays=True)
            if not snap:
                continue
            dlg = None
            try:
                dlg = canvas.parent()
            except Exception:
                dlg = None
            if dlg is not None:
                try:
                    if hasattr(dlg, "isVisible") and not dlg.isVisible():
                        continue
                except Exception:
                    pass
                try:
                    geo = dlg.geometry()
                    snap["window_geometry"] = [geo.x(), geo.y(), geo.width(), geo.height()]
                except Exception:
                    pass
                try:
                    snap["window_title"] = dlg.windowTitle()
                except Exception:
                    pass
            snapshots.append(snap)
        return snapshots

    def _capture_canvas_snapshot(self, canvas, views_dir: Path, prefix: str, include_arrays: bool):
        if canvas is None:
            return None
        snapshot = {
            "view_layout": getattr(canvas, "_view_layout", "grid"),
            "relative_axes_override": getattr(canvas, "_relative_axes_override", None),
            "scale_bar_enabled": bool(getattr(canvas, "scale_bar_enabled", False)),
            "show_title": bool(getattr(canvas, "_show_title", True)),
            "show_acquisition_overlay": bool(getattr(canvas, "_show_acquisition_overlay", False)),
            "view_font_scale": float(getattr(canvas, "_view_font_scale", 1.0) or 1.0),
            "plot_font_family": str(getattr(canvas, "_font_family", "") or ""),
            "plot_font_bold": bool(getattr(canvas, "_plot_font_bold", False)),
            "plot_font_italic": bool(getattr(canvas, "_plot_font_italic", False)),
            "plot_font_underline": bool(getattr(canvas, "_plot_font_underline", False)),
            "profile_label_mode": str(getattr(canvas, "_profile_label_mode", "length") or "length"),
            "profile_state": self._safe_canvas_call(canvas, "export_profile_state"),
            "angle_state": self._safe_canvas_call(canvas, "export_angle_state"),
            "molecule_state": self._safe_canvas_call(canvas, "export_molecule_state"),
            "scale_bar_pos": list(getattr(canvas, "_scale_bar_pos", (0.94, 0.06))),
            "scale_bar_settings": dict(getattr(canvas, "_scale_bar_settings", {}) or {}),
            "views": [],
            "zoom": [],
        }
        pipeline, label = self._view_filter_spec(canvas)
        snapshot["filter_pipeline"] = pipeline
        snapshot["filter_label"] = label
        for idx, view in enumerate(getattr(canvas, "views", []) or []):
            serialized = self._serialize_view_for_session(view, views_dir, f"{prefix}_v{idx}", include_arrays)
            snapshot["views"].append(serialized)
        try:
            snapshot["zoom"] = canvas.export_zoom_states()
        except Exception:
            snapshot["zoom"] = []
        return snapshot

    @staticmethod
    def _view_filter_spec(canvas):
        pipeline = None
        label = None
        for view in getattr(canvas, "views", []) or []:
            steps = view.get("filter_steps")
            if steps:
                pipeline = steps
                label = view.get("filter_label")
                break
        return pipeline, label

    def _serialize_view_for_session(self, view: dict, views_dir: Path, label: str, include_arrays: bool):
        if not view:
            return {}
        state = {}
        arr = view.get("arr")
        if include_arrays:
            arr_name = f"{label}.npy"
            if arr is not None:
                try:
                    np.save(views_dir / arr_name, np.asarray(arr), allow_pickle=False)
                    state["arr_file"] = arr_name
                except Exception:
                    state["arr_file"] = None
            else:
                state["arr_file"] = None
        for key, val in view.items():
            if key == "arr":
                continue
            if isinstance(key, str) and key.startswith("_"):
                continue
            state[key] = val
        state["session_signature"] = self._view_signature(view)
        return state

    @staticmethod
    def _view_signature(view: dict):
        meta = (view or {}).get("meta") or {}
        file_path = meta.get("file_path") or meta.get("path") or ""
        channel = meta.get("channel_index")
        try:
            channel = int(channel) if channel is not None else None
        except Exception:
            channel = None
        return {
            "file": str(file_path),
            "channel": channel,
            "crop_sequence": view.get("crop_sequence"),
            "title": view.get("title"),
        }

    @staticmethod
    def _session_signature_key(signature):
        if not signature:
            return None
        return (
            signature.get("file"),
            signature.get("channel"),
            signature.get("crop_sequence"),
            signature.get("title"),
        )

    @staticmethod
    def _safe_canvas_call(canvas, method_name: str):
        if canvas is None:
            return None
        try:
            method = getattr(canvas, method_name)
        except AttributeError:
            return None
        try:
            return method()
        except Exception:
            return None

    def _restore_canvas_snapshot(
        self,
        canvas,
        snapshot: dict,
        views_dir: Path,
        viewer=None,
        require_view_match: bool = False,
    ):
        if canvas is None or not snapshot:
            return
        snapshot_views = snapshot.get("views") or []
        snapshot_has_arrays = any(bool(entry.get("arr_file")) for entry in snapshot_views)
        if viewer is not None and hasattr(viewer, "_apply_canvas_style_snapshot"):
            try:
                viewer._apply_canvas_style_snapshot(
                    canvas,
                    {
                        "plot_typography": {
                            "family": snapshot.get("plot_font_family") or getattr(canvas, "_font_family", ""),
                            "bold": bool(snapshot.get("plot_font_bold", getattr(canvas, "_plot_font_bold", False))),
                            "italic": bool(snapshot.get("plot_font_italic", getattr(canvas, "_plot_font_italic", False))),
                            "underline": bool(snapshot.get("plot_font_underline", getattr(canvas, "_plot_font_underline", False))),
                        },
                        "view_font_scale": float(snapshot.get("view_font_scale", getattr(canvas, "_view_font_scale", 1.0)) or 1.0),
                        "display_options": {
                            "show_ticks": bool(getattr(canvas, "_show_ticks", True)),
                            "show_colorbar": bool(getattr(canvas, "_show_colorbar", True)),
                            "colorbar_orientation": str(getattr(canvas, "_colorbar_orientation", "vertical") or "vertical"),
                            "show_title": bool(snapshot.get("show_title", getattr(canvas, "_show_title", True))),
                            "show_acquisition_overlay": bool(snapshot.get("show_acquisition_overlay", getattr(canvas, "_show_acquisition_overlay", False))),
                            "show_shortcut_hint": bool(getattr(canvas, "_show_shortcut_hint", True)),
                            "show_profile_overlays": bool(getattr(canvas, "_show_profile_overlays", True)),
                            "show_angle_overlays": bool(getattr(canvas, "_show_angle_overlays", True)),
                            "show_molecules": bool(getattr(canvas, "show_molecules", True)),
                            "scale_bar_enabled": bool(snapshot.get("scale_bar_enabled", getattr(canvas, "scale_bar_enabled", False))),
                            "frame_fill_mode": bool(getattr(canvas, "_frame_fill_mode", False)),
                            "relative_axes_override": snapshot.get("relative_axes_override", getattr(canvas, "_relative_axes_override", None)),
                            "view_layout": snapshot.get("view_layout", getattr(canvas, "_view_layout", "grid")),
                        },
                    },
                    notify=False,
                    redraw=True,
                )
            except Exception:
                pass
        profile_label_mode = snapshot.get("profile_label_mode")
        if profile_label_mode is not None:
            try:
                canvas._profile_label_mode = str(profile_label_mode or "length")
            except Exception:
                pass
        sb_pos = snapshot.get("scale_bar_pos")
        if sb_pos:
            try:
                canvas._scale_bar_pos = tuple(sb_pos)
            except Exception:
                pass
        sb_settings = snapshot.get("scale_bar_settings")
        if sb_settings:
            try:
                canvas._scale_bar_settings = dict(sb_settings)
            except Exception:
                pass
        permit_view_state = True
        if require_view_match:
            permit_view_state = self._canvas_views_match_snapshot(canvas, snapshot_views)
        if permit_view_state:
            prof = snapshot.get("profile_state")
            if prof:
                try:
                    canvas.import_profile_state(prof, emit=False)
                except Exception:
                    pass
            angle_state = snapshot.get("angle_state")
            if angle_state:
                try:
                    canvas.import_angle_state(angle_state)
                except Exception:
                    pass
            molecules = snapshot.get("molecule_state")
            if molecules:
                try:
                    canvas.import_molecule_state(molecules)
                except Exception:
                    pass
            pipeline = snapshot.get("filter_pipeline")
            if pipeline and viewer is not None and not snapshot_has_arrays:
                try:
                    viewer._apply_filter_to_canvas(canvas, pipeline=pipeline, label=snapshot.get("filter_label"))
                except Exception:
                    pass
            self._restore_view_specific_state(canvas, snapshot_views)
            zoom_state = snapshot.get("zoom")
            if zoom_state:
                try:
                    canvas.apply_zoom_states(zoom_state)
                except Exception:
                    pass
        try:
            canvas._apply_view_font_scale()
        except Exception:
            pass

    def _restore_view_specific_state(self, canvas, entries):
        if not entries or canvas is None:
            return
        if not hasattr(canvas, "_session_signature_for_view"):
            return
        try:
            key_fn = canvas._signature_key
        except AttributeError:
            return
        view_map = {}
        for view in getattr(canvas, "views", []) or []:
            try:
                sig = canvas._session_signature_for_view(view)
                key = key_fn(sig)
            except Exception:
                key = None
            if key is None:
                continue
            view_map[key] = view
        changed = False
        for entry in entries:
            sig = entry.get("session_signature")
            if not sig:
                continue
            key = key_fn(sig)
            target = view_map.get(key)
            if not target:
                continue
            clim = entry.get("clim")
            if clim:
                try:
                    new_clim = tuple(clim)
                    if tuple(target.get("clim") or ()) != new_clim:
                        target["clim"] = new_clim
                        changed = True
                except Exception:
                    pass
            if entry.get("relative_axes") is not None:
                rel_axes = bool(entry.get("relative_axes"))
                if bool(target.get("relative_axes")) != rel_axes:
                    target["relative_axes"] = rel_axes
                    changed = True
        if changed:
            try:
                canvas._redraw()
            except Exception:
                canvas.draw_idle()

    def _canvas_view_keys(self, canvas):
        if canvas is None or not hasattr(canvas, "_session_signature_for_view"):
            return set()
        try:
            key_fn = canvas._signature_key
        except AttributeError:
            return set()
        keys = set()
        for view in getattr(canvas, "views", []) or []:
            try:
                sig = canvas._session_signature_for_view(view)
                key = key_fn(sig)
            except Exception:
                key = None
            if key is not None:
                keys.add(key)
        return keys

    def _canvas_views_match_snapshot(self, canvas, snapshot_views):
        if not snapshot_views:
            return True
        current = self._canvas_view_keys(canvas)
        if not current:
            return False
        snapshot_keys = set()
        for entry in snapshot_views:
            key = self._session_signature_key(entry.get("session_signature"))
            if key is not None:
                snapshot_keys.add(key)
        if not snapshot_keys:
            return False
        return not current.isdisjoint(snapshot_keys)

    def _build_view_from_snapshot_entry(self, entry: dict, views_dir: Path):
        if not entry:
            return None
        view = {}
        for key, val in entry.items():
            if key in ("arr_file", "session_signature"):
                continue
            view[key] = val
        arr_file = entry.get("arr_file")
        if arr_file:
            arr_path = views_dir / arr_file
            if not arr_path.exists():
                return None
            try:
                view["arr"] = np.load(arr_path, allow_pickle=False)
            except Exception:
                return None
        return view

    def _restore_popup_canvases(self, popup_defs, views_dir: Path):
        if not popup_defs:
            return {"count": 0, "elapsed": 0.0, "arrays": 0.0, "spawn": 0.0, "state": 0.0, "show": 0.0}
        viewer = self.viewer
        start = time.perf_counter()
        arrays_dt = 0.0
        spawn_dt = 0.0
        state_dt = 0.0
        show_dt = 0.0
        restored = []
        prev_display_sync = bool(getattr(viewer, "_canvas_display_syncing", False))
        viewer._canvas_display_syncing = True
        try:
            for snap in popup_defs:
                entries = snap.get("views") or []
                t_phase = time.perf_counter()
                built_views = []
                for entry in entries:
                    built = self._build_view_from_snapshot_entry(entry, views_dir)
                    if built is None:
                        built_views = []
                        break
                    built_views.append(built)
                arrays_dt += time.perf_counter() - t_phase
                if not built_views:
                    continue
                t_phase = time.perf_counter()
                try:
                    dlg = viewer._spawn_preview_popup(
                        built_views,
                        title=snap.get("window_title") or "Preview",
                        show_immediately=False,
                    )
                except Exception:
                    continue
                spawn_dt += time.perf_counter() - t_phase
                canvas = None
                try:
                    canvases = getattr(viewer, "_popup_canvases", [])
                    canvas = canvases[-1] if canvases else None
                except Exception:
                    canvas = None
                try:
                    if dlg:
                        dlg.setUpdatesEnabled(False)
                except Exception:
                    pass
                t_phase = time.perf_counter()
                if canvas:
                    self._restore_canvas_snapshot(canvas, snap, views_dir, viewer=viewer)
                state_dt += time.perf_counter() - t_phase
                geom = snap.get("window_geometry")
                has_geometry = False
                if dlg and geom and len(geom) == 4:
                    try:
                        x, y, w, h = [int(v) for v in geom]
                        dlg.setGeometry(x, y, w, h)
                        has_geometry = True
                    except Exception:
                        pass
                restored.append((dlg, has_geometry))
        finally:
            viewer._canvas_display_syncing = prev_display_sync
        shown = 0
        show_start = time.perf_counter()
        for dlg, has_geometry in restored:
            if dlg is None:
                continue
            try:
                dlg.setUpdatesEnabled(True)
            except Exception:
                pass
            try:
                if hasattr(dlg, "_resume_preview_resize"):
                    dlg._resume_preview_resize(force=not has_geometry)
                else:
                    dlg._preview_resize_paused = False
                dlg.show()
                shown += 1
            except Exception:
                continue
        show_dt += time.perf_counter() - show_start
        return {
            "count": shown,
            "elapsed": time.perf_counter() - start,
            "arrays": arrays_dt,
            "spawn": spawn_dt,
            "state": state_dt,
            "show": show_dt,
        }

    # ------------------------------------------------------------------
    @staticmethod
    def _jsonify(obj: Any):
        if isinstance(obj, dict):
            return {str(k): SessionController._jsonify(v) for k, v in obj.items()}
        if isinstance(obj, (list, tuple)):
            return [SessionController._jsonify(v) for v in obj]
        if isinstance(obj, set):
            return [SessionController._jsonify(v) for v in sorted(obj)]
        if isinstance(obj, Path):
            return str(obj)
        try:
            if isinstance(obj, np.ndarray):
                return obj.tolist()
        except Exception:
            pass
        try:
            if isinstance(obj, np.generic):
                return obj.item()
        except Exception:
            pass
        return obj
