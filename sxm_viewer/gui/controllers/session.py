"""Session save/load controller for SXM Viewer."""
from __future__ import annotations

import json
import os
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
        os.makedirs(processed_dir, exist_ok=True)
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
            "canvas_state": canvas_state,
            "data_dir": data_dir.name,
        }
        return self._jsonify(payload)

    # ------------------------------------------------------------------
    def _apply_session_state(self, payload: dict, session_path: Path):
        viewer = self.viewer
        if not isinstance(payload, dict):
            return False
        image_folder = payload.get("image_folder") or ""
        if image_folder:
            try:
                viewer.load_folder(Path(image_folder))
            except Exception:
                pass
        spectra_folder = payload.get("spectra_folder") or ""
        if spectra_folder:
            try:
                viewer._set_spec_folder(Path(spectra_folder))
            except Exception:
                pass
        data_dir = payload.get("data_dir") or ""
        data_dir = session_path.parent / data_dir if data_dir else session_path.parent
        processed_dir = data_dir / "processed"
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
        try:
            if hasattr(viewer, "thumb_sort_combo") and ui.get("thumb_sort"):
                viewer.thumb_sort_combo.setCurrentText(ui.get("thumb_sort"))
            if hasattr(viewer, "thumb_filter_combo") and ui.get("thumb_filter"):
                viewer.thumb_filter_combo.setCurrentText(ui.get("thumb_filter"))
            if hasattr(viewer, "thumb_cmap_combo") and ui.get("thumb_cmap"):
                viewer.thumb_cmap_combo.setCurrentText(ui.get("thumb_cmap"))
                viewer.on_thumb_cmap_changed(viewer.thumb_cmap_combo.currentIndex())
            if hasattr(viewer, "preview_cmap_combo") and ui.get("preview_cmap"):
                viewer.preview_cmap_combo.setCurrentText(ui.get("preview_cmap"))
                viewer.on_preview_cmap_changed(viewer.preview_cmap_combo.currentIndex())
            if hasattr(viewer, "channel_dropdown"):
                idx = int(ui.get("channel_index", viewer.channel_dropdown.currentIndex()))
                viewer.channel_dropdown.setCurrentIndex(max(0, idx))
        except Exception:
            pass
        try:
            viewer._apply_mode(int(ui.get("mode", viewer.MODE_BROWSE)), remember=False)
        except Exception:
            pass
        try:
            viewer.on_show_spectra_toggled(bool(ui.get("show_spectra", True)))
            viewer.on_show_preview_spectra_toggled(bool(ui.get("show_preview_spectra", True)))
            viewer.on_show_matrix_markers_toggled(bool(ui.get("show_matrix_markers", True)))
            viewer.on_show_single_markers_toggled(bool(ui.get("show_single_markers", True)))
            viewer.on_compact_markers_toggled(bool(ui.get("compact_markers", True)))
            viewer.on_detail_dark_toggled(bool(ui.get("detail_dark_view", False)))
            viewer.on_detail_grid_toggled(bool(ui.get("detail_grid_view", False)))
            viewer.on_show_molecules_toggled(bool(ui.get("show_molecules", True)))
            viewer.on_relative_axes_toggled(bool(ui.get("relative_axes", False)))
            viewer.on_unit_relative_toggled(bool(ui.get("display_units_relative", False)))
            viewer.on_unit_display_toggled(bool(ui.get("display_units_si", False)))
            viewer.on_scale_bar_toggled(bool(ui.get("scale_bar", False)))
            viewer._on_toggle_preview_title(bool(ui.get("show_preview_title", True)))
        except Exception:
            pass
        try:
            viewer.populate_thumbnails_for_channel(viewer.channel_dropdown.currentIndex())
        except Exception:
            pass
        if pending_preview:
            try:
                viewer.show_file_channel(pending_preview[0], pending_preview[1])
            except Exception:
                pass
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
        canvas_state = payload.get("canvas_state")
        if canvas_state:
            try:
                viewer._on_open_canvas()
                win = viewer._canvas_window_ref()
                if win:
                    win._restore_state(canvas_state)
            except Exception:
                pass
        return True

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

