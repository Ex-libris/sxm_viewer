"""Session save/load controller for SXM Viewer."""
from __future__ import annotations

import copy
import json
import os
import shutil
import time
from pathlib import Path
from typing import Any

from ..._shared import QtCore, QtGui, QtWidgets, log_status, np
from ...data.io import parse_header
from ..viewer import measurement as viewer_measurement


class SessionController:
    """Handles serialising/deserialising the viewer state to JSON session files."""

    def __init__(self, viewer):
        self.viewer = viewer

    def _pump_ui(self):
        app = QtWidgets.QApplication.instance()
        if app is None:
            return
        try:
            app.processEvents(
                QtCore.QEventLoop.ExcludeUserInputEvents | QtCore.QEventLoop.ExcludeSocketNotifiers,
                25,
            )
        except Exception:
            try:
                app.processEvents()
            except Exception:
                pass

    def _set_session_activity(self, message, detail="", value=None, stage="loading", hide_delay_ms=0):
        viewer = self.viewer
        cb = getattr(viewer, "_set_session_activity", None)
        if callable(cb):
            try:
                cb(
                    message,
                    detail=detail,
                    value=value,
                    stage=stage,
                    visible=True,
                    hide_delay_ms=hide_delay_ms,
                )
            except Exception:
                pass
        self._pump_ui()

    def _hide_session_activity(self):
        viewer = self.viewer
        cb = getattr(viewer, "_hide_session_activity", None)
        if callable(cb):
            try:
                cb()
            except Exception:
                pass

    def _deferred_popup_title(self, snapshot: dict):
        if not isinstance(snapshot, dict):
            return "Deferred pop-up"
        title = str(snapshot.get("window_title") or "").strip()
        if title:
            return title
        first_view = ((snapshot.get("views") or [{}]) or [{}])[0]
        meta = first_view.get("meta") or {}
        file_name = meta.get("file_name") or Path(str(first_view.get("path") or "")).name
        channel = meta.get("channel") or ""
        if file_name and channel:
            return f"{file_name} | {channel}"
        if file_name:
            return str(file_name)
        return "Deferred pop-up"

    def _register_deferred_popup(self, snapshot: dict, views_dir: Path):
        viewer = self.viewer
        if not snapshot:
            return None
        try:
            stored_snapshot = copy.deepcopy(snapshot)
        except Exception:
            stored_snapshot = dict(snapshot)
        try:
            viewer._deferred_popup_serial = int(getattr(viewer, "_deferred_popup_serial", 0)) + 1
        except Exception:
            viewer._deferred_popup_serial = 1
        entry = {
            "id": int(getattr(viewer, "_deferred_popup_serial", 1)),
            "snapshot": stored_snapshot,
            "views_dir": Path(views_dir),
            "title": self._deferred_popup_title(stored_snapshot),
        }
        entries = list(getattr(viewer, "_deferred_popup_entries", []) or [])
        entries.append(entry)
        viewer._deferred_popup_entries = entries
        try:
            viewer._refresh_deferred_popup_ui()
        except Exception:
            pass
        return entry

    def _remove_deferred_popup(self, entry_id):
        viewer = self.viewer
        if entry_id is None:
            return
        current = list(getattr(viewer, "_deferred_popup_entries", []) or [])
        viewer._deferred_popup_entries = [e for e in current if int(e.get("id", -1)) != int(entry_id)]
        try:
            viewer._refresh_deferred_popup_ui()
        except Exception:
            pass

    def _lookup_deferred_popup(self, entry_id=None, entry=None):
        viewer = self.viewer
        if isinstance(entry, dict):
            return entry
        if entry_id is None:
            return None
        for candidate in list(getattr(viewer, "_deferred_popup_entries", []) or []):
            try:
                if int(candidate.get("id", -1)) == int(entry_id):
                    return candidate
            except Exception:
                continue
        return None

    # ------------------------------------------------------------------
    def save_session(
        self,
        session_path=None,
        *,
        force_prompt: bool = False,
        prompt_if_missing: bool = True,
        record_recent: bool = True,
        set_current: bool = True,
        autosave: bool = False,
        quiet: bool = False,
    ):
        viewer = self.viewer
        chosen_path = session_path
        if force_prompt:
            chosen_path = None
        if chosen_path is None:
            chosen_path = getattr(viewer, "_current_session_path", None)
        if chosen_path is None and prompt_if_missing:
            default_path = Path(getattr(viewer, "last_dir", ".")).joinpath("sxm_session.json")
            path, _ = QtWidgets.QFileDialog.getSaveFileName(
                viewer,
                "Save session",
                str(default_path),
                "SXM Session (*.json)",
            )
            if not path:
                return None
            chosen_path = path
        if chosen_path is None:
            return None
        session_path = Path(chosen_path)
        if session_path.suffix.lower() != ".json":
            session_path = session_path.with_suffix(".json")
        try:
            payload = self._collect_session_state(session_path)
            session_path.parent.mkdir(parents=True, exist_ok=True)
            with open(session_path, "w", encoding="utf-8") as fh:
                json.dump(payload, fh, indent=2)
            if set_current:
                try:
                    viewer._current_session_path = str(session_path)
                except Exception:
                    pass
            if record_recent:
                record_recent_cb = getattr(viewer, "_record_recent_session", None)
                if callable(record_recent_cb):
                    try:
                        record_recent_cb(session_path)
                    except Exception:
                        pass
            if not autosave:
                log_status(f"Saved session to {session_path}")
                if not quiet:
                    show_saved_cb = getattr(viewer, "_show_saved_path_toast", None)
                    if callable(show_saved_cb):
                        try:
                            show_saved_cb("Session saved", session_path)
                        except Exception:
                            pass
            elif not quiet:
                log_status(f"Updated recovery session at {session_path}")
            return session_path
        except Exception as exc:
            if not quiet:
                QtWidgets.QMessageBox.warning(viewer, "Save session", f"Unable to save session: {exc}")
            return None

    def save_session_as(self):
        """Prompt for a new session target and save there."""
        return self.save_session(force_prompt=True)

    # ------------------------------------------------------------------
    def load_session(self, start_dir=None, session_path=None, *, record_recent: bool = True, set_current: bool = True):
        viewer = self.viewer
        if session_path is None:
            start_path = Path(start_dir) if start_dir is not None else Path(getattr(viewer, "last_dir", "."))
            path, _ = QtWidgets.QFileDialog.getOpenFileName(
                viewer,
                "Load session",
                str(start_path),
                "SXM Session (*.json)",
            )
            if not path:
                return
            session_path = Path(path)
        else:
            session_path = Path(session_path)
        try:
            record_recent_cb = getattr(viewer, "_record_recent_session", None)
            if record_recent and callable(record_recent_cb):
                try:
                    record_recent_cb(session_path)
                except Exception:
                    pass
            self._set_session_activity(
                "Opening session...",
                detail=session_path.name,
                value=6,
                stage="loading",
            )
            with open(session_path, "r", encoding="utf-8") as fh:
                payload = json.load(fh)
            self._set_session_activity(
                "Applying saved workspace...",
                detail="Restoring cached preview state and deferred pop-ups.",
                value=12,
                stage="loading",
            )
            prepare_cb = getattr(viewer, "_prepare_for_workspace_load", None)
            if callable(prepare_cb):
                try:
                    prepare_cb(kind="session")
                except Exception:
                    pass
            if self._apply_session_state(payload, session_path):
                if set_current:
                    try:
                        viewer._current_session_path = str(session_path)
                    except Exception:
                        pass
                log_status(f"Loaded session from {session_path}")
        except Exception as exc:
            self._set_session_activity(
                "Unable to load session",
                detail=str(exc),
                value=100,
                stage="error",
                hide_delay_ms=3500,
            )
            QtWidgets.QMessageBox.warning(viewer, "Load session", f"Unable to load session: {exc}")

    # ------------------------------------------------------------------
    def _collect_session_state(self, session_path: Path):
        viewer = self.viewer
        data_dir = session_path.parent / f"{session_path.stem}_data"
        processed_dir = data_dir / "processed"
        views_dir = data_dir / "views"
        thumbs_dir = data_dir / "thumbs"
        os.makedirs(processed_dir, exist_ok=True)
        os.makedirs(views_dir, exist_ok=True)
        os.makedirs(thumbs_dir, exist_ok=True)
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
                "lock_channel": data.get("lock_channel", True),
                "extent_raw": data.get("extent_raw"),
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
                preview_state["profile_dialog"] = viewer_measurement.export_profile_dialog_state(viewer)
            except Exception:
                preview_state["profile_dialog"] = None
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
            include_arrays=True,
        )
        session_headers = {}
        for key, pair in (getattr(viewer, "headers", {}) or {}).items():
            if getattr(viewer, "_is_processed_key", lambda _k: False)(key):
                continue
            try:
                header, fds = pair
            except Exception:
                continue
            session_headers[str(key)] = {
                "header": header or {},
                "fds": fds or [],
            }
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
            "show_spectro_miniatures": bool(getattr(viewer, "show_spectro_miniatures", False)),
            "spectro_share_overlapping_repeats": bool(getattr(viewer, "spectro_share_overlapping_repeats", False)),
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
            "headers": session_headers,
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
            "thumbnail_snapshot": self._capture_thumbnail_snapshot(thumbs_dir),
            "canvas_state": canvas_state,
            "data_dir": data_dir.name,
            "popup_canvases": self._capture_popup_snapshots(views_dir),
        }
        return self._jsonify(payload)

    # ------------------------------------------------------------------
    def _apply_cached_session_state(self, payload: dict, session_path: Path):
        viewer = self.viewer
        t0 = time.perf_counter()
        phase_t = t0
        try:
            viewer.clear_loaded_images()
        except Exception:
            pass
        self._set_session_activity(
            "Restoring session cache...",
            detail="Applying saved headers and cached image state.",
            value=18,
            stage="loading",
        )

        image_folder = payload.get("image_folder") or ""
        if image_folder:
            try:
                viewer.last_dir = Path(image_folder)
            except Exception:
                pass
            try:
                viewer.path_le.setText(str(image_folder))
            except Exception:
                pass
        spectra_folder = payload.get("spectra_folder") or ""
        if spectra_folder:
            try:
                viewer.spec_folder_path = Path(spectra_folder)
            except Exception:
                pass
            try:
                viewer.spec_folder_le.setText(str(spectra_folder))
            except Exception:
                pass

        data_dir = payload.get("data_dir") or ""
        data_dir = session_path.parent / data_dir if data_dir else session_path.parent
        processed_dir = data_dir / "processed"
        views_dir = data_dir / "views"
        thumbs_dir = data_dir / "thumbs"

        self._restore_headers_from_payload(payload.get("headers") or {})
        self._restore_processed_views_payload(payload.get("processed") or {}, processed_dir)
        session_files = self._session_files_from_payload(payload)
        if session_files:
            viewer.files = session_files
        try:
            viewer._build_image_timestamp_index()
            viewer._rebuild_frame_map_entries()
        except Exception:
            pass
        load_folder_dt = time.perf_counter() - phase_t

        phase_t = time.perf_counter()
        self._set_session_activity(
            "Restoring session cache...",
            detail="Applying saved UI state.",
            value=34,
            stage="loading",
        )
        self._apply_basic_session_payload(payload)
        self._restore_channel_dropdown_from_headers()
        ui = payload.get("ui") or {}
        self._apply_ui_state_fast(ui, load_spectros=False)
        apply_ui_dt = time.perf_counter() - phase_t

        phase_t = time.perf_counter()
        self._set_session_activity(
            "Restoring session cache...",
            detail="Painting cached thumbnails.",
            value=48,
            stage="loading",
        )
        thumb_snapshot = payload.get("thumbnail_snapshot") or {}
        self._restore_thumbnail_snapshot(thumb_snapshot, thumbs_dir)
        thumbs_dt = time.perf_counter() - phase_t

        phase_t = time.perf_counter()
        self._set_session_activity(
            "Restoring session cache...",
            detail="Painting cached preview.",
            value=64,
            stage="loading",
        )
        pending_preview = payload.get("last_preview")
        preview_state = payload.get("preview_state") or {}
        preview_snapshot = payload.get("preview_canvas_snapshot") or {}
        self._restore_preview_from_snapshot(pending_preview, preview_snapshot, preview_state, views_dir)
        preview_build_dt = 0.0
        self._restore_canvas_snapshot(
            getattr(viewer, "preview_canvas", None),
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
        popup_stats = {"count": 0, "elapsed": 0.0, "arrays": 0.0, "spawn": 0.0, "state": 0.0, "show": 0.0, "lazy": 0}
        if popup_defs:
            self._set_session_activity(
                "Restoring session cache...",
                detail=f"Deferring {len(popup_defs)} pop-up{'s' if len(popup_defs) != 1 else ''} to the Pop-ups menu.",
                value=78,
                stage="loading",
            )
            popup_stats = self._restore_popup_canvases(popup_defs, views_dir)
        total_dt = time.perf_counter() - t0

        try:
            viewer._update_toolbar_actions(bool(getattr(viewer, "files", []) or []))
        except Exception:
            pass
        self._schedule_session_hydration(payload, session_path)
        popup_lazy = int(popup_stats.get("lazy", 0) or 0)
        hydrate_detail = "Refreshing live thumbnails and preview in the background."
        if popup_lazy:
            hydrate_detail += f" {popup_lazy} pop-up{'s' if popup_lazy != 1 else ''} are ready in the Pop-ups menu."
        self._set_session_activity(
            "Session ready",
            detail=hydrate_detail,
            value=84,
            stage="hydrating",
        )

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
                popup_lazy = int(popup_stats.get("lazy", 0) or 0)
                if popup_lazy:
                    popup_tail += f" | lazy {popup_lazy}"
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

    # ------------------------------------------------------------------
    def _apply_session_state(self, payload: dict, session_path: Path):
        viewer = self.viewer
        if not isinstance(payload, dict):
            return False
        try:
            viewer._deferred_popup_entries = []
            viewer._deferred_popup_serial = 0
            viewer._refresh_deferred_popup_ui()
        except Exception:
            pass
        preview_snapshot = payload.get("preview_canvas_snapshot") or {}
        if (
            payload.get("headers")
            and any(bool(entry.get("arr_file")) for entry in (preview_snapshot.get("views") or []))
        ):
            return self._apply_cached_session_state(payload, session_path)
        t0 = time.perf_counter()
        phase_t = t0
        self._set_session_activity(
            "Restoring session...",
            detail="Loading source folder and rebuilding live state.",
            value=18,
            stage="loading",
        )
        image_folder = payload.get("image_folder") or ""
        if image_folder:
            try:
                viewer.load_folder(Path(image_folder))
            except Exception:
                pass
        load_folder_dt = time.perf_counter() - phase_t
        phase_t = time.perf_counter()
        self._set_session_activity(
            "Restoring session...",
            detail="Applying saved UI state.",
            value=38,
            stage="loading",
        )
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
                    "lock_channel": entry.get("lock_channel", True),
                    "source": entry.get("source"),
                    "extent_raw": entry.get("extent_raw"),
                    "label": entry.get("label"),
                    "op": entry.get("op"),
                }
                self._hydrate_collection_processed_view(str(key))
                restored = viewer._processed_views.get(str(key), {})
                viewer.headers[str(key)] = (
                    restored.get("header", header),
                    restored.get("fds", fds),
                )
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
        self._set_session_activity(
            "Restoring session...",
            detail="Rebuilding thumbnails.",
            value=56,
            stage="loading",
        )
        try:
            viewer.populate_thumbnails_for_channel(viewer.channel_dropdown.currentIndex())
        except Exception:
            pass
        thumbs_dt = time.perf_counter() - phase_t
        phase_t = time.perf_counter()
        self._set_session_activity(
            "Restoring session...",
            detail="Rebuilding preview.",
            value=70,
            stage="loading",
        )
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
        try:
            viewer_measurement.restore_profile_dialog_state(viewer, preview_state.get("profile_dialog"))
        except Exception:
            pass
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
        popup_stats = {"count": 0, "elapsed": 0.0, "arrays": 0.0, "spawn": 0.0, "state": 0.0, "show": 0.0, "lazy": 0}
        if popup_defs:
            self._set_session_activity(
                "Restoring session...",
                detail=f"Deferring {len(popup_defs)} pop-up{'s' if len(popup_defs) != 1 else ''} to the Pop-ups menu.",
                value=84,
                stage="loading",
            )
            popup_stats = self._restore_popup_canvases(popup_defs, views_dir)
        total_dt = time.perf_counter() - t0
        popup_lazy = int(popup_stats.get("lazy", 0) or 0)
        done_detail = "Session is ready."
        if popup_lazy:
            done_detail += f" {popup_lazy} pop-up{'s' if popup_lazy != 1 else ''} were deferred to the Pop-ups menu."
        self._set_session_activity(
            "Session fully restored",
            detail=done_detail,
            value=100,
            stage="complete",
            hide_delay_ms=2400,
        )
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
                popup_lazy = int(popup_stats.get("lazy", 0) or 0)
                if popup_lazy:
                    popup_tail += f" | lazy {popup_lazy}"
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

    def _apply_ui_state_fast(self, ui: dict, *, load_spectros: bool = True):
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
        viewer.show_spectro_miniatures = bool(ui.get("show_spectro_miniatures", getattr(viewer, "show_spectro_miniatures", False)))
        viewer.spectro_share_overlapping_repeats = bool(
            ui.get("spectro_share_overlapping_repeats", getattr(viewer, "spectro_share_overlapping_repeats", False))
        )
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
        self._set_checked_silent(getattr(viewer, "display_units_si_act", None), viewer.display_units_si)
        self._set_checked_silent(getattr(viewer, "display_units_relative_act", None), viewer.display_units_relative)
        self._set_checked_silent(getattr(viewer, "relative_axes_act", None), viewer.relative_axes)
        self._set_checked_silent(getattr(viewer, "show_spectra_cb", None), viewer.show_preview_spectra)
        self._set_checked_silent(getattr(viewer, "spectro_thumbnail_markers_cb", None), viewer.show_spectra)
        self._set_checked_silent(getattr(viewer, "spectro_preview_markers_cb", None), viewer.show_preview_spectra)
        self._set_checked_silent(getattr(viewer, "spectro_miniatures_cb", None), viewer.show_spectro_miniatures)
        self._set_checked_silent(getattr(viewer, "toolbar_spectro_thumb_btn", None), viewer.show_spectra)
        self._set_checked_silent(getattr(viewer, "toolbar_spectro_preview_btn", None), viewer.show_preview_spectra)
        self._set_checked_silent(getattr(viewer, "toolbar_spectro_miniatures_btn", None), viewer.show_spectro_miniatures)
        self._set_checked_silent(getattr(viewer, "scale_bar_cb", None), bool(ui.get("scale_bar", False)))
        self._set_checked_silent(getattr(viewer, "display_scale_bar_act", None), bool(ui.get("scale_bar", False)))
        self._set_checked_silent(getattr(viewer, "preview_lock_cb", None), viewer.preview_locked)
        self._set_checked_silent(getattr(viewer, "tools_preview_lock_act", None), viewer.preview_locked)

        for action_name, value in (
            ("spectro_overlay_act", viewer.show_spectra),
            ("preview_spectra_toggle_btn", viewer.show_spectra),
            ("spectro_miniatures_act", viewer.show_spectro_miniatures),
            ("toolbar_spectro_markers_act", viewer.show_spectra),
            ("toolbar_spectro_preview_act", viewer.show_preview_spectra),
            ("toolbar_spectro_miniatures_act", viewer.show_spectro_miniatures),
            ("toolbar_spectro_repeat_share_act", viewer.spectro_share_overlapping_repeats),
            ("highlight_glow_act", viewer.spectro_highlight_glow),
            ("toolbar_spectro_highlight_act", viewer.spectro_highlight_glow),
            ("matrix_markers_act", viewer.show_matrix_markers),
            ("single_markers_act", viewer.show_single_markers),
            ("compact_markers_act", viewer.compact_markers),
            ("toolbar_spectro_matrix_markers_act", viewer.show_matrix_markers),
            ("toolbar_spectro_single_markers_act", viewer.show_single_markers),
            ("toolbar_spectro_compact_markers_act", viewer.compact_markers),
            ("detail_dark_act", viewer.detail_dark_view),
            ("detail_grid_act", viewer.detail_grid_view),
            ("preview_grid_toggle_btn", viewer.detail_grid_view),
            ("molecules_act", viewer.show_molecules),
            ("preview_molecules_toggle_btn", viewer.show_molecules),
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
            if getattr(viewer, "spectros", None):
                viewer._assign_spectros_to_images()
        except Exception:
            pass
        if load_spectros:
            try:
                if (
                    viewer.show_spectra
                    or viewer.show_preview_spectra
                    or viewer.show_spectro_miniatures
                    or viewer.show_matrix_markers
                    or viewer.show_single_markers
                ):
                    if not getattr(viewer, "_spectros_loaded", False):
                        viewer.ensure_spectros_loaded(refresh=False)
                    else:
                        viewer._update_spectro_stats_label()
                else:
                    viewer._clear_multi_spec_selection()
                    viewer._update_spectro_stats_label()
            except Exception:
                pass
        else:
            try:
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
    def _capture_thumbnail_snapshot(self, thumbs_dir: Path):
        viewer = self.viewer
        items = []
        for idx, key in enumerate([str(fp) for fp in list(getattr(viewer, "current_thumb_files", []) or []) if str(fp)]):
            lbl = (getattr(viewer, "_thumb_labels", {}) or {}).get(key)
            if lbl is None:
                continue
            try:
                pix = lbl.pixmap()
            except Exception:
                pix = None
            if pix is None or pix.isNull():
                continue
            fname = f"thumb_{idx:03d}.png"
            try:
                pix.save(str(thumbs_dir / fname), "PNG")
            except Exception:
                continue
            try:
                dims = lbl.property("thumb_dims") or (pix.width(), pix.height())
            except Exception:
                dims = (pix.width(), pix.height())
            items.append({
                "file_key": key,
                "png_file": fname,
                "thumb_dims": [int(dims[0]), int(dims[1])],
            })
        if not items:
            return None
        selected = getattr(viewer, "selected_file_for_thumbs", None)
        return {
            "channel_index": int(getattr(viewer, "channel_dropdown", None).currentIndex()) if hasattr(viewer, "channel_dropdown") else 0,
            "thumb_size_px": int(getattr(viewer, "thumb_size_px", 160) or 160),
            "selected_file": str(selected) if selected else None,
            "items": items,
        }

    def _restore_headers_from_payload(self, headers_payload):
        viewer = self.viewer
        viewer.headers = {}
        for key, entry in (headers_payload or {}).items():
            try:
                viewer.headers[str(key)] = (
                    entry.get("header") or {},
                    entry.get("fds") or [],
                )
            except Exception:
                continue

    def _restore_channel_dropdown_from_headers(self):
        viewer = self.viewer
        if not getattr(viewer, "headers", None):
            try:
                viewer.channel_dropdown.blockSignals(True)
                viewer.channel_dropdown.clear()
                viewer.channel_dropdown.setEnabled(False)
            except Exception:
                pass
            finally:
                try:
                    viewer.channel_dropdown.blockSignals(False)
                except Exception:
                    pass
            return
        try:
            first_key = next(iter(viewer.headers))
            _, first_fds = viewer.headers[first_key]
        except Exception:
            first_fds = []
        labels = []
        for idx, fd in enumerate(first_fds or []):
            cap = fd.get("Caption", fd.get("FileName", f"chan{idx}"))
            labels.append(f"{idx}: {cap}")
        try:
            max_channels = max(len(v[1]) for v in viewer.headers.values())
        except Exception:
            max_channels = len(labels)
        if max_channels > len(labels):
            for idx in range(len(labels), max_channels):
                labels.append(f"{idx}: chan{idx}")
        try:
            viewer.channel_dropdown.blockSignals(True)
            viewer.channel_dropdown.clear()
            for lab in labels:
                viewer.channel_dropdown.addItem(lab)
                viewer.channel_dropdown.setItemData(viewer.channel_dropdown.count() - 1, lab, QtCore.Qt.ToolTipRole)
            viewer.channel_dropdown.setMinimumWidth(240)
            viewer.channel_dropdown.setEnabled(bool(labels))
        except Exception:
            pass
        finally:
            try:
                viewer.channel_dropdown.blockSignals(False)
            except Exception:
                pass
        try:
            viewer._sync_channel_nav_buttons()
        except Exception:
            pass

    def _restore_thumbnail_snapshot(self, snapshot: dict, thumbs_dir: Path):
        viewer = self.viewer
        if not snapshot:
            return 0
        viewer.clear_thumbs()
        items = list(snapshot.get("items") or [])
        if not items:
            return 0
        try:
            viewer.thumb_size_px = int(snapshot.get("thumb_size_px", getattr(viewer, "thumb_size_px", 160)) or getattr(viewer, "thumb_size_px", 160))
        except Exception:
            pass
        try:
            vp = getattr(viewer, "_thumb_viewport", None)
            avail_w = vp.width() if vp is not None else (viewer.thumb_container.width() if hasattr(viewer, "thumb_container") else 800)
        except Exception:
            avail_w = 800
        try:
            thumb_w = int(items[0].get("thumb_dims", [viewer.thumb_size_px, int(round(viewer.thumb_size_px * 0.75))])[0])
            thumb_h = int(items[0].get("thumb_dims", [viewer.thumb_size_px, int(round(viewer.thumb_size_px * 0.75))])[1])
        except Exception:
            thumb_w = int(getattr(viewer, "thumb_size_px", 160))
            thumb_h = int(round(thumb_w * 0.75))
        card_w = thumb_w + 24
        cols = max(1, min(12, int(avail_w / max(1, card_w))))
        viewer.thumb_grid_columns = cols
        viewer._thumb_card_height = thumb_h + 48
        viewer.current_thumb_files = []
        viewer.current_thumbnail_entries = []
        viewer.current_thumbnail_kind_by_key = {}
        viewer._thumb_meta = {}
        viewer._thumb_loaded = set()
        viewer._thumb_inflight = set()
        row = 0
        col = 0
        restored = 0
        for item in items:
            key = str(item.get("file_key") or "")
            png_file = str(item.get("png_file") or "")
            if not key or not png_file:
                continue
            pix = QtGui.QPixmap(str(thumbs_dir / png_file))
            if pix.isNull():
                continue
            lbl = QtWidgets.QLabel()
            lbl.setAlignment(QtCore.Qt.AlignCenter)
            lbl.setProperty("file_path", key)
            lbl.setProperty("channel_index", int(snapshot.get("channel_index", getattr(viewer, "last_channel_index", 0)) or 0))
            lbl.setProperty("spec_markers", [])
            lbl.setProperty("thumb_dims", (thumb_w, thumb_h))
            lbl.setProperty("drag_start", None)
            lbl.setProperty("dragging", False)
            lbl.setPixmap(pix)
            lbl.setMouseTracking(True)
            lbl.mousePressEvent = viewer._make_thumb_press_handler(lbl)
            lbl.mouseReleaseEvent = viewer._make_thumb_release_handler(lbl)
            lbl.mouseMoveEvent = viewer._make_thumb_move_handler(lbl)
            lbl.mouseDoubleClickEvent = viewer._make_thumb_double_handler(lbl)
            lbl.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
            lbl.customContextMenuRequested.connect(lambda pos, lb=lbl: viewer._on_thumb_context_menu(lb, pos))
            vbox = QtWidgets.QVBoxLayout()
            vbox.setContentsMargins(0, 0, 0, 0)
            vbox.setSpacing(2)
            card = QtWidgets.QFrame()
            card.setFrameShape(QtWidgets.QFrame.StyledPanel)
            card.setLineWidth(0)
            card_layout = QtWidgets.QVBoxLayout(card)
            card_layout.setContentsMargins(4, 4, 4, 4)
            card_layout.setSpacing(4)
            vbox.addWidget(lbl)
            cap = QtWidgets.QLabel(Path(key).name)
            cap.setAlignment(QtCore.Qt.AlignCenter)
            cap.setMaximumHeight(18)
            cap.setFont(QtGui.QFont("Segoe UI", 9))
            cap.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
            cap.customContextMenuRequested.connect(lambda pos, lb=lbl: viewer._on_thumb_context_menu(lb, pos))
            vbox.addWidget(cap)
            card_layout.addLayout(vbox)
            viewer.thumb_layout.addWidget(card, row, col)
            viewer.thumb_widgets[key] = card
            viewer._thumb_labels[key] = lbl
            viewer.current_thumb_files.append(key)
            viewer.current_thumbnail_entries.append({"kind": "image", "key": key})
            viewer.current_thumbnail_kind_by_key[key] = "image"
            viewer._thumb_loaded.add(key)
            restored += 1
            col += 1
            if col >= cols:
                col = 0
                row += 1
        try:
            viewer.selected_file_for_thumbs = snapshot.get("selected_file")
            viewer._refresh_thumb_selection_styles()
        except Exception:
            pass
        return restored

    def _restore_preview_from_snapshot(self, pending_preview, preview_snapshot: dict, preview_state: dict, views_dir: Path):
        viewer = self.viewer
        preview = getattr(viewer, "preview_canvas", None)
        if preview is None or not preview_snapshot:
            return False
        built_views = []
        for entry in preview_snapshot.get("views") or []:
            built = self._build_view_from_snapshot_entry(entry, views_dir)
            if built is None:
                return False
            built_views.append(built)
        if not built_views:
            return False
        try:
            if pending_preview and len(pending_preview) >= 2:
                viewer.last_preview = (str(pending_preview[0]), int(pending_preview[1]))
            else:
                meta = built_views[0].get("meta") or {}
                file_path = built_views[0].get("path") or meta.get("file_path")
                ch_idx = built_views[0].get("channel_idx", meta.get("channel_index", 0))
                viewer.last_preview = (str(file_path), int(ch_idx))
        except Exception:
            pass
        try:
            if hasattr(viewer, "adjust_image_btn"):
                viewer.adjust_image_btn.setEnabled(False)
        except Exception:
            pass
        try:
            if viewer.last_preview:
                viewer.selected_file_for_thumbs = str(viewer.last_preview[0])
                viewer._refresh_thumb_selection_styles()
                viewer._update_frame_map_active(str(viewer.last_preview[0]))
        except Exception:
            pass
        try:
            preview.set_views(built_views, preserve_profiles=False)
        except Exception:
            return False
        first_view = built_views[0] if built_views else {}
        meta = first_view.get("meta") or {}
        file_key = str(first_view.get("path") or meta.get("file_path") or "")
        try:
            ch_idx = int(first_view.get("channel_idx", meta.get("channel_index", 0)) or 0)
        except Exception:
            ch_idx = 0
        try:
            header_path = Path(file_key)
            header, fds = viewer.headers.get(file_key, (None, None))
            fd = fds[ch_idx] if fds and 0 <= ch_idx < len(fds) else {}
            html = viewer._build_metadata_html(
                header_path,
                header or {},
                fd or {},
                ch_idx,
                first_view.get("unit_normalized") or first_view.get("unit") or fd.get("PhysUnit", ""),
                first_view.get("unit") or "",
                np.asarray(first_view.get("arr")),
                first_view.get("zero_offset"),
            )
            viewer.meta_box.setHtml(html)
        except Exception:
            try:
                viewer.meta_box.setPlainText(f"File: {Path(file_key).name}")
            except Exception:
                pass
        return True

    def _schedule_session_hydration(self, payload: dict, session_path: Path):
        serial = int(getattr(self, "_session_hydration_serial", 0)) + 1
        self._session_hydration_serial = serial
        QtCore.QTimer.singleShot(0, lambda s=serial, p=payload, sp=Path(session_path): self._run_session_hydration(s, p, sp))

    def _run_session_hydration(self, serial: int, payload: dict, session_path: Path):
        if int(getattr(self, "_session_hydration_serial", 0)) != int(serial):
            return
        viewer = self.viewer
        start = time.perf_counter()
        preview_dt = 0.0
        thumbs_dt = 0.0
        spectro_dt = 0.0
        missing = 0
        spectro_rebuilt_preview = False
        self._set_session_activity(
            "Hydrating live session data...",
            detail="Refreshing spectroscopy links, thumbnails and preview from source data.",
            value=86,
            stage="hydrating",
        )
        try:
            files = [str(p) for p in (payload.get("files") or []) if str(p) and not viewer._is_processed_key(str(p))]
            missing = sum(1 for p in files if not Path(p).exists())
        except Exception:
            missing = 0

        spectra_folder = payload.get("spectra_folder") or ""
        if spectra_folder:
            t_phase = time.perf_counter()
            self._set_session_activity(
                "Hydrating live session data...",
                detail="Refreshing spectroscopy links.",
                value=89,
                stage="hydrating",
            )
            try:
                viewer._set_spec_folder(Path(spectra_folder))
            except Exception:
                pass
            try:
                if (
                    viewer.show_spectra
                    or viewer.show_preview_spectra
                    or viewer.show_spectro_miniatures
                    or viewer.show_matrix_markers
                    or viewer.show_single_markers
                ):
                    viewer.ensure_spectros_loaded(refresh=False)
                    spectro_rebuilt_preview = True
            except Exception:
                pass
            spectro_dt = time.perf_counter() - t_phase

        t_phase = time.perf_counter()
        self._set_session_activity(
            "Hydrating live session data...",
            detail="Refreshing live thumbnails.",
            value=92,
            stage="hydrating",
        )
        try:
            viewer.populate_thumbnails_for_channel(viewer.channel_dropdown.currentIndex())
        except Exception:
            pass
        thumbs_dt = time.perf_counter() - t_phase

        preview_ref = payload.get("last_preview")
        preview_key = None
        try:
            if preview_ref and len(preview_ref) >= 2:
                preview_key = (str(preview_ref[0]), int(preview_ref[1]))
        except Exception:
            preview_key = None
        t_phase = time.perf_counter()
        try:
            current_preview = getattr(viewer, "last_preview", None)
            if preview_key and current_preview == preview_key:
                preview_path_ok = viewer._is_processed_key(preview_key[0]) or Path(preview_key[0]).exists()
                if preview_path_ok and not spectro_rebuilt_preview:
                    self._set_session_activity(
                        "Hydrating live session data...",
                        detail="Refreshing live preview.",
                        value=97,
                        stage="hydrating",
                    )
                    viewer.show_file_channel(preview_key[0], preview_key[1])
                data_dir = payload.get("data_dir") or ""
                data_dir = session_path.parent / data_dir if data_dir else session_path.parent
                views_dir = data_dir / "views"
                preview_snapshot = payload.get("preview_canvas_snapshot")
                if preview_snapshot and preview_path_ok:
                    self._restore_canvas_snapshot(
                        getattr(viewer, "preview_canvas", None),
                        preview_snapshot,
                        views_dir,
                        viewer=viewer,
                        require_view_match=True,
                    )
        except Exception:
            pass
        preview_dt = time.perf_counter() - t_phase

        try:
            popup_pending = len(list(getattr(viewer, "_deferred_popup_entries", []) or []))
            detail = "Session is fully live."
            if popup_pending:
                detail += f" {popup_pending} deferred pop-up{'s' if popup_pending != 1 else ''} remain available in the Pop-ups menu."
            self._set_session_activity(
                "Session fully ready",
                detail=detail,
                value=100,
                stage="complete",
                hide_delay_ms=3200,
            )
            log_status(
                "[Session] hydrate %.2fs | thumbs %.2fs | preview %.2fs | spectros %.2fs | missing %d"
                % (
                    time.perf_counter() - start,
                    thumbs_dt,
                    preview_dt,
                    spectro_dt,
                    int(missing),
                )
            )
        except Exception:
            pass

    # ------------------------------------------------------------------
    def _capture_popup_snapshots(self, views_dir: Path):
        viewer = self.viewer
        canvases = list(getattr(viewer, "_popup_canvases", []) or [])
        snapshots = []
        captured_dialog_ids = set()
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
                captured_dialog_ids.add(id(dlg))
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
        lazy_idx = 0
        for dlg in list(getattr(viewer, "_popup_refs", []) or []):
            if dlg is None or id(dlg) in captured_dialog_ids:
                continue
            snap = getattr(dlg, "_lazy_popup_snapshot", None)
            source_dir = getattr(dlg, "_lazy_popup_views_dir", None)
            if not snap or not source_dir:
                continue
            try:
                if hasattr(dlg, "isVisible") and not dlg.isVisible():
                    continue
            except Exception:
                pass
            cloned = self._clone_lazy_popup_snapshot(snap, Path(source_dir), views_dir, prefix=f"lazy{lazy_idx}")
            if not cloned:
                continue
            try:
                geo = dlg.geometry()
                cloned["window_geometry"] = [geo.x(), geo.y(), geo.width(), geo.height()]
            except Exception:
                pass
            try:
                cloned["window_title"] = dlg.windowTitle()
            except Exception:
                pass
            snapshots.append(cloned)
            lazy_idx += 1
        for entry in list(getattr(viewer, "_deferred_popup_entries", []) or []):
            snap = entry.get("snapshot")
            source_dir = entry.get("views_dir")
            if not snap or not source_dir:
                continue
            cloned = self._clone_lazy_popup_snapshot(snap, Path(source_dir), views_dir, prefix=f"deferred{lazy_idx}")
            if not cloned:
                continue
            try:
                cloned["window_title"] = entry.get("title") or cloned.get("window_title")
            except Exception:
                pass
            snapshots.append(cloned)
            lazy_idx += 1
        return snapshots

    def _clone_lazy_popup_snapshot(self, snapshot: dict, source_views_dir: Path, dest_views_dir: Path, prefix: str):
        if not snapshot:
            return None
        try:
            cloned = copy.deepcopy(snapshot)
        except Exception:
            cloned = dict(snapshot)
        cloned_views = []
        for idx, entry in enumerate(snapshot.get("views") or []):
            try:
                new_entry = copy.deepcopy(entry)
            except Exception:
                new_entry = dict(entry or {})
            arr_file = entry.get("arr_file")
            if arr_file:
                src_path = source_views_dir / str(arr_file)
                dst_name = f"{prefix}_v{idx}.npy"
                dst_path = dest_views_dir / dst_name
                try:
                    if src_path.exists():
                        shutil.copyfile(src_path, dst_path)
                        new_entry["arr_file"] = dst_name
                    else:
                        new_entry["arr_file"] = None
                except Exception:
                    new_entry["arr_file"] = None
            cloned_views.append(new_entry)
        cloned["views"] = cloned_views
        return cloned

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
            "profile_dialog": self._safe_canvas_call(canvas, "export_profile_dialog_state"),
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

    def _hydrate_collection_processed_view(self, key: str):
        """Backfill collection processed entries with all source channels when available."""
        viewer = self.viewer
        data = (getattr(viewer, "_processed_views", {}) or {}).get(str(key))
        if not isinstance(data, dict) or str(data.get("op") or "") != "collection":
            return False
        source = str(data.get("source") or "").strip()
        if not source:
            return False
        try:
            source_path = Path(source)
        except Exception:
            return False
        if not source_path.exists():
            return False
        try:
            header, fds = viewer.headers.get(source, (None, None))
            if header is None or fds is None:
                header, fds = parse_header(source_path)
        except Exception:
            return False
        if not fds:
            return False
        existing = data.get("arr_by_channel") or {}
        arr_by_channel = {}
        for idx, fd in enumerate(fds):
            try:
                arr_by_channel[int(idx)] = np.array(
                    viewer._get_channel_array(source, idx, header, fd),
                    copy=True,
                )
            except Exception:
                continue
        if not arr_by_channel:
            arr_by_channel = dict(existing)
        if not arr_by_channel:
            return False
        try:
            sample = arr_by_channel.get(int(data.get("channel_idx") or 0))
        except Exception:
            sample = None
        if sample is None:
            try:
                sample = next(iter(arr_by_channel.values()))
            except Exception:
                sample = None
        header_new = dict(header or {})
        if sample is not None:
            try:
                sample = np.asarray(sample)
                header_new["xPixel"] = int(sample.shape[1])
                header_new["yPixel"] = int(sample.shape[0])
            except Exception:
                pass
        data["arr_by_channel"] = arr_by_channel
        data["header"] = header_new
        data["fds"] = [dict(fd or {}) for fd in fds]
        data["lock_channel"] = False
        viewer.headers[str(key)] = (header_new, data["fds"])
        return True

    def _restore_processed_views_payload(self, processed_payload: dict, processed_dir: Path):
        viewer = self.viewer
        viewer._processed_views = {}
        for key, entry in (processed_payload or {}).items():
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
                    "lock_channel": entry.get("lock_channel", True),
                    "source": entry.get("source"),
                    "label": entry.get("label"),
                    "op": entry.get("op"),
                }
                self._hydrate_collection_processed_view(str(key))
                restored = viewer._processed_views.get(str(key), {})
                viewer.headers[str(key)] = (
                    restored.get("header", header),
                    restored.get("fds", fds),
                )
            except Exception:
                continue

    def _session_files_from_payload(self, payload: dict):
        viewer = self.viewer
        session_files = []
        for fp in payload.get("files", []) or []:
            path_str = str(fp)
            if viewer._is_processed_key(path_str) and path_str in getattr(viewer, "_processed_views", {}):
                session_files.append(Path(path_str))
            elif Path(path_str).exists():
                session_files.append(Path(path_str))
        return session_files

    def _apply_basic_session_payload(self, payload: dict):
        viewer = self.viewer
        viewer.image_adjustments = payload.get("image_adjustments") or {}
        viewer.thumbnail_filters = payload.get("thumbnail_filters") or {}
        viewer.per_file_channel_cmap = payload.get("per_file_channel_cmap") or {}
        viewer.extra_view_specs = payload.get("extra_view_specs") or []
        viewer.tags = payload.get("tags") or {}
        viewer.molecule_overlays = payload.get("molecule_overlays") or {}
        viewer.thumb_multi_select = set(payload.get("thumb_multi_select") or [])
        viewer.selected_file_for_thumbs = payload.get("selected_file_for_thumbs")

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
                            "show_ticks": bool(snapshot.get("show_ticks", getattr(canvas, "_show_ticks", True))),
                            "show_colorbar": bool(snapshot.get("show_colorbar", getattr(canvas, "_show_colorbar", True))),
                            "colorbar_orientation": str(snapshot.get("colorbar_orientation", getattr(canvas, "_colorbar_orientation", "vertical")) or "vertical"),
                            "show_title": bool(snapshot.get("show_title", getattr(canvas, "_show_title", True))),
                            "show_acquisition_overlay": bool(snapshot.get("show_acquisition_overlay", getattr(canvas, "_show_acquisition_overlay", False))),
                            "show_shortcut_hint": bool(snapshot.get("show_shortcut_hint", getattr(canvas, "_show_shortcut_hint", True))),
                            "show_profile_overlays": bool(snapshot.get("show_profile_overlays", getattr(canvas, "_show_profile_overlays", True))),
                            "show_angle_overlays": bool(snapshot.get("show_angle_overlays", getattr(canvas, "_show_angle_overlays", True))),
                            "show_molecules": bool(snapshot.get("show_molecules", getattr(canvas, "show_molecules", True))),
                            "scale_bar_enabled": bool(snapshot.get("scale_bar_enabled", getattr(canvas, "scale_bar_enabled", False))),
                            "frame_fill_mode": bool(snapshot.get("frame_fill_mode", getattr(canvas, "_frame_fill_mode", False))),
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
            self._restore_view_specific_state(canvas, snapshot_views)
            prof = snapshot.get("profile_state")
            if prof:
                try:
                    canvas.import_profile_state(prof, emit=False)
                except Exception:
                    pass
            profile_dialog = snapshot.get("profile_dialog")
            if profile_dialog:
                try:
                    restorer = getattr(canvas, "restore_profile_dialog_state", None)
                    if callable(restorer):
                        restorer(profile_dialog)
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
            for field in (
                "extent",
                "extent_raw",
                "title",
                "colorbar_label",
                "axis_unit",
                "unit",
                "unit_normalized",
                "display_relative_zero",
                "zero_offset",
                "channel_idx",
                "crop_sequence",
                "path",
                "meta",
                "spectra",
                "highlight_spec",
                "spec_pixels",
            ):
                if field not in entry:
                    continue
                try:
                    incoming = copy.deepcopy(entry.get(field))
                except Exception:
                    incoming = entry.get(field)
                if target.get(field) != incoming:
                    target[field] = incoming
                    changed = True
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

    def _restore_popup_dialog_from_snapshot(
        self,
        snapshot: dict,
        views_dir: Path,
        *,
        geometry=None,
        window_state=None,
        visible=True,
        active=False,
        title=None,
    ):
        viewer = self.viewer
        if not snapshot or not views_dir:
            raise RuntimeError("missing popup view data")
        prev_display_sync = bool(getattr(viewer, "_canvas_display_syncing", False))
        viewer._canvas_display_syncing = True
        try:
            entries = snapshot.get("views") or []
            built_views = []
            for entry in entries:
                built = self._build_view_from_snapshot_entry(entry, Path(views_dir))
                if built is None:
                    built_views = []
                    break
                built_views.append(built)
            if not built_views:
                raise RuntimeError("missing popup view data")
            dlg = viewer._spawn_preview_popup(
                built_views,
                title=title or snapshot.get("window_title") or "Preview",
                show_immediately=False,
                restore_mode=True,
            )
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
            if canvas:
                self._restore_canvas_snapshot(canvas, snapshot, Path(views_dir), viewer=viewer)
            geom = geometry or snapshot.get("window_geometry")
            has_geometry = False
            if dlg and geom and len(geom) == 4:
                try:
                    x, y, w, h = [int(v) for v in geom]
                    dlg.setGeometry(x, y, w, h)
                    has_geometry = True
                except Exception:
                    pass
            if dlg and window_state is not None:
                try:
                    dlg.setWindowState(window_state)
                except Exception:
                    pass
            try:
                if dlg:
                    dlg.setUpdatesEnabled(True)
            except Exception:
                pass
            try:
                if canvas is not None and hasattr(canvas, "set_render_suspended"):
                    canvas.set_render_suspended(False)
            except Exception:
                pass
            if dlg:
                try:
                    if hasattr(dlg, "_resume_preview_resize"):
                        dlg._resume_preview_resize(force=not has_geometry)
                    else:
                        dlg._preview_resize_paused = False
                except Exception:
                    pass
                if visible:
                    dlg.show()
                if active:
                    try:
                        dlg.raise_()
                        dlg.activateWindow()
                    except Exception:
                        pass
            return dlg
        finally:
            viewer._canvas_display_syncing = prev_display_sync

    def restore_deferred_popup(self, entry_id=None, entry=None, *, show_activity=True, activate=True):
        popup_entry = self._lookup_deferred_popup(entry_id=entry_id, entry=entry)
        if not popup_entry:
            return None
        title = str(popup_entry.get("title") or self._deferred_popup_title(popup_entry.get("snapshot") or {}))
        if show_activity:
            self._set_session_activity(
                "Restoring deferred pop-up...",
                detail=title,
                value=40,
                stage="popup",
            )
        try:
            dlg = self._restore_popup_dialog_from_snapshot(
                popup_entry.get("snapshot") or {},
                Path(popup_entry.get("views_dir")),
                visible=True,
                active=activate,
                title=title,
            )
        except Exception as exc:
            if show_activity:
                self._set_session_activity(
                    "Unable to restore pop-up",
                    detail=f"{title}: {exc}",
                    value=100,
                    stage="error",
                    hide_delay_ms=3200,
                )
            return None
        self._remove_deferred_popup(popup_entry.get("id"))
        if show_activity:
            remaining = len(list(getattr(self.viewer, "_deferred_popup_entries", []) or []))
            detail = title
            if remaining:
                detail += f" | {remaining} deferred remaining in Pop-ups"
            self._set_session_activity(
                "Pop-up restored",
                detail=detail,
                value=100,
                stage="complete",
                hide_delay_ms=1800,
            )
        return dlg

    def restore_all_deferred_popups(self):
        viewer = self.viewer
        entries = list(getattr(viewer, "_deferred_popup_entries", []) or [])
        total = len(entries)
        if total <= 0:
            self._set_session_activity(
                "No deferred pop-ups",
                detail="There are no pending session pop-outs to restore.",
                value=100,
                stage="complete",
                hide_delay_ms=1400,
            )
            return []
        restored = []
        for idx, entry in enumerate(entries, start=1):
            title = str(entry.get("title") or self._deferred_popup_title(entry.get("snapshot") or {}))
            self._set_session_activity(
                "Restoring deferred pop-ups...",
                detail=f"{title} ({idx}/{total})",
                value=int(round((idx - 1) * 100.0 / max(1, total))),
                stage="popup",
            )
            dlg = self.restore_deferred_popup(
                entry=entry,
                show_activity=False,
                activate=(idx == total),
            )
            if dlg is not None:
                restored.append(dlg)
            self._pump_ui()
        self._set_session_activity(
            "Deferred pop-ups restored",
            detail=f"{len(restored)} pop-up{'s' if len(restored) != 1 else ''} are now open.",
            value=100,
            stage="complete",
            hide_delay_ms=2200,
        )
        return restored

    def _spawn_lazy_popup_shell(self, snapshot: dict, views_dir: Path):
        viewer = self.viewer
        dlg = QtWidgets.QDialog(viewer)
        dlg.setAttribute(QtCore.Qt.WA_DeleteOnClose, True)
        dlg.setWindowFlags(
            dlg.windowFlags()
            | QtCore.Qt.WindowMinimizeButtonHint
            | QtCore.Qt.WindowMaximizeButtonHint
            | QtCore.Qt.WindowSystemMenuHint
        )
        dlg.setMinimumSize(220, 140)
        dlg.setWindowTitle(snapshot.get("window_title") or "Preview")
        layout = QtWidgets.QVBoxLayout(dlg)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)
        title = QtWidgets.QLabel("Preview pending", dlg)
        title_font = title.font()
        title_font.setBold(True)
        title.setFont(title_font)
        layout.addWidget(title)
        body = QtWidgets.QLabel(
            "This pop-out will be restored when you focus or click it.",
            dlg,
        )
        body.setWordWrap(True)
        layout.addWidget(body, 1)
        restore_btn = QtWidgets.QPushButton("Restore now", dlg)
        layout.addWidget(restore_btn)

        dlg._lazy_popup_snapshot = copy.deepcopy(snapshot)
        dlg._lazy_popup_views_dir = Path(views_dir)
        dlg._lazy_popup_ready = False
        dlg._lazy_popup_hydrating = False
        dlg._lazy_popup_message_label = body
        dlg._lazy_popup_restore_btn = restore_btn

        def _queue_hydrate():
            if not getattr(dlg, "_lazy_popup_ready", False):
                return
            if getattr(dlg, "_lazy_popup_hydrating", False):
                return
            dlg._lazy_popup_hydrating = True
            try:
                restore_btn.setEnabled(False)
            except Exception:
                pass
            try:
                body.setText("Restoring preview...")
            except Exception:
                pass
            QtCore.QTimer.singleShot(0, lambda: self._hydrate_lazy_popup_shell(dlg))

        restore_btn.clicked.connect(_queue_hydrate)

        class _LazyPopupShellFilter(QtCore.QObject):
            def eventFilter(self, obj, event):
                if not getattr(dlg, "_lazy_popup_ready", False):
                    return False
                etype = event.type()
                if etype in (
                    QtCore.QEvent.MouseButtonPress,
                    QtCore.QEvent.MouseButtonDblClick,
                    QtCore.QEvent.FocusIn,
                    QtCore.QEvent.WindowActivate,
                ):
                    _queue_hydrate()
                elif etype == QtCore.QEvent.KeyPress:
                    _queue_hydrate()
                    return True
                return False

        dlg._lazy_popup_filter = _LazyPopupShellFilter(dlg)
        dlg.installEventFilter(dlg._lazy_popup_filter)
        body.installEventFilter(dlg._lazy_popup_filter)
        title.installEventFilter(dlg._lazy_popup_filter)

        viewer._popup_refs.append(dlg)
        if hasattr(viewer, "quick_crop_controller"):
            viewer.quick_crop_controller.update_popup_actions()

        def _on_popup_closed(_=None):
            if dlg in getattr(viewer, "_popup_refs", []):
                viewer._popup_refs.remove(dlg)
            if hasattr(viewer, "quick_crop_controller"):
                viewer.quick_crop_controller.update_popup_actions()
            if hasattr(viewer, "_clear_active_preview_popup"):
                try:
                    viewer._clear_active_preview_popup(dlg)
                except Exception:
                    pass

        dlg.finished.connect(_on_popup_closed)
        QtCore.QTimer.singleShot(0, lambda: setattr(dlg, "_lazy_popup_ready", True))
        return dlg

    def _hydrate_lazy_popup_shell(self, shell):
        if shell is None:
            return None
        snapshot = getattr(shell, "_lazy_popup_snapshot", None)
        views_dir = getattr(shell, "_lazy_popup_views_dir", None)
        message = getattr(shell, "_lazy_popup_message_label", None)
        button = getattr(shell, "_lazy_popup_restore_btn", None)
        if not snapshot or not views_dir:
            return None
        try:
            try:
                shell_geom = shell.geometry()
            except Exception:
                shell_geom = None
            try:
                shell_state = shell.windowState()
            except Exception:
                shell_state = QtCore.Qt.WindowNoState
            try:
                shell_visible = bool(shell.isVisible())
            except Exception:
                shell_visible = True
            try:
                shell_active = bool(shell.isActiveWindow())
            except Exception:
                shell_active = False
            dlg = self._restore_popup_dialog_from_snapshot(
                snapshot,
                Path(views_dir),
                geometry=shell_geom,
                window_state=shell_state,
                visible=shell_visible,
                active=shell_active,
                title=shell.windowTitle() or snapshot.get("window_title") or "Preview",
            )
            try:
                shell.close()
            except Exception:
                pass
            return dlg
        except Exception as exc:
            try:
                if message is not None:
                    message.setText(f"Unable to restore popup:\n{exc}")
            except Exception:
                pass
            try:
                if button is not None:
                    button.setEnabled(True)
            except Exception:
                pass
            shell._lazy_popup_hydrating = False
            return None

    def _restore_popup_canvases(self, popup_defs, views_dir: Path):
        if not popup_defs:
            return {"count": 0, "elapsed": 0.0, "arrays": 0.0, "spawn": 0.0, "state": 0.0, "show": 0.0, "lazy": 0}
        viewer = self.viewer
        start = time.perf_counter()
        arrays_dt = 0.0
        spawn_dt = 0.0
        state_dt = 0.0
        show_dt = 0.0
        lazy_count = 0
        restored = []
        lazy_mode = True
        prev_display_sync = bool(getattr(viewer, "_canvas_display_syncing", False))
        viewer._canvas_display_syncing = True
        try:
            for snap in popup_defs:
                t_phase = time.perf_counter()
                if lazy_mode:
                    try:
                        self._register_deferred_popup(snap, views_dir)
                    except Exception:
                        continue
                    lazy_count += 1
                    spawn_dt += time.perf_counter() - t_phase
                    continue
                else:
                    entries = snap.get("views") or []
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
                            restore_mode=True,
                        )
                    except Exception:
                        continue
                    canvas = None
                    try:
                        canvases = getattr(viewer, "_popup_canvases", [])
                        canvas = canvases[-1] if canvases else None
                    except Exception:
                        canvas = None
                spawn_dt += time.perf_counter() - t_phase
                try:
                    if dlg:
                        dlg.setUpdatesEnabled(False)
                except Exception:
                    pass
                if not lazy_mode:
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
                restored.append((dlg, canvas, has_geometry))
        finally:
            viewer._canvas_display_syncing = prev_display_sync
        shown = 0
        show_start = time.perf_counter()
        for dlg, canvas, has_geometry in restored:
            if dlg is None:
                continue
            try:
                dlg.setUpdatesEnabled(True)
            except Exception:
                pass
            try:
                if canvas is not None and hasattr(canvas, "set_render_suspended") and not lazy_mode:
                    canvas.set_render_suspended(False)
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
            "count": shown + lazy_count,
            "elapsed": time.perf_counter() - start,
            "arrays": arrays_dt,
            "spawn": spawn_dt,
            "state": state_dt,
            "show": show_dt,
            "lazy": lazy_count,
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
