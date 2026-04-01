"""Curated cross-folder collection save/load helpers."""
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

from ..._shared import QtCore, QtWidgets, log_status, np
from ...data.io import parse_header


class _CollectionTargetDialog(QtWidgets.QDialog):
    """Prompt for collection destination and linked/portable storage mode."""

    def __init__(self, parent, *, source_summary: str, default_path: str):
        super().__init__(parent)
        self.setWindowTitle("Add to Collection")
        self.setModal(True)
        self.resize(640, 0)

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        intro = QtWidgets.QLabel(
            "<b>Collections</b> are curated workspaces built from selected views across folders or sessions.<br>"
            "Choose whether this save should stay lightweight (<b>Linked</b>) or carry its own image data "
            "for moving/sharing (<b>Portable</b>).",
            self,
        )
        intro.setWordWrap(True)
        intro.setTextFormat(QtCore.Qt.RichText)
        layout.addWidget(intro)

        summary = QtWidgets.QLabel(source_summary, self)
        summary.setWordWrap(True)
        summary.setStyleSheet("color: #555;")
        layout.addWidget(summary)

        mode_group = QtWidgets.QGroupBox("Storage mode", self)
        mode_layout = QtWidgets.QVBoxLayout(mode_group)
        mode_layout.setContentsMargins(10, 10, 10, 10)
        mode_layout.setSpacing(8)

        self.linked_rb = QtWidgets.QRadioButton(
            "Linked (Recommended): keep the collection light and reopen original source views when possible. "
            "Derived crops are cached only when needed.",
            mode_group,
        )
        self.linked_rb.setChecked(True)
        self.portable_rb = QtWidgets.QRadioButton(
            "Portable: cache every selected image array inside the collection. Larger file, but safer to move "
            "to another machine or share with someone else.",
            mode_group,
        )
        mode_layout.addWidget(self.linked_rb)
        mode_layout.addWidget(self.portable_rb)
        layout.addWidget(mode_group)

        path_group = QtWidgets.QGroupBox("Collection file", self)
        path_layout = QtWidgets.QGridLayout(path_group)
        path_layout.setContentsMargins(10, 10, 10, 10)
        path_layout.setHorizontalSpacing(8)
        path_layout.setVerticalSpacing(8)
        path_layout.addWidget(QtWidgets.QLabel("Path", path_group), 0, 0)
        self.path_edit = QtWidgets.QLineEdit(default_path, path_group)
        self.path_edit.setPlaceholderText("Choose an existing collection to append, or type a new file name.")
        path_layout.addWidget(self.path_edit, 0, 1)
        browse_btn = QtWidgets.QPushButton("Browse...", path_group)
        browse_btn.clicked.connect(self._on_browse)
        path_layout.addWidget(browse_btn, 0, 2)
        hint = QtWidgets.QLabel(
            "If the file already exists, the selected items will be appended to it. "
            "New files are created with the extension <code>.sxmcoll.json</code>.",
            path_group,
        )
        hint.setWordWrap(True)
        hint.setTextFormat(QtCore.Qt.RichText)
        hint.setStyleSheet("color: #555;")
        path_layout.addWidget(hint, 1, 0, 1, 3)
        layout.addWidget(path_group)

        buttons = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel,
            QtCore.Qt.Horizontal,
            self,
        )
        help_btn = buttons.addButton("What is a collection?", QtWidgets.QDialogButtonBox.HelpRole)
        help_btn.clicked.connect(self._show_help)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _on_browse(self):
        start = self.path_edit.text().strip() or "analysis_collection.sxmcoll.json"
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Choose collection file",
            start,
            "SXM Collection (*.sxmcoll.json);;JSON (*.json)",
        )
        if path:
            self.path_edit.setText(path)

    def _show_help(self):
        QtWidgets.QMessageBox.information(
            self,
            "Collections",
            (
                "<b>Linked</b> collections keep the file smaller and still remember where each item came from. "
                "They are best when the original data stays on the same machine.<br><br>"
                "<b>Portable</b> collections cache every selected image/crop, so they reopen more safely on a "
                "different machine or after moving files. They take more disk space."
            ),
        )

    def values(self):
        return self.path_edit.text().strip(), ("portable" if self.portable_rb.isChecked() else "linked")


class CollectionController:
    """Create and reopen curated, cross-folder collections of selected analysis items."""

    KIND = "sxm_collection"
    VERSION = 1

    def __init__(self, viewer):
        self.viewer = viewer

    # ------------------------------------------------------------------
    def show_help(self):
        QtWidgets.QMessageBox.information(
            self.viewer,
            "Collections",
            (
                "<b>Collections</b> are curated workspaces made from selected preview views, pop-ups, and crop "
                "snapshots.<br><br>"
                "Use them when you want to compare results from different folders without saving the whole "
                "folder session.<br><br>"
                "<b>Linked</b>: lighter, expects original source data to remain available when possible.<br>"
                "<b>Portable</b>: larger, caches image arrays so the collection can reopen more safely."
            ),
        )

    def add_current_preview(self):
        canvas = getattr(self.viewer, "preview_canvas", None)
        view = ((getattr(canvas, "views", None) or [None]) or [None])[0]
        if canvas is None or not isinstance(view, dict):
            QtWidgets.QMessageBox.information(self.viewer, "Collections", "There is no preview image to add.")
            return
        item = self._build_item_from_canvas(
            canvas,
            source_kind="preview",
            restore_as_popup=False,
            label=self._friendly_item_label(view, prefix="Preview"),
        )
        if item:
            self._save_items([item], source_summary=f"Add the current preview to a collection.\nItem: {item['label']}")

    def add_active_popup(self):
        canvas = getattr(self.viewer, "_active_preview_canvas", None)
        if canvas is None:
            popups = list(getattr(self.viewer, "_popup_canvases", []) or [])
            canvas = popups[-1] if popups else None
        if canvas is None:
            QtWidgets.QMessageBox.information(self.viewer, "Collections", "There is no active pop-up to add.")
            return
        view = ((getattr(canvas, "views", None) or [None]) or [None])[0]
        item = self._build_item_from_canvas(
            canvas,
            source_kind="popup",
            restore_as_popup=True,
            label=self._friendly_item_label(view, prefix="Pop-up"),
        )
        if item:
            self._save_items([item], source_summary=f"Add the active pop-up to a collection.\nItem: {item['label']}")

    def add_all_popups(self):
        canvases = [c for c in list(getattr(self.viewer, "_popup_canvases", []) or []) if c is not None and getattr(c, "views", None)]
        if not canvases:
            QtWidgets.QMessageBox.information(self.viewer, "Collections", "There are no open pop-ups to add.")
            return
        items = []
        for idx, canvas in enumerate(canvases, start=1):
            view = ((getattr(canvas, "views", None) or [None]) or [None])[0]
            item = self._build_item_from_canvas(
                canvas,
                source_kind="popup",
                restore_as_popup=True,
                label=self._friendly_item_label(view, prefix=f"Pop-up {idx}"),
            )
            if item:
                items.append(item)
        if items:
            self._save_items(items, source_summary=f"Add {len(items)} open pop-up(s) to a collection.")

    def add_selected_crop_history(self):
        controller = getattr(self.viewer, "quick_crop_controller", None)
        preview_canvas = getattr(self.viewer, "preview_canvas", None)
        if controller is None or preview_canvas is None:
            QtWidgets.QMessageBox.information(self.viewer, "Collections", "There is no crop history available.")
            return
        seqs = list(getattr(controller, "selected_sequences", []) or [])
        if not seqs:
            active = getattr(controller, "active_sequence", None)
            if active is not None:
                seqs = [active]
        if not seqs:
            QtWidgets.QMessageBox.information(
                self.viewer,
                "Collections",
                "Select one or more crop-history entries first, then add them to a collection.",
            )
            return
        items = []
        for seq in seqs:
            try:
                entry = preview_canvas.get_fixed_crop_history_entry(seq)
            except Exception:
                entry = None
            if not entry:
                continue
            view = entry.get("view_snapshot")
            if not isinstance(view, dict):
                continue
            item = self._build_item_from_view_snapshot(
                view,
                preview_canvas,
                source_kind="crop_history",
                restore_as_popup=False,
                label=self._friendly_item_label(view, prefix=f"Crop #{seq}"),
            )
            if item:
                items.append(item)
        if items:
            self._save_items(items, source_summary=f"Add {len(items)} selected crop snapshot(s) to a collection.")

    def load_collection(self, collection_path=None):
        path = collection_path
        if path is None:
            start = Path(getattr(self.viewer, "last_dir", ".")) / "analysis_collection.sxmcoll.json"
            path, _ = QtWidgets.QFileDialog.getOpenFileName(
                self.viewer,
                "Open collection",
                str(start),
                "SXM Collection (*.sxmcoll.json);;JSON (*.json)",
            )
            if not path:
                return
        collection_path = Path(path)
        try:
            with open(collection_path, "r", encoding="utf-8") as fh:
                payload = json.load(fh)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self.viewer, "Open collection", f"Unable to open collection: {exc}")
            return
        if str(payload.get("kind") or "") != self.KIND:
            QtWidgets.QMessageBox.warning(
                self.viewer,
                "Open collection",
                "This file is not an SXM collection. Use Load Session for normal session files.",
            )
            return
        self._load_payload_into_viewer(payload, collection_path)

    def apply_snapshot_for_file(self, file_key):
        state = (getattr(self.viewer, "_collection_item_snapshots", {}) or {}).get(str(file_key))
        if not state:
            return False
        snapshot = state.get("snapshot") or {}
        views_dir = state.get("views_dir")
        canvas = getattr(self.viewer, "preview_canvas", None)
        if canvas is None or not views_dir:
            return False
        try:
            self.viewer.session_controller._restore_canvas_snapshot(
                canvas,
                snapshot,
                Path(views_dir),
                viewer=self.viewer,
                require_view_match=False,
            )
            return True
        except Exception:
            return False

    def handle_canvas_menu_action(self, action, view, canvas=None):
        if action == "collection_add":
            target_canvas = canvas or getattr(self.viewer, "preview_canvas", None)
            if target_canvas is not None and len(list(getattr(target_canvas, "views", []) or [])) <= 1:
                item = self._build_item_from_canvas(
                    target_canvas,
                    source_kind="view",
                    restore_as_popup=False,
                    label=self._friendly_item_label(view, prefix="View"),
                )
            else:
                item = self._build_item_from_view_snapshot(
                    view,
                    target_canvas,
                    source_kind="view",
                    restore_as_popup=False,
                    label=self._friendly_item_label(view, prefix="View"),
                )
            if item:
                self._save_items([item], source_summary=f"Add this view to a collection.\nItem: {item['label']}")
        elif action == "collection_help":
            self.show_help()

    # ------------------------------------------------------------------
    def _save_items(self, items, *, source_summary: str):
        items = [item for item in list(items or []) if isinstance(item, dict)]
        if not items:
            return
        path, mode = self._prompt_target(source_summary)
        if not path:
            return
        collection_path = Path(path)
        if collection_path.suffix.lower() != ".json" or not collection_path.name.endswith(".sxmcoll.json"):
            if collection_path.suffix.lower() == ".json":
                collection_path = collection_path.with_name(collection_path.stem + ".sxmcoll.json")
            else:
                collection_path = collection_path.with_suffix(".sxmcoll.json")
        try:
            payload = self._load_or_init_payload(collection_path, mode=mode)
            data_dir = collection_path.parent / str(payload.get("data_dir") or f"{collection_path.stem}_collection_data")
            views_dir = data_dir / "views"
            views_dir.mkdir(parents=True, exist_ok=True)
            next_id = int(payload.get("next_item_id", 1) or 1)
            appended = []
            for raw_item in items:
                item = dict(raw_item)
                item_id = int(next_id)
                next_id += 1
                snapshot = self._recapture_item_snapshot(item, views_dir, item_id=item_id, mode=mode)
                if not snapshot:
                    continue
                primary = self._snapshot_primary_meta(snapshot)
                appended.append(
                    {
                        "id": item_id,
                        "label": item.get("label") or primary.get("title") or f"Collection item {item_id}",
                        "source_kind": item.get("source_kind") or "view",
                        "storage_mode": mode,
                        "restore_as_popup": bool(item.get("restore_as_popup", False)),
                        "created_at": datetime.utcnow().isoformat(timespec="seconds") + "Z",
                        "source_file": primary.get("source_file"),
                        "source_folder": primary.get("source_folder"),
                        "channel_index": primary.get("channel_index"),
                        "channel_name": primary.get("channel_name"),
                        "snapshot": snapshot,
                    }
                )
            if not appended:
                QtWidgets.QMessageBox.information(
                    self.viewer,
                    "Collections",
                    "Nothing could be added to the collection from the current selection.",
                )
                return
            payload.setdefault("items", []).extend(appended)
            payload["next_item_id"] = next_id
            payload["updated_at"] = datetime.utcnow().isoformat(timespec="seconds") + "Z"
            collection_path.parent.mkdir(parents=True, exist_ok=True)
            with open(collection_path, "w", encoding="utf-8") as fh:
                json.dump(payload, fh, indent=2)
            QtWidgets.QMessageBox.information(
                self.viewer,
                "Collections",
                (
                    f"Added {len(appended)} item(s) to {collection_path.name}.\n\n"
                    f"{'Linked' if mode == 'linked' else 'Portable'} mode is now stored for these new entries."
                ),
            )
            log_status(f"Saved {len(appended)} collection item(s) to {collection_path}")
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self.viewer, "Collections", f"Unable to save collection: {exc}")

    def _prompt_target(self, source_summary: str):
        default_path = str(Path(getattr(self.viewer, "last_dir", ".")) / "analysis_collection.sxmcoll.json")
        dlg = _CollectionTargetDialog(self.viewer, source_summary=source_summary, default_path=default_path)
        if dlg.exec_() != QtWidgets.QDialog.Accepted:
            return None, None
        return dlg.values()

    def _load_or_init_payload(self, collection_path: Path, *, mode: str):
        if collection_path.exists():
            with open(collection_path, "r", encoding="utf-8") as fh:
                payload = json.load(fh)
            if str(payload.get("kind") or "") != self.KIND:
                raise ValueError("Selected file exists but is not an SXM collection.")
            payload.setdefault("items", [])
            payload.setdefault("data_dir", f"{collection_path.stem}_collection_data")
            payload.setdefault("next_item_id", len(payload.get("items") or []) + 1)
            return payload
        return {
            "kind": self.KIND,
            "version": self.VERSION,
            "created_at": datetime.utcnow().isoformat(timespec="seconds") + "Z",
            "updated_at": datetime.utcnow().isoformat(timespec="seconds") + "Z",
            "data_dir": f"{collection_path.stem}_collection_data",
            "default_mode": mode,
            "items": [],
            "next_item_id": 1,
            "help": {
                "linked": "Lightweight collection that reopens original source views when possible. Derived views cache arrays only when needed.",
                "portable": "Caches every selected image array so the collection can be reopened more safely on another machine.",
            },
        }

    def _recapture_item_snapshot(self, item: dict, views_dir: Path, *, item_id: int, mode: str):
        prefix = f"item{item_id}"
        kind = str(item.get("capture_kind") or "")
        capture_mode = "portable" if bool(item.get("restore_as_popup", False)) else mode
        if kind == "canvas":
            canvas = item.get("canvas")
            if canvas is None:
                return None
            return self._capture_collection_snapshot(canvas, views_dir, prefix=prefix, mode=capture_mode)
        if kind == "view":
            view = item.get("view")
            canvas = item.get("canvas") or getattr(self.viewer, "preview_canvas", None)
            if not isinstance(view, dict):
                return None
            return self._capture_collection_snapshot(
                canvas,
                views_dir,
                prefix=prefix,
                mode=capture_mode,
                views=[view],
                include_state=False,
            )
        return None

    def _build_item_from_canvas(self, canvas, *, source_kind: str, restore_as_popup: bool, label: str):
        if canvas is None or not getattr(canvas, "views", None):
            return None
        return {
            "capture_kind": "canvas",
            "canvas": canvas,
            "source_kind": source_kind,
            "restore_as_popup": bool(restore_as_popup),
            "label": label,
        }

    def _build_item_from_view_snapshot(self, view, canvas, *, source_kind: str, restore_as_popup: bool, label: str):
        if not isinstance(view, dict):
            return None
        return {
            "capture_kind": "view",
            "view": dict(view),
            "canvas": canvas,
            "source_kind": source_kind,
            "restore_as_popup": bool(restore_as_popup),
            "label": label,
        }

    def _capture_collection_snapshot(self, canvas, views_dir: Path, *, prefix: str, mode: str, views=None, include_state: bool = True):
        session = getattr(self.viewer, "session_controller", None)
        if canvas is None or session is None:
            return None
        target_views = [dict(v) for v in list(views if views is not None else (getattr(canvas, "views", []) or [])) if isinstance(v, dict)]
        if not target_views:
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
            "profile_state": session._safe_canvas_call(canvas, "export_profile_state") if include_state else None,
            "profile_dialog": session._safe_canvas_call(canvas, "export_profile_dialog_state") if include_state else None,
            "angle_state": session._safe_canvas_call(canvas, "export_angle_state") if include_state else None,
            "molecule_state": session._safe_canvas_call(canvas, "export_molecule_state") if include_state else None,
            "scale_bar_pos": list(getattr(canvas, "_scale_bar_pos", (0.94, 0.06))),
            "scale_bar_settings": dict(getattr(canvas, "_scale_bar_settings", {}) or {}),
            "filter_pipeline": None,
            "filter_label": None,
            "views": [],
            "zoom": session._safe_canvas_call(canvas, "export_zoom_states") if include_state and len(target_views) == len(getattr(canvas, "views", []) or []) else [],
        }
        try:
            pipeline, label = session._view_filter_spec(canvas)
            snapshot["filter_pipeline"] = pipeline
            snapshot["filter_label"] = label
        except Exception:
            pass
        for idx, view in enumerate(target_views):
            include_arrays = bool(mode == "portable" or self._view_requires_cached_array(view))
            serialized = session._serialize_view_for_session(view, views_dir, f"{prefix}_v{idx}", include_arrays)
            snapshot["views"].append(serialized)
        return snapshot

    def _view_requires_cached_array(self, view: dict):
        path = self._view_source_path(view)
        title = str(view.get("title") or "").lower()
        if view.get("crop_sequence") is not None:
            return True
        if path and self.viewer._is_processed_key(str(path)):
            return True
        if "[crop]" in title or "[copy]" in title:
            return True
        if not path:
            return True
        try:
            return not Path(str(path)).exists()
        except Exception:
            return True

    def _snapshot_primary_meta(self, snapshot: dict):
        first = ((snapshot.get("views") or [{}]) or [{}])[0]
        meta = dict(first.get("meta") or {})
        source_file = str(meta.get("file_path") or meta.get("path") or first.get("path") or "")
        return {
            "title": str(first.get("title") or meta.get("channel") or Path(source_file).name or "Collection item"),
            "source_file": source_file,
            "source_folder": str(Path(source_file).parent) if source_file else "",
            "channel_index": meta.get("channel_index", first.get("channel_idx")),
            "channel_name": meta.get("channel") or "",
        }

    def _friendly_item_label(self, view, *, prefix: str):
        meta = (view or {}).get("meta") or {}
        title = str((view or {}).get("title") or meta.get("channel") or "").strip()
        file_name = str(meta.get("file_name") or Path(str(meta.get("file_path") or (view or {}).get("path") or "")).name or "").strip()
        parts = [str(prefix).strip()]
        if file_name:
            parts.append(file_name)
        if title:
            parts.append(title)
        return " | ".join(part for part in parts if part)

    def _view_source_path(self, view):
        if not isinstance(view, dict):
            return None
        meta = view.get("meta") or {}
        return view.get("path") or meta.get("path") or meta.get("file_path")

    # ------------------------------------------------------------------
    def _load_payload_into_viewer(self, payload: dict, collection_path: Path):
        viewer = self.viewer
        items = list(payload.get("items") or [])
        if not items:
            QtWidgets.QMessageBox.information(viewer, "Collections", "This collection does not contain any items.")
            return
        try:
            viewer.on_close_popouts()
        except Exception:
            pass
        try:
            viewer.clear_loaded_images()
        except Exception:
            pass

        data_dir = collection_path.parent / str(payload.get("data_dir") or f"{collection_path.stem}_collection_data")
        views_dir = data_dir / "views"
        viewer.last_dir = collection_path.parent
        try:
            viewer.path_le.setText(f"[Collection] {collection_path}")
        except Exception:
            pass
        try:
            viewer.spec_folder_le.setText("")
        except Exception:
            pass
        viewer._collection_item_snapshots = {}
        viewer._collection_source = str(collection_path)
        viewer._workspace_kind = "collection"
        loaded_keys = []
        popup_items = []
        skipped = []
        for item in items:
            snapshot = dict(item.get("snapshot") or {})
            primary_view = self._build_primary_view_for_item(item, snapshot, views_dir)
            if primary_view is None:
                skipped.append(str(item.get("label") or item.get("id") or "item"))
                continue
            key = self._register_collection_processed_view(primary_view, item)
            if not key:
                skipped.append(str(item.get("label") or item.get("id") or "item"))
                continue
            viewer._collection_item_snapshots[str(key)] = {
                "snapshot": snapshot,
                "views_dir": str(views_dir),
                "label": item.get("label"),
            }
            molecules = snapshot.get("molecule_state")
            if molecules is not None:
                viewer.molecule_overlays[str(key)] = molecules
            loaded_keys.append(str(key))
            if bool(item.get("restore_as_popup")):
                popup_items.append((item, snapshot))

        self._setup_collection_channel_dropdown()
        try:
            viewer.populate_thumbnails_for_channel(0)
        except Exception:
            pass
        if loaded_keys:
            try:
                viewer.show_file_channel(loaded_keys[0], 0)
            except Exception:
                pass
        for item, snapshot in popup_items:
            try:
                viewer.session_controller._restore_popup_dialog_from_snapshot(
                    snapshot,
                    views_dir,
                    title=item.get("label"),
                    visible=True,
                    active=False,
                )
            except Exception:
                continue

        message = f"Opened collection with {len(loaded_keys)} item(s)."
        if skipped:
            message += f"\n\nSkipped {len(skipped)} item(s) that could not be rebuilt."
        if payload.get("default_mode") == "linked":
            message += "\n\nLinked collection: original source files are preferred when available."
        else:
            message += "\n\nPortable collection: cached image data is being used."
        QtWidgets.QMessageBox.information(viewer, "Collection opened", message)
        log_status(f"Opened collection {collection_path} with {len(loaded_keys)} item(s)")

    def _setup_collection_channel_dropdown(self):
        viewer = self.viewer
        try:
            viewer.channel_dropdown.blockSignals(True)
            viewer.channel_dropdown.clear()
            viewer.channel_dropdown.addItem("0: Collection item")
            viewer.channel_dropdown.setItemData(0, "Collection items are curated single-channel entries.", QtCore.Qt.ToolTipRole)
            viewer.channel_dropdown.setCurrentIndex(0)
            viewer.channel_dropdown.setEnabled(True)
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

    def _build_primary_view_for_item(self, item: dict, snapshot: dict, views_dir: Path):
        entries = list(snapshot.get("views") or [])
        if not entries:
            return None
        first = dict(entries[0])
        if not first.get("arr_file"):
            source_file = self._view_source_path(first)
            channel_idx = (first.get("meta") or {}).get("channel_index", first.get("channel_idx"))
            try:
                if source_file and Path(str(source_file)).exists() and channel_idx is not None:
                    built = self.viewer._build_single_channel_view(str(source_file), int(channel_idx))
                    if built:
                        if first.get("title"):
                            built["title"] = first.get("title")
                        if first.get("cmap"):
                            built["cmap"] = first.get("cmap")
                        return built
            except Exception:
                pass
        try:
            return self.viewer.session_controller._build_view_from_snapshot_entry(first, views_dir)
        except Exception:
            return None

    def _register_collection_processed_view(self, view: dict, item: dict):
        viewer = self.viewer
        path = self._view_source_path(view) or str(item.get("id") or "collection")
        arr = view.get("arr")
        if arr is None:
            return None
        try:
            arr = np.asarray(arr)
        except Exception:
            return None
        if arr.ndim < 2 or arr.size == 0:
            return None
        source_header = {}
        source_fds = []
        source_channel_idx = (view.get("meta") or {}).get("channel_index", view.get("channel_idx"))
        try:
            source_channel_idx = int(source_channel_idx) if source_channel_idx is not None else 0
        except Exception:
            source_channel_idx = 0
        try:
            source_header, source_fds = viewer.headers.get(str(path), (None, None))
            if source_header is None or source_fds is None:
                source_header, source_fds = parse_header(Path(str(path)))
        except Exception:
            source_header, source_fds = {}, []
        source_header = dict(source_header or {})
        fd = {}
        if source_fds and 0 <= source_channel_idx < len(source_fds):
            fd = dict(source_fds[source_channel_idx] or {})
        meta = view.get("meta") or {}
        fd["Caption"] = str(view.get("title") or meta.get("channel") or fd.get("Caption") or item.get("label") or "Collection item")
        fd["FileName"] = str(fd.get("FileName") or Path(str(path)).name or "collection_item")
        header_new = dict(source_header)
        header_new["xPixel"] = int(arr.shape[1])
        header_new["yPixel"] = int(arr.shape[0])
        key = viewer._make_processed_key(str(path), op="collection", channel_idx=0)
        viewer._processed_views[key] = {
            "arr_by_channel": {0: np.array(arr, copy=True)},
            "header": header_new,
            "fds": [fd],
            "channel_idx": 0,
            "source": str(path),
            "label": "[collection]",
            "op": "collection",
        }
        viewer.headers[key] = (header_new, [fd])
        viewer.files.append(Path(key))
        try:
            viewer._set_processed_insert_after(key, after_key="__virtual_copy_start__")
        except Exception:
            pass
        return key
