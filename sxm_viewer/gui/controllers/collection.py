"""Curated cross-folder collection save/load helpers."""
from __future__ import annotations

import copy
import json
from datetime import datetime
from pathlib import Path

from ..._shared import QtCore, QtGui, QtWidgets, log_status, np
from ...data.io import parse_header
from ..viewer import measurement as viewer_measurement
from ..thumbnail_render import array_to_qimage


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

        banner = QtWidgets.QFrame(self)
        banner.setStyleSheet(
            "QFrame {"
            "  background: #eef7ff;"
            "  border: 1px solid #b7d4f7;"
            "  border-radius: 8px;"
            "}"
        )
        banner_layout = QtWidgets.QVBoxLayout(banner)
        banner_layout.setContentsMargins(12, 10, 12, 10)
        banner_layout.setSpacing(4)
        banner_title = QtWidgets.QLabel("Add Selected Analysis To A Collection", banner)
        banner_title.setStyleSheet("font-weight: 700; color: #174a8b;")
        banner_msg = QtWidgets.QLabel(
            "Collections are reusable analysis sets. Keep appending into the same collection while you move across folders.",
            banner,
        )
        banner_msg.setWordWrap(True)
        banner_msg.setStyleSheet("color: #2a4560;")
        banner_layout.addWidget(banner_title)
        banner_layout.addWidget(banner_msg)
        layout.addWidget(banner)

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
            "If the file already exists, the selected items will be appended to it in place. "
            "The collection file is not overwritten. New files are created with the extension <code>.sxmcoll.json</code>.",
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

    def _collection_undo_stack(self):
        stack = getattr(self.viewer, "_collection_undo_stack", None)
        if stack is None:
            stack = []
            self.viewer._collection_undo_stack = stack
        return stack

    def _clear_collection_undo_stack(self):
        try:
            self.viewer._collection_undo_stack = []
        except Exception:
            pass

    def _push_collection_undo_state(self, collection_path: Path, payload: dict, *, description: str = ""):
        try:
            self._collection_undo_stack().append(
                {
                    "path": str(Path(collection_path)),
                    "payload": copy.deepcopy(self.viewer.session_controller._jsonify(payload)),
                    "description": str(description or ""),
                }
            )
        except Exception:
            pass

    def undo_last_collection_action(self):
        stack = self._collection_undo_stack()
        if not stack:
            return False
        current_path = str(getattr(self.viewer, "_collection_source", "") or "").strip()
        if not current_path:
            return False
        entry = stack.pop()
        entry_path = str(entry.get("path") or "").strip()
        if entry_path and str(Path(entry_path)) != str(Path(current_path)):
            stack.append(entry)
            return False
        payload = entry.get("payload")
        if not isinstance(payload, dict):
            return False
        collection_path = Path(entry_path or current_path)
        try:
            collection_path.parent.mkdir(parents=True, exist_ok=True)
            with open(collection_path, "w", encoding="utf-8") as fh:
                json.dump(payload, fh, indent=2)
        except Exception:
            stack.append(entry)
            return False
        try:
            self._load_payload_into_viewer(payload, collection_path)
        except Exception:
            pass
        return True

    def _collection_dialog_start_dir(self):
        current = str(getattr(self.viewer, "_collection_source", "") or "").strip()
        if current:
            try:
                current_path = Path(current)
                if current_path.exists():
                    return current_path.parent
                if current_path.parent:
                    return current_path.parent
            except Exception:
                pass
        last_collection_dir = getattr(self.viewer, "_last_collection_dir", None)
        if last_collection_dir:
            try:
                return Path(last_collection_dir)
            except Exception:
                pass
        return Path(getattr(self.viewer, "last_dir", "."))

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
                "Once you create or open a collection, it becomes the <b>current collection</b> for this app "
                "session. New <i>Add to collection</i> actions append to it by default until you open another "
                "collection.<br><br>"
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

    def add_thumbnail_entries(self, entries):
        """Add plain thumbnail/file entries to the current collection as fresh copies."""
        built_items = []
        for entry in list(entries or []):
            if not isinstance(entry, dict):
                continue
            file_path = str(entry.get("file_path") or "").strip()
            channel_idx = entry.get("channel_index")
            if not file_path or channel_idx is None:
                continue
            try:
                channel_idx = int(channel_idx)
            except Exception:
                continue
            try:
                bundle = self.viewer._build_single_channel_view(file_path, channel_idx)
            except Exception:
                bundle = None
            view = bundle.get("view") if isinstance(bundle, dict) else None
            if not isinstance(view, dict):
                continue
            built_items.append(
                self._build_item_from_view_snapshot(
                    view,
                    getattr(self.viewer, "preview_canvas", None),
                    source_kind="thumbnail",
                    restore_as_popup=False,
                    label=self._friendly_item_label(view, prefix="Thumbnail"),
                )
            )
        built_items = [item for item in built_items if isinstance(item, dict)]
        if built_items:
            self._save_items(
                built_items,
                source_summary=(
                    f"Add {len(built_items)} thumbnail selection(s) to the current collection.\n"
                    "These are stored as fresh collection copies without popup-only overlay state.\n"
                    "If a current collection is selected, the items are appended to it."
                ),
            )

    def add_from_view_drag_payload(self, payload: dict):
        """Add a dragged preview view as a fresh collection item."""
        if not isinstance(payload, dict):
            return
        view = None
        drag_token = payload.get("view_drag_token")
        if drag_token:
            try:
                from ..canvases.detail_preview_canvas import MultiPreviewCanvas
                view = MultiPreviewCanvas.consume_drag_view_snapshot(drag_token)
            except Exception:
                view = None
        if not isinstance(view, dict):
            file_path = str(payload.get("file_path") or "").strip()
            channel_idx = payload.get("channel_index")
            if file_path and channel_idx is not None:
                try:
                    bundle = self.viewer._build_single_channel_view(file_path, int(channel_idx))
                except Exception:
                    bundle = None
                view = bundle.get("view") if isinstance(bundle, dict) else None
        if not isinstance(view, dict):
            QtWidgets.QMessageBox.information(
                self.viewer,
                "Collections",
                "The dragged view could not be added to the collection.",
            )
            return
        item = self._build_item_from_view_snapshot(
            view,
            getattr(self.viewer, "preview_canvas", None),
            source_kind="dragged_view",
            restore_as_popup=False,
            label=self._friendly_item_label(view, prefix="Dragged view"),
        )
        if item:
            self._save_items(
                [item],
                source_summary=(
                    "Add the dragged preview view to the current collection.\n"
                    "Tip: use the popup Collection menu if you want to preserve popup-specific overlay state."
                ),
            )

    def remove_collection_items(self, item_ids):
        """Remove one or more items from the current collection file and refresh the workspace."""
        current_path = str(getattr(self.viewer, "_collection_source", "") or "").strip()
        if not current_path:
            QtWidgets.QMessageBox.information(self.viewer, "Collections", "No collection is currently selected.")
            return False
        ids = []
        for raw in list(item_ids or []):
            try:
                ids.append(int(raw))
            except Exception:
                continue
        ids = sorted(set(ids))
        if not ids:
            return False
        collection_path = Path(current_path)
        try:
            payload = self._load_or_init_payload(collection_path, mode=str(getattr(self.viewer, "_current_collection_mode", "linked") or "linked"))
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self.viewer, "Collections", f"Unable to open current collection: {exc}")
            return False
        items = list(payload.get("items") or [])
        kept = [item for item in items if int(item.get("id", -1)) not in ids]
        removed = len(items) - len(kept)
        if removed <= 0:
            QtWidgets.QMessageBox.information(self.viewer, "Collections", "The selected item was not found in the current collection.")
            return False
        if QtWidgets.QMessageBox.question(
            self.viewer,
            "Remove from collection",
            f"Remove {removed} item(s) from this collection?\n\nThis updates the collection file on disk.",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
            QtWidgets.QMessageBox.No,
        ) != QtWidgets.QMessageBox.Yes:
            return False
        self._push_collection_undo_state(collection_path, payload, description="remove_items")
        payload["items"] = kept
        payload["updated_at"] = datetime.utcnow().isoformat(timespec="seconds") + "Z"
        try:
            collection_path.parent.mkdir(parents=True, exist_ok=True)
            with open(collection_path, "w", encoding="utf-8") as fh:
                json.dump(self.viewer.session_controller._jsonify(payload), fh, indent=2)
            try:
                self.viewer._record_collection_dir(collection_path.parent)
            except Exception:
                pass
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self.viewer, "Collections", f"Unable to update collection: {exc}")
            return False
        try:
            self._load_payload_into_viewer(payload, collection_path)
        except Exception:
            pass
        return True

    def remove_thumbnail_entries(self, entries):
        """Remove thumbnail/file entries from the current collection."""
        current_path = str(getattr(self.viewer, "_collection_source", "") or "").strip()
        if not current_path:
            QtWidgets.QMessageBox.information(self.viewer, "Collections", "No collection is currently selected.")
            return False
        collection_path = Path(current_path)
        try:
            payload = self._load_or_init_payload(collection_path, mode=str(getattr(self.viewer, "_current_collection_mode", "linked") or "linked"))
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self.viewer, "Collections", f"Unable to open current collection: {exc}")
            return False
        requests = []
        for entry in list(entries or []):
            if not isinstance(entry, dict):
                continue
            file_path = str(entry.get("file_path") or "").strip()
            channel_idx = entry.get("channel_index")
            if not file_path or channel_idx is None:
                continue
            try:
                channel_idx = int(channel_idx)
            except Exception:
                continue
            requests.append((str(Path(file_path)), int(channel_idx)))
        if not requests:
            return False
        items = list(payload.get("items") or [])
        remaining = list(items)
        removed = 0
        for file_path, channel_idx in requests:
            for idx in range(len(remaining) - 1, -1, -1):
                item = remaining[idx] or {}
                snapshot = dict(item.get("snapshot") or {})
                first = (((snapshot.get("views") or [{}]) or [{}])[0]) or {}
                meta = dict(first.get("meta") or {})
                item_source = str(item.get("source_file") or meta.get("file_path") or meta.get("path") or first.get("path") or "")
                try:
                    item_channel = int(item.get("channel_index", meta.get("channel_index", first.get("channel_idx", -1))) or -1)
                except Exception:
                    item_channel = -1
                if str(Path(item_source)) == file_path and int(item_channel) == int(channel_idx):
                    remaining.pop(idx)
                    removed += 1
                    break
        if removed <= 0:
            QtWidgets.QMessageBox.information(self.viewer, "Collections", "The selected thumbnail(s) were not found in the current collection.")
            return False
        if QtWidgets.QMessageBox.question(
            self.viewer,
            "Remove from collection",
            f"Remove {removed} thumbnail item(s) from this collection?\n\nThis updates the collection file on disk.",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
            QtWidgets.QMessageBox.No,
        ) != QtWidgets.QMessageBox.Yes:
            return False
        self._push_collection_undo_state(collection_path, payload, description="remove_thumbnails")
        payload["items"] = remaining
        payload["updated_at"] = datetime.utcnow().isoformat(timespec="seconds") + "Z"
        try:
            collection_path.parent.mkdir(parents=True, exist_ok=True)
            with open(collection_path, "w", encoding="utf-8") as fh:
                json.dump(self.viewer.session_controller._jsonify(payload), fh, indent=2)
            try:
                self.viewer._record_collection_dir(collection_path.parent)
            except Exception:
                pass
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self.viewer, "Collections", f"Unable to update collection: {exc}")
            return False
        try:
            self._load_payload_into_viewer(payload, collection_path)
        except Exception:
            pass
        return True

    def load_collection(self, collection_path=None):
        path = collection_path
        if path is None:
            start = self._collection_dialog_start_dir() / "analysis_collection.sxmcoll.json"
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
        try:
            self.viewer._record_collection_dir(collection_path.parent)
        except Exception:
            pass
        self._clear_collection_undo_stack()
        self._load_payload_into_viewer(payload, collection_path)

    def choose_current_collection(self):
        """Pick or create the collection file that future add actions should append to."""
        start = str(self._collection_dialog_start_dir() / "analysis_collection.sxmcoll.json")
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self.viewer,
            "Choose current collection",
            start,
            "SXM Collection (*.sxmcoll.json);;JSON (*.json)",
        )
        if not path:
            return
        collection_path = Path(path)
        if collection_path.suffix.lower() != ".json" or not collection_path.name.endswith(".sxmcoll.json"):
            if collection_path.suffix.lower() == ".json":
                collection_path = collection_path.with_name(collection_path.stem + ".sxmcoll.json")
            else:
                collection_path = collection_path.with_suffix(".sxmcoll.json")
        mode = "linked"
        try:
            if collection_path.exists():
                payload = self._load_or_init_payload(collection_path, mode=mode)
                mode = str(payload.get("default_mode") or mode)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self.viewer, "Collections", f"Unable to use this collection: {exc}")
            return
        try:
            self.viewer._record_collection_dir(collection_path.parent)
        except Exception:
            pass
        self._clear_collection_undo_stack()
        self._remember_current_collection(collection_path, mode=mode)
        QtWidgets.QMessageBox.information(
            self.viewer,
            "Collections",
            f"Current collection set to:\n{collection_path}\n\nNew Add to Collection actions will append to this file in place; it will not be rewritten from scratch.",
        )

    def clear_current_collection(self):
        """Forget the current default collection target for this app session."""
        self.viewer._collection_source = None
        self.viewer._current_collection_mode = None
        self._clear_collection_undo_stack()
        refresh = getattr(self.viewer, "_refresh_collection_ui", None)
        if callable(refresh):
            try:
                refresh()
            except Exception:
                pass

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
            active_profile = None
            saved_profiles = None
            try:
                active_profile, saved_profiles = canvas.export_profile_datasets()
            except Exception:
                active_profile, saved_profiles = None, None
            try:
                self.viewer._last_profile_payload = (
                    active_profile,
                    list(saved_profiles or []),
                ) if (active_profile or saved_profiles) else None
            except Exception:
                pass
            profile_dialog = snapshot.get("profile_dialog")
            if profile_dialog and canvas is getattr(self.viewer, "preview_canvas", None):
                try:
                    viewer_measurement.restore_profile_dialog_state(self.viewer, profile_dialog)
                except Exception:
                    pass
            elif canvas is getattr(self.viewer, "preview_canvas", None):
                try:
                    dlg = getattr(self.viewer, "_profile_dialog", None)
                    if dlg is not None:
                        dlg.close()
                except Exception:
                    pass
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
        elif action == "collection_remove":
            if view is None:
                return
            source_file = str(self._view_source_path(view) or "").strip()
            if not source_file:
                return
            try:
                channel_idx = int((view.get("meta") or {}).get("channel_index", view.get("channel_idx", 0)) or 0)
            except Exception:
                channel_idx = 0
            current_path = str(getattr(self.viewer, "_collection_source", "") or "").strip()
            if not current_path:
                QtWidgets.QMessageBox.information(self.viewer, "Collections", "No collection is currently selected.")
                return
            try:
                payload = self._load_or_init_payload(Path(current_path), mode=str(getattr(self.viewer, "_current_collection_mode", "linked") or "linked"))
            except Exception as exc:
                QtWidgets.QMessageBox.warning(self.viewer, "Collections", f"Unable to open current collection: {exc}")
                return
            matches = []
            for item in list(payload.get("items") or []):
                snapshot = dict(item.get("snapshot") or {})
                first = (((snapshot.get("views") or [{}]) or [{}])[0]) or {}
                meta = dict(first.get("meta") or {})
                item_source = str(item.get("source_file") or meta.get("file_path") or meta.get("path") or first.get("path") or "")
                try:
                    item_channel = int(item.get("channel_index", meta.get("channel_index", first.get("channel_idx", -1))) or -1)
                except Exception:
                    item_channel = -1
                if str(Path(item_source)) == str(Path(source_file)) and int(item_channel) == int(channel_idx):
                    matches.append(int(item.get("id", -1)))
            if matches:
                self.remove_collection_items(matches)
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
            mode = str(payload.get("default_mode") or mode or "linked")
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
            self._push_collection_undo_state(collection_path, payload, description="add_items")
            payload.setdefault("items", []).extend(appended)
            payload["next_item_id"] = next_id
            payload["updated_at"] = datetime.utcnow().isoformat(timespec="seconds") + "Z"
            collection_path.parent.mkdir(parents=True, exist_ok=True)
            with open(collection_path, "w", encoding="utf-8") as fh:
                json.dump(self.viewer.session_controller._jsonify(payload), fh, indent=2)
            try:
                self.viewer._record_collection_dir(collection_path.parent)
            except Exception:
                pass
            self._remember_current_collection(collection_path, mode=mode)
            show_saved_cb = getattr(self.viewer, "_show_saved_path_toast", None)
            if callable(show_saved_cb):
                try:
                    show_saved_cb(
                        "Collection saved",
                        collection_path,
                        detail=f"Added {len(appended)} item(s) | {'Linked' if mode == 'linked' else 'Portable'} mode",
                    )
                except Exception:
                    pass
            show_tray = getattr(self.viewer, "show_collection_tray", None)
            if callable(show_tray):
                try:
                    show_tray(activate=False)
                except Exception:
                    pass
            log_status(f"Updated collection {collection_path} with {len(appended)} item(s)")
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self.viewer, "Collections", f"Unable to save collection: {exc}")

    def _prompt_target(self, source_summary: str):
        current_path = str(getattr(self.viewer, "_collection_source", "") or "").strip()
        if current_path:
            current = Path(current_path)
            if current.exists() and current.is_file():
                payload = self._load_or_init_payload(current, mode=str(getattr(self.viewer, "_current_collection_mode", "linked") or "linked"))
                mode = str(payload.get("default_mode") or getattr(self.viewer, "_current_collection_mode", "linked") or "linked")
                log_status(f"Appending to current collection: {current}")
                return str(current), mode
        default_path = current_path or str(self._collection_dialog_start_dir() / "analysis_collection.sxmcoll.json")
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
            payload.setdefault("default_mode", str(mode or "linked"))
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

    def _remember_current_collection(self, collection_path: Path, *, mode: str | None = None):
        """Keep one collection as the default append target for the current app session."""
        try:
            self.viewer._collection_source = str(Path(collection_path))
        except Exception:
            self.viewer._collection_source = str(collection_path)
        try:
            self.viewer._record_collection_dir(Path(collection_path).parent)
        except Exception:
            pass
        self.viewer._current_collection_mode = str(mode or getattr(self.viewer, "_current_collection_mode", "linked") or "linked")
        refresh = getattr(self.viewer, "_refresh_collection_ui", None)
        if callable(refresh):
            try:
                refresh()
            except Exception:
                pass

    @staticmethod
    def _snapshot_has_spectra(snapshot: dict):
        for entry in list((snapshot or {}).get("views") or []):
            if (entry or {}).get("spectra"):
                return True
        return False

    @staticmethod
    def _snapshot_has_analysis_state(snapshot: dict):
        if not isinstance(snapshot, dict):
            return False
        return bool(
            snapshot.get("profile_state")
            or snapshot.get("profile_dialog")
            or snapshot.get("angle_state")
            or snapshot.get("molecule_state")
        )

    def _should_restore_item_as_popup(self, item: dict, snapshot: dict):
        """Popup-like analysis items should reopen as popups instead of collection thumbnails."""
        if bool(item.get("restore_as_popup")):
            return True
        source_kind = str(item.get("source_kind") or "").strip().lower()
        if source_kind == "popup":
            return True
        return self._snapshot_has_analysis_state(snapshot)

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
            "show_ticks": bool(getattr(canvas, "_show_ticks", True)),
            "show_colorbar": bool(getattr(canvas, "_show_colorbar", True)),
            "colorbar_orientation": str(getattr(canvas, "_colorbar_orientation", "vertical") or "vertical"),
            "show_title": bool(getattr(canvas, "_show_title", True)),
            "show_acquisition_overlay": bool(getattr(canvas, "_show_acquisition_overlay", False)),
            "show_shortcut_hint": bool(getattr(canvas, "_show_shortcut_hint", True)),
            "show_profile_overlays": bool(getattr(canvas, "_show_profile_overlays", True)),
            "show_angle_overlays": bool(getattr(canvas, "_show_angle_overlays", True)),
            "show_molecules": bool(getattr(canvas, "show_molecules", True)),
            "frame_fill_mode": bool(getattr(canvas, "_frame_fill_mode", False)),
            "view_font_scale": float(getattr(canvas, "_view_font_scale", 1.0) or 1.0),
            "plot_font_family": str(getattr(canvas, "_font_family", "") or ""),
            "plot_font_bold": bool(getattr(canvas, "_plot_font_bold", False)),
            "plot_font_italic": bool(getattr(canvas, "_plot_font_italic", False)),
            "plot_font_underline": bool(getattr(canvas, "_plot_font_underline", False)),
            "profile_label_mode": str(getattr(canvas, "_profile_label_mode", "length") or "length"),
            "profile_state": session._safe_canvas_call(canvas, "export_profile_state") if include_state else None,
            "profile_dialog": (
                viewer_measurement.export_profile_dialog_state(self.viewer)
                if include_state and canvas is getattr(self.viewer, "preview_canvas", None)
                else session._safe_canvas_call(canvas, "export_profile_dialog_state") if include_state else None
            ),
            "angle_state": session._safe_canvas_call(canvas, "export_angle_state") if include_state else None,
            "molecule_state": session._safe_canvas_call(canvas, "export_molecule_state") if include_state else None,
            "scale_bar_pos": list(getattr(canvas, "_scale_bar_pos", (0.94, 0.06))),
            "scale_bar_settings": dict(getattr(canvas, "_scale_bar_settings", {}) or {}),
            "show_preview_spectra": bool(getattr(self.viewer, "show_preview_spectra", getattr(self.viewer, "show_spectra", True))),
            "show_spectra": bool(getattr(self.viewer, "show_spectra", True)),
            "show_spectro_miniatures": bool(getattr(self.viewer, "show_spectro_miniatures", False)),
            "share_overlapping_repeats": bool(getattr(self.viewer, "spectro_share_overlapping_repeats", False)),
            "spectro_settings": {
                "show_preview_spectra": bool(getattr(self.viewer, "show_preview_spectra", getattr(self.viewer, "show_spectra", True))),
                "show_spectra": bool(getattr(self.viewer, "show_spectra", True)),
                "show_spectro_miniatures": bool(getattr(self.viewer, "show_spectro_miniatures", False)),
                "share_overlapping_repeats": bool(getattr(self.viewer, "spectro_share_overlapping_repeats", False)),
                "show_matrix_markers": bool(getattr(self.viewer, "show_matrix_markers", True)),
                "show_single_markers": bool(getattr(self.viewer, "show_single_markers", True)),
                "compact_markers": bool(getattr(self.viewer, "compact_markers", True)),
                "highlight_glow": bool(getattr(self.viewer, "spectro_highlight_glow", True)),
                "single_symbol": str(getattr(self.viewer, "spectro_marker_symbol", "circle") or "circle"),
                "single_size": float(getattr(self.viewer, "spectro_marker_size", 5.0) or 5.0),
                "color_cycle": str(getattr(self.viewer, "spectro_color_cycle", "") or ""),
            },
            "filter_pipeline": None,
            "filter_label": None,
            "views": [],
            "zoom": session._safe_canvas_call(canvas, "export_zoom_states") if include_state and len(target_views) == len(getattr(canvas, "views", []) or []) else [],
        }
        try:
            host = canvas.window()
        except Exception:
            host = None
        if host is not None and host is not self.viewer:
            try:
                geo = host.geometry()
                snapshot["window_geometry"] = [int(geo.x()), int(geo.y()), int(geo.width()), int(geo.height())]
            except Exception:
                pass
            try:
                snapshot["window_state"] = int(host.windowState())
            except Exception:
                pass
            try:
                snapshot["window_title"] = str(host.windowTitle() or "")
            except Exception:
                pass
        try:
            pipeline, label = session._view_filter_spec(canvas)
            snapshot["filter_pipeline"] = pipeline
            snapshot["filter_label"] = label
        except Exception:
            pass
        has_analysis_state = bool(
            snapshot.get("profile_state")
            or snapshot.get("profile_dialog")
            or snapshot.get("angle_state")
            or snapshot.get("molecule_state")
        )
        for idx, view in enumerate(target_views):
            include_arrays = bool(
                mode == "portable"
                or self._view_requires_cached_array(view)
                or has_analysis_state
                or bool((view or {}).get("spectra"))
                or bool((view or {}).get("highlight_spec"))
            )
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
        channel_name = self._snapshot_channel_name(first)
        return {
            "title": str(first.get("title") or channel_name or Path(source_file).name or "Collection item"),
            "source_file": source_file,
            "source_folder": str(Path(source_file).parent) if source_file else "",
            "channel_index": meta.get("channel_index", first.get("channel_idx")),
            "channel_name": channel_name,
        }

    @staticmethod
    def _strip_display_unit(label) -> str:
        text = str(label or "").strip()
        if not text or not text.endswith("]"):
            return text
        head, sep, tail = text.rpartition("[")
        if not sep:
            return text
        base = head.rstrip()
        unit = tail[:-1].strip()
        if base and unit:
            return base
        return text

    def _snapshot_channel_name(self, view: dict) -> str:
        if not isinstance(view, dict):
            return ""
        meta = dict(view.get("meta") or {})
        for candidate in (
            meta.get("channel"),
            view.get("channel"),
            self._strip_display_unit(view.get("colorbar_label")),
        ):
            text = str(candidate or "").strip()
            if text:
                return text
        return ""

    def tray_entries_for_current_collection(self, *, icon_size: int = 72):
        """Build lightweight visual summaries for the collection tray in the main window."""
        current = str(getattr(self.viewer, "_collection_source", "") or "").strip()
        if not current:
            return []
        path = Path(current)
        if not path.exists():
            return []
        try:
            payload = self._load_or_init_payload(path, mode=str(getattr(self.viewer, "_current_collection_mode", "linked") or "linked"))
        except Exception:
            return []
        data_dir = path.parent / str(payload.get("data_dir") or f"{path.stem}_collection_data")
        views_dir = data_dir / "views"
        entries = []
        for item in reversed(list(payload.get("items") or [])):
            snapshot = dict(item.get("snapshot") or {})
            first = (((snapshot.get("views") or [{}]) or [{}])[0]) or {}
            meta = dict(first.get("meta") or {})
            source_file = str(item.get("source_file") or meta.get("file_path") or meta.get("path") or first.get("path") or "")
            folder_path = str(item.get("source_folder") or (str(Path(source_file).parent) if source_file else "") or "")
            folder_name = Path(folder_path).name if folder_path else ""
            channel_name = str(item.get("channel_name") or self._snapshot_channel_name(first) or "").strip()
            when = ""
            for key in ("datetime", "time", "time_str", "date", "Date", "Timestamp", "acquisition_time"):
                val = meta.get(key)
                if val:
                    when = str(val)
                    break
            icon = self._collection_item_icon(item, snapshot, views_dir, icon_size=icon_size)
            label = str(item.get("label") or self._snapshot_primary_meta(snapshot).get("title") or "Collection item")
            lines = [label]
            secondary = " | ".join(part for part in (folder_name, channel_name) if part)
            if secondary:
                lines.append(secondary)
            if when:
                lines.append(when)
            entries.append(
                {
                    "id": item.get("id"),
                    "label": label,
                    "text": "\n".join(lines),
                    "tool_tip": (
                        f"{label}\n"
                        f"Folder: {folder_path or '-'}\n"
                        f"Channel: {channel_name or '-'}\n"
                        f"Time: {when or '-'}\n"
                        f"Source: {source_file or '-'}"
                    ),
                    "icon": icon,
                    "source_file": source_file,
                    "folder_name": folder_name,
                    "channel_name": channel_name,
                    "when": when,
                }
            )
        return entries

    def _collection_item_icon(self, item: dict, snapshot: dict, views_dir: Path, *, icon_size: int = 72):
        first = (((snapshot.get("views") or [{}]) or [{}])[0]) or {}
        cmap = str(first.get("cmap") or getattr(self.viewer, "thumb_cmap", "viridis") or "viridis")
        try:
            view = self.viewer.session_controller._build_view_from_snapshot_entry(first, views_dir)
        except Exception:
            view = None
        if isinstance(view, dict):
            try:
                qimg = array_to_qimage(np.asarray(view.get("arr")), cmap_name=cmap)
                pix = QtGui.QPixmap.fromImage(qimg)
                return QtGui.QIcon(pix.scaled(icon_size, icon_size, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation))
            except Exception:
                pass
        source_file = str(item.get("source_file") or (first.get("meta") or {}).get("file_path") or first.get("path") or "")
        channel_idx = item.get("channel_index", (first.get("meta") or {}).get("channel_index", first.get("channel_idx")))
        if source_file and channel_idx is not None:
            try:
                pix = self.viewer._thumbnail_pixmap_for_file(source_file, int(channel_idx), icon_size, icon_size, getattr(self.viewer, "thumb_cmap", "viridis"))
                if pix is not None and not pix.isNull():
                    return QtGui.QIcon(pix)
            except Exception:
                pass
        return QtGui.QIcon()

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
            prepare_cb = getattr(viewer, "_prepare_for_workspace_load", None)
            if callable(prepare_cb):
                prepare_cb(kind="collection")
        except Exception:
            pass
        try:
            viewer.clear_loaded_images()
        except Exception:
            pass
        try:
            viewer.image_adjustments.clear()
        except Exception:
            pass
        try:
            viewer.per_file_channel_cmap.clear()
        except Exception:
            pass
        try:
            viewer.image_meta = []
        except Exception:
            pass
        try:
            viewer._spectro_browser_entries = []
        except Exception:
            pass
        try:
            viewer._multi_spec_selection = []
            viewer._multi_spec_selection_keys = set()
            viewer._last_clicked_spec = None
        except Exception:
            pass

        data_dir = collection_path.parent / str(payload.get("data_dir") or f"{collection_path.stem}_collection_data")
        views_dir = data_dir / "views"
        viewer.last_dir = collection_path.parent
        try:
            viewer._record_collection_dir(collection_path.parent)
        except Exception:
            pass
        try:
            viewer.path_le.setText(f"[Collection] {collection_path}")
        except Exception:
            pass
        try:
            viewer.spec_folder_path = collection_path.parent
            viewer.spec_folder_le.setText("")
        except Exception:
            pass
        viewer._collection_item_snapshots = {}
        viewer._workspace_kind = "collection"
        self._remember_current_collection(collection_path, mode=str(payload.get("default_mode") or "linked"))
        loaded_keys = []
        popup_items = []
        skipped = []
        any_spectra = False
        for item in items:
            snapshot = dict(item.get("snapshot") or {})
            restore_as_popup = self._should_restore_item_as_popup(item, snapshot)
            primary_view = self._build_primary_view_for_item(item, snapshot, views_dir)
            key = None
            if primary_view is not None:
                key = self._register_collection_processed_view(primary_view, item)
            if key:
                if not restore_as_popup:
                    viewer._collection_item_snapshots[str(key)] = {
                        "snapshot": snapshot,
                        "views_dir": str(views_dir),
                        "label": item.get("label"),
                    }
                    molecules = snapshot.get("molecule_state")
                    if molecules is not None:
                        viewer.molecule_overlays[str(key)] = molecules
                    any_spectra = self._register_collection_spectra_for_key(str(key), snapshot) or any_spectra
                loaded_keys.append(str(key))
            if restore_as_popup:
                popup_items.append((item, snapshot))
                any_spectra = self._snapshot_has_spectra(snapshot) or any_spectra
            elif not key:
                skipped.append(str(item.get("label") or item.get("id") or "item"))

        if any_spectra:
            self._apply_collection_spectro_settings(payload, items)

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
                    title=snapshot.get("window_title") or item.get("label"),
                    visible=True,
                    active=False,
                )
            except Exception:
                continue

        message = f"Opened collection with {len(loaded_keys)} library item(s)"
        if popup_items:
            message += f" and {len(popup_items)} restored pop-up(s)"
        message += "."
        if skipped:
            message += f"\n\nSkipped {len(skipped)} item(s) that could not be rebuilt."
        if payload.get("default_mode") == "linked":
            message += "\n\nLinked collection: original source files are preferred when available."
        else:
            message += "\n\nPortable collection: cached image data is being used."
        message += "\n\nThis collection is now the default target for Add to Collection actions in this app session."
        QtWidgets.QMessageBox.information(viewer, "Collection opened", message)
        log_status(f"Opened collection {collection_path} with {len(loaded_keys)} item(s)")

    def _setup_collection_channel_dropdown(self):
        viewer = self.viewer
        session = getattr(viewer, "session_controller", None)
        if session is not None:
            try:
                session._restore_channel_dropdown_from_headers()
                return
            except Exception:
                pass
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
                    bundle = self.viewer._build_single_channel_view(str(source_file), int(channel_idx))
                    built = bundle.get("view") if isinstance(bundle, dict) else None
                    if isinstance(built, dict):
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
        arr_by_channel = {}
        fds_new = []
        header_new = dict(source_header)
        try:
            if path and Path(str(path)).exists() and source_fds:
                for idx, source_fd in enumerate(source_fds):
                    try:
                        raw_arr = viewer._get_channel_array(str(path), idx, source_header, source_fd)
                        arr_by_channel[int(idx)] = np.array(raw_arr, copy=True)
                        fds_new.append(dict(source_fd or {}))
                    except Exception:
                        fds_new.append(dict(source_fd or {}))
                arr_by_channel = {idx: val for idx, val in arr_by_channel.items() if val is not None}
        except Exception:
            arr_by_channel = {}
            fds_new = []
        meta = view.get("meta") or {}
        caption = (
            self._snapshot_channel_name(view)
            or str(item.get("channel_name") or "").strip()
            or str(fd.get("Caption") or "").strip()
            or str(item.get("label") or "").strip()
            or str(view.get("title") or "").strip()
            or "Collection item"
        )
        if arr_by_channel:
            sample_arr = arr_by_channel.get(source_channel_idx)
            if sample_arr is None:
                try:
                    sample_arr = next(iter(arr_by_channel.values()))
                except Exception:
                    sample_arr = arr
            header_new["xPixel"] = int(np.asarray(sample_arr).shape[1])
            header_new["yPixel"] = int(np.asarray(sample_arr).shape[0])
        else:
            fd["Caption"] = caption
            fd["FileName"] = str(fd.get("FileName") or Path(str(path)).name or "collection_item")
            fds_new = [fd]
            arr_by_channel = {0: np.array(arr, copy=True)}
            header_new["xPixel"] = int(arr.shape[1])
            header_new["yPixel"] = int(arr.shape[0])
            source_channel_idx = 0
        key = viewer._make_processed_key(str(path), op="collection", channel_idx=0)
        viewer._processed_views[key] = {
            "arr_by_channel": arr_by_channel,
            "header": header_new,
            "fds": fds_new,
            "channel_idx": int(source_channel_idx),
            "lock_channel": False,
            "source": str(path),
            "label": "[collection]",
            "op": "collection",
        }
        viewer.headers[key] = (header_new, fds_new)
        viewer.files.append(Path(key))
        try:
            viewer._set_processed_insert_after(key, after_key="__virtual_copy_start__")
        except Exception:
            pass
        return key

    def _register_collection_spectra_for_key(self, file_key: str, snapshot: dict):
        """Attach saved spectroscopy entries back onto a curated collection item."""
        specs = []
        seen = set()
        for entry in list(snapshot.get("views") or []):
            for spec in list((entry or {}).get("spectra") or []):
                if not isinstance(spec, dict):
                    continue
                copied = copy.deepcopy(spec)
                copied["image_key"] = str(file_key)
                key = (
                    str(copied.get("path") or ""),
                    copied.get("matrix_index"),
                    copied.get("x"),
                    copied.get("y"),
                    copied.get("order_idx"),
                )
                if key in seen:
                    continue
                seen.add(key)
                specs.append(copied)
        if not specs:
            return False
        viewer = self.viewer
        viewer.spectros_by_image[str(file_key)] = specs
        all_specs = list(getattr(viewer, "spectros", []) or [])
        all_specs.extend(specs)
        viewer.spectros = all_specs
        matrix_specs = [spec for spec in specs if spec.get("matrix_index") is not None]
        if matrix_specs:
            existing = list(getattr(viewer, "matrix_spectros", []) or [])
            existing.extend(matrix_specs)
            viewer.matrix_spectros = existing
            try:
                viewer.files_with_matrix.add(str(file_key))
            except Exception:
                pass
        try:
            viewer._spectros_loaded = True
            viewer._spectros_pending = False
        except Exception:
            pass
        return True

    def _apply_collection_spectro_settings(self, payload: dict, items: list[dict]):
        """Collections carrying spectroscopy should restore the corresponding overlay visibility/settings."""
        viewer = self.viewer
        spectro_settings = {}
        for item in items:
            snapshot = item.get("snapshot") or {}
            settings = snapshot.get("spectro_settings") or {}
            if settings:
                spectro_settings = settings
                break
        if not spectro_settings:
            return
        viewer.show_spectra = bool(spectro_settings.get("show_spectra", True))
        viewer.show_preview_spectra = bool(spectro_settings.get("show_preview_spectra", viewer.show_spectra))
        viewer.show_spectro_miniatures = bool(spectro_settings.get("show_spectro_miniatures", getattr(viewer, "show_spectro_miniatures", False)))
        viewer.spectro_share_overlapping_repeats = bool(
            spectro_settings.get("share_overlapping_repeats", getattr(viewer, "spectro_share_overlapping_repeats", False))
        )
        viewer.show_matrix_markers = bool(spectro_settings.get("show_matrix_markers", True))
        viewer.show_single_markers = bool(spectro_settings.get("show_single_markers", True))
        viewer.compact_markers = bool(spectro_settings.get("compact_markers", True))
        viewer.spectro_highlight_glow = bool(spectro_settings.get("highlight_glow", True))
        try:
            viewer.spectro_marker_symbol = str(spectro_settings.get("single_symbol", viewer.spectro_marker_symbol) or viewer.spectro_marker_symbol)
            viewer.spectro_marker_size = float(spectro_settings.get("single_size", viewer.spectro_marker_size) or viewer.spectro_marker_size)
        except Exception:
            pass
        cycle = str(spectro_settings.get("color_cycle", "") or "").strip()
        if cycle:
            viewer.spectro_color_cycle = cycle
        for attr in (
            "spectro_overlay_act",
            "preview_spectra_toggle_btn",
            "show_spectra_cb",
            "spectro_thumbnail_markers_cb",
            "spectro_preview_markers_cb",
            "spectro_miniatures_cb",
            "spectro_miniatures_act",
            "highlight_glow_act",
            "toolbar_spectro_markers_act",
            "toolbar_spectro_preview_act",
            "toolbar_spectro_miniatures_act",
            "toolbar_spectro_repeat_share_act",
            "toolbar_spectro_matrix_markers_act",
            "toolbar_spectro_single_markers_act",
            "toolbar_spectro_compact_markers_act",
            "toolbar_spectro_highlight_act",
            "toolbar_spectro_thumb_btn",
            "toolbar_spectro_preview_btn",
            "toolbar_spectro_miniatures_btn",
        ):
            widget = getattr(viewer, attr, None)
            if widget is None:
                continue
            try:
                widget.blockSignals(True)
                if attr in {
                    "show_spectra_cb",
                    "spectro_preview_markers_cb",
                    "toolbar_spectro_preview_act",
                    "toolbar_spectro_preview_btn",
                }:
                    widget.setChecked(viewer.show_preview_spectra)
                elif attr in {
                    "spectro_miniatures_cb",
                    "spectro_miniatures_act",
                    "toolbar_spectro_miniatures_act",
                    "toolbar_spectro_miniatures_btn",
                }:
                    widget.setChecked(viewer.show_spectro_miniatures)
                elif attr in {"toolbar_spectro_repeat_share_act"}:
                    widget.setChecked(viewer.spectro_share_overlapping_repeats)
                elif attr in {"matrix_markers_act", "toolbar_spectro_matrix_markers_act"}:
                    widget.setChecked(viewer.show_matrix_markers)
                elif attr in {"single_markers_act", "toolbar_spectro_single_markers_act"}:
                    widget.setChecked(viewer.show_single_markers)
                elif attr in {"compact_markers_act", "toolbar_spectro_compact_markers_act"}:
                    widget.setChecked(viewer.compact_markers)
                elif attr in {"highlight_glow_act", "toolbar_spectro_highlight_act"}:
                    widget.setChecked(viewer.spectro_highlight_glow)
                else:
                    widget.setChecked(viewer.show_spectra)
            except Exception:
                pass
            finally:
                try:
                    widget.blockSignals(False)
                except Exception:
                    pass
        try:
            if getattr(viewer, "spectros", None):
                viewer._assign_spectros_to_images()
        except Exception:
            pass
        try:
            viewer._update_spectro_stats_label()
        except Exception:
            pass
