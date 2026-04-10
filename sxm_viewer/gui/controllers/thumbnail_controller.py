"""Thumbnail interaction helpers."""
from __future__ import annotations

from pathlib import Path

from ..._shared import QtCore


class ThumbnailController:
    """Encapsulates thumbnail-selection and navigation behaviour."""

    def __init__(self, viewer):
        self.viewer = viewer

    # ------------------------------------------------------------------
    def handle_thumbnail_clicked(self, header_path_str, channel_idx):
        """Display the clicked thumbnail in preview and record selection."""
        viewer = self.viewer
        use_local_cmap = False
        try:
            use_local_cmap = bool((getattr(viewer, "per_file_channel_cmap", {}) or {}).get((str(header_path_str), int(channel_idx))))
        except Exception:
            use_local_cmap = False
        viewer.show_file_channel(header_path_str, channel_idx, use_local_cmap=use_local_cmap)
        key = str(header_path_str)
        viewer.selected_file_for_thumbs = key
        viewer._refresh_thumb_selection_styles()
        viewer.current_inspector_header = key
        viewer.current_inspector_channel = int(channel_idx)

    def handle_thumbnail_double_clicked(self, header_path_str, channel_idx):
        """Double-click thumbnail -> preview + popup."""
        viewer = self.viewer
        try:
            current_preview = getattr(viewer, "last_preview", None)
            current_views = getattr(getattr(viewer, "preview_canvas", None), "views", None)
            needs_refresh = not (
                current_preview
                and str(current_preview[0]) == str(header_path_str)
                and int(current_preview[1]) == int(channel_idx)
                and current_views
            )
            if needs_refresh:
                self.handle_thumbnail_clicked(header_path_str, channel_idx)
        except Exception:
            pass
        try:
            views = getattr(viewer.preview_canvas, "views", None)
            if not views:
                return
            copied = [viewer._copy_view_for_popup(v) for v in views]
            default_title = Path(header_path_str).name if header_path_str else "Preview"
            title = viewer._friendly_view_title(views[0], default=default_title)
            viewer._spawn_preview_popup(copied, title=title)
        except Exception:
            pass

    # ------------------------------------------------------------------
    def scroll_to_thumbnail(self, file_key):
        viewer = self.viewer
        if not file_key:
            return
        widget = getattr(viewer, "thumb_widgets", {}).get(str(file_key))
        if widget is None:
            return
        try:
            viewer.scroll.ensureWidgetVisible(widget)
        except Exception:
            try:
                bar = viewer.scroll.verticalScrollBar()
                if bar is not None:
                    bar.setValue(widget.y())
            except Exception:
                pass

    def handle_navigation(self, key, modifiers=QtCore.Qt.NoModifier):
        viewer = self.viewer
        files = list(getattr(viewer, "current_thumb_files", []) or [])
        if not files:
            return False
        cols = max(1, int(getattr(viewer, "thumb_grid_columns", 1) or 1))
        selected = str(getattr(viewer, "selected_file_for_thumbs", "") or "")
        try:
            current_idx = files.index(selected)
        except ValueError:
            current_idx = -1

        if current_idx < 0:
            new_idx = 0
        else:
            new_idx = current_idx
            if key == QtCore.Qt.Key_Left and new_idx > 0:
                new_idx -= 1
            elif key == QtCore.Qt.Key_Right and new_idx < len(files) - 1:
                new_idx += 1
            elif key == QtCore.Qt.Key_Up:
                new_idx = new_idx - cols if new_idx - cols >= 0 else 0
            elif key == QtCore.Qt.Key_Down:
                new_idx = new_idx + cols if new_idx + cols < len(files) else len(files) - 1
            else:
                return False
        if new_idx == current_idx:
            if current_idx == -1:
                return self.activate_thumbnail_by_index(0)
            return False
        return self.activate_thumbnail_by_index(new_idx)

    def activate_thumbnail_by_index(self, index):
        viewer = self.viewer
        files = list(getattr(viewer, "current_thumb_files", []) or [])
        if not files or not (0 <= index < len(files)):
            return False
        file_key = files[index]
        highlight = getattr(viewer, "_highlighted_spec", None)
        if highlight:
            try:
                highlight_path = str(highlight.get("image_key") or highlight.get("path") or "")
            except Exception:
                highlight_path = ""
            if highlight_path and highlight_path != file_key:
                viewer._highlight_spectrum_entry(None)
        viewer._clear_thumb_multi_selection(update_styles=False)
        label = getattr(viewer, "_thumb_labels", {}).get(file_key)
        if label is not None:
            try:
                channel_idx = int(label.property("channel_index") or 0)
            except Exception:
                channel_idx = viewer.channel_dropdown.currentIndex()
        else:
            channel_idx = viewer.channel_dropdown.currentIndex()
        viewer.on_thumbnail_clicked(file_key, channel_idx)
        viewer.last_thumb_anchor = str(file_key)
        self.scroll_to_thumbnail(file_key)
        return True

    # ------------------------------------------------------------------
    def focus_first_matrix_dataset(self):
        viewer = self.viewer
        matrix_files = list(getattr(viewer, "files_with_matrix", set()) or [])
        if not matrix_files:
            return
        target = None
        for path in getattr(viewer, "current_thumb_files", []):
            if path in matrix_files:
                target = path
                break
        if target is None:
            target = matrix_files[0]
        self.scroll_to_thumbnail(target)
        viewer.selected_file_for_thumbs = target
        viewer._refresh_thumb_selection_styles()

    def update_matrix_summary_banner(self):
        viewer = self.viewer
        label = getattr(viewer, "matrix_summary_label", None)
        if label is None:
            return
        matrix_count = len(getattr(viewer, "matrix_datasets", {}) or {})
        if matrix_count <= 0:
            label.hide()
            return
        noun = "Matrix dataset" if matrix_count == 1 else "Matrix datasets"
        label.setText(f"{noun}: {matrix_count} · click to focus")
        label.setToolTip("Click to jump to the first thumbnail containing a matrix spectroscopy grid.")
        label.show()
