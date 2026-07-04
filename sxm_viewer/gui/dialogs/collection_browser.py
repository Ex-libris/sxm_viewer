"""Browse, preview, and choose a collection.

Merges the previous "Choose Current Collection" and "Open Collection" blind file pickers into
one dialog with a live preview (item count, last-modified, thumbnail strip) before committing to
either action - so the user always sees what's inside before deciding whether to just set it as
the append target or fully open it into the workspace.
"""
from __future__ import annotations

from pathlib import Path

from ..._shared import QtCore, QtGui, QtWidgets


class CollectionBrowserDialog(QtWidgets.QDialog):
    """Preview a collection's contents, then choose "Set as Current Collection" (append target
    only, workspace untouched) or "Open" (fully replaces the workspace) - never both at once,
    and never implicit."""

    def __init__(self, parent, *, collection_controller, recent_collections, default_dir: Path):
        super().__init__(parent)
        self.setWindowTitle("Browse Collections")
        self.setModal(True)
        self.resize(680, 420)
        self._controller = collection_controller
        self._default_dir = Path(default_dir)
        self._selected_path = None

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        intro = QtWidgets.QLabel(
            "Pick a collection to preview its contents, then choose whether to set it as your "
            "current collection (append target only - your workspace is not touched) or open it "
            "(replaces your workspace with this collection's contents).",
            self,
        )
        intro.setWordWrap(True)
        layout.addWidget(intro)

        body = QtWidgets.QHBoxLayout()
        body.setSpacing(10)

        left = QtWidgets.QVBoxLayout()
        left.setSpacing(6)
        left.addWidget(QtWidgets.QLabel("Recent collections", self))
        self.list_widget = QtWidgets.QListWidget(self)
        for path in list(recent_collections or []):
            self._add_list_entry(path)
        self.list_widget.currentItemChanged.connect(self._on_selection_changed)
        left.addWidget(self.list_widget, 1)
        browse_btn = QtWidgets.QPushButton("Browse for Another Collection...", self)
        browse_btn.clicked.connect(self._on_browse)
        left.addWidget(browse_btn)
        body.addLayout(left, 1)

        right = QtWidgets.QVBoxLayout()
        right.setSpacing(6)
        self.summary_label = QtWidgets.QLabel("Select a collection on the left to preview it.", self)
        self.summary_label.setWordWrap(True)
        self.summary_label.setStyleSheet("font-weight: 600;")
        right.addWidget(self.summary_label)
        self.preview_list = QtWidgets.QListWidget(self)
        self.preview_list.setViewMode(QtWidgets.QListView.IconMode)
        self.preview_list.setIconSize(QtCore.QSize(72, 72))
        self.preview_list.setResizeMode(QtWidgets.QListView.Adjust)
        self.preview_list.setMovement(QtWidgets.QListView.Static)
        self.preview_list.setSelectionMode(QtWidgets.QAbstractItemView.NoSelection)
        self.preview_list.setFixedHeight(200)
        right.addWidget(self.preview_list, 1)
        body.addLayout(right, 2)

        layout.addLayout(body)

        buttons = QtWidgets.QDialogButtonBox(QtCore.Qt.Horizontal, self)
        self.set_current_btn = buttons.addButton("Set as Current Collection", QtWidgets.QDialogButtonBox.ActionRole)
        self.open_btn = buttons.addButton("Open (Load Into Workspace)", QtWidgets.QDialogButtonBox.ActionRole)
        cancel_btn = buttons.addButton(QtWidgets.QDialogButtonBox.Cancel)
        help_btn = buttons.addButton("What is a collection?", QtWidgets.QDialogButtonBox.HelpRole)
        self.set_current_btn.clicked.connect(self._on_set_current)
        self.open_btn.clicked.connect(self._on_open)
        cancel_btn.clicked.connect(self.reject)
        help_btn.clicked.connect(self._show_help)
        self.set_current_btn.setEnabled(False)
        self.open_btn.setEnabled(False)
        layout.addWidget(buttons)

        if self.list_widget.count():
            self.list_widget.setCurrentRow(0)

    @staticmethod
    def _display_label(path) -> str:
        name = Path(path).name
        if name.lower().endswith(".sxmcoll.json"):
            return name[: -len(".sxmcoll.json")]
        return name

    def _add_list_entry(self, path):
        item = QtWidgets.QListWidgetItem(self._display_label(path))
        item.setData(QtCore.Qt.UserRole, str(path))
        item.setToolTip(str(path))
        self.list_widget.addItem(item)
        return item

    def _on_browse(self):
        start = str(self._default_dir / "analysis_collection.sxmcoll.json")
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Browse for a collection",
            start,
            "SXM Collection (*.sxmcoll.json);;JSON (*.json)",
        )
        if not path:
            return
        path = str(Path(path))
        for i in range(self.list_widget.count()):
            if self.list_widget.item(i).data(QtCore.Qt.UserRole) == path:
                self.list_widget.setCurrentRow(i)
                return
        item = self._add_list_entry(path)
        self.list_widget.setCurrentItem(item)
        self.list_widget.setCurrentRow(self.list_widget.row(item))

    def _on_selection_changed(self, current, previous):
        self.preview_list.clear()
        if current is None:
            self._selected_path = None
            self.summary_label.setText("Select a collection on the left to preview it.")
            self.set_current_btn.setEnabled(False)
            self.open_btn.setEnabled(False)
            return
        path = current.data(QtCore.Qt.UserRole)
        self._selected_path = path
        try:
            summary = self._controller.collection_summary(path)
        except Exception:
            summary = {"exists": False, "item_count": 0, "updated_at": ""}
        if not summary.get("exists"):
            self.summary_label.setText(f"{self._display_label(path)} - file not found or unreadable.")
            self.set_current_btn.setEnabled(False)
            self.open_btn.setEnabled(False)
            return
        count = summary.get("item_count", 0)
        when = summary.get("updated_at") or "unknown"
        self.summary_label.setText(f"{self._display_label(path)} - {count} item(s), last updated {when}")
        self.set_current_btn.setEnabled(True)
        self.open_btn.setEnabled(True)
        try:
            entries = self._controller.tray_entries_for_current_collection(icon_size=72, collection_path=path)
        except Exception:
            entries = []
        if not entries:
            placeholder = QtWidgets.QListWidgetItem("No preview available")
            placeholder.setFlags(QtCore.Qt.NoItemFlags)
            self.preview_list.addItem(placeholder)
        else:
            for entry in entries[:6]:
                icon = entry.get("icon") or QtGui.QIcon()
                item = QtWidgets.QListWidgetItem(icon, entry.get("label") or "")
                item.setToolTip(entry.get("tool_tip") or "")
                item.setFlags(QtCore.Qt.NoItemFlags)
                self.preview_list.addItem(item)

    def _on_set_current(self):
        if not self._selected_path:
            return
        self._controller.choose_current_collection_for(self._selected_path)
        self.accept()

    def _on_open(self):
        if not self._selected_path:
            return
        self._controller.load_collection(Path(self._selected_path))
        self.accept()

    def _show_help(self):
        from ..controllers.collection import _COLLECTION_HELP_HTML
        QtWidgets.QMessageBox.information(self, "Collections", _COLLECTION_HELP_HTML)
