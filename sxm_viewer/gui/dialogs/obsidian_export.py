"""Dialog choosing the Obsidian vault/note an image export is appended to.

Pure UI: reads vault/note candidates through the Qt-free
``sxm_viewer.integrations.obsidian`` helpers and returns the choices via
``values()``; the caller does the rendering and writing.
"""
from __future__ import annotations

from pathlib import Path

from PyQt5 import QtWidgets

from ...integrations import obsidian


def suggested_attachment_setting(vault: Path) -> str:
    """The vault's own attachment folder; when that is the vault root but
    the vault clearly keeps images in a subfolder, suggest that one."""
    setting = obsidian.default_attachment_setting(vault)
    if setting.strip() in ("", "/"):
        for name in ("Images", "images", "Attachments", "attachments", "assets"):
            if (Path(vault) / name).is_dir():
                return name
    return setting


class ObsidianExportDialog(QtWidgets.QDialog):
    def __init__(self, parent=None, *, n_images=0, settings=None):
        super().__init__(parent)
        self.setWindowTitle("Export to Obsidian")
        self._settings = dict(settings or {})
        self._vaults = obsidian.list_vaults()
        form = QtWidgets.QFormLayout()

        self.vault_combo = QtWidgets.QComboBox()
        for v in self._vaults:
            self.vault_combo.addItem(v.name + ("  (open)" if v.is_open else ""), str(v.path))
            self.vault_combo.setItemData(self.vault_combo.count() - 1, str(v.path), 3)  # tooltip
        browse_vault = QtWidgets.QPushButton("Other vault...")
        browse_vault.clicked.connect(self._browse_vault)
        row = QtWidgets.QHBoxLayout()
        row.addWidget(self.vault_combo, 1)
        row.addWidget(browse_vault)
        form.addRow("Vault:", row)

        self.note_combo = QtWidgets.QComboBox()
        self.note_combo.setEditable(True)
        self.note_combo.setInsertPolicy(QtWidgets.QComboBox.NoInsert)
        self.note_combo.setMinimumWidth(360)
        self.note_combo.lineEdit().setPlaceholderText("note path in the vault (new notes are created)")
        browse_note = QtWidgets.QPushButton("Browse...")
        browse_note.clicked.connect(self._browse_note)
        row = QtWidgets.QHBoxLayout()
        row.addWidget(self.note_combo, 1)
        row.addWidget(browse_note)
        form.addRow("Note:", row)

        self.attach_edit = QtWidgets.QLineEdit()
        self.attach_edit.setToolTip("Folder for the PNGs, relative to the vault ('/' = vault root, "
                                    "'./sub' = next to the note), as in Obsidian's own setting")
        form.addRow("Image folder:", self.attach_edit)

        self.width_spin = QtWidgets.QSpinBox()
        self.width_spin.setRange(0, 4000)
        self.width_spin.setSingleStep(50)
        self.width_spin.setSpecialValueText("native")
        self.width_spin.setSuffix(" px")
        self.width_spin.setValue(int(self._settings.get("width", 500) or 0))
        form.addRow("Embed width:", self.width_spin)

        self.heading_cb = QtWidgets.QCheckBox("Add a timestamped heading")
        self.heading_cb.setChecked(bool(self._settings.get("heading", True)))
        self.caption_cb = QtWidgets.QCheckBox("Caption each image (file, channel, bias, date)")
        self.caption_cb.setChecked(bool(self._settings.get("captions", True)))
        form.addRow("", self.heading_cb)
        form.addRow("", self.caption_cb)

        layout = QtWidgets.QVBoxLayout(self)
        if n_images:
            info = QtWidgets.QLabel(f"{n_images} image(s), rendered as shown in the preview "
                                    "(current display preset), appended to the end of the note.")
            info.setWordWrap(True)
            layout.addWidget(info)
        if not self._vaults:
            warn = QtWidgets.QLabel("No Obsidian vaults found - use 'Other vault...'.")
            layout.addWidget(warn)
        layout.addLayout(form)
        buttons = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        buttons.button(QtWidgets.QDialogButtonBox.Ok).setText("Export")
        buttons.accepted.connect(self._accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.vault_combo.currentIndexChanged.connect(self._on_vault_changed)
        last = str(self._settings.get("vault") or "")
        idx = self.vault_combo.findData(last) if last else -1
        if idx < 0 and last and Path(last).is_dir():
            self.vault_combo.addItem(Path(last).name, last)
            idx = self.vault_combo.count() - 1
        if idx >= 0:
            self.vault_combo.setCurrentIndex(idx)
        self._on_vault_changed()

    # ----------------------------------------------------------------
    def vault_path(self):
        data = self.vault_combo.currentData()
        return Path(data) if data else None

    def _on_vault_changed(self, *_):
        vault = self.vault_path()
        self.note_combo.clear()
        if vault is None:
            return
        for note in obsidian.recent_notes(vault):
            self.note_combo.addItem(note.relative_to(vault).as_posix())
        attach = (self._settings.get("attachments") or {}).get(str(vault))
        self.attach_edit.setText(attach if attach is not None else suggested_attachment_setting(vault))

    def _browse_vault(self):
        start = str(self.vault_path() or Path.home())
        folder = QtWidgets.QFileDialog.getExistingDirectory(self, "Select Obsidian vault folder", start)
        if not folder:
            return
        idx = self.vault_combo.findData(folder)
        if idx < 0:
            self.vault_combo.addItem(Path(folder).name, folder)
            idx = self.vault_combo.count() - 1
        self.vault_combo.setCurrentIndex(idx)

    def _browse_note(self):
        vault = self.vault_path()
        if vault is None:
            return
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Choose or create a note", str(vault), "Markdown (*.md)",
            options=QtWidgets.QFileDialog.DontConfirmOverwrite)
        if not path:
            return
        try:
            text = Path(path).relative_to(vault).as_posix()
        except ValueError:
            QtWidgets.QMessageBox.warning(self, "Export to Obsidian", "The note must be inside the vault.")
            return
        self.note_combo.setEditText(text)

    def _accept(self):
        if self.vault_path() is None or not self.note_combo.currentText().strip():
            QtWidgets.QMessageBox.information(self, "Export to Obsidian", "Choose a vault and a note.")
            return
        self.accept()

    def values(self):
        vault = self.vault_path()
        return {
            "vault": vault,
            "note": obsidian.resolve_note_path(vault, self.note_combo.currentText()),
            "attachment": self.attach_edit.text().strip(),
            "width": int(self.width_spin.value()),
            "heading": self.heading_cb.isChecked(),
            "captions": self.caption_cb.isChecked(),
        }
