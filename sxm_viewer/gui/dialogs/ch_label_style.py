"""Style editor for the CH height label (font, colours, outline, box, corner).

Pure UI: emits ``style_changed(dict)`` on every edit; the owner applies it
live (``gui/viewer/ch_height.py``). The dict only holds keys the user
actually set - anything absent keeps the per-mode default (boxed HUD in
Analysis, outlined text in Publication), so an untouched dialog changes
nothing.
"""
from __future__ import annotations

from ..._shared import QtCore, QtGui, QtWidgets

POSITIONS = [
    ("Top right", "top-right"),
    ("Top left", "top-left"),
    ("Bottom right", "bottom-right"),
    ("Bottom left", "bottom-left"),
]

# What the controls show for keys the user hasn't set (Analysis HUD look).
DISPLAY_DEFAULTS = {
    "font_size": 0.0,
    "font_family": "",
    "bold": True,
    "color": "#f5f7fb",
    "outline": False,
    "outline_color": "#000000",
    "box": True,
    "box_color": "#000000",
    "box_alpha": 0.42,
    "position": "top-right",
}


class _ColorButton(QtWidgets.QPushButton):
    color_changed = QtCore.pyqtSignal(str)

    def __init__(self, color, title, parent=None):
        super().__init__(parent)
        self._title = title
        self.setFixedWidth(64)
        self.set_color(color)
        self.clicked.connect(self._pick)

    def color(self):
        return self._color

    def set_color(self, color):
        self._color = QtGui.QColor(color).name()
        self.setStyleSheet(f"background-color: {self._color}; border: 1px solid #888; min-height: 18px;")

    def _pick(self):
        col = QtWidgets.QColorDialog.getColor(QtGui.QColor(self._color), self, self._title)
        if col.isValid():
            self.set_color(col.name())
            self.color_changed.emit(self._color)


class ChLabelStyleDialog(QtWidgets.QDialog):
    style_changed = QtCore.pyqtSignal(dict)

    def __init__(self, style=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("CH height label style")
        self._initial = dict(style or {})
        self._style = dict(self._initial)
        val = lambda key: self._style.get(key, DISPLAY_DEFAULTS[key])

        form = QtWidgets.QFormLayout()

        self.size_spin = QtWidgets.QDoubleSpinBox()
        self.size_spin.setRange(0.0, 48.0)
        self.size_spin.setSingleStep(0.5)
        self.size_spin.setSpecialValueText("Auto")
        self.size_spin.setSuffix(" pt")
        self.size_spin.setValue(float(val("font_size") or 0.0))
        self.size_spin.valueChanged.connect(lambda v: self._set("font_size", float(v)))
        form.addRow("Font size", self.size_spin)

        self.font_check = QtWidgets.QCheckBox("Custom")
        self.font_combo = QtWidgets.QFontComboBox()
        family = val("font_family")
        self.font_check.setChecked(bool(family))
        self.font_combo.setEnabled(bool(family))
        if family:
            self.font_combo.setCurrentFont(QtGui.QFont(family))
        self.font_check.toggled.connect(self._on_font_toggled)
        self.font_combo.currentFontChanged.connect(
            lambda f: self.font_check.isChecked() and self._set("font_family", f.family()))
        font_row = QtWidgets.QHBoxLayout()
        font_row.addWidget(self.font_check)
        font_row.addWidget(self.font_combo, 1)
        form.addRow("Font", font_row)

        self.bold_check = QtWidgets.QCheckBox("Bold")
        self.bold_check.setChecked(bool(val("bold")))
        self.bold_check.toggled.connect(lambda on: self._set("bold", bool(on)))
        form.addRow("", self.bold_check)

        self.color_btn = _ColorButton(val("color"), "Text colour")
        self.color_btn.color_changed.connect(lambda c: self._set("color", c))
        form.addRow("Text colour", self.color_btn)

        self.outline_check = QtWidgets.QCheckBox("Outline")
        self.outline_check.setChecked(bool(val("outline")))
        self.outline_check.toggled.connect(lambda on: self._set("outline", bool(on)))
        self.outline_btn = _ColorButton(val("outline_color"), "Outline colour")
        self.outline_btn.color_changed.connect(lambda c: self._set("outline_color", c))
        outline_row = QtWidgets.QHBoxLayout()
        outline_row.addWidget(self.outline_check)
        outline_row.addWidget(self.outline_btn)
        outline_row.addStretch(1)
        form.addRow("Text outline", outline_row)

        self.box_check = QtWidgets.QCheckBox("Box")
        self.box_check.setChecked(bool(val("box")))
        self.box_check.toggled.connect(lambda on: self._set("box", bool(on)))
        self.box_btn = _ColorButton(val("box_color"), "Background colour")
        self.box_btn.color_changed.connect(lambda c: self._set("box_color", c))
        self.alpha_spin = QtWidgets.QSpinBox()
        self.alpha_spin.setRange(0, 100)
        self.alpha_spin.setSuffix(" %")
        self.alpha_spin.setToolTip("Background opacity")
        self.alpha_spin.setValue(int(round(float(val("box_alpha")) * 100)))
        self.alpha_spin.valueChanged.connect(lambda v: self._set("box_alpha", v / 100.0))
        box_row = QtWidgets.QHBoxLayout()
        box_row.addWidget(self.box_check)
        box_row.addWidget(self.box_btn)
        box_row.addWidget(self.alpha_spin)
        box_row.addStretch(1)
        form.addRow("Background", box_row)

        self.pos_combo = QtWidgets.QComboBox()
        for label, key in POSITIONS:
            self.pos_combo.addItem(label, key)
        idx = self.pos_combo.findData(val("position"))
        self.pos_combo.setCurrentIndex(max(0, idx))
        self.pos_combo.currentIndexChanged.connect(
            lambda _i: self._set("position", self.pos_combo.currentData()))
        form.addRow("Corner", self.pos_combo)

        note = QtWidgets.QLabel(
            "Changes apply live. Settings left untouched keep the default look "
            "(boxed in Analysis, outlined in Publication).")
        note.setWordWrap(True)
        note.setStyleSheet("color: gray;")

        buttons = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel
            | QtWidgets.QDialogButtonBox.RestoreDefaults)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        buttons.button(QtWidgets.QDialogButtonBox.RestoreDefaults).clicked.connect(self._restore_defaults)

        layout = QtWidgets.QVBoxLayout(self)
        layout.addLayout(form)
        layout.addWidget(note)
        layout.addWidget(buttons)

    def style(self):
        return dict(self._style)

    def _set(self, key, value):
        self._style[key] = value
        self.style_changed.emit(dict(self._style))

    def _on_font_toggled(self, on):
        self.font_combo.setEnabled(bool(on))
        if on:
            self._set("font_family", self.font_combo.currentFont().family())
        else:
            self._style.pop("font_family", None)
            self.style_changed.emit(dict(self._style))

    def _restore_defaults(self):
        widgets = (self.size_spin, self.font_check, self.font_combo, self.bold_check,
                   self.outline_check, self.box_check, self.alpha_spin, self.pos_combo)
        for w in widgets:
            w.blockSignals(True)
        self.size_spin.setValue(0.0)
        self.font_check.setChecked(False)
        self.font_combo.setEnabled(False)
        self.bold_check.setChecked(DISPLAY_DEFAULTS["bold"])
        self.outline_check.setChecked(DISPLAY_DEFAULTS["outline"])
        self.box_check.setChecked(DISPLAY_DEFAULTS["box"])
        self.alpha_spin.setValue(int(DISPLAY_DEFAULTS["box_alpha"] * 100))
        self.pos_combo.setCurrentIndex(0)
        for w in widgets:
            w.blockSignals(False)
        self.color_btn.set_color(DISPLAY_DEFAULTS["color"])
        self.outline_btn.set_color(DISPLAY_DEFAULTS["outline_color"])
        self.box_btn.set_color(DISPLAY_DEFAULTS["box_color"])
        self._style = {}
        self.style_changed.emit({})

    def reject(self):
        # Cancel rolls the live preview back to the style the dialog opened with.
        self.style_changed.emit(dict(self._initial))
        super().reject()
