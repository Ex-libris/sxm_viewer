"""Consolidated style controls for 2D molecule overlays (SvgMoleculeOverlay).

Collects atom-color scheme, text size/contrast, and bond-order/bond-length
display into one discoverable dialog instead of scattered context-menu
toggles, and applies changes live against the actual scan data.
"""
from __future__ import annotations

import json

from ..._shared import QtCore, QtGui, QtWidgets, Path
from ..canvases.molecular_overlay import (
    available_atom_palettes,
    QUANTITATIVE_COLORMAPS,
    QUALITATIVE_COLORMAPS,
)
from ..thumbnail_render import _colormap_icon

_PALETTE_LABELS = {
    "cpk": "CPK (standard)",
    "jmol": "Jmol",
    "pymol": "PyMOL",
    "avogadro": "Avogadro",
    "ase": "ASE",
}

_BOND_COLOR_MODES = [
    ("Uniform (default)", "uniform"),
    ("By length — continuous colormap", "length_continuous"),
    ("By length — categories (short/normal/long)", "length_binned"),
    ("By bond order (single/double/triple)", "bond_order"),
]

# ColorBrewer's own documented colorblind-safe qualitative schemes
# (https://colorbrewer2.org). Others are usable but not verified safe -
# flagged rather than hidden, so the choice stays with the user.
_COLORBLIND_SAFE_QUALITATIVE = {"Dark2", "Paired", "Set2"}


class SvgMoleculeStyleDialog(QtWidgets.QDialog):
    """Non-modal style editor for one SvgMoleculeOverlay, applied live so the
    user can see changes immediately against the underlying image."""

    def __init__(self, canvas, overlay, parent=None):
        super().__init__(parent)
        self.canvas = canvas
        self.overlay = overlay
        self.setWindowTitle(f"2D Structure Style — {overlay.name or 'structure'}")
        self._flat_color = str(getattr(overlay, "flat_atom_color", "#f7fafc") or "#f7fafc")
        self._last_legend_state = bool(getattr(overlay, "show_legend", False))

        canvas.push_undo_state("svg_molecule_style")

        form = QtWidgets.QFormLayout(self)

        form.addRow(QtWidgets.QLabel("<b>Atoms</b>"))
        self.palette_combo = QtWidgets.QComboBox()
        for key in available_atom_palettes():
            self.palette_combo.addItem(_PALETTE_LABELS.get(key, key.title()), key)
        self.palette_combo.addItem("Flat / single color", "flat")
        current_mode = str(getattr(overlay, "atom_color_mode", "cpk") or "cpk")
        idx = self.palette_combo.findData(current_mode)
        self.palette_combo.setCurrentIndex(idx if idx >= 0 else 0)
        form.addRow("Atom colors:", self.palette_combo)

        self.flat_color_btn = QtWidgets.QPushButton()
        self.flat_color_btn.setToolTip("Pick a single flat color for all atoms")
        self._update_flat_color_btn()
        form.addRow("Flat color:", self.flat_color_btn)

        form.addRow(QtWidgets.QLabel("<b>Bonds</b>"))
        self.bond_order_check = QtWidgets.QCheckBox("Show bond order (double/triple bonds)")
        self.bond_order_check.setChecked(bool(getattr(overlay, "show_bond_order", False)))
        self.bond_order_check.setToolTip(
            "Reflects the source file's assumed bonding, often a gas-phase\n"
            "reference structure. Bonding and aromaticity can change once a\n"
            "molecule is adsorbed on a surface — enable only if this is meant\n"
            "as a reference/idealized structure, not a measured one."
        )
        form.addRow(self.bond_order_check)

        self.bond_length_check = QtWidgets.QCheckBox("Show bond-length labels")
        self.bond_length_check.setChecked(bool(getattr(overlay, "show_bond_length_labels", False)))
        form.addRow(self.bond_length_check)

        self.bond_color_mode_combo = QtWidgets.QComboBox()
        for label, key in _BOND_COLOR_MODES:
            self.bond_color_mode_combo.addItem(label, key)
        current_bond_mode = str(getattr(overlay, "bond_color_mode", "uniform") or "uniform")
        idx = self.bond_color_mode_combo.findData(current_bond_mode)
        self.bond_color_mode_combo.setCurrentIndex(idx if idx >= 0 else 0)
        form.addRow("Color bonds:", self.bond_color_mode_combo)

        # Quantitative colormap (continuous length deviation) - sequential
        # and diverging colormaps only, since a qualitative/categorical
        # colormap would misleadingly imply unordered categories for a
        # continuous quantity.
        self.quant_cmap_combo = QtWidgets.QComboBox()
        for name in QUANTITATIVE_COLORMAPS:
            self.quant_cmap_combo.addItem(_colormap_icon(name), name, name)
        current_quant = str(getattr(overlay, "bond_colormap", "coolwarm") or "coolwarm")
        idx = self.quant_cmap_combo.findData(current_quant)
        self.quant_cmap_combo.setCurrentIndex(idx if idx >= 0 else 0)
        form.addRow("Colormap (continuous):", self.quant_cmap_combo)

        # Qualitative colormap (discrete categories: binned length, or bond
        # order) - distinct discrete colors, no implied ordering/magnitude.
        # Colorblind-safe options (per ColorBrewer) are flagged, not hidden.
        self.qual_cmap_combo = QtWidgets.QComboBox()
        for name in QUALITATIVE_COLORMAPS:
            label = f"{name} (colorblind-safe)" if name in _COLORBLIND_SAFE_QUALITATIVE else name
            self.qual_cmap_combo.addItem(_colormap_icon(name), label, name)
        current_qual = str(getattr(overlay, "bond_qualitative_cmap", "Set2") or "Set2")
        idx = self.qual_cmap_combo.findData(current_qual)
        self.qual_cmap_combo.setCurrentIndex(idx if idx >= 0 else 0)
        form.addRow("Colormap (categories):", self.qual_cmap_combo)

        self.legend_check = QtWidgets.QCheckBox("Show color legend on the image")
        self.legend_check.setChecked(bool(getattr(overlay, "show_legend", False)))
        self.legend_check.setToolTip(
            "A small on-image key explaining what the bond colors mean - useful\n"
            "for anyone viewing a figure who doesn't know this app's conventions."
        )
        form.addRow(self.legend_check)

        form.addRow(QtWidgets.QLabel("<b>Text</b>"))
        self.font_scale_spin = QtWidgets.QDoubleSpinBox()
        self.font_scale_spin.setRange(0.5, 3.0)
        self.font_scale_spin.setSingleStep(0.1)
        self.font_scale_spin.setValue(float(getattr(overlay, "label_font_scale", 1.0) or 1.0))
        self.font_scale_spin.setSuffix("×")
        form.addRow("Text size:", self.font_scale_spin)

        self.high_contrast_check = QtWidgets.QCheckBox("Thicker text outline (busy backgrounds / low vision)")
        self.high_contrast_check.setChecked(bool(getattr(overlay, "high_contrast_labels", False)))
        form.addRow(self.high_contrast_check)

        form.addRow(QtWidgets.QLabel("<b>Presets</b>"))
        preset_row = QtWidgets.QHBoxLayout()
        self.save_default_btn = QtWidgets.QPushButton("Save as default")
        self.save_default_btn.setToolTip("Use this style for new 2D structures you load from now on")
        self.save_preset_btn = QtWidgets.QPushButton("Save preset...")
        self.save_preset_btn.setToolTip("Save this style to a file, to share or reuse later")
        self.load_preset_btn = QtWidgets.QPushButton("Load preset...")
        preset_row.addWidget(self.save_default_btn)
        preset_row.addWidget(self.save_preset_btn)
        preset_row.addWidget(self.load_preset_btn)
        form.addRow(preset_row)

        buttons = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Close)
        buttons.rejected.connect(self.close)
        buttons.accepted.connect(self.close)
        form.addRow(buttons)

        self.palette_combo.currentIndexChanged.connect(self._apply)
        self.flat_color_btn.clicked.connect(self._pick_flat_color)
        self.bond_order_check.toggled.connect(self._apply)
        self.bond_length_check.toggled.connect(self._apply)
        self.bond_color_mode_combo.currentIndexChanged.connect(self._apply)
        self.quant_cmap_combo.currentIndexChanged.connect(self._apply)
        self.qual_cmap_combo.currentIndexChanged.connect(self._apply)
        self.legend_check.toggled.connect(self._apply)
        self.font_scale_spin.valueChanged.connect(self._apply)
        self.high_contrast_check.toggled.connect(self._apply)
        self.save_default_btn.clicked.connect(self._save_as_default)
        self.save_preset_btn.clicked.connect(self._save_preset)
        self.load_preset_btn.clicked.connect(self._load_preset)
        self._update_enabled_state()

    def _update_flat_color_btn(self):
        self.flat_color_btn.setStyleSheet(f"background-color: {self._flat_color};")
        self.flat_color_btn.setText(self._flat_color)

    def _update_enabled_state(self):
        self.flat_color_btn.setEnabled(self.palette_combo.currentData() == "flat")
        bond_mode = self.bond_color_mode_combo.currentData()
        self.quant_cmap_combo.setEnabled(bond_mode == "length_continuous")
        self.qual_cmap_combo.setEnabled(bond_mode in ("length_binned", "bond_order"))
        self.legend_check.setEnabled(bond_mode != "uniform")

    def _pick_flat_color(self):
        color = QtWidgets.QColorDialog.getColor(QtGui.QColor(self._flat_color), self, "Choose atom color")
        if color.isValid():
            self._flat_color = color.name()
            self._update_flat_color_btn()
            self._apply()

    def _apply(self, *_args):
        self._update_enabled_state()
        overlay = self.overlay
        overlay.atom_color_mode = str(self.palette_combo.currentData() or "cpk")
        overlay.flat_atom_color = self._flat_color
        overlay.label_font_scale = float(self.font_scale_spin.value())
        overlay.high_contrast_labels = bool(self.high_contrast_check.isChecked())
        overlay.show_bond_order = bool(self.bond_order_check.isChecked())
        overlay.show_bond_length_labels = bool(self.bond_length_check.isChecked())
        overlay.bond_color_mode = str(self.bond_color_mode_combo.currentData() or "uniform")
        overlay.bond_colormap = str(self.quant_cmap_combo.currentData() or "coolwarm")
        overlay.bond_qualitative_cmap = str(self.qual_cmap_combo.currentData() or "Set2")
        overlay.show_legend = bool(self.legend_check.isChecked())
        # Legend content doesn't depend on any bond's live length, so unlike
        # every other control here it isn't refreshed by the fast per-frame
        # update path (see _draw_svg_overlay_legend) - only rebuilt on a full
        # redraw. Toggling it is rare/deliberate, so paying that cost here
        # is fine and keeps the hot drag path untouched.
        legend_changed = overlay.show_legend != self._last_legend_state
        self._last_legend_state = overlay.show_legend
        if legend_changed:
            self.canvas._redraw()
        else:
            self.canvas._update_svg_molecule_artists()

    def _save_as_default(self):
        self.canvas.save_svg_molecule_style_as_default(self.overlay)
        QtWidgets.QMessageBox.information(
            self, "Style saved as default",
            "New 2D structures you load will start with this style.",
        )

    def _save_preset(self):
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Save style preset", "molecule_style.json", "Style preset (*.json)",
        )
        if not path:
            return
        out_path = Path(path)
        if not out_path.suffix:
            out_path = out_path.with_suffix(".json")
        try:
            out_path.write_text(json.dumps(self.overlay.get_style_dict(), indent=2), encoding="utf-8")
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Save style preset", f"Unable to save preset:\n{exc}")

    def _load_preset(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Load style preset", "", "Style preset (*.json)",
        )
        if not path:
            return
        try:
            style = json.loads(Path(path).read_text(encoding="utf-8"))
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Load style preset", f"Unable to load preset:\n{exc}")
            return
        self.overlay.apply_style_dict(style)
        self._flat_color = str(getattr(self.overlay, "flat_atom_color", "#f7fafc") or "#f7fafc")
        self._sync_controls_from_overlay()
        self._apply()

    def _sync_controls_from_overlay(self):
        # Block signals while updating widgets: every control here is wired
        # to _apply(), which reads *all* widgets and writes *all* overlay
        # fields at once. Without blocking, the first widget's change fires
        # _apply() before the rest have been synced, so it reads their
        # still-stale values and writes them back over the just-loaded
        # overlay state - corrupting fields before their own sync line even
        # runs. One explicit _apply() call after this method (already done
        # by callers) applies everything correctly in one consistent pass.
        widgets = (
            self.palette_combo, self.bond_order_check, self.bond_length_check,
            self.bond_color_mode_combo, self.quant_cmap_combo, self.qual_cmap_combo,
            self.legend_check, self.font_scale_spin, self.high_contrast_check,
        )
        blockers = [QtCore.QSignalBlocker(w) for w in widgets]
        try:
            overlay = self.overlay
            idx = self.palette_combo.findData(str(getattr(overlay, "atom_color_mode", "cpk") or "cpk"))
            self.palette_combo.setCurrentIndex(idx if idx >= 0 else 0)
            self._update_flat_color_btn()
            self.bond_order_check.setChecked(bool(getattr(overlay, "show_bond_order", False)))
            self.bond_length_check.setChecked(bool(getattr(overlay, "show_bond_length_labels", False)))
            idx = self.bond_color_mode_combo.findData(str(getattr(overlay, "bond_color_mode", "uniform") or "uniform"))
            self.bond_color_mode_combo.setCurrentIndex(idx if idx >= 0 else 0)
            idx = self.quant_cmap_combo.findData(str(getattr(overlay, "bond_colormap", "coolwarm") or "coolwarm"))
            self.quant_cmap_combo.setCurrentIndex(idx if idx >= 0 else 0)
            idx = self.qual_cmap_combo.findData(str(getattr(overlay, "bond_qualitative_cmap", "Set2") or "Set2"))
            self.qual_cmap_combo.setCurrentIndex(idx if idx >= 0 else 0)
            self.legend_check.setChecked(bool(getattr(overlay, "show_legend", False)))
            self.font_scale_spin.setValue(float(getattr(overlay, "label_font_scale", 1.0) or 1.0))
            self.high_contrast_check.setChecked(bool(getattr(overlay, "high_contrast_labels", False)))
        finally:
            del blockers
        self._update_enabled_state()
