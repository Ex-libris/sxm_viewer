"""Molecular overlay support for SXM Viewer."""
from __future__ import annotations

import numpy as np
from pathlib import Path
from ..._shared import QtWidgets, QtCore, QtGui
from ... import cmap_registry
from ...molecules.model import (  # noqa: F401 - re-exported for existing importers
    CPK_COLORS,
    PYMOL_COLORS,
    JMOL_COLORS,
    AVOGADRO_COLORS,
    ATOM_PALETTES,
    DIVERGING_COLORMAPS,
    SEQUENTIAL_COLORMAPS,
    QUANTITATIVE_COLORMAPS,
    QUALITATIVE_COLORMAPS,
    MOLECULE_RENDER_STYLE_OPTIONS,
    _MOLECULE_RENDER_STYLE_ALIASES,
    available_molecule_render_styles,
    normalize_molecule_render_style,
    molecule_render_style_label,
    COVALENT_RADII,
    VDW_RADII,
    ATOMIC_RADII,
    available_atom_palettes,
    get_atom_color,
    get_cpk_color,
    element_key,
    element_color_override,
    resolve_atom_color,
    effective_bond_color_mode,
    molecule_elements,
    get_atom_radius,
    _clamp_opacity,
    SCREEN_RADIUS_AREA_REF,
    atom_marker_area_factor,
    Molecule,
)


class ElementColorEditor(QtWidgets.QWidget):
    """One color swatch per element present in a structure.

    Edits ``color_map`` (element symbol -> hex) in place, so a color applies
    to every atom of that element only - never to the whole structure.
    Unset elements show (and follow) ``base_color(element)`` - the active
    palette/scheme color.
    """

    changed = QtCore.pyqtSignal()

    def __init__(self, elements, color_map, base_color, parent=None, columns=3):
        super().__init__(parent)
        self._color_map = color_map
        self._base_color = base_color
        self._buttons = {}
        grid = QtWidgets.QGridLayout(self)
        grid.setContentsMargins(0, 0, 0, 0)
        grid.setHorizontalSpacing(6)
        grid.setVerticalSpacing(4)
        elems = molecule_elements(elements)
        if not elems:
            grid.addWidget(QtWidgets.QLabel("(no atoms)"), 0, 0)
        for n, el in enumerate(elems):
            btn = QtWidgets.QToolButton()
            btn.setToolButtonStyle(QtCore.Qt.ToolButtonTextBesideIcon)
            btn.setMinimumWidth(64)
            btn.clicked.connect(lambda _checked=False, e=el: self._pick(e))
            btn.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
            btn.customContextMenuRequested.connect(lambda _pos, e=el: self._reset(e))
            grid.addWidget(btn, n // columns, n % columns)
            self._buttons[el] = btn
        self.refresh()

    def set_color_map(self, color_map):
        """Rebind to a replaced map object (e.g. after loading a preset)."""
        self._color_map = color_map
        self.refresh()

    def color_for(self, element):
        override = element_color_override(self._color_map, element)
        if override:
            return override
        try:
            return self._base_color(element)
        except Exception:
            return get_atom_color(element, "cpk")

    def refresh(self):
        for el, btn in self._buttons.items():
            color = QtGui.QColor(self.color_for(el))
            pix = QtGui.QPixmap(14, 14)
            pix.fill(color)
            painter = QtGui.QPainter(pix)
            painter.setPen(QtGui.QColor("#202020"))
            painter.drawRect(0, 0, 13, 13)
            painter.end()
            btn.setIcon(QtGui.QIcon(pix))
            custom = bool(element_color_override(self._color_map, el))
            btn.setText(f"{el} *" if custom else el)
            btn.setToolTip(
                f"{el}: {color.name()} ({'custom' if custom else 'palette'})\n"
                "Click to choose a color for every " + el + " atom."
                + ("\nRight-click to revert to the palette color." if custom else "")
            )

    def _drop_key(self, element):
        key = element_key(element)
        for k in (key, key.upper(), key.lower()):
            self._color_map.pop(k, None)

    def _pick(self, element):
        color = QtWidgets.QColorDialog.getColor(
            QtGui.QColor(self.color_for(element)), self, f"Color for {element} atoms"
        )
        if color.isValid():
            self._drop_key(element)
            self._color_map[element_key(element)] = color.name()
            self.refresh()
            self.changed.emit()

    def _reset(self, element):
        if element_color_override(self._color_map, element):
            self._drop_key(element)
            self.refresh()
            self.changed.emit()

    def reset_all(self):
        self._color_map.clear()
        self.refresh()


class MoleculePropertiesDialog(QtWidgets.QDialog):
    def __init__(self, molecule, parent=None, callback=None, overlay_settings=None):
        super().__init__(parent)
        self.molecule = molecule
        self.callback = callback
        self.overlay_settings = overlay_settings if isinstance(overlay_settings, dict) else None
        self._show_shadow_option = bool((self.overlay_settings or {}).get("show_shadows_available", True))
        self._show_hydrogen_option = bool((self.overlay_settings or {}).get("show_hydrogens_available", True))
        self._show_palette_option = bool((self.overlay_settings or {}).get("palette_available", True))
        self._save_default_callback = (self.overlay_settings or {}).get("save_default_callback")
        self.setWindowTitle("Molecule Appearance & Transform")
        self.setWindowFlags(QtCore.Qt.Tool)
        
        layout = QtWidgets.QVBoxLayout(self)
        
        # Display
        grp_disp = QtWidgets.QGroupBox("Display")
        form_disp = QtWidgets.QFormLayout(grp_disp)
        self.combo_mode = QtWidgets.QComboBox()
        self.combo_mode.addItems(["Atoms + Bonds", "Atoms Only", "Bonds Only"])
        self.combo_mode.setCurrentText(molecule.display_mode)
        self.combo_mode.currentTextChanged.connect(self._on_change)
        form_disp.addRow("Style:", self.combo_mode)
        self.combo_atom_style = QtWidgets.QComboBox()
        self.combo_atom_style.addItems(available_molecule_render_styles())
        self.combo_atom_style.setCurrentText(molecule_render_style_label(molecule.render_style))
        self.combo_atom_style.currentTextChanged.connect(self._on_change)
        self.combo_atom_style.setToolTip(
            "Choose the molecule representation.\n"
            "Shaded/Flat are general-purpose; Ball-and-Stick and CPK are chemistry-style; "
            "Outline and Ghost are optimized for busy scan backgrounds."
        )
        form_disp.addRow("Representation:", self.combo_atom_style)
        self.combo_bond_style = QtWidgets.QComboBox()
        self.combo_bond_style.addItems(["Default", "Thin", "Thick"])
        bs = molecule.bond_style.lower()
        if bs == 'thin': self.combo_bond_style.setCurrentText("Thin")
        elif bs == 'thick': self.combo_bond_style.setCurrentText("Thick")
        else: self.combo_bond_style.setCurrentText("Default")
        self.combo_bond_style.currentTextChanged.connect(self._on_change)
        form_disp.addRow("Bond style:", self.combo_bond_style)
        self.combo_radius_mode = QtWidgets.QComboBox()
        self.combo_radius_mode.addItems(["Covalent", "van der Waals", "Atomic", "Constant"])
        rm = (molecule.radius_mode or "covalent").lower()
        if rm == 'vdw': self.combo_radius_mode.setCurrentText("van der Waals")
        elif rm == 'atomic': self.combo_radius_mode.setCurrentText("Atomic")
        elif rm == 'constant': self.combo_radius_mode.setCurrentText("Constant")
        else: self.combo_radius_mode.setCurrentText("Covalent")
        self.combo_radius_mode.currentTextChanged.connect(self._on_change)
        form_disp.addRow("Radius model:", self.combo_radius_mode)
        self.spin_radius_scale = QtWidgets.QDoubleSpinBox()
        self.spin_radius_scale.setRange(0.1, 5.0)
        self.spin_radius_scale.setSingleStep(0.05)
        self.spin_radius_scale.setValue(molecule.radius_scale)
        self.spin_radius_scale.valueChanged.connect(self._on_change)
        form_disp.addRow("Radius scale:", self.spin_radius_scale)
        self.chk_space_filling = QtWidgets.QCheckBox("Space-filling (true atomic size)")
        self.chk_space_filling.setChecked(bool(getattr(molecule, "space_filling", False)))
        self.chk_space_filling.setToolTip(
            "Draw each atom at its real radius (from the radius model x radius scale)\n"
            "in image units, so atoms touch/overlap like a space-filling model and\n"
            "grow/shrink with zoom. Off: atoms keep a fixed on-screen size."
        )
        self.chk_space_filling.toggled.connect(self._on_change)
        form_disp.addRow("Atom size:", self.chk_space_filling)
        self.slider_atom_opacity, row_atom_op = self._make_opacity_row(getattr(molecule, "atom_opacity", 1.0))
        form_disp.addRow("Atom opacity:", row_atom_op)
        self.slider_bond_opacity, row_bond_op = self._make_opacity_row(getattr(molecule, "bond_opacity", 1.0))
        form_disp.addRow("Bond opacity:", row_bond_op)
        self.chk_depth_cue = QtWidgets.QCheckBox("Depth cue (fade/shrink far atoms)")
        self.chk_depth_cue.setChecked(bool(getattr(molecule, "depth_cue", True)))
        self.chk_depth_cue.setToolTip(
            "Atoms and bonds further from the viewer (lower Z after rotation) are drawn\n"
            "more transparent and smaller. Turn off for uniform opacity/size."
        )
        self.chk_depth_cue.toggled.connect(self._on_change)
        form_disp.addRow("", self.chk_depth_cue)
        self.lbl_atoms_hidden = QtWidgets.QLabel(
            "Atoms are hidden in the 'Bonds Only' style - radius, atom size and\n"
            "atom opacity only take effect with 'Atoms + Bonds' or 'Atoms Only'."
        )
        self.lbl_atoms_hidden.setStyleSheet("color: #c98a1b;")
        form_disp.addRow(self.lbl_atoms_hidden)
        self.combo_atom_style.currentTextChanged.connect(self._on_representation_changed)
        self.combo_palette = None
        if self._show_palette_option:
            self.combo_palette = QtWidgets.QComboBox()
            for pal in available_atom_palettes():
                self.combo_palette.addItem(pal.title(), pal)
            current_palette = str((self.overlay_settings or {}).get("palette", "avogadro") or "avogadro").lower()
            current_idx = max(0, self.combo_palette.findData(current_palette))
            self.combo_palette.setCurrentIndex(current_idx)
            self.combo_palette.currentTextChanged.connect(self._on_change)
            self.combo_palette.setToolTip("Choose the atom color palette used when no per-atom override is set.")
            form_disp.addRow("Palette:", self.combo_palette)
        self.chk_show_shadows = None
        self.chk_show_hydrogens = None
        if self._show_shadow_option or self._show_hydrogen_option:
            row_vis = QtWidgets.QHBoxLayout()
            if self._show_shadow_option:
                self.chk_show_shadows = QtWidgets.QCheckBox("Show shadows")
                self.chk_show_shadows.setChecked(bool((self.overlay_settings or {}).get("show_shadows", True)))
                self.chk_show_shadows.toggled.connect(self._on_change)
                row_vis.addWidget(self.chk_show_shadows)
            if self._show_hydrogen_option:
                self.chk_show_hydrogens = QtWidgets.QCheckBox("Show hydrogens")
                self.chk_show_hydrogens.setChecked(bool((self.overlay_settings or {}).get("show_hydrogens", True)))
                self.chk_show_hydrogens.toggled.connect(self._on_change)
                row_vis.addWidget(self.chk_show_hydrogens)
            row_vis.addStretch(1)
            form_disp.addRow("Canvas:", row_vis)
        # Per-element atom colors. A legacy whole-molecule override
        # (atom_color_override, from before per-element colors existed) is
        # migrated on touch into an equivalent per-element map so it stays
        # visible and editable here, element by element.
        if not isinstance(molecule.atom_color_map, dict):
            molecule.atom_color_map = dict(molecule.atom_color_map or {})
        if molecule.atom_color_override:
            for el in molecule_elements(molecule.elements):
                if not element_color_override(molecule.atom_color_map, el):
                    molecule.atom_color_map[el] = molecule.atom_color_override
            molecule.atom_color_override = None
        self.element_colors = ElementColorEditor(
            molecule.elements, molecule.atom_color_map,
            lambda el: get_atom_color(el, self._current_palette()), parent=grp_disp
        )
        self.element_colors.setToolTip(
            "Color per element: click an element to recolor all its atoms;\n"
            "right-click to revert it to the palette color."
        )
        self.element_colors.changed.connect(self._on_change)
        form_disp.addRow("Element colors:", self.element_colors)
        # Bond color
        bond_row = QtWidgets.QHBoxLayout()
        self.combo_bond_color = QtWidgets.QComboBox()
        self.combo_bond_color.addItem("Automatic (style default)", "default")
        self.combo_bond_color.addItem("Match atom colors", "by_atoms")
        self.combo_bond_color.addItem("Custom color", "single")
        bond_idx = self.combo_bond_color.findData(effective_bond_color_mode(molecule))
        self.combo_bond_color.setCurrentIndex(max(0, bond_idx))
        self.combo_bond_color.currentIndexChanged.connect(self._on_bond_color_mode_changed)
        self.btn_bond_color = QtWidgets.QToolButton()
        self.btn_bond_color.setText("Pick...")
        self.btn_bond_color.clicked.connect(self._pick_bond_color)
        bond_row.addWidget(self.combo_bond_color, 1)
        bond_row.addWidget(self.btn_bond_color)
        form_disp.addRow("Bond color:", bond_row)
        self.btn_reset_colors = QtWidgets.QPushButton("Reset colors")
        self.btn_reset_colors.setToolTip("Revert all element colors to the palette and bonds to the style default.")
        self.btn_reset_colors.clicked.connect(self._reset_colors)
        form_disp.addRow(self.btn_reset_colors)
        self._update_bond_color_button()
        if callable(self._save_default_callback):
            self.btn_save_default = QtWidgets.QPushButton("Save as default for new molecules")
            self.btn_save_default.clicked.connect(self._save_as_default)
            form_disp.addRow(self.btn_save_default)
        layout.addWidget(grp_disp)

        # Rotation
        grp_rot = QtWidgets.QGroupBox("Rotation (deg)")
        form_rot = QtWidgets.QFormLayout(grp_rot)
        self.spin_x = self._make_spin(molecule.angles[0], -360, 360)
        self.spin_y = self._make_spin(molecule.angles[1], -360, 360)
        self.spin_z = self._make_spin(molecule.angles[2], -360, 360)
        form_rot.addRow("X (Pitch):", self.spin_x)
        form_rot.addRow("Y (Roll):", self.spin_y)
        form_rot.addRow("Z (Yaw):", self.spin_z)
        layout.addWidget(grp_rot)
        
        # Scale
        grp_scale = QtWidgets.QGroupBox("Scale")
        form_scale = QtWidgets.QFormLayout(grp_scale)
        self.spin_scale = self._make_spin(molecule.scale, 0.001, 100.0, step=0.01)
        form_scale.addRow("Factor:", self.spin_scale)
        layout.addWidget(grp_scale)
        
        # Mirror
        grp_mirror = QtWidgets.QGroupBox("Mirror")
        hbox_mirror = QtWidgets.QHBoxLayout(grp_mirror)
        self.chk_mx = QtWidgets.QCheckBox("X-Axis")
        self.chk_mx.setChecked(molecule.mirror_x)
        self.chk_my = QtWidgets.QCheckBox("Y-Axis")
        self.chk_my.setChecked(molecule.mirror_y)
        self.chk_mx.toggled.connect(self._on_change)
        self.chk_my.toggled.connect(self._on_change)
        hbox_mirror.addWidget(self.chk_mx)
        hbox_mirror.addWidget(self.chk_my)
        layout.addWidget(grp_mirror)
        
        # Buttons
        bbox = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Close)
        bbox.rejected.connect(self.accept)
        layout.addWidget(bbox)
        self._sync_atom_controls_enabled()

    def _make_opacity_row(self, value):
        slider = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        slider.setRange(0, 100)
        slider.setValue(int(round(_clamp_opacity(value) * 100)))
        label = QtWidgets.QLabel(f"{slider.value()}%")
        label.setMinimumWidth(36)
        slider.valueChanged.connect(lambda v, lbl=label: lbl.setText(f"{v}%"))
        slider.valueChanged.connect(self._on_change)
        row = QtWidgets.QHBoxLayout()
        row.addWidget(slider, 1)
        row.addWidget(label)
        return slider, row

    def _sync_atom_controls_enabled(self):
        atoms_shown = "atoms" in self.combo_mode.currentText().lower()
        for widget in (self.combo_radius_mode, self.spin_radius_scale,
                       self.chk_space_filling, self.slider_atom_opacity):
            widget.setEnabled(atoms_shown)
        self.lbl_atoms_hidden.setVisible(not atoms_shown)

    def _on_representation_changed(self, text):
        # A space-filling representation is only meaningful with atoms drawn
        # at their true size; switch the related controls along with it
        # (the user can still override them afterwards).
        if normalize_molecule_render_style(text) != "cpk":
            return
        widgets = (self.combo_mode, self.chk_space_filling)
        for w in widgets:
            w.blockSignals(True)
        try:
            if "atoms" not in self.combo_mode.currentText().lower():
                self.combo_mode.setCurrentText("Atoms + Bonds")
            self.chk_space_filling.setChecked(True)
        finally:
            for w in widgets:
                w.blockSignals(False)
        self._on_change()

    def _make_spin(self, val, min_val, max_val, step=1.0):
        sb = QtWidgets.QDoubleSpinBox()
        sb.setRange(min_val, max_val)
        sb.setSingleStep(step)
        sb.setValue(val)
        sb.valueChanged.connect(self._on_change)
        return sb
        
    def _on_change(self):
        self.molecule.angles = np.array([
            self.spin_x.value(),
            self.spin_y.value(),
            self.spin_z.value()
        ])
        self.molecule.scale = self.spin_scale.value()
        self.molecule.mirror_x = self.chk_mx.isChecked()
        self.molecule.mirror_y = self.chk_my.isChecked()
        self.molecule.display_mode = self.combo_mode.currentText()
        self.molecule.render_style = normalize_molecule_render_style(self.combo_atom_style.currentText())
        self.molecule.bond_style = self.combo_bond_style.currentText().lower()
        rm_choice = self.combo_radius_mode.currentText()
        if rm_choice.lower().startswith("van"):
            self.molecule.radius_mode = "vdw"
        elif rm_choice.lower().startswith("atomic"):
            self.molecule.radius_mode = "atomic"
        elif rm_choice.lower().startswith("constant"):
            self.molecule.radius_mode = "constant"
        else:
            self.molecule.radius_mode = "covalent"
        self.molecule.radius_scale = self.spin_radius_scale.value()
        self.molecule.space_filling = bool(self.chk_space_filling.isChecked())
        self.molecule.atom_opacity = self.slider_atom_opacity.value() / 100.0
        self.molecule.bond_opacity = self.slider_bond_opacity.value() / 100.0
        self.molecule.depth_cue = bool(self.chk_depth_cue.isChecked())
        self._sync_atom_controls_enabled()
        self.element_colors.refresh()  # unset elements follow the palette
        if self.overlay_settings is not None:
            if self.combo_palette is not None:
                self.overlay_settings["palette"] = str(self.combo_palette.currentData() or "avogadro").lower()
            if self.chk_show_shadows is not None:
                self.overlay_settings["show_shadows"] = bool(self.chk_show_shadows.isChecked())
            if self.chk_show_hydrogens is not None:
                self.overlay_settings["show_hydrogens"] = bool(self.chk_show_hydrogens.isChecked())
        if self.callback:
            self.callback()

    def _current_default_style(self):
        style = {
            "display_mode": self.molecule.display_mode,
            "render_style": normalize_molecule_render_style(self.molecule.render_style),
            "bond_style": str(self.molecule.bond_style or "default").lower(),
            "radius_mode": str(self.molecule.radius_mode or "covalent").lower(),
            "radius_scale": float(self.molecule.radius_scale),
            "space_filling": bool(self.molecule.space_filling),
            "atom_opacity": float(self.molecule.atom_opacity),
            "bond_opacity": float(self.molecule.bond_opacity),
            "depth_cue": bool(self.molecule.depth_cue),
            "atom_color_override": self.molecule.atom_color_override,
            "bond_color_override": self.molecule.bond_color_override,
            "bond_color_mode": self.molecule.bond_color_mode,
            "atom_color_map": dict(self.molecule.atom_color_map or {}),
        }
        if self.overlay_settings is not None:
            if self.combo_palette is not None:
                style["palette"] = str(self.combo_palette.currentData() or self.overlay_settings.get("palette", "avogadro")).lower()
            else:
                style["palette"] = str(self.overlay_settings.get("palette", "avogadro") or "avogadro").lower()
            style["show_shadows"] = bool(self.overlay_settings.get("show_shadows", True))
            style["show_hydrogens"] = bool(self.overlay_settings.get("show_hydrogens", True))
        return style

    def _save_as_default(self):
        self._on_change()
        if not callable(self._save_default_callback):
            return
        ok = False
        try:
            ok = bool(self._save_default_callback(self._current_default_style()))
        except Exception:
            ok = False
        if ok:
            QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), "Default molecule style saved", self, self.rect(), 1800)
        else:
            QtWidgets.QMessageBox.warning(self, "Molecule defaults", "Unable to save the default molecule style.")

    def _current_palette(self):
        if self.combo_palette is not None:
            return str(self.combo_palette.currentData() or "avogadro").lower()
        return str((self.overlay_settings or {}).get("palette", "avogadro") or "avogadro").lower()

    def _update_bond_color_button(self):
        custom = self.combo_bond_color.currentData() == "single"
        color = self.molecule.bond_color_override or "#e0e0e0"
        pix = QtGui.QPixmap(14, 14)
        pix.fill(QtGui.QColor(color))
        self.btn_bond_color.setIcon(QtGui.QIcon(pix))
        self.btn_bond_color.setToolButtonStyle(QtCore.Qt.ToolButtonTextBesideIcon)
        self.btn_bond_color.setToolTip(f"Custom bond color ({color})" if custom else "Pick a custom bond color")

    def _on_bond_color_mode_changed(self, _idx=None):
        mode = str(self.combo_bond_color.currentData() or "default")
        if mode == "single" and not self.molecule.bond_color_override:
            # Choosing "Custom" without a color yet: ask for one right away.
            self._pick_bond_color()
            return
        self.molecule.bond_color_mode = mode
        self._update_bond_color_button()
        self._on_change()

    def _pick_bond_color(self):
        color = QtWidgets.QColorDialog.getColor(
            QtGui.QColor(self.molecule.bond_color_override or "#e0e0e0"),
            self,
            "Select bond color"
        )
        if color.isValid():
            self.molecule.bond_color_override = color.name()
            self.molecule.bond_color_mode = "single"
        elif not self.molecule.bond_color_override:
            self.molecule.bond_color_mode = "default"
        blocker = QtCore.QSignalBlocker(self.combo_bond_color)
        self.combo_bond_color.setCurrentIndex(
            max(0, self.combo_bond_color.findData(effective_bond_color_mode(self.molecule)))
        )
        del blocker
        self._update_bond_color_button()
        self._on_change()

    def _reset_colors(self):
        self.molecule.atom_color_override = None
        self.element_colors.reset_all()
        self.molecule.bond_color_override = None
        self.molecule.bond_color_mode = "default"
        blocker = QtCore.QSignalBlocker(self.combo_bond_color)
        self.combo_bond_color.setCurrentIndex(0)
        del blocker
        self._update_bond_color_button()
        self._on_change()
