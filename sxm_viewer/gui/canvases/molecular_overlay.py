"""Molecular overlay support for SXM Viewer."""
from __future__ import annotations

import numpy as np
from pathlib import Path
from ..._shared import QtWidgets, QtCore, QtGui

# CPK coloring (simplified)
CPK_COLORS = {
    'H': '#FFFFFF', 'C': '#909090', 'N': '#3050F8', 'O': '#FF0D0D',
    'F': '#90E050', 'Cl': '#1FF01F', 'Br': '#A62929', 'I': '#940094',
    'He': '#40FFFF', 'Ne': '#4B5E71', 'Ar': '#80D1E3', 'Kr': '#5CB8D1',
    'Xe': '#D09E73', 'P': '#FF8000', 'S': '#FFFF30', 'B': '#FFB5B5',
    'Li': '#CC80FF', 'Na': '#AB5CF2', 'K': '#8F40D4', 'Rb': '#3D2E85',
    'Cs': '#2E1D4C', 'Be': '#C2FF00', 'Mg': '#8AFF00', 'Ca': '#4DFF00',
    'Sr': '#26FF00', 'Ba': '#00FF00', 'Ra': '#00FF00', 'Ti': '#BFC2C7',
    'Fe': '#E06633', 'Cu': '#C88033', 'Zn': '#C88033', 'Ag': '#C88033',
    'Au': '#FFD123', 'Pt': '#D0D0E0', 'Si': '#F0C8A0', 'Pb': '#575961',
}

COVALENT_RADII = {
    'H': 0.31, 'He': 0.28, 'Li': 1.28, 'Be': 0.96, 'B': 0.84, 'C': 0.76, 
    'N': 0.71, 'O': 0.66, 'F': 0.57, 'Ne': 0.58, 'Na': 1.66, 'Mg': 1.41, 
    'Al': 1.21, 'Si': 1.11, 'P': 1.07, 'S': 1.05, 'Cl': 1.02, 'Ar': 1.06, 
    'K': 2.03, 'Ca': 1.76, 'Sc': 1.70, 'Ti': 1.60, 'V': 1.53, 'Cr': 1.39, 
    'Mn': 1.39, 'Fe': 1.32, 'Co': 1.26, 'Ni': 1.24, 'Cu': 1.32, 'Zn': 1.22, 
    'Ga': 1.22, 'Ge': 1.20, 'As': 1.19, 'Se': 1.20, 'Br': 1.20, 'Kr': 1.16, 
    'Rb': 2.20, 'Sr': 1.95, 'Y': 1.90, 'Zr': 1.75, 'Nb': 1.64, 'Mo': 1.54, 
    'Tc': 1.47, 'Ru': 1.46, 'Rh': 1.42, 'Pd': 1.39, 'Ag': 1.45, 'Cd': 1.44, 
    'In': 1.42, 'Sn': 1.39, 'Sb': 1.39, 'Te': 1.38, 'I': 1.39, 'Xe': 1.40, 
    'Cs': 2.44, 'Ba': 2.15, 'La': 2.07, 'Ce': 2.04, 'Pr': 2.03, 'Nd': 2.01, 
    'Pm': 1.99, 'Sm': 1.98, 'Eu': 1.98, 'Gd': 1.96, 'Tb': 1.94, 'Dy': 1.92, 
    'Ho': 1.92, 'Er': 1.89, 'Tm': 1.90, 'Yb': 1.87, 'Lu': 1.87, 'Hf': 1.75, 
    'Ta': 1.70, 'W': 1.62, 'Re': 1.51, 'Os': 1.44, 'Ir': 1.41, 'Pt': 1.36, 
    'Au': 1.36, 'Hg': 1.32, 'Tl': 1.45, 'Pb': 1.46, 'Bi': 1.48, 'Po': 1.40, 
    'At': 1.50, 'Rn': 1.50, 'Fr': 2.60, 'Ra': 2.21, 'Ac': 2.15, 'Th': 2.06, 
    'Pa': 2.00, 'U': 1.96, 'Np': 1.90, 'Pu': 1.87, 'Am': 1.80, 'Cm': 1.69
}

def get_cpk_color(element):
    return CPK_COLORS.get(element.title(), '#FF1493')

class Molecule:
    def __init__(self, filepath=None):
        self.filepath = filepath
        self.coordinates = np.zeros((0, 3))
        self.elements = []
        self.offset = np.array([0.0, 0.0, 0.0])
        # Rotation in degrees: x (pitch), y (roll), z (yaw)
        self.angles = np.array([0.0, 0.0, 0.0]) 
        self.scale = 0.1  # Angstrom to nm default
        self.mirror_x = False
        self.mirror_y = False
        self.z_height_scale = 1.0 # For visual depth cue
        self.bonds = [] # List of (index1, index2)
        self.display_mode = 'Atoms + Bonds'
        
        if filepath:
            self.load(filepath)

    def load(self, filepath):
        path = Path(filepath)
        suffix = path.suffix.lower()
        if suffix == '.xyz':
            self._load_xyz(path)
        elif suffix == '.pdb':
            self._load_pdb(path)
        elif suffix == '.mol':
            self._load_mol(path)
        
        # Center molecule at origin initially
        if len(self.coordinates) > 0:
            center = np.mean(self.coordinates, axis=0)
            self.coordinates -= center
        self.recalculate_bonds()

    def _load_xyz(self, path):
        try:
            with open(path, 'r') as f:
                lines = f.readlines()
            # Skip header lines (count and comment)
            coords = []
            elems = []
            start = 2
            for line in lines[start:]:
                parts = line.split()
                if len(parts) >= 4:
                    elems.append(parts[0])
                    coords.append([float(parts[1]), float(parts[2]), float(parts[3])])
            self.coordinates = np.array(coords)
            self.elements = elems
        except Exception as e:
            print(f"Error loading XYZ: {e}")

    def _load_pdb(self, path):
        coords = []
        elems = []
        try:
            with open(path, 'r') as f:
                for line in f:
                    if line.startswith("ATOM") or line.startswith("HETATM"):
                        try:
                            x = float(line[30:38])
                            y = float(line[38:46])
                            z = float(line[46:54])
                            coords.append([x, y, z])
                            # Element is often in 76-78, or derived from name
                            elem = line[76:78].strip()
                            if not elem:
                                name = line[12:16].strip()
                                # heuristic for element from atom name
                                elem = ''.join([c for c in name if not c.isdigit()])[:2]
                            elems.append(elem)
                        except:
                            pass
            self.coordinates = np.array(coords)
            self.elements = elems
        except Exception as e:
            print(f"Error loading PDB: {e}")

    def _load_mol(self, path):
        coords = []
        elems = []
        try:
            with open(path, 'r') as f:
                lines = f.readlines()
            counts_idx = None
            counts = []
            for idx, line in enumerate(lines):
                parts = line.split()
                if len(parts) >= 2:
                    try:
                        int(float(parts[0]))
                        counts_idx = idx
                        counts = parts
                        break
                    except Exception:
                        continue
            if counts_idx is None:
                raise ValueError("Counts line not found")
            natoms = int(float(counts[0]))
            start = counts_idx + 1
            for i in range(natoms):
                if start + i >= len(lines):
                    break
                parts = lines[start + i].split()
                if len(parts) >= 4:
                    x = float(parts[0])
                    y = float(parts[1])
                    z = float(parts[2])
                    elem = parts[3]
                    coords.append([x, y, z])
                    elems.append(elem)
            self.coordinates = np.array(coords)
            self.elements = elems
        except Exception as e:
            print(f"Error loading MOL: {e}")

    def recalculate_bonds(self):
        self.bonds = []
        n = len(self.coordinates)
        if n < 2: return
        
        coords = self.coordinates
        radii = np.array([COVALENT_RADII.get(e, 1.5) for e in self.elements])
        
        # Simple pairwise distance check (N^2 is fine for small molecules)
        # Bond if dist < r1 + r2 + tolerance (0.45 A)
        for i in range(n):
            for j in range(i + 1, n):
                dist_sq = np.sum((coords[i] - coords[j])**2)
                thresh = (radii[i] + radii[j] + 0.45)**2
                if 0.1 < dist_sq < thresh:
                    self.bonds.append((i, j))

    def get_transformed_coordinates(self):
        if len(self.coordinates) == 0:
            return np.zeros((0, 3))
            
        # 1. Rotation
        rads = np.radians(self.angles)
        cx, cy, cz = np.cos(rads)
        sx, sy, sz = np.sin(rads)
        
        # Rx
        Rx = np.array([[1, 0, 0], [0, cx, -sx], [0, sx, cx]])
        # Ry
        Ry = np.array([[cy, 0, sy], [0, 1, 0], [-sy, 0, cy]])
        # Rz
        Rz = np.array([[cz, -sz, 0], [sz, cz, 0], [0, 0, 1]])
        
        R = Rz @ Ry @ Rx
        coords = self.coordinates @ R.T
        
        # 2. Mirror
        if self.mirror_x:
            coords[:, 0] *= -1
        if self.mirror_y:
            coords[:, 1] *= -1
            
        # 3. Scale (Angstrom -> nm)
        coords *= self.scale
        
        # 4. Translation
        coords += self.offset
        
        return coords

    def copy(self):
        new_mol = Molecule()
        new_mol.filepath = self.filepath
        new_mol.coordinates = self.coordinates.copy()
        new_mol.elements = list(self.elements)
        new_mol.offset = self.offset.copy()
        new_mol.angles = self.angles.copy()
        new_mol.scale = self.scale
        new_mol.mirror_x = self.mirror_x
        new_mol.mirror_y = self.mirror_y
        new_mol.bonds = list(self.bonds)
        new_mol.display_mode = self.display_mode
        return new_mol

class MoleculePropertiesDialog(QtWidgets.QDialog):
    def __init__(self, molecule, parent=None, callback=None):
        super().__init__(parent)
        self.molecule = molecule
        self.callback = callback
        self.setWindowTitle("Molecule Properties")
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
        if self.callback:
            self.callback()
