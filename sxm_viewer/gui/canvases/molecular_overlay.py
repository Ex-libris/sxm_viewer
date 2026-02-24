"""Molecular overlay support for SXM Viewer."""
from __future__ import annotations

import numpy as np
from pathlib import Path
from ..._shared import QtWidgets, QtCore, QtGui

# Atom palettes (simplified, hex)
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

# Pymol palette (classic)
PYMOL_COLORS = {
    'H': '#ffffff', 'C': '#c8c8c8', 'N': '#8f8fff', 'O': '#f00000',
    'S': '#ffbf00', 'P': '#ffa500', 'F': '#00ff00', 'Cl': '#00ff00',
    'Br': '#a52a2a', 'I': '#940094', 'Fe': '#ffa500', 'Cu': '#c88033',
    'Zn': '#7d7d7d'
}

# Jmol palette (subset)
JMOL_COLORS = {
    'H': '#FFFFFF', 'C': '#909090', 'N': '#3050F8', 'O': '#FF0D0D',
    'S': '#FFFF30', 'P': '#FF8000', 'F': '#90E050', 'Cl': '#1FF01F',
    'Br': '#A62929', 'I': '#940094', 'Fe': '#E06633', 'Cu': '#C88033',
    'Zn': '#7d80b0', 'Si': '#F0C8A0', 'B': '#FFB5B5'
}

# Avogadro palette (subset)
AVOGADRO_COLORS = {
    'H': '#ffffff', 'C': '#000000', 'N': '#3050f8', 'O': '#ff0d0d',
    'S': '#c8c800', 'P': '#ffa500', 'F': '#00ff00', 'Cl': '#00ff00',
    'Br': '#8a0a0a', 'I': '#940094', 'Si': '#da8b45', 'B': '#ffb5b5',
    'Fe': '#b7410e', 'Cu': '#b87333'
}

ATOM_PALETTES = {
    'cpk': CPK_COLORS,
    'pymol': PYMOL_COLORS,
    'jmol': JMOL_COLORS,
    'avogadro': AVOGADRO_COLORS,
    'ase': {
        'H': '#FFFFFF', 'C': '#404040', 'N': '#3050F8', 'O': '#FF0D0D',
        'S': '#FFFF30', 'P': '#FF8000', 'F': '#90E050', 'Cl': '#1FF01F',
        'Br': '#A62929', 'I': '#940094', 'Fe': '#E06633', 'Cu': '#C88033',
        'Zn': '#7d80b0', 'Si': '#DA8B45', 'B': '#FFB5B5'
    }
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

# van der Waals radii (Å) from Bondi/CRC (fallback to covalent if missing)
VDW_RADII = {
    'H': 1.20, 'He': 1.40, 'Li': 1.82, 'Be': 1.53, 'B': 1.92, 'C': 1.70, 'N': 1.55, 'O': 1.52,
    'F': 1.47, 'Ne': 1.54, 'Na': 2.27, 'Mg': 1.73, 'Al': 1.84, 'Si': 2.10, 'P': 1.80, 'S': 1.80,
    'Cl': 1.75, 'Ar': 1.88, 'K': 2.75, 'Ca': 2.31, 'Ni': 1.63, 'Cu': 1.40, 'Zn': 1.39,
    'Ga': 1.87, 'Ge': 2.11, 'As': 1.85, 'Se': 1.90, 'Br': 1.85, 'Kr': 2.02, 'Rb': 3.03,
    'Sr': 2.49, 'Pd': 1.63, 'Ag': 1.72, 'Cd': 1.58, 'In': 1.93, 'Sn': 2.17, 'Sb': 2.06,
    'Te': 2.06, 'I': 1.98, 'Xe': 2.16, 'Cs': 3.43, 'Ba': 2.68, 'Pt': 1.75, 'Au': 1.66, 'Hg': 1.55,
    'Tl': 1.96, 'Pb': 2.02, 'Bi': 2.07, 'Po': 1.97, 'At': 2.02, 'Rn': 2.20
}

# Neutral atomic radii (Å), CRC/Pauling
ATOMIC_RADII = {
    'H': 0.53, 'He': 0.31, 'Li': 1.67, 'Be': 1.12, 'B': 0.87, 'C': 0.67, 'N': 0.56, 'O': 0.48,
    'F': 0.42, 'Ne': 0.38, 'Na': 1.90, 'Mg': 1.45, 'Al': 1.18, 'Si': 1.11, 'P': 0.98, 'S': 0.88,
    'Cl': 0.79, 'Ar': 0.71, 'K': 2.43, 'Ca': 1.94, 'Sc': 1.84, 'Ti': 1.76, 'V': 1.71, 'Cr': 1.66,
    'Mn': 1.61, 'Fe': 1.56, 'Co': 1.52, 'Ni': 1.49, 'Cu': 1.45, 'Zn': 1.42, 'Ga': 1.36, 'Ge': 1.25,
    'As': 1.14, 'Se': 1.03, 'Br': 0.94, 'Kr': 0.88, 'Rb': 2.65, 'Sr': 2.19, 'Y': 2.12, 'Zr': 2.06,
    'Nb': 1.98, 'Mo': 1.90, 'Tc': 1.83, 'Ru': 1.78, 'Rh': 1.73, 'Pd': 1.69, 'Ag': 1.65, 'Cd': 1.61,
    'In': 1.56, 'Sn': 1.45, 'Sb': 1.33, 'Te': 1.23, 'I': 1.15, 'Xe': 1.08, 'Cs': 2.98, 'Ba': 2.53,
    'La': 1.95, 'Hf': 2.08, 'Ta': 2.00, 'W': 1.93, 'Re': 1.88, 'Os': 1.85, 'Ir': 1.80, 'Pt': 1.77,
    'Au': 1.74, 'Hg': 1.71, 'Tl': 1.56, 'Pb': 1.54, 'Bi': 1.43, 'Po': 1.35, 'At': 1.27, 'Rn': 1.20
}

def available_atom_palettes():
    """Return sorted list of available palette keys."""
    return sorted(ATOM_PALETTES.keys())

def get_atom_color(element, palette='cpk'):
    """Return color for element from selected palette (fallback to CPK)."""
    pal = ATOM_PALETTES.get(str(palette).lower()) or CPK_COLORS
    return pal.get(element.title(), CPK_COLORS.get(element.title(), '#FF1493'))

# Backwards compatible alias
def get_cpk_color(element):
    return get_atom_color(element, 'cpk')

def get_atom_radius(element, mode='covalent'):
    """Return an element radius (Å) for the requested mode with sensible fallback."""
    el = (element or '').title()
    mode = (mode or 'covalent').lower()
    if mode == 'vdw':
        val = VDW_RADII.get(el)
    elif mode == 'atomic':
        val = ATOMIC_RADII.get(el)
    elif mode == 'constant':
        val = 1.0
    else:
        val = COVALENT_RADII.get(el)
    if val is None:
        val = COVALENT_RADII.get(el) or VDW_RADII.get(el) or ATOMIC_RADII.get(el) or 1.5
    return float(val)

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
        self.display_mode = 'Bonds Only'
        self.render_style = 'classic'   # classic | flat | spacefill | licorice | wire | sticks | skeletal
        self.bond_style = 'default'    # default | thin | thick
        self.atom_color_override = None  # Hex string or None (use palette)
        self.atom_color_map = {}         # Element -> hex override
        self.bond_color_override = None  # Hex string or None (use default)
        self.bond_color_mode = 'default' # default | single | by_atoms
        self.radius_mode = 'covalent'    # covalent | vdw | atomic | constant
        self.radius_scale = 1.0
        
        # Avoid truth-testing numpy arrays or other iterables; only treat valid paths.
        if isinstance(filepath, (str, Path)) and str(filepath).strip():
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
        new_mol.render_style = self.render_style
        new_mol.bond_style = self.bond_style
        new_mol.atom_color_override = self.atom_color_override
        new_mol.atom_color_map = dict(self.atom_color_map or {})
        new_mol.bond_color_override = self.bond_color_override
        new_mol.bond_color_mode = self.bond_color_mode
        new_mol.radius_mode = self.radius_mode
        new_mol.radius_scale = self.radius_scale
        new_mol.z_height_scale = self.z_height_scale
        return new_mol

    def to_dict(self) -> dict:
        return {
            "filepath": str(self.filepath) if self.filepath else None,
            "coordinates": self.coordinates.tolist() if len(self.coordinates) else [],
            "elements": list(self.elements),
            "offset": self.offset.tolist(),
            "angles": self.angles.tolist(),
            "scale": float(self.scale),
            "mirror_x": bool(self.mirror_x),
            "mirror_y": bool(self.mirror_y),
            "z_height_scale": float(getattr(self, "z_height_scale", 1.0)),
            "bonds": [list(b) for b in (self.bonds or [])],
            "display_mode": self.display_mode,
            "render_style": self.render_style,
            "bond_style": self.bond_style,
            "atom_color_override": self.atom_color_override,
            "atom_color_map": dict(self.atom_color_map or {}),
            "bond_color_override": self.bond_color_override,
            "bond_color_mode": self.bond_color_mode,
            "radius_mode": self.radius_mode,
            "radius_scale": float(self.radius_scale),
        }

    @classmethod
    def from_dict(cls, data: dict) -> "Molecule":
        mol = cls()
        if not isinstance(data, dict):
            return mol
        mol.filepath = data.get("filepath")
        coords = data.get("coordinates") or []
        mol.coordinates = np.asarray(coords, dtype=float) if coords else np.zeros((0, 3))
        mol.elements = list(data.get("elements") or [])
        mol.offset = np.asarray(data.get("offset") or [0.0, 0.0, 0.0], dtype=float)
        mol.angles = np.asarray(data.get("angles") or [0.0, 0.0, 0.0], dtype=float)
        mol.scale = float(data.get("scale", 0.1))
        mol.mirror_x = bool(data.get("mirror_x", False))
        mol.mirror_y = bool(data.get("mirror_y", False))
        mol.z_height_scale = float(data.get("z_height_scale", 1.0))
        mol.bonds = [tuple(b) for b in (data.get("bonds") or [])]
        mol.display_mode = data.get("display_mode", "Atoms + Bonds")
        mol.render_style = data.get("render_style", "classic")
        mol.bond_style = data.get("bond_style", "default")
        mol.atom_color_override = data.get("atom_color_override")
        mol.atom_color_map = dict(data.get("atom_color_map") or {})
        mol.bond_color_override = data.get("bond_color_override")
        mol.bond_color_mode = data.get("bond_color_mode", "default")
        mol.radius_mode = data.get("radius_mode", "covalent")
        mol.radius_scale = float(data.get("radius_scale", 1.0))
        if not mol.bonds:
            mol.recalculate_bonds()
        return mol

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
        self.combo_atom_style = QtWidgets.QComboBox()
        self.combo_atom_style.addItems(["Shaded", "Flat", "Spacefill", "Licorice", "Wireframe", "Sticks", "Stick/Skeletal"])
        style_label = molecule.render_style.lower()
        if style_label == 'flat': self.combo_atom_style.setCurrentText("Flat")
        elif style_label == 'spacefill': self.combo_atom_style.setCurrentText("Spacefill")
        elif style_label == 'licorice': self.combo_atom_style.setCurrentText("Licorice")
        elif style_label == 'wire': self.combo_atom_style.setCurrentText("Wireframe")
        elif style_label == 'sticks': self.combo_atom_style.setCurrentText("Sticks")
        elif style_label == 'skeletal': self.combo_atom_style.setCurrentText("Stick/Skeletal")
        else: self.combo_atom_style.setCurrentText("Shaded")
        self.combo_atom_style.currentTextChanged.connect(self._on_change)
        form_disp.addRow("Atom style:", self.combo_atom_style)
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
        # Color overrides
        self.btn_atom_color = QtWidgets.QPushButton("Atom color...")
        self.btn_atom_color.clicked.connect(self._pick_atom_color)
        self.btn_bond_color = QtWidgets.QPushButton("Bond color...")
        self.btn_bond_color.clicked.connect(self._pick_bond_color)
        self.btn_reset_colors = QtWidgets.QPushButton("Reset colors")
        self.btn_reset_colors.clicked.connect(self._reset_colors)
        form_disp.addRow(self.btn_atom_color, self.btn_bond_color)
        form_disp.addRow(self.btn_reset_colors)
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
        style_choice = self.combo_atom_style.currentText()
        if style_choice == "Stick/Skeletal":
            self.molecule.render_style = "skeletal"
        else:
            self.molecule.render_style = style_choice.lower()
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
        if self.callback:
            self.callback()

    def _pick_atom_color(self):
        color = QtWidgets.QColorDialog.getColor(
            QtGui.QColor(self.molecule.atom_color_override or "#cccccc"),
            self,
            "Select atom color"
        )
        if color.isValid():
            self.molecule.atom_color_override = color.name()
            self._on_change()

    def _pick_bond_color(self):
        color = QtWidgets.QColorDialog.getColor(
            QtGui.QColor(self.molecule.bond_color_override or "#e0e0e0"),
            self,
            "Select bond color"
        )
        if color.isValid():
            self.molecule.bond_color_override = color.name()
            self._on_change()

    def _reset_colors(self):
        self.molecule.atom_color_override = None
        self.molecule.bond_color_override = None
        self._on_change()
