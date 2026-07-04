"""Independent 2D structure overlay support for SXM Viewer."""
from __future__ import annotations

from dataclasses import dataclass
import math
from pathlib import Path
import re
import xml.etree.ElementTree as ET

import numpy as np

_SVG_NS = "{http://www.w3.org/2000/svg}"
_PERIODIC_SYMBOLS = {
    "H", "He",
    "Li", "Be", "B", "C", "N", "O", "F", "Ne",
    "Na", "Mg", "Al", "Si", "P", "S", "Cl", "Ar",
    "K", "Ca", "Sc", "Ti", "V", "Cr", "Mn", "Fe", "Co", "Ni", "Cu", "Zn",
    "Ga", "Ge", "As", "Se", "Br", "Kr",
    "Rb", "Sr", "Y", "Zr", "Nb", "Mo", "Tc", "Ru", "Rh", "Pd", "Ag", "Cd",
    "In", "Sn", "Sb", "Te", "I", "Xe",
    "Cs", "Ba", "La", "Ce", "Pr", "Nd", "Pm", "Sm", "Eu", "Gd", "Tb", "Dy",
    "Ho", "Er", "Tm", "Yb", "Lu",
    "Hf", "Ta", "W", "Re", "Os", "Ir", "Pt", "Au", "Hg",
    "Tl", "Pb", "Bi", "Po", "At", "Rn",
}
_ATOMIC_NUMBERS = {
    1: "H", 2: "He", 3: "Li", 4: "Be", 5: "B", 6: "C", 7: "N", 8: "O", 9: "F", 10: "Ne",
    11: "Na", 12: "Mg", 13: "Al", 14: "Si", 15: "P", 16: "S", 17: "Cl", 18: "Ar",
    19: "K", 20: "Ca", 21: "Sc", 22: "Ti", 23: "V", 24: "Cr", 25: "Mn", 26: "Fe",
    27: "Co", 28: "Ni", 29: "Cu", 30: "Zn", 31: "Ga", 32: "Ge", 33: "As", 34: "Se",
    35: "Br", 36: "Kr", 47: "Ag", 53: "I", 79: "Au",
}
_TRANSFORM_RE = re.compile(r"([a-zA-Z]+)\s*\(([^)]*)\)")
_NUMBER_RE = re.compile(r"[-+]?(?:\d*\.\d+|\d+)(?:[eE][-+]?\d+)?")
_PATH_TOKEN_RE = re.compile(r"[a-zA-Z]|[-+]?(?:\d*\.\d+|\d+)(?:[eE][-+]?\d+)?")
_ELEMENT_RE = re.compile(r"([A-Z][a-z]?)")
_MOL_COUNTS_RE = re.compile(r"^\s*(\d+)\s+(\d+)\b.*\bV2000\b", re.IGNORECASE)
_SUPPORTED_2D_STRUCTURE_SUFFIXES = {".svg", ".mol", ".sdf", ".xyz", ".cdxml", ".cml"}


def _strip_ns(tag: str) -> str:
    if not tag:
        return ""
    return tag.split("}", 1)[-1]


def _parse_float(text, default=None):
    if text is None:
        return default
    match = _NUMBER_RE.search(str(text))
    if match is None:
        return default
    try:
        return float(match.group(0))
    except Exception:
        return default


def _parse_float_list(text) -> list[float]:
    if text is None:
        return []
    values = []
    for match in _NUMBER_RE.finditer(str(text)):
        try:
            values.append(float(match.group(0)))
        except Exception:
            continue
    return values


def _identity_transform() -> np.ndarray:
    return np.eye(3, dtype=float)


def _translate(tx: float, ty: float) -> np.ndarray:
    mat = _identity_transform()
    mat[0, 2] = float(tx)
    mat[1, 2] = float(ty)
    return mat


def _scale(sx: float, sy: float | None = None) -> np.ndarray:
    if sy is None:
        sy = sx
    mat = _identity_transform()
    mat[0, 0] = float(sx)
    mat[1, 1] = float(sy)
    return mat


def _rotate(angle_deg: float, cx: float = 0.0, cy: float = 0.0) -> np.ndarray:
    ang = math.radians(float(angle_deg))
    c = math.cos(ang)
    s = math.sin(ang)
    rot = np.array(
        [
            [c, -s, 0.0],
            [s, c, 0.0],
            [0.0, 0.0, 1.0],
        ],
        dtype=float,
    )
    if abs(cx) < 1e-12 and abs(cy) < 1e-12:
        return rot
    return _translate(cx, cy) @ rot @ _translate(-cx, -cy)


def parse_svg_transform(transform_text) -> np.ndarray:
    if not transform_text:
        return _identity_transform()
    current = _identity_transform()
    for name, args_text in _TRANSFORM_RE.findall(str(transform_text)):
        args = _parse_float_list(args_text)
        op = name.strip().lower()
        if op == "translate":
            if args:
                current = current @ _translate(args[0], args[1] if len(args) > 1 else 0.0)
        elif op == "scale":
            if args:
                current = current @ _scale(args[0], args[1] if len(args) > 1 else None)
        elif op == "rotate":
            if args:
                cx = args[1] if len(args) > 2 else 0.0
                cy = args[2] if len(args) > 2 else 0.0
                current = current @ _rotate(args[0], cx, cy)
        elif op == "matrix" and len(args) >= 6:
            a, b, c, d, e, f = args[:6]
            mat = np.array(
                [
                    [a, c, e],
                    [b, d, f],
                    [0.0, 0.0, 1.0],
                ],
                dtype=float,
            )
            current = current @ mat
    return current


def _apply_transform(point, transform: np.ndarray) -> np.ndarray:
    vec = np.array([float(point[0]), float(point[1]), 1.0], dtype=float)
    mapped = transform @ vec
    return mapped[:2]


def _path_points(d_attr: str) -> list[np.ndarray]:
    if not d_attr:
        return []
    tokens = _PATH_TOKEN_RE.findall(str(d_attr))
    if not tokens:
        return []
    out: list[np.ndarray] = []
    idx = 0
    cmd = None
    current = np.array([0.0, 0.0], dtype=float)
    start = None
    while idx < len(tokens):
        token = tokens[idx]
        if token.isalpha():
            cmd = token
            idx += 1
            if cmd in ("Z", "z") and start is not None:
                out.append(np.array(start, dtype=float))
                current = np.array(start, dtype=float)
            continue
        if cmd is None:
            idx += 1
            continue
        lower = cmd.lower()
        relative = cmd.islower()

        def _take(n: int) -> list[float]:
            nonlocal idx
            vals = []
            for _ in range(n):
                if idx >= len(tokens):
                    break
                try:
                    vals.append(float(tokens[idx]))
                except Exception:
                    break
                idx += 1
            return vals

        if lower in ("m", "l", "t"):
            vals = _take(2)
            if len(vals) < 2:
                break
            point = np.array(vals[:2], dtype=float)
            if relative:
                point = current + point
            current = point
            if start is None or lower == "m":
                start = tuple(point.tolist())
                cmd = "l" if relative else "L"
            out.append(point.copy())
            continue
        if lower == "h":
            vals = _take(1)
            if not vals:
                break
            x_val = current[0] + vals[0] if relative else vals[0]
            current = np.array([x_val, current[1]], dtype=float)
            out.append(current.copy())
            continue
        if lower == "v":
            vals = _take(1)
            if not vals:
                break
            y_val = current[1] + vals[0] if relative else vals[0]
            current = np.array([current[0], y_val], dtype=float)
            out.append(current.copy())
            continue
        if lower in ("s", "q"):
            vals = _take(4)
            if len(vals) < 4:
                break
            point = np.array(vals[2:4], dtype=float)
            if relative:
                point = current + point
            current = point
            out.append(point.copy())
            continue
        if lower == "c":
            vals = _take(6)
            if len(vals) < 6:
                break
            point = np.array(vals[4:6], dtype=float)
            if relative:
                point = current + point
            current = point
            out.append(point.copy())
            continue
        if lower == "a":
            vals = _take(7)
            if len(vals) < 7:
                break
            point = np.array(vals[5:7], dtype=float)
            if relative:
                point = current + point
            current = point
            out.append(point.copy())
            continue
        idx += 1
    return out


def _segment_from_points(points: list[np.ndarray]) -> tuple[np.ndarray, np.ndarray] | None:
    if len(points) < 2:
        return None
    pts = np.asarray(points, dtype=float)
    if pts.ndim != 2 or pts.shape[0] < 2:
        return None
    center = np.mean(pts, axis=0)
    shifted = pts - center
    try:
        _, _, vh = np.linalg.svd(shifted, full_matrices=False)
        axis = np.asarray(vh[0], dtype=float)
    except Exception:
        axis = shifted[-1] - shifted[0]
    norm = float(np.linalg.norm(axis))
    if norm < 1e-12:
        return None
    axis = axis / norm
    proj = shifted @ axis
    p0 = center + (axis * float(np.min(proj)))
    p1 = center + (axis * float(np.max(proj)))
    if float(np.linalg.norm(p1 - p0)) < 1e-6:
        return None
    return p0, p1


def _element_from_text(text: str, default: str = "C") -> str:
    raw = str(text or "").strip()
    if not raw:
        return default
    match = _ELEMENT_RE.search(raw)
    if match is None:
        return default
    symbol = match.group(1).title()
    if symbol in _PERIODIC_SYMBOLS:
        return symbol
    return default


def _element_from_atomic_number(value, default: str = "C") -> str:
    try:
        number = int(value)
    except Exception:
        return default
    return _ATOMIC_NUMBERS.get(number, default)


def _bond_order_from_text(value, default: int = 1) -> int:
    raw = str(value or "").strip().upper()
    if not raw:
        return int(default)
    mapping = {
        "1": 1,
        "2": 2,
        "3": 3,
        "4": 4,
        "S": 1,
        "D": 2,
        "T": 3,
        "A": 1,
        "AR": 1,
    }
    try:
        return max(1, int(mapping.get(raw, int(float(raw)))))
    except Exception:
        return int(default)


def _bond_component_map(num_atoms: int, bonds: list["SvgBond"]) -> list[list[int]]:
    adjacency = [[] for _ in range(max(0, int(num_atoms)))]
    for bond in bonds or []:
        try:
            i_val = int(bond.atom1)
            j_val = int(bond.atom2)
        except Exception:
            continue
        if not (0 <= i_val < len(adjacency) and 0 <= j_val < len(adjacency)):
            continue
        adjacency[i_val].append(j_val)
        adjacency[j_val].append(i_val)
    seen = set()
    components = []
    for start in range(len(adjacency)):
        if start in seen:
            continue
        stack = [start]
        comp = []
        seen.add(start)
        while stack:
            idx = stack.pop()
            comp.append(idx)
            for nxt in adjacency[idx]:
                if nxt in seen:
                    continue
                seen.add(nxt)
                stack.append(nxt)
        components.append(comp)
    return components


def _remove_atoms_and_reindex(atoms: list["SvgAtom"], bonds: list["SvgBond"], remove_ids: set[int]) -> tuple[list["SvgAtom"], list["SvgBond"]]:
    if not remove_ids:
        return atoms, bonds
    index_map = {}
    kept_atoms = []
    for idx, atom in enumerate(atoms):
        if idx in remove_ids:
            continue
        index_map[idx] = len(kept_atoms)
        kept_atoms.append(atom)
    kept_bonds = []
    for bond in bonds:
        try:
            i_val = int(bond.atom1)
            j_val = int(bond.atom2)
        except Exception:
            continue
        if i_val in remove_ids or j_val in remove_ids:
            continue
        if i_val not in index_map or j_val not in index_map:
            continue
        kept_bonds.append(
            SvgBond(
                atom1=index_map[i_val],
                atom2=index_map[j_val],
                source_kind=str(getattr(bond, "source_kind", "parsed") or "parsed"),
                order=max(1, int(getattr(bond, "order", 1) or 1)),
            )
        )
    return kept_atoms, kept_bonds


def _prune_svg_auxiliary_bond_components(atoms: list["SvgAtom"], bonds: list["SvgBond"], reference_scale: float) -> tuple[list["SvgAtom"], list["SvgBond"]]:
    if len(atoms) < 4 or len(bonds) < 2:
        return atoms, bonds
    scale = max(float(reference_scale or 0.0), 1e-6)
    close_threshold = 0.58 * scale
    remove_ids: set[int] = set()
    for comp in _bond_component_map(len(atoms), bonds):
        if len(comp) != 2:
            continue
        a_idx, b_idx = sorted(int(v) for v in comp)
        atom_a = atoms[a_idx]
        atom_b = atoms[b_idx]
        if str(getattr(atom_a, "source_kind", "")).lower() != "inferred":
            continue
        if str(getattr(atom_b, "source_kind", "")).lower() != "inferred":
            continue
        if any(
            not (int(bond.atom1) in comp and int(bond.atom2) in comp)
            for bond in bonds
            if int(getattr(bond, "atom1", -1)) in comp or int(getattr(bond, "atom2", -1)) in comp
        ):
            continue
        keep_component = False
        for idx in comp:
            point = np.array([float(atoms[idx].x), float(atoms[idx].y)], dtype=float)
            nearest_outside = None
            for jdx, other in enumerate(atoms):
                if jdx in comp:
                    continue
                dist = float(math.hypot(float(other.x) - point[0], float(other.y) - point[1]))
                if nearest_outside is None or dist < nearest_outside:
                    nearest_outside = dist
            if nearest_outside is None or nearest_outside > close_threshold:
                keep_component = True
                break
        if not keep_component:
            remove_ids.update(comp)
    if not remove_ids:
        return atoms, bonds
    return _remove_atoms_and_reindex(atoms, bonds, remove_ids)


@dataclass
class SvgAtom:
    element: str
    x: float
    y: float
    label_visible: bool = False
    source_kind: str = "inferred"

    def to_dict(self) -> dict:
        return {
            "element": str(self.element or "C"),
            "x": float(self.x),
            "y": float(self.y),
            "label_visible": bool(self.label_visible),
            "source_kind": str(self.source_kind or "inferred"),
        }

    @classmethod
    def from_dict(cls, data: dict):
        return cls(
            element=str((data or {}).get("element") or "C"),
            x=float((data or {}).get("x", 0.0) or 0.0),
            y=float((data or {}).get("y", 0.0) or 0.0),
            label_visible=bool((data or {}).get("label_visible", False)),
            source_kind=str((data or {}).get("source_kind") or "inferred"),
        )


@dataclass
class SvgBond:
    atom1: int
    atom2: int
    source_kind: str = "parsed"
    order: int = 1

    def to_dict(self) -> dict:
        return {
            "atom1": int(self.atom1),
            "atom2": int(self.atom2),
            "source_kind": str(self.source_kind or "parsed"),
            "order": max(1, int(self.order or 1)),
        }

    @classmethod
    def from_dict(cls, data: dict):
        return cls(
            atom1=int((data or {}).get("atom1", 0) or 0),
            atom2=int((data or {}).get("atom2", 0) or 0),
            source_kind=str((data or {}).get("source_kind") or "parsed"),
            order=max(1, int((data or {}).get("order", 1) or 1)),
        )


class SvgMoleculeOverlay:
    """Editable 2D molecule overlay parsed from text-based chemistry formats."""

    # Display/style fields only (no geometry/identity) - the single source
    # of truth for "what counts as style" used by both the defaults-for-new-
    # molecules mechanism and the save/load style preset feature, so the two
    # can't silently drift apart from the Style dialog's actual fields.
    STYLE_FIELDS = (
        "atom_color_mode", "flat_atom_color",
        "label_font_scale", "high_contrast_labels",
        "show_bond_order", "show_bond_length_labels",
        "bond_color_mode", "bond_colormap", "bond_qualitative_cmap",
        "show_legend",
    )

    def get_style_dict(self) -> dict:
        return {name: getattr(self, name) for name in self.STYLE_FIELDS}

    def apply_style_dict(self, style: dict) -> None:
        for name in self.STYLE_FIELDS:
            if name in (style or {}):
                setattr(self, name, style[name])

    def __init__(self, filepath=None):
        self.filepath = str(filepath) if filepath else None
        self.name = Path(str(filepath)).stem if filepath else "2D Structure"
        self.atoms: list[SvgAtom] = []
        self.bonds: list[SvgBond] = []
        self.reference_bond_length_angstrom = 1.54
        self.source_median_bond_length = 1.0
        self.edit_mode = False
        self.source_format = "svg"
        # Off by default: a label on every single bond is the dominant
        # source of visual clutter on anything bigger than a handful of
        # atoms. Still user-toggleable via the Style dialog.
        self.show_bond_length_labels = False
        # Bond coloring: "uniform" (default), "length_continuous" (a
        # sequential/diverging colormap over length deviation from
        # reference), "length_binned" (discrete short/normal/long
        # categories via a qualitative colormap), or "bond_order" (color by
        # single/double/triple, qualitative colormap). Colormap names are
        # real matplotlib colormaps, user-selectable via the Style dialog.
        self.bond_color_mode = "uniform"
        self.bond_colormap = "coolwarm"
        self.bond_qualitative_cmap = "Set2"
        # User-controllable style, exposed via the Style dialog rather than
        # hardcoded, since font/color needs vary by figure use (screen
        # review vs. publication) and by user (colorblind-safe palettes,
        # larger text for readability).
        self.atom_color_mode = "cpk"  # one of available_atom_palettes(), or "flat"
        self.flat_atom_color = "#f7fafc"
        self.label_font_scale = 1.0
        self.high_contrast_labels = False
        # Off by default: bond order (double/triple bonds) reflects the
        # *source file's* assumed structure (often a gas-phase reference),
        # which is not necessarily what a molecule adsorbed on a surface
        # actually looks like - bonding and aromaticity can be altered by
        # the surface. Rendering it as fact by default would overstate what
        # is actually known/measured, so this is opt-in.
        self.show_bond_order = False
        # Off by default (no clutter until asked for); draws a small key
        # explaining what the current bond-coloring colors/categories mean,
        # useful for readers of a figure who don't know this app's colors.
        self.show_legend = False
        if filepath:
            loaded = self.from_file(filepath)
            self.__dict__.update(loaded.__dict__)

    def to_dict(self) -> dict:
        return {
            "type": "svg_molecule_overlay",
            "filepath": self.filepath,
            "name": self.name,
            "atoms": [atom.to_dict() for atom in self.atoms],
            "bonds": [bond.to_dict() for bond in self.bonds],
            "reference_bond_length_angstrom": float(self.reference_bond_length_angstrom),
            "source_median_bond_length": float(self.source_median_bond_length),
            "edit_mode": bool(self.edit_mode),
            "source_format": str(self.source_format or "svg"),
            "show_bond_length_labels": bool(self.show_bond_length_labels),
            "bond_color_mode": str(getattr(self, "bond_color_mode", "uniform") or "uniform"),
            "bond_colormap": str(getattr(self, "bond_colormap", "coolwarm") or "coolwarm"),
            "bond_qualitative_cmap": str(getattr(self, "bond_qualitative_cmap", "Set2") or "Set2"),
            "atom_color_mode": str(getattr(self, "atom_color_mode", "cpk") or "cpk"),
            "flat_atom_color": str(getattr(self, "flat_atom_color", "#f7fafc") or "#f7fafc"),
            "label_font_scale": float(getattr(self, "label_font_scale", 1.0) or 1.0),
            "high_contrast_labels": bool(getattr(self, "high_contrast_labels", False)),
            "show_bond_order": bool(getattr(self, "show_bond_order", False)),
            "show_legend": bool(getattr(self, "show_legend", False)),
        }

    @classmethod
    def from_dict(cls, data: dict):
        obj = cls()
        obj.filepath = (data or {}).get("filepath")
        obj.name = str((data or {}).get("name") or (Path(str(obj.filepath)).stem if obj.filepath else "2D Structure"))
        obj.atoms = [SvgAtom.from_dict(entry) for entry in (data or {}).get("atoms", []) or []]
        obj.bonds = [SvgBond.from_dict(entry) for entry in (data or {}).get("bonds", []) or []]
        obj.reference_bond_length_angstrom = float((data or {}).get("reference_bond_length_angstrom", 1.54) or 1.54)
        obj.source_median_bond_length = float((data or {}).get("source_median_bond_length", 1.0) or 1.0)
        obj.edit_mode = bool((data or {}).get("edit_mode", False))
        obj.source_format = str((data or {}).get("source_format") or "svg")
        obj.show_bond_length_labels = bool((data or {}).get("show_bond_length_labels", True))
        if "bond_color_mode" in (data or {}):
            obj.bond_color_mode = str((data or {}).get("bond_color_mode") or "uniform")
        else:
            # Migrate pre-existing saved sessions that only had the old
            # boolean toggle.
            obj.bond_color_mode = "length_continuous" if bool((data or {}).get("color_bonds_by_length", False)) else "uniform"
        obj.bond_colormap = str((data or {}).get("bond_colormap") or "coolwarm")
        obj.bond_qualitative_cmap = str((data or {}).get("bond_qualitative_cmap") or "Set2")
        obj.atom_color_mode = str((data or {}).get("atom_color_mode") or "cpk")
        obj.flat_atom_color = str((data or {}).get("flat_atom_color") or "#f7fafc")
        obj.label_font_scale = float((data or {}).get("label_font_scale", 1.0) or 1.0)
        obj.high_contrast_labels = bool((data or {}).get("high_contrast_labels", False))
        obj.show_bond_order = bool((data or {}).get("show_bond_order", False))
        obj.show_legend = bool((data or {}).get("show_legend", False))
        return obj

    @classmethod
    def from_svg_file(cls, filepath, reference_bond_length_angstrom: float = 1.54):
        path = Path(filepath)
        tree = ET.parse(path)
        root = tree.getroot()
        explicit_atoms: list[tuple[str, np.ndarray, bool, str]] = []
        bond_segments: list[tuple[np.ndarray, np.ndarray, str]] = []

        def visit(node, parent_transform):
            transform = parent_transform @ parse_svg_transform(node.attrib.get("transform"))
            tag = _strip_ns(node.tag).lower()

            if tag == "text":
                values_x = _parse_float_list(node.attrib.get("x"))
                values_y = _parse_float_list(node.attrib.get("y"))
                if values_x and values_y:
                    point = _apply_transform((values_x[0], values_y[0]), transform)
                    label = "".join(node.itertext()).strip()
                    element = _element_from_text(label, default="C")
                    if label:
                        explicit_atoms.append((element, point, True, "text"))

            elif tag in ("circle", "ellipse"):
                cx = _parse_float(node.attrib.get("cx"))
                cy = _parse_float(node.attrib.get("cy"))
                if cx is not None and cy is not None:
                    point = _apply_transform((cx, cy), transform)
                    raw_label = node.attrib.get("data-element") or node.attrib.get("element") or node.attrib.get("data-atom")
                    explicit = raw_label is not None
                    element = _element_from_text(raw_label, default="C")
                    explicit_atoms.append((element, point, explicit, tag))

            elif tag == "line":
                x1 = _parse_float(node.attrib.get("x1"))
                y1 = _parse_float(node.attrib.get("y1"))
                x2 = _parse_float(node.attrib.get("x2"))
                y2 = _parse_float(node.attrib.get("y2"))
                if None not in (x1, y1, x2, y2):
                    p0 = _apply_transform((x1, y1), transform)
                    p1 = _apply_transform((x2, y2), transform)
                    bond_segments.append((p0, p1, "line"))

            elif tag == "rect":
                x = _parse_float(node.attrib.get("x"), 0.0)
                y = _parse_float(node.attrib.get("y"), 0.0)
                width = _parse_float(node.attrib.get("width"))
                height = _parse_float(node.attrib.get("height"))
                if width is not None and height is not None:
                    points = [
                        _apply_transform((x, y), transform),
                        _apply_transform((x + width, y), transform),
                        _apply_transform((x + width, y + height), transform),
                        _apply_transform((x, y + height), transform),
                    ]
                    segment = _segment_from_points(points)
                    if segment is not None:
                        bond_segments.append((segment[0], segment[1], "rect"))

            elif tag in ("polygon", "polyline"):
                raw_points = _parse_float_list(node.attrib.get("points"))
                if len(raw_points) >= 4:
                    points = []
                    for idx in range(0, len(raw_points) - 1, 2):
                        points.append(_apply_transform((raw_points[idx], raw_points[idx + 1]), transform))
                    segment = _segment_from_points(points)
                    if segment is not None:
                        bond_segments.append((segment[0], segment[1], tag))

            elif tag == "path":
                points = [_apply_transform(point, transform) for point in _path_points(node.attrib.get("d"))]
                segment = _segment_from_points(points)
                if segment is not None:
                    bond_segments.append((segment[0], segment[1], "path"))

            for child in list(node):
                visit(child, transform)

        visit(root, _identity_transform())
        all_points = []
        for _el, point, _label_visible, _source in explicit_atoms:
            all_points.append(np.asarray(point, dtype=float))
        for p0, p1, _source in bond_segments:
            all_points.append(np.asarray(p0, dtype=float))
            all_points.append(np.asarray(p1, dtype=float))
        if not all_points:
            raise ValueError(f"No atom or bond geometry found in {path.name}")

        max_y = float(max(point[1] for point in all_points))

        def flip_y(point):
            pt = np.asarray(point, dtype=float)
            return np.array([float(pt[0]), max_y - float(pt[1])], dtype=float)

        explicit_atoms = [(el, flip_y(point), label_visible, source) for el, point, label_visible, source in explicit_atoms]
        bond_segments = [(flip_y(p0), flip_y(p1), source) for p0, p1, source in bond_segments]

        segment_lengths = [
            float(np.linalg.norm(np.asarray(p1, dtype=float) - np.asarray(p0, dtype=float)))
            for p0, p1, _source in bond_segments
            if float(np.linalg.norm(np.asarray(p1, dtype=float) - np.asarray(p0, dtype=float))) > 1e-6
        ]
        nearest_scale = 1.0
        if segment_lengths:
            nearest_scale = float(np.median(np.asarray(segment_lengths, dtype=float)))
        elif len(explicit_atoms) >= 2:
            explicit_points = np.asarray([point for _el, point, _lbl, _src in explicit_atoms], dtype=float)
            dists = []
            for idx in range(len(explicit_points)):
                delta = explicit_points - explicit_points[idx]
                norms = np.hypot(delta[:, 0], delta[:, 1])
                norms = norms[norms > 1e-6]
                if norms.size:
                    dists.append(float(np.min(norms)))
            if dists:
                nearest_scale = float(np.median(np.asarray(dists, dtype=float)))
        if nearest_scale <= 1e-12:
            nearest_scale = 1.0

        endpoint_threshold = max(0.35 * nearest_scale, 1.5)
        atoms: list[SvgAtom] = []
        for element, point, label_visible, source in explicit_atoms:
            atoms.append(
                SvgAtom(
                    element=str(element or "C"),
                    x=float(point[0]),
                    y=float(point[1]),
                    label_visible=bool(label_visible),
                    source_kind=str(source or "text"),
                )
            )

        def snap_or_create(point, default_element="C") -> int:
            best_idx = None
            best_dist = None
            point = np.asarray(point, dtype=float)
            for idx, atom in enumerate(atoms):
                dist = float(math.hypot(float(atom.x) - point[0], float(atom.y) - point[1]))
                if best_dist is None or dist < best_dist:
                    best_dist = dist
                    best_idx = idx
            if best_idx is not None and best_dist is not None and best_dist <= endpoint_threshold:
                return int(best_idx)
            atoms.append(
                SvgAtom(
                    element=str(default_element or "C"),
                    x=float(point[0]),
                    y=float(point[1]),
                    label_visible=False,
                    source_kind="inferred",
                )
            )
            return len(atoms) - 1

        bonds: list[SvgBond] = []
        seen_bonds = set()
        for p0, p1, source in bond_segments:
            idx0 = snap_or_create(p0, default_element="C")
            idx1 = snap_or_create(p1, default_element="C")
            if idx0 == idx1:
                continue
            bond_key = tuple(sorted((int(idx0), int(idx1))))
            if bond_key in seen_bonds:
                continue
            seen_bonds.add(bond_key)
            bonds.append(SvgBond(atom1=bond_key[0], atom2=bond_key[1], source_kind=str(source or "parsed")))

        atoms, bonds = _prune_svg_auxiliary_bond_components(atoms, bonds, reference_scale=nearest_scale)

        if not bonds and len(atoms) >= 2:
            positions = np.asarray([[atom.x, atom.y] for atom in atoms], dtype=float)
            nearest = []
            for idx in range(len(positions)):
                delta = positions - positions[idx]
                norms = np.hypot(delta[:, 0], delta[:, 1])
                norms = norms[norms > 1e-6]
                if norms.size:
                    nearest.append(float(np.min(norms)))
            inferred_threshold = 1.8 * (float(np.median(np.asarray(nearest, dtype=float))) if nearest else nearest_scale)
            for i in range(len(atoms)):
                for j in range(i + 1, len(atoms)):
                    dist = float(math.hypot(atoms[i].x - atoms[j].x, atoms[i].y - atoms[j].y))
                    if 1e-6 < dist <= inferred_threshold:
                        bonds.append(SvgBond(atom1=i, atom2=j, source_kind="proximity"))

        if not atoms:
            raise ValueError(f"No atoms could be inferred from {path.name}")

        positions = np.asarray([[atom.x, atom.y] for atom in atoms], dtype=float)
        center = np.mean(positions, axis=0)
        if segment_lengths:
            median_svg_bond = float(np.median(np.asarray(segment_lengths, dtype=float)))
        else:
            median_svg_bond = nearest_scale
        if median_svg_bond <= 1e-12:
            median_svg_bond = 1.0
        scale = float(reference_bond_length_angstrom) / float(median_svg_bond)
        for atom in atoms:
            atom.x = float((atom.x - center[0]) * scale)
            atom.y = float((atom.y - center[1]) * scale)

        overlay = cls()
        overlay.filepath = str(path)
        overlay.name = path.stem
        overlay.atoms = atoms
        overlay.bonds = bonds
        overlay.reference_bond_length_angstrom = float(reference_bond_length_angstrom)
        overlay.source_median_bond_length = float(median_svg_bond)
        overlay.edit_mode = False
        overlay.source_format = "svg"
        return overlay

    @classmethod
    def from_file(cls, filepath, reference_bond_length_angstrom: float = 1.54):
        path = Path(filepath)
        suffix = path.suffix.lower()
        if suffix == ".svg":
            return cls.from_svg_file(path, reference_bond_length_angstrom=reference_bond_length_angstrom)
        if suffix in {".mol", ".sdf"}:
            return cls.from_mol_like_file(path, reference_bond_length_angstrom=reference_bond_length_angstrom)
        if suffix == ".xyz":
            return cls.from_xyz_file(path, reference_bond_length_angstrom=reference_bond_length_angstrom)
        if suffix == ".cdxml":
            return cls.from_cdxml_file(path, reference_bond_length_angstrom=reference_bond_length_angstrom)
        if suffix == ".cml":
            return cls.from_cml_file(path, reference_bond_length_angstrom=reference_bond_length_angstrom)
        raise ValueError(
            f"Unsupported 2D structure format: {path.suffix or path.name}. "
            f"Supported: {', '.join(sorted(_SUPPORTED_2D_STRUCTURE_SUFFIXES))}"
        )

    @classmethod
    def _from_atom_bond_lists(
        cls,
        filepath,
        atoms_raw,
        bonds_raw,
        *,
        source_format: str,
        reference_bond_length_angstrom: float = 1.54,
        infer_bonds_if_missing: bool = True,
    ):
        atoms = []
        for entry in atoms_raw or []:
            try:
                element, x_val, y_val = entry[:3]
            except Exception:
                continue
            label_visible = bool(entry[3]) if len(entry) > 3 else False
            source_kind = str(entry[4]) if len(entry) > 4 else "parsed"
            atoms.append(
                SvgAtom(
                    element=_element_from_text(element, default="C"),
                    x=float(x_val),
                    y=float(y_val),
                    label_visible=label_visible,
                    source_kind=source_kind,
                )
            )
        if not atoms:
            raise ValueError(f"No atom coordinates found in {Path(filepath).name}")

        bond_pairs = []
        seen = set()
        for entry in bonds_raw or []:
            try:
                i_val, j_val = int(entry[0]), int(entry[1])
            except Exception:
                continue
            if i_val == j_val:
                continue
            if not (0 <= i_val < len(atoms) and 0 <= j_val < len(atoms)):
                continue
            key = tuple(sorted((i_val, j_val)))
            if key in seen:
                continue
            seen.add(key)
            source_kind = str(entry[2]) if len(entry) > 2 else "parsed"
            order = _bond_order_from_text(entry[3], default=1) if len(entry) > 3 else 1
            bond_pairs.append(SvgBond(atom1=key[0], atom2=key[1], source_kind=source_kind, order=order))

        if not bond_pairs and infer_bonds_if_missing and len(atoms) >= 2:
            positions = np.asarray([[atom.x, atom.y] for atom in atoms], dtype=float)
            nearest = []
            for idx in range(len(positions)):
                delta = positions - positions[idx]
                norms = np.hypot(delta[:, 0], delta[:, 1])
                norms = norms[norms > 1e-6]
                if norms.size:
                    nearest.append(float(np.min(norms)))
            inferred_threshold = 1.8 * (float(np.median(np.asarray(nearest, dtype=float))) if nearest else 1.0)
            for i_val in range(len(atoms)):
                for j_val in range(i_val + 1, len(atoms)):
                    dist = float(math.hypot(atoms[i_val].x - atoms[j_val].x, atoms[i_val].y - atoms[j_val].y))
                    if 1e-6 < dist <= inferred_threshold:
                        bond_pairs.append(SvgBond(atom1=i_val, atom2=j_val, source_kind="proximity", order=1))

        positions = np.asarray([[atom.x, atom.y] for atom in atoms], dtype=float)
        center = np.mean(positions, axis=0)
        bond_lengths = []
        for bond in bond_pairs:
            atom1 = atoms[int(bond.atom1)]
            atom2 = atoms[int(bond.atom2)]
            bond_lengths.append(float(math.hypot(atom2.x - atom1.x, atom2.y - atom1.y)))
        if bond_lengths:
            median_source_bond = float(np.median(np.asarray(bond_lengths, dtype=float)))
        else:
            nearest = []
            for idx in range(len(positions)):
                delta = positions - positions[idx]
                norms = np.hypot(delta[:, 0], delta[:, 1])
                norms = norms[norms > 1e-6]
                if norms.size:
                    nearest.append(float(np.min(norms)))
            median_source_bond = float(np.median(np.asarray(nearest, dtype=float))) if nearest else 1.0
        if median_source_bond <= 1e-12:
            median_source_bond = 1.0
        scale = float(reference_bond_length_angstrom) / median_source_bond
        for atom in atoms:
            atom.x = float((atom.x - center[0]) * scale)
            atom.y = float((atom.y - center[1]) * scale)

        overlay = cls()
        overlay.filepath = str(filepath)
        overlay.name = Path(str(filepath)).stem
        overlay.atoms = atoms
        overlay.bonds = bond_pairs
        overlay.reference_bond_length_angstrom = float(reference_bond_length_angstrom)
        overlay.source_median_bond_length = float(median_source_bond)
        overlay.edit_mode = False
        overlay.source_format = str(source_format or "unknown")
        return overlay

    @classmethod
    def from_xyz_file(cls, filepath, reference_bond_length_angstrom: float = 1.54):
        path = Path(filepath)
        lines = path.read_text(encoding="utf-8", errors="ignore").splitlines()
        start = 2 if len(lines) >= 2 else 0
        atoms_raw = []
        for line in lines[start:]:
            parts = line.split()
            if len(parts) < 3:
                continue
            element = _element_from_text(parts[0], default="C")
            try:
                x_val = float(parts[1])
                y_val = float(parts[2])
            except Exception:
                continue
            atoms_raw.append((element, x_val, y_val, element != "C", "xyz"))
        return cls._from_atom_bond_lists(
            path,
            atoms_raw,
            [],
            source_format="xyz",
            reference_bond_length_angstrom=reference_bond_length_angstrom,
            infer_bonds_if_missing=True,
        )

    @classmethod
    def from_mol_like_file(cls, filepath, reference_bond_length_angstrom: float = 1.54):
        path = Path(filepath)
        text = path.read_text(encoding="utf-8", errors="ignore")
        for block in text.split("$$$$"):
            block = block.strip("\r\n")
            if not block:
                continue
            if "V3000" in block:
                atoms_raw, bonds_raw = cls._parse_mol_v3000_block(block)
            else:
                atoms_raw, bonds_raw = cls._parse_mol_v2000_block(block)
            if atoms_raw:
                return cls._from_atom_bond_lists(
                    path,
                    atoms_raw,
                    bonds_raw,
                    source_format=path.suffix.lower().lstrip("."),
                    reference_bond_length_angstrom=reference_bond_length_angstrom,
                    infer_bonds_if_missing=not bool(bonds_raw),
                )
        raise ValueError(f"No usable molecule block found in {path.name}")

    @staticmethod
    def _parse_mol_v2000_block(block: str):
        lines = block.splitlines()
        if len(lines) < 2:
            return [], []
        counts_idx = None
        atom_count = 0
        bond_count = 0
        for idx, line in enumerate(lines[:8]):
            match = _MOL_COUNTS_RE.match(line)
            if not match:
                continue
            counts_idx = idx
            atom_count = int(match.group(1))
            bond_count = int(match.group(2))
            break
        if counts_idx is None:
            return [], []
        atoms_raw = []
        bonds_raw = []
        atom_start = counts_idx + 1
        bond_start = atom_start + atom_count
        atom_lines = lines[atom_start:bond_start]
        bond_lines = lines[bond_start:bond_start + bond_count]
        for line in atom_lines:
            if len(line) < 34:
                continue
            try:
                x_val = float(line[0:10].strip())
                y_val = float(line[10:20].strip())
            except Exception:
                continue
            element = _element_from_text(line[31:34].strip(), default="C")
            atoms_raw.append((element, x_val, y_val, element != "C", "mol"))
        for line in bond_lines:
            try:
                i_val = int(line[0:3].strip()) - 1
                j_val = int(line[3:6].strip()) - 1
            except Exception:
                continue
            order = _bond_order_from_text(line[6:9].strip(), default=1)
            bonds_raw.append((i_val, j_val, "mol", order))
        return atoms_raw, bonds_raw

    @staticmethod
    def _parse_mol_v3000_block(block: str):
        atoms_raw = []
        bonds_raw = []
        in_atom = False
        in_bond = False
        for raw_line in block.splitlines():
            line = raw_line.strip()
            if "BEGIN ATOM" in line:
                in_atom = True
                in_bond = False
                continue
            if "END ATOM" in line:
                in_atom = False
                continue
            if "BEGIN BOND" in line:
                in_atom = False
                in_bond = True
                continue
            if "END BOND" in line:
                in_bond = False
                continue
            if in_atom:
                tokens = line.split()
                if len(tokens) < 6:
                    continue
                try:
                    x_val = float(tokens[4])
                    y_val = float(tokens[5])
                except Exception:
                    continue
                element = _element_from_text(tokens[3], default="C")
                atoms_raw.append((element, x_val, y_val, element != "C", "mol"))
            elif in_bond:
                tokens = line.split()
                if len(tokens) < 6:
                    continue
                try:
                    i_val = int(tokens[4]) - 1
                    j_val = int(tokens[5]) - 1
                except Exception:
                    continue
                order = _bond_order_from_text(tokens[3], default=1)
                bonds_raw.append((i_val, j_val, "mol", order))
        return atoms_raw, bonds_raw

    @classmethod
    def from_cml_file(cls, filepath, reference_bond_length_angstrom: float = 1.54):
        path = Path(filepath)
        root = ET.parse(path).getroot()
        atom_map = {}
        atoms_raw = []
        for node in root.iter():
            if _strip_ns(node.tag).lower() != "atom":
                continue
            atom_id = node.attrib.get("id")
            x_val = _parse_float(node.attrib.get("x2"))
            y_val = _parse_float(node.attrib.get("y2"))
            if x_val is None or y_val is None:
                x_val = _parse_float(node.attrib.get("x3"))
                y_val = _parse_float(node.attrib.get("y3"))
            if x_val is None or y_val is None:
                continue
            element = _element_from_text(node.attrib.get("elementType"), default="C")
            atom_map[str(atom_id or len(atoms_raw))] = len(atoms_raw)
            atoms_raw.append((element, x_val, y_val, element != "C", "cml"))
        bonds_raw = []
        for node in root.iter():
            if _strip_ns(node.tag).lower() != "bond":
                continue
            refs = str(node.attrib.get("atomRefs2") or "").split()
            if len(refs) != 2:
                continue
            if refs[0] not in atom_map or refs[1] not in atom_map:
                continue
            order = _bond_order_from_text(node.attrib.get("order"), default=1)
            bonds_raw.append((atom_map[refs[0]], atom_map[refs[1]], "cml", order))
        return cls._from_atom_bond_lists(
            path,
            atoms_raw,
            bonds_raw,
            source_format="cml",
            reference_bond_length_angstrom=reference_bond_length_angstrom,
            infer_bonds_if_missing=not bool(bonds_raw),
        )

    @classmethod
    def from_cdxml_file(cls, filepath, reference_bond_length_angstrom: float = 1.54):
        path = Path(filepath)
        root = ET.parse(path).getroot()
        atom_map = {}
        atoms_raw = []
        for node in root.iter():
            tag = _strip_ns(node.tag).lower()
            if tag != "n":
                continue
            point_vals = _parse_float_list(node.attrib.get("p"))
            if len(point_vals) < 2:
                continue
            atom_id = str(node.attrib.get("id") or len(atoms_raw))
            element = None
            atomic_num = node.attrib.get("Element")
            if atomic_num is not None:
                element = _element_from_atomic_number(atomic_num, default="C")
            if element is None:
                element = _element_from_text("".join(node.itertext()).strip(), default="C")
            atom_map[atom_id] = len(atoms_raw)
            atoms_raw.append((element, point_vals[0], point_vals[1], element != "C", "cdxml"))
        bonds_raw = []
        for node in root.iter():
            if _strip_ns(node.tag).lower() != "b":
                continue
            begin = str(node.attrib.get("B") or "")
            end = str(node.attrib.get("E") or "")
            if begin not in atom_map or end not in atom_map:
                continue
            order = _bond_order_from_text(node.attrib.get("Order"), default=1)
            bonds_raw.append((atom_map[begin], atom_map[end], "cdxml", order))
        return cls._from_atom_bond_lists(
            path,
            atoms_raw,
            bonds_raw,
            source_format="cdxml",
            reference_bond_length_angstrom=reference_bond_length_angstrom,
            infer_bonds_if_missing=not bool(bonds_raw),
        )

    def atom_positions(self) -> np.ndarray:
        if not self.atoms:
            return np.zeros((0, 2), dtype=float)
        return np.asarray([[atom.x, atom.y] for atom in self.atoms], dtype=float)

    def centroid(self) -> np.ndarray:
        pts = self.atom_positions()
        if pts.size == 0:
            return np.zeros(2, dtype=float)
        return np.mean(pts, axis=0)

    def translate(self, dx: float, dy: float):
        for atom in self.atoms:
            atom.x = float(atom.x + dx)
            atom.y = float(atom.y + dy)

    def rotate(self, angle_deg: float, center=None):
        if not self.atoms:
            return
        if center is None:
            center = self.centroid()
        center = np.asarray(center, dtype=float)
        ang = math.radians(float(angle_deg))
        c = math.cos(ang)
        s = math.sin(ang)
        rot = np.array([[c, -s], [s, c]], dtype=float)
        for atom in self.atoms:
            vec = np.array([float(atom.x), float(atom.y)], dtype=float) - center
            mapped = center + (rot @ vec)
            atom.x = float(mapped[0])
            atom.y = float(mapped[1])

    def scale_about_centroid(self, factor: float):
        factor = float(factor)
        if factor <= 0.0 or not self.atoms:
            return
        center = self.centroid()
        for atom in self.atoms:
            atom.x = float(center[0] + ((atom.x - center[0]) * factor))
            atom.y = float(center[1] + ((atom.y - center[1]) * factor))

    def bond_pairs(self) -> list[tuple[int, int]]:
        pairs = []
        for bond in self.bonds:
            i = int(bond.atom1)
            j = int(bond.atom2)
            if i == j:
                continue
            if 0 <= i < len(self.atoms) and 0 <= j < len(self.atoms):
                pairs.append((i, j))
        return pairs
