"""Molecule manipulation (Qt-free) - one gesture model for every window.

The preview / pop-ups (``MultiPreviewCanvas``) and the figure windows
(``FigureWindow``) receive mouse events in different ways (matplotlib events
in data coordinates vs. Qt events on a rendered figure), but what a gesture
*does* to a molecule is defined once, here:

=====================================  ==========================================
gesture                                effect
=====================================  ==========================================
drag                                   ``translate`` - move in the image plane
Shift+drag                             ``rotate_z`` - spin about the model centre
Ctrl+Shift+drag / middle-button drag   ``rotate_3d`` - tumble, 0.5 deg/px
gizmo inner disc                       ``gizmo_rotate_xy`` - tumble, 0.45 deg/px
gizmo rim (r >= 0.58)                  ``gizmo_rotate_z`` - spin
X / Y / Z (Shift = reverse)            rotate 5 deg about that axis
=====================================  ==========================================

Screen conventions: pixel deltas and gizmo-local vectors are y **up**
(matplotlib's convention); a Qt host flips its y before calling in.
"""
from __future__ import annotations

import math

import numpy as np

TRANSLATE = "translate"
ROTATE_Z = "rotate_z"
ROTATE_3D = "rotate_3d"
GIZMO_ROTATE_XY = "gizmo_rotate_xy"
GIZMO_ROTATE_Z = "gizmo_rotate_z"

ROTATE_3D_DEG_PER_PX = 0.5
GIZMO_DEG_PER_PX = 0.45
KEY_STEP_DEG = 5.0
HIT_RADIUS_PX = 14.0
GIZMO_INNER_RADIUS = 0.58    # gizmo-local radius where the rim (spin) starts
GIZMO_OUTER_RADIUS = 1.02    # beyond this the gizmo is not hit
AXIS_COLORS = ("#ef4444", "#22c55e", "#3b82f6")   # X, Y, Z
KEY_AXES = {"x": 0, "y": 1, "z": 2}


# --- orientation ---------------------------------------------------------
def orientation_matrix(mol):
    """3x3 rotation of ``mol`` (Rz @ Ry @ Rx of its ``angles`` in degrees,
    rows negated for ``mirror_x`` / ``mirror_y``); identity for ``None``."""
    if mol is None:
        return np.eye(3, dtype=float)
    rads = np.radians(np.asarray(getattr(mol, "angles", [0.0, 0.0, 0.0]), dtype=float))
    cx, cy, cz = np.cos(rads)
    sx, sy, sz = np.sin(rads)
    rx = np.array([[1.0, 0.0, 0.0], [0.0, cx, -sx], [0.0, sx, cx]], dtype=float)
    ry = np.array([[cy, 0.0, sy], [0.0, 1.0, 0.0], [-sy, 0.0, cy]], dtype=float)
    rz = np.array([[cz, -sz, 0.0], [sz, cz, 0.0], [0.0, 0.0, 1.0]], dtype=float)
    mat = rz @ ry @ rx
    if bool(getattr(mol, "mirror_x", False)):
        mat[0, :] *= -1.0
    if bool(getattr(mol, "mirror_y", False)):
        mat[1, :] *= -1.0
    return mat


def axis_screen_vectors(mol):
    """``[(axis, vector3, screen_xy, colour), ...]`` for X/Y/Z: each axis'
    direction and its oblique 2-D projection (z leaks into x/y so the third
    axis stays visible)."""
    mat = orientation_matrix(mol)
    out = []
    for idx, color in enumerate(AXIS_COLORS):
        vec = np.asarray(mat[:, idx], dtype=float)
        screen = np.array([vec[0] + 0.38 * vec[2], vec[1] + 0.18 * vec[2]], dtype=float)
        out.append((idx, vec, screen, color))
    return out


def gizmo_vectors(mol, reach=0.74):
    """Gizmo arrows ready to paint: ``[(axis, depth, (x, y), colour)]`` in
    gizmo-local units (-1..1, y up), scaled so the longest arrow reaches
    ``reach`` and sorted back to front."""
    vecs = [(i, float(v[2]), sc, c) for i, v, sc, c in axis_screen_vectors(mol)]
    norms = [max(0.22, float(np.linalg.norm(sc))) for _i, _z, sc, _c in vecs]
    scale = reach / max(norms or [1.0])
    return [(i, z, (float(sc[0] * scale), float(sc[1] * scale)), c)
            for i, z, sc, c in sorted(vecs, key=lambda t: t[1])]


# --- hit testing ---------------------------------------------------------
def gizmo_hit_mode(local_x, local_y):
    """``"rotate_xy"`` (inner disc), ``"rotate_z"`` (rim) or ``None`` for a
    point in gizmo-local units (centre 0, radius 1)."""
    r = math.hypot(float(local_x), float(local_y))
    if r > GIZMO_OUTER_RADIUS:
        return None
    return "rotate_z" if r >= GIZMO_INNER_RADIUS else "rotate_xy"


def atoms_hit(atom_px, point_px, radii_px=None):
    """True when ``point_px`` lies within ``HIT_RADIUS_PX`` of any atom
    centre (``atom_px``: N x 2 screen pixels) - or inside an atom, when its
    on-screen radii are given (to-scale / space-filling styles)."""
    atom_px = np.asarray(atom_px, dtype=float).reshape(-1, 2)
    if not atom_px.size:
        return False
    dist = np.hypot(atom_px[:, 0] - float(point_px[0]), atom_px[:, 1] - float(point_px[1]))
    thresh = HIT_RADIUS_PX
    if radii_px is not None and len(radii_px) == len(dist):
        thresh = np.maximum(thresh, np.asarray(radii_px, dtype=float))
    return bool(np.any(dist <= thresh))


# --- gestures ------------------------------------------------------------
def drag_mode(*, middle=False, shift=False, ctrl=False):
    """Gesture for a press on a molecule (see the module table)."""
    if middle or (shift and ctrl):
        return ROTATE_3D
    if shift:
        return ROTATE_Z
    return TRANSLATE


def drag_mode_for_key(button, key):
    """``drag_mode`` from a matplotlib event's ``button`` / ``key`` string
    (``"shift"``, ``"ctrl+shift"``; ``"control"`` accepted too)."""
    k = str(key or "").lower()
    return drag_mode(middle=(button == 2), shift="shift" in k,
                     ctrl=("ctrl" in k or "control" in k))


def drag_update(mode, offset0, angles0, *, start=None, current=None,
                delta_px=(0.0, 0.0), start_local=None, current_local=None):
    """New ``(offset, angles)`` of a molecule dragged from its pose at press
    (``offset0``, ``angles0``).

    ``start`` / ``current`` - pointer in the molecule's data units (used by
    ``translate`` / ``rotate_z``; ``rotate_z`` spins about ``offset0``).
    ``delta_px`` - pointer travel in pixels, y up (``rotate_3d`` /
    ``gizmo_rotate_xy``). ``start_local`` / ``current_local`` - pointer as a
    vector from the gizmo centre, y up, any scale (``gizmo_rotate_z``).
    Returns copies; the inputs are untouched.
    """
    offset = np.array(offset0, dtype=float, copy=True)
    angles = np.array(angles0, dtype=float, copy=True)
    if mode == TRANSLATE:
        if start is not None and current is not None:
            offset = offset + np.array([current[0] - start[0], current[1] - start[1], 0.0])
    elif mode == ROTATE_Z:
        if start is not None and current is not None:
            v0 = np.array([start[0] - offset[0], start[1] - offset[1]], dtype=float)
            v1 = np.array([current[0] - offset[0], current[1] - offset[1]], dtype=float)
            if np.linalg.norm(v0) > 0.01 and np.linalg.norm(v1) > 0.01:
                angles[2] += np.degrees(np.arctan2(v1[1], v1[0]) - np.arctan2(v0[1], v0[0]))
    elif mode in (ROTATE_3D, GIZMO_ROTATE_XY):
        sens = ROTATE_3D_DEG_PER_PX if mode == ROTATE_3D else GIZMO_DEG_PER_PX
        angles[0] += float(delta_px[1]) * sens
        angles[1] += float(delta_px[0]) * sens
    elif mode == GIZMO_ROTATE_Z:
        if start_local is not None and current_local is not None:
            a0 = math.degrees(math.atan2(float(start_local[1]), float(start_local[0])))
            a1 = math.degrees(math.atan2(float(current_local[1]), float(current_local[0])))
            angles[2] += a1 - a0
    return offset, angles


def key_rotation(axis, reverse=False):
    """Angle increment for an X/Y/Z key (``axis`` 0/1/2 or "x"/"y"/"z")."""
    if isinstance(axis, str):
        axis = KEY_AXES[axis.lower()]
    rot = np.zeros(3, dtype=float)
    rot[int(axis)] = -KEY_STEP_DEG if reverse else KEY_STEP_DEG
    return rot


def apply_transform(mol, *, rot=None, move=None, mirror=None):
    """Mutate ``mol``: add ``rot`` (degrees) to its angles, ``move`` (x, y)
    to its offset, toggle ``mirror`` ("x" / "y")."""
    if rot is not None:
        mol.angles = np.asarray(mol.angles, dtype=float) + np.asarray(rot, dtype=float)
    if move is not None:
        mol.offset = np.asarray(mol.offset, dtype=float) + np.array([move[0], move[1], 0.0])
    if mirror == "x":
        mol.mirror_x = not bool(mol.mirror_x)
    elif mirror == "y":
        mol.mirror_y = not bool(mol.mirror_y)
    return mol
