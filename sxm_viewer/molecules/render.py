"""Molecule renderer (Qt-free) - one implementation for every window.

Moved from ``MultiPreviewCanvas._draw_molecules`` (now a thin wrapper); the
only changes: canvas attributes became the ``settings`` dict, and
screen-sized marks take a ``scale`` factor (formerly patched on afterwards
by the figure windows).
"""
from __future__ import annotations

import math

import matplotlib
import matplotlib.patheffects as PathEffects
import numpy as np
from matplotlib.collections import EllipseCollection, LineCollection

from .model import (
    atom_marker_area_factor,
    effective_bond_color_mode,
    get_atom_radius,
    normalize_molecule_render_style,
    resolve_atom_color,
)

DEFAULT_SETTINGS = {
    "palette": "avogadro",
    "bond_color_mode": "default",
    "default_bond_color": (0.9, 0.9, 0.9),
    "show_shadows": False,
    "show_hydrogens": True,
}


def artists_of(entries):
    """Every matplotlib artist in ``draw_molecules`` entries."""
    out = []
    for e in entries or []:
        for k in ("scatter", "atom_underlay", "shadow", "line_underlay", "lines"):
            a = e.get(k)
            if a is not None:
                out.append(a)
    return out


def draw_molecules(ax, molecules, settings=None, *, scale=1.0):
    """Draw ``molecules`` onto ``ax`` (data coords = the molecules' own
    units, nm in the app); return one entry per molecule ``{mol, ax,
    scatter, atom_underlay, atom_order, shadow, line_underlay, lines}``.

    THE molecule renderer of the app: preview, pop-ups, figure windows,
    exports and the publication canvas all draw through it.
    ``settings``: palette, bond_color_mode, default_bond_color,
    show_shadows, show_hydrogens (``DEFAULT_SETTINGS``). ``scale``
    multiplies screen-sized marks (marker areas by scale**2, line widths
    by scale): 1.0 in the preview, smaller for a small printed panel.
    To-scale (space-filling) atoms are data units and never scaled.
    """
    s = dict(DEFAULT_SETTINGS, **(settings or {}))
    entries = []
    if not molecules:
        return entries

    for mol in molecules:
        coords = mol.get_transformed_coordinates()
        if len(coords) == 0:
            continue
        hide_h = not s["show_hydrogens"]

        # Z-range for depth cueing
        z_vals = coords[:, 2]
        z_min = z_vals.min()
        z_range = z_vals.max() - z_vals.min()
        is_flat = z_range < 1e-6
        if is_flat:
            z_range = 1.0
        # Depth cue: far atoms/bonds fade and shrink. A planar molecule
        # has no depth to cue, so it gets the near (full) opacity rather
        # than every atom being treated as "farthest" - that made flat
        # overlays look washed out with no way to make them opaque.
        depth_cue = bool(getattr(mol, "depth_cue", True)) and not is_flat
        atom_opacity = float(np.clip(float(getattr(mol, "atom_opacity", 1.0) or 0.0), 0.0, 1.0))
        bond_opacity = float(np.clip(float(getattr(mol, "bond_opacity", 1.0) or 0.0), 0.0, 1.0))
        space_filling = bool(getattr(mol, "space_filling", False))

        lc = None
        lc_underlay = None
        sc = None
        shadow_sc = None
        atom_underlay_sc = None
        atom_order = None
        atom_style = normalize_molecule_render_style(mol.render_style)
        bond_style = (mol.bond_style or "default").lower()
        style_profiles = {
            "shaded": {
                "size_base": 80, "size_scale": 140, "shadow_alpha": 0.25,
                "show_shadow": True, "atom_alpha": 1.0, "highlight": 0.25,
                "edgecolor": "black", "edgewidth": 0.6, "bond_scale": 1.0,
                "bond_alpha_min": 0.4, "bond_alpha_span": 0.6, "split_bonds": False,
                "force_bond_by_atoms": False, "outline": False,
            },
            "flat": {
                "size_base": 76, "size_scale": 126, "shadow_alpha": 0.0,
                "show_shadow": False, "atom_alpha": 0.98, "highlight": 0.0,
                "edgecolor": "black", "edgewidth": 0.5, "bond_scale": 1.0,
                "bond_alpha_min": 0.55, "bond_alpha_span": 0.4, "split_bonds": False,
                "force_bond_by_atoms": False, "outline": False,
            },
            "ballstick": {
                "size_base": 34, "size_scale": 62, "shadow_alpha": 0.14,
                "show_shadow": True, "atom_alpha": 1.0, "highlight": 0.18,
                "edgecolor": "black", "edgewidth": 0.75, "bond_scale": 1.35,
                "bond_alpha_min": 0.55, "bond_alpha_span": 0.4, "split_bonds": True,
                "force_bond_by_atoms": True, "outline": False,
            },
            "cpk": {
                "size_base": 150, "size_scale": 240, "shadow_alpha": 0.28,
                "show_shadow": True, "atom_alpha": 0.96, "highlight": 0.3,
                "edgecolor": "black", "edgewidth": 0.75, "bond_scale": 0.75,
                "bond_alpha_min": 0.3, "bond_alpha_span": 0.35, "split_bonds": False,
                "force_bond_by_atoms": False, "outline": False,
            },
            "licorice": {
                "size_base": 48, "size_scale": 82, "shadow_alpha": 0.2,
                "show_shadow": True, "atom_alpha": 0.98, "highlight": 0.12,
                "edgecolor": "black", "edgewidth": 0.55, "bond_scale": 1.5,
                "bond_alpha_min": 0.5, "bond_alpha_span": 0.45, "split_bonds": True,
                "force_bond_by_atoms": True, "outline": False,
            },
            "wire": {
                "size_base": 14, "size_scale": 24, "shadow_alpha": 0.0,
                "show_shadow": False, "atom_alpha": 0.82, "highlight": 0.0,
                "edgecolor": "black", "edgewidth": 0.3, "bond_scale": 0.65,
                "bond_alpha_min": 0.45, "bond_alpha_span": 0.35, "split_bonds": False,
                "force_bond_by_atoms": False, "outline": False,
            },
            "line": {
                "size_base": 8, "size_scale": 14, "shadow_alpha": 0.0,
                "show_shadow": False, "atom_alpha": 0.8, "highlight": 0.0,
                "edgecolor": "none", "edgewidth": 0.0, "bond_scale": 0.5,
                "bond_alpha_min": 0.55, "bond_alpha_span": 0.25, "split_bonds": True,
                "force_bond_by_atoms": True, "outline": False,
            },
            "sticks": {
                "size_base": 18, "size_scale": 35, "shadow_alpha": 0.12,
                "show_shadow": True, "atom_alpha": 0.96, "highlight": 0.08,
                "edgecolor": "black", "edgewidth": 0.35, "bond_scale": 1.6,
                "bond_alpha_min": 0.55, "bond_alpha_span": 0.4, "split_bonds": True,
                "force_bond_by_atoms": True, "outline": False,
            },
            "skeletal": {
                "size_base": 10, "size_scale": 20, "shadow_alpha": 0.0,
                "show_shadow": False, "atom_alpha": 0.88, "highlight": 0.0,
                "edgecolor": "black", "edgewidth": 0.25, "bond_scale": 0.85,
                "bond_alpha_min": 0.5, "bond_alpha_span": 0.35, "split_bonds": False,
                "force_bond_by_atoms": False, "outline": False,
            },
            "outline": {
                "size_base": 78, "size_scale": 128, "shadow_alpha": 0.0,
                "show_shadow": False, "atom_alpha": 0.12, "highlight": 0.0,
                "edgecolor": "white", "edgewidth": 1.0, "bond_scale": 1.1,
                "bond_alpha_min": 0.85, "bond_alpha_span": 0.05, "split_bonds": False,
                "force_bond_by_atoms": False, "outline": True,
            },
            "ghost": {
                "size_base": 82, "size_scale": 132, "shadow_alpha": 0.0,
                "show_shadow": False, "atom_alpha": 0.28, "highlight": 0.05,
                "edgecolor": "black", "edgewidth": 0.35, "bond_scale": 0.9,
                "bond_alpha_min": 0.22, "bond_alpha_span": 0.18, "split_bonds": False,
                "force_bond_by_atoms": False, "outline": False,
            },
        }
        profile = style_profiles.get(atom_style, style_profiles["shaded"])
        size_base = profile["size_base"] * scale * scale
        size_scale = profile["size_scale"] * scale * scale
        shadow_alpha = profile["shadow_alpha"]

        def _atom_base_rgba(idx):
            elem = mol.elements[idx] if idx < len(mol.elements) else ""
            return matplotlib.colors.to_rgba(resolve_atom_color(mol, elem, s["palette"]))

        # Draw Bonds
        if 'Bonds' in mol.display_mode and len(mol.bonds) > 0:
            lines = []
            underlay_lines = []
            colors = []
            underlay_colors = []
            linewidths = []
            underlay_widths = []
            lw_scale = profile["bond_scale"]
            if bond_style == "thick":
                lw_scale *= 1.6
            elif bond_style == "thin":
                lw_scale *= 0.7
            display_mode_lower = (mol.display_mode or "").lower()
            force_atom_bond_colors = profile["force_bond_by_atoms"] or ("bonds only" in display_mode_lower)
            bond_outline_rgba = (0.0, 0.0, 0.0, 0.95 * bond_opacity)
            for (i, j) in mol.bonds:
                if i >= len(coords) or j >= len(coords): continue
                ei = (mol.elements[i] or "").strip().upper() if i < len(mol.elements) else ""
                ej = (mol.elements[j] or "").strip().upper() if j < len(mol.elements) else ""
                if hide_h:
                    try:
                        if ei == 'H' or ej == 'H':
                            continue
                    except Exception:
                        pass
                p1 = coords[i]
                p2 = coords[j]
                z_mid = (p1[2] + p2[2]) * 0.5
                z_norm = (z_mid - z_min) / z_range if depth_cue else 0.0
                depth = z_norm if depth_cue else 1.0
                alpha = (profile["bond_alpha_min"] + profile["bond_alpha_span"] * depth) * bond_opacity
                lw = (1.0 + 2.0 * z_norm) * lw_scale
                bond_mode = effective_bond_color_mode(mol, s["bond_color_mode"])
                if force_atom_bond_colors and bond_mode == "default":
                    bond_mode = "by_atoms"
                rgba1 = _atom_base_rgba(i)
                rgba2 = _atom_base_rgba(j)
                if bond_mode == "single" and mol.bond_color_override:
                    br, bg, bb, _ = matplotlib.colors.to_rgba(mol.bond_color_override)
                    segment_color = (br, bg, bb, alpha)
                    lines.append([(p1[0], p1[1]), (p2[0], p2[1])])
                    colors.append(segment_color)
                    linewidths.append(lw)
                    if profile["outline"]:
                        underlay_lines.append([(p1[0], p1[1]), (p2[0], p2[1])])
                        underlay_colors.append(bond_outline_rgba)
                        underlay_widths.append(lw + 2.4)
                elif bond_mode == "by_atoms":
                    r1, g1, b1, _ = rgba1
                    r2, g2, b2, _ = rgba2
                    if profile["split_bonds"]:
                        mid = 0.5 * (p1 + p2)
                        lines.append([(p1[0], p1[1]), (mid[0], mid[1])])
                        colors.append((r1, g1, b1, alpha))
                        linewidths.append(lw)
                        lines.append([(mid[0], mid[1]), (p2[0], p2[1])])
                        colors.append((r2, g2, b2, alpha))
                        linewidths.append(lw)
                        if profile["outline"]:
                            underlay_lines.extend([
                                [(p1[0], p1[1]), (mid[0], mid[1])],
                                [(mid[0], mid[1]), (p2[0], p2[1])],
                            ])
                            underlay_colors.extend([bond_outline_rgba] * 2)
                            underlay_widths.extend([lw + 2.4, lw + 2.4])
                    else:
                        br, bg, bb = (0.5 * (r1 + r2), 0.5 * (g1 + g2), 0.5 * (b1 + b2))
                        lines.append([(p1[0], p1[1]), (p2[0], p2[1])])
                        colors.append((br, bg, bb, alpha))
                        linewidths.append(lw)
                        if profile["outline"]:
                            underlay_lines.append([(p1[0], p1[1]), (p2[0], p2[1])])
                            underlay_colors.append(bond_outline_rgba)
                            underlay_widths.append(lw + 2.4)
                else:
                    br, bg, bb = s["default_bond_color"]
                    lines.append([(p1[0], p1[1]), (p2[0], p2[1])])
                    colors.append((br, bg, bb, alpha))
                    linewidths.append(lw)
                    if profile["outline"]:
                        underlay_lines.append([(p1[0], p1[1]), (p2[0], p2[1])])
                        underlay_colors.append(bond_outline_rgba)
                        underlay_widths.append(lw + 2.4)

            if underlay_lines:
                lc_underlay = LineCollection(underlay_lines, colors=underlay_colors, linewidths=[w * scale for w in underlay_widths], zorder=28.7)
                ax.add_collection(lc_underlay)
                try:
                    lc_underlay.set_capstyle("round")
                    lc_underlay.set_joinstyle("round")
                except Exception:
                    pass
            lc = LineCollection(lines, colors=colors, linewidths=[w * scale for w in linewidths], zorder=29)
            ax.add_collection(lc)
            lc.set_pickradius(5)
            try:
                # Round ends in every style: bonds meeting at an atom (and
                # the two halves of a two-colour bond) blend into one
                # continuous stick, as in Avogadro / VESTA / PyMOL -
                # square ("butt") ends left notches and gaps at each joint.
                lc.set_capstyle("round")
                lc.set_joinstyle("round")
            except Exception:
                pass

        # Draw Atoms
        if 'Atoms' in mol.display_mode:
            # Sort atoms by Z for simple painter's algorithm
            order = np.argsort(z_vals)
            if hide_h:
                order = [idx for idx in order if str(mol.elements[idx]).strip().upper() != 'H']
            if len(order) == 0:
                continue
            atom_order = list(order)
            coords_sorted = coords[order]
            elements_sorted = [mol.elements[i] for i in order]

            x = coords_sorted[:, 0]
            y = coords_sorted[:, 1]
            z = coords_sorted[:, 2]

            z_norm = (z - z_min) / z_range if depth_cue else np.zeros(len(z))
            depth = z_norm if depth_cue else np.ones(len(z))
            # The radius model is honoured as chosen for every
            # representation (CPK used to silently swap covalent->vdW).
            rad_mode = getattr(mol, "radius_mode", "covalent")
            rad_scale = float(getattr(mol, "radius_scale", 1.0) or 1.0)
            if space_filling:
                # True size in data units: Å radius x molecule scale
                # (Å -> axis unit) x radius scale.
                mol_scale = abs(float(getattr(mol, "scale", 0.1) or 0.1))
                radii_data = np.array(
                    [get_atom_radius(e, rad_mode) for e in elements_sorted], dtype=float
                ) * mol_scale * rad_scale
                sizes = None
            else:
                # Scatter `s` is an area, so scale by r**2 to keep the
                # on-screen radius proportional to the radius model.
                radii_data = None
                area_factors = np.array(
                    [atom_marker_area_factor(e, rad_mode, rad_scale) for e in elements_sorted], dtype=float
                )
                sizes = (size_base + size_scale * z_norm) * area_factors

            rgba_colors = [_atom_base_rgba(i) for i in order]
            final_colors = []
            for i, (r, g, b, a) in enumerate(rgba_colors):
                depth_alpha = profile["atom_alpha"] * (0.45 + 0.55 * depth[i]) * atom_opacity
                highlight = profile["highlight"]
                r_h = min(1.0, r + (1 - r) * highlight)
                g_h = min(1.0, g + (1 - g) * highlight)
                b_h = min(1.0, b + (1 - b) * highlight)
                final_colors.append((r_h, g_h, b_h, depth_alpha))
            edge = profile["edgecolor"]
            if str(edge).lower() != "none":
                edge = matplotlib.colors.to_rgba(edge, atom_opacity)

            def _atom_layer(dx, dy, grow, facecolors, edgecolors, linewidths, zorder):
                if space_filling:
                    diam = 2.0 * radii_data * grow
                    coll = EllipseCollection(
                        diam, diam, np.zeros_like(diam),
                        units="xy",
                        offsets=np.c_[x + dx, y + dy],
                        offset_transform=ax.transData,
                        facecolors=facecolors,
                        edgecolors=edgecolors,
                        linewidths=linewidths,
                        zorder=zorder,
                    )
                    ax.add_collection(coll, autolim=False)
                    return coll
                return ax.scatter(
                    x + dx, y + dy,
                    s=sizes * grow * grow,
                    c=facecolors,
                    edgecolors=edgecolors,
                    linewidths=linewidths,
                    zorder=zorder,
                )

            if profile["outline"]:
                atom_underlay_sc = _atom_layer(
                    0.0, 0.0, math.sqrt(1.55),
                    [(0.0, 0.0, 0.0, 0.92 * atom_opacity)] * len(x), 'none', 0, 28.85,
                )
            if s["show_shadows"] and profile["show_shadow"]:
                shadow_sc = _atom_layer(
                    0.05, -0.05, math.sqrt(1.25),
                    [(0, 0, 0, shadow_alpha * atom_opacity)] * len(x), 'none', 0, 28,
                )
            sc = _atom_layer(0.0, 0.0, 1.0, final_colors, edge, profile["edgewidth"] * scale, 30)
            if atom_style == "outline":
                try:
                    sc.set_path_effects([PathEffects.Stroke(linewidth=2.2 * scale, foreground='black'), PathEffects.Normal()])
                except Exception:
                    pass

        entries.append({
            'mol': mol,
            'ax': ax,
            'scatter': sc,
            'atom_underlay': atom_underlay_sc,
            'atom_order': atom_order,
            'shadow': shadow_sc,
            'line_underlay': lc_underlay,
            'lines': lc
        })
    return entries


def panel_scale(ax, user=1.0, reference_in=3.5):
    """``scale`` for ``draw_molecules`` on a printed panel: the axes' width
    in inches against a ~3.5 in preview (clamped to 0.12..1), times the
    user's atom-size factor. Atom and bond marks are sized in points for the
    preview, so a 30 mm figure panel needs them proportionally smaller."""
    try:
        w_in = ax.get_position().width * ax.figure.get_size_inches()[0]
    except Exception:
        w_in = reference_in
    return float(min(1.0, max(0.12, w_in / reference_in)) * float(user or 1.0))
