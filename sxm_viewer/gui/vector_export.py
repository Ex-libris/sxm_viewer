"""Molecule-aware SVG post-processing for vector exports (Qt-free).

Matplotlib writes every artist as a flat ``<g id=...>`` in paint order, so a
molecule's bonds, atoms and labels end up interleaved with the rest of the
figure. At export time the preview canvas tags each molecule artist with a
gid from :func:`molecule_gid`; the helpers here then either

* restructure the SVG into Inkscape layers (``Image`` + one ``Molecule``
  layer per overlay with ``Bonds``/``Atoms``/``Labels`` sub-layers), or
* split it into standalone SVGs (base figure, and one per molecule role),
  which the PowerPoint bridge converts separately and re-nests as groups —
  PowerPoint's "Convert to Shape" flattens SVG ``<g>`` structure, so the
  grouping has to be rebuilt on the PowerPoint side.

All derived SVGs keep the root ``width``/``height``/``viewBox`` of the source,
so they overlay pixel-exactly when placed in the same box.
"""
from __future__ import annotations

import copy
import io
import re
import xml.etree.ElementTree as ET

SVG_NS = "http://www.w3.org/2000/svg"
XLINK_NS = "http://www.w3.org/1999/xlink"
INKSCAPE_NS = "http://www.inkscape.org/namespaces/inkscape"
SODIPODI_NS = "http://sodipodi.sourceforge.net/DTD/sodipodi-0.dtd"

for _prefix, _uri in (
    ("", SVG_NS),
    ("xlink", XLINK_NS),
    ("inkscape", INKSCAPE_NS),
    ("sodipodi", SODIPODI_NS),
    ("rdf", "http://www.w3.org/1999/02/22-rdf-syntax-ns#"),
    ("cc", "http://creativecommons.org/ns#"),
    ("dc", "http://purl.org/dc/elements/1.1/"),
):
    ET.register_namespace(_prefix, _uri)

MOLECULE_GID_PREFIX = "sxmmol"
# Paint order of roles inside a molecule group (bottom -> top).
MOLECULE_ROLES = ("bonds", "atoms", "labels")
ROLE_TITLES = {"bonds": "Bonds", "atoms": "Atoms", "labels": "Labels"}

_GID_RE = re.compile(rf"^{MOLECULE_GID_PREFIX}(\d+)--([a-z]+)--\d+$")
_PASSTHROUGH_TAGS = {f"{{{SVG_NS}}}{tag}" for tag in ("defs", "style", "metadata", "title", "desc")}


def molecule_gid(mol_idx: int, role: str, part_idx: int) -> str:
    """The gid assigned to one molecule artist for the duration of an export."""
    return f"{MOLECULE_GID_PREFIX}{int(mol_idx)}--{role}--{int(part_idx)}"


def _parse_gid(element) -> tuple[int, str] | None:
    match = _GID_RE.match(element.get("id") or "")
    if not match:
        return None
    return int(match.group(1)), match.group(2)


def _load(svg_bytes: bytes):
    return ET.ElementTree(ET.fromstring(svg_bytes))


def _dump(tree) -> bytes:
    out = io.BytesIO()
    tree.write(out, encoding="utf-8", xml_declaration=True)
    return out.getvalue()


def _parent_map(root):
    return {child: parent for parent in root.iter() for child in parent}


def _molecule_elements(root):
    """{mol_idx: {role: [element, ...]}} in document (paint) order."""
    found: dict[int, dict[str, list]] = {}
    for element in root.iter(f"{{{SVG_NS}}}g"):
        parsed = _parse_gid(element)
        if parsed is None:
            continue
        mol_idx, role = parsed
        found.setdefault(mol_idx, {}).setdefault(role, []).append(element)
    return found


def has_molecule_layers(svg_bytes: bytes) -> bool:
    return f'id="{MOLECULE_GID_PREFIX}'.encode() in (svg_bytes or b"")


def _ordered_roles(roles: dict) -> list[str]:
    known = [role for role in MOLECULE_ROLES if role in roles]
    return known + sorted(role for role in roles if role not in MOLECULE_ROLES)


def _keep_only(svg_bytes: bytes, keep_ids: set[str]) -> bytes:
    """A copy of the SVG holding only the groups whose id is in ``keep_ids``
    (plus defs/style, which clip paths and markers reference)."""
    tree = _load(svg_bytes)
    root = tree.getroot()

    def _prune(element) -> bool:
        """Return True when ``element`` should survive."""
        if element.tag in _PASSTHROUGH_TAGS:
            return True
        if (element.get("id") or "") in keep_ids:
            return True
        for child in list(element):
            if not _prune(child):
                element.remove(child)
        return len(element) > 0

    for child in list(root):
        if not _prune(child):
            root.remove(child)
    return _dump(tree)


def split_for_powerpoint(svg_bytes: bytes, molecule_names: dict[int, str] | None = None):
    """Split a tagged figure SVG into ``(base_svg, molecules)``.

    ``molecules`` is a list of ``(name, [(role_title, role_svg), ...])`` in
    paint order. Returns ``(svg_bytes, [])`` when nothing is tagged.
    """
    tree = _load(svg_bytes)
    root = tree.getroot()
    found = _molecule_elements(root)
    if not found:
        return svg_bytes, []

    parents = _parent_map(root)
    molecules = []
    for mol_idx in sorted(found):
        roles = found[mol_idx]
        parts = []
        for role in _ordered_roles(roles):
            ids = {element.get("id") for element in roles[role]}
            parts.append((ROLE_TITLES.get(role, role.title()), _keep_only(svg_bytes, ids)))
        name = (molecule_names or {}).get(mol_idx) or f"Molecule {mol_idx + 1}"
        molecules.append((name, parts))

    for roles in found.values():
        for elements in roles.values():
            for element in elements:
                parent = parents.get(element)
                if parent is not None:
                    parent.remove(element)
    return _dump(tree), molecules


def to_inkscape_layers(svg_bytes: bytes, molecule_names: dict[int, str] | None = None) -> bytes:
    """Restructure the figure SVG into Inkscape layers.

    Everything that is not a molecule goes into an ``Image`` layer; each
    molecule becomes its own layer (above it) with ``Bonds``/``Atoms``/
    ``Labels`` sub-layers. Without tagged molecules the figure still gets a
    single ``Image`` layer so it opens as an editable layer in Inkscape.
    """
    tree = _load(svg_bytes)
    root = tree.getroot()
    found = _molecule_elements(root)
    parents = _parent_map(root)

    def _layer(label: str, layer_id: str):
        layer = ET.Element(f"{{{SVG_NS}}}g")
        layer.set("id", layer_id)
        layer.set(f"{{{INKSCAPE_NS}}}groupmode", "layer")
        layer.set(f"{{{INKSCAPE_NS}}}label", label)
        return layer

    detached: dict[int, dict[str, list]] = {}
    for mol_idx, roles in found.items():
        for role, elements in roles.items():
            for element in elements:
                parent = parents.get(element)
                if parent is not None:
                    parent.remove(element)
                detached.setdefault(mol_idx, {}).setdefault(role, []).append(element)

    image_layer = _layer("Image", "layer_image")
    for child in list(root):
        if child.tag in _PASSTHROUGH_TAGS:
            continue
        root.remove(child)
        image_layer.append(child)
    root.append(image_layer)

    for mol_idx in sorted(detached):
        name = (molecule_names or {}).get(mol_idx) or f"Molecule {mol_idx + 1}"
        mol_layer = _layer(name, f"layer_molecule_{mol_idx + 1}")
        roles = detached[mol_idx]
        for role in _ordered_roles(roles):
            role_layer = _layer(ROLE_TITLES.get(role, role.title()), f"layer_molecule_{mol_idx + 1}_{role}")
            for element in roles[role]:
                role_layer.append(copy.deepcopy(element))
            mol_layer.append(role_layer)
        root.append(mol_layer)
    return _dump(tree)


def find_inkscape() -> str | None:
    """Path to the Inkscape executable, or None when it isn't installed."""
    import os
    import shutil
    import sys

    found = shutil.which("inkscape")
    if found:
        return found
    candidates = []
    if sys.platform == "win32":
        for base in (
            os.environ.get("ProgramFiles"),
            os.environ.get("ProgramFiles(x86)"),
            os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs"),
        ):
            if base:
                candidates.append(os.path.join(base, "Inkscape", "bin", "inkscape.exe"))
    elif sys.platform == "darwin":
        candidates.append("/Applications/Inkscape.app/Contents/MacOS/inkscape")
    for candidate in candidates:
        if os.path.isfile(candidate):
            return candidate
    return None
