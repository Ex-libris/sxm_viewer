"""Inkscape-layered SVG for figure recipes (Qt-free).

``figures.height_series.render`` tags every artist with a ``gid``
(``image_r_f``, ``letter_r_f``, ``scalebar_r_f``, ``zoombox_r_f``,
``colorbar_...``, ``header_f``, ``channel_r``, ...). Matplotlib writes each
as a ``<g id=...>``; this module moves those groups into named Inkscape
layers, so the figure opens in Inkscape as

    Images / Zoom outlines / Notes / Scale bars / Letters / Colour bars / Labels

- each layer can be locked, hidden or restyled at once (e.g. all letters).
Matplotlib's SVG has no group transforms (coordinates are absolute) and
clip paths live in ``<defs>``, so moving a group between parents keeps it
exactly where it was.
"""
from __future__ import annotations

import io
import re
import xml.etree.ElementTree as ET

SVG_NS = "http://www.w3.org/2000/svg"
XLINK_NS = "http://www.w3.org/1999/xlink"
INKSCAPE_NS = "http://www.inkscape.org/namespaces/inkscape"
SODIPODI_NS = "http://sodipodi.sourceforge.net/DTD/sodipodi-0.dtd"

for _prefix, _uri in (("", SVG_NS), ("xlink", XLINK_NS), ("inkscape", INKSCAPE_NS),
                      ("sodipodi", SODIPODI_NS)):
    ET.register_namespace(_prefix, _uri)

# (layer label, id pattern) in bottom-to-top paint order.
LAYERS = (
    ("Images", re.compile(r"^image_\d+_\d+$")),
    ("Zoom outlines", re.compile(r"^zoom(box|label|frame)_\d+_\d+$")),
    ("Notes", re.compile(r"^note_\d+_\d+$")),
    ("Scale bars", re.compile(r"^scalebar_\d+_\d+(_label)?$")),
    ("Letters", re.compile(r"^letter_\d+_\d+$")),
    ("Colour bars", re.compile(r"^colorbar_")),
    ("Labels", re.compile(r"^(header|channel)_")),
)


def layered_svg(svg_bytes):
    """Return ``svg_bytes`` restructured into Inkscape layers (unchanged
    input on any parse problem - a plain SVG is still a valid result)."""
    try:
        tree = ET.parse(io.BytesIO(svg_bytes))
    except ET.ParseError:
        return svg_bytes
    root = tree.getroot()
    parent_of = {child: parent for parent in root.iter() for child in parent}
    figure = None
    for el in root.iter(f"{{{SVG_NS}}}g"):
        if el.get("id", "").startswith("figure_"):
            figure = el
            break
    if figure is None:
        return svg_bytes
    buckets = {label: [] for label, _ in LAYERS}
    for el in list(root.iter()):
        gid = el.get("id", "")
        for label, pattern in LAYERS:
            if pattern.match(gid):
                # Skip groups already inside a matched group (one move each).
                anc, nested = parent_of.get(el), False
                while anc is not None:
                    if any(p.match(anc.get("id", "")) for _l, p in LAYERS):
                        nested = True
                        break
                    anc = parent_of.get(anc)
                if not nested:
                    buckets[label].append(el)
                break
    moved = False
    for label, _pattern in LAYERS:
        items = buckets[label]
        if not items:
            continue
        layer = ET.SubElement(figure, f"{{{SVG_NS}}}g", {
            f"{{{INKSCAPE_NS}}}groupmode": "layer",
            f"{{{INKSCAPE_NS}}}label": label,
            "id": "layer_" + label.lower().replace(" ", "_"),
        })
        for el in items:
            par = parent_of.get(el)
            if par is not None:
                par.remove(el)
            layer.append(el)
            moved = True
    if not moved:
        return svg_bytes
    out = io.BytesIO()
    tree.write(out, encoding="utf-8", xml_declaration=True)
    return out.getvalue()
