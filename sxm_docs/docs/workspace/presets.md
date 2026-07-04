# Display Presets

Display presets provide fast, repeatable visual styles for figure preparation and analysis.

---

## Canvas presets

The publication canvas exposes three global presets:

| Preset | Description |
|---|---|
| Clean | No title, overlay info, metadata/unit bar, or colorbar; scale bar on |
| Analysis | Title, scale bar, and a colorbar with ticks (right-positioned) on; metadata bar and overlay info off |
| Publication | Same toggles as Clean |

!!! note
    Clean and Publication currently apply the identical set of toggles - there is no visible difference between them as implemented.

Manual changes return the canvas to **Custom** state.

See [Publication Canvas](canvas.md).

---

## Figure layout presets (profile and spectroscopy plots)

Separately from the canvas presets above, profile and spectroscopy plot windows offer their own **figure layout preset** picker, sizing the plot window and its typography for a specific output target:

| Preset | Target size |
|---|---|
| Interactive | No fixed size - normal on-screen window |
| Journal 1-col square (88mm) | 88 mm square |
| Journal 1-col square (85mm) | 85 mm square |
| Journal 1.5-col square (114mm) | 114 mm square |
| Journal 2-col square (174mm) | 174 mm square |
| Slides square (127mm) | 127 mm square |

Each preset also sets a matching font family/scale and legend font size and line width, so a plot sized for a journal column reads correctly at that physical size rather than needing manual font tweaks afterward.

---

## Why presets matter

Presets let you switch quickly between:

- an exploratory analysis view
- a cleaner presentation view
- a publication-style figure layout, sized correctly for its destination

This is faster and more reliable than toggling every overlay - or every font size - one by one.

---

## Related workflows

Popup/style workflows also allow one window's style to be copied to others, making presets part of a broader figure-preparation pipeline rather than a canvas-only convenience.

---

## Related pages

- [Colormaps & Contrast](colormaps.md)
- [Dark Mode & Typography](typography.md)
- [Publication Canvas](canvas.md)
