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

## Preview and popup publication mode

The right-click **Display → Preset → Publication** action is separate from the
publication-canvas presets above. On the preview and popups it applies a
compact figure-style treatment independently to each image panel:

- a horizontal colorbar is shown for every panel;
- only the displayed lower and upper color limits are labelled inside the bar;
- common SPM channels receive compact labels: `I` for current, `z` for
  topography, and `Δf` for frequency-shift/nc-AFM channels;
- scale bars are shown, while titles, axes, image dimensions, and
  profile/angle/acquisition overlays are hidden;
- the existing panel order, grid arrangement, and number of channels are
  preserved.

The preset changes presentation only. It does not change the data, colormap,
contrast limits, or filtering.

Preview and popup presets are temporary: they are not saved as application
defaults, session state, or collection state. Use **Restore previous display**
in the Preset menu, or `Ctrl+Z`, to return to the exact display state from
before the preset was applied.

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
