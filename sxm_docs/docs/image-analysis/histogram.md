# Histogram & Contrast

SXM Viewer includes histogram-based range controls for adjusting image contrast without changing the underlying data.

![Histogram live contrast adjustment](../assets/screenshots/histogram-live-contrast.gif){ width="900" }

---

## Opening the histogram tool

Use the image display controls or the relevant right-click menu entry to open the histogram/range dialog for the current preview or pop-out.

The histogram tool has been split into its own controller internally, but from the user side it remains part of the normal image-adjustment workflow.

---

## What you can do

Typical histogram tasks include:

- inspecting the data-value distribution for the current view
- adjusting the active display range
- resetting contrast after experimental changes
- refining the view before export or screenshot capture

---

## Interaction with other tools

Histogram edits are display operations, not destructive processing. They work alongside:

- relative-zero display
- colorbar orientation and display settings
- profiles and overlays
- popup-specific view state

Recent fixes explicitly prevented histogram interactions from creating or updating profiles when profile mode is not active.

---

## Auto-contrast behavior

The project history shows special handling for partial scans and invalid regions during auto-contrast. In practice, that means the contrast logic tries to ignore obvious invalid rows rather than letting them dominate the visible range.

---

## Tips

!!! tip
    If an image looks washed out, first try the histogram/range dialog before applying a filter.

!!! tip
    For constant-height images, test histogram changes together with ++0++ relative-zero mode, since that can produce a more interpretable zero-anchored colorbar.

---

## Related pages

- [Colormaps & Contrast](../workspace/colormaps.md)
- [Preview & Popups](preview-and-popups.md)
- [Filters & Processing](filters.md)
