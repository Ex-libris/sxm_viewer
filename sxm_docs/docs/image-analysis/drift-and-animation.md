# Drift Correction & Animation

Two related workflows for working with a *sequence* of scans - typically repeated images of the same area taken over time - available from the thumbnail right-click menu after selecting two or more images.

---

## Drift correction

Select two or more thumbnails of the same region taken at different times, then right-click -> **Drift-correct and export...**. This aligns every image to the first one in the selection using sub-pixel image registration (scikit-image's phase cross-correlation, falling back to OpenCV's ECC method if scikit-image isn't installed), then crops all frames to their common overlapping region so the whole stack shares one consistent field of view.

!!! note
    Requires `scipy`. For the alignment step itself you also need `scikit-image` or `opencv-python` - without either, frames are still cropped to a common region but not actually shifted into alignment.

The **Drift correction preview** dialog that opens shows:

- the measured shift (in pixels) for every frame, relative to the first
- the resulting common crop size
- an animated preview of the aligned stack, with a colormap picker and a playback-speed slider

From there, three outputs are available:

| Button | Result |
|---|---|
| Save aligned PNGs... | Writes each aligned, cropped frame to disk as a PNG |
| Save animation... | Saves the aligned stack directly as a GIF |
| Save corrected copies to thumbnails | Inserts each aligned frame back into the thumbnail grid as a virtual copy |

---

## Animation from a selection

Select two or more thumbnails, then right-click -> **Create animation from selection...** to build a GIF, MP4, or PNG sequence directly from them - no alignment step. If the frames need registering first, use [Drift correction](#drift-correction) instead (its own **Save animation...** button covers that case).

!!! note
    Requires the `imageio` package.

Frames are always ordered alphabetically by filename, not by the order you selected them in. The **Export animation** dialog offers:

- **Format**: GIF, MP4, or a numbered PNG sequence
- **FPS** and **Duration** - linked, so changing one adjusts the other for the fixed frame count
- **Include scale bar / markers-overlays / molecules** toggles
- **Resolution**: Auto (native size), 720p, 1080p, or a custom width/height

A live preview canvas shows the assembled animation before you save it.

---

## Related pages

- [Thumbnail Grid](../browsing/thumbnail-grid.md)
- [Copy & Export Images](../export-and-sharing/export-images.md)
