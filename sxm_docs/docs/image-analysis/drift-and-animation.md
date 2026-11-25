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

Select two or more thumbnails, then right-click -> **Create animation...** to build a GIF, MP4, or PNG sequence directly from them - no alignment step. If the frames need registering first, use [Drift correction](#drift-correction) instead (its own **Save animation...** button covers that case).

!!! note
    Requires the `imageio` package, plus `imageio-ffmpeg` for MP4
    (`python -m pip install imageio imageio-ffmpeg`).

Every frame is rendered **as the main preview shows it**: each image's own colormap, contrast, filters, Zero, rotation, display preset, scale bar, molecules and overlays. The **Export animation** dialog plays the result live and offers:

| Row | Options |
|---|---|
| Frames | Order: thumbnail order, acquisition time, bias or file name; Reverse |
| Look | **Colormap**: as displayed (each image's own) or one colormap for every frame (favourites, search, gallery); **Colour range**: as displayed, or *Same for all frames* (one range over the whole series, so brightness changes between frames are real); **Label** on each frame: file name, bias, CH height, date/time or frame number |
| Show | Scale bar; Molecules - each image's own, or the shown image's model on every frame; other overlays (spectroscopy markers, profiles, acquisition box) |
| Output | GIF, MP4 or PNG frames; *Whole preview* (with colour bar, title, axes) or *Image only*; width in pixels |
| Timing | Frames per second, a pause on the last frame, Loop / Forth and back / Play once |

Nothing in the viewer is changed by these options: they apply to the exported frames only, and the previously shown image comes back when the dialog closes. The last settings are remembered.

---

## Related pages

- [Thumbnail Grid](../browsing/thumbnail-grid.md)
- [Copy & Export Images](../export-and-sharing/export-images.md)
