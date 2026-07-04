# Colormaps & Contrast

SXM Viewer gives each image view its own display state, including colormap and contrast settings.

---

## Changing the colormap

Use the preview or pop-out colormap controls to switch the active colormap. This isn't just a temporary display change: it's stored as a **per-file, per-channel override** that persists across sessions - so switching the colormap while viewing one file/channel doesn't affect any other file's or channel's colormap, and reopening that same file/channel later reapplies the colormap you chose. The global default colormap (used for anything without its own override) only changes when nothing is currently selected.

Right-clicking a **virtual copy** in the thumbnail grid has its own **Colormap** submenu, with a handful of featured colormaps, a "More..." option for the full list, and a **Use global thumbnail/preview cmap** action to clear that copy's override and fall back to the default.

Colormaps can also be managed in canvas workflows and comparison tools where different panels have different display needs. Examples from the app include:

- scientific grayscale or topography maps for normal viewing
- fixed diverging maps for signed difference views
- dedicated maps for magnitude-style views

---

## Contrast controls

Contrast can be adjusted through the histogram/range workflow and through reset/auto-range actions.

Important related behaviors include:

- **relative-zero** mode can clamp the lower bound to zero
- auto-range respects valid scan regions where possible
- popup views can preserve their own local display state

---

## Shared vs local behavior

Some display decisions propagate across preview and pop-outs, while others are intentionally local to a particular view.

Recent project work also added explicit colorbar orientation control with vertical and horizontal choices.

---

## In comparison workflows

The A/B comparison tool uses panel-specific rendering rules. For example, the signed difference map uses a diverging map, while absolute-difference views use a different magnitude-oriented map.

See [Image Comparison (A/B)](../image-analysis/compare.md).

---

## Related pages

- [Histogram & Contrast](../image-analysis/histogram.md)
- [Overlays](overlays.md)
- [Display Presets](presets.md)
