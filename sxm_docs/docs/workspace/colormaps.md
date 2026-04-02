# Colormaps & Contrast

SXM Viewer gives each image view its own display state, including colormap and contrast settings.

---

## Changing the colormap

Use the preview or pop-out controls to switch the active colormap for the current view. Colormaps can also be managed in canvas workflows and comparison tools where different panels have different display needs.

Examples from the app include:

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
