# Filters & Processing

![Image filter pipeline workflow](../assets/screenshots/filters-pipeline.gif){ width="900" }

## Applying a filter

Right-click the preview or a pop-out → **Apply filter** submenu, or use the thumbnail context menu for batch application to selected thumbnails.

Filters are applied non-destructively: the original data is preserved and the processed view is stored alongside it. ++ctrl+z++ undoes the last filter step.

---

## Available filters

### Laplacian

Highlights edges and fine surface features by computing the second spatial derivative.

Parameters (set via a quick dialog before applying):

| Parameter | Effect |
|---|---|
| Sigma | Pre-smoothing (Gaussian) before Laplacian |
| Stencil | 4-neighbor or 8-neighbor Laplacian kernel |
| Absolute response | Show \|∇²f\| instead of signed response |

Last-used parameters are remembered across sessions.

### Other filters

Additional filters are available in the **Custom pipeline** dialog (see below). The filter set is under active development.

---

## Custom filter pipeline

**Right-click → Apply filter → Custom pipeline...** opens a dialog to chain multiple filter steps into a single processing pipeline. Each step has dynamic parameter controls (only relevant controls are shown for the selected filter type).

The pipeline can be previewed before committing.

---

## Applying filters to multiple images

Select several thumbnails (Shift+Click, Ctrl+Click, or drag rubber-band) then right-click → **Apply filter** to batch-process all selected images with the same filter and parameters.
