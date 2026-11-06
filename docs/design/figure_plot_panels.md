# Figures: plot panels (profiles, spectra) - design sketch

Status: **proposal**, nothing implemented. Written to agree on the model
before code, because this feature has many options and the wrong data model
would be expensive to undo.

## Why

SPM figures are rarely images only. The common cases are an image plus a
line profile across it, and an image plus spectra taken on it. Today users
export the plot from a profile or spectrum window and line the parts up in
PowerPoint or Inkscape, so fonts, letters and sizes stop matching. The figure
window already knows the images, letters, fonts and page width; plot panels
would let the same figure carry the plots, drawn by the same renderer.

## Scope (first version)

| Panel type | Data source in the app | What it shows |
| --- | --- | --- |
| **Line profile** | a line in *panel nm* on one image | height (or signal) along the line, one curve per frame / channel it is drawn on |
| **Point spectra** | point spectra assigned to an image (`SpectroCompareController.all_specs_for_image`) | the chosen spectra (channel vs. bias or z), markers numbered on the image |

Out of scope for v1: grid (3ds) maps (they are images - already possible as
virtual copies), FFTs, histograms, KPFM fits (the matrix-fit dialog exports
those).

## Where plots go - stay inside the grid

The grid stays the only layout, so the composition rules (page width,
alignment, letters) keep working. A plot occupies **one extra row** (frames
across) or **one extra column** (frames down) with one plot cell per frame,
or a **single wide cell** spanning the frames:

```
 Δz = 0    +20 pm   +40 pm
 [img a]   [img b]  [img c]       <- channel row(s), as today
 [plot d]  [plot e] [plot f]      <- per-frame plot row: the profile on each frame
 [        plot g (all)       ]    <- or one spanning plot: every frame overlaid
```

`spec["rows"]` gains rows of `kind == "plot"`; `figure_layout` gives them
their own height (an aspect such as 0.6 x panel width, adjustable), and the
renderer draws axes in them instead of `imshow`. Letters, Ctrl+wheel sizing,
the − / + tabs and drag-to-reorder treat them like any other row.

## Data model (figure state)

```python
state["plots"] = [
  {"id": "p1", "type": "profile",
   "line": [x0, y0, x1, y1],          # panel nm, same convention as molecules
   "width_nm": 0.0,                   # averaging width across the line (0 = 1 px)
   "rows": ["TopoFwd"],               # which channel rows are sampled
   "layout": "per_frame" | "span",    # one plot per frame, or one overlaid plot
   "offset": 0.0, "normalise": False, # waterfall offset / normalisation (span)
   "color_by": "frame" | "fixed",
   "x_range": None, "y_range": None}, # None = automatic
  {"id": "s1", "type": "spectra",
   "specs": [{"path": ..., "matrix_index": None}, ...],   # stable spec keys
   "channel": "df", "x": "bias",
   "layout": "span", "offset": 0.0, "labels": "number", ...},
]
```

- The **line lives in panel nm** (like molecular models), so a crop or a
  reorder keeps it on the same feature, and one line can be sampled on every
  frame of a series - the main use case: one profile across a Δz series.
- Spectra are referenced by the **stable spec key** (path + matrix index,
  `_inset_spec_key`), never `id(spec)` (see the position-inset lessons in
  CLAUDE.md), and their markers are mapped with `_map_spec_to_pixels`,
  turned and cropped like the image (the raster-frame rule).
- **Sampling happens in the controller** (GUI thread, from `panel()` data, so
  per-image filters / Zero apply), and the result is plain arrays passed to
  the Qt-free renderer, the same split as images.

## Interaction

- **Create**: right-click an image > "Line profile…" then drag a line on it
  (the same gesture as crop: a rubber band); or "From preview" imports the
  preview's saved profiles (they are in preview data coords; convert as
  `preview_molecules` does with `_frame_origin_nm`). "Spectra on this image…"
  lists the image's spectra with check boxes.
- **On the image**: the line / markers are drawn in the plot's colours; drag
  the ends to move them (live re-sample), Shift = horizontal / vertical.
- **On the plot**: click = its card (curves, axis ranges, offset, legend,
  units); Ctrl+wheel = size, like everything else; double-click = automatic
  ranges.
- A new **"Plots" chip** in the Format bar lists the figure's plots.

## Rendering rules (journal look)

- Same font, sizes and line widths as the figure's text style (`font_rc`);
  axis lines 0.5-0.75 pt, ticks out, no grid by default, unit in brackets in
  the axis label, SI prefix per axis (one helper with the colour bars).
- Curve colours: per frame from one qualitative palette, matching the marker
  / line colours on the images; a colour-blind-safe default.
- Profiles in display units of the row (the same factor as its colour bar).

## Outputs

- Source data: one CSV per plot (x, then a column per curve), written next
  to the images' CSVs by the submission bundle.
- Caption: "(d-f) Height profiles along the line in (a-c)"; spectra: "(g)
  Δf(V) spectra at the positions marked in (a)".
- Layered SVG: plots get their own Inkscape layer.

## Open questions (to decide before coding)

1. Per-frame plot row vs. one spanning plot as the *default*. A series reads
   best as one overlaid plot; a zoom figure as one plot per frame.
2. Should a profile follow a crop (re-sampled inside the crop only) or keep
   its full length and be clipped? Proposed: clipped to the crop, with a
   note.
3. Spectra across frames: spectra belong to one image. In a series, show
   them under that image only (with empty cells for the others), or as one
   spanning plot? Proposed: spanning.
4. Matrix / grid spectra (1000s of points): pick by click on the image, cap
   the number of curves, or allow averaging over a region?

## Phasing

1. Profile, per frame and spanning, sampled in panel nm; markers on images;
   CSV + caption. (Reuses: crop rubber band, `panel()`, molecules' panel-nm
   convention.)
2. Point spectra (single and Z-stacks), spanning plot, numbered markers.
3. Editing polish: dragging line ends on the figure, "From preview" import,
   waterfall offsets.
