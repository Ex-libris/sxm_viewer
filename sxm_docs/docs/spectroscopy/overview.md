# Spectroscopy Overview

SXM Viewer loads spectroscopy files automatically when a folder is opened, associating them with images using robust timestamp-based matching. Spectroscopy data and scan images live in the same workspace rather than requiring separate tools.

---

## Spectroscopy thumbnails

Associated spectroscopy files appear as **miniature thumbnails** within the main thumbnail grid, positioned near their spatially associated scan image. Markers on the scan thumbnails indicate where each spectroscopy was acquired.

### Thumbnail marker customisation

Right-click a spectroscopy thumbnail → **Miniature channel** to choose which channel drives the miniature plot. Marker size, colour, and symbol are also customisable for better visibility on different backgrounds.

### Selection

| Gesture | Effect |
|---|---|
| Single click | Select spectroscopy |
| Shift+Click | Range-select |
| Ctrl+Click | Add/remove from selection |
| Drag | Rubber-band selection |
| Double-click | Open spectroscopy popup |

---

## Spectroscopy browser

Open the **Spectroscopy Browser** from the toolbar. It presents all associated spectroscopy files as a **multi-column table** with sortable columns. From the browser you can:

- Select single or multiple spectroscopies
- Open them in the spectroscopy popup
- Apply a channel preset to all selected entries

---

## Spectroscopy popup

The spectroscopy popup plots the selected spectrum (or spectra) with full axis labels and units. Multiple selected spectroscopies overlay on the same axes. Waterfall plotting is available for visualising sets of spectra with a vertical offset.

### Colour control

Each spectroscopy curve can be given an individual colour from the expanded palette (includes black, gray, white, and an arbitrary colour picker), or colours can be assigned from a shared palette applied to the whole selection.

### Typography

Font family, bold, italic, and underline are accessible from the right-click **Typography** menu and stay consistent with the rest of the GUI.

### Export

Right-click the spectroscopy plot → **Copy as PNG** or **Copy as SVG**.

---

## Spatial markers on images

When spectroscopy display is enabled, markers appear on the preview and pop-outs at the acquisition positions. Toggling spectroscopy on/off does not reload the data — the association is cached.

Marker positions are correctly placed in both absolute and relative axes display modes.

---

## Supported spectroscopy types

- Single-point I(V), I(z), df(V), df(z) traces
- Grid / matrix spectroscopy (see [Matrix Scans](matrix.md))
- KPFM data (see [KPFM](kpfm.md))
- Parabola fits (see [Parabola Fits](parabolas.md))
- WSxM XYZ export