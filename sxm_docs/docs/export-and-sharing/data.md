# CSV & Data Export

SXM Viewer can export numerical data as well as rendered figures - but exactly what format and how depends on which tool you're exporting from. This page lists every data-export path and what it actually produces.

---

## Profile data - clipboard only

The profile dialog's copy action builds a tab-separated table (one distance/value column pair per active and saved profile) and puts it on the clipboard, ready to paste into a spreadsheet. There is currently no "save profile as CSV" file dialog - if you need a file on disk, paste the clipboard contents into a spreadsheet application and save from there.

---

## Whole-channel/image data - XYZ and WSxM `.stp`, not CSV

There is no direct CSV export for a whole image/channel. Instead:

- **XYZ point cloud** - a plain text file with x, y, z columns and a metadata header, one row per pixel.
- **WSxM `.stp`** - a binary raster file in WSxM's native format.

Both are reachable from the export menu wherever a channel is displayed. If you specifically need a CSV of raw channel values, export as XYZ and reformat - there is no other path.

---

## Image comparison (A/B) - automatic, diff-only CSV

Exporting a comparison figure (see [Image Comparison (A/B)](../image-analysis/compare.md)) only prompts for a PNG path, but it **automatically** writes a CSV alongside it with the same name (`<name>_AminusB.csv`) - no separate export step. That CSV contains only the **A minus B difference** array (not A, not B, not |A-B|, even though all of those appear in the exported figure itself), with the first row/column holding x/y coordinates in the current display unit.

---

## Matrix / grid spectroscopy - selection export

In the matrix spectroscopy viewer, selecting one or more pixel positions and choosing **Export selection to CSV** writes a long-format table with columns `channel, index, x_nm, y_nm, bias, bias_unit, value` - one row per (selected spectrum x bias point). See [Matrix Scans](../spectroscopy/matrix.md).

---

## KPFM fit-vs-height trend - dedicated export

The KPFM fit-trend dialog (plotting LCPD, `a`, `c`, or RMSE against Z/height across a stack of fits - see [KPFM Data](../spectroscopy/kpfm.md)) has its own **Export CSV** button (++ctrl+e++) that saves the fit-derived parameters as a table, separate from the trace-value exports above.

---

## When to use data export

Use these workflows when you need to:

- replot in another tool
- perform custom fitting outside the GUI
- archive the numerical values behind a figure
- compare several exported traces in a spreadsheet or script

---

## Related pages

- [Image Comparison (A/B)](../image-analysis/compare.md)
- [Spectroscopy Overview](../spectroscopy/overview.md)
- [Matrix Scans](../spectroscopy/matrix.md)
- [KPFM Data](../spectroscopy/kpfm.md)
- [Export to SVG & PDF](vector.md)
