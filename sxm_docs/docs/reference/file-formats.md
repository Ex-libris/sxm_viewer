# Supported File Formats

SXM Viewer natively reads Omicron/Anfatec SPM data, and supports Nanonis data through a conversion adapter.

---

## Imaging formats

| Format | Files | Notes |
|---|---|---|
| Omicron / Anfatec | `.txt` header + binary channel files | Native format; full header and multi-channel support |
| Nanonis | `.sxm` | Converted on first load into an Omicron-style header/channel cache, then handled identically to native files |

Every loaded scan is automatically tagged as **constant-height (CH)** or **constant-current (CC)** based on header keywords and a small `dz`-offset heuristic - see [Constant-height frame detection](../browsing/loading.md#constant-height-frame-detection) for how this shows up in the thumbnail grid.

!!! note
    Nanonis conversion caches its results in a `.sxmviewer_nanonis/` folder next to the source data, so repeat loads of the same folder are fast. This folder is safe to delete; it is rebuilt automatically the next time the folder is loaded.

---

## Spectroscopy formats

Supported spectroscopy workflows include:

- single-point spectroscopy traces (Omicron `.dat` files)
- matrix / grid (CITS-style) spectroscopy, where many spectra are arranged on a 2D grid over an image
- KPFM bias-spectroscopy traces, analyzed with the same parabola-fit tooling as any other spectroscopy channel

See [Spectroscopy Overview](../spectroscopy/overview.md) for how these are browsed and analyzed.

---

## Export-only formats

Scan channel data (not the whole session) can be exported as:

- an XYZ point-cloud text file (x, y, z columns with a metadata header)
- a WSxM-style `.stp` binary raster file

See [CSV & Data Export](../export-and-sharing/data.md) for exact formats and entry points. `.xyz` files can also be **imported**, but only as a molecule-model overlay to place on top of a scan - not as scan/image data - see [Molecule Overlays](../image-analysis/molecules.md).

---

## Notes

!!! note
    The exact set of imported channels can depend on the header metadata available in the source files.

!!! note
    Nanonis support is implemented through a conversion adapter that normalizes `.sxm` scans into the same header/channel representation used natively, rather than a fully separate code path - so once loaded, Nanonis and Omicron/Anfatec data behave identically throughout the app.

---

## Related pages

- [Loading Data](../browsing/loading.md)
- [Spectroscopy Overview](../spectroscopy/overview.md)
- [Matrix Scans](../spectroscopy/matrix.md)
