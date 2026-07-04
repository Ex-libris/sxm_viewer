# KPFM Data

KPFM-related data can be browsed and analysed inside the same SXM Viewer workspace as topography and other spectroscopy content - both as image channels and as bias-spectroscopy traces with dedicated fitting tools.

---

## Where KPFM appears

Depending on the acquisition and import path, KPFM information may appear as:

- image channels in the main preview or pop-outs
- spectroscopy traces, analyzed with the same parabola-fit tooling as any other bias spectroscopy (see [Parabola Fits](parabolas.md)) - the fit's headline output, the vertex of the fitted parabola, is labeled **LCPD** (local contact potential difference) throughout the app
- associated data that can be exported or compared like other supported channels

The preview channel selector explicitly supports workflows where KPFM is one of several acquired channels.

---

## Working with KPFM image channels

Treat KPFM image channels like other preview channels: switch to the channel in the preview or popup, adjust contrast and colormap locally, and use cropping, profiles, and export tools as needed.

---

## LCPD vs. Z: the fit-trend dialog

For a stack of KPFM spectra taken at different tip heights, the **fit-trend dialog** ("Fit vs Z", available once a stack has fits with usable Z metadata) plots a chosen fit metric - **LCPD**, `a`, `c`, or RMSE - against Z/height across the whole stack, with:

- **Show errors** - display propagated fit uncertainty as error bars
- **Relative Z** - shift the Z axis so it starts at the fitted-Z minimum, instead of an absolute height
- **Sort by Z** - order points by height rather than acquisition order

Its own **Export CSV** (++ctrl+e++) saves the plotted metric-vs-Z table - see [CSV & Data Export](../export-and-sharing/data.md).

---

## Matrix-wide fitting

For a KPFM matrix/grid dataset, **Fit matrix parabolas...** (see [Matrix Scans](matrix.md)) fits every pixel independently and produces seven 2D maps: `a`, **LCPD**, `c`, each of their fit errors, and RMSE - each with its own colormap and units. The dialog auto-detects which channel in the matrix is the Δf/KPFM channel to fit (scoring channel names containing "df"/"frequency"/"kpfm" and similar).

From there you can:

- adjust the color **scale mode** (full range, clipped percentiles, or centered around zero)
- hover a map to read its exact value at a point
- **Save maps...** - writes the maps plus a JSON metadata sidecar
- **Export WSxM XYZ...** - one XYZ file per parameter map

This turns a whole grid of individual bias sweeps into a spatial map of contact potential (and the other fit parameters) across the scanned area in one action.

---

## Related: force reconstruction

A related tool converts a Δf(z) spectrum into a force-vs-distance curve (Sader-Jarvis method) using the cantilever's resonance frequency, spring constant, amplitude, and Q - see [Parabola Fits](parabolas.md#converting-a-frequency-shift-spectrum-to-force).

---

## Related pages

- [Preview & Popups](../image-analysis/preview-and-popups.md)
- [Spectroscopy Overview](overview.md)
- [Parabola Fits](parabolas.md)
- [Matrix Scans](matrix.md)
- [CSV & Data Export](../export-and-sharing/data.md)
