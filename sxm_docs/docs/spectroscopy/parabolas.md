# Parabola Fits

SXM Viewer fits a quadratic curve (`a·V² + b·V + c`) to bias-dependent spectroscopy traces - the same fit is used everywhere a "Fit parabola" button appears: single-spectrum popups, batch fitting across a comparison set, and per-pixel across a whole matrix grid.

---

## What the fit reports

Fitting a trace reports the quadratic coefficients `a`, `b`, `c` (each with a propagated uncertainty) and the fit's RMSE. The headline result, though, is the fit's **vertex position** - `V = -b/(2a)` - labeled **LCPD** (local contact potential difference) throughout the app, since this fit is used almost exclusively on bias-spectroscopy/KPFM data where the vertex of the parabola *is* the contact potential.

---

## Single-spectrum fit

Open a spectroscopy popup and use **Fit parabola** (see [Spectroscopy Overview](overview.md)). The fitted curve is drawn over the trace, and the coefficients/LCPD/RMSE are shown alongside it.

---

## Batch fitting across several spectra

In the spectroscopy comparison dialog, **Fit selected** (++f++) or **Fit all** runs the same fit independently across every trace in the set, building a results table with each trace's name, X/Y/Z position, `a`, `b` (LCPD), `c` (with uncertainties), and RMSE. **Export CSV** (++ctrl+e++) saves that table.

### Comparing LCPD between two traces

Shift-click two LCPD guide lines in the comparison plot to draw a **ΔLCPD** annotation between them - a quick way to read off the contact-potential difference between two specific spectra without doing the subtraction yourself.

---

## Matrix-wide fitting

Fitting every pixel in a matrix/grid dataset at once produces a set of 2D fit-parameter maps rather than a table - see [Matrix Scans](matrix.md) and [KPFM Data](kpfm.md#matrix-wide-fitting).

---

## Converting a frequency-shift spectrum to force

A related, separate tool in the comparison dialog - **Convert to force** - reconstructs a force-vs-distance curve from a Δf(z) spectrum using the Sader-Jarvis method, given the cantilever's resonance frequency, spring constant, oscillation amplitude, and Q factor. These parameters can be remembered between sessions so you don't have to re-enter them every time.

---

## Related pages

- [Spectroscopy Overview](overview.md)
- [Matrix Scans](matrix.md)
- [KPFM Data](kpfm.md)
- [CSV & Data Export](../export-and-sharing/data.md)
