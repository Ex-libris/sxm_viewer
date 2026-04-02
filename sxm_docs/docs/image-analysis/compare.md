# Image Comparison (A/B)

The A/B comparison tool lets you align and quantitatively compare two images: forward/backward scan pairs, before/after tip changes, or any two frames you want to subtract.

---

## Setting up a comparison

Right-click any preview or pop-out → **Compare** submenu:

| Action | Effect |
|---|---|
| Set This View as Compare A | Freezes the current view as slot A |
| Set This View as Compare B | Freezes the current view as slot B |
| Compare A with This | Sets A to the stored slot, B to this view, opens dialog |
| Compare B with This | Sets B to the stored slot, A to this view, opens dialog |
| Open A/B Comparison | Opens the dialog with current A and B |
| Swap A and B | Swaps the two slots |
| Clear Compare Selection | Resets both slots |

The comparison dialog opens as a normal tracked popup window.

---

## Compare dialog panels

The dialog renders four panels simultaneously:

1. **A** — image A with the current colormap
2. **B aligned** — image B after applying the current transform
3. **A − B** — difference map (always on `RdBu_r`, symmetric)
4. **|A − B|** — absolute difference (on `magma`)

Metrics shown: overlap area, Pearson r, RMSE, fit slope.

---

## Manual alignment controls

| Control | Effect |
|---|---|
| Rotation | Rotate B relative to A (degrees) |
| Shift X / Shift Y | Translate B in physical units |
| Intensity matching | Scale B intensity to minimise A−B |
| Lock A/B range | Fix both images to the same color scale |
| Stretch | Linear or histogram-equalised display |

---

## Automatic alignment

**Auto fit** aligns B to A automatically. Two modes:

- **Translate** — finds the best XY offset
- **Rigid** — finds the best rotation + translation

---

## Landmark-based alignment

For cases where auto fit is insufficient:

1. Enable **Pick landmarks**.
2. Click A1 on image A, then the corresponding point B1 on image B. Continue for A2/B2, etc.
3. Choose **Translate** or **Rigid**.
4. Press **Fit from points**.

Numbered landmark markers are drawn on both panels. Use **Undo point** or **Clear points** to reset.

---

## Profile and overlap analysis

- **Profile** mode adds a linked crosshair across all four panels with a status readout.
- **Overlap scatter** panel shows a pixel-wise scatter plot of A vs B with a 1:1 line, Pearson r, and fit slope.

---

## Export

**Export** saves a 300 DPI 4-panel PNG and an A−B difference CSV.