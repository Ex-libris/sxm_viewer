# Handling Partial/Aborted Scan Contrast (Thumbnails & Preview)

This document lists the functions that suppress blank/aborted regions when computing auto-contrast for thumbnails and the Preview. Adjust thresholds here if partial scans are still too dark or too aggressive.

## Thumbnail auto-contrast
- **File:** `sxm_viewer/gui/thumbnail_render.py`
- **Function:** `robust_limits(arr, low_pct=2.0, high_pct=98.0)`
- **What it does:** Computes percentile-based `vmin`/`vmax` and ignores a single dominant flat bin **only if**:
  - That bin holds more than 80% of finite pixels, **and**
  - Trimming it still leaves at least 1% (or ≥10) pixels.
- **Tweak points:** The dominance threshold (`>0.8`) and minimum retained pixels (`max(10, 0.01 * data.size)`).

## Preview auto-contrast
- **File:** `sxm_viewer/gui/viewer/preview.py`
- **Function:** `_auto_preview_clim(arr)`
- **What it does:** Same relaxed dominant-bin rule as `robust_limits`, returning `(vmin, vmax)` or `None` if no span is found.
- **Where it is used:**
  - `show_file_channel(...)` sets `clim_main = _auto_preview_clim(display_arr)` and, if present, assigns `main['clim'] = clim_main`.
  - For extra views in `show_file_channel(...)`, `clim2 = _auto_preview_clim(arr2_display)` and, if present, `vdict['clim'] = clim2`.
  - These clims are applied by `MultiPreviewCanvas.set_views(...)`, so the canvas uses them automatically.
- **Tweak points:** Dominance threshold (`>0.8`) and retained pixel check (`max(10, 0.01 * finite.size)`), plus percentile range (1–99%).

## Related notes
- Auto-contrast is computed **after** unit scaling and adjustments. If you adjust these thresholds, keep parity between thumbnail and preview behavior.
- If partial scans still vanish, lower the dominance threshold (e.g., 0.7) or increase the minimum retained-pixel ratio. If blank regions bleed into contrast, raise the threshold (e.g., 0.9).
