# Copy & Export Images

## Copying a single image

- ++ctrl+c++ on the preview or a pop-out copies the **displayed view** (with all visible overlays) as a PNG to the clipboard.
- Right-click → **Copy** → **Copy as PNG** or **Copy as SVG** for explicit format choice.

!!! note
    The on-canvas shortcut hint text is **not** included in copied or exported images, even when visible on screen.

---

## Copying multiple thumbnails

1. Select thumbnails with Shift+Click, Ctrl+Click, or drag rubber-band selection.
2. Press ++ctrl+a++ to select all visible thumbnails.
3. Press ++ctrl+c++ to copy all selected images as separate PNG files.

Multi-image copy runs asynchronously so the UI stays responsive. A configurable cap (default 48 images) prevents memory issues on large selections. A non-blocking toast notification confirms completion.

---

## Saving a single image

Right-click → **Save / Export** → **Save as PNG** or **Save as SVG**.

---

## Exporting the preview with overlays

Right-click → **Copy displayed** captures the view exactly as it appears on screen, including profiles, angle overlays, molecules, scale bar, and acquisition HUD, but excluding the UI shortcut hint.

---

## Format notes

| Format | Best for |
|---|---|
| PNG | General use, presentations, quick sharing |
| SVG | Vector figures; editable in Inkscape/Illustrator |
| PDF | Print-ready single-page figures |

See [Export to SVG & PDF](vector.md) and [Export to PowerPoint](powerpoint.md) for those specific workflows.