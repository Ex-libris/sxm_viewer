# Export to PowerPoint

SXM Viewer can send images directly into PowerPoint while preserving the source image proportions.

---

## Requirements

This is **live automation of an already-running PowerPoint**, not a file-based export - it does not create a new `.pptx` file. Before using it:

- PowerPoint must already be **open**, with a **presentation already open** in it.
- This feature is **Windows-only** and requires the `pywin32` package.

If either requirement isn't met, the PowerPoint menu actions are disabled with a tooltip explaining why, rather than failing silently.

---

## What is exported

PowerPoint sending is available from:

- the main preview and pop-out right-click menus ("Send to PowerPoint" / "Send to Current Slide")
- the thumbnail grid's right-click menu, for one or more selected thumbnails
- the Spectroscopy Summary dialog's preview, and the filter preview panel

!!! warning
    The Publication Canvas does **not** have a PowerPoint export action. If you've composed a figure there, export it as an image (see [Publication Canvas](../workspace/canvas.md#export)) and insert that into PowerPoint yourself.

Recent project changes fixed the export path so source aspect ratios are preserved instead of forcing every image into the same fixed rectangle - square scans stay square.

---

## Sending multiple images at once

Selecting several thumbnails and choosing **Send selected to Current Slide** automatically arranges them in a grid on the currently active slide (roughly square, based on how many images you selected), fitting each into its own cell with margins between them.

Choosing **Send selected to PowerPoint** (new slide) instead, or sending just a single image, creates one new slide per image rather than tiling them together.

---

## Notes on appearance

When sending exactly one image with a caption, the on-image title is hidden so it isn't duplicated by the caption in the final slide. The caption text itself is generated automatically from the image's own metadata - it isn't a free-text field you fill in beforehand.

---

## Related pages

- [Copy & Export Images](export-images.md)
- [Export to SVG & PDF](vector.md)
- [Publication Canvas](../workspace/canvas.md)
