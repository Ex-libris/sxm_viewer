# Export to SVG & PDF

SXM Viewer supports vector export for cleaner editing and print-oriented workflows.

---

## When to use vector export

Choose SVG or PDF when you want:

- scalable figure output
- post-editing in illustration tools
- cleaner print or manuscript figures

PNG is still the simplest choice for quick sharing, but vector export is often better for final figure preparation.

---

## Available routes

**Main preview / pop-outs** - the right-click menu has four distinct vector-ish actions:

- **Copy displayed as SVG** - the whole tiled view, including colorbar and labels
- **Copy data view as SVG (vector)** - the image data only, no colorbar/labels
- **Save displayed view as SVG...**
- **Save displayed view as PDF...**

Only the copy action has a data-only variant; saving to a file always includes the full displayed view.

**Publication Canvas** - vector export only exists via the right-click menu on empty canvas space, and only for **selected tiles** (there's no "export the whole canvas as SVG" action - see [Publication Canvas](../workspace/canvas.md#export)). With multiple tiles selected, copying composes them into a single multi-tile SVG document, while saving to a folder writes one SVG file per tile.

---

## Typical uses

| Format | Best for |
|---|---|
| SVG | Editing in tools such as Inkscape or Illustrator |
| PDF | Print-ready figures and archive output |

---

## Related pages

- [Copy & Export Images](export-images.md)
- [Export to PowerPoint](powerpoint.md)
- [Publication Canvas](../workspace/canvas.md)
