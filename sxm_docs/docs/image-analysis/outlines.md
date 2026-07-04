# Outline Extraction

Outline extraction automatically traces the boundary of a prominent feature inside a region you select - it does not require manually tracing along an edge.

---

## Creating an outline

**Alt+drag** a rectangle on the preview or a pop-out around the feature you want outlined. The tool automatically finds the brightest/most prominent connected region inside that rectangle (using an intensity threshold with morphological clean-up) and extracts its boundary as a contour overlay.

Because the segmentation is automatic, the result depends on how tightly the rectangle is drawn around the feature - a tighter selection around just the feature of interest generally gives a cleaner outline than a loose one that includes surrounding background.

---

## Editing outlines

Right-click an existing outline for:

- **Change color...**
- **Line width** (several presets)
- **Line style** (solid, dashed, dotted, dense dash)
- **Undo last outline** - removes only the most recently added outline
- **Clear outlines** - removes every outline on the current view

Outline hit-testing is pixel-based and forgiving, so you don't need to click exactly on the thin contour line to select it.

!!! note "Two different undo mechanisms"
    **Undo last outline** (in the right-click menu) is a dedicated, outline-only undo stack. The general ++ctrl+z++ canvas undo also covers outline add/clear/style actions, but as part of the same shared stack used by filters, crops, and other edits - so Ctrl+Z may undo something else first if you've made other changes since adding the outline.

---

## Saving and restoring

Outline state is preserved through session save/load and virtual-copy workflows, the same as other canvas-side analysis state.

---

## Related tools

- [Cropping](cropping.md)
- [Profiles & Measurements](profiles.md)
- [Overlays](../workspace/overlays.md)
