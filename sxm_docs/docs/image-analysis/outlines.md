# Outline Extraction

Outline tools are used to define and edit extracted shape information directly on the image canvas.

---

## What this tool is for

Use outline extraction when you want a reusable contour-like overlay rather than a simple line or crop box.

Typical uses include:

- tracing a visible feature boundary
- refining a previously extracted outline
- preserving the result as part of the analysed image state

---

## Interaction model

Outlines are part of the canvas editing workflow, so they follow the same general principles as other analysis tools:

- they are edited on the preview or pop-out image canvas
- they participate in undo with ++ctrl+z++
- they can be preserved through saved state workflows

The project history also notes a dedicated fallback/undo path for outline edits, which is one reason they behave consistently with other canvas tools.

---

## Precision and selection

Recent changes improved outline hit-testing so right-click detection uses a more forgiving pixel-based test. That makes existing outlines easier to select and manipulate without requiring an exact click on a thin line.

---

## Saving and restoring

Outline state is intended to survive the same workflows as other canvas-side analysis data, including session save/load and virtual-copy style workflows when the full image state is preserved.

---

## Related tools

- [Cropping](cropping.md)
- [Profiles & Measurements](profiles.md)
- [Overlays](../workspace/overlays.md)
