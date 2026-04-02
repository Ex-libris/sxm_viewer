# Overlays

Overlays add analysis and contextual information directly on top of image views.

---

## Common overlay types

Depending on the current workflow, overlays can include:

- profiles
- angles
- molecules
- scale bars
- acquisition HUD / metadata overlays
- spectroscopy markers
- crop outlines

Some overlays are analysis tools, while others are display aids.

---

## Saved vs active overlays

For profile and angle workflows, SXM Viewer distinguishes between:

- **active** live measurements being edited now, and
- **saved overlays** that persist with the image state

The overlay toggle shortcuts refer to the saved overlays. Active measurements remain visible while you work.

| Shortcut | Effect |
|---|---|
| ++ctrl+1++ | Toggle saved profile overlays |
| ++ctrl+2++ | Toggle saved angle overlays |
| ++ctrl+3++ | Toggle molecule overlays |
| ++ctrl+4++ | Toggle scale bar |
| ++ctrl+5++ | Toggle acquisition HUD |

---

## Overlay sync across windows

Many display-state choices propagate between the main preview and open pop-outs, including profile/angle visibility, molecules, scale bar, acquisition overlay, title, and other display options.

---

## In sessions and collections

Overlay state is part of the saved workspace for sessions and is also preserved for collection items where per-item analysis context matters.

---

## Related pages

- [Profiles & Measurements](../image-analysis/profiles.md)
- [Angle Measurements](../image-analysis/angles.md)
- [Molecule Overlays](../image-analysis/molecules.md)
