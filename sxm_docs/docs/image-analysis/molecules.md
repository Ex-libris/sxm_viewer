# Molecule Overlays

Molecule overlays let you place and manipulate molecular models directly on preview canvases, pop-outs, and canvas tiles.

---

## Showing molecules

Molecule overlays can be toggled from display controls and right-click menus. They are part of the normal overlay system and can be shown or hidden without reloading the underlying image.

Saved molecule state is preserved with sessions and other workspace snapshots.

---

## Loading a molecule

Use the molecule controls in the GUI to load a molecular model onto the current image. The toolbar includes a dedicated molecule control, and the canvas also provides **Show**, **Load onto selected**, and **Clear from selected** actions.

On normal preview/pop-out canvases, molecule visibility is part of the shared display-state system.

---

## Rotating a selected molecule

Once a molecule is selected:

| Shortcut | Action |
|---|---|
| ++x++ | Rotate around X |
| ++y++ | Rotate around Y |
| ++z++ | Rotate around Z |
| ++shift+x++ / ++shift+y++ / ++shift+z++ | Rotate in the opposite direction |

The shortcut guidance appears in the same places where the app already surfaces key workflow hints.

---

## Overlay behavior

Molecule overlays are tied to the current image state rather than being a purely global decoration. In particular:

- preview and pop-outs can show or hide them independently through display state
- virtual copies default to their own molecule state
- sessions can preserve molecule overlays
- copy / clear actions exist for thumbnail and canvas workflows

---

## Default appearance

Recent project changes set new overlays to start in a bond-only display mode with the PyMol palette selected by default.

---

## Related pages

- [Overlays](../workspace/overlays.md)
- [Publication Canvas](../workspace/canvas.md)
- [Sessions & Collections](../browsing/sessions-and-collections.md)
