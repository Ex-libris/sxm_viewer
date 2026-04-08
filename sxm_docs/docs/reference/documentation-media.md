# Documentation Media

This page tracks which screenshots and GIFs are already available for the MkDocs site, and which new captures would improve the documentation further.

All media referenced by the docs should live under:

- `sxm_docs/docs/assets/screenshots/`

That keeps the files inside MkDocs' published `docs_dir`, so they are included in the built site.

Older or duplicate captures can be kept under:

- `sxm_docs/docs/assets/screenshots/legacy/`

---

## Current reusable media

These files are already available and ready to reference from the docs:

| File | Best use |
|---|---|
| `main-window-overview.png` | Home page / general orientation |
| `main-menu.png` | Top-level menu layout / preview header discussion |
| `thumbnail-to-popout.gif` | Basic browsing and opening pop-outs |
| `crop-to-measure-workflow.gif` | Crop-template workflow overview |
| `filters-pipeline.gif` | Image filter workflow |
| `histogram-live-contrast.gif` | Histogram and live contrast adjustment |
| `molecule-overlay-styling.gif` | Molecule overlay placement and appearance |
| `spectro-workflow.gif` | Spectroscopy popup workflow |
| `spectroscopies.png` | Spectroscopy browser / overview |
| `matrix-data.png` | Matrix spectroscopy page |
| `canvas-export-flow.gif` | Publication canvas and export |
| `wheel-zoom-pan-reset.gif` | Preview and pop-out navigation |
| `molecule-gizmo.png` | Static view of the molecule orientation gizmo |
| `molecule-gizmo-rotation.gif` | Interactive gizmo rotation workflow |
| `source-file-context-menu.png` | Source-file submenu reference |
| `spectroscopy-popup-current-layout.png` | Current single-spectrum popup layout |

---

## Current page coverage

These pages now reference in-site media directly from `assets/screenshots`:

| Page | Media now used |
|---|---|
| `index.md` | `main-window-overview.png` |
| `getting-started/first-steps.md` | `thumbnail-to-popout.gif` |
| `image-analysis/preview-and-popups.md` | `main-menu.png`, `wheel-zoom-pan-reset.gif`, `source-file-context-menu.png` |
| `image-analysis/cropping.md` | `crop-to-measure-workflow.gif` |
| `image-analysis/filters.md` | `filters-pipeline.gif` |
| `image-analysis/histogram.md` | `histogram-live-contrast.gif` |
| `image-analysis/molecules.md` | `molecule-overlay-styling.gif`, `molecule-gizmo.png`, `molecule-gizmo-rotation.gif` |
| `spectroscopy/overview.md` | `spectroscopies.png`, `spectroscopy-popup-current-layout.png`, `spectro-workflow.gif` |
| `spectroscopy/browser.md` | `spectroscopies.png` |
| `spectroscopy/matrix.md` | `matrix-data.png` |
| `workspace/canvas.md` | `canvas-export-flow.gif` |

This means MkDocs is now ready to publish the current screenshots and GIFs without depending on the top-level repository `screenshots/` folder.

---

## Existing gaps worth capturing

The current docs still describe some workflows that would benefit from newer screenshots or GIFs:

### High priority

- `spectroscopy-browser-table.png`
  - show the table-based browser rather than only the popup
- `preview-popout-layout.png`
  - preview and pop-out shown together, useful for `Preview & Pop-outs`

### Medium priority

- `molecule-reset-to-file-state.png`
  - context menu or right-click flow for reset
- `crop-template-row.png`
  - only needed if the current crop GIF no longer matches the crop-template toolbar
- `crop-template-edit-frame.png`
  - only needed if the current crop GIF does not show edit/rotate state clearly

---

## Capture order

If you only record a few new assets next, capture them in this order:

1. `spectroscopy-browser-table.png`
2. `preview-popout-layout.png`
3. `molecule-reset-to-file-state.png`

---

## Naming convention

Use lowercase, hyphenated names:

- `page-purpose.png`
- `feature-workflow.gif`

Examples:

- `crop-template-edit-frame.png`
- `molecule-gizmo-rotation.gif`
- `spectroscopy-popup-trace-controls.png`

---

## When to use PNG vs GIF

- Use **PNG** for layout, menus, and static state.
- Use **GIF** only when motion matters:
  - drag workflows
  - rotation
  - crop application
  - zoom/pan/reset

Keep GIFs short and tightly cropped. Large looping recordings make the docs feel heavy very quickly.
