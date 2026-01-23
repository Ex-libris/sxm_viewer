# SXM Viewer keyboard & mouse reference

Centralized list of the gestures that the GUI already exposes. Everything here
maps directly to shortcuts that are wired in the codebase today.

## Global navigation
- `Ctrl+B`, `Ctrl+M`, `Ctrl+S` - jump between **Browse**, **Measure** and
  **Spectro** preview modes (see `SXMGridViewer._init_mode_shortcuts`).
- `Ctrl+C` while the preview canvas has focus - copy the rendered image to the
  clipboard (`MultiPreviewCanvas` copy handler).
- Drag folders from the OS explorer onto the main window to load them.

## Thumbnail grid & minimap
- Grid selection: `Shift+Click` or `Ctrl+Click` adds thumbnails to the current
  selection; `Ctrl+Wheel` resizes thumbnails; `Ctrl+Drag` reorders the export
  selection list.
- Thumbnail actions: right-click a frame to open filters/export commands; drag
  selected thumbnails directly into the Canvas window to build a layout.
- Minimap (Folder layout panel):
  - `Click` focuses a frame.
  - `Shift+Click` hides a frame; **Show all frames** restores every entry.
  - Mouse wheel zooms; click-drag pans the minimap view.
  - `Show real view` toggles channel overlays on the minimap itself.

## Preview & measurement panel
- Right-click inside the preview canvas for copy/save/export actions.
- `Ctrl+C` duplicates the preview image (also listed in the global shortcuts).
- **Measure** mode:
  - Drag in the preview to define a profile line.
  - The profile dialog lists overlays; `Delete` or `Backspace` removes the
    selected overlay; `Ctrl+Wheel` scales every font/label in that dialog.
  - Right-click the profile plot to copy it as PNG or SVG.

## Spectroscopy markers & thumbnails
- `Shift+Click` a spectroscopy marker (thumbnail or preview) to multi-select
  traces for the Spectroscopy Comparison window.
- Drag spectroscopy markers onto an existing spectroscopy popup to stack that
  trace with the currently plotted ones.
- `Ctrl+Wheel` over thumbnail markers still resizes the underlying thumbnails,
  keeping marker sizes proportional if *Compact markers* is enabled.

## Spectroscopy popup window
- Channel selector toolbar: **Fit parabola** fits the active curve, **Copy
  channel** pushes the current dataset to the clipboard.
- Canvas context menu -> **Plot style** toggles grid/legend, switches to log
  axes, hides the connecting line, or switches to markers-only plots. Line-width
  presets and incremental increase/decrease commands live in this submenu.
- **Position inset** toggle adds the miniature thumbnail with marker locations;
  drag the inset to reposition it on the figure.
- `Ctrl+Wheel` anywhere on the dialog scales all fonts (axes, legends, inset
  labels) between 60% and 250%.
- Drag/drop additional spectroscopy markers from thumbnails, the preview, or the
  comparison list to overlay more curves in the same popup.

## Spectroscopy comparison dialog
- Checklist view shortcuts (defined via `QtWidgets.QShortcut`):
  - `F` - fit the selected spectra.
  - `Ctrl+E` - export the current table to CSV.
  - `Ctrl+A` - select all visible spectra.
  - `Ctrl+Shift+A` - invert the selection of all visible spectra.
  - `Delete` - clear selected spectra; `Ctrl+Delete` clears the entire list.
- Plot interactions:
  - `Shift+Click` two LCPD guide lines to annotate the Delta LCPD distance.
  - Waterfall/points/lines toggles control how overlapping traces are drawn.
  - Right-click the canvas for copy/export options plus the same style toggles
    available in single-popup plots.

## Canvas workspace
- Drag thumbnails (single or multi-select) into the canvas window to add them as
  tiles. The canvas exposes Align/Distribute/Polish buttons; they operate on
  whichever tiles are selected.
- Canvas toolbar buttons (Align Top/Bottom/Left/Right, Distribute, Normalize
  size) rescale selected tiles to match the first tile in the selection, keeping
  the grid tidy.
- Right-click a tile for quick exports (PNG/SVG) or to toggle shared colorbars.

> **Tip:** most dialogs remember their screen position. Move them once and
> future openings will stagger so they no longer block the thumbnail region.
