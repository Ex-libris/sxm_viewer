# SXM Viewer interaction guide

This is a code-backed inventory of the user-facing controls currently wired into
the GUI. It is broader than a shortcut cheat sheet: it covers file loading,
mouse gestures, context menus, crop tools, spectroscopy, molecules, and the
publication canvas.

Scope note:
- This guide reflects the current implementation in the GUI code, not an ideal
  future design.
- Shortcuts are context-sensitive. The same key can do different things
  depending on whether focus is in thumbnails, a preview canvas, a popup, a
  spectroscopy dialog, or the canvas workspace.

Primary sources used for this guide:
- `sxm_viewer/gui/main_window.py`
- `sxm_viewer/gui/main_window_layout.py`
- `sxm_viewer/gui/main_window_toolbar.py`
- `sxm_viewer/gui/viewer/thumbnail_ui.py`
- `sxm_viewer/gui/controllers/thumbnail_controller.py`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py`
- `sxm_viewer/gui/controllers/quick_crop.py`
- `sxm_viewer/gui/controllers/histogram.py`
- `sxm_viewer/gui/dialogs/profile_dialog.py`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py`
- `sxm_viewer/gui/canvases/canvas_view.py`
- `sxm_viewer/gui/canvases/canvas_items.py`
- `sxm_viewer/gui/canvases/canvas_window.py`

## Supported inputs and outputs

Inputs:
- Data folders loaded from the `Open folder` button or by dragging folders/files
  from the OS into the main window.
- Omicron / Anfatec `.txt` images with multiple channels.
- Omicron `.dat` spectroscopy files, including single traces and matrix grids.
- Nanonis `.sxm` scans, converted on load into SXM Viewer cache folders.
- Molecule overlays loaded from `.xyz`, `.pdb`, or `.mol`.
- Saved session files loaded from JSON.

Outputs:
- PNG image export.
- SVG / PDF export from preview, popup, profile, spectroscopy, and canvas
  views.
- XYZ export.
- WSxM STP export.
- Animation export.
- Saved viewer sessions.
- Virtual thumbnail copies created from crops, popup views, or selected
  channels.

## Where features are discoverable today

Main discovery surfaces:
- Top toolbar: open folder, canvas, export, adjust image, spectro browser,
  shortcuts panel, display options, molecule loading, dark mode, colormap
  selectors.
- Display Options menu: spectroscopy markers, molecules, acquisition overlay,
  quick crop toggles, crop history/template visibility, profile label mode,
  detail theme, session save/load, arrange pop-outs.
- Thumbnail right-click menu: filters, export, adjust image, spectroscopy
  overlays, marker style, virtual copies, molecule actions, drift correction,
  animation.
- Preview / popup right-click menu: quick tools, display, overlays, analysis,
  copy / export, molecules, layout, typography.
- Plot right-click menus: copy / export plus typography and style controls.

Current UX limitation:
- Most advanced functions are still hidden in context menus. Tooltips exist, but
  the existing in-app shortcuts panel is only a compact hint, not a full guide.

## Global and main window shortcuts

Mode switching:
- `Ctrl+B` - switch to Browse mode.
- `Ctrl+M` - switch to Measure mode.
- `Ctrl+S` - switch to Spectroscopy mode.

Quick crop controls:
- `Ctrl+Shift+C` - toggle quick crop mode.
- `Ctrl+Shift+R` - reapply the current real-size crop template from the quick
  crop controls.
- `Ctrl+Shift+H` - toggle crop history overlay.
- `Ctrl+Shift+T` - toggle crop template overlay.
- `Ctrl+Shift+P` - bring all open pop-outs to the front and restore minimized
  ones.
- `Ctrl+Shift+W` - close the latest quick-crop pop-out.

General:
- `Ctrl+Z` - undo the focused canvas action; if no focused canvas action is
  available, fall back to quick-crop undo.
- `Ctrl+D` - duplicate the current preview into a popup.
- Arrow keys - move thumbnail selection if focus is not in a text-editing
  widget.

Drag and drop:
- Drag a local folder or file from the OS into the main window or thumbnail area
  to load it.

## Thumbnails and the thumbnail column

Basic selection:
- `Left click` - open the clicked thumbnail in the preview.
- `Double click` - open the preview as a popup.
- `Shift+Click` - range-select thumbnails.
- `Ctrl+Click` - toggle a thumbnail in the current multi-selection.
- `Ctrl+A` while focus is in the thumbnail area - select all visible thumbnails
  and spectroscopy miniatures.

Rubber-band selection:
- `Left drag` on empty thumbnail space draws a rubber-band selection box.
- Hold `Shift` or `Ctrl` while doing that to extend or modify an existing
  selection instead of replacing it.

Thumbnail actions:
- `Ctrl+Wheel` over thumbnails - resize thumbnail previews.
- `Ctrl+C` while focus is in the thumbnail area - export the selected thumbnails
  to temporary image files and place those file paths on the clipboard.
- `Left drag` a thumbnail into the publication canvas - add it as a canvas tile.
- Drag multiple selected thumbnails together into the canvas.

Thumbnail context menu:
- Apply a single filter or a custom pipeline to one or many selected files.
- Clear filters for one or many selected files.
- Copy selected thumbnails as SVG using the current rendered view.
- Export PNGs, XYZ, or WSxM STP.
- Open the image-adjust dialog.
- Toggle spectroscopy overlays, miniatures, glow, and marker style.
- Create virtual copies for the current or chosen channel.
- Remove selected virtual copies.
- Clear molecules or copy molecules from a source image into processed copies.
- Clear spectroscopy selections.
- Start drift-correction export.
- Create an animation from the current selection.

Thumbnail navigation:
- Arrow keys navigate the thumbnail grid in reading order.

## Spectroscopy thumbnails

Marker interactions:
- `Left click` a spectroscopy marker - open or select that spectroscopy entry.
- `Shift+Click` - multi-select spectroscopy entries for comparison.
- `Ctrl+Click` - toggle a spectroscopy entry in the current selection.
- Clicking a spectroscopy badge opens the summary for that image.

Context menu:
- Open spectroscopy popup.
- Show spectroscopy metadata in the details panel.
- Choose which channel is used for spectroscopy miniatures.
- Toggle spectroscopy miniatures.
- Copy the source file path.

## Minimap

Navigation:
- `Click` a frame - focus that frame.
- `Shift+Click` a frame - hide that entry.
- `Show all frames` restores hidden entries.
- `Mouse wheel` - zoom the minimap.
- `Left drag` on empty space - pan.
- `Middle drag` or `Right drag` - pan.

Other:
- Hover shows the filename tooltip.
- `Show real view` displays image pixmaps instead of flat frame polygons.
- Toolbar `Pop-ups`: click the main button to recall open pop-outs; use the
  arrow to focus a specific window, arrange/minimize/close all, or restore
  saved deferred pop-ups.

## Preview and popup canvases

Core navigation:
- `Mouse wheel` - zoom centered at the cursor.
- `Left drag` or `Middle drag` while already zoomed - pan the current axes.
- `R` - reset zoom.
- `Double click` a view - open it in a popup.

Clipboard and undo:
- `Ctrl+C` - copy the displayed canvas as PNG, including overlays.
- `Ctrl+Z` - undo the last canvas-side action when possible.

Overlay toggles:
- `Ctrl+1` - show / hide saved profile overlays.
- `Ctrl+2` - show / hide saved angle overlays.
- `Ctrl+3` - show / hide molecule overlays.
- `Ctrl+4` - show / hide scale bar.
- `Ctrl+5` - show / hide acquisition HUD.
- `Ctrl+H` - show / hide the in-canvas shortcut hint.

Contrast:
- `A` - auto contrast using the 1st to 99th percentile of the active image.

Popup / preview context menu:
- Quick tools: profile tool, angle tool, crop frame editor, apply crop,
  histogram, auto contrast, contrast reset, clear overlays.
- Display: presets, scale bar, ticks, colorbar, colorbar orientation, title,
  acquisition HUD, shortcut hint, frame fill, relative axes, grid / stacked
  layout.
- Overlays: saved profiles, saved angles, molecules.
- Analysis: filter menu and angle-style toggle for the active angle frame.
- Copy: displayed PNG / SVG or data-only PNG / SVG.
- Save / Export: PNG, SVG, PDF, WSxM STP.
- Virtual copy: create a virtual thumbnail copy from the current popup view.
- Molecules: load, load recent, clear.
- View: reset zoom, arrange pop-outs.
- Typography: plot font family and typography style controls.

Drag and drop:
- Drag a popup view into the thumbnail column to create a virtual copy at the
  drop position.

## Crop workflows

Manual drag crop:
- `Shift+drag` - create a rectangular crop selection.
- `Ctrl+Shift+drag` - create a square crop selection.
- Releasing the drag creates a cropped popup view.

Quick crop mode:
- Quick crop mode uses a reusable fixed-size template.
- `Click` the preview in quick crop mode to create a crop from the current
  template.
- The quick crop controls let you define real width / height, lock aspect ratio,
  and force square templates.

Crop frame editor:
- `Ctrl+E` or right-click -> `Quick tools -> Edit crop frame` enters crop frame
  transform mode.
- In transform mode:
  - `Drag frame body` - move the crop frame.
  - `Drag corner handles` - resize.
  - `Drag rotate handle` - rotate.
  - `Ctrl+drag` on the frame body - rotate without grabbing the rotate handle.
  - `Enter` - apply crop.
  - `Esc` - exit crop editor without applying.

Crop history:
- The crop history panel tracks recent crop overlays and pop-outs.
- `Select` highlights and focuses a crop; `Shift+click` supports multi-select.
- `Virtual copy` creates a thumbnail virtual copy from a stored crop snapshot.
- `Close view` closes the crop popup linked to that history entry.
- `Remove overlay` removes the crop overlay / history entry.
- `Ctrl+W` inside a quick-crop popup closes that popup.

## Measurement tools in preview / popup canvases

### Profile tool

Activation:
- `Ctrl+Click` on the main preview axes - enable the profile tool and start the
  first line immediately.
- The Measure-mode buttons and popup context menu can also enable or disable the
  profile tool.

Interaction:
- Drag endpoints to reshape the line.
- Drag the line body to move the whole profile.
- Clicking a saved profile overlay can promote it back to the active line.

Context actions:
- Right-click a profile overlay to change color, make it thicker / thinner,
  change label mode, or delete it.
- `Clear profile/angle overlays` removes active measurement overlays.

Related display:
- `Ctrl+1` toggles saved profile overlay visibility.

Profile dialog:
- `Delete` or `Backspace` - remove the selected overlay.
- `Ctrl+Wheel` - scale dialog fonts and labels.
- `V` - toggle draggable markers.
- `G` - toggle plot grid.
- `L` - toggle line rendering.
- `P` - toggle points.
- `M` - toggle multi-channel mode.
- `T` - toggle extra ticks.
- `R` - toggle precision mode.
- `A` - expand / collapse advanced options.
- Drag marker lines and marker arrows directly on the profile plot.
- Right-click the plot to copy PNG / SVG or adjust plot typography.

### Angle tool

Activation:
- `Ctrl+Alt+Click` on the main preview axes - enable the angle tool and insert a
  new angle frame at the click position.
- The Measure-mode buttons and popup context menu can also toggle the angle
  tool.

Interaction:
- Drag the vertex or arm handles to adjust the measurement.
- The active angle can be displayed as dots or arrowheads.

Related display:
- `Ctrl+2` toggles saved angle overlay visibility.

### Outline extraction

Interaction:
- `Alt+Left click` on an image - outline the dominant blob around the clicked
  point.
- `Middle click` does the same as a fallback gesture.
- Right-click an existing outline to open the outline menu.

Outline menu:
- Restyle existing outlines.
- Clear outlines.
- Undo recent outline additions.

### Filters on preview / popup canvases

Available from the `Analysis -> Filters` submenu:
- Flatten.
- Tilt correction.
- 2nd-order plane subtraction.
- Low-pass.
- High-pass.
- Laplacian.
- Custom pipeline.
- Clear filter.

Undo:
- `Ctrl+Z` on the focused preview / popup canvas can undo filter application.

## Histogram and contrast controls

Fast access:
- `A` - auto contrast (1 to 99 percent).
- Right-click -> `Quick tools -> Auto contrast`.
- Right-click -> `Quick tools -> Reset range to data min/max`.
- Right-click -> `Quick tools -> Histogram...`.

Histogram dialog:
- For multi-view popups, a view selector appears at the top.
- Drag the dashed min / max lines directly on the histogram.
- Or edit the numeric `Min` / `Max` spin boxes.
- `Auto (1-99%)` applies percentile-based range selection.
- `Reset` restores the data min / max.
- `Live preview` updates the image while dragging or editing.
- `Apply` confirms and closes the dialog.

Crop interaction:
- New cropped popups open with color limits derived from the cropped image
  itself, rather than inheriting the parent popup range.

## Molecule overlays

Loading:
- Toolbar molecule icon.
- Preview / popup context menu -> `Molecules -> Load Molecule`.
- Recent molecule list from the same menu.

Supported formats:
- `.xyz`
- `.pdb`
- `.mol`

Direct manipulation on the image:
- `Left drag` a molecule - translate it.
- `Shift+drag` - rotate in-plane (Z axis).
- `Middle drag` - rotate in 3D.
- `Ctrl+Shift+drag` - rotate in 3D.

Molecule context menu:
- Properties dialog for rotate / scale.
- Toggle shadows.
- Toggle hydrogens.
- Palette selection.
- Set atom color, per-element color, or bond color.
- Reset colors.
- Reset all molecules.
- `Ctrl+Z` - undo last molecule change.
- `Ctrl+D` - duplicate molecule.
- `Delete` - delete molecule.

## Spectroscopy tools

### Spectroscopy browser and overlays

Entry points:
- Bottom-bar `Spectro Browser` button.
- Toolbar `Spectro browser`.
- Spectroscopy mode controls.

Display toggles:
- Show spectroscopy overlays in thumbnails and preview.
- Show spectroscopy overlays only in preview.
- Toggle matrix markers, single markers, compact markers, marker style, glow,
  and miniatures.
- Treat NxN singles as matrix datasets.
- Force single mode.

### Single spectroscopy popup

General:
- Open from a marker click, spectroscopy lists, or context menus.
- Additional spectroscopy traces can be dragged onto an existing popup to stack
  them in the same window.

Popup interactions:
- Right-click plot for data copy, PNG / SVG copy, plot-style toggles, log axes,
  marker / line visibility, line width, legend, grid, position inset, and reset.
- Drag the position inset to reposition it.
- `Ctrl+Wheel` scales plot fonts.

### Spectroscopy comparison dialog

Shortcuts:
- `F` - fit selected spectra.
- `Ctrl+E` - export CSV.
- `Ctrl+A` - select all visible spectra.
- `Ctrl+Shift+A` - invert selection.
- `Delete` - clear selected spectra.
- `Ctrl+Delete` - clear all spectra.
- `Ctrl+Z` - undo the most recent comparison-side change.
- `Ctrl+Wheel` - scale plot fonts.

Mouse actions:
- `Shift+Click` two LCPD guide lines to draw a Delta LCPD annotation.
- Drag the position inset to move it.
- Drag minima markers and point-label annotations directly on the plot when
  those tools are active.

Context menus:
- Plot copy / export.
- Grid / legend / markers / lines toggles.
- Log axes.
- Line width control.
- Position inset toggle.
- Minima finding and overlap resolution.
- Add / clear point labels.
- Reset plot style.

## Publication canvas workspace

Entry:
- Open from the top toolbar `Open Canvas`.
- Or drag thumbnails into the canvas to create tiles.

Canvas view gestures:
- `Ctrl+Wheel` - zoom the canvas view.
- `Left drag` on empty canvas - rubber-band select items.
- `Left drag` on the background - pan.
- `Fit All in View` available from the canvas context menu.

Canvas keyboard controls:
- Arrow keys - nudge selected items by 1 px.
- `Shift+Arrow` - nudge selected items by 10 px.
- `Ctrl+A` - select all items.
- `Esc` - clear selection.
- `Delete` or `Backspace` - delete selected items.
- `Ctrl+Z` - undo.
- `Ctrl+Y` - redo.

Canvas tile interactions:
- Drag tiles to move them.
- Drag the lower-right resize handle to resize.
- `Alt+drag` a tile - duplicate it.
- Right-click a tile for duplicate, vector copy / save, stacking order, aspect
  lock, reset size, delete, and canvas-wide alignment / overlay actions.

Canvas empty-space context menu:
- Select all / deselect all.
- Copy or save selected items as SVG / PDF.
- Zoom in / out / reset.
- Fit all in view.
- Alignment, color-sync, overlay, grid, snap, layout, and canvas-color actions.

## Sessions, exports, and long-lived state

Saved sessions:
- `Display Options -> Save session...`
- `Display Options -> Load session...`
- Sessions preserve virtual copies, filters, profile and angle state, molecules,
  popup canvases, and the publication canvas state.

Export surfaces:
- Toolbar: export PNGs, export XYZ.
- Thumbnail context menu: PNG / XYZ / STP, animation, drift correction.
- Preview / popup context menu: PNG / SVG / PDF / STP.
- Profile dialog and spectroscopy dialogs: copy / export plots as PNG / SVG.
- Canvas workspace: copy / save vector outputs.

Virtual copies:
- Thumbnail context menu -> `Virtual copy`.
- Popup context menu -> `Create virtual copy in thumbnails`.
- Drag popup views into the thumbnail column.
- Quick-crop history -> `Virtual copy`.

## Summary: what is still hard to discover

What the codebase can do is broader than what the current GUI advertises:
- Many advanced actions are visible only through right-click menus.
- Typography controls exist in several plot/canvas menus but are not presented
  centrally.
- Crop, spectroscopy, molecules, and canvas editing each have their own local
  interaction model.
- The built-in shortcuts panel is useful for onboarding but too small to serve
  as a full manual.

Recommended next UX step:
- Add a dedicated in-app `Controls Guide` dialog that renders this document or a
  condensed version of it, so users do not need to inspect repository docs to
  discover the available workflows.
