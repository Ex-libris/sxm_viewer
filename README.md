# SXM Viewer

Fast thumbnail browsing, spectroscopy overlays, and publication-ready canvases
for Omicron/Anfatec data. Nanonis `.sxm` files are supported through the bundled
adapter (they are converted into Omicron-style headers on load).

## Gallery (animated)
- Wheel zoom + pan + reset  
  ![](screenshots/Wheel-zoom%20+%20pan%20+%20reset.gif)
- Crop → pop-out → measure  
  ![](screenshots/Crop%20%E2%86%92%20pop-out%20%E2%86%92%20measure.gif)
- Thumbnail to pop-out  
  ![](screenshots/Thumbnail%20to%20pop-out.gif)
- Histogram live contrast  
  ![](screenshots/Histogram%20live%20contrast.gif)
- Filters pipeline  
  ![](screenshots/Filters%20pipeline.gif)
- Spectro workflow  
  ![](screenshots/Spectro%20workflow.gif)
- Molecular overlay styling  
  ![](screenshots/Molecular%20overlay%20styling.gif)
- Canvas export flow  
  ![](screenshots/Canvas%20export%20flow.gif)

Static reference:
![Main menu and grid navigation](screenshots/main_menu.png)
![Matrix and KPFM views](screenshots/matrix_data.png)
![Spectroscopy and parabolas](screenshots/spectroscopies.png)

## Quick start
- Prereqs: Python 3.8–3.12 and Git. (We auto-detect a working Python; use python.org builds for least friction.)
- Recommended (self-contained venv):
  ```bash
  python scripts/install.py
  python -m sxm_viewer
  ```
- Existing or Conda environment:
  ```bash
  pip install -r scripts/requirements.txt
  python -m sxm_viewer
  ```
- Windows, just clicking things:
  1) Click and Run `scripts\install_sxm_viewer.bat` (creates `.venv`, installs deps).
  2) Click and Run `scripts\run_sxm_viewer.bat` (launches using `.venv`, auto-loads `.env` if present).
     - Tip: edit `.env` after first launch to set paths/options; the runner will pick it up automatically when `python-dotenv` is installed (bundled in requirem

## Supported data and caches
- Anfatec/Omicron `.txt` images with multiple channels.
- Omicron `.dat` spectroscopies (single traces and matrix grids).
- Nanonis `.sxm` scans: converted on-the-fly under
  `<data folder>/.sxmviewer_nanonis/` (auto-regenerated when source files change).
- Keep personal datasets outside the repo; `.gitignore` excludes `.sxm`, `.dat`,
  `.sxmviewer_nanonis/`, and other caches.

## Core workflows
- **Browse datasets quickly**: responsive thumbnail grid with sorting/filtering,
  tagging, quick colormap swaps, and a minimap to stay oriented in large folders.
- **Inspect & measure**: wheel-zoom/pan, crop/zoom into pop-out windows, scale
  bars, line profiles, angle tool, and easy copy/export of what you see.
- **Spectroscopy & grids**: spectro browser/comparison for single traces and
  matrix grids; anchor grids to reference images; export fits/CSVs when needed.
- **Fix and enhance images**: histogram/contrast with draggable limits and live
  preview; plane/flatten/gamma/flip/rotate pipelines that work the same on
  thumbnails, preview, crops, and pop-outs.
- **Overlay molecules**: drop in XYZ/PDB/MOL, pick a style (ball & stick,
  sticks, wireframe, spacefill, licorice), palette, hydrogens/shadows, and align
  overlays directly on top of images.
- **Build figures**: publication canvas for arranging selected views; export
  PNGs/XYZ from toolbar groups (file/layout/export/tools/theme).

## Highlights
- Disk-backed caches for Nanonis conversions, headers, and spectroscopies for
  fast reloads.
- Double-click anywhere to pop out a full viewer with the same tools (measure,
  filters, overlays).
- Activity/status feedback: color-coded log with auto-scroll and a status bar
  showing current work and cache/selection info.
- Keyboard-friendly: navigation, layout toggle, exports, zoom/pan/reset, and
  measurement tools all have shortcuts (see [`docs/SHORTCUTS.md`](docs/SHORTCUTS.md)).

## Shortcuts and first steps
- Load a folder, pick a channel, and toggle **Show spectroscopies** to reveal
  markers. Shift+Click a marker to open comparison; right-click preview for copy
  options. Full list: [`docs/SHORTCUTS.md`](docs/SHORTCUTS.md).

## Troubleshooting
- Stale Nanonis cache: delete the generated `.sxmviewer_nanonis/` inside your
  data folder and reload.
- Config issues: use the in-app "Purge config" button or remove the config file
  under your user profile if asked by support.
- Run from a terminal to see log lines prefixed with `[SXMViewer]`.

## Repository layout
```
docs/         user/developer docs (structure, shortcuts, overviews)
scripts/      installers, helper launchers, legacy utilities (Python/BAT/MATLAB)
screenshots/  gallery shown above
samples/      optional reference datasets (keep personal data elsewhere)
sxm_viewer/   application package (GUI, data loaders, providers, utils)
```

Legacy `sxm_grid_viewer.py` remains as a shim and forwards to the packaged entry
point.

## License
MIT (see `LICENSE`).
