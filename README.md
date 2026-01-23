# SXM Viewer

Fast thumbnail browsing, spectroscopy overlays, and publication-ready canvases
for Omicron/Anfatec data. Nanonis `.sxm` files are supported through the bundled
adapter (they are converted into Omicron-style headers on load).

![Main menu and grid navigation](screenshots/main_menu.png)
![Matrix and KPFM views](screenshots/matrix_data.png)
![Spectroscopy and parabolas](screenshots/spectroscopies.png)

## Quick start
- Prereqs: Python 3.10+ and Git.
- Recommended (self-contained venv):
  ```bash
  python scripts/install.py
  python -m sxm_viewer
  ```
- Existing or Conda environment:
  ```bash
  pip install -r requirements.txt
  python -m sxm_viewer
  ```
- Windows double-click: run `scripts\install_sxm_viewer.bat`, then
  `scripts\run_sxm_viewer.bat`.

## Supported data and caches
- Anfatec/Omicron `.txt` images with multiple channels.
- Omicron `.dat` spectroscopies (single traces and matrix grids).
- Nanonis `.sxm` scans: converted on-the-fly under
  `<data folder>/.sxmviewer_nanonis/` (auto-regenerated when source files change).
- Keep personal datasets outside the repo; `.gitignore` excludes `.sxm`, `.dat`,
  `.sxmviewer_nanonis/`, and other caches.

## Current capabilities
- Thumbnail grid plus minimap with sort/filter, channel colormaps, and quick
  frame tagging.
- Preview panel with measurement mode (profiles), scale bars, SI/relative axes,
  and copy/export actions.
- Spectroscopy workflow: configurable markers, shift-click multi-selection,
  popups with log/lin axes, inset markers, parabola fits, and a comparison dialog
  with CSV export. Matrix datasets open in a dedicated explorer.
- Canvas workspace: drag thumbnails to build layouts, align/distribute/polish
  tiles, and export PNG/SVG snapshots for figures.

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
