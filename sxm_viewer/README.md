# SXM Viewer package (internals)

This package is what `python -m sxm_viewer` loads. The old `sxm_grid_viewer.py`
remains only as a shim and delegates into this package.

## Package map
```
sxm_viewer/
  __init__.py
  config.py                 # user config, cache limits, defaults
  palettes.py               # color cycles and colormap helpers
  data/
    io.py                   # Omicron/Anfatec header + channel parsing
    spectroscopy.py         # .dat metadata, axis helpers, matrix detection
    matrix.py               # matrix dataset representations
  providers/
    nanonis/
      adapter.py            # Nanonis .sxm -> Omicron-style cache generator
      vendor/               # vendored nanonispy reader
  gui/
    main_window.py          # top-level Qt widget and app state
    main_window_layout.py   # layout helpers and shortcuts panel
    main_window_toolbar.py  # toolbar actions and dark-mode toggle
    main_window_spectro.py  # spectro dock and browser wiring
    viewer/                 # thumbnail load/render, preview, loader, measurement
    spectroscopy/           # overlays, controller, popups for spectroscopies
    dialogs/                # spectroscopy dialogs, profile dialog, exports
    canvases/               # canvas workspace window and tiles
  utils/                    # small helpers (units, logging, thumbnails)
```

## Running from source
- After installing dependencies (see top-level README), launch with:
  ```bash
  python -m sxm_viewer
  ```
- For legacy compatibility, `python sxm_grid_viewer.py` still forwards to the
  package entry point.

## Migration status
- Core image browsing, spectroscopy overlays, matrix explorer, and canvas live
  here.
- Remaining legacy utilities are isolated under `scripts/` or kept as thin
  shims; new features should target the modules above.
