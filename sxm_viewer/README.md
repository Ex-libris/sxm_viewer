# Refactored SXM Viewer

This directory contains a modular rewrite of `sxm_grid_viewer.py`.  The goal is to split
GUI, data-access, and processing logic into focused modules that are easier to test and
maintain.  The layout is intentionally simple so incremental migration from the legacy
script can happen feature-by-feature.

```
refactored_viewer/
+-- __init__.py
+-- __main__.py          # `python -m refactored_viewer` entry point
+-- data/
¦   +-- io.py            # Parsing of Omicron .txt headers and binary channels
+-- processing/
¦   +-- dataset.py       # Folder loader, tag detection, derived metadata
+-- gui/
¦   +-- main_window.py   # Qt widgets and interaction logic
+-- utils/
    +-- logging.py       # Thin helpers for progress + status reporting
```

Only a subset of the legacy features is implemented here (folder load, thumbnail list,
preview pane).  However, each module exposes small, composable classes so remaining
features (spectroscopy markers, filters, exports, etc.) can be migrated one at a time
without turning the GUI into another monolithic file.

## Migration strategy
1. **Data layer first** – keep header/binary parsing inside `data.io`.
2. **Processing/services** – move folder-level state (current files, tags, cached arrays)
   into plain Python classes inside `processing.dataset`.
3. **GUI** – connect widgets to the services via clean signals/slots.  Avoid having the
   widgets themselves manipulate files directly.

Running the refactored viewer:
```
python -m refactored_viewer --folder <path-to-sxm-folder>
```
