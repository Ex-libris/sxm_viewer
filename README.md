# SXM Grid Viewer

Fast installer, predictable launch, all the heavy analysis tools from the legacy script without the
monolithic pain.

## Install

1. Clone or unzip this repo. Work from the folder that contains `install.py`.
2. Run the installer:

   ```powershell
   python install.py
   ```

   - Prefer a double-click? Use `install_sxm_viewer.bat`—it calls the same script.
   - The installer creates `.venv` (or uses Conda) and drops all dependencies there.

3. Launch:

   ```powershell
   python -m sxm_viewer
   ```

   - Shortcut for lab PCs: `run_sxm_viewer.bat` uses the local env automatically.

## Why Use It

- Detects constant-height/current frames automatically, keeps dz tags with each file.
- Thumbnail grid + mini-map to navigate big folders without waiting on blocking previews.
- Spectroscopy panel handles single traces, matrix scans, parabola fits, and WSxM XYZ export.
- Legacy `sxm_grid_viewer.py` is still present; it just imports the package entrypoint.

## Notes

- Re-run `python install.py` any time you pull new dependencies.
- To force a specific interpreter, set `PYTHON` before running the installer.
- Alternate launch names: `python -m sxm_viewer.cli` or `python sxm_grid_viewer.py`.

## License

MIT License (see `LICENSE`).
