# SXM Viewer

**SXM Viewer** is a scientific SPM (Scanning Probe Microscopy) data analysis tool tuned for data acquired with an Anfatec SXM controller running an Omicron Infinity microscope (tribus head, QPlus sensors at 8.6 K). It also supports MATRIX vendor files and Nanonis formats.

![Main menu overview](assets/screenshots/main-menu.png){ width="1100" }


---

## What it does

- **Browse** large SPM data folders quickly via a thumbnail grid and minimap — no blocking on full renders.
- **Analyse** images with profiles, angle measurements, cropping, filters, and A/B comparison.
- **Visualise spectroscopy** — single traces, matrix scans, waterfall plots, and parabola fits — alongside your scan images.
- **Overlay** molecular models, scale bars, and acquisition metadata directly on images.
- **Export** publication-ready figures to PNG, SVG, PDF, and PowerPoint, or compose multi-image figure layouts in the built-in canvas.
- **Save sessions and collections** so you can return exactly where you left off, or curate a cross-folder set of key images.

---

## Quick start

```powershell
git clone https://github.com/Ex-libris/sxm_viewer.git
cd sxm_viewer
conda create -n sxmviewer python=3.11
conda activate sxmviewer
cd .\scripts
python -m pip install -r .\requirements.txt
cd ..
python -m sxm_viewer
```

If you prefer the project-managed installer instead of Conda, see [Installation](getting-started/installation.md). That page also explains the Windows launcher scripts and the supported Python versions.

---

## Navigation

Use the tabs above or the sidebar to find documentation by topic. If you are new, start with [First Steps](getting-started/first-steps.md). The [Keyboard Shortcuts](getting-started/shortcuts.md) reference is worth keeping open in a second tab.
