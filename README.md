### 2026-07-04 UPDATES (performance pass)
A round of work aimed at the two things that show up most during everyday use: browsing thumbnails/previews and loading a folder for the first time.

- clicking between thumbnails and previews is noticeably snappier - removed several redundant full-canvas re-renders that used to happen on every click
- fixed a couple of display glitches introduced while chasing that speed-up (an accumulating scale bar, and image/axis misalignment when switching between very differently-sized scans)
- loading a folder you've never opened before is roughly 4-5x faster for Nanonis `.sxm` conversion, now parallelized across CPU cores with verified byte-identical output
- fixed a few accidentally-quadratic slow paths in spectroscopy loading (matching spectra to the right image, building per-spectrum metadata) - up to ~60x faster in the worst cases
- the on-disk header cache no longer grows without bound across every folder you've ever opened

### 2026-05-07 UPDATES (talking with Kelvin's group)
Nanonis support has been updated with a focus on faster scan loading and reloads:

- converted `.sxm` channel caches now use binary NumPy `.npy` files instead of ASCII text exports
- this reduces conversion I/O overhead and speeds up subsequent channel reads
- automatic CH/CC tag detection now reuses cached results when the header and topography source have not changed
- warm folder reloads therefore avoid unnecessary topography re-reads when auto-tagging is enabled

# SXM Viewer

SXM Viewer is a Python-based desktop application for scientific SPM (Scanning Probe Microscopy) data analysis and visualization, designed for Anfatec/Omicron systems. But also Nanonis. Maybe in the future Matrix. We will see.

---


## Documentation

Full documentation is available at:

https://ex-libris.github.io/sxm_viewer/

Key pages:
- Installation: https://ex-libris.github.io/sxm_viewer/getting-started/installation/
- First Steps: https://ex-libris.github.io/sxm_viewer/getting-started/first-steps/
- Profiles and Measurements: https://ex-libris.github.io/sxm_viewer/image-analysis/profiles/

---

## Overview

SXM Viewer provides an integrated environment for:

- Fast browsing of large SPM datasets
- Image analysis (profiles, angles, cropping, filtering)
- Spectroscopy visualization (traces, matrix scans, KPFM)
- Overlay tools (molecules, metadata, scale bars)
- Publication-ready figure composition (canvas)
- Session and collection management



![Main interface](screenshots/main_menu.png)

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

See the full installation guide in the MkDocs site for the Windows installer helper and troubleshooting notes.
