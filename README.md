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

```bash
git clone https://github.com/Ex-libris/sxm_viewer.git
cd sxm_viewer
python install.py
python -m sxm_viewer
