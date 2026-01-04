# SXM Grid Viewer

This viewer is tuned for data acquired with an Anfatec SXM controller running an Omicron Infinity microscope (tribus head, QPlus sensors at 8.6 K). It started life as a single monolithic script; it is now split into modules for the GUI, thumbnail grid, minimap, and helpers so it stays maintainable and easy to install.

## Screenshots (quick tour)

All images live in `screenshots/` so you can see the workflow before loading your own data.

![Main menu and grid navigation](screenshots/main_menu.png)
![Matrix and KPFM views](screenshots/matrix_data.png)
![Spectroscopy and parabolas](screenshots/spectroscopies.png)

## Why this exists

- Replaces a bulky one-off script with a repeatable install that lab PCs can keep in sync.
- Makes it fast to browse large folders, spot constant-height frames, and keep dz tags attached.
- Keeps spectroscopy, matrix scans, parabola fits, and WSxM XYZ exports in one place instead of scattered tools.
- Gives new users a predictable launcher and keeps the environment self contained in `.venv`.

## What it solves

- Thumbnail grid plus minimap to scan folders without blocking on full renders.
- Automatic detection of constant-height/current frames with dz preserved per file.
- Spectroscopy panel for single traces, matrix scans, parabola fits, and XYZ export (still being refined, so expect updates).
- Legacy `sxm_grid_viewer.py` remains as a shim; it now imports the package entry point.

## INSTALLATION : Choose **one** option below.

### Option A — Local self-contained Python (recommended)

This option creates a local Python environment inside the repo.
Nothing global is changed.

### What you start with
- You already have **some Python** on your system
- That Python is only used to start the installer

### What happens
1. You run:
   ```bash
   python install.py
   ```
2. That Python runs `install.py`
3. `install.py` uses `venv` to create:
   ```text
   .venv/
     bin/ (or Scripts/)
       python
   ```
4. The Python inside `.venv/` is:
   - the **same version** as the one you started with
   - placed in a new directory
5. All required packages are installed **only** into `.venv/`

### Run
No activation needed.

```bash
python -m sxm_viewer
```

Equivalent and explicit:

```bash
.venv/bin/python -m sxm_viewer      # macOS / Linux
.venv\Scripts\python -m sxm_viewer  # Windows
```

### What this option is NOT
- does not touch system Python
- does not create global executables
- does not use Conda

---

### Option B — Use your own existing environment

Use this if you already manage your Python env
(Conda, venv, system Python).

### Steps
```bash
git clone https://github.com/Ex-libris/sxm_viewer
cd sxm_viewer
python -m pip install -r requirements.txt
python -m sxm_viewer
```

Everything installs into **your active Python**.

---

### Option C — Conda environment

Explicit create → activate → install → run.

```bash
conda create -n sxm_viewer python=3.11
conda activate sxm_viewer
git clone https://github.com/Ex-libris/sxm_viewer
cd sxm_viewer
python -m pip install -r requirements.txt
python -m sxm_viewer
```

Leave later with:
```bash
conda deactivate
```

Do not use `install.py` if you choose this option.

---

### Option D — Windows double-click

No terminal required.

- `install_sxm_viewer.bat`  
  Builds `.venv/` and installs packages

- `run_sxm_viewer.bat`  
  Finds the local Python and starts the app

---

## Notes

- Re-run `python install.py` when dependencies change; add `--reset` if the existing `.venv` is broken.
- Set `PYTHON` (or pass `--python`) before running the installer to force a specific interpreter.
- Spectroscopy handling is under active improvement; workflows there may evolve.

## License

MIT License (see `LICENSE`).
