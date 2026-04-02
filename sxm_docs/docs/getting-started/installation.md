# Installation

SXM Viewer is a Python desktop application. The simplest install path is to clone the repository, install the dependencies, and launch the package from the repo root.

---

## Recommended install

```bash
git clone https://github.com/Ex-libris/sxm_viewer.git
cd sxm_viewer
python install.py
```

If you prefer, you can also install dependencies manually in your own environment and then run the package directly.

```bash
python -m sxm_viewer
```

---

## Choosing a Python environment

A dedicated virtual environment is recommended so plotting, Qt, and scientific packages do not conflict with other projects.

Typical workflow:

```bash
python -m venv .venv
.venv\Scripts\activate
python install.py
```

If you already use Conda or another environment manager, install into that environment instead.

!!! tip
    The project history explicitly notes support for Python 3.13, so newer Python environments are expected to work as well.

---

## First launch

From the repository root:

```bash
python -m sxm_viewer
```

On first launch, open a data folder with **Open folder** from the toolbar. See [Loading Data](../browsing/loading.md).

---

## Common launch options

| Task | How |
|---|---|
| Launch from source | `python -m sxm_viewer` |
| Reinstall after updates | rerun `python install.py` |
| Use a specific interpreter | run the command with that interpreter explicitly |
| Keep docs and code together | run from the repo root |

---

## Troubleshooting

### The app does not start

Check that the environment has the GUI and plotting dependencies installed, then rerun the installer.

### A file opens but displays strangely

That is usually a data-format or channel issue rather than an installation issue. See [Supported File Formats](../reference/file-formats.md).

### I want a minimal smoke test

Launch the app, open a folder, click a thumbnail, and confirm that the preview, channel selector, and thumbnail grid all respond. Then open the [First Steps](first-steps.md) page and walk through the basic workflow.
