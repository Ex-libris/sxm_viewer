# Installation

SXM Viewer is a Python desktop application. It runs directly from its repository folder and does not need to be installed as a system package.

!!! warning "Supported Python"
    Supported Python versions are 3.8 through 3.13.

    For Conda on Windows, Python 3.11 is the recommended choice.

    Do not use Python 3.14 yet.

!!! note "Where commands run"
    - `scripts/install.py` lives in `scripts/`
    - `scripts/requirements.txt` lives in `scripts/`
    - `python -m sxm_viewer` must be run from the repository root
    - Do not run `python -m sxm_viewer` from `scripts/`

It is strongly recommended to install SXM Viewer inside a dedicated Python environment such as Conda or `venv`. This keeps the Python version and package set isolated from the rest of your system.

??? note "What is a Python environment?"
    A Python environment is an isolated space with its own interpreter and installed libraries.

    In practice, that means:
    - Installing SXM Viewer does not affect other software using Python
    - Updates to other environments do not break SXM Viewer
    - Troubleshooting is easier because the app has a known Python setup

---

## Installation methods

| Method | Best for |
| --- | --- |
| Windows + Conda | New Windows users who want a clear, repeatable setup |
| Download ZIP | Quick setup without Git |
| Git clone | Keeping the checkout up to date |
| Project installer (`scripts/install.py`) | Users who want the repo to create and manage its own `.venv` |

---

## Option 1 - Windows + Conda or `venv` (recommended)

This is the clearest path for Windows users starting from zero.

Open a terminal in the repository root, then run:

```powershell
cd "path\to\sxm_viewer"
conda create -n sxmviewer python=3.11
conda activate sxmviewer
cd .\scripts
python -m pip install -r .\requirements.txt
cd ..
python -m sxm_viewer
```

If you are using `venv` instead of Conda, the same directory rules apply:

```powershell
cd "path\to\sxm_viewer"
python -m venv .venv
.\.venv\Scripts\Activate.ps1
cd .\scripts
python -m pip install -r .\requirements.txt
cd ..
python -m sxm_viewer
```

---

## Option 2 - Download ZIP

1. Go to https://github.com/Ex-libris/sxm_viewer
2. Click **Code -> Download ZIP**
3. Extract the archive

Then follow the Windows + Conda or `venv` steps above.

If you prefer the project-managed installer instead, open a terminal in `scripts/` and run:

```powershell
python install.py
```

That helper creates `scripts/.venv` and installs dependencies from `scripts/requirements.txt`.

---

## Option 3 - Clone with Git

If you use Git:

```powershell
git clone https://github.com/Ex-libris/sxm_viewer.git
```

Then follow the Windows + Conda or `venv` steps above.

This method makes updates straightforward:

```powershell
git pull
```

??? tip "What is Git?"
    Git is a version control system used to download and track changes in a project.

    Using Git makes it easier to update SXM Viewer without re-downloading the full archive.

??? tip "Installing Git"
    Git can be installed from https://git-scm.com/

    - Windows: use the official installer
    - macOS: install via Homebrew (`brew install git`) or Xcode Command Line Tools
    - Linux: install via your package manager, for example `apt install git` or `dnf install git`

    Verify installation with:

    ```bash
    git --version
    ```

---

## Project installer (`scripts/install.py`)

If you want SXM Viewer to manage its own local environment, use the helper in `scripts/install.py`.

Open a terminal in `scripts/` and run:

```powershell
python install.py
```

That script:

- creates `scripts/.venv` if needed
- installs the packages listed in `scripts/requirements.txt`
- checks that the selected Python version is supported

After it finishes, activate the new environment and launch from the repository root:

```powershell
.\.venv\Scripts\Activate.ps1
cd ..
python -m sxm_viewer
```

---

## First launch

From the repository root:

```powershell
python -m sxm_viewer
```

Then:

1. Click **Open folder**
2. Select a directory containing SXM data
3. Click a thumbnail to load a preview

See [Loading Data](../browsing/loading.md).

If you used the project installer, make sure the helper-created environment is active before launching.

---

## Updating the software

SXM Viewer is under active development. Git is the easiest way to keep a working checkout current.

### Using Git

Open a terminal in the repository root, then run:

```powershell
git pull
```

If you are using your own Conda or `venv` environment, reinstall the dependencies after pulling changes:

```powershell
cd scripts
python -m pip install -r .\requirements.txt
cd ..
```

If you are using the helper installer, rerun it from `scripts/`:

```powershell
python install.py
```

If needed, use `python install.py --reset` to recreate `scripts/.venv` from scratch.

??? tip "What is GitHub Desktop?"
    GitHub Desktop is a graphical interface for Git.

    It lets you clone, update, and manage repositories without using the command line.

### Using ZIP

Download a fresh archive and replace the existing folder. Local modifications are not preserved.

---

## When to use each method

Use Git, or GitHub Desktop, if:

- You want regular updates
- You are testing recent changes
- You plan to modify the code

Use the ZIP method if:

- You need a fixed snapshot
- You do not plan to update frequently

---

## Troubleshooting

### `No module named sxm_viewer`

You are running the launch command from the wrong folder.

Run `python -m sxm_viewer` from the repository root, not from `scripts/`.

If you activated Conda or `venv`, confirm that the same environment is still active:

```powershell
python -c "import sys; print(sys.executable)"
```

---

### `ModuleNotFoundError: No module named 'matplotlib'`

The active environment does not have the required packages installed.

From the repository root:

```powershell
cd scripts
python -m pip install -r .\requirements.txt
cd ..
```

Then run:

```powershell
python -m sxm_viewer
```

---

### `conda install requirements` fails

`requirements` is not a Conda package name.

Use the requirements file instead:

```powershell
cd scripts
python -m pip install -r .\requirements.txt
```

---

### `pip install .\requirements.txt` fails

That syntax is missing the `-r` flag.

Use:

```powershell
cd scripts
python -m pip install -r .\requirements.txt
```

---

### Python 3.14 is rejected

The installer only supports Python 3.8 through 3.13.

For Conda on Windows, use Python 3.11 unless you have a specific reason to choose another supported version.

---

### `python` or `conda` is not recognized

- On Windows, use the terminal that came with Conda, or run `conda init powershell` once and restart PowerShell.
- Check that Python is installed and available on `PATH`.

---

### The application still does not start

Check the Python interpreter and installed packages in the active environment:

```powershell
python --version
python -m pip --version
python -m pip list
```

If you used the helper installer, rerun it from `scripts/` with `--reset`.

---

## Minimal functional test

1. Launch the application
2. Open a data folder
3. Select a file in the thumbnail grid
4. Confirm that the preview and channel selector respond

Then proceed to [First Steps](first-steps.md).
