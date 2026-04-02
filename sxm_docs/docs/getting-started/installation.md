# Installation

SXM Viewer is a Python desktop application. It runs directly from its folder and does not require installation as a system package.

It is strongly recommended to install it inside a dedicated Python environment (e.g. `venv` or Conda). This keeps the required Python version and dependencies isolated from other software on the system (for example system-wide Python installations or external tools such as Origin or Avogadro), and avoids version conflicts.

??? note "What is a Python environment?"
    A Python environment is an isolated space that contains its own Python interpreter and installed libraries.  
    This allows different projects to use different versions of packages without interfering with each other.  

    In practice, this means:
    - Installing SXM Viewer will not affect other software using Python  
    - Updates to other environments will not break SXM Viewer  
    - Reproducibility is improved when sharing setups across machines  

---

## Installation methods

| Method                      | Use case                                        |
| --------------------------- | ----------------------------------------------- |
| Download ZIP                | Quick setup without Git                         |
| Git clone                   | Recommended for keeping the software up to date |
| Conda / virtual environment | Isolated scientific environments                |

---

## Option 1 — Download ZIP

1. Go to: https://github.com/Ex-libris/sxm_viewer  
2. Click **Code → Download ZIP**  
3. Extract the archive  

### Open a terminal in the folder

Navigate into the extracted folder. You should see:

```
install.py
sxm_viewer/
README.md
```

**Windows:**
- Right-click inside the folder → **Open in Terminal** or **Open PowerShell window here**  
- Or click the path bar, type `cmd`, and press Enter  

**macOS:**
- Right-click → **New Terminal at Folder**  
- Or open Terminal and drag the folder into the window  

**Linux:**
- Right-click → **Open in Terminal**

---

### Install and run

```bash
python install.py
python -m sxm_viewer
```

---

## Option 2 — Clone with Git

If you use Git:

```bash
git clone https://github.com/Ex-libris/sxm_viewer.git
```

Then open a terminal inside the downloaded folder (as described above) and run:

```bash
python install.py
python -m sxm_viewer
```

This method allows straightforward updates:

```bash
git pull
```

!!! tip "What is Git?"
    Git is a version control system used to download and track changes in a project.  
    Using Git allows you to update SXM Viewer efficiently without re-downloading the entire project.

!!! tip "Installing Git"
    Git can be installed from: https://git-scm.com/

    - **Windows:** use the official installer  
    - **macOS:** install via Homebrew (`brew install git`) or Xcode Command Line Tools  
    - **Linux:** install via your package manager (e.g. `apt install git`, `dnf install git`)  

    Verify installation with:

    ```bash
    git --version
    ```

---

## Option 3 — Virtual environment (recommended)

Using a dedicated environment avoids conflicts with other Python installations.

### Using venv

```bash
python -m venv .venv
```

Activate the environment:

- **Windows:**
  ```bash
  .venv\Scripts\activate
  ```

- **macOS / Linux:**
  ```bash
  source .venv/bin/activate
  ```

Then install:

```bash
python install.py
```

---

### Using Conda

```bash
conda create -n sxmviewer python=3.11
conda activate sxmviewer
python install.py
```

---

## First launch

From the project folder:

```bash
python -m sxm_viewer
```

Then:

1. Click **Open folder**  
2. Select a directory containing SXM data  
3. Click a thumbnail to load a preview  

See [Loading Data](../browsing/loading.md).

---

## Updating the software

SXM Viewer is under active development. Using Git or GitHub Desktop is recommended if you intend to keep the software up to date.

### Using Git

Open a terminal in the repository folder (the folder containing `install.py`), then run:

```bash
git pull
python install.py
```

---

### Using GitHub Desktop

1. Install: https://desktop.github.com/  
2. Open GitHub Desktop  
3. Click **File → Clone repository**  
4. Select `Ex-libris/sxm_viewer` and choose a local folder  

To update:

- Click **Fetch origin → Pull**

---

### Running SXM Viewer after updating

After pulling updates, run the installer again.

Open a terminal in the repository folder (the folder containing `install.py`):

**Windows:**
- Open the folder in File Explorer  
- Right-click → **Open in Terminal** or **Open PowerShell window here**

**macOS:**
- Right-click → **New Terminal at Folder**  
- Or open Terminal and drag the folder into the window  

**Linux:**
- Right-click → **Open in Terminal**

Then run:

```bash
python install.py
```

If you are using a virtual environment, ensure it is activated before running this command.

!!! tip "What is GitHub Desktop?"
    GitHub Desktop is a graphical interface for Git.  
    It allows you to clone, update, and manage repositories without using the command line.

---

### Using ZIP

Download a fresh archive and replace the existing folder. Note that local modifications are not preserved.

---

## When to use each method

Use Git (or GitHub Desktop) if:

- You want regular updates  
- You are testing recent changes  
- You plan to modify the code  

Use the ZIP method if:

- You need a fixed snapshot  
- You do not plan to update frequently  

---

## Troubleshooting

### Python not found

Check:

```bash
python --version
```

If Python is not installed, download it from:  
https://www.python.org  

---

### Command not recognized

- **Windows:**
  ```bash
  py install.py
  ```

- **macOS / Linux:**
  ```bash
  python3 install.py
  ```

---

### Application does not start

Re-run:

```bash
python install.py
```

Ensure that the environment includes the required GUI and plotting dependencies.

---

## Minimal functional test

1. Launch the application  
2. Open a data folder  
3. Select a file in the thumbnail grid  
4. Confirm that the preview and channel selector respond  

Then proceed to [First Steps](first-steps.md).