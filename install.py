#!/usr/bin/env python3
"""Bootstrap a dedicated virtual environment for the SXM viewer."""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
VENV_DIR = PROJECT_ROOT / ".venv"
REQUIREMENTS = PROJECT_ROOT / "requirements.txt"


def _run(cmd, **kwargs):
    print(f"[install] {' '.join(str(c) for c in cmd)}")
    subprocess.check_call(cmd, **kwargs)


def ensure_venv():
    if VENV_DIR.exists():
        return
    import venv

    print(f"[install] creating virtual environment in {VENV_DIR}")
    venv.EnvBuilder(with_pip=True, clear=False).create(str(VENV_DIR))


def pip_executable() -> Path:
    if os.name == "nt":
        return VENV_DIR / "Scripts" / "pip.exe"
    return VENV_DIR / "bin" / "pip"


def python_executable() -> Path:
    if os.name == "nt":
        return VENV_DIR / "Scripts" / "python.exe"
    return VENV_DIR / "bin" / "python"


def install_requirements():
    if not REQUIREMENTS.exists():
        raise FileNotFoundError("requirements.txt is missing; cannot install dependencies")
    py = python_executable()
    pip = pip_executable()
    runner = [str(py), "-m", "pip"] if py.exists() else [str(pip)]
    _run(runner + ["install", "--upgrade", "pip"])
    _run(runner + ["install", "-r", str(REQUIREMENTS)])


def main():
    ensure_venv()
    install_requirements()
    py = python_executable()
    if py.exists():
        if os.name == "nt":
            activate = VENV_DIR / "Scripts" / "activate"
            activate_cmd = f"{activate}"
        else:
            activate = VENV_DIR / "bin" / "activate"
            activate_cmd = f"source {activate}"
        print("[install] Done. Activate the environment with:\n" f"    {activate_cmd}")
        print("[install] then run:\n    python -m sxm_viewer.cli")


if __name__ == "__main__":
    main()
