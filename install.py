#!/usr/bin/env python3
"""Bootstrap a dedicated virtual environment for the SXM viewer."""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
VENV_DIR = PROJECT_ROOT / ".venv"
REQUIREMENTS = PROJECT_ROOT / "requirements.txt"
MIN_PY = (3, 9)
MAX_PY = (3, 12)


def _run(cmd, **kwargs):
    print(f"[install] {' '.join(str(c) for c in cmd)}")
    subprocess.check_call(cmd, **kwargs)


def parse_args():
    parser = argparse.ArgumentParser(description="Install SXM Viewer dependencies.")
    parser.add_argument(
        "--reset",
        action="store_true",
        help="Recreate the virtual environment from scratch.",
    )
    parser.add_argument(
        "--python",
        help="Path to the Python interpreter to use (overrides PYTHON env).",
    )
    return parser.parse_args()


def supported_python_version(info):
    major, minor = info[:2]
    if (major, minor) < MIN_PY:
        return False
    if (major, minor) > MAX_PY:
        return False
    return True


def assert_supported_runtime():
    if not supported_python_version(sys.version_info):
        raise RuntimeError(
            f"Unsupported Python {sys.version_info.major}.{sys.version_info.minor}. "
            f"Please use a Python between {MIN_PY[0]}.{MIN_PY[1]} and {MAX_PY[0]}.{MAX_PY[1]}."
        )


def maybe_reexec_with_requested_python(requested: str | None):
    """
    If the user requested a specific interpreter (env PYTHON or --python) that differs
    from the current runtime, re-exec this installer with it. This keeps the .venv
    bound to the intended Python without requiring the caller to use that interpreter
    to launch install.py directly.
    """
    if not requested:
        return
    if os.environ.get("SXM_INSTALL_REEXEC"):
        return
    req_path = Path(requested)
    if not req_path.exists():
        return
    current = Path(sys.executable).resolve()
    if req_path.resolve() == current:
        return
    env = os.environ.copy()
    env["SXM_INSTALL_REEXEC"] = "1"
    print(f"[install] Re-running installer with PYTHON={req_path}")
    _run([str(req_path), __file__, *sys.argv[1:]], env=env)
    sys.exit(0)


def pick_base_python(args):
    # Priority: CLI flag, PYTHON env, active conda env, current interpreter.
    candidates: list[Path] = []
    if args.python:
        candidates.append(Path(args.python))
    env_py = os.environ.get("PYTHON")
    if env_py:
        candidates.append(Path(env_py))
    conda_prefix = os.environ.get("CONDA_PREFIX")
    if conda_prefix:
        if os.name == "nt":
            candidates.append(Path(conda_prefix) / "python.exe")
        else:
            candidates.append(Path(conda_prefix) / "bin" / "python")
    candidates.append(Path(sys.executable))
    for cand in candidates:
        try:
            if cand.exists():
                return cand
        except Exception:
            continue
    return Path(sys.executable)


def read_python_version(py_path: Path) -> tuple[int, int]:
    cmd = [str(py_path), "-c", "import sys; print(f'{sys.version_info[0]}.{sys.version_info[1]}')"]
    out = subprocess.check_output(cmd, text=True).strip()
    parts = out.split(".")
    return int(parts[0]), int(parts[1])


def assert_supported_external_python(py_path: Path):
    major, minor = read_python_version(py_path)
    if not supported_python_version((major, minor, 0)):
        raise RuntimeError(
            f"Interpreter {py_path} reports Python {major}.{minor}, "
            f"but this project expects {MIN_PY[0]}.{MIN_PY[1]}–{MAX_PY[0]}.{MAX_PY[1]}. "
            "Pick a supported interpreter (set PYTHON or pass --python) and re-run with --reset if needed."
        )


def ensure_venv(reset: bool, base_python: Path):
    if reset and VENV_DIR.exists():
        print(f"[install] Removing existing environment at {VENV_DIR}")
        shutil.rmtree(VENV_DIR, ignore_errors=True)
    if not VENV_DIR.exists():
        print(f"[install] Creating virtual environment in {VENV_DIR}")
        _run([str(base_python), "-m", "venv", str(VENV_DIR)])
    py = python_executable()
    if not py.exists():
        raise FileNotFoundError(
            f"Virtual environment is missing {py}. Re-run with --reset to recreate."
        )


def pip_executable() -> Path:
    if os.name == "nt":
        return VENV_DIR / "Scripts" / "pip.exe"
    return VENV_DIR / "bin" / "pip"


def python_executable() -> Path:
    if os.name == "nt":
        return VENV_DIR / "Scripts" / "python.exe"
    return VENV_DIR / "bin" / "python"


def ensure_pip(py: Path):
    pip = pip_executable()
    if pip.exists():
        return
    print("[install] ensurepip: restoring pip inside the virtual environment")
    _run([str(py), "-m", "ensurepip", "--upgrade"])


def install_requirements():
    if not REQUIREMENTS.exists():
        raise FileNotFoundError("requirements.txt is missing; cannot install dependencies")
    py = python_executable()
    ensure_pip(py)
    runner = [str(py), "-m", "pip"]
    _run(runner + ["--version"])
    _run(runner + ["install", "--upgrade", "pip"])
    _run(runner + ["install", "-r", str(REQUIREMENTS)])


def main():
    args = parse_args()
    requested = args.python or os.environ.get("PYTHON")
    maybe_reexec_with_requested_python(requested)
    assert_supported_runtime()
    base_python = pick_base_python(args)
    assert_supported_external_python(base_python)
    ensure_venv(reset=args.reset, base_python=base_python)
    assert_supported_external_python(python_executable())
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
