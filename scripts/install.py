#!/usr/bin/env python3
"""Bootstrap a dedicated virtual environment for the SXM viewer."""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import urllib.request
import tempfile
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
VENV_DIR = PROJECT_ROOT / ".venv"
REQUIREMENTS = PROJECT_ROOT / "requirements.txt"
MIN_PY = (3, 8)
MAX_PY = (3, 13)

# Fallback pip installer URL (works without SSL)
GET_PIP_URL = "https://bootstrap.pypa.io/get-pip.py"


def _run(cmd, check=True, **kwargs):
    print(f"[install] {' '.join(str(c) for c in cmd)}")
    if check:
        subprocess.check_call(cmd, **kwargs)
    else:
        return subprocess.run(cmd, **kwargs)


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


def check_ssl_support(py_path: Path) -> bool:
    """Check if Python has SSL support."""
    try:
        result = subprocess.run(
            [str(py_path), "-c", "import ssl"],
            capture_output=True,
            timeout=5
        )
        return result.returncode == 0
    except Exception:
        return False


def find_system_python() -> Path | None:
    """Find a working system Python installation with SSL support."""
    candidates = []
    
    if os.name == "nt":
        # Priority order for Windows
        local_appdata = os.environ.get("LocalAppData", os.path.expanduser("~\\AppData\\Local"))
        program_data = os.environ.get("ProgramData", "C:\\ProgramData")
        program_files = os.environ.get("ProgramFiles", "C:\\Program Files")
        
        # 1. Check python.org installations (most likely to have SSL)
        for version in ["312", "311", "310", "39"]:
            # Per-user install
            candidates.append(Path(local_appdata) / "Programs" / "Python" / f"Python{version}" / "python.exe")
            # All users install
            candidates.append(Path(program_data) / f"Python{version}" / "python.exe")
            candidates.append(Path(program_files) / f"Python{version}" / "python.exe")
            # Old-style C:\Python3X
            candidates.append(Path(f"C:\\Python{version}") / "python.exe")
        
        # 2. Try py launcher to find the default Python
        try:
            result = subprocess.run(
                ["py", "-3", "-c", "import sys; print(sys.executable)"],
                capture_output=True,
                text=True,
                timeout=5
            )
            if result.returncode == 0:
                py_path = Path(result.stdout.strip())
                if py_path.exists():
                    candidates.insert(0, py_path)
        except Exception:
            pass
        
        # 3. Check python in PATH
        python_exe = shutil.which("python")
        if python_exe:
            try:
                # Get actual path (not just "python")
                result = subprocess.run(
                    ["python", "-c", "import sys; print(sys.executable)"],
                    capture_output=True,
                    text=True,
                    timeout=5
                )
                if result.returncode == 0:
                    candidates.insert(0, Path(result.stdout.strip()))
            except Exception:
                candidates.append(Path(python_exe))
        
        # 4. Check Windows Store Python (usually works but check carefully)
        win_store_python = Path(local_appdata) / "Microsoft" / "WindowsApps" / "python.exe"
        if win_store_python.exists():
            candidates.append(win_store_python)
        
    else:
        # Unix-like systems
        for cmd in ["python3.12", "python3.11", "python3.10", "python3.9", "python3", "python"]:
            exe = shutil.which(cmd)
            if exe:
                candidates.append(Path(exe))
    
    # Test candidates for SSL support and correct version
    print("[install] Searching for Python with SSL support...")
    tested = set()
    for cand in candidates:
        try:
            # Avoid testing the same Python multiple times
            cand_resolved = cand.resolve()
            if cand_resolved in tested:
                continue
            tested.add(cand_resolved)
            
            if not cand.exists():
                continue
            
            # Check version first (faster than SSL check)
            major, minor = read_python_version(cand)
            if not supported_python_version((major, minor, 0)):
                continue
            
            # Check SSL support
            if check_ssl_support(cand):
                print(f"[install] Found working Python: {cand} (Python {major}.{minor})")
                return cand
            else:
                print(f"[install] Skipping {cand} (no SSL support)")
        except Exception as e:
            continue
    
    print("[install] No Python with SSL support found")
    return None


def pick_base_python(args):
    """Pick the best Python interpreter to use."""
    candidates: list[Path] = []
    
    # Priority 1: Explicit --python flag
    if args.python:
        candidates.append(Path(args.python))
    
    # Priority 2: PYTHON environment variable
    env_py = os.environ.get("PYTHON")
    if env_py:
        candidates.append(Path(env_py))
    
    # Priority 3: Current interpreter (but check SSL)
    candidates.append(Path(sys.executable))
    
    # Priority 4: Search for system Python with SSL
    system_py = find_system_python()
    if system_py:
        candidates.append(system_py)
    
    for cand in candidates:
        try:
            if cand.exists():
                return cand
        except Exception:
            continue
    
    return Path(sys.executable)


def read_python_version(py_path: Path) -> tuple[int, int]:
    cmd = [str(py_path), "-c", "import sys; print(f'{sys.version_info[0]}.{sys.version_info[1]}')"]
    out = subprocess.check_output(cmd, text=True, timeout=5).strip()
    parts = out.split(".")
    return int(parts[0]), int(parts[1])


def assert_supported_external_python(py_path: Path):
    major, minor = read_python_version(py_path)
    if not supported_python_version((major, minor, 0)):
        raise RuntimeError(
            f"Interpreter {py_path} reports Python {major}.{minor}, "
            f"but this project expects {MIN_PY[0]}.{MIN_PY[1]}--{MAX_PY[0]}.{MAX_PY[1]}. "
            "Pick a supported interpreter (set PYTHON or pass --python) and re-run with --reset if needed."
        )


def download_file(url: str, dest: Path) -> bool:
    """Download a file, trying both with and without SSL verification."""
    import ssl
    
    print(f"[install] Downloading {url}")
    
    # Try with SSL first
    try:
        urllib.request.urlretrieve(url, dest)
        return True
    except Exception as e:
        print(f"[install] SSL download failed: {e}")
    
    # Try without SSL verification
    try:
        print("[install] Retrying without SSL verification...")
        context = ssl._create_unverified_context()
        with urllib.request.urlopen(url, context=context) as response:
            with open(dest, 'wb') as out_file:
                out_file.write(response.read())
        return True
    except Exception as e:
        print(f"[install] Download failed: {e}")
        return False


def install_pip_manually(py: Path):
    """Download and install pip manually if it's missing or broken."""
    print("[install] Installing/upgrading pip manually...")
    
    with tempfile.TemporaryDirectory() as tmpdir:
        get_pip = Path(tmpdir) / "get-pip.py"
        
        if not download_file(GET_PIP_URL, get_pip):
            raise RuntimeError("Failed to download get-pip.py")
        
        # Install pip
        _run([str(py), str(get_pip), "--no-warn-script-location"])


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


def python_executable() -> Path:
    if os.name == "nt":
        return VENV_DIR / "Scripts" / "python.exe"
    return VENV_DIR / "bin" / "python"


def pip_executable() -> Path:
    if os.name == "nt":
        return VENV_DIR / "Scripts" / "pip.exe"
    return VENV_DIR / "bin" / "pip"


def test_pip_ssl(py: Path) -> bool:
    """Test if pip can access PyPI."""
    result = _run(
        [str(py), "-m", "pip", "search", "pip", "--index-url", "https://pypi.org/simple"],
        check=False,
        capture_output=True
    )
    return result.returncode == 0


def install_requirements():
    if not REQUIREMENTS.exists():
        raise FileNotFoundError("requirements.txt is missing; cannot install dependencies")
    
    py = python_executable()
    
    # Ensure pip is installed
    result = _run([str(py), "-m", "pip", "--version"], check=False, capture_output=True)
    if result.returncode != 0:
        print("[install] pip not found in venv, installing manually...")
        install_pip_manually(py)
    
    # Check if we have SSL issues
    has_ssl = check_ssl_support(py)
    
    if not has_ssl:
        print("\n" + "="*70)
        print("WARNING: Virtual environment has no SSL support!")
        print("This happens when the base Python lacks SSL libraries.")
        print("="*70)
        print("\nAttempting to install packages anyway...")
        print("If this fails, you'll need to:")
        print("1. Install Python from python.org (includes SSL)")
        print("2. Or fix your current Python's SSL support")
        print("="*70 + "\n")
    
    # Try to upgrade pip first
    print("[install] Upgrading pip...")
    result = _run(
        [str(py), "-m", "pip", "install", "--upgrade", "pip"],
        check=False,
        capture_output=True
    )
    
    if result.returncode != 0 and not has_ssl:
        # pip upgrade failed due to SSL, try manual installation
        print("[install] pip upgrade failed, installing manually...")
        install_pip_manually(py)
    
    # Install requirements
    print("[install] Installing requirements...")
    runner = [str(py), "-m", "pip", "install", "-r", str(REQUIREMENTS)]
    
    # If no SSL, try with trusted host flag
    if not has_ssl:
        runner.extend(["--trusted-host", "pypi.org", "--trusted-host", "files.pythonhosted.org"])
    
    result = _run(runner, check=False)
    
    if result.returncode != 0:
        print("\n" + "="*70)
        print("ERROR: Package installation failed!")
        print("="*70)
        
        if not has_ssl:
            print("\nYour Python installation lacks SSL support.")
            print("\nRECOMMENDED SOLUTIONS:")
            print("1. Download Python from https://www.python.org/downloads/")
            print("   (Make sure to check 'Add Python to PATH' during installation)")
            print("2. Set PYTHON environment variable to point to the new Python:")
            print("   set PYTHON=C:\\Python312\\python.exe")
            print("3. Re-run this installer with --reset")
            print("\nExample:")
            print("   set PYTHON=C:\\Python312\\python.exe")
            print("   python install.py --reset")
        else:
            print("\nTry running with --reset to start fresh:")
            print("   python install.py --reset")
        
        print("="*70 + "\n")
        raise RuntimeError("Installation failed")


def main():
    args = parse_args()
    assert_supported_runtime()
    
    base_python = pick_base_python(args)
    print(f"Using {base_python}")
    
    assert_supported_external_python(base_python)
    
    # Check SSL and warn user
    has_ssl = check_ssl_support(base_python)
    if not has_ssl:
        print("\n" + "="*70)
        print("ERROR: Selected Python lacks SSL support!")
        print("="*70)
        print(f"Python: {base_python}")
        print("\nThis Python cannot install packages from PyPI.")
        print("A virtual environment created from broken Python will also be broken.")
        print("\n" + "="*70)
        print("SOLUTIONS:")
        print("="*70)
        
        # Try to find alternative
        alt_python = find_system_python()
        if alt_python and alt_python != base_python:
            print(f"\nOption 1: Use alternative Python found on your system")
            print(f"  Found: {alt_python}")
            response = input(f"\nUse this Python instead? [Y/n]: ").strip().lower()
            if response != 'n':
                base_python = alt_python
                has_ssl = True  # We found working Python
                print(f"Switched to {base_python}\n")
            else:
                print("\nOption 2: Install Python from python.org")
                print("  1. Download from: https://www.python.org/downloads/")
                print("  2. Check 'Add Python to PATH' during installation")
                print("  3. Re-run this installer")
                print("\nOption 3: Fix Anaconda SSL")
                print("  Open Anaconda Prompt and run:")
                print("    conda install openssl certifi")
                print("  Then re-run this installer with --reset")
                print("="*70 + "\n")
                print("[install] Cannot continue with broken Python. Aborting.")
                sys.exit(1)
        else:
            print("\nOption 1: Install Python from python.org (RECOMMENDED)")
            print("  1. Download from: https://www.python.org/downloads/")
            print("  2. Check 'Add Python to PATH' during installation")
            print("  3. Re-run this installer")
            print("\nOption 2: Fix Anaconda SSL")
            print("  Open Anaconda Prompt and run:")
            print("    conda install openssl certifi")
            print("  Then re-run this installer with --reset")
            print("\nOption 3: Manually specify working Python")
            print("  set PYTHON=C:\\Python312\\python.exe")
            print("  install_sxm_viewer.bat --reset")
            print("="*70 + "\n")
            print("[install] Cannot continue with broken Python. Aborting.")
            sys.exit(1)
    
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
        
        print("\n" + "="*70)
        print("Installation complete!")
        print("="*70)
        print(f"\nActivate the environment with:")
        print(f"    {activate_cmd}")
        print(f"\nThen run:")
        print(f"    python -m sxm_viewer")
        print("\nOr use the run_sxm_viewer.bat shortcut")
        print("="*70 + "\n")


if __name__ == "__main__":
    main()
