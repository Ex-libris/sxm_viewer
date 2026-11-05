"""Import every module of the ``sxm_viewer`` package under offscreen Qt.

Catches broken imports in modules no smoke test reaches (a removed name
that another module still imports, a typo in a rarely used dialog, a
missing optional dependency that should have been guarded).

    python scripts/check_imports.py

Exit code 0 = every module imported.
"""
from __future__ import annotations

import importlib
import os
import sys
import traceback
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT))
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

# Not importable on purpose: the vendored reader (imported via the provider),
# a script meant to be run directly, and the `python -m` entry point.
SKIP_PARTS = {"vendor"}
SKIP_NAMES = {"smoke_profile_dialog", "__main__"}


def module_names() -> list[str]:
    names = []
    for path in sorted((REPO_ROOT / "sxm_viewer").rglob("*.py")):
        rel = path.relative_to(REPO_ROOT).with_suffix("")
        if SKIP_PARTS & set(rel.parts) or rel.name in SKIP_NAMES:
            continue
        parts = list(rel.parts)
        if parts[-1] == "__init__":
            parts = parts[:-1]
        names.append(".".join(parts))
    return names


def main() -> int:
    from PyQt5 import QtWidgets

    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])  # noqa: F841
    names = module_names()
    failed = []
    for name in names:
        try:
            importlib.import_module(name)
        except Exception:
            failed.append(name)
            traceback.print_exc()
    print(f"imported {len(names) - len(failed)}/{len(names)} modules", flush=True)
    for name in failed:
        print(f"FAILED  {name}", flush=True)
    return 1 if failed else 0


if __name__ == "__main__":
    code = main()
    sys.stdout.flush()
    os._exit(code)  # skip Qt teardown, which can crash offscreen and mask the exit code
