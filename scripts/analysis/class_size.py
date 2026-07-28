"""Track the size of the classes being decomposed.

    python scripts/analysis/class_size.py

Prints method/attribute counts for the god classes and, when
``class_size_baseline.json`` exists, the delta since it was written.
Used as the progress metric for docs/refactor/GOD_CLASS_PLAN.md.
"""
from __future__ import annotations

import ast
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from common import iter_source_files, parse, rel  # noqa: E402

TARGETS = ("SXMGridViewer", "MultiPreviewCanvas", "MatrixSpectroViewer")
BASELINE = Path(__file__).parent / "class_size_baseline.json"


def measure():
    out = {}
    for path in iter_source_files():
        tree = parse(path)
        if tree is None:
            continue
        for node in ast.walk(tree):
            if not (isinstance(node, ast.ClassDef) and node.name in TARGETS):
                continue
            methods = [c.name for c in node.body
                       if isinstance(c, (ast.FunctionDef, ast.AsyncFunctionDef))]
            attrs = {n.attr for n in ast.walk(node)
                     if isinstance(n, ast.Attribute)
                     and isinstance(n.value, ast.Name)
                     and n.value.id == "self"}
            lines = (getattr(node, "end_lineno", node.lineno) - node.lineno + 1)
            out[node.name] = {"methods": len(methods),
                              "attributes": len(attrs),
                              "lines": lines,
                              "file": rel(path)}
    return out


def main():
    now = measure()
    base = {}
    if BASELINE.exists():
        base = json.loads(BASELINE.read_text(encoding="utf-8"))

    if "--update" in sys.argv:
        BASELINE.write_text(json.dumps(now, indent=2, sort_keys=True),
                            encoding="utf-8")
        print(f"baseline written to {BASELINE.name}")

    print(f"{'class':<22} {'methods':>9} {'attrs':>8} {'lines':>8}")
    print("-" * 50)
    for name in TARGETS:
        cur = now.get(name)
        if not cur:
            continue
        old = base.get(name)
        if old:
            print(f"{name:<22} {cur['methods']:>9} {cur['attributes']:>8} "
                  f"{cur['lines']:>8}")
            print(f"{'  vs baseline':<22} "
                  f"{cur['methods'] - old['methods']:>+9} "
                  f"{cur['attributes'] - old['attributes']:>+8} "
                  f"{cur['lines'] - old['lines']:>+8}")
        else:
            print(f"{name:<22} {cur['methods']:>9} {cur['attributes']:>8} "
                  f"{cur['lines']:>8}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
