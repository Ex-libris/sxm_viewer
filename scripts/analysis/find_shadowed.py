"""A6 - shadowed definitions and broken module-level delegations.

Two defect classes that the other tools miss, both of which hide real
breakage behind Python's "last definition wins" rule:

1. **Shadowed methods** - the same method defined twice in one class body.
   The earlier one is dead. If the two bodies differ, the dead one is often
   an older version that no longer matches its callers' expectations, so
   the shadowing is silently load-bearing.

2. **Broken delegations** - a shim like
   ``return some_module.func(self, ...)`` where ``some_module`` does not
   define ``func``. These raise AttributeError if ever reached. Combined
   with (1) they are invisible: a broken shim that is also shadowed never
   runs, so nothing fails until someone deletes the shadowing definition.

Confirmed in this repo: ``SXMGridViewer._map_spec_to_pixels`` /
``_matrix_bbox_pixels`` / ``_fallback_spec_coords`` are defined twice, and
the *earlier* copies delegate to ``viewer_preview`` functions that do not
exist. Same family as the ``_show_spectro_popup`` bug (see CLAUDE.md).

Run:  python scripts/analysis/find_shadowed.py
"""
from __future__ import annotations

import ast
from collections import defaultdict

from common import Report, iter_source_files, parse, rel


def find_shadowed_methods():
    """Methods defined more than once in the same class body."""
    out = []
    for path in iter_source_files():
        tree = parse(path)
        if tree is None:
            continue
        for node in ast.walk(tree):
            if not isinstance(node, ast.ClassDef):
                continue
            seen = defaultdict(list)
            for child in node.body:
                if isinstance(child, (ast.FunctionDef, ast.AsyncFunctionDef)):
                    seen[child.name].append(child)
            for name, defs in seen.items():
                if len(defs) > 1:
                    out.append((rel(path), node.name, name,
                                [d.lineno for d in defs]))
    # Module-level duplicates too.
    for path in iter_source_files():
        tree = parse(path)
        if tree is None:
            continue
        seen = defaultdict(list)
        for child in tree.body:
            if isinstance(child, (ast.FunctionDef, ast.AsyncFunctionDef)):
                seen[child.name].append(child)
        for name, defs in seen.items():
            if len(defs) > 1:
                out.append((rel(path), "(module)", name,
                            [d.lineno for d in defs]))
    return out


def module_exports():
    """{resolved_module_path: set(top-level names)} for the package.

    Keyed by full path (not stem) so `data/io.py` cannot be confused with
    stdlib `io`.
    """
    exports = {}
    for path in iter_source_files():
        tree = parse(path)
        if tree is None:
            continue
        names = set()
        for child in tree.body:
            if isinstance(child, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
                names.add(child.name)
            elif isinstance(child, ast.Assign):
                for target in child.targets:
                    if isinstance(target, ast.Name):
                        names.add(target.id)
            elif isinstance(child, (ast.Import, ast.ImportFrom)):
                for alias in child.names:
                    names.add(alias.asname or alias.name.split(".")[0])
        exports[path.resolve()] = names
    return exports


def import_aliases(tree, path):
    """{local_alias: resolved_module_file} for relative imports that name a
    real module file in this package.

    Resolves ``from .viewer import preview as viewer_preview`` to the actual
    ``gui/viewer/preview.py``. Anything that does not resolve to a file on
    disk is skipped - notably ``from .._shared import io``, where ``io`` is
    a *re-exported stdlib module*, not our ``data/io.py``. Matching on bare
    stems instead flagged every ``io.BytesIO()`` in the codebase.
    """
    aliases = {}
    for node in ast.walk(tree):
        if not (isinstance(node, ast.ImportFrom) and (node.level or 0) > 0):
            continue
        base = path.resolve().parent
        for _ in range(node.level - 1):
            base = base.parent
        if node.module:
            for part in node.module.split("."):
                base = base / part
        for alias in node.names:
            candidate = base / f"{alias.name}.py"
            if candidate.exists():
                aliases[alias.asname or alias.name] = candidate.resolve()
    return aliases


def find_broken_delegations(exports):
    """`module.func(...)` where the resolved module defines no `func`."""
    broken = []
    for path in iter_source_files():
        tree = parse(path)
        if tree is None:
            continue
        aliases = import_aliases(tree, path)
        for node in ast.walk(tree):
            if not (isinstance(node, ast.Call)
                    and isinstance(node.func, ast.Attribute)
                    and isinstance(node.func.value, ast.Name)):
                continue
            local = node.func.value.id
            attr = node.func.attr
            target = aliases.get(local)
            if target is None or target not in exports:
                continue          # not one of our modules; can't judge
            if attr not in exports[target]:
                broken.append((rel(path), node.lineno, local,
                               rel(target), attr))
    return broken


def main():
    shadowed = find_shadowed_methods()
    exports = module_exports()
    broken = find_broken_delegations(exports)

    report = Report(
        "A6 - Shadowed definitions and broken delegations",
        "Defects hidden by Python's last-definition-wins rule. A shadowed "
        "method is dead code; a broken delegation raises AttributeError if "
        "reached. When both apply to the same name, nothing fails until "
        "someone removes the shadowing copy.")

    report.line(f"## Shadowed definitions ({len(shadowed)})\n")
    if shadowed:
        report.line("> The earlier definition never runs.\n")
        report.table(
            ("File", "Owner", "Name", "Lines"),
            [(f"`{f}`", owner, f"`{name}`",
              " , ".join(str(n) for n in lines))
             for f, owner, name, lines in sorted(shadowed)])
    else:
        report.line("None.\n")

    report.line(f"## Broken delegations ({len(broken)})\n")
    if broken:
        report.line("> `module.func(...)` where that module defines no "
                    "`func`. Each raises AttributeError if reached.\n")
        report.table(
            ("File", "Line", "Call"),
            [(f"`{f}`", line, f"`{local}.{attr}()` -> module `{stem}`")
             for f, line, local, stem, attr in sorted(broken)])
    else:
        report.line("None.\n")

    out = report.write("A6_SHADOWED.md")
    print(f"[A6] shadowed definitions: {len(shadowed)} | "
          f"broken delegations: {len(broken)} -> {out}")
    for f, owner, name, lines in shadowed:
        print(f"     SHADOWED {owner}.{name} at lines {lines} in {f}")
    for f, line, local, stem, attr in broken:
        print(f"     BROKEN   {f}:{line} {local}.{attr}() (module {stem})")


if __name__ == "__main__":
    main()
