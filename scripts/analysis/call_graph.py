"""A4 - usage and reachability analysis.

Three simplification levers, each derived from "who calls what":

  * **never called**  -> dead code; the safest possible simplification.
  * **called once**   -> inline candidates (especially 1-3 line shims).
  * **phantom calls** -> `self.x()` / `viewer.x()` where no `def x` exists
    anywhere. This is the `_show_spectro_popup` bug class: a `hasattr`
    guard that is always False, so a feature is silently dead. Worth
    checking on every run.

Name-based, not type-resolved: a method called only via ``getattr`` or a Qt
signal connection may look unused. Treat "never called" as a review queue,
not an automatic delete list - the report flags likely dynamic uses.

Run:  python scripts/analysis/call_graph.py
"""
from __future__ import annotations

import ast
from collections import defaultdict

from common import Report, iter_functions, iter_source_files, parse, read_text, rel

# Names that are called by the framework, not by our code.
QT_LIFECYCLE = {
    "__init__", "__repr__", "__str__", "__eq__", "__hash__", "__enter__",
    "__exit__", "__len__", "__iter__", "__getitem__", "__setitem__",
    "paintEvent", "resizeEvent", "showEvent", "hideEvent", "closeEvent",
    "keyPressEvent", "keyReleaseEvent", "mousePressEvent", "mouseMoveEvent",
    "mouseReleaseEvent", "mouseDoubleClickEvent", "wheelEvent", "moveEvent",
    "enterEvent", "leaveEvent", "focusInEvent", "focusOutEvent",
    "dragEnterEvent", "dragMoveEvent", "dragLeaveEvent", "dropEvent",
    "contextMenuEvent", "eventFilter", "event", "sizeHint",
    "minimumSizeHint", "paint", "data", "rowCount", "columnCount", "flags",
    "headerData", "accept", "reject", "done", "exec_", "run", "main",
    "setData", "index", "parent", "editorEvent", "createEditor",
}


def collect_definitions():
    defs = defaultdict(list)
    for fn in iter_functions():
        defs[fn.name].append(fn)
    return defs


def collect_calls():
    """{called_name: [(file, lineno), ...]} for attribute and bare calls."""
    calls = defaultdict(list)
    for path in iter_source_files():
        tree = parse(path)
        if tree is None:
            continue
        for node in ast.walk(tree):
            if isinstance(node, ast.Call):
                func = node.func
                if isinstance(func, ast.Attribute):
                    calls[func.attr].append((rel(path), node.lineno))
                elif isinstance(func, ast.Name):
                    calls[func.id].append((rel(path), node.lineno))
            # Qt signal wiring: .connect(self.handler) references without calling.
            elif isinstance(node, ast.Attribute):
                calls[node.attr].append((rel(path), node.lineno))
    return calls


def collect_dynamic_name_strings():
    """String literals that match method names - i.e. probable getattr/
    hasattr/signal-name dynamic use. Keeps the dead list honest."""
    names = set()
    for path in iter_source_files():
        tree = parse(path)
        if tree is None:
            continue
        for node in ast.walk(tree):
            if isinstance(node, ast.Constant) and isinstance(node.value, str):
                if node.value.isidentifier():
                    names.add(node.value)
    return names


def _inherited_api_names():
    """Method names provided by Qt / matplotlib base classes.

    Without this, every `self.setWindowTitle()` looks like a call to an
    undefined method. Introspects the actual installed PyQt5/matplotlib
    rather than hardcoding a list, and degrades to a static fallback when
    they are not importable.
    """
    names = set()
    try:
        from PyQt5 import QtWidgets, QtCore, QtGui
        for module in (QtWidgets, QtCore, QtGui):
            for cls_name in dir(module):
                cls = getattr(module, cls_name, None)
                if isinstance(cls, type):
                    names.update(dir(cls))
    except Exception:
        pass
    try:
        from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg
        from matplotlib.figure import Figure
        from matplotlib.axes import Axes
        for cls in (FigureCanvasQTAgg, Figure, Axes):
            names.update(dir(cls))
    except Exception:
        pass
    if not names:  # fallback so the tool still runs bare
        names = {"setWindowTitle", "resize", "update", "show", "hide", "close",
                 "scene", "rect", "setCursor", "setZValue", "mpl_connect"}
    return names


def collect_attribute_assignments():
    """Attribute names ever assigned anywhere (`x.attr = ...`).

    Needed to tell an *injected callback* apart from a genuine phantom:
    this codebase widely assigns callables onto objects from outside
    (`canvas._compare_menu_callback = fn`) and then calls them as
    `self._compare_menu_callback(...)`. Those have no `def`, but they are
    not dead - they are an implicit, untyped interface.
    """
    assigned = set()
    for path in iter_source_files():
        tree = parse(path)
        if tree is None:
            continue
        # Class-body `Name = value` bindings become class attributes, so
        # `self.Name` resolves. Missing these produced a false phantom for
        # the `SpectroSummaryDialog = SpectroSummaryDialog` import alias.
        for node in ast.walk(tree):
            if not isinstance(node, ast.ClassDef):
                continue
            for stmt in node.body:
                targets = (stmt.targets if isinstance(stmt, ast.Assign)
                           else [stmt.target] if isinstance(stmt, (ast.AnnAssign, ast.AugAssign))
                           else [])
                for target in targets:
                    if isinstance(target, ast.Name):
                        assigned.add(target.id)
        for node in ast.walk(tree):
            targets = []
            if isinstance(node, ast.Assign):
                targets = node.targets
            elif isinstance(node, (ast.AnnAssign, ast.AugAssign)):
                targets = [node.target]
            for target in targets:
                if isinstance(target, ast.Attribute):
                    assigned.add(target.attr)
                elif isinstance(target, (ast.Tuple, ast.List)):
                    for elt in target.elts:
                        if isinstance(elt, ast.Attribute):
                            assigned.add(elt.attr)
            # setattr(obj, "name", value)
            if (isinstance(node, ast.Call) and isinstance(node.func, ast.Name)
                    and node.func.id == "setattr" and len(node.args) >= 2
                    and isinstance(node.args[1], ast.Constant)
                    and isinstance(node.args[1].value, str)):
                assigned.add(node.args[1].value)
    return assigned


def find_phantom_calls(defs):
    """`self.foo()` / `viewer.foo()` where no `def foo` exists.

    Returns (true_phantoms, injected_callbacks):
      * **true phantom** - never defined AND never assigned. Dead code
        behind an always-False guard, or a latent AttributeError. This is
        the `_show_spectro_popup` bug class.
      * **injected callback** - never defined but assigned somewhere, i.e.
        a deliberate (if untyped) plug-in point. Not a bug, but an
        undocumented interface worth listing for the handover.
    """
    inherited = _inherited_api_names()
    assigned = collect_attribute_assignments()
    true_phantoms = defaultdict(list)
    injected = defaultdict(list)
    for path in iter_source_files():
        tree = parse(path)
        if tree is None:
            continue
        for node in ast.walk(tree):
            if not (isinstance(node, ast.Call)
                    and isinstance(node.func, ast.Attribute)):
                continue
            recv = node.func.value
            if not (isinstance(recv, ast.Name) and recv.id in ("self", "viewer")):
                continue
            name = node.func.attr
            if name in defs or name in QT_LIFECYCLE or name in inherited:
                continue
            bucket = injected if name in assigned else true_phantoms
            bucket[name].append((rel(path), node.lineno))
    return true_phantoms, injected


def main():
    defs = collect_definitions()
    calls = collect_calls()
    dynamic = collect_dynamic_name_strings()

    never, once = [], []
    for name, fns in defs.items():
        if name in QT_LIFECYCLE or name.startswith("__"):
            continue
        # A definition site also appears as a Name/Attribute in some ASTs;
        # count only references outside the definition's own line.
        sites = [c for c in calls.get(name, [])
                 if not any(c[0] == f.file and c[1] == f.lineno for f in fns)]
        if not sites:
            never.append((name, fns, name in dynamic))
        elif len(sites) == 1:
            once.append((name, fns, sites[0]))

    phantoms, injected = find_phantom_calls(defs)

    report = Report(
        "A4 - Usage and reachability",
        "Name-based call analysis. **Never called** is a review queue, not a "
        "delete list - Qt signal connections and `getattr` dispatch are "
        "invisible here (flagged where a matching string literal exists).")

    report.line("## True phantom calls (never defined, never assigned)\n")
    if phantoms:
        report.line("> Calls to something that exists nowhere - neither a "
                    "`def` nor an assignment. Each is dead code behind an "
                    "always-False guard, or a latent `AttributeError`. "
                    "Precedent: `_show_spectro_popup` was dead for the "
                    "Spectro Browser's entire lifetime this way.\n")
        rows = [(f"`{n}`", len(sites), f"`{sites[0][0]}:{sites[0][1]}`")
                for n, sites in sorted(phantoms.items(), key=lambda kv: -len(kv[1]))]
        report.table(("Name", "Call sites", "First site"), rows)
    else:
        report.line("None found.\n")

    report.line(f"## Injected callbacks ({len(injected)} names)\n")
    report.line("> Called as `self.x(...)` with no `def x`, but assigned "
                "somewhere as an attribute - a deliberate plug-in point. "
                "Not bugs, but an **implicit, untyped interface**: nothing "
                "declares them, so a typo or a rename fails silently at "
                "runtime. Worth documenting as a real extension surface.\n")
    rows = [(f"`{n}`", len(sites), f"`{sites[0][0]}:{sites[0][1]}`")
            for n, sites in sorted(injected.items(), key=lambda kv: -len(kv[1]))[:30]]
    report.table(("Callback", "Call sites", "First site"), rows)

    report.line(f"## Never referenced ({len(never)} definitions)\n")
    dyn = [n for n in never if n[2]]
    clean = [n for n in never if not n[2]]
    report.line(f"- {len(clean)} with no matching string literal "
                "(strongest delete candidates)")
    report.line(f"- {len(dyn)} whose name appears as a string somewhere "
                "(likely dynamic use - verify before touching)\n")
    rows = [(f"`{name}`", f"`{fns[0].file}:{fns[0].lineno}`", fns[0].n_lines,
             fns[0].classname or "(module)")
            for name, fns, _ in sorted(clean, key=lambda n: -n[1][0].n_lines)[:60]]
    report.table(("Name", "Location", "Lines", "Owner"), rows)

    report.line(f"## Called exactly once ({len(once)} definitions)\n")
    report.line("> Inline candidates - strongest for short shims.\n")
    short = [o for o in once if o[1][0].n_lines <= 4]
    report.line(f"{len(short)} of these are <= 4 lines.\n")
    rows = [(f"`{name}`", f"`{fns[0].file}:{fns[0].lineno}`", fns[0].n_lines,
             f"`{site[0]}:{site[1]}`")
            for name, fns, site in sorted(short, key=lambda o: o[1][0].file)[:60]]
    report.table(("Name", "Defined", "Lines", "Only caller"), rows)

    out = report.write("A4_USAGE.md")
    print(f"[A4] phantom calls: {len(phantoms)} | never referenced: "
          f"{len(never)} ({len(clean)} clean) | called once: {len(once)} "
          f"({len(short)} short) -> {out}")


if __name__ == "__main__":
    main()
