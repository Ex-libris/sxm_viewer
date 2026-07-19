"""Census of delegating shims on a class.

A *shim* is a method whose entire body is one call that forwards to
somewhere else (``return mod.func(self, ...)`` /
``self.controller.do(...)``). Shims are how logic gets moved out of a god
class without breaking callers - but they keep the method count up and
leave the class as the discovery surface for everything.

This reports how much of a class is shim vs real logic, and groups shims
by the module they forward to, so a whole group can be retired at once by
pointing callers at the target module directly.

    python scripts/analysis/shim_census.py [--class SXMGridViewer]
"""
from __future__ import annotations

import argparse
import ast
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from common import Report, iter_source_files, parse, rel  # noqa: E402


def _delegation_target(node):
    """Module/attribute this single-statement body forwards to, or None."""
    body = [s for s in node.body
            if not (isinstance(s, ast.Expr) and isinstance(s.value, ast.Constant))]
    if len(body) != 1:
        return None
    stmt = body[0]
    if isinstance(stmt, ast.Return):
        call = stmt.value
    elif isinstance(stmt, ast.Expr):
        call = stmt.value
    else:
        return None
    if not isinstance(call, ast.Call) or not isinstance(call.func, ast.Attribute):
        return None
    receiver = call.func.value
    if isinstance(receiver, ast.Name):
        return receiver.id
    if isinstance(receiver, ast.Attribute):
        try:
            return ast.unparse(receiver)
        except Exception:
            return None
    return None


def find_reachability(names, owner_file):
    """Where each name is reachable from - all three routes that matter.

    Deleting a shim requires proving nothing reaches it, and there are
    three distinct ways it can be reached. Missing any of them deletes
    live code:

    1. ``viewer.X`` / ``self.viewer.X`` - the documented convention.
    2. ``self.X`` **in another module** - modules like
       ``main_window_state.py`` take the viewer as ``self`` so their code
       reads identically to when it lived in the class. Neither a
       ``viewer.X`` scan nor an "inside the owning file" scan sees these.
       This gap deleted a live shim on this branch; the smoke test caught
       it, static analysis did not.
    3. **String literals** - ``getattr(viewer, "X")`` dynamic dispatch,
       invisible to any attribute scan. ``session.py`` reaches
       ``_record_recent_session`` this way.

    Returns {name: {"attr": [files], "string": [files]}}.
    """
    hits = {n: {"attr": set(), "string": set()} for n in names}
    for path in iter_source_files():
        tree = parse(path)
        if tree is None:
            continue
        same_file = path.name == Path(owner_file).name
        for node in ast.walk(tree):
            if isinstance(node, ast.Attribute) and node.attr in names:
                recv = node.value
                is_viewer = isinstance(recv, ast.Name) and recv.id in ("viewer", "self")
                is_viewer = is_viewer or (isinstance(recv, ast.Attribute)
                                          and recv.attr == "viewer")
                if is_viewer and not same_file:
                    hits[node.attr]["attr"].add(path.name)
            elif (isinstance(node, ast.Constant)
                  and isinstance(node.value, str) and node.value in names):
                hits[node.value]["string"].add(path.name)
    return hits


def census(class_name):
    for path in iter_source_files():
        tree = parse(path)
        if tree is None:
            continue
        for node in ast.walk(tree):
            if not (isinstance(node, ast.ClassDef) and node.name == class_name):
                continue
            shims, real = [], []
            for child in node.body:
                if not isinstance(child, (ast.FunctionDef, ast.AsyncFunctionDef)):
                    continue
                lines = (getattr(child, "end_lineno", child.lineno)
                         - child.lineno + 1)
                target = _delegation_target(child)
                if target:
                    shims.append((child.name, target, child.lineno, lines))
                else:
                    real.append((child.name, child.lineno, lines))
            return rel(path), shims, real
    return None, [], []


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--class", dest="class_name", default="SXMGridViewer")
    args = ap.parse_args()

    file, shims, real = census(args.class_name)
    if file is None:
        print(f"class {args.class_name} not found")
        return 1

    total = len(shims) + len(real)
    by_target = defaultdict(list)
    for name, target, line, lines in shims:
        by_target[target].append((name, line, lines))

    report = Report(
        f"Shim census - {args.class_name}",
        "A shim is a method whose whole body forwards elsewhere. Shims let "
        "logic leave the class without breaking callers, but they keep the "
        "method count up and keep the class as the discovery surface. "
        "Retiring a group means pointing its callers at the target module "
        "directly and deleting the shims.")
    report.line(f"`{file}`\n")
    report.line(f"- total methods: **{total}**")
    report.line(f"- pure shims: **{len(shims)}** "
                f"({100.0 * len(shims) / max(1, total):.0f}%)")
    report.line(f"- real logic: **{len(real)}** "
                f"({sum(r[2] for r in real)} lines)\n")

    report.line("## Shims by forwarding target\n")
    rows = [(f"`{target}`", len(items),
             ", ".join(n for n, _l, _c in items[:5])
             + (f" +{len(items) - 5}" if len(items) > 5 else ""))
            for target, items in sorted(by_target.items(),
                                        key=lambda kv: -len(kv[1]))]
    report.table(("Target", "Shims", "Methods"), rows)

    # Which shims are reachable from outside, and by which route?
    shim_names = {n for n, _t, _l, _c in shims}
    reach = find_reachability(shim_names, file)
    unreachable = sorted(n for n in shim_names
                         if not reach[n]["attr"] and not reach[n]["string"])
    report.line(f"## Retirement candidates ({len(unreachable)})\n")
    report.line("> Shims with **no** external `viewer.X` / `self.X` access and "
                "**no** string reference. Still verify each against call sites "
                "inside the owning file, then delete. Run the smoke test after "
                "- static analysis has missed a live call site here before.\n")
    if unreachable:
        report.table(("Shim", "Forwards to"),
                     [(f"`{n}`", f"`{dict((a, b) for a, b, _l, _c in [(x[0], x[1], x[2], x[3]) for x in shims])[n]}`")
                      for n in unreachable])
    string_only = sorted(n for n in shim_names
                         if reach[n]["string"] and not reach[n]["attr"])
    if string_only:
        report.line("### Reached ONLY by string/getattr - never delete blindly\n")
        report.table(("Shim", "Referenced as a string in"),
                     [(f"`{n}`", ", ".join(sorted(reach[n]['string'])))
                      for n in string_only])

    report.line("## Largest remaining real-logic methods\n")
    real.sort(key=lambda r: -r[2])
    report.table(("Lines", "Line", "Method"),
                 [(c, l, f"`{n}`") for n, l, c in real[:40]])

    out = report.write(f"SHIM_CENSUS_{args.class_name}.md")
    print(f"{args.class_name}: {total} methods = {len(shims)} shims "
          f"+ {len(real)} real ({sum(r[2] for r in real)} lines)")
    print(f"\nTop forwarding targets:")
    for target, items in sorted(by_target.items(), key=lambda kv: -len(kv[1]))[:12]:
        print(f"  {len(items):>4}  {target}")
    print(f"\n-> {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
