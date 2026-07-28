"""A1 - scan for known repeated idioms (the "we already know this is
duplicated" catalogue).

Detects each pattern **structurally** (via AST) rather than by regex, so
renamed variables and reformatting don't hide an instance. Every rule
reports file:line so the output doubles as a migration worklist.

Run:  python scripts/analysis/find_idioms.py
Also importable: ``scan_all()`` returns {rule_name: [Hit, ...]} so the
same rules can be used as a regression check after migration (see
Phase D / docs/refactor/PATTERNS.md).
"""
from __future__ import annotations

import ast
from collections import defaultdict
from dataclasses import dataclass

from common import Report, iter_source_files, parse, rel


@dataclass
class Hit:
    rule: str
    file: str
    lineno: int
    detail: str = ""


def _is_call_to(node, attr_name):
    """True when node is a call like `<something>.attr_name(...)`."""
    return (isinstance(node, ast.Call)
            and isinstance(node.func, ast.Attribute)
            and node.func.attr == attr_name)


def _expr_calls(stmt, attr_name):
    return (isinstance(stmt, ast.Expr) and _is_call_to(stmt.value, attr_name))


def _target(node):
    """Rough source-ish name of the object a method is called on."""
    try:
        return ast.unparse(node.func.value)
    except Exception:
        return "?"


# --------------------------------------------------------------------------
# Rules
# --------------------------------------------------------------------------

def rule_block_signals(tree, path, hits):
    """Every `blockSignals(True)` - i.e. every hand-rolled "update this
    widget without emitting signals" dance.

    Each such site implies the full triad, so counting the opens is the
    honest measure of the idiom (matching on strict statement adjacency
    undercounts: the guard is often wrapped in try/except, split across
    branches, or interleaved when several widgets are updated in a row).

    Classifies the restore style by inspecting the *enclosing function*:
    `blockSignals(prev)` is correct, a hardcoded `blockSignals(False)` is a
    latent bug - it unblocks signals a caller may have deliberately
    blocked, so nesting silently breaks.
    """
    # Map each function to the blockSignals calls inside it.
    for fn in ast.walk(tree):
        if not isinstance(fn, (ast.FunctionDef, ast.AsyncFunctionDef, ast.Module)):
            continue
        opens, restores_const_false, restores_dynamic = [], 0, 0
        for node in ast.walk(fn):
            if not _is_call_to(node, "blockSignals"):
                continue
            args = node.args
            if args and isinstance(args[0], ast.Constant):
                if args[0].value is True:
                    opens.append(node)
                elif args[0].value is False:
                    restores_const_false += 1
            else:
                # blockSignals(prev) / blockSignals(some_var)
                restores_dynamic += 1
        if not opens:
            continue
        # Attribute the restore style per function (dominant form).
        if restores_dynamic and not restores_const_false:
            style = "restores-prev"
        elif restores_const_false and not restores_dynamic:
            style = "hardcoded-False"
        elif restores_const_false or restores_dynamic:
            style = "mixed"
        else:
            style = "no-restore-found"
        for node in opens:
            hits.append(Hit("block_signals_triad", rel(path), node.lineno,
                            f"{_target(node)} ({style})"))


def rule_try_except_pass(tree, path, hits):
    """`try: ... except Exception: pass` - swallow-everything guards."""
    for node in ast.walk(tree):
        if not isinstance(node, ast.Try):
            continue
        for handler in node.handlers:
            is_broad = (handler.type is None
                        or (isinstance(handler.type, ast.Name)
                            and handler.type.id == "Exception"))
            body_is_pass = (len(handler.body) == 1
                            and isinstance(handler.body[0], ast.Pass))
            if is_broad and body_is_pass:
                n = len(node.body)
                hits.append(Hit("try_except_pass", rel(path), node.lineno,
                                f"guards {n} stmt(s)"))


def rule_defensive_getattr(tree, path, hits):
    """`getattr(self, "attr", None)` / `getattr(viewer, "attr", None)`.

    A symptom of attributes with no guaranteed initialization order.
    """
    for node in ast.walk(tree):
        if not (isinstance(node, ast.Call) and isinstance(node.func, ast.Name)
                and node.func.id == "getattr" and len(node.args) == 3):
            continue
        obj, name, default = node.args
        if not (isinstance(obj, ast.Name) and obj.id in ("self", "viewer")):
            continue
        if isinstance(name, ast.Constant) and isinstance(name.value, str):
            hits.append(Hit("defensive_getattr", rel(path), node.lineno,
                            f"{obj.id}.{name.value}"))


def rule_config_save(tree, path, hits):
    """`self.config[key] = value` immediately followed by `save_config(...)`."""
    for node in ast.walk(tree):
        body = getattr(node, "body", None)
        if not isinstance(body, list):
            continue
        for i, stmt in enumerate(body):
            if not (isinstance(stmt, ast.Assign) and stmt.targets
                    and isinstance(stmt.targets[0], ast.Subscript)):
                continue
            try:
                target_src = ast.unparse(stmt.targets[0].value)
            except Exception:
                continue
            if not target_src.endswith("config"):
                continue
            for follow in body[i + 1:i + 3]:
                if (isinstance(follow, ast.Expr)
                        and isinstance(follow.value, ast.Call)
                        and isinstance(follow.value.func, ast.Name)
                        and follow.value.func.id == "save_config"):
                    hits.append(Hit("config_write_then_save", rel(path),
                                    stmt.lineno, target_src))
                    break


def rule_pixel_extraction(tree, path, hits):
    """`int(header.get('xPixel', 128))` and friends - repeated header
    unpacking that belongs in one accessor."""
    for node in ast.walk(tree):
        if not (isinstance(node, ast.Call) and _is_call_to(node, "get")):
            continue
        if not node.args:
            continue
        first = node.args[0]
        if isinstance(first, ast.Constant) and first.value in (
                "xPixel", "yPixel", "XScanRange", "YScanRange"):
            hits.append(Hit("header_field_unpack", rel(path), node.lineno,
                            str(first.value)))


RULES = (
    ("block_signals_triad", rule_block_signals),
    ("try_except_pass", rule_try_except_pass),
    ("defensive_getattr", rule_defensive_getattr),
    ("config_write_then_save", rule_config_save),
    ("header_field_unpack", rule_pixel_extraction),
)

RULE_NOTES = {
    "block_signals_triad":
        "Replace with a `set_silent(widget, **props)` helper that restores the "
        "*previous* block state. Sites marked `hardcoded-False` also carry a "
        "latent nesting bug.",
    "try_except_pass":
        "Mostly guards hand-rolled widget pokes; many disappear for free once "
        "the other idioms are replaced by null-safe helpers. Do NOT bulk-delete "
        "- some are load-bearing (Qt teardown, optional imports).",
    "defensive_getattr":
        "Symptom of ~400 attributes with no guaranteed init order. Real fix is "
        "attribute initialization guarantees; deferred to the class-split work.",
    "config_write_then_save":
        "Collapse into a declarative settings registry (one row per setting).",
    "header_field_unpack":
        "Collapse into a single header accessor returning pixel/size fields.",
}


def scan_all():
    results = defaultdict(list)
    for path in iter_source_files():
        tree = parse(path)
        if tree is None:
            continue
        for name, rule in RULES:
            hits: list[Hit] = []
            rule(tree, path, hits)
            results[name].extend(hits)
    return results


def main():
    results = scan_all()
    report = Report(
        "A1 - Known idiom scan",
        "Structural (AST) scan for the repeated idioms identified in the "
        "duplication assessment. Re-run after each migration batch to confirm "
        "counts actually drop, and periodically to catch regressions.")

    report.line("## Summary\n")
    rows = []
    for name, _ in RULES:
        hits = results[name]
        files = len({h.file for h in hits})
        rows.append((name, len(hits), files))
    report.table(("Idiom", "Sites", "Files"), rows)

    for name, _ in RULES:
        hits = results[name]
        if not hits:
            continue
        report.line(f"## {name}  ({len(hits)} sites)\n")
        note = RULE_NOTES.get(name)
        if note:
            report.line(f"> {note}\n")
        by_file = defaultdict(list)
        for h in hits:
            by_file[h.file].append(h)
        rows = sorted(((f, len(hs)) for f, hs in by_file.items()),
                      key=lambda r: -r[1])
        report.table(("File", "Sites"), rows[:20])
        if name == "block_signals_triad":
            bad = [h for h in hits if "hardcoded-False" in h.detail]
            report.line(f"**{len(bad)} of {len(hits)} hardcode "
                        "`blockSignals(False)` instead of restoring the "
                        "previous state.**\n")
        report.line("<details><summary>All sites</summary>\n")
        for h in sorted(hits, key=lambda h: (h.file, h.lineno)):
            report.line(f"- `{h.file}:{h.lineno}` {h.detail}")
        report.line("\n</details>\n")

    out = report.write("A1_IDIOMS.md")
    total = sum(len(v) for v in results.values())
    print(f"[A1] {total} idiom sites across {len(RULES)} rules -> {out}")
    for name, _ in RULES:
        print(f"     {name:<26} {len(results[name])}")


if __name__ == "__main__":
    main()
