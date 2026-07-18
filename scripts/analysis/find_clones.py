"""A2 - AST structural clone detection.

Finds code that has the *same shape* even when identifiers, attribute names
and literals differ - the duplication a text search cannot see. Statement
sequences are normalized (names -> NAME, constants -> CONST, attribute
chains flattened), hashed, and grouped.

Two kinds of clone are reported:
  * whole-function clones - functions whose entire bodies normalize equal;
  * block clones - repeated statement runs of >= MIN_STMTS inside functions.

Run:  python scripts/analysis/find_clones.py
"""
from __future__ import annotations

import ast
import hashlib
from collections import defaultdict

from common import Report, iter_functions, iter_source_files, parse, rel

MIN_STMTS = 4          # shortest run worth reporting as a block clone
MIN_INSTANCES = 3      # a shape must recur at least this often
MAX_GROUPS = 40


# Normalization philosophy: erase identity (variable names, receivers,
# literal values) but KEEP called-method/attribute names, because those
# carry the semantics that make a reported clone actionable.

_SIG_CACHE: dict[int, str] = {}


def _signature(node) -> str:
    """Canonical normalized string for one AST node, computed directly.

    Deliberately *not* implemented as a NodeTransformer over a re-parsed
    copy: that required an ``ast.unparse`` + ``ast.parse`` round-trip per
    statement, which made a full scan of main_window.py (11.7k lines)
    effectively non-terminating. This walks the node once and memoizes per
    node id, so each statement is serialized exactly once no matter how
    many candidate runs contain it.
    """
    cached = _SIG_CACHE.get(id(node))
    if cached is not None:
        return cached
    if isinstance(node, ast.AST):
        if isinstance(node, ast.Name):
            out = "Name(NAME)"
        elif isinstance(node, ast.arg):
            out = "arg(ARG)"
        elif isinstance(node, ast.Constant):
            # Keep True/False/None - they carry control-flow meaning.
            out = (f"Const({node.value!r})"
                   if node.value in (True, False, None) else "Const(CONST)")
        elif isinstance(node, ast.Attribute):
            # Keep the attribute name (semantics), drop the receiver.
            out = f"Attr({node.attr},{_signature(node.value)})"
        else:
            parts = []
            for field, value in ast.iter_fields(node):
                if field in ("ctx", "type_comment", "lineno", "col_offset"):
                    continue
                parts.append(f"{field}={_signature(value)}")
            out = f"{type(node).__name__}({','.join(parts)})"
    elif isinstance(node, list):
        out = "[" + ",".join(_signature(n) for n in node) + "]"
    else:
        out = repr(node)
    _SIG_CACHE[id(node)] = out
    return out


def _normalize_dump(nodes):
    try:
        return "[" + ",".join(_signature(n) for n in nodes) + "]"
    except Exception:
        return None


def _hash(text):
    return hashlib.sha1(text.encode("utf-8")).hexdigest()[:12]


def find_function_clones():
    """Functions whose whole bodies are structurally identical."""
    groups = defaultdict(list)
    for fn in iter_functions():
        body = [s for s in fn.node.body
                if not (isinstance(s, ast.Expr) and isinstance(s.value, ast.Constant))]
        if len(body) < 2:
            continue
        dump = _normalize_dump(body)
        if dump is None:
            continue
        groups[_hash(dump)].append(fn)
    return {k: v for k, v in groups.items() if len(v) >= 2}


def find_block_clones():
    """Repeated statement runs inside function bodies."""
    groups = defaultdict(list)
    for path in iter_source_files():
        tree = parse(path)
        if tree is None:
            continue
        for node in ast.walk(tree):
            body = getattr(node, "body", None)
            if not isinstance(body, list) or len(body) < MIN_STMTS:
                continue
            for size in range(MIN_STMTS, min(len(body), 8) + 1):
                for start in range(0, len(body) - size + 1):
                    run = body[start:start + size]
                    dump = _normalize_dump(run)
                    if dump is None:
                        continue
                    groups[_hash(dump)].append(
                        (rel(path), run[0].lineno, size))
    return {k: v for k, v in groups.items() if len(v) >= MIN_INSTANCES}


def main():
    report = Report(
        "A2 - Structural clone detection",
        "Code with identical *shape* after normalizing away identifiers, "
        "receivers and literals. Catches duplication that renaming hides. "
        "Method-call names are deliberately preserved, so a group here is "
        "usually directly collapsible into a shared helper.")

    fn_groups = find_function_clones()
    report.line(f"## Whole-function clones ({len(fn_groups)} groups)\n")
    ranked = sorted(fn_groups.values(), key=lambda g: -len(g))
    for group in ranked[:MAX_GROUPS]:
        names = ", ".join(f"`{f.qualname}`" for f in group[:6])
        more = f" (+{len(group) - 6} more)" if len(group) > 6 else ""
        report.line(f"- **{len(group)}x** {names}{more}")
        report.line(f"  - first: `{group[0].file}:{group[0].lineno}` "
                    f"({group[0].n_lines} lines)")
    report.line()

    blk_groups = find_block_clones()
    ranked_blocks = sorted(blk_groups.values(),
                           key=lambda g: -(len(g) * g[0][2]))
    report.line(f"## Block clones ({len(blk_groups)} groups, "
                f">= {MIN_STMTS} statements, >= {MIN_INSTANCES} instances)\n")
    rows = []
    for group in ranked_blocks[:MAX_GROUPS]:
        files = len({g[0] for g in group})
        rows.append((len(group), group[0][2], len(group) * group[0][2],
                     files, f"`{group[0][0]}:{group[0][1]}`"))
    report.table(("Instances", "Stmts", "Total lines", "Files", "Example"), rows)

    for group in ranked_blocks[:15]:
        report.line(f"<details><summary>{len(group)}x {group[0][2]}-statement "
                    f"block (example {group[0][0]}:{group[0][1]})</summary>\n")
        for f, line, _size in sorted(group)[:25]:
            report.line(f"- `{f}:{line}`")
        report.line("\n</details>\n")

    out = report.write("A2_CLONES.md")
    fn_total = sum(len(g) for g in fn_groups.values())
    blk_total = sum(len(g) * g[0][2] for g in blk_groups.values())
    print(f"[A2] {len(fn_groups)} function-clone groups ({fn_total} functions)")
    print(f"[A2] {len(blk_groups)} block-clone groups "
          f"(~{blk_total} duplicated statements) -> {out}")


if __name__ == "__main__":
    main()
