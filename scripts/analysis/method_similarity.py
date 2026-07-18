"""A3 - near-duplicate method clustering.

A2 finds *exact* structural clones. This finds the softer case: methods that
are 75-99% the same - the "write the next one by copy-pasting the last one"
family that makes a 558-method class. Bodies are normalized to a token
sequence, compared pairwise with difflib, and clustered greedily.

Run:  python scripts/analysis/method_similarity.py [--min-lines 4] [--threshold 0.80]
"""
from __future__ import annotations

import argparse
import ast
from collections import defaultdict
from difflib import SequenceMatcher

from common import Report, iter_functions

DEFAULT_THRESHOLD = 0.80
DEFAULT_MIN_LINES = 4
MAX_PER_CLASS = 700          # guard against O(n^2) blow-up on huge classes


def tokenize(node):
    """Normalized token stream for a function body.

    Keeps structure and called-method names, drops receivers/identifiers/
    literals - same philosophy as A2's normalizer but as a flat sequence so
    partial (not just exact) similarity is measurable.
    """
    tokens = []
    for child in ast.walk(node):
        if isinstance(child, ast.Attribute):
            tokens.append(f".{child.attr}")
        elif isinstance(child, ast.Call):
            tokens.append("call")
        elif isinstance(child, ast.Name):
            tokens.append("name")
        elif isinstance(child, ast.Constant):
            tokens.append("const" if child.value not in (True, False, None)
                          else str(child.value))
        else:
            tokens.append(type(child).__name__)
    return tokens


def cluster(functions, threshold):
    """Greedy clustering: seed with the largest function, absorb anything
    similar enough, repeat."""
    remaining = sorted(functions, key=lambda f: -len(f[1]))
    clusters = []
    while remaining:
        seed_fn, seed_tokens = remaining.pop(0)
        group = [seed_fn]
        leftovers = []
        for fn, tokens in remaining:
            ratio = SequenceMatcher(None, seed_tokens, tokens).quick_ratio()
            if ratio >= threshold:
                # quick_ratio is an upper bound; confirm with the real one.
                if SequenceMatcher(None, seed_tokens, tokens).ratio() >= threshold:
                    group.append(fn)
                    continue
            leftovers.append((fn, tokens))
        remaining = leftovers
        if len(group) >= 2:
            clusters.append(group)
    return clusters


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--threshold", type=float, default=DEFAULT_THRESHOLD)
    ap.add_argument("--min-lines", type=int, default=DEFAULT_MIN_LINES)
    args = ap.parse_args()

    # Group by owning class/module so we compare like with like and keep the
    # pairwise cost bounded.
    by_owner = defaultdict(list)
    for fn in iter_functions():
        if fn.n_lines < args.min_lines:
            continue
        tokens = tokenize(fn.node)
        if len(tokens) < 12:
            continue
        by_owner[(fn.file, fn.classname)].append((fn, tokens))

    report = Report(
        "A3 - Near-duplicate method clusters",
        f"Methods >= {args.threshold:.0%} similar after normalization, "
        f"grouped within each class/module. These are copy-paste families: "
        "each cluster is a candidate for one parameterized method, a shared "
        "helper, or a declarative table.")

    all_clusters = []
    for (file, classname), functions in by_owner.items():
        if len(functions) < 2 or len(functions) > MAX_PER_CLASS:
            continue
        for group in cluster(functions, args.threshold):
            all_clusters.append((file, classname, group))

    all_clusters.sort(key=lambda c: -sum(f.n_lines for f in c[2]))

    report.line("## Summary\n")
    total_methods = sum(len(g) for _, _, g in all_clusters)
    total_lines = sum(sum(f.n_lines for f in g) for _, _, g in all_clusters)
    report.line(f"- **{len(all_clusters)}** clusters")
    report.line(f"- **{total_methods}** methods involved")
    report.line(f"- **~{total_lines}** lines in clustered methods\n")

    rows = []
    for file, classname, group in all_clusters[:40]:
        rows.append((len(group), sum(f.n_lines for f in group),
                     classname or "(module)", f"`{file}`",
                     ", ".join(f.name for f in group[:4])
                     + (f" +{len(group)-4}" if len(group) > 4 else "")))
    report.table(("Methods", "Lines", "Owner", "File", "Members"), rows)

    report.line("## Cluster detail\n")
    for file, classname, group in all_clusters[:30]:
        owner = classname or "(module-level)"
        report.line(f"### {owner} - {len(group)} similar methods "
                    f"({sum(f.n_lines for f in group)} lines)\n")
        report.line(f"`{file}`\n")
        for fn in sorted(group, key=lambda f: f.lineno):
            report.line(f"- `{fn.name}` - line {fn.lineno} ({fn.n_lines} lines)")
        report.line()

    out = report.write("A3_SIMILARITY.md")
    print(f"[A3] {len(all_clusters)} clusters, {total_methods} methods, "
          f"~{total_lines} lines -> {out}")


if __name__ == "__main__":
    main()
