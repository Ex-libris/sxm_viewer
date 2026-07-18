"""A5 - attribute cohesion: find the classes hiding inside the god object.

Builds the method <-> attribute access graph for a large class, then clusters
attributes that are consistently used by the same methods. Attributes that
always travel together are, by definition, one piece of state - i.e. a
candidate extracted class.

This is read-only evidence for the eventual decomposition: it replaces "I
think preview state belongs together" with a measured grouping, and it also
scores each candidate group by how *isolated* it is (how few methods touch
both it and anything else). Low-coupling groups are the safe ones to extract
first.

Run:  python scripts/analysis/attribute_cohesion.py [--class SXMGridViewer]
"""
from __future__ import annotations

import argparse
import ast
from collections import defaultdict

from common import Report, check_coverage, iter_source_files, parse, rel

DEFAULT_CLASSES = ("SXMGridViewer", "MultiPreviewCanvas")
MIN_METHODS_PER_ATTR = 2
MIN_GROUP_SIZE = 3


def class_nodes(target):
    for path in iter_source_files():
        tree = parse(path)
        if tree is None:
            continue
        for node in ast.walk(tree):
            if isinstance(node, ast.ClassDef) and node.name == target:
                yield rel(path), node


def method_attribute_map(cls_node):
    """{method_name: set(attribute_names)} for `self.attr` accesses."""
    result = {}
    for child in cls_node.body:
        if not isinstance(child, (ast.FunctionDef, ast.AsyncFunctionDef)):
            continue
        attrs = set()
        for node in ast.walk(child):
            if (isinstance(node, ast.Attribute)
                    and isinstance(node.value, ast.Name)
                    and node.value.id == "self"):
                attrs.add(node.attr)
        result[child.name] = attrs
    return result


def cluster_attributes(m2a, methods_by_attr):
    """Greedy cohesion clustering.

    Similarity between two attributes = Jaccard overlap of the method sets
    that touch them. Seeds with the most-used attribute and absorbs anything
    strongly overlapping, so groups come out as "state used by the same
    code".
    """
    attrs = sorted(methods_by_attr, key=lambda a: -len(methods_by_attr[a]))
    attrs = [a for a in attrs if len(methods_by_attr[a]) >= MIN_METHODS_PER_ATTR]
    unassigned = list(attrs)
    groups = []
    while unassigned:
        seed = unassigned.pop(0)
        group = [seed]
        seed_methods = methods_by_attr[seed]
        rest = []
        for attr in unassigned:
            other = methods_by_attr[attr]
            union = seed_methods | other
            if not union:
                rest.append(attr)
                continue
            jaccard = len(seed_methods & other) / len(union)
            if jaccard >= 0.5:
                group.append(attr)
            else:
                rest.append(attr)
        unassigned = rest
        if len(group) >= MIN_GROUP_SIZE:
            groups.append(group)
    return groups


def isolation_score(group, m2a, methods_by_attr):
    """Fraction of touching methods that touch ONLY this group.

    1.0 = perfectly separable (extract with no ripple); low = entangled.
    """
    group_set = set(group)
    touching = set()
    for attr in group:
        touching |= methods_by_attr[attr]
    if not touching:
        return 0.0, 0, 0
    pure = sum(1 for m in touching if m2a[m] and m2a[m] <= group_set)
    return pure / len(touching), pure, len(touching)


def analyse(target, report):
    found = list(class_nodes(target))
    if not found:
        report.line(f"## {target}\n\n_class not found_\n")
        return
    file, node = found[0]
    m2a = method_attribute_map(node)
    methods_by_attr = defaultdict(set)
    for method, attrs in m2a.items():
        for attr in attrs:
            methods_by_attr[attr].add(method)

    groups = cluster_attributes(m2a, methods_by_attr)
    scored = []
    for group in groups:
        score, pure, touching = isolation_score(group, m2a, methods_by_attr)
        scored.append((score, pure, touching, group))
    scored.sort(key=lambda s: (-s[0], -len(s[3])))

    report.line(f"## {target}\n")
    report.line(f"`{file}`\n")
    report.line(f"- methods analysed: **{len(m2a)}**")
    report.line(f"- distinct `self.` attributes: **{len(methods_by_attr)}**")
    report.line(f"- cohesive groups (>= {MIN_GROUP_SIZE} attrs): "
                f"**{len(groups)}**\n")

    report.line("### Candidate extractions, most isolated first\n")
    report.line("`Isolation` = share of methods touching this group that "
                "touch *nothing else*. High = safe to extract.\n")
    rows = []
    for score, pure, touching, group in scored[:25]:
        rows.append((f"{score:.0%}", len(group), f"{pure}/{touching}",
                     ", ".join(f"`{a}`" for a in group[:6])
                     + (f" +{len(group)-6}" if len(group) > 6 else "")))
    report.table(("Isolation", "Attrs", "Pure/Touching methods", "Attributes"),
                 rows)

    report.line("### Group detail\n")
    for score, pure, touching, group in scored[:12]:
        report.line(f"<details><summary>{len(group)} attributes, "
                    f"{score:.0%} isolated ({pure}/{touching} methods)"
                    "</summary>\n")
        for attr in sorted(group):
            report.line(f"- `self.{attr}` - {len(methods_by_attr[attr])} methods")
        report.line("\n</details>\n")

    # Attributes touched by a huge share of methods are the "spine" - they
    # cannot be extracted and identify the true core responsibility.
    spine = sorted(methods_by_attr.items(), key=lambda kv: -len(kv[1]))[:15]
    report.line("### Spine attributes (touched by the most methods)\n")
    report.line("> These resist extraction and define what the class is "
                "*actually* about; everything else is a passenger.\n")
    report.table(("Attribute", "Methods touching"),
                 [(f"`self.{a}`", len(ms)) for a, ms in spine])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--class", dest="classes", action="append",
                    default=None, help="class name (repeatable)")
    args = ap.parse_args()
    targets = args.classes or list(DEFAULT_CLASSES)

    print("[A5] parse coverage check:")
    check_coverage()

    report = Report(
        "A5 - Attribute cohesion (hidden class boundaries)",
        "Clusters a large class's attributes by which methods use them. "
        "Attributes that always travel together are one piece of state - a "
        "candidate extracted class. Isolation scores rank which groups can "
        "be pulled out with the least ripple. Read-only evidence; no "
        "refactoring is implied by inclusion here.")

    for target in targets:
        analyse(target, report)

    out = report.write("A5_COHESION.md")
    print(f"[A5] analysed {', '.join(targets)} -> {out}")


if __name__ == "__main__":
    main()
