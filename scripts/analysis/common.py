"""Shared helpers for the duplication-analysis toolkit.

Every tool in this package is **read-only**: it parses the source tree and
writes reports, never edits code. They are safe to run at any time and are
meant to stay useful after the initial deduplication pass - re-run them to
check whether a pattern is creeping back (see docs/refactor/PATTERNS.md).

Vendored code (``providers/nanonis/vendor``) is always excluded: it mirrors
an upstream package and must not be refactored.
"""
from __future__ import annotations

import ast
import io
from dataclasses import dataclass
from pathlib import Path

# Repo root = three levels up from this file (scripts/analysis/common.py).
REPO_ROOT = Path(__file__).resolve().parents[2]
PKG_ROOT = REPO_ROOT / "sxm_viewer"

# Matched against the path *relative to the scanned root* - an absolute
# match would exclude everything when the checkout itself lives under one
# of these names (e.g. a git worktree under .claude/worktrees/).
EXCLUDE_PARTS = ("vendor", "site-packages", "__pycache__")


def iter_source_files(root: Path | None = None):
    """Yield every analysable .py file under the package."""
    root = root or PKG_ROOT
    for path in sorted(root.rglob("*.py")):
        try:
            parts = set(path.relative_to(root).parts)
        except ValueError:
            parts = set(path.parts)
        if any(bad in parts for bad in EXCLUDE_PARTS):
            continue
        yield path


def rel(path: Path) -> str:
    try:
        return str(Path(path).resolve().relative_to(REPO_ROOT)).replace("\\", "/")
    except ValueError:
        return str(path)


def read_text(path: Path) -> str:
    """Read source, tolerating this repo's encoding quirks.

    Uses ``utf-8-sig``: four files (main_window.py, histogram.py,
    profile.py, quick_crop.py) carry a UTF-8 BOM, and a plain utf-8 read
    leaves U+FEFF at offset 0, which makes ``ast.parse`` raise
    SyntaxError. Reading those as utf-8 silently dropped the single
    biggest module from every analysis - always verify parse coverage
    (see ``check_coverage()``) before trusting a report.
    """
    return path.read_text(encoding="utf-8-sig", errors="replace")


def check_coverage(verbose: bool = True):
    """Report any file that fails to parse. Call this before trusting
    results - a silently-skipped file makes every count wrong."""
    bad = []
    for path in iter_source_files():
        if parse(path) is None:
            bad.append(path)
    if verbose and bad:
        print(f"  !! {len(bad)} file(s) failed to parse and were EXCLUDED:")
        for path in bad:
            print(f"     {rel(path)}")
    return bad


def parse(path: Path):
    """Parse a file to an AST, or None when it cannot be parsed."""
    try:
        return ast.parse(read_text(path), filename=str(path))
    except SyntaxError:
        return None


@dataclass
class FunctionInfo:
    """One function/method found in the tree."""
    file: str
    lineno: int
    end_lineno: int
    qualname: str        # "Class.method" or "module_function"
    name: str
    classname: str | None
    node: ast.AST

    @property
    def n_lines(self) -> int:
        return max(1, (self.end_lineno or self.lineno) - self.lineno + 1)


def iter_functions(root: Path | None = None):
    """Yield FunctionInfo for every function/method in the package."""
    for path in iter_source_files(root):
        tree = parse(path)
        if tree is None:
            continue
        for node in ast.walk(tree):
            if not isinstance(node, ast.ClassDef):
                continue
            for child in node.body:
                if isinstance(child, (ast.FunctionDef, ast.AsyncFunctionDef)):
                    yield FunctionInfo(
                        file=rel(path), lineno=child.lineno,
                        end_lineno=getattr(child, "end_lineno", child.lineno),
                        qualname=f"{node.name}.{child.name}", name=child.name,
                        classname=node.name, node=child)
        # Module-level functions (the gui/viewer/*.py convention).
        for child in tree.body:
            if isinstance(child, (ast.FunctionDef, ast.AsyncFunctionDef)):
                yield FunctionInfo(
                    file=rel(path), lineno=child.lineno,
                    end_lineno=getattr(child, "end_lineno", child.lineno),
                    qualname=child.name, name=child.name, classname=None,
                    node=child)


class Report:
    """Accumulates markdown output and writes it next to the other reports."""

    def __init__(self, title: str, intro: str = ""):
        self.buf = io.StringIO()
        self.buf.write(f"# {title}\n\n")
        if intro:
            self.buf.write(intro.strip() + "\n\n")

    def line(self, text: str = ""):
        self.buf.write(text + "\n")

    def table(self, headers, rows):
        self.line("| " + " | ".join(str(h) for h in headers) + " |")
        self.line("|" + "|".join("---" for _ in headers) + "|")
        for row in rows:
            self.line("| " + " | ".join(str(c) for c in row) + " |")
        self.line()

    def write(self, filename: str) -> Path:
        out_dir = REPO_ROOT / "docs" / "refactor"
        out_dir.mkdir(parents=True, exist_ok=True)
        out = out_dir / filename
        out.write_text(self.buf.getvalue(), encoding="utf-8")
        return out
