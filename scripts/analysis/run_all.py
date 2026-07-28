"""Run the whole duplication-analysis toolkit and write all reports.

    python scripts/analysis/run_all.py

Always starts with a parse-coverage check: a file that fails to parse is
silently excluded from every analysis, which makes all counts quietly wrong.
That is not hypothetical - four files in this repo carry a UTF-8 BOM, and
before ``common.read_text`` switched to ``utf-8-sig`` they were dropped,
taking main_window.py (the single biggest refactor target) with them.
"""
from __future__ import annotations

import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

import common  # noqa: E402
import find_idioms  # noqa: E402
import find_clones  # noqa: E402
import method_similarity  # noqa: E402
import call_graph  # noqa: E402
import attribute_cohesion  # noqa: E402


def main():
    files = list(common.iter_source_files())
    print(f"Analysing {len(files)} source files under {common.PKG_ROOT}\n")
    print("Parse coverage:")
    bad = common.check_coverage()
    if not bad:
        print("  all files parsed OK\n")
    else:
        print("  ^ fix these first; reports below EXCLUDE them\n")

    for label, module in (("A1 idioms", find_idioms),
                          ("A2 clones", find_clones),
                          ("A3 similarity", method_similarity),
                          ("A4 usage", call_graph),
                          ("A5 cohesion", attribute_cohesion)):
        print(f"--- {label} ---")
        start = time.time()
        argv = sys.argv
        sys.argv = [argv[0]]           # tools use argparse with defaults
        try:
            module.main()
        finally:
            sys.argv = argv
        print(f"    ({time.time() - start:.1f}s)\n")

    print("Reports written to docs/refactor/")


if __name__ == "__main__":
    main()
