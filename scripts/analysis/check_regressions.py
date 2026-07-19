"""Phase D - guard against reintroduced boilerplate and dead references.

Compares the current idiom/usage counts against a committed baseline
(``baseline.json``) and fails when a number goes **up**. The point is not
to reach zero - some of these idioms are legitimate - but to make the
direction of travel visible and stop silent regression.

    python scripts/analysis/check_regressions.py          # check
    python scripts/analysis/check_regressions.py --update # accept current

Exit 0 = no regression. Exit 1 = at least one counter increased.

Why this exists: the single most-duplicated idiom in this codebase
(``blockSignals`` triads, 245 sites at baseline) had a *correct* shared
helper available the whole time - it was used 6 times. Availability alone
did not stop the pattern spreading. A counter that fails loudly does.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

import find_idioms  # noqa: E402
import call_graph  # noqa: E402
import find_shadowed  # noqa: E402
from common import check_coverage  # noqa: E402

BASELINE_PATH = Path(__file__).parent / "baseline.json"

# Counters that must stay at 0 - each represents code that cannot work:
#   phantom_calls       call to a name defined nowhere (the
#                       _show_spectro_popup bug class)
#   shadowed_defs       same method defined twice in one class; the first
#                       is dead. Found three such pairs in SXMGridViewer.
#   broken_delegations  `module.func(...)` where module has no `func` -
#                       AttributeError if reached. All three found were
#                       *also* shadowed, so nothing failed until someone
#                       would have deleted the shadowing copy.
ZERO_TOLERANCE = ("phantom_calls", "shadowed_defs", "broken_delegations")


def measure():
    counts = {}
    for name, hits in find_idioms.scan_all().items():
        counts[name] = len(hits)
    defs = call_graph.collect_definitions()
    phantoms, _injected = call_graph.find_phantom_calls(defs)
    counts["phantom_calls"] = len(phantoms)
    counts["shadowed_defs"] = len(find_shadowed.find_shadowed_methods())
    counts["broken_delegations"] = len(
        find_shadowed.find_broken_delegations(find_shadowed.module_exports()))
    return counts


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--update", action="store_true",
                    help="write current counts as the new baseline")
    args = ap.parse_args()

    bad_parse = check_coverage(verbose=True)
    if bad_parse:
        print("ERROR: some files do not parse; counts would be wrong.")
        return 1

    counts = measure()

    if args.update or not BASELINE_PATH.exists():
        BASELINE_PATH.write_text(json.dumps(counts, indent=2, sort_keys=True),
                                 encoding="utf-8")
        print(f"baseline written to {BASELINE_PATH.name}:")
        for key, value in sorted(counts.items()):
            print(f"  {key:<26} {value}")
        return 0

    baseline = json.loads(BASELINE_PATH.read_text(encoding="utf-8"))
    failed = []
    print(f"{'counter':<26} {'baseline':>9} {'now':>7} {'delta':>7}")
    for key in sorted(set(baseline) | set(counts)):
        was = baseline.get(key)
        now = counts.get(key)
        if was is None or now is None:
            print(f"{key:<26} {'-' if was is None else was:>9} "
                  f"{'-' if now is None else now:>7}    (new/removed)")
            continue
        delta = now - was
        flag = ""
        if delta > 0:
            flag = "  <-- REGRESSION"
            failed.append((key, was, now))
        elif delta < 0:
            flag = "  improved"
        print(f"{key:<26} {was:>9} {now:>7} {delta:>+7}{flag}")

    for key in ZERO_TOLERANCE:
        if counts.get(key, 0) > 0 and not any(f[0] == key for f in failed):
            failed.append((key, 0, counts[key]))
            print(f"{key}: must be 0, found {counts[key]}")

    if failed:
        print("\nRegression detected. Either use the shared helper "
              "(see docs/refactor/PATTERNS.md) or, if the new code is "
              "genuinely justified, re-run with --update and say why in the "
              "commit message.")
        return 1
    print("\nNo regressions.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
