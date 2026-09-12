#!/usr/bin/env python3
"""The lowest pension age and the riktålder, per cohort.

`Mcalc` and `startsetup` do not compute these -- they read two cells:

    riktl     = Application.Range("Rng_riktL")     ' Nyckeltal!DR2
    riktalder = Application.Range("Rng_riktage")   ' Nyckeltal!DR1

and both are looked up **by cohort** from the Nyckeltal sheet, whose columns 121
and 122 are headed "Lägsta ålder" and "riktalder".  The author says so twice:
"för aktuell årskull" at VBA_go.bas:155, and at :780 "Nedan avseende årskull men
lagstiftningen ser till inkomstår".

That matters because the VBA also has a `riktage(year, typ)` function, keyed on
the *income year*, and the two disagree for about half the cohorts: a 1970
typfall's lowest age is 65 here, where `riktage(1970 + 64, 0)` gives 64.  The
function is still the right answer where the model pins a rules year -- see
VBA_go.bas 161 and 789 -- but not for the cohort's own limits.

Addressing mirrors the VBA:  Nyckeltal!Cells(5 + born - 1930, 121)
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

try:
    from common import DATA_DIR, REPO_ROOT, Workbook, write_json
except ImportError as exc:  # pragma: no cover - a setup problem, not a data one
    # Reading the .xlsb needs pyxlsb, which lives in the extraction venv. Run
    # ourselves under it rather than making the caller know that.
    _venv = Path(__file__).resolve().parents[2] / ".venv" / "bin" / "python"
    if _venv.exists() and not os.environ.get("_REEXEC"):
        os.environ["_REEXEC"] = "1"
        os.execv(str(_venv), [str(_venv), __file__, *sys.argv[1:]])
    raise SystemExit(
        f"{exc}. Reading the .xlsb needs pyxlsb:\n"
        "  python3 -m venv .venv && .venv/bin/pip install -r tools/extract/requirements.txt"
    ) from exc

SHEET = "Nyckeltal"
ROW_BASE = 1930          # row = 5 + born - 1930
FIRST_ROW = 5
COL_COHORT = 119
COL_LOWEST = 121
COL_RIKTALDER = 122

OUTPUT = DATA_DIR / "retirement-ages.json"


def extract(wb) -> dict:
    sheet = wb.sheet(SHEET)

    cohorts: list[int] = []
    lowest: list[int] = []
    riktalder: list[int] = []

    row = FIRST_ROW
    while True:
        born = sheet.num(row, COL_COHORT)
        low = sheet.num(row, COL_LOWEST)
        rikt = sheet.num(row, COL_RIKTALDER)
        if born is None or low is None or rikt is None:
            break
        born = int(born)
        if born != ROW_BASE + row - FIRST_ROW:
            raise SystemExit(
                f"{SHEET} row {row} says cohort {born}, but the VBA addresses it as "
                f"{ROW_BASE + row - FIRST_ROW} (row = 5 + born - 1930)"
            )
        cohorts.append(born)
        lowest.append(int(low))
        riktalder.append(int(rikt))
        row += 1

    if not cohorts:
        raise SystemExit(f"no cohort rows found on {SHEET} from row {FIRST_ROW}")

    return {
        "source": f"{SHEET}, columns {COL_LOWEST} and {COL_RIKTALDER} (row = 5 + born - {ROW_BASE})",
        "note": (
            "Read by cohort, as Rng_riktL and Rng_riktage are. Not the same as the "
            "riktage(year, typ) function, which is keyed on the income year and is "
            "used only where a rules year is pinned."
        ),
        "firstCohort": cohorts[0],
        "lastCohort": cohorts[-1],
        "lowest": lowest,
        "riktalder": riktalder,
    }


def check(workbook: Path) -> int:
    if not OUTPUT.exists():
        print(f"error: {OUTPUT.relative_to(REPO_ROOT)} does not exist; run the extractor", file=sys.stderr)
        return 1

    fresh = extract(Workbook(workbook))
    committed = json.loads(OUTPUT.read_text(encoding="utf-8"))

    problems = []
    for key in ("firstCohort", "lastCohort"):
        if fresh[key] != committed.get(key):
            problems.append(f"{key}: workbook {fresh[key]}, committed {committed.get(key)}")

    for key in ("lowest", "riktalder"):
        a, b = fresh[key], committed.get(key, [])
        if a == b:
            continue
        for i, (x, y) in enumerate(zip(a, b)):
            if x != y:
                problems.append(f"{key}[{fresh['firstCohort'] + i}]: workbook {x}, committed {y}")
        if len(a) != len(b):
            problems.append(f"{key}: workbook has {len(a)} cohorts, committed has {len(b)}")

    if problems:
        print("retirement ages differ from the workbook:", file=sys.stderr)
        for problem in problems[:20]:
            print(f"  {problem}", file=sys.stderr)
        if len(problems) > 20:
            print(f"  ... and {len(problems) - 20} more", file=sys.stderr)
        return 1

    print(
        f"retirement ages: {len(fresh['lowest'])} cohorts "
        f"({fresh['firstCohort']}-{fresh['lastCohort']}) match {SHEET}"
    )
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("workbook", nargs="?", default=str(REPO_ROOT / "source" / "Typfallsmodellen.xlsb"))
    parser.add_argument("--check", action="store_true", help="compare the workbook against the committed JSON")
    args = parser.parse_args()

    workbook = Path(args.workbook).resolve()
    if not workbook.exists():
        print(f"error: no workbook at {workbook}", file=sys.stderr)
        return 1

    if args.check:
        return check(workbook)

    write_json(OUTPUT, extract(Workbook(workbook)), description="lowest pension age and riktålder per cohort")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
