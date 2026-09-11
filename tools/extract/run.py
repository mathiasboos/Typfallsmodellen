#!/usr/bin/env python3
"""Regenerate packages/data from a Typfallsmodellen .xlsb workbook.

    python tools/extract/run.py source/Typfallsmodellen.xlsb

Run this once a year, when Pensionsmyndigheten publishes a new workbook, then
review the resulting git diff.  See docs/YEARLY-UPDATE.md for the full runbook.

Output is deterministic: the same workbook in gives byte-identical files out, so
a diff shows only what genuinely changed between releases.
"""

from __future__ import annotations

import argparse
import hashlib
import sys
from datetime import datetime, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import extract_arv
import extract_content
import extract_i18n
import extract_mortality
import extract_options
import extract_series
import extract_tax
from common import DATA_DIR, REPO_ROOT, Workbook, write_bytes, write_json
from dump_vba import extract as dump_vba
from formulas import FormulaMap, convert_to_xlsx

VBA_DIR = REPO_ROOT / "reference" / "vba"
FIXTURE_DIR = REPO_ROOT / "reference" / "fixtures"


def file_digest(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("workbook", nargs="?", default=str(REPO_ROOT / "source" / "Typfallsmodellen.xlsb"))
    parser.add_argument("--skip-vba", action="store_true", help="skip the VBA dump (slow; only needed on a new release)")
    args = parser.parse_args()

    xlsb = Path(args.workbook).resolve()
    if not xlsb.exists():
        print(f"error: no workbook at {xlsb}", file=sys.stderr)
        return 1

    print(f"Reading {xlsb.relative_to(REPO_ROOT) if xlsb.is_relative_to(REPO_ROOT) else xlsb}")
    wb = Workbook(xlsb)

    print("Converting to .xlsx to read worksheet formulas (LibreOffice; cached)...")
    formula_map = FormulaMap(convert_to_xlsx(xlsb))

    manifest: dict = {
        "generatedBy": "tools/extract/run.py",
        "generatedAt": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "source": {
            "file": xlsb.name,
            "bytes": xlsb.stat().st_size,
            "sha256": file_digest(xlsb),
        },
    }

    print("Economic series...")
    last_year = extract_series.last_year_with_data(wb)
    series, anchors, cached_series = extract_series.extract(wb, formula_map, last_year)
    write_json(
        DATA_DIR / "economic-series.json",
        {
            "source": f"'Några tal', rows for {extract_series.FIRST_YEAR}-{last_year} (row = year - {extract_series.ROW_BASE})",
            "note": (
                "Actual values only. Years past lastActualYear are projected by the engine "
                "from the user's inflation and growth assumptions; see docs/PROJECTION-RULES.md."
            ),
            "lastSheetYear": last_year,
            "anchors": anchors,
            "series": {name: s.to_json() for name, s in series.items()},
        },
        description="yearly economic actuals",
    )
    write_json(
        FIXTURE_DIR / "economic-series-cached.json",
        {
            "note": (
                "Every value the workbook had cached, projections included, computed with the "
                "shipped assumptions (0% inflation, 0% growth, 1.7% real return). The engine's "
                "projection must reproduce these exactly."
            ),
            "firstYear": extract_series.FIRST_YEAR,
            "lastYear": last_year,
            "assumptions": {"inflation": 0.0, "realGrowth": 0.0, "realReturn": 0.017},
            "series": cached_series,
        },
        description="projection regression fixture",
    )
    manifest["series"] = {
        name: {"lastActualYear": s.last_actual_year, "projection": s.note.replace("projection: ", "")}
        for name, s in series.items()
    }

    print("Mortality...")
    mortality_meta, mortality_bin = extract_mortality.extract(wb)
    write_bytes(DATA_DIR / "mortality-risks.bin", mortality_bin, description="SCB death probabilities")
    write_json(DATA_DIR / "mortality.json", mortality_meta, description="mortality grid layout")
    write_bytes(
        FIXTURE_DIR / "annuity-factors-cached.csv",
        extract_mortality.extract_cached_annuity_factors(wb).encode("utf-8"),
        description="delningstal/arvsvinst regression fixture",
    )
    manifest["mortality"] = {
        "firstYear": mortality_meta["firstYear"],
        "lastYear": mortality_meta["lastYear"],
        "values": mortality_meta["count"],
    }

    print("Inheritance gains...")
    arv = extract_arv.extract(wb)
    write_json(DATA_DIR / "inheritance-gains.json", arv, description="arvsvinstfaktorer")

    print("Tax rates...")
    tax = extract_tax.extract(wb, formula_map)
    write_json(DATA_DIR / "municipal-tax.json", tax, description="K_skatt")

    print("Options and defaults...")
    options = extract_options.extract(wb)
    write_json(DATA_DIR / "options.json", options, description="form contract")

    print("Translations...")
    i18n = extract_i18n.extract(wb)
    write_json(DATA_DIR / "i18n.json", i18n, description="sv/en strings")

    print("Content...")
    content = extract_content.extract(wb)
    write_json(DATA_DIR / "content.json", content, description="glossary, explanations, changelog")
    manifest["modelVersion"] = content.get("modelVersion")

    if args.skip_vba:
        print("Skipping VBA dump (--skip-vba)")
    else:
        print("VBA source...")
        manifest["vba"] = dump_vba(xlsb, VBA_DIR)
        print(f"  wrote {manifest['vba']['moduleCount']} modules to {VBA_DIR.relative_to(REPO_ROOT)}")

    write_json(DATA_DIR / "manifest.json", manifest, description="provenance")

    print()
    print(f"Model version: {manifest.get('modelVersion')}")
    print("Last actual year per series:")
    for name, info in sorted(manifest["series"].items()):
        print(f"  {name:24s} {info['lastActualYear']}   ({info['projection']})")
    print()
    print("Review the git diff before committing. See docs/YEARLY-UPDATE.md.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
