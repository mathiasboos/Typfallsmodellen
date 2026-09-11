"""Yearly economic series, from the 'Några tal' sheet.

'Några tal' is the source of truth.  The Nyckeltal sheet, which is what the VBA
actually reads, is a pass-through of 'Några tal' unless the "syntetisk
framskrivning" option is on -- so extracting here gets the same numbers one step
closer to the origin.

Both sheets put year Y on row ``Y - 1953``.  The VBA writes that offset two ways
(``year - 1959 + 6`` and ``year - 1958 + 5``); they are the same row.

Only *actual* values are extracted.  A series' decided values are typed into the
sheet as literals and its projection begins at the first formula cell, so the
last literal year is the last actual year -- detected exactly, via formulas.py,
rather than guessed.  Everything past that the engine projects from the user's
assumptions, mirroring the sheet formulas recorded in docs/PROJECTION-RULES.md.

Columns marked ``derived`` hold no actuals of their own at all: the sheet
computes them from other columns for every year, and so does the engine.
"""

from __future__ import annotations

from common import Series, round_sig
from formulas import FormulaMap

SHEET = "Några tal"
ROW_BASE = 1953
FIRST_YEAR = 1957

# name -> (column, projection rule, derived?, description)
# The projection rule names the function the engine applies past lastActualYear;
# see packages/engine/src/data/projection.ts and docs/PROJECTION-RULES.md.
COLUMNS: dict[str, tuple[int, str, bool, str]] = {
    "kpiJune":              (2,  "inflation",       False, "KPI, June figure (KPI_j)"),
    "kpiAnnual":            (3,  "inflation",       False, "KPI, annual average"),
    "prisbasbelopp":        (5,  "fromKpiJune",     False, "Prisbasbelopp (PBB)"),
    "medelPgi":             (6,  "followIndex",     False, "Average pension-qualifying income (MPGI)"),
    "inkomstbasbelopp":     (7,  "fromIncomeIndex", False, "Inkomstbasbelopp (IBB)"),
    "forhojtPrisbasbelopp": (8,  "fromKpiJune",     False, "Förhöjt prisbasbelopp (FPB)"),
    "inkomstindex":         (9,  "indexGrowth",     False, "Inkomstindex"),
    "balanstal":            (10, "neutralOne",      False, "Balanstal; projects to 1.0, not to its last value"),
    "balansindex":          (11, "balansindexFn",   False, "Balansindex; computed by the balansindex() function"),
    "gallandeIndex":        (12, "gallandeIndex",   True,  "Index in force: balansindex, or inkomstindex when rng_Senaste_Index_Framskrivning = 2"),
    "totalAvgiftPp":        (16, "feeTotal",        True,  "Total premium-pension fee: admin, plus management when returns are gross"),
    "avkastningPpm":        (17, "assumedReturn",   False, "Premium pension fund return, PPM index"),
    "avkastningAp7":        (18, "assumedReturn",   False, "AP7 Såfa return"),
    "rantaRiksgalden":      (19, "rgkSetting",      False, "Riksgälden rate for temporary management, percent"),
    "adminavgiftPp":        (20, "feeAdmin",        False, "Premium pension administration fee"),
    "forvaltningsavgiftPp": (21, "feeManagement",   False, "Premium pension fund management fee"),
    "skiktgrans1":          (29, "kpiPlusTwo",      False, "Threshold for state income tax (Tax_limit1)"),
    "skiktgrans2":          (30, "carryForward",    False, "Second state threshold (Tax_limit2); 1e16 since värnskatten was abolished"),
    "kvarEfterAdminIp":     (14, "carryForward",    False, "Share left after income-pension admin fees (IP_avg)"),
    "kvarEfterAvgiftPp":    (15, "oneMinusFee",     True,  "Share left after total premium-pension fees (PP_avg) = 1 - totalAvgiftPp"),
}

# Constants the sheet's projection formulas anchor on, read from the workbook so
# that a future release revising them is picked up rather than silently ignored.
ANCHOR_CELLS = {
    "pbbBase":            (4, 5),   # 'Några tal'!E4  - the 1957 prisbasbelopp, 36396
    "fpbBase":            (4, 8),   # 'Några tal'!H4  - 37144
    "ibbBase":            (4, 7),   # 'Några tal'!G4  - 43313
    "kpiJuneDivisor":     (44, 2),  # 'Några tal'!B44 - KPI June 1997, 257.38
    "incomeIndexDivisor": (52, 9),  # 'Några tal'!I52 - inkomstindex 2005, 118.41
    "kpiRoundDecimals":   (2, 3),   # 'Några tal'!C2  - decimals the KPI projection rounds to
}


def last_year_with_data(wb) -> int:
    """Highest year the sheet carries a year label for in column A."""
    sheet = wb.sheet(SHEET)
    year = FIRST_YEAR
    while sheet.num(year + 1 - ROW_BASE, 1) is not None:
        year += 1
    return year


def extract(wb, formula_map: FormulaMap, last_year: int):
    """Return (actual series, anchor constants, full cached series for verification)."""
    sheet = wb.sheet(SHEET)
    first_row, last_row = FIRST_YEAR - ROW_BASE, last_year - ROW_BASE

    series: dict[str, Series] = {}
    cached: dict[str, list] = {}

    for name, (col, projection, derived, description) in COLUMNS.items():
        all_values = [round_sig(sheet.num(y - ROW_BASE, col)) for y in range(FIRST_YEAR, last_year + 1)]
        cached[name] = all_values

        if derived:
            last_actual, keep = FIRST_YEAR - 1, 0
        else:
            literal_row = formula_map.last_literal_row(SHEET, col, first_row, last_row)
            last_actual = literal_row + ROW_BASE if literal_row is not None else FIRST_YEAR - 1
            keep = max(0, last_actual - FIRST_YEAR + 1)

        series[name] = Series(
            name=name,
            first_year=FIRST_YEAR,
            values=all_values[:keep],
            last_actual_year=last_actual,
            source=f"'{SHEET}'!C{col} - {description}",
            note=f"projection: {projection}" + (" (derived every year)" if derived else ""),
        )

    anchors = {key: round_sig(sheet.num(r, c)) for key, (r, c) in ANCHOR_CELLS.items()}
    return series, anchors, cached
