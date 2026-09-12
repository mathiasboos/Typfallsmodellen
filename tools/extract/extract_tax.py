"""Municipal tax rates and related charges, from the K_skatt sheet.

startsetup reads ``wsK_Skatt.Cells(year - 1930 + 2, 2)`` for the municipal rate
and column 8 for the burial fee, both as percentages -- so row = year - 1928.
"""

from __future__ import annotations

from common import exact

SHEET = "K_skatt"
ROW_BASE = 1928
FIRST_YEAR = 1930

COLUMNS = {
    "kommunalskatt": (2, "Average municipal tax rate, percent"),
    "begravningsavgift": (8, "Burial fee, percent; zero before 2000 in historical mode"),
    "kyrkoavgift": (9, "Church fee including burial fee, percent"),
    "arbetsgivaravgift": (10, "Employer social contribution, percent"),
    "varavAlderspension": (11, "of which the old-age pension component, percent"),
}


def extract(wb, formula_map) -> dict:
    sheet = wb.sheet(SHEET)

    last_year = FIRST_YEAR
    while sheet.num(last_year + 1 - ROW_BASE, 1) is not None:
        last_year += 1

    out = {
        "source": f"{SHEET}, row = year - {ROW_BASE}",
        "firstYear": FIRST_YEAR,
        "lastYear": last_year,
        "series": {},
    }
    for name, (col, description) in COLUMNS.items():
        values = [exact(sheet.num(y - ROW_BASE, col)) for y in range(FIRST_YEAR, last_year + 1)]
        literal_row = formula_map.last_literal_row(
            SHEET, col, FIRST_YEAR - ROW_BASE, last_year - ROW_BASE
        )
        out["series"][name] = {
            "column": col,
            "description": description,
            "lastActualYear": literal_row + ROW_BASE if literal_row is not None else None,
            "values": values,
        }
    return out
