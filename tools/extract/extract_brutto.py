"""The Brutto sheet: a worked example, cached, that exercises the earning phase.

Unlike the result sheets -- which the workbook clears on open -- Brutto keeps its
computed values, and its formulas call the model's VBA functions directly as
worksheet UDFs:

    =PGI(E15, I15, G15, H15, ..., 0, A15, 0)
    =IPavgift(J15 + M15, A15, $J$4, ...)
    =ppavgift(...)   =GPAVGIFT(...)

So it is a ready-made fixture for pgi, ipavgift, ppavgift and gpavgift, plus the
basbelopp and inheritance-factor lookups feeding them -- roughly a hundred years
of a single typfall, with no Excel needed.

Its limit is that it is *one* income (an earner above the ceiling in every year),
so it exercises the capped path thoroughly and the uncapped path not at all. The
golden files from a real Excel run remain the acceptance gate.

Column 42, the ATP points, is deliberately excluded: its formula adds `RAND()`
as a tie-breaker, so it is not reproducible.
"""

from __future__ import annotations

from common import round_sig

SHEET = "Brutto"
HEADER_ROW = 14
FIRST_DATA_ROW = 15

# column -> field name. Names follow the VBA, not the sheet's prose headings.
COLUMNS = {
    1: "age",
    2: "year",
    5: "income",
    7: "inkomstbasbelopp",
    8: "forhojtPrisbasbelopp",
    9: "prisbasbelopp",
    10: "pgi",
    11: "arbetsgivaravgift",
    12: "egenavgift",
    13: "pgb",
    14: "ipavgift",
    15: "ppavgift",
    16: "gpavgift",
    19: "forvaltningsfaktorIp",
    20: "forvaltningsfaktorPp",
    21: "arvsvinstIpUnder65",
    22: "arvsvinstIpOver65",
    23: "arvsvinstPp",
    29: "pbhIp",
    35: "pbhPp",
}


def extract(wb) -> dict:
    sheet = wb.sheet(SHEET)

    rows = []
    for r in range(FIRST_DATA_ROW, sheet.nrows + 1):
        year = sheet.num(r, 2)
        age = sheet.num(r, 1)
        if year is None or age is None:
            continue
        rows.append({name: round_sig(sheet.num(r, col)) for col, name in COLUMNS.items()})

    return {
        "source": f"{SHEET}, rows {FIRST_DATA_ROW}+ - cached results of the VBA functions called as worksheet UDFs",
        "note": (
            "One typfall, earning above the contribution ceiling every year. Validates pgi, "
            "ipavgift, ppavgift and gpavgift against the original without Excel. Column 42 "
            "(ATP points) is excluded because its formula adds RAND()."
        ),
        "referenceYear": round_sig(sheet.num(4, 2)),
        "shareOfNewSystem": round_sig(sheet.num(4, 10)),
        "columns": list(COLUMNS.values()),
        "rowCount": len(rows),
        "rows": rows,
    }
