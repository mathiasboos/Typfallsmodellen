"""Two reference series the PGB sheet's own conscription/study formulas read.

'Några tal' col AF ("Studiebidrag/termin utan tillägg") and col AG ("PA avgift
för studier som ger individen ett PGB") are not economic projections the way
extract_series.py's columns are -- the sheet never adjusts them for inflation
or growth, it just holds a literal per-year decision that goes flat once the
sheet's own data runs out. `wsPGB!L15`/`N15` read them with
``INDEX('Några tal'!$AF$7:$AF$261, year-1959)``/``$AG$5:$AG$163`` respectively,
which is the same ``row = year - 1953`` addressing extract_series.py already
uses for every other column on this sheet (both INDEX ranges start at
``ROW_BASE + firstYear``, so the offset falls out the same way regardless of
which row the range itself starts on).

col F ("Medel PGI") on the same sheet needs no new extraction at all -- it is
already ``medelPgi`` in economic-series.json, read by the engine as
``v.mpgi.get(age)``, which is what wsPGB!G ("50% medel efter 1995") turns out
to just be half of, gated to the years conscription actually existed.
"""

from __future__ import annotations

from common import Series, exact

SHEET = "Några tal"
ROW_BASE = 1953
FIRST_YEAR = 1960

# name -> (column, description)
COLUMNS: dict[str, tuple[int, str]] = {
    "studyGrantPerTerm": (32, "Studiebidrag/termin utan tillägg"),
    "studyPgbFactor": (33, "PA avgift för studier som ger individen ett PGB"),
}


def last_real_year(sheet, col: int, first_year: int) -> int:
    """Highest year whose cell is a genuine, nonzero decision.

    Neither column is one contiguous run of real values: col AG (the
    PGB-generating share) reads blank for 1960-1994 -- studies did not earn
    PGB at all until then -- before its real run starts, and col AF (the
    grant rate) is pre-filled with formula-driven zeros well past its own
    last real figure (verified against the source workbook: 20 600 kr
    through 2110, then 0 from 2111 on, for another forty-plus rows before
    finally running out and reading blank) -- an artifact of however far the
    sheet's own fill happens to reach, not a policy that conscription-era
    study grants stop existing. So this scans the sheet's whole used range
    rather than stopping at the first blank or zero it meets, and the engine
    holds the last *real* year's value flat past whatever it finds.
    """
    last = first_year - 1
    for row in range(first_year - ROW_BASE, sheet.nrows + 1):
        v = sheet.num(row, col)
        if v:
            last = row + ROW_BASE
    return last


def extract(wb) -> dict:
    sheet = wb.sheet(SHEET)
    series: dict[str, Series] = {}
    for name, (col, description) in COLUMNS.items():
        last_year = last_real_year(sheet, col, FIRST_YEAR)
        values = [exact(sheet.num(y - ROW_BASE, col)) for y in range(FIRST_YEAR, last_year + 1)]
        values = [0 if v is None else v for v in values]
        series[name] = Series(
            name=name,
            first_year=FIRST_YEAR,
            values=values,
            last_actual_year=last_year,
            source=f"'{SHEET}'!col {col} - {description}",
            note="held flat past lastActualYear, the same as the sheet's own plateau up to it",
        )
    return {
        "source": f"'{SHEET}', rows for {FIRST_YEAR}- (row = year - {ROW_BASE})",
        "note": (
            "Not an economic projection -- these are literal per-year decisions "
            "(a study grant rate, a PGB-generating share of the pension fee) the "
            "sheet itself never adjusts for inflation or growth. The engine "
            "clamps to the last extracted year rather than projecting further."
        ),
        "series": {name: s.to_json() for name, s in series.items()},
    }
