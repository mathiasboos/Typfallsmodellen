"""Death probabilities and the workbook's own cached annuity factors.

Columns A:D of the ``mortality`` sheet hold SCB's death risks as
(year, sex, age, risk) -- the input ``ReadMortality`` loads.  Columns P:Z hold
the delningstal and arvsvinstfaktorer the workbook computed from them, which is
a ready-made regression fixture for the ported ``Calculate_Deltal``.

The risks ship as a Float64 binary rather than JSON: ~30k values are 240 KB
packed against roughly 700 KB of JSON text, and the engine wants a typed array
anyway.  Float64 rather than Float32 because the annuity factors compound
survival probabilities across fifty-odd years, where float32's seven significant
digits would drift away from the VBA's Double arithmetic.
"""

from __future__ import annotations

import struct

SHEET = "mortality"
MAX_AGE = 105  # agegroup = 106 in Mortality.bas, i.e. ages 0..105
SEXES = (1, 2)  # 1 = male, 2 = female


def extract(wb):
    """Return (metadata dict, packed Float64 risks) for the death-probability grid."""
    sheet = wb.sheet(SHEET)

    rows = []
    for r in range(2, sheet.nrows + 1):
        year, sex, age, risk = (sheet.num(r, c) for c in (1, 2, 3, 4))
        if year is None or sex is None or age is None or risk is None:
            continue
        rows.append((int(year), int(sex), int(age), risk))
    if not rows:
        raise RuntimeError("no death probabilities found on the mortality sheet")

    first_year = min(r[0] for r in rows)
    last_year = max(r[0] for r in rows)
    n_years, n_ages = last_year - first_year + 1, MAX_AGE + 1

    # [year][sex][age], flattened; index = ((year - first) * 2 + (sex - 1)) * n_ages + age
    grid = [0.0] * (n_years * len(SEXES) * n_ages)
    seen = 0
    for year, sex, age, risk in rows:
        if sex not in SEXES or age > MAX_AGE:
            continue
        grid[((year - first_year) * len(SEXES) + (sex - 1)) * n_ages + age] = risk
        seen += 1

    expected = n_years * len(SEXES) * n_ages
    meta = {
        "firstYear": first_year,
        "lastYear": last_year,
        "maxAge": MAX_AGE,
        "sexes": list(SEXES),
        "layout": "((year - firstYear) * 2 + (sex - 1)) * (maxAge + 1) + age",
        "dtype": "float64",
        "count": expected,
        "populated": seen,
        "source": f"{SHEET}!A:D (year, sex, age, risk)",
    }
    return meta, struct.pack(f"<{expected}d", *grid)


def extract_cached_annuity_factors(wb) -> str:
    """The workbook's own delningstal/arvsvinst output, P:Z, as a CSV fixture.

    CSV rather than JSON: 17k rows of eleven numbers are far smaller without the
    brackets, and a yearly regeneration then produces a line-by-line diff instead
    of one unreadable blob.
    """
    sheet = wb.sheet(SHEET)
    labels = [sheet.text(1, c) or f"col{c}" for c in range(16, 27)]

    lines = [",".join(labels)]
    for r in range(2, sheet.nrows + 1):
        values = [sheet.num(r, c) for c in range(16, 27)]
        if values[0] is None:
            continue
        lines.append(",".join("" if v is None else repr(v) for v in values))
    return "\n".join(lines) + "\n"
