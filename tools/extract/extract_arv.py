"""Inheritance-gain factors (arvsvinstfaktorer) for income and premium pension.

Addressing mirrors startsetup exactly:

  income pension, under riktålder   arv IP!Cells(age - 17 + 3, year - 1999 + 3)
  income pension, the double factor arv IP!Cells(81 + age - 60, year - 1999 + 3)
  premium pension                   arv PP!Cells(age - 14 + 2, year - 1999 + 3)

i.e. row = age - 14 / age + 21 / age - 12, and column = year - 1996 throughout.
"""

from __future__ import annotations

from common import round_sig

FIRST_YEAR = 1999
COLUMN_BASE = 1996  # column = year - 1996


def _read_block(sheet, row_for_age, ages: range, years: range) -> dict:
    values = []
    for age in ages:
        row = row_for_age(age)
        values.append([round_sig(sheet.num(row, y - COLUMN_BASE)) for y in years])
    return {
        "firstAge": ages.start,
        "lastAge": ages.stop - 1,
        "firstYear": years.start,
        "lastYear": years.stop - 1,
        "values": values,  # [age][year]
    }


def _last_year(sheet, row: int, start: int) -> int:
    year = start
    while sheet.num(row, year + 1 - COLUMN_BASE) is not None:
        year += 1
    return year


def extract(wb) -> dict:
    ip_sheet = wb.sheet("arv IP")
    pp_sheet = wb.sheet("arv PP")

    # Header rows carry the year labels; walk them to find how far the tables run.
    ip_last = _last_year(ip_sheet, 2, FIRST_YEAR)
    pp_last = _last_year(pp_sheet, 2, FIRST_YEAR)

    return {
        "source": "arv IP / arv PP, addressed as in startsetup",
        "incomePension": {
            "note": "IP_arv1: applies below riktålder. Row = age - 14.",
            **_read_block(ip_sheet, lambda a: a - 14, range(17, 92), range(FIRST_YEAR, ip_last + 1)),
        },
        "incomePensionDouble": {
            "note": "IP_arv2: the double inheritance gain from age 60. Row = age + 21.",
            **_read_block(ip_sheet, lambda a: a + 21, range(60, 106), range(FIRST_YEAR, ip_last + 1)),
        },
        "premiumPension": {
            "note": "PP_arv. Row = age - 12; ages past 105 reuse the age-105 row.",
            **_read_block(pp_sheet, lambda a: a - 12, range(15, 106), range(FIRST_YEAR, pp_last + 1)),
        },
    }
