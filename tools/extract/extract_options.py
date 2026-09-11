"""Input options and defaults: the form contract for the website.

Three sources:

  NameRange     the occupational pension schemes and the return-basis choices,
                as numbered lists the Start sheet's drop-downs read
  Input         the birth years and retirement ages the drop-downs allow
  Adv_settings  every advanced setting, with its variable name in column 9 and
                its default ("Normala inställningar") in column 8 -- the
                workbook documents its own settings contract, so extract it
                rather than transcribing 60 values by hand
"""

from __future__ import annotations

from common import clean

# Blocks of numbered options on the NameRange sheet: (heading row, key)
NAMERANGE_LISTS = {
    "returnBasis": 12,             # 1 = chosen real return, 2 = PPM history, 3 = AP7
    "occupationalPension": 17,     # 1..8, the collective agreements
}


def _numbered_list(sheet, heading_row: int) -> list[dict]:
    options = []
    r = heading_row + 1
    while True:
        index, label = sheet.num(r, 1), sheet.text(r, 2)
        if index is None or not label:
            break
        options.append({"value": int(index), "label": label})
        r += 1
    return options


def _column_values(sheet, col: int, first_row: int = 2) -> list:
    values = []
    r = first_row
    while (v := sheet.num(r, col)) is not None:
        values.append(int(v))
        r += 1
    return values


def _advanced_settings(sheet) -> list[dict]:
    settings = []
    for r in range(3, sheet.nrows + 1):
        name = sheet.text(r, 9)
        if not name:
            continue
        settings.append({
            "name": name,
            "row": r,
            "label": sheet.text(r, 1),
            "hint": sheet.text(r, 3),
            "default": clean(sheet.cell(r, 8)),
            "current": clean(sheet.cell(r, 2)),
        })
    return settings


def extract(wb) -> dict:
    name_range = wb.sheet("NameRange")
    input_sheet = wb.sheet("Input")
    start = wb.sheet("Start")
    adv = wb.sheet("Adv_settings")

    lists = {key: _numbered_list(name_range, row) for key, row in NAMERANGE_LISTS.items()}

    return {
        "source": "NameRange, Input, Start and Adv_settings",
        "choices": lists,
        "ranges": {
            "birthYears": _column_values(input_sheet, 1),
            "retirementAges": _column_values(input_sheet, 2),
        },
        "normalDefaults": {
            # The Start sheet's yellow input cells, in sheet order.
            "birthYear": clean(start.cell(4, 2)),
            "retirementAge": clean(start.cell(5, 2)),
            "startWorkAge": clean(start.cell(6, 2)),
            "monthlySalary": clean(start.cell(7, 2)),
            # Start!B9 is driven by a checkbox control, so the cell itself is blank;
            # NameRange row 10 records the checkbox's default state.
            "married": bool(name_range.cell(10, 2)),
            "yearlyInflation": clean(start.cell(12, 2)),
            "realGrowth": clean(start.cell(13, 2)),
            "realReturn": clean(start.cell(14, 2)),
        },
        "advancedSettings": _advanced_settings(adv),
    }
