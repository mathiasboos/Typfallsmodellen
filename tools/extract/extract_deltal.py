"""Published delningstal (annuity divisors) from the Nyckeltal sheet.

`fnDeltal_IP` and `fnDeltal_PP` in aaDeltal.bas do *not* use the annuity factors
the model computes in Calculate_Deltal. They read published tables from
Nyckeltal, and then overwrite them from the mortality sheet's computed output
for cohorts at or after `rngDelnIPMort` / `rngDelnPPMort` (both 1958 by default).
So both sources are live, and the engine needs this one as well as its own port
of Calculate_Deltal.

Extents are found by walking the header row and column for contiguity, the way
the VBA's `End(xlToRight)` and `End(xlDown)` do, rather than being hard-coded --
so a release that adds ages or cohorts is picked up automatically.
"""

from __future__ import annotations

from common import exact

SHEET = "Nyckeltal"

# Anchors, matching the named ranges aaDeltal.bas resolves:
#   rng_Delningtal_FirstAge   ages along row 4 from column 4
#   rng_Delningtal_FirstYear  cohorts down column 3 from row 5
#   rng_PP_deltal_Year        cell (4, 27); ages start one column right,
#                             cohorts one row down
#   rng_Nyckeltal_Deltal_IP   cell (8, 74); the pre-1938 block starts at (9, 73)
IP_AGE_ROW, IP_AGE_FIRST_COL = 4, 4
IP_COHORT_COL, IP_COHORT_FIRST_ROW = 3, 5
PP_ANCHOR_ROW, PP_ANCHOR_COL = 4, 27
OLD_COHORT_COL, OLD_FIRST_ROW = 73, 9
OLD_IP_COL, OLD_PP_COL = 74, 75


def _walk_right(sheet, row: int, first_col: int) -> int:
    col = first_col
    while sheet.num(row, col) is not None:
        col += 1
    return col - 1


def _walk_down(sheet, col: int, first_row: int) -> int:
    row = first_row
    while sheet.num(row, col) is not None:
        row += 1
    return row - 1


def _read_table(sheet, age_row: int, age_first_col: int, cohort_col: int, cohort_first_row: int) -> dict:
    last_col = _walk_right(sheet, age_row, age_first_col)
    last_row = _walk_down(sheet, cohort_col, cohort_first_row)

    first_age = int(sheet.num(age_row, age_first_col))
    last_age = int(sheet.num(age_row, last_col))
    first_cohort = int(sheet.num(cohort_first_row, cohort_col))
    last_cohort = int(sheet.num(last_row, cohort_col))

    values = []
    for row in range(cohort_first_row, last_row + 1):
        values.append(
            [exact(sheet.num(row, col)) for col in range(age_first_col, last_col + 1)]
        )

    return {
        "firstCohort": first_cohort,
        "lastCohort": last_cohort,
        "firstAge": first_age,
        "lastAge": last_age,
        "values": values,  # [cohort][age]
    }


def extract(wb) -> dict:
    sheet = wb.sheet(SHEET)

    income = _read_table(sheet, IP_AGE_ROW, IP_AGE_FIRST_COL, IP_COHORT_COL, IP_COHORT_FIRST_ROW)
    premium = _read_table(
        sheet, PP_ANCHOR_ROW, PP_ANCHOR_COL + 1, PP_ANCHOR_COL, PP_ANCHOR_ROW + 1
    )

    # Cohorts born 1937 or earlier get a single divisor each, not an age table.
    last_old_row = _walk_down(sheet, OLD_COHORT_COL, OLD_FIRST_ROW)
    old_cohorts = {
        "firstCohort": int(sheet.num(OLD_FIRST_ROW, OLD_COHORT_COL)),
        "lastCohort": int(sheet.num(last_old_row, OLD_COHORT_COL)),
        "incomePension": [
            exact(sheet.num(r, OLD_IP_COL)) for r in range(OLD_FIRST_ROW, last_old_row + 1)
        ],
        "premiumPension": [
            exact(sheet.num(r, OLD_PP_COL)) for r in range(OLD_FIRST_ROW, last_old_row + 1)
        ],
    }

    return {
        "source": f"{SHEET}: published delningstal tables read by aaDeltal.bas",
        "note": (
            "For cohorts at or after mortalityOverrideFrom the engine replaces these with its own "
            "Calculate_Deltal output, as subLoadDeltal_IP_PPifneeded does."
        ),
        "incomePension": income,
        "premiumPension": premium,
        "bornBefore1938": old_cohorts,
    }
