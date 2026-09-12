/**
 * The lowest pension age and the riktålder, per cohort.
 *
 * `Rng_riktL` and `Rng_riktage` -- two cells on the Nyckeltal sheet that
 * `startsetup` and `Mcalc` read rather than compute (VBA_go.bas 155, 781-782).
 * They are looked up **by cohort**, which the author says twice: "för aktuell
 * årskull" beside the first, and "Nedan avseende årskull men lagstiftningen ser
 * till inkomstår" beside the second.
 *
 * This is not the same thing as `riktage(year, typ)` in contributions.ts, which
 * is keyed on the income year. The two disagree for about half the cohorts --
 * a 1970 typfall's lowest age is 65 here, where `riktage(1970 + 64, 0)` gives
 * 64 -- so they are not interchangeable. The function is right only where the
 * model pins an expenditure-rule year (VBA_go.bas 161 and 789); everywhere else
 * the cohort table is what the workbook uses.
 */

import retirementAgesJson from "../../../data/retirement-ages.json" with { type: "json" };
import { vbaInt } from "../vba/math.js";

interface RetirementAgeTable {
  readonly firstCohort: number;
  readonly lastCohort: number;
  readonly lowest: readonly number[];
  readonly riktalder: readonly number[];
}

const TABLE = retirementAgesJson as RetirementAgeTable;

/**
 * Reads one cohort's row, clamping to the ends of the table.
 *
 * The sheet runs 1930-2057 and the Start sheet's dropdown stays well inside
 * that, so the clamp is for callers the workbook has no answer for at all --
 * where its own lookup would return an error rather than a number.
 */
function at(values: readonly number[], born: number): number {
  const cohort = vbaInt(born);
  const index = Math.min(Math.max(cohort, TABLE.firstCohort), TABLE.lastCohort) - TABLE.firstCohort;
  return values[index]!;
}

/** `Rng_riktL`: the earliest age this cohort may draw public pension. */
export function lowestPensionAge(born: number): number {
  return at(TABLE.lowest, born);
}

/** `Rng_riktage`: this cohort's riktålder, which the basic protection follows. */
export function riktalderFor(born: number): number {
  return at(TABLE.riktalder, born);
}

/** The cohorts the table covers. */
export const COHORT_RANGE = {
  first: TABLE.firstCohort,
  last: TABLE.lastCohort,
} as const;
