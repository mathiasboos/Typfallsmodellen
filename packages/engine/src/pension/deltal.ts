/**
 * Delningstal lookup -- the divisor that turns pension capital into an annuity.
 *
 * Port of aaDeltal.bas. Two sources feed one table, exactly as
 * `subLoadDeltal_IP_PPifneeded` assembles it:
 *
 *   - published values from the Nyckeltal sheet, for every cohort;
 *   - overwritten, for cohorts from `deltalFromMortalityIp` / `...Pp` (1958 by
 *     default), with the model's own unisex figures from `Calculate_Deltal`.
 *
 * The VBA reads those second values from the mortality sheet's cached output.
 * The port computes them instead, which is equivalent -- `calculateDeltal`
 * reproduces that cached table exactly -- and means only the cohort actually
 * being modelled is ever computed, rather than all 121.
 */

import type { ModelContext } from "../model/context.js";
import { vbaInt, vbaRound } from "../vba/math.js";
import type { DeathProbabilities } from "./mortality.js";
import { calculateDeltal } from "./mortality.js";

/** A published cohort-by-age table from the Nyckeltal sheet. */
export interface PublishedTable {
  readonly firstCohort: number;
  readonly lastCohort: number;
  readonly firstAge: number;
  readonly lastAge: number;
  readonly values: readonly (readonly (number | null)[])[];
}

export interface AnnuityTablesData {
  readonly incomePension: PublishedTable;
  readonly premiumPension: PublishedTable;
  readonly bornBefore1938: {
    readonly firstCohort: number;
    readonly lastCohort: number;
    readonly incomePension: readonly (number | null)[];
    readonly premiumPension: readonly (number | null)[];
  };
}

/** Cohorts below this have no delningstal at all; the VBA halts on them. */
export const EARLIEST_COHORT = 1930;
/** Cohorts at or below this get a single divisor rather than an age table. */
export const LAST_OLD_RULES_COHORT = 1937;

function lookupPublished(table: PublishedTable, cohort: number, age: number): number {
  const row = table.values[cohort - table.firstCohort];
  if (row === undefined) return 0;
  return row[age - table.firstAge] ?? 0;
}

/**
 * The assembled delningstal tables for one model run.
 *
 * Holds the published values and, lazily, the model's own figures for whichever
 * cohorts get asked for.
 */
export class DeltalTables {
  private readonly computed = new Map<number, ReturnType<typeof calculateDeltal>>();

  constructor(
    private readonly published: AnnuityTablesData,
    private readonly deaths: DeathProbabilities,
    private readonly context: ModelContext,
  ) {}

  private ownFigures(cohort: number): ReturnType<typeof calculateDeltal> {
    let factors = this.computed.get(cohort);
    if (factors === undefined) {
      factors = calculateDeltal(cohort, this.deaths);
      this.computed.set(cohort, factors);
    }
    return factors;
  }

  /**
   * Income pension delningstal for a cohort at a whole age.
   *
   * The mortality figures override the published ones only inside the published
   * table's own bounds, matching the VBA's guards -- a cohort past the table's
   * last row, or an age past its last column, keeps the published value.
   */
  incomePension(cohort: number, age: number): number {
    const table = this.published.incomePension;
    const from = this.context.deltalFromMortalityIp;
    if (from !== 0 && cohort >= from && cohort <= table.lastCohort && age <= table.lastAge) {
      // Sex 0 is the unisex basis, which is what the model uses.
      const own = this.ownFigures(cohort).dtalip[0]?.[age];
      if (own !== undefined && own > 0) return own;
    }
    return lookupPublished(table, cohort, age);
  }

  /** Premium pension delningstal for a cohort at a whole age. */
  premiumPension(cohort: number, age: number): number {
    const table = this.published.premiumPension;
    const from = this.context.deltalFromMortalityPp;
    if (from !== 0 && cohort >= from && cohort <= table.lastCohort && age <= table.lastAge) {
      const own = this.ownFigures(cohort).dtalpp[0]?.[age];
      if (own !== undefined && own > 0) return own;
    }
    return lookupPublished(table, cohort, age);
  }

  /** The single divisor for cohorts born 1937 or earlier. */
  oldRules(cohort: number, kind: "income" | "premium"): number {
    const old = this.published.bornBefore1938;
    const index = cohort - old.firstCohort;
    const list = kind === "income" ? old.incomePension : old.premiumPension;
    return list[index] ?? 0;
  }
}

/**
 * Rounds to two decimals the way aaDeltal.bas does: `Int(x * 100 + 0.4999) / 100`.
 *
 * Not `vbaRound` -- this rounds a value ending in exactly .xx5 *down*, where
 * banker's rounding would take it to even. The VBA says so in a comment, and the
 * difference is a whole öre on the divisor.
 */
function roundDeltal(value: number): number {
  return vbaInt(value * 100 + 0.4999) / 100;
}

/**
 * Interpolates between the divisors either side of a part-year age.
 *
 * Someone retiring part-way through a year draws on both. The weighting is by
 * whole months, not by the raw fraction.
 */
function blend(atAge: number, atNextAge: number, curAge: number): number {
  const months = vbaRound((curAge - vbaInt(curAge)) * 12, 0);
  return ((12 - months) / 12) * atAge + (months / 12) * atNextAge;
}

function deltalFor(
  born: number,
  curAge: number,
  adjustProportional: boolean,
  tables: DeltalTables,
  kind: "income" | "premium",
): number {
  if (vbaInt(born) !== born) {
    throw new RangeError(`delningstal: birth year must be whole, got ${born}`);
  }
  if (vbaInt(born) < EARLIEST_COHORT) {
    throw new RangeError(`delningstal: no figures for cohorts before ${EARLIEST_COHORT}`);
  }

  // Cohorts born 1937 or earlier get one divisor, with no month handling.
  if (vbaInt(born) <= LAST_OLD_RULES_COHORT) return tables.oldRules(vbaInt(born), kind);

  let age = curAge;
  let ageFraction = age - vbaInt(age);
  // Retiring on a whole birthday is taken to mean mid-year.
  if (adjustProportional && ageFraction === 0) {
    age += 0.5;
    ageFraction = 0.5;
  }

  const at = (a: number) =>
    kind === "income" ? tables.incomePension(born, a) : tables.premiumPension(born, a);

  const value =
    ageFraction > 0 ? blend(at(vbaInt(age)), at(vbaInt(age) + 1), age) : at(age);
  return roundDeltal(value);
}

/** Income pension delningstal. Mirrors `fnDeltal_IP`. */
export function fnDeltalIp(
  born: number,
  curAge: number,
  tables: DeltalTables,
  adjustProportional = false,
): number {
  return deltalFor(born, curAge, adjustProportional, tables, "income");
}

/** Premium pension delningstal. Mirrors `fnDeltal_PP`. */
export function fnDeltalPp(
  born: number,
  curAge: number,
  tables: DeltalTables,
  adjustProportional = false,
): number {
  return deltalFor(born, curAge, adjustProportional, tables, "premium");
}

/**
 * `fnDeltal_IP2`: blends the neighbouring divisors when this is the year
 * pension is first drawn, or the year a partial withdrawal becomes final.
 */
export function fnDeltalIp2(
  born: number,
  curAge: number,
  tables: DeltalTables,
  par = 0,
  defAr = 0,
): number {
  // Past the final withdrawal age the divisor stops moving.
  const age = curAge > defAr ? defAr + 1 : curAge;
  const adjust = vbaInt(par) === age || vbaInt(defAr) === age;
  return fnDeltalIp(born, age, tables, adjust);
}

/**
 * `fnDeltal_PP2`. Same as its income-pension twin except that it does not clamp
 * the age at the final withdrawal year -- faithful to the VBA, which omits that
 * line here.
 */
export function fnDeltalPp2(
  born: number,
  curAge: number,
  tables: DeltalTables,
  par = 0,
  defAr = 0,
): number {
  const adjust = vbaInt(par) === curAge || vbaInt(defAr) === curAge;
  return fnDeltalPp(born, curAge, tables, adjust);
}
