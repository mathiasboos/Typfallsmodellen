/**
 * Socialstyrelsen's riksnorm for försörjningsstöd -- social assistance.
 *
 * Port of `bistOld` (income years 1985-2005) and `bist25` (2006 onwards) from
 * Bidrag.bas. The two are the same calculation over different tables, and Mcalc
 * picks between them by year.
 *
 * The norm is a monthly cost per person -- one figure per child age band, one
 * per adult, and a shared household amount that grows with family size -- plus
 * the rent, which the model assumes is reasonable. Anything the household's own
 * disposable income does not cover is paid.
 *
 * The amounts live in packages/data/riksnorm.json, parsed straight out of the
 * VBA by tools/extract/extract_riksnorm.py rather than retyped; `npm run
 * check:riksnorm` verifies they still match. The rules around them are here.
 */

import riksnormJson from "../../../data/riksnorm.json" with { type: "json" };
import { vbaArray } from "../vba/ageArray.js";
import { vbaInt } from "../vba/math.js";
import type { RiksnormContext } from "./types.js";

interface Branch {
  readonly op: string;
  readonly year?: number;
  readonly values: readonly number[];
}

const TABLES = riksnormJson.tables as Record<string, Record<string, readonly Branch[]>>;

/**
 * One row of a table, walking the same If/ElseIf chain the VBA does.
 *
 * The branches are stored in source order with the comparison that guards them,
 * so a year that matches two conditions takes the first, exactly as VBA would.
 */
function row(fn: "bistOld" | "bist25", family: "xn" | "vuxna" | "Gn", year: number): number[] {
  for (const branch of TABLES[fn]![family]!) {
    const matches =
      branch.op === "else" ||
      (branch.op === "eq" && year === branch.year) ||
      (branch.op === "ne" && year !== branch.year) ||
      (branch.op === "le" && year <= branch.year!) ||
      (branch.op === "lt" && year < branch.year!) ||
      (branch.op === "ge" && year >= branch.year!) ||
      (branch.op === "gt" && year > branch.year!);
    if (matches) return [...branch.values];
  }
  throw new RangeError(`riksnorm: no ${fn}.${family} row for ${year}`);
}

/**
 * Social assistance under the 1985-2005 tables.
 *
 * NOTE: `wage` is declared but never read here -- it only matters from 2013, in
 * `bist25`. Kept so both functions take the same arguments.
 *
 * @param civ     Number of adults in the household, 1 or 2.
 * @param hyra    Monthly housing cost, assumed reasonable.
 * @param disp    The household's annual disposable income.
 * @param b1      Children aged 0, then 1-2, 3, 4-6, 7-10, 11-14, 15-18, 19-20.
 * @param wage    Net wage; never read.
 * @param context `marginal`.
 * @param year    Income year.
 */
export function bistOld(
  civ: number,
  hyra: number,
  disp: number,
  b1: number,
  b2: number,
  b3: number,
  b4: number,
  b5: number,
  b6: number,
  b7: number,
  b8: number,
  wage: number,
  context: RiksnormContext,
  year = 2005,
): number {
  let bistand = 0;
  if (year < 1985 || year > 2005) return bistand;

  const { marginal } = context;

  // Personal costs for children at home, by age band.
  const xnValues = row("bistOld", "xn", year);
  // Medical and dental care comes out of the norm from 1994.
  if (year === 1994) for (let i = 0; i < 8; i += 1) xnValues[i] = xnValues[i]! - 32;
  if (year === 1995) for (let i = 0; i < 8; i += 1) xnValues[i] = xnValues[i]! - 33;
  if (year === 1996) for (let i = 0; i < 8; i += 1) xnValues[i] = xnValues[i]! - (33 * 9) / 12;
  const xn = vbaArray(...xnValues);

  const barn = vbaArray(b1, b2, b3, b4, b5, b6, b7, b8);

  let antal = 0;
  for (let i = 1; i <= 8; i += 1) {
    bistand = bistand + xn(i) * barn(i);
    antal = antal + barn(i);
  }

  // Household members, adults included.
  antal = antal + civ;

  // Personal costs for the adults: single, then cohabiting.
  const vuxnaValues = row("bistOld", "vuxna", year);
  if (year === 1996) {
    vuxnaValues[0] = vuxnaValues[0]! - (57 * 9) / 12;
    vuxnaValues[1] = vuxnaValues[1]! - (116 * 9) / 12;
  }
  const vuxna = vbaArray(...vuxnaValues);

  bistand = bistand + (civ === 1 ? vuxna(1) : vuxna(2));

  // Shared household costs, by family size. A medium-sized town is assumed.
  const Gn = vbaArray(...row("bistOld", "Gn", year));
  const flera = Gn(7) - Gn(6);

  if (antal > 7) {
    bistand = bistand + Gn(7) + flera * (antal - 7);
  } else {
    bistand = bistand + Gn(antal);
  }

  // Assumed to follow CPI when projected; `bist25` does index it, this does not.
  const just = 1;
  bistand = bistand * just + hyra;

  bistand = vbaInt(bistand + 0.5);

  /**
   * NOTE: `disp` is an annual figure -- Mcalc passes `IndDisp(age)` -- and
   * `bistand` at this point is a monthly one. `bist25` writes the same test as
   * `disp < bistand * 12`; this one does not, so the household has to fall
   * below one month's subsistence need in a whole year's income before
   * anything is paid. Kept as written.
   */
  if (disp < bistand) {
    bistand = 12 * bistand - disp;
    if (marginal === 0) bistand = vbaInt(bistand + 0.5);
  } else {
    bistand = 0;
  }

  return bistand;
}

/**
 * Social assistance under the 2006 and later tables, as an annual amount.
 *
 * @param civ     Number of adults in the household, 1 or 2.
 * @param hyra    Monthly housing cost, assumed reasonable.
 * @param disp    The household's annual disposable income.
 * @param b1      Children aged 0, then 1-2, 3, 4-6, 7-10, 11-14, 15-18, 19-20.
 * @param wage    Net wage; a quarter of it is disregarded from 2013.
 * @param context `marginal`, and the cohort and KPI vector for the indexation.
 * @param year    Income year.
 */
export function bist25(
  civ: number,
  hyra: number,
  disp: number,
  b1: number,
  b2: number,
  b3: number,
  b4: number,
  b5: number,
  b6: number,
  b7: number,
  b8: number,
  wage: number,
  context: RiksnormContext,
  year = 2025,
): number {
  let bistand = 0;
  // Before 2005 the older tables apply; see `bistOld`.
  if (year < 2005) return bistand;

  const { marginal, born, age, vectors } = context;

  const xn = vbaArray(...row("bist25", "xn", year));
  const barn = vbaArray(b1, b2, b3, b4, b5, b6, b7, b8);

  let antal = 0;
  for (let i = 1; i <= 8; i += 1) {
    bistand = bistand + xn(i) * barn(i);
    antal = antal + barn(i);
  }

  antal = antal + civ;

  const vuxna = vbaArray(...row("bist25", "vuxna", year));
  bistand = bistand + (civ === 1 ? vuxna(1) : vuxna(2));

  const Gn = vbaArray(...row("bist25", "Gn", year));
  const flera = Gn(7) - Gn(6);

  if (antal > 7) {
    bistand = bistand + Gn(7) + flera * (antal - 7);
  } else {
    bistand = bistand + Gn(antal);
  }

  // Past the last tabulated year the whole norm follows CPI.
  let just = 1;
  if (year > 2025) just = vectors.kpi(age) / vectors.kpi(2025 - vbaInt(born));

  bistand = bistand * just + hyra;

  // A quarter of earned income is disregarded, for a recipient of six months
  // or more. Mcalc passes 0, so this is inert there.
  if (year >= 2013) disp = disp - wage * 0.25;

  if (disp < bistand * 12) {
    bistand = 12 * bistand - disp;
    if (marginal === 0) bistand = vbaInt(bistand + 0.5);
  } else {
    bistand = 0;
  }

  return bistand;
}
