/**
 * ITP -- the agreement for salaried employees in private industry and trade.
 *
 * ITP 1 is defined contribution, for those born 1979 or later. ITP 2 covers
 * earlier cohorts and is mostly defined benefit, with a small defined
 * contribution part (ITPK) alongside it.
 *
 * Port of `tlITP1`, `tlITP2A` and `tlITP2F` from Tjänstepensioner.bas.
 */

import { vbaInt } from "../vba/math.js";
import { BREAKPOINT_IBB } from "./types.js";
import type { SchemeContext } from "./types.js";

/** ITP 1 pays no premium before 25. */
const ITP1_ENTRY_AGE = 25;
/** ITPK pays no premium before 28. */
const ITPK_ENTRY_AGE = 28;

/**
 * Months of premium payment in a year, shared by ITP 1 and ITPK.
 *
 * A full year once established, part years in the year the entry age is
 * reached, the year work started, and the year of retirement.
 */
function premiumMonths(
  alder: number,
  entryAge: number,
  zpar: number,
  context: SchemeContext,
): number {
  const { born, wStart } = context;

  if (alder < entryAge) return 0;

  if (vbaInt(alder) === entryAge) {
    // Months since the entry-age birthday, capped at a full year.
    let month = vbaInt((vbaInt(born + entryAge + 1) - born - wStart) * 12);
    if (month > 12) month = 12;
    if (month < 0) month = 0;
    return month;
  }

  if (alder < vbaInt(zpar)) {
    if (alder > wStart) return 12;
    return vbaInt((vbaInt(born + wStart + 1) - born - wStart) * 12);
  }

  if (alder === vbaInt(zpar)) {
    if (zpar < 65) return vbaInt((born + zpar - vbaInt(born + zpar)) * 12);
    return vbaInt((born - vbaInt(born)) * 12);
  }

  return 0;
}

/**
 * ITP 1 premium for a year.
 *
 * 4.5% of salary below the 7.5 IBB breakpoint and 30% above it, plus whatever
 * flexpension premium is set, from 2014.
 */
export function tlITP1(
  alder: number,
  inkomstIn: number,
  ibbIn: number,
  zparIn: number,
  context: SchemeContext,
): number {
  const { year, born, tjpPar, flexPension, marginal } = context;

  let zpar = zparIn;
  // No pension rights accrue after 65 -- 66 once the retirement ages rose.
  if (zpar > 65) zpar = 65;
  if (year > 2022) zpar = 66;

  let inkomst = inkomstIn;
  if (inkomst <= 0) return 0;

  const month = premiumMonths(alder, ITP1_ENTRY_AGE, zpar, context);

  // In the retirement year at 65 or later the salary would be rescaled to the
  // months actually worked -- except that this never fires.
  //
  // QUIRK: the guard is `Int(born + tjp_par - Int(born + tjp_par)) > 0`, and
  // that expression is a *fraction* of a year between 0 and 1, so Int() of it is
  // always 0. ITPK's version of the same guard multiplies by 12 inside the Int()
  // and does fire (see tlITP2A below), which is what makes this look like a
  // dropped `* 12` rather than a deliberate difference. Kept as written.
  if (alder === vbaInt(zpar) && zpar >= 65) {
    const retirementMonths = vbaInt((born + tjpPar - vbaInt(born + tjpPar)) * 12);
    if (vbaInt(born + tjpPar - vbaInt(born + tjpPar)) > 0) {
      inkomst = (inkomst * month) / retirementMonths;
    }
  }

  const ibb = ibbIn === 0 ? inkomst : ibbIn;
  if (month === 0) return 0;

  // Rates apply to the monthly salary.
  inkomst = inkomst / month;
  // Note the two different divisors, both as written in the VBA: the *test*
  // divides the breakpoint by the months worked, so it compares annual figures,
  // while the premium formula divides by 12. They agree for a full year and part
  // company in the entry and retirement years.
  const breakpointTest = (BREAKPOINT_IBB * ibb) / month;
  const breakpoint = (BREAKPOINT_IBB * ibb) / 12;

  let premium: number;
  if (inkomst <= breakpointTest) {
    premium = 0.045 * inkomst;
    if (year > 2013) premium = (0.045 + flexPension) * inkomst;
  } else {
    // QUIRK: meant to cap the premium at 30 income base amounts from 2023, but
    // `inkomst` is monthly here while `30 * IBB` is annual, so it only bites
    // above a *monthly* salary of 30 IBB -- around 360 base amounts a year.
    // Effectively inert. Reads like a dropped `/ 12`; kept as written.
    if (year > 2022 && inkomst > 30 * ibb) inkomst = 30 * ibb;
    premium = 0.3 * (inkomst - breakpoint) + 0.045 * breakpoint;
    if (year > 2013) {
      premium =
        (0.3 + flexPension) * (inkomst - breakpoint) + (0.045 + flexPension) * breakpoint;
    }
  }

  if (marginal === 0) premium = vbaInt(premium + 0.5);
  return premium * month;
}

/**
 * ITPK premium -- the defined contribution part of ITP 2.
 *
 * A flat 2% of salary, with no breakpoint. Could not be earned before 1997.
 */
export function tlITP2A(
  alder: number,
  inkomstIn: number,
  zparIn: number,
  context: SchemeContext,
): number {
  const { year, born, tjpPar, marginal } = context;

  if (year < 1997) return 0;
  let inkomst = inkomstIn;
  if (inkomst <= 0) return 0;

  let zpar = zparIn;
  if (zpar > 65) zpar = 65;

  const month = premiumMonths(alder, ITPK_ENTRY_AGE, zpar, context);

  // Unlike ITP 1's, this guard has the `* 12` inside the Int() and does fire.
  if (alder === vbaInt(zpar) && zpar >= 65) {
    const retirementMonths = vbaInt((born + tjpPar - vbaInt(born + tjpPar)) * 12);
    if (retirementMonths > 0) inkomst = (inkomst * month) / retirementMonths;
  }

  if (month === 0) return 0;

  let premium = (0.02 * inkomst * month) / 12;
  if (marginal === 0) premium = vbaInt(premium + 0.5);
  return premium;
}

/**
 * ITP 2's defined benefit -- an annual pension, not a premium.
 *
 * 10% of final salary below 7.5 IBB, 65% between 7.5 and 20, 32.5% between 20
 * and 30, nothing above. Scaled down below thirty years of service, then
 * adjusted for retiring early or late.
 *
 * @param year number of years of service, not a calendar year
 */
export function tlITP2F(
  inkomst: number,
  ibb: number,
  context: SchemeContext,
  year = 30,
): number {
  if (inkomst <= 0) return 0;
  const { tjpPar, marginal } = context;

  let benefit: number;
  if (inkomst <= 7.5 * ibb) {
    benefit = 0.1 * inkomst;
  } else if (inkomst <= 20 * ibb) {
    benefit = 0.65 * (inkomst - 7.5 * ibb) + 0.1 * 7.5 * ibb;
  } else if (inkomst <= 30 * ibb) {
    benefit = 0.325 * (inkomst - 20 * ibb) + 0.65 * (10 - 7.5) * ibb + 0.1 * 7.5 * ibb;
  } else {
    benefit = 0.325 * (30 - 20) * ibb + 0.65 * (10 - 7.5) * ibb + 0.1 * 7.5 * ibb;
  }

  // Short of thirty years of service, reduced in proportion.
  if (year < 30) benefit = (benefit * year) / 30;

  // Note this uses 0.6% a month for deferral, where tilläggspension uses 0.7%.
  const zpar = tjpPar;
  let faktor: number;
  if (vbaInt(zpar) < 64) {
    faktor = 1 - 0.005 * vbaInt((65 - zpar) * 12);
  } else if (vbaInt(zpar) === 64) {
    faktor = 1 - 0.005 * vbaInt((65 - zpar) * 12);
  } else if (vbaInt(zpar) <= 70) {
    faktor = 1 + 0.006 * vbaInt((zpar - 65) * 12);
  } else {
    faktor = 1 + 0.006 * 60;
  }

  benefit *= faktor;
  if (marginal === 0) benefit = vbaInt(benefit / 12 + 0.5) * 12;
  return benefit;
}
