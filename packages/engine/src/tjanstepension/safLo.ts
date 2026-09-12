/**
 * SAF-LO and STP -- the agreement for blue-collar workers in private industry.
 *
 * SAF-LO is defined contribution and replaced STP in 1996. STP lives on as a
 * transitional benefit for cohorts born up to 1967, based on points earned
 * through to 1995.
 *
 * Port of `SAF_LO` and `STP_` from Tjänstepensioner.bas.
 */

import { vbaInt } from "../vba/math.js";
import { BREAKPOINT_IBB } from "./types.js";
import type { SchemeContext } from "./types.js";

/**
 * The age from which premiums are paid, which has been lowered repeatedly.
 */
function entryAge(year: number): number {
  if (year < 2000) return 28;
  if (year < 2021) return 25;
  if (year < 2022) return 24;
  if (year < 2023) return 23;
  return 22;
}

/**
 * Premium rates below and above the breakpoint, by year.
 *
 * The step-up between 2008 and 2013 is the phased introduction of the higher
 * rate above the breakpoint, from 3.5% to 30%.
 */
function premiumRates(year: number, flexPension: number): [below: number, above: number] {
  if (year < 1996) return [0, 0];
  if (year < 2000) return [0.02, 0.02];
  if (year < 2008) return [0.035, 0.035];
  if (year === 2008) return [0.039, 0.06];
  if (year === 2009) return [0.04, 0.12];
  if (year === 2010) return [0.041, 0.18];
  if (year === 2011) return [0.043, 0.24];
  if (year <= 2013) return [0.045, 0.3];
  return [0.045 + flexPension, 0.3 + flexPension];
}

/** SAF-LO premium for a year. */
export function SAF_LO(
  alderIn: number,
  inkomstIn: number,
  ibbIn: number,
  zparIn: number,
  year: number,
  context: SchemeContext,
): number {
  const { born, wStart, tjpPar, flexPension, marginal } = context;

  // `alder As Long` and `IBB As Long` in the VBA: both are rounded on entry.
  const alder = Math.round(alderIn);
  let inkomst = inkomstIn;
  if (inkomst <= 0) return 0;
  if (year < 1996) return 0;

  let zpar = zparIn;
  if (zpar > 65) zpar = 65;

  const lowAge = entryAge(year);
  const ibb = ibbIn === 0 ? Math.round(inkomst) : Math.round(ibbIn);

  const [premie1, premie2] = premiumRates(year, flexPension);

  let month: number;
  if (alder < lowAge) {
    return 0;
  } else if (vbaInt(alder) === lowAge) {
    month = vbaInt((vbaInt(born + lowAge + 1) - born - wStart) * 12);
    if (month > 12) month = 12;
  } else if (alder < vbaInt(zpar)) {
    month = alder > wStart ? 12 : vbaInt((vbaInt(born + wStart + 1) - born - wStart) * 12);
  } else if (alder === vbaInt(zpar)) {
    if (zpar < 65) {
      month = vbaInt((born + zpar - vbaInt(born + zpar)) * 12);
    } else {
      month = vbaInt((born - vbaInt(born)) * 12);
      // QUIRK, the same one as in tlITP1: the guard is Int() of a fraction
      // between 0 and 1, so it never fires. Kept as written.
      if (vbaInt(born + tjpPar - vbaInt(born + tjpPar)) > 0) {
        inkomst = (inkomst * month) / vbaInt((born + tjpPar - vbaInt(born + tjpPar)) * 12);
      }
    }
  } else {
    month = 0;
  }

  if (month === 0) return 0;

  inkomst = inkomst / month;
  // As in ITP 1, the test divides by the months worked and the formula by 12.
  const breakpointTest = (BREAKPOINT_IBB * ibb) / month;
  const breakpoint = (BREAKPOINT_IBB * ibb) / 12;

  let premium =
    inkomst <= breakpointTest
      ? premie1 * inkomst
      : premie2 * (inkomst - breakpoint) + premie1 * breakpoint;

  if (marginal === 0) premium = vbaInt(premium + 0.5);
  return premium * month;
}

/**
 * STP -- the transitional benefit for blue-collar workers, an annual pension
 * rather than a premium.
 *
 * 10% of the price base amount per average point, scaled by years of service
 * against a target that rises with the cohort.
 *
 * SIDE EFFECT IN THE ORIGINAL: `STP_` assigns `born = Int(born)`, truncating the
 * birth-month fraction on the module-level global for everything computed after
 * it in that run. It is called only from `FTJP`, so the truncation is applied
 * there rather than hidden inside this function -- see formansbestamd.ts.
 *
 * @param medel average STP points
 * @param antal years of service through to 1995
 * @param pbb   price base amount in force at 65, or in 1995 if earlier
 */
export function STP_(
  medel: number,
  antalIn: number,
  pbbIn: number,
  born: number,
  tjpPar: number,
  marginal: 0 | 1 = 0,
): number {
  const antal = Math.round(antalIn);
  const pbb = Math.round(pbbIn);
  if (antal < 3) return 0;

  let par = tjpPar;
  // No withdrawal before 65, and no further deferral past 70.
  if (par < 65) par = 65;
  if (par > 70) par = 70;
  const cohort = vbaInt(born);

  let faktor = 1;
  if (par > 65) {
    // DEAD in the original: this branch already requires par > 65.
    if (par < 65) {
      faktor = 0;
    } else if (vbaInt(par) <= 70) {
      faktor = 1 + 0.006 * vbaInt((par - 65) * 12);
    } else {
      faktor = 1 + 0.006 * 60;
    }
  }

  // A 2.5% uplift for the value protection written into the 1996-2000
  // transitional rules.
  if (vbaInt(cohort + par) > 1999) faktor = 1.025 * faktor;

  let targetYears: number;
  if (cohort <= 1937) targetYears = 30;
  else if (cohort <= 1940) targetYears = 32;
  else if (cohort <= 1941) targetYears = 33;
  else if (cohort <= 1942) targetYears = 34;
  else if (cohort <= 1943) targetYears = 35;
  else if (cohort <= 1944) targetYears = 36;
  else targetYears = 37;

  let benefit =
    antal < targetYears
      ? (faktor * medel * pbb * 0.1 * antal) / targetYears
      : faktor * medel * pbb * 0.1;

  // Note the 0.49 here, not the 0.5 used elsewhere -- as written.
  if (marginal === 0) benefit = vbaInt(benefit / 12 + 0.49) * 12;
  return benefit;
}
