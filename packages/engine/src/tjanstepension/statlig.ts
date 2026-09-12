/**
 * PA16 -- the agreement for central government employees.
 *
 * Avdelning 1 covers those born 1988 or later and is a single defined
 * contribution premium (`tlPA16`). Avdelning 2 covers earlier cohorts and is
 * built from two premiums paid side by side: the collectively agreed Kåpan
 * (`Kapan`) and the individually chosen part (`PA_indiv`), which Mcalc sums.
 *
 * Port of `tlPA16`, `Kapan` and `PA_indiv` from Tjänstepensioner.bas. `tlPA03`,
 * the defined benefit that goes with avdelning 2, lives with the rest of the
 * benefit layer in formansbestamd.ts.
 */

import { vbaInt } from "../vba/math.js";
import { BREAKPOINT_IBB } from "./types.js";
import type { SchemeContext } from "./types.js";

/** Contributions are capped at 30 income base amounts of salary from 2003. */
const SALARY_CEILING_IBB = 30;

/**
 * PA16 avdelning 1 premium.
 *
 * 6% below the breakpoint and 31.5% above it, rising slightly as the flex part
 * of the agreement steps up through 2026 and 2027.
 */
export function tlPA16(
  alder: number,
  inkomstIn: number,
  ibbIn: number,
  zparIn: number,
  year: number,
  context: SchemeContext,
): number {
  const { born, wStart, tjpPar, marginal } = context;

  let inkomst = inkomstIn;
  if (inkomst <= 0) return 0;

  // No rights accrue past the LAS age.
  const maxAge = year > 2003 ? 67 : 65;
  const zpar = zparIn > maxAge ? maxAge : zparIn;

  let month: number;
  if (alder < wStart) {
    return 0;
  } else if (vbaInt(alder) === wStart) {
    month = vbaInt((vbaInt(born + wStart + 1) - born - wStart) * 12);
    if (month > 12) month = 12;
  } else if (alder < vbaInt(zpar)) {
    month = 12;
  } else if (alder === vbaInt(zpar)) {
    if (zpar < maxAge) {
      month = vbaInt((born + zpar - vbaInt(born + zpar)) * 12);
    } else {
      month = vbaInt((born - vbaInt(born)) * 12);
      // QUIRK, as in tlITP1 and SAF_LO: Int() of a fraction below 1, so never.
      if (vbaInt(born + tjpPar - vbaInt(born + tjpPar)) > 0) {
        inkomst = (inkomst * month) / vbaInt((born + tjpPar - vbaInt(born + tjpPar)) * 12);
      }
    }
  } else {
    month = 0;
  }

  const ibb = ibbIn === 0 ? inkomst : ibbIn;
  if (month === 0) return 0;

  inkomst = inkomst / month;
  const breakpointTest = (BREAKPOINT_IBB * ibb) / month;
  const breakpoint = (BREAKPOINT_IBB * ibb) / 12;

  // The flex part rises mid-year in 2026, hence the nine-and-three blend.
  let below: number;
  let above: number;
  if (year < 2026) {
    below = 0.06;
    above = 0.315;
  } else if (year === 2026) {
    below = (0.061 * 9 + 0.062 * 3) / 12;
    above = (0.316 * 9 + 0.317 * 3) / 12;
  } else {
    below = 0.062;
    above = 0.317;
  }

  let premium =
    inkomst <= breakpointTest
      ? below * inkomst
      : above * (inkomst - breakpoint) + below * breakpoint;

  if (marginal === 0) premium = vbaInt(premium + 0.5);
  return premium * month;
}

/**
 * Months worked, shared by the two avdelning 2 premiums.
 *
 * Returns the months and, separately, the divisor to use in the retirement
 * year -- where both functions divide by the months of the retirement year
 * rather than by twelve.
 */
function avd2Months(
  alder: number,
  entryAge: number,
  zpar: number,
  highAge: number,
  context: SchemeContext,
): { month: number; divisor: number } | null {
  const { born, wStart, tjpPar } = context;

  if (alder < entryAge) return null;

  if (alder === entryAge) {
    let month = vbaInt((vbaInt(born + entryAge + 1) - born - wStart) * 12);
    if (month > 12) month = 12;
    return { month, divisor: 12 };
  }

  if (alder < vbaInt(zpar)) {
    const month =
      alder > wStart ? 12 : vbaInt((vbaInt(born + wStart + 1) - born - wStart) * 12);
    return { month, divisor: 12 };
  }

  if (alder === vbaInt(zpar)) {
    if (zpar < highAge) {
      return { month: vbaInt((born + zpar - vbaInt(born + zpar)) * 12), divisor: 12 };
    }
    const month = vbaInt((born - vbaInt(born)) * 12);
    // Unlike the ITP-style guard, this one has the `* 12` inside and does fire.
    // VBA treats the bare numeric as a condition, i.e. non-zero.
    const retirementMonths = vbaInt((born + tjpPar - vbaInt(born + tjpPar)) * 12);
    if (retirementMonths === 0) return null;
    return { month, divisor: retirementMonths };
  }

  return null;
}

/**
 * Kåpan -- the collectively agreed premium in PA16 avdelning 2.
 *
 * Paid from 1991, at 2% of salary today. No rounding is applied.
 */
export function Kapan(
  year: number,
  fodar: number,
  ibb: number,
  wageIn: number,
  context: SchemeContext,
  zparIn = 65,
): number {
  if (wageIn <= 0) return 0;

  const alder = year - vbaInt(fodar);
  if (alder > vbaInt(zparIn)) return 0;

  let prem: number;
  if (year < 1991) prem = 0;
  else if (year <= 1993) prem = 0.013;
  else if (year === 1994) prem = 0.015;
  else if (year <= 2002) prem = 0.019;
  else prem = 0.02;

  let wage = wageIn;
  if (year > 2002 && wage > SALARY_CEILING_IBB * ibb) wage = SALARY_CEILING_IBB * ibb;

  // The qualifying age was 28 until 2007 and 23 from 2008.
  const lowAge = year <= 2007 ? 28 : 23;
  const highAge = year > 2023 ? 69 : 65;
  const zpar = zparIn > highAge ? highAge : zparIn;

  const worked = avd2Months(alder, lowAge, zpar, highAge, context);
  if (worked === null) return 0;
  return (wage * prem * worked.month) / worked.divisor;
}

/**
 * The individually chosen premium in PA16 avdelning 2.
 *
 * Paid from 2003. The flex part added from 2024 is higher for those born 1965
 * or later.
 */
export function PA_indiv(
  year: number,
  fodar: number,
  ibb: number,
  wageIn: number,
  context: SchemeContext,
  zparIn = 67,
): number {
  if (wageIn <= 0) return 0;

  const alder = year - vbaInt(fodar);
  // Note this compares against the raw retirement age, where Kapan truncates it.
  if (alder > zparIn) return 0;

  let prem: number;
  if (year < 2003) prem = 0;
  else if (year < 2008) prem = 0.023;
  else prem = 0.025;

  const bornBefore1965 = vbaInt(fodar) < 1965;
  if (year > 2023 && year < 2026) {
    prem = bornBefore1965 ? 0.03 : 0.04;
  } else if (year === 2026) {
    // The flex part rises from October, hence the nine-and-three blend.
    prem = bornBefore1965 ? (0.031 * 9 + 0.032 * 3) / 12 : (0.041 * 9 + 0.042 * 3) / 12;
  } else if (year > 2026) {
    prem = bornBefore1965 ? 0.032 : 0.042;
  }

  let wage = wageIn;
  if (year > 2002 && wage > SALARY_CEILING_IBB * ibb) wage = SALARY_CEILING_IBB * ibb;

  const highAge = year > 2023 ? 69 : 65;
  const zpar = zparIn > highAge ? highAge : zparIn;

  const worked = avd2Months(alder, 23, zpar, highAge, context);
  if (worked === null) return 0;
  return (wage * prem * worked.month) / worked.divisor;
}
