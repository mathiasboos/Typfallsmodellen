/**
 * KAP-KL and AKAP-KR -- the agreements for municipal and regional employees.
 *
 * KAP-KL is the older agreement; AKAP-KL replaced it from 2014 for those born
 * 1986 or later, and AKAP-KR from 2023 for everyone who chose to switch.
 *
 * Port of `tlkap_kl` and `tlakap_kr` from Tjänstepensioner.bas. The two read
 * almost identically and differ in ways that are easy to miss, so the
 * differences are called out at each site rather than factored away.
 *
 * The defined-benefit companions -- `KAPKL_f`, `PA_KL`, `PA_KLBPP` -- live with
 * the rest of the benefit layer in formansbestamd.ts.
 */

import type { ModelContext } from "../model/context.js";
import { vbaInt } from "../vba/math.js";
import { BREAKPOINT_IBB } from "./types.js";
import type { SchemeContext } from "./types.js";

/** Contributions are capped at 30 income base amounts of salary. */
const SALARY_CEILING_IBB = 30;

/** Result of the shared month calculation: months worked and the scaled salary. */
interface MonthsWorked {
  readonly month: number;
  readonly inkomst: number;
}

/**
 * Months of premium payment, and the salary scaled to them.
 *
 * Shared by both agreements from age 21 upward. What differs between them is
 * only what happens *below* 21, which each caller handles itself.
 */
function monthsFrom21(
  alder: number,
  inkomstIn: number,
  zpar: number,
  highAge: number,
  context: SchemeContext,
): MonthsWorked {
  const { born, wStart } = context;
  let inkomst = inkomstIn;
  let month: number;

  if (alder === 21) {
    month = vbaInt((vbaInt(born + 21 + 1) - born - wStart) * 12);
    if (month > 12) month = 12;
    if (month < 0) month = 0;
    inkomst = (inkomst * month) / 12;
  } else if (alder < vbaInt(zpar)) {
    if (alder > wStart) {
      month = 12;
    } else {
      month = vbaInt((vbaInt(born + wStart + 1) - born - wStart) * 12);
      inkomst = (inkomst * month) / 12;
    }
  } else if (alder === vbaInt(zpar)) {
    month =
      zpar < highAge
        ? vbaInt((born + zpar - vbaInt(born + zpar)) * 12)
        : vbaInt((born - vbaInt(born)) * 12);
    inkomst = (inkomst * month) / 12;
  } else {
    month = 0;
  }

  return { month, inkomst };
}

/** Applies the two-rate premium formula both agreements share. */
function premiumFromRates(
  inkomstIn: number,
  ibb: number,
  par1: number,
  par2: number,
  marginal: 0 | 1,
): number {
  // Pensionable salary stops at 30 income base amounts.
  const inkomst = inkomstIn > SALARY_CEILING_IBB * ibb ? SALARY_CEILING_IBB * ibb : inkomstIn;
  const breakpoint = BREAKPOINT_IBB * ibb;

  const premium =
    inkomst <= breakpoint ? par1 * inkomst : par2 * (inkomst - breakpoint) + par1 * breakpoint;

  return marginal === 0 ? vbaInt(premium + 0.5) : premium;
}

/** KAP-KL premium for a year. */
export function tlkap_kl(
  alderIn: number,
  inkomstIn: number,
  ibbIn: number,
  zparIn: number,
  year: number,
  context: SchemeContext,
): number {
  const { marginal } = context;
  const alder = Math.round(alderIn);

  // PA-KL applied before this.
  if (year < 1998) return 0;
  // NOTE this early exit: it is what makes the `born > 1985 && alder < 21`
  // branch further down unreachable here, where the same branch in AKAP-KR --
  // which has no early exit -- does fire. Kept as written in both.
  if (alder < 21) return 0;

  // Rights accrue to the LAS age of 67 from 2003, 65 before that.
  const highAge = year > 2002 ? 67 : 65;
  const zpar = zparIn > highAge ? highAge : zparIn;

  const { month: monthsWorked, inkomst: scaled } = monthsFrom21(
    alder,
    inkomstIn,
    zpar,
    highAge,
    context,
  );
  let month = monthsWorked;

  const ibb = ibbIn === 0 ? Math.round(scaled) : Math.round(ibbIn);
  if (month === 0) return 0;

  let par1 = 0;
  let par2 = 0;
  if (year < 1998) {
    // Unreachable: the function already returned for year < 1998.
    par1 = 0.035;
    par2 = 0.011;
    if (alder < 28) {
      month = 0;
      par1 = 0;
      par2 = 0;
    }
  } else if (year <= 2002) {
    if (alder > 27) {
      par1 = 0.034;
      par2 = 0.01;
    }
  } else if (year < 2005) {
    par1 = 0.035;
    if (alder > 27) par2 = 0.011;
  } else if (year === 2006 || vbaInt(context.born) <= 1946) {
    // Note the `or`: this also catches every year from 2005 for cohorts born
    // 1946 or earlier, which is what makes the next branch unreachable.
    par1 = 0.045;
    par2 = 0.021;
  } else if (year === 2006 && vbaInt(context.born) > 1946) {
    // Unreachable: year === 2006 was already caught above.
    par1 = 0.04;
    par2 = 0.04;
  } else if (year === 2007) {
    par1 = 0.04;
    par2 = 0.04;
  } else if (year <= 2009) {
    par1 = 0.0425;
    par2 = 0.0425;
  } else {
    par1 = 0.045;
    par2 = 0.045;
  }

  if (month === 0) return 0;
  return premiumFromRates(scaled, ibb, par1, par2, marginal);
}

/** AKAP-KR premium for a year. */
export function tlakap_kr(
  alderIn: number,
  inkomstIn: number,
  ibbIn: number,
  zparIn: number,
  year: number,
  context: SchemeContext,
  model: ModelContext,
): number {
  const { born, wStart, marginal } = context;
  const alder = Math.round(alderIn);

  // From 2023 there is no upper age limit for premiums below the breakpoint.
  let highAge: number;
  if (year < 2003) highAge = 65;
  else if (year < 2023) highAge = 67;
  else highAge = 150;

  const zpar = zparIn > highAge ? highAge : zparIn;

  let month: number;
  let inkomst = inkomstIn;

  // Unlike KAP-KL, this branch is reachable: there is no earlier `alder < 21`
  // exit above it. Those born after 1985 accrue from their first working day.
  if (born > 1985 && alder < 21 && year >= 2014) {
    if (alder < wStart) return 0;
    if (wStart === alder) {
      month = vbaInt((vbaInt(born + alder + 1) - born - wStart) * 12);
      if (month > 12) month = 12;
      if (month < 0) month = 0;
      inkomst = (inkomst * month) / 12;
    } else {
      month = 12;
    }
  } else if (alder < 21) {
    return 0;
  } else {
    const worked = monthsFrom21(alder, inkomst, zpar, highAge, context);
    month = worked.month;
    inkomst = worked.inkomst;
  }

  const ibb = ibbIn === 0 ? Math.round(inkomst) : Math.round(ibbIn);
  if (month === 0) return 0;

  let par1 = 0;
  let par2 = 0;
  if (year < 2003) {
    par1 = 0.035;
    par2 = 0.011;
    if (alder < 28) {
      month = 0;
      par1 = 0;
      par2 = 0;
    }
  } else if (year < 2005) {
    par1 = 0.035;
    if (alder > 27) par2 = 0.011;
  } else if (year === 2006 && vbaInt(born) <= 1946) {
    // Note: `and` here, where KAP-KL has `or`. So both 2006 branches are live.
    par1 = 0.045;
    par2 = 0.021;
  } else if (year === 2006 && vbaInt(born) > 1946) {
    par1 = 0.04;
    par2 = 0.04;
  } else if (year <= 2007) {
    par1 = 0.04;
    par2 = 0.04;
  } else if (year <= 2009) {
    par1 = 0.0425;
    par2 = 0.0425;
  } else {
    par1 = 0.045;
    par2 = 0.045;
  }

  // Those born after 1985 moved to the high rate above the breakpoint in 2014.
  if (vbaInt(born) > 1985 && year >= 2014 && year < 2023) {
    par1 = 0.045;
    par2 = 0.3;
  }

  // From 2023 the higher rate is tied to the LAS age, riktålder plus three.
  const lasAge = model.riktage + 3;
  if (year >= 2023 && alder <= lasAge) {
    par1 = 0.06;
    par2 = 0.315;
  } else if (year >= 2023 && alder > lasAge) {
    par1 = 0.06;
    par2 = 0.06;
  }

  if (month === 0) return 0;
  return premiumFromRates(inkomst, ibb, par1, par2, marginal);
}
