/**
 * Turning accumulated occupational pension capital into an annuity.
 *
 * Port of `tjpkassa` and `tjp_ddeltal` from Tjänstepensioner.bas.
 *
 * The divisor comes from the premium pension table -- `deltal(..., "premium")`,
 * the same lookup `ppkassa` uses -- and is then adjusted by `tjp_ddeltal` for
 * the interest rate and life expectancy the agreement's insurer assumes.
 *
 * Under the model's normal settings that adjustment is a pass-through: it only
 * does anything when a temporary withdrawal is set or `rng_Ddelat` is 1.
 */

import type { ModelContext } from "../model/context.js";
import type { DeltalTables } from "../pension/deltal.js";
import { deltal } from "../pension/incomePension.js";
import { vbaInt, vbaRound } from "../vba/math.js";
import { Scheme } from "./types.js";
import type { SchemeId } from "./types.js";

/** Cohorts before 1938 get a flat divisor of 13. */
const PRE_1938_DIVISOR = 13;

/** The premium pension's assumed rate and life expectancy, which the
 *  agreements' own assumptions are measured against. */
function premiumPensionBasis(year: number): { rate: number; life: number } {
  if (year < 2002) return { rate: 4, life: 20.3 };
  if (year < 2018) return { rate: 3, life: 20.3 };
  return { rate: 1.75, life: 22.65 };
}

/** Each agreement's own assumed rate and life expectancy. */
function schemeBasis(scheme: number, year: number): { rate: number; life: number } {
  const after2019 = year > 2019;
  if (scheme === 1) {
    // Private saving, which has no agreement of its own.
    return after2019 ? { rate: 3.5, life: 22.6 } : { rate: 2.5, life: 22 };
  }
  if (scheme < 4) return after2019 ? { rate: 2.2, life: 22 } : { rate: 2.9, life: 22 };
  if (scheme === 4) return after2019 ? { rate: 1.3, life: 20.8 } : { rate: 2.25, life: 21.3 };
  if (scheme < 7) return after2019 ? { rate: 2.75, life: 23.1 } : { rate: 2.75, life: 22.8 };
  return after2019 ? { rate: 2, life: 23.3 } : { rate: 2, life: 21.9 };
}

/**
 * Adjusts a premium pension divisor for one agreement's own assumptions.
 *
 * Returns the divisor unchanged unless a temporary withdrawal is set or
 * `rng_Ddelat` is 1 -- which is to say, unchanged under the model's normal
 * settings.
 *
 * @param tal    the divisor from the premium pension table
 * @param scheme the agreement, numbered as the Start sheet numbers it
 * @param ips    non-zero when this is private saving rather than occupational
 */
export function tjp_ddeltal(
  tal: number,
  scheme: number,
  year: number,
  tjpPar: number,
  context: ModelContext,
  ips = 0,
): number {
  let utbtid = ips === 0 ? context.tempTjpUttag : context.tempIpsUttag;
  if (utbtid <= 0) utbtid = 100;

  // Lifelong withdrawal with no divisor correction asked for: the occupational
  // pension is then valued exactly as the premium pension is. This is the
  // default path.
  if (utbtid > 98 && !context.adjustOccupationalDivisor) return tal;

  const basis = premiumPensionBasis(year);
  // Elasticities from Pensionsmyndigheten's deltal_sensitivity memorandum.
  const [b1, b2] = year > 2019 ? [-0.16216, 0.84517] : [-0.27166, 0.78457];

  const effectiveScheme = ips > 0 ? 1 : scheme;
  const own = schemeBasis(effectiveScheme, year);

  if (effectiveScheme === 0) {
    // Premium pension itself, corrected only for the rate difference.
    let rate = context.rgk;
    if (rate < 0.1) rate = 100 * rate;
    return tal * (1 - 0.1812 * (rate / basis.rate - 1));
  }

  // A temporary withdrawal is valued directly, by discounting survival over the
  // payout years rather than scaling a lifelong divisor.
  if (utbtid > 0 && utbtid < 30) {
    const makeham = (x: number) => 0.0002 + 0.000007 * Math.exp(0.1071 * x);
    const surv = 1 - makeham(tjpPar);
    let s = 0;
    for (let i = tjpPar; i <= tjpPar + vbaInt(utbtid) - 1; i += 1) {
      s += (1 - makeham(i)) / (1 + own.rate / 100) ** (i - tjpPar);
    }
    return s / surv;
  }

  let adjusted = tal * (1 + b1 * (own.rate / basis.rate - 1) + b2 * (own.life / basis.life - 1));
  if (context.marginal === 0) adjusted = vbaRound(adjusted, 2);
  return adjusted;
}

/**
 * The occupational pension paid out in a year.
 *
 * @param pbh accumulated capital, including this year's premium and return
 */
export function tjpkassa(
  tjpPar: number,
  born: number,
  alder: number,
  pbh: number,
  scheme: SchemeId,
  year: number,
  tables: DeltalTables,
  context: ModelContext,
  tmonth: number,
  ips = 0,
): number {
  if (pbh < 1) return 0;

  const konst = vbaInt(born + tjpPar + 1 / 1000) > vbaInt(born) + vbaInt(tjpPar) ? 1 : 0;

  let divisor: number;
  if (born < 1938) {
    divisor = PRE_1938_DIVISOR;
  } else if (alder < 98) {
    divisor = deltal(tjpPar, born, alder, tables, "premium", 999);
    // Where the premium pension table has nothing, step a year on and add a
    // rough correction -- the VBA's own words are "godtyckligt", arbitrary.
    if (divisor <= 0) divisor = deltal(tjpPar + 1, born, alder, tables, "premium", 999) + 0.6;
  } else {
    divisor = 2;
  }

  divisor = tjp_ddeltal(divisor, scheme, year, tjpPar, context, ips);
  if (divisor <= 0) return 0;

  let amount: number;
  let month = tmonth;
  if (alder < vbaInt(tjpPar + konst)) {
    return 0;
  } else if (alder === vbaInt(tjpPar + konst)) {
    amount = pbh / divisor;
  } else {
    amount = pbh / divisor;
    month = 12;
  }

  // Rounded to a whole krona per month, then multiplied back up.
  amount = context.marginal === 0 ? vbaInt(amount / 12 + 0.5) : amount / 12;
  return amount * month;
}

export { Scheme };
