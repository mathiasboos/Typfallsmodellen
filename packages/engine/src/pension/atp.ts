/**
 * Tilläggspension -- the ATP system that preceded the current one.
 *
 * Port of `tp_`, `tp_faktor`, `pts` and `fnorm` from Pensionssystemet.bas.
 *
 * Only cohorts born 1953 or earlier receive any of it: 1937 and earlier are
 * entirely on ATP, and 1938-1953 get a share that falls a twentieth per year
 * (see `andel` in contributions.ts).
 */

import { vbaInt } from "../vba/math.js";

/** The last cohort with any entitlement to tilläggspension. */
export const LAST_ATP_COHORT = 1953;
/** Years of pensionable income needed for a full ATP. */
const YEARS_FOR_FULL_ATP = 30;
/** Below this many years there is no ATP at all. */
const MINIMUM_ATP_YEARS = 3;

/**
 * The follow-up indexation norm, 1.6% -- except in 2000, its first year, when a
 * lower figure was used, and before 2000 when there was none.
 */
export function fnorm(year: number): number {
  if (vbaInt(year) === 2000) return 0.996;
  if (vbaInt(year) < 2000) return 1;
  return 1.016;
}

/**
 * The adjustment for drawing tilläggspension before or after 65.
 *
 * Reduced 0.5% for each month early, increased 0.7% for each month late, and
 * frozen past 70.
 */
export function tpFaktor(par: number): number {
  const month2 = par > 70 ? 0 : vbaInt(12 * (par - vbaInt(par)));

  if (vbaInt(par) < 65) {
    return 1 - 0.005 * (12 * (65 - vbaInt(par)) - month2);
  }
  if (vbaInt(par) === 65) {
    return 1 + 0.007 * month2;
  }
  if (vbaInt(par) <= 70) {
    return 1 + 0.007 * (12 * (vbaInt(par) - 65) + month2);
  }
  return 1 + 0.007 * 12 * (70 - 65);
}

/**
 * Annual tilläggspension at first withdrawal.
 *
 * @param poang  average ATP points earned
 * @param nyear  years with ATP points
 * @param civ    0 unmarried, 1 married -- the folkpension part differs
 * @param par    age at which pension is drawn
 * @param fodar  birth year
 * @param alder  age at 31 December
 * @param pbb    the price base amount for the year. In the VBA this is read as
 *               `pbb(age)` off the module-level array using the *global* loop
 *               age rather than the `alder` argument; every call site passes the
 *               same value for both, so it is taken as a parameter here.
 * @param andel  the share of pension being drawn
 */
export function tp(
  poang: number,
  nyear: number,
  civ: number,
  par: number,
  fodar: number,
  alder: number,
  pbb: number,
  andel = 1,
): number {
  if (alder < vbaInt(par)) return 0;
  if (fodar > LAST_ATP_COHORT) return 0;

  // The ATP part plus the old folkpension part, which was lower for a married
  // person because the household received two.
  let atp = 0.6 * poang * pbb + (civ === 0 ? 0.96 : 0.785) * pbb;

  if (fodar > 1937) {
    // Short of thirty qualifying years, the pension is reduced in proportion.
    if (nyear < YEARS_FOR_FULL_ATP) atp = (atp * nyear) / YEARS_FOR_FULL_ATP;
    if (nyear < MINIMUM_ATP_YEARS) atp = 0;
  }

  const faktor = tpFaktor(par);

  // QUIRK: the factor is applied twice -- once here and again in the return.
  // It reads as a leftover from the 2021 rewrite that reduced this function to
  // the first withdrawal only (most of the original body is commented out
  // around these two lines). It has no effect at exactly 65, where the factor
  // is 1, but squares the adjustment at any other retirement age.
  //
  // Kept, because the aim is parity with the model. It affects only cohorts
  // born 1953 or earlier, and the golden files will confirm it.
  atp = atp * faktor;
  return atp * andel * faktor;
}

/**
 * Pensionstillskottet -- the supplement paid under the old system to those with
 * little or no ATP, which the ATP amount was set against.
 */
export function pts(
  alder: number,
  pbb: number,
  atp = 0,
  marginal: 0 | 1 = 0,
  nyear = 30,
): number {
  const koeff = 0.569;
  let amount = koeff * pbb - atp;

  // Drawn early, the supplement is reduced too.
  if (alder > 60 && alder < 65) amount -= 0.005 * (65 - alder) * pbb;

  // To the nearest whole krona per month.
  if (marginal === 0) amount = vbaInt(amount / 12 + 0.5) * 12;
  if (amount < 0) amount = 0;
  if (nyear < MINIMUM_ATP_YEARS) amount = 0;
  return amount;
}
