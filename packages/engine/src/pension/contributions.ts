/**
 * Pension-qualifying income and the contributions paid on it.
 *
 * Port of the first part of Pensionssystemet.bas: `riktage`, `pgi`, `andel`,
 * `ipavgift`, `ppavgift`, `gpavgift` and `PPMavg`.
 *
 * Verified against reference/fixtures/brutto-cached.json -- the Brutto sheet
 * calls these same functions as worksheet UDFs and keeps the results, so the
 * port is checked against the original across ~80 years of one typfall.
 */

import type { ModelContext } from "../model/context.js";
import { vbaCLng, vbaInt } from "../vba/math.js";

/** What `pgi` should return; the VBA passes this as its `Typ` argument. */
export const PgiResult = {
  /** Pension-qualifying income itself. */
  Income: 0,
  /** PAVG -- the general pension contribution the individual pays. */
  Contribution: 1,
  /** SRED -- the tax reduction that contribution earns. */
  TaxReduction: 2,
  /** The ceiling in force that year. */
  Ceiling: 4,
  /** The assessed income, after rounding. */
  AssessedIncome: 9,
} as const;

export type PgiResultKind = (typeof PgiResult)[keyof typeof PgiResult];

/**
 * Riktålder, and the earliest age public pension may be drawn.
 *
 * @param type 0 for the earliest age, anything else for riktålder
 */
export function riktage(year: number, type: 0 | 1): number {
  let rikt: number;
  let riktl: number;
  if (year < 2020) {
    rikt = 65;
    riktl = 61;
  } else if (year < 2023) {
    rikt = 65;
    riktl = 62;
  } else if (year < 2026) {
    rikt = 66;
    riktl = 63;
  } else if (year < 2038) {
    rikt = 67;
    riktl = 64;
  } else if (year < 2051) {
    rikt = 68;
    riktl = 65;
  } else if (year < 2068) {
    rikt = 69;
    riktl = 66;
  } else {
    rikt = 70;
    riktl = 67;
  }
  return type === 0 ? riktl : rikt;
}

/**
 * The share of a cohort's public pension that comes from the new system, the
 * remainder being ATP.
 *
 * Phased in a twentieth at a time: cohorts born 1938 through 1953 get part of
 * each, 1954 onwards is entirely the new system, 1937 and earlier entirely ATP.
 */
export function andel(kohort: number): number {
  const cohort = vbaInt(kohort);
  let share = 0;
  if (cohort >= 1938) share = (1 / 20) * (cohort - 1935 + 1);
  if (cohort > 1953) share = 1;
  return share;
}

/**
 * Pension-qualifying income (PGI), and the contribution and tax reduction on it.
 *
 * The ceiling is 8.07 income base amounts of gross income, from which the 7%
 * general pension contribution is deducted -- landing at the statutory 7.5 base
 * amounts, which the final clamp enforces exactly.
 */
export function pgi(
  year: number,
  inkomst: number,
  pbb: number,
  ibb: number,
  fhb: number,
  context: ModelContext,
  typ: PgiResultKind = PgiResult.Income,
  alder = 64,
  sjuk = 0,
): number {
  const { marginal, socTak, socialAvg } = context;
  let result = 0;

  // The state pays the whole contribution on sickness and activity compensation.
  let income = inkomst - sjuk;
  if (marginal === 0) income = vbaInt(income / 100) * 100;

  let ceiling = 8.07 * ibb;
  if (year <= 1994) ceiling = 7.5 * pbb;
  else if (year < 1999) ceiling = 7.5 * fhb;
  else if (year === 1999) ceiling = 8.06 * fhb;
  else if (year === 2000) ceiling = 8.07 * fhb;
  // Policy experiment: a raised ceiling from a given year.
  if (socTak > 1999 && year >= socTak) ceiling = 7.5 * ibb;

  let floor = 0.423 * pbb;
  if (year <= 1994) floor = 1 * pbb;
  else if (year < 1999) floor = 1 * fhb;
  else if (year < 2001) floor = 0.24 * pbb;
  else if (year === 2001) floor = 0.27 * pbb;
  else if (year === 2002) floor = 0.293 * pbb;

  if (income <= floor) result = 0;
  else if (income <= ceiling) result = income;
  else result = ceiling;

  let egen = 0.07;
  if (year < 1993) egen = 0;
  else if (year === 1993) egen = 0;
  else if (year === 1994) egen = 0;
  else if (year === 1995) egen = 0.01;
  else if (year === 1996) egen = 0.01;
  else if (year === 1997) egen = 0.01;
  else if (year < 2000) egen = 0.0695;

  let pavg = result * egen;
  if (marginal === 0) pavg = vbaInt((pavg + 49) / 100) * 100;
  if (socialAvg > 1999 && year >= socialAvg) pavg = 0;

  // QUIRK: the VBA declares this share `As Long`, so the fractional phase-in
  // values are rounded to whole numbers on assignment -- 0.25 and 0.5 become 0
  // (0.5 by banker's rounding), 0.75 and 0.875 become 1. The tax reduction is
  // therefore all-or-nothing, phasing in at 2002 rather than gradually from
  // 2000. Faithful to the original; see docs/VBA-MAPPING.md.
  let reductionShare: number;
  if (year < 2000) reductionShare = 0;
  else if (year <= 2000) reductionShare = vbaCLng(0.25);
  else if (year <= 2001) reductionShare = vbaCLng(0.5);
  else if (year <= 2002) reductionShare = vbaCLng(0.75);
  else if (year <= 2005) reductionShare = vbaCLng(0.875);
  else reductionShare = 1;

  let sred = reductionShare * pavg;
  if (marginal === 0) sred = vbaInt(sred / 100) * 100;

  income += sjuk;
  if (marginal === 0) income = vbaInt(income / 100) * 100;

  if (socialAvg > 1999 && year >= socialAvg) {
    // A raised contribution is borne by the employer, so nothing is deducted.
    if (income < floor) result = 0;
    else if (income <= ceiling) result = income;
    else result = ceiling;
  } else if (income < floor) result = 0;
  else if (income <= ceiling) result = income - pavg;
  else result = ceiling - pavg;

  if (marginal === 0) {
    result = year < 1999 ? vbaInt(result / 50) * 50 : vbaInt(result / 100) * 100;
  }

  // The statutory ceiling: 7.5 income base amounts after the contribution.
  if (result > 7.5 * ibb) result = 7.5 * ibb;

  if (typ === PgiResult.Contribution) return pavg;
  if (typ === PgiResult.TaxReduction) return sred;
  if (typ === PgiResult.Ceiling) return ceiling;
  if (typ === PgiResult.AssessedIncome) return income;
  return result;
}

/**
 * The income pension contribution -- 16% of PGI today.
 *
 * Paid only for cohorts born 1938 or later; PGI is zero for the rest anyway.
 */
export function ipavgift(
  year: number,
  pgiAmount: number,
  alder: number,
  share = 1,
  marginal: 0 | 1 = 0,
  fodd = 1938,
): number {
  if (fodd < 1938) return 0;

  let rate = 0.16;
  if (year <= 1994 && marginal === 0) rate = 0.185;
  if (year > 1994 && year <= 1998 && marginal === 0) rate = 0.165;

  // From 65 the whole contribution goes to the new system.
  const effectiveShare = alder > 64 ? 1 : share;
  const amount = rate * pgiAmount * effectiveShare;
  return marginal === 0 ? vbaInt(amount) : amount;
}

/** The premium pension contribution -- 2.5% of PGI today. */
export function ppavgift(
  year: number,
  pgiAmount: number,
  alder: number,
  share = 1,
  marginal: 0 | 1 = 0,
  fodd = 1938,
): number {
  if (fodd < 1938) return 0;

  let rate = 0.025;
  if (year <= 1994) rate = 0;
  if (year > 1994 && year <= 1998) rate = 0.02;

  const effectiveShare = alder > 64 ? 1 : share;
  const amount = rate * pgiAmount * effectiveShare;
  return marginal === 0 ? vbaInt(amount) : amount;
}

/**
 * The notional 18.5% used to work out entitlement to garantipension -- as if
 * the whole contribution had gone to the income pension.
 */
export function gpavgift(
  year: number,
  pgiAmount: number,
  alder: number,
  share = 1,
  marginal: 0 | 1 = 0,
  fodd = 1938,
  rikt = 65,
): number {
  if (fodd < 1938) return 0;

  const effectiveShare = alder >= rikt ? 1 : share;
  const amount = 0.185 * pgiAmount * effectiveShare;
  return marginal === 0 ? vbaInt(amount) : amount;
}

/**
 * The premium pension's management charge, returned as a negative amount.
 *
 * Zero when returns are already quoted net of fund fees, which is the default.
 */
export function ppmavg(
  year: number,
  kapital: number,
  faktor: number,
  avkast: number,
  context: ModelContext,
): number {
  if (context.returnsNetOfFees) return 0;

  // Fund companies' own charge; always zero in the model as it stands.
  const fakt2 = 0;

  let charge =
    year < 2011
      ? kapital * (1 - faktor + fakt2)
      : kapital * (1 - faktor + fakt2) * (1 + avkast) ** 0.5;

  // The cap, raised repeatedly over the years.
  if (year < 2007) {
    // uncapped
  } else if (year === 2007 && charge > 100) charge = 100;
  else if (year < 2010 && charge > 110) charge = 110;
  else if (year === 2010 && charge > 125) charge = 125;
  else if (year < 2014 && charge > 110) charge = 110;
  else if (year < 2017 && charge > 120) charge = 120;
  else if (year < 2018 && charge > 125) charge = 125;
  else if (year < 2019 && charge > 125) charge = 160;
  else charge = 100;

  const fundCharge = kapital * fakt2 * (1 + avkast) ** 0.5;
  return -charge - fundCharge;
}
