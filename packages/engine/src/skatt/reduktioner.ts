/**
 * State income tax, the public service fee, and the tax reductions.
 *
 * Port of the remaining functions in Skatteregler.bas: `statlig`, `PublicAvg`,
 * `sared`, `FAared`, `pandred`, `arbgiv`, `avdragRES` and `avkskatt`.
 */

import type { ModelContext } from "../model/context.js";
import type { RunVectors } from "../model/runVectors.js";
import { vbaInt } from "../vba/math.js";
import { avdragxx } from "./grundavdrag.js";

/**
 * State income tax: 20% above the first threshold, 25% above the second.
 *
 * The second threshold (värnskatten) was abolished in 2020; the data carries
 * 1e16 from then on, so the upper band never binds.
 */
export function statlig(
  besk: number,
  lim1 = 383_000,
  lim2 = 548_300,
  marginal: 0 | 1 = 0,
): number {
  let tax: number;
  if (besk > lim2) {
    tax = (besk - lim2) * 0.25 + (lim2 - lim1) * 0.2;
  } else if (besk > lim1) {
    tax = (besk - lim1) * 0.2;
  } else {
    tax = 0;
  }
  return marginal === 0 ? vbaInt(tax) : tax;
}

/**
 * The public service fee, which replaced the licence fee in 2019.
 *
 * 1% of taxable income up to a ceiling that has been lowered each year.
 */
export function PublicAvg(
  besk: number,
  born: number,
  vectors: RunVectors,
  lim1 = 0.01,
  bald = 18,
  marginal: 0 | 1 = 0,
  year = 2019,
): number {
  if (bald <= 18 || besk < 1 || year < 2019) return 0;

  const ageThatYear = year - vbaInt(born);
  const ibb = vectors.ibb(ageThatYear);

  let ceiling = 2.092 * ibb;
  if (year === 2021) ceiling = 1.95 * ibb;
  if (year === 2022) ceiling = 1.87 * ibb;
  if (year === 2023) ceiling = 1.75 * ibb;
  if (year === 2024) ceiling = 1.6 * ibb;
  if (year === 2025) ceiling = 1.55 * ibb;
  if (year >= 2026) ceiling = 1.42 * ibb;

  const fee = besk > ceiling ? ceiling * lim1 : besk * lim1;
  return marginal === 0 ? vbaInt(fee + 0.5) : fee;
}

/**
 * Tax reduction for recipients of sickness and activity compensation.
 *
 * NOTE: the VBA reads `alder` and `binkomst` here without them being parameters
 * or ever assigned anywhere in the module -- and `Skatteregler.bas` declares no
 * `Option Explicit`, so both are implicit module-level Variants that are always
 * Empty, i.e. 0. They are passed as 0 here for that reason, not by choice.
 */
export function sared(
  inkomst: number,
  year: number,
  marginal: 0 | 1 = 0,
  ksats = 0.3212,
  pbb = 45_500,
  wyear = 21_000,
  IBB = 0,
  kvoten = 1,
  Iyear = 0,
): number {
  if (year < 2018) return 0;

  const alder = 0;
  const binkomst = 0;
  let reduction = 0;

  if (year < 2022) {
    reduction =
      inkomst < 2.53 * pbb
        ? 0.045 * inkomst
        : 0.045 * 2.53 * pbb + 0.025 * (inkomst - 2.53 * pbb);
    reduction = reduction * ksats;
  }

  // Note both of the blocks below run for a year under 2022 as well, overwriting
  // what the first computed. Faithful to the original's `If` rather than `ElseIf`.
  if (year < 2026) {
    if (inkomst <= 0.91 * pbb) {
      reduction = inkomst;
    } else if (inkomst <= 3.24 * pbb) {
      reduction = 0.3405 * (inkomst - 0.91 * pbb) + 0.91 * pbb;
    } else {
      reduction = 0.128 * (inkomst - 1.703 * pbb);
    }
    const avdrag = avdragxx(inkomst + binkomst, pbb, marginal, alder, year, wyear, IBB, kvoten, Iyear);
    reduction = (avdrag - reduction) * ksats;
    if (inkomst * 0.045 * ksats > reduction) reduction = inkomst * 0.045 * ksats;
  }

  if (year > 2025) {
    if (inkomst <= 0.91 * pbb) {
      reduction = inkomst;
    } else if (inkomst <= 3.24 * pbb) {
      reduction = 0.3874 * (inkomst - 0.91 * pbb) + 0.91 * pbb;
    } else {
      reduction = 0.251 * (inkomst - 1.813 * pbb);
    }
    const avdrag = avdragxx(inkomst + binkomst, pbb, marginal, alder, year, wyear, IBB, kvoten, Iyear);
    reduction = (avdrag - reduction) * ksats;
    if (inkomst * 0.045 * ksats > reduction) reduction = inkomst * 0.045 * ksats;
  }

  return marginal === 0 ? vbaInt(reduction) : reduction;
}

/**
 * Tax reduction for earned income, introduced in the 2021 budget.
 *
 * 0.75% of income above 40 000 kr, capped at 1 500 kr.
 */
export function FAared(
  inkomst: number,
  year: number,
  marginal: 0 | 1,
  pbb = 47_600,
): number {
  if (year < 2021) return 0;

  const limit1 = 40_000;
  const limit2 = 240_000;
  const andel = 0.0075;

  let reduction = 0;
  if (inkomst > limit1 && inkomst < limit2) {
    reduction = andel * (inkomst - limit1);
  } else if (inkomst >= limit2) {
    reduction = 1500;
  }

  if (reduction > inkomst) reduction = inkomst;
  return marginal === 0 ? vbaInt(reduction) : reduction;
}

/** The temporary tax reduction for pandemic costs, 2021 to 2023. */
export function pandred(wage: number, year: number, marginal: 0 | 1): number {
  if (year < 2021 || year > 2023) return 0;
  if (wage < 60_000 || wage > 500_000) return 0;

  let reduction = 2250;
  if (wage < 240_000) reduction = 0.0125 * (wage - 60_000);
  if (wage > 300_000) reduction = 2250 - 0.01125 * (wage - 300_000);

  return marginal === 0 ? vbaInt(reduction) : reduction;
}

/** What `arbgiv` should return. */
export const ArbgivResult = {
  /** The contribution in kronor. */
  Amount: 0,
  /** The rate itself. */
  Rate: 1,
} as const;

/**
 * Employer social contributions.
 *
 * Below riktålder the full rate applies; at or above it, only the old-age
 * pension component.
 *
 * NOTE ON SCALE: the K_skatt sheet holds these as percentages (31.42, 10.21)
 * and the VBA multiplies the wage by them directly, so the amount comes out a
 * hundred times too large. The model only ever writes it to the output sheet
 * "för studier av arbetsgivarna" and never feeds it into a pension or a tax, so
 * nothing downstream depends on it. Reproduced as written.
 *
 * @param rates the percentages from K_skatt: full rate and old-age component
 */
export function arbgiv(
  wage: number,
  year: number,
  age: number,
  rates: { full: number; oldAge: number },
  riktage: number,
  marginal: 0 | 1 = 0,
  typ: number = ArbgivResult.Amount,
): number {
  if (wage < 1000) return 0;
  if (year < 1960) return 0;

  const sats = age < riktage && typ === ArbgivResult.Amount ? rates.full : rates.oldAge;

  if (typ === ArbgivResult.Rate) return sats;
  const amount = wage * sats;
  return marginal === 0 ? vbaInt(amount + 0.5) : amount;
}

/** Which deduction `avdragRES` is capping. */
export const DeductionType = {
  Travel: 1,
  PrivatePension: 2,
  Other: 3,
} as const;

/**
 * Caps a deduction at the threshold in force, and nets off the floor below
 * which it is not allowed at all.
 */
export function avdragRES(
  avdragIn: number,
  age: number,
  vectors: RunVectors,
  hasOccupationalPension: boolean,
  typ: number = DeductionType.Travel,
  marginal: 0 | 1 = 0,
  year = 2011,
): number {
  if (avdragIn === 0) return 0;
  if (year < 1994) return 0;

  const tjp = hasOccupationalPension ? 1 : 0;

  // Those without an occupational pension may deduct far more private saving.
  let max = 0.35 * vectors.wage(age);
  if (max > 10 * vectors.pbb(age)) max = 10 * vectors.pbb(age);

  let maxips = 12_000;
  if (year === 2015) maxips = 1800;
  if (year >= 2016) maxips = 0;
  if (typ === DeductionType.PrivatePension && year > 2015 && tjp > 0) return 0;
  if (tjp === 0) maxips += max;

  let travelFloor: number;
  if (year < 2007) travelFloor = 7000;
  else if (year <= 2010) travelFloor = 8000;
  else if (year === 2011) travelFloor = 9000;
  else if (year <= 2016) travelFloor = 10_000;
  else travelFloor = 11_000;

  const limits = [0, travelFloor, maxips, 1000];

  let avdrag = avdragIn;
  if (typ === DeductionType.PrivatePension) {
    if (avdrag > limits[2]!) avdrag = limits[2]!;
  } else if (typ < 4) {
    avdrag -= limits[typ] ?? 0;
  }
  if (typ === 4) avdrag = 0;
  if (avdrag < 0) avdrag = 0;

  return marginal === 0 ? vbaInt(avdrag) : avdrag;
}

/** What `avkskatt` is being asked about. */
export const YieldTaxKind = {
  /** Pension insurance, taxed at 15%. */
  PensionInsurance: 0,
  /** ISK or kapitalförsäkring, taxed at 30%. */
  CapitalInsurance: 1,
} as const;

/**
 * The yield tax rate for a year, from the government borrowing rate.
 *
 * A floor of 0.5% applies to pension insurance from 2017, and capital insurance
 * adds a percentage point with a floor of 0.25% on the rate itself.
 */
export function avkskatt(year: number, context: ModelContext, val: number = YieldTaxKind.CapitalInsurance): number {
  const previousYear = year - 1;

  const rates: Readonly<Record<number, number>> = {
    1998: 0.05, 1999: 0.0489, 2000: 0.0534, 2001: 0.0497, 2002: 0.0515,
    2003: 0.0439, 2004: 0.043, 2005: 0.0325, 2006: 0.0361, 2007: 0.0413,
    2008: 0.0388, 2009: 0.0309, 2010: 0.0277, 2011: 0.0258, 2012: 0.0152,
    2013: 0.02, 2014: 0.0163, 2015: 0.0058, 2016: 0.0034, 2017: 0.0051,
    2018: 0.0048, 2019: 0.0004, 2020: -0.0007, 2021: 0.0016, 2022: 0.0194,
    2023: 0.0262, 2024: 0.0196, 2025: 0.0255,
  };

  let bondi: number;
  if (previousYear < 1998) bondi = 0.0498;
  else bondi = rates[previousYear] ?? 0.025;
  void context;

  if (year > 2016 && val === YieldTaxKind.PensionInsurance) {
    if (bondi < 0.005) bondi = 0.005;
  }

  if (val === YieldTaxKind.PensionInsurance) return bondi * 0.15;
  if (bondi < 0.0025) bondi = 0.0025;
  return (0.01 + bondi) * 0.3;
}
