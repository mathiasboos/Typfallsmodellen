/**
 * Grundavdrag -- the basic allowance, one function per set of rules.
 *
 * These are the only functions in the engine that were translated
 * mechanically rather than written out. `Skatteregler.bas` holds twenty
 * near-identical grundavdrag functions and twelve jobbskatteavdrag ones -- some
 * 1 300 lines of nothing but thresholds and coefficients. Typing those by hand
 * invites exactly the transposed digit no test would catch, so
 * `tools/transpile/skatteregler.py` translates them and the result is reviewed
 * and committed. Re-run it for a new rule year and paste the new function in.
 *
 * Each takes the same arguments as its VBA original, in the same order:
 *
 *   inkomst   assessed earned income
 *   pbb       price base amount
 *   marginal  0 applies the rules' rounding, 1 removes it
 *   bald      age at 31 December, which decides the higher allowance for the old
 *   year      income year
 *   wyear     the year from which thresholds follow the income base amount
 *   IBB       income base amount, for that switch
 *   kvoten    the price/income base amount ratio at the switch
 *   Iyear     the year whose rules apply
 *   Xage      the age at which the higher allowance begins
 */

import { vbaCLng, vbaInt } from "../vba/math.js";
import type { ModelContext } from "../model/context.js";

/**
 * The age at which the higher grundavdrag for older people begins.
 *
 * It has been rising with riktålder, and past 2028 is projected as riktålder
 * plus one.
 */
export function Xage(year: number, context: ModelContext): number {
  if (year <= 2023) return 66;
  if (year <= 2026) return 67;
  if (year <= 2028) return 68;
  // `Rng_riktage + 1` -- the *cohort's* riktålder, which `prepareRun` puts on
  // the context for the run. See pension/retirementAges.ts.
  return context.riktage + 1;
}

export function avdrag95(inkomst: number, pbb = 41000, marginal = 0, bald = 64, year = 2100, wyear = 21000, IBB = 0, kvoten = 1, Iyear = 0, Xage = 66): number {
  let result = 0;
  if (year > wyear && kvoten < 1) {
    pbb = IBB;
    pbb = pbb * kvoten;
  }
  let avdrag = 0;
  let Extra = 0;
  if (inkomst < 0.25 * pbb) {
    avdrag = inkomst;
  } else if (inkomst <= 1.86 * pbb) {
    avdrag = 0.25 * pbb;
  } else if (inkomst <= 2.89 * pbb) {
    avdrag = 0.25 * pbb + 0.25 * (inkomst - 1.86 * pbb);
  } else if (inkomst <= 3.04 * pbb) {
    avdrag = 0.25 * pbb + 0.25 * (2.89 - 1.86) * pbb;
  } else {
    avdrag = 0.25 * pbb + 0.25 * (2.89 - 1.86) * pbb - 0.1 * (inkomst - 3.04 * pbb);
    if (avdrag < 0.25 * pbb) avdrag = 0.25 * pbb;
  }
  if (avdrag > inkomst) avdrag = inkomst;
  if (marginal === 0) avdrag = vbaInt((avdrag + 99.99) / 100) * 100;
  result = avdrag;
  return result;
}

export function avdrag96(inkomst: number, pbb = 41000, marginal = 0, bald = 64, year = 2100, wyear = 21000, IBB = 0, kvoten = 1, Iyear = 0, Xage = 66): number {
  let result = 0;
  if (year > wyear && kvoten < 1) {
    pbb = IBB;
    pbb = pbb * kvoten;
  }
  let avdrag = 0;
  let Extra = 0;
  if (inkomst < 0.24 * pbb) {
    avdrag = inkomst;
  } else if (inkomst <= 1.86 * pbb) {
    avdrag = 0.24 * pbb;
  } else if (inkomst <= 2.89 * pbb) {
    avdrag = 0.24 * pbb + 0.25 * (inkomst - 1.86 * pbb);
  } else if (inkomst <= 3.04 * pbb) {
    avdrag = 0.24 * pbb + 0.25 * (2.89 - 1.86) * pbb;
  } else {
    avdrag = 0.24 * pbb + 0.25 * (2.89 - 1.86) * pbb - 0.1 * (inkomst - 3.04 * pbb);
    if (avdrag < 0.24 * pbb) avdrag = 0.24 * pbb;
  }
  if (avdrag > inkomst) avdrag = inkomst;
  if (marginal === 0) avdrag = vbaInt((avdrag + 99.99) / 100) * 100;
  result = avdrag;
  return result;
}

export function avdrag01(inkomst: number, pbb = 41000, marginal = 0, bald = 64, year = 2100, wyear = 21000, IBB = 0, kvoten = 1, Iyear = 0, Xage = 66): number {
  let result = 0;
  if (year > wyear && kvoten < 1) {
    pbb = IBB;
    pbb = pbb * kvoten;
  }
  let avdrag = 0;
  let Extra = 0;
  if (inkomst < 0.27 * pbb) {
    avdrag = inkomst;
  } else if (inkomst <= 1.86 * pbb) {
    avdrag = 0.27 * pbb;
  } else if (inkomst <= 2.89 * pbb) {
    avdrag = 0.27 * pbb + 0.25 * (inkomst - 1.86 * pbb);
  } else if (inkomst <= 3.04 * pbb) {
    avdrag = 0.27 * pbb + 0.25 * (2.89 - 1.86) * pbb;
  } else {
    avdrag = 0.27 * pbb + 0.25 * (2.89 - 1.86) * pbb - 0.1 * (inkomst - 3.04 * pbb);
    if (avdrag < 0.27 * pbb) avdrag = 0.27 * pbb;
  }
  if (avdrag > inkomst) avdrag = inkomst;
  if (marginal === 0) avdrag = vbaInt((avdrag + 99.99) / 100) * 100;
  result = avdrag;
  return result;
}

export function avdrag03(inkomst: number, pbb = 41000, marginal = 0, bald = 64, year = 2100, wyear = 21000, IBB = 0, kvoten = 1, Iyear = 0, Xage = 66): number {
  let result = 0;
  if (year > wyear && kvoten < 1) {
    pbb = IBB;
    pbb = pbb * kvoten;
  }
  let avdrag = 0;
  let Extra = 0;
  if (inkomst < 0.423 * pbb) {
    avdrag = inkomst;
  } else if (inkomst <= 1.49 * pbb) {
    avdrag = 0.423 * pbb;
  } else if (inkomst <= 2.72 * pbb) {
    avdrag = 0.423 * pbb + 0.2 * (inkomst - 1.49 * pbb);
  } else if (inkomst <= 3.11 * pbb) {
    avdrag = 0.67 * pbb;
  } else if (inkomst <= 6.87 * pbb) {
    avdrag = 0.67 * pbb - 0.1 * (inkomst - 3.11 * pbb);
  } else {
    avdrag = 0.293 * pbb;
  }
  if (avdrag > inkomst) avdrag = inkomst;
  if (marginal === 0) avdrag = vbaInt((avdrag + 99.99) / 100) * 100;
  result = avdrag;
  return result;
}

export function avdrag05(inkomst: number, pbb = 41000, marginal = 0, bald = 64, year = 2100, wyear = 21000, IBB = 0, kvoten = 1, Iyear = 0, Xage = 66): number {
  let result = 0;
  if (year > wyear && kvoten < 1) {
    pbb = IBB;
    pbb = pbb * kvoten;
  }
  let avdrag = 0;
  let Extra = 0;
  if (inkomst < 0.423 * pbb) {
    avdrag = inkomst;
  } else if (inkomst <= 1.185 * pbb) {
    avdrag = 0.423 * pbb;
  } else if (inkomst <= 2.72 * pbb) {
    avdrag = 0.423 * pbb + 0.2 * (inkomst - 1.185 * pbb);
  } else if (inkomst <= 3.11 * pbb) {
    avdrag = 0.73 * pbb;
  } else if (inkomst <= 7.48 * pbb) {
    avdrag = 0.73 * pbb - 0.1 * (inkomst - 3.11 * pbb);
  } else {
    avdrag = 0.293 * pbb;
  }
  if (avdrag > inkomst) avdrag = inkomst;
  if (marginal === 0) avdrag = vbaInt((avdrag + 99.99) / 100) * 100;
  result = avdrag;
  return result;
}

export function avdrag06(inkomst: number, pbb = 41000, marginal = 0, bald = 64, year = 2100, wyear = 21000, IBB = 0, kvoten = 1, Iyear = 0, Xage = 66): number {
  let result = 0;
  if (year > wyear && kvoten < 1) {
    pbb = IBB;
    pbb = pbb * kvoten;
  }
  let avdrag = 0;
  let Extra = 0;
  if (inkomst < 0.423 * pbb) {
    avdrag = inkomst;
  } else if (inkomst <= 0.99 * pbb) {
    avdrag = 0.423 * pbb;
  } else if (inkomst <= 2.72 * pbb) {
    avdrag = 0.423 * pbb + 0.2 * (inkomst - 0.99 * pbb);
  } else if (inkomst <= 3.11 * pbb) {
    avdrag = 0.77 * pbb;
  } else if (inkomst <= 7.88 * pbb) {
    avdrag = 0.77 * pbb - 0.1 * (inkomst - 3.11 * pbb);
  } else {
    avdrag = 0.293 * pbb;
  }
  if (avdrag > inkomst) avdrag = inkomst;
  if (marginal === 0) avdrag = vbaInt((avdrag + 99.99) / 100) * 100;
  result = avdrag;
  return result;
}

export function avdrag09(inkomst: number, pbb = 41000, marginal = 0, bald = 64, year = 2100, wyear = 21000, IBB = 0, kvoten = 1, Iyear = 0, Xage = 66): number {
  let result = 0;
  if (year > wyear && kvoten < 1) {
    pbb = IBB;
    pbb = pbb * kvoten;
  }
  let avdrag = 0;
  let Extra = 0;
  if (inkomst < 0.423 * pbb) {
    avdrag = inkomst;
  } else if (inkomst <= 0.99 * pbb) {
    avdrag = 0.423 * pbb;
  } else if (inkomst <= 2.72 * pbb) {
    avdrag = 0.423 * pbb + 0.2 * (inkomst - 0.99 * pbb);
  } else if (inkomst <= 3.11 * pbb) {
    avdrag = 0.77 * pbb;
  } else if (inkomst <= 7.88 * pbb) {
    avdrag = 0.77 * pbb - 0.1 * (inkomst - 3.11 * pbb);
  } else {
    avdrag = 0.293 * pbb;
  }
  if (inkomst <= 0.425 * pbb) {
    Extra = 0;
  } else if (inkomst <= 0.99 * pbb) {
    Extra = 0.425 * pbb;
  } else if (inkomst <= 2.72 * pbb) {
    Extra = 0.623 * pbb - 0.2 * inkomst;
  } else if (inkomst <= 2.94 * pbb) {
    Extra = 0.078 * pbb;
  } else if (inkomst <= 3.11 * pbb) {
    Extra = 0.372 * pbb - 0.1 * inkomst;
  } else if (inkomst <= 7.88 * pbb) {
    Extra = 0.061 * pbb;
  } else if (inkomst <= 8.49 * pbb) {
    Extra = 0.849 * pbb - 0.1 * inkomst;
  } else {
    Extra = 0 * pbb;
  }
  if (bald >= 66) avdrag = Extra + avdrag;
  if (avdrag > inkomst) avdrag = inkomst;
  if (marginal === 0) avdrag = vbaInt((avdrag + 99.99) / 100) * 100;
  result = avdrag;
  return result;
}

export function avdrag10(inkomst: number, pbb = 41000, marginal = 0, bald = 64, year = 2100, wyear = 21000, IBB = 0, kvoten = 1, Iyear = 0, Xage = 66): number {
  let result = 0;
  if (year > wyear && kvoten < 1) {
    pbb = IBB;
    pbb = pbb * kvoten;
  }
  let avdrag = 0;
  let Extra = 0;
  if (inkomst < 0.423 * pbb) {
    avdrag = inkomst;
  } else if (inkomst <= 0.99 * pbb) {
    avdrag = 0.423 * pbb;
  } else if (inkomst <= 2.72 * pbb) {
    avdrag = 0.225 * pbb + 0.2 * inkomst;
  } else if (inkomst <= 3.11 * pbb) {
    avdrag = 0.77 * pbb;
  } else if (inkomst <= 7.88 * pbb) {
    avdrag = 1.081 * pbb - 0.1 * inkomst;
  } else {
    avdrag = 0.293 * pbb;
  }
  if (inkomst <= 0.423 * pbb) {
    Extra = 0;
  } else if (inkomst <= 0.99 * pbb) {
    Extra = 0.5094 * pbb;
  } else if (inkomst <= 2.72 * pbb) {
    Extra = 0.7074 * pbb - 0.2 * inkomst;
  } else if (inkomst <= 2.94 * pbb) {
    Extra = 0.1624 * pbb;
  } else if (inkomst <= 3.11 * pbb) {
    Extra = 0.1624 * pbb;
  } else if (inkomst <= 3.9 * pbb) {
    Extra = -0.1486 * pbb + 0.1 * inkomst;
  } else if (inkomst <= 7.88 * pbb) {
    Extra = 0.2219 * pbb + 0.005 * inkomst;
  } else if (inkomst <= 9.1568 * pbb) {
    Extra = 1.0099 * pbb - 0.095 * inkomst;
  } else {
    Extra = 0.14 * pbb;
  }
  if (bald >= 66) avdrag = Extra + avdrag;
  if (avdrag > inkomst) avdrag = inkomst;
  if (marginal === 0) avdrag = vbaInt((avdrag + 99.99) / 100) * 100;
  result = avdrag;
  return result;
}

export function avdrag11(inkomst: number, pbb = 41000, marginal = 0, bald = 64, year = 2100, wyear = 21000, IBB = 0, kvoten = 1, Iyear = 0, Xage = 66): number {
  let result = 0;
  if (year > wyear && kvoten < 1) {
    pbb = IBB;
    pbb = pbb * kvoten;
  }
  let avdrag = 0;
  let Extra = 0;
  if (inkomst < 0.423 * pbb) {
    avdrag = inkomst;
  } else if (inkomst <= 0.99 * pbb) {
    avdrag = 0.423 * pbb;
  } else if (inkomst <= 2.72 * pbb) {
    avdrag = 0.225 * pbb + 0.2 * inkomst;
  } else if (inkomst <= 3.11 * pbb) {
    avdrag = 0.77 * pbb;
  } else if (inkomst <= 7.88 * pbb) {
    avdrag = 1.081 * pbb - 0.1 * inkomst;
  } else {
    avdrag = 0.293 * pbb;
  }
  if (inkomst <= 0.98 * pbb) {
    Extra = 0.557 * pbb;
  } else if (inkomst <= 0.99 * pbb) {
    Extra = 0.459 * pbb + 0.1 * inkomst;
  } else if (inkomst <= 2.72 * pbb) {
    Extra = 0.657 * pbb - 0.1 * inkomst;
  } else if (inkomst <= 3.11 * pbb) {
    Extra = 0.112 * pbb + 0.1 * inkomst;
  } else if (inkomst <= 3.85 * pbb) {
    Extra = -0.199 * pbb + 0.2 * inkomst;
  } else if (inkomst <= 4.8 * pbb) {
    Extra = 0.186 * pbb + 0.1 * inkomst;
  } else if (inkomst <= 7.88 * pbb) {
    Extra = 0.619 * pbb + 0.01 * inkomst;
  } else if (inkomst <= 12.21 * pbb) {
    Extra = 1.407 * pbb - 0.09 * inkomst;
  } else {
    Extra = 0.307 * pbb;
  }
  if (bald >= 66) avdrag = avdrag + Extra;
  if (avdrag > inkomst) avdrag = inkomst;
  if (marginal === 0) avdrag = vbaInt((avdrag + 99.99) / 100) * 100;
  result = avdrag;
  return result;
}

export function avdrag13(inkomst: number, pbb = 41000, marginal = 0, bald = 64, year = 2100, wyear = 21000, IBB = 0, kvoten = 1, Iyear = 0, Xage = 66): number {
  let result = 0;
  if (year > wyear && kvoten < 1) {
    pbb = IBB;
    pbb = pbb * kvoten;
  }
  let avdrag = 0;
  let Extra = 0;
  if (inkomst < 0.423 * pbb) {
    avdrag = inkomst;
  } else if (inkomst <= 0.99 * pbb) {
    avdrag = 0.423 * pbb;
  } else if (inkomst <= 2.72 * pbb) {
    avdrag = 0.225 * pbb + 0.2 * inkomst;
  } else if (inkomst <= 3.11 * pbb) {
    avdrag = 0.77 * pbb;
  } else if (inkomst <= 7.88 * pbb) {
    avdrag = 1.081 * pbb - 0.1 * inkomst;
  } else {
    avdrag = 0.293 * pbb;
  }
  let korr = 0;
  korr = 0;
  if (inkomst <= 0.99 * pbb) {
    Extra = 0.567 * pbb;
    korr = 0;
  } else if (inkomst <= 1.01 * pbb) {
    Extra = 0.785 * pbb - 0.2 * inkomst;
    korr = 0;
  } else if (inkomst <= 2.72 * pbb) {
    Extra = 0.674 * pbb - 0.09 * inkomst;
    korr = 0;
  } else if (inkomst <= 3.11 * pbb) {
    Extra = 0.129 * pbb + 0.11 * inkomst;
    korr = 0;
  } else if (inkomst <= 3.75 * pbb) {
    Extra = -0.182 * pbb + 0.21 * inkomst;
    korr = 0;
  } else if (inkomst <= 4.77 * pbb) {
    Extra = 0.233 * pbb + 0.1 * inkomst;
    korr = 0;
  } else if (inkomst <= 7.88 * pbb) {
    Extra = 0.66 * pbb + 0.01 * inkomst;
    korr = 0;
  } else if (inkomst <= 12.14 * pbb) {
    Extra = 1.448 * pbb - 0.09 * inkomst;
    korr = 0;
  } else {
    Extra = 0.357 * pbb;
    korr = 0;
  }
  if (bald >= 66) avdrag = avdrag + Extra + korr;
  if (avdrag > inkomst) avdrag = inkomst;
  if (marginal === 0) avdrag = vbaInt((avdrag + 99.99) / 100) * 100;
  result = avdrag;
  return result;
}

export function avdrag14(inkomst: number, pbb = 41000, marginal = 0, bald = 64, year = 2100, wyear = 21000, IBB = 0, kvoten = 1, Iyear = 0, Xage = 66): number {
  let result = 0;
  if (year > wyear && kvoten < 1) {
    pbb = IBB;
    pbb = pbb * kvoten;
  }
  let avdrag = 0;
  let Extra = 0;
  if (inkomst < 0.423 * pbb) {
    avdrag = inkomst;
  } else if (inkomst <= 0.99 * pbb) {
    avdrag = 0.423 * pbb;
  } else if (inkomst <= 2.72 * pbb) {
    avdrag = 0.225 * pbb + 0.2 * inkomst;
  } else if (inkomst <= 3.11 * pbb) {
    avdrag = 0.77 * pbb;
  } else if (inkomst <= 7.88 * pbb) {
    avdrag = 1.081 * pbb - 0.1 * inkomst;
  } else {
    avdrag = 0.293 * pbb;
  }
  if (inkomst <= 0.99 * pbb) {
    Extra = 0.682 * pbb;
  } else if (inkomst <= 1.105 * pbb) {
    Extra = 0.88 * pbb - 0.2 * inkomst;
  } else if (inkomst <= 2.72 * pbb) {
    Extra = 0.753 * pbb - 0.085 * inkomst;
  } else if (inkomst <= 3.11 * pbb) {
    Extra = 0.208 * pbb + 0.115 * inkomst;
  } else if (inkomst <= 3.69 * pbb) {
    Extra = -0.103 * pbb + 0.215 * inkomst;
  } else if (inkomst <= 4.785 * pbb) {
    Extra = 0.322 * pbb + 0.1 * inkomst;
  } else if (inkomst <= 7.88 * pbb) {
    Extra = 0.753 * pbb + 0.01 * inkomst;
  } else if (inkomst <= 12.43 * pbb) {
    Extra = 1.541 * pbb - 0.09 * inkomst;
  } else {
    Extra = 0.422 * pbb;
  }
  if (bald >= 66) avdrag = avdrag + Extra;
  if (avdrag > inkomst) avdrag = inkomst;
  if (marginal === 0) avdrag = vbaInt((avdrag + 99.99) / 100) * 100;
  result = avdrag;
  return result;
}

export function avdrag16(inkomst: number, pbb = 41000, marginal = 0, bald = 64, year = 2100, wyear = 21000, IBB = 0, kvoten = 1, Iyear = 0, Xage = 66): number {
  let result = 0;
  if (year > wyear && kvoten < 1) {
    pbb = IBB;
    pbb = pbb * kvoten;
  }
  let avdrag = 0;
  let Extra = 0;
  if (inkomst < 0.423 * pbb) {
    avdrag = inkomst;
  } else if (inkomst <= 0.99 * pbb) {
    avdrag = 0.423 * pbb;
  } else if (inkomst <= 2.72 * pbb) {
    avdrag = 0.225 * pbb + 0.2 * inkomst;
  } else if (inkomst <= 3.11 * pbb) {
    avdrag = 0.77 * pbb;
  } else if (inkomst <= 7.88 * pbb) {
    avdrag = 1.081 * pbb - 0.1 * inkomst;
  } else {
    avdrag = 0.293 * pbb;
  }
  if (inkomst <= 0.99 * pbb) {
    Extra = 0.687 * pbb;
  } else if (inkomst <= 1.11 * pbb) {
    Extra = 0.885 * pbb - 0.2 * inkomst;
  } else if (inkomst <= 2.72 * pbb) {
    Extra = 0.609 * pbb + 0.049 * inkomst;
  } else if (inkomst <= 3.11 * pbb) {
    Extra = 0.741 * pbb;
  } else if (inkomst <= 3.77 * pbb) {
    Extra = 0.43 * pbb + 0.1 * inkomst;
  } else if (inkomst <= 5.4 * pbb) {
    Extra = 0.807 * pbb;
  } else if (inkomst <= 7.88 * pbb) {
    Extra = 0.753 * pbb + 0.01 * inkomst;
  } else if (inkomst <= 12.43 * pbb) {
    Extra = 1.541 * pbb - 0.09 * inkomst;
  } else {
    Extra = 0.422 * pbb;
  }
  if (bald >= 66) avdrag = avdrag + Extra;
  if (avdrag > inkomst) avdrag = inkomst;
  if (marginal === 0) avdrag = vbaInt((avdrag + 99.99) / 100) * 100;
  result = avdrag;
  return result;
}

export function avdrag18(inkomst: number, pbb = 45500, marginal = 0, bald = 64, year = 2100, wyear = 21000, IBB = 0, kvoten = 1, Iyear = 0, Xage = 66): number {
  let result = 0;
  if (marginal === 0) inkomst = vbaInt(inkomst / 100) * 100;
  if (year > wyear && kvoten < 1) {
    pbb = IBB;
    pbb = pbb * kvoten;
  }
  let avdrag = 0;
  let Extra = 0;
  if (inkomst <= 0.99 * pbb) {
    avdrag = 0.423 * pbb;
  } else if (inkomst <= 2.72 * pbb) {
    avdrag = 0.423 * pbb + 0.2 * (inkomst - 0.99 * pbb);
  } else if (inkomst <= 3.11 * pbb) {
    avdrag = 0.77 * pbb;
  } else if (inkomst <= 7.88 * pbb) {
    avdrag = 0.77 * pbb - 0.1 * (inkomst - 3.11 * pbb);
  } else {
    avdrag = 0.293 * pbb;
  }
  if (inkomst <= 0.99 * pbb) {
    Extra = 0.687 * pbb;
  } else if (inkomst <= 1.11 * pbb) {
    Extra = 0.885 * pbb - 0.2 * inkomst;
  } else if (inkomst <= 2.72 * pbb) {
    Extra = 0.609 * pbb + 0.049 * inkomst;
  } else if (inkomst <= 2.94 * pbb) {
    Extra = -0.162 * pbb + 0.332 * inkomst;
  } else if (inkomst <= 3.11 * pbb) {
    Extra = 0.482 * pbb + 0.113 * inkomst;
  } else if (inkomst <= 4.45 * pbb) {
    Extra = 0.171 * pbb + 0.213 * inkomst;
  } else if (inkomst <= 7.88 * pbb) {
    Extra = 1.376 * pbb - 0.058 * inkomst;
  } else if (inkomst <= 9.15 * pbb) {
    Extra = 2.164 * pbb - 0.158 * inkomst;
  } else if (inkomst <= 12.43 * pbb) {
    Extra = 1.541 * pbb - 0.09 * inkomst;
  } else {
    Extra = 0.422 * pbb;
  }
  if (bald > 65) avdrag = avdrag + Extra;
  if (avdrag > inkomst) avdrag = inkomst;
  if (marginal === 0) avdrag = vbaInt((avdrag + 99.99) / 100) * 100;
  result = avdrag;
  return result;
}

export function avdrag19(inkomst: number, pbb = 46500, marginal = 0, bald = 64, year = 2100, wyear = 21000, IBB = 0, kvoten = 1, Iyear = 0, Xage = 66): number {
  let result = 0;
  if (marginal === 0) inkomst = vbaInt(inkomst / 100) * 100;
  if (year > wyear && kvoten < 1) {
    pbb = IBB;
    pbb = pbb * kvoten;
  }
  let avdrag = 0;
  let Extra = 0;
  if (inkomst <= 0.99 * pbb) {
    avdrag = 0.423 * pbb;
  } else if (inkomst <= 2.72 * pbb) {
    avdrag = 0.423 * pbb + 0.2 * (inkomst - 0.99 * pbb);
  } else if (inkomst <= 3.11 * pbb) {
    avdrag = 0.77 * pbb;
  } else if (inkomst <= 7.88 * pbb) {
    avdrag = 0.77 * pbb - 0.1 * (inkomst - 3.11 * pbb);
  } else {
    avdrag = 0.293 * pbb;
  }
  if (inkomst <= 0.99 * pbb) {
    Extra = 0.687 * pbb;
  } else if (inkomst <= 1.11 * pbb) {
    Extra = 0.885 * pbb - 0.2 * inkomst;
  } else if (inkomst <= 2.72 * pbb) {
    Extra = 0.6 * pbb + 0.057 * inkomst;
  } else if (inkomst <= 3.11 * pbb) {
    Extra = -0.169 * pbb + 0.34 * inkomst;
  } else if (inkomst <= 3.21 * pbb) {
    Extra = -0.48 * pbb + 0.44 * inkomst;
  } else if (inkomst <= 4.45 * pbb) {
    Extra = 0.207 * pbb + 0.228 * inkomst;
  } else if (inkomst <= 5.31 * pbb) {
    Extra = 1.397 * pbb - 0.039 * inkomst;
  } else if (inkomst <= 7.88 * pbb) {
    Extra = 0.763 * pbb + 0.08 * inkomst;
  } else if (inkomst <= 8.08 * pbb) {
    Extra = 1.551 * pbb - 0.02 * inkomst;
  } else if (inkomst <= 13.54 * pbb) {
    Extra = 2.399 * pbb - 0.125 * inkomst;
  } else if (inkomst <= 34 * pbb) {
    Extra = 1.031 * pbb - 0.024 * inkomst;
  } else {
    Extra = 0.215 * pbb;
  }
  if (bald > 65) avdrag = avdrag + Extra;
  if (avdrag > inkomst) avdrag = inkomst;
  if (marginal === 0) avdrag = vbaInt((avdrag + 99.99) / 100) * 100;
  result = avdrag;
  return result;
}

export function avdrag20(inkomst: number, pbb = 47300, marginal = 0, bald = 64, year = 2100, wyear = 2100, IBB = 0, kvoten = 1, Iyear = 0, Xage = 66): number {
  let result = 0;
  if (marginal === 0) inkomst = vbaInt(inkomst / 100) * 100;
  if ((year > Iyear && Iyear > 0) && kvoten < 1) {
    pbb = IBB;
    pbb = pbb * kvoten;
  }
  let avdrag = 0;
  let Extra = 0;
  if (inkomst <= 0.99 * pbb) {
    avdrag = 0.423 * pbb;
  } else if (inkomst <= 2.72 * pbb) {
    avdrag = 0.423 * pbb + 0.2 * (inkomst - 0.99 * pbb);
  } else if (inkomst <= 3.11 * pbb) {
    avdrag = 0.77 * pbb;
  } else if (inkomst <= 7.88 * pbb) {
    avdrag = 0.77 * pbb - 0.1 * (inkomst - 3.11 * pbb);
  } else {
    avdrag = 0.293 * pbb;
  }
  if (inkomst <= 0.99 * pbb) {
    Extra = 0.687 * pbb;
  } else if (inkomst <= 1.11 * pbb) {
    Extra = 0.885 * pbb - 0.2 * inkomst;
  } else if (inkomst <= 2.72 * pbb) {
    Extra = 0.6 * pbb + 0.057 * inkomst;
  } else if (inkomst <= 3.11 * pbb) {
    Extra = -0.169 * pbb + 0.34 * inkomst;
  } else if (inkomst <= 3.21 * pbb) {
    Extra = -0.48 * pbb + 0.44 * inkomst;
  } else if (inkomst <= 4.45 * pbb) {
    Extra = 0.207 * pbb + 0.228 * inkomst;
  } else if (inkomst <= 7.88 * pbb) {
    Extra = 0.488 * pbb + 0.165 * inkomst;
  } else if (inkomst <= 8.08 * pbb) {
    Extra = 1.276 * pbb + 0.065 * inkomst;
  } else if (inkomst <= 11.06 * pbb) {
    Extra = 2.205 * pbb - 0.05 * inkomst;
  } else if (inkomst <= 12.15 * pbb) {
    Extra = 7.182 * pbb - 0.5 * inkomst;
  } else if (inkomst <= 29.65 * pbb) {
    Extra = 1.654 * pbb - 0.045 * inkomst;
  } else if (inkomst <= 34 * pbb) {
    Extra = 1.031 * pbb - 0.024 * inkomst;
  } else {
    Extra = 0.215 * pbb;
  }
  if (bald >= Xage) avdrag = avdrag + Extra;
  if (avdrag > inkomst) avdrag = inkomst;
  if (marginal === 0) avdrag = vbaInt((avdrag + 99.99) / 100) * 100;
  result = avdrag;
  return result;
}

export function avdrag21(inkomst: number, pbb = 47300, marginal = 0, bald = 64, year = 2100, wyear = 2100, IBB = 0, kvoten = 1, Iyear = 0, Xage = 66): number {
  let result = 0;
  if (marginal === 0) inkomst = vbaInt(inkomst / 100) * 100;
  if ((year > Iyear && Iyear > 0) && kvoten < 1) {
    pbb = IBB;
    pbb = pbb * kvoten;
  }
  let avdrag = 0;
  let Extra = 0;
  if (inkomst <= 0.99 * pbb) {
    avdrag = 0.423 * pbb;
  } else if (inkomst <= 2.72 * pbb) {
    avdrag = 0.423 * pbb + 0.2 * (inkomst - 0.99 * pbb);
  } else if (inkomst <= 3.11 * pbb) {
    avdrag = 0.77 * pbb;
  } else if (inkomst <= 7.88 * pbb) {
    avdrag = 0.77 * pbb - 0.1 * (inkomst - 3.11 * pbb);
  } else {
    avdrag = 0.293 * pbb;
  }
  if (inkomst <= 0.99 * pbb) {
    Extra = 0.687 * pbb;
  } else if (inkomst <= 1.11 * pbb) {
    Extra = 0.885 * pbb - 0.2 * inkomst;
  } else if (inkomst <= 2.72 * pbb) {
    Extra = 0.6 * pbb + 0.057 * inkomst;
  } else if (inkomst <= 3.11 * pbb) {
    Extra = -0.169 * pbb + 0.34 * inkomst;
  } else if (inkomst <= 3.21 * pbb) {
    Extra = -0.48 * pbb + 0.44 * inkomst;
  } else if (inkomst <= 7.88 * pbb) {
    Extra = 0.207 * pbb + 0.228 * inkomst;
  } else if (inkomst <= 8.08 * pbb) {
    Extra = 0.995 * pbb + 0.128 * inkomst;
  } else if (inkomst <= 11.28 * pbb) {
    Extra = 2.029 * pbb;
  } else if (inkomst <= 12.53 * pbb) {
    Extra = 9.023 * pbb - 0.62 * inkomst;
  } else if (inkomst <= 13.54 * pbb) {
    Extra = 1.253 * pbb;
  } else if (inkomst <= 35.36 * pbb) {
    Extra = 2.03 * pbb - 0.0574 * inkomst;
  } else {
    Extra = 0;
  }
  if (bald >= Xage) avdrag = avdrag + Extra;
  if (avdrag > inkomst) avdrag = inkomst;
  if (marginal === 0) avdrag = vbaInt((avdrag + 99.99) / 100) * 100;
  result = avdrag;
  return result;
}

export function avdrag22(inkomst: number, pbb = 47300, marginal = 0, bald = 64, year = 2100, wyear = 2100, IBB = 0, kvoten = 1, Iyear = 0, Xage = 66): number {
  let result = 0;
  if (marginal === 0) inkomst = vbaInt(inkomst / 100) * 100;
  if ((year > Iyear && Iyear > 0) && kvoten < 1) {
    pbb = IBB;
    pbb = pbb * kvoten;
  }
  let avdrag = 0;
  let Extra = 0;
  if (inkomst <= 0.99 * pbb) {
    avdrag = 0.423 * pbb;
  } else if (inkomst <= 2.72 * pbb) {
    avdrag = 0.423 * pbb + 0.2 * (inkomst - 0.99 * pbb);
  } else if (inkomst <= 3.11 * pbb) {
    avdrag = 0.77 * pbb;
  } else if (inkomst <= 7.88 * pbb) {
    avdrag = 0.77 * pbb - 0.1 * (inkomst - 3.11 * pbb);
  } else {
    avdrag = 0.293 * pbb;
  }
  if (inkomst <= 0.91 * pbb) {
    Extra = 0.687 * pbb;
  } else if (inkomst <= 1.11 * pbb) {
    Extra = 0.885 * pbb - 0.2 * inkomst;
  } else if (inkomst <= 1.965 * pbb) {
    Extra = 0.6 * pbb + 0.057 * inkomst;
  } else if (inkomst <= 2.72 * pbb) {
    Extra = 0.333 * pbb + 0.1949 * inkomst;
  } else if (inkomst <= 3.11 * pbb) {
    Extra = -0.212 * pbb + 0.3949 * inkomst;
  } else if (inkomst <= 3.24 * pbb) {
    Extra = -0.523 * pbb + 0.4949 * inkomst;
  } else if (inkomst <= 5.53 * pbb) {
    Extra = 0.325 * pbb + 0.233 * inkomst;
  } else if (inkomst <= 7.88 * pbb) {
    Extra = 0.441 * pbb + 0.212 * inkomst;
  } else if (inkomst <= 8.08 * pbb) {
    Extra = 1.104 * pbb + 0.128 * inkomst;
  } else if (inkomst <= 11.48 * pbb) {
    Extra = 2.139 * pbb;
  } else if (inkomst <= 12.8 * pbb) {
    Extra = 9.257 * pbb - 0.62 * inkomst;
  } else if (inkomst <= 13.54 * pbb) {
    Extra = 1.32 * pbb;
  } else if (inkomst <= 36.54 * pbb) {
    Extra = 2.097 * pbb - 0.0574 * inkomst;
  } else {
    Extra = 0;
  }
  if (bald >= Xage) avdrag = avdrag + Extra;
  if (avdrag > inkomst) avdrag = inkomst;
  if (marginal === 0) avdrag = vbaInt((avdrag + 99.99) / 100) * 100;
  result = avdrag;
  return result;
}

export function avdrag24(inkomst: number, pbb = 47300, marginal = 0, bald = 64, year = 2100, wyear = 2100, IBB = 0, kvoten = 1, Iyear = 0, Xage = 66): number {
  let result = 0;
  if (marginal === 0) inkomst = vbaInt(inkomst / 100) * 100;
  if ((year > Iyear && Iyear > 0) && kvoten < 1) {
    pbb = IBB;
    pbb = pbb * kvoten;
  }
  let avdrag = 0;
  let Extra = 0;
  if (inkomst <= 0.99 * pbb) {
    avdrag = 0.423 * pbb;
  } else if (inkomst <= 2.72 * pbb) {
    avdrag = 0.225 * pbb + 0.2 * inkomst;
  } else if (inkomst <= 3.11 * pbb) {
    avdrag = 0.77 * pbb;
  } else if (inkomst <= 7.88 * pbb) {
    avdrag = 1.081 * pbb - 0.1 * inkomst;
  } else {
    avdrag = 0.293 * pbb;
  }
  if (inkomst < 0.91 * pbb) {
    Extra = 0.687 * pbb;
  } else if (inkomst < 1.11 * pbb) {
    Extra = 0.885 * pbb - 0.2 * inkomst;
  } else if (inkomst < 1.965 * pbb) {
    Extra = 0.6 * pbb + 0.057 * inkomst;
  } else if (inkomst < 2.72 * pbb) {
    Extra = 0.333 * pbb + 0.1949 * inkomst;
  } else if (inkomst < 3.11 * pbb) {
    Extra = -0.212 * pbb + 0.3949 * inkomst;
  } else if (inkomst < 3.24 * pbb) {
    Extra = -0.523 * pbb + 0.4949 * inkomst;
  } else if (inkomst < 5 * pbb) {
    Extra = 0.208 * pbb + 0.2693 * inkomst;
  } else if (inkomst < 7.88 * pbb) {
    Extra = 0.3 * pbb + 0.2513 * inkomst;
  } else if (inkomst < 8.08 * pbb) {
    Extra = 0.986 * pbb + 0.1643 * inkomst;
  } else if (inkomst < 10.74 * pbb) {
    Extra = 2.313 * pbb;
  } else if (inkomst < 12.16 * pbb) {
    Extra = 8.972 * pbb - 0.62 * inkomst;
  } else if (inkomst < 13.54 * pbb) {
    Extra = 1.43 * pbb;
  } else if (inkomst < 38.42 * pbb) {
    Extra = 2.206 * pbb - 0.0574 * inkomst;
  } else {
    Extra = 0;
  }
  if (bald >= Xage) avdrag = avdrag + Extra;
  if (avdrag > inkomst) avdrag = inkomst;
  if (marginal === 0) avdrag = vbaInt((avdrag + 99.99) / 100) * 100;
  result = avdrag;
  return result;
}

export function avdrag25(inkomst: number, pbb = 47300, marginal = 0, bald = 64, year = 2100, wyear = 2100, IBB = 0, kvoten = 1, Iyear = 0, Xage = 66): number {
  let result = 0;
  if (marginal === 0) inkomst = vbaInt(inkomst / 100) * 100;
  if ((year > Iyear && Iyear > 0) && kvoten < 1) {
    pbb = IBB;
    pbb = pbb * kvoten;
  }
  let avdrag = 0;
  let Extra = 0;
  if (inkomst <= 0.99 * pbb) {
    avdrag = 0.423 * pbb;
  } else if (inkomst <= 2.72 * pbb) {
    avdrag = 0.423 * pbb + 0.2 * (inkomst - 0.99 * pbb);
  } else if (inkomst <= 3.11 * pbb) {
    avdrag = 0.77 * pbb;
  } else if (inkomst <= 7.88 * pbb) {
    avdrag = 0.77 * pbb - 0.1 * (inkomst - 3.11 * pbb);
  } else {
    avdrag = 0.293 * pbb;
  }
  if (inkomst < 0.91 * pbb) {
    Extra = 0.687 * pbb;
  } else if (inkomst < 1.11 * pbb) {
    Extra = 0.885 * pbb - 0.2 * inkomst;
  } else if (inkomst < 1.965 * pbb) {
    Extra = 0.6 * pbb + 0.057 * inkomst;
  } else if (inkomst < 2.72 * pbb) {
    Extra = 0.333 * pbb + 0.1949 * inkomst;
  } else if (inkomst < 3.11 * pbb) {
    Extra = -0.212 * pbb + 0.3949 * inkomst;
  } else if (inkomst < 3.24 * pbb) {
    Extra = -0.523 * pbb + 0.4949 * inkomst;
  } else if (inkomst < 5 * pbb) {
    Extra = 0.096 * pbb + 0.304 * inkomst;
  } else if (inkomst < 7.88 * pbb) {
    Extra = 0.186 * pbb + 0.286 * inkomst;
  } else if (inkomst < 8.08 * pbb) {
    Extra = 0.872 * pbb + 0.199 * inkomst;
  } else if (inkomst < 10.94 * pbb) {
    Extra = 2.48 * pbb;
  } else if (inkomst < 12.47 * pbb) {
    Extra = 9.263 * pbb - 0.62 * inkomst;
  } else {
    Extra = 1.532 * pbb;
  }
  if (bald >= Xage) avdrag = avdrag + Extra;
  if (avdrag > inkomst) avdrag = inkomst;
  if (marginal === 0) avdrag = vbaInt((avdrag + 99.99) / 100) * 100;
  result = avdrag;
  return result;
}

export function avdrag26(inkomst: number, pbb = 47300, marginal = 0, bald = 64, year = 2100, wyear = 2100, IBB = 0, kvoten = 1, Iyear = 0, Xage = 66): number {
  let result = 0;
  if (marginal === 0) inkomst = vbaInt(inkomst / 100) * 100;
  if ((year > Iyear && Iyear > 0) && kvoten < 1) {
    pbb = IBB;
    pbb = pbb * kvoten;
  }
  let avdrag = 0;
  let Extra = 0;
  if (inkomst <= 0.99 * pbb) {
    avdrag = 0.423 * pbb;
  } else if (inkomst <= 2.72 * pbb) {
    avdrag = 0.423 * pbb + 0.2 * (inkomst - 0.99 * pbb);
  } else if (inkomst <= 3.11 * pbb) {
    avdrag = 0.77 * pbb;
  } else if (inkomst <= 7.88 * pbb) {
    avdrag = 0.77 * pbb - 0.1 * (inkomst - 3.11 * pbb);
  } else {
    avdrag = 0.293 * pbb;
  }
  if (bald >= Xage) {
    if (inkomst < 0.91 * pbb) {
      Extra = 0.687 * pbb;
    } else if (inkomst < 1.11 * pbb) {
      Extra = 0.885 * pbb - 0.2 * inkomst;
    } else if (inkomst < 1.965 * pbb) {
      Extra = 0.6 * pbb + 0.057 * inkomst;
    } else if (inkomst < 2.72 * pbb) {
      Extra = 0.333 * pbb + 0.1949 * inkomst;
    } else if (inkomst < 3.11 * pbb) {
      Extra = -0.212 * pbb + 0.3949 * inkomst;
    } else if (inkomst < 3.24 * pbb) {
      Extra = -0.523 * pbb + 0.4949 * inkomst;
    } else if (inkomst < 5 * pbb) {
      Extra = -0.073 * pbb + 0.356 * inkomst;
    } else if (inkomst < 7.88 * pbb) {
      Extra = 0.017 * pbb + 0.338 * inkomst;
    } else if (inkomst < 8.08 * pbb) {
      Extra = 0.703 * pbb + 0.251 * inkomst;
    } else if (inkomst < 11.16 * pbb) {
      Extra = 2.732 * pbb;
    } else if (inkomst < 12.84 * pbb) {
      Extra = 9.652 * pbb - 0.62 * inkomst;
    } else {
      Extra = 1.691 * pbb;
    }
    avdrag = avdrag + Extra;
  }
  if (avdrag > inkomst) avdrag = inkomst;
  if (marginal === 0) avdrag = vbaInt((avdrag + 99.99) / 100) * 100;
  result = avdrag;
  return result;
}

/**
 * Picks the grundavdrag rules for a year. Mirrors `avdragxx`.
 *
 * `Iyear` lets a run pin the rules to one year while the income year moves --
 * the "unchanged rules" assumption the advanced settings offer.
 */
export function avdragxx(
  inkomst: number,
  pbb = 52500,
  marginal = 0,
  bald = 64,
  year = 2100,
  wyear = 21000,
  IBB = 0,
  kvoten = 1,
  IyearIn = 0,
  xage = 66,
): number {
  const Iyear = IyearIn === 0 ? year : IyearIn;
  const a = [inkomst, pbb, marginal, bald, year, wyear, IBB, kvoten, Iyear] as const;

  if (Iyear < 1996) return avdrag95(...a);
  if (Iyear < 2001) return avdrag96(...a);
  if (Iyear < 2003) return avdrag01(...a);
  if (Iyear < 2005) return avdrag03(...a);
  if (Iyear === 2005) return avdrag05(...a);
  if (Iyear < 2009) return avdrag06(...a);
  if (Iyear < 2010) return avdrag09(...a);
  if (Iyear < 2011) return avdrag10(...a);
  if (Iyear < 2013) return avdrag11(...a);
  if (Iyear < 2014) return avdrag13(...a);
  if (Iyear < 2016) return avdrag14(...a);
  if (Iyear < 2018) return avdrag16(...a);
  if (Iyear < 2019) return avdrag18(...a);
  if (Iyear < 2020) return avdrag19(...a);
  // From 2020 the higher allowance for the old follows Xage.
  if (Iyear < 2021) return avdrag20(...a, xage);
  if (Iyear < 2022) return avdrag21(...a, xage);
  if (Iyear <= 2023) return avdrag22(...a, xage);
  if (Iyear < 2025) return avdrag24(...a, xage);
  if (Iyear === 2025) return avdrag25(...a, xage);
  return avdrag26(...a, xage);
}
