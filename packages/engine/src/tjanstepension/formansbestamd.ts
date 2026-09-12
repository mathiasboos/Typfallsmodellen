/**
 * Defined-benefit occupational pension.
 *
 * Port of `TjänstepensionerFörmån.bas` (`FTJP`) together with the benefit
 * functions it calls -- `DC_underlag`, `tlPA03`, `KAPKL_f`, `PA_KL` and
 * `PA_KLBPP`.
 *
 * Four agreements carry a benefit-based amount on top of their premiums: ITP 2,
 * SAF-LO (through the legacy STP), KAP-KL and PA16 avdelning 2. `FTJP` adds it
 * to whatever the defined contribution part has already produced.
 *
 * These functions read a great deal of Mcalc's state -- the wage vector, the
 * basbelopp series, the income year for each age. That is gathered into
 * `RunVectors` and passed in.
 */

import type { ModelContext } from "../model/context.js";
import type { RunVectors } from "../model/runVectors.js";
import type { DeltalTables } from "../pension/deltal.js";
import { deltal } from "../pension/incomePension.js";
import { vbaInt, vbaRound, wsLarge } from "../vba/math.js";
import { STP_ } from "./safLo.js";
import { tlITP2F } from "./itp.js";
import { tjp_ddeltal } from "./tjpkassa.js";
export type { RunVectors };

import { Scheme } from "./types.js";
import type { SchemeContext, SchemeId } from "./types.js";

/**
 * The pensionable salary basis, after the salary-capping rules.
 *
 * Averages the five years before retirement, in fixed prices, having first held
 * back rises that outpaced the income base amount by more than the agreement
 * allows. `stat` of 1 skips the capping, which does not apply in the state
 * sector.
 */
export function DC_underlag(
  stat: 0 | 1,
  tjpPar: number,
  vectors: RunVectors,
  marginal: 0 | 1 = 0,
): number {
  const zpar = tjpPar > 65 ? 65 : tjpPar;
  const at = vbaInt(zpar);

  let w1 = vectors.wage(at - 1);
  let w2 = vectors.wage(at - 2);
  let w3 = vectors.wage(at - 3);
  let w4 = vectors.wage(at - 4);
  let w5 = vectors.wage(at - 5);
  const w6 = vectors.wage(at - 6);

  const pbb1 = vectors.pbb(at - 1);
  const pbb2 = vectors.pbb(at - 2);
  const pbb3 = vectors.pbb(at - 3);
  const pbb4 = vectors.pbb(at - 4);
  const pbb5 = vectors.pbb(at - 5);

  const ibb1 = vectors.ibb(at - 1);
  const ibb2 = vectors.ibb(at - 2);
  const ibb3 = vectors.ibb(at - 3);
  const ibb4 = vectors.ibb(at - 4);
  const ibb5 = vectors.ibb(at - 5);
  const ibb6 = vectors.ibb(at - 6);

  // A gap in any of the six years and there is no basis at all.
  if (w6 === 0 || w5 === 0 || w4 === 0 || w3 === 0 || w2 === 0 || w1 === 0) return 0;

  if (stat === 0) {
    // Collectum's salary capping: the allowance tightens each year closer to
    // retirement, from 20% down to nothing in the final year.
    if (w5 / w6 > (1.2 * ibb5) / ibb6) w5 = w6 * ((1.2 * ibb6) / ibb5);
    if (w4 / w5 > (1.15 * ibb4) / ibb5) w4 = w5 * ((1.15 * ibb4) / ibb5);
    if (w3 / w4 > (1.1 * ibb3) / ibb4) w3 = w4 * ((1.1 * ibb3) / ibb4);
    if (w2 / w3 > (1.05 * ibb2) / ibb3) w2 = w3 * ((1.05 * ibb2) / ibb3);
    if (w1 / w2 > ibb1 / ibb2) w1 = w2 * (ibb1 / ibb2);
  }

  // Then restate each year in the price level of the year before retirement.
  w5 = (w5 * pbb1) / pbb5;
  w4 = (w4 * pbb1) / pbb4;
  w3 = (w3 * pbb1) / pbb3;
  w2 = (w2 * pbb1) / pbb2;

  const average = (w1 + w2 + w3 + w4 + w5) / 5;
  return marginal === 0 ? vbaInt(average + 0.5) : average;
}

/**
 * PA03 -- the defined benefit for central government employees in avdelning 2.
 *
 * The replacement rates fall cohort by cohort as the benefit is phased out,
 * reaching zero below the breakpoint for those born 1973 or later.
 */
export function tlPA03(inkomst: number, aar: number, inkbas: number, fodar: number, scheme: SchemeId): number {
  if (inkomst <= 0) return 0;
  // Avdelning 1 has no defined benefit at all.
  if (scheme === Scheme.Pa16Avd1) return 0;

  let par1: number;
  let par2: number;
  let par3: number;

  if (fodar < 1943) {
    [par1, par2, par3] = [0.1, 0.65, 0.325];
  } else if (fodar >= 1973) {
    [par1, par2, par3] = [0, 0.6, 0.3];
  } else {
    // 1943 to 1972, one row per cohort.
    const table: Readonly<Record<number, [number, number, number]>> = {
      1943: [0.095, 0.6485, 0.324],
      1944: [0.093, 0.647, 0.323],
      1945: [0.091, 0.6455, 0.322],
      1946: [0.089, 0.644, 0.321],
      1947: [0.087, 0.6425, 0.32],
      1948: [0.084, 0.641, 0.319],
      1949: [0.082, 0.6395, 0.318],
      1950: [0.079, 0.638, 0.317],
      1951: [0.077, 0.6365, 0.316],
      1952: [0.074, 0.635, 0.315],
      1953: [0.072, 0.6335, 0.314],
      1954: [0.069, 0.632, 0.313],
      1955: [0.066, 0.6305, 0.312],
      1956: [0.063, 0.629, 0.311],
      1957: [0.06, 0.6275, 0.31],
      1958: [0.057, 0.626, 0.309],
      1959: [0.054, 0.6245, 0.308],
      1960: [0.051, 0.623, 0.307],
      1961: [0.047, 0.6215, 0.306],
      1962: [0.043, 0.62, 0.305],
      1963: [0.039, 0.6185, 0.304],
      1964: [0.036, 0.617, 0.303],
      1965: [0.032, 0.615, 0.302],
      1966: [0.029, 0.613, 0.301],
      1967: [0.025, 0.611, 0.3],
      1968: [0.021, 0.609, 0.3],
      1969: [0.017, 0.607, 0.3],
      1970: [0.013, 0.605, 0.3],
      1971: [0.009, 0.603, 0.3],
      1972: [0.005, 0.601, 0.3],
    };
    [par1, par2, par3] = table[vbaInt(fodar)] ?? [0, 0.6, 0.3];
  }

  let benefit: number;
  if (inkomst > 30 * inkbas) {
    benefit = (30 - 20) * inkbas * par3 + (20 - 7.5) * par2 * inkbas + 7.5 * inkbas * par1;
  } else if (inkomst > 20 * inkbas) {
    benefit = (inkomst - 20 * inkbas) * par3 + (20 - 7.5) * par2 * inkbas + 7.5 * inkbas * par1;
  } else if (inkomst > 7.5 * inkbas) {
    benefit = (inkomst - 7.5 * inkbas) * par2 + 7.5 * inkbas * par1;
  } else {
    benefit = inkomst * par1;
  }

  if (aar < 30) benefit = (benefit * aar) / 30;
  return benefit;
}

/**
 * KAP-KL's defined benefit, which applies only above 7.5 income base amounts.
 *
 * Replacement rates fall cohort by cohort, and reach zero for those born 1986
 * or later, who are on AKAP-KL instead.
 */
export function KAPKL_f(arsmedel: number, aar: number, ibb: number, fodar: number): number {
  let par2: number;
  let par3: number;

  const cohort = vbaInt(fodar);
  if (cohort <= 1946) {
    [par2, par3] = [0.625, 0.3125];
  } else if (cohort >= 1986) {
    [par2, par3] = [0, 0];
  } else {
    const table: Readonly<Record<number, [number, number]>> = {
      1947: [0.6214, 0.3107],
      1948: [0.6179, 0.3089],
      1949: [0.6143, 0.3071],
      1950: [0.6107, 0.3054],
      1951: [0.6071, 0.3036],
      1952: [0.6036, 0.3018],
      1953: [0.6, 0.3],
      1954: [0.5964, 0.2982],
      1955: [0.5929, 0.2964],
      1956: [0.5893, 0.2946],
      1957: [0.5857, 0.2929],
      1958: [0.5821, 0.2911],
      1959: [0.5786, 0.2893],
      1960: [0.575, 0.2875],
      1961: [0.5714, 0.2857],
      1962: [0.5679, 0.2839],
      1963: [0.5643, 0.2821],
      1964: [0.5607, 0.2804],
      1965: [0.5571, 0.2786],
      // Note 0.268 here, where the sequence would suggest 0.2768. As written.
      1966: [0.5536, 0.268],
    };
    [par2, par3] = table[cohort] ?? [0.55, 0.275];
  }

  let benefit: number;
  if (arsmedel > 30 * ibb) {
    benefit = (30 - 20) * ibb * par3 + (20 - 7.5) * par2 * ibb;
  } else if (arsmedel > 20 * ibb) {
    benefit = (arsmedel - 20 * ibb) * par3 + (20 - 7.5) * par2 * ibb;
  } else if (arsmedel > 7.5 * ibb) {
    benefit = (arsmedel - 7.5 * ibb) * par2;
  } else {
    benefit = 0;
  }

  if (aar < 30) benefit = (benefit * aar) / 30;
  return benefit;
}

/**
 * PA-KL gross pension points, from the best five of the seven years ending two
 * years before retirement.
 *
 * QUIRK: the fallbacks for a short working life never fire. Each reads
 * `ElseIf yearpoint7 = 0 & yearpoint6 <> 0`, and VBA's `&` is string
 * concatenation, not `And`. The expression parses as
 * `yearpoint7 = ((0 & yearpoint6) <> 0)` -- a comparison against a boolean,
 * so `yearpoint7 = True`, i.e. -1, which a salary ratio never is. With fewer
 * than seven years of salary the chain therefore falls through to the final
 * `Else` and averages the single best year, never the four, three or two the
 * intermediate branches intend. Kept as written.
 */
export function PA_KLBPP(zpar: number, vectors: RunVectors): number {
  const at = vbaInt(zpar);
  const points: number[] = [];
  for (let offset = 2; offset <= 8; offset += 1) {
    const fpb = vectors.fpb(at - offset);
    points.push(fpb === 0 ? 0 : vectors.wage(at - offset) / fpb);
  }

  const averageOfLargest = (k: number) => {
    let total = 0;
    for (let i = 1; i <= k; i += 1) total += wsLarge(points, i);
    return total / k;
  };

  // points[6] is yearpoint7, the earliest of the seven years.
  const yearpointavg = points[6] !== 0 ? averageOfLargest(5) : averageOfLargest(1);

  // Gross pension points, banded.
  let bpp: number;
  if (yearpointavg <= 1) {
    bpp = yearpointavg * 0.96;
  } else if (yearpointavg <= 2.5) {
    bpp = 0.96 + (yearpointavg - 1) * 0.785;
  } else if (yearpointavg <= 3.5) {
    bpp = 0.96 + 1.5 * 0.785 + (yearpointavg - 2.5) * 0.6;
  } else if (yearpointavg <= 7.5) {
    bpp = 0.96 + 1.5 * 0.785 + 0.6 + (yearpointavg - 3.5) * 0.64;
  } else if (yearpointavg <= 20) {
    bpp = 0.96 + 1.5 * 0.785 + 0.6 + 4 * 0.64 + (yearpointavg - 7.5) * 0.65;
  } else {
    bpp = 0.96 + 1.5 * 0.785 + 0.6 + 4 * 0.64 + 12.5 * 0.65 + (yearpointavg - 20) * 0.325;
  }
  return bpp;
}

/** PA-KL's annual pension, which PFA98 replaced from 1998. */
export function PA_KL(
  zpar: number,
  year: number,
  bpp: number,
  wStart: number,
  vectors: RunVectors,
  marginal: 0 | 1 = 0,
): number {
  if (year > 1997) return 0;

  let tidfaktor =
    year < 1997 ? zpar - wStart : 1997 - vectors.year(wStart);
  if (tidfaktor > 30) tidfaktor = 30;

  const base = year < 1997 ? vectors.fpb(vbaInt(zpar)) : vectors.pbb(vbaInt(zpar));
  let pension = (tidfaktor / 30) * bpp * base;

  if (marginal === 0) pension = vbaInt(pension / 12 + 0.5) * 12;
  return pension;
}

/**
 * Adds the defined benefit to whatever the defined contribution part produced.
 *
 * Mirrors `FTJP`. Each agreement computes a benefit and then scales it by the
 * ratio of the divisor at 65 to the divisor at the actual retirement age, which
 * is how drawing early or late is priced.
 *
 * @param tjpAge     the occupational pension so far, from tjpkassa
 * @param startage   the first age the model computes
 * @param stpPoints  ATP points, reused as STP points for SAF-LO
 * @param tp         tilläggspension by age, which PA-KL is offset against
 * @param tmonth     months of occupational pension drawn in the first year
 */
export function FTJP(
  tjpAge: number,
  startage: number,
  age: number,
  tjpPar: number,
  born: number,
  avtal: SchemeId,
  stpPoints: (age: number) => number,
  tp: (age: number) => number,
  vectors: RunVectors,
  tables: DeltalTables,
  context: SchemeContext,
  model: ModelContext,
  tmonth: number,
): number {
  const { wStart, marginal } = context;
  const cohort = vbaInt(born);

  /**
   * The early- or late-withdrawal adjustment, and the correction applied on top
   * of it when the pension is drawn for a fixed term rather than for life.
   */
  const withdrawalFactor = (bornForDeltal = cohort): number => {
    let diverse =
      age === 65
        ? 1
        : deltal(65, bornForDeltal, 65, tables, "premium", 99) /
          deltal(tjpPar, bornForDeltal, age, tables, "premium", 99);

    if (model.tempTjpUttag > 0) {
      const reference =
        born > 1938
          ? deltal(tjpPar, cohort, age, tables, "premium", 99)
          : // The VBA reads a fixed cell here, which holds the income pension
            // divisor at 65 for this cohort -- computed, not stored data.
            tables.incomePension(cohort, 65);
      diverse =
        (diverse * reference) /
        tjp_ddeltal(15, Scheme.SafLo, vectors.year(age), tjpPar, model);
    }
    return diverse;
  };

  let result = tjpAge;

  if (avtal === Scheme.Itp2) {
    let tpYear = 0;
    for (let counter = startage; counter <= 68; counter += 1) {
      if (vectors.wage(counter) > 0) tpYear += 1;
    }
    const diverse = withdrawalFactor();
    const underlag =
      tlITP2F(DC_underlag(0, tjpPar, vectors, marginal), vectors.ibb(age - 1), context, tpYear) *
      diverse;
    result += (underlag * tmonth) / 12;
  }

  if (avtal === Scheme.SafLo) {
    let points = 0;
    let tpYear = 0;
    for (let counter = 28; counter <= 64; counter += 1) {
      if (born < 1968) {
        // Points earned up to 1995, when SAF-LO replaced STP.
        if (vectors.year(counter) < 1996) {
          tpYear += 1;
          points += stpPoints(counter);
        }
      } else if (counter >= 55 && counter <= 59 && stpPoints(counter) > 0) {
        // Later cohorts use the average of ages 55 to 59.
        tpYear += 1;
        points += stpPoints(counter) / 5;
      }
    }

    // STP needs a price base amount that could be determined, i.e. after 1959.
    let base = 0;
    if (vectors.year(65) > 1960) {
      base = tjpPar < 65 ? vectors.fpb(vbaInt(tjpPar)) : vectors.fpb(65);
    }

    const benefit =
      tpYear > 0 && points > 0 ? STP_(points / tpYear + 1, tpYear, base, born, tjpPar, marginal) : 0;
    result += (benefit * withdrawalFactor() * tmonth) / 12;
  }

  if (avtal === Scheme.KapKl || avtal === Scheme.AkapKr) {
    // The seven years ending two years before retirement, capped at 30 income
    // base amounts and restated in the price level of the year before.
    const salaries: number[] = [];
    let tpYear = 0;
    for (let counter = 28; counter <= 70; counter += 1) {
      if (stpPoints(counter) > 0) tpYear += 1;
      if (tjpPar - counter >= 2 && tjpPar - counter < 9) {
        let salary = Math.min(vectors.wage(counter), vectors.ibb(counter) * 30);
        if (vectors.year(counter) >= 1960) {
          salary = (salary * vectors.kpiJune(tjpPar - 1)) / vectors.kpiJune(counter);
        }
        salaries.push(salary);
      }
    }
    while (salaries.length < 7) salaries.push(0);

    // The average of the best five of those seven.
    let underlag = 0;
    for (let i = 1; i <= 5; i += 1) underlag += wsLarge(salaries, i) / 5;
    if (marginal === 0) underlag = vbaRound(underlag, 0);

    const diverse = withdrawalFactor();

    if (vectors.year(age) > 1997) {
      // PFA98, later KAP-KL, took over. Anything earned under PA-KL up to 1997
      // is carried across, uprated 8% for the loss of survivor cover.
      if (vectors.year(28) < 1995) {
        const paklbpp = PA_KLBPP(1997 - cohort, vectors);
        let workyear = 1997 - vectors.year(wStart);
        if (workyear > 30) workyear = 30;
        const yearOffset = age - vectors.year(age);
        const uprate = vectors.ibb(age - 1) / vectors.ibb(yearOffset + 1998);

        let carried: number;
        if (cohort < 1938) {
          // Full ATP, deducted whole.
          carried =
            (PA_KL(1997 - cohort, 1997, paklbpp, wStart, vectors, marginal) - tp(age)) *
            1.08 *
            uprate;
        } else {
          carried =
            (PA_KL(1997 - cohort, 1997, paklbpp, wStart, vectors, marginal) -
              0.6 * stpPoints(yearOffset + 1997) * (workyear / 30) * vectors.pbb(yearOffset + 1997) -
              0.96 * vectors.pbb(yearOffset + 1997) * (workyear / 30)) *
            1.08 *
            uprate;
        }

        // A negative carry-over is passed into KAPKL_f as the salary basis
        // rather than added -- as written.
        underlag =
          carried < 0
            ? KAPKL_f(carried, tpYear, vectors.ibb(tjpPar - 1), born)
            : KAPKL_f(underlag, tpYear, vectors.ibb(tjpPar - 1), born) + carried;
      } else {
        underlag = KAPKL_f(underlag, tpYear, vectors.ibb(tjpPar - 1), born);
      }
    } else {
      const paklbpp = PA_KLBPP(tjpPar, vectors);
      underlag = PA_KL(tjpPar, vectors.year(age), paklbpp, wStart, vectors, marginal) - tp(age);
      // PA-KL guaranteed 100 kr a month to those retiring under it.
      if (underlag < 1200) underlag = 1200;
    }

    result += underlag * diverse;
  }

  if (avtal === Scheme.Pa16Avd2) {
    let tpYear = 0;
    for (let counter = 28; counter <= 64; counter += 1) {
      if (vectors.wage(counter) > 0) tpYear += 1;
    }
    // Note the floor at 1938 here, which the other agreements do not apply.
    const diverse = withdrawalFactor(Math.max(1938, cohort));
    let underlag =
      (tlPA03(DC_underlag(1, tjpPar, vectors, marginal), tpYear, vectors.ibb(age - 1), born, avtal) *
        diverse) /
      12;
    if (marginal === 0) underlag = vbaInt(underlag + 0.49);
    result += underlag * tmonth;
  }

  return result;
}
