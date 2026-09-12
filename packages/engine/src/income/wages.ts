/**
 * The wage vector: what the typfall earns at each age.
 *
 * Port of `wages` in Lön_mm.bas. The entered salary is a single figure for one
 * reference year; this spreads it across a working life by following the income
 * index, optionally bends it with an age profile, and scales the entry and
 * retirement years by the months actually worked.
 *
 * DEVIATION: the VBA computes the pension withdrawal share into the global
 * `uttagIP` as a side effect of asking for a wage. Here that is
 * `withdrawalShare`, an exported function the caller applies -- same arithmetic,
 * but a wage function that only returns a wage.
 */

import type { EconomicData } from "../data/projection.js";
import { vbaInt } from "../vba/math.js";

/** The run-level facts about a typfall that `wages` reads from globals. */
export interface WageProfile {
  /** Birth year, possibly with a fraction of a year for the birth month. */
  readonly born: number;
  /** Age at which public pension is first drawn. */
  readonly par: number;
  /** Age at which withdrawal becomes final and full. */
  readonly defAr: number;
  /**
   * `Wage_profil`: 0 a straight wage path, 1-4 an age profile fitted to
   * observed earnings. See the coefficients below.
   */
  readonly wageProfile: number;
  /**
   * `rngLönPartUttag`: a number means work is fixed at that share while pension
   * is drawn partially; anything else means work follows the withdrawal.
   */
  readonly workDuringPartialWithdrawal: number | string;
  /** `UttagIP`: the share of pension drawn during partial withdrawal. */
  readonly partialWithdrawalShare: number;
}

/**
 * Fifth-degree age profiles for earnings, from Pensionsmyndigheten's 2017
 * memorandum on the regleringsbrev. Profile 1 is flattened above 61.
 */
const WAGE_PROFILES: Readonly<Record<number, readonly [number, number, number, number, number, number]>> = {
  1: [-33.01341746, 4.354364353, -0.212974964, 0.005053746, -0.0000583887, 0.000000263128],
  2: [-20.6648551, 2.496893883, -0.10843247, 0.002287573, -0.0000236622, 0.0000000962164],
  3: [-23.73767932, 3.104093675, -0.149408507, 0.003477166, -0.0000392148, 0.000000171834],
  4: [-17.18584424, 2.203777173, -0.10280891, 0.00234388, -0.0000261194, 0.000000113668],
};

function polynomial(
  coefficients: readonly [number, number, number, number, number, number],
  age: number,
): number {
  const [b0, b1, b2, b3, b4, b5] = coefficients;
  return b0 + b1 * age + b2 * age ** 2 + b3 * age ** 3 + b4 * age ** 4 + b5 * age ** 5;
}

/**
 * The share of pension drawn at a given age.
 *
 * Zero before pension is first taken, the partial share between first and final
 * withdrawal, and one from the final withdrawal on. `previous` carries the last
 * value forward, which is what the VBA's global does in the branch where it
 * assigns nothing.
 *
 * Note the year in which pension is both first and finally drawn (par === defAr,
 * the ordinary case) falls through every branch and keeps the previous value --
 * so the share is still 0 in that year and only becomes 1 the year after. It
 * makes no difference to the wage: the months worked in that year are computed
 * by a formula whose withdrawal term is multiplied by (defAr - par), i.e. zero.
 */
export function withdrawalShare(age: number, profile: WageProfile, previous = 0): number {
  const { par, defAr, workDuringPartialWithdrawal, partialWithdrawalShare } = profile;

  if (age < vbaInt(par)) return 0;

  if (age >= vbaInt(par) && par <= defAr && age < defAr) {
    // DEAD in the original too, and kept for that reason: entering this branch
    // needs age >= par and age < defAr at once, which par === defAr forbids.
    if (par === defAr) return 1;
    // A number here fixes work at a share independent of the withdrawal.
    return typeof workDuringPartialWithdrawal === "number"
      ? 1 - workDuringPartialWithdrawal
      : partialWithdrawalShare;
  }

  return age > defAr ? 1 : previous;
}

/**
 * Earnings at `age`, in the year `year`.
 *
 * @param wstart      age of entry to the labour market
 * @param wage        the salary entered, for the reference income year
 * @param wTime       the age the entered salary refers to
 * @param refYear     the year whose price level results are expressed in
 * @param nominalWage true for current prices, false for fixed prices
 * @param uttagIp     the withdrawal share at this age, from `withdrawalShare`
 */
export function wages(
  age: number,
  year: number,
  wstart: number,
  wage: number,
  profile: WageProfile,
  data: EconomicData,
  uttagIp: number,
  wTime = 0,
  refYear = 2011,
  nominalWage = true,
): number {
  if (age < wstart) return 0;

  const { born, par, defAr, wageProfile } = profile;
  let amount = wage;

  // The income year the entered salary refers to.
  const wSlut = vbaInt(born + wTime);

  // A birthday part-way through a year pushes the relevant income year on by
  // one; these three flags carry that through each comparison below.
  const konst = vbaInt(born + par + 1 / 1000) > vbaInt(born) + vbaInt(par) ? 1 : 0;
  const konst2 = vbaInt(born + defAr + 1 / 1000) > vbaInt(born) + vbaInt(defAr) ? 1 : 0;
  const konst3 = vbaInt(born + wstart + 1 / 1000) > vbaInt(born) + vbaInt(wstart) ? 1 : 0;

  if (age < vbaInt(wstart + konst3)) return 0;

  if (age <= vbaInt(defAr + konst) && age >= vbaInt(wstart)) {
    let indexRatio: number;
    if (year < 1958) {
      // No income index before 1957; the model assumes 8% nominal growth back.
      indexRatio = data.inkomstindex.at(1958) / data.inkomstindex.at(wSlut);
      indexRatio /= 1.08 ** (1958 - year);
    } else {
      indexRatio = data.inkomstindex.at(year) / data.inkomstindex.at(wSlut);
    }
    amount *= indexRatio;

    const coefficients = WAGE_PROFILES[wageProfile];
    if (wageProfile > 0 && coefficients !== undefined) {
      // Profile 1 is held flat from 61; the others run on.
      const atAge = wageProfile === 1 ? polynomial(coefficients, Math.min(age, 61)) : polynomial(coefficients, age);
      const atReference =
        wageProfile === 1
          ? polynomial(coefficients, Math.min(wTime, 61))
          : polynomial(coefficients, wTime);
      let shape = atAge / atReference;
      if (shape < 0) shape = 0;
      amount *= shape;
    }
  }

  if (!nominalWage) {
    // The series is in current prices, so this ratio is what puts the result in
    // the reference year's fixed prices.
    amount *= data.kpiAnnual.at(wSlut) / data.kpiAnnual.at(refYear);
  }

  // Months actually worked in this year.
  let month: number;
  if (age === vbaInt(wstart + konst3)) {
    const intoYear = 12 * (born + wstart - vbaInt(born + wstart + 1 / 1000));
    month = vbaInt(12 - intoYear);
  } else if (age < vbaInt(par + konst)) {
    month = 12;
  } else if (age === vbaInt(par + konst)) {
    if (age === vbaInt(defAr + konst2)) {
      month =
        vbaInt(12 * (born + par - vbaInt(born + par + 1 / 1000))) +
        vbaInt(12 * (defAr - par) * (1 - uttagIp));
    } else {
      const whole = vbaInt(12 * (born + par - vbaInt(born + par + 1 / 1000)));
      month = whole + (12 - whole) * (1 - uttagIp);
    }
  } else if (age < vbaInt(defAr + konst2)) {
    month = 12 * (1 - uttagIp);
  } else if (age === vbaInt(defAr + konst2)) {
    month = 12 * (born + defAr - vbaInt(born + defAr + 1 / 1000)) * (1 - uttagIp);
  } else {
    month = 0;
  }

  return (month * amount) / 12;
}
