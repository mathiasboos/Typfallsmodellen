/**
 * Death probabilities and the annuity factors derived from them.
 *
 * Port of `Mortality.bas`: `ReadMortality` (loading) and `Calculate_Deltal`
 * (the actuarial core). It produces
 *
 *   delningstal      the divisor turning pension capital into an annual pension,
 *                    separately for income pension, premium pension and
 *                    occupational pension, which discount at different rates
 *   arvsvinstfaktorer the inheritance gains redistributed from those who die
 *   expected lifetime
 *
 * Verified against reference/fixtures/annuity-factors-cached.csv -- the 17 685
 * rows the workbook computed into `mortality!P:Z`.
 *
 * Two quirks below are faithful to the VBA rather than to actuarial intent. Both
 * are marked; a clean reimplementation would "fix" them and diverge.
 */

import { vbaRound } from "../vba/math.js";

/** The share of newborns who are boys, used to weight the unisex table. */
const SHARE_OF_BOYS = 0.5145;

/** Weight on age 0 rather than age 1 when smoothing the first year of life. */
const BETA0 = 0.9;

/** Real return norm for income-pension delningstal: 1.6% forskottsranta. */
const NORM_INCOME_PENSION = 1.016;

/** Ages the annuity factors are produced for. */
export const FIRST_ANNUITY_AGE = 50;
export const LAST_ANNUITY_AGE = 105;
/** Inheritance gains start at 61. */
export const FIRST_INHERITANCE_AGE = 61;

/** SCB's death risks, indexed by year, sex and age. */
export class DeathProbabilities {
  constructor(
    private readonly risks: Float64Array,
    readonly firstYear: number,
    readonly lastYear: number,
    readonly maxAge: number,
  ) {}

  static fromPacked(
    risks: Float64Array,
    meta: { firstYear: number; lastYear: number; maxAge: number },
  ): DeathProbabilities {
    return new DeathProbabilities(risks, meta.firstYear, meta.lastYear, meta.maxAge);
  }

  /**
   * Risk of death at `age` for someone of `sex` in `year`.
   *
   * The year is clamped to the table, matching the VBA's
   * `mini(Maxy, maxi(Miny, year))` at every call site.
   */
  at(year: number, sex: 1 | 2, age: number): number {
    const clampedYear = Math.min(this.lastYear, Math.max(this.firstYear, year));
    const ages = this.maxAge + 1;
    const index = ((clampedYear - this.firstYear) * 2 + (sex - 1)) * ages + age;
    return this.risks[index] ?? 0;
  }
}

/** Annuity factors for one cohort, indexed `[sex 0..2][age]`. */
export interface AnnuityFactors {
  readonly cohort: number;
  /** Income pension delningstal (N/Q), ages 50..105. */
  readonly dtalip: readonly number[][];
  /** Premium pension delningstal, ages 50..105. */
  readonly dtalpp: readonly number[][];
  /** Occupational pension delningstal -- premium pension net of 15% dividend tax. */
  readonly dtalpp2: readonly number[][];
  /** Expected remaining lifetime on the income-pension table. */
  readonly expLife: readonly number[][];
  /** Expected remaining lifetime on the premium-pension table. */
  readonly expLife2: readonly number[][];
  /** Inheritance-gain factor, ages 61..105. */
  readonly arvf: readonly number[][];
  /** Inheritance-gain factor on the premium-pension table, ages 61..105. */
  readonly arvf2: readonly number[][];
}

/** Riktålder in force in a given calendar year, as Calculate_Deltal steps it. */
function riktageFor(year: number, current: number): number {
  let riktage = current;
  if (year >= 2023) riktage = 66;
  if (year >= 2026) riktage = 67;
  if (year >= 2034) riktage = 68;
  if (year >= 2049) riktage = 69;
  return riktage;
}

/**
 * Premium pension's forskottsranta, which has been cut repeatedly.
 *
 * QUIRK: the VBA evaluates this against a `year` left over from the preceding
 * age loop, which holds `cohort + 106` rather than the year being discounted.
 * For every cohort the model runs (1930-2050) that lands past 2017, so the rate
 * is 1.0165 throughout -- but the port keeps the leftover so that a future
 * cohort range cannot silently change the answer here without changing it in
 * the original too.
 */
function forskottsranta(year: number): number {
  if (year < 2003) return 1.037;
  if (year < 2008) return 1.027;
  if (year < 2015) return 1.039;
  if (year < 2018) return 1.029;
  return 1.0165;
}

function grid(rows: number, cols: number): number[][] {
  return Array.from({ length: rows }, () => new Array<number>(cols).fill(0));
}

/**
 * Computes the annuity factors for one birth cohort.
 *
 * Mirrors `Calculate_Deltal` in Mortality.bas.
 */
export function calculateDeltal(cohort: number, deaths: DeathProbabilities): AnnuityFactors {
  const MAX = 106;

  // Survivors out of 100 000, on two different mortality bases:
  //   b  income pension -- a five-year average of the tables around riktålder
  //   b2 premium pension -- the cohort's own year-by-year table
  const b = grid(MAX + 1, 3);
  const b2 = grid(MAX + 1, 3);

  // QUIRK: riktage is initialised once, outside the sex loop, and the loop
  // leaves it at its final value. So for sex = 2 the early ages inherit
  // whatever sex = 1 ended on, which shifts which mortality year they read.
  // This is faithful to the VBA; the cached fixture confirms it.
  let riktage = 65;
  let year = cohort;

  for (let sex: 1 | 2 = 1; sex <= 2; sex = (sex + 1) as 1 | 2) {
    let pop = 100_000;
    let pop2 = 100_000;

    for (let age = 0; age <= MAX; age += 1) {
      year = cohort + age;
      riktage = riktageFor(year, riktage);

      let qLag = 0;
      let qLag2 = 0;
      if (age > 0) {
        const lastAge = Math.min(100, age - 1);
        // Income pension averages the five tables ending three years before
        // riktålder for the cohort.
        qLag = deaths.at(cohort + riktage - 3, sex, lastAge);
        for (let jj = 1; jj <= 4; jj += 1) {
          qLag += deaths.at(cohort + riktage - 2 - jj, sex, lastAge);
        }
        qLag *= 1 / 5;
        qLag2 = deaths.at(cohort + age - 2, sex, lastAge);
      }

      pop *= 1 - qLag;
      pop2 *= 1 - qLag2;
      b[age]![sex] = pop;
      b2[age]![sex] = pop2;
    }
  }

  // Smooth the first year of life, then weight into the unisex table. Note the
  // one-year shift: lx_ at index `age` holds lx at `age`, read from b[age].
  const lx = grid(MAX + 1, 3);
  const lx2 = grid(MAX + 1, 3);
  for (let sex = 1; sex <= 2; sex += 1) {
    for (let age = 0; age <= MAX; age += 1) {
      if (age === 0) {
        lx[0]![sex] = BETA0 * b[0]![sex]! + (1 - BETA0) * b[1]![sex]!;
        lx2[0]![sex] = BETA0 * b2[0]![sex]! + (1 - BETA0) * b2[1]![sex]!;
      } else {
        lx[age]![sex] = b[age]![sex]!;
        lx2[age]![sex] = b2[age]![sex]!;
      }
    }
  }

  // lxBySex[sex][age]; sex 0 is unisex. Male and female are *weighted* by the
  // sex ratio, not raw survivor counts -- so that the three sum consistently.
  const lxBySex = grid(3, MAX + 2);
  const lx2BySex = grid(3, MAX + 2);
  for (let age = 1; age <= MAX; age += 1) {
    lxBySex[0]![age] = lx[age]![1]! * SHARE_OF_BOYS + lx[age]![2]! * (1 - SHARE_OF_BOYS);
    lx2BySex[0]![age] = lx2[age]![1]! * SHARE_OF_BOYS + lx2[age]![2]! * (1 - SHARE_OF_BOYS);
    lxBySex[1]![age] = lx[age]![1]! * SHARE_OF_BOYS;
    lx2BySex[1]![age] = lx2[age]![1]! * SHARE_OF_BOYS;
    lxBySex[2]![age] = lx[age]![2]! * (1 - SHARE_OF_BOYS);
    lx2BySex[2]![age] = lx2[age]![2]! * (1 - SHARE_OF_BOYS);
  }

  const ageCount = LAST_ANNUITY_AGE + 1;
  const dtalip = grid(3, ageCount);
  const dtalpp = grid(3, ageCount);
  const dtalpp2 = grid(3, ageCount);
  const expLife = grid(3, ageCount);
  const expLife2 = grid(3, ageCount);

  // `year` here is the leftover from the age loop -- see forskottsranta().
  const normpp = forskottsranta(year);
  const normpp2 = 1 + (normpp - 1) * (1 - 0.15);

  for (let n = FIRST_ANNUITY_AGE; n <= LAST_ANNUITY_AGE; n += 1) {
    for (let sex = 0; sex <= 2; sex += 1) {
      let d = 0;
      let e = 0;
      let e2 = 0;
      let r = 0;
      let r2 = 0;

      // Sum discounted survivors over remaining years, interpolating linearly
      // within each year across its twelve monthly payments.
      for (let k = n; k <= LAST_ANNUITY_AGE; k += 1) {
        const lxK = lxBySex[sex]![k]!;
        const lxNext = lxBySex[sex]![k + 1]!;
        const lx2K = lx2BySex[sex]![k]!;
        const lx2Next = lx2BySex[sex]![k + 1]!;

        for (let x = 0; x <= 11; x += 1) {
          const survivors = lxK + (lxNext - lxK) * (x / 12);
          const survivors2 = lx2K + (lx2Next - lx2K) * (x / 12);

          d += NORM_INCOME_PENSION ** -(k - n) * survivors * NORM_INCOME_PENSION ** (-x / 12);
          e += normpp ** -(k - n) * survivors2 * normpp ** (-x / 12);
          e2 += normpp2 ** -(k - n) * survivors2 * normpp2 ** (-x / 12);
          r += survivors;
          r2 += survivors2;
        }
      }

      dtalip[sex]![n] = vbaRound(d / (12 * lxBySex[sex]![n]!), 2);
      dtalpp[sex]![n] = vbaRound(e / (12 * lx2BySex[sex]![n]!), 2);
      dtalpp2[sex]![n] = vbaRound(e2 / (12 * lx2BySex[sex]![n]!), 2);
      expLife[sex]![n] = vbaRound(r / (12 * lxBySex[sex]![n]!), 2);
      // Note expLife2 divides by the *income pension* table, as the VBA does.
      expLife2[sex]![n] = vbaRound(r2 / (12 * lxBySex[sex]![n]!), 2);
    }
  }

  const arvf = grid(3, ageCount);
  const arvf2 = grid(3, ageCount);
  for (let sex = 0; sex <= 2; sex += 1) {
    for (let age = FIRST_INHERITANCE_AGE; age <= LAST_ANNUITY_AGE; age += 1) {
      arvf[sex]![age] = vbaRound(lxBySex[sex]![age]! / lxBySex[sex]![age + 1]!, 6);
      arvf2[sex]![age] = vbaRound(lx2BySex[sex]![age]! / lx2BySex[sex]![age + 1]!, 6);
    }
  }

  return { cohort, dtalip, dtalpp, dtalpp2, expLife, expLife2, arvf, arvf2 };
}
