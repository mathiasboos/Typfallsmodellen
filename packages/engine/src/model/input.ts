/**
 * The typfall itself -- what the Start sheet asks for.
 *
 * The workbook splits its inputs in two: the Start sheet holds the handful of
 * facts about the person, and Adv_settings holds everything about how the model
 * behaves. That split is what lets normal mode show a short form, so the port
 * keeps it: this is the Start sheet, `ModelContext` is Adv_settings, and a
 * normal-mode run passes only a `TypfallInput`.
 *
 * Defaults come from the workbook, through packages/data/options.json.
 */

import optionsJson from "../../../data/options.json" with { type: "json" };
import { Scheme as SchemeEnum } from "../tjanstepension/types.js";
import type { SchemeId } from "../tjanstepension/types.js";

/**
 * The eight collective agreements, numbered as the Start sheet's drop-down is.
 *
 * Re-exported from the occupational module, which is where the premium
 * functions switch on them, so the input contract and the calculation cannot
 * drift apart.
 */
export { Scheme } from "../tjanstepension/types.js";
export type { SchemeId } from "../tjanstepension/types.js";

/** One year of a hand-entered income vector, replacing the wage profile. */
export interface OwnIncomeYear {
  readonly age: number;
  /** Taxerad förvärvsinkomst -- `Income_(age)`. */
  readonly income: number;
  /** Of which earned income -- `Wage_(age)`. */
  readonly wage: number;
}

/**
 * A typfall: the person the model computes an outcome for.
 *
 * `born` may carry a fraction of a year for the birth month -- the VBA's
 * `YYYY,yy` -- though the shipped workbook only offers whole years.
 */
export interface TypfallInput {
  /** `BornYear`: year of birth. */
  readonly born: number;
  /** `ParYear`: the age at which the public pension is first drawn. */
  readonly retirementAge: number;
  /** `wStartYear`: the age at which working life starts. */
  readonly startWorkAge: number;
  /** `Wage_Monthly`: the salary, per month, in the reference year's money. */
  readonly monthlySalary: number;
  /** `Gift`: cohabiting or married. */
  readonly married: boolean;
  /** `rng_TJP_Val`: which collective agreement applies. */
  readonly scheme: SchemeId;

  /** `rng_Yearly_Inflation`. Prognosstandard: 0. */
  readonly yearlyInflation: number;
  /** `rng_Real_Growth`, general income growth. Prognosstandard: 0. */
  readonly realGrowth: number;
  /** `rng_FondAvkastning`, the return in excess of wage growth. Prognosstandard: 0.017. */
  readonly realReturn: number;

  /**
   * `rng_Egen_Lon` with the Egen inkomst sheet: a hand-entered income vector
   * instead of a wage profile. Present means on, and forces `startage` to 15.
   */
  readonly ownIncome?: readonly OwnIncomeYear[];

  /**
   * The PGB sheet's hand-typed pensionsgrundande belopp for sickness and
   * activity compensation, conscription and study. The shipped workbook has
   * none, so childcare years are the only PGB a default run earns.
   */
  readonly pgbManual?: readonly PgbManualYear[];
}

/** One age's manual entries on the PGB sheet. */
export interface PgbManualYear {
  readonly age: number;
  /** Column 5: sickness and activity compensation. */
  readonly sa: number;
  /** Column 9: conscription. */
  readonly vpl: number;
  /** Column 16: study. */
  readonly studier: number;
}

const NORMAL = optionsJson.normalDefaults as {
  birthYear: number;
  retirementAge: number;
  startWorkAge: number;
  monthlySalary: number;
  married: boolean;
  yearlyInflation: number;
  realGrowth: number;
  realReturn: number;
};

/**
 * The typfall the workbook opens with.
 *
 * The shipped Start sheet offers no occupational pension by default, which is
 * the one field `normalDefaults` does not carry.
 */
export function defaultInput(overrides: Partial<TypfallInput> = {}): TypfallInput {
  const base: TypfallInput = {
    born: NORMAL.birthYear,
    retirementAge: NORMAL.retirementAge,
    startWorkAge: NORMAL.startWorkAge,
    monthlySalary: NORMAL.monthlySalary,
    married: NORMAL.married,
    scheme: SchemeEnum.None,
    yearlyInflation: NORMAL.yearlyInflation,
    realGrowth: NORMAL.realGrowth,
    realReturn: NORMAL.realReturn,
  };
  return Object.freeze({ ...base, ...overrides });
}

/** Birth years the Start sheet's dropdown offers. */
export const BIRTH_YEARS = optionsJson.ranges.birthYears as readonly number[];
/** Retirement ages the Start sheet's dropdown offers. */
export const RETIREMENT_AGES = optionsJson.ranges.retirementAges as readonly number[];
/** The agreement dropdown, value and label, as the workbook words it. */
export const SCHEME_CHOICES = optionsJson.choices.occupationalPension as readonly {
  value: number;
  label: string;
}[];
