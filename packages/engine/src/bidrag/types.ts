/**
 * The Mcalc globals and workbook settings the housing supplement reads.
 *
 * `BTP` and `SBTP` reach for `Application.Range` and for the `Public` variables
 * declared in VBA_go.bas. Gathering those behind one object keeps the ported
 * functions pure, the same way `SchemeContext` already does for the
 * occupational pension.
 */

import type { RunVectors } from "../model/runVectors.js";

/** What `BTP` needs beyond its own arguments. */
export interface BtpContext {
  /**
   * `RNG_dela`: for a cohabiting household, 1 reports the combined amount,
   * 2 the individual's, 3 an equal split. Default 2.
   */
  readonly dela: number;
}

/** What `SBTP` needs beyond its own arguments. */
export interface SbtpContext extends BtpContext {
  /** `RulesfromUtg`: expenditure rules fixed from this year, 0 for current rules. */
  readonly rulesFromUtg: number;
  /** `born`: year of birth. */
  readonly born: number;
  /** `age`: the age Mcalc is currently at. */
  readonly age: number;
  /** `slutage`: the last age the run covers. */
  readonly slutage: number;
  /** `Iyear`: the reference year for switching to earnings indexation. */
  readonly Iyear: number;
  /** `year_()` and the other age-indexed vectors. */
  readonly vectors: RunVectors;
}
