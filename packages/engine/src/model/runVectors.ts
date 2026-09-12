/**
 * The age-indexed series the rule functions read off Mcalc's globals.
 *
 * The VBA keeps these as module-level arrays -- `Wage_()`, `IBB()`, `pbb()` and
 * so on -- that any function can reach. Gathering them behind one interface lets
 * the ported functions stay pure while still being called the same way.
 */
export interface RunVectors {
  /** `Wage_(age)` -- earnings at each age. */
  wage(age: number): number;
  /** `IBB(age)` -- income base amount. */
  ibb(age: number): number;
  /** `pbb(age)` -- price base amount. */
  pbb(age: number): number;
  /** `FPB(age)` -- raised price base amount. */
  fpb(age: number): number;
  /** `KPI_j(age)` -- KPI, June figure. */
  kpiJune(age: number): number;
  /** `year_(age)` -- the income year at each age. */
  year(age: number): number;
}
