/**
 * The case sets `reference/golden/ExportGoldenCases.bas` generates.
 *
 * Two things need them. The report groups divergences by block, which is what
 * turns "32 cases are wrong" into "every KAP-KL case is wrong"; and the
 * harness's own tests need a golden file to run against before a real one
 * exists, which means generating the same inputs here.
 *
 * Keep in step with `BuildQuickCases` and `BuildFullCases` in the macro. The
 * block totals are asserted against the generated cases, so a set that drifts
 * fails rather than mislabels.
 */

import { lowestPensionAge } from "../../src/index.js";
import type { SchemeId } from "../../src/index.js";

/** One generated case, in the ten input columns' order. */
export interface GeneratedCase {
  readonly born: number;
  readonly startWorkAge: number;
  readonly retirementAge: number;
  readonly annualSalary: number;
  readonly inflation: number;
  readonly realGrowth: number;
  readonly realReturn: number;
  readonly scheme: SchemeId;
}

export interface Block {
  readonly name: string;
  readonly count: number;
  readonly what: string;
}

export type CaseSetName = "quick" | "full";

class Builder {
  readonly cases: GeneratedCase[] = [];
  readonly blocks: Block[] = [];
  private marked = 0;

  /**
   * The macro's `AddCase`: clamp the retirement age to the cohort's earliest,
   * and drop a case whose working life would be empty.
   *
   * The clamp reads Nyckeltal column 121, which `lowestPensionAge` now serves
   * from the same table -- so the generated cases match the macro's exactly.
   */
  add(
    born: number,
    startWorkAge: number,
    retirementAge: number,
    annualSalary: number,
    inflation: number,
    realGrowth: number,
    realReturn: number,
    scheme: number,
  ): void {
    const clamped = Math.max(retirementAge, lowestPensionAge(born));
    if (startWorkAge >= clamped) return;
    this.cases.push({
      born,
      startWorkAge,
      retirementAge: clamped,
      annualSalary,
      inflation,
      realGrowth,
      realReturn,
      scheme: scheme as SchemeId,
    });
  }

  block(name: string, what: string): void {
    this.blocks.push({ name, count: this.cases.length - this.marked, what });
    this.marked = this.cases.length;
  }
}

const SALARIES_FULL = [180_000, 264_000, 324_000, 396_000, 462_000, 540_000, 660_000, 840_000, 1_080_000];
const COHORTS_FULL = [
  1937, 1938, 1940, 1945, 1950, 1953, 1954, 1955, 1960,
  1965, 1970, 1975, 1980, 1985, 1990, 1995, 2000, 2005,
];

function buildQuick(): Builder {
  const b = new Builder();

  for (const born of [1937, 1938, 1953, 1954]) {
    for (let scheme = 1; scheme <= 8; scheme += 1) b.add(born, 23, 66, 462_000, 0, 0, 0.017, scheme);
  }
  b.block("A", "the ATP boundaries against every agreement");

  for (const salary of [180_000, 462_000, 660_000, 1_080_000]) {
    for (const age of [63, 66, 70]) b.add(1959, 23, age, salary, 0, 0, 0.017, 4);
  }
  b.block("B", "salary against retirement age, cohort 1959");

  for (const born of [1970, 1990]) {
    for (const start of [20, 30]) {
      for (const salary of [264_000, 840_000]) b.add(born, start, 67, salary, 0, 0, 0.017, 2);
    }
  }
  b.block("C", "entry age against salary, modern cohorts");

  b.add(1959, 23, 66, 462_000, 0.02, 0, 0.017, 4);
  b.add(1959, 23, 66, 462_000, 0, 0.016, 0.017, 4);
  b.add(1959, 23, 66, 462_000, 0, 0, 0.035, 4);
  b.add(1959, 23, 66, 462_000, 0.02, 0.016, 0.035, 4);
  b.block("D", "away from the forecasting standard");

  for (const born of [1945, 1980]) {
    b.add(born, 25, 66, 180_000, 0, 0, 0.017, 1);
    b.add(born, 20, 66, 1_080_000, 0, 0, 0.017, 3);
  }
  b.block("E", "low and high earners");

  b.add(1960, 20, 67, 324_000, 0, 0, 0.017, 5);
  b.block("F", "the user manual's care assistant");

  for (const born of [1957, 1958]) {
    for (const salary of [180_000, 462_000]) b.add(born, 23, 66, salary, 0, 0, 0.017, 4);
  }
  b.block("G", "retirement in 2023 and 2024");

  return b;
}

function buildFull(): Builder {
  const b = new Builder();

  for (const born of COHORTS_FULL) {
    for (let scheme = 1; scheme <= 8; scheme += 1) b.add(born, 23, 66, 462_000, 0, 0, 0.017, scheme);
  }
  b.block("A", "every cohort against every agreement");

  for (const salary of SALARIES_FULL) {
    for (const age of [63, 65, 66, 67, 68, 70, 75]) b.add(1959, 23, age, salary, 0, 0, 0.017, 4);
  }
  b.block("B", "salary against retirement age, cohort 1959");

  for (const start of [18, 20, 23, 25, 30]) {
    for (const salary of SALARIES_FULL) b.add(1970, start, 67, salary, 0, 0, 0.017, 2);
  }
  b.block("C", "entry age against salary, cohort 1970");

  for (let i = 0; i < COHORTS_FULL.length; i += 3) {
    const born = COHORTS_FULL[i]!;
    b.add(born, 23, 66, 462_000, 0.02, 0, 0.017, 4);
    b.add(born, 23, 66, 462_000, 0, 0.016, 0.017, 4);
    b.add(born, 23, 66, 462_000, 0, 0, 0.035, 4);
    b.add(born, 23, 66, 462_000, 0.02, 0.016, 0.035, 4);
  }
  b.block("D", "away from the forecasting standard");

  for (let i = 0; i < COHORTS_FULL.length; i += 2) {
    const born = COHORTS_FULL[i]!;
    b.add(born, 25, 66, 180_000, 0, 0, 0.017, 1);
    b.add(born, 20, 66, 1_080_000, 0, 0, 0.017, 3);
  }
  b.block("E", "low and high earners across cohorts");

  b.add(1960, 20, 67, 324_000, 0, 0, 0.017, 5);
  b.block("F", "the user manual's care assistant");

  for (const born of [1957, 1958]) {
    for (const salary of [180_000, 462_000]) b.add(born, 23, 66, salary, 0, 0, 0.017, 4);
  }
  b.block("G", "retirement in 2023 and 2024");

  return b;
}

const SETS: Record<CaseSetName, Builder> = {
  quick: buildQuick(),
  full: buildFull(),
};

export function caseSet(name: CaseSetName): readonly GeneratedCase[] {
  return SETS[name].cases;
}

export function blocksOf(name: CaseSetName): readonly Block[] {
  return SETS[name].blocks;
}

/**
 * Which case set a golden file came from.
 *
 * The macro writes `# caseset:` when it knows, which it does for a run it
 * generated. A file exported from the sheet after a restart carries no key, so
 * the count decides: an exact total first, then the smallest set that could
 * contain this many cases -- a halted quick run of 40 is still labelled from
 * the quick blocks, correctly for the rows it has.
 */
export function resolveCaseSet(declared: string | undefined, caseCount: number): CaseSetName | null {
  if (declared === "quick" || declared === "full") return declared;

  const names: CaseSetName[] = ["quick", "full"];
  for (const name of names) {
    if (caseSet(name).length === caseCount) return name;
  }
  for (const name of names) {
    if (caseSet(name).length > caseCount) return name;
  }
  return null;
}
