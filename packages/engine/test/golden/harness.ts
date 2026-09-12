/**
 * Runs the engine over every case in the golden file.
 */

import { existsSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { loadDeathProbabilities } from "../../src/data/nodeLoader.js";
import { run } from "../../src/index.js";
import { checkLabels, compareCase, failedCase } from "./compare.js";
import type { CaseComparison } from "./compare.js";
import { readSettings, toContext, toInput } from "./map.js";
import type { GoldenSettings } from "./map.js";
import { readGoldenFile } from "./parse.js";
import type { GoldenFile } from "./parse.js";

/** Where the export is committed. */
export const GOLDEN_PATH = fileURLToPath(
  new URL("../../../../reference/golden/golden-cases.csv", import.meta.url),
);

export function goldenFileExists(): boolean {
  return existsSync(GOLDEN_PATH);
}

/**
 * The blocks `BuildCases` generates, in order.
 *
 * Kept because a failure confined to one block says far more than a flat list
 * of case numbers: block A varies the agreement across every cohort, E the
 * extremes of the salary range, F is the user manual's worked example.
 */
export const BLOCKS: readonly { name: string; count: number; what: string }[] = [
  { name: "A", count: 144, what: "every cohort against every agreement" },
  { name: "B", count: 63, what: "salary against retirement age, cohort 1959" },
  { name: "C", count: 45, what: "entry age against salary, cohort 1970" },
  { name: "D", count: 24, what: "away from the forecasting standard" },
  { name: "E", count: 18, what: "low and high earners across cohorts" },
  { name: "F", count: 1, what: "the user manual's care assistant" },
];

/** Which block a case number falls in. */
export function blockOf(index: number): string {
  let start = 1;
  for (const block of BLOCKS) {
    if (index < start + block.count) return block.name;
    start += block.count;
  }
  return "?";
}

export interface ComparisonRun {
  readonly file: GoldenFile;
  readonly settings: GoldenSettings;
  readonly cases: readonly CaseComparison[];
  /** Labels whose text does not match the column their position says they are. */
  readonly labelComplaints: readonly string[];
  readonly elapsedMs: number;
}

export function compareGoldenFile(path: string = GOLDEN_PATH): ComparisonRun {
  const file = readGoldenFile(path);
  const settings = readSettings(file);

  const fatal = settings.issues.filter((issue) => issue.fatal);
  if (fatal.length > 0) {
    throw new Error(
      `the golden file was exported with settings the comparison cannot work ` +
        `around:\n` +
        fatal.map((issue) => `  - ${issue.message}`).join("\n"),
    );
  }

  const deaths = loadDeathProbabilities();
  const context = toContext();
  const started = Date.now();

  const cases = file.cases.map((row): CaseComparison => {
    try {
      const result = run(toInput(row, settings), context, { deaths });
      return compareCase(row, result, settings, file.outputLabels);
    } catch (error) {
      return failedCase(row, error);
    }
  });

  return {
    file,
    settings,
    cases,
    labelComplaints: checkLabels(file.outputLabels, settings.compareTo),
    elapsedMs: Date.now() - started,
  };
}
