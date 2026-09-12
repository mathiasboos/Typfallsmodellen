/**
 * `run(input, context)` -- one typfall, start to finish.
 *
 * This is `Mcalc` as a function: validate and build the vectors, walk the age
 * loop, then assemble the tables. Everything the workbook reads from a named
 * range arrives on the two arguments, and everything it writes to the Start
 * sheet comes back in the result.
 */

import annuityTablesJson from "../../../data/annuity-tables.json" with { type: "json" };
import type { DeathProbabilities } from "../pension/mortality.js";
import { DeltalTables } from "../pension/deltal.js";
import type { AnnuityTablesData } from "../pension/deltal.js";
import { defaultContext } from "./context.js";
import type { ModelContext } from "./context.js";
import { defaultInput } from "./input.js";
import type { TypfallInput } from "./input.js";
import { prepareRun, runLoop } from "./mcalc.js";
import type { Run } from "./mcalc.js";
import { buildResult } from "./result.js";
import type { TypfallResult } from "./result.js";
import { startsetup } from "./setup.js";
import type { Warning } from "./setup.js";
import { createRunState } from "./state.js";

const PUBLISHED = annuityTablesJson as unknown as AnnuityTablesData;

/**
 * The death probabilities the run needs.
 *
 * They live in a 240 KB binary rather than JSON, so how they are loaded depends
 * on where the engine runs: `src/data/nodeLoader.ts` reads the file, and a
 * browser fetches it. The caller supplies them either way.
 */
export interface RunOptions {
  readonly deaths: DeathProbabilities;
  /** Published delningstal, if something other than the committed tables. */
  readonly annuityTables?: AnnuityTablesData;
}

/** The prepared run, for a caller that wants the state as well as the result. */
export interface PreparedRun {
  readonly run: Run;
  readonly warnings: Warning[];
}

/** Builds a run without walking it, which the tests and the harness both use. */
export function prepare(
  input: TypfallInput,
  context: ModelContext,
  options: RunOptions,
): PreparedRun {
  const setup = startsetup(input, context);
  const warnings: Warning[] = [...setup.warnings];
  const state = createRunState(setup.vectors.startage, setup.vectors.slutage);
  const tables = new DeltalTables(options.annuityTables ?? PUBLISHED, options.deaths, context);
  const run = prepareRun(setup, input, context, state, tables, warnings);
  return { run, warnings };
}

/**
 * Computes one typfall.
 *
 * @param input   the person, or `defaultInput()` for the workbook's own typfall
 * @param context the advanced settings, or `defaultContext()` for its normal ones
 * @param options where the mortality data comes from
 */
export function run(
  input: TypfallInput = defaultInput(),
  context: ModelContext = defaultContext(),
  options: RunOptions,
): TypfallResult {
  const prepared = prepare(input, context, options);
  runLoop(prepared.run);
  return buildResult(prepared.run, prepared.warnings);
}
