/**
 * Turns an exported case into the arguments `run` takes, and checks that the
 * settings the workbook exported are the ones the engine was built from.
 *
 * The input columns are read in `mdlIndataInputOutput.bas` 279-301. Nothing
 * here guesses: each column below cites the line that consumes it.
 */

import optionsJson from "../../../data/options.json" with { type: "json" };
import { SCHEME_CHOICES, contextFromSettings, defaultInput } from "../../src/index.js";
import type { ModelContext, SchemeId, TypfallInput } from "../../src/index.js";
import type { GoldenCase, GoldenFile } from "./parse.js";

interface AdvancedSettingSpec {
  readonly name: string;
  readonly row: number;
  readonly default: unknown;
  readonly current: unknown;
}

const SPECS = optionsJson.advancedSettings as readonly AdvancedSettingSpec[];

/** A setting the run used that is not the model's normal value. */
export interface SettingIssue {
  readonly setting: string;
  /** What the run used. */
  readonly exported: string;
  /** What packages/data/options.json says the model's normal value is. */
  readonly expected: string;
  /** True when the file cannot be compared at all. See `UNUSABLE_SETTINGS`. */
  readonly fatal: boolean;
  readonly message: string;
}

/** What the harness has to read out of the file to interpret the numbers. */
export interface GoldenSettings {
  /** `Rng_belopp12`: 1 means every output was divided by twelve. */
  readonly monthly: boolean;
  /** `Rng_CompareTo`: 0 slutlon, 1 netto, 2 disponibel -- which one column M holds. */
  readonly compareTo: 0 | 1 | 2;
  /** `Gift`, from the Start sheet. Applies to every case. */
  readonly married: boolean;
  /** `# caseset`: which set the macro generated, when it said. */
  readonly caseSet: string | undefined;
  /** Every Adv_settings value the run recorded, lower-cased, for the context. */
  readonly advanced: ReadonlyMap<string, number>;
  /** The model version string, for the report. */
  readonly modelVersion: string;
  /**
   * `# publicavg`: the macro's own verdict on whether the exporting workbook is
   * the build this port was written against.
   *
   * Undefined for a file exported before the check existed, which is not the
   * same as a failure -- see `buildCheck` in compare.ts. It is recorded because
   * the first golden file came from a different build of the model, and a
   * version number did not distinguish them: both said "Version 4.8".
   */
  readonly buildCheck: string | undefined;
  readonly exportedAt: string;
  /** Settings that were absent from the file, so the engine's default was assumed. */
  readonly assumed: readonly string[];
  readonly issues: readonly SettingIssue[];
}

/**
 * Settings that make the file itself unusable, rather than merely unusual.
 *
 * Everything else on Adv_settings is a *configuration*, and a golden file
 * records the whole sheet -- so the engine is configured to match it rather
 * than refusing because a setting was off its normal value. These two cannot be
 * matched, because they break the relationship between the file's inputs and
 * its outputs:
 *
 * - `Alt_p_age` at 1 or 2 makes the batch runner take the retirement age from
 *   the Nyckeltal sheet instead of column D
 *   (mdlIndataInputOutput.bas:283-288), so the input in the file is not the
 *   input that produced the output beside it.
 * - `Risk` above 0 draws the fund return at random, so no second run of the
 *   same inputs -- by the workbook or by the engine -- can reproduce it.
 */
const UNUSABLE_SETTINGS = new Set(["Alt_p_age", "Risk"].map((name) => name.toLowerCase()));

/** Reads a setting under any of the keys it may have been written with. */
function lookup(file: GoldenFile, ...keys: string[]): string | undefined {
  for (const key of keys) {
    const value = file.settings.get(key.toLowerCase());
    if (value !== undefined && value !== "") return value;
  }
  return undefined;
}

/** The workbook writes booleans as 1/0; older exports wrote -1 for TRUE. */
function asBoolean(value: string | undefined, fallback: boolean): boolean {
  if (value === undefined) return fallback;
  const n = Number(value);
  if (Number.isFinite(n)) return n !== 0;
  return /^(true|sant|yes|ja)$/i.test(value.trim());
}

/**
 * The normal value of a setting, as extracted from Adv_settings column 8, and
 * the spelling the sheet gives it -- the provenance keys are lower-cased for
 * lookup, and `Rng_belopp12` reads better in a report than `rng_belopp12`.
 */
function normalValue(name: string): { canonical: string; value: number } | undefined {
  const spec = SPECS.find((s) => s.name.toLowerCase() === name.toLowerCase());
  if (spec === undefined) return undefined;
  const value = spec.default;
  if (typeof value === "number") return { canonical: spec.name, value };
  if (typeof value === "boolean") return { canonical: spec.name, value: value ? 1 : 0 };
  return undefined;
}

/**
 * Settings whose Adv_settings row documents a name and a normal value but holds
 * no value cell of its own -- the live setting is a named range somewhere else
 * on the sheet. `rng_Bara_fastapriser` is the clearest: its row says the name
 * and the normal value 1, while the range itself is at A47.
 *
 * An export that reads the row rather than the name finds a blank cell there,
 * and VBA's `IsNumeric(Empty)` is True, so the blank is written out as 0. That
 * is not a setting of 0; it is no reading at all. Taking it at face value put
 * the whole first comparison on the wrong price basis.
 */
const NO_VALUE_CELL = new Set(
  SPECS.filter((spec) => spec.current === null || spec.current === undefined).map((spec) =>
    spec.name.toLowerCase(),
  ),
);

/**
 * Every Adv_settings value the run recorded, keyed by lower-cased name.
 *
 * This is what the engine's context is built from, so it is configured exactly
 * as the workbook was. Free-text rows are skipped; there is nothing to compare
 * a label against. So are the rows above, which carry no reading to trust.
 */
function advancedSettings(file: GoldenFile): Map<string, number> {
  const values = new Map<string, number>();
  for (const [key, raw] of file.settings) {
    if (!key.startsWith("adv.")) continue;
    const name = key.slice("adv.".length);
    const value = Number(raw);
    if (!Number.isFinite(value)) continue;
    // A row with no value cell reads as blank, and VBA writes a blank as 0 --
    // so 0 from one of those rows is "no reading" and has to be ignored. A
    // non-zero one cannot come from a blank cell, so it is a real value that
    // somebody typed into the row, and ignoring it is how 2 800 000 kr of
    // hand-entered pension balance passed for an ordinary run.
    if (NO_VALUE_CELL.has(name) && value === 0) continue;
    values.set(name, value);
  }
  // `# live.<name>` is the same setting read through its defined name rather
  // than off its Adv_settings row, which is the only way to see the eight whose
  // row holds no value. It is a real reading where `adv.` was a blank cell, so
  // it wins -- including for the settings skipped just above.
  for (const [key, raw] of file.settings) {
    if (!key.startsWith("live.")) continue;
    const value = Number(raw);
    if (Number.isFinite(value)) values.set(key.slice("live.".length), value);
  }
  return values;
}

/**
 * Opening balances that replace the earning phase.
 *
 * Set any of these and the model starts from a hand-entered pension behallning
 * instead of computing a working life: every final salary comes out 0 and every
 * pension comes from the balance rather than the salary. That is a legitimate
 * thing to ask the model and useless as reference data, so a file that records
 * one cannot be compared against.
 *
 * They are checked separately from the settings below because they are exactly
 * the ones the old export could not see: their Adv_settings rows hold no value,
 * so the dump wrote 0 and half an hour of exporting read as a normal run.
 */
const OPENING_BALANCES = ["rng_PBHYear", "rng_PBH_IP", "rng_PBH_PP", "rng_PBH_tjp", "rng_PBH_IPS"];

/**
 * Notes where the run's settings are not the model's normal ones.
 *
 * These are not failures. HOWTO.md asks for "Anvand normala installningar"
 * before a run, but a file that records something else is still perfectly
 * comparable -- the engine is built from the same settings. What matters is
 * saying so, because it changes what the comparison proves: it checks the
 * engine against the model *under those settings*, which may not be the ones a
 * website would use.
 *
 * The exception is `UNUSABLE_SETTINGS`, which no configuration can rescue.
 */
function checkAdvancedSettings(file: GoldenFile, used: ReadonlyMap<string, number>): SettingIssue[] {
  const issues: SettingIssue[] = [];

  for (const name of OPENING_BALANCES) {
    const actual = used.get(name.toLowerCase());
    if (actual === undefined || actual === 0) continue;
    issues.push({
      setting: name,
      exported: String(actual),
      expected: "0",
      fatal: true,
      message:
        `${name} was ${actual} during the export. The model starts from that balance ` +
        `instead of computing a working life, so every Slutlon in the file is 0 and every ` +
        `pension comes from the balance rather than the salary. Click "Anvand normala ` +
        `installningar" on Adv_settings and re-run the export.`,
    });
  }

  for (const [name, actual] of used) {
    const normal = normalValue(name);
    const fatal = UNUSABLE_SETTINGS.has(name);
    if (normal === undefined) continue; // a setting the extraction does not model
    // The export formats to ten decimals, so an extracted 0.010000000000000009
    // comes back as 0.01. That is the same number, not a changed setting.
    const same = Math.abs(actual - normal.value) <= Math.abs(normal.value) * 1e-9;
    if (same && !fatal) continue;
    if (fatal && actual === 0) continue; // 0 is the usable value for both of these

    issues.push({
      setting: normal.canonical,
      exported: String(actual),
      expected: String(normal.value),
      fatal,
      message: fatal
        ? `${normal.canonical} was ${actual} during the export. ` +
          (name === "risk"
            ? "That draws the fund return at random, so nothing can reproduce this run."
            : "That makes the batch runner take the retirement age from the Nyckeltal " +
              "sheet rather than from the input column, so the inputs in the file are not " +
              "the ones that produced the outputs beside them.") +
          " Re-run the export with it at 0."
        : `${normal.canonical} was ${actual}, against a normal value of ${normal.value}. ` +
          `The engine was configured to match.`,
    });
  }

  return issues;
}

export function readSettings(file: GoldenFile): GoldenSettings {
  const advanced = advancedSettings(file);
  const issues = checkAdvancedSettings(file, advanced);
  const assumed: string[] = [];

  const belopp12 = lookup(file, "adv.rng_belopp12", "belopp12");
  const compareTo = lookup(file, "adv.rng_compareto", "compareTo");
  const gift = lookup(file, "start.gift");

  if (belopp12 === undefined) assumed.push("Rng_belopp12 (assumed annual amounts)");
  if (compareTo === undefined) assumed.push("Rng_CompareTo (assumed column M is slutlon)");
  if (gift === undefined) assumed.push("Gift (assumed the engine's default)");

  const compareToValue = Number(compareTo ?? 0);
  if (![0, 1, 2].includes(compareToValue)) {
    throw new Error(`Rng_CompareTo in the golden file is ${compareTo}, expected 0, 1 or 2`);
  }

  return {
    monthly: asBoolean(belopp12, false),
    compareTo: compareToValue as 0 | 1 | 2,
    married: asBoolean(gift, defaultInput().married),
    caseSet: lookup(file, "caseset"),
    advanced,
    modelVersion: lookup(file, "model") ?? "unknown",
    buildCheck: lookup(file, "publicavg"),
    exportedAt: lookup(file, "exported") ?? "unknown",
    assumed,
    issues,
  };
}

/**
 * The typfall one exported row describes.
 *
 * Columns, as `InputXGetY` reads them:
 *   1 B  fodelsear          -> BornYear          (:279)
 *   2 C  borjar arbeta      -> wStartYear        (:280)
 *   3 D  pensionsalder      -> PARYear           (:282, unless Alt_p_age)
 *   4 E  arslon             -> Wage_Monthly, /12 (:291)
 *   5 F  inflation          -> rng_Yearly_Inflation (:293)
 *   6 G  real tillvaxt      -> rng_Real_Growth   (:294)
 *   7 H  fondavkastning     -> rng_FondAvkastning (:295)
 *   8 I  egen lonelista     -> rng_Egen_Lon      (:296-299)
 *   9 J  IPS                -> IPS               (:300)
 *  10 K  tjanstepension     -> rng_TJP_Listbox   (:301)
 */
export function toInput(row: GoldenCase, settings: GoldenSettings): TypfallInput {
  const [born, startWorkAge, retirementAge, annualSalary, inflation, growth, fund, ownList, ips, scheme] =
    row.inputs as [number, number, number, number, number, number, number, number, number, number];

  // The generated cases use neither, and both would need data the file does not
  // carry -- an Egen inkomst column, an IPS start year. Refuse rather than
  // compare against a case the engine was not given the same inputs for.
  if (ownList > 1) {
    throw new Error(
      `case ${row.index} (line ${row.line}) uses the workbook's own wage list ` +
        `(column I = ${ownList}), which the golden file does not carry`,
    );
  }
  if (ips !== 0) {
    throw new Error(
      `case ${row.index} (line ${row.line}) has private saving of ${ips}, which the ` +
        `harness does not map`,
    );
  }
  if (!SCHEME_CHOICES.some((choice) => choice.value === scheme)) {
    throw new Error(
      `case ${row.index} (line ${row.line}) names occupational scheme ${scheme}, ` +
        `which is not one of the eight the Start sheet offers`,
    );
  }

  return defaultInput({
    born,
    startWorkAge,
    retirementAge,
    monthlySalary: annualSalary / 12,
    yearlyInflation: inflation,
    realGrowth: growth,
    realReturn: fund,
    married: settings.married,
    scheme: scheme as SchemeId,
  });
}

/**
 * The settings every case ran under.
 *
 * Built from the Adv_settings block the file carries, falling back to the
 * workbook's normal value for anything it does not record. The mapping from
 * sheet name to context field is `defaultContext`'s own -- see
 * `contextFromSettings` in src/model/context.ts -- so there is nothing to keep
 * in step here.
 */
export function toContext(settings: GoldenSettings): ModelContext {
  return contextFromSettings(settings.advanced);
}
