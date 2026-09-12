/**
 * Turns an exported case into the arguments `run` takes, and checks that the
 * settings the workbook exported are the ones the engine was built from.
 *
 * The input columns are read in `mdlIndataInputOutput.bas` 279-301. Nothing
 * here guesses: each column below cites the line that consumes it.
 */

import optionsJson from "../../../data/options.json" with { type: "json" };
import { SCHEME_CHOICES, defaultContext, defaultInput } from "../../src/index.js";
import type { ModelContext, SchemeId, TypfallInput } from "../../src/index.js";
import type { GoldenCase, GoldenFile } from "./parse.js";

interface AdvancedSettingSpec {
  readonly name: string;
  readonly row: number;
  readonly default: unknown;
  readonly current: unknown;
}

const SPECS = optionsJson.advancedSettings as readonly AdvancedSettingSpec[];

/** A setting in the file that the engine cannot vouch for. */
export interface SettingIssue {
  readonly setting: string;
  /** What the workbook exported. */
  readonly exported: string;
  /** What packages/data/options.json says the model's normal value is. */
  readonly expected: string;
  /** Fatal issues make the comparison meaningless rather than merely suspect. */
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
  /** The model version string, for the report. */
  readonly modelVersion: string;
  readonly exportedAt: string;
  /** Settings that were absent from the file, so the engine's default was assumed. */
  readonly assumed: readonly string[];
  readonly issues: readonly SettingIssue[];
}

/**
 * Settings that decide what the exported numbers *mean*. A disagreement here is
 * not a note in the report -- it makes every comparison after it worthless, so
 * the harness refuses rather than reporting agreement it cannot vouch for.
 *
 * `Alt_p_age` earns its place: at 1 or 2 the batch runner overrides the
 * retirement age in column D from the Nyckeltal sheet
 * (mdlIndataInputOutput.bas:283-288), so the input in the file is not the input
 * that produced the output beside it.
 */
const FATAL_SETTINGS = new Set(
  [
    "rng_Bara_fastapriser",
    "Alt_p_age",
    "Risk",
    "Average_Earning",
    "marginal",
    "w_ref",
    "rng_Sista_PensRatt",
    "rngPens_Inflation",
  ].map((name) => name.toLowerCase()),
);

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
 * Checks every setting the file carries against the model's normal value.
 *
 * HOWTO.md asks for "Anvand normala installningar" before the run, so the two
 * should agree. Where they do not, either the export ran with something other
 * than the shipped model or packages/data/options.json is stale -- both worth
 * knowing before reading a single number.
 */
function checkAdvancedSettings(file: GoldenFile): SettingIssue[] {
  const issues: SettingIssue[] = [];

  for (const [key, exported] of file.settings) {
    if (!key.startsWith("adv.")) continue;
    const name = key.slice("adv.".length);
    const normal = normalValue(name);
    if (normal === undefined) continue; // a setting the extraction does not model
    const actual = Number(exported);
    if (!Number.isFinite(actual)) continue; // free text, nothing to compare
    if (actual === normal.value) continue;

    const fatal = FATAL_SETTINGS.has(name);
    issues.push({
      setting: normal.canonical,
      exported,
      expected: String(normal.value),
      fatal,
      message: fatal
        ? `${normal.canonical} was ${exported} during the export but the model's normal ` +
          `value is ${normal.value}. This changes what the exported numbers mean, so the ` +
          `comparison would be meaningless. Re-run the export after clicking "Anvand ` +
          `normala installningar" on Adv_settings.`
        : `${normal.canonical} was ${exported} during the export, against a normal value ` +
          `of ${normal.value}.`,
    });
  }

  return issues;
}

export function readSettings(file: GoldenFile): GoldenSettings {
  const issues = checkAdvancedSettings(file);
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
    modelVersion: lookup(file, "model") ?? "unknown",
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
 * `defaultContext()` is already built from the workbook's own Adv_settings via
 * packages/data/options.json, so there is nothing to translate -- the file's
 * settings are checked against it in `readSettings` rather than copied over it.
 */
export function toContext(): ModelContext {
  return defaultContext();
}
