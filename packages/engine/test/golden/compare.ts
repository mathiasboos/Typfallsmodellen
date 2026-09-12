/**
 * Diffs one engine run against one exported case.
 *
 * Every exported value is read from **column D of Table 1** on the Start sheet
 * -- `Col_top_tabell1 + 1`, and `Col_top_tabell1` is column C (VBA_go.bas:801)
 * -- so the comparison is against `Table1Row.adjusted`, never `.nominal`. When
 * `Rng_belopp12` is 1 the macro divides by twelve on the way out.
 */

import { Table1Key } from "../../src/index.js";
import type { Table1Row, TypfallResult } from "../../src/index.js";
import type { GoldenCase } from "./parse.js";
import type { GoldenSettings } from "./map.js";

/**
 * What each output column holds.
 *
 * The order is the contract, not the header labels: `InputXGetY` picks the
 * Table 1 row by a `Select Case` on the column number
 * (mdlIndataInputOutput.bas:328-356), and the labels are whatever text happens
 * to sit in row 7 of the Mikrosim sheet. The labels are cross-checked below
 * rather than trusted.
 *
 * `offset` is the row offset the VBA uses into the Table 1 block, kept so the
 * mapping can be checked against the source by eye.
 */
interface ColumnSpec {
  readonly offset: number;
  readonly vba: string;
  /** The Table 1 row to compare against, or null where the port has no such row. */
  readonly key: Table1Key | null;
  /** Words the sheet's own label is expected to contain, accent-insensitive. */
  readonly expectLabel: readonly string[];
  /** Why the port cannot produce this column yet. */
  readonly notPortedBecause?: string;
}

const NOT_PORTED =
  "Table 1's Efterskatt, Bidrag and Disp_efterskatt rows come from the second tax " +
  "and benefits pass at the retirement age (VBA_go.bas 2607-2800), which the port " +
  "does not have yet";

/** Column 1 follows `Rng_CompareTo` (mdlIndataInputOutput.bas:331-336). */
const FIRST_COLUMN: Record<0 | 1 | 2, ColumnSpec> = {
  0: { offset: -3, vba: "ys(1), lon t-1", key: Table1Key.FinalSalary, expectLabel: ["slutlon", "lon"] },
  1: { offset: -2, vba: "ys(2), netto t-1", key: Table1Key.SalaryAfterTax, expectLabel: ["netto"] },
  2: {
    offset: -1,
    vba: "ys(3), disp t-1",
    key: Table1Key.DisposableBeforeRetirement,
    expectLabel: ["disp"],
  },
};

const REST: readonly ColumnSpec[] = [
  {
    offset: 9,
    vba: "brutto(Int(PAR))",
    key: Table1Key.TotalGross,
    expectLabel: ["total", "brutto"],
  },
  {
    offset: 1,
    vba: "ip(PAR)",
    key: Table1Key.IncomePension,
    expectLabel: ["inkomstp", "income"],
  },
  {
    offset: 2,
    vba: "tp(PAR)",
    key: Table1Key.SupplementaryPension,
    expectLabel: ["tillaggspension", "atp", "(a)tp", "supplement"],
  },
  { offset: 3, vba: "pp(PAR)", key: Table1Key.PremiumPension, expectLabel: ["premie", "premium"] },
  {
    offset: 4,
    vba: "garp(PAR)",
    key: Table1Key.GuaranteePension,
    expectLabel: ["garanti", "garp", "guarantee"],
  },
  {
    offset: 5,
    vba: "ptillagg(PAR)",
    key: Table1Key.IncomePensionSupplement,
    expectLabel: ["tillagg", "supplement"],
  },
  {
    offset: 7,
    vba: "TJP(PAR)",
    key: Table1Key.OccupationalPension,
    expectLabel: ["tjanste", "tjp", "occupational"],
  },
  {
    offset: 8,
    vba: "ips(PAR)",
    key: Table1Key.PrivateSaving,
    expectLabel: ["ips", "privat", "sparande", "saving"],
  },
  {
    offset: 15,
    vba: "Netto(Int(PAR))",
    key: null,
    expectLabel: ["skatt", "netto", "tax"],
    notPortedBecause: NOT_PORTED,
  },
  {
    offset: 16,
    vba: "Bidrag(Int(PAR))",
    key: null,
    expectLabel: ["bidrag", "benefit"],
    notPortedBecause: NOT_PORTED,
  },
  {
    offset: 18,
    vba: "IndDisp(Int(PAR))",
    key: null,
    expectLabel: ["disp"],
    notPortedBecause: NOT_PORTED,
  },
];

export function columnSpecs(compareTo: 0 | 1 | 2): readonly ColumnSpec[] {
  return [FIRST_COLUMN[compareTo], ...REST];
}

/** Lower-cases and drops accents, so "Tjanstep" matches "Tjänstep". */
function fold(text: string): string {
  return text
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .toLowerCase();
}

/**
 * Checks the exported labels look like the columns the order says they are.
 *
 * A mismatch does not stop the comparison -- the `Select Case` decides what is
 * in each column, not the sheet's labels -- but it is worth surfacing, because
 * a reordered Mikrosim sheet in a future model version would show up here first.
 */
export function checkLabels(labels: readonly string[], compareTo: 0 | 1 | 2): string[] {
  const specs = columnSpecs(compareTo);
  const complaints: string[] = [];
  labels.forEach((label, i) => {
    const spec = specs[i];
    if (spec === undefined) return;
    const folded = fold(label);
    if (!spec.expectLabel.some((word) => folded.includes(word))) {
      complaints.push(
        `output column ${i + 1} is labelled "${label}", but by its position it holds ` +
          `${spec.vba}`,
      );
    }
  });
  return complaints;
}

export type CellStatus = "exact" | "close" | "off" | "bad" | "missing" | "not-ported";

export interface CellComparison {
  readonly column: number;
  readonly label: string;
  readonly vba: string;
  readonly status: CellStatus;
  readonly expected: number;
  readonly actual: number;
  readonly absolute: number;
  readonly relative: number;
  /**
   * True when the gap is a whole number of the model's own rounding steps: with
   * `marginal = 0` amounts are rounded to whole kronor per month, so a single
   * krona of disagreement lands as exactly 12 a year.
   */
  readonly roundingStep: boolean;
}

export interface CaseComparison {
  readonly index: number;
  readonly line: number;
  readonly inputs: readonly number[];
  readonly cells: readonly CellComparison[];
  readonly warnings: readonly { field: string; given: number | string; used: number | string }[];
  readonly failed?: string;
}

/** Half an öre: below this the two sides agree to the precision the CSV carries. */
const EXACT = 1e-6;
const CLOSE_ABSOLUTE = 0.005;
const CLOSE_RELATIVE = 1e-6;
const OFF_RELATIVE = 0.01;

function classify(
  expected: number,
  actual: number,
): { status: CellStatus; absolute: number; relative: number } {
  if (!Number.isFinite(expected)) return { status: "missing", absolute: Number.NaN, relative: Number.NaN };

  const absolute = actual - expected;
  const magnitude = Math.abs(absolute);
  // Scaled against the amount itself, with a one-krona floor so a pension of
  // zero against twelve kronor does not divide by nothing.
  const relative = magnitude / Math.max(Math.abs(expected), 1);

  if (magnitude <= EXACT) return { status: "exact", absolute, relative };
  if (magnitude <= CLOSE_ABSOLUTE || relative <= CLOSE_RELATIVE)
    return { status: "close", absolute, relative };
  if (relative <= OFF_RELATIVE) return { status: "off", absolute, relative };
  return { status: "bad", absolute, relative };
}

function isRoundingStep(absolute: number, monthly: boolean): boolean {
  const step = monthly ? 1 : 12;
  const magnitude = Math.abs(absolute);
  if (magnitude === 0 || magnitude > step * 5) return false;
  return Math.abs(magnitude / step - Math.round(magnitude / step)) < 1e-6;
}

const amount = (row: Table1Row | undefined, monthly: boolean): number =>
  row === undefined ? Number.NaN : monthly ? row.adjusted / 12 : row.adjusted;

/** Compares one case's twelve columns. */
export function compareCase(
  row: GoldenCase,
  result: TypfallResult,
  settings: GoldenSettings,
  labels: readonly string[],
): CaseComparison {
  const specs = columnSpecs(settings.compareTo);
  const byKey = new Map(result.table1.map((r) => [r.key, r]));

  const cells = specs.map((spec, i): CellComparison => {
    const expected = row.outputs[i] ?? Number.NaN;
    const label = labels[i] ?? `column ${i + 1}`;

    if (spec.key === null) {
      return {
        column: i + 1,
        label,
        vba: spec.vba,
        status: "not-ported",
        expected,
        actual: Number.NaN,
        absolute: Number.NaN,
        relative: Number.NaN,
        roundingStep: false,
      };
    }

    const actual = amount(byKey.get(spec.key), settings.monthly);
    const { status, absolute, relative } = classify(expected, actual);
    return {
      column: i + 1,
      label,
      vba: spec.vba,
      status,
      expected,
      actual,
      absolute,
      relative,
      roundingStep: isRoundingStep(absolute, settings.monthly),
    };
  });

  return {
    index: row.index,
    line: row.line,
    inputs: row.inputs,
    cells,
    warnings: result.warnings.map((w) => ({ field: w.field, given: w.given, used: w.used })),
  };
}

/** A case the engine could not run at all, which is a result in itself. */
export function failedCase(row: GoldenCase, error: unknown): CaseComparison {
  return {
    index: row.index,
    line: row.line,
    inputs: row.inputs,
    cells: [],
    warnings: [],
    failed: error instanceof Error ? error.message : String(error),
  };
}

export const NOT_PORTED_REASON = NOT_PORTED;
