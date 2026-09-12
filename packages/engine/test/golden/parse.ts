/**
 * Reads the golden file the Excel export writes.
 *
 * The format is described in reference/golden/HOWTO.md and produced by
 * reference/golden/ExportGoldenCases.bas: a block of `#` provenance lines, a
 * header row taken from row 7 of the Mikrosim sheet, then one line per case
 * with ten input columns (B-K) and twelve output columns (M-X).
 */

import { readFileSync } from "node:fs";

/** Columns B-K on the Mikrosim sheet, in the order `InputXGetY` reads them. */
export const INPUT_COLUMNS = 10;
/** Columns M-X, in the order `InputXGetY` writes them. */
export const OUTPUT_COLUMNS = 12;

/** One exported case: the inputs the workbook was given and what it returned. */
export interface GoldenCase {
  /** Position in the file, from 1. Also the case number `BuildCases` generated. */
  readonly index: number;
  /** Line number in the CSV, so the report can point at the source. */
  readonly line: number;
  /** Columns B-K. */
  readonly inputs: readonly number[];
  /** Columns M-X. `NaN` where the workbook wrote an empty cell. */
  readonly outputs: readonly number[];
}

export interface GoldenFile {
  /**
   * The `#` block, keyed by the name before the colon, lower-cased. Settings
   * from the Adv_settings dump are keyed `adv.<name>`; the Start sheet's
   * civilstand is `start.gift`; the hand-picked older keys (`referensar`,
   * `belopp12`, ...) are kept as they are, so a file exported before the block
   * was widened still parses.
   */
  readonly settings: ReadonlyMap<string, string>;
  readonly inputLabels: readonly string[];
  readonly outputLabels: readonly string[];
  readonly cases: readonly GoldenCase[];
}

/**
 * Decodes the file.
 *
 * VBA's `Print #` writes the machine's ANSI code page, not UTF-8, so a Swedish
 * Excel puts single high bytes where the labels have a-ring or o-umlaut. Try
 * UTF-8 first and fall back to windows-1252, which is what those bytes are.
 */
function decode(bytes: Buffer): string {
  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch {
    return new TextDecoder("windows-1252").decode(bytes);
  }
}

/**
 * Parses one exported number.
 *
 * `CsvNum` writes an empty field for a cell that is not numeric, which is how a
 * row that never ran looks. A comma means the macro's decimal-separator guard
 * failed, which would silently shift every column after it, so it is an error
 * rather than something to work around.
 */
function parseNumber(field: string, line: number, column: string): number {
  const value = field.trim();
  if (value === "") return Number.NaN;
  if (value.includes(",")) {
    throw new Error(
      `${column} on line ${line} is "${value}": the export wrote a comma decimal ` +
        `separator. See "Numbers in the CSV use commas" in reference/golden/HOWTO.md.`,
    );
  }
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) {
    throw new Error(`${column} on line ${line} is not a number: "${value}"`);
  }
  return parsed;
}

export function parseGoldenFile(text: string): GoldenFile {
  const lines = text.split(/\r?\n/);
  const settings = new Map<string, string>();

  let index = 0;
  for (; index < lines.length; index += 1) {
    const line = lines[index]!;
    if (!line.startsWith("#")) break;
    const body = line.slice(1).trim();
    const colon = body.indexOf(":");
    if (colon < 0) continue; // the title line, which carries no value
    const key = body.slice(0, colon).trim().toLowerCase();
    // The macro appends "   (row 47)" or "   (name)" after some values.
    const value = body.slice(colon + 1).replace(/\s{2,}\(.*\)\s*$/, "").trim();
    settings.set(key, value);
  }

  const headerLine = lines[index];
  if (headerLine === undefined) throw new Error("the golden file has no header row");
  const header = headerLine.split(",").map((cell) => cell.trim());
  if (header.length !== INPUT_COLUMNS + OUTPUT_COLUMNS) {
    throw new Error(
      `the header row has ${header.length} columns, expected ` +
        `${INPUT_COLUMNS + OUTPUT_COLUMNS} (${INPUT_COLUMNS} inputs and ${OUTPUT_COLUMNS} outputs)`,
    );
  }
  const headerLineNumber = index + 1;

  const cases: GoldenCase[] = [];
  for (index += 1; index < lines.length; index += 1) {
    const line = lines[index]!;
    if (line.trim() === "") continue;
    const lineNumber = index + 1;
    const fields = line.split(",");
    if (fields.length !== header.length) {
      // The usual cause: the macro's decimal-separator guard failed and a
      // Swedish "1,7" split itself across two fields.
      const commaDecimals = /\d,\d/.test(line) && fields.length > header.length;
      throw new Error(
        `line ${lineNumber} has ${fields.length} fields, but the header has ` +
          `${header.length}` +
          (commaDecimals
            ? `. The line looks like it carries comma decimal separators, which split ` +
              `the fields. See "Numbers in the CSV use commas" in reference/golden/HOWTO.md.`
            : ""),
      );
    }
    const caseIndex = cases.length + 1;
    cases.push({
      index: caseIndex,
      line: lineNumber,
      inputs: fields
        .slice(0, INPUT_COLUMNS)
        .map((field, i) => parseNumber(field, lineNumber, `input column ${i + 1}`)),
      outputs: fields
        .slice(INPUT_COLUMNS)
        .map((field, i) => parseNumber(field, lineNumber, `output column ${i + 1}`)),
    });
  }

  if (cases.length === 0) {
    throw new Error(`the golden file has a header on line ${headerLineNumber} but no cases`);
  }

  return {
    settings,
    inputLabels: header.slice(0, INPUT_COLUMNS),
    outputLabels: header.slice(INPUT_COLUMNS),
    cases,
  };
}

export function readGoldenFile(path: string): GoldenFile {
  return parseGoldenFile(decode(readFileSync(path)));
}
