/**
 * Mikrosim's own CSV: the 9 input columns, then the 12 output columns,
 * headers exactly as `mikrosim.ts`'s `INPUT_COLUMNS`/`OUTPUT_COLUMNS` define
 * them (see that file's own comment on why those headers are their own
 * hardcoded table rather than `apps/web/src/i18n.ts`'s `t()`).
 *
 * Every numeric cell, on both sides, is a plain unformatted number -- no
 * thousands grouping, and a decimal point regardless of language. That is a
 * deliberate departure from `table1ToCsv`'s own locale-formatted style
 * (`kronor`/`percent`, comma decimals for Swedish): those exports are meant
 * to be read once in Excel, never read back by this app, so formatting them
 * for a human is the right call. This file's CSV is a genuine round trip
 * (export, hand-edit, re-import; or a file built entirely by hand) and nine
 * of its columns are re-parsed, so a plain, unambiguous number is safer than
 * a locale-formatted one -- it sidesteps thousands-separator parsing
 * entirely and never collides with either delimiter convention (`;` or `,`).
 * Only `delimiterFor` -- the cell separator, not the numbers themselves -- is
 * still reused from `tables.ts`'s existing per-language choice. On *import*,
 * both a comma and a point are still accepted as the decimal mark (see
 * `parseNumber`), since a real Mikrosim export from Excel, or a file someone
 * has hand-edited, may well use a Swedish comma.
 *
 * Output columns are written for reference only -- read back never parses
 * them, so a value here does not need to round-trip. A row with no `result`
 * yet exports with its 12 output cells blank rather than being dropped:
 * silently omitting an uncalculated row would make the file's row count
 * disagree with the table's own.
 */
import type { Lang } from "./i18n.js";
import type { InputColumnDef, MikrosimRow } from "./mikrosim.js";
import { INPUT_COLUMNS, OUTPUT_COLUMNS, headerName, newMikrosimRow, validateMikrosimRow } from "./mikrosim.js";
import { csvLine, delimiterFor } from "./tables.js";

const say = (l: Lang, sv: string, en: string) => (l === "sv" ? sv : en);

export interface MikrosimParseResult {
  readonly rows: MikrosimRow[];
  /** Set when a required column is missing -- the whole file is refused,
   * since there is nowhere to attribute the problem to one row. */
  readonly fileError?: string;
}

/** Whichever of `;`/`,` splits the header line into more pieces wins -- the
 * header is the one line guaranteed to contain only column names, no numbers
 * that could themselves contain either character. */
function sniffDelimiter(headerLine: string): string {
  const semicolons = headerLine.split(";").length;
  const commas = headerLine.split(",").length;
  return semicolons >= commas ? ";" : ",";
}

/** Accepts a typed "," or "." either way, the same tolerant rule
 * `controls.ts`'s own `percent` field applies to a typed value. */
function parseNumber(raw: string): number | undefined {
  const trimmed = raw.trim();
  if (trimmed === "") return undefined;
  const value = Number(trimmed.replace(",", "."));
  return Number.isFinite(value) ? value : undefined;
}

/**
 * Parses a Mikrosim CSV -- the 9 input columns, matched by header text (case-
 * insensitive, either language), not by position, so a hand-adapted real
 * Excel export (which may have dropped or reordered a column) still lines
 * up. No RFC4180 quoted-field handling: none of the 9 columns are ever free
 * text, so a plain per-line split on the sniffed delimiter is sufficient,
 * and a stray delimiter inside one of the (ignored, on import) output
 * columns can only ever add harmless extra trailing cells, never disturb an
 * input column that comes before it in the row.
 */
export function parseMikrosimCsv(text: string, lang: Lang): MikrosimParseResult {
  const withoutBom = text.replace(/^﻿/, "");
  const lines = withoutBom.split(/\r\n|\r|\n/);

  let i = 0;
  while (i < lines.length && (lines[i]!.trim() === "" || lines[i]!.trim().startsWith("#"))) {
    i += 1;
  }
  if (i >= lines.length) {
    return { rows: [], fileError: say(lang, "Filen är tom.", "The file is empty.") };
  }

  const delimiter = sniffDelimiter(lines[i]!);
  const headerCells = lines[i]!.split(delimiter).map((cell) => cell.trim().toLowerCase());

  const cellIndexFor: number[] = [];
  const missing: string[] = [];
  for (const col of INPUT_COLUMNS) {
    const at = headerCells.findIndex((cell) => cell === col.sv.toLowerCase() || cell === col.en.toLowerCase());
    cellIndexFor.push(at);
    if (at === -1) missing.push(headerName(col, lang));
  }
  if (missing.length > 0) {
    return {
      rows: [],
      fileError: say(
        lang,
        `Filen saknar kolumnen "${missing.join('", "')}".`,
        `The file is missing the column "${missing.join('", "')}".`,
      ),
    };
  }

  const rows: MikrosimRow[] = [];
  let nextId = 1;
  for (let li = i + 1; li < lines.length; li += 1) {
    const line = lines[li]!;
    if (line.trim() === "") continue;
    const cells = line.split(delimiter);
    const row = newMikrosimRow(String(nextId));
    nextId += 1;

    let error: string | undefined;
    INPUT_COLUMNS.forEach((col: InputColumnDef, colIdx) => {
      const raw = cells[cellIndexFor[colIdx]!];
      if (raw === undefined) {
        error ??= say(lang, "Raden saknar en eller flera kolumner.", "The row is missing one or more columns.");
        return;
      }
      const value = parseNumber(raw);
      if (value === undefined) {
        error ??= say(
          lang,
          `Kan inte tolka "${raw.trim()}" som ett tal i kolumnen "${headerName(col, lang)}".`,
          `Can't read "${raw.trim()}" as a number in the column "${headerName(col, lang)}".`,
        );
        return;
      }
      if (col.discrete) {
        if (!col.discrete.includes(value)) {
          error ??= say(
            lang,
            `Ogiltigt värde i kolumnen "${headerName(col, lang)}": ${raw.trim()}.`,
            `Invalid value in the column "${headerName(col, lang)}": ${raw.trim()}.`,
          );
          return;
        }
        col.set(row, value);
      } else {
        col.set(row, col.clamp(value));
      }
    });

    const rowError = error ?? validateMikrosimRow(row, lang);
    if (rowError !== undefined) row.error = rowError;
    rows.push(row);
  }

  return { rows };
}

function formatNumber(value: number): string {
  return String(value);
}

/** Mikrosim's rows as a CSV -- see the file comment on why every cell is a
 * plain unformatted number rather than `kronor`/`percent`-styled text. */
export function mikrosimRowsToCsv(rows: readonly MikrosimRow[], lang: Lang): string {
  const delimiter = delimiterFor(lang);
  const header = [
    ...INPUT_COLUMNS.map((col) => headerName(col, lang)),
    ...OUTPUT_COLUMNS.map((col) => headerName(col, lang)),
  ];
  const lines = [csvLine(header, delimiter)];
  for (const row of rows) {
    const inputs = INPUT_COLUMNS.map((col) => formatNumber(col.get(row)));
    const outputs = OUTPUT_COLUMNS.map((col) => {
      const found = row.result?.table1.find((r) => r.key === col.key);
      return found ? formatNumber(Math.round(found.adjusted)) : "";
    });
    lines.push(csvLine([...inputs, ...outputs], delimiter));
  }
  return lines.join("\r\n");
}
