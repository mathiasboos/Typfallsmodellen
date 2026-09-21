/**
 * A thin wrapper around `write-excel-file` -- this app's only third-party
 * runtime dependency (everything else is this port's own code, or the
 * generated `@typfallsmodellen/data`/`@typfallsmodellen/engine` packages).
 * Kept in its own module so that dependency is isolated to one file, the
 * same way `deaths.ts` isolates the generated mortality data.
 *
 * `/universal` rather than `/browser`: the `/browser` entry point runs in a
 * Web Worker, which needs its own script URL -- incompatible with this
 * app's single offline HTML file, which must open from `file://` with zero
 * outbound requests (`verify-offline.mjs`'s own "outbound blocked" check).
 * `/universal` does the same work on the main thread and returns a `Blob`
 * directly, at the cost of a (typically sub-millisecond) main-thread pause
 * for a table this small.
 */
import writeXlsxFile from "write-excel-file/universal";
import type { SheetData } from "write-excel-file/universal";

export type { Cell, Row, SheetData } from "write-excel-file/universal";

export function xlsxBlob(sheetData: SheetData): Promise<Blob> {
  return writeXlsxFile(sheetData).toBlob();
}

/** A kronor amount: the real number, formatted with a thousands separator
 * and no decimals -- the same rounding `format.ts`'s own `kronor()` shows,
 * but as a live number Excel can still add up, not text. */
export function kronorCell(value: number) {
  return { value, type: Number, format: "#,##0" };
}

/** A share, stored as the 0-1 fraction Excel's own percent format expects --
 * the same convention `format.ts`'s own `percent()` already uses. */
export function percentCell(value: number, digits = 1) {
  return { value, type: Number, format: `0.${"0".repeat(digits)}%` };
}
