import { mkdirSync, writeFileSync } from "node:fs";
import { dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { beforeAll, describe, expect, it } from "vitest";

import { compareCase, checkLabels } from "./compare.js";
import { GOLDEN_PATH, compareGoldenFile, blockOf, goldenFileExists } from "./harness.js";
import type { ComparisonRun } from "./harness.js";
import { readSettings, toInput } from "./map.js";
import { parseGoldenFile } from "./parse.js";
import { formatReport, formatSummary, summariseColumns } from "./report.js";

/**
 * The acceptance gate: the engine against the real Typfallsmodellen.
 *
 * The comparison needs `reference/golden/golden-cases.csv`, which only Excel can
 * produce -- see reference/golden/HOWTO.md. Until it is committed these tests
 * skip, and the harness itself is covered by the unit tests below, which run on
 * a fixture written out here rather than on the workbook's output.
 */

const REPORT_PATH = fileURLToPath(
  new URL("../../../../reference/golden/report.md", import.meta.url),
);

/** A miniature golden file, in the exact shape the export macro writes. */
const FIXTURE = [
  "# Typfallsmodellen reference results",
  "# model: Typfallsmodellen 2025-01-15",
  "# workbook: Typfallsmodellen.xlsb",
  "# exported: 2026-09-12 10:00:00",
  "# referensar: 2025",
  "# marginal: 0",
  "# belopp12: 0",
  "# compareTo: 0",
  "# cases: 2",
  "# start.Gift: 0",
  "# adv.w_ref: 2025   (row 64)",
  "# adv.Rng_belopp12: 0   (row 66)",
  "# adv.Average_Earning: 5   (row 103)",
  "# adv.Alt_p_age: 0   (row 67)",
  "# adv.Risk: 0   (row 95)",
  "Fodelsear,Borjar arbeta,Pensionsalder,Arslon,Inflation,Real tillvaxt,Avkastning,Egen lon,IPS,Avtal," +
    "Slutlon,Total pension,Inkomstpension,Tillaggspension,Premiepension,Garantipension,Inkomstpensionstillagg," +
    "Tjanstepension,IPS,Efter skatt,Bidrag,Disponibel inkomst",
  "1959,23,66,462000,0,0,0.017,0,0,1,458000,244000,200000,0,44000,0,0,0,0,190000,0,190000",
  "1970,25,67,324000,0,0,0.017,0,0,5,321000,190000,150000,0,40000,0,0,0,0,150000,0,150000",
].join("\n");

describe("the golden-file harness", () => {
  const file = parseGoldenFile(FIXTURE);

  it("splits the provenance block from the data", () => {
    expect(file.cases).toHaveLength(2);
    expect(file.settings.get("model")).toBe("Typfallsmodellen 2025-01-15");
    // The trailing "   (row 64)" the macro appends is not part of the value.
    expect(file.settings.get("adv.w_ref")).toBe("2025");
    expect(file.outputLabels).toHaveLength(12);
    expect(file.cases[0]!.line).toBe(17);
  });

  it("reads the settings the numbers have to be interpreted through", () => {
    const settings = readSettings(file);
    expect(settings.monthly).toBe(false);
    expect(settings.compareTo).toBe(0);
    expect(settings.married).toBe(false);
    expect(settings.assumed).toEqual([]);
    expect(settings.issues.filter((issue) => issue.fatal)).toEqual([]);
  });

  it("maps the input columns the way InputXGetY reads them", () => {
    const settings = readSettings(file);
    const input = toInput(file.cases[0]!, settings);
    expect(input.born).toBe(1959);
    expect(input.startWorkAge).toBe(23);
    expect(input.retirementAge).toBe(66);
    // Column E is an annual salary; the runner divides it by twelve.
    expect(input.monthlySalary).toBe(462000 / 12);
    expect(input.realReturn).toBe(0.017);
    expect(input.scheme).toBe(1);
  });

  it("refuses a file whose rows do not match its own header", () => {
    expect(() => parseGoldenFile(FIXTURE.replace(",1,458000", ",1,458000,99"))).toThrow(
      /has 23 fields/,
    );
  });

  it("refuses comma decimals rather than reading them as extra columns", () => {
    // A comma decimal splits the field, so the row no longer matches the header
    // -- which is where the reader lands, so that is where it has to be named.
    const swedish = FIXTURE.replace("0.017,0,0,1,", "0,017,0,0,1,");
    expect(() => parseGoldenFile(swedish)).toThrow(/comma decimal separators/);
  });

  it("treats a setting that changes what the numbers mean as fatal", () => {
    const wrong = parseGoldenFile(FIXTURE.replace("# adv.Alt_p_age: 0", "# adv.Alt_p_age: 1"));
    const issues = readSettings(wrong).issues;
    // Reported under the spelling Adv_settings gives it, not the lower-cased key.
    expect(issues.some((issue) => issue.fatal && issue.setting === "Alt_p_age")).toBe(true);
  });

  it("assumes and says so when the file predates the widened block", () => {
    const older = FIXTURE.split("\n")
      .filter((line) => !line.startsWith("# adv.") && !line.startsWith("# start."))
      .join("\n")
      .replace("# belopp12: 0\n", "")
      .replace("# compareTo: 0\n", "");
    const settings = readSettings(parseGoldenFile(older));
    expect(settings.assumed.length).toBeGreaterThan(0);
  });

  it("checks the labels against the column order without trusting them", () => {
    expect(checkLabels(file.outputLabels, 0)).toEqual([]);
    const shuffled = [...file.outputLabels];
    [shuffled[2], shuffled[3]] = [shuffled[3]!, shuffled[2]!];
    expect(checkLabels(shuffled, 0).length).toBeGreaterThan(0);
  });

  it("knows which block a case number belongs to", () => {
    expect(blockOf(1)).toBe("A");
    expect(blockOf(144)).toBe("A");
    expect(blockOf(145)).toBe("B");
    expect(blockOf(295)).toBe("F");
  });

  it("reports the three unported columns rather than scoring them", () => {
    const settings = readSettings(file);
    const result = {
      table1: [{ key: "slutlon", nominal: 0, adjusted: 458000, monthly: 0, shareOfFinalSalary: 0 }],
      warnings: [],
    } as unknown as Parameters<typeof compareCase>[1];
    const comparison = compareCase(file.cases[0]!, result, settings, file.outputLabels);
    expect(comparison.cells.filter((cell) => cell.status === "not-ported")).toHaveLength(3);
    expect(comparison.cells[0]!.status).toBe("exact");
  });
});

describe.skipIf(!goldenFileExists())("the engine against the real model", () => {
  let comparison: ComparisonRun;

  beforeAll(() => {
    comparison = compareGoldenFile();
    mkdirSync(dirname(REPORT_PATH), { recursive: true });
    writeFileSync(REPORT_PATH, `${formatReport(comparison)}\n`, "utf8");
    // eslint-disable-next-line no-console
    console.log(`\n${formatSummary(comparison)}\n\n  full report: ${REPORT_PATH}\n`);
  }, 600_000);

  it("ran every case in the file", () => {
    const failed = comparison.cases.filter((c) => c.failed !== undefined);
    expect(
      failed.map((c) => `case ${c.index} (line ${c.line}): ${c.failed}`),
      `${failed.length} of ${comparison.cases.length} cases could not be run`,
    ).toEqual([]);
  });

  it("used inputs the workbook did not have to correct", () => {
    // The export macro clamps the retirement age to each cohort's earliest, so
    // a correction here means the two sides computed different people.
    const corrected = comparison.cases.filter((c) => c.warnings.length > 0);
    expect(corrected.map((c) => `case ${c.index}: ${JSON.stringify(c.warnings)}`)).toEqual([]);
  });

  it.each(
    // One test per column, so a failure names the column rather than the file.
    Array.from({ length: 12 }, (_, i) => i + 1),
  )("agrees with the workbook in output column %i", (column) => {
    const summaries = summariseColumns(comparison.cases);
    const summary = summaries.find((s) => s.column === column);
    expect(summary, `column ${column} is missing from the comparison`).toBeDefined();
    if (summary!.notPorted) return;

    const divergent = comparison.cases
      .map((c) => c.cells.find((cell) => cell.column === column))
      .filter((cell) => cell !== undefined && (cell.status === "off" || cell.status === "bad" || cell.status === "missing"));

    expect(
      divergent.length,
      `${divergent.length} of ${summary!.compared} cases diverge in column ${column} ` +
        `(${summary!.label}); worst is case ${summary!.worstCase}. See ${REPORT_PATH}.`,
    ).toBe(0);
  });
});
