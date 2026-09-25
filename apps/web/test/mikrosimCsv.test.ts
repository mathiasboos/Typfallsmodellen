import { describe, expect, it } from "vitest";

import { mikrosimRowsToCsv, parseMikrosimCsv } from "../src/mikrosimCsv.js";
import { newMikrosimRow } from "../src/mikrosim.js";

/**
 * `parseMikrosimCsv`/`mikrosimRowsToCsv` are the one piece of this feature
 * that does not need a browser -- reading and writing a CSV is plain string
 * processing, with the DOM-facing file-picker and download plumbing left to
 * `tools/build/verify-offline.mjs`, the same split `compare.test.ts` already
 * uses for `applyScenario`.
 */
describe("parsing a Mikrosim CSV", () => {
  const HEADER =
    "Välj tjänstepension;Födelseår;Årslön;Börjar arbeta vid ålder;Går i pension vid ålder;" +
    "Årlig inflation;Real tillväxt;Real fondavkastning;Privat pensionsförsäkring";

  it("matches columns by header text, in scrambled order, not by position", () => {
    const csv = `${HEADER}\n3;1970;480000;20;68;0,02;0,01;0,03;500\n`;
    const result = parseMikrosimCsv(csv, "sv");
    expect(result.fileError).toBeUndefined();
    expect(result.rows).toHaveLength(1);
    const row = result.rows[0]!;
    expect(row.scheme).toBe(3);
    expect(row.born).toBe(1970);
    expect(row.annualSalary).toBe(480_000);
    expect(row.startWorkAge).toBe(20);
    expect(row.retirementAge).toBe(68);
    expect(row.yearlyInflation).toBeCloseTo(0.02);
    expect(row.realGrowth).toBeCloseTo(0.01);
    expect(row.realReturn).toBeCloseTo(0.03);
    expect(row.ipsMonthly).toBe(500);
  });

  it("accepts a ,-delimited, .-decimal file, a leading BOM and a #-comment provenance line", () => {
    const en =
      "Occupational pension scheme,Birth year,Annual salary,Starts working at age,Retires at age," +
      "Yearly inflation,Real growth,Real fund return,Private pension insurance";
    const csv = `﻿# exported by typfallsmodellen\n${en}\n1,1959,462000,23,66,0,0,0.017,0\n`;
    const result = parseMikrosimCsv(csv, "en");
    expect(result.fileError).toBeUndefined();
    expect(result.rows).toHaveLength(1);
    expect(result.rows[0]!.realReturn).toBeCloseTo(0.017);
  });

  it("refuses the whole file, naming every missing column, when one is absent", () => {
    const result = parseMikrosimCsv("Födelseår;Går i pension vid ålder\n1970;65\n", "sv");
    expect(result.rows).toHaveLength(0);
    expect(result.fileError).toContain("Börjar arbeta vid ålder");
    expect(result.fileError).toContain("Årslön");
  });

  it("flags only the offending row, not the file, for an unparseable cell or an invalid scheme", () => {
    const csv =
      `${HEADER}\n` +
      "3;1970;480000;20;68;0,02;0,01;0,03;500\n" + // good
      "3;abc;480000;20;68;0,02;0,01;0,03;500\n" + // unparseable birth year
      "99;1970;480000;20;68;0,02;0,01;0,03;500\n"; // invalid scheme
    const result = parseMikrosimCsv(csv, "sv");
    expect(result.fileError).toBeUndefined();
    expect(result.rows).toHaveLength(3);
    expect(result.rows[0]!.error).toBeUndefined();
    expect(result.rows[1]!.error).toBeDefined();
    expect(result.rows[2]!.error).toBeDefined();
  });

  it("clamps a continuous value that is out of bounds rather than rejecting the row", () => {
    // 5 is below START_WORK's own minimum of 15.
    const csv = `${HEADER}\n3;1970;480000;5;68;0,02;0,01;0,03;500\n`;
    const result = parseMikrosimCsv(csv, "sv");
    expect(result.fileError).toBeUndefined();
    expect(result.rows).toHaveLength(1);
    expect(result.rows[0]!.error).toBeUndefined();
    expect(result.rows[0]!.startWorkAge).toBe(15);
  });
});

describe("round-tripping Mikrosim's own CSV", () => {
  it("recovers equivalent rows from its own export, in both languages", () => {
    const rows = [
      { ...newMikrosimRow("1"), born: 1970, annualSalary: 480_000, realReturn: 0.017 },
      { ...newMikrosimRow("2"), born: 1985, annualSalary: 360_000, scheme: 5 as const },
    ];
    for (const lang of ["sv", "en"] as const) {
      const csv = mikrosimRowsToCsv(rows, lang);
      const parsed = parseMikrosimCsv(csv, lang);
      expect(parsed.fileError).toBeUndefined();
      expect(parsed.rows).toHaveLength(rows.length);
      parsed.rows.forEach((row, i) => {
        const original = rows[i]!;
        expect(row.born).toBe(original.born);
        expect(row.startWorkAge).toBe(original.startWorkAge);
        expect(row.retirementAge).toBe(original.retirementAge);
        expect(row.annualSalary).toBe(original.annualSalary);
        expect(row.yearlyInflation).toBeCloseTo(original.yearlyInflation);
        expect(row.realGrowth).toBeCloseTo(original.realGrowth);
        expect(row.realReturn).toBeCloseTo(original.realReturn);
        expect(row.ipsMonthly).toBe(original.ipsMonthly);
        expect(row.scheme).toBe(original.scheme);
      });
    }
  });

  it("writes every row, including one with no result yet, with its output columns blank", () => {
    const rows = [newMikrosimRow("1")];
    const csv = mikrosimRowsToCsv(rows, "sv");
    const lines = csv.split("\r\n");
    expect(lines).toHaveLength(2);
    const cells = lines[1]!.split(";");
    // 9 input cells, then 12 output cells, all blank since `result` is unset.
    expect(cells).toHaveLength(21);
    expect(cells.slice(9)).toEqual(Array(12).fill(""));
  });
});
