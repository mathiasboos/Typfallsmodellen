/**
 * The KPI cards' arithmetic, checked without a browser.
 *
 * `apps/web`'s test environment has no DOM, so `renderKpis` itself -- like
 * every other rendering function here -- is checked against the real built
 * page by `tools/build/verify-offline.mjs` instead. What can and should be
 * unit-tested is the pure computation `computeKpis` does before any of that:
 * which Table 1 row it reads, and the one new piece of arithmetic,
 * `averageAnnualGross`.
 */
import { describe, expect, it } from "vitest";

import { Table1Key } from "@typfallsmodellen/engine";
import type { MvaluesRow, Table1Row, TypfallResult } from "@typfallsmodellen/engine";

import { averageAnnualGross, computeKpis } from "../src/kpis.js";

function row(age: number, brutto: number): MvaluesRow {
  return {
    year: 2000 + age,
    age,
    income: 0,
    ip: 0,
    tp: 0,
    pp: 0,
    garp: 0,
    ptillagg: 0,
    tjp: 0,
    ips: 0,
    brutto,
    netto: 0,
    bidrag: 0,
    pps: 0,
    indDisp: 0,
    kpiFactor: 1,
    indexFactor: 1,
    municipalTax: 0,
    stateTax: 0,
  };
}

/** Only `rows`, `table1` and `lifeIncome` matter here; the rest is padding. */
function stubResult(overrides: {
  rows?: readonly MvaluesRow[];
  table1?: readonly Table1Row[];
  throughAge?: number;
}): TypfallResult {
  return {
    table1: overrides.table1 ?? [],
    table2: [],
    rows: overrides.rows ?? [],
    wagePath: [],
    lifeIncome: { gross: 0, net: 0, disposable: 0, throughAge: overrides.throughAge ?? 0 },
    qualifyingYears: 40,
    warnings: [],
  };
}

describe("averageAnnualGross", () => {
  it("averages brutto over whole ages from retirement through the expected age at death", () => {
    // par = 66, throughAge truncates to 68 -- ages 64, 65 and 69 sit outside
    // the window and must not pull the average toward their (very different)
    // values.
    const result = stubResult({
      rows: [row(64, 999_999), row(65, 999_999), row(66, 100), row(67, 200), row(68, 300), row(69, -999_999)],
      throughAge: 68.7,
    });
    expect(averageAnnualGross(result, 66)).toBeCloseTo((100 + 200 + 300) / 3);
  });

  it("includes both endpoints of the window", () => {
    const result = stubResult({ rows: [row(70, 10), row(71, 20)], throughAge: 71 });
    expect(averageAnnualGross(result, 70)).toBeCloseTo(15);
  });

  it("returns 0 rather than NaN when no row falls in the window", () => {
    const result = stubResult({ rows: [row(50, 1)], throughAge: 40 });
    expect(averageAnnualGross(result, 66)).toBe(0);
  });
});

describe("computeKpis", () => {
  it("reads Table 1's own TotalGross row rather than re-deriving it from rows", () => {
    // Deliberately different from the row-based figure at the retirement age,
    // the way the real engine's two numbers differ (closeRetirementYear
    // rebuilds the retirement year's gross for Table 1 after that age's row
    // was already written) -- computeKpis must report Table 1's, not this
    // row's.
    const result = stubResult({
      table1: [
        {
          key: Table1Key.TotalGross,
          nominal: 0,
          adjusted: 0,
          monthly: 20_154,
          shareOfFinalSalary: 0.5275,
        },
      ],
      rows: [row(66, 239_460)],
      throughAge: 66.4,
    });

    const kpis = computeKpis(result, 66);
    expect(kpis.monthlyAtRetirement).toBe(20_154);
    expect(kpis.replacementRate).toBe(0.5275);
  });

  it("falls back to 0 rather than throwing when Table 1 has no TotalGross row", () => {
    const kpis = computeKpis(stubResult({ throughAge: 66 }), 66);
    expect(kpis.monthlyAtRetirement).toBe(0);
    expect(kpis.replacementRate).toBe(0);
  });

  it("truncates both ends of the age window it reports", () => {
    const result = stubResult({ rows: [row(66, 1200)], throughAge: 86.9 });
    const kpis = computeKpis(result, 66);
    expect(kpis.firstAge).toBe(66);
    expect(kpis.lastAge).toBe(86);
    expect(kpis.averageMonthly).toBeCloseTo(100);
  });
});
