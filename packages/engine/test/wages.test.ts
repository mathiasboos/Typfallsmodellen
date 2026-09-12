import { describe, expect, it } from "vitest";

import economicSeriesJson from "../../data/economic-series.json" with { type: "json" };
import { PROGNOSSTANDARD, projectEconomicData } from "../src/data/index.js";
import type { RawEconomicSeries } from "../src/data/index.js";
import { wages, withdrawalShare } from "../src/income/wages.js";
import type { WageProfile } from "../src/income/wages.js";

const data = projectEconomicData(economicSeriesJson as unknown as RawEconomicSeries, PROGNOSSTANDARD);

/**
 * No fixture exists for the wage vector -- the Brutto sheet is fed a hand-typed
 * income list rather than `wages` output. These assert shape and the arithmetic
 * that is stated outright in the VBA, not amounts that only Excel could confirm.
 */
const profile = (overrides: Partial<WageProfile> = {}): WageProfile => ({
  born: 1959,
  par: 66,
  defAr: 66,
  wageProfile: 0,
  workDuringPartialWithdrawal: "Arbetar deltid",
  partialWithdrawalShare: 1,
  ...overrides,
});

const at = (age: number, p = profile(), share = 1) =>
  wages(age, p.born + age, 23, 462_000, p, data, share, 66, 2025, false);

describe("withdrawalShare", () => {
  const p = profile({ par: 66, defAr: 70, partialWithdrawalShare: 0.5 });

  it("is nothing before pension is first drawn", () => {
    expect(withdrawalShare(60, p)).toBe(0);
    expect(withdrawalShare(65, p)).toBe(0);
  });

  it("is the partial share between first and final withdrawal", () => {
    expect(withdrawalShare(66, p)).toBe(0.5);
    expect(withdrawalShare(69, p)).toBe(0.5);
  });

  it("is everything once withdrawal is final", () => {
    expect(withdrawalShare(71, p)).toBe(1);
  });

  it("falls through in the year pension is both first and finally drawn", () => {
    // par === defAr is the ordinary case, and it matches no branch: the VBA
    // leaves its global untouched, so the share is still 0 that year and only
    // becomes 1 the year after. The line that looks like it handles this sits
    // inside a branch par === defAr can never reach.
    const full = profile({ par: 66, defAr: 66 });
    expect(withdrawalShare(66, full, 0)).toBe(0);
    expect(withdrawalShare(67, full, 0)).toBe(1);
  });

  it("fixes work at a set share when one is given as a number", () => {
    const fixed = profile({ par: 66, defAr: 70, workDuringPartialWithdrawal: 0.4 });
    expect(withdrawalShare(67, fixed)).toBeCloseTo(0.6, 12);
  });

  it("carries the previous value forward where the VBA assigns nothing", () => {
    // age === defAr falls through every branch; the global keeps its value.
    expect(withdrawalShare(70, p, 0.5)).toBe(0.5);
  });
});

describe("wages", () => {
  it("earns nothing before entering the labour market", () => {
    expect(at(20)).toBe(0);
    expect(at(22)).toBe(0);
    expect(at(23)).toBeGreaterThan(0);
  });

  it("earns nothing after final withdrawal", () => {
    expect(at(67)).toBe(0);
    expect(at(80)).toBe(0);
  });

  it("returns the entered salary in the reference year", () => {
    // Salary refers to age 66 in fixed 2025 prices; at that age the wage path
    // has neither grown nor shrunk relative to itself.
    const p = profile({ par: 80, defAr: 80 });
    expect(wages(66, p.born + 66, 23, 462_000, p, data, 0, 66, 2025, false)).toBeCloseTo(462_000, 6);
  });

  it("follows the income index across a working life", () => {
    const p = profile({ par: 80, defAr: 80 });
    const earlier = wages(30, p.born + 30, 23, 462_000, p, data, 0, 66, 2025, false);
    const later = wages(60, p.born + 60, 23, 462_000, p, data, 0, 66, 2025, false);
    expect(later).toBeGreaterThan(earlier);
    const indexRatio =
      data.inkomstindex.at(p.born + 60) / data.inkomstindex.at(p.born + 30);
    expect(later / earlier).toBeCloseTo(indexRatio, 9);
  });

  it("scales the entry year by the months actually worked", () => {
    // Born mid-year, so the first working year is a part year.
    const midYear = profile({ born: 1959.5, par: 80, defAr: 80 });
    const first = wages(23, 1982, 23, 462_000, midYear, data, 0, 66, 2025, false);
    const second = wages(24, 1983, 23, 462_000, midYear, data, 0, 66, 2025, false);
    expect(first).toBeGreaterThan(0);
    expect(first).toBeLessThan(second);
  });

  it("stops earning in proportion to how much pension is drawn", () => {
    const p = profile({ par: 66, defAr: 70 });
    const full = wages(67, p.born + 67, 23, 462_000, p, data, 0, 66, 2025, false);
    const half = wages(67, p.born + 67, 23, 462_000, p, data, 0.5, 66, 2025, false);
    const none = wages(67, p.born + 67, 23, 462_000, p, data, 1, 66, 2025, false);
    expect(half).toBeCloseTo(full / 2, 6);
    expect(none).toBe(0);
  });

  it("bends the wage path when an age profile is chosen", () => {
    const straight = profile({ par: 80, defAr: 80, wageProfile: 0 });
    const curved = profile({ par: 80, defAr: 80, wageProfile: 2 });
    const young = (p: WageProfile) => wages(25, p.born + 25, 23, 462_000, p, data, 0, 66, 2025, false);
    expect(young(curved)).not.toBeCloseTo(young(straight), 0);
    // A profile should still leave the reference age unchanged.
    expect(wages(66, curved.born + 66, 23, 462_000, curved, data, 0, 66, 2025, false)).toBeCloseTo(
      462_000,
      6,
    );
  });

  it("holds profile 1 flat above 61, unlike the others", () => {
    const p = profile({ par: 80, defAr: 80, wageProfile: 1 });
    const w = (age: number) => wages(age, p.born + age, 23, 462_000, p, data, 0, 61, 2025, false);
    // Above 61 only the index moves the wage, not the profile.
    const ratio62 = w(62) / w(61);
    const indexRatio = data.inkomstindex.at(p.born + 62) / data.inkomstindex.at(p.born + 61);
    expect(ratio62).toBeCloseTo(indexRatio, 9);
  });

  it("differs between current and fixed prices only by the KPI ratio", () => {
    const p = profile({ par: 80, defAr: 80 });
    const nominal = wages(50, p.born + 50, 23, 462_000, p, data, 0, 66, 2025, true);
    const fixed = wages(50, p.born + 50, 23, 462_000, p, data, 0, 66, 2025, false);
    const wSlut = Math.floor(p.born + 66);
    expect(fixed / nominal).toBeCloseTo(
      data.kpiAnnual.at(wSlut) / data.kpiAnnual.at(2025),
      9,
    );
  });

  it("assumes 8% nominal growth for years before the income index exists", () => {
    const p = profile({ born: 1930, par: 80, defAr: 80 });
    const early = wages(20, 1950, 15, 462_000, p, data, 0, 66, 2025, false);
    const later = wages(28, 1958, 15, 462_000, p, data, 0, 66, 2025, false);
    expect(early).toBeGreaterThan(0);
    expect(later / early).toBeCloseTo(1.08 ** 8, 6);
  });
});
