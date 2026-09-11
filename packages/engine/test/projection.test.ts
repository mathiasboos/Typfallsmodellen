import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import economicSeriesJson from "../../data/economic-series.json" with { type: "json" };
import { PROGNOSSTANDARD, projectEconomicData } from "../src/data/index.js";
import type { EconomicAssumptions, RawEconomicSeries, SeriesName } from "../src/data/index.js";

/**
 * The workbook's own cached values -- every year, projections included --
 * computed with the assumptions it shipped with. Reproducing this exactly is
 * what proves the projection rules in docs/PROJECTION-RULES.md were recovered
 * correctly.
 */
interface CachedFixture {
  readonly firstYear: number;
  readonly lastYear: number;
  readonly assumptions: { inflation: number; realGrowth: number; realReturn: number };
  readonly series: Readonly<Record<string, readonly (number | null)[]>>;
}

const cached = JSON.parse(
  readFileSync(
    fileURLToPath(new URL("../../../reference/fixtures/economic-series-cached.json", import.meta.url)),
    "utf8",
  ),
) as CachedFixture;

const raw = economicSeriesJson as unknown as RawEconomicSeries;

/** The assumptions the shipped workbook was saved with. */
const shipped: EconomicAssumptions = {
  ...PROGNOSSTANDARD,
  yearlyInflation: cached.assumptions.inflation,
  realGrowth: cached.assumptions.realGrowth,
  realReturn: cached.assumptions.realReturn,
};

const data = projectEconomicData(raw, shipped, {
  firstYear: cached.firstYear,
  lastYear: cached.lastYear,
});

/**
 * Series the workbook computes for every year, including years it also holds
 * actuals for. Their "actual" cells are the officially published figures, which
 * for the balancing years 2010-2018 differ from what the damping rule recomputes
 * -- so those are compared only from the year the sheet starts computing them.
 */
const COMPARE_FROM: Partial<Record<SeriesName, number>> = {
  balansindex: 2022,
};

/**
 * inkomstbasbelopp reads the *next* year's income index. The workbook's last row
 * has no next row, so it holds a 0 there; the engine computes the real value.
 * Compare everything except that one artifact.
 */
const COMPARE_UNTIL: Partial<Record<SeriesName, number>> = {
  inkomstbasbelopp: cached.lastYear - 1,
};

describe("economic projection", () => {
  const names = Object.keys(cached.series) as SeriesName[];

  it.each(names)("reproduces the workbook's cached %s", (name) => {
    const expected = cached.series[name]!;
    const series = data[name];
    expect(series, `no projected series named ${name}`).toBeDefined();

    const from = COMPARE_FROM[name] ?? raw.series[name]!.lastActualYear + 1;
    const until = COMPARE_UNTIL[name] ?? cached.lastYear;
    const mismatches: string[] = [];
    for (let year = from; year <= until; year += 1) {
      const want = expected[year - cached.firstYear];
      if (want === null || want === undefined) continue;
      const got = series.at(year);
      // The workbook's cached doubles carry the usual binary noise; compare to
      // twelve significant digits, which is the precision the data is stored at.
      const tolerance = Math.max(Math.abs(want) * 1e-12, 1e-12);
      if (Math.abs(got - want) > tolerance) {
        mismatches.push(`${year}: got ${got}, workbook has ${want}`);
      }
    }
    expect(mismatches.slice(0, 5).join("\n")).toBe("");
  });

  it("covers every series the engine knows about", () => {
    for (const name of Object.keys(raw.series)) {
      expect(names, `${name} missing from the fixture`).toContain(name);
    }
  });
});

describe("projection responds to assumptions", () => {
  it("grows prices with inflation, which the shipped 0% run cannot show", () => {
    const withInflation = projectEconomicData(raw, { ...shipped, yearlyInflation: 0.02 });
    const flat = data.kpiJune;
    const last = raw.series.kpiJune!.lastActualYear;

    expect(flat.at(last + 10)).toBe(flat.at(last));
    expect(withInflation.kpiJune.at(last + 10)).toBeGreaterThan(withInflation.kpiJune.at(last));
    // Ten years of 2% compounds to about 22%.
    expect(withInflation.kpiJune.at(last + 10) / withInflation.kpiJune.at(last)).toBeCloseTo(
      1.02 ** 10,
      2,
    );
  });

  it("grows the income index with real growth as well as inflation", () => {
    const grown = projectEconomicData(raw, { ...shipped, realGrowth: 0.016 });
    const last = raw.series.inkomstindex!.lastActualYear;
    expect(grown.inkomstindex.at(last + 5) / grown.inkomstindex.at(last)).toBeCloseTo(1.016 ** 5, 6);
  });

  it("carries inflation through to the prisbasbelopp, rounded to hundreds", () => {
    const withInflation = projectEconomicData(raw, { ...shipped, yearlyInflation: 0.02 });
    const last = raw.series.prisbasbelopp!.lastActualYear;
    const projected = withInflation.prisbasbelopp.at(last + 5);
    expect(projected).toBeGreaterThan(withInflation.prisbasbelopp.at(last));
    expect(projected % 100).toBe(0);
  });

  it("stops rounding when marginal is 1", () => {
    const unrounded = projectEconomicData(raw, {
      ...shipped,
      yearlyInflation: 0.02,
      marginal: 1,
    });
    const last = raw.series.prisbasbelopp!.lastActualYear;
    expect(unrounded.prisbasbelopp.at(last + 5) % 100).not.toBe(0);
  });

  it("uses inkomstindex for the index in force when latestIndexBasis is 2", () => {
    const incomeOnly = projectEconomicData(raw, { ...shipped, latestIndexBasis: 2 });
    for (const year of [2015, 2030]) {
      expect(incomeOnly.gallandeIndex.at(year)).toBe(incomeOnly.inkomstindex.at(year));
    }
  });

  it("subtracts fund fees from the return when returns are quoted gross", () => {
    const gross = projectEconomicData(raw, { ...shipped, returnsNetOfFees: false });
    const last = raw.series.avkastningPpm!.lastActualYear;
    expect(gross.avkastningPpm.at(last + 1)).toBeLessThan(data.avkastningPpm.at(last + 1));
    expect(gross.kvarEfterAvgiftPp.at(last + 1)).toBeLessThan(1);
  });
});
