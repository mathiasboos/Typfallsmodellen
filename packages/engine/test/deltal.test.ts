import { describe, expect, it } from "vitest";

import annuityTables from "../../data/annuity-tables.json" with { type: "json" };
import { loadDeathProbabilities } from "../src/data/nodeLoader.js";
import { defaultContext } from "../src/model/context.js";
import {
  DeltalTables,
  fnDeltalIp,
  fnDeltalIp2,
  fnDeltalPp,
  fnDeltalPp2,
} from "../src/pension/deltal.js";
import type { AnnuityTablesData } from "../src/pension/deltal.js";
import { calculateDeltal } from "../src/pension/mortality.js";

const published = annuityTables as unknown as AnnuityTablesData;
const deaths = loadDeathProbabilities();
const context = defaultContext();
const tables = new DeltalTables(published, deaths, context);

describe("DeltalTables sources", () => {
  it("takes cohorts before the override year from the published table", () => {
    // 1950 is below the 1958 default, so Nyckeltal governs.
    const row = published.incomePension.values[1950 - published.incomePension.firstCohort]!;
    const at65 = row[65 - published.incomePension.firstAge]!;
    expect(tables.incomePension(1950, 65)).toBe(at65);
  });

  it("takes cohorts from the override year from the model's own figures", () => {
    const own = calculateDeltal(1960, deaths);
    expect(tables.incomePension(1960, 66)).toBe(own.dtalip[0]![66]);
    expect(tables.premiumPension(1960, 66)).toBe(own.dtalpp[0]![66]);
  });

  it("keeps the published value past the mortality table's ages", () => {
    // The income-pension table stops at 82; ages beyond it are not overridden.
    expect(published.incomePension.lastAge).toBe(82);
    const row = published.incomePension.values[1960 - published.incomePension.firstCohort]!;
    expect(tables.incomePension(1960, 82)).not.toBe(row[82 - published.incomePension.firstAge]);
  });

  it("honours an override year of zero by using the published table throughout", () => {
    const noOverride = new DeltalTables(
      published,
      deaths,
      defaultContext({ deltalFromMortalityIp: 0 }),
    );
    const row = published.incomePension.values[1960 - published.incomePension.firstCohort]!;
    expect(noOverride.incomePension(1960, 66)).toBe(row[66 - published.incomePension.firstAge]);
  });
});

describe("fnDeltalIp", () => {
  it("returns the single divisor for cohorts born 1937 or earlier", () => {
    const old = published.bornBefore1938;
    expect(fnDeltalIp(1935, 65, tables)).toBe(old.incomePension[1935 - old.firstCohort]);
    expect(fnDeltalPp(1935, 70, tables)).toBe(old.premiumPension[1935 - old.firstCohort]);
  });

  it("ignores the age entirely for those cohorts", () => {
    expect(fnDeltalIp(1935, 61, tables)).toBe(fnDeltalIp(1935, 80, tables));
  });

  it("rejects cohorts the model has no figures for", () => {
    expect(() => fnDeltalIp(1929, 65, tables)).toThrow(RangeError);
    expect(() => fnDeltalIp(1960.5, 65, tables)).toThrow(RangeError);
  });

  it("returns the table value at a whole age", () => {
    expect(fnDeltalIp(1960, 66, tables)).toBeCloseTo(tables.incomePension(1960, 66), 2);
  });

  it("blends the neighbouring divisors at a part-year age", () => {
    const at66 = tables.incomePension(1960, 66);
    const at67 = tables.incomePension(1960, 67);
    const blended = fnDeltalIp(1960, 66.5, tables);
    expect(blended).toBeLessThan(at66);
    expect(blended).toBeGreaterThan(at67);
    expect(blended).toBeCloseTo((at66 + at67) / 2, 2);
  });

  it("weights by whole months, not the raw fraction", () => {
    const at66 = tables.incomePension(1960, 66);
    const at67 = tables.incomePension(1960, 67);
    // Three months in: a quarter of the way from one divisor to the next.
    expect(fnDeltalIp(1960, 66.25, tables)).toBeCloseTo(0.75 * at66 + 0.25 * at67, 2);
  });

  it("treats retirement on a whole birthday as mid-year when asked to", () => {
    expect(fnDeltalIp(1960, 66, tables, true)).toBe(fnDeltalIp(1960, 66.5, tables));
    expect(fnDeltalIp(1960, 66, tables, true)).not.toBe(fnDeltalIp(1960, 66, tables, false));
  });

  it("rounds .xx5 down rather than to even", () => {
    // aaDeltal.bas uses Int(x * 100 + 0.4999) / 100, which the VBA comments on
    // explicitly: a value ending in exactly .xx5 goes down.
    for (const age of [61, 64, 66, 70, 75]) {
      const value = fnDeltalIp(1960, age, tables);
      expect(Math.round(value * 100) / 100).toBe(value);
    }
  });
});

describe("fnDeltalIp2 and fnDeltalPp2", () => {
  it("blends in the year pension is first drawn", () => {
    expect(fnDeltalIp2(1960, 66, tables, 66, 66)).toBe(fnDeltalIp(1960, 66, tables, true));
  });

  it("does not blend in a year that is neither the first nor the final withdrawal", () => {
    expect(fnDeltalIp2(1960, 68, tables, 66, 70)).toBe(fnDeltalIp(1960, 68, tables, false));
  });

  it("stops the income-pension divisor moving past the final withdrawal age", () => {
    expect(fnDeltalIp2(1960, 75, tables, 66, 70)).toBe(fnDeltalIp2(1960, 71, tables, 66, 70));
  });

  it("does not clamp the premium-pension divisor, as the VBA omits that line", () => {
    expect(fnDeltalPp2(1960, 75, tables, 66, 70)).not.toBe(
      fnDeltalPp2(1960, 71, tables, 66, 70),
    );
  });
});

describe("delningstal shape across the model's range", () => {
  it("falls as retirement is deferred, for every cohort", () => {
    for (const cohort of [1940, 1950, 1958, 1960, 1970, 1980, 1990, 2000]) {
      for (let age = 62; age <= 80; age += 1) {
        const earlier = fnDeltalIp(cohort, age - 1, tables);
        const later = fnDeltalIp(cohort, age, tables);
        expect(later, `cohort ${cohort} age ${age}`).toBeLessThanOrEqual(earlier);
      }
    }
  });

  it("is positive and plausible everywhere the model can reach", () => {
    for (let cohort = 1938; cohort <= 2005; cohort += 1) {
      for (let age = 61; age <= 80; age += 1) {
        const value = fnDeltalIp(cohort, age, tables);
        expect(value, `cohort ${cohort} age ${age}`).toBeGreaterThan(5);
        expect(value, `cohort ${cohort} age ${age}`).toBeLessThan(30);
      }
    }
  });
});
