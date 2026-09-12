import { describe, expect, it } from "vitest";

import cached from "../../../reference/fixtures/economic-series-cached.json" with { type: "json" };
import inheritance from "../../data/inheritance-gains.json" with { type: "json" };
import municipalTax from "../../data/municipal-tax.json" with { type: "json" };
import { defaultContext } from "../src/model/context.js";
import { defaultInput } from "../src/model/input.js";
import { startsetup } from "../src/model/setup.js";

/**
 * `startsetup` mostly re-reads series the projection already reproduces exactly,
 * so the tests that matter are about wiring: that each vector is fed by the
 * right series at the right age, that the guards around the inheritance-gain and
 * tax grids fire where the VBA's do, and that the validations correct the way
 * the workbook does when its dialogs are accepted.
 *
 * The economic vectors are checked against `economic-series-cached.json` -- the
 * values the workbook itself had cached under the shipped assumptions -- so an
 * off-by-one in the age-to-year mapping shows up against the original rather
 * than against the projection that feeds it.
 */

const run = (
  input: Parameters<typeof defaultInput>[0] = {},
  context: Parameters<typeof defaultContext>[0] = {},
) => startsetup(defaultInput(input), defaultContext(context));

/** A cached series value for a year, or undefined where the workbook had none. */
function cachedAt(series: string, year: number): number | undefined {
  const values = (cached.series as Record<string, (number | null)[]>)[series]!;
  return values[year - cached.firstYear] ?? undefined;
}

describe("the economic vectors", () => {
  it("map age to year as born + age", () => {
    const { vectors: v, profile: p } = run();
    expect(p.born).toBe(1959);
    for (let age = v.startage; age <= v.slutage; age += 1) {
      expect(v.year.get(age)).toBe(1959 + age);
    }
  });

  it("carry the workbook's own cached values from 1960 on", () => {
    const { vectors: v } = run();
    const pairs: [string, (age: number) => number][] = [
      ["kpiJune", (a) => v.kpiJune.get(a)],
      ["kpiAnnual", (a) => v.kpi.get(a)],
      ["prisbasbelopp", (a) => v.pbb.get(a)],
      ["inkomstbasbelopp", (a) => v.ibb.get(a)],
      ["forhojtPrisbasbelopp", (a) => v.fpb.get(a)],
      ["inkomstindex", (a) => v.iindex.get(a)],
      ["medelPgi", (a) => v.mpgi.get(a)],
    ];
    let compared = 0;
    for (const [series, read] of pairs) {
      for (let age = v.startage; age <= v.slutage; age += 1) {
        const year = v.year.get(age);
        if (year <= 1959) continue;
        const want = cachedAt(series, year);
        if (want === undefined) continue;
        expect(read(age), `${series} at ${year}`).toBeCloseTo(want, 6);
        compared += 1;
      }
    }
    // Guard against the loop quietly comparing nothing.
    expect(compared).toBeGreaterThan(500);
  });

  it("holds prices at their 1960 level before 1960", () => {
    // Born 1930, so the run reaches back to 1931.
    const { vectors: v } = run({ born: 1930 });
    expect(v.year.get(1)).toBe(1931);
    expect(v.kpi.get(1)).toBe(25.39);
    expect(v.kpiJune.get(1)).toBe(25.39);
    expect(v.pbb.get(1)).toBe(4200);
    expect(v.ibb.get(1)).toBe(4200);
    expect(v.fpb.get(1)).toBe(4200);
    // The return is assumed at 9%, and none of it reaches a pension.
    expect(v.yieldFactor.get(1)).toBe(1.09);
    // Income index runs back at 6% a year from 6.54 in 1960.
    expect(v.iindex.get(1)).toBeCloseTo(6.54 / 1.06 ** 29, 9);
  });

  it("never lets the index pensions follow exceed inkomstindex", () => {
    const { vectors: v } = run();
    for (let age = v.startage; age <= v.slutage; age += 1) {
      if (v.year.get(age) <= 1959) continue;
      expect(v.pindex.get(age)).toBeLessThanOrEqual(v.iindex.get(age));
      expect(v.pindex.get(age)).toBeGreaterThanOrEqual(1);
    }
  });
});

describe("the fund return", () => {
  it("follows the historical index by default", () => {
    const { vectors: v } = run();
    const year = v.year.get(50);
    expect(v.yieldFactor.get(50)).toBeCloseTo(1 + cachedAt("avkastningPpm", year)!, 9);
  });

  it("takes the AP7 series when the basis is 3", () => {
    const { vectors: v } = run({}, { returnBasis: 3 });
    const year = v.year.get(50);
    expect(v.yieldFactor.get(50)).toBeCloseTo(1 + cachedAt("avkastningAp7", year)!, 9);
  });

  it("uses the chosen real return, carried through inflation, when the basis is 1", () => {
    const { vectors: v } = run({ yearlyInflation: 0.02 }, { returnBasis: 1 });
    const ratio = v.kpi.get(50) / v.kpi.get(49);
    expect(v.yieldFactor.get(50)).toBeCloseTo(1.017 * ratio, 9);
    // The first age has no previous year to carry.
    expect(v.yieldFactor.get(v.startage)).toBeCloseTo(1.017, 9);
  });

  it("takes 0.2% off when the entered return is gross of fund fees", () => {
    const net = run({}, { returnBasis: 1 }).vectors.yieldFactor.get(50);
    const gross = run({}, { returnBasis: 1, returnsNetOfFees: false }).vectors.yieldFactor.get(50);
    expect(net - gross).toBeCloseTo(0.002, 9);
  });

  it("pays the Riksgälden rate only from 1996", () => {
    const { vectors: v } = run();
    const before = 1995 - 1959;
    expect(v.rgk.get(before)).toBe(1);
    expect(v.rgk.get(before + 1)).toBeCloseTo(1 + cachedAt("rantaRiksgalden", 1996)! / 100, 9);
  });
});

describe("fees and inheritance gains", () => {
  it("charges no management cost before the systems existed", () => {
    const { vectors: v } = run();
    expect(v.ipAvg.get(1959 - 1959 || 1)).toBe(1);
    expect(v.ppAvg.get(1999 - 1959)).toBe(1);
    expect(v.ppAvg.get(2000 - 1959)).toBeCloseTo(cachedAt("kvarEfterAvgiftPp", 2000)!, 9);
    // Occupational funds are assumed as efficient as the PPM platform.
    for (let age = v.startage; age <= v.slutage; age += 1) {
      expect(v.tpAvg.get(age)).toBe(v.ppAvg.get(age));
    }
  });

  it("pays no inheritance gain before age 17 or before 2000", () => {
    const { vectors: v } = run();
    expect(v.ipArv1.get(16)).toBe(1);
    expect(v.ipArv2.get(16)).toBe(1);
    expect(v.ipArv1.get(2000 - 1959 - 1)).toBe(1);
  });

  it("reads the income-pension grid at the VBA's own address", () => {
    const { vectors: v, profile: p } = run();
    const grid = inheritance.incomePension;
    const age = 50;
    const year = 1959 + age;
    expect(age).toBeLessThan(p.riktl);
    const want = grid.values[age - grid.firstAge]![year - grid.firstYear]!;
    expect(v.ipArv1.get(age)).toBeCloseTo(want, 12);
    // Below the riktålder the double factor is inert.
    expect(v.ipArv2.get(age)).toBe(1);
  });

  it("switches to the double factor from the riktålder, and the year before", () => {
    const { vectors: v, profile: p } = run();
    const grid = inheritance.incomePensionDouble;
    for (const age of [p.riktl - 1, p.riktl, p.riktl + 5]) {
      const year = 1959 + age;
      const want = grid.values[age - grid.firstAge]![year - grid.firstYear]!;
      expect(v.ipArv2.get(age), `age ${age}`).toBeCloseTo(want, 12);
    }
    // And the single factor stops at the riktålder.
    expect(v.ipArv1.get(p.riktl)).toBe(1);
  });

  it("QUIRK: the single factor's grid stops at 61, so 62 falls back to 1", () => {
    // The sheet only carries rows up to age 61 -- the double factor takes over
    // at the riktålder -- and the VBA's `If IP_arv1 = 0 Then IP_arv1 = 1` is
    // what covers the years where the riktålder is higher than that.
    const { vectors: v, profile: p } = run();
    expect(p.riktl).toBe(62);
    expect(v.ipArv1.get(60)).toBeGreaterThan(1);
    expect(v.ipArv1.get(62)).toBe(1);
  });

  it("pays no premium-pension gain before 2003, which is where its grid starts", () => {
    const { vectors: v } = run();
    expect(v.ppArv.get(2002 - 1959)).toBe(1);
    const grid = inheritance.premiumPension;
    const age = 2003 - 1959;
    const want = grid.values[age - grid.firstAge]![2003 - grid.firstYear]!;
    expect(v.ppArv.get(age)).toBeCloseTo(want, 12);
  });
});

describe("taxes", () => {
  it("takes the historical average rate when no rate is entered", () => {
    const { vectors: v } = run();
    const kommunalskatt = municipalTax.series.kommunalskatt.values;
    const year = 2010;
    const want = kommunalskatt[year - municipalTax.firstYear]! / 100;
    expect(v.komSkatt.get(year - 1959)).toBeCloseTo(want, 9);
  });

  it("uses an entered rate for every year, history included", () => {
    const { vectors: v } = run({}, { kommunalskatt: 0.32, begravningsavgift: 0.003 });
    for (const age of [30, 50, 66, 90]) {
      expect(v.komSkatt.get(age)).toBe(0.32);
      expect(v.begravavg.get(age)).toBe(0.003);
    }
  });

  it("charges no burial fee before 2000 in historical mode", () => {
    const { vectors: v } = run();
    expect(v.begravavg.get(1999 - 1959)).toBe(0);
    expect(v.begravavg.get(2000 - 1959)).toBeGreaterThan(0);
  });

  it("holds the second threshold open and the first flat before 1959", () => {
    const { vectors: v } = run({ born: 1930 });
    // The oldest cohort the workbook offers starts in 1931, so the `<= 1930`
    // branch of the VBA -- an assumed 8.44% -- is unreachable.
    expect(v.year.get(1)).toBe(1931);
    expect(v.komSkatt.get(1)).toBeCloseTo(
      municipalTax.series.kommunalskatt.values[1931 - municipalTax.firstYear]! / 100,
      9,
    );
    expect(v.taxLimit2.get(1)).toBe(10 ** 9);
    expect(v.taxLimit1.get(1958 - 1930)).toBe(11_000);
  });

  it("takes the thresholds from skiktgräns once the sheet has them", () => {
    const { vectors: v } = run();
    expect(v.taxLimit1.get(2020 - 1959)).toBeCloseTo(cachedAt("skiktgrans1", 2020)!, 6);
    expect(v.taxLimit2.get(2020 - 1959)).toBeCloseTo(cachedAt("skiktgrans2", 2020)!, 6);
  });

  it("pins every year to one year's tax rules when asked", () => {
    const { vectors: v } = run({}, { rulesFromSkatt: 2020, rules: 1 });
    const want = cachedAt("skiktgrans1", 2020)!;
    for (const age of [30, 50, 66, 90]) {
      expect(v.taxLimit1.get(age)).toBeCloseTo(want, 6);
    }
  });
});

describe("the income and wage vectors", () => {
  it("pays nothing before work starts and nothing after the pension is final", () => {
    const { vectors: v, profile: p } = run();
    expect(p.wStart).toBe(23);
    expect(v.income.get(22)).toBe(0);
    expect(v.income.get(23)).toBeGreaterThan(0);
    expect(v.income.get(65)).toBeGreaterThan(0);
    // Born on 1 January and retiring at 66 means no salary in the retirement year.
    expect(v.income.get(66)).toBe(0);
  });

  it("keeps income and wage equal unless the typfall stops working", () => {
    const { vectors: v } = run();
    for (let age = v.startage; age <= v.slutage; age += 1) {
      expect(v.wage.get(age)).toBe(v.income.get(age));
    }
  });

  it("zeroes the wage but not the income for a compensation recipient", () => {
    const { vectors: v } = run({}, { satagare: 2000 });
    const age = 2005 - 1959;
    expect(v.income.get(age)).toBeGreaterThan(0);
    expect(v.wage.get(age)).toBe(0);
    // Before that year the salary is untouched.
    expect(v.wage.get(1999 - 1959)).toBe(v.income.get(1999 - 1959));
  });

  it("applies a salary change from a given income year", () => {
    const plain = run().vectors.income;
    const halved = run({}, { tlSpecYear: 2000, tlSpecial: 0.5 }).vectors.income;
    expect(halved.get(1999 - 1959)).toBeCloseTo(plain.get(1999 - 1959), 9);
    expect(halved.get(2005 - 1959)).toBeCloseTo(plain.get(2005 - 1959) * 0.5, 9);
  });

  it("takes a hand-entered income vector in place of the wage profile", () => {
    const { vectors: v } = run({
      ownIncome: [{ age: 30, income: 500_000, wage: 450_000 }],
    });
    // An own vector forces the model to start at 15.
    expect(v.startage).toBe(15);
    expect(v.income.get(30)).toBe(500_000);
    expect(v.wage.get(30)).toBe(450_000);
    // Ages the vector does not mention stay at zero.
    expect(v.income.get(31)).toBe(0);
  });
});

describe("validation", () => {
  it("moves a retirement age below the riktålder up, and says so", () => {
    // Cohort 1959 retiring at 61 would be in 2020, when 62 was the floor.
    const { profile, warnings } = run({ retirementAge: 61 });
    expect(profile.par).toBe(62);
    expect(warnings).toHaveLength(1);
    expect(warnings[0]).toMatchObject({ field: "ParYear", given: 61, used: 62 });
    expect(warnings[0]!.message).toContain("62");
  });

  it("accepts a retirement age at the floor without comment", () => {
    const { profile, warnings } = run({ retirementAge: 62 });
    expect(profile.par).toBe(62);
    expect(warnings).toEqual([]);
  });

  it("lets the occupational pension follow the public one", () => {
    // TJP_PAR is a formula falling back to the retirement age, so 0 means follow.
    expect(run({ retirementAge: 70 }).profile.tjpPar).toBe(70);
    expect(run({}, { tjpPar: 62 }).profile.tjpPar).toBe(62);
  });

  it("refuses an occupational pension before 55", () => {
    const { profile, warnings } = run({}, { tjpPar: 50 });
    expect(profile.tjpPar).toBe(66);
    expect(warnings[0]).toMatchObject({ field: "tjp_par", given: 50, used: 66 });
  });

  it("treats a blank or too-early final withdrawal age as the retirement age", () => {
    expect(run().profile.defAr).toBe(66);
    expect(run({}, { defAr: 60 }).profile.defAr).toBe(66);
    expect(run({}, { defAr: 70 }).profile.defAr).toBe(70);
  });

  it("rejects an impossible final withdrawal age", () => {
    const { profile, warnings } = run({}, { defAr: 120 });
    expect(profile.defAr).toBe(66);
    expect(warnings[0]).toMatchObject({ field: "rng_def_ar", given: 120, used: 66 });
  });

  it("resolves the model year and the salary's reference age", () => {
    const { profile } = run();
    // w_ref is YEAR(NOW()) - 1, so the model year is one past it...
    expect(profile.modelYear).toBe(2026);
    // ...and w_time is the typfall's age in the reference year, the shipped 66.
    expect(profile.wTime).toBe(66);
  });

  it("resolves the cohort's share of the new system and its riktålder", () => {
    expect(run().profile.andelnya).toBe(1);
    // A twentieth per cohort from 1938, which puts 1945 at 11/20.
    expect(run({ born: 1945 }).profile.andelnya).toBeCloseTo(0.55, 9);
    // Both come from the Nyckeltal cohort table, not from riktage(year, typ):
    // the 1959 row reads 66 and 62. See pension/retirementAges.ts.
    expect(run().profile.riktalder).toBe(66);
    expect(run().profile.riktl).toBe(62);
    // And they do not move when the typfall retires later, as the year-keyed
    // function would have them do.
    expect(run({ retirementAge: 70 }).profile.riktl).toBe(62);
    expect(run({ retirementAge: 70 }).profile.riktalder).toBe(66);
    // A 1970 typfall is where the two rules part company most plainly: the
    // sheet says 65, riktage(1970 + 64, 0) would say 64.
    expect(run({ born: 1970, retirementAge: 67 }).profile.riktl).toBe(65);
  });
});
