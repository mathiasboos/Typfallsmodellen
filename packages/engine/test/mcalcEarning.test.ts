import { describe, expect, it } from "vitest";

import annuityTables from "../../data/annuity-tables.json" with { type: "json" };
import brutto from "../../../reference/fixtures/brutto-cached.json" with { type: "json" };
import { loadDeathProbabilities } from "../src/data/nodeLoader.js";
import { defaultContext } from "../src/model/context.js";
import { Scheme, defaultInput } from "../src/model/input.js";
import type { TypfallInput } from "../src/model/input.js";
import type { ModelContext } from "../src/model/context.js";
import { earningPhase, prepareRun, ruleYears } from "../src/model/mcalc.js";
import { startsetup } from "../src/model/setup.js";
import { createRunState } from "../src/model/state.js";
import { DeltalTables } from "../src/pension/deltal.js";
import type { AnnuityTablesData } from "../src/pension/deltal.js";

const deaths = loadDeathProbabilities();
const published = annuityTables as unknown as AnnuityTablesData;

/** Builds a run and walks its earning phase over every age. */
function walk(input: TypfallInput, context: ModelContext) {
  const setup = startsetup(input, context);
  const state = createRunState(setup.vectors.startage, setup.vectors.slutage);
  const warnings = [...setup.warnings];
  const run = prepareRun(
    setup,
    input,
    context,
    state,
    new DeltalTables(published, deaths, context),
    warnings,
  );
  for (let age = run.v.startage; age <= run.v.slutage; age += 1) earningPhase(run, age);
  return { run, state, warnings };
}

const plain = (
  input: Parameters<typeof defaultInput>[0] = {},
  context: Parameters<typeof defaultContext>[0] = {},
) => walk(defaultInput(input), defaultContext(context));

describe("rule years", () => {
  it("follow the income year unless a rule year is pinned", () => {
    const { run } = plain();
    expect(ruleYears(run, 50)).toEqual({ utgyear: 2009, skyear: 2009 });
  });

  it("stop at the pinned year once it is passed", () => {
    const { run } = plain({}, { rulesFromUtg: 2020, rulesFromSkatt: 2018 });
    // 2009 is before both pinned years, so it keeps its own rules.
    expect(ruleYears(run, 50)).toEqual({ utgyear: 2009, skyear: 2009 });
    expect(ruleYears(run, 70)).toEqual({ utgyear: 2020, skyear: 2018 });
  });

  it("apply the pinned year to every age when rules is 1", () => {
    const { run } = plain({}, { rulesFromUtg: 2020, rules: 1 });
    expect(ruleYears(run, 30).utgyear).toBe(2020);
    expect(ruleYears(run, 90).utgyear).toBe(2020);
  });
});

describe("pension rights", () => {
  it("levies the contribution on the previous year's income", () => {
    const { run, state } = plain();
    // The assessment lags a year, so age 50's right is 16% of age 49's PGI --
    // an age whose contribution year is past 1998, where the rate settles.
    expect(state.ipRatt.get(50)).toBeCloseTo(0.16 * state.pgi.get(49), 0);
    expect(state.ppRatt.get(50)).toBeCloseTo(0.025 * state.pgi.get(49), 0);
    expect(state.gpRatt.get(50)).toBeCloseTo(0.185 * state.pgi.get(49), 0);
    expect(run.p.andelnya).toBe(1);
    // The higher transitional rate applies to the 1995-1998 contribution years.
    expect(state.ipRatt.get(40)).toBeCloseTo(0.165 * state.pgi.get(39), 0);
  });

  it("earns nothing before work starts", () => {
    const { state } = plain();
    expect(state.pgi.get(22)).toBe(0);
    expect(state.ipRatt.get(23)).toBe(0);
    expect(state.ipRatt.get(24)).toBeGreaterThan(0);
  });

  it("stops ATP points at 65 and keeps an age-ordered copy for SAF-LO", () => {
    const { state } = plain({ born: 1950 });
    expect(state.tpPoints.get(64)).toBeGreaterThan(0);
    expect(state.tpPoints.get(65)).toBe(0);
    for (let age = 15; age <= 64; age += 1) {
      expect(state.stpPoints.get(age)).toBe(state.tpPoints.get(age));
    }
  });

  it("counts the years that carry a pension right or an ATP point", () => {
    const { run, state } = plain();
    // Work from 23 to the riktålder, one year's lag at each end.
    expect(state.pgiYears).toBeGreaterThan(35);
    expect(state.pgiYears).toBeLessThanOrEqual(run.p.riktalder - 23 + 1);
  });
});

describe("pensionsgrundande belopp", () => {
  it("earns none without children, because the PGB sheet ships empty", () => {
    const { state } = plain();
    for (let age = 15; age <= 70; age += 1) expect(state.pgb.get(age)).toBe(0);
  });

  it("counts a qualifying year on income alone, despite its name", () => {
    // `If (PGB_(age) + pgi_(age)) > 0` -- so `pgbyears` counts every year with a
    // pension base, not years with a pensionsgrundande belopp. It is what
    // decides which of the three contribution branches a run takes, and a
    // normal working life puts it well past five.
    const { state } = plain();
    expect(state.pgbYears).toBeGreaterThanOrEqual(40);
  });

  it("earns childcare years for a child born while working", () => {
    const { state } = plain({}, { childBirthYears: [1990, 0, 0, 0] });
    const ageAtBirth = 1990 - 1959;
    // The credit runs for the birth year and the three years after.
    const credited = [0, 1, 2, 3].map((n) => state.pgb.get(ageAtBirth + n));
    expect(credited.filter((x) => x > 0).length).toBeGreaterThan(0);
    expect(state.pgbYears).toBeGreaterThan(0);
  });

  it("takes the PGB sheet's manual entries where they are given", () => {
    const { state } = plain({
      pgbManual: [{ age: 30, sa: 50_000, vpl: 0, studier: 0 }],
    });
    expect(state.pgb.get(30)).toBe(50_000);
    expect(state.pgb.get(31)).toBe(0);
  });

  it("caps income plus PGB at 7.5 inkomstbasbelopp", () => {
    const { run, state } = plain({ monthlySalary: 60_000 }, {
      childBirthYears: [0, 0, 0, 0],
      // A large manual entry, to push against the ceiling.
    });
    const withManual = walk(
      defaultInput({ monthlySalary: 60_000, pgbManual: [{ age: 40, sa: 5_000_000, vpl: 0, studier: 0 }] }),
      defaultContext(),
    );
    const ceiling = 7.5 * run.v.ibb.get(40);
    expect(withManual.state.pgi.get(40) + withManual.state.pgb.get(40)).toBeCloseTo(ceiling, 6);
    expect(state.pgb.get(40)).toBe(0);
  });
});

describe("the occupational premium", () => {
  it("is zero without an agreement", () => {
    const { state } = plain();
    for (let age = 23; age <= 65; age += 1) expect(state.tjpRatt.get(age)).toBe(0);
  });

  it("is paid from each agreement's own entry age", () => {
    const itp1 = plain({ born: 1985, scheme: Scheme.Itp1 }).state;
    expect(itp1.tjpRatt.get(24)).toBe(0);
    expect(itp1.tjpRatt.get(25)).toBeGreaterThan(0);
  });

  it("pays no ITPK premium before 1997", () => {
    const { state } = plain({ born: 1970, scheme: Scheme.Itp2 });
    expect(state.tjpRatt.get(1996 - 1970)).toBe(0);
    expect(state.tjpRatt.get(1998 - 1970)).toBeGreaterThan(0);
  });

  it("QUIRK: Mcalc's own 1976 guard on ITP 2 is dead", () => {
    // `tlITP2A` returns 0 below 1997 on its own, so the `If year_(age) > 1976`
    // around the call can never change an answer. Kept, and asserted so that a
    // future workbook moving one of the two years shows up here.
    const { state } = plain({ born: 1930, scheme: Scheme.Itp2 });
    for (let age = 15; age <= 66; age += 1) expect(state.tjpRatt.get(age)).toBe(0);
  });
});

describe("private saving", () => {
  it("puts nothing aside by default", () => {
    const { state } = plain();
    expect(state.ipsRatt).toBe(0);
  });

  it("saves a monthly amount from its start year until the occupational pension", () => {
    const input = defaultInput();
    const context = defaultContext({ ipsMonthly: 1000, ipsStart: 2000 });
    const setup = startsetup(input, context);
    const state = createRunState(setup.vectors.startage, setup.vectors.slutage);
    const run = prepareRun(setup, input, context, state, new DeltalTables(published, deaths, context));

    const at = (age: number) => {
      earningPhase(run, age);
      return state.ipsRatt;
    };
    expect(at(1999 - 1959)).toBe(0);
    expect(at(2000 - 1959)).toBe(12_000);
    // Nothing after the occupational pension starts, which follows PAR at 66.
    expect(at(66)).toBe(0);
  });

  it("reads a value at or below 1 as a share of income", () => {
    const input = defaultInput();
    const context = defaultContext({ ipsMonthly: 0.05, ipsStart: 2000 });
    const setup = startsetup(input, context);
    const state = createRunState(setup.vectors.startage, setup.vectors.slutage);
    const run = prepareRun(setup, input, context, state, new DeltalTables(published, deaths, context));
    const age = 2005 - 1959;
    earningPhase(run, age);
    expect(state.ipsRatt).toBeCloseTo(0.05 * setup.vectors.income.get(age), 9);
  });
});

describe("withdrawal share and divisors", () => {
  it("draws nothing before the pension starts", () => {
    const input = defaultInput();
    const context = defaultContext();
    const setup = startsetup(input, context);
    const state = createRunState(setup.vectors.startage, setup.vectors.slutage);
    const run = prepareRun(setup, input, context, state, new DeltalTables(published, deaths, context));
    for (let age = run.v.startage; age <= 65; age += 1) earningPhase(run, age);
    expect(state.uttagIp).toBe(0);
    expect(state.uttagPp).toBe(0);
    expect(state.dtalIp).toBe(0);
    expect(state.dtalPp).toBe(0);
  });

  it("draws the whole pension from the retirement year on", () => {
    const input = defaultInput();
    const context = defaultContext();
    const setup = startsetup(input, context);
    const state = createRunState(setup.vectors.startage, setup.vectors.slutage);
    const run = prepareRun(setup, input, context, state, new DeltalTables(published, deaths, context));
    for (let age = run.v.startage; age <= 70; age += 1) earningPhase(run, age);
    expect(state.uttagIp).toBe(1);
    expect(state.uttagPp).toBe(1);
    expect(state.dtalIp).toBeGreaterThan(10);
    expect(state.dtalPp).toBeGreaterThan(10);
  });
});

describe("Mcalc's own preamble", () => {
  it("gives a whole year of pension to a cohort born on 1 January", () => {
    const { run } = plain();
    expect(run.pmonth).toBe(12);
    expect(run.tmonth).toBe(12);
  });

  it("raises försäkringstid to the working life when it is set too low", () => {
    const { run, warnings } = plain({}, { insuranceYears: 20 });
    // Working 23 to 66 is 42 years, so the 40-year cap applies.
    expect(run.forstid).toBe(40);
    expect(warnings.at(-1)).toMatchObject({ field: "rng_Försäkringstid_vid_65", given: 20 });
  });

  it("leaves försäkringstid alone when the working life is shorter", () => {
    const { run, warnings } = plain({ startWorkAge: 50 }, { insuranceYears: 20 });
    expect(run.forstid).toBe(20);
    expect(warnings).toEqual([]);
  });

  it("QUIRK: kvoten is always 1, because its guard compares a year against 105", () => {
    // `If Iyear < slutage` -- slutage is the constant 105, so no calendar year
    // satisfies it and the earnings-indexed branches of gp, BTP and SBTP never
    // fire.
    expect(plain().run.kvoten).toBe(1);
    expect(plain({}, { boundaryYear: 2030 }).run.kvoten).toBe(1);
    expect(plain({ born: 1930 }, { boundaryYear: 1990 }).run.kvoten).toBe(1);
  });

  it("defaults the boundary year past the end of the run", () => {
    expect(plain().run.iyear).toBe(1959 + 105);
  });
});

describe("against the Brutto sheet", () => {
  /**
   * The Brutto sheet is a per-age trace of one typfall -- income above the
   * contribution ceiling every year -- with the VBA's own functions called as
   * worksheet UDFs. Driving the loop with its income vector checks the wiring
   * between them, which the per-function tests cannot: that PGI is taken on the
   * right year's income, that the contribution lags a year, and that the
   * ceiling is applied where the sheet applies it.
   *
   * It is one cohort at a fixed nominal income, so it exercises the loop, not
   * the wage profile that normally feeds it. The balances further down the
   * sheet need the closing-balance section, which is not in this commit.
   */
  const rows = brutto.rows as unknown as Record<string, number | null>[];
  const bornYear = rows[0]!.year! - rows[0]!.age!;

  const fixtureRun = () => {
    const input = defaultInput({
      born: bornYear,
      ownIncome: rows.map((r) => ({ age: r.age!, income: r.income!, wage: r.income! })),
    });
    const context = defaultContext({ referenceYear: brutto.referenceYear });
    return walk(input, context);
  };

  it("is one cohort born in 2000, earning above the ceiling", () => {
    expect(bornYear).toBe(2000);
    expect(brutto.shareOfNewSystem).toBe(1);
    expect(rows.length).toBe(brutto.rowCount);
  });

  it("reproduces the sheet's pension-qualifying income at every age", () => {
    const { state } = fixtureRun();
    let compared = 0;
    for (const row of rows) {
      if (row.pgi === null) continue;
      expect(state.pgi.get(row.age!), `pgi at age ${row.age}`).toBeCloseTo(row.pgi!, 6);
      compared += 1;
    }
    expect(compared).toBeGreaterThan(50);
  });

  it("earns no pensionsgrundande belopp, as the sheet has none", () => {
    const { state } = fixtureRun();
    for (const row of rows) {
      if (row.pgb === null) continue;
      expect(state.pgb.get(row.age!)).toBe(row.pgb);
    }
  });

  it("reproduces the three contributions at every age", () => {
    const { state } = fixtureRun();
    let compared = 0;
    for (const row of rows) {
      const age = row.age!;
      if (row.ipavgift === null) continue;
      expect(state.ipRatt.get(age), `IP right at ${age}`).toBeCloseTo(row.ipavgift!, 6);
      expect(state.ppRatt.get(age), `PP right at ${age}`).toBeCloseTo(row.ppavgift!, 6);
      expect(state.gpRatt.get(age), `GP right at ${age}`).toBeCloseTo(row.gpavgift!, 6);
      compared += 1;
    }
    expect(compared).toBeGreaterThan(40);
  });

  it("takes the contribution on the previous year's income, as the sheet does", () => {
    // Age 17's contribution is 16% of age 16's PGI, which is what fixes the
    // lag: the sheet's own column carries it, and the loop has to match.
    const at = (age: number) => rows.find((r) => r.age === age)!;
    expect(at(17).ipavgift).toBeCloseTo(0.16 * at(16).pgi!, 6);
    expect(at(17).ipavgift).not.toBeCloseTo(0.16 * at(17).pgi!, 6);
    const { state } = fixtureRun();
    expect(state.ipRatt.get(17)).toBeCloseTo(0.16 * state.pgi.get(16), 6);
  });
});
