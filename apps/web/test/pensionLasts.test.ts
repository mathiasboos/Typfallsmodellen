/**
 * "Hur länge räcker pensionen?"'s arithmetic, checked without a browser.
 *
 * `apps/web`'s test environment has no DOM, so the panel and its charts are
 * checked against the real built page by `tools/build/verify-offline.mjs`
 * instead (the same split `pensionAge.test.ts` makes). What is pinned here is
 * everything `pensionLasts.ts` computes before any of that: the payout periods
 * and budgets, how the input is held inside what the model allows, the engine
 * runs one payout period each, and how a pension is set against a budget -- year
 * by year, added up, before tax and after.
 *
 * The anchor numbers are model outputs, so a legitimate change to the engine
 * moves them; `npm run compare` (the golden cases) is what keeps the engine
 * honest, and these move with it on purpose.
 */
import { describe, expect, it } from "vitest";

import { Scheme, defaultContext, lowestPensionAge, run } from "@typfallsmodellen/engine";

import { loadDeathProbabilities } from "../src/deaths.js";
import { LAST_AGE, LIFESPAN_BOUNDS } from "../src/pensionAge.js";
import {
  BUDGETS,
  CHART_PAYOUTS,
  LIFELONG,
  PAYOUTS,
  PAYOUT_YEARS,
  TOLERANCE,
  assess,
  balanceText,
  budgetIdOf,
  budgetMatrix,
  computeLasts,
  coverageText,
  defaultLastsInput,
  evaluate,
  finalSalaryOf,
  lastsContext,
  lastsTypfall,
  matrixToCsv,
  matrixToXlsxRows,
  monthlySalaryForTarget,
  normaliseLasts,
  optionsToCsv,
  optionsToXlsxRows,
  payoutName,
  runPayouts,
  runsKey,
  savingYear,
  seriesToCsv,
  seriesToXlsxRows,
  slotOfPayout,
} from "../src/pensionLasts.js";
import type { Lasts, LastsInput } from "../src/pensionLasts.js";

const deaths = loadDeathProbabilities();
const context = defaultContext();
const base = defaultLastsInput();
const lasts = computeLasts(base, context, deaths);
const optionOf = (l: Lasts, payout: number) => l.options.find((o) => o.payout === payout)!;

/** A flat income by age, for the pure arithmetic. */
const flat = (from: number, to: number, value: (age: number) => number): number[] =>
  Array.from({ length: LAST_AGE + 1 }, (_, age) => (age >= from && age < to ? value(age) : 0));

describe("the payout periods", () => {
  it("are every whole year from 5 to 20, then lifelong", () => {
    expect(PAYOUT_YEARS).toEqual([5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20]);
    expect(PAYOUTS).toHaveLength(17);
    expect(PAYOUTS[0]).toBe(5);
    expect(PAYOUTS[16]).toBe(LIFELONG);
    expect(LIFELONG).toBe(0);
  });

  it("draw five of them, each with a colour of its own", () => {
    expect(CHART_PAYOUTS).toEqual([5, 10, 15, 20, LIFELONG]);
    expect(CHART_PAYOUTS.map((p) => slotOfPayout(p))).toEqual(["y5", "y10", "y15", "y20", "life"]);
    expect(slotOfPayout(7)).toBeUndefined();
  });

  it("are named in both languages", () => {
    expect(payoutName(10, "sv")).toBe("10 år");
    expect(payoutName(10, "en")).toBe("10 years");
    expect(payoutName(LIFELONG, "sv")).toBe("Livsvarigt");
    expect(payoutName(LIFELONG, "en")).toBe("Lifelong");
  });
});

describe("the budgets", () => {
  it("are minPension's three examples, the amounts Pensionsöversikten uses", () => {
    expect(BUDGETS.map((b) => [b.id, b.sv, b.amount])).toEqual([
      ["bas", "Baspensionär", 17_600],
      ["medel", "Medelpensionär", 18_400],
      ["plus", "Pluspensionär", 27_300],
    ]);
    expect(budgetIdOf(18_400)).toBe("medel");
    expect(budgetIdOf(18_401)).toBeUndefined();
  });
});

describe("the input", () => {
  it("opens on a typfall that has something to compare", () => {
    expect(base).toMatchObject({
      born: 1980,
      startWorkAge: 25,
      retirementAge: 68,
      lifespan: 90,
      scheme: Scheme.Itp1,
      budget: 18_400,
      capital: 0,
      monthlySaving: 0,
      afterTax: true,
    });
  });

  it("is held inside what the cohort may draw, with the lifespan after it", () => {
    expect(normaliseLasts({ ...base, retirementAge: 40 }).retirementAge).toBe(lowestPensionAge(1980));
    expect(normaliseLasts({ ...base, retirementAge: 99 }).retirementAge).toBe(80);
    expect(normaliseLasts({ ...base, retirementAge: 80, lifespan: 75 }).lifespan).toBe(81);
    expect(normaliseLasts({ ...base, lifespan: 40 }).lifespan).toBe(LIFESPAN_BOUNDS.min);
    expect(normaliseLasts({ ...base, lifespan: 140 }).lifespan).toBe(LIFESPAN_BOUNDS.max);
    expect(normaliseLasts(base)).toEqual(base);
  });

  it("makes a typfall of the person and nothing else", () => {
    expect(lastsTypfall({ ...base, monthlySalary: 41_000, realGrowth: 0.018 })).toMatchObject({
      born: 1980,
      startWorkAge: 25,
      monthlySalary: 41_000,
      retirementAge: 68,
      scheme: Scheme.Itp1,
      realGrowth: 0.018,
      married: false,
    });
  });

  it("keys the runs by the person, the pensions and the assumptions -- not the budget, the lifespan or the tax view", () => {
    const key = runsKey(base);
    expect(runsKey({ ...base, budget: 27_300, lifespan: 95, afterTax: false })).toBe(key);
    expect(runsKey({ ...base, monthlySalary: 40_000 })).not.toBe(key);
    expect(runsKey({ ...base, capital: 1 })).not.toBe(key);
    expect(runsKey({ ...base, retirementAge: 69 })).not.toBe(key);
    expect(runsKey({ ...base, realReturn: 0.03 })).not.toBe(key);
  });
});

describe("the context a payout is run with", () => {
  it("sets the payout for both pensions, starts the occupational pension with the public one, and uses fixed prices", () => {
    const shared = { ...context, tjpPar: 62, priceBasis: -1, tempTjpUttag: 3, tempIpsUttag: 4, privateSavingKind: 2 };
    const used = lastsContext(shared, { ...base, monthlySaving: 1_500, capital: 120_000 }, 12);
    expect(used).toMatchObject({
      tjpPar: 0,
      priceBasis: 1,
      tempTjpUttag: 12,
      tempIpsUttag: 12,
      privateSavingKind: 0,
      ipsMonthly: 1_500,
      ipsOpening: 120_000,
    });
    // Everything else is the shared context untouched.
    expect(used.referenceYear).toBe(shared.referenceYear);
    expect(used.marginal).toBe(shared.marginal);
    expect(used.kommunalskatt).toBe(shared.kommunalskatt);
  });

  it("reads a monthly saving of 1 or less as none, not as a share of income", () => {
    expect(lastsContext(context, { ...base, monthlySaving: 1 }, 10).ipsMonthly).toBe(0);
    expect(lastsContext(context, { ...base, monthlySaving: 0.5 }, 10).ipsMonthly).toBe(0);
    expect(lastsContext(context, { ...base, monthlySaving: 2 }, 10).ipsMonthly).toBe(2);
    expect(lastsContext(context, { ...base, capital: -5 }, 10).ipsOpening).toBe(0);
  });

  it("counts the saving from the model's today, or from the retirement year if that has passed", () => {
    expect(savingYear(base, context)).toBe(context.ipsStart);
    expect(savingYear({ ...base, born: 1959, retirementAge: 66 }, context)).toBe(2025);
    expect(savingYear(base, { ...context, ipsStart: 2040 })).toBe(2040);
  });

  it("never counts it from so early that the run cannot book the capital", () => {
    // Booked at the end of the year before, in a year the run covers: from the model's first age
    // (1) or the age work starts, whichever is earlier. 1990 + 1 + 1 is the earliest for a 1990 cohort.
    const early = { ...context, ipsStart: 1950 };
    expect(savingYear({ ...base, born: 1990, retirementAge: 68 }, early)).toBe(1992);
    expect(savingYear({ ...base, born: 1990, startWorkAge: 15, retirementAge: 68 }, { ...early, modelStartAge: 15 })).toBe(2006);
    expect(savingYear({ ...base, born: 1990, retirementAge: 68 }, { ...context, ipsStart: 1992 })).toBe(1992);
    expect(savingYear({ ...base, born: 1990, retirementAge: 68 }, { ...context, ipsStart: 1991 })).toBe(1992);
  });

  it("still pays out the capital when Avancerat's saving year is from before the person's time", () => {
    const young = { ...base, born: 1990, retirementAge: 68, lifespan: 90, capital: 300_000 };
    const withCapital = computeLasts(young, { ...context, ipsStart: 1950 }, deaths);
    const without = computeLasts({ ...young, capital: 0 }, { ...context, ipsStart: 1950 }, deaths);
    const total = (l: Lasts) => optionOf(l, 10).run.individual.reduce((sum, v) => sum + v, 0);
    expect(withCapital.runs.savingYear).toBe(1992);
    expect(total(without)).toBe(0);
    expect(total(withCapital)).toBeGreaterThan(300_000);
  });
});

describe("the runs", () => {
  const runs = lasts.runs;
  const R = base.retirementAge;

  it("are one per payout period, in order, and raise no warning", () => {
    expect(runs.runs.map((r) => r.payout)).toEqual([...PAYOUTS]);
    expect(runs.warnings).toEqual([]);
    expect(runs.referenceYear).toBe(context.referenceYear);
  });

  it("pay the same public pension whatever the payout", () => {
    const first = runs.runs[0]!.publicPension;
    for (const r of runs.runs) expect(r.publicPension).toEqual(first);
    expect(first[R - 1]).toBe(0);
    expect(first[R]).toBeGreaterThan(0);
  });

  it("pay the occupational pension for the payout's years, and not a year longer", () => {
    for (const payout of [5, 10, 15, 20]) {
      const r = runs.runs.find((x) => x.payout === payout)!;
      expect(r.occupational[R - 1]).toBe(0);
      expect(r.occupational[R]).toBeGreaterThan(0);
      expect(r.occupational[R + payout - 1], `last year of ${payout}`).toBeGreaterThan(0);
      expect(r.occupational[R + payout], `year after ${payout}`).toBe(0);
    }
    const life = runs.runs.find((x) => x.payout === LIFELONG)!;
    expect(life.occupational[LAST_AGE - 1]).toBeGreaterThan(0);
  });

  it("pay more a year the shorter the payout", () => {
    const first = PAYOUTS.map((p) => runs.runs.find((x) => x.payout === p)!.occupational[R]!);
    // 5, 6, ... 20, lifelong: strictly falling.
    for (let i = 1; i < first.length; i += 1) expect(first[i]!, `payout ${PAYOUTS[i]}`).toBeLessThan(first[i - 1]!);
    // And the figures the example in the file comment reads.
    expect(first[0]! / 12).toBeCloseTo(31_781, -2);
    expect(first[first.length - 1]! / 12).toBeCloseTo(8_207, -2);
  });

  it("have gross as the sum of its parts, and net below it", () => {
    for (const r of runs.runs) {
      for (let age = R; age <= LAST_AGE; age += 1) {
        expect(r.gross[age]).toBeCloseTo(r.publicPension[age]! + r.occupational[age]! + r.individual[age]!, 6);
        expect(r.net[age]!, `${r.payout} at ${age}`).toBeLessThan(r.gross[age]!);
        expect(r.net[age]!).toBeGreaterThan(0);
      }
    }
  });

  it("match a plain run of the model for the lifelong payout", () => {
    const plain = run(lastsTypfall(base), lastsContext(context, base, LIFELONG), { deaths });
    const life = runs.runs.find((x) => x.payout === LIFELONG)!;
    for (const row of plain.rows.filter((x) => x.age > R && x.age <= LAST_AGE)) {
      expect(life.net[row.age], `net at ${row.age}`).toBe(row.netto);
      expect(life.gross[row.age], `gross at ${row.age}`).toBe(row.brutto);
    }
  });

  it("have no individual pension without capital or saving, and pay it over the payout's years with them", () => {
    for (const r of runs.runs) expect(r.individual.every((v) => v === 0)).toBe(true);

    const own = runPayouts(normaliseLasts({ ...base, capital: 300_000, monthlySaving: 2_000 }), context, deaths);
    const ten = own.runs.find((x) => x.payout === 10)!;
    expect(ten.individual[R - 1]).toBe(0);
    expect(ten.individual[R]).toBeGreaterThan(0);
    expect(ten.individual[R + 9]).toBeGreaterThan(0);
    expect(ten.individual[R + 10]).toBe(0);
    const life = own.runs.find((x) => x.payout === LIFELONG)!;
    expect(life.individual[R]).toBeGreaterThan(0);
    // A shorter payout of the same pot is paid out faster.
    expect(ten.individual[R]!).toBeGreaterThan(life.individual[R]!);
  });

  it("count capital and saving separately: each adds to the pension", () => {
    const at = (patch: Partial<LastsInput>) =>
      runPayouts(normaliseLasts({ ...base, ...patch }), context, deaths).runs.find((x) => x.payout === 15)!.individual[R]!;
    const capitalOnly = at({ capital: 300_000 });
    const savingOnly = at({ monthlySaving: 2_000 });
    const both = at({ capital: 300_000, monthlySaving: 2_000 });
    expect(capitalOnly).toBeGreaterThan(0);
    expect(savingOnly).toBeGreaterThan(0);
    expect(both).toBeCloseTo(capitalOnly + savingOnly, -2);
  });

  it("leave out the shared private saving and the occupational start age Avancerat may have set", () => {
    const shared = { ...context, tjpPar: 62, ipsMonthly: 5_000, tempTjpUttag: 3, priceBasis: -1 };
    const overridden = runPayouts(base, shared, deaths);
    expect(overridden.runs.map((r) => r.net)).toEqual(runs.runs.map((r) => r.net));
  });
});

describe("setting a pension against a budget", () => {
  // Round thousands of kronor a year, so that the half-krona-a-month tolerance never decides anything here.
  const k = 1000;
  const from = 65;
  const to = 75;
  const need = (value: number) => () => value;

  it("covers the years the income reaches the need, and says when it stops", () => {
    // 30 a year for five years, then 10; the need is 20.
    const income = flat(from, to, (age) => (age < 70 ? 30 * k : 10 * k));
    const a = assess(5, income, need(20 * k), from, to);
    expect(a.covered).toEqual([true, true, true, true, true, false, false, false, false, false]);
    expect(a.coveredYears).toBe(5);
    expect(a.stretches).toEqual([{ from: 65, to: 70 }]);
    expect(a.firstCovered).toBe(65);
    expect(a.lastCovered).toBe(69);
    expect(a.runsOutAt).toBe(70);
    expect(a.lasts).toBe(false);
    expect(a.surplus).toBe(50 * k);
    expect(a.shortfall).toBe(50 * k);
    expect(coverageText(a, "sv")).toBe("Till 70 år");
    expect(coverageText(a, "en")).toBe("To 70");
  });

  it("lasts when every year is covered, and has no end to give", () => {
    const a = assess(LIFELONG, flat(from, to, () => 25 * k), need(20 * k), from, to);
    expect(a.lasts).toBe(true);
    expect(a.runsOutAt).toBeNull();
    expect(a.shortfall).toBe(0);
    expect(coverageText(a, "sv")).toBe("Hela vägen till 75 år");
  });

  it("does not last when no year is covered", () => {
    const a = assess(LIFELONG, flat(from, to, () => 15 * k), need(20 * k), from, to);
    expect(a.lasts).toBe(false);
    expect(a.coveredYears).toBe(0);
    expect(a.firstCovered).toBeNull();
    expect(a.lastCovered).toBeNull();
    expect(a.runsOutAt).toBeNull();
    expect(a.stretches).toEqual([]);
    expect(coverageText(a, "sv")).toBe("Räcker inte");
  });

  it("reports a first year that falls short as a stretch that begins later", () => {
    // The model taxes the first year harder: 17, then 25 for three years, then 10.
    const income = flat(from, to, (age) => (age === 65 ? 17 * k : age < 69 ? 25 * k : 10 * k));
    const a = assess(5, income, need(20 * k), from, to);
    expect(a.stretches).toEqual([{ from: 66, to: 69 }]);
    expect(a.firstCovered).toBe(66);
    expect(a.runsOutAt).toBe(69);
    expect(a.coveredYears).toBe(3);
    expect(coverageText(a, "sv")).toBe("66–69 år");

    const goes = assess(LIFELONG, flat(from, to, (age) => (age === 65 ? 17 * k : 25 * k)), need(20 * k), from, to);
    expect(goes.runsOutAt).toBeNull();
    expect(goes.lasts).toBe(false);
    expect(coverageText(goes, "sv")).toBe("Från 66 år");
  });

  it("counts years when the covered ones are not one stretch", () => {
    const income = flat(from, to, (age) => (age % 3 === 0 ? 10 * k : 25 * k));
    const a = assess(LIFELONG, income, need(20 * k), from, to);
    expect(a.stretches.length).toBeGreaterThan(1);
    expect(coverageText(a, "sv")).toBe(`${a.coveredYears} av 10 år`);
  });

  it("forgives half a krona a month", () => {
    const yearlyNeed = 12 * 18_400;
    const near = assess(LIFELONG, flat(from, to, () => yearlyNeed - TOLERANCE), need(yearlyNeed), from, to);
    expect(near.lasts).toBe(true);
    const short = assess(LIFELONG, flat(from, to, () => yearlyNeed - TOLERANCE - 1), need(yearlyNeed), from, to);
    expect(short.lasts).toBe(false);
  });

  describe("the running total", () => {
    const income = flat(from, 80, (age) => (age < 70 ? 30 * k : 10 * k));

    it("adds the years' differences up, from the retirement age", () => {
      const a = assess(5, income, need(20 * k), from, 80);
      expect(a.balance.slice(0, 5)).toEqual([10, 20, 30, 40, 50].map((v) => v * k));
      expect(a.balance.slice(5)).toEqual([40, 30, 20, 10, 0, -10, -20, -30, -40, -50].map((v) => v * k));
    });

    it("turns negative inside a year, to a tenth", () => {
      // A need of 22: +8 for five years, then -12 a year. 40, 28, 16, 4 are left at the ends of 69..72, -8 at 73's.
      const a = assess(5, income, need(22 * k), from, 80);
      expect(a.balance.slice(4, 9)).toEqual([40, 28, 16, 4, -8].map((v) => v * k));
      // The 4 is gone a third of the way into the year at 73 (12 a year): 73.33.
      expect(a.balanceRunsOutAt).toBeCloseTo(73 + 4 / 12, 9);
    });

    it("lasts when it never goes below zero, and is gone at once when it never gets there", () => {
      const exact = assess(5, income, need(20 * k), from, 75);
      expect(exact.balanceRunsOutAt).toBeNull();
      expect(exact.balance[9]).toBe(0);
      const none = assess(5, flat(from, 80, () => 10 * k), need(20 * k), from, 80);
      expect(none.balanceRunsOutAt).toBe(from);
      expect(balanceText(none, "sv")).toBe("Räcker inte");
    });

    it("is not gone in year one because year one fell a little short", () => {
      const dip = flat(from, 80, (age) => (age === 65 ? 17 * k : age < 70 ? 30 * k : 10 * k));
      const a = assess(5, dip, need(20 * k), from, 80);
      expect(a.balance[0]).toBe(-3 * k);
      // -3, 7, 17, 27, 37 at the ends of 65..69, then 27, 17, 7, -3 at the ends of 70..73: gone in 73.
      expect(a.balanceRunsOutAt!).toBeGreaterThan(73);
      expect(a.balanceRunsOutAt!).toBeLessThan(74);
      expect(balanceText(a, "sv")).toMatch(/^Till 73,\d år$/);
      expect(balanceText(a, "en")).toMatch(/^To 73\.\d$/);
      expect(balanceText(assess(LIFELONG, flat(from, to, () => 25 * k), need(20 * k), from, to), "sv")).toBe("Hela vägen till 75 år");
    });
  });

  describe("the phases", () => {
    const income = flat(from, to, (age) => (age < 70 ? 30 * k : 10 * k));

    it("split the years at the end of the payout", () => {
      const a = assess(5, income, need(20 * k), from, to);
      expect(a.during).toMatchObject({ from: 65, to: 70, years: 5, covered: true });
      expect(a.after).toMatchObject({ from: 70, to: 75, years: 5, covered: false });
      // Kronor a year to a month.
      expect(a.during!.income).toBeCloseTo((30 * k) / 12, 9);
      expect(a.during!.difference).toBeCloseTo((10 * k) / 12, 9);
      expect(a.after!.difference).toBeCloseTo((-10 * k) / 12, 9);
      expect(a.after!.lowest).toBeCloseTo((10 * k) / 12, 9);
    });

    it("are one phase for a lifelong payout and for one that outlasts the lifespan", () => {
      const life = assess(LIFELONG, income, need(20 * k), from, to);
      expect(life.during).toMatchObject({ from: 65, to: 75, years: 10 });
      expect(life.after).toBeNull();
      const long = assess(20, income, need(20 * k), from, to);
      expect(long.during).toMatchObject({ from: 65, to: 75, years: 10 });
      expect(long.after).toBeNull();
    });
  });
});

describe("the default person", () => {
  const life = optionOf(lasts, LIFELONG);

  it("is assessed from the retirement age to the lifespan", () => {
    expect(lasts.from).toBe(68);
    expect(lasts.to).toBe(90);
    expect(lasts.options).toHaveLength(17);
    expect(life.view.income).toHaveLength(22);
    expect(life.view.during!.years).toBe(22);
    expect(lasts.flexible).toBe(true);
  });

  it("can live on a lifelong payout on an average pensioner's budget, all the way", () => {
    expect(life.view.lasts).toBe(true);
    expect(life.view.coveredYears).toBe(22);
    expect(life.view.shortfall).toBe(0);
    expect(lasts.best).toBe(life);
    expect(life.view.during!.income).toBeCloseTo(22_289, -2);
  });

  it("is short of that budget on the public pension alone", () => {
    expect(lasts.publicOnly.lasts).toBe(false);
    expect(lasts.publicOnly.coveredYears).toBe(0);
    expect(lasts.publicOnly.during!.income).toBeCloseTo(16_094, -2);
  });

  it("is paid more a month, and for fewer years, the shorter the payout", () => {
    const monthly = PAYOUT_YEARS.map((p) => optionOf(lasts, p).view.during!.income);
    for (let i = 1; i < monthly.length; i += 1) expect(monthly[i]!).toBeLessThan(monthly[i - 1]!);
    expect(monthly[monthly.length - 1]!).toBeGreaterThan(life.view.during!.income);
    for (const payout of PAYOUT_YEARS) {
      const a = optionOf(lasts, payout).view;
      // 68 plus the payout is the age the payout ends at, and with it the budget.
      expect(a.runsOutAt, `${payout}`).toBe(68 + payout);
      expect(a.coveredYears).toBe(payout);
      // What is left once the payout ends is the public pension, taxed on its own.
      const alone: number[] = [];
      for (let age = 68 + payout; age < 90; age += 1) alone.push(lasts.runs.netAt(age, lasts.runs.runs[0]!.publicPension[age]!) / 12);
      expect(a.after!.income, `${payout}`).toBeCloseTo(alone.reduce((x, y) => x + y, 0) / alone.length, 0);
    }
  });

  it("saves the early surplus into a payout that outlasts the budget's gap", () => {
    // 5 years of +20 600 a month, then -2 400: the surplus covers the rest to 90.
    expect(optionOf(lasts, 5).net.balanceRunsOutAt).toBeNull();
    expect(optionOf(lasts, 5).view.shortfall).toBeGreaterThan(0);
  });

  it("has a budget before tax that is higher than after, and higher still in the years taxed harder", () => {
    const gross = computeLasts({ ...base, afterTax: false }, context, deaths);
    expect(gross.budgetGross.typical).toBeGreaterThan(base.budget);
    expect(gross.budgetGross.typical).toBeCloseTo(22_204, -2);
    expect(gross.budgetGross.first).toBeGreaterThan(gross.budgetGross.last);
    expect(gross.budgetGross.first).toBeCloseTo(25_946, -2);
    expect(gross.budgetMonthly).toHaveLength(22);
    expect(gross.budgetMonthly[0]).toBe(gross.budgetGross.first);
    // After tax the budget is the budget.
    expect(lasts.budgetMonthly.every((v) => v === base.budget)).toBe(true);
  });

  it("gives back the budget after tax when the budget before tax is taxed", () => {
    const gross = computeLasts({ ...base, afterTax: false }, context, deaths);
    for (const age of [68, 69, 75, 89]) {
      const monthly = gross.budgetMonthly[age - gross.from]!;
      expect(gross.runs.netAt(age, monthly * 12) / 12, `age ${age}`).toBeCloseTo(base.budget, 0);
    }
  });
});

describe("before tax and after", () => {
  const net = lasts;
  const gross = computeLasts({ ...base, afterTax: false }, context, deaths);

  it("cover the same years, since tax never leaves less for more", () => {
    for (const budget of [BUDGETS[0].amount, BUDGETS[1].amount, BUDGETS[2].amount, 12_000, 21_000]) {
      const a = computeLasts({ ...base, budget }, context, deaths);
      const b = computeLasts({ ...base, budget, afterTax: false }, context, deaths);
      a.options.forEach((option, i) => {
        const other = b.options[i]!.view;
        expect(other.covered, `${option.payout} at ${budget}`).toEqual(option.view.covered);
        expect(other.stretches).toEqual(option.view.stretches);
        expect(other.runsOutAt).toBe(option.view.runsOutAt);
      });
      expect(b.best?.payout).toBe(a.best?.payout);
    }
  });

  it("show a larger pension and a larger budget before tax", () => {
    net.options.forEach((option, i) => {
      expect(gross.options[i]!.view.during!.income).toBeGreaterThan(option.view.during!.income);
      expect(gross.options[i]!.view.during!.need).toBeGreaterThan(option.view.during!.need);
    });
  });

  it("read the running total after tax in both views", () => {
    gross.options.forEach((option, i) => {
      expect(option.net.balance).toEqual(net.options[i]!.net.balance);
      expect(option.net.balanceRunsOutAt).toBe(net.options[i]!.net.balanceRunsOutAt);
    });
    expect(net.options[0]!.view).toBe(net.options[0]!.net);
    expect(gross.options[0]!.view).not.toBe(gross.options[0]!.net);
  });

  it("leave the engine runs alone: the same runs serve both", () => {
    const runs = runPayouts(base, context, deaths);
    const a = evaluate(runs, base);
    const b = evaluate(runs, { ...base, afterTax: false, budget: 27_300, lifespan: 95 });
    expect(a.runs).toBe(b.runs);
    expect(b.to).toBe(95);
    expect(b.options[0]!.view.income).toHaveLength(27);
  });
});

describe("other budgets", () => {
  const plus = computeLasts({ ...base, budget: BUDGETS[2].amount }, context, deaths);
  const bas = computeLasts({ ...base, budget: BUDGETS[0].amount }, context, deaths);

  it("last all the way on a basic pensioner's, as on an average one's", () => {
    expect(optionOf(bas, LIFELONG).view.lasts).toBe(true);
    expect(bas.best).toBe(optionOf(bas, LIFELONG));
  });

  it("cannot be met for life by a plus pensioner's, and are met longest by a payout in between", () => {
    expect(optionOf(plus, LIFELONG).view.lasts).toBe(false);
    expect(optionOf(plus, LIFELONG).view.coveredYears).toBe(0);
    expect(plus.best).not.toBeNull();
    expect(plus.best!.payout).toBeGreaterThan(5);
    expect(plus.best!.payout).toBeLessThan(15);
    for (const option of plus.options) {
      expect(option.view.coveredYears).toBeLessThanOrEqual(plus.best!.view.coveredYears);
    }
    // The model taxes the first year harder, so the best payout falls short in it alone.
    expect(plus.best!.view.firstCovered).toBe(69);
  });

  it("have no best payout when no payout meets them in any year", () => {
    const none = computeLasts({ ...base, scheme: Scheme.None }, context, deaths);
    expect(none.best).toBeNull();
    expect(none.flexible).toBe(false);
    for (const option of none.options) expect(option.view.coveredYears).toBe(0);
  });

  it("make a matrix: more years the smaller the budget, and the most a budget can be", () => {
    const matrix = budgetMatrix(lasts, [BUDGETS[0].amount, BUDGETS[1].amount, BUDGETS[2].amount]);
    expect(matrix).toHaveLength(17);
    for (const row of matrix) {
      const years = row.cells.map((c) => c.coveredYears);
      expect(years[0]!, `payout ${row.payout}`).toBeGreaterThanOrEqual(years[1]!);
      expect(years[1]!).toBeGreaterThanOrEqual(years[2]!);
      // A budget at or below `highest` is met in every year of the payout.
      const met = assess(row.payout, optionOf(lasts, row.payout).run.net, () => row.highest * 12, lasts.from, lasts.to);
      expect(met.during!.covered, `payout ${row.payout}`).toBe(true);
      const over = assess(row.payout, optionOf(lasts, row.payout).run.net, () => (row.highest + 1) * 12, lasts.from, lasts.to);
      expect(over.during!.covered, `payout ${row.payout} over`).toBe(false);
    }
    const lifelong = matrix[16]!;
    expect(lifelong.cells[1]!.lasts).toBe(true);
    expect(lifelong.cells[2]!.lasts).toBe(false);
    // A shorter payout pays more a month, so a higher budget can be met during it.
    expect(matrix[0]!.highest).toBeGreaterThan(matrix[16]!.highest);
  });

  it("agree with the assessments of the person's own budget", () => {
    const matrix = budgetMatrix(lasts, [base.budget]);
    matrix.forEach((row, i) => {
      expect(row.cells[0]!.covered).toEqual(lasts.options[i]!.view.covered);
      expect(row.cells[0]!.balanceRunsOutAt).toBe(lasts.options[i]!.net.balanceRunsOutAt);
    });
  });
});

describe("the individual occupational pension", () => {
  const own = computeLasts({ ...base, capital: 300_000, monthlySaving: 2_000 }, context, deaths);

  it("adds to every payout and is paid for the same years", () => {
    own.options.forEach((option, i) => {
      expect(option.view.during!.income, `payout ${option.payout}`).toBeGreaterThan(lasts.options[i]!.view.during!.income);
    });
    expect(own.flexible).toBe(true);
  });

  it("counts for a person with no agreement", () => {
    const only = computeLasts({ ...base, scheme: Scheme.None, capital: 300_000, monthlySaving: 2_000 }, context, deaths);
    expect(only.flexible).toBe(true);
    expect(optionOf(only, 10).run.occupational.every((v) => v === 0)).toBe(true);
    expect(optionOf(only, 10).run.individual[68]).toBeGreaterThan(0);
    expect(optionOf(only, 5).view.during!.income).toBeGreaterThan(optionOf(only, 20).view.during!.income);
  });

  it("is split out before tax and only there", () => {
    const rows = optionsToCsv(computeLasts({ ...base, capital: 300_000, afterTax: false }, context, deaths), "sv").split("\r\n");
    expect(rows[0]!.split(";")).toContain("varav individuell tjänstepension");
    const net = optionsToCsv(own, "sv").split("\r\n");
    expect(net[0]!.split(";")).not.toContain("varav individuell tjänstepension");
  });
});

describe("a final salary", () => {
  it("is found for the monthly salary it comes from", () => {
    const final = finalSalaryOf(base, context, deaths);
    expect(final).toBeCloseTo(39_872, -1);
    expect(lasts.runs.finalSalary).toBe(final);
    const found = monthlySalaryForTarget(base, final, context, deaths);
    expect(Math.abs(found - base.monthlySalary)).toBeLessThanOrEqual(1);
  });

  it("lands on the salary asked for", () => {
    const monthly = monthlySalaryForTarget(base, 55_000, context, deaths);
    expect(monthly).toBeGreaterThan(base.monthlySalary);
    expect(finalSalaryOf({ ...base, monthlySalary: monthly }, context, deaths)).toBeCloseTo(55_000, -1);
    expect(monthlySalaryForTarget(base, 0, context, deaths)).toBe(0);
  });
});

describe("the exports", () => {
  it("write one line per payout, a header first, in the language's own delimiter", () => {
    const sv = optionsToCsv(lasts, "sv").split("\r\n");
    expect(sv).toHaveLength(18);
    expect(sv[0]!.split(";")).toHaveLength(sv[1]!.split(";").length);
    expect(sv[1]!.startsWith("5 år;")).toBe(true);
    expect(sv[17]!.startsWith("Livsvarigt;")).toBe(true);
    const en = optionsToCsv(lasts, "en").split("\r\n");
    expect(en[0]!.split(",").length).toBeGreaterThan(5);
    expect(en[17]!.startsWith("Lifelong,")).toBe(true);
  });

  it("write plain integer kronor, and leave out what a lifelong payout has no 'after' for", () => {
    const life = optionsToCsv(lasts, "sv").split("\r\n")[17]!.split(";");
    expect(life[2]).toBe("");
    expect(life[1]).toMatch(/^\d+$/);
    expect(optionsToXlsxRows(lasts, "sv")).toHaveLength(18);
  });

  it("write a series with a row for every age from retirement and a column for every payout", () => {
    const rows = seriesToCsv(lasts, "sv").split("\r\n");
    expect(rows).toHaveLength(1 + 22);
    expect(rows[0]!.split(";")).toHaveLength(1 + 1 + 17 + 17);
    expect(rows[1]!.startsWith("68;")).toBe(true);
    expect(rows[22]!.startsWith("89;")).toBe(true);
    expect(seriesToXlsxRows(lasts, "en")).toHaveLength(23);
  });

  it("write the matrix with one column per budget", () => {
    const budgets = BUDGETS.map((b) => ({ label: b.sv, amount: b.amount }));
    const matrix = budgetMatrix(lasts, budgets.map((b) => b.amount));
    const rows = matrixToCsv(matrix, budgets, "sv").split("\r\n");
    expect(rows).toHaveLength(18);
    expect(rows[0]!.split(";")).toHaveLength(1 + 3 + 1);
    expect(matrixToXlsxRows(matrix, budgets, "sv")).toHaveLength(18);
  });
});
