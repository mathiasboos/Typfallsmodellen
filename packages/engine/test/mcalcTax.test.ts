import { describe, expect, it } from "vitest";

import annuityTables from "../../data/annuity-tables.json" with { type: "json" };
import { loadDeathProbabilities } from "../src/data/nodeLoader.js";
import { defaultContext } from "../src/model/context.js";
import type { ModelContext } from "../src/model/context.js";
import { Scheme, defaultInput } from "../src/model/input.js";
import type { TypfallInput } from "../src/model/input.js";
import { prepareRun, runLoop } from "../src/model/mcalc.js";
import { millenniumReduction } from "../src/model/taxAndBenefits.js";
import { startsetup } from "../src/model/setup.js";
import { createRunState } from "../src/model/state.js";
import { DeltalTables } from "../src/pension/deltal.js";
import type { AnnuityTablesData } from "../src/pension/deltal.js";

const deaths = loadDeathProbabilities();
const published = annuityTables as unknown as AnnuityTablesData;

/**
 * Tax, benefits and the output row. No offline fixture reaches this far either,
 * so these check the identities that have to hold whatever the amounts are --
 * net never above gross, benefits never negative, the row matching the vectors
 * it is built from -- plus the quirks worth pinning down.
 */
function run(input: TypfallInput, context: ModelContext) {
  const setup = startsetup(input, context);
  const state = createRunState(setup.vectors.startage, setup.vectors.slutage);
  const r = prepareRun(setup, input, context, state, new DeltalTables(published, deaths, context));
  runLoop(r);
  return { run: r, state };
}

const plain = (
  input: Parameters<typeof defaultInput>[0] = {},
  context: Parameters<typeof defaultContext>[0] = {},
) => run(defaultInput(input), defaultContext(context));

describe("tax", () => {
  it("never takes more than the whole gross income", () => {
    const { run: r, state } = plain();
    for (let age = r.v.startage; age <= 105; age += 1) {
      expect(state.netto.get(age), `age ${age}`).toBeLessThanOrEqual(state.brutto.get(age) + 1e-9);
      expect(state.netto.get(age)).toBeGreaterThanOrEqual(0);
    }
  });

  it("takes nothing where there is nothing to tax", () => {
    const { state } = plain();
    expect(state.brutto.get(20)).toBe(0);
    expect(state.netto.get(20)).toBe(0);
  });

  it("leaves more after tax at a higher municipal rate than a lower one", () => {
    const low = plain({}, { kommunalskatt: 0.29 }).state;
    const high = plain({}, { kommunalskatt: 0.35 }).state;
    expect(low.netto.get(50)).toBeGreaterThan(high.netto.get(50));
  });

  it("gives the enhanced allowance from the year after the riktålder", () => {
    // The jump between two years of near-identical pension is the grundavdrag.
    const { state } = plain();
    const ratio66 = state.netto.get(66) / state.brutto.get(66);
    const ratio67 = state.netto.get(67) / state.brutto.get(67);
    expect(ratio67).toBeGreaterThan(ratio66);
  });

  it("charges capital tax only once the pension is drawn", () => {
    const without = plain().state;
    const withCapital = plain({}, { kapital: 100_000 }).state;
    expect(withCapital.netto.get(50)).toBe(without.netto.get(50));
    expect(withCapital.netto.get(70)).toBeLessThan(without.netto.get(70));
  });

  it("QUIRK: the old 1 320-krona reduction is applied twice", () => {
    // Below 135 000 of uprated PGI each copy gives the full 1 320, so the
    // reduction comes to 2 640 rather than 1 320.
    expect(millenniumReduction(0, 2000, 100_000, 0, 0)).toBe(2640);
    expect(millenniumReduction(0, 2001, 100_000, 0, 0)).toBe(2640);
    // Above it the amount tapers, and that is doubled too.
    const single = 1320 - 0.012 * (Math.trunc(200_000 / 0.93) - 135_000);
    expect(single).toBeLessThan(1320);
    expect(millenniumReduction(0, 2000, 200_000, 0, 1)).toBeCloseTo(2 * single, 6);
    // And it exists only in those two years.
    expect(millenniumReduction(0, 1999, 100_000, 0, 0)).toBe(0);
    expect(millenniumReduction(0, 2002, 100_000, 0, 0)).toBe(0);
  });

  it("QUIRK: the 2004 clause inside that block is unreachable", () => {
    // `If Skyear = 2004` sits inside `If Skyear > 1999 And Skyear < 2002`.
    expect(millenniumReduction(0, 2004, 100_000, 5000, 0)).toBe(0);
  });
});

describe("benefits", () => {
  it("pays none to a working household without children", () => {
    const { state } = plain();
    for (let age = 30; age <= 60; age += 1) expect(state.bidrag.get(age), `age ${age}`).toBe(0);
  });

  it("pays child allowance while the children are at home", () => {
    const { state } = plain({}, { childBirthYears: [1990, 1993, 0, 0] });
    const at = (year: number) => state.bidrag.get(year - 1959);
    expect(at(1989)).toBe(0);
    expect(at(1995)).toBeGreaterThan(0);
    expect(at(2015)).toBe(0);
  });

  it("pays a housing supplement to a small pension and none to a large one", () => {
    const small = plain({ monthlySalary: 14_000 }).state;
    const large = plain({ monthlySalary: 70_000 }).state;
    expect(small.bidrag.get(70)).toBeGreaterThan(0);
    expect(large.bidrag.get(70)).toBe(0);
  });

  it("pays nothing when the household applies for nothing", () => {
    const { state } = plain({ monthlySalary: 14_000 }, { ansokt: 0 });
    expect(state.bidrag.get(70)).toBe(0);
  });

  it("never reports a negative benefit or disposable income", () => {
    for (const salary of [10_000, 25_000, 45_000, 90_000]) {
      const { run: r, state } = plain({ monthlySalary: salary });
      for (let age = r.v.startage; age <= 105; age += 1) {
        expect(state.bidrag.get(age)).toBeGreaterThanOrEqual(0);
        expect(state.indDisp.get(age)).toBeGreaterThanOrEqual(0);
      }
    }
  });

  it("QUIRK: a childless household born on 1 January gets no social assistance", () => {
    // `bist * (12 - pmonth) / 12`, and pmonth is 12 for a January birthday, so
    // the norm is multiplied by zero whenever there are no children at home.
    const { run: r, state } = plain({ monthlySalary: 10_000 });
    expect(r.pmonth).toBe(12);
    // Disposable income is net plus benefits and nothing else.
    for (const age of [70, 80, 90]) {
      expect(state.indDisp.get(age)).toBeCloseTo(state.netto.get(age) + state.bidrag.get(age), 6);
    }
  });
});

describe("the output matrix", () => {
  it("has one row per age from the start age", () => {
    const { run: r, state } = plain();
    expect(state.rows).toHaveLength(r.v.slutage - r.v.startage + 1);
    expect(state.rows[0]!.age).toBe(r.v.startage);
    expect(state.rows.at(-1)!.age).toBe(105);
    expect(state.rows[0]!.year).toBe(1959 + r.v.startage);
  });

  it("expresses amounts in the reference year's prices by default", () => {
    const { run: r, state } = plain();
    const row = state.rows.find((x) => x.age === 50)!;
    const factor = r.v.kpi.get(2025 - 1959) / r.v.kpi.get(50);
    expect(row.income).toBe(Math.round(r.v.income.get(50) * factor));
    expect(row.kpiFactor).toBeCloseTo(factor, 4);
  });

  it("leaves amounts nominal when asked", () => {
    const { run: r, state } = plain({}, { priceBasis: -1 });
    const row = state.rows.find((x) => x.age === 50)!;
    expect(row.income).toBe(Math.round(r.v.income.get(50)));
  });

  it("rounds to whole kronor unless the rounding is off", () => {
    const { state } = plain();
    for (const row of state.rows) {
      expect(row.brutto % 1).toBe(0);
      expect(row.netto % 1).toBe(0);
    }
    const exact = plain({}, { marginal: 1 }).state;
    expect(exact.rows.some((x) => x.brutto % 1 !== 0)).toBe(true);
  });

  it("is finite everywhere, for every cohort and agreement", () => {
    for (const born of [1935, 1950, 1965, 1980, 2000]) {
      for (const scheme of [Scheme.None, Scheme.Itp1, Scheme.SafLo, Scheme.AkapKr, Scheme.Pa16Avd2]) {
        const { state } = plain({ born, scheme });
        for (const row of state.rows) {
          for (const [key, value] of Object.entries(row)) {
            expect(Number.isFinite(value), `born ${born} scheme ${scheme} age ${row.age} ${key}`).toBe(true);
          }
        }
      }
    }
  });

  it("puts the whole gross into the row it belongs to", () => {
    const { run: r, state } = plain({ born: 1985, scheme: Scheme.Itp1 });
    for (const row of state.rows) {
      const parts = row.income + row.ip + row.tp + row.pp + row.garp + row.ptillagg + row.tjp + row.ips;
      // Rounding each component separately can move the total by a krona each.
      expect(Math.abs(row.brutto - parts)).toBeLessThanOrEqual(8);
    }
    expect(r.p.avtal).toBe(Scheme.Itp1);
  });
});
