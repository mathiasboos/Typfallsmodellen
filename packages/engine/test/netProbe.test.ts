import { describe, expect, it } from "vitest";

import { loadDeathProbabilities } from "../src/data/nodeLoader.js";
import { defaultContext } from "../src/model/context.js";
import { Scheme, defaultInput } from "../src/model/input.js";
import { runWithTaxProbe } from "../src/model/netProbe.js";
import { run } from "../src/model/run.js";
import { AgeArray } from "../src/vba/ageArray.js";

const deaths = loadDeathProbabilities();

/**
 * The tax probe re-runs the model's own tax section on another income, so the
 * one thing to prove is that it is the same section: handed the income a run
 * actually had, it gives back the net that run actually reported.
 */
describe("the tax probe", () => {
  const input = defaultInput({ born: 1980, retirementAge: 68, startWorkAge: 25, monthlySalary: 38_500, scheme: Scheme.Itp1 });
  const context = defaultContext({ tjpPar: 0 });
  const probe = runWithTaxProbe(input, context, { deaths });
  const rows = probe.result.rows.filter((r) => r.age >= 69 && r.age <= 100);

  it("returns the result run() returns", () => {
    const plain = run(input, context, { deaths });
    expect(probe.result.rows).toEqual(plain.rows);
    expect(probe.result.table1).toEqual(plain.table1);
  });

  it("gives back the net a year of pension actually had", () => {
    expect(rows.length).toBeGreaterThan(20);
    for (const row of rows) {
      // Within the krona the row itself is rounded to.
      expect(Math.abs(probe.netAt(row.age, row.brutto) - row.netto), `net at ${row.age}`).toBeLessThanOrEqual(1);
    }
  });

  it("gives it back at the retirement age itself too, and with capital income in the model", () => {
    // The first pension year is the one the part-year rules touch and the one the tab's
    // budget is first set against; `kapital` adds a tax on capital from that age on.
    for (const kapital of [0, 50_000, -50_000]) {
      for (const [born, retirementAge] of [[1980, 68], [1975, 66], [1960, 65], [1990, 70]] as const) {
        const typfall = defaultInput({ born, retirementAge, startWorkAge: 25, monthlySalary: 38_500, scheme: Scheme.Itp1 });
        const other = runWithTaxProbe(typfall, defaultContext({ tjpPar: 0, kapital }), { deaths });
        for (const row of other.result.rows.filter((r) => r.age >= retirementAge && r.age <= retirementAge + 3)) {
          expect(
            Math.abs(other.netAt(row.age, row.brutto) - row.netto),
            `born ${born}, retires at ${retirementAge}, kapital ${kapital}, net at ${row.age}`,
          ).toBeLessThanOrEqual(1);
        }
      }
    }
  });

  it("follows the price basis the rows are expressed in", () => {
    for (const priceBasis of [0, -1]) {
      const other = runWithTaxProbe(input, defaultContext({ tjpPar: 0, priceBasis }), { deaths });
      for (const row of other.result.rows.filter((r) => r.age >= 69 && r.age <= 95 && r.age % 5 === 0)) {
        expect(Math.abs(other.netAt(row.age, row.brutto) - row.netto), `basis ${priceBasis} at ${row.age}`).toBeLessThanOrEqual(1);
      }
    }
  });

  it("leaves the run as it found it, however often it is asked", () => {
    const age = 75;
    const first = probe.netAt(age, 300_000);
    probe.netAt(age, 900_000);
    probe.netAt(age, 120_000);
    expect(probe.netAt(age, 300_000)).toBe(first);
    const row = rows.find((r) => r.age === age)!;
    expect(Math.abs(probe.netAt(age, row.brutto) - row.netto)).toBeLessThanOrEqual(1);
  });

  it("leaves the whole state of the run as it found it, not just the net", () => {
    // Every age array, every running scalar, the loop's leftovers and the output rows.
    const snapshot = (): string =>
      JSON.stringify(
        Object.entries(probe.run.s).map(([name, value]) => [
          name,
          value instanceof AgeArray ? value.slice() : value,
        ]),
      );
    const before = snapshot();
    // The last age is an early one, whose rules year differs from the loop's own last
    // (`gage` is 66 there and 69 at the end), so a leak into the leftovers would show.
    for (const age of [105, 90, 75, 69, 68, 60, 30]) {
      probe.netAt(age, 250_000);
      probe.grossFor(age, 180_000);
    }
    expect(snapshot()).toBe(before);
  });

  it("never leaves more than it was given, and leaves more the more it is given", () => {
    for (const age of [69, 75, 90]) {
      let previous = 0;
      for (let gross = 0; gross <= 1_200_000; gross += 25_000) {
        const net = probe.netAt(age, gross);
        expect(net, `net of ${gross} at ${age}`).toBeLessThanOrEqual(gross);
        expect(net, `net of ${gross} at ${age}`).toBeGreaterThanOrEqual(previous - 1e-6);
        previous = net;
      }
    }
  });

  it("taxes the same income more in the year before the higher grundavdrag begins", () => {
    // The riktålder of 1980 is 68 and the allowance for the old starts a year
    // later: 29 000 kr a month is taxed harder at 68 than at 69.
    expect(probe.netAt(68, 12 * 29_000)).toBeLessThan(probe.netAt(69, 12 * 29_000));
  });

  it("inverts: the gross that leaves a net, left through the tax, is that net", () => {
    for (const age of [69, 75, 90]) {
      for (const net of [120_000, 220_000, 330_000, 600_000]) {
        const gross = probe.grossFor(age, net);
        expect(gross, `gross for ${net} at ${age}`).toBeGreaterThanOrEqual(net);
        expect(probe.netAt(age, gross), `${net} at ${age}`).toBeGreaterThanOrEqual(net - 0.01);
        // And it is the least such gross: a few kronor less leaves less.
        expect(probe.netAt(age, gross - 200), `${net} at ${age}`).toBeLessThan(net);
      }
    }
  });

  it("has no gross for no net, and refuses an age outside the run", () => {
    expect(probe.grossFor(75, 0)).toBe(0);
    expect(probe.netAt(75, 0)).toBe(0);
    expect(() => probe.netAt(500, 100_000)).toThrow(RangeError);
  });
});
