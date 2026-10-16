/**
 * The model's own annuity factors are computed once per cohort and death table,
 * not once per run.
 *
 * `calculateDeltal` is the most expensive thing a run does, and a run builds a
 * fresh `DeltalTables` every time, so before they were shared a caller that ran
 * the model dozens of times for one person (the web app's retirement-age study)
 * recomputed the same figures each time. What is pinned here is the contract
 * that makes sharing safe: one computation per (table, cohort) however many
 * `DeltalTables` ask, no leakage from one death table to another, and results
 * that do not change when a run is repeated.
 */
import { describe, expect, it, vi } from "vitest";

import annuityTables from "../../data/annuity-tables.json" with { type: "json" };
import mortalityMeta from "../../data/mortality.json" with { type: "json" };
import { loadDeathProbabilities } from "../src/data/nodeLoader.js";
import { defaultContext } from "../src/model/context.js";
import { defaultInput } from "../src/model/input.js";
import { run } from "../src/model/run.js";
import { DeltalTables } from "../src/pension/deltal.js";
import type { AnnuityTablesData } from "../src/pension/deltal.js";
import { DeathProbabilities } from "../src/pension/mortality.js";

// `vi.mock` is hoisted above the imports, so what it records has to be too.
const computed = vi.hoisted(() => [] as { cohort: number; deaths: unknown }[]);

vi.mock("../src/pension/mortality.js", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../src/pension/mortality.js")>();
  return {
    ...actual,
    calculateDeltal: (cohort: number, deaths: import("../src/pension/mortality.js").DeathProbabilities) => {
      computed.push({ cohort, deaths });
      return actual.calculateDeltal(cohort, deaths);
    },
  };
});

const published = annuityTables as unknown as AnnuityTablesData;
const context = defaultContext();
const timesComputed = (cohort: number, deaths: unknown): number =>
  computed.filter((c) => c.cohort === cohort && c.deaths === deaths).length;

describe("the model's own annuity factors", () => {
  it("are computed once per cohort and death table, however many DeltalTables ask", () => {
    const deaths = loadDeathProbabilities();
    const lives = [1, 2, 3].map(() => new DeltalTables(published, deaths, context).expectedLife(1980, 65));
    expect(new Set(lives).size).toBe(1);
    expect(timesComputed(1980, deaths)).toBe(1);

    // Another cohort is another computation; the first stays as it was.
    new DeltalTables(published, deaths, context).expectedLife(1975, 65);
    expect(timesComputed(1975, deaths)).toBe(1);
    expect(timesComputed(1980, deaths)).toBe(1);
  });

  it("are never lent from one death table to another", () => {
    const real = loadDeathProbabilities();
    // A table in which nobody survives for long: the same shape, very different figures.
    const grim = DeathProbabilities.fromPacked(new Float64Array(mortalityMeta.count).fill(0.2), mortalityMeta);

    const realLife = new DeltalTables(published, real, context).expectedLife(1980, 65);
    const grimLife = new DeltalTables(published, grim, context).expectedLife(1980, 65);
    expect(grimLife).toBeLessThan(realLife);
    expect(timesComputed(1980, grim)).toBe(1);

    // And asking the real table again still gets the real answer.
    expect(new DeltalTables(published, real, context).expectedLife(1980, 65)).toBe(realLife);
    expect(timesComputed(1980, grim)).toBe(1);
  });

  it("do not change what a run says when it is repeated", () => {
    const deaths = loadDeathProbabilities();
    const input = defaultInput({ born: 1980, retirementAge: 68, startWorkAge: 25, monthlySalary: 38_500 });
    const first = run(input, context, { deaths });
    const again = run(input, context, { deaths });
    expect(again.rows).toEqual(first.rows);
    expect(again.table1).toEqual(first.table1);
    // Two runs for one cohort, one computation.
    expect(timesComputed(1980, deaths)).toBe(1);
  });
});
