import { describe, expect, it } from "vitest";

import { loadDeathProbabilities } from "../src/data/nodeLoader.js";
import { Scheme, defaultContext, defaultInput, run } from "../src/index.js";
import type { ModelContext, TypfallInput, TypfallResult } from "../src/index.js";

/**
 * `ModelContext.formogenhetPerAge`: the housing supplement's wealth test year by year,
 * for the web app's "Uttagsstrategi" tab. Not in the workbook, which holds one amount.
 */

const deaths = loadDeathProbabilities();

/** A low income and a rent, so that the housing supplement is paid. */
const typfall = defaultInput({ born: 1961, monthlySalary: 24_000, scheme: Scheme.SafLo }) as TypfallInput;

const compute = (context: Partial<ModelContext>): TypfallResult =>
  run(typfall, defaultContext({ hyra: 7_000, ...context }) as ModelContext, { deaths });

const benefits = (result: TypfallResult): number[] => result.rows.map((r) => r.bidrag);
const everyAge = (wealth: number): number[] => new Array<number>(106).fill(wealth);

describe("wealth by age in the housing supplement", () => {
  const none = compute({});
  const par = typfall.retirementAge;
  const at = (result: TypfallResult, age: number) => result.rows.find((r) => r.age === age)!.bidrag;

  it("pays a housing supplement to this typfall at all", () => {
    expect(at(none, par + 2)).toBeGreaterThan(1_000);
  });

  it("disregards wealth up to 100 000 kronor, where the workbook's own setting counts it as income", () => {
    expect(benefits(compute({ formogenhetPerAge: everyAge(100_000) }))).toEqual(benefits(none));
    expect(benefits(compute({ formogenhetPerAge: everyAge(50_000) }))).toEqual(benefits(none));
    // The workbook's setting, kept as written: 100 000 kronor goes into income whole.
    expect(at(compute({ formogenhet: 100_000 }), par + 2)).toBeLessThan(at(none, par + 2));
  });

  it("tests wealth above 100 000 kronor exactly as the workbook's setting does", () => {
    expect(benefits(compute({ formogenhetPerAge: everyAge(400_000) }))).toEqual(benefits(compute({ formogenhet: 400_000 })));
  });

  it("follows the wealth from year to year", () => {
    // 400 000 kronor until two years into retirement, then nothing.
    const falling = everyAge(400_000).map((w, age) => (age >= par + 2 ? 0 : w));
    const result = compute({ formogenhetPerAge: falling });
    expect(at(result, par + 1)).toBeLessThan(at(none, par + 1));
    expect(at(result, par + 3)).toBe(at(none, par + 3));
  });

  it("changes nothing when it is absent", () => {
    expect(compute({}).table1).toEqual(none.table1);
  });
});
