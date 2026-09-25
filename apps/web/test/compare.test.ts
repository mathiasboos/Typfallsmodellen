import { describe, expect, it } from "vitest";

import { defaultInput } from "@typfallsmodellen/engine";

import { applyScenario } from "../src/compare.js";

/**
 * `applyScenario`'s whole job is to lay a variant's five fields over the
 * baseline and leave everything else alone -- the one piece of the
 * comparison tab with real logic, and the only piece that does not need a
 * browser to check (every result renderer it feeds is already checked
 * against the real page by `tools/build/verify-offline.mjs`).
 */
describe("applying a scenario override", () => {
  const baseline = defaultInput();

  it("overrides only birth year, salary, retirement age, start-of-work age and scheme", () => {
    const applied = applyScenario(baseline, {
      id: "1",
      label: "Scenario 1",
      born: baseline.born - 5,
      monthlySalary: baseline.monthlySalary + 5000,
      retirementAge: baseline.retirementAge + 1,
      startWorkAge: baseline.startWorkAge + 2,
      scheme: baseline.scheme,
    });
    expect(applied.born).toBe(baseline.born - 5);
    expect(applied.monthlySalary).toBe(baseline.monthlySalary + 5000);
    expect(applied.retirementAge).toBe(baseline.retirementAge + 1);
    expect(applied.startWorkAge).toBe(baseline.startWorkAge + 2);
    expect(applied.scheme).toBe(baseline.scheme);
    // Everything else -- the economic assumptions, any advanced-mode
    // override already on the baseline -- stays exactly the baseline's own
    // value.
    const { born: _b, monthlySalary: _s, retirementAge: _r, startWorkAge: _w, scheme: _c, ...restBaseline } =
      baseline;
    const { born: _b2, monthlySalary: _s2, retirementAge: _r2, startWorkAge: _w2, scheme: _c2, ...restApplied } =
      applied;
    expect(restApplied).toEqual(restBaseline);
  });

  it("carries a typed salary vector or PGB entries through untouched", () => {
    const withExtras = {
      ...baseline,
      ownIncome: [{ age: 30, income: 400_000, wage: 380_000 }],
      pgbManual: [{ age: 40, sa: 10_000, studySemesters: 0 }],
    };
    const applied = applyScenario(withExtras, {
      id: "1",
      label: "Scenario 1",
      born: withExtras.born,
      monthlySalary: withExtras.monthlySalary,
      retirementAge: withExtras.retirementAge,
      startWorkAge: withExtras.startWorkAge,
      scheme: withExtras.scheme,
    });
    expect(applied.ownIncome).toBe(withExtras.ownIncome);
    expect(applied.pgbManual).toBe(withExtras.pgbManual);
  });

  it("changes with the baseline it is applied to, not a frozen copy", () => {
    const variant = {
      id: "1",
      label: "Scenario 1",
      born: baseline.born,
      monthlySalary: baseline.monthlySalary,
      retirementAge: baseline.retirementAge,
      startWorkAge: baseline.startWorkAge,
      scheme: baseline.scheme,
    };
    // `yearlyInflation` is not one of the five overridden fields, so it has
    // to keep tracking whichever baseline `applyScenario` is given, not a
    // value frozen from whenever the variant itself was built.
    const higherInflation = { ...baseline, yearlyInflation: baseline.yearlyInflation + 0.01 };
    expect(applyScenario(baseline, variant).yearlyInflation).toBe(baseline.yearlyInflation);
    expect(applyScenario(higherInflation, variant).yearlyInflation).toBe(higherInflation.yearlyInflation);
  });
});
