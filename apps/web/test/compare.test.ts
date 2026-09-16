import { describe, expect, it } from "vitest";

import { defaultInput } from "@typfallsmodellen/engine";

import { applyScenario } from "../src/compare.js";

/**
 * `applyScenario`'s whole job is to lay a variant's three fields over the
 * baseline and leave everything else alone -- the one piece of the
 * comparison tab with real logic, and the only piece that does not need a
 * browser to check (every result renderer it feeds is already checked
 * against the real page by `tools/build/verify-offline.mjs`).
 */
describe("applying a scenario override", () => {
  const baseline = defaultInput();

  it("overrides only salary, retirement age and scheme", () => {
    const applied = applyScenario(baseline, {
      id: "1",
      label: "Scenario 1",
      monthlySalary: baseline.monthlySalary + 5000,
      retirementAge: baseline.retirementAge + 1,
      scheme: baseline.scheme,
    });
    expect(applied.monthlySalary).toBe(baseline.monthlySalary + 5000);
    expect(applied.retirementAge).toBe(baseline.retirementAge + 1);
    expect(applied.scheme).toBe(baseline.scheme);
    // Everything else -- birth year, started-work age, the economic
    // assumptions, any advanced-mode override already on the baseline --
    // stays exactly the baseline's own value.
    const { monthlySalary: _s, retirementAge: _r, scheme: _c, ...restBaseline } = baseline;
    const { monthlySalary: _s2, retirementAge: _r2, scheme: _c2, ...restApplied } = applied;
    expect(restApplied).toEqual(restBaseline);
  });

  it("carries a typed salary vector or PGB entries through untouched", () => {
    const withExtras = {
      ...baseline,
      ownIncome: [{ age: 30, income: 400_000, wage: 380_000 }],
      pgbManual: [{ age: 40, sa: 10_000, vpl: 0, studier: 0 }],
    };
    const applied = applyScenario(withExtras, {
      id: "1",
      label: "Scenario 1",
      monthlySalary: withExtras.monthlySalary,
      retirementAge: withExtras.retirementAge,
      scheme: withExtras.scheme,
    });
    expect(applied.ownIncome).toBe(withExtras.ownIncome);
    expect(applied.pgbManual).toBe(withExtras.pgbManual);
  });

  it("changes with the baseline it is applied to, not a frozen copy", () => {
    const variant = {
      id: "1",
      label: "Scenario 1",
      monthlySalary: baseline.monthlySalary,
      retirementAge: baseline.retirementAge,
      scheme: baseline.scheme,
    };
    const olderBaseline = { ...baseline, born: baseline.born - 10 };
    expect(applyScenario(baseline, variant).born).toBe(baseline.born);
    expect(applyScenario(olderBaseline, variant).born).toBe(olderBaseline.born);
  });
});
