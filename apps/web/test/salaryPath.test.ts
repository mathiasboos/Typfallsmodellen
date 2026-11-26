import { describe, expect, it } from "vitest";

import type { OwnIncomeYear } from "@typfallsmodellen/engine";

import { applyPartTime } from "../src/salaryPath.js";

/**
 * "Deltid": the part-time periods that scale the computed salary path.
 *
 * The UI around them is covered by the offline check (tools/build/
 * verify-offline.mjs); what is worth pinning here is the arithmetic the
 * screen and the model both read -- which years a period covers, what it
 * does to the two columns, and how two periods that overlap resolve.
 */
describe("applyPartTime", () => {
  const born = 1959;
  const baseline: readonly OwnIncomeYear[] = [
    { age: 14, income: 100_000, wage: 100_000 }, // below FIRST_AGE: never read
    { age: 15, income: 200_000, wage: 180_000 }, // 1974
    { age: 60, income: 500_000, wage: 450_000 }, // 2019
    { age: 61, income: 510_000, wage: 460_000 }, // 2020
    { age: 62, income: 520_000, wage: 470_000 }, // 2021
  ];

  it("drops the ages an own path is never read at", () => {
    const path = applyPartTime(baseline, born, []);
    expect(path.map((r) => r.age)).toEqual([15, 60, 61, 62]);
  });

  it("leaves every year alone when there are no periods", () => {
    expect(applyPartTime(baseline, born, [])).toEqual([
      { age: 15, income: 200_000, wage: 180_000 },
      { age: 60, income: 500_000, wage: 450_000 },
      { age: 61, income: 510_000, wage: 460_000 },
      { age: 62, income: 520_000, wage: 470_000 },
    ]);
  });

  it("scales both columns, and only inside the period's own years", () => {
    // 2020-2021 is age 61-62 for someone born in 1959; 2019 is not in it.
    const path = applyPartTime(baseline, born, [{ fromYear: 2020, toYear: 2021, share: 0.5 }]);
    expect(path).toEqual([
      { age: 15, income: 200_000, wage: 180_000 },
      { age: 60, income: 500_000, wage: 450_000 },
      { age: 61, income: 255_000, wage: 230_000 },
      { age: 62, income: 260_000, wage: 235_000 },
    ]);
  });

  it("is inclusive at both ends, so one year is a one-year period", () => {
    const path = applyPartTime(baseline, born, [{ fromYear: 2020, toYear: 2020, share: 0.8 }]);
    expect(path.map((r) => r.income)).toEqual([200_000, 500_000, 408_000, 520_000]);
  });

  it("lets a later period carve out of an earlier one", () => {
    // 2019-2021 at 50 %, then 2021 alone at 25 %: the narrow one is last, so
    // it wins its own year and the wide one keeps the other two.
    const path = applyPartTime(baseline, born, [
      { fromYear: 2019, toYear: 2021, share: 0.5 },
      { fromYear: 2021, toYear: 2021, share: 0.25 },
    ]);
    expect(path.map((r) => r.income)).toEqual([200_000, 250_000, 255_000, 130_000]);
  });

  it("rounds once, after scaling", () => {
    // 0.333 of 500 000 is 166 500 exactly; rounding the baseline first and
    // then scaling would be the same here, but a third of an odd figure is
    // where the two part -- 0.5 of 510 000 is 255 000, never 255 000.5.
    const path = applyPartTime([{ age: 60, income: 333, wage: 333 }], born, [
      { fromYear: 2019, toYear: 2019, share: 0.5 },
    ]);
    expect(path).toEqual([{ age: 60, income: 167, wage: 167 }]);
  });

  it("maps years through the birth year, so the same period moves with it", () => {
    const period = [{ fromYear: 2020, toYear: 2020, share: 0.5 }];
    // Born a year later: 2020 is age 60, not 61.
    const path = applyPartTime(baseline, 1960, period);
    expect(path.map((r) => r.income)).toEqual([200_000, 250_000, 510_000, 520_000]);
  });

  it("leaves everything alone when the period covers no worked year", () => {
    const path = applyPartTime(baseline, born, [{ fromYear: 2090, toYear: 2095, share: 0.5 }]);
    expect(path).toEqual(applyPartTime(baseline, born, []));
  });
});
