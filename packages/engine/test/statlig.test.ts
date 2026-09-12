import { describe, expect, it } from "vitest";

import { Kapan, PA_indiv, tlPA16 } from "../src/tjanstepension/statlig.js";
import type { SchemeContext } from "../src/tjanstepension/types.js";

/** No fixture exists for these agreements; see the note in itp.test.ts. */
const context = (overrides: Partial<SchemeContext> = {}): SchemeContext => ({
  year: 2024,
  born: 1990,
  wStart: 23,
  tjpPar: 66,
  flexPension: 0,
  marginal: 0,
  ...overrides,
});

const IBB = 76_200;
const SALARY = 462_000;

describe("tlPA16 (avdelning 1)", () => {
  it("pays nothing before work starts or on no salary", () => {
    expect(tlPA16(22, SALARY, IBB, 66, 2024, context({ wStart: 23 }))).toBe(0);
    expect(tlPA16(40, 0, IBB, 66, 2024, context())).toBe(0);
  });

  it("charges 6% below the breakpoint and 31.5% above, before 2026", () => {
    const step = 60_000;
    const breakpoint = 7.5 * IBB;
    const ctx = context({ year: 2024 });
    const belowRate =
      (tlPA16(40, breakpoint - step, IBB, 66, 2024, ctx) -
        tlPA16(40, breakpoint - 2 * step, IBB, 66, 2024, ctx)) /
      step;
    const aboveRate =
      (tlPA16(40, breakpoint + 2 * step, IBB, 66, 2024, ctx) -
        tlPA16(40, breakpoint + step, IBB, 66, 2024, ctx)) /
      step;
    expect(belowRate).toBeCloseTo(0.06, 3);
    expect(aboveRate).toBeCloseTo(0.315, 3);
  });

  it("steps the flex part up through 2026 and again after", () => {
    const at = (year: number) => tlPA16(40, SALARY, IBB, 66, year, context({ year }));
    expect(at(2026)).toBeGreaterThan(at(2025));
    expect(at(2027)).toBeGreaterThan(at(2026));
    expect(at(2028)).toBe(at(2027));
  });

  it("accrues to 67 from 2004 but only to 65 before", () => {
    expect(tlPA16(66, SALARY, IBB, 67, 2003, context({ year: 2003 }))).toBe(0);
    expect(tlPA16(66, SALARY, IBB, 67, 2004, context({ year: 2004 }))).toBeGreaterThan(0);
  });
});

describe("Kapan (avdelning 2, collectively agreed)", () => {
  it("did not exist before 1991", () => {
    expect(Kapan(1990, 1960, IBB, SALARY, context({ year: 1990, born: 1960 }), 65)).toBe(0);
    expect(
      Kapan(1995, 1960, IBB, SALARY, context({ year: 1995, born: 1960, wStart: 23 }), 65),
    ).toBeGreaterThan(0);
  });

  it("lowers its qualifying age from 28 to 23 in 2008", () => {
    const young = (year: number) =>
      Kapan(year, year - 25, IBB, SALARY, context({ year, born: year - 25, wStart: 20 }), 65);
    expect(young(2007)).toBe(0);
    expect(young(2008)).toBeGreaterThan(0);
  });

  it("caps pensionable salary at 30 income base amounts from 2003", () => {
    const ctx = context({ year: 2010, born: 1970, wStart: 23 });
    expect(Kapan(2010, 1970, IBB, 40 * IBB, ctx, 65)).toBe(
      Kapan(2010, 1970, IBB, 30 * IBB, ctx, 65),
    );
  });

  it("raises its upper age from 65 to 69 after 2023", () => {
    const ctx = (year: number) => context({ year, born: year - 67, wStart: 23, tjpPar: 70 });
    expect(Kapan(2023, 2023 - 67, IBB, SALARY, ctx(2023), 70)).toBe(0);
    expect(Kapan(2024, 2024 - 67, IBB, SALARY, ctx(2024), 70)).toBeGreaterThan(0);
  });

  it("applies no rounding", () => {
    const value = Kapan(2024, 1990, IBB, 462_137, context({ born: 1990 }), 65);
    expect(Number.isInteger(value)).toBe(false);
  });
});

describe("PA_indiv (avdelning 2, individually chosen)", () => {
  it("did not exist before 2003", () => {
    const ctx = context({ year: 2002, born: 1970, wStart: 23 });
    expect(PA_indiv(2002, 1970, IBB, SALARY, ctx, 65)).toBe(0);
  });

  it("rises from 2.3% to 2.5% in 2008", () => {
    const at = (year: number) =>
      PA_indiv(year, 1970, IBB, SALARY, context({ year, born: 1970, wStart: 23 }), 65);
    expect(at(2008) / at(2007)).toBeCloseTo(0.025 / 0.023, 3);
  });

  it("adds a higher flex premium for those born 1965 or later, from 2024", () => {
    const older = PA_indiv(2024, 1960, IBB, SALARY, context({ year: 2024, born: 1960 }), 65);
    const younger = PA_indiv(2024, 1970, IBB, SALARY, context({ year: 2024, born: 1970 }), 65);
    expect(younger).toBeGreaterThan(older);
    expect(younger / older).toBeCloseTo(0.04 / 0.03, 3);
  });

  it("blends the 2026 rate across the year the flex part rises", () => {
    const at = (year: number) =>
      PA_indiv(year, 1970, IBB, SALARY, context({ year, born: 1970 }), 65);
    expect(at(2026)).toBeGreaterThan(at(2025));
    expect(at(2027)).toBeGreaterThan(at(2026));
  });

  it("pays nothing before 23", () => {
    const ctx = context({ year: 2024, born: 2002, wStart: 20 });
    expect(PA_indiv(2024, 2002, IBB, SALARY, ctx, 65)).toBe(0);
    expect(PA_indiv(2025, 2002, IBB, SALARY, context({ year: 2025, born: 2002, wStart: 20 }), 65)).toBeGreaterThan(0);
  });
});

describe("avdelning 2 as a whole", () => {
  it("is the sum of the two premiums, as Mcalc adds them", () => {
    const ctx = context({ year: 2024, born: 1970, wStart: 23 });
    const kapan = Kapan(2024, 1970, IBB, SALARY, ctx, 66);
    const indiv = PA_indiv(2024, 1970, IBB, SALARY, ctx, 66);
    expect(kapan).toBeGreaterThan(0);
    expect(indiv).toBeGreaterThan(0);
    // Together they come to roughly 6% of salary, close to avdelning 1's rate.
    expect((kapan + indiv) / SALARY).toBeGreaterThan(0.04);
    expect((kapan + indiv) / SALARY).toBeLessThan(0.08);
  });
});
