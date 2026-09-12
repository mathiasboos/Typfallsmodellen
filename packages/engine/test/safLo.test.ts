import { describe, expect, it } from "vitest";

import { SAF_LO, STP_ } from "../src/tjanstepension/safLo.js";
import type { SchemeContext } from "../src/tjanstepension/types.js";

/**
 * As with ITP, there is no fixture for this agreement. These assert structure
 * and the transitions the VBA states outright; the golden files are the judge of
 * whether the rates themselves are right.
 */
const context = (overrides: Partial<SchemeContext> = {}): SchemeContext => ({
  year: 2024,
  born: 1985,
  wStart: 23,
  tjpPar: 66,
  flexPension: 0,
  marginal: 0,
  ...overrides,
});

const IBB = 76_200;
const SALARY = 462_000;
const PBB = 58_800;

describe("SAF_LO", () => {
  it("did not exist before 1996", () => {
    expect(SAF_LO(40, SALARY, IBB, 65, 1995, context({ year: 1995 }))).toBe(0);
    expect(SAF_LO(40, SALARY, IBB, 65, 1996, context({ year: 1996 }))).toBeGreaterThan(0);
  });

  it("pays nothing on a salary of zero or less", () => {
    expect(SAF_LO(40, 0, IBB, 65, 2024, context())).toBe(0);
  });

  it("lowers its entry age over time", () => {
    // 28 before 2000, 25 to 2020, then 24 in 2021, 23 in 2022, 22 from 2023.
    const at = (age: number, year: number) =>
      SAF_LO(age, SALARY, IBB, 65, year, context({ year, wStart: 20 }));

    expect(at(26, 1999)).toBe(0);
    expect(at(28, 1999)).toBeGreaterThan(0);

    expect(at(24, 2020)).toBe(0);
    expect(at(24, 2021)).toBeGreaterThan(0);

    expect(at(23, 2021)).toBe(0);
    expect(at(23, 2022)).toBeGreaterThan(0);

    expect(at(22, 2022)).toBe(0);
    expect(at(22, 2023)).toBeGreaterThan(0);
  });

  it("phases in the higher rate above the breakpoint between 2008 and 2013", () => {
    const high = 12 * IBB;
    const premiums = [2007, 2008, 2009, 2010, 2011, 2013].map((year) =>
      SAF_LO(40, high, IBB, 65, year, context({ year })),
    );
    for (let i = 1; i < premiums.length; i += 1) {
      expect(premiums[i]!, `year index ${i}`).toBeGreaterThan(premiums[i - 1]!);
    }
  });

  it("charges the same rate either side of the breakpoint before 2008", () => {
    const year = 2005;
    const ctx = context({ year });
    const low = SAF_LO(40, 300_000, IBB, 65, year, ctx);
    const high = SAF_LO(40, 900_000, IBB, 65, year, ctx);
    expect(high / low).toBeCloseTo(3, 2);
  });

  it("adds the flexpension premium only from 2014", () => {
    const with2020 = SAF_LO(40, SALARY, IBB, 65, 2020, context({ year: 2020, flexPension: 0.02 }));
    const without2020 = SAF_LO(40, SALARY, IBB, 65, 2020, context({ year: 2020 }));
    expect(with2020).toBeGreaterThan(without2020);

    const with2013 = SAF_LO(40, SALARY, IBB, 65, 2013, context({ year: 2013, flexPension: 0.02 }));
    const without2013 = SAF_LO(40, SALARY, IBB, 65, 2013, context({ year: 2013 }));
    expect(with2013).toBe(without2013);
  });

  it("rises with salary throughout", () => {
    let previous = 0;
    for (const salary of [200_000, 400_000, 600_000, 900_000]) {
      const premium = SAF_LO(40, salary, IBB, 65, 2024, context());
      expect(premium).toBeGreaterThan(previous);
      previous = premium;
    }
  });

  it("pays nothing past the occupational retirement age", () => {
    expect(SAF_LO(66, SALARY, IBB, 65, 2024, context())).toBe(0);
  });
});

describe("STP_", () => {
  it("needs at least three years of service", () => {
    expect(STP_(5, 2, PBB, 1950, 65)).toBe(0);
    expect(STP_(5, 3, PBB, 1950, 65)).toBeGreaterThan(0);
  });

  it("rises with average points", () => {
    expect(STP_(6, 30, PBB, 1950, 65)).toBeGreaterThan(STP_(3, 30, PBB, 1950, 65));
  });

  it("scales down below the cohort's target years of service", () => {
    // Born 1950 targets 37 years.
    const full = STP_(5, 37, PBB, 1950, 65);
    expect(STP_(5, 18, PBB, 1950, 65)).toBeLessThan(full);
    expect(STP_(5, 40, PBB, 1950, 65)).toBe(full);
  });

  it("raises the target years for later cohorts", () => {
    // The same service is worth proportionally less to a cohort with a higher
    // target, so a 1950 cohort gets less than a 1937 one for 30 years.
    expect(STP_(5, 30, PBB, 1950, 65)).toBeLessThan(STP_(5, 30, PBB, 1937, 65));
  });

  it("increases for deferral past 65 and does not decrease before it", () => {
    const at65 = STP_(5, 30, PBB, 1950, 65);
    expect(STP_(5, 30, PBB, 1950, 68)).toBeGreaterThan(at65);
    // There is no withdrawal before 65: the age is raised to 65 first.
    expect(STP_(5, 30, PBB, 1950, 62)).toBe(at65);
  });

  it("stops increasing past 70", () => {
    expect(STP_(5, 30, PBB, 1950, 73)).toBe(STP_(5, 30, PBB, 1950, 70));
  });

  it("rounds to a whole krona per month unless marginal turns rounding off", () => {
    expect(STP_(5.3, 29, PBB, 1950, 66) % 12).toBe(0);
    expect(STP_(5.3, 29, PBB, 1950, 66, 1) % 12).not.toBe(0);
  });
});
