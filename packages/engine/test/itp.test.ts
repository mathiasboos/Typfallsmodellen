import { describe, expect, it } from "vitest";

import { tlITP1, tlITP2A, tlITP2F } from "../src/tjanstepension/itp.js";
import type { SchemeContext } from "../src/tjanstepension/types.js";

/**
 * No fixture exists for the occupational pension agreements: the Brutto sheet
 * covers the public pension only, and the model's own `deltal_tjp` cell is
 * cleared when the workbook opens.
 *
 * So these tests assert structure and the transitions the VBA states outright --
 * where a rate steps, which years a rule applies from -- and never claim that a
 * rate is the legally correct one, because nothing available offline can confirm
 * that. The golden files from a real Excel run are the judge of these.
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

describe("tlITP1", () => {
  it("pays nothing before 25", () => {
    expect(tlITP1(24, SALARY, IBB, 66, context())).toBe(0);
    expect(tlITP1(26, SALARY, IBB, 66, context())).toBeGreaterThan(0);
  });

  it("pays nothing on a salary of zero or less", () => {
    expect(tlITP1(40, 0, IBB, 66, context())).toBe(0);
    expect(tlITP1(40, -100, IBB, 66, context())).toBe(0);
  });

  it("part-pays in the year the entry age is reached", () => {
    const partYear = tlITP1(25, SALARY, IBB, 66, context({ born: 1985, wStart: 23 }));
    const fullYear = tlITP1(30, SALARY, IBB, 66, context({ born: 1985, wStart: 23 }));
    expect(partYear).toBeGreaterThan(0);
    expect(partYear).toBeLessThanOrEqual(fullYear);
  });

  it("charges 4.5% below the breakpoint", () => {
    const below = 7.5 * IBB - 60_000;
    expect(tlITP1(40, below, IBB, 66, context())).toBeCloseTo(0.045 * below, -1);
  });

  it("charges roughly 30% on salary above the breakpoint, against 4.5% below", () => {
    const breakpoint = 7.5 * IBB;
    const step = 60_000;
    const belowRate =
      (tlITP1(40, breakpoint - step, IBB, 66, context()) -
        tlITP1(40, breakpoint - 2 * step, IBB, 66, context())) /
      step;
    const aboveRate =
      (tlITP1(40, breakpoint + 2 * step, IBB, 66, context()) -
        tlITP1(40, breakpoint + step, IBB, 66, context())) /
      step;
    expect(belowRate).toBeCloseTo(0.045, 3);
    expect(aboveRate).toBeCloseTo(0.3, 3);
  });

  it("rises with salary throughout", () => {
    let previous = 0;
    for (const salary of [200_000, 400_000, 600_000, 900_000, 1_500_000]) {
      const premium = tlITP1(40, salary, IBB, 66, context());
      expect(premium).toBeGreaterThan(previous);
      previous = premium;
    }
  });

  it("leaves the 2023 cap inert at any realistic salary", () => {
    // The cap compares a *monthly* salary against 30 *annual* income base
    // amounts, so it only bites above roughly 360 base amounts a year. See the
    // note in itp.ts -- it is kept as written, not corrected.
    const huge = 60 * IBB;
    expect(tlITP1(40, huge, IBB, 66, context({ year: 2023 }))).toBe(
      tlITP1(40, huge, IBB, 66, context({ year: 2022 })),
    );

    // Past the threshold it does fire, which is how we know it is wired up.
    const absurd = 400 * IBB * 12;
    expect(tlITP1(40, absurd, IBB, 66, context({ year: 2023 }))).toBeLessThan(
      tlITP1(40, absurd, IBB, 66, context({ year: 2022 })),
    );
  });

  it("adds the flexpension premium only from 2014", () => {
    const withFlex = context({ year: 2020, flexPension: 0.02 });
    const withoutFlex = context({ year: 2020, flexPension: 0 });
    expect(tlITP1(40, SALARY, IBB, 66, withFlex)).toBeGreaterThan(
      tlITP1(40, SALARY, IBB, 66, withoutFlex),
    );

    const before = context({ year: 2013, flexPension: 0.02 });
    const beforeNone = context({ year: 2013, flexPension: 0 });
    expect(tlITP1(40, SALARY, IBB, 66, before)).toBe(tlITP1(40, SALARY, IBB, 66, beforeNone));
  });

  it("rounds the premium to whole kronor unless marginal turns rounding off", () => {
    const rounded = tlITP1(40, 462_123, IBB, 66, context());
    const unrounded = tlITP1(40, 462_123, IBB, 66, context({ marginal: 1 }));
    expect(Number.isInteger(rounded / 12)).toBe(true);
    expect(rounded).not.toBe(unrounded);
  });
});

describe("tlITP2A (ITPK)", () => {
  it("could not be earned before 1997", () => {
    expect(tlITP2A(40, SALARY, 65, context({ year: 1996 }))).toBe(0);
    expect(tlITP2A(40, SALARY, 65, context({ year: 1997 }))).toBeGreaterThan(0);
  });

  it("pays nothing before 28", () => {
    expect(tlITP2A(27, SALARY, 65, context())).toBe(0);
    expect(tlITP2A(29, SALARY, 65, context())).toBeGreaterThan(0);
  });

  it("charges a flat 2% with no breakpoint", () => {
    expect(tlITP2A(40, 300_000, 65, context())).toBeCloseTo(0.02 * 300_000, -1);
    expect(tlITP2A(40, 900_000, 65, context())).toBeCloseTo(0.02 * 900_000, -1);
  });

  it("scales linearly with salary, unlike ITP 1", () => {
    const single = tlITP2A(40, 300_000, 65, context());
    const double = tlITP2A(40, 600_000, 65, context());
    expect(double / single).toBeCloseTo(2, 3);
  });
});

describe("tlITP2F (defined benefit)", () => {
  it("pays nothing on a salary of zero", () => {
    expect(tlITP2F(0, IBB, context())).toBe(0);
  });

  it("replaces 10% of salary below the breakpoint", () => {
    const below = 7.5 * IBB - 60_000;
    expect(tlITP2F(below, IBB, context({ tjpPar: 65 }))).toBeCloseTo(0.1 * below, -2);
  });

  it("replaces far more of the salary between 7.5 and 20 base amounts", () => {
    const atBreakpoint = tlITP2F(7.5 * IBB, IBB, context({ tjpPar: 65 }));
    const above = tlITP2F(7.5 * IBB + 100_000, IBB, context({ tjpPar: 65 }));
    expect(above - atBreakpoint).toBeCloseTo(0.65 * 100_000, -2);
  });

  it("stops rising above 30 income base amounts", () => {
    expect(tlITP2F(40 * IBB, IBB, context())).toBe(tlITP2F(30 * IBB, IBB, context()));
  });

  it("scales down below thirty years of service", () => {
    const full = tlITP2F(SALARY, IBB, context(), 30);
    expect(tlITP2F(SALARY, IBB, context(), 15)).toBeCloseTo(full / 2, -2);
  });

  it("reduces for early retirement and increases for late", () => {
    const at65 = tlITP2F(SALARY, IBB, context({ tjpPar: 65 }));
    expect(tlITP2F(SALARY, IBB, context({ tjpPar: 62 }))).toBeLessThan(at65);
    expect(tlITP2F(SALARY, IBB, context({ tjpPar: 68 }))).toBeGreaterThan(at65);
  });

  it("defers at 0.6% a month here, where tilläggspension uses 0.7%", () => {
    const at65 = tlITP2F(SALARY, IBB, context({ tjpPar: 65 }));
    const at66 = tlITP2F(SALARY, IBB, context({ tjpPar: 66 }));
    expect(at66 / at65).toBeCloseTo(1 + 0.006 * 12, 2);
  });

  it("stops increasing past 70", () => {
    expect(tlITP2F(SALARY, IBB, context({ tjpPar: 72 }))).toBeCloseTo(
      tlITP2F(SALARY, IBB, context({ tjpPar: 70 })),
      6,
    );
  });
});
