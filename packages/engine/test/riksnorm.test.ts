import { describe, expect, it } from "vitest";

import { bistOld, bist25 } from "../src/bidrag/riksnorm.js";
import type { RiksnormContext } from "../src/bidrag/types.js";
import type { RunVectors } from "../src/model/runVectors.js";

/**
 * The amounts come from packages/data/riksnorm.json, parsed out of the VBA by
 * tools/extract/extract_riksnorm.py; `npm run check:riksnorm` is what keeps
 * those honest. These tests cover the rules around them.
 *
 * One genuine reference value exists: a commented-out `verb()` in Bidrag.bas
 * records 46 240 kr a month for a named household in 2025, which the first test
 * reproduces. It is the only figure the workbook's author left behind for this
 * function, and the only offline check the benefits module has.
 */

const vectors: RunVectors = {
  wage: () => 0,
  ibb: () => 80_600,
  pbb: () => 58_800,
  fpb: () => 58_800,
  kpiJune: () => 400,
  // 2% a year from age 66, so the indexation past 2025 is visible.
  kpi: (age) => 400 * 1.02 ** (age - 66),
  year: (age) => 1959 + age,
};

const context = (overrides: Partial<RiksnormContext> = {}): RiksnormContext => ({
  marginal: 0,
  born: 1959,
  age: 66,
  vectors,
  ...overrides,
});

/** `bist25(civ, hyra, disp, b1..b8, wage, context, year)`, with a childless default. */
const bist = (
  civ: number,
  hyra: number,
  disp: number,
  barn: number[] = [0, 0, 0, 0, 0, 0, 0, 0],
  over: { wage?: number; year?: number; context?: RiksnormContext } = {},
) =>
  bist25(
    civ,
    hyra,
    disp,
    barn[0]!,
    barn[1]!,
    barn[2]!,
    barn[3]!,
    barn[4]!,
    barn[5]!,
    barn[6]!,
    barn[7]!,
    over.wage ?? 0,
    over.context ?? context(),
    over.year ?? 2025,
  );

const old = (
  civ: number,
  hyra: number,
  disp: number,
  barn: number[] = [0, 0, 0, 0, 0, 0, 0, 0],
  year = 2005,
) =>
  bistOld(
    civ,
    hyra,
    disp,
    barn[0]!,
    barn[1]!,
    barn[2]!,
    barn[3]!,
    barn[4]!,
    barn[5]!,
    barn[6]!,
    barn[7]!,
    0,
    context(),
    year,
  );

describe("bist25", () => {
  it("reproduces the reference value the VBA's own test records", () => {
    // `verb()` in Bidrag.bas: two adults, 20 600 kr rent, children in the 4-6,
    // 7-10 and 11-14 bands, 2025 -- and 46 240 kr a month.
    const annual = bist(2, 20_600, 0, [0, 0, 0, 1, 1, 2, 0, 0]);
    expect(annual / 12).toBe(46_240);
    expect(annual).toBe(554_880);
  });

  it("sums the norm for a childless single adult", () => {
    // 3 910 for the adult plus 1 270 of shared household costs.
    expect(bist(1, 0, 0)).toBe(12 * (3910 + 1270));
  });

  it("adds the rent on top, and pays what income does not cover", () => {
    expect(bist(1, 5000, 0)).toBe(12 * (3910 + 1270 + 5000));
    expect(bist(1, 5000, 50_000)).toBe(12 * (3910 + 1270 + 5000) - 50_000);
  });

  it("pays nothing once income covers the whole year's need", () => {
    const need = 12 * (3910 + 1270);
    expect(bist(1, 0, need - 1)).toBe(1);
    expect(bist(1, 0, need)).toBe(0);
    expect(bist(1, 0, need + 100_000)).toBe(0);
  });

  it("extrapolates the shared household cost past seven members", () => {
    // The last step, 2 810 - 2 620, repeats for each further member.
    const seven = bist(1, 0, 0, [0, 0, 0, 0, 6, 0, 0, 0]);
    const eight = bist(1, 0, 0, [0, 0, 0, 0, 7, 0, 0, 0]);
    const nine = bist(1, 0, 0, [0, 0, 0, 0, 8, 0, 0, 0]);
    // Each further child adds its own cost plus one step of the shared cost.
    expect(eight - seven).toBe(12 * (3790 + 190));
    expect(nine - eight).toBe(12 * (3790 + 190));
  });

  it("follows CPI past the last tabulated year", () => {
    // The stub puts KPI 2% higher a year on.
    expect(bist(1, 0, 0, undefined, { year: 2026, context: context({ age: 67 }) })).toBe(
      Math.trunc(12 * (3910 + 1270) * 1.02 + 0.5),
    );
  });

  it("disregards a quarter of earned income from 2013", () => {
    expect(bist(1, 0, 40_000, undefined, { wage: 40_000, year: 2014 })).toBe(
      bist(1, 0, 30_000, undefined, { year: 2014 }),
    );
    // Not before 2013, where the argument is ignored.
    expect(bist(1, 0, 40_000, undefined, { wage: 40_000, year: 2012 })).toBe(
      bist(1, 0, 40_000, undefined, { year: 2012 }),
    );
  });

  it("does not apply before 2005", () => {
    expect(bist(1, 5000, 0, undefined, { year: 2004 })).toBe(0);
  });
});

describe("bistOld", () => {
  it("applies only to 1985-2005", () => {
    expect(old(1, 5000, 0, undefined, 1984)).toBe(0);
    expect(old(1, 5000, 0, undefined, 1985)).toBeGreaterThan(0);
    expect(old(1, 5000, 0, undefined, 2005)).toBeGreaterThan(0);
    expect(old(1, 5000, 0, undefined, 2006)).toBe(0);
  });

  it("QUIRK: compares a year's income against one month's need", () => {
    // 2 590 for the adult plus 770 shared, a month.
    const monthly = 2590 + 770;
    expect(old(1, 0, 0)).toBe(12 * monthly);
    // A krona below one month's norm still pays a full year's shortfall...
    expect(old(1, 0, monthly - 1)).toBe(12 * monthly - (monthly - 1));
    // ...and a krona above it pays nothing, on an annual income of 3 360 kr.
    expect(old(1, 0, monthly)).toBe(0);
    // `bist25` writes the same test as `disp < bistand * 12`, so the household
    // there keeps the benefit until a year's income covers a year's need.
    expect(bist(1, 0, 12 * (3910 + 1270) - 1)).toBe(1);
  });

  it("takes medical and dental care out of the norm from 1994", () => {
    // The 1994 table is 1 643 for an infant, less 32 for care.
    const withChild = old(1, 0, 0, [1, 0, 0, 0, 0, 0, 0, 0], 1994);
    expect(withChild).toBe(12 * (1643 - 32 + 3403));
    // 1995 deducts 33, and 1996 deducts it for the nine months it applied.
    expect(old(1, 0, 0, [1, 0, 0, 0, 0, 0, 0, 0], 1995)).toBe(12 * (1666 - 33 + (3451 - 57)));
  });

  it("has no shared household cost before 1996", () => {
    // The Gn table is all zeros up to 1995, so only the adult's norm is paid.
    expect(old(1, 0, 0, undefined, 1995)).toBe(12 * (3451 - 57));
    expect(old(1, 0, 0, undefined, 1997)).toBe(12 * (2190 + 760));
  });

  it("throws on a household with no members, as VBA's subscript check does", () => {
    expect(() => old(0, 0, 0)).toThrow(RangeError);
  });
});
