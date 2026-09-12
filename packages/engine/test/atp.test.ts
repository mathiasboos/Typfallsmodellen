import { describe, expect, it } from "vitest";

import { fnorm, LAST_ATP_COHORT, pts, tp, tpFaktor } from "../src/pension/atp.js";

const PBB = 58_800;

describe("fnorm", () => {
  it("is 1.6% from 2001 on", () => {
    expect(fnorm(2001)).toBe(1.016);
    expect(fnorm(2030)).toBe(1.016);
  });

  it("is the lower figure in 2000, its first year", () => {
    expect(fnorm(2000)).toBe(0.996);
  });

  it("does not apply before 2000", () => {
    expect(fnorm(1999)).toBe(1);
  });
});

describe("tpFaktor", () => {
  it("is neutral at exactly 65", () => {
    expect(tpFaktor(65)).toBe(1);
  });

  it("reduces by 0.5% a month for early withdrawal", () => {
    expect(tpFaktor(64)).toBeCloseTo(1 - 0.005 * 12, 12);
    expect(tpFaktor(62)).toBeCloseTo(1 - 0.005 * 36, 12);
  });

  it("increases by 0.7% a month for deferral", () => {
    expect(tpFaktor(66)).toBeCloseTo(1 + 0.007 * 12, 12);
    expect(tpFaktor(65.5)).toBeCloseTo(1 + 0.007 * 6, 12);
  });

  it("stops increasing past 70", () => {
    expect(tpFaktor(71)).toBeCloseTo(tpFaktor(70), 12);
    expect(tpFaktor(80)).toBeCloseTo(tpFaktor(70), 12);
  });
});

describe("tp", () => {
  it("pays nothing to cohorts after 1953", () => {
    expect(tp(4, 30, 0, 65, LAST_ATP_COHORT + 1, 65, PBB)).toBe(0);
  });

  it("pays nothing before pension is drawn", () => {
    expect(tp(4, 30, 0, 65, 1950, 64, PBB)).toBe(0);
  });

  it("pays a married person less, the folkpension part being lower", () => {
    expect(tp(4, 30, 1, 65, 1950, 65, PBB)).toBeLessThan(tp(4, 30, 0, 65, 1950, 65, PBB));
  });

  it("reduces in proportion below thirty qualifying years", () => {
    const full = tp(4, 30, 0, 65, 1950, 65, PBB);
    const short = tp(4, 15, 0, 65, 1950, 65, PBB);
    expect(short).toBeCloseTo(full / 2, 6);
  });

  it("pays nothing below three qualifying years", () => {
    expect(tp(4, 2, 0, 65, 1950, 65, PBB)).toBe(0);
  });

  it("does not apply the qualifying-year reduction to cohorts before 1938", () => {
    // The residence criterion was met differently for them.
    expect(tp(4, 15, 0, 65, 1930, 65, PBB)).toBe(tp(4, 30, 0, 65, 1930, 65, PBB));
  });

  it("rises with average ATP points", () => {
    expect(tp(6, 30, 0, 65, 1950, 65, PBB)).toBeGreaterThan(tp(3, 30, 0, 65, 1950, 65, PBB));
  });

  it("applies the withdrawal factor twice, as the original does", () => {
    // Not a typo in the port: `ATP = ATP * faktor` is followed by
    // `tp_ = ATP * andel_ * faktor`. See the note in atp.ts.
    const atExactly65 = tp(4, 30, 0, 65, 1950, 65, PBB);
    const deferredOneYear = tp(4, 30, 0, 66, 1950, 66, PBB);
    expect(deferredOneYear / atExactly65).toBeCloseTo(tpFaktor(66) ** 2, 9);
  });

  it("scales with the share of pension drawn", () => {
    const full = tp(4, 30, 0, 65, 1950, 65, PBB, 1);
    expect(tp(4, 30, 0, 65, 1950, 65, PBB, 0.5)).toBeCloseTo(full / 2, 9);
  });
});

describe("pts", () => {
  it("falls as ATP rises, and stops at zero", () => {
    expect(pts(65, PBB, 0)).toBeGreaterThan(0);
    expect(pts(65, PBB, 10_000)).toBeLessThan(pts(65, PBB, 0));
    expect(pts(65, PBB, 1_000_000)).toBe(0);
  });

  it("is reduced for withdrawal before 65", () => {
    expect(pts(62, PBB, 0)).toBeLessThan(pts(65, PBB, 0));
  });

  it("pays nothing below three qualifying years", () => {
    expect(pts(65, PBB, 0, 0, 2)).toBe(0);
  });

  it("rounds to a whole krona per month unless marginal turns rounding off", () => {
    expect(pts(65, PBB, 1234) % 12).toBe(0);
    expect(pts(65, PBB, 1234, 1) % 12).not.toBe(0);
  });
});
