import { describe, expect, it } from "vitest";

import {
  BURIAL_ONLY_RATE,
  CHURCH_MEMBER_RATE,
  KOMMUNALSKATT,
  KOMMUNALSKATT_YEAR,
  STOCKHOLM_BURIAL_RATE,
  TRANAS_BURIAL_RATE,
} from "../src/kommunalskatt.js";

/**
 * A structural check on the hand-copied municipality table, not a
 * re-derivation of it -- the 290 rates themselves are trusted from the
 * reference calculator's own SCB-sourced source, per kommunalskatt.ts's own
 * header. What's checkable without an independent source is that the copy
 * came through intact: a plausible count, a plausible range, no duplicate or
 * empty keys, and a year that isn't stale by construction.
 */
describe("the municipality tax table", () => {
  it("has close to Sweden's actual number of municipalities, none duplicated or empty", () => {
    const names = Object.keys(KOMMUNALSKATT);
    // 290 kommuner as of 2026; a wide band rather than the exact count, so a
    // yearly municipality merger/split doesn't make this test the blocker.
    expect(names.length).toBeGreaterThan(280);
    expect(names.length).toBeLessThan(300);
    expect(new Set(names).size, "no duplicate municipality names").toBe(names.length);
    for (const name of names) expect(name.trim(), `"${name}" is not empty`).not.toBe("");
  });

  it("keeps every rate inside the real range Swedish municipal tax takes", () => {
    // The lowest and highest rows any recent year has shown are around 29%
    // and 36%; a wide 25-45% band catches a transcription slip (a misplaced
    // decimal point, a percentage typed as a fraction) without being tripped
    // by next year's real movement within the actual range.
    for (const [name, rate] of Object.entries(KOMMUNALSKATT)) {
      expect(rate, `${name}`).toBeGreaterThan(25);
      expect(rate, `${name}`).toBeLessThan(45);
    }
  });

  it("carries a real year, not a placeholder", () => {
    expect(KOMMUNALSKATT_YEAR).toBeGreaterThanOrEqual(2026);
    expect(KOMMUNALSKATT_YEAR).toBeLessThan(2100);
  });
});

/**
 * The two defaults offered for `begravningsavgift` -- the workbook's own
 * K_skatt data, not new numbers, so what's checkable here is that the
 * clamped `lastActualYear` logic picked a real, sane year and a fraction in
 * the right ballpark, not the raw percentage or something a hundredfold off.
 */
describe("the church/burial-fee defaults", () => {
  it("reads the member rate as a fraction, in the range a combined church+burial fee actually takes", () => {
    expect(CHURCH_MEMBER_RATE.rate).toBeGreaterThan(0.005);
    expect(CHURCH_MEMBER_RATE.rate).toBeLessThan(0.03);
    expect(CHURCH_MEMBER_RATE.year).toBeGreaterThanOrEqual(2020);
    expect(CHURCH_MEMBER_RATE.year).toBeLessThanOrEqual(2030);
  });

  it("reads the burial-only rate as a fraction, smaller than the member rate", () => {
    expect(BURIAL_ONLY_RATE.rate).toBeGreaterThan(0);
    expect(BURIAL_ONLY_RATE.rate).toBeLessThan(0.01);
    expect(BURIAL_ONLY_RATE.year).toBeGreaterThanOrEqual(2020);
    expect(BURIAL_ONLY_RATE.year).toBeLessThanOrEqual(2030);
  });

  it("keeps the member rate strictly above the burial-only rate", () => {
    // `kyrkoavgift` is documented as "church fee including burial fee" --
    // members pay both, so the combined figure has to exceed the burial-only
    // figure, or the two series would have been read backwards.
    expect(CHURCH_MEMBER_RATE.rate).toBeGreaterThan(BURIAL_ONLY_RATE.rate);
  });

  it("keeps Stockholm's and Tranås's own rates as small fractions", () => {
    expect(STOCKHOLM_BURIAL_RATE).toBeGreaterThan(0);
    expect(STOCKHOLM_BURIAL_RATE).toBeLessThan(0.01);
    expect(TRANAS_BURIAL_RATE).toBeGreaterThan(0);
    expect(TRANAS_BURIAL_RATE).toBeLessThan(0.01);
  });
});
