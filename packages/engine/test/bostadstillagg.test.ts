import { describe, expect, it } from "vitest";

import { BTP, SBTP, btp_sbtp } from "../src/bidrag/bostadstillagg.js";
import type { SbtpContext } from "../src/bidrag/types.js";
import type { RunVectors } from "../src/model/runVectors.js";

/**
 * No offline fixture covers the benefits module. These tests assert rent
 * bracket boundaries, the year transitions the VBA states outright, and the
 * quirks worth pinning down -- never that an amount is the legally correct one.
 * The golden files from a real Excel run are the judge of that.
 */

const PBB = 58_800;

const vectors: RunVectors = {
  wage: () => 0,
  ibb: () => 80_600,
  pbb: () => PBB,
  fpb: () => PBB,
  kpiJune: () => 400,
  kpi: () => 400,
  year: (age) => 1959 + age,
};

const sbtpContext = (overrides: Partial<SbtpContext> = {}): SbtpContext => ({
  dela: 2,
  rulesFromUtg: 0,
  born: 1959,
  age: 66,
  slutage: 100,
  Iyear: 0,
  vectors,
  ...overrides,
});

describe("BTP", () => {
  /** A single pensioner in 2025 with 7 000 kr a month in rent. */
  const single = (inkomst: number, over: Partial<Record<string, number>> = {}) =>
    BTP(
      inkomst,
      0,
      over.hyra ?? 84_000,
      0,
      PBB,
      { dela: 2 },
      over.ap ?? 1,
      1,
      over.form ?? 0,
      0,
      0,
      1,
      0,
      over.year ?? 2025,
      2100,
      0,
      1,
      over.bald ?? 66,
    );

  it("did not exist before 1978", () => {
    expect(single(0, { year: 1977 })).toBe(0);
    expect(single(0, { year: 1978 })).toBeGreaterThan(0);
  });

  it("covers the rent in four bands from 2022, plus the consumption supplement", () => {
    // 100% to 36 000, 90% to 60 000, 70% to 84 000, 50% to 90 000; then 840/mo.
    const cover = 36_000 + 24_000 * 0.9 + 24_000 * 0.7;
    expect(single(0)).toBeCloseTo(cover + 840 * 12, 9);
  });

  it("puts the third band's edge at 84 001, not 84 000", () => {
    const at84000 = single(0, { hyra: 84_000 });
    const at84001 = single(0, { hyra: 84_001 });
    expect(at84001 - at84000).toBeCloseTo(0.5, 9);
  });

  it("caps the rent at 7 500 a month", () => {
    expect(single(0, { hyra: 90_000 })).toBeCloseTo(single(0, { hyra: 120_000 }), 9);
  });

  it("reduces by 62 öre per krona of reduction income above the free amount", () => {
    const fri = 2.43 * PBB;
    const full = single(0);
    // From 2020 the reduction income is 93% of everything above garantipension.
    expect(single(200_000)).toBeCloseTo(full - 0.62 * (200_000 * 0.93 - fri), 9);
  });

  it("pays nothing once the reduction eats the whole amount", () => {
    expect(single(1_000_000)).toBe(0);
  });

  it("QUIRK: the 0.95 share for non-pensioners after 2017 is dead code", () => {
    // `If year > 2017 Then PAR = 0.96` overwrites it unconditionally.
    const covered = single(0, { year: 2018, ap: 0, hyra: 50_000 });
    expect(covered).toBeCloseTo(50_000 * 0.96 + 340 * 12, 9);
    expect(covered).not.toBeCloseTo(50_000 * 0.95 + 340 * 12, 9);
  });

  it("QUIRK: wealth below the threshold is added to income in full", () => {
    // 50 000 is under the 100 000 floor, so nothing of it should count -- but
    // the VBA only scales the excess and leaves a sub-threshold amount intact.
    expect(single(150_000, { form: 50_000 })).toBeCloseTo(single(200_000), 9);
    expect(single(150_000, { form: 50_000 })).toBeLessThan(single(150_000));
  });

  it("scales only the excess when wealth is above the threshold", () => {
    expect(single(150_000, { form: 300_000 })).toBeCloseTo(
      single(150_000 + 0.15 * 200_000),
      9,
    );
  });

  it("QUIRK: the age guard on the consumption supplement lapses after 2021", () => {
    const at = (year: number, bald: number) => single(0, { year, bald, hyra: 36_000 });
    // 2012-2021: only from 65.
    expect(at(2021, 60) + 340 * 12).toBeCloseTo(at(2021, 70), 9);
    // 2023: paid at any age.
    expect(at(2023, 60)).toBeCloseTo(at(2023, 70), 9);
    expect(at(2023, 60)).toBeCloseTo(36_000 + 840 * 12, 9);
  });

  describe("dela", () => {
    const couple = (dela: number, ap = 1) =>
      BTP(0, 0, 84_000, 1, PBB, { dela }, ap, 1, 0, 0, 0, 1, 0, 2025, 2100, 0, 1, 66);

    it("reports the individual's amount by default", () => {
      // Half the rent each: 100% to 18 000, 90% to 30 000, 70% to 42 000.
      const cover = 18_000 + 12_000 * 0.9 + 12_000 * 0.7;
      expect(couple(2)).toBeCloseTo(cover + (840 * 12) / 2, 9);
    });

    it("QUIRK: doubles the amount for a cohabiting pensioner when dela is 1", () => {
      // `BTPm` is assigned the finished amount, then added to it.
      expect(couple(1)).toBeCloseTo(2 * couple(2), 9);
      // An equal split lands back on the individual's amount.
      expect(couple(3)).toBeCloseTo(couple(2), 9);
    });

    it("leaves the spouse's amount at zero when the household is not a pensioner", () => {
      expect(couple(1, 0)).toBeCloseTo(couple(2, 0), 9);
      expect(couple(3, 0)).toBeCloseTo(couple(2, 0) / 2, 9);
    });
  });
});

describe("SBTP", () => {
  /** A single pensioner in 2025 with a small pension and 6 000 kr a month rent. */
  const sbtp = (over: Partial<Record<string, number>> = {}, context = sbtpContext()) =>
    SBTP(
      over.inkomst ?? 100_000,
      over.hyra ?? 6_000,
      0,
      over.btpb ?? 0,
      over.avdrag ?? 43_600,
      context,
      0.32,
      over.ap ?? 1,
      over.form ?? 0,
      PBB,
      0,
      over.year ?? 2025,
      21_000,
      0,
      0,
      over.bald ?? 66,
      over.kapital ?? 0,
      0,
      over.marginal ?? 0,
      -99,
      over.ftid ?? 1,
    );

  it("did not exist before 1978", () => {
    expect(sbtp({ year: 1977 })).toBe(0);
  });

  it("falls as income rises and never goes negative", () => {
    const amounts = [0, 50_000, 100_000, 150_000, 250_000, 500_000].map((inkomst) =>
      sbtp({ inkomst }),
    );
    for (let i = 1; i < amounts.length; i += 1) {
      expect(amounts[i]!).toBeLessThanOrEqual(amounts[i - 1]!);
      expect(amounts[i]!).toBeGreaterThanOrEqual(0);
    }
  });

  it("QUIRK: reads a rent under 10 000 as monthly and multiplies it by twelve", () => {
    expect(sbtp({ hyra: 6_000 })).toBeCloseTo(sbtp({ hyra: 72_000 }), 9);
    // The heuristic flips at exactly 10 000, which is a cliff in the rent.
    expect(sbtp({ hyra: 9_999 })).not.toBeCloseTo(sbtp({ hyra: 10_000 }), 9);
  });

  it("QUIRK: pays nothing to a non-pensioner in 2012-2017, because Lev is never set", () => {
    for (const year of [2012, 2015, 2017]) {
      expect(sbtp({ year, ap: 0 })).toBe(0);
      expect(sbtp({ year, ap: 1 })).toBeGreaterThan(0);
    }
    // From 2018 the level is assigned regardless of pensioner status.
    expect(sbtp({ year: 2018, ap: 0 })).toBeGreaterThan(0);
  });

  it("QUIRK: skipping to äldreförsörjningsstöd leaves Xage at 0", () => {
    // With ftid >= 1, Xage is riktålder + 1 and the grundavdrag depends on age.
    expect(sbtp({ bald: 66 })).not.toBe(sbtp({ bald: 67 }));
    // With ftid < 1 the GoTo jumps over the assignment, so Xage stays 0 --
    // every age then clears it, and the age stops mattering.
    expect(sbtp({ bald: 66, ftid: 0 })).toBe(sbtp({ bald: 67, ftid: 0 }));
    expect(sbtp({ bald: 20, ftid: 0 })).toBe(sbtp({ bald: 67, ftid: 0 }));
  });

  it("QUIRK: weights wealth 0.7 in äldreförsörjningsstöd, but only when rounding is on", () => {
    // Only the excess above 100 000 counts, scaled by 15%.
    const excess = 0.15 * 200_000;
    const drop = (marginal: number) =>
      sbtp({ ftid: 0, form: 100_000, marginal }) - sbtp({ ftid: 0, form: 300_000, marginal });
    expect(drop(0)).toBeCloseTo(0.7 * excess, 6);
    expect(drop(1)).toBeCloseTo(excess, 6);
  });

  it("adds the bostadstillägg to disposable income, reducing itself krona for krona", () => {
    const without = sbtp({ btpb: 0 });
    expect(without - sbtp({ btpb: 10_000 })).toBeCloseTo(10_000, 6);
  });

  it("reports the individual's amount by default and splits it when dela is 3", () => {
    const couple = (dela: number) =>
      SBTP(100_000, 6_000, 1, 0, 43_600, sbtpContext({ dela }), 0.32, 1, 0, PBB, 0, 2025);
    expect(couple(3)).toBeCloseTo(couple(1) / 2, 9);
    expect(couple(2)).toBeLessThanOrEqual(couple(1));
  });
});

describe("btp_sbtp", () => {
  it("rounds the total to a whole krona a month", () => {
    expect(btp_sbtp(100, 0, 0, 2025)).toBe(96);
    expect(btp_sbtp(100, 0, 1, 2025)).toBe(100);
    expect(btp_sbtp(84_480, 0, 0, 2025)).toBe(84_480);
  });

  it("suppressed amounts under 25 kr a month before 2014, but only without SBTP", () => {
    expect(btp_sbtp(100, 0, 0, 2013)).toBe(0);
    expect(btp_sbtp(100, 1, 0, 2013)).toBe(96);
    expect(btp_sbtp(100, 0, 0, 2014)).toBe(96);
  });
});
