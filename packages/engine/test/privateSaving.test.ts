import { beforeEach, describe, expect, it } from "vitest";

import {
  AvkastningsskattKFISK,
  DEPOT_GAIN_TAX,
  isUntaxedSaving,
  PrivatSpar,
  resetPrivateSavingState,
  SavingType,
  Schablonintakt,
} from "../src/saving/privateSaving.js";

beforeEach(() => {
  resetPrivateSavingState();
});

describe("Schablonintakt", () => {
  it("uses the preceding year's borrowing rate", () => {
    // 2024's notional return is built from the 30 November 2023 rate of 2.62%,
    // plus the 1 percentage point margin added since 2018.
    expect(Schablonintakt(2024)).toBeCloseTo(0.0262 + 0.01, 6);
  });

  it("applies a floor of 1.25%", () => {
    // 2021's rate came from a negative figure at the end of 2020.
    expect(Schablonintakt(2021)).toBe(0.0125);
  });

  it("adds 0.75 points between 2016 and 2017, and 1 point from 2018", () => {
    expect(Schablonintakt(2017)).toBeCloseTo(Math.max(0.0027 + 0.0075, 0.0125), 6);
    expect(Schablonintakt(2018)).toBeCloseTo(Math.max(0.0049 + 0.01, 0.0125), 6);
  });

  it("forecasts 2.5% beyond the table, plus the margin", () => {
    expect(Schablonintakt(2030)).toBeCloseTo(0.025 + 0.01, 6);
  });

  it("carries the previous value forward for a year the table does not cover", () => {
    // No Case exists before 1986, and the VBA's implicit module variable keeps
    // whatever was there. See the note in privateSaving.ts.
    expect(Schablonintakt(1980)).toBe(0);
    const recent = Schablonintakt(2024);
    expect(Schablonintakt(1980)).toBe(recent);
  });
});

describe("AvkastningsskattKFISK", () => {
  it("taxes 27% of the notional return before 2012 and 30% after", () => {
    const balance = 1_000_000;
    expect(AvkastningsskattKFISK(balance, 2010)).toBeCloseTo(
      balance * 0.27 * Schablonintakt(2010),
      6,
    );
    expect(AvkastningsskattKFISK(balance, 2020)).toBeCloseTo(
      balance * 0.3 * Schablonintakt(2020),
      6,
    );
  });

  it("exempts the first 150 000 in 2025 and the first 300 000 from 2026", () => {
    expect(AvkastningsskattKFISK(100_000, 2025)).toBe(0);
    expect(AvkastningsskattKFISK(200_000, 2025)).toBeGreaterThan(0);
    expect(AvkastningsskattKFISK(200_000, 2026)).toBe(0);
    expect(AvkastningsskattKFISK(400_000, 2026)).toBeGreaterThan(0);
  });

  it("returns the previous call's tax at exactly the allowance", () => {
    // The VBA tests > and < and covers neither the boundary nor, having no
    // Option Explicit, resets its variable. Reproduced deliberately.
    const previous = AvkastningsskattKFISK(1_000_000, 2025);
    expect(previous).toBeGreaterThan(0);
    expect(AvkastningsskattKFISK(150_000, 2025)).toBe(previous);
  });
});

describe("PrivatSpar", () => {
  it("grows the balance and deducts the yield tax", () => {
    const opening = 500_000;
    const grown = PrivatSpar(opening, 0, 1.05, 2020, SavingType.KF);
    expect(grown).toBeGreaterThan(opening);
    expect(grown).toBeLessThan(opening * 1.05);
  });

  it("adds the year's saving", () => {
    const withSaving = PrivatSpar(500_000, 60_000, 1.05, 2020, SavingType.KF);
    resetPrivateSavingState();
    const without = PrivatSpar(500_000, 0, 1.05, 2020, SavingType.KF);
    expect(withSaving).toBeGreaterThan(without);
  });

  it("leaves an ISK untouched before it existed", () => {
    expect(PrivatSpar(500_000, 60_000, 1.05, 2011, SavingType.ISK)).toBe(500_000);
    expect(PrivatSpar(500_000, 60_000, 1.05, 2012, SavingType.ISK)).not.toBe(500_000);
  });

  it("taxes an ISK on a quarterly average, so it differs from a KF", () => {
    const kf = PrivatSpar(500_000, 60_000, 1.05, 2020, SavingType.KF);
    resetPrivateSavingState();
    const isk = PrivatSpar(500_000, 60_000, 1.05, 2020, SavingType.ISK);
    expect(kf).not.toBe(isk);
  });

  it("compounds over a working life without going negative", () => {
    let balance = 0;
    for (let year = 2000; year <= 2040; year += 1) {
      balance = PrivatSpar(balance, 24_000, 1.04, year, SavingType.KF);
      expect(balance, `year ${year}`).toBeGreaterThanOrEqual(0);
      expect(Number.isFinite(balance), `year ${year}`).toBe(true);
    }
    expect(balance).toBeGreaterThan(41 * 24_000);
  });
});

describe("a depot (an ordinary securities account), which is not in the workbook", () => {
  it("is taxed 30% on the gain and keeps the rest, the balance being after tax already", () => {
    expect(DEPOT_GAIN_TAX).toBe(0.3);
    expect(PrivatSpar(1_000_000, 0, 1.05, 2030, SavingType.Depot)).toBeCloseTo(1_000_000 * (1 + 0.05 * 0.7), 6);
    expect(PrivatSpar(1_000_000, 0, 1.017, 2030, SavingType.Depot)).toBeCloseTo(1_000_000 * (1 + 0.017 * 0.7), 6);
  });

  it("adds the year's saving at half a year's growth after tax", () => {
    const kept = 1 + 0.05 * 0.7;
    expect(PrivatSpar(0, 12_000, 1.05, 2030, SavingType.Depot)).toBeCloseTo(12_000 * kept ** 0.5, 6);
    expect(PrivatSpar(500_000, 12_000, 1.05, 2030, SavingType.Depot)).toBeCloseTo(500_000 * kept + 12_000 * kept ** 0.5, 6);
  });

  it("gives a loss no refund", () => {
    expect(PrivatSpar(1_000_000, 0, 0.95, 2030, SavingType.Depot)).toBeCloseTo(950_000, 6);
    expect(PrivatSpar(1_000_000, 0, 1, 2030, SavingType.Depot)).toBe(1_000_000);
  });

  it("is taxed on no notional return: no allowance and no year it depends on", () => {
    const early = PrivatSpar(2_000_000, 0, 1.04, 2012, SavingType.Depot);
    expect(PrivatSpar(2_000_000, 0, 1.04, 2040, SavingType.Depot)).toBe(early);
    expect(PrivatSpar(100_000, 0, 1.04, 2040, SavingType.Depot)).toBeCloseTo(100_000 * (1 + 0.04 * 0.7), 6);
  });

  it("leaves what the yield tax carries between calls alone", () => {
    // The two wrappers' tax functions remember their last answer (see the note at the top of the file).
    resetPrivateSavingState();
    PrivatSpar(2_000_000, 120_000, 1.04, 2030, SavingType.Depot);
    expect(AvkastningsskattKFISK(150_000, 2025)).toBe(0);
  });

  it("grows less than the untaxed return and more than nothing over a working life", () => {
    let balance = 0;
    let untaxed = 0;
    for (let year = 2026; year <= 2066; year += 1) {
      balance = PrivatSpar(balance, 24_000, 1.04, year, SavingType.Depot);
      untaxed = untaxed * 1.04 + 24_000 * 1.04 ** 0.5;
      expect(Number.isFinite(balance), `year ${year}`).toBe(true);
    }
    expect(balance).toBeGreaterThan(41 * 24_000);
    expect(balance).toBeLessThan(untaxed);
  });

  it("is one of the savings that are built up in the pps pot and paid out untaxed, the IPS (0) not", () => {
    expect([0, 1, 2, 3, 4, -1].map((kind) => isUntaxedSaving(kind))).toEqual([false, true, true, true, false, false]);
  });
});
