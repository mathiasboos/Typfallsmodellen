import { describe, expect, it } from "vitest";

import {
  CalcAntalBarn,
  CalcBarnPerAlder,
  barnbidraget,
  ustod,
} from "../src/bidrag/barn.js";
import { bobid } from "../src/bidrag/bostadsbidrag.js";

/**
 * No offline fixture covers the benefits module. These tests assert the
 * boundaries and year transitions the VBA states outright, and the quirks worth
 * pinning down. The pre-1990 branches of `barnbidraget` and `bobid` will get no
 * golden coverage either, because no typfall the model offers reaches them --
 * they are ported, not verified.
 */

describe("CalcAntalBarn", () => {
  it("counts a child from its birth year for nineteen years", () => {
    expect(CalcAntalBarn(1999, 2000, 0, 0, 0)).toBe(0);
    expect(CalcAntalBarn(2000, 2000, 0, 0, 0)).toBe(1);
    expect(CalcAntalBarn(2018, 2000, 0, 0, 0)).toBe(1);
    // The nineteenth year is zeroed, so the child is out a year before the
    // twenty-year window the comment describes closes.
    expect(CalcAntalBarn(2019, 2000, 0, 0, 0)).toBe(0);
    expect(CalcAntalBarn(2020, 2000, 0, 0, 0)).toBe(0);
  });

  it("adds up several children", () => {
    expect(CalcAntalBarn(2010, 2000, 2003, 0, 0)).toBe(2);
    expect(CalcAntalBarn(2010, 2000, 2003, 2006, 2008)).toBe(4);
  });

  it("drops one child at a time as each reaches nineteen", () => {
    expect(CalcAntalBarn(2019, 2000, 2003, 0, 0)).toBe(1);
    expect(CalcAntalBarn(2022, 2000, 2003, 0, 0)).toBe(0);
  });

  it("QUIRK: the fourth child's window is bounded by the third child's year", () => {
    // A fourth child born fifteen years after the third is six years old in
    // 2021 and still does not count, because 2021 is past `barn3 + 20`.
    expect(CalcAntalBarn(2021, 0, 0, 2000, 2015)).toBe(0);
    // And one born before the third keeps counting past its own nineteenth.
    expect(CalcAntalBarn(2019, 0, 0, 2000, 1995)).toBe(1);
  });
});

describe("CalcBarnPerAlder", () => {
  it("places a child in the right age band", () => {
    const bands = (ar: number) => CalcBarnPerAlder(ar, 2000, 0, 0, 0);
    expect(bands(2000)).toMatchObject({ b1: 1 });
    expect(bands(2002)).toMatchObject({ b2: 1 });
    expect(bands(2003)).toMatchObject({ b3: 1 });
    expect(bands(2006)).toMatchObject({ b4: 1 });
    expect(bands(2010)).toMatchObject({ b5: 1 });
    expect(bands(2014)).toMatchObject({ b6: 1 });
    expect(bands(2018)).toMatchObject({ b7: 1 });
    expect(bands(2020)).toMatchObject({ b8: 1 });
  });

  it("counts nobody once the twenty-year window closes", () => {
    const total = (ar: number) =>
      Object.values(CalcBarnPerAlder(ar, 2000, 0, 0, 0)).reduce((a, b) => a + b, 0);
    expect(total(2020)).toBe(1);
    expect(total(2021)).toBe(0);
    expect(total(1999)).toBe(0);
  });

  it("adds children in the same band together", () => {
    expect(CalcBarnPerAlder(2010, 2000, 2001, 0, 0)).toMatchObject({ b5: 2 });
  });

  it("uses each child's own window, unlike CalcAntalBarn", () => {
    expect(CalcBarnPerAlder(2021, 0, 0, 2000, 2015)).toMatchObject({ b4: 1 });
  });
});

describe("barnbidraget", () => {
  it("pays 1 250 kr a month per child from 2019", () => {
    expect(barnbidraget(1, 2025)).toBe(1250 * 12);
    expect(barnbidraget(0, 2025)).toBe(0);
  });

  it("adds the large-family supplement from the second child", () => {
    expect(barnbidraget(2, 2025)).toBe(2 * 1250 * 12 + 150 * 12);
    expect(barnbidraget(3, 2025)).toBe(3 * 1250 * 12 + 150 * 12 + 580 * 12);
    expect(barnbidraget(4, 2025)).toBe(4 * 1250 * 12 + (150 + 580 + 1010) * 12);
  });

  it("holds the supplement flat past the fifth child", () => {
    const fifth = barnbidraget(5, 2025) - barnbidraget(4, 2025) - 1250 * 12;
    const sixth = barnbidraget(6, 2025) - barnbidraget(5, 2025) - 1250 * 12;
    expect(sixth).toBe(fifth);
    expect(sixth).toBe(1250 * 12);
  });

  it("blends the 2018 raise over the year it landed", () => {
    expect(barnbidraget(1, 2017)).toBe(1050 * 12);
    expect(barnbidraget(1, 2018)).toBe(1050 * 3 + 1250 * 9);
    expect(barnbidraget(1, 2019)).toBe(1250 * 12);
  });

  it("steps the supplement for a third child at 2017", () => {
    expect(barnbidraget(3, 2016) - 3 * 1050 * 12).toBe((150 + 454) * 12);
    expect(barnbidraget(3, 2017) - 3 * 1050 * 12).toBe((150 + 580) * 12);
  });

  it("accumulates the supplement year on year in the 1980s rules", () => {
    // 1983-84: 825 per child for the third, 1 650 for the fourth, and 1 650
    // more for each after that.
    expect(barnbidraget(5, 1983) - 5 * 825 * 4).toBe(825 + 1650 + 3300);
  });
});

describe("ustod", () => {
  it("is paid only to a single parent", () => {
    expect(ustod(2, 1, 2025)).toBe(2 * 1823 * 12);
    expect(ustod(2, 0, 2025)).toBe(0);
  });

  it("steps at 2006, 2015, 2017 and 2022", () => {
    expect(ustod(1, 1, 2005)).toBe(1173 * 12);
    expect(ustod(1, 1, 2006)).toBe(1273 * 12);
    expect(ustod(1, 1, 2015)).toBe(1573 * 12);
    expect(ustod(1, 1, 2017)).toBe(1673 * 12);
    expect(ustod(1, 1, 2022)).toBe(1823 * 12);
  });
});

describe("bobid", () => {
  /** A single parent in 2025 paying 8 000 kr a month for 80 square metres. */
  const single = (barn: number, inkomst = 200_000, over: Record<string, number> = {}) =>
    bobid(1, inkomst, 0, barn, over.uboende ?? 8000, over.yta ?? 80, 0, 0, over.year ?? 2025);

  it("QUIRK: pays nothing to a one-child family, because Array is 1-based", () => {
    // `Option Base 1` shifts `Array(0, xn1, xn2, xn3)` by one place, so a
    // one-child family reads 0 for every limit and for the särskilda bidrag.
    expect(single(1)).toBe(0);
    expect(single(1, 0)).toBe(0);
    expect(single(2)).toBeGreaterThan(0);
  });

  it("QUIRK: a two-child family is calculated on the one-child limits", () => {
    // 1 300 kr särskilt bidrag and half the cost from 1 400 up to 5 300.
    const beforeReduction = 12 * (1300 + (5300 - 1400) * 0.5);
    const reduced = beforeReduction - 0.2 * (200_000 - 150_000);
    expect(single(2)).toBe(12 * Math.trunc(reduced / 12 / 100) * 100);
  });

  it("QUIRK: four children are treated exactly as three", () => {
    // `min(3, ...)` caps the lookup, and the 140 m² limit is never reached.
    expect(single(4)).toBe(single(3));
    expect(single(5)).toBe(single(3));
  });

  it("reduces by 20 öre per krona above the income floor", () => {
    const step = single(3, 200_000) - single(3, 250_000);
    // 0.2 * 50 000 = 10 000 a year, then rounded down to whole hundreds a month.
    expect(step).toBeGreaterThan(9000);
    expect(step).toBeLessThanOrEqual(10_000);
  });

  it("pays nothing once income has eaten the whole allowance", () => {
    expect(single(3, 2_000_000)).toBe(0);
  });

  it("QUIRK: reads a housing cost over 50 000 as annual and divides it by twelve", () => {
    expect(single(3, 200_000, { uboende: 96_000 })).toBe(single(3));
    // The threshold is a cliff: 50 000 is still read as a monthly cost, which
    // clears every limit, while a krona more is divided down to 4 167 a month.
    expect(single(3, 200_000, { uboende: 50_000 })).toBeGreaterThan(
      single(3, 200_000, { uboende: 50_001 }),
    );
  });

  it("scales the cost down for a dwelling above the size limit", () => {
    expect(single(3, 200_000, { yta: 240 })).toBeLessThan(single(3, 200_000, { yta: 120 }));
    expect(single(3, 200_000, { yta: 120 })).toBe(single(3));
  });

  it("suppresses an allowance under 100 kr a month", () => {
    // Just above the floor the amount is rounded away rather than paid.
    const amounts = [];
    for (let inkomst = 300_000; inkomst <= 400_000; inkomst += 5000) {
      amounts.push(single(3, inkomst));
    }
    expect(amounts.filter((a) => a > 0 && a < 1200)).toHaveLength(0);
  });

  it("raises the allowance by an eighth in 2020", () => {
    // A quarter more for half the year, as the VBA assumes the proposal passed.
    expect(single(3, 200_000, { year: 2020 })).toBeGreaterThan(
      single(3, 200_000, { year: 2019 }),
    );
  });
});
