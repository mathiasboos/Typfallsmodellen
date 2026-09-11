import { describe, expect, it } from "vitest";

import {
  AgeArray,
  excelRound,
  vbaCLng,
  vbaFix,
  vbaInt,
  vbaIntDiv,
  vbaRound,
  vbaSingle,
  wsAverage,
  wsLarge,
  wsMax,
  wsMin,
} from "../src/vba/index.js";

describe("vbaInt", () => {
  it("rounds toward negative infinity, as VBA Int does", () => {
    expect(vbaInt(2.7)).toBe(2);
    expect(vbaInt(2.0)).toBe(2);
    expect(vbaInt(-2.7)).toBe(-3);
    expect(vbaInt(-2.1)).toBe(-3);
  });

  it("differs from Math.trunc for negatives, which is the whole point", () => {
    expect(vbaInt(-2.7)).not.toBe(Math.trunc(-2.7));
  });

  it("supports the model's rounding idiom Int(x / 12 + 0.5) * 12", () => {
    const roundToWholeMonths = (annual: number) => vbaInt(annual / 12 + 0.5) * 12;
    expect(roundToWholeMonths(24_005)).toBe(24_000);
    expect(roundToWholeMonths(24_007)).toBe(24_012);
  });
});

describe("vbaFix", () => {
  it("truncates toward zero", () => {
    expect(vbaFix(2.7)).toBe(2);
    expect(vbaFix(-2.7)).toBe(-2);
  });
});

describe("vbaRound", () => {
  it("rounds half to even, as VBA's intrinsic Round does", () => {
    expect(vbaRound(0.5)).toBe(0);
    expect(vbaRound(1.5)).toBe(2);
    expect(vbaRound(2.5)).toBe(2);
    expect(vbaRound(3.5)).toBe(4);
    expect(vbaRound(-2.5)).toBe(-2);
    expect(vbaRound(-3.5)).toBe(-4);
  });

  it("rounds non-halves to nearest", () => {
    expect(vbaRound(2.4)).toBe(2);
    expect(vbaRound(2.6)).toBe(3);
    expect(vbaRound(-2.6)).toBe(-3);
  });

  it("scales by a power of ten before testing the half, as the VBA runtime does", () => {
    // 2.675 is held as 2.67499999999999982, but multiplying by 100 rounds it up
    // to exactly 267.5, so the banker's rule applies and 267 is odd. A
    // decimal-aware implementation would return 2.67 instead. See the note on
    // vbaRound: this is unverified against a running VBA.
    expect(2.675 * 100).toBe(267.5);
    expect(vbaRound(2.675, 2)).toBe(2.68);
  });

  it("rounds at a given number of digits", () => {
    expect(vbaRound(3.14159, 2)).toBe(3.14);
    expect(vbaRound(1250, -2)).toBe(1200); // half to even: 12 is even
    expect(vbaRound(1350, -2)).toBe(1400); // half to even: 13 is odd, so up
  });
});

describe("excelRound", () => {
  it("rounds half away from zero, unlike vbaRound", () => {
    expect(excelRound(2.5)).toBe(3);
    expect(excelRound(3.5)).toBe(4);
    expect(excelRound(-2.5)).toBe(-3);
    expect(vbaRound(2.5)).not.toBe(excelRound(2.5));
  });

  it("rounds the decimal value Excel shows, so 2.675 at two digits gives 2.68", () => {
    expect(excelRound(2.675, 2)).toBe(2.68);
  });

  it("rounds to hundreds, as the basbelopp projections do", () => {
    expect(excelRound(59_249, -2)).toBe(59_200);
    expect(excelRound(59_250, -2)).toBe(59_300);
    expect(excelRound(58_777.4, -2)).toBe(58_800);
  });
});

describe("vbaCLng", () => {
  it("rounds rather than truncating, as assignment to Integer does", () => {
    expect(vbaCLng(2.5)).toBe(2);
    expect(vbaCLng(2.6)).toBe(3);
    expect(vbaCLng(-2.6)).toBe(-3);
  });
});

describe("vbaIntDiv", () => {
  it("truncates the quotient toward zero", () => {
    expect(vbaIntDiv(7, 2)).toBe(3);
    expect(vbaIntDiv(-7, 2)).toBe(-3);
  });

  it("rounds its operands to Long first", () => {
    expect(vbaIntDiv(7.6, 2)).toBe(4);
  });
});

describe("vbaSingle", () => {
  it("narrows to 32-bit float", () => {
    expect(vbaSingle(0.1)).toBeCloseTo(0.1, 7);
    expect(vbaSingle(0.1)).not.toBe(0.1);
  });

  it("leaves the year and count values the model stores as Single exact", () => {
    for (const year of [1957, 2021, 2026, 2110]) expect(vbaSingle(year)).toBe(year);
  });
});

describe("worksheet functions", () => {
  it("takes the largest and smallest", () => {
    expect(wsMax(1, 9, 3)).toBe(9);
    expect(wsMin(1, 9, 3)).toBe(1);
  });

  it("takes the k-th largest, 1-indexed, as Large does", () => {
    const values = [10, 50, 20, 40, 30];
    expect(wsLarge(values, 1)).toBe(50);
    expect(wsLarge(values, 3)).toBe(30);
    expect(wsLarge(values, 5)).toBe(10);
  });

  it("leaves the caller's array untouched", () => {
    const values = [10, 50, 20];
    wsLarge(values, 1);
    expect(values).toEqual([10, 50, 20]);
  });

  it("rejects a k outside the array", () => {
    expect(() => wsLarge([1, 2], 3)).toThrow(RangeError);
  });

  it("averages", () => {
    expect(wsAverage([10, 20, 30])).toBe(20);
  });
});

describe("AgeArray", () => {
  it("indexes by age, not from zero", () => {
    const ages = new AgeArray(15, 105);
    ages.set(66, 1234);
    expect(ages.get(66)).toBe(1234);
    expect(ages.get(15)).toBe(0);
  });

  it("throws outside its bounds rather than yielding undefined", () => {
    const ages = new AgeArray(15, 105);
    expect(() => ages.get(14)).toThrow(RangeError);
    expect(() => ages.get(106)).toThrow(RangeError);
    expect(() => ages.set(14, 1)).toThrow(RangeError);
  });

  it("reads one below the first age as zero, for the x(age - 1) idiom", () => {
    const ages = new AgeArray(15, 105);
    expect(ages.getOrZero(14)).toBe(0);
    expect(ages.getOrZero(200)).toBe(0);
  });

  it("accumulates and slices", () => {
    const ages = new AgeArray(15, 20);
    ages.add(16, 5);
    ages.add(16, 7);
    expect(ages.get(16)).toBe(12);
    expect(ages.slice(15, 17)).toEqual([0, 12, 0]);
  });

  it("iterates age and value together", () => {
    const ages = new AgeArray(60, 62);
    ages.set(61, 9);
    expect([...ages.entries()]).toEqual([
      [60, 0],
      [61, 9],
      [62, 0],
    ]);
  });
});
