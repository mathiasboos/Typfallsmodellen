import { describe, expect, it } from "vitest";

import { conscriptionDaysByYear, conscriptionPgb, studyPgb } from "../src/model/pgb.js";

describe("conscriptionDaysByYear", () => {
  it("credits nothing under the 120-day minimum", () => {
    const days = conscriptionDaysByYear({ start: "2005-01-01", end: "2005-03-01" });
    expect(days.size).toBe(0);
  });

  it("credits a whole period within one calendar year to that year alone", () => {
    const days = conscriptionDaysByYear({ start: "2005-01-15", end: "2005-08-20" });
    expect([...days.keys()]).toEqual([2005]);
    // H5-H4 exactly, since the whole period sits inside one year's own bound.
    expect(days.get(2005)).toBe((Date.UTC(2005, 7, 20) - Date.UTC(2005, 0, 15)) / 86_400_000);
  });

  it("splits a period crossing one year boundary between both years", () => {
    const days = conscriptionDaysByYear({ start: "2005-10-01", end: "2006-03-01" });
    expect([...days.keys()]).toEqual([2005, 2006]);
    expect(days.get(2005)).toBe(92); // Oct 1 - Dec 31 2005, inclusive
    expect(days.get(2006)).toBe(60); // Jan 1 - Mar 1 2006, inclusive
  });

  it("splits a period crossing two year boundaries across all three years, capped at 730 days", () => {
    const days = conscriptionDaysByYear({ start: "2005-06-01", end: "2007-06-01" });
    expect([...days.keys()]).toEqual([2005, 2006, 2007]);
    const total = [...days.values()].reduce((a, b) => a + b, 0);
    expect(total).toBeLessThanOrEqual(730);
    expect(days.get(2006)).toBe(365); // a full year in the middle
  });

  it("gives nothing for a reversed or empty period", () => {
    expect(conscriptionDaysByYear({ start: "2005-08-20", end: "2005-01-15" }).size).toBe(0);
    expect(conscriptionDaysByYear({ start: "2005-01-15", end: "2005-01-15" }).size).toBe(0);
  });
});

describe("conscriptionPgb", () => {
  it("is zero outside the years conscription actually existed (2011-2017)", () => {
    expect(conscriptionPgb(2015, 200, 300_000, 0)).toBe(0);
  });

  it("is zero before 1995 even with real days and a real medelPgi", () => {
    expect(conscriptionPgb(1990, 200, 100_000, 0)).toBe(0);
  });

  it("is half of medelPgi, pro-rated by days over 365, inside the eligible window", () => {
    const exact = conscriptionPgb(1998, 365, 200_000, 1);
    expect(exact).toBeCloseTo(0.5 * 200_000, 6);
  });

  it("rounds down to the nearest 100 kr unless marginal removes it", () => {
    const rounded = conscriptionPgb(1998, 100, 199_999, 0);
    const exact = conscriptionPgb(1998, 100, 199_999, 1);
    expect(rounded % 100).toBe(0);
    expect(rounded).toBeLessThanOrEqual(exact);
    expect(exact - rounded).toBeLessThan(100);
  });

  it("is eligible again from 2018 on", () => {
    expect(conscriptionPgb(2018, 200, 300_000, 1)).toBeGreaterThan(0);
    expect(conscriptionPgb(2010, 200, 300_000, 1)).toBeGreaterThan(0);
  });
});

describe("studyPgb", () => {
  it("is zero before 1995, when studies did not yet earn PGB", () => {
    expect(studyPgb(1990, 2, 0)).toBe(0);
    expect(studyPgb(1980, 1, 0)).toBe(0);
  });

  it("is positive from 1995 on, for a typed semester count", () => {
    expect(studyPgb(1998, 1, 1)).toBeGreaterThan(0);
    expect(studyPgb(1998, 2, 1)).toBeGreaterThan(studyPgb(1998, 1, 1));
  });

  it("scales linearly with the number of semesters", () => {
    const one = studyPgb(2010, 1, 1);
    const two = studyPgb(2010, 2, 1);
    expect(two).toBeCloseTo(one * 2, 6);
  });

  it("rounds down to the nearest 100 kr unless marginal removes it", () => {
    const rounded = studyPgb(2010, 1, 0);
    const exact = studyPgb(2010, 1, 1);
    expect(rounded % 100).toBe(0);
    expect(rounded).toBeLessThanOrEqual(exact);
    expect(exact - rounded).toBeLessThan(100);
  });

  it("holds the last real year's rate flat for a year past the extracted range", () => {
    const atEdge = studyPgb(2110, 1, 1);
    const wellPast = studyPgb(2200, 1, 1);
    expect(wellPast).toBe(atEdge);
  });
});
