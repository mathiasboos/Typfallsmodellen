import { describe, expect, it } from "vitest";

import { gp, tillagg } from "../src/bidrag/garantipension.js";

/**
 * No offline fixture covers the benefits module -- the Brutto sheet inside the
 * workbook reaches only the earning phase of the public pension. So these tests
 * assert bracket boundaries, the year transitions the VBA states outright, and
 * the quirks worth pinning down, and never claim an amount is the legally
 * correct one. The golden files from a real Excel run are the judge of that.
 */

const PBB = 58_800;

/** `gp` with the rounding off, so the raw formula is visible. */
const raw = (...args: Parameters<typeof gp>) => gp(...args);

describe("gp, cohorts born 1938 and later", () => {
  it("pays nothing before the riktålder", () => {
    expect(raw(0, 0, 1958, PBB, 40, 1, 65, 2025, 21_000, 0, 1, 66)).toBe(0);
    expect(raw(0, 0, 1958, PBB, 40, 1, 66, 2025, 21_000, 0, 1, 66)).toBeGreaterThan(0);
  });

  it("pays the full single amount at zero income", () => {
    expect(raw(0, 0, 1958, PBB, 40, 1, 66, 2025, 21_000, 0, 1, 66)).toBeCloseTo(2.43 * PBB, 9);
  });

  it("pays less to a cohabitant than to a single person", () => {
    const single = raw(0, 0, 1958, PBB, 40, 1, 66, 2025, 21_000, 0, 1, 66);
    const couple = raw(0, 1, 1958, PBB, 40, 1, 66, 2025, 21_000, 0, 1, 66);
    expect(couple).toBeCloseTo(2.2 * PBB, 9);
    expect(couple).toBeLessThan(single);
  });

  it("switches to the shallower taper above 1.26 pbb for a single person", () => {
    const below = raw(1.25 * PBB, 0, 1958, PBB, 40, 1, 66, 2025, 21_000, 0, 1, 66);
    const above = raw(1.27 * PBB, 0, 1958, PBB, 40, 1, 66, 2025, 21_000, 0, 1, 66);
    // Krona for krona below the breakpoint, 48 öre above it.
    expect(below).toBeCloseTo(2.43 * PBB - 1.25 * PBB, 9);
    expect(above).toBeCloseTo(1.17 * PBB - 0.48 * (1.27 * PBB - 1.26 * PBB), 9);
  });

  it("puts the single breakpoint at 1.26 pbb and the cohabiting one at 1.14", () => {
    // `<=` in the VBA, so the breakpoint itself is on the steep side.
    expect(raw(1.26 * PBB, 0, 1958, PBB, 40, 1, 66, 2025, 21_000, 0, 1, 66)).toBeCloseTo(
      2.43 * PBB - 1.26 * PBB,
      9,
    );
    expect(raw(1.14 * PBB, 0, 1958, PBB, 40, 1, 66, 2025, 21_000, 0, 1, 66)).not.toBeCloseTo(
      2.2 * PBB - 1.14 * PBB,
      9,
    );
    expect(raw(1.14 * PBB, 1, 1958, PBB, 40, 1, 66, 2025, 21_000, 0, 1, 66)).toBeCloseTo(
      2.2 * PBB - 1.14 * PBB,
      9,
    );
  });

  it("steps the single base amount at 2020, 2022 and 2023", () => {
    const at = (year: number) => raw(0, 0, 1958, PBB, 40, 1, 66, year, 21_000, 0, 1, 66);
    expect(at(2019)).toBeCloseTo(2.13 * PBB, 9);
    expect(at(2020)).toBeCloseTo(2.181 * PBB, 9);
    expect(at(2021)).toBeCloseTo(2.181 * PBB, 9);
    // 2022 blends seven months of the old level with five of the new.
    expect(at(2022)).toBeCloseTo(((2.181 * 7 + 2.43 * 5) * PBB) / 12, 9);
    expect(at(2023)).toBeCloseTo(2.43 * PBB, 9);
  });

  it("scales by försäkringstid out of 40 years", () => {
    const full = raw(0, 0, 1958, PBB, 40, 1, 66, 2025, 21_000, 0, 1, 66);
    expect(raw(0, 0, 1958, PBB, 20, 1, 66, 2025, 21_000, 0, 1, 66)).toBeCloseTo(full / 2, 9);
  });

  it("QUIRK: requires four years of residence, though its comment says three", () => {
    expect(raw(0, 0, 1958, PBB, 3, 1, 66, 2025, 21_000, 0, 1, 66)).toBe(0);
    expect(raw(0, 0, 1958, PBB, 4, 1, 66, 2025, 21_000, 0, 1, 66)).toBeCloseTo(
      (2.43 * PBB * 4) / 40,
      9,
    );
  });

  it("scales by the withdrawal share only when it is below a full one", () => {
    const full = raw(0, 0, 1958, PBB, 40, 1, 66, 2025, 21_000, 0, 1, 66);
    expect(raw(0, 0, 1958, PBB, 40, 1, 66, 2025, 21_000, 0, 1, 66, 0.5)).toBeCloseTo(full / 2, 9);
    expect(raw(0, 0, 1958, PBB, 40, 1, 66, 2025, 21_000, 0, 1, 66, 1)).toBe(full);
  });

  it("rounds to a whole krona a month when marginal is 0", () => {
    const rounded = gp(1.3 * PBB, 0, 1958, PBB, 40, 0, 66, 2025, 21_000, 0, 1, 66);
    const exact = gp(1.3 * PBB, 0, 1958, PBB, 40, 1, 66, 2025, 21_000, 0, 1, 66);
    expect(rounded % 12).toBe(0);
    expect(Math.abs(rounded - exact)).toBeLessThan(12);
  });

  it("QUIRK: swaps pbb for IBB only past wyear -- strictly after, unlike BTP", () => {
    const args = [0, 0, 1958, PBB, 40, 1, 66] as const;
    const atWyear = gp(...args, 2030, 2030, 70_000, 0.8, 66);
    const pastWyear = gp(...args, 2031, 2030, 70_000, 0.8, 66);
    expect(atWyear).toBeCloseTo(2.43 * PBB, 9);
    expect(pastWyear).toBeCloseTo(2.43 * 70_000 * 0.8, 9);
  });
});

describe("gp, cohorts born 1937 and earlier", () => {
  const OLD = 1930;

  it("builds the amount up from income rather than reducing it down", () => {
    // Below 0.25 pbb the old rules give 104% of income, so the net top-up is 4%.
    expect(raw(0.1 * PBB, 0, OLD, PBB, 40, 1, 66, 2005, 21_000, 0, 1, 66)).toBeCloseTo(
      0.1 * PBB * 1.04 - 0.1 * PBB,
      9,
    );
  });

  it("pays nothing above the top knot", () => {
    expect(raw(3.2 * PBB, 0, OLD, PBB, 40, 1, 66, 2005, 21_000, 0, 1, 66)).toBe(0);
    expect(raw(2.9 * PBB, 1, OLD, PBB, 40, 1, 66, 2005, 21_000, 0, 1, 66)).toBe(0);
  });

  it("QUIRK: the 2020 supplement enters income for a single person but not a cohabitant", () => {
    // Single: income is raised by 0.051 pbb before the brackets, so the 1.04
    // factor applies to the supplement too.
    expect(raw(0, 0, OLD, PBB, 40, 1, 66, 2021, 21_000, 0, 1, 66)).toBeCloseTo(
      0.051 * PBB * 1.04,
      9,
    );
    // Cohabiting: income is untouched, so the supplement passes through flat.
    expect(raw(0, 1, OLD, PBB, 40, 1, 66, 2021, 21_000, 0, 1, 66)).toBeCloseTo(0.051 * PBB, 9);
  });

  it("QUIRK: has no four-year residence floor, unlike the later cohorts", () => {
    expect(raw(0.1 * PBB, 0, OLD, PBB, 3, 1, 66, 2005, 21_000, 0, 1, 66)).toBeGreaterThan(0);
  });
});

describe("tillagg", () => {
  const at = (underl: number, marginal = 0, born = 1958, ftid = 40) =>
    tillagg(underl, 2021, born, 186.52, 186.52, 1, ftid, marginal);

  it("does not exist before 2021", () => {
    expect(tillagg(150_000, 2020, 1958, 186.52, 186.52)).toBe(0);
    expect(tillagg(150_000, 2021, 1958, 186.52, 186.52)).toBeGreaterThan(0);
  });

  it("needs at least one year of residence", () => {
    expect(tillagg(150_000, 2021, 1958, 186.52, 186.52, 1, 0)).toBe(0);
  });

  it("pays nothing outside 108 000 to 204 000", () => {
    expect(at(107_999)).toBe(0);
    expect(at(108_000)).toBeGreaterThan(0);
    expect(at(203_999)).toBeGreaterThan(0);
    expect(at(204_000)).toBe(0);
  });

  it("rises, plateaus at 7 200 and tapers away", () => {
    expect(at(120_000, 1)).toBeCloseTo((120_000 - 108_000) * 0.3, 4);
    expect(at(150_000, 1)).toBe(7200);
    expect(at(180_000, 1)).toBeCloseTo(7200 - (180_000 - 168_000) * 0.2, 4);
  });

  it("QUIRK: the 600-krona rounding adds 300, so it can exceed the exact amount", () => {
    expect(at(120_000, 1)).toBe(3600);
    expect(at(120_000, 0)).toBe(3900);
    // The taper rounds the same way, down to a 300-krona residue at the very top.
    expect(at(180_000, 0)).toBe(5100);
    expect(at(203_999, 0)).toBe(300);
  });

  it("QUIRK: the flat middle bracket is not rounded at all", () => {
    expect(at(150_000, 0)).toBe(7200);
    expect(at(150_000, 1)).toBe(7200);
  });

  it("deflates the brackets by inkomstindex net of the 1.6% norm", () => {
    // Index up by exactly the norm leaves the real amount unchanged.
    expect(tillagg(150_000 * 1.016, 2022, 1958, 186.52 * 1.016, 186.52)).toBe(7200);
    // The exponent compounds, so two years of the norm behave the same.
    expect(tillagg(150_000 * 1.016 ** 2, 2023, 1958, 186.52 * 1.016 ** 2, 186.52)).toBe(7200);
  });

  it("varies the years needed for a full supplement by cohort", () => {
    const full = 7200;
    expect(at(150_000, 1, 1950, 20)).toBeCloseTo((full * 20) / 40, 3);
    expect(at(150_000, 1, 1940, 20)).toBeCloseTo((full * 20) / 35, 3);
    expect(at(150_000, 1, 1930, 20)).toBeCloseTo((full * 20) / 30, 3);
    // Born 1915-1924: 20 years plus one for each year after 1915.
    expect(at(150_000, 1, 1920, 20)).toBeCloseTo((full * 20) / 25, 3);
    expect(at(150_000, 1, 1910, 10)).toBeCloseTo((full * 10) / 20, 3);
  });

  it("QUIRK: returns a 32-bit float, so the cohort scaling narrows", () => {
    const scaled = at(150_000, 1, 1940, 20);
    expect(scaled).toBe(Math.fround((7200 * 20) / 35));
    expect(scaled).not.toBe((7200 * 20) / 35);
  });
});
