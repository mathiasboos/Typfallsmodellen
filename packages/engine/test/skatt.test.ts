import { describe, expect, it } from "vitest";

import { defaultContext } from "../src/model/context.js";
import type { RunVectors } from "../src/model/runVectors.js";
import {
  arbgiv,
  avdrag22,
  avdrag26,
  avdragxx,
  avkskatt,
  FAared,
  Jobb26,
  Jobbxx,
  pandred,
  PublicAvg,
  sared,
  statlig,
  Xage,
} from "../src/skatt/index.js";

/**
 * The grundavdrag and jobbskatteavdrag functions are translated mechanically
 * from the VBA, and `npm run check:transpile` verifies the committed code still
 * matches its source. These tests cover what that check cannot: that the
 * dispatchers pick the right year, and that the results behave sensibly.
 *
 * As with the occupational agreements, no offline fixture reaches this layer, so
 * nothing here claims an amount is the legally correct one.
 */
const model = defaultContext();
const PBB = 58_800;
const IBB = 76_200;

const vectors: RunVectors = {
  wage: () => 462_000,
  ibb: () => IBB,
  pbb: () => PBB,
  fpb: () => 60_000,
  kpiJune: () => 400,
  kpi: () => 400,
  year: (age) => 1959 + age,
};

describe("Xage", () => {
  it("rises with riktålder and is projected beyond 2028", () => {
    expect(Xage(2023, model)).toBe(66);
    expect(Xage(2026, model)).toBe(67);
    expect(Xage(2028, model)).toBe(68);
    expect(Xage(2035, model)).toBe(model.riktage + 1);
  });
});

describe("grundavdrag", () => {
  it("never exceeds the income it is deducted from", () => {
    for (let income = 0; income <= 2_000_000; income += 25_000) {
      for (const year of [1995, 2005, 2015, 2024, 2026]) {
        const avdrag = avdragxx(income, PBB, 0, 40, year);
        expect(avdrag, `${year} at ${income}`).toBeLessThanOrEqual(income);
        expect(avdrag, `${year} at ${income}`).toBeGreaterThanOrEqual(0);
      }
    }
  });

  it("rounds up to whole hundreds unless marginal turns rounding off", () => {
    expect(avdragxx(437_651, PBB, 0, 40, 2024) % 100).toBe(0);
    expect(avdragxx(437_651, PBB, 1, 40, 2024) % 100).not.toBe(0);
  });

  it("gives a far larger allowance at and above Xage", () => {
    const income = 300_000;
    const young = avdrag26(income, PBB, 0, 50, 2026, 2100, 0, 1, 0, 67);
    const old = avdrag26(income, PBB, 0, 67, 2026, 2100, 0, 1, 0, 67);
    expect(old).toBeGreaterThan(young);
  });

  it("applies the higher allowance exactly at Xage, not before", () => {
    const income = 300_000;
    const at66 = avdrag26(income, PBB, 0, 66, 2026, 2100, 0, 1, 0, 67);
    const at67 = avdrag26(income, PBB, 0, 67, 2026, 2100, 0, 1, 0, 67);
    expect(at66).toBeLessThan(at67);
  });

  it("dispatches to the right rules for a year", () => {
    const args = [300_000, PBB, 0, 40, 2023, 21_000, 0, 1, 2023, 66] as const;
    expect(avdragxx(...args)).toBe(avdrag22(...args));
  });

  it("lets Iyear pin the rules while the income year moves", () => {
    // The "unchanged rules" assumption: 2040 income, 2022 rules.
    const pinned = avdragxx(300_000, PBB, 0, 40, 2040, 21_000, 0, 1, 2023, 66);
    expect(pinned).toBe(avdrag22(300_000, PBB, 0, 40, 2040, 21_000, 0, 1, 2023, 66));
  });

  it("is finite and sensible across every year and income the model reaches", () => {
    for (let year = 1990; year <= 2040; year += 1) {
      for (const income of [0, 50_000, 200_000, 500_000, 1_000_000]) {
        for (const age of [40, 70]) {
          const avdrag = avdragxx(income, PBB, 0, age, year, 21_000, 0, 1, 0, 67);
          expect(Number.isFinite(avdrag), `${year} age ${age} at ${income}`).toBe(true);
          expect(avdrag, `${year} age ${age} at ${income}`).toBeGreaterThanOrEqual(0);
        }
      }
    }
  });
});

describe("jobbskatteavdrag", () => {
  it("did not exist before 2007", () => {
    expect(Jobbxx(400_000, 40, 0.32, PBB, 0, 0, 2006)).toBe(0);
    expect(Jobbxx(400_000, 40, 0.32, PBB, 0, 0, 2007)).toBeGreaterThan(0);
  });

  it("is never negative", () => {
    for (let income = 0; income <= 1_500_000; income += 25_000) {
      for (const year of [2007, 2012, 2019, 2024, 2026]) {
        const credit = Jobbxx(income, 40, 0.32, PBB, 0, 0, year);
        expect(credit, `${year} at ${income}`).toBeGreaterThanOrEqual(0);
        expect(Number.isFinite(credit), `${year} at ${income}`).toBe(true);
      }
    }
  });

  it("rises, plateaus, then tapers away for high earners", () => {
    const at = (income: number) => Jobbxx(income, 40, 0.32, PBB, 0, 0, 2024);

    // Rising through the low and middle bands.
    expect(at(300_000)).toBeGreaterThan(at(150_000));
    expect(at(475_000)).toBeGreaterThan(at(300_000));

    // A plateau between 8.08 and 13.54 price base amounts.
    expect(at(600_000)).toBe(at(700_000));

    // Then the taper that phases the credit out for high earners, at 3% of
    // income above 13.54 base amounts.
    const taperStart = 13.54 * PBB;
    expect(at(taperStart + 200_000)).toBeLessThan(at(taperStart));
    expect(at(2_000_000)).toBeLessThan(at(1_000_000));
  });

  it("gives a larger credit past Xage", () => {
    const income = 300_000;
    const young = Jobb26(income, 50, 0.32, PBB, 0, 0, 2026, 21_000, 0, 1, 0, 67);
    const old = Jobb26(income, 70, 0.32, PBB, 0, 0, 2026, 21_000, 0, 1, 0, 67);
    expect(old).toBeGreaterThan(young);
  });

  it("dispatches to the right rules for a year", () => {
    const args = [400_000, 40, 0.32, PBB, 0, 0, 2026, 21_000, 0, 1, 2026, 67] as const;
    expect(Jobbxx(...args)).toBe(Jobb26(...args));
  });
});

describe("statlig", () => {
  it("is nothing below the threshold", () => {
    expect(statlig(500_000, 625_800, 1e16)).toBe(0);
  });

  it("takes 20% above the threshold", () => {
    expect(statlig(725_800, 625_800, 1e16)).toBe(Math.floor(0.2 * 100_000));
  });

  it("takes 25% above the second threshold, where one still applies", () => {
    const tax = statlig(800_000, 500_000, 700_000);
    expect(tax).toBe(Math.floor(0.25 * 100_000 + 0.2 * 200_000));
  });
});

describe("PublicAvg", () => {
  it("did not exist before 2019", () => {
    expect(PublicAvg(400_000, 1959, vectors, 0.01, 40, 0, 2018)).toBe(0);
  });

  it("is 1% of income up to a ceiling that falls each year", () => {
    const at = (year: number) => PublicAvg(1_000_000, 1959, vectors, 0.01, 40, 0, year);
    expect(at(2022)).toBeGreaterThan(at(2023));
    expect(at(2023)).toBeGreaterThan(at(2024));
    expect(at(2026)).toBeLessThan(at(2025));
  });

  it("is not charged to those 18 or under", () => {
    expect(PublicAvg(400_000, 2005, vectors, 0.01, 18, 0, 2024)).toBe(0);
  });
});

describe("tax reductions", () => {
  it("gives no sickness-compensation reduction before 2018", () => {
    expect(sared(200_000, 2017)).toBe(0);
    expect(sared(200_000, 2024)).not.toBe(0);
  });

  it("gives no earned-income reduction before 2021, then caps it at 1 500", () => {
    expect(FAared(300_000, 2020, 0)).toBe(0);
    expect(FAared(300_000, 2024, 0)).toBe(1500);
    expect(FAared(100_000, 2024, 0)).toBeLessThan(1500);
  });

  it("gives the pandemic reduction only in 2021 to 2023 and only in a salary band", () => {
    expect(pandred(300_000, 2020, 0)).toBe(0);
    expect(pandred(300_000, 2024, 0)).toBe(0);
    expect(pandred(50_000, 2022, 0)).toBe(0);
    expect(pandred(600_000, 2022, 0)).toBe(0);
    expect(pandred(300_000, 2022, 0)).toBeGreaterThan(0);
  });
});

describe("arbgiv", () => {
  const rates = { full: 31.42, oldAge: 10.21 };

  it("charges nothing on a negligible wage or before 1960", () => {
    expect(arbgiv(500, 2024, 40, rates, 66)).toBe(0);
    expect(arbgiv(400_000, 1959, 40, rates, 66)).toBe(0);
  });

  it("charges only the old-age component at and above riktålder", () => {
    const below = arbgiv(400_000, 2024, 40, rates, 66);
    const above = arbgiv(400_000, 2024, 66, rates, 66);
    expect(above).toBeLessThan(below);
  });

  it("returns the rate itself when asked for it", () => {
    expect(arbgiv(400_000, 2024, 40, rates, 66, 0, 1)).toBe(rates.oldAge);
  });
});

describe("avkskatt", () => {
  it("taxes pension insurance at 15% of the borrowing rate, with a floor from 2017", () => {
    // 2021's rate came from a negative figure at the end of 2020, so the floor
    // of 0.5% applies.
    expect(avkskatt(2021, model, 0)).toBeCloseTo(0.005 * 0.15, 9);
  });

  it("taxes capital insurance at 30% of the rate plus a point", () => {
    expect(avkskatt(2024, model, 1)).toBeCloseTo((0.01 + 0.0262) * 0.3, 9);
  });

  it("forecasts 2.5% beyond the table", () => {
    expect(avkskatt(2040, model, 1)).toBeCloseTo((0.01 + 0.025) * 0.3, 9);
  });
});
