import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { defaultContext } from "../src/model/context.js";
import {
  andel,
  gpavgift,
  ipavgift,
  pgi,
  ppavgift,
  PgiResult,
  riktage,
} from "../src/pension/contributions.js";

/**
 * The Brutto sheet's cached values. Its formulas call PGI, IPavgift, ppavgift
 * and GPAVGIFT as worksheet UDFs, so these are the original functions' own
 * results for ~80 years of one typfall -- an earner above the ceiling
 * throughout, born 2000.
 */
interface BruttoRow {
  age: number;
  year: number;
  income: number | null;
  inkomstbasbelopp: number | null;
  forhojtPrisbasbelopp: number | null;
  prisbasbelopp: number | null;
  pgi: number | null;
  egenavgift: number | null;
  ipavgift: number | null;
  ppavgift: number | null;
  gpavgift: number | null;
  pgb: number | null;
}

const fixture = JSON.parse(
  readFileSync(
    fileURLToPath(new URL("../../../reference/fixtures/brutto-cached.json", import.meta.url)),
    "utf8",
  ),
) as { rows: BruttoRow[]; shareOfNewSystem: number };

const context = defaultContext();

/** Rows where the sheet has both the inputs and a computed PGI to compare. */
const rows = fixture.rows.filter(
  (r) =>
    r.income !== null &&
    r.inkomstbasbelopp !== null &&
    r.prisbasbelopp !== null &&
    r.forhojtPrisbasbelopp !== null &&
    r.pgi !== null &&
    r.year >= 1960,
);

describe("riktage", () => {
  it("steps the earliest withdrawal age and riktålder together", () => {
    expect(riktage(2019, 0)).toBe(61);
    expect(riktage(2019, 1)).toBe(65);
    expect(riktage(2023, 0)).toBe(63);
    expect(riktage(2023, 1)).toBe(66);
    expect(riktage(2026, 0)).toBe(64);
    expect(riktage(2026, 1)).toBe(67);
    expect(riktage(2100, 1)).toBe(70);
  });
});

describe("andel", () => {
  it("gives cohorts before 1938 no share of the new system", () => {
    expect(andel(1930)).toBe(0);
    expect(andel(1937)).toBe(0);
  });

  it("phases the new system in a twentieth at a time from 1938", () => {
    expect(andel(1938)).toBeCloseTo(4 / 20, 12);
    expect(andel(1944)).toBeCloseTo(10 / 20, 12);
    expect(andel(1953)).toBeCloseTo(19 / 20, 12);
  });

  it("gives cohorts after 1953 the new system in full", () => {
    expect(andel(1954)).toBe(1);
    expect(andel(1990)).toBe(1);
  });

  it("truncates a fractional birth year", () => {
    expect(andel(1954.7)).toBe(andel(1954));
  });
});

describe("against the Brutto sheet's cached values", () => {
  it("has rows to compare", () => {
    expect(rows.length).toBeGreaterThan(50);
  });

  // What this fixture does and does not reach. Its typfall earns above the
  // ceiling in every year from 2016, so it pins down the ceiling, the
  // contribution and the rounding thoroughly -- removing the final 7.5 IBB
  // clamp breaks it immediately -- but it never exercises an income below the
  // ceiling, nor any year before 2016. In particular the `As Long` quirk in the
  // tax-reduction share only bites in 2000-2005 and is NOT covered here; the
  // golden files from a real Excel run are what will settle it.
  it("only covers incomes at the ceiling, from 2016 on", () => {
    expect(Math.min(...rows.map((r) => r.year))).toBeGreaterThanOrEqual(2016);
    const working = rows.filter((r) => r.income! > 0);
    expect(working.every((r) => r.pgi === 7.5 * r.inkomstbasbelopp!)).toBe(true);
  });

  it("reproduces the pension-qualifying income", () => {
    const mismatches = rows
      .filter(
        (r) =>
          pgi(
            r.year,
            r.income!,
            r.prisbasbelopp!,
            r.inkomstbasbelopp!,
            r.forhojtPrisbasbelopp!,
            context,
            PgiResult.Income,
            r.age,
            0,
          ) !== r.pgi,
      )
      .map((r) => {
        const got = pgi(
          r.year,
          r.income!,
          r.prisbasbelopp!,
          r.inkomstbasbelopp!,
          r.forhojtPrisbasbelopp!,
          context,
          PgiResult.Income,
          r.age,
          0,
        );
        return `${r.year}: got ${got}, workbook has ${r.pgi}`;
      });
    expect(mismatches.slice(0, 5).join("\n")).toBe("");
  });

  it("reproduces the general pension contribution", () => {
    const mismatches = rows
      .filter((r) => r.egenavgift !== null)
      .map((r) => {
        const got = pgi(
          r.year,
          r.income!,
          r.prisbasbelopp!,
          r.inkomstbasbelopp!,
          r.forhojtPrisbasbelopp!,
          context,
          PgiResult.Contribution,
          r.age,
          0,
        );
        return got === r.egenavgift ? null : `${r.year}: got ${got}, workbook has ${r.egenavgift}`;
      })
      .filter((m): m is string => m !== null);
    expect(mismatches.slice(0, 5).join("\n")).toBe("");
  });

  it("reproduces the income, premium and garantipension contributions", () => {
    // Brutto computes each year's contribution from the *previous* row's PGI
    // plus PGB, which is how the pension right lands a year after assessment.
    const byYear = new Map(rows.map((r) => [r.year, r]));
    const mismatches: string[] = [];

    for (const row of rows) {
      const previous = byYear.get(row.year - 1);
      if (previous === undefined || row.ipavgift === null) continue;
      const base = previous.pgi! + (previous.pgb ?? 0);

      const checks: [string, number, number | null][] = [
        [
          "ipavgift",
          ipavgift(previous.year, base, previous.age, fixture.shareOfNewSystem, context.marginal),
          row.ipavgift,
        ],
        [
          "ppavgift",
          ppavgift(previous.year, base, previous.age, fixture.shareOfNewSystem, context.marginal),
          row.ppavgift,
        ],
        [
          "gpavgift",
          gpavgift(previous.year, base, row.age, fixture.shareOfNewSystem, context.marginal),
          row.gpavgift,
        ],
      ];
      for (const [name, got, want] of checks) {
        if (want === null) continue;
        if (Math.abs(got - want) > 1e-9) {
          mismatches.push(`${row.year} ${name}: got ${got}, workbook has ${want}`);
        }
      }
    }
    expect(mismatches.slice(0, 8).join("\n")).toBe("");
  });
});

/**
 * Structural properties of pgi that the Brutto fixture cannot reach, since it
 * only ever feeds an income above the ceiling.
 *
 * These assert shape, not amounts: no figure here is claimed to be the legally
 * correct one, because nothing available offline can confirm that. They exist to
 * catch regressions in the paths the fixture leaves untested.
 */
describe("pgi below and around the ceiling", () => {
  const year = 2024;
  const pbb = 57_300;
  const ibb = 76_200;
  const fhb = 58_500;
  const of = (income: number, typ = PgiResult.Income) =>
    pgi(year, income, pbb, ibb, fhb, context, typ, 40, 0);

  it("gives nothing below the floor", () => {
    expect(of(0)).toBe(0);
    expect(of(Math.floor(0.423 * pbb) - 1000)).toBe(0);
  });

  it("rises with income between floor and ceiling, and then stops", () => {
    expect(of(300_000)).toBeGreaterThan(of(200_000));
    expect(of(500_000)).toBeGreaterThan(of(300_000));
    expect(of(2_000_000)).toBe(of(1_000_000));
  });

  it("never exceeds 7.5 income base amounts", () => {
    for (const income of [500_000, 700_000, 1_000_000, 5_000_000]) {
      expect(of(income)).toBeLessThanOrEqual(7.5 * ibb);
    }
  });

  it("deducts the 7% contribution below the ceiling", () => {
    const income = 400_000;
    const contribution = of(income, PgiResult.Contribution);
    expect(contribution).toBeGreaterThan(0.06 * income);
    expect(contribution).toBeLessThan(0.08 * income);
    expect(of(income)).toBeLessThan(income);
  });

  it("rounds to whole hundreds unless marginal turns rounding off", () => {
    expect(of(400_050) % 100).toBe(0);
    const unrounded = pgi(
      year,
      400_055,
      pbb,
      ibb,
      fhb,
      defaultContext({ marginal: 1 }),
      PgiResult.Income,
      40,
      0,
    );
    expect(unrounded % 100).not.toBe(0);
  });
});

describe("contribution rates", () => {
  it("pays nothing for cohorts born before 1938", () => {
    expect(ipavgift(2024, 400_000, 50, 1, 0, 1937)).toBe(0);
    expect(ppavgift(2024, 400_000, 50, 1, 0, 1937)).toBe(0);
    expect(gpavgift(2024, 400_000, 50, 1, 0, 1937)).toBe(0);
  });

  it("applies 16% to the income pension and 2.5% to the premium pension today", () => {
    expect(ipavgift(2024, 400_000, 50, 1, 1)).toBeCloseTo(0.16 * 400_000, 9);
    expect(ppavgift(2024, 400_000, 50, 1, 1)).toBeCloseTo(0.025 * 400_000, 9);
  });

  it("credits the whole contribution to the new system from age 65", () => {
    const half = 0.5;
    expect(ipavgift(2024, 400_000, 64, half, 1)).toBeCloseTo(0.16 * 400_000 * half, 9);
    expect(ipavgift(2024, 400_000, 65, half, 1)).toBeCloseTo(0.16 * 400_000, 9);
  });

  it("rounds down to whole kronor unless marginal turns rounding off", () => {
    expect(Number.isInteger(ipavgift(2024, 400_123, 50, 1, 0))).toBe(true);
    expect(Number.isInteger(ipavgift(2024, 400_123, 50, 1, 1))).toBe(false);
  });
});
