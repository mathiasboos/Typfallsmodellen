import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import annuityTables from "../../data/annuity-tables.json" with { type: "json" };
import { loadDeathProbabilities } from "../src/data/nodeLoader.js";
import { defaultContext } from "../src/model/context.js";
import { DeltalTables } from "../src/pension/deltal.js";
import type { AnnuityTablesData } from "../src/pension/deltal.js";
import {
  deltal,
  incomePensionYear,
  pgbBarn,
  ppkassa,
  pUttag,
} from "../src/pension/incomePension.js";
import type { WithdrawalShares } from "../src/pension/incomePension.js";

interface BruttoRow {
  age: number;
  year: number;
  ipavgift: number | null;
  arvsvinstIpUnder65: number | null;
  arvsvinstIpOver65: number | null;
  forvaltningsfaktorIp: number | null;
  gallandeIndexNiva: number | null;
  ipArvsvinst: number | null;
  ipIndexering: number | null;
  ipForvaltningskostnad: number | null;
  ipBegransadUppskrivning: number | null;
  pbhIp: number | null;
}

const fixture = JSON.parse(
  readFileSync(
    fileURLToPath(new URL("../../../reference/fixtures/brutto-cached.json", import.meta.url)),
    "utf8",
  ),
) as { rows: BruttoRow[] };

const byYear = new Map(fixture.rows.map((r) => [r.year, r]));
/**
 * The Brutto typfall: born 2000, and the sheet never draws a pension -- it keeps
 * accumulating right through to age 98. So the retirement age passed here is set
 * beyond the fixture's range, which keeps every row in the earning-phase branch,
 * the one Brutto's inline formulas reproduce.
 *
 * This fixture therefore checks the accumulation thoroughly and the drawdown not
 * at all. The retirement branches wait on the golden files.
 */
const BORN = 2000;
const NEVER_RETIRES = 120;

const tables = new DeltalTables(
  annuityTables as unknown as AnnuityTablesData,
  loadDeathProbabilities(),
  defaultContext(),
);
const shares: WithdrawalShares = { incomePension: 1, premiumPension: 1 };

/**
 * Years where the fixture has everything needed to roll the balance forward:
 * an opening balance, a pension right, the factors, and next year's index.
 */
const accumulationYears = fixture.rows.filter((r) => {
  const previous = byYear.get(r.year - 1);
  const next = byYear.get(r.year + 1);
  return (
    previous?.pbhIp != null &&
    r.ipavgift != null &&
    r.arvsvinstIpUnder65 != null &&
    r.forvaltningsfaktorIp != null &&
    r.gallandeIndexNiva != null &&
    next?.gallandeIndexNiva != null &&
    r.ipArvsvinst != null &&
    previous.pbhIp > 0
  );
});

function rollForward(row: BruttoRow) {
  const previous = byYear.get(row.year - 1)!;
  const next = byYear.get(row.year + 1)!;
  // A year's pension right is credited the year after, so the indexation that
  // applies is next year's level over this one's -- which is what Brutto uses.
  const index = next.gallandeIndexNiva! / row.gallandeIndexNiva!;
  return incomePensionYear(
    row.year,
    NEVER_RETIRES,
    BORN,
    row.ipavgift!,
    row.arvsvinstIpUnder65!,
    row.arvsvinstIpOver65!,
    row.forvaltningsfaktorIp!,
    previous.pbhIp!,
    0,
    0,
    0,
    index,
    NEVER_RETIRES,
    // marginal 1 so nothing is truncated here; Brutto truncates the total
    // instead, which the balance check below reproduces.
    1,
    0,
  );
}

describe("income pension accumulation, against the Brutto sheet", () => {
  it("has years to compare", () => {
    expect(accumulationYears.length).toBeGreaterThan(50);
  });

  it("covers accumulation only, since the sheet never draws a pension", () => {
    // Worth stating: nothing here exercises IP_'s drawdown branches.
    expect(fixture.rows.every((r) => r.year - BORN < NEVER_RETIRES)).toBe(true);
  });

  it("reproduces the inheritance gains credited each year", () => {
    const bad = accumulationYears
      .map((r) => {
        const got = rollForward(r).inheritanceGains;
        return Math.abs(got - r.ipArvsvinst!) < 1e-6
          ? null
          : `${r.year}: got ${got}, workbook has ${r.ipArvsvinst}`;
      })
      .filter((m): m is string => m !== null);
    expect(bad.slice(0, 5).join("\n")).toBe("");
  });

  it("reproduces the indexation credited each year", () => {
    const bad = accumulationYears
      .map((r) => {
        const got = rollForward(r).indexation;
        return Math.abs(got - r.ipIndexering!) < 1e-6
          ? null
          : `${r.year}: got ${got}, workbook has ${r.ipIndexering}`;
      })
      .filter((m): m is string => m !== null);
    expect(bad.slice(0, 5).join("\n")).toBe("");
  });

  it("reproduces the management cost deducted each year", () => {
    const bad = accumulationYears
      .map((r) => {
        const got = rollForward(r).managementCost;
        return Math.abs(got - r.ipForvaltningskostnad!) < 1e-6
          ? null
          : `${r.year}: got ${got}, workbook has ${r.ipForvaltningskostnad}`;
      })
      .filter((m): m is string => m !== null);
    expect(bad.slice(0, 5).join("\n")).toBe("");
  });

  it("reproduces the closing balance", () => {
    const bad = accumulationYears
      .map((r) => {
        // Brutto carries the limited uprating of pension rights as its own
        // column; IP_ folds it into pratt through bindex, so it is added back
        // here before truncating, exactly as the sheet does.
        const got = Math.floor(rollForward(r).balance + (r.ipBegransadUppskrivning ?? 0));
        return got === r.pbhIp ? null : `${r.year}: got ${got}, workbook has ${r.pbhIp}`;
      })
      .filter((m): m is string => m !== null);
    expect(bad.slice(0, 5).join("\n")).toBe("");
  });
});

describe("income pension, structural", () => {
  it("gives cohorts before 1938 nothing at all", () => {
    const result = incomePensionYear(2000, 65, 1937, 50_000, 1, 1, 1, 100_000, 1, 0, 15, 1.02);
    expect(result.balance).toBe(0);
    expect(result.pension).toBe(0);
  });

  it("trims the pension right by the balance index ratio from 2015", () => {
    const withTrim = incomePensionYear(
      2016, 66, 1960, 100_000, 1, 1, 1, 0, 0, 0, 0, 1, 66, 1, 0.9,
    );
    const without = incomePensionYear(
      2016, 66, 1960, 100_000, 1, 1, 1, 0, 0, 0, 0, 1, 66, 1, 0,
    );
    expect(withTrim.balance).toBeCloseTo(90_000, 6);
    expect(without.balance).toBeCloseTo(100_000, 6);
  });

  it("does not trim before 2015", () => {
    const result = incomePensionYear(2014, 66, 1960, 100_000, 1, 1, 1, 0, 0, 0, 0, 1, 66, 1, 0.9);
    expect(result.balance).toBeCloseTo(100_000, 6);
  });

  it("rounds the annual pension to a whole krona per month", () => {
    const result = incomePensionYear(
      2030, 66, 1960, 0, 1, 1, 1, 3_000_000, 1, 0, 17.5, 1.02, 66, 0, 0,
    );
    expect(result.pension % 12).toBe(0);
  });
});

describe("pUttag", () => {
  const partial: WithdrawalShares = { incomePension: 0.5, premiumPension: 0.25 };

  it("is nothing before pension is drawn", () => {
    expect(pUttag(66, 60, partial, 70)).toBe(0);
  });

  it("is the partial share between first and final withdrawal", () => {
    expect(pUttag(66, 67, partial, 70)).toBe(0.5);
    expect(pUttag(66, 67, partial, 70, false)).toBe(0.25);
  });

  it("is everything once withdrawal is final", () => {
    expect(pUttag(66, 70, partial, 70)).toBe(1);
  });

  it("collapses a final age below the first back onto the first", () => {
    expect(pUttag(66, 66, partial, 60)).toBe(1);
  });
});

describe("deltal from the published tables", () => {
  it("returns the single divisor for cohorts born 1937 or earlier", () => {
    expect(deltal(65, 1935, 65, tables, "income")).toBeGreaterThan(0);
  });

  it("is zero before pension is drawn", () => {
    expect(deltal(66, 1960, 60, tables, "income", 66)).toBe(0);
  });

  it("falls as retirement is deferred", () => {
    const at66 = deltal(66, 1960, 66, tables, "income", 66);
    const at70 = deltal(70, 1960, 70, tables, "income", 70);
    expect(at70).toBeLessThan(at66);
  });

  it("extrapolates below 61 from the premium pension table, whichever pension is asked for", () => {
    // A quirk of the original: the sub-61 branch reads fixed columns belonging
    // to the premium pension table regardless of `kind`.
    expect(deltal(58, 1960, 58, tables, "income", 58)).toBe(
      deltal(58, 1960, 58, tables, "premium", 58),
    );
  });
});

describe("ppkassa", () => {
  it("pays nothing for cohorts before 1938", () => {
    expect(ppkassa(66, 1937, 66, 500_000, tables, shares)).toBe(0);
  });

  it("pays nothing before pension is drawn", () => {
    expect(ppkassa(66, 1960, 60, 500_000, tables, shares, 66)).toBe(0);
  });

  it("pays a full year once retired, near balance over divisor", () => {
    const balance = 600_000;
    const paid = ppkassa(66, 1960, 70, balance, tables, shares, 66);
    const divisor = deltal(66, 1960, 70, tables, "premium", 66);
    expect(paid).toBeGreaterThan(0);
    expect(paid).toBeCloseTo(balance / divisor, -1);
  });

  it("treats a negative balance as zero", () => {
    expect(ppkassa(66, 1960, 70, -1000, tables, shares, 66)).toBe(0);
  });

  it("ignores a withdrawal share below a tenth", () => {
    expect(ppkassa(66, 1960, 70, 500_000, tables, shares, 66, 0.05)).toBe(0);
  });
});

describe("pgbBarn", () => {
  it("credits nothing outside the first four years of a child's life", () => {
    expect(pgbBarn(2010, 400_000, 200_000, 2000, 35, 350_000, 74_300)).toBe(0);
  });

  it("credits at least one income base amount during those years", () => {
    expect(pgbBarn(2002, 400_000, 380_000, 2000, 35, 350_000, 74_300)).toBeGreaterThanOrEqual(
      74_300,
    );
  });

  it("credits the income drop when that is larger", () => {
    const amount = pgbBarn(2002, 500_000, 100_000, 2000, 35, 350_000, 74_300);
    expect(amount).toBeGreaterThan(74_300);
  });

  it("credits nothing to a parent at or past riktålder", () => {
    expect(pgbBarn(2002, 400_000, 200_000, 2000, 66, 350_000, 74_300, 0, 65)).toBe(0);
  });
});
