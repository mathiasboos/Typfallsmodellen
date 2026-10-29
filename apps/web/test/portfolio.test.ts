/**
 * "Portföljvärde"'s arithmetic, checked without a browser.
 *
 * The panel and its charts are checked against the built page by
 * `tools/build/verify-offline.mjs`; what is pinned here is what `portfolio.ts`
 * makes of the engine's ledger: which accounts there are, the units, the chaining
 * of one year's closing balance into the next one's opening, the tax-adjusted view,
 * the all-accounts summary and the downloads.
 */
import { describe, expect, it } from "vitest";

import { Scheme, Table1Key, defaultContext, defaultInput, run } from "@typfallsmodellen/engine";
import type { ModelContext, TypfallInput, TypfallResult } from "@typfallsmodellen/engine";

import { loadDeathProbabilities } from "../src/deaths.js";
import { retirementAge } from "../src/kpis.js";
import {
  LINE_KEYS,
  accountName,
  accountRows,
  buildPortfolio,
  pensionTaxRate,
  portfolioToCsv,
  portfolioToXlsxRows,
  summaryRows,
  totalOf,
} from "../src/portfolio.js";
import type { Portfolio, PortfolioMode } from "../src/portfolio.js";

const deaths = loadDeathProbabilities();

interface Case {
  readonly result: TypfallResult;
  readonly context: ModelContext;
  readonly par: number;
}

function compute(input: Partial<TypfallInput> = {}, settings: Partial<ModelContext> = {}): Case {
  const typfall = defaultInput(input) as TypfallInput;
  const context = defaultContext(settings) as ModelContext;
  const result = run(typfall, context, { deaths });
  return { result, context, par: retirementAge(typfall, result) };
}

const portfolio = (c: Case, mode: PortfolioMode = "total"): Portfolio =>
  buildPortfolio(c.result, c.par, mode, c.context.privateSavingKind);

/** The workbook's typfall: no occupational pension, no saving. */
const plain = compute();
/** Every kind of account but the IPS: ITP 1 and an ISK. */
const full = compute({ born: 1980, scheme: Scheme.Itp1 }, { ipsMonthly: 2000, privateSavingKind: 2 });
/** An IPS rather than an ISK. */
const withIps = compute({ born: 1975, scheme: Scheme.Itp1 }, { ipsMonthly: 1000, privateSavingKind: 0 });
/** Prices that move, so a year's units are not the next one's. */
const inflatedInput = { born: 1975, scheme: Scheme.Itp1, yearlyInflation: 0.02 } as const;
const moving = compute(inflatedInput, { ipsMonthly: 1000, privateSavingKind: 0 });

describe("the accounts", () => {
  it("are the pots that hold anything, in stack order", () => {
    expect(portfolio(plain).accounts.map((a) => a.id)).toEqual(["ip", "pp"]);
    expect(portfolio(full).accounts.map((a) => a.id)).toEqual(["ip", "pp", "tjp", "pps"]);
    expect(portfolio(withIps).accounts.map((a) => a.id)).toEqual(["ip", "pp", "tjp", "ips"]);
  });

  it("are named for what they are, private saving by its kind", () => {
    expect(accountName("pps", "sv", 2)).toBe("Privat sparande (ISK)");
    expect(accountName("pps", "en", 3)).toBe("Private saving (securities account)");
    expect(accountName("tjp", "sv")).toBe("Tjänstepension");
  });

  it("run from the first year anything is saved to the model's last age", () => {
    const p = portfolio(plain);
    const first = p.years[0]!;
    expect(first.age).toBeGreaterThan(15);
    expect(p.years[p.years.length - 1]!.age).toBe(105);
    expect(p.years.every((y, i) => i === 0 || y.age === p.years[i - 1]!.age + 1)).toBe(true);
    expect(p.retirementYear).toBe(defaultInput().born + p.retirementAge);
  });
});

describe("each account's year", () => {
  it.each([
    ["no inflation", plain],
    ["ITP 1 and an ISK", full],
    ["2 % inflation, in fixed prices", moving],
  ] as const)("adds up, and opens on last year's closing balance: %s", (_, c) => {
    for (const mode of ["total", "taxadj"] as const) {
      const p = portfolio(c, mode);
      for (const a of p.accounts) {
        a.lines.forEach((l, i) => {
          const sum = l.open + l.contribution + l.growth - l.withdrawal - l.fee - l.yieldTax + l.adjustment;
          expect(sum).toBeCloseTo(l.close, 6);
          if (i > 0) expect(l.open).toBeCloseTo(a.lines[i - 1]!.close, 6);
        });
      }
    }
  });

  it("pays out what Table 2's rows pay, in the same units", () => {
    for (const c of [plain, full, moving]) {
      const p = portfolio(c);
      const byAge = new Map(c.result.rows.map((r) => [r.age, r]));
      const ip = p.accounts.find((a) => a.id === "ip")!;
      const pp = p.accounts.find((a) => a.id === "pp")!;
      p.years.forEach((y, i) => {
        const row = byAge.get(y.age)!;
        expect(Math.abs(ip.lines[i]!.withdrawal - row.ip)).toBeLessThan(0.51);
        expect(Math.abs(pp.lines[i]!.withdrawal - row.pp)).toBeLessThan(0.51);
      });
    }
  });

  it("in fixed prices under inflation, a balance is worth less than its nominal kronor later on", () => {
    const p = portfolio(moving);
    const tjp = p.accounts.find((a) => a.id === "tjp")!;
    const at = p.years.findIndex((y) => y.age === p.retirementAge - 1);
    const nominal = moving.result.ledger.find((r) => r.age === p.retirementAge - 1)!.pots.tjp.close;
    expect(tjp.lines[at]!.close).toBeLessThan(nominal);
    expect(tjp.lines[at]!.close).toBeGreaterThan(nominal * 0.3);
  });
});

describe("Skattekorrigerat", () => {
  it("takes the average tax on the pension at retirement from Table 1", () => {
    const byKey = new Map(full.result.table1.map((r) => [r.key, r]));
    const rate = 1 - byKey.get(Table1Key.PensionAfterTax)!.nominal / byKey.get(Table1Key.TotalGross)!.nominal;
    expect(pensionTaxRate(full.result)).toBeCloseTo(rate, 12);
    expect(rate).toBeGreaterThan(0.15);
    expect(rate).toBeLessThan(0.45);
  });

  it("scales every amount of a taxed account by it, and leaves the private saving alone", () => {
    const total = portfolio(full, "total");
    const adjusted = portfolio(full, "taxadj");
    const keep = 1 - adjusted.taxRate;
    for (const a of adjusted.accounts) {
      const before = total.accounts.find((b) => b.id === a.id)!;
      const factor = a.meta.kind === "taxed" ? 1 : keep;
      a.lines.forEach((l, i) => {
        for (const key of LINE_KEYS) expect(l[key]).toBeCloseTo(before.lines[i]![key] * factor, 6);
      });
    }
  });
});

describe("the all-accounts summary", () => {
  it("adds the accounts up, contributions and withdrawals split by how they are taxed", () => {
    const p = portfolio(full);
    const rows = new Map(summaryRows(p, "sv").map((r) => [r.key, r.values]));
    const close = totalOf(p, "close");
    expect(rows.get("close")).toEqual(close);
    const pps = p.accounts.find((a) => a.id === "pps")!;
    p.years.forEach((_, i) => {
      expect(rows.get("contribution-after")![i]).toBeCloseTo(pps.lines[i]!.contribution, 6);
      expect(rows.get("withdrawal-after")![i]).toBeCloseTo(pps.lines[i]!.withdrawal, 6);
      const contributions = rows.get("contribution-after")![i]! + rows.get("contribution-before")![i]!;
      expect(contributions).toBeCloseTo(totalOf(p, "contribution")[i]!, 6);
    });
    // No balance was typed in, so there is no adjustment row.
    expect(rows.has("adjustment")).toBe(false);
  });

  it("has the planner's rows, and one account's table its own", () => {
    const p = portfolio(full);
    expect(summaryRows(p, "sv").map((r) => r.label)).toEqual([
      "Ingående saldo",
      "Insättning, konton efter skatt",
      "Insättning, konton före skatt",
      "Uttag, konton efter skatt",
      "Uttag, konton före skatt",
      "Avkastning",
      "Avkastningsskatt",
      "Förvaltningsavgift",
      "Utgående saldo",
    ]);
    const tjp = p.accounts.find((a) => a.id === "tjp")!;
    expect(accountRows(p, tjp, "en").map((r) => r.label)).toEqual([
      "Opening balance",
      "Contribution, pre-tax account",
      "Withdrawal, pre-tax account",
      "Portfolio return",
      "Yield tax",
      "Management fee",
      "Closing balance",
    ]);
  });

  it("shows a balance typed in under Avancerat on a row of its own", () => {
    const typed = compute({ born: 1975, scheme: Scheme.Itp1 }, { pbhYear: 2020, pbhIp: 1_000_000, pbhPp: 200_000, pbhTjp: 300_000 });
    const p = portfolio(typed);
    expect(p.hasAdjustment).toBe(true);
    expect(summaryRows(p, "sv").some((r) => r.key === "adjustment")).toBe(true);
  });
});

describe("the downloads", () => {
  it("write a block per table, with ages, years and outflows as negative numbers", () => {
    const p = portfolio(full);
    const csv = portfolioToCsv(p, "all", "sv").split("\r\n");
    expect(csv[0]).toBe("Alla konton (summa)");
    expect(csv[1]!.startsWith("Ålder;")).toBe(true);
    expect(csv[2]!.startsWith("År;")).toBe(true);
    const withdrawals = csv.find((line) => line.startsWith("Uttag, konton före skatt;"))!;
    expect(withdrawals).toMatch(/;−\d/);
    // One block per account when every account is shown.
    const each = portfolioToCsv(p, "each", "en");
    expect(each.split("\r\n").filter((line) => line.startsWith("Age,")).length).toBe(p.accounts.length);
    const one = portfolioToCsv(p, "tjp", "en").split("\r\n");
    expect(one[0]).toBe("Occupational pension (pension insurance · taxed when paid)");
  });

  it("give Excel numbers, not text", () => {
    const p = portfolio(full);
    const sheet = portfolioToXlsxRows(p, "all", "sv");
    const close = sheet.find((row) => (row[0] as { value?: unknown } | null)?.value === "Utgående saldo")!;
    const cell = close[1] as { value: number; type: unknown };
    expect(typeof cell.value).toBe("number");
    expect(cell.value).toBe(Math.round(totalOf(p, "close")[0]!));
  });
});
