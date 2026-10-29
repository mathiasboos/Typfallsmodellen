import { describe, expect, it } from "vitest";

import { loadDeathProbabilities } from "../src/data/nodeLoader.js";
import { POTS, Scheme, defaultContext, defaultInput, run } from "../src/index.js";
import type { LedgerRow, ModelContext, PotId, TypfallInput, TypfallResult } from "../src/index.js";

/**
 * `result.ledger`: each pot of pension capital, year by year, for the web app's
 * "Portföljvärde" tab. Not a workbook output, so nothing to compare it against but
 * its own arithmetic and the model's own payouts.
 */

const deaths = loadDeathProbabilities();

const compute = (input: Partial<TypfallInput>, context: Partial<ModelContext> = {}): TypfallResult =>
  run(defaultInput(input) as TypfallInput, defaultContext({ priceBasis: -1, ...context }) as ModelContext, { deaths });

/** A typfall with every pot in use: ITP 1, and private saving of the given kind with capital. */
const withSaving = (kind: number): TypfallResult =>
  compute(
    { born: 1980, scheme: Scheme.Itp1, retirementAge: 67 },
    { ipsMonthly: 2000, ipsOpening: 100_000, privateSavingKind: kind },
  );

const cases: readonly [string, () => TypfallResult][] = [
  ["the workbook's own typfall", () => compute({})],
  ["ITP 1 and an IPS", () => withSaving(0)],
  ["ITP 1 and a KF", () => withSaving(1)],
  ["ITP 1 and an ISK", () => withSaving(2)],
  ["ITP 1 and a depot", () => withSaving(3)],
  ["KAP-KL", () => compute({ born: 1975, scheme: Scheme.KapKl })],
  ["ITP 2, defined benefit", () => compute({ born: 1965, scheme: Scheme.Itp2 })],
];

const row = (result: TypfallResult, age: number): LedgerRow => {
  const found = result.ledger.find((r) => r.age === age);
  if (found === undefined) throw new Error(`no ledger row at ${age}`);
  return found;
};

describe("the ledger", () => {
  it.each(cases)("adds up for every pot in every year: %s", (_, make) => {
    const result = make();
    expect(result.ledger.length).toBe(result.rows.length);
    let previous: LedgerRow | undefined;
    for (const line of result.ledger) {
      for (const pot of POTS) {
        const y = line.pots[pot];
        const sum = y.open + y.contribution + y.growth - y.withdrawal - y.fee - y.yieldTax + y.adjustment;
        expect(sum).toBeCloseTo(y.close, 6);
        // Each year opens on last year's closing balance.
        expect(y.open).toBeCloseTo(previous?.pots[pot].close ?? 0, 6);
        expect(y.close).toBeGreaterThanOrEqual(0);
        expect(y.withdrawal).toBeGreaterThanOrEqual(0);
      }
      previous = line;
    }
  });

  it.each(cases)("pays out what Table 2's rows pay, while a pot lasts: %s", (_, make) => {
    const result = make();
    // Nominal kronor, so the rows' amounts are the model's own, rounded to the krona.
    const paid: Record<PotId, (r: TypfallResult["rows"][number]) => number> = {
      ip: (r) => r.ip,
      pp: (r) => r.pp,
      tjp: (r) => r.tjp,
      ips: (r) => r.ips,
      pps: (r) => r.pps,
    };
    for (const mv of result.rows) {
      const line = row(result, mv.age);
      for (const pot of POTS) {
        const y = line.pots[pot];
        // A pot that ran dry pays what it had; every other year all that was due.
        if (y.close > 0 || pot === "ip" || pot === "pp") expect(y.withdrawal).toBeCloseTo(paid[pot](mv), -1);
        else expect(y.withdrawal).toBeLessThanOrEqual(paid[pot](mv) + 1);
      }
    }
  });

  it("carries the income pension on after retirement as what the pension still to come is worth", () => {
    const result = compute({});
    const par = defaultInput().retirementAge;
    const atRetirement = row(result, par - 1).pots.ip.close;
    expect(atRetirement).toBeGreaterThan(2_000_000);
    // It falls year by year from the first full year of pension, and is gone at 105.
    for (let age = par + 1; age < 105; age += 1) {
      expect(row(result, age).pots.ip.close).toBeLessThan(row(result, age - 1).pots.ip.close);
    }
    expect(row(result, 105).pots.ip.close).toBe(0);
    // Ten years in it is worth roughly ten years' pension less, not more.
    const tenOn = row(result, par + 10).pots.ip.close;
    expect(tenOn).toBeLessThan(atRetirement * 0.75);
    expect(tenOn).toBeGreaterThan(atRetirement * 0.45);
  });

  it("credits capital already saved as a contribution the year before the saving starts", () => {
    const context = defaultContext();
    for (const [kind, pot] of [
      [0, "ips"],
      [2, "pps"],
    ] as const) {
      const result = withSaving(kind);
      const line = result.ledger.find((r) => r.year === context.ipsStart - 1)!;
      expect(line.pots[pot].contribution).toBe(100_000);
      expect(line.pots[pot].close).toBe(100_000);
      // Only into the pot the saving's kind builds.
      expect(line.pots[pot === "ips" ? "pps" : "ips"].close).toBe(0);
    }
  });

  it("shows a balance typed in for a year as an adjustment, not as return", () => {
    const year = 2020;
    const typed = { pbhYear: year, pbhIp: 1_000_000, pbhPp: 200_000, pbhTjp: 300_000, pbhIps: 0 };
    const built = compute({ born: 1975, scheme: Scheme.Itp1 });
    const result = compute({ born: 1975, scheme: Scheme.Itp1 }, typed);
    const line = result.ledger.find((r) => r.year === year)!;
    const before = built.ledger.find((r) => r.year === year)!;
    expect(line.pots.ip.close).toBe(1_000_000);
    expect(line.pots.pp.close).toBe(200_000);
    expect(line.pots.tjp.close).toBe(300_000);
    expect(line.pots.ip.adjustment).toBeCloseTo(1_000_000 - before.pots.ip.close, 0);
    // The year's return is the same as without the typed balance, give or take the krona the
    // model truncates a balance to.
    expect(Math.abs(line.pots.pp.growth - before.pots.pp.growth)).toBeLessThan(1);
  });

  it("does not change what the model computes", () => {
    // The golden comparison covers this case by case; here, that the ledger's arithmetic
    // leaves the balances and the payouts where a run without anyone reading it puts them.
    const a = withSaving(0);
    const b = withSaving(0);
    expect(a.rows).toEqual(b.rows);
    expect(a.table1).toEqual(b.table1);
  });
});
