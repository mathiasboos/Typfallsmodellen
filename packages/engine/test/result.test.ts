import { describe, expect, it } from "vitest";

import snapshotJson from "../../../reference/fixtures/default-run.json" with { type: "json" };
import { loadDeathProbabilities } from "../src/data/nodeLoader.js";
import {
  Scheme,
  Table1Key,
  defaultContext,
  defaultInput,
  run,
} from "../src/index.js";
import type { ModelContext, TypfallInput, TypfallResult } from "../src/index.js";

const deaths = loadDeathProbabilities();
const snapshot = snapshotJson as unknown as {
  input: TypfallInput;
  qualifyingYears: number;
  table1: { key: string; nominal: number; adjusted: number; monthly: number }[];
  lifeIncome: { gross: number; net: number; disposable: number };
  rows: Record<string, number>[];
};

/**
 * The result model: Table 1, Table 2, the life-income sums.
 *
 * None of these amounts is verified against the workbook -- that is Phase 2 --
 * so these tests assert the relationships the tables have to satisfy whatever
 * the amounts are, and the quirks worth pinning down.
 */
const compute = (
  input: Parameters<typeof defaultInput>[0] = {},
  context: Parameters<typeof defaultContext>[0] = {},
): TypfallResult =>
  run(defaultInput(input) as TypfallInput, defaultContext(context) as ModelContext, { deaths });

const row = (result: TypfallResult, key: string) => result.table1.find((r) => r.key === key)!;

describe("run", () => {
  it("computes the workbook's own typfall without complaint", () => {
    const result = compute();
    expect(result.warnings).toEqual([]);
    expect(result.rows).toHaveLength(105);
    expect(result.table1.length).toBeGreaterThan(10);
    expect(result.table2.length).toBeGreaterThan(40);
  });

  it("reports the corrections it made", () => {
    const result = compute({ retirementAge: 61 });
    expect(result.warnings).toHaveLength(1);
    expect(result.warnings[0]!.field).toBe("ParYear");
  });
});

describe("Table 1", () => {
  it("adds its components up to the totals", () => {
    const r = compute();
    const parts = [
      Table1Key.IncomePension,
      Table1Key.SupplementaryPension,
      Table1Key.PremiumPension,
      Table1Key.GuaranteePension,
      Table1Key.IncomePensionSupplement,
    ].reduce((sum, key) => sum + row(r, key).nominal, 0);
    expect(row(r, Table1Key.TotalPublicPension).nominal).toBeCloseTo(parts, 6);

    const gross =
      row(r, Table1Key.TotalPublicPension).nominal +
      row(r, Table1Key.OccupationalPension).nominal +
      row(r, Table1Key.PrivateSaving).nominal;
    expect(row(r, Table1Key.TotalGross).nominal).toBeCloseTo(gross, 6);
  });

  /**
   * The last three rows measure themselves against something other than the
   * final salary, so they are checked separately below.
   */
  const OWN_DENOMINATOR: readonly string[] = [
    Table1Key.PensionAfterTax,
    Table1Key.BenefitsAtRetirement,
    Table1Key.DisposableAtRetirement,
  ];

  it("keeps its four columns consistent", () => {
    const r = compute();
    const salary = row(r, Table1Key.FinalSalary);
    for (const line of r.table1) {
      expect(line.monthly).toBeCloseTo(line.adjusted / 12, 6);
      if (salary.adjusted > 0 && !OWN_DENOMINATOR.includes(line.key)) {
        expect(line.shareOfFinalSalary).toBeCloseTo(line.adjusted / salary.adjusted, 9);
      }
    }
    expect(salary.shareOfFinalSalary).toBe(1);
  });

  it("measures the last three rows against net and disposable income", () => {
    // VBA_go.bas 2679, 2776 and 2790: the pension after tax is compared with
    // the salary after tax, and both the benefits and the disposable income
    // with last year's disposable income -- not with the final salary.
    const r = compute();
    const net = row(r, Table1Key.SalaryAfterTax);
    const disp = row(r, Table1Key.DisposableBeforeRetirement);

    expect(row(r, Table1Key.PensionAfterTax).shareOfFinalSalary).toBeCloseTo(
      row(r, Table1Key.PensionAfterTax).adjusted / net.adjusted,
      9,
    );
    expect(row(r, Table1Key.DisposableAtRetirement).shareOfFinalSalary).toBeCloseTo(
      row(r, Table1Key.DisposableAtRetirement).adjusted / disp.adjusted,
      9,
    );
  });

  it("gives a plausible compensation level for the shipped typfall", () => {
    // Not a check against the workbook -- just that the model is in the right
    // world: a full career on an average salary, with no occupational pension.
    const r = compute();
    const share = row(r, Table1Key.TotalPublicPension).shareOfFinalSalary;
    expect(share).toBeGreaterThan(0.35);
    expect(share).toBeLessThan(0.75);
  });

  it("adds an occupational pension on top", () => {
    const without = compute({ born: 1985 });
    const withItp = compute({ born: 1985, scheme: Scheme.Itp1 });
    expect(row(without, Table1Key.OccupationalPension).nominal).toBe(0);
    expect(row(withItp, Table1Key.OccupationalPension).nominal).toBeGreaterThan(0);
    expect(row(withItp, Table1Key.TotalGross).nominal).toBeGreaterThan(
      row(without, Table1Key.TotalGross).nominal,
    );
  });

  it("credits a final year of pension rights by default", () => {
    const withRight = compute();
    const without = compute({}, { lastPensionRight: 0 });
    expect(row(withRight, Table1Key.IncomePension).nominal).toBeGreaterThan(
      row(without, Table1Key.IncomePension).nominal,
    );
  });

  it("averages the final salary over the years the setting asks for", () => {
    const oneYear = compute({}, { finalSalaryYears: 1 });
    const fiveYears = compute();
    expect(row(fiveYears, Table1Key.FinalSalary).nominal).not.toBeCloseTo(
      row(oneYear, Table1Key.FinalSalary).nominal,
      0,
    );
    // Five years reaching further back is a lower average in a rising series.
    expect(row(fiveYears, Table1Key.FinalSalary).nominal).toBeLessThan(
      row(oneYear, Table1Key.FinalSalary).nominal,
    );
  });

  it("leaves the nominal and adjusted columns equal when the basis is nominal", () => {
    const r = compute({}, { priceBasis: -1 });
    for (const line of r.table1) {
      expect(line.adjusted).toBeCloseTo(line.nominal, 6);
    }
  });

  it("QUIRK: the retirement year's salary is zeroed before the summary", () => {
    // `Income_(PAR)` and `Wage_(PAR)` are set to 0 before Table 1 is drawn, so
    // the table reports a full year of pension against no salary at all. Every
    // cohort the workbook offers is born on 1 January, which already leaves no
    // salary in the retirement year, so the zeroing changes nothing -- but it
    // would for a mid-year birthday, and the per-age matrix keeps the income
    // either way.
    const r = compute();
    const atPar = r.rows.find((x) => x.age === 66)!;
    expect(atPar.income).toBe(0);
    expect(row(r, Table1Key.TotalGross).nominal).toBeGreaterThan(0);
    // The matrix is untouched by the summary: a working age still has income.
    expect(r.rows.find((x) => x.age === 60)!.income).toBeGreaterThan(0);
  });
});

describe("Table 2", () => {
  it("starts ten years before retirement by default", () => {
    const r = compute();
    expect(r.table2[0]!.age).toBe(56);
    expect(r.table2.at(-1)!.age).toBe(105);
  });

  it("takes an explicit start age when one is set", () => {
    expect(compute({}, { table2StartAge: 40 }).table2[0]!.age).toBe(40);
  });

  it("pairs the columns the way the figures do", () => {
    const r = compute({ born: 1945 });
    for (const t2 of r.table2) {
      const m = r.rows.find((x) => x.age === t2.age)!;
      expect(t2.incomeAndSupplementary).toBeCloseTo(m.ip + m.tp, 6);
      expect(t2.occupationalAndPrivate).toBeCloseTo(m.tjp + m.ips, 6);
      expect(t2.guaranteeAndSupplement).toBeCloseTo(m.garp + m.ptillagg, 6);
      expect(t2.gross).toBeCloseTo(m.brutto, 6);
      expect(t2.municipalTax).toBeCloseTo(m.municipalTax, 6);
      expect(t2.stateTax).toBeCloseTo(m.stateTax, 6);
      expect(t2.disposable).toBeCloseTo(m.indDisp, 6);
    }
  });

  it("divides by twelve when the figures are shown per month", () => {
    const yearly = compute();
    const monthly = compute({}, { chartEarningFactor: 12 });
    expect(monthly.table2[0]!.salary).toBeCloseTo(yearly.table2[0]!.salary / 12, 6);
  });

  it("shows a salary before retirement and a pension after", () => {
    const r = compute();
    const before = r.table2.find((x) => x.age === 60)!;
    const after = r.table2.find((x) => x.age === 70)!;
    expect(before.salary).toBeGreaterThan(0);
    expect(before.incomeAndSupplementary).toBe(0);
    expect(after.salary).toBe(0);
    expect(after.incomeAndSupplementary).toBeGreaterThan(0);
  });
});

describe("life income", () => {
  it("runs from the pension age to the expected age at death", () => {
    const r = compute();
    expect(r.lifeIncome.throughAge).toBeGreaterThan(80);
    expect(r.lifeIncome.throughAge).toBeLessThan(95);
  });

  it("orders the three sums the way the amounts are ordered", () => {
    const r = compute();
    expect(r.lifeIncome.gross).toBeGreaterThan(r.lifeIncome.net);
    expect(r.lifeIncome.disposable).toBeGreaterThan(r.lifeIncome.net);
    expect(r.lifeIncome.disposable).toBeLessThan(r.lifeIncome.gross);
  });

  it("falls as the discount rate rises", () => {
    const low = compute({}, { discountRate: 0.01 }).lifeIncome.gross;
    const high = compute({}, { discountRate: 0.1 }).lifeIncome.gross;
    expect(high).toBeLessThan(low);
  });

  it("is larger for a later retirement, which pays more for fewer years", () => {
    const early = compute({ retirementAge: 63 });
    const late = compute({ retirementAge: 70 });
    expect(row(late, Table1Key.IncomePension).monthly).toBeGreaterThan(
      row(early, Table1Key.IncomePension).monthly,
    );
  });
});

describe("the municipal/state tax split", () => {
  it("sums to the same total tax netto already implies, at every age", () => {
    const r = compute();
    const par = (defaultInput() as TypfallInput).retirementAge;
    // `r()` rounds municipalTax, stateTax and netto each to a whole kronor
    // independently, so their sum can be a kronor off the rounded `tax` a
    // second, separate rounding of the same figures would give -- allow that
    // much slack, not the tighter one `toBeCloseTo`'s digit count would ask for.
    const reconciles = (m: (typeof r.rows)[number]) => {
      if (m.brutto === 0 && m.netto === 0) return;
      const tax = m.brutto - m.netto;
      expect(Math.abs(m.municipalTax + m.stateTax - tax), `age ${m.age}`).toBeLessThanOrEqual(1);
    };
    for (const m of r.rows) reconciles(m);
    // The retirement-transition age is the one row `atRetirement.ts`'s
    // second tax pass rewrites -- check it explicitly, not just in passing.
    reconciles(r.rows.find((row) => row.age === par)!);
  });
});

describe("across the model's whole range", () => {
  it("produces a finite result for every cohort, retirement age and agreement", () => {
    for (const born of [1930, 1945, 1960, 1980, 2000, 2024]) {
      for (const retirementAge of [63, 66, 70]) {
        for (const scheme of [Scheme.None, Scheme.Itp2, Scheme.KapKl, Scheme.Pa16Avd1]) {
          const r = compute({ born, retirementAge, scheme });
          const label = `born ${born} par ${retirementAge} scheme ${scheme}`;
          for (const line of r.table1) {
            expect(Number.isFinite(line.nominal), `${label} ${line.key}`).toBe(true);
            expect(Number.isFinite(line.adjusted), `${label} ${line.key}`).toBe(true);
          }
          expect(Number.isFinite(r.lifeIncome.gross), label).toBe(true);
          expect(r.table2.length).toBeGreaterThan(0);
        }
      }
    }
  });

  it("never reports a negative pension", () => {
    for (const salary of [8_000, 20_000, 38_500, 80_000]) {
      const r = compute({ monthlySalary: salary });
      for (const line of r.table1) {
        expect(line.nominal, `salary ${salary} ${line.key}`).toBeGreaterThanOrEqual(0);
      }
    }
  });
});

describe("the default run, against its committed snapshot", () => {
  /**
   * Not a check against the workbook -- nothing here is, until Phase 2. It is a
   * regression snapshot: an unintended change to any rule in the engine shows
   * up as a diff in reference/fixtures/default-run.json rather than silently.
   */
  it("still produces the figures that were committed", () => {
    const r = compute();
    const round = (x: number) => Math.round(x * 100) / 100;

    expect(snapshot.input).toEqual(defaultInput());
    expect(r.qualifyingYears).toBe(snapshot.qualifyingYears);

    for (const want of snapshot.table1) {
      const got = row(r, want.key);
      expect(round(got.nominal), `${want.key} nominal`).toBe(want.nominal);
      expect(round(got.adjusted), `${want.key} adjusted`).toBe(want.adjusted);
      expect(round(got.monthly), `${want.key} monthly`).toBe(want.monthly);
    }

    expect(round(r.lifeIncome.gross)).toBe(snapshot.lifeIncome.gross);
    expect(round(r.lifeIncome.net)).toBe(snapshot.lifeIncome.net);
    expect(round(r.lifeIncome.disposable)).toBe(snapshot.lifeIncome.disposable);

    for (const want of snapshot.rows) {
      const got = r.rows.find((x) => x.age === want.age)!;
      expect(got, `age ${want.age}`).toEqual(want);
    }
  });
});

/**
 * The wage path the advanced form's salary grid starts from.
 *
 * `Egen löneutveckling` (manual section 3.1) is filled in from the computed
 * path and then edited, so the fill has to be an exact description of the run
 * it came from -- otherwise ticking the box would move the answer before
 * anyone typed anything.
 */
describe("the wage path", () => {
  it("covers every age from the first, contiguously, and stops at the last worked one", () => {
    const r = compute();
    const path = r.wagePath;

    expect(path.length).toBeGreaterThan(0);
    for (let i = 1; i < path.length; i += 1) {
      expect(path[i]!.age, `row ${i}`).toBe(path[i - 1]!.age + 1);
    }

    const last = path[path.length - 1]!;
    expect(last.income === 0 && last.wage === 0).toBe(false);
    // Retirement is at 66 in the shipped typfall, so the last worked age is 65.
    expect(last.age).toBeLessThan((defaultInput() as TypfallInput).retirementAge);
  });

  it("is nominal, where the per-age matrix's income is in reference-year prices", () => {
    // `recordRow` multiplies every amount by the price factor before storing
    // it; `Income_(age)` on the Indata_lista sheet is the unadjusted figure,
    // and has to stay that way or feeding it back would deflate the run twice.
    // `kpiFactor` is that same factor, so the two reconcile through it.
    const r = compute();
    let differed = 0;
    for (const m of r.rows) {
      const entry = r.wagePath.find((x) => x.age === m.age);
      if (entry === undefined) continue;
      // `kpiFactor` is itself stored to four decimals, so the slack has to
      // scale with the amount it multiplies: a 5e-5 error on a 400 000 kr
      // income is 20 kr, and nothing is wrong with it.
      const slack = 1 + entry.income * 5e-5;
      expect(Math.abs(entry.income * m.kpiFactor - m.income), `age ${m.age}`).toBeLessThanOrEqual(slack);
      if (Math.round(entry.income) !== m.income) differed += 1;
    }
    // If they never differed the factor would be 1 throughout and this would
    // prove nothing.
    expect(differed).toBeGreaterThan(0);
  });

  it("reproduces the run when handed straight back as ownIncome", () => {
    const r = compute();
    const again = run(
      { ...(defaultInput() as TypfallInput), ownIncome: r.wagePath },
      defaultContext(),
      { deaths },
    );
    for (const want of r.table1) {
      const got = again.table1.find((x) => x.key === want.key)!;
      expect(got.nominal, `${want.key} nominal`).toBeCloseTo(want.nominal, 6);
      expect(got.monthly, `${want.key} monthly`).toBeCloseTo(want.monthly, 6);
    }
  });

  it("reproduces it from age 15 up too, which is all an own vector reads", () => {
    // `validate` pins `startage` to 15 whenever `ownIncome` is present, so the
    // grid shows 15 and up and sends only that slice back.
    const r = compute();
    const again = run(
      { ...(defaultInput() as TypfallInput), ownIncome: r.wagePath.filter((x) => x.age >= 15) },
      defaultContext(),
      { deaths },
    );
    for (const want of r.table1) {
      const got = again.table1.find((x) => x.key === want.key)!;
      expect(got.monthly, `${want.key} monthly`).toBeCloseTo(want.monthly, 6);
    }
  });
});
