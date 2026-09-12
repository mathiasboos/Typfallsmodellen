import { describe, expect, it } from "vitest";

import { loadDeathProbabilities } from "../src/data/nodeLoader.js";
import { Table1Key, defaultContext, defaultInput, prepare, run, runLoop } from "../src/index.js";
import type { ModelContext, TypfallInput, TypfallResult } from "../src/index.js";

/**
 * The second tax and benefits pass at the retirement age.
 *
 * None of these amounts is verified against the workbook -- that is what the
 * golden file is for, and these are three of the twelve columns it carries.
 * What is asserted here is the shape: that the pass runs when the settings say
 * so, that it reads the loop's leftovers where the original does, and the
 * quirks that follow from it.
 */

const deaths = loadDeathProbabilities();

const compute = (
  input: Parameters<typeof defaultInput>[0] = {},
  context: Parameters<typeof defaultContext>[0] = {},
): TypfallResult =>
  run(defaultInput(input) as TypfallInput, defaultContext(context) as ModelContext, { deaths });

const value = (r: TypfallResult, key: string): number =>
  r.table1.find((line) => line.key === key)!.nominal;

const atPar = (r: TypfallResult, age = 66) => r.rows.find((row) => row.age === age)!;

describe("the retirement-year recomputation", () => {
  it("does not run when the final pension right is switched off", () => {
    // VBA_go.bas 2607 and 2684 are both inside `If Last_pratt > 0`, while the
    // rows they feed are written unconditionally -- so with it off Table 1
    // reports exactly what the loop computed at that age.
    const r = compute({}, { lastPensionRight: 0 });
    const row = atPar(r);
    expect(value(r, Table1Key.PensionAfterTax)).toBe(row.netto);
    expect(value(r, Table1Key.BenefitsAtRetirement)).toBe(row.bidrag);
    expect(value(r, Table1Key.DisposableAtRetirement)).toBe(row.indDisp);
  });

  it("changes the answer when it does run", () => {
    // The final pension right raises the gross, so the tax on it has to be
    // worked out again; if these agreed, the pass would be pointless.
    const on = compute();
    const off = compute({}, { lastPensionRight: 0 });
    expect(value(on, Table1Key.PensionAfterTax)).not.toBe(value(off, Table1Key.PensionAfterTax));
    expect(value(on, Table1Key.TotalGross)).toBeGreaterThan(value(off, Table1Key.TotalGross));
  });

  it("adds disposable income up from its parts", () => {
    const r = compute();
    const expected =
      value(r, Table1Key.PensionAfterTax) +
      value(r, Table1Key.BenefitsAtRetirement) +
      value(r, Table1Key.PrivateSavingAfterTax);
    expect(value(r, Table1Key.DisposableAtRetirement)).toBeCloseTo(expected, 6);
  });

  it("counts capital income in the pension after tax, taxed at 30 per cent", () => {
    // The loop only counts capital income from the retirement age; this pass
    // adds `kapital` to the gross unconditionally (VBA_go.bas:2663).
    const base = value(compute(), Table1Key.PensionAfterTax);
    const withCapital = value(compute({}, { kapital: 50_000 }), Table1Key.PensionAfterTax);
    expect(withCapital - base).toBeCloseTo(50_000 * 0.7, 6);
  });

  it("leaves the loop's counter one past the last age, as VBA does", () => {
    // `year_(mini(age, 100))` at :2614 and `year_(mini(age, slutage))` at :2749
    // index with the loop counter, which a completed `For` leaves one past its
    // limit. See `RunState.leftovers`.
    const prepared = prepare(defaultInput(), defaultContext(), { deaths });
    runLoop(prepared.run);
    expect(prepared.run.s.leftovers.age).toBe(prepared.run.v.slutage + 1);
  });

  it("QUIRK: the benefits row loses the child and housing allowances", () => {
    // `barnbidrag` and `bostadsbidrag` are not recomputed here -- the blocks
    // that would are commented out at :2687 -- so what `Bidrag(Int(PAR))` adds
    // is whatever the loop last left in them, which is the value at `slutage`,
    // decades after the children have gone. A household with children at
    // retirement therefore sees them in the cash-flow table but not in Table 1.
    const withChildren = compute(
      { monthlySalary: 8000 },
      { childBirthYears: [2012, 2014, 0, 0] },
    );
    const row = atPar(withChildren);
    const table = value(withChildren, Table1Key.BenefitsAtRetirement);

    expect(row.bidrag).toBeGreaterThan(table + 100_000);
    // And the childless run, whose loop leftovers are the same, lands within a
    // housing supplement of it.
    const childless = value(compute({ monthlySalary: 8000 }), Table1Key.BenefitsAtRetirement);
    expect(Math.abs(table - childless)).toBeLessThan(5000);
  });

  it("QUIRK: a household below the riktalder gets no housing supplement here", () => {
    // The pass gates on `PAR >= riktalder And uttagIP = 1 And uttagPP = 1`,
    // where the loop asks only that the income pension is being drawn at all.
    const early = compute({ retirementAge: 63, monthlySalary: 8000 });
    expect(early.warnings).toEqual([]);
    const row = early.rows.find((r) => r.age === 63)!;
    expect(row.bidrag).toBe(0);
    expect(value(early, Table1Key.BenefitsAtRetirement)).toBe(0);
  });
});
