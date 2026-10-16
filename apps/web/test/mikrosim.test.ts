import { describe, expect, it } from "vitest";

import { defaultContext, defaultInput } from "@typfallsmodellen/engine";

import {
  mikrosimRowFromInput,
  mikrosimRowToContext,
  mikrosimRowToInput,
  newMikrosimRow,
  validateMikrosimRow,
} from "../src/mikrosim.js";
import type { MikrosimRow } from "../src/mikrosim.js";

/**
 * `mikrosimRowToInput`/`mikrosimRowToContext` are this feature's one piece of
 * real mapping logic -- a Mikrosim row is a fully independent typfall (see
 * mikrosim.ts's own file comment), so unlike `compare.ts`'s `applyScenario`
 * there is no baseline to diff against; what needs proving instead is that
 * the mapping touches exactly the nine Mikrosim fields and nothing else,
 * that the annual-to-monthly salary conversion is exact, and that the IPS
 * figure lands on the shared context rather than the input.
 */
describe("mapping a Mikrosim row onto the engine's own types", () => {
  function fullRow(): MikrosimRow {
    return {
      id: "1",
      born: 1970,
      startWorkAge: 20,
      retirementAge: 68,
      annualSalary: 480_000,
      yearlyInflation: 0.02,
      realGrowth: 0.01,
      realReturn: 0.03,
      ipsMonthly: 500,
      scheme: 3,
    };
  }

  it("maps all nine fields onto TypfallInput, converting Årslön to a monthly figure", () => {
    const row = fullRow();
    const input = mikrosimRowToInput(row);
    expect(input.born).toBe(row.born);
    expect(input.startWorkAge).toBe(row.startWorkAge);
    expect(input.retirementAge).toBe(row.retirementAge);
    expect(input.monthlySalary).toBe(row.annualSalary / 12);
    expect(input.yearlyInflation).toBe(row.yearlyInflation);
    expect(input.realGrowth).toBe(row.realGrowth);
    expect(input.realReturn).toBe(row.realReturn);
    expect(input.scheme).toBe(row.scheme);
  });

  it("never sets married, ownIncome, pgbManual or pgbConscription", () => {
    const input = mikrosimRowToInput(fullRow());
    expect(input.married).toBe(false);
    expect(input.ownIncome).toBeUndefined();
    expect(input.pgbManual).toBeUndefined();
    expect(input.pgbConscription).toBeUndefined();
  });

  it("puts ipsMonthly on the shared context and leaves every other field untouched", () => {
    const row = fullRow();
    const shared = defaultContext({ finalSalaryYears: 5, hyra: 7200 });
    const rowContext = mikrosimRowToContext(row, shared);
    expect(rowContext.ipsMonthly).toBe(row.ipsMonthly);
    const { ipsMonthly: _shared, ...restShared } = shared;
    const { ipsMonthly: _row, ...restRow } = rowContext;
    expect(restRow).toEqual(restShared);
  });

  it("is a pure function of the row and the shared context -- not a frozen snapshot", () => {
    const row = fullRow();
    const contextA = defaultContext({ finalSalaryYears: 3 });
    const contextB = defaultContext({ finalSalaryYears: 8 });
    expect(mikrosimRowToContext(row, contextA).finalSalaryYears).toBe(3);
    expect(mikrosimRowToContext(row, contextB).finalSalaryYears).toBe(8);
  });
});

describe("mikrosimRowFromInput", () => {
  it("is the reverse of mikrosimRowToInput -- round-trips a TypfallInput's own nine fields", () => {
    const input = defaultInput({
      born: 1975,
      startWorkAge: 21,
      retirementAge: 67,
      monthlySalary: 42_000,
      yearlyInflation: 0.015,
      realGrowth: 0.008,
      realReturn: 0.02,
      scheme: 5,
    });
    const row = mikrosimRowFromInput("1", input, 350);
    expect(row.born).toBe(input.born);
    expect(row.startWorkAge).toBe(input.startWorkAge);
    expect(row.retirementAge).toBe(input.retirementAge);
    expect(row.annualSalary).toBe(input.monthlySalary * 12);
    expect(row.yearlyInflation).toBe(input.yearlyInflation);
    expect(row.realGrowth).toBe(input.realGrowth);
    expect(row.realReturn).toBe(input.realReturn);
    expect(row.scheme).toBe(input.scheme);
    expect(row.ipsMonthly).toBe(350);
    expect(mikrosimRowToInput(row).monthlySalary).toBeCloseTo(input.monthlySalary);
  });

  it("takes ipsMonthly from its own argument, not from the TypfallInput", () => {
    const input = defaultInput();
    const row = mikrosimRowFromInput("1", input, 1200);
    expect(row.ipsMonthly).toBe(1200);
  });

  it("clamps a continuous field outside Mikrosim's own bounds rather than carrying it through", () => {
    const input = defaultInput({ startWorkAge: 5 });
    const row = mikrosimRowFromInput("1", input, 0);
    expect(row.startWorkAge).toBeGreaterThanOrEqual(15);
  });
});

describe("newMikrosimRow", () => {
  it("seeds every row identically from the workbook's own normal defaults, not a live baseline", () => {
    const a = newMikrosimRow("1");
    const b = newMikrosimRow("2");
    expect(a.id).not.toBe(b.id);
    const { id: _a, ...restA } = a;
    const { id: _b, ...restB } = b;
    expect(restA).toEqual(restB);
    expect(a.ipsMonthly).toBe(0);
  });
});

describe("validateMikrosimRow", () => {
  it("accepts a freshly seeded row", () => {
    expect(validateMikrosimRow(newMikrosimRow("1"), "sv")).toBeUndefined();
  });

  it("rejects a birth year outside the workbook's own dropdown span", () => {
    const row = { ...newMikrosimRow("1"), born: 1800 };
    expect(validateMikrosimRow(row, "sv")).toBeDefined();
  });

  it("rejects a start-of-work age after the retirement age", () => {
    const row = { ...newMikrosimRow("1"), startWorkAge: 70, retirementAge: 66 };
    expect(validateMikrosimRow(row, "sv")).toBeDefined();
  });

  it("rejects a scheme outside the workbook's own eight agreements", () => {
    const row = { ...newMikrosimRow("1"), scheme: 99 as MikrosimRow["scheme"] };
    expect(validateMikrosimRow(row, "en")).toBeDefined();
  });
});
