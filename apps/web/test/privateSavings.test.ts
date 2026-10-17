/**
 * The savings both "Hur länge räcker pensionen?" and "Bästa pensionsålder" ask for,
 * and how the model is asked to carry them -- the arithmetic that is the same in
 * both, checked without a browser or an engine run. What each tab then does with
 * the runs is checked in its own test file.
 */
import { describe, expect, it } from "vitest";

import { defaultContext } from "@typfallsmodellen/engine";

import {
  SAVINGS_KINDS,
  defaultSavings,
  savingSettings,
  savingYear,
  savingsKindName,
  taxFreeSaving,
  taxedSaving,
} from "../src/privateSavings.js";
import type { SavingsInput } from "../src/privateSavings.js";

const context = defaultContext();
const person = { born: 1980, startWorkAge: 25, retirementAge: 68 };
const saved = (patch: Partial<SavingsInput>): SavingsInput => ({ ...defaultSavings(), ...patch });
const isk = { savingsKind: "isk" as const, savingsCapital: 400_000, savingsMonthly: 1_500 };

describe("the kinds of private saving", () => {
  it("are two, as the money is taxed, the schablon-taxed ones first, each with the model's own number", () => {
    expect(SAVINGS_KINDS.map((k) => [k.id, k.kind, k.sv])).toEqual([
      ["isk", 2, "ISK / Kapitalförsäkring / Depåförsäkring"],
      ["depot", 3, "Depå / VP-konto / Fondkonto"],
    ]);
  });

  it("are named in both languages", () => {
    expect(savingsKindName("isk", "en")).toBe("ISK / Endowment insurance / Depot insurance");
    expect(savingsKindName("depot", "en")).toBe("Securities account / Fund account");
    expect(savingsKindName("depot", "sv")).toBe("Depå / VP-konto / Fondkonto");
  });

  it("open on an ISK with nothing in it, or in anything else", () => {
    expect(defaultSavings()).toEqual({
      individualCapital: 0,
      individualMonthly: 0,
      privateCapital: 0,
      privateMonthly: 0,
      savingsKind: "isk",
      savingsCapital: 0,
      savingsMonthly: 0,
    });
    expect(taxedSaving(defaultSavings())).toEqual({ capital: 0, monthly: 0 });
    expect(taxFreeSaving(defaultSavings())).toBeNull();
  });
});

describe("what is paid out as taxed income, and what is not", () => {
  it("adds up the individual pension and the private pension, and nothing else", () => {
    const both = saved({ individualCapital: 100_000, individualMonthly: 1_000, privateCapital: 50_000, privateMonthly: 500 });
    expect(taxedSaving(both)).toEqual({ capital: 150_000, monthly: 1_500 });
    // Private savings are not taxed as income, whichever kind they are in, so they are no part of what is.
    for (const savingsKind of ["isk", "depot"] as const) {
      expect(taxedSaving({ ...both, savingsKind, savingsCapital: 25_000, savingsMonthly: 250 })).toEqual({
        capital: 150_000,
        monthly: 1_500,
      });
    }
  });

  it("keeps the tax-free saving apart, as the model's ISK (2) or a depot (3), and none without anything in it", () => {
    expect(taxFreeSaving(saved(isk))).toEqual({ capital: 400_000, monthly: 1_500, kind: 2 });
    expect(taxFreeSaving(saved({ ...isk, savingsKind: "depot" }))).toEqual({ capital: 400_000, monthly: 1_500, kind: 3 });
    expect(taxFreeSaving(saved({ savingsKind: "depot", savingsMonthly: 1_000 }))).toMatchObject({ capital: 0, kind: 3 });
    expect(taxFreeSaving(saved({ savingsKind: "isk", savingsCapital: -5, savingsMonthly: -5 }))).toBeNull();
    expect(taxFreeSaving(saved({ savingsKind: "isk", savingsCapital: Number.NaN }))).toBeNull();
  });

  it("counts a negative or a broken amount as none", () => {
    expect(taxedSaving(saved({ individualCapital: -10, privateMonthly: Number.NaN, privateCapital: 5 }))).toEqual({ capital: 5, monthly: 0 });
  });
});

describe("the year the saving counts from", () => {
  it("is the model's today, or the retirement year if that has passed", () => {
    expect(savingYear(person, context)).toBe(context.ipsStart);
    expect(savingYear({ ...person, born: 1959, retirementAge: 66 }, context)).toBe(2025);
    expect(savingYear(person, { ...context, ipsStart: 2040 })).toBe(2040);
  });

  it("is never so early that the run cannot book the capital", () => {
    // Booked at the end of the year before, in a year the run covers: from the model's first age
    // (1) or the age work starts, whichever is earlier. 1990 + 1 + 1 is the earliest for a 1990 cohort.
    const early = { ...context, ipsStart: 1950 };
    expect(savingYear({ ...person, born: 1990 }, early)).toBe(1992);
    expect(savingYear({ ...person, born: 1990, startWorkAge: 15 }, { ...early, modelStartAge: 15 })).toBe(2006);
    expect(savingYear({ ...person, born: 1990 }, { ...context, ipsStart: 1992 })).toBe(1992);
    expect(savingYear({ ...person, born: 1990 }, { ...context, ipsStart: 1991 })).toBe(1992);
  });
});

describe("the model's settings for each part", () => {
  const all = saved({ individualCapital: 100_000, individualMonthly: 1_000, privateCapital: 50_000, privateMonthly: 500, ...isk });

  it("give the pension part the sum as IPS (0), counted from the saving year", () => {
    expect(savingSettings(context, person, all, "pension")).toEqual({
      privateSavingKind: 0,
      ipsMonthly: 1_500,
      ipsStart: context.ipsStart,
      ipsOpening: 150_000,
    });
  });

  it("give the savings part its own amounts as an ISK (2) or a depot (3)", () => {
    expect(savingSettings(context, person, all, "savings")).toMatchObject({ privateSavingKind: 2, ipsMonthly: 1_500, ipsOpening: 400_000 });
    expect(savingSettings(context, person, { ...all, savingsKind: "depot" }, "savings")).toMatchObject({ privateSavingKind: 3 });
  });

  it("give a savings part with nothing in it no saving at all", () => {
    expect(savingSettings(context, person, defaultSavings(), "savings")).toMatchObject({ privateSavingKind: 0, ipsMonthly: 0, ipsOpening: 0 });
  });

  it("read a monthly saving of 1 or less as none, not as a share of income", () => {
    const monthly = (n: number) => savingSettings(context, person, saved({ individualMonthly: n }), "pension").ipsMonthly;
    expect(monthly(1)).toBe(0);
    expect(monthly(0.5)).toBe(0);
    expect(monthly(2)).toBe(2);
  });

  it("follow the retirement age, so that each alternative counts the saving to its own", () => {
    const at = (retirementAge: number) => savingSettings({ ...context, ipsStart: 2060 }, { ...person, retirementAge }, all, "pension").ipsStart;
    expect(at(65)).toBe(1980 + 65);
    expect(at(70)).toBe(1980 + 70);
  });
});
