import { describe, expect, it } from "vitest";

import optionsJson from "../../data/options.json" with { type: "json" };
import {
  BIRTH_YEARS,
  RETIREMENT_AGES,
  SCHEME_CHOICES,
  Scheme,
  defaultInput,
} from "../src/model/input.js";
import { defaultContext } from "../src/model/context.js";
import { SLUTAGE, createRunState, runVectors } from "../src/model/state.js";

/**
 * The input contract is extracted from the workbook, not transcribed, so these
 * tests check that the extraction and the typed shape still agree -- and pin
 * down the handful of settings whose shipped value is a formula's result rather
 * than the setting itself.
 */

describe("defaultInput", () => {
  it("is the typfall the workbook opens with", () => {
    expect(defaultInput()).toEqual({
      born: 1959,
      retirementAge: 66,
      startWorkAge: 23,
      monthlySalary: 38_500,
      married: false,
      scheme: Scheme.None,
      yearlyInflation: 0,
      realGrowth: 0,
      realReturn: 0.017,
    });
  });

  it("takes every field from options.json rather than a transcription", () => {
    const n = optionsJson.normalDefaults;
    const input = defaultInput();
    expect(input.born).toBe(n.birthYear);
    expect(input.retirementAge).toBe(n.retirementAge);
    expect(input.startWorkAge).toBe(n.startWorkAge);
    expect(input.monthlySalary).toBe(n.monthlySalary);
    expect(input.married).toBe(n.married);
    expect(input.realReturn).toBe(n.realReturn);
  });

  it("is frozen, and overrides replace one field at a time", () => {
    const input = defaultInput({ retirementAge: 63 });
    expect(input.retirementAge).toBe(63);
    expect(input.born).toBe(1959);
    expect(Object.isFrozen(input)).toBe(true);
  });

  it("offers the drop-downs the Start sheet does", () => {
    expect(BIRTH_YEARS[0]).toBe(1930);
    expect(BIRTH_YEARS.at(-1)).toBe(2024);
    expect(RETIREMENT_AGES).toEqual([
      61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80,
    ]);
  });

  it("numbers the agreements the way the workbook's drop-down does", () => {
    expect(SCHEME_CHOICES.map((c) => c.value)).toEqual([1, 2, 3, 4, 5, 6, 7, 8]);
    expect(SCHEME_CHOICES[Scheme.None - 1]!.label).toMatch(/^Saknar/);
    expect(SCHEME_CHOICES[Scheme.SafLo - 1]!.label).toMatch(/^SAF-LO/);
    expect(SCHEME_CHOICES[Scheme.Pa16Avd1 - 1]!.label).toMatch(/PA16 \(Avd 1\)/);
  });
});

describe("defaultContext", () => {
  it("carries the workbook's normal settings", () => {
    const c = defaultContext();
    expect(c.marginal).toBe(0);
    expect(c.finalSalaryYears).toBe(5);
    expect(c.referenceYear).toBe(2025);
    expect(c.dela).toBe(2);
    expect(c.insuranceYears).toBe(40);
    expect(c.hyra).toBe(6300);
    expect(c.discountRate).toBe(0.05);
    expect(c.lastPensionRight).toBe(1);
    expect(c.priceBasis).toBe(1);
  });

  it("leaves the derived settings at 0 for setup to resolve", () => {
    // Each of these is a formula on Adv_settings, so the extracted value is
    // what the formula happened to produce for the shipped typfall -- 66 for
    // TJP_PAR because the retirement age is 66, 56 for the table start because
    // it is `INT(par - 10)`. Storing the result would freeze a normal-mode run
    // at the shipped retirement age.
    const c = defaultContext();
    expect(c.tjpPar).toBe(0);
    expect(c.wTime).toBe(0);
    expect(c.makeBorn).toBe(0);
    expect(c.table2StartAge).toBe(0);
    expect(c.modelYear).toBe(0);
    // The extracted values those formulas produced, for comparison.
    const adv = (name: string) =>
      optionsJson.advancedSettings.find((s) => s.name.toLowerCase() === name.toLowerCase())
        ?.default;
    expect(adv("TJP_PAR")).toBe(66);
    expect(adv("rng_tabell2_startAge")).toBe(56);
  });

  it("switches the advanced-mode extras off", () => {
    const c = defaultContext();
    expect(c.risk).toBe(0);
    expect(c.wealth).toBe(0);
    expect(c.rules).toBe(0);
    expect(c.rulesFromUtg).toBe(0);
    expect(c.rulesFromSkatt).toBe(0);
    expect(c.altLastPensionRight).toBe(0);
    expect(c.tlSpecYear).toBe(0);
    expect(c.pbhYear).toBe(0);
  });

  it("starts with no children and no private saving", () => {
    const c = defaultContext();
    expect(c.childBirthYears).toEqual([0, 0, 0, 0]);
    expect(c.ipsMonthly).toBe(0);
  });
});

describe("createRunState", () => {
  it("sizes every vector from startage to slutage", () => {
    const state = createRunState(15);
    expect(SLUTAGE).toBe(105);
    expect(state.ipPbh.firstAge).toBe(15);
    expect(state.ipPbh.lastAge).toBe(105);
    expect(state.ipPbh.get(15)).toBe(0);
    expect(() => state.ipPbh.get(14)).toThrow(RangeError);
    expect(() => state.ipPbh.get(106)).toThrow(RangeError);
  });

  it("stops the pre-1994 ATP points at 64, as the VBA's ReDim does", () => {
    const state = createRunState(15);
    expect(state.tp94p.lastAge).toBe(64);
    expect(() => state.tp94p.get(65)).toThrow(RangeError);
    // Every other vector runs the full span.
    expect(state.tpPoints.lastAge).toBe(105);
  });

  it("keeps the ATP points and their STP copy as separate vectors", () => {
    // FTJP reads STP_points after Mcalc has sorted TP_points in place, so the
    // two cannot be the same array.
    const state = createRunState(15);
    state.tpPoints.set(40, 6.5);
    expect(state.stpPoints.get(40)).toBe(0);
  });

  it("starts every running scalar at zero", () => {
    const state = createRunState(15);
    expect(state.atpYear).toBe(0);
    expect(state.atpPoints).toBe(0);
    expect(state.pgiYears).toBe(0);
    expect(state.uttagIp).toBe(0);
    expect(state.rows).toEqual([]);
  });
});

describe("runVectors", () => {
  it("reads the setup vectors through, and yields 0 outside their range", () => {
    const state = createRunState(15);
    const setup = {
      startage: 15,
      slutage: 105,
      year: state.ip,
      ibb: state.pp,
      pbb: state.tp,
      fpb: state.garp,
      kpi: state.brutto,
      kpiJune: state.netto,
      iindex: state.bidrag,
      pindex: state.indDisp,
      yieldFactor: state.ips,
      rgk: state.pps,
      mpgi: state.ptillagg,
      ipAvg: state.tjp,
      ppAvg: state.gbelopp,
      tpAvg: state.ipRatt,
      ipArv1: state.ppRatt,
      ipArv2: state.gpRatt,
      ppArv: state.tjpRatt,
      komSkatt: state.ipPbh,
      begravavg: state.ppPbh,
      taxLimit1: state.gpPbh,
      taxLimit2: state.tjpPbh,
      income: state.ipsPbh,
      wage: state.ppsPbh,
    };
    setup.wage.set(30, 400_000);
    setup.ibb.set(30, 80_600);
    const vectors = runVectors(setup);
    expect(vectors.wage(30)).toBe(400_000);
    expect(vectors.ibb(30)).toBe(80_600);
    // `x(age - 1)` at the lower edge is a real idiom in the VBA and must not throw.
    expect(vectors.wage(14)).toBe(0);
  });
});
