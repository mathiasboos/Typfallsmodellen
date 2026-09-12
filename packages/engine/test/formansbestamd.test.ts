import { describe, expect, it } from "vitest";

import annuityTables from "../../data/annuity-tables.json" with { type: "json" };
import { loadDeathProbabilities } from "../src/data/nodeLoader.js";
import { defaultContext } from "../src/model/context.js";
import { DeltalTables } from "../src/pension/deltal.js";
import type { AnnuityTablesData } from "../src/pension/deltal.js";
import { calculateDeltal } from "../src/pension/mortality.js";
import {
  DC_underlag,
  FTJP,
  KAPKL_f,
  PA_KL,
  PA_KLBPP,
  Scheme,
  tlPA03,
} from "../src/tjanstepension/index.js";
import type { RunVectors, SchemeContext } from "../src/tjanstepension/index.js";

const model = defaultContext();
const deaths = loadDeathProbabilities();
const tables = new DeltalTables(
  annuityTables as unknown as AnnuityTablesData,
  deaths,
  model,
);

const BORN = 1959;
const IBB = 76_200;
const PBB = 58_800;
const FPB = 60_000;

/**
 * A simple typfall: a flat salary from 23 to 65, with the basbelopp held
 * constant so the arithmetic under test is not entangled with indexation.
 */
const vectors = (salary = 462_000, wStart = 23): RunVectors => ({
  wage: (age) => (age >= wStart && age <= 65 ? salary : 0),
  ibb: () => IBB,
  pbb: () => PBB,
  fpb: () => FPB,
  kpiJune: () => 400,
  kpi: () => 400,
  year: (age) => BORN + age,
});

const context = (overrides: Partial<SchemeContext> = {}): SchemeContext => ({
  year: 2025,
  born: BORN,
  wStart: 23,
  tjpPar: 66,
  flexPension: 0,
  marginal: 0,
  ...overrides,
});

describe("the fixed mortality cell FTJP reads", () => {
  it("is the income pension divisor at 65 for the cohort, which is computed not stored", () => {
    // The workbook reads mortality!J27, part of a block that INDEXes into the
    // table Calculate_Deltal writes. For the default cohort of 1959 the sheet
    // shows 17.6 at age 65, which is what the ported calculation gives.
    const own = calculateDeltal(1959, deaths);
    expect(own.dtalip[0]![65]).toBeCloseTo(17.6, 2);
  });
});

describe("DC_underlag", () => {
  it("averages the five years before retirement", () => {
    expect(DC_underlag(0, 66, vectors(), 0)).toBeCloseTo(462_000, -1);
  });

  it("gives nothing when any of the six years is missing", () => {
    const late = vectors(462_000, 62);
    expect(DC_underlag(0, 66, late, 0)).toBe(0);
  });

  it("caps a salary that rose faster than the agreement allows", () => {
    // A sharp rise in the final years should be held back.
    const jumpy: RunVectors = {
      ...vectors(),
      wage: (age) => (age >= 23 && age <= 65 ? (age >= 62 ? 1_500_000 : 400_000) : 0),
    };
    const capped = DC_underlag(0, 66, jumpy, 0);
    const uncapped = DC_underlag(1, 66, jumpy, 0);
    expect(capped).toBeLessThan(uncapped);
  });

  it("does not cap in the state sector", () => {
    expect(DC_underlag(1, 66, vectors(), 0)).toBeCloseTo(462_000, -1);
  });
});

describe("tlPA03", () => {
  it("pays nothing to avdelning 1", () => {
    expect(tlPA03(600_000, 30, IBB, 1959, Scheme.Pa16Avd1)).toBe(0);
  });

  it("pays nothing on no income", () => {
    expect(tlPA03(0, 30, IBB, 1959, Scheme.Pa16Avd2)).toBe(0);
  });

  it("falls cohort by cohort as the benefit is phased out", () => {
    let previous = Number.POSITIVE_INFINITY;
    for (const cohort of [1950, 1955, 1960, 1965, 1970, 1973]) {
      const benefit = tlPA03(600_000, 30, IBB, cohort, Scheme.Pa16Avd2);
      expect(benefit, `cohort ${cohort}`).toBeLessThan(previous);
      previous = benefit;
    }
  });

  it("replaces far more of the salary above the breakpoint", () => {
    const below = tlPA03(7 * IBB, 30, IBB, 1959, Scheme.Pa16Avd2);
    const above = tlPA03(10 * IBB, 30, IBB, 1959, Scheme.Pa16Avd2);
    expect((above - below) / (3 * IBB)).toBeGreaterThan(0.5);
  });

  it("scales down below thirty years of service", () => {
    const full = tlPA03(600_000, 30, IBB, 1959, Scheme.Pa16Avd2);
    expect(tlPA03(600_000, 15, IBB, 1959, Scheme.Pa16Avd2)).toBeCloseTo(full / 2, 6);
  });
});

describe("KAPKL_f", () => {
  it("pays nothing below the breakpoint", () => {
    expect(KAPKL_f(7 * IBB, 30, IBB, 1959)).toBe(0);
  });

  it("pays nothing to cohorts born 1986 or later, who are on AKAP-KL", () => {
    expect(KAPKL_f(20 * IBB, 30, IBB, 1986)).toBe(0);
    expect(KAPKL_f(20 * IBB, 30, IBB, 1985)).toBeGreaterThan(0);
  });

  it("falls cohort by cohort", () => {
    expect(KAPKL_f(15 * IBB, 30, IBB, 1960)).toBeLessThan(KAPKL_f(15 * IBB, 30, IBB, 1950));
  });

  it("scales down below thirty years of service", () => {
    const full = KAPKL_f(15 * IBB, 30, IBB, 1959);
    expect(KAPKL_f(15 * IBB, 15, IBB, 1959)).toBeCloseTo(full / 2, 6);
  });
});

describe("PA_KLBPP", () => {
  it("rises with salary", () => {
    expect(PA_KLBPP(66, vectors(900_000))).toBeGreaterThan(PA_KLBPP(66, vectors(300_000)));
  });

  it("takes only the single best year when the seventh is missing", () => {
    // The fallbacks meant for a short working life can never fire -- the VBA
    // writes `&` where `And` was intended, so the chain falls through to the
    // final Else and averages one year. See the note in formansbestamd.ts.
    const rising = (age: number) => (age >= 59 && age <= 64 ? 100_000 * (age - 58) : 0);
    const short: RunVectors = { ...vectors(), wage: rising };

    // The same points, but with all seven years present at the best salary, so
    // the five-year average is also the best year. If the fallback were doing
    // what it looks like it does, these two would differ.
    const bestOnly: RunVectors = { ...vectors(), wage: (age) => (age >= 58 && age <= 64 ? 600_000 : 0) };

    expect(PA_KLBPP(66, short)).toBe(PA_KLBPP(66, bestOnly));

    // And it is not the average of the six years that do carry a salary.
    const sixYearAverage: RunVectors = {
      ...vectors(),
      wage: (age) => (age >= 58 && age <= 64 ? 350_000 : 0),
    };
    expect(PA_KLBPP(66, short)).not.toBe(PA_KLBPP(66, sixYearAverage));
  });
});

describe("PA_KL", () => {
  it("pays nothing after PFA98 replaced it in 1998", () => {
    expect(PA_KL(66, 1998, 2, 23, vectors())).toBe(0);
  });

  it("rises with the years worked, up to thirty", () => {
    const short = PA_KL(50, 1996, 2, 40, vectors());
    const long = PA_KL(66, 1996, 2, 23, vectors());
    expect(long).toBeGreaterThan(short);
  });
});

describe("FTJP", () => {
  const noPoints = () => 0;
  const noTp = () => 0;

  it("leaves agreements without a defined benefit untouched", () => {
    for (const scheme of [Scheme.None, Scheme.Itp1, Scheme.Pa16Avd1]) {
      const result = FTJP(
        5000, 15, 66, 66, BORN, scheme, noPoints, noTp, vectors(), tables, context(), model, 12,
      );
      expect(result, `scheme ${scheme}`).toBe(5000);
    }
  });

  it("adds a benefit for ITP 2", () => {
    const result = FTJP(
      5000, 15, 66, 66, BORN, Scheme.Itp2, noPoints, noTp, vectors(), tables, context(), model, 12,
    );
    expect(result).toBeGreaterThan(5000);
  });

  it("adds a benefit for PA16 avdelning 2", () => {
    const result = FTJP(
      5000, 15, 66, 66, BORN, Scheme.Pa16Avd2, noPoints, noTp, vectors(), tables, context(), model, 12,
    );
    expect(result).toBeGreaterThan(5000);
  });

  it("adds a benefit for KAP-KL", () => {
    const points = (age: number) => (age >= 28 && age <= 64 ? 4 : 0);
    const result = FTJP(
      5000, 15, 66, 66, BORN, Scheme.KapKl, points, noTp, vectors(), tables, context(), model, 12,
    );
    expect(result).toBeGreaterThan(5000);
  });

  it("scales the benefit down for early retirement and up for late", () => {
    const at = (tjpPar: number, age: number) =>
      FTJP(0, 15, age, tjpPar, BORN, Scheme.Itp2, noPoints, noTp, vectors(), tables,
        context({ tjpPar }), model, 12);
    const early = at(63, 63);
    const late = at(68, 68);
    expect(early).toBeLessThan(late);
  });

  it("stays finite for every agreement across a range of retirement ages", () => {
    for (const scheme of Object.values(Scheme)) {
      for (const tjpPar of [63, 65, 66, 68, 70]) {
        const result = FTJP(
          1000, 15, tjpPar, tjpPar, BORN, scheme, noPoints, noTp, vectors(), tables,
          context({ tjpPar }), model, 12,
        );
        expect(Number.isFinite(result), `scheme ${scheme} at ${tjpPar}`).toBe(true);
        expect(result, `scheme ${scheme} at ${tjpPar}`).toBeGreaterThanOrEqual(0);
      }
    }
  });
});
