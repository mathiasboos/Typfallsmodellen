import { describe, expect, it } from "vitest";

import annuityTables from "../../data/annuity-tables.json" with { type: "json" };
import { loadDeathProbabilities } from "../src/data/nodeLoader.js";
import { defaultContext } from "../src/model/context.js";
import { DeltalTables } from "../src/pension/deltal.js";
import type { AnnuityTablesData } from "../src/pension/deltal.js";
import { deltal } from "../src/pension/incomePension.js";
import { premiumFor, Scheme, tjp_ddeltal, tjpkassa } from "../src/tjanstepension/index.js";
import type { SchemeContext, SchemeId } from "../src/tjanstepension/index.js";

const model = defaultContext();
const tables = new DeltalTables(
  annuityTables as unknown as AnnuityTablesData,
  loadDeathProbabilities(),
  model,
);

const context = (overrides: Partial<SchemeContext> = {}): SchemeContext => ({
  year: 2024,
  born: 1970,
  wStart: 23,
  tjpPar: 66,
  flexPension: 0,
  marginal: 0,
  ...overrides,
});

const IBB = 76_200;
const SALARY = 462_000;
const ALL_SCHEMES = Object.values(Scheme) as SchemeId[];

describe("premiumFor", () => {
  it("pays nothing for the no-occupational-pension option", () => {
    expect(premiumFor(Scheme.None, 40, SALARY, IBB, 66, 2024, context(), model)).toBe(0);
  });

  it("pays something for every other agreement at a working age", () => {
    for (const scheme of ALL_SCHEMES) {
      if (scheme === Scheme.None) continue;
      const premium = premiumFor(scheme, 40, SALARY, IBB, 66, 2024, context(), model);
      expect(premium, `scheme ${scheme}`).toBeGreaterThan(0);
    }
  });

  it("stays finite and non-negative across the model's whole range", () => {
    for (const scheme of ALL_SCHEMES) {
      for (let born = 1938; born <= 2005; born += 7) {
        for (let alder = 16; alder <= 70; alder += 3) {
          const year = born + alder;
          const premium = premiumFor(
            scheme,
            alder,
            SALARY,
            IBB,
            66,
            year,
            context({ year, born }),
            model,
          );
          expect(Number.isFinite(premium), `scheme ${scheme} born ${born} age ${alder}`).toBe(true);
          expect(premium, `scheme ${scheme} born ${born} age ${alder}`).toBeGreaterThanOrEqual(0);
        }
      }
    }
  });

  it("rises with salary for every agreement that pays", () => {
    for (const scheme of ALL_SCHEMES) {
      if (scheme === Scheme.None) continue;
      const low = premiumFor(scheme, 40, 300_000, IBB, 66, 2024, context(), model);
      const high = premiumFor(scheme, 40, 600_000, IBB, 66, 2024, context(), model);
      expect(high, `scheme ${scheme}`).toBeGreaterThan(low);
    }
  });

  it("adds flexpension to ITP 1 and SAF-LO, and to nothing else", () => {
    const plain = context({ year: 2020 });
    const flex = context({ year: 2020, flexPension: 0.02 });
    for (const scheme of ALL_SCHEMES) {
      const without = premiumFor(scheme, 40, SALARY, IBB, 66, 2020, plain, model);
      const withFlex = premiumFor(scheme, 40, SALARY, IBB, 66, 2020, flex, model);
      if (scheme === Scheme.Itp1 || scheme === Scheme.SafLo) {
        expect(withFlex, `scheme ${scheme}`).toBeGreaterThan(without);
      } else {
        expect(withFlex, `scheme ${scheme}`).toBe(without);
      }
    }
  });
});

describe("tjp_ddeltal", () => {
  it("passes the divisor through under the model's normal settings", () => {
    for (const scheme of ALL_SCHEMES) {
      expect(tjp_ddeltal(18.5, scheme, 2024, 66, model)).toBe(18.5);
    }
  });

  it("adjusts the divisor once the correction is switched on", () => {
    const adjusting = defaultContext({ adjustOccupationalDivisor: true });
    const adjusted = tjp_ddeltal(18.5, Scheme.Itp1, 2024, 66, adjusting);
    expect(adjusted).not.toBe(18.5);
    expect(adjusted).toBeGreaterThan(10);
    expect(adjusted).toBeLessThan(30);
  });

  it("values a temporary withdrawal directly rather than scaling", () => {
    const temporary = defaultContext({ tempTjpUttag: 5 });
    const divisor = tjp_ddeltal(18.5, Scheme.Itp1, 2024, 66, temporary);
    // Five years of payments discount to well under a lifelong divisor.
    expect(divisor).toBeGreaterThan(4);
    expect(divisor).toBeLessThan(6);
  });
});

describe("tjpkassa", () => {
  it("pays nothing on an empty pot", () => {
    expect(tjpkassa(66, 1970, 66, 0, Scheme.Itp1, 2036, tables, model, 12)).toBe(0);
  });

  it("pays nothing before the occupational retirement age", () => {
    expect(tjpkassa(66, 1970, 60, 1_000_000, Scheme.Itp1, 2030, tables, model, 12)).toBe(0);
  });

  it("uses the same divisor as the premium pension, by default", () => {
    // The two must not drift apart: tjpkassa reads the premium pension table and
    // tjp_ddeltal leaves it alone under normal settings.
    const pot = 1_200_000;
    const paid = tjpkassa(66, 1970, 68, pot, Scheme.Itp1, 2038, tables, model, 12);
    const divisor = deltal(66, 1970, 68, tables, "premium", 999);
    expect(divisor).toBeGreaterThan(0);
    expect(paid).toBeCloseTo(Math.floor(pot / divisor / 12 + 0.5) * 12, 6);
  });

  it("pays more from a bigger pot", () => {
    const small = tjpkassa(66, 1970, 68, 600_000, Scheme.Itp1, 2038, tables, model, 12);
    const large = tjpkassa(66, 1970, 68, 1_200_000, Scheme.Itp1, 2038, tables, model, 12);
    expect(large).toBeGreaterThan(small);
  });

  it("gives cohorts before 1938 a flat divisor of 13", () => {
    const pot = 1_300_000;
    const paid = tjpkassa(65, 1935, 66, pot, Scheme.Itp1, 2001, tables, model, 12);
    expect(paid).toBeCloseTo(Math.floor(pot / 13 / 12 + 0.5) * 12, 6);
  });

  it("scales the first year by the months of pension drawn", () => {
    const full = tjpkassa(66, 1970, 66, 1_000_000, Scheme.Itp1, 2036, tables, model, 12);
    const half = tjpkassa(66, 1970, 66, 1_000_000, Scheme.Itp1, 2036, tables, model, 6);
    expect(half).toBeCloseTo(full / 2, 6);
  });
});
