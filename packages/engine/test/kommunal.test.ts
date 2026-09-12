import { describe, expect, it } from "vitest";

import { defaultContext } from "../src/model/context.js";
import { tlakap_kr, tlkap_kl } from "../src/tjanstepension/kommunal.js";
import type { SchemeContext } from "../src/tjanstepension/types.js";

/** No fixture exists for these agreements; see the note in itp.test.ts. */
const context = (overrides: Partial<SchemeContext> = {}): SchemeContext => ({
  year: 2024,
  born: 1990,
  wStart: 23,
  tjpPar: 66,
  flexPension: 0,
  marginal: 0,
  ...overrides,
});

const model = defaultContext();
const IBB = 76_200;
const SALARY = 462_000;

describe("tlkap_kl", () => {
  it("did not exist before 1998", () => {
    expect(tlkap_kl(40, SALARY, IBB, 66, 1997, context({ year: 1997 }))).toBe(0);
    expect(tlkap_kl(40, SALARY, IBB, 66, 1998, context({ year: 1998 }))).toBeGreaterThan(0);
  });

  it("pays nothing before 21", () => {
    expect(tlkap_kl(20, SALARY, IBB, 66, 2024, context({ wStart: 18 }))).toBe(0);
    expect(tlkap_kl(22, SALARY, IBB, 66, 2024, context({ wStart: 18 }))).toBeGreaterThan(0);
  });

  it("pays nothing before 21 even for the cohorts AKAP-KR admits earlier", () => {
    // The `born > 1985 && alder < 21` branch is unreachable here because of the
    // early exit above it. AKAP-KR, which has no such exit, does pay. See the
    // note in kommunal.ts.
    const young = context({ born: 1990, wStart: 18, year: 2014 });
    expect(tlkap_kl(20, SALARY, IBB, 66, 2014, young)).toBe(0);
    expect(tlakap_kr(20, SALARY, IBB, 66, 2014, young, model)).toBeGreaterThan(0);
  });

  it("rises with salary and steps up above the breakpoint", () => {
    const step = 60_000;
    const breakpoint = 7.5 * IBB;
    const belowRate =
      (tlkap_kl(40, breakpoint - step, IBB, 66, 2024, context()) -
        tlkap_kl(40, breakpoint - 2 * step, IBB, 66, 2024, context())) /
      step;
    expect(belowRate).toBeCloseTo(0.045, 3);
  });

  it("caps pensionable salary at 30 income base amounts", () => {
    const ctx = context();
    expect(tlkap_kl(40, 40 * IBB, IBB, 66, 2024, ctx)).toBe(
      tlkap_kl(40, 30 * IBB, IBB, 66, 2024, ctx),
    );
  });

  it("accrues to 67 from 2003 but only to 65 before", () => {
    expect(tlkap_kl(66, SALARY, IBB, 67, 2002, context({ year: 2002 }))).toBe(0);
    expect(tlkap_kl(66, SALARY, IBB, 67, 2003, context({ year: 2003 }))).toBeGreaterThan(0);
  });

  it("gives cohorts born up to 1946 the lower rate above the breakpoint from 2005", () => {
    // The branch reads `year = 2006 Or born <= 1946`, so the `or` catches every
    // year from 2005 onward for those cohorts.
    const old = context({ born: 1946, year: 2010 });
    const young = context({ born: 1947, year: 2010 });
    const high = 12 * IBB;
    expect(tlkap_kl(60, high, IBB, 66, 2010, old)).toBeLessThan(
      tlkap_kl(60, high, IBB, 66, 2010, young),
    );
  });
});

describe("tlakap_kr", () => {
  it("pays from the first working day for those born after 1985, from 2014", () => {
    const ctx = context({ born: 1990, wStart: 19, year: 2014 });
    expect(tlakap_kr(19, SALARY, IBB, 66, 2014, ctx, model)).toBeGreaterThan(0);
    expect(tlakap_kr(18, SALARY, IBB, 66, 2014, ctx, model)).toBe(0);
  });

  it("moves those cohorts to the high rate above the breakpoint in 2014", () => {
    const high = 12 * IBB;
    const before = tlakap_kr(40, high, IBB, 66, 2013, context({ year: 2013 }), model);
    const after = tlakap_kr(40, high, IBB, 66, 2014, context({ year: 2014 }), model);
    expect(after).toBeGreaterThan(before);
  });

  it("raises both rates from 2023", () => {
    const ctx2022 = context({ year: 2022 });
    const ctx2023 = context({ year: 2023 });
    const below = 5 * IBB;
    expect(tlakap_kr(40, below, IBB, 66, 2023, ctx2023, model)).toBeGreaterThan(
      tlakap_kr(40, below, IBB, 66, 2022, ctx2022, model),
    );
  });

  it("drops the high rate above the LAS age from 2023", () => {
    const lasAge = model.riktage + 3;
    const high = 12 * IBB;
    const ctx = context({ year: 2023, tjpPar: 75 });
    const atLas = tlakap_kr(lasAge, high, IBB, 75, 2023, ctx, model);
    const pastLas = tlakap_kr(lasAge + 1, high, IBB, 75, 2023, ctx, model);
    expect(pastLas).toBeLessThan(atLas);
  });

  it("removes the upper age limit from 2023", () => {
    // Before 2023 rights stopped at the LAS age of 67.
    const ctx2022 = context({ year: 2022, tjpPar: 70 });
    const ctx2023 = context({ year: 2023, tjpPar: 70 });
    expect(tlakap_kr(69, SALARY, IBB, 70, 2022, ctx2022, model)).toBe(0);
    expect(tlakap_kr(69, SALARY, IBB, 70, 2023, ctx2023, model)).toBeGreaterThan(0);
  });

  it("keeps both 2006 branches live, unlike KAP-KL", () => {
    // AKAP-KR joins the two conditions with `and`, so a cohort after 1946 takes
    // the second branch rather than being swallowed by the first.
    const high = 12 * IBB;
    const old = tlakap_kr(60, high, IBB, 66, 2006, context({ born: 1946, year: 2006 }), model);
    const young = tlakap_kr(55, high, IBB, 66, 2006, context({ born: 1947, year: 2006 }), model);
    expect(old).not.toBe(young);
  });
});
