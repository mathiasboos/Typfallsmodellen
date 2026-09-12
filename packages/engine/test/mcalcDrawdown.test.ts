import { describe, expect, it } from "vitest";

import annuityTables from "../../data/annuity-tables.json" with { type: "json" };
import { loadDeathProbabilities } from "../src/data/nodeLoader.js";
import { defaultContext } from "../src/model/context.js";
import type { ModelContext } from "../src/model/context.js";
import { Scheme, defaultInput } from "../src/model/input.js";
import type { TypfallInput } from "../src/model/input.js";
import { prepareRun, simulateYear } from "../src/model/mcalc.js";
import { startsetup } from "../src/model/setup.js";
import { createRunState } from "../src/model/state.js";
import { DeltalTables } from "../src/pension/deltal.js";
import type { AnnuityTablesData } from "../src/pension/deltal.js";

const deaths = loadDeathProbabilities();
const published = annuityTables as unknown as AnnuityTablesData;

/**
 * The drawdown half of the loop has no offline fixture.
 *
 * The Brutto sheet cannot supply one: it exercises `IP_` for a typfall that
 * never retires, and pairs each year's pension right with the *following*
 * year's indexation, where Mcalc credits the right at the current age and
 * indexes with this year's level over last year's. The two are a year apart by
 * construction, so a per-age balance comparison is not the same calculation.
 * `incomePension.test.ts` checks `IP_` itself against that sheet term by term;
 * what the loop does with it waits on the golden files.
 *
 * So these tests assert structure -- when each pension starts, that balances
 * are drawn down by what is paid, that nothing goes negative -- and the quirks
 * worth pinning down.
 */
function walk(input: TypfallInput, context: ModelContext, lastAge = 105) {
  const setup = startsetup(input, context);
  const state = createRunState(setup.vectors.startage, setup.vectors.slutage);
  const run = prepareRun(
    setup,
    input,
    context,
    state,
    new DeltalTables(published, deaths, context),
  );
  for (let age = run.v.startage; age <= lastAge; age += 1) simulateYear(run, age);
  return { run, state };
}

const plain = (
  input: Parameters<typeof defaultInput>[0] = {},
  context: Parameters<typeof defaultContext>[0] = {},
) => walk(defaultInput(input), defaultContext(context));

describe("when each pension starts", () => {
  it("pays no public pension before the retirement age", () => {
    const { state } = plain();
    for (let age = 15; age <= 65; age += 1) {
      expect(state.ip.get(age), `ip at ${age}`).toBe(0);
      expect(state.pp.get(age), `pp at ${age}`).toBe(0);
      expect(state.garp.get(age), `garp at ${age}`).toBe(0);
      expect(state.ptillagg.get(age), `ptillagg at ${age}`).toBe(0);
    }
    expect(state.ip.get(66)).toBeGreaterThan(0);
    expect(state.pp.get(66)).toBeGreaterThan(0);
  });

  it("keeps paying for the rest of the run", () => {
    const { state } = plain();
    for (let age = 66; age <= 105; age += 1) {
      expect(state.ip.get(age), `ip at ${age}`).toBeGreaterThan(0);
    }
  });

  it("pays no occupational pension without an agreement", () => {
    const { state } = plain();
    for (let age = 15; age <= 105; age += 1) expect(state.tjp.get(age)).toBe(0);
  });

  it("pays the occupational pension from its own retirement age", () => {
    const { state } = plain({ born: 1985, scheme: Scheme.Itp1 }, { tjpPar: 62 });
    expect(state.tjp.get(61)).toBe(0);
    expect(state.tjp.get(62)).toBeGreaterThan(0);
    expect(state.tjp.get(70)).toBeGreaterThan(0);
  });

  it("stops a temporary occupational withdrawal after its years", () => {
    const { state } = plain({ born: 1985, scheme: Scheme.Itp1 }, { tempTjpUttag: 5 });
    expect(state.tjp.get(66)).toBeGreaterThan(0);
    expect(state.tjp.get(70)).toBeGreaterThan(0);
    expect(state.tjp.get(72)).toBe(0);
  });
});

describe("balances", () => {
  it("never go negative", () => {
    const { state } = plain({ born: 1985, scheme: Scheme.SafLo }, { ipsMonthly: 2000, ipsStart: 2015 });
    for (let age = 15; age <= 105; age += 1) {
      for (const [name, v] of [
        ["ipPbh", state.ipPbh],
        ["ppPbh", state.ppPbh],
        ["tjpPbh", state.tjpPbh],
        ["ipsPbh", state.ipsPbh],
        ["ppsPbh", state.ppsPbh],
      ] as const) {
        expect(v.get(age), `${name} at ${age}`).toBeGreaterThanOrEqual(0);
      }
    }
  });

  it("build up while working and fall once drawn", () => {
    const { state } = plain();
    expect(state.ipPbh.get(65)).toBeGreaterThan(state.ipPbh.get(50));
    expect(state.ppPbh.get(65)).toBeGreaterThan(state.ppPbh.get(50));
    // The premium pension pot is drawn down from retirement.
    expect(state.ppPbh.get(80)).toBeLessThan(state.ppPbh.get(66));
  });

  it("takes an opening balance typed in for one year", () => {
    const { state } = plain({}, { pbhYear: 2020, pbhIp: 1_500_000, pbhPp: 400_000 });
    expect(state.ipPbh.get(2020 - 1959)).toBe(1_500_000);
    expect(state.ppPbh.get(2020 - 1959)).toBe(400_000);
  });

  it("QUIRK: an opening balance divides the garantipension base by the income one", () => {
    // `GP_pbh(age) = rng_PBH_IP / IP_pbh(age)`, where the commented-out
    // alternative multiplies by 185/160. It only fires when a balance is typed
    // in, and it is what the workbook computes.
    const age = 2020 - 1959;
    const { state } = plain({}, { pbhYear: 2020, pbhIp: 1_500_000 });
    const untouched = plain().state;
    expect(untouched.gpPbh.get(age)).toBeGreaterThan(1_000_000);
    // The division leaves a ratio near one, which the rounding then truncates.
    expect(state.gpPbh.get(age)).toBeLessThan(2);
  });

  it("holds every figure to whole kronor unless the rounding is off", () => {
    const { state } = plain();
    for (const age of [40, 60, 66, 80]) {
      expect(state.ipPbh.get(age) % 1).toBe(0);
      expect(state.ip.get(age) % 12).toBe(0);
      expect(state.pp.get(age) % 12).toBe(0);
    }
    const exact = plain({}, { marginal: 1 }).state;
    expect(exact.ip.get(70) % 12).not.toBe(0);
  });
});

describe("the ATP cohorts", () => {
  it("pay no ATP to a cohort entirely in the new system", () => {
    const { run, state } = plain();
    expect(run.p.andelnya).toBe(1);
    for (let age = 66; age <= 105; age += 1) expect(state.tp.get(age)).toBe(0);
  });

  it("pay ATP to a cohort with a share of the old system", () => {
    const { run, state } = plain({ born: 1945 });
    expect(run.p.andelnya).toBeCloseTo(0.55, 9);
    expect(state.tp.get(run.p.par)).toBeGreaterThan(0);
  });

  it("QUIRK: the pre-1994 points are summed unsorted, the ATP points sorted", () => {
    // The VBA's comment above the first loop says "Sorteras i stigande ordning"
    // and no sort follows; only the ATP loop calls QuickSort. So the
    // garantibelopp averages the last fifteen *ages* and the ATP the fifteen
    // *best* points.
    const { run, state } = plain({ born: 1945 });
    const par = run.p.par;
    // After the sort the points vector is ascending over its whole span.
    for (let age = run.v.startage + 1; age <= run.v.slutage; age += 1) {
      expect(state.tpPoints.get(age)).toBeGreaterThanOrEqual(state.tpPoints.get(age - 1));
    }
    // The copy FTJP reads is untouched by it, and is not ascending.
    const stp: number[] = [];
    for (let age = 28; age <= 64; age += 1) stp.push(state.stpPoints.get(age));
    expect(stp.some((x, i) => i > 0 && x < stp[i - 1]!)).toBe(true);
    expect(state.gbelopp.get(par)).toBeGreaterThanOrEqual(0);
  });
});

describe("garantipension", () => {
  it("tops up a small pension and not a large one", () => {
    const small = plain({ monthlySalary: 12_000 }).state;
    const large = plain({ monthlySalary: 60_000 }).state;
    expect(small.garp.get(66)).toBeGreaterThan(0);
    expect(large.garp.get(66)).toBe(0);
  });

  it("is not paid before the riktålder", () => {
    // Retiring at 63 is allowed for this cohort, and its riktålder is the one
    // in force in that year -- 65, because 2022 is before the step to 66.
    const { run, state } = plain({ retirementAge: 63, monthlySalary: 12_000 });
    expect(run.p.riktalder).toBe(65);
    expect(state.garp.get(63)).toBe(0);
    expect(state.garp.get(64)).toBe(0);
    expect(state.garp.get(65)).toBeGreaterThan(0);
  });
});

describe("private saving", () => {
  it("pays nothing without a saving plan", () => {
    const { state } = plain();
    for (let age = 15; age <= 105; age += 1) {
      expect(state.ips.get(age)).toBe(0);
      expect(state.pps.get(age)).toBe(0);
    }
  });

  it("annuitises an IPS pot over the expected remaining life", () => {
    const { run, state } = plain({}, { ipsMonthly: 2000, ipsStart: 2000 });
    const expected = run.deltalTables.expectedLife(1959, 66);
    expect(expected).toBeGreaterThan(15);
    expect(state.ipsPbh.get(65)).toBeGreaterThan(0);
    expect(state.ips.get(66)).toBeGreaterThan(0);
    // Roughly the balance over the remaining life, uprated for half of it.
    expect(state.ips.get(66)).toBeCloseTo(
      Math.round((state.ipsPbh.get(65) * 1.017 ** (expected / 2)) / expected / 12) * 12,
      -1,
    );
  });

  it("routes an ISK or KF pot to the untaxed payout instead", () => {
    const isk = plain({}, { ipsMonthly: 2000, ipsStart: 2000, privateSavingKind: 2 }).state;
    expect(isk.ppsPbh.get(65)).toBeGreaterThan(0);
    expect(isk.pps.get(66)).toBeGreaterThan(0);
    expect(isk.ipsPbh.get(65)).toBe(0);
  });
});

describe("gross income", () => {
  it("is the sum of every component paid", () => {
    const { run, state } = plain({ born: 1985, scheme: Scheme.Itp1 }, { ipsMonthly: 1000, ipsStart: 2015 });
    for (const age of [50, 66, 70, 90]) {
      const parts =
        run.v.income.get(age) +
        state.ip.get(age) +
        state.tp.get(age) +
        state.garp.get(age) +
        state.pp.get(age) +
        state.tjp.get(age) +
        state.ptillagg.get(age) +
        state.ips.get(age);
      expect(state.brutto.get(age)).toBeCloseTo(parts, 6);
    }
  });

  it("leaves an ISK or KF payout out, because it is not taxed as income", () => {
    const { run, state } = plain({}, { ipsMonthly: 2000, ipsStart: 2000, privateSavingKind: 2 });
    expect(state.pps.get(70)).toBeGreaterThan(0);
    expect(state.brutto.get(70)).toBeCloseTo(
      run.v.income.get(70) + state.ip.get(70) + state.pp.get(70) + state.garp.get(70) + state.ptillagg.get(70),
      6,
    );
  });

  it("is finite and non-negative for every cohort, age and agreement", () => {
    for (const born of [1940, 1955, 1970, 1990, 2005]) {
      for (const scheme of [Scheme.None, Scheme.Itp2, Scheme.SafLo, Scheme.KapKl, Scheme.Pa16Avd1]) {
        const { run, state } = plain({ born, scheme });
        for (let age = run.v.startage; age <= 105; age += 1) {
          const value = state.brutto.get(age);
          expect(Number.isFinite(value), `born ${born} scheme ${scheme} age ${age}`).toBe(true);
          expect(value).toBeGreaterThanOrEqual(0);
        }
      }
    }
  });
});
