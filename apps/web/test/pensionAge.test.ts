/**
 * "Hur maxa pensionen?"'s arithmetic, checked without a browser.
 *
 * `apps/web`'s test environment has no DOM, so the panel and its charts are
 * checked against the real built page by `tools/build/verify-offline.mjs`
 * instead (the same split `kpis.test.ts` and `compare.test.ts` already make).
 * What is pinned here is everything `pensionAge.ts` computes before any of
 * that: which alternatives exist, what each one adds up to, who wins at each
 * lifespan, where the totals cross -- and the two shared-context overrides the
 * question depends on.
 *
 * The anchor numbers are the typfall a static report on the same question
 * used first (born 1980, work from 25, 38 500 kr/month, the workbook's own
 * assumptions, no collective agreement): this module has to reproduce that
 * report to the krona. They are model outputs, so a legitimate change to the
 * engine moves them -- `npm run compare` (the golden cases) is what keeps the
 * engine honest, and these move with it on purpose.
 */
import { describe, expect, it } from "vitest";

import { SCHEME_CHOICES, Scheme, defaultContext, lowestPensionAge } from "@typfallsmodellen/engine";

import { loadDeathProbabilities } from "../src/deaths.js";
import {
  AGREEMENTS,
  LAST_AGE,
  LIFESPAN_BOUNDS,
  NEAR_TIE,
  SCHEME_COHORTS,
  SCHEME_ORDER,
  alternativeAges,
  averageMonthly,
  bestAgeSpread,
  bestRuns,
  breakEven,
  centreBounds,
  cohortText,
  comparisonKey,
  comparisonRows,
  comparisonToCsv,
  comparisonToXlsxRows,
  computeComparison,
  computeStudy,
  defaultSchemes,
  defaultStudyInput,
  finalSalaryOf,
  kindsPaid,
  lifespanRow,
  matrixToCsv,
  monthlySalaryForFinalSalary,
  monthlySalaryForTarget,
  optimalRows,
  optimalToCsv,
  optimalToXlsxRows,
  pensionContext,
  referenceAge,
  savingsContext,
  schemeFitsCohort,
  schemeName,
  schemeSector,
  seriesToCsv,
  seriesToXlsxRows,
  slotOf,
  studyContext,
  studyTypfall,
  totalFor,
  totalSeries,
} from "../src/pensionAge.js";
import type { Study, StudyOption } from "../src/pensionAge.js";
import { LIFELONG, defaultLastsInput, normaliseLasts, runPayouts } from "../src/pensionLasts.js";
import { kronorCell } from "../src/xlsx.js";

const deaths = loadDeathProbabilities();
const context = defaultContext();
const base = defaultStudyInput();
const study = computeStudy(base, context, deaths);

const parsOf = (s: ReturnType<typeof computeStudy>) => s.options.map((o) => o.par);
const runsOf = (s: ReturnType<typeof computeStudy>) =>
  bestRuns(s).map((r) => `${r.from}-${r.to}:${r.option.par}`);

describe("the alternatives", () => {
  it("are the reference age and three years either side, clamped to what the cohort may draw", () => {
    expect(alternativeAges(1980, 68)).toEqual([65, 66, 67, 68, 69, 70, 71]);
    // 1980's lowest age is 65, so a reference of 66 loses its two earliest.
    expect(alternativeAges(1980, 66)).toEqual([65, 66, 67, 68, 69]);
    // And the sheet stops at 80.
    expect(alternativeAges(1980, 80)).toEqual([77, 78, 79, 80]);
    // A reference outside the allowed range is moved inside it first.
    expect(alternativeAges(1980, 40)).toEqual([65, 66, 67, 68]);
    expect(alternativeAges(1980, 99)).toEqual([77, 78, 79, 80]);
  });

  it("take the reference age moved inside what the cohort may draw", () => {
    expect(referenceAge(1980, 68)).toBe(68);
    expect(referenceAge(1980, 40)).toBe(65);
    expect(referenceAge(1980, 99)).toBe(80);
    expect(referenceAge(1980, 67.6)).toBe(68);
  });

  it("never offer an age below the cohort's own lowest", () => {
    for (const born of [1950, 1959, 1965, 1980, 1995]) {
      const ages = alternativeAges(born, 60);
      expect(Math.min(...ages)).toBeGreaterThanOrEqual(lowestPensionAge(born));
      expect(ages.length).toBeGreaterThan(0);
      expect(centreBounds(born).min).toBe(lowestPensionAge(born));
    }
  });

  it("take their colour from the offset, so a clamped neighbour does not repaint the rest", () => {
    expect(slotOf(-3)).toBe("m3");
    expect(slotOf(0)).toBe("z");
    expect(slotOf(3)).toBe("p3");
    expect(slotOf(9)).toBe("p3");
    expect(slotOf(-9)).toBe("m3");
    const clamped = computeStudy({ ...base, centre: 66 }, context, deaths);
    // 65 is offset -1 here, whatever its position in the list.
    expect(clamped.options.map((o) => [o.par, o.slot])).toEqual([
      [65, "m1"],
      [66, "z"],
      [67, "p1"],
      [68, "p2"],
      [69, "p3"],
    ]);
  });
});

describe("the typfall and context an alternative is run with", () => {
  it("carry the study's own fields and nothing else", () => {
    const typfall = studyTypfall({ ...base, scheme: Scheme.Itp1, monthlySalary: 41_000, realGrowth: 0.018 }, 70);
    expect(typfall).toMatchObject({
      born: 1980,
      startWorkAge: 25,
      monthlySalary: 41_000,
      retirementAge: 70,
      scheme: Scheme.Itp1,
      realGrowth: 0.018,
      married: false,
    });
    expect(typfall.ownIncome).toBeUndefined();
    expect(typfall.pgbManual).toBeUndefined();
  });

  it("drop the agreement when occupational pension is switched off", () => {
    expect(studyTypfall({ ...base, scheme: Scheme.SafLo, includeOccupational: false }, 68).scheme).toBe(Scheme.None);
    expect(studyTypfall({ ...base, scheme: Scheme.SafLo, includeOccupational: true }, 68).scheme).toBe(Scheme.SafLo);
  });

  it("force occupational pension to follow retirement", () => {
    const shared = { ...context, tjpPar: 62 };
    const used = studyContext(shared);
    expect(used.tjpPar).toBe(0);
    // Everything else is the shared context untouched.
    expect(used.priceBasis).toBe(shared.priceBasis);
    expect(used.referenceYear).toBe(shared.referenceYear);
    expect(used.marginal).toBe(shared.marginal);
    expect(used.ipsMonthly).toBe(shared.ipsMonthly);
  });

  it("leave the model's private saving to the study's own savings, whatever Avancerat says", () => {
    const shared = { ...context, ipsMonthly: 5000, privateSavingKind: 2, ipsOpening: 123_456 };
    // No savings typed: no saving at all, and no ISK either.
    const none = pensionContext(shared, base, 67);
    expect([none.privateSavingKind, none.ipsMonthly, none.ipsOpening]).toEqual([0, 0, 0]);
    expect(savingsContext(shared, base, 67).ipsOpening).toBe(0);
    // Typed: the individual and the private pension are one pot, taxed as income (kind 0); the savings are another, tax-free.
    const saved = {
      ...base,
      individualCapital: 100_000,
      individualMonthly: 1_000,
      privateCapital: 200_000,
      privateMonthly: 2_000,
      savingsKind: "depot" as const,
      savingsCapital: 50_000,
      savingsMonthly: 500,
    };
    const pension = pensionContext(shared, saved, 67);
    expect([pension.privateSavingKind, pension.ipsMonthly, pension.ipsOpening]).toEqual([0, 3_000, 300_000]);
    const savings = savingsContext(shared, saved, 67);
    expect([savings.privateSavingKind, savings.ipsMonthly, savings.ipsOpening]).toEqual([3, 500, 50_000]);
    // Both carry the study's own overrides.
    expect(pension.tjpPar).toBe(0);
    expect(savings.tjpPar).toBe(0);
  });

  it("count the saving from the model's own start year until the alternative's retirement", () => {
    const early = pensionContext(context, { ...base, individualCapital: 1 }, 62);
    const late = pensionContext(context, { ...base, individualCapital: 1 }, 70);
    expect(early.ipsStart).toBe(context.ipsStart);
    expect(late.ipsStart).toBe(context.ipsStart);
    // A cohort whose retirement falls before the model's start year counts from the retirement year.
    const old = pensionContext(context, { ...base, born: 1960, individualCapital: 1 }, 62);
    expect(old.ipsStart).toBe(Math.min(context.ipsStart, 1960 + 62));
  });
});

describe("the default study (the static report's typfall, the first year as Table 1 has it)", () => {
  it("runs seven alternatives around the riktålder", () => {
    expect(parsOf(study)).toEqual([65, 66, 67, 68, 69, 70, 71]);
    expect(study.riktalder).toBe(68);
    expect(study.options[study.reference]!.par).toBe(68);
    expect(study.rows).toHaveLength(16);
    expect(study.rows[0]!.lifespan).toBe(80);
    expect(study.rows[15]!.lifespan).toBe(95);
    expect(study.warnings).toEqual([]);
  });

  it("averages what a start pays a month over the years it is paid", () => {
    const sixtyFive = study.options[0]!;
    // 15 years of 12 months to 80, and the total is the matrix's own 2 977 236.
    expect(averageMonthly(sixtyFive, 80)).toBeCloseTo(2977236 / 180, 6);
    expect(averageMonthly(sixtyFive, 80)! * 15 * 12).toBeCloseTo(sixtyFive.cumulative[80]!, 4);
    // Nothing has been paid by the start age, so there is no monthly amount to speak of.
    expect(averageMonthly(sixtyFive, 65)).toBeUndefined();
    expect(averageMonthly(study.options[6]!, 70)).toBeUndefined();
  });

  it("adds up whole years: lifespan minus retirement age", () => {
    const sixtyFive = study.options[0]!;
    // 80 - 65 = 15 years, a year each from the model's age matrix.
    let years = 0;
    for (let age = 65; age < 80; age += 1) years += sixtyFive.annual[age]!;
    expect(sixtyFive.cumulative[80]).toBe(years);
    // Nothing is paid before the start, and the first year is paid in full by the next birthday.
    expect(sixtyFive.cumulative[65]).toBe(0);
    expect(sixtyFive.cumulative[66]).toBe(sixtyFive.annual[65]);
    const seventyOne = study.options[6]!;
    expect(seventyOne.cumulative[71]).toBe(0);
    let late = 0;
    for (let age = 71; age < 80; age += 1) late += seventyOne.annual[age]!;
    expect(seventyOne.cumulative[80]).toBe(late);
  });

  it("matches the report's totals to the krona", () => {
    expect(study.rows[0]!.totals).toEqual([2_977_236, 2_947_200, 2_906_316, 2_853_552, 2_765_892, 2_673_120, 2_573_640]);
    expect(study.rows[0]!.margin).toBe(30_036);
    expect(study.rows[3]!.totals).toEqual([3_554_376, 3_556_644, 3_551_808, 3_539_244, 3_490_152, 3_440_832, 3_390_660]);
    expect(study.rows[3]!.margin).toBe(2_268);
  });

  it("picks 65, then 66, then the riktålder, then 71 as the lifespan grows", () => {
    expect(runsOf(study)).toEqual(["80-82:65", "83-83:66", "84-86:68", "87-95:71"]);
  });

  it("calls two totals within 1 % practically equal", () => {
    const ties = study.rows.filter((r) => r.marginShare < NEAR_TIE).map((r) => r.lifespan);
    expect(ties).toEqual([81, 82, 83, 84, 85, 86, 87, 88, 89]);
    const first = study.rows[0]!;
    expect(first.marginShare).toBeCloseTo(first.margin / first.totals[first.best]!, 12);
  });

  it("finds the break-even lifespans the report lists", () => {
    const at = (earlier: number, later: number) => {
      const found = [...study.versusReference, ...study.chain].find((b) => b.earlier.par === earlier && b.later.par === later);
      return found?.at;
    };
    expect(at(65, 68)).toBeCloseTo(83.428, 2);
    expect(at(66, 68)).toBeCloseTo(83.701, 2);
    expect(at(67, 68)).toBeCloseTo(83.961, 2);
    expect(at(68, 69)).toBeCloseTo(86.978, 2);
    expect(at(68, 70)).toBeCloseTo(86.747, 2);
    expect(at(68, 71)).toBeCloseTo(86.531, 2);
    // The step from the riktålder to one year later is the slowest, not the fastest.
    expect(at(68, 69)!).toBeGreaterThan(at(69, 70)!);
    expect(study.versusReference).toHaveLength(6);
    expect(study.chain).toHaveLength(6);
  });

  it("starts every alternative at Table 1's own figure, as Prognos shows it", () => {
    const reference = study.options[study.reference]!;
    // Table 1 counts the last year's pension right at once; the yearly flow picks it up in year two, so its own
    // first row says 20 194. The first year is Table 1's -- what Prognos shows -- and the flow goes on from there.
    expect(reference.firstYearMonthly).toBe(20_398);
    expect(reference.table1Monthly).toBe(20_398);
    expect(reference.annual[reference.par + 1]! / 12).toBeLessThan(reference.firstYearMonthly);
    expect(reference.finalSalaryMonthly).toBeCloseTo(39_872.3, 1);
    expect(reference.table1Share).toBeCloseTo(0.5116, 3);
  });
});

describe("the first year", () => {
  const savings = { ...base, individualCapital: 300_000, individualMonthly: 2_000, savingsKind: "isk" as const, savingsCapital: 400_000, savingsMonthly: 1_000 };
  const cases: [string, typeof base][] = [
    ["before tax", base],
    ["after tax", { ...base, afterTax: true }],
    ["with an agreement", { ...base, scheme: Scheme.Itp1 }],
    ["with every kind of saving", { ...savings, scheme: Scheme.Itp1 }],
    ["after tax with savings", { ...savings, afterTax: true }],
  ];

  it.each(cases)("is Table 1's own figure for every alternative, %s -- the one Prognos shows", (_name, input) => {
    const result = computeStudy(input, context, deaths);
    for (const option of result.options) {
      expect(option.firstYearMonthly).toBeCloseTo(option.table1Monthly, 6);
      // The parts add up to it: the occupational and private pensions and the savings are Table 1's lines too.
      expect(option.occupational[option.par]).toBeCloseTo(option.table1OccupationalMonthly * 12, 6);
      expect(option.savings[option.par]).toBeCloseTo(option.table1SavingsMonthly * 12, 6);
      // Before tax the yearly flow goes on from there without a jump: the next year is within a few per cent.
      // (After tax it is not: the model's own tax rules change in the year after, which is not this tab's to smooth.)
      if (!input.afterTax) {
        const next = option.annual[option.par + 1]! / 12;
        expect(Math.abs(next / option.firstYearMonthly - 1)).toBeLessThan(0.05);
      }
    }
  });
});

describe("the economic assumptions", () => {
  it("move the break-evens years earlier with real wage growth", () => {
    const growth = computeStudy({ ...base, yearlyInflation: 0.02, realGrowth: 0.018 }, context, deaths);
    expect(runsOf(growth)).toEqual(["80-80:67", "81-81:68", "82-82:70", "83-95:71"]);
    const first = growth.versusReference.find((b) => b.earlier.par === 65)!.at!;
    expect(first).toBeCloseTo(79.451, 2);
    expect(first).toBeLessThan(study.versusReference.find((b) => b.earlier.par === 65)!.at! - 3);
  });

  it("pay no guarantee pension once the income pension keeps its value", () => {
    const growth = computeStudy({ ...base, yearlyInflation: 0.02, realGrowth: 0.018 }, context, deaths);
    expect(runsOf(growth)).not.toEqual(runsOf(study));
    // Standard assumptions erode the income pension ~1.6 %/year, so the yearly flow falls with age.
    const sixtyEight = study.options[3]!;
    expect(sixtyEight.annual[95]!).toBeLessThan(sixtyEight.annual[70]!);
  });
});

describe("the collective agreements", () => {
  const itp = computeStudy({ ...base, scheme: Scheme.Itp1 }, context, deaths);

  it("add exactly the occupational pension they pay", () => {
    for (let i = 0; i < study.options.length; i += 1) {
      const without = study.options[i]!;
      const withIt = itp.options[i]!;
      let occupational = 0;
      for (let age = withIt.par; age < 80; age += 1) occupational += withIt.occupational[age]!;
      expect(occupational).toBeGreaterThan(0);
      expect(withIt.cumulative[80]).toBe(without.cumulative[80]! + occupational);
    }
    expect(itp.rows[0]!.totals[0]).toBe(4_242_816);
  });

  it("can move the winning ages", () => {
    expect(runsOf(itp)).toEqual(["80-82:65", "83-84:66", "85-85:67", "86-87:68", "88-95:71"]);
  });

  it("are left out when occupational pension is switched off", () => {
    const off = computeStudy({ ...base, scheme: Scheme.Itp1, includeOccupational: false }, context, deaths);
    expect(off.rows.map((r) => r.totals)).toEqual(study.rows.map((r) => r.totals));
  });

  it("give every area but \"Saknar tjänstepension\" a pension of its own", () => {
    const totals = [Scheme.Itp1, Scheme.Itp2, Scheme.SafLo, Scheme.KapKl, Scheme.AkapKr, Scheme.Pa16Avd2, Scheme.Pa16Avd1].map(
      (scheme) => computeStudy({ ...base, scheme }, context, deaths).rows[0]!.totals[3]!,
    );
    for (const total of totals) expect(total).toBeGreaterThan(study.rows[0]!.totals[3]!);
    expect(new Set(totals).size).toBe(totals.length);
  });
});

describe("the shared context", () => {
  it("is not allowed to start occupational pension early or add private saving of its own", () => {
    const shared = { ...context, tjpPar: 62, ipsMonthly: 5000, privateSavingKind: 2, ipsOpening: 250_000 };
    const overridden = computeStudy({ ...base, scheme: Scheme.Itp1 }, shared, deaths);
    const plain = computeStudy({ ...base, scheme: Scheme.Itp1 }, context, deaths);
    expect(overridden.rows.map((r) => r.totals)).toEqual(plain.rows.map((r) => r.totals));
  });
});

describe("the person's own savings", () => {
  const plain = study;
  const pensionOnly = { ...base, individualCapital: 300_000, individualMonthly: 2_000 };
  const savingsOnly = { ...base, savingsKind: "isk" as const, savingsCapital: 400_000, savingsMonthly: 1_000 };
  /** Both kinds at once: the pension fields of `pensionOnly` and the savings of `savingsOnly`. */
  const everyKind = { ...pensionOnly, savingsKind: savingsOnly.savingsKind, savingsCapital: savingsOnly.savingsCapital, savingsMonthly: savingsOnly.savingsMonthly };
  const withPension = computeStudy(pensionOnly, context, deaths);
  const withSavings = computeStudy(savingsOnly, context, deaths);

  it("change nothing when there are none -- or one krona a month, which the model would read as a share of income", () => {
    const nothing = computeStudy({ ...base, privateMonthly: 1, savingsMonthly: 1, savingsKind: "depot" }, context, deaths);
    expect(nothing.rows.map((r) => r.totals)).toEqual(plain.rows.map((r) => r.totals));
    for (const option of nothing.options) {
      expect(option.individual.every((v) => v === 0)).toBe(true);
      expect(option.savings.every((v) => v === 0)).toBe(true);
      expect(option.table1SavingsMonthly).toBe(0);
    }
  });

  it("add a lifelong pension of their own from each alternative's retirement age, taxed as income", () => {
    withPension.options.forEach((option, i) => {
      const without = plain.options[i]!;
      expect(option.individual[option.par - 1]).toBe(0);
      expect(option.individual[option.par]).toBeGreaterThan(0);
      expect(option.individual[LAST_AGE - 1]).toBe(option.individual[option.par]);
      expect(option.savings.every((v) => v === 0)).toBe(true);
      for (let age = option.par; age <= LAST_AGE; age += 1) {
        expect(option.annual[age]).toBeCloseTo(without.annual[age]! + option.individual[age]!, 6);
      }
    });
  });

  it("are one pot of taxed income whichever of the two pension fields they are typed in", () => {
    const split = computeStudy(
      { ...base, individualCapital: 100_000, privateCapital: 200_000, individualMonthly: 500, privateMonthly: 1_500 },
      context,
      deaths,
    );
    expect(split.rows.map((r) => r.totals)).toEqual(withPension.rows.map((r) => r.totals));
  });

  it("pay a later start more, since it saves for longer and the capital earns longer", () => {
    const monthly = computeStudy({ ...base, individualMonthly: 3_000 }, context, deaths);
    const first = monthly.options.map((o) => o.individual[o.par]!);
    first.forEach((amount, i) => {
      if (i > 0) expect(amount).toBeGreaterThan(first[i - 1]!);
    });
    const savings = computeStudy({ ...base, savingsMonthly: 3_000 }, context, deaths);
    const kept = savings.options.map((o) => o.savings[o.par]!);
    kept.forEach((amount, i) => {
      if (i > 0) expect(amount).toBeGreaterThan(kept[i - 1]!);
    });
  });

  it("can move the best age: waiting is worth more with saving to wait for", () => {
    expect(runsOf(plain)).toEqual(["80-82:65", "83-83:66", "84-86:68", "87-95:71"]);
    expect(runsOf(withPension)).toEqual(["80-82:65", "83-83:67", "84-85:68", "86-95:71"]);
  });

  it("are paid out tax-free when they are private savings, so the same before and after tax", () => {
    const net = computeStudy({ ...savingsOnly, afterTax: true }, context, deaths);
    const plainNet = computeStudy({ ...base, afterTax: true }, context, deaths);
    withSavings.options.forEach((option, i) => {
      expect(option.savings[option.par]).toBeGreaterThan(0);
      expect(net.options[i]!.savings).toEqual(option.savings);
      // The tax on everything else is what it was: the savings come on top of the net pension.
      for (let age = option.par; age <= LAST_AGE; age += 1) {
        expect(net.options[i]!.annual[age]! - net.options[i]!.savings[age]!).toBeCloseTo(plainNet.options[i]!.annual[age]!, 6);
        expect(option.annual[age]).toBeCloseTo(plain.options[i]!.annual[age]! + option.savings[age]!, 6);
      }
    });
  });

  it("spend the money until it is gone, where a pension of their own is paid for life", () => {
    const option = withSavings.options[3]!;
    expect(option.savings[option.par]).toBeGreaterThan(0);
    expect(option.savings[LAST_AGE - 1]).toBe(0);
    const lifelong = withPension.options[3]!;
    expect(lifelong.individual[LAST_AGE - 1]).toBeGreaterThan(0);
  });

  it("take the tax on a pension of their own out of the after-tax total, but not the tax-free savings", () => {
    const grossGain = computeStudy(pensionOnly, context, deaths).rows[0]!.totals[3]! - plain.rows[0]!.totals[3]!;
    const netPlain = computeStudy({ ...base, afterTax: true }, context, deaths);
    const netGain = computeStudy({ ...pensionOnly, afterTax: true }, context, deaths).rows[0]!.totals[3]! - netPlain.rows[0]!.totals[3]!;
    expect(netGain).toBeGreaterThan(0);
    expect(netGain).toBeLessThan(grossGain);

    const savingsGain = computeStudy({ ...savingsOnly, afterTax: true }, context, deaths).rows[0]!.totals[3]! - netPlain.rows[0]!.totals[3]!;
    const savingsGross = withSavings.rows[0]!.totals[3]! - plain.rows[0]!.totals[3]!;
    expect(savingsGain).toBeCloseTo(savingsGross, 6);
  });

  it("pay out as the other tab pays them, for the same savings and the same retirement age", () => {
    const savings = {
      individualCapital: 250_000,
      individualMonthly: 1_500,
      privateCapital: 50_000,
      privateMonthly: 500,
      savingsCapital: 120_000,
      savingsMonthly: 800,
    };
    const shared = { ...context, priceBasis: 1, tempIpsUttag: 0, tempTjpUttag: 0 };
    for (const savingsKind of ["isk", "depot"] as const) {
      const here = computeStudy({ ...base, ...savings, savingsKind }, shared, deaths);
      for (const option of here.options) {
        const there = runPayouts(
          normaliseLasts({ ...defaultLastsInput(), ...savings, savingsKind, retirementAge: option.par }),
          shared,
          deaths,
        );
        const lifelong = there.runs.find((r) => r.payout === LIFELONG)!;
        expect(option.individual, `${savingsKind} ${option.par}: individual`).toEqual(lifelong.individual);
        expect(option.savings, `${savingsKind} ${option.par}: savings`).toEqual(lifelong.savings);
      }
    }
  });

  it("differ between an ISK and a depot, which are taxed differently while they are saved", () => {
    const isk = withSavings.options[3]!;
    const depot = computeStudy({ ...savingsOnly, savingsKind: "depot" }, context, deaths).options[3]!;
    expect(depot.savings[depot.par]).toBeGreaterThan(0);
    expect(depot.savings[depot.par]).not.toBe(isk.savings[isk.par]);
  });

  it("count the private savings in Table 1's total, and the pension of their own in its gross", () => {
    const at = 3;
    const without = plain.options[at]!;
    const saved = withSavings.options[at]!;
    expect(saved.table1SavingsMonthly).toBeGreaterThan(0);
    expect(saved.table1IndividualMonthly).toBe(0);
    expect(saved.table1Monthly).toBeCloseTo(without.table1Monthly + saved.table1SavingsMonthly, 6);
    expect(saved.table1Share).toBeGreaterThan(without.table1Share);

    const own = withPension.options[at]!;
    expect(own.table1IndividualMonthly).toBeGreaterThan(0);
    expect(own.table1SavingsMonthly).toBe(0);
    expect(own.table1Monthly).toBeCloseTo(without.table1Monthly + own.table1IndividualMonthly, 6);
  });

  it("count the private savings in the after-tax Table 1 total too", () => {
    const net = computeStudy({ ...base, afterTax: true }, context, deaths).options[3]!;
    const saved = computeStudy({ ...savingsOnly, afterTax: true }, context, deaths).options[3]!;
    expect(saved.table1Monthly).toBeCloseTo(net.table1Monthly + saved.table1SavingsMonthly, 6);
    expect(saved.table1Share).toBeGreaterThan(net.table1Share);
  });

  it("are told apart from the agreement's pension by what the first year is made of", () => {
    expect(kindsPaid(plain)).toEqual({ occupational: false, individual: false, savings: false });
    expect(kindsPaid(withPension)).toEqual({ occupational: false, individual: true, savings: false });
    expect(kindsPaid(withSavings)).toEqual({ occupational: false, individual: false, savings: true });
    const everything = computeStudy({ ...everyKind, scheme: Scheme.Itp1 }, context, deaths);
    expect(kindsPaid(everything)).toEqual({ occupational: true, individual: true, savings: true });
    const agreementOff = computeStudy({ ...pensionOnly, scheme: Scheme.Itp1, includeOccupational: false }, context, deaths);
    expect(kindsPaid(agreementOff)).toEqual({ occupational: false, individual: true, savings: false });
  });

  it("leave the final salary alone", () => {
    expect(finalSalaryOf(everyKind, 67, context, deaths)).toBe(finalSalaryOf(base, 67, context, deaths));
  });

  it("are worked out once per age however many agreements the same person is run under", () => {
    const known = new Map();
    const first = computeStudy({ ...savingsOnly, scheme: Scheme.None }, context, deaths, known);
    expect(known.size).toBe(first.options.length);
    const kept = [...known.values()];
    const second = computeStudy({ ...savingsOnly, scheme: Scheme.Itp1 }, context, deaths, known);
    expect(known.size).toBe(first.options.length);
    [...known.values()].forEach((payout, i) => expect(payout).toBe(kept[i]));
    // ... and an agreement makes no difference to them.
    second.options.forEach((option, i) => expect(option.savings).toEqual(first.options[i]!.savings));
    // Nothing is kept when there is nothing to keep.
    const none = new Map();
    computeStudy(base, context, deaths, none);
    expect(none.size).toBe(0);
  });

  it("are part of what a comparison of agreements is cached for, and of every agreement's total", () => {
    expect(comparisonKey({ ...base, savingsCapital: 1_000 })).not.toBe(comparisonKey(base));
    expect(comparisonKey({ ...base, individualMonthly: 1_000 })).not.toBe(comparisonKey(base));
    expect(comparisonKey({ ...base, savingsKind: "depot" })).not.toBe(comparisonKey(base));

    const schemes = [Scheme.Itp1, Scheme.SafLo];
    const own = computeComparison(everyKind, schemes, context, deaths);
    const bare = computeComparison(base, schemes, context, deaths);
    const rowsOwn = comparisonRows(own, 85, 67);
    const rowsBare = comparisonRows(bare, 85, 67);
    rowsOwn.forEach((row, i) => {
      // Everyone's savings pay out the same, so each total is higher by the same amount ...
      expect(row.total - rowsBare[i]!.total).toBeCloseTo(rowsOwn[0]!.total - rowsBare[0]!.total, 6);
      // ... and what an agreement adds is what it was.
      expect(row.extra).toBeCloseTo(rowsBare[i]!.extra, 6);
    });
    expect(rowsOwn[0]!.total).toBeGreaterThan(rowsBare[0]!.total);
    own.entries.forEach((entry) => {
      expect(entry.study.options.map((o) => o.savings)).toEqual(own.entries[0]!.study.options.map((o) => o.savings));
    });
  });

  it("are not switched off by leaving the agreement out", () => {
    const off = computeStudy({ ...pensionOnly, scheme: Scheme.Itp1, includeOccupational: false }, context, deaths);
    expect(off.rows.map((r) => r.totals)).toEqual(withPension.rows.map((r) => r.totals));
    expect(off.options[0]!.individual[off.options[0]!.par]).toBeGreaterThan(0);
  });

  it("are counted from the model's start year, as Avancerat sets it", () => {
    const later = computeStudy(pensionOnly, { ...context, ipsStart: 2030 }, deaths);
    // Four years of the capital's return, and of the saving, are missing.
    expect(later.options[3]!.individual[later.options[3]!.par]).toBeLessThan(withPension.options[3]!.individual[withPension.options[3]!.par]!);
  });
});

describe("after tax", () => {
  const net = computeStudy({ ...base, afterTax: true }, context, deaths);

  it("sums the model's net income from retirement, so every total is lower", () => {
    net.rows.forEach((row, i) => {
      row.totals.forEach((total, j) => {
        expect(total).toBeLessThan(study.rows[i]!.totals[j]!);
        expect(total).toBeGreaterThan(study.rows[i]!.totals[j]! * 0.6);
      });
    });
  });

  it("keeps the earliest start ahead at 80 and the latest ahead at 95", () => {
    expect(net.options[net.rows[0]!.best]!.par).toBe(65);
    expect(net.options[net.rows[15]!.best]!.par).toBe(71);
  });
});

describe("break-even", () => {
  const option = (par: number, annual: number): StudyOption => {
    const flow = Array.from({ length: LAST_AGE + 1 }, (_, age) => (age >= par ? annual : 0));
    const cumulative: number[] = [];
    let sum = 0;
    for (let age = 0; age <= LAST_AGE; age += 1) {
      cumulative.push(sum);
      sum += flow[age]!;
    }
    return {
      offset: 0,
      slot: "z",
      par,
      annual: flow,
      occupational: flow.map(() => 0),
      individual: flow.map(() => 0),
      savings: flow.map(() => 0),
      cumulative,
      firstYearMonthly: annual / 12,
      finalSalaryMonthly: 0,
      table1Monthly: 0,
      table1Share: 0,
      table1OccupationalMonthly: 0,
      table1IndividualMonthly: 0,
      table1SavingsMonthly: 0,
      warnings: [],
    };
  };

  it("is linear inside a year", () => {
    // 100 a year from 65 against 150 a year from 66: behind by 100, gaining 50 a year -> 2 years later, age 68.
    expect(breakEven(option(65, 100), option(66, 150))).toBeCloseTo(68, 10);
    // 100 against 130: behind by 100, gaining 30 -> 3.333 years after 66.
    expect(breakEven(option(65, 100), option(66, 130))).toBeCloseTo(66 + 100 / 30, 10);
  });

  it("is null when the later start never catches up", () => {
    expect(breakEven(option(65, 100), option(66, 100))).toBeNull();
    expect(breakEven(option(65, 100), option(66, 101))).toBeNull();
  });
});

describe("final salary as the input", () => {
  it("finds the monthly salary that gives a stated final salary at the reference age", () => {
    for (const target of [30_000, 40_000, 55_000, 80_000]) {
      const salary = monthlySalaryForTarget(base, target, context, deaths);
      expect(Math.abs(finalSalaryOf({ ...base, monthlySalary: salary }, 68, context, deaths) - target)).toBeLessThan(1);
      // A final salary is a few percent above today's salary under the workbook's own wage path.
      expect(salary).toBeLessThan(target);
      expect(salary).toBeGreaterThan(target * 0.9);
    }
    expect(monthlySalaryForTarget(base, 40_000, context, deaths)).toBe(38_623);
  });

  it("matches at the reference age, wherever the reference sits", () => {
    const later = { ...base, centre: 72 };
    const salary = monthlySalaryForTarget(later, 40_000, context, deaths);
    expect(Math.abs(finalSalaryOf({ ...later, monthlySalary: salary }, 72, context, deaths) - 40_000)).toBeLessThan(1);
  });

  it("is exact for a linear relation and gives up on a dead one", () => {
    expect(monthlySalaryForFinalSalary(31_000, (s) => s * 1.24)).toBe(25_000);
    expect(monthlySalaryForFinalSalary(0, (s) => s)).toBe(0);
    expect(monthlySalaryForFinalSalary(1000, () => 0)).toBe(1000);
  });
});

describe("the best age by lifespan", () => {
  it("has one column per lifespan: the winning age, what it pays a month on average and in all", () => {
    const rows = optimalRows(study);
    expect(rows).toHaveLength(16);
    expect(rows.map((r) => r.lifespan)).toEqual(study.rows.map((r) => r.lifespan));
    study.rows.forEach((r, i) => {
      const best = study.options[r.best]!;
      expect(rows[i]!.par).toBe(best.par);
      expect(rows[i]!.total).toBe(r.totals[r.best]);
      expect(rows[i]!.monthly! * (r.lifespan - best.par) * 12).toBeCloseTo(r.totals[r.best]!, 4);
    });
    expect(rows[0]).toMatchObject({ lifespan: 80, par: 65, total: 2977236 });
    expect(rows[0]!.monthly).toBeCloseTo(2977236 / 180, 6);
  });

  it("downloads as the table is laid out: a row per measure, whole kronor, millions with two decimals", () => {
    const sv = optimalToCsv(study, "sv").split("\r\n");
    expect(sv).toHaveLength(4);
    expect(sv[0]).toBe("Livslängd;80;81;82;83;84;85;86;87;88;89;90;91;92;93;94;95");
    expect(sv[1]!.startsWith("Pensionsålder;65;65;65;")).toBe(true);
    expect(sv[2]!.startsWith("Månadsbelopp kr;16540;")).toBe(true);
    expect(sv[3]!.startsWith("Livsutbetalning mkr;2,98;")).toBe(true);
    const en = optimalToCsv(study, "en").split("\r\n");
    expect(en[3]!.startsWith("Lifetime payout MSEK,2.98,")).toBe(true);
    const sheet = optimalToXlsxRows(study, "sv");
    expect(sheet).toHaveLength(4);
    expect(sheet.every((line) => line.length === 17)).toBe(true);
  });
});

describe("the exports", () => {
  it("write the lifespan matrix with whole kronor, a delimiter per language and the winner", () => {
    const sv = matrixToCsv(study, "sv").split("\r\n");
    expect(sv).toHaveLength(17);
    expect(sv[0]).toBe("Lever till (år);65;66;67;68;69;70;71;Bäst;Försprång mot näst bäst (kr)");
    expect(sv[1]).toBe("80;2977236;2947200;2906316;2853552;2765892;2673120;2573640;65;30036");
    const en = matrixToCsv(study, "en").split("\r\n");
    expect(en[0]).toBe("Lifespan (years),65,66,67,68,69,70,71,Best,Lead over runner-up (SEK)");
    expect(en[16]!.split(",")[8]).toBe("71");
  });

  it("write one line per age, blank before an alternative starts", () => {
    const lines = seriesToCsv(study, "sv", "monthly").split("\r\n");
    // 65..95 inclusive, plus the header.
    expect(lines).toHaveLength(1 + 31);
    const header = lines[0]!.split(";");
    expect(header).toHaveLength(1 + 7);
    expect(header[0]).toBe("Ålder");
    expect(header[1]).toBe("Pension per månad 65");
    const first = lines[1]!.split(";");
    expect(first[0]).toBe("65");
    expect(first[1]).toBe("16858");
    expect(first.slice(2, 8).every((cell) => cell === "")).toBe(true);
  });

  it("hold either the monthly pension or the total paid out, never both", () => {
    const total = seriesToCsv(study, "sv", "total").split("\r\n");
    expect(total).toHaveLength(1 + 31);
    const header = total[0]!.split(";");
    expect(header).toHaveLength(1 + 7);
    expect(header[0]).toBe("Före ålder");
    expect(header[1]).toBe("Total utbetalning 65");
    // Nothing is paid out before the first age; one year later it is that year's pension.
    const first = study.options[0]!;
    expect(total[1]!.split(";")[1]).toBe("0");
    expect(total[2]!.split(";")[1]).toBe(String(Math.round(first.cumulative[66] ?? 0)));
    expect(Number(total[2]!.split(";")[1])).toBeGreaterThan(0);
    const sheet = seriesToXlsxRows(study, "en", "total");
    expect(sheet).toHaveLength(1 + 31);
    expect(sheet[0]).toHaveLength(1 + 7);
    expect(sheet[0]![0]).toBe("Before age");
  });

  it("keep the lifespan range inside what the study supports", () => {
    expect(LIFESPAN_BOUNDS.max).toBe(LAST_AGE);
    expect(LIFESPAN_BOUNDS.min).toBeLessThan(base.lifespanFrom);
  });
});

// ------------------------------------------------------------------ agreements ---

describe("which agreements fit which birth years", () => {
  it("follows the cohort each label names", () => {
    const fits = (born: number) => SCHEME_ORDER.filter((scheme) => schemeFitsCohort(scheme, born));
    // ITP-1 from 1979, ITP-2 before it; PA16 avd 2 before 1988, avd 1 from it; AKAP-KR from 1986.
    expect(fits(1960)).toEqual([Scheme.None, Scheme.Itp2, Scheme.SafLo, Scheme.KapKl, Scheme.Pa16Avd2]);
    expect(fits(1978)).toEqual([Scheme.None, Scheme.Itp2, Scheme.SafLo, Scheme.KapKl, Scheme.Pa16Avd2]);
    expect(fits(1979)).toEqual([Scheme.None, Scheme.Itp1, Scheme.SafLo, Scheme.KapKl, Scheme.Pa16Avd2]);
    expect(fits(1980)).toEqual([Scheme.None, Scheme.Itp1, Scheme.SafLo, Scheme.KapKl, Scheme.Pa16Avd2]);
    expect(fits(1986)).toEqual([Scheme.None, Scheme.Itp1, Scheme.SafLo, Scheme.KapKl, Scheme.AkapKr, Scheme.Pa16Avd2]);
    expect(fits(1987)).toEqual([Scheme.None, Scheme.Itp1, Scheme.SafLo, Scheme.KapKl, Scheme.AkapKr, Scheme.Pa16Avd2]);
    expect(fits(1988)).toEqual([Scheme.None, Scheme.Itp1, Scheme.SafLo, Scheme.KapKl, Scheme.AkapKr, Scheme.Pa16Avd1]);
    expect(fits(2000)).toEqual([Scheme.None, Scheme.Itp1, Scheme.SafLo, Scheme.KapKl, Scheme.AkapKr, Scheme.Pa16Avd1]);
  });

  it("agrees with the workbook's own labels, so the table cannot drift from them", () => {
    expect(SCHEME_CHOICES.map((c) => c.value)).toEqual([...SCHEME_ORDER]);
    for (const choice of SCHEME_CHOICES) {
      const years = choice.label.match(/\b(?:19|20)\d{2}\b/g);
      const range = SCHEME_COHORTS[choice.value as (typeof SCHEME_ORDER)[number]];
      // A label that names a year has a range, and one that does not has none.
      expect(range !== undefined, choice.label).toBe(years !== null);
      if (range !== undefined) {
        // "födda 1979-" / "födda 1986 -" / "födda 1988 och senare" name the first birth year;
        // "födda före 1979" / "före 1988" name the first one *excluded*.
        const boundary = range.from ?? range.to! + 1;
        expect(Number(years![0]), choice.label).toBe(boundary);
        expect(/före/.test(choice.label), choice.label).toBe(range.to !== undefined);
      }
    }
  });

  it("is worded in both languages, and only for agreements that have a cohort", () => {
    expect(cohortText(Scheme.Itp1, "sv")).toBe("födda 1979 och senare");
    expect(cohortText(Scheme.Itp1, "en")).toBe("born 1979 or later");
    expect(cohortText(Scheme.Itp2, "sv")).toBe("födda före 1979");
    expect(cohortText(Scheme.Pa16Avd2, "en")).toBe("born before 1988");
    expect(cohortText(Scheme.SafLo, "sv")).toBeUndefined();
    expect(cohortText(Scheme.None, "en")).toBeUndefined();
  });

  it("is named briefly, with who it is for", () => {
    expect(schemeName(Scheme.None, "sv")).toBe("Saknar tjänstepension");
    expect(schemeName(Scheme.None, "en")).toBe("No occupational pension");
    expect(schemeName(Scheme.Itp1, "sv")).toBe("ITP-1");
    expect(schemeName(Scheme.Pa16Avd1, "sv")).toBe("PA16 avd 1");
    expect(schemeName(Scheme.Pa16Avd1, "en")).toBe("PA16 section 1");
    expect(schemeSector(Scheme.SafLo, "sv")).toBe("Privatanställda arbetare");
    expect(schemeSector(Scheme.KapKl, "en")).toBe("Municipal and regional employees");
    for (const scheme of SCHEME_ORDER) {
      expect(schemeName(scheme, "sv").length).toBeGreaterThan(0);
      expect(schemeSector(scheme, "en").length).toBeGreaterThan(0);
    }
  });

  it("picks, until someone picks, the agreements the cohort can belong to and the one chosen", () => {
    expect(defaultSchemes(1980, Scheme.None)).toEqual([Scheme.Itp1, Scheme.SafLo, Scheme.KapKl, Scheme.Pa16Avd2]);
    expect(defaultSchemes(1960, Scheme.None)).toEqual([Scheme.Itp2, Scheme.SafLo, Scheme.KapKl, Scheme.Pa16Avd2]);
    expect(defaultSchemes(1990, Scheme.None)).toEqual([
      Scheme.Itp1,
      Scheme.SafLo,
      Scheme.KapKl,
      Scheme.AkapKr,
      Scheme.Pa16Avd1,
    ]);
    // An agreement chosen in the inputs is shown even if it does not fit -- it is the one being asked about.
    expect(defaultSchemes(1980, Scheme.AkapKr)).toEqual([
      Scheme.Itp1,
      Scheme.SafLo,
      Scheme.KapKl,
      Scheme.AkapKr,
      Scheme.Pa16Avd2,
    ]);
    expect(AGREEMENTS).not.toContain(Scheme.None);
    expect(AGREEMENTS).toHaveLength(7);
  });
});

// ------------------------------------------------------------------ comparison ---

const everyAgreement = computeComparison(base, AGREEMENTS, context, deaths);
const best = (entry: { study: Study }) =>
  bestRuns(entry.study)
    .map((r) => `${r.from === r.to ? r.from : `${r.from}-${r.to}`}:${r.option.par}`)
    .join(" ");
const entryFor = (scheme: (typeof SCHEME_ORDER)[number]) => everyAgreement.entries.find((e) => e.scheme === scheme)!;

describe("a comparison of agreements", () => {
  it("runs the baseline first and then the chosen agreements in the workbook's order", () => {
    expect(everyAgreement.entries.map((e) => e.scheme)).toEqual([...SCHEME_ORDER]);
    expect(everyAgreement.entries.map((e) => e.fits)).toEqual([true, true, false, true, true, false, true, false]);
    const some = computeComparison(base, [Scheme.Pa16Avd2, Scheme.Itp1, Scheme.Pa16Avd2], context, deaths);
    expect(some.entries.map((e) => e.scheme)).toEqual([Scheme.None, Scheme.Itp1, Scheme.Pa16Avd2]);
    // The baseline is always there, whether or not it is asked for.
    expect(computeComparison(base, [], context, deaths).entries.map((e) => e.scheme)).toEqual([Scheme.None]);
    expect(computeComparison(base, [Scheme.None], context, deaths).entries).toHaveLength(1);
  });

  it("offers the same alternatives under every agreement, and the lifespans asked for", () => {
    expect(everyAgreement.ages).toEqual([65, 66, 67, 68, 69, 70, 71]);
    expect(everyAgreement.ages[everyAgreement.reference]).toBe(68);
    expect(everyAgreement.lifespans[0]).toBe(80);
    expect(everyAgreement.lifespans).toHaveLength(16);
    for (const input of [
      base,
      { ...base, born: 1960, centre: 66 },
      { ...base, born: 1990, centre: 68 },
      { ...base, centre: 66, lifespanFrom: 84, lifespanTo: 88 },
    ]) {
      const comparison = computeComparison(input, AGREEMENTS, context, deaths);
      for (const entry of comparison.entries) {
        expect(entry.study.options.map((o) => o.par)).toEqual([...comparison.ages]);
      }
    }
  });

  it("is, for the baseline, exactly the single study with no agreement", () => {
    const baseline = everyAgreement.entries[0]!;
    expect(baseline.study.rows.map((r) => r.totals)).toEqual(study.rows.map((r) => r.totals));
    expect(best(baseline)).toBe("80-82:65 83:66 84-86:68 87-95:71");
  });

  it("is, for an agreement, the single study with that agreement -- whatever the occupational switch says", () => {
    const single = computeStudy({ ...base, scheme: Scheme.Itp1 }, context, deaths);
    expect(entryFor(Scheme.Itp1).study.rows.map((r) => r.totals)).toEqual(single.rows.map((r) => r.totals));
    // The switch belongs to the single view; a comparison is about the agreements themselves.
    const off = computeComparison({ ...base, includeOccupational: false }, [Scheme.Itp1], context, deaths);
    expect(off.entries[1]!.study.rows.map((r) => r.totals)).toEqual(single.rows.map((r) => r.totals));
  });

  it("finds each agreement's best age per lifespan", () => {
    expect(best(entryFor(Scheme.Itp1))).toBe("80-82:65 83-84:66 85:67 86-87:68 88-95:71");
    expect(best(entryFor(Scheme.Itp2))).toBe("80-81:65 82:67 83-84:68 85:69 86-88:70 89-95:71");
    expect(best(entryFor(Scheme.SafLo))).toBe("80-83:65 84:66 85-87:68 88-95:71");
    expect(best(entryFor(Scheme.KapKl))).toBe("80-82:65 83:66 84-85:67 86-87:68 88-95:71");
    expect(best(entryFor(Scheme.AkapKr))).toBe("80-82:65 83:66 84-85:68 86-95:71");
    expect(best(entryFor(Scheme.Pa16Avd2))).toBe("80-82:65 83:66 84-85:68 86-87:69 88-95:71");
    expect(best(entryFor(Scheme.Pa16Avd1))).toBe("80-82:65 83:66 84-85:67 86-87:68 88-95:71");
  });

  it("totals what each agreement pays by 80 if retiring at 65, to the krona", () => {
    const at80 = (scheme: (typeof SCHEME_ORDER)[number]) => Math.round(totalFor(entryFor(scheme), 80, 65));
    expect(at80(Scheme.None)).toBe(2_977_236);
    expect(at80(Scheme.Itp1)).toBe(4_242_816);
    expect(at80(Scheme.Itp2)).toBe(4_185_576);
    expect(at80(Scheme.SafLo)).toBe(4_211_496);
    expect(at80(Scheme.KapKl)).toBe(4_226_436);
    expect(at80(Scheme.AkapKr)).toBe(4_357_116);
    expect(at80(Scheme.Pa16Avd2)).toBe(4_310_136);
    expect(at80(Scheme.Pa16Avd1)).toBe(4_679_676);
  });

  it("is a lifespan row of the study that was computed", () => {
    const options = study.options;
    expect(lifespanRow(options, 85)).toEqual(study.rows[5]);
    expect(lifespanRow(options, 80)).toEqual(study.rows[0]);
  });

  it("reuses studies it was handed, and only computes the rest", () => {
    const cache = new Map<(typeof SCHEME_ORDER)[number], Study>();
    const first = computeComparison(base, [Scheme.Itp1, Scheme.SafLo], context, deaths, cache);
    expect([...cache.keys()].sort()).toEqual([Scheme.None, Scheme.Itp1, Scheme.SafLo]);
    const again = computeComparison(base, [Scheme.Itp1, Scheme.SafLo, Scheme.KapKl], context, deaths, cache);
    expect(cache.size).toBe(4);
    // The studies it already had are the very same objects, not recomputed copies.
    expect(again.entries[0]!.study).toBe(first.entries[0]!.study);
    expect(again.entries[1]!.study).toBe(first.entries[1]!.study);
    expect(again.entries[2]!.study).toBe(first.entries[2]!.study);
  });

  it("keys the cache on everything but the agreement", () => {
    const key = comparisonKey(base);
    expect(comparisonKey({ ...base, scheme: Scheme.Itp2 })).toBe(key);
    expect(comparisonKey({ ...base, includeOccupational: false })).toBe(key);
    for (const changed of [
      { monthlySalary: 41_000 },
      { born: 1979 },
      { startWorkAge: 26 },
      { yearlyInflation: 0.02 },
      { realGrowth: 0.018 },
      { realReturn: 0.02 },
      { centre: 67 },
      { lifespanFrom: 81 },
      { lifespanTo: 90 },
      { afterTax: true },
    ]) {
      expect(comparisonKey({ ...base, ...changed }), JSON.stringify(changed)).not.toBe(key);
    }
  });

  it("carries the distinct warnings of every run", () => {
    expect(everyAgreement.warnings).toEqual([]);
  });
});

describe("reading a comparison", () => {
  it("ranks the agreements at 85 by what each pays at its own best age", () => {
    const rows = comparisonRows(everyAgreement, 85, "best");
    const row = (scheme: (typeof SCHEME_ORDER)[number]) => rows.find((r) => r.entry.scheme === scheme)!;
    expect(rows.map((r) => r.entry.scheme)).toEqual([...SCHEME_ORDER]);
    expect(Math.round(row(Scheme.None).total)).toBe(3_990_588);
    expect(row(Scheme.None).extra).toBe(0);
    expect(row(Scheme.None).extraShare).toBe(0);
    expect(Math.round(row(Scheme.Itp1).total)).toBe(5_665_188);
    expect(Math.round(row(Scheme.Itp1).extra)).toBe(1_674_600);
    expect(row(Scheme.Itp1).extraShare).toBeCloseTo(0.4196, 4);
    // ITP-1's best age at 85 is 67, one year before the riktålder, so its figures are the 67's.
    expect(row(Scheme.Itp1).option.par).toBe(67);
    expect(Math.round(row(Scheme.Itp1).occupationalMonthly)).toBe(7815);
    expect(Math.round(row(Scheme.Itp1).monthly)).toBe(26_911);
    // PA16 avd 1 is the best of them, and at 67 rather than 68.
    expect(row(Scheme.Pa16Avd1).option.par).toBe(67);
    expect(Math.round(row(Scheme.Pa16Avd1).total)).toBe(6_276_900);
    const totals = rows.map((r) => r.total);
    expect(rows[totals.indexOf(Math.max(...totals))]!.entry.scheme).toBe(Scheme.Pa16Avd1);
    expect(rows[totals.indexOf(Math.min(...totals))]!.entry.scheme).toBe(Scheme.None);
    // At 85 the 66, 67 and 68 totals are within a percent of each other under every agreement.
    expect(rows.every((r) => r.tie)).toBe(true);
  });

  it("reads every agreement at one retirement age when asked to", () => {
    const rows = comparisonRows(everyAgreement, 85, 65);
    expect(rows.every((r) => r.option.par === 65)).toBe(true);
    expect(Math.round(rows[0]!.total)).toBe(3_935_496);
    expect(Math.round(rows[1]!.total)).toBe(5_622_936);
    expect(Math.round(rows[1]!.extra)).toBe(1_687_440);
    // The best age is still what it was, whichever age the figures are read at.
    expect(rows.map((r) => r.best.par)).toEqual([68, 67, 69, 68, 67, 68, 68, 67]);
    expect(Math.round(rows[1]!.monthly)).toBe(23_889);
    expect(Math.round(rows[1]!.occupationalMonthly)).toBe(7031);
  });

  it("falls back to the reference age for a retirement age the study does not offer", () => {
    const offered = comparisonRows(everyAgreement, 85, 68).map((r) => Math.round(r.total));
    expect(comparisonRows(everyAgreement, 85, 90).map((r) => Math.round(r.total))).toEqual(offered);
    expect(comparisonRows(everyAgreement, 85, 90).every((r) => r.option.par === 68)).toBe(true);
  });

  it("draws one line per agreement over the lifespans", () => {
    const totals = totalSeries(everyAgreement, "best", "total");
    expect(totals.map((s) => s.entry.scheme)).toEqual([...SCHEME_ORDER]);
    expect(totals.every((s) => s.points.length === 16)).toBe(true);
    expect(totals[1]!.points[0]).toEqual([80, totalFor(entryFor(Scheme.Itp1), 80, "best")]);
    expect(Math.round(totals[1]!.points[0]![1])).toBe(4_242_816);
    // Paid-out totals never fall as the lifespan grows.
    for (const series of totals) {
      series.points.forEach((point, i) => {
        if (i > 0) expect(point[1]).toBeGreaterThanOrEqual(series.points[i - 1]![1]);
      });
    }
  });

  it("draws what an agreement adds, with no line for the baseline", () => {
    const extra = totalSeries(everyAgreement, "best", "extra");
    expect(extra.map((s) => s.entry.scheme)).toEqual([...AGREEMENTS]);
    const itp1 = extra[0]!.points.find(([lifespan]) => lifespan === 85)!;
    expect(Math.round(itp1[1])).toBe(1_674_600);
    // The same number the table reads for it.
    expect(itp1[1]).toBe(comparisonRows(everyAgreement, 85, "best")[1]!.extra);
    // And at a fixed age it is the difference at that age.
    const fixed = totalSeries(everyAgreement, 65, "extra")[0]!.points.find(([lifespan]) => lifespan === 85)!;
    expect(Math.round(fixed[1])).toBe(1_687_440);
  });

  it("shows how far apart the agreements' best ages are at each lifespan", () => {
    const spread = bestAgeSpread(everyAgreement);
    expect(spread).toHaveLength(16);
    const at = (lifespan: number) => spread.find((s) => s.lifespan === lifespan)!;
    expect([at(80).min, at(80).max]).toEqual([65, 65]);
    expect([at(83).min, at(83).max]).toEqual([65, 68]);
    expect([at(85).min, at(85).max]).toEqual([67, 69]);
    expect([at(86).min, at(86).max]).toEqual([68, 71]);
    expect([at(95).min, at(95).max]).toEqual([71, 71]);
    expect(spread.filter((s) => s.min === s.max).map((s) => s.lifespan)).toEqual([80, 81, 89, 90, 91, 92, 93, 94, 95]);
  });

  it("reads after tax too, where an agreement still adds something", () => {
    const net = computeComparison({ ...base, afterTax: true }, [Scheme.Itp1], context, deaths);
    const gross = computeComparison(base, [Scheme.Itp1], context, deaths);
    const netRows = comparisonRows(net, 85, "best");
    const grossRows = comparisonRows(gross, 85, "best");
    expect(netRows[0]!.total).toBeLessThan(grossRows[0]!.total);
    expect(netRows[1]!.extra).toBeGreaterThan(0);
    expect(netRows[1]!.extra).toBeLessThan(grossRows[1]!.extra);
  });
});

describe("the comparison's exports", () => {
  it("write each agreement's best age and total, per lifespan", () => {
    const lines = comparisonToCsv(everyAgreement, "sv").split("\r\n");
    expect(lines).toHaveLength(17);
    const header = lines[0]!.split(";");
    expect(header).toHaveLength(1 + 2 * 8);
    expect(header.slice(0, 5)).toEqual([
      "Lever till (år)",
      "Saknar tjänstepension: bästa ålder",
      "Saknar tjänstepension: total före skatt (kr)",
      "ITP-1: bästa ålder",
      "ITP-1: total före skatt (kr)",
    ]);
    expect(lines[1]).toBe("80;65;2977236;65;4242816;65;4185576;65;4211496;65;4226436;65;4357116;65;4310136;65;4679676");
    expect(lines[16]!.split(";").filter((_, i) => i > 0 && i % 2 === 1).every((age) => age === "71")).toBe(true);
  });

  it("use the other language's wording and delimiter, and say when it is after tax", () => {
    const en = comparisonToCsv({ ...everyAgreement, input: { ...base, afterTax: true } }, "en").split("\r\n");
    expect(en[0]!.split(",").slice(0, 3)).toEqual([
      "Lifespan (years)",
      "No occupational pension: best age",
      "No occupational pension: total after tax (SEK)",
    ]);
    expect(en[1]!.startsWith("80,65,2977236,")).toBe(true);
  });

  it("write the same figures as spreadsheet rows", () => {
    const rows = comparisonToXlsxRows(everyAgreement, "sv");
    expect(rows).toHaveLength(17);
    expect(rows[0]).toHaveLength(17);
    expect(rows[1]!.slice(0, 5)).toEqual([80, 65, kronorCell(2_977_236), 65, kronorCell(4_242_816)]);
  });
});
