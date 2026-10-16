/**
 * "Bästa pensionsålder"'s arithmetic, checked without a browser.
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
  lifespanRow,
  matrixToCsv,
  monthlySalaryForFinalSalary,
  monthlySalaryForTarget,
  referenceAge,
  schemeFitsCohort,
  schemeName,
  schemeSector,
  seriesToCsv,
  slotOf,
  studyContext,
  studyTypfall,
  totalFor,
  totalSeries,
} from "../src/pensionAge.js";
import type { Study, StudyOption } from "../src/pensionAge.js";
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

  it("force occupational pension to follow retirement and leave private saving out", () => {
    const shared = { ...context, tjpPar: 62, ipsMonthly: 5000, pbhIps: 250_000 };
    const used = studyContext(shared);
    expect(used.tjpPar).toBe(0);
    expect(used.ipsMonthly).toBe(0);
    expect(used.pbhIps).toBe(0);
    // Everything else is the shared context untouched.
    expect(used.priceBasis).toBe(shared.priceBasis);
    expect(used.referenceYear).toBe(shared.referenceYear);
    expect(used.marginal).toBe(shared.marginal);
  });
});

describe("the default study reproduces the static report", () => {
  it("runs seven alternatives around the riktålder", () => {
    expect(parsOf(study)).toEqual([65, 66, 67, 68, 69, 70, 71]);
    expect(study.riktalder).toBe(68);
    expect(study.options[study.reference]!.par).toBe(68);
    expect(study.rows).toHaveLength(16);
    expect(study.rows[0]!.lifespan).toBe(80);
    expect(study.rows[15]!.lifespan).toBe(95);
    expect(study.warnings).toEqual([]);
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
    expect(study.rows[0]!.totals).toEqual([2_972_916, 2_942_724, 2_901_684, 2_851_104, 2_760_972, 2_668_008, 2_568_324]);
    expect(study.rows[0]!.margin).toBe(30_192);
    expect(study.rows[3]!.totals).toEqual([3_550_056, 3_552_168, 3_547_176, 3_536_796, 3_485_232, 3_435_720, 3_385_344]);
    expect(study.rows[3]!.margin).toBe(2_112);
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
    expect(at(65, 68)).toBeCloseTo(83.375, 2);
    expect(at(66, 68)).toBeCloseTo(83.619, 2);
    expect(at(67, 68)).toBeCloseTo(83.794, 2);
    expect(at(68, 69)).toBeCloseTo(87.184, 2);
    expect(at(68, 70)).toBeCloseTo(86.851, 2);
    expect(at(68, 71)).toBeCloseTo(86.6, 2);
    // The step from the riktålder to one year later is the slowest, not the fastest.
    expect(at(68, 69)!).toBeGreaterThan(at(69, 70)!);
    expect(study.versusReference).toHaveLength(6);
    expect(study.chain).toHaveLength(6);
  });

  it("reads the first year off the age matrix and keeps Table 1's own figure beside it", () => {
    const reference = study.options[study.reference]!;
    expect(reference.firstYearMonthly).toBe(20_194);
    // Table 1 counts the last year's pension right at once; the yearly flow picks it up in year two.
    expect(reference.table1Monthly).toBe(20_398);
    expect(reference.table1Monthly).toBeGreaterThan(reference.firstYearMonthly);
    expect(reference.finalSalaryMonthly).toBeCloseTo(39_872.3, 1);
    expect(reference.table1Share).toBeCloseTo(0.5116, 3);
  });
});

describe("the economic assumptions", () => {
  it("move the break-evens years earlier with real wage growth", () => {
    const growth = computeStudy({ ...base, yearlyInflation: 0.02, realGrowth: 0.018 }, context, deaths);
    expect(runsOf(growth)).toEqual(["80-80:67", "81-81:68", "82-82:70", "83-95:71"]);
    const first = growth.versusReference.find((b) => b.earlier.par === 65)!.at!;
    expect(first).toBeCloseTo(79.465, 2);
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
    expect(itp.rows[0]!.totals[0]).toBe(4_238_496);
  });

  it("can move the winning ages", () => {
    expect(runsOf(itp)).toEqual(["80-82:65", "83-84:66", "85-87:68", "88-95:71"]);
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
  it("is not allowed to start occupational pension early or add private saving", () => {
    const shared = { ...context, tjpPar: 62, ipsMonthly: 5000, pbhIps: 250_000 };
    const overridden = computeStudy({ ...base, scheme: Scheme.Itp1 }, shared, deaths);
    const plain = computeStudy({ ...base, scheme: Scheme.Itp1 }, context, deaths);
    expect(overridden.rows.map((r) => r.totals)).toEqual(plain.rows.map((r) => r.totals));
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
      cumulative,
      firstYearMonthly: annual / 12,
      finalSalaryMonthly: 0,
      table1Monthly: 0,
      table1Share: 0,
      table1OccupationalMonthly: 0,
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

describe("the exports", () => {
  it("write the lifespan matrix with whole kronor, a delimiter per language and the winner", () => {
    const sv = matrixToCsv(study, "sv").split("\r\n");
    expect(sv).toHaveLength(17);
    expect(sv[0]).toBe("Lever till (år);65;66;67;68;69;70;71;Bäst;Försprång mot näst bäst (kr)");
    expect(sv[1]).toBe("80;2972916;2942724;2901684;2851104;2760972;2668008;2568324;65;30192");
    const en = matrixToCsv(study, "en").split("\r\n");
    expect(en[0]).toBe("Lifespan (years),65,66,67,68,69,70,71,Best,Lead over runner-up (SEK)");
    expect(en[16]!.split(",")[8]).toBe("71");
  });

  it("write one line per age, blank before an alternative starts", () => {
    const lines = seriesToCsv(study, "sv").split("\r\n");
    // 65..95 inclusive, plus the header.
    expect(lines).toHaveLength(1 + 31);
    const header = lines[0]!.split(";");
    expect(header).toHaveLength(1 + 7 + 7);
    const first = lines[1]!.split(";");
    expect(first[0]).toBe("65");
    expect(first[1]).toBe("16498");
    expect(first.slice(2, 8).every((cell) => cell === "")).toBe(true);
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
    expect(best(entryFor(Scheme.Itp1))).toBe("80-82:65 83-84:66 85-87:68 88-95:71");
    expect(best(entryFor(Scheme.Itp2))).toBe("80-81:65 82:67 83-85:68 86-88:70 89-95:71");
    expect(best(entryFor(Scheme.SafLo))).toBe("80-83:65 84:66 85-87:68 88-95:71");
    expect(best(entryFor(Scheme.KapKl))).toBe("80-82:65 83:66 84:67 85-87:68 88-95:71");
    expect(best(entryFor(Scheme.AkapKr))).toBe("80-82:65 83:66 84-85:68 86-95:71");
    expect(best(entryFor(Scheme.Pa16Avd2))).toBe("80-82:65 83:66 84-85:68 86-87:69 88-95:71");
    expect(best(entryFor(Scheme.Pa16Avd1))).toBe("80-82:65 83:66 84-85:67 86-87:68 88-95:71");
  });

  it("totals what each agreement pays by 80 if retiring at 65, to the krona", () => {
    const at80 = (scheme: (typeof SCHEME_ORDER)[number]) => Math.round(totalFor(entryFor(scheme), 80, 65));
    expect(at80(Scheme.None)).toBe(2_972_916);
    expect(at80(Scheme.Itp1)).toBe(4_238_496);
    expect(at80(Scheme.Itp2)).toBe(4_181_256);
    expect(at80(Scheme.SafLo)).toBe(4_207_176);
    expect(at80(Scheme.KapKl)).toBe(4_222_116);
    expect(at80(Scheme.AkapKr)).toBe(4_352_796);
    expect(at80(Scheme.Pa16Avd2)).toBe(4_305_816);
    expect(at80(Scheme.Pa16Avd1)).toBe(4_675_356);
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
    expect(Math.round(row(Scheme.None).total)).toBe(3_988_140);
    expect(row(Scheme.None).extra).toBe(0);
    expect(row(Scheme.None).extraShare).toBe(0);
    expect(Math.round(row(Scheme.Itp1).total)).toBe(5_662_368);
    expect(Math.round(row(Scheme.Itp1).extra)).toBe(1_674_228);
    expect(row(Scheme.Itp1).extraShare).toBeCloseTo(0.4198, 4);
    expect(Math.round(row(Scheme.Itp1).occupationalMonthly)).toBe(8207);
    expect(Math.round(row(Scheme.Itp1).monthly)).toBe(28_401);
    // PA16 avd 1 is the best of them, and at 67 rather than 68.
    expect(row(Scheme.Pa16Avd1).option.par).toBe(67);
    expect(Math.round(row(Scheme.Pa16Avd1).total)).toBe(6_272_268);
    const totals = rows.map((r) => r.total);
    expect(rows[totals.indexOf(Math.max(...totals))]!.entry.scheme).toBe(Scheme.Pa16Avd1);
    expect(rows[totals.indexOf(Math.min(...totals))]!.entry.scheme).toBe(Scheme.None);
    // At 85 the 66, 67 and 68 totals are within a percent of each other under every agreement.
    expect(rows.every((r) => r.tie)).toBe(true);
  });

  it("reads every agreement at one retirement age when asked to", () => {
    const rows = comparisonRows(everyAgreement, 85, 65);
    expect(rows.every((r) => r.option.par === 65)).toBe(true);
    expect(Math.round(rows[0]!.total)).toBe(3_931_176);
    expect(Math.round(rows[1]!.total)).toBe(5_618_616);
    expect(Math.round(rows[1]!.extra)).toBe(1_687_440);
    // The best age is still what it was, whichever age the figures are read at.
    expect(rows.map((r) => r.best.par)).toEqual([68, 68, 68, 68, 68, 68, 68, 67]);
    expect(Math.round(rows[1]!.monthly)).toBe(23_529);
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
    expect(Math.round(totals[1]!.points[0]![1])).toBe(4_238_496);
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
    expect(Math.round(itp1[1])).toBe(1_674_228);
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
    expect([at(85).min, at(85).max]).toEqual([67, 68]);
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
    expect(lines[1]).toBe("80;65;2972916;65;4238496;65;4181256;65;4207176;65;4222116;65;4352796;65;4305816;65;4675356");
    expect(lines[16]!.split(";").filter((_, i) => i > 0 && i % 2 === 1).every((age) => age === "71")).toBe(true);
  });

  it("use the other language's wording and delimiter, and say when it is after tax", () => {
    const en = comparisonToCsv({ ...everyAgreement, input: { ...base, afterTax: true } }, "en").split("\r\n");
    expect(en[0]!.split(",").slice(0, 3)).toEqual([
      "Lifespan (years)",
      "No occupational pension: best age",
      "No occupational pension: total after tax (SEK)",
    ]);
    expect(en[1]!.startsWith("80,65,2972916,")).toBe(true);
  });

  it("write the same figures as spreadsheet rows", () => {
    const rows = comparisonToXlsxRows(everyAgreement, "sv");
    expect(rows).toHaveLength(17);
    expect(rows[0]).toHaveLength(17);
    expect(rows[1]!.slice(0, 5)).toEqual([80, 65, kronorCell(2_972_916), 65, kronorCell(4_238_496)]);
  });
});
