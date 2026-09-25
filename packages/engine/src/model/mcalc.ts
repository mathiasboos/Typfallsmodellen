/**
 * `Mcalc` -- the age loop that computes a typfall's whole life.
 *
 * Port of the main loop in VBA_go.bas, lines 899-2118. The VBA writes it as one
 * 1 200-line body; here it is one function per section comment in the original,
 * called in the original's order, so that next year's diff still lines up
 * against it.
 *
 * Everything the loop reads that the VBA takes from a global or a named range
 * arrives on `Run`, and everything it writes lands in `RunState`.
 */

import municipalTaxJson from "../../../data/municipal-tax.json" with { type: "json" };
import { withdrawalShare } from "../income/wages.js";
import type { WageProfile } from "../income/wages.js";
import { fnDeltalIp2, fnDeltalPp2 } from "../pension/deltal.js";
import type { DeltalTables } from "../pension/deltal.js";
import { gpavgift, ipavgift, pgi, ppavgift, riktage } from "../pension/contributions.js";
import { pgbBarn } from "../pension/incomePension.js";
import { arbgiv } from "../skatt/reduktioner.js";
import { premiumFor } from "../tjanstepension/index.js";
import type { SchemeContext } from "../tjanstepension/types.js";
import type { SchemeId } from "../tjanstepension/types.js";
import { vbaInt, vbaRound, wsMax } from "../vba/math.js";
import { rgkFor } from "./context.js";
import type { ModelContext } from "./context.js";
import type { TypfallInput } from "./input.js";
import { conscriptionDaysByYear, conscriptionPgb, studyPgb } from "./pgb.js";
import type { RunProfile, SetupResult, Warning } from "./setup.js";
import type { RunState, SetupVectors } from "./state.js";
import { drawdownPhase } from "./drawdown.js";
import { taxAndBenefits } from "./taxAndBenefits.js";

/** Everything a loop section needs, gathered from Mcalc's globals and ranges. */
export interface Run {
  readonly input: TypfallInput;
  readonly context: ModelContext;
  /** `startsetup`'s vectors. */
  readonly v: SetupVectors;
  /** Its resolved scalars. */
  readonly p: RunProfile;
  /** What the loop fills in. */
  readonly s: RunState;
  readonly deltalTables: DeltalTables;
  /** `pmonth` / `Tmonth` -- months of pension in the year it starts. */
  readonly pmonth: number;
  readonly tmonth: number;
  /** `forstid` -- försäkringstid at 65, after Mcalc's own correction. */
  readonly forstid: number;
  /** `kvoten` -- pbb/IBB at the boundary year, or 1 when there is none. */
  readonly kvoten: number;
  /** `Iyear` -- the year expenditure rules switch to earnings indexation. */
  readonly iyear: number;
  /** The PGB sheet's own entries, resolved to kronor and keyed by age --
   * `earnPgb` reads this the same way regardless of which of `sa`, `vpl` or
   * `studier` were typed directly and which `buildPgbManual` computed. */
  readonly pgbManual: ReadonlyMap<number, ResolvedPgbYear>;
}

/**
 * `earnPgb`'s own view of one age's PGB sheet: everything already in kronor
 * -- `vplDays` is the one exception, kept alongside `vpl` only because the
 * web UI shows it as its own column next to the kronor it produced, the same
 * way the sheet itself shows a day count and a PGB amount side by side.
 */
export interface ResolvedPgbYear {
  readonly sa: number;
  readonly vpl: number;
  readonly vplDays: number;
  readonly studier: number;
}

/**
 * Resolves `TypfallInput.pgbManual`/`pgbConscription` into one kronor figure
 * per age, the shape `earnPgb` has always read. `sa` is typed kronor, passed
 * through; `vpl` and `studier` are computed -- conscription from a single
 * date range split across the years it touches (`conscriptionDaysByYear`),
 * study from a per-age semester count -- so this is the one place either
 * touches a calendar year or a lookup table rather than a typed number.
 */
function buildPgbManual(
  input: TypfallInput,
  v: SetupVectors,
  born: number,
  marginal: number,
): ReadonlyMap<number, ResolvedPgbYear> {
  const byAge = new Map<number, { sa: number; studier: number }>(
    (input.pgbManual ?? []).map((row) => [
      row.age,
      { sa: row.sa, studier: studyPgb(vbaInt(born) + row.age, row.studySemesters, marginal) },
    ]),
  );

  const vplByAge = new Map<number, { vpl: number; vplDays: number }>();
  if (input.pgbConscription) {
    for (const [year, days] of conscriptionDaysByYear(input.pgbConscription)) {
      const age = year - vbaInt(born);
      vplByAge.set(age, {
        vpl: conscriptionPgb(year, days, v.mpgi.getOrZero(age), marginal),
        vplDays: days,
      });
    }
  }

  const ages = new Set([...byAge.keys(), ...vplByAge.keys()]);
  const result = new Map<number, ResolvedPgbYear>();
  for (const age of ages) {
    const row = byAge.get(age);
    const vplEntry = vplByAge.get(age);
    result.set(age, {
      sa: row?.sa ?? 0,
      vpl: vplEntry?.vpl ?? 0,
      vplDays: vplEntry?.vplDays ?? 0,
      studier: row?.studier ?? 0,
    });
  }
  return result;
}

const TAX = municipalTaxJson as {
  firstYear: number;
  series: Record<string, { values: (number | null)[] }>;
};

/** Employer contribution rates for a year, from K_skatt columns 10 and 11. */
function employerRates(year: number): { full: number; oldAge: number } {
  const read = (name: string) => {
    const values = TAX.series[name]!.values;
    const index = Math.min(Math.max(year - TAX.firstYear, 0), values.length - 1);
    return values[index] ?? 0;
  };
  return { full: read("arbetsgivaravgift"), oldAge: read("varavAlderspension") };
}

/**
 * `Utgyear` and `Skyear` -- the rule years this age is computed under.
 *
 * Both follow the income year unless the run pins expenditure or tax rules to
 * one year, either from a given year on or throughout.
 */
export function ruleYears(run: Run, age: number): { utgyear: number; skyear: number } {
  const year = run.v.year.get(age);
  const { rules, rulesFromUtg, rulesFromSkatt } = run.context;

  let utgyear = year;
  if (rulesFromUtg !== 0 && ((year > rulesFromUtg && rules === 0) || rules === 1)) {
    utgyear = rulesFromUtg;
  }

  let skyear = year;
  if (rulesFromSkatt !== 0 && ((year > rulesFromSkatt && rules === 0) || rules === 1)) {
    skyear = rulesFromSkatt;
  }

  return { utgyear, skyear };
}

/** Pension-qualifying income, the employer's contribution and the individual's own. */
export function earnPgi(run: Run, age: number, utgyear: number, skyear: number): void {
  const { v, s, context } = run;

  s.pgi.set(age, pgi(utgyear, v.income.get(age), v.pbb.get(age), v.ibb.get(age), v.fpb.get(age), context, 0, age, 0));

  // Only ever written out "för studier av arbetsgivarna"; nothing reads it.
  s.arbAvg.set(
    age,
    arbgiv(v.wage.get(age), skyear, age, employerRates(skyear), run.p.riktalder, context.marginal, 0),
  );
  s.egenAvg.set(
    age,
    pgi(utgyear, v.income.get(age), v.pbb.get(age), v.ibb.get(age), v.fpb.get(age), context, 1, age, 0),
  );
}

/** The contribution ceiling has been 7.5 base amounts since 1960. */
const TAK = 7.5;

/**
 * Pensionsgrundande belopp: sickness compensation, childcare years,
 * conscription and study.
 *
 * Sickness/activity compensation is still typed kronor. Childcare years,
 * conscription and study are all computed -- `buildPgbManual` resolves the
 * latter two from a date range and a semester count before the loop starts,
 * so `manual?.vpl`/`manual?.studier` below are already final kronor. The
 * shipped workbook has none of the three, so childcare years alone move a
 * default run.
 */
export function earnPgb(run: Run, age: number, utgyear: number): void {
  const { v, s, p, context } = run;
  const year = v.year.get(age);
  const manual = run.pgbManual.get(age);

  if (!(age < p.par && vbaInt(p.born) > 1937)) return;

  s.uttagIp = 0;
  s.uttagPp = 0;
  let diverse = 0;

  if (!(age > 15 && age <= p.riktalder)) return;

  const capped = () => wsMax(TAK * v.ibb.get(age) - s.pgi.get(age), 0);

  s.pgb.set(age, manual?.sa ?? 0);
  if (s.pgi.get(age) + s.pgb.get(age) > TAK * v.ibb.get(age)) s.pgb.set(age, capped());

  // Childcare years, for up to four children.
  const [barn1, barn2, barn3, barn4] = context.childBirthYears;

  // NOTE: the first child's age offset has no `- 1` and its income base is
  // uprated by a year of CPI, where the other three have neither. Both
  // asymmetries are in the original.
  let counter = barn1 - vbaInt(p.born);
  if (barn1 > 1960 && counter > 15) {
    diverse = pgbBarn(
      year,
      v.income.getOrZero(counter),
      (s.pgi.get(age) * v.kpi.getOrZero(counter)) / v.kpi.getOrZero(counter - 1) + s.pgb.get(age),
      barn1,
      age,
      v.mpgi.get(age),
      v.ibb.get(age),
      context.marginal,
      p.riktalder,
    );
  }

  for (const barn of [barn2, barn3, barn4]) {
    counter = barn - vbaInt(p.born) - 1;
    if (barn > 1960 && diverse === 0 && counter > 15) {
      diverse = pgbBarn(
        year,
        v.income.getOrZero(counter),
        s.pgi.get(age) + s.pgb.get(age),
        barn,
        age,
        v.mpgi.get(age),
        v.ibb.get(age),
        context.marginal,
        p.riktalder,
      );
    }
  }

  s.pgbBarn.set(age, diverse);
  s.pgb.set(age, s.pgb.get(age) + diverse);

  // Conscription -- `buildPgbManual` has already applied wsPGB!I's own
  // round-down-to-100 (or not, per `marginal`), the same as `study` below.
  s.pgb.set(age, s.pgb.get(age) + (manual?.vpl ?? 0));
  if (s.pgi.get(age) + s.pgb.get(age) > TAK * v.ibb.get(age)) s.pgb.set(age, capped());

  // Study -- `buildPgbManual` has already turned the typed semester count
  // into kronor and applied wsPGB!P's own rounding.
  s.pgb.set(age, s.pgb.get(age) + (manual?.studier ?? 0));
  if (s.pgi.get(age) + s.pgb.get(age) > TAK * v.ibb.get(age)) s.pgb.set(age, capped());

  if (s.pgb.get(age) + s.pgi.get(age) > 0 && age < 71) s.pgbYears += 1;
}

/**
 * The year's pension rights: ATP points, and the three public contributions.
 *
 * The contributions are levied on the *previous* year's income, because a
 * pension-qualifying income is only settled by the following year's assessment.
 */
export function earnRights(run: Run, age: number): void {
  const { v, s, p, context } = run;
  const year = v.year.get(age);

  if (age === v.startage) {
    s.ipRatt.set(age, 0);
    if (age < 65 && year > 1959) {
      // NOTE: no PGB term here, where every later age adds one. At the model's
      // own start age there is none to add, so the two agree in practice.
      let points = wsMax(s.pgi.get(age) / v.fpb.get(age) - 1, 0);
      if (context.marginal === 0) points = vbaRound(points, 2);
      s.tpPoints.set(age, points);
      s.stpPoints.set(age, points);
    }
    s.gpRatt.set(age, 0);
    s.ppRatt.set(age, 0);
    return;
  }

  // ATP points accrue to the 64th birthday.
  if (age < 65) {
    if (year > 1959) {
      let points = wsMax((s.pgi.get(age) + s.pgb.get(age)) / v.fpb.get(age) - 1, 0);
      if (context.marginal === 0) points = vbaRound(points, 2);
      s.tpPoints.set(age, points);
      // The copy FTJP reads after Mcalc sorts `TP_points` in place.
      s.stpPoints.set(age, points);
    } else {
      s.tpPoints.set(age, 0);
      s.stpPoints.set(age, 0);
    }
  }

  const lastYear = year - 1;
  const lastAge = age - 1;
  const lastPgi = s.pgi.getOrZero(lastAge);
  const lastPgb = s.pgb.getOrZero(lastAge);
  const { andelnya, riktalder } = p;
  const born = vbaInt(p.born);
  const marginal = context.marginal;

  if (s.pgbYears >= 5) {
    s.ipRatt.set(age, ipavgift(lastYear, lastPgi + lastPgb, lastAge, andelnya, marginal, born));
    s.ppRatt.set(age, ppavgift(lastYear, lastPgi + lastPgb, lastAge, andelnya, marginal, born));
    s.gpRatt.set(age, gpavgift(lastYear, lastPgi + lastPgb, lastAge, andelnya, marginal, born, riktalder));
  } else if (s.pgbYears > 0) {
    // Without five qualifying years the whole PGB goes to the income pension at
    // 18.5%, so it is scaled up by 185/160 before the 16% rate is applied.
    s.ipRatt.set(
      age,
      ipavgift(lastYear, lastPgi + vbaInt((lastPgb * 185) / 160), lastAge, andelnya, marginal, born),
    );
    s.ppRatt.set(age, ppavgift(lastYear, lastPgi, lastAge, andelnya, marginal, born));
    s.gpRatt.set(
      age,
      gpavgift(lastYear, lastPgi + (lastPgb * 185) / 160, lastAge, andelnya, marginal, born, riktalder),
    );
  } else {
    s.ipRatt.set(age, ipavgift(lastYear, lastPgi, lastAge, andelnya, marginal, born));
    s.ppRatt.set(age, ppavgift(lastYear, lastPgi, lastAge, andelnya, marginal, born));
    s.gpRatt.set(age, gpavgift(lastYear, lastPgi, lastAge, andelnya, marginal, born, riktalder));
  }

  // Years with either a pension right or an ATP point, for inkomstpensionstillägg.
  if (age <= riktalder) {
    if (s.ipRatt.get(age) > 0 || s.tpPoints.getOrZero(age - 1) > 0) s.pgiYears += 1;
  }
}

/** This year's occupational premium, whichever agreement applies. */
export function earnOccupational(run: Run, age: number): void {
  const { v, s, p, context } = run;
  const year = v.year.get(age);

  const schemeContext: SchemeContext = {
    year,
    born: p.born,
    wStart: p.wStart,
    tjpPar: p.tjpPar,
    flexPension: context.flexPension,
    marginal: context.marginal,
  };

  // NOTE: dead. `tlITP2A` returns 0 below 1997 on its own, so this guard can
  // never change an answer. Kept because the original has it, and asserted, so
  // that a future workbook moving either year shows up.
  if (p.avtal === 3 && year <= 1976) {
    s.tjpRatt.set(age, 0);
    return;
  }

  s.tjpRatt.set(
    age,
    premiumFor(
      p.avtal as SchemeId,
      age,
      v.wage.get(age),
      v.ibb.get(age),
      p.tjpPar,
      year,
      schemeContext,
      context,
    ),
  );
}

/** This year's private saving premium; nothing is put aside after retirement. */
export function earnPrivateSaving(run: Run, age: number): void {
  const { v, s, p, context } = run;

  s.ipsRatt = 0;
  s.ips.set(age, 0);
  s.pps.set(age, 0);

  if (context.ipsStart <= vbaInt(p.born) + age && age < p.tjpPar) {
    // Above 1 the setting is an amount per month; at or below it, a share of
    // income. No check of how many months the first year has.
    s.ipsRatt =
      context.ipsMonthly > 1
        ? context.ipsMonthly * 12
        : context.ipsMonthly * v.income.get(age);
    if (context.marginal !== 0) s.ipsRatt = vbaInt(s.ipsRatt + 0.5);
  }
}

/**
 * The share of each pension drawn this year, and the divisors it is drawn on.
 *
 * Nothing before the pension starts; the chosen shares between first and final
 * withdrawal; the whole pension from the final withdrawal on.
 */
export function resolveWithdrawal(run: Run, age: number): void {
  const { s, p, context, deltalTables } = run;
  const born = vbaInt(p.born);

  if (age < vbaInt(p.par)) {
    s.dtalIp = 0;
    s.dtalPp = 0;
    s.uttagIp = 0;
    s.uttagPp = 0;
    return;
  }

  if (age >= vbaInt(p.par) && p.par <= p.defAr && age < p.defAr) {
    s.uttagIp = context.uttagIp;
    s.uttagPp = context.uttagPp;
    // DEAD in the original too: reaching here needs age >= par and age < defAr
    // at once, which par === defAr forbids.
    if (p.par === p.defAr) {
      s.uttagIp = 1;
      s.uttagPp = 1;
    }
  } else if (age >= p.defAr) {
    s.uttagIp = 1;
    s.uttagPp = 1;
  }

  s.dtalIp = fnDeltalIp2(born, age, deltalTables, p.par, p.defAr);
  s.dtalPp = fnDeltalPp2(born, age, deltalTables, p.par, p.defAr);
}

/**
 * The scalars `Mcalc` itself resolves before the loop, and the `Run` they go on.
 *
 * `startsetup` has already validated the typfall; this is the rest of Mcalc's
 * own preamble -- months of pension in the first year, försäkringstid, the
 * boundary year for earnings indexation, and the riktålder under a pinned rule
 * year.
 */
export function prepareRun(
  setup: SetupResult,
  input: TypfallInput,
  contextIn: ModelContext,
  state: RunState,
  deltalTables: DeltalTables,
  warnings: Warning[] = [],
): Run {
  const { vectors: v, profile } = setup;
  const born = profile.born;

  // `Rng_riktage` is a cell on Nyckeltal, looked up by cohort, and two rules
  // read it: the age the higher grundavdrag starts at (`Xage`, via
  // Skatteregler.bas:26) and the LAS age that KAP-KL's premium steps down at
  // (Tjänstepensioner.bas:1335). Both take the context, so the cohort's value
  // goes on it here rather than each of them reaching for the cohort.
  const context: ModelContext = Object.freeze({
    ...contextIn,
    riktage: profile.riktalder,
    // The same sheet cell the projection reads, so `tjpkassa`'s divisor
    // adjustment sees the rate this run's assumptions produce rather than the
    // base. See `rgkFor`.
    rgk: rgkFor(contextIn.rgk, input.yearlyInflation, input.realGrowth),
  });

  // Pension starts on the first of the month the typfall turns PAR.
  const pmonth = 12 - vbaInt(12 * (born + profile.par - vbaInt(born + profile.par)));
  const tmonth = 12 - vbaInt(12 * (born + profile.tjpPar - vbaInt(born + profile.tjpPar)));

  // Försäkringstid cannot be shorter than the working life implies.
  let forstid = context.insuranceYears;
  if (forstid < 40 && vbaInt(profile.par - profile.wStart - 1) > forstid) {
    if (input.ownIncome === undefined) {
      const corrected = Math.min(profile.par - profile.wStart - 1, 40);
      warnings.push({
        field: "rng_Försäkringstid_vid_65",
        given: forstid,
        used: corrected,
        message:
          `Försäkringstid angiven under 40 år, men antal arbetande år är fler än angivet; ` +
          `försäkringstiden är ändrad till ${corrected} år`,
      });
      forstid = corrected;
    }
  }

  // The year expenditure rules switch from prices to earnings.
  let iyear = context.boundaryYear;
  if (iyear === 0) iyear = vbaInt(born) + profile.slutage;
  if (iyear < profile.modelYear && iyear > 0) {
    warnings.push({
      field: "rng_Boundray_Year",
      given: context.boundaryYear,
      used: vbaInt(born) + profile.slutage,
      message: `Inte möjligt att retroaktivt ändra reglerna, bortser från ${iyear}`,
    });
    iyear = vbaInt(born) + profile.slutage;
  }

  /**
   * NOTE: this opening `kvoten` is only ever 1. Its guard reads `If Iyear <
   * slutage`, comparing a calendar year against the constant 105, which no real
   * year satisfies -- so the pbb/IBB ratio is never taken here. The
   * housing-supplement block later in the loop reassigns `kvoten` on a guard
   * that *can* hold, which is why it lives on `RunState`.
   */
  let kvoten = 1;
  const iyearAge = iyear - vbaInt(born);
  if (iyear < profile.slutage && iyearAge > profile.startage && iyearAge < profile.slutage) {
    kvoten = v.pbb.get(iyearAge) / v.ibb.get(iyearAge);
  }

  // A pinned expenditure-rule year pins the riktålder with it.
  let { riktalder, riktl } = profile;
  if (context.rulesFromUtg !== 0) {
    const parYear = v.year.getOrZero(vbaInt(profile.par));
    if ((parYear > context.rulesFromUtg && context.rules === 0) || context.rules === 1) {
      riktalder = riktage(context.rulesFromUtg, 1);
      riktl = riktage(context.rulesFromUtg, 0);
    }
  }

  const pgbManual = buildPgbManual(input, v, born, context.marginal);

  state.kvoten = kvoten;

  return {
    input,
    context,
    v,
    p: { ...profile, riktalder, riktl },
    s: state,
    deltalTables,
    pmonth,
    tmonth,
    forstid,
    kvoten,
    iyear,
    pgbManual,
  };
}

/** The wage profile Mcalc passes to `withdrawalShare`. */
export function wageProfileOf(run: Run): WageProfile {
  return {
    born: run.p.born,
    par: run.p.par,
    defAr: run.p.defAr,
    wageProfile: run.context.wageProfile,
    workDuringPartialWithdrawal: run.context.workDuringPartialWithdrawal,
    partialWithdrawalShare: run.context.uttagIp,
  };
}

/**
 * One age's earning phase, in the VBA's order.
 *
 * Returns the rule years, which the drawdown and tax sections reuse.
 */
export function earningPhase(run: Run, age: number): { utgyear: number; skyear: number } {
  const years = ruleYears(run, age);
  earnPgi(run, age, years.utgyear, years.skyear);
  earnPgb(run, age, years.utgyear);
  earnRights(run, age);
  earnOccupational(run, age);
  earnPrivateSaving(run, age);
  resolveWithdrawal(run, age);
  return years;
}

/**
 * One age, earning phase then drawdown.
 *
 * The tax, benefit and output sections follow in `taxAndBenefits`.
 */
export function simulateYear(run: Run, age: number): { utgyear: number; skyear: number } {
  const years = earningPhase(run, age);
  drawdownPhase(run, age, years.utgyear);
  taxAndBenefits(run, age, years.utgyear, years.skyear);
  return years;
}

/** The whole loop, `startage` to `slutage`. */
export function runLoop(run: Run): void {
  for (let age = run.v.startage; age <= run.v.slutage; age += 1) simulateYear(run, age);
}

export { withdrawalShare };
