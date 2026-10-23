/**
 * Bästa pensionsålder: which retirement age pays out the most over a lifetime?
 *
 * There is no equivalent sheet in the workbook -- this is genuinely new, the
 * same class of addition as "Jämför scenarier" and Mikrosim. It answers one
 * question the single forecast cannot: the monthly pension rises with every
 * year of delay, but each year of delay also costs a year of payments, so
 * which age wins depends on how long the person lives. The model deliberately
 * does not know that (the annuity divisors are population averages), which is
 * what makes the answer a table of "if you live to D" rather than one number.
 *
 * Everything here is pure -- no DOM, no `document` -- so the arithmetic can be
 * unit-tested directly (`apps/web/test/pensionAge.test.ts`), the same split
 * `kpis.ts`/`computeKpis` and `compare.ts`/`applyScenario` already make. What
 * it reuses is `run()` itself, once per alternative retirement age. A run takes
 * a couple of milliseconds (the engine shares the annuity factors between runs
 * of one cohort -- `deltal.ts` says why that matters), so seven of them per
 * keystroke need no worker and no debounce, which is `main.ts`'s own stated
 * philosophy. The agreement comparison below is eight times that; it is only
 * computed while its view is showing, and its studies are cached by the caller.
 *
 * The method is the one a static report on the same typfall used first
 * (born 1980, work from 25, 38 500 kr/month; 7 ages around the riktålder):
 *
 * - **Alternatives.** The reference age and three years either side, clamped to
 *   the cohort's lowest permitted age (`lowestPensionAge`) and the sheet's
 *   highest (80). Colours follow the *offset* from the reference, not the
 *   position in the list, so an alternative keeps its colour when a neighbour
 *   is clamped away (`Slot`).
 * - **Pension** is the model's own per-age matrix (`result.rows`, the same
 *   numbers Figur 2 and Table 2 are drawn from): allmän pension (income,
 *   supplementary and premium pension, guarantee pension and IPT) plus
 *   occupational pension and the person's own pension saving (see below), before
 *   tax; or `netto`, after tax, from retirement. Whatever is paid out without
 *   tax (the private savings) is the same before and after. Amounts are in
 *   whatever price basis the shared context asks for (fixed reference-year
 *   prices by default) and are **not discounted**: a krona at 90 counts as much
 *   as one at 70.
 * - **"Lever till D år"** means pension is paid for every year from retirement
 *   up to the year before the D-year birthday -- `D - retirement age` whole
 *   years. The model's row for age `a` is the year in which the person turns
 *   `a` (an integer birth year means 1 January), a full year of payments, so
 *   the sum is `rows[par] + ... + rows[D - 1]`. The benefit does not depend on
 *   how long the person lives; the lifespan only decides how many years are
 *   added up. No probability weighting.
 * - **Break-even** between two alternatives is the lifespan at which the later
 *   start has caught up with the earlier one, interpolated linearly inside the
 *   year (payments are monthly, so the running total is piecewise linear).
 *
 * One setting of the shared (Avancerat) context is overridden, on purpose,
 * because it would break the question being asked:
 *
 * - `tjpPar: 0` -- occupational pension follows each alternative's own
 *   retirement age. An Avancerat "draw occupational pension from age X" would
 *   otherwise pay it before the public pension starts for some alternatives
 *   and not others, and `netto` (which still contains salary before
 *   retirement) could not be summed from that age.
 *
 * Every other shared setting (tax, wage profile, own salary path excluded --
 * that is a `TypfallInput` field this tab does not read -- price basis, the
 * "temporary occupational withdrawal" years, ...) applies as in the other tabs.
 *
 * **Own saving.** The person's pension saving beyond the public pension and the
 * agreement is asked for, and carried by the model, exactly as "Hur länge räcker
 * pensionen?" does (`privateSavings.ts` says what the kinds are and how the model
 * is made to carry them): an individual occupational pension and a private pension
 * -- one pot of taxed income, the model's IPS -- and private savings in an ISK or a
 * depot, paid out tax-free and run on their own. Capital counts from the year
 * Avancerat calls "Sparandet börjar år", and the monthly saving goes on until the
 * alternative's own retirement age -- so a later start also means more years of
 * saving and of return, which is part of what waiting is worth. They are paid out
 * over the shared "temporary withdrawal" years (lifelong unless Avancerat says
 * otherwise): an IPS for as long as the person lives, money until it is spent. The
 * private savings do not depend on the agreement, so a comparison of agreements
 * runs them once per age and not once per agreement.
 *
 * **Comparing agreements** (`computeComparison`). The same person, the same
 * seven alternatives, run once per collective agreement: "Saknar tjänstepension"
 * is the baseline, and everything an agreement adds is measured against it, so a
 * figure reads as "what this agreement is worth" and not as a level that mixes
 * the public pension in. The agreement is the *only* thing that varies -- not
 * the retirement age, which is shared by every agreement -- so the best age of
 * each agreement is a fair like-for-like answer. Some agreements are written
 * for particular birth years (`SCHEME_COHORTS`, taken from the workbook's own
 * labels). The model does not stop a 1980 cohort from being given ITP-2, and
 * neither does this: such an agreement is a "what if", marked as one, and left
 * out of the default selection.
 */
import {
  RETIREMENT_AGES,
  Scheme,
  Table1Key,
  defaultInput,
  lowestPensionAge,
  riktalderFor,
  run,
} from "@typfallsmodellen/engine";
import type {
  DeathProbabilities,
  ModelContext,
  SchemeId,
  TypfallInput,
  TypfallResult,
} from "@typfallsmodellen/engine";

import type { Lang } from "./i18n.js";
import { retirementAge } from "./kpis.js";
import { defaultSavings, savingSettings, taxFreeSaving } from "./privateSavings.js";
import type { SavingPerson, SavingsInput } from "./privateSavings.js";
import { csvLine, delimiterFor } from "./tables.js";
import { kronorCell } from "./xlsx.js";
import type { SheetData } from "./xlsx.js";

const say = (l: Lang, sv: string, en: string) => (l === "sv" ? sv : en);

/** Alternatives are the reference age and this many years either side. */
export const SPREAD = 3;

/**
 * The last age the study sums to. The model's age matrix runs to 105, but its
 * final years carry end-of-horizon artefacts (the public pension jumps at 101),
 * and nobody is asking about a lifespan past 100.
 */
export const LAST_AGE = 100;

/** The "lever till" range a user can ask for. */
export const LIFESPAN_BOUNDS = { min: 70, max: LAST_AGE } as const;

/** Two best totals closer than this share of the best are "practically equal". */
export const NEAR_TIE = 0.01;

// ---------------------------------------------------------------- colours ---

/** One per offset from the reference age: warm earlier, cool later, ink for the reference. */
export type Slot = "m3" | "m2" | "m1" | "z" | "p1" | "p2" | "p3";

export const SLOTS: readonly Slot[] = ["m3", "m2", "m1", "z", "p1", "p2", "p3"];

export function slotOf(offset: number): Slot {
  const clamped = Math.min(Math.max(Math.round(offset), -SPREAD), SPREAD);
  return SLOTS[clamped + SPREAD]!;
}

// ------------------------------------------------------------------ input ---

export interface StudyInput extends SavingsInput {
  readonly born: number;
  readonly startWorkAge: number;
  /** Per month, in the reference year's money -- `TypfallInput.monthlySalary`. */
  readonly monthlySalary: number;
  readonly scheme: SchemeId;
  readonly yearlyInflation: number;
  readonly realGrowth: number;
  readonly realReturn: number;
  /** The age the alternatives are centred on (the riktålder, unless asked otherwise). */
  readonly centre: number;
  /** The first and last "lever till" age. */
  readonly lifespanFrom: number;
  readonly lifespanTo: number;
  /** Sum `netto` (after tax) instead of the pension before tax. */
  readonly afterTax: boolean;
  /** False runs every alternative as if no collective agreement applied; the person's own saving still counts. */
  readonly includeOccupational: boolean;
}

/**
 * The typfall the static report this tab grew out of used: born 1980, work from
 * 25, 38 500 kr a month, the workbook's own economic assumptions (no inflation,
 * no real wage growth), no collective agreement, centred on the riktålder.
 */
export function defaultStudyInput(): StudyInput {
  const base = defaultInput();
  const born = 1980;
  return {
    born,
    startWorkAge: 25,
    monthlySalary: base.monthlySalary,
    scheme: Scheme.None,
    yearlyInflation: base.yearlyInflation,
    realGrowth: base.realGrowth,
    realReturn: base.realReturn,
    centre: riktalderFor(born),
    lifespanFrom: 80,
    lifespanTo: 95,
    afterTax: false,
    includeOccupational: true,
    ...defaultSavings(),
  };
}

/** The ages a reference age can sit at for a cohort: from its lowest permitted to the sheet's highest. */
export function centreBounds(born: number): { readonly min: number; readonly max: number } {
  return { min: lowestPensionAge(born), max: Math.max(...RETIREMENT_AGES) };
}

/** The reference age actually used: the asked-for one, moved inside what the cohort may draw. */
export function referenceAge(born: number, centre: number): number {
  const { min, max } = centreBounds(born);
  return Math.min(Math.max(Math.round(centre), min), max);
}

/**
 * The alternative retirement ages: the reference and `SPREAD` years either side,
 * ascending, dropping any the cohort may not draw or the sheet does not offer.
 */
export function alternativeAges(born: number, centre: number): readonly number[] {
  const { min, max } = centreBounds(born);
  const middle = referenceAge(born, centre);
  const ages: number[] = [];
  for (let offset = -SPREAD; offset <= SPREAD; offset += 1) {
    const age = middle + offset;
    if (age >= min && age <= max) ages.push(age);
  }
  return ages;
}

/** The typfall for one alternative. Nothing but the study's own fields varies. */
export function studyTypfall(input: StudyInput, retirementAgeValue: number): TypfallInput {
  return defaultInput({
    born: input.born,
    startWorkAge: input.startWorkAge,
    monthlySalary: input.monthlySalary,
    retirementAge: retirementAgeValue,
    scheme: input.includeOccupational ? input.scheme : Scheme.None,
    yearlyInflation: input.yearlyInflation,
    realGrowth: input.realGrowth,
    realReturn: input.realReturn,
  });
}

/** The shared context with the setting the question cannot survive turned off (see the file comment). */
export function studyContext(shared: ModelContext): ModelContext {
  return { ...shared, tjpPar: 0 };
}

/** The person as the saving sees them: it runs until the alternative's own retirement age. */
function studyPerson(input: StudyInput, retirementAgeValue: number): SavingPerson {
  return { born: input.born, startWorkAge: input.startWorkAge, retirementAge: retirementAgeValue };
}

/**
 * The context for one alternative's run: the study's, with the model's private saving carrying
 * the pension side of the savings -- the individual occupational pension and the private
 * pension, the model's IPS, taxed as income.
 */
export function pensionContext(shared: ModelContext, input: StudyInput, retirementAgeValue: number): ModelContext {
  return { ...studyContext(shared), ...savingSettings(shared, studyPerson(input, retirementAgeValue), input, "pension") };
}

/** The same for the other run: the private savings, an ISK or a depot, paid out tax-free (`savingSettings`). */
export function savingsContext(shared: ModelContext, input: StudyInput, retirementAgeValue: number): ModelContext {
  return { ...studyContext(shared), ...savingSettings(shared, studyPerson(input, retirementAgeValue), input, "savings") };
}

// ---------------------------------------------------------------- results ---

export interface StudyOption {
  /** Years from the reference age: -3 .. +3. */
  readonly offset: number;
  readonly slot: Slot;
  /** The retirement age the run actually used. */
  readonly par: number;
  /** The measure, in kronor per year at each age 0..`LAST_AGE`; zero before `par`. */
  readonly annual: readonly number[];
  /** Of which occupational pension, before tax -- also in the after-tax view, where it cannot be split out of net income and is not shown. */
  readonly occupational: readonly number[];
  /** Of which individual occupational pension and private pension (the model's IPS), before tax; taxed as income, so not split out of net income either. */
  readonly individual: readonly number[];
  /** Of which private savings (ISK, depot): not taxed as income, so the same before and after tax. */
  readonly savings: readonly number[];
  /** `cumulative[x]`: kronor paid out before the x-th birthday. */
  readonly cumulative: readonly number[];
  /** The first year's measure per month -- Table 1's own figure, the one Prognos shows at retirement. */
  readonly firstYearMonthly: number;
  /** Table 1's own "Slutlön", per month. */
  readonly finalSalaryMonthly: number;
  /** Table 1's own total at retirement (gross or after tax), per month, with the private savings -- which counts
   * the last year's pension right straight away, where the yearly flow's own first row leaves it out. The first
   * year is this figure (`firstYearMonthly` equals it); every later year is the flow's. */
  readonly table1Monthly: number;
  /** Table 1's share of final salary for the same line. */
  readonly table1Share: number;
  /** Table 1's occupational pension line, per month, before tax. */
  readonly table1OccupationalMonthly: number;
  /** Table 1's private pension saving line (individual and private pension), per month, before tax. */
  readonly table1IndividualMonthly: number;
  /** Table 1's private savings line, per month; not taxed. Also in `table1Monthly`. */
  readonly table1SavingsMonthly: number;
  readonly warnings: readonly string[];
}

export interface LifespanRow {
  readonly lifespan: number;
  /** One total per alternative, in `Study.options` order. */
  readonly totals: readonly number[];
  readonly best: number;
  readonly runnerUp: number;
  /** Best total minus the runner-up's. */
  readonly margin: number;
  /** `margin` as a share of the best total. */
  readonly marginShare: number;
}

export interface BreakEven {
  readonly earlier: StudyOption;
  readonly later: StudyOption;
  /** The lifespan at which `later` has caught up with `earlier`; null if it does not by `LAST_AGE`. */
  readonly at: number | null;
}

export interface Study {
  readonly input: StudyInput;
  readonly options: readonly StudyOption[];
  /** Index into `options` of the reference age. */
  readonly reference: number;
  readonly rows: readonly LifespanRow[];
  /** Every alternative but the reference, against the reference. */
  readonly versusReference: readonly BreakEven[];
  /** Each alternative against the one a year earlier. */
  readonly chain: readonly BreakEven[];
  readonly riktalder: number;
  readonly lowestAge: number;
  readonly priceBasis: number;
  readonly referenceYear: number;
  /** The distinct warnings the runs raised, in the order met. */
  readonly warnings: readonly string[];
}

function cumulativeOf(annual: readonly number[]): number[] {
  const out: number[] = [];
  let sum = 0;
  for (let x = 0; x <= LAST_AGE; x += 1) {
    out.push(sum);
    sum += annual[x] ?? 0;
  }
  return out;
}

function table1Row(result: TypfallResult, key: string): TypfallResult["table1"][number] | undefined {
  return result.table1.find((r) => r.key === key);
}

function table1(result: TypfallResult, key: string): { monthly: number; share: number } {
  const row = table1Row(result, key);
  return { monthly: row?.monthly ?? 0, share: row?.shareOfFinalSalary ?? 0 };
}

/** What the tax-free private savings pay out when the person retires at one age. */
export interface SavingsPayout {
  /** Kronor a year at each age 0..`LAST_AGE`; zero before the retirement age. */
  readonly annual: readonly number[];
  /** Table 1's private savings line at retirement: per month, and in kronor a year at retirement's prices. */
  readonly table1Monthly: number;
  readonly table1Adjusted: number;
  readonly warnings: readonly string[];
}

/** Payouts already worked out, by retirement age -- valid for studies of one person, whatever their agreement. */
export type SavingsPayouts = Map<number, SavingsPayout>;

/**
 * The private savings' payouts for one retirement age; undefined when there are no private
 * savings. They are a run of their own -- the model carries one kind of private saving at a
 * time, and this one is not taxed as income (`savingSettings`) -- and the agreement plays no
 * part in them, so the run is made without one and `known` keeps it for the next study of the
 * same person.
 */
function savingsPayout(
  input: StudyInput,
  shared: ModelContext,
  deaths: DeathProbabilities,
  requested: number,
  known: SavingsPayouts,
): SavingsPayout | undefined {
  if (taxFreeSaving(input) === null) return undefined;
  const kept = known.get(requested);
  if (kept !== undefined) return kept;

  const typfall = { ...studyTypfall(input, requested), scheme: Scheme.None };
  const result = run(typfall, savingsContext(shared, input, requested), { deaths });
  const par = retirementAge(typfall, result);
  const annual = new Array<number>(LAST_AGE + 1).fill(0);
  for (const row of result.rows) {
    if (row.age >= par && row.age <= LAST_AGE) annual[row.age] = row.pps;
  }
  const line = table1Row(result, Table1Key.PrivateSavingAfterTax);
  // Table 1's own figure is the first year's, as it is for the pensions (see `buildOption`).
  if (line !== undefined && par <= LAST_AGE) annual[par] = line.monthly * 12;
  const payout: SavingsPayout = {
    annual,
    table1Monthly: line?.monthly ?? 0,
    table1Adjusted: line?.adjusted ?? 0,
    warnings: result.warnings.map((w) => w.message),
  };
  known.set(requested, payout);
  return payout;
}

function buildOption(
  input: StudyInput,
  shared: ModelContext,
  deaths: DeathProbabilities,
  requested: number,
  middle: number,
  known: SavingsPayouts,
): StudyOption | undefined {
  const typfall = studyTypfall(input, requested);
  const result = run(typfall, pensionContext(shared, input, requested), { deaths });
  const par = retirementAge(typfall, result);
  // A run that corrected the age is another alternative's twin -- skip it
  // rather than draw the same line twice under the wrong label.
  if (par !== requested) return undefined;
  const own = savingsPayout(input, shared, deaths, requested, known);

  const annual = new Array<number>(LAST_AGE + 1).fill(0);
  const occupational = new Array<number>(LAST_AGE + 1).fill(0);
  const individual = new Array<number>(LAST_AGE + 1).fill(0);
  const savings = new Array<number>(LAST_AGE + 1).fill(0);
  for (const row of result.rows) {
    if (row.age < par || row.age > LAST_AGE) continue;
    const publicPension = row.ip + row.tp + row.pp + row.garp + row.ptillagg;
    const untaxed = own?.annual[row.age] ?? 0;
    annual[row.age] = (input.afterTax ? row.netto : publicPension + row.tjp + row.ips) + untaxed;
    occupational[row.age] = row.tjp;
    individual[row.age] = row.ips;
    savings[row.age] = untaxed;
  }

  // Table 1's total counts the private savings too: they are not in its gross, and its own line for them is a separate row.
  const total = table1(result, input.afterTax ? Table1Key.PensionAfterTax : Table1Key.TotalGross);
  // The first year is what Table 1 -- and so Prognos -- says. The model draws Table 1 in a second pass at the
  // retirement age, once the last working year's pension right is in; the yearly flow has it from the next year, so
  // its own first row is a little lower. Every later year stays the flow's. The pension is paid out against the
  // amount the person was told at retirement, which is what Prognos shows.
  if (par <= LAST_AGE) {
    annual[par] = total.monthly * 12 + (own?.annual[par] ?? 0);
    occupational[par] = table1(result, Table1Key.OccupationalPension).monthly * 12;
    individual[par] = table1(result, Table1Key.PrivateSaving).monthly * 12;
    savings[par] = own?.annual[par] ?? 0;
  }
  const base = table1Row(result, input.afterTax ? Table1Key.SalaryAfterTax : Table1Key.FinalSalary)?.adjusted ?? 0;
  const savingsShare = own !== undefined && base > 0 ? own.table1Adjusted / base : 0;
  return {
    offset: requested - middle,
    slot: slotOf(requested - middle),
    par,
    annual,
    occupational,
    individual,
    savings,
    cumulative: cumulativeOf(annual),
    firstYearMonthly: (annual[par] ?? 0) / 12,
    finalSalaryMonthly: table1(result, Table1Key.FinalSalary).monthly,
    table1Monthly: total.monthly + (own?.table1Monthly ?? 0),
    table1Share: total.share + savingsShare,
    table1OccupationalMonthly: table1(result, Table1Key.OccupationalPension).monthly,
    table1IndividualMonthly: table1(result, Table1Key.PrivateSaving).monthly,
    table1SavingsMonthly: own?.table1Monthly ?? 0,
    warnings: [...result.warnings.map((w) => w.message), ...(own?.warnings ?? [])],
  };
}

/**
 * The lifespan at which `later` has caught up with `earlier`, or null.
 *
 * `later` starts behind (it has paid nothing while `earlier` has) and, being a
 * larger monthly amount, closes the gap each year; the running totals are
 * linear inside a year, so the crossing is interpolated exactly.
 */
export function breakEven(earlier: StudyOption, later: StudyOption): number | null {
  for (let x = later.par; x < LAST_AGE; x += 1) {
    const here = later.cumulative[x]! - earlier.cumulative[x]!;
    const next = later.cumulative[x + 1]! - earlier.cumulative[x + 1]!;
    if (here >= 0) return x;
    if (next >= 0) return x + -here / (next - here);
  }
  return null;
}

/**
 * Who wins at one lifespan: every alternative's total, the best and the
 * runner-up, and how far apart they are. Ties go to the earlier age.
 */
export function lifespanRow(options: readonly StudyOption[], lifespan: number): LifespanRow {
  const totals = options.map((o) => o.cumulative[lifespan] ?? 0);
  const order = totals
    .map((total, index) => ({ total, index }))
    .sort((a, b) => b.total - a.total || a.index - b.index);
  const best = order[0]!;
  const second = order[1];
  const margin = second === undefined ? 0 : best.total - second.total;
  return {
    lifespan,
    totals,
    best: best.index,
    runnerUp: second?.index ?? best.index,
    margin,
    marginShare: best.total > 0 ? margin / best.total : 0,
  };
}

/**
 * What one alternative pays a month on average if its person lives to `lifespan`: the total paid out before
 * that age spread over the months from the retirement age to it. Undefined when nothing is paid out by then.
 */
export function averageMonthly(option: StudyOption, lifespan: number): number | undefined {
  const months = (lifespan - option.par) * 12;
  return months > 0 ? (option.cumulative[lifespan] ?? 0) / months : undefined;
}

/**
 * Runs the study: one `run()` per alternative -- two, with private savings -- then the sums and
 * break-evens. `known` holds the private savings' payouts already worked out for this person,
 * and is filled in with any that were not (`computeComparison` shares one between agreements).
 */
export function computeStudy(
  input: StudyInput,
  shared: ModelContext,
  deaths: DeathProbabilities,
  known: SavingsPayouts = new Map(),
): Study {
  const context = studyContext(shared);
  const ages = alternativeAges(input.born, input.centre);
  const middle = referenceAge(input.born, input.centre);

  const options: StudyOption[] = [];
  for (const age of ages) {
    const option = buildOption(input, shared, deaths, age, middle, known);
    if (option !== undefined) options.push(option);
  }

  let reference = options.findIndex((o) => o.offset === 0);
  if (reference === -1) reference = 0;

  const rows: LifespanRow[] = [];
  for (let lifespan = input.lifespanFrom; lifespan <= input.lifespanTo; lifespan += 1) {
    rows.push(lifespanRow(options, lifespan));
  }

  const ref = options[reference];
  const versusReference: BreakEven[] = [];
  if (ref !== undefined) {
    for (const option of options) {
      if (option === ref) continue;
      const [earlier, later] = option.par < ref.par ? [option, ref] : [ref, option];
      versusReference.push({ earlier, later, at: breakEven(earlier, later) });
    }
  }
  const chain: BreakEven[] = [];
  for (let i = 1; i < options.length; i += 1) {
    chain.push({ earlier: options[i - 1]!, later: options[i]!, at: breakEven(options[i - 1]!, options[i]!) });
  }

  const warnings: string[] = [];
  for (const option of options) {
    for (const message of option.warnings) if (!warnings.includes(message)) warnings.push(message);
  }

  return {
    input,
    options,
    reference,
    rows,
    versusReference,
    chain,
    riktalder: riktalderFor(input.born),
    lowestAge: lowestPensionAge(input.born),
    priceBasis: context.priceBasis,
    referenceYear: context.referenceYear,
    warnings,
  };
}

/** Which kinds of pension a study pays in the first year of its alternatives -- what a table can split that year into. */
export interface KindsPaid {
  /** The agreement's occupational pension. */
  readonly occupational: boolean;
  /** The individual occupational pension and the private pension. */
  readonly individual: boolean;
  /** The private savings. */
  readonly savings: boolean;
}

export function kindsPaid(study: Study): KindsPaid {
  const paid = (pick: (option: StudyOption) => readonly number[]): boolean =>
    study.options.some((option) => (pick(option)[option.par] ?? 0) > 0);
  return {
    occupational: paid((o) => o.occupational),
    individual: paid((o) => o.individual),
    savings: paid((o) => o.savings),
  };
}

/** Which alternative is best over each unbroken run of lifespans: "80-82: 65, 83: 66, ...". */
export function bestRuns(study: Study): readonly { from: number; to: number; option: StudyOption }[] {
  const runs: { from: number; to: number; option: StudyOption }[] = [];
  for (const row of study.rows) {
    const option = study.options[row.best]!;
    const last = runs[runs.length - 1];
    if (last !== undefined && last.option === option) last.to = row.lifespan;
    else runs.push({ from: row.lifespan, to: row.lifespan, option });
  }
  return runs;
}

/**
 * The monthly salary whose final salary (Table 1's "Slutlön", per month) is
 * `target`, found by running the model through `finalAt`.
 *
 * The final salary is the average of the last years' income, and income is the
 * monthly salary times a wage path that does not depend on it -- so the
 * relationship is linear and a couple of secant steps land within rounding.
 */
export function monthlySalaryForFinalSalary(target: number, finalAt: (monthlySalary: number) => number): number {
  if (!(target > 0)) return 0;
  let salary = target;
  for (let step = 0; step < 4; step += 1) {
    const got = finalAt(salary);
    if (!(got > 0)) break;
    const next = salary * (target / got);
    const settled = Math.abs(next - salary) < 0.25;
    salary = next;
    if (settled) break;
  }
  return Math.round(salary);
}

/** Table 1's "Slutlön" per month for one alternative -- what a typed final salary is matched against. */
export function finalSalaryOf(
  input: StudyInput,
  retirementAgeValue: number,
  shared: ModelContext,
  deaths: DeathProbabilities,
): number {
  const result = run(studyTypfall(input, retirementAgeValue), pensionContext(shared, input, retirementAgeValue), { deaths });
  return table1(result, Table1Key.FinalSalary).monthly;
}

/**
 * The monthly salary that gives `target` as the *reference* alternative's final
 * salary. The other alternatives may end up with a different final salary (a
 * wage path that grows, or an age profile, makes it depend on the retirement
 * age); the typed figure is matched at the reference age, where the question
 * "what do I earn when I retire?" is asked.
 */
export function monthlySalaryForTarget(
  input: StudyInput,
  target: number,
  shared: ModelContext,
  deaths: DeathProbabilities,
): number {
  const age = referenceAge(input.born, input.centre);
  return monthlySalaryForFinalSalary(target, (monthlySalary) =>
    finalSalaryOf({ ...input, monthlySalary }, age, shared, deaths),
  );
}

// ------------------------------------------------------------- agreements ---

/** The agreements in the workbook's own order -- its drop-down's, and so this tab's -- "Saknar tjänstepension" first. */
export const SCHEME_ORDER: readonly SchemeId[] = [
  Scheme.None,
  Scheme.Itp1,
  Scheme.Itp2,
  Scheme.SafLo,
  Scheme.KapKl,
  Scheme.AkapKr,
  Scheme.Pa16Avd2,
  Scheme.Pa16Avd1,
];

/** The seven that are an agreement: all but the baseline. */
export const AGREEMENTS: readonly SchemeId[] = SCHEME_ORDER.filter((scheme) => scheme !== Scheme.None);

/** Birth years an agreement is written for; either end may be open. */
export interface CohortRange {
  readonly from?: number;
  readonly to?: number;
}

/**
 * The birth years each agreement covers, as the workbook's own labels word them
 * ("ITP-1 ... födda 1979-", "ITP-2 ... födda före 1979", "AKAP-KR ... födda
 * 1986-", "PA16 (Avd 2) ... födda före 1988", "PA16 (Avd 1) ... födda 1988 och
 * senare"). The others carry no year in their label and fit every cohort.
 * Reality is looser than a label -- KAP-KL and AKAP-KR overlap for those who
 * chose to switch -- which is why this only marks an agreement as a "what if"
 * and never removes it.
 */
export const SCHEME_COHORTS: Readonly<Partial<Record<SchemeId, CohortRange>>> = {
  [Scheme.Itp1]: { from: 1979 },
  [Scheme.Itp2]: { to: 1978 },
  [Scheme.AkapKr]: { from: 1986 },
  [Scheme.Pa16Avd2]: { to: 1987 },
  [Scheme.Pa16Avd1]: { from: 1988 },
};

export function schemeFitsCohort(scheme: SchemeId, born: number): boolean {
  const range = SCHEME_COHORTS[scheme];
  if (range === undefined) return true;
  return (range.from === undefined || born >= range.from) && (range.to === undefined || born <= range.to);
}

/** "födda 1979 och senare", "born before 1979" -- undefined for an agreement with no cohort in its label. */
export function cohortText(scheme: SchemeId, lang: Lang): string | undefined {
  const range = SCHEME_COHORTS[scheme];
  if (range?.from !== undefined) return say(lang, `födda ${range.from} och senare`, `born ${range.from} or later`);
  if (range?.to !== undefined) return say(lang, `födda före ${range.to + 1}`, `born before ${range.to + 1}`);
  return undefined;
}

const SCHEME_NAMES: Readonly<
  Record<SchemeId, { readonly sv: string; readonly en: string; readonly sectorSv: string; readonly sectorEn: string }>
> = {
  [Scheme.None]: {
    sv: "Saknar tjänstepension",
    en: "No occupational pension",
    sectorSv: "Inget kollektivavtal",
    sectorEn: "No collective agreement",
  },
  [Scheme.Itp1]: { sv: "ITP-1", en: "ITP-1", sectorSv: "Privatanställda tjänstemän", sectorEn: "Private-sector salaried employees" },
  [Scheme.Itp2]: { sv: "ITP-2", en: "ITP-2", sectorSv: "Privatanställda tjänstemän", sectorEn: "Private-sector salaried employees" },
  [Scheme.SafLo]: { sv: "SAF-LO", en: "SAF-LO", sectorSv: "Privatanställda arbetare", sectorEn: "Private-sector blue-collar workers" },
  [Scheme.KapKl]: {
    sv: "KAP-KL",
    en: "KAP-KL",
    sectorSv: "Kommunal- och regionalanställda",
    sectorEn: "Municipal and regional employees",
  },
  [Scheme.AkapKr]: {
    sv: "AKAP-KR",
    en: "AKAP-KR",
    sectorSv: "Kommunal- och regionalanställda",
    sectorEn: "Municipal and regional employees",
  },
  [Scheme.Pa16Avd2]: { sv: "PA16 avd 2", en: "PA16 section 2", sectorSv: "Statligt anställda", sectorEn: "Central government employees" },
  [Scheme.Pa16Avd1]: { sv: "PA16 avd 1", en: "PA16 section 1", sectorSv: "Statligt anställda", sectorEn: "Central government employees" },
};

/** The short name a table row or a chart label carries. */
export function schemeName(scheme: SchemeId, lang: Lang): string {
  const names = SCHEME_NAMES[scheme];
  return say(lang, names.sv, names.en);
}

/** Who the agreement is for: "Privatanställda tjänstemän". */
export function schemeSector(scheme: SchemeId, lang: Lang): string {
  const names = SCHEME_NAMES[scheme];
  return say(lang, names.sectorSv, names.sectorEn);
}

/** The agreements a comparison shows until someone picks: those the cohort can belong to, and the one chosen in the inputs. */
export function defaultSchemes(born: number, chosen: SchemeId): readonly SchemeId[] {
  return AGREEMENTS.filter((scheme) => schemeFitsCohort(scheme, born) || scheme === chosen);
}

export interface AgreementEntry {
  readonly scheme: SchemeId;
  /** False when the agreement is written for other birth years than this person's. */
  readonly fits: boolean;
  /** The seven alternatives under this agreement; its `input` carries the agreement. */
  readonly study: Study;
}

export interface Comparison {
  readonly input: StudyInput;
  /** The baseline (no agreement) first, then the chosen agreements in the workbook's order. */
  readonly entries: readonly AgreementEntry[];
  /** The alternative retirement ages, ascending -- the same under every agreement. */
  readonly ages: readonly number[];
  /** Index into `ages` of the reference age. */
  readonly reference: number;
  /** The "lever till" ages the study covers. */
  readonly lifespans: readonly number[];
  /** The distinct warnings any of the runs raised, in the order met. */
  readonly warnings: readonly string[];
}

/**
 * What a cached agreement study is valid for: everything in the input but the
 * agreement itself, which is what a comparison varies. A caller keeps its
 * studies while this and the shared context stay the same, and empties the
 * cache when either changes.
 */
export function comparisonKey(input: StudyInput): string {
  const { scheme, includeOccupational, ...rest } = input;
  return JSON.stringify(rest);
}

/**
 * One study per agreement. `cache` holds studies already computed for the same
 * `comparisonKey` and context, and is filled in with any that were not.
 */
export function computeComparison(
  input: StudyInput,
  schemes: readonly SchemeId[],
  shared: ModelContext,
  deaths: DeathProbabilities,
  cache: Map<SchemeId, Study> = new Map(),
): Comparison {
  const entries: AgreementEntry[] = [];
  // The same person under every agreement, so their private savings pay out the same: worked out once per age.
  const known: SavingsPayouts = new Map();
  for (const scheme of SCHEME_ORDER) {
    if (scheme !== Scheme.None && !schemes.includes(scheme)) continue;
    let study = cache.get(scheme);
    if (study === undefined) {
      study = computeStudy({ ...input, scheme, includeOccupational: true }, shared, deaths, known);
      cache.set(scheme, study);
    }
    entries.push({ scheme, fits: schemeFitsCohort(scheme, input.born), study });
  }

  const baseline = entries[0]!.study;
  const warnings: string[] = [];
  for (const entry of entries) {
    for (const message of entry.study.warnings) if (!warnings.includes(message)) warnings.push(message);
  }
  const lifespans: number[] = [];
  for (let lifespan = input.lifespanFrom; lifespan <= input.lifespanTo; lifespan += 1) lifespans.push(lifespan);
  return {
    input,
    entries,
    ages: baseline.options.map((o) => o.par),
    reference: baseline.reference,
    lifespans,
    warnings,
  };
}

/** What a total is measured at: the best alternative for each lifespan, or one retirement age. */
export type AgeBasis = "best" | number;

export type Metric = "total" | "extra";

/** The alternative a figure is read at: the best one for `lifespan`, or the asked-for age (the reference if it is not offered). */
function optionAt(entry: AgreementEntry, lifespan: number, basis: AgeBasis): StudyOption {
  const { options, reference } = entry.study;
  if (basis !== "best") {
    const fixed = options.find((o) => o.par === basis);
    if (fixed !== undefined) return fixed;
    return options[reference]!;
  }
  return options[lifespanRow(options, lifespan).best]!;
}

/** Kronor paid out before the `lifespan`-year birthday, at the best age or at a fixed one. */
export function totalFor(entry: AgreementEntry, lifespan: number, basis: AgeBasis): number {
  return optionAt(entry, lifespan, basis).cumulative[lifespan] ?? 0;
}

export interface ComparisonRow {
  readonly entry: AgreementEntry;
  /** The alternative the row's figures are at. */
  readonly option: StudyOption;
  /** The best alternative at this lifespan, whatever the basis. */
  readonly best: StudyOption;
  /** The best and the runner-up are within `NEAR_TIE` of each other. */
  readonly tie: boolean;
  readonly total: number;
  /** `total` minus the baseline's at the same basis. */
  readonly extra: number;
  /** `extra` as a share of the baseline's total. */
  readonly extraShare: number;
  /** The first year's pension per month, at `option`. */
  readonly monthly: number;
  /** Of which occupational pension, before tax. */
  readonly occupationalMonthly: number;
}

/** One row per entry, the baseline first: what each agreement pays by `lifespan`. */
export function comparisonRows(comparison: Comparison, lifespan: number, basis: AgeBasis): ComparisonRow[] {
  const baseline = comparison.entries[0]!;
  const baselineTotal = totalFor(baseline, lifespan, basis);
  return comparison.entries.map((entry) => {
    const row = lifespanRow(entry.study.options, lifespan);
    const option = optionAt(entry, lifespan, basis);
    const total = option.cumulative[lifespan] ?? 0;
    const extra = total - baselineTotal;
    return {
      entry,
      option,
      best: entry.study.options[row.best]!,
      tie: row.marginShare < NEAR_TIE,
      total,
      extra,
      extraShare: baselineTotal > 0 ? extra / baselineTotal : 0,
      monthly: option.firstYearMonthly,
      occupationalMonthly: (option.occupational[option.par] ?? 0) / 12,
    };
  });
}

export interface AgreementSeries {
  readonly entry: AgreementEntry;
  /** (lifespan, kronor): the total, or what the agreement adds to the baseline's. */
  readonly points: readonly (readonly [number, number])[];
}

/** One line per agreement over the lifespans. The baseline adds nothing to itself, so it has no line in `extra`. */
export function totalSeries(comparison: Comparison, basis: AgeBasis, metric: Metric): AgreementSeries[] {
  const baseline = comparison.entries[0]!;
  const out: AgreementSeries[] = [];
  for (const entry of comparison.entries) {
    if (metric === "extra" && entry === baseline) continue;
    const points = comparison.lifespans.map((lifespan) => {
      const total = totalFor(entry, lifespan, basis);
      return [lifespan, metric === "total" ? total : total - totalFor(baseline, lifespan, basis)] as const;
    });
    out.push({ entry, points });
  }
  return out;
}

export interface AgeSpread {
  readonly lifespan: number;
  /** The lowest and the highest best age any entry has at this lifespan. */
  readonly min: number;
  readonly max: number;
}

/** How far apart the agreements' best ages are, lifespan by lifespan. */
export function bestAgeSpread(comparison: Comparison): AgeSpread[] {
  return comparison.lifespans.map((lifespan) => {
    const ages = comparison.entries.map((entry) => entry.study.options[lifespanRow(entry.study.options, lifespan).best]!.par);
    return { lifespan, min: Math.min(...ages), max: Math.max(...ages) };
  });
}

// ----------------------------------------------------------------- export ---

/** Plain integer kronor -- an export meant for Excel's own arithmetic, not for reading aloud. */
const whole = (n: number): string => String(Math.round(n));

/** The lifespan matrix: one row per "lever till" age, one total per alternative, then the winner. */
export function matrixToCsv(study: Study, lang: Lang): string {
  const delimiter = delimiterFor(lang);
  const lines = [
    csvLine(
      [
        say(lang, "Lever till (år)", "Lifespan (years)"),
        ...study.options.map((o) => String(o.par)),
        say(lang, "Bäst", "Best"),
        say(lang, "Försprång mot näst bäst (kr)", "Lead over runner-up (SEK)"),
      ],
      delimiter,
    ),
  ];
  for (const row of study.rows) {
    lines.push(
      csvLine(
        [String(row.lifespan), ...row.totals.map(whole), String(study.options[row.best]!.par), whole(row.margin)],
        delimiter,
      ),
    );
  }
  return lines.join("\r\n");
}

export function matrixToXlsxRows(study: Study, lang: Lang): SheetData {
  const rows: SheetData = [
    [
      say(lang, "Lever till (år)", "Lifespan (years)"),
      ...study.options.map((o) => String(o.par)),
      say(lang, "Bäst", "Best"),
      say(lang, "Försprång mot näst bäst (kr)", "Lead over runner-up (SEK)"),
    ],
  ];
  for (const row of study.rows) {
    rows.push([
      row.lifespan,
      ...row.totals.map((t) => kronorCell(t)),
      study.options[row.best]!.par,
      kronorCell(row.margin),
    ]);
  }
  return rows;
}

/** The first age any alternative pays out, through the last "lever till" age. */
function seriesAges(study: Study): number[] {
  const first = Math.min(...study.options.map((o) => o.par));
  const ages: number[] = [];
  for (let age = first; age <= study.input.lifespanTo; age += 1) ages.push(age);
  return ages;
}

/**
 * What the per-age table, and so its download, shows: the monthly pension at each age, or the total paid out
 * before each age. The table has a switch for it, and the download is of what the switch shows.
 */
export type SeriesMeasure = "monthly" | "total";

function seriesHeader(study: Study, lang: Lang, measure: SeriesMeasure): string[] {
  return measure === "monthly"
    ? [
        say(lang, "Ålder", "Age"),
        ...study.options.map((o) => `${say(lang, "Pension per månad", "Pension per month")} ${o.par}`),
      ]
    : [
        say(lang, "Före ålder", "Before age"),
        ...study.options.map((o) => `${say(lang, "Total utbetalning", "Total payout")} ${o.par}`),
      ];
}

/** Per age: the monthly pension, or the sum paid out before the age, for each alternative. Blank before it starts. */
function seriesCell(o: StudyOption, age: number, measure: SeriesMeasure): number | undefined {
  if (age < o.par) return undefined;
  return measure === "monthly" ? (o.annual[age] ?? 0) / 12 : (o.cumulative[age] ?? 0);
}

export function seriesToCsv(study: Study, lang: Lang, measure: SeriesMeasure): string {
  const delimiter = delimiterFor(lang);
  const lines = [csvLine(seriesHeader(study, lang, measure), delimiter)];
  for (const age of seriesAges(study)) {
    lines.push(
      csvLine(
        [
          String(age),
          ...study.options.map((o) => {
            const value = seriesCell(o, age, measure);
            return value === undefined ? "" : whole(value);
          }),
        ],
        delimiter,
      ),
    );
  }
  return lines.join("\r\n");
}

export function seriesToXlsxRows(study: Study, lang: Lang, measure: SeriesMeasure): SheetData {
  const rows: SheetData = [seriesHeader(study, lang, measure)];
  for (const age of seriesAges(study)) {
    rows.push([
      age,
      ...study.options.map((o) => {
        const value = seriesCell(o, age, measure);
        return value === undefined ? "" : kronorCell(value);
      }),
    ]);
  }
  return rows;
}

// --------------------------------------------------- optimal age by lifespan ---

/** What the best start age is, and pays, for one lifespan: a column of "Optimal pensionsålder i förhållande till livslängd". */
export interface OptimalRow {
  readonly lifespan: number;
  /** The best retirement age for this lifespan. */
  readonly par: number;
  /** What it pays a month on average (`averageMonthly`); undefined when nothing is paid by then. */
  readonly monthly: number | undefined;
  /** What it pays in total before the lifespan, kronor. */
  readonly total: number;
}

export function optimalRows(study: Study): OptimalRow[] {
  return study.rows.map((r) => {
    const best = study.options[r.best]!;
    return { lifespan: r.lifespan, par: best.par, monthly: averageMonthly(best, r.lifespan), total: r.totals[r.best] ?? 0 };
  });
}

/** The labels of the table's rows, in its own order. */
function optimalLabels(lang: Lang): { lifespan: string; age: string; monthly: string; total: string } {
  return {
    lifespan: say(lang, "Livslängd", "Lifespan"),
    age: say(lang, "Pensionsålder", "Retirement age"),
    monthly: say(lang, "Månadsbelopp kr", "Monthly amount SEK"),
    total: say(lang, "Livsutbetalning mkr", "Lifetime payout MSEK"),
  };
}

/** Millions of kronor with two decimals, the decimal sign of the delimiter's own language. */
function millions(value: number, lang: Lang): string {
  const text = (Math.round(value / 1e4) / 100).toFixed(2);
  return lang === "sv" ? text.replace(".", ",") : text;
}

/** The table as it is shown: one row per measure, one column per lifespan. */
export function optimalToCsv(study: Study, lang: Lang): string {
  const delimiter = delimiterFor(lang);
  const rows = optimalRows(study);
  const label = optimalLabels(lang);
  return [
    csvLine([label.lifespan, ...rows.map((r) => String(r.lifespan))], delimiter),
    csvLine([label.age, ...rows.map((r) => String(r.par))], delimiter),
    csvLine([label.monthly, ...rows.map((r) => (r.monthly === undefined ? "" : whole(r.monthly)))], delimiter),
    csvLine([label.total, ...rows.map((r) => millions(r.total, lang))], delimiter),
  ].join("\r\n");
}

export function optimalToXlsxRows(study: Study, lang: Lang): SheetData {
  const rows = optimalRows(study);
  const label = optimalLabels(lang);
  return [
    [label.lifespan, ...rows.map((r) => r.lifespan)],
    [label.age, ...rows.map((r) => r.par)],
    [label.monthly, ...rows.map((r) => (r.monthly === undefined ? "" : kronorCell(r.monthly)))],
    [label.total, ...rows.map((r) => ({ value: r.total / 1e6, type: Number, format: "0.0" }))],
  ];
}

/** The comparison's two numbers per agreement and lifespan: its best age, and what it pays in total at that age. */
function comparisonHeader(comparison: Comparison, lang: Lang): string[] {
  const tax = comparison.input.afterTax ? say(lang, "efter skatt", "after tax") : say(lang, "före skatt", "before tax");
  return [
    say(lang, "Lever till (år)", "Lifespan (years)"),
    ...comparison.entries.flatMap((entry) => {
      const name = schemeName(entry.scheme, lang);
      return [
        `${name}: ${say(lang, "bästa ålder", "best age")}`,
        `${name}: ${say(lang, `total ${tax} (kr)`, `total ${tax} (SEK)`)}`,
      ];
    }),
  ];
}

function comparisonCells(comparison: Comparison, lifespan: number): { age: number; total: number }[] {
  return comparison.entries.map((entry) => {
    const option = optionAt(entry, lifespan, "best");
    return { age: option.par, total: option.cumulative[lifespan] ?? 0 };
  });
}

export function comparisonToCsv(comparison: Comparison, lang: Lang): string {
  const delimiter = delimiterFor(lang);
  const lines = [csvLine(comparisonHeader(comparison, lang), delimiter)];
  for (const lifespan of comparison.lifespans) {
    lines.push(
      csvLine(
        [String(lifespan), ...comparisonCells(comparison, lifespan).flatMap((c) => [String(c.age), whole(c.total)])],
        delimiter,
      ),
    );
  }
  return lines.join("\r\n");
}

export function comparisonToXlsxRows(comparison: Comparison, lang: Lang): SheetData {
  const rows: SheetData = [comparisonHeader(comparison, lang)];
  for (const lifespan of comparison.lifespans) {
    rows.push([lifespan, ...comparisonCells(comparison, lifespan).flatMap((c) => [c.age, kronorCell(c.total)])]);
  }
  return rows;
}
