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
 * it reuses is `run()` itself, once per alternative retirement age (a run is
 * well under a millisecond, so ten of them per keystroke needs no worker, no
 * memoisation and no debounce -- `main.ts`'s own stated philosophy).
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
 *   occupational pension, before tax; or `netto`, after tax, from retirement.
 *   Amounts are in whatever price basis the shared context asks for (fixed
 *   reference-year prices by default) and are **not discounted**: a krona at 90
 *   counts as much as one at 70.
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
 * Two settings of the shared (Avancerat) context are overridden, on purpose,
 * because they would break the question being asked:
 *
 * - `tjpPar: 0` -- occupational pension follows each alternative's own
 *   retirement age. An Avancerat "draw occupational pension from age X" would
 *   otherwise pay it before the public pension starts for some alternatives
 *   and not others, and `netto` (which still contains salary before
 *   retirement) could not be summed from that age.
 * - `ipsMonthly`/`pbhIps: 0` -- private saving is out of scope here; the study
 *   is about allmän pension and the collective agreements.
 *
 * Every other shared setting (tax, wage profile, own salary path excluded --
 * that is a `TypfallInput` field this tab does not read -- price basis, the
 * "temporary occupational withdrawal" years, ...) applies as in the other tabs.
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

export interface StudyInput {
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
  /** False runs every alternative as if no collective agreement applied. */
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

/** The shared context with the two settings the question cannot survive turned off (see the file comment). */
export function studyContext(shared: ModelContext): ModelContext {
  return { ...shared, tjpPar: 0, ipsMonthly: 0, pbhIps: 0 };
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
  /** Of which occupational pension (before tax; zero in the after-tax view, which cannot be split). */
  readonly occupational: readonly number[];
  /** `cumulative[x]`: kronor paid out before the x-th birthday. */
  readonly cumulative: readonly number[];
  /** The first year's measure per month -- what the first row of the age matrix says. */
  readonly firstYearMonthly: number;
  /** Table 1's own "Slutlön", per month. */
  readonly finalSalaryMonthly: number;
  /** Table 1's own total at retirement (gross or after tax), per month -- which counts
   * the last year's pension right straight away, where the yearly flow picks it up in year two. */
  readonly table1Monthly: number;
  /** Table 1's share of final salary for the same line. */
  readonly table1Share: number;
  /** Table 1's occupational pension line, per month, before tax. */
  readonly table1OccupationalMonthly: number;
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

function table1(result: TypfallResult, key: string): { monthly: number; share: number } {
  const row = result.table1.find((r) => r.key === key);
  return { monthly: row?.monthly ?? 0, share: row?.shareOfFinalSalary ?? 0 };
}

function buildOption(
  input: StudyInput,
  context: ModelContext,
  deaths: DeathProbabilities,
  requested: number,
  middle: number,
): StudyOption | undefined {
  const typfall = studyTypfall(input, requested);
  const result = run(typfall, context, { deaths });
  const par = retirementAge(typfall, result);
  // A run that corrected the age is another alternative's twin -- skip it
  // rather than draw the same line twice under the wrong label.
  if (par !== requested) return undefined;

  const annual = new Array<number>(LAST_AGE + 1).fill(0);
  const occupational = new Array<number>(LAST_AGE + 1).fill(0);
  for (const row of result.rows) {
    if (row.age < par || row.age > LAST_AGE) continue;
    const publicPension = row.ip + row.tp + row.pp + row.garp + row.ptillagg;
    annual[row.age] = input.afterTax ? row.netto : publicPension + row.tjp;
    occupational[row.age] = row.tjp;
  }

  const total = table1(result, input.afterTax ? Table1Key.PensionAfterTax : Table1Key.TotalGross);
  return {
    offset: requested - middle,
    slot: slotOf(requested - middle),
    par,
    annual,
    occupational,
    cumulative: cumulativeOf(annual),
    firstYearMonthly: (annual[par] ?? 0) / 12,
    finalSalaryMonthly: table1(result, Table1Key.FinalSalary).monthly,
    table1Monthly: total.monthly,
    table1Share: total.share,
    table1OccupationalMonthly: table1(result, Table1Key.OccupationalPension).monthly,
    warnings: result.warnings.map((w) => w.message),
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

/** Runs the study: one `run()` per alternative, then the sums and break-evens. */
export function computeStudy(input: StudyInput, shared: ModelContext, deaths: DeathProbabilities): Study {
  const context = studyContext(shared);
  const ages = alternativeAges(input.born, input.centre);
  const middle = referenceAge(input.born, input.centre);

  const options: StudyOption[] = [];
  for (const age of ages) {
    const option = buildOption(input, context, deaths, age, middle);
    if (option !== undefined) options.push(option);
  }

  let reference = options.findIndex((o) => o.offset === 0);
  if (reference === -1) reference = 0;

  const rows: LifespanRow[] = [];
  for (let lifespan = input.lifespanFrom; lifespan <= input.lifespanTo; lifespan += 1) {
    const totals = options.map((o) => o.cumulative[lifespan] ?? 0);
    const order = totals
      .map((total, index) => ({ total, index }))
      .sort((a, b) => b.total - a.total || a.index - b.index);
    const best = order[0]!;
    const second = order[1];
    const margin = second === undefined ? 0 : best.total - second.total;
    rows.push({
      lifespan,
      totals,
      best: best.index,
      runnerUp: second?.index ?? best.index,
      margin,
      marginShare: best.total > 0 ? margin / best.total : 0,
    });
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
  const result = run(studyTypfall(input, retirementAgeValue), studyContext(shared), { deaths });
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

/** Per age: the monthly pension and the running total for each alternative. */
export function seriesToCsv(study: Study, lang: Lang): string {
  const delimiter = delimiterFor(lang);
  const lines = [
    csvLine(
      [
        say(lang, "Ålder", "Age"),
        ...study.options.map((o) => `${say(lang, "Pension per månad", "Pension per month")} ${o.par}`),
        ...study.options.map((o) => `${say(lang, "Ackumulerat före årsdagen", "Cumulative before the birthday")} ${o.par}`),
      ],
      delimiter,
    ),
  ];
  for (const age of seriesAges(study)) {
    lines.push(
      csvLine(
        [
          String(age),
          ...study.options.map((o) => (age >= o.par ? whole((o.annual[age] ?? 0) / 12) : "")),
          ...study.options.map((o) => (age >= o.par ? whole(o.cumulative[age] ?? 0) : "")),
        ],
        delimiter,
      ),
    );
  }
  return lines.join("\r\n");
}

export function seriesToXlsxRows(study: Study, lang: Lang): SheetData {
  const rows: SheetData = [
    [
      say(lang, "Ålder", "Age"),
      ...study.options.map((o) => `${say(lang, "Pension per månad", "Pension per month")} ${o.par}`),
      ...study.options.map((o) => `${say(lang, "Ackumulerat före årsdagen", "Cumulative before the birthday")} ${o.par}`),
    ],
  ];
  for (const age of seriesAges(study)) {
    rows.push([
      age,
      ...study.options.map((o) => (age >= o.par ? kronorCell((o.annual[age] ?? 0) / 12) : "")),
      ...study.options.map((o) => (age >= o.par ? kronorCell(o.cumulative[age] ?? 0) : "")),
    ]);
  }
  return rows;
}
