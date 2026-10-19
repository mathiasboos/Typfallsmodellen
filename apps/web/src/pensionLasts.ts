/**
 * Hur länge räcker pensionen? How long does the pension last, for a budget?
 *
 * Another question the single forecast cannot answer. The public pension is paid
 * for life and nothing the person chooses changes that. The occupational pension
 * and any private pension saving are different: the person can often choose to
 * have them paid out over 5 to 20 years, or for life. A short payout is a larger
 * amount for a shorter time; a lifelong one is a smaller amount for as long as
 * the person lives. Whether that pays for the life one wants is a question of a
 * budget, so the tab sets the pension -- under each way of taking it out -- against
 * one.
 *
 * Everything here is pure -- no DOM -- so the arithmetic can be unit-tested
 * directly (`apps/web/test/pensionLasts.test.ts`), the same split
 * `pensionAge.ts` makes. What it reuses is `run()` itself, once per payout
 * period, through the two settings the workbook already has for this:
 * `rng_Temp_Tjp_Uttag` and `rng_Temp_IPS_Uttag`, "temporary withdrawal of the
 * occupational pension / of private saving, in years; 0 is lifelong". Seventeen
 * runs (every whole year from 5 to 20, and lifelong) take a few tens of
 * milliseconds, so there is no worker and no debounce, which is `main.ts`'s own
 * stated philosophy. The runs do not depend on the budget, the lifespan or
 * before/after tax, so a caller keeps them (`PayoutRuns`) while only those change.
 *
 * - **The person** is a typfall: birth year, when work started, the retirement
 *   age (public and occupational pension both start then), a monthly salary --
 *   or a final salary, found the way `pensionAge.ts` finds it -- and a collective
 *   agreement. The retirement age is moved inside what the cohort may draw.
 * - **The individual occupational pension and the private pension** -- the latter a
 *   pension insurance or an IPS, both -- are the model's private saving (IPS,
 *   `privateSavingKind: 0`, taxed as income when paid out, like an occupational pension):
 *   each a capital already saved and a monthly saving until retirement. The model treats them alike -- the same return, the same payout period,
 *   the same tax -- so they are one pot and what is saved in each adds up. They are
 *   paid out over the same period as the agreement's pension. The model has no place
 *   for the capital (its own `pbhYear` replaces every pot at once, and garantipension
 *   with it), so `ipsOpening` was added to the engine. Capital and saving count from
 *   the year Avancerat calls "Sparandet börjar år" -- the model's idea of today -- or
 *   from the retirement year if that has passed. The monthly saving is a fixed number
 *   of kronor, as in the model: it is not raised with inflation.
 * - **Private savings** are held in one of two kinds, as the money is taxed: an ISK,
 *   a kapitalförsäkring or a depåförsäkring, taxed on a standard return while saved; or
 *   a depå, a VP-konto or a fondkonto, taxed 30% on the gain when it is sold. Both are
 *   paid out tax-free, and the person types the capital after that tax on the second --
 *   what there is to live on -- while its return is taxed each year (the engine's
 *   `SavingType.Depot`, which the workbook does not have). The model carries one kind
 *   of private saving per run, so they are a second run (`lastsContext`'s
 *   `"savings"`), and the two payouts add -- what is tax-free is no part of what the
 *   tax is taken from. They are drawn over the same payout period as everything else;
 *   only "lifelong" differs, since money, unlike an IPS, is not paid for as long as the
 *   person lives: the model spreads it over the expected remaining life and it runs out.
 * - **The budget** is what a pensioner spends a month *after tax*, in the
 *   reference year's money -- minPension's three examples, the same as the site's
 *   Pensionsöversikt uses, or one's own. The pension is set against it at every
 *   age from the retirement age up to the lifespan (`lifespan - retirement`
 *   whole years: the lifespan is the age at which the person is gone, as in
 *   `pensionAge.ts`). Everything is in fixed prices (`priceBasis: 1`) whatever
 *   Avancerat says, since a budget is an amount of purchasing power. Benefits
 *   such as housing supplement are not counted; the pension after tax is the
 *   model's `netto`.
 * - **Before tax** the budget is what it takes to leave the budget after tax. That
 *   is not a rate applied to the budget: tax is progressive and the model's rules
 *   change with age, so it is the gross income the model's own tax section turns
 *   into the budget, age by age (`runWithTaxProbe`). Since tax never leaves less
 *   for more, a pension meets the budget before tax exactly when it meets it
 *   after, and the view only changes the units. A payout that is not taxed as income
 *   (the private savings) is grossed up the same way, to the gross income that leaves
 *   as much: added at face value it would count for no more than a taxed krona, and
 *   the two views would part.
 *
 * **Two readings of "lasts".** Month by month: an age is *covered* if the pension
 * then is at least the budget. A 5-year payout covers the budget with room to
 * spare for five years and then leaves only the public pension, which may fall
 * short. Added up: if what is left over in the good years is saved (without
 * interest), the money lasts until the running total of pension minus budget
 * turns negative. The first is the plain answer to "can I live on it each month";
 * the second says whether a short payout, saved with discipline, still gets the
 * person through. Both are reported and neither is hidden in the other. The
 * running total is always after tax -- only that is left to save -- whichever
 * basis the rest is shown in.
 *
 * **The recommended payout** is the one that covers the budget in the most years;
 * a tie goes to the longer payout, lifelong being the longest. Shorter and
 * shorter payouts pay more and more for fewer years, so the count rises up to the
 * longest payout that still reaches the budget and then falls. When no payout
 * reaches the budget in any year there is none to recommend.
 *
 * **The model's tax steps.** The higher grundavdrag for the old begins at an age
 * of its own (the riktålder plus one, for the cohorts to come), so the years
 * before it are taxed harder and leave less. That is the model's own rule, the one
 * Table 2 shows too, and it is not smoothed away: an age that falls short only
 * for that reason is reported as short. Reported as a stretch, though --
 * "från 69 år" -- not as a failure of the whole.
 */
import { Scheme, Table1Key, defaultInput, riktalderFor, run, runWithTaxProbe } from "@typfallsmodellen/engine";
import type { DeathProbabilities, ModelContext, MvaluesRow, SchemeId, TypfallInput } from "@typfallsmodellen/engine";

import type { Lang } from "./i18n.js";
import { retirementAge } from "./kpis.js";
import { LAST_AGE, LIFESPAN_BOUNDS, monthlySalaryForFinalSalary, referenceAge } from "./pensionAge.js";
import { defaultSavings, savingSettings, savingYear, taxFreeSaving } from "./privateSavings.js";
import type { SavingPart, SavingsInput } from "./privateSavings.js";
import { csvLine, delimiterFor } from "./tables.js";
import { kronorCell } from "./xlsx.js";
import type { SheetData } from "./xlsx.js";

const say = (l: Lang, sv: string, en: string) => (l === "sv" ? sv : en);

// --------------------------------------------------------------- payouts ---

/** The payout period 0: for life, as the workbook writes "Livsvarigt". */
export const LIFELONG = 0;

/** The shortest and the longest fixed payout the agreements offer. */
export const PAYOUT_BOUNDS = { min: 5, max: 20 } as const;

/** Every whole year from the shortest payout to the longest. */
export const PAYOUT_YEARS: readonly number[] = Array.from(
  { length: PAYOUT_BOUNDS.max - PAYOUT_BOUNDS.min + 1 },
  (_, i) => PAYOUT_BOUNDS.min + i,
);

/** What the tab compares, shortest payout first and lifelong last. */
export const PAYOUTS: readonly number[] = [...PAYOUT_YEARS, LIFELONG];

/** The periods the charts draw: more lines than these cannot be told apart. */
export const CHART_PAYOUTS: readonly number[] = [5, 10, 15, 20, LIFELONG];

/** One per drawn period; the table rows of the other years have no colour. */
export type PayoutSlot = "y5" | "y10" | "y15" | "y20" | "life";

export function slotOfPayout(payout: number): PayoutSlot | undefined {
  switch (payout) {
    case 5:
      return "y5";
    case 10:
      return "y10";
    case 15:
      return "y15";
    case 20:
      return "y20";
    case LIFELONG:
      return "life";
    default:
      return undefined;
  }
}

/** "10 år", "Livsvarigt". */
export function payoutName(payout: number, lang: Lang): string {
  if (payout === LIFELONG) return say(lang, "Livsvarigt", "Lifelong");
  return `${payout} ${say(lang, "år", "years")}`;
}

// --------------------------------------------------------------- budgets ---

/**
 * minPension's examples of what a pensioner spends a month after tax ("Hur mycket
 * pengar behöver en pensionär?"), in today's money. The same three amounts as the
 * site's Pensionsöversikt uses; if they change there, they change here.
 */
export const BUDGETS = [
  { id: "bas", sv: "Baspensionär", en: "Basic pensioner", amount: 17_600 },
  { id: "medel", sv: "Medelpensionär", en: "Average pensioner", amount: 18_400 },
  { id: "plus", sv: "Pluspensionär", en: "Plus pensioner", amount: 27_300 },
] as const;

export type BudgetId = (typeof BUDGETS)[number]["id"];

/** The example whose amount is exactly `amount`, if any. */
export function budgetIdOf(amount: number): BudgetId | undefined {
  return BUDGETS.find((b) => b.amount === amount)?.id;
}

/** What a budget a month can be typed as, kronor. */
export const BUDGET_BOUNDS = { min: 0, max: 200_000, step: 100 } as const;

// ----------------------------------------------------------------- input ---

export interface LastsInput extends SavingsInput {
  readonly born: number;
  readonly startWorkAge: number;
  /** The age the public and the occupational pension start at. */
  readonly retirementAge: number;
  /** The age the person is gone at: pension is set against the budget up to the year before. */
  readonly lifespan: number;
  /** Per month, in the reference year's money -- `TypfallInput.monthlySalary`. */
  readonly monthlySalary: number;
  readonly scheme: SchemeId;
  /** A month after tax, in the reference year's money. */
  readonly budget: number;
  readonly yearlyInflation: number;
  readonly realGrowth: number;
  readonly realReturn: number;
  /** Show the pension and the budget after tax instead of before. */
  readonly afterTax: boolean;
}

/**
 * A typfall born 1980 who started working at 25, earns the model's normal salary,
 * has ITP 1, retires at the riktålder, wants to last to 95 on what minPension calls
 * an average pensioner's budget, has no private pension or private saving, and takes
 * the workbook's own economic assumptions.
 */
export function defaultLastsInput(): LastsInput {
  const base = defaultInput();
  const born = 1980;
  return {
    born,
    startWorkAge: 25,
    retirementAge: riktalderFor(born),
    lifespan: 95,
    monthlySalary: base.monthlySalary,
    scheme: Scheme.Itp1,
    budget: BUDGETS[1].amount,
    ...defaultSavings(),
    yearlyInflation: base.yearlyInflation,
    realGrowth: base.realGrowth,
    realReturn: base.realReturn,
    afterTax: true,
  };
}

/** The retirement age moved inside what the cohort may draw, and the lifespan kept after it. */
export function normaliseLasts(input: LastsInput): LastsInput {
  const retirement = referenceAge(input.born, input.retirementAge);
  const lifespan = Math.min(
    Math.max(Math.round(input.lifespan), LIFESPAN_BOUNDS.min, retirement + 1),
    LIFESPAN_BOUNDS.max,
  );
  return { ...input, retirementAge: retirement, lifespan };
}

export function lastsTypfall(input: LastsInput): TypfallInput {
  return defaultInput({
    born: input.born,
    startWorkAge: input.startWorkAge,
    monthlySalary: input.monthlySalary,
    retirementAge: input.retirementAge,
    scheme: input.scheme,
    yearlyInflation: input.yearlyInflation,
    realGrowth: input.realGrowth,
    realReturn: input.realReturn,
  });
}

/**
 * The shared context with what the question needs: the payout period for both
 * pensions, the occupational pension starting with the public one, fixed prices,
 * and the model's private saving carrying one of two things (`savingSettings`):
 * `"pension"`, everything paid out as taxed income, or `"savings"`, the tax-free
 * private savings, which are a run of their own.
 */
export function lastsContext(
  shared: ModelContext,
  input: LastsInput,
  payout: number,
  part: SavingPart = "pension",
): ModelContext {
  return {
    ...shared,
    tjpPar: 0,
    priceBasis: 1,
    tempTjpUttag: payout,
    tempIpsUttag: payout,
    ...savingSettings(shared, input, input, part),
  };
}

// ------------------------------------------------------------------ runs ---

/** One payout period's pension, kronor a year in the reference year's money, by age 0..`LAST_AGE`; zero before retirement. */
export interface PayoutRun {
  /** 0 is lifelong. */
  readonly payout: number;
  readonly publicPension: readonly number[];
  readonly occupational: readonly number[];
  /** Individual occupational pension and private pension (pension insurance, IPS): the model's IPS, paid out as taxed income. */
  readonly individual: readonly number[];
  /** The three together, before tax: all that the model takes tax from. */
  readonly gross: readonly number[];
  /** The private savings paid out. Not taxed as income, so the same before and after tax. */
  readonly savings: readonly number[];
  /** What is left of `gross` after the model's tax, with the tax-free `savings` on top. */
  readonly net: readonly number[];
}

/**
 * Everything the engine is run for -- independent of the budget, the lifespan and
 * before/after tax, so a caller keeps it while only those change.
 */
export interface PayoutRuns {
  readonly input: LastsInput;
  /** In `PAYOUTS` order. */
  readonly runs: readonly PayoutRun[];
  /** What is left of `gross` kronor a year at `age` after the model's tax, and the gross that leaves `net`. */
  readonly netAt: (age: number, gross: number) => number;
  readonly grossFor: (age: number, net: number) => number;
  readonly referenceYear: number;
  /** The year the capital and the saving count from. */
  readonly savingYear: number;
  /** Table 1's "Slutlön" per month -- what a typed final salary is matched against. */
  readonly finalSalary: number;
  /** The distinct warnings the runs raised, in the order met. */
  readonly warnings: readonly string[];
}

/** What a set of runs is valid for: the person, the pensions and the assumptions -- not the budget, the lifespan or the tax view. */
export function runsKey(input: LastsInput): string {
  const { budget: _budget, lifespan: _lifespan, afterTax: _afterTax, ...rest } = input;
  return JSON.stringify(rest);
}

function annual(rows: readonly MvaluesRow[], par: number, pick: (row: MvaluesRow) => number): number[] {
  const out = new Array<number>(LAST_AGE + 1).fill(0);
  for (const row of rows) {
    if (row.age >= par && row.age <= LAST_AGE) out[row.age] = pick(row);
  }
  return out;
}

/** The seventeen runs. `input` should already be `normaliseLasts`'d. */
export function runPayouts(input: LastsInput, shared: ModelContext, deaths: DeathProbabilities): PayoutRuns {
  const typfall = lastsTypfall(input);
  const runs: PayoutRun[] = [];
  const warnings: string[] = [];
  let netAt: PayoutRuns["netAt"] = () => 0;
  let grossFor: PayoutRuns["grossFor"] = () => 0;
  let finalSalary = 0;
  const taxFree = taxFreeSaving(input);
  const remember = (messages: readonly { readonly message: string }[]): void => {
    for (const { message } of messages) if (!warnings.includes(message)) warnings.push(message);
  };

  for (const payout of PAYOUTS) {
    const probe = runWithTaxProbe(typfall, lastsContext(shared, input, payout), { deaths });
    const { rows } = probe.result;
    const par = retirementAge(typfall, probe.result);

    const publicPension = annual(rows, par, (r) => r.ip + r.tp + r.pp + r.garp + r.ptillagg);
    const occupational = annual(rows, par, (r) => r.tjp);
    const individual = annual(rows, par, (r) => r.ips);
    remember(probe.result.warnings);

    // The private savings are a second run: the model carries one kind of private saving at a time.
    let savings: readonly number[] = new Array<number>(LAST_AGE + 1).fill(0);
    if (taxFree !== null) {
      const own = run(typfall, lastsContext(shared, input, payout, "savings"), { deaths });
      savings = annual(own.rows, retirementAge(typfall, own), (r) => r.pps);
      remember(own.warnings);
    }

    runs.push({
      payout,
      publicPension,
      occupational,
      individual,
      gross: publicPension.map((p, age) => p + occupational[age]! + individual[age]!),
      savings,
      net: annual(rows, par, (r) => r.netto).map((n, age) => n + savings[age]!),
    });
    // The tax section and the salary are the same whatever the payout, so the last run's are kept.
    netAt = probe.netAt;
    grossFor = probe.grossFor;
    finalSalary = probe.result.table1.find((r) => r.key === Table1Key.FinalSalary)?.monthly ?? 0;
  }

  return {
    input,
    runs,
    netAt,
    grossFor,
    referenceYear: shared.referenceYear,
    savingYear: savingYear(input, shared),
    finalSalary,
    warnings,
  };
}

// ------------------------------------------------------------ assessment ---

/** A shortfall or surplus this small, kronor a year, is rounding: half a krona a month. */
export const TOLERANCE = 6;

/** Some of the years: the payout itself, or what follows it. */
export interface Phase {
  /** First age, and the age after the last. */
  readonly from: number;
  readonly to: number;
  readonly years: number;
  /** Kronor a month, averaged over the years. */
  readonly income: number;
  readonly need: number;
  /** `income - need`. */
  readonly difference: number;
  /** The lowest year's income, a month: the most a budget can be and still be met every year. */
  readonly lowest: number;
  /** Every year meets the budget. */
  readonly covered: boolean;
}

/** Ages from `from` up to, not including, `to`. */
export interface Stretch {
  readonly from: number;
  readonly to: number;
}

/** How the pension of one payout period meets one budget, age by age. */
export interface Assessment {
  readonly payout: number;
  /** The retirement age, and the lifespan. */
  readonly from: number;
  readonly to: number;
  /** Kronor a year at each age from `from` up to the year before `to`, in the basis shown. */
  readonly income: readonly number[];
  readonly need: readonly number[];
  readonly covered: readonly boolean[];
  readonly coveredYears: number;
  /** The unbroken runs of covered ages, in order. */
  readonly stretches: readonly Stretch[];
  /** The first and the last age that meets the budget; null if none does. */
  readonly firstCovered: number | null;
  readonly lastCovered: number | null;
  /**
   * The age after the last covered one -- where the last stretch that meets the budget
   * ends, however many earlier ones there were; null if the budget is met up to the
   * lifespan, or never.
   */
  readonly runsOutAt: number | null;
  /** Every age from the retirement age to the lifespan meets the budget. */
  readonly lasts: boolean;
  /** Kronor, the years' shortfalls and surpluses added up. */
  readonly shortfall: number;
  readonly surplus: number;
  /** The running total of income minus budget at the end of each age. */
  readonly balance: readonly number[];
  /** The age, to a tenth, at which that total turns negative; null if it never does. */
  readonly balanceRunsOutAt: number | null;
  /** The years of the payout, and the years after it; null if there are none. */
  readonly during: Phase | null;
  readonly after: Phase | null;
}

function phaseOf(income: readonly number[], need: readonly number[], from: number, first: number, last: number): Phase | null {
  if (last <= first) return null;
  let sumIncome = 0;
  let sumNeed = 0;
  let lowest = Number.POSITIVE_INFINITY;
  let covered = true;
  for (let age = first; age < last; age += 1) {
    const i = age - from;
    sumIncome += income[i]!;
    sumNeed += need[i]!;
    lowest = Math.min(lowest, income[i]!);
    if (income[i]! - need[i]! < -TOLERANCE) covered = false;
  }
  const years = last - first;
  return {
    from: first,
    to: last,
    years,
    income: sumIncome / years / 12,
    need: sumNeed / years / 12,
    difference: (sumIncome - sumNeed) / years / 12,
    lowest: lowest / 12,
    covered,
  };
}

/**
 * Sets one payout period's income against the need, a year at a time.
 *
 * `income` is indexed by age 0..`LAST_AGE`, `need` is asked for the ages assessed:
 * `from` up to the year before `to`. The payout lasts `payout` whole years from
 * `from` -- the model pays the last of them in full and none after.
 */
export function assess(
  payout: number,
  income: readonly number[],
  need: (age: number) => number,
  from: number,
  to: number,
): Assessment {
  const yearly: number[] = [];
  const wanted: number[] = [];
  for (let age = from; age < to; age += 1) {
    yearly.push(income[age] ?? 0);
    wanted.push(need(age));
  }
  const covered = yearly.map((v, i) => v - wanted[i]! >= -TOLERANCE);

  const stretches: Stretch[] = [];
  covered.forEach((is, i) => {
    if (!is) return;
    const last = stretches[stretches.length - 1];
    if (last !== undefined && last.to === from + i) stretches[stretches.length - 1] = { from: last.from, to: from + i + 1 };
    else stretches.push({ from: from + i, to: from + i + 1 });
  });
  const firstCovered = stretches[0]?.from ?? null;
  const lastCovered = stretches.length === 0 ? null : stretches[stretches.length - 1]!.to - 1;

  let shortfall = 0;
  let surplus = 0;
  const balance: number[] = [];
  let running = 0;
  yearly.forEach((v, i) => {
    const difference = v - wanted[i]!;
    if (difference < 0) shortfall -= difference;
    else surplus += difference;
    running += difference;
    balance.push(running);
  });

  // The money is gone when the running total turns negative for the last time. A
  // total that dips below zero in the first years (the model taxes the years
  // before the higher grundavdrag harder) and then climbs is not "gone" in year
  // one -- the dip is a shortfall, reported as one -- but one that never reaches
  // zero is gone at once.
  let balanceRunsOutAt: number | null = null;
  const lastIndex = balance.length - 1;
  let positive = lastIndex;
  while (positive >= 0 && balance[positive]! < -TOLERANCE) positive -= 1;
  if (lastIndex >= 0 && positive < lastIndex) {
    if (positive < 0) {
      balanceRunsOutAt = from;
    } else {
      // Linear inside the year, as `breakEven` does: the running total is.
      const before = balance[positive]!;
      const after = balance[positive + 1]!;
      const share = before > 0 ? before / (before - after) : 0;
      balanceRunsOutAt = from + positive + 1 + Math.min(Math.max(share, 0), 1);
    }
  }

  const payoutEnd = payout === LIFELONG ? to : Math.min(from + payout, to);
  return {
    payout,
    from,
    to,
    income: yearly,
    need: wanted,
    covered,
    coveredYears: covered.filter(Boolean).length,
    stretches,
    firstCovered,
    lastCovered,
    runsOutAt: lastCovered === null || lastCovered === to - 1 ? null : lastCovered + 1,
    lasts: covered.length > 0 && covered.every(Boolean),
    shortfall,
    surplus,
    balance,
    balanceRunsOutAt,
    during: phaseOf(yearly, wanted, from, from, payoutEnd),
    after: payout === LIFELONG ? null : phaseOf(yearly, wanted, from, payoutEnd, to),
  };
}

// --------------------------------------------------------------- results ---

export interface PayoutOption {
  readonly payout: number;
  readonly run: PayoutRun;
  /** In the basis shown: after tax or before, against the budget in the same terms. */
  readonly view: Assessment;
  /**
   * The same after tax, always. What is left over is only worth saving after tax, so
   * the running total -- "the money lasts if the surplus is saved" -- is read here,
   * not from `view`. It is `view` itself when the basis shown is after tax.
   */
  readonly net: Assessment;
  /**
   * The private savings payout in the basis shown, by age 0..`LAST_AGE`. After
   * tax it is what is paid out. Before tax it is grossed up, so that the parts of the
   * pension still add up to `view.income` -- see `evaluate`. All zero without one.
   */
  readonly savings: readonly number[];
}

export interface Lasts {
  readonly input: LastsInput;
  readonly runs: PayoutRuns;
  /** The retirement age, and the lifespan. */
  readonly from: number;
  readonly to: number;
  /** In `PAYOUTS` order. */
  readonly options: readonly PayoutOption[];
  /** The public pension alone, for the years from the retirement age: what is left when a payout ends. */
  readonly publicOnly: Assessment;
  /** What the budget is in the basis shown, a month, at each age from `from` up to `to`. */
  readonly budgetMonthly: readonly number[];
  /**
   * The budget a month before tax: at the first age, at the last, and the middle
   * one -- the first is higher when the first years are taxed harder.
   */
  readonly budgetGross: { readonly first: number; readonly last: number; readonly typical: number };
  /**
   * The payout period that covers the budget in the most years -- the longer
   * payout on a tie. Null when no payout covers it in any year.
   */
  readonly best: PayoutOption | null;
  /** Whether the person has an occupational or individual pension, or private saving, for a payout period to apply to. */
  readonly flexible: boolean;
  /** The private savings pay out something, which is not taxed as income. */
  readonly taxFree: boolean;
}

/** Whether a run holds any occupational pension or private saving at all. */
function hasFlexible(run: PayoutRun): boolean {
  return run.occupational.some((v) => v > 0) || run.individual.some((v) => v > 0) || run.savings.some((v) => v > 0);
}

/** Sets the runs against the input's budget, lifespan and tax view. */
export function evaluate(runs: PayoutRuns, input: LastsInput): Lasts {
  const from = input.retirementAge;
  const to = input.lifespan;
  const yearlyBudget = input.budget * 12;

  // Before tax the budget is the gross that leaves it after tax -- see the file comment.
  const grossBudget = new Array<number>(LAST_AGE + 1).fill(0);
  for (let age = from; age < to; age += 1) grossBudget[age] = runs.grossFor(age, yearlyBudget);
  const need = (age: number) => (input.afterTax ? yearlyBudget : (grossBudget[age] ?? 0));
  const publicGross = runs.runs[0]!.publicPension;

  const options = runs.runs.map((run): PayoutOption => {
    const net = assess(run.payout, run.net, () => yearlyBudget, from, to);
    // Before tax, a payout that is not taxed as income is what the same money would have
    // been as taxed income: the gross income the model's tax turns into the same net.
    // Only that keeps this view meeting the budget in exactly the years the other does --
    // added at face value, an untaxed krona would count for no more than a taxed one.
    const grossIncome = input.afterTax
      ? run.net
      : run.gross.map((g, age) => (age >= from && age < to && run.savings[age]! > 0 ? runs.grossFor(age, run.net[age]!) : g));
    return {
      payout: run.payout,
      run,
      view: input.afterTax ? net : assess(run.payout, grossIncome, need, from, to),
      net,
      savings: input.afterTax
        ? run.savings
        : grossIncome.map((g, age) => (run.savings[age]! > 0 ? Math.max(g - run.gross[age]!, 0) : 0)),
    };
  });

  let best: PayoutOption | null = null;
  for (const option of options) {
    // Shortest first, so `>=` lets a longer payout win a tie; lifelong comes last.
    if (option.view.coveredYears > 0 && (best === null || option.view.coveredYears >= best.view.coveredYears)) best = option;
  }

  const publicIncome = publicGross.map((v, age) => (input.afterTax ? runs.netAt(age, v) : v));
  const grossMonthly = Array.from({ length: to - from }, (_, i) => (grossBudget[from + i] ?? 0) / 12);
  const middle = [...grossMonthly].sort((a, b) => a - b)[Math.floor(grossMonthly.length / 2)] ?? 0;
  return {
    input,
    runs,
    from,
    to,
    options,
    publicOnly: assess(LIFELONG, publicIncome, need, from, to),
    budgetMonthly: Array.from({ length: to - from }, (_, i) => need(from + i) / 12),
    budgetGross: { first: grossMonthly[0] ?? 0, last: grossMonthly[grossMonthly.length - 1] ?? 0, typical: middle },
    best,
    flexible: runs.runs.some(hasFlexible),
    taxFree: runs.runs.some((r) => r.savings.some((v) => v > 0)),
  };
}

/** `runPayouts` and `evaluate` in one, for a caller that has no reason to keep the runs. */
export function computeLasts(input: LastsInput, shared: ModelContext, deaths: DeathProbabilities): Lasts {
  const normal = normaliseLasts(input);
  return evaluate(runPayouts(normal, shared, deaths), normal);
}

// ---------------------------------------------------------------- matrix ---

export interface MatrixRow {
  readonly payout: number;
  /** One per budget asked for, after tax. */
  readonly cells: readonly Assessment[];
  /** The most a budget a month after tax can be and still be met in every year of the payout. */
  readonly highest: number;
}

/**
 * Every payout period against several budgets, after tax -- how long each
 * budget lasts under each way of taking the pension out. After tax because the
 * budgets are: the answer is the same before tax, only the units differ.
 */
export function budgetMatrix(lasts: Lasts, budgets: readonly number[]): MatrixRow[] {
  return lasts.options.map((option) => {
    const cells = budgets.map((budget) => assess(option.payout, option.run.net, () => budget * 12, lasts.from, lasts.to));
    const during = assess(option.payout, option.run.net, () => 0, lasts.from, lasts.to).during;
    return { payout: option.payout, cells, highest: during?.lowest ?? 0 };
  });
}

// -------------------------------------------------------------- salaries ---

/** Table 1's "Slutlön" per month for the person -- what a typed final salary is matched against. */
export function finalSalaryOf(input: LastsInput, shared: ModelContext, deaths: DeathProbabilities): number {
  const result = run(lastsTypfall(input), lastsContext(shared, input, LIFELONG), { deaths });
  return result.table1.find((r) => r.key === Table1Key.FinalSalary)?.monthly ?? 0;
}

/** The monthly salary whose final salary is `target`, found by running the model -- see `monthlySalaryForFinalSalary`. */
export function monthlySalaryForTarget(
  input: LastsInput,
  target: number,
  shared: ModelContext,
  deaths: DeathProbabilities,
): number {
  return monthlySalaryForFinalSalary(target, (monthlySalary) => finalSalaryOf({ ...input, monthlySalary }, shared, deaths));
}

// ----------------------------------------------------------------- export ---

const whole = (n: number): string => String(Math.round(n));

function basisText(input: LastsInput, lang: Lang): string {
  return input.afterTax ? say(lang, "efter skatt", "after tax") : say(lang, "före skatt", "before tax");
}

/**
 * Where a budget stands under a payout, in words that fit a table cell and a CSV. `short`
 * leaves the lifespan out of "all the way", for where it has just been said.
 */
export function coverageText(a: Assessment, lang: Lang, short = false): string {
  if (a.lasts) return short ? say(lang, "Hela vägen", "All the way") : say(lang, `Hela vägen till ${a.to} år`, `All the way to ${a.to}`);
  if (a.firstCovered === null) return say(lang, "Räcker inte", "Does not last");
  if (a.stretches.length > 1) {
    return say(lang, `${a.coveredYears} av ${a.to - a.from} år`, `${a.coveredYears} of ${a.to - a.from} years`);
  }
  const first = a.firstCovered;
  const early = first > a.from;
  if (a.runsOutAt === null) return say(lang, `Från ${first} år`, `From ${first}`);
  return early ? `${first}–${a.runsOutAt} ${say(lang, "år", "yrs")}` : say(lang, `Till ${a.runsOutAt} år`, `To ${a.runsOutAt}`);
}

/** When the running total turns negative: a tenth of a year, or the whole way. */
export function balanceText(a: Assessment, lang: Lang): string {
  if (a.balanceRunsOutAt === null) return say(lang, `Hela vägen till ${a.to} år`, `All the way to ${a.to}`);
  if (a.balanceRunsOutAt <= a.from) return say(lang, "Räcker inte", "Does not last");
  const text = new Intl.NumberFormat(lang === "sv" ? "sv-SE" : "en-GB", { minimumFractionDigits: 1, maximumFractionDigits: 1 }).format(
    a.balanceRunsOutAt,
  );
  return say(lang, `Till ${text} år`, `To ${text}`);
}

function monthlyOf(lasts: Lasts, option: PayoutOption, age: number): number {
  return option.view.income[age - lasts.from]! / 12;
}

/** The first header row and the cells for one payout, in the table's own order. */
function optionsHeader(lasts: Lasts, lang: Lang): string[] {
  const tax = basisText(lasts.input, lang);
  return [
    say(lang, "Utbetalningstid", "Payout period"),
    say(lang, `Under utbetalningen, kr/mån ${tax}`, `During the payout, SEK/month ${tax}`),
    ...(lasts.input.afterTax
      ? []
      : [
          say(lang, "varav allmän pension", "of which public pension"),
          say(lang, "varav tjänstepension", "of which occupational pension"),
          say(lang, "varav individuell och privat pension", "of which individual and private pension"),
          ...(lasts.taxFree
            ? [say(lang, "varav privat sparande, uppräknat", "of which private savings, grossed up")]
            : []),
        ]),
    say(lang, `Efter utbetalningen, kr/mån ${tax}`, `After the payout, SEK/month ${tax}`),
    say(lang, "Mot budgeten under utbetalningen, kr/mån", "Against the budget during the payout, SEK/month"),
    say(lang, "Mot budgeten efter utbetalningen, kr/mån", "Against the budget after the payout, SEK/month"),
    say(lang, "Budgeten räcker", "The budget lasts"),
    say(lang, "År som budgeten täcks", "Years the budget is met"),
    say(lang, "Saknas totalt, kr", "Short in total, SEK"),
    say(lang, "Pengarna räcker om överskottet sparas (efter skatt)", "The money lasts if the surplus is saved (after tax)"),
  ];
}

function optionRow(lasts: Lasts, option: PayoutOption, lang: Lang): (string | number)[] {
  const { view } = option;
  const { during, after } = view;
  const afterTax = lasts.input.afterTax;
  /** A series' average a month over the years of the payout. */
  const average = (series: readonly number[]): number => {
    if (during === null) return 0;
    let sum = 0;
    for (let age = during.from; age < during.to; age += 1) sum += series[age] ?? 0;
    return sum / during.years / 12;
  };
  return [
    payoutName(option.payout, lang),
    during?.income ?? 0,
    ...(afterTax
      ? []
      : [
          average(option.run.publicPension),
          average(option.run.occupational),
          average(option.run.individual),
          ...(lasts.taxFree ? [average(option.savings)] : []),
        ]),
    after === null ? "" : after.income,
    during?.difference ?? 0,
    after === null ? "" : after.difference,
    coverageText(view, lang),
    view.coveredYears,
    view.shortfall,
    balanceText(option.net, lang),
  ];
}

export function optionsToCsv(lasts: Lasts, lang: Lang): string {
  const delimiter = delimiterFor(lang);
  const lines = [csvLine(optionsHeader(lasts, lang), delimiter)];
  for (const option of lasts.options) {
    lines.push(csvLine(optionRow(lasts, option, lang).map((cell) => (typeof cell === "number" ? whole(cell) : cell)), delimiter));
  }
  return lines.join("\r\n");
}

export function optionsToXlsxRows(lasts: Lasts, lang: Lang): SheetData {
  const rows: SheetData = [optionsHeader(lasts, lang)];
  for (const option of lasts.options) {
    rows.push(optionRow(lasts, option, lang).map((cell) => (typeof cell === "number" ? kronorCell(cell) : cell)));
  }
  return rows;
}

/** "Hur mycket får jag?": what each payout pays a month, during it and after it, and in all to the lifespan. */
function amountsHeader(lasts: Lasts, lang: Lang): string[] {
  const tax = basisText(lasts.input, lang);
  return [
    say(lang, "Utbetalningstid", "Payout period"),
    say(lang, `Under utbetalningen, kr/mån ${tax}`, `During the payout, SEK/month ${tax}`),
    ...(lasts.input.afterTax
      ? []
      : [
          say(lang, "varav allmän pension", "of which public pension"),
          say(lang, "varav tjänstepension", "of which occupational pension"),
          say(lang, "varav individuell och privat pension", "of which individual and private pension"),
          ...(lasts.taxFree
            ? [say(lang, "varav privat sparande, uppräknat", "of which private savings, grossed up")]
            : []),
        ]),
    say(lang, `Efter utbetalningen, kr/mån ${tax}`, `After the payout, SEK/month ${tax}`),
    say(lang, `Totalt till ${lasts.to} år, kr ${tax}`, `In all to ${lasts.to}, SEK ${tax}`),
  ];
}

/** The total over the years from the retirement age to the lifespan, in the basis shown. */
export const totalPaid = (a: Assessment): number => a.income.reduce((sum, value) => sum + value, 0);

function amountRow(lasts: Lasts, option: PayoutOption, lang: Lang): (string | number)[] {
  const { view } = option;
  const { during, after } = view;
  const average = (series: readonly number[]): number => {
    if (during === null) return 0;
    let sum = 0;
    for (let age = during.from; age < during.to; age += 1) sum += series[age] ?? 0;
    return sum / during.years / 12;
  };
  return [
    payoutName(option.payout, lang),
    during?.income ?? 0,
    ...(lasts.input.afterTax
      ? []
      : [
          average(option.run.publicPension),
          average(option.run.occupational),
          average(option.run.individual),
          ...(lasts.taxFree ? [average(option.savings)] : []),
        ]),
    after === null ? "" : after.income,
    totalPaid(view),
  ];
}

/** The payout periods the tab shows, one row each. */
const amountOptions = (lasts: Lasts): PayoutOption[] => CHART_PAYOUTS.map((p) => lasts.options.find((o) => o.payout === p)!);

export function amountsToCsv(lasts: Lasts, lang: Lang): string {
  const delimiter = delimiterFor(lang);
  const lines = [csvLine(amountsHeader(lasts, lang), delimiter)];
  for (const option of amountOptions(lasts)) {
    lines.push(csvLine(amountRow(lasts, option, lang).map((cell) => (typeof cell === "number" ? whole(cell) : cell)), delimiter));
  }
  return lines.join("\r\n");
}

export function amountsToXlsxRows(lasts: Lasts, lang: Lang): SheetData {
  const rows: SheetData = [amountsHeader(lasts, lang)];
  for (const option of amountOptions(lasts)) {
    rows.push(amountRow(lasts, option, lang).map((cell) => (typeof cell === "number" ? kronorCell(cell) : cell)));
  }
  return rows;
}

function amountSeriesHeader(lasts: Lasts, lang: Lang): string[] {
  const unit = say(lang, "kr/mån", "SEK/month");
  const tax = basisText(lasts.input, lang);
  return [say(lang, "Ålder", "Age"), ...amountOptions(lasts).map((o) => `${payoutName(o.payout, lang)} ${tax}, ${unit}`)];
}

/** Per age: each payout's pension a month. */
export function amountSeriesToCsv(lasts: Lasts, lang: Lang): string {
  const delimiter = delimiterFor(lang);
  const lines = [csvLine(amountSeriesHeader(lasts, lang), delimiter)];
  for (let age = lasts.from; age < lasts.to; age += 1) {
    lines.push(csvLine([String(age), ...amountOptions(lasts).map((o) => whole(monthlyOf(lasts, o, age)))], delimiter));
  }
  return lines.join("\r\n");
}

export function amountSeriesToXlsxRows(lasts: Lasts, lang: Lang): SheetData {
  const rows: SheetData = [amountSeriesHeader(lasts, lang)];
  for (let age = lasts.from; age < lasts.to; age += 1) rows.push([age, ...amountOptions(lasts).map((o) => kronorCell(monthlyOf(lasts, o, age)))]);
  return rows;
}

/** The ages the series exports cover: the retirement age up to the year before the lifespan. */
function seriesAges(lasts: Lasts): number[] {
  const ages: number[] = [];
  for (let age = lasts.from; age < lasts.to; age += 1) ages.push(age);
  return ages;
}

function seriesHeader(lasts: Lasts, lang: Lang): string[] {
  const unit = say(lang, "kr/mån", "SEK/month");
  const tax = basisText(lasts.input, lang);
  return [
    say(lang, "Ålder", "Age"),
    `${say(lang, "Budget", "Budget")} ${tax}, ${unit}`,
    ...lasts.options.map((o) => `${payoutName(o.payout, lang)} ${tax}, ${unit}`),
    ...lasts.options.map(
      (o) =>
        `${payoutName(o.payout, lang)}: ${say(lang, "ackumulerat mot budgeten, efter skatt", "running total against the budget, after tax")}, ${say(lang, "kr", "SEK")}`,
    ),
  ];
}

function seriesCells(lasts: Lasts, age: number): number[] {
  const i = age - lasts.from;
  return [
    lasts.budgetMonthly[i] ?? 0,
    ...lasts.options.map((o) => monthlyOf(lasts, o, age)),
    ...lasts.options.map((o) => o.net.balance[i] ?? 0),
  ];
}

/** Per age: the budget, each payout's pension a month, and each payout's running total against the budget. */
export function seriesToCsv(lasts: Lasts, lang: Lang): string {
  const delimiter = delimiterFor(lang);
  const lines = [csvLine(seriesHeader(lasts, lang), delimiter)];
  for (const age of seriesAges(lasts)) lines.push(csvLine([String(age), ...seriesCells(lasts, age).map(whole)], delimiter));
  return lines.join("\r\n");
}

export function seriesToXlsxRows(lasts: Lasts, lang: Lang): SheetData {
  const rows: SheetData = [seriesHeader(lasts, lang)];
  for (const age of seriesAges(lasts)) rows.push([age, ...seriesCells(lasts, age).map(kronorCell)]);
  return rows;
}

function matrixHeader(budgets: readonly { label: string; amount: number }[], lang: Lang): string[] {
  return [
    say(lang, "Utbetalningstid", "Payout period"),
    ...budgets.map((b) => `${b.label} (${whole(b.amount)} ${say(lang, "kr/mån efter skatt", "SEK/month after tax")})`),
    say(lang, "Högsta budget som klaras varje år, kr/mån efter skatt", "Highest budget met every year, SEK/month after tax"),
  ];
}

export function matrixToCsv(rows: readonly MatrixRow[], budgets: readonly { label: string; amount: number }[], lang: Lang): string {
  const delimiter = delimiterFor(lang);
  const lines = [csvLine(matrixHeader(budgets, lang), delimiter)];
  for (const row of rows) {
    lines.push(csvLine([payoutName(row.payout, lang), ...row.cells.map((c) => coverageText(c, lang)), whole(row.highest)], delimiter));
  }
  return lines.join("\r\n");
}

export function matrixToXlsxRows(rows: readonly MatrixRow[], budgets: readonly { label: string; amount: number }[], lang: Lang): SheetData {
  const sheet: SheetData = [matrixHeader(budgets, lang)];
  for (const row of rows) {
    sheet.push([payoutName(row.payout, lang), ...row.cells.map((c) => coverageText(c, lang)), kronorCell(row.highest)]);
  }
  return sheet;
}
