/**
 * Uttagsstrategi: in which order, when and from which pot -- the question the other tabs leave open.
 *
 * "Hur länge räcker pensionen?" and "Hur mycket får jag?" ask how much a pension pays and for how
 * long; this tab asks how to take it. For a typfall of its own (as those two tabs, with "Hämta från
 * Prognos"), it runs the model for a handful of ways of drawing the pots -- the occupational pension,
 * the individual and private pension (the model's IPS) and the private savings (an ISK, a KF or a
 * depot) -- each with its own payout period and, for the taxed pots, its own start, and reads three
 * things off the runs:
 *
 * - **What most people miss** (`insights`): effects the model computes but no other tab shows, each
 *   measured in kronor for this person -- that the first year of the pension is taxed harder than the
 *   next, which payout period pays the most up to the expected lifespan and from which lifespan the
 *   others catch up, what savings cost in housing supplement, the ISK tax and its tax-free amount,
 *   state tax from a short payout, and how little it costs to spend the savings early.
 * - **Three scenarios**: balanced (everything for life), the person's own plan (everything over the
 *   shorter period they give) and the optimal one (`optimalCombo`: of every way of setting each pot's
 *   payout period, the one that leaves the most to live on in all up to the expected lifespan while the
 *   lowest month stays at or above the floor), each with what it leaves to live on a month after tax,
 *   housing supplement included, in three stretches of life; the total at four lifespans; the lowest
 *   month up to 95; and the lifespan from which the balanced way has paid more. The optimal one is the
 *   one suggested.
 * - **Tax and tax reductions** (`taxPicture`): the rules on what of the pension is kept -- always draw at least the
 *   tax-free annual income, ROT and RUT and the other tax reductions only take off tax actually paid, interest is set
 *   against an ISK's schablonintäkt first -- with Skatteverket's amounts for the year (`TAX_YEARS`, typed in yearly)
 *   and the person's own figures.
 * - **A checklist** of at most five steps, dated where a date matters (jämkning).
 *
 * Everything is in fixed prices, as in the other question-mark tabs, and what is left to live on is the
 * model's: the pension after tax (`netto`), the housing supplement and other benefits (`bidrag`), and the
 * private savings paid out tax-free. The first year is Table 1's, as "Hur länge räcker pensionen?" takes
 * it, so it matches Prognos. The housing supplement is tested against the private savings' balance as it
 * is spent (the engine's `formogenhetPerAge`), not against one amount for the whole retirement -- spending
 * the savings early is part of what the tab weighs.
 *
 * Pure: no DOM, so the arithmetic is tested directly (`apps/web/test/strategy.test.ts`).
 */
import { Scheme, Table1Key, defaultContext, defaultInput, riktalderFor, runWithTaxProbe, run } from "@typfallsmodellen/engine";
import type { DeathProbabilities, ModelContext, MvaluesRow, SchemeId, TypfallInput } from "@typfallsmodellen/engine";

import { kronor, percent } from "./format.js";
import type { Lang } from "./i18n.js";
import { retirementAge } from "./kpis.js";
import { LAST_AGE, monthlySalaryForFinalSalary, referenceAge } from "./pensionAge.js";
import { BUDGETS } from "./pensionLasts.js";
import { defaultSavings, savingSettings, savingYear, taxFreeSaving, taxedSaving } from "./privateSavings.js";
import type { SavingPart, SavingsInput, SavingsKind } from "./privateSavings.js";
import { csvLine, delimiterFor } from "./tables.js";
import { kronorCell } from "./xlsx.js";
import type { SheetData } from "./xlsx.js";

const say = (l: Lang, sv: string, en: string) => (l === "sv" ? sv : en);

// ----------------------------------------------------------------- input ---

/** The front-loaded scenarios' payout period, years: what the agreements offer short of life. */
export const FRONT_BOUNDS = { min: 5, max: 20, step: 1 } as const;
export const FLOOR_BOUNDS = { min: 0, max: 200_000, step: 100 } as const;
export const RENT_BOUNDS = { min: 0, max: 50_000, step: 100 } as const;
/** The expected lifespan one can type: an age; 0 in the input stands for the model's own. */
export const LIFESPAN_BOUNDS = { min: 60, max: 100, step: 1 } as const;
/** A long life: the lowest month is looked for up to here, and it is one of the lifespans the totals are given at. */
export const LONG_LIFE = 95;
/** The first years of retirement, when most people want the most: what "early" averages. */
export const EARLY_YEARS = 10;
/** The fixed payout periods the occupational pension is compared over, besides life. */
export const PAYOUT_OPTIONS = [5, 10, 15, 20] as const;
/** The period 0: for life. */
export const LIFELONG = 0;
/** The housing supplement's wealth test disregards this much (the engine's `btpWealth`). */
export const WEALTH_DISREGARDED = 100_000;
/** What an ISK, a KF or a depot insurance is not taxed on, per person, from 2026 (the engine's `AvkastningsskattKFISK`). */
export const ISK_TAX_FREE = 300_000;
/** The tax on capital income: what an ISK's schablonintäkt is taxed at. */
export const CAPITAL_TAX = 0.3;
/** The most occupational pensions, and the most private savings, that can be added beside the first. */
export const MAX_EXTRA = 5;

/**
 * Skatteverket's amounts the tax rules quote, by income year. They change every year and are updated by hand
 * (docs/YEARLY-UPDATE.md). `taxFreeIncome` is "Skattefri årsinkomst" on Belopp och procent: 42,3 % of the
 * prisbasbelopp, the most a year that carries no income tax (a return is due from a krona more), and the line
 * for income from work to earn a public pension. `rotRutMax` is the most ROT and RUT together can take off one
 * person's tax in a year, `rotMax` the most of it ROT can be. A year not listed takes the latest one before it.
 */
export const TAX_YEARS: Readonly<Record<number, { readonly taxFreeIncome: number; readonly rotRutMax: number; readonly rotMax: number }>> = {
  2026: { taxFreeIncome: 25_041, rotRutMax: 75_000, rotMax: 50_000 },
};

/** Skatteverket's pages the amounts come from and the tax rules point to. */
export const SKATTEVERKET = {
  amounts: "https://www.skatteverket.se/privat/skatter/beloppochprocent.4.18e1b10334ebe8bc80004109.html",
  rotRut: "https://www.skatteverket.se/privat/fastigheterochbostad/rotarbeteochrutarbete.4.2e56d4ba1202f95012080002966.html",
  interest: "https://www.skatteverket.se/privat/deklaration/avdragforprivatpersoner/avdragforranteutgifter.106.1b39a64a1919eabb488b87.html",
  reductions: "https://www.skatteverket.se/privat/skatter/arbeteochinkomst/skattereduktioner.4.3810a01c150939e893f1a17e.html",
} as const;

/** The amounts for `year`: its own, or the latest listed year's before it (the first year's, before any). */
export function taxYear(year: number): { readonly year: number; readonly taxFreeIncome: number; readonly rotRutMax: number; readonly rotMax: number } {
  const years = Object.keys(TAX_YEARS)
    .map(Number)
    .sort((a, b) => a - b);
  const pick = years.filter((y) => y <= year).pop() ?? years[0]!;
  return { year: pick, ...TAX_YEARS[pick]! };
}

export interface StrategyInput extends SavingsInput {
  readonly born: number;
  readonly startWorkAge: number;
  /** The age the public pension starts at. */
  readonly retirementAge: number;
  /** Per month, in the reference year's money. */
  readonly monthlySalary: number;
  readonly scheme: SchemeId;
  readonly yearlyInflation: number;
  readonly realGrowth: number;
  readonly realReturn: number;
  /** The housing cost a month, kronor: what the housing supplement is worked out from. */
  readonly rent: number;
  /** Count the housing supplement and the other benefits the model gives. */
  readonly housingSupplement: boolean;
  /** The least a month, after tax, the person wants to have to live on as long as they live. */
  readonly floor: number;
  /** The own plan's payout period, years (and the hybrid insight's). */
  readonly frontYears: number;
  /** The age the person expects to live to: what the totals, insights and checklist are read up to. 0 is the model's own. */
  readonly lifespan: number;
  /** More occupational pensions, from earlier employers or other agreements, beside the agreement's own: taxed as income. */
  readonly extraPensions: readonly ExtraPension[];
  /** More private savings beside the first, each held in one of the two kinds. */
  readonly extraSavings: readonly ExtraSaving[];
}

/** An occupational pension of its own: capital already saved, and what is still paid in a month until retirement. */
export interface ExtraPension {
  readonly capital: number;
  readonly monthly: number;
}

/** A private saving of its own, held in one of the two kinds. */
export interface ExtraSaving {
  readonly kind: SavingsKind;
  readonly capital: number;
  readonly monthly: number;
}

/**
 * Someone a year or two from retirement, so that the first-year effect and the checklist's dates are
 * near: born 1961, the model's normal salary, SAF-LO (the workers' agreement, premium-based, with a
 * payout period of one's own choosing), half a million in an ISK, the model's own housing cost, and
 * minPension's basic pensioner as the floor.
 */
export function defaultStrategyInput(): StrategyInput {
  const base = defaultInput();
  const born = 1961;
  return {
    born,
    startWorkAge: 25,
    retirementAge: riktalderFor(born),
    monthlySalary: base.monthlySalary,
    scheme: Scheme.SafLo,
    ...defaultSavings(),
    savingsCapital: 500_000,
    yearlyInflation: base.yearlyInflation,
    realGrowth: base.realGrowth,
    realReturn: base.realReturn,
    rent: defaultContext().hyra,
    housingSupplement: true,
    floor: BUDGETS[0].amount,
    frontYears: 10,
    lifespan: 0,
    extraPensions: [],
    extraSavings: [],
  };
}

/** What an added pension's or saving's capital and monthly saving may be: as the first one's fields. */
const CAPITAL_BOUNDS = { min: 0, max: 20_000_000 } as const;
const SAVING_BOUNDS = { min: 0, max: 100_000 } as const;

const clamp = (value: number, bounds: { readonly min: number; readonly max: number }): number =>
  Math.min(Math.max(Number.isFinite(value) ? value : bounds.min, bounds.min), bounds.max);

/** The retirement age moved inside what the cohort may draw, and the rest inside its bounds. */
export function normaliseStrategy(input: StrategyInput): StrategyInput {
  return {
    ...input,
    retirementAge: referenceAge(input.born, input.retirementAge),
    frontYears: Math.round(clamp(input.frontYears, FRONT_BOUNDS)),
    floor: clamp(input.floor, FLOOR_BOUNDS),
    rent: clamp(input.rent, RENT_BOUNDS),
    lifespan:
      input.lifespan > 0
        ? Math.round(Math.max(clamp(input.lifespan, LIFESPAN_BOUNDS), referenceAge(input.born, input.retirementAge) + 1))
        : 0,
    extraPensions: (input.extraPensions ?? []).slice(0, MAX_EXTRA).map((p) => ({ capital: clamp(p.capital, CAPITAL_BOUNDS), monthly: clamp(p.monthly, SAVING_BOUNDS) })),
    extraSavings: (input.extraSavings ?? []).slice(0, MAX_EXTRA).map((p) => ({
      kind: p.kind === "depot" ? "depot" : "isk",
      capital: clamp(p.capital, CAPITAL_BOUNDS),
      monthly: clamp(p.monthly, SAVING_BOUNDS),
    })),
  };
}

/**
 * What is paid out as taxed income, as the model's private pension pot carries it: the individual occupational
 * pension, the private pension and every occupational pension added, summed. The model's pot for them is one --
 * the same return, payout period and tax -- so the added pensions are paid over the private pension's period.
 */
export function pooledPension(input: StrategyInput): SavingsInput {
  const extras = input.extraPensions ?? [];
  return {
    ...input,
    individualCapital: input.individualCapital + extras.reduce((sum, p) => sum + Math.max(p.capital, 0), 0),
    individualMonthly: input.individualMonthly + extras.reduce((sum, p) => sum + Math.max(p.monthly, 0), 0),
  };
}

/**
 * The private savings by kind, summed: the first saving and every one added. The model carries one kind of private
 * saving per run, so each kind with anything in it is a run of its own; within a kind the savings are one pot.
 */
export function savingsByKind(input: StrategyInput): readonly SavingsInput[] {
  const all = [{ kind: input.savingsKind, capital: input.savingsCapital, monthly: input.savingsMonthly }, ...(input.extraSavings ?? [])];
  const kinds: SavingsKind[] = ["isk", "depot"];
  return kinds
    .map((kind) => {
      const of = all.filter((s) => s.kind === kind);
      return {
        ...input,
        savingsKind: kind,
        savingsCapital: of.reduce((sum, s) => sum + Math.max(s.capital, 0), 0),
        savingsMonthly: of.reduce((sum, s) => sum + Math.max(s.monthly, 0), 0),
      };
    })
    .filter((s) => taxFreeSaving(s) !== null);
}

/** Whether there is a private saving of `kind` in the plan. */
export function hasSavingsOf(input: StrategyInput, kind: SavingsKind): boolean {
  return savingsByKind(input).some((s) => s.savingsKind === kind);
}

export function strategyTypfall(input: StrategyInput): TypfallInput {
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

// ------------------------------------------------------------------ runs ---

/** One way of drawing the pots. A payout period of 0 is for life. */
export interface Combo {
  readonly occupational: number;
  /** The individual and private pension: the model's IPS, taxed as income. */
  readonly pension: number;
  /** The private savings, paid out tax-free. */
  readonly savings: number;
  /** Years the occupational and private pension wait after retirement: 0, or 1 to start the January after. */
  readonly delay: number;
}

/**
 * The context of one run: the payout periods, the start, fixed prices, the housing, and the model's private
 * saving carrying one of the two parts (`savingSettings`): every taxed pension pooled (`pooledPension`), or one
 * kind of private saving, `savings` (the first in the plan unless given). The taxed part's run also gets the wealth the
 * housing supplement is tested against, year by year.
 */
export function strategyContext(
  shared: ModelContext,
  input: StrategyInput,
  combo: Combo,
  part: SavingPart,
  wealth?: readonly number[],
  savings: SavingsInput = savingsByKind(input)[0] ?? { ...input, savingsCapital: 0, savingsMonthly: 0 },
): ModelContext {
  return {
    ...shared,
    priceBasis: 1,
    tjpPar: part === "pension" && combo.delay > 0 ? input.retirementAge + combo.delay : 0,
    tempTjpUttag: combo.occupational,
    tempIpsUttag: part === "pension" ? combo.pension : combo.savings,
    hyra: input.rent,
    ansokt: input.housingSupplement ? 1 : 0,
    ...savingSettings(shared, input, part === "pension" ? pooledPension(input) : savings, part),
    ...(wealth === undefined ? {} : { formogenhetPerAge: wealth }),
  };
}

/** Kronor a year by age, `0..LAST_AGE`, from the retirement age; zero before it. */
function annual(rows: readonly MvaluesRow[], par: number, pick: (row: MvaluesRow) => number): number[] {
  const out = new Array<number>(LAST_AGE + 1).fill(0);
  for (const row of rows) if (row.age >= par && row.age <= LAST_AGE) out[row.age] = pick(row);
  return out;
}

/** The private savings runs, summed: what they pay, their balance and their tax, and the wealth they make by age. */
interface SavingsRun {
  readonly savings: readonly number[];
  /** Fixed prices, at the end of each age. */
  readonly balance: readonly number[];
  readonly tax: readonly number[];
  /** The part of `balance` and `tax` held in the ISK kind: what the ISK's schablonintäkt is read from. */
  readonly iskBalance: readonly number[];
  readonly iskTax: readonly number[];
  /** Nominal kronor at the start of each age, `0..105`, other wealth included: the housing supplement's test. */
  readonly wealth: readonly number[];
}

/** One run per kind of private saving in the plan (`savingsByKind`), the two added up. */
function savingsRun(typfall: TypfallInput, input: StrategyInput, shared: ModelContext, deaths: DeathProbabilities, payout: number): SavingsRun {
  const zero = () => new Array<number>(LAST_AGE + 1).fill(0);
  const other = Math.max(shared.formogenhet, 0);
  const savings = zero();
  const balance = zero();
  const tax = zero();
  const iskBalance = zero();
  const iskTax = zero();
  const wealth = new Array<number>(106).fill(other);
  const combo: Combo = { occupational: LIFELONG, pension: LIFELONG, savings: payout, delay: 0 };
  for (const kind of savingsByKind(input)) {
    const result = run(typfall, strategyContext(shared, input, combo, "savings", undefined, kind), { deaths });
    const par = retirementAge(typfall, result);
    const paid = annual(result.rows, par, (r) => r.pps);
    // The first year is Table 1's, as "Hur länge räcker pensionen?" takes it.
    const line = result.table1.find((r) => r.key === Table1Key.PrivateSavingAfterTax);
    if (line !== undefined && par <= LAST_AGE) paid[par] = line.monthly * 12;
    paid.forEach((v, age) => (savings[age] = savings[age]! + v));
    const isk = kind.savingsKind === "isk";
    for (const row of result.ledger) {
      if (row.age >= 0 && row.age <= LAST_AGE) {
        const close = row.pots.pps.close * row.priceFactor;
        const yieldTax = row.pots.pps.yieldTax * row.priceFactor;
        balance[row.age] = balance[row.age]! + close;
        tax[row.age] = tax[row.age]! + yieldTax;
        if (isk) {
          iskBalance[row.age] = iskBalance[row.age]! + close;
          iskTax[row.age] = iskTax[row.age]! + yieldTax;
        }
      }
      if (row.age >= 0 && row.age < wealth.length) wealth[row.age] = wealth[row.age]! + Math.max(row.pots.pps.open, 0);
    }
  }
  return { savings, balance, tax, iskBalance, iskTax, wealth };
}

/** One way of drawing the pots, run: kronor a year by age, fixed prices. */
export interface ComboRun {
  readonly combo: Combo;
  readonly par: number;
  readonly publicPension: readonly number[];
  /** The public pension's parts: income pension (with ATP), premium pension, guarantee pension and its supplement. */
  readonly incomePension: readonly number[];
  readonly premiumPension: readonly number[];
  readonly guaranteePension: readonly number[];
  readonly occupational: readonly number[];
  readonly individual: readonly number[];
  /** All that is taxed: the three above. */
  readonly gross: readonly number[];
  /** What is left of `gross` after the model's tax. */
  readonly net: readonly number[];
  /** Housing supplement and the model's other benefits. */
  readonly benefits: readonly number[];
  /** The private savings paid out, tax-free. */
  readonly savings: readonly number[];
  /** What there is to live on: `net + benefits + savings`. */
  readonly disposable: readonly number[];
  readonly stateTax: readonly number[];
  /** The private savings' balance at the end of each age and the tax on them, all kinds. */
  readonly savingsBalance: readonly number[];
  readonly savingsTax: readonly number[];
  /** The same, for what is held in the ISK kind only. */
  readonly iskBalance: readonly number[];
  readonly iskTax: readonly number[];
  /** What is left of `gross` kronor a year at `age` after the model's tax. */
  readonly netAt: (age: number, gross: number) => number;
  /** What a nominal amount at each age is multiplied by to be in this run's fixed prices, `0..LAST_AGE`. */
  readonly kpiFactor: readonly number[];
  /** The model's expected age at death for the person at retirement. */
  readonly expectedDeath: number;
  /** Table 1's "Slutlön" per month: what a typed final salary is matched against. */
  readonly finalSalary: number;
  readonly warnings: readonly string[];
}

function pensionRun(
  typfall: TypfallInput,
  input: StrategyInput,
  shared: ModelContext,
  deaths: DeathProbabilities,
  combo: Combo,
  own: SavingsRun,
  wealth: readonly number[] = own.wealth,
): ComboRun {
  const probe = runWithTaxProbe(typfall, strategyContext(shared, input, combo, "pension", wealth), { deaths });
  const { rows } = probe.result;
  const par = retirementAge(typfall, probe.result);
  const publicPension = annual(rows, par, (r) => r.ip + r.tp + r.pp + r.garp + r.ptillagg);
  const incomePension = annual(rows, par, (r) => r.ip + r.tp);
  const premiumPension = annual(rows, par, (r) => r.pp);
  const guaranteePension = annual(rows, par, (r) => r.garp + r.ptillagg);
  const occupational = annual(rows, par, (r) => r.tjp);
  const individual = annual(rows, par, (r) => r.ips);
  const net = annual(rows, par, (r) => r.netto);
  const benefits = annual(rows, par, (r) => r.bidrag);
  const stateTax = annual(rows, par, (r) => r.stateTax);

  // The first year is Table 1's -- what Prognos shows: the model draws Table 1 in a second pass at the
  // retirement age, with the last working year's pension right in. What is left after tax follows the gross.
  const yearly = (key: Table1Key): number => (probe.result.table1.find((r) => r.key === key)?.monthly ?? 0) * 12;
  if (par <= LAST_AGE) {
    const before = publicPension[par]! + occupational[par]! + individual[par]!;
    publicPension[par] = yearly(Table1Key.TotalPublicPension);
    occupational[par] = yearly(Table1Key.OccupationalPension);
    individual[par] = yearly(Table1Key.PrivateSaving);
    const after = publicPension[par]! + occupational[par]! + individual[par]!;
    net[par] = net[par]! + probe.netAt(par, after) - probe.netAt(par, before);
    benefits[par] = yearly(Table1Key.BenefitsAtRetirement);
  }
  const gross = publicPension.map((p, age) => p + occupational[age]! + individual[age]!);
  const kpiFactor = new Array<number>(LAST_AGE + 1).fill(1);
  for (const row of rows) if (row.age >= 0 && row.age <= LAST_AGE) kpiFactor[row.age] = row.kpiFactor;
  return {
    combo,
    par,
    publicPension,
    incomePension,
    premiumPension,
    guaranteePension,
    occupational,
    individual,
    gross,
    net,
    benefits,
    savings: own.savings,
    disposable: net.map((n, age) => n + benefits[age]! + own.savings[age]!),
    stateTax,
    savingsBalance: own.balance,
    savingsTax: own.tax,
    iskBalance: own.iskBalance,
    iskTax: own.iskTax,
    netAt: probe.netAt,
    kpiFactor,
    expectedDeath: probe.result.lifeIncome.throughAge,
    finalSalary: yearly(Table1Key.FinalSalary) / 12,
    warnings: probe.result.warnings.map((w) => w.message),
  };
}

// ------------------------------------------------------------- scenarios ---

export type ScenarioId = "balanced" | "front" | "optimal";

export interface Scenario {
  readonly id: ScenarioId;
  readonly run: ComboRun;
  /** What there is to live on a month, on average, in each stretch of life. */
  readonly phases: readonly { readonly from: number; readonly to: number; readonly monthly: number }[];
  /** In all, from retirement up to each lifespan (the age the person is gone at). */
  readonly totals: readonly { readonly lifespan: number; readonly total: number }[];
  /** The least a month up to `LONG_LIFE`, and the age it falls at. */
  readonly lowest: { readonly age: number; readonly monthly: number };
  /** A month on average over the first `EARLY_YEARS`. */
  readonly early: number;
  /** The lifespan from which the balanced scenario has paid more in all, where this one is ahead before it. */
  readonly breakEven: number | undefined;
  readonly meetsFloor: boolean;
}

/** In all, from retirement up to the age the person is gone at (`lifespan` itself not counted). */
export function totalTo(run: ComboRun, lifespan: number): number {
  let sum = 0;
  for (let age = run.par; age < Math.min(lifespan, LAST_AGE + 1); age += 1) sum += run.disposable[age]!;
  return sum;
}

const average = (values: readonly number[], from: number, to: number): number => {
  let sum = 0;
  let count = 0;
  for (let age = from; age <= Math.min(to, LAST_AGE); age += 1) {
    sum += values[age]!;
    count += 1;
  }
  return count > 0 ? sum / count : 0;
};

/**
 * The lifespan from which `other` has paid at least as much as `run` in all, for every longer life up to the
 * last age the model runs to -- where `run` is ahead at some shorter lifespan; undefined when it never is, or
 * is still ahead at the end.
 */
export function overtakenAt(run: ComboRun, other: ComboRun): number | undefined {
  let ahead = false;
  let from: number | undefined;
  let lead = 0;
  for (let lifespan = run.par + 1; lifespan <= LAST_AGE + 1; lifespan += 1) {
    lead += run.disposable[lifespan - 1]! - other.disposable[lifespan - 1]!;
    if (lead > 0.5) {
      ahead = true;
      from = undefined;
    } else if (from === undefined) {
      from = lifespan;
    }
  }
  return ahead ? from : undefined;
}

/** The lifespans the totals are given at: 80, the expected one, 90 and a long life, past retirement. */
export function planAges(run: ComboRun, expectedDeath: number): readonly number[] {
  const expected = Math.round(expectedDeath);
  return [...new Set([80, expected, 90, LONG_LIFE])].filter((age) => age > run.par).sort((a, b) => a - b);
}

function scenario(id: ScenarioId, run: ComboRun, balanced: ComboRun | undefined, floor: number, expectedDeath: number): Scenario {
  const par = run.par;
  const phases = [
    { from: par, to: par + EARLY_YEARS - 1 },
    { from: par + EARLY_YEARS, to: par + 2 * EARLY_YEARS - 1 },
    { from: par + 2 * EARLY_YEARS, to: LONG_LIFE },
  ]
    .filter((phase) => phase.from <= Math.min(phase.to, LONG_LIFE))
    .map((phase) => ({ ...phase, to: Math.min(phase.to, LONG_LIFE), monthly: average(run.disposable, phase.from, phase.to) / 12 }));
  let lowest = { age: par, monthly: Number.POSITIVE_INFINITY };
  for (let age = par; age <= Math.min(LONG_LIFE, LAST_AGE); age += 1) {
    const monthly = run.disposable[age]! / 12;
    if (monthly < lowest.monthly) lowest = { age, monthly };
  }
  return {
    id,
    run,
    phases,
    totals: planAges(run, expectedDeath).map((lifespan) => ({ lifespan, total: totalTo(run, lifespan) })),
    lowest,
    early: average(run.disposable, par, par + EARLY_YEARS - 1) / 12,
    breakEven: balanced === undefined || id === "balanced" ? undefined : overtakenAt(run, balanced),
    meetsFloor: lowest.monthly + 0.5 >= floor,
  };
}

/** The payout periods the optimal scenario chooses among for each pot: the agreements' fixed ones, and life. */
export const OPTIMAL_PERIODS: readonly number[] = [...PAYOUT_OPTIONS, LIFELONG];

/** How many of a combo's three periods are for life: the simpler way, which a tie goes to. */
const lifelongCount = (c: Combo) => [c.occupational, c.pension, c.savings].filter((p) => p === LIFELONG).length;
/** The total length of a combo's fixed periods: a tie between as many lifelong pots goes to the longer periods. */
const spread = (c: Combo) => [c.occupational, c.pension, c.savings].reduce((sum, p) => sum + (p === LIFELONG ? 0 : p), 0);

/**
 * The optimal way of drawing the pots, of the runs given: the one that leaves the most to live on in all up to
 * `lifespan` among those whose lowest month up to `LONG_LIFE` stays at or above the floor, or -- if none does --
 * the one whose lowest month is highest. A tie (within a krona) goes to the one with more pots for life, then to
 * the longer periods: the simpler, safer way.
 */
export function optimalCombo(runs: readonly ComboRun[], floor: number, lifespan: number): ComboRun {
  const lowest = (r: ComboRun) => {
    let low = Number.POSITIVE_INFINITY;
    for (let age = r.par; age <= Math.min(LONG_LIFE, LAST_AGE); age += 1) low = Math.min(low, r.disposable[age]! / 12);
    return low;
  };
  const scored = runs.map((run) => ({ run, low: lowest(run), total: totalTo(run, lifespan) }));
  const safe = scored.filter((s) => s.low + 0.5 >= floor);
  const pool = safe.length > 0 ? safe : scored;
  const score = (s: (typeof scored)[number]) => (safe.length > 0 ? s.total : s.low * 12);
  let best = pool[0]!;
  for (const s of pool) {
    const diff = score(s) - score(best);
    if (diff > 1) best = s;
    else if (diff >= -1) {
      const simpler = lifelongCount(s.run.combo) - lifelongCount(best.run.combo) || spread(s.run.combo) - spread(best.run.combo);
      if (simpler > 0) best = s;
    }
  }
  return best.run;
}

// ------------------------------------------------------------- tax rules ---

/** An amount a year at an age. */
export interface AtAge {
  readonly age: number;
  readonly amount: number;
}

/**
 * What the tax rules come to for this person, read off the suggested scenario: kronor a year in the tab's fixed
 * prices, beside Skatteverket's amounts for the year they are for (`TAX_YEARS`).
 */
export interface TaxPicture {
  /** The year the amounts are for, and the amounts in that year's kronor. */
  readonly year: number;
  readonly taxFreeIncome: number;
  readonly rotRutMax: number;
  readonly rotMax: number;
  /** The tax-free income in the tab's fixed prices. */
  readonly taxFreeFixed: number;
  readonly scenario: ScenarioId;
  /** What the model leaves untaxed a year once the higher basic deduction applies, from the age it applies. */
  readonly pensionerTaxFree: AtAge;
  /** The least pension before tax a year, retirement to `LONG_LIFE`, and the ages it is under the tax-free income. */
  readonly lowestTaxable: AtAge;
  readonly below: readonly number[];
  /** An occupational or private pension in the plan, which could be drawn earlier to fill those years. */
  readonly canTopUp: boolean;
  /** The tax on the pension a year after the first year (taken higher, see `firstYear`): the most and the least. */
  readonly taxHigh: AtAge;
  readonly taxLow: AtAge;
  /** The ages, retirement to `LONG_LIFE`, with no tax on the pension: no tax reduction can be had then. */
  readonly untaxed: readonly number[];
  /** Whether the private savings are taxed on a schablonintäkt (the ISK kind), and that intäkt the first year. */
  readonly hasIsk: boolean;
  readonly iskIncome: number;
}

/** The most a year that is left untaxed at `age`: where the model's tax starts, to the hundred below. */
export function zeroTaxLevel(run: ComboRun, age: number): number {
  let lo = 0;
  let hi = 1_000_000;
  for (let i = 0; i < 40; i += 1) {
    const mid = (lo + hi) / 2;
    if (run.netAt(age, mid) >= mid - 0.5) lo = mid;
    else hi = mid;
  }
  return Math.floor(lo / 100) * 100;
}

export function taxPicture(scenario: Scenario, input: StrategyInput, today: Date): TaxPicture {
  const { run } = scenario;
  const par = run.par;
  const last = Math.min(LONG_LIFE, LAST_AGE);
  const amounts = taxYear(today.getFullYear());
  const factor = run.kpiFactor[Math.min(Math.max(amounts.year - input.born, 0), LAST_AGE)] ?? 1;
  const taxFreeFixed = amounts.taxFreeIncome * factor;

  let lowestTaxable: AtAge = { age: par, amount: Number.POSITIVE_INFINITY };
  const below: number[] = [];
  for (let a = par; a <= last; a += 1) {
    if (run.gross[a]! < lowestTaxable.amount) lowestTaxable = { age: a, amount: run.gross[a]! };
    if (run.gross[a]! < taxFreeFixed - 0.5) below.push(a);
  }
  const tax = (a: number) => Math.max(run.gross[a]! - run.net[a]!, 0);
  const untaxed: number[] = [];
  for (let a = par; a <= last; a += 1) if (tax(a) < 1) untaxed.push(a);
  const from = Math.min(par + 1, last);
  let taxHigh: AtAge = { age: from, amount: tax(from) };
  let taxLow: AtAge = taxHigh;
  for (let a = from; a <= last; a += 1) {
    if (tax(a) > taxHigh.amount + 0.5) taxHigh = { age: a, amount: tax(a) };
    if (tax(a) < taxLow.amount - 0.5) taxLow = { age: a, amount: tax(a) };
  }
  const raised = Math.min(Math.max(par, riktalderFor(input.born)) + 1, LAST_AGE);
  const hasIsk = hasSavingsOf(input, "isk");
  return {
    ...amounts,
    taxFreeFixed,
    scenario: scenario.id,
    pensionerTaxFree: { age: raised, amount: zeroTaxLevel(run, raised) },
    lowestTaxable,
    below,
    canTopUp: run.occupational.some((v) => v > 0) || run.individual.some((v) => v > 0),
    taxHigh,
    taxLow,
    untaxed,
    hasIsk,
    iskIncome: hasIsk ? Math.max(run.iskTax[par] ?? 0, 0) / CAPITAL_TAX : 0,
  };
}

// -------------------------------------------------------------- insights ---

export type Insight =
  | {
      /** The housing supplement the savings cost, against none. */
      readonly kind: "btpWealth";
      readonly cost: number;
      readonly lifespan: number;
      readonly savingsCapital: number;
      /** Less housing supplement a month for each 100 000 kronor over the disregarded amount. */
      readonly perHundred: number;
    }
  | {
      /** Which occupational payout period pays the most in all up to the expected lifespan. */
      readonly kind: "payout";
      readonly lifespan: number;
      readonly best: number;
      readonly worst: number;
      readonly bestTotal: number;
      readonly worstTotal: number;
      /** The shortest period, and the lifespan from which `best` has paid more than it -- if it is not the best. */
      readonly shortest: number;
      readonly shortestUntil: number | undefined;
      /** The best is short, and life overtakes it only at this lifespan (undefined: not before the model's last age). */
      readonly lifelongFrom: number | undefined;
      /** The housing supplement makes most of the difference. */
      readonly housingDriven: boolean;
    }
  | {
      /** The first year of the pension is taxed harder than the next: the higher basic deduction comes the year after. */
      readonly kind: "firstYear";
      readonly year: number;
      readonly gross: number;
      readonly rateFirst: number;
      readonly rateNext: number;
      /** Kronor a year more left after tax, on the same gross, the year after. */
      readonly gain: number;
      /** The occupational and private pension a year, and what starting them the January after saves in tax. */
      readonly slice: number;
      readonly saving: number;
      /** What the first year then needs from the savings, after tax. */
      readonly bridge: number;
      /** The savings there are to take it from at retirement. */
      readonly savingsAtRetirement: number;
    }
  | {
      /** The yearly tax on an ISK, a KF or a depot insurance, and the amount it is not taken on. */
      readonly kind: "iskTax";
      readonly perYear: number;
      readonly total: number;
      readonly lifespan: number;
      readonly capital: number;
    }
  | {
      /** State tax from the front-loaded scenario's largest years. */
      readonly kind: "stateTax";
      readonly total: number;
      readonly years: number;
      readonly avoidWith: number | undefined;
    }
  | {
      /** Spending the savings over the first years while the pensions are kept for life. */
      readonly kind: "hybrid";
      readonly years: number;
      readonly monthlyEarly: number;
      readonly gainBy80: number;
      readonly costByLongLife: number;
    };

export type InsightKind = Insight["kind"];

/** The order the insights are offered in; the first three that apply are shown. */
export const INSIGHT_ORDER: readonly InsightKind[] = ["btpWealth", "payout", "firstYear", "iskTax", "stateTax", "hybrid"];

/** Kronor below this are not worth a person's attention. */
const NOTEWORTHY = 1_000;

// ------------------------------------------------------------ checklist ---

export type ChecklistItem =
  | {
      readonly kind: "jamkning";
      /** The income year the jämkning is for. */
      readonly year: number;
      /** "D-M-YYYY" where the last day is known. */
      readonly deadline: string | undefined;
      /** Before the last day this autumn; after it but still this autumn's; or an autumn to come. */
      readonly when: "soon" | "late" | "later";
    }
  | { readonly kind: "minpension" }
  | {
      /** Years when the pension before tax is under the tax-free income, and whether a pension could be drawn earlier. */
      readonly kind: "taxFree";
      readonly from: number;
      readonly to: number;
      readonly amount: number;
      readonly year: number;
      readonly topUp: boolean;
    }
  | { readonly kind: "payout"; readonly best: number; readonly worst: number; readonly difference: number; readonly lifespan: number }
  | { readonly kind: "start"; readonly year: number; readonly saving: number; readonly bridge: number; readonly covered: boolean }
  | {
      readonly kind: "btp";
      readonly monthly: number;
      /** The age it is first paid at in the recommended scenario, and whether that is the retirement age. */
      readonly fromAge: number;
      readonly atStart: boolean;
    }
  | { readonly kind: "isk"; readonly perYear: number }
  | { readonly kind: "scenario"; readonly id: ScenarioId; readonly years: number; readonly combo: Combo; readonly kinds: PotKinds };

/**
 * Skatteverket's last day to apply for jämkning that holds for the whole of the next year, by the year it
 * falls in. It moves from year to year (2025: 9 November); a year not listed is told "early November".
 */
export const JAMKNING_DEADLINES: Readonly<Record<number, { readonly month: number; readonly day: number }>> = {
  2025: { month: 11, day: 9 },
  2026: { month: 11, day: 8 },
};

// -------------------------------------------------------------- strategy ---

export interface Strategy {
  readonly input: StrategyInput;
  readonly par: number;
  readonly retirementYear: number;
  /** The lifespan everything is read up to: the person's own, or the model's. */
  readonly expectedDeath: number;
  /** The model's expected age at death for the person at retirement. */
  readonly modelExpectedDeath: number;
  readonly scenarios: readonly Scenario[];
  readonly recommended: ScenarioId;
  /** The occupational payout periods compared, for life last. */
  readonly payouts: readonly ComboRun[];
  readonly insights: readonly Insight[];
  /** The tax rules for this person: the tax-free income, ROT and RUT, interest. */
  readonly tax: TaxPicture;
  readonly checklist: readonly ChecklistItem[];
  readonly referenceYear: number;
  readonly savingYear: number;
  readonly warnings: readonly string[];
}

/** The engine runs, kept while only the floor changes. */
export interface StrategyRuns {
  readonly balanced: ComboRun;
  readonly front: ComboRun;
  readonly hybrid: ComboRun | undefined;
  readonly payouts: readonly ComboRun[];
  /** Every way of setting the pots' payout periods the optimal scenario chooses among (`OPTIMAL_PERIODS`). */
  readonly candidates: readonly ComboRun[];
  readonly withoutSavings: ComboRun | undefined;
  readonly extraWealth: readonly [ComboRun, ComboRun] | undefined;
}

/** What a set of runs is valid for: everything but the floor and the lifespan, which only the reading of them uses. */
export function strategyRunsKey(input: StrategyInput): string {
  const { floor: _floor, lifespan: _lifespan, ...rest } = input;
  return JSON.stringify(rest);
}

export function runStrategy(input: StrategyInput, shared: ModelContext, deaths: DeathProbabilities): StrategyRuns {
  const typfall = strategyTypfall(input);
  const n = input.frontYears;
  const hasTaxFree = savingsByKind(input).length > 0;
  const taxed = taxedSaving(pooledPension(input));
  const hasTaxed = taxed.capital > 0 || taxed.monthly > 1;
  const life: Combo = { occupational: LIFELONG, pension: LIFELONG, savings: LIFELONG, delay: 0 };

  // The savings and the pension runs, each done once however many of the lists below it is in.
  const savingsRuns = new Map<number, SavingsRun>();
  const savingsFor = (payout: number): SavingsRun => {
    const key = hasTaxFree ? payout : LIFELONG;
    if (!savingsRuns.has(key)) savingsRuns.set(key, savingsRun(typfall, input, shared, deaths, key));
    return savingsRuns.get(key)!;
  };
  const pensionRuns = new Map<string, ComboRun>();
  const runFor = (combo: Combo): ComboRun => {
    const key = `${combo.occupational}|${combo.pension}|${combo.savings}`;
    if (!pensionRuns.has(key)) pensionRuns.set(key, pensionRun(typfall, input, shared, deaths, combo, savingsFor(combo.savings)));
    return pensionRuns.get(key)!;
  };
  const savingsLife = savingsFor(LIFELONG);
  const balanced = runFor(life);

  const hasOccupational = balanced.occupational.some((v, age) => age > balanced.par && v > 0);
  const front = runFor({ occupational: n, pension: n, savings: n, delay: 0 });
  const hybrid = hasTaxFree ? runFor({ ...life, savings: n }) : undefined;

  const payouts = hasOccupational ? [...PAYOUT_OPTIONS.map((years) => runFor({ ...life, occupational: years })), balanced] : [];

  // The optimal scenario's choice: each pot that has anything in it over every period, the others for life,
  // and the own plan and the hybrid besides, so that the optimal one is never worse than either.
  const periods = (has: boolean) => (has ? OPTIMAL_PERIODS : [LIFELONG]);
  for (const occupational of periods(hasOccupational))
    for (const pension of periods(hasTaxed))
      for (const savings of periods(hasTaxFree)) runFor({ occupational, pension, savings, delay: 0 });
  const candidates = [...pensionRuns.values()];

  // The housing supplement against the savings: none at all, and two amounts above the disregarded one.
  const paysHousing = input.housingSupplement && balanced.benefits.some((v) => v > 0);
  const other = Math.max(shared.formogenhet, 0);
  const constant = (wealth: number) => new Array<number>(106).fill(other + wealth);
  const withoutSavings =
    input.housingSupplement && hasTaxFree ? pensionRun(typfall, input, shared, deaths, life, savingsLife, constant(0)) : undefined;
  const extraWealth =
    withoutSavings !== undefined && (paysHousing || withoutSavings.benefits.some((v) => v > 0))
      ? ([
          pensionRun(typfall, input, shared, deaths, life, savingsLife, constant(200_000)),
          pensionRun(typfall, input, shared, deaths, life, savingsLife, constant(300_000)),
        ] as const)
      : undefined;

  return { balanced, front, hybrid, payouts, candidates, withoutSavings, extraWealth };
}

/** What the runs say, for a floor: the scenarios, the recommendation, the insights and the checklist. */
export function evaluateStrategy(runs: StrategyRuns, input: StrategyInput, shared: ModelContext, today: Date): Strategy {
  const { balanced } = runs;
  const par = balanced.par;
  const expectedDeath = input.lifespan > 0 ? input.lifespan : balanced.expectedDeath;
  const lifespan = Math.round(expectedDeath);
  const scenarios = [
    scenario("balanced", balanced, undefined, input.floor, expectedDeath),
    scenario("front", runs.front, balanced, input.floor, expectedDeath),
    scenario("optimal", optimalCombo(runs.candidates, input.floor, lifespan), balanced, input.floor, expectedDeath),
  ];
  const recommended: ScenarioId = "optimal";
  const insights = findInsights(runs, input, lifespan);
  const tax = taxPicture(scenarios.find((s) => s.id === recommended)!, input, today);
  const retirementYear = input.born + par;
  const warnings = [...new Set(scenarios.flatMap((s) => s.run.warnings))];
  return {
    input,
    par,
    retirementYear,
    expectedDeath,
    modelExpectedDeath: balanced.expectedDeath,
    scenarios,
    recommended,
    payouts: runs.payouts,
    insights,
    tax,
    checklist: buildChecklist(insights, scenarios, recommended, input, retirementYear, today, tax, lifespan),
    referenceYear: shared.referenceYear,
    savingYear: savingYear(input, shared),
    warnings,
  };
}

/** Table 1's "Slutlön" per month for the input: one run, everything for life. */
export function strategyFinalSalary(input: StrategyInput, shared: ModelContext, deaths: DeathProbabilities): number {
  const life: Combo = { occupational: LIFELONG, pension: LIFELONG, savings: LIFELONG, delay: 0 };
  const result = run(strategyTypfall(input), strategyContext(shared, input, life, "pension"), { deaths });
  return result.table1.find((r) => r.key === Table1Key.FinalSalary)?.monthly ?? 0;
}

/** The monthly salary, in the reference year's money, whose final salary is `target` (`monthlySalaryForFinalSalary`). */
export function strategySalaryForFinal(input: StrategyInput, target: number, shared: ModelContext, deaths: DeathProbabilities): number {
  return monthlySalaryForFinalSalary(target, (monthlySalary) => strategyFinalSalary({ ...input, monthlySalary }, shared, deaths));
}

export function computeStrategy(input: StrategyInput, shared: ModelContext, deaths: DeathProbabilities, today: Date): Strategy {
  return evaluateStrategy(runStrategy(input, shared, deaths), input, shared, today);
}

function sumBetween(values: readonly number[], from: number, toExclusive: number): number {
  let sum = 0;
  for (let age = Math.max(from, 0); age < Math.min(toExclusive, LAST_AGE + 1); age += 1) sum += values[age]!;
  return sum;
}

/** Every rule that applies, in `INSIGHT_ORDER`, at most three. */
function findInsights(runs: StrategyRuns, input: StrategyInput, lifespan: number): Insight[] {
  const { balanced } = runs;
  const par = balanced.par;
  const found: Insight[] = [];

  // The housing supplement the savings cost.
  if (runs.withoutSavings !== undefined) {
    const cost = sumBetween(runs.withoutSavings.benefits, par, lifespan) - sumBetween(balanced.benefits, par, lifespan);
    if (cost >= NOTEWORTHY) {
      const age = Math.min(par + 2, LAST_AGE);
      const perHundred =
        runs.extraWealth === undefined ? 0 : Math.max(runs.extraWealth[0].benefits[age]! - runs.extraWealth[1].benefits[age]!, 0) / 12;
      found.push({ kind: "btpWealth", cost, lifespan, savingsCapital: balanced.savingsBalance[par - 1] ?? 0, perHundred });
    }
  }

  // Which occupational payout period pays the most up to the expected lifespan.
  if (runs.payouts.length > 1) {
    const totals = runs.payouts.map((r) => ({ run: r, total: totalTo(r, lifespan) }));
    const best = totals.reduce((a, b) => (b.total > a.total + 0.5 ? b : a));
    const worst = totals.reduce((a, b) => (b.total < a.total - 0.5 ? b : a));
    if (best.total - worst.total >= NOTEWORTHY) {
      const shortest = runs.payouts[0]!;
      const life = runs.payouts[runs.payouts.length - 1]!;
      const benefitsGap = sumBetween(best.run.benefits, par, lifespan) - sumBetween(life.benefits, par, lifespan);
      found.push({
        kind: "payout",
        lifespan,
        best: best.run.combo.occupational,
        worst: worst.run.combo.occupational,
        bestTotal: best.total,
        worstTotal: worst.total,
        shortest: shortest.combo.occupational,
        shortestUntil: best.run === shortest ? undefined : overtakenAt(shortest, best.run),
        lifelongFrom: best.run.combo.occupational === LIFELONG ? undefined : overtakenAt(best.run, life),
        housingDriven: best.run !== life && benefitsGap > 0.5 * Math.abs(totalTo(best.run, lifespan) - totalTo(life, lifespan)),
      });
    }
  }

  // The first year is taxed harder than the next.
  if (par + 1 <= LAST_AGE) {
    const gross = balanced.gross[par + 1]!;
    const first = balanced.netAt(par, gross);
    const next = balanced.netAt(par + 1, gross);
    const slice = balanced.occupational[par + 1]! + balanced.individual[par + 1]!;
    if (gross > 0 && next - first >= NOTEWORTHY) {
      const marginal = (age: number) =>
        slice > 0 ? 1 - (balanced.netAt(age, gross) - balanced.netAt(age, gross - slice)) / slice : 0;
      const saving = slice * (marginal(par) - marginal(par + 1));
      found.push({
        kind: "firstYear",
        year: input.born + par,
        gross,
        rateFirst: 1 - first / gross,
        rateNext: 1 - next / gross,
        gain: next - first,
        slice,
        saving,
        bridge: slice * (1 - marginal(par)),
        savingsAtRetirement: balanced.savingsBalance[par - 1] ?? 0,
      });
    }
  }

  // The ISK tax, where the savings are an ISK, a KF or a depot insurance.
  if (hasSavingsOf(input, "isk")) {
    const total = sumBetween(balanced.iskTax, 0, lifespan);
    if (total >= NOTEWORTHY) {
      found.push({
        kind: "iskTax",
        perYear: balanced.iskTax[par] ?? 0,
        total,
        lifespan,
        capital: balanced.iskBalance[par - 1] ?? 0,
      });
    }
  }

  // State tax from the front-loaded scenario's largest years.
  if (runs.front.disposable.some((v, age) => Math.abs(v - balanced.disposable[age]!) > 0.5)) {
    const total = sumBetween(runs.front.stateTax, par, LONG_LIFE);
    if (total >= NOTEWORTHY) {
      const clean = runs.payouts.find((r) => r.combo.occupational !== LIFELONG && sumBetween(r.stateTax, par, LONG_LIFE) < NOTEWORTHY && r.combo.occupational > input.frontYears);
      found.push({ kind: "stateTax", total, years: input.frontYears, avoidWith: clean?.combo.occupational });
    }
  }

  // Spending the savings early, the pensions kept for life.
  if (runs.hybrid !== undefined) {
    const early = (r: ComboRun) => average(r.disposable, par, par + EARLY_YEARS - 1) / 12;
    const monthlyEarly = early(runs.hybrid) - early(balanced);
    if (monthlyEarly >= 100) {
      found.push({
        kind: "hybrid",
        years: input.frontYears,
        monthlyEarly,
        gainBy80: totalTo(runs.hybrid, 80) - totalTo(balanced, 80),
        costByLongLife: totalTo(balanced, LONG_LIFE) - totalTo(runs.hybrid, LONG_LIFE),
      });
    }
  }

  return INSIGHT_ORDER.flatMap((kind) => found.filter((i) => i.kind === kind)).slice(0, 3);
}

/** At most five steps, the ones with a date first. */
function buildChecklist(
  insights: readonly Insight[],
  scenarios: readonly Scenario[],
  recommended: ScenarioId,
  input: StrategyInput,
  retirementYear: number,
  today: Date,
  tax: TaxPicture,
  lifespan: number,
): ChecklistItem[] {
  const items: { priority: number; item: ChecklistItem }[] = [];
  const thisYear = today.getFullYear();

  // Jämkning: for a pensioner, or someone who retires next year, it is this autumn's.
  const applyIn = Math.max(thisYear, retirementYear - 1);
  const known = JAMKNING_DEADLINES[applyIn];
  const deadline = known === undefined ? undefined : new Date(applyIn, known.month - 1, known.day, 23, 59);
  const before = deadline === undefined ? today.getMonth() <= 10 : today <= deadline;
  const when = applyIn > thisYear ? "later" : before ? "soon" : "late";
  items.push({
    priority: when === "later" ? 7 : 0,
    item: {
      kind: "jamkning",
      year: applyIn + 1,
      deadline: known === undefined ? undefined : `${known.day}-${known.month}-${applyIn}`,
      when,
    },
  });
  items.push({ priority: 1, item: { kind: "minpension" } });
  if (tax.below.length > 0) {
    items.push({
      priority: 1.5,
      item: {
        kind: "taxFree",
        from: tax.below[0]!,
        to: tax.below[tax.below.length - 1]!,
        amount: tax.taxFreeIncome,
        year: tax.year,
        topUp: tax.canTopUp,
      },
    });
  }

  for (const insight of insights) {
    if (insight.kind === "payout") {
      items.push({
        priority: 2,
        item: { kind: "payout", best: insight.best, worst: insight.worst, difference: insight.bestTotal - insight.worstTotal, lifespan: insight.lifespan },
      });
    } else if (insight.kind === "firstYear" && insight.saving >= 500) {
      items.push({
        priority: 3,
        item: { kind: "start", year: insight.year + 1, saving: insight.saving, bridge: insight.bridge, covered: insight.savingsAtRetirement >= insight.bridge },
      });
    } else if (insight.kind === "iskTax") {
      items.push({ priority: 5, item: { kind: "isk", perYear: insight.perYear } });
    }
  }
  const chosen = scenarios.find((s) => s.id === recommended)!;
  if (input.housingSupplement) {
    let fromAge = -1;
    for (let a = chosen.run.par; a <= Math.min(LONG_LIFE, LAST_AGE); a += 1) {
      if (chosen.run.benefits[a]! / 12 >= 100) {
        fromAge = a;
        break;
      }
    }
    if (fromAge >= 0) {
      // What it comes to a month, on average over the years it is paid up to the expected lifespan.
      const until = Math.max(lifespan, fromAge + 1);
      let sum = 0;
      let years = 0;
      for (let a = fromAge; a < Math.min(until, LAST_AGE + 1); a += 1) {
        if (chosen.run.benefits[a]! > 0) {
          sum += chosen.run.benefits[a]!;
          years += 1;
        }
      }
      items.push({
        priority: 4,
        item: { kind: "btp", monthly: years > 0 ? sum / years / 12 : 0, fromAge, atStart: fromAge <= chosen.run.par + 1 },
      });
    }
  }
  items.push({ priority: 6, item: { kind: "scenario", id: recommended, years: input.frontYears, combo: chosen.run.combo, kinds: comboKinds(chosen.run) } });

  return items
    .sort((a, b) => a.priority - b.priority)
    .slice(0, 5)
    .map((entry) => entry.item);
}

// ------------------------------------------------------------------ text ---

/** "10 år", "livsvarigt". */
export function payoutText(years: number, lang: Lang): string {
  return years === LIFELONG ? say(lang, "livsvarigt", "for life") : say(lang, `${years} år`, `${years} years`);
}

export function scenarioName(id: ScenarioId, lang: Lang): string {
  switch (id) {
    case "balanced":
      return say(lang, "Balanserad", "Balanced");
    case "front":
      return say(lang, "Din plan", "Your plan");
    case "optimal":
      return say(lang, "Optimalt", "Optimal");
  }
}

/** Which of the three pots there is anything in. */
export interface PotKinds {
  readonly occupational: boolean;
  readonly pension: boolean;
  readonly savings: boolean;
}

/** "tjänstepension 10 år, privat pension livsvarigt, sparande 5 år": the pots a scenario draws, with their periods. */
export function comboText(combo: Combo, kinds: PotKinds, lang: Lang): string {
  const parts = [
    ...(kinds.occupational ? [`${say(lang, "tjänstepension", "occupational pension")} ${payoutText(combo.occupational, lang)}`] : []),
    ...(kinds.pension ? [`${say(lang, "privat pension", "private pension")} ${payoutText(combo.pension, lang)}`] : []),
    ...(kinds.savings ? [`${say(lang, "sparande", "savings")} ${payoutText(combo.savings, lang)}`] : []),
  ];
  return parts.join(", ");
}

/** Which pots a run pays anything from after the first year: what `comboText` names. */
export function comboKinds(run: ComboRun): PotKinds {
  const paid = (values: readonly number[]) => values.some((v, age) => age > run.par && v > 0);
  return { occupational: paid(run.occupational), pension: paid(run.individual), savings: paid(run.savings) };
}

/**
 * What a scenario does, in a sentence. The optimal one says what it chose (`combo`, with the pots there are,
 * `kinds`); without them it says how it chooses.
 */
export function scenarioDescription(
  id: ScenarioId,
  years: number,
  lang: Lang,
  combo?: Combo,
  kinds?: PotKinds,
): string {
  switch (id) {
    case "balanced":
      return say(
        lang,
        "Tjänstepension, privat pension och sparande livsvarigt, liksom den allmänna pensionen.",
        "Occupational pension, private pension and savings for life, like the public pension.",
      );
    case "front":
      return say(
        lang,
        `Uttag på kortare tid, som du valt under Din plan: tjänstepension, privat pension och sparande på ${years} år. Mer de första åren, sedan bara den allmänna pensionen, som alltid är livsvarig.`,
        `Shorter payout, as you chose under Your plan: occupational pension, private pension and savings over ${years} years. More in the first years, then only the public pension, which is always for life.`,
      );
    case "optimal": {
      const chosen = combo === undefined || kinds === undefined ? "" : comboText(combo, kinds, lang);
      const how = say(
        lang,
        "det sätt som ger mest att leva på totalt fram till din förväntade livslängd, utan att någon månad hamnar under din lägsta nivå. Den allmänna pensionen är alltid livsvarig.",
        "the way that leaves the most to live on in all up to your expected lifespan, with no month below your lowest level. The public pension is always for life.",
      );
      return chosen === ""
        ? capitalise(how)
        : say(lang, `${capitalise(chosen)}: ${how}`, `${capitalise(chosen)}: ${how}`);
    }
  }
}

const kr = (value: number, l: Lang) => `${kronor(Math.round(value), l)} ${say(l, "kr", "SEK")}`;
const perMonth = (value: number, l: Lang) => `${kronor(Math.round(value), l)} ${say(l, "kr/mån", "SEK/month")}`;
const age = (value: number, l: Lang) => say(l, `${value} år`, `${value}`);
/** About so many kronor: to the hundred. */
const about = (value: number, l: Lang) => kr(Math.round(value / 100) * 100, l);
/** "vid 63 år", "vid 63–66 år". */
const ageRange = (from: number, to: number, l: Lang) =>
  from === to ? say(l, `vid ${from} år`, `at ${from}`) : say(l, `vid ${from}–${to} år`, `at ${from}–${to}`);
const capitalise = (text: string) => text.charAt(0).toUpperCase() + text.slice(1);

/** A scenario's description, the optimal one's with what it chose. */
export function scenarioSummary(s: Scenario, years: number, lang: Lang): string {
  return scenarioDescription(s.id, years, lang, s.run.combo, comboKinds(s.run));
}

/** An insight as the reader sees it: a headline, what the numbers are, and what to do. */
export function insightText(insight: Insight, lang: Lang): { readonly title: string; readonly body: string; readonly action: string } {
  const l = lang;
  switch (insight.kind) {
    case "btpWealth":
      return {
        title: say(l, "Sparpengarna kostar dig bostadstillägg", "Your savings cost you housing supplement"),
        body: say(
          l,
          `Bostadstillägget prövas mot det du har sparat. Ditt sparande gör att du får ${kr(insight.cost, l)} mindre i bostadstillägg fram till ${age(insight.lifespan, l)}.${insight.perHundred > 0 ? ` Varje 100 000 kr över 100 000 kr sänker tillägget med ungefär ${perMonth(insight.perHundred, l)} – det motsvarar ${percent((insight.perHundred * 12) / 100_000, l)} av pengarna per år, mer än de ger i avkastning.` : ""}`,
          `The housing supplement is tested against what you have saved. Your savings mean ${kr(insight.cost, l)} less in housing supplement up to ${age(insight.lifespan, l)}.${insight.perHundred > 0 ? ` Each 100,000 SEK over 100,000 SEK lowers it by about ${perMonth(insight.perHundred, l)} – that is ${percent((insight.perHundred * 12) / 100_000, l)} of the money a year, more than it earns.` : ""}`,
        ),
        action: say(
          l,
          "Använd sparpengarna tidigt till sådant du ändå vill göra. Bostaden du bor i räknas inte som förmögenhet.",
          "Spend the savings early on what you want to do anyway. The home you live in does not count as wealth.",
        ),
      };
    case "payout": {
      const best = payoutText(insight.best, l);
      const worst = payoutText(insight.worst, l);
      const reasons: string[] = [];
      if (insight.housingDriven) {
        reasons.push(
          say(
            l,
            "Det är bostadstillägget som gör skillnaden: det kan inte bli lägre än noll de år tjänstepensionen betalas ut, och blir högre när den har tagit slut.",
            "The housing supplement makes the difference: it cannot fall below zero in the years the occupational pension is paid, and is higher once it has ended.",
          ),
        );
      }
      if (insight.shortestUntil !== undefined) {
        reasons.push(
          say(
            l,
            `Uttag på ${payoutText(insight.shortest, l)} ger bara mer om du dör före ${age(insight.shortestUntil, l)}.`,
            `A ${payoutText(insight.shortest, l)} payout only pays more if you die before ${age(insight.shortestUntil, l)}.`,
          ),
        );
      }
      if (insight.best !== LIFELONG) {
        reasons.push(
          insight.lifelongFrom === undefined
            ? say(l, "Livsvarigt hinner inte ikapp före 100 år.", "For life does not catch up before 100.")
            : say(l, `Livsvarigt går om först vid ${age(insight.lifelongFrom, l)}.`, `For life catches up only at ${age(insight.lifelongFrom, l)}.`),
        );
      }
      return {
        title: say(l, `Tjänstepension på ${best} ger mest`, `An occupational pension over ${best} pays the most`),
        body: say(
          l,
          `Fram till din förväntade livslängd, ${age(insight.lifespan, l)}, ger ${best} ${kr(insight.bestTotal - insight.worstTotal, l)} mer att leva på än ${worst}. ${reasons.join(" ")}`,
          `Up to your expected lifespan, ${age(insight.lifespan, l)}, ${best} leaves ${kr(insight.bestTotal - insight.worstTotal, l)} more to live on than ${worst}. ${reasons.join(" ")}`,
        ).trim(),
        action: say(
          l,
          `Välj uttagstid när du ansöker om tjänstepensionen, och väg in hur länge du själv tror att du lever.`,
          `Choose the payout period when you apply for the occupational pension, and weigh how long you yourself expect to live.`,
        ),
      };
    }
    case "firstYear":
      return {
        title: say(l, "Första pensionsåret är det dyraste i skatt", "The first year of the pension is the most heavily taxed"),
        body: say(
          l,
          `På samma pension, ${perMonth(insight.gross / 12, l)} före skatt, får du ${perMonth(insight.gain / 12, l)} mer efter skatt året efter än det första (skatt ${percent(insight.rateFirst, l)} mot ${percent(insight.rateNext, l)}). Det högre grundavdraget gäller först från året efter att du fyllt riktåldern.`,
          `On the same pension, ${perMonth(insight.gross / 12, l)} before tax, you keep ${perMonth(insight.gain / 12, l)} more after tax the year after than in the first (tax ${percent(insight.rateFirst, l)} against ${percent(insight.rateNext, l)}). The higher basic deduction applies only from the year after you reach the riktålder.`,
        ),
        action:
          insight.saving >= 500
            ? say(
                l,
                `Låt tjänstepensionen och den privata pensionen börja 1 januari ${insight.year + 1} och lev ${insight.year} på sparpengar (ungefär ${kr(insight.bridge, l)} efter skatt): det sparar ungefär ${kr(insight.saving, l)} i skatt.`,
                `Start the occupational and private pension on 1 January ${insight.year + 1} and live on savings in ${insight.year} (about ${kr(insight.bridge, l)} after tax): it saves about ${kr(insight.saving, l)} in tax.`,
              )
            : say(
                l,
                `Räkna med det i budgeten för ${insight.year}: nettot stiger året efter utan att pensionen ändras.`,
                `Allow for it in the budget for ${insight.year}: what is left after tax rises the year after without the pension changing.`,
              ),
      };
    case "iskTax":
      return {
        title: say(l, "ISK-skatten kan bli lägre", "The ISK tax can be lowered"),
        body: say(
          l,
          `Med ${kr(insight.capital, l)} på ISK när du går i pension blir skatten ungefär ${kr(insight.perYear, l)} om året, ${kr(insight.total, l)} fram till ${age(insight.lifespan, l)}. De första 300 000 kronorna är skattefria (från 2026), för var och en.`,
          `With ${kr(insight.capital, l)} in an ISK at retirement, the tax is about ${kr(insight.perYear, l)} a year, ${kr(insight.total, l)} up to ${age(insight.lifespan, l)}. The first 300,000 SEK are tax-free (from 2026), for each person.`,
        ),
        action: say(
          l,
          "Är ni två: fördela sparandet så att var och en har högst 300 000 kr på ISK. Behöver du ta av sparandet, ta då i första hand av det som ligger över 300 000 kr.",
          "If you are two: split the savings so that each has at most 300,000 SEK in an ISK. When you take from the savings, take first from what lies above 300,000 SEK.",
        ),
      };
    case "stateTax":
      return {
        title: say(l, "Kort uttag kan ge statlig skatt", "A short payout can bring state tax"),
        body: say(
          l,
          `Tas allt ut på ${payoutText(insight.years, l)} blir inkomsten så hög de åren att du betalar ${kr(insight.total, l)} i statlig skatt.`,
          `Drawing everything over ${payoutText(insight.years, l)} lifts the income so high in those years that you pay ${kr(insight.total, l)} in state tax.`,
        ),
        action:
          insight.avoidWith === undefined
            ? say(l, "Sprid uttagen över fler år eller ta dem livsvarigt.", "Spread the withdrawals over more years or take them for life.")
            : say(
                l,
                `Med tjänstepensionen på ${payoutText(insight.avoidWith, l)} slipper du den.`,
                `With the occupational pension over ${payoutText(insight.avoidWith, l)} you avoid it.`,
              ),
      };
    case "hybrid":
      return {
        title: say(l, "Mer tidigt, nästan gratis", "More early, almost for free"),
        body: say(
          l,
          `Tar du sparpengarna på ${payoutText(insight.years, l)} men pensionerna livsvarigt får du ${perMonth(insight.monthlyEarly, l)} mer de första tio åren och ${kr(insight.gainBy80, l)} mer fram till 80 år. Lever du till 95 kostar det bara ${kr(insight.costByLongLife, l)}, och golvet från pensionerna är detsamma.`,
          `Taking the savings over ${payoutText(insight.years, l)} but the pensions for life gives ${perMonth(insight.monthlyEarly, l)} more in the first ten years and ${kr(insight.gainBy80, l)} more up to 80. If you live to 95 it costs only ${kr(insight.costByLongLife, l)}, and the floor from the pensions is the same.`,
        ),
        action: say(
          l,
          "Planera de första årens resor och projekt med sparpengarna, inte med tjänstepensionen.",
          "Plan the first years' travel and projects with the savings, not with the occupational pension.",
        ),
      };
  }
}

const MONTHS_SV = ["januari", "februari", "mars", "april", "maj", "juni", "juli", "augusti", "september", "oktober", "november", "december"];
const MONTHS_EN = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

function dateText(deadline: string, l: Lang): string {
  const [day, month, year] = deadline.split("-").map(Number) as [number, number, number];
  return l === "sv" ? `${day} ${MONTHS_SV[month - 1]} ${year}` : `${day} ${MONTHS_EN[month - 1]} ${year}`;
}

export function checklistText(item: ChecklistItem, lang: Lang): string {
  const l = lang;
  switch (item.kind) {
    case "jamkning":
      if (item.when === "late") {
        return say(
          l,
          `Ansök om jämkning för ${item.year} hos Skatteverket så snart som möjligt, eller begär extra skatt på pensionen – efter början av november gäller beslutet inte hela året.`,
          `Apply to Skatteverket for jämkning for ${item.year} as soon as you can, or ask for extra tax on the pension – after early November the decision does not hold for the whole year.`,
        );
      }
      return item.when === "soon"
        ? say(
            l,
            `Ansök om jämkning för ${item.year} hos Skatteverket${item.deadline === undefined ? " i början av november" : ` senast ${dateText(item.deadline, l)}`}, eller begär extra skatt på pensionen – flera utbetalare drar ofta för lite skatt tillsammans.`,
            `Apply to Skatteverket for jämkning for ${item.year}${item.deadline === undefined ? " in early November" : ` by ${dateText(item.deadline, l)}`}, or ask for extra tax on the pension – several payers often withhold too little together.`,
          )
        : say(
            l,
            `Ansök om jämkning för ${item.year} i början av november ${item.year - 1}, så att rätt skatt dras från början.`,
            `Apply for jämkning for ${item.year} in early November ${item.year - 1}, so that the right tax is withheld from the start.`,
          );
    case "minpension":
      return say(
        l,
        "Hämta dina saldon på minPension och fyll i dem här: tjänstepensionens avtal och uttagsregler, privat pension och sparande.",
        "Fetch your balances at minPension and enter them here: the occupational pension's agreement and payout rules, private pension and savings.",
      );
    case "taxFree":
      return item.topUp
        ? say(
            l,
            `Ta ut minst ${kr(item.amount, l)} om året (${item.year}) i pension också ${ageRange(item.from, item.to, l)}, till exempel genom att börja ta ut tjänstepensionen eller den privata pensionen tidigare. Så mycket är skattefritt, och utrymmet går förlorat om du inte använder det.`,
            `Draw at least ${kr(item.amount, l)} a year (${item.year}) in pension ${ageRange(item.from, item.to, l)} too, for instance by starting the occupational or private pension earlier. That much is tax-free, and the room is lost if you do not use it.`,
          )
        : say(
            l,
            `${capitalise(ageRange(item.from, item.to, l))} är din pension före skatt under den skattefria årsinkomsten, ${kr(item.amount, l)} (${item.year}). Då betalar du ingen inkomstskatt, men du kan inte heller få ROT, RUT eller ränteavdrag – lägg därför större ROT- och RUT-arbeten till andra år.`,
            `${capitalise(ageRange(item.from, item.to, l))} your pension before tax is under the tax-free annual income, ${kr(item.amount, l)} (${item.year}). You pay no income tax then, but cannot get ROT, RUT or the interest deduction either – so put larger ROT and RUT work in other years.`,
          );
    case "payout":
      return say(
        l,
        `Välj ${payoutText(item.best, l)} för tjänstepensionen när du ansöker – ${kr(item.difference, l)} mer än ${payoutText(item.worst, l)} fram till ${age(item.lifespan, l)}.`,
        `Choose ${payoutText(item.best, l)} for the occupational pension when you apply – ${kr(item.difference, l)} more than ${payoutText(item.worst, l)} up to ${age(item.lifespan, l)}.`,
      );
    case "start":
      return say(
        l,
        `Låt tjänstepensionen börja 1 januari ${item.year} och betala året innan med sparpengar (ungefär ${kr(item.bridge, l)})${item.covered ? "" : " – ditt sparande räcker inte hela vägen"}: ungefär ${kr(item.saving, l)} lägre skatt.`,
        `Start the occupational pension on 1 January ${item.year} and pay for the year before with savings (about ${kr(item.bridge, l)})${item.covered ? "" : " – your savings do not cover all of it"}: about ${kr(item.saving, l)} less tax.`,
      );
    case "btp":
      return item.atStart
        ? say(
            l,
            `Ansök om bostadstillägg hos Pensionsmyndigheten när pensionen börjar – i snitt ungefär ${perMonth(item.monthly, l)} enligt modellen.`,
            `Apply to Pensionsmyndigheten for housing supplement when the pension starts – about ${perMonth(item.monthly, l)} on average in the model.`,
          )
        : say(
            l,
            `Ansök om bostadstillägg hos Pensionsmyndigheten från ${age(item.fromAge, l)}, när inkomsten sjunker – i snitt ungefär ${perMonth(item.monthly, l)} enligt modellen.`,
            `Apply to Pensionsmyndigheten for housing supplement from ${age(item.fromAge, l)}, when the income falls – about ${perMonth(item.monthly, l)} on average in the model.`,
          );
    case "isk":
      return say(
        l,
        `Fördela ISK-sparandet så att var och en har högst 300 000 kr: ISK-skatten, ungefär ${kr(item.perYear, l)} om året, minskar eller försvinner.`,
        `Split the ISK savings so that each person has at most 300,000 SEK: the ISK tax, about ${kr(item.perYear, l)} a year, falls or disappears.`,
      );
    case "scenario":
      return say(
        l,
        `Bestäm uttagsplan: ${scenarioName(item.id, l).toLowerCase()} – ${scenarioDescription(item.id, item.years, l, item.combo, item.kinds).toLowerCase()}`,
        `Settle the withdrawal plan: ${scenarioName(item.id, l).toLowerCase()} – ${scenarioDescription(item.id, item.years, l, item.combo, item.kinds).toLowerCase()}`,
      );
  }
}

export type TaxRule = "taxFree" | "rotRut" | "interest" | "reductions";
export const TAX_RULES: readonly TaxRule[] = ["taxFree", "rotRut", "interest", "reductions"];

/** A tax rule as the reader sees it: the rule, what it comes to for this person, and Skatteverket's page on it. */
export function taxRuleText(
  rule: TaxRule,
  tax: TaxPicture,
  input: StrategyInput,
  lang: Lang,
): { readonly title: string; readonly body: string; readonly mine: string; readonly link: { readonly href: string; readonly label: string } } {
  const l = lang;
  const plan = scenarioName(tax.scenario, l);
  const page = (href: string, name: string) => ({ href, label: say(l, `Skatteverket: ${name}`, `Skatteverket (in Swedish): ${name}`) });
  switch (rule) {
    case "taxFree": {
      const raised = tax.pensionerTaxFree.amount > tax.taxFreeFixed + 1_000;
      const lowest = tax.lowestTaxable;
      const below = tax.below;
      return {
        title: say(l, "Ta alltid ut minst den skattefria nivån", "Always draw at least the tax-free amount"),
        body: say(
          l,
          `Skattefri årsinkomst ${tax.year} är ${kr(tax.taxFreeIncome, l)}, 42,3 % av prisbasbeloppet. Upp till den betalar du ingen inkomstskatt och behöver inte deklarera, och det är också gränsen för att en arbetsinkomst ska ge allmän pension. Ta därför ut minst så mycket i pension varje år, också år då du lever på sparpengar – utrymmet går inte att spara till ett annat år.${raised ? ` Som pensionär är ännu mer skattefritt: från ${age(tax.pensionerTaxFree.age, l)} ger det högre grundavdraget ungefär ${kr(tax.pensionerTaxFree.amount, l)} om året utan skatt, enligt modellen.` : ""}`,
          `The tax-free annual income for ${tax.year} is ${kr(tax.taxFreeIncome, l)}, 42.3 % of the price base amount. Up to it you pay no income tax and need not file a return, and it is also the line for income from work to earn a public pension. So draw at least that much in pension every year, also in years you live on savings – the room cannot be saved for another year.${raised ? ` As a pensioner even more is tax-free: from ${age(tax.pensionerTaxFree.age, l)} the higher basic deduction leaves about ${kr(tax.pensionerTaxFree.amount, l)} a year untaxed, in the model.` : ""}`,
        ),
        mine:
          below.length === 0
            ? say(
                l,
                `I förslaget (${plan}) är din pension före skatt som lägst ${about(lowest.amount, l)} om året, vid ${age(lowest.age, l)}. Det är över nivån, så det skattefria utrymmet används varje år.`,
                `In the suggestion (${plan}) your pension before tax is at least ${about(lowest.amount, l)} a year, at ${age(lowest.age, l)}. That is above the amount, so the tax-free room is used every year.`,
              )
            : say(
                l,
                `I förslaget (${plan}) är din pension före skatt under nivån ${ageRange(below[0]!, below[below.length - 1]!, l)}, som lägst ${about(lowest.amount, l)} om året.${tax.canTopUp ? " Ta ut tjänstepension eller privat pension tidigare, så att du når nivån de åren." : " Då betalar du ingen inkomstskatt de åren, men kan inte heller få ROT, RUT eller ränteavdrag."}`,
                `In the suggestion (${plan}) your pension before tax is under the amount ${ageRange(below[0]!, below[below.length - 1]!, l)}, at least ${about(lowest.amount, l)} a year.${tax.canTopUp ? " Draw occupational or private pension earlier, so that you reach the amount in those years." : " You pay no income tax in those years, but cannot get ROT, RUT or the interest deduction either."}`,
              ),
        link: page(SKATTEVERKET.amounts, "Belopp och procent"),
      };
    }
    case "rotRut": {
      const { taxHigh: high, taxLow: low, untaxed } = tax;
      let mine: string;
      if (untaxed.length > 0) {
        const when = ageRange(untaxed[0]!, untaxed[untaxed.length - 1]!, l);
        const most =
          high.amount >= 1
            ? say(l, ` Som mest betalar du ungefär ${about(high.amount, l)} om året, vid ${age(high.age, l)}.`, ` At most you pay about ${about(high.amount, l)} a year, at ${age(high.age, l)}.`)
            : "";
        mine = say(
          l,
          `I förslaget (${plan}) betalar du ingen skatt på pensionen ${when}, och då kan du inte få ROT, RUT eller ränteavdrag.${most}`,
          `In the suggestion (${plan}) you pay no tax on the pension ${when}, so you cannot get ROT, RUT or the interest deduction then.${most}`,
        );
      } else if (high.age === low.age) {
        mine = say(
          l,
          `I förslaget (${plan}) betalar du ungefär ${about(high.amount, l)} i skatt på pensionen om året. Så mycket kan ROT, RUT och ränteavdrag som mest bli tillsammans.`,
          `In the suggestion (${plan}) you pay about ${about(high.amount, l)} in tax on the pension a year. That is the most ROT, RUT and the interest deduction can come to together.`,
        );
      } else {
        const plan2 = low.amount < 0.8 * high.amount;
        mine = say(
          l,
          `I förslaget (${plan}) betalar du ungefär ${about(high.amount, l)} i skatt på pensionen om året vid ${age(high.age, l)} och ${about(low.amount, l)} vid ${age(low.age, l)}. Så mycket kan ROT, RUT och ränteavdrag som mest bli tillsammans de åren.${plan2 ? " Gör därför större ROT-arbeten de år skatten är som högst." : ""}`,
          `In the suggestion (${plan}) you pay about ${about(high.amount, l)} in tax on the pension a year at ${age(high.age, l)} and ${about(low.amount, l)} at ${age(low.age, l)}. That is the most ROT, RUT and the interest deduction can come to together in those years.${plan2 ? " So do larger ROT work in the years the tax is highest." : ""}`,
        );
      }
      return {
        title: say(l, "ROT och RUT kräver att du betalar skatt", "ROT and RUT need tax to take off"),
        body: say(
          l,
          `ROT- och RUT-avdraget är en skattereduktion. Det kan bara minska skatt som du faktiskt betalar, och högst ${kr(tax.rotRutMax, l)} per person och år (${tax.year}), varav ROT högst ${kr(tax.rotMax, l)}. Har du ingen inkomst av tjänst, som pension eller lön, över den skattefria nivån får du inget avdrag alls.`,
          `ROT and RUT are tax reductions. They can only take off tax you actually pay, and at most ${kr(tax.rotRutMax, l)} per person and year (${tax.year}), of which ROT at most ${kr(tax.rotMax, l)}. With no income from work or pension above the tax-free amount you get no deduction at all.`,
        ),
        mine,
        link: page(SKATTEVERKET.rotRut, "ROT- och RUT-arbete"),
      };
    }
    case "interest": {
      let mine: string;
      if (tax.hasIsk && tax.iskIncome >= 100) {
        mine = say(
          l,
          `Om ditt sparande ligger på ISK är schablonintäkten ungefär ${about(tax.iskIncome, l)} det första året. Räntor upp till det beloppet kvittas direkt mot den, och bara resten minskar skatten på pensionen.`,
          `If your savings are in an ISK, the notional income is about ${about(tax.iskIncome, l)} the first year. Interest up to that amount is set against it directly, and only the rest takes tax off the pension.`,
        );
      } else if (tax.hasIsk) {
        mine = say(
          l,
          `Ditt sparande ger nästan ingen schablonintäkt, eftersom de första ${kr(ISK_TAX_FREE, l)} är skattefria. Därför minskar räntan i stället skatten på pensionen.`,
          `Your savings give almost no notional income, as the first ${kr(ISK_TAX_FREE, l)} are tax-free, so interest takes tax off the pension instead.`,
        );
      } else if (hasSavingsOf(input, "depot")) {
        mine = say(
          l,
          "Utdelningar och vinster i din depå är också inkomst av kapital, och räntan kvittas först mot dem. Bara resten minskar skatten på pensionen.",
          "Dividends and gains in your securities account are income from capital too, and the interest is set against them first. Only the rest takes tax off the pension.",
        );
      } else {
        mine = say(
          l,
          "Du har inget sparande med inkomst av kapital i planen, så räntan minskar skatten på pensionen – så länge du betalar skatt.",
          "There are no savings with income from capital in your plan, so interest takes tax off the pension – as long as you pay tax.",
        );
      }
      return {
        title: say(l, "Ränteavdrag: först mot ISK, sedan mot pensionen", "Interest: against an ISK first, then the pension"),
        body: say(
          l,
          `Räntor på bolån och andra lån som uppfyller kraven på säkerhet dras först av mot inkomst av kapital, till exempel schablonintäkten på ett ISK. Det som blir över minskar skatten på pension eller lön. Är din enda inkomst det skattefria beloppet, högst ${kr(tax.taxFreeIncome, l)} (${tax.year}), finns ingen skatt att minska och du får inget ränteavdrag. Har du ett ISK kan räntan ändå kvittas mot schablonintäkten. Det gäller inte kapitalförsäkring, eftersom det är försäkringsbolaget som äger tillgångarna.`,
          `Interest on mortgages and other loans that meet the security requirements is set first against income from capital, such as an ISK's notional income (schablonintäkt). What is left over takes tax off a pension or salary. If your only income is the tax-free amount, at most ${kr(tax.taxFreeIncome, l)} (${tax.year}), there is no tax to take off and no interest deduction. With an ISK the interest can still be set against its notional income. That does not apply to an endowment insurance (kapitalförsäkring), as the insurer owns the assets.`,
        ),
        mine,
        link: page(SKATTEVERKET.interest, "Avdrag för ränteutgifter"),
      };
    }
    case "reductions":
      return {
        title: say(l, "Andra skattereduktioner", "Other tax reductions"),
        body: say(
          l,
          "Samma sak gäller andra skattereduktioner: de kan bara minska skatt som du faktiskt betalar. Skatteverket har en lista över alla, med belopp och villkor.",
          "The same holds for other tax reductions: they can only take off tax you actually pay. Skatteverket lists them all, with amounts and conditions.",
        ),
        mine: "",
        link: page(SKATTEVERKET.reductions, "Skattereduktioner"),
      };
  }
}

// --------------------------------------------------------------- payouts ---

/** The kinds of money a scenario pays, in the order the timeline lists them. */
export type PayoutKind = "income" | "premium" | "guarantee" | "occupational" | "individual" | "savings" | "benefits";
export const PAYOUT_KINDS: readonly PayoutKind[] = ["income", "premium", "guarantee", "occupational", "individual", "savings", "benefits"];

/** One stretch of years a kind of money is paid, ages `from` to `to` inclusive, and what it pays a month on average. */
export interface PayoutPhase {
  readonly kind: PayoutKind;
  readonly from: number;
  readonly to: number;
  /** Still paid at `LONG_LIFE`: for life, as far as the tab reads. */
  readonly lifelong: boolean;
  readonly monthly: number;
}

/**
 * When each kind of money is paid in a run, from the retirement age to `LONG_LIFE`: the stretches of years it is more
 * than a krona a month. Before tax for the pensions, as paid for the savings and the housing supplement.
 */
export function payoutPhases(run: ComboRun): PayoutPhase[] {
  const series: Record<PayoutKind, readonly number[]> = {
    income: run.incomePension,
    premium: run.premiumPension,
    guarantee: run.guaranteePension,
    occupational: run.occupational,
    individual: run.individual,
    savings: run.savings,
    benefits: run.benefits,
  };
  const last = Math.min(LONG_LIFE, LAST_AGE);
  const phases: PayoutPhase[] = [];
  for (const kind of PAYOUT_KINDS) {
    const values = series[kind];
    let from = -1;
    for (let a = run.par; a <= last + 1; a += 1) {
      const paid = a <= last && values[a]! / 12 >= 1;
      if (paid && from < 0) from = a;
      if (!paid && from >= 0) {
        const to = a - 1;
        phases.push({ kind, from, to, lifelong: to === last, monthly: average(values, from, to) / 12 });
        from = -1;
      }
    }
  }
  return phases;
}

/** What the savings' row is named by: the one kind in the plan, or `"mixed"` for both. */
export type SavingsLabel = SavingsKind | "mixed";

export function savingsLabel(input: StrategyInput): SavingsLabel {
  const kinds = savingsByKind(input);
  return kinds.length > 1 ? "mixed" : (kinds[0]?.savingsKind ?? input.savingsKind);
}

export function payoutKindName(kind: PayoutKind, lang: Lang, savingsKind: SavingsLabel = "isk"): string {
  switch (kind) {
    case "income":
      return say(lang, "Inkomstpension", "Income pension");
    case "premium":
      return say(lang, "Premiepension", "Premium pension");
    case "guarantee":
      return say(lang, "Garantipension m.m.", "Guarantee pension etc.");
    case "occupational":
      return say(lang, "Tjänstepension", "Occupational pension");
    case "individual":
      return say(lang, "Privat pension", "Private pension");
    case "savings":
      if (savingsKind === "mixed") return say(lang, "Privat sparande", "Private savings");
      return savingsKind === "depot" ? say(lang, "Depå / fondkonto", "Securities account") : say(lang, "ISK / kapitalförsäkring", "ISK / endowment insurance");
    case "benefits":
      return say(lang, "Bostadstillägg", "Housing supplement");
  }
}

// --------------------------------------------------------------- exports ---

function scenarioHeader(strategy: Strategy, lang: Lang): string[] {
  const first = strategy.scenarios[0]!;
  return [
    say(lang, "Scenario", "Scenario"),
    ...first.phases.map((p) => say(lang, `Per månad ${p.from}–${p.to} år`, `Per month ages ${p.from}–${p.to}`)),
    ...first.totals.map((t) => say(lang, `Totalt till ${t.lifespan} år`, `Total to ${t.lifespan}`)),
    say(lang, "Lägsta månad", "Lowest month"),
    say(lang, "Balanserad går om vid", "Balanced catches up at"),
  ];
}

function scenarioCells(s: Scenario, lang: Lang, years: number): (string | number)[] {
  return [
    `${scenarioName(s.id, lang)} (${scenarioSummary(s, years, lang)})`,
    ...s.phases.map((p) => Math.round(p.monthly)),
    ...s.totals.map((t) => Math.round(t.total)),
    Math.round(s.lowest.monthly),
    s.id === "balanced" ? "" : (s.breakEven ?? say(lang, "inte före 100", "not before 100")),
  ];
}

/** Whole kronor, no thousands separator, as in the other tabs' downloads. */
const whole = (n: number): string => String(Math.round(n));

export function scenariosToCsv(strategy: Strategy, lang: Lang): string {
  const d = delimiterFor(lang);
  const lines = [csvLine(scenarioHeader(strategy, lang), d)];
  for (const s of strategy.scenarios) {
    lines.push(csvLine(scenarioCells(s, lang, strategy.input.frontYears).map((c) => (typeof c === "number" ? whole(c) : c)), d));
  }
  return lines.join("\r\n");
}

export function scenariosToXlsxRows(strategy: Strategy, lang: Lang): SheetData {
  return [
    scenarioHeader(strategy, lang),
    ...strategy.scenarios.map((s) =>
      scenarioCells(s, lang, strategy.input.frontYears).map((c, i) => (typeof c === "number" && i < scenarioHeader(strategy, lang).length - 1 ? kronorCell(c) : c)),
    ),
  ];
}

/** What there is to live on a month, by age, per scenario. */
export function seriesToCsv(strategy: Strategy, lang: Lang): string {
  const d = delimiterFor(lang);
  const lines = [csvLine([say(lang, "Ålder", "Age"), ...strategy.scenarios.map((s) => scenarioName(s.id, lang))], d)];
  for (let a = strategy.par; a <= LONG_LIFE; a += 1) {
    lines.push(csvLine([String(a), ...strategy.scenarios.map((s) => whole(s.run.disposable[a]! / 12))], d));
  }
  return lines.join("\r\n");
}

export function seriesToXlsxRows(strategy: Strategy, lang: Lang): SheetData {
  const rows: SheetData = [[say(lang, "Ålder", "Age"), ...strategy.scenarios.map((s) => scenarioName(s.id, lang))]];
  for (let a = strategy.par; a <= LONG_LIFE; a += 1) {
    rows.push([a, ...strategy.scenarios.map((s) => kronorCell(Math.round(s.run.disposable[a]! / 12)))]);
  }
  return rows;
}
