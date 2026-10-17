/**
 * What a person has saved besides the public pension and the collective
 * agreement, and how the model is asked to carry it.
 *
 * Two tabs ask about it the same way -- "Hur länge räcker pensionen?" and "Bästa
 * pensionsålder" -- so the arithmetic is here, and the screen's own half is in
 * `savingsInputs.ts`. Nothing here touches the DOM.
 *
 * Three kinds of saving, each a capital already saved and a monthly saving until
 * retirement, and two pots in the model:
 *
 * - **The individual occupational pension and the private pension** -- the latter a
 *   pension insurance or an IPS, both -- are the model's private saving (IPS,
 *   `privateSavingKind: 0`, taxed as income when paid out, like an occupational
 *   pension). The model treats them alike -- the same return, the same payout
 *   period, the same tax -- so they are one pot and what is saved in each adds up.
 *   The model has no place for the capital (its own `pbhYear` replaces every pot at
 *   once, and garantipension with it), so `ipsOpening` was added to the engine.
 *   Capital and saving count from the year Avancerat calls "Sparandet börjar år" --
 *   the model's idea of today -- or from the retirement year if that has passed.
 *   The monthly saving is a fixed number of kronor, as in the model: it is not
 *   raised with inflation.
 * - **Private savings** are held in one of two kinds, as the money is taxed: an ISK,
 *   a kapitalförsäkring or a depåförsäkring, taxed on a standard return while saved;
 *   or a depå, a VP-konto or a fondkonto, taxed 30% on the gain when it is sold.
 *   Both are paid out tax-free, and the person types the capital after that tax on
 *   the second -- what there is to live on -- while its return is taxed each year
 *   (the engine's `SavingType.Depot`, which the workbook does not have). The model
 *   carries one kind of private saving per run, so they are a second run, and the
 *   two payouts add -- what is tax-free is no part of what the tax is taken from.
 *   They are drawn over the same payout period as everything else; only "lifelong"
 *   differs, since money, unlike an IPS, is not paid for as long as the person
 *   lives: the model spreads it over the expected remaining life and it runs out.
 */
import type { ModelContext } from "@typfallsmodellen/engine";

import type { Lang } from "./i18n.js";

const say = (l: Lang, sv: string, en: string) => (l === "sv" ? sv : en);

/**
 * What private saving can be held in -- two kinds, as the money is taxed, in the order the
 * drop-down offers them.
 *
 * - `isk`: ISK, kapitalförsäkring and depåförsäkring, which are taxed on a standard return
 *   while the money is saved (schablonbeskattade) and paid out tax-free. The model's
 *   ISK (2); a kapitalförsäkring (1) differs from it by a hair, in how the tax base is
 *   taken, so one stands for the group.
 * - `depot`: depå, VP-konto and fondkonto, where the gain is taxed 30% when it is sold.
 *   The model has no such saving (the engine got `SavingType.Depot`, 3, for this): the
 *   capital is taken after that tax, as the person is asked to type it, and the return
 *   is taxed each year, as if the gain were realised then. Paid out tax-free, the tax
 *   being paid.
 */
export const SAVINGS_KINDS = [
  { id: "isk", kind: 2, sv: "ISK / Kapitalförsäkring / Depåförsäkring", en: "ISK / Endowment insurance / Depot insurance" },
  { id: "depot", kind: 3, sv: "Depå / VP-konto / Fondkonto", en: "Securities account / Fund account" },
] as const;

export type SavingsKind = (typeof SAVINGS_KINDS)[number]["id"];

export function savingsKindName(kind: SavingsKind, lang: Lang): string {
  const found = SAVINGS_KINDS.find((k) => k.id === kind) ?? SAVINGS_KINDS[0];
  return say(lang, found.sv, found.en);
}

/** The savings a tab asks for, in kronor. A tab's own input extends this. */
export interface SavingsInput {
  /** Individual occupational pension already saved. */
  readonly individualCapital: number;
  /** Individual occupational pension saved each month until retirement. */
  readonly individualMonthly: number;
  /** Private pension (pension insurance, IPS) already saved. */
  readonly privateCapital: number;
  /** Private pension saved each month until retirement. */
  readonly privateMonthly: number;
  /** What the private saving below is held in. */
  readonly savingsKind: SavingsKind;
  /** Private saving already saved -- after the tax on its gain, for a depot. */
  readonly savingsCapital: number;
  /** Private saving put aside each month until retirement. */
  readonly savingsMonthly: number;
}

/** No saving at all, the ISK being what the drop-down opens on. */
export function defaultSavings(): SavingsInput {
  return {
    individualCapital: 0,
    individualMonthly: 0,
    privateCapital: 0,
    privateMonthly: 0,
    savingsKind: "isk",
    savingsCapital: 0,
    savingsMonthly: 0,
  };
}

/** A saving's capital already saved and what is put aside each month, kronor. */
export interface Saving {
  readonly capital: number;
  readonly monthly: number;
}

const amount = (n: number): number => (Number.isFinite(n) ? Math.max(n, 0) : 0);

/**
 * What is paid out as taxed income: the individual occupational pension and the private
 * pension -- a pension insurance or an IPS. The model has one pot for both (its IPS): the
 * same return, the same payout period, the same tax, so what is saved in each simply adds up.
 */
export function taxedSaving(input: SavingsInput): Saving {
  return {
    capital: amount(input.individualCapital) + amount(input.privateCapital),
    monthly: amount(input.individualMonthly) + amount(input.privateMonthly),
  };
}

/**
 * What is paid out tax-free -- the private savings, whichever of the two kinds they are held
 * in, the tax on them having been paid already or being paid on the way -- and the model's
 * number for the kind; null when there is nothing in them.
 */
export function taxFreeSaving(input: SavingsInput): (Saving & { readonly kind: 2 | 3 }) | null {
  const found = SAVINGS_KINDS.find((k) => k.id === input.savingsKind) ?? SAVINGS_KINDS[0];
  const saving = { capital: amount(input.savingsCapital), monthly: amount(input.savingsMonthly) };
  if (saving.capital <= 0 && saving.monthly <= 0) return null;
  return { ...saving, kind: found.kind };
}

/** What the year the saving counts from depends on in the person. */
export interface SavingPerson {
  readonly born: number;
  readonly startWorkAge: number;
  readonly retirementAge: number;
}

/**
 * The year the capital and the saving count from -- the model's today, or the retirement year if that has passed.
 *
 * Never so early that the run cannot book the capital: it is credited at the end of
 * the year before, which has to be a year the run covers, and the run starts at the
 * model's first age or the age work starts, whichever is earlier. A "Sparandet börjar
 * år" in Avancerat from before the person's time would otherwise drop the capital.
 */
export function savingYear(person: SavingPerson, shared: ModelContext): number {
  const earliest = person.born + Math.min(shared.modelStartAge, person.startWorkAge) + 1;
  return Math.max(Math.min(shared.ipsStart, person.born + person.retirementAge), earliest);
}

/** The two runs the savings need: the one paid out as taxed income, and the one that is not. */
export type SavingPart = "pension" | "savings";

/**
 * The four settings of the model's private saving for one part, to spread into a context.
 *
 * `"pension"` is everything paid out as taxed income (`taxedSaving`): the model's IPS,
 * `privateSavingKind: 0`. `"savings"` is the private savings (`taxFreeSaving`), an ISK or
 * a depot, run on its own because the model carries one kind of private saving per run.
 * Nothing else differs between the two, and the two payouts just add: the tax-free one is
 * no part of what the tax is taken from. A `"savings"` part with nothing in it carries no
 * saving at all.
 *
 * `ipsMonthly` of 1 or less would be read as a share of income (a landmine `advanced.ts`
 * describes), so a saving that small is none.
 */
export function savingSettings(
  shared: ModelContext,
  person: SavingPerson,
  savings: SavingsInput,
  part: SavingPart,
): Pick<ModelContext, "privateSavingKind" | "ipsMonthly" | "ipsStart" | "ipsOpening"> {
  const saving = part === "pension" ? { ...taxedSaving(savings), kind: 0 } : (taxFreeSaving(savings) ?? { capital: 0, monthly: 0, kind: 0 });
  return {
    privateSavingKind: saving.kind,
    ipsMonthly: saving.monthly > 1 ? saving.monthly : 0,
    ipsStart: savingYear(person, shared),
    ipsOpening: saving.capital,
  };
}

/**
 * What a person is told about these savings in the notes under a tab's results: the same two
 * paragraphs in both tabs. `from` is the year the capital and the saving count from. A tab
 * adds to the first what is its own -- how long the pension is paid out for.
 */
export function savingsNotes(lang: Lang, from: number): { readonly pension: string; readonly savings: string } {
  return {
    pension: say(
      lang,
      `Individuell tjänstepension och privat pension (en pensionsförsäkring eller IPS) räknas som modellens privata pensionssparande (IPS) och beskattas som inkomst vid utbetalning, precis som tjänstepension. De behandlas lika och läggs därför ihop. Kapitalet du har, och månadssparandet fram till pensionen, räknas från ${from} (Avancerat > Sparandet börjar år; har pensionen redan börjat räknas kapitalet vid pensionsstart). Månadssparandet är ett fast belopp i kronor och räknas inte upp med inflationen.`,
      `Individual occupational pension and private pension (a pension insurance or an IPS) are counted as the model's private pension saving (IPS) and are taxed as income when paid out, like an occupational pension. They are treated alike and so are added together. The capital you have, and the monthly saving until retirement, count from ${from} (Advanced > Saving starts in; if the pension has already started, the capital counts at the start). The monthly saving is a fixed number of kronor and is not raised with inflation.`,
    ),
    savings: say(
      lang,
      `Privat sparande är pengar som redan är inkomstbeskattade, och betalas ut skattefritt. På ISK, kapitalförsäkring och depåförsäkring tas en schablonskatt ut på kapitalet varje år (modellens ISK). På depå, VP-konto och fondkonto beskattas vinsten med 30 % när du säljer: ange därför kapitalvärdet efter reavinstskatt, det du har att leva av. Avkastningen på det beskattas med 30 % varje år, som om vinsten realiserades då. Månadssparandet sätts av efter skatt.`,
      `Private savings are money that has already been taxed as income, and are paid out tax-free. An ISK, an endowment insurance or a depot insurance pays a standard tax on the capital each year (the model's ISK). On a securities account or a fund account the gain is taxed at 30% when you sell: enter the capital value after capital gains tax, what you have to live on. Its return is taxed at 30% each year, as if the gain were realised then. The monthly saving is put aside after tax.`,
    ),
  };
}
