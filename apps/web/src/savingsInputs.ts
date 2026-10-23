/**
 * The savings inputs "Hur länge räcker pensionen?" and "Hur maxa pensionen?"
 * share: the individual occupational pension and the private pension, under the
 * agreement's drop-down, and the "Privat sparande" group.
 *
 * They are built here, once, for both panels, so that they look the same and
 * work the same by construction -- what the fields are called, how capital and
 * monthly saving sit side by side on one row, the small print under them, what
 * the drop-down of private savings offers and what it asks for. What they mean
 * to the model is `privateSavings.ts`; what a panel does with them is its own.
 *
 * A panel gives the module its grid and its relabel list, a way to read the
 * savings it holds and a way to change them, and calls `appendPension()` where
 * the pairs go (right under the agreement's drop-down) and `appendPrivateSavings()`
 * where the group goes. It calls `sync()` whenever its own state changes, which
 * writes the state into the controls -- setting a value does not fire `change` --
 * and redoes the text that depends on it.
 */
import type { Bounds, Relabel } from "./controls.js";
import { fieldSet } from "./controls.js";
import type { Lang } from "./i18n.js";
import { h, say } from "./pensionAgeUi.js";
import { SAVINGS_KINDS, savingsKindName } from "./privateSavings.js";
import type { SavingsInput } from "./privateSavings.js";

const CAPITAL: Bounds = { min: 0, max: 20_000_000, step: 10_000 };
const SAVING: Bounds = { min: 0, max: 100_000, step: 100 };

export interface SavingsInputsOptions {
  /** What the inputs' `data-setting` names start with: `plCapital`, `paCapital`, ... */
  readonly prefix: string;
  readonly grid: HTMLElement;
  readonly relabels: Relabel[];
  readonly lang: Lang;
  /** The savings the panel holds now. */
  readonly value: () => SavingsInput;
  /** Changes them, and redraws the panel. */
  readonly update: (patch: Partial<SavingsInput>) => void;
  /** The year capital and saving count from, as the engine will use it. */
  readonly savingStart: () => number;
  /** Starts a group of the panel's inputs, as its own `group()` does. */
  readonly group: (sv: string, en: string) => void;
}

export interface SavingsInputs {
  /** The individual occupational pension and the private pension, with what a private pension is. */
  appendPension(): void;
  /** The "Privat sparande" group: what it is held in, how that is taxed, and the capital and monthly saving. */
  appendPrivateSavings(): void;
  /** Writes the savings into the controls and redoes the text that depends on them. */
  sync(lang: Lang): void;
  /** Redoes the text that depends on the savings or on the saving year, without touching a value. */
  relabel(lang: Lang): void;
}

/** One kind of saving's two fields. */
interface Pair {
  readonly capital: { setValue(v: number): void };
  readonly monthly: { setValue(v: number): void };
  relabel(l: Lang): void;
}

export function createSavingsInputs(options: SavingsInputsOptions): SavingsInputs {
  const { prefix, grid, relabels, lang, value, update, savingStart, group } = options;
  const own = fieldSet(grid, relabels, lang);

  /**
   * A line of small print across the grid, under the field it explains; the returned function
   * redoes it. `prompt` marks one that asks the person to do something, which is set in ink.
   */
  function noteLine(text: (l: Lang) => string, prompt: () => boolean = () => false): { relabel: (l: Lang) => void } {
    const element = h("p", "saving-note");
    const relabel = (l: Lang): void => {
      element.textContent = text(l);
      element.classList.toggle("saving-note-prompt", prompt());
    };
    relabel(lang);
    relabels.push(relabel);
    grid.append(element);
    return { relabel };
  }

  /**
   * Capital and monthly saving of one kind of saving, side by side. The grid would
   * put two fields wherever its columns happen to fall; a pair is one row at every
   * width, each caption over its own box, and is relabelled by `caption`.
   */
  function savingPair(
    caption: (l: Lang) => { capital: string; monthly: string },
    capital: { value: number; set: (value: number) => Partial<SavingsInput>; setting: string },
    monthly: { value: number; set: (value: number) => Partial<SavingsInput>; setting: string },
    capitalHint: (l: Lang) => string = (l) => say(l, `kr · värde vid ${savingStart()} års början`, `SEK · value at the start of ${savingStart()}`),
  ): Pair {
    const box = h("div", "field-pair");
    grid.append(box);
    const inBox = fieldSet(box, relabels, lang);
    const capitalField = inBox.number(capital.value, CAPITAL, (v) => update(capital.set(v)));
    capitalField.element.dataset.setting = capital.setting;
    const relabelCapital = inBox.field(capitalField.element, (l) => ({ label: caption(l).capital, hint: capitalHint(l) }), "field field-stack");
    const monthlyField = inBox.number(monthly.value, SAVING, (v) => update(monthly.set(v)));
    monthlyField.element.dataset.setting = monthly.setting;
    const relabelMonthly = inBox.field(
      monthlyField.element,
      (l) => ({ label: caption(l).monthly, hint: say(l, "kr/mån fram till pensionen", "SEK/month until retirement") }),
      "field field-stack",
    );
    /** The captions depend on the drop-down and the hint on the saving year, so they are redone when either changes. */
    const relabel = (l: Lang): void => {
      relabelCapital(l);
      relabelMonthly(l);
    };
    return { capital: capitalField, monthly: monthlyField, relabel };
  }

  let individual: Pair | undefined;
  let privatePension: Pair | undefined;
  let savings: Pair | undefined;
  let kind: { element: HTMLSelectElement } | undefined;
  let kindNote: { relabel: (l: Lang) => void } | undefined;

  const relabel = (l: Lang): void => {
    for (const pair of [individual, privatePension, savings]) pair?.relabel(l);
    kindNote?.relabel(l);
  };

  return {
    appendPension() {
      const v = value();
      individual = savingPair(
        (l) => ({
          capital: say(l, "Individuell tjänstepension, kapital", "Individual occupational pension, capital"),
          monthly: say(l, "Individuell tjänstepension, månadssparande", "Individual occupational pension, monthly saving"),
        }),
        { value: v.individualCapital, set: (individualCapital) => ({ individualCapital }), setting: `${prefix}Capital` },
        { value: v.individualMonthly, set: (individualMonthly) => ({ individualMonthly }), setting: `${prefix}Saving` },
      );
      privatePension = savingPair(
        (l) => ({
          capital: say(l, "Privat pension, kapital", "Private pension, capital"),
          monthly: say(l, "Privat pension, månadssparande", "Private pension, monthly saving"),
        }),
        { value: v.privateCapital, set: (privateCapital) => ({ privateCapital }), setting: `${prefix}PrivCapital` },
        { value: v.privateMonthly, set: (privateMonthly) => ({ privateMonthly }), setting: `${prefix}PrivSaving` },
      );
      noteLine((l) =>
        say(
          l,
          "Privat pension är en pensionsförsäkring eller IPS (individuellt pensionssparande), och beskattas som inkomst när den betalas ut.",
          "Private pension is a pension insurance or an IPS (individual pension saving), and is taxed as income when paid out.",
        ),
      );
    },

    appendPrivateSavings() {
      const v = value();
      // What it is held in decides how it is taxed, and names the two fields.
      group("Privat sparande", "Private savings");
      const select = own.select(
        SAVINGS_KINDS.map((k, index) => ({ value: index, label: (l: Lang) => say(l, k.sv, k.en) })),
        SAVINGS_KINDS.findIndex((k) => k.id === v.savingsKind),
        (index) => update({ savingsKind: SAVINGS_KINDS[index]!.id }),
      );
      select.element.dataset.setting = `${prefix}SavingsKind`;
      relabels.push(select.relabel);
      own.field(select.element, (l) => ({ label: say(l, "Typ av sparande", "Type of saving") }), "field field-wide");
      kind = select;
      // What each kind means for the tax; for a depot the capital is asked for after the tax on its gain.
      kindNote = noteLine(
        (l) =>
          value().savingsKind === "depot"
            ? say(
                l,
                "Vinsten beskattas med 30 % reavinstskatt när du säljer. Ange kapitalvärdet efter reavinstskatt – det du har att leva av.",
                "The gain is taxed at 30% when you sell. Enter the capital value after capital gains tax – what you have to live on.",
              )
            : say(
                l,
                "Schablonbeskattat: skatten tas ut på kapitalet varje år, och utbetalningen är skattefri.",
                "Taxed on a standard return: the tax is taken on the capital each year, and the payout is tax-free.",
              ),
        () => value().savingsKind === "depot",
      );
      savings = savingPair(
        (l) => ({
          capital: `${savingsKindName(value().savingsKind, l)}, ${say(l, "kapital", "capital")}`,
          monthly: `${savingsKindName(value().savingsKind, l)}, ${say(l, "månadssparande", "monthly saving")}`,
        }),
        { value: v.savingsCapital, set: (savingsCapital) => ({ savingsCapital }), setting: `${prefix}SavingsCapital` },
        { value: v.savingsMonthly, set: (savingsMonthly) => ({ savingsMonthly }), setting: `${prefix}SavingsMonthly` },
        (l) =>
          value().savingsKind === "depot"
            ? say(
                l,
                `kr · efter reavinstskatt, värde vid ${savingStart()} års början`,
                `SEK · after capital gains tax, value at the start of ${savingStart()}`,
              )
            : say(l, `kr · värde vid ${savingStart()} års början`, `SEK · value at the start of ${savingStart()}`),
      );
    },

    sync(l) {
      const v = value();
      individual?.capital.setValue(v.individualCapital);
      individual?.monthly.setValue(v.individualMonthly);
      privatePension?.capital.setValue(v.privateCapital);
      privatePension?.monthly.setValue(v.privateMonthly);
      if (kind !== undefined) kind.element.value = String(SAVINGS_KINDS.findIndex((k) => k.id === v.savingsKind));
      savings?.capital.setValue(v.savingsCapital);
      savings?.monthly.setValue(v.savingsMonthly);
      relabel(l);
    },

    relabel,
  };
}
