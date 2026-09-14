/**
 * The Start sheet's input panel.
 *
 * Every choice comes from packages/data/options.json, which is extracted from
 * the workbook: the cohort range, the retirement ages, the eight agreements in
 * the model's own wording, and the defaults a fresh workbook opens with. None
 * of it is retyped here, so next year's extraction updates the form by itself.
 *
 * The numeric variables are cells you type into, the way the workbook's own
 * input cells are. A typed value is clamped to the range the extraction allows
 * and written back into the field, so the form can never show a number the run
 * did not use.
 *
 * The salary field asks for a monthly wage directly rather than the Start
 * sheet's own annual one -- a deliberate departure from `mdlIndataInputOutput.
 * bas:291`, which is Årslön divided by twelve on the way in. `TypfallInput.
 * monthlySalary` is already that monthly figure, so this removes a round trip
 * rather than adding one.
 *
 * The riktålder checkbox mirrors `rng_Riktålder` on the Start sheet
 * (`wsStart.cls:33`, `mdlAlternativePensYear.bas:51`): checked, a birth-year
 * change writes that cohort's riktålder into the retirement age instead of
 * leaving it where it was, and the field is locked to it until unchecked.
 *
 * The form is built once and then left alone -- re-creating the fields on every
 * keystroke would take the focus with it. Only the results re-render.
 */
import { options } from "@typfallsmodellen/data";
import { RETIREMENT_AGES, riktalderFor } from "@typfallsmodellen/engine";
import type { SchemeId, TypfallInput } from "@typfallsmodellen/engine";

import { t } from "./i18n.js";
import type { Lang } from "./i18n.js";

const BIRTH_YEARS = options.ranges.birthYears as readonly number[];
const SCHEMES = options.choices.occupationalPension as readonly { value: number; label: string }[];

/** The workbook's own dropdowns, read as the range they cover. */
const span = (values: readonly number[]): { min: number; max: number } => ({
  min: Math.min(...values),
  max: Math.max(...values),
});

const BORN = span(BIRTH_YEARS);
const RETIREMENT = span(RETIREMENT_AGES);
// `Börjar arbeta vid ålder` has no extracted list; the sheet offers 15 to 40.
const START_WORK = { min: 15, max: 40 };

export interface FormHandle {
  readonly element: HTMLElement;
  /** Re-labels the fields after a language change, keeping their values. */
  relabel(lang: Lang): void;
}

type Relabel = (lang: Lang) => void;

/** A number field, and a way to set its value from outside without losing the
 * invalid-input fallback the field keeps for itself. */
interface NumberField {
  readonly element: HTMLInputElement;
  setValue(v: number): void;
}

/**
 * Explains the three economic-assumption fields: not a SysLang row (there is
 * no "överavkastning" wording on the Start sheet), so -- like the KPI cards'
 * text -- this is genuine per-language copy rather than a `t()` lookup.
 */
function assumptionsNote(l: Lang): { readonly summary: string; readonly paragraphs: readonly string[] } {
  return l === "sv"
    ? {
        summary: "Om pris- och avkastningsantaganden",
        paragraphs: [
          "Pensionsprognosen beräknas i fasta priser, det vill säga med antagandet om 0 procent " +
            "framtida inflation och 0 procent real löneutveckling. Prognosresultatet uttrycks " +
            "därmed i dagens pris- och löneläge, vilket gör det möjligt att jämföra den " +
            "prognostiserade pensionen direkt med dagens lön.",
          "I prognosstandarden antas en avkastning på 1,7 procent. För att pensionens " +
            "prognosvärde ska sättas i relation till den framtida inkomsten uttrycks " +
            "avkastningen som en så kallad överavkastning – det vill säga hur mycket kapitalets " +
            "avkastning överstiger den generella löneutvecklingen.",
        ],
      }
    : {
        summary: "About the price and return assumptions",
        paragraphs: [
          "The pension forecast is calculated in fixed prices, meaning it assumes 0 percent " +
            "future inflation and 0 percent real wage growth. The forecast result is therefore " +
            "expressed at today's price and wage level, which makes it possible to compare the " +
            "forecasted pension directly with today's salary.",
          "The forecast standard assumes a return of 1.7 percent. To relate the pension's " +
            "forecast value to future income, the return is expressed as a so-called excess " +
            "return – that is, how much the return on capital exceeds general wage growth.",
        ],
      };
}

export function createForm(
  initial: TypfallInput,
  lang: Lang,
  onChange: (patch: Partial<TypfallInput>) => void,
): FormHandle {
  const element = document.createElement("form");
  element.className = "panel inputs";
  element.addEventListener("submit", (e) => e.preventDefault());

  const relabels: Relabel[] = [];
  // `lang` is the language the form was built with; the language can change
  // later (main.ts calls `relabel`), and the riktålder checkbox's own handler
  // needs the current one rather than the one it closed over at construction.
  let currentLang = lang;

  const field = (
    control: HTMLElement,
    relabel: (l: Lang) => { label: string; hint?: string },
    className = "field",
  ): ((l: Lang) => void) => {
    const wrap = document.createElement("label");
    wrap.className = className;
    const caption = document.createElement("span");
    caption.className = "field-label";
    const note = document.createElement("span");
    note.className = "field-hint";

    const apply = (l: Lang) => {
      const text = relabel(l);
      caption.textContent = text.label;
      note.textContent = text.hint ?? "";
      note.hidden = text.hint === undefined;
    };
    apply(lang);

    if (className.includes("field-check")) wrap.append(control, caption, note);
    else wrap.append(caption, control, note);
    relabels.push(apply);
    element.append(wrap);
    // Returned so a linked control (the riktålder checkbox) can re-run this
    // field's own hint text without guessing its position in `relabels`.
    return apply;
  };

  /**
   * A cell you type a number into.
   *
   * `change` rather than `input`: re-running the model on every keystroke would
   * fire on the half-typed "19" of "1960". The value is clamped to the workbook's
   * range and written back, so what the field shows is what the model was given.
   *
   * `setValue` lets a linked control (the riktålder checkbox) update the field
   * the same way a typed change would, so a later invalid entry reverts to that
   * value rather than to whatever was there before the link took over.
   */
  const number = (
    value: number,
    { min, max, step }: { min: number; max: number; step: number },
    apply: (v: number) => void,
  ): NumberField => {
    const el = document.createElement("input");
    el.type = "number";
    el.inputMode = "numeric";
    el.min = String(min);
    el.max = String(max);
    el.step = String(step);
    let current = value;
    const setValue = (v: number) => {
      current = v;
      el.value = String(v);
    };
    setValue(value);
    el.addEventListener("change", () => {
      const typed = Number(el.value);
      if (!Number.isFinite(typed) || el.value.trim() === "") {
        el.value = String(current);
        return;
      }
      const clamped = Math.min(Math.max(typed, min), max);
      setValue(clamped);
      apply(clamped);
    });
    return { element: el, setValue };
  };

  const percentField = (value: number, apply: (v: number) => void) => {
    const el = document.createElement("input");
    el.type = "number";
    el.value = String(Math.round(value * 1000) / 10);
    el.step = "0.1";
    el.className = "percent";
    el.addEventListener("change", () => {
      const v = Number(el.value);
      if (Number.isFinite(v)) apply(v / 100);
    });
    return el;
  };

  const range = ({ min, max }: { min: number; max: number }) => `${min}–${max}`;

  // Retirement age is built before birth year appends it, so birth year's
  // change handler can reach it when the riktålder checkbox is locked -- DOM
  // order still follows the sheet's own layout, since that is decided by
  // when `field()` appends each wrapper, not by when the control is built.
  let riktalderLocked = false;
  let born = initial.born;
  const retirement = number(initial.retirementAge, { ...RETIREMENT, step: 1 }, (retirementAge) =>
    onChange({ retirementAge }),
  );

  /** `mdlAlternativePensYear.bas:51`: this cohort's riktålder, applied now. */
  const applyRiktalder = () => {
    const value = riktalderFor(born);
    retirement.setValue(value);
    onChange({ retirementAge: value });
  };

  field(
    number(initial.born, { ...BORN, step: 1 }, (value) => {
      born = value;
      onChange({ born: value });
      if (riktalderLocked) applyRiktalder();
    }).element,
    (l) => ({ label: t("birthYear", l), hint: range(BORN) }),
  );

  const relabelRetirement = field(retirement.element, (l) => ({
    label: t("retirementAge", l),
    // The range hint would be misleading while the checkbox has taken over --
    // the field's own disabled state and the checked box beside it already
    // say why, so this just steps aside rather than repeating "Riktålder".
    ...(riktalderLocked ? {} : { hint: range(RETIREMENT) }),
  }));

  const riktalderBox = document.createElement("input");
  riktalderBox.type = "checkbox";
  riktalderBox.addEventListener("change", () => {
    riktalderLocked = riktalderBox.checked;
    retirement.element.disabled = riktalderLocked;
    if (riktalderLocked) applyRiktalder();
    // The hint switches between the allowed range and naming the lock, so it
    // has to be re-applied along with the checkbox state, not just on a
    // language change.
    relabelRetirement(currentLang);
  });
  field(riktalderBox, (l) => ({ label: t("riktalderCheckbox", l) }), "field field-check");

  field(
    number(initial.startWorkAge, { ...START_WORK, step: 1 }, (startWorkAge) =>
      onChange({ startWorkAge }),
    ).element,
    (l) => ({ label: t("startWorkAge", l), hint: range(START_WORK) }),
  );

  field(
    number(
      Math.round(initial.monthlySalary),
      { min: 0, max: 1_000_000, step: 100 },
      (monthlySalary) => onChange({ monthlySalary }),
    ).element,
    (l) => ({ label: t("monthlySalary", l), hint: `${t("kronor", l)} / ${t("perMonth", l).toLowerCase()}` }),
  );

  const schemeSelect = document.createElement("select");
  for (const choice of SCHEMES) {
    const option = document.createElement("option");
    option.value = String(choice.value);
    option.textContent = choice.label;
    if (choice.value === initial.scheme) option.selected = true;
    schemeSelect.append(option);
  }
  schemeSelect.addEventListener("change", () =>
    onChange({ scheme: Number(schemeSelect.value) as SchemeId }),
  );
  field(schemeSelect, (l) => ({ label: t("occupational", l) }), "field field-wide");

  const marriedBox = document.createElement("input");
  marriedBox.type = "checkbox";
  marriedBox.checked = initial.married;
  marriedBox.addEventListener("change", () => onChange({ married: marriedBox.checked }));
  field(marriedBox, (l) => ({ label: t("married", l) }), "field field-check");

  field(
    percentField(initial.yearlyInflation, (yearlyInflation) => onChange({ yearlyInflation })),
    (l) => ({ label: t("inflation", l), hint: "%" }),
  );

  field(
    percentField(initial.realGrowth, (realGrowth) => onChange({ realGrowth })),
    (l) => ({ label: t("realGrowth", l), hint: "%" }),
  );

  field(
    percentField(initial.realReturn, (realReturn) => onChange({ realReturn })),
    (l) => ({ label: t("realReturn", l), hint: "%" }),
  );

  const assumptionsInfo = document.createElement("details");
  assumptionsInfo.className = "field-note";
  const assumptionsSummary = document.createElement("summary");
  const assumptionsBody = document.createElement("div");
  assumptionsBody.className = "field-note-body";
  const applyAssumptionsNote = (l: Lang) => {
    const note = assumptionsNote(l);
    assumptionsSummary.textContent = note.summary;
    assumptionsBody.replaceChildren(
      ...note.paragraphs.map((text) => {
        const p = document.createElement("p");
        p.textContent = text;
        return p;
      }),
    );
  };
  applyAssumptionsNote(lang);
  assumptionsInfo.append(assumptionsSummary, assumptionsBody);
  relabels.push(applyAssumptionsNote);
  element.append(assumptionsInfo);

  return {
    element,
    relabel(l) {
      currentLang = l;
      for (const r of relabels) r(l);
    },
  };
}
