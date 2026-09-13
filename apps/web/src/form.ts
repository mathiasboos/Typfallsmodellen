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
 * The form is built once and then left alone -- re-creating the fields on every
 * keystroke would take the focus with it. Only the results re-render.
 */
import { options } from "@typfallsmodellen/data";
import { RETIREMENT_AGES } from "@typfallsmodellen/engine";
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

export function createForm(
  initial: TypfallInput,
  lang: Lang,
  onChange: (patch: Partial<TypfallInput>) => void,
): FormHandle {
  const element = document.createElement("form");
  element.className = "panel inputs";
  element.addEventListener("submit", (e) => e.preventDefault());

  const relabels: Relabel[] = [];

  const field = (
    control: HTMLElement,
    relabel: (l: Lang) => { label: string; hint?: string },
    className = "field",
  ) => {
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
  };

  /**
   * A cell you type a number into.
   *
   * `change` rather than `input`: re-running the model on every keystroke would
   * fire on the half-typed "19" of "1960". The value is clamped to the workbook's
   * range and written back, so what the field shows is what the model was given.
   */
  const number = (
    value: number,
    { min, max, step }: { min: number; max: number; step: number },
    apply: (v: number) => void,
  ) => {
    const el = document.createElement("input");
    el.type = "number";
    el.inputMode = "numeric";
    el.value = String(value);
    el.min = String(min);
    el.max = String(max);
    el.step = String(step);
    let current = value;
    el.addEventListener("change", () => {
      const typed = Number(el.value);
      if (!Number.isFinite(typed) || el.value.trim() === "") {
        el.value = String(current);
        return;
      }
      const clamped = Math.min(Math.max(typed, min), max);
      current = clamped;
      el.value = String(clamped);
      apply(clamped);
    });
    return el;
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

  field(
    number(initial.born, { ...BORN, step: 1 }, (born) => onChange({ born })),
    (l) => ({ label: t("birthYear", l), hint: range(BORN) }),
  );

  field(
    number(initial.retirementAge, { ...RETIREMENT, step: 1 }, (retirementAge) =>
      onChange({ retirementAge }),
    ),
    (l) => ({ label: t("retirementAge", l), hint: range(RETIREMENT) }),
  );

  field(
    number(initial.startWorkAge, { ...START_WORK, step: 1 }, (startWorkAge) =>
      onChange({ startWorkAge }),
    ),
    (l) => ({ label: t("startWorkAge", l), hint: range(START_WORK) }),
  );

  // The workbook asks for an annual salary and divides by twelve on the way in
  // (mdlIndataInputOutput.bas:291), so the field is annual here too.
  field(
    number(
      Math.round(initial.monthlySalary * 12),
      { min: 0, max: 100_000_000, step: 1000 },
      (annual) => onChange({ monthlySalary: annual / 12 }),
    ),
    (l) => ({ label: t("annualSalary", l), hint: `${t("kronor", l)} / ${t("years", l)}` }),
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

  return {
    element,
    relabel(l) {
      for (const r of relabels) r(l);
    },
  };
}
