/**
 * The Start sheet's input panel.
 *
 * Every choice comes from packages/data/options.json, which is extracted from
 * the workbook: the cohort range, the retirement ages, the eight agreements in
 * the model's own wording, and the defaults a fresh workbook opens with. None
 * of it is retyped here, so next year's extraction updates the form by itself.
 *
 * The form is built once and then left alone -- re-creating the fields on every
 * keystroke would take the focus with it. Only the results re-render.
 */
import { options } from "@typfallsmodellen/data";
import { RETIREMENT_AGES, Scheme } from "@typfallsmodellen/engine";
import type { SchemeId, TypfallInput } from "@typfallsmodellen/engine";

import { t } from "./i18n.js";
import type { Lang } from "./i18n.js";

const BIRTH_YEARS = options.ranges.birthYears as readonly number[];
const SCHEMES = options.choices.occupationalPension as readonly { value: number; label: string }[];

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
    labelText: string,
    control: HTMLElement,
    relabel: Relabel,
    hint?: string,
    className = "field",
  ) => {
    const wrap = document.createElement("label");
    wrap.className = className;
    const caption = document.createElement("span");
    caption.className = "field-label";
    caption.textContent = labelText;
    if (className.includes("field-check")) wrap.append(control, caption);
    else wrap.append(caption, control);
    if (hint !== undefined) {
      const note = document.createElement("span");
      note.className = "field-hint";
      note.textContent = hint;
      wrap.append(note);
    }
    relabels.push((l) => relabel(l));
    element.append(wrap);
    return caption;
  };

  const select = (values: readonly number[], selected: number, apply: (v: number) => void) => {
    const el = document.createElement("select");
    for (const v of values) {
      const option = document.createElement("option");
      option.value = String(v);
      option.textContent = String(v);
      if (v === selected) option.selected = true;
      el.append(option);
    }
    el.addEventListener("change", () => apply(Number(el.value)));
    return el;
  };

  const number = (value: number, step: number, apply: (v: number) => void) => {
    const el = document.createElement("input");
    el.type = "number";
    el.value = String(value);
    el.step = String(step);
    el.addEventListener("change", () => {
      const v = Number(el.value);
      if (Number.isFinite(v)) apply(v);
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

  const bornCaption = field(
    t("birthYear", lang),
    select(BIRTH_YEARS, initial.born, (born) => onChange({ born })),
    (l) => (bornCaption.textContent = t("birthYear", l)),
  );

  const parCaption = field(
    t("retirementAge", lang),
    select(RETIREMENT_AGES, initial.retirementAge, (retirementAge) => onChange({ retirementAge })),
    (l) => (parCaption.textContent = t("retirementAge", l)),
  );

  const startCaption = field(
    t("startWorkAge", lang),
    select(
      Array.from({ length: 26 }, (_, i) => 15 + i),
      initial.startWorkAge,
      (startWorkAge) => onChange({ startWorkAge }),
    ),
    (l) => (startCaption.textContent = t("startWorkAge", l)),
  );

  // The workbook asks for an annual salary and divides by twelve on the way in
  // (mdlIndataInputOutput.bas:291), so the field is annual here too.
  const salaryCaption = field(
    t("annualSalary", lang),
    number(Math.round(initial.monthlySalary * 12), 1000, (annual) =>
      onChange({ monthlySalary: annual / 12 }),
    ),
    (l) => (salaryCaption.textContent = t("annualSalary", l)),
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
  const schemeCaption = field(t("occupational", lang), schemeSelect, (l) => {
    schemeCaption.textContent = t("occupational", l);
  });

  const marriedBox = document.createElement("input");
  marriedBox.type = "checkbox";
  marriedBox.checked = initial.married;
  marriedBox.addEventListener("change", () => onChange({ married: marriedBox.checked }));
  const marriedCaption = field(
    t("married", lang),
    marriedBox,
    (l) => {
      marriedCaption.textContent = t("married", l);
    },
    undefined,
    "field field-check",
  );

  const inflationCaption = field(
    t("inflation", lang),
    percentField(initial.yearlyInflation, (yearlyInflation) => onChange({ yearlyInflation })),
    (l) => (inflationCaption.textContent = t("inflation", l)),
    "%",
  );

  const growthCaption = field(
    t("realGrowth", lang),
    percentField(initial.realGrowth, (realGrowth) => onChange({ realGrowth })),
    (l) => (growthCaption.textContent = t("realGrowth", l)),
    "%",
  );

  const returnCaption = field(
    t("realReturn", lang),
    percentField(initial.realReturn, (realReturn) => onChange({ realReturn })),
    (l) => (returnCaption.textContent = t("realReturn", l)),
    "%",
  );

  void Scheme;

  return {
    element,
    relabel(l) {
      for (const r of relabels) r(l);
    },
  };
}
