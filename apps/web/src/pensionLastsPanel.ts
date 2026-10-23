/**
 * "Hur länge räcker pensionen?": the tab that sets a pension against a budget --
 * for a typfall of your own, a collective agreement of your own, and every way of
 * taking the occupational pension out that the agreements offer, 5 to 20 years or
 * for life -- and shows how long the budget lasts under each.
 *
 * `pensionLasts.ts` does the arithmetic (its file comment is where the method and
 * its conventions live: the payout periods, the budget after tax and before, the
 * two readings of "lasts"); this file is the panel around it -- the inputs, the
 * short answer, the two line charts, the tables and the exports. The charts are
 * `pensionAgeChart.ts`'s `renderLineChart`, which this tab only asks for two
 * more things (a dashed line, a label where a payout ends).
 *
 * Like "Hur maxa pensionen?", this tab's inputs are its own: it is not a diff
 * against the Prognos form and does not follow it. "Hämta från Prognos" copies the
 * form's person, agreement and assumptions across once, at the moment of the click.
 * What it does read live is the shared `ModelContext` (Avancerat: tax, wage
 * profile, ...), which `main.ts` hands over through `setContext` on every render --
 * `pensionLasts.ts` says which of those settings it overrides.
 *
 * Results are live: a change that matters to the engine runs the model seventeen
 * times (a few tens of milliseconds) and rebuilds the results area; a change to the
 * budget, the lifespan or before/after tax reuses the runs. The inputs are built
 * once and never rebuilt, so a field you are typing into keeps its focus.
 *
 * **The budget** is a drop-down with minPension's three examples -- the amounts
 * the site's Pensionsöversikt uses -- and "Egen budget", beside the amount itself.
 * Typing an amount that is one of the examples selects it, as Pensionsöversikten's
 * does; choosing "Egen budget" keeps the amount and lets it be changed.
 *
 * **Salary as monthly salary or as final salary** works as in "Hur maxa pensionen?":
 * a typed final salary is turned into the monthly salary behind it by running the
 * model (`monthlySalaryForTarget`), and both numbers are shown beside the field.
 *
 * Every visible string is a plain per-language literal, as in the other new tabs;
 * the field captions that have a SysLang row (birth year, start of work, the
 * agreement, the three assumptions) use it, so they read exactly as in Prognos.
 */
import { BIRTH_YEARS, RETIREMENT_AGES, SCHEME_CHOICES, defaultContext, riktalderFor } from "@typfallsmodellen/engine";
import type { DeathProbabilities, ModelContext, SchemeId, TypfallInput } from "@typfallsmodellen/engine";

import type { Bounds, Relabel } from "./controls.js";
import { fieldSet, rangeHint, span } from "./controls.js";
import { csvExportButton, xlsxExportButton } from "./export.js";
import { kronor, percent as share } from "./format.js";
import { t } from "./i18n.js";
import type { Lang } from "./i18n.js";
import { LIFESPAN_BOUNDS, centreBounds } from "./pensionAge.js";
import { niceTicks, renderLineChart } from "./pensionAgeChart.js";
import type { ChartSeries } from "./pensionAgeChart.js";
import { capital, decimals, h, mkr, panel, row, say, signed, tableOf } from "./pensionAgeUi.js";
import {
  BUDGETS,
  BUDGET_BOUNDS,
  amountSeriesToCsv,
  amountSeriesToXlsxRows,
  amountsToCsv,
  amountsToXlsxRows,
  totalPaid,
  CHART_PAYOUTS,
  LIFELONG,
  balanceText,
  budgetIdOf,
  budgetMatrix,
  coverageText,
  defaultLastsInput,
  evaluate,
  matrixToCsv,
  matrixToXlsxRows,
  monthlySalaryForTarget,
  normaliseLasts,
  optionsToCsv,
  optionsToXlsxRows,
  payoutName,
  runPayouts,
  runsKey,
  seriesToCsv,
  seriesToXlsxRows,
  slotOfPayout,
} from "./pensionLasts.js";
import type { Assessment, Lasts, LastsInput, PayoutOption, PayoutRuns } from "./pensionLasts.js";
import { savingYear, savingsNotes } from "./privateSavings.js";
import { createSavingsInputs } from "./savingsInputs.js";
import { cell, headCell } from "./tables.js";

// The sheet's own ranges, duplicated rather than imported for the same reason
// `pensionAgePanel.ts` duplicates them.
const BORN: Bounds = { ...span(BIRTH_YEARS), step: 1 };
const START_WORK: Bounds = { min: 15, max: 40, step: 1 };
const RETIREMENT: Bounds = { ...span(RETIREMENT_AGES), step: 1 };
const LIFESPAN: Bounds = { ...LIFESPAN_BOUNDS, step: 1 };
const SALARY: Bounds = { min: 0, max: 1_000_000, step: 100 };
const BUDGET: Bounds = { ...BUDGET_BOUNDS };

const clampTo = (value: number, bounds: Bounds): number => Math.min(Math.max(value, bounds.min), bounds.max);

/** The settings the saving year is read from until the shared context arrives: the workbook's own. */
const DEFAULT_CONTEXT = defaultContext();

/** The select's value for "Egen budget"; the examples are 0, 1 and 2. */
const OWN_BUDGET = BUDGETS.length;

/** The budget line's handle in the charts, in the place a payout period's goes. */
export const BUDGET_SERIES = -1;

/** The two sets of economic assumptions the tab offers a button for -- the same two as "Hur maxa pensionen?". */
const PRESETS = [
  { id: "standard", sv: "Modellens standard", en: "Model default", inflation: 0, growth: 0, ret: 0.017 },
  { id: "growth", sv: "Med löneökning", en: "With wage growth", inflation: 0.02, growth: 0.018, ret: 0.017 },
] as const;

export interface PensionLastsHandle {
  readonly element: HTMLElement;
  relabel(lang: Lang): void;
  /** Stashes the shared context and deaths and redraws the results. */
  setContext(context: ModelContext, deaths: DeathProbabilities): void;
}

// ---------------------------------------------------------------- phrases ---

const perMonth = (l: Lang): string => say(l, "kr/mån", "SEK/month");
const amount = (value: number, l: Lang): string => `${kronor(Math.round(value), l)} ${perMonth(l)}`;

/** "1 200 kr/mån över budgeten", "300 kr/mån under budgeten". */
function againstBudget(difference: number, l: Lang): string {
  const rounded = Math.round(Math.abs(difference));
  if (rounded === 0) return say(l, "lika med budgeten", "equal to the budget");
  return difference >= 0
    ? say(l, `${kronor(rounded, l)} kr/mån över budgeten`, `${kronor(rounded, l)} SEK/month above the budget`)
    : say(l, `${kronor(rounded, l)} kr/mån under budgeten`, `${kronor(rounded, l)} SEK/month below the budget`);
}

const taxText = (afterTax: boolean, l: Lang): string => (afterTax ? say(l, "efter skatt", "after tax") : say(l, "före skatt", "before tax"));

const priceText = (lasts: Lasts, l: Lang): string => {
  const year = lasts.runs.referenceYear;
  return say(l, `fasta priser (${year})`, `constant prices (${year})`);
};

/** A box with a payout period in it, coloured as the period's line is. */
function chip(payout: number, text: string): HTMLElement {
  const slot = slotOfPayout(payout);
  return h("span", `pa-chip ${slot === undefined ? "pa-chip-none" : `pa-chip-${slot}`}`, text);
}

/** A table cell carrying a payout period's line key beside its name. */
function payoutCell(payout: number, text: string): HTMLTableCellElement {
  const td = cell("");
  const slot = slotOfPayout(payout);
  if (slot !== undefined) td.append(h("span", `pa-key pa-key-${slot}`));
  td.append(document.createTextNode(text));
  return td;
}

/** How a budget stands, as a class for the table cell: met all the way, met for a time, never met. */
function statusOf(a: Assessment): "ok" | "part" | "no" {
  if (a.lasts) return "ok";
  return a.coveredYears > 0 ? "part" : "no";
}

/**
 * The years from the retirement age to the lifespan as a thin bar, the ones that meet the
 * budget filled: how long it lasts, at a glance, beside the words that say it. Decorative --
 * the cell's text carries the same.
 */
function coverageBar(a: Assessment): HTMLElement {
  const bar = h("span", "pl-bar");
  bar.setAttribute("aria-hidden", "true");
  const years = a.covered.length;
  const stops: string[] = [];
  for (let i = 0; i < years; ) {
    let j = i;
    while (j < years && a.covered[j] === a.covered[i]) j += 1;
    stops.push(`${a.covered[i] ? "var(--accent)" : "transparent"} ${(i / years) * 100}% ${(j / years) * 100}%`);
    i = j;
  }
  bar.style.backgroundImage = `linear-gradient(to right, ${stops.join(", ")})`;
  return bar;
}

function coverageCell(a: Assessment, l: Lang): HTMLTableCellElement {
  const td = cell("", `pl-cell pl-cell-${statusOf(a)}`);
  td.append(h("span", "pl-cell-text", coverageText(a, l)), coverageBar(a));
  td.dataset.status = statusOf(a);
  return td;
}

const optionOf = (lasts: Lasts, payout: number): PayoutOption => lasts.options.find((o) => o.payout === payout)!;

/** A payout period's pension a month at an age, in the basis shown. */
const monthlyAt = (lasts: Lasts, option: PayoutOption, age: number): number => option.view.income[age - lasts.from]! / 12;

// ------------------------------------------------------------------ panel ---

/**
 * Which tab the panel is. "lasts" is "Hur länge räcker pensionen?"; "amount" is "Hur mycket får jag?", the
 * same person, agreement, savings and assumptions and the same runs, answering what the pension pays a month
 * under each payout period rather than how long a budget lasts -- so it has no budget, and none of what is
 * about one.
 */
export type LastsVariant = "lasts" | "amount";

export function createPensionLastsPanel(
  lang: Lang,
  getForecastInput: () => TypfallInput,
  variant: LastsVariant = "lasts",
): PensionLastsHandle {
  let currentLang = lang;
  let input: LastsInput = normaliseLasts(defaultLastsInput());
  let salaryAs: "monthly" | "final" = "monthly";
  /** The typed final salary, per month; read only while `salaryAs` is "final". */
  let finalTarget = 0;
  let riktalderLocked = true;
  /** "Egen budget" chosen in the drop-down; otherwise the example whose amount the budget is, if any. */
  let ownBudget = false;
  let context: ModelContext | undefined;
  let deaths: DeathProbabilities | undefined;
  /** What the last run worked out, for the two field hints that quote it. */
  let derived = { finalSalary: 0, monthlySalary: input.monthlySalary };
  /** The engine runs, kept while only the budget, the lifespan or the tax view change. */
  let cache: { key: string; context: ModelContext; deaths: DeathProbabilities; runs: PayoutRuns } | undefined;
  /** The monthly salary found for a typed final salary, kept for the same reason. */
  let salaryCache: { key: string; context: ModelContext; value: number } | undefined;

  const element = h("div", `${variant === "lasts" ? "pension-lasts-panel" : "pension-amount-panel"} pension-age-panel`);
  element.dataset.role = variant === "lasts" ? "pension-lasts-panel" : "pension-amount-panel";

  // ---- the inputs, built once -------------------------------------------------
  const inputsSection = h("section", "panel");
  const inputsHead = h("h2");
  const inputsTitle = h("span");
  const inputsActions = h("div", "panel-actions");
  inputsHead.append(inputsTitle, inputsActions);
  const intro = h("p", "pa-intro");
  const grid = h("div", "pa-grid");
  inputsSection.append(inputsHead, intro, grid);

  const results = h("div", "pa-results");
  element.append(inputsSection, results);

  const importButton = h("button", "export-btn");
  importButton.type = "button";
  importButton.dataset.action = "pl-import-forecast";
  const resetButton = h("button", "export-btn");
  resetButton.type = "button";
  resetButton.dataset.action = "pl-reset";
  inputsActions.append(importButton, resetButton);

  const relabels: Relabel[] = [];
  const { field, number, percent, check, select } = fieldSet(grid, relabels, lang);

  function group(sv: string, en: string): void {
    const title = h("h3", "pa-group");
    const apply = (l: Lang) => {
      title.textContent = say(l, sv, en);
    };
    apply(lang);
    relabels.push(apply);
    grid.append(title);
  }

  const savingsUi = createSavingsInputs({
    prefix: "pl",
    grid,
    relabels,
    lang,
    value: () => input,
    update: (patch) => update(patch),
    savingStart: () => savingStart(),
    group,
  });

  // Person
  group("Person", "Person");
  const bornField = number(input.born, BORN, (born) => update({ born }));
  bornField.element.dataset.setting = "plBorn";
  const relabelBorn = field(bornField.element, (l) => ({
    label: t("birthYear", l),
    hint: say(l, `Riktålder ${riktalderFor(input.born)} år`, `${t("riktalderCheckbox", l)} ${riktalderFor(input.born)}`),
  }));

  const startField = number(input.startWorkAge, START_WORK, (startWorkAge) => update({ startWorkAge }));
  startField.element.dataset.setting = "plStartWork";
  field(startField.element, (l) => ({ label: t("startWorkAge", l), hint: rangeHint(START_WORK) }));

  const retirementField = number(input.retirementAge, RETIREMENT, (retirementAge) => update({ retirementAge }));
  retirementField.element.dataset.setting = "plRetirement";
  const relabelRetirement = field(retirementField.element, (l) => ({
    label: say(l, "Pensionsålder", "Retirement age"),
    hint: say(l, `Tidigast ${centreBounds(input.born).min} år`, `Earliest ${centreBounds(input.born).min}`),
  }));
  const riktalderBox = check(riktalderLocked, (locked) => {
    riktalderLocked = locked;
    update({});
  });
  riktalderBox.dataset.setting = "plRiktalder";
  field(riktalderBox, (l) => ({ label: t("riktalderCheckbox", l) }), "field field-check");

  const lifespanField = number(input.lifespan, LIFESPAN, (lifespan) => update({ lifespan }));
  lifespanField.element.dataset.setting = "plLifespan";
  const relabelLifespan = field(lifespanField.element, (l) => ({
    label: say(l, "Uppskattad livslängd", "Estimated lifespan"),
    hint: say(
      l,
      `år · ${input.lifespan - input.retirementAge} år som pensionär`,
      `years · ${input.lifespan - input.retirementAge} years retired`,
    ),
  }));

  // Salary
  group("Lön", "Salary");
  const salaryAsSelect = select(
    [
      { value: 0, label: (l) => say(l, "Månadslön", "Monthly salary") },
      { value: 1, label: (l) => say(l, "Slutlön (vid pensionsåldern)", "Final salary (at the retirement age)") },
    ],
    0,
    (value) => {
      if (value === 1 && salaryAs === "monthly") finalTarget = Math.round(derived.finalSalary);
      if (value === 0 && salaryAs === "final") input = { ...input, monthlySalary: derived.monthlySalary };
      salaryAs = value === 1 ? "final" : "monthly";
      sync();
      redraw();
    },
  );
  salaryAsSelect.element.dataset.setting = "plSalaryAs";
  relabels.push(salaryAsSelect.relabel);
  field(salaryAsSelect.element, (l) => ({ label: say(l, "Ange lön som", "Enter salary as") }), "field field-wide");

  const salaryField = number(input.monthlySalary, SALARY, (value) => {
    if (salaryAs === "monthly") update({ monthlySalary: value });
    else {
      finalTarget = value;
      redraw();
    }
  });
  salaryField.element.dataset.setting = "plSalary";
  const relabelSalary = field(salaryField.element, (l) => ({
    label: salaryAs === "monthly" ? t("monthlySalary", l) : t("finalSalary", l),
    hint:
      salaryAs === "monthly"
        ? say(
            l,
            `kr/mån · slutlön ${kronor(derived.finalSalary, l)} kr/mån`,
            `SEK/month · final salary ${kronor(derived.finalSalary, l)} SEK/month`,
          )
        : say(
            l,
            `kr/mån · månadslön ${kronor(derived.monthlySalary, l)} kr`,
            `SEK/month · monthly salary ${kronor(derived.monthlySalary, l)} SEK`,
          ),
  }));

  // Budget -- built either way, which keeps `sync` as it is, but only shown where there is a budget to read.
  const budgetSet = fieldSet(variant === "lasts" ? grid : h("div"), relabels, lang);
  if (variant === "lasts") group("Budget", "Budget");
  const budgetSelect = budgetSet.select(
    [
      ...BUDGETS.map((b, index) => ({
        value: index,
        label: (l: Lang) => `${say(l, b.sv, b.en)} – ${kronor(b.amount, l)} ${perMonth(l)}`,
      })),
      { value: OWN_BUDGET, label: (l: Lang) => say(l, "Egen budget", "Own budget") },
    ],
    1,
    (value) => {
      ownBudget = value === OWN_BUDGET;
      if (!ownBudget) input = { ...input, budget: BUDGETS[value]!.amount };
      sync();
      redraw();
    },
  );
  budgetSelect.element.dataset.setting = "plBudgetChoice";
  relabels.push(budgetSelect.relabel);
  budgetSet.field(budgetSelect.element, (l) => ({ label: say(l, "Budget som pensionär", "Budget as a pensioner") }), "field field-wide");

  const budgetField = budgetSet.number(input.budget, BUDGET, (value) => {
    ownBudget = budgetIdOf(value) === undefined;
    update({ budget: value });
  });
  budgetField.element.dataset.setting = "plBudget";
  budgetSet.field(budgetField.element, (l) => ({
    label: say(l, "Budget per månad efter skatt", "Budget per month after tax"),
    hint: say(l, "kr/mån i dagens penningvärde", "SEK/month in today's money"),
  }));

  // Pension
  group("Pension", "Pension");
  const schemeSelect = select(
    SCHEME_CHOICES.map((choice) => ({ value: choice.value, label: () => choice.label })),
    input.scheme,
    (scheme) => update({ scheme: scheme as SchemeId }),
  );
  schemeSelect.element.dataset.setting = "plScheme";
  relabels.push(schemeSelect.relabel);
  field(schemeSelect.element, (l) => ({ label: t("occupational", l) }), "field field-wide");

  // The individual occupational pension, the private pension and the private savings: built by
  // `savingsInputs.ts`, which "Hur maxa pensionen?" is built with too, so the two tabs ask the same way.
  savingsUi.appendPension();
  savingsUi.appendPrivateSavings();

  // Assumptions: two ready-made sets, then the three numbers behind them.
  group("Antaganden", "Assumptions");
  const presetBox = h("div", "panel-toggle pa-toggle");
  presetBox.setAttribute("role", "group");
  const presetButtons = PRESETS.map((preset) => {
    const button = h("button", "panel-btn");
    button.type = "button";
    button.dataset.preset = preset.id;
    button.addEventListener("click", () =>
      update({ yearlyInflation: preset.inflation, realGrowth: preset.growth, realReturn: preset.ret }),
    );
    presetBox.append(button);
    return { preset, button };
  });
  grid.append(presetBox);

  const inflationField = percent(input.yearlyInflation, (yearlyInflation) => update({ yearlyInflation }));
  inflationField.element.dataset.setting = "plInflation";
  field(inflationField.element, (l) => ({ label: t("inflation", l), hint: "%" }));
  relabels.push(inflationField.relabel);
  const growthField = percent(input.realGrowth, (realGrowth) => update({ realGrowth }));
  growthField.element.dataset.setting = "plGrowth";
  field(growthField.element, (l) => ({ label: t("realGrowth", l), hint: "%" }));
  relabels.push(growthField.relabel);
  const returnField = percent(input.realReturn, (realReturn) => update({ realReturn }));
  returnField.element.dataset.setting = "plReturn";
  field(returnField.element, (l) => ({ label: t("realReturn", l), hint: "%" }));
  relabels.push(returnField.relabel);

  // Show
  group("Visa", "Show");
  const basisBox = h("div", "panel-toggle pa-toggle");
  basisBox.setAttribute("role", "group");
  const basisButtons = [
    { afterTax: false, sv: "Före skatt", en: "Before tax", id: "gross" },
    { afterTax: true, sv: "Efter skatt", en: "After tax", id: "net" },
  ].map((choice) => {
    const button = h("button", "panel-btn");
    button.type = "button";
    button.dataset.basis = choice.id;
    button.addEventListener("click", () => update({ afterTax: choice.afterTax }));
    basisBox.append(button);
    return { choice, button };
  });
  grid.append(basisBox);

  // ---- state -> controls -------------------------------------------------------

  /** The year capital and saving count from, as the engine will use it. */
  function savingStart(): number {
    return savingYear(input, context ?? DEFAULT_CONTEXT);
  }

  /** The riktålder tracks the birth year while its box is ticked; either way the input is kept inside what the cohort may draw. */
  function normalise(): void {
    const wanted = riktalderLocked ? riktalderFor(input.born) : input.retirementAge;
    input = normaliseLasts({ ...input, retirementAge: wanted });
  }

  /** Writes the state into every control. Setting a value does not fire `change`. */
  function sync(): void {
    bornField.setValue(input.born);
    startField.setValue(input.startWorkAge);
    retirementField.setValue(input.retirementAge);
    retirementField.element.disabled = riktalderLocked;
    riktalderBox.checked = riktalderLocked;
    lifespanField.setValue(input.lifespan);
    salaryAsSelect.element.value = salaryAs === "final" ? "1" : "0";
    salaryField.setValue(salaryAs === "final" ? finalTarget : input.monthlySalary);
    const tier = ownBudget ? undefined : BUDGETS.findIndex((b) => b.amount === input.budget);
    budgetSelect.element.value = String(tier === undefined || tier < 0 ? OWN_BUDGET : tier);
    budgetField.setValue(input.budget);
    schemeSelect.element.value = String(input.scheme);
    savingsUi.sync(currentLang);
    inflationField.setValue(input.yearlyInflation);
    growthField.setValue(input.realGrowth);
    returnField.setValue(input.realReturn);
    for (const { preset, button } of presetButtons) {
      const active =
        input.yearlyInflation === preset.inflation && input.realGrowth === preset.growth && input.realReturn === preset.ret;
      button.classList.toggle("active", active);
    }
    for (const { choice, button } of basisButtons) {
      button.classList.toggle("active", choice.afterTax === input.afterTax);
      button.setAttribute("aria-pressed", String(choice.afterTax === input.afterTax));
    }
    for (const relabel of [
      relabelBorn,
      relabelRetirement,
      relabelLifespan,
      relabelSalary,
    ]) {
      relabel(currentLang);
    }
  }

  function update(patch: Partial<LastsInput>): void {
    input = { ...input, ...patch };
    normalise();
    sync();
    redraw();
  }

  importButton.addEventListener("click", () => {
    const forecast = getForecastInput();
    salaryAs = "monthly";
    finalTarget = 0;
    riktalderLocked = false;
    update({
      born: clampTo(forecast.born, BORN),
      startWorkAge: clampTo(forecast.startWorkAge, START_WORK),
      retirementAge: clampTo(forecast.retirementAge, RETIREMENT),
      monthlySalary: Math.round(clampTo(forecast.monthlySalary, SALARY)),
      scheme: forecast.scheme,
      yearlyInflation: forecast.yearlyInflation,
      realGrowth: forecast.realGrowth,
      realReturn: forecast.realReturn,
    });
  });
  resetButton.addEventListener("click", () => {
    input = defaultLastsInput();
    salaryAs = "monthly";
    finalTarget = 0;
    riktalderLocked = true;
    ownBudget = false;
    normalise();
    sync();
    redraw();
  });

  // ---- results -----------------------------------------------------------------

  function warningsBox(messages: readonly string[]): HTMLElement {
    const box = h("aside", "warnings");
    for (const message of messages) box.append(h("p", undefined, message));
    return box;
  }

  /** The budget as the person will recognise it: an example's name, or "own". */
  function budgetName(lasts: Lasts, l: Lang): string {
    const id = budgetIdOf(lasts.input.budget);
    const example = BUDGETS.find((b) => b.id === id);
    return example === undefined || ownBudget ? say(l, "egen budget", "own budget") : say(l, example.sv, example.en);
  }

  function answerPanel(lasts: Lasts, l: Lang): HTMLElement {
    const { section, body } = panel(say(l, "Svar i korthet", "Short answer"));
    section.dataset.role = "pl-answer";
    section.classList.add("pa-answer");
    const { input: current, from, to } = lasts;
    const years = to - from;
    const tax = taxText(current.afterTax, l);
    const life = optionOf(lasts, LIFELONG);

    body.append(
      h(
        "p",
        "pa-lead",
        say(
          l,
          "Allmän pension betalas ut livet ut. Tjänstepension, individuell tjänstepension, privat pension och privat sparande kan du ofta ta ut under 5–20 år eller livsvarigt: kortare tid ger mer per månad men tar slut. Varje ruta visar hur länge din budget räcker med en utbetalningstid, månad för månad, till dess att du är " +
            `${to} år.`,
          "The public pension is paid for life. Occupational pension, individual occupational pension, private pension and private savings can often be drawn over 5–20 years or for life: a shorter time pays more a month but runs out. Each box shows how long your budget lasts with one payout period, month by month, until you are " +
            `${to}.`,
        ),
      ),
    );

    // One box per drawn payout period, and the recommended one if it is not among them.
    const shown: PayoutOption[] = CHART_PAYOUTS.map((p) => optionOf(lasts, p));
    if (lasts.best !== null && !shown.includes(lasts.best)) shown.push(lasts.best);
    shown.sort((a, b) => (a.payout === LIFELONG ? 1000 : a.payout) - (b.payout === LIFELONG ? 1000 : b.payout));
    const ribbonBox = h("div", "pa-ribbon-box");
    ribbonBox.append(h("div", "pa-ribbon-label", say(l, "Utbetalningstid och hur länge budgeten räcker", "Payout period and how long the budget lasts")));
    const ribbon = h("ol", "pa-ribbon pl-ribbon");
    ribbon.dataset.role = "pl-ribbon";
    for (const option of shown) {
      const a = option.view;
      const item = h("li", "pa-ribbon-item pl-ribbon-item");
      item.dataset.payout = String(option.payout);
      item.dataset.status = statusOf(a);
      item.dataset.best = lasts.best === option ? "1" : "0";
      item.title = say(
        l,
        `${payoutName(option.payout, l)}: ${coverageText(a, l)}. ${a.coveredYears} av ${years} år täcks.`,
        `${payoutName(option.payout, l)}: ${coverageText(a, l)}. ${a.coveredYears} of ${years} years covered.`,
      );
      if (lasts.best === option) item.append(h("span", "pl-best-tag", say(l, "Bäst", "Best")));
      item.append(
        chip(option.payout, payoutName(option.payout, l)),
        h("span", `pl-ribbon-d pl-ribbon-${statusOf(a)}`, coverageText(a, l, true)),
        h(
          "span",
          "pa-ribbon-m",
          option.view.during === null
            ? ""
            : `${signed(kronor(Math.round(option.view.during.difference), l), option.view.during.difference)} ${perMonth(l)}`,
        ),
      );
      ribbon.append(item);
    }
    ribbonBox.append(ribbon);
    ribbonBox.append(
      h(
        "div",
        "pa-ribbon-label",
        say(
          l,
          `Siffran under är pensionen minus budgeten under själva utbetalningen, per månad ${tax}.`,
          `The figure beneath is the pension minus the budget during the payout itself, per month ${tax}.`,
        ),
      ),
    );
    body.append(ribbonBox);

    // The budget and the public pension alone, in the terms of the basis shown.
    const publicAlone = lasts.publicOnly.during?.income ?? 0;
    const budgetNow = current.afterTax ? current.budget : lasts.budgetGross.typical;
    const example = budgetName(lasts, l);
    body.append(
      h(
        "p",
        "pa-kpi",
        say(
          l,
          `Budget (${example}): ${amount(current.budget, l)} efter skatt` +
            (current.afterTax ? "" : `, vilket motsvarar ungefär ${amount(lasts.budgetGross.typical, l)} före skatt`) +
            `. Bara allmän pension ger ${amount(publicAlone, l)} ${tax}, ${againstBudget(publicAlone - budgetNow, l)}.`,
          `Budget (${example}): ${amount(current.budget, l)} after tax` +
            (current.afterTax ? "" : `, which is about ${amount(lasts.budgetGross.typical, l)} before tax`) +
            `. The public pension alone gives ${amount(publicAlone, l)} ${tax}, ${againstBudget(publicAlone - budgetNow, l)}.`,
        ),
      ),
    );

    const list = h("ul", "pa-points");
    const bullet = (text: string): void => {
      list.append(h("li", undefined, text));
    };

    // 1. what to choose
    const best = lasts.best;
    if (!lasts.flexible) {
      bullet(
        say(
          l,
          "Du har ingen tjänstepension, privat pension eller annat sparande att välja utbetalningstid för, så allmän pension är allt. Välj ett avtalsområde eller fyll i kapital och sparande för att se hur utbetalningstiden påverkar.",
          "You have no occupational pension, private pension or other saving to choose a payout period for, so the public pension is all there is. Pick an agreement area, or fill in capital and saving, to see what the payout period does.",
        ),
      );
    } else if (best === null) {
      const short = optionOf(lasts, 5);
      bullet(
        say(
          l,
          `Ingen utbetalningstid täcker budgeten något enda år. Mest ger 5 års utbetalning: ${amount(short.view.during?.income ?? 0, l)} ${tax}, ${againstBudget(short.view.during?.difference ?? 0, l)}.`,
          `No payout period covers the budget in a single year. The most comes from a 5-year payout: ${amount(short.view.during?.income ?? 0, l)} ${tax}, ${againstBudget(short.view.during?.difference ?? 0, l)}.`,
        ),
      );
    } else if (best.payout === LIFELONG && best.view.lasts) {
      bullet(
        say(
          l,
          `Livsvarig utbetalning täcker budgeten hela vägen till ${to} år: pensionen är i snitt ${amount(best.view.during!.income, l)} ${tax}, ${againstBudget(best.view.during!.difference, l)}. Det är också det tryggaste valet – det tar inte slut om du lever längre än ${to} år.`,
          `A lifelong payout covers the budget all the way to ${to}: the pension averages ${amount(best.view.during!.income, l)} ${tax}, ${againstBudget(best.view.during!.difference, l)}. It is also the safest choice – it does not run out if you live past ${to}.`,
        ),
      );
    } else if (best.payout === LIFELONG) {
      bullet(
        say(
          l,
          `Livsvarig utbetalning täcker budgeten längst: ${best.view.coveredYears} av ${years} år (${coverageText(best.view, l)}). Pensionen är i snitt ${amount(best.view.during!.income, l)} ${tax}, ${againstBudget(best.view.during!.difference, l)}.`,
          `A lifelong payout covers the budget longest: ${best.view.coveredYears} of ${years} years (${coverageText(best.view, l)}). The pension averages ${amount(best.view.during!.income, l)} ${tax}, ${againstBudget(best.view.during!.difference, l)}.`,
        ),
      );
    } else {
      const during = best.view.during!;
      const after = best.view.after;
      const lifeNote =
        life.view.coveredYears === 0
          ? say(
              l,
              ` Livsvarigt ger ${amount(life.view.during!.income, l)}, ${againstBudget(life.view.during!.difference, l)}.`,
              ` A lifelong payout gives ${amount(life.view.during!.income, l)}, ${againstBudget(life.view.during!.difference, l)}.`,
            )
          : "";
      bullet(
        say(
          l,
          `Längst täcks budgeten med ${best.payout} års utbetalning: ${best.view.coveredYears} av ${years} år (${coverageText(best.view, l)}). Under utbetalningen är pensionen i snitt ${amount(during.income, l)} ${tax}, ${againstBudget(during.difference, l)}.` +
            (after === null
              ? ""
              : ` Därefter är det bara allmän pension kvar: ${amount(after.income, l)}, ${againstBudget(after.difference, l)} i ${after.years} år till ${to} år.`) +
            lifeNote,
          `The budget lasts longest with the ${best.payout}-year payout: ${best.view.coveredYears} of ${years} years (${coverageText(best.view, l)}). During the payout the pension averages ${amount(during.income, l)} ${tax}, ${againstBudget(during.difference, l)}.` +
            (after === null
              ? ""
              : ` After that only the public pension is left: ${amount(after.income, l)}, ${againstBudget(after.difference, l)} for ${after.years} years to ${to}.`) +
            lifeNote,
        ),
      );
    }

    // 2. what saving the surplus would do
    if (lasts.flexible) {
      const lasting = lasts.options.filter((o) => o.net.balanceRunsOutAt === null);
      // The shortest payout from which every longer one also lasts.
      let suffix = lasts.options.length;
      for (let i = lasts.options.length - 1; i >= 0 && lasts.options[i]!.net.balanceRunsOutAt === null; i -= 1) suffix = i;
      const saved = say(l, "Sparar du det som blir över (utan ränta, efter skatt)", "If you save what is left over (no interest, after tax)");
      const phrase = (p: number): string =>
        p === LIFELONG ? say(l, "livsvarig utbetalning", "the lifelong payout") : say(l, `${p} års utbetalning`, `the ${p}-year payout`);
      if (lasting.length > 0) {
        // The shortest payout from which every longer one lasts too -- or, if lifelong does not, the first that does.
        const first = lasts.options[suffix < lasts.options.length ? suffix : lasts.options.indexOf(lasting[0]!)]!;
        let text: string;
        if (suffix === 0) {
          text = say(
            l,
            `${saved} räcker pengarna hela vägen till ${to} år med alla utbetalningstider, även 5 år.`,
            `${saved} the money lasts all the way to ${to} with every payout period, even 5 years.`,
          );
        } else if (suffix === lasts.options.length - 1) {
          text = say(
            l,
            `${saved} räcker pengarna hela vägen till ${to} år bara med livsvarig utbetalning.`,
            `${saved} the money lasts all the way to ${to} only with the lifelong payout.`,
          );
        } else if (suffix < lasts.options.length) {
          text = say(
            l,
            `${saved} räcker pengarna hela vägen till ${to} år med ${phrase(first.payout)} eller längre.`,
            `${saved} the money lasts all the way to ${to} with ${phrase(first.payout)} or longer.`,
          );
        } else {
          text = say(
            l,
            `${saved} räcker pengarna hela vägen till ${to} år med bland annat ${phrase(first.payout)}, men inte med alla.`,
            `${saved} the money lasts all the way to ${to} with, among others, ${phrase(first.payout)}, but not with all.`,
          );
        }
        bullet(text);
      } else {
        const longest = lasts.options.reduce((a, b) => ((b.net.balanceRunsOutAt ?? 0) > (a.net.balanceRunsOutAt ?? 0) ? b : a));
        bullet(
          say(
            l,
            `Inte heller om du sparar det som blir över (utan ränta, efter skatt) räcker pengarna till ${to} år. Längst räcker de med ${phrase(longest.payout)}: ${balanceText(longest.net, l).toLowerCase()}.`,
            `Nor does the money last to ${to} if you save what is left over (no interest, after tax). It lasts longest with ${phrase(longest.payout)}: ${balanceText(longest.net, l).toLowerCase()}.`,
          ),
        );
      }

      // 3. what a payout leaves behind
      const floor = lasts.publicOnly.during;
      if (floor !== null) {
        bullet(
          say(
            l,
            `När en tidsbegränsad utbetalning är slut återstår bara allmän pension: i snitt ${amount(floor.income, l)} ${tax} i dagens penningvärde, ${share(floor.income / Math.max(budgetNow, 1), l, 0)} av budgeten.`,
            `When a temporary payout ends only the public pension is left: on average ${amount(floor.income, l)} ${tax} in today's money, ${share(floor.income / Math.max(budgetNow, 1), l, 0)} of the budget.`,
          ),
        );
      }
    }

    // 4. which of minPension's budgets a lifelong payout meets
    const examples = budgetMatrix(lasts, BUDGETS.map((b) => b.amount)).find((r) => r.payout === LIFELONG)!;
    const met = BUDGETS.filter((_, i) => examples.cells[i]!.lasts).map((b) => say(l, b.sv, b.en));
    const missed = BUDGETS.filter((_, i) => !examples.cells[i]!.lasts).map((b) => say(l, b.sv, b.en));
    const join = (names: string[]): string => names.join(say(l, " och ", " and "));
    bullet(
      met.length === BUDGETS.length
        ? say(l, "Med livsvarig utbetalning klaras alla tre av minPensions budgetar hela vägen.", "With a lifelong payout all three of minPension's budgets are met all the way.")
        : met.length === 0
          ? say(l, "Med livsvarig utbetalning klaras ingen av minPensions tre budgetar hela vägen.", "With a lifelong payout none of minPension's three budgets is met all the way.")
          : say(
              l,
              `Med livsvarig utbetalning klaras ${join(met)} hela vägen, men inte ${join(missed)}.`,
              `With a lifelong payout ${join(met)} ${met.length > 1 ? "are" : "is"} met all the way, but not ${join(missed)}.`,
            ),
    );

    body.append(list);
    return section;
  }

  // ---- charts ---------------------------------------------------------------

  /** The ages the charts cover: the retirement age up to the year before the lifespan. */
  function chartAges(lasts: Lasts): number[] {
    return Array.from({ length: lasts.to - lasts.from }, (_, i) => lasts.from + i);
  }

  function incomeFigure(lasts: Lasts, l: Lang): HTMLElement {
    const ages = chartAges(lasts);
    const measure = capital(`${say(l, "pension", "pension")} ${taxText(lasts.input.afterTax, l)}`);
    const series: ChartSeries[] = CHART_PAYOUTS.map((payout) => {
      const option = optionOf(lasts, payout);
      const endAt = payout === LIFELONG ? undefined : lasts.from + payout - 1;
      return {
        payout,
        name: payout === LIFELONG ? say(l, "Livsvarig utbetalning", "Lifelong payout") : say(l, `${payout} års utbetalning`, `${payout}-year payout`),
        end: payoutName(payout, l),
        colour: `--fig-pu-${slotOfPayout(payout)}`,
        emphasis: payout === LIFELONG,
        essential: payout === LIFELONG,
        ...(endAt !== undefined && endAt < lasts.to - 1 ? { endAt } : {}),
        points: ages.map((age) => [age, monthlyAt(lasts, option, age)] as const),
        format: (value: number) => `${kronor(value, l)} ${perMonth(l)}`,
      };
    });
    if (variant === "lasts") {
      series.push({
        payout: BUDGET_SERIES,
        name: `${say(l, "Budget", "Budget")}, ${taxText(lasts.input.afterTax, l)}`,
        end: say(l, "Budget", "Budget"),
        colour: "--text-muted",
        dashed: true,
        essential: true,
        points: ages.map((age) => [age, lasts.budgetMonthly[age - lasts.from]!] as const),
        format: (value: number) => `${kronor(value, l)} ${perMonth(l)}`,
      });
    }
    const top = Math.max(1, ...series.flatMap((s) => s.points.map((p) => p[1])));
    const ticks = niceTicks(0, top, 5);
    return renderLineChart({
      role: variant === "lasts" ? "pl-fig-income" : "pm-fig-income",
      lang: l,
      title:
        variant === "lasts"
          ? say(l, "Pension per månad mot budgeten", "Pension per month against the budget")
          : say(l, "Pension per månad", "Pension per month"),
      subtitle:
        variant === "lasts"
          ? say(
              l,
              `${measure}, ${priceText(lasts, l)}. Varje linje är pensionen vid en utbetalningstid; den streckade är budgeten. Linjerna slutar vid ${lasts.to - 1} år, året före ${lasts.to}-årsdagen.`,
              `${measure}, ${priceText(lasts, l)}. Each line is the pension under one payout period; the dashed one is the budget. The lines stop at ${lasts.to - 1}, the year before the ${lasts.to}th birthday.`,
            )
          : say(
              l,
              `${measure}, ${priceText(lasts, l)}. Varje linje är pensionen vid en utbetalningstid. Linjerna slutar vid ${lasts.to - 1} år, året före ${lasts.to}-årsdagen.`,
              `${measure}, ${priceText(lasts, l)}. Each line is the pension under one payout period. The lines stop at ${lasts.to - 1}, the year before the ${lasts.to}th birthday.`,
            ),
      xMin: lasts.from,
      xMax: Math.max(lasts.to - 1, lasts.from + 1),
      xTitle: say(l, "Ålder", "Age"),
      yTicks: ticks.ticks,
      yTick: (v) => kronor(v, l),
      yUnit: say(l, "kr per månad", "SEK per month"),
      heading: (x) => `${say(l, "Ålder", "Age")} ${x}`,
      series,
      endRoom: 74,
      ...(lasts.input.afterTax || variant === "amount"
        ? {}
        : {
            note: say(
              l,
              "Budgeten före skatt är den bruttoinkomst som efter modellens skatt ger budgeten efter skatt; den är högre de första åren, då skatten är högre.",
              "The budget before tax is the gross income that leaves the budget after the model's tax; it is higher in the first years, when the tax is higher.",
            ),
          }),
    });
  }

  function balanceFigure(lasts: Lasts, l: Lang): HTMLElement {
    const ages = chartAges(lasts);
    const series: ChartSeries[] = CHART_PAYOUTS.map((payout) => {
      const option = optionOf(lasts, payout);
      return {
        payout,
        name: payout === LIFELONG ? say(l, "Livsvarig utbetalning", "Lifelong payout") : say(l, `${payout} års utbetalning`, `${payout}-year payout`),
        end: payoutName(payout, l),
        colour: `--fig-pu-${slotOfPayout(payout)}`,
        emphasis: payout === LIFELONG,
        essential: payout === LIFELONG || payout === 5 || payout === 20,
        points: ages.map((age) => [age, option.net.balance[age - lasts.from]!] as const),
        format: (value: number) => `${signed(mkr(value, l), value)} ${say(l, "Mkr", "MSEK")}`,
      };
    });
    const values = series.flatMap((s) => s.points.map((p) => p[1]));
    const ticks = niceTicks(Math.min(0, ...values), Math.max(0, ...values), 6);
    return renderLineChart({
      role: "pl-fig-balance",
      lang: l,
      title: say(l, "Räcker pengarna om överskottet sparas?", "Does the money last if the surplus is saved?"),
      subtitle: say(
        l,
        `Summan av pensionen minus budgeten sedan pensionsstart, efter skatt, miljoner kronor, ${priceText(lasts, l)}, utan ränta. Över noll finns pengar kvar; där en linje når noll är pengarna slut.`,
        `The pension minus the budget added up since the start, after tax, millions of kronor, ${priceText(lasts, l)}, with no interest. Above zero there is money left; where a line reaches zero the money is gone.`,
      ),
      xMin: lasts.from,
      xMax: Math.max(lasts.to - 1, lasts.from + 1),
      xTitle: say(l, "Ålder", "Age"),
      yTicks: ticks.ticks,
      yTick: (v) => (v === 0 ? "0" : `${v < 0 ? "−" : "+"}${decimals(Math.abs(v) / 1e6, 1, l)}`),
      yUnit: say(l, "Mkr", "MSEK"),
      heading: (x) => `${say(l, "Ålder", "Age")} ${x}`,
      series,
      endRoom: 74,
      annotations: {
        top: say(l, "▲ pengar kvar", "▲ money left"),
        // Nothing is "below" when no line goes below zero.
        ...(ticks.lo < 0 ? { bottom: say(l, "▼ pengarna slut", "▼ money gone") } : {}),
      },
    });
  }

  // ---- tables ---------------------------------------------------------------

  /** The first chart's numbers, at the retirement age and every fifth age after it -- its table twin, with the per-year export. */
  function timePanel(lasts: Lasts, l: Lang): HTMLElement {
    const actions =
      variant === "lasts"
        ? [
            csvExportButton(l, say(l, "pension-per-ar.csv", "pension-per-year.csv"), () => seriesToCsv(lasts, l)),
            xlsxExportButton(l, say(l, "pension-per-ar.xlsx", "pension-per-year.xlsx"), () => seriesToXlsxRows(lasts, l)),
          ]
        : [
            csvExportButton(l, say(l, "pension-per-ar.csv", "pension-per-year.csv"), () => amountSeriesToCsv(lasts, l)),
            xlsxExportButton(l, say(l, "pension-per-ar.xlsx", "pension-per-year.xlsx"), () => amountSeriesToXlsxRows(lasts, l)),
          ];
    const { section, body } = panel(say(l, "Pension per månad vid olika åldrar", "Pension per month at different ages"), actions);
    section.dataset.role = variant === "lasts" ? "pl-time" : "pm-time";
    const ages: number[] = [lasts.from];
    for (let age = Math.ceil((lasts.from + 1) / 5) * 5; age < lasts.to; age += 5) ages.push(age);

    const head = [headCell(say(l, "Utbetalningstid", "Payout period")), ...ages.map((age) => headCell(`${age} ${say(l, "år", "yrs")}`))];
    const rows: HTMLTableRowElement[] = [];
    if (variant === "lasts") {
      const budget = cell("");
      budget.append(h("span", "pa-key pl-key-budget"), document.createTextNode(`${say(l, "Budget", "Budget")}, ${taxText(lasts.input.afterTax, l)}`));
      rows.push(row([budget, ...ages.map((age) => cell(kronor(lasts.budgetMonthly[age - lasts.from]!, l), "num"))], { payout: String(BUDGET_SERIES) }));
    }
    for (const payout of CHART_PAYOUTS) {
      const option = optionOf(lasts, payout);
      rows.push(
        row(
          [payoutCell(payout, payoutName(payout, l)), ...ages.map((age) => cell(kronor(monthlyAt(lasts, option, age), l), "num"))],
          { payout: String(payout) },
        ),
      );
    }
    body.append(tableOf("pl-time-table", "pl-time", head, rows));
    return section;
  }

  function optionsPanel(lasts: Lasts, l: Lang): HTMLElement {
    const afterTax = lasts.input.afterTax;
    const actions = [
      csvExportButton(l, say(l, "utbetalningstider.csv", "payout-periods.csv"), () => optionsToCsv(lasts, l)),
      xlsxExportButton(l, say(l, "utbetalningstider.xlsx", "payout-periods.xlsx"), () => optionsToXlsxRows(lasts, l)),
    ];
    const { section, body } = panel(say(l, "Utbetalningstid och budget", "Payout period and budget"), actions);
    section.dataset.role = "pl-options-panel";
    body.append(
      h(
        "p",
        "pa-lead",
        say(
          l,
          `Alla utbetalningstider, per månad ${taxText(afterTax, l)}, i ${priceText(lasts, l)}. "Under" är själva utbetalningen, "efter" åren därefter då bara allmän pension återstår. Den rad som täcker budgeten i flest år är markerad.`,
          `Every payout period, per month ${taxText(afterTax, l)}, in ${priceText(lasts, l)}. "During" is the payout itself, "after" the years that follow when only the public pension is left. The row that covers the budget in the most years is highlighted.`,
        ),
      ),
    );

    const head = [
      headCell(say(l, "Utbetalningstid", "Payout period")),
      headCell(say(l, "Under utbetalningen", "During the payout")),
      ...(afterTax
        ? []
        : [
            headCell(say(l, "varav allmän", "of which public")),
            headCell(say(l, "varav tjänstepension", "of which occupational")),
            headCell(say(l, "varav individuell och privat", "of which individual and private")),
            ...(lasts.taxFree ? [headCell(say(l, "varav privat sparande, uppräknat", "of which private savings, grossed up"))] : []),
          ]),
      headCell(say(l, "Efter utbetalningen", "After the payout")),
      headCell(say(l, "Mot budgeten under / efter", "Against the budget during / after")),
      headCell(say(l, "Budgeten räcker", "The budget lasts")),
      headCell(say(l, "År som täcks", "Years covered")),
      headCell(say(l, "Saknas totalt, tkr", "Short in total, thousand kr")),
      headCell(say(l, "Räcker om överskottet sparas", "Lasts if the surplus is saved")),
    ];

    const average = (option: PayoutOption, pick: (o: PayoutOption) => readonly number[]): number => {
      const during = option.view.during;
      if (during === null) return 0;
      let sum = 0;
      for (let age = during.from; age < during.to; age += 1) sum += pick(option)[age] ?? 0;
      return sum / during.years / 12;
    };
    const diff = (d: number): string => `${signed(kronor(Math.round(d), l), Math.round(d))}`;

    const rows: HTMLTableRowElement[] = [];
    // The public pension alone: the floor every temporary payout ends at.
    {
      const a = lasts.publicOnly;
      const cells: HTMLTableCellElement[] = [cell(say(l, "Bara allmän pension", "Public pension only"), "pl-floor-name"), cell(kronor(a.during?.income ?? 0, l), "num")];
      if (!afterTax) cells.push(...Array.from({ length: lasts.taxFree ? 4 : 3 }, () => cell("–", "num")));
      cells.push(
        cell("–", "num"),
        cell(`${diff(a.during?.difference ?? 0)} / –`, "num"),
        coverageCell(a, l),
        cell(`${a.coveredYears} / ${a.to - a.from}`, "num"),
        cell(kronor(Math.round(a.shortfall / 1000), l), "num"),
        cell("–", "num"),
      );
      rows.push(row(cells, { payout: "floor" }));
    }
    for (const option of lasts.options) {
      const a = option.view;
      const cells: HTMLTableCellElement[] = [
        payoutCell(option.payout, payoutName(option.payout, l)),
        cell(kronor(a.during?.income ?? 0, l), "num"),
      ];
      if (!afterTax) {
        cells.push(
          cell(kronor(average(option, (o) => o.run.publicPension), l), "num"),
          cell(kronor(average(option, (o) => o.run.occupational), l), "num"),
          cell(kronor(average(option, (o) => o.run.individual), l), "num"),
          ...(lasts.taxFree ? [cell(kronor(average(option, (o) => o.savings), l), "num")] : []),
        );
      }
      cells.push(
        cell(a.after === null ? "–" : kronor(a.after.income, l), "num"),
        cell(`${diff(a.during?.difference ?? 0)} / ${a.after === null ? "–" : diff(a.after.difference)}`, "num"),
        coverageCell(a, l),
        cell(`${a.coveredYears} / ${a.to - a.from}`, "num"),
        cell(kronor(Math.round(a.shortfall / 1000), l), "num"),
        cell(balanceText(option.net, l), "pl-cell"),
      );
      const tr = row(cells, { payout: String(option.payout) });
      if (lasts.best === option) tr.classList.add("pl-best");
      rows.push(tr);
    }
    body.append(tableOf("pl-options", "pl-options", head, rows));
    return section;
  }

  function matrixPanel(lasts: Lasts, l: Lang): HTMLElement {
    // The three examples, and the person's own budget if it is none of them.
    const budgets: { label: string; amount: number }[] = BUDGETS.map((b) => ({ label: say(l, b.sv, b.en), amount: b.amount }));
    if (budgetIdOf(lasts.input.budget) === undefined || ownBudget) {
      if (!budgets.some((b) => b.amount === lasts.input.budget)) {
        budgets.push({ label: say(l, "Egen budget", "Own budget"), amount: lasts.input.budget });
      }
    }
    const matrix = budgetMatrix(lasts, budgets.map((b) => b.amount));
    const actions = [
      csvExportButton(l, say(l, "budgetar.csv", "budgets.csv"), () => matrixToCsv(matrix, budgets, l)),
      xlsxExportButton(l, say(l, "budgetar.xlsx", "budgets.xlsx"), () => matrixToXlsxRows(matrix, budgets, l)),
    ];
    const { section, body } = panel(say(l, "Hur länge räcker olika budgetar?", "How long do different budgets last?"), actions);
    section.dataset.role = "pl-matrix-panel";
    body.append(
      h(
        "p",
        "pa-lead",
        say(
          l,
          `Samma pension mot minPensions tre budgetar och din egen, efter skatt: till vilken ålder varje budget täcks, månad för månad. Sista kolumnen är den högsta budget som klaras varje år av utbetalningen.`,
          `The same pension against minPension's three budgets and your own, after tax: the age each budget is covered to, month by month. The last column is the highest budget the payout meets every year.`,
        ),
      ),
    );
    const head = [
      headCell(say(l, "Utbetalningstid", "Payout period")),
      ...budgets.map((b) => headCell(`${b.label} (${kronor(b.amount, l)})`)),
      headCell(say(l, "Högsta budget som klaras varje år", "Highest budget met every year")),
    ];
    const rows = matrix.map((m) => {
      const cells: HTMLTableCellElement[] = [payoutCell(m.payout, payoutName(m.payout, l))];
      m.cells.forEach((a, i) => {
        const td = coverageCell(a, l);
        td.dataset.budget = String(budgets[i]!.amount);
        cells.push(td);
      });
      cells.push(cell(kronor(m.highest, l), "num"));
      return row(cells, { payout: String(m.payout) });
    });
    body.append(tableOf("pl-matrix", "pl-matrix", head, rows));
    return section;
  }

  function notesPanel(lasts: Lasts, l: Lang): HTMLElement {
    const details = h("details", "field-note pa-notes");
    details.dataset.role = "pl-notes";
    details.append(h("summary", undefined, say(l, "Antaganden och metod", "Assumptions and method")));
    const bodyBox = h("div", "field-note-body");
    const savings = savingsNotes(l, lasts.runs.savingYear);
    const paragraphs = [
      say(
        l,
        `Pensionen är allmän pension (inkomstpension, premiepension, garantipension och inkomstpensionstillägg), tjänstepension från avtalsområdet och eventuell individuell tjänstepension, privat pension och privat sparande, från pensionsstart till och med året före ${lasts.to}-årsdagen (${lasts.to - lasts.from} hela år). Belopp i ${priceText(lasts, l)}, utan diskontering. Bidrag som bostadstillägg räknas inte med; efter skatt är modellens nettoinkomst.`,
        `The pension is the public pension (income pension, premium pension, guarantee pension and the income-pension supplement), occupational pension from the agreement area and any individual occupational pension, private pension and private savings, from the start up to the year before the ${lasts.to}th birthday (${lasts.to - lasts.from} whole years). Amounts are in ${priceText(lasts, l)}, not discounted. Benefits such as housing supplement are not counted; after tax is the model's net income.`,
      ),
      say(
        l,
        `Utbetalningstiden är modellens inställning "Temporärt uttag" för tjänstepension och privat sparande: pensionen betalas ut jämnt under 5–20 hela år från pensionsstart och slutar sedan, eller livsvarigt. Allmän pension påverkas inte. Kortare tid ger mer per månad eftersom samma kapital delas på färre år; hur mycket styrs av avtalets egna antaganden om ränta och livslängd.`,
        `The payout period is the model's "Temporary withdrawal" setting for occupational pension and private saving: the pension is paid evenly over 5–20 whole years from the start and then stops, or for life. The public pension is not affected. A shorter time pays more a month because the same capital is shared over fewer years; how much is set by the agreement's own assumptions about interest and life expectancy.`,
      ),
      // The two paragraphs about the savings are the other tab's too (`savingsNotes`); this tab adds how long they are paid.
      `${savings.pension} ${say(l, "Allt betalas ut över samma period som avtalets tjänstepension.", "Everything is paid out over the same period as the agreement's occupational pension.")}`,
      savings.savings,
      say(
        l,
        `Budgeten är ett belopp per månad efter skatt i dagens penningvärde: minPensions exempel (samma belopp som Pensionsöversikten) eller ditt eget. Före skatt visas den bruttoinkomst som modellens skatt gör till budgeten efter skatt, år för år, inte en genomsnittlig skattesats. Eftersom skatten aldrig lämnar mindre för mer täcker en pension budgeten före skatt precis när den täcker den efter skatt.`,
        `The budget is an amount per month after tax in today's money: minPension's examples (the same amounts as Pensionsöversikten) or your own. Before tax shows the gross income that the model's tax turns into the budget after tax, year by year, not an average tax rate. Since tax never leaves less for more, a pension covers the budget before tax exactly when it covers it after.`,
      ),
      say(
        l,
        `"Räcker" betyder att pensionen vid en ålder är minst lika stor som budgeten; "Budgeten räcker" visar vilka åldrar som täcks. Modellens skatt är högre åren innan det förhöjda grundavdraget för äldre börjar, så de första åren kan ligga lägre än de följande; det visas som det är. "Räcker om överskottet sparas" lägger ihop pension minus budget år för år (utan ränta, efter skatt) och visar när summan blir negativ.`,
        `"Lasts" means the pension at an age is at least the budget; "The budget lasts" shows which ages are covered. The model's tax is higher in the years before the higher basic allowance for older people begins, so the first years can lie lower than the ones after; it is shown as it is. "Lasts if the surplus is saved" adds up pension minus budget year by year (no interest, after tax) and shows when the total turns negative.`,
      ),
      say(
        l,
        `Den rekommenderade utbetalningstiden är den som täcker budgeten i flest år; vid lika många år vinner den längre. Fliken har egna indata och följer inte formuläret till vänster; "Hämta från Prognos" kopierar dem en gång. Skatteinställningar och löneprofil kommer från Avancerat; prisbasen är alltid fasta priser, och tjänstepensionen börjar alltid vid pensionsåldern. Kända kapitalvärden som anges under Avancerat gäller också här.`,
        `The recommended payout period is the one that covers the budget in the most years; with an equal number, the longer wins. The tab has its own inputs and does not follow the form on the left; "Use the Forecast values" copies them once. Tax settings and wage profile come from Advanced; prices are always constant, and the occupational pension always starts at the retirement age. Known capital values entered under Advanced apply here too.`,
      ),
      ...(lasts.taxFree
        ? [
            say(
              l,
              "Privat sparande har ingen utbetalning för livet: Livsvarigt betyder här att pengarna fördelas över din förväntade återstående livslängd (modellens regel), och sedan är de slut. En pension betalas i stället ut så länge du lever.",
              "Private savings have no payout for life: Lifelong here means the money is spread over your expected remaining life (the model's rule), and then it is gone. A pension, by contrast, is paid for as long as you live.",
            ),
            say(
              l,
              "Före skatt räknas utbetalningen från det privata sparandet upp till den bruttoinkomst som ger lika mycket efter skatt (med modellens skatt, år för år). Annars skulle en skattefri krona räknas som en beskattad och vyerna före och efter skatt ge olika svar. Efter skatt visas utbetalningen som den är.",
              "Before tax, the payout from the private savings is grossed up to the gross income that leaves the same amount after tax (with the model's tax, year by year). Otherwise an untaxed krona would count as a taxed one and the views before and after tax would give different answers. After tax the payout is shown as it is.",
            ),
          ]
        : []),
    ];
    for (const text of paragraphs) bodyBox.append(h("p", undefined, text));
    details.append(bodyBox);
    return details;
  }

  // ---- "Hur mycket får jag?" ----------------------------------------------------

  /** The payout periods the tab shows: 5, 10, 15 and 20 years, and for life. */
  const shownOptions = (lasts: Lasts): PayoutOption[] => CHART_PAYOUTS.map((p) => optionOf(lasts, p));

  function amountAnswer(lasts: Lasts, l: Lang): HTMLElement {
    const { section, body } = panel(say(l, "Svar i korthet", "Short answer"));
    section.dataset.role = "pm-answer";
    section.classList.add("pa-answer");
    const { from, to } = lasts;
    const tax = taxText(lasts.input.afterTax, l);
    const life = optionOf(lasts, LIFELONG);
    const short = optionOf(lasts, 5);

    body.append(
      h(
        "p",
        "pa-lead",
        say(
          l,
          `Så mycket får du i månaden ${tax}, i ${priceText(lasts, l)}, från pensionsstart vid ${from} år om du tar ut tjänstepensionen, den individuella tjänstepensionen, den privata pensionen och det privata sparandet under 5, 10, 15 eller 20 år eller livsvarigt. Allmän pension betalas alltid ut livet ut. Kortare tid ger mer per månad men tar slut.`,
          `This is what you get a month ${tax}, in ${priceText(lasts, l)}, from the start at ${from} if you draw the occupational pension, the individual occupational pension, the private pension and the private savings over 5, 10, 15 or 20 years or for life. The public pension is always paid for life. A shorter time pays more a month but runs out.`,
        ),
      ),
    );

    const ribbonBox = h("div", "pa-ribbon-box");
    ribbonBox.append(h("div", "pa-ribbon-label", say(l, "Utbetalningstid och pension per månad", "Payout period and pension per month")));
    const ribbon = h("ol", "pa-ribbon pl-ribbon");
    ribbon.dataset.role = "pm-ribbon";
    for (const option of shownOptions(lasts)) {
      const { during, after } = option.view;
      const item = h("li", "pa-ribbon-item pl-ribbon-item");
      item.dataset.payout = String(option.payout);
      item.append(
        chip(option.payout, payoutName(option.payout, l)),
        h("span", "pl-ribbon-d", during === null ? "" : amount(during.income, l)),
        h(
          "span",
          "pa-ribbon-m",
          option.payout === LIFELONG
            ? say(l, "livet ut", "for life")
            : after === null
              ? say(l, `till ${to} år`, `to ${to}`)
              : say(l, `därefter ${amount(after.income, l)}`, `then ${amount(after.income, l)}`),
        ),
      );
      ribbon.append(item);
    }
    ribbonBox.append(ribbon);
    ribbonBox.append(
      h(
        "div",
        "pa-ribbon-label",
        say(
          l,
          "Siffran överst är pensionen per månad under själva utbetalningen (i snitt); under den vad som återstår därefter, när bara allmän pension är kvar.",
          "The figure on top is the pension a month during the payout itself (on average); beneath it what is left afterwards, when only the public pension remains.",
        ),
      ),
    );
    body.append(ribbonBox);

    const list = h("ul", "pa-points");
    const bullet = (text: string): void => {
      list.append(h("li", undefined, text));
    };
    if (!lasts.flexible) {
      bullet(
        say(
          l,
          "Du har ingen tjänstepension, privat pension eller annat sparande att välja utbetalningstid för, så allmän pension är allt. Välj ett avtalsområde eller fyll i kapital och sparande för att se vad utbetalningstiden gör.",
          "You have no occupational pension, private pension or other saving to choose a payout period for, so the public pension is all there is. Pick an agreement area, or fill in capital and saving, to see what the payout period does.",
        ),
      );
    } else {
      const shortDuring = short.view.during?.income ?? 0;
      const lifeDuring = life.view.during?.income ?? 0;
      bullet(
        say(
          l,
          `5 års utbetalning ger ${amount(shortDuring, l)} ${tax} under utbetalningen, livsvarig ${amount(lifeDuring, l)}.`,
          `A 5-year payout gives ${amount(shortDuring, l)} ${tax} during the payout, a lifelong one ${amount(lifeDuring, l)}.`,
        ),
      );
      const floor = lasts.publicOnly.during;
      if (floor !== null) {
        bullet(
          say(
            l,
            `När en tidsbegränsad utbetalning är slut återstår bara allmän pension: i snitt ${amount(floor.income, l)} ${tax}.`,
            `When a temporary payout ends only the public pension is left: on average ${amount(floor.income, l)} ${tax}.`,
          ),
        );
      }
      bullet(
        say(
          l,
          `Summerat till ${to} år blir det ${mkr(totalPaid(short.view), l)} Mkr med 5 års utbetalning och ${mkr(totalPaid(life.view), l)} Mkr med livsvarig, ${tax}.`,
          `Added up to ${to} that is ${mkr(totalPaid(short.view), l)} MSEK with a 5-year payout and ${mkr(totalPaid(life.view), l)} MSEK with a lifelong one, ${tax}.`,
        ),
      );
    }
    body.append(list);
    return section;
  }

  function amountPanel(lasts: Lasts, l: Lang): HTMLElement {
    const afterTax = lasts.input.afterTax;
    const actions = [
      csvExportButton(l, say(l, "belopp-per-manad.csv", "monthly-amounts.csv"), () => amountsToCsv(lasts, l)),
      xlsxExportButton(l, say(l, "belopp-per-manad.xlsx", "monthly-amounts.xlsx"), () => amountsToXlsxRows(lasts, l)),
    ];
    const { section, body } = panel(say(l, "Pension per månad och utbetalningstid", "Pension per month and payout period"), actions);
    section.dataset.role = "pm-options-panel";
    body.append(
      h(
        "p",
        "pa-lead",
        say(
          l,
          `Per månad ${taxText(afterTax, l)}, i ${priceText(lasts, l)}. "Under" är själva utbetalningen (i snitt), "efter" åren därefter då bara allmän pension återstår.`,
          `Per month ${taxText(afterTax, l)}, in ${priceText(lasts, l)}. "During" is the payout itself (on average), "after" the years that follow when only the public pension is left.`,
        ),
      ),
    );
    const head = [
      headCell(say(l, "Utbetalningstid", "Payout period")),
      headCell(say(l, "Under utbetalningen", "During the payout")),
      ...(afterTax
        ? []
        : [
            headCell(say(l, "varav allmän", "of which public")),
            headCell(say(l, "varav tjänstepension", "of which occupational")),
            headCell(say(l, "varav individuell och privat", "of which individual and private")),
            ...(lasts.taxFree ? [headCell(say(l, "varav privat sparande, uppräknat", "of which private savings, grossed up"))] : []),
          ]),
      headCell(say(l, "Efter utbetalningen", "After the payout")),
      headCell(say(l, `Totalt till ${lasts.to} år, tkr`, `In all to ${lasts.to}, thousand kr`)),
    ];
    const average = (option: PayoutOption, pick: (o: PayoutOption) => readonly number[]): number => {
      const during = option.view.during;
      if (during === null) return 0;
      let sum = 0;
      for (let age = during.from; age < during.to; age += 1) sum += pick(option)[age] ?? 0;
      return sum / during.years / 12;
    };
    const rows: HTMLTableRowElement[] = [];
    {
      const a = lasts.publicOnly;
      const cells: HTMLTableCellElement[] = [cell(say(l, "Bara allmän pension", "Public pension only"), "pl-floor-name"), cell(kronor(a.during?.income ?? 0, l), "num")];
      if (!afterTax) cells.push(...Array.from({ length: lasts.taxFree ? 4 : 3 }, () => cell("–", "num")));
      cells.push(cell("–", "num"), cell(kronor(Math.round(totalPaid(a) / 1000), l), "num"));
      rows.push(row(cells, { payout: "floor" }));
    }
    for (const option of shownOptions(lasts)) {
      const a = option.view;
      const cells: HTMLTableCellElement[] = [payoutCell(option.payout, payoutName(option.payout, l)), cell(kronor(a.during?.income ?? 0, l), "num")];
      if (!afterTax) {
        cells.push(
          cell(kronor(average(option, (o) => o.run.publicPension), l), "num"),
          cell(kronor(average(option, (o) => o.run.occupational), l), "num"),
          cell(kronor(average(option, (o) => o.run.individual), l), "num"),
          ...(lasts.taxFree ? [cell(kronor(average(option, (o) => o.savings), l), "num")] : []),
        );
      }
      cells.push(cell(a.after === null ? "–" : kronor(a.after.income, l), "num"), cell(kronor(Math.round(totalPaid(a) / 1000), l), "num"));
      rows.push(row(cells, { payout: String(option.payout) }));
    }
    body.append(tableOf("pm-amounts", "pl-options", head, rows));
    return section;
  }

  function amountNotes(lasts: Lasts, l: Lang): HTMLElement {
    const details = h("details", "field-note pa-notes");
    details.dataset.role = "pm-notes";
    details.append(h("summary", undefined, say(l, "Antaganden och metod", "Assumptions and method")));
    const bodyBox = h("div", "field-note-body");
    const savings = savingsNotes(l, lasts.runs.savingYear);
    const paragraphs = [
      say(
        l,
        `Pensionen är allmän pension (inkomstpension, premiepension, garantipension och inkomstpensionstillägg), tjänstepension från avtalsområdet och eventuell individuell tjänstepension, privat pension och privat sparande, från pensionsstart till och med året före ${lasts.to}-årsdagen. Belopp i ${priceText(lasts, l)}, utan diskontering. Bidrag som bostadstillägg räknas inte med; efter skatt är modellens nettoinkomst.`,
        `The pension is the public pension (income pension, premium pension, guarantee pension and the income-pension supplement), occupational pension from the agreement area and any individual occupational pension, private pension and private savings, from the start up to the year before the ${lasts.to}th birthday. Amounts are in ${priceText(lasts, l)}, not discounted. Benefits such as housing supplement are not counted; after tax is the model's net income.`,
      ),
      say(
        l,
        `Utbetalningstiden är modellens inställning "Temporärt uttag" för tjänstepension och privat sparande: pensionen betalas ut jämnt under 5–20 hela år från pensionsstart och slutar sedan, eller livsvarigt. Allmän pension påverkas inte. Kortare tid ger mer per månad eftersom samma kapital delas på färre år; hur mycket styrs av avtalets egna antaganden om ränta och livslängd.`,
        `The payout period is the model's "Temporary withdrawal" setting for occupational pension and private saving: the pension is paid evenly over 5–20 whole years from the start and then stops, or for life. The public pension is not affected. A shorter time pays more a month because the same capital is shared over fewer years; how much is set by the agreement's own assumptions about interest and life expectancy.`,
      ),
      `${savings.pension} ${say(l, "Allt betalas ut över samma period som avtalets tjänstepension.", "Everything is paid out over the same period as the agreement's occupational pension.")}`,
      savings.savings,
      say(
        l,
        `"Under utbetalningen" är genomsnittet per månad över utbetalningsåren; "efter" är genomsnittet de åren som följer fram till ${lasts.to} år, då bara allmän pension återstår. Modellens skatt är högre åren innan det förhöjda grundavdraget för äldre börjar, så de första åren kan ligga lägre än de följande; det visas som det är. Fliken har egna indata och följer inte formuläret till vänster; "Hämta från Prognos" kopierar dem en gång. Skatteinställningar och löneprofil kommer från Avancerat; prisbasen är alltid fasta priser, och tjänstepensionen börjar alltid vid pensionsåldern.`,
        `"During the payout" is the monthly average over the years of the payout; "after" is the average over the years that follow up to ${lasts.to}, when only the public pension is left. The model's tax is higher in the years before the higher basic allowance for older people begins, so the first years can lie lower than the ones after; it is shown as it is. The tab has its own inputs and does not follow the form on the left; "Use the Forecast values" copies them once. Tax settings and wage profile come from Advanced; prices are always constant, and the occupational pension always starts at the retirement age.`,
      ),
      ...(lasts.taxFree
        ? [
            say(
              l,
              "Privat sparande har ingen utbetalning för livet: Livsvarigt betyder här att pengarna fördelas över din förväntade återstående livslängd (modellens regel), och sedan är de slut. En pension betalas i stället ut så länge du lever.",
              "Private savings have no payout for life: Lifelong here means the money is spread over your expected remaining life (the model's rule), and then it is gone. A pension, by contrast, is paid for as long as you live.",
            ),
            say(
              l,
              "Före skatt räknas utbetalningen från det privata sparandet upp till den bruttoinkomst som ger lika mycket efter skatt (med modellens skatt, år för år). Annars skulle en skattefri krona räknas som en beskattad och vyerna före och efter skatt ge olika svar. Efter skatt visas utbetalningen som den är.",
              "Before tax, the payout from the private savings is grossed up to the gross income that leaves the same amount after tax (with the model's tax, year by year). Otherwise an untaxed krona would count as a taxed one and the views before and after tax would give different answers. After tax the payout is shown as it is.",
            ),
          ]
        : []),
    ];
    for (const text of paragraphs) bodyBox.append(h("p", undefined, text));
    details.append(bodyBox);
    return details;
  }

  function buildResults(lasts: Lasts, l: Lang): HTMLElement[] {
    const out: HTMLElement[] = [];
    if (lasts.runs.warnings.length > 0) out.push(warningsBox(lasts.runs.warnings));
    if (variant === "amount") {
      out.push(amountAnswer(lasts, l), incomeFigure(lasts, l), timePanel(lasts, l), amountPanel(lasts, l), amountNotes(lasts, l));
      return out;
    }
    out.push(
      answerPanel(lasts, l),
      incomeFigure(lasts, l),
      timePanel(lasts, l),
      optionsPanel(lasts, l),
      balanceFigure(lasts, l),
      matrixPanel(lasts, l),
      notesPanel(lasts, l),
    );
    return out;
  }

  /** The engine runs for the input, reused while nothing the engine depends on has changed. */
  function runsFor(effective: LastsInput, shared: ModelContext, shownDeaths: DeathProbabilities): PayoutRuns {
    const key = runsKey(effective);
    if (cache !== undefined && cache.key === key && cache.context === shared && cache.deaths === shownDeaths) return cache.runs;
    const runs = runPayouts(effective, shared, shownDeaths);
    cache = { key, context: shared, deaths: shownDeaths, runs };
    return runs;
  }

  function redraw(): void {
    if (context === undefined || deaths === undefined) return;
    let effective = input;
    if (salaryAs === "final") {
      const key = `${runsKey({ ...input, monthlySalary: 0 })}|${finalTarget}`;
      if (salaryCache === undefined || salaryCache.key !== key || salaryCache.context !== context) {
        salaryCache = { key, context, value: monthlySalaryForTarget(input, finalTarget, context, deaths) };
      }
      effective = { ...input, monthlySalary: salaryCache.value };
    }
    const runs = runsFor(effective, context, deaths);
    const lasts = evaluate(runs, effective);
    derived = { finalSalary: runs.finalSalary, monthlySalary: effective.monthlySalary };
    relabelSalary(currentLang);
    savingsUi.relabel(currentLang);
    results.replaceChildren(...buildResults(lasts, currentLang));
  }

  // ---- language ---------------------------------------------------------------

  const applyText = (l: Lang): void => {
    inputsTitle.textContent =
      variant === "lasts" ? say(l, "Hur länge räcker pensionen?", "How long does the pension last?") : say(l, "Hur mycket får jag?", "How much do I get?");
    intro.textContent = variant === "amount" ? say(
      l,
      "Se hur mycket pengar du får i månaden om du tar ut tjänstepensionen, den individuella tjänstepensionen, den privata pensionen och det privata sparandet under 5, 10, 15 eller 20 år eller livsvarigt. Allmän pension betalas alltid ut livet ut. Ändra personen, avtalsområdet, sparandet eller antagandena så räknas allt om direkt. Fliken har egna indata och påverkas bara av inställningarna under Avancerat.",
      "See how much money you get a month if you draw the occupational pension, the individual occupational pension, the private pension and the private savings over 5, 10, 15 or 20 years or for life. The public pension is always paid for life. Change the person, the agreement area, the savings or the assumptions and everything is recalculated at once. The tab has its own inputs and is affected only by the settings under Advanced.",
    ) : say(
      l,
      "Sätt din pension mot en budget och se hur länge den räcker, beroende på hur du tar ut tjänstepensionen, den individuella tjänstepensionen, den privata pensionen och det privata sparandet: under 5–20 år eller livsvarigt. Allmän pension betalas alltid ut livet ut. Ändra personen, avtalsområdet, budgeten eller antagandena så räknas allt om direkt. Fliken har egna indata och påverkas bara av inställningarna under Avancerat.",
      "Set your pension against a budget and see how long it lasts, depending on how you draw the occupational pension, the individual occupational pension, the private pension and the private savings: over 5–20 years or for life. The public pension is always paid for life. Change the person, the agreement area, the budget or the assumptions and everything is recalculated at once. The tab has its own inputs and is affected only by the settings under Advanced.",
    );
    importButton.textContent = say(l, "Hämta från Prognos", "Use the Forecast values");
    resetButton.textContent = say(l, "Återställ", "Reset");
    for (const { preset, button } of presetButtons) {
      button.textContent = say(l, preset.sv, preset.en);
      button.title = say(
        l,
        `${decimals(preset.inflation * 100, 1, "sv")} % inflation, ${decimals(preset.growth * 100, 1, "sv")} % real tillväxt, ${decimals(preset.ret * 100, 1, "sv")} % real avkastning`,
        `${decimals(preset.inflation * 100, 1, "en")}% inflation, ${decimals(preset.growth * 100, 1, "en")}% real growth, ${decimals(preset.ret * 100, 1, "en")}% real return`,
      );
    }
    for (const { choice, button } of basisButtons) button.textContent = say(l, choice.sv, choice.en);
  };
  applyText(lang);
  normalise();
  sync();

  return {
    element,
    relabel(l) {
      currentLang = l;
      applyText(l);
      for (const relabel of relabels) relabel(l);
      sync();
      // Hidden, the results are drawn when the tab is shown again (`setContext`), not now.
      if (element.isConnected) redraw();
    },
    setContext(nextContext, nextDeaths) {
      context = nextContext;
      deaths = nextDeaths;
      sync();
      redraw();
    },
  };
}
