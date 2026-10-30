/**
 * "Uttagsstrategi": `strategy.ts`'s findings on the page -- what most people miss, three ways of drawing
 * the pots set side by side, and a checklist -- for a typfall of the tab's own.
 *
 * Built the way "Hur länge räcker pensionen?" is: the inputs are the tab's own and built once (so a field
 * being typed in keeps its focus), "Hämta från Prognos" copies the forecast's person across at the click,
 * the individual and private pension and the private savings are `savingsInputs.ts`'s sections, and what
 * the tab reads live is the shared context (Avancerat), which `main.ts` hands over through `setContext`.
 * A change reruns the model (the runs are kept while only the floor changes) and rebuilds the results.
 *
 * The scenarios' chart is `pensionAgeChart.ts`'s line chart, with the floor as a dashed line; the colours
 * are the comparison tab's scenario steps (`--fig-scenario-*`), validated there for both themes.
 */
import { BIRTH_YEARS, RETIREMENT_AGES, SCHEME_CHOICES, defaultContext, riktalderFor } from "@typfallsmodellen/engine";
import type { DeathProbabilities, ModelContext, SchemeId, TypfallInput } from "@typfallsmodellen/engine";

import type { Bounds, Relabel } from "./controls.js";
import { fieldSet, rangeHint, span } from "./controls.js";
import { csvExportButton, xlsxExportButton } from "./export.js";
import { kronor } from "./format.js";
import { t } from "./i18n.js";
import type { Lang } from "./i18n.js";
import { centreBounds } from "./pensionAge.js";
import { niceTicks, renderLineChart } from "./pensionAgeChart.js";
import type { ChartSeries } from "./pensionAgeChart.js";
import { decimals, h, panel, row, say, tableOf } from "./pensionAgeUi.js";
import { BUDGETS } from "./pensionLasts.js";
import { SAVINGS_KINDS, savingYear, savingsNotes } from "./privateSavings.js";
import type { SavingsKind } from "./privateSavings.js";
import { createSavingsInputs } from "./savingsInputs.js";
import {
  FLOOR_BOUNDS,
  FRONT_BOUNDS,
  LIFESPAN_BOUNDS,
  MAX_EXTRA,
  PAYOUT_KINDS,
  LIFELONG,
  LONG_LIFE,
  RENT_BOUNDS,
  checklistText,
  payoutKindName,
  payoutPhases,
  defaultStrategyInput,
  evaluateStrategy,
  insightText,
  normaliseStrategy,
  overtakenAt,
  payoutText,
  planAges,
  savingsLabel,
  runStrategy,
  scenarioDescription,
  scenarioName,
  scenariosToCsv,
  scenariosToXlsxRows,
  seriesToCsv,
  seriesToXlsxRows,
  strategyRunsKey,
  TAX_RULES,
  taxRuleText,
  totalTo,
} from "./strategy.js";
import type { ExtraPension, ExtraSaving, PayoutKind, ScenarioId, Strategy, StrategyInput, StrategyRuns } from "./strategy.js";
import { cell, headCell } from "./tables.js";

const BORN: Bounds = { ...span(BIRTH_YEARS), step: 1 };
const START_WORK: Bounds = { min: 15, max: 40, step: 1 };
const RETIREMENT: Bounds = { ...span(RETIREMENT_AGES), step: 1 };
const SALARY: Bounds = { min: 0, max: 1_000_000, step: 100 };
/** An added pension's or saving's two fields: as the first one's (`savingsInputs.ts`). */
const CAPITAL: Bounds = { min: 0, max: 20_000_000, step: 10_000 };
const SAVING: Bounds = { min: 0, max: 100_000, step: 100 };

const clampTo = (value: number, bounds: Bounds): number => Math.min(Math.max(value, bounds.min), bounds.max);

/** Each kind of money's colour in the timeline: the Portföljvärde tab's account colours, validated there. */
const KIND_COLOUR: Readonly<Record<PayoutKind, string>> = {
  income: "--fig-pv-ip",
  premium: "--fig-pv-pp",
  guarantee: "--fig-pv-fee",
  occupational: "--fig-pv-tjp",
  individual: "--fig-pv-ips",
  savings: "--fig-pv-pps",
  benefits: "--fig-pv-tax",
};

/** The settings the saving year is read from until the shared context arrives: the workbook's own. */
const DEFAULT_CONTEXT = defaultContext();

/** The same two sets of economic assumptions as the other question-mark tabs. */
const PRESETS = [
  { id: "standard", sv: "Modellens standard", en: "Model default", inflation: 0, growth: 0, ret: 0.017 },
  { id: "growth", sv: "Med löneökning", en: "With wage growth", inflation: 0.02, growth: 0.018, ret: 0.017 },
] as const;

/** One colour per scenario: the comparison tab's validated steps. */
const COLOUR: Readonly<Record<ScenarioId, string>> = {
  balanced: "--fig-scenario-0",
  front: "--fig-scenario-3",
  hybrid: "--fig-scenario-1",
};

const perMonth = (l: Lang): string => say(l, "kr/mån", "SEK/month");
const money = (value: number, l: Lang): string => kronor(Math.round(value), l);
const mkr = (value: number, l: Lang): string => decimals(value / 1e6, 2, l);

export interface StrategyHandle {
  readonly element: HTMLElement;
  relabel(lang: Lang): void;
  /** Stashes the shared context and deaths and redraws the results. */
  setContext(context: ModelContext, deaths: DeathProbabilities): void;
}

export function createStrategyPanel(lang: Lang, getForecastInput: () => TypfallInput): StrategyHandle {
  let currentLang = lang;
  let input: StrategyInput = normaliseStrategy(defaultStrategyInput());
  let riktalderLocked = true;
  let context: ModelContext | undefined;
  let deaths: DeathProbabilities | undefined;
  /** The engine runs, kept while only the floor changes. */
  let cache: { key: string; context: ModelContext; deaths: DeathProbabilities; runs: StrategyRuns } | undefined;

  const element = h("div", "strategy-panel pension-age-panel");
  element.dataset.role = "strategy-panel";

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
  importButton.dataset.action = "st-import-forecast";
  const resetButton = h("button", "export-btn");
  resetButton.type = "button";
  resetButton.dataset.action = "st-reset";
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
    prefix: "st",
    grid,
    relabels,
    lang,
    value: () => input,
    update: (patch) => update(patch),
    savingStart: () => savingStart(),
    group,
  });

  /**
   * A list of added occupational pensions or private savings: a row each -- for a saving, what it is held in --
   * with its capital and monthly saving side by side, a button to take it away, and one to add another, up to
   * `MAX_EXTRA`. The rows are rebuilt only when one is added or taken away (or the input is replaced), so a field
   * being typed in keeps its focus; a typed value only updates the input.
   */
  function extrasList<T extends ExtraPension>(options: {
    readonly name: string;
    readonly items: () => readonly T[];
    readonly set: (items: readonly T[]) => void;
    readonly blank: () => T;
    readonly title: (index: number, l: Lang) => string;
    readonly add: (l: Lang) => string;
    readonly note: (l: Lang) => string;
    readonly capitalHint?: (item: T, l: Lang) => string;
    /** A row's own fields before its capital and monthly saving (a saving's kind), with what its captions depend on. */
    readonly head?: (row: HTMLElement, item: () => T, change: (patch: Partial<T>) => void, rowRelabels: Relabel[]) => void;
  }): { sync(): void } {
    const box = h("div", "st-extras");
    box.dataset.role = `st-${options.name}`;
    const rows = h("div", "st-extra-rows");
    const note = h("p", "saving-note");
    const addButton = h("button", "export-btn st-extra-add");
    addButton.type = "button";
    addButton.dataset.action = `st-add-${options.name}`;
    box.append(rows, addButton, note);
    grid.append(box);
    let rowRelabels: Relabel[] = [];
    let shown: readonly T[] = [];

    const relabel = (l: Lang): void => {
      addButton.textContent = `+ ${options.add(l)}`;
      note.textContent = options.note(l);
      for (const r of rowRelabels) r(l);
    };
    relabels.push((l) => relabel(l));

    /** A typed value: the rows already show it, so they are not rebuilt. */
    const change = (index: number, patch: Partial<T>): void => {
      const items = options.items().map((item, i) => (i === index ? { ...item, ...patch } : item));
      shown = items;
      options.set(items);
    };

    function build(): void {
      rows.replaceChildren();
      rowRelabels = [];
      shown = options.items();
      shown.forEach((item, index) => {
        const row = h("div", "st-extra");
        row.dataset.index = String(index);
        const top = h("div", "st-extra-head");
        const title = h("h4", "st-extra-title");
        const remove = h("button", "panel-btn st-extra-remove");
        remove.type = "button";
        remove.dataset.action = `st-remove-${options.name}`;
        remove.addEventListener("click", () => options.set(options.items().filter((_, i) => i !== index)));
        top.append(title, remove);
        row.append(top);
        rowRelabels.push((l) => {
          title.textContent = options.title(index, l);
          remove.textContent = say(l, "Ta bort", "Remove");
          remove.setAttribute("aria-label", `${say(l, "Ta bort", "Remove")} ${options.title(index, l).toLowerCase()}`);
        });
        const current = () => options.items()[index] ?? item;
        options.head?.(row, current, (patch) => change(index, patch), rowRelabels);
        const pair = h("div", "field-pair");
        row.append(pair);
        const inPair = fieldSet(pair, rowRelabels, currentLang);
        const capital = inPair.number(item.capital, CAPITAL, (v) => change(index, { capital: v } as Partial<T>));
        capital.element.dataset.setting = `st-${options.name}-capital`;
        inPair.field(
          capital.element,
          (l) => ({
            label: say(l, "Kapital", "Capital"),
            hint: options.capitalHint?.(current(), l) ?? say(l, `kr · värde vid ${savingStart()} års början`, `SEK · value at the start of ${savingStart()}`),
          }),
          "field field-stack",
        );
        const monthly = inPair.number(item.monthly, SAVING, (v) => change(index, { monthly: v } as Partial<T>));
        monthly.element.dataset.setting = `st-${options.name}-monthly`;
        inPair.field(
          monthly.element,
          (l) => ({ label: say(l, "Månadssparande", "Monthly saving"), hint: say(l, "kr/mån fram till pensionen", "SEK/month until retirement") }),
          "field field-stack",
        );
        rows.append(row);
      });
      addButton.disabled = shown.length >= MAX_EXTRA;
      relabel(currentLang);
    }

    addButton.addEventListener("click", () => {
      if (options.items().length < MAX_EXTRA) options.set([...options.items(), options.blank()]);
    });
    build();
    return {
      /** Rebuilds the rows when the list is not the one they show: one added or taken away, a reset. */
      sync() {
        if (JSON.stringify(options.items()) === JSON.stringify(shown)) relabel(currentLang);
        else build();
      },
    };
  }

  // Person
  group("Person", "Person");
  const bornField = number(input.born, BORN, (born) => update({ born }));
  bornField.element.dataset.setting = "stBorn";
  const relabelBorn = field(bornField.element, (l) => ({
    label: t("birthYear", l),
    hint: say(l, `Riktålder ${riktalderFor(input.born)} år`, `${t("riktalderCheckbox", l)} ${riktalderFor(input.born)}`),
  }));
  const startField = number(input.startWorkAge, START_WORK, (startWorkAge) => update({ startWorkAge }));
  startField.element.dataset.setting = "stStartWork";
  field(startField.element, (l) => ({ label: t("startWorkAge", l), hint: rangeHint(START_WORK) }));
  const retirementField = number(input.retirementAge, RETIREMENT, (retirementAge) => update({ retirementAge }));
  retirementField.element.dataset.setting = "stRetirement";
  const relabelRetirement = field(retirementField.element, (l) => ({
    label: say(l, "Pensionsålder", "Retirement age"),
    hint: say(l, `Tidigast ${centreBounds(input.born).min} år`, `Earliest ${centreBounds(input.born).min}`),
  }));
  const riktalderBox = check(riktalderLocked, (locked) => {
    riktalderLocked = locked;
    update({});
  });
  riktalderBox.dataset.setting = "stRiktalder";
  field(riktalderBox, (l) => ({ label: t("riktalderCheckbox", l) }), "field field-check");
  /** The model's expected lifespan for the person, from the last results: what the field shows until one is typed. */
  let modelLifespan = 0;
  /** The scenario the payout timeline shows; the suggested one until another is picked. */
  let timelineScenario: ScenarioId | undefined;
  const lifespanField = number(Math.round(LIFESPAN_BOUNDS.max), LIFESPAN_BOUNDS, (lifespan) => update({ lifespan }));
  lifespanField.element.dataset.setting = "stLifespan";
  const relabelLifespan = field(lifespanField.element, (l) => ({
    label: say(l, "Förväntad livslängd (ålder)", "Expected lifespan (age)"),
    hint:
      modelLifespan > 0
        ? say(l, `år · förifyllt med modellens antagande, ${decimals(modelLifespan, 1, l)} år`, `years · filled in with the model's assumption, ${decimals(modelLifespan, 1, l)}`)
        : say(l, "år · förifyllt med modellens antagande", "years · filled in with the model's assumption"),
  }));

  // Salary
  group("Lön", "Salary");
  const salaryField = number(input.monthlySalary, SALARY, (monthlySalary) => update({ monthlySalary }));
  salaryField.element.dataset.setting = "stSalary";
  field(salaryField.element, (l) => ({ label: t("monthlySalary", l), hint: say(l, "kr/mån", "SEK/month") }));

  // Pension
  group("Pension", "Pension");
  const schemeSelect = select(
    SCHEME_CHOICES.map((choice) => ({ value: choice.value, label: () => choice.label })),
    input.scheme,
    (scheme) => update({ scheme: scheme as SchemeId }),
  );
  schemeSelect.element.dataset.setting = "stScheme";
  relabels.push(schemeSelect.relabel);
  field(schemeSelect.element, (l) => ({ label: t("occupational", l) }), "field field-wide");
  savingsUi.appendPension();
  const extraPensions = extrasList<ExtraPension>({
    name: "pension",
    items: () => input.extraPensions,
    set: (extraPensions) => update({ extraPensions }),
    blank: () => ({ capital: 0, monthly: 0 }),
    title: (index, l) => say(l, `Tjänstepension ${index + 2}`, `Occupational pension ${index + 2}`),
    add: (l) => say(l, "Lägg till tjänstepension", "Add an occupational pension"),
    note: (l) =>
      say(
        l,
        "Har du tjänstepension från fler arbetsgivare eller avtal, till exempel ett fribrev från ett tidigare jobb? Lägg till dem här. De beskattas som inkomst och betalas ut som den privata pensionen: livsvarigt eller på den kortare tiden.",
        "Do you have occupational pension from more employers or agreements, such as a paid-up policy from an earlier job? Add them here. They are taxed as income and paid out like the private pension: for life or over the shorter period.",
      ),
  });
  savingsUi.appendPrivateSavings();
  const extraSavings = extrasList<ExtraSaving>({
    name: "saving",
    items: () => input.extraSavings,
    set: (extraSavings) => update({ extraSavings }),
    blank: () => ({ kind: "isk", capital: 0, monthly: 0 }),
    title: (index, l) => say(l, `Sparande ${index + 2}`, `Saving ${index + 2}`),
    add: (l) => say(l, "Lägg till sparande", "Add a saving"),
    note: (l) =>
      say(
        l,
        "Till exempel ett ISK och en depå, eller sparande i flera banker. Sparanden av samma typ räknas ihop.",
        "For example an ISK and a securities account, or savings at more than one bank. Savings of the same type are added together.",
      ),
    capitalHint: (item, l) =>
      item.kind === "depot"
        ? say(l, `kr · efter reavinstskatt, värde vid ${savingStart()} års början`, `SEK · after capital gains tax, value at the start of ${savingStart()}`)
        : say(l, `kr · värde vid ${savingStart()} års början`, `SEK · value at the start of ${savingStart()}`),
    head: (row, item, change, rowRelabels) => {
      const inRow = fieldSet(row, rowRelabels, currentLang);
      const kind = inRow.select(
        SAVINGS_KINDS.map((k, index) => ({ value: index, label: (l: Lang) => say(l, k.sv, k.en) })),
        SAVINGS_KINDS.findIndex((k) => k.id === item().kind),
        (index) => {
          change({ kind: SAVINGS_KINDS[index]!.id as SavingsKind });
          for (const r of rowRelabels) r(currentLang);
        },
      );
      kind.element.dataset.setting = "st-saving-kind";
      rowRelabels.push(kind.relabel);
      inRow.field(kind.element, (l) => ({ label: say(l, "Typ av sparande", "Type of saving") }), "field field-wide");
    },
  });

  // Housing: what the housing supplement is worked out from.
  group("Boende", "Housing");
  const rentField = number(input.rent, RENT_BOUNDS, (rent) => update({ rent }));
  rentField.element.dataset.setting = "stRent";
  field(rentField.element, (l) => ({
    label: say(l, "Boendekostnad per månad", "Housing cost per month"),
    hint: say(l, "kr/mån · hyra eller avgift och drift", "SEK/month · rent, or fee and running costs"),
  }));
  const housingBox = check(input.housingSupplement, (housingSupplement) => update({ housingSupplement }));
  housingBox.dataset.setting = "stHousing";
  field(housingBox, (l) => ({ label: say(l, "Räkna med bostadstillägg", "Count the housing supplement") }), "field field-check");

  // The plan: the floor and the front-loaded period.
  group("Din plan", "Your plan");
  const floorField = number(input.floor, FLOOR_BOUNDS, (floor) => update({ floor }));
  floorField.element.dataset.setting = "stFloor";
  field(floorField.element, (l) => ({
    label: say(l, "Lägsta nivå per månad", "Lowest level per month"),
    hint: say(
      l,
      `kr/mån efter skatt, hela livet · minPension: ${BUDGETS.map((b) => `${b.sv.replace("pensionär", "").toLowerCase()} ${kronor(b.amount, "sv")}`).join(", ")}`,
      `SEK/month after tax, for life · minPension: ${BUDGETS.map((b) => `${b.en.replace(" pensioner", "").toLowerCase()} ${kronor(b.amount, "en")}`).join(", ")}`,
    ),
  }));
  const frontField = number(input.frontYears, FRONT_BOUNDS, (frontYears) => update({ frontYears }));
  frontField.element.dataset.setting = "stFrontYears";
  field(frontField.element, (l) => ({
    label: say(l, "Uttag på kortare tid", "Shorter payout"),
    hint: say(l, `år, ${FRONT_BOUNDS.min}–${FRONT_BOUNDS.max} · i scenarierna Framtung och Hybrid`, `years, ${FRONT_BOUNDS.min}–${FRONT_BOUNDS.max} · in the Front-loaded and Hybrid scenarios`),
  }));

  // Assumptions
  group("Antaganden", "Assumptions");
  const presetBox = h("div", "panel-toggle pa-toggle");
  presetBox.setAttribute("role", "group");
  const presetButtons = PRESETS.map((preset) => {
    const button = h("button", "panel-btn");
    button.type = "button";
    button.dataset.preset = preset.id;
    button.addEventListener("click", () => update({ yearlyInflation: preset.inflation, realGrowth: preset.growth, realReturn: preset.ret }));
    presetBox.append(button);
    return { preset, button };
  });
  grid.append(presetBox);
  const inflationField = percent(input.yearlyInflation, (yearlyInflation) => update({ yearlyInflation }));
  inflationField.element.dataset.setting = "stInflation";
  field(inflationField.element, (l) => ({ label: t("inflation", l), hint: "%" }));
  relabels.push(inflationField.relabel);
  const growthField = percent(input.realGrowth, (realGrowth) => update({ realGrowth }));
  growthField.element.dataset.setting = "stGrowth";
  field(growthField.element, (l) => ({ label: t("realGrowth", l), hint: "%" }));
  relabels.push(growthField.relabel);
  const returnField = percent(input.realReturn, (realReturn) => update({ realReturn }));
  returnField.element.dataset.setting = "stReturn";
  field(returnField.element, (l) => ({ label: t("realReturn", l), hint: "%" }));
  relabels.push(returnField.relabel);

  // ---- state -> controls -------------------------------------------------------

  function savingStart(): number {
    return savingYear(input, context ?? DEFAULT_CONTEXT);
  }

  function normalise(): void {
    const wanted = riktalderLocked ? riktalderFor(input.born) : input.retirementAge;
    input = normaliseStrategy({ ...input, retirementAge: wanted });
  }

  function sync(): void {
    bornField.setValue(input.born);
    startField.setValue(input.startWorkAge);
    retirementField.setValue(input.retirementAge);
    retirementField.element.disabled = riktalderLocked;
    riktalderBox.checked = riktalderLocked;
    salaryField.setValue(input.monthlySalary);
    schemeSelect.element.value = String(input.scheme);
    savingsUi.sync(currentLang);
    extraPensions.sync();
    extraSavings.sync();
    rentField.setValue(input.rent);
    housingBox.checked = input.housingSupplement;
    floorField.setValue(input.floor);
    frontField.setValue(input.frontYears);
    if (input.lifespan > 0 || modelLifespan > 0) lifespanField.setValue(input.lifespan > 0 ? input.lifespan : Math.round(modelLifespan));
    inflationField.setValue(input.yearlyInflation);
    growthField.setValue(input.realGrowth);
    returnField.setValue(input.realReturn);
    for (const { preset, button } of presetButtons) {
      const active = input.yearlyInflation === preset.inflation && input.realGrowth === preset.growth && input.realReturn === preset.ret;
      button.classList.toggle("active", active);
    }
    relabelBorn(currentLang);
    relabelRetirement(currentLang);
  }

  function update(patch: Partial<StrategyInput>): void {
    input = { ...input, ...patch };
    normalise();
    sync();
    redraw();
  }

  importButton.addEventListener("click", () => {
    const forecast = getForecastInput();
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
    input = defaultStrategyInput();
    riktalderLocked = true;
    normalise();
    sync();
    redraw();
  });

  // ---- results -----------------------------------------------------------------

  function insightsPanel(strategy: Strategy, l: Lang): HTMLElement {
    const { section, body } = panel(say(l, "Det de flesta missar", "What most people miss"));
    section.dataset.role = "st-insights";
    if (strategy.insights.length === 0) {
      body.append(
        h(
          "p",
          "pa-lead",
          say(
            l,
            "Inga dolda effekter på över 1 000 kr hittades för dig. Scenarierna nedan visar ändå hur uttagen kan fördelas.",
            "No hidden effects of more than 1,000 SEK were found for you. The scenarios below still show how the withdrawals can be spread.",
          ),
        ),
      );
      return section;
    }
    const list = h("div", "st-insights");
    strategy.insights.forEach((insight, index) => {
      const text = insightText(insight, l);
      const card = h("article", "st-insight");
      card.dataset.insight = insight.kind;
      const number = h("span", "st-insight-number", String(index + 1));
      number.setAttribute("aria-hidden", "true");
      const title = h("h3", "st-insight-title", text.title);
      const bodyText = h("p", "st-insight-body", text.body);
      const action = h("p", "st-insight-action");
      action.append(h("strong", undefined, say(l, "Gör så här: ", "What to do: ")), document.createTextNode(text.action));
      card.append(number, title, bodyText, action);
      list.append(card);
    });
    body.append(list);
    return section;
  }

  function recommendationText(strategy: Strategy, l: Lang): string {
    const chosen = strategy.scenarios.find((s) => s.id === strategy.recommended)!;
    const floor = money(strategy.input.floor, l);
    if (!chosen.meetsFloor) {
      return say(
        l,
        `Inget scenario håller ${floor} ${perMonth(l)} hela livet. ${scenarioName(chosen.id, l)} kommer närmast: lägst ${money(chosen.lowest.monthly, l)} ${perMonth(l)}.`,
        `No scenario keeps ${floor} ${perMonth(l)} for life. ${scenarioName(chosen.id, l)} comes closest: at least ${money(chosen.lowest.monthly, l)} ${perMonth(l)}.`,
      );
    }
    return say(
      l,
      `Förslag: ${scenarioName(chosen.id, l)}. Det ger mest de första tio åren, ${money(chosen.early, l)} ${perMonth(l)}, och blir aldrig lägre än ${money(chosen.lowest.monthly, l)} ${perMonth(l)} – över din lägsta nivå på ${floor} ${perMonth(l)}.`,
      `Suggestion: ${scenarioName(chosen.id, l)}. It pays the most in the first ten years, ${money(chosen.early, l)} ${perMonth(l)}, and never falls below ${money(chosen.lowest.monthly, l)} ${perMonth(l)} – above your lowest level of ${floor} ${perMonth(l)}.`,
    );
  }

  function scenariosPanel(strategy: Strategy, l: Lang): HTMLElement {
    const { section, body } = panel(say(l, "Strategiska scenarier", "Strategic scenarios"), [
      csvExportButton(l, say(l, "uttagsstrategi.csv", "withdrawal-strategy.csv"), () => scenariosToCsv(strategy, l)),
      xlsxExportButton(l, say(l, "uttagsstrategi.xlsx", "withdrawal-strategy.xlsx"), () => scenariosToXlsxRows(strategy, l)),
    ]);
    section.dataset.role = "st-scenarios";
    const lead = h("p", "pa-lead st-recommendation", recommendationText(strategy, l));
    lead.dataset.recommended = strategy.recommended;
    body.append(lead);

    const first = strategy.scenarios[0]!;
    /** A header with its unit on a line of its own, small. */
    const unitCell = (text: string, unit: string, info?: string): HTMLTableCellElement => {
      const th = headCell(text, info);
      th.append(h("small", undefined, unit));
      return th;
    };
    // The totals at 80, at the expected lifespan and at a long life: 90 is in the payout table below.
    const shownTotals = first.totals.filter((t) => t.lifespan !== 90 || t.lifespan === Math.round(strategy.expectedDeath));
    const head = [
      headCell(say(l, "Scenario", "Scenario")),
      ...first.phases.map((p) =>
        unitCell(say(l, `${p.from}–${p.to} år`, `Ages ${p.from}–${p.to}`), perMonth(l), say(l, "Att leva på per månad, i snitt", "To live on a month, on average")),
      ),
      ...shownTotals.map((t) =>
        unitCell(
          say(l, `Till ${t.lifespan} år`, `To ${t.lifespan}`),
          say(l, "Mkr totalt", "MSEK in all"),
          t.lifespan === Math.round(strategy.expectedDeath)
            ? say(l, "Till din förväntade livslängd", "To your expected lifespan")
            : undefined,
        ),
      ),
      unitCell(say(l, "Lägsta månad", "Lowest month"), perMonth(l), say(l, `Fram till ${LONG_LIFE} år`, `Up to ${LONG_LIFE}`)),
      unitCell(
        say(l, "Balanserad går om", "Balanced catches up"),
        say(l, "vid livslängd", "at lifespan"),
        say(l, "Den livslängd från vilken Balanserad har gett mer totalt", "The lifespan from which Balanced has paid more in all"),
      ),
    ];
    const rows = strategy.scenarios.map((s) => {
      const name = cell("", "st-scenario-name");
      const key = h("span", "pa-key");
      key.style.background = `var(${COLOUR[s.id]})`;
      name.append(key, h("strong", undefined, scenarioName(s.id, l)));
      if (s.id === strategy.recommended) name.append(h("span", "st-chip", say(l, "★ Förslag", "★ Suggested")));
      name.append(h("small", undefined, scenarioDescription(s.id, strategy.input.frontYears, l)));
      const lowest = cell(money(s.lowest.monthly, l), `num${s.meetsFloor ? "" : " st-below"}`);
      if (!s.meetsFloor) lowest.title = say(l, "Under din lägsta nivå", "Below your lowest level");
      return row(
        [
          name,
          ...s.phases.map((p) => cell(money(p.monthly, l), "num")),
          ...s.totals.filter((t) => shownTotals.some((shown) => shown.lifespan === t.lifespan)).map((t) => cell(mkr(t.total, l), "num")),
          lowest,
          cell(s.id === "balanced" ? "–" : s.breakEven === undefined ? say(l, "inte före 100", "not before 100") : say(l, `${s.breakEven} år`, `${s.breakEven}`), "num"),
        ],
        { scenario: s.id },
      );
    });
    body.append(tableOf("st-scenario-table", "st-table", head, rows), scenarioFigure(strategy, l), timeline(strategy, l), payoutTable(strategy, l));
    const downloads = h("div", "panel-actions st-downloads");
    downloads.append(
      csvExportButton(l, say(l, "att-leva-pa-per-ar.csv", "to-live-on-per-age.csv"), () => seriesToCsv(strategy, l)),
      xlsxExportButton(l, say(l, "att-leva-pa-per-ar.xlsx", "to-live-on-per-age.xlsx"), () => seriesToXlsxRows(strategy, l)),
    );
    body.append(downloads);
    return section;
  }

  /**
   * When each kind of money is paid, from the retirement year, as a Gantt chart: one row per kind, a bar per stretch
   * of years, the expected lifespan marked. A switch picks the scenario; the suggested one is shown first.
   */
  function timeline(strategy: Strategy, l: Lang): HTMLElement {
    const figure = h("figure", "figure st-timeline");
    figure.dataset.role = "st-fig-timeline";
    const ids = strategy.scenarios.map((s) => s.id);
    if (timelineScenario === undefined || !ids.includes(timelineScenario)) timelineScenario = strategy.recommended;
    const head = h("div", "st-timeline-head");
    head.append(h("h3", "st-subhead", say(l, "Utbetalningarnas tidslinje", "When each payout is paid")));
    const toggle = h("div", "panel-toggle pa-toggle");
    toggle.setAttribute("role", "group");
    const draw = h("div", "scroll");
    const note = h("p", "pa-lead");
    const render = (): void => {
      const s = strategy.scenarios.find((x) => x.id === timelineScenario)!;
      for (const button of toggle.querySelectorAll<HTMLButtonElement>("button")) {
        const on = button.dataset.scenario === s.id;
        button.classList.toggle("active", on);
        button.setAttribute("aria-pressed", String(on));
      }
      draw.replaceChildren(timelineSvg(strategy, s.id, l));
      note.textContent = say(
        l,
        `${scenarioName(s.id, l)}: ${scenarioDescription(s.id, strategy.input.frontYears, l)} Staplarna visar från vilken ålder till vilken varje del betalas ut, i snitt per månad före skatt. Där staplarna ligger över varandra betalas delarna ut samtidigt.`,
        `${scenarioName(s.id, l)}: ${scenarioDescription(s.id, strategy.input.frontYears, l)} The bars show from which age to which each part is paid, on average a month before tax. Where bars lie above one another, the parts are paid at the same time.`,
      );
    };
    for (const id of ids) {
      const button = h("button", "panel-btn", scenarioName(id, l));
      button.type = "button";
      button.dataset.scenario = id;
      button.addEventListener("click", () => {
        timelineScenario = id;
        render();
      });
      toggle.append(button);
    }
    head.append(toggle);
    figure.append(head, note, draw);
    render();
    return figure;
  }

  function timelineSvg(strategy: Strategy, id: ScenarioId, l: Lang): SVGSVGElement {
    const NS = "http://www.w3.org/2000/svg";
    const el = <K extends keyof SVGElementTagNameMap>(tag: K, attrs: Record<string, string | number>, text?: string) => {
      const node = document.createElementNS(NS, tag);
      for (const [k, v] of Object.entries(attrs)) node.setAttribute(k, String(v));
      if (text !== undefined) node.textContent = text;
      return node;
    };
    const s = strategy.scenarios.find((x) => x.id === id)!;
    const phases = payoutPhases(s.run);
    const kinds = PAYOUT_KINDS.filter((k) => phases.some((p) => p.kind === k));
    const par = strategy.par;
    const end = LONG_LIFE + 1;
    const W = 760;
    const left = 175;
    const right = 24;
    const rowH = 44;
    const top = 26;
    const H = top + kinds.length * rowH + 34;
    const x = (age: number) => left + ((age - par) / (end - par)) * (W - left - right);
    const svg = el("svg", { viewBox: `0 0 ${W} ${H}`, class: "st-timeline-svg", role: "img" });
    svg.setAttribute(
      "aria-label",
      phases
        .map((p) => `${payoutKindName(p.kind, l, savingsLabel(strategy.input))}: ${p.from}–${p.lifelong ? say(l, "livsvarigt", "for life") : p.to}`)
        .join("; "),
    );
    // The grid: every fifth age, with its year.
    for (let age = par; age <= end; age += 1) {
      if ((age - par) % 5 !== 0 && age !== end) continue;
      const gx = x(age);
      svg.append(el("line", { x1: gx, x2: gx, y1: top - 6, y2: H - 30, class: "st-tl-grid" }));
      if (age < end) {
        svg.append(el("text", { x: gx, y: H - 16, class: "st-tl-tick", "text-anchor": "middle" }, say(l, `${age} år`, `${age}`)));
        svg.append(el("text", { x: gx, y: H - 3, class: "st-tl-tick st-tl-year", "text-anchor": "middle" }, String(strategy.input.born + age)));
      }
    }
    // The expected lifespan.
    const life = strategy.expectedDeath;
    if (life > par && life < end) {
      const lx = x(life);
      svg.append(el("line", { x1: lx, x2: lx, y1: 12, y2: H - 30, class: "st-tl-life" }));
      svg.append(
        el("text", { x: lx, y: 10, class: "st-tl-tick", "text-anchor": lx > W - 140 ? "end" : "middle" }, say(l, `Förväntad livslängd ${decimals(life, 1, l)} år`, `Expected lifespan ${decimals(life, 1, l)}`)),
      );
    }
    kinds.forEach((kind, row) => {
      const y = top + row * rowH;
      svg.append(el("text", { x: 0, y: y + 13, class: "st-tl-label" }, payoutKindName(kind, l, savingsLabel(strategy.input))));
      for (const p of phases.filter((q) => q.kind === kind)) {
        const x0 = x(p.from);
        const x1 = x(p.to + 1);
        const bar = el("rect", { x: x0, y: y + 1, width: Math.max(x1 - x0, 3), height: 16, rx: 4, fill: `var(${KIND_COLOUR[kind]})`, class: "st-tl-bar" });
        bar.dataset.kind = kind;
        const span = p.lifelong
          ? say(l, `från ${p.from} år, livsvarigt`, `from ${p.from}, for life`)
          : say(l, `${p.from}–${p.to} år (${p.to - p.from + 1} år)`, `${p.from}–${p.to} (${p.to - p.from + 1} years)`);
        const text = `${span} · ${say(l, "ca", "about")} ${kronor(Math.round(p.monthly / 10) * 10, l)} ${say(l, "kr/mån", "SEK/month")}`;
        bar.append(el("title", {}, `${payoutKindName(kind, l, savingsLabel(strategy.input))}: ${text}`));
        svg.append(bar);
        if (p.lifelong) svg.append(el("path", { d: `M${x1 - 1} ${y + 1} l7 8 l-7 8 z`, fill: `var(${KIND_COLOUR[kind]})` }));
        const tx = Math.min(x0, W - right - 230);
        svg.append(el("text", { x: tx, y: y + 32, class: "st-tl-text" }, text));
      }
    });
    return svg;
  }

  function scenarioFigure(strategy: Strategy, l: Lang): HTMLElement {
    const from = strategy.par;
    const to = LONG_LIFE;
    const ages = Array.from({ length: to - from + 1 }, (_, i) => from + i);
    const format = (value: number) => `${kronor(Math.round(value), l)} ${perMonth(l)}`;
    const series: ChartSeries[] = strategy.scenarios.map((s) => ({
      name: scenarioName(s.id, l),
      end: scenarioName(s.id, l),
      colour: COLOUR[s.id],
      emphasis: s.id === strategy.recommended,
      essential: true,
      points: ages.map((age) => [age, s.run.disposable[age]! / 12] as const),
      format,
    }));
    series.push({
      name: say(l, "Din lägsta nivå", "Your lowest level"),
      end: say(l, "Golv", "Floor"),
      colour: "--text-muted",
      dashed: true,
      essential: true,
      points: ages.map((age) => [age, strategy.input.floor] as const),
      format,
    });
    const top = Math.max(1, ...series.flatMap((s) => s.points.map((p) => p[1])));
    return renderLineChart({
      role: "st-fig-scenarios",
      lang: l,
      title: say(l, "Att leva på per månad", "To live on per month"),
      subtitle: say(
        l,
        `Pension efter skatt, bostadstillägg och sparpengar, i fasta priser. Den streckade linjen är din lägsta nivå. ${strategy.input.lifespan > 0 ? `Du räknar med att leva till ${strategy.input.lifespan} år (modellen: ${decimals(strategy.modelExpectedDeath, 1, l)} år).` : `Modellens förväntade livslängd för dig är ${decimals(strategy.modelExpectedDeath, 1, l)} år.`}`,
        `Pension after tax, housing supplement and savings, in constant prices. The dashed line is your lowest level. ${strategy.input.lifespan > 0 ? `You count on living to ${strategy.input.lifespan} (the model: ${decimals(strategy.modelExpectedDeath, 1, l)}).` : `The model's expected lifespan for you is ${decimals(strategy.modelExpectedDeath, 1, l)}.`}`,
      ),
      xMin: from,
      xMax: to,
      xTitle: say(l, "Ålder", "Age"),
      yTicks: niceTicks(0, top, 5).ticks,
      autoTicks: 5,
      yTick: (v) => kronor(v, l),
      yUnit: say(l, "kr per månad", "SEK per month"),
      heading: (x) => `${say(l, "Ålder", "Age")} ${x}`,
      series,
      endRoom: 84,
    });
  }

  /** The occupational pension's payout periods: in all to each lifespan, and from where life catches up. */
  function payoutTable(strategy: Strategy, l: Lang): HTMLElement {
    const box = h("div", "st-payouts");
    if (strategy.payouts.length < 2) return box;
    const life = strategy.payouts[strategy.payouts.length - 1]!;
    const lifespans = planAges(life, strategy.expectedDeath);
    const expected = Math.round(strategy.expectedDeath);
    const best = strategy.payouts.reduce((a, b) => (totalTo(b, expected) > totalTo(a, expected) + 0.5 ? b : a));
    box.append(
      h("h3", "st-subhead", say(l, "Tjänstepensionens uttagstid", "The occupational pension's payout period")),
      h(
        "p",
        "pa-lead",
        say(
          l,
          "Allt att leva på totalt, i miljoner kronor, om tjänstepensionen tas ut på olika lång tid och resten livsvarigt.",
          "All there is to live on in total, in millions of kronor, if the occupational pension is drawn over different periods and the rest for life.",
        ),
      ),
    );
    const head = [
      headCell(say(l, "Uttagstid", "Payout period")),
      ...lifespans.map((a) => headCell(say(l, `Till ${a} år`, `To ${a}`))),
      headCell(say(l, "Livsvarigt går om vid", "For life catches up at")),
    ];
    const rows = strategy.payouts.map((r) => {
      const isBest = r === best;
      const name = cell(payoutText(r.combo.occupational, l), isBest ? "st-best" : undefined);
      const caught = r.combo.occupational === LIFELONG ? "–" : (() => {
        const at = overtakenAt(r, life);
        return at === undefined ? say(l, "inte före 100 år", "not before 100") : say(l, `${at} år`, `${at}`);
      })();
      return row(
        [name, ...lifespans.map((a) => cell(mkr(totalTo(r, a), l), `num${isBest && a === expected ? " st-best" : ""}`)), cell(caught, "num")],
        { payout: String(r.combo.occupational) },
      );
    });
    box.append(tableOf("st-payout-table", "st-table", head, rows));
    return box;
  }

  /** The tax rules: the tax-free income, ROT and RUT, interest, and the other tax reductions, each with its page. */
  function taxPanel(strategy: Strategy, l: Lang): HTMLElement {
    const { tax } = strategy;
    const { section, body } = panel(say(l, "Skatt och skattereduktioner", "Tax and tax reductions"));
    section.dataset.role = "st-tax";
    body.append(
      h(
        "p",
        "pa-lead",
        say(
          l,
          `Fyra regler som avgör hur mycket av pensionen du får behålla. Beloppen är Skatteverkets för ${tax.year} och ändras varje år. Siffrorna för dig kommer från förslaget, i fasta priser.`,
          `Four rules that decide how much of the pension you keep. The amounts are Skatteverket's for ${tax.year} and change every year. Your own figures come from the suggestion, in constant prices.`,
        ),
      ),
    );
    const list = h("div", "st-insights st-rules");
    for (const rule of TAX_RULES) {
      const text = taxRuleText(rule, tax, strategy.input, l);
      const card = h("article", "st-insight st-rule");
      card.dataset.rule = rule;
      card.append(h("h3", "st-insight-title", text.title), h("p", "st-insight-body", text.body));
      if (text.mine !== "") {
        const mine = h("p", "st-insight-action");
        mine.append(h("strong", undefined, say(l, "För dig: ", "For you: ")), document.createTextNode(text.mine));
        card.append(mine);
      }
      const link = h("a", "st-rule-link", text.link.label);
      link.href = text.link.href;
      link.target = "_blank";
      link.rel = "noopener noreferrer";
      card.append(link);
      list.append(card);
    }
    body.append(list);
    return section;
  }

  function checklistPanel(strategy: Strategy, l: Lang): HTMLElement {
    const { section, body } = panel(say(l, "Checklista för den här månaden", "Checklist for this month"));
    section.dataset.role = "st-checklist";
    const list = h("ol", "st-checklist");
    for (const item of strategy.checklist) {
      const li = h("li", undefined, checklistText(item, l));
      li.dataset.item = item.kind;
      list.append(li);
    }
    body.append(list);
    return section;
  }

  function notesPanel(strategy: Strategy, l: Lang): HTMLElement {
    const { section, body } = panel(say(l, "Så har vi räknat", "How this is computed"));
    section.dataset.role = "st-notes";
    const notes = savingsNotes(l, strategy.savingYear);
    const list = h("ul", "pa-notes");
    for (const text of [
      say(
        l,
        "Att leva på är pensionen efter skatt, bostadstillägget och de andra bidrag modellen räknar med, och det privata sparandet som tas ut skattefritt. Allt i fasta priser, som i de andra frågeflikarna. Det första året är samma belopp som Prognos visar vid pensioneringen (Tabell 1).",
        "To live on is the pension after tax, the housing supplement and the other benefits the model counts, and the private savings paid out tax-free. Everything in constant prices, as in the other question-mark tabs. The first year is the amount Prognos shows at retirement (Table 1).",
      ),
      say(
        l,
        "Bostadstillägget prövas mot det privata sparandet år för år, i takt med att det tas ut. Förmögenhet upp till 100 000 kr räknas bort. Bostaden du bor i räknas inte.",
        "The housing supplement is tested against the private savings year by year, as they are drawn. Wealth up to 100,000 SEK is disregarded. The home you live in does not count.",
      ),
      say(
        l,
        "Livsvarigt sparande sprids över modellens förväntade återstående livslängd och tar då slut, som i de andra flikarna. Uttagstiden gäller tjänstepensionen och den privata pensionen tillsammans, som i modellen.",
        "Savings drawn for life are spread over the model's expected remaining lifetime and then run out, as in the other tabs. The payout period applies to the occupational and the private pension together, as in the model.",
      ),
      notes.pension,
      notes.savings,
      say(
        l,
        "Det här är räkneexempel i en modell, inte råd. Kontrollera dina egna villkor hos minPension, Pensionsmyndigheten och ditt pensionsbolag innan du väljer.",
        "These are worked examples in a model, not advice. Check your own terms with minPension, Pensionsmyndigheten and your pension company before you choose.",
      ),
    ]) {
      list.append(h("li", undefined, text));
    }
    body.append(list);
    return section;
  }

  function runsFor(shared: ModelContext, shownDeaths: DeathProbabilities): StrategyRuns {
    const key = strategyRunsKey(input);
    if (cache !== undefined && cache.key === key && cache.context === shared && cache.deaths === shownDeaths) return cache.runs;
    const runs = runStrategy(input, shared, shownDeaths);
    cache = { key, context: shared, deaths: shownDeaths, runs };
    return runs;
  }

  function redraw(): void {
    if (context === undefined || deaths === undefined) return;
    const strategy = evaluateStrategy(runsFor(context, deaths), input, context, new Date());
    modelLifespan = strategy.modelExpectedDeath;
    if (input.lifespan === 0) lifespanField.setValue(Math.round(modelLifespan));
    relabelLifespan(currentLang);
    savingsUi.relabel(currentLang);
    const l = currentLang;
    const out: HTMLElement[] = [];
    if (strategy.warnings.length > 0) {
      const box = h("aside", "warnings");
      for (const message of strategy.warnings) box.append(h("p", undefined, message));
      out.push(box);
    }
    out.push(insightsPanel(strategy, l), scenariosPanel(strategy, l), taxPanel(strategy, l), checklistPanel(strategy, l), notesPanel(strategy, l));
    results.replaceChildren(...out);
  }

  // ---- language ---------------------------------------------------------------

  const applyText = (l: Lang): void => {
    inputsTitle.textContent = say(l, "Uttagsstrategi", "Withdrawal strategy");
    intro.textContent = say(
      l,
      "I vilken ordning, när och från vilket konto ska du ta ut pengarna? Fyll i dig själv, dina pensioner, ditt sparande och ditt boende. Fliken räknar fram det de flesta missar, jämför tre sätt att ta ut pengarna och ger en checklista. Fliken har egna indata och påverkas bara av inställningarna under Avancerat.",
      "In which order, when and from which account should you draw the money? Enter yourself, your pensions, your savings and your housing. The tab works out what most people miss, compares three ways of drawing the money and gives a checklist. The tab has its own inputs and is affected only by the settings under Advanced.",
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
