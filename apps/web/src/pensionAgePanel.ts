/**
 * "Hur maxa pensionen?": the tab that answers *which retirement age pays out the
 * most over a lifetime* -- for a typfall of your own, a collective agreement of
 * your own, and a lifespan you can only guess at.
 *
 * `pensionAge.ts` does the arithmetic (its file comment is where the method and
 * its conventions live: "lever till D år", break-even, the two overridden
 * context settings); `pensionAgeChart.ts` draws the three line charts; this file
 * is the panel around them -- the inputs, the short answer, the tables and the
 * exports.
 *
 * Like Mikrosim's rows, this tab's inputs are its own: it is not a diff against
 * the Prognos form and does not follow it. "Hämta från Prognos" copies the
 * form's person, agreement and assumptions across once, at the moment of the
 * click, the same one-time snapshot Mikrosim's import button takes. What it does
 * read live is the shared `ModelContext` (Avancerat: tax, wage profile, price
 * basis, ...), which `main.ts` hands over through `setContext` on every render --
 * `pensionAge.ts` says which of those settings it deliberately ignores.
 *
 * Results are live, in the same way as everywhere else here: a change runs the
 * model once per alternative retirement age (seven runs, a few milliseconds) and
 * rebuilds the results area. The inputs are built once and never rebuilt, so a
 * field you are typing into keeps its focus.
 *
 * **Salary as monthly salary or as final salary.** The model's own input is a
 * monthly salary; the final salary (Table 1's "Slutlön") is what comes out of
 * it. The "Ange lön som" switch lets either be the thing you type. Typed as a
 * final salary, the monthly salary behind it is found by running the model
 * (`monthlySalaryForTarget`) -- the relation is linear, so a couple of secant
 * steps land within a krona -- and both numbers are shown beside the field.
 * Which alternative's final salary is matched is the reference age's: the
 * others can differ a little once wages grow.
 *
 * **Two views.** Below the inputs a switch chooses what the results are: the
 * single-agreement analysis this file builds, or the side-by-side comparison of
 * all the collective agreements (`pensionAgeCompare.ts`, which has its own
 * controls and takes the person and the assumptions from the inputs here). Only
 * the view showing is computed. The inputs, the switch and the comparison's
 * controls are built once; the results area is rebuilt on every change.
 *
 * Every visible string is a plain per-language literal, the same way
 * `compare.ts`'s and `mikrosim.ts`'s own new UI text is -- there is no SysLang
 * row for any of it. The field captions that do have one (birth year, start of
 * work, occupational pension, ...) use it, so they read exactly as in Prognos.
 */
import { BIRTH_YEARS, RETIREMENT_AGES, SCHEME_CHOICES, Scheme, defaultContext, riktalderFor } from "@typfallsmodellen/engine";
import type { DeathProbabilities, ModelContext, SchemeId, TypfallInput } from "@typfallsmodellen/engine";

import type { Bounds, Relabel } from "./controls.js";
import { fieldSet, rangeHint, span } from "./controls.js";
import { csvExportButton, xlsxExportButton } from "./export.js";
import { kronor, percent as share } from "./format.js";
import { t } from "./i18n.js";
import type { Lang } from "./i18n.js";
import {
  LAST_AGE,
  LIFESPAN_BOUNDS,
  NEAR_TIE,
  averageMonthly,
  bestRuns,
  centreBounds,
  computeStudy,
  defaultStudyInput,
  kindsPaid,
  matrixToCsv,
  matrixToXlsxRows,
  monthlySalaryForTarget,
  optimalRows,
  optimalToCsv,
  optimalToXlsxRows,
  referenceAge,
  seriesToCsv,
  seriesToXlsxRows,
} from "./pensionAge.js";
import type { BreakEven, LifespanRow, SeriesMeasure, Study, StudyInput, StudyOption } from "./pensionAge.js";
import { createAgreementCompare } from "./pensionAgeCompare.js";
import { niceTicks, renderLineChart } from "./pensionAgeChart.js";
import { renderOptimalChart } from "./pensionAgeOptimalChart.js";
import {
  capital,
  chip,
  decimals,
  h,
  measureText,
  mkr,
  optionCell,
  optionName,
  panel,
  priceText,
  refName,
  row,
  say,
  signed,
  tableOf,
  tkr,
} from "./pensionAgeUi.js";
import { savingYear, savingsNotes } from "./privateSavings.js";
import { createSavingsInputs } from "./savingsInputs.js";
import { cell, headCell } from "./tables.js";

// The sheet's own ranges; duplicated rather than imported for the same reason
// `compare.ts` and `mikrosim.ts` each duplicate them (neither exports its own).
const BORN: Bounds = { ...span(BIRTH_YEARS), step: 1 };
const START_WORK: Bounds = { min: 15, max: 40, step: 1 };
const RETIREMENT: Bounds = { ...span(RETIREMENT_AGES), step: 1 };
const SALARY: Bounds = { min: 0, max: 1_000_000, step: 100 };
const LIFESPAN: Bounds = { ...LIFESPAN_BOUNDS, step: 1 };

const clampTo = (value: number, bounds: Bounds): number => Math.min(Math.max(value, bounds.min), bounds.max);

/** What the saving year is before the shared context has arrived. */
const DEFAULT_CONTEXT = defaultContext();

/** The two sets of economic assumptions the tab offers a button for. */
const PRESETS = [
  { id: "standard", sv: "Modellens standard", en: "Model default", inflation: 0, growth: 0, ret: 0.017 },
  { id: "growth", sv: "Med löneökning", en: "With wage growth", inflation: 0.02, growth: 0.018, ret: 0.017 },
] as const;

export interface PensionAgeHandle {
  readonly element: HTMLElement;
  relabel(lang: Lang): void;
  /** Stashes the shared context and deaths and redraws the results. */
  setContext(context: ModelContext, deaths: DeathProbabilities): void;
}

// ------------------------------------------------------------------ panel ---

export function createPensionAgePanel(lang: Lang, getForecastInput: () => TypfallInput): PensionAgeHandle {
  let currentLang = lang;
  let input: StudyInput = defaultStudyInput();
  let salaryAs: "monthly" | "final" = "monthly";
  /** The typed final salary, per month; read only while `salaryAs` is "final". */
  let finalTarget = 0;
  let riktalderLocked = true;
  let view: "single" | "compare" = "single";
  /** What the per-age table shows -- and so what its download holds. Kept across recalculations. */
  let timeMeasure: SeriesMeasure = "monthly";
  let context: ModelContext | undefined;
  let deaths: DeathProbabilities | undefined;
  /** What the last run worked out, for the two field hints that quote it. */
  let derived = { finalSalary: 0, monthlySalary: input.monthlySalary };

  const element = h("div", "pension-age-panel");
  element.dataset.role = "pension-age-panel";

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

  // The switch between the two views, and the comparison's own controls.
  const viewBar = h("div", "pa-view");
  const viewLabel = h("span", "pa-pick-label");
  const viewBox = h("div", "panel-toggle pa-toggle");
  viewBox.setAttribute("role", "group");
  const viewButtons = (
    [
      { view: "single", sv: "Ett avtalsområde", en: "One agreement" },
      { view: "compare", sv: "Jämför avtalsområden", en: "Compare agreements" },
    ] as const
  ).map((choice) => {
    const button = h("button", "panel-btn");
    button.type = "button";
    button.dataset.view = choice.view;
    button.addEventListener("click", () => {
      view = choice.view;
      sync();
      redraw();
    });
    viewBox.append(button);
    return { choice, button };
  });
  viewBar.append(viewLabel, viewBox);
  const compare = createAgreementCompare(lang, () => redraw());
  element.append(inputsSection, viewBar, compare.controls, results);

  const importButton = h("button", "export-btn");
  importButton.type = "button";
  importButton.dataset.action = "pa-import-forecast";
  const resetButton = h("button", "export-btn");
  resetButton.type = "button";
  resetButton.dataset.action = "pa-reset";
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

  // The individual occupational pension, the private pension and the private savings are asked for by
  // `savingsInputs.ts`, the module "Hur länge räcker pensionen?" is built with too, so the two tabs ask the same way.
  const savingsUi = createSavingsInputs({
    prefix: "pa",
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
  bornField.element.dataset.setting = "paBorn";
  const relabelBorn = field(bornField.element, (l) => ({
    label: t("birthYear", l),
    hint: say(l, `Riktålder ${riktalderFor(input.born)} år`, `${t("riktalderCheckbox", l)} ${riktalderFor(input.born)}`),
  }));

  const startField = number(input.startWorkAge, START_WORK, (startWorkAge) => update({ startWorkAge }));
  startField.element.dataset.setting = "paStartWork";
  field(startField.element, (l) => ({ label: t("startWorkAge", l), hint: rangeHint(START_WORK) }));

  const centreField = number(input.centre, RETIREMENT, (centre) => update({ centre }));
  centreField.element.dataset.setting = "paCentre";
  const relabelCentre = field(centreField.element, (l) => {
    const alternatives = alternativeRange();
    return {
      label: say(l, "Referensålder", "Reference age"),
      hint: say(l, `Jämför ${alternatives}`, `Compares ${alternatives}`),
    };
  });
  const riktalderBox = check(riktalderLocked, (locked) => {
    riktalderLocked = locked;
    update({});
  });
  riktalderBox.dataset.setting = "paRiktalder";
  field(riktalderBox, (l) => ({ label: t("riktalderCheckbox", l) }), "field field-check");

  const salaryAsSelect = select(
    [
      { value: 0, label: (l) => say(l, "Månadslön", "Monthly salary") },
      { value: 1, label: (l) => say(l, "Slutlön (vid referensåldern)", "Final salary (at the reference age)") },
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
  salaryAsSelect.element.dataset.setting = "paSalaryAs";
  relabels.push(salaryAsSelect.relabel);
  field(salaryAsSelect.element, (l) => ({ label: say(l, "Ange lön som", "Enter salary as") }), "field field-wide");

  const salaryField = number(input.monthlySalary, SALARY, (value) => {
    if (salaryAs === "monthly") update({ monthlySalary: value });
    else {
      finalTarget = value;
      redraw();
    }
  });
  salaryField.element.dataset.setting = "paSalary";
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

  // Pension
  group("Pension", "Pension");
  const schemeSelect = select(
    SCHEME_CHOICES.map((choice) => ({ value: choice.value, label: () => choice.label })),
    input.scheme,
    (scheme) => update({ scheme: scheme as SchemeId }),
  );
  schemeSelect.element.dataset.setting = "paScheme";
  relabels.push(schemeSelect.relabel);
  field(schemeSelect.element, (l) => ({ label: t("occupational", l) }), "field field-wide");

  const includeBox = check(input.includeOccupational, (includeOccupational) => update({ includeOccupational }));
  includeBox.dataset.setting = "paIncludeOccupational";
  const relabelInclude = field(
    includeBox,
    (l) => ({
      label: say(l, "Räkna med avtalets tjänstepension", "Include the agreement's occupational pension"),
      ...(input.scheme === Scheme.None
        ? { hint: say(l, "Välj ett avtalsområde ovan", "Pick an agreement area above") }
        : {}),
    }),
    "field field-check",
  );

  // The individual occupational pension and the private pension, then the group of private savings.
  savingsUi.appendPension();
  savingsUi.appendPrivateSavings();

  // Lifespan
  group("Livslängd", "Lifespan");
  const fromField = number(input.lifespanFrom, LIFESPAN, (lifespanFrom) =>
    update({ lifespanFrom, lifespanTo: Math.max(input.lifespanTo, lifespanFrom) }),
  );
  fromField.element.dataset.setting = "paLifespanFrom";
  field(fromField.element, (l) => ({
    label: say(l, "Lever till, från (år)", "Lifespan from (years)"),
    hint: rangeHint(LIFESPAN),
  }));
  const toField = number(input.lifespanTo, LIFESPAN, (lifespanTo) =>
    update({ lifespanTo, lifespanFrom: Math.min(input.lifespanFrom, lifespanTo) }),
  );
  toField.element.dataset.setting = "paLifespanTo";
  field(toField.element, (l) => ({
    label: say(l, "Lever till, till (år)", "Lifespan to (years)"),
    hint: rangeHint(LIFESPAN),
  }));

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
  inflationField.element.dataset.setting = "paInflation";
  field(inflationField.element, (l) => ({ label: t("inflation", l), hint: "%" }));
  relabels.push(inflationField.relabel);
  const growthField = percent(input.realGrowth, (realGrowth) => update({ realGrowth }));
  growthField.element.dataset.setting = "paGrowth";
  field(growthField.element, (l) => ({ label: t("realGrowth", l), hint: "%" }));
  relabels.push(growthField.relabel);
  const returnField = percent(input.realReturn, (realReturn) => update({ realReturn }));
  returnField.element.dataset.setting = "paReturn";
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

  /** The year capital and saving count from, as the engine will use it -- for the reference age, the alternatives' middle. */
  function savingStart(): number {
    const retirement = referenceAge(input.born, input.centre);
    return savingYear({ born: input.born, startWorkAge: input.startWorkAge, retirementAge: retirement }, context ?? DEFAULT_CONTEXT);
  }

  function alternativeRange(): string {
    const { min } = centreBounds(input.born);
    const middle = referenceAge(input.born, input.centre);
    const first = Math.max(middle - 3, min);
    const last = Math.min(middle + 3, centreBounds(input.born).max);
    return `${first}–${last}`;
  }

  /** The riktålder tracks the birth year while its box is ticked; either way the age is kept inside what the cohort may draw. */
  function normalise(): void {
    const wanted = riktalderLocked ? riktalderFor(input.born) : input.centre;
    input = { ...input, centre: referenceAge(input.born, wanted) };
  }

  /** Writes the state into every control. Setting a value does not fire `change`. */
  function sync(): void {
    bornField.setValue(input.born);
    startField.setValue(input.startWorkAge);
    salaryAsSelect.element.value = salaryAs === "final" ? "1" : "0";
    salaryField.setValue(salaryAs === "final" ? finalTarget : input.monthlySalary);
    schemeSelect.element.value = String(input.scheme);
    includeBox.checked = input.includeOccupational;
    includeBox.disabled = input.scheme === Scheme.None;
    savingsUi.sync(currentLang);
    centreField.setValue(input.centre);
    centreField.element.disabled = riktalderLocked;
    riktalderBox.checked = riktalderLocked;
    fromField.setValue(input.lifespanFrom);
    toField.setValue(input.lifespanTo);
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
    }
    for (const { choice, button } of viewButtons) {
      button.classList.toggle("active", choice.view === view);
      button.setAttribute("aria-pressed", String(choice.view === view));
    }
    compare.controls.hidden = view !== "compare";
    for (const relabel of [relabelBorn, relabelSalary, relabelInclude, relabelCentre]) relabel(currentLang);
  }

  function update(patch: Partial<StudyInput>): void {
    input = { ...input, ...patch };
    normalise();
    sync();
    redraw();
  }

  importButton.addEventListener("click", () => {
    const forecast = getForecastInput();
    salaryAs = "monthly";
    update({
      born: clampTo(forecast.born, BORN),
      startWorkAge: clampTo(forecast.startWorkAge, START_WORK),
      monthlySalary: Math.round(clampTo(forecast.monthlySalary, SALARY)),
      scheme: forecast.scheme,
      yearlyInflation: forecast.yearlyInflation,
      realGrowth: forecast.realGrowth,
      realReturn: forecast.realReturn,
    });
  });
  resetButton.addEventListener("click", () => {
    input = defaultStudyInput();
    salaryAs = "monthly";
    riktalderLocked = true;
    compare.reset();
    normalise();
    sync();
    redraw();
  });

  // ---- results -----------------------------------------------------------------

  /** "83,4–83,8", or "efter 100" when a later start never catches up within the study. */
  const atText = (items: readonly BreakEven[], l: Lang): string => {
    if (items.some((b) => b.at === null)) return say(l, `efter ${LAST_AGE}`, `after ${LAST_AGE}`);
    const values = items.map((b) => b.at as number);
    const low = Math.min(...values);
    const high = Math.max(...values);
    return low === high ? decimals(low, 1, l) : `${decimals(low, 1, l)}–${decimals(high, 1, l)}`;
  };

  function warningsBox(messages: readonly string[]): HTMLElement {
    const box = h("aside", "warnings");
    for (const message of messages) box.append(h("p", undefined, message));
    return box;
  }

  /** One sentence about what a row's lifespan is worth: the winner against the reference and the worst. */
  function stakes(study: Study, row: LifespanRow, l: Lang): string {
    const best = study.options[row.best]!;
    const worstIndex = row.totals.indexOf(Math.min(...row.totals));
    const worst = study.options[worstIndex]!;
    const bestTotal = row.totals[row.best]!;
    const worstTotal = row.totals[worstIndex]!;
    const referenceTotal = row.totals[study.reference]!;
    const more = (amount: number, base: number): string =>
      `${tkr(amount, l)} ${say(l, "tkr", "thousand kr")}${base > 0 ? ` (${share(amount / base, l)})` : ""}`;
    const name = refName(study, l);
    const years = say(l, "år", "yrs");
    if (row.best === study.reference) {
      return say(
        l,
        `dör du vid ${row.lifespan} ${years} ger ${name} (${best.par}) mest – ${more(bestTotal - worstTotal, worstTotal)} mer än det sämsta (${worst.par})`,
        `if you live to ${row.lifespan}, ${name} (${best.par}) pays the most – ${more(bestTotal - worstTotal, worstTotal)} more than the worst (${worst.par})`,
      );
    }
    return say(
      l,
      `dör du vid ${row.lifespan} ${years} ger bästa alternativ (${best.par}) ${more(bestTotal - referenceTotal, referenceTotal)} mer än ${name} och ${more(bestTotal - worstTotal, worstTotal)} mer än det sämsta (${worst.par})`,
      `if you live to ${row.lifespan}, the best (${best.par}) pays ${more(bestTotal - referenceTotal, referenceTotal)} more than ${name} and ${more(bestTotal - worstTotal, worstTotal)} more than the worst (${worst.par})`,
    );
  }

  const monthlyText = (option: StudyOption, lifespan: number, l: Lang): string => {
    const monthly = averageMonthly(option, lifespan);
    return monthly === undefined ? "–" : kronor(monthly, l);
  };

  function answerPanel(study: Study, l: Lang): HTMLElement {
    const { section, body } = panel(say(l, "Svar i korthet", "Short answer"));
    section.dataset.role = "pa-answer";
    section.classList.add("pa-answer");
    const reference = study.options[study.reference]!;
    const name = refName(study, l);

    body.append(
      h(
        "p",
        "pa-lead",
        say(
          l,
          "Vilken pensionsålder som ger mest pension beror på hur länge du lever. Varje ruta gäller en livslängd och visar bästa pensionsålder, vad den ger i genomsnitt per månad, vad den ger sammanlagt och försprånget mot näst bästa alternativ (≈ = under 1 % av summan, praktiskt taget lika).",
          "Which retirement age pays the most depends on how long you live. Each box is for one lifespan and shows the best retirement age, what it pays a month on average, what it pays in total and the lead over the runner-up (≈ = under 1 % of the total, practically equal).",
        ),
      ),
    );

    const ribbonBox = h("div", "pa-ribbon-box");
    ribbonBox.append(h("div", "pa-ribbon-label", say(l, "Lever till (år)", "Lifespan (years)")));
    const ribbon = h("ol", "pa-ribbon");
    ribbon.dataset.role = "pa-ribbon";
    for (const r of study.rows) {
      const best = study.options[r.best]!;
      const second = study.options[r.runnerUp]!;
      const tie = r.marginShare < NEAR_TIE;
      const item = h("li", "pa-ribbon-item");
      item.dataset.lifespan = String(r.lifespan);
      item.dataset.best = String(best.par);
      item.dataset.tie = tie ? "1" : "0";
      item.title = say(
        l,
        `Lever du till ${r.lifespan}: bäst är start vid ${best.par} (${tkr(r.totals[r.best]!, l)} tkr, i snitt ${monthlyText(best, r.lifespan, l)} kr/mån), därefter ${second.par} (${tkr(r.totals[r.runnerUp]!, l)} tkr).`,
        `If you live to ${r.lifespan}: best is ${best.par} (${tkr(r.totals[r.best]!, l)} thousand kr, ${monthlyText(best, r.lifespan, l)} SEK/month on average), then ${second.par} (${tkr(r.totals[r.runnerUp]!, l)} thousand kr).`,
      );
      const monthlyCell = h("span", "pa-ribbon-v", monthlyText(best, r.lifespan, l));
      monthlyCell.dataset.row = "monthly";
      const totalCell = h("span", "pa-ribbon-v", decimals(r.totals[r.best]! / 1e6, 1, l));
      totalCell.dataset.row = "total";
      item.append(
        h("span", "pa-ribbon-d", String(r.lifespan)),
        chip(best),
        monthlyCell,
        totalCell,
        h("span", "pa-ribbon-m", `${tie ? "≈ " : ""}+${tkr(r.margin, l)}`),
      );
      ribbon.append(item);
    }
    ribbonBox.append(
      ribbon,
      h(
        "div",
        "pa-ribbon-key",
        say(
          l,
          "I varje ruta, uppifrån: livslängd · pensionsålder · månadsbelopp (kr, genomsnitt) · livsutbetalning (Mkr) · försprång (tkr)",
          "In each box, from the top: lifespan · retirement age · monthly amount (SEK, average) · lifetime payout (MSEK) · lead (thousand SEK)",
        ),
      ),
    );
    body.append(ribbonBox);

    // What the reference age pays at the start, in Table 1's own terms, and before tax what it is made of.
    const kinds = kindsPaid(study);
    const parts: string[] = [];
    if (!study.input.afterTax) {
      if (kinds.occupational) {
        parts.push(say(l, `tjänstepension ${kronor(reference.table1OccupationalMonthly, l)}`, `occupational pension ${kronor(reference.table1OccupationalMonthly, l)}`));
      }
      if (kinds.individual) {
        parts.push(say(l, `individuell och privat pension ${kronor(reference.table1IndividualMonthly, l)}`, `individual and private pension ${kronor(reference.table1IndividualMonthly, l)}`));
      }
      if (kinds.savings) {
        parts.push(say(l, `privat sparande ${kronor(reference.table1SavingsMonthly, l)}`, `private savings ${kronor(reference.table1SavingsMonthly, l)}`));
      }
    }
    const occupationalPart =
      parts.length === 0 ? "" : say(l, ` (varav ${parts.join(", ")} kr/mån)`, ` (of which ${parts.join(", ")} SEK/month)`);
    body.append(
      h(
        "p",
        "pa-kpi",
        say(
          l,
          `Vid ${reference.par} år (${name.replace(/^the /, "")}): ${kronor(reference.table1Monthly, l)} kr/mån${occupationalPart}, ${share(reference.table1Share, l)} av slutlönen ${kronor(reference.finalSalaryMonthly, l)} kr/mån. ${capital(measureText(study, l))}.`,
          `At ${reference.par} (${name.replace(/^the /, "")}): ${kronor(reference.table1Monthly, l)} SEK/month${occupationalPart}, ${share(reference.table1Share, l)} of the final salary ${kronor(reference.finalSalaryMonthly, l)} SEK/month. ${capital(measureText(study, l))}.`,
        ),
      ),
    );

    const list = h("ul", "pa-points");
    const bullet = (...parts: (string | Node)[]): void => {
      const li = h("li");
      li.append(...parts);
      list.append(li);
    };

    // 1. the winner per lifespan
    const runs: (string | Node)[] = [
      say(l, "Bäst pensionsålder beroende på hur länge du lever – ", "Best retirement age depending on how long you live – "),
    ];
    bestRuns(study).forEach((run, i) => {
      if (i > 0) runs.push(" · ");
      const range = run.from === run.to ? `${run.from}` : `${run.from}–${run.to}`;
      runs.push(h("strong", undefined, `${range} ${say(l, "år", "yrs")}:`));
      runs.push(
        ` ${say(l, "gå vid", "retire at")} ${run.option.par}${run.option.offset === 0 ? ` (${name.replace(/^the /, "")})` : ""}`,
      );
    });
    bullet(...runs);

    // 2. the break-evens against the reference
    const early = study.versusReference.filter((b) => b.later === reference);
    const late = study.versusReference.filter((b) => b.earlier === reference);
    const sentences: string[] = [];
    if (early.length > 0) {
      sentences.push(
        say(
          l,
          `Att gå före ${name} (${reference.par}) ger mest om du dör före ${atText(early, l)} år; därefter har ${name} hunnit ikapp.`,
          `Retiring before ${name} (${reference.par}) pays the most if you die before ${atText(early, l)}; after that ${name} has caught up.`,
        ),
      );
    }
    if (late.length > 0) {
      sentences.push(
        say(
          l,
          `Att vänta längre än ${name} lönar sig först om du lever längre än ${atText(late, l)} år.`,
          `Waiting longer than ${name} only pays if you live past ${atText(late, l)}.`,
        ),
      );
    }
    if (sentences.length > 0) bullet(sentences.join(" "));

    // 3. where the choice hardly matters
    const ties = study.rows.filter((r) => r.marginShare < NEAR_TIE).map((r) => r.lifespan);
    if (ties.length > 0) {
      const first = ties[0]!;
      const last = ties[ties.length - 1]!;
      const contiguous = last - first + 1 === ties.length;
      bullet(
        contiguous
          ? say(
              l,
              `Mellan ${first} och ${last} år skiljer det mindre än 1 % mellan de två bästa alternativen, så ett val nära det bästa kostar nästan ingenting.`,
              `Between ${first} and ${last} the two best alternatives are within 1 % of each other, so a choice close to the best costs next to nothing.`,
            )
          : say(
              l,
              `Vid ${ties.join(", ")} år skiljer det mindre än 1 % mellan de två bästa alternativen.`,
              `At ${ties.join(", ")} the two best alternatives are within 1 % of each other.`,
            ),
      );
    }

    // 4. what is at stake at each end of the range
    const first = study.rows[0];
    const last = study.rows[study.rows.length - 1];
    if (first !== undefined && last !== undefined) {
      const lines = [first, ...(last === first ? [] : [last])].map((r) => stakes(study, r, l));
      const [one, two] = lines;
      bullet(
        say(l, "Vad står på spel: ", "What is at stake: ") +
          (two === undefined ? `${one}.` : `${one}. ${capital(two)}.`),
      );
    }

    body.append(list);
    return section;
  }

  // ---- charts ---------------------------------------------------------------

  function monthlyFigure(study: Study, l: Lang): HTMLElement {
    const last = study.input.lifespanTo;
    const series = study.options.map((o) => ({
      slot: o.slot,
      par: o.par,
      name: optionName(study, o, l),
      end: String(o.par),
      points: Array.from({ length: Math.max(last - o.par + 1, 0) }, (_, i) => [o.par + i, (o.annual[o.par + i] ?? 0) / 12] as const),
      format: (value: number) => `${kronor(value, l)} ${say(l, "kr/mån", "SEK/month")}`,
    }));
    const top = Math.max(1, ...series.flatMap((s) => s.points.map((p) => p[1])));
    const ticks = niceTicks(0, top, 5);
    return renderLineChart({
      role: "pa-fig-monthly",
      lang: l,
      title: say(l, "Pension per månad", "Pension per month"),
      subtitle: `${capital(measureText(study, l))}, ${priceText(study, l)}. ${say(l, "Varje linje startar vid sin pensionsålder.", "Each line starts at its own retirement age.")}`,
      xMin: Math.min(...study.options.map((o) => o.par)),
      xMax: Math.max(last, Math.min(...study.options.map((o) => o.par)) + 1),
      xTitle: say(l, "Ålder", "Age"),
      yTicks: ticks.ticks,
      autoTicks: 5,
      yTick: (v) => kronor(v, l),
      yUnit: say(l, "kr per månad", "SEK per month"),
      heading: (x) => `${say(l, "Ålder", "Age")} ${x}`,
      series,
      startDots: true,
    });
  }

  function totalFigure(study: Study, l: Lang): HTMLElement {
    const last = study.input.lifespanTo;
    const series = study.options.map((o) => ({
      slot: o.slot,
      par: o.par,
      name: optionName(study, o, l),
      end: String(o.par),
      points: Array.from({ length: Math.max(last - o.par + 1, 0) }, (_, i) => [o.par + i, o.cumulative[o.par + i] ?? 0] as const),
      format: (value: number) => `${mkr(value, l)} Mkr`,
    }));
    const top = Math.max(1, ...series.flatMap((s) => s.points.map((p) => p[1])));
    const ticks = niceTicks(0, top, 7);
    return renderLineChart({
      role: "pa-fig-total",
      lang: l,
      title: say(l, "Totalt utbetald pension", "Total pension paid out"),
      subtitle: say(
        l,
        `Summan som betalats ut före respektive ålder, ${measureText(study, l)}, miljoner kronor, ${priceText(study, l)}, utan diskontering.`,
        `The sum paid out before each age, ${measureText(study, l)}, millions of kronor, ${priceText(study, l)}, not discounted.`,
      ),
      xMin: Math.min(...study.options.map((o) => o.par)),
      xMax: Math.max(last, Math.min(...study.options.map((o) => o.par)) + 1),
      xTitle: say(l, "Ålder", "Age"),
      yTicks: ticks.ticks,
      autoTicks: 7,
      yTick: (v) => (v === 0 ? "0" : `${kronor(v / 1e6, l)} Mkr`),
      yUnit: "Mkr",
      heading: (x) => `${say(l, "Utbetalt före ålder", "Paid out before age")} ${x}`,
      series,
      startDots: true,
    });
  }

  function deltaFigure(study: Study, l: Lang): HTMLElement | undefined {
    if (study.rows.length < 2) return undefined;
    const name = refName(study, l);
    const reference = study.options[study.reference]!;
    const series = study.options.map((o, i) => ({
      slot: o.slot,
      par: o.par,
      name: optionName(study, o, l),
      end: String(o.par),
      points: study.rows.map((r) => [r.lifespan, r.totals[i]! - r.totals[study.reference]!] as const),
      format: (value: number) => `${signed(kronor(value / 1000, l), value)} ${say(l, "tkr", "thousand kr")}`,
    }));
    const values = series.flatMap((s) => s.points.map((p) => p[1]));
    const ticks = niceTicks(Math.min(0, ...values), Math.max(0, ...values), 6);
    return renderLineChart({
      role: "pa-fig-delta",
      lang: l,
      title: say(l, `Skillnad mot ${name} (${reference.par} år), per livslängd`, `Difference from ${name} (${reference.par}), by lifespan`),
      subtitle: say(
        l,
        `Total pension minus summan för ${name}, tusen kronor. Över noll = mer än ${name}. Där en linje korsar noll är brytpunkten.`,
        `Total pension minus the total for ${name}, thousands of kronor. Above zero = more than ${name}. Where a line crosses zero is the break-even point.`,
      ),
      xMin: study.input.lifespanFrom,
      xMax: study.input.lifespanTo,
      xTitle: say(l, "Lever till (år)", "Lifespan (years)"),
      yTicks: ticks.ticks,
      autoTicks: 6,
      yTick: (v) => (v === 0 ? "0" : `${v < 0 ? "−" : "+"}${kronor(Math.abs(v) / 1000, l)}`),
      yUnit: say(l, "tkr", "thousand kr"),
      heading: (x) => `${say(l, "Lever till", "Lifespan")} ${x}`,
      series,
      annotations: {
        top: say(l, `▲ mer än ${name}`, `▲ more than ${name}`),
        bottom: say(l, `▼ mindre än ${name}`, `▼ less than ${name}`),
      },
    });
  }

  // ---- tables ---------------------------------------------------------------

  function timePanel(study: Study, l: Lang): HTMLElement {
    // Before tax the first year splits into what each kind pays; after tax the tax is not split out of net income.
    const kinds = kindsPaid(study);
    const split = !study.input.afterTax && (kinds.occupational || kinds.individual || kinds.savings);
    const last = study.input.lifespanTo;
    const firstPar = Math.min(...study.options.map((o) => o.par));
    const ages: number[] = [];
    for (let age = Math.ceil((firstPar + 1) / 5) * 5; age <= last; age += 5) ages.push(age);

    const title = (measure: SeriesMeasure): string =>
      measure === "monthly"
        ? say(l, "Pension per månad vid olika åldrar", "Pension per month at different ages")
        : say(l, "Total utbetalning före olika åldrar", "Total payout before different ages");
    // The downloads are of what the table shows, so they are made again when the switch is pressed.
    const downloads = (measure: SeriesMeasure): HTMLElement[] => {
      const file = measure === "monthly" ? say(l, "pension-per-ar", "pension-per-year") : say(l, "total-utbetalning", "total-payout");
      return [
        csvExportButton(l, `${file}.csv`, () => seriesToCsv(study, l, measure)),
        xlsxExportButton(l, `${file}.xlsx`, () => seriesToXlsxRows(study, l, measure)),
      ];
    };
    const { section, body } = panel(title(timeMeasure), downloads(timeMeasure));
    section.dataset.role = "pa-time";

    const monthlyTable = (): HTMLElement[] => {
      const head = [
        headCell(say(l, "Start vid", "Start at")),
        headCell(say(l, "Första året", "First year")),
        ...(split ? [headCell(say(l, "varav allmän", "of which public"))] : []),
        ...(split && kinds.occupational ? [headCell(say(l, "varav tjänstepension", "of which occupational"))] : []),
        ...(split && kinds.individual ? [headCell(say(l, "varav individuell och privat", "of which individual and private"))] : []),
        ...(split && kinds.savings ? [headCell(say(l, "varav privat sparande", "of which private savings"))] : []),
        ...ages.map((age) => headCell(`${age} ${say(l, "år", "yrs")}`)),
        headCell(say(l, "Slutlön", "Final salary")),
      ];
      const rows = study.options.map((o) => {
        const cells: HTMLTableCellElement[] = [
          optionCell(o, optionName(study, o, l)),
          cell(kronor(o.firstYearMonthly, l), "num"),
        ];
        if (split) {
          const occ = (o.occupational[o.par] ?? 0) / 12;
          const ind = (o.individual[o.par] ?? 0) / 12;
          const sav = (o.savings[o.par] ?? 0) / 12;
          cells.push(cell(kronor(o.firstYearMonthly - occ - ind - sav, l), "num"));
          if (kinds.occupational) cells.push(cell(kronor(occ, l), "num"));
          if (kinds.individual) cells.push(cell(kronor(ind, l), "num"));
          if (kinds.savings) cells.push(cell(kronor(sav, l), "num"));
        }
        for (const age of ages) cells.push(cell(age >= o.par ? kronor((o.annual[age] ?? 0) / 12, l) : "–", "num"));
        cells.push(cell(kronor(o.finalSalaryMonthly, l), "num"));
        return row(cells, { par: String(o.par) });
      });
      return [
        tableOf("pa-time-table", "pa-time", head, rows),
        h(
          "p",
          "figure-note",
          say(
            l,
            "Första året är samma belopp som Prognos visar vid pensionering (Tabell 1). Därefter följer pensionen modellens år-för-år-flöde, som diagrammen och summorna bygger på.",
            "The first year is the same amount Prognos shows at retirement (Table 1). After that the pension follows the model's year-by-year flow, which the charts and totals are built on.",
          ),
        ),
      ];
    };

    const totalTable = (): HTMLElement[] => {
      const head = [headCell(say(l, "Start vid", "Start at")), ...ages.map((age) => headCell(`${age} ${say(l, "år", "yrs")}`))];
      const rows = study.options.map((o) =>
        row(
          [
            optionCell(o, optionName(study, o, l)),
            ...ages.map((age) => cell(age >= o.par ? kronor(o.cumulative[age] ?? 0, l) : "–", "num")),
          ],
          { par: String(o.par) },
        ),
      );
      return [
        tableOf("pa-time-table", "pa-time", head, rows),
        h(
          "p",
          "figure-note",
          say(
            l,
            `Summan som betalats ut före respektive ålder, ${measureText(study, l)}, kronor, ${priceText(study, l)}, utan diskontering.`,
            `The sum paid out before each age, ${measureText(study, l)}, kronor, ${priceText(study, l)}, not discounted.`,
          ),
        ),
      ];
    };

    // The switch sits above the table; pressing it changes the table, the heading and the downloads in place.
    const content = h("div");
    const bar = h("div", "pa-view pa-time-view");
    const box = h("div", "panel-toggle pa-toggle");
    box.setAttribute("role", "group");
    box.setAttribute("aria-label", say(l, "Vad tabellen visar", "What the table shows"));
    bar.append(h("span", "pa-pick-label", say(l, "Visa", "Show")), box);
    const choices = (
      [
        { measure: "monthly", text: say(l, "Månad", "Month") },
        { measure: "total", text: say(l, "Total utbetalning", "Total payout") },
      ] as const
    ).map((choice) => {
      const button = h("button", "panel-btn", choice.text);
      button.type = "button";
      button.dataset.measure = choice.measure;
      button.addEventListener("click", () => {
        timeMeasure = choice.measure;
        paint();
      });
      box.append(button);
      return { choice, button };
    });
    function paint(): void {
      section.dataset.measure = timeMeasure;
      section.querySelector("h2 > span")!.textContent = title(timeMeasure);
      section.querySelector(".panel-actions")!.replaceChildren(...downloads(timeMeasure));
      for (const { choice, button } of choices) {
        button.classList.toggle("active", choice.measure === timeMeasure);
        button.setAttribute("aria-pressed", String(choice.measure === timeMeasure));
      }
      content.replaceChildren(...(timeMeasure === "monthly" ? monthlyTable() : totalTable()));
    }
    paint();
    body.append(bar, content);
    return section;
  }

  function matrixPanel(study: Study, l: Lang): HTMLElement {
    const actions = [
      csvExportButton(l, say(l, "livstidssummor.csv", "lifetime-totals.csv"), () => matrixToCsv(study, l)),
      xlsxExportButton(l, say(l, "livstidssummor.xlsx", "lifetime-totals.xlsx"), () => matrixToXlsxRows(study, l)),
    ];
    const { section, body } = panel(say(l, "Total pension per livslängd", "Total pension by lifespan"), actions);
    section.dataset.role = "pa-matrix-panel";
    body.append(
      h(
        "p",
        "pa-lead",
        say(
          l,
          "Summan som betalats ut från pensionsstart tills du fyller den ålder som anges i första kolumnen (livslängd − pensionsålder hela år), i tusen kronor. Markerad ruta = högst summa för den livslängden.",
          "The sum paid out from the start until you reach the age in the first column (lifespan − retirement age whole years), in thousands of kronor. Highlighted box = the highest total for that lifespan.",
        ),
      ),
    );

    const head = [
      headCell(say(l, "Lever till (år)", "Lifespan (years)")),
      ...study.options.map((o) => {
        const th = headCell("");
        th.append(
          h("span", `pa-key pa-key-${o.slot}`),
          document.createTextNode(String(o.par)),
          h("small", undefined, o.offset === 0 ? refName(study, l).replace(/^the /, "") : `${o.offset > 0 ? "+" : "−"}${Math.abs(o.offset)}`),
        );
        return th;
      }),
      headCell(say(l, "Bäst", "Best")),
      headCell(say(l, "Försprång mot näst bäst, tkr", "Lead over runner-up, thousand kr")),
    ];
    const rows = study.rows.map((r) => {
      const best = study.options[r.best]!;
      const tie = r.marginShare < NEAR_TIE;
      const cells: HTMLTableCellElement[] = [cell(String(r.lifespan))];
      r.totals.forEach((total, i) => {
        const td = cell(tkr(total, l), i === r.best ? "num pa-best" : "num");
        td.dataset.par = String(study.options[i]!.par);
        td.dataset.kr = String(Math.round(total));
        if (i === r.best) td.classList.add(`pa-best-${best.slot}`);
        cells.push(td);
      });
      const bestCell = cell("", "num");
      bestCell.append(chip(best));
      const lead = cell("", "num");
      lead.append(document.createTextNode(`${tie ? "≈ " : ""}${tkr(r.margin, l)}`), h("small", undefined, share(r.marginShare, l)));
      cells.push(bestCell, lead);
      return row(cells, { lifespan: String(r.lifespan), best: String(best.par) });
    });
    body.append(tableOf("pa-matrix", "pa-matrix", head, rows));
    return section;
  }

  function optimalPanel(study: Study, l: Lang): HTMLElement {
    const actions = [
      csvExportButton(l, say(l, "optimal-pensionsalder.csv", "optimal-retirement-age.csv"), () => optimalToCsv(study, l)),
      xlsxExportButton(l, say(l, "optimal-pensionsalder.xlsx", "optimal-retirement-age.xlsx"), () => optimalToXlsxRows(study, l)),
    ];
    const { section, body } = panel(
      say(l, "Optimal pensionsålder i förhållande till livslängd", "Optimal retirement age in relation to lifespan"),
      actions,
    );
    section.dataset.role = "pa-optimal-panel";
    const points = optimalRows(study);
    const text = {
      age: say(l, "Pensionsålder", "Retirement age"),
      lifespan: say(l, "Livslängd", "Lifespan"),
      monthly: say(l, "Månadsbelopp kr", "Monthly amount SEK"),
      total: say(l, "Livsutbetalning mkr", "Lifetime payout MSEK"),
    };
    body.append(
      h(
        "p",
        "pa-lead",
        say(
          l,
          "För varje livslängd: den pensionsålder som ger mest pension sammanlagt, vad den ger i genomsnitt per månad och vad den ger totalt, i miljoner kronor.",
          "For each lifespan: the retirement age that pays the most in all, what it pays a month on average and what it pays in total, in millions of kronor.",
        ),
      ),
    );

    const head = [headCell(text.lifespan), ...points.map((p) => headCell(String(p.lifespan)))];
    const ageRow = row(
      [
        cell(text.age, "pa-opt-label"),
        ...points.map((p, i) => {
          const td = cell("", "num");
          td.dataset.lifespan = String(p.lifespan);
          td.append(chip(study.options[study.rows[i]!.best]!));
          return td;
        }),
      ],
      { row: "age" },
    );
    const figureRow = (key: "monthly" | "total", show: (p: (typeof points)[number]) => string): HTMLTableRowElement =>
      row(
        [
          cell(text[key], "pa-opt-label"),
          ...points.map((p) => {
            const td = cell(show(p), "num");
            td.dataset.lifespan = String(p.lifespan);
            return td;
          }),
        ],
        { row: key },
      );
    body.append(
      tableOf("pa-optimal", "pa-optimal", head, [
        ageRow,
        figureRow("monthly", (p) => (p.monthly === undefined ? "–" : kronor(p.monthly, l))),
        figureRow("total", (p) => decimals(p.total / 1e6, 1, l)),
      ]),
    );

    body.append(
      renderOptimalChart({
        role: "pa-fig-optimal",
        lang: l,
        title: say(l, "Pension per månad och totalt vid bästa pensionsålder", "Pension per month and in total at the best retirement age"),
        subtitle: say(
          l,
          `Bästa pensionsålder för varje livslängd. Stapel: vad den ger i genomsnitt per månad (kr, vänster axel). Linje: vad den ger sammanlagt (miljoner kronor, höger axel). ${capital(measureText(study, l))}, ${priceText(study, l)}, utan diskontering.`,
          `The best retirement age for each lifespan. Bar: what it pays a month on average (SEK, left axis). Line: what it pays in all (millions of kronor, right axis). ${capital(measureText(study, l))}, ${priceText(study, l)}, not discounted.`,
        ),
        points,
        text: {
          ...text,
          leftUnit: say(l, "Kr/mån", "SEK/month"),
          rightUnit: "Mkr",
          heading: (lifespan) => say(l, `Lever du till ${lifespan} år`, `If you live to ${lifespan}`),
        },
        format: {
          kronor: (v) => kronor(v, l),
          monthly: (v) => `${kronor(v, l)} ${say(l, "kr/mån", "SEK/month")}`,
          millions: (v) => decimals(v, 1, l),
        },
      }),
    );
    return section;
  }

  function breakEvenPanel(study: Study, l: Lang): HTMLElement {
    const { section, body } = panel(say(l, "Brytpunkter", "Break-even points"));
    section.dataset.role = "pa-breakeven-panel";
    const name = refName(study, l);
    const reference = study.options[study.reference]!;
    const at = (b: BreakEven): string => (b.at === null ? say(l, `efter ${LAST_AGE}`, `after ${LAST_AGE}`) : decimals(b.at, 1, l));

    body.append(
      h(
        "p",
        "pa-lead",
        say(
          l,
          `Brytpunkten är den livslängd då det senare alternativet hunnit ikapp det tidigare, så att summorna är lika. Under den ger det tidigare mest, över den det senare.`,
          `The break-even point is the lifespan at which the later alternative has caught up with the earlier one, so that the totals are equal. Below it the earlier pays more, above it the later one.`,
        ),
      ),
    );

    const versus = study.versusReference.map((b) => {
      const other = b.earlier === reference ? b.later : b.earlier;
      const text =
        b.earlier === reference
          ? say(
              l,
              `Start vid ${other.par} ger mer än ${name} om du lever längre än ${at(b)} år.`,
              `Retiring at ${other.par} pays more than ${name} if you live past ${at(b)}.`,
            )
          : say(
              l,
              `Start vid ${other.par} ger mest om du dör före ${at(b)} år; lever du längre ger ${name} mer.`,
              `Retiring at ${other.par} pays the most if you die before ${at(b)}; if you live longer ${name} pays more.`,
            );
      const explain = cell(text, "txt");
      return row([optionCell(other, optionName(study, other, l)), cell(at(b), "num"), explain], { par: String(other.par) });
    });
    body.append(
      tableOf(
        "pa-breakeven",
        "pa-breakeven",
        [headCell(say(l, "Alternativ", "Alternative")), headCell(say(l, "Lika vid (år)", "Equal at (years)")), headCell(say(l, "Så läser du det", "How to read it"))],
        versus,
      ),
    );

    const chain = study.chain.map((b) => row([cell(`${b.earlier.par} → ${b.later.par} ${say(l, "år", "yrs")}`), cell(at(b), "num")], { to: String(b.later.par) }));
    const chainTable = tableOf(
      "pa-chain",
      "pa-chain",
      [headCell(say(l, "Från → till", "From → to")), headCell(say(l, "Lönar sig efter (år)", "Pays off after (years)"))],
      chain,
    );
    chainTable.classList.add("pa-narrow");
    body.append(h("p", "pa-lead", say(l, "Steg för steg: att vänta ett år till lönar sig om du lever längre än …", "Step by step: waiting one more year pays off if you live longer than …")), chainTable);
    return section;
  }

  function notesPanel(study: Study, l: Lang): HTMLElement {
    const details = h("details", "field-note pa-notes");
    details.dataset.role = "pa-notes";
    details.append(h("summary", undefined, say(l, "Antaganden och metod", "Assumptions and method")));
    const bodyBox = h("div", "field-note-body");
    const savings = savingsNotes(l, savingStart());
    const paragraphs = [
      say(
        l,
        `Pension är allmän pension (inkomstpension, premiepension, garantipension och inkomstpensionstillägg) plus tjänstepension om du räknar med den, och eventuell individuell tjänstepension, privat pension och privat sparande, från pensionsstart. Efter skatt är modellens nettoinkomst; privat sparande beskattas inte vid utbetalning och räknas lika före och efter skatt. Belopp i ${priceText(study, l)}, utan diskontering: en krona vid 90 års ålder räknas som en krona vid 70.`,
        `Pension is public pension (income pension, premium pension, guarantee pension and the income-pension supplement) plus occupational pension if you include it, and any individual occupational pension, private pension and private savings, from the start. After tax is the model's net income; private savings are not taxed when paid out and count the same before and after tax. Amounts are in ${priceText(study, l)}, not discounted: a krona at 90 counts as much as one at 70.`,
      ),
      say(
        l,
        `”Lever till D år” betyder att pension betalas för alla år från pensionsstart till och med året före D-årsdagen (D − pensionsålder hela år). Pensionens storlek beror inte på hur länge du lever; livslängden avgör bara hur många år som summeras. Ingen sannolikhetsviktning görs.`,
        `"Lifespan D" means pension is paid for every year from the start up to the year before the D-year birthday (D − retirement age whole years). The size of the pension does not depend on how long you live; the lifespan only decides how many years are added up. No probability weighting is applied.`,
      ),
      say(
        l,
        `Tjänstepensionen tas ut samtidigt som den allmänna pensionen i den här fliken, oavsett inställningen under Avancerat för uttagsålder. Avtalsområdet avgör beloppet; modellen varnar inte om det inte hör ihop med födelseåret (ITP-1 gäller födda 1979 och senare, AKAP-KR födda 1986 och senare, PA16 avd 1 födda 1988 och senare).`,
        `Occupational pension is drawn together with the public pension in this tab, whatever the Advanced setting for its start age says. The agreement area decides the amount; the model does not warn if it does not fit the birth year (ITP-1 is for those born 1979 or later, AKAP-KR 1986 or later, PA16 avd 1 1988 or later).`,
      ),
      // The two paragraphs about the savings are the other tab's too (`savingsNotes`); this tab adds what is its own.
      `${savings.pension} ${say(
        l,
        "Sparandet fortsätter fram till respektive pensionsålder, så ett senare val ger också fler år av sparande och avkastning. Allt betalas ut under samma tid som avtalets tjänstepension (Avancerat > Temporärt uttag; annars livsvarigt).",
        "The saving goes on until each retirement age, so a later choice also means more years of saving and of return. Everything is paid out over the same time as the agreement's occupational pension (Advanced > Temporary withdrawal; otherwise for life).",
      )}`,
      `${savings.savings} ${say(
        l,
        "Privat sparande har ingen utbetalning för livet: pengarna fördelas över din förväntade återstående livslängd (modellens regel) och är sedan slut, medan en pension betalas ut så länge du lever.",
        "Private savings have no payout for life: the money is spread over your expected remaining life (the model's rule) and is then gone, where a pension is paid for as long as you live.",
      )}`,
      say(
        l,
        `Standardantagandena (0 % inflation, 0 % real löneökning) ger i fast penningvärde en inkomstpension som sjunker ungefär 1,6 % per år, och garantipensionen, som följer priserna, fyller på allt mer i hög ålder. Det skjuter brytpunkterna flera år framåt jämfört med antagandet om en växande lön – pröva knappen "Med löneökning" för att se hur mycket svaret beror på det.`,
        `The default assumptions (0 % inflation, 0 % real wage growth) make the income pension fall by about 1.6 % a year in constant prices, and the guarantee pension, which follows prices, tops it up more and more in old age. That pushes the break-even points several years later than if wages grow – try "With wage growth" to see how much the answer depends on it.`,
      ),
      say(
        l,
        `Fliken har egna indata och följer inte formuläret till vänster; "Hämta från Prognos" kopierar dem en gång (sparandet följer inte med). Skatteinställningar, löneprofil och prisbas kommer från Avancerat. Kända kapitalvärden som anges under Avancerat gäller också här.`,
        `The tab has its own inputs and does not follow the form on the left; "Use the Forecast values" copies them once (the savings do not come along). Tax settings, wage profile and price basis come from Advanced. Known capital values entered under Advanced apply here too.`,
      ),
    ];
    for (const text of paragraphs) bodyBox.append(h("p", undefined, text));
    details.append(bodyBox);
    return details;
  }

  function buildResults(study: Study, l: Lang): HTMLElement[] {
    const out: HTMLElement[] = [];
    if (study.warnings.length > 0) out.push(warningsBox(study.warnings));
    out.push(
      answerPanel(study, l),
      monthlyFigure(study, l),
      timePanel(study, l),
      totalFigure(study, l),
      matrixPanel(study, l),
      optimalPanel(study, l),
    );
    const delta = deltaFigure(study, l);
    if (delta) out.push(delta);
    out.push(breakEvenPanel(study, l), notesPanel(study, l));
    return out;
  }

  function redraw(): void {
    // The year the capital counts from is the shared context's, which may have changed.
    savingsUi.relabel(currentLang);
    if (context === undefined || deaths === undefined) return;
    let effective = input;
    if (salaryAs === "final") {
      effective = { ...input, monthlySalary: monthlySalaryForTarget(input, finalTarget, context, deaths) };
    }
    if (view === "compare") {
      const shown = compare.render({ input: effective, chosen: input.scheme, context, deaths });
      derived = shown.derived;
      relabelSalary(currentLang);
      results.replaceChildren(...shown.nodes);
      return;
    }
    const study = computeStudy(effective, context, deaths);
    derived = {
      finalSalary: study.options[study.reference]?.finalSalaryMonthly ?? 0,
      monthlySalary: effective.monthlySalary,
    };
    relabelSalary(currentLang);
    results.replaceChildren(...buildResults(study, currentLang));
  }

  // ---- language ---------------------------------------------------------------

  const applyText = (l: Lang): void => {
    inputsTitle.textContent = say(l, "Hur maxa pensionen?", "How to maximise the pension?");
    intro.textContent = say(
      l,
      "Jämför sju pensionsåldrar kring en referensålder och se vilken som ger mest pension totalt, beroende på hur länge du lever. Ändra personen, avtalsområdet, sparandet eller antagandena så räknas allt om direkt. Fliken har egna indata och påverkas bara av inställningarna under Avancerat.",
      "Compares seven retirement ages around a reference age and shows which pays the most in total, depending on how long you live. Change the person, the agreement area, the savings or the assumptions and everything is recalculated at once. The tab has its own inputs and is affected only by the settings under Advanced.",
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
    viewLabel.textContent = say(l, "Resultat", "Results");
    viewBox.setAttribute("aria-label", say(l, "Vilket resultat som visas", "Which results are shown"));
    for (const { choice, button } of viewButtons) button.textContent = say(l, choice.sv, choice.en);
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
      compare.relabel(l);
      sync();
      redraw();
    },
    setContext(nextContext, nextDeaths) {
      context = nextContext;
      deaths = nextDeaths;
      redraw();
    },
  };
}
