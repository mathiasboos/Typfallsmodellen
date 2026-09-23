/**
 * The page: inputs on the left, results on the right, recomputed on every change.
 *
 * One model run is about ninety ages of arithmetic, well under a millisecond, so
 * there is no debounce and no incremental update -- every change re-runs the
 * model and rebuilds the results. That keeps the state in one place: whatever is
 * on screen is what `run()` returned for what the form says.
 *
 * The results follow the Start sheet's own order: Table 1, Figur 1, Figur 2,
 * the disposable income chart, the tax-per-year chart, Table 2. The last of
 * those is not a workbook figure -- see `renderTaxChart`'s own comment.
 *
 * The workbook has two modes, chosen by a pair of radio circles on the Start
 * sheet, and this keeps them: normal mode runs on a `TypfallInput` alone, which
 * is the whole reason `viewContext` could get this far handing `run()` an empty
 * settings map. Avancerat adds `advanced.ts`'s settings and `salaryPath.ts`'s
 * own salary path on top.
 */
import { contextFromSettings, defaultInput, run } from "@typfallsmodellen/engine";
import type { ModelContext, TypfallInput, TypfallResult } from "@typfallsmodellen/engine";

import { createAdvancedPanel } from "./advanced.js";
import { renderDisposable, renderFigure1, renderFigure2, renderTaxChart } from "./chart.js";
import type { FigureView } from "./chart.js";
import { createComparePanel } from "./compare.js";
import { loadDeathProbabilities } from "./deaths.js";
import { csvExportButton, xlsxExportButton } from "./export.js";
import { createForm } from "./form.js";
import { LANGS, dropHeadingNumber, t } from "./i18n.js";
import type { Lang, LabelName } from "./i18n.js";
import { renderKpis, retirementAge } from "./kpis.js";
import { createPgbGrid } from "./pgb.js";
import { createSalaryPath } from "./salaryPath.js";
import {
  renderTable1,
  renderTable2,
  table1ToCsv,
  table1ToXlsxRows,
  table2ToCsv,
  table2ToXlsxRows,
} from "./tables.js";
import type { Table1View } from "./tables.js";
import "./styles.css";

interface View {
  readonly lang: Lang;
  /** The workbook's own month/year switch. */
  readonly monthly: boolean;
}

/** The workbook's two radio circles: `Normalt` and `Avancerat`. */
type Mode = "normal" | "advanced";

/** Not a workbook mode -- a second top-level view alongside the forecast. */
type Screen = "single" | "compare";

const deaths = loadDeathProbabilities();

/**
 * The context for a view.
 *
 * `rng_Chart_Earning_factor` is the workbook's month/year switch: `buildTable2`
 * divides by it, so asking for monthly amounts is a model setting rather than a
 * division in the view. Table 1 carries all four of the sheet's columns on every
 * row, so it needs nothing here.
 */
function viewContext(view: View): ModelContext {
  return contextFromSettings(new Map(), {
    ...(mode === "advanced" ? advanced : {}),
    chartEarningFactor: view.monthly ? 12 : 1,
  });
}

/**
 * The typfall as the run sees it.
 *
 * Normal mode suspends the advanced entries rather than dropping them -- the
 * workbook leaves them sitting on its sheet too, and switching modes to compare
 * would be useless if the second switch came back to an empty form. Only
 * `Använd normala inställningar` clears them.
 */
function runInput(): TypfallInput {
  return mode === "advanced" ? { ...input, ...advancedInput } : input;
}

let input: TypfallInput = defaultInput();
let view: View = { lang: "sv", monthly: true };
let mode: Mode = "normal";
let screen: Screen = "single";
/** Adv_settings, as overrides on top of the workbook's own normal values. */
let advanced: Partial<ModelContext> = {};
/** The Start-sheet side of advanced mode: the own salary path and PGB. */
let advancedInput: Partial<TypfallInput> = {};

const root = document.querySelector("#app");
if (!(root instanceof HTMLElement)) throw new Error("#app is missing from the page");

const header = document.createElement("header");
header.className = "masthead";

const results = document.createElement("div");
results.className = "results";

const form = createForm(input, view.lang, (patch) => {
  input = Object.freeze({ ...input, ...patch });
  render();
});

const advancedPanel = createAdvancedPanel(view.lang, (patch) => {
  advanced = { ...advanced, ...patch };
  render();
});

const salaryPath = createSalaryPath(view.lang, (path) => {
  const next = { ...advancedInput };
  if (path === undefined) delete next.ownIncome;
  else next.ownIncome = path;
  advancedInput = next;
  render();
});

const pgbGrid = createPgbGrid(view.lang, (patch) => {
  const next = { ...advancedInput };
  if (patch.pgbManual === undefined) delete next.pgbManual;
  else next.pgbManual = patch.pgbManual;
  if (patch.pgbConscription === undefined) delete next.pgbConscription;
  else next.pgbConscription = patch.pgbConscription;
  advancedInput = next;
  // `childBirthYears` is a ModelContext (Adv_settings) field, unlike the two
  // above (TypfallInput, the Start sheet) -- it rides in the same patch
  // since it lives in the same panel, but goes into `advanced` instead.
  advanced = { ...advanced, childBirthYears: patch.childBirthYears };
  render();
});

const comparePanel = createComparePanel(view.lang, input, () => render());

/** Everything advanced mode adds, hidden until the mode is switched. */
const advancedBox = document.createElement("div");
advancedBox.className = "advanced-box";
advancedBox.hidden = true;
advancedBox.append(advancedPanel.element, salaryPath.element, pgbGrid.element);

/**
 * `Använd normala inställningar` -- the button the Adv_settings sheet carries.
 *
 * Clears both halves of advanced mode and puts every control back to the
 * workbook's own value, rather than merely suspending them the way switching
 * back to Normalt does.
 */
const resetAdvanced = document.createElement("button");
resetAdvanced.type = "button";
resetAdvanced.className = "export-btn";
resetAdvanced.dataset.action = "reset-advanced";
resetAdvanced.addEventListener("click", () => {
  advanced = {};
  advancedInput = {};
  advancedPanel.reset();
  salaryPath.reset();
  pgbGrid.reset();
  render();
});
advancedBox.append(resetAdvanced);

function modeToggle(): HTMLElement {
  const box = document.createElement("div");
  box.className = "panel-toggle mode-toggle";
  box.setAttribute("role", "group");
  const choices: readonly { mode: Mode; label: (l: Lang) => string }[] = [
    { mode: "normal", label: (l) => (l === "sv" ? "Normalt" : "Normal") },
    { mode: "advanced", label: (l) => (l === "sv" ? "Avancerat" : "Advanced") },
  ];
  for (const choice of choices) {
    const button = document.createElement("button");
    button.type = "button";
    button.dataset.mode = choice.mode;
    button.textContent = choice.label(view.lang);
    button.className = choice.mode === mode ? "panel-btn active" : "panel-btn";
    button.addEventListener("click", () => {
      mode = choice.mode;
      render();
    });
    box.append(button);
  }
  return box;
}

/** Prognos / Jämför scenarier -- not a workbook toggle, so plain per-language
 * literals, the same way `modeToggle`'s own labels are. */
function screenToggle(): HTMLElement {
  const box = document.createElement("div");
  box.className = "panel-toggle screen-toggle";
  box.setAttribute("role", "group");
  const choices: readonly { screen: Screen; label: (l: Lang) => string }[] = [
    { screen: "single", label: (l) => (l === "sv" ? "Prognos" : "Forecast") },
    { screen: "compare", label: (l) => (l === "sv" ? "Jämför scenarier" : "Compare scenarios") },
  ];
  for (const choice of choices) {
    const button = document.createElement("button");
    button.type = "button";
    button.dataset.screen = choice.screen;
    button.textContent = choice.label(view.lang);
    button.className = choice.screen === screen ? "panel-btn active" : "panel-btn";
    button.addEventListener("click", () => {
      screen = choice.screen;
      render();
    });
    box.append(button);
  }
  return box;
}

const modeBox = document.createElement("div");
modeBox.className = "mode-row";

// Built once, not inside `renderHeading()` -- that function reruns on every
// `render()`, which is every input change anywhere on the page, and a fresh
// `document.createElement("details")` each time would reset `.open` back to
// closed the instant someone who had expanded it touched anything else.
// Moving the same node around the DOM (what `header.replaceChildren` below
// still does every render) does not reset it; only recreating the element
// would. Collapsed by default -- a freshly created `<details>` starts closed,
// and nothing here ever sets `.open`.
const noticeSummary = document.createElement("summary");
const noticeStrong = document.createElement("strong");
const noticeRest = document.createElement("span");
const noticeMail = document.createElement("a");
noticeMail.href = "mailto:typfallsmodellen@pensionsmyndigheten.se";
noticeMail.textContent = "typfallsmodellen@pensionsmyndigheten.se";
const notice = document.createElement("details");
notice.className = "disclaimer";
notice.dataset.role = "disclaimer";
notice.append(noticeSummary, noticeStrong, noticeRest, noticeMail, document.createTextNode("."));
applyNoticeText(view.lang);

/** Rebuilt on a language change, so the subtitle and the active chip follow. */
function renderHeading(): void {
  const title = document.createElement("h1");
  title.textContent = "Pensionsprognos";

  const sub = document.createElement("p");
  sub.className = "subtitle";
  sub.textContent =
    view.lang === "sv"
      ? `Alla beräkningar sker i din webbläsare; ingenting skickas någonstans.`
      : `Everything is computed in your browser; nothing is sent anywhere.`;

  const langs = document.createElement("div");
  langs.className = "langs";
  langs.setAttribute("role", "group");
  langs.setAttribute("aria-label", t("chooseLanguage", view.lang));
  for (const lang of LANGS) {
    const button = document.createElement("button");
    button.type = "button";
    button.textContent = lang.toUpperCase();
    button.className = lang === view.lang ? "chip active" : "chip";
    button.addEventListener("click", () => {
      view = { ...view, lang };
      applyNoticeText(lang);
      form.relabel(lang);
      advancedPanel.relabel(lang);
      salaryPath.relabel(lang);
      pgbGrid.relabel(lang);
      comparePanel.relabel(lang);
      document.documentElement.lang = lang;
      render();
    });
    langs.append(button);
  }

  header.replaceChildren(title, sub, notice, langs);
}

/**
 * That this is not Pensionsmyndigheten's own tool, on the page rather than only
 * in the README.
 *
 * The site computes a pension forecast and looks like it knows what it is
 * talking about, which is exactly why it has to say whose model it is and what
 * a forecast is worth. The agency's own address is here because a question
 * about the model belongs with the people who wrote it, not with this port.
 *
 * A `<details>`, matching Ordlista, rather than an always-open box: the text
 * doesn't change while someone works, only the language does, so this just
 * fills in the four text nodes `notice` was built from (see where it's
 * constructed, above) rather than rebuilding the element.
 */
function applyNoticeText(l: Lang): void {
  noticeSummary.textContent = l === "sv" ? "Om modellen" : "About the model";
  noticeStrong.textContent = l === "sv" ? "Inofficiell version." : "Unofficial version.";
  noticeRest.textContent =
    l === "sv"
      ? " Den här sidan är inte utvecklad av, kopplad till eller godkänd av " +
        "Pensionsmyndigheten. Modellen, dess data och dess användarmanual är deras. " +
        "Resultatet är en prognos under de antaganden du anger – inte ett besked om din " +
        "pension. Frågor om själva modellen går till "
      : " This page is not built by, affiliated with or endorsed by Pensionsmyndigheten, " +
        "the Swedish Pensions Agency. The model, its data and its user manual are theirs. " +
        "What it shows is a forecast under the assumptions you enter – not a statement " +
        "about your pension. Questions about the model itself go to ";
}

/**
 * Årsvis / Månadsvis -- Table 2 and the three figures' month/year scale.
 *
 * Moved into Table 2's own header on request, in place of a standalone toggle
 * that used to sit above Table 1 and no longer relates to anything there
 * (Table 1 shows all four of the sheet's columns at once, its own scale).
 */
function scaleToggle(): HTMLElement {
  const box = document.createElement("div");
  box.className = "panel-toggle";
  box.setAttribute("role", "group");
  box.setAttribute("aria-label", t("show", view.lang));

  const choices: readonly { monthly: boolean; label: LabelName }[] = [
    { monthly: false, label: "yearlyView" },
    { monthly: true, label: "monthlyView" },
  ];
  for (const choice of choices) {
    const button = document.createElement("button");
    button.type = "button";
    button.textContent = t(choice.label, view.lang);
    button.className = choice.monthly === view.monthly ? "panel-btn active" : "panel-btn";
    button.addEventListener("click", () => {
      view = { ...view, monthly: choice.monthly };
      render();
    });
    box.append(button);
  }
  return box;
}

function warnings(result: TypfallResult): HTMLElement | undefined {
  if (result.warnings.length === 0) return undefined;
  // The workbook raises these as dialogs mid-run. They mean it changed the
  // question before answering it, which the reader has to be told.
  const box = document.createElement("aside");
  box.className = "warnings";
  for (const warning of result.warnings) {
    const line = document.createElement("p");
    line.textContent = warning.message;
    box.append(line);
  }
  return box;
}

function section(title: string, body: HTMLElement, actions?: HTMLElement): HTMLElement {
  const wrap = document.createElement("section");
  wrap.className = "panel";
  const head = document.createElement("h2");
  const heading = document.createElement("span");
  heading.textContent = title;
  head.append(heading);
  if (actions) head.append(actions);
  wrap.append(head, body);
  return wrap;
}

function render(): void {
  renderHeading();
  const lang = view.lang;

  modeBox.replaceChildren(modeToggle());
  advancedBox.hidden = mode !== "advanced";
  resetAdvanced.textContent =
    lang === "sv" ? "Använd normala inställningar" : "Use the normal settings";

  const context = viewContext(view);
  const typfall = runInput();
  const result = run(typfall, context, { deaths });
  const par = retirementAge(typfall, result);

  // The salary grid fills from, and resets to, the path the model derives for
  // the Start sheet as it currently stands. Once an own path is in use the
  // run's own `wagePath` just echoes it back, so the baseline has to come from
  // a run without it. A run is well under a millisecond, and this second one
  // only happens in advanced mode.
  if (mode === "advanced") {
    const { ownIncome, ...withoutOwnIncome } = typfall;
    const baseline = ownIncome === undefined ? result : run(withoutOwnIncome, context, { deaths });
    salaryPath.setBaseline(baseline.wagePath, input.born);
    pgbGrid.setBaseline(input.born, context.marginal, result.pgbBreakdown);
  }

  results.replaceChildren();
  results.append(screenToggle());

  if (screen === "compare") {
    results.append(comparePanel.element);
    comparePanel.renderResults(typfall, context, deaths, lang, view.monthly);
    return;
  }

  const figures: FigureView = {
    lang,
    par,
    perMonth: view.monthly,
    priceBasis: context.priceBasis,
  };

  // The sheet's own heading, minus its "Tabell 2." numbering (dropped on
  // request), with this run's start age written into what's left. The age is
  // the label's *last* number -- "Månadsinkomster från 56 ålder" -- so the
  // table's own number is left alone.
  const table2Title = dropHeadingNumber(t("table2", lang)).replace(
    /\d+(?=\D*$)/,
    String(result.table2[0]?.age ?? Math.trunc(par) - 10),
  );

  // Replaces SysLang row 24 ("Tabell 1. Specificerat resultat över slutlön och
  // pensionsinkomster") outright, on request, rather than trimming it the way
  // Table 2 and the figures trim theirs.
  const table1Title = lang === "sv" ? "Pensionsinkomst" : "Pension income";

  const warned = warnings(result);
  if (warned) results.append(warned);
  results.append(renderKpis(result, lang, par));
  const table1View: Table1View = {
    par,
    finalSalaryYears: context.finalSalaryYears,
    lastPensionRight: context.lastPensionRight > 0,
  };

  const table1Actions = document.createElement("div");
  table1Actions.className = "panel-actions";
  table1Actions.append(
    csvExportButton(lang, lang === "sv" ? "tabell1.csv" : "table1.csv", () =>
      table1ToCsv(result, lang, table1View),
    ),
    xlsxExportButton(lang, lang === "sv" ? "tabell1.xlsx" : "table1.xlsx", () =>
      table1ToXlsxRows(result, lang, table1View),
    ),
  );

  const table2Actions = document.createElement("div");
  table2Actions.className = "panel-actions";
  table2Actions.append(
    scaleToggle(),
    csvExportButton(lang, lang === "sv" ? "tabell2.csv" : "table2.csv", () =>
      table2ToCsv(result, lang),
    ),
    xlsxExportButton(lang, lang === "sv" ? "tabell2.xlsx" : "table2.xlsx", () =>
      table2ToXlsxRows(result, lang),
    ),
  );

  results.append(
    section(table1Title, wrapScroll(renderTable1(result, lang, table1View)), table1Actions),
    renderFigure1(result, figures),
    renderFigure2(result, figures),
    renderDisposable(result, figures),
    renderTaxChart(result, figures),
    section(table2Title, wrapScroll(renderTable2(result, lang)), table2Actions),
  );

  const foot = document.createElement("p");
  foot.className = "footnote";
  foot.textContent = t("kpiFootnote", lang);
  results.append(foot);
}

/** Both tables are wider than a phone; give each its own scroller rather than
 * letting it widen the page. */
function wrapScroll(table: HTMLElement): HTMLElement {
  const box = document.createElement("div");
  box.className = "scroll";
  box.append(table);
  return box;
}

const inputs = document.createElement("div");
inputs.className = "input-column";
inputs.append(modeBox, form.element, advancedBox);

/**
 * Print every disclosure open, and put them back afterwards.
 *
 * On paper a collapsed `<details>` is a heading that withholds what it covers:
 * the price assumptions, the advanced settings the figures were computed
 * under, the salary path that was typed. The reader cannot click it.
 *
 * This is script rather than a print rule because CSS cannot reach it -- a
 * closed `<details>` hides its content through an internal slot, so
 * `display: block` on the child computes correctly and still lays out a
 * zero-height box. Firefox and Chromium fire `beforeprint`/`afterprint`;
 * Safari only changes the `print` media query, so both are wired.
 */
function openForPrint(): void {
  let reclose: HTMLDetailsElement[] = [];

  const open = () => {
    reclose = [...document.querySelectorAll("details")].filter((d) => !d.open);
    for (const d of reclose) d.open = true;
  };
  const restore = () => {
    for (const d of reclose) d.open = false;
    reclose = [];
  };

  window.addEventListener("beforeprint", open);
  window.addEventListener("afterprint", restore);

  const printing = window.matchMedia("print");
  printing.addEventListener("change", (e) => (e.matches ? open() : restore()));
}

openForPrint();

const layout = document.createElement("div");
layout.className = "layout";
layout.append(inputs, results);

root.append(header, layout);
document.documentElement.lang = view.lang;
render();
