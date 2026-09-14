/**
 * The page: inputs on the left, results on the right, recomputed on every change.
 *
 * One model run is about ninety ages of arithmetic, well under a millisecond, so
 * there is no debounce and no incremental update -- every change re-runs the
 * model and rebuilds the results. That keeps the state in one place: whatever is
 * on screen is what `run()` returned for what the form says.
 *
 * The results follow the Start sheet's own order: Table 1, Figur 1, Figur 2,
 * the disposable income chart, Table 2.
 */
import { contextFromSettings, defaultInput, run } from "@typfallsmodellen/engine";
import type { ModelContext, TypfallInput, TypfallResult } from "@typfallsmodellen/engine";

import { renderDisposable, renderFigure1, renderFigure2 } from "./chart.js";
import type { FigureView } from "./chart.js";
import { loadDeathProbabilities } from "./deaths.js";
import { createForm } from "./form.js";
import { LANGS, dropHeadingNumber, t } from "./i18n.js";
import type { Lang, LabelName } from "./i18n.js";
import { renderKpis } from "./kpis.js";
import { renderTable1, renderTable2, table1ToCsv, table2ToCsv } from "./tables.js";
import type { Table1View } from "./tables.js";
import "./styles.css";

interface View {
  readonly lang: Lang;
  /** The workbook's own month/year switch. */
  readonly monthly: boolean;
}

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
  return contextFromSettings(new Map(), { chartEarningFactor: view.monthly ? 12 : 1 });
}

/**
 * The retirement age the run used, which is not always the one asked for.
 *
 * `startsetup` raises a retirement age below the cohort's earliest and says so
 * in a warning; every heading that names the age reads it back from there.
 */
function retirementAge(input: TypfallInput, result: TypfallResult): number {
  const corrected = result.warnings.find((w) => w.field === "ParYear");
  return typeof corrected?.used === "number" ? corrected.used : input.retirementAge;
}

let input: TypfallInput = defaultInput();
let view: View = { lang: "sv", monthly: true };

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
      form.relabel(lang);
      document.documentElement.lang = lang;
      render();
    });
    langs.append(button);
  }

  header.replaceChildren(title, sub, langs);
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

/**
 * A CSV download for one table. Not a workbook feature -- there is no SysLang
 * row for it -- so the button text is a plain per-language literal, the same
 * way the subtitle above is.
 */
function exportButton(lang: Lang, filename: string, csv: () => string): HTMLButtonElement {
  const button = document.createElement("button");
  button.type = "button";
  button.className = "export-btn";
  button.textContent = lang === "sv" ? "Ladda ner CSV" : "Download CSV";
  button.addEventListener("click", () => {
    // A BOM, so Excel reads å/ä/ö as UTF-8 instead of guessing a legacy
    // codepage from the bytes.
    const blob = new Blob(["﻿" + csv()], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    link.click();
    URL.revokeObjectURL(url);
  });
  return button;
}

function render(): void {
  renderHeading();
  const context = viewContext(view);
  const result = run(input, context, { deaths });
  const lang = view.lang;
  const par = retirementAge(input, result);

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

  results.replaceChildren();
  const warned = warnings(result);
  if (warned) results.append(warned);
  results.append(renderKpis(result, lang, par));
  const table1View: Table1View = {
    par,
    finalSalaryYears: context.finalSalaryYears,
    lastPensionRight: context.lastPensionRight > 0,
  };

  const table2Actions = document.createElement("div");
  table2Actions.className = "panel-actions";
  table2Actions.append(
    scaleToggle(),
    exportButton(lang, lang === "sv" ? "tabell2.csv" : "table2.csv", () =>
      table2ToCsv(result, lang),
    ),
  );

  results.append(
    section(
      table1Title,
      wrapScroll(renderTable1(result, lang, table1View)),
      exportButton(lang, lang === "sv" ? "tabell1.csv" : "table1.csv", () =>
        table1ToCsv(result, lang, table1View),
      ),
    ),
    renderFigure1(result, figures),
    renderFigure2(result, figures),
    renderDisposable(result, figures),
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

const layout = document.createElement("div");
layout.className = "layout";
layout.append(form.element, results);

root.append(header, layout);
document.documentElement.lang = view.lang;
render();
