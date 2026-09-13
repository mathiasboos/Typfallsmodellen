/**
 * The page: inputs on the left, results on the right, recomputed on every change.
 *
 * One model run is about ninety ages of arithmetic, well under a millisecond, so
 * there is no debounce and no incremental update -- every change re-runs the
 * model and rebuilds the results. That keeps the state in one place: whatever is
 * on screen is what `run()` returned for what the form says.
 */
import { content } from "@typfallsmodellen/data";
import { contextFromSettings, defaultInput, run } from "@typfallsmodellen/engine";
import type { ModelContext, TypfallInput, TypfallResult } from "@typfallsmodellen/engine";

import { renderComposition, renderIncome } from "./chart.js";
import { loadDeathProbabilities } from "./deaths.js";
import { createForm } from "./form.js";
import { LANGS, t } from "./i18n.js";
import type { Lang } from "./i18n.js";
import { renderTable1, renderTable2 } from "./tables.js";
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
 * division in the view. Table 1 carries both scales on every row, so it needs
 * nothing here.
 */
function viewContext(view: View): ModelContext {
  return contextFromSettings(new Map(), { chartEarningFactor: view.monthly ? 12 : 1 });
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
  title.textContent = "Typfallsmodellen";

  const sub = document.createElement("p");
  sub.className = "subtitle";
  sub.textContent =
    view.lang === "sv"
      ? `En portering av Pensionsmyndighetens typfallsmodell, ${content.modelVersion}. ` +
        `Alla beräkningar sker i din webbläsare; ingenting skickas någonstans.`
      : `A port of the Swedish Pensions Agency's typfallsmodell, ${content.modelVersion}. ` +
        `Everything is computed in your browser; nothing is sent anywhere.`;

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

function scaleSwitch(): HTMLElement {
  const box = document.createElement("div");
  box.className = "toggle";
  box.setAttribute("role", "group");
  box.setAttribute("aria-label", t("show", view.lang));

  const choices: readonly { monthly: boolean; text: string }[] = [
    { monthly: true, text: t("perMonth", view.lang) },
    { monthly: false, text: `${t("kronor", view.lang)} / ${t("years", view.lang)}` },
  ];
  for (const choice of choices) {
    const button = document.createElement("button");
    button.type = "button";
    button.textContent = choice.text;
    button.className = choice.monthly === view.monthly ? "chip active" : "chip";
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

function section(title: string, body: HTMLElement): HTMLElement {
  const wrap = document.createElement("section");
  wrap.className = "panel";
  const head = document.createElement("h2");
  head.textContent = title;
  wrap.append(head, body);
  return wrap;
}

function render(): void {
  renderHeading();
  const result = run(input, viewContext(view), { deaths });
  const lang = view.lang;

  const table2Title =
    lang === "sv"
      ? `Tabell 2. ${view.monthly ? "Månadsinkomster" : "Årsinkomster"} från ${
          result.table2[0]?.age ?? ""
        } års ålder`
      : `Table 2. ${view.monthly ? "Monthly" : "Annual"} income from age ${
          result.table2[0]?.age ?? ""
        }`;

  const composition =
    lang === "sv" ? "Vad inkomsten består av" : "What the income is made of";
  const levels = lang === "sv" ? "Brutto, efter skatt och disponibelt" : "Gross, after tax and disposable";

  results.replaceChildren();
  const warned = warnings(result);
  if (warned) results.append(warned);
  results.append(
    scaleSwitch(),
    section(t("table1", lang), renderTable1(result, lang, view.monthly)),
    renderComposition(result, lang, composition),
    renderIncome(result, lang, levels),
    section(table2Title, wrapScroll(renderTable2(result, lang))),
  );

  const foot = document.createElement("p");
  foot.className = "footnote";
  foot.textContent = t("kpiFootnote", lang);
  results.append(foot);
}

/** Table 2 is wider than a phone; give it its own scroller rather than the page. */
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
