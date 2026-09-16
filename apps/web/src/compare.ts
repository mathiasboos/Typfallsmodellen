/**
 * Jamfor scenarier: the baseline typfall, alongside a handful of variants that
 * each change only a few fields.
 *
 * There is no equivalent sheet in the workbook -- this is genuinely new UI,
 * the same class of addition as the tax-per-year chart. What it reuses is
 * everything downstream of a `TypfallInput`: `run()` itself (confirmed cheap,
 * well under a millisecond, the same reasoning `main.ts` already gives for
 * running it twice per render in Avancerat), and `renderKpis`/`renderTable1`,
 * which are pure functions of a result and take no more work to call four
 * times than once.
 *
 * A variant is not a diff against the baseline in the sense of "undefined
 * means inherit" -- it is a plain copy of the baseline's salary, retirement
 * age and occupational scheme, taken the moment the variant is added, that
 * the three controls below then edit independently. That is simpler than an
 * inherit/override state machine and reads the same way to whoever is using
 * it: a new scenario starts out identical to the baseline, and diverges only
 * where it is typed into. Every other field -- birth year, inflation, any
 * advanced setting, a typed salary vector -- keeps coming from the live
 * baseline on every run, so only these three are ever "frozen" per scenario.
 */
import { options } from "@typfallsmodellen/data";
import { RETIREMENT_AGES, run } from "@typfallsmodellen/engine";
import type {
  DeathProbabilities,
  ModelContext,
  SchemeId,
  TypfallInput,
} from "@typfallsmodellen/engine";

import { fieldSet, span } from "./controls.js";
import type { Lang } from "./i18n.js";
import { t } from "./i18n.js";
import { renderKpis, retirementAge } from "./kpis.js";
import { renderTable1 } from "./tables.js";
import type { Table1View } from "./tables.js";

const RETIREMENT = span(RETIREMENT_AGES);
const SCHEMES = options.choices.occupationalPension as readonly { value: number; label: string }[];
const MAX_VARIANTS = 3;

export interface ScenarioOverride {
  readonly id: string;
  label: string;
  monthlySalary: number;
  retirementAge: number;
  scheme: SchemeId;
}

/** The baseline, with just a variant's three fields laid over it. */
export function applyScenario(baseline: TypfallInput, s: ScenarioOverride): TypfallInput {
  return {
    ...baseline,
    monthlySalary: s.monthlySalary,
    retirementAge: s.retirementAge,
    scheme: s.scheme,
  };
}

function newVariant(id: string, label: string, baseline: TypfallInput): ScenarioOverride {
  return {
    id,
    label,
    monthlySalary: Math.round(baseline.monthlySalary),
    retirementAge: baseline.retirementAge,
    scheme: baseline.scheme,
  };
}

export interface CompareHandle {
  readonly element: HTMLElement;
  relabel(lang: Lang): void;
  /** Rebuilds every scenario's KPI row and Table 1 against the live baseline. */
  renderResults(baseline: TypfallInput, context: ModelContext, deaths: DeathProbabilities, lang: Lang): void;
}

const say = (l: Lang, sv: string, en: string) => (l === "sv" ? sv : en);

export function createComparePanel(
  lang: Lang,
  initialBaseline: TypfallInput,
  onChange: () => void,
): CompareHandle {
  let nextId = 1;
  // Updated on every `renderResults` call, so "add scenario" seeds a new
  // variant from the baseline as it stands now, not from the one the panel
  // happened to be constructed with.
  let lastBaseline = initialBaseline;
  const variants: ScenarioOverride[] = [newVariant(String(nextId), "Scenario 1", initialBaseline)];

  const element = document.createElement("div");
  element.className = "compare-panel";
  element.dataset.role = "compare-panel";

  const intro = document.createElement("p");
  intro.className = "field-hint";

  const row = document.createElement("div");
  row.className = "compare-row";

  const addButton = document.createElement("button");
  addButton.type = "button";
  addButton.className = "export-btn";
  addButton.dataset.action = "add-scenario";
  addButton.addEventListener("click", () => {
    if (variants.length >= MAX_VARIANTS) return;
    nextId += 1;
    variants.push(newVariant(String(nextId), `Scenario ${nextId}`, lastBaseline));
    build();
    onChange();
  });

  element.append(intro, row, addButton);

  let currentLang = lang;

  /** One card's controls and result area, rebuilt whenever the variant list changes. */
  function variantCard(variant: ScenarioOverride): HTMLElement {
    const card = document.createElement("section");
    card.className = "panel compare-card";
    card.dataset.scenario = variant.id;

    const head = document.createElement("div");
    head.className = "compare-card-head";
    const labelInput = document.createElement("input");
    labelInput.type = "text";
    labelInput.className = "compare-card-label";
    labelInput.value = variant.label;
    labelInput.addEventListener("change", () => {
      variant.label = labelInput.value.trim() || variant.label;
      onChange();
    });
    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "export-btn";
    remove.dataset.action = "remove-scenario";
    remove.textContent = say(currentLang, "Ta bort", "Remove");
    remove.disabled = variants.length <= 1;
    remove.addEventListener("click", () => {
      const index = variants.findIndex((v) => v.id === variant.id);
      if (index === -1 || variants.length <= 1) return;
      variants.splice(index, 1);
      build();
      onChange();
    });
    head.append(labelInput, remove);

    const controls = document.createElement("div");
    controls.className = "compare-card-controls";
    const relabels: ((l: Lang) => void)[] = [];
    const { field, number } = fieldSet(controls, relabels, currentLang);

    field(
      number(variant.monthlySalary, { min: 0, max: 1_000_000, step: 100 }, (v) => {
        variant.monthlySalary = v;
        onChange();
      }).element,
      (l) => ({ label: t("monthlySalary", l) }),
    );
    field(
      number(variant.retirementAge, { ...RETIREMENT, step: 1 }, (v) => {
        variant.retirementAge = v;
        onChange();
      }).element,
      (l) => ({ label: t("retirementAge", l) }),
    );

    const schemeSelect = document.createElement("select");
    for (const choice of SCHEMES) {
      const option = document.createElement("option");
      option.value = String(choice.value);
      option.textContent = choice.label;
      if (choice.value === variant.scheme) option.selected = true;
      schemeSelect.append(option);
    }
    schemeSelect.addEventListener("change", () => {
      variant.scheme = Number(schemeSelect.value) as SchemeId;
      onChange();
    });
    field(schemeSelect, (l) => ({ label: t("occupational", l) }), "field field-wide");

    const results = document.createElement("div");
    results.className = "compare-card-results";

    card.append(head, controls, results);
    return card;
  }

  function baselineCard(): HTMLElement {
    const card = document.createElement("section");
    card.className = "panel compare-card compare-card-baseline";
    card.dataset.scenario = "baseline";
    const head = document.createElement("div");
    head.className = "compare-card-head";
    const heading = document.createElement("strong");
    heading.textContent = say(currentLang, "Utgångsläge", "Baseline");
    head.append(heading);
    const results = document.createElement("div");
    results.className = "compare-card-results";
    card.append(head, results);
    return card;
  }

  /** Rebuilds the row of cards from scratch -- add/remove change how many
   * there are, so there is no cheaper update than redrawing all of them. */
  function build(): void {
    row.replaceChildren(baselineCard(), ...variants.map(variantCard));
    addButton.disabled = variants.length >= MAX_VARIANTS;
  }
  build();

  const applyText = (l: Lang) => {
    intro.textContent = say(
      l,
      "Jämför upp till tre alternativa scenarier mot utgångsläget till vänster -- var och ett " +
        "med sin egen månadslön, pensionsålder och tjänstepension, allt annat oförändrat.",
      "Compare up to three alternative scenarios against the baseline on the left -- each with " +
        "its own monthly salary, retirement age and occupational pension, everything else unchanged.",
    );
    addButton.textContent = say(l, "+ Lägg till scenario", "+ Add scenario");
  };
  applyText(lang);

  return {
    element,
    relabel(l) {
      currentLang = l;
      applyText(l);
      build();
    },
    renderResults(baseline, context, deaths, runLang) {
      lastBaseline = baseline;
      const cards = row.children;

      const renderInto = (resultsEl: Element, typfall: TypfallInput) => {
        const result = run(typfall, context, { deaths });
        const par = retirementAge(typfall, result);
        const table1View: Table1View = {
          par,
          finalSalaryYears: context.finalSalaryYears,
          lastPensionRight: context.lastPensionRight > 0,
        };
        // Table 1 is wider than a card, the same way it is wider than a phone
        // in the single-scenario view -- `.scroll` gives it its own scroller
        // instead of widening the card (`main.ts`'s own `wrapScroll`).
        const table1Box = document.createElement("div");
        table1Box.className = "scroll";
        table1Box.append(renderTable1(result, runLang, table1View));
        resultsEl.replaceChildren(renderKpis(result, runLang, par), table1Box);
      };

      const baselineResults = cards[0]?.querySelector(".compare-card-results");
      if (baselineResults) renderInto(baselineResults, baseline);

      for (const [index, variant] of variants.entries()) {
        const resultsEl = cards[index + 1]?.querySelector(".compare-card-results");
        if (resultsEl) renderInto(resultsEl, applyScenario(baseline, variant));
      }
    },
  };
}
