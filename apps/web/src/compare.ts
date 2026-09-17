/**
 * Jamfor scenarier: the baseline typfall, alongside a handful of variants that
 * each change only a few fields.
 *
 * There is no equivalent sheet in the workbook -- this is genuinely new UI,
 * the same class of addition as the tax-per-year chart. What it reuses is
 * everything downstream of a `TypfallInput`: `run()` itself (confirmed cheap,
 * well under a millisecond, the same reasoning `main.ts` already gives for
 * running it twice per render in Avancerat), and `renderCompareTable`/
 * `renderCompareChart` (`tables.ts`/`chart.ts`), which take no more work to
 * run against four scenarios than one.
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
 *
 * Results used to live inside each scenario's own card (a KPI row plus a full
 * Table 1) until a reviewer found that layout hard to actually compare row by
 * row and hand-built a spreadsheet mockup of what they wanted instead: one
 * table, every row lined up across scenarios, plus an overlay chart. Both now
 * live in one shared section below the cards, which are input-only.
 */
import { options } from "@typfallsmodellen/data";
import { RETIREMENT_AGES, run } from "@typfallsmodellen/engine";
import type {
  DeathProbabilities,
  ModelContext,
  SchemeId,
  TypfallInput,
} from "@typfallsmodellen/engine";

import { renderCompareChart } from "./chart.js";
import type { FigureView } from "./chart.js";
import { fieldSet, span } from "./controls.js";
import type { Lang } from "./i18n.js";
import { t } from "./i18n.js";
import { retirementAge } from "./kpis.js";
import { renderCompareTable } from "./tables.js";
import type { ScenarioColumn } from "./tables.js";

const RETIREMENT = span(RETIREMENT_AGES);
const SCHEMES = options.choices.occupationalPension as readonly { value: number; label: string }[];
const MAX_VARIANTS = 3;

/**
 * One colour per card position, not per `id` -- so a card's own accent, its
 * line in the chart and its column in the table always agree even as
 * variants are added or removed, at the cost of a variant's colour shifting
 * if an earlier one is removed. Validated (`dataviz` skill, `validate_palette.
 * js`) against both this app's chart surfaces (`--surface-raised`, light
 * `#ffffff` and dark `#003824`) as a fixed-order categorical set.
 */
const SCENARIO_COLOURS = ["--fig-scenario-0", "--fig-scenario-1", "--fig-scenario-2", "--fig-scenario-3"] as const;

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
  /** Rebuilds the comparison chart and table against the live baseline. */
  renderResults(
    baseline: TypfallInput,
    context: ModelContext,
    deaths: DeathProbabilities,
    lang: Lang,
    perMonth: boolean,
  ): void;
}

const say = (l: Lang, sv: string, en: string) => (l === "sv" ? sv : en);

function swatch(colour: string): HTMLElement {
  const el = document.createElement("span");
  el.className = "compare-swatch";
  el.style.background = `var(${colour})`;
  return el;
}

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
  // Plain text, like a variant's own label -- typed once, then kept as-is
  // across language switches rather than re-translated on every `relabel`.
  let baselineLabel = say(lang, "Utgångsläge", "Baseline");

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

  const results = document.createElement("div");
  results.className = "compare-results";

  element.append(intro, row, addButton, results);

  let currentLang = lang;

  /** One variant's controls, rebuilt whenever the variant list changes. */
  function variantCard(variant: ScenarioOverride, colour: string): HTMLElement {
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
    head.append(swatch(colour), labelInput, remove);

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

    card.append(head, controls);
    return card;
  }

  function baselineCard(): HTMLElement {
    const card = document.createElement("section");
    card.className = "panel compare-card compare-card-baseline";
    card.dataset.scenario = "baseline";
    const head = document.createElement("div");
    head.className = "compare-card-head";
    const labelInput = document.createElement("input");
    labelInput.type = "text";
    labelInput.className = "compare-card-label";
    labelInput.value = baselineLabel;
    labelInput.addEventListener("change", () => {
      baselineLabel = labelInput.value.trim() || baselineLabel;
      onChange();
    });
    head.append(swatch(SCENARIO_COLOURS[0]), labelInput);
    card.append(head);
    return card;
  }

  /** Rebuilds the row of cards from scratch -- add/remove change how many
   * there are, so there is no cheaper update than redrawing all of them. */
  function build(): void {
    row.replaceChildren(
      baselineCard(),
      ...variants.map((variant, i) => variantCard(variant, SCENARIO_COLOURS[i + 1] ?? SCENARIO_COLOURS[0])),
    );
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
    renderResults(baseline, context, deaths, runLang, perMonth) {
      lastBaseline = baseline;

      const columns: ScenarioColumn[] = [baseline, ...variants.map((v) => applyScenario(baseline, v))].map(
        (typfall, i) => {
          const result = run(typfall, context, { deaths });
          return {
            label: i === 0 ? baselineLabel : variants[i - 1]!.label,
            colour: SCENARIO_COLOURS[i] ?? SCENARIO_COLOURS[0],
            result,
            par: retirementAge(typfall, result),
          };
        },
      );

      const figureView: FigureView = {
        lang: runLang,
        par: columns[0]!.par,
        perMonth,
        priceBasis: context.priceBasis,
      };

      const table = document.createElement("div");
      table.className = "scroll";
      table.append(
        renderCompareTable(columns, runLang, {
          finalSalaryYears: context.finalSalaryYears,
          lastPensionRight: context.lastPensionRight > 0,
        }),
      );

      results.replaceChildren(renderCompareChart(columns, figureView), table);
    },
  };
}
