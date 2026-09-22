/**
 * Egen löneutveckling -- the Indata_lista sheet, as a grid.
 *
 * Manual section 3.1: tick the box and the model stops deriving the wage path
 * and reads a typed one instead, one row per age with two columns. `Inkomst` is
 * taxed earned income; `Varav lön` is the part of it that pension contributions
 * are taken from, and private saving entered as a share of income is a share of
 * the first. Zeroing a year is how a year of leave is simulated; halving a run
 * of years is how part-time is.
 *
 * The grid is filled from the path the model just computed rather than left
 * empty, which is what the workbook does too -- `setup.ts:424` reads an age the
 * array does not mention as 0, so an empty grid would mean a lifetime of no
 * income, not "derive it for me". `TypfallResult.wagePath` is that fill, shaped
 * so it can be handed straight back.
 *
 * Ages below 15 are not shown: `validate` in setup.ts pins `startage` to 15
 * whenever an own vector is present, so those rows could never be read.
 *
 * Amounts are rounded to whole kronor, both shown and used, keeping form.ts's
 * rule that the form never shows a number the run did not use. Measured cost of
 * the rounding against the unrounded path: 0.02 kr per month on the final
 * salary, and less on everything else.
 */
import type { OwnIncomeYear } from "@typfallsmodellen/engine";

import type { Lang } from "./i18n.js";

/** The first age an own vector is ever read at -- `setup.ts:139`. */
const FIRST_AGE = 15;

export interface SalaryPathHandle {
  readonly element: HTMLElement;
  relabel(lang: Lang): void;
  /**
   * The path the model derives for the Start sheet as it currently stands,
   * ignoring anything typed here. Kept fresh so the grid fills from the right
   * numbers whenever it is switched on, and so `Återställ` has somewhere to go
   * back to after the birth year or the salary has moved.
   */
  setBaseline(path: readonly OwnIncomeYear[], born: number): void;
  /** Switches the grid off and drops the typed path. */
  reset(): void;
}

const say = (l: Lang, sv: string, en: string) => (l === "sv" ? sv : en);

export function createSalaryPath(
  lang: Lang,
  onChange: (path: readonly OwnIncomeYear[] | undefined) => void,
): SalaryPathHandle {
  let currentLang = lang;
  let baseline: readonly OwnIncomeYear[] = [];
  let born = 0;
  let rows: OwnIncomeYear[] = [];

  const element = document.createElement("details");
  element.className = "adv-group";
  element.dataset.group = "salary-path";

  const summary = document.createElement("summary");
  const body = document.createElement("div");
  body.className = "adv-body";

  const toggleWrap = document.createElement("label");
  toggleWrap.className = "field field-check";
  const toggle = document.createElement("input");
  toggle.type = "checkbox";
  toggle.dataset.setting = "ownIncome";
  const toggleLabel = document.createElement("span");
  toggleLabel.className = "field-label";
  const toggleHint = document.createElement("span");
  toggleHint.className = "field-hint";
  toggleWrap.append(toggle, toggleLabel, toggleHint);

  const actions = document.createElement("div");
  actions.className = "panel-actions adv-actions";
  const refill = document.createElement("button");
  refill.type = "button";
  refill.className = "panel-btn";
  actions.append(refill);

  const scroll = document.createElement("div");
  scroll.className = "adv-grid-scroll";
  const table = document.createElement("table");
  table.className = "table adv-grid";
  const head = document.createElement("thead");
  const headRow = document.createElement("tr");
  const headCells = [
    document.createElement("th"),
    document.createElement("th"),
    document.createElement("th"),
    document.createElement("th"),
  ];
  headRow.append(...headCells);
  head.append(headRow);
  const tbody = document.createElement("tbody");
  table.append(head, tbody);
  scroll.append(table);

  body.append(toggleWrap, actions, scroll);
  element.append(summary, body);

  /** Rebuilt whenever the path is refilled; edits mutate `rows` in place. */
  function drawGrid(): void {
    tbody.replaceChildren();
    for (const [index, row] of rows.entries()) {
      const tr = document.createElement("tr");

      const year = document.createElement("td");
      year.className = "num";
      year.textContent = born > 0 ? String(Math.trunc(born) + row.age) : "";

      const age = document.createElement("td");
      age.className = "num";
      age.textContent = String(row.age);

      const cell = (get: () => number, put: (v: number) => void) => {
        const td = document.createElement("td");
        const input = document.createElement("input");
        input.type = "number";
        input.inputMode = "numeric";
        input.min = "0";
        input.step = "100";
        input.value = String(get());
        input.addEventListener("change", () => {
          const typed = Number(input.value);
          if (!Number.isFinite(typed) || input.value.trim() === "") {
            input.value = String(get());
            return;
          }
          const value = Math.max(Math.round(typed), 0);
          put(value);
          input.value = String(value);
          onChange(rows);
        });
        td.append(input);
        return td;
      };

      tr.append(
        year,
        age,
        cell(
          () => rows[index]!.income,
          (v) => {
            rows[index] = { ...rows[index]!, income: v };
          },
        ),
        cell(
          () => rows[index]!.wage,
          (v) => {
            rows[index] = { ...rows[index]!, wage: v };
          },
        ),
      );
      tbody.append(tr);
    }
  }

  /** The baseline, trimmed to the ages an own vector is read at and rounded. */
  function fill(): OwnIncomeYear[] {
    return baseline
      .filter((r) => r.age >= FIRST_AGE)
      .map((r) => ({ age: r.age, income: Math.round(r.income), wage: Math.round(r.wage) }));
  }

  toggle.addEventListener("change", () => {
    if (toggle.checked) {
      rows = fill();
      drawGrid();
      scroll.hidden = false;
      actions.hidden = false;
      onChange(rows);
    } else {
      scroll.hidden = true;
      actions.hidden = true;
      onChange(undefined);
    }
  });

  refill.addEventListener("click", () => {
    rows = fill();
    drawGrid();
    onChange(rows);
  });

  scroll.hidden = true;
  actions.hidden = true;

  const applyText = (l: Lang) => {
    summary.textContent = say(l, "Egen löneutveckling", "Own salary path");
    toggleLabel.textContent = say(l, "Använd egen lönevektor", "Use an own salary vector");
    toggleHint.textContent = say(
      l,
      "Fylls i från den beräknade lönebanan, som du sedan kan ändra",
      "Filled in from the computed path, which you can then edit",
    );
    refill.textContent = say(l, "Återställ till beräknad lönebana", "Reset to the computed path");
    const heads = [
      say(l, "År", "Year"),
      say(l, "Ålder", "Age"),
      say(l, "Inkomst", "Income"),
      say(l, "Varav lön", "Of which salary"),
    ];
    for (const [i, cell] of headCells.entries()) cell.textContent = heads[i]!;
  };
  applyText(lang);

  return {
    element,
    relabel(l) {
      currentLang = l;
      applyText(currentLang);
    },
    setBaseline(path, bornYear) {
      baseline = path;
      // While the grid is off it tracks the Start sheet; once it is on, the
      // typed rows are the authority and only Återställ overwrites them.
      //
      // The year column still has to follow a changed birth year, but it is
      // the only thing that does -- and redrawing on every call would tear the
      // cell being edited out of the DOM inside its own change event, which
      // the offline check caught as a run of `replaceChildren` NotFoundErrors.
      // So the years are relabelled in place, and the rows are left alone.
      if (bornYear === born) return;
      born = bornYear;
      if (!toggle.checked) return;
      for (const [index, tr] of [...tbody.children].entries()) {
        const row = rows[index];
        if (row === undefined) continue;
        tr.firstElementChild!.textContent = born > 0 ? String(Math.trunc(born) + row.age) : "";
      }
    },
    reset() {
      toggle.checked = false;
      scroll.hidden = true;
      actions.hidden = true;
      rows = [];
    },
  };
}
