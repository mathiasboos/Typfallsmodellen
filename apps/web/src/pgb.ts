/**
 * Pensionsgrundande belopp -- the PGB sheet, as a grid.
 *
 * The workbook credits pension rights for years without ordinary income in two
 * ways: childcare years, which the model computes on its own from the Start
 * sheet's `childBirthYears`, and sickness/activity compensation, conscription
 * and study, which are hand-typed on the PGB sheet and ship as zero
 * (`packages/engine/src/model/input.ts`'s comment on `pgbManual`: "the shipped
 * workbook has none"). This grid is that sheet.
 *
 * Unlike `salaryPath.ts`'s grid, there is no computed path to fill from or
 * fall back to -- every cell starts at zero, and only the nonzero rows are
 * ever handed to the engine (an all-zero row and an absent one behave
 * identically in `earnPgb`, which only ever adds a manual amount in).
 *
 * The row range is a fixed 16 through 70, not tied to the wage path the way
 * `salaryPath.ts`'s rows are: `earnPgb` (packages/engine/src/model/mcalc.ts)
 * only ever reads a manual entry for `age > 15 && age <= riktalder`, and
 * `context.riktage` -- the only riktålder this port has today -- defaults to
 * 66 for every cohort ("cohort table pending"), so 70 leaves a safe margin
 * without inventing a per-cohort range the workbook itself does not model yet.
 * A fixed range also means the row list never has to be rebuilt (and typed
 * values never have to be reconciled) when unrelated inputs change.
 */
import type { PgbManualYear } from "@typfallsmodellen/engine";

import type { Lang } from "./i18n.js";

const FIRST_AGE = 16;
const LAST_AGE = 70;

export interface PgbGridHandle {
  readonly element: HTMLElement;
  relabel(lang: Lang): void;
  /** Keeps the year column in step with the Start sheet's birth year. */
  setBaseline(born: number): void;
  /** Clears every cell back to zero. */
  reset(): void;
}

const say = (l: Lang, sv: string, en: string) => (l === "sv" ? sv : en);

export function createPgbGrid(
  lang: Lang,
  onChange: (pgb: readonly PgbManualYear[] | undefined) => void,
): PgbGridHandle {
  let born = 0;
  let rows: PgbManualYear[] = [];
  for (let age = FIRST_AGE; age <= LAST_AGE; age += 1) rows.push({ age, sa: 0, vpl: 0, studier: 0 });

  const element = document.createElement("details");
  element.className = "adv-group";
  element.dataset.group = "pgb";

  const summary = document.createElement("summary");
  const body = document.createElement("div");
  body.className = "adv-body";

  const intro = document.createElement("p");
  intro.className = "field-hint";

  // Five columns split across the sidebar's own width clip a sixth digit
  // (see `.pgb-grid`'s own comment in styles.css), so `.pgb-grid` carries a
  // 380px floor and the grid scrolls sideways in its narrow `.adv-grid-scroll`
  // box to show it all -- reported as having to scroll to see the rest of the
  // table. `expandBtn` moves the same table (same nodes, same listeners, no
  // duplicated state) into a `<dialog>` instead: freed from the sidebar, the
  // dialog is wide enough on its own that the grid needs no min-width floor
  // and no horizontal scroll to show every column at once.
  const expandBtn = document.createElement("button");
  expandBtn.type = "button";
  expandBtn.className = "export-btn";
  expandBtn.dataset.action = "pgb-expand";

  const scroll = document.createElement("div");
  scroll.className = "adv-grid-scroll";
  const table = document.createElement("table");
  table.className = "table adv-grid pgb-grid";
  const head = document.createElement("thead");
  const headRow = document.createElement("tr");
  const headCells = [
    document.createElement("th"),
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

  const dialog = document.createElement("dialog");
  dialog.className = "pgb-dialog";
  const dialogHead = document.createElement("div");
  dialogHead.className = "pgb-dialog-head";
  const dialogTitle = document.createElement("h3");
  const closeBtn = document.createElement("button");
  closeBtn.type = "button";
  closeBtn.className = "dialog-close";
  closeBtn.textContent = "×";
  closeBtn.addEventListener("click", () => dialog.close());
  dialogHead.append(dialogTitle, closeBtn);
  dialog.append(dialogHead);
  // Clicking the backdrop -- a click landing on the <dialog> element itself
  // rather than anything inside it -- closes it the same way Escape does.
  dialog.addEventListener("click", (event) => {
    if (event.target === dialog) dialog.close();
  });
  // Fires for every path out (the close button, the backdrop, Escape), so the
  // grid always ends up back where it started, never stranded in a closed
  // dialog: `intro.after(scroll)` rather than `body.append(scroll)` so a
  // future field appended after `scroll` would not silently reorder it.
  dialog.addEventListener("close", () => {
    intro.after(scroll);
  });
  document.body.append(dialog);

  expandBtn.addEventListener("click", () => {
    dialog.append(scroll);
    dialog.showModal();
  });

  body.append(intro, expandBtn, scroll);
  element.append(summary, body);

  function emit(): void {
    const nonzero = rows.filter((r) => r.sa !== 0 || r.vpl !== 0 || r.studier !== 0);
    onChange(nonzero.length > 0 ? nonzero : undefined);
  }

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
          emit();
        });
        td.append(input);
        return td;
      };

      tr.append(
        year,
        age,
        cell(
          () => rows[index]!.sa,
          (v) => {
            rows[index] = { ...rows[index]!, sa: v };
          },
        ),
        cell(
          () => rows[index]!.vpl,
          (v) => {
            rows[index] = { ...rows[index]!, vpl: v };
          },
        ),
        cell(
          () => rows[index]!.studier,
          (v) => {
            rows[index] = { ...rows[index]!, studier: v };
          },
        ),
      );
      tbody.append(tr);
    }
  }

  drawGrid();

  const applyText = (l: Lang) => {
    const title = say(l, "Pensionsgrundande belopp (PGB)", "Pension-qualifying amounts (PGB)");
    summary.textContent = title;
    dialogTitle.textContent = title;
    closeBtn.setAttribute("aria-label", say(l, "Stäng", "Close"));
    expandBtn.textContent = say(l, "Visa alla kolumner", "Show all columns");
    intro.textContent = say(
      l,
      "Sjuk- eller aktivitetsersättning, värnplikt och studier ger pensionsrätt utöver " +
        "barnår, som modellen redan räknar med. Fyll i den ålder respektive belopp gäller.",
      "Sickness or activity compensation, conscription and study earn pension rights on " +
        "top of childcare years, which the model already accounts for. Fill in the age each " +
        "amount applies at.",
    );
    const heads = [
      say(l, "År", "Year"),
      say(l, "Ålder", "Age"),
      say(l, "Sjuk-/aktivitetsersättning", "Sickness/activity comp."),
      say(l, "Värnplikt", "Conscription"),
      say(l, "Studier", "Study"),
    ];
    for (const [i, cell] of headCells.entries()) cell.textContent = heads[i]!;
  };
  applyText(lang);

  return {
    element,
    relabel(l) {
      applyText(l);
    },
    setBaseline(bornYear) {
      if (bornYear === born) return;
      born = bornYear;
      for (const [index, tr] of [...tbody.children].entries()) {
        const row = rows[index];
        if (row === undefined) continue;
        tr.firstElementChild!.textContent = born > 0 ? String(Math.trunc(born) + row.age) : "";
      }
    },
    reset() {
      rows = rows.map((row) => ({ ...row, sa: 0, vpl: 0, studier: 0 }));
      drawGrid();
    },
  };
}
