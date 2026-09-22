/**
 * Pensionsgrundande belopp -- the PGB sheet.
 *
 * The workbook credits pension rights for years without ordinary income four
 * ways: childcare years, which the model computes on its own from the Start
 * sheet's `childBirthYears`; sickness/activity compensation, hand-typed
 * kronor (PGB!row 4: "Ange manuellt"); conscription, a single date range the
 * sheet turns into kronor itself (PGB!row 5: "Lägg in datum"); and study, a
 * semester count the sheet turns into kronor from a per-year grant rate
 * (PGB!row 6: "Lägg in antingen antal terminer eller bidragsbelopp" -- the
 * "or a kronor amount" half of that instruction is not ported; a semester
 * count is what the manual's own example asks for). `packages/engine/src/
 * model/pgb.ts` is where the conscription and study kronor actually get
 * computed; this file only collects the dates and the semester counts.
 *
 * Unlike `salaryPath.ts`'s grid, there is no computed path to fill from or
 * fall back to -- every cell starts at zero (or empty, for the dates), and
 * only the nonzero/non-empty entries are ever handed to the engine.
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
import { conscriptionDaysByYear, studyPgb } from "@typfallsmodellen/engine";
import type { PgbConscriptionPeriod, PgbManualYear } from "@typfallsmodellen/engine";

import type { Lang } from "./i18n.js";

const FIRST_AGE = 16;
const LAST_AGE = 70;
const MAX_SEMESTERS = 2;

export interface PgbGridHandle {
  readonly element: HTMLElement;
  relabel(lang: Lang): void;
  /** Keeps the year column, and the study readout, in step with the Start
   * sheet's birth year and the "Marginal, avrundningar" setting. */
  setBaseline(born: number, marginal: number): void;
  /** Clears every cell, and the conscription dates, back to empty. */
  reset(): void;
}

interface Row {
  age: number;
  sa: number;
  studySemesters: number;
}

const say = (l: Lang, sv: string, en: string) => (l === "sv" ? sv : en);

export function createPgbGrid(
  lang: Lang,
  onChange: (patch: {
    pgbManual: readonly PgbManualYear[] | undefined;
    pgbConscription: PgbConscriptionPeriod | undefined;
  }) => void,
): PgbGridHandle {
  let born = 0;
  let marginal = 0;
  let rows: Row[] = [];
  for (let age = FIRST_AGE; age <= LAST_AGE; age += 1) rows.push({ age, sa: 0, studySemesters: 0 });

  const element = document.createElement("details");
  element.className = "adv-group";
  element.dataset.group = "pgb";

  const summary = document.createElement("summary");
  const body = document.createElement("div");
  body.className = "adv-body";

  const intro = document.createElement("p");
  intro.className = "field-hint";

  // Värnplikt: `wsPGB!H4`/`H5` are a single date range, not a per-age row --
  // the sheet's own instruction is "Cell H4 och H5: Lägg in datum", nothing
  // about a grid. `startInput`/`endInput` feed `PgbConscriptionPeriod`
  // straight through to the engine; `vplReadout` shows the same
  // year-by-year day split `conscriptionDaysByYear` computes there, so
  // typing a date pair reads back as confirmation before it ever shows up
  // in the pension figures.
  const vplBox = document.createElement("div");
  vplBox.className = "pgb-vpl";
  const vplFields = document.createElement("div");
  vplFields.className = "pgb-vpl-fields";
  const startWrap = document.createElement("label");
  startWrap.className = "field";
  const startLabel = document.createElement("span");
  startLabel.className = "field-label";
  const startInput = document.createElement("input");
  startInput.type = "date";
  startInput.min = "1995-01-01";
  startInput.dataset.setting = "pgbConscriptionStart";
  const startHint = document.createElement("span");
  startHint.className = "field-hint";
  startWrap.append(startLabel, startInput, startHint);

  const endWrap = document.createElement("label");
  endWrap.className = "field";
  const endLabel = document.createElement("span");
  endLabel.className = "field-label";
  const endInput = document.createElement("input");
  endInput.type = "date";
  endInput.dataset.setting = "pgbConscriptionEnd";
  endWrap.append(endLabel, endInput);

  vplFields.append(startWrap, endWrap);
  const vplReadout = document.createElement("p");
  vplReadout.className = "field-hint pgb-vpl-readout";
  vplBox.append(vplFields, vplReadout);

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
  // dialog: `expandBtn.after(scroll)` puts it back in the one place it can
  // go, right after the button that moves it, regardless of what else this
  // body gains later.
  dialog.addEventListener("close", () => {
    expandBtn.after(scroll);
  });
  document.body.append(dialog);

  expandBtn.addEventListener("click", () => {
    dialog.append(scroll);
    dialog.showModal();
  });

  body.append(intro, vplBox, expandBtn, scroll);
  element.append(summary, body);

  function conscriptionPeriod(): PgbConscriptionPeriod | undefined {
    if (!startInput.value || !endInput.value) return undefined;
    return { start: startInput.value, end: endInput.value };
  }

  function updateVplReadout(l: Lang): void {
    const period = conscriptionPeriod();
    if (period === undefined) {
      vplReadout.textContent = "";
      return;
    }
    const days = conscriptionDaysByYear(period);
    if (days.size === 0) {
      vplReadout.textContent = say(
        l,
        "Perioden är kortare än 120 dagar och ger ingen pensionsrätt.",
        "The period is under 120 days and earns no pension rights.",
      );
      return;
    }
    const perYear = [...days.entries()]
      .map(([year, count]) => `${year}: ${Math.round(count)} ${say(l, "dagar", "days")}`)
      .join(", ");
    vplReadout.textContent = say(l, `Registrerat: ${perYear}.`, `Recorded: ${perYear}.`);
  }

  function emit(): void {
    const nonzero = rows.filter((r) => r.sa !== 0 || r.studySemesters !== 0);
    onChange({
      pgbManual: nonzero.length > 0 ? nonzero.map((r) => ({ ...r })) : undefined,
      pgbConscription: conscriptionPeriod(),
    });
  }

  startInput.addEventListener("change", () => {
    updateVplReadout(currentLang);
    emit();
  });
  endInput.addEventListener("change", () => {
    updateVplReadout(currentLang);
    emit();
  });

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

      const saCell = document.createElement("td");
      const saInput = document.createElement("input");
      saInput.type = "number";
      saInput.inputMode = "numeric";
      saInput.min = "0";
      saInput.step = "100";
      saInput.value = String(row.sa);
      saInput.addEventListener("change", () => {
        const typed = Number(saInput.value);
        if (!Number.isFinite(typed) || saInput.value.trim() === "") {
          saInput.value = String(rows[index]!.sa);
          return;
        }
        const value = Math.max(Math.round(typed), 0);
        rows[index] = { ...rows[index]!, sa: value };
        saInput.value = String(value);
        emit();
      });
      saCell.append(saInput);

      const semesterCell = document.createElement("td");
      const semesterInput = document.createElement("input");
      semesterInput.type = "number";
      semesterInput.inputMode = "numeric";
      semesterInput.min = "0";
      semesterInput.max = String(MAX_SEMESTERS);
      semesterInput.step = "1";
      semesterInput.value = String(row.studySemesters);
      const studyCell = document.createElement("td");
      studyCell.className = "num pgb-study-kr";
      const updateStudyReadout = () => {
        const year = born > 0 ? Math.trunc(born) + rows[index]!.age : 0;
        const kr = born > 0 ? studyPgb(year, rows[index]!.studySemesters, marginal) : 0;
        studyCell.textContent = kr > 0 ? kr.toLocaleString("sv-SE") : "";
      };
      semesterInput.addEventListener("change", () => {
        const typed = Number(semesterInput.value);
        if (!Number.isFinite(typed) || semesterInput.value.trim() === "") {
          semesterInput.value = String(rows[index]!.studySemesters);
          return;
        }
        const value = Math.max(Math.min(Math.round(typed), MAX_SEMESTERS), 0);
        rows[index] = { ...rows[index]!, studySemesters: value };
        semesterInput.value = String(value);
        updateStudyReadout();
        emit();
      });
      semesterCell.append(semesterInput);
      updateStudyReadout();

      tr.append(year, age, saCell, semesterCell, studyCell);
      tbody.append(tr);
    }
  }

  drawGrid();

  let currentLang = lang;

  const applyText = (l: Lang) => {
    const title = say(l, "Pensionsgrundande belopp (PGB)", "Pension-qualifying amounts (PGB)");
    summary.textContent = title;
    dialogTitle.textContent = title;
    closeBtn.setAttribute("aria-label", say(l, "Stäng", "Close"));
    expandBtn.textContent = say(l, "Visa alla kolumner", "Show all columns");
    intro.textContent = say(
      l,
      "Sjuk- eller aktivitetsersättning, värnplikt och studier ger pensionsrätt utöver " +
        "barnår, som modellen redan räknar med.",
      "Sickness or activity compensation, conscription and study earn pension rights on " +
        "top of childcare years, which the model already accounts for.",
    );
    startLabel.textContent = say(l, "Värnplikt, startdatum", "Conscription, start date");
    startHint.textContent = say(
      l,
      "ej före 1995-01-01; ger bara pensionsrätt 1995–2010 och från 2018",
      "not before 1995-01-01; only earns pension rights 1995-2010 and from 2018 on",
    );
    endLabel.textContent = say(l, "Muck (slutdatum)", "End date");
    const heads = [
      say(l, "År", "Year"),
      say(l, "Ålder", "Age"),
      say(l, "Sjuk-/aktivitetsersättning", "Sickness/activity comp."),
      say(l, "Antal terminer", "Semesters"),
      say(l, "PGB studier, kr", "Study PGB, kr"),
    ];
    for (const [i, cell] of headCells.entries()) cell.textContent = heads[i]!;
    updateVplReadout(l);
  };
  applyText(lang);

  return {
    element,
    relabel(l) {
      currentLang = l;
      applyText(l);
    },
    setBaseline(bornYear, marginalValue) {
      if (bornYear === born && marginalValue === marginal) return;
      born = bornYear;
      marginal = marginalValue;
      // Every cell's own study readout reads `born`/`marginal` from this
      // closure, so a rebuild is the only way either change reaches it --
      // relabelling the year column in place, the way `salaryPath.ts` does,
      // would leave the study kronor stale.
      drawGrid();
    },
    reset() {
      rows = rows.map((row) => ({ ...row, sa: 0, studySemesters: 0 }));
      startInput.value = "";
      endInput.value = "";
      updateVplReadout(currentLang);
      drawGrid();
    },
  };
}
