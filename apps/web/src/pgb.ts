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
 * The grid has two grouped columns, "Studier" and "Värnplikt", each an input
 * beside the kronor (and, for Värnplikt, the days) it produces -- mirroring
 * the sheet's own PGB tab, which lays conscription and study out as Year,
 * Age, Days, PGB amount and Year, Age, Semesters, PGB amount respectively,
 * rather than as a single line of text summarising the whole period.
 *
 * The two conscription dates live in the Värnplikt group's own header cell,
 * not in a box above the grid -- reported as unintuitive that typing a date
 * range somewhere else changed a table further down the page, with nothing
 * visually tying the two together. `wsPGB!H4`/`H5` are still a single
 * period, not a per-age entry (the sheet's own instruction is "Lägg in
 * datum", nothing about a grid), so the inputs stay singular rather than
 * becoming two more per-row columns; what moved is only *where* that single
 * pair sits, from a separate `<div>` to the same `<th>` that already names
 * the Dagar/PGB värnplikt columns the dates fill in, right there in the table.
 *
 * The conscription kronor need `medelPgi`, an economic projection that
 * depends on the run's own inflation/growth/price-basis assumptions --
 * unlike study's kronor, which are a pure function of a year and a semester
 * count and so are still computed client-side. There is no reaching that
 * projection from here without duplicating a chunk of `setup.ts`, so the
 * conscription column instead reads back `TypfallResult.pgbBreakdown`, the
 * same figures `earnPgb` used in the pension it just computed -- `setBaseline`
 * takes it as a third argument, refreshed on every render the way
 * `salaryPath.ts`'s own baseline is.
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
import type {
  PgbBreakdownYear,
  PgbConscriptionPeriod,
  PgbManualYear,
} from "@typfallsmodellen/engine";

import type { Lang } from "./i18n.js";

const FIRST_AGE = 16;
const LAST_AGE = 70;
const MAX_SEMESTERS = 2;

export interface PgbGridHandle {
  readonly element: HTMLElement;
  relabel(lang: Lang): void;
  /** Keeps the year column, the study readout and the conscription
   * days/kronor columns in step with the Start sheet's birth year, the
   * "Marginal, avrundningar" setting, and the latest run's own PGB figures. */
  setBaseline(born: number, marginal: number, pgbBreakdown: readonly PgbBreakdownYear[]): void;
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
  let breakdown = new Map<number, PgbBreakdownYear>();
  let breakdownDigest = "";
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

  // Seven columns split across the sidebar's own width clip the later ones
  // (see `.pgb-grid`'s own comment in styles.css), so `.pgb-grid` carries a
  // min-width floor and the grid scrolls sideways in its narrow `.adv-grid-scroll`
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

  // One <col> per physical column, in order -- styles.css sets every width
  // on these rather than on the header cells, since the two-row header below
  // means a header cell's position in its own row no longer lines up with
  // its column index once earlier cells start spanning rows or columns.
  const colgroup = document.createElement("colgroup");
  for (let i = 0; i < 7; i += 1) colgroup.append(document.createElement("col"));
  table.append(colgroup);

  // Two header rows: Year/Age/Sjuk- span both (`rowSpan`, so their own label
  // is not repeated), and Studier/Värnplikt each span two columns in the top
  // row -- the input beside the kronor (and, for Värnplikt, the days) it
  // produces -- so the pairing reads as one group rather than two unrelated
  // columns that happen to sit next to each other.
  const head = document.createElement("thead");
  const groupRow = document.createElement("tr");
  const yearHead = document.createElement("th");
  yearHead.rowSpan = 2;
  const ageHead = document.createElement("th");
  ageHead.rowSpan = 2;
  const saHead = document.createElement("th");
  saHead.rowSpan = 2;
  const studyGroupHead = document.createElement("th");
  studyGroupHead.colSpan = 2;
  const vplGroupHead = document.createElement("th");
  vplGroupHead.colSpan = 2;
  vplGroupHead.className = "pgb-vpl-head";
  groupRow.append(yearHead, ageHead, saHead, studyGroupHead, vplGroupHead);

  // The one pair of date inputs `wsPGB!H4`/`H5` are, right in the Värnplikt
  // group's own header cell -- reported as unintuitive that typing a date
  // range in a box elsewhere on the page changed a table further down,
  // nothing showing the two were connected. `startInput`/`endInput` feed
  // `PgbConscriptionPeriod` straight through to the engine; `vplReadout`
  // covers the one case the Dagar/PGB värnplikt columns below cannot show on
  // their own -- a period under 120 days has no touched-year row for a zero.
  const vplTitle = document.createElement("span");
  vplTitle.className = "pgb-vpl-title";
  const startLabel = document.createElement("span");
  const startInput = document.createElement("input");
  startInput.type = "date";
  startInput.min = "1995-01-01";
  startInput.dataset.setting = "pgbConscriptionStart";
  const startWrap = document.createElement("label");
  startWrap.className = "pgb-vpl-date";
  startWrap.append(startLabel, startInput);

  const endLabel = document.createElement("span");
  const endInput = document.createElement("input");
  endInput.type = "date";
  endInput.dataset.setting = "pgbConscriptionEnd";
  const endWrap = document.createElement("label");
  endWrap.className = "pgb-vpl-date";
  endWrap.append(endLabel, endInput);

  const vplReadout = document.createElement("p");
  vplReadout.className = "pgb-vpl-readout";
  vplGroupHead.append(vplTitle, startWrap, endWrap, vplReadout);

  const subRow = document.createElement("tr");
  const semesterHead = document.createElement("th");
  const studyKrHead = document.createElement("th");
  const vplDaysHead = document.createElement("th");
  const vplKrHead = document.createElement("th");
  subRow.append(semesterHead, studyKrHead, vplDaysHead, vplKrHead);

  head.append(groupRow, subRow);
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

  body.append(intro, expandBtn, scroll);
  element.append(summary, body);

  function conscriptionPeriod(): PgbConscriptionPeriod | undefined {
    if (!startInput.value || !endInput.value) return undefined;
    return { start: startInput.value, end: endInput.value };
  }

  // Only the one case the grid itself cannot show: a period too short to earn
  // anything has no touched-year rows to display a zero in. A valid period's
  // own days and kronor show in the grid, per year, so this stays blank then
  // rather than repeating the same figures as a line of text above it.
  function updateVplReadout(l: Lang): void {
    const period = conscriptionPeriod();
    if (period === undefined) {
      vplReadout.textContent = "";
      return;
    }
    const days = conscriptionDaysByYear(period);
    vplReadout.textContent =
      days.size === 0
        ? say(
            l,
            "Perioden är kortare än 120 dagar och ger ingen pensionsrätt.",
            "The period is under 120 days and earns no pension rights.",
          )
        : "";
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
      studyCell.className = "num pgb-computed";
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

      // Read back, not computed here: conscription's kronor need `medelPgi`,
      // an economic projection this file has no way to reach on its own (see
      // the file's own top comment) -- `breakdown` is the latest run's own
      // figures, refreshed by `setBaseline` on every render.
      const entry = breakdown.get(row.age);
      const vplDaysCell = document.createElement("td");
      vplDaysCell.className = "num pgb-computed";
      vplDaysCell.textContent = entry && entry.vplDays > 0 ? String(Math.round(entry.vplDays)) : "";
      const vplKrCell = document.createElement("td");
      vplKrCell.className = "num pgb-computed";
      vplKrCell.textContent = entry && entry.vpl > 0 ? entry.vpl.toLocaleString("sv-SE") : "";

      tr.append(year, age, saCell, semesterCell, studyCell, vplDaysCell, vplKrCell);
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
        "barnår, som modellen redan räknar med. Värnpliktens datum fylls i uppe i tabellen, " +
        "under Värnplikt (ger bara pensionsrätt 1995–2010 och från 2018).",
      "Sickness or activity compensation, conscription and study earn pension rights on " +
        "top of childcare years, which the model already accounts for. Conscription's own " +
        "dates are filled in up in the table, under Värnplikt (only earns pension rights " +
        "1995-2010 and from 2018 on).",
    );
    startLabel.textContent = say(l, "Från", "From");
    startInput.setAttribute("aria-label", say(l, "Värnplikt, startdatum", "Conscription, start date"));
    endLabel.textContent = say(l, "Till", "To");
    endInput.setAttribute("aria-label", say(l, "Muck (slutdatum)", "End date"));
    yearHead.textContent = say(l, "År", "Year");
    ageHead.textContent = say(l, "Ålder", "Age");
    saHead.textContent = say(l, "Sjuk-/aktivitetsersättning", "Sickness/activity comp.");
    studyGroupHead.textContent = say(l, "Studier", "Study");
    vplTitle.textContent = say(l, "Värnplikt", "Conscription");
    semesterHead.textContent = say(l, "Antal terminer", "Semesters");
    studyKrHead.textContent = say(l, "PGB studier, kr", "Study PGB, kr");
    vplDaysHead.textContent = say(l, "Dagar", "Days");
    vplKrHead.textContent = say(l, "PGB värnplikt, kr", "Conscription PGB, kr");
    updateVplReadout(l);
  };
  applyText(lang);

  return {
    element,
    relabel(l) {
      currentLang = l;
      applyText(l);
    },
    setBaseline(bornYear, marginalValue, pgbBreakdown) {
      // `pgbBreakdown` is a handful of entries at most (only touched ages),
      // so a stringified comparison is cheap -- and it is what actually
      // decides whether the conscription columns need to change, since
      // `medelPgi` can move with the run's own economic assumptions without
      // `born` or `marginal` moving at all.
      const digest = JSON.stringify(pgbBreakdown);
      if (bornYear === born && marginalValue === marginal && digest === breakdownDigest) return;
      born = bornYear;
      marginal = marginalValue;
      breakdown = new Map(pgbBreakdown.map((row) => [row.age, row]));
      breakdownDigest = digest;
      // Every cell's own study readout reads `born`/`marginal` from this
      // closure, and the conscription cells read `breakdown`, so a rebuild
      // is the only way any of the three reaches them -- relabelling the
      // year column in place, the way `salaryPath.ts` does, would leave the
      // rest stale.
      drawGrid();
    },
    reset() {
      rows = rows.map((row) => ({ ...row, sa: 0, studySemesters: 0 }));
      startInput.value = "";
      endInput.value = "";
      breakdown = new Map();
      breakdownDigest = "";
      updateVplReadout(currentLang);
      drawGrid();
    },
  };
}
