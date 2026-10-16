/**
 * Pensionsgrundande belopp -- the PGB sheet.
 *
 * The workbook credits pension rights for years without ordinary income four
 * ways: childcare years, computed from the Start sheet's `childBirthYears`;
 * sickness/activity compensation, hand-typed kronor (PGB!row 4: "Ange
 * manuellt"); conscription, a single date range the sheet turns into kronor
 * itself (PGB!row 5: "Lägg in datum"); and study, a semester count the sheet
 * turns into kronor from a per-year grant rate (PGB!row 6). `packages/
 * engine/src/model/pgb.ts` is where the conscription and study kronor
 * actually get computed; this file only collects the raw inputs.
 *
 * This used to be an always-visible 55-row grid (ages 16-70), mostly blank,
 * with the four child-birth-year fields above it -- reported as hard to use:
 * almost every cell is empty for almost every user, and the one thing anyone
 * actually wants (which years have a PGB credit, and how much) was buried in
 * a wall of zeros. This is a small add-entry form instead -- pick a type,
 * fill in that type's own fields, click "Lägg till" -- plus a summary table
 * that only ever shows years and categories that actually have data, per the
 * user's own mockup.
 *
 * Showing a "Child PGB, kr" column needed a real engine change, not just a
 * client-side read: `context.childBirthYears` already drove a per-age credit
 * inside `earnPgb` (`packages/engine/src/model/mcalc.ts`), but it was folded
 * straight into `RunState.pgb` with no way to tell it apart from the other
 * three sources, and `TypfallResult.pgbBreakdown` only ever flattened
 * `TypfallInput.pgbManual` -- which structurally can't see a `ModelContext`
 * field. `RunState.pgbBarn` and `PgbBreakdownYear.barn` (both new) expose it.
 *
 * A single child's credit can land in up to four consecutive calendar years
 * (the birth year and the three after), and if two children's four-year
 * windows overlap the same year, only the higher-priority child's credit
 * counts for that year -- the other is silently dropped (`pgbBarn`,
 * `packages/engine/src/pension/incomePension.ts`, original workbook
 * behaviour, not a bug). The "Child PGB, kr" column header carries a short
 * info tooltip explaining this, matching `tables.ts`'s own `TABLE2_COLUMNS`
 * pattern -- the column shows a number either way, so this is not otherwise
 * visible.
 *
 * Sickness/activity compensation and study are each keyed by (year, that
 * type), so re-adding the same year replaces whatever was there -- editing
 * is "type it again, or remove and re-add". Conscription stays a single
 * period, same as before. Children live in their own small list beside the
 * form, one always-editable birth-year field per filled slot -- the summary
 * table's own "Child PGB, kr" column stays read-only, since attributing a
 * given year's credit back to a specific child for removal would mean
 * duplicating `pgbBarn`'s own window/priority logic client-side just to
 * target a click.
 */
import { conscriptionDaysByYear, conscriptionEligible, defaultContext } from "@typfallsmodellen/engine";
import type {
  PgbBreakdownYear,
  PgbConscriptionPeriod,
  PgbManualYear,
} from "@typfallsmodellen/engine";

import { fieldSet } from "./controls.js";
import { kronor } from "./format.js";
import type { Lang } from "./i18n.js";
import { t } from "./i18n.js";
import { cell, headCell } from "./tables.js";

const FIRST_AGE = 16;
const LAST_AGE = 70;
const MAX_SEMESTERS = 2;

const TYPE_CHILD = 0;
const TYPE_CONSCRIPTION = 1;
const TYPE_SICKNESS = 2;
const TYPE_STUDY = 3;
type PgbEntryType = typeof TYPE_CHILD | typeof TYPE_CONSCRIPTION | typeof TYPE_SICKNESS | typeof TYPE_STUDY;

export interface PgbGridHandle {
  readonly element: HTMLElement;
  relabel(lang: Lang): void;
  /** Keeps the age readout and the summary table in step with the Start
   * sheet's birth year and the latest run's own PGB figures. */
  setBaseline(born: number, marginal: number, pgbBreakdown: readonly PgbBreakdownYear[]): void;
  /** Clears every entry, the conscription period and the children list. */
  reset(): void;
}

const say = (l: Lang, sv: string, en: string) => (l === "sv" ? sv : en);

const childOrdinal = (slot: number, l: Lang): string =>
  [
    say(l, "1:a barnet", "1st child"),
    say(l, "2:a barnet", "2nd child"),
    say(l, "3:e barnet", "3rd child"),
    say(l, "4:e barnet", "4th child"),
  ][slot] ?? "";

const TYPE_CHOICES: readonly { value: PgbEntryType; label: (l: Lang) => string }[] = [
  { value: TYPE_CHILD, label: (l) => say(l, "Barn", "Child") },
  { value: TYPE_CONSCRIPTION, label: (l) => say(l, "Värnplikt", "Conscription") },
  { value: TYPE_SICKNESS, label: (l) => say(l, "Sjuk-/aktivitetsersättning", "Sickness/activity comp.") },
  { value: TYPE_STUDY, label: (l) => say(l, "Studier", "Study") },
];

const SLOT_CHOICES: readonly { value: number; label: (l: Lang) => string }[] = [1, 2, 3, 4].map((n) => ({
  value: n,
  label: (l: Lang) => childOrdinal(n - 1, l),
}));

interface SummaryRow {
  readonly age: number;
  readonly year: number;
  readonly sa: number;
  readonly vpl: number;
  readonly studier: number;
  readonly barn: number;
  readonly total: number;
}

const CHILD_INFO = (l: Lang) =>
  say(
    l,
    "Ett barns pensionsgrundande belopp kan falla ut för upp till fyra år (födelseåret och de tre " +
      "följande). Om flera barns fyraårsperioder överlappar samma år räknas bara ett barns belopp det året.",
    "A child's PGB credit can land in up to four years (the birth year and the three after). If more " +
      "than one child's four-year window overlaps the same year, only one child's credit counts that year.",
  );

const SUMMARY_COLUMNS: readonly {
  readonly key: "barn" | "studier" | "vpl" | "sa";
  readonly head: (l: Lang) => string;
  readonly info?: (l: Lang) => string;
  readonly get: (r: SummaryRow) => number;
}[] = [
  { key: "barn", head: (l) => say(l, "Barn-PGB, kr", "Child PGB, kr"), info: CHILD_INFO, get: (r) => r.barn },
  { key: "studier", head: (l) => say(l, "PGB studier, kr", "Study PGB, kr"), get: (r) => r.studier },
  { key: "vpl", head: (l) => say(l, "PGB värnplikt, kr", "Conscription PGB, kr"), get: (r) => r.vpl },
  {
    key: "sa",
    head: (l) => say(l, "Sjuk-/aktivitetsersättning, kr", "Sickness/activity comp, kr"),
    get: (r) => r.sa,
  },
];

export function createPgbGrid(
  lang: Lang,
  onChange: (patch: {
    pgbManual: readonly PgbManualYear[] | undefined;
    pgbConscription: PgbConscriptionPeriod | undefined;
    childBirthYears: readonly [number, number, number, number];
  }) => void,
): PgbGridHandle {
  let born = 0;
  let marginal = 0;
  let breakdown: readonly PgbBreakdownYear[] = [];
  let breakdownDigest = "";

  let sicknessByAge = new Map<number, number>();
  let studyByAge = new Map<number, number>();
  let conscription: PgbConscriptionPeriod | undefined;
  let childBirthYears = [...defaultContext().childBirthYears] as [number, number, number, number];

  let typeValue: PgbEntryType = TYPE_CHILD;
  let yearValue = 0;
  let slotValue = 1;
  let amountValue = 0;
  let semesterValue = 1;

  const element = document.createElement("details");
  element.className = "adv-group";
  element.dataset.group = "pgb";

  const summary = document.createElement("summary");
  const body = document.createElement("div");
  body.className = "adv-body";

  const intro = document.createElement("p");
  intro.className = "pgb-intro";

  // ---------------------------------------------------------------- form --
  const formBox = document.createElement("div");
  formBox.className = "pgb-form";
  const relabels: ((l: Lang) => void)[] = [];
  const { field, number, select } = fieldSet(formBox, relabels, lang);

  const typeSelect = select(TYPE_CHOICES, typeValue, (v) => {
    typeValue = v as PgbEntryType;
    onTypeChange();
  });
  typeSelect.element.dataset.setting = "pgbEntryType";
  relabels.push(typeSelect.relabel);
  field(typeSelect.element, (l) => ({ label: say(l, "Typ", "Type") }), "field field-wide");

  // Year (or, for a child, the birth year itself) plus a read-only age
  // readout beside it -- its own composite row, the same reason
  // `savingAmountOrShare` (advanced.ts) hand-builds one for the IPS toggle:
  // the standard field shape has no room for a second, dependent readout.
  const yearRow = document.createElement("div");
  yearRow.className = "pgb-year";
  const yearField = number(0, { min: 1900, max: 2100, step: 1 }, (v) => {
    yearValue = v;
    updateAgeReadout();
  });
  yearField.element.dataset.setting = "pgbEntryYear";
  const ageReadout = document.createElement("span");
  ageReadout.className = "pgb-year-age";
  yearRow.append(yearField.element, ageReadout);
  const yearCaption = (l: Lang) => ({
    label: typeValue === TYPE_CHILD ? say(l, "Barnets födelseår", "Child's birth year") : say(l, "År", "Year"),
  });
  const relabelYear = field(yearRow, yearCaption, "field field-wide");
  const yearWrap = formBox.lastElementChild as HTMLElement;

  const slotSelect = select(SLOT_CHOICES, slotValue, (v) => {
    slotValue = v;
  });
  slotSelect.element.dataset.setting = "pgbEntryChildSlot";
  relabels.push(slotSelect.relabel);
  field(slotSelect.element, (l) => ({ label: say(l, "Vilket barn", "Which child") }), "field field-wide");
  const slotWrap = formBox.lastElementChild as HTMLElement;

  const amountField = number(0, { min: 0, max: 10_000_000, step: 100 }, (v) => {
    amountValue = v;
  });
  amountField.element.dataset.setting = "pgbEntryAmount";
  field(amountField.element, (l) => ({ label: say(l, "Belopp, kr", "Amount, kr") }));
  const amountWrap = formBox.lastElementChild as HTMLElement;

  const semesterField = number(1, { min: 1, max: MAX_SEMESTERS, step: 1 }, (v) => {
    semesterValue = v;
  });
  semesterField.element.dataset.setting = "pgbEntryStudySemesters";
  field(semesterField.element, (l) => ({ label: say(l, "Antal terminer", "Semesters") }));
  const semesterWrap = formBox.lastElementChild as HTMLElement;

  // Conscription's own from/to pair, plus a "too short" readout and a way to
  // clear a period that already exists -- carried over from the old grid's
  // header-cell widget, just no longer glued to the Värnplikt table column.
  const conscriptionRow = document.createElement("div");
  conscriptionRow.className = "pgb-conscription";
  const startLabel = document.createElement("span");
  const startInput = document.createElement("input");
  startInput.type = "date";
  startInput.min = "1995-01-01";
  startInput.dataset.setting = "pgbEntryConscriptionStart";
  const startWrap = document.createElement("label");
  startWrap.className = "pgb-conscription-date";
  startWrap.append(startLabel, startInput);

  const endLabel = document.createElement("span");
  const endInput = document.createElement("input");
  endInput.type = "date";
  endInput.dataset.setting = "pgbEntryConscriptionEnd";
  const endWrap = document.createElement("label");
  endWrap.className = "pgb-conscription-date";
  endWrap.append(endLabel, endInput);

  const clearBtn = document.createElement("button");
  clearBtn.type = "button";
  clearBtn.className = "pgb-conscription-clear";
  clearBtn.dataset.action = "pgb-clear-conscription";
  clearBtn.addEventListener("click", () => {
    conscription = undefined;
    startInput.value = "";
    endInput.value = "";
    updateConscriptionReadout();
    emit();
  });

  const conscriptionReadout = document.createElement("p");
  conscriptionReadout.className = "pgb-conscription-readout";
  conscriptionRow.append(startWrap, endWrap, clearBtn, conscriptionReadout);
  field(
    conscriptionRow,
    (l) => ({ label: say(l, "Värnplikt, period", "Conscription period") }),
    "field field-wide",
  );
  const conscriptionWrap = formBox.lastElementChild as HTMLElement;

  startInput.addEventListener("change", () => updateConscriptionReadout());
  endInput.addEventListener("change", () => updateConscriptionReadout());

  const addButton = document.createElement("button");
  addButton.type = "button";
  addButton.className = "export-btn pgb-add";
  addButton.dataset.action = "pgb-add-entry";
  addButton.addEventListener("click", onAdd);
  formBox.append(addButton);

  // Reports back whether the entry just added actually earned anything --
  // read from the run's own `pgbBreakdown` after `emit()`, not re-derived
  // eligibility rules, so it can never disagree with what the engine did.
  const addFeedback = document.createElement("p");
  addFeedback.className = "pgb-add-feedback";
  formBox.append(addFeedback);

  // ------------------------------------------------------------ children --
  const childrenBox = document.createElement("div");
  childrenBox.className = "pgb-children";
  const childrenHeading = document.createElement("h4");
  const childrenList = document.createElement("div");
  childrenList.className = "pgb-children-list";
  childrenBox.append(childrenHeading, childrenList);

  // --------------------------------------------------------------- table --
  const summaryBox = document.createElement("div");
  summaryBox.className = "pgb-summary scroll";
  const summaryHead = document.createElement("div");
  summaryHead.className = "pgb-summary-head";
  // Hidden in the sidebar, where the group's own <summary> already labels
  // the panel -- shown only once the table floats away from it, so the
  // floating panel still says what it is.
  const summaryTitle = document.createElement("span");
  summaryTitle.className = "pgb-summary-title";
  summaryTitle.hidden = true;
  const expandBtn = document.createElement("button");
  expandBtn.type = "button";
  expandBtn.className = "export-btn pgb-expand";
  expandBtn.dataset.action = "pgb-expand";
  expandBtn.addEventListener("click", () => setExpanded(!expanded));
  summaryHead.append(summaryTitle, expandBtn);
  const summaryTable = document.createElement("table");
  summaryTable.className = "table pgb-summary-table";
  const summaryThead = summaryTable.createTHead();
  const theadRow = summaryThead.insertRow();
  const summaryTbody = summaryTable.createTBody();
  const summaryEmpty = document.createElement("p");
  summaryEmpty.className = "pgb-summary-empty";
  summaryBox.append(summaryHead, summaryTable, summaryEmpty);

  body.append(intro, formBox, childrenBox, summaryBox);
  element.append(summary, body);

  // "Visa alla kolumner": unlike the old grid's own modal `<dialog>`, this
  // moves the same live table (same node, same remove-button listeners, not
  // a copy) into a plain floating panel instead -- no `showModal()`, so
  // nothing about the rest of the page (in particular the add-entry form
  // just above, still in its own place in the sidebar) is made inert. Asked
  // for by name after the redesign shipped: the whole point of adding one
  // entry at a time is to see the table update as you go, which a blocking
  // dialog that hides everything else defeats.
  let expanded = false;
  const drawer = document.createElement("div");
  drawer.className = "pgb-summary-drawer";
  drawer.hidden = true;
  document.body.append(drawer);

  function setExpanded(next: boolean): void {
    expanded = next;
    if (expanded) {
      drawer.append(summaryBox);
      drawer.hidden = false;
    } else {
      childrenBox.after(summaryBox);
      drawer.hidden = true;
    }
    applyExpandText(currentLang);
  }

  function applyExpandText(l: Lang): void {
    expandBtn.textContent = expanded
      ? say(l, "Dölj tabellen", "Hide the table")
      : say(l, "Visa alla kolumner", "Show all columns");
    expandBtn.setAttribute(
      "aria-label",
      expanded
        ? say(l, "Dölj den utökade PGB-tabellen", "Hide the expanded PGB table")
        : say(l, "Visa PGB-tabellen större, utan att stänga formuläret", "Show the PGB table larger, without closing the form"),
    );
    summaryTitle.hidden = !expanded;
    summaryTitle.textContent = say(l, "Pensionsgrundande belopp (PGB)", "Pension-qualifying amounts (PGB)");
  }

  function onTypeChange(): void {
    addFeedback.textContent = "";
    yearWrap.hidden = typeValue === TYPE_CONSCRIPTION;
    slotWrap.hidden = typeValue !== TYPE_CHILD;
    amountWrap.hidden = typeValue !== TYPE_SICKNESS;
    semesterWrap.hidden = typeValue !== TYPE_STUDY;
    conscriptionWrap.hidden = typeValue !== TYPE_CONSCRIPTION;
    relabelYear(currentLang);
    if (typeValue === TYPE_CONSCRIPTION) {
      startInput.value = conscription?.start ?? "";
      endInput.value = conscription?.end ?? "";
      updateConscriptionReadout();
    }
    refreshChildSlotOptions();
    refreshAddButtonState();
    updateAgeReadout();
  }

  function updateAgeReadout(): void {
    ageReadout.textContent = born > 0 && yearValue > 0 ? String(yearValue - Math.trunc(born)) : "";
  }

  // Only the one case the table itself cannot show: a period too short to
  // earn anything has no touched-year row to display a zero in.
  function updateConscriptionReadout(): void {
    clearBtn.hidden = conscription === undefined;
    if (!startInput.value || !endInput.value) {
      conscriptionReadout.textContent = "";
      return;
    }
    const days = conscriptionDaysByYear({ start: startInput.value, end: endInput.value });
    if (days.size === 0) {
      conscriptionReadout.textContent = say(
        currentLang,
        "Perioden är kortare än 120 dagar och ger ingen pensionsrätt.",
        "The period is under 120 days and earns no pension rights.",
      );
      return;
    }
    // Long enough is not the same as eligible: `wsPGB!F`'s own window is
    // 1995-2010 and from 2018 on, regardless of the day count.
    const eligible = [...days.keys()].some(conscriptionEligible);
    conscriptionReadout.textContent = eligible
      ? ""
      : say(
          currentLang,
          "Perioden ger ingen pensionsrätt -- värnplikt ger bara pensionsrätt 1995–2010 och från 2018.",
          "The period earns no pension rights -- conscription only earns pension rights 1995-2010 and from 2018 on.",
        );
  }

  function refreshChildSlotOptions(): void {
    const options = slotSelect.element.options;
    for (let i = 0; i < options.length; i += 1) options[i]!.disabled = (childBirthYears[i] ?? 0) !== 0;
    if ((childBirthYears[slotValue - 1] ?? 0) !== 0) {
      const free = childBirthYears.findIndex((y) => y === 0);
      if (free !== -1) {
        slotValue = free + 1;
        slotSelect.element.value = String(slotValue);
      }
    }
  }

  function refreshAddButtonState(): void {
    addButton.disabled = typeValue === TYPE_CHILD && childBirthYears.every((y) => y !== 0);
  }

  /** `age > 15 && age <= riktalder` -- `earnPgb`'s own gate on any manual
   * entry, riktålder defaulting to 66 (`LAST_AGE` leaves a small margin). A
   * year outside this is rejected outright rather than silently moved to
   * the nearest valid one, which would credit a different year than the one
   * actually typed with no sign anything had changed. */
  function ageInRange(year: number): number | undefined {
    if (born <= 0) return undefined;
    const age = year - Math.trunc(born);
    return age >= FIRST_AGE && age <= LAST_AGE ? age : undefined;
  }

  function onAdd(): void {
    addFeedback.textContent = "";
    let childYears: readonly number[] = [];
    let committedAge: number | undefined;

    switch (typeValue) {
      case TYPE_CHILD: {
        if (!Number.isFinite(yearValue) || yearValue <= 0) return;
        const slot = slotValue - 1;
        if ((childBirthYears[slot] ?? 0) !== 0) return;
        childBirthYears = [...childBirthYears] as [number, number, number, number];
        childBirthYears[slot] = yearValue;
        childYears = [0, 1, 2, 3].map((n) => yearValue + n);
        break;
      }
      case TYPE_CONSCRIPTION: {
        if (!startInput.value || !endInput.value) return;
        conscription = { start: startInput.value, end: endInput.value };
        break;
      }
      case TYPE_SICKNESS: {
        if (!Number.isFinite(yearValue) || yearValue <= 0 || amountValue <= 0) return;
        const age = ageInRange(yearValue);
        if (age === undefined) {
          addFeedback.textContent = say(
            currentLang,
            `Året motsvarar en ålder utanför ${FIRST_AGE}–${LAST_AGE} år, och ger ingen pensionsrätt.`,
            `That year is an age outside ${FIRST_AGE}-${LAST_AGE}, and earns no pension rights.`,
          );
          return;
        }
        committedAge = age;
        sicknessByAge.set(age, amountValue);
        break;
      }
      case TYPE_STUDY: {
        if (!Number.isFinite(yearValue) || yearValue <= 0 || semesterValue <= 0) return;
        const age = ageInRange(yearValue);
        if (age === undefined) {
          addFeedback.textContent = say(
            currentLang,
            `Året motsvarar en ålder utanför ${FIRST_AGE}–${LAST_AGE} år, och ger ingen pensionsrätt.`,
            `That year is an age outside ${FIRST_AGE}-${LAST_AGE}, and earns no pension rights.`,
          );
          return;
        }
        committedAge = age;
        studyByAge.set(age, semesterValue);
        break;
      }
    }
    if (typeValue !== TYPE_CONSCRIPTION) {
      yearValue = 0;
      yearField.setValue(0);
    }
    if (typeValue === TYPE_SICKNESS) {
      amountValue = 0;
      amountField.setValue(0);
    }
    if (typeValue === TYPE_STUDY) {
      semesterValue = 1;
      semesterField.setValue(1);
    }
    updateAgeReadout();
    refreshChildSlotOptions();
    refreshAddButtonState();
    drawChildrenList();
    emit();
    // `emit` -> `onChange` -> `render()` -> `setBaseline` is synchronous (no
    // `await` anywhere in that chain), so `breakdown` already reflects this
    // entry -- reading it back, rather than re-deriving eligibility here,
    // is the only check that can't disagree with what the engine actually
    // credited. The age range above already rules out sickness/study's one
    // real failure mode, so a still-zero figure here means the 7.5
    // inkomstbasbelopp income+PGB cap absorbed it, not a bad year.
    if (typeValue === TYPE_SICKNESS || typeValue === TYPE_STUDY) {
      if (committedAge === undefined) return;
      const row = breakdown.find((r) => r.age === committedAge);
      const earned = typeValue === TYPE_SICKNESS ? (row?.sa ?? 0) > 0 : (row?.studier ?? 0) > 0;
      if (!earned) {
        addFeedback.textContent = say(
          currentLang,
          "Beloppet gav ingen pensionsrätt det här året -- inkomsten är redan vid taket (7,5 inkomstbasbelopp).",
          "The amount earned no pension rights this year -- income is already at the cap (7.5 income base amounts).",
        );
      }
      return;
    }
    if (typeValue === TYPE_CHILD) {
      const earned = childYears.some((year) => {
        const age = born > 0 ? year - Math.trunc(born) : undefined;
        return age !== undefined && (breakdown.find((r) => r.age === age)?.barn ?? 0) > 0;
      });
      if (!earned) {
        addFeedback.textContent = say(
          currentLang,
          "Födelseåret gav ingen pensionsrätt -- du måste ha varit minst 16 år vid barnets födelse.",
          "The birth year earned no pension rights -- you must have been at least 16 at the child's birth.",
        );
      }
    }
  }

  function removeEntry(category: "sa" | "studier", age: number): void {
    if (category === "sa") sicknessByAge.delete(age);
    else studyByAge.delete(age);
    emit();
  }

  function removeChild(slot: number): void {
    childBirthYears = [...childBirthYears] as [number, number, number, number];
    childBirthYears[slot] = 0;
    refreshChildSlotOptions();
    refreshAddButtonState();
    drawChildrenList();
    emit();
  }

  function drawChildrenList(): void {
    const any = childBirthYears.some((y) => y !== 0);
    childrenBox.hidden = !any;
    childrenList.replaceChildren();
    childBirthYears.forEach((year, slot) => {
      if (year === 0) return;
      const row = document.createElement("div");
      row.className = "pgb-child-row";
      row.dataset.childSlot = String(slot + 1);

      const label = document.createElement("span");
      label.textContent = childOrdinal(slot, currentLang);

      const input = document.createElement("input");
      input.type = "number";
      input.inputMode = "numeric";
      input.min = "1900";
      input.max = "2100";
      input.step = "1";
      input.value = String(year);
      input.dataset.setting = `childBirthYear${slot + 1}`;
      input.addEventListener("change", () => {
        const typed = Math.round(Number(input.value));
        if (!Number.isFinite(typed) || typed <= 0) {
          input.value = String(childBirthYears[slot] ?? 0);
          return;
        }
        childBirthYears = [...childBirthYears] as [number, number, number, number];
        childBirthYears[slot] = typed;
        input.value = String(typed);
        emit();
      });

      const removeBtn = document.createElement("button");
      removeBtn.type = "button";
      removeBtn.className = "pgb-remove";
      removeBtn.dataset.action = "pgb-remove-child";
      removeBtn.dataset.childSlot = String(slot + 1);
      removeBtn.textContent = "×";
      removeBtn.setAttribute("aria-label", say(currentLang, "Ta bort barn", "Remove child"));
      removeBtn.addEventListener("click", () => removeChild(slot));

      row.append(label, input, removeBtn);
      childrenList.append(row);
    });
  }

  function summaryRows(): SummaryRow[] {
    return breakdown
      .filter((r) => r.sa !== 0 || r.vpl !== 0 || r.studier !== 0 || r.barn !== 0)
      .map((r) => ({
        age: r.age,
        year: born > 0 ? Math.trunc(born) + r.age : r.age,
        sa: r.sa,
        vpl: r.vpl,
        studier: r.studier,
        barn: r.barn,
        total: r.sa + r.vpl + r.studier + r.barn,
      }))
      .sort((a, b) => a.age - b.age);
  }

  function drawSummary(): void {
    const rows = summaryRows();
    summaryEmpty.hidden = rows.length > 0;
    summaryTable.hidden = rows.length === 0;
    expandBtn.hidden = rows.length === 0;
    if (rows.length === 0 && expanded) setExpanded(false);
    theadRow.replaceChildren();
    summaryTbody.replaceChildren();
    if (rows.length === 0) return;

    const visibleColumns = SUMMARY_COLUMNS.filter((col) => rows.some((r) => col.get(r) > 0));
    theadRow.append(
      headCell(t("year", currentLang)),
      headCell(t("age", currentLang)),
      ...visibleColumns.map((col) => headCell(col.head(currentLang), col.info?.(currentLang))),
      headCell(say(currentLang, "Summa PGB, kr", "Total PGB, kr")),
    );

    for (const r of rows) {
      const tr = document.createElement("tr");
      tr.dataset.year = String(r.year);
      tr.append(cell(String(r.year)), cell(String(r.age)));
      for (const col of visibleColumns) {
        const value = col.get(r);
        const td = cell(value > 0 ? kronor(value, currentLang) : "", "num");
        if (value > 0 && (col.key === "sa" || col.key === "studier")) {
          const category = col.key;
          const removeBtn = document.createElement("button");
          removeBtn.type = "button";
          removeBtn.className = "pgb-remove";
          removeBtn.dataset.action = "pgb-remove";
          removeBtn.dataset.category = category;
          removeBtn.dataset.year = String(r.year);
          removeBtn.textContent = "×";
          removeBtn.setAttribute("aria-label", say(currentLang, "Ta bort", "Remove"));
          removeBtn.addEventListener("click", () => removeEntry(category, r.age));
          td.append(removeBtn);
        }
        tr.append(td);
      }
      tr.append(cell(kronor(r.total, currentLang), "num strong"));
      summaryTbody.append(tr);
    }
  }

  function emit(): void {
    const ages = new Set([...sicknessByAge.keys(), ...studyByAge.keys()]);
    const pgbManual: PgbManualYear[] = [...ages].map((age) => ({
      age,
      sa: sicknessByAge.get(age) ?? 0,
      studySemesters: studyByAge.get(age) ?? 0,
    }));
    onChange({
      pgbManual: pgbManual.length > 0 ? pgbManual : undefined,
      pgbConscription: conscription,
      childBirthYears,
    });
  }

  let currentLang = lang;

  const applyText = (l: Lang) => {
    const title = say(l, "Pensionsgrundande belopp (PGB)", "Pension-qualifying amounts (PGB)");
    summary.textContent = title;
    intro.textContent = say(
      l,
      "Barnår, sjuk- eller aktivitetsersättning, värnplikt och studier ger alla pensionsrätt " +
        "utöver den vanliga inkomsten. Välj typ nedan, fyll i det som gäller och klicka Lägg till.",
      "Child years, sickness or activity compensation, conscription and study all earn pension " +
        "rights on top of ordinary income. Pick a type below, fill in what it needs and click Add.",
    );
    for (const fn of relabels) fn(l);
    relabelYear(l);
    startLabel.textContent = say(l, "Från", "From");
    startInput.setAttribute("aria-label", say(l, "Värnplikt, startdatum", "Conscription, start date"));
    endLabel.textContent = say(l, "Till", "To");
    endInput.setAttribute("aria-label", say(l, "Muck (slutdatum)", "End date"));
    clearBtn.textContent = say(l, "Rensa", "Clear");
    clearBtn.setAttribute("aria-label", say(l, "Rensa värnpliktsperiod", "Clear the conscription period"));
    addButton.textContent = say(l, "Lägg till", "Add");
    childrenHeading.textContent = say(l, "Barn", "Children");
    summaryEmpty.textContent = say(l, "Inga poster ännu.", "No entries yet.");
    applyExpandText(l);
    updateConscriptionReadout();
    drawChildrenList();
    drawSummary();
  };
  applyText(lang);
  onTypeChange();

  return {
    element,
    relabel(l) {
      currentLang = l;
      applyText(l);
      onTypeChange();
    },
    setBaseline(bornYear, marginalValue, pgbBreakdown) {
      // `pgbBreakdown` is a handful of entries at most (only touched ages),
      // so a stringified comparison is cheap -- and it is what actually
      // decides whether the summary table needs to change, since the run's
      // own economic assumptions can move conscription's kronor without
      // `born` or `marginal` moving at all.
      const digest = JSON.stringify(pgbBreakdown);
      if (bornYear === born && marginalValue === marginal && digest === breakdownDigest) return;
      born = bornYear;
      marginal = marginalValue;
      breakdown = pgbBreakdown;
      breakdownDigest = digest;
      updateAgeReadout();
      drawSummary();
    },
    reset() {
      sicknessByAge = new Map();
      studyByAge = new Map();
      conscription = undefined;
      childBirthYears = [...defaultContext().childBirthYears] as [number, number, number, number];
      typeValue = TYPE_CHILD;
      typeSelect.element.value = String(TYPE_CHILD);
      yearValue = 0;
      yearField.setValue(0);
      slotValue = 1;
      slotSelect.element.value = "1";
      amountValue = 0;
      amountField.setValue(0);
      semesterValue = 1;
      semesterField.setValue(1);
      startInput.value = "";
      endInput.value = "";
      breakdown = [];
      breakdownDigest = "";
      onTypeChange();
      drawChildrenList();
      drawSummary();
    },
  };
}
