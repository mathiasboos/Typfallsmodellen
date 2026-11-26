/**
 * "Salary" -- the Indata_lista sheet as a grid, renamed from "Own salary
 * path" and, on request, now also the home for manual 3.6's "Slutlön" pair
 * (`FINAL_SALARY_YEARS`/`PENSION_SAME_YEAR_AS_FINAL_SALARY`) and row 141's
 * "Löneprofil" (`WAGE_PROFILE`), all three imported from `advanced.ts`: each
 * describes the derived wage path this file's own typed path also feeds or
 * overrides, so they render here, above the path toggle, using the same
 * `fieldSet` builders `advanced.ts`'s own loop uses for a plain number or
 * select setting -- rather than in a `<details>` of their own (that used to
 * be `advanced.ts`'s "Övrigt"/"Other" group, now removed since nothing was
 * left in it once this and `flexPension` moved out). A second
 * `onSettingChange` callback carries their own `ModelContext` patches up to
 * `main.ts`'s `advanced` state, alongside the existing `onChange` for the
 * wage path itself (a `TypfallInput` field) -- the same two-callbacks-one-
 * component shape `pgb.ts`'s own `childBirthYears` already established.
 *
 * Manual section 3.1: tick the box and the model stops deriving the wage path
 * and reads a typed one instead, one row per age with two columns, read
 * independently of each other -- neither is derived from the other once this
 * is on. `Inkomst` is taxed earned income: it is what PGI (pension-qualifying
 * income) is computed from, and PGI is what actually earns income-pension,
 * premium-pension and guarantee-pension rights -- `ipavgift`/`ppavgift`/
 * `gpavgift` all read PGI, never `v.wage` directly. `Inkomst` is also the
 * final salary Table 1 reports, private saving entered as a share of income,
 * and the tax base `taxes()` reads. `Varav lön` is the narrower, employer-
 * paid part of it: what an occupational scheme's own premium (`premiumFor`)
 * and the employer-contribution reduction (`arbgiv`) are computed from, and
 * what a benefit or tax calculation subtracts back out of gross income to
 * isolate salary from everything else taxable. Zeroing a year is how a year
 * of leave is simulated; halving a run of years is how part-time is; typing
 * `Inkomst` with `Varav lön` left at 0 is how taxable income that still
 * earns public pension rights but no occupational scheme's own premium
 * (some taxable benefits) is.
 *
 * "Deltid" is that halving made a control rather than a chore. A part-time
 * period is a range of calendar years and a share of full time, any number of
 * them can be listed, and `applyPartTime` scales both `Inkomst` and `Varav
 * lön` of every year a period covers. Nothing downstream had to change for
 * it: reduced earnings reach PGI, the occupational premium, the employer-
 * contribution reduction, income tax and the housing supplement through the
 * very same own-income path a typed grid uses, so the engine is untouched and
 * the golden cases with it. Adv_settings row 84 (`rngLönPartUttag`) is the
 * workbook's own part-time switch, but a far narrower one: it governs only
 * the window between the first, partial public-pension withdrawal and the
 * final one, and its shipped value already ties work to the share drawn
 * (`withdrawalShare`, packages/engine/src/income/wages.ts), so it cannot
 * express a period of its own at all. Its own sheet comment -- "obs om egen
 * lönevektor antas de gälla" -- says an own wage vector wins over it, which
 * is exactly how these periods are built.
 *
 * Periods apply whether or not the grid below is switched on. With it off the
 * generated path follows the salary and the birth year as they are typed,
 * which a hand-typed grid deliberately does not; switching the grid on freezes
 * those numbers into editable rows, and from then on the rows are the
 * authority until a period is added or removed, or `Återställ` is pressed.
 * Overlapping periods resolve last-wins, so a narrow period added after a wide
 * one carves out of it.
 *
 * The grid is filled from the path the model just computed rather than left
 * empty, which is what the workbook does too -- `setup.ts:424` reads an age the
 * array does not mention as 0, so an empty grid would mean a lifetime of no
 * income, not "derive it for me". `TypfallResult.wagePath` is that fill, shaped
 * so it can be handed straight back.
 *
 * Ages below 15 are not shown: `validate` in setup.ts pins `startage` to 15
 * whenever an own path is present, so those rows could never be read.
 *
 * Amounts are rounded to whole kronor, both shown and used, keeping form.ts's
 * rule that the form never shows a number the run did not use. Measured cost of
 * the rounding against the unrounded path: 0.02 kr per month on the final
 * salary, and less on everything else.
 */
import { defaultContext } from "@typfallsmodellen/engine";
import type { ModelContext, OwnIncomeYear } from "@typfallsmodellen/engine";

import { FINAL_SALARY_YEARS, PENSION_SAME_YEAR_AS_FINAL_SALARY, WAGE_PROFILE } from "./advanced.js";
import { fieldSet } from "./controls.js";
import { locale } from "./format.js";
import type { Lang } from "./i18n.js";

/** The first age an own path is ever read at -- `setup.ts:139`. */
const FIRST_AGE = 15;

/** One stretch of part-time work: the calendar years, and the share worked. */
export interface PartTimePeriod {
  /** First calendar year worked part-time, inclusive. */
  readonly fromYear: number;
  /** Last calendar year worked part-time, inclusive. */
  readonly toYear: number;
  /** Share of full time, as a fraction -- 0.5 is half time. */
  readonly share: number;
}

/**
 * `baseline`, trimmed to the ages an own path is read at, with every year a
 * period covers scaled to that period's share, and rounded.
 *
 * A row's calendar year is `trunc(born) + age`, the same arithmetic the grid's
 * own Year column shows, so the years the period list names and the years the
 * grid shows are the same ones. Periods apply in list order and the last one
 * covering a year wins, which is what lets a narrow period be added on top of
 * a wide one. The rounding happens once, after scaling, so half of a rounded
 * figure is never rounded a second time.
 */
export function applyPartTime(
  baseline: readonly OwnIncomeYear[],
  born: number,
  periods: readonly PartTimePeriod[],
): OwnIncomeYear[] {
  return baseline
    .filter((r) => r.age >= FIRST_AGE)
    .map((r) => {
      const year = Math.trunc(born) + r.age;
      let share = 1;
      for (const period of periods) {
        if (year >= period.fromYear && year <= period.toYear) share = period.share;
      }
      return {
        age: r.age,
        income: Math.round(r.income * share),
        wage: Math.round(r.wage * share),
      };
    });
}

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
  onSettingChange: (patch: Partial<ModelContext>) => void,
): SalaryPathHandle {
  let currentLang = lang;
  let baseline: readonly OwnIncomeYear[] = [];
  let born = 0;
  let rows: OwnIncomeYear[] = [];
  let periods: PartTimePeriod[] = [];
  /** `JSON.stringify` of the path last handed up -- see `push` below. */
  let pushed = "";
  const relabels: ((l: Lang) => void)[] = [];
  const restores: (() => void)[] = [];
  const normal = defaultContext();

  const element = document.createElement("details");
  element.className = "adv-group";
  element.dataset.group = "salary-path";

  const summary = document.createElement("summary");
  const body = document.createElement("div");
  body.className = "adv-body";

  // The "Slutlön" pair (see this file's own header comment) -- plain number
  // fields, rendered with the same builder `advanced.ts`'s own loop uses for
  // a setting of this `control.kind`, since neither needs that loop's other
  // branches (percent, select, check, the ipsMonthly special case).
  const { field, number, select } = fieldSet(body, relabels, lang);
  for (const setting of [FINAL_SALARY_YEARS, PENSION_SAME_YEAR_AS_FINAL_SALARY]) {
    if (setting.control.kind !== "number") continue; // both are; guards the destructure below
    const initial = setting.get(normal);
    const { min, max, step } = setting.control;
    const control = number(initial, { min, max, step }, (v) => onSettingChange(setting.set(v)));
    control.element.dataset.setting = setting.key;
    field(control.element, (l) => ({
      label: setting.label(l),
      ...(setting.hint ? { hint: setting.hint(l) } : {}),
    }));
    restores.push(() => control.setValue(initial));
  }

  // "Löneprofil" -- a select, not a number, so it sits outside the loop
  // above rather than widening that loop's own single-kind guard.
  if (WAGE_PROFILE.control.kind === "select") {
    const initial = WAGE_PROFILE.get(normal);
    const control = select(WAGE_PROFILE.control.choices, initial, (v) => onSettingChange(WAGE_PROFILE.set(v)));
    control.element.dataset.setting = WAGE_PROFILE.key;
    relabels.push(control.relabel);
    field(
      control.element,
      (l) => ({
        label: WAGE_PROFILE.label(l),
        ...(WAGE_PROFILE.hint ? { hint: WAGE_PROFILE.hint(l) } : {}),
      }),
      "field field-wide",
    );
    restores.push(() => {
      control.element.value = String(initial);
    });
  }

  // ------------------------------------------------------------- Deltid --
  // Its own box and its own `fieldSet`, bound to that box rather than to
  // `body`: the three fields are one add-entry form -- the shape pgb.ts
  // settled on, down to editing being "remove it and add it again" -- not
  // three more settings in this group's own single column.
  const partTimeBox = document.createElement("div");
  partTimeBox.className = "salary-parttime";
  const partTimeHeading = document.createElement("h4");
  const partTimeIntro = document.createElement("p");
  partTimeIntro.className = "pgb-intro";
  const partTimeForm = document.createElement("div");
  partTimeForm.className = "pgb-form";
  partTimeBox.append(partTimeHeading, partTimeIntro, partTimeForm);

  const thisYear = new Date().getFullYear();
  let fromYear = thisYear;
  let toYear = thisYear;
  let share = 0.5;

  const partTime = fieldSet(partTimeForm, relabels, lang);
  const fromField = partTime.number(fromYear, { min: 1900, max: 2100, step: 1 }, (v) => {
    fromYear = v;
    drawPartTimeAges();
  });
  fromField.element.dataset.setting = "partTimeFrom";
  partTime.field(fromField.element, (l) => ({ label: say(l, "Från år", "From year") }));

  const toField = partTime.number(toYear, { min: 1900, max: 2100, step: 1 }, (v) => {
    toYear = v;
    drawPartTimeAges();
  });
  toField.element.dataset.setting = "partTimeTo";
  partTime.field(toField.element, (l) => ({ label: say(l, "Till år", "To year") }));

  // A percent field, not a number one: it reads and writes the fraction the
  // period actually stores, and renders it with the page's own decimal mark
  // rather than the browser's -- see controls.ts's own note on why.
  const shareField = partTime.percent(share, (v) => {
    // Clamped and written back rather than only clamped on the way in: the
    // form never shows a number the run did not use (form.ts's own rule,
    // which `fieldSet`'s `number` already keeps for a bounded number field --
    // `percent` has no bounds of its own to keep it with).
    share = Math.min(Math.max(v, 0), 1);
    if (share !== v) shareField.setValue(share);
  });
  shareField.element.dataset.setting = "partTimeShare";
  relabels.push(shareField.relabel);
  partTime.field(shareField.element, (l) => ({
    label: say(l, "Sysselsättningsgrad", "Share of full time"),
    hint: say(l, "t.ex. 50 % för halvtid", "e.g. 50% for half time"),
  }));

  // Which ages the typed years come to for this typfall, so a period can be
  // aimed without counting from the birth year by hand.
  const partTimeAges = document.createElement("p");
  partTimeAges.className = "salary-parttime-ages";
  const addPeriod = document.createElement("button");
  addPeriod.type = "button";
  addPeriod.className = "export-btn pgb-add";
  addPeriod.dataset.action = "salary-add-parttime";
  // Reports back a period that changed nothing -- read from the path the grid
  // would now be filled with, not from re-derived working-age rules, so it
  // can never disagree with what the model was handed.
  const partTimeNote = document.createElement("p");
  partTimeNote.className = "pgb-add-feedback";
  partTimeForm.append(partTimeAges, addPeriod, partTimeNote);

  const partTimeList = document.createElement("div");
  partTimeList.className = "pgb-children-list salary-parttime-list";
  partTimeBox.append(partTimeList);

  // "Använd normala inställningar" puts the add-entry form back where it
  // opened too, the same as every other advanced-mode field.
  restores.push(() => {
    fromYear = thisYear;
    toYear = thisYear;
    share = 0.5;
    fromField.setValue(fromYear);
    toField.setValue(toYear);
    shareField.setValue(share);
    drawPartTimeAges();
  });

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
  refill.dataset.action = "salary-refill";
  const zeroAll = document.createElement("button");
  zeroAll.type = "button";
  zeroAll.className = "panel-btn";
  zeroAll.dataset.action = "salary-zero";
  actions.append(refill, zeroAll);

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

  body.append(partTimeBox, toggleWrap, actions, scroll);
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
          push();
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

  /** The baseline, trimmed, scaled by the part-time periods and rounded. */
  function fill(): OwnIncomeYear[] {
    return applyPartTime(baseline, born, periods);
  }

  /**
   * Hand the current path up, but only when it actually moved.
   *
   * `setBaseline` calls this on every render, so that a generated path (one or
   * more periods, grid off) keeps up with the salary and the birth year as
   * they are typed -- and `onChange` re-enters `render`. The digest is what
   * makes that terminate: the second pass derives the same path from the same
   * baseline and stops here. The baseline itself is a run with `ownIncome`
   * removed, so pushing a path never moves it.
   */
  function push(): void {
    const path = toggle.checked ? rows : periods.length > 0 ? fill() : undefined;
    const digest = path === undefined ? "" : JSON.stringify(path);
    if (digest === pushed) return;
    pushed = digest;
    onChange(path);
  }

  /** What the typed years come to in ages, for the typfall on screen. */
  function drawPartTimeAges(): void {
    if (born <= 0) {
      partTimeAges.textContent = "";
      return;
    }
    const first = fromYear - Math.trunc(born);
    const last = toYear - Math.trunc(born);
    partTimeAges.textContent = say(currentLang, `ålder ${first}–${last}`, `age ${first}–${last}`);
  }

  function drawPeriods(): void {
    partTimeList.replaceChildren();
    const l = currentLang;
    if (periods.length === 0) {
      const empty = document.createElement("p");
      // Its own class, not pgb.ts's `.pgb-summary-empty`: both panels are on
      // screen at once, and the offline check addresses that one unscoped.
      empty.className = "salary-parttime-empty";
      empty.textContent = say(l, "Inga deltidsperioder", "No part-time periods");
      partTimeList.append(empty);
      return;
    }
    for (const [index, period] of periods.entries()) {
      const row = document.createElement("div");
      row.className = "pgb-child-row";
      const label = document.createElement("span");
      const percent = new Intl.NumberFormat(locale(l), { maximumFractionDigits: 1 }).format(
        Math.round(period.share * 1000) / 10,
      );
      label.textContent = say(
        l,
        `${period.fromYear}–${period.toYear}: ${percent} % av heltid`,
        `${period.fromYear}–${period.toYear}: ${percent}% of full time`,
      );
      const remove = document.createElement("button");
      remove.type = "button";
      remove.className = "pgb-remove";
      remove.dataset.action = "salary-remove-parttime";
      remove.textContent = "×";
      remove.title = say(l, "Ta bort", "Remove");
      remove.addEventListener("click", () => {
        periods.splice(index, 1);
        partTimeNote.textContent = "";
        drawPeriods();
        regenerate();
      });
      row.append(label, remove);
      partTimeList.append(row);
    }
  }

  /**
   * A period was added or removed: the grid, when it is on, is refilled from
   * the baseline rather than scaled in place -- scaling what is already there
   * would compound the moment a period's share was changed twice. It is the
   * same refill `Återställ` does, and it does cost hand edits, which the
   * toggle's own hint says.
   */
  function regenerate(): void {
    if (toggle.checked) {
      rows = fill();
      drawGrid();
    }
    push();
  }

  addPeriod.addEventListener("click", () => {
    const l = currentLang;
    if (toYear < fromYear) {
      partTimeNote.textContent = say(
        l,
        "Till år kan inte vara före från år.",
        "The to year cannot be before the from year.",
      );
      return;
    }
    const before = JSON.stringify(fill());
    periods.push({ fromYear, toYear, share });
    drawPeriods();
    regenerate();
    partTimeNote.textContent =
      JSON.stringify(fill()) === before
        ? say(
            l,
            "Perioden ändrar ingenting: den ligger utanför de år typfallet arbetar, eller är på heltid.",
            "The period changes nothing: it falls outside the years this typfall works, or it is full time.",
          )
        : "";
  });

  toggle.addEventListener("change", () => {
    if (toggle.checked) {
      rows = fill();
      drawGrid();
    }
    scroll.hidden = !toggle.checked;
    actions.hidden = !toggle.checked;
    // Switching it off does not drop the periods: they keep scaling the
    // derived path, which is the half of this section that needs no grid.
    push();
  });

  refill.addEventListener("click", () => {
    rows = fill();
    drawGrid();
    push();
  });

  // A full year of unpaid leave typed one row at a time is tedious at this
  // range's own length (age 15 through the last worked age) -- this zeroes
  // every row in one click, the same way "Använd normala inställningar"
  // zeroes a whole setting but scoped to just this grid.
  zeroAll.addEventListener("click", () => {
    rows = rows.map((row) => ({ ...row, income: 0, wage: 0 }));
    drawGrid();
    push();
  });

  scroll.hidden = true;
  actions.hidden = true;

  const applyText = (l: Lang) => {
    summary.textContent = say(l, "Lön", "Salary");
    toggleLabel.textContent = say(l, "Använd egen löneutveckling", "Use an own salary path");
    toggleHint.textContent = say(
      l,
      "Fylls i från den beräknade lönebanan, deltidsperioderna inräknade, och kan sedan ändras fritt",
      "Filled in from the computed path, part-time periods included, which you can then edit",
    );
    partTimeHeading.textContent = say(l, "Deltid", "Part-time work");
    partTimeIntro.textContent = say(
      l,
      "Lägg till en eller flera perioder på deltid, med eller utan egen lönebana nedan. Inkomsten och lönen sänks till den angivna sysselsättningsgraden, vilket sänker pensionsrätterna och tjänstepensionens premier med dem. Överlappar två perioder gäller den som står sist.",
      "Add one or more periods of part-time work, with or without the own salary path below. Income and salary drop to the share of full time you give, and the pension rights and occupational premiums drop with them. Where two periods overlap, the lower one wins.",
    );
    addPeriod.textContent = say(l, "Lägg till period", "Add period");
    refill.textContent = say(l, "Återställ till beräknad lönebana", "Reset to the computed path");
    zeroAll.textContent = say(l, "Nollställ alla värden", "Set all values to 0");
    const heads = [
      say(l, "År", "Year"),
      say(l, "Ålder", "Age"),
      say(l, "Inkomst", "Income"),
      say(l, "Varav lön", "Of which salary"),
    ];
    for (const [i, cell] of headCells.entries()) cell.textContent = heads[i]!;
    drawPartTimeAges();
    drawPeriods();
  };
  applyText(lang);

  return {
    element,
    relabel(l) {
      currentLang = l;
      applyText(currentLang);
      for (const r of relabels) r(l);
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
      if (bornYear !== born) {
        born = bornYear;
        drawPartTimeAges();
        if (toggle.checked) {
          for (const [index, tr] of [...tbody.children].entries()) {
            const row = rows[index];
            if (row === undefined) continue;
            tr.firstElementChild!.textContent = born > 0 ? String(Math.trunc(born) + row.age) : "";
          }
        }
      }
      // With the grid off, a path generated from the periods has to follow the
      // baseline it was scaled from, or typing a new salary would leave the
      // part-time years at the old one's. `push` calls back only when the
      // numbers really moved, which is what keeps this out of a render loop;
      // with the grid on, `rows` has not moved and this is a no-op.
      push();
    },
    reset() {
      toggle.checked = false;
      scroll.hidden = true;
      actions.hidden = true;
      rows = [];
      periods = [];
      // `main.ts` drops its own `ownIncome` in the same click and re-renders,
      // so nothing is handed up here -- but the digest has to agree with that,
      // or the next `push` would think the empty path was already sent.
      pushed = "";
      partTimeNote.textContent = "";
      drawPeriods();
      for (const r of restores) r();
    },
  };
}
