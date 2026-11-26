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
 * Periods apply whether or not the grid below is switched on, and either way
 * the numbers follow the salary and the birth year as they are typed. Only a
 * row somebody has actually typed into stops following -- see `handEdited`
 * below, which is the whole of the rule. Overlapping periods resolve
 * last-wins, so a narrow period added after a wide one carves out of it.
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

/**
 * What `savedCases.ts` has to keep for this section, beyond the income path.
 *
 * The path itself rides in `TypfallInput.ownIncome` and is saved with the rest
 * of Avancerat; these two are not recoverable from it. A Deltid period list
 * cannot be read back out of the numbers it produced, and whether the grid was
 * open decides whether those numbers are authoritative or get regenerated.
 */
export interface SalaryPathState {
  /** Whether `Använd egen löneutveckling` was on. */
  readonly grid: boolean;
  /**
   * Whether a grid cell had been typed into. Saved rather than recomputed:
   * it is what decides whether a restored case's rows are its own numbers or
   * a view that the next salary change may refill, and the rows alone cannot
   * say which they were.
   */
  readonly edited: boolean;
  readonly periods: readonly PartTimePeriod[];
}

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
   * Takes the wage path the model derives for the Start sheet as it currently
   * stands -- ignoring anything typed here -- and answers with the path the
   * run should actually be given.
   *
   * It both feeds and answers because the answer is an *input* to that run, so
   * `main.ts` settles it before running rather than being called back
   * afterwards. It used to be called back: `setBaseline` pushed a new path
   * through `onChange`, which re-entered `render`, which finished and returned
   * into the outer `render` -- which then overwrote the corrected results with
   * its own, computed moments earlier from the path that had just been
   * replaced. Every figure on the page sat one edit behind, which is what
   * "change the salary and nothing updates" turned out to be.
   *
   * Keeping the baseline fresh is the other half of its job: the grid fills
   * from it, `Återställ` goes back to it, and an untouched grid follows it.
   */
  sync(path: readonly OwnIncomeYear[], born: number): readonly OwnIncomeYear[] | undefined;
  /** Switches the grid off and drops the typed path. */
  reset(): void;
  /** The part of this section a saved typfall has to carry itself. */
  state(): SalaryPathState;
  /**
   * Puts a saved typfall back on screen.
   *
   * `path` is that case's own `ownIncome`, which fills the grid when the case
   * had it open. Nothing is handed up: `main.ts` has set its own state
   * already, and the `render` it does next asks `sync` for the path anyway --
   * which is where a case saved with the grid untouched, or with periods and
   * no grid, gets its numbers rebuilt against the restored salary.
   */
  restore(
    state: SalaryPathState,
    path: readonly OwnIncomeYear[] | undefined,
    context: ModelContext,
  ): void;
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
  /**
   * Whether a grid cell has been typed into (or zeroed in one click) since the
   * grid was last filled.
   *
   * This is what decides whether the rows are a *view* of the derived path or
   * a replacement for it, and switching the grid on is deliberately not enough
   * to make them the latter. Reported as a bug, and it was one: tick the box
   * to look at the numbers, change the salary, and Table 2 silently kept
   * answering for the old one, because the rows it was still being handed had
   * been frozen at the moment of the tick. An untouched grid now refills from
   * the baseline like everything else here, and only a typed cell takes over --
   * which is the case the "only Återställ overwrites them" rule existed to
   * protect in the first place.
   */
  let handEdited = false;
  const relabels: ((l: Lang) => void)[] = [];
  const restores: (() => void)[] = [];
  // The three Adv_settings this file renders rather than `advanced.ts`'s own
  // loop, so `AdvancedHandle.apply` cannot reach them -- a restored typfall
  // sets them from here instead. See `restore` below.
  const settingSetters: ((c: ModelContext) => void)[] = [];
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
    settingSetters.push((c) => control.setValue(Math.min(Math.max(setting.get(c), min), max)));
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
    const choices = WAGE_PROFILE.control.choices;
    const control = select(choices, initial, (v) => onSettingChange(WAGE_PROFILE.set(v)));
    control.element.dataset.setting = WAGE_PROFILE.key;
    settingSetters.push((c) => {
      const value = WAGE_PROFILE.get(c);
      control.element.value = String(choices.some((choice) => choice.value === value) ? value : initial);
    });
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
          // Before `push`, which re-enters `render` and so `setBaseline`:
          // the flag has to be true by the time that runs, or it would
          // refill the grid out from under the cell being edited.
          handEdited = true;
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
   * The path the run should be given as things stand.
   *
   * A typed grid is the authority; failing that, the periods generate one; and
   * with neither, nothing -- which is how the model is told to derive the wage
   * path itself, rather than to read a lifetime of zeroes.
   */
  function current(): readonly OwnIncomeYear[] | undefined {
    return toggle.checked ? rows : periods.length > 0 ? fill() : undefined;
  }

  /**
   * A click or a keystroke in this section changed the path: hand it up.
   *
   * Only ever called from a user's own event, never from `sync` -- which is
   * what keeps `onChange`, and so `render`, out of its own call stack.
   */
  function push(): void {
    onChange(current());
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
      handEdited = false;
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
      handEdited = false;
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
    handEdited = false;
    drawGrid();
    push();
  });

  // A full year of unpaid leave typed one row at a time is tedious at this
  // range's own length (age 15 through the last worked age) -- this zeroes
  // every row in one click, the same way "Använd normala inställningar"
  // zeroes a whole setting but scoped to just this grid.
  zeroAll.addEventListener("click", () => {
    rows = rows.map((row) => ({ ...row, income: 0, wage: 0 }));
    handEdited = true;
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
    sync(path, bornYear) {
      baseline = path;
      if (bornYear !== born) {
        born = bornYear;
        drawPartTimeAges();
      }

      // Everything here follows the Start sheet -- the path generated from the
      // periods while the grid is off, and the grid's own rows while nobody
      // has typed into them -- so that changing the salary or the birth year
      // moves the numbers, which is what anyone would expect of a figure the
      // page recomputes on every keystroke. A typed cell is the one exception,
      // and then only until `Återställ`.
      //
      // Redrawing unconditionally is what must not happen: it would tear the
      // cell being edited out of the DOM inside its own change event, which an
      // earlier offline check caught as a run of `replaceChildren`
      // NotFoundErrors. `handEdited` is set before that cell's own `push`, so
      // the branch below is already closed by the time this runs for it; the
      // year column is relabelled in place for a typed grid instead.
      if (toggle.checked && !handEdited) {
        const next = fill();
        if (JSON.stringify(next) !== JSON.stringify(rows)) {
          rows = next;
          drawGrid();
        }
      } else if (toggle.checked) {
        for (const [index, tr] of [...tbody.children].entries()) {
          const row = rows[index];
          if (row === undefined) continue;
          tr.firstElementChild!.textContent = born > 0 ? String(Math.trunc(born) + row.age) : "";
        }
      }

      return current();
    },
    state() {
      return {
        grid: toggle.checked,
        edited: handEdited,
        periods: periods.map((p) => ({ ...p })),
      };
    },
    restore(state, path, context) {
      for (const set of settingSetters) set(context);
      periods = state.periods.map((p) => ({ ...p }));
      toggle.checked = state.grid;
      scroll.hidden = !state.grid;
      actions.hidden = !state.grid;
      rows = state.grid ? (path ?? []).map((r) => ({ ...r })) : [];
      handEdited = state.grid && state.edited;
      if (state.grid) drawGrid();
      partTimeNote.textContent = "";
      drawPeriods();
      drawPartTimeAges();
    },
    reset() {
      toggle.checked = false;
      scroll.hidden = true;
      actions.hidden = true;
      rows = [];
      handEdited = false;
      periods = [];
      partTimeNote.textContent = "";
      drawPeriods();
      for (const r of restores) r();
    },
  };
}
