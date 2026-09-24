/**
 * Mikrosim: the workbook's own batch runner, reproduced as a tab.
 *
 * The original sheet holds one independent pension case ("typfall") per row --
 * 10 input columns, a "Beräkna" button that runs the model over every row, and
 * 12 output columns fill in (`reference/golden/HOWTO.md` documents it at
 * length, since it's also what this project's own golden-file export drives).
 * This reproduces that model closely, minus one input column:
 *
 * - Every row is a fully independent `TypfallInput`/`ModelContext` pair, never
 *   a diff against the form on the left the way "Jämför scenarier"'s variants
 *   are -- a row doesn't need to know what the current form contains, which is
 *   what makes CSV import make sense here (a file doesn't carry a baseline).
 * - "Egen Lönelista" (a whole per-age income table, the same shape as the
 *   Advanced-mode salary path) is out of scope: it doesn't fit one flat batch
 *   row or one CSV cell. Every row uses the standard wage-growth model.
 * - Results are never live-recomputed. "Beräkna" is explicit, on purpose, not
 *   as a performance workaround (`run()` is sub-millisecond even for hundreds
 *   of rows) but because it's what the sheet itself calls this, because it
 *   cleanly separates "editing a batch" from "viewing its results" -- which
 *   matters once CSV import can load a hundred rows in one action -- and
 *   because it needs no extra state: a row's output cells are populated iff
 *   `row.result` is set, editing any input cell clears that row's own
 *   `result`/`error` straight back to blank, and "Beräkna" (re)computes every
 *   row unconditionally, the same way the sheet's own batch runner reruns its
 *   whole range each time rather than tracking which rows are dirty.
 *
 * No `onChange` callback up to `main.ts`, unlike `compare.ts` -- nothing
 * outside this panel depends on a Mikrosim row's contents, so there is
 * nothing to notify. `setContext` only stashes the shared `ModelContext`/
 * `DeathProbabilities` for the next "Beräkna" click; it does not itself
 * recompute anything, matching the explicit-calculate rule above -- so
 * changing an Advanced setting while old results are showing does not
 * invalidate them either. That is a deliberate simplification, not an
 * oversight: `main.ts`'s own `viewContext` builds a fresh `ModelContext`
 * object on every render regardless of whether anything relevant changed, so
 * there is no cheap way to tell "the settings actually changed" from "the
 * page merely re-rendered" -- and the sheet itself does not auto-rerun on a
 * settings change elsewhere in the workbook either.
 *
 * `INPUT_COLUMNS`/`OUTPUT_COLUMNS` are the one canonical description of
 * Mikrosim's own columns -- headers (in both languages), bounds, and the
 * mapping to/from the engine's own types -- shared with `mikrosimCsv.ts` so
 * the on-screen table and the CSV file can never drift apart. Their header
 * text is deliberately its own hardcoded table, not sourced from
 * `apps/web/src/i18n.ts`'s `t()`: cross-checking against the real workbook's
 * own Mikrosim header row (`reference/golden/golden-cases.csv`) turns up real
 * wording mismatches with the Start sheet's own labels for the same concept
 * (e.g. `t("realReturn")` is "Real avkastning", the sheet's own Mikrosim
 * column is "Real fondavkastning") -- mixing sources would silently mislabel
 * a column. English text is this port's own translation; there is no English
 * Mikrosim sheet, the same situation every other new-UI string in this app is
 * already in.
 */
import {
  BIRTH_YEARS,
  RETIREMENT_AGES,
  SCHEME_CHOICES,
  Table1Key,
  defaultContext,
  defaultInput,
  run,
} from "@typfallsmodellen/engine";
import type {
  DeathProbabilities,
  ModelContext,
  SchemeId,
  TypfallInput,
  TypfallResult,
} from "@typfallsmodellen/engine";

import type { Bounds, FieldSet } from "./controls.js";
import { fieldSet, span } from "./controls.js";
import { csvExportButton } from "./export.js";
import { kronor } from "./format.js";
import type { Lang } from "./i18n.js";
import { mikrosimRowsToCsv, parseMikrosimCsv } from "./mikrosimCsv.js";
import { headCell } from "./tables.js";

// `mikrosimCsv.ts` imports `MikrosimRow`/`newMikrosimRow`/`validateMikrosimRow`
// /`INPUT_COLUMNS`/`OUTPUT_COLUMNS`/`headerName` back from this file -- a
// circular import, but a safe one: every value each side uses from the other
// is only ever read inside a function body (a click handler, `parseMikrosimCsv`
// itself), never at module-evaluation time, so it does not matter which of
// the two finishes evaluating its own top level first.

const say = (l: Lang, sv: string, en: string) => (l === "sv" ? sv : en);

/** No cap on CSV import -- a file's row count is its own business -- but
 * "Lägg till rad" is capped so a slip of the mouse can't quietly build a
 * table too large to be useful. */
const MAX_ROWS = 500;

export const BORN: Bounds = { ...span(BIRTH_YEARS), step: 1 };
/** `Börjar arbeta vid ålder` has no extracted list; the sheet offers 15 to 40
 * -- the same bound `form.ts`/`compare.ts` already duplicate for the same
 * reason (neither exports it either). */
export const START_WORK: Bounds = { min: 15, max: 40, step: 1 };
export const RETIREMENT: Bounds = { ...span(RETIREMENT_AGES), step: 1 };
/** The annual equivalent of `form.ts`'s monthly `{min:0,max:1_000_000,step:100}`. */
export const SALARY: Bounds = { min: 0, max: 12_000_000, step: 1200 };
export const IPS: Bounds = { min: 0, max: 100_000, step: 100 };

function clampTo(value: number, bounds: Bounds): number {
  return Math.min(Math.max(value, bounds.min), bounds.max);
}

export interface MikrosimRow {
  readonly id: string;
  born: number;
  startWorkAge: number;
  retirementAge: number;
  /** Årslön -- annual, kr/year. Converted to `TypfallInput.monthlySalary`
   * (÷12) only at mapping time, so the on-screen cell, the CSV cell and the
   * workbook's own column all show literally the same number. */
  annualSalary: number;
  yearlyInflation: number;
  realGrowth: number;
  realReturn: number;
  /** `ModelContext.ipsMonthly`, raw -- a value >1 is kronor/month, a value
   * <=1 is a share of income instead, the same dual meaning `advanced.ts`'s
   * own `savingAmountOrShare` already documents and preserves. */
  ipsMonthly: number;
  scheme: SchemeId;
  /** Set only by "Beräkna". `undefined` means "not currently calculated for
   * the inputs this row holds right now". */
  result?: TypfallResult;
  /** Set instead of `result` when the row could not be run at all. */
  error?: string;
}

export function newMikrosimRow(id: string): MikrosimRow {
  const base = defaultInput();
  return {
    id,
    born: base.born,
    startWorkAge: base.startWorkAge,
    retirementAge: base.retirementAge,
    annualSalary: Math.round(base.monthlySalary * 12),
    yearlyInflation: base.yearlyInflation,
    realGrowth: base.realGrowth,
    realReturn: base.realReturn,
    ipsMonthly: 0,
    scheme: base.scheme,
  };
}

/** The row, as a typfall -- `married`, `ownIncome`, `pgbManual` and
 * `pgbConscription` all stay at `defaultInput()`'s own defaults, since none
 * of them are Mikrosim columns. */
export function mikrosimRowToInput(row: MikrosimRow): TypfallInput {
  return defaultInput({
    born: row.born,
    startWorkAge: row.startWorkAge,
    retirementAge: row.retirementAge,
    monthlySalary: row.annualSalary / 12,
    yearlyInflation: row.yearlyInflation,
    realGrowth: row.realGrowth,
    realReturn: row.realReturn,
    scheme: row.scheme,
  });
}

/** The row's own IPS figure laid over the shared, currently-in-force
 * `ModelContext` -- `ipsStart`/`privateSavingKind` are not Mikrosim columns,
 * so every row shares whatever the app's own Advanced settings currently
 * hold for those two. */
export function mikrosimRowToContext(row: MikrosimRow, shared: ModelContext): ModelContext {
  return { ...shared, ipsMonthly: row.ipsMonthly };
}

export function validateMikrosimRow(row: MikrosimRow, lang: Lang): string | undefined {
  if (row.born < BORN.min || row.born > BORN.max) {
    return say(
      lang,
      `Födelseår måste vara mellan ${BORN.min} och ${BORN.max}.`,
      `Birth year must be between ${BORN.min} and ${BORN.max}.`,
    );
  }
  if (row.startWorkAge < START_WORK.min || row.startWorkAge > START_WORK.max) {
    return say(
      lang,
      `Ålder vid arbetslivets start måste vara mellan ${START_WORK.min} och ${START_WORK.max}.`,
      `The start-of-work age must be between ${START_WORK.min} and ${START_WORK.max}.`,
    );
  }
  if (row.retirementAge < RETIREMENT.min || row.retirementAge > RETIREMENT.max) {
    return say(
      lang,
      `Pensionsåldern måste vara mellan ${RETIREMENT.min} och ${RETIREMENT.max}.`,
      `The retirement age must be between ${RETIREMENT.min} and ${RETIREMENT.max}.`,
    );
  }
  if (row.startWorkAge > row.retirementAge) {
    return say(
      lang,
      "Ålder vid arbetslivets start måste vara före pensionsåldern.",
      "The start-of-work age must be before the retirement age.",
    );
  }
  if (!SCHEME_CHOICES.some((choice) => choice.value === row.scheme)) {
    return say(lang, "Ogiltig tjänstepension.", "Invalid occupational pension.");
  }
  return undefined;
}

type InputControl =
  | { readonly kind: "number"; readonly bounds: Bounds }
  | { readonly kind: "percent" }
  | { readonly kind: "select" };

export interface InputColumnDef {
  readonly sv: string;
  readonly en: string;
  readonly control: InputControl;
  readonly info?: (lang: Lang) => string;
  readonly get: (row: MikrosimRow) => number;
  readonly set: (row: MikrosimRow, value: number) => void;
  /** Clamps a continuous value to this column's own bounds -- not used for
   * `discrete` columns, which are validated against the list instead. */
  readonly clamp: (value: number) => number;
  /** Set only for a column whose valid values are a fixed list rather than a
   * range -- `scheme`, the one Mikrosim column that isn't a plain number
   * line. */
  readonly discrete?: readonly number[];
}

/** Mikrosim's nine input columns (`Egen Lönelista` dropped, see the file
 * comment) -- one shared description driving the on-screen table, CSV
 * parsing and CSV export alike, so the three can never disagree with each
 * other about what a column means or where its bounds are. */
export const INPUT_COLUMNS: readonly InputColumnDef[] = [
  {
    sv: "Födelseår",
    en: "Birth year",
    control: { kind: "number", bounds: BORN },
    get: (r) => r.born,
    set: (r, v) => (r.born = v),
    clamp: (v) => clampTo(v, BORN),
  },
  {
    sv: "Börjar arbeta vid ålder",
    en: "Starts working at age",
    control: { kind: "number", bounds: START_WORK },
    get: (r) => r.startWorkAge,
    set: (r, v) => (r.startWorkAge = v),
    clamp: (v) => clampTo(v, START_WORK),
  },
  {
    sv: "Går i pension vid ålder",
    en: "Retires at age",
    control: { kind: "number", bounds: RETIREMENT },
    get: (r) => r.retirementAge,
    set: (r, v) => (r.retirementAge = v),
    clamp: (v) => clampTo(v, RETIREMENT),
  },
  {
    sv: "Årslön",
    en: "Annual salary",
    control: { kind: "number", bounds: SALARY },
    get: (r) => r.annualSalary,
    set: (r, v) => (r.annualSalary = v),
    clamp: (v) => clampTo(v, SALARY),
  },
  {
    sv: "Årlig inflation",
    en: "Yearly inflation",
    control: { kind: "percent" },
    get: (r) => r.yearlyInflation,
    set: (r, v) => (r.yearlyInflation = v),
    clamp: (v) => v,
  },
  {
    sv: "Real tillväxt",
    en: "Real growth",
    control: { kind: "percent" },
    get: (r) => r.realGrowth,
    set: (r, v) => (r.realGrowth = v),
    clamp: (v) => v,
  },
  {
    sv: "Real fondavkastning",
    en: "Real fund return",
    control: { kind: "percent" },
    get: (r) => r.realReturn,
    set: (r, v) => (r.realReturn = v),
    clamp: (v) => v,
  },
  {
    sv: "Privat pensionssparande (med avdragsrätt)",
    en: "Private pension saving (tax-deductible)",
    control: { kind: "number", bounds: IPS },
    info: (l) =>
      say(
        l,
        "Över 1 tolkas som kronor/månad; 1 eller mindre tolkas som andel av inkomsten (samma regel som i Avancerat läge).",
        "Above 1 is read as kronor/month; 1 or less is read as a share of income (the same rule as in Advanced mode).",
      ),
    get: (r) => r.ipsMonthly,
    set: (r, v) => (r.ipsMonthly = v),
    clamp: (v) => clampTo(v, IPS),
  },
  {
    sv: "Välj tjänstepension",
    en: "Occupational pension scheme",
    control: { kind: "select" },
    get: (r) => r.scheme,
    set: (r, v) => (r.scheme = v as SchemeId),
    clamp: (v) => v,
    discrete: SCHEME_CHOICES.map((c) => c.value),
  },
];

export interface OutputColumnDef {
  readonly sv: string;
  readonly en: string;
  readonly key: Table1Key;
}

/**
 * Mikrosim's twelve output columns, in the real sheet's own order
 * (`reference/golden/golden-cases.csv`'s own header row). `Eget sparande` is
 * `Table1Key.PrivateSaving` ("ips"), not `PrivateSavingAfterTax` ("pps") --
 * confirmed by laying the real header row against `tables.ts`'s
 * `TABLE1_LINES` order: every column matches a strictly-increasing,
 * adjacency-preserving walk through it with `TotalGross` pulled forward, and
 * `PrivateSavingAfterTax` is not one of Mikrosim's twelve columns at all --
 * `PrivateSaving` is the one immediately after `OccupationalPension`, exactly
 * where "Eget sparande" sits after "Tjänstepension" in the real header row.
 * `DisposableAtRetirement` ("dispEfterSkatt"), not `DisposableBeforeRetirement`
 * ("dispInkomst", the year *before* retirement), for the same reason.
 */
export const OUTPUT_COLUMNS: readonly OutputColumnDef[] = [
  { sv: "Slutlön", en: "Final salary", key: Table1Key.FinalSalary },
  { sv: "Brutto-pension", en: "Gross pension", key: Table1Key.TotalGross },
  { sv: "Inkomstpension", en: "Income pension", key: Table1Key.IncomePension },
  { sv: "Tilläggspension", en: "Supplementary pension", key: Table1Key.SupplementaryPension },
  { sv: "Premiepension", en: "Premium pension", key: Table1Key.PremiumPension },
  { sv: "Garanti-pension", en: "Guarantee pension", key: Table1Key.GuaranteePension },
  { sv: "P_tillägg", en: "Income pension supplement", key: Table1Key.IncomePensionSupplement },
  { sv: "Tjänstepension", en: "Occupational pension", key: Table1Key.OccupationalPension },
  { sv: "Eget sparande", en: "Private saving", key: Table1Key.PrivateSaving },
  { sv: "Efter skatt", en: "After tax", key: Table1Key.PensionAfterTax },
  { sv: "Bostadstillägg + ÄFS", en: "Housing supplement + old-age support", key: Table1Key.BenefitsAtRetirement },
  { sv: "Disponibel inkomst", en: "Disposable income", key: Table1Key.DisposableAtRetirement },
];

export function headerName(col: { readonly sv: string; readonly en: string }, lang: Lang): string {
  return lang === "sv" ? col.sv : col.en;
}

export interface MikrosimHandle {
  readonly element: HTMLElement;
  relabel(lang: Lang): void;
  /** Stashes the context/deaths the next "Beräkna" click runs against. Does
   * not itself compute anything -- see the file comment. */
  setContext(context: ModelContext, deaths: DeathProbabilities): void;
}

export function createMikrosimPanel(lang: Lang): MikrosimHandle {
  let currentLang = lang;
  let context: ModelContext = defaultContext();
  let deaths: DeathProbabilities | undefined;

  let nextId = 1;
  const freshId = () => String(nextId++);
  let rows: MikrosimRow[] = [newMikrosimRow(freshId())];

  const element = document.createElement("div");
  element.className = "mikrosim-panel";
  element.dataset.role = "mikrosim-panel";

  const intro = document.createElement("p");
  intro.className = "field-hint";

  const importNote = document.createElement("p");
  importNote.className = "field-hint";

  const fileError = document.createElement("p");
  fileError.className = "mikrosim-file-error";

  const actions = document.createElement("div");
  actions.className = "panel-actions";

  const addButton = document.createElement("button");
  addButton.type = "button";
  addButton.className = "export-btn";
  addButton.dataset.action = "add-mikrosim-row";
  addButton.addEventListener("click", () => {
    if (rows.length >= MAX_ROWS) return;
    rows.push(newMikrosimRow(freshId()));
    rebuildTable();
  });

  const calcButton = document.createElement("button");
  calcButton.type = "button";
  calcButton.className = "export-btn";
  calcButton.dataset.action = "calculate-mikrosim";
  calcButton.addEventListener("click", () => {
    if (!deaths) return;
    const activeDeaths = deaths;
    for (const row of rows) {
      // A row already flagged (e.g. an invalid scheme a CSV import couldn't
      // apply, so the field itself silently sits at its old, individually
      // valid value) is left alone until an edit clears it -- otherwise
      // re-validating would find nothing wrong with the field as it stands
      // now and quietly compute the row anyway, hiding the very problem the
      // flag exists to surface.
      if (row.error) continue;
      const invalid = validateMikrosimRow(row, currentLang);
      if (invalid) {
        row.error = invalid;
        delete row.result;
        continue;
      }
      try {
        row.result = run(mikrosimRowToInput(row), mikrosimRowToContext(row, context), {
          deaths: activeDeaths,
        });
        delete row.error;
      } catch {
        delete row.result;
        row.error = say(
          currentLang,
          "Beräkningen misslyckades för raden.",
          "The calculation failed for this row.",
        );
      }
    }
    rebuildTable();
  });

  const fileInput = document.createElement("input");
  fileInput.type = "file";
  fileInput.accept = ".csv,text/csv";
  fileInput.className = "mikrosim-file-input";
  fileInput.dataset.action = "mikrosim-import-file";
  fileInput.addEventListener("change", () => {
    const file = fileInput.files?.[0];
    // Cleared immediately so choosing the same file again still fires `change`.
    fileInput.value = "";
    if (!file) return;
    void file.text().then((text) => {
      const parsed = parseMikrosimCsv(text, currentLang);
      if (parsed.fileError) {
        fileError.textContent = parsed.fileError;
        return;
      }
      fileError.textContent = "";
      rows = parsed.rows.length > 0 ? [...parsed.rows] : [newMikrosimRow(freshId())];
      nextId = rows.length + 1;
      rebuildTable();
    });
  });
  const importLabel = document.createElement("label");
  importLabel.className = "export-btn mikrosim-import-label";
  const importText = document.createElement("span");
  importLabel.append(importText, fileInput);

  let exportBtn = csvExportButton(lang, "mikrosim.csv", () => mikrosimRowsToCsv(rows, currentLang));

  actions.append(addButton, calcButton, importLabel, exportBtn);

  const tableScroll = document.createElement("div");
  tableScroll.className = "scroll mikrosim-table-scroll";
  const table = document.createElement("table");
  table.className = "table mikrosim-table";
  const thead = table.createTHead();
  const tbody = table.createTBody();
  tableScroll.append(table);

  element.append(intro, actions, importNote, fileError, tableScroll);

  function buildHead(): void {
    thead.replaceChildren();
    const groupRow = thead.insertRow();
    const corner = document.createElement("th");
    corner.colSpan = 2;
    const inputsHead = document.createElement("th");
    inputsHead.colSpan = INPUT_COLUMNS.length;
    inputsHead.textContent = say(currentLang, "Indata", "Inputs");
    const outputsHead = document.createElement("th");
    outputsHead.colSpan = OUTPUT_COLUMNS.length;
    outputsHead.textContent = say(currentLang, "Resultat", "Results");
    const removeHead = document.createElement("th");
    groupRow.append(corner, inputsHead, outputsHead, removeHead);

    const nameRow = thead.insertRow();
    nameRow.append(document.createElement("th"), document.createElement("th"));
    for (const col of INPUT_COLUMNS) {
      nameRow.append(headCell(headerName(col, currentLang), col.info?.(currentLang)));
    }
    for (const col of OUTPUT_COLUMNS) {
      nameRow.append(headCell(headerName(col, currentLang)));
    }
    nameRow.append(document.createElement("th"));
  }

  function buildRow(row: MikrosimRow, index: number, cells: FieldSet): HTMLTableRowElement {
    const tr = document.createElement("tr");
    tr.dataset.mikrosimRow = row.id;

    const numCell = document.createElement("td");
    numCell.textContent = String(index + 1);
    numCell.className = "mikrosim-rownum";

    const statusCell = document.createElement("td");
    statusCell.className = "mikrosim-status";

    function clearResult(): void {
      delete row.result;
      delete row.error;
      renderStatus();
      renderOutputs();
    }

    const inputCells: HTMLTableCellElement[] = [];
    for (const col of INPUT_COLUMNS) {
      const td = document.createElement("td");
      if (col.control.kind === "number") {
        const field = cells.number(col.get(row), col.control.bounds, (v) => {
          col.set(row, v);
          clearResult();
        });
        td.append(field.element);
      } else if (col.control.kind === "percent") {
        const field = cells.percent(col.get(row), (v) => {
          col.set(row, v);
          clearResult();
        });
        td.append(field.element);
      } else {
        const field = cells.select(
          SCHEME_CHOICES.map((choice) => ({ value: choice.value, label: () => choice.label })),
          col.get(row),
          (v) => {
            col.set(row, v);
            clearResult();
          },
        );
        td.append(field.element);
      }
      inputCells.push(td);
    }

    const outputCells = OUTPUT_COLUMNS.map(() => document.createElement("td"));
    outputCells.forEach((td) => (td.className = "num"));

    function renderOutputs(): void {
      OUTPUT_COLUMNS.forEach((col, i) => {
        const found = row.result?.table1.find((r) => r.key === col.key);
        outputCells[i]!.textContent = found ? kronor(found.adjusted, currentLang) : "";
      });
    }

    function renderStatus(): void {
      statusCell.replaceChildren();
      if (row.error) {
        const mark = document.createElement("abbr");
        mark.className = "table-term mikrosim-error";
        mark.title = row.error;
        mark.textContent = "!";
        statusCell.append(mark);
      } else if (row.result && row.result.warnings.length > 0) {
        const mark = document.createElement("abbr");
        mark.className = "table-term";
        mark.title = row.result.warnings.map((w) => w.message).join(" ");
        mark.textContent = "⚠";
        statusCell.append(mark);
      }
    }

    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.className = "pgb-remove";
    removeBtn.dataset.action = "mikrosim-remove-row";
    removeBtn.textContent = "×";
    removeBtn.disabled = rows.length <= 1;
    removeBtn.setAttribute("aria-label", say(currentLang, "Ta bort raden", "Remove the row"));
    removeBtn.addEventListener("click", () => {
      const at = rows.findIndex((r) => r.id === row.id);
      if (at === -1 || rows.length <= 1) return;
      rows.splice(at, 1);
      rebuildTable();
    });
    const removeCell = document.createElement("td");
    removeCell.append(removeBtn);

    renderStatus();
    renderOutputs();

    tr.append(numCell, statusCell, ...inputCells, ...outputCells, removeCell);
    return tr;
  }

  function rebuildTable(): void {
    buildHead();
    // A throwaway container: `number`/`percent`/`select` never touch it or
    // the relabel list (only `field()` does), so this is purely a source of
    // fresh, correctly-clamped cell builders for the current language.
    const cellBuilders = fieldSet(document.createElement("div"), [], currentLang);
    tbody.replaceChildren(...rows.map((row, i) => buildRow(row, i, cellBuilders)));
    addButton.disabled = rows.length >= MAX_ROWS;
  }

  function applyText(l: Lang): void {
    intro.textContent = say(
      l,
      'Mikrosim körs som ett eget läge: varje rad är ett fristående typfall, inte en avvikelse mot ' +
        'formuläret till vänster. Lägg till rader för hand eller importera en CSV-fil, och klicka ' +
        'sedan på "Beräkna" för att fylla i resultatkolumnerna.',
      'Mikrosim runs as its own mode: each row is a standalone case, not a variation on the form to ' +
        'the left. Add rows by hand or import a CSV file, then click "Calculate" to fill in ' +
        'the result columns.',
    );
    addButton.textContent = say(l, "+ Lägg till rad", "+ Add row");
    calcButton.textContent = say(l, "Beräkna", "Calculate");
    importText.textContent = say(l, "Importera CSV", "Import CSV");
    importNote.textContent = say(
      l,
      "Att importera en fil ersätter alla rader i tabellen.",
      "Importing a file replaces every row in the table.",
    );
    // `csvExportButton` bakes its own button text in at creation time and
    // exposes no relabel hook, so the button itself is replaced; the CSV
    // callback closes over the live `rows`/`currentLang` variables (not a
    // snapshot), so only the visible label actually needed refreshing here.
    const next = csvExportButton(l, "mikrosim.csv", () => mikrosimRowsToCsv(rows, currentLang));
    exportBtn.replaceWith(next);
    exportBtn = next;
  }

  applyText(lang);
  rebuildTable();

  return {
    element,
    relabel(l) {
      currentLang = l;
      applyText(l);
      rebuildTable();
    },
    setContext(nextContext, nextDeaths) {
      context = nextContext;
      deaths = nextDeaths;
    },
  };
}
