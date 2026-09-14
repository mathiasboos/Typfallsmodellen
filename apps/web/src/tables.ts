/**
 * Table 1 and Table 2, as the Start sheet lays them out.
 *
 * Table 1 is the summary at retirement: four columns per line, A) to D), with
 * the sheet's own row order. Every heading, row label and footnote below is
 * cell-for-cell what the Start sheet composes -- `A23` is the corner cell,
 * `C23:F23` the four column heads, `A24:A45` the rows, `A37`, `A38` and `F37`
 * the notes underneath -- read out of the workbook with the formula reader in
 * tools/extract/formulas.py.
 *
 * Table 2 is the year-by-year cash flow from `rng_tabell2_startAge`, in the
 * twelve columns `VBA_go.bas:2183` writes.
 */
import { Table1Key } from "@typfallsmodellen/engine";
import type { Table1Row, Table2Row, TypfallResult } from "@typfallsmodellen/engine";

import { kronor, percent } from "./format.js";
import { t } from "./i18n.js";
import type { Lang, LabelName } from "./i18n.js";

/** What Table 1's headings and notes need beyond the result itself. */
export interface Table1View {
  /** The retirement age the run actually used. */
  readonly par: number;
  /** `Average_Earning`: how many years the final salary averages over. */
  readonly finalSalaryYears: number;
  /** `rng_Sista_PensRatt`, which decides whether the first note is shown. */
  readonly lastPensionRight: boolean;
}

/**
 * Table 1's lines, in the Start sheet's order.
 *
 * The blank line after `dispInkomst` is the sheet's own: everything above it is
 * the year *before* retirement -- `InputXGetY` calls those rows "lon t-1",
 * "netto t-1" and "disp t-1" -- and everything below is the retirement year,
 * which is what the corner cell "Pension vid N års ålder" names.
 */
type Line =
  | { readonly kind: "spacer" }
  | { readonly kind: "notes" }
  | { readonly kind: "row"; key: Table1Key; label: LabelName; strong?: boolean };

const TABLE1_LINES: readonly Line[] = [
  { kind: "row", key: Table1Key.FinalSalary, label: "finalSalary" },
  { kind: "row", key: Table1Key.SalaryAfterTax, label: "salaryAfterTax" },
  { kind: "row", key: Table1Key.DisposableBeforeRetirement, label: "disposable" },
  { kind: "spacer" },
  { kind: "row", key: Table1Key.IncomePension, label: "incomePension" },
  { kind: "row", key: Table1Key.SupplementaryPension, label: "supplementary" },
  { kind: "row", key: Table1Key.PremiumPension, label: "premium" },
  { kind: "row", key: Table1Key.GuaranteePension, label: "guarantee" },
  { kind: "row", key: Table1Key.IncomePensionSupplement, label: "iptFull" },
  { kind: "row", key: Table1Key.TotalPublicPension, label: "totalPublic", strong: true },
  { kind: "row", key: Table1Key.OccupationalPension, label: "occupationalShort" },
  { kind: "row", key: Table1Key.PrivateSaving, label: "privateSaving" },
  { kind: "row", key: Table1Key.TotalGross, label: "totalGross", strong: true },
  { kind: "notes" },
  { kind: "row", key: Table1Key.PensionAfterTax, label: "afterTax" },
  { kind: "row", key: Table1Key.BenefitsAtRetirement, label: "housingSupplement" },
  { kind: "row", key: Table1Key.PrivateSavingAfterTax, label: "privateSavingIsk" },
  { kind: "row", key: Table1Key.DisposableAtRetirement, label: "disposable", strong: true },
];

/** The four columns, as `C23:F23` composes their headings. */
const TABLE1_COLUMNS: readonly {
  readonly name: string;
  readonly letter: string;
  readonly head: (lang: Lang) => string;
  readonly text: (row: Table1Row, lang: Lang) => string;
}[] = [
  {
    name: "nominal",
    letter: "A",
    head: (l) => t("currentPricesKronor", l),
    text: (r, l) => kronor(r.nominal, l),
  },
  {
    name: "adjusted",
    letter: "B",
    head: (l) => `${t("fixedPrices", l)}, ${t("kronor", l)}`,
    text: (r, l) => kronor(r.adjusted, l),
  },
  {
    name: "monthly",
    letter: "C",
    head: (l) => `${t("perMonth", l)}, ${t("kronor", l)}`,
    text: (r, l) => kronor(r.monthly, l),
  },
  {
    name: "share",
    letter: "D",
    head: (l) => t("shareOfFinalSalaryShort", l),
    text: (r, l) => percent(r.shareOfFinalSalary, l),
  },
];

function cell(text: string, className?: string): HTMLTableCellElement {
  const td = document.createElement("td");
  td.textContent = text;
  if (className) td.className = className;
  return td;
}

function headCell(text: string): HTMLTableCellElement {
  const th = document.createElement("th");
  th.textContent = text;
  return th;
}

/**
 * Two row labels carry a number the run decides.
 *
 * `A24` writes the final salary's averaging window -- "Slutlön, 61 - 65 års
 * ålder" -- from `Average_Earning`, and `rng_tabell1_pt` recomputes the
 * supplement's qualifying years into its own label, which the extracted string
 * has the shipped typfall's 40/40 baked into.
 */
function label(name: LabelName, lang: Lang, result: TypfallResult, view: Table1View): string {
  const text = t(name, lang);
  if (name === "iptFull") {
    return text.replace(/\(\d+\/40\)/, `(${Math.min(result.qualifyingYears, 40)}/40)`);
  }
  if (name !== "finalSalary") return text;
  const years = view.finalSalaryYears;
  const from = years > 1 ? view.par - years : view.par - 1;
  return `${text}, ${from} - ${view.par - 1} ${t("yearsAge", lang).toLowerCase()}`;
}

export function renderTable1(
  result: TypfallResult,
  lang: Lang,
  view: Table1View,
): HTMLElement {
  const byKey = new Map(result.table1.map((r) => [r.key, r]));
  const table = document.createElement("table");
  table.className = "table table1";

  const head = table.createTHead().insertRow();
  // `A23`: the corner cell names the age everything below it is measured at.
  head.append(
    headCell(`${t("pensionWord", lang)} ${t("at", lang)} ${view.par} ${t("yearsAge", lang)}`),
  );
  for (const column of TABLE1_COLUMNS) head.append(headCell(`${column.letter}) ${column.head(lang)}`));

  const body = table.createTBody();
  for (const line of TABLE1_LINES) {
    if (line.kind === "spacer") {
      const tr = body.insertRow();
      tr.className = "spacer";
      const gap = cell("");
      gap.colSpan = 5;
      tr.append(gap);
      continue;
    }

    if (line.kind === "notes") {
      appendNotes(body, lang, view);
      continue;
    }

    const row: Table1Row | undefined = byKey.get(line.key);
    if (row === undefined) continue;
    const tr = body.insertRow();
    // The model's own key, so a test can find a row without reading the label
    // in whichever language happens to be showing.
    tr.dataset.key = line.key;
    if (line.strong) tr.className = "strong";
    tr.append(cell(label(line.label, lang, result, view)));
    for (const column of TABLE1_COLUMNS) {
      const td = cell(column.text(row, lang), "num");
      td.dataset.col = column.name;
      tr.append(td);
    }
  }

  return table;
}

/**
 * `A37`, `A38` and `F37`: the two notes under the gross total, and the heading
 * column D's lower block is measured against.
 *
 * The occupational note is the lifelong branch of `A38`; the alternative reads
 * "Valt uttag av tjänstepension är på N år" and belongs to a payout-length
 * setting that normal mode does not offer.
 */
function appendNotes(body: HTMLTableSectionElement, lang: Lang, view: Table1View): void {
  const lines = [
    view.lastPensionRight ? `${t("tableAboveShows", lang)} ${t("lastRightShort", lang)}` : "",
    t("occupationalLifelong", lang),
  ].filter((text) => text !== "");

  lines.forEach((text, i) => {
    const tr = body.insertRow();
    tr.className = "note";
    const td = cell(text, "note-text");
    td.colSpan = 4;
    tr.append(td);
    if (i === 0) {
      const against = cell(t("shareOfIncomeBefore", lang), "num col-note");
      against.rowSpan = lines.length;
      tr.append(against);
    }
  });
}

const TABLE2_COLUMNS: readonly { label: LabelName; get: (r: Table2Row) => number }[] = [
  { label: "year", get: (r) => r.year },
  { label: "age", get: (r) => r.age },
  { label: "salary", get: (r) => r.salary },
  { label: "incomeAndSupplementary", get: (r) => r.incomeAndSupplementary },
  { label: "premium", get: (r) => r.premium },
  { label: "occupationalPlusIps", get: (r) => r.occupationalAndPrivate },
  { label: "guarantee", get: (r) => r.guaranteeAndSupplement },
  { label: "grossIncome", get: (r) => r.gross },
  { label: "incomeAfterTax", get: (r) => r.net },
  { label: "benefits", get: (r) => r.benefits },
  { label: "privateSavingIsk", get: (r) => r.privateAfterTax },
  { label: "disposable", get: (r) => r.disposable },
];

/**
 * Table 2's amounts are already scaled by the model.
 *
 * `rng_Chart_Earning_factor` is 1 for annual and 12 for monthly, and
 * `buildTable2` divides by it, so the month/year switch belongs on the context
 * and not in this function -- see `viewContext` in main.ts.
 */
export function renderTable2(result: TypfallResult, lang: Lang): HTMLElement {
  const table = document.createElement("table");
  table.className = "table table2";

  const head = table.createTHead().insertRow();
  for (const col of TABLE2_COLUMNS) head.append(headCell(t(col.label, lang)));

  const body = table.createTBody();
  for (const row of result.table2) {
    const tr = body.insertRow();
    TABLE2_COLUMNS.forEach((col, i) => {
      const value = col.get(row);
      // The first two columns are a year and an age, not money.
      const text = i < 2 ? String(value) : kronor(value, lang);
      tr.append(cell(text, i < 2 ? undefined : "num"));
    });
  }

  return table;
}

// -------------------------------------------------------------- CSV export --

/**
 * A field for a delimited row: quoted, with internal quotes doubled, only when
 * it contains the delimiter, a quote or a newline -- the ordinary CSV rule.
 */
function csvField(value: string, delimiter: string): string {
  if (value.includes(delimiter) || value.includes('"') || value.includes("\n")) {
    return `"${value.replace(/"/g, '""')}"`;
  }
  return value;
}

function csvLine(cells: readonly string[], delimiter: string): string {
  return cells.map((c) => csvField(c, delimiter)).join(delimiter);
}

/**
 * Swedish Excel opens a CSV with `;` fields and a decimal comma; everyone
 * else's with `,` fields and a decimal point. Both tables are already
 * formatted in the chosen locale (`kronor`, `percent`), so the delimiter is
 * the one thing this has to choose for itself.
 */
function delimiterFor(lang: Lang): string {
  return lang === "sv" ? ";" : ",";
}

/**
 * Table 1 as a CSV, one line per data row in the sheet's own order. The
 * spacer and the two footnotes are prose, not data, and are left out.
 */
export function table1ToCsv(result: TypfallResult, lang: Lang, view: Table1View): string {
  const delimiter = delimiterFor(lang);
  const byKey = new Map(result.table1.map((r) => [r.key, r]));
  const lines: string[] = [
    csvLine(
      [
        `${t("pensionWord", lang)} ${t("at", lang)} ${view.par} ${t("yearsAge", lang)}`,
        ...TABLE1_COLUMNS.map((c) => `${c.letter}) ${c.head(lang)}`),
      ],
      delimiter,
    ),
  ];
  for (const line of TABLE1_LINES) {
    if (line.kind !== "row") continue;
    const row = byKey.get(line.key);
    if (row === undefined) continue;
    lines.push(
      csvLine(
        [label(line.label, lang, result, view), ...TABLE1_COLUMNS.map((c) => c.text(row, lang))],
        delimiter,
      ),
    );
  }
  return lines.join("\r\n");
}

/** Table 2 as a CSV, in the same twelve columns and row order as the table. */
export function table2ToCsv(result: TypfallResult, lang: Lang): string {
  const delimiter = delimiterFor(lang);
  const lines: string[] = [csvLine(TABLE2_COLUMNS.map((c) => t(c.label, lang)), delimiter)];
  for (const row of result.table2) {
    lines.push(
      csvLine(
        TABLE2_COLUMNS.map((col, i) => {
          const value = col.get(row);
          return i < 2 ? String(value) : kronor(value, lang);
        }),
        delimiter,
      ),
    );
  }
  return lines.join("\r\n");
}
