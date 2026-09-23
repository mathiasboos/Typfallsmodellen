/**
 * Table 1 and Table 2, as the Start sheet lays them out.
 *
 * Table 1 is the summary at retirement: four columns per line, with the
 * sheet's own row order. Every heading, row label and footnote below is
 * cell-for-cell what the Start sheet composes -- `A23` is the corner cell,
 * `C23:F23` the four column heads (the sheet leads each with a letter, A) to
 * D), dropped here on request), `A24:A45` the rows, `A37`, `A38` and `F37`
 * the notes underneath -- read out of the workbook with the formula reader in
 * tools/extract/formulas.py.
 *
 * Table 2 is the year-by-year cash flow from `rng_tabell2_startAge`, in the
 * twelve columns `VBA_go.bas:2183` writes, plus two the workbook never shows:
 * municipal and state tax, split out of `netto` for this port's own tax-per-
 * year chart (see `MvaluesRow.municipalTax`/`stateTax`).
 */
import { Table1Key } from "@typfallsmodellen/engine";
import type { Table1Row, Table2Row, TypfallResult } from "@typfallsmodellen/engine";

import { kronor, percent } from "./format.js";
import { t } from "./i18n.js";
import type { Lang, LabelName } from "./i18n.js";
import { computeKpis, kpiLabels } from "./kpis.js";
import { kronorCell, percentCell } from "./xlsx.js";
import type { SheetData } from "./xlsx.js";

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

/**
 * The four columns, as `C23:F23` composes their headings.
 *
 * The sheet leads each with a letter -- "A) Löpande priser, kronor" -- which
 * is dropped here on request; `name` is what still identifies a column
 * (`data-col`, and the CSV export's own header) now that the letter doesn't.
 */
const TABLE1_COLUMNS: readonly {
  readonly name: string;
  readonly head: (lang: Lang) => string;
  readonly text: (row: Table1Row, lang: Lang) => string;
  /** The same figure `text` formats, unrounded -- what an Excel cell holds
   * once the display formatting is a number format string instead of text. */
  readonly value: (row: Table1Row) => number;
  readonly kind: "kronor" | "percent";
}[] = [
  {
    name: "nominal",
    head: (l) => t("currentPricesKronor", l),
    text: (r, l) => kronor(r.nominal, l),
    value: (r) => r.nominal,
    kind: "kronor",
  },
  {
    name: "adjusted",
    head: (l) => `${t("fixedPrices", l)}, ${t("kronor", l)}`,
    text: (r, l) => kronor(r.adjusted, l),
    value: (r) => r.adjusted,
    kind: "kronor",
  },
  {
    name: "monthly",
    head: (l) => `${t("perMonth", l)}, ${t("kronor", l)}`,
    text: (r, l) => kronor(r.monthly, l),
    value: (r) => r.monthly,
    kind: "kronor",
  },
  {
    name: "share",
    head: (l) => t("shareOfFinalSalaryShort", l),
    text: (r, l) => percent(r.shareOfFinalSalary, l),
    value: (r) => r.shareOfFinalSalary,
    kind: "percent",
  },
];

export function cell(text: string, className?: string): HTMLTableCellElement {
  const td = document.createElement("td");
  td.textContent = text;
  if (className) td.className = className;
  return td;
}

/**
 * `info`, when given, names what the column's own figure is built from --
 * shown on hover/focus via a plain `<abbr title>` rather than a custom
 * tooltip: unlike the charts' hover readout, this is one static string per
 * column, not a value that changes with the pointer's position, so the
 * browser's own mechanism already does the job.
 */
export function headCell(text: string, info?: string): HTMLTableCellElement {
  const th = document.createElement("th");
  if (info === undefined) {
    th.textContent = text;
  } else {
    const abbr = document.createElement("abbr");
    abbr.className = "table-term";
    abbr.title = info;
    abbr.textContent = text;
    th.append(abbr);
  }
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
  for (const column of TABLE1_COLUMNS) head.append(headCell(column.head(lang)));

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

/** One scenario's own column-pair in `renderCompareTable`. */
export interface ScenarioColumn {
  /** The scenario's own display name, e.g. "Baseline" or a typed label. */
  readonly label: string;
  /** A CSS custom-property name, shared with that scenario's card and its
   * line in `renderCompareChart` -- one colour identifies it everywhere. */
  readonly colour: string;
  readonly result: TypfallResult;
  /** This scenario's own retirement age, post-correction. */
  readonly par: number;
}

/**
 * `renderTable1`'s own rows, lined up against several scenarios at once
 * instead of one -- built for `compare.ts`. Reuses `TABLE1_LINES`'s row order
 * and labels verbatim, and only the "monthly"/"share" entries of
 * `TABLE1_COLUMNS`: the nominal and price-adjusted columns say less once
 * salary itself is the thing being varied, so they are dropped here.
 *
 * The three KPI cards' own figures lead the table, each in whichever of the
 * two sub-columns it naturally has (a kronor amount or a percentage, never
 * both) and blank in the other -- `kpis.ts`'s `computeKpis`, not a second
 * calculation.
 *
 * Two row labels embed a number that can differ per scenario (`finalSalary`'s
 * age-range suffix, `iptFull`'s qualifying-years suffix) but the table has
 * only one row to show it in. Both use the *first* column's own scenario for
 * that text -- a deliberate simplification, not a bug: the row's *numbers*
 * are still each scenario's real ones regardless of which one the label text
 * happens to be read against.
 *
 * The two footnotes under the gross total (`appendNotes`) still apply --
 * `finalSalaryYears`/`lastPensionRight` come from the shared `ModelContext`,
 * never overridden per scenario -- but span the whole table rather than
 * repeating the "share of last year's income" side note under every
 * scenario's own percentage column, which would need a genuinely different
 * layout for what is disclosure text, not a number.
 */
export function renderCompareTable(
  columns: readonly ScenarioColumn[],
  lang: Lang,
  shared: { readonly finalSalaryYears: number; readonly lastPensionRight: boolean },
): HTMLElement {
  const anchor = columns[0]!;
  const anchorView: Table1View = { par: anchor.par, ...shared };
  const totalCols = 1 + columns.length * 2;

  const table = document.createElement("table");
  table.className = "table table1 compare-table";

  const head = table.createTHead();
  const nameRow = head.insertRow();
  nameRow.append(headCell(""));
  for (const column of columns) {
    const th = document.createElement("th");
    th.colSpan = 2;
    th.dataset.scenario = column.label;
    const swatch = document.createElement("span");
    swatch.className = "compare-swatch";
    swatch.style.background = `var(${column.colour})`;
    th.append(swatch, document.createTextNode(column.label));
    nameRow.append(th);
  }
  const subRow = head.insertRow();
  subRow.append(headCell(t("pensionWord", lang)));
  for (let i = 0; i < columns.length; i += 1) {
    subRow.append(
      headCell(`${t("perMonth", lang)}, ${t("kronor", lang)}`),
      headCell(t("shareOfFinalSalaryShort", lang)),
    );
  }

  const body = table.createTBody();
  const spacer = () => {
    const tr = body.insertRow();
    tr.className = "spacer";
    const gap = cell("");
    gap.colSpan = totalCols;
    tr.append(gap);
  };

  const [pensionLabel, replacementLabel, averageLabel] = kpiLabels(lang);
  const kpis = columns.map((column) => computeKpis(column.result, column.par));
  const kpiRow = (
    key: string,
    rowLabel: string,
    sub: "monthly" | "share",
    get: (i: number) => number,
  ) => {
    const tr = body.insertRow();
    // Not a `Table1Key` -- these three rows come from `computeKpis`, not
    // `result.table1` -- but the same stable-attribute idea `TABLE1_LINES`'s
    // own rows use below, so a check can find a value without reading its
    // label in whichever language happens to be showing.
    tr.dataset.key = key;
    tr.append(cell(rowLabel));
    columns.forEach((column, i) => {
      const value = get(i);
      const monthly = sub === "monthly" ? cell(kronor(value, lang), "num") : cell("", "num");
      monthly.dataset.col = "monthly";
      monthly.dataset.scenario = column.label;
      const share = sub === "share" ? cell(percent(value, lang), "num") : cell("", "num");
      share.dataset.col = "share";
      share.dataset.scenario = column.label;
      tr.append(monthly, share);
    });
  };
  kpiRow("kpi-pension", pensionLabel, "monthly", (i) => kpis[i]!.monthlyAtRetirement);
  kpiRow("kpi-replacement", replacementLabel, "share", (i) => kpis[i]!.replacementRate);
  kpiRow("kpi-average", averageLabel, "monthly", (i) => kpis[i]!.averageMonthly);

  spacer();

  const byKey = columns.map((column) => new Map(column.result.table1.map((r) => [r.key, r])));
  let notesShown = false;
  for (const line of TABLE1_LINES) {
    if (line.kind === "spacer") {
      spacer();
      continue;
    }
    if (line.kind === "notes") {
      if (!notesShown) {
        notesShown = true;
        const lines = [
          shared.lastPensionRight ? `${t("tableAboveShows", lang)} ${t("lastRightShort", lang)}` : "",
          t("occupationalLifelong", lang),
        ].filter((text) => text !== "");
        for (const text of lines) {
          const tr = body.insertRow();
          tr.className = "note";
          const td = cell(text, "note-text");
          td.colSpan = totalCols;
          tr.append(td);
        }
      }
      continue;
    }

    const tr = body.insertRow();
    tr.dataset.key = line.key;
    if (line.strong) tr.className = "strong";
    tr.append(cell(label(line.label, lang, anchor.result, anchorView)));
    columns.forEach((column, i) => {
      const row: Table1Row | undefined = byKey[i]!.get(line.key);
      const monthly = cell(row ? kronor(row.monthly, lang) : "", "num");
      monthly.dataset.col = "monthly";
      monthly.dataset.scenario = column.label;
      const share = cell(row ? percent(row.shareOfFinalSalary, lang) : "", "num");
      share.dataset.col = "share";
      share.dataset.scenario = column.label;
      tr.append(monthly, share);
    });
  }

  return table;
}

/**
 * What each column's own figure is built from, in one or two sentences --
 * not a SysLang extraction (the sheet has no per-column explanations of its
 * own), so written from `MvaluesRow`'s own field comments
 * (`packages/engine/src/model/state.ts`) and `buildTable2`'s composition of
 * them (`packages/engine/src/model/result.ts`), the same source this port's
 * own tax-per-year chart note already draws on for the tax split. Year and
 * age need none of this -- they carry no composition to explain.
 */
const TABLE2_COLUMNS: readonly {
  readonly head: (lang: Lang) => string;
  readonly info?: (lang: Lang) => string;
  readonly get: (r: Table2Row) => number;
}[] = [
  { head: (l) => t("year", l), get: (r) => r.year },
  { head: (l) => t("age", l), get: (r) => r.age },
  {
    head: (l) => t("salary", l),
    info: (l) => (l === "sv" ? "Årets lön eller annan förvärvsinkomst." : "The year's salary or other earned income."),
    get: (r) => r.salary,
  },
  {
    head: (l) => t("incomeAndSupplementary", l),
    info: (l) =>
      l === "sv"
        ? "Inkomstpension, den största delen av den allmänna pensionen, plus tilläggspension för den som är född 1953 eller tidigare."
        : "Income pension, the largest part of the public pension, plus tilläggspension (ATP) for those born in 1953 or earlier.",
    get: (r) => r.incomeAndSupplementary,
  },
  {
    head: (l) => t("premium", l),
    info: (l) =>
      l === "sv"
        ? "Premiepension: den del av den allmänna pensionen du själv väljer placering för."
        : "Premium pension: the part of the public pension you choose how to invest.",
    get: (r) => r.premium,
  },
  {
    head: (l) => t("occupationalPlusIps", l),
    info: (l) =>
      l === "sv"
        ? "Tjänstepension från arbetsgivaren plus eventuellt individuellt pensionssparande (IPS), före skatt."
        : "Occupational pension from the employer plus any individual pension saving (IPS), before tax.",
    get: (r) => r.occupationalAndPrivate,
  },
  {
    head: (l) => t("guarantee", l),
    info: (l) =>
      l === "sv"
        ? "Garantipension (lägstanivå för den som haft låg eller ingen inkomst) plus inkomstpensionstillägg och, för den som är född 1938–1953, garantitillägg."
        : "Guarantee pension (a minimum level for those with low or no income) plus the income pension supplement and, for those born 1938–1953, the guarantee supplement.",
    get: (r) => r.guaranteeAndSupplement,
  },
  {
    head: (l) => t("grossIncome", l),
    info: (l) =>
      l === "sv"
        ? "Summan av lön, allmän pension, tjänstepension och privat pensionssparande, före skatt."
        : "The sum of salary, public pension, occupational pension and private pension saving, before tax.",
    get: (r) => r.gross,
  },
  // Not a SysLang extraction -- the workbook never shows this split either.
  {
    head: (l) => (l === "sv" ? "Kommunal skatt" : "Municipal tax"),
    info: (l) =>
      l === "sv"
        ? "Kommunal inkomstskatt och kyrko-/begravningsavgift, efter jobbskatteavdrag och andra skattereduktioner."
        : "Municipal income tax and the church/burial fee, net of the earned-income tax credit and other tax reductions.",
    get: (r) => r.municipalTax,
  },
  {
    head: (l) => (l === "sv" ? "Statlig skatt" : "State tax"),
    // The formula also has a 25% band above a second, higher threshold --
    // "värnskatten", abolished in 2020 -- left out here on request: the
    // data sets that threshold to an unreachable 10^16 from 2020 on (see
    // NO_SECOND_THRESHOLD in packages/engine/src/data/projection.ts), so it
    // never actually fires for a present-day or future run, and naming it
    // here read as if it were still in force.
    info: (l) =>
      l === "sv"
        ? "Statlig inkomstskatt (20 % över brytpunkten), public service-avgiften, och eventuell skatt på kapital efter pensionering."
        : "State income tax (20% above the threshold), the public-service fee, and any capital-income tax after retirement.",
    get: (r) => r.stateTax,
  },
  {
    head: (l) => t("incomeAfterTax", l),
    info: (l) => (l === "sv" ? "Bruttoinkomst minus kommunal och statlig skatt." : "Gross income minus municipal and state tax."),
    get: (r) => r.net,
  },
  {
    head: (l) => t("benefits", l),
    info: (l) =>
      l === "sv"
        ? "Bostadstillägg, äldreförsörjningsstöd och andra behovsprövade tillägg."
        : "Housing supplement, income support for the elderly, and other means-tested benefits.",
    get: (r) => r.benefits,
  },
  {
    head: (l) => t("privateSavingIsk", l),
    info: (l) =>
      l === "sv"
        ? "Utbetalning från privat pensionssparande (ISK eller kapitalförsäkring), som inte beskattas som inkomst."
        : "Payout from private pension saving (an ISK or endowment insurance), which is not taxed as income.",
    get: (r) => r.privateAfterTax,
  },
  {
    head: (l) => t("disposable", l),
    info: (l) =>
      l === "sv"
        ? "Inkomst efter skatt plus bidrag och privat pensionssparande efter skatt."
        : "Income after tax plus benefits and private pension saving after tax.",
    get: (r) => r.disposable,
  },
];

/**
 * Table 2's columns, minus any that are zero for every row in this run --
 * Bidrag or Privat pensionssparande (ISK/KF), say, when this typfall never
 * pays them (private saving has no normal-mode input yet, so that column is
 * empty on every run until Phase 4 adds one). Year and age are identifiers,
 * not amounts, and always show; requested on the same principle as the
 * figures' own empty-legend-entry trim.
 */
function visibleTable2Columns(
  result: TypfallResult,
): readonly (typeof TABLE2_COLUMNS)[number][] {
  return TABLE2_COLUMNS.filter((col, i) => i < 2 || result.table2.some((row) => col.get(row) > 0));
}

/**
 * Table 2's amounts are already scaled by the model.
 *
 * `rng_Chart_Earning_factor` is 1 for annual and 12 for monthly, and
 * `buildTable2` divides by it, so the month/year switch belongs on the context
 * and not in this function -- see `viewContext` in main.ts.
 */
export function renderTable2(result: TypfallResult, lang: Lang): HTMLElement {
  const columns = visibleTable2Columns(result);
  const table = document.createElement("table");
  table.className = "table table2";

  const head = table.createTHead().insertRow();
  for (const col of columns) head.append(headCell(col.head(lang), col.info?.(lang)));

  const body = table.createTBody();
  for (const row of result.table2) {
    const tr = body.insertRow();
    columns.forEach((col, i) => {
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
        ...TABLE1_COLUMNS.map((c) => c.head(lang)),
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
  const columns = visibleTable2Columns(result);
  const lines: string[] = [csvLine(columns.map((c) => c.head(lang)), delimiter)];
  for (const row of result.table2) {
    lines.push(
      csvLine(
        columns.map((col, i) => {
          const value = col.get(row);
          return i < 2 ? String(value) : kronor(value, lang);
        }),
        delimiter,
      ),
    );
  }
  return lines.join("\r\n");
}

/**
 * `renderCompareTable` as a CSV -- the same two rows of headers (a scenario's
 * name, blank beside it standing in for its colspan; then its two sub-column
 * heads), the three KPI rows, and `TABLE1_LINES`'s own rows, in the same
 * order. The spacer and the two footnotes are prose, not data, and are left
 * out, the same as `table1ToCsv`.
 */
export function compareTableToCsv(
  columns: readonly ScenarioColumn[],
  lang: Lang,
  shared: { readonly finalSalaryYears: number; readonly lastPensionRight: boolean },
): string {
  const delimiter = delimiterFor(lang);
  const anchor = columns[0]!;
  const anchorView: Table1View = { par: anchor.par, ...shared };

  const nameRow = [""];
  const subRow = [t("pensionWord", lang)];
  for (const column of columns) {
    nameRow.push(column.label, "");
    subRow.push(`${t("perMonth", lang)}, ${t("kronor", lang)}`, t("shareOfFinalSalaryShort", lang));
  }
  const lines: string[] = [csvLine(nameRow, delimiter), csvLine(subRow, delimiter)];

  const [pensionLabel, replacementLabel, averageLabel] = kpiLabels(lang);
  const kpis = columns.map((column) => computeKpis(column.result, column.par));
  const kpiLine = (rowLabel: string, sub: "monthly" | "share", get: (i: number) => number) => {
    const cells = [rowLabel];
    columns.forEach((_column, i) => {
      const value = get(i);
      cells.push(sub === "monthly" ? kronor(value, lang) : "", sub === "share" ? percent(value, lang) : "");
    });
    lines.push(csvLine(cells, delimiter));
  };
  kpiLine(pensionLabel, "monthly", (i) => kpis[i]!.monthlyAtRetirement);
  kpiLine(replacementLabel, "share", (i) => kpis[i]!.replacementRate);
  kpiLine(averageLabel, "monthly", (i) => kpis[i]!.averageMonthly);

  const byKey = columns.map((column) => new Map(column.result.table1.map((r) => [r.key, r])));
  for (const line of TABLE1_LINES) {
    if (line.kind !== "row") continue;
    const cells = [label(line.label, lang, anchor.result, anchorView)];
    columns.forEach((_column, i) => {
      const row: Table1Row | undefined = byKey[i]!.get(line.key);
      cells.push(row ? kronor(row.monthly, lang) : "", row ? percent(row.shareOfFinalSalary, lang) : "");
    });
    lines.push(csvLine(cells, delimiter));
  }
  return lines.join("\r\n");
}

// ------------------------------------------------------------- Excel export --

/**
 * Table 1 as `write-excel-file`'s sheet-data shape -- the same rows
 * `table1ToCsv` writes, but each amount a live number with a display format
 * instead of pre-formatted text, so the numbers stay usable in a formula
 * once they're in Excel.
 */
export function table1ToXlsxRows(result: TypfallResult, lang: Lang, view: Table1View): SheetData {
  const byKey = new Map(result.table1.map((r) => [r.key, r]));
  const rows: SheetData = [
    [
      `${t("pensionWord", lang)} ${t("at", lang)} ${view.par} ${t("yearsAge", lang)}`,
      ...TABLE1_COLUMNS.map((c) => c.head(lang)),
    ],
  ];
  for (const line of TABLE1_LINES) {
    if (line.kind !== "row") continue;
    const row = byKey.get(line.key);
    if (row === undefined) continue;
    rows.push([
      label(line.label, lang, result, view),
      ...TABLE1_COLUMNS.map((c) =>
        c.kind === "kronor" ? kronorCell(c.value(row)) : percentCell(c.value(row)),
      ),
    ]);
  }
  return rows;
}

/** Table 2 as `write-excel-file`'s sheet-data shape -- year and age as plain
 * numbers, every amount a kronor-formatted number. */
export function table2ToXlsxRows(result: TypfallResult, lang: Lang): SheetData {
  const columns = visibleTable2Columns(result);
  const rows: SheetData = [columns.map((c) => c.head(lang))];
  for (const row of result.table2) {
    rows.push(columns.map((col, i) => (i < 2 ? col.get(row) : kronorCell(col.get(row)))));
  }
  return rows;
}

/** `renderCompareTable` as `write-excel-file`'s sheet-data shape -- the same
 * rows `compareTableToCsv` writes, with live numbers in place of text. */
export function compareTableToXlsxRows(
  columns: readonly ScenarioColumn[],
  lang: Lang,
  shared: { readonly finalSalaryYears: number; readonly lastPensionRight: boolean },
): SheetData {
  const anchor = columns[0]!;
  const anchorView: Table1View = { par: anchor.par, ...shared };

  const nameRow: SheetData[number] = [""];
  const subRow: SheetData[number] = [t("pensionWord", lang)];
  for (const column of columns) {
    nameRow.push(column.label, "");
    subRow.push(`${t("perMonth", lang)}, ${t("kronor", lang)}`, t("shareOfFinalSalaryShort", lang));
  }
  const rows: SheetData = [nameRow, subRow];

  const [pensionLabel, replacementLabel, averageLabel] = kpiLabels(lang);
  const kpis = columns.map((column) => computeKpis(column.result, column.par));
  const kpiRow = (rowLabel: string, sub: "monthly" | "share", get: (i: number) => number) => {
    const cells: SheetData[number] = [rowLabel];
    columns.forEach((_column, i) => {
      const value = get(i);
      cells.push(sub === "monthly" ? kronorCell(value) : "", sub === "share" ? percentCell(value) : "");
    });
    rows.push(cells);
  };
  kpiRow(pensionLabel, "monthly", (i) => kpis[i]!.monthlyAtRetirement);
  kpiRow(replacementLabel, "share", (i) => kpis[i]!.replacementRate);
  kpiRow(averageLabel, "monthly", (i) => kpis[i]!.averageMonthly);

  const byKey = columns.map((column) => new Map(column.result.table1.map((r) => [r.key, r])));
  for (const line of TABLE1_LINES) {
    if (line.kind !== "row") continue;
    const cells: SheetData[number] = [label(line.label, lang, anchor.result, anchorView)];
    columns.forEach((_column, i) => {
      const row: Table1Row | undefined = byKey[i]!.get(line.key);
      cells.push(row ? kronorCell(row.monthly) : "", row ? percentCell(row.shareOfFinalSalary) : "");
    });
    rows.push(cells);
  }
  return rows;
}
