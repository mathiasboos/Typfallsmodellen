/**
 * Table 1 and Table 2, as the Start sheet lays them out.
 *
 * Table 1 is the summary at the retirement age, four columns per line; Table 2
 * is the year-by-year cash flow from `rng_tabell2_startAge`. Both come straight
 * off `TypfallResult`, which is the workbook's own output model rather than
 * something this app invents -- see docs/ARCHITECTURE.md.
 */
import { Table1Key } from "@typfallsmodellen/engine";
import type { Table1Row, Table2Row, TypfallResult } from "@typfallsmodellen/engine";

import { kronor, percent } from "./format.js";
import { t } from "./i18n.js";
import type { Lang, LabelName } from "./i18n.js";

/**
 * Table 1's lines, in the Start sheet's order.
 *
 * Two sections, because the sheet has two: the three rows above the break are
 * the year *before* retirement -- `InputXGetY` calls them "lon t-1", "netto
 * t-1" and "disp t-1" -- and everything below is the retirement year itself.
 * Without the break the table shows "Disponibel inkomst" twice with different
 * numbers and no way to tell which is which.
 */
type Line =
  | { readonly kind: "section"; readonly sv: string; readonly en: string }
  | { readonly kind: "row"; key: Table1Key; label: LabelName; strong?: boolean };

const TABLE1_LINES: readonly Line[] = [
  { kind: "section", sv: "Året före pension", en: "The year before retirement" },
  { kind: "row", key: Table1Key.FinalSalary, label: "finalSalary" },
  { kind: "row", key: Table1Key.SalaryAfterTax, label: "salaryAfterTax" },
  { kind: "row", key: Table1Key.DisposableBeforeRetirement, label: "disposable" },
  { kind: "section", sv: "Vid pension", en: "At retirement" },
  { kind: "row", key: Table1Key.IncomePension, label: "incomePension" },
  { kind: "row", key: Table1Key.SupplementaryPension, label: "supplementary" },
  { kind: "row", key: Table1Key.PremiumPension, label: "premium" },
  { kind: "row", key: Table1Key.GuaranteePension, label: "guarantee" },
  { kind: "row", key: Table1Key.IncomePensionSupplement, label: "iptFull" },
  { kind: "row", key: Table1Key.TotalPublicPension, label: "totalPublic", strong: true },
  { kind: "row", key: Table1Key.OccupationalPension, label: "occupationalShort" },
  { kind: "row", key: Table1Key.PrivateSaving, label: "privateSaving" },
  { kind: "row", key: Table1Key.TotalGross, label: "totalGross", strong: true },
  { kind: "row", key: Table1Key.PensionAfterTax, label: "afterTax" },
  { kind: "row", key: Table1Key.BenefitsAtRetirement, label: "benefits" },
  { kind: "row", key: Table1Key.DisposableAtRetirement, label: "disposable", strong: true },
];

function cell(text: string, className?: string): HTMLTableCellElement {
  const td = document.createElement("td");
  td.textContent = text;
  if (className) td.className = className;
  return td;
}

/**
 * The supplement's label carries the qualifying years behind it.
 *
 * The workbook writes it into `rng_tabell1_pt` as "Pensionstillägg (37/40)",
 * recomputed per run. The extracted label has the shipped typfall's 40/40 baked
 * in, so the ratio is replaced with this run's own.
 */
function label(name: LabelName, lang: Lang, result: TypfallResult): string {
  const text = t(name, lang);
  if (name !== "iptFull") return text;
  return text.replace(/\(\d+\/40\)/, `(${Math.min(result.qualifyingYears, 40)}/40)`);
}

function headCell(text: string): HTMLTableCellElement {
  const th = document.createElement("th");
  th.textContent = text;
  return th;
}

export function renderTable1(result: TypfallResult, lang: Lang, monthly: boolean): HTMLElement {
  const byKey = new Map(result.table1.map((r) => [r.key, r]));
  const table = document.createElement("table");
  table.className = "table table1";

  const head = table.createTHead().insertRow();
  head.append(
    headCell(""),
    headCell(monthly ? t("perMonth", lang) : t("kronor", lang)),
    headCell(t("shareOfFinalSalary", lang)),
  );

  const body = table.createTBody();
  for (const line of TABLE1_LINES) {
    if (line.kind === "section") {
      const tr = body.insertRow();
      tr.className = "section";
      const th = document.createElement("th");
      th.colSpan = 3;
      th.scope = "colgroup";
      th.textContent = lang === "sv" ? line.sv : line.en;
      tr.append(th);
      continue;
    }

    const row: Table1Row | undefined = byKey.get(line.key);
    if (row === undefined) continue;
    const tr = body.insertRow();
    // The model's own key, so a test can find a row without reading the label
    // in whichever language happens to be showing.
    tr.dataset.key = line.key;
    if (line.strong) tr.className = "strong";
    tr.append(cell(label(line.label, lang, result)));
    tr.append(cell(kronor(monthly ? row.monthly : row.adjusted, lang), "num"));
    tr.append(
      cell(row.shareOfFinalSalary > 0 ? percent(row.shareOfFinalSalary, lang) : "", "num"),
    );
  }

  return table;
}

const TABLE2_COLUMNS: readonly { label: LabelName; get: (r: Table2Row) => number }[] = [
  { label: "age", get: (r) => r.age },
  { label: "year", get: (r) => r.year },
  { label: "salary", get: (r) => r.salary },
  { label: "incomeAndSupplementary", get: (r) => r.incomeAndSupplementary },
  { label: "premium", get: (r) => r.premium },
  { label: "occupationalPlusIps", get: (r) => r.occupationalAndPrivate },
  { label: "guarantee", get: (r) => r.guaranteeAndSupplement },
  { label: "grossIncome", get: (r) => r.gross },
  { label: "incomeAfterTax", get: (r) => r.net },
  { label: "benefits", get: (r) => r.benefits },
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
      // The first two columns are an age and a year, not money.
      const text = i < 2 ? String(value) : kronor(value, lang);
      tr.append(cell(text, i < 2 ? undefined : "num"));
    });
  }

  return table;
}
