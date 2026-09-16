/**
 * Three headline figures above Table 1.
 *
 * None of this is a new calculation, and none of it is SysLang text -- there
 * is no "replacement rate" row on the Start sheet. Every value reads a number
 * Table 1 or the per-age matrix already computes, the same way the CSV export
 * reuses the tables' own data rather than a second source of truth.
 *
 * The values are split out from the DOM building (`computeKpis` vs
 * `renderKpis`) so the arithmetic can be unit-tested directly: `apps/web` has
 * no DOM in its test environment, the same reason every other rendering
 * function here (`renderTable1`, the figures) is checked against the real
 * built page instead, by `tools/build/verify-offline.mjs`.
 */
import { Table1Key } from "@typfallsmodellen/engine";
import type { TypfallInput, TypfallResult } from "@typfallsmodellen/engine";

import { kronor, percent } from "./format.js";
import { t } from "./i18n.js";
import type { Lang } from "./i18n.js";

/**
 * The retirement age the run used, which is not always the one asked for.
 *
 * `startsetup` raises a retirement age below the cohort's earliest and says so
 * in a warning; every heading that names the age, and the KPI cards below,
 * read it back from there rather than from the input.
 */
export function retirementAge(input: TypfallInput, result: TypfallResult): number {
  const corrected = result.warnings.find((w) => w.field === "ParYear");
  return typeof corrected?.used === "number" ? corrected.used : input.retirementAge;
}

export interface KpiValues {
  /** `Table1Key.TotalGross`'s own monthly figure -- not re-derived from
   * `result.rows`, which disagrees with it: `closeRetirementYear` rebuilds
   * the retirement year's gross for Table 1 after that age's row has already
   * been written, so the two are not interchangeable. */
  readonly monthlyAtRetirement: number;
  /** The same row's `shareOfFinalSalary`: total pension gross over slutlön. */
  readonly replacementRate: number;
  /** `averageAnnualGross`, expressed per month. */
  readonly averageMonthly: number;
  readonly firstAge: number;
  readonly lastAge: number;
}

/**
 * The average annual gross pension over the years the pension is expected to
 * be paid, from retirement to `lifeIncome.throughAge` -- the same span
 * `lifeIncome` sums over, but a plain arithmetic mean rather than a sum
 * discounted at `context.discountRate`. Those answer different questions:
 * `lifeIncome.gross` is what the pension stream is worth today, this is what
 * a typical year actually pays.
 *
 * A simplification kept for this first cut: whole ages only, unweighted, no
 * partial final year (`lifeIncome()`'s own `Int(diverse) = counter` branch
 * is left out).
 */
export function averageAnnualGross(result: TypfallResult, par: number): number {
  const first = Math.trunc(par);
  const last = Math.trunc(result.lifeIncome.throughAge);
  const rows = result.rows.filter((r) => r.age >= first && r.age <= last);
  if (rows.length === 0) return 0;
  return rows.reduce((sum, r) => sum + r.brutto, 0) / rows.length;
}

export function computeKpis(result: TypfallResult, par: number): KpiValues {
  const totalGross = result.table1.find((r) => r.key === Table1Key.TotalGross);
  return {
    monthlyAtRetirement: totalGross?.monthly ?? 0,
    replacementRate: totalGross?.shareOfFinalSalary ?? 0,
    averageMonthly: averageAnnualGross(result, par) / 12,
    firstAge: Math.trunc(par),
    lastAge: Math.trunc(result.lifeIncome.throughAge),
  };
}

interface Kpi {
  readonly label: string;
  readonly value: string;
  /** A short suffix after the value, in a smaller weight -- "kr", not a
   * second bold word competing with the number for the reader's eye. */
  readonly unit?: string;
  readonly note: string;
}

function card(kpi: Kpi): HTMLElement {
  const el = document.createElement("div");
  el.className = "kpi-card";
  const label = document.createElement("p");
  label.className = "kpi-label";
  label.textContent = kpi.label;
  const value = document.createElement("p");
  value.className = "kpi-value";
  value.textContent = kpi.value;
  if (kpi.unit !== undefined) {
    const unit = document.createElement("span");
    unit.className = "kpi-unit";
    unit.textContent = ` ${kpi.unit}`;
    value.append(unit);
  }
  const note = document.createElement("p");
  note.className = "kpi-note";
  note.textContent = kpi.note;
  el.append(label, value, note);
  return el;
}

export function renderKpis(result: TypfallResult, lang: Lang, par: number): HTMLElement {
  const v = computeKpis(result, par);

  const kpis: readonly Kpi[] = [
    {
      label: lang === "sv" ? "Pension vid pensionering" : "Pension at retirement",
      value: kronor(v.monthlyAtRetirement, lang),
      unit: t("kr", lang),
      note:
        lang === "sv"
          ? `${t("totalGross", lang)}, ${t("perMonth", lang).toLowerCase()}, före skatt`
          : `${t("totalGross", lang)}, ${t("perMonth", lang).toLowerCase()}, before tax`,
    },
    {
      label:
        lang === "sv" ? "Kompensationsgrad vid pensionering" : "Replacement rate at retirement",
      value: percent(v.replacementRate, lang),
      note: `${t("totalGross", lang)}, ${t("shareOfFinalSalaryShort", lang).toLowerCase()}`,
    },
    {
      label:
        lang === "sv"
          ? "Genomsnittlig pension under pensionstiden"
          : "Average pension through retirement",
      value: kronor(v.averageMonthly, lang),
      unit: t("kr", lang),
      note:
        lang === "sv"
          ? `${t("perMonth", lang)}, ${v.firstAge}–${v.lastAge} års ålder, före skatt`
          : `${t("perMonth", lang)}, ages ${v.firstAge}–${v.lastAge}, before tax`,
    },
  ];

  const box = document.createElement("div");
  box.className = "kpi-row";
  box.append(...kpis.map(card));
  return box;
}
