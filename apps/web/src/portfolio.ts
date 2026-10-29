/**
 * Portföljvärde: the pension capital behind the forecast, account by account and
 * year by year -- what was in it, what was paid in, what it earned, what was paid
 * out and what it cost.
 *
 * The same tables and charts as the "Portfolio Value" tab of the site's
 * Pensionsplaneraren (Retirement_Planner.html), on this model's own figures: the
 * engine's `result.ledger`, which follows each pot the model keeps -- inkomstpension,
 * premiepension, the occupational pension, private pension saving (IPS) and private
 * saving in a KF, an ISK or a depot. The tab follows Prognos: the same person, the
 * same settings (Avancerat included) and the same run, so a withdrawal here is the
 * same amount Table 2 shows paid that year.
 *
 * Garantipension, inkomstpensionstillägg and tilläggspension (ATP) are paid without
 * any balance behind them, and a defined-benefit occupational pension is paid by the
 * employer, so none of them is an account here. The income pension has no balance in
 * the model after retirement either -- it is turned into a pension for life -- so from
 * the retirement year the engine values the pension still to come instead (see
 * `finishLedger`).
 *
 * Amounts are in the units Table 2 is: fixed prices, the reference year's wage level
 * or nominal, as Avancerat says. Each year is expressed in its own year's units and
 * opens on the last year's closing balance as shown, so the return is the real one.
 *
 * "Skattekorrigerat" (tax-adjusted) values every account whose withdrawals are taxed
 * as income -- all but the private saving -- after the average tax Prognos's Table 1
 * takes from the pension at retirement. One rate for every amount of the account, so
 * the account still adds up.
 *
 * Nothing here touches the DOM, so the arithmetic is tested directly
 * (`apps/web/test/portfolio.test.ts`).
 */
import { POTS, Table1Key } from "@typfallsmodellen/engine";
import type { PotId, PotYear, TypfallResult } from "@typfallsmodellen/engine";

import { kronor } from "./format.js";
import type { Lang } from "./i18n.js";
import { csvLine, delimiterFor } from "./tables.js";
import { kronorCell } from "./xlsx.js";
import type { SheetData } from "./xlsx.js";

const say = (l: Lang, sv: string, en: string) => (l === "sv" ? sv : en);

/** One of the five accounts, in the order they stack. */
export type AccountId = PotId;

export type AccountKind = "public" | "insurance" | "taxed";

export interface AccountMeta {
  readonly id: AccountId;
  /** How it is taxed: the public pension and pension insurances on the way out, private saving on the way in. */
  readonly kind: AccountKind;
  /** A custom property in styles.css. */
  readonly colour: string;
}

/**
 * Stack order, bottom up. Validated with the `dataviz` skill's checker in that order
 * (see the `--fig-pv-*` note in styles.css): no two neighbours are hard to tell apart.
 */
export const ACCOUNTS: readonly AccountMeta[] = [
  { id: "ip", kind: "public", colour: "--fig-pv-ip" },
  { id: "pp", kind: "public", colour: "--fig-pv-pp" },
  { id: "tjp", kind: "insurance", colour: "--fig-pv-tjp" },
  { id: "ips", kind: "insurance", colour: "--fig-pv-ips" },
  { id: "pps", kind: "taxed", colour: "--fig-pv-pps" },
];

export function accountMeta(id: AccountId): AccountMeta {
  return ACCOUNTS.find((a) => a.id === id)!;
}

/** The model's private saving kinds, as `ModelContext.privateSavingKind` numbers them. */
const SAVING_KIND: Readonly<Record<number, { sv: string; en: string }>> = {
  1: { sv: "kapitalförsäkring", en: "endowment insurance" },
  2: { sv: "ISK", en: "ISK" },
  3: { sv: "depå", en: "securities account" },
};

export function accountName(id: AccountId, lang: Lang, savingKind = 2): string {
  switch (id) {
    case "ip":
      return say(lang, "Inkomstpension", "Income pension");
    case "pp":
      return say(lang, "Premiepension", "Premium pension");
    case "tjp":
      return say(lang, "Tjänstepension", "Occupational pension");
    case "ips":
      return say(lang, "Privat pensionssparande", "Private pension saving");
    case "pps": {
      const kind = SAVING_KIND[savingKind] ?? SAVING_KIND[2]!;
      return say(lang, `Privat sparande (${kind.sv})`, `Private saving (${kind.en})`);
    }
  }
}

/** The badge beside an account's name. */
export function kindName(kind: AccountKind, lang: Lang): string {
  switch (kind) {
    case "public":
      return say(lang, "allmän pension · skatt vid uttag", "public pension · taxed when paid");
    case "insurance":
      return say(lang, "pensionsförsäkring · skatt vid uttag", "pension insurance · taxed when paid");
    case "taxed":
      return say(lang, "beskattat sparande · skattefritt uttag", "taxed saving · paid tax-free");
  }
}

/** Totalt, or every taxed account after the average tax on the pension. */
export type PortfolioMode = "total" | "taxadj";

/** One account's year as the tab shows it. `open + contribution + growth - withdrawal - fee - yieldTax + adjustment = close`. */
export interface Line {
  readonly open: number;
  readonly contribution: number;
  readonly growth: number;
  readonly withdrawal: number;
  readonly fee: number;
  readonly yieldTax: number;
  readonly adjustment: number;
  readonly close: number;
}

export const LINE_KEYS = ["open", "contribution", "growth", "withdrawal", "fee", "yieldTax", "adjustment", "close"] as const;

export type LineKey = (typeof LINE_KEYS)[number];

export interface PortfolioYear {
  readonly year: number;
  readonly age: number;
}

export interface Account {
  readonly id: AccountId;
  readonly meta: AccountMeta;
  /** One per year of `Portfolio.years`. */
  readonly lines: readonly Line[];
}

export interface Portfolio {
  readonly years: readonly PortfolioYear[];
  /** The accounts that hold or move anything, in stack order. */
  readonly accounts: readonly Account[];
  /** The age the public pension starts at, and its year. */
  readonly retirementAge: number;
  readonly retirementYear: number;
  readonly mode: PortfolioMode;
  /** The average tax on the pension at retirement, from Table 1 (0-1). */
  readonly taxRate: number;
  /** `ModelContext.privateSavingKind`, which names the private saving account. */
  readonly savingKind: number;
  /** Whether a balance typed in under Avancerat replaces one somewhere, which earns its own row. */
  readonly hasAdjustment: boolean;
}

/**
 * The average tax Table 1 takes from the pension at retirement: 1 - pension after tax / gross.
 * The private saving paid out tax-free is in neither.
 */
export function pensionTaxRate(result: TypfallResult): number {
  const byKey = new Map(result.table1.map((r) => [r.key, r]));
  const gross = byKey.get(Table1Key.TotalGross)?.nominal ?? 0;
  const net = byKey.get(Table1Key.PensionAfterTax)?.nominal ?? 0;
  if (gross <= 0) return 0;
  return Math.min(Math.max(1 - net / gross, 0), 1);
}

const isEmpty = (y: PotYear): boolean =>
  y.open === 0 && y.close === 0 && y.contribution === 0 && y.withdrawal === 0 && y.adjustment === 0;

/**
 * The tab's figures from a run.
 *
 * @param par        the retirement age the run used (`retirementAge` in kpis.ts)
 * @param savingKind `ModelContext.privateSavingKind`
 */
export function buildPortfolio(
  result: TypfallResult,
  par: number,
  mode: PortfolioMode,
  savingKind: number,
): Portfolio {
  const taxRate = pensionTaxRate(result);

  // From the first year anything is in any account to the model's last age.
  const firstUsed = result.ledger.findIndex((line) => POTS.some((pot) => !isEmpty(line.pots[pot])));
  const ledger = firstUsed < 0 ? [] : result.ledger.slice(firstUsed);
  const years = ledger.map((line) => ({ year: line.year, age: line.age }));

  const accounts: Account[] = [];
  for (const meta of ACCOUNTS) {
    if (!ledger.some((line) => !isEmpty(line.pots[meta.id]))) continue;
    const tax = mode === "taxadj" && meta.kind !== "taxed" ? 1 - taxRate : 1;
    let shown = 0;
    const lines = ledger.map((line, i): Line => {
      const y = line.pots[meta.id];
      // Table 2's units, which the ledger carries the exact factor for.
      const f = line.priceFactor * tax;
      const open = i === 0 ? y.open * f : shown;
      const close = y.close * f;
      const contribution = y.contribution * f;
      const withdrawal = y.withdrawal * f;
      const fee = y.fee * f;
      const yieldTax = y.yieldTax * f;
      const adjustment = y.adjustment * f;
      shown = close;
      return {
        open,
        contribution,
        // The rest of the change: the return in this year's units, and the inflation the
        // move from last year's units takes off it.
        growth: close - open - contribution + withdrawal + fee + yieldTax - adjustment,
        withdrawal,
        fee,
        yieldTax,
        adjustment,
        close,
      };
    });
    accounts.push({ id: meta.id, meta, lines });
  }

  const par0 = Math.trunc(par);
  const retirementYear = years.find((y) => y.age === par0)?.year ?? (years[0] ? years[0].year + (par0 - years[0].age) : 0);
  return {
    years,
    accounts,
    retirementAge: par0,
    retirementYear,
    mode,
    taxRate,
    savingKind,
    hasAdjustment: accounts.some((a) => a.lines.some((l) => Math.abs(l.adjustment) >= 0.5)),
  };
}

// ------------------------------------------------------------ the summary ---

/** A row of the all-accounts table: a label, and a value per year. */
export interface SummaryRow {
  readonly key: string;
  readonly label: string;
  readonly values: readonly number[];
  readonly strong?: boolean;
  /** Explains the row on hover. */
  readonly title?: string;
}

const sumOver = (accounts: readonly Account[], i: number, key: LineKey): number =>
  accounts.reduce((sum, a) => sum + (a.lines[i]?.[key] ?? 0), 0);

/**
 * The all-accounts summary, the planner's rows: contributions and withdrawals split
 * between the accounts paid out tax-free and those taxed when paid out.
 */
export function summaryRows(p: Portfolio, lang: Lang): readonly SummaryRow[] {
  const afterTax = p.accounts.filter((a) => a.meta.kind === "taxed");
  const preTax = p.accounts.filter((a) => a.meta.kind !== "taxed");
  const each = (f: (i: number) => number) => p.years.map((_, i) => f(i));
  const adjTitle =
    p.mode === "taxadj"
      ? say(lang, "Efter den genomsnittliga skatten på pensionen (Skattekorrigerat)", "After the average tax on the pension (Tax-adjusted)")
      : undefined;
  const rows: SummaryRow[] = [
    { key: "open", label: lineLabel("open", lang), values: each((i) => sumOver(p.accounts, i, "open")), strong: true },
    {
      key: "contribution-after",
      label: say(lang, "Insättning, konton efter skatt", "Contribution, after-tax accounts"),
      values: each((i) => sumOver(afterTax, i, "contribution")),
    },
    {
      key: "contribution-before",
      label: say(lang, "Insättning, konton före skatt", "Contribution, pre-tax accounts"),
      values: each((i) => sumOver(preTax, i, "contribution")),
      ...(adjTitle === undefined ? {} : { title: adjTitle }),
    },
    {
      key: "withdrawal-after",
      label: say(lang, "Uttag, konton efter skatt", "Withdrawal, after-tax accounts"),
      values: each((i) => sumOver(afterTax, i, "withdrawal")),
    },
    {
      key: "withdrawal-before",
      label: say(lang, "Uttag, konton före skatt", "Withdrawal, pre-tax accounts"),
      values: each((i) => sumOver(preTax, i, "withdrawal")),
      ...(adjTitle === undefined ? {} : { title: adjTitle }),
    },
    { key: "growth", label: lineLabel("growth", lang), values: each((i) => sumOver(p.accounts, i, "growth")) },
    { key: "yieldTax", label: lineLabel("yieldTax", lang), values: each((i) => sumOver(p.accounts, i, "yieldTax")) },
    { key: "fee", label: lineLabel("fee", lang), values: each((i) => sumOver(p.accounts, i, "fee")) },
  ];
  if (p.hasAdjustment) {
    rows.push({ key: "adjustment", label: lineLabel("adjustment", lang), values: each((i) => sumOver(p.accounts, i, "adjustment")) });
  }
  rows.push({ key: "close", label: lineLabel("close", lang), values: each((i) => sumOver(p.accounts, i, "close")), strong: true });
  return rows;
}

/** One account's rows, in the planner's order. */
export function accountRows(p: Portfolio, account: Account, lang: Lang): readonly SummaryRow[] {
  const keys: LineKey[] = ["open", "contribution", "withdrawal", "growth", "yieldTax", "fee"];
  if (p.hasAdjustment) keys.push("adjustment");
  keys.push("close");
  const after = account.meta.kind === "taxed";
  return keys.map((key) => ({
    key,
    label:
      key === "contribution"
        ? say(lang, after ? "Insättning, konto efter skatt" : "Insättning, konto före skatt", after ? "Contribution, after-tax account" : "Contribution, pre-tax account")
        : key === "withdrawal"
          ? say(lang, after ? "Uttag, konto efter skatt" : "Uttag, konto före skatt", after ? "Withdrawal, after-tax account" : "Withdrawal, pre-tax account")
          : lineLabel(key, lang),
    values: account.lines.map((l) => l[key]),
    ...(key === "open" || key === "close" ? { strong: true } : {}),
  }));
}

export function lineLabel(key: LineKey, lang: Lang): string {
  switch (key) {
    case "open":
      return say(lang, "Ingående saldo", "Opening balance");
    case "contribution":
      return say(lang, "Insättning", "Contribution");
    case "growth":
      return say(lang, "Avkastning", "Portfolio return");
    case "withdrawal":
      return say(lang, "Uttag", "Withdrawal");
    case "fee":
      return say(lang, "Förvaltningsavgift", "Management fee");
    case "yieldTax":
      return say(lang, "Avkastningsskatt", "Yield tax");
    case "adjustment":
      return say(lang, "Angivet saldo (justering)", "Balance entered (adjustment)");
    case "close":
      return say(lang, "Utgående saldo", "Closing balance");
  }
}

/** The all-accounts sum of one line, for the charts. */
export function totalOf(p: Portfolio, key: LineKey): readonly number[] {
  return p.years.map((_, i) => sumOver(p.accounts, i, key));
}

// --------------------------------------------------------------- exports ---

/** What the downloads hold: the summary, or one account, or every account. */
export type Selection = "all" | AccountId | "each";

function blocks(p: Portfolio, selection: Selection, lang: Lang): { title: string; rows: readonly SummaryRow[] }[] {
  const named = (a: Account) => `${accountName(a.id, lang, p.savingKind)} (${kindName(a.meta.kind, lang)})`;
  if (selection === "all") return [{ title: say(lang, "Alla konton (summa)", "All accounts (sum)"), rows: summaryRows(p, lang) }];
  const chosen = selection === "each" ? p.accounts : p.accounts.filter((a) => a.id === selection);
  return chosen.map((a) => ({ title: named(a), rows: accountRows(p, a, lang) }));
}

/** Withdrawals, taxes and fees go out as negative numbers, the planner's sign convention. */
const signed = (key: string, value: number): number =>
  key.startsWith("withdrawal") || key === "fee" || key === "yieldTax" ? -value : value;

export function portfolioToCsv(p: Portfolio, selection: Selection, lang: Lang): string {
  const d = delimiterFor(lang);
  const lines: string[] = [];
  for (const block of blocks(p, selection, lang)) {
    lines.push(csvLine([block.title], d));
    lines.push(csvLine([say(lang, "Ålder", "Age"), ...p.years.map((y) => String(y.age))], d));
    lines.push(csvLine([say(lang, "År", "Year"), ...p.years.map((y) => String(y.year))], d));
    for (const r of block.rows) {
      lines.push(csvLine([r.label, ...r.values.map((v) => kronor(Math.round(signed(r.key, v)) || 0, lang))], d));
    }
    lines.push("");
  }
  return lines.join("\r\n");
}

export function portfolioToXlsxRows(p: Portfolio, selection: Selection, lang: Lang): SheetData {
  const rows: SheetData = [];
  for (const block of blocks(p, selection, lang)) {
    rows.push([{ value: block.title, fontWeight: "bold" }]);
    rows.push([{ value: say(lang, "Ålder", "Age"), fontWeight: "bold" }, ...p.years.map((y) => ({ value: y.age, type: Number }))]);
    rows.push([{ value: say(lang, "År", "Year"), fontWeight: "bold" }, ...p.years.map((y) => ({ value: y.year, type: Number }))]);
    for (const r of block.rows) {
      rows.push([
        { value: r.label, ...(r.strong ? { fontWeight: "bold" as const } : {}) },
        ...r.values.map((v) => kronorCell(Math.round(signed(r.key, v)) || 0)),
      ]);
    }
    rows.push([null]);
  }
  return rows;
}
