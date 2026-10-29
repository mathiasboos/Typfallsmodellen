/**
 * The Portföljvärde tab: `portfolio.ts`'s figures on the page, laid out as the
 * planner's "Portfolio Value" tab lays them out -- an account picker, "Visa alla
 * konton", the Totalt / Skattekorrigerat switch and the two downloads in one row; the
 * balance chart and the flow chart side by side; then the all-accounts table, or one
 * table per account.
 *
 * It has no inputs of its own: it shows the run Prognos shows, so `main.ts` hands it
 * the result on every redraw. What it keeps is how it is being looked at.
 */
import type { ModelContext, TypfallResult } from "@typfallsmodellen/engine";

import { csvExportButton, xlsxExportButton } from "./export.js";
import { kronor, percent } from "./format.js";
import type { Lang } from "./i18n.js";
import { accountName, accountRows, buildPortfolio, kindName, portfolioToCsv, portfolioToXlsxRows, summaryRows, totalOf } from "./portfolio.js";
import type { Account, AccountId, Portfolio, PortfolioMode, Selection, SummaryRow } from "./portfolio.js";
import { renderPortfolioChart } from "./portfolioChart.js";
import type { PortfolioSeries } from "./portfolioChart.js";

const say = (l: Lang, sv: string, en: string) => (l === "sv" ? sv : en);

function h<K extends keyof HTMLElementTagNameMap>(tag: K, className?: string, text?: string): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text !== undefined) node.textContent = text;
  return node;
}

export interface PortfolioHandle {
  readonly element: HTMLElement;
  relabel(lang: Lang): void;
  /** Draws the tab for Prognos's run. */
  show(result: TypfallResult, context: ModelContext, par: number): void;
}

/** The rows whose amounts leave the account, shown with a minus. */
const OUTFLOWS = new Set(["withdrawal", "withdrawal-after", "withdrawal-before", "yieldTax", "fee"]);

function priceText(context: ModelContext, l: Lang): string {
  if (context.priceBasis === 1) return say(l, `fasta priser (${context.referenceYear})`, `constant prices (${context.referenceYear})`);
  if (context.priceBasis === 0) return say(l, `${context.referenceYear} års lönenivå`, `the ${context.referenceYear} wage level`);
  return say(l, "löpande priser", "current prices");
}

export function createPortfolioPanel(lang: Lang): PortfolioHandle {
  let currentLang = lang;
  let selection: "all" | AccountId = "all";
  let showEach = false;
  let mode: PortfolioMode = "total";
  let last: { result: TypfallResult; context: ModelContext; par: number } | undefined;

  const element = h("div", "pv");
  element.dataset.role = "portfolio";

  function draw(): void {
    if (last === undefined) return;
    const l = currentLang;
    const { result, context, par } = last;
    const p = buildPortfolio(result, par, mode, context.privateSavingKind);
    if (selection !== "all" && !p.accounts.some((a) => a.id === selection)) selection = "all";
    const exportSelection: Selection = showEach ? "each" : selection;

    const section = h("section", "panel");
    const head = h("h2");
    head.append(h("span", undefined, say(l, "Portföljvärde", "Portfolio value")));
    const actions = h("div", "panel-actions");
    const file = say(l, "portfoljvarde", "portfolio-value");
    actions.append(
      csvExportButton(l, `${file}.csv`, () => portfolioToCsv(p, exportSelection, l)),
      xlsxExportButton(l, `${file}.xlsx`, () => portfolioToXlsxRows(p, exportSelection, l)),
    );
    head.append(actions);
    const body = h("div", "pa-body");
    section.append(head, body);

    body.append(
      h(
        "p",
        "pa-lead",
        say(
          l,
          `Kontovis redovisning · ingående och utgående saldo per år. Samma person och inställningar som Prognos, i ${priceText(context, l)}.`,
          `Account by account · opening and closing balance for each year. The same person and settings as the Forecast, in ${priceText(context, l)}.`,
        ),
      ),
      controls(p, l),
    );

    if (!showEach) body.append(charts(p, l));

    if (showEach) for (const account of p.accounts) body.append(accountTable(p, account, l));
    else if (selection === "all") body.append(summaryTable(p, l));
    else body.append(accountTable(p, p.accounts.find((a) => a.id === selection)!, l));

    body.append(notes(p, l, par));
    element.replaceChildren(section);
    // The retirement year in the middle of each table, once the tables have a width.
    requestAnimationFrame(() => centreRetirement(element));
  }

  function controls(p: Portfolio, l: Lang): HTMLElement {
    const row = h("div", "pv-controls");

    const pick = h("label", "pv-pick");
    pick.append(h("span", "pa-pick-label", say(l, "Konto:", "Account:")));
    const select = h("select", "pv-select");
    select.dataset.role = "pv-account";
    const option = (value: string, text: string) => {
      const o = h("option", undefined, text);
      o.value = value;
      o.selected = value === selection;
      select.append(o);
    };
    option("all", say(l, "Alla konton (summa)", "All accounts (sum)"));
    for (const a of p.accounts) option(a.id, accountName(a.id, l, p.savingKind));
    select.disabled = showEach;
    select.addEventListener("change", () => {
      selection = select.value as "all" | AccountId;
      draw();
      element.querySelector<HTMLElement>("[data-role=pv-account]")?.focus();
    });
    pick.append(select);

    const each = h("label", "pv-each");
    const box = h("input");
    box.type = "checkbox";
    box.checked = showEach;
    box.dataset.role = "pv-each";
    box.addEventListener("change", () => {
      showEach = box.checked;
      draw();
      element.querySelector<HTMLElement>("[data-role=pv-each]")?.focus();
    });
    each.append(box, h("span", undefined, say(l, "Visa alla konton", "Show all accounts")));

    const toggle = h("div", "panel-toggle pv-mode");
    toggle.setAttribute("role", "group");
    for (const [value, text] of [
      ["total", say(l, "Totalt", "Total")],
      ["taxadj", say(l, "Skattekorrigerat", "Tax-adjusted")],
    ] as const) {
      const button = h("button", value === mode ? "panel-btn active" : "panel-btn", text);
      button.type = "button";
      button.dataset.mode = value;
      button.setAttribute("aria-pressed", String(value === mode));
      button.addEventListener("click", () => {
        mode = value;
        draw();
        element.querySelector<HTMLElement>(`.pv-mode [data-mode="${value}"]`)?.focus();
      });
      toggle.append(button);
    }

    row.append(pick, each, toggle);
    return row;
  }

  function charts(p: Portfolio, l: Lang): HTMLElement {
    const grid = h("div", "pv-charts");
    const years = p.years;
    const single = selection === "all" ? undefined : p.accounts.find((a) => a.id === selection);
    const subtitle = p.mode === "taxadj" ? say(l, "Skattekorrigerat", "Tax-adjusted") : undefined;

    const balance: PortfolioSeries[] = single
      ? [{ key: single.id, name: accountName(single.id, l, p.savingKind), colour: single.meta.colour, values: single.lines.map((x) => x.close) }]
      : p.accounts.map((a) => ({ key: a.id, name: accountName(a.id, l, p.savingKind), colour: a.meta.colour, values: a.lines.map((x) => x.close) }));
    grid.append(
      renderPortfolioChart({
        role: single ? `pv-fig-value-${single.id}` : "pv-fig-value",
        lang: l,
        title: single
          ? say(l, `Portföljvärde – ${accountName(single.id, l, p.savingKind)}`, `Portfolio value – ${accountName(single.id, l, p.savingKind)}`)
          : say(l, "Portföljvärde per konto", "Portfolio value by account"),
        ...(subtitle === undefined ? {} : { subtitle }),
        years,
        retirementAge: p.retirementAge,
        kind: single ? "area" : "stack",
        series: balance,
        totalName: say(l, "Summa", "Total"),
      }),
    );

    const pick = (key: "contribution" | "growth" | "withdrawal" | "yieldTax" | "fee"): readonly number[] =>
      single ? single.lines.map((x) => x[key]) : totalOf(p, key);
    const flows: PortfolioSeries[] = ([
      { key: "contribution", name: say(l, "Insättning", "Contribution"), colour: "--fig-pv-in", values: pick("contribution") },
      { key: "growth", name: say(l, "Avkastning", "Portfolio return"), colour: "--fig-pv-return", values: pick("growth") },
      { key: "withdrawal", name: say(l, "Uttag", "Withdrawal"), colour: "--fig-pv-out", values: pick("withdrawal").map((v) => -v) },
      { key: "yieldTax", name: say(l, "Avkastningsskatt", "Yield tax"), colour: "--fig-pv-tax", values: pick("yieldTax").map((v) => -v) },
      { key: "fee", name: say(l, "Förvaltningsavgift", "Management fee"), colour: "--fig-pv-fee", values: pick("fee").map((v) => -v) },
    ] as PortfolioSeries[]).filter((series) => series.values.some((v) => Math.abs(v) >= 0.5));
    grid.append(
      renderPortfolioChart({
        role: single ? `pv-fig-flow-${single.id}` : "pv-fig-flow",
        lang: l,
        title: say(l, "Årsvis in- och utflöde", "Annual in- and outflow"),
        ...(subtitle === undefined ? {} : { subtitle }),
        years,
        retirementAge: p.retirementAge,
        kind: "stack",
        series: flows,
      }),
    );
    return grid;
  }

  /** A card with a heading and a year-by-year table: the first column stays put while the years scroll. */
  function card(title: string, badge: string | undefined, rows: readonly SummaryRow[], p: Portfolio, role: string, l: Lang): HTMLElement {
    const box = h("div", "pv-card");
    box.dataset.role = role;
    const heading = h("h3", "pv-card-title", title);
    if (badge !== undefined) heading.append(h("span", "pv-badge", badge));
    box.append(heading);

    const wrap = h("div", "scroll pv-scroll");
    const table = h("table", "table pv-table");
    const thead = h("thead");
    const ageRow = h("tr");
    const yearRow = h("tr");
    ageRow.append(h("th", "pv-label", say(l, "Ålder", "Age")));
    yearRow.append(h("th", "pv-label", say(l, "År", "Year")));
    for (const yr of p.years) {
      const ret = yr.age === p.retirementAge;
      const a = h("th", ret ? "pv-ret" : undefined, String(yr.age));
      const y = h("th", ret ? "pv-ret pv-ret-year" : undefined, String(yr.year));
      if (ret) y.append(h("span", "pv-star", ` ★ ${say(l, "Pension", "Retirement")}`));
      ageRow.append(a);
      yearRow.append(y);
    }
    thead.append(ageRow, yearRow);
    const tbody = h("tbody");
    for (const r of rows) {
      const tr = h("tr", r.strong ? "pv-strong" : undefined);
      tr.dataset.row = r.key;
      const name = h("th", "pv-label", r.label);
      name.scope = "row";
      if (r.title !== undefined) name.title = r.title;
      tr.append(name);
      r.values.forEach((v, i) => {
        const rounded = Math.round(v);
        const out = OUTFLOWS.has(r.key) && rounded !== 0;
        const td = h("td", `num${rounded === 0 ? " pv-zero" : ""}${p.years[i]!.age === p.retirementAge ? " pv-ret" : ""}`);
        td.textContent = out ? `−${kronor(Math.abs(rounded), l)}` : kronor(rounded || 0, l);
        tr.append(td);
      });
      tbody.append(tr);
    }
    table.append(thead, tbody);
    wrap.append(table);
    box.append(wrap);
    return box;
  }

  function summaryTable(p: Portfolio, l: Lang): HTMLElement {
    return card(say(l, "Alla konton", "All accounts"), say(l, "summa av alla konton", "sum of all accounts"), summaryRows(p, l), p, "pv-summary", l);
  }

  function accountTable(p: Portfolio, account: Account, l: Lang): HTMLElement {
    return card(accountName(account.id, l, p.savingKind), kindName(account.meta.kind, l), accountRows(p, account, l), p, `pv-account-${account.id}`, l);
  }

  function notes(p: Portfolio, l: Lang, par: number): HTMLElement {
    const list = h("ul", "pa-notes pv-notes");
    const add = (sv: string, en: string) => list.append(h("li", undefined, say(l, sv, en)));
    if (p.mode === "taxadj") {
      add(
        `Skattekorrigerat: konton vars uttag beskattas som inkomst visas efter den genomsnittliga skatten på pensionen vid ${Math.trunc(par)} år i Prognos, ${percent(p.taxRate, l)}. Privat sparande i kapitalförsäkring, ISK eller depå tas ut skattefritt och visas som det är.`,
        `Tax-adjusted: accounts whose withdrawals are taxed as income are shown after the average tax on the pension at ${Math.trunc(par)} in the Forecast, ${percent(p.taxRate, l)}. Private saving in an endowment insurance, an ISK or a securities account is paid out tax-free and shown as it is.`,
      );
    }
    add(
      "Uttagen är samma belopp som Prognos visar utbetalda per år (Tabell 2). Avkastningen är kontots förändring utöver insättningar, uttag, skatt och avgift: avkastning, indexering, arvsvinster och ränta.",
      "The withdrawals are the amounts the Forecast shows paid each year (Table 2). The return is the account's change beyond contributions, withdrawals, tax and fee: return, indexation, inheritance gains and interest.",
    );
    if (p.accounts.some((a) => a.id === "ip")) add(
      "Inkomstpensionen görs om till en livsvarig pension när den börjar betalas ut, och modellen har inget saldo för den efter det. Från pensionsåret visas i stället vad den pension som återstår är värd: nästa års pension gånger delningstalet vid nästa års ålder.",
      "The income pension is turned into a pension for life when it starts to be paid, and the model keeps no balance for it after that. From the retirement year the tab shows what the pension still to come is worth instead: next year's pension times the annuity divisor at next year's age.",
    );
    add(
      "En livsvarig pension betalas så länge man lever, även när kontots saldo i modellen har tagit slut. Garantipension, inkomstpensionstillägg, tilläggspension och förmånsbestämd tjänstepension har inget saldo och ingår inte.",
      "A pension for life is paid for as long as one lives, even once the account's balance in the model has run out. Guarantee pension, income pension supplement, supplementary pension and defined-benefit occupational pension have no balance and are not included.",
    );
    return list;
  }

  return {
    element,
    // `main.ts` redraws after a language change, and only the tab on screen is shown.
    relabel(l: Lang) {
      currentLang = l;
    },
    show(result, context, par) {
      last = { result, context, par };
      draw();
    },
  };
}

/** Scrolls each table so the retirement year's column sits mid-way across what the label column leaves. */
function centreRetirement(root: HTMLElement): void {
  for (const scroller of root.querySelectorAll<HTMLElement>(".pv-scroll")) {
    const target = scroller.querySelector<HTMLElement>("thead tr:last-child th.pv-ret");
    const label = scroller.querySelector<HTMLElement>("thead th.pv-label");
    if (!target || !label || scroller.clientWidth === 0) continue;
    const frozen = label.getBoundingClientRect().width;
    const box = scroller.getBoundingClientRect();
    const cell = target.getBoundingClientRect();
    const offset = cell.left - box.left + cell.width / 2 - (frozen + (scroller.clientWidth - frozen) / 2);
    scroller.scrollLeft = Math.max(0, scroller.scrollLeft + offset);
  }
}
