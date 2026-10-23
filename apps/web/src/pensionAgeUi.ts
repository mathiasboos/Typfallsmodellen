/**
 * The small pieces "Hur maxa pensionen?" builds its panels from, shared by the
 * tab's own panel (`pensionAgePanel.ts`) and by the agreement comparison
 * (`pensionAgeCompare.ts`): element and table builders, the number formats the
 * two use, and the few phrases both need to describe a study ("fasta priser
 * (2025)", "allmän pension, före skatt"). "Hur länge räcker pensionen?"
 * (`pensionLastsPanel.ts`) builds its own panels from the same element, table and
 * number helpers; the phrases and the line keys are about retirement ages and stay
 * here.
 *
 * Nothing here holds state. It lives in a module of its own because the two
 * panels would otherwise have to import each other.
 */
import { kronor, locale } from "./format.js";
import { t } from "./i18n.js";
import type { Lang } from "./i18n.js";
import { kindsPaid } from "./pensionAge.js";
import type { Study, StudyOption } from "./pensionAge.js";
import { cell } from "./tables.js";

export const say = (l: Lang, sv: string, en: string) => (l === "sv" ? sv : en);

export function h<K extends keyof HTMLElementTagNameMap>(tag: K, className?: string, text?: string): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text !== undefined) node.textContent = text;
  return node;
}

export function decimals(value: number, digits: number, l: Lang): string {
  return new Intl.NumberFormat(locale(l), { minimumFractionDigits: digits, maximumFractionDigits: digits }).format(value);
}

/** Thousands of kronor, no decimals. */
export const tkr = (value: number, l: Lang): string => kronor(Math.round(value / 1000), l);
export const signed = (text: string, value: number): string => (value > 0 ? `+${text}` : text);

/** An amount in millions of kronor, two decimals. */
export const mkr = (value: number, l: Lang): string => decimals(value / 1e6, 2, l);

export const capital = (text: string): string => text.charAt(0).toUpperCase() + text.slice(1);

/** A section with the sheet's own gold title bar, the way `main.ts`'s `section()` builds one. */
export function panel(title: string, actions: readonly HTMLElement[] = []): { section: HTMLElement; body: HTMLElement } {
  const section = h("section", "panel");
  const head = h("h2");
  head.append(h("span", undefined, title));
  if (actions.length > 0) {
    const box = h("div", "panel-actions");
    box.append(...actions);
    head.append(box);
  }
  const body = h("div", "pa-body");
  section.append(head, body);
  return { section, body };
}

export function tableOf(
  role: string,
  extra: string,
  head: readonly HTMLTableCellElement[],
  rows: readonly HTMLTableRowElement[],
): HTMLElement {
  const wrap = h("div", "scroll");
  const table = h("table", `table pa-table ${extra}`);
  table.dataset.role = role;
  const thead = h("thead");
  const headRow = h("tr");
  headRow.append(...head);
  thead.append(headRow);
  const tbody = h("tbody");
  tbody.append(...rows);
  table.append(thead, tbody);
  wrap.append(table);
  return wrap;
}

export function row(cells: readonly HTMLTableCellElement[], data: Record<string, string> = {}): HTMLTableRowElement {
  const tr = h("tr");
  for (const [key, value] of Object.entries(data)) tr.dataset[key] = value;
  tr.append(...cells);
  return tr;
}

/** A table cell carrying an alternative's line key beside its label. */
export function optionCell(option: StudyOption, text: string): HTMLTableCellElement {
  const td = cell("");
  td.append(h("span", `pa-key pa-key-${option.slot}`), document.createTextNode(text));
  return td;
}

/** A box with a retirement age in it, coloured by how far it sits from the reference age. */
export function chip(option: StudyOption): HTMLElement {
  return h("span", `pa-chip pa-chip-${option.slot}`, String(option.par));
}

// ---------------------------------------------------------------- phrases ---

/** What the reference age is called: the riktålder when it is one, "the reference age" when it is not. */
export const refName = (study: Study, l: Lang): string =>
  study.options[study.reference]?.par === study.riktalder
    ? say(l, "riktåldern", `the ${t("riktalderCheckbox", l).toLowerCase()}`)
    : say(l, "referensåldern", "the reference age");

export const optionName = (study: Study, option: StudyOption, l: Lang): string => {
  const years = say(l, "år", "yrs");
  if (option.offset === 0) return `${option.par} ${years} (${refName(study, l).replace(/^the /, "")})`;
  return `${option.par} ${years} (${option.offset > 0 ? "+" : "−"}${Math.abs(option.offset)})`;
};

export const priceText = (study: Study, l: Lang): string => {
  const year = study.referenceYear;
  if (study.priceBasis === 1) return say(l, `fasta priser (${year})`, `constant prices (${year})`);
  if (study.priceBasis === 0) return say(l, `dagens (${year}) lönenivå`, `current (${year}) wage level`);
  return say(l, "löpande priser", "current prices");
};

/** What a study adds up, in a few words: "allmän pension + tjänstepension + eget sparande, före skatt". */
export const measureText = (study: Study, l: Lang): string => {
  const kinds = kindsPaid(study);
  const kind = kinds.occupational
    ? say(l, "allmän pension + tjänstepension", "public + occupational pension")
    : say(l, "allmän pension", "public pension");
  const own = kinds.individual || kinds.savings ? say(l, " + eget sparande", " + own saving") : "";
  const tax = study.input.afterTax ? say(l, "efter skatt", "after tax") : say(l, "före skatt", "before tax");
  return `${kind}${own}, ${tax}`;
};
