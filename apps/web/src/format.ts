/**
 * Numbers as the workbook writes them: space-grouped thousands, no decimals on
 * kronor. `Intl` does the work, so both locales get their own conventions.
 */
import type { Lang } from "./i18n.js";

const LOCALE: Record<Lang, string> = { sv: "sv-SE", en: "en-GB" };

export function kronor(value: number, lang: Lang): string {
  return new Intl.NumberFormat(LOCALE[lang], { maximumFractionDigits: 0 }).format(value);
}

export function percent(value: number, lang: Lang, digits = 1): string {
  return new Intl.NumberFormat(LOCALE[lang], {
    style: "percent",
    minimumFractionDigits: digits,
    maximumFractionDigits: digits,
  }).format(value);
}

export function locale(lang: Lang): string {
  return LOCALE[lang];
}
