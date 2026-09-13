/**
 * Every visible string comes from the workbook.
 *
 * `packages/data/i18n.json` carries 481 labels, 50 messages and 129 captions in
 * Swedish and English, extracted from the SysLang sheet. They are keyed by that
 * sheet's row number, which is meaningful to the workbook and to nobody else --
 * so this file gives each one the name the code uses, with the Swedish text
 * beside it. A reviewer can check any line against the sheet without leaving
 * the file, and next year's extraction cannot silently renumber a string
 * without `checkLabels` below noticing.
 *
 * Where the workbook's own text embeds a number that this port computes -- "2025
 * ars lonenivå", "Tabell 2. Manadsinkomster fran 56 alder" -- the label is not
 * used: those are composed from the run instead, so they stay true when the
 * inputs change.
 */
import { i18n } from "@typfallsmodellen/data";

export type Lang = "sv" | "en";
export const LANGS: readonly Lang[] = ["sv", "en"];

type Entry = Partial<Record<Lang, string>>;
const LABELS = i18n.labels as Record<string, Entry>;

/** Label key → the SysLang row it came from, with that row's Swedish text. */
export const L = {
  chooseLanguage: "2", //   Välj språk
  inputs: "4", //           Ingångsvärden
  birthYear: "5", //        Födelseår
  retirementAge: "6", //    Går i pension vid ålder
  startWorkAge: "7", //     Börjar arbeta vid ålder
  retires: "8", //          Går i pension
  beforeEarliest: "9", //   Fel kan inte beräkna före tidigaste uttagsålder
  annualSalary: "10", //    Årslön
  yearsOld: "11", //        år gammal
  married: "14", //         Gift
  inflation: "15", //       Årlig inflation
  realReturn: "16", //      Real avkastning
  realGrowth: "18", //      Real tillväxt
  occupational: "19", //    Välj tjänstepension
  currentPrices: "21", //   Löpande priser
  fixedPrices: "22", //     Fasta priser (2025)
  table1: "24", //          Tabell 1. Specificerat resultat över slutlön och pensionsinkomster
  year: "27", //            År
  pension: "28", //         Pension
  yearsAge: "29", //        års ålder
  finalSalary: "31", //     Slutlön
  salaryAfterTax: "32", //  Lön efter skatt
  disposable: "33", //      Disponibel inkomst
  totalGross: "34", //      Total pension brutto
  totalPublic: "35", //     Total allmän pension
  kronor: "37", //          kronor
  show: "38", //            Visa
  years: "39", //           år
  occupationalShort: "40", // Tjänstepension
  privateSaving: "41", //   Privat pensionssparande (med avdragsrätt)
  shareOfFinalSalary: "44", // Som andel av slutlön i fasta priser
  perMonth: "45", //        Per månad
  afterTax: "46", //        Efter skatt
  kpiFootnote: "48", //     *Mätt med årsmedeltal för KPI
  age: "51", //             Ålder
  salary: "52", //          Lön
  incomeAndSupplementary: "53", // Inkomst- och tilläggspension
  premium: "54", //         Premiepension
  occupationalPlusIps: "55", //  Tjänste-pension+IPS
  guarantee: "56", //       Garantipension
  grossIncome: "57", //     Inkomst brutto
  incomeAfterTax: "58", //  Inkomst efter skatt
  benefits: "59", //        Bidrag (BT, ÄFS, m.m.)
  incomePension: "242", //  Inkomstpension
  supplementary: "330", //  Tilläggspension
  iptFull: "481", //        Pensionstillägg (IPT) (40/40)
} as const;

export type LabelName = keyof typeof L;

/** The Swedish text each key is expected to hold, so a renumbering is caught. */
const EXPECTED_SV: Partial<Record<LabelName, string>> = {
  birthYear: "Födelseår",
  retirementAge: "Går i pension vid ålder",
  startWorkAge: "Börjar arbeta vid ålder",
  annualSalary: "Årslön",
  occupational: "Välj tjänstepension",
  finalSalary: "Slutlön",
  age: "Ålder",
  premium: "Premiepension",
  guarantee: "Garantipension",
  incomePension: "Inkomstpension",
  supplementary: "Tilläggspension",
};

/**
 * Checks the keys still point at the strings this file claims.
 *
 * Cheap insurance for the yearly update: if next year's SysLang sheet gains a
 * row, every key after it shifts by one and the UI would relabel itself with
 * plausible-looking nonsense. Run from the tests.
 */
export function checkLabels(): string[] {
  const wrong: string[] = [];
  for (const [name, sv] of Object.entries(EXPECTED_SV) as [LabelName, string][]) {
    const got = LABELS[L[name]]?.sv;
    if (got !== sv) wrong.push(`${name} (key ${L[name]}): expected "${sv}", found "${got ?? ""}"`);
  }
  return wrong;
}

/**
 * The label in the chosen language.
 *
 * Many of the workbook's advanced-settings labels have no English text, so a
 * missing translation falls back to Swedish rather than showing a key. Normal
 * mode's labels are all translated; the fallback is for the ones Phase 4 reaches.
 */
export function t(name: LabelName, lang: Lang): string {
  const entry = LABELS[L[name]];
  return entry?.[lang] || entry?.sv || "";
}
