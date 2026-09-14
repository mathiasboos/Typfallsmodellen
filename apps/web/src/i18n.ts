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
// Keyed by identifier rather than SysLang row -- these are Start-sheet control
// captions (checkbox and button text), extracted separately from the numbered
// labels but looked up the same way below.
const CAPTIONS = i18n.captions as Record<string, Entry>;

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
  // table1 (SysLang row 24, "Tabell 1. Specificerat resultat över slutlön och
  // pensionsinkomster") is not used: Table 1's title was replaced outright with
  // "Pensionsinkomst" on request, in main.ts, rather than composed from it.
  table2: "49", //          Tabell 2. Månadsinkomster från 56 ålder
  year: "27", //            År
  pension: "28", //         Pension
  yearsAge: "29", //        års ålder
  finalSalary: "31", //     Slutlön
  salaryAfterTax: "32", //  Lön efter skatt
  disposable: "33", //      Disponibel inkomst
  totalGross: "34", //      Total pension brutto
  totalPublic: "35", //     Total allmän pension
  kronor: "37", //          kronor
  kr: "476", //             kr
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

  // Table 1's own headings, footnotes and the three rows the first draft left
  // out. The letters A)-D) are typed into the sheet's header cells rather than
  // stored in SysLang, so they are composed in tables.ts.
  at: "47", //              vid
  pensionWord: "28", //     Pension
  currentPricesKronor: "42", // Löpande priser, kronor
  shareOfFinalSalaryShort: "393", // Som andel av slutlön
  shareOfIncomeBefore: "479", // Som andel av inkomsten året innan
  housingSupplement: "185", // Bostadstillägg för pensionärer m.m.
  privateSavingIsk: "485", // Privat pensionssparande (ISK / KF)
  tableAboveShows: "456", // Tabellen ovan visar värden
  lastRightShort: "64", //  Inkl sista pensionsrätten som ... medräknas först året efter.
  occupationalLifelong: "457", // Tjänstepension betalas ut livsvarigt

  // The figures. These are the rows `Data_till_Start` itself looks up for each
  // chart's series names -- Y2:Y10 for Figur 2, AN2:AN6 for the disposable
  // income chart, N1:T2 for Figur 1 -- so they follow the workbook's choice
  // even where an equivalent string exists on another row.
  figure1: "68", //         Figur 1. Löneinkomst mellan 22
  andPensionFrom: "420", // och pension från
  fixedPricesNote: "439", // Fasta priser (reala priser) - priset justerat för inflation
  currentPricesNote: "440", // Löpande priser (nominella priser) - priset anges i aktuell prisnivå
  figure2: "374", //        Figur 2. Månadsinkomster: 56 - 76 års ålder
  todayWageLevel: "465", // Dagens (2025) lönenivå
  earningsAndPension: "335", // Lön och pension
  salaryChart: "422", //    Lön
  continuedWork: "328", //  Lön vid fortsatt arbete
  occupationalChart: "332", // Tjänstepension
  grossIncomeChart: "333", // Bruttoinkomst
  netIncomeChart: "334", // Nettoinkomst

  // A monthly rather than annual salary field, and the riktålder checkbox --
  // both requested directly rather than following the Start sheet's own
  // Årslön field and its `rng_Riktålder` checkbox one-for-one.
  monthlySalary: "435", // Månadslön
  riktalderCheckbox: "chkRecPensAge", // Riktålder (a control caption, not a SysLang row)

  // The Table 2 scale toggle, moved into its header on request. Distinct rows
  // from `perMonth` (45, "Per månad"), which still heads Table 1's column C.
  yearlyView: "464", //     Årsvis
  monthlyView: "463", //    Månadsvis
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
  currentPricesKronor: "Löpande priser, kronor",
  shareOfIncomeBefore: "Som andel av inkomsten året innan",
  housingSupplement: "Bostadstillägg för pensionärer m.m.",
  privateSavingIsk: "Privat pensionssparande (ISK / KF)",
  todayWageLevel: "Dagens (2025) lönenivå",
  earningsAndPension: "Lön och pension",
  continuedWork: "Lön vid fortsatt arbete",
  grossIncomeChart: "Bruttoinkomst",
  netIncomeChart: "Nettoinkomst",
  monthlySalary: "Månadslön",
  riktalderCheckbox: "Riktålder",
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
    const got = LABELS[L[name]]?.sv ?? CAPTIONS[L[name]]?.sv;
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
  const entry = LABELS[L[name]] ?? CAPTIONS[L[name]];
  return entry?.[lang] || entry?.sv || "";
}

/**
 * Strips the sheet's own "Tabell 2.", "Figur 1.", "Table 1.", "Chart 2." lead-in
 * off a heading, on request: Table 2 and both figures keep the rest of the
 * SysLang text, just not the number the sheet gives itself.
 */
export function dropHeadingNumber(text: string): string {
  return text.replace(/^(Tabell|Tabel|Table|Figur|Figure|Chart)\s*\d+\.?\s*/i, "");
}
