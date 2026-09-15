import { municipalTax } from "@typfallsmodellen/data";

/**
 * Sweden's 290 total municipal + regional tax rates, and the two church/burial
 * defaults the engine already extracted but never had a UI path to.
 *
 * === The municipality table ===
 *
 * Not from the workbook. `packages/data/municipal-tax.json` (extracted from
 * the workbook's own `K_skatt` sheet) is a national *average* -- there is no
 * per-municipality table anywhere in this repo to reuse, so this one is new,
 * the same class of addition as the tax-per-year chart or the KPI cards.
 *
 * Copied verbatim from a reference calculator the user supplied, which cites:
 * Statistiska centralbyrån (SCB), "Totala kommunala skattesatser, kommunvis"
 * (2025-12-15) -- https://www.scb.se/hitta-statistik/statistik-efter-amne/
 * offentlig-ekonomi/finanser-for-den-kommunala-sektorn/kommunalskatterna/pong/
 * tabell-och-diagram/totala-kommunala-skattesatser-2026-kommunvis/
 *
 * The 290 rates themselves are trusted from that source, not independently
 * re-verified against SCB's own table. A rate here is a snapshot for the
 * stated year -- next year's list will differ, the way `w_ref`'s reference
 * year already dates other assumptions in this port. Update by replacing the
 * whole table at the next release.
 */
export const KOMMUNALSKATT_YEAR = 2026;

/** Percent, not a fraction -- matches how the percent field displays it. */
export const KOMMUNALSKATT: Readonly<Record<string, number>> = {
  "Ale": 32.8,
  "Alingsås": 32.84,
  "Alvesta": 33.42,
  "Aneby": 33.85,
  "Arboga": 33.29,
  "Arjeplog": 34.84,
  "Arvidsjaur": 34.14,
  "Arvika": 34.03,
  "Askersund": 34.15,
  "Avesta": 33.95,
  "Bengtsfors": 34.4,
  "Berg": 34.22,
  "Bjurholm": 35.0,
  "Bjuv": 32.17,
  "Boden": 33.94,
  "Bollebygd": 33.07,
  "Bollnäs": 33.37,
  "Borgholm": 33.44,
  "Borlänge": 34.4,
  "Borås": 32.79,
  "Botkyrka": 32.23,
  "Boxholm": 33.37,
  "Bromölla": 33.74,
  "Bräcke": 35.09,
  "Burlöv": 31.27,
  "Båstad": 31.41,
  "Dals-Ed": 34.69,
  "Danderyd": 30.58,
  "Degerfors": 35.3,
  "Dorotea": 35.65,
  "Eda": 34.55,
  "Ekerö": 31.45,
  "Eksjö": 34.02,
  "Emmaboda": 33.82,
  "Enköping": 33.05,
  "Eskilstuna": 32.85,
  "Eslöv": 31.72,
  "Essunga": 33.05,
  "Fagersta": 32.99,
  "Falkenberg": 32.5,
  "Falköping": 33.43,
  "Falun": 34.05,
  "Filipstad": 34.55,
  "Finspång": 33.7,
  "Flen": 33.1,
  "Forshaga": 34.63,
  "Färgelanda": 34.39,
  "Gagnef": 34.27,
  "Gislaved": 33.75,
  "Gnesta": 32.95,
  "Gnosjö": 34.0,
  "Gotland": 33.6,
  "Grums": 34.78,
  "Grästorp": 33.47,
  "Gullspång": 33.97,
  "Gällivare": 33.89,
  "Gävle": 33.77,
  "Göteborg": 32.6,
  "Götene": 33.6,
  "Habo": 33.93,
  "Hagfors": 34.3,
  "Hallsberg": 33.85,
  "Hallstahammar": 32.69,
  "Halmstad": 32.38,
  "Hammarö": 35.05,
  "Haninge": 31.28,
  "Haparanda": 33.84,
  "Heby": 34.21,
  "Hedemora": 34.15,
  "Helsingborg": 31.39,
  "Herrljunga": 33.62,
  "Hjo": 33.48,
  "Hofors": 34.37,
  "Huddinge": 31.71,
  "Hudiksvall": 33.12,
  "Hultsfred": 33.77,
  "Hylte": 33.85,
  "Hällefors": 34.35,
  "Härjedalen": 34.17,
  "Härnösand": 34.63,
  "Härryda": 31.98,
  "Hässleholm": 32.38,
  "Håbo": 33.3,
  "Höganäs": 30.91,
  "Högsby": 34.07,
  "Hörby": 32.26,
  "Höör": 32.63,
  "Jokkmokk": 34.29,
  "Järfälla": 31.52,
  "Jönköping": 33.4,
  "Kalix": 33.89,
  "Kalmar": 33.67,
  "Karlsborg": 32.8,
  "Karlshamn": 33.8,
  "Karlskoga": 34.3,
  "Karlskrona": 33.69,
  "Karlstad": 33.55,
  "Katrineholm": 32.95,
  "Kil": 34.63,
  "Kinda": 33.0,
  "Kiruna": 34.39,
  "Klippan": 31.93,
  "Knivsta": 32.62,
  "Kramfors": 34.43,
  "Kristianstad": 32.64,
  "Kristinehamn": 34.25,
  "Krokom": 33.87,
  "Kumla": 33.84,
  "Kungsbacka": 32.58,
  "Kungsör": 32.91,
  "Kungälv": 32.92,
  "Kävlinge": 29.59,
  "Köping": 33.04,
  "Laholm": 32.8,
  "Landskrona": 31.42,
  "Laxå": 35.0,
  "Lekeberg": 33.73,
  "Leksand": 33.8,
  "Lerum": 32.03,
  "Lessebo": 33.81,
  "Lidingö": 29.67,
  "Lidköping": 32.74,
  "Lilla Edet": 33.85,
  "Lindesberg": 34.6,
  "Linköping": 31.75,
  "Ljungby": 33.07,
  "Ljusdal": 33.87,
  "Ljusnarsberg": 33.8,
  "Lomma": 30.72,
  "Ludvika": 34.05,
  "Luleå": 33.84,
  "Lund": 32.42,
  "Lycksele": 34.9,
  "Lysekil": 33.94,
  "Malmö": 32.42,
  "Malung-Sälen": 34.45,
  "Malå": 35.2,
  "Mariestad": 32.74,
  "Mark": 32.99,
  "Markaryd": 33.31,
  "Mellerud": 34.08,
  "Mjölby": 33.45,
  "Mora": 34.32,
  "Motala": 33.25,
  "Mullsjö": 34.1,
  "Munkedal": 34.86,
  "Munkfors": 34.3,
  "Mölndal": 31.99,
  "Mönsterås": 34.07,
  "Mörbylånga": 34.07,
  "Nacka": 30.11,
  "Nora": 34.55,
  "Norberg": 33.54,
  "Nordanstig": 34.02,
  "Nordmaling": 35.1,
  "Norrköping": 33.3,
  "Norrtälje": 32.05,
  "Norsjö": 35.2,
  "Nybro": 34.19,
  "Nykvarn": 32.3,
  "Nyköping": 32.25,
  "Nynäshamn": 32.18,
  "Nässjö": 34.3,
  "Ockelbo": 34.27,
  "Olofström": 33.75,
  "Orsa": 34.3,
  "Orust": 33.69,
  "Osby": 33.99,
  "Oskarshamn": 34.21,
  "Ovanåker": 33.37,
  "Oxelösund": 33.05,
  "Pajala": 34.74,
  "Partille": 31.36,
  "Perstorp": 31.99,
  "Piteå": 33.59,
  "Ragunda": 34.92,
  "Robertsfors": 35.0,
  "Ronneby": 33.68,
  "Rättvik": 33.8,
  "Sala": 33.19,
  "Salem": 32.0,
  "Sandviken": 33.12,
  "Sigtuna": 31.83,
  "Simrishamn": 31.69,
  "Sjöbo": 32.1,
  "Skara": 33.38,
  "Skellefteå": 34.45,
  "Skinnskatteberg": 33.34,
  "Skurup": 31.6,
  "Skövde": 33.09,
  "Smedjebacken": 34.44,
  "Sollefteå": 34.68,
  "Sollentuna": 30.45,
  "Solna": 29.7,
  "Sorsele": 35.45,
  "Sotenäs": 33.47,
  "Staffanstorp": 30.12,
  "Stenungsund": 33.12,
  "Stockholm": 30.55,
  "Storfors": 34.98,
  "Storuman": 34.95,
  "Strängnäs": 32.5,
  "Strömstad": 33.39,
  "Strömsund": 34.92,
  "Sundbyberg": 31.58,
  "Sundsvall": 33.88,
  "Sunne": 33.75,
  "Surahammar": 33.19,
  "Svalöv": 31.92,
  "Svedala": 31.42,
  "Svenljunga": 33.53,
  "Säffle": 33.8,
  "Säter": 34.3,
  "Sävsjö": 33.68,
  "Söderhamn": 33.17,
  "Söderköping": 33.53,
  "Södertälje": 32.38,
  "Sölvesborg": 33.86,
  "Tanum": 33.04,
  "Tibro": 33.19,
  "Tidaholm": 33.55,
  "Tierp": 33.0,
  "Timrå": 34.38,
  "Tingsryd": 34.0,
  "Tjörn": 33.19,
  "Tomelilla": 31.79,
  "Torsby": 34.3,
  "Torsås": 33.79,
  "Tranemo": 32.98,
  "Tranås": 33.77,
  "Trelleborg": 31.58,
  "Trollhättan": 33.84,
  "Trosa": 32.03,
  "Tyresö": 31.83,
  "Täby": 29.88,
  "Töreboda": 33.2,
  "Uddevalla": 33.64,
  "Ulricehamn": 32.53,
  "Umeå": 34.65,
  "Upplands Väsby": 31.75,
  "Upplands-Bro": 31.73,
  "Uppsala": 32.85,
  "Uppvidinge": 33.8,
  "Vadstena": 34.35,
  "Vaggeryd": 33.25,
  "Valdemarsvik": 34.03,
  "Vallentuna": 31.23,
  "Vansbro": 34.28,
  "Vara": 33.25,
  "Varberg": 31.73,
  "Vaxholm": 31.63,
  "Vellinge": 29.68,
  "Vetlanda": 33.77,
  "Vilhelmina": 35.5,
  "Vimmerby": 34.22,
  "Vindeln": 35.2,
  "Vingåker": 33.5,
  "Vänersborg": 33.69,
  "Vännäs": 35.2,
  "Värmdö": 31.31,
  "Värnamo": 33.28,
  "Västervik": 33.02,
  "Västerås": 31.24,
  "Växjö": 32.19,
  "Vårgårda": 33.09,
  "Ydre": 34.1,
  "Ystad": 31.29,
  "Älmhult": 33.86,
  "Älvdalen": 34.77,
  "Älvkarleby": 34.4,
  "Älvsbyn": 33.79,
  "Ängelholm": 31.35,
  "Åmål": 33.94,
  "Ånge": 34.62,
  "Åre": 33.92,
  "Årjäng": 34.25,
  "Åsele": 35.45,
  "Åstorp": 31.47,
  "Åtvidaberg": 33.94,
  "Öckerö": 33.04,
  "Ödeshög": 33.95,
  "Örebro": 33.65,
  "Örkelljunga": 30.24,
  "Örnsköldsvik": 33.85,
  "Östersund": 33.72,
  "Österåker": 28.93,
  "Östhammar": 33.3,
  "Östra Göinge": 32.17,
  "Överkalix": 34.14,
  "Övertorneå": 33.84,
};

// === The church/burial defaults ==========================================

interface TaxSeries {
  readonly firstYear: number;
  readonly lastYear: number;
  readonly series: Record<string, { readonly lastActualYear: number | null; readonly values: readonly (number | null)[] }>;
}
const TAX = municipalTax as unknown as TaxSeries;

/**
 * The latest year a K_skatt column actually carries a number, as a fraction.
 *
 * `kyrkoavgift`'s own `lastActualYear` metadata reads 2120 -- `begravnings-
 * avgift`'s own last real value, held flat by a projection formula the sheet
 * applies to every column but this one -- despite `kyrkoavgift` itself having
 * nothing typed past 2026 (its values run out to `null` there). The engine
 * has never read this column (`grep -rln kyrkoavgift packages/engine/src`
 * finds nothing), so nobody has had reason to notice or fix it there.
 *
 * Clamping `lastActualYear` to the last age the column actually holds a
 * number gets both series right without treating either as a special case:
 * `kyrkoavgift` clamps down to 2026 (correct), `begravningsavgift`'s own
 * claim of 2026 already sits before its mechanically-held-flat tail so the
 * clamp changes nothing (also correct).
 */
function latestRate(name: "kyrkoavgift" | "begravningsavgift"): { readonly year: number; readonly rate: number } {
  const s = TAX.series[name]!;
  let lastNonNull = s.values.length - 1;
  while (lastNonNull >= 0 && (s.values[lastNonNull] === null || s.values[lastNonNull] === undefined)) {
    lastNonNull -= 1;
  }
  const claimedIndex = (s.lastActualYear ?? TAX.firstYear + lastNonNull) - TAX.firstYear;
  const index = Math.min(claimedIndex, lastNonNull);
  return { year: TAX.firstYear + index, rate: (s.values[index] ?? 0) / 100 };
}

/** Svenska kyrkan/annat trossamfund, riksgenomsnitt -- kyrkoavgift including
 *  the burial fee, which is exactly what `context.begravningsavgift` wants:
 *  one combined number, matching Adv_settings row 40 ("Begravningsavgiften
 *  samt avgiften till kyrkan/trossamfundet") and taxAndBenefits.ts's own
 *  single `kyrkskatt = cbefvi * v.begravavg.get(age)` line. */
export const CHURCH_MEMBER_RATE = latestRate("kyrkoavgift");

/** Not a member -- the burial fee alone, riksgenomsnitt. This is also what
 *  `historicalTaxRate` in setup.ts falls back to whenever `kommunalskatt`
 *  is left at 0, so it is this port's own honest default for "not a member,
 *  and no municipality picked either" -- see the footgun note in advanced.ts. */
export const BURIAL_ONLY_RATE = latestRate("begravningsavgift");

/** Stockholms stad and Tranås kommun run their own burial-fee-only regimes,
 *  separate from Svenska kyrkan's collection of everyone else's. Cited
 *  external facts (Skatteverket SKV 433 / SKVFS 2025:20, via the same
 *  reference calculator), not derived from workbook or extracted data --
 *  the same kind of addition `returnBasis`'s PPM/AP7 choices already are. */
export const STOCKHOLM_BURIAL_RATE = 0.0007;
export const TRANAS_BURIAL_RATE = 0.00285;
