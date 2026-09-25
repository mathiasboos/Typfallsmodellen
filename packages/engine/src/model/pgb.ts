/**
 * The two halves of the PGB sheet that compute themselves: conscription
 * (`wsPGB!H4`/`H5`, a single date range) and study (`wsPGB!M`, a semester
 * count). Sjuk-/aktivitetsersättning stays typed kronor -- `PgbManualYear.sa`
 * -- because the sheet's own row 4 says so: "Ange manuellt".
 *
 * Every day-count and lookup below is translated cell by cell from the
 * sheet's own formulas (`wsPGB!H6, H10:J11, G15, I15, L15, N15, P15`), read
 * off the source workbook rather than guessed -- LibreOffice's own .xlsx
 * conversion of the .xlsb preserves them as plain formula text, unlike
 * pyxlsb, which only ever returns the cached numbers.
 */
import pgbStudyJson from "../../../data/pgb-study.json" with { type: "json" };
import { vbaInt } from "../vba/math.js";

const STUDY = pgbStudyJson as {
  series: {
    studyGrantPerTerm: { firstYear: number; lastActualYear: number; values: readonly number[] };
    studyPgbFactor: { firstYear: number; lastActualYear: number; values: readonly number[] };
  };
};

/** Clamped lookup into an extracted per-year table -- held flat past both
 * ends, the same convention `employerRates` in mcalc.ts already uses for
 * K_skatt. */
function atYear(series: { firstYear: number; values: readonly number[] }, year: number): number {
  const index = Math.min(Math.max(year - series.firstYear, 0), series.values.length - 1);
  return series.values[index] ?? 0;
}

/** `wsPGB!H4`/`H5`: a single conscription period, not one entry per age. */
export interface ConscriptionPeriod {
  /** ISO `YYYY-MM-DD`. `wsPGB!H4`'s own hint: "ej före 1995-01-01". */
  readonly start: string;
  /** ISO `YYYY-MM-DD` -- "Muck", `wsPGB!H5`. */
  readonly end: string;
}

const DAY_MS = 86_400_000;
const MIN_DAYS = 120; // wsPGB!H7
const MAX_DAYS = 730; // wsPGB!H8

function parseIsoDate(iso: string): number {
  const [y, m, d] = iso.split("-").map(Number);
  return Date.UTC(y!, m! - 1, d!);
}

function daysInYear(year: number): number {
  return (Date.UTC(year, 11, 31) - Date.UTC(year, 0, 1)) / DAY_MS + 1;
}

/**
 * `wsPGB!H6` and `H10:J11`: splits a conscription period across the (up to
 * three) calendar years it touches, each clipped to that year's own
 * boundaries and to a lifetime 730-day cap; below a 120-day minimum the
 * whole period earns nothing.
 *
 * The sheet's own `H11`/`I11` additionally gate each year on being `>= 1995`
 * -- redundant here on purpose: `conscriptionPgb` below applies wsPGB!G's own
 * eligibility window (1995-2010, or 2018 on), which is a *subset* of "year >=
 * 1995", so nothing before 1995 ever survives regardless of which formula
 * drops it. `wsPGB!J11`'s own `IF(J10="-", ...)` is dead in the original for
 * the same reason `Mcalc.ts` already documents elsewhere: `J10` is a year
 * number or a literal `0`, never the string `"-"` its own sibling cells use,
 * so that branch can never take. Kept out here rather than translated
 * uselessly.
 *
 * Returns a year -> days map, at most three entries, all zero for an empty,
 * reversed or under-minimum period.
 */
export function conscriptionDaysByYear(period: ConscriptionPeriod): ReadonlyMap<number, number> {
  const start = parseIsoDate(period.start);
  const end = parseIsoDate(period.end);
  const result = new Map<number, number>();
  if (!(end > start)) return result;

  const totalDays = (end - start) / DAY_MS; // wsPGB!H6, H5-H4
  if (totalDays < MIN_DAYS) return result;

  const y1 = new Date(start).getUTCFullYear();
  const endYear = new Date(end).getUTCFullYear();
  const y2 = endYear > y1 ? y1 + 1 : undefined;
  const y3 = endYear >= y1 + 2 ? y1 + 2 : undefined;

  // wsPGB!H11: MIN(DATE(H10,12,31)-H4+1, H5-H4)
  const days1 = Math.min((Date.UTC(y1, 11, 31) - start) / DAY_MS + 1, totalDays);
  result.set(y1, days1);

  let days2 = 0;
  if (y2 !== undefined) {
    // wsPGB!I11: a full year if a third year follows, else up to the end date.
    days2 = y3 !== undefined ? daysInYear(y2) : (end - Date.UTC(y2, 0, 1)) / DAY_MS + 1;
    result.set(y2, days2);
  }

  if (y3 !== undefined) {
    // wsPGB!J11: from the start of that year to the end date, capped by
    // whatever of the 730-day lifetime maximum days1/days2 have not already used.
    const days3 = Math.min((end - Date.UTC(y3, 0, 1)) / DAY_MS + 1, MAX_DAYS - days1 - days2);
    result.set(y3, days3);
  }

  return result;
}

/** `wsPGB!F`: conscription only ever earned PGB 1995-2010, and again from 2018. */
export function conscriptionEligible(year: number): boolean {
  return (year >= 1995 && year <= 2010) || year >= 2018;
}

/**
 * `wsPGB!G` and `I`: this year's PGB from its own share of the conscription
 * period -- half of `medelPgi` (the average pension-qualifying income,
 * `v.mpgi` in mcalc.ts), pro-rated by days over 365, zero outside the
 * eligible window regardless of the day count `conscriptionDaysByYear` gives it.
 *
 * `marginal`: `ModelContext.marginal` -- 1 removes the round-down-to-100 the
 * sheet's own `INT(x/100)*100` applies, the same flag `earnPgb`'s own
 * `study` branch already reads for the identical rounding idiom.
 */
export function conscriptionPgb(year: number, days: number, medelPgi: number, marginal: number): number {
  if (!conscriptionEligible(year)) return 0;
  const raw = (days * (0.5 * medelPgi)) / 365;
  return marginal === 0 ? vbaInt(raw / 100) * 100 : raw;
}

/**
 * `wsPGB!L`, `N` and `P`: this year's PGB from a typed number of semesters --
 * the per-semester grant rate times the semesters times the year's own
 * PGB-generating share of the pension fee (0 for any year before 1995, when
 * studies first began earning PGB at all).
 */
export function studyPgb(year: number, semesters: number, marginal: number): number {
  const rate = atYear(STUDY.series.studyGrantPerTerm, year);
  const factor = atYear(STUDY.series.studyPgbFactor, year);
  const raw = rate * semesters * factor;
  return marginal === 0 ? vbaInt(raw / 100) * 100 : raw;
}
