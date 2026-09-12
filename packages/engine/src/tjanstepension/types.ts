/**
 * Shared types for the occupational pension agreements.
 *
 * Every `tl*` function in Tjänstepensioner.bas reaches for the same handful of
 * module-level globals -- the income year, the birth year, the age work started,
 * the occupational retirement age. They are gathered here and passed in, so the
 * ported functions stay pure.
 *
 * DEVIATION worth knowing: several of these functions index `year_(age)` using
 * the **global** loop age rather than their own `alder` argument, while others
 * use `year_(alder)`. Every call site in Mcalc passes the same value for both,
 * so `SchemeContext.year` stands in for both readings.
 */

/** The eight options the Start sheet offers, numbered as its drop-down is. */
export const Scheme = {
  None: 1,
  Itp1: 2,
  Itp2: 3,
  SafLo: 4,
  KapKl: 5,
  AkapKr: 6,
  /** PA16 avdelning 2 -- state employees born before 1988. */
  Pa16Avd2: 7,
  /** PA16 avdelning 1 -- state employees born 1988 or later. */
  Pa16Avd1: 8,
} as const;

export type SchemeId = (typeof Scheme)[keyof typeof Scheme];

/** Labels as the workbook words them, for the UI and for error messages. */
export const SCHEME_LABELS: Readonly<Record<SchemeId, string>> = {
  [Scheme.None]: "Saknar tjänstepension",
  [Scheme.Itp1]: "ITP-1, Privatanställda tjänstemän, födda 1979-",
  [Scheme.Itp2]: "ITP-2, Privatanställda tjänstemän, födda före 1979",
  [Scheme.SafLo]: "SAF-LO + STP, Privatanställda arbetare",
  [Scheme.KapKl]: "KAP-KL, Kommunal- & Regionalanställda",
  [Scheme.AkapKr]: "AKAP-KR, Kommunal- & Regionalanställda, födda 1986-",
  [Scheme.Pa16Avd2]: "PA16 (Avd 2), Statligt anställda, födda före 1988",
  [Scheme.Pa16Avd1]: "PA16 (Avd 1), Statligt anställda, födda 1988 och senare",
};

/** The run-level facts the agreement functions read from globals in the VBA. */
export interface SchemeContext {
  /** The income year being computed. */
  readonly year: number;
  /** Birth year, possibly with a fraction for the birth month. */
  readonly born: number;
  /** Age at which work started. */
  readonly wStart: number;
  /** Occupational pension retirement age (`tjp_par`). */
  readonly tjpPar: number;
  /** `rng_FlexPens`: extra flexpension premium, ITP 1 and SAF-LO, from 2014. */
  readonly flexPension: number;
  /** `marginal`: 0 applies the rules' rounding, 1 removes it. */
  readonly marginal: 0 | 1;
}

/** The contribution ceiling every agreement steps at: 7.5 income base amounts. */
export const BREAKPOINT_IBB = 7.5;
