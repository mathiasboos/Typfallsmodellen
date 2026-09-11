/** Shapes of the generated data in packages/data, and of the assumptions that drive projection. */

/** One yearly series: actual values, and the year they stop at. */
export interface RawSeries {
  readonly firstYear: number;
  /** Last year whose value is an actual. Past this the engine projects. */
  readonly lastActualYear: number;
  readonly source: string;
  readonly note: string;
  readonly values: readonly (number | null)[];
}

export interface RawEconomicSeries {
  readonly source: string;
  readonly note: string;
  readonly lastSheetYear: number;
  readonly anchors: {
    readonly pbbBase: number;
    readonly fpbBase: number;
    readonly ibbBase: number;
    readonly kpiJuneDivisor: number;
    readonly incomeIndexDivisor: number;
    readonly kpiRoundDecimals: number;
  };
  readonly series: Readonly<Record<string, RawSeries>>;
}

/**
 * The economic assumptions that drive projection.
 *
 * The defaults are the forecasting standard (prognosstandard) the model ships
 * with: fixed prices and no real growth, so results come out in today's price
 * and wage level and compare directly against the entered salary.
 */
export interface EconomicAssumptions {
  /** rng_Yearly_Inflation. Prognosstandard: 0. */
  readonly yearlyInflation: number;
  /** rng_Real_Growth, i.e. general income growth. Prognosstandard: 0. */
  readonly realGrowth: number;
  /**
   * rng_FondAvkastning: return in excess of general wage growth, so that the
   * forecast stays comparable with current income. Prognosstandard: 0.017.
   */
  readonly realReturn: number;
  /** `marginal`: 0 applies the real rules' rounding, 1 removes all of it. */
  readonly marginal: 0 | 1;
  /** rng_Avkastning_fondavgifter: true when the return is already net of fund fees. */
  readonly returnsNetOfFees: boolean;
  /** rng_Senaste_Index_Framskrivning: 2 means use inkomstindex for every year, ignoring balansindex. */
  readonly latestIndexBasis: number;
  /** Rgk: the Riksgälden rate used while contributions await assessment. */
  readonly rgk: number;
}

export const PROGNOSSTANDARD: EconomicAssumptions = {
  yearlyInflation: 0,
  realGrowth: 0,
  realReturn: 0.017,
  marginal: 0,
  returnsNetOfFees: true,
  latestIndexBasis: 1,
  rgk: 0.01,
};

/** Names of the series the engine projects and reads. */
export type SeriesName =
  | "kpiJune"
  | "kpiAnnual"
  | "prisbasbelopp"
  | "medelPgi"
  | "inkomstbasbelopp"
  | "forhojtPrisbasbelopp"
  | "inkomstindex"
  | "balanstal"
  | "balansindex"
  | "gallandeIndex"
  | "totalAvgiftPp"
  | "avkastningPpm"
  | "avkastningAp7"
  | "rantaRiksgalden"
  | "adminavgiftPp"
  | "forvaltningsavgiftPp"
  | "skiktgrans1"
  | "skiktgrans2"
  | "kvarEfterAdminIp"
  | "kvarEfterAvgiftPp";
