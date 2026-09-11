/**
 * Projects the economic series past their last actual year.
 *
 * `packages/data/economic-series.json` holds actuals only. The workbook's own
 * projection columns are live Excel formulas driven by the Start sheet's
 * inflation and growth inputs, so snapshotting them would have frozen the
 * shipped file's 0% / 0% assumptions into the data and the site would return
 * the same numbers whatever the user typed.
 *
 * Each rule below mirrors one sheet formula. They are written up, with the
 * originals, in docs/PROJECTION-RULES.md. Order matters: prisbasbelopp reads
 * the previous year's KPI, inkomstbasbelopp reads the *next* year's income
 * index, and balansindex chains off its own previous value.
 *
 * Verified against reference/fixtures/economic-series-cached.json, which holds
 * every value the workbook had cached under the shipped assumptions.
 */

import { excelRound, vbaInt } from "../vba/math.js";
import type { EconomicAssumptions, RawEconomicSeries, SeriesName } from "./types.js";

/** A full series of yearly values, indexed by year. */
export class Series {
  constructor(
    readonly name: string,
    readonly firstYear: number,
    private readonly values: Float64Array,
  ) {}

  get lastYear(): number {
    return this.firstYear + this.values.length - 1;
  }

  /** Value for `year`, clamped to the ends of the range. */
  at(year: number): number {
    const index = Math.min(Math.max(year - this.firstYear, 0), this.values.length - 1);
    return this.values[index]!;
  }

  /** Value for `year`, or undefined outside the range. */
  tryAt(year: number): number | undefined {
    const index = year - this.firstYear;
    return index < 0 || index >= this.values.length ? undefined : this.values[index];
  }
}

export type EconomicData = Readonly<Record<SeriesName, Series>>;

/**
 * The income pension's balance index (`Balansindex` in Pensionssystemet.bas).
 *
 * Caps at the income index: the balance index catches up to it but never
 * exceeds it.
 *
 * @param bi    the previous year's balance index
 * @param btal  the damped balance ratio, (balanstal - 1) / 3 + 1
 * @param ital1 income index for this year
 * @param ital2 income index for last year
 */
export function balansindex(
  bi: number,
  btal: number,
  ital1: number,
  ital2: number,
  marginal = 0,
): number {
  let result = 0;
  if (bi === 0) {
    if (btal < 1) result = btal * ital1;
  } else {
    result = (btal * bi * ital1) / ital2;
  }
  if (result > ital1) result = ital1;
  // Int(x * 100 + 0.49) / 100 -- the VBA's own two-decimal rounding idiom.
  if (marginal === 0) result = vbaInt(result * 100 + 0.49) / 100;
  return result;
}

/**
 * The stand-in for "no second state tax threshold" -- what the sheet writes
 * (`=10^16`) for every year since värnskatten was abolished in 2020.
 */
export const NO_SECOND_THRESHOLD = 1e16;

/** First year the extracted actuals carry a value for a series. */
function firstYearWithValue(raw: RawEconomicSeries, name: SeriesName): number {
  const series = raw.series[name];
  if (series === undefined) throw new Error(`economic-series.json has no series "${name}"`);
  const offset = series.values.findIndex((value) => value !== null && value !== 0);
  return offset === -1 ? Number.POSITIVE_INFINITY : series.firstYear + offset;
}

interface Working {
  readonly firstYear: number;
  readonly lastYear: number;
  readonly raw: RawEconomicSeries;
  readonly assumptions: EconomicAssumptions;
  readonly out: Map<SeriesName, Float64Array>;
}

function actualsOf(w: Working, name: SeriesName): { values: Float64Array; lastActual: number } {
  const raw = w.raw.series[name];
  if (raw === undefined) throw new Error(`economic-series.json has no series "${name}"`);
  const values = new Float64Array(w.lastYear - w.firstYear + 1);
  for (let i = 0; i < raw.values.length; i += 1) {
    const year = raw.firstYear + i;
    const index = year - w.firstYear;
    if (index >= 0 && index < values.length) values[index] = raw.values[i] ?? 0;
  }
  return { values, lastActual: raw.lastActualYear };
}

function get(w: Working, name: SeriesName, year: number): number {
  const values = w.out.get(name);
  if (values === undefined) throw new Error(`series "${name}" projected out of order`);
  const index = Math.min(Math.max(year - w.firstYear, 0), values.length - 1);
  return values[index]!;
}

/**
 * Fills a series: actuals as given, then `project` for each later year.
 *
 * `project` returns the value for `year`; it may read any series already
 * completed, and earlier years of its own via `self`.
 */
function fill(
  w: Working,
  name: SeriesName,
  project: (year: number, self: (year: number) => number) => number,
  options: { derived?: boolean } = {},
): void {
  const { values, lastActual } = actualsOf(w, name);
  w.out.set(name, values);
  const self = (year: number) => get(w, name, year);
  const from = options.derived === true ? w.firstYear : Math.max(w.firstYear, lastActual + 1);
  for (let year = from; year <= w.lastYear; year += 1) {
    values[year - w.firstYear] = project(year, self);
  }
}

/**
 * Builds the full set of series for `firstYear..lastYear` from the extracted
 * actuals and the given assumptions.
 */
export function projectEconomicData(
  raw: RawEconomicSeries,
  assumptions: EconomicAssumptions,
  options: { firstYear?: number; lastYear?: number } = {},
): EconomicData {
  const firstYear = options.firstYear ?? 1957;
  const requestedLastYear = options.lastYear ?? 2153;
  // inkomstbasbelopp reads the *next* year's income index, so compute one year
  // beyond what was asked for and let that spare year absorb the lookup. The
  // workbook cannot do this and leaves a 0 in its own final row.
  const lastYear = requestedLastYear + 1;
  const w: Working = { firstYear, lastYear, raw, assumptions, out: new Map() };

  const { pbbBase, fpbBase, ibbBase, kpiJuneDivisor, incomeIndexDivisor, kpiRoundDecimals } =
    raw.anchors;
  const { yearlyInflation, realGrowth, realReturn, marginal, returnsNetOfFees, rgk } = assumptions;
  const inflationFactor = 1 + yearlyInflation;
  const indexFactor = (1 + realGrowth) * (1 + yearlyInflation);
  /** Applies the rules' rounding unless `marginal` switches it off. */
  const round = (value: number, digits: number) =>
    marginal === 1 ? value : excelRound(value, digits);

  // Prices and the income index drive everything else, so they go first.
  fill(w, "kpiJune", (year, self) =>
    excelRound(self(year - 1) * inflationFactor, kpiRoundDecimals),
  );
  fill(w, "kpiAnnual", (year, self) =>
    excelRound(self(year - 1) * inflationFactor, kpiRoundDecimals),
  );
  fill(w, "inkomstindex", (year, self) => self(year - 1) * indexFactor);

  // Basbelopp follow from them. Note each reads the *previous* year's KPI, and
  // that inkomstbasbelopp reads the *next* year's income index -- both quirks of
  // the sheet, both deliberate.
  fill(w, "prisbasbelopp", (year) =>
    round((pbbBase * get(w, "kpiJune", year - 1)) / kpiJuneDivisor, -2),
  );
  fill(w, "forhojtPrisbasbelopp", (year) =>
    round((get(w, "kpiJune", year - 1) / kpiJuneDivisor) * fpbBase, -2),
  );
  fill(w, "inkomstbasbelopp", (year) =>
    round((get(w, "inkomstindex", year + 1) / incomeIndexDivisor) * ibbBase, -2),
  );
  fill(w, "medelPgi", (year, self) =>
    (self(year - 1) * get(w, "inkomstindex", year)) / get(w, "inkomstindex", year - 1),
  );

  // Balancing. Balanstal projects to its neutral 1.0 rather than carrying its
  // last outcome forward, which is why a "repeat the last value" heuristic
  // cannot be used for these.
  fill(w, "balanstal", () => 1);
  fill(w, "balansindex", (year, self) =>
    balansindex(
      self(year - 1),
      (get(w, "balanstal", year) - 1) / 3 + 1,
      get(w, "inkomstindex", year),
      get(w, "inkomstindex", year - 1),
      marginal,
    ),
  );
  // Before balancing began the sheet simply takes the income index (`=I7`);
  // from the first balancing year it chooses between the two. The boundary is
  // read from the data -- the first year balansindex has a value -- rather than
  // hard-coded, so a revised history is picked up automatically.
  const balancingFrom = firstYearWithValue(raw, "balansindex");
  fill(
    w,
    "gallandeIndex",
    (year) =>
      year < balancingFrom || assumptions.latestIndexBasis === 2
        ? get(w, "inkomstindex", year)
        : get(w, "balansindex", year),
    { derived: true },
  );

  // Fees, and the returns that are quoted net of them.
  fill(w, "adminavgiftPp", () => (returnsNetOfFees ? 0 : 0.0002662 / 100));
  fill(w, "forvaltningsavgiftPp", () => (returnsNetOfFees ? 0 : 0.14 / 100));
  fill(
    w,
    "totalAvgiftPp",
    (year) =>
      returnsNetOfFees
        ? get(w, "adminavgiftPp", year)
        : get(w, "adminavgiftPp", year) + get(w, "forvaltningsavgiftPp", year),
    { derived: true },
  );
  fill(w, "kvarEfterAvgiftPp", (year) => 1 - get(w, "totalAvgiftPp", year), { derived: true });
  fill(w, "kvarEfterAdminIp", (year, self) => self(year - 1));

  const projectedReturn = (year: number) => {
    const fees = returnsNetOfFees
      ? 0
      : get(w, "adminavgiftPp", year) + get(w, "forvaltningsavgiftPp", year);
    return (1 + realReturn - fees) * inflationFactor - 1;
  };
  fill(w, "avkastningPpm", projectedReturn);
  fill(w, "avkastningAp7", projectedReturn);
  fill(w, "rantaRiksgalden", () => rgk * 100);

  // State income tax thresholds. The +51 before truncating to hundreds adds
  // 5 100 kr; it reproduces the sheet, though not the real decided figures.
  // See docs/PROJECTION-RULES.md -- parity with the model is the goal.
  fill(w, "skiktgrans1", (year, self) => {
    const kpiRatio = get(w, "kpiJune", year - 1) / get(w, "kpiJune", year - 2) + 0.02;
    return marginal === 1
      ? self(year - 1) * kpiRatio
      : vbaInt((self(year - 1) * kpiRatio) / 100 + 51) * 100;
  });
  // Värnskatten was abolished in 2020, and the sheet writes a literal 10^16
  // from then on rather than carrying the last real threshold forward.
  fill(w, "skiktgrans2", () => NO_SECOND_THRESHOLD);

  const data: Partial<Record<SeriesName, Series>> = {};
  for (const [name, values] of w.out) data[name] = new Series(name, firstYear, values);
  return data as EconomicData;
}
