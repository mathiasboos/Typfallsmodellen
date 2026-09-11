/**
 * VBA and Excel arithmetic semantics that differ from JavaScript's.
 *
 * These are the quietest source of wrong answers in a port like this. The model
 * rounds deliberately at dozens of points -- `marginal = 0` switches real-world
 * rounding on across the whole rule set -- so a rounding function that is
 * *nearly* right produces figures that are nearly right, which is worse than
 * failing loudly.
 *
 * Use these instead of `Math.round`, `Math.trunc` and friends anywhere the
 * original VBA rounds. See docs/VBA-MAPPING.md.
 */

/**
 * VBA's `Int()`: rounds toward negative infinity.
 *
 * The single most important function here -- the VBA calls `Int` 1 061 times,
 * usually as the rounding idiom `Int(x / 12 + 0.5) * 12`. JavaScript's
 * `Math.trunc` rounds toward zero instead and disagrees for every negative
 * non-integer, so it is never a substitute.
 *
 * Excel's worksheet `INT()` behaves the same way.
 */
export function vbaInt(value: number): number {
  return Math.floor(value);
}

/** VBA's `Fix()`: truncates toward zero. Differs from {@link vbaInt} for negatives. */
export function vbaFix(value: number): number {
  return Math.trunc(value);
}

/**
 * VBA's intrinsic `Round()`: banker's rounding, half to even.
 *
 * `Round(2.5)` is 2 and `Round(3.5)` is 4. Note this is *not*
 * `Application.WorksheetFunction.Round`, which rounds half away from zero --
 * see {@link excelRound}. Every one of the ~90 `Round(` calls in the VBA is the
 * intrinsic, so this is the one to reach for when porting them.
 *
 * Scaling by a power of ten and then testing the half exactly mirrors what the
 * VBA runtime does, rather than trying to recover the decimal value the user
 * typed. The two differ: 2.675 is held as 2.67499999999999982, but multiplying
 * by 100 rounds it up to exactly 267.5, so this returns 2.68 where a
 * decimal-aware implementation would return 2.67.
 *
 * UNVERIFIED at digits > 0. The model calls `Round(x, 2)` about ten times, on
 * continuous monetary and actuarial quantities where landing exactly on a half
 * is rare, so this is unlikely to bite -- but it has not been checked against a
 * running VBA. The Phase 2 golden-file comparison settles it; if a case ever
 * disagrees by a hundredth, look here first.
 */
export function vbaRound(value: number, digits = 0): number {
  if (!Number.isFinite(value)) return value;
  const factor = 10 ** digits;
  const scaled = value * factor;
  const floor = Math.floor(scaled);
  const result =
    scaled - floor === 0.5
      ? floor % 2 === 0
        ? floor
        : floor + 1
      : Math.round(scaled);
  return result / factor;
}

/**
 * Excel's worksheet `ROUND()`: half away from zero.
 *
 * Used by the sheet formulas the projection rules mirror, e.g.
 * `ROUND(pbbBase * kpiJune / 257.38, -2)`.
 *
 * Excel rounds the decimal value a user would see rather than the raw binary
 * double, so the scaled value is first collapsed to 15 significant digits --
 * Excel's own working precision. Without that, `ROUND(2.675, 2)` would give
 * 2.67 here and 2.68 in Excel.
 */
export function excelRound(value: number, digits = 0): number {
  if (!Number.isFinite(value)) return value;
  const sign = value < 0 ? -1 : 1;
  const factor = 10 ** digits;
  const scaled = Number((Math.abs(value) * factor).toPrecision(15));
  return (sign * Math.floor(scaled + 0.5)) / factor;
}

/**
 * Coercion to VBA's `Integer`/`Long`: banker's rounding, not truncation.
 *
 * Assigning a Double to an integer variable rounds in VBA. Apply this where the
 * original declares `As Integer` or `As Long` and assigns a non-integral value.
 */
export function vbaCLng(value: number): number {
  return vbaRound(value, 0);
}

/**
 * VBA's integer division operator `\`: operands rounded to Long, then the
 * quotient truncated toward zero.
 */
export function vbaIntDiv(left: number, right: number): number {
  return Math.trunc(vbaCLng(left) / vbaCLng(right));
}

/**
 * Coercion to VBA's `Single` (32-bit float).
 *
 * A handful of the model's variables are declared `As Single` -- `Iyear`,
 * `pgi_years`, `forstid` among them. They hold years and counts where the
 * narrower type never bites, but boundary comparisons are exactly where it
 * would, so the helper exists for when a divergence needs ruling out.
 */
export function vbaSingle(value: number): number {
  return Math.fround(value);
}

/** `Application.WorksheetFunction.Max`, over a list or array. */
export function wsMax(...values: number[]): number {
  return values.reduce((a, b) => (b > a ? b : a), Number.NEGATIVE_INFINITY);
}

/** `Application.WorksheetFunction.Min`, over a list or array. */
export function wsMin(...values: number[]): number {
  return values.reduce((a, b) => (b < a ? b : a), Number.POSITIVE_INFINITY);
}

/**
 * `Application.WorksheetFunction.Large(array, k)`: the k-th largest value,
 * 1-indexed.
 *
 * The model uses it to pick the best earning years -- the ATP fifteen-year rule
 * and the final-salary average.
 */
export function wsLarge(values: readonly number[], k: number): number {
  if (k < 1 || k > values.length) {
    throw new RangeError(`Large: k=${k} outside 1..${values.length}`);
  }
  const sorted = [...values].sort((a, b) => b - a);
  return sorted[k - 1]!;
}

/** `Application.WorksheetFunction.Average`. */
export function wsAverage(values: readonly number[]): number {
  if (values.length === 0) throw new RangeError("Average: no values");
  return values.reduce((a, b) => a + b, 0) / values.length;
}
