/**
 * What the model leaves of an income, for any income.
 *
 * `run()` taxes the one income the typfall has and reports it. A budget is an
 * amount *after* tax, so comparing it with a pension also asks the opposite
 * question -- what would it be before tax? -- and the honest answer is the
 * model's own tax section run on another income, not an average rate: the tax
 * is progressive, and its rules change with the person's age (the higher
 * grundavdrag for the old begins at a certain age, not at retirement).
 *
 * The tax section reads the age's gross income from the run state and writes
 * the net back, so after the loop has finished it can simply be run again with
 * the gross swapped. The probe puts everything it touches back, so the run is
 * left exactly as `run()` leaves it.
 *
 * Only meaningful for ages with no earned income -- from the retirement age on --
 * where gross income is pension and nothing else; before it, the age's own
 * salary is taxed along with whatever is probed.
 */

import type { ModelContext } from "./context.js";
import type { TypfallInput } from "./input.js";
import { runLoop, ruleYears } from "./mcalc.js";
import type { Run } from "./mcalc.js";
import { buildResult } from "./result.js";
import type { TypfallResult } from "./result.js";
import { prepare } from "./run.js";
import type { RunOptions } from "./run.js";
import { priceFactor, taxes } from "./taxAndBenefits.js";

export interface TaxProbe {
  /** The same result `run()` returns. */
  readonly result: TypfallResult;
  /**
   * The finished run the probe works on. Reading it is what the tests do to show that
   * probing leaves it untouched; nothing else needs it.
   */
  readonly run: Run;
  /**
   * What is left of `gross` kronor a year at `age` after the model's tax, both in
   * the run's price basis (the amounts `result.rows` holds).
   */
  netAt(age: number, gross: number): number;
  /**
   * The gross income a year at `age` that leaves `net` after tax -- the inverse of
   * `netAt`, found by bisection (the net never falls as the gross rises).
   */
  grossFor(age: number, net: number): number;
}

/** Gross income searched up to when inverting, in nominal kronor a year. */
const SEARCH_CEILING = 1e10;

/** `run()`, with the tax section left open for `netAt` and `grossFor`. */
export function runWithTaxProbe(input: TypfallInput, context: ModelContext, options: RunOptions): TaxProbe {
  const prepared = prepare(input, context, options);
  const { run } = prepared;
  runLoop(run);
  const result = buildResult(run, prepared.warnings);
  const { s, v } = run;

  const netAt = (age: number, gross: number): number => {
    if (!(gross > 0)) return 0;
    if (age < v.startage || age > v.slutage) throw new RangeError(`age ${age} is outside the run (${v.startage}-${v.slutage})`);
    const factor = priceFactor(run, age);
    const { utgyear, skyear } = ruleYears(run, age);

    const kept = {
      brutto: s.brutto.get(age),
      netto: s.netto.get(age),
      municipalTax: s.municipalTax.get(age),
      stateTax: s.stateTax.get(age),
      leftovers: { ...s.leftovers },
    };
    s.brutto.set(age, gross / factor);
    taxes(run, age, utgyear, skyear);
    const net = s.netto.get(age) * factor;

    s.brutto.set(age, kept.brutto);
    s.netto.set(age, kept.netto);
    s.municipalTax.set(age, kept.municipalTax);
    s.stateTax.set(age, kept.stateTax);
    Object.assign(s.leftovers, kept.leftovers);
    return net;
  };

  const grossFor = (age: number, net: number): number => {
    if (!(net > 0)) return 0;
    // Tax is never negative, so the gross is at least the net.
    let low = net;
    let high = net * 2;
    while (netAt(age, high) < net) {
      low = high;
      high *= 2;
      if (high > SEARCH_CEILING) return Number.POSITIVE_INFINITY;
    }
    for (let step = 0; step < 80 && high - low > 0.005; step += 1) {
      const middle = (low + high) / 2;
      if (netAt(age, middle) >= net) high = middle;
      else low = middle;
    }
    return high;
  };

  return { result, run, netAt, grossFor };
}
