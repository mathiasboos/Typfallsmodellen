/**
 * Private saving in a kapitalförsäkring (KF) or investeringssparkonto (ISK).
 *
 * Port of `PrivatSparande.bas`. Neither wrapper is taxed on its actual return:
 * both pay a flat yield tax on a notional return (schablonintäkt) derived from
 * the government borrowing rate.
 *
 * ---
 *
 * A WARNING ABOUT STATE.
 *
 * `PrivatSparande.bas` does not declare `Option Explicit`, so its undeclared
 * variables -- `avkskatten`, `Slr`, `AvkSkaUl`, `underlag` -- are implicit
 * module-level Variants that **persist between calls**. That matters in two
 * places where the original's branches do not cover every input:
 *
 *   - `AvkastningsskattKFISK` tests `> 150000` and `< 150000` (and the same at
 *     300 000). At exactly the threshold neither fires, and the function
 *     returns whatever the previous call left behind.
 *   - `Schablonintakt` has no `Case` for a year before 1986, and the same
 *     applies.
 *
 * Both are reproduced here with module-level variables, because parity with the
 * model is the goal and a golden-file case could land on a round 150 000. The
 * consequence is that these functions are **not pure**: the same arguments can
 * give different answers depending on what was computed before. Call
 * {@link resetPrivateSavingState} between independent runs. The VBA has no such
 * reset, so a model run that follows another in the same Excel session inherits
 * the earlier one's leftovers — which the port cannot reproduce and does not try to.
 */

/** Which wrapper the saving sits in. */
export const SavingType = {
  /** Kapitalförsäkring. */
  KF: 1,
  /** Investeringssparkonto, which did not exist before 2012. */
  ISK: 2,
} as const;

export type SavingTypeId = (typeof SavingType)[keyof typeof SavingType];

/** Mirrors the VBA's implicit module-level `avkskatten`. */
let lastYieldTax = 0;
/** Mirrors the VBA's implicit module-level `Slr`. */
let lastBorrowingRate = 0;

/** Clears the carried-over values. No equivalent exists in the VBA. */
export function resetPrivateSavingState(): void {
  lastYieldTax = 0;
  lastBorrowingRate = 0;
}

/**
 * The government borrowing rate behind the notional return.
 *
 * An annual average up to 2011, and the rate on 30 November of the preceding
 * year from 2012. Since 2016 a margin is added, with a floor.
 */
export function Schablonintakt(ar: number): number {
  const previousYear = ar - 1;

  const rates: Readonly<Record<number, number>> = {
    1986: 0.1077, 1987: 0.1167, 1988: 0.1135, 1989: 0.1118, 1990: 0.1313,
    1991: 0.1072, 1992: 0.1003, 1993: 0.0855, 1994: 0.0957, 1995: 0.1014,
    1996: 0.0789, 1997: 0.0647, 1998: 0.0498, 1999: 0.0489, 2000: 0.0534,
    2001: 0.0498, 2002: 0.0515, 2003: 0.0439, 2004: 0.043, 2005: 0.0324,
    2006: 0.0362, 2007: 0.0414, 2008: 0.0387, 2009: 0.0311, 2010: 0.0277,
    // From here the figure is the rate on 30 November rather than an average.
    2011: 0.0165, 2012: 0.0149, 2013: 0.0209, 2014: 0.009, 2015: 0.0065,
    2016: 0.0027, 2017: 0.0049, 2018: 0.0051, 2019: -0.0009, 2020: -0.001,
    2021: 0.0023, 2022: 0.0194, 2023: 0.0262, 2024: 0.0196, 2025: 0.0255,
  };

  // A year with no entry leaves the previous call's value in place -- see the
  // note at the top of this file.
  let slr = rates[previousYear] ?? lastBorrowingRate;

  if (ar > 2025) slr = 0.025;

  if (ar > 2015 && ar < 2018) {
    slr += 0.0075;
    if (slr < 0.0125) slr = 0.0125;
  } else if (ar > 2017) {
    slr += 0.01;
    if (slr < 0.0125) slr = 0.0125;
  }

  lastBorrowingRate = slr;
  return slr;
}

/**
 * Yield tax on a KF or ISK.
 *
 * 27% of the notional return before 2012, 30% after. A tax-free allowance was
 * introduced in 2025 at 150 000 kr and raised to 300 000 kr in 2026.
 */
export function AvkastningsskattKFISK(ingaendeBalans: number, ar: number): number {
  let tax: number;

  if (ar < 2012) {
    tax = ingaendeBalans * 0.27 * Schablonintakt(ar);
  } else if (ar < 2025) {
    tax = ingaendeBalans * 0.3 * Schablonintakt(ar);
  } else {
    const allowance = ar === 2025 ? 150_000 : 300_000;
    if (ingaendeBalans > allowance) {
      tax = (ingaendeBalans - allowance) * 0.3 * Schablonintakt(ar);
    } else if (ingaendeBalans < allowance) {
      tax = 0;
    } else {
      // Exactly at the allowance neither branch fires in the VBA, so the
      // previous call's value stands. See the note at the top of this file.
      tax = lastYieldTax;
    }
  }

  lastYieldTax = tax;
  return tax;
}

/**
 * Rolls a private saving balance forward one year.
 *
 * @param ingaendeBalans opening balance
 * @param sparandePerAr  saved during the year
 * @param yield_         the year's growth factor
 */
export function PrivatSpar(
  ingaendeBalans: number,
  sparandePerAr: number,
  yield_: number,
  ar: number,
  typ: SavingTypeId,
): number {
  let balance = ingaendeBalans;

  if (typ === SavingType.KF) {
    // KF is taxed on the opening value plus deposits, those in the second half
    // of the year counting half.
    const taxBase = balance + sparandePerAr * (6 / 12) + sparandePerAr * (6 / 12) * 0.5;
    balance =
      balance * yield_ + sparandePerAr * yield_ ** (180 / 360) - AvkastningsskattKFISK(taxBase, ar);
  } else if (typ === SavingType.ISK) {
    // ISK did not exist before 2012; before then the balance is left untouched.
    if (ar > 2011) {
      // Taxed on the average of the value at the start of each quarter, plus
      // the year's deposits divided by four.
      const taxBase =
        (balance +
          balance * yield_ ** (90 / 360) +
          sparandePerAr * (3 / 12) * yield_ ** (45 / 360) +
          balance * yield_ ** (180 / 360) +
          sparandePerAr * (6 / 12) * yield_ ** (90 / 360) +
          balance * yield_ ** (270 / 360) +
          sparandePerAr * (9 / 12) * yield_ ** (135 / 360) +
          sparandePerAr) /
        4;
      balance =
        balance * yield_ +
        sparandePerAr * yield_ ** (180 / 360) -
        AvkastningsskattKFISK(taxBase, ar);
    }
  }

  return balance;
}
