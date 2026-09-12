/**
 * Garantipension and inkomstpensionstillägg -- the two flat top-ups to a small
 * public pension.
 *
 * Port of `gp` and `tillagg` from Bidrag.bas.
 *
 * `gp` carries two entirely separate rule sets side by side: cohorts born 1938
 * and later get the guarantee pension proper (SFS 1998:702), a flat amount
 * reduced against the income pension; cohorts born 1937 and earlier stay on the
 * folkpension/pensionstillskott construction that preceded it, which builds the
 * amount up from income rather than reducing it down.
 *
 * `tillagg` is the income pension supplement introduced in 2021 -- a trapezoid
 * in pension income, deflated to 2021 money before the brackets are applied.
 */

import { vbaInt, vbaSingle } from "../vba/math.js";

/**
 * Garantipension for a year.
 *
 * @param inkomst  Beräkningsunderlag at a full withdrawal.
 * @param civ      0 single, 1 married or cohabiting.
 * @param fodar    Year of birth.
 * @param pbb      The year's prisbasbelopp.
 * @param ftid     Försäkringstid, years of residence.
 * @param marginal 0 applies the rounding, 1 removes it.
 * @param alder    Age.
 * @param year     Income year.
 * @param wyear    The year from which the floor is indexed to earnings.
 * @param IBB      That year's inkomstbasbelopp.
 * @param kvoten   The pbb/IBB ratio at `wyear`.
 * @param rikt     Riktålder -- nothing is paid before it.
 * @param andel    Withdrawal share.
 */
export function gp(
  inkomst: number,
  civ: number,
  fodar: number,
  pbb: number,
  ftid = 40,
  marginal = 0,
  alder = 65,
  year = 2100,
  wyear = 21000,
  IBB = 0,
  kvoten = 1,
  rikt = 65,
  andel = 1,
): number {
  let result = 0;

  // Garantipension starts at 65, or at the riktålder from 2020.
  if (alder < rikt) return result;

  // NOTE: `>` here, where BTP and SBTP write `>=` for the same switch to
  // earnings indexation. Kept as written.
  if (year > wyear && kvoten < 1) {
    // Swap pbb for IBB and scale the thresholds back to the old level: the
    // pensions themselves follow the indexation, so in time everyone reaches
    // the guarantee.
    pbb = IBB;
    pbb = pbb * kvoten;
  }

  if (fodar > 1937) {
    if (civ === 0) {
      if (inkomst <= 1.26 * pbb) {
        result = 2.13 * pbb - inkomst;
        if (year > 2019) result = 2.181 * pbb - inkomst;
        if (year === 2022) result = ((2.181 * 7 + 2.43 * 5) * pbb) / 12 - inkomst;
        if (year > 2022) result = 2.43 * pbb - inkomst;
      } else {
        result = 0.87 * pbb - 0.48 * (inkomst - 1.26 * pbb);
        if (year > 2019) result = 0.921 * pbb - 0.48 * (inkomst - 1.26 * pbb);
        if (year === 2022) {
          result = ((7 * 0.921 + 5 * 1.17) * pbb) / 12 - 0.48 * (inkomst - 1.26 * pbb);
        }
        if (year > 2022) result = 1.17 * pbb - 0.48 * (inkomst - 1.26 * pbb);
      }
    } else {
      if (inkomst <= 1.14 * pbb) {
        result = 1.9 * pbb - inkomst;
        if (year > 2019) result = 1.951 * pbb - inkomst;
        if (year === 2022) result = ((7 * 1.951 + 5 * 2.2) * pbb) / 12 - inkomst;
        if (year > 2022) result = 2.2 * pbb - inkomst;
      } else {
        result = 0.76 * pbb - 0.48 * (inkomst - 1.14 * pbb);
        if (year > 2019) result = 0.811 * pbb - 0.48 * (inkomst - 1.14 * pbb);
        if (year === 2022) {
          result = ((7 * 0.811 + 5 * 1.06) * pbb) / 12 - 0.48 * (inkomst - 1.14 * pbb);
        }
        if (year > 2022) result = 1.06 * pbb - 0.48 * (inkomst - 1.14 * pbb);
      }
    }

    // NOTE: the comment in the VBA says three years of residence are enough,
    // but the test is `ftid < 4`, so four are required. Kept as written.
    if (result < 0 || ftid < 4) result = 0;
  } else {
    if (civ === 0) {
      // NOTE: the single branch adds the 2020 supplement to *income* before the
      // brackets are applied; the cohabiting branch below does not. Both add it
      // to the benefit afterwards. Asymmetric in the original, kept.
      if (year > 2019 && year < 2022) inkomst = inkomst + 0.051 * pbb;
      if (year === 2022) inkomst = inkomst + (0.051 * pbb * 7) / 12 + (0.3 * pbb * 5) / 12;
      if (year > 2022) inkomst = inkomst + 0.3 * pbb;

      if (inkomst <= 0.25 * pbb) {
        result = inkomst * 1.04;
      } else if (inkomst < 1.354 * pbb) {
        result = 1.5174 * inkomst - 0.1193 * pbb;
      } else if (inkomst <= 1.529 * pbb) {
        result = 1.343 * inkomst + 0.1168 * pbb;
      } else if (inkomst <= 3.16 * pbb) {
        result = 2.17 * pbb + 0.6 * (inkomst - 1.51 * pbb);
      } else {
        result = 0;
      }
      if (year > 2019 && inkomst < 3.16 * pbb) result = result + 0.051 * pbb;
    } else {
      if (inkomst <= 0.25 * pbb) {
        result = inkomst * 1.04;
      } else if (inkomst < 1.354 * pbb) {
        result = 1.5174 * inkomst - 0.1193 * pbb;
      } else if (inkomst <= 2.8275 * pbb) {
        result = 1.935 * pbb + 0.6 * (inkomst - 1.34 * pbb);
      } else {
        result = 0;
      }

      if (year > 2019 && inkomst < 2.8275 * pbb) result = result + 0.051 * pbb;
    }

    result = result - inkomst;
    if (result < 0) result = 0;
  }

  // Försäkringstid, out of 40 years. NOTE: always 40 here, where `tillagg`
  // varies the denominator by cohort.
  if (ftid < 40 && result > 0) result = (result * ftid) / 40;
  // A withdrawal share below a full one.
  if (andel < 1) result = result * andel;
  if (marginal === 0) result = 12 * vbaInt(result / 12 + 0.5);

  return result;
}

/**
 * Inkomstpensionstillägg, the supplement introduced in September 2021.
 *
 * The brackets are written in 2021 money, so the pension is deflated by the
 * change in inkomstindex net of the 1.6% follow-up norm before they are applied.
 *
 * NOTE: declared `As Single` in the VBA, so every assignment to the result
 * narrows to a 32-bit float. `vbaSingle` reproduces that.
 *
 * NOTE: `born` and `marginal` are not parameters in the VBA -- they are the
 * `Public` globals declared in VBA_go.bas, which `tillagg` reads directly.
 * Bidrag.bas does declare `Option Explicit`, so these carry real values rather
 * than being empty variants. They are explicit here, placed after the required
 * arguments the way `PublicAvg` already handles the same situation.
 *
 * @param underl    Annual pension underlying the supplement.
 * @param year      Income year.
 * @param born      Year of birth (a global in the VBA).
 * @param index     The year's inkomstindex.
 * @param index2021 Inkomstindex in 2021.
 * @param andel     Withdrawal share.
 * @param ftid      Försäkringstid.
 * @param marginal  0 applies the rounding, 1 removes it (a global in the VBA).
 */
export function tillagg(
  underl: number,
  year: number,
  born: number,
  index = 202.84,
  index2021 = 186.52,
  andel = 1,
  ftid = 40,
  marginal = 0,
): number {
  let result = 0;
  if (year < 2021 || ftid < 1) return result;

  // Every year's change in the thresholds compounds, hence the exponent.
  const i = index / index2021 / Math.pow(1.016, year - 2021);
  const u = underl / i;

  if (u < 108_000 || u >= 204_000) return result;

  let tlIpt: number;
  if (u < 132_000) {
    tlIpt = (u - 108_000) * 0.3;
    // NOTE: this rounding adds 300 after flooring to a 600 step, so it can
    // exceed the unrounded amount by up to 300 kr a year. Kept as written.
    if (marginal === 0) tlIpt = vbaInt(((u - 108_000) * 0.3) / 600) * 600 + 300;
  } else if (u <= 168_000) {
    // NOTE: the flat middle bracket is not rounded at all.
    tlIpt = 7200;
  } else {
    tlIpt = 7200 - (u - 168_000) * 0.2;
    if (marginal === 0) tlIpt = vbaInt((7200 - (u - 168_000) * 0.2) / 600) * 600 + 300;
  }

  result = vbaSingle(tlIpt);
  if (andel <= 1) result = vbaSingle(result * andel);

  // The years of residence needed for a full supplement fall with age.
  if (vbaInt(born) > 1944) {
    if (ftid < 40) result = vbaSingle((result * ftid) / 40);
  } else if (vbaInt(born) > 1937) {
    if (ftid < 35) result = vbaSingle((result * ftid) / 35);
  } else if (vbaInt(born) > 1924) {
    if (ftid < 30) result = vbaSingle((result * ftid) / 30);
  } else if (vbaInt(born) > 1914) {
    const full = 20 + vbaInt(born) - 1915;
    if (ftid < full) result = vbaSingle((result * ftid) / full);
  } else {
    if (ftid < 20) result = vbaSingle((result * ftid) / 20);
  }

  return result;
}
