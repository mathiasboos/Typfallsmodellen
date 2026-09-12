/**
 * Bostadsbidrag for families with children.
 *
 * Port of `bobid` from Bidrag.bas.
 *
 * A fixed "särskilt bidrag" per family, plus a share of the housing cost
 * between a lower and an upper limit, reduced by 20 öre per krona of income
 * above a floor. The limits depend on the number of children.
 */

import { vbaArray } from "../vba/ageArray.js";
import { vbaInt } from "../vba/math.js";

/**
 * Bostadsbidrag for a year, as an annual amount.
 *
 * NOTE: `ungdom` is declared but never read -- the VBA notes that the rules for
 * young adults without children are missing. It is kept so the signature
 * matches the call sites.
 *
 * NOTE: the VBA guards on `pblnCloseOrSave` and traps errors into `#VALUE!`.
 * Both are worksheet-function housekeeping and are dropped.
 *
 * @param vuxna    Number of adults in the household, 1 or 2.
 * @param inkomst  The household head's income.
 * @param make_ink The spouse's income.
 * @param barn     Number of children at home.
 * @param Uboende  Housing cost; see the note on the monthly/annual heuristic.
 * @param yta      Dwelling size in square metres.
 * @param marginal 0 applies the rounding, 1 removes it.
 * @param ungdom   Declared but never read.
 * @param year     Income year.
 */
export function bobid(
  vuxna: number,
  inkomst: number,
  make_ink: number,
  barn: number,
  Uboende: number,
  yta = 80,
  marginal = 0,
  ungdom = 0,
  year = 2013,
): number {
  // Housing cost limits, by family size: lower, middle, upper.
  let xn1 = 0;
  let xm1 = 0;
  let xo1 = 0;
  let xn2 = 0;
  let xm2 = 0;
  let xo2 = 0;
  let xn3 = 0;
  let xm3 = 0;
  let xo3 = 0;

  // The särskilda bidrag, a fixed monthly amount by family size.
  let xg1 = 0;
  let xg2 = 0;
  let xg3 = 0;

  let bhyra1 = 0;
  let bhyra2 = 0;
  let bhyra3 = 0;
  let bhyra4 = 0;
  let bhyra5 = 0;

  /** The income floor above which the allowance is reduced. */
  let xfmb = 0;
  /** The reduction rate. */
  const xrfmb = 0.2;

  if (year < 1995) {
    xn1 = 3200;
    xm1 = 4800;
    // NOTE: `xo1 = xm2` is read before `xm2` is assigned three lines below, so
    // the upper limit for a one-child family is 0. Kept as written.
    xo1 = xm2;
    xn2 = 3200;
    xm2 = 5300;
    xo2 = xm2;
    xn3 = 5300;
    xm3 = 6100;
    xo3 = xm3;
    xg1 = 600;
    xg2 = 900;
    xg3 = 1200;
    bhyra1 = 2900;
    bhyra2 = 3200;
    bhyra3 = 3500;
    bhyra4 = 3800;
    bhyra5 = 4100;
    xfmb = 110_000;
  } else if (year < 1996) {
    xn1 = 3300;
    xm1 = 5200;
    // NOTE: the same premature read as above.
    xo1 = xm2;
    xn2 = 3300;
    xm2 = 5800;
    xo2 = xm2;
    xn3 = 5400;
    xm3 = 6500;
    xo3 = xm3;
    xg1 = 600;
    xg2 = 900;
    xg3 = 1200;
    bhyra1 = 3000;
    bhyra2 = 3300;
    bhyra3 = 3600;
    bhyra4 = 3900;
    bhyra5 = 4200;
    xfmb = 115_000;
  } else if (year < 2017) {
    xn1 = 2000;
    xm1 = 3000;
    xo1 = 5300;
    xn2 = 2000;
    xm2 = 3300;
    xo2 = 5900;
    xn3 = 2000;
    xm3 = 3600;
    xo3 = 6600;
    xg1 = 600;
    xg2 = 900;
    xg3 = 1200;
    bhyra1 = 3000;
    bhyra2 = 3300;
    bhyra3 = 3600;
    bhyra4 = 3900;
    bhyra5 = 4200;
    xfmb = 117_000;
  } else {
    xn1 = 2000;
    xm1 = 5300;
    xo1 = xm1;
    xn2 = 2000;
    xm2 = 5900;
    xo2 = xm2;
    xn3 = 2000;
    xm3 = 6600;
    xo3 = xm3;
    xg1 = 1300;
    xg2 = 1750;
    xg3 = 2350;
    bhyra1 = 3000;
    bhyra2 = 3300;
    bhyra3 = 3600;
    bhyra4 = 3900;
    bhyra5 = 4200;
    // From 2017 the floor is per person.
    xfmb = 117_500;
  }

  const xg4 = xg3;
  const xg5 = xg3;
  if (year >= 2017) xfmb = 127_000;

  if (year > 2012) {
    xn1 = 1400;
    xn2 = 1400;
    xn3 = 1400;
  }

  // NOTE: a monthly/annual heuristic, as in SBTP.
  if (Uboende > 50_000) Uboende = Uboende / 12;

  if (year > 2017) {
    // Changed on 1 March.
    xfmb = 135_000;
    if (year === 2019) xfmb = 142_000;
    if (year === 2020) xfmb = 148_000;
    if (year > 2020) xfmb = 150_000;
  }

  // The dwelling size the allowance is calculated on, by family size.
  const boyta1 = 80;
  const boyta2 = 100;
  const boyta3 = 120;

  let zbost = Uboende;

  if (barn === 1 && yta > boyta1) {
    if (Uboende > bhyra1) zbost = ((Uboende - bhyra1) * boyta1) / yta;
  }
  if (barn === 2 && yta > boyta2) {
    if (Uboende > bhyra2) zbost = ((Uboende - bhyra2) * boyta2) / yta;
  }
  if (barn === 3 && yta > boyta3) {
    if (Uboende > bhyra3) zbost = ((Uboende - bhyra3) * boyta3) / yta;
  }
  // NOTE: four and five children are compared against `boyta3` and scaled by it
  // too, so the 140 and 160 square metre limits the VBA declares are never
  // used. Kept as written.
  if (barn === 4 && yta > boyta3) {
    if (Uboende > bhyra4) zbost = ((Uboende - bhyra4) * boyta3) / yta;
  }
  if (barn > 4 && yta > boyta3) {
    if (Uboende > bhyra5) zbost = ((Uboende - bhyra5) * boyta3) / yta;
  }

  // The share of the cost covered between the lower and middle, and the middle
  // and upper, limits.
  let xand1: number;
  const xand2 = 0.5;
  if (year <= 2004) {
    xand1 = 0.5;
  } else if (year <= 2012) {
    xand1 = 0.75;
  } else {
    xand1 = 0.5;
  }

  /**
   * NOTE: `Array()` is 1-based here, because Bidrag.bas declares
   * `Option Base 1` -- so the leading 0 the VBA writes as a zero-index pad
   * lands on index 1, and every lookup below is shifted one place. A one-child
   * family reads 0 for all three limits and for the särskilda bidrag, so it
   * gets nothing; a two-child family reads the one-child figures; and `xn3`,
   * `xm3`, `xo3`, `xg4` and `xg5` are never reached at all.
   *
   * This is almost certainly not what the author meant, but it is what the
   * workbook computes, so it is what the port computes. See docs/VBA-MAPPING.md.
   */
  const ZN = vbaArray(0, xn1, xn2, xn3);
  const ZM = vbaArray(0, xm1, xm2, xm3);
  const ZO = vbaArray(0, xo1, xo2, xo3);
  const ZG = vbaArray(0, xg1, xg2, xg3, xg4, xg5);

  let ibostbh = 0;

  let csbink: number;
  let csbinkM: number;
  if (vuxna === 1) {
    csbink = inkomst;
    csbinkM = 0;
  } else {
    csbink = inkomst;
    csbinkM = make_ink;
  }

  const antbarn = 0;
  const hemmabarn = barn;

  const zbarnsum = hemmabarn + antbarn;
  /** Children counted when setting the housing cost limits. */
  const zbantbrn = Math.min(3, zbarnsum);
  /** Children counted for the särskilda bidrag. */
  const zbantsar = Math.min(3, hemmabarn);

  // The monthly housing cost is rounded down to a whole 25 kronor.
  if (marginal === 0) zbost = 25 * vbaInt(zbost / 25);

  if (zbarnsum > 0) {
    if (zbost <= ZN(zbantbrn)) {
      ibostbh = 12 * ZG(zbantsar);
    } else if (zbost <= ZM(zbantbrn)) {
      ibostbh = 12 * (ZG(zbantsar) + (zbost - ZN(zbantbrn)) * xand1);
    } else if (zbost <= ZO(zbantbrn)) {
      ibostbh =
        12 *
        (ZG(zbantsar) +
          (ZM(zbantbrn) - ZN(zbantbrn)) * xand1 +
          (zbost - ZM(zbantbrn)) * xand2);
    } else {
      ibostbh =
        12 *
        (ZG(zbantsar) +
          (ZM(zbantbrn) - ZN(zbantbrn)) * xand1 +
          (ZO(zbantbrn) - ZM(zbantbrn)) * xand2);
    }
  }

  // The reduction: 20 öre per krona above the floor, which a couple splits.
  if (vuxna === 1) {
    if (csbink > xfmb && ibostbh > 0) {
      ibostbh = ibostbh - xrfmb * (csbink - xfmb);
    }
  } else if (vuxna === 2) {
    if (csbink > xfmb / 2 && ibostbh > 0) {
      ibostbh = ibostbh - xrfmb * (csbink - xfmb / 2);
    }
    if (csbinkM > xfmb / 2 && ibostbh > 0) {
      ibostbh = ibostbh - xrfmb * (csbinkM - xfmb / 2);
    }
  }

  if (ibostbh < 0) ibostbh = 0;

  // Never more than the annual housing cost.
  if (ibostbh > 12 * Uboende) ibostbh = Uboende * 12;

  // A temporary raise proposed in May 2020 and assumed to pass, for half a year.
  if (ibostbh > 0 && year === 2020) ibostbh = ibostbh + ibostbh * 0.25 * (6 / 12);

  // At least 100 kr a month, and paid in whole hundreds.
  if (marginal === 0 && ibostbh < 1200) ibostbh = 0;
  ibostbh = ibostbh / 12;
  if (marginal === 0) ibostbh = vbaInt(ibostbh / 100) * 100;

  return 12 * ibostbh;
}
