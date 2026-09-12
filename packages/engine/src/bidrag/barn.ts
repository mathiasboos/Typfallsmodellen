/**
 * Benefits that follow from having children at home.
 *
 * Port of `CalcAntalBarn`, `CalcBarnPerAlder`, `barnbidraget` and `ustod` from
 * Bidrag.bas.
 *
 * The model offers up to four children, each given by its year of birth, and
 * these turn that into the counts the benefits are calculated on -- a plain
 * total for the child allowance, and a split into eight age bands for the
 * social assistance norm.
 */

import { vbaSingle } from "../vba/math.js";

/** The eight age bands Socialstyrelsen's riksnorm is written in. */
export interface BarnPerAlder {
  /** 0 years old. */
  b1: number;
  /** 1-2 years. */
  b2: number;
  /** 3 years. */
  b3: number;
  /** 4-6 years. */
  b4: number;
  /** 7-10 years. */
  b5: number;
  /** 11-14 years. */
  b6: number;
  /** 15-18 years. */
  b7: number;
  /** 19-20 years. */
  b8: number;
}

/**
 * How many children are at home in a given year.
 *
 * Each child counts for twenty years from its birth year. A child's year of
 * birth is 0 when the slot is unused, which no calendar year matches.
 *
 * NOTE: declared `As Single` in the VBA.
 *
 * NOTE: the fourth child's window is tested against the *third* child's birth
 * year -- `ar < (barn3 + 20)` where every other block uses its own. A
 * copy-paste slip, kept as written: a fourth child born more than twenty years
 * after the third never counts, and one born earlier drops out on the third
 * child's schedule instead of its own.
 *
 * @param ar    The calendar year.
 * @param barn1 Year of birth of the first child, 0 for none.
 */
export function CalcAntalBarn(
  ar: number,
  barn1: number,
  barn2: number,
  barn3: number,
  barn4: number,
): number {
  let result = 0;

  if (ar >= barn1 && ar < barn1 + 20) {
    result = 1;
    if (ar === barn1 + 19) result = 0;
  }
  if (ar >= barn2 && ar < barn2 + 20) {
    if (ar >= barn2) result = result + 1;
    if (ar === barn2 + 19) result = Math.max(result - 1, 0);
  }
  if (ar >= barn3 && ar < barn3 + 20) {
    if (ar >= barn3) result = result + 1;
    if (ar === barn3 + 19) result = Math.max(result - 1, 0);
  }
  // NOTE: `barn3 + 20`, not `barn4 + 20`. See the note above.
  if (ar >= barn4 && ar < barn3 + 20) {
    if (ar >= barn4) result = result + 1;
    if (ar === barn4 + 19) result = Math.max(result - 1, 0);
  }

  return vbaSingle(result);
}

/**
 * The same children, split into the eight age bands the riksnorm uses.
 *
 * The VBA returns these through eight `ByRef` arguments; here they come back as
 * one object. Ages are taken from the birth year alone, with no regard to when
 * in the year a birthday falls -- the VBA says as much.
 *
 * The VBA writes the same block out four times, once per child. The first
 * assigns (`b1 = 1`) where the others add, which comes to the same thing from
 * an all-zero start, so the four are folded into one loop here.
 *
 * @param ar    The calendar year.
 * @param barn1 Year of birth of the first child, 0 for none.
 */
export function CalcBarnPerAlder(
  ar: number,
  barn1: number,
  barn2: number,
  barn3: number,
  barn4: number,
): BarnPerAlder {
  const out: BarnPerAlder = { b1: 0, b2: 0, b3: 0, b4: 0, b5: 0, b6: 0, b7: 0, b8: 0 };

  for (const born of [barn1, barn2, barn3, barn4]) {
    if (ar >= born && ar <= born + 20) {
      if (ar === born) out.b1 += 1;
      else if (ar < born + 3) out.b2 += 1;
      else if (ar < born + 4) out.b3 += 1;
      else if (ar < born + 7) out.b4 += 1;
      else if (ar < born + 11) out.b5 += 1;
      else if (ar < born + 15) out.b6 += 1;
      else if (ar < born + 19) out.b7 += 1;
      else if (ar < born + 21) out.b8 += 1;
    }
  }

  return out;
}

/**
 * Child allowance for a year, including the large-family supplement.
 *
 * The VBA writes the history back to 1974 as a `Select Case` on the year, with
 * the supplement per child held in `xtill`. Adoption support is in the code but
 * switched off.
 *
 * NOTE: the VBA guards on `pblnCloseOrSave` and traps errors into `#VALUE!`.
 * Both are worksheet-function housekeeping and are dropped.
 *
 * @param antal Number of children.
 * @param year  Income year.
 */
export function barnbidraget(antal: number, year = 2013): number {
  // The supplement is held per child; the VBA sizes the array to at least 11.
  const xtill = new Array<number>(Math.max(antal, 11) + 1).fill(0);
  let grund = 0;

  if (year < 1974) {
    grund = 330 * 4;
  } else if (year === 1974) {
    grund = 375 * 4 + 200;
  } else if (year === 1975) {
    grund = 375 * 3 + 450;
  } else if (year === 1976) {
    grund = 4 * 450;
  } else if (year === 1977) {
    grund = 2 * 450 + 2 * 525;
  } else if (year === 1978) {
    grund = 1 * 525 + 3 * 565;
  } else if (year === 1979) {
    grund = 4 * 625;
  } else if (year === 1980) {
    grund = 3 * 700 + 750;
  } else if (year === 1981) {
    grund = 750 * 4;
  } else if (year === 1982) {
    grund = 750 * 4;
    xtill[3] = 750;
    xtill[4] = 1500;
    for (let i = 5; i <= antal; i += 1) xtill[i] = xtill[i - 1]! + 1500;
  } else if (year < 1985) {
    grund = 825 * 4;
    xtill[3] = 825;
    xtill[4] = 1650;
    for (let i = 5; i <= antal; i += 1) xtill[i] = xtill[i - 1]! + 1650;
  } else if (year < 1987) {
    // Förlängt barnbidrag and some adopted children are added in 1985.
    grund = 4800;
    xtill[3] = 1200;
    xtill[4] = 2400;
    for (let i = 5; i <= antal; i += 1) xtill[i] = xtill[i - 1]! + 2400;
  } else if (year < 1988) {
    grund = 5820;
    xtill[3] = 1455;
    for (let i = 4; i <= antal; i += 1) xtill[i] = xtill[i - 1]! + 2890;
  } else if (year < 1989) {
    grund = 5820;
    xtill[3] = 1455;
    xtill[4] = 1.6 * grund;
    for (let i = 5; i <= antal; i += 1) xtill[i] = 1.6 * grund;
  } else if (year < 1990) {
    grund = 5820;
    xtill[3] = 1455;
    xtill[4] = 1.9 * grund;
    xtill[5] = 2.4 * grund;
    for (let i = 6; i <= antal; i += 1) xtill[i] = 1.6 * grund;
  } else if (year < 1991) {
    grund = 6720;
    xtill[3] = 1455;
    xtill[4] = 1.9 * grund;
    xtill[5] = 2.4 * grund;
    for (let i = 6; i <= antal; i += 1) xtill[i] = 1.6 * grund;
  } else if (year < 1994) {
    grund = 750 * 12;
    xtill[3] = 1455;
    xtill[4] = 1 * grund;
    xtill[5] = 1.5 * grund;
    for (let i = 6; i <= antal; i += 1) xtill[i] = 1.5 * grund;
  } else if (year === 1994) {
    grund = 750 * 12;
    xtill[3] = 1455;
    xtill[4] = 1 * grund;
    xtill[5] = 1.25 * grund;
    for (let i = 6; i <= antal; i += 1) xtill[i] = 1.25 * grund;
  } else if (year < 1996) {
    grund = 750 * 12;
    xtill[3] = 2400;
    xtill[4] = 7200;
    for (let i = 5; i <= antal; i += 1) xtill[i] = 9000;
  } else if (year < 1998) {
    // The large-family supplement is abolished for children born after 1995,
    // which the VBA notes and does not model.
    grund = 640 * 12;
    xtill[3] = 2400;
    xtill[4] = 7200;
    for (let i = 5; i <= antal; i += 1) xtill[i] = 9000;
  } else if (year < 2000) {
    // And comes back.
    grund = 750 * 12;
    xtill[3] = 2400;
    xtill[4] = 7200;
    for (let i = 5; i <= antal; i += 1) xtill[i] = 9000;
  } else if (year < 2001) {
    grund = 850 * 12;
    xtill[1] = 0;
    xtill[2] = 0;
    xtill[3] = 227 * 12;
    xtill[4] = 680 * 12;
    xtill[5] = 850 * 12;
    for (let i = 6; i <= antal; i += 1) xtill[i] = xtill[5]!;
  } else if (year < 2006) {
    grund = 950 * 12;
    xtill[1] = 0;
    xtill[2] = 0;
    xtill[3] = 254 * 12;
    xtill[4] = 760 * 12;
    xtill[5] = 950 * 12;
    for (let i = 6; i <= antal; i += 1) xtill[i] = xtill[5]!;
  } else if (year < 2011) {
    grund = 12_600;
    xtill[1] = 0;
    xtill[2] = 1200;
    xtill[3] = 4248;
    xtill[4] = 10_320;
    xtill[5] = 12_600;
    for (let i = 6; i <= antal; i += 1) xtill[i] = xtill[5]!;
  } else if (year <= 2016) {
    grund = 1050 * 12;
    xtill[1] = 0;
    xtill[2] = 150 * 12;
    xtill[3] = 454 * 12;
    xtill[4] = 1010 * 12;
    xtill[5] = 1250 * 12;
    for (let i = 6; i <= antal; i += 1) xtill[i] = xtill[5]!;
  } else {
    grund = 1050 * 12;
    xtill[1] = 0;
    xtill[2] = 150 * 12;
    xtill[3] = 580 * 12;
    xtill[4] = 1010 * 12;
    xtill[5] = 1250 * 12;
    for (let i = 6; i <= antal; i += 1) xtill[i] = xtill[5]!;
  }

  // The 2018 raise landed in April, so that year blends the two levels.
  if (year === 2018) grund = 1050 * 3 + 1250 * 9;
  if (year > 2018) grund = 1250 * 12;

  grund = antal * grund;

  // The supplement is paid from the second child onwards.
  let fbtill = 0;
  for (let i = 2; i <= antal; i += 1) fbtill = fbtill + xtill[i]!;

  // Adoption support is switched off, as the VBA leaves it.
  const adoptionsbidrag = 0;
  if (year > 2016) return grund + fbtill + adoptionsbidrag * 40_000;
  return grund + fbtill + adoptionsbidrag * 75_000;
}

/**
 * Underhållsstöd, paid only to a single parent.
 *
 * NOTE: the amount does not vary by the child's age here, though the real rules
 * have done so for some years -- the VBA says as much.
 *
 * @param barn    Number of children.
 * @param ensamst 1 for a single parent; anything else pays nothing.
 * @param year    Income year.
 */
export function ustod(barn: number, ensamst = 1, year = 2013): number {
  let result = 0;

  if (ensamst !== 1) return result;

  if (year < 1996) {
    result = 1173 * 12;
  } else if (year < 2006) {
    result = 1173 * 12;
  } else if (year < 2015) {
    result = 1273 * 12;
  } else if (year < 2017) {
    result = 1573 * 12;
  } else if (year < 2022) {
    result = 1673 * 12;
  } else {
    result = 1823 * 12;
  }

  return result * barn;
}
