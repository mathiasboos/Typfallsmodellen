/**
 * Bostadstillägg, särskilt bostadstillägg and äldreförsörjningsstöd.
 *
 * Port of `BTP`, `SBTP` and `btp_sbtp` from Bidrag.bas.
 *
 * `BTP` pays a share of the rent, reduced against a "reduceringsinkomst" above
 * a free amount. `SBTP` is a second, needs-tested layer that tops the household
 * up to a reasonable standard of living -- and the same function also computes
 * äldreförsörjningsstöd, for those whose residence history leaves them without
 * a guarantee pension. `btp_sbtp` adds the two and rounds.
 */

import { vbaCLng, vbaInt, wsMax } from "../vba/math.js";
import { avdragxx } from "../skatt/grundavdrag.js";
import { PublicAvg } from "../skatt/reduktioner.js";
import { riktage } from "../pension/contributions.js";
import type { BtpContext, SbtpContext } from "./types.js";

/**
 * Bostadstillägg for a year, annual amounts throughout.
 *
 * NOTE: `marginal`, `mgarp`, `month`, `ftid`, `index`, `uttag` and `ansokt` are
 * declared but never read. They are kept so the signature matches the VBA call
 * sites, which do pass them.
 *
 * @param inkomst  Income.
 * @param inkomstm The spouse's income, for a cohabiting household.
 * @param hyra     Housing cost.
 * @param gift     1 cohabiting, 0 single.
 * @param pbb      The year's prisbasbelopp.
 * @param context  `RNG_dela`.
 * @param ap       1 if an old-age pensioner.
 * @param apm      1 if the spouse is. Declared but never read.
 * @param form     The household's wealth.
 * @param arb      Earned income.
 * @param ArbM     The spouse's earned income.
 * @param maxhyra  A replacement ceiling on the rent, used with `wyear`.
 * @param year     Income year.
 * @param wyear    The year from which the thresholds follow earnings.
 * @param IBB      That year's inkomstbasbelopp.
 * @param kvoten   The pbb/IBB ratio at `wyear`.
 * @param bald     Age.
 * @param TJP      Occupational pension.
 * @param tjpm     The spouse's occupational pension.
 * @param garp     Garantipension.
 * @param garpm    The spouse's garantipension.
 * @param IBB0     The inkomstbasbelopp at `wyear`.
 */
export function BTP(
  inkomst: number,
  inkomstm: number,
  hyra: number,
  gift: number,
  pbb: number,
  context: BtpContext,
  ap = 1,
  apm = 1,
  form = 0,
  arb = 0,
  ArbM = 0,
  marginal = 0,
  maxhyra = 0,
  year = 2014,
  wyear = 2100,
  IBB = 0,
  kvoten = 1,
  bald = 65,
  TJP = 0,
  tjpm = 0,
  garp = 0,
  garpm = 0,
  IBB0 = 0,
  mgarp = 0,
  month = 12,
  ftid = 40,
  index = 100,
  uttag = 1,
  ansokt = 1,
): number {
  let result = 0;
  let max = 0;
  let maxm = 0;

  // Kommunalt bostadstillägg was introduced on 1 January 1978.
  if (year < 1978) return result;

  if (year >= wyear && kvoten < 1) {
    // Swap pbb for IBB and scale the thresholds back to the old level.
    pbb = IBB;
    pbb = pbb * kvoten;
  }

  /** The highest annual rent the supplement is calculated on. */
  let maxohyra = 0;
  /** The spouse's own BTP. */
  let BTPm = 0;

  // The share of the rent that is covered, in up to four bands.
  let PAR = 0.9;
  let par2 = 0;
  let par3 = 0;
  let par4 = 0;

  if (year < 1995) {
    PAR = 0.83;
    maxohyra = hyra > 150 ? vbaCLng(3500) * 12 : 0;
  } else if (year < 2001) {
    PAR = 0.85;
    maxohyra = hyra > 100 ? vbaCLng(3900) * 12 : 0;
    if (year === 1998) PAR = 0.85;
    if (year === 1999) PAR = 0.9;
  } else if (year <= 2003) {
    PAR = 0.9;
    maxohyra = vbaCLng(4500) * 12;
    if (year > 2001) PAR = 0.91;
  } else if (year <= 2005) {
    PAR = 0.91;
    maxohyra = vbaCLng(4670) * 12;
  } else if (year === 2006) {
    // Raised partway through the year; assuming even income and rent, this
    // understates a pension taken mid-year.
    PAR = (0.91 * 3 + 0.93 * 9) / 12;
    if (ap === 0) PAR = 0.91;
    maxohyra = vbaCLng(4850) * 12;
  } else if (year === 2007) {
    PAR = 0.93;
    if (ap === 0) PAR = 0.91;
    maxohyra = vbaCLng(5000) * 12;
  } else {
    PAR = 0.93;
    maxohyra = vbaCLng(5000) * 12;
    if (ap === 0 && year < 2009) PAR = 0.91;
  }

  if (year === 2015 && ap === 1) {
    PAR = (0.93 * 8) / 12 + (0.95 * 4) / 12;
  } else if (year > 2015 && ap === 1) {
    PAR = 0.95;
  } else if (year > 2017 && ap === 0) {
    // NOTE: dead. The `year > 2017` block immediately below overwrites PAR
    // unconditionally, so this 0.95 never reaches the calculation. Kept.
    PAR = 0.95;
  }

  if (year > 2017) {
    // 96% of the rent, and 70% of the part between 5 000 and 5 600 a month.
    PAR = 0.96;
    par2 = 0.7;
    maxohyra = vbaCLng(5600) * 12;
  }

  if (year > 2019) {
    // Introduced on 1 December 2019, applied here from the turn of the year.
    PAR = 1;
    par2 = 0.9;
    par3 = 0.7;
    maxohyra = vbaCLng(7000) * 12;

    if (year >= 2022) {
      maxohyra = vbaCLng(7500) * 12;
      par4 = 0.5;
    }
  }

  if (maxhyra > 0 && year >= wyear) maxohyra = maxhyra * maxohyra;
  if (hyra > maxohyra) hyra = maxohyra;

  /** The spouse's share of the rent. */
  let hyram = 0;

  // If the thresholds follow earnings, express the rent in the old money.
  if (year >= wyear) hyra = (hyra * IBB0) / IBB;

  if (gift === 1) {
    hyra = hyra / 2;
    hyram = hyra;
  }

  if (year < 2018) {
    max = hyra * PAR;
    maxm = hyram * PAR;
  } else if (year < 2020) {
    if (gift === 0) {
      if (hyra < 60_000) {
        max = hyra * PAR;
      } else if (hyra < 67_200) {
        max = 60_000 * PAR + (hyra - 60_000) * par2;
      } else {
        max = 60_000 * PAR + (67_200 - 60_000) * par2;
      }
    } else {
      if (hyra < 30_000) {
        max = hyra * PAR;
      } else if (hyra < 33_600) {
        max = 30_000 * PAR + (hyra - 30_000) * par2;
      } else {
        max = 30_000 * PAR + (33_600 - 30_000) * par2;
      }
      maxm = max;
    }
  } else if (year < 2022) {
    if (gift === 0) {
      if (hyra < 36_001) {
        max = hyra * PAR;
      } else if (hyra < 60_001) {
        max = 36_000 * PAR + (hyra - 36_000) * par2;
      } else if (hyra < 84_001) {
        max = 36_000 * PAR + (60_000 - 36_000) * par2 + (hyra - 60_000) * par3;
      } else {
        max = 36_000 * PAR + (60_000 - 36_000) * par2 + (84_000 - 60_000) * par3;
      }
    } else {
      if (hyra < 18_001) {
        max = hyra * PAR;
      } else if (hyra < 30_001) {
        max = 18_000 * PAR + (hyra - 18_000) * par2;
      } else if (hyra < 42_001) {
        max = 18_000 * PAR + (30_000 - 18_000) * par2 + (hyra - 30_000) * par3;
      } else {
        max = 18_000 * PAR + (30_000 - 18_000) * par2 + (42_000 - 30_000) * par3;
      }
      maxm = max;
    }
  } else {
    if (gift === 0) {
      if (hyra < 36_001) {
        max = hyra * PAR;
      } else if (hyra < 60_001) {
        max = 36_000 * PAR + (hyra - 36_000) * par2;
      } else if (hyra < 84_001) {
        max = 36_000 * PAR + (60_000 - 36_000) * par2 + (hyra - 60_000) * par3;
      } else if (hyra < 90_001) {
        max =
          36_000 * PAR +
          (60_000 - 36_000) * par2 +
          (84_000 - 60_000) * par3 +
          (hyra - 84_000) * par4;
      } else {
        max =
          36_000 * PAR +
          (60_000 - 36_000) * par2 +
          (84_000 - 60_000) * par3 +
          (90_000 - 84_000) * par4;
      }
    } else {
      if (hyra < 18_001) {
        max = hyra * PAR;
      } else if (hyra < 30_001) {
        max = 18_000 * PAR + (hyra - 18_000) * par2;
      } else if (hyra < 42_001) {
        max = 18_000 * PAR + (30_000 - 18_000) * par2 + (hyra - 30_000) * par3;
      } else if (hyra < 45_001) {
        max =
          18_000 * PAR +
          (30_000 - 18_000) * par2 +
          (42_000 - 30_000) * par3 +
          (hyra - 42_000) * par4;
      } else {
        max =
          18_000 * PAR +
          (30_000 - 18_000) * par2 +
          (42_000 - 30_000) * par3 +
          (45_000 - 42_000) * par4;
      }
      maxm = max;
    }
  }

  /** Extra consumption support, paid alongside the housing supplement. */
  let Extra = 0;

  // NOTE: the age guard applies only to 2012-2021. From 2022 the supplement is
  // added at any age. Kept as written.
  if (year >= 2012 && bald > 64 && year < 2022) {
    Extra = (340 * 12) / (1 + gift);
  }

  if (year === 2022) {
    // Raised by 200 kr on 1 August, and again to 840 from the turn of the year.
    Extra = (540 * 7 + 840 * 5) / (gift + 1);
  }

  if (year > 2022) Extra = (840 * 12) / (gift + 1);

  max = max + Extra;
  maxm = maxm + Extra * gift;

  if (year >= wyear) {
    // And back into the year's own money.
    max = (max * IBB) / IBB0;
    maxm = (maxm * IBB) / IBB0;
  }

  // NOTE: only wealth *above* the threshold is scaled -- below it, `form` is
  // left untouched and then added to income in full, rather than disregarded.
  // SBTP does the opposite. Kept as written.
  if (year < 2001) {
    if (gift === 1) {
      if (form > 75_000) form = 0.05 * (form - 75_000);
    } else {
      if (form > 120_000) form = 0.1 * (form - 120_000);
    }
  } else {
    if (form > 100_000) form = 0.15 * (form - 100_000);
  }
  if (form < 0) form = 0;

  inkomst = inkomst + form;
  inkomstm = inkomstm + form;

  /** Free amount, expressed in prisbasbelopp. */
  let fri = 0;
  let Red = 0;
  let Redm = 0;

  if (year >= 2014 && (arb > 0 || ArbM > 0)) {
    if (year >= wyear) {
      arb = wsMax(arb - (24_000 * IBB) / IBB0, 0);
      ArbM = wsMax(ArbM - (24_000 * IBB) / IBB0, 0);
    } else {
      arb = wsMax(arb - 24_000, 0);
      ArbM = wsMax(ArbM - 24_000, 0);
    }
  }

  if (gift === 0) {
    if (ap === 1) {
      fri = pbb * 2.17;
      if (year > 2019) fri = 2.181 * pbb;
      if (year === 2022) fri = ((7 * 2.181 + 5 * 2.43) * pbb) / 12;
      if (year > 2022) fri = 2.43 * pbb;
    } else {
      fri = pbb * 2.4;
    }

    // Public pensions count in full, earned income at 50% -- 80% before 2008.
    if (year < 2008) {
      Red = inkomst - arb - TJP + 0.8 * arb + 0.8 * TJP - fri;
    } else {
      Red = inkomst - arb - TJP + 0.5 * arb + 0.8 * TJP - fri;
    }

    if (year > 2019) {
      Red = garp + (inkomst - garp) * 0.93 - fri;
    }

    if (Red < 0) Red = 0;
  }

  if (gift === 1) {
    if (ap === 1) {
      fri = pbb * 1.935;
      if (year > 2019) fri = 1.951 * pbb;
      if (year === 2022) fri = ((7 * 1.951 + 5 * 2.2) * pbb) / 12;
      if (year > 2022) fri = 2.2 * pbb;
    } else {
      fri = pbb * 2.4;
    }

    if (year < 2008) {
      Red = inkomst - arb - TJP + 0.8 * arb + 0.8 * TJP - fri;
      Redm = inkomstm - ArbM - tjpm + 0.8 * ArbM + 0.8 * tjpm - fri;
    } else {
      Red = inkomst - arb - TJP + 0.5 * arb + 0.8 * TJP - fri;
      Redm = inkomstm - ArbM - tjpm + 0.5 * ArbM + 0.8 * tjpm - fri;
    }

    if (year > 2019) {
      Red = garp + (inkomst - garp) * 0.93 - fri;
      Redm = garpm + (inkomstm - garpm) * 0.93 - fri;
    }

    if (Red < 0) Red = 0;
    if (Redm < 0) Redm = 0;
    // The reduction is shared between the spouses.
    Red = vbaInt((Red + Redm) / 2);
  }

  if (Red < pbb) {
    result = max - 0.62 * Red;
  } else {
    result = max - (Red - pbb) * 0.5 - pbb * 0.62;
  }

  if (year > 2019) {
    result = max - 0.62 * Red;
  }

  if (result < 0) result = 0;
  // NOTE: assigned only once the amount is final, so under `dela = 1` a
  // cohabiting pensioner's supplement doubles. Kept as written.
  if (gift === 1 && ap === 1) BTPm = result;

  if (context.dela === 1) {
    // SBTP reduces the whole amount.
    result = result + BTPm;
  } else if (context.dela === 2) {
    // The VBA writes `BTP = BTP` here. Kept as an explicit branch so the three
    // settings read as the three cases they are.
  } else {
    result = (result + BTPm) / (gift + 1);
  }

  // NOTE: the VBA's closing comment promises rounding to whole kronor here, and
  // suppression of monthly amounts under 25 kr. Both happen in `btp_sbtp`
  // instead. There is also no reduction for försäkringstid.
  return result;
}

/**
 * Särskilt bostadstillägg and äldreförsörjningsstöd, in one function.
 *
 * Both compare the household's disposable income after rent against a
 * "skälig levnadsnivå" and pay the difference. `ftid < 1` -- no entitlement to
 * the housing supplement at all -- skips straight to the äldreförsörjningsstöd
 * part.
 *
 * @param inkomst    Income.
 * @param hyra       Housing cost; see the note on the monthly/annual heuristic.
 * @param gift       1 cohabiting, 0 single.
 * @param btpb       The bostadstillägg to be added in.
 * @param avdrag     Grundavdrag, used for the free amount net of tax.
 * @param context    Mcalc's globals and the workbook settings.
 * @param skattesats The municipal tax rate.
 * @param ap         1 if an old-age pensioner.
 * @param form       Wealth.
 * @param pbb        The year's prisbasbelopp.
 * @param maxhyra    A replacement ceiling on the rent, used with `wyear`.
 * @param year       Income year.
 * @param wyear      The year from which the thresholds follow earnings.
 * @param IBB        That year's inkomstbasbelopp.
 * @param kvoten     The pbb/IBB ratio at `wyear`.
 * @param bald       Age.
 * @param kapital    Capital income.
 * @param inkomstm   The spouse's income.
 * @param marginal   0 applies the rounding, 1 removes it.
 * @param baldm      The spouse's age; below 0 means the same as `bald`.
 * @param ftid       Below 1, only äldreförsörjningsstöd is calculated.
 */
export function SBTP(
  inkomst: number,
  hyra: number,
  gift: number,
  btpb: number,
  avdrag: number,
  context: SbtpContext,
  skattesats = 0.316,
  ap = 1,
  form = 0,
  pbb = 42_800,
  maxhyra = 0,
  year = 2014,
  wyear = 21_000,
  IBB = 0,
  kvoten = 0,
  bald = 65,
  kapital = 0,
  inkomstm = 0,
  marginal = 0,
  baldm = -99,
  ftid = 1,
): number {
  let result = 0;

  if (year < 1978) return result;

  bald = vbaInt(bald);
  baldm = vbaInt(baldm);
  if (baldm < 0) baldm = bald;

  // NOTE: a monthly/annual heuristic. Mcalc passes an annual rent to BTP but a
  // monthly one to SBTP at one of its call sites, and this is what reconciles
  // them. Kept as written.
  if (hyra < 10_000) hyra = hyra * 12;

  if (year >= wyear && kvoten < 1) {
    pbb = IBB;
    pbb = pbb * kvoten;
  }

  // NOTE: unlike BTP, wealth below the threshold becomes 0 here.
  form = ((form - 100_000) * 0.15) / (gift + 1);
  if (form < 0) form = 0;

  let maxohyra: number;
  if (year <= 1997) {
    maxohyra = vbaCLng(5200) * 12;
  } else if (year <= 2000) {
    maxohyra = vbaCLng(5200) * 12;
  } else if (year <= 2004) {
    maxohyra = vbaCLng(5700) * 12;
  } else if (year === 2005) {
    maxohyra = vbaCLng(5870) * 12;
  } else if (year === 2006) {
    maxohyra = vbaCLng(6050) * 12;
  } else if (year < 2010) {
    maxohyra = vbaCLng(6050) * 12;
  } else if (year < 2017) {
    maxohyra = vbaCLng(6200) * 12;
  } else if (year < 2020) {
    maxohyra = vbaCLng(6620) * 12;
  } else {
    maxohyra = vbaCLng(7500) * 12;
  }

  if (maxhyra > 0 && year >= wyear) maxohyra = maxhyra * maxohyra;

  maxhyra = maxohyra / (gift + 1);

  /** Skälig levnadsnivå. */
  let Lev = 0;
  let Besk = 0;
  let Beskm = 0;
  let disp = 0;
  let Dispm = 0;

  // The income the calculation uses is rounded down to a whole hundred.
  let ctxfvi = inkomst;
  let ctxfvim = inkomstm;
  if (marginal === 0) {
    ctxfvi = vbaInt(inkomst / 100) * 100;
    ctxfvim = vbaInt(inkomstm / 100) * 100;
  }

  hyra = hyra / (gift + 1);
  if (hyra > maxhyra) hyra = maxhyra;

  // NOTE: for `ap = 0` in 2012-2017 no branch below assigns Lev, so it stays 0
  // and both benefits come out at nothing. Kept as written.
  if (year < 2009) {
    Lev = gift === 0 ? 1.294 * pbb : 1.084 * pbb;
  } else if (year <= 2011) {
    Lev = gift === 0 ? 1.3546 * pbb : 1.1446 * pbb;
  }

  if (year === 2012 && ap === 1) {
    Lev = gift === 0 ? 1.401 * pbb : 1.191 * pbb;
  } else if (year > 2012 && ap === 1) {
    Lev = gift === 0 ? 1.4468 * pbb : 1.191 * pbb;
  }

  if (year > 2015 && ap === 1) {
    Lev = gift === 0 ? 1.473 * pbb : 1.204 * pbb;
  }

  if (year > 2017) {
    Lev = gift === 0 ? 1.486 * pbb : 1.2105 * pbb;
  }
  if (year > 2021) {
    Lev = gift === 0 ? 1.5357 * pbb : 1.2353 * pbb;
  }

  if (marginal === 0) Lev = vbaInt(Lev / 12) * 12;

  /**
   * NOTE: the VBA's `GoTo A_F_S` jumps over the declaration *and* the
   * assignment of `Xage`, so the äldreförsörjningsstöd block below runs with
   * `Xage = 0` whenever `ftid < 1` -- which picks a different grundavdrag under
   * the 2020 and later rules. Kept as written.
   */
  let Xage = 0;
  /** The spouse's särskilda bostadstillägg. */
  let sbtpm = 0;
  const publicAvgMarginal: 0 | 1 = marginal === 0 ? 0 : 1;

  if (ftid >= 1) {
    Xage = riktage(vbaCLng(year), 1) + 1;

    // NOTE: the taxable amount subtracts a grundavdrag computed on the
    // *unrounded* income from the rounded-down income. Kept as written.
    Besk = ctxfvi - avdragxx(inkomst, pbb, marginal, bald, year, 21_000, IBB, kvoten, context.Iyear, Xage);

    if (gift === 1 && context.rulesFromUtg === 0) {
      Beskm =
        ctxfvim -
        avdragxx(inkomstm, pbb, marginal, bald, year, 21_000, IBB, kvoten, context.Iyear, Xage);
    } else {
      // NOTE: taken for a single household too, and then `rulesFromUtg` -- 0 by
      // default -- is passed as the income *year*. Kept as written.
      Beskm =
        ctxfvim -
        avdragxx(
          inkomstm,
          pbb,
          marginal,
          bald,
          context.rulesFromUtg,
          21_000,
          IBB,
          kvoten,
          context.Iyear,
          Xage,
        );
    }

    const kapitalm = kapital;

    Dispm = 0;
    if (marginal === 0) {
      disp = ctxfvi - vbaInt(Besk * skattesats) + vbaInt(kapital * 0.7);
      if (gift === 1) Dispm = ctxfvim - vbaInt(Beskm * skattesats) + vbaInt(kapitalm * 0.7);
    } else {
      disp = ctxfvi - Besk * skattesats + kapital * 0.7;
      if (gift === 1) Dispm = ctxfvim - Beskm * skattesats + kapitalm * 0.7;
    }

    if (context.age <= context.slutage) {
      const { age, born, vectors } = context;
      disp = disp - PublicAvg(Besk, born, vectors, 0.01, age, publicAvgMarginal, vectors.year(age));
      Dispm =
        Dispm - PublicAvg(Beskm, born, vectors, 0.01, age, publicAvgMarginal, vectors.year(age));
    }
    if (disp < 0) disp = 0;
    if (Dispm < 0) Dispm = 0;

    // Disposable income is compared against a free amount as its floor.
    let fri: number;
    if (year > 2007) {
      fri = gift === 0 ? 2.17 * pbb : 1.935 * pbb;
    } else {
      fri = gift === 0 ? 2.13 * pbb : 1.9 * pbb;
    }

    if (year > 2019) {
      fri = gift === 0 ? 2.181 * pbb : 1.951 * pbb;
      // Cohorts born 1937 and earlier have higher levels.
      if (year - bald < 1938) {
        fri = gift === 0 ? 2.221 * pbb : 1.986 * pbb;
        if (year > 2022) {
          fri = gift === 0 ? 2.47 * pbb : 2.235 * pbb;
        }
      }
    }

    let frinetto: number;
    if (gift === 0) {
      frinetto =
        context.born > 1937
          ? 2.13 * pbb - (2.13 * pbb - avdrag) * skattesats
          : fri - (fri - avdrag) * skattesats;
    } else {
      frinetto =
        context.born > 1937
          ? 1.9 * pbb - (1.9 * pbb - avdrag) * skattesats
          : fri - (fri - avdrag) * skattesats;
    }

    if (marginal === 0) frinetto = vbaInt(frinetto);
    // Which means turning 66 and gaining a larger grundavdrag has an effect of
    // its own, through `disp > fri`.
    if (disp < frinetto) disp = frinetto;
    if (gift === 1 && Dispm < frinetto) Dispm = frinetto;

    // Add wealth and the housing supplement, take off the rent.
    disp = disp + form - hyra + btpb / (gift + 1);
    Dispm = Dispm + form - hyra + btpb / (gift + 1);

    if (disp < 0) disp = 0;
    if (Dispm < 0) Dispm = 0;

    result = Lev - disp;
    sbtpm = 0;
    if (gift === 1) sbtpm = Lev - Dispm;

    if (result < 0) result = 0;
    if (sbtpm < 0) sbtpm = 0;
  }

  // Äldreförsörjningsstöd -- which looks at actual income.
  Besk =
    ctxfvi - avdragxx(inkomst, pbb, marginal, bald, year, 21_000, IBB, kvoten, context.Iyear, Xage);

  // NOTE: wealth is weighted 0.7 alongside capital income when marginal is 0,
  // but counted in full when it is 1. Kept as written.
  if (marginal === 0) {
    disp =
      ctxfvi - vbaInt(Besk * skattesats) + vbaInt(form + kapital) * 0.7 + btpb / (gift + 1) + result - hyra;
  } else {
    disp = ctxfvi - Besk * skattesats + form + kapital * 0.7 + btpb / (gift + 1) + result - hyra;
  }

  if (disp < 0) disp = 0;

  Beskm = 0;
  if (gift === 1 && context.rulesFromUtg === 0) {
    Beskm =
      ctxfvim -
      avdragxx(inkomstm, pbb, marginal, baldm, year, 21_000, IBB, kvoten, context.Iyear, Xage);
  } else {
    Beskm =
      ctxfvim -
      avdragxx(
        inkomstm,
        pbb,
        marginal,
        baldm,
        context.rulesFromUtg,
        21_000,
        IBB,
        kvoten,
        context.Iyear,
        Xage,
      );
  }

  Dispm = 0;
  if (gift === 1) {
    Dispm = ctxfvim - Beskm * skattesats + form + kapital * 0.7 + btpb / (gift + 1) + sbtpm - hyra;
    if (marginal === 0) {
      Dispm =
        ctxfvim -
        vbaInt(Beskm * skattesats) +
        vbaInt(form + kapital) * 0.7 +
        btpb / (gift + 1) +
        sbtpm -
        hyra;
    }
  }
  if (Dispm < 0) Dispm = 0;

  let AFS = Lev - disp;
  if (AFS < 0) AFS = 0;
  let AFSm = 0;
  if (gift === 1) AFSm = Lev - Dispm;
  if (AFSm < 0) AFSm = 0;

  if (context.dela === 1) {
    result = result + AFS + sbtpm + AFSm;
  } else if (context.dela === 2) {
    result = result + AFS;
  } else {
    result = (result + AFS + sbtpm + AFSm) / (gift + 1);
  }

  return result;
}

/**
 * The two housing supplements together, rounded to a whole krona a month.
 *
 * NOTE: the VBA guards on `pblnCloseOrSave` and traps errors into `#VALUE!`.
 * Both are worksheet-function housekeeping with no bearing on the result, and
 * are dropped.
 *
 * @param marginal 0 applies the rounding, 1 removes it.
 * @param year     Income year.
 */
export function btp_sbtp(BTP: number, SBTP: number, marginal = 0, year = 2016): number {
  // The division between spouses happens in the main model, not here.
  let result = BTP + SBTP;
  if (marginal === 0) result = vbaInt(result / 12 + 0.5) * 12;

  // Before 2014, a monthly amount under 25 kr was not paid out -- but only when
  // there is no särskilt bostadstillägg alongside it.
  if (marginal === 0 && year < 2014 && SBTP === 0 && result < 25 * 12) result = 0;

  return result;
}
