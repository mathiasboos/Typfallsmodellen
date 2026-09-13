/**
 * The drawdown half of Mcalc's age loop: pensions paid, and closing balances.
 *
 * Port of VBA_go.bas lines 1290-1744. Three sections in the original's order:
 * the public pensions, the occupational pension and private saving, and the
 * balances each of them carries into the next year.
 *
 * This is the most intricate stretch of the whole model. The retirement year is
 * computed by a separate branch from every later year, partial withdrawal
 * (`def_ar` above `PAR`) splits the ATP in two, and a scratch variable carries
 * a value between branches of the same `If`. The structure is kept as it is,
 * with the sections split where the VBA's own comments split them.
 */

import { withdrawalShare } from "../income/wages.js";
import { incomePensionYear, ppkassa } from "../pension/incomePension.js";
import { fnDeltalIp } from "../pension/deltal.js";
import { ppmavg } from "../pension/contributions.js";
import { tp, tpFaktor } from "../pension/atp.js";
import { gp, tillagg } from "../bidrag/garantipension.js";
import { PrivatSpar } from "../saving/privateSaving.js";
import type { SavingTypeId } from "../saving/privateSaving.js";
import { avkskatt } from "../skatt/reduktioner.js";
import { FTJP, tjpkassa } from "../tjanstepension/index.js";
import type { SchemeContext, SchemeId } from "../tjanstepension/types.js";
import { vbaInt, vbaRound, wsMax, wsMin } from "../vba/math.js";
import type { Run } from "./mcalc.js";
import { runVectors } from "./state.js";

/** The follow-up indexation the ATP is divided by, which differs in 2000. */
function fnormDivisor(year: number): number {
  if (year < 2000) return 1;
  if (year === 2000) return 0.996;
  return 1.016;
}

/** `IP_`'s arguments, which six different call sites repeat verbatim. */
function ipArgs(run: Run, age: number, par: number, ratt: number, opening: number, andelUttag: number, previousPension: number) {
  const { v, p, s, context } = run;
  const year = v.year.get(age);
  return {
    year,
    par,
    born: p.born,
    ratt,
    arv1: v.ipArv1.get(age),
    arv2: v.ipArv2.get(age),
    kost: v.ipAvg.get(age),
    opening,
    andelUttag,
    previousPension,
    dtal: s.dtalIp,
    index: v.pindex.get(age) / v.pindex.getOrZero(age - 1),
    defAr: p.defAr,
    marginal: context.marginal,
    bindex: v.pindex.get(age) / v.iindex.get(age),
  };
}

function callIp(a: ReturnType<typeof ipArgs>, kind: number): number {
  const result = incomePensionYear(
    a.year, a.par, a.born, a.ratt, a.arv1, a.arv2, a.kost, a.opening, a.andelUttag,
    a.previousPension, a.dtal, a.index, a.defAr, a.marginal, a.bindex,
  );
  switch (kind) {
    case 4: return result.balance;
    default: return result.pension;
  }
}

/**
 * The public pensions paid this year: income pension, ATP, garantipension,
 * premium pension and inkomstpensionstillägg.
 */
export function payPublicPensions(run: Run, age: number, utgyear: number): void {
  const { v, s, p, context, deltalTables } = run;
  const year = v.year.get(age);
  const par = p.par;
  const born = p.born;
  const marginal = context.marginal;

  if (age < vbaInt(par)) {
    s.ip.set(age, 0);
    s.tp.set(age, 0);
    s.pp.set(age, 0);
    s.garp.set(age, 0);
    s.ptillagg.set(age, 0);
    s.mpension = 0;
    return;
  }

  if (age === vbaInt(par)) {
    payAtRetirement(run, age, utgyear);
    return;
  }

  // ---- Every year after the retirement year ------------------------------
  let ppmonth = run.pmonth;
  if (vbaInt(par) + 1 < age) ppmonth = 12;

  if (age <= p.defAr) s.dtalIp = fnDeltalIp(born, age, deltalTables);

  s.ip.set(
    age,
    callIp(ipArgs(run, age, par, s.ipRatt.get(age), s.ipPbh.getOrZero(age - 1), s.uttagIp, s.ip.getOrZero(age - 1)), 0),
  );
  s.tp.set(age, 0);

  // `diverse`, the VBA's scratch variable, carries between the branches below.
  let diverse = 0;

  if (p.andelnya < 1) {
    const indexRatio = v.pindex.get(age) / v.pindex.getOrZero(age - 1);
    const divisor = fnormDivisor(year);
    s.tp.set(age, (s.tp.getOrZero(age - 1) * indexRatio * (12 / ppmonth)) / divisor);
    // NOTE: before 2000 the garantibelopp is written back into the *previous*
    // age; from 2000 it is written to this one. Kept as written.
    if (s.gbelopp.getOrZero(age - 1) > 0) {
      const uprated = (s.gbelopp.getOrZero(age - 1) * indexRatio * (12 / ppmonth)) / divisor;
      if (year < 2000) s.gbelopp.set(age - 1, uprated);
      else s.gbelopp.set(age, uprated);
    }

    if (s.uttagIp < 1) {
      // Partial withdrawal, so def_ar is above par and there is ATP left.
      if (p.defAr === age) {
        // The remaining share, and the months it is drawn over.
        ppmonth = 12 - vbaInt(12 * (born + p.defAr - vbaInt(born + p.defAr)));
        diverse = tp(s.atpPoints, s.atpYear, p.civ, p.defAr, born, age, v.pbb.get(age), 1 - s.uttagIp);
        diverse = (diverse * (1 - p.andelnya) * ppmonth) / 12;
        if (marginal === 0) {
          s.tp.set(age, vbaInt(s.tp.get(age) / 12 + 0.5) * 12);
          diverse = vbaInt(diverse / ppmonth + 0.5) * run.pmonth;
        }
        s.tp.set(age, s.tp.get(age) + diverse);
      } else if (p.defAr > par && p.defAr === age + 1) {
        s.tp.set(age, s.tp.getOrZero(age - 1) - diverse);
        if (year < 2000) {
          diverse = (diverse * indexRatio * 12) / ppmonth;
          s.tp.set(age, s.tp.getOrZero(age - 1) * indexRatio);
        } else {
          diverse = (diverse * indexRatio * (12 / ppmonth)) / divisor;
          s.tp.set(age, (s.tp.getOrZero(age - 1) * indexRatio) / divisor);
        }
        s.tp.set(age, s.tp.get(age) + diverse);
        if (marginal === 0) s.tp.set(age, vbaInt(s.tp.get(age) / 12 + 0.5) * 12);
      } else {
        ppmonth = 12;
        s.tp.set(age, (s.tp.getOrZero(age - 1) * indexRatio) / divisor);
        if (marginal === 0) s.tp.set(age, vbaInt(s.tp.get(age) / 12 + 0.5) * 12);
      }
    }
  }

  // The garantibelopp on rights earned up to 1994 tops the ATP up.
  s.gbelopp.set(age, tp(s.atp94, s.atp94Year, p.civ, par, born, age, v.pbb.get(age), s.uttagIp));
  if (s.gbelopp.get(age) - (s.tp.get(age) + (s.ip.get(age) * 185) / 160) > 0) {
    s.tp.set(age, s.tp.get(age) + s.gbelopp.get(age) - (s.tp.get(age) + (s.ip.get(age) * 185) / 160));
  }

  // Garantipension, on its own divisor.
  s.dtalIp = garantipensionDivisor(run, age);
  const gpUnd = callIp(
    ipArgs(run, age, wsMax(par, p.riktalder), s.gpRatt.get(age), s.gpPbh.getOrZero(age - 1), 1, s.gpUnd),
    0,
  );
  s.gpUnd = gpUnd;

  if (age >= p.riktalder) {
    s.garp.set(age, garantipension(run, age, gpUnd + s.tp.get(age) * wsMax(1, tpFaktor(par)), utgyear));
  } else {
    s.garp.set(age, 0);
  }

  s.mpension = (gpUnd + s.tp.get(age) * wsMax(1, tpFaktor(par))) / run.pmonth;
  if (s.mpension === 0 && s.ip.get(age) > 0) {
    // A simplification, as the VBA calls it.
    s.mpension =
      (s.ip.get(age) * (185 / 160) * (1 / s.uttagIp)) / ppmonth +
      (s.tp.get(age) * wsMax(1, tpFaktor(par)) * (1 / s.uttagIp)) / run.pmonth;
  }

  payPremiumPension(run, age, ppmonth);
  payIncomePensionSupplement(run, age, utgyear);
}

/**
 * The ATP earned over a working life, settled in the retirement year.
 *
 * Two averages of the fifteen years that count, taken differently:
 *
 * - the garantibelopp uses points earned to 1994, and the VBA sums the last
 *   fifteen *ages* of that window;
 * - the ATP itself sorts the points ascending first, so its fifteen are the
 *   fifteen *best*.
 *
 * NOTE: the comment above the first loop in the original says "Sorteras i
 * stigande ordning" and no sort happens there. Only the second is sorted. Kept.
 *
 * NOTE: the sort is in place on `TP_points`, and `FTJP` reads `STP_points`
 * afterwards for SAF-LO -- which is why the earning phase keeps that copy in
 * age order. The order of the two is load-bearing.
 */
function accrueAtp(run: Run, age: number): void {
  const { v, s, p, context } = run;
  const year = v.year.get(age);
  const marginal = context.marginal;

  if (p.andelnya >= 1) return;

  for (let counter = v.startage; counter <= s.tp94p.lastAge; counter += 1) {
    s.tp94p.set(counter, v.year.get(counter) <= 1994 ? s.tpPoints.getOrZero(counter) : 0);
  }

  for (let counter = v.startage; counter <= s.tp94p.lastAge; counter += 1) {
    if (s.tp94p.get(counter) > 0) s.atp94Year += 1;
    if (counter > s.tp94p.lastAge - 15) s.atp94 += s.tp94p.get(counter) / 15;
  }
  if (marginal === 0) s.atp94 = vbaRound(s.atp94, 2);

  s.gbelopp.set(age, tp(s.atp94, s.atp94Year, p.civ, p.par, p.born, age, v.pbb.get(age), s.uttagIp));
  s.gbelopp.set(age, (s.gbelopp.get(age) * run.pmonth) / 12);

  // Ascending, in place -- see the note above.
  const sorted = s.tpPoints.slice().sort((a, b) => a - b);
  for (let i = 0; i < sorted.length; i += 1) s.tpPoints.set(v.startage + i, sorted[i]!);

  for (let counter = v.startage; counter <= v.slutage; counter += 1) {
    if (s.tpPoints.get(counter) > 0) s.atpYear += 1;
    if (counter > v.slutage - 15) s.atpPoints += s.tpPoints.get(counter) / 15;
  }
  if (marginal === 0) s.atpPoints = vbaRound(s.atpPoints, 2);

  s.tp.set(age, tp(s.atpPoints, s.atpYear, p.civ, p.par, p.born, age, v.pbb.get(age), s.uttagIp));
  s.tp.set(age, (s.tp.get(age) * (1 - p.andelnya) * run.pmonth) / 12);

  // The garantibelopp tops the ATP up where it falls short.
  if (s.gbelopp.get(age) - (s.tp.get(age) + (s.ip.get(age) * 185) / 160) > 0) {
    s.tp.set(age, s.tp.get(age) + s.gbelopp.get(age) - (s.tp.get(age) + (s.ip.get(age) * 185) / 160));
  }
  if (marginal === 0) s.tp.set(age, vbaInt(s.tp.get(age) / run.pmonth + 0.5) * run.pmonth);
  void year;
}

/** The divisor garantipension is measured against. */
function garantipensionDivisor(run: Run, age: number): number {
  const { p, context, deltalTables } = run;
  if (!context.ownDeltal && vbaInt(p.born) <= 1958) {
    return deltalTables.incomePension(vbaInt(p.born), age);
  }
  // QUIRK: the workbook reads `mortality!L5` (VBA_go.bas:1356 and :1496), a
  // worksheet formula holding this cohort's income pension delningstal at the
  // riktalder -- 16.67 for 1959, beside the cells holding 1959 and 66. It does
  // not vary with `age`, so the guarantee underlag is the pension as it would
  // be at the riktalder rather than the larger one a later withdrawal buys.
  // That is what a guarantee underlag should be, and it is also what feeds the
  // inkomstpensionstillagg.
  //
  // The port had `age` here, which is the same figure whenever retirement is
  // at the riktalder and only then. It went unnoticed until the full golden
  // set, whose block B draws the 1959 cohort at 67, 68, 70 and 75: nine cases
  // where the supplement was wrong by up to 525 kr a month.
  return deltalTables.incomePension(vbaInt(p.born), p.riktalder);
}

/** `gp` with the arguments Mcalc always passes. */
function garantipension(run: Run, age: number, underlag: number, utgyear: number): number {
  const { v, p, s, context } = run;
  // Imported lazily to keep the benefit module's imports in one place.
  return gp(
    underlag,
    p.civ,
    vbaInt(p.born),
    v.pbb.get(age),
    run.forstid,
    context.marginal,
    age,
    utgyear,
    run.iyear,
    v.ibb.get(age),
    s.kvoten,
    p.riktalder,
    s.uttagIp,
  );
}

/** Premium pension, whose return accrues evenly across the year. */
function payPremiumPension(run: Run, age: number, ppmonth: number): void {
  const { v, s, p, context, deltalTables } = run;
  const year = v.year.get(age);
  const opening = s.ppPbh.getOrZero(age - 1);

  if (!context.simplifiedPremiumPension) {
    s.pp.set(age, ppkassa(p.par, p.born, age, opening, deltalTables, { incomePension: s.uttagIp, premiumPension: s.uttagPp }, p.defAr, 1, context.marginal));
    return;
  }

  if (age > p.defAr || (age > p.par && age < p.defAr) || s.uttagPp === 1) {
    // The forskottsranta built into the divisor, unwound year by year.
    const rate =
      year > 2017 ? 1.0165 : year > 2014 ? 1.029 : year > 2007 ? 1.039 : year > 2002 ? 1.027 : 1.036;
    s.pp.set(age, (s.pp.getOrZero(age - 1) * v.yieldFactor.get(age)) / rate);
    return;
  }

  if (year > p.born + p.par + 1 && s.uttagPp === 1 && s.ppRatt.getOrZero(age - 1) === 0) {
    s.pp.set(age, (s.pp.getOrZero(age - 1) * v.yieldFactor.get(age)) / 1.035);
  } else {
    s.pp.set(age, ppkassa(p.par, p.born, age, opening, deltalTables, { incomePension: s.uttagIp, premiumPension: s.uttagPp }, p.defAr, 1, context.marginal));
  }
  void ppmonth;
}

/** Inkomstpensionstillägg, from the riktålder and only once retired. */
function payIncomePensionSupplement(run: Run, age: number, utgyear: number): void {
  const { v, s, p, context } = run;
  const year = v.year.get(age);

  if (age < p.riktalder || year < vbaInt(p.par) + p.born) {
    s.ptillagg.set(age, 0);
    return;
  }

  // NOTE: the fallback index differs between the two call sites -- 186.52 in
  // the retirement year, 182.58 afterwards. Kept as written.
  const fallback = age === vbaInt(p.par) ? 186.52 : 182.58;
  const index2021 =
    2021 - vbaInt(p.born) > v.startage ? v.iindex.getOrZero(2021 - vbaInt(p.born)) : fallback;

  s.ptillagg.set(
    age,
    tillagg(12 * s.mpension, utgyear, p.born, v.iindex.get(age), index2021, s.uttagIp, s.pgiYears, context.marginal),
  );

  // Introduced on 1 September 2021, so that year pays at most four months.
  if (year === 2021) {
    s.ptillagg.set(age, (s.ptillagg.get(age) * wsMin(4, run.pmonth)) / 12);
  } else {
    s.ptillagg.set(age, (s.ptillagg.get(age) * run.pmonth) / 12);
  }
}

/** The retirement year, which the VBA computes by its own branch. */
function payAtRetirement(run: Run, age: number, utgyear: number): void {
  const { v, s, p, context, deltalTables } = run;
  const year = v.year.get(age);
  const par = p.par;
  const born = p.born;

  // Income pension, on the balance built up to now.
  s.ip.set(
    age,
    callIp(ipArgs(run, age, par, s.ipRatt.get(age), s.ipPbh.getOrZero(age - 1), s.uttagIp, s.ip.getOrZero(age - 1)), 0),
  );

  accrueAtp(run, age);

  s.gpundtab1 = 0;

  if (age < p.riktalder || year < vbaInt(par) + born) {
    s.garp.set(age, 0);
  } else {
    s.dtalIp = garantipensionDivisor(run, age);

    // NOTE: `If Int(PAR + 1) = age` can never hold inside a branch reached only
    // when `age = Int(PAR)`. Dead in the original; kept.
    if (vbaInt(par + 1) === age) s.garp.set(age, (s.garp.get(age) * 12) / run.pmonth);

    const gpUnd = callIp(
      ipArgs(run, age, wsMax(par, p.riktalder), s.gpRatt.get(age), s.gpPbh.getOrZero(age - 1), 1, s.garp.getOrZero(age - 1)),
      0,
    );
    s.gpUnd = gpUnd;

    s.garp.set(
      age,
      garantipension(run, age, (gpUnd * 12) / run.pmonth + s.tp.get(age) * wsMax(1, tpFaktor(par)), utgyear),
    );
    if (vbaInt(par) === age) s.garp.set(age, ((s.garp.get(age) * run.pmonth) / 12) * s.uttagIp);

    s.mpension = (gpUnd + s.tp.get(age) * wsMax(1, tpFaktor(par))) / run.pmonth;
    if (age === vbaInt(par)) s.gpundtab1 = s.mpension;
  }

  // Premium pension: the return accrues evenly over the months before it starts.
  s.pp.set(
    age,
    ppkassa(
      par,
      born,
      age,
      s.ppPbh.getOrZero(age - 1) * v.yieldFactor.get(age) ** ((12 - run.pmonth) / 12),
      deltalTables,
      { incomePension: s.uttagIp, premiumPension: s.uttagPp },
      p.defAr,
      1,
      context.marginal,
    ),
  );

  payIncomePensionSupplement(run, age, utgyear);
}

/**
 * The occupational pension and private saving, paid from `tjp_par`.
 */
export function payOccupationalAndSaving(run: Run, age: number): void {
  const { v, s, p, context, deltalTables } = run;
  const year = v.year.get(age);
  const tjpPar = p.tjpPar;
  const born = p.born;

  if (age < vbaInt(tjpPar)) {
    s.tjp.set(age, 0);
    s.ips.set(age, 0);
    s.pps.set(age, 0);
    return;
  }

  const schemeContext: SchemeContext = {
    year,
    born,
    wStart: p.wStart,
    tjpPar,
    flexPension: context.flexPension,
    marginal: context.marginal,
  };

  if (year === vbaInt(born + tjpPar)) {
    s.tjp.set(
      age,
      tjpkassa(
        tjpPar,
        born,
        age,
        (s.tjpRatt.get(age) + s.tjpPbh.getOrZero(age - 1)) * v.yieldFactor.get(age) ** ((12 - run.tmonth) / 12),
        p.avtal as SchemeId,
        year,
        deltalTables,
        context,
        run.tmonth,
        0,
      ),
    );

    // Private saving is annuitised over the expected remaining life, or over
    // the temporary withdrawal when one is set.
    const uttag =
      context.tempIpsUttag === 0
        ? deltalTables.expectedLife(vbaInt(born), vbaInt(p.par))
        : context.tempIpsUttag;
    s.ips.set(age, (s.ipsPbh.getOrZero(age - 1) * (1 + run.input.realReturn) ** (uttag / 2)) / uttag);
    s.pps.set(age, (s.ppsPbh.getOrZero(age - 1) * (1 + run.input.realReturn) ** (uttag / 2)) / uttag);

    // The defined-benefit part on top.
    s.tjp.set(
      age,
      FTJP(
        s.tjp.get(age),
        v.startage,
        age,
        tjpPar,
        born,
        p.avtal as SchemeId,
        (a) => s.stpPoints.getOrZero(a),
        (a) => s.tp.getOrZero(a),
        runVectors(v),
        deltalTables,
        schemeContext,
        context,
        run.tmonth,
      ),
    );
    return;
  }

  // Later years follow prisbasbelopp, with the first full year scaled up.
  const priceRatio = v.pbb.get(age) / v.pbb.getOrZero(age - 1);
  const scale = year === vbaInt(tjpPar + born + 1) ? (12 / run.tmonth) : 1;
  s.tjp.set(age, s.tjp.getOrZero(age - 1) * priceRatio * scale);
  s.ips.set(age, s.ips.getOrZero(age - 1) * priceRatio * scale);
  s.pps.set(age, s.pps.getOrZero(age - 1) * priceRatio * scale);

  if (context.tempTjpUttag > 0) {
    if (age === tjpPar + context.tempTjpUttag) {
      s.tjp.set(age, (s.tjp.get(age) * (12 - run.tmonth)) / 12);
    } else if (age > tjpPar + context.tempTjpUttag) {
      s.tjp.set(age, 0);
    }
  }
  if (context.tempIpsUttag > 0) {
    if (age === tjpPar + context.tempIpsUttag) {
      s.ips.set(age, (s.ips.get(age) * (12 - run.tmonth)) / 12);
      s.pps.set(age, (s.pps.get(age) * (12 - run.tmonth)) / 12);
    } else if (age > tjpPar + context.tempIpsUttag) {
      s.ips.set(age, 0);
      s.pps.set(age, 0);
    }
  }

  // A payout larger than the balance is capped at it.
  if (s.ppsPbh.getOrZero(age - 1) < s.pps.get(age)) {
    s.pps.set(age, s.ppsPbh.getOrZero(age - 1) * v.yieldFactor.get(age));
  }
}

/** Rounds every pension paid to a whole krona a month. */
export function roundPensions(run: Run, age: number): void {
  const { s, context } = run;
  if (context.marginal !== 0) return;
  for (const vector of [s.ip, s.pp, s.tjp, s.ips, s.pps]) {
    vector.set(age, vbaInt(vector.get(age) / 12 + 0.5) * 12);
  }
  // ptillagg is already rounded inside `tillagg`.
}

/** The closing balance of each pot. */
export function closeBalances(run: Run, age: number, utgyear: number): void {
  const { v, s, p, context } = run;
  const year = v.year.get(age);
  const marginal = context.marginal;

  if (age === v.startage) {
    for (const vector of [s.ipPbh, s.gpPbh, s.ppPbh, s.tjpPbh, s.ipsPbh, s.ppsPbh]) {
      vector.set(age, 0);
    }
  } else if (year > 1960) {
    if (age < v.slutage) {
      s.ipPbh.set(
        age,
        callIp(ipArgs(run, age, p.par, s.ipRatt.get(age), s.ipPbh.getOrZero(age - 1), s.uttagIp, s.ip.getOrZero(age - 1)), 4),
      );
      s.gpPbh.set(
        age,
        callIp(ipArgs(run, age, wsMax(p.par, p.riktalder), s.gpRatt.get(age), s.gpPbh.getOrZero(age - 1), 1, s.ip.getOrZero(age - 1)), 4),
      );
    }
    // The premium pension's deposit lands at the end of the year.
    const opening = s.ppPbh.getOrZero(age - 1);
    const yieldFactor = v.yieldFactor.get(age);
    const inheritance =
      v.year.getOrZero(age - 1) < 2010
        ? opening * (v.ppArv.get(age) - 1)
        : (opening * (v.ppArv.get(age) - 1)) / yieldFactor ** (6 / 12);
    s.ppPbh.set(
      age,
      opening * yieldFactor +
        inheritance +
        v.rgk.get(age) * s.ppRatt.get(age) +
        ppmavg(utgyear, opening, v.ppAvg.get(age), yieldFactor, context),
    );
  }

  // The occupational and private pots, adjusted for what was drawn.
  if (age === v.startage) {
    const half = v.yieldFactor.get(age) ** 0.5;
    s.tjpPbh.set(age, s.tjpRatt.get(age) * half);
    s.ipsPbh.set(age, s.ipsRatt * half);
    s.ppsPbh.set(age, s.ipsRatt * half);
  } else {
    const half = v.yieldFactor.get(age) ** 0.5;
    // Interest on the amount paid out during the year.
    s.ppPbh.set(age, s.ppPbh.get(age) - s.pp.get(age) * half);

    s.tjpPbh.set(
      age,
      s.tjpRatt.get(age) * half +
        s.tjpPbh.getOrZero(age - 1) * v.yieldFactor.get(age) * v.tpAvg.get(age) +
        s.tjpPbh.getOrZero(age - 1) * (v.ppArv.getOrZero(age - 1) - 1) * context.occupationalInheritanceGains,
    );
    let kskatt = s.tjpPbh.get(age) * avkskatt(year, context, 0);
    s.tjpPbh.set(age, s.tjpPbh.get(age) - kskatt - s.tjp.get(age) * half);

    const spar = context.privateSavingKind;
    if (spar === 1 || spar === 2) {
      s.ppsPbh.set(
        age,
        PrivatSpar(s.ppsPbh.getOrZero(age - 1), s.ipsRatt, v.yieldFactor.get(age), year, spar as SavingTypeId) -
          s.pps.get(age) * half,
      );
    } else {
      kskatt = 0;
      s.ipsPbh.set(
        age,
        s.ipsRatt * half + s.ipsPbh.getOrZero(age - 1) * v.yieldFactor.get(age) * v.tpAvg.get(age),
      );
      kskatt = s.ipsPbh.get(age) * avkskatt(year, context, 0);
      s.ipsPbh.set(age, s.ipsPbh.get(age) - kskatt - s.ips.get(age) * half);
    }
    // NOTE: rounded after its last use, so this changes nothing. Kept.
    if (marginal === 0) kskatt = vbaInt(kskatt * 100) / 100;
  }

  for (const vector of [s.tjpPbh, s.ipsPbh, s.ppsPbh]) {
    if (vector.get(age) < 0) vector.set(age, 0);
  }

  // An opening balance typed in for one income year replaces what was built.
  if (context.pbhYear > 0 && year === context.pbhYear) {
    // NOTE: a division, where the commented-out alternative is a multiplication
    // by 185/160. Almost certainly not intended, but it is what the workbook
    // computes, and it only fires when an opening balance is entered.
    s.gpPbh.set(age, context.pbhIp / s.ipPbh.get(age));
    s.ipPbh.set(age, context.pbhIp);
    s.ppPbh.set(age, context.pbhPp);
    s.tjpPbh.set(age, context.pbhTjp);
    s.ipsPbh.set(age, context.pbhIps);
  }

  if (marginal === 0) {
    for (const vector of [s.ipPbh, s.gpPbh, s.ppPbh, s.tjpPbh, s.ipsPbh, s.ppsPbh]) {
      vector.set(age, vbaInt(vector.get(age)));
    }
  }
}

/** Gross income before tax, which the tax section starts from. */
export function grossIncome(run: Run, age: number): void {
  const { v, s, context } = run;
  let amount =
    v.income.get(age) +
    s.ip.get(age) +
    s.tp.get(age) +
    s.garp.get(age) +
    s.pp.get(age) +
    s.tjp.get(age) +
    s.ptillagg.get(age);
  // An IPS payout is taxed as income; a KF or ISK payout is not.
  if (context.privateSavingKind === 0) amount += s.ips.get(age);
  s.brutto.set(age, amount);
}

/** One age's drawdown phase, in the VBA's order. */
export function drawdownPhase(run: Run, age: number, utgyear: number): void {
  payPublicPensions(run, age, utgyear);
  payOccupationalAndSaving(run, age);
  roundPensions(run, age);
  closeBalances(run, age, utgyear);
  grossIncome(run, age);
}

export { withdrawalShare };
