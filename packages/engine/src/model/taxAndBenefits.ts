/**
 * Tax, benefits and the output row -- the last third of Mcalc's age loop.
 *
 * Port of VBA_go.bas 1745-2117: the deductions and the three taxes, the tax
 * reductions, the child and housing benefits, social assistance, and the
 * seventeen-column `mvalues` row that every table and both figures are drawn
 * from.
 */

import {
  BTP,
  SBTP,
  btp_sbtp,
  CalcAntalBarn,
  CalcBarnPerAlder,
  barnbidraget,
  bist25,
  bistOld,
  bobid,
  gp,
  ustod,
} from "../bidrag/index.js";
import { pgi } from "../pension/contributions.js";
import {
  DeductionType,
  FAared,
  PublicAvg,
  avdragRES,
  avdragxx,
  Jobbxx,
  Xage,
  pandred,
  sared,
  statlig,
} from "../skatt/index.js";
import { vbaInt, vbaRound, wsMax, wsMin } from "../vba/math.js";
import type { Run } from "./mcalc.js";
import { runVectors } from "./state.js";
import type { MvaluesRow } from "./state.js";

/**
 * The 1 320-krona reduction of 2000 and 2001, taken off the PGI ceiling.
 *
 * QUIRK: the original writes this block out twice in a row, so the reduction is
 * added *twice* in both years. Kept, and separated out here so the doubling can
 * be asserted rather than only described.
 *
 * The first copy carries an `If Skyear = 2004` inside a block reached only when
 * Skyear is 2000 or 2001, which is dead.
 */
export function millenniumReduction(
  rakassaIn: number,
  skyear: number,
  pgiAmount: number,
  cbefvi: number,
  marginal: 0 | 1,
): number {
  let rakassa = rakassaIn;
  for (let copy = 0; copy < 2; copy += 1) {
    if (skyear > 1999 && skyear < 2002) {
      const base = vbaInt(pgiAmount / 0.93);
      rakassa += base < 135_000 ? wsMin(1320, base) : 1320 - 0.012 * (base - 135_000);
      if (marginal === 0) rakassa = vbaInt(rakassa);
      if (copy === 0 && skyear === 2004 && cbefvi >= 100) rakassa += 200;
    }
  }
  return rakassa;
}

/** What the tax section works out, which the benefit section then reads. */
interface TaxResult {
  readonly ctxfvi: number;
  readonly grundavdrag: number;
  readonly cbefvi: number;
  readonly kapskatt: number;
}

/**
 * Deductions, the three taxes and the reductions, ending at net income.
 */
export function taxes(run: Run, age: number, utgyear: number, skyear: number): TaxResult {
  const { v, s, p, context } = run;
  const year = v.year.get(age);
  const marginal = context.marginal;
  const income = v.income.get(age);
  const wage = v.wage.get(age);

  // Deductions from income. Only private pension saving is ever taken; travel
  // and other deductions are commented out in the original.
  let kostnadsavd = 0;
  if (s.brutto.get(age) > 0 && context.privateSavingKind === 0) {
    const allowed =
      context.rulesFromSkatt === 0
        ? s.tjpRatt.get(age) === 0 || year < 2016
        : context.rulesFromSkatt < 2016;
    if (allowed) {
      kostnadsavd +=
        avdragRES(
          s.ipsRatt,
          age,
          runVectors(v),
          s.tjpRatt.get(age) > 0,
          DeductionType.PrivatePension,
          marginal,
          skyear,
        );
    }
  }

  let ctxfvi = s.brutto.get(age) - kostnadsavd;
  if (marginal === 0) ctxfvi = vbaInt(ctxfvi / 100) * 100;

  // The higher allowance for the old follows the riktålder from 2020.
  const gage = Xage(skyear, context);
  const grundavdrag = avdragxx(ctxfvi, v.pbb.get(age), marginal, age, year, 2100, v.ibb.get(age), s.kvoten, skyear, gage);

  // The individual's own pension contribution, on the expenditure rule year.
  const pensionavgift = pgi(utgyear, income, v.pbb.get(age), v.ibb.get(age), v.fpb.get(age), context, 1, age, 0);
  const cbefvi =
    ctxfvi -
    grundavdrag -
    pensionavgift +
    pgi(utgyear, income, v.pbb.get(age), v.ibb.get(age), v.fpb.get(age), context, 2, age, 0);

  let kinkskatt = cbefvi * v.komSkatt.get(age);
  if (marginal === 0) kinkskatt = vbaInt(kinkskatt);

  let kyrkskatt = cbefvi * v.begravavg.get(age);
  if (marginal === 0) kyrkskatt = vbaInt(kyrkskatt);

  // The public service fee rides along with the state tax.
  let statskatt =
    statlig(cbefvi, v.taxLimit1.get(age), v.taxLimit2.get(age), marginal) +
    PublicAvg(cbefvi, p.born, runVectors(v), 0.01, age, marginal, skyear);

  // Capital income, and the tax or reduction on it, are assumed to start at
  // retirement.
  let kapskatt = 0;
  if (age >= vbaInt(p.par) && context.kapital !== 0) {
    kapskatt =
      context.kapital < -100_000
        ? -30_000 + (context.kapital + 100_000) * 0.7 * 0.3
        : context.kapital * 0.3;
    if (kapskatt < 0) {
      if (statskatt > -kapskatt) {
        statskatt += kapskatt;
      } else {
        kinkskatt += statskatt + kapskatt;
        statskatt = 0;
      }
    } else {
      statskatt += kapskatt;
    }
  }

  // The pension contribution is given back as a reduction.
  const pensredukt = pgi(skyear, income, v.pbb.get(age), v.ibb.get(age), v.fpb.get(age), context, 2, age, 0);

  let jobbavdrag = Jobbxx(
    wage, age, v.komSkatt.get(age), v.pbb.get(age), marginal, ctxfvi - wage, skyear, 2199, v.ibb.get(age), s.kvoten, skyear, gage,
  );
  if (kinkskatt - pensredukt < jobbavdrag) jobbavdrag = kinkskatt - pensredukt;

  // Union and unemployment-insurance fees, and the reductions on them.
  let akassa = context.akasseavg;
  let fack = context.fack;
  let rakassa = 0;
  let rfack = 0;

  if (age >= p.wStart && age <= p.par) {
    if (2019 - vbaInt(p.born) > v.startage) {
      // NOTE: `KPI(2019 - born)` indexes with the unrounded birth year, where
      // every neighbouring expression uses `Int(born)`. Kept as written.
      akassa = (akassa * v.kpi.get(age)) / v.kpi.getOrZero(2019 - p.born);
      fack = (fack * v.kpi.get(age)) / v.kpi.getOrZero(2019 - p.born);
    }
    if (age === p.par) {
      akassa = akassa * (12 - run.pmonth);
      fack = fack * (12 - run.pmonth);
    }
  } else {
    akassa = 0;
    fack = 0;
  }

  if (skyear > 1995 && skyear <= 2006) {
    rakassa = 0.75 * akassa;
    if (fack >= 400) rfack = 0.25 * fack;
  }
  if (skyear === 2018) {
    rakassa = 0;
    if (fack >= 400) rfack = (0.25 * fack * 6) / 12;
  }
  if (skyear === 2019) {
    rakassa = 0;
    if (fack >= 400) rfack = (0.25 * fack * 3) / 12;
  }
  if (skyear >= 2022) {
    rakassa = 0.25 * akassa;
    if (skyear === 2022) rakassa = rakassa / 2;
  }

  rakassa = millenniumReduction(rakassa, skyear, s.pgi.get(age), cbefvi, marginal);

  if (marginal === 0) {
    rakassa = vbaInt(rakassa);
    rfack = vbaInt(rfack);
  }

  // The reduction for sickness and activity compensation.
  let saavdrag = 0;
  if (context.satagare > 0 && context.satagare <= year) {
    rakassa = 0;
    rfack = 0;
    saavdrag = sared(income, skyear, marginal, v.komSkatt.get(age), v.pbb.get(age), 2100, v.ibb.get(age), s.kvoten, year);
  }

  // Computed and then never used, as in the original.
  void pandred(wage, skyear, marginal);

  if (kinkskatt - saavdrag - jobbavdrag < saavdrag) saavdrag = kinkskatt - saavdrag - jobbavdrag;

  let faavdrag = FAared(cbefvi, skyear, marginal, v.pbb.get(age));
  if (kinkskatt - saavdrag - faavdrag - jobbavdrag < faavdrag) {
    faavdrag = kinkskatt - saavdrag - faavdrag - jobbavdrag;
  }

  const tax = wsMax(
    kinkskatt + kyrkskatt + statskatt + pensionavgift - pensredukt - jobbavdrag - rakassa - rfack - saavdrag - faavdrag,
    0,
  );
  s.netto.set(age, s.brutto.get(age) - tax);

  // Still in scope when the loop ends, and read by the recomputation at the
  // retirement age. See `RunState.leftovers`.
  s.leftovers.gage = gage;
  s.leftovers.rakassa = rakassa;
  s.leftovers.saavdrag = saavdrag;

  return { ctxfvi, grundavdrag, cbefvi, kapskatt };
}

/**
 * Child allowance, housing allowance, housing supplement and social assistance.
 */
export function benefits(run: Run, age: number, utgyear: number, tax: TaxResult): void {
  const { v, s, p, context } = run;
  const year = v.year.get(age);
  const marginal = context.marginal;
  const civ = p.civ;
  const [barn1, barn2, barn3, barn4] = context.childBirthYears;

  let bidragovr = 0;

  const antalBarn = CalcAntalBarn(vbaInt(p.born) + age, barn1, barn2, barn3, barn4);
  const barnbidrag = barnbidraget(antalBarn, utgyear) + ustod(antalBarn, civ + 1, utgyear);

  // Rent is assumed to follow CPI from the reference year.
  let hyraT = 0;
  if (v.startage < context.referenceYear - p.born) {
    if (year >= 1960) hyraT = context.hyra * (v.kpi.get(age) / v.kpi.getOrZero(context.referenceYear - p.born));
  } else {
    hyraT = context.hyra * (v.kpi.get(age) / v.kpi.get(v.startage));
  }

  // Under 29 the young-adult rules apply; the model ignores part years.
  let bostadsbidrag = bobid(
    civ + 1,
    s.brutto.get(age) - s.ptillagg.get(age),
    0,
    antalBarn,
    hyraT,
    80,
    marginal,
    age < 29 ? 1 : 0,
    utgyear,
  );

  let bostadstillagg = 0;
  let sbostadstillagg = 0;

  // NOTE: `ftid` reads the *setting*, where SBTP below is passed the corrected
  // `forstid`. The two differ whenever Mcalc had to raise it. Kept.
  const ftid = context.insuranceYears;
  const ansoker = context.ansokt;

  if (p.defAr > p.par && p.defAr <= age) {
    s.uttagIp = 1;
    s.uttagPp = 1;
  }

  const manualSa = run.pgbManual.get(wsMax(age, 15))?.sa ?? 0;
  if ((age >= vbaInt(p.par) && age >= p.riktalder && s.uttagIp > 0) || manualSa > 0) {
    const iyearAge = run.iyear - vbaInt(p.born);
    if (run.iyear > 0 && run.iyear <= year && iyearAge >= v.startage && iyearAge <= v.slutage) {
      s.maxhyra = v.ibb.get(age) / v.ibb.get(iyearAge);
      s.kvoten = v.pbb.get(iyearAge) / v.ibb.get(iyearAge);
    }

    const ap = age >= vbaInt(p.par) ? 1 : 0;
    let apm = 0;
    let tjpm = 0;
    let garpm = 0;
    let makaInk = 0;

    if (civ !== 0) {
      // The spouse is assumed the same age, with income in proportion.
      apm = ap;
      const share = context.makensInkomst / (12 * run.input.monthlySalary);
      if (context.makensInkomst > 0) {
        makaInk = (s.brutto.get(age) - tax.kapskatt - s.ptillagg.get(age)) * share;
        garpm = gp(
          s.gpUnd * share, civ, vbaInt(p.born), v.pbb.get(age), context.makeInsuranceYears, marginal,
          age, utgyear, run.iyear, v.ibb.get(age), s.kvoten, p.riktalder, s.uttagIp,
        );
        tjpm = s.tjp.get(age) * share;
      } else {
        garpm = gp(
          0, civ, vbaInt(p.born), v.pbb.get(age), context.makeInsuranceYears, marginal,
          age, utgyear, run.iyear, v.ibb.get(age), s.kvoten, p.riktalder, s.uttagIp,
        );
      }
    }

    const iyearIbb = v.ibb.getOrZero(run.iyear - vbaInt(p.born));
    // In the final withdrawal year the amounts are put on a full-year footing.
    const partYear = vbaInt(p.defAr) === age;
    const scale = partYear ? 12 / run.pmonth : 1;
    const base = s.brutto.get(age) - (partYear ? v.wage.get(age) : 0) - tax.kapskatt - s.ptillagg.get(age);

    bostadstillagg =
      s.uttagIp *
      BTP(
        base * scale, makaInk * scale, 12 * hyraT, civ, v.pbb.get(age), context,
        ap, apm, context.formogenhet, partYear ? 0 : v.wage.get(age), 0, marginal, s.maxhyra,
        utgyear, run.iyear, v.ibb.get(age), s.kvoten, age,
        s.tjp.get(age) * scale, tjpm * scale, s.garp.get(age) * scale, s.garp.get(age) * scale,
        iyearIbb, s.mpension, run.pmonth, ftid, v.iindex.get(age), s.uttagIp, ansoker,
      );

    sbostadstillagg = SBTP(
      base * scale, 12 * hyraT, civ, bostadstillagg + bostadsbidrag, tax.grundavdrag,
      {
        dela: context.dela,
        rulesFromUtg: context.rulesFromUtg,
        born: p.born,
        age,
        slutage: v.slutage,
        Iyear: run.iyear,
        vectors: runVectors(v),
      },
      v.komSkatt.get(age), ap, context.formogenhet, v.pbb.get(age), s.maxhyra,
      utgyear, run.iyear, v.ibb.get(age), s.kvoten, age, context.kapital, makaInk, marginal, -99, run.forstid,
    );

    if (context.ansokt === 9) {
      bidragovr = bostadsbidrag + bostadstillagg;
      bostadstillagg = 0;
    }
    if (ansoker === 0) {
      bostadstillagg = 0;
      sbostadstillagg = 0;
      // Only inside this block, so a household below the retirement age keeps
      // its housing allowance even when the run says nothing is applied for.
      bostadsbidrag = 0;
    }
    bostadstillagg = (btp_sbtp(bostadstillagg, sbostadstillagg, marginal, utgyear) * run.pmonth) / 12;

    // NOTE: these four are set only inside this block, so what survives the
    // loop is the last age that *reached* it, not the last age of the loop.
    s.leftovers.ap = ap;
    s.leftovers.apm = apm;
    s.leftovers.makaInk = makaInk;
    s.leftovers.tjpm = tjpm;
  }

  s.bidrag.set(age, barnbidrag + bostadsbidrag + bostadstillagg);
  if (s.bidrag.get(age) < 0) s.bidrag.set(age, 0);
  s.indDisp.set(age, s.netto.get(age) + s.bidrag.get(age) + bidragovr + s.pps.get(age));

  // As above: the recomputation at the retirement age does not work these out
  // again, it reads whatever the loop left behind.
  s.leftovers.barnbidrag = barnbidrag;
  s.leftovers.bostadsbidrag = bostadsbidrag;
  s.leftovers.hyraT = hyraT;
  s.leftovers.bidragovr = bidragovr;

  // Social assistance, on Socialstyrelsen's norm plus the rent.
  const bands = CalcBarnPerAlder(vbaInt(p.born) + age, barn1, barn2, barn3, barn4);
  const riksnorm = {
    marginal,
    born: p.born,
    age,
    vectors: runVectors(v),
  };
  let bist = 0;
  if (year >= 1985) {
    const args = [
      civ + 1, hyraT, s.indDisp.get(age),
      bands.b1, bands.b2, bands.b3, bands.b4, bands.b5, bands.b6, bands.b7, bands.b8,
      v.wage.get(age) * 0,
    ] as const;
    bist = year < 2006 ? bistOld(...args, riksnorm, year) : bist25(...args, riksnorm, year);
    // NOTE: with a January birthday `pmonth` is 12, so a childless household's
    // social assistance is multiplied by zero. The comment says the old-age
    // basic protection is assumed to take over.
    if (antalBarn === 0) bist = (bist * (12 - run.pmonth)) / 12;
    if (civ === 1 && bist > 0) bist = bist / 2;
  }
  if (bist > 0) {
    s.bidrag.set(age, s.bidrag.get(age) + bist);
    s.indDisp.set(age, s.indDisp.get(age) + bist);
  }
}

/** The row this age contributes to the output matrix. */
export function recordRow(run: Run, age: number): void {
  const { v, s, p, context } = run;
  const born = vbaInt(p.born);

  // Amounts are expressed in the reference year's prices, its wage level, or
  // nominally, depending on the setting.
  let factor: number;
  const refAge = context.referenceYear - born;
  if (context.priceBasis === 1) {
    factor = v.startage <= refAge ? v.kpi.get(refAge) / v.kpi.get(age) : v.kpi.get(v.startage) / v.kpi.get(age);
  } else if (context.priceBasis === 0) {
    factor = v.startage <= refAge ? v.iindex.get(refAge) / v.iindex.get(age) : v.iindex.get(v.startage) / v.iindex.get(age);
  } else {
    factor = 1;
  }

  // `Round(x, i - 1)`: whole kronor normally, two decimals without rounding.
  // `Round(x, i - 1)` with i = 1 normally and 3 without rounding, so whole
  // kronor or two decimals; the two index factors use `i + 3`.
  const i = context.marginal === 1 ? 3 : 1;
  const r = (value: number) => vbaRound(value * factor, i - 1);
  const rf = (value: number) => vbaRound(value, i + 3);

  const row: MvaluesRow = {
    year: v.year.get(age),
    age,
    income: r(v.income.get(age)),
    ip: r(s.ip.get(age)),
    tp: r(s.tp.get(age)),
    pp: r(s.pp.get(age)),
    garp: r(s.garp.get(age)),
    ptillagg: r(s.ptillagg.get(age)),
    tjp: r(s.tjp.get(age)),
    ips: r(s.ips.get(age)),
    brutto: r(s.brutto.get(age)),
    netto: r(s.netto.get(age)),
    bidrag: r(s.bidrag.get(age)),
    pps: r(s.pps.get(age)),
    indDisp: r(s.indDisp.get(age)),
    kpiFactor: rf(v.kpi.get(wsMax(v.startage, refAge)) / v.kpi.get(age)),
    indexFactor: rf(v.iindex.get(wsMax(v.startage, refAge)) / v.iindex.get(age)),
  };
  s.rows.push(row);
}

/** The tax, benefit and output sections for one age, in the VBA's order. */
export function taxAndBenefits(run: Run, age: number, utgyear: number, skyear: number): void {
  const result = taxes(run, age, utgyear, skyear);
  benefits(run, age, utgyear, result);
  recordRow(run, age);
}
