/**
 * The result model: Table 1, Table 2, the life-income sums and the warnings.
 *
 * Port of VBA_go.bas 2119-2844, the part of Mcalc that runs once the age loop
 * is done. In the original it writes straight into the Start sheet; here it
 * returns the same figures as data.
 *
 * Both of the workbook's figures are Excel charts bound to Table 2's range --
 * `mdlChartData.bas` only shows and hides them -- so Table 2 *is* the chart
 * data and there is no separate computation for them.
 */

import { gp, tillagg } from "../bidrag/garantipension.js";
import { recomputeAtRetirement } from "./atRetirement.js";
import { deltal, incomePensionYear } from "../pension/incomePension.js";
import { tpFaktor } from "../pension/atp.js";
import { vbaInt, wsMax, wsMin } from "../vba/math.js";
import type { Run } from "./mcalc.js";
import type { MvaluesRow } from "./state.js";
import type { Warning } from "./setup.js";

/** One row of Table 1: an amount, four ways. */
export interface Table1Row {
  /** Which line of the table this is. */
  readonly key: Table1Key;
  /** Column A: the amount as the model computed it, nominal. */
  readonly nominal: number;
  /** Column B: in the reference year's prices, wage level, or nominal. */
  readonly adjusted: number;
  /** Column C: the same, per month. */
  readonly monthly: number;
  /** Column D: as a share of the final salary. */
  readonly shareOfFinalSalary: number;
}

export const Table1Key = {
  FinalSalary: "slutlon",
  SalaryAfterTax: "lonEfterSkatt",
  DisposableBeforeRetirement: "dispInkomst",
  IncomePension: "ip",
  SupplementaryPension: "tp",
  PremiumPension: "pp",
  GuaranteePension: "garp",
  IncomePensionSupplement: "ptillagg",
  TotalPublicPension: "totAllmanPension",
  OccupationalPension: "tjp",
  PrivateSaving: "ips",
  TotalGross: "totBrutto",
  PrivateSavingAfterTax: "pps",
  /** `rng_Tabell1_Efterskatt`: the pension after tax. */
  PensionAfterTax: "efterSkatt",
  /** `rng_Tabell1_Bidrag`: benefits at the retirement age. */
  BenefitsAtRetirement: "bidrag",
  /** `rng_Tabell1_Disp_efterskatt`: disposable income in retirement. */
  DisposableAtRetirement: "dispEfterSkatt",
} as const;

export type Table1Key = (typeof Table1Key)[keyof typeof Table1Key];

/** One row of Table 2, the cash-flow table both figures are drawn from. */
export interface Table2Row {
  readonly year: number;
  readonly age: number;
  /** Earned income. */
  readonly salary: number;
  /** Income pension and tilläggspension together. */
  readonly incomeAndSupplementary: number;
  readonly premium: number;
  /** Occupational pension and private saving together. */
  readonly occupationalAndPrivate: number;
  /** Garantipension and inkomstpensionstillägg together. */
  readonly guaranteeAndSupplement: number;
  readonly gross: number;
  /** Kommunal inkomstskatt and kyrkoskatt, net of the credits capped against it. */
  readonly municipalTax: number;
  /** Statlig inkomstskatt, the public-service fee, and any capital-gains tax. */
  readonly stateTax: number;
  readonly net: number;
  readonly benefits: number;
  /** Private saving paid out of an ISK or KF, which is not taxed as income. */
  readonly privateAfterTax: number;
  readonly disposable: number;
}

/** The three discounted sums over a retirement. */
export interface LifeIncome {
  readonly gross: number;
  readonly net: number;
  readonly disposable: number;
  /** Expected age at death, which the sums run to. */
  readonly throughAge: number;
}

/** Everything a run produces. */
export interface TypfallResult {
  /** The summary at retirement. */
  readonly table1: readonly Table1Row[];
  /** The cash-flow table, and the data behind both figures. */
  readonly table2: readonly Table2Row[];
  /** The per-age matrix Table 2 is cut from. */
  readonly rows: readonly MvaluesRow[];
  readonly lifeIncome: LifeIncome;
  /** Years with a pension right, capped at 40, which labels the IPT row. */
  readonly qualifyingYears: number;
  /** Corrections the run made to the inputs it was given. */
  readonly warnings: readonly Warning[];
}

/**
 * `Last_pratt`: one more year of pension rights, credited at retirement.
 *
 * A pension right is settled by the following year's assessment, so the year of
 * retirement would otherwise lose its own. With `rng_Sista_PensRatt` on -- which
 * is the shipped setting -- the retirement year's balance is grossed back up by
 * the divisor, the final right added, and the pension recomputed.
 */
function creditLastPensionRight(run: Run): void {
  const { v, s, p, context, deltalTables } = run;
  const par = vbaInt(p.par);
  const year = v.year.get(par);
  const born = vbaInt(p.born);
  const marginal = context.marginal;

  // `deltal(PAR, Int(born), PAR, def_ar, 4)` and `..., 19` (VBA_go.bas 2408 and
  // 2411, then again at 2434 and 2436) -- the Nyckeltal sheet, not the spliced
  // arrays `fnDeltal_*` read. It also gives cohorts born 1937 or earlier a
  // single divisor rather than an age table, which has no rows for them.
  const dtalIp = deltal(p.par, born, par, deltalTables, "income", p.defAr);
  const dtalPp = deltal(p.par, born, par, deltalTables, "premium", p.defAr);

  // Undo the annuitisation, add the final right, and divide again.
  s.ipPbh.set(par, (12 / run.pmonth) * s.ip.get(par) * dtalIp + s.ipPbh.get(par));
  s.tp.set(par, s.tp.get(par) * (12 / run.pmonth));
  s.ppPbh.set(par, (12 / run.pmonth) * s.pp.get(par) * dtalPp + s.ppRatt.get(par));

  s.ip.set(par, dtalIp > 0 ? s.ipPbh.get(par) / dtalIp : 0);
  s.pp.set(par, dtalPp > 0 ? s.ppPbh.get(par) / dtalPp : 0);

  // `age = PAR` (VBA_go.bas:2445). The loop's counter is reset here, inside the
  // final-pension-right block and before the tax and benefit recomputation that
  // reads it -- so `year_(mini(age, 100))` there is the retirement year, not the
  // year the loop happened to stop at. Every other loop local keeps its last
  // value; see `RunState.leftovers`.
  s.leftovers.age = par;

  // Garantipension is measured on its own divisor, at the riktålder:
  // `deltal(riktalder, Int(born), riktalder, riktalder, 4)` (VBA_go.bas:2441).
  const gpDtal = deltal(p.riktalder, born, p.riktalder, deltalTables, "income", p.riktalder);

  if (gpDtal > 0 && p.par >= p.riktalder) {
    let gpUnd = incomePensionYear(
      year, wsMax(p.par, p.riktalder), p.born, s.gpRatt.get(par),
      v.ipArv1.get(par), v.ipArv2.get(par), v.ipAvg.get(par), s.gpPbh.getOrZero(par - 1), 1,
      s.garp.getOrZero(par - 1), gpDtal, v.pindex.get(par) / v.pindex.getOrZero(par - 1),
      p.defAr, marginal, v.pindex.get(par) / v.iindex.get(par),
    ).pension;
    gpUnd += s.gpPbh.get(par) / gpDtal;
    if (marginal === 0) gpUnd = vbaInt((gpUnd + 0.49) / 12) * 12;

    const utgyear =
      context.rulesFromUtg !== 0 &&
      ((born + p.par > context.rulesFromUtg && context.rules === 0) || context.rules === 1)
        ? context.rulesFromUtg
        : year;

    s.garp.set(
      par,
      gp(
        gpUnd + s.tp.get(par), p.civ, born, v.pbb.get(par), run.forstid, marginal,
        par, utgyear, run.iyear, v.ibb.get(par), s.kvoten, p.riktalder, s.uttagIp,
      ),
    );
  } else {
    // A part-year between par and the riktålder can still qualify later.
    s.garp.set(par, 0);
  }

  // The supplement is recomputed on the amount Table 1 kept from the loop.
  const mpension = s.gpundtab1 * 12;
  const index2021 =
    2021 - born > v.startage ? v.iindex.getOrZero(2021 - born) : 182.58;
  s.ptillagg.set(
    par,
    tillagg(mpension, year, p.born, v.iindex.get(par), index2021, s.uttagIp, s.pgiYears, marginal),
  );
  if (year === 2021) {
    s.ptillagg.set(par, (s.ptillagg.get(par) * wsMin(4, run.pmonth)) / 12);
  }

  if (marginal === 0) {
    s.ip.set(par, vbaInt((s.ip.get(par) + 0.49) / 12) * 12);
    s.pp.set(par, vbaInt((s.pp.get(par) + 0.49) / 12) * 12);
  }
}

/**
 * The price or wage-level factors Table 1 expresses its amounts in.
 *
 * `justering0` is for the final salary, which is measured a year before
 * retirement; `justering(0)` for the pensions, measured in the retirement year;
 * `justering(i)` for each earlier year the final salary averages over.
 */
function adjustmentFactors(run: Run): { atRetirement: number; beforeRetirement: number; perYear: number[] } {
  const { v, p, context } = run;
  const born = vbaInt(p.born);
  const par = vbaInt(p.par);
  // A birthday part-way through the year puts retirement in the following one.
  const korr = vbaInt(p.par + p.born) > par + born ? 1 : 0;
  const timeLag = context.pensionSameYearAsFinalSalary;
  const refAge = wsMax(v.startage, context.referenceYear - born + korr);
  const years = context.finalSalaryYears;

  if (context.priceBasis === -1) {
    return { atRetirement: 1, beforeRetirement: 1, perYear: new Array<number>(years + 1).fill(1) };
  }

  const series = context.priceBasis === 1 ? v.kpi : v.iindex;
  const at = (age: number) => series.getOrZero(age);
  const perYear: number[] = [1];
  if (years > 1) {
    for (let i = 1; i <= years; i += 1) perYear[i] = at(refAge) / at(par - i + korr);
  }
  return {
    atRetirement: at(refAge) / at(par + korr),
    beforeRetirement: at(refAge) / at(par - 1 + timeLag + korr),
    perYear,
  };
}

/**
 * The retirement year's salary is zeroed and the gross recomputed from it.
 *
 * VBA_go.bas 2507 and 2533, between the final pension right and Table 1: the
 * table reports a full year of pension against the previous year's salary, and
 * the gross has to be rebuilt because `creditLastPensionRight` has just changed
 * the components. `pps` and `kapital` are outside it, as the original leaves
 * them (`'+ kapital` is commented out on the line itself).
 */
function closeRetirementYear(run: Run): void {
  const { v, s, p } = run;
  const par = vbaInt(p.par);

  v.income.set(par, 0);
  v.wage.set(par, 0);

  s.brutto.set(
    par,
    v.income.get(par) + s.ip.get(par) + s.tp.get(par) + s.pp.get(par) + s.garp.get(par) +
      s.tjp.get(par) + s.ips.get(par) + s.ptillagg.get(par),
  );
}

/**
 * Table 1: the summary at retirement.
 *
 * `closeRetirementYear` must have run first, and `recomputeAtRetirement` too
 * where the settings call for it -- `buildResult` does both in order.
 */
export function buildTable1(run: Run): Table1Row[] {
  const { v, s, p, context } = run;
  const par = vbaInt(p.par);
  const born = vbaInt(p.born);
  const korr = vbaInt(p.par + p.born) > par + born ? 1 : 0;
  const years = context.finalSalaryYears;
  const factors = adjustmentFactors(run);

  // `growth` -- reused as a scratch variable in the original, overwriting the
  // real growth assumption -- puts a part-year salary on a full-year footing.
  const previousSalary = v.income.getOrZero(par + korr - 1);
  const growth =
    run.pmonth === 12 || previousSalary === 0
      ? 1
      : (v.income.getOrZero(par + korr) * 12) / (12 - run.pmonth) / previousSalary;

  let finalSalaryNominal: number;
  let finalSalaryAdjusted: number;
  let netNominal: number;
  let netAdjusted: number;
  let dispNominal: number;
  let dispAdjusted: number;

  if (years > 1) {
    let sIncome = v.income.getOrZero(par - 1);
    let sNet = s.netto.getOrZero(par - 1);
    let sDisp = s.indDisp.getOrZero(par - 1);
    for (let i = 2; i <= years; i += 1) {
      sIncome += v.income.getOrZero(par - i);
      sNet += s.netto.getOrZero(par - i);
      sDisp += s.indDisp.getOrZero(par - i);
    }
    finalSalaryNominal = sIncome / years;
    netNominal = sNet / years;
    dispNominal = sDisp / years;

    // The adjusted column weights the most recent year differently from the
    // rest, which is what the 2024 and 2025 corrections in the VBA are about.
    let aIncome = v.income.getOrZero(par - 1) * factors.beforeRetirement;
    let aNet = s.netto.getOrZero(par - 1) * factors.beforeRetirement;
    let aDisp = s.indDisp.getOrZero(par - 1) * factors.beforeRetirement;
    for (let i = 2; i <= years; i += 1) {
      aIncome += v.income.getOrZero(par - i) * (factors.perYear[i] ?? 1);
      aNet += s.netto.getOrZero(par - i) * (factors.perYear[i] ?? 1);
      aDisp += s.indDisp.getOrZero(par - i) * (factors.perYear[i] ?? 1);
    }
    finalSalaryAdjusted = aIncome / years;
    netAdjusted = aNet / years;
    dispAdjusted = aDisp / years;
  } else {
    finalSalaryNominal = v.income.getOrZero(par - 1 + korr) * growth;
    netNominal = s.netto.getOrZero(par - 1 + korr) * growth;
    dispNominal = s.indDisp.getOrZero(par - 1 + korr) * growth;
    finalSalaryAdjusted = finalSalaryNominal * factors.beforeRetirement;
    netAdjusted = netNominal * factors.beforeRetirement;
    dispAdjusted = dispNominal * factors.beforeRetirement;
  }

  // `brutto(Int(PAR))` as `closeRetirementYear` rebuilt it, rather than a second
  // sum of the same components -- so this row cannot drift from what the tax
  // and benefit recomputation was given.
  const gross = s.brutto.get(par);
  const publicTotal =
    s.ip.get(par) + s.tp.get(par) + s.pp.get(par) + s.garp.get(par) + s.ptillagg.get(par);

  /**
   * What the share column divides by.
   *
   * NOTE: the original reassigns `Income_(PAR - 1)` for this (VBA_go.bas 2568),
   * and the reassignment does not match the value rows it is compared against.
   * With `Average_Earning > 1` it takes `ys(1)`, which is `finalSalaryAdjusted`
   * exactly; with a single year it takes `Income_(PAR - 1 - korr2)` where the
   * value row reads `Income_(PAR - 1 + korr)`, and it drops the `growth`
   * factor the value row applies.
   *
   * Both differences are unreachable from the Start sheet: `korr2` is
   * `rng_alt_last_pratt`, which is 0, and `korr` and `growth` only move when the
   * birth year carries a fraction, which the workbook's dropdown does not offer.
   * So this uses `finalSalaryAdjusted` for both branches -- but a later feature
   * that allows a birth month would have to come back to this line.
   */
  const base = finalSalaryAdjusted;
  const pensionRow = (key: Table1Key, amount: number): Table1Row => ({
    key,
    nominal: amount,
    adjusted: amount * factors.atRetirement,
    monthly: (amount * factors.atRetirement) / 12,
    shareOfFinalSalary: base > 0 ? (amount * factors.atRetirement) / base : 0,
  });

  const salaryRow = (key: Table1Key, nominal: number, adjusted: number): Table1Row => ({
    key,
    nominal,
    adjusted,
    monthly: adjusted / 12,
    shareOfFinalSalary: base > 0 ? adjusted / base : 0,
  });

  const table: Table1Row[] = [
    salaryRow(Table1Key.FinalSalary, finalSalaryNominal, finalSalaryAdjusted),
    salaryRow(Table1Key.SalaryAfterTax, netNominal, netAdjusted),
    salaryRow(Table1Key.DisposableBeforeRetirement, dispNominal, dispAdjusted),
    pensionRow(Table1Key.IncomePension, s.ip.get(par)),
    pensionRow(Table1Key.SupplementaryPension, s.tp.get(par)),
    pensionRow(Table1Key.PremiumPension, s.pp.get(par)),
    pensionRow(Table1Key.GuaranteePension, s.garp.get(par)),
    pensionRow(Table1Key.IncomePensionSupplement, s.ptillagg.get(par)),
    pensionRow(Table1Key.TotalPublicPension, publicTotal),
    pensionRow(Table1Key.OccupationalPension, s.tjp.get(par)),
    pensionRow(Table1Key.PrivateSaving, s.ips.get(par)),
    pensionRow(Table1Key.PrivateSavingAfterTax, s.pps.get(par)),
    pensionRow(Table1Key.TotalGross, gross),
  ];

  // The last three rows measure themselves against different denominators from
  // the pension rows: net against net, and both benefits and disposable income
  // against disposable income (VBA_go.bas 2679, 2776 and 2790).
  const against = (amount: number, base: number, guard: number): number =>
    guard > 0 ? (amount * factors.atRetirement) / base : 0;

  const netAtRetirement = s.netto.get(par);
  const benefitsAtRetirement = s.bidrag.get(par);
  const dispAtRetirement = s.indDisp.get(par);

  table.push(
    {
      key: Table1Key.PensionAfterTax,
      nominal: netAtRetirement,
      adjusted: netAtRetirement * factors.atRetirement,
      monthly: (netAtRetirement * factors.atRetirement) / 12,
      shareOfFinalSalary: against(netAtRetirement, netAdjusted, s.netto.getOrZero(par - 1)),
    },
    {
      key: Table1Key.BenefitsAtRetirement,
      nominal: benefitsAtRetirement,
      adjusted: benefitsAtRetirement * factors.atRetirement,
      monthly: (benefitsAtRetirement * factors.atRetirement) / 12,
      // QUIRK: the guard tests last year's *benefits* but the division is by
      // last year's *disposable income* (VBA_go.bas:2776). So a household with
      // no benefits the year before retirement reports a share of 0 however
      // much it receives after.
      shareOfFinalSalary: against(benefitsAtRetirement, dispAdjusted, s.bidrag.getOrZero(par - 1)),
    },
    {
      key: Table1Key.DisposableAtRetirement,
      nominal: dispAtRetirement,
      adjusted: dispAtRetirement * factors.atRetirement,
      monthly: (dispAtRetirement * factors.atRetirement) / 12,
      shareOfFinalSalary: against(dispAtRetirement, dispAdjusted, s.indDisp.getOrZero(par - 1)),
    },
  );

  return table;
}

/** Table 2, from the start age the settings ask for. */
export function buildTable2(run: Run): Table2Row[] {
  const { p, s, context } = run;
  const delat = context.chartEarningFactor === 0 ? 1 : context.chartEarningFactor;

  // 0 derives the start age as `Int(par - 10)`, which is the sheet's own rule.
  let start = context.table2StartAge === 0 ? vbaInt(p.par - 10) : context.table2StartAge;
  if (start > run.v.slutage || start < run.v.startage) {
    start = start < run.v.startage ? run.v.startage : vbaInt(p.par) - 5;
  }

  return s.rows
    .filter((row) => row.age >= start)
    .map((row) => ({
      year: row.year,
      age: row.age,
      salary: row.income / delat,
      incomeAndSupplementary: (row.ip + row.tp) / delat,
      premium: row.pp / delat,
      occupationalAndPrivate: (row.tjp + row.ips) / delat,
      guaranteeAndSupplement: (row.garp + row.ptillagg) / delat,
      gross: row.brutto / delat,
      municipalTax: row.municipalTax / delat,
      stateTax: row.stateTax / delat,
      net: row.netto / delat,
      benefits: row.bidrag / delat,
      privateAfterTax: row.pps / delat,
      disposable: row.indDisp / delat,
    }));
}

/**
 * The three discounted sums over a retirement, from the pension age to the
 * expected age at death.
 */
export function lifeIncome(run: Run): LifeIncome {
  const { p, s, context, deltalTables } = run;
  const par = p.par;
  const rate = context.discountRate;
  const throughAge = deltalTables.expectedLife(vbaInt(p.born), vbaInt(par)) + par;

  let gross = 0;
  let net = 0;
  let disposable = 0;

  for (const row of s.rows) {
    const counter = row.age;
    const discount = (1 + rate) ** (counter - par);
    if (counter >= par && vbaInt(throughAge) > counter) {
      gross += row.brutto / discount;
      net += row.netto / discount;
      disposable += row.indDisp / discount;
    }
    if (vbaInt(throughAge) === counter && throughAge > par) {
      // NOTE: the part-year at the end adds the *disposable* figure to all
      // three sums, including the gross and net ones. Kept as written.
      const part = ((throughAge - vbaInt(throughAge)) * row.indDisp) / discount;
      gross += part;
      net += part;
      disposable += part;
    }
  }

  return { gross, net, disposable, throughAge };
}

/** Assembles the whole result, after the loop has run. */
export function buildResult(run: Run, warnings: readonly Warning[]): TypfallResult {
  if (run.context.lastPensionRight > 0) creditLastPensionRight(run);

  closeRetirementYear(run);

  // With `rng_Sista_PensRatt` off, Table 1 reports the tax, benefits and
  // disposable income the loop computed at that age; the second pass exists
  // only because the final pension right has just changed the gross. It writes
  // its results back into the run state, which is where `buildTable1` reads
  // them from either way.
  if (run.context.lastPensionRight > 0) recomputeAtRetirement(run);

  return {
    table1: buildTable1(run),
    table2: buildTable2(run),
    rows: run.s.rows,
    lifeIncome: lifeIncome(run),
    qualifyingYears: wsMin(run.s.pgiYears, 40),
    warnings,
  };
}

export { tpFaktor };
