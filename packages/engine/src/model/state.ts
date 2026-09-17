/**
 * The vectors and running totals `Mcalc` carries across the age loop.
 *
 * The VBA declares about fifty-five `ReDim x(startage To slutage)` arrays and a
 * dozen scalars that accumulate between iterations. They fall into two groups
 * with different lifetimes, and keeping them apart is what makes the loop
 * readable:
 *
 * - `SetupVectors` are built once by `startsetup` and only read afterwards --
 *   the index series, the base amounts, the fee and inheritance-gain factors,
 *   the tax rates, and the income and wage path.
 * - `RunState` is what the loop itself fills in: pension rights, balances,
 *   pensions paid, taxes, benefits, and the output matrix.
 *
 * `AgeArray` keeps the VBA's own indexing, so `x.get(age - 1)` reads the way it
 * does in the original and an off-by-one shows up as a thrown range error
 * rather than a silent `undefined`.
 */

import { AgeArray } from "../vba/ageArray.js";
import type { RunVectors } from "./runVectors.js";

/** The last age the model computes, `Public Const slutage = 105`. */
export const SLUTAGE = 105;

/** What `startsetup` produces: one value per age, read-only for the rest of the run. */
export interface SetupVectors {
  /** `startage`: the first age with an income year, 15 unless an own wage vector is used. */
  readonly startage: number;
  /** `slutage`: always 105. */
  readonly slutage: number;

  /** `year_(age)` -- the income year, `Int(born) + age`. */
  readonly year: AgeArray;

  /** `IBB(age)`, `pbb(age)`, `FPB(age)` -- the three base amounts. */
  readonly ibb: AgeArray;
  readonly pbb: AgeArray;
  readonly fpb: AgeArray;

  /** `KPI(age)` annual average, `KPI_j(age)` the June figure. */
  readonly kpi: AgeArray;
  readonly kpiJune: AgeArray;

  /** `Iindex(age)` inkomstindex, `Pindex(age)` the balanced index pensions follow. */
  readonly iindex: AgeArray;
  readonly pindex: AgeArray;

  /** `yield(age)` -- the fund return factor, and `RGK(age)` the Riksgälden rate. */
  readonly yieldFactor: AgeArray;
  readonly rgk: AgeArray;

  /** `MPGI(age)` -- average pension-qualifying income. */
  readonly mpgi: AgeArray;

  /** `IP_avg`, `PP_avg`, `TP_avg` -- what is left after management costs. */
  readonly ipAvg: AgeArray;
  readonly ppAvg: AgeArray;
  readonly tpAvg: AgeArray;

  /** `IP_arv1` up to the riktålder, `IP_arv2` from it, `PP_arv` for premium pension. */
  readonly ipArv1: AgeArray;
  readonly ipArv2: AgeArray;
  readonly ppArv: AgeArray;

  /** `Kom_skatt`, `Begravavg`, `Tax_limit1`, `Tax_limit2`. */
  readonly komSkatt: AgeArray;
  readonly begravavg: AgeArray;
  readonly taxLimit1: AgeArray;
  readonly taxLimit2: AgeArray;

  /** `Income_(age)` taxable earned income, `Wage_(age)` the part that is salary. */
  readonly income: AgeArray;
  readonly wage: AgeArray;
}

/** One row of `mvalues` -- the per-age output the tables and figures are drawn from. */
export interface MvaluesRow {
  /** 1: `year_(age)`. */
  readonly year: number;
  /** 2: age. */
  readonly age: number;
  /** 3: earned income. */
  readonly income: number;
  /** 4: inkomstpension. */
  readonly ip: number;
  /** 5: tilläggspension. */
  readonly tp: number;
  /** 6: premiepension. */
  readonly pp: number;
  /** 7: garantipension. */
  readonly garp: number;
  /** 8: inkomstpensionstillägg. */
  readonly ptillagg: number;
  /** 9: tjänstepension. */
  readonly tjp: number;
  /** 10: private saving, before tax. */
  readonly ips: number;
  /** 11: gross income. */
  readonly brutto: number;
  /** 12: net income. */
  readonly netto: number;
  /** 13: benefits. */
  readonly bidrag: number;
  /** 14: private saving, after tax. */
  readonly pps: number;
  /** 15: individual disposable income. */
  readonly indDisp: number;
  /** 16: the price factor this row was expressed with. */
  readonly kpiFactor: number;
  /** 17: the wage-level factor. */
  readonly indexFactor: number;
  /**
   * 18: municipal tax -- not one of the workbook's own seventeen `mvalues`
   * columns. Kommunal inkomstskatt plus kyrkoskatt, net of the job/union/
   * sickness credits the tax calculation caps against it; what's left of the
   * year's total tax once `stateTax` is taken out. Added for this port's own
   * tax-per-year chart and Table 2's matching columns, the same kind of
   * addition the KPI cards already are.
   */
  readonly municipalTax: number;
  /**
   * 19: state tax -- statlig inkomstskatt, the public-service fee, and any
   * capital-gains tax, none of which the credits above touch. Also not one
   * of the workbook's seventeen.
   */
  readonly stateTax: number;
}

/** Everything the age loop accumulates. */
export interface RunState {
  // ---- Pension-qualifying income and contributions -------------------------
  /** `pgi_(age)`. */
  readonly pgi: AgeArray;
  /** `EgenAvg_(age)` -- the individual's own pension contribution. */
  readonly egenAvg: AgeArray;
  /** `PGB_(age)` -- pensionsgrundande belopp for childcare, conscription, study. */
  readonly pgb: AgeArray;
  /** `SAPGB_(age)` -- the same for sickness and activity compensation. */
  readonly saPgb: AgeArray;
  /** `arb_avg(age)` -- employer contributions, written out for study only. */
  readonly arbAvg: AgeArray;

  // ---- Yearly rights and premiums -----------------------------------------
  /** `IP_ratt`, `PP_ratt`, `GP_ratt`, `TJP_ratt`. */
  readonly ipRatt: AgeArray;
  readonly ppRatt: AgeArray;
  readonly gpRatt: AgeArray;
  readonly tjpRatt: AgeArray;
  /** `TP_points(age)` -- ATP points; sorted in place at retirement, see the note below. */
  readonly tpPoints: AgeArray;
  /**
   * `STP_points(age)` -- a copy of the ATP points taken during the earning
   * phase, kept in age order.
   *
   * NOTE: `Mcalc` sorts `TP_points` in place at retirement and `FTJP` then reads
   * `STP_points` for SAF-LO, which is why the copy exists. The VBA flags it
   * itself ("OBS TP_points är omsorterad"). The order of those two steps is
   * load-bearing.
   */
  readonly stpPoints: AgeArray;
  /** `TP94p(age)` -- ATP points earned up to 1994, for the garantibelopp. Ends at 64. */
  readonly tp94p: AgeArray;

  // ---- Balances ------------------------------------------------------------
  /** `IP_pbh`, `PP_pbh`, `GP_pbh`, `TJP_pbh`, `IPS_pbh`, `PPS_pbh`. */
  readonly ipPbh: AgeArray;
  readonly ppPbh: AgeArray;
  readonly gpPbh: AgeArray;
  readonly tjpPbh: AgeArray;
  readonly ipsPbh: AgeArray;
  readonly ppsPbh: AgeArray;

  // ---- Pensions paid -------------------------------------------------------
  /** `ip`, `tp`, `pp`, `garp`, `ptillagg`, `TJP`, `ips`, `pps`. */
  readonly ip: AgeArray;
  readonly tp: AgeArray;
  readonly pp: AgeArray;
  readonly garp: AgeArray;
  readonly ptillagg: AgeArray;
  readonly tjp: AgeArray;
  readonly ips: AgeArray;
  readonly pps: AgeArray;
  /** `Gbelopp(age)` -- the garantibelopp on rights earned up to 1994. */
  readonly gbelopp: AgeArray;

  // ---- Income, tax and benefits -------------------------------------------
  /** `brutto`, `Netto`, `Bidrag`, `IndDisp`. */
  readonly brutto: AgeArray;
  readonly netto: AgeArray;
  readonly bidrag: AgeArray;
  readonly indDisp: AgeArray;
  /** `kinkskatt`/`statskatt`, kept apart -- see `MvaluesRow.municipalTax`/`stateTax`. */
  readonly municipalTax: AgeArray;
  readonly stateTax: AgeArray;

  // ---- Running scalars -----------------------------------------------------
  /** `IPS_ratt` -- a scalar in the VBA, not a vector. */
  ipsRatt: number;
  /** `atp_year` / `atp_points` -- years with ATP points and the fifteen-year average. */
  atpYear: number;
  atpPoints: number;
  /** `atp94year` / `atp94` -- the same, counted only to 1994. */
  atp94Year: number;
  atp94: number;
  /** `tp_year` -- years with occupational pension. */
  tpYear: number;
  /** `pgi_years` -- years with a pension-qualifying income, for inkomstpensionstillägg. */
  pgiYears: number;
  /** `pgbyears` -- income years up to the first pensionsgrundande belopp. */
  pgbYears: number;
  /** `uttagIP` / `uttagPP` -- the share of each pension currently drawn. */
  uttagIp: number;
  uttagPp: number;
  /** `dtal_ip` / `dtal_pp` -- this year's delningstal. */
  dtalIp: number;
  dtalPp: number;
  /** `mpension` -- the underlying amount inkomstpensionstillägg is computed on. */
  mpension: number;
  /** `gpundtab1` -- that amount as it stood at retirement, which Table 1 reuses. */
  gpundtab1: number;
  /** `gp_und` -- the garantipension's own underlying balance, carried year to year. */
  gpUnd: number;
  /**
   * `kvoten` -- pbb over IBB at the boundary year.
   *
   * Seeded by `prepareRun` and then *reassigned inside the loop*, in the
   * housing-supplement block, so it is run state rather than a constant.
   */
  kvoten: number;
  /** `maxhyra` -- a replacement rent ceiling, set in the same block. */
  maxhyra: number;

  /**
   * What the age loop's own local variables still held when it ended.
   *
   * Mcalc's tax and benefit sections are written inline in one very long
   * procedure, so when the loop finishes its locals are still in scope --
   * and the recomputation at the retirement age (VBA_go.bas 2607-2800) reads
   * several of them rather than working them out again. `pensredukt = rakassa`
   * at :2652 is the plainest example: the a-kassa reduction from the *last*
   * age the loop ran stands in for the pension contribution reduction.
   *
   * Splitting the procedure into functions loses that, so the values the
   * recomputation reaches for are kept here deliberately. See atRetirement.ts.
   */
  readonly leftovers: Leftovers;

  /** The output matrix, one row per age from `startage`. */
  readonly rows: MvaluesRow[];
}

/** The loop locals the post-loop recomputation reads. See `RunState.leftovers`. */
export interface Leftovers {
  /**
   * The loop counter.
   *
   * A completed VBA `For` leaves its counter one past the limit, so it starts
   * at `slutage + 1` -- but `creditLastPensionRight` resets it to the
   * retirement age (`age = PAR`, VBA_go.bas:2445) before the recomputation that
   * reads it. So it is `slutage + 1` only when the final pension right is
   * switched off, and the retirement age otherwise. Both cases are indexed by
   * `year_(mini(age, 100))` at :2614 and `year_(mini(age, slutage))` at :2749.
   */
  age: number;
  /** `Gage` -- the age the higher grundavdrag starts at, from the last iteration. */
  gage: number;
  /** `rakassa` -- which :2652 uses as `pensredukt`. */
  rakassa: number;
  /** `SAavdrag` -- read by the FAavdrag cap at :2658. */
  saavdrag: number;
  /** `ap` / `apm` -- whether the individual and the spouse draw a pension. */
  ap: number;
  apm: number;
  /** `MakaInk` -- the spouse's income. */
  makaInk: number;
  /** `tjpm` -- the spouse's occupational pension. */
  tjpm: number;
  /** `barnbidrag` and `bostadsbidrag`, whose recomputation at :2687 is commented out. */
  barnbidrag: number;
  bostadsbidrag: number;
  /** `hyra_t` -- the rent, indexed to the year. */
  hyraT: number;
  /** `bidragovr` -- benefits paid under the `Rng_Ansokt = 9` special case. */
  bidragovr: number;
}

/** A fresh set of zeroed vectors for `startage..slutage`. */
export function createRunState(startage: number, slutage = SLUTAGE): RunState {
  const age = () => new AgeArray(startage, slutage);
  return {
    pgi: age(),
    egenAvg: age(),
    pgb: age(),
    saPgb: age(),
    arbAvg: age(),

    ipRatt: age(),
    ppRatt: age(),
    gpRatt: age(),
    tjpRatt: age(),
    tpPoints: age(),
    stpPoints: age(),
    // `ReDim TP94p(startage To 64)` -- points after 1994 do not count.
    tp94p: new AgeArray(startage, 64),

    ipPbh: age(),
    ppPbh: age(),
    gpPbh: age(),
    tjpPbh: age(),
    ipsPbh: age(),
    ppsPbh: age(),

    ip: age(),
    tp: age(),
    pp: age(),
    garp: age(),
    ptillagg: age(),
    tjp: age(),
    ips: age(),
    pps: age(),
    gbelopp: age(),

    brutto: age(),
    netto: age(),
    bidrag: age(),
    indDisp: age(),
    municipalTax: age(),
    stateTax: age(),

    ipsRatt: 0,
    atpYear: 0,
    atpPoints: 0,
    atp94Year: 0,
    atp94: 0,
    tpYear: 0,
    pgiYears: 0,
    pgbYears: 0,
    uttagIp: 0,
    uttagPp: 0,
    dtalIp: 0,
    dtalPp: 0,
    mpension: 0,
    gpundtab1: 0,
    gpUnd: 0,
    kvoten: 1,
    maxhyra: 0,

    leftovers: {
      age: slutage + 1,
      gage: 0,
      rakassa: 0,
      saavdrag: 0,
      ap: 0,
      apm: 0,
      makaInk: 0,
      tjpm: 0,
      barnbidrag: 0,
      bostadsbidrag: 0,
      hyraT: 0,
      bidragovr: 0,
    },

    rows: [],
  };
}

/**
 * The `RunVectors` view the rule modules already take.
 *
 * They were written against Mcalc's globals one module at a time, before the
 * loop existed; this is the adapter that finally supplies them for real.
 */
export function runVectors(setup: SetupVectors): RunVectors {
  return {
    wage: (age) => setup.wage.getOrZero(age),
    ibb: (age) => setup.ibb.getOrZero(age),
    pbb: (age) => setup.pbb.getOrZero(age),
    fpb: (age) => setup.fpb.getOrZero(age),
    kpiJune: (age) => setup.kpiJune.getOrZero(age),
    kpi: (age) => setup.kpi.getOrZero(age),
    year: (age) => setup.year.getOrZero(age),
  };
}
