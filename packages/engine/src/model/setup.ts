/**
 * `startsetup` -- everything the age loop needs before it can start.
 *
 * Port of `startsetup` in VBA_go.bas. It validates the typfall, then fills one
 * value per age for every series the loop reads: the index series and base
 * amounts, the fund return, the fee and inheritance-gain factors, the tax rates,
 * and the income and wage path.
 *
 * Most of the index work in the original is reading Nyckeltal columns 83-93 and
 * falling back to a projection once the sheet runs out. `src/data/projection.ts`
 * already reproduces all twenty series against the workbook's own cached values,
 * so that reading is a lookup here and the fallback is the projection itself.
 *
 * DEVIATION: the original validates by dialog -- a Yes/No box for a retirement
 * age below the riktålder, message boxes elsewhere -- and writes its corrections
 * back into the sheet. Here every correction is the one the workbook makes when
 * you accept it, and is recorded in `warnings` so the caller can say what it did.
 */

import inheritanceJson from "../../../data/inheritance-gains.json" with { type: "json" };
import municipalTaxJson from "../../../data/municipal-tax.json" with { type: "json" };
import rawSeriesJson from "../../../data/economic-series.json" with { type: "json" };
import { projectEconomicData } from "../data/projection.js";
import type { EconomicData } from "../data/projection.js";
import type { RawEconomicSeries } from "../data/types.js";
import { wages, withdrawalShare } from "../income/wages.js";
import type { WageProfile } from "../income/wages.js";
import { andel, riktage } from "../pension/contributions.js";
import { lowestPensionAge, riktalderFor } from "../pension/retirementAges.js";
import { AgeArray } from "../vba/ageArray.js";
import { vbaInt } from "../vba/math.js";
import type { ModelContext } from "./context.js";
import type { TypfallInput } from "./input.js";
import { SLUTAGE } from "./state.js";
import type { SetupVectors } from "./state.js";

const RAW = rawSeriesJson as unknown as RawEconomicSeries;

/** A correction the workbook would have asked about, and what was used instead. */
export interface Warning {
  /** The named range or input field, as the workbook calls it. */
  readonly field: string;
  /** What the run was asked for. */
  readonly given: number | string;
  /** What it used instead. */
  readonly used: number | string;
  /** The workbook's own wording, where it has one. */
  readonly message: string;
}

/** The scalars `startsetup` resolves, which the loop and the tables both read. */
export interface RunProfile {
  /** `born`. */
  readonly born: number;
  /** `startage`, and `slutage` which is always 105. */
  readonly startage: number;
  readonly slutage: number;
  /** `PAR` -- the age the public pension is first drawn, after validation. */
  readonly par: number;
  /** `tjp_par` -- the same for the occupational pension. */
  readonly tjpPar: number;
  /** `def_ar` -- the age withdrawal becomes final and full. */
  readonly defAr: number;
  /** `W_start` -- the age working life starts, after validation. */
  readonly wStart: number;
  /** `Income` -- the entered salary, annualised. */
  readonly income: number;
  /** `avtal` -- the collective agreement. */
  readonly avtal: number;
  /** `civ` -- 0 single, 1 cohabiting. */
  readonly civ: number;
  /** `andelnya` -- the share of the new system for this cohort. */
  readonly andelnya: number;
  /** `riktl` / `riktalder` -- the lower and the ordinary riktålder. */
  readonly riktl: number;
  readonly riktalder: number;
  /** `Modell_year` -- the workbook's own current year. */
  readonly modelYear: number;
  /** `w_time` -- the age the entered salary refers to. */
  readonly wTime: number;
}

export interface SetupResult {
  readonly vectors: SetupVectors;
  readonly profile: RunProfile;
  readonly data: EconomicData;
  readonly warnings: readonly Warning[];
}

interface Grid {
  readonly firstAge: number;
  readonly lastAge: number;
  readonly firstYear: number;
  readonly lastYear: number;
  readonly values: readonly (readonly (number | null)[])[];
}

/** One cell of an inheritance-gain grid, 0 outside it -- as an empty sheet cell reads. */
function gridAt(grid: Grid, age: number, year: number): number {
  const row = grid.values[age - grid.firstAge];
  if (row === undefined) return 0;
  return row[year - grid.firstYear] ?? 0;
}

const ARV_IP = inheritanceJson.incomePension as Grid;
const ARV_IP_DOUBLE = inheritanceJson.incomePensionDouble as Grid;
const ARV_PP = inheritanceJson.premiumPension as Grid;

const TAX = municipalTaxJson as {
  firstYear: number;
  lastYear: number;
  series: Record<string, { values: (number | null)[] }>;
};

/** A K_skatt column for one year, as a rate rather than a percentage. */
function taxRate(name: string, year: number): number {
  const values = TAX.series[name]!.values;
  const index = Math.min(Math.max(year - TAX.firstYear, 0), values.length - 1);
  return (values[index] ?? 0) / 100;
}

/**
 * Runs the workbook's validations, correcting as it would when accepted.
 *
 * The retirement age is the one the original refuses outright: below the lower
 * riktålder it offers to move it up, and `Mcalc` exits without computing if the
 * offer is declined. Taking the offer is the only branch that produces a result.
 */
function validate(
  input: TypfallInput,
  context: ModelContext,
  warnings: Warning[],
): { par: number; tjpPar: number; defAr: number; wStart: number; startage: number; riktl: number } {
  const born = input.born;
  const ownIncome = input.ownIncome !== undefined;

  // `startage` comes from Adv_settings, lowered further if work starts earlier.
  let startage = ownIncome
    ? 15
    : Math.min(context.modelStartAge, input.startWorkAge);

  // `Rng_riktL`, a cell looked up by cohort -- not `riktage(year, 0)`, which is
  // keyed on the income year and gives a different answer for about half the
  // cohorts. See pension/retirementAges.ts.
  let riktl = lowestPensionAge(born);
  if (context.rulesFromUtg !== 0) {
    // Pinning the expenditure rules pins the riktålder with them.
    if (context.rules === 1 || vbaInt(born) + riktl > context.rulesFromUtg) {
      riktl = riktage(context.rulesFromUtg, 0);
    }
  }

  let par = input.retirementAge;
  if (par < riktl) {
    warnings.push({
      field: "ParYear",
      given: par,
      used: riktl,
      message: `Allmän pension först möjlig vid ${riktl} års ålder, räknar pensionsålder vid ${riktl}`,
    });
    par = riktl;
  }

  // 0 follows the public pension age; see the note on ModelContext.tjpPar.
  let tjpPar = context.tjpPar === 0 ? par : context.tjpPar;
  if (tjpPar < 55) {
    warnings.push({
      field: "tjp_par",
      given: tjpPar,
      used: par,
      message: `Tjänstepension först möjlig vid 55 år, räknar pensionsålder vid ${par}`,
    });
    tjpPar = par;
  }

  let defAr = context.defAr;
  if (defAr < par || defAr === 0) defAr = par;
  if (defAr > 100) {
    warnings.push({
      field: "rng_def_ar",
      given: context.defAr,
      used: par,
      message: `Definitiv pensionsålder är inte möjlig/sannolik, ändrad till ${par}`,
    });
    defAr = par;
  }

  let wStart = input.startWorkAge;
  if (wStart < startage) {
    warnings.push({
      field: "wStartYear",
      given: wStart,
      used: startage,
      message: `Beräkningen börjar vid ${startage} års ålder`,
    });
    wStart = startage;
  }

  // An own income vector always starts at 15; otherwise the span has to reach
  // back to whichever of the two start ages is earlier.
  if (!ownIncome) startage = Math.min(startage, wStart);

  return { par, tjpPar, defAr, wStart, startage, riktl };
}

/**
 * Builds the run's vectors and resolved scalars.
 *
 * @param input   the typfall
 * @param context the advanced settings
 */
export function startsetup(input: TypfallInput, context: ModelContext): SetupResult {
  const warnings: Warning[] = [];
  const born = input.born;
  const slutage = SLUTAGE;

  const { par, tjpPar, defAr, wStart, startage, riktl } = validate(input, context, warnings);

  // `w_ref` is `YEAR(NOW()) - 1`, so the model year is one past it.
  const modelYear = context.modelYear === 0 ? context.referenceYear + 1 : context.modelYear;
  // `w_time` is the typfall's age in the reference year: the Adv_settings cell
  // is a subtraction of two named ranges, and its shipped 66 is exactly
  // `w_ref (2025) - born (1959)`.
  const wTime = context.wTime === 0 ? context.referenceYear - vbaInt(born) : context.wTime;

  const data = projectEconomicData(
    RAW,
    {
      yearlyInflation: input.yearlyInflation,
      realGrowth: input.realGrowth,
      realReturn: input.realReturn,
      marginal: context.marginal,
      returnsNetOfFees: context.returnsNetOfFees,
      latestIndexBasis: context.latestIndexBasis,
      rgk: context.rgk,
    },
    { firstYear: 1957, lastYear: vbaInt(born) + slutage + 1 },
  );

  const make = () => new AgeArray(startage, slutage);
  const year = make();
  const ibb = make();
  const pbb = make();
  const fpb = make();
  const kpi = make();
  const kpiJune = make();
  const iindex = make();
  const pindex = make();
  const yieldFactor = make();
  const rgk = make();
  const mpgi = make();
  const ipAvg = make();
  const ppAvg = make();
  const tpAvg = make();
  const ipArv1 = make();
  const ipArv2 = make();
  const ppArv = make();
  const komSkatt = make();
  const begravavg = make();
  const taxLimit1 = make();
  const taxLimit2 = make();
  const income = make();
  const wage = make();

  // Below 0.1 the setting means "use the historical average rate".
  const historicalTaxRate = context.kommunalskatt < 0.1;

  const profile: WageProfile = {
    born,
    par,
    defAr,
    wageProfile: context.wageProfile,
    workDuringPartialWithdrawal: context.workDuringPartialWithdrawal,
    partialWithdrawalShare: context.uttagIp,
  };

  const ownIncomeByAge =
    input.ownIncome === undefined
      ? undefined
      : new Map(input.ownIncome.map((row) => [row.age, row] as const));

  let uttagIp = 0;

  for (let age = startage; age <= slutage; age += 1) {
    const y = vbaInt(born) + age;
    year.set(age, y);

    if (y > 1959) {
      kpiJune.set(age, data.kpiJune.at(y));
      kpi.set(age, data.kpiAnnual.at(y));
      pbb.set(age, data.prisbasbelopp.at(y));
      ibb.set(age, data.inkomstbasbelopp.at(y));
      fpb.set(age, data.forhojtPrisbasbelopp.at(y));
      mpgi.set(age, data.medelPgi.at(y));
      iindex.set(age, data.inkomstindex.at(y));
      pindex.set(age, data.gallandeIndex.at(y));

      // The VBA rebuilds these from the previous year when the sheet has run
      // out of rows. The projection is that same computation, so there is
      // nothing left to fall back to here.
      if (pindex.get(age) < 1 || pindex.get(age) > iindex.get(age)) {
        pindex.set(age, iindex.get(age));
      }
    } else {
      // Before 1960 the return is assumed at 9% and prices held at their 1960
      // level. None of it reaches a pension.
      yieldFactor.set(age, 1.09);
      kpiJune.set(age, 25.39);
      kpi.set(age, 25.39);
      pbb.set(age, 4200);
      ibb.set(age, 4200);
      fpb.set(age, 4200);
      iindex.set(age, 6.54 / 1.06 ** (1960 - y));
    }

    if (y > 1959) {
      if (context.returnBasis === 1) {
        // The chosen real return, carried through inflation.
        let factor =
          age > startage
            ? (1 + input.realReturn) * (kpi.get(age) / kpi.get(age - 1))
            : 1 + input.realReturn;
        // Gross of fund fees: take 0.2% off.
        if (!context.returnsNetOfFees) factor = factor - 0.002;
        yieldFactor.set(age, factor);
      } else {
        // NOTE: the VBA splits this on `Modell_year > year_(age)` into two
        // branches that are character for character the same. Kept as one.
        const series = context.returnBasis === 2 ? data.avkastningPpm : data.avkastningAp7;
        yieldFactor.set(age, 1 + series.at(y));
      }
    }

    // Temporary management by Riksgälden, from 1996.
    rgk.set(age, y > 1995 ? 1 + data.rantaRiksgalden.at(y) / 100 : 1);

    // What is left after management costs.
    ipAvg.set(age, y < 1960 ? 1 : data.kvarEfterAdminIp.at(y));
    ppAvg.set(age, y < 2000 ? 1 : data.kvarEfterAvgiftPp.at(y));
    // The occupational fund market is assumed as efficient as the PPM one.
    tpAvg.set(age, ppAvg.get(age));

    // Inheritance gains.
    if (age < 17 || y < 2000) {
      ipArv1.set(age, 1);
      ipArv2.set(age, 1);
    } else {
      if (age < riktl) {
        // NOTE: the sheet has no rows above age 61, because the double factor
        // takes over at the riktålder; the VBA's `If ... = 0 Then 1` is what
        // covers the gap between 61 and a riktålder above it.
        const factor = gridAt(ARV_IP, age, y);
        ipArv1.set(age, factor === 0 ? 1 : factor);
        ipArv2.set(age, 1);
        // The year before the riktålder already carries the double gain.
        if (age === riktl - 1) ipArv2.set(age, gridAt(ARV_IP_DOUBLE, age, y));
        if (ipArv2.get(age) === 0) ipArv2.set(age, 1);
      }
      if (age >= riktl) {
        ipArv1.set(age, 1);
        ipArv2.set(age, gridAt(ARV_IP_DOUBLE, age, y));
        if (ipArv2.get(age) === 0) ipArv2.set(age, ipArv2.getOrZero(age - 1));
        if (ipArv2.get(age) === 0) ipArv2.set(age, 1);
      }
    }

    if (age < 15 || y < 2003) {
      ppArv.set(age, 1);
    } else {
      // The premium pension sheet stops at 105; past that its last row repeats.
      ppArv.set(age, gridAt(ARV_PP, age < 106 ? age : 105, y));
    }

    // Taxes.
    if (y <= 1930 && historicalTaxRate) {
      // Missing from K_skatt, so assumed. NOTE: unreachable for every cohort the
      // workbook offers -- the earliest is 1930, whose first modelled age lands
      // in 1931.
      komSkatt.set(age, 0.0844);
      taxLimit1.set(age, 10_000);
      taxLimit2.set(age, 10 ** 9);
    } else if (y <= 1958 && historicalTaxRate) {
      komSkatt.set(age, taxRate("kommunalskatt", y));
      taxLimit1.set(age, 11_000);
      taxLimit2.set(age, 10 ** 9);
    } else if (historicalTaxRate) {
      komSkatt.set(age, taxRate("kommunalskatt", y));
      taxLimit1.set(age, data.skiktgrans1.at(y));
      taxLimit2.set(age, data.skiktgrans2.at(y));
    } else {
      komSkatt.set(age, context.kommunalskatt);
      if (y < 1958) {
        // Arbitrary, as the VBA says.
        taxLimit1.set(age, vbaInt(11_000 / 1.06 ** (1958 - y)));
        taxLimit2.set(age, 10 ** 9);
      } else {
        taxLimit1.set(age, data.skiktgrans1.at(y));
        taxLimit2.set(age, data.skiktgrans2.at(y));
      }
    }

    if (y < 2000 && historicalTaxRate) {
      begravavg.set(age, 0);
    } else if (historicalTaxRate) {
      begravavg.set(age, taxRate("begravningsavgift", y));
    } else {
      begravavg.set(age, context.begravningsavgift);
    }

    // Another year's tax rules applied throughout.
    if (context.rulesFromSkatt !== 0) {
      if ((y > context.rulesFromSkatt && context.rules === 0) || context.rules === 1) {
        taxLimit1.set(age, data.skiktgrans1.at(context.rulesFromSkatt));
        taxLimit2.set(age, data.skiktgrans2.at(context.rulesFromSkatt));
      }
      if (historicalTaxRate) {
        komSkatt.set(age, taxRate("kommunalskatt", context.rulesFromSkatt));
        begravavg.set(age, taxRate("begravningsavgift", context.rulesFromSkatt));
      }
    }

    // Taxable earned income.
    if (ownIncomeByAge !== undefined) {
      // The Egen inkomst sheet supplies a whole column, so an age it does not
      // mention reads as an empty cell, i.e. 0 -- the wage profile is not
      // consulted for it.
      const own = ownIncomeByAge.get(age);
      income.set(age, own?.income ?? 0);
      wage.set(age, own?.wage ?? 0);
    } else {
      // `wages` reads the withdrawal share for this age, which the VBA computes
      // at the top of the same function; so it is updated first here.
      uttagIp = withdrawalShare(age, profile, uttagIp);
      let amount = wages(
        age,
        y,
        wStart,
        input.monthlySalary * 12,
        profile,
        data,
        uttagIp,
        wTime,
        context.referenceYear,
        context.nominalWage,
      );

      // A salary changed from a given income year onwards.
      if (context.tlSpecYear > 1959 && amount > 0 && y >= context.tlSpecYear) {
        if (context.tlSpecial !== 1) amount = amount * context.tlSpecial;
      }
      income.set(age, amount);

      // On sickness or activity compensation from a given year: no more salary.
      if (context.satagare > 1960) {
        wage.set(age, context.satagare <= y ? 0 : amount);
      } else {
        wage.set(age, amount);
      }
    }
  }

  const vectors: SetupVectors = {
    startage,
    slutage,
    year,
    ibb,
    pbb,
    fpb,
    kpi,
    kpiJune,
    iindex,
    pindex,
    yieldFactor,
    rgk,
    mpgi,
    ipAvg,
    ppAvg,
    tpAvg,
    ipArv1,
    ipArv2,
    ppArv,
    komSkatt,
    begravavg,
    taxLimit1,
    taxLimit2,
    income,
    wage,
  };

  return {
    vectors,
    data,
    warnings,
    profile: {
      born,
      startage,
      slutage,
      par,
      tjpPar,
      defAr,
      wStart,
      income: input.monthlySalary * 12,
      avtal: input.scheme,
      civ: input.married ? 1 : 0,
      andelnya: andel(born),
      riktl,
      // `Rng_riktage`, by cohort, for the same reason as `riktl` above.
      riktalder: riktalderFor(born),
      modelYear,
      wTime,
    },
  };
}
