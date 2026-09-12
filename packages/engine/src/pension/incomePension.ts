/**
 * The income pension and premium pension: capital in, annuity out.
 *
 * Port of the second half of Pensionssystemet.bas -- `deltal`, `P_uttag`,
 * `IP_`, `ppkassa` and `pgb_barn`.
 *
 * `IP_` is the heart of the model. Called once per age, it rolls the income
 * pension balance forward through inheritance gains, indexation and management
 * costs, and once pension is drawn it converts the balance into an annuity.
 * Which of those five quantities it returns is chosen by its `Typ` argument --
 * the workbook's Brutto sheet calls it once per component, which is how the
 * fixture can check each one separately.
 */

import { vbaInt } from "../vba/math.js";
import type { DeltalTables } from "./deltal.js";

/** What `IP_` should return; the VBA passes this as its `Typ` argument. */
export const IpResult = {
  /** The pension paid out this year. */
  Pension: 0,
  /** Inheritance gains credited to the balance. */
  InheritanceGains: 1,
  /** Indexation credited to the balance. */
  Indexation: 2,
  /** Management costs deducted from the balance. */
  ManagementCost: 3,
  /** The closing balance at 31 December. */
  Balance: 4,
} as const;

export type IpResultKind = (typeof IpResult)[keyof typeof IpResult];

/** Cohorts before this have no income pension at all -- they are on ATP. */
export const FIRST_INCOME_PENSION_COHORT = 1938;

/** The norm the follow-up indexation divides out: 1.6%, but 0.996 in 2000. */
function foljsamhetsindex(year: number, index: number): number {
  return year === 2000 ? index / 0.996 : index / 1.016;
}

/** Withdrawal shares in force, which the VBA reads from named ranges. */
export interface WithdrawalShares {
  /** `UttagIP`: share of income pension drawn during a partial withdrawal. */
  readonly incomePension: number;
  /** `UttagPP`: share of premium pension drawn during a partial withdrawal. */
  readonly premiumPension: number;
}

/**
 * Whether a birthday part-way through the year pushes the relevant income year
 * on by one. The VBA computes this as `konst` in half a dozen places.
 */
function yearShift(born: number, age: number): number {
  return vbaInt(born + age + 1 / 1000) > vbaInt(born) + vbaInt(age) ? 1 : 0;
}

/**
 * Delningstal from the published tables, interpolated across the retirement
 * year. Mirrors `deltal` in Pensionssystemet.bas.
 *
 * Distinct from `fnDeltalIp` in deltal.ts: that one is the newer lookup with
 * the mortality override, this one indexes the published tables directly and is
 * what `ppkassa` calls.
 */
export function deltal(
  par: number,
  fodar: number,
  alder: number,
  tables: DeltalTables,
  kind: "income" | "premium",
  defArIn = 999,
): number {
  if (fodar <= 1937) return tables.oldRules(vbaInt(fodar), kind);

  let defAr = defArIn;
  if (par > defAr) defAr = par;
  if (defAr > 99) defAr = par;

  const konst = yearShift(fodar, par);
  const cohort = vbaInt(fodar);
  const at = (age: number) =>
    kind === "income" ? tables.incomePension(cohort, age) : tables.premiumPension(cohort, age);

  // Months already elapsed when pension is taken, rounded down to whole months.
  const month1 = 12 - vbaInt((par - vbaInt(par)) * 12);
  const month2 = 12 - vbaInt((defAr - vbaInt(defAr)) * 12);

  const lookupAge = vbaInt(alder - konst);
  let value: number;

  if (par <= 60) {
    // QUIRK: below 61 the value is extrapolated from the *premium pension*
    // table's first two ages, whichever pension is being asked about. Faithful
    // to the VBA, which reads fixed columns 28 and 29 here.
    const first = tables.premiumPension(cohort, 61);
    const second = tables.premiumPension(cohort, 62);
    value = (first - second) * (61 - par) + first;
  } else if (alder < vbaInt(par + konst)) {
    value = 0;
  } else if (alder === vbaInt(par + konst)) {
    value = at(lookupAge) * (month1 / 12) + at(lookupAge + 1) * ((12 - month1) / 12);
  } else if (par === defAr) {
    value = at(lookupAge);
  } else if (alder < vbaInt(defAr + 1)) {
    value = at(lookupAge);
  } else if (alder === vbaInt(defAr + 1)) {
    value = at(lookupAge) * (month2 / 12) + at(lookupAge + 1) * ((12 - month2) / 12);
  } else {
    value = at(lookupAge);
  }

  // Two decimals, with an exact .xx5 going down -- as everywhere delningstal
  // are rounded in this model.
  return vbaInt(value * 100 + 0.4999) / 100;
}

/** The share of pension being drawn at a given age. Mirrors `P_uttag`. */
export function pUttag(
  par: number,
  alder: number,
  shares: WithdrawalShares,
  defArIn = 999,
  incomePension = true,
): number {
  let defAr = defArIn;
  if (par > defAr) defAr = par;
  if (defAr > 99) defAr = par;

  const share = incomePension ? shares.incomePension : shares.premiumPension;
  if (alder < vbaInt(par)) return 0;
  if (alder < vbaInt(defAr)) return share;
  return 1;
}

/** Everything `IP_` computes in a year; the VBA returns one at a time. */
export interface IncomePensionYear {
  /** Pension paid out this year. */
  readonly pension: number;
  /** Inheritance gains credited. */
  readonly inheritanceGains: number;
  /** Indexation credited. */
  readonly indexation: number;
  /** Management cost deducted (negative). */
  readonly managementCost: number;
  /** Closing balance at 31 December. */
  readonly balance: number;
}

/**
 * Rolls the income pension forward one year.
 *
 * @param pratt    the pension right earned this year
 * @param arvsf1   inheritance-gain factor below riktålder
 * @param arvsf2   inheritance-gain factor at and above riktålder
 * @param kostf    what is left after management costs, as a factor
 * @param pbhIng   opening balance at 31 December last year
 * @param andel    the share of pension being drawn
 * @param pens     last year's pension
 * @param deltalValue the divisor in force
 * @param index    this year's indexation factor
 * @param bindex   balance index over income index; below 1 it trims the right
 */
export function incomePensionYear(
  year: number,
  par: number,
  born: number,
  pratt: number,
  arvsf1: number,
  arvsf2: number,
  kostf: number,
  pbhIng: number,
  andel: number,
  pens: number,
  deltalValue: number,
  index: number,
  defArIn = 999,
  marginal: 0 | 1 = 0,
  bindex = 0,
): IncomePensionYear {
  const empty: IncomePensionYear = {
    pension: 0,
    inheritanceGains: 0,
    indexation: 0,
    managementCost: 0,
    balance: 0,
  };
  if (born < FIRST_INCOME_PENSION_COHORT) return empty;

  let right = pratt;
  // From 2015 the pension right is trimmed by the balance/income index ratio.
  if (year > 2014 && bindex < 1 && bindex > 0) right *= bindex;

  const konst = yearShift(born, par);
  let defAr = defArIn;
  const konst2 = yearShift(born, defAr);

  const alder = year - vbaInt(born);
  let month = 12 - vbaInt(12 * (born + par - vbaInt(born + par)));

  if (par > defAr) defAr = par;
  if (defAr > 100) defAr = par;

  const finx = foljsamhetsindex(year, index);

  let arv = 0;
  let inx = 0;
  let kost = 0;
  let pbh = 0;
  let pension = 0;
  let share = andel;
  let underlag = 0;

  if (alder < vbaInt(par + konst)) {
    // Still earning: inheritance gains accrue on the whole balance.
    arv = (arvsf1 - 1) * pbhIng;
    arv = arv * arvsf2 + (pbhIng + right) * (arvsf2 - 1);
    inx = (arv + right + pbhIng) * (index - 1);
    kost = (inx + arv + right + pbhIng) * (kostf - 1);
    pbh = right + arv + inx + kost + pbhIng;
  } else if (alder === vbaInt(par + konst)) {
    // The year pension is first drawn, part-way through.
    arv = ((arvsf2 - 1) * pbhIng * (12 - month)) / 12 + right * (arvsf2 - 1);
    inx = (arv + right) * (index - 1) + (((index * 0.016 * pbhIng) / 1.016) * (12 - month)) / 12;
    kost = (inx + arv + right + pbhIng) * (kostf - 1);
    if (deltalValue > 0) pension = (pbhIng * share) / deltalValue;
    pbh = pbhIng * (1 - share) + (right + arv + inx + kost);
  } else if (alder > vbaInt(par + konst) && alder < vbaInt(defAr)) {
    // Drawing partially, still earning rights.
    arv = (arvsf2 - 1) * pbhIng + (arvsf2 - 1) * right;
    inx = (arv + right) * (index - 1) + (((index * 0.016 * pbhIng) / 1.016) * (12 - month)) / 12;
    kost = (inx + arv + right + pbhIng) * (kostf - 1);

    if (pbhIng > 0) {
      if (alder + 1 > vbaInt(par + konst)) month = 12;
      if (deltalValue > 0) {
        pension = ((pbhIng + pens * finx * deltalValue * (12 / month)) * share) / deltalValue;
      }
      underlag = pens * finx * deltalValue * (12 / month);
      if (marginal === 0) underlag = vbaInt(vbaInt(pens * finx + 0.49) * deltalValue + 0.49);
      pbh = (pbhIng + underlag) * (1 - share) + (right + arv + inx + kost);
    } else {
      pension = pens * finx * (12 / month) * deltalValue;
      pbh = 0;
    }
    month = 12;
  } else if (alder === vbaInt(defAr + konst2)) {
    // The year withdrawal becomes final and full.
    share = 1;
    arv = (arvsf2 - 1) * (pbhIng + right);
    inx = (arv + right) * (index - 1) + (((index * 0.016 * pbhIng) / 1.016) * (12 - month)) / 12;
    kost = (inx + arv + right + pbhIng) * (kostf - 1);
    if (deltalValue > 0) {
      pension = ((pbhIng + pens * finx * deltalValue * (12 / month)) * share) / deltalValue;
    }
    underlag = pens * finx * deltalValue;
    if (marginal === 0) underlag = vbaInt(vbaInt(pens * finx + 0.49) * deltalValue + 0.49);
    pbh = (pbhIng + underlag) * (1 - share) + (right + arv + inx + kost);
  } else if (alder > vbaInt(defAr + konst2)) {
    // Fully retired: the balance is re-derived from last year's pension.
    share = 1;
    if (alder > vbaInt(defAr + konst2 + 1)) month = 12;
    if (pbhIng > 0 || right > 0) {
      if (deltalValue > 0) {
        pension = ((pbhIng + pens * finx * deltalValue * (12 / month)) * share) / deltalValue;
      }
      underlag = pens * finx * deltalValue * (12 / month);
      if (marginal === 0) underlag = vbaInt(underlag + 0.49);
      pbh = pbhIng + underlag + (right + arv + inx + kost);
      if (deltalValue > 0) pension = pbh / deltalValue;
      pbh *= 1 - share;
    } else if (alder <= 65) {
      underlag = pens * finx * deltalValue * (12 / month);
      if (marginal === 0) underlag = vbaInt(vbaInt(pens * finx + 0.49) * deltalValue + 0.49);
      if (deltalValue > 0) pension = underlag / deltalValue;
    } else {
      pension = pens * finx;
    }
    month = 12;
  }

  if (pbh < 0) pbh = 0;
  if (marginal === 0) {
    pbh = vbaInt(pbh);
    // The annual pension is rounded to a whole krona per month.
    pension = vbaInt(pension / 12 + 0.5) * 12;
  }
  pension *= month / 12;

  return {
    pension,
    inheritanceGains: arv,
    indexation: inx,
    managementCost: kost,
    balance: pbh,
  };
}

/** `IP_` with its `Typ` argument, for call sites that mirror the VBA directly. */
export function ipResult(result: IncomePensionYear, typ: IpResultKind): number {
  switch (typ) {
    case IpResult.InheritanceGains:
      return result.inheritanceGains;
    case IpResult.Indexation:
      return result.indexation;
    case IpResult.ManagementCost:
      return result.managementCost;
    case IpResult.Balance:
      return result.balance;
    default:
      return result.pension;
  }
}

/**
 * The premium pension paid out in a year. Mirrors `ppkassa`.
 *
 * Simpler than the income pension: the balance is whatever the funds are worth,
 * and the annuity is that divided by the divisor.
 */
export function ppkassa(
  par: number,
  born: number,
  alder: number,
  pbhIn: number,
  tables: DeltalTables,
  shares: WithdrawalShares,
  defArIn = 999,
  andel = 1,
  marginal: 0 | 1 = 0,
): number {
  let defAr = defArIn;
  if (par > defAr) defAr = par;
  if (defAr < 61) defAr = par;
  if (defAr > 99) defAr = par;

  if (born < FIRST_INCOME_PENSION_COHORT) return 0;

  const konst = yearShift(born, par);
  const konst2 = yearShift(born, defAr);

  let share = andel;
  if (share > 1) share = 1;
  if (share < 0.1) return 0;

  const divisor =
    alder < 101 ? deltal(par, born, alder, tables, "premium", defAr) : 2;

  const pbh = pbhIn < 0 ? 0 : pbhIn;
  const uttagPp = shares.premiumPension;

  let amount = 0;
  let month = 0;

  if (alder < vbaInt(par + konst)) {
    amount = 0;
  } else if (alder === vbaInt(par + konst)) {
    amount = (share * pbh) / divisor;
    month = 12 - vbaInt(12 * (born + par - vbaInt(born + par)));
    if (alder === vbaInt(defAr + konst2)) {
      month += (12 - 12 * (born + defAr - vbaInt(born + defAr + 1 / 1000))) * (1 - uttagPp);
    }
  } else if (alder < vbaInt(defAr + konst)) {
    amount = pbh / divisor;
    month = 12 * uttagPp;
  } else if (alder === vbaInt(defAr + konst)) {
    amount = pbh / divisor;
    month = 12 * uttagPp;
    month += (12 - 12 * (born + defAr - vbaInt(born + defAr + 1 / 1000))) * (1 - uttagPp);
  } else {
    amount = pbh / divisor;
    month = 12;
  }

  if (month > 12) month = 12;
  if (month < 0) month = 0;

  // Rounded to a whole krona per month, then multiplied back up.
  amount = marginal === 0 ? vbaInt(amount / 12 + 0.5) : amount / 12;
  return amount * month;
}

/**
 * Pension-qualifying amount for years spent caring for a small child.
 *
 * Whichever of three comparisons is most favourable: the drop in the parent's
 * own income, three quarters of the average income, or one income base amount.
 */
export function pgbBarn(
  ar: number,
  jink: number,
  uink: number,
  barn: number,
  parent: number,
  medel: number,
  ibb = 52_100,
  marginal: 0 | 1 = 0,
  rikt = 65,
): number {
  let amount = 0;
  if (barn > 1959 && parent < rikt) {
    if (ar === barn || (ar < barn + 4 && ar > barn)) {
      amount = jink - uink;
      if (0.75 * medel - uink > amount) amount = medel * 0.75 - uink;
      if (ibb > amount) amount = ibb;
      if (marginal === 0) amount = vbaInt((amount + 49) / 100) * 100;
    }
  }
  return amount;
}
