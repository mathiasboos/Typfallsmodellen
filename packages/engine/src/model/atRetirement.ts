/**
 * The second tax and benefits pass, at the retirement age.
 *
 * Port of VBA_go.bas 2607-2800, the `Last_pratt > 0` blocks that run after the
 * age loop and after `creditLastPensionRight`. They produce the last three rows
 * of Table 1 -- pension after tax, benefits, and disposable income at
 * retirement -- and they exist because the final year's pension right changes
 * the gross, so the tax and the housing supplement the loop worked out at that
 * age no longer match it.
 *
 * It is not the loop's pass repeated. It applies pensioner rules: no cost
 * deductions, no pension contribution, no jobbskatteavdrag, capital income
 * included, and no social assistance. It also reads several of the loop's own
 * local variables where the original left them in scope -- see
 * `RunState.leftovers` -- and the differences from the loop's arguments are
 * marked one by one below.
 *
 * When `rng_Sista_PensRatt` is 0 this does not run, and Table 1 reports what
 * the loop computed.
 */

import { BTP, SBTP, btp_sbtp } from "../bidrag/index.js";
import { FAared, PublicAvg, avdragxx, statlig } from "../skatt/index.js";
import { vbaInt, wsMax, wsMin } from "../vba/math.js";
import type { Run } from "./mcalc.js";
import { runVectors } from "./state.js";

/** What the recomputation leaves behind, for the three Table 1 rows. */
export interface AtRetirement {
  /** `Netto(Int(PAR))`: the pension after tax. */
  readonly netto: number;
  /** `Bidrag(Int(PAR))`: benefits, without the social assistance the loop adds. */
  readonly bidrag: number;
  /** `IndDisp(Int(PAR))`: disposable income. */
  readonly disposable: number;
}

/**
 * The tax rule year this pass uses.
 *
 * NOTE: not `ruleYears`. The original writes the test out again here as
 * `RulesfromSkatt = 0 Or year_(PAR) < RulesfromSkatt`, without the `Rules = 1`
 * arm the loop's version has, so the two disagree when rules are pinned to a
 * single year *and* that year is in the future. Both come to `year` under the
 * shipped settings, where `RulesfromSkatt` is 0.
 */
function skattYear(year: number, rulesFromSkatt: number): number {
  return rulesFromSkatt === 0 || year < rulesFromSkatt ? year : rulesFromSkatt;
}

/** The same again for the expenditure rules (VBA_go.bas:2727). */
function utgYear(year: number, rulesFromUtg: number): number {
  return rulesFromUtg === 0 || year < rulesFromUtg ? year : rulesFromUtg;
}

/** Tax on the pension at the retirement age (VBA_go.bas 2607-2666). */
function taxAtRetirement(run: Run): { netto: number; grundavdrag: number; kapskatt: number } {
  const { v, s, p, context } = run;
  const par = vbaInt(p.par);
  const year = v.year.get(par);
  const marginal = context.marginal;
  const { leftovers } = s;

  // No cost deductions: the comment says "inga kostnadsavdrag som pensionar".
  let ctxfvi = s.brutto.get(par);
  if (marginal === 0) ctxfvi = vbaInt(ctxfvi / 100) * 100;

  // NOTE: the loop passes 2100 for `wyear` (VBA_go.bas:1773); this passes
  // `Iyear`. And the tax year is `year_(mini(age, 100))` -- the loop counter as
  // VBA left it, one past `slutage` -- not this age's year.
  const skyearFromLoopAge = v.year.getOrZero(wsMin(leftovers.age, 100));
  const grundavdrag = avdragxx(
    ctxfvi, v.pbb.get(par), marginal, par, year, run.iyear, v.ibb.get(par), s.kvoten,
    skyearFromLoopAge, leftovers.gage,
  );

  // "Ingen lon antas" -- no salary, so no contribution.
  const pensionavgift = 0;
  const cbefvi = ctxfvi - grundavdrag;

  let kinkskatt = cbefvi * v.komSkatt.get(par);
  let kyrkskatt = cbefvi * v.begravavg.get(par);
  if (marginal === 0) {
    kinkskatt = vbaInt(kinkskatt);
    kyrkskatt = vbaInt(kyrkskatt);
  }

  const skyear = skattYear(year, context.rulesFromSkatt);
  let statskatt =
    statlig(cbefvi, v.taxLimit1.get(par), v.taxLimit2.get(par), marginal) +
    // NOTE: `age` here is the loop's leftover counter, which only has to clear
    // 18 for the public service fee to apply, so it never changes the answer.
    PublicAvg(cbefvi, p.born, runVectors(v), 0.01, leftovers.age, marginal, skyear);

  // Capital income, unconditionally -- the loop only counts it from `par`.
  const kapskatt =
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
  if (marginal === 0) statskatt = vbaInt(statskatt);

  // QUIRK: `pensredukt = rakassa` (VBA_go.bas:2652). The pension contribution
  // reduction is the a-kassa reduction the loop last computed, at `slutage`.
  // The line that would have computed it properly is commented out beside it.
  const pensredukt = leftovers.rakassa;
  const jobbavdrag = 0;

  let faavdrag = FAared(cbefvi, skyear, marginal, v.pbb.get(par));
  // NOTE: `SAavdrag` is the loop's leftover too, and the cap subtracts
  // `FAavdrag` from itself, as in the loop.
  if (kinkskatt - leftovers.saavdrag - faavdrag - jobbavdrag < faavdrag) {
    faavdrag = kinkskatt - leftovers.saavdrag - faavdrag - jobbavdrag;
  }

  // NOTE: `kapital` is added to the gross here but `SAavdrag` is not subtracted
  // from the tax, though it was in the loop. Kept as written.
  const netto =
    s.brutto.get(par) +
    context.kapital -
    wsMax(kinkskatt + kyrkskatt + statskatt + pensionavgift - pensredukt - jobbavdrag - faavdrag, 0);

  return { netto, grundavdrag, kapskatt };
}

/** Housing supplement and aldreforsorjningsstod at retirement (VBA_go.bas 2684-2752). */
function benefitsAtRetirement(
  run: Run,
  grundavdrag: number,
  kapskatt: number,
): { bidrag: number; bidragovr: number } {
  const { v, s, p, context } = run;
  const par = vbaInt(p.par);
  const year = v.year.get(par);
  const marginal = context.marginal;
  const civ = p.civ;
  const { leftovers } = s;

  // The child allowance and the housing allowance are *not* recomputed -- the
  // blocks that would are commented out in the original, with the note that
  // they are unaffected by the final pension right. So the loop's values stand.
  const barnbidrag = leftovers.barnbidrag;
  const bostadsbidrag = leftovers.bostadsbidrag;
  let bidragovr = leftovers.bidragovr;

  let bostadstillagg = 0;

  // A narrower gate than the loop's: both pensions have to be drawn in full.
  if (p.par >= p.riktalder && s.uttagIp === 1 && s.uttagPp === 1) {
    // NOTE: pbb over pbb, where the loop's version of this line uses IBB over
    // IBB -- and `kvoten` is not updated alongside it here.
    if (run.iyear > 0 && run.iyear <= year) {
      s.maxhyra = v.pbb.get(par) / v.pbb.getOrZero(run.iyear - vbaInt(p.born));
    }

    const rules = utgYear(year, context.rulesFromUtg);
    const ansoker = context.ansokt;
    const scale = 12 / run.pmonth;
    const iyearIbb = v.ibb.getOrZero(run.iyear - vbaInt(p.born));

    // NOTE: not multiplied by `uttagIP` the way the loop's BTP is, and the
    // `12 / pmonth` scaling is unconditional rather than only in the final
    // withdrawal year. The wage argument is written `0 * Wage_(Int(PAR))`.
    bostadstillagg = BTP(
      (s.brutto.get(par) - v.wage.get(par)) * scale - kapskatt - s.ptillagg.get(par),
      leftovers.makaInk * scale,
      12 * leftovers.hyraT,
      civ,
      v.pbb.get(par),
      context,
      leftovers.ap,
      leftovers.apm,
      context.formogenhet,
      0,
      0,
      marginal,
      s.maxhyra,
      rules,
      run.iyear,
      v.ibb.get(par),
      s.kvoten,
      par,
      s.tjp.get(par) * scale,
      leftovers.tjpm * scale,
      s.garp.get(par) * scale,
      s.garp.get(par) * scale,
      iyearIbb,
      s.mpension / 12,
      run.pmonth,
      context.insuranceYears,
      v.iindex.get(par),
      s.uttagIp,
      ansoker,
    );

    // QUIRK: `hyra_t`, not `12 * hyra_t`. Every other SBTP call in the model
    // passes the annual rent (VBA_go.bas:1996 and :2008); this one passes the
    // monthly figure, so aldreforsorjningsstod is measured against a rent a
    // twelfth of the real one. Kept as written.
    const sbostadstillaggRaw = SBTP(
      s.brutto.get(par) + context.kapital - kapskatt - s.ptillagg.get(par),
      leftovers.hyraT,
      civ,
      bostadstillagg + bostadsbidrag,
      grundavdrag,
      {
        dela: context.dela,
        rulesFromUtg: context.rulesFromUtg,
        born: p.born,
        age: par,
        slutage: v.slutage,
        Iyear: run.iyear,
        vectors: runVectors(v),
      },
      v.komSkatt.get(par),
      leftovers.ap,
      context.formogenhet,
      v.pbb.get(par),
      s.maxhyra,
      rules,
      run.iyear,
      v.ibb.get(par),
      s.kvoten,
      par,
      0,
      leftovers.makaInk,
      marginal,
      -99,
      run.forstid,
    );

    let sbostadstillagg = sbostadstillaggRaw;
    if (context.ansokt === 9) {
      bidragovr = bostadsbidrag + bostadstillagg;
      bostadstillagg = 0;
    }
    // NOTE: only the aldreforsorjningsstod is zeroed when nothing is applied
    // for. The loop zeroes the housing supplement and the housing allowance too.
    if (ansoker === 0) sbostadstillagg = 0;

    // NOTE: `year_(mini(age, slutage))` -- the loop's leftover counter again,
    // which clamps to `slutage`, not this age's year.
    bostadstillagg = btp_sbtp(
      bostadstillagg,
      sbostadstillagg,
      marginal,
      v.year.get(wsMin(leftovers.age, v.slutage)),
    );
  }

  return { bidrag: barnbidrag + bostadsbidrag + bostadstillagg, bidragovr };
}

/**
 * Recomputes tax, benefits and disposable income at the retirement age.
 *
 * Writes them back into the run state as the original does, so anything else
 * reading `Netto(PAR)` sees the same values, and returns them for Table 1.
 */
export function recomputeAtRetirement(run: Run): AtRetirement {
  const { s, p } = run;
  const par = vbaInt(p.par);

  const { netto, grundavdrag, kapskatt } = taxAtRetirement(run);
  s.netto.set(par, netto);

  const { bidrag, bidragovr } = benefitsAtRetirement(run, grundavdrag, kapskatt);
  s.bidrag.set(par, bidrag);

  // NOTE: no floor at zero here, where the loop clamps `Bidrag` to 0. And no
  // social assistance: the loop's `bist25` block has no counterpart.
  const disposable = netto + bidrag + bidragovr + s.pps.get(par);
  s.indDisp.set(par, disposable);

  return { netto, bidrag, disposable };
}
