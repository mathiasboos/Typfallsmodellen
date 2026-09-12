/**
 * Occupational pension: the eight collective agreements the model offers.
 *
 * `premiumFor` is the dispatcher, mirroring the `Select Case avtal` in Mcalc
 * that chooses which agreement's premium function to call.
 */

import type { ModelContext } from "../model/context.js";
import { Kapan, PA_indiv, tlPA16 } from "./statlig.js";
import { SAF_LO } from "./safLo.js";
import { tlakap_kr, tlkap_kl } from "./kommunal.js";
import { tlITP1, tlITP2A } from "./itp.js";
import { Scheme } from "./types.js";
import type { SchemeContext, SchemeId } from "./types.js";

export * from "./types.js";
export * from "./itp.js";
export * from "./safLo.js";
export * from "./kommunal.js";
export * from "./statlig.js";
export * from "./tjpkassa.js";

/**
 * The occupational pension premium earned in one year.
 *
 * @param scheme  the agreement, numbered as the Start sheet's drop-down numbers it
 * @param alder   age at 31 December
 * @param wage    earnings for the year
 * @param ibb     income base amount
 * @param tjpPar  occupational pension retirement age
 */
export function premiumFor(
  scheme: SchemeId,
  alder: number,
  wage: number,
  ibb: number,
  tjpPar: number,
  year: number,
  context: SchemeContext,
  model: ModelContext,
): number {
  switch (scheme) {
    case Scheme.None:
      return 0;
    case Scheme.Itp1:
      return tlITP1(alder, wage, ibb, tjpPar, context);
    case Scheme.Itp2:
      // Only the ITPK premium accrues here; the benefit part is added by FTJP.
      return tlITP2A(alder, wage, tjpPar, context);
    case Scheme.SafLo:
      return SAF_LO(alder, wage, ibb, tjpPar, year, context);
    case Scheme.KapKl:
      return tlkap_kl(alder, wage, ibb, tjpPar, year, context);
    case Scheme.AkapKr:
      return tlakap_kr(alder, wage, ibb, tjpPar, year, context, model);
    case Scheme.Pa16Avd2:
      // Two premiums paid side by side, which Mcalc sums.
      return (
        Kapan(year, context.born, ibb, wage, context, tjpPar) +
        PA_indiv(year, context.born, ibb, wage, context, tjpPar)
      );
    case Scheme.Pa16Avd1:
      return tlPA16(alder, wage, ibb, tjpPar, year, context);
    default: {
      const exhaustive: never = scheme;
      throw new RangeError(`unknown occupational pension scheme: ${String(exhaustive)}`);
    }
  }
}
