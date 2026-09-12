/**
 * The settings the rule functions read, in place of Excel named ranges.
 *
 * The VBA reaches out to `Application.Range("...")` from inside the rule
 * modules -- about forty distinct names. Threading one frozen context object
 * through instead keeps those functions pure and testable, and puts every
 * setting the model has in one typed place.
 *
 * Defaults come from the workbook itself: the `Adv_settings` sheet carries its
 * own variable name in column 9 and its "normal" value in column 8, both
 * extracted into packages/data/options.json. So these are the model's defaults,
 * not a transcription of them.
 */

import optionsJson from "../../../data/options.json" with { type: "json" };

/** One advanced setting as the workbook documents it. */
interface AdvancedSettingSpec {
  readonly name: string;
  readonly row: number;
  readonly label: string;
  readonly hint: string;
  readonly default: unknown;
  readonly current: unknown;
}

const SPECS = optionsJson.advancedSettings as readonly AdvancedSettingSpec[];

/** The workbook's default for a named setting, or `fallback` if it has none. */
export function workbookDefault(name: string, fallback: number): number {
  const spec = SPECS.find((s) => s.name.toLowerCase() === name.toLowerCase());
  const value = spec?.default;
  return typeof value === "number" ? value : typeof value === "boolean" ? (value ? 1 : 0) : fallback;
}

/**
 * Settings read by the pension, tax and benefit rules.
 *
 * Named after the workbook's variables so each one can be traced back. Where a
 * setting is a policy experiment rather than current law, the comment says what
 * switches it on.
 */
export interface ModelContext {
  /** `marginal`: 0 applies the real rules' rounding, 1 removes all of it. */
  readonly marginal: 0 | 1;

  /** `Soc_tak`: raise the contribution ceiling from this year. Inert at or below 1999. */
  readonly socTak: number;
  /** `Social_avg`: raise the contribution from this year. Inert at or below 1999. */
  readonly socialAvg: number;

  /** `Rules`: 0 changes rules from a given year, 1 applies that year's rules throughout. */
  readonly rules: 0 | 1;
  /** `RulesfromUtg`: expenditure rules fixed from this year. 0 means current rules per year. */
  readonly rulesFromUtg: number;
  /** `RulesfromSkatt`: tax rules fixed from this year. 0 means current rules per year. */
  readonly rulesFromSkatt: number;

  /** `rng_Avkastning_fondavgifter`: true when the return is already net of fund fees. */
  readonly returnsNetOfFees: boolean;
  /** `rng_Avkastning_val`: 1 the chosen real return throughout, 2 PPM history, 3 AP7. */
  readonly returnBasis: number;
  /** `rng_Forenklad_berakning`: 1 uses the faster, simpler premium pension formula. */
  readonly simplifiedPremiumPension: boolean;

  /** `rng_Forsakringstid_vid_65`: years of residence, which caps garantipension below 40. */
  readonly insuranceYears: number;
  /** `rng_nypropp`: 1 includes tax proposals from the latest budget bill. */
  readonly newBudgetProposals: boolean;

  /** `rngDelnIPMort` / `rngDelnPPMort`: cohorts from here take delningstal from the
   *  model's own mortality calculation rather than the published table. */
  readonly deltalFromMortalityIp: number;
  readonly deltalFromMortalityPp: number;

  /** `Mortality`: 1 computes delningstal from the model's own tables rather than SCB's. */
  readonly ownDeltal: boolean;
  /** `Average_Earning`: how many of the final years the final salary averages. */
  readonly finalSalaryYears: number;
  /** `w_ref`: the year whose price level results are expressed in. */
  readonly referenceYear: number;
  /** `rng_Senaste_Index_Framskrivning`: 2 uses inkomstindex for every year. */
  readonly latestIndexBasis: number;
  /** `Rgk`: the Riksgälden rate applied while contributions await assessment. */
  readonly rgk: number;
  /** `rng_Arvsvinster_TJP`: 0 simulates survivor cover, forgoing inheritance gains. */
  readonly occupationalInheritanceGains: number;
  /** `rng_FlexPens`: extra flexpension premium for ITP1 and SAF-LO from 2014. */
  readonly flexPension: number;
  /**
   * `Rng_riktage`: riktålder for the current cohort, read from Nyckeltal.
   * AKAP-KR ties its higher premium to the LAS age, which is this plus three.
   */
  readonly riktage: number;
  /**
   * `RNG_dela`: how a cohabiting household's benefits are reported -- 1 the
   * household's combined amount, 2 the individual's, 3 an equal split between
   * the spouses. At the default 2 every spouse term in BTP and SBTP is inert.
   */
  readonly dela: number;
  /** `rng_Temp_Tjp_Uttag`: years of temporary occupational withdrawal; 0 is lifelong. */
  readonly tempTjpUttag: number;
  /** `rng_Temp_IPS_Uttag`: the same for private saving. */
  readonly tempIpsUttag: number;
  /**
   * `rng_Ddelat`: 1 corrects the occupational pension divisor for each
   * agreement's own assumed interest rate and life expectancy. Off by default,
   * which makes that correction a pass-through.
   */
  readonly adjustOccupationalDivisor: boolean;
}

/**
 * The model's normal settings -- what `Använd normala inställningar` restores.
 *
 * `socTak` and `socialAvg` are policy-experiment switches with no entry on the
 * Adv_settings sheet; both are guarded by `> 1999` in the VBA, so 0 leaves them
 * inert, which is how the shipped workbook behaves.
 *
 * `riktage` defaults to the value the shipped workbook carries. It properly
 * varies by cohort, and Mcalc will set it from `riktage(year)` once the main
 * loop lands.
 */
export function defaultContext(overrides: Partial<ModelContext> = {}): ModelContext {
  const base: ModelContext = {
    marginal: workbookDefault("marginal", 0) === 1 ? 1 : 0,
    socTak: 0,
    socialAvg: 0,
    rules: workbookDefault("Rules", 0) === 1 ? 1 : 0,
    rulesFromUtg: workbookDefault("RulesFromUtg", 0),
    rulesFromSkatt: workbookDefault("RulesFromSkatt", 0),
    returnsNetOfFees: workbookDefault("rng_Avkastning_fondavgifter", 1) === 1,
    returnBasis: workbookDefault("rng_Avkastning_val", 2),
    simplifiedPremiumPension: workbookDefault("rng_Förenklad_beräkning", 0) === 1,
    insuranceYears: workbookDefault("rng_Försäkringstid_vid_65", 40),
    newBudgetProposals: workbookDefault("rng_nypropp", 0) === 1,
    deltalFromMortalityIp: workbookDefault("rngDelnIPMort", 1958),
    deltalFromMortalityPp: workbookDefault("rngDelnPPMort", 1958),
    ownDeltal: workbookDefault("Mortality", 0) === 1,
    finalSalaryYears: workbookDefault("Average_Earning", 5),
    referenceYear: workbookDefault("w_ref", 2025),
    latestIndexBasis: workbookDefault("rng_Senaste_Index_Framskrivning", 1),
    rgk: workbookDefault("Rgk", 0.01),
    occupationalInheritanceGains: workbookDefault("rng_Arvsvinster_TJP", 1),
    flexPension: workbookDefault("rng_FlexPens", 0),
    riktage: 66,
    dela: workbookDefault("RNG_dela", 2),
    tempTjpUttag: workbookDefault("rng_Temp_Tjp_Uttag", 0),
    tempIpsUttag: workbookDefault("rng_Temp_IPS_Uttag", 0),
    adjustOccupationalDivisor: workbookDefault("Rng_ddelat", 0) === 1,
  };
  return Object.freeze({ ...base, ...overrides });
}
