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

  // ---- Retirement ages and withdrawal -------------------------------------

  /**
   * `TJP_PAR`: the age the occupational pension is drawn from. 0 follows the
   * public pension age.
   *
   * The Adv_settings cell is `=IF(ISBLANK(...), ..., ...)`, so a blank entry
   * falls back to the retirement age -- which is why the extracted value reads
   * 66, the shipped retirement age, rather than a constant. `setup.ts` resolves
   * 0 the same way.
   */
  readonly tjpPar: number;
  /**
   * `rng_def_ar`: the age at which withdrawal becomes final and full. Below the
   * retirement age, or 0, means the same as it -- which is what `startsetup`
   * does with the blank cell the workbook ships.
   */
  readonly defAr: number;
  /** `UttagIP`: the share of income pension drawn during partial withdrawal. */
  readonly uttagIp: number;
  /** `UttagPP`: the same for premium pension. */
  readonly uttagPp: number;
  /**
   * `rngLönPartUttag`: a number fixes work at that share while pension is drawn
   * partially; any other value means work follows the withdrawal.
   */
  readonly workDuringPartialWithdrawal: number | string;
  /** `Wage_profil`: 0 a straight wage path, 1-4 an age profile. */
  readonly wageProfile: number;
  /**
   * `w_time`: the age the entered salary refers to. 0 derives it.
   *
   * The Adv_settings cell is a subtraction of two named ranges. Its hint reads
   * "Inkomståret 2025" and its shipped value is 66, which is exactly
   * `w_ref (2025) - born (1959)` -- so it is the typfall's age in the reference
   * year. The operands are pinned by that arithmetic rather than read directly,
   * because the LibreOffice conversion the extractor uses cannot resolve
   * defined names inside formulas.
   */
  readonly wTime: number;
  /** `Nominal`: the salary is given in nominal money rather than fixed prices. */
  readonly nominalWage: boolean;
  /** `wStartYear_AdvSettings`: the model's first income age, 1 in the shipped workbook. */
  readonly modelStartAge: number;

  // ---- Household -----------------------------------------------------------

  /**
   * `Hyra`: housing cost per month in the reference year. The sheet holds
   * `=6300 + 1200 * x`, which comes to the shipped 6 300.
   */
  readonly hyra: number;
  /** `rng_Formogenhet`: wealth beyond the home. */
  readonly formogenhet: number;
  /** `rng_Makens_inkomst`: the spouse's annual income. */
  readonly makensInkomst: number;
  /** `rng_Make_Bald`: the spouse's year of birth; 0 takes the typfall's own. */
  readonly makeBorn: number;
  /** `rng_forstidM`: the spouse's försäkringstid. */
  readonly makeInsuranceYears: number;
  /** `rng_Kapital_pens`: gross capital income, assumed to start at retirement. */
  readonly kapital: number;
  /** `Rng_Ansokt`: 1 when the household applies for the benefits it is entitled to. */
  readonly ansokt: number;
  /** `rng_Född_Barn1`..`4`: the children's years of birth, 0 for an empty slot. */
  readonly childBirthYears: readonly [number, number, number, number];

  // ---- Private saving ------------------------------------------------------

  /** `IPS_Monthly`: monthly private pension saving. */
  readonly ipsMonthly: number;
  /** `IPS_start`: the year that saving starts. The sheet holds `=YEAR(TODAY())`. */
  readonly ipsStart: number;
  /** `rng_Kapitalförsäkring`: 0 IPS, 1 kapitalförsäkring, 2 ISK. */
  readonly privateSavingKind: number;
  /** `rng_Fondförsäkring`: fund insurance rather than traditional. */
  readonly fundInsurance: number;

  // ---- Taxes and fees ------------------------------------------------------

  /** `rng_Kommunalskatt`: a fixed municipal rate; below 0.1 uses the historical average. */
  readonly kommunalskatt: number;
  /** `rng_Begravningsavgift`: the burial fee and any church fee, used with the above. */
  readonly begravningsavgift: number;
  /** `fack`: union fee per month. */
  readonly fack: number;
  /** `akasseavg`: unemployment insurance fee per month. */
  readonly akasseavg: number;
  /** `satagare`: on sickness or activity compensation since this year; 0 for never. */
  readonly satagare: number;

  // ---- Opening balances ----------------------------------------------------

  /** `rng_PBHYear`: the income year an opening balance is injected for; 0 for none. */
  readonly pbhYear: number;
  /** `rng_PBH_IP`: that year's income pension balance. */
  readonly pbhIp: number;
  /** `rng_PBH_PP`: premium pension. */
  readonly pbhPp: number;
  /** `rng_PBH_tjp`: occupational pension. */
  readonly pbhTjp: number;
  /** `rng_PBH_IPS`: private saving. */
  readonly pbhIps: number;

  // ---- Rule-year experiments ----------------------------------------------

  /** `rng_Boundray_Year` (`Iyear`): the year from which expenditure rules follow earnings. */
  readonly boundaryYear: number;
  /** `Tl_spec_y`: an income year from which the salary is multiplied by `tlSpecial`. */
  readonly tlSpecYear: number;
  /** `TL_special`: that multiplier. */
  readonly tlSpecial: number;
  /** `rng_Sista_PensRatt`: 1 credits a final year of pension rights at retirement. */
  readonly lastPensionRight: number;
  /** `rng_alt_last_pratt`: years after retirement for an alternative final right. */
  readonly altLastPensionRight: number;
  /** `rngPens_Inflation`: 1 puts retirement in the same year as the final salary. */
  readonly pensionSameYearAsFinalSalary: number;

  // ---- Presentation --------------------------------------------------------

  /**
   * `rng_Bara_fastapriser`: 1 fixed prices, 0 expressed in the reference year's
   * wage level, -1 nominal.
   */
  readonly priceBasis: number;
  /**
   * `rng_tabell2_startAge`: the age table 2 and the figures start at. 0 derives
   * it as `Int(retirement age - 10)`, which is what the sheet's `=INT(x - 10)`
   * does and what the shipped 56 comes to at a retirement age of 66.
   */
  readonly table2StartAge: number;
  /** `rng_Chart_Earning_factor`: 1 shows annual amounts, 12 monthly. */
  readonly chartEarningFactor: number;
  /** `rng_discount`: the rate the life-income sums are discounted at. */
  readonly discountRate: number;
  /**
   * `Modell_year`: the workbook's own current year, which the yield series
   * switches on. 0 derives it as `w_ref + 1`: the reference year is
   * `=YEAR(NOW()) - 1`, so the two differ by exactly one.
   */
  readonly modelYear: number;

  // ---- Advanced-mode extras, off by default -------------------------------

  /** `Risk`: the standard deviation of the return. 0 makes the yield deterministic. */
  readonly risk: number;
  /** `lognormal`: 1 draws the return lognormally, anything else normally. */
  readonly lognormal: number;
  /** `Wealth`: 1 shows the pension wealth and respektavstånd box. Not yet ported. */
  readonly wealth: number;
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

    // 0, not the extracted 66: that value is the formula's fallback to the
    // shipped retirement age, and following it is the behaviour, not the 66.
    tjpPar: 0,
    defAr: workbookDefault("rng_def_ar", 0),
    uttagIp: 1,
    uttagPp: 1,
    workDuringPartialWithdrawal: "Arbetar deltid",
    wageProfile: workbookDefault("Wage_profil", 0),
    wTime: 0,
    nominalWage: false,
    modelStartAge: workbookDefault("wStartYear_AdvSettings", 1),

    hyra: workbookDefault("Hyra", 6300),
    formogenhet: workbookDefault("rng_Formogenhet", 0),
    makensInkomst: workbookDefault("rng_Makens_inkomst", 0),
    makeBorn: 0,
    makeInsuranceYears: workbookDefault("rng_forstidM", 40),
    kapital: workbookDefault("rng_Kapital_pens", 0),
    ansokt: workbookDefault("Rng_Ansokt", 1),
    childBirthYears: [
      workbookDefault("rng_Född_Barn1", 0),
      workbookDefault("rng_Född_Barn2", 0),
      workbookDefault("rng_Född_Barn3", 0),
      workbookDefault("rng_Född_Barn4", 0),
    ],

    ipsMonthly: workbookDefault("IPS_Monthly", 0),
    ipsStart: workbookDefault("IPS_start", 2026),
    privateSavingKind: workbookDefault("rng_Kapitalförsäkring", 0),
    fundInsurance: workbookDefault("rng_Fondförsäkring", 1),

    kommunalskatt: workbookDefault("rng_Kommunalskatt", 0),
    begravningsavgift: workbookDefault("rng_Begravningsavgift", 0),
    fack: workbookDefault("fack", 0),
    akasseavg: workbookDefault("akasseavg", 0),
    satagare: workbookDefault("satagare", 0),

    pbhYear: workbookDefault("rng_PBHYear", 0),
    pbhIp: workbookDefault("rng_PBH_IP", 0),
    pbhPp: workbookDefault("rng_PBH_PP", 0),
    pbhTjp: workbookDefault("rng_PBH_tjp", 0),
    pbhIps: workbookDefault("rng_PBH_IPS", 0),

    boundaryYear: workbookDefault("rng_Boundray_Year", 0),
    tlSpecYear: workbookDefault("Tl_spec_y", 0),
    tlSpecial: workbookDefault("TL_special", 1),
    lastPensionRight: workbookDefault("rng_Sista_PensRatt", 1),
    altLastPensionRight: workbookDefault("rng_alt_last_pratt", 0),
    pensionSameYearAsFinalSalary: workbookDefault("rngPens_Inflation", 0),

    priceBasis: workbookDefault("rng_Bara_fastapriser", 1),
    table2StartAge: 0,
    chartEarningFactor: 1,
    discountRate: workbookDefault("rng_discount", 0.05),
    modelYear: 0,

    risk: workbookDefault("Risk", 0),
    lognormal: workbookDefault("lognormal", 1),
    wealth: workbookDefault("Wealth", 0),
  };
  return Object.freeze({ ...base, ...overrides });
}
