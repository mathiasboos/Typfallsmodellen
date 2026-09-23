/**
 * Avancerat läge: the Adv_settings the web port exposes.
 *
 * The workbook splits its inputs in two -- the Start sheet holds the facts
 * about the person, `Adv_settings` holds everything about how the model behaves
 * -- and `packages/engine/src/model/input.ts` keeps that split: `TypfallInput`
 * is the Start sheet, `ModelContext` is Adv_settings. Normal mode runs on the
 * first alone, which is why `main.ts` could get this far passing an empty
 * settings map. This panel is the second.
 *
 * Thirty of the sheet's seventy-six rows are here, grouped as sections 3.2
 * to 3.8 of the user manual group them (plus 3.7's own "Partiellt uttag" half --
 * its "Barnår" half, `rng_Född_Barn1..4`, lives in pgb.ts instead, alongside
 * the sheet's other three pension-qualifying-amount sources). The rest are
 * left out on purpose: some
 * are Excel's own business (`Visa_process`, `rngTurboMode`, `Verbose`), some
 * feed a sheet this port does not have (`Rng_belopp12`, `Alt_p_age`,
 * `Rng_compareTo` are Mikrosim's), some are not ported (`Wealth` and the
 * respektavstånd box, `rng_Syntetisk`, `Scenario`), and the remainder are
 * policy experiments -- pinning a year's tax rules, removing the rounding from
 * the rule system -- that want a more careful UI than a number box. Adding any
 * of them is a row in the table below.
 *
 * **The labels are written here, in both languages, and that is a deliberate
 * departure from form.ts's "none of it is retyped here".** options.json does
 * carry a Swedish `label` and `hint` per row, but the workbook never translated
 * this sheet: 59 of its 90 SysLang rows have an empty English string. What
 * Swedish it does carry is sheet shorthand -- "Marginal, avrundningar mm",
 * "(1)-> (dvs ingen förändring...)" -- written for whoever maintains the
 * workbook. The manual's own prose is the better source, so each setting below
 * quotes the sheet's wording in a comment and carries its row number, and the
 * mapping stays checkable by eye. Values are still never retyped: every default
 * comes from `defaultContext()`, which reads options.json -- except `tjpPar`,
 * `uttagIp` and `uttagPp`, whose row on the sheet does not follow the sheet's
 * usual name-in-column-9 layout `extract_options.py` reads (`ModelContext`'s
 * own comment on `tjpPar` says why), so `buildContext` gives them a literal
 * default instead. That default is still the sheet's own, read by hand off
 * `Adv_settings` rather than the extractor -- see each setting's own comment
 * below for the row it came from.
 */
import { defaultContext } from "@typfallsmodellen/engine";
import type { ModelContext } from "@typfallsmodellen/engine";

import { fieldSet } from "./controls.js";
import type { FieldSet, Relabel } from "./controls.js";
import type { Lang } from "./i18n.js";
import {
  BURIAL_ONLY_RATE,
  CHURCH_MEMBER_RATE,
  KOMMUNALSKATT,
  KOMMUNALSKATT_YEAR,
  STOCKHOLM_BURIAL_RATE,
  TRANAS_BURIAL_RATE,
} from "./kommunalskatt.js";

type Text = (l: Lang) => string;

const text = (sv: string, en: string): Text => (l) => (l === "sv" ? sv : en);

interface Choice {
  readonly value: number;
  readonly label: Text;
}

type Control =
  | { readonly kind: "number"; readonly min: number; readonly max: number; readonly step: number }
  | { readonly kind: "percent" }
  | { readonly kind: "check" }
  | { readonly kind: "select"; readonly choices: readonly Choice[] };

/**
 * One exposed setting.
 *
 * `get` and `set` work in numbers whatever the context field's own type is, so
 * a checkbox over a `boolean` and one over a 0/1 `number` are the same control
 * here. They are a pair of functions rather than a `keyof ModelContext`,
 * following `TABLE2_COLUMNS`: it costs a line each and buys the whole table
 * type-safety without a cast.
 */
export interface Setting {
  /** Stable id -- the DOM attribute the offline check reads, and the test key. */
  readonly key: string;
  /** The `Adv_settings` row this came from. */
  readonly row: number;
  readonly control: Control;
  readonly label: Text;
  readonly hint?: Text;
  readonly get: (c: ModelContext) => number;
  readonly set: (v: number) => Partial<ModelContext>;
}

export interface Group {
  readonly key: string;
  /** The user manual's section number, which is where the wording comes from. */
  readonly section: string;
  readonly title: Text;
  readonly settings: readonly Setting[];
}

const kr = (min: number, max: number) => ({ kind: "number", min, max, step: 100 }) as const;
const years = (min: number, max: number) => ({ kind: "number", min, max, step: 1 }) as const;
const flag = { kind: "check" } as const;
const percent = { kind: "percent" } as const;

/** Whole kronor per month or year; nobody's typfall needs more than this. */
const MONEY = 100_000_000;

/**
 * "Partiellt uttag IP"/"...PP", rows 82/83: the sheet offers exactly these
 * four shares, not an open percent field -- confirmed against the sheet
 * itself (`Adv_settings!D82`/`D83`), since neither row follows the
 * name-in-column-9 layout `extract_options.py` reads and so neither made it
 * into `options.json`.
 */
const WITHDRAWAL_SHARE: readonly Choice[] = [
  { value: 1, label: text("100 %", "100%") },
  { value: 0.75, label: text("75 %", "75%") },
  { value: 0.5, label: text("50 %", "50%") },
  { value: 0.25, label: text("25 %", "25%") },
];

export const GROUPS: readonly Group[] = [
  {
    key: "saving",
    section: "3.2",
    title: text("Privat sparande och tjänstepension", "Private saving and occupational pension"),
    settings: [
      {
        // row 11 "Privat pensionssparande belopp" / "kronor per månad, sedan 2026"
        key: "ipsMonthly",
        row: 11,
        control: kr(0, MONEY),
        label: text("Privat pensionssparande", "Private pension saving"),
        get: (c) => c.ipsMonthly,
        set: (ipsMonthly) => ({ ipsMonthly }),
      },
      {
        // row 12 "Privat pensionssparande sedan när"
        key: "ipsStart",
        row: 12,
        control: years(1960, 2100),
        label: text("Sparandet börjar år", "Saving starts in"),
        get: (c) => c.ipsStart,
        set: (ipsStart) => ({ ipsStart }),
      },
      {
        // row 13 "Typ av privat pensionssparande, (0) IPS, (1) KF, (2) ISK"
        key: "privateSavingKind",
        row: 13,
        control: {
          kind: "select",
          choices: [
            { value: 0, label: text("IPS / pensionsförsäkring", "IPS / pension insurance") },
            { value: 1, label: text("Kapitalförsäkring", "Endowment insurance") },
            { value: 2, label: text("ISK", "Investment savings account") },
          ],
        },
        label: text("Typ av sparande", "Kind of saving"),
        hint: text(
          "IPS ger avdrag under spartiden och beskattas som inkomst; KF och ISK gör tvärtom",
          "IPS is deductible while saving and taxed as income; KF and ISK are the other way round",
        ),
        get: (c) => c.privateSavingKind,
        set: (privateSavingKind) => ({ privateSavingKind }),
      },
      {
        // row 14 "Uttagsålder för tjänstepension (och ev. IPS eller pensionsförsäkring)"
        key: "tjpPar",
        row: 14,
        control: years(0, 100),
        label: text("Uttagsålder för tjänstepension", "Occupational pension drawn from age"),
        hint: text("0 = samma som den allmänna pensionen", "0 = the same as the public pension"),
        get: (c) => c.tjpPar,
        set: (tjpPar) => ({ tjpPar }),
      },
      {
        // row 17 "Temporärt uttag av tjänstepensionen (ange antal år)" / "Livsvarigt"
        key: "tempTjpUttag",
        row: 17,
        control: years(0, 40),
        label: text("Temporärt uttag av tjänstepension", "Occupational pension drawn over"),
        hint: text("antal år, 0 = livsvarigt", "number of years, 0 = lifelong"),
        get: (c) => c.tempTjpUttag,
        set: (tempTjpUttag) => ({ tempTjpUttag }),
      },
      {
        // row 18 "Temporärt uttag av privatsparande (ange antal år)" / "Livsvarigt"
        key: "tempIpsUttag",
        row: 18,
        control: years(0, 40),
        label: text("Temporärt uttag av privat sparande", "Private saving drawn over"),
        hint: text("antal år, 0 = livsvarigt", "number of years, 0 = lifelong"),
        get: (c) => c.tempIpsUttag,
        set: (tempIpsUttag) => ({ tempIpsUttag }),
      },
      {
        // row 15 "Arvsvinster tjänstepension". Manual 3.2: with återbetalnings-
        // skydd the capital goes to survivors, so no inheritance gains accrue.
        key: "occupationalInheritanceGains",
        row: 15,
        control: flag,
        label: text("Arvsvinster på tjänstepensionen", "Inheritance gains on occupational pension"),
        hint: text(
          "Avmarkera för återbetalningsskydd, som avstår arvsvinsterna",
          "Clear it for survivor cover, which forgoes the inheritance gains",
        ),
        get: (c) => c.occupationalInheritanceGains,
        set: (occupationalInheritanceGains) => ({ occupationalInheritanceGains }),
      },
    ],
  },
  {
    key: "insurance",
    section: "3.3",
    title: text("Försäkringstid", "Insurance time"),
    settings: [
      {
        // row 21 "Försäkringstid (bosättningsår fram till 65 års ålder)"
        key: "insuranceYears",
        row: 21,
        control: years(0, 40),
        label: text("Bosättningsår i Sverige fram till 65", "Years resident in Sweden up to 65"),
        hint: text(
          "Färre än 40 år sänker garantipensionen",
          "Fewer than 40 years reduces the guarantee pension",
        ),
        get: (c) => c.insuranceYears,
        set: (insuranceYears) => ({ insuranceYears }),
      },
    ],
  },
  {
    key: "housing",
    section: "3.4",
    title: text("Underlag för bostadstillägg", "Housing supplement basis"),
    settings: [
      {
        // row 29 "Ansöker (=1)" / "Ansöker om bostadsstöd"
        key: "ansokt",
        row: 29,
        control: flag,
        label: text("Ansöker om bostadstillägg", "Applies for housing supplement"),
        get: (c) => c.ansokt,
        set: (ansokt) => ({ ansokt }),
      },
      {
        // row 30 "Boendekostnad per månad för 2025" / "Påverkar BT och ÄFS"
        key: "hyra",
        row: 30,
        control: kr(0, MONEY),
        label: text("Boendekostnad per månad", "Housing cost per month"),
        hint: text("i referensårets priser", "in the reference year's prices"),
        get: (c) => c.hyra,
        set: (hyra) => ({ hyra }),
      },
      {
        // row 31 "Makens / makas årsinkomst"
        key: "makensInkomst",
        row: 31,
        control: kr(0, MONEY),
        label: text("Makens/makans årsinkomst", "Spouse's annual income"),
        hint: text("räknas bara om Gift är ikryssat", "counted only when Married is ticked"),
        get: (c) => c.makensInkomst,
        set: (makensInkomst) => ({ makensInkomst }),
      },
      {
        // row 32 "Förmögenhet (utöver vistelsebostaden)" / "obs: nominellt låst"
        key: "formogenhet",
        row: 32,
        control: kr(0, MONEY),
        label: text("Förmögenhet utöver bostaden", "Wealth beyond the home"),
        get: (c) => c.formogenhet,
        set: (formogenhet) => ({ formogenhet }),
      },
      {
        // row 34 "Kapitalinkomster brutto"
        key: "kapital",
        row: 34,
        control: kr(0, MONEY),
        label: text("Kapitalinkomster brutto per år", "Gross capital income per year"),
        hint: text("antas börja vid pensioneringen", "assumed to start at retirement"),
        get: (c) => c.kapital,
        set: (kapital) => ({ kapital }),
      },
    ],
  },
  {
    key: "tax",
    section: "3.5",
    title: text("Underlag för inkomstskatt", "Income tax basis"),
    settings: [
      {
        // row 39 "Kommunalskatten är antagen till, utelämnas (0) om historiska
        // genomsnitt ska användas"
        key: "kommunalskatt",
        row: 39,
        control: percent,
        label: text("Kommunalskatt", "Municipal tax rate"),
        hint: text(
          "0 använder det historiska genomsnittet",
          "0 uses the historical average",
        ),
        get: (c) => c.kommunalskatt,
        set: (kommunalskatt) => ({ kommunalskatt }),
      },
      {
        // row 40 "Begravningsavgiften samt avgiften till kyrkan/trossamfundet"
        key: "begravningsavgift",
        row: 40,
        control: percent,
        label: text("Begravningsavgift och samfundsavgift", "Burial fee and religious community fee"),
        hint: text(
          `medlem ~${pct(CHURCH_MEMBER_RATE.rate)} %, ej medlem ~${pct(BURIAL_ONLY_RATE.rate)} % (${CHURCH_MEMBER_RATE.year})`,
          `member ~${pct(CHURCH_MEMBER_RATE.rate)}%, non-member ~${pct(BURIAL_ONLY_RATE.rate)}% (${CHURCH_MEMBER_RATE.year})`,
        ),
        get: (c) => c.begravningsavgift,
        set: (begravningsavgift) => ({ begravningsavgift }),
      },
      {
        // row 60 "Fackföreningsavgift"
        key: "fack",
        row: 60,
        control: kr(0, MONEY),
        label: text("Fackföreningsavgift", "Union fee"),
        hint: text("kronor per månad", "kronor per month"),
        get: (c) => c.fack,
        set: (fack) => ({ fack }),
      },
      {
        // row 61 "Avgift -akassa"
        key: "akasseavg",
        row: 61,
        control: kr(0, MONEY),
        label: text("A-kasseavgift", "Unemployment insurance fee"),
        hint: text("kronor per månad", "kronor per month"),
        get: (c) => c.akasseavg,
        set: (akasseavg) => ({ akasseavg }),
      },
    ],
  },
  {
    // Manual 3.7's own example: "Partiellt uttag av inkomstpensionen och
    // premiepensionen kan läggas in här, för att till exempel simulera ett
    // typfall som är jobbonär under en viss period" -- take out a reduced
    // share of the public pension for a few years while still working
    // part-time, then retire in full. `withdrawalShare` (packages/engine/
    // src/income/wages.ts) ties Lön to whichever share is drawn by default,
    // so these three fields are the whole feature -- nothing else has to
    // move for "jobbonär" to show up in Table 2 as a reduced salary
    // alongside a reduced pension.
    //
    // Manual 3.7's other half, "Barnår" (`childBirthYears`), lives in
    // pgb.ts instead, alongside the sheet's other three pension-qualifying-
    // amount sources -- see that file's own comment on why.
    key: "partialWithdrawal",
    section: "3.7",
    title: text("Partiellt uttag", "Partial withdrawal"),
    settings: [
      {
        // row 82 "Partiellt uttag IP" / "... vid 66 med 100 % uttag"
        key: "uttagIp",
        row: 82,
        control: { kind: "select", choices: WITHDRAWAL_SHARE },
        label: text("Andel uttag, inkomstpension", "Income pension withdrawn"),
        hint: text(
          "mellan pensionsåldern och \"Definitivt vid\" nedan",
          "between the retirement age and \"Final at\" below",
        ),
        get: (c) => c.uttagIp,
        set: (uttagIp) => ({ uttagIp }),
      },
      {
        // row 83 "Partiellt uttag PP" / "... vid 66 med 100 % uttag"
        key: "uttagPp",
        row: 83,
        control: { kind: "select", choices: WITHDRAWAL_SHARE },
        label: text("Andel uttag, premiepension", "Premium pension withdrawn"),
        get: (c) => c.uttagPp,
        set: (uttagPp) => ({ uttagPp }),
      },
      {
        // row 81 "Definitivt vid" / "års ålder"
        key: "defAr",
        row: 81,
        control: years(0, 100),
        label: text("Definitivt uttag vid ålder", "Withdrawal becomes final at age"),
        hint: text(
          "0 eller pensionsåldern = fullt uttag direkt, som idag",
          "0 or the retirement age = full withdrawal right away, as today",
        ),
        get: (c) => c.defAr,
        set: (defAr) => ({ defAr }),
      },
    ],
  },
  {
    key: "capital",
    section: "3.8",
    title: text("Känt pensionskapital och avkastning", "Known pension capital and return"),
    settings: [
      {
        // row 87 "Ange Inkomstår"
        key: "pbhYear",
        row: 87,
        control: years(0, 2100),
        label: text("Inkomstår som behållningen avser", "Income year the balances apply to"),
        hint: text("0 = inget känt kapital", "0 = no known capital"),
        get: (c) => c.pbhYear,
        set: (pbhYear) => ({ pbhYear }),
      },
      {
        // row 88 "Behållningen för inkomstpensionen"
        key: "pbhIp",
        row: 88,
        control: kr(0, MONEY),
        label: text("Behållning inkomstpension", "Income pension balance"),
        get: (c) => c.pbhIp,
        set: (pbhIp) => ({ pbhIp }),
      },
      {
        // row 89 "Behållningen för premiepensionen"
        key: "pbhPp",
        row: 89,
        control: kr(0, MONEY),
        label: text("Behållning premiepension", "Premium pension balance"),
        get: (c) => c.pbhPp,
        set: (pbhPp) => ({ pbhPp }),
      },
      {
        // row 90 "Eventuell behållning för tjänstepensionen"
        key: "pbhTjp",
        row: 90,
        control: kr(0, MONEY),
        label: text("Behållning tjänstepension", "Occupational pension balance"),
        get: (c) => c.pbhTjp,
        set: (pbhTjp) => ({ pbhTjp }),
      },
      {
        // row 91 "Eventuellt eget sparande med avdragsrätt (IPS eller p-försäkring)"
        key: "pbhIps",
        row: 91,
        control: kr(0, MONEY),
        label: text("Behållning privat sparande", "Private saving balance"),
        get: (c) => c.pbhIps,
        set: (pbhIps) => ({ pbhIps }),
      },
      {
        // row 94 "1: Angiven real avkastning även historiskt, 2: Historiskt PPM,
        // 3: Historiskt SÅFan"
        key: "returnBasis",
        row: 94,
        control: {
          kind: "select",
          choices: [
            { value: 1, label: text("Angiven real avkastning", "The stated real return") },
            { value: 2, label: text("Historiskt PPM", "Premium pension history") },
            { value: 3, label: text("Historiskt AP7 Såfa", "AP7 Såfa history") },
          ],
        },
        label: text("Historisk avkastning", "Historical return"),
        hint: text(
          "framtiden följer alltid den reala avkastningen på startsidan",
          "the future always follows the real return set above",
        ),
        get: (c) => c.returnBasis,
        set: (returnBasis) => ({ returnBasis }),
      },
      {
        // row 96 "Avkastningen avser efter fondavgifter (1) eller före (0)"
        key: "returnsNetOfFees",
        row: 96,
        control: flag,
        label: text("Avkastningen är efter fondavgifter", "The return is net of fund fees"),
        get: (c) => (c.returnsNetOfFees ? 1 : 0),
        set: (v) => ({ returnsNetOfFees: v === 1 }),
      },
    ],
  },
  {
    key: "other",
    section: "3.6",
    title: text("Övrigt", "Other"),
    settings: [
      {
        // row 103 "Slutlön: Medel av de senaste angivna årens inkomster"
        key: "finalSalaryYears",
        row: 103,
        control: years(1, 40),
        label: text("Slutlönen är medel av de senaste", "The final salary averages the last"),
        hint: text("årens inkomster, och används i Pensionsinkomst", "years, and is used in Pension income"),
        get: (c) => c.finalSalaryYears,
        set: (finalSalaryYears) => ({ finalSalaryYears }),
      },
      {
        // row 43 "(1)-> Pensioneringen sker samma år som slutlönen" -- the
        // row's own "(1)->" shorthand reads like a flag, but manual 3.6 is
        // explicit that it is not one: "Om 0 anges sker pensionering samma år
        // som slutlön. Om större siffra än 0 anges sker pensionering så många
        // år efter slutlönen. [Detta] påverkar resultatet som skrivs ut i
        // Tabell 1." (0 = same year as the final salary, any larger number =
        // that many years after it; the manual's own last sentence scopes the
        // whole effect to Table 1.) `adjustmentFactors` (packages/engine/src/
        // model/result.ts) reads it as `timeLag`, added into the price-index
        // lookup that feeds `beforeRetirement` -- which only rescales the
        // *price-adjusted* ("Fasta priser") column of Table 1's Slutlön/Lön
        // efter skatt/Disponibel inkomst rows. It does not move `par` (the
        // retirement age) or anything in Table 2: raising this to 5 does NOT
        // delay when Income/Premium/Occupational pension start being paid by
        // five years, confirmed against a user's own test after this control
        // first shipped -- only Table 1's own real-terms Slutlön figure moves.
        // A genuine "stop working before the pension starts" scenario is
        // modeled today by setting "Går i pension vid ålder" to the later age
        // and zeroing the gap years in the own salary-path grid instead; this
        // setting is a narrower price-basis knob the manual itself scopes to
        // Table 1, not a withdrawal-timing control. A checkbox here (writing
        // only 0 or 1) could still only ever reach a one-year shift, which is
        // why it is a plain year count now -- that part of the fix stands
        // regardless of the setting's own narrow real-world scope.
        key: "pensionSameYearAsFinalSalary",
        row: 43,
        control: years(0, 40),
        label: text("Slutlönens referensår efter pensioneringen", "Final salary's reference year after retiring"),
        hint: text(
          "Justerar bara Slutlönens belopp i Tabell 1 (Fasta priser) -- flyttar inte när pensionen betalas ut",
          "Only adjusts the Slutlön figure in Table 1 (Fasta priser) -- does not move when the pension itself starts",
        ),
        get: (c) => c.pensionSameYearAsFinalSalary,
        set: (pensionSameYearAsFinalSalary) => ({ pensionSameYearAsFinalSalary }),
      },
      {
        // row 46 "Flexpension för ITP 1 och SAF-LO från och med 2014". Manual
        // 3.6: "lägger till en extra premie till de ovan nämnda
        // tjänstepensionsavtalen från 2014 och framåt. Anges 0 läggs ingen
        // premie till, om större procentsats än 0 läggs den angivna premien
        // till" -- already wired into itp.ts/safLo.ts (`flexPension`, added
        // straight onto both agreements' own premium rates for `year > 2013`),
        // just not reachable from this panel before now.
        key: "flexPension",
        row: 46,
        control: percent,
        label: text("Flexpension, ITP 1 och SAF-LO", "Flex pension, ITP 1 and SAF-LO"),
        hint: text(
          "Extra premie i %, från och med 2014. 0 = ingen premie",
          "Extra premium in %, from 2014 onward. 0 = no premium",
        ),
        get: (c) => c.flexPension,
        set: (flexPension) => ({ flexPension }),
      },
    ],
  },
];

/** Every exposed setting, flattened -- what the tests walk. */
export const SETTINGS: readonly Setting[] = GROUPS.flatMap((g) => g.settings);

/**
 * A companion `<select>` beside a Setting's own percent field -- not through
 * `fieldSet.select()`, which assumes every option *is* a valid `ModelContext`
 * value that stays selected as that value. These two exist only to fire a
 * one-shot fill into the percent field below them: switching municipality is
 * a fresh choice every time, not a value the model itself keeps track of, so
 * there is nothing for the select's own selected option to persist.
 */
function pct(v: number): string {
  return (v * 100).toLocaleString("sv-SE", { minimumFractionDigits: 2, maximumFractionDigits: 3 });
}

/**
 * "Kommunal skattesats": picking a municipality writes its exact
 * `KOMMUNALSKATT_YEAR` rate into the percent field below -- the field itself
 * is still what the run reads, so the number shown is always the one used,
 * the same rule a linked control follows in form.ts (the riktålder checkbox
 * does the same to the retirement-age field).
 *
 * Not from the workbook: see kommunalskatt.ts's own header for the source.
 */
function municipalitySelect(
  lang: Lang,
  onPick: (fraction: number) => void,
): { element: HTMLElement; relabel: Relabel; select: HTMLSelectElement } {
  const wrap = document.createElement("div");
  wrap.className = "field field-wide adv-fill";

  const select = document.createElement("select");
  select.dataset.setting = "kommunalskatt-municipality";
  const placeholder = document.createElement("option");
  placeholder.value = "";
  select.append(placeholder);
  for (const name of Object.keys(KOMMUNALSKATT).sort((a, b) => a.localeCompare(b, "sv"))) {
    const option = document.createElement("option");
    option.value = String(KOMMUNALSKATT[name]);
    option.textContent = name;
    select.append(option);
  }
  select.addEventListener("change", () => {
    const rate = Number(select.value);
    if (select.value !== "" && Number.isFinite(rate)) onPick(rate / 100);
  });

  const link = document.createElement("a");
  link.href =
    "https://www.scb.se/hitta-statistik/statistik-efter-amne/offentlig-ekonomi/finanser-for-den-kommunala-sektorn/kommunalskatterna/pong/tabell-och-diagram/totala-kommunala-skattesatser-2026-kommunvis/";
  link.target = "_blank";
  link.rel = "noopener";
  link.className = "field-hint field-link";

  const relabel = (l: Lang) => {
    // Short: a long placeholder truncates inside a native <select> at this
    // sidebar's width, with no ellipsis or way to read the rest without
    // opening it -- the year belongs on the link below instead, which has a
    // whole line to itself.
    placeholder.textContent = l === "sv" ? "— Välj kommun —" : "— Choose a municipality —";
    link.textContent = l === "sv" ? `SCB:s lista (${KOMMUNALSKATT_YEAR}) ↗` : `SCB's list (${KOMMUNALSKATT_YEAR}) ↗`;
  };
  relabel(lang);

  wrap.append(select, link);
  return { element: wrap, relabel, select };
}

/** Which option `churchBurialSelect` last picked -- `""` is "own rate", the
 * only choice that leaves the percent field alone rather than filling it. */
type ChurchChoice = "member" | "stockholm" | "tranas" | "rest" | "";

/**
 * "Begravningsavgift": covers both the church-fee and burial-fee halves of
 * the manual request, because `context.begravningsavgift` is one number --
 * Adv_settings row 40, "Begravningsavgiften samt avgiften till kyrkan/
 * trossamfundet" -- not two, and taxAndBenefits.ts reads it as a single
 * `kyrkskatt` line. Two of these five options reuse data the workbook's own
 * K_skatt sheet already extracts (`kyrkoavgift`, `begravningsavgift` in
 * kommunalskatt.ts); the Stockholm/Tranås figures are cited external facts
 * (Skatteverket), the same way `returnBasis`'s PPM/AP7 choices already are.
 */
function churchBurialSelect(
  lang: Lang,
  onPick: (fraction: number, choice: ChurchChoice) => void,
): { element: HTMLElement; relabel: Relabel; select: HTMLSelectElement } {
  const wrap = document.createElement("div");
  wrap.className = "field field-wide adv-fill";

  const options: readonly { readonly choice: ChurchChoice; readonly rate: number | null }[] = [
    { choice: "member", rate: CHURCH_MEMBER_RATE.rate },
    { choice: "stockholm", rate: STOCKHOLM_BURIAL_RATE },
    { choice: "tranas", rate: TRANAS_BURIAL_RATE },
    { choice: "rest", rate: BURIAL_ONLY_RATE.rate },
    { choice: "", rate: null },
  ];
  const select = document.createElement("select");
  select.dataset.setting = "begravningsavgift-select";
  const optionEls = options.map(({ choice }) => {
    const el = document.createElement("option");
    el.value = choice;
    select.append(el);
    return el;
  });
  select.addEventListener("change", () => {
    const picked = options.find((o) => o.choice === select.value);
    if (picked?.rate !== null && picked?.rate !== undefined) onPick(picked.rate, picked.choice);
  });

  const link = document.createElement("a");
  link.href = "https://www.svenskakyrkan.se/medlem/kyrkoavgiften";
  link.target = "_blank";
  link.rel = "noopener";
  link.className = "field-hint field-link";

  const relabel = (l: Lang) => {
    // Short, same reason municipalitySelect's placeholder is: a long option
    // truncates inside a native <select> at this sidebar's width with no way
    // to read the rest short of opening it. The exact figures a label would
    // have carried are one line down instead, in the field's own hint text
    // (see the "begravningsavgift" Setting's `hint` below) and in the number
    // the pick itself fills in.
    const labels: Record<ChurchChoice, string> =
      l === "sv"
        ? {
            member: "Medlem i Svenska kyrkan/annat trossamfund",
            stockholm: "Inte medlem, Stockholms stad",
            tranas: "Inte medlem, Tranås kommun",
            rest: "Inte medlem, övriga Sverige",
            "": "— Egen sats —",
          }
        : {
            member: "Member of the Church of Sweden/another faith community",
            stockholm: "Not a member, City of Stockholm",
            tranas: "Not a member, Tranås municipality",
            rest: "Not a member, rest of Sweden",
            "": "— Own rate —",
          };
    for (const [i, el] of optionEls.entries()) el.textContent = labels[options[i]!.choice];
    link.textContent = l === "sv" ? "Hitta din församling ↗" : "Find your parish ↗";
  };
  relabel(lang);

  wrap.append(select, link);
  return { element: wrap, relabel, select };
}

/**
 * "Privat pensionssparande" (row 11): `ipsMonthly` means two different things
 * depending on its own size -- `earnPrivateSaving` (packages/engine/src/
 * model/mcalc.ts) reads a value over 1 as kronor per month and a value at or
 * below 1 as a share of income instead, one cell's dual meaning inherited
 * from the workbook. `options.json`'s own `IPS_start` entry carries a
 * giveaway hint straight off `Adv_settings!C12` in the real sheet, "0 procent
 * av årsinkomsten, sedan 2026" -- sitting on row 12's own line there, whether
 * by design or by how the original sheet happens to be laid out, rather than
 * on row 11's where the setting it explains actually lives. Neither this
 * project's own extractor nor this file ever carried that hint into either
 * field's own UI before now, so the second meaning was reachable only by
 * already knowing to type a fraction into a box labelled "kronor per månad".
 *
 * This toggle makes both meanings their own labelled choice, each with its
 * own kind of field (kronor, percent) rather than one box whose meaning
 * silently depends on how big the number typed into it happens to be.
 * Switching resets the value to 0 rather than converting between them: a
 * kronor figure and a share of a still-varying income have no single right
 * conversion, and 0 means "nothing set" the same way under either reading.
 *
 * The underlying landmine survives on purpose, faithfully: typing exactly
 * "1" into the kronor field is still 1 kr/month by the field's own label,
 * but `ipsMonthly > 1` reads it as the *share* branch instead (100% of
 * income) -- the workbook's own off-by-one, not smoothed over here, and
 * vanishingly unlikely in practice since every real kronor figure in this
 * app is a multiple of 100.
 */
function savingAmountOrShare(
  lang: Lang,
  initial: number,
  percentControl: FieldSet["percent"],
  onChange: (v: number) => void,
): { element: HTMLElement; relabel: Relabel; setValue: (v: number) => void } {
  const wrap = document.createElement("div");
  wrap.className = "adv-ips";
  wrap.dataset.setting = "ipsMonthly";

  const toggle = document.createElement("div");
  toggle.className = "panel-toggle";
  toggle.setAttribute("role", "group");
  const amountBtn = document.createElement("button");
  amountBtn.type = "button";
  amountBtn.dataset.mode = "amount";
  const shareBtn = document.createElement("button");
  shareBtn.type = "button";
  shareBtn.dataset.mode = "share";
  toggle.append(amountBtn, shareBtn);

  const amountInput = document.createElement("input");
  amountInput.type = "number";
  amountInput.inputMode = "numeric";
  amountInput.min = "0";
  amountInput.max = String(MONEY);
  amountInput.step = "100";
  amountInput.dataset.setting = "ipsMonthly-amount";

  // Reuses the same control every other percent field in this app does --
  // reported as showing "1.7" instead of "1,7" for another percent field, a
  // bug this one would otherwise have repeated with its own hand-rolled input.
  const share = percentControl(0, onChange);
  share.element.dataset.setting = "ipsMonthly-share";

  // The one thing a bare number can't say for itself: which of the two this
  // is. "kr" is not spelled out the same way for the amount field, matching
  // every other kronor field in this app, none of which do either. Its own
  // row, not `field()`'s usual hint slot below the label -- that slot is one
  // line for the whole widget and can't toggle with the mode the way this
  // one, sitting right beside the share input itself, does.
  const shareRow = document.createElement("div");
  shareRow.className = "adv-ips-share";
  const hint = document.createElement("span");
  hint.className = "field-hint";
  hint.textContent = "%";
  shareRow.append(share.element, hint);

  // A value at or below 1 already means "share" to the engine; use the same
  // rule here to decide which mode a freshly loaded or reset value opens in.
  let mode: "amount" | "share" = initial > 0 && initial <= 1 ? "share" : "amount";

  const applyMode = () => {
    amountBtn.className = mode === "amount" ? "panel-btn active" : "panel-btn";
    shareBtn.className = mode === "share" ? "panel-btn active" : "panel-btn";
    amountInput.hidden = mode !== "amount";
    shareRow.hidden = mode !== "share";
  };

  const setValue = (v: number) => {
    mode = v > 0 && v <= 1 ? "share" : "amount";
    amountInput.value = String(mode === "amount" ? Math.round(v) : 0);
    share.setValue(mode === "share" ? v : 0);
    applyMode();
  };
  setValue(initial);

  const switchTo = (next: "amount" | "share") => {
    if (mode === next) return;
    mode = next;
    amountInput.value = "0";
    share.setValue(0);
    applyMode();
    onChange(0);
  };
  amountBtn.addEventListener("click", () => switchTo("amount"));
  shareBtn.addEventListener("click", () => switchTo("share"));

  amountInput.addEventListener("change", () => {
    const typed = Number(amountInput.value);
    if (!Number.isFinite(typed) || amountInput.value.trim() === "") {
      amountInput.value = "0";
      return;
    }
    const clamped = Math.min(Math.max(Math.round(typed), 0), MONEY);
    amountInput.value = String(clamped);
    onChange(clamped);
  });

  const relabel = (l: Lang) => {
    amountBtn.textContent = l === "sv" ? "Belopp" : "Amount";
    shareBtn.textContent = l === "sv" ? "Andel av inkomst" : "Share of income";
    share.relabel(l);
  };
  relabel(lang);

  wrap.append(toggle, amountInput, shareRow);
  return { element: wrap, relabel, setValue };
}

export interface AdvancedHandle {
  readonly element: HTMLElement;
  relabel(lang: Lang): void;
  /** `Använd normala inställningar`: every control back to the workbook's own. */
  reset(): void;
}

/**
 * The panel.
 *
 * One `<details>` per group, closed to begin with: twenty-five fields open at
 * once would bury the nine on the Start sheet that most runs only ever touch.
 */
export function createAdvancedPanel(
  lang: Lang,
  onChange: (patch: Partial<ModelContext>) => void,
): AdvancedHandle {
  const element = document.createElement("div");
  element.className = "advanced";

  const relabels: Relabel[] = [];
  const restores: (() => void)[] = [];
  const normal = defaultContext();

  for (const group of GROUPS) {
    const box = document.createElement("details");
    box.className = "adv-group";
    box.dataset.group = group.key;
    const summary = document.createElement("summary");
    const body = document.createElement("div");
    body.className = "adv-body";

    const applyTitle = (l: Lang) => {
      summary.textContent = group.title(l);
    };
    applyTitle(lang);
    relabels.push(applyTitle);

    const { field, number, percent: percentControl, check, select } = fieldSet(body, relabels, lang);

    // Filled in when the "tax" group reaches `begravningsavgift`, below --
    // `kommunalskatt` comes first in that group's own settings array and
    // needs to call into it on a user's *later* pick, once the whole loop (and
    // so this closure) has run, which is why a `let` assigned out of order
    // here is safe rather than a forward-reference bug.
    let applyChurchAutoDefault: () => void = () => {};
    let churchExplicit = false;

    for (const setting of group.settings) {
      const initial = setting.get(normal);
      const label = (l: Lang) => ({
        label: setting.label(l),
        ...(setting.hint ? { hint: setting.hint(l) } : {}),
      });

      if (setting.control.kind === "number" && setting.key === "ipsMonthly") {
        // field-wide: a toggle plus a number needs a whole row, not the
        // 104px second column every other field's control gets.
        const widget = savingAmountOrShare(lang, initial, percentControl, (v) => onChange(setting.set(v)));
        relabels.push(widget.relabel);
        field(widget.element, label, "field field-wide");
        restores.push(() => widget.setValue(initial));
      } else if (setting.control.kind === "number") {
        const { min, max, step } = setting.control;
        const control = number(initial, { min, max, step }, (v) => onChange(setting.set(v)));
        control.element.dataset.setting = setting.key;
        field(control.element, label);
        restores.push(() => control.setValue(initial));
      } else if (setting.control.kind === "percent") {
        // These two get three decimal places, not the usual one: both can be
        // filled from a picked, cited rate (a municipality's own published
        // rate, Tranås's 0.285%) rather than typed, and the field showing a
        // coarser number than the one it just picked is exactly the mismatch
        // an early reviewer caught, comparing this field's "1.3" against its
        // own hint's "1,32 %" underneath it.
        const precise = setting.key === "kommunalskatt" || setting.key === "begravningsavgift";
        const control = percentControl(
          initial,
          (v) => {
            if (setting.key === "begravningsavgift") churchExplicit = true;
            onChange(setting.set(v));
          },
          precise ? { maxDecimals: 3 } : undefined,
        );
        control.element.dataset.setting = setting.key;
        relabels.push(control.relabel);

        if (setting.key === "kommunalskatt") {
          // Picking a municipality flips `historicalTaxRate` off in setup.ts,
          // which stops `begravavg` from following the historical-average
          // series it was reading and starts reading this panel's own
          // `begravningsavgift` field instead -- silently 0 unless something
          // has set it. Rather than let a one-click municipality picker trip
          // that footgun, an untouched church/burial field is auto-filled
          // with this port's own honest default for it: the same population
          // average `historicalTaxRate` would have used anyway.
          const municipality = municipalitySelect(lang, (fraction) => {
            control.setValue(fraction);
            onChange(setting.set(fraction));
            if (!churchExplicit) applyChurchAutoDefault();
          });
          relabels.push(municipality.relabel);
          body.append(municipality.element);
          restores.push(() => {
            municipality.select.value = "";
          });
        }
        if (setting.key === "begravningsavgift") {
          const church = churchBurialSelect(lang, (fraction, choice) => {
            churchExplicit = true;
            church.select.value = choice;
            control.setValue(fraction);
            onChange(setting.set(fraction));
          });
          relabels.push(church.relabel);
          body.append(church.element);
          applyChurchAutoDefault = () => {
            churchExplicit = true;
            church.select.value = "rest";
            control.setValue(BURIAL_ONLY_RATE.rate);
            onChange(setting.set(BURIAL_ONLY_RATE.rate));
          };
          restores.push(() => {
            church.select.value = "";
            churchExplicit = false;
          });
        }

        field(control.element, label);
        restores.push(() => control.setValue(initial));
      } else if (setting.control.kind === "check") {
        const control = check(initial === 1, (on) => onChange(setting.set(on ? 1 : 0)));
        control.dataset.setting = setting.key;
        field(control, label, "field field-check");
        restores.push(() => {
          control.checked = initial === 1;
        });
      } else {
        const control = select(setting.control.choices, initial, (v) => onChange(setting.set(v)));
        control.element.dataset.setting = setting.key;
        relabels.push(control.relabel);
        field(control.element, label, "field field-wide");
        restores.push(() => {
          control.element.value = String(initial);
        });
      }
    }

    box.append(summary, body);
    element.append(box);
  }

  return {
    element,
    relabel(l) {
      for (const r of relabels) r(l);
    },
    reset() {
      for (const r of restores) r();
    },
  };
}
