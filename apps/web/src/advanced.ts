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
 * Twenty-five of the sheet's seventy-six rows are here, grouped as sections 3.2
 * to 3.8 of the user manual group them. The rest are left out on purpose: some
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
 * comes from `defaultContext()`, which reads options.json.
 */
import { defaultContext } from "@typfallsmodellen/engine";
import type { ModelContext } from "@typfallsmodellen/engine";

import { fieldSet } from "./controls.js";
import type { Relabel } from "./controls.js";
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
        hint: text("kronor per månad", "kronor per month"),
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
        // row 43 "(1)-> Pensioneringen sker samma år som slutlönen"
        key: "pensionSameYearAsFinalSalary",
        row: 43,
        control: flag,
        label: text("Pensionering samma år som slutlönen", "Retire in the same year as the final salary"),
        get: (c) => c.pensionSameYearAsFinalSalary,
        set: (pensionSameYearAsFinalSalary) => ({ pensionSameYearAsFinalSalary }),
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

      if (setting.control.kind === "number") {
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
