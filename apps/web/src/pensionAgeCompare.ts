/**
 * "Jämför avtalsområden": the second view of the "Bästa pensionsålder" tab.
 *
 * The tab's first view answers *which retirement age pays the most* for one
 * collective agreement. This one asks the same of all of them at once, side by
 * side: for each agreement the best age at every lifespan, what it pays in
 * total, and what it adds to having no occupational pension at all.
 * `pensionAge.ts` does the arithmetic (`computeComparison` and what reads it --
 * its file comment says what is held constant and what varies); this file is
 * the view around it.
 *
 * What it takes from the tab and what it does not. The person, the economic
 * assumptions, the lifespan range and the before/after-tax choice come from the
 * tab's own inputs, so the two views always describe the same person. The
 * agreement chosen there and its "include occupational pension" switch belong
 * to the single view: a comparison is *about* the agreements, so it ignores the
 * switch. The agreement chosen above is shown along with the cohort's own until
 * somebody picks (it is the one being asked about) and is marked "ditt val"
 * wherever it appears.
 *
 * **Which agreements.** Until someone picks, the ones the cohort can belong to
 * (`schemeFitsCohort`) plus the one chosen above; picking any chip freezes the
 * selection as it stands. An agreement written for other birth years can be
 * switched on -- the model does not stop it -- but is always labelled as a
 * "what if". "Saknar tjänstepension" is the baseline, always there.
 *
 * **Cost.** One study is seven model runs, a couple of milliseconds each, so
 * all eight agreements are about 60 ms. They are only computed while this view
 * shows, and kept (by `comparisonKey` and the identity of the shared context)
 * across everything that does not change the inputs: picking an agreement,
 * choosing a retirement age, moving the lifespan to read at, the chart toggle.
 *
 * The controls are built once and never rebuilt, so a field being typed into
 * keeps its focus; only the result panels are rebuilt on every change. Colours
 * are the agreements' own (`--fig-ag-*`, categorical, fixed per agreement and
 * not per position, so switching one off does not repaint the rest); the
 * baseline is ink. Every figure's numbers are also in a table, and the
 * agreements are named by text everywhere a colour is used.
 */
import { Scheme } from "@typfallsmodellen/engine";
import type { DeathProbabilities, ModelContext, SchemeId } from "@typfallsmodellen/engine";

import { fieldSet } from "./controls.js";
import type { Bounds, Relabel } from "./controls.js";
import { csvExportButton, xlsxExportButton } from "./export.js";
import { kronor, locale, percent as share } from "./format.js";
import { t } from "./i18n.js";
import type { Lang } from "./i18n.js";
import {
  AGREEMENTS,
  LIFESPAN_BOUNDS,
  NEAR_TIE,
  bestAgeSpread,
  cohortText,
  comparisonKey,
  comparisonRows,
  comparisonToCsv,
  comparisonToXlsxRows,
  computeComparison,
  defaultSchemes,
  kindsPaid,
  schemeFitsCohort,
  schemeName,
  schemeSector,
  totalSeries,
} from "./pensionAge.js";
import type { AgeBasis, Comparison, ComparisonRow, Metric, Study, StudyInput } from "./pensionAge.js";
import { niceTicks, renderLineChart } from "./pensionAgeChart.js";
import { chip, h, mkr, panel, priceText, refName, row, say, tableOf, tkr } from "./pensionAgeUi.js";
import { cell, headCell } from "./tables.js";

const LIFESPAN: Bounds = { ...LIFESPAN_BOUNDS, step: 1 };

/** The "lever till" age the table and the short answer are read at until someone changes it. */
const READ_AT = 85;

export interface CompareOutput {
  readonly nodes: HTMLElement[];
  /** What the baseline study says, for the salary field's hint in the tab's inputs. */
  readonly derived: { readonly finalSalary: number; readonly monthlySalary: number };
}

export interface CompareHandle {
  /** The agreement picker and the two read-off controls -- built once, shown only in this view. */
  readonly controls: HTMLElement;
  relabel(lang: Lang): void;
  /** Back to the defaults: the cohort's own agreements, the best age, 85, the totals chart. */
  reset(): void;
  /** Computes and builds the results for the inputs as they are now, and brings the controls in line with them. */
  render(args: {
    readonly input: StudyInput;
    readonly chosen: SchemeId;
    readonly context: ModelContext;
    readonly deaths: DeathProbabilities;
  }): CompareOutput;
}

// ---------------------------------------------------------------- helpers ---

const clamp = (value: number, low: number, high: number): number => Math.min(Math.max(value, low), high);

const agreementColour = (scheme: SchemeId): string => `--fig-ag-${scheme}`;

/** Thousands of kronor with their unit: "1 674 tkr". */
const kk = (value: number, l: Lang): string => `${tkr(value, l)} ${say(l, "tkr", "thousand kr")}`;

/** "a, b och c". */
function joinNames(names: readonly string[], l: Lang): string {
  if (names.length <= 1) return names.join("");
  return `${names.slice(0, -1).join(", ")} ${say(l, "och", "and")} ${names[names.length - 1]}`;
}

/** What stands beside the end of a line: short enough for the chart's modest right margin. */
function endName(scheme: SchemeId, l: Lang): string {
  if (scheme === Scheme.None) return say(l, "Saknar", "None");
  if (scheme === Scheme.Pa16Avd2) return "PA16-2";
  if (scheme === Scheme.Pa16Avd1) return "PA16-1";
  return schemeName(scheme, l);
}

/** A line key in the agreement's own colour -- the table's counterpart of the legend swatch. */
const agreementKey = (scheme: SchemeId): HTMLElement => h("span", `pa-key pa-key-ag${scheme}`);

/** What the retirement age is, in a sentence: "vid den bästa åldern för respektive avtal" / "vid 65 år". */
function atPhrase(basis: AgeBasis, l: Lang): string {
  return basis === "best"
    ? say(l, "vid den bästa åldern för respektive avtal", "at each agreement's best age")
    : say(l, `vid ${basis} år`, `at ${basis}`);
}

/** The state the builders read besides the comparison itself. */
interface View {
  readonly lang: Lang;
  readonly basis: AgeBasis;
  readonly readAt: number;
  readonly metric: Metric;
  /** The agreement chosen in the tab's inputs. */
  readonly chosen: SchemeId;
}

const referenceStudy = (comparison: Comparison): Study => comparison.entries[0]!.study;

function measureOf(comparison: Comparison, l: Lang): string {
  const tax = comparison.input.afterTax ? say(l, "efter skatt", "after tax") : say(l, "före skatt", "before tax");
  // The same savings in every entry, so the baseline's say whether there are any.
  const kinds = kindsPaid(referenceStudy(comparison));
  const own = kinds.individual || kinds.savings ? say(l, " + eget sparande", " + own saving") : "";
  return say(l, `allmän pension + tjänstepension${own}, ${tax}`, `public + occupational pension${own}, ${tax}`);
}

// --------------------------------------------------------------- the answer ---

function answerPanel(comparison: Comparison, rows: readonly ComparisonRow[], view: View): HTMLElement {
  const l = view.lang;
  const { section, body } = panel(say(l, "Svar i korthet", "Short answer"));
  section.dataset.role = "pa-compare-answer";
  section.classList.add("pa-answer");
  body.append(
    h(
      "p",
      "pa-lead",
      say(
        l,
        "Samma person och samma sju pensionsåldrar under olika avtalsområden. Varje avtalsområde räknas för sig och mäts mot att sakna tjänstepension.",
        "The same person and the same seven retirement ages under different collective agreements. Each agreement is calculated on its own and measured against having no occupational pension.",
      ),
    ),
  );

  const baseline = rows[0]!;
  const agreements = rows.slice(1);
  if (agreements.length === 0) {
    body.append(
      h("p", "pa-kpi", say(l, "Välj minst ett avtalsområde ovan för att jämföra.", "Pick at least one agreement above to compare.")),
    );
    return section;
  }

  const name = (r: ComparisonRow): string => schemeName(r.entry.scheme, l);
  const list = h("ul", "pa-points");
  const bullet = (text: string): void => {
    list.append(h("li", undefined, text));
  };
  const D = view.readAt;
  const at = atPhrase(view.basis, l);
  const ranked = [...agreements].sort((a, b) => b.total - a.total);
  const top = ranked[0]!;
  const low = ranked[ranked.length - 1]!;

  // 1. the level: who pays most, and what an agreement is worth
  const worth =
    top.extra > 0
      ? say(
          l,
          `, alltså ${kk(top.extra, l)} (${share(top.extraShare, l)}) mer än att sakna tjänstepension (${kk(baseline.total, l)})`,
          `, that is ${kk(top.extra, l)} (${share(top.extraShare, l)}) more than having no occupational pension (${kk(baseline.total, l)})`,
        )
      : "";
  let level = say(
    l,
    `Om du lever till ${D} år och går i pension ${at} ger ${name(top)} mest pension: ${kk(top.total, l)}${worth}.`,
    `If you live to ${D} and retire ${at}, ${name(top)} pays the most: ${kk(top.total, l)}${worth}.`,
  );
  if (ranked.length > 1 && low.total > 0) {
    const gap = top.total - low.total;
    level += say(
      l,
      ` ${name(low)} ger minst av avtalen, ${kk(low.total, l)}; skillnaden mellan det bästa och det sämsta avtalet är ${kk(gap, l)} (${share(gap / low.total, l)}).`,
      ` ${name(low)} pays the least of the agreements, ${kk(low.total, l)}; the difference between the best and the worst agreement is ${kk(gap, l)} (${share(gap / low.total, l)}).`,
    );
  }
  bullet(level);

  // 2. the agreement chosen in the inputs
  const mine = agreements.find((r) => r.entry.scheme === view.chosen);
  if (mine !== undefined) {
    const place = ranked.indexOf(mine) + 1;
    bullet(
      say(
        l,
        `Ditt val i fältet ”${t("occupational", l)}”, ${name(mine)}, hamnar på plats ${place} av ${ranked.length} avtal vid ${D} år (${kk(mine.total, l)}); bästa pensionsålder är ${mine.best.par} år.`,
        `Your choice in the “${t("occupational", l)}” field, ${name(mine)}, comes ${place} of ${ranked.length} agreements at ${D} (${kk(mine.total, l)}); its best retirement age is ${mine.best.par}.`,
      ),
    );
  }

  // 3. the timing: does the agreement move the best age?
  const byAge = new Map<number, string[]>();
  for (const r of rows) {
    const names = byAge.get(r.best.par) ?? [];
    names.push(name(r));
    byAge.set(r.best.par, names);
  }
  const ages = [...byAge.keys()].sort((a, b) => a - b);
  let timing =
    ages.length === 1
      ? say(
          l,
          `Bästa pensionsålder vid ${D} år är ${ages[0]} för alla jämförda alternativ.`,
          `The best retirement age at ${D} is ${ages[0]} for every option compared.`,
        )
      : say(
          l,
          `Bästa pensionsålder vid ${D} år: ${ages.map((age) => `${age} för ${joinNames(byAge.get(age)!, l)}`).join("; ")}.`,
          `The best retirement age at ${D}: ${ages.map((age) => `${age} for ${joinNames(byAge.get(age)!, l)}`).join("; ")}.`,
        );
  const spread = bestAgeSpread(comparison);
  if (spread.length > 1) {
    const same = spread.filter((s) => s.min === s.max).length;
    const widest = spread.reduce((w, s) => (s.max - s.min > w.max - w.min ? s : w), spread[0]!);
    const range = `${comparison.lifespans[0]}–${comparison.lifespans[comparison.lifespans.length - 1]}`;
    timing +=
      same === spread.length
        ? say(l, ` Över hela intervallet ${range} år är den densamma för alla.`, ` Across the whole range ${range} it is the same for all of them.`)
        : say(
            l,
            ` Över intervallet ${range} år är den densamma för alla vid ${same} av ${spread.length} livslängder; som mest skiljer den ${widest.max - widest.min} år (vid ${widest.lifespan} år).`,
            ` Across ${range} it is the same for all of them at ${same} of ${spread.length} lifespans; at most it differs by ${widest.max - widest.min} years (at ${widest.lifespan}).`,
          );
  }
  bullet(timing);

  // 4. agreements written for other birth years
  const whatIf = agreements.filter((r) => !r.entry.fits);
  if (whatIf.length > 0) {
    const born = comparison.input.born;
    const text = whatIf.map((r) => `${name(r)} (${cohortText(r.entry.scheme, l) ?? ""})`);
    bullet(
      say(
        l,
        `Avtal som normalt inte gäller födda ${born} är med som ”vad om”-jämförelse: ${joinNames(text, l)}.`,
        `Agreements that do not normally cover people born in ${born} are included as a “what if”: ${joinNames(text, l)}.`,
      ),
    );
  }

  body.append(list);
  return section;
}

// ------------------------------------------------------------------- grid ---

/** The first column of a table row: key, name, who it is for, and the tags that say why it may not belong. */
function nameCell(entry: { scheme: SchemeId; fits: boolean }, view: View): HTMLTableCellElement {
  const l = view.lang;
  const td = cell("", "pa-name");
  const line = h("span", "pa-name-line");
  line.append(agreementKey(entry.scheme), h("span", "pa-name-text", schemeName(entry.scheme, l)));
  td.append(line);
  const notes: string[] = [];
  if (entry.scheme !== Scheme.None) notes.push(schemeSector(entry.scheme, l));
  const born = cohortText(entry.scheme, l);
  if (born !== undefined) notes.push(born);
  if (notes.length > 0) td.append(h("small", undefined, notes.join(" · ")));
  const tags = h("span", "pa-tags");
  if (!entry.fits) tags.append(h("span", "pa-tag pa-tag-off", say(l, "vad om", "what if")));
  if (entry.scheme === view.chosen && entry.scheme !== Scheme.None) tags.append(h("span", "pa-tag", say(l, "ditt val", "your choice")));
  if (tags.childElementCount > 0) td.append(tags);
  return td;
}

function gridPanel(comparison: Comparison, view: View): HTMLElement {
  const l = view.lang;
  const actions = [
    csvExportButton(l, say(l, "avtalsomraden.csv", "agreements.csv"), () => comparisonToCsv(comparison, l)),
    xlsxExportButton(l, say(l, "avtalsomraden.xlsx", "agreements.xlsx"), () => comparisonToXlsxRows(comparison, l)),
  ];
  const { section, body } = panel(say(l, "Bästa pensionsålder per avtalsområde", "Best retirement age by agreement"), actions);
  section.dataset.role = "pa-compare-grid-panel";
  body.append(
    h(
      "p",
      "pa-lead",
      say(
        l,
        "Varje ruta visar bästa pensionsålder för avtalsområdet om du lever till åldern överst; ≈ betyder att de två bästa åldrarna ligger inom 1 % från varandra. Kolumnen för den ålder tabellen nedan läses av vid är markerad.",
        "Each box shows the best retirement age for the agreement if you live to the age at the top; ≈ means the two best ages are within 1 % of each other. The column for the age the table below is read at is highlighted.",
      ),
    ),
  );

  const key = h("div", "pa-grid-key");
  key.append(
    h("span", undefined, say(l, "Färg = pensionsåldern mot referensåldern, varmare tidigare och kallare senare:", "Colour = retirement age against the reference age, warmer earlier and cooler later:")),
    ...referenceStudy(comparison).options.map((option) => chip(option)),
  );
  body.append(key);

  const head = [
    headCell(say(l, "Avtalsområde", "Agreement")),
    ...comparison.lifespans.map((lifespan) => {
      const th = headCell(String(lifespan));
      th.title = say(l, `Lever till ${lifespan} år`, `Lifespan ${lifespan}`);
      if (lifespan === view.readAt) th.classList.add("pa-read");
      return th;
    }),
  ];
  const rows = comparison.entries.map((entry) => {
    const cells: HTMLTableCellElement[] = [nameCell(entry, view)];
    for (const lifespan of comparison.lifespans) {
      const result = entry.study.rows.find((r) => r.lifespan === lifespan)!;
      const best = entry.study.options[result.best]!;
      const second = entry.study.options[result.runnerUp]!;
      const tie = result.marginShare < NEAR_TIE;
      const td = cell("", lifespan === view.readAt ? "num pa-gcell pa-read" : "num pa-gcell");
      td.dataset.lifespan = String(lifespan);
      td.dataset.best = String(best.par);
      td.dataset.tie = tie ? "1" : "0";
      td.title = say(
        l,
        `${schemeName(entry.scheme, l)}, lever till ${lifespan}: bäst är pension vid ${best.par} (${kk(result.totals[result.best]!, l)}), därefter ${second.par} (${kk(result.totals[result.runnerUp]!, l)}).`,
        `${schemeName(entry.scheme, l)}, lifespan ${lifespan}: best is retiring at ${best.par} (${kk(result.totals[result.best]!, l)}), then ${second.par} (${kk(result.totals[result.runnerUp]!, l)}).`,
      );
      const box = chip(best);
      if (tie) box.append(h("sup", undefined, "≈"));
      td.append(box);
      cells.push(td);
    }
    return row(cells, { scheme: String(entry.scheme) });
  });
  body.append(tableOf("pa-compare-grid", "pa-gridtable", head, rows));
  return section;
}

// ----------------------------------------------------------------- figure ---

/** Two lines closer than this share of the axis, everywhere, cover each other on screen. */
const OVERLAP = 0.015;

/**
 * Groups of agreements whose lines lie on top of one another. A reader looking
 * for ITP-1 under KAP-KL would otherwise find an orange key in the legend and no
 * orange line, so the figure says which lines hide which (they differ by a few
 * kronor a month; the table and the hover readout hold the difference).
 */
function coincident(series: readonly { entry: { scheme: SchemeId }; points: readonly (readonly [number, number])[] }[], tolerance: number): SchemeId[][] {
  const groups: SchemeId[][] = [];
  const placed = new Set<SchemeId>();
  series.forEach((a, i) => {
    if (placed.has(a.entry.scheme)) return;
    const group = [a.entry.scheme];
    for (const b of series.slice(i + 1)) {
      if (placed.has(b.entry.scheme)) continue;
      const apart = Math.max(...a.points.map((point, k) => Math.abs(point[1] - (b.points[k]?.[1] ?? Infinity))));
      if (apart <= tolerance) group.push(b.entry.scheme);
    }
    if (group.length > 1) {
      group.forEach((scheme) => placed.add(scheme));
      groups.push(group);
    }
  });
  return groups;
}

function totalFigure(comparison: Comparison, view: View): HTMLElement | undefined {
  const l = view.lang;
  const series = totalSeries(comparison, view.basis, view.metric);
  if (series.length === 0) return undefined;
  const ends = series.map((s) => s.points[s.points.length - 1]![1]);
  const highest = Math.max(...ends);
  const lowest = Math.min(...ends);
  const total = view.metric === "total";
  const values = series.flatMap((s) => s.points.map((p) => p[1]));
  const ticks = niceTicks(0, Math.max(1, ...values), 6);
  const mkrTick = new Intl.NumberFormat(locale(l), { maximumFractionDigits: 2 });
  const first = comparison.lifespans[0]!;
  const last = comparison.lifespans[comparison.lifespans.length - 1]!;
  const at = atPhrase(view.basis, l);
  const price = priceText(referenceStudy(comparison), l);
  const hidden = coincident(series, OVERLAP * (ticks.hi - ticks.lo)).map((group) =>
    joinNames(group.map((scheme) => schemeName(scheme, l)), l),
  );
  const overlapNote =
    hidden.length === 0
      ? undefined
      : say(
          l,
          `${hidden.join("; ")} ligger så nära varandra att linjerna täcker varandra – skillnaden finns i tabellen och i hovervisningen.`,
          `${hidden.join("; ")} are so close that their lines cover each other – the difference is in the table and the hover readout.`,
        );

  return renderLineChart({
    role: "pa-fig-compare",
    lang: l,
    title: total
      ? say(l, "Total pension per avtalsområde", "Total pension by agreement")
      : say(l, "Vad avtalsområdet tillför", "What each agreement adds"),
    subtitle: total
      ? say(
          l,
          `Summan som betalats ut före varje ålder, ${at}, ${measureOf(comparison, l)}, miljoner kronor, ${price}, utan diskontering.`,
          `The sum paid out before each age, ${at}, ${measureOf(comparison, l)}, millions of kronor, ${price}, not discounted.`,
        )
      : say(
          l,
          `Total pension minus summan utan tjänstepension, ${at}, tusen kronor, ${price}. Det avtalsområdet lägger till.`,
          `Total pension minus the total with no occupational pension, ${at}, thousands of kronor, ${price}. What the agreement adds.`,
        ),
    xMin: first,
    xMax: Math.max(last, first + 1),
    xTitle: say(l, "Lever till (år)", "Lifespan (years)"),
    yTicks: ticks.ticks,
    yTick: total
      ? (v) => (v === 0 ? "0" : `${mkrTick.format(v / 1e6)} Mkr`)
      : (v) => (v === 0 ? "0" : `+${kronor(v / 1000, l)}`),
    yUnit: total ? "Mkr" : say(l, "tkr", "thousand kr"),
    heading: (x) => `${say(l, "Lever till", "Lifespan")} ${x}`,
    endRoom: 66,
    ...(overlapNote === undefined ? {} : { note: overlapNote }),
    series: series.map((s, i) => ({
      agreement: s.entry.scheme,
      colour: agreementColour(s.entry.scheme),
      emphasis: s.entry.scheme === Scheme.None,
      essential:
        s.entry.scheme === Scheme.None || ends[i] === highest || ends[i] === lowest || s.entry.scheme === view.chosen,
      name: s.entry.fits ? schemeName(s.entry.scheme, l) : `${schemeName(s.entry.scheme, l)} (${say(l, "vad om", "what if")})`,
      end: endName(s.entry.scheme, l),
      points: s.points,
      format: total ? (v: number) => `${mkr(v, l)} Mkr` : (v: number) => `+${kk(v, l)}`,
    })),
  });
}

// ------------------------------------------------------------------ table ---

function tablePanel(comparison: Comparison, rows: readonly ComparisonRow[], view: View): HTMLElement {
  const l = view.lang;
  const D = view.readAt;
  const fixed = view.basis !== "best";
  const occupational = !comparison.input.afterTax;
  const { section, body } = panel(say(l, `Avtalsområdena vid ${D} år`, `The agreements at ${D}`));
  section.dataset.role = "pa-compare-table-panel";
  body.append(
    h(
      "p",
      "pa-lead",
      say(
        l,
        `Summan som betalats ut från pensionsstart tills du fyller ${D} år, ${measureOf(comparison, l)}, i tusen kronor. ${
          fixed
            ? `Alla avtal räknas vid pensionsålder ${view.basis}; tillägget är skillnaden mot att sakna tjänstepension vid samma ålder.`
            : "Varje avtal räknas vid sin egen bästa pensionsålder; tillägget är skillnaden mot att sakna tjänstepension vid dess bästa ålder."
        }`,
        `The sum paid out from the start until you turn ${D}, ${measureOf(comparison, l)}, in thousands of kronor. ${
          fixed
            ? `Every agreement is read at retirement age ${view.basis}; the addition is the difference from having no occupational pension at the same age.`
            : "Each agreement is read at its own best retirement age; the addition is the difference from having no occupational pension at its best age."
        }`,
      ),
    ),
  );

  const agreements = rows.slice(1);
  const topTotal = Math.max(...agreements.map((r) => r.total));
  const head = [
    headCell(say(l, "Avtalsområde", "Agreement")),
    headCell(fixed ? say(l, "Pensionsålder", "Retirement age") : say(l, "Bästa pensionsålder", "Best retirement age")),
    ...(fixed ? [headCell(say(l, "Bästa ålder vid", "Best age at") + ` ${D}`)] : []),
    headCell(say(l, "Total pension, tkr", "Total pension, thousand kr")),
    headCell(say(l, "Tillägg, tkr", "Added, thousand kr"), say(l, "Skillnad mot att sakna tjänstepension", "Difference from having no occupational pension")),
    headCell(say(l, "Tillägg, %", "Added, %")),
    headCell(say(l, "Månadspension vid start, kr", "Monthly pension at start, SEK")),
    ...(occupational ? [headCell(say(l, "varav tjänstepension, kr/mån", "of which occupational, SEK/month"))] : []),
  ];
  const bodyRows = rows.map((r) => {
    const isBase = r.entry.scheme === Scheme.None;
    const total = cell(tkr(r.total, l), agreements.length > 0 && !isBase && r.total === topTotal ? "num pa-top" : "num");
    total.dataset.kr = String(Math.round(r.total));
    const added = cell(isBase ? "–" : tkr(r.extra, l), "num");
    added.dataset.kr = String(Math.round(r.extra));
    const cells: HTMLTableCellElement[] = [nameCell(r.entry, view)];
    const age = cell("", "num");
    age.append(chip(r.option));
    cells.push(age);
    if (fixed) {
      const bestAge = cell("", "num");
      bestAge.append(chip(r.best));
      cells.push(bestAge);
    }
    cells.push(total, added, cell(isBase ? "–" : share(r.extraShare, l), "num"), cell(kronor(r.monthly, l), "num"));
    if (occupational) cells.push(cell(isBase ? "–" : kronor(r.occupationalMonthly, l), "num"));
    const tr = row(cells, { scheme: String(r.entry.scheme), age: String(r.option.par) });
    if (isBase) tr.classList.add("pa-base-row");
    if (r.entry.scheme === view.chosen && !isBase) tr.classList.add("pa-mine");
    return tr;
  });
  body.append(tableOf("pa-compare-table", "pa-agrtable", head, bodyRows));
  return section;
}

// ------------------------------------------------------------------ notes ---

function notesPanel(comparison: Comparison, l: Lang): HTMLElement {
  const study = referenceStudy(comparison);
  const details = h("details", "field-note pa-notes");
  details.dataset.role = "pa-compare-notes";
  details.append(h("summary", undefined, say(l, "Antaganden och metod", "Assumptions and method")));
  const bodyBox = h("div", "field-note-body");
  const tax = comparison.input.afterTax ? say(l, "efter skatt", "after tax") : say(l, "före skatt", "before tax");
  const paragraphs = [
    say(
      l,
      `Alla avtalsområden räknas för samma person och med samma sju pensionsåldrar; bara avtalsområdet skiljer. ”Saknar tjänstepension” är jämförelsebasen: det ett avtalsområde ”tillför” är dess summa minus basens vid samma pensionsålder – eller, vid bästa ålder, vid var och ens egen bästa ålder.`,
      `Every agreement is calculated for the same person and the same seven retirement ages; only the agreement differs. “No occupational pension” is the baseline: what an agreement “adds” is its total minus the baseline's at the same retirement age – or, at the best age, at each one's own best age.`,
    ),
    say(
      l,
      `Pension är allmän pension plus tjänstepension enligt avtalsområdet, ${tax}, från pensionsstart. Eventuell individuell tjänstepension, privat pension och privat sparande från fälten ovan ingår lika för alla avtalsområden (privat sparande är skattefritt och räknas lika före och efter skatt), så det är avtalets tillägg som skiljer dem åt. Belopp i ${priceText(study, l)}, utan diskontering: en krona vid 90 års ålder räknas som en krona vid 70. Tjänstepensionen tas ut samtidigt som den allmänna pensionen, oavsett inställningen under Avancerat för uttagsålder.`,
      `Pension is public pension plus the occupational pension of the agreement, ${tax}, from the start. Any individual occupational pension, private pension and private savings from the fields above are included alike for every agreement (private savings are not taxed and count the same before and after tax), so it is the agreement's addition that sets them apart. Amounts are in ${priceText(study, l)}, not discounted: a krona at 90 counts as much as one at 70. Occupational pension is drawn together with the public pension, whatever the Advanced setting for its start age says.`,
    ),
    say(
      l,
      `”Lever till D år” betyder att pension betalas för alla år från pensionsstart till och med året före D-årsdagen (D − pensionsålder hela år). Pensionens storlek beror inte på hur länge du lever; ingen sannolikhetsviktning görs.`,
      `“Lifespan D” means pension is paid for every year from the start up to the year before the D-year birthday (D − retirement age whole years). The size of the pension does not depend on how long you live; no probability weighting is applied.`,
    ),
    say(
      l,
      `Vissa avtalsområden är skrivna för vissa födelseår: ITP-1 för födda 1979 och senare, ITP-2 före 1979, AKAP-KR 1986 och senare, PA16 avd 2 före 1988 och PA16 avd 1 1988 och senare. Modellen hindrar inte att en annan årskull räknas på dem, och det gör inte heller den här jämförelsen, men sådana avtal är en ”vad om”-jämförelse och väljs därför inte från början – slå på dem med knapparna.`,
      `Some agreements are written for particular birth years: ITP-1 for those born 1979 or later, ITP-2 before 1979, AKAP-KR 1986 or later, PA16 section 2 before 1988 and PA16 section 1 1988 or later. The model does not stop another cohort from being calculated on them, and neither does this comparison, but such an agreement is a “what if” and so is not selected to begin with – switch them on with the buttons.`,
    ),
    say(
      l,
      `Personen, sparandet, antagandena och livslängdsintervallet kommer från fälten ovan; valet av avtalsområde där och ”Räkna med avtalets tjänstepension” påverkar inte jämförelsen. Skatteinställningar, löneprofil och prisbas kommer från Avancerat. Standardantagandena (0 % inflation, 0 % real löneökning) skjuter brytpunkterna flera år framåt jämfört med antagandet om en växande lön – pröva ”Med löneökning” för att se hur mycket svaret beror på det.`,
      `The person, the savings, the assumptions and the lifespan range come from the fields above; the agreement chosen there and “Include the agreement's occupational pension” do not affect the comparison. Tax settings, wage profile and price basis come from Advanced. The default assumptions (0 % inflation, 0 % real wage growth) push the break-even points several years later than if wages grow – try “With wage growth” to see how much the answer depends on it.`,
    ),
    say(
      l,
      `Månadspension vid start är första årets belopp i år-för-år-flödet som diagrammen och summorna bygger på; appens Tabell 1 visar något högre belopp vid pensionering, eftersom den räknar in sista årets pensionsrätt direkt.`,
      `Monthly pension at the start is the first year's amount in the year-by-year flow the charts and totals are built on; the app's Table 1 shows a slightly higher amount at retirement, because it counts the last year's pension right at once.`,
    ),
  ];
  for (const text of paragraphs) bodyBox.append(h("p", undefined, text));
  details.append(bodyBox);
  return details;
}

function warningsBox(messages: readonly string[]): HTMLElement {
  const box = h("aside", "warnings");
  for (const message of messages) box.append(h("p", undefined, message));
  return box;
}

function buildResults(comparison: Comparison, view: View): HTMLElement[] {
  const rows = comparisonRows(comparison, view.readAt, view.basis);
  const out: HTMLElement[] = [];
  if (comparison.warnings.length > 0) out.push(warningsBox(comparison.warnings));
  out.push(answerPanel(comparison, rows, view), gridPanel(comparison, view));
  const figure = totalFigure(comparison, view);
  if (figure !== undefined) out.push(figure);
  if (rows.length > 1) out.push(tablePanel(comparison, rows, view));
  out.push(notesPanel(comparison, view.lang));
  return out;
}

// ------------------------------------------------------------------ panel ---

export function createAgreementCompare(lang: Lang, onChange: () => void): CompareHandle {
  let currentLang = lang;
  /** The agreements somebody picked; null until anybody does, which means the defaults. */
  let picked: ReadonlySet<SchemeId> | null = null;
  let basis: AgeBasis = "best";
  let readAt = READ_AT;
  let metric: Metric = "total";
  /** The agreements an unpicked selection is, for the inputs last rendered. */
  let defaults: readonly SchemeId[] = [];
  let born = 0;
  let chosen: SchemeId = Scheme.None;
  let cache:
    | { key: string; context: ModelContext; deaths: DeathProbabilities; studies: Map<SchemeId, Study> }
    | undefined;
  let lastComparison: Comparison | undefined;

  const selection = (): ReadonlySet<SchemeId> => picked ?? new Set(defaults);

  // ---- the controls, built once ---------------------------------------------
  const controls = h("section", "panel pa-compare-controls");
  controls.dataset.role = "pa-compare-controls";
  const head = h("h2");
  const title = h("span");
  head.append(title);
  const intro = h("p", "pa-intro");

  const pick = h("div", "pa-pick");
  const pickLabel = h("span", "pa-pick-label");
  const chipsBox = h("div", "pa-chips");
  chipsBox.setAttribute("role", "group");
  const baseChip = h("span", "pa-agr pa-agr-base");
  const baseName = h("span", "pa-agr-name");
  const baseNote = h("small", "pa-agr-note");
  baseChip.append(agreementKey(Scheme.None), baseName, baseNote);
  chipsBox.append(baseChip);

  const chips = new Map<SchemeId, { button: HTMLButtonElement; name: HTMLElement; note: HTMLElement }>();
  for (const scheme of AGREEMENTS) {
    const button = h("button", "pa-agr");
    button.type = "button";
    button.dataset.action = "pa-agr";
    button.dataset.scheme = String(scheme);
    const name = h("span", "pa-agr-name");
    const note = h("small", "pa-agr-note");
    button.append(agreementKey(scheme), name, note);
    button.addEventListener("click", () => {
      const next = new Set(selection());
      if (next.has(scheme)) next.delete(scheme);
      else next.add(scheme);
      picked = next;
      onChange();
    });
    chipsBox.append(button);
    chips.set(scheme, { button, name, note });
  }
  pick.append(pickLabel, chipsBox);

  const grid = h("div", "pa-grid");
  controls.append(head, intro, pick, grid);

  const relabels: Relabel[] = [];
  const { field, number } = fieldSet(grid, relabels, lang);

  // The retirement age the totals are read at. Its options follow the study (the
  // alternatives depend on the birth year and the reference age), so they are
  // rebuilt when those change -- the select itself is not.
  const basisSelect = h("select");
  basisSelect.dataset.setting = "paCompareBasis";
  basisSelect.addEventListener("change", () => {
    basis = basisSelect.value === "best" ? "best" : Number(basisSelect.value);
    onChange();
  });
  field(
    basisSelect,
    (l) => ({
      label: say(l, "Pensionsålder", "Retirement age"),
      hint: say(l, "Åldern som summorna avser", "The age the totals are for"),
    }),
    "field field-wide",
  );
  let optionsKey = "";

  const readField = number(readAt, LIFESPAN, (value) => {
    readAt = value;
    onChange();
  });
  readField.element.dataset.setting = "paCompareLifespan";
  field(readField.element, (l) => ({
    label: say(l, "Läs av vid, lever till (år)", "Read at, lifespan (years)"),
    hint: say(l, "Gäller tabellen och svaret", "For the table and the answer"),
  }));

  const metricRow = h("div", "pa-pick pa-metric");
  const metricLabel = h("span", "pa-pick-label");
  const metricBox = h("div", "panel-toggle");
  metricBox.setAttribute("role", "group");
  metricRow.append(metricLabel, metricBox);
  const metricButtons = (
    [
      { metric: "total", sv: "Total pension", en: "Total pension" },
      { metric: "extra", sv: "Tillägg mot utan avtal", en: "Added vs no agreement" },
    ] as const
  ).map((choice) => {
    const button = h("button", "panel-btn");
    button.type = "button";
    button.dataset.metric = choice.metric;
    button.addEventListener("click", () => {
      metric = choice.metric;
      onChange();
    });
    metricBox.append(button);
    return { choice, button };
  });
  grid.append(metricRow);

  const applyText = (l: Lang): void => {
    title.textContent = say(l, "Jämför avtalsområden", "Compare agreements");
    intro.textContent = say(
      l,
      `Samma person och samma sju pensionsåldrar, men olika avtalsområden, sida vid sida. Personen, sparandet, antagandena och livslängdsintervallet kommer från fälten ovan; ”${t("occupational", l)}” och ”Räkna med avtalets tjänstepension” där gäller bara det enskilda läget.`,
      `The same person and the same seven retirement ages, but different collective agreements, side by side. The person, the savings, the assumptions and the lifespan range come from the fields above; “${t("occupational", l)}” and “Include the agreement's occupational pension” there only apply to the single view.`,
    );
    pickLabel.textContent = say(l, "Avtalsområden", "Agreements");
    chipsBox.setAttribute("aria-label", say(l, "Avtalsområden att jämföra", "Agreements to compare"));
    baseName.textContent = schemeName(Scheme.None, l);
    baseNote.textContent = say(l, "jämförelsebas", "baseline");
    metricLabel.textContent = say(l, "Diagrammet visar", "The chart shows");
    metricBox.setAttribute("aria-label", say(l, "Vad diagrammet visar", "What the chart shows"));
    for (const { choice, button } of metricButtons) button.textContent = say(l, choice.sv, choice.en);
  };
  applyText(lang);

  /** Brings every control in line with the state and the comparison just computed. */
  function syncControls(comparison: Comparison): void {
    const l = currentLang;
    const selected = selection();
    for (const [scheme, { button, name, note }] of chips) {
      const on = selected.has(scheme);
      const fits = schemeFitsCohort(scheme, born);
      button.setAttribute("aria-pressed", String(on));
      button.classList.toggle("active", on);
      button.classList.toggle("pa-agr-off", !fits);
      name.textContent = schemeName(scheme, l);
      note.textContent = fits ? "" : say(l, "vad om", "what if");
      note.hidden = fits;
      const cohort = cohortText(scheme, l);
      button.title = fits
        ? `${schemeSector(scheme, l)}${cohort === undefined ? "" : `, ${cohort}`}`
        : say(
            l,
            `${schemeName(scheme, l)} gäller ${cohort ?? ""}; du är född ${born}. Tas med som ”vad om”-jämförelse.`,
            `${schemeName(scheme, l)} is for people ${cohort ?? ""}; you were born in ${born}. Included as a “what if” comparison.`,
          );
    }

    // The offered retirement ages: the best for each lifespan, then each alternative.
    const study = referenceStudy(comparison);
    const key = `${currentLang}|${comparison.ages.join(",")}|${comparison.reference}|${refName(study, l)}`;
    if (key !== optionsKey) {
      optionsKey = key;
      basisSelect.replaceChildren();
      const add = (value: string, text: string): void => {
        const option = h("option", undefined, text);
        option.value = value;
        basisSelect.append(option);
      };
      add("best", say(l, "Bästa för varje livslängd", "Best for each lifespan"));
      comparison.ages.forEach((age, index) => {
        add(String(age), `${age} ${say(l, "år", "yrs")}${index === comparison.reference ? ` (${refName(study, l).replace(/^the /, "")})` : ""}`);
      });
    }
    basisSelect.value = basis === "best" ? "best" : String(basis);

    readField.setValue(readAt);
    readField.element.min = String(comparison.lifespans[0] ?? LIFESPAN.min);
    readField.element.max = String(comparison.lifespans[comparison.lifespans.length - 1] ?? LIFESPAN.max);
    for (const { choice, button } of metricButtons) {
      button.classList.toggle("active", choice.metric === metric);
      button.setAttribute("aria-pressed", String(choice.metric === metric));
    }
  }

  return {
    controls,
    relabel(l) {
      currentLang = l;
      applyText(l);
      for (const relabel of relabels) relabel(l);
      optionsKey = "";
      if (lastComparison !== undefined) syncControls(lastComparison);
    },
    reset() {
      picked = null;
      basis = "best";
      readAt = READ_AT;
      metric = "total";
    },
    render(args) {
      const { input, context, deaths } = args;
      chosen = args.chosen;
      born = input.born;
      defaults = defaultSchemes(input.born, args.chosen);

      const key = comparisonKey(input);
      if (cache === undefined || cache.key !== key || cache.context !== context || cache.deaths !== deaths) {
        cache = { key, context, deaths, studies: new Map() };
      }
      const comparison = computeComparison(input, [...selection()], context, deaths, cache.studies);
      lastComparison = comparison;

      // The state a changed input can invalidate: an age the study no longer offers, a lifespan outside its range.
      if (basis !== "best" && !comparison.ages.includes(basis)) basis = "best";
      readAt = clamp(readAt, comparison.lifespans[0] ?? LIFESPAN.min, comparison.lifespans[comparison.lifespans.length - 1] ?? LIFESPAN.max);
      syncControls(comparison);

      const view: View = { lang: currentLang, basis, readAt, metric, chosen };
      return {
        nodes: buildResults(comparison, view),
        derived: {
          finalSalary: referenceStudy(comparison).options[comparison.reference]?.finalSalaryMonthly ?? 0,
          monthlySalary: input.monthlySalary,
        },
      };
    },
  };
}
