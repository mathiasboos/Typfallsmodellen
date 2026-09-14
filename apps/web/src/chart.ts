/**
 * The Start sheet's three figures, as inline SVG.
 *
 * The chart objects themselves are unrecoverable -- `Workbook_Open` clears the
 * result sheets, so every series reference in the distributed file reads
 * `#REFERENS!`. What survives is the sheet the charts are fed from:
 * `Data_till_Start` holds one block per figure, with live headings and live
 * formulas, and its columns say exactly what is plotted:
 *
 *   A2:L22   ages Int(par)-10 .. Int(par)+10 (A12 = Int(par)) -- Ålder, Lön,
 *            Lön vid fortsatt arbete, Inkomstpension, Premiepension,
 *            Tilläggspension, Garantipension+Pensionstillägg, Tjänstepension,
 *            Bruttoinkomst, Inkomst efter skatt, Disponibel inkomst, Bidrag.
 *            Feeds "Diagram 1 i start" (Figur 2) and "Diagram 2 i start".
 *   N3:T78   ages 0..75 -- År, Ålder, Löpande priser, Fasta priser (2025),
 *            Dagens (2025) lönenivå, CPI, Income index. Feeds Figur 1.
 *   Y2:Y10   the series names of Figur 2, as SysLang row lookups;
 *   AN2:AN6  the same for the disposable income chart.
 *
 * Every source column is one of the seventeen `mvalues` columns, which this
 * port returns as `result.rows` -- including columns 16 and 17, the price and
 * wage-level factors, which exist so Figur 1's three series can be derived from
 * a single run.
 *
 * `renderTaxChart` below is not one of these three: it is not a workbook
 * chart at all, the same kind of addition the KPI cards are. It reads columns
 * 18 and 19, municipal and state tax, which this port itself splits out of
 * `netto` (see `MvaluesRow.municipalTax`/`stateTax`) since the workbook
 * computes and discards that split rather than ever showing it.
 */
import type { MvaluesRow, TypfallResult } from "@typfallsmodellen/engine";

import { kronor } from "./format.js";
import { dropHeadingNumber, t } from "./i18n.js";
import type { Lang } from "./i18n.js";

const SVG = "http://www.w3.org/2000/svg";

const WIDTH = 760;
const HEIGHT = 330;
const PAD = { top: 26, right: 14, bottom: 42, left: 68 } as const;
const PLOT = {
  left: PAD.left,
  right: WIDTH - PAD.right,
  top: PAD.top,
  bottom: HEIGHT - PAD.bottom,
} as const;

/** What the figures need beyond the result itself. */
export interface FigureView {
  readonly lang: Lang;
  /** The retirement age the run used. */
  readonly par: number;
  /** `rng_Chart_Earning_factor`: the amounts are shown per month. */
  readonly perMonth: boolean;
  /** `rng_Bara_fastapriser`: 1 fixed prices, 0 wage level, -1 nominal. */
  readonly priceBasis: number;
}

/** One plotted series: where its colour comes from and how it reads a row. */
interface Series {
  /** Stable across languages, so it can name a CSS class. */
  readonly key: string;
  /** Composed rather than a plain key: `Y8` concatenates two of the sheet's. */
  readonly name: (lang: Lang) => string;
  /** A custom property in styles.css, so light and dark are separate steps. */
  readonly colour: string;
  readonly mark: "fill" | "line";
  readonly get: (row: MvaluesRow) => number;
}

function el<K extends keyof SVGElementTagNameMap>(
  name: K,
  attrs: Record<string, string | number> = {},
): SVGElementTagNameMap[K] {
  const node = document.createElementNS(SVG, name);
  for (const [k, v] of Object.entries(attrs)) node.setAttribute(k, String(v));
  return node;
}

/** "Pensionstillägg (IPT) (n/40)", with this run's years rather than 40/40. */
function ipt(lang: Lang, result: TypfallResult): string {
  return t("iptFull", lang).replace(/\(\d+\/40\)/, `(${Math.min(result.qualifyingYears, 40)}/40)`);
}

function label(text: string, x: number, y: number, className: string): SVGTextElement {
  const node = el("text", { x, y, class: className });
  node.textContent = text;
  return node;
}

/** A tick step that lands on 1/2/5 x 10^n, so the axis reads in round numbers. */
function niceStep(span: number, target: number): number {
  const raw = span / Math.max(target, 1);
  const magnitude = 10 ** Math.floor(Math.log10(Math.max(raw, 1)));
  for (const m of [1, 2, 5, 10]) {
    if (magnitude * m >= raw) return magnitude * m;
  }
  return magnitude * 10;
}

interface Vertical {
  /** The top gridline, which sits above the tallest mark rather than on it. */
  readonly top: number;
  readonly step: number;
  readonly y: (value: number) => number;
}

/** The vertical scale, shared by all three figures, with Excel's headroom. */
function vertical(max: number): Vertical {
  const step = niceStep(Math.max(max, 1), 8);
  const top = Math.max(Math.ceil((max * 1.02) / step) * step, step);
  return { top, step, y: (value) => PLOT.bottom - (value / top) * (PLOT.bottom - PLOT.top) };
}

function gridlines(svg: SVGSVGElement, scale: Vertical, lang: Lang): void {
  for (let v = 0; v <= scale.top + scale.step / 2; v += scale.step) {
    const y = scale.y(v);
    svg.append(el("line", { x1: PLOT.left, x2: PLOT.right, y1: y, y2: y, class: "grid" }));
    svg.append(label(kronor(v, lang), PLOT.left - 8, y + 4, "tick tick-y"));
  }
  svg.append(
    el("line", {
      x1: PLOT.left,
      x2: PLOT.right,
      y1: scale.y(0),
      y2: scale.y(0),
      class: "axis",
    }),
  );
}

/**
 * The years from the retirement age on are shaded, as both column charts in the
 * workbook shade them.
 */
function shadeRetirement(svg: SVGSVGElement, from: number): void {
  if (from >= PLOT.right) return;
  svg.append(
    el("rect", {
      x: Math.max(from, PLOT.left),
      y: PLOT.top,
      width: PLOT.right - Math.max(from, PLOT.left),
      height: PLOT.bottom - PLOT.top,
      class: "retired",
    }),
  );
}

function legend(series: readonly Series[], lang: Lang): HTMLElement {
  const box = document.createElement("ul");
  box.className = "legend";
  for (const item of series) {
    const li = document.createElement("li");
    const swatch = document.createElement("span");
    swatch.className = item.mark === "line" ? "swatch swatch-line" : "swatch";
    if (item.mark === "line") swatch.style.borderTopColor = `var(${item.colour})`;
    else swatch.style.background = `var(${item.colour})`;
    const text = document.createElement("span");
    text.textContent = item.name(lang);
    li.append(swatch, text);
    box.append(li);
  }
  return box;
}

/** One column of the readout: an age, and every series' value at it. */
interface Point {
  readonly x: number;
  readonly heading: string;
  readonly row: MvaluesRow;
}

/** Crosshair and a readout of every series at the nearest age. */
function hover(
  svg: SVGSVGElement,
  figure: HTMLElement,
  points: readonly Point[],
  series: readonly Series[],
  lang: Lang,
): void {
  if (points.length === 0) return;
  const line = el("line", { y1: PLOT.top, y2: PLOT.bottom, class: "crosshair", x1: 0, x2: 0 });
  line.style.display = "none";
  svg.append(line);

  const tip = document.createElement("div");
  tip.className = "tooltip";
  tip.hidden = true;
  figure.append(tip);

  svg.addEventListener("pointermove", (event: PointerEvent) => {
    const box = svg.getBoundingClientRect();
    const xInView = ((event.clientX - box.left) / box.width) * WIDTH;
    let nearest = points[0]!;
    for (const point of points) {
      if (Math.abs(point.x - xInView) < Math.abs(nearest.x - xInView)) nearest = point;
    }

    line.setAttribute("x1", String(nearest.x));
    line.setAttribute("x2", String(nearest.x));
    line.style.display = "";

    tip.replaceChildren();
    const head = document.createElement("strong");
    head.textContent = nearest.heading;
    tip.append(head);
    for (const item of series) {
      const row = document.createElement("div");
      row.className = "tooltip-row";
      const swatch = document.createElement("span");
      swatch.className = item.mark === "line" ? "swatch swatch-line" : "swatch";
      if (item.mark === "line") swatch.style.borderTopColor = `var(${item.colour})`;
      else swatch.style.background = `var(${item.colour})`;
      const name = document.createElement("span");
      name.textContent = item.name(lang);
      const value = document.createElement("span");
      value.className = "num";
      value.textContent = kronor(item.get(nearest.row), lang);
      row.append(swatch, name, value);
      tip.append(row);
    }
    tip.hidden = false;
    const left = (nearest.x / WIDTH) * box.width;
    tip.style.left = `${Math.min(Math.max(left, 8), box.width - 8)}px`;
  });

  svg.addEventListener("pointerleave", () => {
    line.style.display = "none";
    tip.hidden = true;
  });
}

function newSvg(title: string): SVGSVGElement {
  return el("svg", {
    viewBox: `0 0 ${WIDTH} ${HEIGHT}`,
    role: "img",
    "aria-label": title,
    preserveAspectRatio: "xMidYMid meet",
  });
}

function frame(
  title: string,
  subtitle: string | undefined,
  svg: SVGSVGElement,
  key: HTMLElement,
  notes: readonly string[] = [],
): HTMLElement {
  const figure = document.createElement("figure");
  figure.className = "figure";
  const caption = document.createElement("figcaption");
  caption.textContent = title;
  figure.append(caption);
  if (subtitle !== undefined) {
    const sub = document.createElement("p");
    sub.className = "figure-subtitle";
    sub.textContent = subtitle;
    figure.append(sub);
  }
  const plot = document.createElement("div");
  plot.className = "plot";
  plot.append(svg, key);
  figure.append(plot);
  for (const note of notes) {
    const line = document.createElement("p");
    line.className = "figure-note";
    line.textContent = note;
    figure.append(line);
  }
  return figure;
}

/** The amounts are annual; `rng_Chart_Earning_factor` divides them. */
function scale(view: FigureView): number {
  return view.perMonth ? 12 : 1;
}

/**
 * The nominal amount behind a row.
 *
 * `recordRow` multiplies every amount by the price factor the settings ask for,
 * and columns 16 and 17 carry both factors, so any of the three views can be
 * recovered from the row -- which is what Figur 1's `P3` does before `Q3` and
 * `R3` re-express it.
 */
function nominal(value: number, row: MvaluesRow, priceBasis: number): number {
  const factor = priceBasis === 1 ? row.kpiFactor : priceBasis === 0 ? row.indexFactor : 1;
  return factor > 0 ? value / factor : value;
}

// ---------------------------------------------------------------- Figur 1 ---

/**
 * Figur 1: the whole life, in the three price views.
 *
 * `Data_till_Start!P:R` takes the gross income of each age -- `Utdata!L`, the
 * `brutto` column -- back to nominal and then re-expresses it: `Q = P x CPI
 * factor`, `R = P x income-index factor`. The block's first 23 rows read the
 * separate nominal wage column instead, which is the same number while nobody
 * is drawing a pension; this uses `brutto` throughout.
 */
export function renderFigure1(result: TypfallResult, view: FigureView): HTMLElement {
  const { lang } = view;
  const per = scale(view);
  const rows = result.rows;

  const series: readonly Series[] = [
    {
      key: "current",
      name: (l) => t("currentPrices", l),
      colour: "--fig-current",
      mark: "line",
      get: (r) => nominal(r.brutto, r, view.priceBasis) / per,
    },
    {
      key: "fixed",
      name: (l) => t("fixedPrices", l),
      colour: "--fig-fixed",
      mark: "line",
      get: (r) => (nominal(r.brutto, r, view.priceBasis) * r.kpiFactor) / per,
    },
    {
      key: "wage-level",
      name: (l) => t("todayWageLevel", l),
      colour: "--fig-wage-level",
      mark: "line",
      get: (r) => (nominal(r.brutto, r, view.priceBasis) * r.indexFactor) / per,
    },
  ];

  // "Figur 1." is dropped on request; the rest of the sheet's own heading stays.
  const title =
    `${dropHeadingNumber(t("figure1", lang))}-${view.par} ${t("andPensionFrom", lang)} ` +
    `${view.par} ${t("yearsAge", lang)}`;
  const svg = newSvg(title);

  const first = rows[0]?.age ?? 0;
  const last = rows[rows.length - 1]?.age ?? first + 1;
  const x = (age: number) =>
    PLOT.left + ((age - first) / Math.max(last - first, 1)) * (PLOT.right - PLOT.left);
  const max = Math.max(...rows.flatMap((r) => series.map((s) => s.get(r))), 1);
  const axis = vertical(max);
  const y = axis.y;

  gridlines(svg, axis, lang);

  // Age ticks every five years, with the calendar year under every second one,
  // as the workbook's own axis carries both.
  for (let age = Math.ceil(first / 5) * 5; age <= last; age += 5) {
    svg.append(label(String(age), x(age), PLOT.bottom + 16, "tick tick-x"));
    if (age % 10 === 0) {
      const row = rows.find((r) => r.age === age);
      if (row) svg.append(label(String(row.year), x(age), PLOT.bottom + 31, "tick tick-x tick-year"));
    }
  }

  for (const item of series) {
    const points = rows.map((row) => `${x(row.age)},${y(item.get(row))}`).join(" ");
    const line = el("polyline", {
      points,
      fill: "none",
      stroke: `var(${item.colour})`,
      class: `line line-${item.key}`,
    });
    svg.append(line);
  }

  const figure = frame(title, undefined, svg, legend(series, lang), [
    t("fixedPricesNote", lang),
    t("currentPricesNote", lang),
  ]);
  hover(
    svg,
    figure,
    rows.map((row) => ({ x: x(row.age), heading: `${row.year} · ${row.age}`, row })),
    series,
    lang,
  );
  return figure;
}

// ------------------------------------------------- the two column charts ---

/**
 * `A2:A22` on the sheet is ten years either side of retirement; widened here
 * on request to five years before retirement through age 100, so the working
 * years get less room and the retirement years more.
 */
function aroundRetirement(result: TypfallResult, par: number): readonly MvaluesRow[] {
  const from = Math.trunc(par) - 5;
  return result.rows.filter((row) => row.age >= from && row.age <= 100);
}

/**
 * Age ticks along a column chart's x-axis: every fifth age once the window is
 * wide enough that one label per column would collide, but always the first
 * and last column so the reader can see exactly where the chart starts and
 * ends.
 */
function ageTicks(
  svg: SVGSVGElement,
  rows: readonly MvaluesRow[],
  centre: (i: number) => number,
): void {
  const step = rows.length > 25 ? 5 : 1;
  rows.forEach((row, i) => {
    const edge = i === 0 || i === rows.length - 1;
    if (!edge && row.age % step !== 0) return;
    svg.append(label(String(row.age), centre(i), PLOT.bottom + 16, "tick tick-x"));
  });
}

/**
 * Drops a series from the legend when it never has a value in the rows being
 * drawn -- on request, for a typfall with no occupational pension or no
 * garantipension, say. The stacked band or line itself is still drawn
 * (harmlessly invisible at zero); this only trims the legend, not the
 * chart's own arithmetic.
 */
function visibleSeries(rows: readonly MvaluesRow[], series: readonly Series[]): readonly Series[] {
  return series.filter((item) => rows.some((row) => item.get(row) > 0));
}

/** Where each column sits, and how wide it is. */
function columns(count: number): {
  centre: (i: number) => number;
  width: number;
  step: number;
} {
  const step = (PLOT.right - PLOT.left) / Math.max(count, 1);
  return { centre: (i) => PLOT.left + step * (i + 0.5), width: step * 0.66, step };
}

/** The shaded band starts at the left edge of the retirement year's column. */
function retirementEdge(
  rows: readonly MvaluesRow[],
  par: number,
  centre: (i: number) => number,
  step: number,
): number {
  const index = rows.findIndex((row) => row.age >= Math.trunc(par));
  return index < 0 ? PLOT.right : centre(index) - step / 2;
}

/** Draws one stack of columns, bottom band first. */
function stack(
  svg: SVGSVGElement,
  rows: readonly MvaluesRow[],
  series: readonly Series[],
  y: (v: number) => number,
  centre: (i: number) => number,
  width: number,
): void {
  rows.forEach((row, i) => {
    let base = 0;
    for (const item of series) {
      const value = item.get(row);
      if (value <= 0) continue;
      const top = base + value;
      svg.append(
        el("rect", {
          x: centre(i) - width / 2,
          y: y(top),
          width,
          height: Math.max(y(base) - y(top), 0),
          fill: `var(${item.colour})`,
          class: "bar",
        }),
      );
      base = top;
    }
  });
}

function overlay(
  svg: SVGSVGElement,
  rows: readonly MvaluesRow[],
  item: Series,
  y: (v: number) => number,
  centre: (i: number) => number,
  from?: number,
): void {
  const points = rows
    .map((row, i) => ({ row, i }))
    .filter(({ row }) => from === undefined || row.age >= from)
    .map(({ row, i }) => `${centre(i)},${y(item.get(row))}`)
    .join(" ");
  if (points === "") return;
  svg.append(
    el("polyline", {
      points,
      fill: "none",
      stroke: `var(${item.colour})`,
      class: `line line-${item.key}`,
    }),
  );
}

/**
 * Figur 2: what the income is made of, year by year around retirement.
 *
 * The stack is `Data_till_Start` columns B, D, G, E, H in the order the chart
 * lays them: Lön, Inkomstpension (which the sheet defines as inkomst- plus
 * tilläggspension), Garantipension with Pensionstillägget, Premiepension and
 * Tjänstepension. The sheet's separate Tilläggspension column, F, is left out
 * on purpose: D already contains it, so stacking both would count ATP twice for
 * everyone born before 1954.
 *
 * "Lön vid fortsatt arbete" (`Data_till_Start!C`, the salary the run would
 * have paid had work continued) was drawn as a fourth overlay here; removed
 * on request, series and legend both.
 */
export function renderFigure2(result: TypfallResult, view: FigureView): HTMLElement {
  const { lang } = view;
  const per = scale(view);
  const rows = aroundRetirement(result, view.par);

  const bands: readonly Series[] = [
    {
      key: "salary",
      name: (l) => t("salaryChart", l),
      colour: "--fig-salary",
      mark: "fill",
      get: (r) => r.income / per,
    },
    {
      key: "income-pension",
      name: (l) => t("incomePension", l),
      colour: "--fig-income-pension",
      mark: "fill",
      get: (r) => (r.ip + r.tp) / per,
    },
    {
      key: "guarantee",
      // `Y8`: "Garantipension" and the supplement's own label, joined, with this
      // run's qualifying years in place of the shipped typfall's 40/40.
      name: (l) => `${t("guarantee", l)} ${ipt(l, result)}`,
      colour: "--fig-guarantee",
      mark: "fill",
      get: (r) => (r.garp + r.ptillagg) / per,
    },
    {
      key: "premium",
      name: (l) => t("premium", l),
      colour: "--fig-premium",
      mark: "fill",
      get: (r) => r.pp / per,
    },
    {
      key: "occupational",
      name: (l) => t("occupationalChart", l),
      colour: "--fig-occupational",
      mark: "fill",
      get: (r) => r.tjp / per,
    },
  ];
  const lines: readonly Series[] = [
    {
      key: "after-tax",
      name: (l) => t("incomeAfterTax", l),
      colour: "--fig-after-tax",
      mark: "line",
      get: (r) => r.netto / per,
    },
  ];

  const first = rows[0]?.age ?? view.par;
  const lastAge = rows[rows.length - 1]?.age ?? first;
  // "Figur 2." is dropped on request, same as Figur 1 and Table 2.
  const title = dropHeadingNumber(t("figure2", lang)).replace(
    /\d+\s*-\s*\d+/,
    `${first} - ${lastAge}`,
  );
  const subtitle = view.priceBasis === 1 ? t("fixedPrices", lang) : undefined;

  const svg = newSvg(title);
  const { centre, width, step } = columns(rows.length);
  const max = Math.max(
    ...rows.map((row) => bands.reduce((sum, s) => sum + Math.max(s.get(row), 0), 0)),
    ...rows.flatMap((row) => lines.map((s) => s.get(row))),
    1,
  );
  const axis = vertical(max);
  const y = axis.y;

  gridlines(svg, axis, lang);
  shadeRetirement(svg, retirementEdge(rows, view.par, centre, step));
  svg.append(label(t("earningsAndPension", lang), PLOT.left - 56, PLOT.top - 10, "axis-title"));
  stack(svg, rows, bands, y, centre, width);
  overlay(svg, rows, lines[0]!, y, centre);
  ageTicks(svg, rows, centre);

  const series = [...bands, ...lines];
  const legendSeries = visibleSeries(rows, [...bands].reverse().concat(lines));
  const figure = frame(title, subtitle, svg, legend(legendSeries, lang));
  hover(
    svg,
    figure,
    rows.map((row, i) => ({ x: centre(i), heading: `${t("age", lang)} ${row.age} · ${row.year}`, row })),
    series,
    lang,
  );
  return figure;
}

/**
 * "Diagram 2 i start": gross and net, with the disposable income over them.
 *
 * `Data_till_Start!I` is the gross *above* the net -- the cell reads
 * `brutto - netto` -- so the two bands stack to the gross income, and column K
 * is drawn as a line on top.
 */
export function renderDisposable(result: TypfallResult, view: FigureView): HTMLElement {
  const { lang } = view;
  const per = scale(view);
  const rows = aroundRetirement(result, view.par);

  const bands: readonly Series[] = [
    {
      key: "net",
      name: (l) => t("netIncomeChart", l),
      colour: "--fig-net",
      mark: "fill",
      get: (r) => r.netto / per,
    },
    {
      key: "gross",
      name: (l) => t("grossIncomeChart", l),
      colour: "--fig-gross",
      mark: "fill",
      get: (r) => (r.brutto - r.netto) / per,
    },
  ];
  const line: Series = {
    key: "disposable",
    name: (l) => t("disposable", l),
    colour: "--fig-disposable",
    mark: "line",
    get: (r) => r.indDisp / per,
  };

  const title = t("disposable", lang);
  const svg = newSvg(title);
  const { centre, width, step } = columns(rows.length);
  const max = Math.max(...rows.map((row) => row.brutto / per), 1);
  const axis = vertical(max);
  const y = axis.y;

  gridlines(svg, axis, lang);
  shadeRetirement(svg, retirementEdge(rows, view.par, centre, step));
  stack(svg, rows, bands, y, centre, width);
  overlay(svg, rows, line, y, centre);
  // The workbook marks this series with diamonds; at one per age they stay
  // separable without reading the line itself.
  rows.forEach((row, i) => {
    const cx = centre(i);
    const cy = y(line.get(row));
    svg.append(
      el("path", {
        d: `M ${cx} ${cy - 4} L ${cx + 4} ${cy} L ${cx} ${cy + 4} L ${cx - 4} ${cy} Z`,
        fill: `var(${line.colour})`,
        class: "marker",
      }),
    );
  });
  ageTicks(svg, rows, centre);

  const series = [...bands, line];
  const legendSeries = visibleSeries(rows, [bands[1]!, bands[0]!, line]);
  const figure = frame(title, undefined, svg, legend(legendSeries, lang));
  hover(
    svg,
    figure,
    rows.map((row, i) => ({ x: centre(i), heading: `${t("age", lang)} ${row.age} · ${row.year}`, row })),
    series,
    lang,
  );
  return figure;
}

/**
 * "Skatt per år" / "Tax per year": not a workbook figure -- see the header
 * comment above. Kommunal and statlig skatt are computed the same way for
 * every age, working or retired, so unlike a payslip-style breakdown that
 * separates salary withholding from pension withholding, this shows the true
 * two-part split across the whole window rather than inventing a third,
 * merged category for the working years the engine doesn't actually have.
 */
export function renderTaxChart(result: TypfallResult, view: FigureView): HTMLElement {
  const { lang } = view;
  const per = scale(view);
  const rows = aroundRetirement(result, view.par);

  const bands: readonly Series[] = [
    {
      key: "municipal-tax",
      name: (l) => (l === "sv" ? "Kommunal skatt" : "Municipal tax"),
      colour: "--fig-municipal-tax",
      mark: "fill",
      get: (r) => r.municipalTax / per,
    },
    {
      key: "state-tax",
      name: (l) => (l === "sv" ? "Statlig skatt" : "State tax"),
      colour: "--fig-state-tax",
      mark: "fill",
      get: (r) => r.stateTax / per,
    },
  ];

  // No "(SEK)" suffix, matching every other figure here -- the kronor-
  // formatted y-axis ticks already say what unit this is.
  const title = lang === "sv" ? "Skatt per år" : "Tax per year";
  const subtitle = view.priceBasis === 1 ? t("fixedPrices", lang) : undefined;
  const note =
    lang === "sv"
      ? "Statlig skatt inkluderar public service-avgiften och eventuell kapitalskatt."
      : "State tax includes the public-service fee and any capital-gains tax.";

  const svg = newSvg(title);
  const { centre, width, step } = columns(rows.length);
  const max = Math.max(
    ...rows.map((row) => bands.reduce((sum, s) => sum + Math.max(s.get(row), 0), 0)),
    1,
  );
  const axis = vertical(max);
  const y = axis.y;

  gridlines(svg, axis, lang);
  shadeRetirement(svg, retirementEdge(rows, view.par, centre, step));
  stack(svg, rows, bands, y, centre, width);
  ageTicks(svg, rows, centre);

  const legendSeries = visibleSeries(rows, bands);
  const figure = frame(title, subtitle, svg, legend(legendSeries, lang), [note]);
  hover(
    svg,
    figure,
    rows.map((row, i) => ({ x: centre(i), heading: `${t("age", lang)} ${row.age} · ${row.year}`, row })),
    bands,
    lang,
  );
  return figure;
}
