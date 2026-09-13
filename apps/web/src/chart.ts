/**
 * The two figures, as inline SVG.
 *
 * The workbook draws these with Excel chart objects rather than VBA, and in the
 * distributed file every one of their series references reads `#REFERENS!` --
 * `Workbook_Open` clears the result sheets, which breaks them. So the exact
 * series list cannot be recovered from the file. What can be read is their
 * shape: two line charts of sixteen series each, against a seventeen-column
 * per-age matrix. They plot `mvalues`.
 *
 * Sixteen lines on one pair of axes is not readable, so the same data is shown
 * as the two questions it answers: what the income is made of, and what is left
 * of it. Both are drawn from `result.table2`, which is that matrix with its
 * columns paired the way the Start sheet pairs them.
 *
 * Colours are the validated categorical slots in fixed order -- never cycled,
 * never reassigned when a series drops out. Light and dark are separate steps
 * of the same hues, set as custom properties in styles.css.
 */
import type { Table2Row, TypfallResult } from "@typfallsmodellen/engine";

import { kronor } from "./format.js";
import { t } from "./i18n.js";
import type { Lang, LabelName } from "./i18n.js";

const SVG = "http://www.w3.org/2000/svg";

const WIDTH = 720;
const HEIGHT = 300;
const PAD = { top: 16, right: 16, bottom: 34, left: 64 } as const;

interface Series {
  readonly label: LabelName;
  readonly slot: number;
  readonly get: (row: Table2Row) => number;
}

/** The parts that add up to the gross income, in the order they stack. */
const COMPOSITION: readonly Series[] = [
  { label: "salary", slot: 1, get: (r) => r.salary },
  { label: "incomeAndSupplementary", slot: 2, get: (r) => r.incomeAndSupplementary },
  { label: "premium", slot: 3, get: (r) => r.premium },
  { label: "occupationalPlusIps", slot: 4, get: (r) => r.occupationalAndPrivate },
  { label: "guarantee", slot: 5, get: (r) => r.guaranteeAndSupplement },
];

/** What is left of it, before and after tax and benefits. */
const INCOME_LINES: readonly Series[] = [
  { label: "grossIncome", slot: 1, get: (r) => r.gross },
  { label: "incomeAfterTax", slot: 2, get: (r) => r.net },
  { label: "disposable", slot: 3, get: (r) => r.disposable },
];

function el<K extends keyof SVGElementTagNameMap>(
  name: K,
  attrs: Record<string, string | number> = {},
): SVGElementTagNameMap[K] {
  const node = document.createElementNS(SVG, name);
  for (const [k, v] of Object.entries(attrs)) node.setAttribute(k, String(v));
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

interface Scales {
  x(age: number): number;
  y(value: number): number;
  readonly ages: readonly number[];
  readonly max: number;
}

function scalesFor(rows: readonly Table2Row[], max: number): Scales {
  const ages = rows.map((r) => r.age);
  const first = ages[0] ?? 0;
  const last = ages[ages.length - 1] ?? first + 1;
  const spanX = Math.max(last - first, 1);
  const plotW = WIDTH - PAD.left - PAD.right;
  const plotH = HEIGHT - PAD.top - PAD.bottom;
  const top = max === 0 ? 1 : max;
  return {
    x: (age) => PAD.left + ((age - first) / spanX) * plotW,
    y: (value) => PAD.top + plotH - (value / top) * plotH,
    ages,
    max: top,
  };
}

function axes(svg: SVGSVGElement, s: Scales, lang: Lang): void {
  const step = niceStep(s.max, 4);
  for (let v = 0; v <= s.max + step / 2; v += step) {
    const y = s.y(v);
    svg.append(el("line", { x1: PAD.left, x2: WIDTH - PAD.right, y1: y, y2: y, class: "grid" }));
    const label = el("text", { x: PAD.left - 8, y: y + 4, class: "tick tick-y" });
    label.textContent = kronor(v, lang);
    svg.append(label);
  }

  const first = s.ages[0] ?? 0;
  const last = s.ages[s.ages.length - 1] ?? first;
  const ageStep = Math.max(Math.ceil((last - first) / 8 / 5) * 5, 5);
  for (let age = Math.ceil(first / ageStep) * ageStep; age <= last; age += ageStep) {
    const label = el("text", { x: s.x(age), y: HEIGHT - PAD.bottom + 18, class: "tick tick-x" });
    label.textContent = String(age);
    svg.append(label);
  }

  svg.append(
    el("line", {
      x1: PAD.left,
      x2: WIDTH - PAD.right,
      y1: s.y(0),
      y2: s.y(0),
      class: "axis",
    }),
  );
}

function legend(series: readonly Series[], lang: Lang): HTMLElement {
  const box = document.createElement("ul");
  box.className = "legend";
  for (const item of series) {
    const li = document.createElement("li");
    const swatch = document.createElement("span");
    swatch.className = "swatch";
    swatch.style.background = `var(--series-${item.slot})`;
    const text = document.createElement("span");
    text.textContent = t(item.label, lang);
    li.append(swatch, text);
    box.append(li);
  }
  return box;
}

/** Crosshair and a readout of every series at the nearest age. */
function hover(
  svg: SVGSVGElement,
  figure: HTMLElement,
  rows: readonly Table2Row[],
  series: readonly Series[],
  s: Scales,
  lang: Lang,
): void {
  const line = el("line", {
    y1: PAD.top,
    y2: HEIGHT - PAD.bottom,
    class: "crosshair",
    x1: 0,
    x2: 0,
  });
  line.style.display = "none";
  svg.append(line);

  const tip = document.createElement("div");
  tip.className = "tooltip";
  tip.hidden = true;
  figure.append(tip);

  const move = (event: PointerEvent) => {
    const box = svg.getBoundingClientRect();
    const xInView = ((event.clientX - box.left) / box.width) * WIDTH;
    let nearest = rows[0]!;
    for (const row of rows) {
      if (Math.abs(s.x(row.age) - xInView) < Math.abs(s.x(nearest.age) - xInView)) nearest = row;
    }

    const x = s.x(nearest.age);
    line.setAttribute("x1", String(x));
    line.setAttribute("x2", String(x));
    line.style.display = "";

    tip.replaceChildren();
    const head = document.createElement("strong");
    head.textContent = `${t("age", lang)} ${nearest.age} · ${nearest.year}`;
    tip.append(head);
    for (const item of series) {
      const row = document.createElement("div");
      row.className = "tooltip-row";
      const swatch = document.createElement("span");
      swatch.className = "swatch";
      swatch.style.background = `var(--series-${item.slot})`;
      const name = document.createElement("span");
      name.textContent = t(item.label, lang);
      const value = document.createElement("span");
      value.className = "num";
      value.textContent = kronor(item.get(nearest), lang);
      row.append(swatch, name, value);
      tip.append(row);
    }
    tip.hidden = false;
    const left = (x / WIDTH) * box.width;
    tip.style.left = `${Math.min(Math.max(left, 8), box.width - 8)}px`;
  };

  svg.addEventListener("pointermove", move);
  svg.addEventListener("pointerleave", () => {
    line.style.display = "none";
    tip.hidden = true;
  });
}

function frame(title: string, svg: SVGSVGElement, key: HTMLElement): HTMLElement {
  const figure = document.createElement("figure");
  figure.className = "figure";
  const caption = document.createElement("figcaption");
  caption.textContent = title;
  figure.append(caption, svg, key);
  return figure;
}

function newSvg(title: string): SVGSVGElement {
  const svg = el("svg", {
    viewBox: `0 0 ${WIDTH} ${HEIGHT}`,
    role: "img",
    "aria-label": title,
    preserveAspectRatio: "xMidYMid meet",
  });
  return svg;
}

/** What the income is made of, stacked, age by age. */
export function renderComposition(result: TypfallResult, lang: Lang, title: string): HTMLElement {
  const rows = result.table2;
  const max = Math.max(...rows.map((r) => COMPOSITION.reduce((sum, s) => sum + s.get(r), 0)), 1);
  const s = scalesFor(rows, max);
  const svg = newSvg(title);
  axes(svg, s, lang);

  // Stack from the baseline up, each band drawn over the one below with a 2px
  // surface gap so adjacent fills stay separable without relying on hue alone.
  const running = new Map<number, number>();
  for (const row of rows) running.set(row.age, 0);

  for (const item of COMPOSITION) {
    const upper: string[] = [];
    const lower: string[] = [];
    for (const row of rows) {
      const base = running.get(row.age)!;
      const top = base + item.get(row);
      upper.push(`${s.x(row.age)},${s.y(top)}`);
      lower.push(`${s.x(row.age)},${s.y(base)}`);
      running.set(row.age, top);
    }
    svg.append(
      el("polygon", {
        points: [...upper, ...lower.reverse()].join(" "),
        fill: `var(--series-${item.slot})`,
        class: "band",
      }),
    );
  }

  const figure = frame(title, svg, legend(COMPOSITION, lang));
  hover(svg, figure, rows, COMPOSITION, s, lang);
  return figure;
}

/** Gross, after tax, and disposable, as lines. */
export function renderIncome(result: TypfallResult, lang: Lang, title: string): HTMLElement {
  const rows = result.table2;
  const max = Math.max(...rows.flatMap((r) => INCOME_LINES.map((s) => s.get(r))), 1);
  const s = scalesFor(rows, max);
  const svg = newSvg(title);
  axes(svg, s, lang);

  for (const item of INCOME_LINES) {
    const points = rows.map((row) => `${s.x(row.age)},${s.y(item.get(row))}`).join(" ");
    svg.append(
      el("polyline", { points, fill: "none", stroke: `var(--series-${item.slot})`, class: "line" }),
    );
  }

  const figure = frame(title, svg, legend(INCOME_LINES, lang));
  hover(svg, figure, rows, INCOME_LINES, s, lang);
  return figure;
}
