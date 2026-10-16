/**
 * "Bästa pensionsålder"'s three line charts, in the frame every other figure
 * here uses (`chart.ts`'s `frame`/`legend`/`newSvg`, the same `.figure`,
 * `.grid`, `.tick`, `.line`, `.tooltip` and `.crosshair` rules in styles.css).
 *
 * They are line charts with up to seven series that differ only in *when they
 * start* -- one per alternative retirement age -- which none of the existing
 * figures draws: `renderFigure1` plots one run's rows against age, and
 * `renderCompareChart` one row set per scenario, all over the same ages. The
 * hover readout is therefore its own (`chart.ts`'s `hover()` prints every
 * series at every point, and here a series has no value before its start).
 *
 * Colours are not categorical hues but a diverging scale around the reference
 * age: warm steps for alternatives earlier than it, cool steps for later ones,
 * and plain ink (`--text-primary`) for the reference itself -- the same
 * "always-legible ink" reasoning `renderMikrosimChart`'s disposable-income
 * line already uses. Seven series cannot be told apart by seven categorical
 * hues (the dataviz method caps a fixed order at eight and says to fold or
 * facet past it), and what the reader needs here is the *order*: how much
 * earlier or later than the reference an alternative is. The six steps are one
 * hue per side stepped in OKLCH lightness (`styles.css`, `--fig-pa-*`), each
 * arm validated as an ordinal ramp against this app's own chart surfaces.
 * Distance from the reference reads as contrast against the surface in both
 * themes (darker in light, lighter in dark), so the extremes stand out and the
 * alternatives next to the reference recede. Identity does not rest on hue
 * alone: every chart carries a legend and the hover readout names every value,
 * and the two tables next to them hold every number plotted.
 *
 * The last point of each line gets a short end label (its retirement age) when
 * the labels can be spread apart without moving any of them far from their own
 * line; otherwise only the two extremes and the reference are labelled, rather
 * than nudging seven labels into a column that no longer sits beside its line.
 */
import { HEIGHT, PLOT, WIDTH, el, frame, label, legend, newSvg } from "./chart.js";
import type { SeriesMeta } from "./chart.js";
import type { Lang } from "./i18n.js";
import type { Slot } from "./pensionAge.js";

/** Room on the right for the end labels the shared `PLOT` leaves none for. */
const RIGHT = WIDTH - 38;
const FRAME = { left: PLOT.left, right: RIGHT, top: PLOT.top, bottom: PLOT.bottom } as const;

export interface ChartSeries {
  readonly slot: Slot;
  /** The retirement age this line is -- a stable handle for tests and the end label. */
  readonly par: number;
  /** Legend and tooltip text. */
  readonly name: string;
  /** Short label beside the last point. */
  readonly end: string;
  /** (x, y) with whole-number x. */
  readonly points: readonly (readonly [number, number])[];
  /** How the tooltip writes a value. */
  readonly format: (value: number) => string;
}

export interface LineChartSpec {
  readonly role: string;
  readonly lang: Lang;
  readonly title: string;
  readonly subtitle?: string;
  readonly note?: string;
  readonly xMin: number;
  readonly xMax: number;
  readonly xTitle: string;
  readonly yTicks: readonly number[];
  readonly yTick: (value: number) => string;
  /** The unit, written above the y axis. */
  readonly yUnit: string;
  /** The tooltip's heading for an x. */
  readonly heading: (x: number) => string;
  readonly series: readonly ChartSeries[];
  /** Mark where each line starts (the charts whose lines begin at different ages). */
  readonly startDots?: boolean;
  /** Short notes inside the plot, at its top-left and bottom-left. */
  readonly annotations?: { readonly top?: string; readonly bottom?: string };
}

// ---------------------------------------------------------------- scales ---

export interface Ticks {
  readonly ticks: readonly number[];
  readonly lo: number;
  readonly hi: number;
}

/**
 * Round tick values covering `[min, max]`: the smallest 1/2/2.5/5 x 10^n step
 * that gives at most about 1.35 x the wanted number of intervals.
 */
export function niceTicks(min: number, max: number, target = 5): Ticks {
  const range = max - min || 1;
  let step = range;
  const exponent = Math.floor(Math.log10(range));
  search: for (let e = exponent - 2; e <= exponent + 1; e += 1) {
    for (const multiple of [1, 2, 2.5, 5]) {
      const candidate = multiple * 10 ** e;
      if (range / candidate <= target * 1.35) {
        step = candidate;
        break search;
      }
    }
  }
  const lo = Math.floor(min / step + 1e-9) * step;
  const hi = Math.ceil(max / step - 1e-9) * step;
  const count = Math.round((hi - lo) / step);
  return { ticks: Array.from({ length: count + 1 }, (_, i) => lo + i * step), lo, hi };
}

/**
 * Moves label positions at least `gap` apart with the least total shift;
 * returns the new positions and the furthest any one of them moved.
 */
function spread(ys: readonly number[], gap: number): { positions: number[]; shift: number } {
  const order = ys.map((y, index) => ({ y, index })).sort((a, b) => a.y - b.y);
  const pos = order.map((o) => o.y);
  for (let pass = 0; pass < 80; pass += 1) {
    let moved = false;
    for (let k = 1; k < pos.length; k += 1) {
      const close = gap - (pos[k]! - pos[k - 1]!);
      if (close > 0.01) {
        pos[k - 1] = pos[k - 1]! - close / 2;
        pos[k] = pos[k]! + close / 2;
        moved = true;
      }
    }
    if (!moved) break;
  }
  const positions = new Array<number>(ys.length).fill(0);
  let shift = 0;
  order.forEach((o, k) => {
    positions[o.index] = pos[k]!;
    shift = Math.max(shift, Math.abs(pos[k]! - o.y));
  });
  return { positions, shift };
}

const cssVar = (slot: Slot): string => `--fig-pa-${slot}`;

// ----------------------------------------------------------------- chart ---

export function renderLineChart(spec: LineChartSpec): HTMLElement {
  const { lang } = spec;
  const svg = newSvg(spec.title);
  svg.setAttribute("tabindex", "0");

  const yLo = spec.yTicks[0] ?? 0;
  const yHi = spec.yTicks[spec.yTicks.length - 1] ?? 1;
  const x = (age: number) => FRAME.left + ((age - spec.xMin) / Math.max(spec.xMax - spec.xMin, 1)) * (FRAME.right - FRAME.left);
  const y = (value: number) => FRAME.bottom - ((value - yLo) / Math.max(yHi - yLo, 1e-9)) * (FRAME.bottom - FRAME.top);

  // Horizontal grid and the y labels; a zero line gets the axis weight.
  for (const tick of spec.yTicks) {
    const py = y(tick);
    svg.append(el("line", { x1: FRAME.left, x2: FRAME.right, y1: py, y2: py, class: tick === 0 ? "axis" : "grid" }));
    svg.append(label(spec.yTick(tick), FRAME.left - 8, py + 4, "tick tick-y"));
  }
  svg.append(el("line", { x1: FRAME.left, x2: FRAME.right, y1: FRAME.bottom, y2: FRAME.bottom, class: "axis" }));
  // Left-aligned at the very edge: right-aligned over the tick labels, "kr per månad" ran off the frame.
  svg.append(label(spec.yUnit, 4, FRAME.top - 10, "axis-title"));

  // X labels: every year is a tick, labelled at a spacing that keeps them apart.
  const span = spec.xMax - spec.xMin;
  const every = span <= 12 ? 1 : span <= 32 ? 5 : 10;
  for (let age = spec.xMin; age <= spec.xMax; age += 1) {
    const major = (age - spec.xMin) % every === 0;
    svg.append(
      el("line", { x1: x(age), x2: x(age), y1: FRAME.bottom, y2: FRAME.bottom + (major ? 5 : 3), class: "axis" }),
    );
    if (major) svg.append(label(String(age), x(age), FRAME.bottom + 18, "tick tick-x"));
  }
  const xTitle = label(spec.xTitle, (FRAME.left + FRAME.right) / 2, HEIGHT - 4, "axis-title");
  xTitle.setAttribute("text-anchor", "middle");
  svg.append(xTitle);

  if (spec.annotations?.top) svg.append(label(spec.annotations.top, FRAME.left + 8, FRAME.top + 14, "tick pa-note"));
  if (spec.annotations?.bottom) svg.append(label(spec.annotations.bottom, FRAME.left + 8, FRAME.bottom - 8, "tick pa-note"));

  // The lines, the reference last so it lies on top of its neighbours.
  const drawn = [...spec.series].sort((a, b) => Number(a.slot === "z") - Number(b.slot === "z"));
  const ends: { series: ChartSeries; y: number }[] = [];
  for (const series of drawn) {
    if (series.points.length === 0) continue;
    const points = series.points.map(([age, value]) => `${x(age)},${y(value)}`).join(" ");
    const line = el("polyline", {
      points,
      fill: "none",
      stroke: `var(${cssVar(series.slot)})`,
      class: series.slot === "z" ? "line pa-line pa-line-z" : "line pa-line",
    });
    line.dataset.par = String(series.par);
    svg.append(line);

    const first = series.points[0]!;
    const last = series.points[series.points.length - 1]!;
    if (spec.startDots) svg.append(dot(x(first[0]), y(first[1]), series.slot));
    svg.append(dot(x(last[0]), y(last[1]), series.slot));
    ends.push({ series, y: y(last[1]) });
  }

  // End labels: all of them if they fit with a small nudge, else the extremes and the reference.
  const sets = [ends, ends.filter((e) => e.series.slot === "m3" || e.series.slot === "z" || e.series.slot === "p3")];
  for (const [attempt, set] of sets.entries()) {
    const { positions, shift } = spread(set.map((e) => e.y), 13);
    if (shift > (attempt === 0 ? 5 : 9)) continue;
    set.forEach((e, i) => svg.append(label(e.series.end, FRAME.right + 9, positions[i]! + 4, "tick pa-end")));
    break;
  }

  const key: SeriesMeta[] = spec.series.map((s) => ({
    key: s.slot,
    name: () => s.name,
    colour: cssVar(s.slot),
    mark: "line",
    dashed: false,
  }));
  const figure = frame(spec.title, spec.subtitle, svg, legend(key, lang), spec.note ? [spec.note] : []);
  figure.dataset.role = spec.role;

  attachHover(svg, figure, spec, x, y);
  return figure;
}

function dot(cx: number, cy: number, slot: Slot): SVGCircleElement {
  return el("circle", { cx, cy, r: 4, fill: `var(${cssVar(slot)})`, stroke: "var(--surface-raised)", "stroke-width": 2 });
}

// ----------------------------------------------------------------- hover ---

/**
 * A crosshair that snaps to whole ages and a readout of every line that has a
 * value there. The pointer is the obvious way in; the arrow keys do the same
 * on the focused chart, so the readout is not mouse-only.
 */
function attachHover(
  svg: SVGSVGElement,
  figure: HTMLElement,
  spec: LineChartSpec,
  x: (age: number) => number,
  y: (value: number) => number,
): void {
  const values = spec.series.map((s) => new Map(s.points.map(([age, value]) => [age, value])));

  const cross = el("line", { y1: FRAME.top, y2: FRAME.bottom, x1: 0, x2: 0, class: "crosshair" });
  cross.style.display = "none";
  const marks = el("g");
  svg.append(cross, marks);

  const tip = document.createElement("div");
  tip.className = "tooltip";
  tip.hidden = true;
  figure.append(tip);

  let current: number | undefined;

  const hide = () => {
    cross.style.display = "none";
    marks.replaceChildren();
    tip.hidden = true;
    current = undefined;
  };

  const show = (age: number) => {
    const at = Math.min(Math.max(Math.round(age), spec.xMin), spec.xMax);
    current = at;
    const px = x(at);
    cross.setAttribute("x1", String(px));
    cross.setAttribute("x2", String(px));
    cross.style.display = "";

    marks.replaceChildren();
    tip.replaceChildren();
    const head = document.createElement("strong");
    head.textContent = spec.heading(at);
    tip.append(head);
    spec.series.forEach((series, i) => {
      const value = values[i]!.get(at);
      if (value === undefined) return;
      const row = document.createElement("div");
      row.className = "tooltip-row";
      const swatch = document.createElement("span");
      swatch.className = "swatch swatch-line-solid";
      swatch.style.borderTopColor = `var(${cssVar(series.slot)})`;
      const name = document.createElement("span");
      name.textContent = series.name;
      const amount = document.createElement("span");
      amount.className = "num";
      amount.textContent = series.format(value);
      row.append(swatch, name, amount);
      tip.append(row);
      marks.append(dot(px, y(value), series.slot));
    });
    tip.hidden = false;

    const figureBox = figure.getBoundingClientRect();
    const svgBox = svg.getBoundingClientRect();
    const left = svgBox.left - figureBox.left + (px / WIDTH) * svgBox.width;
    tip.style.left = `${Math.min(Math.max(left, 112), Math.max(figureBox.width - 112, 112))}px`;
  };

  svg.addEventListener("pointermove", (event: PointerEvent) => {
    const box = svg.getBoundingClientRect();
    const inView = ((event.clientX - box.left) / box.width) * WIDTH;
    if (inView < FRAME.left - 12 || inView > FRAME.right + 12) {
      hide();
      return;
    }
    show(spec.xMin + ((inView - FRAME.left) / (FRAME.right - FRAME.left)) * (spec.xMax - spec.xMin));
  });
  svg.addEventListener("pointerleave", hide);
  svg.addEventListener("focus", () => show(current ?? spec.xMin));
  svg.addEventListener("blur", hide);
  svg.addEventListener("keydown", (event: KeyboardEvent) => {
    if (event.key === "ArrowRight") show((current ?? spec.xMin - 1) + 1);
    else if (event.key === "ArrowLeft") show((current ?? spec.xMin + 1) - 1);
    else if (event.key === "Escape") hide();
    else return;
    event.preventDefault();
  });
}
