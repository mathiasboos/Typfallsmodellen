/**
 * "Optimal pensionsålder i förhållande till livslängd": one column per lifespan, the best retirement age
 * for it underneath, what that age pays a month as a bar and what it pays in all as a line.
 *
 * The two measures are of different size (tens of thousands of kronor a month, millions in all), so the
 * bars have the axis on the left and the line the one on the right. That is the one place here where two
 * scales share a plot, and it is drawn to be read as two: each axis has its unit written above it, the
 * legend says which mark is which, the right axis has exactly as many steps as the left so every gridline
 * is a tick on both, both start at zero, and the line has a surface-coloured halo where it crosses the
 * bars. The hover readout and the table beside the chart give every number as a number.
 *
 * Frame, grid, ticks, legend, tooltip and crosshair are the other figures' own (`chart.ts`, the same
 * `.figure`, `.grid`, `.tick`, `.tooltip` and `.crosshair` rules in styles.css). The colours are the
 * page's existing green and blue (`--fig-premium`, `--fig-salary`), which have dark-theme steps already.
 */
import { WIDTH, el, frame, label, legend } from "./chart.js";
import type { SeriesMeta } from "./chart.js";
import type { Lang } from "./i18n.js";
import { niceTicks } from "./pensionAgeChart.js";

const HEIGHT = 360;
const PLOT = { left: 80, right: WIDTH - 58, top: 34, bottom: HEIGHT - 62 } as const;
const BAR = "--fig-premium";
const LINE = "--fig-salary";

export interface OptimalPoint {
  readonly lifespan: number;
  readonly par: number;
  /** Kronor a month; undefined when nothing is paid out by the lifespan. */
  readonly monthly: number | undefined;
  /** Kronor in all. */
  readonly total: number;
}

export interface OptimalChartSpec {
  readonly role: string;
  readonly lang: Lang;
  readonly title: string;
  readonly subtitle: string;
  readonly points: readonly OptimalPoint[];
  readonly text: {
    readonly age: string;
    readonly lifespan: string;
    readonly monthly: string;
    readonly total: string;
    readonly leftUnit: string;
    readonly rightUnit: string;
    readonly heading: (lifespan: number) => string;
  };
  readonly format: {
    readonly kronor: (value: number) => string;
    readonly monthly: (value: number) => string;
    readonly millions: (value: number) => string;
  };
}

/**
 * The right axis: `steps` equal steps from zero, the smallest round ones that reach `max`, so each of its
 * ticks lies on a gridline the left axis already drew.
 */
export function alignedTicks(max: number, steps: number): readonly number[] {
  let step = max / steps || 1;
  const exponent = Math.floor(Math.log10(step));
  search: for (let e = exponent; e <= exponent + 1; e += 1) {
    for (const multiple of [1, 2, 2.5, 5, 10]) {
      const candidate = multiple * 10 ** e;
      if (candidate * steps >= max - 1e-9) {
        step = candidate;
        break search;
      }
    }
  }
  return Array.from({ length: steps + 1 }, (_, i) => i * step);
}

/** A bar with its two top corners rounded and its foot square on the baseline. */
function barPath(x: number, top: number, width: number, bottom: number): string {
  const r = Math.max(0, Math.min(4, (bottom - top) / 2, width / 2));
  return `M${x},${bottom}V${top + r}Q${x},${top} ${x + r},${top}H${x + width - r}Q${x + width},${top} ${x + width},${top + r}V${bottom}Z`;
}

export function renderOptimalChart(spec: OptimalChartSpec): HTMLElement {
  const { lang, points } = spec;
  const svg = el("svg", {
    viewBox: `0 0 ${WIDTH} ${HEIGHT}`,
    role: "img",
    "aria-label": spec.title,
    preserveAspectRatio: "xMidYMid meet",
    tabindex: 0,
  });

  const leftTicks = niceTicks(0, Math.max(1, ...points.map((p) => p.monthly ?? 0)), 5);
  const steps = leftTicks.ticks.length - 1;
  // The right axis is in millions of kronor.
  const rightTicks = alignedTicks(Math.max(1e-9, ...points.map((p) => p.total / 1e6)), steps);
  const yLeft = (v: number) => PLOT.bottom - (v / (leftTicks.hi || 1)) * (PLOT.bottom - PLOT.top);
  const yRight = (kronor: number) => PLOT.bottom - (kronor / 1e6 / (rightTicks[steps] || 1)) * (PLOT.bottom - PLOT.top);
  const slot = (PLOT.right - PLOT.left) / Math.max(points.length, 1);
  const cx = (i: number) => PLOT.left + (i + 0.5) * slot;

  // Gridlines and both axes' numbers; the unit over each axis, in ink.
  leftTicks.ticks.forEach((tick, i) => {
    const py = yLeft(tick);
    svg.append(el("line", { x1: PLOT.left, x2: PLOT.right, y1: py, y2: py, class: tick === 0 ? "axis" : "grid" }));
    svg.append(label(spec.format.kronor(tick), PLOT.left - 8, py + 4, "tick tick-y"));
    const right = label(spec.format.millions(rightTicks[i] ?? 0), PLOT.right + 8, py + 4, "tick");
    right.style.textAnchor = "start";
    svg.append(right);
  });
  svg.append(label(spec.text.leftUnit, 4, PLOT.top - 14, "axis-title"));
  const rightUnit = label(spec.text.rightUnit, WIDTH - 4, PLOT.top - 14, "axis-title");
  rightUnit.style.textAnchor = "end";
  svg.append(rightUnit);

  // Under the baseline, two rows like the table's: the best retirement age, then the lifespan.
  for (const [text, py] of [
    [spec.text.age, PLOT.bottom + 20],
    [spec.text.lifespan, PLOT.bottom + 40],
  ] as const) {
    const row = label(text, 4, py, "tick");
    row.style.textAnchor = "start";
    svg.append(row);
  }
  points.forEach((p, i) => {
    const age = label(String(p.par), cx(i), PLOT.bottom + 20, "tick tick-x pa-opt-age");
    age.style.textAnchor = "middle";
    const span = label(String(p.lifespan), cx(i), PLOT.bottom + 40, "tick tick-x");
    span.style.textAnchor = "middle";
    svg.append(age, span);
  });

  // The bars: what the best age pays a month.
  const width = Math.min(slot * 0.5, 24);
  points.forEach((p, i) => {
    if (p.monthly === undefined || p.monthly <= 0) return;
    const bar = el("path", { d: barPath(cx(i) - width / 2, yLeft(p.monthly), width, PLOT.bottom), fill: `var(${BAR})`, class: "pa-opt-bar" });
    bar.dataset.lifespan = String(p.lifespan);
    svg.append(bar);
  });

  // The line: what it pays in all, with a surface ring so it reads where it crosses a bar.
  const linePoints = points.map((p, i) => `${cx(i)},${yRight(p.total)}`).join(" ");
  if (points.length > 1) {
    svg.append(el("polyline", { points: linePoints, fill: "none", stroke: "var(--surface-raised)", "stroke-width": 5, "stroke-linejoin": "round" }));
    const line = el("polyline", { points: linePoints, fill: "none", stroke: `var(${LINE})`, class: "line pa-line pa-opt-line" });
    svg.append(line);
  }
  points.forEach((p, i) => {
    svg.append(el("circle", { cx: cx(i), cy: yRight(p.total), r: 3.5, fill: `var(${LINE})`, stroke: "var(--surface-raised)", "stroke-width": 2 }));
  });

  const key: SeriesMeta[] = [
    { key: "monthly", name: () => spec.text.monthly, colour: BAR, mark: "fill" },
    { key: "total", name: () => spec.text.total, colour: LINE, mark: "line", dashed: false },
  ];
  const figure = frame(spec.title, spec.subtitle, svg, legend(key, lang));
  figure.dataset.role = spec.role;
  attachHover(svg, figure, spec, cx, slot);
  return figure;
}

/** Crosshair and a readout of the column under the pointer; the arrow keys do the same on the focused chart. */
function attachHover(
  svg: SVGSVGElement,
  figure: HTMLElement,
  spec: OptimalChartSpec,
  cx: (i: number) => number,
  slot: number,
): void {
  const { points } = spec;
  if (points.length === 0) return;
  const cross = el("line", { y1: PLOT.top, y2: PLOT.bottom, x1: 0, x2: 0, class: "crosshair" });
  cross.style.display = "none";
  svg.append(cross);

  const tip = document.createElement("div");
  tip.className = "tooltip";
  tip.hidden = true;
  figure.append(tip);

  let current: number | undefined;
  const hide = () => {
    cross.style.display = "none";
    tip.hidden = true;
    current = undefined;
  };

  const row = (swatchClass: string, colour: string, name: string, value: string): HTMLElement => {
    const line = document.createElement("div");
    line.className = "tooltip-row";
    const swatch = document.createElement("span");
    swatch.className = swatchClass;
    if (swatchClass === "swatch") swatch.style.background = `var(${colour})`;
    else swatch.style.borderTopColor = `var(${colour})`;
    const label = document.createElement("span");
    label.textContent = name;
    const amount = document.createElement("span");
    amount.className = "num";
    amount.textContent = value;
    line.append(swatch, label, amount);
    return line;
  };

  const show = (index: number) => {
    const at = Math.min(Math.max(index, 0), points.length - 1);
    current = at;
    const p = points[at]!;
    const px = cx(at);
    cross.setAttribute("x1", String(px));
    cross.setAttribute("x2", String(px));
    cross.style.display = "";
    tip.replaceChildren();
    const head = document.createElement("strong");
    head.textContent = spec.text.heading(p.lifespan);
    tip.append(
      head,
      row("swatch", "--text-primary", spec.text.age, String(p.par)),
      row("swatch", BAR, spec.text.monthly, p.monthly === undefined ? "–" : spec.format.monthly(p.monthly)),
      row("swatch-line-solid", LINE, spec.text.total, `${spec.format.millions(p.total / 1e6)} Mkr`),
    );
    tip.hidden = false;
    const figureBox = figure.getBoundingClientRect();
    const svgBox = svg.getBoundingClientRect();
    const left = svgBox.left - figureBox.left + (px / WIDTH) * svgBox.width;
    tip.style.left = `${Math.min(Math.max(left, 112), Math.max(figureBox.width - 112, 112))}px`;
  };

  svg.addEventListener("pointermove", (event: PointerEvent) => {
    const box = svg.getBoundingClientRect();
    const inView = ((event.clientX - box.left) / box.width) * WIDTH;
    if (inView < PLOT.left - 12 || inView > PLOT.right + 12) {
      hide();
      return;
    }
    show(Math.floor((inView - PLOT.left) / slot));
  });
  svg.addEventListener("pointerleave", hide);
  svg.addEventListener("focus", () => show(current ?? 0));
  svg.addEventListener("blur", hide);
  svg.addEventListener("keydown", (event: KeyboardEvent) => {
    if (event.key === "ArrowRight") show((current ?? -1) + 1);
    else if (event.key === "ArrowLeft") show((current ?? 1) - 1);
    else if (event.key === "Escape") hide();
    else return;
    event.preventDefault();
  });
}
