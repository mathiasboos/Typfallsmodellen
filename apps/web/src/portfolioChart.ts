/**
 * The Portföljvärde tab's two charts, in the frame the other figures use (`chart.ts`).
 *
 * - **Stacked columns**, one per year: each account's closing balance on top of the
 *   last, or the year's flows -- what came in (contribution, return) above zero and
 *   what went out (withdrawal, yield tax, fee) below it, each side stacked on its own.
 * - **An area**, for one account's balance on its own.
 *
 * The planner draws them with Chart.js; this draws them as the page's other figures
 * are drawn, as SVG, with the same pieces: the legend is a switch per series (as in
 * "När ta ut pensionen?"), a crosshair and readout follow the pointer and the arrow
 * keys, the retirement year is marked, and the table under the charts carries every
 * number. Stacked fills are parted by a hairline of the surface colour.
 */
import { WIDTH, el, frame, label, newSvg } from "./chart.js";
import { kronor } from "./format.js";
import type { Lang } from "./i18n.js";
import { niceTicks } from "./pensionAgeChart.js";

const HEIGHT = 330;
const PLOT = { left: 64, right: WIDTH - 12, top: 30, bottom: HEIGHT - 40 } as const;

export interface PortfolioSeries {
  /** Stable across redraws and languages: the handle the legend switch keeps. */
  readonly key: string;
  readonly name: string;
  /** A custom property in styles.css. */
  readonly colour: string;
  /** One per year; a negative value stacks below zero. */
  readonly values: readonly number[];
}

export interface PortfolioChartSpec {
  readonly role: string;
  readonly lang: Lang;
  readonly title: string;
  readonly subtitle?: string;
  readonly years: readonly { readonly year: number; readonly age: number }[];
  readonly retirementAge: number;
  readonly kind: "stack" | "area";
  readonly series: readonly PortfolioSeries[];
  /** The readout's last line: the sum of what is shown (the stacked balance chart). */
  readonly totalName?: string;
  readonly note?: string;
}

/** "1,5 Mkr", "250 tkr", "0" -- the axis in round units. */
export function axisAmount(value: number, lang: Lang): string {
  const abs = Math.abs(value);
  const sign = value < 0 ? "−" : "";
  if (abs >= 1e6) {
    const digits = abs % 1e6 === 0 ? 0 : (abs * 10) % 1e6 === 0 ? 1 : 2;
    return `${sign}${new Intl.NumberFormat(lang === "sv" ? "sv-SE" : "en-GB", { minimumFractionDigits: digits, maximumFractionDigits: digits }).format(abs / 1e6)} ${lang === "sv" ? "Mkr" : "MSEK"}`;
  }
  if (abs >= 1e3) return `${sign}${kronor(abs / 1e3, lang)} ${lang === "sv" ? "tkr" : "kSEK"}`;
  return `${sign}${kronor(abs, lang)}`;
}

/** The series each chart has switched off, by chart role: kept for as long as the page is open. */
const hiddenKeys = new Map<string, Set<string>>();

function hiddenFor(role: string, keys: readonly string[]): Set<string> {
  const kept = new Set([...(hiddenKeys.get(role) ?? [])].filter((key) => keys.includes(key)));
  if (kept.size >= keys.length) kept.clear();
  hiddenKeys.set(role, kept);
  return kept;
}

function switchLegend(spec: PortfolioChartSpec, off: ReadonlySet<string>, toggle: (key: string) => void): HTMLElement {
  const box = document.createElement("ul");
  box.className = "legend";
  const sv = spec.lang === "sv";
  for (const series of spec.series) {
    const item = document.createElement("li");
    const swatch = document.createElement("span");
    swatch.className = "swatch";
    swatch.style.background = `var(${series.colour})`;
    const text = document.createElement("span");
    text.className = "legend-name";
    text.textContent = series.name;
    if (spec.series.length < 2) {
      item.append(swatch, text);
    } else {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "legend-toggle";
      button.dataset.series = series.key;
      button.setAttribute("aria-pressed", String(!off.has(series.key)));
      button.title = off.has(series.key) ? (sv ? "Visa serien" : "Show the series") : sv ? "Dölj serien" : "Hide the series";
      button.append(swatch, text);
      button.addEventListener("click", () => toggle(series.key));
      item.append(button);
    }
    box.append(item);
  }
  return box;
}

export function renderPortfolioChart(spec: PortfolioChartSpec): HTMLElement {
  const { lang, years } = spec;
  const keys = spec.series.map((s) => s.key);
  const off = hiddenFor(spec.role, keys);
  const shown = spec.series.filter((s) => !off.has(s.key));

  const svg = newSvg(spec.title);
  svg.setAttribute("viewBox", `0 0 ${WIDTH} ${HEIGHT}`);
  svg.setAttribute("tabindex", "0");

  // What each year reaches up to and down to, over what is shown.
  const n = years.length;
  const ups = years.map((_, i) => shown.reduce((s, series) => s + Math.max(series.values[i] ?? 0, 0), 0));
  const downs = years.map((_, i) => shown.reduce((s, series) => s + Math.min(series.values[i] ?? 0, 0), 0));
  const scale = niceTicks(Math.min(0, ...downs), Math.max(1, ...ups), 5);
  const y = (v: number) => PLOT.bottom - ((v - scale.lo) / Math.max(scale.hi - scale.lo, 1e-9)) * (PLOT.bottom - PLOT.top);
  const slot = (PLOT.right - PLOT.left) / Math.max(n, 1);
  const cx = (i: number) => PLOT.left + (i + 0.5) * slot;

  // The years from retirement on are shaded, as the workbook's own column charts shade them.
  const retiredAt = years.findIndex((yr) => yr.age >= spec.retirementAge);
  if (retiredAt >= 0) {
    svg.append(el("rect", { x: PLOT.left + retiredAt * slot, y: PLOT.top, width: PLOT.right - PLOT.left - retiredAt * slot, height: PLOT.bottom - PLOT.top, class: "retired" }));
  }

  for (const tick of scale.ticks) {
    const py = y(tick);
    svg.append(el("line", { x1: PLOT.left, x2: PLOT.right, y1: py, y2: py, class: tick === 0 ? "axis" : "grid" }));
    svg.append(label(axisAmount(tick, lang), PLOT.left - 8, py + 4, "tick tick-y"));
  }
  svg.append(label(lang === "sv" ? "Kronor" : "SEK", 4, PLOT.top - 12, "axis-title"));

  // Ages along the bottom, every fifth, and the retirement age always.
  years.forEach((yr, i) => {
    if (yr.age % 5 !== 0 && yr.age !== spec.retirementAge) return;
    if (yr.age !== spec.retirementAge && Math.abs(yr.age - spec.retirementAge) < 3) return;
    const tick = label(String(yr.age), cx(i), PLOT.bottom + 16, yr.age === spec.retirementAge ? "tick tick-x pv-ret-tick" : "tick tick-x");
    tick.style.textAnchor = "middle";
    svg.append(tick);
  });
  const axisName = label(lang === "sv" ? "Ålder" : "Age", PLOT.right, PLOT.bottom + 32, "tick");
  axisName.style.textAnchor = "end";
  svg.append(axisName);

  if (spec.kind === "stack") {
    const width = Math.max(1, Math.min(slot * 0.72, 18));
    years.forEach((_, i) => {
      let up = 0;
      let down = 0;
      for (const series of shown) {
        const v = series.values[i] ?? 0;
        if (Math.abs(v) < 0.5) continue;
        const from = v > 0 ? up : down;
        const to = from + v;
        if (v > 0) up = to;
        else down = to;
        const top = y(Math.max(from, to));
        const bottom = y(Math.min(from, to));
        const bar = el("rect", {
          x: cx(i) - width / 2,
          y: top,
          width,
          height: Math.max(bottom - top, 0.5),
          fill: `var(${series.colour})`,
          class: "pv-bar",
        });
        bar.dataset.series = series.key;
        svg.append(bar);
      }
    });
  } else {
    for (const series of shown) {
      const points = years.map((_, i) => `${cx(i)},${y(series.values[i] ?? 0)}`);
      if (points.length === 0) continue;
      const area = el("path", {
        d: `M${cx(0)},${y(0)}L${points.join("L")}L${cx(n - 1)},${y(0)}Z`,
        fill: `var(${series.colour})`,
        class: "pv-area",
      });
      const line = el("polyline", { points: points.join(" "), fill: "none", stroke: `var(${series.colour})`, class: "line pv-line" });
      svg.append(area, line);
    }
  }

  // The retirement year: a dashed rule through the plot.
  if (retiredAt >= 0) {
    const px = cx(retiredAt);
    svg.append(el("line", { x1: px, x2: px, y1: PLOT.top, y2: PLOT.bottom, class: "pv-ret-line" }));
    const tag = label(lang === "sv" ? `Pension ${spec.retirementAge}` : `Retirement ${spec.retirementAge}`, px + 4, PLOT.top + 10, "tick pv-ret-tag");
    tag.style.textAnchor = "start";
    svg.append(tag);
  }

  const toggle = (key: string): void => {
    const hidden = hiddenKeys.get(spec.role) ?? new Set<string>();
    if (hidden.has(key)) hidden.delete(key);
    else if (keys.length - hidden.size > 1) hidden.add(key);
    else return;
    hiddenKeys.set(spec.role, hidden);
    const next = renderPortfolioChart(spec);
    figure.replaceWith(next);
    next.querySelector<HTMLElement>(`[data-series="${key}"]`)?.focus();
  };
  const hint =
    spec.series.length < 2
      ? []
      : [lang === "sv" ? "Klicka på en post i teckenförklaringen för att dölja eller visa den." : "Click an entry in the legend to hide or show it."];
  const figure = frame(spec.title, spec.subtitle, svg, switchLegend(spec, off, toggle), [...(spec.note ? [spec.note] : []), ...hint]);
  figure.dataset.role = spec.role;
  attachHover(svg, figure, spec, shown, cx, slot);
  return figure;
}

/** Crosshair and a readout of the year under the pointer; the arrow keys do the same on the focused chart. */
function attachHover(
  svg: SVGSVGElement,
  figure: HTMLElement,
  spec: PortfolioChartSpec,
  shown: readonly PortfolioSeries[],
  cx: (i: number) => number,
  slot: number,
): void {
  const { years, lang } = spec;
  if (years.length === 0) return;
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

  const line = (colour: string | undefined, name: string, value: number): HTMLElement => {
    const row = document.createElement("div");
    row.className = "tooltip-row";
    const swatch = document.createElement("span");
    swatch.className = "swatch";
    if (colour !== undefined) swatch.style.background = `var(${colour})`;
    else swatch.style.visibility = "hidden";
    const text = document.createElement("span");
    text.textContent = name;
    const amount = document.createElement("span");
    amount.className = "num";
    amount.textContent = `${kronor(Math.round(value), lang)} kr`;
    row.append(swatch, text, amount);
    return row;
  };

  const show = (index: number) => {
    const at = Math.min(Math.max(index, 0), years.length - 1);
    current = at;
    const yr = years[at]!;
    const px = cx(at);
    cross.setAttribute("x1", String(px));
    cross.setAttribute("x2", String(px));
    cross.style.display = "";
    tip.replaceChildren();
    const head = document.createElement("strong");
    head.textContent = lang === "sv" ? `${yr.age} år · ${yr.year}` : `Age ${yr.age} · ${yr.year}`;
    tip.append(head);
    for (const series of shown) tip.append(line(series.colour, series.name, series.values[at] ?? 0));
    if (spec.totalName !== undefined && shown.length > 1) {
      const total = line(undefined, spec.totalName, shown.reduce((s, series) => s + (series.values[at] ?? 0), 0));
      total.classList.add("tooltip-total");
      tip.append(total);
    }
    tip.hidden = false;
    const figureBox = figure.getBoundingClientRect();
    const svgBox = svg.getBoundingClientRect();
    const left = svgBox.left - figureBox.left + (px / WIDTH) * svgBox.width;
    tip.style.left = `${Math.min(Math.max(left, 112), Math.max(figureBox.width - 112, 112))}px`;
  };

  svg.addEventListener("pointermove", (event: PointerEvent) => {
    const box = svg.getBoundingClientRect();
    const inView = ((event.clientX - box.left) / box.width) * WIDTH;
    if (inView < PLOT.left - 6 || inView > PLOT.right + 6) {
      hide();
      return;
    }
    show(Math.floor((inView - PLOT.left) / slot));
  });
  svg.addEventListener("pointerleave", hide);
  svg.addEventListener("focus", () => show(current ?? Math.max(years.findIndex((yr) => yr.age >= spec.retirementAge), 0)));
  svg.addEventListener("blur", hide);
  svg.addEventListener("keydown", (event: KeyboardEvent) => {
    if (event.key === "ArrowRight") show((current ?? -1) + 1);
    else if (event.key === "ArrowLeft") show((current ?? 1) - 1);
    else if (event.key === "Home") show(0);
    else if (event.key === "End") show(years.length - 1);
    else if (event.key === "Escape") hide();
    else return;
    event.preventDefault();
  });
}
