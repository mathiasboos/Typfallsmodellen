/**
 * Turns a comparison into something a person can act on.
 *
 * A flat list of failing cases says almost nothing: one agreement wrong across
 * every cohort and one cohort boundary wrong across every agreement look
 * identical in it. So the report groups the same divergences several ways and
 * lets the shape of the failure point at the module that caused it.
 */

import type { CaseComparison, CellStatus } from "./compare.js";
import { blockOf } from "./harness.js";
import type { ComparisonRun } from "./harness.js";

const DIVERGENT: readonly CellStatus[] = ["off", "bad", "missing"];

const isDivergent = (status: CellStatus): boolean => DIVERGENT.includes(status);

export interface ColumnSummary {
  readonly column: number;
  readonly label: string;
  readonly vba: string;
  readonly compared: number;
  readonly matched: number;
  readonly notPorted: boolean;
  readonly worstAbsolute: number;
  readonly worstRelative: number;
  readonly worstCase: number | null;
  readonly roundingSteps: number;
}

type MutableSummary = { -readonly [K in keyof ColumnSummary]: ColumnSummary[K] };

export function summariseColumns(cases: readonly CaseComparison[]): ColumnSummary[] {
  const summaries = new Map<number, MutableSummary>();

  for (const c of cases) {
    for (const cell of c.cells) {
      let summary = summaries.get(cell.column);
      if (summary === undefined) {
        summary = {
          column: cell.column,
          label: cell.label,
          vba: cell.vba,
          compared: 0,
          matched: 0,
          notPorted: cell.status === "not-ported",
          worstAbsolute: 0,
          worstRelative: 0,
          worstCase: null,
          roundingSteps: 0,
        };
        summaries.set(cell.column, summary);
      }
      if (cell.status === "not-ported") continue;

      summary.compared += 1;
      if (cell.status === "exact" || cell.status === "close") summary.matched += 1;
      if (cell.roundingStep) summary.roundingSteps += 1;
      if (Number.isFinite(cell.relative) && cell.relative > summary.worstRelative) {
        summary.worstRelative = cell.relative;
        summary.worstAbsolute = cell.absolute;
        summary.worstCase = c.index;
      }
    }
  }

  return [...summaries.values()].sort((a, b) => a.column - b.column);
}

/** How many cases in a group diverge on at least one column. */
function groupCounts(
  cases: readonly CaseComparison[],
  keyOf: (c: CaseComparison) => string,
): { key: string; total: number; divergent: number }[] {
  const counts = new Map<string, { total: number; divergent: number }>();
  for (const c of cases) {
    const key = keyOf(c);
    const entry = counts.get(key) ?? { total: 0, divergent: 0 };
    entry.total += 1;
    if (c.failed !== undefined || c.cells.some((cell) => isDivergent(cell.status))) {
      entry.divergent += 1;
    }
    counts.set(key, entry);
  }
  return [...counts.entries()]
    .map(([key, value]) => ({ key, ...value }))
    .sort((a, b) => (a.key < b.key ? -1 : a.key > b.key ? 1 : 0));
}

const SCHEME_NAMES = [
  "",
  "ingen",
  "ITP1",
  "ITP2",
  "SAF-LO",
  "KAP-KL",
  "AKAP-KR",
  "PA16 avd 2",
  "PA16 avd 1",
];

const kr = (value: number): string =>
  // `+ 0` so a negative zero prints as 0 rather than -0.
  Number.isFinite(value) ? (value + 0).toLocaleString("en-US", { maximumFractionDigits: 2 }) : "-";

const pct = (value: number): string =>
  Number.isFinite(value) ? `${(value * 100).toFixed(4)}%` : "-";

/** The worst cases, sorted by their largest relative divergence. */
function worstCases(cases: readonly CaseComparison[], limit: number): CaseComparison[] {
  const score = (c: CaseComparison): number => {
    if (c.failed !== undefined) return Number.POSITIVE_INFINITY;
    return Math.max(0, ...c.cells.map((cell) => (Number.isFinite(cell.relative) ? cell.relative : 1e9)));
  };
  return [...cases]
    .filter((c) => c.failed !== undefined || c.cells.some((cell) => isDivergent(cell.status)))
    .sort((a, b) => score(b) - score(a))
    .slice(0, limit);
}

export function formatReport(comparison: ComparisonRun, worstLimit = 20): string {
  const { cases, settings, file } = comparison;
  const out: string[] = [];
  const columns = summariseColumns(cases);
  const comparable = columns.filter((c) => !c.notPorted);
  const failed = cases.filter((c) => c.failed !== undefined);
  const divergentCases = cases.filter(
    (c) => c.failed !== undefined || c.cells.some((cell) => isDivergent(cell.status)),
  );

  out.push("# Golden-file comparison");
  out.push("");
  out.push(`- model: ${settings.modelVersion}`);
  out.push(`- exported: ${settings.exportedAt}`);
  const blocks = comparison.blocks;
  out.push(
    `- cases: ${cases.length}` +
      (blocks.length > 0
        ? ` of the ${comparison.caseSet} set (${blocks.map((b) => `${b.name} ${b.count}`).join(", ")} when complete)`
        : " (no case set recognised, so divergences are not grouped by block)"),
  );
  out.push(`- amounts: ${settings.monthly ? "per month" : "per year"}, Table 1 column D (adjusted)`);
  out.push(`- ran in ${(comparison.elapsedMs / 1000).toFixed(1)}s`);
  out.push("");

  const matched = comparable.reduce((sum, c) => sum + c.matched, 0);
  const compared = comparable.reduce((sum, c) => sum + c.compared, 0);
  out.push(
    `**${matched} of ${compared} comparable cells match** ` +
      `(${comparable.length} of ${columns.length} columns; ` +
      `${divergentCases.length} of ${cases.length} cases diverge).`,
  );
  out.push("");

  if (settings.assumed.length > 0) {
    out.push("## Settings the file did not carry");
    out.push("");
    out.push("The export predates the widened provenance block, so these were assumed:");
    out.push("");
    for (const assumption of settings.assumed) out.push(`- ${assumption}`);
    out.push("");
  }

  const notes = settings.issues.filter((issue) => !issue.fatal);
  if (notes.length > 0) {
    out.push("## Settings that differ from the model's normal values");
    out.push("");
    out.push(
      "Not fatal, but worth knowing: either the export ran with something other than " +
        "the shipped settings, or packages/data/options.json is stale.",
    );
    out.push("");
    for (const issue of notes) out.push(`- ${issue.message}`);
    out.push("");
  }

  if (comparison.labelComplaints.length > 0) {
    out.push("## Column labels that do not match their position");
    out.push("");
    out.push(
      "The column order is the contract (`InputXGetY`'s `Select Case`), not the labels, " +
        "so the comparison proceeded. But a reordered Mikrosim sheet would look exactly " +
        "like this.",
    );
    out.push("");
    for (const complaint of comparison.labelComplaints) out.push(`- ${complaint}`);
    out.push("");
  }

  out.push("## By column");
  out.push("");
  out.push("| # | column | VBA | compared | matched | worst abs | worst rel | case |");
  out.push("|---|---|---|---:|---:|---:|---:|---:|");
  for (const column of columns) {
    if (column.notPorted) {
      out.push(`| ${column.column} | ${column.label} | \`${column.vba}\` | not ported | - | - | - | - |`);
      continue;
    }
    out.push(
      `| ${column.column} | ${column.label} | \`${column.vba}\` | ${column.compared} | ` +
        `${column.matched} | ${kr(column.worstAbsolute)} | ${pct(column.worstRelative)} | ` +
        `${column.worstCase ?? "-"} |`,
    );
  }
  out.push("");

  const unported = columns.filter((c) => c.notPorted);
  if (unported.length > 0) {
    out.push(
      `${unported.length} column(s) have no Table 1 row in the port yet and were not ` +
        `scored: ${unported.map((c) => `${c.column} (${c.label})`).join(", ")}.`,
    );
    out.push("");
  }

  const steps = columns.reduce((sum, c) => sum + c.roundingSteps, 0);
  if (steps > 0) {
    out.push(
      `${steps} divergent cells are a whole number of rounding steps ` +
        `(${settings.monthly ? "1 kr" : "12 kr"}), which points at where the model rounds ` +
        `rather than at what it computes.`,
    );
    out.push("");
  }

  if (divergentCases.length > 0) {
    out.push("## Where the divergences sit");
    out.push("");
    for (const [title, keyOf] of [
      ["block", (c: CaseComparison) => blockOf(c.index, blocks)],
      ["cohort", (c: CaseComparison) => String(c.inputs[0])],
      ["agreement", (c: CaseComparison) => SCHEME_NAMES[c.inputs[9] ?? 0] ?? String(c.inputs[9])],
      ["retirement age", (c: CaseComparison) => String(c.inputs[2])],
    ] as const) {
      const rows = groupCounts(cases, keyOf).filter((r) => r.divergent > 0);
      if (rows.length === 0) continue;
      out.push(`**By ${title}:** ` + rows.map((r) => `${r.key} ${r.divergent}/${r.total}`).join(", "));
      out.push("");
    }
  }

  if (failed.length > 0) {
    out.push("## Cases the engine could not run");
    out.push("");
    for (const c of failed) out.push(`- case ${c.index} (line ${c.line}): ${c.failed}`);
    out.push("");
  }

  const corrected = cases.filter((c) => c.warnings.length > 0);
  if (corrected.length > 0) {
    out.push("## Cases whose input the engine corrected");
    out.push("");
    out.push(
      "The two sides may not have computed the same person here: the workbook corrects " +
        "such inputs through a dialog, the engine through a warning.",
    );
    out.push("");
    for (const c of corrected.slice(0, 20)) {
      const what = c.warnings.map((w) => `${w.field} ${w.given} -> ${w.used}`).join("; ");
      out.push(`- case ${c.index} (line ${c.line}): ${what}`);
    }
    if (corrected.length > 20) out.push(`- ... and ${corrected.length - 20} more`);
    out.push("");
  }

  const worst = worstCases(cases, worstLimit);
  if (worst.length > 0) {
    out.push(`## The ${worst.length} worst cases`);
    out.push("");
    for (const c of worst) {
      const [born, start, retire, salary, , , , , , scheme] = c.inputs;
      out.push(
        `### Case ${c.index} (block ${blockOf(c.index, blocks)}, line ${c.line}) — born ${born}, ` +
          `${start}→${retire}, ${kr(salary ?? 0)} kr/yr, ${SCHEME_NAMES[scheme ?? 0] ?? scheme}`,
      );
      out.push("");
      if (c.failed !== undefined) {
        out.push(`Could not run: ${c.failed}`);
        out.push("");
        continue;
      }
      out.push("| column | expected | engine | diff | rel |");
      out.push("|---|---:|---:|---:|---:|");
      for (const cell of c.cells) {
        if (cell.status === "not-ported") continue;
        const label = isDivergent(cell.status) ? `**${cell.label}**` : cell.label;
        out.push(
          `| ${label} | ${kr(cell.expected)} | ${kr(cell.actual)} | ${kr(cell.absolute)} | ` +
            `${pct(cell.relative)} |`,
        );
      }
      out.push("");
    }
  }

  out.push("## Reading a failure");
  out.push("");
  out.push(
    "One agreement wrong across every cohort points at `src/tjanstepension/`. One cohort " +
      "wrong across every agreement points at a rule-year boundary in `contributions.ts` " +
      "or `atp.ts`. Every column off by the same factor points at `adjustmentFactors` in " +
      "`result.ts` and the price basis. A single block failing points at whatever that " +
      "block varies — see `BuildCases` in the export macro.",
  );
  out.push("");
  out.push(`Source: \`${file.cases.length}\` cases from the committed golden file.`);

  return out.join("\n");
}

/** The one-screen version, for a test failure message or a terminal. */
export function formatSummary(comparison: ComparisonRun): string {
  const columns = summariseColumns(comparison.cases).filter((c) => !c.notPorted);
  const lines = columns.map((column) => {
    const state = column.matched === column.compared ? "ok" : `${column.compared - column.matched} off`;
    return (
      `  ${String(column.column).padStart(2)}  ${column.label.padEnd(22).slice(0, 22)} ` +
      `${String(column.matched).padStart(4)}/${String(column.compared).padEnd(4)} ${state}` +
      (column.worstCase === null || column.matched === column.compared
        ? ""
        : `  worst ${pct(column.worstRelative)} at case ${column.worstCase}`)
    );
  });
  return lines.join("\n");
}
