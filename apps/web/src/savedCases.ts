/**
 * Sparade typfall: name the typfall on screen, come back to it later.
 *
 * There is no equivalent in the workbook -- a workbook *is* the saved file, so
 * the original never needed this. The web port has the opposite problem: one
 * page, one typfall, and everything you typed gone the moment it reloads. This
 * is genuinely new UI, the same class of addition as `compare.ts`.
 *
 * **What a case holds.** All of it: the Start-sheet form (`TypfallInput`), every
 * Adv_settings override (`Partial<ModelContext>`), the Start-sheet half of
 * Avancerat (`ownIncome`/`pgbManual`/`pgbConscription`), `salaryPath.ts`'s own
 * Deltid periods and grid toggle, whether the riktålder checkbox had taken over
 * the retirement age, and which mode was on. Opening a case therefore
 * reproduces the run exactly, rather than approximately -- which is the whole
 * point, and the reason every panel grew a setter for this (`FormHandle.
 * setValues`, `AdvancedHandle.apply`, `SalaryPathHandle.restore`,
 * `PgbGridHandle.restore`). A case that only carried the nine form fields would
 * quietly drop the work that takes longest to redo.
 *
 * `pgb.ts` needs no state of its own saved: its internal `sicknessByAge`/
 * `studyByAge` maps are exactly what it emits as `pgbManual`, so `restore`
 * rebuilds them from the saved `TypfallInput` half. `salaryPath.ts` does, since
 * a Deltid period list is not recoverable from the income path it generated.
 *
 * **Where it lives.** `localStorage`, under one key holding one JSON document.
 * Not a per-case key: the whole set is read, rewritten and exported together,
 * and one key keeps that atomic and keeps `Ångra` (below) to a single
 * assignment. `localStorage` is the only durable store a page with no outbound
 * requests has -- and it is not always there. Private windows, blocked site
 * data, and Safari on a `file://` page all either throw or silently forget, so
 * every read and write here is wrapped and the panel says so out loud rather
 * than appearing to save and losing the lot. That is also why the file export
 * exists beside it, not merely as a convenience: it is the escape hatch when
 * the browser will not store, when the cases need to reach another device, or
 * when they need to cross between pensionslyft.se's iframe and a downloaded
 * copy of the page, which are separate origins and so separate stores.
 *
 * **Nothing is confirmed with a dialog.** `confirm()` is not reliably available
 * to a page inside someone else's iframe, and a modal is a poor fit for an undo
 * that costs one assignment here. Deleting, overwriting and importing each stash
 * the previous list in `undo` and offer `Ångra` in the status line instead, which
 * is both modal-free and strictly more forgiving than a confirmation would be.
 *
 * **Anything read back is treated as foreign.** `localStorage` is editable by
 * hand and an imported file can be anything at all, so `sanitiseCase` rebuilds
 * each case field by field and drops whatever does not fit, rather than casting
 * the parse result and handing it to `run()`. Ranges are not checked here: the
 * panels' own setters clamp to the bounds their controls already carry, which
 * is the same clamp a typed value goes through and keeps the rule that the form
 * never shows a number the run did not use.
 */
import { defaultContext } from "@typfallsmodellen/engine";
import type { ModelContext, TypfallInput } from "@typfallsmodellen/engine";

import type { Lang } from "./i18n.js";
import type { PartTimePeriod, SalaryPathState } from "./salaryPath.js";

/** The workbook's two radio circles, as `main.ts` names them. */
export type CaseMode = "normal" | "advanced";

/** Everything that makes the page show what it shows, bar the language. */
export interface CaseSnapshot {
  readonly mode: CaseMode;
  readonly input: TypfallInput;
  readonly advanced: Partial<ModelContext>;
  readonly advancedInput: Partial<TypfallInput>;
  readonly salary: SalaryPathState;
  /** Whether the riktålder checkbox had taken over the retirement age. */
  readonly riktalder: boolean;
}

export interface SavedCase extends CaseSnapshot {
  readonly name: string;
  /** ISO 8601, when it was last written. */
  readonly saved: string;
}

/** One key, one JSON document -- see this file's own header on why. */
const KEY = "typfallsmodellen.cases.v1";
/** Bumped only by a change this file cannot read back; `sanitiseCase` handles
 * the rest, so added fields do not need one. */
const FORMAT = 1;
const MAX_CASES = 100;
const MAX_NAME = 80;
export const EXPORT_FILE = "typfall.json";

const say = (l: Lang, sv: string, en: string) => (l === "sv" ? sv : en);

const isRecord = (v: unknown): v is Record<string, unknown> =>
  typeof v === "object" && v !== null && !Array.isArray(v);

const finite = (v: unknown): v is number => typeof v === "number" && Number.isFinite(v);

// ---------------------------------------------------------------- reading --

function sanitiseOwnIncome(raw: unknown): TypfallInput["ownIncome"] | undefined {
  if (!Array.isArray(raw)) return undefined;
  const rows = raw.filter(
    (r): r is { age: number; income: number; wage: number } =>
      isRecord(r) && finite(r.age) && finite(r.income) && finite(r.wage),
  );
  return rows.length > 0 ? rows.map((r) => ({ age: r.age, income: r.income, wage: r.wage })) : undefined;
}

function sanitisePgbManual(raw: unknown): TypfallInput["pgbManual"] | undefined {
  if (!Array.isArray(raw)) return undefined;
  const rows = raw.filter(
    (r): r is { age: number; sa: number; studySemesters: number } =>
      isRecord(r) && finite(r.age) && finite(r.sa) && finite(r.studySemesters),
  );
  return rows.length > 0
    ? rows.map((r) => ({ age: r.age, sa: r.sa, studySemesters: r.studySemesters }))
    : undefined;
}

function sanitiseConscription(raw: unknown): TypfallInput["pgbConscription"] | undefined {
  if (!isRecord(raw)) return undefined;
  const { start, end } = raw;
  const iso = /^\d{4}-\d{2}-\d{2}$/;
  if (typeof start !== "string" || typeof end !== "string") return undefined;
  return iso.test(start) && iso.test(end) ? { start, end } : undefined;
}

/** The nine Start-sheet fields; a case missing any of them is not a case. */
function sanitiseInput(raw: unknown): TypfallInput | undefined {
  if (!isRecord(raw)) return undefined;
  const { born, retirementAge, startWorkAge, monthlySalary, married, scheme } = raw;
  const { yearlyInflation, realGrowth, realReturn } = raw;
  if (
    !finite(born) ||
    !finite(retirementAge) ||
    !finite(startWorkAge) ||
    !finite(monthlySalary) ||
    !finite(scheme) ||
    !finite(yearlyInflation) ||
    !finite(realGrowth) ||
    !finite(realReturn) ||
    typeof married !== "boolean"
  ) {
    return undefined;
  }
  const ownIncome = sanitiseOwnIncome(raw.ownIncome);
  const pgbManual = sanitisePgbManual(raw.pgbManual);
  const pgbConscription = sanitiseConscription(raw.pgbConscription);
  return {
    born,
    retirementAge,
    startWorkAge,
    monthlySalary,
    married,
    scheme: scheme as TypfallInput["scheme"],
    yearlyInflation,
    realGrowth,
    realReturn,
    ...(ownIncome ? { ownIncome } : {}),
    ...(pgbManual ? { pgbManual } : {}),
    ...(pgbConscription ? { pgbConscription } : {}),
  };
}

/**
 * Adv_settings overrides, kept only where they match `defaultContext()`.
 *
 * Checking each key against the shipped context rather than listing sixty
 * fields here means this cannot drift out of step with `ModelContext` -- a
 * field added there is saved and read back with no change to this file, and a
 * field removed there is dropped on read instead of reaching `run()`.
 *
 * `workDuringPartialWithdrawal` is the one field whose type is a union
 * (`number | string`, Adv_settings row 84). Its shipped value is the string, so
 * a saved number would be dropped here and fall back to the default. Nothing
 * can produce one today -- the panel does not expose that row at all -- and the
 * day it does, this is the line to widen.
 */
function sanitiseContext(raw: unknown): Partial<ModelContext> {
  if (!isRecord(raw)) return {};
  const normal = defaultContext() as unknown as Record<string, unknown>;
  const out: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(raw)) {
    if (!Object.hasOwn(normal, key)) continue;
    const expected = normal[key];
    if (Array.isArray(expected)) {
      if (
        Array.isArray(value) &&
        value.length === expected.length &&
        value.every((v, i) => typeof v === typeof expected[i])
      ) {
        out[key] = value;
      }
      continue;
    }
    if (typeof value !== typeof expected) continue;
    if (typeof value === "number" && !Number.isFinite(value)) continue;
    out[key] = value;
  }
  return out as Partial<ModelContext>;
}

/** The Start-sheet half of Avancerat, which is a `TypfallInput` subset. */
function sanitiseAdvancedInput(raw: unknown): Partial<TypfallInput> {
  if (!isRecord(raw)) return {};
  const ownIncome = sanitiseOwnIncome(raw.ownIncome);
  const pgbManual = sanitisePgbManual(raw.pgbManual);
  const pgbConscription = sanitiseConscription(raw.pgbConscription);
  return {
    ...(ownIncome ? { ownIncome } : {}),
    ...(pgbManual ? { pgbManual } : {}),
    ...(pgbConscription ? { pgbConscription } : {}),
  };
}

function sanitiseSalary(raw: unknown): SalaryPathState {
  if (!isRecord(raw)) return { grid: false, edited: false, periods: [] };
  const periods: PartTimePeriod[] = Array.isArray(raw.periods)
    ? raw.periods
        .filter(
          (p): p is { fromYear: number; toYear: number; share: number } =>
            isRecord(p) && finite(p.fromYear) && finite(p.toYear) && finite(p.share),
        )
        .map((p) => ({
          fromYear: p.fromYear,
          toYear: p.toYear,
          share: Math.min(Math.max(p.share, 0), 1),
        }))
    : [];
  return { grid: raw.grid === true, edited: raw.edited === true, periods };
}

/** One stored entry, rebuilt field by field. `undefined` means unreadable. */
export function sanitiseCase(raw: unknown): SavedCase | undefined {
  if (!isRecord(raw)) return undefined;
  const name = typeof raw.name === "string" ? raw.name.trim().slice(0, MAX_NAME) : "";
  if (name === "") return undefined;
  const input = sanitiseInput(raw.input);
  if (input === undefined) return undefined;
  return {
    name,
    saved: typeof raw.saved === "string" ? raw.saved : new Date().toISOString(),
    mode: raw.mode === "advanced" ? "advanced" : "normal",
    input,
    advanced: sanitiseContext(raw.advanced),
    advancedInput: sanitiseAdvancedInput(raw.advancedInput),
    salary: sanitiseSalary(raw.salary),
    riktalder: raw.riktalder === true,
  };
}

/**
 * The cases in a parsed JSON document, in name order.
 *
 * Takes both the `{format, cases}` document this writes and a bare array, so a
 * file someone has pulled a list out of by hand still imports.
 */
export function readDocument(raw: unknown): SavedCase[] {
  const list = Array.isArray(raw) ? raw : isRecord(raw) && Array.isArray(raw.cases) ? raw.cases : [];
  const cases: SavedCase[] = [];
  for (const entry of list) {
    const one = sanitiseCase(entry);
    // Later entries win on a repeated name, the same rule `save` follows.
    if (one) cases.push(one);
  }
  return dedupe(cases);
}

function dedupe(cases: readonly SavedCase[]): SavedCase[] {
  const byName = new Map<string, SavedCase>();
  for (const one of cases) byName.set(one.name, one);
  return [...byName.values()].slice(0, MAX_CASES);
}

export function sortCases(cases: readonly SavedCase[], lang: Lang): SavedCase[] {
  return [...cases].sort((a, b) => a.name.localeCompare(b.name, lang === "sv" ? "sv" : "en"));
}

export function toDocument(cases: readonly SavedCase[]): string {
  return JSON.stringify({ format: FORMAT, cases }, null, 2);
}

// ---------------------------------------------------------------- storage --

/** Whether this browser will actually keep anything for this page. */
export function storageWorks(): boolean {
  try {
    const probe = `${KEY}.probe`;
    localStorage.setItem(probe, "1");
    localStorage.removeItem(probe);
    return true;
  } catch {
    return false;
  }
}

function load(): SavedCase[] {
  try {
    const raw = localStorage.getItem(KEY);
    return raw === null ? [] : readDocument(JSON.parse(raw));
  } catch {
    // Unreadable, hand-edited or blocked -- an empty list, never a throw that
    // would take the whole page down on load.
    return [];
  }
}

function store(cases: readonly SavedCase[]): boolean {
  try {
    localStorage.setItem(KEY, toDocument(cases));
    return true;
  } catch {
    return false;
  }
}

// -------------------------------------------------------------------- UI ---

export interface SavedCasesHandle {
  readonly element: HTMLElement;
  relabel(lang: Lang): void;
}

export function createSavedCases(
  lang: Lang,
  snapshot: () => CaseSnapshot,
  open: (one: SavedCase) => void,
): SavedCasesHandle {
  let currentLang = lang;
  let cases = load();
  let selected = "";
  /** The list as it was before the last destructive change, for `Ångra`. */
  let undo: SavedCase[] | undefined;
  let status: (l: Lang) => string = () => "";

  const element = document.createElement("section");
  element.className = "panel saved-cases";
  element.dataset.panel = "saved-cases";

  const heading = document.createElement("h2");

  const picker = document.createElement("select");
  picker.className = "saved-cases-picker";
  picker.dataset.action = "case-select";

  const nameRow = document.createElement("div");
  nameRow.className = "saved-cases-row";
  const nameField = document.createElement("input");
  nameField.type = "text";
  nameField.className = "saved-cases-name";
  nameField.dataset.action = "case-name";
  nameField.maxLength = MAX_NAME;
  const saveBtn = document.createElement("button");
  saveBtn.type = "button";
  saveBtn.className = "export-btn";
  saveBtn.dataset.action = "case-save";
  const deleteBtn = document.createElement("button");
  deleteBtn.type = "button";
  deleteBtn.className = "export-btn";
  deleteBtn.dataset.action = "case-delete";
  nameRow.append(nameField, saveBtn, deleteBtn);

  const fileRow = document.createElement("div");
  fileRow.className = "saved-cases-row";
  const exportBtn = document.createElement("button");
  exportBtn.type = "button";
  exportBtn.className = "export-btn";
  exportBtn.dataset.file = EXPORT_FILE;
  exportBtn.dataset.action = "case-export";
  const importBtn = document.createElement("button");
  importBtn.type = "button";
  importBtn.className = "export-btn";
  importBtn.dataset.action = "case-import";
  const filePicker = document.createElement("input");
  filePicker.type = "file";
  filePicker.accept = "application/json,.json";
  filePicker.hidden = true;
  fileRow.append(exportBtn, importBtn, filePicker);

  const statusLine = document.createElement("p");
  statusLine.className = "saved-cases-status";
  const statusText = document.createElement("span");
  const undoBtn = document.createElement("button");
  undoBtn.type = "button";
  undoBtn.className = "saved-cases-undo";
  undoBtn.dataset.action = "case-undo";
  undoBtn.hidden = true;
  statusLine.append(statusText, undoBtn);

  element.append(heading, picker, nameRow, fileRow, statusLine);

  /** Replaces the list and rewrites it, reporting a browser that will not. */
  function commit(next: readonly SavedCase[], previous: readonly SavedCase[] | undefined): boolean {
    cases = dedupe(next);
    undo = previous === undefined ? undefined : [...previous];
    if (store(cases)) return true;
    undo = undefined;
    status = (l) =>
      say(
        l,
        "Webbläsaren sparar inget lokalt här — exportera till fil i stället.",
        "This browser will not store anything locally here — export to a file instead.",
      );
    return false;
  }

  function drawPicker(): void {
    const sorted = sortCases(cases, currentLang);
    picker.replaceChildren();
    const blank = document.createElement("option");
    blank.value = "";
    blank.textContent =
      cases.length === 0
        ? say(currentLang, "Inga sparade typfall ännu", "No saved cases yet")
        : say(currentLang, "Välj ett sparat typfall …", "Pick a saved case …");
    picker.append(blank);
    for (const one of sorted) {
      const option = document.createElement("option");
      option.value = one.name;
      option.textContent = one.name;
      picker.append(option);
    }
    picker.value = cases.some((c) => c.name === selected) ? selected : "";
    deleteBtn.disabled = picker.value === "";
    exportBtn.disabled = cases.length === 0;
  }

  function drawStatus(): void {
    statusText.textContent = status(currentLang);
    undoBtn.hidden = undo === undefined;
    statusLine.hidden = statusText.textContent === "" && undo === undefined;
  }

  function draw(): void {
    drawPicker();
    drawStatus();
  }

  picker.addEventListener("change", () => {
    const one = cases.find((c) => c.name === picker.value);
    if (one === undefined) {
      selected = "";
      status = () => "";
      undo = undefined;
      draw();
      return;
    }
    selected = one.name;
    nameField.value = one.name;
    undo = undefined;
    open(one);
    status = (l) => say(l, `Öppnade ”${one.name}”.`, `Opened “${one.name}”.`);
    draw();
  });

  saveBtn.addEventListener("click", () => {
    const name = nameField.value.trim().slice(0, MAX_NAME);
    if (name === "") {
      status = (l) => say(l, "Ge typfallet ett namn först.", "Give the case a name first.");
      undo = undefined;
      draw();
      nameField.focus();
      return;
    }
    const existing = cases.some((c) => c.name === name);
    if (!existing && cases.length >= MAX_CASES) {
      status = (l) =>
        say(
          l,
          `Det går att spara ${MAX_CASES} typfall. Ta bort något först.`,
          `There is room for ${MAX_CASES} saved cases. Delete one first.`,
        );
      undo = undefined;
      draw();
      return;
    }
    const one: SavedCase = { name, saved: new Date().toISOString(), ...snapshot() };
    const before = cases;
    // An overwrite is the one save worth being able to take back; a new name
    // takes nothing away, so it offers no Ångra.
    if (commit([...cases.filter((c) => c.name !== name), one], existing ? before : undefined)) {
      selected = name;
      status = (l) =>
        existing
          ? say(l, `Uppdaterade ”${name}”.`, `Updated “${name}”.`)
          : say(l, `Sparade ”${name}”.`, `Saved “${name}”.`);
    }
    draw();
  });

  deleteBtn.addEventListener("click", () => {
    const name = picker.value;
    if (name === "") {
      status = (l) =>
        say(l, "Välj ett sparat typfall att ta bort.", "Pick a saved case to delete.");
      undo = undefined;
      draw();
      return;
    }
    const before = cases;
    if (commit(cases.filter((c) => c.name !== name), before)) {
      selected = "";
      status = (l) => say(l, `Tog bort ”${name}”.`, `Deleted “${name}”.`);
    }
    draw();
  });

  undoBtn.addEventListener("click", () => {
    if (undo === undefined) return;
    const back = undo;
    if (commit(back, undefined)) {
      status = (l) => say(l, "Ångrade.", "Undone.");
    }
    draw();
  });

  exportBtn.addEventListener("click", () => {
    const blob = new Blob([toDocument(sortCases(cases, currentLang))], {
      type: "application/json;charset=utf-8;",
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = EXPORT_FILE;
    link.click();
    URL.revokeObjectURL(url);
    status = (l) =>
      say(l, `Exporterade ${cases.length} typfall.`, `Exported ${cases.length} cases.`);
    undo = undefined;
    draw();
  });

  importBtn.addEventListener("click", () => filePicker.click());

  filePicker.addEventListener("change", () => {
    const file = filePicker.files?.[0];
    // Cleared straight away, so picking the same file twice still fires.
    filePicker.value = "";
    if (!file) return;
    void file
      .text()
      .then((text) => {
        const incoming = readDocument(JSON.parse(text) as unknown);
        if (incoming.length === 0) {
          status = (l) =>
            say(
              l,
              "Filen innehöll inga typfall som gick att läsa.",
              "The file held no cases that could be read.",
            );
          undo = undefined;
          draw();
          return;
        }
        const before = cases;
        const names = new Set(incoming.map((c) => c.name));
        if (commit([...cases.filter((c) => !names.has(c.name)), ...incoming], before)) {
          status = (l) =>
            say(
              l,
              `Importerade ${incoming.length} typfall.`,
              `Imported ${incoming.length} cases.`,
            );
        }
        draw();
      })
      .catch(() => {
        status = (l) =>
          say(l, "Filen gick inte att läsa som JSON.", "The file could not be read as JSON.");
        undo = undefined;
        draw();
      });
  });

  const applyText = (l: Lang) => {
    heading.textContent = say(l, "Sparade typfall", "Saved cases");
    nameField.placeholder = say(l, "Namn på typfallet", "Name for this case");
    saveBtn.textContent = say(l, "Spara", "Save");
    deleteBtn.textContent = say(l, "Ta bort", "Delete");
    exportBtn.textContent = say(l, "Exportera till fil", "Export to a file");
    importBtn.textContent = say(l, "Importera från fil", "Import from a file");
    undoBtn.textContent = say(l, "Ångra", "Undo");
  };
  applyText(lang);

  // Said once, on the way in, rather than only when a save fails: somebody who
  // cannot save at all should find that out before typing a case, not after.
  if (!storageWorks()) {
    status = (l) =>
      say(
        l,
        "Webbläsaren sparar inget lokalt här (privat läge, blockerad lagring, eller filen öppnad direkt från disk). Exportera till fil i stället.",
        "This browser will not store anything locally here (a private window, blocked site data, or the file opened straight from disk). Export to a file instead.",
      );
  }
  draw();

  return {
    element,
    relabel(l) {
      currentLang = l;
      applyText(l);
      draw();
    },
  };
}
