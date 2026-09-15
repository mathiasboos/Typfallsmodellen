/**
 * The input controls the two panels share.
 *
 * `createForm` grew these as closures over its own `<form>` element and its own
 * list of relabel callbacks. The advanced panel needs the same controls against
 * a different container -- several of them, one per settings group -- so they
 * are a factory now. The bodies are the ones that were in form.ts; what changed
 * is only where they append and which relabel list they register with.
 */
import type { Lang } from "./i18n.js";

export type Relabel = (lang: Lang) => void;

/** What a field's caption and hint read in a given language. */
export interface FieldText {
  readonly label: string;
  readonly hint?: string;
}

/** A number field, and a way to set its value from outside without losing the
 * invalid-input fallback the field keeps for itself. */
export interface NumberField {
  readonly element: HTMLInputElement;
  setValue(v: number): void;
}

/** The same, for a percent field -- `setValue` takes the fraction the field
 * represents (0.324), not the percentage it displays (32.4), matching `apply`. */
export interface PercentField {
  readonly element: HTMLInputElement;
  setValue(v: number): void;
}

export interface PercentOptions {
  /**
   * How many decimal digits `setValue` may show, trimmed of trailing zeros --
   * not padded out to this width, so a round number still reads "35", not
   * "35.000". Defaults to 1, which is every field before this option existed:
   * inflation, growth and return are all round-ish numbers a single digit
   * already describes exactly. A rate read from an outside source can need
   * more -- Tranås's own burial-fee rate is 0.285%, and rounding a picked
   * municipality's real published rate to one decimal (Danderyd's 30.58%
   * becoming "30.6") would show a different number than the one it names.
   */
  readonly maxDecimals?: number;
}

export interface Bounds {
  readonly min: number;
  readonly max: number;
  readonly step: number;
}

export interface FieldSet {
  /**
   * Wraps a control in its caption and hint and appends it.
   *
   * Returns the relabel callback so a linked control can re-run this field's
   * own text without guessing its position in the shared list.
   */
  field(control: HTMLElement, relabel: (l: Lang) => FieldText, className?: string): Relabel;
  number(value: number, bounds: Bounds, apply: (v: number) => void): NumberField;
  percent(value: number, apply: (v: number) => void, options?: PercentOptions): PercentField;
  check(checked: boolean, apply: (v: boolean) => void): HTMLInputElement;
  select(
    choices: readonly { value: number; label: (l: Lang) => string }[],
    selected: number,
    apply: (v: number) => void,
  ): { element: HTMLSelectElement; relabel: Relabel };
}

/** Binds the builders to one container and one relabel list. */
export function fieldSet(container: HTMLElement, relabels: Relabel[], lang: Lang): FieldSet {
  const field: FieldSet["field"] = (control, relabel, className = "field") => {
    const wrap = document.createElement("label");
    wrap.className = className;
    const caption = document.createElement("span");
    caption.className = "field-label";
    const note = document.createElement("span");
    note.className = "field-hint";

    const apply = (l: Lang) => {
      const text = relabel(l);
      caption.textContent = text.label;
      note.textContent = text.hint ?? "";
      note.hidden = text.hint === undefined;
    };
    apply(lang);

    if (className.includes("field-check")) wrap.append(control, caption, note);
    else wrap.append(caption, control, note);
    relabels.push(apply);
    container.append(wrap);
    return apply;
  };

  /**
   * A cell you type a number into.
   *
   * `change` rather than `input`: re-running the model on every keystroke would
   * fire on the half-typed "19" of "1960". The value is clamped to the allowed
   * range and written back, so what the field shows is what the model was given.
   *
   * `setValue` lets a linked control (the riktålder checkbox) update the field
   * the same way a typed change would, so a later invalid entry reverts to that
   * value rather than to whatever was there before the link took over.
   */
  const number: FieldSet["number"] = (value, { min, max, step }, apply) => {
    const el = document.createElement("input");
    el.type = "number";
    el.inputMode = "numeric";
    el.min = String(min);
    el.max = String(max);
    el.step = String(step);
    let current = value;
    const setValue = (v: number) => {
      current = v;
      el.value = String(v);
    };
    setValue(value);
    el.addEventListener("change", () => {
      const typed = Number(el.value);
      if (!Number.isFinite(typed) || el.value.trim() === "") {
        el.value = String(current);
        return;
      }
      const clamped = Math.min(Math.max(typed, min), max);
      setValue(clamped);
      apply(clamped);
    });
    return { element: el, setValue };
  };

  const percent: FieldSet["percent"] = (value, apply, options = {}) => {
    const maxDecimals = options.maxDecimals ?? 1;
    const scale = 10 ** maxDecimals;
    const el = document.createElement("input");
    el.type = "number";
    el.step = String(1 / scale);
    el.className = "percent";
    // `maxDecimals: 1` reduces to `Math.round(v * 1000) / 10` exactly -- the
    // formula every field used before this option existed, unchanged.
    const setValue = (v: number) => {
      el.value = String(Math.round(v * 100 * scale) / scale);
    };
    setValue(value);
    el.addEventListener("change", () => {
      const v = Number(el.value);
      if (Number.isFinite(v)) apply(v / 100);
    });
    return { element: el, setValue };
  };

  const check: FieldSet["check"] = (checked, apply) => {
    const el = document.createElement("input");
    el.type = "checkbox";
    el.checked = checked;
    el.addEventListener("change", () => apply(el.checked));
    return el;
  };

  const select: FieldSet["select"] = (choices, selected, apply) => {
    const el = document.createElement("select");
    const options = choices.map((choice) => {
      const option = document.createElement("option");
      option.value = String(choice.value);
      if (choice.value === selected) option.selected = true;
      el.append(option);
      return { choice, option };
    });
    const relabel = (l: Lang) => {
      for (const { choice, option } of options) option.textContent = choice.label(l);
    };
    relabel(lang);
    el.addEventListener("change", () => apply(Number(el.value)));
    return { element: el, relabel };
  };

  return { field, number, percent, check, select };
}

/** The workbook's own dropdowns, read as the range they cover. */
export const span = (values: readonly number[]): { min: number; max: number } => ({
  min: Math.min(...values),
  max: Math.max(...values),
});

export const rangeHint = ({ min, max }: { min: number; max: number }): string => `${min}–${max}`;
