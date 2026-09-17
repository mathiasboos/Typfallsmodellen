import { describe, expect, it } from "vitest";

import { defaultContext } from "@typfallsmodellen/engine";

import { GROUPS, SETTINGS } from "../src/advanced.js";
import { LANGS } from "../src/i18n.js";

/**
 * The advanced panel's descriptor table.
 *
 * Its whole job is to be a faithful, typed index into `ModelContext`, so these
 * check the properties that make it one: that each row reads and writes the
 * field it claims, that the panel opens on the workbook's own values rather
 * than a transcription of them, and that the hand-written labels are complete
 * in both languages -- the one place this port does retype the workbook, and
 * so the one place a silent gap could appear.
 */
describe("the advanced settings table", () => {
  const normal = defaultContext();

  it("opens every control on the value the run would use anyway", () => {
    // `get` is what the panel seeds each control with. If it disagreed with
    // `defaultContext()` the form would open showing a number the model was
    // not given -- the thing form.ts's clamp-and-write-back exists to prevent.
    for (const setting of SETTINGS) {
      const value = setting.get(normal);
      expect(Number.isFinite(value), `${setting.key} reads a number`).toBe(true);
    }
  });

  it("writes back into the same field it reads from", () => {
    for (const setting of SETTINGS) {
      const before = setting.get(normal);
      // A flag's only other legal value is its opposite -- several sit over a
      // real `boolean` in ModelContext, and `set(2)` on one of those would read
      // back as 0. Everything else is a count, an amount, a year or a small
      // enumeration, where stepping one up is legal.
      const after = setting.control.kind === "check" ? 1 - before : before + 1;
      const patched = { ...normal, ...setting.set(after) };
      expect(setting.get(patched), `${setting.key} round-trips`).toBe(after);
      // And it must not have moved anything else.
      expect(Object.keys(setting.set(after)).length, `${setting.key} writes one field`).toBe(1);
    }
  });

  it("gives every setting a label in both languages", () => {
    for (const setting of SETTINGS) {
      for (const lang of LANGS) {
        expect(setting.label(lang).trim(), `${setting.key} label ${lang}`).not.toBe("");
        if (setting.hint) {
          expect(setting.hint(lang).trim(), `${setting.key} hint ${lang}`).not.toBe("");
        }
      }
    }
  });

  it("gives every group, and every dropdown choice, both languages too", () => {
    for (const group of GROUPS) {
      for (const lang of LANGS) {
        expect(group.title(lang).trim(), `${group.key} title ${lang}`).not.toBe("");
      }
      for (const setting of group.settings) {
        if (setting.control.kind !== "select") continue;
        for (const choice of setting.control.choices) {
          for (const lang of LANGS) {
            expect(choice.label(lang).trim(), `${setting.key} choice ${choice.value}`).not.toBe("");
          }
        }
      }
    }
  });

  it("says the same thing in Swedish and English, never the same string twice", () => {
    // A label that reads identically in both languages is almost always one
    // that was pasted and not translated. Numbers and units are the exception.
    for (const setting of SETTINGS) {
      const sv = setting.label("sv");
      const en = setting.label("en");
      expect(sv === en, `${setting.key} is untranslated: "${sv}"`).toBe(false);
    }
  });

  it("keys each setting uniquely, and names the Adv_settings row it came from", () => {
    const keys = SETTINGS.map((s) => s.key);
    expect(new Set(keys).size, "keys are unique").toBe(keys.length);
    const rows = SETTINGS.map((s) => s.row);
    expect(new Set(rows).size, "rows are unique").toBe(rows.length);
    for (const setting of SETTINGS) {
      expect(setting.row, `${setting.key} row`).toBeGreaterThan(0);
    }
  });

  it("keeps a dropdown's default among the choices it offers", () => {
    for (const setting of SETTINGS) {
      if (setting.control.kind !== "select") continue;
      const values = setting.control.choices.map((c) => c.value);
      expect(values, `${setting.key}`).toContain(setting.get(normal));
    }
  });

  it("keeps a checkbox's default to 0 or 1", () => {
    for (const setting of SETTINGS) {
      if (setting.control.kind !== "check") continue;
      expect([0, 1], `${setting.key}`).toContain(setting.get(normal));
    }
  });

  it("brackets every typed number's default inside its own bounds", () => {
    for (const setting of SETTINGS) {
      if (setting.control.kind !== "number") continue;
      const value = setting.get(normal);
      expect(value, `${setting.key} min`).toBeGreaterThanOrEqual(setting.control.min);
      expect(value, `${setting.key} max`).toBeLessThanOrEqual(setting.control.max);
    }
  });
});
