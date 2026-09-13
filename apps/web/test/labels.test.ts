/**
 * The UI's two standing risks, checked without a browser.
 *
 * Neither is about how the page looks. The first is that next year's extraction
 * renumbers the SysLang rows and every label after the insertion shifts by one,
 * relabelling the UI with plausible nonsense; the second is that the page and
 * the engine drift apart. The rendered output is checked separately, against
 * the real built file, by `npm run verify:offline`.
 */
import { describe, expect, it } from "vitest";

import snapshot from "../../../reference/fixtures/default-run.json" with { type: "json" };
import { i18n, options } from "@typfallsmodellen/data";
import { defaultInput } from "@typfallsmodellen/engine";

import { L, LANGS, checkLabels, t } from "../src/i18n.js";

describe("the labels the UI reads", () => {
  it("still point at the strings the code names them by", () => {
    // A renumbered SysLang sheet is the failure this catches, and it would
    // otherwise be invisible: every label would still resolve, to the wrong text.
    expect(checkLabels()).toEqual([]);
  });

  it("resolves every key the UI uses, in both languages", () => {
    const missing: string[] = [];
    for (const name of Object.keys(L) as (keyof typeof L)[]) {
      for (const lang of LANGS) {
        if (t(name, lang).trim() === "") missing.push(`${name} (${lang})`);
      }
    }
    expect(missing).toEqual([]);
  });

  it("carries both languages in the extracted data", () => {
    expect(i18n.languages).toEqual(["sv", "en"]);
  });
});

describe("the form's choices", () => {
  it("comes from the workbook rather than from this app", () => {
    // If these ever have to be typed here, the yearly update stops being a
    // data change and becomes a code change.
    expect(options.ranges.birthYears.length).toBeGreaterThan(50);
    expect(options.ranges.retirementAges).toContain(66);
    expect(options.choices.occupationalPension).toHaveLength(8);
  });

  it("opens on the typfall the workbook opens on", () => {
    expect(defaultInput().born).toBe(snapshot.input.born);
    expect(defaultInput().monthlySalary).toBe(snapshot.input.monthlySalary);
    expect(defaultInput().retirementAge).toBe(snapshot.input.retirementAge);
  });
});
