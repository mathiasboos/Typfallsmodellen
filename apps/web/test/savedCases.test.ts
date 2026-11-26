import { describe, expect, it } from "vitest";

import { defaultContext, defaultInput } from "@typfallsmodellen/engine";

import { readDocument, sanitiseCase, sortCases, toDocument } from "../src/savedCases.js";
import type { CaseSnapshot, SavedCase } from "../src/savedCases.js";

/**
 * Sparade typfall: what survives the round trip through storage.
 *
 * The panel itself is covered by the offline check (tools/build/
 * verify-offline.mjs), which drives the real thing in a browser. What is worth
 * pinning here is the reading: both `localStorage` and an imported file are
 * editable by hand, so `sanitiseCase` is the boundary between somebody's JSON
 * and `run()`, and it has to drop what does not fit rather than pass it on.
 */
const snapshot = (over: Partial<CaseSnapshot> = {}): CaseSnapshot => ({
  mode: "normal",
  input: defaultInput(),
  advanced: {},
  advancedInput: {},
  salary: { grid: false, edited: false, periods: [] },
  riktalder: false,
  ...over,
});

const saved = (over: Partial<SavedCase> = {}): SavedCase => ({
  name: "Ett typfall",
  saved: "2026-11-26T12:00:00.000Z",
  ...snapshot(),
  ...over,
});

describe("sanitiseCase", () => {
  it("keeps a case this app wrote, unchanged", () => {
    const one = saved();
    expect(sanitiseCase(JSON.parse(JSON.stringify(one)))).toEqual(one);
  });

  it("refuses anything without a name or a usable form", () => {
    expect(sanitiseCase(undefined)).toBeUndefined();
    expect(sanitiseCase("a case")).toBeUndefined();
    expect(sanitiseCase({ ...saved(), name: "   " })).toBeUndefined();
    expect(sanitiseCase({ ...saved(), name: 7 })).toBeUndefined();
    expect(sanitiseCase({ ...saved(), input: undefined })).toBeUndefined();
    // One missing field is enough: a partial typfall is not a typfall.
    const { monthlySalary, ...rest } = defaultInput();
    expect(sanitiseCase({ ...saved(), input: rest })).toBeUndefined();
    // NaN and Infinity read as numbers but are not ones the model can use;
    // JSON cannot carry them, but `localStorage` is not only written by JSON.
    expect(sanitiseCase({ ...saved(), input: { ...defaultInput(), born: Number.NaN } })).toBeUndefined();
  });

  it("trims the name and caps its length", () => {
    expect(sanitiseCase({ ...saved(), name: "  Vårt typfall  " })?.name).toBe("Vårt typfall");
    expect(sanitiseCase({ ...saved(), name: "x".repeat(500) })?.name.length).toBe(80);
  });

  it("falls back on a mode or a flag it does not recognise", () => {
    expect(sanitiseCase({ ...saved(), mode: "sideways" })?.mode).toBe("normal");
    expect(sanitiseCase({ ...saved(), mode: "advanced" })?.mode).toBe("advanced");
    expect(sanitiseCase({ ...saved(), riktalder: "yes" })?.riktalder).toBe(false);
  });

  it("keeps the Adv_settings it knows and drops the rest", () => {
    const one = sanitiseCase({
      ...saved(),
      advanced: {
        hyra: 12_000, //        a real number setting
        ansokt: 1, //           a real flag
        madeUpSetting: 4, //    no such field in ModelContext
        uttagIp: "halva", //    right field, wrong type
        formogenhet: Number.POSITIVE_INFINITY,
      },
    });
    expect(one?.advanced).toEqual({ hyra: 12_000, ansokt: 1 });
  });

  it("keeps a tuple setting only at its own length and type", () => {
    const years = [1991, 1994, 0, 0];
    expect(sanitiseCase({ ...saved(), advanced: { childBirthYears: years } })?.advanced).toEqual({
      childBirthYears: years,
    });
    expect(sanitiseCase({ ...saved(), advanced: { childBirthYears: [1991] } })?.advanced).toEqual({});
    expect(
      sanitiseCase({ ...saved(), advanced: { childBirthYears: ["1991", 0, 0, 0] } })?.advanced,
    ).toEqual({});
  });

  it("reads back every field the shipped context carries", () => {
    // The guard is `defaultContext()` itself, so a field added to ModelContext
    // is saved and read back with no change to savedCases.ts. This walks the
    // whole of it rather than a sample, which is what makes that true.
    const normal = defaultContext() as unknown as Record<string, unknown>;
    const whole = Object.fromEntries(
      Object.entries(normal).filter(([, v]) => typeof v !== "string"),
    );
    expect(sanitiseCase({ ...saved(), advanced: whole })?.advanced).toEqual(whole);
  });

  it("keeps the salary grid's own rows and part-time periods", () => {
    const one = sanitiseCase({
      ...saved(),
      advancedInput: {
        ownIncome: [{ age: 30, income: 300_000, wage: 280_000 }, { age: 31, nope: 1 }],
      },
      salary: {
        grid: true,
        edited: true,
        periods: [
          { fromYear: 1990, toYear: 1999, share: 0.5 },
          { fromYear: 2000, toYear: 2001, share: 4 }, // clamped to full time
          { fromYear: 2002 }, // incomplete, dropped
        ],
      },
    });
    expect(one?.advancedInput.ownIncome).toEqual([{ age: 30, income: 300_000, wage: 280_000 }]);
    expect(one?.salary).toEqual({
      grid: true,
      edited: true,
      periods: [
        { fromYear: 1990, toYear: 1999, share: 0.5 },
        { fromYear: 2000, toYear: 2001, share: 1 },
      ],
    });
  });

  it("keeps a conscription period only when both dates are real ISO dates", () => {
    const period = { start: "1996-08-01", end: "1997-05-30" };
    expect(
      sanitiseCase({ ...saved(), advancedInput: { pgbConscription: period } })?.advancedInput
        .pgbConscription,
    ).toEqual(period);
    expect(
      sanitiseCase({ ...saved(), advancedInput: { pgbConscription: { start: "1996", end: "1997" } } })
        ?.advancedInput.pgbConscription,
    ).toBeUndefined();
  });
});

describe("the stored document", () => {
  it("round-trips through the file it exports", () => {
    const cases = [saved({ name: "A" }), saved({ name: "B", mode: "advanced" })];
    expect(readDocument(JSON.parse(toDocument(cases)))).toEqual(cases);
  });

  it("reads a bare array as well as the document this writes", () => {
    const cases = [saved({ name: "A" })];
    expect(readDocument(cases)).toEqual(cases);
    expect(readDocument({ format: 1, cases })).toEqual(cases);
  });

  it("returns nothing rather than throwing on rubbish", () => {
    expect(readDocument(null)).toEqual([]);
    expect(readDocument("{}")).toEqual([]);
    expect(readDocument({ cases: "lots" })).toEqual([]);
    expect(readDocument({ cases: [1, 2, {}] })).toEqual([]);
  });

  it("keeps the last of a repeated name, the same rule saving follows", () => {
    const first = saved({ name: "A", saved: "2020-01-01T00:00:00.000Z" });
    const second = saved({ name: "A", saved: "2026-01-01T00:00:00.000Z" });
    expect(readDocument([first, second])).toEqual([second]);
  });

  it("drops the unreadable entries and keeps the rest", () => {
    expect(readDocument([saved({ name: "A" }), { name: "" }, saved({ name: "B" })])).toHaveLength(2);
  });

  it("sorts by name in the language on screen", () => {
    const names = ["Östen", "Bo", "Åke", "anna"];
    const cases = names.map((name) => saved({ name }));
    expect(sortCases(cases, "sv").map((c) => c.name)).toEqual(["anna", "Bo", "Åke", "Östen"]);
    // Not a Swedish alphabet: å/ö sort with a/o rather than after z.
    expect(sortCases(cases, "en").map((c) => c.name)).toEqual(["Åke", "anna", "Bo", "Östen"]);
  });
});
