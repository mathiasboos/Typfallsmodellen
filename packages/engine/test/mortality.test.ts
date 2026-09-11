import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { beforeAll, describe, expect, it } from "vitest";

import { loadDeathProbabilities } from "../src/data/nodeLoader.js";
import { calculateDeltal } from "../src/pension/mortality.js";
import type { AnnuityFactors } from "../src/pension/mortality.js";

/**
 * The workbook's own output from ReadMortality/Calculate_Deltal, written into
 * mortality!P:Z. 17 685 rows covering cohorts 1930-2050, ages 61-105 and all
 * three sex bases -- so the port is checked against the original at every
 * combination the model can reach, without needing Excel.
 */
interface CachedRow {
  sex: number;
  age: number;
  cohort: number;
  dtalip: number;
  dtalpp: number;
  dtalpp2: number;
  expLife: number;
  expLife2: number;
  arvf: number;
  arvf2: number;
}

function loadCachedRows(): CachedRow[] {
  const text = readFileSync(
    fileURLToPath(new URL("../../../reference/fixtures/annuity-factors-cached.csv", import.meta.url)),
    "utf8",
  );
  const [, ...lines] = text.trim().split("\n");
  return lines
    .map((line) => line.split(",").map(Number))
    .filter((f) => f[3]! > 0) // trailing rows the workbook sized for but never wrote
    .map(([, sex, age, cohort, dtalip, dtalpp, dtalpp2, expLife, expLife2, arvf, arvf2]) => ({
      sex: sex!,
      age: age!,
      cohort: cohort!,
      dtalip: dtalip!,
      dtalpp: dtalpp!,
      dtalpp2: dtalpp2!,
      expLife: expLife!,
      expLife2: expLife2!,
      arvf: arvf!,
      arvf2: arvf2!,
    }));
}

const FIELDS = [
  ["dtalip", "dtalip"],
  ["dtalpp", "dtalpp"],
  ["dtalpp2", "dtalpp2"],
  ["expLife", "expLife"],
  ["expLife2", "expLife2"],
  ["arvf", "arvf"],
  ["arvf2", "arvf2"],
] as const;

describe("Calculate_Deltal against the workbook's cached factors", () => {
  let rows: CachedRow[];
  const computed = new Map<number, AnnuityFactors>();

  beforeAll(() => {
    rows = loadCachedRows();
    const deaths = loadDeathProbabilities();
    for (const cohort of new Set(rows.map((r) => r.cohort))) {
      computed.set(cohort, calculateDeltal(cohort, deaths));
    }
  });

  it("has a substantial fixture to check against", () => {
    expect(rows.length).toBeGreaterThan(16_000);
    expect(new Set(rows.map((r) => r.cohort)).size).toBeGreaterThan(100);
  });

  it("reproduces every value exactly", () => {
    const mismatches: string[] = [];
    let compared = 0;

    for (const row of rows) {
      const factors = computed.get(row.cohort)!;
      for (const [fixtureField, factorField] of FIELDS) {
        const want = row[fixtureField];
        const got = factors[factorField][row.sex]![row.age]!;
        compared += 1;
        // Both sides are already rounded -- to two decimals for the delningstal
        // and expected lifetimes, six for the inheritance factors -- so any
        // difference beyond float noise is a real divergence.
        if (Math.abs(got - want) > 5e-7) {
          mismatches.push(
            `cohort ${row.cohort} sex ${row.sex} age ${row.age} ${fixtureField}: got ${got}, workbook has ${want}`,
          );
        }
      }
    }

    expect(compared).toBeGreaterThan(100_000);
    expect(
      `${mismatches.length} mismatches of ${compared}\n${mismatches.slice(0, 10).join("\n")}`,
    ).toBe(`0 mismatches of ${compared}\n`);
  });
});
