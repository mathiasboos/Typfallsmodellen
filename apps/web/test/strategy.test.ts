/**
 * "Uttagsstrategi"'s arithmetic, checked without a browser: the runs behind the three scenarios, what
 * the scenarios add up to, the rules that find what most people miss, the recommendation and the
 * checklist. The panel is checked against the built page by `tools/build/verify-offline.mjs`.
 *
 * The anchor numbers are model outputs; `npm run compare` is what keeps the engine honest, and these
 * move with it on purpose.
 */
import { describe, expect, it } from "vitest";

import { Scheme, defaultContext, riktalderFor } from "@typfallsmodellen/engine";
import type { ModelContext } from "@typfallsmodellen/engine";

import { loadDeathProbabilities } from "../src/deaths.js";
import { LIFELONG as LASTS_LIFELONG, defaultLastsInput, normaliseLasts, runPayouts } from "../src/pensionLasts.js";
import {
  JAMKNING_DEADLINES,
  checklistText,
  computeStrategy,
  defaultStrategyInput,
  insightText,
  normaliseStrategy,
  overtakenAt,
  recommend,
  runStrategy,
  evaluateStrategy,
  scenariosToCsv,
  scenariosToXlsxRows,
  seriesToCsv,
  strategyContext,
  strategyRunsKey,
  totalTo,
} from "../src/strategy.js";
import type { ComboRun, Insight, Scenario, StrategyInput } from "../src/strategy.js";

const deaths = loadDeathProbabilities();
const shared = defaultContext() as ModelContext;
/** The day the anchors were taken: the autumn before the jämkning deadline. */
const today = new Date(2026, 9, 30);

const strategy = (patch: Partial<StrategyInput> = {}, day = today) =>
  computeStrategy(normaliseStrategy({ ...defaultStrategyInput(), ...patch }), shared, deaths, day);

const kinds = (insights: readonly Insight[]) => insights.map((i) => i.kind);
const insight = <K extends Insight["kind"]>(insights: readonly Insight[], kind: K) =>
  insights.find((i): i is Extract<Insight, { kind: K }> => i.kind === kind)!;

/** The default: born 1961, the model's salary, SAF-LO, 500 000 kronor in an ISK, the model's rent. */
const plain = strategy();
/** A low income and a higher rent: the housing supplement. */
const low = strategy({ monthlySalary: 24_000, rent: 7_000 });

describe("the input", () => {
  it("is someone a year or two from retirement, with savings, at the riktålder", () => {
    const input = defaultStrategyInput();
    expect(input.born).toBe(1961);
    expect(input.retirementAge).toBe(riktalderFor(1961));
    expect(input.scheme).toBe(Scheme.SafLo);
    expect(input.savingsKind).toBe("isk");
    expect(input.savingsCapital).toBe(500_000);
    expect(input.floor).toBe(17_600);
  });

  it("is kept inside its bounds", () => {
    const input = normaliseStrategy({ ...defaultStrategyInput(), frontYears: 2, floor: -5, rent: 99_999_999, retirementAge: 40 });
    expect(input.frontYears).toBe(5);
    expect(input.floor).toBe(0);
    expect(input.rent).toBe(50_000);
    expect(input.retirementAge).toBeGreaterThanOrEqual(61);
  });

  it("asks the model for each pot's payout period and start separately", () => {
    const input = defaultStrategyInput();
    const combo = { occupational: 10, pension: 15, savings: 5, delay: 1 };
    const pension = strategyContext(shared, input, combo, "pension", [1, 2, 3]);
    expect(pension.tempTjpUttag).toBe(10);
    expect(pension.tempIpsUttag).toBe(15);
    expect(pension.tjpPar).toBe(input.retirementAge + 1);
    expect(pension.privateSavingKind).toBe(0);
    expect(pension.formogenhetPerAge).toEqual([1, 2, 3]);
    expect(pension.hyra).toBe(input.rent);
    expect(pension.priceBasis).toBe(1);
    const savings = strategyContext(shared, input, combo, "savings");
    expect(savings.tempIpsUttag).toBe(5);
    expect(savings.tjpPar).toBe(0);
    expect(savings.privateSavingKind).toBe(2);
    expect(savings.ipsOpening).toBe(500_000);
    expect("formogenhetPerAge" in savings).toBe(false);
    expect(strategyContext(shared, { ...input, housingSupplement: false }, combo, "pension").ansokt).toBe(0);
  });

  it("keeps its runs while only the floor changes", () => {
    const input = defaultStrategyInput();
    expect(strategyRunsKey({ ...input, floor: 1 })).toBe(strategyRunsKey({ ...input, floor: 99_000 }));
    expect(strategyRunsKey({ ...input, rent: 1 })).not.toBe(strategyRunsKey(input));
  });
});

describe("the scenarios", () => {
  it("are balanced, front-loaded and hybrid, starting the year the person turns 67", () => {
    expect(plain.par).toBe(67);
    expect(plain.retirementYear).toBe(2028);
    expect(plain.scenarios.map((s) => s.id)).toEqual(["balanced", "front", "hybrid"]);
    expect(plain.expectedDeath).toBeCloseTo(86.6, 1);
  });

  it("leave what Hur länge räcker pensionen? leaves, when there is no housing supplement to add", () => {
    // Same person, same savings, everything for life, benefits off: the lasts tab's lifelong net, Table 1's first year included.
    const input = normaliseStrategy({ ...defaultStrategyInput(), housingSupplement: false });
    const ours = computeStrategy(input, shared, deaths, today).scenarios[0]!.run;
    const lasts = runPayouts(
      normaliseLasts({ ...defaultLastsInput(), born: input.born, retirementAge: input.retirementAge, monthlySalary: input.monthlySalary, scheme: input.scheme, savingsKind: "isk", savingsCapital: 500_000 }),
      shared,
      deaths,
    ).runs.find((r) => r.payout === LASTS_LIFELONG)!;
    for (let age = ours.par; age <= 100; age += 1) expect(ours.disposable[age]).toBeCloseTo(lasts.net[age]!, 6);
    expect(ours.benefits.every((b) => b === 0)).toBe(true);
  });

  it("add up what there is to live on", () => {
    const balanced = plain.scenarios[0]!;
    const run = balanced.run;
    for (let age = run.par; age <= 100; age += 1) {
      expect(run.disposable[age]).toBeCloseTo(run.net[age]! + run.benefits[age]! + run.savings[age]!, 6);
    }
    const at80 = balanced.totals.find((t) => t.lifespan === 80)!.total;
    expect(at80).toBeCloseTo(totalTo(run, 80), 6);
    expect(balanced.totals.map((t) => t.lifespan)).toEqual([80, 87, 90, 95]);
    expect(balanced.phases.map((p) => [p.from, p.to])).toEqual([
      [67, 76],
      [77, 86],
      [87, 95],
    ]);
  });

  it("put more in the first years when front-loaded, and less after", () => {
    const [balanced, front, hybrid] = plain.scenarios;
    expect(front!.early).toBeGreaterThan(hybrid!.early);
    expect(hybrid!.early).toBeGreaterThan(balanced!.early);
    expect(front!.phases[1]!.monthly).toBeLessThan(balanced!.phases[1]!.monthly);
    // The hybrid keeps the pensions for life, so its lowest month is the balanced one's.
    expect(hybrid!.lowest.monthly).toBeCloseTo(balanced!.lowest.monthly, 6);
  });

  it("recommend the one that pays the most early among those above the floor", () => {
    expect(plain.recommended).toBe("front");
    // A floor only the pensions for life keep: the hybrid, which pays more early than the balanced.
    const front = plain.scenarios.find((s) => s.id === "front")!;
    const between = strategy({ floor: Math.round(front.lowest.monthly) + 50 });
    expect(between.recommended).toBe("hybrid");
    // A floor no scenario keeps: the one with the highest lowest month. The savings are spent by 95 either
    // way, so the hybrid's lowest month is the balanced one's, and the tie goes to the simpler way.
    const above = strategy({ floor: 60_000 });
    expect(above.scenarios.every((s) => !s.meetsFloor)).toBe(true);
    expect(above.recommended).toBe("balanced");
  });

  it("are only the balanced one when there is nothing to draw any other way", () => {
    const bare = strategy({ scheme: Scheme.None, savingsCapital: 0 });
    expect(bare.scenarios.map((s) => s.id)).toEqual(["balanced"]);
    expect(bare.payouts).toEqual([]);
  });

  it("test the housing supplement against the savings as they are spent", () => {
    // Everything for life: the ISK falls year by year, and the housing supplement rises with it.
    const run = low.scenarios[0]!.run;
    expect(run.benefits[run.par + 15]!).toBeGreaterThan(run.benefits[run.par + 2]!);
    expect(run.savingsBalance[run.par + 15]!).toBeLessThan(run.savingsBalance[run.par + 2]!);
  });
});

describe("what most people miss", () => {
  it("finds, for the default, the payout period, the first year's tax and the ISK tax", () => {
    expect(kinds(plain.insights)).toEqual(["payout", "firstYear", "iskTax"]);

    const payout = insight(plain.insights, "payout");
    expect(payout.best).toBe(10);
    expect(payout.worst).toBe(LASTS_LIFELONG);
    expect(payout.lifespan).toBe(87);
    expect(payout.shortestUntil).toBe(77);
    expect(payout.lifelongFrom).toBe(99);
    expect(Math.round(payout.bestTotal - payout.worstTotal)).toBe(219_769);

    const first = insight(plain.insights, "firstYear");
    expect(first.year).toBe(2028);
    expect(first.rateFirst).toBeCloseTo(0.295, 3);
    expect(first.rateNext).toBeCloseTo(0.185, 3);
    expect(Math.round(first.gain / 12)).toBe(3_061);
    expect(first.saving).toBeGreaterThan(5_000);
    expect(first.bridge).toBeGreaterThan(first.saving);

    const isk = insight(plain.insights, "iskTax");
    expect(isk.capital).toBeGreaterThan(300_000);
    expect(isk.perYear).toBeGreaterThan(1_000);
  });

  it("puts the housing supplement first where the savings cost some of it", () => {
    expect(kinds(low.insights)).toEqual(["btpWealth", "payout", "firstYear"]);
    const btp = insight(low.insights, "btpWealth");
    expect(btp.cost).toBeGreaterThan(100_000);
    // 15 % of each 100 000 kronor counts as income, and the supplement falls by most of that.
    expect(btp.perHundred).toBeGreaterThan(600);
    expect(btp.perHundred).toBeLessThan(800);
    const payout = insight(low.insights, "payout");
    expect(payout.best).toBe(5);
    expect(payout.housingDriven).toBe(true);
    expect(payout.lifelongFrom).toBeUndefined();
  });

  it("finds nothing about the housing supplement when it is not counted", () => {
    const off = strategy({ monthlySalary: 24_000, rent: 7_000, housingSupplement: false });
    expect(kinds(off.insights)).not.toContain("btpWealth");
    expect(off.scenarios.every((s) => s.run.benefits.every((b) => b === 0))).toBe(true);
  });

  it("still tells of the first year's tax when there is nothing to move", () => {
    const bare = strategy({ scheme: Scheme.None, savingsCapital: 0 });
    expect(kinds(bare.insights)).toEqual(["firstYear"]);
    const first = insight(bare.insights, "firstYear");
    expect(first.slice).toBe(0);
    expect(first.saving).toBe(0);
    expect(insightText(first, "sv").action).toContain("budgeten för 2028");
  });

  it("reads as text in both languages, with the numbers in it", () => {
    for (const i of plain.insights) {
      const sv = insightText(i, "sv");
      const en = insightText(i, "en");
      expect(sv.title.length).toBeGreaterThan(10);
      expect(en.body).not.toBe(sv.body);
    }
    expect(insightText(insight(plain.insights, "payout"), "sv").body).toContain("Uttag på 5 år ger bara mer om du dör före 77 år.");
    expect(insightText(insight(plain.insights, "firstYear"), "sv").body).toContain("29,5");
  });
});

describe("the checklist", () => {
  it("has at most five steps, minPension first when no deadline is near", () => {
    expect(plain.checklist.length).toBeLessThanOrEqual(5);
    expect(plain.checklist.map((c) => c.kind)).toEqual(["minpension", "payout", "start", "btp", "isk"]);
    const housing = plain.checklist.find((c) => c.kind === "btp")!;
    expect(checklistText(housing, "sv")).toContain("från 77 år");
  });

  it("puts jämkning first in the autumn before the first pension year, with the day it must be in by", () => {
    // Born 1960: the riktålder 67 falls in 2027, so this autumn's jämkning is for 2027.
    const next = strategy({ born: 1960 });
    const first = next.checklist[0]!;
    expect(first).toEqual({ kind: "jamkning", year: 2027, deadline: "8-11-2026", when: "soon" });
    expect(checklistText(first, "sv")).toContain("senast 8 november 2026");
    expect(JAMKNING_DEADLINES[2026]).toEqual({ month: 11, day: 8 });
    // After the day it is still worth doing, but not for the whole year.
    const late = strategy({ born: 1960 }, new Date(2026, 10, 20)).checklist[0]!;
    expect(late).toMatchObject({ kind: "jamkning", when: "late" });
    expect(checklistText(late, "sv")).toContain("så snart som möjligt");
  });
});

describe("the arithmetic behind the comparisons", () => {
  const fake = (disposable: number[], par = 0): ComboRun =>
    ({ par, disposable: [...disposable, ...new Array<number>(101 - disposable.length).fill(0)] }) as unknown as ComboRun;

  it("finds the lifespan from which one way has paid more than another for good", () => {
    // Ahead by 10 for two years, then behind by 4 a year: level after seven years, so not ahead from a
    // lifespan of seven on.
    const early = fake([20, 20, 6, 6, 6, 6, 6, 6]);
    const steady = fake([10, 10, 10, 10, 10, 10, 10, 10]);
    expect(overtakenAt(early, steady)).toBe(7);
    // Never ahead, and still ahead at the end: nothing to report.
    expect(overtakenAt(steady, fake([20, 20, 20]))).toBeUndefined();
    expect(overtakenAt(fake(new Array<number>(101).fill(11)), steady)).toBeUndefined();
  });

  it("recommends the simpler way when the numbers are equal", () => {
    const s = (id: Scenario["id"], early: number, lowest: number, meetsFloor: boolean) =>
      ({ id, early, lowest: { age: 0, monthly: lowest }, meetsFloor }) as unknown as Scenario;
    expect(recommend([s("balanced", 100, 50, true), s("front", 100, 40, true)])).toBe("balanced");
    expect(recommend([s("balanced", 100, 50, true), s("front", 120, 40, false)])).toBe("balanced");
    expect(recommend([s("balanced", 100, 30, false), s("hybrid", 110, 35, false)])).toBe("hybrid");
  });
});

describe("the downloads", () => {
  it("give the scenarios and what there is to live on by age", () => {
    const csv = scenariosToCsv(plain, "sv").split("\r\n");
    expect(csv).toHaveLength(4);
    expect(csv[0]).toContain("Totalt till 87 år");
    expect(csv[2]!.startsWith("Framtung")).toBe(true);
    // Whole kronor without a thousands separator, as in the other tabs' downloads, and the lifespan from
    // which the balanced way has paid more in all: at 100 for the front-loaded one, at 85 for the hybrid.
    expect(csv[1]).toMatch(/^Balanserad \(.*\);\d+;\d+;\d+;\d+;\d+;\d+;\d+;\d+;$/);
    expect(csv[2]!.endsWith(";100")).toBe(true);
    expect(csv[3]!.endsWith(";85")).toBe(true);
    const series = seriesToCsv(plain, "en").split("\r\n");
    expect(series[0]).toBe("Age,Balanced,Front-loaded,Hybrid");
    expect(series[1]).toMatch(/^67,\d+,\d+,\d+$/);
    expect(series).toHaveLength(95 - 67 + 2);
    const sheet = scenariosToXlsxRows(plain, "sv");
    const cell = sheet[1]![1] as { value: number };
    expect(typeof cell.value).toBe("number");
  });

  it("are the same strategy when worked out from runs kept for another floor", () => {
    const input = normaliseStrategy(defaultStrategyInput());
    const runs = runStrategy(input, shared, deaths);
    const again = evaluateStrategy(runs, { ...input, floor: 25_000 }, shared, today);
    expect(again.scenarios[0]!.totals).toEqual(plain.scenarios[0]!.totals);
    expect(again.scenarios.find((s) => s.id === "front")!.meetsFloor).toBe(false);
  });
});
