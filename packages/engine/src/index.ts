/**
 * The Typfallsmodellen engine.
 *
 * A TypeScript port of Pensionsmyndigheten's Excel model, computing the whole
 * Swedish pension outcome for one hypothetical individual. `run` is the entry
 * point; everything else is exported so a caller can reach a single rule.
 *
 * The model's own quirks are preserved deliberately, not accidentally. See
 * docs/VBA-MAPPING.md for the catalogue and the reasoning.
 */

export * from "./model/input.js";
export * from "./model/context.js";
export * from "./model/setup.js";
export * from "./model/state.js";
export * from "./model/mcalc.js";
export * from "./model/drawdown.js";
export * from "./model/taxAndBenefits.js";
export * from "./model/result.js";
export * from "./model/run.js";
export * from "./model/runVectors.js";

export * from "./data/index.js";
export * from "./vba/index.js";
export * from "./income/wages.js";
export * from "./pension/mortality.js";
export * from "./pension/deltal.js";
export * from "./pension/contributions.js";
export * from "./pension/retirementAges.js";
export * from "./pension/incomePension.js";
export * from "./pension/atp.js";
export * from "./tjanstepension/index.js";
export * from "./saving/privateSaving.js";
export * from "./skatt/index.js";
export * from "./bidrag/index.js";
