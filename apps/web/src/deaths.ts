/**
 * The death-probability grid, decoded from the bundle.
 *
 * The Node counterpart is packages/engine/src/data/nodeLoader.ts, which reads
 * the same bytes off disk. Neither lives in the engine's public barrel, so the
 * engine core stays free of both a filesystem and a fetch.
 */
import { mortality } from "@typfallsmodellen/data";
import { DeathProbabilities } from "@typfallsmodellen/engine";

import { MORTALITY_RISKS_BASE64 } from "./generated/mortalityRisks.js";

let cached: DeathProbabilities | undefined;

export function loadDeathProbabilities(): DeathProbabilities {
  if (cached) return cached;

  const binary = atob(MORTALITY_RISKS_BASE64);
  // Copy through a fresh buffer: `Float64Array` needs 8-byte alignment, which
  // only an ArrayBuffer allocated here can promise.
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);

  const risks = new Float64Array(bytes.buffer);
  if (risks.length !== mortality.count) {
    throw new Error(
      `the embedded mortality grid holds ${risks.length} values, but mortality.json ` +
        `expects ${mortality.count}. Re-run tools/build/embed-mortality.mjs.`,
    );
  }

  cached = DeathProbabilities.fromPacked(risks, mortality);
  return cached;
}
