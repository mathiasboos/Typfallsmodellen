/**
 * Reads the packed death-probability grid in Node.
 *
 * The browser gets its copy over fetch instead; the engine core stays free of
 * both by taking a DeathProbabilities instance.
 */
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import mortalityMeta from "../../../data/mortality.json" with { type: "json" };
import { DeathProbabilities } from "../pension/mortality.js";

export function loadDeathProbabilities(): DeathProbabilities {
  const path = fileURLToPath(new URL("../../../data/mortality-risks.bin", import.meta.url));
  const bytes = readFileSync(path);
  const risks = new Float64Array(
    bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
  );
  if (risks.length !== mortalityMeta.count) {
    throw new Error(
      `mortality-risks.bin holds ${risks.length} values, but mortality.json expects ${mortalityMeta.count}`,
    );
  }
  return DeathProbabilities.fromPacked(risks, mortalityMeta);
}
