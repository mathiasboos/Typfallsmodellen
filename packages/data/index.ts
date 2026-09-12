/**
 * Generated data, extracted from the published .xlsb by tools/extract.
 *
 * Do not edit these files by hand. Regenerate them from a new workbook with
 *   python tools/extract/run.py source/Typfallsmodellen.xlsb
 * and review the diff; see docs/YEARLY-UPDATE.md.
 */
import content from "./content.json" with { type: "json" };
import economicSeries from "./economic-series.json" with { type: "json" };
import i18n from "./i18n.json" with { type: "json" };
import inheritanceGains from "./inheritance-gains.json" with { type: "json" };
import manifest from "./manifest.json" with { type: "json" };
import mortality from "./mortality.json" with { type: "json" };
import municipalTax from "./municipal-tax.json" with { type: "json" };
import options from "./options.json" with { type: "json" };
import riksnorm from "./riksnorm.json" with { type: "json" };

export {
  content,
  economicSeries,
  i18n,
  inheritanceGains,
  manifest,
  mortality,
  municipalTax,
  options,
  riksnorm,
};

/** Path of the packed death-probability grid, for loaders that read it as bytes. */
export const mortalityRisksFile = "mortality-risks.bin";
