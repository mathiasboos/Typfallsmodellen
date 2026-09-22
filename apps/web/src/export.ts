/**
 * The CSV and Excel download buttons beside a table's own title bar. Not a
 * workbook feature -- there is no SysLang row for either -- so both button
 * texts are plain per-language literals, the same way `compare.ts`'s own new
 * UI text is.
 */
import type { Lang } from "./i18n.js";
import { xlsxBlob } from "./xlsx.js";
import type { SheetData } from "./xlsx.js";

function download(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

function exportButton(text: string): HTMLButtonElement {
  const button = document.createElement("button");
  button.type = "button";
  button.className = "export-btn";
  button.textContent = text;
  return button;
}

export function csvExportButton(lang: Lang, filename: string, csv: () => string): HTMLButtonElement {
  const button = exportButton(lang === "sv" ? "Ladda ner CSV" : "Download CSV");
  button.addEventListener("click", () => {
    // A BOM, so Excel reads å/ä/ö as UTF-8 instead of guessing a legacy
    // codepage from the bytes.
    download(new Blob(["﻿" + csv()], { type: "text/csv;charset=utf-8;" }), filename);
  });
  return button;
}

export function xlsxExportButton(
  lang: Lang,
  filename: string,
  sheetData: () => SheetData,
): HTMLButtonElement {
  const button = exportButton(lang === "sv" ? "Ladda ner Excel" : "Download Excel");
  button.addEventListener("click", () => {
    void xlsxBlob(sheetData()).then((blob) => download(blob, filename));
  });
  return button;
}
