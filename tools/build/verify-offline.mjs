/**
 * Opens the built file the way a reader will: from disk, with no network.
 *
 * The delivery promise is "double-click it, no installation, nothing sent
 * anywhere". That is a claim about a browser, so it is checked in one rather
 * than asserted here in prose. Chromium loads dist/typfallsmodellen.html over
 * `file://` with every request that is not the file itself aborted, and the
 * script fails if anything was blocked, if the page logged an error, or if
 * Table 1 does not read back what the engine computed.
 *
 * Playwright is deliberately not a dependency of this repo: this is a check you
 * run before shipping, and making it one would put a browser download in front
 * of every `npm ci` for a workspace whose tests are otherwise pure Node. Run
 * `npm i -D playwright` when you need it.
 *
 *   node tools/build/verify-offline.mjs [--shots <dir>]
 */
import { existsSync, mkdirSync, readFileSync, readdirSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const repo = resolve(here, "../..");
const page = join(repo, "apps/web/dist/typfallsmodellen.html");

const args = process.argv.slice(2);
const shotIndex = args.indexOf("--shots");
const shots = shotIndex === -1 ? undefined : args[shotIndex + 1];

if (!existsSync(page)) {
  fail(`${page} is not there. Build it first: npm run build -w @typfallsmodellen/web`);
}

/**
 * A static check on the file itself, before a browser is even involved:
 * Mobile Safari can fail to run a `<script type="module">` at all when the
 * page is opened from Files/Mail/Messages over `file://` -- reported as a
 * black screen on an iPhone -- so the build produces a classic script
 * instead (`vite.config.ts`'s own comment on why). A classic script has no
 * implicit defer the way a module does, so it also has to sit after `#app`
 * in the document -- `inline-single-file.mjs` moves it there rather than
 * leaving it where Vite placed it, in `<head>`.
 */
const html = readFileSync(page, "utf8");
if (/<script[^>]*type=["']module["']/.test(html)) {
  fail(
    'the built file has a <script type="module"> -- Mobile Safari can fail to run it at all over ' +
      "file://. See vite.config.ts's own comment on why this app uses a classic script instead.",
  );
}
const scriptIndex = html.indexOf("<script>");
const appIndex = html.indexOf('id="app"');
if (scriptIndex === -1) fail("the built file has no plain <script> tag to run the app.");
if (appIndex === -1 || scriptIndex < appIndex) {
  fail(
    "the built file's <script> tag isn't after #app in the document -- a classic script runs the " +
      "instant the parser reaches it, unlike a module script, so it has to come after #app exists.",
  );
}
console.log("script tag      : classic, after #app -- runs on Mobile Safari too");

let chromium;
try {
  ({ chromium } = await import("playwright"));
} catch {
  fail("playwright is not installed. It is not a dependency of this repo -- npm i -D playwright");
}

/**
 * Where Chromium is.
 *
 * Playwright normally knows; on a machine where the browsers live somewhere
 * else -- a prepared container, say -- PLAYWRIGHT_BROWSERS_PATH or an explicit
 * CHROMIUM_PATH says so instead. Guessing quietly and failing with a stack
 * trace three calls later is the thing to avoid.
 */
function browserPath() {
  if (process.env.CHROMIUM_PATH) return process.env.CHROMIUM_PATH;

  const declared = (() => {
    try {
      return chromium.executablePath();
    } catch {
      return undefined;
    }
  })();
  if (declared && existsSync(declared)) return declared;

  const root = process.env.PLAYWRIGHT_BROWSERS_PATH;
  if (root && existsSync(root)) {
    for (const entry of readdirSync(root)) {
      if (!entry.startsWith("chromium")) continue;
      for (const candidate of ["chrome-linux/chrome", "chrome-mac/Chromium.app/Contents/MacOS/Chromium", "chrome-win/chrome.exe"]) {
        const path = join(root, entry, candidate);
        if (existsSync(path)) return path;
      }
    }
  }

  fail(
    `no Chromium found. Tried ${declared ?? "playwright's own path"}` +
      `${root ? ` and ${root}` : ""}. Run npx playwright install chromium, or set CHROMIUM_PATH.`,
  );
}

/**
 * What Table 1 must say.
 *
 * From the engine's own regression snapshot, so this checks the page against
 * the model rather than against numbers typed into this file. Table 1 shows the
 * Start sheet's four columns at once, and each is asserted by name -- the cells
 * carry `data-col`, so this cannot silently start reading a different one.
 */
const fixture = JSON.parse(readFileSync(join(repo, "reference/fixtures/default-run.json"), "utf8"));
const COLUMNS = [
  { col: "nominal", of: (row) => Math.round(row.nominal) },
  { col: "adjusted", of: (row) => Math.round(row.adjusted) },
  { col: "monthly", of: (row) => Math.round(row.monthly) },
  // Column D is a percentage with one decimal, so it is compared at that.
  { col: "share", of: (row) => Math.round(row.shareOfFinalSalary * 1000) / 10 },
];

const expected = ["slutlon", "totBrutto", "efterSkatt"].map((key) => {
  const row = fixture.table1.find((r) => r.key === key);
  if (row === undefined) throw new Error(`default-run.json has no ${key} row`);
  return { key, want: COLUMNS.map((column) => column.of(row)) };
});

// Figur 2 and the disposable income chart cover five years before retirement
// through age 100 (widened on request from ten years either side of
// retirement). retirementAge is fixed at par for the default typfall -- no
// warning corrects it -- so this is exact, not an estimate.
const WINDOW = 100 - (fixture.input.retirementAge - 5) + 1;

const browser = await chromium.launch({ executablePath: browserPath() });
const context = await browser.newContext({ viewport: { width: 1280, height: 1000 } });

// Nothing may leave the page. Any request that is not the file itself fails,
// so a stylesheet, a font or a telemetry beacon that crept in shows up here.
// `blob:` is let through alongside `file://` -- a CSV/Excel download is
// `URL.createObjectURL(blob)` plus `<a download>`, which never leaves the
// page either; Chromium still routes the click through here.
const outbound = [];
await context.route("**/*", (route) => {
  const url = route.request().url();
  if (url.startsWith("file://") || url.startsWith("blob:")) return route.continue();
  outbound.push(url);
  return route.abort();
});

const tab = await context.newPage();
const errors = [];
tab.on("pageerror", (error) => errors.push(String(error)));
tab.on("console", (message) => {
  if (message.type() === "error") errors.push(message.text());
});

await tab.goto(pathToFileURL(page).href);
await tab.waitForSelector(".table1 tbody tr[data-key]", { timeout: 15000 });

// The disclaimer is a disclosure now, not an always-open box -- proven closed
// as loaded, before any other interaction could have opened it.
const problems = [];
const noticeAsLoaded = await tab.locator('[data-role="disclaimer"]').getAttribute("open");
if (noticeAsLoaded !== null) {
  problems.push("the disclaimer is open on load, expected collapsed by default");
}

/** One cell of a Table 1 row, found by the model's key and the column's name. */
async function shown(key, col) {
  const text = await tab
    .locator(`.table1 tbody tr[data-key="${key}"] td[data-col="${col}"]`)
    .textContent();
  if (text === null) return Number.NaN;
  // Intl groups with a narrow no-break space in sv-SE, and writes the decimal
  // comma the percentage column uses.
  return Number(text.replace(/[^\d,-]/g, "").replace(",", "."));
}

/**
 * Clicks a download button and reads back what it actually produced --
 * `Blob`/`<a download>` never touches the network, so `context.route`'s own
 * block-everything rule doesn't cover it, and nothing else here proves the
 * file this triggers is real rather than an empty or broken one.
 */
let downloadCount = 0;
async function download(button) {
  const [saved] = await Promise.all([tab.waitForEvent("download"), button.click()]);
  const path = join(tmpdir(), `verify-offline-${downloadCount++}-${saved.suggestedFilename()}`);
  await saved.saveAs(path);
  return { filename: saved.suggestedFilename(), buffer: readFileSync(path) };
}

/** The ZIP local-file-header signature every `.xlsx` starts with -- proof
 * this is a real OOXML package, not just a file with an `.xlsx` name. */
function looksLikeXlsx(buffer) {
  return buffer.length > 4 && buffer[0] === 0x50 && buffer[1] === 0x4b;
}

for (const { key, want } of expected) {
  const got = [];
  for (const column of COLUMNS) got.push(await shown(key, column.col));
  const wrong = COLUMNS.filter((column, i) => got[i] !== want[i]);
  for (const column of wrong) {
    const i = COLUMNS.indexOf(column);
    problems.push(`${key}.${column.col}: the page shows ${got[i]}, the engine says ${want[i]}`);
  }
  console.log(
    `${wrong.length === 0 ? "OK      " : "MISMATCH"} ${key.padEnd(12)} ` +
      COLUMNS.map((column, i) => `${column.col}=${got[i]}`).join("  "),
  );
}

/**
 * The three KPI cards above Table 1.
 *
 * The first two read straight off the same `totBrutto` row Table 1 was just
 * checked against, so they are checked exactly, against `expected`'s own
 * rounding rather than re-deriving it. The third averages over the whole
 * retirement span, which this fixture only samples every few years (no
 * `throughAge` in it at all, in fact) -- there is no cheap exact figure to
 * check it against here, so it gets a sanity range instead: finite, positive,
 * and not wildly off the retirement-year figure beside it.
 */
async function kpiValue(index) {
  const text = await tab.locator(".kpi-card").nth(index).locator(".kpi-value").textContent();
  if (text === null) return Number.NaN;
  return Number(text.replace(/[^\d,-]/g, "").replace(",", "."));
}

const kpiCards = await tab.locator(".kpi-card").count();
console.log(`KPI cards       : ${kpiCards}`);
if (kpiCards !== 3) problems.push(`${kpiCards} KPI cards rendered, expected 3`);

if (kpiCards === 3) {
  const totBruttoWant = expected.find((e) => e.key === "totBrutto")?.want;
  const [, , wantMonthly, wantShare] = totBruttoWant ?? [];

  const gotMonthly = await kpiValue(0);
  const gotShare = await kpiValue(1);
  const gotAverage = await kpiValue(2);
  console.log(`KPI 1 pension at retirement : ${gotMonthly}`);
  console.log(`KPI 2 replacement rate      : ${gotShare}%`);
  console.log(`KPI 3 average through retirement: ${gotAverage}`);

  if (gotMonthly !== wantMonthly) {
    problems.push(`KPI "pension at retirement": the page shows ${gotMonthly}, the engine says ${wantMonthly}`);
  }
  if (gotShare !== wantShare) {
    problems.push(`KPI "replacement rate": the page shows ${gotShare}%, the engine says ${wantShare}%`);
  }
  if (!(gotAverage > 0 && gotAverage < wantMonthly * 2)) {
    problems.push(
      `KPI "average pension": ${gotAverage} is outside the sane range (0, ${wantMonthly * 2})`,
    );
  }
}

// Table 1's own CSV and Excel downloads. Scoped to the `.panel` that holds
// `table.table1` itself, not just "the first `.panel-actions`" -- the
// Avancerat column's own salary-grid actions row is already in the DOM at
// this point (just hidden until that mode is on), and sits before the
// results column, so it would otherwise win a plain `.first()`.
const table1Actions = tab
  .locator("section.panel", { has: tab.locator("table.table1") })
  .locator(".panel-actions");
const table1Csv = await download(table1Actions.getByRole("button", { name: "Ladda ner CSV" }));
console.log(`table1 CSV      : ${table1Csv.filename}, ${table1Csv.buffer.length} bytes`);
if (!table1Csv.filename.endsWith(".csv") || !table1Csv.buffer.toString("utf8").includes("Slutlön")) {
  problems.push(`Table 1's CSV download ("${table1Csv.filename}") doesn't look like Table 1's own data`);
}
const table1Xlsx = await download(table1Actions.getByRole("button", { name: "Ladda ner Excel" }));
console.log(`table1 Excel    : ${table1Xlsx.filename}, ${table1Xlsx.buffer.length} bytes`);
if (!table1Xlsx.filename.endsWith(".xlsx") || !looksLikeXlsx(table1Xlsx.buffer)) {
  problems.push(`Table 1's Excel download ("${table1Xlsx.filename}") isn't a real .xlsx file`);
}

const table2Rows = await tab.locator(".table2 tbody tr").count();
const figures = await tab.locator("figure.figure").count();
// Figur 2 is the second figure; its stack must cover the whole `A2:A22` window,
// one column of bars per age, or the window has silently moved.
const stacked = await tab.locator("figure.figure").nth(1).locator("svg rect.bar").evaluateAll(
  (nodes) => new Set(nodes.map((node) => node.getAttribute("x"))).size,
);
console.log(`rows in Table 2 : ${table2Rows}`);
console.log(`figures         : ${figures}`);
console.log(`columns in Fig 2: ${stacked}`);
if (table2Rows === 0) problems.push("Table 2 is empty");
if (figures !== 4) problems.push(`${figures} figures rendered, expected 4`);
if (stacked !== WINDOW) problems.push(`Figur 2 has ${stacked} columns, expected ${WINDOW}`);

// The default typfall has no occupational pension (scheme 1, "Saknar
// tjänstepension") and no private saving, so Tjänstepension is zero for
// every age in the window -- its legend entry must not appear, while Lön
// (never zero before retirement) must.
const fig2Legend = await tab.locator("figure.figure").nth(1).locator(".legend li").allTextContents();
console.log(`Figur 2 legend  : ${fig2Legend.join(" | ")}`);
if (fig2Legend.some((text) => text.includes("Tjänstepension") || text.includes("Occupational"))) {
  problems.push("Figur 2's legend still names Tjänstepension, which is zero for this typfall");
}
if (!fig2Legend.some((text) => text.includes("Lön") || text.includes("Earnings"))) {
  problems.push("Figur 2's legend is missing Lön, which is never zero before retirement");
}

// Table 2 thins its own columns the same way: a column with no value
// anywhere in this run gets no header. The default typfall has no
// occupational pension and no private saving (no normal-mode input drives
// it), so both must be absent; Lön (never zero before retirement) must be
// present.
const table2Headers = await tab.locator(".table2 thead th").allTextContents();
console.log(`Table 2 columns : ${table2Headers.length} (${table2Headers.join(", ")})`);
if (table2Headers.some((text) => text.includes("Tjänste-pension") || text.includes("Occupational"))) {
  problems.push("Table 2 still has an occupational-pension column, which is zero for this typfall");
}
if (table2Headers.some((text) => text.includes("Privat pensionssparande") || text.includes("Private pension"))) {
  problems.push("Table 2 still has a private-saving column, which is zero for this typfall");
}
if (!table2Headers.some((text) => text === "Lön" || text === "Earnings")) {
  problems.push("Table 2 is missing Lön, which is never zero before retirement");
}
if (!table2Headers.some((text) => text === "Kommunal skatt" || text === "Municipal tax")) {
  problems.push("Table 2 is missing the municipal-tax column");
}
if (!table2Headers.some((text) => text === "Statlig skatt" || text === "State tax")) {
  problems.push("Table 2 is missing the state-tax column");
}

// Hovering/focusing a Table 2 header names what its own figure is built
// from -- an `<abbr title>`, so the check reads the attribute a screen
// reader or a real hover would surface, not just that the cell has text.
const stateTaxInfo = await tab
  .locator(".table2 thead abbr", { hasText: /Statlig skatt|State tax/ })
  .getAttribute("title");
console.log(`state tax info  : ${stateTaxInfo}`);
if (!stateTaxInfo || !/20\s*%/.test(stateTaxInfo) || !/public.service/i.test(stateTaxInfo)) {
  problems.push(
    `the "Statlig skatt"/"State tax" header's title is "${stateTaxInfo}", expected it to name the 20% ` +
      "threshold and the public-service fee",
  );
}
const yearHeaderIsAbbr = await tab.locator(".table2 thead th").first().locator("abbr").count();
if (yearHeaderIsAbbr !== 0) {
  problems.push("the Year column got an explanatory <abbr> it doesn't need");
}

// Table 2's own CSV and Excel downloads -- scoped the same way, to the
// `.panel` holding `table.table2` (beside the Årsvis/Månadsvis toggle).
const table2Actions = tab
  .locator("section.panel", { has: tab.locator("table.table2") })
  .locator(".panel-actions");
const table2Csv = await download(table2Actions.getByRole("button", { name: "Ladda ner CSV" }));
console.log(`table2 CSV      : ${table2Csv.filename}, ${table2Csv.buffer.length} bytes`);
if (!table2Csv.filename.endsWith(".csv") || !table2Csv.buffer.toString("utf8").includes("Statlig skatt")) {
  problems.push(`Table 2's CSV download ("${table2Csv.filename}") doesn't look like Table 2's own data`);
}
const table2Xlsx = await download(table2Actions.getByRole("button", { name: "Ladda ner Excel" }));
console.log(`table2 Excel    : ${table2Xlsx.filename}, ${table2Xlsx.buffer.length} bytes`);
if (!table2Xlsx.filename.endsWith(".xlsx") || !looksLikeXlsx(table2Xlsx.buffer)) {
  problems.push(`Table 2's Excel download ("${table2Xlsx.filename}") isn't a real .xlsx file`);
}

/** One cell of Table 2, found by row and column position -- Table 2's cells
 * carry no `data-key`/`data-col` the way Table 1's do, so position is all
 * there is; the header row just found gives the column indices. */
async function table2Cell(rowIndex, colIndex) {
  const text = await tab
    .locator(".table2 tbody tr")
    .nth(rowIndex)
    .locator("td")
    .nth(colIndex)
    .textContent();
  if (text === null) return Number.NaN;
  return Number(text.replace(/[^\d,-]/g, "").replace(",", "."));
}

// The split must reconcile with the gross/net columns Table 2 already shows --
// checked against the page's own numbers, not a second fixture value, since
// the fixture's sparse rows don't carry this split (see result.test.ts for the
// engine-side version of the same check, against every row).
const grossCol = table2Headers.findIndex((text) => text === "Inkomst brutto" || text === "Gross income");
const municipalCol = table2Headers.findIndex((text) => text === "Kommunal skatt" || text === "Municipal tax");
const stateCol = table2Headers.findIndex((text) => text === "Statlig skatt" || text === "State tax");
const netCol = table2Headers.findIndex((text) => text === "Inkomst efter skatt" || text === "Income after tax");
if (grossCol >= 0 && municipalCol >= 0 && stateCol >= 0 && netCol >= 0 && table2Rows > 0) {
  const sampleRows = [...new Set([0, Math.floor(table2Rows / 2), table2Rows - 1])];
  for (const rowIndex of sampleRows) {
    const gross = await table2Cell(rowIndex, grossCol);
    const municipal = await table2Cell(rowIndex, municipalCol);
    const state = await table2Cell(rowIndex, stateCol);
    const net = await table2Cell(rowIndex, netCol);
    if (Math.abs(municipal + state - (gross - net)) > 1) {
      problems.push(
        `Table 2 row ${rowIndex}: municipal (${municipal}) + state (${state}) tax ` +
          `does not reconcile with gross (${gross}) - net (${net})`,
      );
    }
  }
}

// The tax-per-year chart: two series (municipal, state), always both visible
// for this typfall, and the same age window Figur 2 covers.
const taxChartLegend = await tab.locator("figure.figure").nth(3).locator(".legend li").allTextContents();
console.log(`Tax chart legend: ${taxChartLegend.join(" | ")}`);
if (taxChartLegend.length !== 2) {
  problems.push(`Tax-per-year chart shows ${taxChartLegend.length} legend entries, expected 2`);
}
const taxChartColumns = await tab
  .locator("figure.figure")
  .nth(3)
  .locator("svg rect.bar")
  .evaluateAll((nodes) => new Set(nodes.map((node) => node.getAttribute("x"))).size);
console.log(`columns in tax chart: ${taxChartColumns}`);
if (taxChartColumns !== WINDOW) {
  problems.push(`Tax-per-year chart has ${taxChartColumns} columns, expected ${WINDOW}`);
}

// ---- Avancerat lage -----------------------------------------------------
//
// Normal mode is everything above; this drives the workbook's other mode. The
// point of each check is that the setting reaches the *model*, not merely the
// DOM: a control that renders but is never read would pass a "the field exists"
// assertion and fail every user.

const modeButtons = await tab.locator(".mode-toggle .panel-btn").count();
if (modeButtons !== 2) {
  problems.push(`mode toggle has ${modeButtons} buttons, expected 2 (Normalt / Avancerat)`);
}

// Advanced settings must not be reachable until the mode is switched.
if (await tab.locator(".advanced-box").isVisible()) {
  problems.push("the advanced panel is visible in normal mode");
}

await tab.locator('.mode-toggle .panel-btn[data-mode="advanced"]').click();
await tab.waitForTimeout(50);

const advGroups = await tab.locator(".advanced-box .adv-group").count();
console.log(`advanced groups : ${advGroups}`);
// Seven settings groups, plus the salary-path grid and the PGB grid.
if (advGroups !== 9) {
  problems.push(`advanced mode shows ${advGroups} groups, expected 9`);
}

/** Opens the group holding a setting and returns its control. */
async function setting(key) {
  const control = tab.locator(`[data-setting="${key}"]`);
  const group = control.locator("xpath=ancestor::details[1]");
  if (!(await group.evaluate((node) => node.open))) {
    await group.locator("summary").click();
  }
  return control;
}

// A concrete municipal rate, against the historical average the default uses.
// 35% is well above the ~32.4% average, so the tax has to rise.
const municipalBefore = await table2Cell(table2Rows - 1, municipalCol);
const kommunalskatt = await setting("kommunalskatt");
await kommunalskatt.fill("35");
await kommunalskatt.dispatchEvent("change");
await tab.waitForTimeout(50);
const municipalAfter = await table2Cell(table2Rows - 1, municipalCol);
console.log(`municipal tax   : ${municipalBefore} -> ${municipalAfter} at 35%`);
if (!(municipalAfter > municipalBefore)) {
  problems.push(
    `a 35% municipal rate did not raise the municipal tax (${municipalBefore} -> ${municipalAfter})`,
  );
}

// Aterstall: the reset button has to undo it.
await tab.locator('[data-action="reset-advanced"]').click();
await tab.waitForTimeout(50);
const municipalReset = await table2Cell(table2Rows - 1, municipalCol);
if (municipalReset !== municipalBefore) {
  problems.push(
    `the reset button left the municipal tax at ${municipalReset}, expected ${municipalBefore}`,
  );
}

// A named municipality: a concrete rate, computed independently here rather
// than trusted from the app's own module, so a wrong copy of the SCB table
// would show up as a mismatch rather than the check agreeing with itself.
const municipalTaxData = JSON.parse(readFileSync(join(repo, "packages/data/municipal-tax.json"), "utf8"));
function latestSeriesRate(name) {
  const series = municipalTaxData.series[name];
  let lastNonNull = series.values.length - 1;
  while (lastNonNull >= 0 && series.values[lastNonNull] === null) lastNonNull -= 1;
  const claimedIndex = (series.lastActualYear ?? municipalTaxData.firstYear + lastNonNull) - municipalTaxData.firstYear;
  const index = Math.min(claimedIndex, lastNonNull);
  return series.values[index] / 100;
}
const expectedBurialOnly = latestSeriesRate("begravningsavgift");
const expectedChurchMember = latestSeriesRate("kyrkoavgift");

const kommunSelect = await setting("kommunalskatt-municipality");
await kommunSelect.selectOption({ label: "Danderyd" });
await tab.waitForTimeout(50);
const kommunalskattField = tab.locator('[data-setting="kommunalskatt"]');
const kommunalskattValue = await kommunalskattField.inputValue();
console.log(`kommun select   : Danderyd -> ${kommunalskattValue}%`);
// The exact published rate, read from the same source the app itself reads
// rather than retyped here -- a field showing a rounder number than the one
// it was just told to show is exactly the mismatch a reviewer caught between
// this field and its own hint text underneath it (1.3 next to "~1,32 %").
const kommunalskattSource = Number(
  readFileSync(join(repo, "apps/web/src/kommunalskatt.ts"), "utf8").match(/"Danderyd":\s*([\d.]+)/)?.[1],
);
if (kommunalskattValue !== String(kommunalskattSource)) {
  problems.push(
    `picking Danderyd set kommunalskatt to "${kommunalskattValue}", expected the exact published ` +
      `rate "${kommunalskattSource}"`,
  );
}
const municipalAfterPick = await table2Cell(table2Rows - 1, municipalCol);
if (!(municipalAfterPick < municipalBefore)) {
  problems.push(
    `picking a lower-tax municipality did not lower the municipal tax (${municipalBefore} -> ${municipalAfterPick})`,
  );
}

// The footgun this closes: picking a municipality flips the model off the
// historical-average church/burial rate and onto this panel's own field,
// which defaults to 0 -- silently zeroing it unless something fills it. It
// should have auto-filled with the burial-only rate instead.
const begravningsavgiftField = tab.locator('[data-setting="begravningsavgift"]');
const churchSelect = tab.locator('[data-setting="begravningsavgift-select"]');
const autoFilled = await begravningsavgiftField.inputValue();
const autoChoice = await churchSelect.inputValue();
// Exact strings, not a tolerance -- this field carries three decimal places
// specifically so a picked rate shows exactly, matching the reference figure
// its own hint text quotes underneath it. A reviewer caught this field
// showing "1.3" next to a hint reading "~1,32 %"; the fix was giving the
// field the same precision, and checking the exact string is what would
// catch a regression back to the coarser one.
const expectedBurialOnlyStr = String(Math.round(expectedBurialOnly * 100_000) / 1000);
console.log(`church/burial   : auto-filled to ${autoFilled}% (${autoChoice}), expected ${expectedBurialOnlyStr}%`);
if (autoFilled === "0") {
  problems.push("picking a municipality left the church/burial field at 0 instead of auto-filling it");
}
if (autoFilled !== expectedBurialOnlyStr) {
  problems.push(
    `the auto-filled church/burial rate is "${autoFilled}", expected the exact burial-only average ` +
      `"${expectedBurialOnlyStr}"`,
  );
}
if (autoChoice !== "rest") {
  problems.push(`the church/burial select reads "${autoChoice}" after auto-fill, expected "rest"`);
}

// Explicitly picking "member" overrides the auto-fill with the combined
// church+burial rate -- a materially larger number, computed independently
// from the same K_skatt data the app's own module reads.
await churchSelect.selectOption({ value: "member" });
await tab.waitForTimeout(50);
const memberFilled = await begravningsavgiftField.inputValue();
const expectedChurchMemberStr = String(Math.round(expectedChurchMember * 100_000) / 1000);
console.log(`church member   : ${memberFilled}%, expected ${expectedChurchMemberStr}%`);
if (memberFilled !== expectedChurchMemberStr) {
  problems.push(`picking "member" set the rate to "${memberFilled}", expected "${expectedChurchMemberStr}"`);
}

// Aterstall undoes both selects, not just the numbers they filled.
await tab.locator('[data-action="reset-advanced"]').click();
await tab.waitForTimeout(50);
const kommunSelectAfterReset = await kommunSelect.inputValue();
const churchSelectAfterReset = await churchSelect.inputValue();
if (kommunSelectAfterReset !== "") {
  problems.push(`the reset button left the municipality select at "${kommunSelectAfterReset}", expected blank`);
}
if (churchSelectAfterReset !== "") {
  problems.push(`the reset button left the church/burial select at "${churchSelectAfterReset}", expected "own rate"`);
}

// Housing supplement: the shipped typfall draws none, and a rent high enough
// has to bring some in. `Bidrag` is Table 2's benefits column.
const benefitsCol = table2Headers.findIndex(
  (text) => text.startsWith("Bidrag") || text.startsWith("Benefits"),
);
if (benefitsCol === -1) {
  problems.push("Table 2 has no benefits column to check the housing supplement against");
} else {
  const benefitsBefore = await table2Cell(table2Rows - 1, benefitsCol);
  const hyra = await setting("hyra");
  await hyra.fill("12000");
  await hyra.dispatchEvent("change");
  await tab.waitForTimeout(50);
  const benefitsAfter = await table2Cell(table2Rows - 1, benefitsCol);
  console.log(`benefits at rent: ${benefitsBefore} -> ${benefitsAfter} at 12 000 kr/month`);
  if (!(benefitsAfter > benefitsBefore)) {
    problems.push(
      `doubling the rent did not raise the housing supplement (${benefitsBefore} -> ${benefitsAfter})`,
    );
  }
  await tab.locator('[data-action="reset-advanced"]').click();
  await tab.waitForTimeout(50);
}

// The salary grid fills from the computed path rather than empty, and zeroing
// a year has to cost pension. An empty grid would read as a lifetime of no
// income, which is the failure this guards.
//
// Scoped to salaryPath.ts's own group: the PGB grid below shares the
// `.adv-grid` class for its CSS, and is on screen at the same time, so an
// unscoped `.adv-grid tbody tr` would count both tables' rows together.
const salaryGrid = tab.locator('details[data-group="salary-path"] .adv-grid tbody tr');
const ownIncome = await setting("ownIncome");
await ownIncome.check();
await tab.waitForTimeout(80);
const gridRows = await salaryGrid.count();
console.log(`salary grid rows: ${gridRows}`);
// Ages 15 up to the last worked one; retirement is at 66 for this typfall.
if (gridRows < 40 || gridRows > 60) {
  problems.push(`the salary grid has ${gridRows} rows, expected about fifty`);
}
const firstIncome = await salaryGrid.last().locator("input").first().inputValue();
if (!(Number(firstIncome) > 0)) {
  problems.push(`the salary grid filled the last worked year with ${firstIncome}, expected an income`);
}

const grossBefore = await table2Cell(table2Rows - 1, grossCol);
// Zero the last ten worked years -- a decade of leave.
for (let i = 0; i < 10; i += 1) {
  const cells = salaryGrid.nth(gridRows - 1 - i).locator("input");
  await cells.first().fill("0");
  await cells.first().dispatchEvent("change");
  await cells.last().fill("0");
  await cells.last().dispatchEvent("change");
}
await tab.waitForTimeout(80);
const grossAfter = await table2Cell(table2Rows - 1, grossCol);
console.log(`gross pension   : ${grossBefore} -> ${grossAfter} after zeroing ten years`);
if (!(grossAfter < grossBefore)) {
  problems.push(
    `zeroing ten years of salary did not lower the pension (${grossBefore} -> ${grossAfter})`,
  );
}


// PGB: sickness/activity compensation is still hand-typed kronor; conscription
// (a single date range) and study (a per-age semester count) compute their
// own kronor instead, the same way the real PGB sheet does. The shipped
// workbook has none of the three (`pgbManual`'s own comment: "childcare years
// are the only PGB a default run earns"), so an untouched grid must not
// silently change anything, and a filled-in one must.
const pgbGroup = tab.locator('details[data-group="pgb"]');
if (!(await pgbGroup.evaluate((node) => node.open))) {
  await pgbGroup.locator("summary").click();
}
const pgbGrid = pgbGroup.locator(".pgb-grid tbody tr");
const pgbRows = await pgbGrid.count();
console.log(`pgb grid rows   : ${pgbRows}`);
// Ages 16 through 70 -- see pgb.ts's own comment on that range.
if (pgbRows !== 55) {
  problems.push(`the PGB grid has ${pgbRows} rows, expected 55 (ages 16-70)`);
}

// Studier's own semester count and its kronor, and Värnplikt's own days and
// kronor, are grouped under one named header each -- reported as unclear
// that "Antal terminer" and "PGB studier, kr" were even related fields.
const pgbColCount = await pgbGrid.first().locator("td").count();
if (pgbColCount !== 7) {
  problems.push(
    `a PGB grid row has ${pgbColCount} cells, expected 7 (Year, Age, SA, semesters, study kr, days, conscription kr)`,
  );
}
const pgbGroupHeads = await pgbGroup.locator(".pgb-grid thead tr").first().locator("th").allTextContents();
console.log(`pgb group heads : ${pgbGroupHeads.join(" | ")}`);
if (!pgbGroupHeads.some((h) => /studier/i.test(h))) {
  problems.push(`the PGB grid's header row has no "Studier" group over Antal terminer/PGB studier, got: ${pgbGroupHeads.join(" | ")}`);
}
if (!pgbGroupHeads.some((h) => /värnplikt/i.test(h))) {
  problems.push(`the PGB grid's header row has no "Värnplikt" group over Dagar/PGB värnplikt, got: ${pgbGroupHeads.join(" | ")}`);
}

// Värnplikt: wsPGB!H4/H5, a single date range rather than a row-per-age
// entry -- the days and the kronor it earns show per touched year in the
// grid itself (columns 6 and 7, "Dagar" / "PGB värnplikt, kr"), the same
// way the sheet shows them, and the standalone readout above the grid is
// left for the one thing the grid cannot show: a period too short to earn
// anything, with no touched-year row to hold a zero.
const pensionBeforeVpl = await kpiValue(0);
const vplRow1998 = pgbGrid.nth(1998 - 1959 - 16); // this typfall's born 1959
await pgbGroup.locator('[data-setting="pgbConscriptionStart"]').fill("1998-01-01");
await pgbGroup.locator('[data-setting="pgbConscriptionStart"]').dispatchEvent("change");
await pgbGroup.locator('[data-setting="pgbConscriptionEnd"]').fill("1998-02-01"); // 31 days, under the 120 minimum
await pgbGroup.locator('[data-setting="pgbConscriptionEnd"]').dispatchEvent("change");
await tab.waitForTimeout(80);
const vplShortReadout = await pgbGroup.locator(".pgb-vpl-readout").textContent();
console.log(`vpl short period: "${vplShortReadout}"`);
if (!vplShortReadout || !/120/.test(vplShortReadout)) {
  problems.push(`a conscription period under 120 days reads "${vplShortReadout}", expected the "under 120 days" warning`);
}
const vplDaysTooShort = await vplRow1998.locator("td").nth(5).textContent();
if (vplDaysTooShort !== "") {
  problems.push(`a conscription period under 120 days still shows "${vplDaysTooShort}" in the grid's days column, expected blank`);
}
const pensionAfterShortVpl = await kpiValue(0);
if (pensionAfterShortVpl !== pensionBeforeVpl) {
  problems.push(
    `a conscription period under 120 days moved the pension (${pensionBeforeVpl} -> ${pensionAfterShortVpl}), expected no change`,
  );
}

await pgbGroup.locator('[data-setting="pgbConscriptionEnd"]').fill("1998-12-31"); // now a full year, well over 120 days
await pgbGroup.locator('[data-setting="pgbConscriptionEnd"]').dispatchEvent("change");
await tab.waitForTimeout(80);
const vplReadout = await pgbGroup.locator(".pgb-vpl-readout").textContent();
if (vplReadout !== "") {
  problems.push(
    `the conscription readout reads "${vplReadout}" for a valid period, expected empty -- the grid's own ` +
      "Dagar/PGB värnplikt columns show the detail instead of a line of text above it",
  );
}
const vplDays = await vplRow1998.locator("td").nth(5).textContent();
const vplKr = await vplRow1998.locator("td").nth(6).textContent();
console.log(`vpl days/kr row : ${vplDays} dagar, ${vplKr} kr (1998)`);
if (!vplDays || Number(vplDays.replace(/[^\d]/g, "")) !== 364) {
  problems.push(`the 1998 row's own days cell reads "${vplDays}", expected 364`);
}
if (!vplKr || !(Number(vplKr.replace(/[^\d]/g, "")) > 0)) {
  problems.push(`the 1998 row's own PGB värnplikt cell reads "${vplKr}", expected a positive kronor figure`);
}
const pensionAfterVpl = await kpiValue(0);
console.log(`pension w/ vpl  : ${pensionBeforeVpl} -> ${pensionAfterVpl} kr after a conscription period`);
if (!(pensionAfterVpl > pensionBeforeVpl)) {
  problems.push(
    `entering a conscription date range did not raise the pension (${pensionBeforeVpl} -> ${pensionAfterVpl})`,
  );
}

// "Rensa": a native date input's own clear affordance is easy to miss
// packed into a header cell this small, and there was no way to clear just
// the period without "Använd normala inställningar" clearing all of
// advanced mode -- asked for by name after this shipped.
const vplClear = pgbGroup.locator(".pgb-vpl-clear");
if (!(await vplClear.isVisible())) {
  problems.push('the "Rensa" button is not visible with a conscription period entered');
}
await vplClear.click();
await tab.waitForTimeout(80);
const vplStartAfterClear = await pgbGroup.locator('[data-setting="pgbConscriptionStart"]').inputValue();
const vplEndAfterClear = await pgbGroup.locator('[data-setting="pgbConscriptionEnd"]').inputValue();
const vplDaysAfterClear = await vplRow1998.locator("td").nth(5).textContent();
const pensionAfterClear = await kpiValue(0);
console.log(
  `pgb vpl cleared : start "${vplStartAfterClear}", end "${vplEndAfterClear}", ` +
    `pension ${pensionAfterVpl} -> ${pensionAfterClear}`,
);
if (vplStartAfterClear !== "" || vplEndAfterClear !== "") {
  problems.push(
    `"Rensa" left the dates at "${vplStartAfterClear}"/"${vplEndAfterClear}", expected both empty`,
  );
}
if (vplDaysAfterClear !== "") {
  problems.push(`"Rensa" left "${vplDaysAfterClear}" in the 1998 row's days cell, expected blank`);
}
if (pensionAfterClear !== pensionBeforeVpl) {
  problems.push(
    `"Rensa" left the pension at ${pensionAfterClear}, expected it back at ${pensionBeforeVpl} (before the period)`,
  );
}
if (await vplClear.isVisible()) {
  problems.push('the "Rensa" button is still visible after clearing, expected hidden with nothing to clear');
}

// Re-enter the same period so the dialog/study checks below still have a
// conscription entry to carry through the rest of this run.
await pgbGroup.locator('[data-setting="pgbConscriptionStart"]').fill("1998-01-01");
await pgbGroup.locator('[data-setting="pgbConscriptionStart"]').dispatchEvent("change");
await pgbGroup.locator('[data-setting="pgbConscriptionEnd"]').fill("1998-12-31");
await pgbGroup.locator('[data-setting="pgbConscriptionEnd"]').dispatchEvent("change");
await tab.waitForTimeout(80);

// Antal terminer: a per-age semester count, auto-computed into its own
// kronor right there in the grid -- wsPGB!P shows the same figure beside
// its own "Antal terminer" cell. 2005 (age 46 for this typfall's 1959 birth
// year) is a year study-PGB actually existed (1995 on), distinct from the
// conscription test's own 1998 so the two effects stay untangled.
const pensionBeforeStudy = await kpiValue(0);
const studyRow = pgbGrid.nth(2005 - 1959 - 16);
const semesterInput = studyRow.locator("td").nth(3).locator("input");
await semesterInput.fill("1");
await semesterInput.dispatchEvent("change");
await tab.waitForTimeout(80);
const studyKr = await studyRow.locator("td").nth(4).textContent();
console.log(`study kr readout: ${studyKr}`);
if (!studyKr || !(Number(studyKr.replace(/[^\d]/g, "")) > 0)) {
  problems.push(`entering 1 semester's own readout reads "${studyKr}", expected a positive kronor figure`);
}
const pensionAfterStudy = await kpiValue(0);
console.log(`pension w/ study: ${pensionBeforeStudy} -> ${pensionAfterStudy} kr after 1 semester`);
if (!(pensionAfterStudy > pensionBeforeStudy)) {
  problems.push(`entering a semester count did not raise the pension (${pensionBeforeStudy} -> ${pensionAfterStudy})`);
}

const pensionBeforePgb = await kpiValue(0);
const pgbSaInput = pgbGrid.first().locator("td").nth(2).locator("input");
await pgbSaInput.fill("200000");
await pgbSaInput.dispatchEvent("change");
await tab.waitForTimeout(80);
const pensionAfterPgb = await kpiValue(0);
console.log(`pension w/ pgb  : ${pensionBeforePgb} -> ${pensionAfterPgb} kr after a sickness-comp entry`);
if (!(pensionAfterPgb > pensionBeforePgb)) {
  problems.push(
    `entering a PGB sickness-compensation amount did not raise the pension (${pensionBeforePgb} -> ${pensionAfterPgb})`,
  );
}

// "Visa alla kolumner" moves the same table into a <dialog> -- reported as
// having to scroll sideways to see the rest of it in the sidebar's own
// ~280px-wide scroller. The point of each check below is that it is the same
// table (the 200 000 typed above is still there, and a further edit still
// reaches the model), not a copy, and that the dialog itself never needs that
// horizontal scrollbar.
await pgbGroup.locator('[data-action="pgb-expand"]').click();
await tab.waitForTimeout(100);
const pgbDialog = tab.locator(".pgb-dialog");
const pgbDialogOpen = await pgbDialog.evaluate((node) => node.open);
console.log(`pgb dialog open : ${pgbDialogOpen}`);
if (!pgbDialogOpen) problems.push("clicking \"Visa alla kolumner\" did not open the PGB dialog");

const pgbDialogScroll = pgbDialog.locator(".adv-grid-scroll");
const pgbDialogMetrics = await pgbDialogScroll.evaluate((el) => ({
  scrollWidth: el.scrollWidth,
  clientWidth: el.clientWidth,
}));
console.log(`pgb dialog fit  : scrollWidth ${pgbDialogMetrics.scrollWidth} <= clientWidth ${pgbDialogMetrics.clientWidth}?`);
if (pgbDialogMetrics.scrollWidth > pgbDialogMetrics.clientWidth) {
  problems.push(
    `the PGB dialog still needs horizontal scroll (scrollWidth ${pgbDialogMetrics.scrollWidth} > ` +
      `clientWidth ${pgbDialogMetrics.clientWidth}) -- the whole point of "Visa alla kolumner" is to avoid that`,
  );
}

const pgbDialogFirstSa = await pgbDialog.locator(".pgb-grid tbody tr").first().locator("td").nth(2).locator("input").inputValue();
if (pgbDialogFirstSa !== "200000") {
  problems.push(
    `the PGB dialog shows "${pgbDialogFirstSa}" for the value typed before it opened, expected "200000" -- ` +
      "it should be the same table, not a copy",
  );
}
// A second edit, made from inside the dialog this time, still has to reach
// the model -- the dialog is not a read-only preview.
const pgbDialogSecondSa = pgbDialog.locator(".pgb-grid tbody tr").nth(1).locator("td").nth(2).locator("input");
await pgbDialogSecondSa.fill("100000");
await pgbDialogSecondSa.dispatchEvent("change");
await tab.waitForTimeout(80);
const pensionAfterDialogEdit = await kpiValue(0);
console.log(`pension w/ dialog edit: ${pensionAfterPgb} -> ${pensionAfterDialogEdit}`);
if (!(pensionAfterDialogEdit > pensionAfterPgb)) {
  problems.push(
    `entering a PGB amount from inside the dialog did not raise the pension further ` +
      `(${pensionAfterPgb} -> ${pensionAfterDialogEdit})`,
  );
}

// A native modal <dialog> blocks pointer events on the rest of the page by
// design (confirmed the hard way: a first draft of this check tried to click
// "Använd normala inställningar" while the dialog was still open and Playwright
// timed out with "dialog intercepts pointer events") -- so the only way out is
// its own close button, same as a person has.
await pgbDialog.locator(".dialog-close").click();
await tab.waitForTimeout(80);
const pgbDialogClosed = await pgbDialog.evaluate((node) => !node.open);
const pgbBackInPanel = await pgbGrid.count();
const pgbDialogSecondSaKept = await pgbGrid.nth(1).locator("td").nth(2).locator("input").inputValue();
console.log(`pgb dialog closed: ${pgbDialogClosed}, rows back in panel: ${pgbBackInPanel}, edit kept: ${pgbDialogSecondSaKept}`);
if (!pgbDialogClosed) problems.push("clicking the PGB dialog's own close button did not close it");
if (pgbBackInPanel !== pgbRows) {
  problems.push(
    `after the dialog closed, the PGB grid shows ${pgbBackInPanel} rows back in its panel, expected ${pgbRows}`,
  );
}
if (pgbDialogSecondSaKept !== "100000") {
  problems.push(
    `after the dialog closed, the panel shows "${pgbDialogSecondSaKept}" for the amount typed inside the ` +
      'dialog, expected "100000" -- the grid moved back with the table it is, not a copy of it',
  );
}

// "Privat pensionssparande": the same number means kronor/month or a share of
// income depending on its own size -- exposed as an explicit toggle now,
// rather than a single box whose meaning depended on how big the typed value
// happened to be. Starts in "Belopp" (amount) mode with the share input
// hidden; switching modes has to show the other input and hide this one, and
// a percentage typed after switching has to reach the model the same way an
// amount would. The default "Sparandet börjar år" (2026) is after this
// typfall's own retirement, so nothing would ever accrue regardless of mode
// unless it moves earlier first -- that part is unrelated to the toggle
// itself, just what the eligibility window needs to actually open. Run here,
// with the PGB checks above already done and the full reset below still to
// come, so raising private saving cannot perturb any earlier assumption.
const ipsStart = await setting("ipsStart");
await ipsStart.fill("1990");
await ipsStart.dispatchEvent("change");
const ipsAmountInput = await setting("ipsMonthly-amount");
const ipsShareInput = tab.locator('[data-setting="ipsMonthly-share"]');
if (!(await ipsAmountInput.isVisible()) || (await ipsShareInput.isVisible())) {
  problems.push("the private-saving toggle does not open in amount mode with the share input hidden");
}
const ipsBefore = await shown("ips", "monthly");
await tab.locator('.adv-ips [data-mode="share"]').click();
await tab.waitForTimeout(50);
if ((await ipsAmountInput.isVisible()) || !(await ipsShareInput.isVisible())) {
  problems.push('switching to "Andel av inkomst" did not hide the amount input and show the share input');
}
await ipsShareInput.fill("5");
await ipsShareInput.dispatchEvent("change");
await tab.waitForTimeout(50);
const ipsAfter = await shown("ips", "monthly");
console.log(`private saving  : ${ipsBefore} -> ${ipsAfter} kr/month at 5% of income`);
if (!(ipsAfter > ipsBefore)) {
  problems.push(`setting private saving to 5% of income did not raise it (${ipsBefore} -> ${ipsAfter})`);
}

// "Nollställ alla värden": every row's own income and wage cell has to read
// 0, not just the ten already zeroed above. Not a pension-direction check --
// a whole working life at 0 kr leans on garantipension, whose own means-
// tested taper can (correctly, faithfully) leave *more* total gross pension
// than a life with some income in it does, so "lower" is not a safe
// assumption here. Run last, right before the full reset below: every PGB
// check above assumes a normal-income economic regime where more PGB is
// strictly more pension, which a fully zeroed salary grid no longer is.
await tab.locator('[data-group="salary-path"] [data-action="salary-zero"]').click();
await tab.waitForTimeout(80);
const nonZeroCells = await salaryGrid.locator("input").evaluateAll(
  (inputs) => inputs.filter((el) => el.value !== "0").length,
);
console.log(`salary zero all : ${nonZeroCells} cell(s) left non-zero, expected 0`);
if (nonZeroCells !== 0) {
  problems.push(`"Nollställ alla värden" left ${nonZeroCells} salary-grid cell(s) not at 0`);
}

// Aterstall clears the grid, and the conscription dates, back to empty --
// the same as every other advanced-mode field -- checked with the dialog
// closed, since it is not reachable any other way (see above).
await tab.locator('[data-action="reset-advanced"]').click();
await tab.waitForTimeout(50);
const pgbAfterReset = await pgbSaInput.inputValue();
if (pgbAfterReset !== "0") {
  problems.push(`the reset button left the PGB field at "${pgbAfterReset}", expected "0"`);
}
const vplStartAfterReset = await pgbGroup.locator('[data-setting="pgbConscriptionStart"]').inputValue();
if (vplStartAfterReset !== "") {
  problems.push(`the reset button left the conscription start date at "${vplStartAfterReset}", expected empty`);
}

// Partiellt uttag: "Andel uttag, inkomstpension/premiepension" and
// "Definitivt uttag vid ålder" simulate combining part-time work with a
// partial pension for a few years -- manual 3.7's own example, a "jobbonär
// under en viss period". `withdrawalShare` ties Lön to whichever share is
// drawn by default, so the point of each check is that setting the share
// alone moves both Lön (which the setting never mentions) and the pension
// columns together, and that "Definitivt vid" actually bounds the period
// rather than leaving it open-ended.
const pwGroup = tab.locator('details[data-group="partialWithdrawal"]');
if (!(await pwGroup.evaluate((node) => node.open))) {
  await pwGroup.locator("summary").click();
}
const pwHeaders = await tab.locator(".table2 thead th").allTextContents();
const pwAlderCol = pwHeaders.findIndex((t) => t === "Ålder");
const pwLonCol = pwHeaders.findIndex((t) => t === "Lön");
const pwIncomeCol = pwHeaders.findIndex((t) => t === "Inkomst- och tilläggspension");
const pwPremiumCol = pwHeaders.findIndex((t) => t === "Premiepension");

async function rowForAge(age) {
  const rows = await tab.locator(".table2 tbody tr").count();
  for (let i = 0; i < rows; i += 1) {
    if ((await table2Cell(i, pwAlderCol)) === age) return i;
  }
  return -1;
}

// The default typfall draws its pension in full right at the retirement age
// (66), so Lön at 67 already reads 0 -- the baseline a partial withdrawal has
// to move away from.
const row67 = await rowForAge(67);
if (row67 === -1) problems.push("could not find age 67 in Table 2 to test partial withdrawal against");
const lonBefore = row67 === -1 ? Number.NaN : await table2Cell(row67, pwLonCol);
const incomeBefore = row67 === -1 ? Number.NaN : await table2Cell(row67, pwIncomeCol);
const premiumBefore = row67 === -1 ? Number.NaN : await table2Cell(row67, pwPremiumCol);
console.log(`age 67, full    : lön ${lonBefore}, inkomst-/tilläggspension ${incomeBefore}, premiepension ${premiumBefore}`);
if (lonBefore !== 0) {
  problems.push(`the default typfall already shows a nonzero salary at 67 (${lonBefore}) before touching partial withdrawal`);
}

await tab.locator('[data-setting="defAr"]').fill("70");
await tab.locator('[data-setting="defAr"]').dispatchEvent("change");
await tab.locator('[data-setting="uttagIp"]').selectOption("0.5");
await tab.locator('[data-setting="uttagPp"]').selectOption("0.5");
await tab.waitForTimeout(80);

const lonDuring = row67 === -1 ? Number.NaN : await table2Cell(row67, pwLonCol);
const incomeDuring = row67 === -1 ? Number.NaN : await table2Cell(row67, pwIncomeCol);
const premiumDuring = row67 === -1 ? Number.NaN : await table2Cell(row67, pwPremiumCol);
console.log(`age 67, 50%     : lön ${lonDuring}, inkomst-/tilläggspension ${incomeDuring}, premiepension ${premiumDuring}`);
if (!(lonDuring > 0)) {
  problems.push(
    `a 50% partial withdrawal still shows no salary at 67 (${lonDuring}) -- work should follow the ` +
      "withdrawal and continue part-time",
  );
}
if (!(incomeDuring > 0 && incomeDuring < incomeBefore)) {
  problems.push(
    `a 50% partial withdrawal did not reduce inkomst-/tilläggspension at 67 (${incomeBefore} -> ${incomeDuring})`,
  );
}
if (!(premiumDuring > 0 && premiumDuring < premiumBefore)) {
  problems.push(`a 50% partial withdrawal did not reduce premiepension at 67 (${premiumBefore} -> ${premiumDuring})`);
}

// "Definitivt vid" bounds the partial period: at and after that age, Lön and
// the pension are back to what a full retirement always looked like.
const row70 = await rowForAge(70);
const lonFinal = row70 === -1 ? Number.NaN : await table2Cell(row70, pwLonCol);
const incomeFinal = row70 === -1 ? Number.NaN : await table2Cell(row70, pwIncomeCol);
console.log(`age 70, final   : lön ${lonFinal}, inkomst-/tilläggspension ${incomeFinal}`);
if (lonFinal !== 0) {
  problems.push(`salary is still nonzero at 70 (${lonFinal}), the age "Definitivt vid" was set to`);
}
if (!(incomeFinal >= incomeDuring)) {
  problems.push(
    `inkomst-/tilläggspension at 70 (${incomeFinal}) is not back up to a full withdrawal's level ` +
      `(${incomeDuring} during the partial period)`,
  );
}

await tab.locator('[data-action="reset-advanced"]').click();
await tab.waitForTimeout(50);
const uttagIpAfterReset = await tab.locator('[data-setting="uttagIp"]').inputValue();
const defArAfterReset = await tab.locator('[data-setting="defAr"]').inputValue();
console.log(`partial withdrawal reset: uttagIp ${uttagIpAfterReset}, defAr ${defArAfterReset}`);
if (uttagIpAfterReset !== "1") {
  problems.push(`reset left "Andel uttag, inkomstpension" at "${uttagIpAfterReset}", expected "1" (100%)`);
}
if (defArAfterReset !== "0") {
  problems.push(`reset left "Definitivt uttag vid ålder" at "${defArAfterReset}", expected "0"`);
}

// Back to normal mode for the screenshots, and to leave the page as found.
await tab.locator('[data-action="reset-advanced"]').click();
await tab.locator('.mode-toggle .panel-btn[data-mode="normal"]').click();
await tab.waitForTimeout(50);

// ---- Jamfor scenarier -----------------------------------------------------
//
// A second top-level view, not a workbook mode: the baseline plus up to three
// variants, each overriding only salary, retirement age, start-of-work age
// and occupational scheme, compared in one shared table and chart rather than
// a KPI row and a full Table 1 repeated inside every scenario's own card. The
// point of each check, as with Avancerat above, is that a variant's own
// controls reach its own run and nothing else's.

const screenButtons = await tab.locator(".screen-toggle .panel-btn").count();
if (screenButtons !== 2) {
  problems.push(`screen toggle has ${screenButtons} buttons, expected 2 (Prognos / Jämför scenarier)`);
}

await tab.locator('.screen-toggle .panel-btn[data-screen="compare"]').click();
await tab.waitForTimeout(50);

const compareCards = await tab.locator(".compare-card").count();
console.log(`compare cards   : ${compareCards}`);
if (compareCards !== 2) {
  problems.push(`the compare tab shows ${compareCards} cards, expected 2 (baseline + one scenario)`);
}

const compareTable = tab.locator("table.compare-table");
// Three KPI rows, a spacer, sixteen Table 1 rows, TABLE1_LINES's own spacer,
// and the two footnotes -- see renderCompareTable's own comment for the count.
const compareRows = await compareTable.locator("tbody tr").count();
console.log(`compare rows    : ${compareRows}`);
if (compareRows !== 23) {
  problems.push(`the compare table has ${compareRows} rows, expected 23`);
}

async function compareCell(scenario, key, col) {
  return tab
    .locator(`table.compare-table tbody tr[data-key="${key}"] td[data-scenario="${scenario}"][data-col="${col}"]`)
    .textContent();
}

const baselinePension = await compareCell("Utgångsläge", "kpi-pension", "monthly");
const scenario1PensionBefore = await compareCell("Scenario 1", "kpi-pension", "monthly");
console.log(`compare baseline: ${baselinePension}, scenario 1 before: ${scenario1PensionBefore}`);
if (scenario1PensionBefore !== baselinePension) {
  problems.push(
    `a freshly added scenario shows "${scenario1PensionBefore}", expected to start identical to the ` +
      `baseline's "${baselinePension}"`,
  );
}

// Raising the scenario's own salary well above the baseline's must move only
// that scenario's own columns in the table, not the baseline's.
const scenario1Salary = tab.locator(".compare-card").nth(1).locator(".compare-card-controls input").first();
await scenario1Salary.fill("80000");
await scenario1Salary.dispatchEvent("change");
await tab.waitForTimeout(50);
const scenario1PensionAfter = await compareCell("Scenario 1", "kpi-pension", "monthly");
const baselinePensionAfter = await compareCell("Utgångsläge", "kpi-pension", "monthly");
console.log(`scenario salary raised: pension ${scenario1PensionBefore} -> ${scenario1PensionAfter}`);
if (scenario1PensionAfter === scenario1PensionBefore) {
  problems.push(
    `raising a scenario's own salary did not raise its pension in the table (still "${scenario1PensionAfter}")`,
  );
}
if (baselinePensionAfter !== baselinePension) {
  problems.push(
    `raising a scenario's salary changed the baseline's own pension in the table ("${baselinePension}" -> ` +
      `"${baselinePensionAfter}")`,
  );
}

// Start-of-work age is the fourth override, added on request -- a later start
// shortens a scenario's own working life (setup.ts: startage = min(modelStartAge,
// startWorkAge)), so it has to move that scenario's pension on its own, with
// salary and retirement age held fixed, not just ride along with the salary
// control already proven above.
const scenario1StartWork = tab.locator(".compare-card").nth(1).locator(".compare-card-controls input").nth(2);
await scenario1StartWork.fill("35");
await scenario1StartWork.dispatchEvent("change");
await tab.waitForTimeout(50);
const scenario1PensionAfterStartWork = await compareCell("Scenario 1", "kpi-pension", "monthly");
const baselinePensionAfterStartWork = await compareCell("Utgångsläge", "kpi-pension", "monthly");
console.log(
  `scenario start-work raised: pension ${scenario1PensionAfter} -> ${scenario1PensionAfterStartWork}`,
);
if (scenario1PensionAfterStartWork === scenario1PensionAfter) {
  problems.push(
    `raising a scenario's own start-of-work age did not move its pension in the table (still ` +
      `"${scenario1PensionAfterStartWork}")`,
  );
}
if (baselinePensionAfterStartWork !== baselinePension) {
  problems.push(
    `raising a scenario's start-of-work age changed the baseline's own pension in the table ` +
      `("${baselinePension}" -> "${baselinePensionAfterStartWork}")`,
  );
}

// The replacement-rate row: never at or above 100% just because a scenario's
// salary moved. A reviewer's own pasted mockup of this table once showed
// "152,80%" in this exact row -- which turned out, on inspection, to be an
// artifact of assembling that mockup by hand rather than a live figure, but
// this is the check that would catch it if it ever were one.
const shareCells = await compareTable
  .locator('tbody tr[data-key="kpi-replacement"] td[data-col="share"]')
  .allTextContents();
for (const text of shareCells) {
  const value = Number(text.replace(/[^\d,.-]/g, "").replace(",", "."));
  if (Number.isFinite(value) && value >= 100) {
    problems.push(`a replacement-rate cell reads "${text}", which is 100% or more`);
  }
}

// The chart draws one solid line per active scenario.
const dashedLineSelector =
  ".compare-results svg .line-current, .compare-results svg .line-wage-level, .compare-results svg .line-after-tax";
const chartLines = await tab.locator(".compare-results svg .line").count();
const dashedChartLines = await tab.locator(dashedLineSelector).count();
console.log(`chart lines     : ${chartLines}, dashed: ${dashedChartLines}`);
if (chartLines !== 2) {
  problems.push(`the compare chart draws ${chartLines} lines, expected 2`);
}
if (dashedChartLines !== 0) {
  problems.push(`the compare chart has ${dashedChartLines} dashed line(s), expected solid lines only`);
}

// The baseline's own label is editable, like a variant's -- not just a
// fixed "Baseline" heading -- and drives its column's own heading text.
const baselineLabelInput = tab.locator(".compare-card-baseline .compare-card-label");
await baselineLabelInput.fill("Idag");
await baselineLabelInput.dispatchEvent("change");
await tab.waitForTimeout(50);
const renamedHeader = await tab.locator('table.compare-table thead th[data-scenario="Idag"]').count();
console.log(`baseline renamed: ${renamedHeader === 1 ? "yes" : "no"}`);
if (renamedHeader !== 1) {
  problems.push(
    'renaming the baseline card did not retitle its own column (expected a th[data-scenario="Idag"])',
  );
}

// A dashed reference line marks the retirement year -- shared by both
// scenarios here, since a freshly added variant starts at the baseline's own
// retirement age.
const retirementLines = await tab.locator(".compare-results svg .line-retirement").count();
console.log(`retirement line : ${retirementLines}`);
if (retirementLines !== 1) {
  problems.push(`the compare chart draws ${retirementLines} retirement-year reference line(s), expected 1`);
}

// Hovering the chart shows a tooltip with one readout row per active
// scenario -- each its own separate run, unlike every other figure's shared-
// row tooltip.
const compareChartBox = await tab.locator(".compare-results svg").first().boundingBox();
if (compareChartBox === null) {
  problems.push("the compare chart's <svg> has no bounding box to hover");
} else {
  await tab.mouse.move(compareChartBox.x + compareChartBox.width / 2, compareChartBox.y + compareChartBox.height / 2);
  await tab.waitForTimeout(50);
  const tooltipRows = await tab.locator(".compare-results .tooltip .tooltip-row").count();
  console.log(`hover tooltip rows: ${tooltipRows}`);
  if (tooltipRows !== 2) {
    problems.push(`hovering the compare chart shows ${tooltipRows} tooltip row(s), expected 2 (one per scenario)`);
  }
}

// The comparison table's own CSV and Excel downloads -- entirely new; there
// was no export at all here before.
const compareActions = tab.locator(".compare-results .panel-actions");
const compareCsv = await download(compareActions.getByRole("button", { name: "Ladda ner CSV" }));
console.log(`compare CSV     : ${compareCsv.filename}, ${compareCsv.buffer.length} bytes`);
if (
  !compareCsv.filename.endsWith(".csv") ||
  !compareCsv.buffer.toString("utf8").includes("Scenario 1")
) {
  problems.push(`the comparison table's CSV download ("${compareCsv.filename}") doesn't look right`);
}
const compareXlsx = await download(compareActions.getByRole("button", { name: "Ladda ner Excel" }));
console.log(`compare Excel   : ${compareXlsx.filename}, ${compareXlsx.buffer.length} bytes`);
if (!compareXlsx.filename.endsWith(".xlsx") || !looksLikeXlsx(compareXlsx.buffer)) {
  problems.push(`the comparison table's Excel download ("${compareXlsx.filename}") isn't a real .xlsx file`);
}

// Add up to the cap -- baseline plus three variants -- then confirm add
// disables and the table/chart both track the new count.
await tab.locator('[data-action="add-scenario"]').click();
await tab.waitForTimeout(50);
await tab.locator('[data-action="add-scenario"]').click();
await tab.waitForTimeout(50);
const cardsAtCap = await tab.locator(".compare-card").count();
const addDisabled = await tab.locator('[data-action="add-scenario"]').isDisabled();
const headerCellsAtCap = await compareTable.locator("thead tr").first().locator("th").count();
const chartLinesAtCap = await tab.locator(".compare-results svg .line").count();
console.log(
  `cards at cap    : ${cardsAtCap}, add disabled: ${addDisabled}, header cells: ${headerCellsAtCap}, ` +
    `chart lines: ${chartLinesAtCap}`,
);
if (cardsAtCap !== 4) {
  problems.push(`adding scenarios up to the cap left ${cardsAtCap} cards, expected 4`);
}
if (!addDisabled) {
  problems.push("the add-scenario button is not disabled at the four-card cap");
}
if (headerCellsAtCap !== 5) {
  problems.push(
    `the compare table header has ${headerCellsAtCap} cells at the cap, expected 5 (corner + 4 scenarios)`,
  );
}
if (chartLinesAtCap !== 4) {
  problems.push(`the compare chart draws ${chartLinesAtCap} lines at the cap, expected 4`);
}

// Remove down to the floor -- baseline plus one variant -- then confirm
// remove disables, and the table/chart both track the count back down.
const removeButton = tab.locator('.compare-card:not([data-scenario="baseline"]) [data-action="remove-scenario"]');
await removeButton.first().click();
await tab.waitForTimeout(50);
await removeButton.first().click();
await tab.waitForTimeout(50);
const cardsAtFloor = await tab.locator(".compare-card").count();
const removeDisabled = await removeButton.first().isDisabled();
const chartLinesAtFloor = await tab.locator(".compare-results svg .line").count();
console.log(`cards at floor  : ${cardsAtFloor}, remove disabled: ${removeDisabled}, chart lines: ${chartLinesAtFloor}`);
if (cardsAtFloor !== 2) {
  problems.push(`removing scenarios down to the floor left ${cardsAtFloor} cards, expected 2`);
}
if (!removeDisabled) {
  problems.push("the remove-scenario button is not disabled with only one scenario left");
}
if (chartLinesAtFloor !== 2) {
  problems.push(`the compare chart draws ${chartLinesAtFloor} lines at the floor, expected 2`);
}

// Back to the single-scenario view for the screenshots, and to leave the
// page as found.
await tab.locator('.screen-toggle .panel-btn[data-screen="single"]').click();
await tab.waitForTimeout(50);

// ---- Provenance and paper ------------------------------------------------
//
// The page computes a pension forecast and looks authoritative, so it has to
// say whose model it is; and people print a forecast to take to a meeting, so
// paper is a real output rather than an afterthought.

const notice = tab.locator('[data-role="disclaimer"]');
const noticeText = (await notice.count()) > 0 ? ((await notice.textContent()) ?? "") : "";
if (!/inofficiell|unofficial/i.test(noticeText)) {
  problems.push("the page does not say it is an unofficial version");
}
if (!noticeText.includes("Pensionsmyndigheten")) {
  problems.push("the disclaimer does not name Pensionsmyndigheten");
}
if (!noticeText.includes("typfallsmodellen@pensionsmyndigheten.se")) {
  problems.push("the disclaimer does not give the agency's address for model questions");
}

const closedOnScreen = await tab.locator("details:not([open])").count();
await tab.emulateMedia({ media: "print" });
await tab.waitForTimeout(120);

/** One computed style under the print stylesheet. */
async function printed(selector, property) {
  const target = tab.locator(selector).first();
  if ((await target.count()) === 0) return undefined;
  return target.evaluate((node, p) => getComputedStyle(node)[p], property);
}

for (const [label, selector, property, want] of [
  ["the language chips", ".langs", "display", "none"],
  ["the mode toggle", ".mode-row", "display", "none"],
  ["the CSV button", ".export-btn", "display", "none"],
  ["the disclaimer", '[data-role="disclaimer"]', "display", "block"],
  ["the table scroller", ".scroll", "overflowX", "visible"],
  ["the input column", ".input-column", "position", "static"],
  // A dark-theme page would otherwise print as a black rectangle.
  ["the page ground", "body", "backgroundColor", "rgb(255, 255, 255)"],
]) {
  const got = await printed(selector, property);
  if (got !== want) problems.push(`in print, ${label} has ${property} ${got}, expected ${want}`);
}

// Every disclosure opens for print -- a collapsed <details> on paper withholds
// what it covers from a reader who cannot click it. CSS cannot do this, so the
// check is that the box actually lays out, not merely that `display` computes.
const closedInPrint = await tab.locator("details:not([open])").count();
if (closedInPrint !== 0) {
  problems.push(`${closedInPrint} disclosure(s) would print collapsed`);
}
const disclosureBox = await tab.locator(".field-note-body").first().boundingBox();
if (!disclosureBox || disclosureBox.height === 0) {
  problems.push("a disclosure's body lays out zero-height under print");
}
console.log(`print           : ${closedOnScreen} disclosures opened, controls hidden, light ground`);

await tab.emulateMedia({ media: "screen" });
await tab.waitForTimeout(120);
const closedAfter = await tab.locator("details:not([open])").count();
if (closedAfter !== closedOnScreen) {
  problems.push(`printing left ${closedAfter} disclosures closed, expected ${closedOnScreen} as before`);
}

// Ordlista: the workbook's own glossary sheet, surfaced as a disclosure in
// the left column below "Om pris- och avkastningsantaganden" -- the terms
// themselves stay Swedish always (the sheet has no English column), but the
// disclosure's own summary and a Swedish-only note still switch with the
// page's language, like every other piece of this port's own UI chrome.
const glossaryDetails = tab.locator("details", { has: tab.locator(".glossary") });
const glossaryTerms = await glossaryDetails.locator(".glossary dt").count();
console.log(`glossary terms  : ${glossaryTerms}`);
if (glossaryTerms < 50) {
  problems.push(`Ordlista shows ${glossaryTerms} terms, expected around 57`);
}
const firstTermSv = await glossaryDetails.locator(".glossary dt").first().textContent();
if (firstTermSv !== "Administrationsavgift") {
  problems.push(`Ordlista's first term is "${firstTermSv}", expected "Administrationsavgift" (alphabetical)`);
}
const glossarySummarySv = await glossaryDetails.locator("summary").textContent();
if (glossarySummarySv !== "Ordlista") {
  problems.push(`Ordlista's Swedish heading reads "${glossarySummarySv}", expected "Ordlista"`);
}

await tab.getByRole("button", { name: "EN", exact: true }).click();
await tab.waitForTimeout(50);
const glossarySummaryEn = await glossaryDetails.locator("summary").textContent();
if (glossarySummaryEn !== "Glossary") {
  problems.push(`Ordlista's English heading reads "${glossarySummaryEn}", expected "Glossary"`);
}
const glossaryNoteEn = await glossaryDetails.locator(".field-note-body > p").first().textContent();
if (!glossaryNoteEn || !/only available in Swedish/i.test(glossaryNoteEn)) {
  problems.push(
    `Ordlista's English note reads "${glossaryNoteEn}", expected it to say the glossary is Swedish-only`,
  );
}
const firstTermEn = await glossaryDetails.locator(".glossary dt").first().textContent();
if (firstTermEn !== "Administrationsavgift") {
  problems.push("Ordlista's own terms changed with the language, expected them to stay Swedish");
}
await tab.getByRole("button", { name: "SV", exact: true }).click();
await tab.waitForTimeout(50);

if (shots) {
  mkdirSync(shots, { recursive: true });
  await tab.screenshot({ path: join(shots, "sv.png"), fullPage: true });
  // Exact match, or this also resolves "Jämför scenarier" (contains "en").
  await tab.getByRole("button", { name: "EN", exact: true }).click();
  await tab.waitForTimeout(150);
  await tab.screenshot({ path: join(shots, "en.png"), fullPage: true });
  await tab.emulateMedia({ colorScheme: "dark" });
  await tab.screenshot({ path: join(shots, "dark.png"), fullPage: true });
  console.log(`screenshots     : ${shots}`);
}

// The PGB dialog's own fit was checked above at this suite's standard 1280px
// desktop width, where its ~480px cap already clears the grid's old 380px
// floor on its own -- not a meaningful check of `.pgb-dialog .pgb-grid {
// min-width: 0; }` specifically, since that rule only matters once the
// dialog itself is narrower than the floor it removes, which only happens on
// a phone. Re-checked here at 375px (iPhone SE, the narrowest of the widths
// this was hand-verified at before this check existed: 375, 390, 1280) in its
// own short-lived context, rather than resizing `tab` and disturbing every
// assertion above that assumes the 1280px layout.
const narrowContext = await browser.newContext({ viewport: { width: 375, height: 700 } });
const narrowTab = await narrowContext.newPage();
await narrowTab.goto(pathToFileURL(page).href);
await narrowTab.waitForSelector("#app");
await narrowTab.locator('.mode-toggle .panel-btn[data-mode="advanced"]').click();
await narrowTab.locator('.adv-group[data-group="pgb"] summary').click();
await narrowTab.locator('[data-action="pgb-expand"]').click();
await narrowTab.waitForTimeout(120);
const narrowMetrics = await narrowTab.locator(".pgb-dialog .adv-grid-scroll").evaluate((el) => ({
  scrollWidth: el.scrollWidth,
  clientWidth: el.clientWidth,
}));
console.log(
  `pgb dialog @375px: scrollWidth ${narrowMetrics.scrollWidth} <= clientWidth ${narrowMetrics.clientWidth}?`,
);
if (narrowMetrics.scrollWidth > narrowMetrics.clientWidth) {
  problems.push(
    `at a 375px phone width, the PGB dialog still needs horizontal scroll (scrollWidth ` +
      `${narrowMetrics.scrollWidth} > clientWidth ${narrowMetrics.clientWidth})`,
  );
}
await narrowContext.close();

await browser.close();

console.log(`outbound blocked: ${outbound.length > 0 ? outbound.join(", ") : "(none)"}`);
console.log(`page errors     : ${errors.length > 0 ? errors.join(" | ") : "(none)"}`);

if (outbound.length > 0) problems.push(`the page asked for ${outbound.length} thing(s) off-file`);
if (errors.length > 0) problems.push(`${errors.length} page error(s)`);

if (problems.length > 0) {
  console.error(`\nverify-offline FAILED:\n  ${problems.join("\n  ")}`);
  process.exit(1);
}
console.log("\nverify-offline: the built file runs from disk with no network and the right numbers.");

function fail(message) {
  console.error(`verify-offline: ${message}`);
  process.exit(1);
}
