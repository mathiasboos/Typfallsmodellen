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
// Swedish decimal comma, not the source file's own JS-literal period -- the
// field renders in the page's own language, which defaults to "sv" here.
const kommunalskattExpected = String(kommunalskattSource).replace(".", ",");
if (kommunalskattValue !== kommunalskattExpected) {
  problems.push(
    `picking Danderyd set kommunalskatt to "${kommunalskattValue}", expected the exact published ` +
      `rate "${kommunalskattExpected}"`,
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
const expectedBurialOnlyStr = String(Math.round(expectedBurialOnly * 100_000) / 1000).replace(".", ",");
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
const expectedChurchMemberStr = String(Math.round(expectedChurchMember * 100_000) / 1000).replace(".", ",");
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


// PGB: a compact add-entry form (pick a type, fill in that type's own
// fields, click "Lägg till") plus a summary table that only shows years and
// categories that actually have data -- replacing the old always-visible
// 55-row grid. The shipped workbook has no manual PGB entries at all
// (`pgbManual`'s own comment: "childcare years are the only PGB a default
// run earns"), so an untouched panel must show nothing, and each add below
// must move both the table and the pension.
const pgbGroup = tab.locator('details[data-group="pgb"]');
if (!(await pgbGroup.evaluate((node) => node.open))) {
  await pgbGroup.locator("summary").click();
}
const PGB_TYPE = { child: "0", conscription: "1", sickness: "2", study: "3" };
async function pgbType(type) {
  await pgbGroup.locator('[data-setting="pgbEntryType"]').selectOption(PGB_TYPE[type]);
  await tab.waitForTimeout(30);
}
async function pgbYear(year) {
  const el = pgbGroup.locator('[data-setting="pgbEntryYear"]');
  await el.fill(String(year));
  await el.dispatchEvent("change");
}
async function pgbAdd() {
  await pgbGroup.locator('[data-action="pgb-add-entry"]').click();
  await tab.waitForTimeout(80);
}
// Scoped to the whole tab, not `pgbGroup`: "Visa alla kolumner" (below)
// moves the table -- and the paragraph beside it -- out into a floating
// panel appended to `document.body`, so a `pgbGroup`-scoped locator would
// stop finding either while it's expanded.
const pgbSummaryRows = tab.locator(".pgb-summary-table tbody tr");
async function pgbHeads() {
  return tab.locator(".pgb-summary-table thead th").allTextContents();
}
const pgbSummaryEmptyMsg = tab.locator(".pgb-summary-empty");

// 1. Empty state: nothing typed yet.
const pgbEmptyBefore = await pgbSummaryEmptyMsg.isVisible();
const pgbRowsBefore = await pgbSummaryRows.count();
console.log(`pgb empty state : empty message visible=${pgbEmptyBefore}, rows=${pgbRowsBefore}`);
if (!pgbEmptyBefore || pgbRowsBefore !== 0) {
  problems.push(
    `the PGB panel shows ${pgbRowsBefore} row(s) with the empty message visible=${pgbEmptyBefore} before any ` +
      "entry, expected 0 rows and the empty message",
  );
}

// The form's own fields show and hide with the selected type -- "Barn" (the
// default) shows the birth-year field and the child-slot select; switching
// to sickness swaps in the kronor field instead.
const pgbYearField = pgbGroup.locator('[data-setting="pgbEntryYear"]');
const pgbSlotField = pgbGroup.locator('[data-setting="pgbEntryChildSlot"]');
const pgbAmountField = pgbGroup.locator('[data-setting="pgbEntryAmount"]');
const pgbSemesterField = pgbGroup.locator('[data-setting="pgbEntryStudySemesters"]');
const pgbStartField = pgbGroup.locator('[data-setting="pgbEntryConscriptionStart"]');
if (
  !(await pgbYearField.isVisible()) ||
  !(await pgbSlotField.isVisible()) ||
  (await pgbAmountField.isVisible())
) {
  problems.push(
    'the PGB form does not open on "Barn" with the birth-year field and child slot shown, amount hidden',
  );
}
await pgbType("sickness");
if (
  (await pgbSlotField.isVisible()) ||
  !(await pgbAmountField.isVisible()) ||
  (await pgbStartField.isVisible())
) {
  problems.push(
    'switching the PGB type to "Sjuk-/aktivitetsersättning" did not show the amount field and hide the ' +
      "child slot and conscription dates",
  );
}

// A year whose age falls outside 16-70 is rejected outright, not silently
// moved to the nearest valid one -- 2030 is age 71 for this typfall's 1959
// birth year, one past `earnPgb`'s own riktålder-based ceiling, so there is
// visible feedback and no row rather than a different year quietly earning
// the credit instead.
const pgbAddFeedback = pgbGroup.locator(".pgb-add-feedback");
await pgbYear(2030);
await pgbAmountField.fill("50000");
await pgbAmountField.dispatchEvent("change");
await pgbAdd();
const pgbRowsAfterBadYear = await pgbSummaryRows.count();
const pgbFeedbackBadYear = await pgbAddFeedback.textContent();
console.log(`pgb bad sa year : rows=${pgbRowsAfterBadYear}, feedback="${pgbFeedbackBadYear}"`);
if (pgbRowsAfterBadYear !== 0) {
  problems.push(`adding a sickness entry for an out-of-range year created ${pgbRowsAfterBadYear} row(s), expected 0 (rejected)`);
}
if (!pgbFeedbackBadYear || !/16.*70|16-70/.test(pgbFeedbackBadYear)) {
  problems.push(`adding a sickness entry for an out-of-range year shows feedback "${pgbFeedbackBadYear}", expected a message naming the 16-70 range`);
}

// 2. A sickness entry: one new row, the pension rises, and a category only
// gets a column once something is actually in it.
const pensionBeforeSa = await kpiValue(0);
await pgbYear(2005); // age 46 for this typfall's 1959 birth year
const pgbAgeReadout = await pgbGroup.locator(".pgb-year-age").textContent();
if (pgbAgeReadout !== "46") {
  problems.push(`typing year 2005 shows age readout "${pgbAgeReadout}", expected "46" (born 1959)`);
}
await pgbAmountField.fill("200000");
await pgbAmountField.dispatchEvent("change");
await pgbAdd();
const pgbRowsAfterSa = await pgbSummaryRows.count();
const pgbHeadsAfterSa = await pgbHeads();
console.log(`pgb rows w/ sa  : ${pgbRowsAfterSa}, heads: ${pgbHeadsAfterSa.join(" | ")}`);
if (pgbRowsAfterSa !== 1) {
  problems.push(`adding a sickness entry left ${pgbRowsAfterSa} summary row(s), expected 1`);
}
if (!pgbHeadsAfterSa.some((h) => /sjuk/i.test(h))) {
  problems.push(`the summary table has no "Sjuk-" column after a sickness entry: ${pgbHeadsAfterSa.join(" | ")}`);
}
if (pgbHeadsAfterSa.some((h) => /studier|värnplikt|barn/i.test(h))) {
  problems.push(`the summary table shows an unused column with only a sickness entry present: ${pgbHeadsAfterSa.join(" | ")}`);
}
const pensionAfterSa = await kpiValue(0);
console.log(`pension w/ sa   : ${pensionBeforeSa} -> ${pensionAfterSa} kr after a sickness entry`);
if (!(pensionAfterSa > pensionBeforeSa)) {
  problems.push(`entering a sickness-compensation amount did not raise the pension (${pensionBeforeSa} -> ${pensionAfterSa})`);
}

// "Visa alla kolumner" moves the same live table (not a copy) into a
// floating panel rather than a modal `<dialog>` -- nothing about the rest
// of the page is made inert, so the add-entry form just above has to keep
// working while it's open. Proven by doing check 3, below, with the panel
// expanded the whole time, rather than as a separate step of its own.
// Scoped to the whole tab: the button lives inside the same node that
// moves into the floating panel, so it travels there with the table too --
// still the way to collapse it again, just no longer under `pgbGroup`.
const pgbExpandBtn = tab.locator('[data-action="pgb-expand"]');
const pgbDrawer = tab.locator(".pgb-summary-drawer");
await pgbExpandBtn.click();
await tab.waitForTimeout(80);
if (!(await pgbDrawer.isVisible())) {
  problems.push('"Visa alla kolumner" did not show the floating PGB summary panel');
}
if ((await pgbDrawer.locator(".pgb-summary-table").count()) !== 1) {
  problems.push("the floating panel does not hold the PGB summary table -- expected the same live table, moved");
}

// 3. A study entry in a different year: its own column, a second row, a
// further rise. 2005 and 2010 both postdate 1995, the year study-PGB
// actually started. Added here with the table still expanded (see above).
await pgbType("study");
if (!(await pgbYearField.isVisible())) {
  problems.push("the add-entry form's own Year field is not usable while the PGB table is expanded");
}
await pgbYear(2010); // age 51
await pgbSemesterField.fill("1");
await pgbSemesterField.dispatchEvent("change");
await pgbAdd();
const pgbRowsAfterStudy = await pgbSummaryRows.count();
const pgbHeadsAfterStudy = await pgbHeads();
console.log(`pgb rows w/study: ${pgbRowsAfterStudy}, heads: ${pgbHeadsAfterStudy.join(" | ")}`);
if (pgbRowsAfterStudy !== 2) {
  problems.push(`adding a study entry in a new year left ${pgbRowsAfterStudy} summary row(s), expected 2`);
}
if (!pgbHeadsAfterStudy.some((h) => /studier/i.test(h))) {
  problems.push(`the summary table has no "Studier" column after a study entry: ${pgbHeadsAfterStudy.join(" | ")}`);
}
const pensionAfterStudy = await kpiValue(0);
console.log(`pension w/ study: ${pensionAfterSa} -> ${pensionAfterStudy} kr after a study entry`);
if (!(pensionAfterStudy > pensionAfterSa)) {
  problems.push(`entering a study semester did not raise the pension further (${pensionAfterSa} -> ${pensionAfterStudy})`);
}

// Collapsing puts the very same table (the row just added included) back
// in its normal place in the sidebar, not a stale copy.
await pgbExpandBtn.click();
await tab.waitForTimeout(80);
if (await pgbDrawer.isVisible()) {
  problems.push('"Dölj tabellen" did not hide the floating PGB summary panel');
}
if ((await pgbGroup.locator(".pgb-summary-table").count()) !== 1) {
  problems.push("collapsing the PGB summary panel did not put the table back under its own group");
}
if ((await pgbSummaryRows.count()) !== pgbRowsAfterStudy) {
  problems.push("collapsing the PGB summary panel lost the row added while it was expanded");
}

// 4. A study entry in the SAME year as the sickness one: still one row for
// 2005, now with two populated cells -- the mixed-category-row case.
await pgbYear(2005);
await pgbSemesterField.fill("1");
await pgbSemesterField.dispatchEvent("change");
await pgbAdd();
const pgbRowsAfterMixed = await pgbSummaryRows.count();
console.log(`pgb rows mixed  : ${pgbRowsAfterMixed}, expected still 2 (2005 gained a second figure, not a new row)`);
if (pgbRowsAfterMixed !== 2) {
  problems.push(`adding a study entry for 2005 (already a sickness year) left ${pgbRowsAfterMixed} rows, expected 2`);
}
const row2005 = tab.locator('.pgb-summary-table tbody tr[data-year="2005"]');
const row2005Cells = await row2005.locator("td").allTextContents();
console.log(`2005 row        : ${row2005Cells.join(" | ")}`);
const row2005NonEmpty = row2005Cells.slice(2, -1).filter((c) => c.trim() !== "").length;
if (row2005NonEmpty !== 2) {
  problems.push(`the 2005 row has ${row2005NonEmpty} populated category cell(s), expected 2 (sickness and study)`);
}

// A per-cell remove clears just that one category, not the whole row.
await row2005.locator('[data-action="pgb-remove"][data-category="studier"]').click();
await tab.waitForTimeout(80);
const row2005AfterRemove = await row2005.locator("td").allTextContents();
const pgbRowsAfterRemove = await pgbSummaryRows.count();
console.log(`2005 after x    : ${row2005AfterRemove.join(" | ")}, rows: ${pgbRowsAfterRemove}`);
if (pgbRowsAfterRemove !== 2) {
  problems.push(`removing 2005's study figure changed the row count to ${pgbRowsAfterRemove}, expected 2 (2005's sickness figure keeps the row)`);
}
const row2005NonEmptyAfter = row2005AfterRemove.slice(2, -1).filter((c) => c.trim() !== "").length;
if (row2005NonEmptyAfter !== 1) {
  problems.push(`removing 2005's study figure left ${row2005NonEmptyAfter} populated cell(s), expected 1 (sickness only)`);
}

// 5. A conscription period under 120 days adds no row -- the one case the
// table cannot show a zero in on its own, so a standalone readout covers it.
const pensionBeforeVpl = await kpiValue(0);
await pgbType("conscription");
await pgbStartField.fill("1998-01-01");
await pgbStartField.dispatchEvent("change");
const pgbEndField = pgbGroup.locator('[data-setting="pgbEntryConscriptionEnd"]');
await pgbEndField.fill("1998-02-01"); // 31 days, under the 120 minimum
await pgbEndField.dispatchEvent("change");
await tab.waitForTimeout(80);
const vplShortReadout = await pgbGroup.locator(".pgb-conscription-readout").textContent();
console.log(`vpl short period: "${vplShortReadout}"`);
if (!vplShortReadout || !/120/.test(vplShortReadout)) {
  problems.push(`a conscription period under 120 days reads "${vplShortReadout}", expected the "under 120 days" warning`);
}
await pgbAdd();
const pgbRowsAfterShortVpl = await pgbSummaryRows.count();
const pensionAfterShortVpl = await kpiValue(0);
if (pgbRowsAfterShortVpl !== 2 || pensionAfterShortVpl !== pensionBeforeVpl) {
  problems.push(
    `adding a conscription period under 120 days left ${pgbRowsAfterShortVpl} rows and pension ` +
      `${pensionAfterShortVpl} (was ${pensionBeforeVpl}), expected no change`,
  );
}

// A period long enough to matter is not the same as an eligible one --
// `wsPGB!F`'s own window is 1995-2010 and from 2018 on, and 2012 falls in
// neither. The readout catches this while typing, before "Lägg till" is
// even clicked, the same way it already catches a too-short period.
await pgbStartField.fill("2012-01-01");
await pgbStartField.dispatchEvent("change");
await pgbEndField.fill("2012-12-31"); // 366 days, well over 120 -- long enough, just not eligible
await pgbEndField.dispatchEvent("change");
await tab.waitForTimeout(80);
const vplIneligibleReadout = await pgbGroup.locator(".pgb-conscription-readout").textContent();
console.log(`vpl 2012 window : "${vplIneligibleReadout}"`);
if (!vplIneligibleReadout || !/1995|2018/.test(vplIneligibleReadout)) {
  problems.push(`a conscription period outside 1995-2010/2018+ reads "${vplIneligibleReadout}", expected a message naming the eligible window`);
}
await pgbAdd();
const pgbRowsAfterIneligibleVpl = await pgbSummaryRows.count();
const pensionAfterIneligibleVpl = await kpiValue(0);
if (pgbRowsAfterIneligibleVpl !== 2 || pensionAfterIneligibleVpl !== pensionBeforeVpl) {
  problems.push(
    `adding a conscription period outside the eligible window left ${pgbRowsAfterIneligibleVpl} rows and pension ` +
      `${pensionAfterIneligibleVpl} (was ${pensionBeforeVpl}), expected no change`,
  );
}

// 6. A valid conscription period: its own column, its own row, a rise.
await pgbType("conscription");
await pgbStartField.fill("1998-01-01"); // back to the eligible period this suite carries forward
await pgbStartField.dispatchEvent("change");
await pgbEndField.fill("1998-12-31"); // now a full year, well over 120 days
await pgbEndField.dispatchEvent("change");
await tab.waitForTimeout(80);
const vplReadout = await pgbGroup.locator(".pgb-conscription-readout").textContent();
if (vplReadout !== "") {
  problems.push(`the conscription readout reads "${vplReadout}" for a valid period, expected empty`);
}
await pgbAdd();
const pgbRowsAfterVpl = await pgbSummaryRows.count();
const pgbHeadsAfterVpl = await pgbHeads();
console.log(`pgb rows w/ vpl : ${pgbRowsAfterVpl}, heads: ${pgbHeadsAfterVpl.join(" | ")}`);
if (pgbRowsAfterVpl !== 3) {
  problems.push(`adding a valid conscription period left ${pgbRowsAfterVpl} rows, expected 3`);
}
if (!pgbHeadsAfterVpl.some((h) => /värnplikt/i.test(h))) {
  problems.push(`the summary table has no "Värnplikt" column after a conscription entry: ${pgbHeadsAfterVpl.join(" | ")}`);
}
const pensionAfterVpl = await kpiValue(0);
console.log(`pension w/ vpl  : ${pensionBeforeVpl} -> ${pensionAfterVpl} kr after a conscription period`);
if (!(pensionAfterVpl > pensionBeforeVpl)) {
  problems.push(`entering a conscription date range did not raise the pension (${pensionBeforeVpl} -> ${pensionAfterVpl})`);
}

// 7. Removing the conscription period: its row disappears (nothing else
// populated 1998), the column disappears, the pension falls back.
const vplClear = pgbGroup.locator('[data-action="pgb-clear-conscription"]');
if (!(await vplClear.isVisible())) {
  problems.push('the conscription "Rensa" button is not visible with a period entered');
}
await vplClear.click();
await tab.waitForTimeout(80);
const pgbRowsAfterClear = await pgbSummaryRows.count();
const pgbHeadsAfterClear = await pgbHeads();
const pensionAfterClear = await kpiValue(0);
console.log(`pgb after clear : rows=${pgbRowsAfterClear}, heads: ${pgbHeadsAfterClear.join(" | ")}, pension=${pensionAfterClear}`);
if (pgbRowsAfterClear !== 2) {
  problems.push(`clearing the conscription period left ${pgbRowsAfterClear} rows, expected 2 (1998 had nothing else)`);
}
if (pgbHeadsAfterClear.some((h) => /värnplikt/i.test(h))) {
  problems.push(`clearing the conscription period left a "Värnplikt" column: ${pgbHeadsAfterClear.join(" | ")}`);
}
if (pensionAfterClear !== pensionBeforeVpl) {
  problems.push(`clearing the conscription period left the pension at ${pensionAfterClear}, expected it back at ${pensionBeforeVpl}`);
}
const vplStartAfterClear = await pgbStartField.inputValue();
const vplEndAfterClear = await pgbEndField.inputValue();
if (vplStartAfterClear !== "" || vplEndAfterClear !== "") {
  problems.push(`clearing the conscription period left the dates at "${vplStartAfterClear}"/"${vplEndAfterClear}", expected both empty`);
}
if (await vplClear.isVisible()) {
  problems.push('the conscription "Rensa" button is still visible after clearing, expected hidden');
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

// A birth year too early for the parent to have been at least 16 earns
// nothing (`pgbBarn`'s own gate) -- committed to the slot regardless (it is
// still that child's own birth year, just an ineligible one), but with
// visible feedback and no summary row, rather than the slot silently doing
// nothing with no sign why. Removed again once checked, freeing slot 1 for
// the real child below.
const pgbRowsBeforeBadChild = await pgbSummaryRows.count();
await pgbType("child");
await pgbYear(1965); // age 6 for this typfall's 1959 birth year -- too young to be a parent
await pgbAdd();
const pgbRowsAfterBadChild = await pgbSummaryRows.count();
const pgbFeedbackBadChild = await pgbAddFeedback.textContent();
console.log(`pgb bad child yr: rows=${pgbRowsAfterBadChild}, feedback="${pgbFeedbackBadChild}"`);
if (pgbRowsAfterBadChild !== pgbRowsBeforeBadChild) {
  problems.push(`adding an ineligible child birth year changed the row count (${pgbRowsBeforeBadChild} -> ${pgbRowsAfterBadChild}), expected no change`);
}
if (!pgbFeedbackBadChild || !/16/.test(pgbFeedbackBadChild)) {
  problems.push(`adding an ineligible child birth year shows feedback "${pgbFeedbackBadChild}", expected a message naming the age-16 rule`);
}
await pgbGroup.locator('.pgb-children-list [data-action="pgb-remove-child"][data-child-slot="1"]').click();
await tab.waitForTimeout(50);
// Freeing slot 1 does not by itself move the select back to it (only
// filling the *selected* slot advances it) -- picked explicitly so the
// real child below lands in slot 1, matching this check's own assertions.
await pgbSlotField.selectOption("1");

// 8. A child's birth year: the credit can land in up to four years, so this
// can add up to four new rows, and needs a real engine change to show at all
// -- `RunState.pgbBarn`/`PgbBreakdownYear.barn` are new for this panel; the
// old grid folded the credit straight into `RunState.pgb` with no way to
// tell it apart from the other three sources. Checked with a birth year old
// enough that `pgbBarn`'s own gate (the parent's age at the birth, > 15) is
// satisfied for the default typfall (born 1959). Run here, in the same
// "last, right before the full reset" spot every other test that raises the
// model's own pension figures runs in.
const pensionBeforeChild = await kpiValue(0);
const pgbRowsBeforeChild = await pgbSummaryRows.count();
await pgbType("child");
await pgbYear(1990);
await pgbAdd();
const pgbRowsAfterChild = await pgbSummaryRows.count();
const pgbHeadsAfterChild = await pgbHeads();
console.log(`pgb rows w/child: ${pgbRowsBeforeChild} -> ${pgbRowsAfterChild}, heads: ${pgbHeadsAfterChild.join(" | ")}`);
if (!(pgbRowsAfterChild > pgbRowsBeforeChild)) {
  problems.push(`adding a child's birth year added no summary row(s) (${pgbRowsBeforeChild} -> ${pgbRowsAfterChild})`);
}
if (!pgbHeadsAfterChild.some((h) => /barn/i.test(h))) {
  problems.push(`the summary table has no "Barn" column after a child's birth year: ${pgbHeadsAfterChild.join(" | ")}`);
}
const pensionAfterChild = await kpiValue(0);
console.log(`pension w/ child: ${pensionBeforeChild} -> ${pensionAfterChild} kr after a child's birth year`);
if (!(pensionAfterChild > pensionBeforeChild)) {
  problems.push(`adding a child's birth year did not raise the pension (${pensionBeforeChild} -> ${pensionAfterChild})`);
}
if (!(await pgbGroup.locator('.pgb-children-list > [data-child-slot="1"]').isVisible())) {
  problems.push("adding a child's birth year did not show it in the children list");
}

// The slot select won't offer an already-filled slot, and Add gives up once
// all four are taken.
const pgbSlotOptions = await pgbSlotField
  .locator("option")
  .evaluateAll((opts) => opts.map((o) => ({ value: o.value, disabled: o.disabled })));
if (pgbSlotOptions.find((o) => o.value === "1")?.disabled !== true) {
  problems.push("the child-slot select still offers slot 1 after it was filled");
}
for (const year of [1993, 1996, 1999]) {
  await pgbYear(year);
  await pgbAdd();
}
const pgbAddButton = pgbGroup.locator('[data-action="pgb-add-entry"]');
if (!(await pgbAddButton.isDisabled())) {
  problems.push("the PGB Add button is not disabled with all four child slots filled");
}

// 9. Removing a child: back down to slot 1's own 1990, then remove that too
// -- the rows it alone earned disappear, and the column with them.
for (const slot of [4, 3, 2]) {
  await pgbGroup.locator(`.pgb-children-list [data-action="pgb-remove-child"][data-child-slot="${slot}"]`).click();
  await tab.waitForTimeout(50);
}
if (await pgbAddButton.isDisabled()) {
  problems.push("the PGB Add button is still disabled after freeing three of the four child slots");
}
await pgbGroup.locator('.pgb-children-list [data-action="pgb-remove-child"][data-child-slot="1"]').click();
await tab.waitForTimeout(80);
const pgbRowsAfterRemoveChild = await pgbSummaryRows.count();
const pgbHeadsAfterRemoveChild = await pgbHeads();
const pensionAfterRemoveChild = await kpiValue(0);
console.log(
  `pgb after remove child: rows=${pgbRowsAfterRemoveChild}, heads: ${pgbHeadsAfterRemoveChild.join(" | ")}, ` +
    `pension=${pensionAfterRemoveChild}`,
);
if (pgbRowsAfterRemoveChild !== pgbRowsBeforeChild) {
  problems.push(`removing the last child left ${pgbRowsAfterRemoveChild} rows, expected back to ${pgbRowsBeforeChild}`);
}
if (pgbHeadsAfterRemoveChild.some((h) => /barn/i.test(h))) {
  problems.push(`removing the last child left a "Barn" column: ${pgbHeadsAfterRemoveChild.join(" | ")}`);
}
if (pensionAfterRemoveChild !== pensionBeforeChild) {
  problems.push(`removing the last child left the pension at ${pensionAfterRemoveChild}, expected back at ${pensionBeforeChild}`);
}
if (await pgbGroup.locator(".pgb-children").isVisible()) {
  problems.push("the children list is still visible after every slot was emptied");
}

// 10. Column-hiding in general: only the categories with data show, in
// whichever combination is currently present -- checked here against what
// checks 2-9 above left behind (sickness and study, nothing else).
const pgbHeadsFinal = await pgbHeads();
console.log(`pgb heads final : ${pgbHeadsFinal.join(" | ")}`);
if (!pgbHeadsFinal.some((h) => /sjuk/i.test(h)) || !pgbHeadsFinal.some((h) => /studier/i.test(h))) {
  problems.push(`expected the sickness and study columns to remain: ${pgbHeadsFinal.join(" | ")}`);
}
if (pgbHeadsFinal.some((h) => /värnplikt|barn/i.test(h))) {
  problems.push(`expected no conscription or child column with neither present: ${pgbHeadsFinal.join(" | ")}`);
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

// 11. Aterstall clears the summary table, the form and the children list
// back to empty -- the same as every other advanced-mode field.
await tab.locator('[data-action="reset-advanced"]').click();
await tab.waitForTimeout(50);
const pgbRowsAfterReset = await pgbSummaryRows.count();
const pgbEmptyAfterReset = await pgbSummaryEmptyMsg.isVisible();
if (pgbRowsAfterReset !== 0 || !pgbEmptyAfterReset) {
  problems.push(
    `the reset button left ${pgbRowsAfterReset} PGB summary row(s) (empty message visible=${pgbEmptyAfterReset}), ` +
      "expected 0 and the empty message",
  );
}
const pgbYearAfterReset = await pgbYearField.inputValue();
const pgbAmountAfterReset = await pgbAmountField.inputValue();
if (pgbYearAfterReset !== "0" || pgbAmountAfterReset !== "0") {
  problems.push(`the reset button left the PGB form at year "${pgbYearAfterReset}"/amount "${pgbAmountAfterReset}", expected both "0"`);
}
const vplStartAfterReset = await pgbStartField.inputValue();
if (vplStartAfterReset !== "") {
  problems.push(`the reset button left the conscription start date at "${vplStartAfterReset}", expected empty`);
}
if (await pgbGroup.locator(".pgb-children").isVisible()) {
  problems.push("the reset button left the children list visible, expected empty and hidden");
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

// "Flexpension för ITP 1 och SAF-LO från och med 2014": an extra premium
// added to both agreements' own rates from 2014 onward -- already wired into
// the engine (itp.ts/safLo.ts) before this panel exposed it, so the check is
// that the field actually reaches `flexPension`, not that the maths is new.
// The default typfall carries no occupational scheme ("Saknar
// tjänstepension"), so this is the one place in the script that switches the
// normal-mode scheme select, to SAF-LO, and switches it back after -- nothing
// later may see it moved.
await tab.getByLabel("Välj tjänstepension").selectOption("4");
await tab.waitForTimeout(80);
const tjpBeforeFlex = await shown("tjp", "monthly");
const flexPensionField = await setting("flexPension");
await flexPensionField.fill("10");
await flexPensionField.dispatchEvent("change");
await tab.waitForTimeout(80);
const tjpAfterFlex = await shown("tjp", "monthly");
console.log(`flexpension     : tjp ${tjpBeforeFlex} -> ${tjpAfterFlex} kr/month at 10% (SAF-LO)`);
if (!(tjpAfterFlex > tjpBeforeFlex)) {
  problems.push(
    `setting flexpension to 10% did not raise SAF-LO's occupational pension (${tjpBeforeFlex} -> ${tjpAfterFlex})`,
  );
}
await tab.locator('[data-action="reset-advanced"]').click();
await tab.waitForTimeout(50);
const flexPensionAfterReset = await flexPensionField.inputValue();
if (flexPensionAfterReset !== "0") {
  problems.push(`the reset button left flexpension at "${flexPensionAfterReset}", expected "0"`);
}
await tab.getByLabel("Välj tjänstepension").selectOption("1");
await tab.waitForTimeout(50);

// "Slutlönens referensår efter pensioneringen" (manual 3.6, row 43): a plain
// year count, not a flag (advanced.ts's own comment on this setting explains
// why the row's "(1)->" shorthand reads like one but isn't). It is a narrower
// setting than its name suggests -- see advanced.ts's comment -- feeding only
// `adjustmentFactors`'s `beforeRetirement` factor (packages/engine/src/model/
// result.ts), which rescales the "Slutlön" row's own price-adjusted ("Fasta
// priser") column in Table 1. It does not delay when the pension itself is
// paid (Table 2), so this checks only the one figure it does move.
const finalSalaryAdjustedBefore = await shown("slutlon", "adjusted");
const pensionGapField = await setting("pensionSameYearAsFinalSalary");
await pensionGapField.fill("5");
await pensionGapField.dispatchEvent("change");
await tab.waitForTimeout(80);
const finalSalaryAdjustedAfter = await shown("slutlon", "adjusted");
console.log(
  `slutlön ref year: slutlön (fasta priser) ${finalSalaryAdjustedBefore} -> ${finalSalaryAdjustedAfter} at 5 years`,
);
if (finalSalaryAdjustedAfter === finalSalaryAdjustedBefore) {
  problems.push(
    'setting the "Slutlönens referensår" gap to 5 years did not change the adjusted "Slutlön" figure',
  );
}

// Back to normal mode for the screenshots, and to leave the page as found.
await tab.locator('[data-action="reset-advanced"]').click();
await tab.waitForTimeout(50);
const pensionGapAfterReset = await pensionGapField.inputValue();
if (pensionGapAfterReset !== "0") {
  problems.push(`the reset button left "Slutlönens referensår" at "${pensionGapAfterReset}", expected "0"`);
}
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
if (screenButtons !== 3) {
  problems.push(`screen toggle has ${screenButtons} buttons, expected 3 (Prognos / Jämför scenarier / Mikrosim)`);
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

// ---- Mikrosim --------------------------------------------------------------
//
// The workbook's own batch runner, reproduced as a third top-level view: each
// row is a fully independent typfall (not a diff against the form on the
// left, unlike Jämför scenarier's variants). Unlike the real sheet's own
// explicit-calculate batch runner, results here are live -- there is no
// "Beräkna" button anywhere in this build, and a row's own output columns
// fill in immediately on load, on every edit, on adding a row, and on a CSV
// import finishing. The point of each check is the same independence
// property Jämför scenarier's own checks establish -- one row's own inputs
// reach only that row's own run -- plus CSV import/export (no precedent
// elsewhere in this app) and the stacked chart below the table.

await tab.locator('.screen-toggle .panel-btn[data-screen="mikrosim"]').click();
await tab.waitForTimeout(50);

const mikrosimRow = (i) => tab.locator(".mikrosim-table tbody tr").nth(i);
const mikrosimCell = (i, col) => mikrosimRow(i).locator("td").nth(col);
// td indices: 0 row#, 1 status, 2-10 the nine inputs, 11-22 the twelve
// outputs, 23 remove -- see mikrosim.ts's own INPUT_COLUMNS/OUTPUT_COLUMNS
// order.
const MI_SALARY_TD = 5;
const MI_SCHEME_TD = 10;
const MI_FIRST_OUTPUT_TD = 11;

// There is no "Beräkna" button to look for -- confirm outright.
const mikrosimCalcButton = await tab.locator('[data-action="calculate-mikrosim"]').count();
if (mikrosimCalcButton !== 0) {
  problems.push(`found ${mikrosimCalcButton} "Beräkna"-style button(s) on the Mikrosim tab, expected none`);
}

// A single seed row, already computed with no click needed.
const mikrosimRowsAtStart = await tab.locator(".mikrosim-table tbody tr").count();
console.log(`mikrosim rows   : ${mikrosimRowsAtStart}`);
if (mikrosimRowsAtStart !== 1) {
  problems.push(`the Mikrosim tab starts with ${mikrosimRowsAtStart} row(s), expected 1`);
}
const mikrosimOutputsAtStart = await mikrosimCell(0, MI_FIRST_OUTPUT_TD).textContent();
console.log(`mikrosim seed row output (no click): "${mikrosimOutputsAtStart}"`);
if ((mikrosimOutputsAtStart ?? "").trim() === "") {
  problems.push("the seed Mikrosim row shows no output on load -- live recompute did not run");
}

// Adding rows: three clicks from one row should leave four, each already
// computed (no click needed).
await tab.locator('[data-action="add-mikrosim-row"]').click();
await tab.locator('[data-action="add-mikrosim-row"]').click();
await tab.locator('[data-action="add-mikrosim-row"]').click();
await tab.waitForTimeout(50);
const mikrosimRowsAfterAdd = await tab.locator(".mikrosim-table tbody tr").count();
console.log(`mikrosim rows   : ${mikrosimRowsAfterAdd} after adding three`);
if (mikrosimRowsAfterAdd !== 4) {
  problems.push(`adding three Mikrosim rows left ${mikrosimRowsAfterAdd}, expected 4`);
}
const mikrosimNewRowOutput = await mikrosimCell(3, MI_FIRST_OUTPUT_TD).textContent();
if ((mikrosimNewRowOutput ?? "").trim() === "") {
  problems.push("a freshly added Mikrosim row shows no output -- expected it computed immediately");
}

// Row 1 gets a distinct salary and scheme; rows 2-4 are left at their shared
// defaults. Editing alone -- no click anywhere -- has to update row 1's own
// output from its own inputs, while rows 2-4 keep agreeing with each other.
await mikrosimCell(0, MI_SALARY_TD).locator("input").fill("600000");
await mikrosimCell(0, MI_SALARY_TD).locator("input").dispatchEvent("change");
await mikrosimCell(0, MI_SCHEME_TD).locator("select").selectOption("3");
await tab.waitForTimeout(100);

const mikrosimRow1Salary = await mikrosimCell(0, MI_FIRST_OUTPUT_TD).textContent();
const mikrosimRow2Salary = await mikrosimCell(1, MI_FIRST_OUTPUT_TD).textContent();
const mikrosimRow3Salary = await mikrosimCell(2, MI_FIRST_OUTPUT_TD).textContent();
console.log(`mikrosim Slutlön: row1 ${mikrosimRow1Salary}, row2 ${mikrosimRow2Salary}, row3 ${mikrosimRow3Salary}`);
if ((mikrosimRow1Salary ?? "").trim() === "") {
  problems.push('editing row 1 left its own "Slutlön" blank -- expected an immediate recompute');
}
if (mikrosimRow1Salary === mikrosimRow2Salary) {
  problems.push(
    `row 1's own salary raise did not change its "Slutlön" relative to row 2's ("${mikrosimRow1Salary}" both)`,
  );
}
if (mikrosimRow2Salary !== mikrosimRow3Salary) {
  problems.push(
    `two untouched Mikrosim rows show different "Slutlön" ("${mikrosimRow2Salary}" vs "${mikrosimRow3Salary}")`,
  );
}

// Editing an input again immediately recomputes that row's own output to the
// new value -- not blank, not the old one -- without touching any other
// row's own already-computed figures.
await mikrosimCell(0, MI_SALARY_TD).locator("input").fill("650000");
await mikrosimCell(0, MI_SALARY_TD).locator("input").dispatchEvent("change");
await tab.waitForTimeout(100);
const mikrosimRow1AfterEdit = await mikrosimCell(0, MI_FIRST_OUTPUT_TD).textContent();
const mikrosimRow2AfterEdit = await mikrosimCell(1, MI_FIRST_OUTPUT_TD).textContent();
console.log(`mikrosim after edit: row1 "${mikrosimRow1AfterEdit}", row2 "${mikrosimRow2AfterEdit}"`);
if ((mikrosimRow1AfterEdit ?? "").trim() === "" || mikrosimRow1AfterEdit === mikrosimRow1Salary) {
  problems.push(
    `editing row 1's salary again left its own "Slutlön" as "${mikrosimRow1AfterEdit}", expected a fresh value`,
  );
}
if (mikrosimRow2AfterEdit !== mikrosimRow2Salary) {
  problems.push(
    `editing row 1's salary changed row 2's already-computed "Slutlön" ("${mikrosimRow2Salary}" -> ` +
      `"${mikrosimRow2AfterEdit}")`,
  );
}

// The chart below the table draws one bar group per row, live -- the same
// edit that just changed row 1's own table cells has to change its own bar.
const mikrosimChart = tab.locator(".mikrosim-chart");
const mikrosimBarsBeforeRemove = await mikrosimChart.locator("svg .bar").count();
const mikrosimLegendItems = await mikrosimChart.locator(".legend li").count();
console.log(`mikrosim chart  : ${mikrosimBarsBeforeRemove} bar segments, ${mikrosimLegendItems} legend entries`);
if (mikrosimBarsBeforeRemove === 0) {
  problems.push("the Mikrosim chart draws no bars at all");
}
if (mikrosimLegendItems === 0 || mikrosimLegendItems > 7) {
  problems.push(`the Mikrosim chart legend has ${mikrosimLegendItems} entries, expected 1-7`);
}

// Removing a row drops the count, keeps the remaining rows' own values, and
// the chart's own bar count drops with it.
await mikrosimRow(1).locator('[data-action="mikrosim-remove-row"]').click();
await tab.waitForTimeout(50);
const mikrosimRowsAfterRemove = await tab.locator(".mikrosim-table tbody tr").count();
const mikrosimBarsAfterRemove = await mikrosimChart.locator("svg .bar").count();
console.log(`mikrosim rows   : ${mikrosimRowsAfterRemove} after removing one, chart bars: ${mikrosimBarsAfterRemove}`);
if (mikrosimRowsAfterRemove !== 3) {
  problems.push(`removing a Mikrosim row left ${mikrosimRowsAfterRemove}, expected 3`);
}
if (mikrosimBarsAfterRemove >= mikrosimBarsBeforeRemove) {
  problems.push(
    `removing a Mikrosim row did not shrink the chart (bars ${mikrosimBarsBeforeRemove} -> ${mikrosimBarsAfterRemove})`,
  );
}

// CSV export: the canonical Swedish headers, byte for byte, plus row 1's own
// already-live-computed number -- no click needed before downloading either.
const mikrosimPanel = tab.locator(".mikrosim-panel");
const mikrosimCsv = await download(mikrosimPanel.getByRole("button", { name: "Ladda ner CSV" }));
const mikrosimCsvText = mikrosimCsv.buffer.toString("utf8");
console.log(`mikrosim CSV    : ${mikrosimCsv.filename}, ${mikrosimCsv.buffer.length} bytes`);
const expectedHeader =
  "Födelseår;Börjar arbeta vid ålder;Går i pension vid ålder;Årslön;Årlig inflation;Real tillväxt;" +
  "Real fondavkastning;Privat pensionssparande (med avdragsrätt);Välj tjänstepension;Slutlön;" +
  "Brutto-pension;Inkomstpension;Tilläggspension;Premiepension;Garanti-pension;P_tillägg;" +
  "Tjänstepension;Eget sparande;Efter skatt;Bostadstillägg + ÄFS;Disponibel inkomst";
if (!mikrosimCsvText.includes(expectedHeader)) {
  problems.push(`the Mikrosim CSV's header line doesn't match the workbook's own Mikrosim columns`);
}
if (!mikrosimCsvText.includes("650000")) {
  problems.push(`the Mikrosim CSV doesn't contain row 1's own salary (650000)`);
}

// CSV import replaces the table -- a file with the same nine columns in a
// scrambled order, matched by header text rather than position -- and the
// imported rows are computed immediately, no click needed.
const scrambledCsv =
  "Välj tjänstepension;Födelseår;Årslön;Börjar arbeta vid ålder;Går i pension vid ålder;" +
  "Årlig inflation;Real tillväxt;Real fondavkastning;Privat pensionssparande (med avdragsrätt)\n" +
  "2;1980;500000;22;65;0;0;0,017;0\n" +
  "4;1965;300000;19;66;0;0;0,017;0\n";
await tab
  .locator('[data-action="mikrosim-import-file"]')
  .setInputFiles({ name: "scrambled.csv", mimeType: "text/csv", buffer: Buffer.from(scrambledCsv, "utf8") });
await tab.waitForTimeout(100);
const mikrosimRowsAfterImport = await tab.locator(".mikrosim-table tbody tr").count();
const importedBorn = await mikrosimCell(0, 2).locator("input").inputValue();
const mikrosimImportedOutput = await mikrosimCell(0, MI_FIRST_OUTPUT_TD).textContent();
console.log(
  `mikrosim import : ${mikrosimRowsAfterImport} rows, row 1 born ${importedBorn}, output "${mikrosimImportedOutput}"`,
);
if (mikrosimRowsAfterImport !== 2) {
  problems.push(`importing a 2-row CSV left ${mikrosimRowsAfterImport} rows, expected 2 (replace, not append)`);
}
if (importedBorn !== "1980") {
  problems.push(`the imported row's birth year reads "${importedBorn}", expected "1980" (scrambled columns)`);
}
if ((mikrosimImportedOutput ?? "").trim() === "") {
  problems.push("an imported Mikrosim row shows no output -- expected it computed immediately, no click");
}

// A file missing a required column is refused outright, with the missing
// column named, and the table already on screen is left untouched.
await tab.locator('[data-action="mikrosim-import-file"]').setInputFiles({
  name: "missing-column.csv",
  mimeType: "text/csv",
  buffer: Buffer.from("Födelseår;Går i pension vid ålder\n1970;65\n", "utf8"),
});
await tab.waitForTimeout(100);
const mikrosimFileError = await tab.locator(".mikrosim-file-error").textContent();
console.log(`mikrosim file error: ${mikrosimFileError}`);
if (!(mikrosimFileError ?? "").includes("Årslön")) {
  problems.push(`importing a file missing "Årslön" did not name it in the error ("${mikrosimFileError}")`);
}
const mikrosimRowsAfterBadImport = await tab.locator(".mikrosim-table tbody tr").count();
if (mikrosimRowsAfterBadImport !== 2) {
  problems.push(
    `a refused import changed the row count to ${mikrosimRowsAfterBadImport}, expected the previous 2 to remain`,
  );
}

// A row with an invalid scheme is flagged, not dropped, and left blank --
// the good row next to it is still computed immediately, no click -- and
// fixing the flagged row's own scheme through its own dropdown recomputes it
// on the spot, proving an edit always gets a fresh attempt even though a
// bulk recompute (the import that just ran) left it alone.
await tab.locator('[data-action="mikrosim-import-file"]').setInputFiles({
  name: "one-bad-row.csv",
  mimeType: "text/csv",
  buffer: Buffer.from(
    "Födelseår;Börjar arbeta vid ålder;Går i pension vid ålder;Årslön;Årlig inflation;Real tillväxt;" +
      "Real fondavkastning;Privat pensionssparande (med avdragsrätt);Välj tjänstepension\n" +
      "1970;20;66;480000;0;0;0,017;0;3\n" +
      "1970;20;66;480000;0;0;0,017;0;99\n",
    "utf8",
  ),
});
await tab.waitForTimeout(100);
const mikrosimGoodRowOutput = await mikrosimCell(0, MI_FIRST_OUTPUT_TD).textContent();
const mikrosimBadRowOutput = await mikrosimCell(1, MI_FIRST_OUTPUT_TD).textContent();
const mikrosimBadRowFlagged = await mikrosimRow(1).locator(".mikrosim-error").count();
console.log(
  `mikrosim mixed  : good "${mikrosimGoodRowOutput}", bad "${mikrosimBadRowOutput}", flagged ${mikrosimBadRowFlagged}`,
);
if ((mikrosimGoodRowOutput ?? "").trim() === "") {
  problems.push("a valid row next to an invalid one was not computed immediately");
}
if ((mikrosimBadRowOutput ?? "").trim() !== "") {
  problems.push(`the row with an invalid scheme shows an output ("${mikrosimBadRowOutput}"), expected blank`);
}
if (mikrosimBadRowFlagged !== 1) {
  problems.push(`the row with an invalid scheme is not visibly flagged (found ${mikrosimBadRowFlagged} marker(s))`);
}

await mikrosimCell(1, MI_SCHEME_TD).locator("select").selectOption("2");
await tab.waitForTimeout(100);
const mikrosimFixedOutput = await mikrosimCell(1, MI_FIRST_OUTPUT_TD).textContent();
const mikrosimFixedFlag = await mikrosimRow(1).locator(".mikrosim-error").count();
console.log(`mikrosim fixed  : output "${mikrosimFixedOutput}", flagged ${mikrosimFixedFlag}`);
if ((mikrosimFixedOutput ?? "").trim() === "") {
  problems.push("fixing a flagged row's scheme through its own dropdown did not recompute it");
}
if (mikrosimFixedFlag !== 0) {
  problems.push("fixing a flagged row's scheme left it still marked as an error");
}

// Round trip: importing this app's own export back in recovers the same
// rows, computed immediately.
await tab
  .locator('[data-action="mikrosim-import-file"]')
  .setInputFiles({ name: "roundtrip.csv", mimeType: "text/csv", buffer: mikrosimCsv.buffer });
await tab.waitForTimeout(100);
const mikrosimRowsAfterRoundtrip = await tab.locator(".mikrosim-table tbody tr").count();
const roundtripSalary = await mikrosimCell(0, MI_SALARY_TD).locator("input").inputValue();
const roundtripOutput = await mikrosimCell(0, MI_FIRST_OUTPUT_TD).textContent();
console.log(
  `mikrosim roundtrip: ${mikrosimRowsAfterRoundtrip} rows, row 1 salary ${roundtripSalary}, output "${roundtripOutput}"`,
);
if (mikrosimRowsAfterRoundtrip !== 3) {
  problems.push(`re-importing this app's own CSV export left ${mikrosimRowsAfterRoundtrip} rows, expected 3`);
}
if (roundtripSalary !== "650000") {
  problems.push(`re-importing this app's own CSV export lost row 1's own salary (read back "${roundtripSalary}")`);
}
if ((roundtripOutput ?? "").trim() === "") {
  problems.push("re-importing this app's own CSV export left row 1 uncalculated");
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

// The old PGB grid needed a `<dialog>` to escape the sidebar's own width at
// a phone size, where even its own `.adv-grid-scroll` box was not enough to
// keep it from widening the page. The sparse summary table replacing it
// wraps in the same `.scroll` box every other wide table here already uses
// (Table 2, the comparison table): a long header can still make that one box
// scroll sideways on its own, same as Table 2's does, but the page itself
// must not follow it. Re-checked here at 375px (iPhone SE, the narrowest of
// the widths this app is hand-verified at: 375, 390, 1280) in its own
// short-lived context, rather than resizing `tab` and disturbing every
// assertion above that assumes the 1280px layout.
const narrowContext = await browser.newContext({ viewport: { width: 375, height: 700 } });
const narrowTab = await narrowContext.newPage();
await narrowTab.goto(pathToFileURL(page).href);
await narrowTab.waitForSelector("#app");
await narrowTab.locator('.mode-toggle .panel-btn[data-mode="advanced"]').click();
await narrowTab.locator('.adv-group[data-group="pgb"] summary').click();
await narrowTab.locator('[data-setting="pgbEntryType"]').selectOption("2"); // Sickness
const narrowYear = narrowTab.locator('[data-setting="pgbEntryYear"]');
await narrowYear.fill("2005");
await narrowYear.dispatchEvent("change");
const narrowAmount = narrowTab.locator('[data-setting="pgbEntryAmount"]');
await narrowAmount.fill("200000");
await narrowAmount.dispatchEvent("change");
await narrowTab.locator('[data-action="pgb-add-entry"]').click();
await narrowTab.waitForTimeout(120);
const narrowMetrics = await narrowTab.evaluate(() => ({
  scrollWidth: document.documentElement.scrollWidth,
  clientWidth: document.documentElement.clientWidth,
}));
console.log(
  `pgb page @375px : scrollWidth ${narrowMetrics.scrollWidth} <= clientWidth ${narrowMetrics.clientWidth}?`,
);
if (narrowMetrics.scrollWidth > narrowMetrics.clientWidth) {
  problems.push(
    `at a 375px phone width, the PGB summary table widened the whole page (scrollWidth ` +
      `${narrowMetrics.scrollWidth} > clientWidth ${narrowMetrics.clientWidth}) instead of scrolling in its own box`,
  );
}
await narrowContext.close();

// A first draft of the expanded PGB panel floated full-width at every size
// above 375px too, and a short viewport (or a sidebar scrolled far enough)
// could put the add-entry form's own fields directly under it -- caught by
// hand at 900px wide, not by the 375px or 1280px checks around it, since
// neither happens to put the form there. `.layout`'s own two-column
// breakpoint is 861px; checked here just above it, in its own short-lived
// context, by adding a second entry *while the panel is expanded* and
// confirming the click actually lands instead of hitting the floating
// panel on top of it.
const midContext = await browser.newContext({ viewport: { width: 900, height: 700 } });
const midTab = await midContext.newPage();
await midTab.goto(pathToFileURL(page).href);
await midTab.waitForSelector("#app");
await midTab.locator('.mode-toggle .panel-btn[data-mode="advanced"]').click();
await midTab.locator('.adv-group[data-group="pgb"] summary').click();
await midTab.locator('[data-setting="pgbEntryType"]').selectOption("2"); // Sickness
const midYear = midTab.locator('[data-setting="pgbEntryYear"]');
await midYear.fill("2005");
await midYear.dispatchEvent("change");
const midAmount = midTab.locator('[data-setting="pgbEntryAmount"]');
await midAmount.fill("200000");
await midAmount.dispatchEvent("change");
await midTab.locator('[data-action="pgb-add-entry"]').click();
await midTab.waitForTimeout(80);
await midTab.locator('[data-action="pgb-expand"]').click();
await midTab.waitForTimeout(80);
await midTab.locator('[data-setting="pgbEntryType"]').selectOption("3"); // Study
await midYear.fill("2010");
await midYear.dispatchEvent("change");
const midSemesters = midTab.locator('[data-setting="pgbEntryStudySemesters"]');
await midSemesters.fill("1");
await midSemesters.dispatchEvent("change");
// A plain `.click()` throws (rather than silently mis-clicking) if the
// floating panel intercepts the pointer here -- exactly the failure mode
// the right-hand docking above 860px exists to prevent.
await midTab.locator('[data-action="pgb-add-entry"]').click({ timeout: 5000 });
await midTab.waitForTimeout(80);
const midRows = await midTab.locator(".pgb-summary-table tbody tr").count();
console.log(`pgb expand @900px: rows=${midRows}, add-entry form reachable while expanded`);
if (midRows !== 2) {
  problems.push(`adding a second PGB entry while expanded at 900px left ${midRows} rows, expected 2`);
}
await midContext.close();

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
