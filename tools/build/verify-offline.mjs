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

// Figur 2's own after-tax/before-tax toggle, added on request -- the overlay
// line switches between two totals (`netto`/`brutto`) already on every row,
// so this checks the legend label, the active button and the line's own
// drawn points all change together, not just one of the three.
const fig2Figure = tab.locator("figure.figure").nth(1);
const basisButtonCount = await fig2Figure.locator(".panel-toggle button").count();
if (basisButtonCount !== 2) {
  problems.push(`Figur 2's after-tax/before-tax toggle has ${basisButtonCount} button(s), expected 2`);
}
const afterTaxClass = await fig2Figure.locator('button[data-basis="after-tax"]').getAttribute("class");
if (!afterTaxClass?.includes("active")) {
  problems.push(`Figur 2 does not default to "after-tax" (class was "${afterTaxClass}")`);
}
const linePointsAfterTax = await fig2Figure.locator("svg polyline.line").getAttribute("points");
await fig2Figure.locator('button[data-basis="before-tax"]').click();
await tab.waitForTimeout(80);
const fig2LegendBeforeTax = await fig2Figure.locator(".legend li").allTextContents();
console.log(`Figur 2 basis   : ${fig2LegendBeforeTax.join(" | ")}`);
if (!fig2LegendBeforeTax.some((text) => text.includes("Inkomst brutto") || text.includes("Income gross"))) {
  problems.push(
    `Figur 2's legend did not switch to "Inkomst brutto"/"Income gross" after picking before-tax (${fig2LegendBeforeTax.join(" | ")})`,
  );
}
const beforeTaxClass = await fig2Figure.locator('button[data-basis="before-tax"]').getAttribute("class");
if (!beforeTaxClass?.includes("active")) {
  problems.push("Figur 2's before-tax button did not become active after it was clicked");
}
const linePointsBeforeTax = await fig2Figure.locator("svg polyline.line").getAttribute("points");
if (linePointsBeforeTax === linePointsAfterTax) {
  problems.push("Figur 2's overlay line did not move after switching to before-tax (brutto vs netto)");
}
// Back to the default, so the screenshots below show the usual view.
await fig2Figure.locator('button[data-basis="after-tax"]').click();
await tab.waitForTimeout(80);

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

// The old combined "saving" group split into two, on request, and "Övrigt"/
// "Other" was removed entirely once flexPension and the two Slutlön settings
// moved out of it -- confirm the new titles exist and the old ones don't.
const advGroupTitles = await tab.locator(".advanced-box .adv-group summary").allTextContents();
console.log(`advanced titles : ${JSON.stringify(advGroupTitles)}`);
if (!advGroupTitles.includes("Privat sparande")) {
  problems.push('expected an advanced-mode group titled "Privat sparande"');
}
if (!advGroupTitles.includes("Tjänstepension")) {
  problems.push('expected an advanced-mode group titled "Tjänstepension"');
}
if (advGroupTitles.includes("Privat sparande och tjänstepension")) {
  problems.push('the old combined "Privat sparande och tjänstepension" group title is still shown');
}
if (advGroupTitles.includes("Övrigt")) {
  problems.push('the "Övrigt" group is still shown, expected it removed');
}

// Five groups renamed on request -- confirm the new titles exist and the old
// ones are gone.
const oldTitles = [
  "Partiellt uttag",
  "Underlag för bostadstillägg",
  "Försäkringstid",
  "Underlag för inkomstskatt",
  "Känt pensionskapital och avkastning",
];
for (const old of oldTitles) {
  if (advGroupTitles.includes(old)) {
    problems.push(`the old group title "${old}" is still shown, expected it renamed`);
  }
}

// The nine sections now render in plain Swedish alphabetical order by their
// own displayed title, spanning this file's own seven `.adv-group`s plus
// salary-path's "Lön" and pgb's "Pensionsgrundande belopp (PGB)" interleaved
// among them -- see advanced.ts's own header comment and main.ts's
// `advancedBox.append` call.
const expectedAdvOrder = [
  "Allmän pension",
  "Bostadstillägg",
  "Garantipension",
  "Inkomstskatt",
  "Kapital och avkastning",
  "Lön",
  "Pensionsgrundande belopp (PGB)",
  "Privat sparande",
  "Tjänstepension",
];
if (JSON.stringify(advGroupTitles) !== JSON.stringify(expectedAdvOrder)) {
  problems.push(
    `advanced-mode sections are ordered ${JSON.stringify(advGroupTitles)}, expected ${JSON.stringify(expectedAdvOrder)}`,
  );
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
// Moved into "Tjänstepension"/"Occupational pension" from "Övrigt", on request.
const flexPensionGroup = await flexPensionField
  .locator("xpath=ancestor::details[1]")
  .getAttribute("data-group");
if (flexPensionGroup !== "occupational") {
  problems.push(`flexPension sits in the "${flexPensionGroup}" group, expected "occupational"`);
}
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
// Moved into the renamed "Salary" section (formerly "Own salary path") from
// "Övrigt", on request -- alongside `finalSalaryYears`.
const pensionGapGroup = await pensionGapField.locator("xpath=ancestor::details[1]").getAttribute("data-group");
if (pensionGapGroup !== "salary-path") {
  problems.push(`pensionSameYearAsFinalSalary sits in the "${pensionGapGroup}" group, expected "salary-path"`);
}
const finalSalaryYearsField = tab.locator('[data-setting="finalSalaryYears"]');
const finalSalaryYearsGroup = await finalSalaryYearsField
  .locator("xpath=ancestor::details[1]")
  .getAttribute("data-group");
if (finalSalaryYearsGroup !== "salary-path") {
  problems.push(`finalSalaryYears sits in the "${finalSalaryYearsGroup}" group, expected "salary-path"`);
}
const salaryGroupTitle = await tab.locator('[data-group="salary-path"] summary').textContent();
if (salaryGroupTitle !== "Lön") {
  problems.push(`the salary-path section is titled "${salaryGroupTitle}", expected "Lön"`);
}

// Row 141 "Löneprofil": bends the derived wage curve by age rather than
// moving one figure directly, so this checks a lifetime-integrated figure
// (income pension) rather than Slutlön, whose narrow look-back window a
// particular profile's own shape could leave nearly unchanged by
// coincidence. The five options are the workbook's own dropdown, read off
// a LibreOffice-converted copy of the .xlsb during development -- option 0's
// exact wording differs (this port's own shorter text, not retyped from the
// sheet, matching every other label in advanced.ts), the four profiled ones
// don't.
const wageProfileField = await setting("wageProfile");
const wageProfileGroup = await wageProfileField.locator("xpath=ancestor::details[1]").getAttribute("data-group");
if (wageProfileGroup !== "salary-path") {
  problems.push(`wageProfile sits in the "${wageProfileGroup}" group, expected "salary-path"`);
}
const wageProfileOptions = await wageProfileField.locator("option").allTextContents();
const expectedWageProfiles = [
  "Rak – följer den allmänna löneutvecklingen",
  "Låg inkomst, kvinna",
  "Låg inkomst, man",
  "Normal inkomst, kvinna",
  "Normal inkomst, man",
];
if (JSON.stringify(wageProfileOptions) !== JSON.stringify(expectedWageProfiles)) {
  problems.push(
    `the Löneprofil dropdown shows ${JSON.stringify(wageProfileOptions)}, expected ${JSON.stringify(expectedWageProfiles)}`,
  );
}
const ipBeforeProfile = await shown("ip", "adjusted");
await wageProfileField.selectOption("2"); // Låg inkomst, man
await tab.waitForTimeout(80);
const ipAfterProfile = await shown("ip", "adjusted");
console.log(
  `wage profile    : inkomstpension (fasta priser) ${ipBeforeProfile} -> ${ipAfterProfile} at "Låg inkomst, man"`,
);
if (ipAfterProfile === ipBeforeProfile) {
  problems.push(
    `picking a wage profile did not change the income-pension figure (${ipBeforeProfile} -> ${ipAfterProfile})`,
  );
}

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
const wageProfileAfterReset = await wageProfileField.inputValue();
if (wageProfileAfterReset !== "0") {
  problems.push(`the reset button left Löneprofil at "${wageProfileAfterReset}", expected "0"`);
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
if (screenButtons !== 7) {
  problems.push(
    `screen toggle has ${screenButtons} buttons, expected 7 (Prognos / Jämför scenarier / Mikrosim / Hur maxa pensionen? / Hur länge räcker pensionen? / Hur mycket får jag? / Om Pensionslyft)`,
  );
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

// Birth year is the fifth override, added on request -- a much younger
// cohort has to move that scenario's own pension, with salary, retirement
// age and start-of-work age held fixed, without touching the baseline's.
const scenario1Born = tab.locator(".compare-card").nth(1).locator(".compare-card-controls input").first();
await scenario1Born.fill("1990");
await scenario1Born.dispatchEvent("change");
await tab.waitForTimeout(50);
const scenario1PensionAfterBorn = await compareCell("Scenario 1", "kpi-pension", "monthly");
const baselinePensionAfterBorn = await compareCell("Utgångsläge", "kpi-pension", "monthly");
console.log(`scenario birth year lowered: pension ${scenario1PensionBefore} -> ${scenario1PensionAfterBorn}`);
if (scenario1PensionAfterBorn === scenario1PensionBefore) {
  problems.push(
    `changing a scenario's own birth year did not move its pension in the table (still "${scenario1PensionAfterBorn}")`,
  );
}
if (baselinePensionAfterBorn !== baselinePension) {
  problems.push(
    `changing a scenario's birth year changed the baseline's own pension in the table ("${baselinePension}" -> ` +
      `"${baselinePensionAfterBorn}")`,
  );
}

// Raising the scenario's own salary well above the baseline's must move only
// that scenario's own columns in the table, not the baseline's.
const scenario1Salary = tab.locator(".compare-card").nth(1).locator(".compare-card-controls input").nth(1);
await scenario1Salary.fill("80000");
await scenario1Salary.dispatchEvent("change");
await tab.waitForTimeout(50);
const scenario1PensionAfter = await compareCell("Scenario 1", "kpi-pension", "monthly");
const baselinePensionAfter = await compareCell("Utgångsläge", "kpi-pension", "monthly");
console.log(`scenario salary raised: pension ${scenario1PensionAfterBorn} -> ${scenario1PensionAfter}`);
if (scenario1PensionAfter === scenario1PensionAfterBorn) {
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
const scenario1StartWork = tab.locator(".compare-card").nth(1).locator(".compare-card-controls input").nth(3);
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

// Inputs and Results are two independent tables sharing the same row order
// (Phase 12), each with its own `data-role` since both carry the plain
// `.mikrosim-table` class -- a bare `.mikrosim-table tbody tr` locator would
// match rows from both tables at once.
const mikrosimInputRow = (i) => tab.locator('[data-role="mikrosim-inputs-table"] tbody tr').nth(i);
const mikrosimInputCell = (i, col) => mikrosimInputRow(i).locator("td").nth(col);
const mikrosimResultRow = (i) => tab.locator('[data-role="mikrosim-results-table"] tbody tr').nth(i);
const mikrosimResultCell = (i, col) => mikrosimResultRow(i).locator("td").nth(col);
const mikrosimRowCount = (which) => tab.locator(`[data-role="mikrosim-${which}-table"] tbody tr`).count();
// Inputs table td indices: 0 row#, 1 status, 2-10 the nine inputs, 11 remove.
// Results table td indices: 0 row#, 1-12 the twelve outputs. See mikrosim.ts's
// own INPUT_COLUMNS/OUTPUT_COLUMNS order.
const MI_SALARY_TD = 5;
const MI_SCHEME_TD = 10;
const MI_FIRST_OUTPUT_TD = 1;

// Each table's own header row -- 12 columns (row#, status, nine inputs,
// remove) and 13 (row#, twelve outputs) -- locks in the two-table split
// itself, not just values inside it.
const mikrosimInputHeadCols = await tab
  .locator('[data-role="mikrosim-inputs-table"] thead tr')
  .first()
  .locator("th")
  .count();
const mikrosimResultHeadCols = await tab
  .locator('[data-role="mikrosim-results-table"] thead tr')
  .first()
  .locator("th")
  .count();
console.log(`mikrosim tables : inputs header ${mikrosimInputHeadCols} cols, results header ${mikrosimResultHeadCols} cols`);
if (mikrosimInputHeadCols !== 12) {
  problems.push(`the Mikrosim inputs table header has ${mikrosimInputHeadCols} columns, expected 12`);
}
if (mikrosimResultHeadCols !== 13) {
  problems.push(`the Mikrosim results table header has ${mikrosimResultHeadCols} columns, expected 13`);
}
const mikrosimLabels = await tab.locator(".mikrosim-table-label").allTextContents();
console.log(`mikrosim labels : ${JSON.stringify(mikrosimLabels)}`);
if (mikrosimLabels.length !== 2 || !mikrosimLabels[0] || !mikrosimLabels[1]) {
  problems.push(`expected two Mikrosim table labels (Inputs, Results), found ${JSON.stringify(mikrosimLabels)}`);
}

// Both tables' own column headers, not just their body cells, read 8px bold.
const mikrosimHeadStyle = await tab
  .locator('[data-role="mikrosim-inputs-table"] thead th')
  .nth(2)
  .evaluate((el) => {
    const cs = getComputedStyle(el);
    return { fontSize: cs.fontSize, fontWeight: cs.fontWeight };
  });
console.log(`mikrosim head style: ${mikrosimHeadStyle.fontSize} / ${mikrosimHeadStyle.fontWeight}`);
if (mikrosimHeadStyle.fontSize !== "8px" || Number(mikrosimHeadStyle.fontWeight) < 700) {
  problems.push(
    `the Mikrosim inputs table's own column headers read ${mikrosimHeadStyle.fontSize}/${mikrosimHeadStyle.fontWeight}, expected 8px/700+`,
  );
}

// There is no "Beräkna" button to look for -- confirm outright.
const mikrosimCalcButton = await tab.locator('[data-action="calculate-mikrosim"]').count();
if (mikrosimCalcButton !== 0) {
  problems.push(`found ${mikrosimCalcButton} "Beräkna"-style button(s) on the Mikrosim tab, expected none`);
}

// A single seed row, already computed with no click needed -- in both
// tables, which have to always agree on row count since they share one
// underlying row array.
const mikrosimRowsAtStart = await mikrosimRowCount("inputs");
const mikrosimResultRowsAtStart = await mikrosimRowCount("results");
console.log(`mikrosim rows   : ${mikrosimRowsAtStart} inputs, ${mikrosimResultRowsAtStart} results`);
if (mikrosimRowsAtStart !== 1 || mikrosimResultRowsAtStart !== 1) {
  problems.push(
    `the Mikrosim tab starts with ${mikrosimRowsAtStart} input row(s)/${mikrosimResultRowsAtStart} result row(s), expected 1/1`,
  );
}
const mikrosimOutputsAtStart = await mikrosimResultCell(0, MI_FIRST_OUTPUT_TD).textContent();
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
const mikrosimRowsAfterAdd = await mikrosimRowCount("inputs");
const mikrosimResultRowsAfterAdd = await mikrosimRowCount("results");
console.log(`mikrosim rows   : ${mikrosimRowsAfterAdd} inputs, ${mikrosimResultRowsAfterAdd} results, after adding three`);
if (mikrosimRowsAfterAdd !== 4 || mikrosimResultRowsAfterAdd !== 4) {
  problems.push(
    `adding three Mikrosim rows left ${mikrosimRowsAfterAdd} input row(s)/${mikrosimResultRowsAfterAdd} result row(s), expected 4/4`,
  );
}
const mikrosimNewRowOutput = await mikrosimResultCell(3, MI_FIRST_OUTPUT_TD).textContent();
if ((mikrosimNewRowOutput ?? "").trim() === "") {
  problems.push("a freshly added Mikrosim row shows no output -- expected it computed immediately");
}

// Row 1 gets a distinct salary and scheme; rows 2-4 are left at their shared
// defaults. Editing alone -- no click anywhere -- has to update row 1's own
// output from its own inputs, while rows 2-4 keep agreeing with each other.
await mikrosimInputCell(0, MI_SALARY_TD).locator("input").fill("600000");
await mikrosimInputCell(0, MI_SALARY_TD).locator("input").dispatchEvent("change");
await mikrosimInputCell(0, MI_SCHEME_TD).locator("select").selectOption("3");
await tab.waitForTimeout(100);

const mikrosimRow1Salary = await mikrosimResultCell(0, MI_FIRST_OUTPUT_TD).textContent();
const mikrosimRow2Salary = await mikrosimResultCell(1, MI_FIRST_OUTPUT_TD).textContent();
const mikrosimRow3Salary = await mikrosimResultCell(2, MI_FIRST_OUTPUT_TD).textContent();
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
// row's own already-computed figures. This edit alone (no add/remove/import)
// also has to move the chart's own bars, not just the table -- Phase 12's
// real bug was `editAndRecompute()` never touching the chart at all, which a
// bar-*count* check alone (kept further below) would not have caught, since
// the count stays nonzero whether or not the bars' own values refresh.
const mikrosimChart = tab.locator(".mikrosim-chart");
const mikrosimChartBeforePlainEdit = await mikrosimChart.locator("svg").innerHTML();

await mikrosimInputCell(0, MI_SALARY_TD).locator("input").fill("650000");
await mikrosimInputCell(0, MI_SALARY_TD).locator("input").dispatchEvent("change");
await tab.waitForTimeout(100);
const mikrosimRow1AfterEdit = await mikrosimResultCell(0, MI_FIRST_OUTPUT_TD).textContent();
const mikrosimRow2AfterEdit = await mikrosimResultCell(1, MI_FIRST_OUTPUT_TD).textContent();
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

const mikrosimChartAfterPlainEdit = await mikrosimChart.locator("svg").innerHTML();
console.log(
  `mikrosim chart  : markup ${mikrosimChartAfterPlainEdit === mikrosimChartBeforePlainEdit ? "unchanged" : "changed"} after a plain edit`,
);
if (mikrosimChartAfterPlainEdit === mikrosimChartBeforePlainEdit) {
  problems.push(
    "editing a Mikrosim cell alone (no add/remove/import) left the chart's own markup unchanged -- expected its bars to move with the new value",
  );
}

// The chart below the table draws one bar group per row, live -- the same
// edit that just changed row 1's own table cells has to change its own bar.
const mikrosimBarsBeforeRemove = await mikrosimChart.locator("svg .bar").count();
const mikrosimLegendItems = await mikrosimChart.locator(".legend li").count();
console.log(`mikrosim chart  : ${mikrosimBarsBeforeRemove} bar segments, ${mikrosimLegendItems} legend entries`);
if (mikrosimBarsBeforeRemove === 0) {
  problems.push("the Mikrosim chart draws no bars at all");
}
// Up to 7 stacked bands plus the Disponibel inkomst overlay line.
if (mikrosimLegendItems === 0 || mikrosimLegendItems > 8) {
  problems.push(`the Mikrosim chart legend has ${mikrosimLegendItems} entries, expected 1-8`);
}

// Disponibel inkomst is drawn as a black overlay line, not another stacked
// band, with its own legend entry -- one point per row, so its own line
// count tracks the row count the same way the bars already do.
const mikrosimDisposableLines = await mikrosimChart.locator("svg .mikrosim-disposable-line").count();
const mikrosimDisposableLegend = await mikrosimChart
  .locator(".legend li", { hasText: "Disponibel inkomst" })
  .count();
console.log(`mikrosim chart  : ${mikrosimDisposableLines} disposable-income line(s), legend entry present: ${mikrosimDisposableLegend > 0}`);
if (mikrosimDisposableLines !== 1) {
  problems.push(`the Mikrosim chart draws ${mikrosimDisposableLines} disposable-income line(s), expected 1`);
}
if (mikrosimDisposableLegend === 0) {
  problems.push('the Mikrosim chart legend is missing a "Disponibel inkomst" entry for its overlay line');
}

// A row's own Bruttopension total is labelled above its bar -- annual by
// default, matching the Results table's own "Brutto-pension" column exactly
// (both read the same `TotalGross`, and the sum of the seven stacked bands
// equals it by construction) -- with 4 rows, at or under the 10-row cap.
const mikrosimBrutto = await mikrosimResultCell(0, 2).textContent();
const mikrosimBarValueLabels = await mikrosimChart.locator(".mikrosim-bar-value").allTextContents();
console.log(`mikrosim brutto : table "${mikrosimBrutto}", chart labels ${JSON.stringify(mikrosimBarValueLabels)}`);
if (mikrosimBarValueLabels.length !== 4) {
  problems.push(
    `the Mikrosim chart shows ${mikrosimBarValueLabels.length} Bruttopension label(s), expected 4 (one per row, at or under the 10-row cap)`,
  );
}
if (mikrosimBarValueLabels[0] !== mikrosimBrutto?.trim()) {
  problems.push(
    `the chart's own row-1 Bruttopension label ("${mikrosimBarValueLabels[0]}") doesn't match the Results table's own value ("${mikrosimBrutto}")`,
  );
}

// The chart's own Årsvis/Månadsvis toggle is local to the chart -- it moves
// the bars and this label, never the Results table above, which always
// reads annual.
const mikrosimChartToggle = tab.locator(".mikrosim-panel .panel-toggle").last();
await mikrosimChartToggle.locator("button", { hasText: "Månadsvis" }).click();
await tab.waitForTimeout(100);
const mikrosimBarValueLabelsMonthly = await mikrosimChart.locator(".mikrosim-bar-value").allTextContents();
const mikrosimBruttoAfterToggle = await mikrosimResultCell(0, 2).textContent();
console.log(`mikrosim brutto (månadsvis): chart "${mikrosimBarValueLabelsMonthly[0]}", table "${mikrosimBruttoAfterToggle}"`);
if (mikrosimBarValueLabelsMonthly[0] === mikrosimBarValueLabels[0]) {
  problems.push("switching the Mikrosim chart to Månadsvis did not change its own Bruttopension label");
}
if (mikrosimBruttoAfterToggle !== mikrosimBrutto) {
  problems.push(
    "switching the Mikrosim chart's own Årsvis/Månadsvis toggle changed the Results table too, expected it to stay annual",
  );
}
await mikrosimChartToggle.locator("button", { hasText: "Årsvis" }).click();
await tab.waitForTimeout(100);

// Removing a row drops the count, keeps the remaining rows' own values, and
// the chart's own bar count drops with it.
await mikrosimInputRow(1).locator('[data-action="mikrosim-remove-row"]').click();
await tab.waitForTimeout(50);
const mikrosimRowsAfterRemove = await mikrosimRowCount("inputs");
const mikrosimResultRowsAfterRemove = await mikrosimRowCount("results");
const mikrosimBarsAfterRemove = await mikrosimChart.locator("svg .bar").count();
console.log(
  `mikrosim rows   : ${mikrosimRowsAfterRemove} inputs, ${mikrosimResultRowsAfterRemove} results, after removing one, chart bars: ${mikrosimBarsAfterRemove}`,
);
if (mikrosimRowsAfterRemove !== 3 || mikrosimResultRowsAfterRemove !== 3) {
  problems.push(
    `removing a Mikrosim row left ${mikrosimRowsAfterRemove} input row(s)/${mikrosimResultRowsAfterRemove} result row(s), expected 3/3`,
  );
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
  "Real fondavkastning;Privat pensionsförsäkring;Välj tjänstepension;Slutlön;" +
  "Brutto-pension;Inkomstpension;Tilläggspension;Premiepension;Garanti-pension;P_tillägg;" +
  "Tjänstepension;Privat pensionsförsäkring;Efter skatt;Bostadstillägg + ÄFS;Disponibel inkomst";
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
  "Årlig inflation;Real tillväxt;Real fondavkastning;Privat pensionsförsäkring\n" +
  "2;1980;500000;22;65;0;0;0,017;0\n" +
  "4;1965;300000;19;66;0;0;0,017;0\n";
await tab
  .locator('[data-action="mikrosim-import-file"]')
  .setInputFiles({ name: "scrambled.csv", mimeType: "text/csv", buffer: Buffer.from(scrambledCsv, "utf8") });
await tab.waitForTimeout(100);
const mikrosimRowsAfterImport = await mikrosimRowCount("inputs");
const mikrosimResultRowsAfterImport = await mikrosimRowCount("results");
const importedBorn = await mikrosimInputCell(0, 2).locator("input").inputValue();
const mikrosimImportedOutput = await mikrosimResultCell(0, MI_FIRST_OUTPUT_TD).textContent();
console.log(
  `mikrosim import : ${mikrosimRowsAfterImport} inputs, ${mikrosimResultRowsAfterImport} results, row 1 born ${importedBorn}, output "${mikrosimImportedOutput}"`,
);
if (mikrosimRowsAfterImport !== 2 || mikrosimResultRowsAfterImport !== 2) {
  problems.push(
    `importing a 2-row CSV left ${mikrosimRowsAfterImport} input row(s)/${mikrosimResultRowsAfterImport} result row(s), expected 2/2 (replace, not append)`,
  );
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
const mikrosimRowsAfterBadImport = await mikrosimRowCount("inputs");
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
      "Real fondavkastning;Privat pensionsförsäkring;Välj tjänstepension\n" +
      "1970;20;66;480000;0;0;0,017;0;3\n" +
      "1970;20;66;480000;0;0;0,017;0;99\n",
    "utf8",
  ),
});
await tab.waitForTimeout(100);
const mikrosimGoodRowOutput = await mikrosimResultCell(0, MI_FIRST_OUTPUT_TD).textContent();
const mikrosimBadRowOutput = await mikrosimResultCell(1, MI_FIRST_OUTPUT_TD).textContent();
const mikrosimBadRowFlagged = await mikrosimInputRow(1).locator(".mikrosim-error").count();
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

await mikrosimInputCell(1, MI_SCHEME_TD).locator("select").selectOption("2");
await tab.waitForTimeout(100);
const mikrosimFixedOutput = await mikrosimResultCell(1, MI_FIRST_OUTPUT_TD).textContent();
const mikrosimFixedFlag = await mikrosimInputRow(1).locator(".mikrosim-error").count();
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
const mikrosimRowsAfterRoundtrip = await mikrosimRowCount("inputs");
const mikrosimResultRowsAfterRoundtrip = await mikrosimRowCount("results");
const roundtripSalary = await mikrosimInputCell(0, MI_SALARY_TD).locator("input").inputValue();
const roundtripOutput = await mikrosimResultCell(0, MI_FIRST_OUTPUT_TD).textContent();
console.log(
  `mikrosim roundtrip: ${mikrosimRowsAfterRoundtrip} inputs, ${mikrosimResultRowsAfterRoundtrip} results, row 1 salary ${roundtripSalary}, output "${roundtripOutput}"`,
);
if (mikrosimRowsAfterRoundtrip !== 3 || mikrosimResultRowsAfterRoundtrip !== 3) {
  problems.push(
    `re-importing this app's own CSV export left ${mikrosimRowsAfterRoundtrip} input row(s)/${mikrosimResultRowsAfterRoundtrip} result row(s), expected 3/3`,
  );
}
if (roundtripSalary !== "650000") {
  problems.push(`re-importing this app's own CSV export lost row 1's own salary (read back "${roundtripSalary}")`);
}
if ((roundtripOutput ?? "").trim() === "") {
  problems.push("re-importing this app's own CSV export left row 1 uncalculated");
}

// "Hämta från Prognos" appends one row read from Prognos's own current
// baseline -- unchanged throughout this script (born 1959, salary 462000) --
// computed immediately, no click needed beyond the import itself.
await tab.locator('[data-action="mikrosim-import-forecast"]').click();
await tab.waitForTimeout(100);
const mikrosimRowsAfterForecastImport = await mikrosimRowCount("inputs");
const forecastRowIndex = mikrosimRowsAfterForecastImport - 1;
const forecastRowBorn = await mikrosimInputCell(forecastRowIndex, 2).locator("input").inputValue();
const forecastRowSalary = await mikrosimInputCell(forecastRowIndex, MI_SALARY_TD).locator("input").inputValue();
const forecastRowOutput = await mikrosimResultCell(forecastRowIndex, MI_FIRST_OUTPUT_TD).textContent();
console.log(
  `mikrosim from Prognos: ${mikrosimRowsAfterForecastImport} rows, born ${forecastRowBorn}, salary ${forecastRowSalary}, output "${forecastRowOutput}"`,
);
if (mikrosimRowsAfterForecastImport !== 4) {
  problems.push(`"Hämta från Prognos" left ${mikrosimRowsAfterForecastImport} rows, expected 4 (3 + 1)`);
}
if (forecastRowBorn !== "1959" || forecastRowSalary !== "462000") {
  problems.push(
    `the row imported from Prognos reads born ${forecastRowBorn}/salary ${forecastRowSalary}, expected 1959/462000`,
  );
}
if ((forecastRowOutput ?? "").trim() === "") {
  problems.push("the row imported from Prognos shows no output -- expected it computed immediately");
}

// "Hämta från Jämför scenarier" appends one row per current scenario --
// baseline plus whichever scenario the earlier cap/floor cycle left behind
// (not necessarily the one whose birth year was edited above, since removal
// there goes in DOM order) -- given a fresh, distinct birth year here so the
// check proves each row reflects that scenario's own resolved fields, not
// just the shared baseline repeated.
await tab.locator('.screen-toggle .panel-btn[data-screen="compare"]').click();
await tab.waitForTimeout(50);
const compareSurvivorBorn = tab
  .locator(".compare-card:not([data-scenario='baseline']) .compare-card-controls input")
  .first();
await compareSurvivorBorn.fill("1975");
await compareSurvivorBorn.dispatchEvent("change");
await tab.waitForTimeout(50);
await tab.locator('.screen-toggle .panel-btn[data-screen="mikrosim"]').click();
await tab.waitForTimeout(50);
await tab.locator('[data-action="mikrosim-import-compare"]').click();
await tab.waitForTimeout(100);
const mikrosimRowsAfterCompareImport = await mikrosimRowCount("inputs");
const compareBaselineRowBorn = await mikrosimInputCell(4, 2).locator("input").inputValue();
const compareScenarioRowBorn = await mikrosimInputCell(5, 2).locator("input").inputValue();
console.log(
  `mikrosim from Jämför scenarier: ${mikrosimRowsAfterCompareImport} rows, borns ${compareBaselineRowBorn}/${compareScenarioRowBorn}`,
);
if (mikrosimRowsAfterCompareImport !== 6) {
  problems.push(
    `"Hämta från Jämför scenarier" left ${mikrosimRowsAfterCompareImport} rows, expected 6 (4 + 2, baseline and one scenario)`,
  );
}
if (compareBaselineRowBorn !== "1959") {
  problems.push(
    `the baseline row imported from Jämför scenarier reads born ${compareBaselineRowBorn}, expected 1959`,
  );
}
if (compareScenarioRowBorn !== "1975") {
  problems.push(
    `the scenario row imported from Jämför scenarier reads born ${compareScenarioRowBorn}, expected 1975 (its own edited birth year)`,
  );
}

// Past 10 rows, no Bruttopension labels are drawn at all -- not some of
// them -- since they would overlap and clutter more than they inform.
for (let i = 0; i < 8; i += 1) {
  await tab.locator('[data-action="add-mikrosim-row"]').click();
}
await tab.waitForTimeout(100);
const mikrosimRowsOver10 = await mikrosimRowCount("inputs");
const mikrosimBarValueLabelsOver10 = await mikrosimChart.locator(".mikrosim-bar-value").count();
console.log(`mikrosim >10 rows: ${mikrosimRowsOver10} rows, ${mikrosimBarValueLabelsOver10} Bruttopension label(s)`);
if (mikrosimRowsOver10 <= 10) {
  problems.push(`expected more than 10 Mikrosim rows for this check, found ${mikrosimRowsOver10}`);
}
if (mikrosimBarValueLabelsOver10 !== 0) {
  problems.push(`the Mikrosim chart shows ${mikrosimBarValueLabelsOver10} Bruttopension label(s) with over 10 rows, expected 0`);
}
// Unlike the Bruttopension labels, the disposable-income line has no 10-row
// cap -- it is one continuous polyline covering every row, not a label per
// row that would start overlapping.
const mikrosimDisposableLinesOver10 = await mikrosimChart.locator("svg .mikrosim-disposable-line").count();
if (mikrosimDisposableLinesOver10 !== 1) {
  problems.push(
    `the Mikrosim chart draws ${mikrosimDisposableLinesOver10} disposable-income line(s) with over 10 rows, expected 1`,
  );
}

// Back to the single-scenario view for the screenshots, and to leave the
// page as found.
await tab.locator('.screen-toggle .panel-btn[data-screen="single"]').click();
await tab.waitForTimeout(50);

// ---- Hur maxa pensionen? ---------------------------------------------------
//
// A fourth top-level view: seven retirement ages around a reference age, run
// as one typfall each, summed over a lifespan and compared. The point of each
// check, as above, is that a control reaches the runs it should and the page
// says what the engine said. The numbers pinned here are the ones a static
// report on this very typfall (born 1980, work from 25, 38 500 kr/month, the
// workbook's own assumptions, no agreement) printed first -- the tab has to
// reproduce that report to the krona -- and `pensionAge.test.ts` pins the
// same ones against the engine without a browser.

const paRiktalderTable = JSON.parse(readFileSync(join(repo, "packages/data/retirement-ages.json"), "utf8"));
const paRiktalder = (born) => paRiktalderTable.riktalder[born - paRiktalderTable.firstCohort];

const PA_BEST_DEFAULT = [65, 65, 65, 66, 68, 68, 68, 71, 71, 71, 71, 71, 71, 71, 71, 71];
const PA_BEST_ITP1 = [65, 65, 65, 66, 66, 68, 68, 68, 71, 71, 71, 71, 71, 71, 71, 71];

await tab.locator('.screen-toggle .panel-btn[data-screen="retirement"]').click();
await tab.waitForTimeout(100);

const pa = tab.locator('[data-role="pension-age-panel"]');
const paSetting = (name) => pa.locator(`[data-setting="${name}"]`);
const paType = async (name, value) => {
  const field = paSetting(name);
  await field.fill(String(value));
  await field.dispatchEvent("change");
  // A field that was typed into fires one more change when it loses focus, and that redraws the
  // results. Let it happen here, so that each step starts from a page that has settled; a press
  // made while a field is still being typed in is what "The first press" below checks.
  await field.evaluate((el) => el.blur());
  await tab.waitForTimeout(60);
};
const paBest = () =>
  pa.locator('[data-role="pa-ribbon"] li').evaluateAll((items) => items.map((li) => Number(li.dataset.best)));
const paTies = () =>
  pa.locator('[data-role="pa-ribbon"] li').evaluateAll((items) => items.map((li) => li.dataset.tie));
const paCell = async (lifespan, par) =>
  Number(await pa.locator(`[data-role="pa-matrix"] tr[data-lifespan="${lifespan}"] td[data-par="${par}"]`).getAttribute("data-kr"));
const paLines = (figure) => pa.locator(`[data-role="${figure}"] svg polyline.pa-line`).evaluateAll((lines) => lines.map((l) => Number(l.dataset.par)));
const sameList = (a, b) => a.length === b.length && a.every((value, i) => value === b[i]);

if ((await pa.count()) !== 1) problems.push("the Hur maxa pensionen? tab did not open");

// The inputs start as the report's typfall, centred on the cohort's riktålder.
const paStart = {
  born: await paSetting("paBorn").inputValue(),
  start: await paSetting("paStartWork").inputValue(),
  salary: await paSetting("paSalary").inputValue(),
  centre: await paSetting("paCentre").inputValue(),
  centreDisabled: await paSetting("paCentre").isDisabled(),
  riktalder: await paSetting("paRiktalder").isChecked(),
};
console.log(`pension age     : inputs ${JSON.stringify(paStart)}`);
if (paStart.born !== "1980" || paStart.start !== "25" || paStart.salary !== "38500" || paStart.centre !== "68") {
  problems.push(`the Hur maxa pensionen? inputs start at ${JSON.stringify(paStart)}, expected 1980/25/38500 around 68`);
}
if (!paStart.centreDisabled || !paStart.riktalder) {
  problems.push("the reference age is not locked to the riktålder as loaded");
}

// The short answer: sixteen lifespans, the report's winners, its near-ties.
const paBestDefault = await paBest();
console.log(`pension age     : best per lifespan ${paBestDefault.join(",")}`);
if (!sameList(paBestDefault, PA_BEST_DEFAULT)) {
  problems.push(`the short answer picks ${paBestDefault.join(",")} for lifespans 80-95, expected ${PA_BEST_DEFAULT.join(",")}`);
}
// Each box also carries the average monthly amount and the lifetime total of its best start: 80 -> 65 is the
// matrix's 2 972 916 over 180 months, and the lifetime total is the same sum in millions.
const paBoxFigures = await pa
  .locator('[data-role="pa-ribbon"] li[data-lifespan="80"] .pa-ribbon-v')
  .evaluateAll((cells) => cells.map((cell) => (cell.textContent ?? "").replace(/\s/g, "")));
console.log(`pension age     : box 80 -> monthly ${paBoxFigures[0]}, total ${paBoxFigures[1]} Mkr`);
if (!sameList(paBoxFigures, ["16516", "3,0"])) {
  problems.push(`the answer box for lifespan 80 shows ${paBoxFigures.join(" / ")}, expected 16516 kr/month and 3,0 Mkr`);
}
const paTiesDefault = (await paTies()).join("");
if (paTiesDefault !== "0111111111000000") {
  problems.push(`the short answer flags near-ties as ${paTiesDefault}, expected 0111111111000000 (81-89)`);
}

// The matrix: the report's totals to the krona, the winner marked.
const paMatrixRows = await pa.locator('[data-role="pa-matrix"] tbody tr').count();
const paTotal80 = await paCell(80, 65);
const paTotal83 = await paCell(83, 66);
console.log(`pension age     : matrix ${paMatrixRows} rows, 80/65=${paTotal80}, 83/66=${paTotal83}`);
if (paMatrixRows !== 16) problems.push(`the matrix has ${paMatrixRows} rows, expected 16 (lifespans 80-95)`);
if (paTotal80 !== 2972916) problems.push(`the matrix says ${paTotal80} for living to 80 with a start at 65, the report says 2972916`);
if (paTotal83 !== 3552168) problems.push(`the matrix says ${paTotal83} for living to 83 with a start at 66, the report says 3552168`);
const paBestCells = await pa.locator('[data-role="pa-matrix"] td.pa-best').count();
if (paBestCells !== 16) problems.push(`the matrix marks ${paBestCells} winning cells, expected one per row (16)`);

// The three charts: one line per alternative, the reference drawn heavier.
const paMonthlyLines = await paLines("pa-fig-monthly");
const paTotalLines = await paLines("pa-fig-total");
const paDeltaLines = await paLines("pa-fig-delta");
console.log(`pension age     : lines monthly ${paMonthlyLines.length}, total ${paTotalLines.length}, delta ${paDeltaLines.length}`);
for (const [name, lines] of [["monthly", paMonthlyLines], ["total", paTotalLines], ["difference", paDeltaLines]]) {
  if (!sameList([...lines].sort((a, b) => a - b), [65, 66, 67, 68, 69, 70, 71])) {
    problems.push(`the ${name} chart draws lines for ${lines.join(",")}, expected 65-71`);
  }
}
const paHeavy = await pa.locator("svg polyline.pa-line-z").evaluateAll((lines) => [...new Set(lines.map((l) => Number(l.dataset.par)))]);
if (!sameList(paHeavy, [68])) problems.push(`the heavy reference line is ${paHeavy.join(",")}, expected 68 only`);

// The hover readout names every line that has a value at the pointer's age.
const paTotalFigure = pa.locator('[data-role="pa-fig-total"]');
await paTotalFigure.scrollIntoViewIfNeeded();
const paBox = await paTotalFigure.locator("svg").boundingBox();
if (paBox) {
  const k = paBox.width / 760;
  // 80 of 65-95, in the chart's own geometry (68 left, 722 right of 760).
  await tab.mouse.move(paBox.x + 300 * k, paBox.y + 150 * k);
  await tab.mouse.move(paBox.x + (68 + 0.5 * (722 - 68)) * k, paBox.y + 150 * k);
  await tab.waitForTimeout(100);
}
const paTooltipRows = await paTotalFigure.locator(".tooltip .tooltip-row").count();
const paTooltipHead = (await paTotalFigure.locator(".tooltip strong").textContent()) ?? "";
console.log(`pension age     : hover "${paTooltipHead}", ${paTooltipRows} rows`);
if (paTooltipRows !== 7 || !/80/.test(paTooltipHead)) {
  problems.push(`the total chart's hover shows "${paTooltipHead}" with ${paTooltipRows} rows at age 80, expected 7 rows`);
}
// The same readout from the keyboard.
await paTotalFigure.locator("svg").focus();
await tab.keyboard.press("ArrowRight");
const paKeyHead = (await paTotalFigure.locator(".tooltip strong").textContent()) ?? "";
if (!/81/.test(paKeyHead)) problems.push(`the arrow key moved the readout to "${paKeyHead}", expected age 81`);
await tab.keyboard.press("Escape");

// A collective agreement adds exactly its occupational pension, and can move the winners.
await paSetting("paScheme").selectOption("2"); // ITP-1
await tab.waitForTimeout(80);
const paItpBest = await paBest();
const paItpTotal = await paCell(80, 65);
console.log(`pension age     : ITP-1 best ${paItpBest.join(",")}, 80/65=${paItpTotal}`);
if (paItpTotal !== 4238496) problems.push(`with ITP-1 the total for 80/65 is ${paItpTotal}, expected 4238496`);
if (!sameList(paItpBest, PA_BEST_ITP1)) {
  problems.push(`with ITP-1 the short answer picks ${paItpBest.join(",")}, expected ${PA_BEST_ITP1.join(",")}`);
}
const paTimeHeads = await pa.locator('[data-role="pa-time-table"] thead th').allTextContents();
if (!paTimeHeads.some((text) => /tjänstepension/i.test(text))) {
  problems.push(`the per-age table has no occupational-pension column with ITP-1 chosen (${paTimeHeads.join(" | ")})`);
}
if (await paSetting("paIncludeOccupational").isDisabled()) {
  problems.push("the include-occupational box is still disabled with an agreement chosen");
}
// ...and switching it off is the same as choosing no agreement.
await paSetting("paIncludeOccupational").uncheck();
await tab.waitForTimeout(80);
const paOffTotal = await paCell(80, 65);
if (paOffTotal !== 2972916) problems.push(`with occupational pension switched off the total for 80/65 is ${paOffTotal}, expected 2972916`);
await paSetting("paIncludeOccupational").check();
await paSetting("paScheme").selectOption("1");
await tab.waitForTimeout(80);

// The economic assumptions: the second preset moves the break-evens years earlier.
await pa.locator('button[data-preset="growth"]').click();
await tab.waitForTimeout(80);
const paGrowthBest = await paBest();
const paGrowthActive = await pa.locator('button[data-preset="growth"]').evaluate((b) => b.classList.contains("active"));
const paGrowthFields = [await paSetting("paInflation").inputValue(), await paSetting("paGrowth").inputValue()];
console.log(`pension age     : growth best ${paGrowthBest.slice(0, 4).join(",")}..., fields ${paGrowthFields.join("/")}`);
if (!sameList(paGrowthBest.slice(0, 4), [67, 68, 70, 71]) || !paGrowthBest.slice(3).every((par) => par === 71)) {
  problems.push(`with 2 %/1.8 % assumptions the short answer picks ${paGrowthBest.join(",")}, expected 67,68,70 then 71 from 83`);
}
if (!paGrowthActive || paGrowthFields[0] !== "2" || paGrowthFields[1] !== "1,8") {
  problems.push(`the growth preset reads active=${paGrowthActive}, fields ${paGrowthFields.join("/")}, expected active 2/1,8`);
}
await pa.locator('button[data-preset="standard"]').click();
await tab.waitForTimeout(80);
if (!sameList(await paBest(), PA_BEST_DEFAULT)) problems.push("the standard preset did not restore the default short answer");

// After tax sums the model's net income.
await pa.locator('button[data-basis="net"]').click();
await tab.waitForTimeout(80);
const paNetTotal = await paCell(80, 65);
const paNetBest = await paBest();
console.log(`pension age     : after tax 80/65=${paNetTotal}, best ${paNetBest[0]}..${paNetBest[paNetBest.length - 1]}`);
if (paNetTotal !== 2449482) problems.push(`after tax the total for 80/65 is ${paNetTotal}, expected 2449482`);
if (paNetBest[0] !== 65 || paNetBest[paNetBest.length - 1] !== 71) {
  problems.push(`after tax the winners run ${paNetBest[0]}..${paNetBest[paNetBest.length - 1]}, expected 65..71`);
}
await pa.locator('button[data-basis="gross"]').click();
await tab.waitForTimeout(80);

// The reference age: unlock the riktålder and move it; ages outside the cohort's range are clamped.
await paSetting("paRiktalder").uncheck();
await paType("paCentre", 72);
const paShifted = await paLines("pa-fig-total");
if (!sameList([...paShifted].sort((a, b) => a - b), [69, 70, 71, 72, 73, 74, 75])) {
  problems.push(`a reference age of 72 draws lines for ${paShifted.join(",")}, expected 69-75`);
}
await paType("paCentre", 62);
const paClamped = await paLines("pa-fig-total");
const paCentreShown = await paSetting("paCentre").inputValue();
console.log(`pension age     : reference 62 -> shows ${paCentreShown}, lines ${paClamped.join(",")}`);
if (paCentreShown !== "65" || !sameList([...paClamped].sort((a, b) => a - b), [65, 66, 67, 68])) {
  problems.push(`a reference age of 62 (cohort's lowest is 65) shows ${paCentreShown} with lines ${paClamped.join(",")}, expected 65 and 65-68`);
}
await paSetting("paRiktalder").check();
await tab.waitForTimeout(80);
if ((await paSetting("paCentre").inputValue()) !== "68") problems.push("ticking the riktålder box did not put the reference age back on 68");

// The riktålder follows the birth year while its box is ticked.
await paType("paBorn", 1990);
const paCentre1990 = await paSetting("paCentre").inputValue();
if (paCentre1990 !== String(paRiktalder(1990))) {
  problems.push(`born 1990 puts the reference age at ${paCentre1990}, the cohort table says ${paRiktalder(1990)}`);
}
await paType("paBorn", 1980);

// A final salary as the input: the monthly salary behind it is worked out and shown.
await paSetting("paSalaryAs").selectOption("1");
await tab.waitForTimeout(80);
const paFinalShown = await paSetting("paSalary").inputValue();
await paType("paSalary", 40000);
const paMonthlyHint = (await pa.locator('label:has([data-setting="paSalary"]) .field-hint').textContent()) ?? "";
console.log(`pension age     : final salary shown ${paFinalShown}, hint "${paMonthlyHint}"`);
if (paFinalShown !== "39872") problems.push(`switching to final salary shows ${paFinalShown}, expected 39872 (the model's own final salary)`);
if (!/38\D?623/.test(paMonthlyHint)) {
  problems.push(`typing a final salary of 40000 gives the hint "${paMonthlyHint}", expected the monthly salary 38 623`);
}
const paFinalRaised = await paCell(80, 65);
if (!(paFinalRaised > 2972916)) problems.push("a higher final salary did not raise the totals");
await paSetting("paSalaryAs").selectOption("0");
await tab.waitForTimeout(80);
if ((await paSetting("paSalary").inputValue()) !== "38623") {
  problems.push(`switching back to monthly salary shows ${await paSetting("paSalary").inputValue()}, expected the derived 38623`);
}
await paType("paSalary", 38500);

// The lifespan range, linked so it never inverts; a single lifespan has no difference chart.
await paType("paLifespanFrom", 85);
await paType("paLifespanTo", 90);
const paItems = await pa.locator('[data-role="pa-ribbon"] li').count();
if (paItems !== 6) problems.push(`lifespans 85-90 give ${paItems} answer boxes, expected 6`);
await paType("paLifespanFrom", 92);
const paToLinked = await paSetting("paLifespanTo").inputValue();
const paDeltaWhenOne = await pa.locator('[data-role="pa-fig-delta"]').count();
if (paToLinked !== "92" || paDeltaWhenOne !== 0) {
  problems.push(`raising "from" to 92 left "to" at ${paToLinked} and ${paDeltaWhenOne} difference chart(s), expected 92 and none`);
}

// Reset puts everything back; "Hämta från Prognos" copies the form's person once.
await pa.locator('[data-action="pa-reset"]').click();
await tab.waitForTimeout(80);
if (!sameList(await paBest(), PA_BEST_DEFAULT) || (await paSetting("paLifespanFrom").inputValue()) !== "80") {
  problems.push("Återställ did not restore the default typfall");
}
await pa.locator('[data-action="pa-import-forecast"]').click();
await tab.waitForTimeout(80);
const paImported = [
  await paSetting("paBorn").inputValue(),
  await paSetting("paStartWork").inputValue(),
  await paSetting("paSalary").inputValue(),
  await paSetting("paCentre").inputValue(),
];
console.log(`pension age     : from Prognos ${paImported.join("/")}`);
if (paImported.join("/") !== `1959/23/38500/${paRiktalder(1959)}`) {
  problems.push(`"Hämta från Prognos" gives ${paImported.join("/")}, expected 1959/23/38500/${paRiktalder(1959)}`);
}
await pa.locator('[data-action="pa-reset"]').click();
await tab.waitForTimeout(80);

// Own savings: the individual occupational pension, the private pension and the private savings are
// asked for by the module "Hur länge räcker pensionen?" is built with, so the two tabs look the same and
// work the same -- the same fields in the same order, the same labels and small print, the same drop-down.
// The numbers are the engine's, and `pensionAge.test.ts` pins them against it and against the other tab's
// runs: 300 000 kr and 2 000 kr/month in a pension pot pay 4 869 kr/month at 68 on top of the public
// pension's 20 194, taxed as income; 400 000 kr and 1 000 kr/month in an ISK pay 3 939, in a depot 3 951,
// not taxed.

const paNorm = (text) => (text ?? "").replace(/\s/g, "").replace(/−/g, "-");
const paLabel = async (name) => paNorm(await pa.locator(`label:has([data-setting="${name}"]) .field-label`).textContent());
const paSequence = () =>
  pa.locator(".pa-grid").first().evaluate((grid) =>
    [...grid.children].map((el) =>
      el.classList.contains("pa-group")
        ? `#${el.textContent}`
        : el.classList.contains("field-pair")
          ? [...el.querySelectorAll("[data-setting]")].map((i) => i.dataset.setting).join("+")
          : (el.querySelector("[data-setting]")?.dataset.setting ?? el.className),
    ),
  );
const paPairRows = () =>
  pa.locator(".field-pair").evaluateAll((pairs) =>
    pairs.map((pair) => {
      const [a, b] = [...pair.querySelectorAll("input")].map((i) => i.getBoundingClientRect());
      return { sameRow: Math.abs(a.top - b.top) < 1.5, side: b.left > a.right - 1 };
    }),
  );
const paTimeHead = async () => (await pa.locator('[data-role="pa-time-table"] thead th').allTextContents()).map(paNorm);
const paTimeRow = async (par = 68) => (await pa.locator(`[data-role="pa-time-table"] tr[data-par="${par}"] td`).allTextContents()).map(paNorm);
const paFirstYear = async (par = 68) => Number((await paTimeRow(par))[1]);
const paKpi = async () => paNorm(await pa.locator('[data-role="pa-answer"] .pa-kpi').textContent());
const paNotes = async () => (await pa.locator(".pa-grid .saving-note").allTextContents()).map(paNorm);

// Where the fields sit, what they are called, and that capital and monthly saving share a row.
const paSeq = await paSequence();
const paWantedSeq = [
  "#Person", "paBorn", "paStartWork", "paCentre", "paRiktalder", "paSalaryAs", "paSalary",
  "#Pension", "paScheme", "paIncludeOccupational", "paCapital+paSaving", "paPrivCapital+paPrivSaving", "saving-note",
  "#Privat sparande", "paSavingsKind", "saving-note", "paSavingsCapital+paSavingsMonthly",
  "#Livslängd",
];
const paSeqAt = paSeq.indexOf("#Person");
if (paSeqAt === -1 || !sameList(paSeq.slice(paSeqAt, paSeqAt + paWantedSeq.length), paWantedSeq)) {
  problems.push(`the Hur maxa pensionen? inputs run ${paSeq.slice(paSeqAt, paSeqAt + paWantedSeq.length).join(" > ")}, expected ${paWantedSeq.join(" > ")}`);
}
const paLabelsSv = [await paLabel("paCapital"), await paLabel("paSaving"), await paLabel("paPrivCapital"), await paLabel("paPrivSaving"), await paLabel("paSavingsCapital"), await paLabel("paSavingsMonthly")];
console.log(`pension age     : saving labels ${paLabelsSv.join(" | ")}`);
if (!sameList(paLabelsSv, ["Individuelltjänstepension,kapital", "Individuelltjänstepension,månadssparande", "Privatpension,kapital", "Privatpension,månadssparande", "ISK/Kapitalförsäkring/Depåförsäkring,kapital", "ISK/Kapitalförsäkring/Depåförsäkring,månadssparande"])) {
  problems.push(`the age tab's capital and saving fields are labelled ${paLabelsSv.join(" | ")}`);
}
if ((await paLabel("paIncludeOccupational")) !== "Räknamedavtaletstjänstepension") {
  problems.push(`the include box is labelled "${await paLabel("paIncludeOccupational")}", expected that it is the agreement's occupational pension that it switches off`);
}
const paPairsDesktop = await paPairRows();
if (paPairsDesktop.length !== 3 || !paPairsDesktop.every((pair) => pair.sameRow && pair.side)) {
  problems.push(`at 1280px the age tab's capital and monthly saving are not side by side on one row: ${JSON.stringify(paPairsDesktop)}`);
}
const paKindOptions = (await paSetting("paSavingsKind").locator("option").allTextContents()).map((text) => text.trim());
if (!sameList(paKindOptions, ["ISK / Kapitalförsäkring / Depåförsäkring", "Depå / VP-konto / Fondkonto"])) {
  problems.push(`the age tab's kind of saving offers ${paKindOptions.join(" | ")}, expected the two tax categories`);
}
const paNotesStart = await paNotes();
if (!/Privatpensionärenpensionsförsäkringellerips\(individuelltpensionssparande\)/i.test(paNotesStart[0] ?? "")) {
  problems.push(`under the private pension fields the age tab says "${paNotesStart[0]}", expected that it is a pension insurance or an IPS`);
}
if (!/Schablonbeskattat/.test(paNotesStart[1] ?? "") || (await pa.locator(".pa-grid .saving-note-prompt").count()) !== 0) {
  problems.push(`for an ISK the age tab says "${paNotesStart[1]}" (prompt ${await pa.locator(".pa-grid .saving-note-prompt").count()}), expected the standard-return note and no prompt`);
}

// The Pension and Privat sparande sections' markup, kept for the other tab to be compared with (its panel is not on screen
// then). The box that switches the agreement off is the age tab's own and is left out.
const sectionsOf = (panel, prefix) =>
  panel.locator(".pa-grid").first().evaluate((grid, prefix) => {
    const kids = [...grid.children];
    const isGroup = (el, text) => el.classList.contains("pa-group") && (text === undefined || el.textContent === text);
    const start = kids.findIndex((el) => isGroup(el, "Pension"));
    const savings = kids.findIndex((el) => isGroup(el, "Privat sparande"));
    const end = kids.findIndex((el, i) => i > savings && isGroup(el));
    return kids
      .slice(start, end)
      .filter((el) => !el.querySelector('[data-setting$="IncludeOccupational"]'))
      .map((el) => el.outerHTML.replaceAll(`data-setting="${prefix}`, 'data-setting="X'));
  }, prefix);
const paSection = await sectionsOf(pa, "pa");

// A pension of one's own: one pot, taxed as income, from the alternative's own retirement age.
const paPlainFirst = await paFirstYear();
await paType("paCapital", 300000);
await paType("paSaving", 2000);
const paIpsFirst = await paFirstYear();
const paIpsHead = await paTimeHead();
const paIpsRow = await paTimeRow();
const paIpsKpi = await paKpi();
const paIpsBest = await paBest();
console.log(`pension age     : pension of one's own 300 000 + 2 000/month -> ${paPlainFirst} to ${paIpsFirst} at 68, best ${paIpsBest.join(",")}`);
if (paPlainFirst !== 20194 || paIpsFirst !== 25063) problems.push(`the first year at 68 reads ${paPlainFirst} without and ${paIpsFirst} with 300 000 kr and 2 000 kr/month of one's own pension, expected 20194 and 25063`);
if (!sameList(paIpsBest, [65, 65, 65, 67, 68, 68, 71, 71, 71, 71, 71, 71, 71, 71, 71, 71])) {
  problems.push(`with a pension of one's own the short answer picks ${paIpsBest.join(",")}, expected 65,65,65,67,68,68 then 71 from 86`);
}
if (!paIpsHead.includes("varavallmän") || !paIpsHead.includes("varavindividuellochprivat") || paIpsHead.includes("varavtjänstepension") || paIpsHead.includes("varavprivatsparande")) {
  problems.push(`before tax the per-age table has columns ${paIpsHead.join("|")}, expected public and individual-and-private only`);
}
if (paIpsRow[1] !== "25063" || paIpsRow[2] !== "20194" || paIpsRow[3] !== "4869") problems.push(`the first-year row splits ${paIpsRow.slice(1, 4).join(" + ")}, expected 25063 = 20194 + 4869`);
if (!/^Vid68år\(riktåldern\):25267kr\/mån\(varavindividuellochprivatpension4869kr\/mån\),[\d,]+%avslutlönen39872kr\/mån\.Allmänpension\+egetsparande,föreskatt\./.test(paIpsKpi)) {
  problems.push(`the short answer's figure reads "${paIpsKpi}", expected it to name the individual and private pension and the own saving`);
}
// The private pension is the same pot: the same numbers whichever field they are typed in, and the two add up.
await paType("paCapital", 0);
await paType("paSaving", 0);
await paType("paPrivCapital", 300000);
await paType("paPrivSaving", 2000);
const paPrivFirst = await paFirstYear();
await paType("paCapital", 150000);
await paType("paSaving", 1000);
await paType("paPrivCapital", 150000);
await paType("paPrivSaving", 1000);
const paSplitFirst = await paFirstYear();
if (paPrivFirst !== 25063 || paSplitFirst !== 25063) problems.push(`as private pension the first year reads ${paPrivFirst} and half and half ${paSplitFirst}, expected 25063 for both`);
for (const name of ["paCapital", "paSaving", "paPrivCapital", "paPrivSaving"]) await paType(name, 0);

// Private savings: what it is held in names the fields and says how it is taxed; paid out tax-free.
await paType("paSavingsCapital", 400000);
await paType("paSavingsMonthly", 1000);
const paIskFirst = await paFirstYear();
const paIskHead = await paTimeHead();
const paIskRow = await paTimeRow();
const paIskKpi = await paKpi();
console.log(`pension age     : ISK 400 000 + 1 000/month -> ${paIskFirst} at 68 (${paIskRow.slice(2, 4).join(" + ")})`);
if (paIskFirst !== 24133) problems.push(`with 400 000 kr and 1 000 kr/month in an ISK the first year at 68 reads ${paIskFirst}, expected 24133`);
if (!paIskHead.includes("varavprivatsparande") || paIskHead.includes("varavindividuellochprivat")) problems.push(`with an ISK the per-age table has columns ${paIskHead.join("|")}, expected one for the private savings`);
if (paIskRow[2] !== "20194" || paIskRow[3] !== "3939") problems.push(`the ISK's first-year row splits ${paIskRow.slice(2, 4).join(" + ")}, expected 20194 + 3939`);
if (!/\(varavprivatsparande3939kr\/mån\)/.test(paIskKpi) || !/Allmänpension\+egetsparande,föreskatt\.$/.test(paIskKpi)) problems.push(`with an ISK the short answer's figure reads "${paIskKpi}"`);
await paSetting("paSavingsKind").selectOption("1"); // Depå / VP-konto / Fondkonto
await tab.waitForTimeout(100);
const paDepotFirst = await paFirstYear();
const paDepotLabels = [await paLabel("paSavingsCapital"), await paLabel("paSavingsMonthly")];
const paDepotNote = (await paNotes())[1] ?? "";
const paDepotHint = paNorm(await pa.locator('label:has([data-setting="paSavingsCapital"]) .field-hint').textContent());
console.log(`pension age     : depot 400 000 + 1 000/month -> ${paDepotFirst} at 68, labels ${paDepotLabels.join(" | ")}`);
if (paDepotFirst !== 24145) problems.push(`in a depot the first year at 68 reads ${paDepotFirst}, expected 24145`);
if (!sameList(paDepotLabels, ["Depå/VP-konto/Fondkonto,kapital", "Depå/VP-konto/Fondkonto,månadssparande"])) problems.push(`choosing the depot labels the age tab's fields ${paDepotLabels.join(" | ")}`);
if (!/30%reavinstskatt.*Angekapitalvärdetefterreavinstskatt–detduharattleva/.test(paDepotNote) || (await pa.locator(".pa-grid .saving-note-prompt").count()) !== 1) {
  problems.push(`for a depot the age tab says "${paDepotNote}" (prompt ${await pa.locator(".pa-grid .saving-note-prompt").count()}), expected the prompt to enter the capital after capital gains tax`);
}
if (!/efterreavinstskatt/.test(paDepotHint)) problems.push(`the age tab's capital hint for a depot reads "${paDepotHint}", expected it to say after capital gains tax`);

// After tax the private savings are the same and the tax is not split out of the rest.
await pa.locator('button[data-basis="net"]').click();
await tab.waitForTimeout(100);
const paNetDepotFirst = await paFirstYear();
const paNetHead = await paTimeHead();
const paNetKpi = await paKpi();
if (paNetDepotFirst !== 18640) problems.push(`after tax a depot's first year at 68 reads ${paNetDepotFirst}, expected 18640 (14689 net public + 3951 untaxed)`);
if (paNetHead.some((text) => text.startsWith("varav"))) problems.push(`after tax the per-age table has columns ${paNetHead.join("|")}, expected no split`);
if (!/Allmänpension\+egetsparande,efterskatt\.$/.test(paNetKpi) || /varav/.test(paNetKpi)) problems.push(`after tax the short answer's figure reads "${paNetKpi}"`);
await pa.locator('button[data-basis="gross"]').click();
await tab.waitForTimeout(100);

// Everything at once, with an agreement: each kind in its own column, and the agreement's switch leaves the rest alone.
await paType("paCapital", 300000);
await paType("paSaving", 2000);
await paSetting("paScheme").selectOption("2"); // ITP-1
await tab.waitForTimeout(100);
const paAllFirst = await paFirstYear();
const paAllHead = await paTimeHead();
const paAllRow = await paTimeRow();
const paAllKpi = await paKpi();
console.log(`pension age     : everything + ITP-1 -> ${paAllFirst} at 68 (${paAllRow.slice(2, 6).join(" + ")})`);
if (paAllFirst !== 37221 || !sameList(paAllRow.slice(1, 6), ["37221", "20194", "8207", "4869", "3951"])) {
  problems.push(`everything at once splits ${paAllRow.slice(1, 6).join(" ")}, expected 37221 = 20194 + 8207 + 4869 + 3951`);
}
if (!sameList(paAllHead.slice(1, 6), ["Förstaåret", "varavallmän", "varavtjänstepension", "varavindividuellochprivat", "varavprivatsparande"])) {
  problems.push(`everything at once gives the per-age table columns ${paAllHead.slice(0, 7).join("|")}`);
}
if (!/\(varavtjänstepension8207,individuellochprivatpension4869,privatsparande3951kr\/mån\)/.test(paAllKpi)) problems.push(`everything at once, the short answer's figure reads "${paAllKpi}"`);
await paSetting("paIncludeOccupational").uncheck();
await tab.waitForTimeout(100);
const paAgreementOffFirst = await paFirstYear();
if (paAgreementOffFirst !== 29014) problems.push(`with the agreement switched off the first year reads ${paAgreementOffFirst}, expected 29014 (the own saving stays)`);
await paSetting("paIncludeOccupational").check();
await paSetting("paScheme").selectOption("1");
await tab.waitForTimeout(80);

// The notes say what is counted and how.
const paAllNotes = paNorm(await pa.locator('[data-role="pa-notes"]').textContent());
if (!/Individuelltjänstepensionochprivatpension\(enpensionsförsäkringellerIPS\)räknassommodellensprivatapensionssparande\(IPS\)/.test(paAllNotes)) problems.push("the notes do not explain the individual and private pension");
if (!/Privatsparandeärpengarsomredanärinkomstbeskattade.*Pådepå,VP-kontoochfondkontobeskattasvinstenmed30%närdusäljer/.test(paAllNotes)) problems.push("the notes do not explain the two tax categories of private savings");
if (!/Sparandetfortsätterframtillrespektivepensionsålder/.test(paAllNotes) || !/Privatsparandeharingenutbetalningförlivet/.test(paAllNotes)) problems.push("the notes do not say that the saving goes on until each retirement age and that private savings have no payout for life");

// English: the same fields, named in English.
await tab.locator(".langs button", { hasText: "EN" }).click();
await tab.waitForTimeout(100);
const paLabelsEn = [await paLabel("paCapital"), await paLabel("paPrivCapital"), await paLabel("paSavingsCapital"), await paLabel("paSavingsKind")];
if (!sameList(paLabelsEn, ["Individualoccupationalpension,capital", "Privatepension,capital", "Securitiesaccount/Fundaccount,capital", "Typeofsaving"])) {
  problems.push(`in English the age tab's saving fields are labelled ${paLabelsEn.join(" | ")}`);
}
await tab.locator(".langs button", { hasText: "SV" }).click();
await tab.waitForTimeout(100);

// Reset puts the savings back to none, on the ISK the drop-down opens on.
await pa.locator('[data-action="pa-reset"]').click();
await tab.waitForTimeout(100);
const paResetSavings = [];
for (const name of ["paCapital", "paSaving", "paPrivCapital", "paPrivSaving", "paSavingsKind", "paSavingsCapital", "paSavingsMonthly"]) paResetSavings.push(await paSetting(name).inputValue());
if (paResetSavings.join("/") !== "0/0/0/0/0/0/0" || (await paFirstYear()) !== 20194) {
  problems.push(`Återställ left the savings at ${paResetSavings.join("/")} (first year ${await paFirstYear()}), expected none and 20194`);
}
if (!sameList(await paBest(), PA_BEST_DEFAULT)) problems.push("Återställ did not restore the short answer after savings were typed");

// The exports: the matrix as CSV (whole kronor, Swedish delimiter) and a real .xlsx.
const paMatrixActions = pa.locator('[data-role="pa-matrix-panel"] .panel-actions');
const paCsv = await download(paMatrixActions.getByRole("button", { name: "Ladda ner CSV" }));
const paCsvLines = paCsv.buffer.toString("utf8").replace(/^﻿/, "").split("\r\n");
console.log(`pension age     : CSV ${paCsv.filename}, ${paCsvLines.length} lines`);
if (paCsvLines.length !== 17 || paCsvLines[0] !== "Lever till (år);65;66;67;68;69;70;71;Bäst;Försprång mot näst bäst (kr)") {
  problems.push(`the matrix CSV starts "${paCsvLines[0]}" and has ${paCsvLines.length} lines, expected the 17-line lifespan matrix`);
}
if (paCsvLines[1] !== "80;2972916;2942724;2901684;2851104;2760972;2668008;2568324;65;30192") {
  problems.push(`the matrix CSV's first row reads "${paCsvLines[1]}"`);
}
const paXlsx = await download(paMatrixActions.getByRole("button", { name: "Ladda ner Excel" }));
if (!looksLikeXlsx(paXlsx.buffer)) problems.push(`the matrix's Excel download ("${paXlsx.filename}") isn't a real .xlsx file`);
const paSeries = await download(pa.locator('[data-role="pa-time"] .panel-actions').getByRole("button", { name: "Ladda ner CSV" }));
const paSeriesLines = paSeries.buffer.toString("utf8").replace(/^﻿/, "").split("\r\n");
if (paSeriesLines.length !== 32 || !paSeriesLines[1]?.startsWith("65;16498;")) {
  problems.push(`the per-age CSV has ${paSeriesLines.length} lines starting "${paSeriesLines[1]}", expected 32 lines from age 65 at 16498`);
}

// "Optimal pensionsålder i förhållande till livslängd": a table and a chart of the best age per lifespan,
// straight after "Total pension per livslängd".
const paOptimal = pa.locator('[data-role="pa-optimal-panel"]');
const paOptimalAfter = await pa.locator(".pa-results").evaluate((box) => {
  const roles = [...box.children].map((node) => node.dataset.role);
  return roles[roles.indexOf("pa-matrix-panel") + 1];
});
const paOptimalCells = async (key) =>
  (await paOptimal.locator(`tr[data-row="${key}"] td.num`).allTextContents()).map((text) => text.replace(/\s/g, ""));
const paOptimalAges = await paOptimalCells("age");
const paOptimalMonthly = await paOptimalCells("monthly");
const paOptimalTotal = await paOptimalCells("total");
console.log(`pension age     : optimal ${paOptimalAfter}, ages ${paOptimalAges.join(",")}, 80 -> ${paOptimalMonthly[0]} kr, ${paOptimalTotal[0]} Mkr`);
if (paOptimalAfter !== "pa-optimal-panel") problems.push(`"Optimal pensionsålder" is not straight after the matrix (it is followed by ${paOptimalAfter})`);
if (!sameList(paOptimalAges.map(Number), PA_BEST_DEFAULT)) {
  problems.push(`the optimal-age table lists ${paOptimalAges.join(",")}, expected ${PA_BEST_DEFAULT.join(",")}`);
}
if (paOptimalMonthly.length !== 16 || paOptimalMonthly[0] !== "16516" || paOptimalTotal[0] !== "3,0" || paOptimalTotal[15] !== "6,5") {
  problems.push(`the optimal-age table's monthly row starts ${paOptimalMonthly[0]} and its total row ${paOptimalTotal[0]}...${paOptimalTotal[15]}, expected 16516 and 3,0...6,5`);
}
const paOptimalChart = paOptimal.locator('[data-role="pa-fig-optimal"]');
const paOptimalMarks = await paOptimalChart.evaluate((figure) => ({
  bars: figure.querySelectorAll("path.pa-opt-bar").length,
  linePoints: (figure.querySelector("polyline.pa-opt-line")?.getAttribute("points") ?? "").split(" ").length,
  grid: figure.querySelectorAll("svg line.grid, svg line.axis").length,
  left: figure.querySelectorAll("svg text.tick-y").length,
  right: [...figure.querySelectorAll("svg text")].filter((t) => t.getAttribute("x") === String(760 - 58 + 8)).length,
  legend: [...figure.querySelectorAll(".legend li")].map((li) => li.textContent),
}));
console.log(`pension age     : optimal chart ${paOptimalMarks.bars} bars, ${paOptimalMarks.linePoints} line points, ticks ${paOptimalMarks.left}/${paOptimalMarks.right}`);
if (paOptimalMarks.bars !== 16 || paOptimalMarks.linePoints !== 16) {
  problems.push(`the optimal-age chart has ${paOptimalMarks.bars} bars and ${paOptimalMarks.linePoints} line points, expected 16 and 16`);
}
// Two axes, but the same number of ticks on both, so every gridline is a tick on each.
if (paOptimalMarks.left < 3 || paOptimalMarks.left !== paOptimalMarks.right || paOptimalMarks.grid < paOptimalMarks.left) {
  problems.push(`the optimal-age chart's axes have ${paOptimalMarks.left} left and ${paOptimalMarks.right} right ticks over ${paOptimalMarks.grid} gridlines`);
}
if (!sameList(paOptimalMarks.legend, ["Månadsbelopp kr", "Livsutbetalning mkr"])) {
  problems.push(`the optimal-age chart's legend reads ${paOptimalMarks.legend.join(" / ")}`);
}
// The hover names the column under the pointer: a real mouse over the 88 bar.
await paOptimalChart.scrollIntoViewIfNeeded();
const paOptimalBar = await paOptimalChart.locator('path.pa-opt-bar[data-lifespan="88"]').boundingBox();
await tab.mouse.move(paOptimalBar.x + paOptimalBar.width / 2, paOptimalBar.y + paOptimalBar.height / 2);
await tab.waitForTimeout(100);
const paOptimalTip = ((await paOptimalChart.locator(".tooltip").textContent()) ?? "").replace(/\s+/g, " ");
if (!/Lever du till 88 år/.test(paOptimalTip) || !/Pensionsålder71/.test(paOptimalTip.replace(/ /g, "")) || !/23\s?106/.test(paOptimalTip) || !/4,7\s?Mkr/.test(paOptimalTip)) {
  problems.push(`the optimal-age chart's hover reads "${paOptimalTip}"`);
}
await tab.mouse.move(5, 5);
const paOptimalCsv = await download(paOptimal.locator(".panel-actions").getByRole("button", { name: "Ladda ner CSV" }));
const paOptimalCsvLines = paOptimalCsv.buffer.toString("utf8").replace(/^\uFEFF/, "").split("\r\n");
if (
  paOptimalCsv.filename !== "optimal-pensionsalder.csv" ||
  paOptimalCsvLines.length !== 4 ||
  paOptimalCsvLines[0] !== "Livslängd;80;81;82;83;84;85;86;87;88;89;90;91;92;93;94;95" ||
  !paOptimalCsvLines[2]?.startsWith("Månadsbelopp kr;16516;")
) {
  problems.push(`the optimal-age CSV is "${paOptimalCsv.filename}" with ${paOptimalCsvLines.length} lines, header "${paOptimalCsvLines[0]}"`);
}
const paOptimalXlsx = await download(paOptimal.locator(".panel-actions").getByRole("button", { name: "Ladda ner Excel" }));
if (paOptimalXlsx.filename !== "optimal-pensionsalder.xlsx" || !looksLikeXlsx(paOptimalXlsx.buffer)) {
  problems.push(`the optimal-age Excel download ("${paOptimalXlsx.filename}") isn't a real .xlsx file`);
}
await tab.locator(".langs button", { hasText: "EN" }).click();
await tab.waitForTimeout(100);
const paOptimalEn = [
  ((await paOptimal.locator("> h2 > span").textContent()) ?? "").trim(),
  (await paOptimal.locator('tr[data-row="total"] td').first().textContent()) ?? "",
];
await tab.locator(".langs button", { hasText: "SV" }).click();
await tab.waitForTimeout(100);
if (paOptimalEn[0] !== "Optimal retirement age in relation to lifespan" || paOptimalEn[1] !== "Lifetime payout MSEK") {
  problems.push(`the optimal-age section in English reads ${JSON.stringify(paOptimalEn)}`);
}

// The per-age table switches between the monthly pension and the total paid out, and its download follows.
const paTimeSwitch = (measure) => pa.locator(`[data-role="pa-time"] button[data-measure="${measure}"]`);
const paTimeHeading = async () => ((await pa.locator('[data-role="pa-time"] > h2 > span').textContent()) ?? "").trim();
if ((await paTimeSwitch("monthly").getAttribute("aria-pressed")) !== "true" || (await paTimeHeading()) !== "Pension per månad vid olika åldrar") {
  problems.push(`the per-age table does not start on its monthly view ("${await paTimeHeading()}")`);
}
await paTimeSwitch("total").click();
await tab.waitForTimeout(100);
const paSwHeads = await paTimeHead();
if (paSwHeads.includes("Förstaåret") || paSwHeads.includes("Slutlön") || paSwHeads[0] !== "Startvid" || paSwHeads[1] !== "70år") {
  problems.push(`the total view's table heads read ${paSwHeads.join(" | ")}`);
}
if ((await paTimeHeading()) !== "Total utbetalning före olika åldrar") {
  problems.push(`the total view's heading reads "${await paTimeHeading()}"`);
}
const paSwCsv = await download(pa.locator('[data-role="pa-time"] .panel-actions').getByRole("button", { name: "Ladda ner CSV" }));
const paSwLines = paSwCsv.buffer.toString("utf8").replace(/^\uFEFF/, "").split("\r\n");
const paSwAt70 = (paSwLines.find((line) => line.startsWith("70;")) ?? "").split(";");
const paSwCell = (await paTimeRow(65))[1];
console.log(`pension age     : total view "${paSwCsv.filename}", header "${paSwLines[0]}", 65 -> 70 = ${paSwAt70[1]} (table ${paSwCell})`);
if (paSwCsv.filename !== "total-utbetalning.csv" || paSwLines.length !== 32 || paSwLines[0]?.startsWith("Före ålder;Total utbetalning 65;") !== true) {
  problems.push(`the total view's CSV is "${paSwCsv.filename}", ${paSwLines.length} lines, header "${paSwLines[0]}"`);
}
if (paSwLines[0]?.includes("Pension per månad") || !paSwLines[1]?.startsWith("65;0;")) {
  problems.push(`the total view's CSV mixes in the monthly pension or does not start at nothing paid ("${paSwLines[1]}")`);
}
if (paSwAt70[1] === undefined || paNorm(paSwCell) !== paSwAt70[1]) {
  problems.push(`the total paid out before 70 reads ${paSwCell} in the table and ${paSwAt70[1]} in the CSV`);
}
const paSwXlsx = await download(pa.locator('[data-role="pa-time"] .panel-actions').getByRole("button", { name: "Ladda ner Excel" }));
if (paSwXlsx.filename !== "total-utbetalning.xlsx" || !looksLikeXlsx(paSwXlsx.buffer)) {
  problems.push(`the total view's Excel download ("${paSwXlsx.filename}") isn't the total .xlsx`);
}
// The choice survives a language change, and back to the month restores table, heading and downloads.
await tab.locator(".langs button", { hasText: "EN" }).click();
await tab.waitForTimeout(100);
const paSwEn = [await paTimeHeading(), await paTimeSwitch("total").textContent(), await paTimeSwitch("total").getAttribute("aria-pressed")];
await tab.locator(".langs button", { hasText: "SV" }).click();
await tab.waitForTimeout(100);
if (paSwEn[0] !== "Total payout before different ages" || paSwEn[1] !== "Total payout" || paSwEn[2] !== "true") {
  problems.push(`the total view in English reads ${JSON.stringify(paSwEn)}`);
}
await paTimeSwitch("monthly").click();
await tab.waitForTimeout(100);
const paSwMonthly = await download(pa.locator('[data-role="pa-time"] .panel-actions').getByRole("button", { name: "Ladda ner CSV" }));
const paSwMonthlyLines = paSwMonthly.buffer.toString("utf8").replace(/^\uFEFF/, "").split("\r\n");
if (
  paSwMonthly.filename !== "pension-per-ar.csv" ||
  paSwMonthlyLines[0] !== "Ålder;Pension per månad 65;Pension per månad 66;Pension per månad 67;Pension per månad 68;Pension per månad 69;Pension per månad 70;Pension per månad 71" ||
  (await paTimeHeading()) !== "Pension per månad vid olika åldrar" ||
  !(await paTimeHead()).includes("Förstaåret")
) {
  problems.push(`back on the month, the CSV is "${paSwMonthly.filename}" with header "${paSwMonthlyLines[0]}"`);
}

// The theme tokens reach the boxes: the reference age is ink, black on light and white on dark.
const paInk = async () => pa.locator(".pa-ribbon .pa-chip-z").first().evaluate((el) => getComputedStyle(el).backgroundColor);
const paInkLight = await paInk();
await tab.emulateMedia({ colorScheme: "dark" });
const paInkDark = await paInk();
await tab.emulateMedia({ colorScheme: "light" });
console.log(`pension age     : reference chip ${paInkLight} light, ${paInkDark} dark`);
if (paInkLight !== "rgb(0, 0, 0)" || paInkDark !== "rgb(255, 255, 255)") {
  problems.push(`the reference chip is ${paInkLight}/${paInkDark} on light/dark, expected black/white`);
}

// Swedish and English.
const paHeadSv = (await pa.locator("section.panel > h2 span").first().textContent()) ?? "";
await tab.locator(".langs button", { hasText: "EN" }).click();
await tab.waitForTimeout(100);
const paHeadEn = (await pa.locator("section.panel > h2 span").first().textContent()) ?? "";
const paLabelEn = (await pa.locator(".pa-ribbon-label").textContent()) ?? "";
const paButtonEn = (await tab.locator('.screen-toggle .panel-btn[data-screen="retirement"]').textContent()) ?? "";
console.log(`pension age     : "${paHeadSv}" / "${paHeadEn}", "${paLabelEn}", button "${paButtonEn}"`);
if (paHeadSv !== "Hur maxa pensionen?" || paHeadEn !== "How to maximise the pension?" || paLabelEn !== "Lifespan (years)" || paButtonEn !== "How to maximise the pension?") {
  problems.push(`the tab's language switch reads "${paHeadSv}"/"${paHeadEn}"/"${paLabelEn}"/"${paButtonEn}"`);
}
if (!sameList(await paBest(), PA_BEST_DEFAULT)) problems.push("switching language changed the short answer");
await tab.locator(".langs button", { hasText: "SV" }).click();
await tab.waitForTimeout(100);

// A phone: nothing widens the page, and the sixteen boxes wrap instead of overflowing.
await tab.setViewportSize({ width: 375, height: 800 });
await tab.waitForTimeout(100);
const paOverflow = await tab.evaluate(() => document.documentElement.scrollWidth - document.documentElement.clientWidth);
const paRibbonRows = await pa.locator(".pa-ribbon li").evaluateAll((items) => new Set(items.map((li) => Math.round(li.getBoundingClientRect().top))).size);
console.log(`pension age @375px: page overflow ${paOverflow}px, ribbon in ${paRibbonRows} rows`);
if (paOverflow > 1) problems.push(`the Hur maxa pensionen? tab widens the page by ${paOverflow}px at 375px`);
if (paRibbonRows < 2) problems.push("the sixteen answer boxes did not wrap onto several rows at 375px");
await tab.setViewportSize({ width: 1280, height: 1000 });
await tab.waitForTimeout(100);

// ---- Jämför avtalsområden ---------------------------------------------------
//
// The tab's second view: the same person under each collective agreement,
// side by side. What is checked is that the switch reaches it, that the
// agreements are the ones the cohort can belong to until somebody picks, that
// every number on the page is the engine's (the totals below are what
// `pensionAge.test.ts` pins against it, and the ITP-1 line reproduces the
// single view's 80/65 total above to the krona), and that each control --
// agreement chips, retirement age, the lifespan to read at, the chart toggle --
// reaches the figure, the grid and the table it should.

const expandRuns = (text) =>
  text.split(" ").flatMap((run) => {
    const [range, age] = run.split(":");
    const [from, to = from] = range.split("-").map(Number);
    return Array.from({ length: to - from + 1 }, () => Number(age));
  });
// Each agreement's best age at lifespans 80..95, born 1980, the workbook's assumptions.
const PA_AG_RUNS = {
  1: "80-82:65 83:66 84-86:68 87-95:71",
  2: "80-82:65 83-84:66 85-87:68 88-95:71",
  3: "80-81:65 82:67 83-85:68 86-88:70 89-95:71",
  4: "80-83:65 84:66 85-87:68 88-95:71",
  5: "80-82:65 83:66 84:67 85-87:68 88-95:71",
  6: "80-82:65 83:66 84-85:68 86-95:71",
  7: "80-82:65 83:66 84-85:68 86-87:69 88-95:71",
  8: "80-82:65 83:66 84-85:67 86-87:68 88-95:71",
};
// Total / added over having no agreement, in kronor, at the best age, by agreement (1 = none, 2 = ITP-1, 4 = SAF-LO, 5 = KAP-KL, 7 = PA16 avd 2).
const PA_AG_AT_85 = { 1: [3988140, 0], 2: [5662368, 1674228], 4: [5606472, 1618332], 5: [5661348, 1673208], 7: [5813328, 1825188] };
const PA_AG_AT_90 = { 1: [5233356, 0], 2: [7422612, 2189256], 4: [7353072, 2119716], 5: [7421244, 2187888], 7: [7655400, 2422044] };
const PA_AG_AT_85_RETIRING_65 = { 1: [3931176, 0], 2: [5618616, 1687440], 4: [5576856, 1645680], 5: [5596776, 1665600], 7: [5708376, 1777200] };

const cmpControls = pa.locator('[data-role="pa-compare-controls"]');
const cmpPressed = () =>
  cmpControls
    .locator("button.pa-agr")
    .evaluateAll((buttons) => buttons.filter((b) => b.getAttribute("aria-pressed") === "true").map((b) => Number(b.dataset.scheme)));
const cmpGrid = () =>
  pa.locator('[data-role="pa-compare-grid"] tbody tr').evaluateAll((rows) =>
    rows.map((tr) => ({ scheme: Number(tr.dataset.scheme), best: [...tr.querySelectorAll("td.pa-gcell")].map((td) => Number(td.dataset.best)) })),
  );
const cmpRow = (scheme) => pa.locator(`[data-role="pa-compare-table"] tr[data-scheme="${scheme}"]`);
const cmpKr = async (scheme, nth) => Number(await cmpRow(scheme).locator("td[data-kr]").nth(nth).getAttribute("data-kr"));
const cmpFigures = async (expected, label) => {
  const got = {};
  for (const scheme of Object.keys(expected)) got[scheme] = [await cmpKr(scheme, 0), await cmpKr(scheme, 1)];
  if (JSON.stringify(got) !== JSON.stringify(expected)) {
    problems.push(`${label}: the table reads ${JSON.stringify(got)}, expected ${JSON.stringify(expected)}`);
  }
};
const cmpLines = () => pa.locator('[data-role="pa-fig-compare"] svg polyline.pa-line').evaluateAll((lines) => lines.map((l) => Number(l.dataset.agreement)));
const cmpTitle = (role) => pa.locator(`[data-role="${role}"] > h2 span, [data-role="${role}"] figcaption`).first().textContent();

// The switch: closed by default, the single view's boxes gone once it is open.
const cmpHiddenAtFirst = await cmpControls.isHidden();
const cmpViewButtons = await pa.locator(".pa-view:not(.pa-time-view) .panel-btn").count();
await pa.locator('.pa-view .panel-btn[data-view="compare"]').click();
await tab.waitForTimeout(150);
const cmpShown = await cmpControls.isVisible();
const cmpRibbonGone = (await pa.locator('[data-role="pa-ribbon"]').count()) === 0;
console.log(`agreements      : switch ${cmpViewButtons} buttons, controls hidden at first ${cmpHiddenAtFirst}, shown ${cmpShown}, single boxes gone ${cmpRibbonGone}`);
if (cmpViewButtons !== 2 || !cmpHiddenAtFirst || !cmpShown || !cmpRibbonGone) {
  problems.push("the Jämför avtalsområden switch does not open and close the comparison");
}

// Until somebody picks: the agreements a 1980 cohort can belong to.
const cmpDefault = await cmpPressed();
const cmpOffCohort = await cmpControls.locator("button.pa-agr.pa-agr-off").evaluateAll((buttons) => buttons.map((b) => Number(b.dataset.scheme)));
console.log(`agreements      : picked ${cmpDefault.join(",")}, marked "what if" ${cmpOffCohort.join(",")}`);
if (cmpDefault.join(",") !== "2,4,5,7" || cmpOffCohort.join(",") !== "3,6,8") {
  problems.push(`a 1980 cohort starts with agreements ${cmpDefault.join(",")} picked and ${cmpOffCohort.join(",")} marked "what if", expected 2,4,5,7 and 3,6,8`);
}

// The grid: every agreement's best age at every lifespan, to the engine's.
const cmpRows = await cmpGrid();
console.log(`agreements      : grid rows ${cmpRows.map((r) => r.scheme).join(",")}`);
if (cmpRows.map((r) => r.scheme).join(",") !== "1,2,4,5,7") {
  problems.push(`the grid has rows for ${cmpRows.map((r) => r.scheme).join(",")}, expected the baseline and 2,4,5,7`);
}
for (const row of cmpRows) {
  if (!sameList(row.best, expandRuns(PA_AG_RUNS[row.scheme]))) {
    problems.push(`agreement ${row.scheme}'s best ages read ${row.best.join(",")}, expected ${expandRuns(PA_AG_RUNS[row.scheme]).join(",")}`);
  }
}
if (!sameList(cmpRows[0].best, PA_BEST_DEFAULT)) problems.push("the baseline row does not match the single view's short answer");

// The table, read at 85 by default: what each pays and what it adds, to the krona.
const cmpTitle85 = (await cmpTitle("pa-compare-table-panel")) ?? "";
await cmpFigures(PA_AG_AT_85, "at 85, best age");
console.log(`agreements      : "${cmpTitle85.trim()}", ITP-1 ${await cmpKr(2, 0)} / +${await cmpKr(2, 1)}`);
if (!/85/.test(cmpTitle85)) problems.push(`the table is titled "${cmpTitle85}", expected it to read at 85`);
// The ITP-1 line is the single view's 80/65 total, the same engine number reached by another route.
const cmpCsvLike = await pa.locator('[data-role="pa-compare-grid"] tr[data-scheme="2"] td[data-lifespan="80"]').getAttribute("title");
if (!/bäst är pension vid 65 \(4\s238\stkr\)/.test(cmpCsvLike ?? "")) problems.push(`the grid cell for ITP-1 at 80 carries "${cmpCsvLike}", expected its best age (65) and total (4 238 tkr)`);

// The figure: one line per row, the baseline heavy, keyed by agreement; the lines are the agreements' own colours.
const cmpLinesDefault = await cmpLines();
const cmpHeavy = await pa.locator('[data-role="pa-fig-compare"] svg polyline.pa-line-z').evaluateAll((lines) => lines.map((l) => Number(l.dataset.agreement)));
const cmpLegend = await pa.locator('[data-role="pa-fig-compare"] .legend li').count();
console.log(`agreements      : lines ${cmpLinesDefault.join(",")}, heavy ${cmpHeavy.join(",")}, legend ${cmpLegend}`);
if ([...cmpLinesDefault].sort().join(",") !== "1,2,4,5,7" || cmpHeavy.join(",") !== "1" || cmpLegend !== 5) {
  problems.push(`the comparison figure draws ${cmpLinesDefault.join(",")} (heavy ${cmpHeavy.join(",")}, ${cmpLegend} legend entries), expected 1,2,4,5,7 with the baseline heavy`);
}
const cmpStroke = (scheme) =>
  pa.locator(`[data-role="pa-fig-compare"] svg polyline[data-agreement="${scheme}"]`).evaluate((line) => getComputedStyle(line).stroke);
const cmpKey = (scheme) => cmpControls.locator(`.pa-key-ag${scheme}`).first().evaluate((el) => getComputedStyle(el).backgroundColor);
const cmpBlueLight = await cmpStroke(4);
const cmpInkLight = await cmpStroke(1);
const cmpKeyLight = await cmpKey(4);
await tab.emulateMedia({ colorScheme: "dark" });
const cmpBlueDark = await cmpStroke(4);
const cmpInkDark = await cmpStroke(1);
const cmpKeyDark = await cmpKey(4);
await tab.emulateMedia({ colorScheme: "light" });
console.log(`agreements      : SAF-LO ${cmpBlueLight} light, ${cmpBlueDark} dark; baseline ${cmpInkLight} / ${cmpInkDark}`);
if (cmpBlueLight !== "rgb(42, 120, 214)" || cmpBlueDark !== "rgb(57, 135, 229)" || cmpKeyLight !== cmpBlueLight || cmpKeyDark !== cmpBlueDark) {
  problems.push(`SAF-LO's line reads ${cmpBlueLight}/${cmpBlueDark} (key ${cmpKeyLight}/${cmpKeyDark}) on light/dark, expected rgb(42, 120, 214) and rgb(57, 135, 229)`);
}
if (cmpInkLight !== "rgb(0, 0, 0)" || cmpInkDark !== "rgb(255, 255, 255)") {
  problems.push(`the baseline line is ${cmpInkLight}/${cmpInkDark} on light/dark, expected black/white`);
}

// Hover: one row per line, named by text.
const cmpFigure = pa.locator('[data-role="pa-fig-compare"]');
await cmpFigure.scrollIntoViewIfNeeded();
const cmpBox = await cmpFigure.locator("svg").boundingBox();
await tab.mouse.move(cmpBox.x + cmpBox.width * 0.4, cmpBox.y + cmpBox.height * 0.5);
await tab.mouse.move(cmpBox.x + cmpBox.width * 0.45, cmpBox.y + cmpBox.height * 0.5);
await tab.waitForTimeout(100);
const cmpTipRows = await cmpFigure.locator(".tooltip .tooltip-row").allTextContents();
const cmpTipHead = (await cmpFigure.locator(".tooltip strong").textContent()) ?? "";
console.log(`agreements      : hover "${cmpTipHead}", ${cmpTipRows.length} rows`);
if (cmpTipRows.length !== 5 || !/^Lever till \d+$/.test(cmpTipHead) || !cmpTipRows.some((text) => /ITP-1/.test(text) && /Mkr/.test(text))) {
  problems.push(`hovering the comparison figure shows "${cmpTipHead}" with ${cmpTipRows.length} rows, expected one per line naming the agreement`);
}
await tab.mouse.move(5, 5);

// The agreement chosen in the inputs is shown whether or not the cohort fits it, and marked.
await paSetting("paScheme").selectOption("6");
await tab.waitForTimeout(150);
const cmpWithMine = await cmpPressed();
const cmpMineRow = await cmpRow(6).evaluate((tr) => tr.classList.contains("pa-mine"));
const cmpMineText = (await pa.locator('[data-role="pa-compare-answer"]').textContent()) ?? "";
console.log(`agreements      : AKAP-KR chosen -> picked ${cmpWithMine.join(",")}, marked row ${cmpMineRow}`);
if (cmpWithMine.join(",") !== "2,4,5,6,7" || !cmpMineRow || !/Ditt val/.test(cmpMineText) || (await cmpKr(6, 0)) !== 5872080) {
  problems.push(`choosing AKAP-KR above gives picked ${cmpWithMine.join(",")}, marked row ${cmpMineRow}, total ${await cmpKr(6, 0)}; expected 2,4,5,6,7, marked, 5872080`);
}
await paSetting("paScheme").selectOption("1");
await tab.waitForTimeout(100);

// The cohort decides what is shown until somebody picks: born 1990 swaps the agreements.
await paType("paBorn", 1990);
const cmp1990 = await cmpPressed();
console.log(`agreements      : born 1990 -> picked ${cmp1990.join(",")}`);
if (cmp1990.join(",") !== "2,4,5,6,8") problems.push(`born 1990 starts with agreements ${cmp1990.join(",")}, expected 2,4,5,6,8`);
await paType("paBorn", 1980);
if ((await cmpPressed()).join(",") !== "2,4,5,7") problems.push("going back to 1980 did not bring back the 1980 cohort's agreements");

// Picking: a chip switches an agreement on, and a "what if" is labelled as one.
await cmpControls.locator('button.pa-agr[data-scheme="3"]').click();
await tab.waitForTimeout(150);
const cmpWith3 = await cmpGrid();
const cmpItp2 = cmpWith3.find((r) => r.scheme === 3);
const cmpWhatIfTag = await cmpRow(3).locator(".pa-tag-off").count();
console.log(`agreements      : ITP-2 on -> rows ${cmpWith3.map((r) => r.scheme).join(",")}, "what if" tags ${cmpWhatIfTag}, lines ${(await cmpLines()).length}`);
if (cmpWith3.map((r) => r.scheme).join(",") !== "1,2,3,4,5,7" || !cmpItp2 || !sameList(cmpItp2.best, expandRuns(PA_AG_RUNS[3])) || cmpWhatIfTag !== 1) {
  problems.push("switching ITP-2 on did not add its row, its best ages or its \"what if\" tag");
}
if ((await cmpLines()).length !== 6) problems.push("switching ITP-2 on did not add its line to the figure");
await cmpControls.locator('button.pa-agr[data-scheme="3"]').click();
await tab.waitForTimeout(100);
await cmpControls.locator('button.pa-agr[data-scheme="2"]').click();
await tab.waitForTimeout(100);
const cmpWithout2 = (await cmpGrid()).map((r) => r.scheme).join(",");
if (cmpWithout2 !== "1,4,5,7" || (await cmpPressed()).join(",") !== "4,5,7") {
  problems.push(`switching ITP-2 back off and ITP-1 off leaves rows ${cmpWithout2}, expected 1,4,5,7`);
}
// A colour follows its agreement, not its position: SAF-LO is still blue with ITP-1 gone.
if ((await cmpStroke(4)) !== cmpBlueLight) problems.push("switching an agreement off repainted the others");
await cmpControls.locator('button.pa-agr[data-scheme="2"]').click();
await tab.waitForTimeout(100);

// The retirement age the totals are read at.
await paSetting("paCompareBasis").selectOption("65");
await tab.waitForTimeout(150);
await cmpFigures(PA_AG_AT_85_RETIRING_65, "at 85, retiring at 65");
const cmpAgeCells = await pa.locator('[data-role="pa-compare-table"] tbody tr').evaluateAll((rows) => rows.map((tr) => Number(tr.dataset.age)));
const cmpFixedHead = (await pa.locator('[data-role="pa-compare-table"] thead th').allTextContents()).join("|");
console.log(`agreements      : retiring at 65 -> ages ${cmpAgeCells.join(",")}, columns ${cmpFixedHead.length > 0}`);
if (!cmpAgeCells.every((age) => age === 65) || !/Bästa ålder vid 85/.test(cmpFixedHead)) {
  problems.push(`reading at retirement age 65 gives ages ${cmpAgeCells.join(",")} and columns "${cmpFixedHead}", expected all 65 and a best-age column`);
}
await paSetting("paCompareBasis").selectOption("best");
await tab.waitForTimeout(150);
await cmpFigures(PA_AG_AT_85, "back to the best age");

// The lifespan the table and the answer are read at.
await paType("paCompareLifespan", 90);
await cmpFigures(PA_AG_AT_90, "at 90, best age");
const cmpTitle90 = (await cmpTitle("pa-compare-table-panel")) ?? "";
const cmpReadColumns = await pa.locator('[data-role="pa-compare-grid"] th.pa-read').evaluateAll((cells) => cells.map((th) => th.textContent));
console.log(`agreements      : read at 90 -> "${cmpTitle90.trim()}", highlighted column ${cmpReadColumns.join(",")}`);
if (!/90/.test(cmpTitle90) || cmpReadColumns.join(",") !== "90") problems.push("reading at 90 did not move the table's title and the grid's highlighted column");
await paType("paCompareLifespan", 85);

// The chart toggle: what each agreement adds, no line for the baseline.
await cmpControls.locator('button[data-metric="extra"]').click();
await tab.waitForTimeout(150);
const cmpExtraLines = await cmpLines();
const cmpExtraTitle = (await cmpTitle("pa-fig-compare")) ?? "";
const cmpExtraTicks = await cmpFigure.locator("svg text.tick-y").allTextContents();
console.log(`agreements      : added -> lines ${cmpExtraLines.join(",")}, "${cmpExtraTitle.trim()}", ticks ${cmpExtraTicks.slice(0, 3).join(" ")}`);
if (cmpExtraLines.join(",") !== "2,4,5,7" || !cmpExtraTicks.some((text) => text.startsWith("+"))) {
  problems.push(`the "added" chart draws ${cmpExtraLines.join(",")} with ticks ${cmpExtraTicks.join(" ")}, expected 2,4,5,7 and "+" ticks`);
}
await cmpControls.locator('button[data-metric="total"]').click();
await tab.waitForTimeout(100);

// After tax: the figure drops the occupational column (it cannot be split out of net income) and the totals fall.
await pa.locator('button[data-basis="net"]').click();
await tab.waitForTimeout(150);
const cmpNetHead = (await pa.locator('[data-role="pa-compare-table"] thead th').allTextContents()).join("|");
const cmpNetBaseline = await cmpKr(1, 0);
const cmpNetBest = await cmpGrid();
console.log(`agreements      : after tax -> baseline ${cmpNetBaseline}, split column ${/varav/.test(cmpNetHead)}`);
if (/varav/.test(cmpNetHead) || !(cmpNetBaseline > 0 && cmpNetBaseline < PA_AG_AT_85[1][0]) || cmpNetBest.length !== 5) {
  problems.push(`after tax the comparison reads baseline ${cmpNetBaseline} with columns "${cmpNetHead}", expected a lower total and no occupational column`);
}
await pa.locator('button[data-basis="gross"]').click();
await tab.waitForTimeout(150);

// Own savings are in every agreement's total, the same for each, so what an agreement adds is what it was.
await paType("paCapital", 300000);
await paType("paSaving", 2000);
await cmpFigures({ 1: [4981416, 0], 2: [6655644, 1674228], 4: [6599748, 1618332], 5: [6654624, 1673208], 7: [6806604, 1825188] }, "with a pension of one's own");
const cmpOwnMonthly = paNorm(await cmpRow(1).locator("td").nth(5).textContent());
const cmpOwnLead = paNorm(await pa.locator('[data-role="pa-compare-table-panel"] .pa-lead').first().textContent());
console.log(`agreements      : with a pension of one's own -> baseline ${await cmpKr(1, 0)} / monthly ${cmpOwnMonthly}, ITP-1 +${await cmpKr(2, 1)}`);
if (cmpOwnMonthly !== "25063") problems.push(`with a pension of one's own the baseline's monthly pension at the start reads ${cmpOwnMonthly}, expected 25063`);
if (!/allmänpension\+tjänstepension\+egetsparande,föreskatt/.test(cmpOwnLead)) problems.push(`with own savings the comparison describes its sums as "${cmpOwnLead.slice(0, 160)}"`);
await paType("paCapital", 0);
await paType("paSaving", 0);
await cmpFigures(PA_AG_AT_85, "once the pension of one's own is taken away");

// The exports: best ages and totals per agreement and lifespan, whole kronor.
const cmpActions = pa.locator('[data-role="pa-compare-grid-panel"] .panel-actions');
const cmpCsv = await download(cmpActions.getByRole("button", { name: "Ladda ner CSV" }));
const cmpCsvLines = cmpCsv.buffer.toString("utf8").replace(/^﻿/, "").split("\r\n");
console.log(`agreements      : CSV ${cmpCsv.filename}, ${cmpCsvLines.length} lines`);
if (cmpCsvLines.length !== 17 || !cmpCsvLines[0].startsWith("Lever till (år);Saknar tjänstepension: bästa ålder;Saknar tjänstepension: total före skatt (kr);ITP-1: bästa ålder;")) {
  problems.push(`the comparison CSV starts "${cmpCsvLines[0]}" and has ${cmpCsvLines.length} lines, expected the 17-line per-agreement table`);
}
if (cmpCsvLines[1] !== "80;65;2972916;65;4238496;65;4207176;65;4222116;65;4305816") {
  problems.push(`the comparison CSV's first row reads "${cmpCsvLines[1]}"`);
}
const cmpXlsx = await download(cmpActions.getByRole("button", { name: "Ladda ner Excel" }));
if (!looksLikeXlsx(cmpXlsx.buffer)) problems.push(`the comparison's Excel download ("${cmpXlsx.filename}") isn't a real .xlsx file`);

// Swedish and English; the numbers do not move with the language.
const cmpHeadSv = (await cmpTitle("pa-compare-controls")) ?? "";
await tab.locator(".langs button", { hasText: "EN" }).click();
await tab.waitForTimeout(150);
const cmpHeadEn = (await cmpTitle("pa-compare-controls")) ?? "";
const cmpTableEn = (await cmpTitle("pa-compare-table-panel")) ?? "";
const cmpNameEn = (await cmpRow(7).locator(".pa-name-text").textContent()) ?? "";
const cmpViewEn = (await pa.locator('.pa-view .panel-btn[data-view="compare"]').textContent()) ?? "";
console.log(`agreements      : "${cmpHeadSv.trim()}" / "${cmpHeadEn.trim()}", "${cmpTableEn.trim()}", PA16 as "${cmpNameEn}", switch "${cmpViewEn}"`);
if (cmpHeadSv.trim() !== "Jämför avtalsområden" || cmpHeadEn.trim() !== "Compare agreements" || cmpTableEn.trim() !== "The agreements at 85" || cmpNameEn !== "PA16 section 2" || cmpViewEn !== "Compare agreements") {
  problems.push(`the comparison's language switch reads "${cmpHeadSv}"/"${cmpHeadEn}"/"${cmpTableEn}"/"${cmpNameEn}"/"${cmpViewEn}"`);
}
await cmpFigures(PA_AG_AT_85, "in English");
await tab.locator(".langs button", { hasText: "SV" }).click();
await tab.waitForTimeout(150);

// A phone: nothing widens the page; the grid scrolls inside its own box and the chips wrap.
await tab.setViewportSize({ width: 375, height: 800 });
await tab.waitForTimeout(150);
const cmpOverflow = await tab.evaluate(() => document.documentElement.scrollWidth - document.documentElement.clientWidth);
const cmpGridScrolls = await pa.locator('[data-role="pa-compare-grid-panel"] .scroll').evaluate((box) => box.scrollWidth > box.clientWidth);
const cmpChipRows = await cmpControls.locator(".pa-chips > *").evaluateAll((items) => new Set(items.map((el) => Math.round(el.getBoundingClientRect().top))).size);
console.log(`agreements @375px: page overflow ${cmpOverflow}px, grid scrolls ${cmpGridScrolls}, chips in ${cmpChipRows} rows`);
if (cmpOverflow > 1 || !cmpGridScrolls || cmpChipRows < 2) {
  problems.push(`at 375px the comparison widens the page by ${cmpOverflow}px (grid scrolls ${cmpGridScrolls}, chips in ${cmpChipRows} rows)`);
}
await tab.setViewportSize({ width: 1280, height: 1000 });
await tab.waitForTimeout(100);

// Reset puts the picks, the age, the lifespan and the chart back; the single view still says what it did.
await cmpControls.locator('button.pa-agr[data-scheme="8"]').click();
await paSetting("paCompareBasis").selectOption("66");
await cmpControls.locator('button[data-metric="extra"]').click();
await tab.waitForTimeout(100);
await pa.locator('[data-action="pa-reset"]').click();
await tab.waitForTimeout(150);
const cmpAfterReset = [(await cmpPressed()).join(","), await paSetting("paCompareBasis").inputValue(), await paSetting("paCompareLifespan").inputValue(), (await cmpLines()).length];
console.log(`agreements      : after reset ${cmpAfterReset.join(" / ")}`);
if (cmpAfterReset.join("/") !== "2,4,5,7/best/85/5") problems.push(`Återställ left the comparison at ${cmpAfterReset.join("/")}, expected 2,4,5,7/best/85/5`);
await pa.locator('.pa-view .panel-btn[data-view="single"]').click();
await tab.waitForTimeout(150);
if (!sameList(await paBest(), PA_BEST_DEFAULT) || !(await cmpControls.isHidden())) {
  problems.push("going back to the single view did not restore its short answer and hide the comparison's controls");
}

await tab.locator('.screen-toggle .panel-btn[data-screen="single"]').click();
await tab.waitForTimeout(50);

// ---- Hur länge räcker pensionen? --------------------------------------------
//
// A fifth top-level view: one typfall's pension set against a budget under every
// way of taking the occupational pension out -- 5 to 20 years, or for life --
// run as one typfall each. As above, the point of each check is that a control
// reaches the runs it should and that the page says what the engine said. The
// numbers pinned are the ones `pensionLasts.test.ts` and the engine give for
// born 1980, work from 25, 38 500 kr/month, ITP-1, an average pensioner's
// budget (18 400 kr/month after tax) and retiring at the riktålder (68). The tab
// opens on a lifespan of 95, which is pinned once; the rest of the block works
// at 90, where its anchors were worked out. The page has no numbers of its own.

/** A cell's text with the spaces thousands are grouped by taken out, and the typographic minus made plain. */
const plNorm = (text) => (text ?? "").replace(/\s/g, "").replace(/−/g, "-");

await tab.locator('.screen-toggle .panel-btn[data-screen="lasts"]').click();
await tab.waitForTimeout(150);

const pl = tab.locator('[data-role="pension-lasts-panel"]');
const plSetting = (name) => pl.locator(`[data-setting="${name}"]`);
const plType = async (name, value) => {
  const field = plSetting(name);
  await field.fill(String(value));
  await field.dispatchEvent("change");
  // A field that was typed into fires one more change when it loses focus, and that redraws the
  // results. Let it happen here, so that each step starts from a page that has settled; a press
  // made while a field is still being typed in is what "The first press" below checks.
  await field.evaluate((el) => el.blur());
  await tab.waitForTimeout(80);
};
const plRibbon = () =>
  pl.locator('[data-role="pl-ribbon"] li').evaluateAll((items) =>
    items.map((li) => ({
      payout: Number(li.dataset.payout),
      status: li.dataset.status,
      best: li.dataset.best === "1",
      text: (li.querySelector(".pl-ribbon-d")?.textContent ?? "").replace(/\s/g, ""),
    })),
  );
const plRow = (payout) => pl.locator(`[data-role="pl-options"] tbody tr[data-payout="${payout}"]`);
const plCells = async (payout) => (await plRow(payout).locator("td").allTextContents()).map(plNorm);
const plLines = (figure) => pl.locator(`[data-role="${figure}"] svg polyline.pa-line`).evaluateAll((lines) => lines.map((l) => Number(l.dataset.payout)));
const plMatrixHead = async () => (await pl.locator('[data-role="pl-matrix"] thead th').allTextContents()).map(plNorm);
const plMatrixStatus = (payout) =>
  pl.locator(`[data-role="pl-matrix"] tbody tr[data-payout="${payout}"] td[data-status]`).evaluateAll((cells) => cells.map((td) => td.dataset.status));
const plTimeCells = async (payout) => (await pl.locator(`[data-role="pl-time-table"] tbody tr[data-payout="${payout}"] td`).allTextContents()).map(plNorm);

const PL_RIBBON_DEFAULT = ["Till73år", "Till78år", "Till83år", "Till88år", "Helavägen"];

if ((await pl.count()) !== 1) problems.push("the Hur länge räcker pensionen? tab did not open");

// The inputs start as the person the file comment describes.
const plStart = {
  born: await plSetting("plBorn").inputValue(),
  start: await plSetting("plStartWork").inputValue(),
  retirement: await plSetting("plRetirement").inputValue(),
  retirementLocked: (await plSetting("plRetirement").isDisabled()) && (await plSetting("plRiktalder").isChecked()),
  lifespan: await plSetting("plLifespan").inputValue(),
  salary: await plSetting("plSalary").inputValue(),
  budgetChoice: await plSetting("plBudgetChoice").inputValue(),
  budget: await plSetting("plBudget").inputValue(),
  scheme: await plSetting("plScheme").inputValue(),
  capital: await plSetting("plCapital").inputValue(),
  saving: await plSetting("plSaving").inputValue(),
  privCapital: await plSetting("plPrivCapital").inputValue(),
  privSaving: await plSetting("plPrivSaving").inputValue(),
  savingsKind: await plSetting("plSavingsKind").inputValue(),
  savingsCapital: await plSetting("plSavingsCapital").inputValue(),
  savingsMonthly: await plSetting("plSavingsMonthly").inputValue(),
  net: await pl.locator('button[data-basis="net"]').evaluate((b) => b.classList.contains("active")),
};
console.log(`pension lasts   : inputs ${JSON.stringify(plStart)}`);
if (
  plStart.born !== "1980" || plStart.start !== "25" || plStart.retirement !== "68" || !plStart.retirementLocked ||
  plStart.lifespan !== "95" || plStart.salary !== "38500" || plStart.budgetChoice !== "1" || plStart.budget !== "18400" ||
  plStart.scheme !== "2" || plStart.capital !== "0" || plStart.saving !== "0" || plStart.privCapital !== "0" ||
  plStart.privSaving !== "0" || plStart.savingsKind !== "0" || plStart.savingsCapital !== "0" || plStart.savingsMonthly !== "0" || !plStart.net
) {
  problems.push(`the Hur länge räcker pensionen? inputs start at ${JSON.stringify(plStart)}`);
}

// The short answer: five boxes, the lifelong payout the best, every fixed payout ending where it should.
const plDefault = await plRibbon();
console.log(`pension lasts   : ribbon ${plDefault.map((item) => `${item.payout}:${item.text}`).join(" ")}`);
if (!sameList(plDefault.map((item) => item.payout), [5, 10, 15, 20, 0]) || !sameList(plDefault.map((item) => item.text), PL_RIBBON_DEFAULT)) {
  problems.push(`the answer boxes read ${plDefault.map((item) => `${item.payout}:${item.text}`).join(" ")}, expected 5,10,15,20,0 and ${PL_RIBBON_DEFAULT.join(",")}`);
}
if (!sameList(plDefault.map((item) => item.status), ["part", "part", "part", "part", "ok"]) || !sameList(plDefault.filter((item) => item.best).map((item) => item.payout), [0])) {
  problems.push(`the answer boxes' statuses ${plDefault.map((item) => item.status).join(",")} / best ${plDefault.filter((item) => item.best).map((item) => item.payout).join(",")}, expected part x4 then ok, best 0`);
}
// The lifespan it opens on, 95, pinned once: 27 pension years, and the numbers that moves. Then on to 90.
const plAnswer95 = plNorm(await pl.locator('[data-role="pl-answer"]').textContent());
const plLife95 = await plCells(0);
const plFive95 = await plCells(5);
const plFloor95 = await plCells("floor");
console.log(`pension lasts   : at 95 lifelong ${plLife95.slice(1, 7).join(" ")}, 5 years ${plFive95.slice(1, 6).join(" ")}, public only ${plFloor95[1]}`);
if (!/15928kr\/mån/.test(plAnswer95)) problems.push("at a lifespan of 95 the short answer does not give the public pension alone as 15 928 kr/month");
if (plLife95[1] !== "22131" || plLife95[3] !== "+3731/–" || plLife95[4] !== "Helavägentill95år" || plLife95[5] !== "27/27" || plLife95[6] !== "0") {
  problems.push(`at 95 the lifelong row reads ${plLife95.join(" | ")}, expected 22131, +3731/–, all the way to 95, 27/27, 0`);
}
if (plFive95[1] !== "38997" || plFive95[2] !== "15822" || plFive95[3] !== "+20597/-2578" || plFive95[4] !== "Till73år" || plFive95[5] !== "5/27") {
  problems.push(`at 95 the 5-year row reads ${plFive95.join(" | ")}, expected 38997, 15822, +20597/-2578, to 73, 5/27`);
}
if (plFloor95[1] !== "15928" || plFloor95[4] !== "Räckerinte") problems.push(`at 95 the public-pension-only row reads ${plFloor95.join(" | ")}, expected 15928 that does not last`);
await plType("plLifespan", 90);

const plAnswerText = plNorm(await pl.locator('[data-role="pl-answer"]').textContent());
if (!/18400kr\/månefterskatt/.test(plAnswerText) || !/Medelpensionär/.test(plAnswerText) || !/16094kr\/mån/.test(plAnswerText)) {
  problems.push("the short answer does not name the average pensioner's 18 400 kr/month and the public pension alone (16 094 kr/month)");
}

// The options table: the public pension alone, then seventeen payouts, the best one marked.
const plRowCount = await pl.locator('[data-role="pl-options"] tbody tr').count();
const plBestRows = await pl.locator('[data-role="pl-options"] tbody tr.pl-best').evaluateAll((rows) => rows.map((tr) => tr.dataset.payout));
const plLife = await plCells(0);
const plFive = await plCells(5);
const plFloor = await plCells("floor");
console.log(`pension lasts   : table ${plRowCount} rows, best ${plBestRows.join(",")}, lifelong ${plLife.slice(1, 5).join(" ")}, 5 years ${plFive.slice(1, 6).join(" ")}`);
if (plRowCount !== 18 || !sameList(plBestRows, ["0"])) problems.push(`the options table has ${plRowCount} rows with ${plBestRows.join(",")} marked, expected 18 and the lifelong row`);
if (plLife[1] !== "22289" || plLife[2] !== "–" || plLife[3] !== "+3889/–" || plLife[4] !== "Helavägentill90år" || plLife[5] !== "22/22" || plLife[6] !== "0") {
  problems.push(`the lifelong row reads ${plLife.join(" | ")}, expected 22289, –, +3889/–, all the way to 90, 22/22, 0`);
}
if (plFive[1] !== "38997" || plFive[2] !== "16005" || plFive[3] !== "+20597/-2395" || plFive[4] !== "Till73år" || plFive[5] !== "5/22") {
  problems.push(`the 5-year row reads ${plFive.join(" | ")}, expected 38997, 16005, +20597/-2395, to 73, 5/22`);
}
if (plFloor[1] !== "16094" || plFloor[4] !== "Räckerinte") problems.push(`the public-pension-only row reads ${plFloor.join(" | ")}, expected 16094 that does not last`);

// The first chart: one line per drawn payout and the budget, the lifelong line heavy, only the budget dashed.
const plIncomeLines = await plLines("pl-fig-income");
const plHeavy = await pl.locator('[data-role="pl-fig-income"] svg polyline.pa-line-z').evaluateAll((lines) => lines.map((l) => Number(l.dataset.payout)));
const plDashed = await pl.locator('[data-role="pl-fig-income"] svg polyline.pa-line-dashed').evaluateAll((lines) => lines.map((l) => Number(l.dataset.payout)));
const plEndLabels = (await pl.locator('[data-role="pl-fig-income"] svg text.pa-end').allTextContents()).join(",");
const plLegend = await pl.locator('[data-role="pl-fig-income"] .legend li').count();
console.log(`pension lasts   : income lines ${plIncomeLines.join(",")}, heavy ${plHeavy.join(",")}, dashed ${plDashed.join(",")}, labels ${plEndLabels}, legend ${plLegend}`);
if (!sameList([...plIncomeLines].sort((a, b) => a - b), [-1, 0, 5, 10, 15, 20])) problems.push(`the income chart draws lines for ${plIncomeLines.join(",")}, expected the budget (-1) and 0,5,10,15,20`);
if (!sameList(plHeavy, [0]) || !sameList(plDashed, [-1])) problems.push(`the income chart's heavy line is ${plHeavy.join(",")} and its dashed one ${plDashed.join(",")}, expected 0 and the budget (-1)`);
for (const wanted of ["5 år", "10 år", "15 år", "20 år", "Livsvarigt", "Budget"]) {
  if (!plEndLabels.split(",").includes(wanted)) problems.push(`the income chart has no label "${wanted}" (labels: ${plEndLabels})`);
}
if (plLegend !== 6) problems.push(`the income chart's legend has ${plLegend} entries, expected 6`);

// Its hover readout names the budget and every payout; the arrow keys move it.
// A field that was typed into fires one more change when it loses focus, and that redraws the results under the chart; let it settle first.
await tab.evaluate(() => document.activeElement instanceof HTMLElement && document.activeElement.blur());
await tab.waitForTimeout(150);
const plIncomeFigure = pl.locator('[data-role="pl-fig-income"]');
await plIncomeFigure.scrollIntoViewIfNeeded();
const plBox = await plIncomeFigure.locator("svg").boundingBox();
if (plBox) {
  const k = plBox.width / 760;
  // 78 of 68-89, in the chart's own geometry (68 left, 686 right of 760: 74 kept for the labels).
  await tab.mouse.move(plBox.x + 300 * k, plBox.y + 150 * k);
  await tab.mouse.move(plBox.x + (68 + (10 / 21) * (686 - 68)) * k, plBox.y + 150 * k);
  await tab.waitForTimeout(100);
}
const plTipRows = await plIncomeFigure.locator(".tooltip .tooltip-row").count();
const plTipHead = (await plIncomeFigure.locator(".tooltip strong").textContent()) ?? "";
console.log(`pension lasts   : hover "${plTipHead}", ${plTipRows} rows`);
if (plTipRows !== 6 || !/78/.test(plTipHead)) problems.push(`the income chart's hover shows "${plTipHead}" with ${plTipRows} rows at age 78, expected 6 rows`);
await plIncomeFigure.locator("svg").focus();
await tab.keyboard.press("ArrowRight");
const plKeyHead = (await plIncomeFigure.locator(".tooltip strong").textContent()) ?? "";
if (!/79/.test(plKeyHead)) problems.push(`the arrow key moved the readout to "${plKeyHead}", expected age 79`);
await tab.keyboard.press("Escape");

// The table twin of that chart, and the second chart: the running total, with nothing "below" yet.
const plTimeBudget = await plTimeCells(-1);
const plTimeFive = await plTimeCells(5);
console.log(`pension lasts   : per-age table budget ${plTimeBudget.slice(1).join(" ")}, 5 years ${plTimeFive.slice(1).join(" ")}`);
if (!sameList(plTimeBudget.slice(1), ["18400", "18400", "18400", "18400", "18400"]) || !sameList(plTimeFive.slice(1), ["35510", "39906", "16456", "16073", "15705"])) {
  problems.push(`the per-age table reads ${plTimeBudget.slice(1).join(",")} for the budget and ${plTimeFive.slice(1).join(",")} for 5 years, expected 18400 x5 and 35510,39906,16456,16073,15705`);
}
const plBalanceLines = await plLines("pl-fig-balance");
const plBalanceNotes = await pl.locator('[data-role="pl-fig-balance"] svg text.pa-note').count();
if (!sameList([...plBalanceLines].sort((a, b) => a - b), [0, 5, 10, 15, 20])) problems.push(`the running-total chart draws ${plBalanceLines.join(",")}, expected 0,5,10,15,20`);
if (plBalanceNotes !== 1) problems.push(`the running-total chart has ${plBalanceNotes} margin notes while no line goes below zero, expected 1`);

// The budgets table: three examples, every payout, and what each budget lasts to.
const plMatrixRows = await pl.locator('[data-role="pl-matrix"] tbody tr').count();
const plLifeStatus = await plMatrixStatus(0);
const plFiveMatrix = (await pl.locator('[data-role="pl-matrix"] tbody tr[data-payout="5"] td').allTextContents()).map(plNorm);
console.log(`pension lasts   : budgets ${plMatrixRows} rows, head ${(await plMatrixHead()).length}, lifelong ${plLifeStatus.join(",")}, 5 years ${plFiveMatrix.join(" ")}`);
if (plMatrixRows !== 17 || (await plMatrixHead()).length !== 5 || !sameList(plLifeStatus, ["ok", "ok", "no"])) {
  problems.push(`the budgets table has ${plMatrixRows} rows, ${(await plMatrixHead()).length} columns and lifelong statuses ${plLifeStatus.join(",")}, expected 17, 5 and ok,ok,no`);
}
if (!sameList(plFiveMatrix, ["5år", "Till73år", "Till73år", "Till73år", "35510"])) problems.push(`the 5-year row of the budgets table reads ${plFiveMatrix.join(" | ")}`);

// The budget: minPension's examples set the amount; typing one's own selects "Egen budget".
await plSetting("plBudgetChoice").selectOption("2"); // Pluspensionär
await tab.waitForTimeout(100);
const plPlus = await plRibbon();
const plPlusBudget = await plSetting("plBudget").inputValue();
console.log(`pension lasts   : Pluspensionär -> budget ${plPlusBudget}, ribbon ${plPlus.map((item) => `${item.payout}:${item.text}`).join(" ")}, best ${plPlus.filter((item) => item.best).map((item) => item.payout).join(",")}`);
if (plPlusBudget !== "27300") problems.push(`choosing Pluspensionär puts ${plPlusBudget} in the budget field, expected 27300`);
if (!sameList(plPlus.map((item) => item.payout), [5, 10, 11, 15, 20, 0]) || !sameList(plPlus.filter((item) => item.best).map((item) => item.payout), [11])) {
  problems.push(`with Pluspensionär the boxes are ${plPlus.map((item) => item.payout).join(",")} and the best ${plPlus.filter((item) => item.best).map((item) => item.payout).join(",")}, expected 5,10,11,15,20,0 and 11`);
}
if (!sameList(plPlus.map((item) => item.text), ["Till73år", "69–78år", "69–79år", "Räckerinte", "Räckerinte", "Räckerinte"]) || plPlus[5]?.status !== "no") {
  problems.push(`with Pluspensionär the boxes read ${plPlus.map((item) => item.text).join(",")}, expected Till73år,69–78år,69–79år and three that do not last`);
}
const plPlusLife = await plMatrixStatus(0);
if (!sameList(plPlusLife, ["ok", "ok", "no"])) problems.push("the budgets table changed with the chosen budget; it should read the same three examples");
const plPlusNotes = await pl.locator('[data-role="pl-fig-balance"] svg text.pa-note').count();
if (plPlusNotes !== 2) problems.push(`the running-total chart has ${plPlusNotes} margin notes once lines go below zero, expected 2`);
const plPlusAnswer = plNorm(await pl.locator('[data-role="pl-answer"]').textContent());
if (!/Längsttäcksbudgetenmed11årsutbetalning:10av22år\(69–79år\)/.test(plPlusAnswer)) problems.push("the short answer does not say that 11 years covers the budget longest, 10 of 22 years (69-79)");

await plType("plBudget", 12345);
const plOwnChoice = await plSetting("plBudgetChoice").inputValue();
const plOwn = await plRibbon();
console.log(`pension lasts   : own 12 345 -> choice ${plOwnChoice}, ribbon ${plOwn.map((item) => item.text).join(",")}, budgets columns ${(await plMatrixHead()).length}`);
if (plOwnChoice !== "3" || !plOwn.every((item) => item.text === "Helavägen") || (await plMatrixHead()).length !== 6) {
  problems.push(`an own budget of 12 345 gives choice ${plOwnChoice}, boxes ${plOwn.map((item) => item.text).join(",")} and ${(await plMatrixHead()).length} budget columns, expected 3, all the way x5 and 6`);
}
await plType("plBudget", 18400);
const plBack = await plSetting("plBudgetChoice").inputValue();
if (plBack !== "1" || (await plMatrixHead()).length !== 5) problems.push(`typing 18 400 gives choice ${plBack} and ${(await plMatrixHead()).length} columns, expected the average pensioner (1) and 5`);
await plSetting("plBudgetChoice").selectOption("3"); // Egen budget keeps the amount; it is no column of its own while it is an example's
await tab.waitForTimeout(80);
if ((await plSetting("plBudget").inputValue()) !== "18400" || (await plSetting("plBudgetChoice").inputValue()) !== "3" || (await plMatrixHead()).length !== 5) {
  problems.push("choosing \"Egen budget\" should keep the amount and stay chosen, without repeating an example's column");
}
await plType("plBudget", 19000);
if ((await plSetting("plBudgetChoice").inputValue()) !== "3" || (await plMatrixHead()).length !== 6) {
  problems.push("an own budget of 19 000 should stay chosen and be a column of its own in the budgets table");
}
await plType("plBudget", 18400);
await plSetting("plBudgetChoice").selectOption("1");
await tab.waitForTimeout(80);

// Before tax: the same years covered, in gross amounts, with the budget grossed up by the model's own tax.
await pl.locator('button[data-basis="gross"]').click();
await tab.waitForTimeout(150);
const plGrossRibbon = await plRibbon();
const plGrossHead = (await pl.locator('[data-role="pl-options"] thead th').allTextContents()).map(plNorm);
const plGrossFive = await plCells(5);
const plGrossBudget = await plTimeCells(-1);
const plGrossAnswer = plNorm(await pl.locator('[data-role="pl-answer"]').textContent());
console.log(`pension lasts   : before tax -> ribbon ${plGrossRibbon.map((item) => item.text).join(",")}, 5 years ${plGrossFive[1]}, budget row ${plGrossBudget.slice(1, 3).join(" ")}`);
if (!sameList(plGrossRibbon.map((item) => item.text), PL_RIBBON_DEFAULT)) problems.push(`before tax the boxes read ${plGrossRibbon.map((item) => item.text).join(",")}, expected the same years as after tax`);
if (!plGrossHead.includes("varavallmän") || plGrossFive[1] !== "51927") problems.push(`before tax the options table has columns ${plGrossHead.join("|")} and 51927 for 5 years? ${plGrossFive[1]}`);
if (plGrossBudget[1] !== "25946" || plGrossBudget[2] !== "22204") problems.push(`the budget before tax reads ${plGrossBudget.slice(1, 3).join(",")} at 68 and 70, expected 25946 (the first year is taxed harder) and 22204`);
if (!/ungefär22204kr\/månföreskatt/.test(plGrossAnswer)) problems.push("the short answer does not say that 18 400 after tax is about 22 204 before tax");
if ((await pl.locator('[data-role="pl-fig-income"] .figure-note').count()) !== 1) problems.push("the income chart has no note on the budget before tax");
await pl.locator('button[data-basis="net"]').click();
await tab.waitForTimeout(100);

// Individual occupational pension: capital and monthly saving add to every payout.
await plType("plCapital", 300000);
await plType("plSaving", 2000);
const plIndLife = await plCells(0);
const plIndFive = await plCells(5);
const plCapitalHint = (await pl.locator('label:has([data-setting="plCapital"]) .field-hint').textContent()) ?? "";
console.log(`pension lasts   : 300 000 + 2 000/month -> lifelong ${plIndLife[1]}, 5 years ${plIndFive[1]}, hint "${plCapitalHint}"`);
if (plIndLife[1] !== "25949" || plIndFive[1] !== "49495") problems.push(`with 300 000 kr and 2 000 kr/month the lifelong payout reads ${plIndLife[1]} and 5 years ${plIndFive[1]}, expected 25949 and 49495`);
if (!/2026/.test(plCapitalHint)) problems.push(`the capital field's hint says "${plCapitalHint}", expected it to name the year the capital is counted from (2026)`);
await plType("plCapital", 0);
await plType("plSaving", 0);

// The inputs: where they sit, what they are called, and that capital and monthly saving share a row.
const plLabel = async (name) => plNorm((await pl.locator(`label:has([data-setting="${name}"]) .field-label`).textContent()) ?? "");
const plSequence = () =>
  pl.locator(".pa-grid").evaluate((grid) =>
    [...grid.children].map((el) =>
      el.classList.contains("pa-group")
        ? `#${el.textContent}`
        : el.classList.contains("field-pair")
          ? [...el.querySelectorAll("[data-setting]")].map((i) => i.dataset.setting).join("+")
          : (el.querySelector("[data-setting]")?.dataset.setting ?? el.className),
    ),
  );
const plPairRows = () =>
  pl.locator(".field-pair").evaluateAll((pairs) =>
    pairs.map((pair) => {
      const [a, b] = [...pair.querySelectorAll("input")].map((i) => i.getBoundingClientRect());
      return { sameRow: Math.abs(a.top - b.top) < 1.5, side: b.left > a.right - 1 };
    }),
  );
const plSeq = await plSequence();
const plWantedSeq = ["#Pension", "plScheme", "plCapital+plSaving", "plPrivCapital+plPrivSaving", "saving-note", "#Privat sparande", "plSavingsKind", "saving-note", "plSavingsCapital+plSavingsMonthly", "#Antaganden"];
const plSeqAt = plSeq.indexOf("#Pension");
if (plSeqAt === -1 || !sameList(plSeq.slice(plSeqAt, plSeqAt + plWantedSeq.length), plWantedSeq)) {
  problems.push(`the input order from "Pension" on is ${plSeq.slice(plSeqAt, plSeqAt + plWantedSeq.length).join(" > ")}, expected ${plWantedSeq.join(" > ")}`);
}
const plLabelsSv = [await plLabel("plCapital"), await plLabel("plSaving"), await plLabel("plPrivCapital"), await plLabel("plPrivSaving"), await plLabel("plSavingsCapital"), await plLabel("plSavingsMonthly")];
console.log(`pension lasts   : labels ${plLabelsSv.join(" | ")}`);
if (!sameList(plLabelsSv, ["Individuelltjänstepension,kapital", "Individuelltjänstepension,månadssparande", "Privatpension,kapital", "Privatpension,månadssparande", "ISK/Kapitalförsäkring/Depåförsäkring,kapital", "ISK/Kapitalförsäkring/Depåförsäkring,månadssparande"])) {
  problems.push(`the capital and saving fields are labelled ${plLabelsSv.join(" | ")}`);
}
const plPairsDesktop = await plPairRows();
if (plPairsDesktop.length !== 3 || !plPairsDesktop.every((pair) => pair.sameRow && pair.side)) {
  problems.push(`at 1280px the capital and monthly saving are not side by side on one row: ${JSON.stringify(plPairsDesktop)}`);
}
const plKindOptions = (await plSetting("plSavingsKind").locator("option").allTextContents()).map((text) => text.trim());
if (!sameList(plKindOptions, ["ISK / Kapitalförsäkring / Depåförsäkring", "Depå / VP-konto / Fondkonto"])) {
  problems.push(`the kind of saving offers ${plKindOptions.join(" | ")}, expected the two tax categories: ISK / Kapitalförsäkring / Depåförsäkring and Depå / VP-konto / Fondkonto`);
}
// What the person is told: a private pension is a pension insurance or an IPS; each kind of saving says how it is taxed.
const plNotes = async () => (await pl.locator(".pa-grid .saving-note").allTextContents()).map(plNorm);
const plNotesStart = await plNotes();
if (!/Privatpensionärenpensionsförsäkringellerips\(individuelltpensionssparande\)/i.test(plNotesStart[0] ?? "")) {
  problems.push(`under the private pension fields the page says "${plNotesStart[0]}", expected that it is a pension insurance or an IPS`);
}
if (!/Schablonbeskattat/.test(plNotesStart[1] ?? "") || (await pl.locator(".pa-grid .saving-note-prompt").count()) !== 0) {
  problems.push(`for an ISK the page says "${plNotesStart[1]}" (prompt ${await pl.locator(".pa-grid .saving-note-prompt").count()}), expected the standard-return note and no prompt`);
}

// Private pension: the same pot as the individual occupational pension, so the same numbers; the two add up.
await plType("plPrivCapital", 300000);
await plType("plPrivSaving", 2000);
const plPrivLife = await plCells(0);
const plPrivFive = await plCells(5);
console.log(`pension lasts   : private pension 300 000 + 2 000/month -> lifelong ${plPrivLife[1]}, 5 years ${plPrivFive[1]}`);
if (plPrivLife[1] !== "25949" || plPrivFive[1] !== "49495") problems.push(`with 300 000 kr and 2 000 kr/month as private pension the lifelong payout reads ${plPrivLife[1]} and 5 years ${plPrivFive[1]}, expected 25949 and 49495, as for an individual pension`);
await plType("plPrivCapital", 150000);
await plType("plPrivSaving", 1000);
await plType("plCapital", 150000);
await plType("plSaving", 1000);
const plSplitLife = await plCells(0);
const plSplitFive = await plCells(5);
if (plSplitLife[1] !== "25949" || plSplitFive[1] !== "49495") problems.push(`half as individual and half as private pension the lifelong payout reads ${plSplitLife[1]} and 5 years ${plSplitFive[1]}, expected 25949 and 49495`);
for (const name of ["plPrivCapital", "plPrivSaving", "plCapital", "plSaving"]) await plType(name, 0);

// Private savings: what it is held in decides how it is taxed, and names the two fields.
await plType("plSavingsCapital", 300000);
await plType("plSavingsMonthly", 2000);
const plIskLife = await plCells(0);
const plIskFive = await plCells(5);
console.log(`pension lasts   : ISK 300 000 + 2 000/month -> lifelong ${plIskLife[1]}, 5 years ${plIskFive[1]}`);
if (plIskLife[1] !== "26441" || plIskFive[1] !== "55871") problems.push(`with 300 000 kr and 2 000 kr/month in an ISK the lifelong payout reads ${plIskLife[1]} and 5 years ${plIskFive[1]}, expected 26441 and 55871 (tax-free, so more than as a pension)`);
await plSetting("plSavingsKind").selectOption("1"); // Depå / VP-konto / Fondkonto
await tab.waitForTimeout(100);
const plDepotLife = await plCells(0);
const plDepotFive = await plCells(5);
const plDepotLabels = [await plLabel("plSavingsCapital"), await plLabel("plSavingsMonthly")];
const plDepotNote = (await plNotes())[1] ?? "";
const plDepotHint = plNorm(await pl.locator('label:has([data-setting="plSavingsCapital"]) .field-hint').textContent());
console.log(`pension lasts   : depot 300 000 + 2 000/month -> lifelong ${plDepotLife[1]}, 5 years ${plDepotFive[1]}, labels ${plDepotLabels.join(" | ")}`);
if (plDepotLife[1] !== "26483" || plDepotFive[1] !== "56022") problems.push(`in a depot the lifelong payout reads ${plDepotLife[1]} and 5 years ${plDepotFive[1]}, expected 26483 and 56022`);
if (!sameList(plDepotLabels, ["Depå/VP-konto/Fondkonto,kapital", "Depå/VP-konto/Fondkonto,månadssparande"])) problems.push(`choosing the depot labels the fields ${plDepotLabels.join(" | ")}`);
if (!/30%reavinstskatt.*Angekapitalvärdetefterreavinstskatt–detduharattleva/.test(plDepotNote) || (await pl.locator(".pa-grid .saving-note-prompt").count()) !== 1) {
  problems.push(`for a depot the page says "${plDepotNote}" (prompt ${await pl.locator(".pa-grid .saving-note-prompt").count()}), expected the prompt to enter the capital after capital gains tax`);
}
if (!/efterreavinstskatt/.test(plDepotHint)) problems.push(`the capital field's hint for a depot reads "${plDepotHint}", expected it to say after capital gains tax`);

// Before tax, an ISK's payout is grossed up so that the same years are covered; the table says so.
await plSetting("plSavingsKind").selectOption("0");
await tab.waitForTimeout(100);
const plIskNetRibbon = (await plRibbon()).map((item) => item.text);
await pl.locator('button[data-basis="gross"]').click();
await tab.waitForTimeout(150);
const plIskGrossRibbon = (await plRibbon()).map((item) => item.text);
const plIskGrossHead = (await pl.locator('[data-role="pl-options"] thead th').allTextContents()).map(plNorm);
const plIskGrossFive = await plCells(5);
const plIskNotes = plNorm(await pl.locator('[data-role="pl-notes"]').textContent());
console.log(`pension lasts   : ISK before tax -> ribbon ${plIskGrossRibbon.join(",")}, 5 years ${plIskGrossFive.slice(1, 6).join(" ")}`);
if (!sameList(plIskGrossRibbon, plIskNetRibbon)) problems.push(`with an ISK the boxes read ${plIskGrossRibbon.join(",")} before tax and ${plIskNetRibbon.join(",")} after, expected the same`);
if (!plIskGrossHead.includes("varavprivatsparande,uppräknat")) problems.push(`before tax with an ISK the table has columns ${plIskGrossHead.join("|")}, expected one for the private savings grossed up`);
if (!sameList(plIskGrossFive.slice(1, 6), ["78965", "20146", "31781", "0", "27039"])) {
  problems.push(`before tax the 5-year row reads ${plIskGrossFive.slice(1, 6).join(" ")}, expected 78965 as 20146 + 31781 + 0 + 27039 (the ISK payout grossed up)`);
}
if (!/Föreskatträknasutbetalningenfråndetprivatasparandetupptillden/.test(plIskNotes)) problems.push("the notes do not say that a tax-free payout is grossed up before tax");
if (!/Privatsparandeharingenutbetalningförlivet/.test(plIskNotes)) problems.push("the notes do not say that private savings have no payout for life");
if (!/Privatsparandeärpengarsomredanärinkomstbeskattade.*Pådepå,VP-kontoochfondkontobeskattasvinstenmed30%närdusäljer/.test(plIskNotes)) {
  problems.push("the notes do not explain the two tax categories of private savings");
}
await pl.locator('button[data-basis="net"]').click();
await tab.waitForTimeout(100);
await plType("plSavingsCapital", 0);
await plType("plSavingsMonthly", 0);
const plIskGone = await plCells(0);
if (plIskGone[1] !== "22289") problems.push(`emptying the ISK left the lifelong payout at ${plIskGone[1]}, expected 22289`);
if ((await pl.locator('[data-role="pl-options"] thead th').count()) !== 8) problems.push("without anything tax-free the options table should not have the ISK column");

// The age tab's Pension and Privat sparande sections are this tab's, field by field: the same markup, by construction
// (one module builds both; the age tab's one extra box, which switches its agreement off, is left out). Its markup was taken
// while that tab was on screen (`paSection`).
const plSection = await sectionsOf(pl, "pl");
const sectionDiffers = paSection.findIndex((html, i) => html !== plSection[i]);
console.log(`pension lasts   : Pension and Privat sparande sections ${paSection.length} elements vs ${plSection.length} in the age tab, first difference at ${sectionDiffers}`);
if (paSection.length < 8 || paSection.length !== plSection.length || sectionDiffers !== -1) {
  problems.push(`the age tab's Pension and Privat sparande sections are not this tab's: ${paSection.length} against ${plSection.length} elements, first difference at ${sectionDiffers}: ${paSection[sectionDiffers] ?? ""} vs ${plSection[sectionDiffers] ?? ""}`);
}

// No agreement and nothing saved: nothing for a payout period to apply to, and the page says so.
await plSetting("plScheme").selectOption("1");
await tab.waitForTimeout(100);
const plNone = await plRibbon();
const plNoneAnswer = plNorm(await pl.locator('[data-role="pl-answer"]').textContent());
console.log(`pension lasts   : no agreement -> ${plNone.map((item) => item.status).join(",")}, best ${plNone.filter((item) => item.best).length}`);
if (!plNone.every((item) => item.status === "no") || plNone.some((item) => item.best) || !/ingentjänstepension/.test(plNoneAnswer)) {
  problems.push("with no agreement the answer should mark no payout best and say there is no occupational pension to choose a payout period for");
}
await plSetting("plScheme").selectOption("2");
await tab.waitForTimeout(100);

// Retirement age: unlock the riktålder and move it; an age below the cohort's lowest is clamped.
await plSetting("plRiktalder").uncheck();
await plType("plRetirement", 65);
const plR65 = await plRibbon();
console.log(`pension lasts   : retiring at 65 -> ${plR65.map((item) => item.text).join(",")}`);
if (!sameList(plR65.map((item) => item.text), ["Till70år", "Till75år", "Till80år", "69–85år", "Från69år"])) {
  problems.push(`retiring at 65 the boxes read ${plR65.map((item) => item.text).join(",")}, expected Till70år,Till75år,Till80år,69–85år,Från69år`);
}
await plType("plRetirement", 62);
const plClamped = await plSetting("plRetirement").inputValue();
if (plClamped !== "65") problems.push(`a retirement age of 62 (the cohort's lowest is 65) shows ${plClamped}, expected 65`);
await plSetting("plRiktalder").check();
await tab.waitForTimeout(80);
if ((await plSetting("plRetirement").inputValue()) !== "68") problems.push("ticking the riktålder box did not put the retirement age back on 68");

// The lifespan: a shorter one lets the longer payouts last all the way.
await plType("plLifespan", 80);
const plL80 = await plRibbon();
const plL80Five = await plCells(5);
console.log(`pension lasts   : living to 80 -> ${plL80.map((item) => item.text).join(",")}, 5 years ${plL80Five[5]}`);
if (!sameList(plL80.map((item) => item.text), ["Till73år", "Till78år", "Helavägen", "Helavägen", "Helavägen"]) || plL80Five[5] !== "5/12") {
  problems.push(`living to 80 the boxes read ${plL80.map((item) => item.text).join(",")} and 5 years ${plL80Five[5]}, expected Till73år,Till78år then all the way x3 and 5/12`);
}
await plType("plLifespan", 60);
if ((await plSetting("plLifespan").inputValue()) !== "70") problems.push(`a lifespan of 60 shows ${await plSetting("plLifespan").inputValue()}, expected 70`);
await plType("plLifespan", 90);

// The riktålder follows the birth year while its box is ticked.
await plType("plBorn", 1990);
const plBorn90 = await plSetting("plRetirement").inputValue();
const plBorn90Ribbon = await plRibbon();
console.log(`pension lasts   : born 1990 -> retirement ${plBorn90}, ${plBorn90Ribbon.map((item) => item.text).join(",")}`);
if (plBorn90 !== String(paRiktalder(1990)) || !sameList(plBorn90Ribbon.map((item) => item.text), ["Till74år", "Till79år", "Till84år", "Till89år", "Från70år"])) {
  problems.push(`born 1990 puts retirement at ${plBorn90} with boxes ${plBorn90Ribbon.map((item) => item.text).join(",")}, expected ${paRiktalder(1990)} and Till74år,Till79år,Till84år,Till89år,Från70år`);
}
await plType("plBorn", 1980);

// A final salary as the input: the monthly salary behind it is worked out and shown.
await plSetting("plSalaryAs").selectOption("1");
await tab.waitForTimeout(80);
const plFinalShown = await plSetting("plSalary").inputValue();
await plType("plSalary", 40000);
const plMonthlyHint = (await pl.locator('label:has([data-setting="plSalary"]) .field-hint').textContent()) ?? "";
console.log(`pension lasts   : final salary shown ${plFinalShown}, hint "${plMonthlyHint}"`);
if (plFinalShown !== "39872") problems.push(`switching to final salary shows ${plFinalShown}, expected 39872 (the model's own final salary)`);
if (!/38\D?623/.test(plMonthlyHint)) problems.push(`typing a final salary of 40000 gives the hint "${plMonthlyHint}", expected the monthly salary 38 623`);
const plRaised = await plCells(0);
if (!(Number(plRaised[1]) > 22289)) problems.push("a higher final salary did not raise the lifelong payout");
await plSetting("plSalaryAs").selectOption("0");
await tab.waitForTimeout(80);
if ((await plSetting("plSalary").inputValue()) !== "38623") problems.push(`switching back to monthly salary shows ${await plSetting("plSalary").inputValue()}, expected the derived 38623`);
await plType("plSalary", 38500);

// The economic assumptions: a preset sets the three fields.
await pl.locator('button[data-preset="growth"]').click();
await tab.waitForTimeout(100);
const plGrowthFields = [await plSetting("plInflation").inputValue(), await plSetting("plGrowth").inputValue(), await plSetting("plReturn").inputValue()];
const plGrowthActive = await pl.locator('button[data-preset="growth"]').evaluate((b) => b.classList.contains("active"));
console.log(`pension lasts   : growth preset -> ${plGrowthFields.join("/")}, ribbon ${(await plRibbon()).map((item) => item.text).join(",")}`);
if (!plGrowthActive || plGrowthFields.join("/") !== "2/1,8/1,7") problems.push(`the growth preset reads active=${plGrowthActive}, fields ${plGrowthFields.join("/")}, expected active 2/1,8/1,7`);
await pl.locator('button[data-preset="standard"]').click();
await tab.waitForTimeout(100);
if (!sameList((await plRibbon()).map((item) => item.text), PL_RIBBON_DEFAULT)) problems.push("the standard preset did not restore the default short answer");

// Reset puts everything back; "Hämta från Prognos" copies the form's person once.
await plType("plBudget", 27300);
await pl.locator('[data-action="pl-reset"]').click();
await tab.waitForTimeout(100);
if (!sameList((await plRibbon()).map((item) => item.text), PL_RIBBON_DEFAULT) || (await plSetting("plBudgetChoice").inputValue()) !== "1") {
  problems.push("Återställ did not restore the default typfall and budget");
}
if ((await plSetting("plLifespan").inputValue()) !== "95") problems.push(`Återställ left the lifespan at ${await plSetting("plLifespan").inputValue()}, expected 95`);
if ((await plSetting("plSavingsKind").inputValue()) !== "0" || (await plSetting("plSavingsCapital").inputValue()) !== "0") problems.push("Återställ did not put the private savings back on the first kind with nothing in it");
await pl.locator('[data-action="pl-import-forecast"]').click();
await tab.waitForTimeout(100);
const plImported = [
  await plSetting("plBorn").inputValue(),
  await plSetting("plStartWork").inputValue(),
  await plSetting("plSalary").inputValue(),
  await plSetting("plRetirement").inputValue(),
  await plSetting("plScheme").inputValue(),
];
console.log(`pension lasts   : from Prognos ${plImported.join("/")}`);
if (plImported.join("/") !== "1959/23/38500/66/1") problems.push(`"Hämta från Prognos" gives ${plImported.join("/")}, expected 1959/23/38500/66/1`);
await pl.locator('[data-action="pl-reset"]').click();
await tab.waitForTimeout(100);
await plType("plLifespan", 90);

// The exports: the options table, the per-age series and the budgets as CSV (whole kronor, Swedish delimiter), and a real .xlsx.
const plOptionsActions = pl.locator('[data-role="pl-options-panel"] .panel-actions');
const plCsv = await download(plOptionsActions.getByRole("button", { name: "Ladda ner CSV" }));
const plCsvLines = plCsv.buffer.toString("utf8").replace(/^﻿/, "").split("\r\n");
console.log(`pension lasts   : CSV ${plCsv.filename}, ${plCsvLines.length} lines`);
if (plCsvLines.length !== 18 || !plCsvLines[0].startsWith("Utbetalningstid;Under utbetalningen, kr/mån efter skatt;Efter utbetalningen, kr/mån efter skatt;")) {
  problems.push(`the options CSV starts "${plCsvLines[0]}" and has ${plCsvLines.length} lines, expected the 18-line table`);
}
if (plCsvLines[1] !== "5 år;38997;16005;20597;-2395;Till 73 år;5;488494;Hela vägen till 90 år") problems.push(`the options CSV's first row reads "${plCsvLines[1]}"`);
if (plCsvLines[17] !== "Livsvarigt;22289;;3889;;Hela vägen till 90 år;22;0;Hela vägen till 90 år") problems.push(`the options CSV's last row reads "${plCsvLines[17]}"`);
const plXlsx = await download(plOptionsActions.getByRole("button", { name: "Ladda ner Excel" }));
if (!looksLikeXlsx(plXlsx.buffer)) problems.push(`the options' Excel download ("${plXlsx.filename}") isn't a real .xlsx file`);
const plSeries = await download(pl.locator('[data-role="pl-time"] .panel-actions').getByRole("button", { name: "Ladda ner CSV" }));
const plSeriesLines = plSeries.buffer.toString("utf8").replace(/^﻿/, "").split("\r\n");
if (plSeriesLines.length !== 23 || !plSeriesLines[1]?.startsWith("68;18400;35510;")) problems.push(`the per-age CSV has ${plSeriesLines.length} lines starting "${plSeriesLines[1]}", expected 23 lines from age 68 at 18400 then 35510`);
const plMatrixCsv = await download(pl.locator('[data-role="pl-matrix-panel"] .panel-actions').getByRole("button", { name: "Ladda ner CSV" }));
const plMatrixLines = plMatrixCsv.buffer.toString("utf8").replace(/^﻿/, "").split("\r\n");
if (plMatrixLines.length !== 18 || plMatrixLines[1] !== "5 år;Till 73 år;Till 73 år;Till 73 år;35510") problems.push(`the budgets CSV has ${plMatrixLines.length} lines, the first row "${plMatrixLines[1]}"`);

// The theme tokens reach the boxes: the lifelong chip is ink (black on light, white on dark), the 5-year one the darkest gold.
const plChip = async (payout) => pl.locator(`.pl-ribbon li[data-payout="${payout}"] .pa-chip`).first().evaluate((el) => getComputedStyle(el).backgroundColor);
const plChipsLight = [await plChip(0), await plChip(5)];
await tab.emulateMedia({ colorScheme: "dark" });
const plChipsDark = [await plChip(0), await plChip(5)];
await tab.emulateMedia({ colorScheme: "light" });
console.log(`pension lasts   : chips ${plChipsLight.join(" / ")} light, ${plChipsDark.join(" / ")} dark`);
if (plChipsLight.join("/") !== "rgb(0, 0, 0)/rgb(91, 68, 0)" || plChipsDark.join("/") !== "rgb(255, 255, 255)/rgb(255, 200, 59)") {
  problems.push(`the lifelong and 5-year chips are ${plChipsLight.join("/")} on light and ${plChipsDark.join("/")} on dark, expected black/#5b4400 and white/#ffc83b`);
}

// Swedish and English; the numbers do not move with the language.
const plHeadSv = (await pl.locator("section.panel > h2 span").first().textContent()) ?? "";
await tab.locator(".langs button", { hasText: "EN" }).click();
await tab.waitForTimeout(150);
const plHeadEn = (await pl.locator("section.panel > h2 span").first().textContent()) ?? "";
const plLabelEn = (await pl.locator(".pa-ribbon-label").first().textContent()) ?? "";
const plButtonEn = (await tab.locator('.screen-toggle .panel-btn[data-screen="lasts"]').textContent()) ?? "";
const plEnFive = await plCells(5);
console.log(`pension lasts   : "${plHeadSv}" / "${plHeadEn}", "${plLabelEn}", button "${plButtonEn}", 5 years "${plEnFive[4]}"`);
if (plHeadSv !== "Hur länge räcker pensionen?" || plHeadEn !== "How long does the pension last?" || plLabelEn !== "Payout period and how long the budget lasts" || plButtonEn !== "How long does the pension last?") {
  problems.push(`the tab's language switch reads "${plHeadSv}"/"${plHeadEn}"/"${plLabelEn}"/"${plButtonEn}"`);
}
if (plEnFive[1] !== "38,997" && plEnFive[1] !== "38997") problems.push(`in English the 5-year payout reads ${plEnFive[1]}, expected 38,997`);
const plLabelsEn = [await plLabel("plCapital"), await plLabel("plPrivSaving"), await plLabel("plSavingsCapital")];
const plGroupsEn = (await pl.locator(".pa-group").allTextContents()).map((text) => text.trim());
if (!sameList(plLabelsEn, ["Individualoccupationalpension,capital", "Privatepension,monthlysaving", "ISK/Endowmentinsurance/Depotinsurance,capital"]) || !plGroupsEn.includes("Private savings")) {
  problems.push(`in English the fields are labelled ${plLabelsEn.join(" | ")} and the groups ${plGroupsEn.join(" | ")}, expected Individual occupational pension, capital / Private pension, monthly saving / ISK / Endowment insurance / Depot insurance, capital and a "Private savings" group`);
}
if (!sameList((await plRibbon()).map((item) => item.text), ["To73", "To78", "To83", "To88", "Alltheway"])) {
  problems.push(`in English the boxes read ${(await plRibbon()).map((item) => item.text).join(",")}, expected To73,To78,To83,To88,Alltheway`);
}
await tab.locator(".langs button", { hasText: "SV" }).click();
await tab.waitForTimeout(150);

// A phone: nothing widens the page, and the boxes wrap instead of overflowing.
await tab.setViewportSize({ width: 375, height: 800 });
await tab.waitForTimeout(150);
const plOverflow = await tab.evaluate(() => document.documentElement.scrollWidth - document.documentElement.clientWidth);
const plRibbonRows = await pl.locator(".pl-ribbon li").evaluateAll((items) => new Set(items.map((li) => Math.round(li.getBoundingClientRect().top))).size);
const plTableScrolls = await pl.locator('[data-role="pl-options-panel"] .scroll').evaluate((box) => box.scrollWidth > box.clientWidth);
console.log(`pension lasts @375px: page overflow ${plOverflow}px, boxes in ${plRibbonRows} rows, table scrolls ${plTableScrolls}`);
if (plOverflow > 1) problems.push(`the Hur länge räcker pensionen? tab widens the page by ${plOverflow}px at 375px`);
const plPairsPhone = await plPairRows();
if (plPairsPhone.length !== 3 || !plPairsPhone.every((pair) => pair.sameRow && pair.side)) {
  problems.push(`at 375px the capital and monthly saving are not side by side on one row: ${JSON.stringify(plPairsPhone)}`);
}
if (plRibbonRows < 2 || !plTableScrolls) problems.push(`at 375px the answer boxes sit in ${plRibbonRows} row(s) and the options table scrolls ${plTableScrolls}, expected several rows and a scrolling table`);
await tab.setViewportSize({ width: 1280, height: 1000 });
await tab.waitForTimeout(100);

await tab.locator('.screen-toggle .panel-btn[data-screen="single"]').click();
await tab.waitForTimeout(50);

// ---- Hur mycket får jag? -------------------------------------------------------
//
// The lasts tab's person, agreement, savings and assumptions without the budget, answering what each
// payout pays a month: 5, 10, 15 and 20 years, and for life.

await tab.locator('.screen-toggle .panel-btn[data-screen="amount"]').click();
await tab.waitForSelector('[data-role="pension-amount-panel"] [data-role="pm-ribbon"]');
await tab.waitForTimeout(250);
const am = tab.locator('[data-role="pension-amount-panel"]');
const amRibbon = async () =>
  am.locator('[data-role="pm-ribbon"] li').evaluateAll((items) =>
    items.map((li) => ({ payout: li.dataset.payout, text: li.innerText.replace(/\s+/g, " ").trim() })),
  );
const amDuring = (items) => items.map((item) => Number(item.text.match(/(\d[\d\s]*) kr\/mån/)?.[1].replace(/\s/g, "")));
const amStart = await amRibbon();
const amHeadSv = ((await am.locator("section.panel > h2 span").first().textContent()) ?? "").trim();
const amBudgetFields = await am.locator('[data-setting="plBudget"], [data-setting="plBudgetChoice"]').count();
const amGroups = (await am.locator(".pa-group").allTextContents()).map((text) => text.trim());
console.log(`how much        : "${amHeadSv}", ${amStart.map((i) => i.payout).join("/")} -> ${amDuring(amStart).join("/")} kr/month, groups ${amGroups.join("/")}, budget fields ${amBudgetFields}`);
if (amHeadSv !== "Hur mycket får jag?") problems.push(`the tab is headed "${amHeadSv}", expected "Hur mycket får jag?"`);
if (!sameList(amStart.map((i) => i.payout), ["5", "10", "15", "20", "0"])) problems.push(`the answer boxes are for ${amStart.map((i) => i.payout).join(",")}, expected 5,10,15,20,0 (life)`);
if (!sameList(amDuring(amStart), [38997, 28977, 25242, 23294, 22131])) problems.push(`the answer boxes read ${amDuring(amStart).join(",")} kr/month, expected 38997,28977,25242,23294,22131 (the same runs as "Hur länge räcker pensionen?")`);
if (amBudgetFields !== 0 || amGroups.includes("Budget")) problems.push(`the tab has a budget (${amBudgetFields} fields, groups ${amGroups.join("/")}), expected none`);
if (!amGroups.includes("Person") || !amGroups.includes("Pension") || !amGroups.includes("Privat sparande") || !amGroups.includes("Antaganden")) {
  problems.push(`the tab's groups are ${amGroups.join("/")}, expected Person, Lön, Pension, Privat sparande, Antaganden, Visa`);
}
// A shorter payout pays more a month, and the temporary ones leave the public pension alone behind them.
const amTable = await am.locator('[data-role="pm-options-panel"] tbody tr').evaluateAll((trs) =>
  trs.map((tr) => [...tr.querySelectorAll("td")].map((td) => td.innerText.replace(/\s+/g, " ").trim())),
);
console.log(`how much        : table ${amTable.map((r) => r.join(" | ")).join(" ; ")}`);
if (amTable.length !== 6 || !amTable[0][0].startsWith("Bara allmän pension")) problems.push(`the amounts table has ${amTable.length} rows, expected the public pension alone and five payouts`);
const amAfter = amTable.slice(1, 5).map((r) => Number(r[2].replace(/\s/g, "")));
if (!amAfter.every((v, i) => i === 0 || v < amAfter[i - 1] + 1) || amTable[5][2] !== "–") {
  problems.push(`what is left after the temporary payouts reads ${amAfter.join(",")} and for life "${amTable[5][2]}"`);
}

// Capital of one's own raises every payout, the 5-year one most.
await am.locator('[data-setting="plCapital"]').fill("300000");
await am.locator('[data-setting="plCapital"]').dispatchEvent("change");
await am.locator('[data-setting="plCapital"]').evaluate((el) => el.blur());
await tab.waitForTimeout(150);
const amWith = amDuring(await amRibbon());
console.log(`how much        : with an individual pension of 300 000 -> ${amWith.join("/")}`);
if (!amWith.every((v, i) => v > amDuring(amStart)[i])) problems.push(`an individual pension of 300 000 did not raise every payout: ${amWith.join(",")} against ${amDuring(amStart).join(",")}`);
if (amWith[0] - amDuring(amStart)[0] <= amWith[4] - amDuring(amStart)[4]) problems.push("the 5-year payout did not gain more a month than the lifelong one from the same capital");

// The downloads, and a phone.
const amCsv = await download(am.locator('.export-btn[data-file="belopp-per-manad.csv"]'));
const amSeries = await download(am.locator('.export-btn[data-file="pension-per-ar.csv"]'));
const amXlsx = await download(am.locator('.export-btn[data-file="belopp-per-manad.xlsx"]'));
const amCsvLines = amCsv.buffer.toString("utf8").replace(/^\uFEFF/, "").split("\r\n");
console.log(`how much        : ${amCsv.filename} ${amCsvLines.length} lines, ${amSeries.filename}, ${amXlsx.filename}`);
if (amCsvLines.length !== 6 || !looksLikeXlsx(amXlsx.buffer) || amSeries.buffer.length < 100) problems.push(`the tab's downloads are ${amCsvLines.length} lines / xlsx ${looksLikeXlsx(amXlsx.buffer)} / series ${amSeries.buffer.length} bytes`);

await tab.locator(".langs button", { hasText: "EN" }).click();
await tab.waitForTimeout(150);
const amHeadEn = ((await am.locator("section.panel > h2 span").first().textContent()) ?? "").trim();
const amButtonEn = ((await tab.locator('.screen-toggle .panel-btn[data-screen="amount"]').textContent()) ?? "").trim();
await tab.locator(".langs button", { hasText: "SV" }).click();
await tab.waitForTimeout(150);
if (amHeadEn !== "How much do I get?" || amButtonEn !== "How much do I get?") problems.push(`in English the tab reads "${amHeadEn}" / "${amButtonEn}"`);
await tab.setViewportSize({ width: 375, height: 800 });
await tab.waitForTimeout(150);
const amOverflow = await tab.evaluate(() => document.documentElement.scrollWidth - document.documentElement.clientWidth);
if (amOverflow > 1) problems.push(`the Hur mycket får jag? tab widens the page by ${amOverflow}px at 375px`);
await tab.setViewportSize({ width: 1280, height: 1000 });
await tab.waitForTimeout(100);
await tab.locator('.screen-toggle .panel-btn[data-screen="single"]').click();
await tab.waitForTimeout(50);

// ---- Om Pensionslyft -----------------------------------------------------------
//
// The last entry of the main menu: what the site is for, the disclaimer and how to reach it.

const aboutTexts = async () => ({
  head: ((await tab.locator('[data-role="about-panel"] section.panel > h2 span').first().textContent()) ?? "").trim(),
  headings: (await tab.locator('[data-role="about-panel"] h3').allTextContents()).map((text) => text.trim()),
  body: (await tab.locator('[data-role="about-panel"]').textContent()) ?? "",
  link: await tab.locator('[data-role="about-panel"] a').last().getAttribute("href"),
  sourceLink: await tab.locator('[data-role="about-panel"] a').first().getAttribute("href"),
});
const tabNamesSv = (await tab.locator(".screen-toggle .panel-btn").allTextContents()).map((text) => text.trim());
await tab.locator('.screen-toggle .panel-btn[data-screen="about"]').click();
await tab.waitForTimeout(100);
const aboutSv = await aboutTexts();
await tab.locator(".langs button", { hasText: "EN" }).click();
await tab.waitForTimeout(100);
const aboutEn = await aboutTexts();
const tabNameEn = ((await tab.locator('.screen-toggle .panel-btn[data-screen="about"]').textContent()) ?? "").trim();
await tab.locator(".langs button", { hasText: "SV" }).click();
await tab.waitForTimeout(100);
console.log(`about           : menu ${tabNamesSv.join(" / ")}; "${aboutSv.head}" / "${aboutEn.head}", ${aboutSv.headings.join("/")}, link ${aboutSv.link}`);
if (!sameList(tabNamesSv, ["Prognos", "Jämför scenarier", "Mikrosim", "Hur maxa pensionen?", "Hur länge räcker pensionen?", "Hur mycket får jag?", "Om Pensionslyft"])) {
  problems.push(`the main menu reads ${tabNamesSv.join(" / ")}, expected Prognos / Jämför scenarier / Mikrosim / Hur maxa pensionen? / Hur länge räcker pensionen? / Hur mycket får jag? / Om Pensionslyft`);
}
if (aboutSv.head !== "Om Pensionslyft" || aboutEn.head !== "About Pensionslyft" || tabNameEn !== "About Pensionslyft") {
  problems.push(`the about page is headed "${aboutSv.head}" / "${aboutEn.head}" and its button reads "${tabNameEn}" in English`);
}
if (!sameList(aboutSv.headings, ["Syfte", "Källa och bearbetning", "Friskrivning", "Kontakt"]) || !sameList(aboutEn.headings, ["Purpose", "Source and adaptation", "Disclaimer", "Contact"])) {
  problems.push(`the about page's sections are ${aboutSv.headings.join("/")} and ${aboutEn.headings.join("/")}`);
}
for (const [what, ok] of [
  ["the purpose", aboutSv.body.includes("förmedla kunskap, inspirera och lyfta pensionsfrågan") && aboutSv.body.includes("välinformerade beslut som passar dig och din situation")],
  ["the 'not financial advice' disclaimer", aboutSv.body.includes("Inte finansiell rådgivning") && /not financial advice/i.test(aboutEn.body)],
  ["the source: Pensionsmyndigheten, with a link and the model's version", aboutSv.sourceLink?.startsWith("https://www.pensionsmyndigheten.se/") === true && aboutSv.body.includes("Källa: Pensionsmyndigheten") && aboutSv.body.includes("ver. 4.8") && /Source: Pensionsmyndigheten/.test(aboutEn.body) && aboutSv.body.includes("bearbetad version")],
  ["the contact: the website", aboutSv.link === "https://pensionslyft.se" && aboutSv.body.includes("pensionslyft.se")],
]) {
  if (!ok) problems.push(`the about page lacks ${what}`);
}
await tab.locator('.screen-toggle .panel-btn[data-screen="single"]').click();
await tab.waitForTimeout(50);

// ---- The first press ---------------------------------------------------------
//
// A number is applied when its field loses focus, and the press on a button is what takes the focus
// from it. The page used to be redrawn between the button's mousedown and its mouseup, so the button
// the pointer came up on was not the one it went down on, the browser dropped the click, and the first
// press did nothing. Every press here is a real one -- the mouse moved, pressed and released through
// the browser's own input, with none of the retrying `locator.click()` does for a button that was
// replaced under it -- made while a number has been typed and its field not left.

const fp = await context.newPage();
fp.on("pageerror", (error) => errors.push(String(error)));
fp.on("console", (message) => {
  if (message.type() === "error") errors.push(message.text());
});
const fpFiles = [];
fp.on("download", (saved) => fpFiles.push(saved));
await fp.goto(pathToFileURL(page).href);
await fp.waitForSelector(".table1 tbody tr[data-key]", { timeout: 15000 });

const fpPresses = [];
const fpExpect = (what, ok) => fpPresses.push({ what, ok });
const fpPress = async (page, locator, how = "mouse") => {
  await locator.scrollIntoViewIfNeeded();
  const box = await locator.boundingBox();
  if (box === null) throw new Error("the button has no box to press");
  const [x, y] = [box.x + box.width / 2, box.y + box.height / 2];
  if (how === "touch") return page.touchscreen.tap(x, y);
  await page.mouse.move(x, y);
  await page.mouse.down();
  await page.mouse.up();
};
const fpTyped = async (field, text) => {
  await field.click();
  await field.press("Control+A");
  await field.pressSequentially(text);
};
const fpActive = async (locator) => ((await locator.getAttribute("class")) ?? "").split(/\s+/).includes("active");
const fpSalary = fp.locator("form.inputs input[type=number]").nth(3);
const fpBorn = fp.locator("form.inputs input[type=number]").first();
const fpScreen = (screen) => fp.locator(`.screen-toggle .panel-btn[data-screen="${screen}"]`);

/** Press a download button and read back what it wrote: how many files, and the first. */
const fpWritten = async (page, button, how = "mouse") => {
  const from = fpFiles.length;
  await fpPress(page, button, how);
  await page.waitForTimeout(350);
  const written = fpFiles.slice(from);
  return { count: written.length, buffer: written.length > 0 ? readFileSync(await written[0].path()) : Buffer.alloc(0) };
};
/**
 * The three presses of one download button: before anything was typed, the first press after a number
 * was typed, and a press once that number has been applied and left. The first has to write the file
 * the third writes -- the one with the new number in it, not the one the page showed when the
 * button was pressed -- once, and not the file the first one wrote.
 */
const fpFresh = async (what, button, field, text, leave) => {
  const baseline = await fpWritten(fp, button());
  await fpTyped(field, text);
  const first = await fpWritten(fp, button());
  await leave();
  const settled = await fpWritten(fp, button());
  fpExpect(`${what}: one file`, first.count === 1);
  fpExpect(`${what}: not the file from before the number was typed`, first.buffer.length > 0 && !first.buffer.equals(baseline.buffer));
  fpExpect(`${what}: the file a press after the number is applied writes`, first.buffer.equals(settled.buffer));
};
const fpLeave = () => fp.locator("h1").click();

// Prognos: the tabs, the mode, the figure's and the table's own toggles, the language.
await fpTyped(fpSalary, "512345");
await fpPress(fp, fpScreen("retirement"));
await fp.waitForTimeout(150);
fpExpect("Prognos: the 'Hur maxa pensionen?' tab", await fpActive(fpScreen("retirement")));
await fpPress(fp, fpScreen("single"));
await fp.waitForTimeout(150);

await fpTyped(fpSalary, "498765");
await fpPress(fp, fpScreen("about"));
await fp.waitForTimeout(150);
fpExpect("Prognos: the 'Om Pensionslyft' entry", await fpActive(fpScreen("about")));
await fpPress(fp, fpScreen("single"));
await fp.waitForTimeout(150);

for (const [what, press, restore] of [
  ["the Avancerat toggle", '.mode-toggle .panel-btn[data-mode="advanced"]', '.mode-toggle .panel-btn[data-mode="normal"]'],
  ["Figur 2's 'Före skatt' toggle", '.panel-toggle .panel-btn[data-basis="before-tax"]', '.panel-toggle .panel-btn[data-basis="after-tax"]'],
]) {
  await fpTyped(fpSalary, "498765");
  await fpPress(fp, fp.locator(press));
  await fp.waitForTimeout(150);
  fpExpect(`Prognos: ${what}`, await fpActive(fp.locator(press)));
  await fpTyped(fpSalary, "512345");
  await fpPress(fp, fp.locator(restore));
  await fp.waitForTimeout(150);
  fpExpect(`Prognos: and back again`, await fpActive(fp.locator(restore)));
}
await fpTyped(fpSalary, "498765");
await fpPress(fp, fp.locator(".panel-toggle .panel-btn", { hasText: "Årsvis" }));
await fp.waitForTimeout(150);
fpExpect("Prognos: Table 2's 'Årsvis' toggle", await fpActive(fp.locator(".panel-toggle .panel-btn", { hasText: "Årsvis" })));
await fpTyped(fpSalary, "512345");
await fpPress(fp, fp.locator(".panel-toggle .panel-btn", { hasText: "Månadsvis" }));
await fp.waitForTimeout(150);

await fpFresh("Prognos: Table 1's CSV", () => fp.locator('.export-btn[data-file="tabell1.csv"]'), fpSalary, "530000", fpLeave);
await fpFresh("Prognos: Table 2's CSV", () => fp.locator('.export-btn[data-file="tabell2.csv"]'), fpSalary, "470000", fpLeave);

await fpTyped(fpSalary, "498765");
await fpPress(fp, fp.locator(".langs button", { hasText: "EN" }));
await fp.waitForTimeout(150);
fpExpect("Prognos: the language chip 'EN'", (await fp.evaluate(() => document.documentElement.lang)) === "en");
await fpTyped(fpSalary, "512345");
await fpPress(fp, fp.locator(".langs button", { hasText: "SV" }));
await fp.waitForTimeout(150);
fpExpect("Prognos: and back to 'SV'", (await fp.evaluate(() => document.documentElement.lang)) === "sv");

// Hur länge räcker pensionen?
await fpPress(fp, fpScreen("lasts"));
await fp.waitForSelector('[data-setting="plCapital"]');
await fp.waitForTimeout(250);
const fpPlCapital = fp.locator('[data-setting="plCapital"]');
await fpFresh("Lasts: the CSV", () => fp.locator('.export-btn[data-file="pension-per-ar.csv"]'), fpPlCapital, "300000", fpLeave);
await fpTyped(fpPlCapital, "310000");
const fpPlExcel = await fpWritten(fp, fp.locator('.export-btn[data-file="pension-per-ar.xlsx"]'));
fpExpect("Lasts: the Excel file, once", fpPlExcel.count === 1 && looksLikeXlsx(fpPlExcel.buffer));
await fpTyped(fpPlCapital, "123456");
await fpPress(fp, fp.locator("[data-preset]").first());
await fp.waitForTimeout(250);
fpExpect("Lasts: a button among the inputs keeps what was typed", (await fpPlCapital.inputValue()) === "123456");

// Hur maxa pensionen?, one agreement and all of them.
await fpPress(fp, fpScreen("retirement"));
await fp.waitForSelector('[data-setting="paCapital"]');
await fp.waitForTimeout(250);
const fpPaCapital = fp.locator('[data-setting="paCapital"]');
await fpFresh("Age: the CSV", () => fp.locator('.export-btn[data-file="pension-per-ar.csv"]'), fpPaCapital, "300000", fpLeave);
await fpPress(fp, fp.locator('button[data-view="compare"]'));
await fp.waitForTimeout(400);
await fpFresh("Age, the agreements: the CSV", () => fp.locator('.export-btn[data-file="avtalsomraden.csv"]'), fpPaCapital, "320000", fpLeave);
await fpTyped(fpPaCapital, "330000");
await fpPress(fp, fp.locator('button[data-view="single"]'));
await fp.waitForTimeout(300);
fpExpect("Age: the switch between one agreement and all", await fpActive(fp.locator('button[data-view="single"]')));

// A number out of range is shown clamped, and the press still lands. A field that was only clicked
// into is left as an ordinary press leaves it.
await fpPress(fp, fpScreen("single"));
await fp.waitForTimeout(150);
await fpTyped(fpBorn, "3000");
await fpPress(fp, fpScreen("retirement"));
await fp.waitForTimeout(250);
fpExpect("a clamped entry: the press lands", await fpActive(fpScreen("retirement")));
fpExpect("a clamped entry: it reads the largest year", (await fpBorn.inputValue()) === (await fpBorn.getAttribute("max")));
await fpPress(fp, fpScreen("single"));
await fp.waitForTimeout(150);
await fpSalary.click();
await fpPress(fp, fpScreen("mikrosim"));
await fp.waitForTimeout(250);
fpExpect("a field with nothing typed: the press lands", await fpActive(fpScreen("mikrosim")));
await fp.close();

// A finger: the same press, as a tap, on a phone.
const touchContext = await browser.newContext({ viewport: { width: 420, height: 900 }, hasTouch: true, isMobile: true, acceptDownloads: true });
await touchContext.route("**/*", (route) => {
  const url = route.request().url();
  if (url.startsWith("file://") || url.startsWith("blob:")) return route.continue();
  outbound.push(url);
  return route.abort();
});
const finger = await touchContext.newPage();
finger.on("pageerror", (error) => errors.push(String(error)));
finger.on("download", (saved) => fpFiles.push(saved));
await finger.goto(pathToFileURL(page).href);
await finger.waitForSelector(".table1 tbody tr[data-key]", { timeout: 15000 });
const fingerSalary = finger.locator("form.inputs input[type=number]").nth(3);
const fingerTyped = async (text) => {
  await fingerSalary.tap();
  await fingerSalary.press("Control+A");
  await fingerSalary.pressSequentially(text);
};
await fingerTyped("512345");
const fingerCsv = await fpWritten(finger, finger.locator('.export-btn[data-file="tabell1.csv"]'), "touch");
fpExpect("a tap on Table 1's CSV, once", fingerCsv.count === 1 && fingerCsv.buffer.length > 100);
await fingerTyped("498765");
await fpPress(finger, finger.locator('.screen-toggle .panel-btn[data-screen="retirement"]'), "touch");
await finger.waitForTimeout(250);
fpExpect("a tap on the 'Hur maxa pensionen?' tab", await fpActive(finger.locator('.screen-toggle .panel-btn[data-screen="retirement"]')));
await touchContext.close();

const fpFailed = fpPresses.filter((press) => !press.ok);
console.log(`first press     : ${fpPresses.length - fpFailed.length}/${fpPresses.length} landed first time, with the number that was just typed${fpFailed.length > 0 ? ` -- not: ${fpFailed.map((press) => press.what).join("; ")}` : ""}`);
for (const press of fpFailed) problems.push(`a press right after typing a number did not do what it should: ${press.what}`);

// ---- Provenance and paper ------------------------------------------------
//
// The page computes a pension forecast and looks authoritative, so it has to
// say whose model it is; and people print a forecast to take to a meeting, so
// paper is a real output rather than an afterthought.

const pageTitle = await tab.locator("h1").first().textContent();
const documentTitle = await tab.title();
console.log(`page title      : h1 "${pageTitle}", document title "${documentTitle}"`);
if (pageTitle !== "Typfallsmodellen web version") {
  problems.push(`the page's own h1 reads "${pageTitle}", expected "Typfallsmodellen web version"`);
}
if (documentTitle !== "Typfallsmodellen web version") {
  problems.push(`the document title reads "${documentTitle}", expected "Typfallsmodellen web version"`);
}

const notice = tab.locator('[data-role="disclaimer"]');
const noticeText = (await notice.count()) > 0 ? ((await notice.textContent()) ?? "") : "";
if (!/inofficiell|unofficial/i.test(noticeText)) {
  problems.push("the page does not say it is an unofficial version");
}
if (!noticeText.includes("Pensionsmyndigheten")) {
  problems.push("the disclaimer does not name Pensionsmyndigheten");
}
const noticeLink = await tab.locator('[data-role="disclaimer"] a').first().getAttribute("href");
if (!noticeText.includes("Källa:") || !noticeText.includes("ver. 4.8") || !noticeText.includes("bearbetad version") || !noticeLink?.startsWith("https://www.pensionsmyndigheten.se/")) {
  problems.push(`the notice does not name Pensionsmyndigheten as the source with a link and the version, nor say it is an adapted version (link ${noticeLink})`);
}
if (noticeText.includes("typfallsmodellen@pensionsmyndigheten.se")) {
  problems.push("the disclaimer still names the agency's address for model questions, expected it removed");
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
