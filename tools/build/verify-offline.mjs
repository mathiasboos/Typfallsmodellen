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
const outbound = [];
await context.route("**/*", (route) => {
  const url = route.request().url();
  if (url.startsWith("file://")) return route.continue();
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

const problems = [];
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
// Six settings groups plus the salary-path grid.
if (advGroups !== 7) {
  problems.push(`advanced mode shows ${advGroups} groups, expected 7`);
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
const ownIncome = await setting("ownIncome");
await ownIncome.check();
await tab.waitForTimeout(80);
const gridRows = await tab.locator(".adv-grid tbody tr").count();
console.log(`salary grid rows: ${gridRows}`);
// Ages 15 up to the last worked one; retirement is at 66 for this typfall.
if (gridRows < 40 || gridRows > 60) {
  problems.push(`the salary grid has ${gridRows} rows, expected about fifty`);
}
const firstIncome = await tab
  .locator(".adv-grid tbody tr")
  .last()
  .locator("input")
  .first()
  .inputValue();
if (!(Number(firstIncome) > 0)) {
  problems.push(`the salary grid filled the last worked year with ${firstIncome}, expected an income`);
}

const grossBefore = await table2Cell(table2Rows - 1, grossCol);
// Zero the last ten worked years -- a decade of leave.
for (let i = 0; i < 10; i += 1) {
  const cells = tab.locator(".adv-grid tbody tr").nth(gridRows - 1 - i).locator("input");
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

// Back to normal mode for the screenshots, and to leave the page as found.
await tab.locator('[data-action="reset-advanced"]').click();
await tab.locator('.mode-toggle .panel-btn[data-mode="normal"]').click();
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

if (shots) {
  mkdirSync(shots, { recursive: true });
  await tab.screenshot({ path: join(shots, "sv.png"), fullPage: true });
  await tab.getByRole("button", { name: "EN" }).click();
  await tab.waitForTimeout(150);
  await tab.screenshot({ path: join(shots, "en.png"), fullPage: true });
  await tab.emulateMedia({ colorScheme: "dark" });
  await tab.screenshot({ path: join(shots, "dark.png"), fullPage: true });
  console.log(`screenshots     : ${shots}`);
}

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
