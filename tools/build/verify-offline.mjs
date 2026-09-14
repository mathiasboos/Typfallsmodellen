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

/** How many ages Figur 2 and the disposable income chart cover: `A2:A22`. */
const WINDOW = 21;

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
if (figures !== 3) problems.push(`${figures} figures rendered, expected 3`);
if (stacked !== WINDOW) problems.push(`Figur 2 has ${stacked} columns, expected ${WINDOW}`);

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
