# Typfallsmodellen på webben

[![CI](https://github.com/mathiasboos/Typfallsmodellen/actions/workflows/ci.yml/badge.svg)](https://github.com/mathiasboos/Typfallsmodellen/actions/workflows/ci.yml)

A web version of Pensionsmyndigheten's **Typfallsmodellen** — the model that computes the full
Swedish pension outcome for a hypothetical individual (a *typfall*): inkomstpension,
premiepension, tilläggspension, garantipension, inkomstpensionstillägg, occupational pension
across eight collective agreements, private saving, income tax and bostadstillägg.

The original is a Windows-only Excel workbook with ~20 000 lines of VBA, published at
[pensionsmyndigheten.se](https://www.pensionsmyndigheten.se/statistik-och-rapporter/pensionsmodellen/typfallsmodellen)
and updated roughly once a year. This project ports it to a static website that runs anywhere,
computes entirely in the browser, and can absorb the yearly data release through a script.

> **Unofficial.** Not affiliated with or endorsed by Pensionsmyndigheten. The model, its data and
> its user manual are theirs.
>
> **Source.** Based on Pensionsmyndigheten's Typfallsmodellen (ver. 4.8),
> [pensionsmyndigheten.se](https://www.pensionsmyndigheten.se/statistik-och-rapporter/pensionsmodellen/typfallsmodellen).
> The agency allows the tool to be spread and modified as long as it is named as the source. This is an
> adapted version: the model is rewritten for the web, and the views "Portföljvärde", "Jämför scenarier", "Mikrosim",
> "När ta ut pensionen?", "Hur länge räcker pensionen?" and "Hur mycket får jag?", and the private saving,
> are additions that are not part of the agency's model.

## Layout

```
source/            the published .xlsb this build is derived from
tools/extract/     Python: .xlsb -> packages/data (run once a year)
tools/transpile/   Python: the 32 tax-year functions, mechanically translated
tools/build/       embeds the mortality grid, folds the build into one HTML file, verifies it
packages/data/     generated data: economic series, mortality, tax, i18n, content
packages/engine/   TypeScript port of the VBA calculation engine; no UI dependencies
                   run(input, context, { deaths }) -> TypfallResult
apps/web/          the website
reference/         VBA source dump and regression fixtures; never shipped
docs/              architecture, projection rules, yearly-update runbook, VBA mapping
```

## Run it

```bash
npm install
npm run build -w @typfallsmodellen/web
```

`npm install` also writes `apps/web/src/generated/mortalityRisks.ts`, through the root `prepare`
script — the death-probability grid packed into a module, since a browser will not `fetch` it over
`file://`. It is derived from `packages/data` rather than committed, so a fresh clone has to make it
before anything will typecheck.

That writes **`apps/web/dist/typfallsmodellen.html`** — one self-contained file, about 850 kB.
The same file is published two ways: at
[mathiasboos.github.io/Typfallsmodellen](https://mathiasboos.github.io/Typfallsmodellen), and as an
asset on each [release](https://github.com/mathiasboos/Typfallsmodellen/releases), which is how the
original is distributed — something you save and keep.
Open it in any browser: double-click it, email it, put it on a stick. No installation, no server,
no network. Everything is computed in the browser and nothing is sent anywhere, which is not a
privacy claim to take on trust — `npm run verify:offline` opens the built file in Chromium with
every request that is not the file itself blocked, and fails if the page asks for anything or
disagrees with the engine. (That check needs `npm i -D playwright`; nothing else here does.)

`npm run dev -w @typfallsmodellen/web` serves it with hot reload while working on it.

## Status

| Phase | | |
|---|---|---|
| 0. Extraction pipeline | done | `.xlsb` → committed, diffable data |
| 1. Engine core | done | the whole model runs: `run(input, context)` returns Table 1, Table 2, the life-income sums and the per-age matrix |
| 2. Golden-file harness vs. Excel | done | 299 typfall out of the real model, twelve output columns, every cell exact — `npm run compare` |
| 3. Normal-mode website | done | the Start sheet: eight typed input cells, Table 1's four columns, Table 2, Figur 1, Figur 2 and the disposable income chart, Swedish and English — one offline HTML file |
| 4. Advanced mode | done | the workbook's second mode: thirty of the `Adv_settings` in the manual's own groups (including partial withdrawal of the public pension and the ITP 1/SAF-LO flexpension premium), the Indata_lista wage path as an editable grid, and the PGB sheet's pension-qualifying amounts as a compact add-entry form (pick sickness/activity compensation, conscription, study or a child, fill in that type's own fields) plus a summary table that only ever shows the years and categories with data |
| 5. Polish, CI, deploy | done | every check above runs on each push; the site publishes to GitHub Pages and each tag attaches the file to a release |
| 6. Compare scenarios *(not from the workbook)* | done | a second top-level view: the baseline typfall alongside up to three variants, each with its own birth year, salary, retirement age, start-of-work age and occupational pension, lined up in one shared comparison table and an overlay chart |
| 7. Mikrosim *(ported from the workbook's own Mikrosim sheet)* | done | a third top-level view: a batch of independent typfall in two linked tables (Inputs, then Results below it), each row's twelve output columns computed live as its own inputs are edited, added or imported, add/remove rows, one-click import from Forecast's own baseline or every Compare-scenarios scenario, CSV import/export, and a stacked column chart below breaking each row's pension into its seven components, with a black Disponibel inkomst line overlaid on top |
| 8. When to take out the pension *(not from the workbook)* | done | a fourth top-level view that answers *which retirement age pays out the most over a lifetime?*: seven ages around a reference age (the riktålder unless moved) run as one typfall each, summed over lifespans of 80-95 years (or any range), for a typfall of its own (birth year, start of work, monthly or final salary, one of the eight agreements, an individual occupational pension, a private pension and private savings -- the same sections and the same arithmetic as in the next tab, each saving going on until the alternative's own retirement age -- three economic assumptions, before or after tax) -- a short answer with the winner per lifespan, three line charts, the lifetime totals as a table, the break-even lifespans, and CSV/Excel downloads (the per-age table switches between the monthly pension and the total paid out, and its download follows); a table and a chart of the best age per lifespan with what it pays a month and in all; the line charts of the three question-mark tabs let each line be switched off and on from the legend; a second view runs the same person under every collective agreement side by side (best age per agreement and lifespan, total and added-over-no-agreement, with the agreements a cohort does not normally belong to marked as "what if") |
| 9. How long does the pension last? *(not from the workbook)* | done | a fifth top-level view that sets a pension against a budget: the public pension is paid for life, but the occupational and individual occupational pension can often be drawn over 5 to 20 years or for life, so the tab runs the typfall once per payout period (the workbook's own temporary-withdrawal settings, seventeen runs) and shows how long the budget lasts under each -- month by month, and added up if the surplus is saved -- for a typfall of its own (birth year, start of work, retirement age, estimated lifespan, monthly or final salary, one of the eight agreements, an individual occupational pension, a private pension (a pension insurance or an IPS) and private savings -- an ISK / kapitalförsäkring / depåförsäkring or a depå / VP-konto / fondkonto, the latter with its capital typed after capital gains tax -- each as capital and monthly saving, three economic assumptions, before or after tax) against minPension's three budgets (Baspensionär, Medelpensionär, Pluspensionär -- the amounts the site's Pensionsöversikt uses) or one's own: a short answer with the recommended payout, two line charts, every payout in a table, every budget against every payout, and CSV/Excel downloads; before tax the budget is grossed up with the model's own tax, age by age, and a tax-free private-savings payout the same way; its Pension and Privat sparande sections are built by one module that "Best retirement age" uses too |
| 10. How much do I get? *(not from the workbook)* | done | a sixth top-level view, the fifth's twin without the budget (`createPensionLastsPanel(…, "amount")`: same inputs, same runs): what the pension pays a month if the occupational pension, the individual and private pension and the private savings are drawn over 5, 10, 15 or 20 years or for life, during the payout and after it, per age, and added up to the lifespan; CSV/Excel |
| 11. About Pensionslyft *(not from the workbook)* | done | the last entry of the main menu (`about.ts`): what the site is for, a disclaimer (not financial advice, forecasts under the reader's own assumptions, unofficial, own responsibility) and the website as contact; static text in both languages. The fourth view is named "Bästa pensionsålder?" |
| 12. Portfolio value *(not from the workbook)* | done | a second entry in the main menu, after the forecast, with the same tables and charts as the "Portfolio Value" tab of the site's Pensionsplaneraren: the pension capital behind Prognos's own run (the form and Avancerat, no inputs of its own), account by account -- income pension, premium pension, occupational pension, private pension saving, private saving -- year by year: opening balance, contribution, withdrawal, return, yield tax, management fee, closing balance. An account picker or every account at once, Totalt or Skattekorrigerat (the accounts taxed when paid out, after the average tax on the pension at retirement), a stacked balance chart (an area for one account) and a chart of the year's in- and outflows, the all-accounts table and one per account, CSV/Excel. The engine's `result.ledger` carries the figures: each pot the model keeps, with the withdrawals Table 2 pays; the income pension, which the model turns into a pension for life at retirement, is valued from then on as next year's pension times the divisor at next year's age |

### Verified so far

| What | Against | Result |
|---|---|---|
| Economic projection | the workbook's cached values for 20 series across ~190 years | exact |
| Annuity factors (delningstal, arvsvinstfaktorer) | the 114 345 values the workbook computed into `mortality!P:Z` | exact |
| Pension-qualifying income and contributions | the `Brutto` sheet, which calls those VBA functions as worksheet UDFs | exact |
| Income pension accumulation | the same sheet, term by term — gains, indexation, cost, balance | exact |
| The earning phase of the main loop, end to end | the `Brutto` sheet's per-age trace — PGI, PGB and all three contributions | exact |
| **Every Table 1 figure, end to end** | **299 typfall the real model computed, all twelve output columns** | **3 588 of 3 588 cells exact** |
| The built HTML file, opened from disk with the network cut | all four Table 1 columns read back out of Chromium, against the engine's own snapshot | no outbound requests, no page errors, values exact — `npm run verify:offline` |
| Advanced mode, driven in the same browser | a municipal rate, a rent, ten zeroed salary years, and a manual PGB entry, each read back out of the rendered tables | every setting moves the model, and `Använd normala inställningar` puts it back |
| Compare scenarios, driven in the same browser | a scenario's own birth year, salary and start-of-work age each changed well away from the baseline's, the add/remove scenario limits, and the comparison table's own replacement-rate row | only that scenario's own column in the table and line in the chart move, the baseline stays put, the card/column/line count stays between one and four, and no replacement-rate cell ever reaches 100% |
| Table 2's header tooltips and Ordlista, driven in the same browser | the "Statlig skatt" header's own `<abbr title>`, and Ordlista's term count and language-switched summary | names the 20% threshold and the public-service fee; 57 terms present, staying Swedish across a language switch that still relabels the disclosure itself |
| CSV and Excel downloads, driven in the same browser | Table 1's, Table 2's and the comparison table's own download buttons, each actually clicked | a real file every time: the CSV's own content matches that table's data, the `.xlsx`'s first bytes are the ZIP signature every real OOXML package starts with |
| The built file's own script tag | a static check of the file itself, no browser involved | a classic `<script>`, not `type="module"` (which Mobile Safari can refuse to run at all over `file://`, seen as a black screen on an iPhone), placed after `#app` in the document |
| The PGB add-entry form and its sparse summary table, driven in the same browser | adding a sickness, study, conscription and child entry one at a time (including two in the same year, and a too-short conscription period), removing each one, and the whole panel at 375px | the summary table starts empty and only ever shows the years and categories that have data; each add/remove moves the pension and the table together; a child's birth year raises `PgbBreakdownYear.barn`, the field `pgbBarn` (new in the engine) exists to expose; the page never grows a horizontal scrollbar at 375px |
| Conscription and study's own PGB arithmetic | 15 unit tests pinning the day-split across one, two and three calendar years by hand, and the browser reading back a date range's own PGB värnplikt figure and a semester count's own PGB studier figure, each its own column in the summary table once it has data | exact day counts and kronor at every split; a 1998 conscription period and a 2005 semester both raise the pension on their own, distinct from each other and from a typed sickness/activity amount |
| Partial withdrawal of the public pension, driven in the same browser | age 67's own Lön and pension columns in Table 2, before and after setting a 50% withdrawal share and a final age of 70 | 0 kr salary and a full pension become a part-time salary alongside almost exactly half the full pension, and age 70 returns to 0 kr salary and a pension above the full-at-66 figure |
| Mikrosim, driven in the same browser | adding and removing rows, editing one row's salary and scheme with no button anywhere to click, importing a row from Forecast's own baseline and one per Compare-scenarios scenario, CSV export and re-importing that same file, a file missing a required column, a row with an invalid scheme next to a valid one, and the stacked chart (plus its Disponibel inkomst overlay line) below the two tables | each row's own output columns (a separate Results table below the Inputs one) fill in immediately from its own inputs only, with no effect on any other row's; the Forecast/Compare-scenarios imports each add the expected number of rows with each row's own fields intact; a bad file is refused by name with the table left untouched; an invalid row is flagged and left blank while the rest still compute, and fixing its own field recomputes it on the spot; the chart's own bar count tracks the table's row count, a plain edit alone moves the chart's own bar values, not just its count, and the overlay line (with its own legend entry) is drawn regardless of the 10-row cap on the Bruttopension labels |
| When to take out the pension, driven in the same browser | the static report's own typfall (born 1980, work from 25, 38 500 kr/month) read back out of the matrix and the short answer, then a chosen agreement, both economic presets, after tax, the reference age moved/clamped/re-locked, a typed final salary, the linked lifespan range, reset, the one-time import from Prognos, the CSV/Excel downloads, both languages, the hover and the arrow keys, and 375px; the individual and private pension, an ISK and a depot as the same numbers the unit tests give (and the same markup, field by field, as in "How long does the pension last?"), each kind of pension in its own column and in the short answer's figure; then the agreement comparison: the switch, the cohort's default picks, the best-age grid and the table to the krona, the lines and their colours in both themes, picking, the retirement age, the lifespan to read at, the chart toggle, after tax, the cohort swapping the defaults, both downloads and 375px | the totals match the report to the krona (e.g. 2 972 916 kr for living to 80 with a start at 65), the winners run 65, 66, 68 then 71 as the lifespan grows, an agreement adds exactly its occupational pension (savings or not), and nothing widens the page |
| How long does the pension last?, driven in the same browser | born 1980, ITP-1, 38 500 kr/month, an average pensioner's budget read back out of the answer boxes, the options table, both charts and the budgets table, then each of minPension's budgets and an own one, before tax, an individual pension, a private pension, an ISK and a depot (each as the same numbers the unit tests give, the tax-free ones grossed up before tax, the notes under the fields), each capital/saving pair on one row at desktop and phone width, no agreement, the retirement age moved and clamped, the lifespan, a typed final salary, the riktålder following the birth year, a preset, reset, the one-time import from Prognos, three CSV downloads and an Excel one, the hover and the arrow keys, both themes, both languages and 375px | the lifelong payout covers an average pensioner's budget all the way to 95 (the tab opens on a lifespan of 95) and a 5-year one to 73; a Pluspensionär's budget is covered longest by 11 years; before tax the budget of 18 400 kr is 22 204 kr, and the same years are covered; and nothing widens the page |
| A press on a button right after a number was typed, driven in the same browser | a real mouse (move, down, up through the browser's own input, none of the retrying a test click does) and a finger on a phone, with a number typed and its field not left: the tabs, Normalt/Avancerat, the language, Figur 2's and Table 2's toggles, the download buttons of Prognos, of "How long does the pension last?", of "When to take out the pension?" and of the agreement comparison, a button among the inputs, a number out of range, and a field that was only clicked into; and 10 unit tests for which button a press is handed on to | every press lands the first time, a download is written once and holds the number that was just typed (it is the file a press after the field is left writes), and a build without the fix fails it |
| VBA arithmetic semantics, delningstal, wages, ATP, the eight occupational agreements, private saving, tax rules, benefits, the main loop | 614 unit and property tests | — |
| The 32 mechanically translated tax functions | re-translated from the VBA by `npm run check:transpile` | match |
| The riksnorm tables | re-parsed from the VBA by `npm run check:riksnorm` | 113 rows match |
| The lowest pension age and riktålder, per cohort | re-read from the workbook by `npm run check:ages` | 128 cohorts match |
| The workbook addresses the export and the harness depend on | re-read from the workbook by `npm run check:names` | 11 defined names match |
| The social assistance norm for one 2025 household | the workbook author's own `verb()` comment | 46 240 kr/month, exact |

The `Brutto` fixture covers the earning phase only — that sheet never draws a pension, and it
pairs each year's pension right with the following year's indexation where the loop pairs it with
this year's. **The golden file is what closes the rest**, and it is the only check that compares
the engine against the real model end to end: `reference/golden/golden-cases.csv` holds 299 typfall
the workbook itself computed, and `npm run compare` runs the engine over the same inputs and diffs
all twelve output columns — the final salary, the five public pensions and their total, the
occupational pension, private saving, and the tax, benefits and disposable income at retirement.
Every one of the 3 588 comparable cells matches, to the last decimal the CSV carries.

The file certifies the workbook that made it: `# publicavg:` records a check that its tax ceilings
are the ones this port mirrors, and `# live.<name>:` records the settings whose `Adv_settings` row
holds no value of its own. Both are read back by the harness, which refuses a file that fails
either. They exist because two earlier attempts at this file were wrong in ways nothing could see.
See [`reference/golden/HOWTO.md`](reference/golden/HOWTO.md) for how to produce it and how to read
the report.

`reference/fixtures/default-run.json` holds the engine's own output for the shipped typfall. It is
a regression snapshot, not a check against the workbook: it makes an unintended change to any rule
show up as a diff.

## Continuous integration and publishing

`.github/workflows/ci.yml` runs everything the tables above claim — typecheck, the 614 unit and
property tests, the four re-derivation checks, the golden file against Excel, the build, and the
offline check in a real Chromium — on every push and pull request. The claims are the point of this
repository, and a claim nothing re-runs is a claim about the day someone last ran it by hand. The
whole run is a couple of minutes, so nothing is held back for a nightly job.

`check:ages` reads the `.xlsb` itself and so needs `pyxlsb`; CI builds the same `.venv` from
`tools/extract/requirements.txt` that the instructions below describe, rather than hand-picking that
one package, so a new extraction dependency reaches CI without anyone remembering to add it. The
other three `check:*` scripts re-derive from committed text and need nothing. LibreOffice is still
only for the full extraction. Playwright is installed for the offline step alone rather than declared
as a dependency — the same thing the instructions above tell a human to do.

`.github/workflows/deploy.yml` builds from source, re-runs the offline check, and then publishes:
GitHub Pages from `main`, and a release asset on a `v*` tag. Publishing a page that reaches for the
network is the one thing this project must never do, so the check runs again on the way out rather
than trusting the artifact.

> Two settings have no API and were done by hand: **Settings → Pages → Source** set to *GitHub
> Actions*, and the repository's default branch. Deploy fails at the `pages` job until the first is
> set.

## Regenerating the data

```bash
python3 -m venv .venv
.venv/bin/pip install -r tools/extract/requirements.txt
sudo apt-get install -y libreoffice-calc     # needed to read worksheet formulas
.venv/bin/python tools/extract/run.py source/Typfallsmodellen.xlsb
```

Deterministic: the same workbook in gives byte-identical files out, so `git diff` shows only what
changed between releases. Full procedure in [docs/YEARLY-UPDATE.md](docs/YEARLY-UPDATE.md).

## How it fits together

The engine is a faithful port, not a reimplementation. It mirrors the VBA function by function —
same names, same parameter order, same optional defaults — so that next year's diff can be
applied by eye. What that costs in elegance it repays in being checkable against the original.

Correctness rests on three things:

1. **Fixtures from inside the workbook.** It caches its own computed annuity factors and economic
   projections, which validate the actuarial core and the projection rules without needing Excel.
2. **Golden files from real Excel runs.** The acceptance gate — every Table 1 figure within
   1 kr/month across a few hundred typfall.
3. **A VBA-semantics shim.** `Int` floors rather than truncates, `Round` is banker's rounding
   while `WorksheetFunction.Round` is not. The model rounds deliberately at dozens of points, so
   these differences are the difference between right and nearly right.
