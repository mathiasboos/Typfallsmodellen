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
> its user manual are theirs. Questions about the model itself go to
> `typfallsmodellen@pensionsmyndigheten.se`.

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
| 6. Compare scenarios *(not from the workbook)* | done | a second top-level view: the baseline typfall alongside up to three variants, each with its own salary, retirement age, start-of-work age and occupational pension, lined up in one shared comparison table and an overlay chart |
| 7. Mikrosim *(ported from the workbook's own Mikrosim sheet)* | done | a third top-level view: a batch of independent typfall in two linked tables (Inputs, then Results below it), each row's twelve output columns computed live as its own inputs are edited, added or imported, add/remove rows, CSV import/export, and a stacked column chart below breaking each row's pension into its seven components |

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
| Compare scenarios, driven in the same browser | a scenario's own salary and start-of-work age each raised well above the baseline's, the add/remove scenario limits, and the comparison table's own replacement-rate row | only that scenario's own column in the table and line in the chart move, the baseline stays put, the card/column/line count stays between one and four, and no replacement-rate cell ever reaches 100% |
| Table 2's header tooltips and Ordlista, driven in the same browser | the "Statlig skatt" header's own `<abbr title>`, and Ordlista's term count and language-switched summary | names the 20% threshold and the public-service fee; 57 terms present, staying Swedish across a language switch that still relabels the disclosure itself |
| CSV and Excel downloads, driven in the same browser | Table 1's, Table 2's and the comparison table's own download buttons, each actually clicked | a real file every time: the CSV's own content matches that table's data, the `.xlsx`'s first bytes are the ZIP signature every real OOXML package starts with |
| The built file's own script tag | a static check of the file itself, no browser involved | a classic `<script>`, not `type="module"` (which Mobile Safari can refuse to run at all over `file://`, seen as a black screen on an iPhone), placed after `#app` in the document |
| The PGB add-entry form and its sparse summary table, driven in the same browser | adding a sickness, study, conscription and child entry one at a time (including two in the same year, and a too-short conscription period), removing each one, and the whole panel at 375px | the summary table starts empty and only ever shows the years and categories that have data; each add/remove moves the pension and the table together; a child's birth year raises `PgbBreakdownYear.barn`, the field `pgbBarn` (new in the engine) exists to expose; the page never grows a horizontal scrollbar at 375px |
| Conscription and study's own PGB arithmetic | 15 unit tests pinning the day-split across one, two and three calendar years by hand, and the browser reading back a date range's own PGB värnplikt figure and a semester count's own PGB studier figure, each its own column in the summary table once it has data | exact day counts and kronor at every split; a 1998 conscription period and a 2005 semester both raise the pension on their own, distinct from each other and from a typed sickness/activity amount |
| Partial withdrawal of the public pension, driven in the same browser | age 67's own Lön and pension columns in Table 2, before and after setting a 50% withdrawal share and a final age of 70 | 0 kr salary and a full pension become a part-time salary alongside almost exactly half the full pension, and age 70 returns to 0 kr salary and a pension above the full-at-66 figure |
| Mikrosim, driven in the same browser | adding and removing rows, editing one row's salary and scheme with no button anywhere to click, CSV export and re-importing that same file, a file missing a required column, a row with an invalid scheme next to a valid one, and the stacked chart below the two tables | each row's own output columns (a separate Results table below the Inputs one) fill in immediately from its own inputs only, with no effect on any other row's; a bad file is refused by name with the table left untouched; an invalid row is flagged and left blank while the rest still compute, and fixing its own field recomputes it on the spot; the chart's own bar count tracks the table's row count, and a plain edit alone moves the chart's own bar values, not just its count |
| VBA arithmetic semantics, delningstal, wages, ATP, the eight occupational agreements, private saving, tax rules, benefits, the main loop | 581 unit and property tests | — |
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

`.github/workflows/ci.yml` runs everything the tables above claim — typecheck, the 581 unit and
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
