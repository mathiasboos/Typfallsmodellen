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
| 4. Advanced mode | done | the workbook's second mode: twenty-five of the `Adv_settings` in the manual's own groups, the Indata_lista wage vector as an editable grid, and the PGB sheet's manual pension-qualifying amounts as a second grid |
| 5. Polish, CI, deploy | done | every check above runs on each push; the site publishes to GitHub Pages and each tag attaches the file to a release |
| 6. Compare scenarios *(not from the workbook)* | done | a second top-level view: the baseline typfall alongside up to three variants, each with its own salary, retirement age and occupational pension |

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
| Compare scenarios, driven in the same browser | a scenario's own salary raised well above the baseline's, and the add/remove scenario limits | only that scenario's own result moves, the baseline stays put, and the card count stays between one and four |
| VBA arithmetic semantics, delningstal, wages, ATP, the eight occupational agreements, private saving, tax rules, benefits, the main loop | 561 unit and property tests | — |
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

`.github/workflows/ci.yml` runs everything the tables above claim — typecheck, the 566 unit and
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
