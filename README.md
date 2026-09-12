# Typfallsmodellen på webben

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
packages/data/     generated data: economic series, mortality, tax, i18n, content
packages/engine/   TypeScript port of the VBA calculation engine; no UI dependencies
                   run(input, context, { deaths }) -> TypfallResult
apps/web/          the website
reference/         VBA source dump and regression fixtures; never shipped
docs/              architecture, projection rules, yearly-update runbook, VBA mapping
```

## Status

| Phase | | |
|---|---|---|
| 0. Extraction pipeline | done | `.xlsb` → committed, diffable data |
| 1. Engine core | done | the whole model runs: `run(input, context)` returns Table 1, Table 2, the life-income sums and the per-age matrix |
| 2. Golden-file harness vs. Excel | awaiting the CSV | harness written and `npm run compare` wired up; needs `reference/golden/golden-cases.csv` out of Excel — `ExportGoldenCasesQuick`, about half an hour |
| 3. Normal-mode website | | |
| 4. Advanced mode | | |
| 5. Polish, CI, deploy | | |

### Verified so far

| What | Against | Result |
|---|---|---|
| Economic projection | the workbook's cached values for 20 series across ~190 years | exact |
| Annuity factors (delningstal, arvsvinstfaktorer) | the 114 345 values the workbook computed into `mortality!P:Z` | exact |
| Pension-qualifying income and contributions | the `Brutto` sheet, which calls those VBA functions as worksheet UDFs | exact |
| Income pension accumulation | the same sheet, term by term — gains, indexation, cost, balance | exact |
| The earning phase of the main loop, end to end | the `Brutto` sheet's per-age trace — PGI, PGB and all three contributions | exact |
| VBA arithmetic semantics, delningstal, wages, ATP, the eight occupational agreements, private saving, tax rules, benefits, the main loop | 542 unit and property tests | — |
| The 32 mechanically translated tax functions | re-translated from the VBA by `npm run check:transpile` | match |
| The riksnorm tables | re-parsed from the VBA by `npm run check:riksnorm` | 113 rows match |
| The lowest pension age and riktålder, per cohort | re-read from the workbook by `npm run check:ages` | 128 cohorts match |
| The workbook addresses the export and the harness depend on | re-read from the workbook by `npm run check:names` | 11 defined names match |
| The social assistance norm for one 2025 household | the workbook author's own `verb()` comment | 46 240 kr/month, exact |

The `Brutto` fixture covers the earning phase only — that sheet never draws a pension, and it
pairs each year's pension right with the following year's indexation where the loop pairs it with
this year's — so no **drawdown** figure and no end-to-end result is verified against the workbook
yet.

That is what Phase 2 closes. The comparison harness is written and `npm run compare` is wired up;
what it still needs is `reference/golden/golden-cases.csv`, which only a run of the real model on
Windows can produce. Until that file lands the comparison skips rather than fails. See
[`reference/golden/HOWTO.md`](reference/golden/HOWTO.md) for how to produce it and how to read the
report.

`reference/fixtures/default-run.json` holds the engine's own output for the shipped typfall. It is
a regression snapshot, not a check against the workbook: it makes an unintended change to any rule
show up as a diff.

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
