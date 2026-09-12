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
apps/web/          the website
reference/         VBA source dump and regression fixtures; never shipped
docs/              architecture, projection rules, yearly-update runbook, VBA mapping
```

## Status

| Phase | | |
|---|---|---|
| 0. Extraction pipeline | done | `.xlsb` → committed, diffable data |
| 1. Engine core | in progress | arithmetic shim ✅ · economic projection ✅ · annuity factors ✅ · settings ✅ · delningstal ✅ · wages ✅ · contributions ✅ · income & premium pension ✅ · tilläggspension ✅ · occupational pension ✅ · private saving ✅ · tax rules ✅ · benefits ✅ · main loop — to do |
| 2. Golden-file harness vs. Excel | awaiting an Excel run | export kit ready in `reference/golden/` |
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
| VBA arithmetic semantics, delningstal, wages, ATP, the eight occupational agreements, private saving, tax rules, benefits | 383 unit and property tests | — |
| The 32 mechanically translated tax functions | re-translated from the VBA by `npm run check:transpile` | match |
| The riksnorm tables | re-parsed from the VBA by `npm run check:riksnorm` | 113 rows match |
| The social assistance norm for one 2025 household | the workbook author's own `verb()` comment | 46 240 kr/month, exact |

The `Brutto` fixture covers the earning phase only — that sheet never draws a pension — so no
**drawdown** figure and no end-to-end result is verified yet. That is Phase 2, and it needs a run
of the real model on Windows: see [`reference/golden/HOWTO.md`](reference/golden/HOWTO.md).

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
