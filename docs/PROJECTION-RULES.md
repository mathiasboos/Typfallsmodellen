# Projection rules for the economic series

`packages/data/economic-series.json` carries **actual values only**. Every year past a series'
`lastActualYear` is projected at runtime from the user's assumptions. This file is the
specification for that projection.

## Why the data stops at the actuals

The workbook's 'Några tal' sheet holds each series' decided values as typed-in literals and
computes everything after them with formulas that read the Start sheet's inflation, growth and
return inputs. Snapshotting those computed cells would silently bake in the shipped file's
assumptions (0% inflation, 0% growth, 1.7% real return) and the site would then return the same
numbers no matter what the user typed.

So `tools/extract` detects, per column, the last cell that is a literal rather than a formula —
that is the last actual year — and keeps only up to there.

## How these formulas were recovered

`.xlsb` stores formulas as parsed token streams, so `tools/extract/formulas.py` converts the
workbook to `.xlsx` with LibreOffice and reads them with openpyxl. LibreOffice drops the
defined-name targets, rendering each as `#NAME?`; the names below were identified from context
and from the `Adv_settings` variable-name column, then **confirmed by recomputation** — applying
these rules with the shipped settings reproduces the workbook's own cached values, which
`reference/fixtures/economic-series-cached.json` exists to keep verifying.

## Sheet geography

Year `Y` is on row `Y - 1953` on both 'Några tal' and Nyckeltal. The VBA writes that offset two
different ways — `year - 1959 + 6` and `year - 1958 + 5` — which are the same row.

Nyckeltal is what `startsetup` reads, but it is a pass-through of 'Några tal' unless "syntetisk
framskrivning" is on:

```
Nyckeltal = IF(OR(rng_Since_Year_Nyckeltal = 0, year <= rng_Since_Year_Nyckeltal),
               'Några tal' value,
               locally projected value)
```

With the default `rng_Since_Year_Nyckeltal = 0` every year passes through, so 'Några tal' is the
source of truth and the extractor reads it directly.

## Anchors

Read from the workbook into `economic-series.json.anchors`, so a release that revises them is
picked up rather than silently ignored.

| Anchor | Cell | Value | Meaning |
|---|---|---|---|
| `pbbBase` | `E4` | 36 396 | Prisbasbelopp basis |
| `fpbBase` | `H4` | 37 144 | Förhöjt prisbasbelopp basis |
| `ibbBase` | `G4` | 43 313 | Inkomstbasbelopp basis |
| `kpiJuneDivisor` | `B44` | 257.38 | KPI June 1997 |
| `incomeIndexDivisor` | `I52` | 118.41 | Inkomstindex 2005 |
| `kpiRoundDecimals` | `C2` | 2 | Decimals the KPI projection rounds to |

Note the VBA's own fallback (`VBA_go.bas`, `startsetup`) divides by **257.4**, not 257.38. That
fallback only runs past the sheet's last row (2153), so the sheet's 257.38 governs every year the
engine will realistically compute. Use 257.38.

## The rules

`infl` = `rng_Yearly_Inflation`, `growth` = `rng_Real_Growth`, `avk` = `rng_FondAvkastning`,
`marginal` = the rounding switch (0 = round as the real rules do, 1 = no rounding).

| Series | Rule | Sheet formula |
|---|---|---|
| `kpiJune`, `kpiAnnual` | `inflation` | `ROUND(prev * (1 + infl), 2)` |
| `prisbasbelopp` | `fromKpiJune` | `ROUND(pbbBase * kpiJune[t-1] / 257.38, -2)`; unrounded when `marginal = 1` |
| `forhojtPrisbasbelopp` | `fromKpiJune` | `ROUND((kpiJune[t-1] / 257.38) * fpbBase, -2)` |
| `inkomstbasbelopp` | `fromIncomeIndex` | `ROUND((inkomstindex[t+1] / 118.41) * ibbBase, -2)` — note **t+1** |
| `inkomstindex` | `indexGrowth` | `prev * (1 + growth) * (1 + infl)` |
| `medelPgi` | `followIndex` | `prev * inkomstindex[t] / inkomstindex[t-1]` |
| `balanstal` | `neutralOne` | projects to `1.0`, *not* to its last actual |
| `balansindex` | `balansindexFn` | `balansindex((balanstal[t] - 1) / 3 + 1, inkomstindex[t], inkomstindex[t-1])` — the VBA function in `Pensionssystemet.bas` |
| `gallandeIndex` | `gallandeIndex` | before balancing began: `= inkomstindex[t]`; from then on `IF(rng_Senaste_Index_Framskrivning = 2, inkomstindex[t], balansindex[t])` — derived every year |
| `skiktgrans1` | `kpiPlusTwo` | `INT(prev * (kpiJune[t-1]/kpiJune[t-2] + 0.02) / 100 + 51) * 100`; unrounded `prev * (…)` when `marginal = 1` |
| `skiktgrans2` | `constantCeiling` | `10^16` — a literal ceiling, not the last real threshold carried forward. Värnskatten was abolished in 2020 |
| `kvarEfterAdminIp` | `carryForward` | `prev` |
| `adminavgiftPp` | `feeAdmin` | `IF(rng_Avkastning_fondavgifter = 1, 0, 0.0002662%)` |
| `forvaltningsavgiftPp` | `feeManagement` | `IF(rng_Avkastning_fondavgifter = 1, 0, 0.14%)` |
| `totalAvgiftPp` | `feeTotal` | `IF(rng_Avkastning_fondavgifter = 1, admin, admin + management)` — derived every year |
| `kvarEfterAvgiftPp` | `oneMinusFee` | `1 - totalAvgiftPp[t]` — derived every year |
| `avkastningPpm`, `avkastningAp7` | `assumedReturn` | `(1 + avk - IF(rng_Avkastning_fondavgifter = 0, admin + management, 0)) * (1 + infl) - 1` |
| `rantaRiksgalden` | `rgkSetting` | `Rgk * 100` |

### Balancing only exists from 2010

The sheet writes `gallandeIndex` two different ways. Before the balancing mechanism first bit, it
is simply `=I7` — the income index. From 2010 it becomes
`IF(rng_Senaste_Index_Framskrivning = 2, inkomstindex, balansindex)`, and `balansindex` is blank
before that year. The engine reads the boundary out of the data — the first year `balansindex`
carries a value — rather than hard-coding 2010, so a revised history is picked up on the next
extraction.

`balansindex` itself is the VBA function, and it takes **four** arguments:
`Balansindex(previousBalansindex, dampedBalanstal, inkomstindex[t], inkomstindex[t-1])`.
LibreOffice renders the sheet's call with only three, dropping the first; the commented-out test
line in `Pensionssystemet.bas` shows the real shape, and the four-argument reading reproduces
every projected year exactly while the three-argument one divides by zero.

### The `+ 51` in skiktgrans1

`INT(prev * (kpiRatio + 2%) / 100 + 51) * 100` adds 5 100 kr before truncating to hundreds. It
reproduces the sheet exactly (2026's 643 000 projects to 660 900 for 2027, which is what the
sheet holds), but it does **not** reproduce the real decided figures — 2025's 625 800 would give
643 400 rather than the actual 643 000. That is a quirk of the workbook's projection, not of the
law. Mirror it: the goal is parity with the model, not a better model.

### The medium-term forecast branch

Several formulas carry a further branch, guarded by `rng_mftid` (Adv_settings, "Medelfristig
prognos om fem år"), that takes values from the `Pensioner` sheet for a five-year window
(`Pensioner!R3..R7`, currently 2025–2029). It is **off by default** (`rng_mftid = 0`). Phase 1
implements the default path; the branch belongs with the rest of the advanced settings in
Phase 4.

### One place the engine is deliberately better than the sheet

`inkomstbasbelopp` reads the **next** year's income index. The workbook's final row has no next
row, so it holds a 0 there. The engine computes one year beyond the requested range and returns
the real value; the fixture comparison skips that single cell.

## Verifying a change to this code

`reference/fixtures/economic-series-cached.json` holds every value the workbook had cached,
projections included, computed with the shipped assumptions. The engine's projection, run with
those same assumptions, must reproduce it exactly. That fixture is regenerated with the data, so
it tracks each yearly release.
