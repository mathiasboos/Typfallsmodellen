# VBA → TypeScript mapping

The engine is a faithful port, not a reimplementation. It mirrors the VBA function by function —
same names, same parameter order, same optional defaults — so that next year's diff of
`reference/vba/` can be applied by eye. What that costs in elegance it repays in being checkable
against the original.

This file records the mapping and, more importantly, every place the port deliberately deviates.

## Conventions

- VBA `Function gp(...)` → TypeScript `gp(...)`, same argument order, VBA's `Optional x = 40`
  becoming a default parameter.
- Swedish identifiers are kept. `garp`, `pbb`, `IBB`, `deltal` are the names in every source
  document about this model; translating them would make the port harder to check, not easier.
- Age-indexed arrays keep their VBA indexing via `AgeArray` (see below).
- Anything reachable from `Sub Mcalc` gets ported; anything not, does not. `akassa`, `sjuk`, `SA`
  and `forbehall` are inherited from a wider household model and are unreachable here.

## Arithmetic semantics — `src/vba/math.ts`

The quietest source of wrong answers. The model rounds deliberately at dozens of points, and
`marginal = 0` switches real-world rounding on across the whole rule set.

| VBA / Excel | TypeScript | Note |
|---|---|---|
| `Int(x)` | `vbaInt` | Floors. **Not** `Math.trunc` — they differ for every negative non-integer. Called 1 061 times, usually as `Int(x / 12 + 0.5) * 12` |
| `Fix(x)` | `vbaFix` | Truncates toward zero |
| `Round(x, n)` | `vbaRound` | Banker's rounding, half to even. All ~90 `Round(` calls in the VBA are this intrinsic |
| `WorksheetFunction.Round` / sheet `ROUND` | `excelRound` | Half **away from zero** — a different function, kept separate on purpose |
| assignment to `As Integer`/`As Long` | `vbaCLng` | Rounds, does not truncate |
| `a \ b` | `vbaIntDiv` | Operands rounded to Long, quotient truncated |
| `As Single` | `vbaSingle` | 32-bit narrowing, for ruling out a boundary divergence |
| `WorksheetFunction.Large/Average/Max/Min` | `wsLarge`, `wsAverage`, `wsMax`, `wsMin` | `Large` is 1-indexed and does not mutate its input |

**`vbaRound` at more than zero digits — resolved.** This was an open question: the implementation
scales by a power of ten and tests the half exactly, mirroring the VBA runtime, which differs from
a decimal-aware implementation for values like 2.675. `Calculate_Deltal` rounds every one of its
outputs with `Round(x, 2)` or `Round(x, 6)`, and the port reproduces all 114 345 cached values
exactly — so the behaviour is now confirmed against the original at both digit counts.

## Age-indexed arrays — `src/vba/ageArray.ts`

The VBA works in `ReDim x(startage To slutage)` arrays — age 15 to 105 — and reads `x(age - 1)`
freely. `AgeArray` keeps that indexing so the ported code stays line-comparable; rebasing to zero
would mean adjusting an offset at several hundred call sites. Out-of-bounds access throws rather
than yielding `undefined`, except `getOrZero` for the `x(age - 1)` idiom at the lower edge.

## Data and projection — `src/data/`

| VBA | TypeScript | Note |
|---|---|---|
| `startsetup`'s sheet reads | `packages/data` + `projectEconomicData` | Actuals extracted; later years projected. See `PROJECTION-RULES.md` |
| `Balansindex` (`Pensionssystemet.bas:1296`) | `balansindex` | Four arguments, per the VBA's own commented test line |

## Mortality and annuity factors — `src/pension/mortality.ts`

| VBA | TypeScript |
|---|---|
| `ReadMortality` (loading half) | `DeathProbabilities`, `loadDeathProbabilities` |
| `Calculate_Deltal` | `calculateDeltal` |

### Two quirks kept on purpose

`Calculate_Deltal` contains two pieces of behaviour that look like slips and that a clean
reimplementation would quietly correct. Both are kept, and both are load-bearing — reverting
either one breaks thousands of the cached values:

1. **`riktage` carries across the sex loop.** It is initialised once, before `For sex = 1 To 2`,
   and the loop leaves it at its final value — so the early ages for women inherit whatever the
   men's pass ended on, shifting which mortality year they read. Resetting it per sex changes
   **10 660** of 114 345 values.
2. **The discount rate reads a leftover `year`.** By the time the premium pension's
   forskottsränta is selected, `year` holds `cohort + 106` from the preceding age loop rather
   than the year being discounted. For every cohort the model runs this lands past 2017, so the
   rate is 1.0165 throughout. Using the year actually being discounted changes **23 540** values.

The port keeps both and says so at the site. Parity with the model is the goal; a better model is
a different project.

### Deviations

- **The projection follows the worksheet, not `startsetup`'s fallback.** `startsetup` contains a
  simplified projection for years past the sheet's last row (2153), dividing by 257.4 where the
  sheet uses 257.38. The sheet governs every year that will realistically be computed, so the
  engine mirrors the sheet.
- **`inkomstbasbelopp` in the final year.** It reads the *next* year's income index; the
  workbook's last row has no next row and holds 0. The engine computes a year beyond the
  requested range and returns the real value.

## Settings — `src/model/context.ts`

The rule modules read about forty Excel named ranges. `ModelContext` replaces them with one
frozen object threaded through the engine, so the functions stay pure and testable.

Defaults come from the workbook itself: `Adv_settings` carries its own variable name in column 9
and its "normal" value in column 8, both extracted into `packages/data/options.json`. Two
exceptions, `Soc_tak` and `Social_avg`, are policy-experiment switches with no row on that sheet;
both are guarded by `> 1999` in the VBA, so they default to 0 and stay inert.

## Contributions — `src/pension/contributions.ts`

| VBA | TypeScript |
|---|---|
| `riktage`, `andel` | same names |
| `pgi` | `pgi`, with the `Typ` argument as the `PgiResult` enum |
| `ipavgift`, `ppavgift`, `gpavgift`, `PPMavg` | same names, `ppmavg` |

### A quirk kept on purpose

`pgi` declares the tax-reduction phase-in share `Dim andel As Long`, so assigning 0.25 and 0.5
rounds them to **0** (0.5 by banker's rounding) and 0.75 and 0.875 to **1**. The reduction is
therefore all-or-nothing, arriving in 2002 rather than phasing in from 2000 as the surrounding
code reads as intending. Kept, and marked at the site.

**Unverified.** The Brutto fixture's typfall starts in 2016, so the 2000-2005 years where this
bites are never exercised — reverting the quirk does not break any test. The golden files will
settle it.

## Delningstal — `src/pension/deltal.ts`

Ports `aaDeltal.bas`, which assembles one table from two sources: published values from the
Nyckeltal sheet for every cohort, overwritten from 1958 with the model's own unisex figures.

**Deviation.** The VBA reads those second values from the mortality sheet's cached output; the
port computes them with `calculateDeltal`, which reproduces that table exactly. Equivalent, and it
means only the cohort being modelled is computed rather than all 121 — which matters in a browser.

Two details worth knowing before touching this file:

- The two-decimal rounding here is `Int(x * 100 + 0.4999) / 100`, which sends an exact `.xx5`
  **down**. It is not `vbaRound`, and the VBA comments on the difference.
- Part-year ages are weighted by **whole months**, not the raw fraction.

`fnDeltal_PP2` omits the clamp at the final withdrawal age that `fnDeltal_IP2` applies. Kept, with
a test asserting the asymmetry so the two do not get tidied into agreement.

## Wages — `src/income/wages.ts`

**Deviation.** The VBA computes the pension withdrawal share into the global `uttagIP` as a side
effect of asking for a wage. Here that is `withdrawalShare`, an exported function the caller
applies — same arithmetic, but a wage function that only returns a wage.

### Dead code kept

`If PAR = def_ar Then uttagIP = 1` sits inside a branch that requires `age >= par` **and**
`age < defAr` at once — impossible when the two are equal, which is the ordinary case. So in the
year pension is both first and finally drawn the share falls through untouched and stays at the
previous year's value. It changes no result: the months worked that year are computed by a formula
whose withdrawal term is multiplied by `defAr - par`, i.e. zero. Kept and marked, because a reader
would otherwise "fix" it.

## Income and premium pension — `src/pension/incomePension.ts`

| VBA | TypeScript |
|---|---|
| `IP_` | `incomePensionYear`, returning all five quantities at once; `ipResult` picks one by `Typ` |
| `ppkassa`, `deltal`, `P_uttag`, `pgb_barn` | `ppkassa`, `deltal`, `pUttag`, `pgbBarn` |

Note there are **two** functions called `deltal` in the model, and they are not the same: the one
here indexes the published tables directly by column arithmetic and is what `ppkassa` calls, while
`fnDeltal_IP` in `deltal.ts` is the newer lookup with the mortality override.

### A quirk kept on purpose

For a retirement age below 61, `deltal` extrapolates from the **premium pension** table's first two
ages — whichever pension is being asked about — because the VBA reads fixed columns 28 and 29 there.

### What the fixture reaches

The Brutto sheet's inline formulas reproduce `IP_`'s earning-phase branch term by term, so
inheritance gains, indexation, management cost and the closing balance are each checked against the
original across ~50 years. That sheet never draws a pension, so **none of the drawdown branches are
verified** — those wait on the golden files.

## A precision trap in the extracted data

The extractor originally rounded every value to twelve significant digits, to keep the generated
JSON tidy. That was wrong, and it took the Brutto comparison to reveal it.

Several series are factors just above or below 1 — inheritance gains at 1.0003, management-cost
factors at 0.9997 — and the engine uses them as `x - 1`. That subtraction cancels the leading
digits, so a twelve-digit value carries a **~1e-8 relative error** into the result: small enough to
look like noise, large enough to miss the workbook's figures.

The extractor now stores the exact double. Python writes floats at shortest round-trip precision,
so the output is both exact and byte-stable between runs. If you are tempted to tidy those numbers
again, this is why not.

## Status

| Area | VBA source | Ported |
|---|---|---|
| Arithmetic semantics | — | ✅ `src/vba/` |
| Economic series + projection | `startsetup`, `Balansindex` | ✅ `src/data/` |
| Mortality, delningstal, arvsvinster | `Mortality.bas` | ✅ `src/pension/mortality.ts` |
| Delningstal lookup | `aaDeltal.bas` | ✅ `src/pension/deltal.ts` |
| Contributions (PGI, avgifter) | `Pensionssystemet.bas` (first part) | ✅ `src/pension/contributions.ts` |
| Settings | ~40 named ranges | ✅ `src/model/context.ts` |
| Wage vector | `Lön_mm.bas` | ✅ `src/income/wages.ts` |
| Public pension | `Pensionssystemet.bas` | ✅ `src/pension/incomePension.ts` (ATP `tp_` still to do) |
| Occupational pension | `Tjänstepensioner.bas`, `TjänstepensionerFörmån.bas` | |
| Private saving | `PrivatSparande.bas` | |
| Tax rules | `Skatteregler.bas` | |
| Benefits | `Bidrag.bas` | |
| Main loop | `Mcalc` | |
