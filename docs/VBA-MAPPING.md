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

## Tilläggspension — `src/pension/atp.ts`

| VBA | TypeScript |
|---|---|
| `tp_`, `tp_faktor`, `pts`, `fnorm` | `tp`, `tpFaktor`, `pts`, `fnorm` |

### The withdrawal factor is applied twice

`tp_` runs `ATP = ATP * faktor` and then returns `ATP * andel_ * faktor`. It reads as a leftover
from the 2021 rewrite that cut the function down to the first withdrawal — most of the original
body sits commented out between those two lines. At exactly 65 the factor is 1 and it makes no
difference; at any other retirement age it squares the adjustment.

Kept, marked, and asserted in a test. It affects only cohorts born 1953 or earlier. The golden
files will confirm it.

### One coupling made explicit

`tp_` reads the price base amount as `pbb(age)` off the module-level array, using the **global**
loop age rather than its own `alder` argument. Every call site passes the same value for both, so
the port takes `pbb` as a parameter.

## Occupational pension — `src/tjanstepension/`

| VBA | TypeScript |
|---|---|
| `tlITP1`, `tlITP2A`, `tlITP2F` | `itp.ts` |
| `SAF_LO`, `STP_` | `safLo.ts` |
| `tlkap_kl`, `tlakap_kr` | `kommunal.ts` |
| `tlPA16`, `Kapan`, `PA_indiv` | `statlig.ts` |
| `tjpkassa`, `tjp_ddeltal` | `tjpkassa.ts` |
| the `Select Case avtal` in Mcalc | `premiumFor` in `index.ts` |

**No fixture reaches this layer.** The `Brutto` sheet covers the public pension only, and the
model's own `deltal_tjp` cell is cleared when the workbook opens. The tests assert structure and
the transitions the VBA states outright — never that a rate is legally correct, because nothing
offline can confirm that. The golden files are the judge here, and this is the layer most likely
to need correcting when they arrive.

### The same guard, written three ways

In the retirement year these functions rescale salary to the months actually worked. The guard on
that rescale appears in two forms, and which one a function got decides whether it ever fires:

| Form | Where | Fires? |
|---|---|---|
| `Int(born + tjp_par - Int(born + tjp_par))` | `tlITP1`, `SAF_LO`, `tlPA16` | **Never** — `Int()` of a fraction below 1 |
| `Int((born + tjp_par - Int(born + tjp_par)) * 12)` | `tlITP2A`, `Kapan`, `PA_indiv` | Yes |

Both forms are kept as written. The first reads like a dropped `* 12`.

### Other quirks kept

- **The breakpoint is tested and applied with different divisors.** The `Select Case` compares
  salary against `7.5 * IBB / month`, while the premium formula uses `7.5 * IBB / 12`. They agree
  for a full year and part company in the entry and retirement years.
- **ITP 1's 2023 cap at 30 income base amounts is effectively inert**: it compares a *monthly*
  salary against an *annual* figure, so it only bites above roughly 360 base amounts a year.
- **KAP-KL's `alder < 21` early exit** makes its later `born > 1985 And alder < 21` branch
  unreachable; AKAP-KR has no such exit, so the same branch there does fire.
- **KAP-KL's 2006 rate branch reads `year = 2006 Or born <= 1946`**, and the `or` swallows the
  branch after it. AKAP-KR writes `and` in the same place, leaving both live.
- **`STP_` rounds with 0.49** where the rest of the module uses 0.5, and has a dead inner branch
  testing `par < 65` inside a `par > 65` block.

### Defined benefit — `formansbestamd.ts`

`FTJP` adds a benefit-based amount for the four agreements that carry one: ITP 2, SAF-LO (through
the legacy STP), KAP-KL and PA16 avdelning 2. Each scales its benefit by the ratio of the divisor
at 65 to the divisor at the actual retirement age, which is how drawing early or late is priced.

The fixed cell `mortality!J27` that the temporary-withdrawal correction reads is not stored data:
it indexes the table `Calculate_Deltal` writes, and holds the income pension divisor at 65 for the
cohort. For the default 1959 cohort the sheet shows 17.6, which the ported calculation reproduces —
asserted by a test, so the reading is pinned rather than assumed.

#### `PA_KLBPP`'s fallbacks can never fire

The function averages the best five of seven years' pension points, with a chain of fallbacks for a
shorter working life. Every one of them reads

```vba
ElseIf yearpoint7 = 0 & yearpoint6 <> 0 Then
```

and VBA's `&` is **string concatenation**, not `And`. The expression parses as
`yearpoint7 = ((0 & yearpoint6) <> 0)` — a comparison against a boolean, so `yearpoint7 = True`,
i.e. `-1`, which a salary ratio never is. With fewer than seven years of salary the chain therefore
falls through to the final `Else` and averages the **single** best year, never the four, three or
two the intermediate branches intend.

Kept as written, with a test that distinguishes the two behaviours so it cannot be silently
"corrected".

#### Smaller things kept

- `tlPA03` has a duplicated `ElseIf fodar = 1943` branch; the second is unreachable.
- `KAPKL_f`'s 1966 row reads `par3 = 0.268` where the sequence around it would suggest `0.2768`.
- In `FTJP`'s KAP-KL branch a negative carry-over from PA-KL is passed into `KAPKL_f` *as the
  salary basis* rather than added to the result.
- Only the PA16 avdelning 2 branch floors the cohort at 1938 when looking up divisors.

### A side effect for Mcalc to own

`STP_` assigns `born = Int(born)`, truncating the birth-month fraction on the module-level global
for everything computed after it in a run — and `born` does carry a month fraction whenever a
birth month other than January is chosen. It is called only from `FTJP`, so the truncation belongs
there rather than hidden inside the function. **Still to apply**: `FTJP` is ported, but the port
takes `born` as a value, so the truncation will be wired in when `Mcalc` lands and owns that state.

## Private saving — `src/saving/privateSaving.ts`

| VBA | TypeScript |
|---|---|
| `PrivatSpar`, `AvkastningsskattKFISK`, `Schablonintakt` | same names |

### The one module that is not pure

`PrivatSparande.bas` declares no `Option Explicit`, so its undeclared variables — `avkskatten`,
`Slr` — are implicit **module-level** Variants that persist between calls. Two branches do not
cover every input, and there the previous call's value stands:

- `AvkastningsskattKFISK` tests `> 150000` and `< 150000` (and the same at 300 000 from 2026).
  At exactly the allowance neither fires.
- `Schablonintakt` has no `Case` for a year before 1986.

Both are reproduced with module-level variables, because a golden-file case could land on a round
150 000. The cost is that these functions are **not pure**: the same arguments can give different
answers depending on what ran before. `resetPrivateSavingState()` clears them between independent
runs; the VBA has no equivalent, so a run following another in the same Excel session inherits the
earlier one's leftovers, which the port does not try to reproduce.

## Tax rules — `src/skatt/`

| VBA | TypeScript |
|---|---|
| `avdrag95`…`avdrag26`, `avdragxx`, `Xage` | `grundavdrag.ts` |
| `Jobb07`…`Jobb26`, `Jobbxx` | `jobbskatteavdrag.ts` |
| `statlig`, `PublicAvg`, `sared`, `FAared`, `pandred`, `arbgiv`, `avdragRES`, `avkskatt` | `reduktioner.ts` |

### The only mechanically translated code in the engine

`Skatteregler.bas` holds 32 near-identical rule functions — one per tax year — that are nothing but
thresholds and coefficients: roughly 1 300 lines of numbers. Typing those by hand invites exactly
the transposed digit no test would catch, so `tools/transpile/skatteregler.py` translates them, and
the result is reviewed and committed.

`npm run check:transpile` re-translates from `reference/vba/Skatteregler.bas` and verifies the
committed TypeScript still matches. That catches a hand-edit of a generated function, and it
catches the committed code drifting from a newly downloaded workbook. **Do not edit those
functions by hand** — change the VBA dump or the translator and re-run.

The translator handles a deliberately tiny subset (If/ElseIf/Else, assignment, arithmetic) and
raises on anything else, so an unhandled construct fails loudly rather than producing plausible
wrong code. Two things it has to get right:

- `Dim jobb As Long` in `Jobb07` and `Jobb08` means **every assignment there rounds** to a whole
  krona; the other ten declare `As Double` and do not. Those assignments are wrapped in `vbaCLng`.
- VBA reads and assigns the function's own name as a variable (`If Jobb14 < 0 Then Jobb14 = 0`), on
  both sides of a condition.

### Things kept as written

- **`sared` reads `alder` and `binkomst` that are not its parameters** and are never assigned
  anywhere in the module. With no `Option Explicit` they are implicit module-level Variants, always
  Empty — so both are passed as 0.
- **`sared`'s year blocks are `If`, not `ElseIf`**, so for a year before 2022 the second block
  overwrites what the first computed.
- **`korr` in `avdrag13` is computed and then zeroed**; the comments show the values it was meant
  to carry.
- **`arbgiv` multiplies the wage by a percentage** taken straight from K_skatt (31.42), so its
  amount is a hundred times too large. It is only ever written to the output sheet "för studier av
  arbetsgivarna" and never feeds a pension or a tax, so nothing downstream depends on it.

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
| Occupational pension, defined contribution | `Tjänstepensioner.bas` | ✅ `src/tjanstepension/` |
| Occupational pension, defined benefit | `TjänstepensionerFörmån.bas` (`FTJP`), `tlPA03`, `KAPKL_f`, `PA_KL`, `PA_KLBPP` | ✅ `src/tjanstepension/formansbestamd.ts` |
| Private saving | `PrivatSparande.bas` | ✅ `src/saving/privateSaving.ts` |
| Tax rules | `Skatteregler.bas` | ✅ `src/skatt/` |
| Benefits | `Bidrag.bas` | |
| Main loop | `Mcalc` | |
