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

**Open question.** `vbaRound` at more than zero digits is unverified against a running VBA. It
scales by a power of ten and then tests the half exactly, which is what the VBA runtime does, but
that differs from a decimal-aware implementation for values like 2.675. The model calls
`Round(x, 2)` about ten times, on continuous quantities where landing exactly on a half is rare.
Phase 2's golden-file comparison settles it; a disagreement of a hundredth points here first.

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

### Deviations

- **The projection follows the worksheet, not `startsetup`'s fallback.** `startsetup` contains a
  simplified projection for years past the sheet's last row (2153), dividing by 257.4 where the
  sheet uses 257.38. The sheet governs every year that will realistically be computed, so the
  engine mirrors the sheet.
- **`inkomstbasbelopp` in the final year.** It reads the *next* year's income index; the
  workbook's last row has no next row and holds 0. The engine computes a year beyond the
  requested range and returns the real value.

## Status

| Area | VBA source | Ported |
|---|---|---|
| Arithmetic semantics | — | ✅ `src/vba/` |
| Economic series + projection | `startsetup`, `Balansindex` | ✅ `src/data/` |
| Mortality, delningstal, arvsvinster | `Mortality.bas`, `aaDeltal.bas` | |
| Wage vector | `Lön_mm.bas` | |
| Public pension | `Pensionssystemet.bas` | |
| Occupational pension | `Tjänstepensioner.bas`, `TjänstepensionerFörmån.bas` | |
| Private saving | `PrivatSparande.bas` | |
| Tax rules | `Skatteregler.bas` | |
| Benefits | `Bidrag.bas` | |
| Main loop | `Mcalc` | |
