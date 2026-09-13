# Producing reference results from the real model

The web port has to prove it returns the same numbers as the Excel model. Most of that proof
comes free — the workbook caches its own annuity factors and economic projections, and the engine
is tested against both. But it caches **no end-to-end pension figures**: `Workbook_Open` clears
the result sheets, so the distributed file arrives empty.

So the acceptance gate needs a real Excel run. This is the one step that cannot be automated from
here — it needs Windows and Excel.

**You only need to do this when the model version changes.** The resulting CSV is committed and
re-used on every subsequent build.

## What you need

- Windows with Excel 2007 or later
- **The same build as `source/Typfallsmodellen.xlsb` in this repository.** Not merely the same
  version number: two downloads have already been seen calling themselves "Version 4.8" with
  different code in `Skatteregler.bas`. The one this port is built from is the one whose
  `Versionsinformation` sheet carries *"Mindre buggfixar, t.ex. löpande priser, fasta löner och för
  lång kod i huvudmodulen."* under Version 4.8. Step 5 below checks it in a second; if you skip it
  and the builds differ, the comparison checks the engine against rules it was never given.

## Steps

1. **Open the workbook** and click **Aktivera innehåll** when Excel warns about macros.

2. **Check the advanced settings are at their defaults.** Go to the `Adv_settings` sheet and click
   **Använd normala inställningar**. The reference run should exercise the model as it ships;
   anything else, and the engine would be compared against settings it was not asked to reproduce.

3. **Import the macro.** Press `Alt+F11` for the VBA editor, then **File → Import File…** and
   choose `ExportGoldenCases.bas` from this folder. If an older copy of the module is already there,
   right-click it → **Remove ExportGoldenCases…** → **No** (do not export it) before importing,
   or the import lands beside it as `ExportGoldenCases1` and you will run the wrong one.

4. **Check the two runners agree.** Put the cursor in **`CompareRunners`** and press `F5`. It runs
   four cases both through the model's own batch runner and through this module's own per-row
   driver, and compares all twelve results. It should report that they agree exactly. This takes a
   couple of minutes and is worth doing once per model version — see "Why there are two runners"
   below.

5. **Check it is the right build.** Put the cursor in **`ReportPublicAvg`** and press `F5`. It runs
   one typfall and prints one line per rule year, each ending `OK` or `MISMATCH`, then a verdict.
   There is nothing to interpret. You do not have to remember to run it — every export runs the
   same check first and refuses to start on a mismatch — but running it now costs half a minute
   instead of finding out later. See below for why it exists.

6. **Run it.** Still in the VBA editor, put the cursor in **`ExportGoldenCasesQuick`** and press
   `F5` (or **Run → Run Sub/UserForm**). Excel asks where to save the CSV.

7. **Wait.** It runs **65** typfall through the model — a bit over half an hour. Excel will look
   unresponsive while it works, and the status bar shows how far it has got. A dialog reporting the
   number of cases means it finished.

   **The CSV is rewritten after every case**, so the file on disk is always complete for everything
   that has finished. A halt costs the case in progress, nothing else.

8. **Save the workbook** if you might want to add the rest later. The results live on the `Mikrosim`
   sheet, and closing without saving loses them — which is what makes `ExportGoldenCasesResume`
   possible or impossible.

9. **Put the CSV here**, as `reference/golden/golden-cases.csv`, and commit it.

10. **Run the comparison**: `npm run compare` from the repository root. See below for what it tells
   you.

### The other entry points

| Sub | What it does |
|---|---|
| `ExportGoldenCasesQuick` | 65 cases, ~32 min. The boundaries that carry the most information. |
| `ExportGoldenCases` | all 299, a couple of hours. Maximum coverage. |
| `ExportGoldenCasesResume` | keeps the inputs on the sheet and runs only the rows without results. Use it after a halt, or to work through the full set in sittings. |
| `ExportGoldenCasesFromSheet` | writes the CSV from what is on the sheet, recomputing nothing. |
| `ReportMikrosimState` | says what is actually on the sheet. **Start here when something looks wrong.** |
| `CompareRunners` | runs a few cases both ways and checks they agree. |
| `ReportPublicAvg` | says whether this workbook is the build the port was written against. Every export runs the same check first. See below. |

### `ReportPublicAvg`

Says whether the workbook in front of you is the build this port was written against.

It runs one typfall to fill `born` and `IBB()`, then asks the workbook's own `PublicAvg` what
ceiling it applies in each rule year — calling it with an income far above any conceivable ceiling,
which returns the ceiling over a hundred — and compares that against the factors the engine mirrors
from `Skatteregler.bas`. Each row ends `OK` or `MISMATCH`. The comparison is on the whole-krona fee
both sides round to, so it is exact; it cannot see a factor difference below about 0.0012, which is
one krona, and no real rule change is that small.

`ExportGoldenCasesQuick`, `ExportGoldenCases`, `ExportGoldenCasesResume` and
`ExportGoldenCasesFromSheet` all run the same check first and refuse to continue on a mismatch, and
the verdict is written into the CSV as `# publicavg:` so a committed golden file carries proof of
which build produced it.

**Why it exists.** The first golden file was exported from a build whose `PublicAvg` stopped at
2022 — every year from 2023 on used the 2022 factor of 1.87 — while `packages/data` was extracted
from the corrected build. Ten of the twelve columns matched anyway, which is exactly what makes
that failure mode dangerous: two builds can share every Nyckeltal, mortality and delningstal table
and differ in one rule, and only the rule that differs shows. It cost a week, and 258 kronor a year
of pension looked for all that time like a bug in the port. Both files called themselves
"Version 4.8"; the corrected one is distinguished only by an extra line in its `Versionsinformation`
sheet.

To go from the quick set to the full one: run `ExportGoldenCases` and let it re-run everything, or
run the quick set, save, and add cases by hand — there is no merge step, because the CSV is always
written from the sheet as a whole.

## What the macro does

It fills the `Mikrosim` sheet — the batch runner already built into the workbook — with a spread of
cases, calls the model's own `InputXGetY` over them, and writes the inputs and results to CSV
along with the model version and the settings in force.

Where things sit on the sheet is read from the workbook's own defined names (`rngXTopleft`,
`rngYtopleft`) rather than hardcoded, so a moved column in a future release cannot send it looking
in the wrong place. `npm run check:names` asserts those addresses from the committed workbook.

The cases are chosen for **coverage of rule boundaries**, not for row count. A thousand ordinary
cases would prove less than a few hundred that straddle the places where the rules change:

- cohorts either side of 1938 and 1954, where tilläggspension appears and disappears
- every one of the eight occupational pension agreements
- salaries from 15 000 to 90 000 kr/month, spanning the garantipension phase-out at the bottom and
  the state income tax threshold at the top
- retirement ages from each cohort's earliest through 75
- entry ages from 18 to 30, giving both short and long working lives
- economic assumptions away from the forecasting standard, varied one at a time as well as together
- the care-assistant example from the user manual, so the published figures can be checked directly

Retirement ages are clamped to each cohort's lowest permitted age, read from the `Nyckeltal`
sheet. Below that the model opens a Yes/No dialog mid-run and, answered Yes, quietly changes the
case — which would write an input into the file that did not produce the output beside it.

## What the CSV holds

Two blocks, then the data.

**The provenance block**, every line starting with `#`. The model version, the workbook name, the
export time, and then **every setting on `Adv_settings`** as `# adv.<name>: <value>`, taken from the
sheet itself — column 9 carries the name the VBA reads the setting by, column 2 its value. Plus
`# start.Gift`, which lives on the Start sheet rather than `Adv_settings` and which the batch runner
never sets per case, so whatever it held applied to all of them.

This block is not decoration. The comparison harness checks the engine's own defaults against it
setting by setting and refuses to report agreement it cannot vouch for: `rng_Bara_fastapriser` alone
rescales every value in the file, and `Alt_p_age` can silently override the retirement age in the
input column, which would pair an input with an output that did not come from it.

**The header row**, taken from row 7 of the `Mikrosim` sheet so it tracks any future column change.
Ten input columns (B–K) then twelve output columns (M–X). The first output column's label is
written by the model itself and says which of Slutlön, Nettolön or Disp. that column holds,
following `Rng_CompareTo`.

Every output value is read from **column D of Table 1** on the Start sheet — the adjusted column,
expressed in the reference year's prices or wage level — and divided by 12 when `Rng_belopp12` is 1.

## Reading the comparison

`npm run compare` runs the engine over every case in the file and diffs all twelve output columns.
It prints a column summary and writes the full report to `reference/golden/report.md`, which is not
committed.

Each cell is scored by how far apart the two sides are: **exact** below a millionth of a krona,
**close** below half an öre or a millionth in relative terms, **off** below one per cent, and
**bad** above it. Both sides do the same IEEE-754 arithmetic on the same inputs, so a faithful port
lands in exact or close. Anything else is a real difference in the rules, not floating-point noise.

The report groups the same divergences four ways — by case block, cohort, agreement and retirement
age — because that is what tells you where to look:

| What the grouping shows | Where to look |
|---|---|
| One agreement wrong across every cohort | `packages/engine/src/tjanstepension/` |
| One cohort wrong across every agreement | a rule-year boundary in `contributions.ts` or `atp.ts` |
| Every column off by the same factor | `adjustmentFactors` in `result.ts`, and the price basis |
| Only the last three columns wrong | `atRetirement.ts`, the second pass at the retirement age |
| One block wrong | whatever that block varies — see `BuildCases` in the macro |

Divergences that are a whole number of rounding steps (12 kr a year, or 1 kr a month) are counted
separately: they point at where the model rounds rather than at what it computes.

If the settings in the file's provenance block disagree with the model's normal values, the report
says so. For the eight that decide what the numbers *mean* — the price basis, `Alt_p_age`, `Risk`,
`Average_Earning`, `marginal`, `w_ref`, `rng_Sista_PensRatt` and `rngPens_Inflation` — the run stops
instead, because a comparison against the wrong settings would report agreement that is not there.

## Why there are two runners

The workbook has its own batch runner, `InputXGetY`, and it has a breakpoint that stops the whole
run in the VBA debugger. `WaitIfCalculationStateIsNotDone` (`mdlIndataInputOutput.bas:12`) waits for
`Application.CalculationState` to reach `xlDone`, and hits a bare `Stop` if it has not within
**0.2 seconds** — two 0.1-second waits — with the author's own comment beside it: *"Should never
happen, the full rebuild should fix the calculation state"*.

It does happen. The three lines before it are:

```vba
Application_Rest                   ' sets calculation to MANUAL (mdlTools.bas:1377)
wsIndata.Range("rngRunTime") = ... ' dirties formulas on the sheet
Application.StatusBar = ...
WaitIfCalculationStateIsNotDone    ' now demands a settled calculation state
```

In manual calculation Excel reports `xlPending` and stays there, because it will not recalculate on
its own. `CalculateFullRebuild` is meant to clear that; on a workbook this size, with `w_ref` on a
volatile `=YEAR(NOW())-1`, it does not reliably do so inside 0.2 seconds. The watchdog runs **after
every row**, so a long run gets a fresh chance to trip on each one. In practice it has stopped a
295-case run on case 294 and a 61-case run on case 25.

That routine is called from two places, both inside `InputXGetY`. Nothing else in the model calls
it — in particular `StartUp_Indata`, which is what actually runs `Mcalc`, does not. So this module
drives the per-row loop itself: it sets the same named ranges from columns B–K, calls
`StartUp_Indata`, reads the same twelve cells back, and leaves the calculation state alone instead
of demanding it be settled. `CompareRunners` is what shows that the watchdog is the only thing this
skips.

To go back to the model's own runner, set `USE_MODEL_BATCH_RUNNER` to `True` at the top of the
module.

## If the run halts in the VBA debugger

If you are on the model's runner — or running `CompareRunners`, which exercises it deliberately —
you can still land on that `Stop`. **Nothing is lost when it fires**: results are written to the row
before the watchdog runs.

Do not reset. In the Immediate window (`Ctrl+G`):

```
Application.Calculation = xlCalculationAutomatic
```

then press **F5**. That satisfies the condition the loop is waiting for, so it returns normally and
the run carries on. Repeat if it trips again.

To stop instead: **Run → Reset**, then the same line plus `Application.EnableEvents = True:
Application.ScreenUpdating = True`, **save the workbook**, and `ExportGoldenCasesFromSheet` to bank
what finished. `ExportGoldenCasesResume` picks up from the first row without results — the runner
only clears the result block for the rows it is about to run, so the finished ones survive.

## If something else goes wrong

**Run `ReportMikrosimState` first.** It prints the resolved layout, the used range, how many rows
carry inputs, how many carry results, and the first rows in full. Nearly every question below is
answered by it in one click.

**"No completed rows found on the Mikrosim sheet."** The results are not there. Either the workbook
was closed without saving after a run — the batch runner writes to the sheet, and the sheet is only
kept if you save — or the sheet was cleared. The message says whether the inputs are still there:
if they are, `ExportGoldenCasesResume` will finish the job; if not, start again with
`ExportGoldenCasesQuick`. This is also why the CSV is now written after every chunk: the file on
disk is the copy that survives.

**A dialog appears mid-run.** Note what it says and stop the run; the resulting CSV may contain
rows whose inputs and outputs disagree. Please report it rather than working around it.

**"Sub or Function not defined" on `InputXGetY`.** The macro was imported into a different
workbook than the model. Make sure the model workbook is the active project in the VBA editor.

**Numbers in the CSV use commas.** The macro forces a dot as the decimal separator, so this should
not happen — but if it does, say so rather than converting the file by hand; it would mean the
formatting helper needs fixing.

**It is very slow.** Each case is a full model run. If it is impractically slow, reducing the case
list is better than not producing the file at all — the boundary cases in blocks A and E matter
most.

## A note on trust

This macro was written without access to Excel. Read it before you run it — it is about 400 lines,
and it writes to the `Mikrosim` sheet (a scratch sheet the model provides for exactly this) and to
the CSV you choose. It does not modify the model or save the workbook.

It has been run once against the 2025 workbook. 294 of the 295 cases completed before the
workbook's own watchdog tripped, and the results were then lost with the unsaved workbook — because
that version wrote the CSV only at the end. That is what the per-chunk write fixes.
