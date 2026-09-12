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
- `Typfallsmodellen-yyyy-mm-dd.xlsb` — the same version as `source/Typfallsmodellen.xlsb` in this
  repository. If they differ, the comparison checks the engine against the wrong rules.

## Steps

1. **Open the workbook** and click **Aktivera innehåll** when Excel warns about macros.

2. **Check the advanced settings are at their defaults.** Go to the `Adv_settings` sheet and click
   **Använd normala inställningar**. The reference run should exercise the model as it ships;
   anything else, and the engine would be compared against settings it was not asked to reproduce.

3. **Import the macro.** Press `Alt+F11` for the VBA editor, then **File → Import File…** and
   choose `ExportGoldenCases.bas` from this folder. If an older copy of the module is already there,
   right-click it → **Remove ExportGoldenCases…** → **No** (do not export it) before importing,
   or the import lands beside it as `ExportGoldenCases1` and you will run the wrong one.

4. **Run it.** Still in the VBA editor, press `F5` (or **Run → Run Sub/UserForm**) with
   `ExportGoldenCases` selected. Excel asks where to save the CSV.

5. **Wait.** It runs **295** typfall through the model, twenty-five at a time. Expect a couple of
   hours; Excel will look unresponsive while it works, and the status bar shows how far it has got.
   A dialog reporting the number of cases means it finished.

6. **Put the CSV here**, as `reference/golden/golden-cases.csv`, and commit it.

## What the macro does

It fills the `Mikrosim` sheet — the batch runner already built into the workbook — with a spread of
cases, calls the model's own `InputXGetY` over them, and writes the inputs and results to CSV
along with the model version and the settings in force.

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

## If the run halts in the VBA debugger

This has happened, and it is worth knowing it is not a failure.

The workbook has its own breakpoint: `Stop` at line 33 of `mdlIndataInputOutput`, inside
`WaitIfCalculationStateIsNotDone`, with the author's own comment beside it — *"Should never happen,
the full rebuild should fix the calculation state"*. It fires when Excel has not finished
recalculating within **0.2 seconds** (two 0.1-second waits) of being asked. On a sheet with a lot of
dirty cells that is a hair-trigger, not a fault, and it is the workbook's code rather than this
macro's.

**Nothing is lost when it fires.** The batch runner writes each row's results beside it as it goes,
so every case that finished is still on the `Mikrosim` sheet.

To recover:

1. In the VBA editor: **Run → Reset**.
2. Open the Immediate window (`Ctrl+G`) and run this one line, which puts back the application
   state the model changes while it runs:
   ```
   Application.Calculation = xlCalculationAutomatic: Application.EnableEvents = True: Application.ScreenUpdating = True
   ```
3. Run **`ExportGoldenCasesFromSheet`**. It writes the CSV from what is already on the sheet, runs
   nothing, and stops at the first row with no results — so a half-finished row is left out rather
   than exported as zeros.

If you would rather finish the remaining cases first, set `P3` and `U3` on the `Mikrosim` sheet to
the rows you still want and call `InputXGetY` directly, then export from the sheet.

## If something else goes wrong

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

It has been run once, against the 2025 workbook: 294 of the 295 cases completed before the
workbook's own watchdog tripped, and the run was recovered with `ExportGoldenCasesFromSheet` as
described above.
