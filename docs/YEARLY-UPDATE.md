# Yearly update runbook

Pensionsmyndigheten publishes a new `Typfallsmodellen-yyyy-mm-dd.xlsb` roughly once a year. Each
release brings two kinds of change:

- **new data** — indices, basbelopp, mortality, tax rates, last year's outcomes. Mechanical.
- **new rule functions** — a new grundavdrag and jobbskatteavdrag for the new tax year, and
  whatever else changed in law. These need code.

The workbook's own changelog shows the pattern. Version 4.8: *"Regelförändringar för 2026, nytt
grundavdrag"*, *"Ökad avsättning inom PA16 Avd I och Avd II för flexpension"*. Version 4.7:
*"Uppdateringar av nyckeltal och några tal (delningstal, förra årets utfall…)"*.

## Steps

### 1. Get the new workbook

Download from
<https://www.pensionsmyndigheten.se/statistik-och-rapporter/pensionsmodellen/typfallsmodellen>
and replace `source/Typfallsmodellen.xlsb`.

**Write down which build it is, and keep that file.** The version number on the
`Versionsinformation` sheet is not a build identity: two downloads have already been seen calling
themselves "Version 4.8" with different code in `Skatteregler.bas`, distinguishable only by an extra
bullet in the version history. Everything downstream — the extracted data, the VBA diff, the golden
file — has to come from the *same* file, so keep the one you extracted from and export the golden
cases from that same copy, not from whatever is in the downloads folder later. See the
`ReportPublicAvg` section of `reference/golden/HOWTO.md` for what it cost to learn this.

### 2. Regenerate the data

```bash
python tools/extract/run.py source/Typfallsmodellen.xlsb
```

Needs LibreOffice for the formula pass:

```bash
sudo apt-get install -y libreoffice-calc
```

The run prints the model version and the detected `lastActualYear` for every series. Output is
deterministic, so `git diff` shows only what genuinely changed.

### 3. Review the data diff

Check `git diff packages/data/`:

- **`economic-series.json`** — each series should have gained one year, and `lastActualYear`
  should have advanced. A series that did *not* advance is worth a look: either the value is
  genuinely not decided yet, or the new workbook changed how that column is built.
- **`municipal-tax.json`** — a new year of kommunalskatt.
- **`mortality-risks.bin`** — SCB revises its forecast periodically; a changed file is normal.
- **`options.json`** — new advanced settings, or changed defaults, appear here. Anything new
  needs UI work in `apps/web`.
- **`riksnorm.json`** — Socialstyrelsen's social assistance norm, parsed out of `Bidrag.bas`
  rather than the workbook. Expect one new row in each of the three table families (`xn`, `vuxna`,
  `Gn`) for both `bistOld` and `bist25`, and the previous `else` row to become an `eq` row for
  the year it covered. Anything else — a changed historical row, a row of a different width — means
  the VBA changed shape and wants reading. Never edit this file by hand.

`lastActualYear` is detected exactly rather than guessed: a series' decided values are typed into
the sheet as literals and its projection starts at the first formula cell. If a future release
starts *computing* a value that used to be typed in, the detected year will drop — that is the
signal to look at the sheet, not to override the number.

### 4. Review the VBA diff

```bash
git diff reference/vba/
```

This is the part that needs judgement. Look for:

- **New rule functions.** They follow a strict naming pattern: `avdrag27` after `avdrag26`
  (grundavdrag), `Jobb27` after `Jobb26` (jobbskatteavdrag). Each is ~60 lines of bracket
  thresholds shaped exactly like its predecessor.
- **Changed dispatchers.** `avdragxx` and `Jobbxx` in `Skatteregler.bas` grow a branch for the
  new year.
- **Changed contribution rates** in `Tjänstepensioner.bas`, which move when a collective
  agreement is renegotiated.
- **Anything else.** A diff touching `VBA_go.bas` or `Pensionssystemet.bas` means the model's
  core changed and deserves real attention.

### 5. Port the code changes

Mirror each new VBA function in `packages/engine/src/`, keeping the 1:1 structure described in
`docs/VBA-MAPPING.md`. Add the dispatcher branch. Update the mapping doc.

### 6. Re-verify

```bash
npm test              # unit tests, including the cached-table fixtures
npm run check:transpile   # the 32 tax functions still match Skatteregler.bas
npm run check:riksnorm    # riksnorm.json still matches Bidrag.bas
npm run check:ages    # the per-cohort retirement ages still match Nyckeltal
npm run check:names   # the sheet addresses the golden-file export depends on
npm run compare       # engine vs. the Excel golden files in reference/
```

The two `check:` scripts re-derive generated code and data from the VBA dump and fail if the
committed version has drifted. They are the safety net for step 5: if a rule function was
mechanically translated and you changed it by hand, or a riksnorm row moved, they say so.

The fixtures in `reference/fixtures/` are regenerated from the new workbook, so they check the
new data. The golden files in `reference/golden/` are **not** — they came from a previous Excel
run. Regenerate them from the new workbook too (see `reference/golden/HOWTO.md`), otherwise the
comparison is checking the new engine against last year's rules. Do not skip that procedure's
step 5, `ReportPublicAvg` — it takes a second and it is what catches exporting from a different
build than the one you extracted the data from.

Keep the previous golden file as well: re-running it with the rule-year overrides pinned to last
year's rules should still reproduce last year's numbers, which catches accidental changes to
historical behaviour.

### 7. Ship

Update the version banner (it reads `packages/data/manifest.json`), commit, deploy.

## What needs a human, and what does not

| Step | Automated |
|---|---|
| Extracting data values | Yes |
| Detecting `lastActualYear` | Yes — from literal-vs-formula, not a heuristic |
| Spotting new VBA rule functions | Yes — the diff shows them |
| Writing the TypeScript for a new rule | No |
| Deciding a rule change is faithfully ported | No |
| Producing golden files from Excel | No — needs Windows and Excel |
