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
npm run compare       # engine vs. the Excel golden files in reference/
```

The fixtures in `reference/fixtures/` are regenerated from the new workbook, so they check the
new data. The golden files in `reference/golden/` are **not** — they came from a previous Excel
run. Regenerate them from the new workbook too (see `reference/README.md`), otherwise the
comparison is checking the new engine against last year's rules.

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
