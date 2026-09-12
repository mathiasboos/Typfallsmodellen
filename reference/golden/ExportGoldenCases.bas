Attribute VB_Name = "ExportGoldenCases"
'==============================================================================
' Exports reference results from the real Typfallsmodellen for the web port.
'
' Fills the Mikrosim sheet with a spread of typfall, runs the model's own batch
' runner over them, and writes inputs and outputs to a CSV.
'
' See reference/golden/HOWTO.md in the web port's repository for how to use it.
'
' Entry points:
'
'   ExportGoldenCasesQuick      61 cases, about half an hour. The rule
'                               boundaries that carry the most information.
'   ExportGoldenCases           all 295 cases. A couple of hours.
'   ExportGoldenCasesResume     keeps the inputs already on the sheet and runs
'                               only the rows without results.
'   ExportGoldenCasesFromSheet  writes the CSV from what is on the sheet,
'                               recomputing nothing.
'   ReportMikrosimState         says what is actually on the sheet. Start here
'                               when something looks wrong.
'
' The CSV is rewritten after every chunk, not once at the end, so a halt, a
' crash or a workbook closed without saving still leaves a complete file for
' every case that finished. An earlier version wrote it only on the way out,
' and a halted run left two hours of work in an unsaved workbook.
'
' The cases deliberately straddle rule boundaries -- cohorts either side of 1938
' and 1954 for ATP, salaries around the garantipension phase-out and the state
' tax threshold, every occupational pension agreement -- because coverage is
' what makes a golden file worth having, not row count.
'
' Retirement ages are clamped to each cohort's lowest permitted age, read from
' the Nyckeltal sheet. Below that the model opens a Yes/No dialog mid-run and,
' if answered Yes, silently changes the case -- which would put an input in the
' file that did not produce the output beside it. It also trips Mcalc's guard at
' VBA_go.bas:529, which exits without restoring automatic calculation.
'==============================================================================
Option Explicit

' The sheet's geometry is read from the workbook's own defined names. These are
' what those names resolve to in the 2025 workbook, used only if a name is
' missing -- so a renamed range degrades to the layout we know rather than to a
' silent wrong answer.
Private Const FALLBACK_HEADER_ROW As Long = 7           ' rngXTopleft is B7
Private Const FALLBACK_FIRST_INPUT_COL As Long = 2      ' B: Fodelsear
Private Const FALLBACK_FIRST_OUTPUT_COL As Long = 13    ' M: rngYtopleft is M7
Private Const FALLBACK_FROM_ROW As String = "P3"        ' rngExecuteFromRow
Private Const FALLBACK_UNTIL_ROW As String = "U3"       ' rngExecuteUntilRow

Private Const INPUT_COLUMNS As Long = 10                ' B to K
Private Const OUTPUT_COLUMNS As Long = 12               ' M to X

' Rows per call to the batch runner. Small enough that a halt costs little and
' the CSV is refreshed often, large enough that the per-call overhead stays out
' of the way.
Private Const CHUNK_ROWS As Long = 25

' A bound on every scan down the sheet, so an odd sheet state cannot walk to
' row a million.
Private Const MAX_SCAN_ROWS As Long = 100000

' Where the inputs and results sit on the Mikrosim sheet.
Private Type MikrosimLayout
    HeaderRow As Long
    FirstDataRow As Long
    FirstInputCol As Long
    LastInputCol As Long
    FirstOutputCol As Long
    LastOutputCol As Long
    ResolvedFromNames As Boolean
End Type

Private mCases() As Variant
Private mCaseCount As Long
Private mCaseSet As String


'--- entry points -------------------------------------------------------------

Public Sub ExportGoldenCasesQuick()
    RunExport "quick"
End Sub


Public Sub ExportGoldenCases()
    RunExport "full"
End Sub


'==============================================================================
' Runs the rows that have inputs but no results yet.
'
' Nothing is cleared and no cases are generated: whatever is on the sheet stays.
' Use it after a halt, or to work through the full set in sittings. The workbook
' has to have been saved for the inputs to still be there in a new session.
'==============================================================================
Public Sub ExportGoldenCasesResume()
    Dim ws As Worksheet
    Dim layout As MikrosimLayout
    Dim savePath As String
    Dim withInputs As Long, withOutputs As Long

    Set ws = ThisWorkbook.Worksheets("Mikrosim")
    layout = ResolveLayout()
    CountRows ws, layout, withInputs, withOutputs

    If withInputs = 0 Then
        MsgBox "There are no inputs on the Mikrosim sheet to resume from." & vbCrLf & vbCrLf & _
               "Run ExportGoldenCasesQuick or ExportGoldenCases to start a run, or " & _
               "ReportMikrosimState to see what is on the sheet.", vbExclamation
        Exit Sub
    End If

    If withOutputs >= withInputs Then
        MsgBox "All " & withInputs & " rows already have results." & vbCrLf & _
               "Run ExportGoldenCasesFromSheet to write the CSV.", vbInformation
        Exit Sub
    End If

    savePath = AskForPath()
    If Len(savePath) = 0 Then Exit Sub

    RunChunks ws, layout, savePath, _
              layout.FirstDataRow + withOutputs, layout.FirstDataRow + withInputs - 1, withInputs

    CountRows ws, layout, withInputs, withOutputs
    MsgBox withOutputs & " of " & withInputs & " cases now have results." & vbCrLf & _
           savePath, vbInformation
End Sub


'==============================================================================
' Writes the CSV from whatever is already on the Mikrosim sheet.
'
' Nothing is recomputed. Rows with no result are left out.
'==============================================================================
Public Sub ExportGoldenCasesFromSheet()
    Dim ws As Worksheet
    Dim layout As MikrosimLayout
    Dim savePath As String
    Dim withInputs As Long, withOutputs As Long

    Set ws = ThisWorkbook.Worksheets("Mikrosim")
    layout = ResolveLayout()
    CountRows ws, layout, withInputs, withOutputs

    If withOutputs = 0 Then
        MsgBox EmptySheetMessage(layout, withInputs), vbExclamation
        Exit Sub
    End If

    savePath = AskForPath()
    If Len(savePath) = 0 Then Exit Sub

    WriteCsv savePath, ws, layout, withOutputs
    MsgBox withOutputs & " completed cases exported to" & vbCrLf & savePath, vbInformation
End Sub


'==============================================================================
' Says what is actually on the Mikrosim sheet.
'
' Run this first when an export reports nothing to write. It resolves the same
' geometry the export uses and reports what it finds there, so the answer comes
' from the sheet rather than from guessing at a distance.
'==============================================================================
Public Sub ReportMikrosimState()
    Dim ws As Worksheet
    Dim layout As MikrosimLayout
    Dim withInputs As Long, withOutputs As Long
    Dim msg As String
    Dim r As Long, c As Long
    Dim sample As String

    Set ws = ThisWorkbook.Worksheets("Mikrosim")
    layout = ResolveLayout()
    CountRows ws, layout, withInputs, withOutputs

    msg = "Mikrosim sheet" & vbCrLf & vbCrLf
    msg = msg & "Layout " & IIf(layout.ResolvedFromNames, _
                "(from the workbook's defined names):", _
                "(FALLBACK -- a defined name is missing):") & vbCrLf
    msg = msg & "  header row      " & layout.HeaderRow & vbCrLf
    msg = msg & "  first data row  " & layout.FirstDataRow & vbCrLf
    msg = msg & "  inputs          " & ColLetter(layout.FirstInputCol) & " to " & _
                ColLetter(layout.LastInputCol) & vbCrLf
    msg = msg & "  results         " & ColLetter(layout.FirstOutputCol) & " to " & _
                ColLetter(layout.LastOutputCol) & vbCrLf & vbCrLf

    msg = msg & "Used range: " & ws.UsedRange.Address & vbCrLf & vbCrLf
    msg = msg & "Rows with an input:  " & withInputs & vbCrLf
    msg = msg & "Rows with a result:  " & withOutputs & vbCrLf & vbCrLf

    If withInputs = 0 Then
        msg = msg & "The sheet is empty. A run writes its results here as it goes, but " & _
              "they are only kept if the workbook is saved -- closing without saving " & _
              "loses them. Start a fresh run with ExportGoldenCasesQuick."
    ElseIf withOutputs = 0 Then
        msg = msg & "The inputs are there but nothing has been computed yet. " & _
              "ExportGoldenCasesResume will run them."
    ElseIf withOutputs < withInputs Then
        msg = msg & "A run stopped part way. ExportGoldenCasesResume will finish it; " & _
              "ExportGoldenCasesFromSheet will export the " & withOutputs & " that are done."
    Else
        msg = msg & "Every row has a result. ExportGoldenCasesFromSheet will write the CSV."
    End If

    ' The first two data rows in full, so a shifted column shows up immediately.
    For r = layout.FirstDataRow To layout.FirstDataRow + 1
        sample = sample & vbCrLf & "Row " & r & ":"
        For c = layout.FirstInputCol To layout.LastOutputCol
            sample = sample & " " & Trim$(CStr(ws.Cells(r, c).Value))
        Next c
    Next r
    Debug.Print msg
    Debug.Print sample

    MsgBox msg & vbCrLf & vbCrLf & "The first rows are in the Immediate window (Ctrl+G).", _
           vbInformation
End Sub


'--- the run ------------------------------------------------------------------

Private Sub RunExport(ByVal caseSet As String)
    Dim ws As Worksheet
    Dim layout As MikrosimLayout
    Dim savePath As String
    Dim i As Long, c As Long

    Set ws = ThisWorkbook.Worksheets("Mikrosim")
    layout = ResolveLayout()

    savePath = AskForPath()
    If Len(savePath) = 0 Then Exit Sub

    BuildCases caseSet
    If mCaseCount = 0 Then
        MsgBox "No cases were generated.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' Clear anything already on the sheet, inputs and results alike, so a
    ' shorter run cannot leave a previous run's results behind.
    '
    ' Bounded deliberately. Clearing to ws.Rows.Count is about 24 million cells,
    ' which is slow and leaves Excel with a very large dirty range -- and the
    ' model's own calculation watchdog allows only 0.2 seconds for the sheet to
    ' settle before it hits a Stop (mdlIndataInputOutput.bas:33).
    ClearPreviousRun ws, layout, mCaseCount

    For i = 1 To mCaseCount
        For c = 0 To INPUT_COLUMNS - 1
            ws.Cells(layout.FirstDataRow + i - 1, layout.FirstInputCol + c).Value = mCases(c + 1, i)
        Next c
    Next i

    Application.ScreenUpdating = True

    RunChunks ws, layout, savePath, _
              layout.FirstDataRow, layout.FirstDataRow + mCaseCount - 1, mCaseCount

    MsgBox mCaseCount & " cases exported to" & vbCrLf & savePath, vbInformation
End Sub


'==============================================================================
' Runs the batch runner over firstRow..lastRow, a chunk at a time.
'
' The CSV is rewritten after every chunk. Rewriting a few hundred rows costs
' milliseconds, and it is what makes a halt cost one chunk rather than the run.
'==============================================================================
Private Sub RunChunks(ByVal ws As Worksheet, ByRef layout As MikrosimLayout, _
                      ByVal savePath As String, ByVal firstRow As Long, _
                      ByVal lastRow As Long, ByVal totalCases As Long)
    Dim chunkStart As Long, chunkEnd As Long
    Dim done As Long

    chunkStart = firstRow
    Do While chunkStart <= lastRow
        chunkEnd = chunkStart + CHUNK_ROWS - 1
        If chunkEnd > lastRow Then chunkEnd = lastRow

        SetRunRange chunkStart, chunkEnd

        ' Mcalc sets calculation to manual on its third line and only restores
        ' it at its normal end (VBA_go.bas:523 and :2834); the guard at :529
        ' exits in between. Putting it back before each chunk keeps one such
        ' case from affecting the rest.
        Application.Calculation = xlCalculationAutomatic

        ' The model's own batch runner: reads each input row, runs Mcalc,
        ' writes the results beside it.
        InputXGetY

        Application.Calculation = xlCalculationAutomatic

        ' Save what we have before starting the next chunk, not at the end.
        done = chunkEnd - layout.FirstDataRow + 1
        WriteCsv savePath, ws, layout, done

        Application.StatusBar = "Typfall " & done & " of " & totalCases & _
                                " done -- CSV saved to " & savePath
        chunkStart = chunkEnd + 1
    Loop

    SetRunRange layout.FirstDataRow, lastRow
    Application.StatusBar = False
End Sub


Private Sub SetRunRange(ByVal firstRow As Long, ByVal lastRow As Long)
    NamedRange("rngExecuteFromRow", FALLBACK_FROM_ROW).Value = firstRow
    NamedRange("rngExecuteUntilRow", FALLBACK_UNTIL_ROW).Value = lastRow
End Sub


'--- the sheet ----------------------------------------------------------------

'==============================================================================
' Where the inputs and results sit, from the workbook's own defined names.
'
' rngXTopleft is the top-left input label and rngYtopleft the first result
' label; the batch runner itself reads both (mdlIndataInputOutput.bas:97 and
' :189), so following them means the export cannot drift from the runner. In the
' 2025 workbook they are B7 and M7.
'==============================================================================
Private Function ResolveLayout() As MikrosimLayout
    Dim layout As MikrosimLayout
    Dim x As Range, y As Range

    layout.ResolvedFromNames = True

    Set x = NamedRange("rngXTopleft", "")
    If x Is Nothing Then
        layout.HeaderRow = FALLBACK_HEADER_ROW
        layout.FirstInputCol = FALLBACK_FIRST_INPUT_COL
        layout.ResolvedFromNames = False
    Else
        layout.HeaderRow = x.Row
        layout.FirstInputCol = x.Column
    End If

    Set y = NamedRange("rngYtopleft", "")
    If y Is Nothing Then
        layout.FirstOutputCol = FALLBACK_FIRST_OUTPUT_COL
        layout.ResolvedFromNames = False
    Else
        layout.FirstOutputCol = y.Column
    End If

    layout.FirstDataRow = layout.HeaderRow + 1
    layout.LastInputCol = layout.FirstInputCol + INPUT_COLUMNS - 1
    layout.LastOutputCol = layout.FirstOutputCol + OUTPUT_COLUMNS - 1

    ResolveLayout = layout
End Function


' A defined range, or the fallback address, or Nothing when neither is given.
Private Function NamedRange(ByVal name As String, ByVal fallback As String) As Range
    On Error Resume Next
    Set NamedRange = ThisWorkbook.Names(name).RefersToRange
    On Error GoTo 0
    If NamedRange Is Nothing And Len(fallback) > 0 Then
        Set NamedRange = ThisWorkbook.Worksheets("Mikrosim").Range(fallback)
    End If
End Function


'==============================================================================
' How many rows carry an input, and how many of those also carry a result.
'
' Counted separately and deliberately: a sheet with no inputs at all and a sheet
' whose run never finished need different answers, and an earlier version
' reported both as "columns M to X are empty", which sent the diagnosis in the
' wrong direction.
'==============================================================================
Private Sub CountRows(ByVal ws As Worksheet, ByRef layout As MikrosimLayout, _
                      ByRef withInputs As Long, ByRef withOutputs As Long)
    Dim r As Long
    Dim stillComplete As Boolean

    withInputs = 0
    withOutputs = 0
    stillComplete = True

    r = layout.FirstDataRow
    Do While r < layout.FirstDataRow + MAX_SCAN_ROWS
        If Not HasValue(ws, r, layout.FirstInputCol, layout.LastInputCol) Then Exit Do
        withInputs = withInputs + 1
        ' Results are written in row order, so the completed rows are the
        ' unbroken run from the top; a gap ends the count.
        If stillComplete Then
            If HasValue(ws, r, layout.FirstOutputCol, layout.LastOutputCol) Then
                withOutputs = withOutputs + 1
            Else
                stillComplete = False
            End If
        End If
        r = r + 1
    Loop
End Sub


Private Function HasValue(ByVal ws As Worksheet, ByVal r As Long, _
                          ByVal firstCol As Long, ByVal lastCol As Long) As Boolean
    Dim c As Long
    For c = firstCol To lastCol
        If Len(Trim$(CStr(ws.Cells(r, c).Value))) > 0 Then
            HasValue = True
            Exit Function
        End If
    Next c
End Function


Private Function EmptySheetMessage(ByRef layout As MikrosimLayout, _
                                   ByVal withInputs As Long) As String
    If withInputs = 0 Then
        EmptySheetMessage = _
            "The Mikrosim sheet is empty -- no inputs in columns " & _
            ColLetter(layout.FirstInputCol) & " to " & ColLetter(layout.LastInputCol) & _
            " from row " & layout.FirstDataRow & " on." & vbCrLf & vbCrLf & _
            "A run's results are only kept if the workbook is saved, so closing " & _
            "without saving loses them." & vbCrLf & vbCrLf & _
            "Start a fresh run with ExportGoldenCasesQuick, or run " & _
            "ReportMikrosimState for a fuller picture."
    Else
        EmptySheetMessage = _
            withInputs & " rows have inputs, but none has a result in columns " & _
            ColLetter(layout.FirstOutputCol) & " to " & ColLetter(layout.LastOutputCol) & _
            "." & vbCrLf & vbCrLf & _
            "ExportGoldenCasesResume will compute them."
    End If
End Function


' Clears the input and result block, bounded to the rows that could hold data:
' whatever the sheet already used, or this run's case count, whichever is more.
Private Sub ClearPreviousRun(ByVal ws As Worksheet, ByRef layout As MikrosimLayout, _
                             ByVal caseCount As Long)
    Dim lastRow As Long
    lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If lastRow < layout.FirstDataRow + caseCount Then
        lastRow = layout.FirstDataRow + caseCount
    End If
    lastRow = lastRow + 10      ' a margin, in case the used range under-reports

    ws.Range(ws.Cells(layout.FirstDataRow, layout.FirstInputCol), _
             ws.Cells(lastRow, layout.LastOutputCol)).ClearContents
End Sub


'--- case design --------------------------------------------------------------

Private Sub AddCase(ByVal bornYear As Long, ByVal startWorkAge As Long, ByVal retireAge As Long, _
                    ByVal annualSalary As Double, ByVal inflation As Double, ByVal realGrowth As Double, _
                    ByVal realReturn As Double, ByVal occupationalScheme As Long)
    Dim lowest As Long
    lowest = LowestRetirementAge(bornYear)
    If retireAge < lowest Then retireAge = lowest
    If startWorkAge >= retireAge Then Exit Sub

    mCaseCount = mCaseCount + 1
    ReDim Preserve mCases(1 To INPUT_COLUMNS, 1 To mCaseCount)

    mCases(1, mCaseCount) = bornYear
    mCases(2, mCaseCount) = startWorkAge
    mCases(3, mCaseCount) = retireAge
    mCases(4, mCaseCount) = annualSalary
    mCases(5, mCaseCount) = inflation
    mCases(6, mCaseCount) = realGrowth
    mCases(7, mCaseCount) = realReturn
    mCases(8, mCaseCount) = 0                      ' no own wage list
    mCases(9, mCaseCount) = 0                      ' no private saving
    mCases(10, mCaseCount) = occupationalScheme
End Sub


Private Sub BuildCases(ByVal caseSet As String)
    Erase mCases
    mCaseCount = 0
    mCaseSet = caseSet
    ReDim mCases(1 To INPUT_COLUMNS, 1 To 1)

    If caseSet = "quick" Then
        BuildQuickCases
    Else
        BuildFullCases
    End If
End Sub


'==============================================================================
' 61 cases: the boundaries that carry the most information, in about half an
' hour rather than a couple of hours.
'
' The blocks are the same shape as the full set's, so the comparison report
' groups them the same way. Keep them in step with QUICK_BLOCKS in
' packages/engine/test/golden/harness.ts.
'==============================================================================
Private Sub BuildQuickCases()
    Dim bornYears As Variant, salaries As Variant, retireAges As Variant
    Dim b As Long, s As Long, r As Long, a As Long, scheme As Long

    ' A (32): the two ATP boundaries, both sides, against every agreement.
    bornYears = Array(1937, 1938, 1953, 1954)
    For b = LBound(bornYears) To UBound(bornYears)
        For scheme = 1 To 8
            AddCase bornYears(b), 23, 66, 462000, 0, 0, 0.017, scheme
        Next scheme
    Next b

    ' B (12): salary against retirement age on the shipped cohort, spanning the
    ' garantipension phase-out and the state tax threshold.
    salaries = Array(180000, 462000, 660000, 1080000)
    retireAges = Array(63, 66, 70)
    For s = LBound(salaries) To UBound(salaries)
        For r = LBound(retireAges) To UBound(retireAges)
            AddCase 1959, 23, retireAges(r), salaries(s), 0, 0, 0.017, 4
        Next r
    Next s

    ' C (8): entry age against salary, on cohorts with a full new-rules career.
    bornYears = Array(1970, 1990)
    salaries = Array(264000, 840000)
    For b = LBound(bornYears) To UBound(bornYears)
        For a = 0 To 1
            For s = LBound(salaries) To UBound(salaries)
                AddCase bornYears(b), IIf(a = 0, 20, 30), 67, salaries(s), 0, 0, 0.017, 2
            Next s
        Next a
    Next b

    ' D (4): away from the forecasting standard. Inflation, real growth and
    ' return each move the answer through a different path.
    AddCase 1959, 23, 66, 462000, 0.02, 0, 0.017, 4
    AddCase 1959, 23, 66, 462000, 0, 0.016, 0.017, 4
    AddCase 1959, 23, 66, 462000, 0, 0, 0.035, 4
    AddCase 1959, 23, 66, 462000, 0.02, 0.016, 0.035, 4

    ' E (4): low and high earners, where garantipension and the state tax
    ' thresholds bite.
    AddCase 1945, 25, 66, 180000, 0, 0, 0.017, 1
    AddCase 1945, 20, 66, 1080000, 0, 0, 0.017, 3
    AddCase 1980, 25, 66, 180000, 0, 0, 0.017, 1
    AddCase 1980, 20, 66, 1080000, 0, 0, 0.017, 3

    ' F (1): the worked example from the user manual -- a care assistant born
    ' 1960, working from 20 to 67 on 27 000 kr/month under KAP-KL.
    AddCase 1960, 20, 67, 324000, 0, 0, 0.017, 5
End Sub


Private Sub BuildFullCases()
    Dim bornYears As Variant, salaries As Variant, retireAges As Variant, startAges As Variant
    Dim b As Long, s As Long, r As Long, a As Long, scheme As Long

    ' Cohorts chosen around the ATP boundaries: 1937/1938 (full ATP), and
    ' 1953/1954 (the last cohorts with any tilläggspension).
    bornYears = Array(1937, 1938, 1940, 1945, 1950, 1953, 1954, 1955, 1960, _
                      1965, 1970, 1975, 1980, 1985, 1990, 1995, 2000, 2005)
    ' Salaries around the garantipension phase-out at the bottom and the state
    ' income tax threshold at the top.
    salaries = Array(180000, 264000, 324000, 396000, 462000, 540000, 660000, 840000, 1080000)
    retireAges = Array(63, 65, 66, 67, 68, 70, 75)
    startAges = Array(18, 20, 23, 25, 30)

    ' A: every cohort against every occupational pension agreement.
    For b = LBound(bornYears) To UBound(bornYears)
        For scheme = 1 To 8
            AddCase bornYears(b), 23, 66, 462000, 0, 0, 0.017, scheme
        Next scheme
    Next b

    ' B: salary against retirement age, on the default cohort.
    For s = LBound(salaries) To UBound(salaries)
        For r = LBound(retireAges) To UBound(retireAges)
            AddCase 1959, 23, retireAges(r), salaries(s), 0, 0, 0.017, 4
        Next r
    Next s

    ' C: entry age against salary -- short and long working lives.
    For a = LBound(startAges) To UBound(startAges)
        For s = LBound(salaries) To UBound(salaries)
            AddCase 1970, startAges(a), 67, salaries(s), 0, 0, 0.017, 2
        Next s
    Next a

    ' D: away from the forecasting standard. Inflation, real growth and return
    ' each move the answer through a different path, so vary them apart as well
    ' as together.
    For b = LBound(bornYears) To UBound(bornYears) Step 3
        AddCase bornYears(b), 23, 66, 462000, 0.02, 0, 0.017, 4
        AddCase bornYears(b), 23, 66, 462000, 0, 0.016, 0.017, 4
        AddCase bornYears(b), 23, 66, 462000, 0, 0, 0.035, 4
        AddCase bornYears(b), 23, 66, 462000, 0.02, 0.016, 0.035, 4
    Next b

    ' E: low and high earners across cohorts, where garantipension and the state
    ' tax thresholds bite.
    For b = LBound(bornYears) To UBound(bornYears) Step 2
        AddCase bornYears(b), 25, 66, 180000, 0, 0, 0.017, 1
        AddCase bornYears(b), 20, 66, 1080000, 0, 0, 0.017, 3
    Next b

    ' F: the worked example from the user manual -- a care assistant born 1960,
    ' working from 20 to 67 on 27 000 kr/month under KAP-KL.
    AddCase 1960, 20, 67, 324000, 0, 0, 0.017, 5
End Sub


'--- helpers ------------------------------------------------------------------

Private Function AskForPath() As String
    Dim savePath As Variant
    savePath = Application.GetSaveAsFilename( _
        InitialFileName:="golden-cases.csv", _
        FileFilter:="CSV files (*.csv), *.csv", _
        Title:="Save reference results as")
    If VarType(savePath) = vbBoolean Then
        AskForPath = ""
    Else
        AskForPath = CStr(savePath)
    End If
End Function


' Lowest age this cohort may draw public pension at, from Nyckeltal column 121.
Private Function LowestRetirementAge(ByVal bornYear As Long) As Long
    Dim ws As Worksheet
    Dim value As Variant

    On Error GoTo Fallback
    Set ws = ThisWorkbook.Worksheets("Nyckeltal")
    value = ws.Cells(5 + bornYear - 1930, 121).Value
    If IsNumeric(value) Then
        If value >= 60 And value <= 75 Then
            LowestRetirementAge = CLng(value)
            Exit Function
        End If
    End If
Fallback:
    LowestRetirementAge = 66
End Function


Private Function ColLetter(ByVal col As Long) As String
    Dim n As Long
    n = col
    Do While n > 0
        ColLetter = Chr$(65 + ((n - 1) Mod 26)) & ColLetter
        n = (n - 1) \ 26
    Loop
End Function


' Formats a number for CSV with a dot as the decimal separator, whatever the
' machine's locale. A Swedish Excel writes "1,7", which would split the field.
Private Function CsvNum(ByVal value As Variant) As String
    If Not IsNumeric(value) Then
        CsvNum = ""
        Exit Function
    End If
    CsvNum = Replace(Format$(CDbl(value), "0.##########"), ",", ".")
End Function


' Formats any cell value for the provenance block. Booleans are written as 1/0
' rather than True/False, which a Swedish Excel spells "Sant" and "Falskt".
Private Function CsvValue(ByVal value As Variant) As String
    If VarType(value) = vbBoolean Then
        If value Then CsvValue = "1" Else CsvValue = "0"
    ElseIf IsNumeric(value) Then
        CsvValue = CsvNum(value)
    Else
        CsvValue = Replace(Replace(Trim$(CStr(value)), vbCr, " "), vbLf, " ")
    End If
End Function


' Writes every Adv_settings row that names a variable: column 9 carries the name
' the VBA reads it by, column 2 the value. The row number goes in beside it, so
' a name mangled by the .bas encoding can still be identified.
Private Sub WriteAdvSettings(ByVal f As Integer, ByVal adv As Worksheet)
    Dim lastRow As Long, r As Long
    Dim settingName As String

    lastRow = adv.UsedRange.Row + adv.UsedRange.Rows.Count - 1
    If lastRow > 400 Then lastRow = 400      ' a bound; the sheet ends long before this

    For r = 1 To lastRow
        settingName = Trim$(CStr(adv.Cells(r, 9).Value))
        If Len(settingName) > 0 Then
            Print #f, "# adv." & settingName & ": " & CsvValue(adv.Cells(r, 2).Value) & _
                      "   (row " & r & ")"
        End If
    Next r
End Sub


Private Sub WriteCsv(ByVal path As String, ByVal ws As Worksheet, _
                     ByRef layout As MikrosimLayout, ByVal caseCount As Long)
    Dim f As Integer
    Dim i As Long, c As Long
    Dim line As String
    Dim adv As Worksheet

    Set adv = ThisWorkbook.Worksheets("Adv_settings")
    f = FreeFile
    Open path For Output As #f

    ' Provenance. The comparison needs to know which model version and which
    ' advanced settings produced these numbers.
    Print #f, "# Typfallsmodellen reference results"
    Print #f, "# model: " & ThisWorkbook.Worksheets("Versionsinformation").Range("A2").Value
    Print #f, "# workbook: " & ThisWorkbook.Name
    Print #f, "# exported: " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Print #f, "# referensar: " & CsvNum(Application.Range("w_ref").Value)
    Print #f, "# marginal: " & CsvNum(Application.Range("marginal").Value)
    Print #f, "# belopp12: " & CsvNum(Application.Range("Rng_belopp12").Value)
    Print #f, "# compareTo: " & CsvNum(Application.Range("Rng_CompareTo").Value)
    Print #f, "# avkastning_val: " & CsvNum(Application.Range("rng_Avkastning_val").Value)
    Print #f, "# hyra: " & CsvNum(Application.Range("Hyra").Value)
    Print #f, "# ansokt_bt: " & CsvNum(Application.Range("Rng_Ansokt").Value)
    Print #f, "# cases: " & caseCount
    If Len(mCaseSet) > 0 Then Print #f, "# caseset: " & mCaseSet

    ' Gift lives on the Start sheet, not Adv_settings, and the batch runner never
    ' sets it per row -- so whatever it held applied to all of them. It moves
    ' garantipension, bostadstillagg and every tax row, so it has to be recorded.
    Print #f, "# start.Gift: " & CsvValue(Application.Range("Gift").Value)

    ' And then every setting on Adv_settings, rather than the hand-picked few
    ' above. A run compared against the wrong settings produces confident
    ' nonsense, and which settings matter is not obvious from the outside:
    ' rng_Bara_fastapriser alone rescales every value in the file.
    WriteAdvSettings f, adv

    ' Header row, taken from the sheet so it tracks any future column change.
    line = ""
    For c = layout.FirstInputCol To layout.LastInputCol
        line = line & Replace(Trim$(CStr(ws.Cells(layout.HeaderRow, c).Value)), ",", " ") & ","
    Next c
    For c = layout.FirstOutputCol To layout.LastOutputCol
        line = line & Replace(Trim$(CStr(ws.Cells(layout.HeaderRow, c).Value)), ",", " ")
        If c < layout.LastOutputCol Then line = line & ","
    Next c
    Print #f, line

    For i = 0 To caseCount - 1
        line = ""
        For c = layout.FirstInputCol To layout.LastInputCol
            line = line & CsvNum(ws.Cells(layout.FirstDataRow + i, c).Value) & ","
        Next c
        For c = layout.FirstOutputCol To layout.LastOutputCol
            line = line & CsvNum(ws.Cells(layout.FirstDataRow + i, c).Value)
            If c < layout.LastOutputCol Then line = line & ","
        Next c
        Print #f, line
    Next i

    Close #f
End Sub
