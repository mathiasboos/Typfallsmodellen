Attribute VB_Name = "ExportGoldenCases"
'==============================================================================
' Exports reference results from the real Typfallsmodellen for the web port.
'
' Fills the Mikrosim sheet with a spread of typfall, runs the model's own batch
' runner over them, and writes inputs and outputs to a CSV.
'
' See reference/golden/HOWTO.md in the web port's repository for how to use it.
'
' The cases deliberately straddle rule boundaries -- cohorts either side of 1938
' and 1954 for ATP, salaries around the garantipension phase-out and the state
' tax threshold, every occupational pension agreement -- because coverage is
' what makes a golden file worth having, not row count.
'
' Retirement ages are clamped to each cohort's lowest permitted age, read from
' the Nyckeltal sheet. Below that the model opens a Yes/No dialog mid-run and,
' if answered Yes, silently changes the case -- which would put an input in the
' file that did not produce the output beside it.
'==============================================================================
Option Explicit

Private Const MIKROSIM_FIRST_DATA_ROW As Long = 8
Private Const COL_FIRST_INPUT As Long = 2      ' B: Födelseår
Private Const COL_LAST_INPUT As Long = 11      ' K: Välj tjänstepension
Private Const COL_FIRST_OUTPUT As Long = 13    ' M: Slutlön
Private Const COL_LAST_OUTPUT As Long = 24     ' X: Disponibel inkomst
Private Const CELL_FROM_ROW As String = "P3"
Private Const CELL_UNTIL_ROW As String = "U3"

Private mCases() As Variant
Private mCaseCount As Long


Public Sub ExportGoldenCases()
    Dim ws As Worksheet
    Dim savePath As String
    Dim i As Long, c As Long

    Set ws = ThisWorkbook.Worksheets("Mikrosim")

    savePath = Application.GetSaveAsFilename( _
        InitialFileName:="golden-cases.csv", _
        FileFilter:="CSV files (*.csv), *.csv", _
        Title:="Save reference results as")
    If VarType(savePath) = vbBoolean Then Exit Sub    ' user cancelled

    BuildCases
    If mCaseCount = 0 Then
        MsgBox "No cases were generated.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' Clear anything already on the sheet, inputs and outputs alike, so a
    ' shorter run cannot leave a previous run's results behind.
    ws.Range(ws.Cells(MIKROSIM_FIRST_DATA_ROW, COL_FIRST_INPUT), _
             ws.Cells(ws.Rows.Count, COL_LAST_OUTPUT)).ClearContents

    For i = 1 To mCaseCount
        For c = 0 To COL_LAST_INPUT - COL_FIRST_INPUT
            ws.Cells(MIKROSIM_FIRST_DATA_ROW + i - 1, COL_FIRST_INPUT + c).Value = mCases(c + 1, i)
        Next c
    Next i

    ws.Range(CELL_FROM_ROW).Value = MIKROSIM_FIRST_DATA_ROW
    ws.Range(CELL_UNTIL_ROW).Value = MIKROSIM_FIRST_DATA_ROW + mCaseCount - 1

    Application.ScreenUpdating = True

    ' The model's own batch runner: reads each input row, runs Mcalc, writes the
    ' results beside it.
    InputXGetY

    WriteCsv savePath, ws, mCaseCount
    MsgBox mCaseCount & " cases exported to" & vbCrLf & savePath, vbInformation
End Sub


'--- case design -------------------------------------------------------------

Private Sub AddCase(ByVal bornYear As Long, ByVal startWorkAge As Long, ByVal retireAge As Long, _
                    ByVal annualSalary As Double, ByVal inflation As Double, ByVal realGrowth As Double, _
                    ByVal realReturn As Double, ByVal occupationalScheme As Long)
    Dim lowest As Long
    lowest = LowestRetirementAge(bornYear)
    If retireAge < lowest Then retireAge = lowest
    If startWorkAge >= retireAge Then Exit Sub

    mCaseCount = mCaseCount + 1
    ReDim Preserve mCases(1 To 10, 1 To mCaseCount)

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


Private Sub BuildCases()
    Dim bornYears As Variant, salaries As Variant, retireAges As Variant, startAges As Variant
    Dim b As Long, s As Long, r As Long, a As Long, scheme As Long

    Erase mCases
    mCaseCount = 0
    ReDim mCases(1 To 10, 1 To 1)

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


'--- helpers -----------------------------------------------------------------

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


' Formats a number for CSV with a dot as the decimal separator, whatever the
' machine's locale. A Swedish Excel writes "1,7", which would split the field.
Private Function CsvNum(ByVal value As Variant) As String
    If Not IsNumeric(value) Then
        CsvNum = ""
        Exit Function
    End If
    CsvNum = Replace(Format$(CDbl(value), "0.##########"), ",", ".")
End Function


Private Sub WriteCsv(ByVal path As String, ByVal ws As Worksheet, ByVal caseCount As Long)
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
    ' Read by row rather than by name: rng_Forenklad_berakning and
    ' rng_Forsakringstid_vid_65 are spelled with a-ring and o-umlaut in the
    ' workbook, which a .bas file's encoding can mangle. Column 9 of
    ' Adv_settings carries the variable name, so the rows are checkable.
    Print #f, "# forenklad_pp: " & CsvNum(adv.Cells(49, 2).Value) & "   (" & adv.Cells(49, 9).Value & ")"
    Print #f, "# forsakringstid: " & CsvNum(adv.Cells(21, 2).Value) & "   (" & adv.Cells(21, 9).Value & ")"
    Print #f, "# hyra: " & CsvNum(Application.Range("Hyra").Value)
    Print #f, "# ansokt_bt: " & CsvNum(Application.Range("Rng_Ansokt").Value)

    ' Header row, taken from the sheet so it tracks any future column change.
    line = ""
    For c = COL_FIRST_INPUT To COL_LAST_INPUT
        line = line & Replace(Trim$(CStr(ws.Cells(7, c).Value)), ",", " ") & ","
    Next c
    For c = COL_FIRST_OUTPUT To COL_LAST_OUTPUT
        line = line & Replace(Trim$(CStr(ws.Cells(7, c).Value)), ",", " ")
        If c < COL_LAST_OUTPUT Then line = line & ","
    Next c
    Print #f, line

    For i = 0 To caseCount - 1
        line = ""
        For c = COL_FIRST_INPUT To COL_LAST_INPUT
            line = line & CsvNum(ws.Cells(MIKROSIM_FIRST_DATA_ROW + i, c).Value) & ","
        Next c
        For c = COL_FIRST_OUTPUT To COL_LAST_OUTPUT
            line = line & CsvNum(ws.Cells(MIKROSIM_FIRST_DATA_ROW + i, c).Value)
            If c < COL_LAST_OUTPUT Then line = line & ","
        Next c
        Print #f, line
    Next i

    Close #f
End Sub
