Attribute VB_Name = "mdlDefalut_value"
Option Explicit

Sub get_economical_Default_Value()

    Dim lng_Yearly_Inflation_Default As Double
    Dim lng_Real_Growth_Default As Double
    Dim lng_FondAvkastning_Default As Double
    
    lng_Yearly_Inflation_Default = Application.Range("rng_Yearly_Inflation_Default").Value
    lng_Real_Growth_Default = Application.Range("rng_Real_Growth_Default").Value
    lng_FondAvkastning_Default = Application.Range("rng_FondAvkastning_Default").Value
    
    If Application.Range("syslang").Value = 0 Then
            MsgBox ("Default värdena utgår från att beräkningarna anges i dagens löneläge")
    Else
            MsgBox ("Default values assume that calculation are in today's payroll")
    End If
    
    With wsStart
        .Range("rng_Yearly_Inflation") = lng_Yearly_Inflation_Default
        .Range("rng_Real_Growth") = lng_Real_Growth_Default
        .Range("rng_FondAvkastning") = lng_FondAvkastning_Default
    End With
End Sub

Sub Show_defaultvalues()
    frm_DefaultValues.Show
End Sub
Sub BeginSaveAsMydefault(ByVal strTypfallName As String)
    If Len(Trim(strTypfallName)) > 0 Then
        SaveAsMyDefault (strTypfallName)
        Application.Calculate
        frmTypfalllist.lstTypfalllist.RowSource = "rng_Typfalllist"
    End If
End Sub
Sub SaveAsMyDefault(ByVal strName As String)
    If Len(Trim(strName)) = 0 Then Exit Sub
    Dim lngCnMyTyfall As Integer
    Dim dblOffsetrow As Double

    Dim x
    Dim r As Range
    Dim rngTopLeftCell As Range
    Dim rngTypfallcell As Range
    
    lngCnMyTyfall = Application.Range("rng_Top_Left_AntalEgentypfall")
    Set rngTopLeftCell = Application.Range("rng_Top_Left_EgentypfallList").Offset(1)
    Set r = rngTopLeftCell.Resize(rngTopLeftCell.CurrentRegion.Rows.Count - 1, 1)

    On Error Resume Next
    x = Application.Match(strName, r, 0)
    If Not IsNumeric(x) Or x < 1 Then
        On Error GoTo 0
        Set rngTypfallcell = Application.Range("rng_Top_Left_EgentypfallList").Offset(lngCnMyTyfall + 1)
        rngTypfallcell = strName
        Set r = Nothing
    End If
    On Error GoTo 0
        
    Set rngTopLeftCell = Application.Range("rng_Top_Left_Egen_Standard").Offset(1)
    Set r = rngTopLeftCell.Resize(rngTopLeftCell.CurrentRegion.Rows.Count - 2, 1)

    On Error Resume Next
    x = Application.Match(strName, r, 0)
    If Not IsNumeric(x) Or x < 1 Then
        Set rngTypfallcell = Application.Range("rng_Top_Left_Egen_Standard").Offset(lngCnMyTyfall + 1)
        Set r = Nothing
    Else
        Set rngTypfallcell = rngTopLeftCell.Offset(x - 1)
        Set r = Nothing
    End If
    On Error GoTo 0
    
    DoEvents
    rngTypfallcell = strName
    With rngTypfallcell
        .Offset(, 1).Value = Application.Range("BornYear").Value
        .Offset(, 2).Value = Application.Range("PARYear").Value
        .Offset(, 3).Value = Application.Range("wStartYear").Value
        .Offset(, 4).Value = Application.Range("Wage_Monthly").Value
        .Offset(, 5).Value = Application.Range("rng_Yearly_Inflation").Value
        .Offset(, 6).Value = Application.Range("rng_Real_Growth").Value
        .Offset(, 7).Value = Application.Range("rng_FondAvkastning").Value
        .Offset(, 8).Value = Application.Range("IPS_Monthly").Value
        .Offset(, 9).Value = Application.Range("rng_TJP_Listbox").Value
    End With
    
    Set rngTopLeftCell = Nothing
    Set r = Nothing
    Set rngTypfallcell = Nothing

End Sub
Sub BeginGetMyDefault()
    frmTypfalllist.Show
End Sub

Sub GetMyDefault(ByVal strName As String)
    If Len(Trim(strName)) = 0 Then Exit Sub
    Dim x
    Dim r As Range
    Dim rngTopLeftCell As Range
    
    Set rngTopLeftCell = Application.Range("rng_Top_Left_Egen_Standard").Offset(1)
    Set r = rngTopLeftCell.Resize(rngTopLeftCell.CurrentRegion.Rows.Count - 2, 1)

    On Error Resume Next
    x = Application.Match(strName, r, 0)
    If Not IsNumeric(x) Or x < 1 Then
        On Error GoTo 0
        Set rngTopLeftCell = Nothing
        Set r = Nothing
        Application_Wakeup
        Exit Sub
    End If
    On Error GoTo 0
    DoEvents
    With Application
        .Range("BornYear").Value = Application.Range("rng_Top_Left_Egen_Standard").Offset(x, 1).Value
        .Range("PARYear").Value = Application.Range("rng_Top_Left_Egen_Standard").Offset(x, 2).Value
        .Range("wStartYear").Value = Application.Range("rng_Top_Left_Egen_Standard").Offset(x, 3).Value
        .Range("Wage_Monthly").Value = Application.Range("rng_Top_Left_Egen_Standard").Offset(x, 4).Value
        .Range("rng_Yearly_Inflation").Value = Application.Range("rng_Top_Left_Egen_Standard").Offset(x, 5).Value
        .Range("rng_Real_Growth").Value = Application.Range("rng_Top_Left_Egen_Standard").Offset(x, 6).Value
        .Range("rng_FondAvkastning").Value = Application.Range("rng_Top_Left_Egen_Standard").Offset(x, 7).Value
        .Range("IPS_Monthly").Value = Application.Range("rng_Top_Left_Egen_Standard").Offset(x, 8).Value
        .Range("rng_TJP_Listbox").Value = Application.Range("rng_Top_Left_Egen_Standard").Offset(x, 9).Value
    End With
    Set rngTopLeftCell = Nothing
    Set r = Nothing
End Sub
Sub RemoveTypfall(ByVal strName As String)
    If Len(Trim(strName)) = 0 Then Exit Sub
    Dim x
    Dim r As Range
    Dim rngTopLeftCell As Range
    Dim rngRemovecell As Range
    
    Set rngTopLeftCell = Application.Range("rng_Top_Left_EgentypfallList").Offset(1)
    Set r = rngTopLeftCell.Resize(rngTopLeftCell.CurrentRegion.Rows.Count - 1, 1)

    On Error Resume Next
    x = Application.Match(strName, r, 0)
    If Not IsNumeric(x) Or x < 1 Then
        On Error GoTo 0
        Set rngTopLeftCell = Nothing
        Set r = Nothing
        Application_Wakeup
        Exit Sub
    End If
    On Error GoTo 0
    
    Set rngRemovecell = Application.Range("rng_Top_Left_EgentypfallList").Offset(x)
    rngRemovecell.Delete Shift:=xlUp
        
    
    Set rngTopLeftCell = Application.Range("rng_Top_Left_Egen_Standard").Offset(1)
    Set r = rngTopLeftCell.Resize(rngTopLeftCell.CurrentRegion.Rows.Count - 2, 1)

    On Error Resume Next
    x = Application.Match(strName, r, 0)
    If Not IsNumeric(x) Or x < 1 Then
        On Error GoTo 0
        Set rngTopLeftCell = Nothing
        Set r = Nothing
        Application_Wakeup
        Exit Sub
    End If
    On Error GoTo 0
    DoEvents
    
    Set rngRemovecell = Application.Range("rng_Top_Left_Egen_Standard").Offset(x)
    Set rngRemovecell = rngRemovecell.Resize(1, rngRemovecell.CurrentRegion.Columns.Count)
    rngRemovecell.Delete Shift:=xlUp

    Set rngTopLeftCell = Nothing
    Set r = Nothing
    Set rngRemovecell = Nothing
End Sub

Sub Put_DefaultValue_AdvSetting()
    Dim rng As Range
    Dim rngSearch As Range
    Dim iCol As Integer
    Dim lLastRow As Long
    Dim c As Range
    Dim strNamerange As String
    Dim strMessage As String
    
    If RangeNameExists("rng_Top_default_Value") = False Or RangeNameExists("w_time") = False Then
        MsgBox "Minst en av följande namerange saknas:" & chr(10) & "rng_Top_default_Value" & chr(10) & "w_time"
        Exit Sub
    End If
    
    Application_Rest
    Set rng = Application.Range("rng_Top_default_Value")
    iCol = Application.Range("IPS_Monthly").Column
    With ActiveSheet
        lLastRow = fnLast(1, .Range(.Cells(1, iCol), .Cells(1000000, iCol)))
        Set rng = rng.Resize(lLastRow, 1)
        For Each c In rng
            strNamerange = c.Offset(, 1)
            If Trim(c) <> "" Or Trim(strNamerange) <> "" Then
                If LCase(Left(strNamerange, 4)) = "cell" Then
                    .Range(Trim(Mid(strNamerange, 5, Len(strNamerange) - 4))) = Trim(c)
                ElseIf Len(Trim(strNamerange)) > 0 Then
                   If RangeNameExists(strNamerange) Then
                        If Trim(LCase(strNamerange)) = "sant" Then
                            Application.Range(strNamerange) = True
                        ElseIf Trim(LCase(Left(strNamerange, 4))) = "falsk" Then
                            Application.Range(strNamerange) = False
                        'För att kopiera formler för angivna variabler - Tillagt 250613
                        ElseIf c.HasFormula = True Then
                            .Cells(c.Row, iCol).Formula = c.Formula
                        Else
                            Application.Range(strNamerange) = Trim(c)
                        End If
                    Else
                        strMessage = strMessage & chr(10) & " Namerange: " & strNamerange & " (rad: " & c.Row & ") finns inte."
                    End If
                ElseIf c.HasFormula = True Then
                    .Cells(c.Row, iCol).Formula = c.Formula
                ElseIf Left(c, 1) = chr(34) Then
                    .Cells(c.Row, iCol) = Mid(c, 2, 1)
                Else
                    .Cells(c.Row, iCol) = c
                End If
            End If
            'Något konstigt c ger 0,x inte 0.1 varför knas och abs funkar inte ? Tillagt 250512
            If (c > 0 And c < 1) Or (c < 0 And c < 1) Then
            'Stop
                .Cells(c.Row, iCol) = c
            End If
        Next c
    End With
    Set rng = Nothing
    Set rngSearch = Nothing
    Application_Wakeup
    If Len(strMessage) > 0 Then MsgBox strMessage
End Sub

Function fnLast(Typ As Long, rng As Range)
' 1 = Last row
' 2 = Last column
' 3 = Last cell
    Dim lRow As Long
    Dim lCol As Long

    Select Case Typ

    Case 1:
        On Error Resume Next
        fnLast = rng.Find(What:="*", _
                        After:=rng.Cells(1), _
                        Lookat:=xlPart, _
                        LookIn:=xlFormulas, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlPrevious, _
                        MatchCase:=False).Row
        On Error GoTo 0

    Case 2:
        On Error Resume Next
        fnLast = rng.Find(What:="*", _
                        After:=rng.Cells(1), _
                        Lookat:=xlPart, _
                        LookIn:=xlFormulas, _
                        SearchOrder:=xlByColumns, _
                        SearchDirection:=xlPrevious, _
                        MatchCase:=False).Column
        On Error GoTo 0

    Case 3:
        On Error Resume Next
        lRow = rng.Find(What:="*", _
                       After:=rng.Cells(1), _
                       Lookat:=xlPart, _
                       LookIn:=xlFormulas, _
                       SearchOrder:=xlByRows, _
                       SearchDirection:=xlPrevious, _
                       MatchCase:=False).Row
        On Error GoTo 0

        On Error Resume Next
        lCol = rng.Find(What:="*", _
                        After:=rng.Cells(1), _
                        Lookat:=xlPart, _
                        LookIn:=xlFormulas, _
                        SearchOrder:=xlByColumns, _
                        SearchDirection:=xlPrevious, _
                        MatchCase:=False).Column
        On Error GoTo 0

        On Error Resume Next
        fnLast = rng.Parent.Cells(lRow, lCol).Address(False, False)
        If err.Number > 0 Then
            fnLast = rng.Cells(1).Address(False, False)
            err.Clear
        End If
        On Error GoTo 0

    End Select
End Function

