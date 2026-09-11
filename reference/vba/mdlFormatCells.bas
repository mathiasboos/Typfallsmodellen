Attribute VB_Name = "mdlFormatCells"
Option Explicit

Sub MonthCellsFontColor()
    Dim rng As Range
    Call Unprotect_sheet(wsStart.CodeName)
    Set rng = Application.Range("PARMonth")
    If Application.Range("rng_Application_Mode") = 1 Then
        If 0 = 1 Then 'Application.Range("rng_MonthCal") Then
            With rng
                .Locked = False
                .Font.ColorIndex = 0
                With .Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                    xlBetween, Formula1:="=rng_Manader_Index"
                    .InCellDropdown = True
                End With
            End With
            ''Application.Range("w_ref").Font.ColorIndex = 0
        Else
            rng.Clear
            rng.Locked = True
''            Application.Range("w_ref").Font.ColorIndex = 2
            Application.Range("PARMonth").Font.ColorIndex = 2
        End If
        Application.Range("rng_besk_RörligaPriser").Font.ColorIndex = 0
        Application.Range("rng_besk_FastaPriser").Font.ColorIndex = 0
        Application.Range("rng_besk_Figur1").Font.ColorIndex = 0
        'Application.Range("rng_IncomeChartHeader").Font.ColorIndex = 0
        'rng.Value = Application.range("rng_MånaderIndexTop")
        'ChangeBorderColor ("rng_besk_Figur1")
        ChangeBorderColor ("rng_besk_FastaPriser")
        ChangeBorderColor ("rng_besk_RörligaPriser")
    Else
        rng.Clear
        rng.Locked = True
        'Application.range("PARMonth") = Application.range("rng_MånaderIndexTop")
        Call ChangeBorderNoneColor

    End If
    Call Protect_sheet(wsStart.CodeName)
    Set rng = Nothing
End Sub

Sub ChangeBorderColor(ByVal strRangename As String)
    
    Dim rng As Range
    Set rng = Application.Range(strRangename)
    With rng
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
    
        If strRangename <> "rng_besk_Figur1" Then
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ThemeColor = 1
                .TintAndShade = 0
                .Weight = xlThin
            End With
        End If
    End With
    If strRangename = "rng_besk_RörligaPriser" Then
        Set rng = wsStart.Range("G38:M38")
        
        With rng.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
        Set rng = wsStart.Range("M36:M38")
            With rng.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .Weight = xlThin
            End With
    End If
    Set rng = Nothing
End Sub
Sub ChangeBorderNoneColor()
    
    Dim rng As Range
    Set rng = wsStart.Range("G36:M38")
    With rng
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlContinuous
    End With
    Set rng = Nothing
    Application.Range("rng_besk_RörligaPriser").Font.ColorIndex = 2
    Application.Range("rng_besk_FastaPriser").Font.ColorIndex = 2
    Application.Range("rng_besk_Figur1").Font.ColorIndex = 2
End Sub
