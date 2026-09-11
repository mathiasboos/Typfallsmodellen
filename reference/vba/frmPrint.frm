Attribute VB_Name = "frmPrint"
Attribute VB_Base = "0{E50EC401-8F61-4ED6-9B12-DAECEECCFB79}{E05B2401-3C69-4D90-B44A-B7C6979C1070}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdUpdate_Click()
    TranslateFormsControls Me
End Sub

Private Sub cmdGo_Click()
    Dim i As Integer
    Dim blnwkHide As Boolean
    Dim rng As Range
    Dim iSelectedItem As Integer
    Dim iListRow As Integer
    
    On Error GoTo errTag
    
    With lstPrintArea
        For i = 0 To .ListIndex
            If .Selected(i) Then
                iSelectedItem = iSelectedItem + 1
                iListRow = i
            End If
        Next i
        If iSelectedItem = 1 Then
                
            If LCase(Left(.List(iListRow, 1), 5)) = "chart" Then
                blnwkHide = Worksheets(.List(iListRow, 0)).Visible
                If Not blnwkHide Then Worksheets(.List(iListRow, 0)).Visible = True
                    Worksheets(.List(iListRow, 0)).Unprotect
                    Worksheets(.List(iListRow, 0)).Activate
                If LCase(Right(.List(iListRow, 1), 5)) = "month" Then
                    Worksheets(.List(iListRow, 0)).Range("rngChartSection_month").Select
                ElseIf LCase(Right(.List(iListRow, 1), 4)) = "year" Then
                    Worksheets(.List(iListRow, 0)).Range("rngChartSection").Select
                Else
                    Worksheets(.List(iListRow, 0)).Range("rng_IncomeChartHeader").Select
                End If
                Worksheets(.List(iListRow, 0)).ChartObjects(.List(iListRow, 1)).Activate
                
                If Worksheets(.List(iListRow, 0)).Visible <> blnwkHide Then Worksheets(.List(iListRow, 0)).Visible = blnwkHide
            Else
                 blnwkHide = Worksheets(.List(iListRow, 0)).Visible
                 If Not blnwkHide Then Worksheets(.List(iListRow, 0)).Visible = True
                 Set rng = Worksheets(.List(iListRow, 0)).Range(.List(iListRow, 1))
                 Worksheets(.List(iListRow, 0)).Unprotect
                 Worksheets(.List(iListRow, 0)).Activate
                 rng.Select
                 If Worksheets(.List(iListRow, 0)).Visible <> blnwkHide Then Worksheets(.List(iListRow, 0)).Visible = blnwkHide
            End If
            Worksheets(.List(iListRow, 0)).Protect
        End If
    End With
CleanTag:
    'Unload Me
    Exit Sub
     
errTag:
    If err.Number = -2147024809 Then
        MsgBox fngetMsgBoxtext(31)
    Else
        MsgBox err.Description & " Error number: " & err.Number, vbCritical
    End If
    GoTo CleanTag
End Sub

Private Sub cmdPDF_Click()
    Dim i As Integer
    Dim argRangeName As String
    Dim strPrintRange As String
    Dim blnwkHide As Boolean
    Dim rng As Range
    Dim strFilename As String
    
    On Error GoTo errTag
    
    Application_Rest
    With lstPrintArea
        For i = 0 To .ListIndex
            If .Selected(i) Then
                strFilename = ThisWorkbook.Path & "\" & .List(i, 2) & ".pdf"
                'Hantera om filen är öppen.
                If fnFileIsOpen(strFilename) Then
                    MsgBox fngetMsgBoxtext(21), vbCritical
                    Application.ScreenUpdating = True
                    Exit Sub
                End If
            
                If LCase(Left(.List(i, 1), 5)) = "chart" Then
                    blnwkHide = Worksheets(.List(i, 0)).Visible
                    If Not blnwkHide Then Worksheets(.List(i, 0)).Visible = True
                    Worksheets(.List(i, 0)).Unprotect
                    Worksheets(.List(i, 0)).ChartObjects(.List(i, 1)).Activate
                    ActiveChart.ExportAsFixedFormat Type:=xlTypePDF, FileName:=strFilename, _
                        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
                        :=False, OpenAfterPublish:=False
                    If Worksheets(.List(i, 0)).Visible <> blnwkHide Then Worksheets(.List(i, 0)).Visible = blnwkHide
                Else
                     argRangeName = .List(i, 0) & "!" & (.List(i, 2))
                    
                     strPrintRange = argRangeName
                     blnwkHide = Worksheets(.List(i, 0)).Visible
                     If Not blnwkHide Then Worksheets(.List(i, 0)).Visible = True
                     Set rng = Worksheets(.List(i, 0)).Range(.List(i, 1))
                     rng.ExportAsFixedFormat Type:=xlTypePDF, FileName:=strFilename, _
                        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
                        :=False, OpenAfterPublish:=False
                     If Worksheets(.List(i, 0)).Visible <> blnwkHide Then Worksheets(.List(i, 0)).Visible = blnwkHide
                End If
            End If
        Next i
    End With
CleanTag:
    wsStart.Activate
    Range("rngModelHeader").Select
    Application_Wakeup
    Unload Me
    Exit Sub
errTag:
    If err.Number = -2147024809 Then
        MsgBox fngetMsgBoxtext(31)
    Else
        MsgBox err.Description & " Error number: " & err.Number, vbCritical
    End If
    GoTo CleanTag
End Sub

Private Sub cmdPrint_Form_Click()
Dim i As Integer
    Dim argRangeName As String
    Dim strPrintRange As String
    Dim blnwkHide As Boolean

    On Error GoTo errTag
    Application_Rest
    With lstPrintArea
        For i = 0 To .ListIndex
            If .Selected(i) Then
                If LCase(Left(.List(i, 1), 5)) = "chart" Then
                    wsStart.Unprotect
                    wsStart.ChartObjects(.List(i, 1)).Activate
                    ActiveChart.PrintOut
                    wsStart.Protect
                Else
                     argRangeName = .List(i, 0) & "!" & (.List(i, 2))
                    
                     strPrintRange = argRangeName
                     blnwkHide = Worksheets(.List(i, 0)).Visible
                     If Not blnwkHide Then Worksheets(.List(i, 0)).Visible = True
                     With Worksheets(.List(i, 0)).PageSetup
                         .FitToPagesWide = 1
                         .FitToPagesTall = 1
                     End With
                     Worksheets(.List(i, 0)).Range(.List(i, 1)).PrintOut
                     If Worksheets(.List(i, 0)).Visible <> blnwkHide Then Worksheets(.List(i, 0)).Visible = blnwkHide
                     End If
            End If
        Next i

    End With
CleanTag:
    wsStart.Activate
    Range("rngModelHeader").Select
    Application_Wakeup
    Unload Me
    Exit Sub
    
errTag:
    If err.Number = -2147024809 Then
        MsgBox fngetMsgBoxtext(31)
    Else
        MsgBox err.Description & " Error number: " & err.Number, vbCritical
    End If
    GoTo CleanTag
End Sub
Private Sub UserForm_Activate()
    Dim N As Name
    Dim i As Long
    Dim j As Long
    Dim iPos As Integer
    Dim lngLangIndex As Long
    Dim wk As Worksheet
    Dim chrobj As ChartObject
    Dim x
    Dim r As Range
    Dim rng As Range
    
    
    On Error GoTo errTag
    lstPrintArea.Clear
    TranslateFormsControls Me
    Me.Caption = fngetMsgBoxtext(17)
    lngLangIndex = fnGetRangeFromNames(ThisWorkbook, "sysLang")
    Select Case lngLangIndex
        Case 0
            i = 0
            With lstPrintArea
                For Each N In ThisWorkbook.Names
                    If InStr(1, N.Name, "__prt_sv_") > 0 Then
                        If InStr(1, N.Name, "!") Then
                            For j = 1 To Len(N.Name)
                                If Mid(N.Name, j, 1) = "!" Then
                                    iPos = j
                                    Exit For
                                End If
                            Next j
                        End If
                        
                        If LCase(Left(N.Name, 9)) = "__prt_sv_" Then
                            .AddItem
                            .List(i, 0) = fnGetRangeFromNames(ThisWorkbook, N.Name).Parent.Name
                            .List(i, 1) = fnGetRangeFromNames(ThisWorkbook, N.Name).Address
                            .List(i, 2) = Mid(N.Name, 10)
                            i = i + 1
                        End If
                    End If
                Next N
            End With
        Case Else
            i = 0
            With lstPrintArea
                For Each N In ThisWorkbook.Names
                    If InStr(1, N.Name, "__prt_en_") > 0 Then
                        If InStr(1, N.Name, "!") Then
                            For j = 1 To Len(N.Name)
                                If Mid(N.Name, j, 1) = "!" Then
                                    iPos = j
                                    Exit For
                                End If
                            Next j
                        End If
                        
                        If LCase(Left(N.Name, 9)) = "__prt_en_" Then
                            .AddItem
                            .List(i, 0) = fnGetRangeFromNames(ThisWorkbook, N.Name).Parent.Name
                            .List(i, 1) = fnGetRangeFromNames(ThisWorkbook, N.Name).Address
                            .List(i, 2) = Mid(N.Name, 10)
                            i = i + 1
                        End If
                    End If
                Next N
            End With
    End Select

    With lstPrintArea
        Set rng = wsSysdata.Range("rngTopShapeList")
        Set r = rng.Resize(rng.CurrentRegion.Rows.Count, 1)
        For Each chrobj In wsStart.ChartObjects
            .AddItem
            .List(i, 0) = wsStart.Name
            .List(i, 1) = chrobj.Name
            On Error Resume Next
            x = Application.Match(chrobj.Name, r, 0) + 2
            If Not IsNumeric(x) Then
                Set rng = Nothing
                Set r = Nothing
                On Error GoTo errTag
                Exit Sub
            End If
            .List(i, 2) = wsSysdata.Cells(x, 3).Offset(, lngLangIndex)
            i = i + 1
            On Error GoTo errTag
        Next chrobj
        Set rng = Nothing
        Set r = Nothing
    End With
Exit Sub
    
errTag:
    MsgBox err.Description & " Error number: " & err.Number, vbCritical
End Sub


