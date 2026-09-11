Attribute VB_Name = "mdlTranslate"
 Option Explicit

'Public blnValidMonth As Boolean
''Sub ShowInput()
''    frmReadInput.Show
''End Sub
''Sub ShowComment()
''    frmComment.Show
''End Sub
Sub ShowSettings()
    frmSettings.Show
End Sub

Sub ShapeTranslate()
    Dim rng As Range
    Dim x
    Dim a
    Dim strActLang As String
    Dim r As Range
    Dim wk As Worksheet
    Dim Sh As Shape
    Dim cmd As Shape
    
    Application.ScreenUpdating = False
    strActLang = Application.Range("rngValSysLang")

    Set rng = wsSysdata.Range("rngTopShapeList")
    Set r = rng.Resize(1, rng.CurrentRegion.Columns.Count)
    On Error Resume Next
    x = Application.Match(strActLang, r, 0)
    If Not IsNumeric(x) Then
        Set rng = Nothing
        Set r = Nothing
        On Error GoTo 0
        Exit Sub
    End If
    Set rng = rng.Offset(1).Resize(rng.CurrentRegion.Rows.Count - 1, 1)
    For Each wk In ThisWorkbook.Worksheets
        If wk.Visible = xlSheetVisible And wk.Shapes.Count > 0 Then
            If wk.Name = "Start" Then wk.Unprotect
            For Each Sh In wk.Shapes
                For Each r In rng
                    If Sh.Name = r.Value Then
                        If LCase(r.Offset(, 1)) = "shape" Then
                            ' för formulär knappar
                            wk.Activate
                            wk.Shapes.Range(Array(r.Value)).Select
                            a = Selection
                            If TypeName(a) <> "Range" Then
                                Selection.Characters.text = r.Offset(, x - 1)
                            End If
                        Else
                            ' för Active X knappar
                            Set cmd = wk.Shapes(Sh.Name)
                            With cmd.OLEFormat.Object
                                .Object.Caption = r.Offset(, x - 1)
                            End With
                        End If
                    End If
                Next r
            Next Sh
            If wk.Name = "Start" Then wk.Protect
        End If
        wk.Range("A1").Select
    Next wk
    Set cmd = Nothing
    On Error GoTo 0

    Set rng = Nothing
    Set r = Nothing

    Application.ScreenUpdating = True
    wsStart.Activate
    wsStart.Range("rngModelHeader").Select
   
End Sub

Sub TranslateFormsControls(frm As UserForm)
    Dim ct As MSForms.Control
    Dim x
    Dim strName As String

    DefineTranslateRange
    For Each ct In frm.Controls
        'If TypeOf ct Is MSForms.CommandButton Then
            strName = frm.Controls(ct.Name).Name
            On Error Resume Next
            x = Application.Match(strName, rngTranslateRange, 0)
            If Not IsNumeric(x) Then
                On Error GoTo 0
            Else
                x = x + wsSysdata.Range("rngTopShapeList").Row - 1
                frm.Controls(ct.Name).Caption = wsSysdata.Cells(x, wsSysdata.Range("rng_FirstLang").Column).Offset(, wsNameRange.Range("sysLang"))
            End If
        'End If
    Next ct
    'TranslateFormsMultiPage 2024-02-06
    On Error GoTo 0
End Sub


Sub DefineTranslateRange()
    Set rngTranslateRange = wsSysdata.Range("rngTopShapeList").Resize(wsSysdata.Range("rngTopShapeList").CurrentRegion.Rows.Count, 1)
End Sub
Sub TranslateMonth(ByVal LangIndex As Integer)

    Dim rng As Range
    Dim rng2 As Range
    Dim x
    Dim strRangeMonth As String
    Dim strRangeMonthIndex As String
    Dim strRangeMonthTop As String
    Dim strRangeMonthIndexTop As String
    Dim intCol As Integer
    Dim strBorn_Month_Value As String
    Dim strPens_Month_Value As String
    Dim strWork_Month_Value As String
    

    If pblnDisableListClick Then Exit Sub
    strRangeMonth = "rng_Month_"
    strRangeMonthIndex = "rng_Month_Index_"
    
    If LangIndex = 0 Then
        LangIndex = 1
        intCol = 1
    Else
        LangIndex = 0
        intCol = 2
    End If
    strRangeMonth = strRangeMonth & LangIndex
    strRangeMonthIndex = strRangeMonthIndex & LangIndex
    
    strRangeMonthTop = strRangeMonth & "_Top"
    strRangeMonthIndexTop = strRangeMonthIndex & "_Top"
    
    
    Set rng = wsNameRange.Range(strRangeMonth)
    Set rng2 = wsNameRange.Range(strRangeMonthIndex)
    On Error Resume Next
    x = Application.Match(wsStart.Range("PARMonth"), rng2, 0) + wsNameRange.Range(strRangeMonthIndexTop).Row
    strPens_Month_Value = wsNameRange.Cells(x, intCol)

    Call Unprotect_sheet(wsStart.CodeName)

    wsStart.Range("PARMonth") = strPens_Month_Value 'ReStore month

    Call Protect_sheet(wsStart.CodeName)
    On Error GoTo 0

    ShapeTranslate

    If ActiveSheet.Name = wsStart.Name Then
        wsStart.Activate
        wsStart.Range("rngModelHeader").Select
    End If
End Sub
Sub Translate_Forms()
    Dim i As Integer
    For i = 0 To VBA.UserForms.Count - 1
         If UserForms(i).Visible Then
            UserForms(i).Hide
            UserForms(i).Show
        End If
    Next i
End Sub
Function IsUserFormLoaded(ByVal UFName As String) As Boolean
    Dim UForm As Object
    IsUserFormLoaded = False
    For Each UForm In VBA.UserForms
        If UForm.Name = UFName Then
            IsUserFormLoaded = True
            Exit For
        End If
    Next
End Function

Public Function fngetMsgBoxtext(ByVal index As Long) As String

    Dim x
    Dim rng As Range
    'Genererar msgbox meddelande. Texten läses in från rader i bladet SysMsgBoxLang
    Set rng = wsSysMsg.Range("rng_Lang_msgTextIndex")
    Set rng = rng.Resize(rng.CurrentRegion.Rows.Count, 1)

    On Error Resume Next
    x = Application.Match(index, rng, 0)
    If Not IsNumeric(x) Or x < 2 Then
        On Error GoTo 0
        fngetMsgBoxtext = wsSysMsg.Cells(2, 2).Offset(, wsNameRange.Range("sysLang"))
        Exit Function
    End If
    fngetMsgBoxtext = wsSysMsg.Cells(x, 2).Offset(, wsNameRange.Range("sysLang"))

    On Error GoTo 0

End Function

Public Function fngetMsgBoxtextByColumn(ByVal index As Long, ByVal ColOffset As Integer) As String

    Dim x
    Dim rng As Range
    'Genererar msgbox meddelande. Texten läses in från rader i bladet SysMsgBoxLang
    Set rng = wsSysMsg.Range("rng_Lang_msgTextIndex")
    Set rng = rng.Resize(rng.CurrentRegion.Rows.Count, 1)

    On Error Resume Next
    x = Application.Match(index, rng, 0)
    If Not IsNumeric(x) Or x < 2 Then
        On Error GoTo 0
        fngetMsgBoxtextByColumn = "No value"
        Exit Function
    End If
    fngetMsgBoxtextByColumn = wsSysMsg.Cells(x, ColOffset + 2)

    On Error GoTo 0

End Function


