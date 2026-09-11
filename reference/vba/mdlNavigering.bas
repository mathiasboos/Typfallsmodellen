Attribute VB_Name = "mdlNavigering"
Option Explicit

Sub Navigate_Headers()
'    Dim x
'    Dim rng As Range
'    Dim strHeader As String
'
'    strHeader = Application.Range("rng_Data_Headers")
'    Set rng = Application.Range("rng_Data_Top_Left")
'    Set rng = rng.Resize(1, rng.CurrentRegion.Columns.Count)
'
'    On Error Resume Next
'
'    x = Application.Match(strHeader, rng, 0)
'    If Not IsNumeric(x) Then
'        Set rng = Nothing
'        On Error GoTo 0
'        Exit Sub
'    End If
'    .Cells(Application.Range("rng_Data_Top_Left").Row, x).Select
'    On Error GoTo 0
'    Set rng = Nothing
End Sub

Sub Navigate_Headers_sort()
'    Dim rng As Range
'    Set rng = Application.Range("rng_Rubriker_sorterad").Offset(1)
'    Set rng = rng.Resize(1200, 1)
'    rng.Clear
'    With wsNavigering
'        Set rng = .Cells(2, 1).Resize(Application.Range("rng_navigering_len") - 1, 1)
'        rng.Copy
'        Application.Range("rng_Rubriker_sorterad").Offset(1).PasteSpecial Paste:=xlPasteValues
'
'        Set rng = Application.Range("rng_Rubriker_sorterad").CurrentRegion
'        .Sort.SetRange rng
'        .Sort.Header = xlYes
'        .Sort.MatchCase = False
'        .Sort.Orientation = xlTopToBottom
'        .Sort.SortMethod = xlPinYin
'        .Sort.Apply
'    End With
'    Set rng = Nothing
End Sub
Sub MoveToSheet(ByVal strName As String)
    If Worksheets(strName).Visible = True Then
        Worksheets(strName).Activate
    End If
End Sub
Sub SelectRangeInSheet(ByVal strNameSheet As String, ByVal strRange As String)
    If Worksheets(strNameSheet).Visible = True Then
        Worksheets(strNameSheet).Activate
        ActiveSheet.Range(strRange).Select
    End If
End Sub
Sub CustomView(ByVal strCustomView As String)
'    If Len(Trim(strCustomView)) = 0 Then Exit Sub
'    wsData.Activate
'    Application_Rest
'    Call Unprotect_Allsheets
'    ThisWorkbook.CustomViews(strCustomView).Show
'    Application_Wakeup
'    Call Protect_sheet(wsStart.CodeName)
End Sub

