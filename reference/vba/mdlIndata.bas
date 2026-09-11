Attribute VB_Name = "mdlIndata"
Option Explicit
Sub CleanIndata()
    If vbYes = MsgBox(fngetMsgBoxtext(50), vbYesNo) Then
        Dim rngX As Range
        Dim rngY As Range
        Set rngX = fnGetRangeFromNames(ThisWorkbook, "rngXTopleft")
        Set rngY = fnGetRangeFromNames(ThisWorkbook, "rngYTopleft")
        
        Set rngX = rngX.Offset(1).Resize(rngX.CurrentRegion.Rows.Count, rngX.CurrentRegion.Columns.Count)
        Set rngY = rngY.Offset(1).Resize(rngY.CurrentRegion.Rows.Count, rngY.CurrentRegion.Columns.Count)
        
        rngX.ClearContents
        rngY.ClearContents
        ' Rensa även formaten
'        rngX.Clear
'        rngY.Clear
        Set rngX = Nothing
        Set rngY = Nothing
    End If
End Sub

