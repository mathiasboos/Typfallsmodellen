Attribute VB_Name = "mdlHelp"
Option Explicit

Sub QuickRef()
    
    Dim strLink As String
    strLink = fngetMsgBoxtext(12)
    On Error GoTo err_handler
    
    ActiveWorkbook.FollowHyperlink Address:=strLink, NewWindow:=True
    Exit Sub
    
err_handler:
        If err.Number = -2147467260 Then
        Else
            MsgBox fngetMsgBoxtext(27)
            strLink = fngetMsgBoxtext(26)
            ActiveWorkbook.FollowHyperlink Address:=strLink, NewWindow:=True
        End If
End Sub

Sub GetFAQ()
    Dim strLink As String
    strLink = fngetMsgBoxtext(30)
    On Error GoTo err_handler
    
    ActiveWorkbook.FollowHyperlink Address:=strLink, NewWindow:=True
    Exit Sub
    
err_handler:
        If err.Number = -2147467260 Then
        Else
            MsgBox fngetMsgBoxtext(27)
            strLink = fngetMsgBoxtext(26)
            ActiveWorkbook.FollowHyperlink Address:=strLink, NewWindow:=True
        End If
End Sub
