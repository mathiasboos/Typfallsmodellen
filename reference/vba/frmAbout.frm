Attribute VB_Name = "frmAbout"
Attribute VB_Base = "0{C77EBE5D-C66E-4279-B2A7-8FCBE9AF616E}{FB022EB7-D2A2-4EAD-8C4B-69595E82D70F}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub cmdClose_Click()
    If blnShowCloseButtonAbout Then
        Unload Me
    Else
        Dim lngRow As Long
        Dim lngCol As Long
        Dim rng As Range
        Dim r As Range
        Dim blnUserExists As Boolean
        blnUserExists = False
        Application_Rest
        
        Set rng = Application.Range("rng_TopLeftAbout")
        lngRow = rng.Row
        Set rng = rng.Resize(rng.CurrentRegion.Rows.Count, 1)
        If Not rng Is Nothing Then
            For Each r In rng
                If LCase(r.Value) = LCase(GetTheNameNETWORK) Then
                    blnUserExists = True
                    lngRow = r.Row
                    Exit For
                End If
            Next r
        End If
        If blnUserExists Then
            lngRow = r.Row + 1
        Else
            lngRow = lngRow + 1
        End If
        lngCol = rng.Column
        With wsNameRange
            .Cells(lngRow, lngCol) = GetTheNameNETWORK
            .Cells(lngRow, lngCol + 1) = ThisWorkbook.Path
            .Cells(lngRow, lngCol + 2) = Date
        End With
        Unload Me
        ThisWorkbook.Save
        Application_Wakeup
    End If
End Sub

Private Sub lblEmail_Click()
    Dim Link As String
    Link = "mailto:" & fnGetRangeFromNames(ThisWorkbook, "rng_Email_To")
       On Error GoTo Err_Handel
       ActiveWorkbook.FollowHyperlink Address:=Link, NewWindow:=True
       Unload Me
       Exit Sub
Err_Handel:
       MsgBox fngetMsgBoxtext(41)
End Sub

Private Sub UserForm_Activate()
    TranslateFormsControls Me
    Me.Caption = fngetMsgBoxtext(15)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        MsgBox fngetMsgBoxtext(9), vbCritical
    End If
     
End Sub


