Attribute VB_Name = "frmTypfalllist"
Attribute VB_Base = "0{ECA89956-C171-4EF9-9B3E-4475E0ACBBDC}{C3F3295C-516A-426B-B0A9-2193CC8BC779}"
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

Private Sub cmdDelete_Click()
    If Len(lstTypfalllist) > 0 And vbYes = MsgBox(fngetMsgBoxtext(44) & ": " & lstTypfalllist & "?", vbYesNo) Then
        blnGetTyfall = True
        Application_Rest
        blnRemoveTypfall = True
        Call RemoveTypfall(lstTypfalllist)
        Application_Wakeup
        blnRemoveTypfall = False
        blnGetTyfall = False
    End If
End Sub

Private Sub cmdSaveAs_Click()
    Dim strTypfallName As String
    If Len(Trim(txtTypfallName)) = 0 And IsNull(lstTypfalllist) Then
        MsgBox fngetMsgBoxtext(47)
        Exit Sub
    End If
    If Len(Trim(txtTypfallName)) > 0 Then
        strTypfallName = txtTypfallName
    Else
        strTypfallName = lstTypfalllist
    End If
    blnGetTyfall = True
    Application_Rest
    BeginSaveAsMydefault (strTypfallName)
    blnGetTyfall = False
    Application_Wakeup
End Sub

Private Sub cmdShow1_Click()
    
    If Len(lstTypfalllist) > 0 Then
        blnGetTyfall = True
        Application_Rest
        Call GetMyDefault(lstTypfalllist)
        blnGetTyfall = False
        Application_Wakeup
    End If
End Sub

Private Sub lstTypfalllist_Click()
    If Not blnRemoveTypfall Then _
        cmdSaveAs.Caption = fngetMsgBoxtext(45)
End Sub


Private Sub txtTypfallName_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    cmdSaveAs.Caption = fngetMsgBoxtext(46)
End Sub

Private Sub txtTypfallName_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    cmdSaveAs.Caption = fngetMsgBoxtext(45)
End Sub

Private Sub txtTypfallName_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    cmdSaveAs.Caption = fngetMsgBoxtext(46)
End Sub

Private Sub txtTypfallName_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    cmdSaveAs.Caption = fngetMsgBoxtext(46)
End Sub

Private Sub UserForm_Activate()
    TranslateFormsControls Me
    Me.Caption = fngetMsgBoxtext(42)
End Sub


