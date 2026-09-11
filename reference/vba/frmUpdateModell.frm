Attribute VB_Name = "frmUpdateModell"
Attribute VB_Base = "0{FDB2B0CE-0645-4EB7-9405-0E11E9FD0AF2}{93B9CB9F-6363-4DF6-B820-A12B20315C02}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub cmdClose_Click()
    If chkUpdateModell Then
        Application.Range("rng_ShowUpdateModell") = False
    End If
    Unload Me
End Sub

Private Sub UserForm_Activate()
    TranslateFormsControls Me
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
    End If
End Sub
