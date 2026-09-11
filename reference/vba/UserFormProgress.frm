Attribute VB_Name = "UserFormProgress"
Attribute VB_Base = "0{FAC5AD00-2294-4C6C-9B2D-B2A8E122EBEB}{D621342B-0F34-4666-A5B5-045FC98D0CB0}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit


Private Sub UserForm_Activate()
    Me.Caption = fngetMsgBoxtext(37)
    Call StartUp 'Kolla villkor, se modulen i mdlCalculate, och därifrån kallas
    'Call Mcalc '-som kör Petra, i modul VBA_go
End Sub

