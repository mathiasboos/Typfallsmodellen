Attribute VB_Name = "frmComment"
Attribute VB_Base = "0{AE3D5477-DE93-4507-97E6-59CA7C352D78}{C633461A-7155-4CEE-ADE8-679E37065B87}"
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

Private Sub lbl_FAQ_Click()
    'Call GetFAQ
End Sub

Private Sub lbl_Read_More_Click()
    Call QuickRef
End Sub

Private Sub UserForm_Activate()
    TranslateFormsControls Me
    Me.Caption = fngetMsgBoxtext(16)
End Sub


