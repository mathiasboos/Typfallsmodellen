Attribute VB_Name = "wsInkomstvektor"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
' Egen lönelista mm som används för egen lönevektor

Private Sub optBalance_ratio_Click()
    Range("rng_Top_Balance_ratio_array").Select
End Sub

Private Sub optFound_Click()
     Range("rng_Top_Yield_array").Select
End Sub

Private Sub optEarn_Click()
    Range("rng_Top_earn_array").Select
End Sub

Private Sub optPrivate_save_Click()
    Range("rng_Top_Private_save_array").Select
End Sub

Private Sub OptReal_Growth_Click()
    Range("rng_Top_Real_Growth_array").Select
End Sub
