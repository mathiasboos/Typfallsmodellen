Attribute VB_Name = "mdGetFormValue"
Option Explicit

Sub getFormValue()
    
    Dim Pens_Month As Double
    Dim Pens_Month_PP As Double
    
    On Error Resume Next
    Pens_Month = Application.Range("rng_Pens_Month")
    Pens_Month_PP = Application.Range("rng_Pens_MonthPP")
    
    If Pens_Month <= 0 Then Pens_Month = 1
    If Pens_Month_PP <= 0 Then Pens_Month_PP = 1
    
    On Error GoTo 0
End Sub



