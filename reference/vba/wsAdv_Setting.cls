Attribute VB_Name = "wsAdv_Setting"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "ComboBox4, 34, 0, MSForms, ComboBox"
Attribute VB_Control = "ComboBox3, 33, 1, MSForms, ComboBox"
Option Explicit
'------------------ Adv_settings Worksheet ------------------
Private Sub Worksheet_Change(ByVal Target As Range)
    
    If Application.Range("rng_Run_From_Indata") Then Exit Sub
    If blnGetTyfall Then Exit Sub
    
    'Step 1: Does the changed range intersect specified range?
    If Intersect(Target, Range("Rng_resp_rel")) Is Nothing Then

    'Step 2: If there is no intersection, exit procedure
    Exit Sub

    'Step 3: If there is an intersection ...
    Else
       wsStart.Unprotect
       If wsNameRange.Range("rng_OpeningProgress") Then wsNameRange.Range("rng_OpeningProgress") = False
       On Error Resume Next
       Application.ScreenUpdating = False
      
       'Stop
       
       If Application.Range("Rng_resp_rel") = 1 Then
                Sheets("Start").Activate
                ActiveSheet.Range("I42:I44").Select
                Selection.Style = "Percent"
                Selection.NumberFormat = "0.0%"
            Else
                Sheets("Start").Activate
                ActiveSheet.Range("I42:I44").Select
                Selection.NumberFormat = "#,##0"
                'Rows("46:46").RowHeight = 12
        End If
                wsStart.Protect
                Sheets("adv_settings").Activate
                Range("Rng_resp_rel").Select
        'Close out the If statement
    End If

End Sub
