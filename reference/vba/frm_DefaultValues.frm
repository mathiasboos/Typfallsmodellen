Attribute VB_Name = "frm_DefaultValues"
Attribute VB_Base = "0{334AFD50-95C2-40E8-983A-20A56BC32259}{31C8E34C-9A65-40EF-B9B5-06619A588E73}"
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

Private Sub cmdGo_Click()

On Error GoTo errTag
    
    Dim rng As Range
    Set rng = fnGetRangeFromNames(ThisWorkbook, "rng_TopDefault_value")
    Worksheets(rng.Parent.Name).Activate
    rng.Select
    Set rng = Nothing
CleanTag:
    Set rng = Nothing
    Exit Sub
errTag:
    If err.Number = -2147024809 Then
        MsgBox fngetMsgBoxtext(31)
    Else
        MsgBox err.Description & " Error number: " & err.Number, vbCritical
    End If
    GoTo CleanTag
End Sub

Private Sub UserForm_Activate()
     Dim rng As Range
     Dim c As Range
     Dim i As Integer
     Dim FormatValue
     Me.Caption = fngetMsgBoxtext(38)
     Set rng = fnGetRangeFromNames(ThisWorkbook, "rng_TopDefault_value")
     Set rng = rng.Resize(rng.CurrentRegion.Rows.Count, 1)
     With Me.lstDefaultValues
        .Clear
        For Each c In rng
            .AddItem
            .List(i, 0) = c.Value
            If c.Offset(, 1) < 1 Then
                FormatValue = Format(c.Offset(, 1), "0.00%")
            ElseIf Right(c, 6) = "(åååå)" Or Right(c, 6) = "(yyyy)" Then
                FormatValue = c.Offset(, 1)
            ElseIf c.Offset(, 1) > 1000 Then
                FormatValue = Format(c.Offset(, 1), "# ###")
            Else
                FormatValue = c.Offset(, 1)
            End If
            
            .List(i, 1) = FormatValue
            i = i + 1
        Next c
     
     End With
     Set rng = Nothing
End Sub


