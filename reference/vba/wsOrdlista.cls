Attribute VB_Name = "wsOrdlista"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
' Ordlista - Funkar inte ThisWorkbook ?
'Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'    Dim rngSearch As Range
'    Dim C As Range
'    Dim blnWordExists As Boolean
'    blnWordExists = False
'    'Här stannar koden. unkar med äldre version från 2000? Tanken är att gå till aktuell Bokstav
'    If Target.Row <> fnGetRangeFromNames(ThisWorkbook, "rng_Ordlista_Top").Row Then Exit Sub
'
'    Set rngSearch = fnGetRangeFromNames(ThisWorkbook, "rng_Ordlista_Top")
'    Set rngSearch = rngSearch.Offset(1).Resize(Columns("A:A").Range("A1000000").End(xlUp).Row, 1)
'    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationManual
'    With wsOrdlista
'        For Each C In rngSearch
'            If Len(Trim(C)) = 0 Then
'                .Rows("" & C.Row & ":" & C.Row & "").RowHeight = 4.5
'            End If
'        Next C
'    End With
'    For Each C In rngSearch
'        If LCase(Left(Trim(C), 1)) = LCase(Trim(Target)) And C.Font.Bold = True Then
'            C.Select
'            ActiveWindow.ScrollRow = C.Row
'            blnWordExists = True
'            Exit For
'        End If
'    Next C
'    Set rngSearch = Nothing
'    If blnWordExists Then
'        Application.Range("rng_Ord_saknas") = ""
'    Else
'         Application.Range("rng_Ord_saknas") = "Inget ord har hittats(sökt bokstav: " & Target.Value & ")"
'    End If
'    Application.ScreenUpdating = True
'    Application.Calculation = xlCalculationAutomatic
'End Sub
