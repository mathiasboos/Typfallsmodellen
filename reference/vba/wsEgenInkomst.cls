Attribute VB_Name = "wsEgenInkomst"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
' Färdiga typfall
Sub cmd_Save_Tyfall()
    'Ändrar X och Y axlarna på diagrammet i fliken Typfall efter önskade minimun, maximum och steglängd
    Dim min As Long
    Dim max As Long
    Dim aby As Long
    
    Dim Ymin As Long
    Dim Ymax As Long
    Dim Yaby As Long
    
    min = fnGetRangeFromNames(ThisWorkbook, "rng_Chart_egen_min_Xaxel")
    max = fnGetRangeFromNames(ThisWorkbook, "rng_Chart_egen_max_Xaxel")
    aby = fnGetRangeFromNames(ThisWorkbook, "rng_Chart_egen_step_Xaxel")
    If aby = 0 Then
        aby = Int((max - min) / 10 + 0.5)
        Application.Range("rng_Chart_egen_step_Xaxel") = aby
    End If
    Ymin = fnGetRangeFromNames(ThisWorkbook, "rng_Chart_egen_min_Yaxel")
    Ymax = fnGetRangeFromNames(ThisWorkbook, "rng_Chart_egen_max_Yaxel")
    Yaby = fnGetRangeFromNames(ThisWorkbook, "rng_Chart_egen_step_Yaxel")
    If Yaby = 0 Then
        Yaby = Int((Ymax - Ymin) / 10 + 0.5)
        Application.Range("rng_Chart_egen_step_Yaxel") = Yaby
    End If
    
    ActiveSheet.ChartObjects("Diagram 6").Activate
    ActiveChart.Axes(xlCategory).Select
    With ActiveChart.Axes(xlCategory)
        .MinimumScale = min
        .MaximumScale = max
        .MinorUnitIsAuto = True
        .MajorUnit = aby
        .Crosses = xlAutomatic
        .ReversePlotOrder = False
        .ScaleType = xlLinear
        .DisplayUnit = xlNone
    End With
    ActiveChart.Axes(xlValue).Select
    With ActiveChart.Axes(xlValue)
        .MinimumScale = Ymin
        .MaximumScale = Ymax
        .MinorUnitIsAuto = True
        .MajorUnit = Yaby
        .Crosses = xlAutomatic
        .ReversePlotOrder = False
        .ScaleType = xlLinear
        .DisplayUnit = xlNone
    End With
    fnGetRangeFromNames(ThisWorkbook, "rng_Chart_egen_min_Xaxel").Select
End Sub

Private Sub Worksheet_Activate()
    If fnGetRangeFromNames(ThisWorkbook, "rng_EgenVetor_Typis") <> "Egen lönelista" Then
        fnGetRangeFromNames(ThisWorkbook, "rng_EgenVetor_Typis") = "Egen lönelista"
    End If
End Sub

Private Sub Worksheet_Deactivate()
    If fnGetRangeFromNames(ThisWorkbook, "rng_EgenVetor_Typis") <> "Egen lönelista" Then
        fnGetRangeFromNames(ThisWorkbook, "rng_EgenVetor_Typis") = "Egen lönelista"
    End If
End Sub

