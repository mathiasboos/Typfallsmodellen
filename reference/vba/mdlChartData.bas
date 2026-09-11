Attribute VB_Name = "mdlChartData"
Option Explicit
Sub ShowChartData(ByVal index As Long)
    Range("rng_Chart_data_" & index).Select
End Sub
Sub ShowChartData_1()
    ShowChartData (1)
End Sub
Sub Show_Hide_Chart(ByVal Ws As Worksheet, ByVal ChartName As String)
    If Not SheetExists(Ws.Name) Then Exit Sub
    If Ws.Name = wsIndata.Name Then Ws.Range("rngXTopleft").Select
    If Ws.ChartObjects(ChartName).Visible = True Then
        Ws.ChartObjects(ChartName).Visible = False
    Else
        Ws.ChartObjects(ChartName).Visible = True
    End If
End Sub

Sub Show_Hide_AllChart(ByVal Ws As Worksheet)
    If Not SheetExists(Ws.Name) Then Exit Sub
    If Ws.Name = wsIndata.Name Then Ws.Range("rngXTopleft").Select
    Dim chr As ChartObject
    For Each chr In Ws.ChartObjects
        If Ws.ChartObjects(chr.Name).Visible = True Then
            Ws.ChartObjects(chr.Name).Visible = False
        Else
            Ws.ChartObjects(chr.Name).Visible = True
        End If
    Next chr
End Sub

Sub ShowChart_in_ArvIP()
    Call Show_Hide_Chart(wsArv_ip, "chart_arvIP")
End Sub

Sub ShowChart_in_ArvPP()
    Call Show_Hide_Chart(wsArv_pp, "chart_arvPP")
End Sub
Sub ShowChart_in_Indata()
    Call Show_Hide_Chart(wsIndata, "chart_Kompens")
End Sub
Sub ShowAllChart_in_Indata()
    Call Show_Hide_AllChart(wsIndata)
End Sub
Sub ShowChart_in_Egeninkomst()
    Call Show_Hide_Chart(wsEgenInkomst, "Diagram 6")
End Sub
Sub Show_Hide_Chart_Priser(ByVal blnShow As Boolean)
    Call Unprotect_sheet(wsStart.CodeName)
    Application.Range("rng_Chart_Eran_Show") = blnShow
    If blnShow Then
        With wsStart
            .ChartObjects("chart_Priser").Visible = True
            .Shapes("OptChart_Month").Visible = True
            .Shapes("OptChart_Year").Visible = True
        End With
        Application.Range("rng_IncomeChartHeader").Font.Color = RGB(0, 0, 0)
        Application.Range("rng_IncomeChartHeader_Tidsperiod").Font.ColorIndex = 0
    Else
        With wsStart
            .ChartObjects("chart_Priser").Visible = False
            .Shapes("OptChart_Month").Visible = False
            .Shapes("OptChart_Year").Visible = False
        End With
        'Application.Range("rng_IncomeChartHeader").Font.Color = RGB(255, 192, 0)
        'Application.Range("rng_IncomeChartHeader_Tidsperiod").Font.ColorIndex = 2
    End If
    Call Protect_sheet(wsStart.CodeName)
End Sub
