Attribute VB_Name = "MdlCalculate"
Option Explicit
' Denna modul förbreder beräkningar

Sub Start_Modell()
    'Knappen beräkna, startar upp baren med % klart
    Dim lngAgeMax As Long
    Dim lngHeight_Nagratal As Long
    Dim lngFirstValue As Long
    'Antal inkomstår som ska beräknas, se flik "input" och J1 (ålder), framgår av namngranskare under formler för "rng_age_top"
    'Maximalt antal beräkningar (106), förenklar börjar ju vanligtvis vid 15 och slutar tidigare än 105
    lngAgeMax = fnGetRangeFromNames(ThisWorkbook, "rng_Age_Top").CurrentRegion.Rows.Count - 1
    lngTotTasks = lngAgeMax
    lngDoneTask = 0
    PctDone = 0
    lngHeight_Nagratal = fnGetRangeFromNames(ThisWorkbook, "rng_Height_Tabell_Nagratal")
    lngFirstValue = fnGetRangeFromNames(ThisWorkbook, "rng_Top_NagraTal")
    If fnGetRangeFromNames(ThisWorkbook, "IPS").HasFormula = False Then 'Se flik NameRange
        fnGetRangeFromNames(ThisWorkbook, "IPS").Formula = "=IPS_Monthly*12"
    End If
    If Application.Range("Visa_process").Value = 1 Then
        Call ShowUserForm(247, 199, 70) 'med rGb färger se modul mdlProgress
            'Userform_activate i formuläret UserFormProgress (höger klicka och visa kod)
    Else
        Call StartUp
    End If
End Sub

Sub StartUp_Indata(ByVal blnFirstRowRun As Boolean)
    'Debug.Print 10631
    Application_Rest
    'Debug.Print 10632
    pblnCloseOrSave = False
    If blnFirstRowRun Then
        Dim lngAgeMax As Long
        Dim lngHeight_Nagratal As Long
        Dim lngFirstValue As Long
        'Debug.Print 10633
        lngAgeMax = Application.Range("rng_Age_Top").CurrentRegion.Rows.Count - 1
        lngHeight_Nagratal = Application.Range("rng_Height_Tabell_Nagratal")
        lngFirstValue = Application.Range("rng_Top_NagraTal")
        pblnDisableListClick = True
        'Debug.Print 10634
    End If
    
    Call Mcalc
    'Debug.Print 10635
    Application_Wakeup
    
    fnGetRangeFromNamesTw("rngSenasteStartChangeCells").Value = fnGetRangeFromNamesTw("rngCurrentStartChangeCells").Value

    pblnDisableListClick = False
    
End Sub


Sub StartUp()
    pblnCloseOrSave = False
    Application.Range("rng_Run_From_Indata") = False
    Application.ScreenUpdating = False
    Dim T
    T = Time
    
    If Not Application.Range("rng_Run_From_Indata") Then
        'Kontroll att vissa nödvändiga parametrar är ifyllda
        If Application.Range("BornYear") = "" Or _
            Application.Range("PARYear") = "" Or _
            Application.Range("wStartYear") = "" Or _
            Application.Range("Wage_Monthly") = "" Then
            
                If Application.Range("BornYear") = "" Then
                    Application.Range("BornYear").Select
                    MsgBox fngetMsgBoxtext(22), vbCritical
                    Unload UserFormProgress
                    Application_Wakeup
                    Exit Sub
                End If
                If Application.Range("PARYear") = "" Then
                    Application.Range("PARYear").Select
                    MsgBox fngetMsgBoxtext(23), vbCritical
                    Unload UserFormProgress
                    Application_Wakeup
                    Exit Sub
                End If
                If Application.Range("wStartYear") = "" Then
                    Application.Range("wStartYear").Select
                    MsgBox fngetMsgBoxtext(24), vbCritical
                    Unload UserFormProgress
                    Application_Wakeup
                    Exit Sub
                End If
                If Application.Range("Wage_Monthly") = "" And Application.Range("rng_Egen_Lon") = False Then
                    Application.Range("Wage_Monthly").Select
                    MsgBox fngetMsgBoxtext(25), vbCritical
                    Unload UserFormProgress
                    Application_Wakeup
                    Exit Sub
                End If
                
        End If
''        'Koll av att antal arbetande år...
''        If Application.Range("PARYear") - Application.Range("wStartYear") < 5 Then
''            'Application.Range("PARYear").Select
''            MsgBox fngetMsgBoxtext(40), vbCritical
''            'Unload UserFormProgress
''            'Application_Wakeup
''            'Exit Sub
''        End If
        
'        If Not Application.Range("rng_Run_From_Indata") Then 'Indata körs inte denna subrutin
'            Call Unprotect_sheet(wsStart.CodeName)
'        End If

        On Error Resume Next
'''        With wsStart
'''            .optNormal.Width = 72.75
'''            .optNormal.Height = 14.25
'''            .optAdvanced.Width = 72.75
'''            .optAdvanced.Height = 14.25
'''            .chkShowSkattRows.Width = 120
'''            .OptChart_Month.Width = 76.5
'''            .OptChart_Month.Height = 22.5
'''            .OptChart_Year.Width = 60
'''            .OptChart_Year.Height = wsStart.OptChart_Month.Height
'''        End With
        ValidateCellValue 'Kontroll av De ekonomiska valen Inflation , Real avkastning , Real tillväxt . Ger ev meddelande om att: Valen är kanske inte fullt rimliga/konsekventa på längre sikt.
    End If
    On Error GoTo 0
    
    pblnDisableListClick = True 'Behövs
    
    Application_Rest 'Behövs
    
    Call Mcalc 'Här körs Petra.
    
''    If Not Application.Range("rng_Run_From_Indata") Then _
''        Unload UserFormProgress 'Vet att den inte körs så
    Unload UserFormProgress
    
    fnGetRangeFromNamesTw("rngSenasteStartChangeCells").Value = fnGetRangeFromNamesTw("rngCurrentStartChangeCells").Value
    
'''    On Error Resume Next
''''    Call Fix_Chart_Start
'''    FixControllSize
'''    On Error GoTo 0
    pblnDisableListClick = False
    'Unload frmReadInput
'    Call getFormValue
    'Application.Range("Value_LastRun").Value = Application.Range("Wage_Monthly").Value 'Se nameRange bl2, Why??
    
    'If Not Application.Range("rng_Run_From_Indata") Then
    'Vet att den inte körs men skydda startfliken
    Call Protect_sheet(wsStart.CodeName) 'skydda startfliken
    'End If
    
    'Debug.Print Format(Time - t, "hh:mm:ss")
'    wsIndata.Range("rngRunTime") = Time - t 'Indata kör ju inte denna sub
    Application_Wakeup
End Sub



Sub Start_Modell_Indata()
    Dim lngAgeMax As Long
    lngAgeMax = fnGetRangeFromNames(ThisWorkbook, "rng_Age_Top").CurrentRegion.Rows.Count - 1
    If Application.Range("IPS").HasFormula = False Then
        Application.Range("IPS").Formula = "=IPS_Monthly*12"
    End If
End Sub
Sub Abort_run()
    Application.Range("rng_Abort_run") = True
End Sub

Sub StopRunning(ByVal iRow As Integer, ByVal iStopRow As Integer)
    If iRow = iStopRow Then Stop
End Sub

Sub Fix_Chart_Start()
    Dim dblMinAxel As Double
    Dim dblMaxAxel As Double
    Dim dblChartHeight As Double
    
    Call Unprotect_sheet(wsStart.CodeName)
    With wsStart
        .ChartObjects(2).Activate
        ActiveChart.Legend.Select
        Selection.Top = 2.51
        
        dblMinAxel = ActiveChart.Axes(xlValue).MinimumScale
        dblMaxAxel = ActiveChart.Axes(xlValue).MaximumScale
        dblChartHeight = ActiveSheet.ChartObjects(2).Height
    
        
        .ChartObjects(3).Activate
        ActiveChart.Legend.Select
        Selection.Top = 2.51
        
        ActiveChart.Axes(xlValue).MinimumScale = dblMinAxel
        ActiveChart.Axes(xlValue).MaximumScale = dblMaxAxel
        'ActiveChart.Height = dblChartHeight
        ActiveSheet.ChartObjects(3).Height = dblChartHeight
    
        
        .ChartObjects(5).Activate
        ActiveChart.Legend.Select
        Selection.Top = 2.51
        
        dblMinAxel = ActiveChart.Axes(xlValue).MinimumScale
        dblMaxAxel = ActiveChart.Axes(xlValue).MaximumScale
        'dblChartHeight = ActiveChart.Height
        dblChartHeight = ActiveSheet.ChartObjects(5).Height
        
        .ChartObjects(4).Activate
        ActiveChart.Legend.Select
        Selection.Top = 2.51
        ActiveChart.Axes(xlValue).MinimumScale = dblMinAxel
        ActiveChart.Axes(xlValue).MaximumScale = dblMaxAxel
        ActiveSheet.ChartObjects(4).Height = dblChartHeight
End With
    
    Range("rngModelHeader").Select
    Call Protect_sheet(wsStart.CodeName)
End Sub


''Sub Calculate_IndataVector()
''    'Egen inkomst
''    Dim iRow As Integer
''    Dim lngRow As Integer
''    Dim rng As Range
''    Dim r As Range
''    Dim lngCol As Integer
''    Dim lngOffsetCol As Integer
''
''    lngOffsetCol = 0
''    lngCol = Application.Range("rng_Data_PBB_Header").Column
''    iRow = 1
''    Set rng = Application.Range("rng_Top_Ersattning_EgenVektor")
''
''    Set rng = rng.Resize(lngAgeMax, 1)
''
''    With wsData
''        For Each r In rng
''            ' Dessa formler instämmer med orginal version blad Egen inkomst
''
''            'Range(G15) =akassa(E15;G$13*30;7)
''            r.Value = akassa(r.Offset(, -1), int_Index_Akassa * 30, 7)
''            'Range(H15) =Sjuk(E15/12;H$13*30;1;7,5;0,97;0,8;B15)
''            r.Offset(, 1) = sjuk(r.Offset(, -1) / 12, int_Index_Sjukpenning * 30, 1, 7.5, 0.97, 0.8, .Cells(iRow + 14, lngCol))
''            'Range(I15) =Sjuk($E15/12;I$13*30;2;7,5;0,97;0,8;B15)
''            r.Offset(, 2) = sjuk(r.Offset(, -1) / 12, int_Index_Sjuklon * 30, 2, 7.5, 0.97, 0.8, .Cells(iRow + 14, lngCol))
''            'Range(J15) =Sjuk($E15/12;J$13*30+14;1;10;0,97;0,8;B15)
''            r.Offset(, 3) = sjuk(r.Offset(, -1) / 12, int_Index_Foraldrarpenning * 30 + 14, 1, 10, 0.97, 0.8, .Cells(iRow + 14, lngCol))
''
''            If r.Row <> rng.Cells(1, 1).Row Then
''                If r.Row = rng.Cells(2, 1).Row Then
''                'Range(I16) =SA(E15;D16;B16;99;1;0;0)
''                    r.Offset(, 4) = SA(r.Offset(-1, -1), r.Offset(, -3), r.Offset(, -5), 99, 1, 0, 0)
''                ElseIf .Cells(iRow + 14 - 1, lngCol) > 0 Then
''                    If r.Row = rng.Cells(3, 1).Row Then
''                        If .Cells(iRow + 14 - 2, lngCol) > 0 Then
''                            'Range(I17) =SA(E16*B17/B16;D17;B17;99;1;E15*B17/B15;0)
''                            r.Offset(, 4) = SA(r.Offset(-1, -1) * .Cells(iRow + 14, lngCol) / wsData.Cells(iRow + 14 - 1, lngCol), _
''                                            .Cells(iRow - 1, 1), .Cells(iRow + 14, lngCol), 99, 1, _
''                                            r.Offset(-2, -1) * .Cells(iRow + 14, lngCol) / .Cells(iRow + 14 - 2, lngCol), 0)
''                        End If
''                    ElseIf .Cells(iRow + 14 - 2, lngCol) > 0 And .Cells(iRow + 14 - 3, lngCol) > 0 Then
''                        'Range(I18) =SA(E17*B18/B17;D18;B18;99;1;E16*B18/B16;E15*B18/B15)
''                        r.Offset(, 4) = SA(r.Offset(-1, -1) * .Cells(iRow + 14, lngCol) / .Cells(iRow + 14 - 1, lngCol), _
''                                        .Cells(iRow, 1), .Cells(iRow + 14, lngCol), 99, 1, _
''                                        r.Offset(-2, -1) * .Cells(iRow + 14, lngCol) / .Cells(iRow + 14 - 2, lngCol), _
''                                        r.Offset(-3, -1) * .Cells(iRow + 14, lngCol) / .Cells(iRow + 14 - 3, lngCol))
''                    End If
''                End If
''            End If
''        iRow = iRow + 1
''        Next r
''    End With
''    Set rng = Nothing
''   Application.Calculate
''End Sub

''Sub Calculate_Tak_EgenInkomst()
''    'Typisar tidigare Egen inkomst
''
''    Dim iRow As Integer
''    Dim lngRow As Integer
''    Dim rng As Range
''    Dim r As Range
''    Dim lngCol As Integer
''    Dim lngOffsetCol As Integer
''
''    lngOffsetCol = 0
''    iRow = 1
''    Set rng = wsEgenInkomst.Range("AL15")
''    Set rng = rng.Resize(lngAgeMax, 1)
''
''    For Each r In rng
''        If r.Row > 15 Then
''            'Kolumn AL (orgina kolumn BA)
''            'Range(BA16) =PGI(Brutto!B16;10^6;Brutto!I16;Brutto!G16;Brutto!H16)
''            r.Value = PGI(CLng(Arr(iRow, 2)), 1000000, wsData.Cells(r.Row, Application.Range("rng_Data_PBB_Header").Column), _
''                    wsData.Cells(r.Row, Application.Range("rng_Data_IBB_Header").Column), wsData.Cells(r.Row, Application.Range("rng_Data_FBB_Header").Column), _
''                    lngMarginal, 0, wsData.Cells(r.Row, Application.Range("rng_Data_Top_Left").Column), 0)
''        End If
''        iRow = iRow + 1
''    Next r
''    Application.Calculate
''End Sub


