Attribute VB_Name = "mdlModellMode"
Option Explicit
Sub Modell_Normal()
    Application_Rest
    
    Call User_mode
    Application.Range("rng_Application_Mode") = 0
    Call Show_Hide_Chart_Priser(False)
    Call getFormValue
    Call MonthCellsFontColor
    
    Application_Wakeup
End Sub
Sub Modell_Advanced()
    Application_Rest
    
    Call Advanced_mode
    Application.Range("rng_Application_Mode") = 1
    Call getFormValue
    Call MonthCellsFontColor
    Call Show_Hide_Chart_Priser(True)
    
    Application_Wakeup
End Sub

Sub OptionButton225_Click()
    Application.Range("rng_Chart_Earning_factor") = 1
End Sub
Sub OptionButton226_Click()
    Application.Range("rng_Chart_Earning_factor") = 12
End Sub

Sub Advanced_mode()
    Dim wk As Worksheet
    Dim iRow As Integer
    iRow = Application.Range("rng_Gift_Row").Row
    With wsStart
        Call Developers_Sheet
        .Activate
        .Unprotect
'        .Rows("" & iRow - 1 & ":" & iRow + 3 & "").Select
'         Selection.EntireRow.Hidden = False
        .Rows("" & iRow - 1 & ":" & iRow + 3 & "").EntireRow.Hidden = False
         .Shapes("chkGift").Visible = True
        ' .Shapes("chkOwnVecktor").Visible = True
        .Range("rngModelHeader").Select
        .Protect
    End With

End Sub
Sub User_mode()
    Dim wk As Worksheet
    Dim iRow As Integer
    iRow = Application.Range("rng_Gift_Row").Row

    With wsStart
        Call User_Sheet
        .Activate
        .Unprotect
'         .Rows("" & iRow & ":" & iRow + 1 & "").Select
'         Selection.EntireRow.Hidden = True
         .Rows("" & iRow & ":" & iRow + 1 & "").EntireRow.Hidden = True
         .Shapes("chkGift").Visible = False
         '.Shapes("chkOwnVecktor").Visible = False
        .Range("rngModelHeader").Select
        .Protect
    End With

End Sub

Sub Developers_Sheet()
    Dim rng As Range
    Dim r As Range
    Dim wk As Worksheet
    Set rng = Application.Range("rngDeActivatedSheets")
    Set rng = rng.Offset(1).Resize(rng.CurrentRegion.Rows.Count - 1, 1)
    For Each wk In ThisWorkbook.Worksheets
        For Each r In rng
            If LCase(wk.Name) = LCase(r.Value) Then
                  wk.Visible = xlSheetHidden
                  Exit For
            Else
                If Not wk.Visible Then wk.Visible = xlSheetVisible
            End If
        Next r
    Next wk
    Set rng = Nothing
    Set r = Nothing
End Sub

Sub User_Sheet()
    Dim rng As Range
    Dim r As Range
    Dim wk As Worksheet
    Set rng = Application.Range("rngUserSheets")
    Set rng = rng.Offset(1).Resize(rng.CurrentRegion.Rows.Count - 1, 1)
    For Each wk In ThisWorkbook.Worksheets
        For Each r In rng
            If LCase(wk.Name) = LCase(r.Value) Then
                  wk.Visible = xlSheetVisible
                  Exit For
            Else
                If wk.Visible Then wk.Visible = xlSheetHidden
            End If
        Next r
    Next wk
    Set rng = Nothing
    Set r = Nothing
End Sub

Sub OwnVectors()
    If wsNameRange.Range("rng_Egen_Lon") Then
         'MsgBox ("Egen inkomstvektor eller välja av färdiga typisar")
         'OBS färdiga typisar finns i arket "Typfall"
        
         Rem Innan vi går till sidan med egen inkomst. Beräkna löner mm till den egna vektorn
         'samt börja med att "nollställa" egna typfall (=1)
         Application.Range("rng_EgenLönelista_Index") = 1
          
         Call startsetup 'Ger lönevektorn mm
         
         Dim mvalues(15 To slutage, 1 To 4) As Double
         Dim txt(15 To slutage) As String
         
         For age = 15 To slutage
               'mvalues(age, 1) = PBB(age)
               mvalues(age, 1) = Int(born) + age
               mvalues(age, 2) = age
               mvalues(age, 3) = wages(age, year_(age), W_start, Income, w_time, w_ref, Nominal)
               mvalues(age, 4) = mvalues(age, 3)
               txt(age) = "=om(och(satagere>1960;satagare<=B" & age & "0);0;D" & age & ")"
               'Debug.Print txt(age)
         Next age
         
         Worksheets("Inkomstvektor").Range("B15:e" & slutage) = mvalues
        ' Worksheets("Indata_lista").Range("E15:E" & slutage) = txt
        With wsInkomstvektor
            .Visible = xlSheetVisible
            .Select
            .Range("rng_Top_earn_array").Select
        End With
        
    'Else
         'wsIndata_Lista.Visible = xlSheetHidden
    End If
'    Application_Wakeup
'    Erase Arr

End Sub

Sub Update_IPS()
    Dim rngFrom As Range
    Dim r As Range
    Dim PAR As Long
    Dim lngMax As Long
    Dim Arr()
    Dim lngOffsetRow As Integer
    Dim lngRow As Long
    Dim iColYear As Integer
    Application_Rest
    lngMax = Application.WorksheetFunction.Count(wsInkomstvektor.Range("A15:A10000"))
    ReDim Arr(1 To lngMax, 1 To 2)

    lngOffsetRow = Application.Range("rng_Top_Year_array").Row - 1
    iColYear = Application.Range("rng_Top_Year_array").Column
    With wsInkomstvektor
        For lngRow = 1 To lngMax
            Arr(lngRow, 1) = .Cells(lngRow + lngOffsetRow, iColYear)
            Arr(lngRow, 2) = .Cells(lngRow + lngOffsetRow, iColYear + 1)
        Next lngRow
    End With

    IPS_year = Application.Range("IPS_year")
    PAR = Application.Range("PAR")
    Set rngFrom = wsInkomstvektor.Range("rng_Top_Private_save_array")
    For Each r In rngFrom.Resize(lngMax, 1)
        If Arr(r.Row - rngFrom.Row + 1, 1) >= IPS_year And Arr(r.Row - rngFrom.Row + 1, 2) < PAR Then
            r.Value = Application.Range("IPS")
        End If
    Next r
    Set rngFrom = Nothing
    Application_Wakeup
    Erase Arr
End Sub
''
''Sub ShowTaxRows_Start()
''    If Application.Range("rng_ShowTaxRows") Then
''        Call ShowTaxRows
''        Application.Range("rng_ShowTaxRows") = False
''    Else
''        Call HideTaxRows
''        Application.Range("rng_ShowTaxRows") = True
''    End If
''End Sub

