Attribute VB_Name = "mdlIndataInputOutput"
Option Explicit
Dim mastrIndataLabel() As Variant
            
Sub testWait()
    Dim T As Double
    T = Timer
    InputXGetY
    T = Timer - T
    Stop
End Sub
Sub WaitIfCalculationStateIsNotDone(Optional blnForceFullRecalculaltion As Boolean = False)
    'Checking calculationState and resolves if needed.
    Dim xx As Long
    Const dblWaitSeconds As Double = 0.1
    Dim T As Double
    
    'Testing reset variables changed during save
'    If pblnCloseOrSave = True Then pblnCloseOrSave = False
'    If pblnDisableListClick = True Then pblnDisableListClick = False
    
    If blnForceFullRecalculaltion Then Application.CalculateFullRebuild
    
    Do While Application.CalculationState <> xlDone
        wsIndata.Cells.SpecialCells(xlCellTypeFormulas).Calculate
        Application.CalculateFullRebuild
        T = Timer()
        Do While Timer() - T < dblWaitSeconds
            DoEvents
        Loop
        xx = xx + 1
        If xx > 1 Then
            Stop
            Debug.Print xx 'Should never happen, the full rebuild should fix the calculation state
        End If
    Loop
End Sub
Sub InputXGetY()
    If pblnCloseOrSave = True Then pblnCloseOrSave = False 'Stability
    'Debug.Print "startar körning " & Now
    WaitIfCalculationStateIsNotDone
    cases = 1 'Är definierad i modul vba_go som public för att fånga antal fall
    
    Dim wsMod As Worksheet
    Dim wsDta As Worksheet
    Set wsMod = Worksheets(wsStart.Name)
    Set wsDta = Worksheets(wsIndata.Name)
    
    Dim rngXdta As Range '#antal variabler som läses in
    Dim rngRow As Range 'Aktuell rad som läses
    Dim lngFirstRow As Long 'Första rad som ska läsas
    Dim rngUntilRow As Range 'Sista rad som ska läsas
    Dim blnRowConditionEnded As Boolean '
    Dim aColumnNamedRange() As String 'Variabelnamn
    Dim aColumnNamedRangeValidDiffer() As String 'Om valideringsrangen skiljer sig, händer för activeX-kontroller
    Dim c As Long 'Används för radnr
    Dim T1 'Koll av tid
    Dim Ttot As Double ' Tid för hela körning
    Dim dblLastRowRun As Double 'Koll på rad
    Dim dblLastRowRunNumber As Double 'Sista radnr
    Dim dblcnActTypfall As Double 'Vilket fall
    Dim i As Integer 'Heltalsräknare
    Dim iStartRow As Integer 'Startrad...
    Dim rngResult As Range 'Output
    Dim mValue 'Resultatvärden
    Dim cnColResult As Integer 'Antal variabler som ska ut
    Dim blnFirstRowRun As Boolean 'Kontoll av första rad
    Dim rng As Range
    Dim blnLabelWageExists As Boolean   'Om lönevariabeln finns
    Dim blnLabelOwnIncomeListExists As Boolean
    Dim blnRunModelOk As Boolean
    Dim mtommy As Single '0 - slutlön, 1 - Netto 2 - disponibelt
    Dim mtommy12 As Single '0 - Årsbelopp, 1 månadsbelopp
    Ttot = Timer 'Start-tid sekunder
    mtommy = Application.Range("Rng_CompareTo")
    mtommy12 = Application.Range("Rng_belopp12")

    If mtommy = 0 Then
        Application.Range("rngYtopleft") = "Slutlön"
    ElseIf mtommy = 1 Then
        Application.Range("rngYtopleft") = "Nettolön"
    Else
        Application.Range("rngYtopleft") = "Disp."
    End If
   
    Application_Rest
    T1 = Time  'Tidsangivelser
    Application.Range("rng_Abort_run") = False 'Nollställ avbrytning
    Application.Range("rng_Chartbar") = 0 'Nollställ procenträkningen
     'Sätt allt manuellt ingen skärmuppdatering, finns i mdtools
    
    'set worksheet variables
    blnRowConditionEnded = False
    
    Application.Range("rng_Run_From_Indata") = True
    'Define indata range (input)
    Set rngXdta = Application.Range("rngXTopleft") 'För tillf. rad 7 och kolumn 2 för att hitta antal rader o kolumner
    c = rngXdta.Row 'Första raden med etiketter
    
    Set rngXdta = rngXdta.CurrentRegion 'Innebär alla närgränsade ifyllda celler till höger och nedåt ...
    Dim lngAntalIfylldaRader As Long
    lngAntalIfylldaRader = rngXdta.Rows.Count - 1 'new 2024-01-30
    
    'koll av start och slutrader som ska läsas in - Lagt till TL
    If Application.Range("rngExecuteFromRow") <= c Then _
        Application.Range("rngExecuteFromRow").Value = c + 1
    If Len(Trim(Application.Range("rngExecuteUntilRow"))) = 0 Then
            Dim blnUntilRowWasblank As Boolean
            blnUntilRowWasblank = True
            Application.Range("rngExecuteUntilRow").Value = rngXdta.CurrentRegion.Rows.Count + c - 1
    End If
        
    If (Application.Range("rngExecuteUntilRow") > 0 And Application.Range("rngExecuteUntilRow") < Application.Range("rngXTopleft").Row) Then
        MsgBox "Var god och ange värdet större än " & Application.Range("rngXTopleft").Row & " för att Stanna vid rad(Du kan ange 0 för att köra alla rader)"
        Application.Range("rngExecuteUntilRow").Select
        Application_Wakeup
        Exit Sub
    End If
    If Len(Trim(Application.Range("rngExecuteFromRow"))) = 0 Then
        Application.Range("rngExecuteFromRow").Value = c + 1
        'MsgBox "Var god och ange värdet större än " & Application.Range("rngXTopleft").Row & " för att börja vid rad."
        'Application.Range("rngExecuteFromRow").Select
        'Application_Wakeup
        'Exit Sub
    End If
    
    If Application.Range("rngExecuteFromRow") > Application.Range("rngExecuteUntilRow") _
        And (Application.Range("rngExecuteFromRow") > 0 And Application.Range("rngExecuteUntilRow") > 0) Then
         MsgBox "Värdet till Börja vid rad måste vara mindre eller lika med värdet till Stanna vid rad."
        Application.Range("rngExecuteFromRow").Select
        Application_Wakeup
        Exit Sub
    End If
    'Antal typfall/Rader att läsa
    If Application.Range("rngExecuteUntilRow") = 0 Then
        dblcnActTypfall = rngXdta.CurrentRegion.Rows.Count - 1 - Application.Range("rngExecuteFromRow") + Application.Range("rngYtopleft").Row + 1
    Else
        dblcnActTypfall = Application.Range("rngExecuteUntilRow") - Application.Range("rngExecuteFromRow") + 1
    End If
    
    If dblcnActTypfall > 1 Then
       ' dblcnActTypfall = dblcnActTypfall
    Else: dblcnActTypfall = 0
    dblcnActTypfall = 1
    End If
    
    With Worksheets(Range("rngExecuteFromRow").Parent.Name)
        If Not fnValidateRegion(.Range("rngExecuteFromRow"), 2, .Range("rngExecuteFromRow") + dblcnActTypfall - 1, 4, Application.Range("rngExecuteFromRow").Parent.Name) Then
            Set rngXdta = Nothing
            Application_Wakeup
            Exit Sub
        End If
    End With
    
    'obs om tomma celler exit... koll för
    'Fyll Array med kolumnens etikett som senare ersättas med variabelnamnet
    Erase mastrIndataLabel 'Töm (blanka) ifall ändring sen senast
    ReDim aColumnNamedRange(1 To rngXdta.Columns.Count) 'Antal Input kolumner
    ReDim aColumnNamedRangeValidDiffer(1 To rngXdta.Columns.Count)
    
''    For c = 1 To rngXdta.Columns.Count  'Läs in etiketterna
''        aColumnNamedRange(c) = fngetNamedRangeForLabel(rngXdta.Cells(1, c).Value)
''        aColumnNamedRangeValidDiffer(c) = fngetNamedRangeForLabelIfValidDiffer(rngXdta.Cells(1, c).Value)
''        Debug.Print "Läser in i tur och ordning: " & aColumnNamedRange(c)
''    Next

    
    If Application.Range("rngExecuteFromRow") > Application.Range("rngXTopleft").Row Then 'OBS idag rad 7
        Set rngXdta = rngXdta.Offset(Application.Range("rngExecuteFromRow") - _
            Application.Range("rngXTopleft").Row).Resize(dblcnActTypfall)
            'debug.print rngXdta.address
    Else
        Set rngXdta = rngXdta.Offset(1).Resize(rngXdta.Rows.Count - 1) 'Ska inte kunna ske
    End If
    lngFirstRow = rngXdta.Cells(1, 1).Row                       'Första radnumret
    Set rngUntilRow = Application.Range("rngExecuteUntilRow")   'och sista radnumret
    
    If rngUntilRow.Value < 1 Then           'Kan inte förekomma
        dblLastRowRun = rngXdta.Rows.Count
        dblLastRowRunNumber = dblLastRowRunNumber + Application.Range("rngYtopleft").Row
    Else
        dblLastRowRun = rngUntilRow.Value - wsIndata.Range("rngExecuteFromRow") + 1 'Antal typfall
        dblLastRowRunNumber = Application.Range("rngExecuteUntilRow")               'Och sista finns på radnr
    End If
'''    Application.Calculation = xlCalculationManual 'Do not calculate during input - slows down

    'Rensa outputen för de typisar som ska beräknas
    Dim rng2 As Range
    Set rng2 = wsIndata.Range("rngYtopleft") 'Första cellen bland etiketterna som skrivs ut är cellen "Slutlön"
    cnColResult = rng2.CurrentRegion.Columns.Count  'Antal variabler som ska skrivas ut
    If rng2.CurrentRegion.Rows.Count > 1 Then
        Set rng2 = rng2.Offset(Application.Range("rngExecuteFromRow") - Application.Range("rngXTopleft").Row).Resize(dblLastRowRun, rng2.CurrentRegion.Columns.Count)
        rng2.ClearContents
    End If
    
    'Rensa bort överflödiga resultatrader, vilket händer om antalet inputrader har minskat 2024-01-30
    Set rng2 = wsIndata.Range("rngYtopleft").CurrentRegion
    Set rng2 = rng2.Offset(lngAntalIfylldaRader + 1).Resize(1000)
    rng2.ClearContents
    
    Set rng2 = Nothing 'Frigör minne
    
    Dim Kol As Integer
    Kol = Application.Range("rngYtopleft").Column - 1 'Start på utdata
    pblnCloseOrSave = True  'Excel går igenom ... när cellvärden ändras
    Call Unprotect_sheet(wsStart.CodeName) 'För att kunna skjuta in värden
    
    wsIndata.Activate 'Återgå till indata
    
    For Each rngRow In rngXdta.Rows 'För varje rad ...
        Application.Range("rng_Egen_Lon") = False 'Utifall ta bort egen lönelista
        'Application_Wakeup
        
        If Application.Range("rng_Abort_run") = True Then 'Koll om nödstopp
            MsgBox "Körningen kommer att avbrytas!"
            Application.Range("rng_Abort_run") = False
            Application.Range("rng_Run_From_Indata") = False
            Call Protect_sheet(wsStart.CodeName)
            Application.StatusBar = ""
            Set rngUntilRow = Nothing
            Set wsMod = Nothing
            Set wsDta = Nothing
            Set rngXdta = Nothing
            Application_Wakeup
            Exit Sub
        End If

'        Application_Rest 'Manuellt...
        
        If rngRow.Row = Application.Range("rngExecuteFromRow") Then
            blnFirstRowRun = True 'för senare meddelanden
        End If
        
        blnLabelOwnIncomeListExists = False
        blnLabelWageExists = False
        
        T1 = Time

        If Cells(rngRow.Row, 2) < 1930 Or Cells(rngRow.Row, 2) > 2100 Then
           MsgBox "Fel i radens födelseår"
           Exit Sub
        End If
        
        '**** Denna rad har ändrat pga Börja arbeta kan vara mindre än 15
        'If Cells(rngRow.Row, 3) < 15 Or Cells(rngRow.Row, 3) > 100 Then
        If Cells(rngRow.Row, 3) < 0 Then
           MsgBox "Fel i raden: Börjar arbeta"
           Exit Sub
        End If
        If Cells(rngRow.Row, 4) < 61 Or Cells(rngRow.Row, 4) > 100 Then
           MsgBox "Fel i raden: Pensionsålder"
           Exit Sub
        End If
        If Cells(rngRow.Row, 3) > Cells(rngRow.Row, 4) Then
           MsgBox "Inkonsekevens Start år och pensionsålder i raden"
           Exit Sub
        End If
        
        ' .............Marginal..............
        'Funkar inte i Outputen
        Dim marg As Integer
        marg = 0 '1000 'OBS per månad
        Dim ber As Integer
        Dim jber As Integer
        
        ber = 1
        If marg > 0 Then
           cnColResult = cnColResult + 1
           ber = 2
        End If
        
        ReDim mValue(1 To cnColResult)
        Dim alt_p_age As Integer
        alt_p_age = Application.Range("Alt_p_age")
        
        For jber = 1 To ber
           Application.Range("BornYear") = Cells(rngRow.Row, 2)
           
           
           Application.Range("wStartYear") = Cells(rngRow.Row, 3)
           'Debug.Print 1060
           Application.Range("PARYear") = Cells(rngRow.Row, 4)
           If alt_p_age = 1 Then
               Application.Range("PARYear") = Worksheets("Nyckeltal").Cells(5 + Cells(rngRow.Row, 2) - 1930, 120)
           ElseIf alt_p_age = 2 Then
               Application.Range("PARYear") = Worksheets("Nyckeltal").Cells(5 + Cells(rngRow.Row, 2) - 1930, 122)
           Else
               Application.Range("PARYear") = Cells(rngRow.Row, 4)
           End If
           
           Application.Range("Wage_Monthly") = Cells(rngRow.Row, 5) / 12 + (jber - 1) * marg 'Månadslön
           'Debug.Print 1061
           Application.Range("rng_Yearly_Inflation") = Cells(rngRow.Row, 6)
           Application.Range("rng_Real_Growth") = Cells(rngRow.Row, 7)
           Application.Range("rng_FondAvkastning") = Cells(rngRow.Row, 8)
           If Cells(rngRow.Row, 9) > 1 Then
               Application.Range("rng_Egen_Lon") = True
               Application.Range("rng_EgenLönelista_Index") = Cells(rngRow.Row, 9)
           End If
           Application.Range("IPS") = Cells(rngRow.Row, 10)
           Application.Range("rng_TJP_Listbox") = Cells(rngRow.Row, 11)
           'Debug.Print 1062
            
           blnRunModelOk = True
           If rngRow.Row = Application.Range("rngExecuteFromRow") Then
               'Call CheckDefaultValueNeed
               Application.Range("rng_Indata_infocell") = "Var god vänta! beräkning pågår..."
           End If
           wsIndata.Range("rng_Current_Row_Indata") = rngRow.Row
           '******
           pblnCloseOrSave = False
           'If blnRunModelOk Then
           'Debug.Print 1063
           
           StartUp_Indata (blnFirstRowRun)
           'Debug.Print 1064
        
           Application_Rest
           'Debug.Print 1065
           'om ändrad input fixas här----'
           blnFirstRowRun = False
           pblnCloseOrSave = True
           '********** Get output Y-data - Values *********
           Set rng = ThisWorkbook.Names("rngTopXYOutput").RefersToRange 'Nu rad 28 col 4, jfr D27 i startfliken
           
           ' Resultat som Array
           'Debug.Print 107
           For i = 1 To cnColResult
               Select Case i
                   Case 1
                       If mtommy = 0 Then
                           iStartRow = -3  'Lön t-1
                       ElseIf mtommy = 1 Then
                           iStartRow = -2  'Netto t-1
                       Else
                           iStartRow = -1  'Disp t-1
                       End If
                   Case 2
                       iStartRow = 9   'Total pension
                   Case 3
                       iStartRow = 1   'InkomstP
                   Case 4
                       iStartRow = 2   '(A)TP
                   Case 5
                       iStartRow = 3   'PremieP
                   Case 6
                       iStartRow = 4   'Garp
                   Case 7
                       iStartRow = 5   'Ptillagg
                   Case 8
                       iStartRow = 7   'Tjänstep
                   Case 9
                       iStartRow = 8   'IPS
                   Case 10
                       iStartRow = 15  'Efter skatt
                   Case 11
                       iStartRow = 16  'Bidrag
                   Case 12
                       iStartRow = 18  'Disp
               End Select
               If jber = 1 Then
                   If mtommy12 = 0 Then
                       mValue(i) = rng.Cells(iStartRow, 1)
                   Else
                       mValue(i) = rng.Cells(iStartRow, 1) / 12
                   End If
               ElseIf jber = 2 And i = 12 Then
                   If mtommy12 = 0 Then
                       mValue(12) = 1 - (rng.Cells(16, 1) - mValue(11)) / _
                       (rng.Cells(8, 1) - rng.Cells(4, 1) - (mValue(2) - mValue(6)))
                   Else
                   
                       mValue(12) = 1 - (rng.Cells(16, 1) / 12 - mValue(11)) / _
                       ((rng.Cells(8, 1) - rng.Cells(4, 1)) / 12 - (mValue(2) - mValue(6)))
                       Debug.Print rng.Cells(16, 1) / 12; (rng.Cells(16, 1) / 12 - mValue(11)); (rng.Cells(8, 1) - rng.Cells(4, 1)) / 12 - (mValue(2) - mValue(6))
                   End If
        
               End If
               
               'Redigera nedan för att skriva ut något värde i fliken Indata
               'wsIndata.Cells(rngRow.Row, 26) = wsStart.Range("F42")
           Next i
        Next jber
        
        'Debug.Print 108
           
        If cases = 0 Then Exit Sub
        cases = cases + 1 'Public in VBA_go
        Set rngResult = rngRow.Cells(1, Kol).Resize(1, cnColResult)
        rngResult.Value = mValue
        Set rngResult = Nothing
        
        If Application.Range("Risk").Value > 0 Then
            copyRnd 'beräkna nya slumptal
        End If
        Application_Wakeup
        Application.Range("rng_Chartbar") = (rngRow.Row - lngFirstRow + 1) / dblLastRowRun
        Application_Rest
        wsIndata.Range("rngRunTime") = Time - T1
        Application.StatusBar = Format(((rngRow.Row - lngFirstRow + 1) / dblLastRowRun), "0%") & " Klar. Aktuell rad " & rngRow.Row & " (Sista rad är: " & dblLastRowRunNumber & ")" & _
                                 " " & "Tid för senaste beräkning: " & Format(wsIndata.Range("rngRunTime"), "hh:mm:ss")
        DoEvents
        'Debug.Print 109
        'Application.Wait TimeSerial(Hour(Now), Minute(Now), Second(Now) + 2)
        If rngUntilRow.Value <= rngRow.Cells.Row And rngUntilRow <> 0 Then
            'Avbryt om UntilRow<>0 och radnr överskrids
            blnRowConditionEnded = True
            Exit For
        End If
        If rngRow.Row = Application.Range("rngExecuteFromRow") Then
            Application.Range("rng_Indata_infocell") = "Beräkning av " & dblLastRowRun & " 'typfall' uppskattas att bli klar vid: " & vbCrLf & Format(Now() + dblLastRowRun * (wsIndata.Range("rngRunTime") + TimeValue("0:00:01")), "hh:mm") ' + (Time - T1)
        End If
        
        'Debug.Print 110
        WaitIfCalculationStateIsNotDone
    Next 'Typfall
    
    Application_Rest
   ' Call Fix_Chart_Start
    'Debug.Print 111
    Application.Range("rng_EgenLönelista_Index").Value = 1
    Application.Range("rng_Egen_Lon") = False
    Call Protect_sheet(wsStart.CodeName)
    'Call CleanArr
    Set rngUntilRow = Nothing
    Erase aColumnNamedRange
    wsIndata.Activate
    Application_Wakeup
'ResetExcelEnvironment:

'''    'reset calculation
    wsIndata.Range("rng_Current_Row_Indata") = ""
    Application.StatusBar = ""
    wsIndata.Range("AA1") = Time
    Ttot = Timer - Ttot
    If blnRowConditionEnded Then
        MsgBox "Körning avslutad på rad " & rngRow.Cells.Row & " på " & Round(Ttot, 1) & " sekunder", vbInformation
    Else
        'MsgBox "Körning klar! " & rngXdta.Rows.Count & " rader beräknade."
        MsgBox "Körning klar! " & rngXdta.Rows.Count & " rader beräknade på " & Round(Ttot, 1) & " sekunder", vbInformation
        Set rngXdta = Nothing
    End If
    
    On Error Resume Next
    'wsPivo.PivotTables("pvtIndata").RefreshTable
    On Error GoTo 0
    wsIndata.Activate
    wsIndata.Range("rngRunTime") = ""
    Application.Range("rng_Indata_infocell") = ""
    Application.Range("rng_Run_From_Indata") = False
    Application.Range("rng_Chartbar") = ""

    'Keep until row blank if it was blank
    If blnUntilRowWasblank Then
        Application.Range("rngExecuteUntilRow").Value = ""
    End If
    Set wsMod = Nothing
    Set wsDta = Nothing
    Set rng = Nothing
    Set rngXdta = Nothing
End Sub


Sub subGetStartSheetValue()
    '
    Dim wsMod As Worksheet
    Dim wsDta As Worksheet
    Dim rnglbl As Range
    Dim cll As Range
    Dim c As Long
    Application_Rest
    Set wsMod = ThisWorkbook.Worksheets(wsStart.Name)
    Set wsDta = ThisWorkbook.Worksheets(wsIndata.Name)
   
    'Define X-data range (input)
    Set rnglbl = Application.Range("rngXTopleft") 'För tillf. rad 6 och kolumn 2
    Set rnglbl = rnglbl.CurrentRegion.Rows(1)

    'Fyll Array med NamedRange beroende på columnens label
    Erase mastrIndataLabel 'Töm ifall ändring sen senast
    ReDim aColumnNamedRange(1 To rnglbl.Columns.Count)
    For c = 1 To rnglbl.Columns.Count
        aColumnNamedRange(c) = fngetNamedRangeForLabel(rnglbl.Cells(1, c).Value, True)
        'Debug.Print C; aColumnNamedRange(C)
    Next

    For c = 1 To rnglbl.Columns.Count
        'If C = rnglbl.Columns.Count Then Stop
        rnglbl.Cells(1, c).Offset(1, 0).Value = fnGetRangeFromNamesTw(aColumnNamedRange(c)).Value
    Next
    Application_Wakeup
    Set rnglbl = Nothing
    Set wsMod = Nothing
    Set wsDta = Nothing

End Sub

' Kolla indata
Function fnIsAllIndataOK(Optional ByVal blnMarkError As Boolean = False) As Boolean
    'Går igenom infylld indata och returnerar True om allt ok, annars false.
    'Optional: att markera felaktig indata
    Dim rng As Range
    Dim col As Range
    Dim a
    Dim cll As Range
    
    Set rng = Application.Range("rngXTopleft")
    Set rng = rng.CurrentRegion
    Set rng = rng.Resize(rng.Rows.Count - 1, rng.Columns.Count).Offset(1, 0)

    fnIsAllIndataOK = True

    For Each col In rng.Columns
        'Load validation list
        For Each cll In col.Cells
            If cll.Validation.Value = False Then
                If blnMarkError Then
                    'markera fel, avbryt inte.
                    cll.Interior.Color = 9737946 'Red
                    fnIsAllIndataOK = False
                Else
                    fnIsAllIndataOK = False
                    Exit Function 'avbryt
                End If
            Else
                If cll.Interior.Color = 9737946 Then cll.Interior.Color = 16777215 'Normal
            End If
            
        Next
    Next

    Set rng = Nothing

End Function

Sub subManageNewValidationRanges(ByRef rngChangedLabel As Range)
    Dim strNamedRange As String
    Dim oValidation As Excel.Validation
    'Lägger till validering för aktuell kolumn.
    Dim rng As Range
    Dim strWsName As String
    Dim strFormula1NoName As String
    Const cNbrOfRows As Long = 2000
    
    
    Set rng = rngChangedLabel.Offset(1, 0).Resize(cNbrOfRows, 1)
    If fnHaveCurRangeValidationList(rng) Then rng.Validation.Delete   'Ta bort existerande Validering (om den finns)
    
    strNamedRange = fngetNamedRangeForLabel(rngChangedLabel.Value, True)
    rng.NumberFormat = fnGetRangeFromNames(ThisWorkbook, strNamedRange).NumberFormat 'Fix numberformat
    
    If fnHaveCurRangeNameValidationList(strNamedRange) Then
        
        Set oValidation = fnGetRangeFromNames(ThisWorkbook, strNamedRange).Validation
        
        If oValidation.Type = xlValidateList Then
        
            rng.Validation.Delete
            strWsName = fnGetRangeFromNames(ThisWorkbook, strNamedRange).Parent.Name
            strFormula1NoName = "=" & strWsName & "!" & Mid(oValidation.Formula1, 2)
            
            With rng.Validation
                If InStr(1, oValidation.Formula1, ":", vbTextCompare) = 0 Then
                    'Validation with namedRange
                    .Add Type:=oValidation.Type, AlertStyle:=oValidation.AlertStyle, Operator:=oValidation.Operator, Formula1:=oValidation.Formula1
                Else
                    'validation with Adress (e.g A4:A8)
                    Stop
                    .Add Type:=oValidation.Type, AlertStyle:=oValidation.AlertStyle, Operator:=oValidation.Operator, Formula1:="=" & strWsName & "!" & Mid(oValidation.Formula1, 2)
                End If
                .IgnoreBlank = oValidation.IgnoreBlank
                .InCellDropdown = oValidation.InCellDropdown
                .InputTitle = oValidation.InputTitle
                .ErrorTitle = oValidation.ErrorTitle
                .InputMessage = oValidation.InputMessage
                .ErrorMessage = oValidation.ErrorMessage
                .ShowInput = oValidation.ShowInput
                .ShowError = oValidation.ShowError
    
            End With
        Else
            'Copy validation that is not lists
            fnGetRangeFromNames(ThisWorkbook, strNamedRange).Copy
            rng.PasteSpecial xlPasteValidation
            Application.CutCopyMode = False

        End If
        
    End If
    
End Sub

Function fnHaveCurRangeNameValidationList(ByVal strRngName As String) As Boolean

    Dim currentValidation As Excel.Validation
    Dim targetRange As Excel.Range
    Dim validationType As Excel.XlDVType
    
    On Error Resume Next
    Set currentValidation = fnGetRangeFromNamesTw(strRngName).Validation
    If currentValidation.Type = xlValidateList Then
        'dummy to trigger error
    End If
    If err.Number > 0 Then
        fnHaveCurRangeNameValidationList = False
        Exit Function
    End If
    
    'If (currentValidation.Type = xlValidateList) Then
        fnHaveCurRangeNameValidationList = True
    'Else
    '    fnHaveCurRangeNameValidationList = False
    'End If
    On Error GoTo 0
    
End Function
Function fnHaveCurRangeValidationList(ByRef rng As Range) As Boolean

    Dim currentValidation As Excel.Validation
    Dim targetRange As Excel.Range
    Dim validationType As Excel.XlDVType
    
    On Error Resume Next
    Set currentValidation = rng.Validation
    
    
    If (currentValidation.Type = xlValidateList) Then
        fnHaveCurRangeValidationList = True
    Else
        fnHaveCurRangeValidationList = False
    End If
    On Error GoTo 0
    Set currentValidation = Nothing
End Function
Function fnGetValidationRange(ByVal strRngName As String) As Excel.Range
    'Anropa först fnHaveCurRangeNameValidationList för att säkerställa att aktuell range har en validation list. Annars körfel
    Dim currentValidation As Excel.Validation
    Dim targetRange As Excel.Range
    
    Set currentValidation = fnGetRangeFromNamesTw(strRngName).Validation  ' check for no existing validation, or multiple validation criteria
    
    If (currentValidation.Type = xlValidateList Or currentValidation.Type = xlValidateCustom) Then    ' test for range reference and not a list of values
    On Error Resume Next
    Set targetRange = Excel.Range(currentValidation.Formula1)
    On Error GoTo 0
    If Not targetRange Is Nothing Then
        Set fnGetValidationRange = targetRange
        End If
    End If
    Set targetRange = Nothing
End Function


Sub subAddValidationForColumnIfNeeded(ByVal rngLabelCell As Range)
    'Lägger till validering för aktuell kolumn.
    Dim rng As Range
    Const cNbrOfRows As Long = 2000
    
    Set rng = rngLabelCell.Offset(1, 0).Resize(cNbrOfRows, 1)
    
    With rng.Validation
        .Delete
        .Add xlValidateList
    End With
    Set rng = Nothing
End Sub



Function fngetNamedRangeForLabel(ByVal strLabel As String, Optional blnReturnValidationRangeIfDifferent As Boolean = False) As String
    'Returnerar Namngivet område för aktuell kolumn-etikett
    'Om aktuellt område har en valideringsrange som är skild från aktuella området (gäller t ex vid listbox) returneras denna om blnReturnValidationRangeIfDifferent sätt till sann
    Dim rng As Range
    Dim r As Long
    
    On Error Resume Next ' tells VBA to ignore the error and continue on
    If UBound(mastrIndataLabel) > 0 Then
        'Dummy
    End If
    If err.Number > 0 Then
        'Debug.Print Err.Description
        
        On Error GoTo 0
        'Initiera Lookup-array
        Set rng = fnGetRangeFromNamesTw("rngIndataLabel") 'Se NameRange j19:j29
        Set rng = rng.Resize(rng.Rows.Count, 3)
        ReDim mastrIndataLabel(1 To rng.Rows.Count, 1 To rng.Columns.Count)
        mastrIndataLabel = rng.Value
    End If
    On Error GoTo 0
    
    For r = 1 To UBound(mastrIndataLabel, 1)
        If mastrIndataLabel(r, 1) = strLabel Then
            If blnReturnValidationRangeIfDifferent Then
                fngetNamedRangeForLabel = mastrIndataLabel(r, 3)
                If fngetNamedRangeForLabel = "" Then fngetNamedRangeForLabel = mastrIndataLabel(r, 2)
            Else
                fngetNamedRangeForLabel = mastrIndataLabel(r, 2)
            End If
        End If
    Next
    Set rng = Nothing
End Function

Function fngetNamedRangeForLabelIfValidDiffer(ByVal strLabel As String) As String
    'Returnerar sant/falskt beroende om det  är värde eller index för Namngivet område för aktuell kolumn-etikett
    'Detta behövs för active-x-kontroller som refererar till ett index, inte till värdet som matas in.
    Dim rng As Range
    Dim r As Long
    
    On Error Resume Next
    If UBound(mastrIndataLabel) > 0 Then
        'Dummy
    End If
    If err.Number > 0 Then
        On Error GoTo 0
        'Initiera Lookup-array
        Set rng = fnGetRangeFromNamesTw("rngIndataLabel")
        Set rng = rng.Resize(rng.Rows.Count, 3)
        ReDim mastrIndataLabel(1 To rng.Rows.Count, 1 To rng.Columns.Count)
        mastrIndataLabel = rng.Value
    End If
    On Error GoTo 0
    
    For r = 1 To UBound(mastrIndataLabel, 1)
        If mastrIndataLabel(r, 1) = strLabel Then
            fngetNamedRangeForLabelIfValidDiffer = mastrIndataLabel(r, 3)
        End If
    Next
    
    Set rng = Nothing
End Function



Sub RecalculateFullrebuild()
    Application.Calculation = xlCalculationAutomatic 'reset calculation
    Application.CalculateFullRebuild
    Application.StatusBar = ""
    MsgBox "Calculation finished!"
End Sub

Sub CheckDefaultValueNeed()
    If Application.Range("BornYear") = "" Then
        Application.Range("BornYear") = Application.Range("BornYear_Default")
    End If
    If Application.Range("PARYear") = "" Then
        Application.Range("PARYear") = Application.Range("PARYear_Default")
    End If
    If Application.Range("wStartYear") = "" Then
        Application.Range("wStartYear") = Application.Range("wStartYear_Default")
    End If
    If Application.Range("Wage_Monthly") = "" Then
        Application.Range("Wage_Monthly") = Application.Range("Wage_Monthly_Default")
    End If
    If Application.Range("rng_FondAvkastning") = "" Then
        Application.Range("rng_FondAvkastning") = Application.Range("rng_FondAvkastning_Default")
    End If
    If Application.Range("rng_Real_Growth") = "" Then
        Application.Range("rng_Real_Growth") = Application.Range("rng_Real_Growth_Default")
    End If
    If Application.Range("rng_Real_Growth") = "" Then
        Application.Range("rng_Real_Growth") = Application.Range("rng_Real_Growth_Default")
    End If
End Sub
