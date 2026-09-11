Attribute VB_Name = "Övrigt_"
'Subrutiner ...Fnns en räntafunktion kallad koll samt PW
Option Explicit
Option Base 1
Public timerflag As Integer  'Status på R

Sub QuickSort(Arr, LO As Long, Hi As Long)
  Dim varPivot As Variant
  Dim varTmp As Variant
  Dim tmpLow As Long
  Dim tmpHi As Long
  tmpLow = LO
  tmpHi = Hi
  varPivot = Arr(Int((LO + Hi) \ 2)) 'Det mittersta värdet i vektorn
  
  Do While tmpLow <= tmpHi
    Do While Arr(tmpLow) < varPivot And tmpLow < Hi
      tmpLow = tmpLow + 1
    Loop
    Do While varPivot < Arr(tmpHi) And tmpHi > LO
      tmpHi = tmpHi - 1
    Loop
    If tmpLow <= tmpHi Then
      varTmp = Arr(tmpLow)
      Arr(tmpLow) = Arr(tmpHi)
      Arr(tmpHi) = varTmp
      tmpLow = tmpLow + 1
      tmpHi = tmpHi - 1
    End If
  Loop
  If LO < tmpHi Then QuickSort Arr, LO, tmpHi
  If tmpLow < Hi Then QuickSort Arr, tmpLow, Hi
End Sub

Sub GoToStart()
    wsStart.Activate
End Sub

Sub GoToData()
    wsOutput.Activate
End Sub

'''Sub figur1()
'''    Dim koll As Byte
'''    koll = 0
'''    If Application.Range("Pic1") = 1 Then koll = 1
'''    'ActiveChart.PlotArea.Select
'''    ActiveSheet.Shapes.Range(Array("Picture 1")).Select
'''    If koll = 1 Then
'''        ActiveSheet.ChartObjects("chart_Priser").Activate
'''    Else
'''         ActiveSheet.ChartObjects("chart_Priser").Deactivate
'''    End If
'''
'''End Sub

Sub Listruta1028_Ändra()
    'Att gå till olika delar av arket
    '1-Utfall pensionsdata
    '2- Diagram
    '3-Kontrafaktiska
    '4-Marginal
    '5-Progressivitet
    '6- Gå till start
    '7- Gå till start
    '8-Förenklad kalkyl
    '9- Pensions wealth?
    'Range("a1").Select
   If pblnCloseOrSave Then Exit Sub
    Dim Z As Long
    Dim strGoto As String
    Dim rng As Range
    
    Z = Application.Range("rngNavigeringIndex")
    
    Set rng = Application.Range("rng_Navigeringstabell1_Top")
    
    strGoto = wsNameRange.Cells(Z + rng.Row, rng.Column + 1)
    If strGoto <> "" Then
        Select Case strGoto
            Case "rng_result_Start"
                wsStart.Activate
            Case "Inget"
            Case Else
                
        End Select
    End If
    Set rng = Nothing
 End Sub
 
 
Sub Kryssruta_PGB_Klicka()
  '    MsgBox ("Pensionsgrundande belopp PGB" )
  wsPGB.Visible = True
  
  Dim Z As Byte
    Z = Application.Range("B77").Select
    'Z = "start"
    If ActiveCell.FormulaR1C1 = "FALSE" Then Z = 0
    If ActiveCell.FormulaR1C1 = "TRUE" Then Z = 1
    
     If Z = 1 Then
      'Sheets("PGB").Select
      ActiveCell.FormulaR1C1 = "FALSE"
       wsPGB.Activate
     End If
End Sub
 
Sub Kryssruta_MFT_Klicka()
  '    MsgBox ("Medelfristig prognos" )
  'WsPensioner.Visible = True
  
  Dim Z As Byte
   ' Z = Application.Range("rng_mftid").Select
    Application.Range("rng_mftid") = Not Application.Range("rng_mftid")
      
    'If ActiveCell.FormulaR1C1 = "FALSE" Then Z = 0
    'If ActiveCell.FormulaR1C1 = "TRUE" Then Z = 1
    
    'If Z = 1 Then
      'Sheets("Pensioner").Select
      'ActiveCell.FormulaR1C1 = "FALSE"
     ' WsPensioner.Activate
      
     'End If
End Sub



Function Omsatt(start, slut, Optional growth = 1.015) As Variant
    'Beräkning av omsättningstiden.
    'Eftersom inkomsten växer konstant med growth så är omsatt oberoende sv initial nivå
    If pblnCloseOrSave Then Exit Function
    On Error GoTo errTag
    'Application.Volatile
    Dim talj As Double
    Dim namn As Double
    Dim inkomst As Double
    
    talj = 0
    namn = 0
    Dim i As Long
    For i = start To (slut - 1)
        inkomst = 100 * growth ^ (i - 1)
        talj = talj + (inkomst * (slut - i - 0.5))
        namn = namn + inkomst
    Next i
    Omsatt = talj / namn
    Exit Function
errTag:
    Omsatt = CVErr(xlErrValue)
End Function
'''
'''Function koll(belopp, start, slut, Optional i = 0.03) As Variant
'''    If pblnCloseOrSave Then Exit Function
'''    On Error GoTo ErrTag
'''    'Application.Volatile
'''    Dim rest As Double
'''    Dim Last As Double
'''
'''    rest = slut - Int(slut)
'''    Do While start <= Int(slut)
'''         koll = koll + belopp * (1 + i) ^ (start)
'''         Last = belopp * (1 + i) ^ (start + 1) 'Det sista året
'''         start = start + 1
'''         'If start = slut Then test = 1
'''        'Exit Do
'''    Loop
'''    'Koll = last * rest
'''    koll = koll + rest * Last
'''    Exit Function
'''ErrTag:
'''        koll = CVErr(xlErrValue)
'''End Function


Function PW_(belopp, slut, Optional d = 1.02) As Variant
    If pblnCloseOrSave Then Exit Function
    On Error GoTo errTag
    'Application.Volatile
    Dim start As Long
    Dim rest As Double
    Dim Last As Double
    Dim koll As Double
    start = 1
    'koll = belopp
    
    rest = (slut - Int(slut)) * 10 / 12 'Antal månader
    Do While start <= Int(slut)
     koll = koll + belopp / d ^ (start - 1)
     Last = belopp / d ^ (start) 'Det sista året
     start = start + 1
     'If start = slut Then test = 1
    'Exit Do
    Loop
    'Koll = last * rest
    PW_ = koll + rest * Last
    Exit Function
errTag:
    PW_ = CVErr(xlErrValue)
End Function

Sub kollpw()
 Dim kollpw As Long
  kollpw = PW_(10, 20, 1.019)
End Sub

''Function deflator(ByVal perioder As Integer, ByVal typ As Integer) As Double
''
''    deflator = 1
''
''    If typ < 0 Then Exit Function
''    If perioder < 2 Then Exit Function
''
''    Dim I As Integer
''    If typ = 1 Then 'KPI
''        deflator = Income_(Int(par)) * KPI(Int(par)) / KPI(Int(par - 1))
''        For I = 2 To perioder
''            deflator = deflator + Income_(Int(par - I)) * KPI(Int(par)) / KPI(Int(par - I))
''        Next I
''        deflator = deflator / perioder
''    Else
''        deflator = Income_(Int(par)) * Iindex(Int(par)) / Iindex(Int(par - 1))
''        For I = 2 To perioder
''            deflator = deflator + Income_(Int(par - I)) * Iindex(Int(par)) / Iindex(Int(par - I))
''        Next I
''        deflator = deflator / perioder
''    End If
''
''End Function

'----------------- R Tur och retur borttaget finns i 2020401 ---------------

Sub mstatus(mText)
    Worksheets("data_till_start").Range("B28") = mText 'Fast adress
End Sub

'Sub mrakn(mtal) 'Räkna antal rader mm
'    If mtal > 0 Then
'        Worksheets("blad1").Range("A28") = mtal 'Fast adress
'    Else
'           Worksheets("blad1").Range("A28") = "Rör inte rad A30..."
'    End If
'End Sub

Function FileExists(FilePath As String) As Boolean
    Dim TestStr As String
    TestStr = ""
    On Error Resume Next
    TestStr = Dir(FilePath)
    On Error GoTo 0
    If TestStr = "" Then
        FileExists = False
    Else
        FileExists = True
    End If
End Function

Sub tabort(fil)
    Dim fil2 As String
    fil2 = "c:\temp\" + fil
    fil = FileExists(fil2)
    
    If fil = False Then
        Call mstatus("file" & fil & " do not exist")
    Else
        Call mstatus("Tar bort " & fil)
        Kill fil2
        'MsgBox (VBA.Len(VBA.Dir(fil2)))
    End If
End Sub


Sub cellcheck() 'Koll av en cell och vad innehållet, format, text belopp mm är
    Dim rcell As Range
    Dim sMyString As String

    On Error GoTo ErrorHandle
    Set rcell = Range("c2")
    
    'We set our range variable = cell C2 in the active sheet.
    'The following control tests if a cell is empty by testing the length of what may be in the cell.
    'Therefore we use the property "Formula" instead of "Value". If you used "Value",
    'it would say the cell is empty, even if it contains spaces/blanks.
    If Len(rcell.Formula) = 0 Then
       MsgBox "Cell " & rcell.Address & " is empty."
    End If

    'You can make the same check using the VBA-function IsEmpty.
    'IsEmpty returns True, if the cell is empty. IsEmpty is better than the control above, when it comes to
    'the content of a cell, but not if it was a string variable. It is a matter of speed.
    If IsEmpty(rcell) Then
       MsgBox "Cell " & rcell.Address & " is empty."
    End If

    'The next example checks if the cell contains a number, i.e. a numeric value.
    If IsNumeric(rcell.Value) Then
        MsgBox "Cell " & rcell.Address & " is a numeric value."
    End If

    'This control checks if a cell contains an error, for instance
    'division by zero or a formula with a reference to a non-existing named cell.
    If IsError(rcell.Value) Then
       MsgBox "Cell " & rcell.Address & " contains an error."
    End If

    'Here we check if a cell (or an expression) is a date.
    If IsDate(rcell.Value) Then
       MsgBox "Cell " & rcell.Address & " is a date."
    End If

    'The following checks, if a cell contains text.
    'If it isn't a numeric value or an error like division by zero, we assume it is text of the data type String.
    'Date values can be mistaken for text, so if you want dates eliminated as well, you must add
    '"IsDate(rCell.Value) = False" to your control.
    If IsNumeric(rcell.Value) = False And _
    IsError(rcell.Value) = False Then
    'The VBA function Trim removes leading and trailing blanks. If the length after trimming is 0, the content was blanks only.
       sMyString = Trim(rcell.Value)
       If Len(sMyString) > 0 Then
          MsgBox "Cell " & rcell.Address & " is a text with " & _
          Len(sMyString) & " characters."
       Else
          MsgBox "The cell contains blanks only"
       End If
    End If

    'Checks if a cell has conditional formatting
    If rcell.FormatConditions.Count > 0 Then
       MsgBox rcell.Address & " has conditional formatting."
    Else
       MsgBox "No conditional formatting."
    End If

    'Checks if a cell contains a formula.
    If rcell.HasFormula Then
       MsgBox "Cell " & rcell.Address & " contains a formula."
    Else
       MsgBox "The cell has no formula."
    End If

    'Checks if a cell has a comment.
    If rcell.Comment Is Nothing Then
       'If not, add a comment.
'       With rcell.AddComment
'          .Visible = False
'          .text "Comment added " & Date
'       End With
    Else
       MsgBox rcell.Address & " has a comment."
    End If

BeforeExit:
    'Set rcell = Nothing
    Exit Sub
ErrorHandle:
    MsgBox err.Description & " Error in procedure CellCheck."
    Resume BeforeExit
End Sub


Sub tocsv()
    'Skriver till csv-fil
    Dim txt As String
    Dim utfil As String
    Dim utdata As String
    utfil = Worksheets("data_till_start").Range("C27") 'Fast adress '"typis.csv"
    utfil = "c:\temp\" & utfil
    'Kolla att "c:\temp finns annars ev. skapa
''  If Dir(utfil) = "" Then
''        mstatus ("C:\temp finns inte avslutar")
''        ExitSub
''    End If
    
    Close #1 'För säkerhetsskull
    Dim MyRange As Range
    Set MyRange = Range("A2:L22") 'Fast adress
    
    Open utfil For Output As #1
    
    Dim i As Integer
    Dim j As Integer
    utdata = ""
    'Inga rubriker - bortkommenterat
'    Dim Xvar As Variant 'Obs fast antal
'    Xvar = Array("Age", "Wage", "w_", "ip", "pp", "atp", "gp", "tjp", "gross", "net", "disp", "cont")
'    For i = 1 To myrange.Columns.Count
'        'utdata = utdata & chr(34) & Xvar(i) & chr(34) & ";"
'         utdata = utdata & Xvar(i) & ";"
'    Next i
'    Print #1, Left(utdata, Len(utdata) - 2)
'    utdata = ""
    
    For i = 1 To MyRange.Rows.Count
        For j = 1 To MyRange.Columns.Count
            If IsError(MyRange(i, j)) Then
                txt = -1
            Else
               txt = Int(MyRange(i, j).Value + 0.5) 'Klarar inte decimaltal?
            End If
            'utdata = utdata & chr(34) & txt & chr(34) & ";"
            utdata = utdata & txt & ";"
        Next j
        Print #1, Left(utdata, Len(utdata) - 2)
        utdata = ""
    Next i

    Close #1

End Sub

Sub kollaok()
  'Static starttid
  Dim res As Boolean
  
  res = FileExists("d:\tem\delme.txt")
  If res = True Then
    'starttid = Now
     Call mstatus("Jobbet är klart")
     Exit Sub
  Else
     Application.StatusBar = timerflag
  End If
  
  'Worksheets("Kontroll").Cells(3, 1) = res
  Call mstatus("R in progress..., please be patient.")
  
  Do While timerflag < 5000
      res = FileExists("d:\tem\delme.txt")
      timerflag = timerflag + 1
      If res = False And timerflag = 4999 Then
            mstatus ("Inte klart men avbryter")
            Exit Sub
      End If
      If res = True Then
          Call mstatus("R finished at")
          Worksheets("Data_till_Start").Range("D30") = FileLen("c:\temp\" & res)
          Worksheets("Data_till_Start").Range("D31") = "c:\temp\" & res
          'Call mstatus("Hämta data")
          Application.StatusBar = ""
      End If
      'Om Kalkylbladlista("Kontroll").Kryssrutor(1).Värde = 1 Så Anropa hämtalog
      'If Worksheets("Kontroll").Cells(2, 2) = 0 Then      ' körning gick bra - rita graf
      'Sheets("DATA").Select
 Loop
 
  'Worksheets("Kontroll").Cells(3, 3) = timerflag
  'Cells(1, 3) = timerflag
   Worksheets("Data_till_Start").Range("D28") = Now '- startid
   Worksheets("Data_till_Start").Range("D28").NumberFormat = "hh:mm:ss"
   Worksheets("Data_till_Start").Range("D29") = Now
   Worksheets("Data_till_Start").Range("D29").NumberFormat = "YY-MM-DD"
 
End Sub


Sub rbatch()
    Dim infil As String
    Dim utfil As String
    Dim txt As String   'För att läsa indata
    
    Dim tomma_rader As Integer
    tomma_rader = 0
    
    Static starttid
    Dim skapa_csv As Byte
    skapa_csv = Worksheets("data_till_start").Range("a27") 'Fast adress
    If skapa_csv > 0 Then Call mstatus("CSV " & skapa_csv & " skapas")
    
    
    Close #1 'För säkehetsskull
    
    starttid = Now
    'Open infil For Output As #1
    Worksheets("data_till_start").Range("d28") = starttid
    Worksheets("data_till_start").Range("D28").NumberFormat = "hh:mm:ss"
    
    Dim rad As Integer 'Startrad för att läsa in R-kod
    Dim kolumn As Integer
    
    txt = ""
    'Dim Xvar As Variant 'Obs fast antal
    'Xvar = Array("Age", "Wage", "w_", "ip", "pp", "atp", "gp", "tjp", "gross", "net", "disp", "cont")
    
    Dim indata As String
''    infil = "df<-data.frame(" & Xvar(1) & "=c("
''    For kolumn = 1 To 12
''        For rad = 2 To 22
''        'If kolumn = 2 And rad = 12 Then Stop
''            indata = Worksheets("data_till_start").Cells(rad, kolumn) 'Haltar om det är formler
''            If indata = "" Then indata = "NA"
''            txt = txt & indata & ","
''
''            If rad = 22 Then
''                txt = txt & "NA)"
''                infil = infil & txt
''                txt = ""
''                If kolumn < UBound(Xvar) Then
''                    infil = infil & "," & chr(13)
''                    infil = infil & Xvar(kolumn + 1) & "=c("
''                Else
''                End If
''                Debug.Print infil
''            End If
''        Next rad
''    Next kolumn
''
''    infil = infil & ")" & chr(13)
''    infil = "df<-df[1:21,]" & chr(13)

    '--- skapar csv fil för att avända till R ---
    If skapa_csv > 0 Then
        Call tabort("delme.txt")
        Call tocsv
    End If
    
    'Skapa R-koden
    If skapa_csv > 1 Then
        txt = ""
        rad = 30 'Startrad för att läsa in R-kod
        Open "c:\temp\delme.R" For Output As #2
        Do Until tomma_rader = 5
            txt = Worksheets("data_till_start").Cells(rad, 1)
            If txt = "" Then tomma_rader = tomma_rader + 1
            If txt <> "" Then
                tomma_rader = 0
                txt = txt '& chr(10) '"10-Carriage return"
            End If
            'infil = infil & txt & chr(10)
            Print #2, txt
            rad = rad + 1
            'If rad < 33 Then Debug.Print txt 'chr(10); chr(13
        Loop
    
        'Open "c:\temp\delme.R" For Output As #2
        'Print #2, infil
        Close #2
        'Call mrakn(0)
    End If
    
    'Call tabort("delme.txt")
    
    txt = "C:\Program Files\R-3.5\bin\R.exe"
    
    If Dir(txt) = "R.exe" Then
        Call mstatus("Startar R i batchmode")
        txt = Shell("C:\Program Files\R-3.5\bin\R CMD BATCH -q c:\temp\delme.R c:\temp\delme.txt", 6)
        'Constant           Value Description
        'vbHide             0     The window is hidden, and focus is passed to the hidden window.
        'vbNormalFocus      1     The window has focus and appears in its most recent size and position.
        'vbMinimizedFocus   2     The window is minimized but has focus.
        'vbMaximizedFocus   3     The window is maximized with focus.
        'vbNormalNoFocus    4     The window appears in its most recent size and position, and the currently active program retains focus.
        'vbMinimizedNoFocus 6     The window is minimized, the currently active program retains focus.
    End If
    
    Call mstatus("Koll om jobbet är klart")
   Call kollaok 'Kollar ...."
    
    'Läsa in resultatet från R
    txt = Worksheets("data_till_start").Range("i27")
    If txt > 0 Then
        'txt = "H31"
        Call rensa_yta("H31") 'Fast adress
        If txt = 2 Then Call logg("delme.txt")
        If txt = 3 Then
            txt = Shell("Notepad c:\temp\delme.txt", 4)
        End If
    
    End If
    timerflag = 0
    
    mstatus ("R-kod klar Uppdatera figuren?")
    Application.StatusBar = Now()
    
End Sub

Sub kopiera(fil1, fil2)
    fil1 = "c:\temp\" & fil1
    fil2 = "c:\temp\" & fil2
    mstatus (fil1 & " kopieras till " & fil2)
    'Note A run-time error will occur if you try to copy a file that is currently open.
    Dim txt As String
    FileCopy fil1, fil2
    'MsgBox "Successfully file Copied.", vbInformation
End Sub

Sub rensa_yta(cell)
    'Obs rensar angränsande celler med värden
    Dim c As Range
    Set c = Application.Range(cell)
        Set c = c.CurrentRegion
        c.Select 'Show the region
        c.Clear
End Sub
Sub q_R_run()
    Dim txt As Variant
    'Call kopiera("delme2.R", "delme.R")
    'txt = shell("C:\Program Files\R-3.5\bin\R CMD BATCH -q  c:\temp\delme.R c:\temp\delme.txt", 2)
    'Call rensa_yta("H31")
    'Call logg("delme.txt")
    txt = Shell("Notepad c:\temp\delme.txt", 4)
End Sub


Sub logg(loggfil)
    Application.StatusBar = "Init"
    If loggfil = "" Then
        loggfil = Worksheets("data_till_start").Range("d31") 'fast adress
        Worksheets("data_till_start").Range("d32") = loggfil
    Else
        Worksheets("data_till_start").Range("d32") = ""
    End If
    
    loggfil = "c:\temp\" + loggfil
    'if dir(loggfil) do not exist
    Dim fsize As Long
    fsize = Len(Dir(loggfil))
    
    If (fsize > 0) Then
        mstatus ("Rensar tidigare utdata")
    Else
        mstatus ("Utfilen" & loggfil & " finns inte")
        Exit Sub
    End If

    mstatus ("Hämtar och skriver " & loggfil)
    
    Close #1    ' Stänger filen för säkerhetsskull.
    
    Dim fdt As Date
    fdt = FileDateTime(loggfil)
    Worksheets("data_till_start").Range("h28") = fdt
    Worksheets("data_till_start").Range("h29") = fdt
    Worksheets("data_till_start").Range("h28").NumberFormat = "hh:mm"
    Worksheets("data_till_start").Range("h29").NumberFormat = "YY-MM-DD"
    
    Open loggfil For Binary As #1    ' Öppnar filen.
    
    Dim flagga As Byte
    Dim indata As String
    Dim i, r, k As Integer
    flagga = 0
    indata = ""
    r = 31
    k = 8 'Kolumn H
    
    Dim tkna As String
    Dim pct As Long
    pct = 0
    
    For i = 1 To FileLen(loggfil)
        tkna = Input(1, #1)
        pct = pct + 1
        Worksheets("data_till_start").Cells(r, k) = indata 'Range("g31")
        'If tkna = "," Then tkna = "."
        If tkna = chr(142) Then tkna = "Ä"
        If tkna = chr(143) Then tkna = "Å"
        If tkna = chr(153) Then tkna = "Ö"
        If tkna = chr(10) Or tkna = chr(13) Then flagga = 1
        If flagga = 0 Then
             indata = indata + tkna 'Fyll på strängen
        ElseIf flagga = 1 Then
           Worksheets("data_till_start").Cells(r, k) = indata
            k = k + 1                 ' Byt kolumn om chr(13)
          If tkna = chr(10) Then    ' Byt rad och gå till första kolumnen
                r = r + 1
                k = 8 'kolumn H
          End If
          indata = ""
          flagga = 0
          tkna = ""
        End If
        pct = 0
        Application.StatusBar = pct
        
  Next
  Close #1    ' Stänger filen.
  
  Application.StatusBar = ""
  
End Sub

''Sub slas(fil)
''    'Dim fil As String
''    fil = "c:\temp\' & fil"
''    Close #1
''
''End Sub


Sub civ()
    'Method 1
    Debug.Print String(65535, vbCr)
''
''    'Method 2
''    Application.VBE.Windows("Immediate").SetFocus
''    Application.SendKeys "^a {DEL} {HOME}"
End Sub

'------------------------------------------
Sub Tommy()
    'Givet vissa förutsättningar,  itererar ev den lön som krävs för att nå en viss pensionsnivå
    'Cell Y3 saknas eller noll - vanlig körning anars sökning
    '<999 Den lön som krävs för att grundskyddet inte ska fås eller bidra till högre marginaleffekt
    '>999 den lön som krävs för att pensionen efter skatt och bidrag  ska nå gränsvärdet
    'Var försiktig
    
    Dim wsMod As Worksheet 'Startfliken
    Dim wsDta As Worksheet 'Utdata
    Dim rngXdta As Range    'och dess vektor
    Dim rngRow As Range     'Aktuel rad i input till utdata
    Dim lngFirstRow As Long
    Dim rngUntilRow As Range
    Dim blnRowConditionEnded As Boolean
    
    blnRowConditionEnded = False
    
    Set wsMod = ThisWorkbook.Worksheets("Start")
    'Set wsMod = ThisWorkbook.Worksheets("Utdata")
    Set wsDta = ThisWorkbook.Worksheets("Mikrosim")
    
    'Define X-data range (input)
    Set rngXdta = wsDta.Range("rngXTopleft") 'För tillf. rad 7 och kolumn 2
    Set rngXdta = rngXdta.CurrentRegion
    
    lngFirstRow = Application.Range("rngExecuteFromRow") - rngXdta.Cells(1, 1).Row
    If lngFirstRow < 1 Then lngFirstRow = 1
    
    Set rngXdta = rngXdta.Offset(lngFirstRow).Resize(rngXdta.Rows.Count - lngFirstRow)
    Set rngUntilRow = Application.Range("rngExecuteUntilRow")
    
    Application.Calculation = xlCalculationManual
    
    'cases = 1 'Cases är publikt och syboliserar antal typfall, kvar som rest
  
    Dim Kol As Integer
    Kol = 12 ' rngXdta.Columns.Count + 2 '10 kolumner input och output (Y) skrivas till den 12:e
    
    Dim wage, wdelta As Double
    Dim rest, sok As Double
    Dim Lrest As Double
    Dim Niter, maxiter As Integer
    
    wdelta = 0 'Inkomständring
    rest = 9999 'Initial rest för att starta
    Lrest = -9999 'laggad rest
    Niter = 0 '#antal iterationer- se endan för maximalt antal
    maxiter = 40
    sok = wsIndata.Range("Y3") 'Sökt värde
    ''If sok > 999 Then maxiter = 30
    'If sok < 999 Then sok inget bidrag
    Dim Ylag, Xlag As Double 'För att mer nyttja utväxlingen...
    
    Dim test As Integer
    test = 0
    'Call Unprotect_sheet(wsStart.CodeName) 'För att kunna skjuta in värden
    wsIndata.Activate 'Återgå till indata
    
    For Each rngRow In rngXdta.Rows
        With rngRow
            If IsMissing(sok) Or sok = 0 Then
                 wage = .Cells(1, 4) / 12  '4 lön månadsbelopp W0 Skulle kunna för rad >1 ta förra personens lön ...
            ElseIf sok > 999 Then
                If .Cells.Row = Application.Range("rngExecuteFromRow") Then
                    wage = .Cells(1, 4) / 12
                Else
                    wage = Xlag * 1.04 'Ett år senare
                End If
            Else
                If .Cells.Row = Application.Range("rngExecuteFromRow") Then
                    wage = 25000
                Else
                    wage = Xlag * 1.04
                End If
            End If
            
            Dim tab2, tab2b As Integer
            tab2 = wsIndata.Range("Y5") 'om specifikt ålder kassamässigt
            tab2b = 79 + tab2 - Application.Range("rng_tabell2_startAge") 'rad att läsa in och ut
            'If (.Cells(1, 3) > tab2 And tab2 > 60) Then .Cells(1, 3) = tab2
            
            'Feed model sheet with x-data, loop through each row
            Application.Range("BornYear") = .Cells(1, 1)      '1 Född
            Application.Range("ParYear") = .Cells(1, 3)       '2 Pensionsålder
            Application.Range("WstartYear") = .Cells(1, 2)    '3 Börjar arbeta
            'Application.Range("Wage_Monthly") = wage
            Application.Range("rng_Yearly_Inflation") = .Cells(1, 5)  '5 Inflation
            Application.Range("rng_Real_Growth") = .Cells(1, 6)       '6 Tillväxt
            Application.Range("rng_FondAvkastning") = .Cells(1, 7)    '7 Avkastning
            Application.Range("rng_TJP_Listbox") = .Cells(1, 10)      '10 Avtalsområde
   
    
            Do While Niter < (maxiter + 1)
                    
                    Application.Range("Wage_Monthly") = wage
                    'wsMod.Calculate
                    Application.Calculate 'Force calculation
                    DoEvents
                    Call Mcalc
                    DoEvents
                    
                    'Get output Y-data - Values
                    If tab2 > 60 And sok > 999 Then
                        rest = sok - wsMod.Range("K" & 79 + tab2 - Application.Range("rng_tabell2_startAge"))
                        'förutätter att tabell2 startar på rad 79
                    ElseIf sok > 999 Then
                        rest = sok - wsMod.Range("D44") 'Disp d44 i start, 054 i utdata 66 år
                    End If
                    
                    If sok > 0 And sok < 999 Then
                        'OM ÄFS slopa garp, dvs 0*garp
                        rest = -(1 * wsMod.Range("D31") + wsMod.Range("D43")) '-Bidrag d31+d43, h54+ n54
                        'If wage > 42625 Then rest = -wsMod.Range("D43")   'D43, N54
                    End If
                    
                    If IsMissing(sok) Or sok = 0 Then rest = 0
                    'If (wage > 100000 And wsMod.Range("N54") = 0) Then Niter = 24 'D43 'dvs BTP/ÄFS
                    rest = Round(rest, 0)
                    
''''                    'If sok < 999 Then
''''                        If Lrest = rest And test < 5 Then
''''                            test = test + 1
''''                            Lrest = rest - 2
''''
''''                           'Lrest = -9999 'Dvs gå ur men backa tillbaka
''''                            If wage < 45865 Then
''''                                If test = 3 And Abs(rest) < 30 Then Lrest = rest 'Gå ur
''''                                If sok < 999 And test = 2 Then
''''                                    wage = Xlag
''''                                    Application.Range("Wage_Monthly") = wage
''''                                    'wsMod.Calculate
''''                                    Application.Calculate 'Force calculation
''''                                    DoEvents
''''                                    Call Mcalc
''''                                    DoEvents
''''                                End If
''''                            End If
''''                        End If
''''                   'End If
                    
                    Application.StatusBar = "Calculates row " & rngRow.Row - lngFirstRow + 1 & " (" & Niter & ")"
                                    '& " (last row is: " & rngXdta.Rows.Count & ")"
        
                    If (Abs(rest) < 10) Or Niter = maxiter Then  'Or (rest = Lrest)
                        test = 0
                        If tab2 < 61 Then
                            .Cells(1, Kol + 0).Value = wsMod.Range("D24")   'Slutlön d24 D52
                            .Cells(1, Kol + 1).Value = wsMod.Range("D36")   'Sa pension d36 L54
                            .Cells(1, Kol + 2).Value = wsMod.Range("D28")   'IP d28 E54
                            .Cells(1, Kol + 3).Value = wsMod.Range("D29")   'ATP d29 F54
                            .Cells(1, Kol + 4).Value = wsMod.Range("D30")   'PP d30 G54
                            .Cells(1, Kol + 5).Value = wsMod.Range("D31")   'Garp d31 H54
                            .Cells(1, Kol + 6).Value = wsMod.Range("D32")   'IPT d32 I54
                            
                            .Cells(1, Kol + 7).Value = wsMod.Range("D34")   'TJP d34 J54
                            .Cells(1, Kol + 8).Value = wsMod.Range("D35")   'IPS d35 K54
                            .Cells(1, Kol + 9).Value = wsMod.Range("D42")   'Netto d42 N54
                            .Cells(1, Kol + 10).Value = wsMod.Range("D43")   'Bidrag d43 N54
                            .Cells(1, Kol + 11).Value = wsMod.Range("D45")  'Disp d45 O54
                        Else
                            .Cells(1, Kol + 0).Value = wsMod.Range("D24")   'Slutlön d24 D52
                            .Cells(1, Kol + 1).Value = wsMod.Range("D" & tab2b)  'Sa pension d36 L54
                            .Cells(1, Kol + 2).Value = wsMod.Range("E" & tab2b)   'IP d28 E54
                            .Cells(1, Kol + 3).Value = 0 'wsMod.Range("D29")   'ATP d29 F54
                            .Cells(1, Kol + 4).Value = wsMod.Range("F" & tab2b)   'PP d30 G54
                            .Cells(1, Kol + 5).Value = wsMod.Range("G" & tab2b)   'Garp d31 H54
                            .Cells(1, Kol + 6).Value = 0 ' wsMod.Range("D" & tab2b)   'IPT d32 I54
                            
                            .Cells(1, Kol + 7).Value = wsMod.Range("F" & tab2b)   'TJP d34 J54
                            .Cells(1, Kol + 8).Value = 0 'wsMod.Range("D" & tab2b)   'IPS d35 K54
                            .Cells(1, Kol + 9).Value = wsMod.Range("I" & tab2b)   'Netto d42 N54
                            .Cells(1, Kol + 10).Value = wsMod.Range("J" & tab2b)   'Bidrag d43 N54
                            .Cells(1, Kol + 11).Value = wsMod.Range("L" & tab2b)  'Disp d45 O54
                        End If
                        If sok > 0 Then .Cells(1, Kol + 12).Value = Application.Range("Wage_Monthly")
                        
''                        If Abs(wage - Xlag) > 0 And Niter > 0 Then
''                           ''.Cells(1, kol + 13).Value = 1 - (wsMod.Range("D44") - Ylag) / (12 * (wage - Xlag))
''                            Debug.Print .Cells.Row; Niter; rest; wage; Xlag; wsMod.Range("D44"); Ylag; _
''                            1 - (wsMod.Range("D44") - Ylag) / 12 / (wage - Xlag)
''                        Else
''                            If sok > 999 Then
''                             .Cells(1, kol + 13).Value = -Niter
''                             Debug.Print .Cells.Row; Niter; wage; Xlag; wsMod.Range("D44"); Ylag; rest; Lrest
''                            Else
''                             Debug.Print .Cells.Row; Niter; wage; wdelta; wsMod.Range("D31"); wsMod.Range("D43"); _
''                             Ylag; Xlag; rest; Lrest; test
''                            End If
''                        End If
                        
                        Niter = 0
                        Exit Do
                    End If
                    
                    
                    DoEvents
                    Niter = Niter + 1
                    'Här skulle man i framtiden kunna ta fram en bättre algoritm som nyttjar första differensen och multiplikativt....

                    If sok > 999 Then
                        If tab2 < 61 Then
                            Debug.Print .Cells.Row; Niter - 1; Round(wage, 0); Round(wdelta / 12, 0); wsMod.Range("D44"); rest; Lrest; test
                        Else
                            Debug.Print .Cells.Row; Niter - 1; Round(wage, 0); Round(wdelta / 12, 0); wsMod.Range("K" & tab2b); Round(rest); Lrest; test
                        End If
                    Else
                        Debug.Print .Cells.Row; Niter; wage; Round(wdelta / 12, 0); wsMod.Range("D31"); wsMod.Range("D43"); rest; Lrest; test
                    End If

                    If sok > 999 Then 'Specifik inkomst
                        If Niter < 2 Then
                            wdelta = rest * 1.3
                            Xlag = 0: Ylag = 0
                        Else
                            wdelta = rest * 1.3

                            If Abs(wage - Xlag) > 10 Then
                                If tab2 < 61 Then
                                    wdelta = ((wsMod.Range("d44") - Ylag) / 12) / (wage - Xlag) 'utväxling
                                Else
                                    wdelta = ((wsMod.Range("K" & tab2b) - Ylag) / 12) / (wage - Xlag) 'utväxling
                                End If

                                ''If Abs(wdelta) > 1 Then wdelta = 0.7 * rest / wdelta

                                If Abs(wdelta) = 0 Then
                                    wdelta = rest * 15
                                ElseIf Abs(wdelta) < 0.3 Then 'Låg utväxling
                                    wdelta = 1.3 * rest / wdelta
                                Else
                                    wdelta = 0.5 * rest
                                End If
                                ''If Abs(wdelta) * 1.1 > Abs(rest) Then wdelta = rest * 1.1
                                ''If Niter > 12 Then wdelta = rest * 0.8
                            End If

                           'Överskott Rest negativ: Rest= Sök - Y
                            If Abs(rest) < 1200 Then
                                    wdelta = rest * 0.4
                                    If Abs(wdelta) < 12 Then wdelta = Sgn(rest) * 12
                                If rest > -100 And Niter > 25 Then Niter = maxiter
                            End If
                        End If
                    Else 'Här söks när bidraget precis upphör
                        wdelta = -rest * 1.4
                        If Abs(rest) > 100 Then wdelta = -rest * 1.7
                        If wage > 42625 Then
                            wdelta = -rest * 0.4
                        End If
                        If wsMod.Range("D31") = 0 Then
                            wdelta = -rest * 1.4
                        End If
                    End If
                    
                    Lrest = rest
                    
                    Xlag = wage
                    wage = wage + wdelta / 12
                    If wage < 0 Then
                        wage = 0
                        Niter = maxiter
                    End If
                    wage = Round(wage, 0)
                    Ylag = wsMod.Range("D44")
                    If tab2b > 61 Then Ylag = wsMod.Range("K" & tab2b)
                    
                    
                                       
                Loop
            
                If rngUntilRow.Value <= .Cells.Row And rngUntilRow <> 0 Then
                        blnRowConditionEnded = True
                        Exit For
                End If
        End With
    Next
    
'ResetExcelEnvironment:

    Application.Calculation = xlCalculationAutomatic 'reset calculation
    Application.StatusBar = ""
    Worksheets("Mikrosim").Select
''    If blnRowConditionEnded Then
''        MsgBox "Körning avslutad på rad " & rngRow.Cells.Row, vbInformation
''    Else
''        MsgBox "Körning klar! " & rngXdta.Rows.Count & " rader beräknade."
''    End If
    'Nollställer
    'Run_from_indata = 0
    'cases = 1
End Sub

Function respekt(ByVal belopp, ByVal hyra, ByVal TJP, ByVal garp, ByVal IPT) As Double
   
    Dim brutto, ctxfvi As Double
    Dim Grundavdrag, GA As Double
    Dim cbefvi As Double
    Dim skatt1, skatt2 As Double
    Dim bidrag1, bidrag2 As Double
    Dim civ As Byte
    Dim kvoten As Byte
    Dim Xage As Integer
    
    Xage = Application.Range("Rng_riktage") + 2 'Application.Range("rng_alt_last_pratt")
    If Application.Range("Rng_resp_max") = 1 Then Xage = Int(PAR) 'Kan påverka Garp och skatt
    civ = 0: kvoten = 1 'antas vara single
    Dim yy As Integer
    
    If Application.Range("RulesFromUtg") = 0 Or IsMissing(Application.Range("RulesFromUtg")) Then
       yy = Int(Application.Range("born") + Application.Range("par")) + Application.Range("rng_alt_last_pratt")
    Else
       yy = Application.Range("RulesFromUtg")
    End If
       
    Dim rad, Kol As Integer
    rad = Application.Range("rng_resp_utdata").Row + 2
    Kol = Application.Range("rng_resp_utdata").Column + 1
       
       
    brutto = gp(0, civ, Int(born), pbb(Int(PAR)), 40, marginal, Xage, yy, _
            Iyear, IBB(Int(PAR)), kvoten)
    
    wsDataTillStart.Cells(rad, Kol) = brutto
    wsDataTillStart.Cells(rad, Kol + 1) = belopp
     
    'Efter skatt och först belopp inkomst ska marginal*0?
    If marginal = 0 Then
            ctxfvi = Int(brutto / 100) * 100
    End If
        
    Grundavdrag = avdragxx(ctxfvi, pbb(Int(PAR)), marginal, Xage, yy, 2100, IBB(Int(PAR)), kvoten, yy, Xage)
    cbefvi = ctxfvi - Grundavdrag
    
    skatt1 = cbefvi * (Kom_skatt(Int(PAR)) + Begravavg(Int(PAR))) _
    + statlig(cbefvi, Tax_limit1(Int(PAR)), Tax_limit2(Int(PAR)), marginal) + PublicAvg(cbefvi, 0.01, Int(PAR), marginal, yy) _
    - FAared(cbefvi, yy, marginal, pbb(Int(PAR)), Iyear, IBB(Int(PAR)), kvoten)
    If skatt1 < 0 Then skatt1 = 0
    If marginal = 0 Then skatt1 = Int(skatt1)
    
    wsDataTillStart.Cells(rad + 1, Kol) = brutto - skatt1
        
    'Resp För typfallet
    If marginal = 0 Then
            ctxfvi = Int(belopp / 100) * 100
    End If
    GA = avdragxx(ctxfvi, pbb(Int(PAR)), marginal, Int(PAR), yy, 2100, IBB(Int(PAR)), kvoten, yy, Xage)

    cbefvi = ctxfvi - GA
    
    skatt2 = cbefvi * (Kom_skatt(Int(PAR)) + Begravavg(Int(PAR))) _
    + statlig(cbefvi, Tax_limit1(Int(PAR)), Tax_limit2(Int(PAR)), marginal) + PublicAvg(cbefvi, 0.01, Int(PAR), marginal, yy) _
    - FAared(cbefvi, yy, marginal, pbb(Int(PAR)), Iyear, IBB(Int(PAR)), kvoten)
    If skatt2 < 0 Then skatt2 = 0
    If marginal = 0 Then skatt2 = Int(skatt2)
    
    wsDataTillStart.Cells(rad + 1, Kol + 1) = belopp - skatt2
    
''        RESP = 1 - RESP / (belopp - skatt1)
   'disponibelt först typisen obs ingen förmgenhet eller kapitalinkomster
    Dim formog, wage, hyram As Double
    formog = 0: wage = 0: hyram = hyra
    
    If Application.Range("Rng_resp_max") = 2 Then hyram = 50000 'månadshyran tillräckligt hög för att klara sig ett tag, BTP ersätter med max
        
     bidrag1 = BTP(brutto, civ * brutto, 12 * hyram, civ, pbb(Int(PAR)), 1, 1 * civ, formog, wage, 0, marginal, 1, year_(Int(PAR)), Iyear, IBB(Int(PAR)), _
                 kvoten, Int(PAR), 0, 0, brutto, brutto * civ, _
                 IBB(Iyear - Int(born)), brutto / 12, 12, 40, Iindex(Int(PAR)), 1, 1)
                 
          '    BTP((brutto(Int(par)) - Wage_(Int(par))) * 12 / pmonth - kapskatt - ptillagg(Int(par)), MakaInk * 12 / pmonth, 12 * hyra_t, civ, pbb(Int(par)), _
'                        ap, apm, formog, 0 * Wage_(Int(par)), 0, marginal, maxhyra, year_(Int(par)), Iyear, IBB(Int(par)), _

'                        kvoten, Int(par), TJP(Int(par)) * 12 / pmonth, tjpm * 12 / pmonth, garp(Int(par)) * 12 / pmonth, garp(Int(par)) * 12 / pmonth, _
'                        IBB(Iyear - Int(born)), mpension / 12, pmonth, ftid, Iindex(Int(par)), Iindex(2022 - Int(born)), uttagIP, ansoker)
     bidrag2 = SBTP(belopp, hyra, civ, bidrag1, Grundavdrag, Kom_skatt(Int(PAR)), 1, formog, _
                 pbb(Int(PAR)), 1, year_(Int(PAR)), Iyear, IBB(Int(PAR)), kvoten, Int(PAR), 0, 0, marginal, -99, 40)
     bidrag1 = btp_sbtp(bidrag1, bidrag2, marginal, year_(Int(PAR)))
     
     wsDataTillStart.Cells(rad + 2, Kol) = brutto - skatt1 + bidrag1
     
     If Application.Range("Rng_Ansokt") > 0 Then

         bidrag1 = BTP(belopp - IPT, civ * belopp, 12 * hyra, civ, pbb(Int(PAR)), 1, 1 * civ, formog, wage, 0, marginal, 1, year_(Int(PAR)), Iyear, IBB(Int(PAR)), _
                     kvoten, Int(PAR), TJP, TJP * civ, garp, garp * civ, _
                     IBB(Iyear - Int(born)), ctxfvi / 12, 12, 40, Iindex(Int(PAR)), 1, 1)
    
         bidrag2 = SBTP(belopp - IPT, hyra, civ, bidrag1, GA, Kom_skatt(Int(PAR)), 1, formog, _
                     pbb(Int(PAR)), 1, year_(Int(PAR)), Iyear, IBB(Int(PAR)), kvoten, Int(PAR), 0, 0, marginal, -99, 40)
         bidrag2 = btp_sbtp(bidrag1, bidrag2, marginal, year_(Int(PAR)))
    Else
        bidrag2 = 0
        'wsDataTillStart.Cells(rad + 2, Kol + 1) = belopp - skatt2 + bidrag2
         
    End If
    wsDataTillStart.Cells(rad + 2, Kol + 1) = belopp - skatt2 + bidrag2
''    RESP = 1 - (maxi(ctxfvi - skatt2, 0) + bidrag1) / belopp

End Function

Function fnTest_Folder_Exist_With_Dir() As Boolean
    'Return True if folder exist, else False
    Dim sFolderPath As String
    Dim sRndFileName As String
    Dim fileNbr As Long

    sFolderPath = Trim(WsPensioner.Range("PathOutput").Value)
    If Right(sFolderPath, 1) <> "\" Then
        sFolderPath = sFolderPath & "\"
    End If
    
    If Dir(sFolderPath, vbDirectory) <> vbNullString Then
        'check if write authority exist
        sRndFileName = Int(Rnd(1758) * 100000000000#) & ".tst"
        fileNbr = FreeFile
        On Error Resume Next
        Open sFolderPath & sRndFileName For Output As #fileNbr
        If err.Number > 0 Then
            fnTest_Folder_Exist_With_Dir = False
            MsgBox "Du har inte skrivbehörighet till sökvägen: " & sFolderPath & ".", vbCritical, "Kontroll filsökväg"
        Else
            fnTest_Folder_Exist_With_Dir = True
        End If
        Close fileNbr
        On Error GoTo 0
    Else
        fnTest_Folder_Exist_With_Dir = False
        MsgBox "Angiven sökväg (" & sFolderPath & ") finns inte. Körning avbryts!", vbCritical, "Kontroll filsökväg"
    End If
End Function

Public Sub cmd_BrowseForFolder_Click()

    On Error GoTo errTag
    Dim fileExplorer As FileDialog
    Set fileExplorer = Application.FileDialog(msoFileDialogFolderPicker)
    Dim folderPath As String

    'To allow or disable to multi select
    fileExplorer.AllowMultiSelect = False
    fileExplorer.InitialFileName = WsPensioner.Range("G18")
    If Right(fileExplorer.InitialFileName, 1) <> "\" Then fileExplorer.InitialFileName = fileExplorer.InitialFileName & "\" 'Add backslash if needed
    Application.Calculation = xlCalculationManual
    With fileExplorer
        If .Show = -1 Then 'Folder is selected
            folderPath = .SelectedItems.Item(1)
            WsPensioner.Range("G18") = folderPath
        Else ' else dialog is cancelled
            MsgBox "Ingen mapp valdes.", vbInformation
        End If
    End With
    Application.Calculation = xlCalculationAutomatic
errTag:

End Sub
Public Sub cmd_OpenFolderInExplorer()
    Dim sFolderPath As String
    If fnTest_Folder_Exist_With_Dir Then
        sFolderPath = Trim(WsPensioner.Range("PathOutput").Value)
'        If Right(sFolderPath, 1) <> "\" Then
'            sFolderPath = sFolderPath & "\"
'        End If
        Call Shell("explorer.exe """ & sFolderPath & "", vbNormalFocus)
    End If
End Sub
'Sub MakeMyFolder()
''Updateby Extendoffice
'    Dim fdObj As Object
'    Application.ScreenUpdating = False
'
'    Dim sFolderPath As String
'    sFolderPath = Range("c3").Value
'
'    Set fdObj = CreateObject("Scripting.FileSystemObject")
'    If fdObj.FolderExists(sFolderPath) Then
'        MsgBox "Found it.", vbInformation, "Kutools for Excel"
'    Else
'        fdObj.CreateFolder (sFolderPath)
'        MsgBox "It has been created.", vbInformation, "Kutools for Excel"
'    End If
'    Application.ScreenUpdating = True
'End Sub
