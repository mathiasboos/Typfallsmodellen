Attribute VB_Name = "aaaErrHlpTmp"
'Function IP_(ByVal year As Long, ByVal par As Double, ByVal born As Double, ByVal pratt As Double, ByVal arvsf1 As Double, _
'            ByVal arvsf2 As Double, ByVal kostf As Double, ByVal pbh_ing As Double, ByVal andel As Double, _
'            ByVal pens As Double, ByVal deltal As Double, ByVal index As Double, _
'            Optional ByVal Def_ar = 999, Optional marginal = 0, Optional typ = 0, _
'            Optional Bindex = 0) As Double
'    subStoreParametersTemp
    'Application.Volatile
    Rem Beräkning av underlaget för arvsvinster, indexeringen förvaltingskostnaden samt underlag för pensioneringen eller _
        utgående pensionsbehållningen 31/12 och inkomstpensionen beroende på val av typ.
       
    'Year    - Inkomstår
    'Par     - pensionsålder slutgiltiga uttaget, ÅÅ,åå
    'Born    - födelseår, ÅÅÅÅ,åå
    'Pratt   - Pensionsrätt
    'arvsf1  - Arvsvinstfaktor <=60 'Numera riktl-1
    'arvsf2  - Arvsvinstfaktor >=60 'Dito
    'Kostf   - förvaltningskostnaden
    'PBH_ing - Ingående Pensionsbehållningen, 31/12 t-1 samt justering för återtagande och eller överförda rätter
    'Andel   - uttagsdel vid pensioneringen
    'Pens    - Inkomstpensionen året innan
    'Deltal  - Aktuellt delningstal
    'Index     - förräntningen eller Indexeringen (inkomst/balansindex)
    'def_ar  - Definitiv pensionsavgång (100%)
    'Marginal- Avrundning och Oavsett kohort
    'Typ     - 0-4; Vad som ska återges: 0-Pension, 1-Arv, 2=Index, 3-Kostnad, 4=PBH
    'Bindex  - Kvoten Balans/Inkomstindex infört 20151118
    
    'par = Worksheets("Start").Range("Par").Value
    'def_ar = Worksheets("Start").Range("Def_ar").Value

Sub ErrHelpLogger(ByVal CodeLine As Integer, ByVal FnValue As Double, ByVal age As Long, ByVal year As Long, ByVal PAR As Double, ByVal born As Double, ByVal pratt As Double, ByVal arvsf1 As Double, _
            ByVal arvsf2 As Double, ByVal kostf As Double, ByVal pbh_ing As Double, ByVal andel As Double, _
            ByVal pens As Double, ByVal deltal As Double, ByVal index As Double, _
            Optional ByVal def_ar = 999, Optional marginal = 0, Optional Typ = 0, _
            Optional Bindex = 0)
'   Debugverktyg som skriver ut samma variabler som används vid anrop av funktionen IP_
'   Skrivs ut i excelfil (mall) vid namn OutputErrHelpLog.xlsx, som ska ligga i samma mapp som Typfallsmodellen.
'   Enda kravet är att den ska ett namngivet område: rngtopleft, som indikerar utdata top-left-position
'   Fyra ytterligare fem variabler skrivs ut, men bara två ytterligare används vid anrop (se första och femte nedan)
'       den första är Codeline, som lämpligtvis anger kodradens Line, används för att veta vilket anrop som görs (IP_ anropas många gånger i Mcalc.
'       den andra är Partiell uttags-ålder
'       den tredje är uttag IP %
'       den fjärde är uttag PP %
'       den femte är det returnerade värdet av funktionen IP_.
'   Hur man kan anropen denna sub: sök på ErrHelpLogger i hela projektet
'
'   Kortsluta denna rutin: se några rader ner: Exit sub
'   För att begränsa output till det man för tillfället är intresserad av, se exempel någon rad ner: If age > 70
'   Output skrivs aldrig över, utan fylls på successivt.
            
            
    Dim arrLbl() As Variant
    Dim a() As Variant
    Dim rngPartielltUttag As Range
    Exit Sub 'Måste kommenteras bort för att aktivera
    'If age < 66 Then Exit Sub 'Används för att begränsa output
    'If age > 70 Then Exit Sub  'Används för att begränsa output
    
    
    arrLbl = Array("Line in code", "PartDefÅr", "PartUttagIP", "PartUttagPP", "IP_value/12", "age", "Year", "PAr", "bORN", "Pensionsrätt", "Arvsvinstfaktor1", "Arvsvinstfaktor2", "FörvKost", "Pbh_Ing age-1", "andel", "Pens-1", "Deltal", "Index", "beräknat Def_ar", "marginal", "typ", "Bindex")
    ReDim a(1 To UBound(arrLbl) + 1)
    
    Set rngPartielltUttag = ThisWorkbook.Names("rng_Val_partiellt").RefersToRange
    
    
    'Kolumnrubriker, om de inte redan finns.
    a(1) = CodeLine
    a(2) = ThisWorkbook.Names("rng_def_ar").RefersToRange.Value
    a(3) = rngPartielltUttag.Cells(Worksheets("Adv_settings").Range("rng_def_ar").Offset(1).Value, 1) 'hämta utifrån val (1-5) rätt rad
    a(4) = rngPartielltUttag.Cells(Worksheets("Adv_settings").Range("rng_def_ar").Offset(2).Value, 1) 'hämta utifrån val (1-5) rätt rad
    a(5) = FnValue / 12
    a(6) = age
    a(7) = year
    a(8) = PAR
    a(9) = born
    a(10) = pratt
    a(11) = arvsf1
    a(12) = arvsf2
    a(13) = kostf
    a(14) = pbh_ing
    a(15) = andel
    a(16) = pens
    a(17) = deltal
    a(18) = index
    a(19) = def_ar
    a(20) = marginal
    a(21) = Typ
    a(22) = Bindex

    'Output
    Dim wb As Workbook
    Dim rng As Range
    Dim i As Long
    
    fnOpenWb wb, "OutputErrHelpLog.xlsx"
    If wb Is Nothing Then Stop 'Se till att att det finns en arbetsbok med namn enligt ovan i samma mapp som typfallsmodellen. Ska finnas namngivet område "rngtopleft" i första fliken.
    If ActiveWorkbook.Name <> ThisWorkbook.Name Then ThisWorkbook.Activate 'Undvik problem när output-filen öppnas
    
    Set rng = wb.Names("rngtopleft").RefersToRange
    If rng.Value = "" Then
        For i = 0 To UBound(arrLbl)
            rng.Offset(, i).Value = arrLbl(i)
        Next
    End If
    Set rng = rng.Offset(rng.CurrentRegion.Rows.Count).Resize(1, rng.CurrentRegion.Columns.Count)
    rng.Value = a
        
    Set rng = Nothing
    Erase a
    Erase arrLbl
    
End Sub
