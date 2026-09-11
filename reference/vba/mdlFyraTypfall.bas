Attribute VB_Name = "mdlFyraTypfall"
Option Explicit

'Givet ett antal typfall, hur har pensionsutvecklingen varit för dessa
Function fnCheckConstants(ByRef rng As Range, ByVal constantVar As Variant)
    'Kontrollerar och återställer om konstanta värde ändrats i fliken.
    If rng.Value <> constantVar Then
        MsgBox "Du kan inte ändra '" & rng.Offset(, -1) & "'. Aktuellt värde återställs. " & vbCrLf & "Om en ändring ska genomföras, ändra konstanter i koden.", vbInformation
        If rng.Address = WsPensioner.Range("iYear").Address Then rng.Formula = "=refar-2" 'Special. Formel för Inkomsterna avser.
         
    End If
End Function
Sub typisar()
    'Först lite data över typisar, år mm
    'Koll om det går att skriva ut csv filerna
    Dim rng As Range
    Const cAntProgAr As Long = 3 'Antal prognosår
    Const cAntFastsAr As Long = 10 'Antal fastställda år
    Const calderRefar As Long = 74 'Antagen ålder vid referensår
    Const cDiffRefarInkar As Long = 2
    Dim refar As Integer 'Referensår - i R senare aar
    Dim Iyear  As Integer 'året för inläst pensionsinkomster
    
    refar = WsPensioner.Range("refar").Value 'läs in referns år mm för indata-kontroll
    Iyear = WsPensioner.Range("iYear") 'De år inkomsterna avser
       
    'Kontroll sökväg
    If Not fnTest_Folder_Exist_With_Dir() Then
        Exit Sub
    End If
    
    'Kontroll antaganden (som inte får ändras i arket)
    fnCheckConstants WsPensioner.Range("AntProgAr"), cAntProgAr
    fnCheckConstants WsPensioner.Range("AntFastsAr"), cAntFastsAr
    fnCheckConstants WsPensioner.Range("alderRefar"), calderRefar
    fnCheckConstants WsPensioner.Range("iYear"), refar - cDiffRefarInkar
    
    Dim visaTypfall As Integer
    visaTypfall = WsPensioner.Range("visaTypfall").Value 'Koll av ett typfall som skriva ut i excel, 0 inget skrivs ut
    Dim libref As String
    'Typisarna skapas i utfil typis.csv och KPI mm i nyckeltal.csv
    'som lagras i biblioteket libref
    libref = fnPathAddBackslashIfNeeded(WsPensioner.Range("PathOutput").Value) ' "c:\tommy\Rapport\X typisar\"
    
    Dim arkull As Integer
    arkull = refar - calderRefar 'I regleringsbrevet avses de som är 74 år gamla vid referensåret
        
    Dim typerAntal As Integer 'Antal typfall som ska beräknas
    'typerAntal = Range(WsPensioner.Range("D9"), WsPensioner.Range("D9").End(xlToRight)).Columns.Count
    typerAntal = Range(WsPensioner.Range("typfallTopLeft"), WsPensioner.Range("typfallTopLeft").End(xlToRight)).Columns.Count
    If visaTypfall > typerAntal Then
'        visaTypfall = 0
'        WsPensioner.Range("visaTypfall") = 0
        MsgBox "Du har angivt ett typfall som inte existerar. Det finns förnärvarande " & typerAntal & " typfall.", vbCritical
        Exit Sub
    End If
    
    Dim startar As Integer
    Dim slutar As Integer
    startar = refar - cAntFastsAr ' 2023 - 10 'Året innan för årsomräkningar
    slutar = refar + cAntProgAr 'Sista året + 3 prognosår
    Dim year  As Integer
       
       
    'Resten skapar typfallens årsinkomster och pensionsutfall
    'När allt är klart, starta R och kör xtypisar.R
    Dim psambo() As Variant 'Option att välja sambo
    Dim ptillp() As Double
    Dim pinkp() As Double
    Dim pprepeg() As Double
    Dim Pavpen() As Double
    Dim Tpriv() As Double
    Dim Uboende() As Double
    'beräknas
    Dim pgarp() As Double
    Dim pipt() As Double
    
    Dim brutto() As Double
    Dim Gavd() As Double
    Dim Besk() As Double
    Dim Skatt() As Double
    Dim Netto() As Double
    Dim Bidrag() As Double
    Dim disp() As Double
        
    Dim omr1() As Double
    Dim omr2() As Double
    Dim bnp() As Double 'BNP per capita
    Dim ii As Integer 'Rad/kolumnhållare
    Dim r As Long ' current column
    Dim c As Long 'current row
    
    marginal = 0 '1-Marginalberäkningar, inga avrundningar
    
    ReDim psambo(1 To typerAntal)
    ReDim ptillp(startar To slutar, 1 To typerAntal)
    ReDim pinkp(startar To slutar, 1 To typerAntal)
    ReDim pprepeg(startar To slutar, 1 To typerAntal)
    ReDim Pavpen(startar To slutar, 1 To typerAntal)
    ReDim Tpriv(startar To slutar, 1 To typerAntal)
    ReDim pgarp(startar To slutar, 1 To typerAntal)
    ReDim pipt(startar To slutar, 1 To typerAntal)
    ReDim brutto(startar To slutar, 1 To typerAntal)
    ReDim Gavd(startar To slutar, 1 To typerAntal)
    ReDim Besk(startar To slutar, 1 To typerAntal)
    ReDim Skatt(startar To slutar, 1 To typerAntal)
    ReDim Netto(startar To slutar, 1 To typerAntal)
    ReDim Bidrag(startar To slutar, 1 To typerAntal)
    ReDim disp(startar To slutar, 1 To typerAntal)
    
    ReDim Uboende(startar To slutar, 1 To typerAntal)
    
    ReDim omr1(startar To slutar) As Double
    ReDim omr2(startar To slutar) As Double
    
    r = WsPensioner.Range("typfallTopLeft").Row + 1 'location typfall top left row
    c = WsPensioner.Range("typfallTopLeft").Column - 1 'location typfall top left column
    
    For ii = 1 To typerAntal
        psambo(ii) = Trim(WsPensioner.Cells(r, c + ii))
        ptillp(Iyear, ii) = WsPensioner.Cells(r + 1, c + ii)
        pinkp(Iyear, ii) = WsPensioner.Cells(r + 2, c + ii)
        pprepeg(Iyear, ii) = WsPensioner.Cells(r + 3, c + ii)
        Pavpen(Iyear, ii) = WsPensioner.Cells(r + 4, c + ii)
        Tpriv(Iyear, ii) = WsPensioner.Cells(r + 5, c + ii)
        Uboende(Iyear, ii) = WsPensioner.Cells(r + 6, c + ii)
    Next ii
    
    'verfiera sambo
    For ii = 1 To typerAntal
        If psambo(ii) <> "" Then
            If Not IsNumeric(psambo(ii)) Then
                MsgBox "Fel inmatning av sambo för typfall " & ii & ". Tillåtna värden är blankt eller nummer för aktuellt sambo-typfall.", vbCritical
                Exit Sub
            Else
                'Siffra
                If psambo(ii) = 0 Then
                    psambo(ii) = "" 'Ignorera
                Else
                    If psambo(ii) > typerAntal Then
                        MsgBox "Fel inmatning av sambo för typfall " & ii & ". Tillåtna värden är blankt eller nummer för aktuellt sambo-typfall.", vbCritical
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next
    
    'Läs in IBB, PBB mm en del är globala se modul VBA
    ReDim year_(startar To slutar) As Long
    ReDim IBB(startar To slutar) As Double
    ReDim pbb(startar To slutar) As Double
    ReDim KPI(startar To slutar) As Double
    ReDim Iindex(startar To slutar) As Double
    ReDim Pindex(startar To slutar) As Double
    ReDim yield(startar To slutar) As Double
    ReDim Kom_skatt(startar To slutar) As Double
    ReDim Tax_limit1(startar To slutar) As Double
    ReDim Tax_limit2(startar To slutar) As Double
    ReDim bnp(startar To slutar) As Double
      
    For year = startar To slutar
        year_(year) = year
        KPI(year) = wsNyckelTal.Cells(year - 1959 + 6, 84)
        IBB(year) = wsNyckelTal.Cells(year - 1959 + 6, 88)
        pbb(year) = wsNyckelTal.Cells(year - 1959 + 6, 86)
        Iindex(year) = wsNyckelTal.Cells(year - 1959 + 6, 90)
        Pindex(year) = wsNyckelTal.Cells(year - 1959 + 6, 93)
        yield(year) = 1 + wsTal.Cells(year - 1959 + 6, 17) 'Fondrörelsen
        Kom_skatt(year) = wsK_Skatt.Cells(year - 1930 + 2, 2) / 100 'Obs efter 1930
        Tax_limit1(year) = wsNyckelTal.Cells(year - 1958 + Application.Range("rng_StartRad_Nyckeltal").Row, Application.Range("rng_Tax_limit1").Column)
        Tax_limit2(year) = wsNyckelTal.Cells(year - 1958 + Application.Range("rng_StartRad_Nyckeltal").Row, Application.Range("rng_Tax_limit2").Column)
        bnp(year) = wsNyckelTal.Cells(year - 1958 + Application.Range("rng_StartRad_Nyckeltal").Row, Application.Range("rng_Tax_limit2").Column + 1) 'Löpande priser
        
    Next year
    
    'Inkomst + tilläggspensionen
    For year = (startar + 1) To slutar
          omr1(year) = (Pindex(year) / Pindex(year - 1)) / 1.016
          bnp(year) = bnp(year)
         'Debug.Print year; Round(omr1(year), 3)
    Next year
    
    omr2(Iyear) = 100 '/ omr1(Iyear)
    year = Iyear - 1 'Året före och stega tillbaka
    Do While year >= startar
      omr2(year) = omr2(year + 1) / omr1(year + 1)
      'Debug.Print year; Round(omr2(year), 3)
      year = year - 1
    Loop
    year = Iyear + 1 'Året efter
    Do While year <= slutar
      omr2(year) = omr2(year - 1) * omr1(year)
      'Debug.Print year; Round(omr2(year), 3)
      year = year + 1
    Loop
    
    For year = startar To slutar
        For ii = 1 To typerAntal
            If year <= Iyear Then
                ptillp(year, ii) = (ptillp(Iyear, ii) * omr2(year) / 100) / 12
                pinkp(year, ii) = (pinkp(Iyear, ii) * omr2(year) / 100) / 12
            Else
                ptillp(year, ii) = (ptillp(Iyear, ii) * omr2(year) / 100)
                pinkp(year, ii) = (pinkp(Iyear, ii) * omr2(year) / 100)
            End If
            ptillp(year, ii) = Round(ptillp(year, ii), 0)
            pinkp(year, ii) = Round(pinkp(year, ii), 0)
            
        Next ii
        'If year > startar Then Debug.Print year, Round(ptillp(year, 2), 0); Round(pinkp(year, 2), 0)
    Next year
    
    'Omräkning av premiepension mht förskottsräntan
    For year = (startar + 1) To slutar
          omr1(year) = yield(year)
          If year < 2017 Then
            omr1(year) = omr1(year) / 1.039
          ElseIf year < 2020 Then
            omr1(year) = omr1(year) / 1.029
          Else
            omr1(year) = omr1(year) / 1.0165
          End If
          'Debug.Print year; Round(omr1(year), 3)
    Next year
          
    Erase yield
    
    omr2(Iyear) = 100 '/ omr1(Iyear)
    year = Iyear - 1
    Do While year > startar
      omr2(year) = omr2(year + 1) / omr1(year)
      'Debug.Print year; Round(omr2(year), 3)
          year = year - 1
    Loop
    year = Iyear + 1
    Do While year <= slutar
      omr2(year) = omr2(year - 1) * omr1(year)
      'Debug.Print year; Round(omr2(year), 3)
      year = year + 1
    Loop
    
    For year = startar To slutar
        For ii = 1 To typerAntal
             If year <= Iyear Then
                pprepeg(year, ii) = (pprepeg(Iyear, ii) * omr2(year) / 100) / 12
            Else
                pprepeg(year, ii) = (pprepeg(Iyear, ii) * omr2(year) / 100)
            End If
            pprepeg(year, ii) = Round(pprepeg(year, ii), 0)
        Next ii
        'If year > startar Then Debug.Print year, Round(pprepeg(year, 2), 0); Round(omr2(year), 2)
    Next year
     
    'Omräkning av tjänstepension och privat sparandde (pbb). Hyran (KPI)
    For year = startar To slutar
          omr1(year) = KPI(year) / KPI(Iyear)
          omr2(year) = pbb(year) / pbb(Iyear)
           For ii = 1 To typerAntal
            If year <= Iyear Then
                Pavpen(year, ii) = Pavpen(Iyear, ii) * omr2(year) / 12
                Tpriv(year, ii) = Tpriv(Iyear, ii) * omr2(year) / 12
                
                Uboende(year, ii) = Uboende(Iyear, ii) * omr1(year) / 12
            Else
                Pavpen(year, ii) = Pavpen(Iyear, ii) * omr2(year)
                Tpriv(year, ii) = Tpriv(Iyear, ii) * omr2(year)
                
                Uboende(year, ii) = Uboende(Iyear, ii) * omr1(year)
            End If
            Pavpen(year, ii) = Round(Pavpen(year, ii), 0)
            Tpriv(year, ii) = Round(Tpriv(year, ii), 0)
            Uboende(year, ii) = Round(Uboende(year, ii), 0)
            
        Next ii
       'If year > startar Then Debug.Print year; Round(Pavpen(year, 2), 0); Round(Uboende(year, 2), 0); Round(omr2(year), 3); Round(omr1(year), 3)
    Next year
          
    'Ingen Omräkning av lön behövs
    
    
    Dim bald As Integer
    Dim rikt As Integer
    rikt = 65
    
    Dim civ As Integer
    For year = startar To slutar
        bald = year - arkull
        
        If year > 2022 Then rikt = 66 'OBS Påverkar även grundavdraget OBS födda 1957, 1959 men inte aktuellt för typisarna ovan
        
        For ii = 1 To typerAntal
            'civ = 0
            'If ii > 4 Then civ = 1
            If psambo(ii) = "" Then
                civ = 0
            Else
                'Sambo
                civ = 1
            End If
       
            pgarp(year, ii) = gp(12 * (pinkp(year, ii) * 185 / 160 + ptillp(year, ii)), civ, Iyear - 72, pbb(year), 40, marginal, bald, year, IBB(year), 1, rikt) / 12
            
            pipt(year, ii) = tillagg(12 * (pinkp(year, ii) * 185 / 160 + ptillp(year, ii)), year, Iindex(year), Iindex(2021), 1, 40) / 12
            If year = 2021 Then pipt(year, ii) = pipt(year, ii) * 4 / 12
                
            
            'Bruttoinkomst
            brutto(year, ii) = pinkp(year, ii) + ptillp(year, ii) + pprepeg(year, ii) + Pavpen(year, ii) + Tpriv(year, ii) + _
                   pgarp(year, ii) + pipt(year, ii)
                   'Kolla rikt som styr det förhöjda grundavdraget för äldre, obs årsinkomst
            'brutto(year, ii) = Round(brutto(year, ii), 0)
            
            Gavd(year, ii) = avdragxx(12 * brutto(year, ii), pbb(year), marginal, bald, year, 2100, IBB(year), 1, 0, rikt)
            
            Besk(year, ii) = 12 * brutto(year, ii) - Gavd(year, ii)
            If marginal = 0 Then Besk(year, ii) = Int(Besk(year, ii) / 100) * 100
            
            'Kommunal skatt
            Skatt(year, ii) = Besk(year, ii) * Kom_skatt(year)
            If marginal = 0 Then Skatt(year, ii) = Int(Skatt(year, ii))
            Skatt(year, ii) = (Skatt(year, ii) + statlig(Besk(year, ii), Tax_limit1(year), Tax_limit2(year), marginal, year) + _
                PublicAvg(Besk(year, ii), 0.01, bald, marginal, year)) / 12
            'Skatt(year, ii) = Round(Skatt(year, ii), 0)
            
            Netto(year, ii) = (brutto(year, ii) - Skatt(year, ii))
            
            'Bidrag
            Bidrag(year, ii) = BTP(12 * brutto(year, ii) - 12 * pipt(year, ii), 0, 12 * Uboende(year, ii), civ, pbb(year), 1, 0, 0, 0, 0, marginal, 0, year, _
                                2100, 0, 1, bald, 12 * Pavpen(year, ii), 0, pgarp(year, ii), 0, 0) / 12
            'Bidrag(year, ii) = Round(Bidrag(year, ii), 0)
            
            'Disp
            disp(year, ii) = Netto(year, ii) + Bidrag(year, ii)
            'Disp(year, ii) = Round(Disp(year, ii), 0)
            
        Next ii
        
    ''    ii = 1
    ''    If year > startar Then Debug.Print year; Round(12 * brutto(year, ii), 0); Gavd(year, 2); Round(Besk(year, ii), 0); Skatt(year, ii); Round(12 * Netto(year, ii), 0); _
    ''    Round(Bidrag(year, ii), 0); Round(12 * Disp(year, ii), 0)
  
Next year

Dim curSambo As Integer
'Samboende paret och bidraget ...
For year = startar To slutar
    For ii = 1 To typerAntal
        If psambo(ii) <> "" Then
            curSambo = psambo(ii)
            civ = 1
            Bidrag(year, ii) = BTP(12 * brutto(year, ii) - 12 * pipt(year, ii), 12 * brutto(year, ii) - 12 * pipt(year, curSambo), Uboende(year, ii), civ, pbb(year), 1, 1, 0, 0, 0, marginal, 0, year, _
                2100, 0, 1, bald, 12 * Pavpen(year, ii), 12 * Pavpen(year, curSambo), pgarp(year, ii), pgarp(year, curSambo), 0)
            'Disp
            disp(year, ii) = Netto(year, ii) + Bidrag(year, ii) 'Summeras?
        End If
    Next ii
  
Next year

'Informationen sparas som CSV fil: YEAR INKOMSTER TYPFALL
'Const Labels As Integer = 13
'Dim rnglbl As Range         'Etiketter för utskrift

Dim Mlabels(1 To 12) As String    'För utskrift i utdata
    Mlabels(1) = "Year"
    Mlabels(2) = "Inkomstpension"
    Mlabels(3) = "Garp"
    Mlabels(4) = "IPT"
    Mlabels(5) = "Premie"
    Mlabels(6) = "TJP"
    Mlabels(7) = "IPS"
    Mlabels(8) = "Brutto"
    Mlabels(9) = "Skatt"
    Mlabels(10) = "Bidrag"
    Mlabels(11) = "Disp"
    Mlabels(12) = "KPI"
    'Mlabels(13) = "BNP_cap"

    With WsPensioner
        'Set rnglbl = Application.Range("B21:m21")
        Set rng = Application.Range("outputTopLeft")
        Set rng = rng.Resize(1, UBound(Mlabels))
        rng.Value = Mlabels
        'Call Format_Range(rnglbl, True, 65535)
        'Set rnglbl = Nothing
    End With

    Dim mvalues() As Variant
    'Dim mfast() As Variant 'Fasta priser sker i R eller direkt i arket

    ReDim mvalues(startar + 1 To slutar, 1 To UBound(Mlabels))
    'ReDim mfast(startar + 1 To slutar, 1 To UBound(Mlabels))
    
    'Dim m_value As Double
 
    'Debug.Print libref & "typis.csv"
    Open libref & "typis.csv" For Output As #3
    Print #3, "year; TP; IP; GP; IPT; PP; TJP; IPS; Brutto; Skatt; Netto; BTP; Disp; Typis"

    For ii = 1 To typerAntal
          'Utskrift till arket ska ändras, för denna lösning är seg se mlabel och mvalues i vba_go
        If visaTypfall = ii Then
            If psambo(ii) = "" Then
                WsPensioner.Range("outputTopLeft").Offset(-1) = "Typfall: " & ii & ", löpande priser"
            Else
                Dim strTmp As String
                If ii < psambo(ii) Then
                    strTmp = ii & " & " & psambo(ii)
                Else
                    strTmp = psambo(ii) & " & " & ii
                End If
                WsPensioner.Range("outputTopLeft").Offset(-1) = "Snitt typfall: " & strTmp & ", löpande priser"
            End If
        End If
     
        For year = (startar + 1) To slutar
            
            Print #3, year & ";" & _
            Replace(ptillp(year, ii), ",", ".") & ";" & _
            Replace(pinkp(year, ii), ",", ".") & ";" & _
            Replace(pgarp(year, ii), ",", ".") & ";" & _
            Replace(pipt(year, ii), ",", ".") & ";" & _
            Replace(pprepeg(year, ii), ",", ".") & ";" & _
            Replace(Pavpen(year, ii), ",", ".") & ";" & _
            Replace(Tpriv(year, ii), ",", ".") & ";" & _
            Replace(brutto(year, ii), ",", ".") & ";" & _
            Replace(Skatt(year, ii), ",", ".") & ";" & _
            Replace(Netto(year, ii), ",", ".") & ";" & _
            Replace(Bidrag(year, ii), ",", ".") & ";" & _
            Replace(disp(year, ii), ",", ".") & ";" & ii

''            Gavd(year, ii)
''            Besk(year, ii)
            
            If visaTypfall = ii Then
                If psambo(ii) <> "" Then
                    'sambo
                    curSambo = psambo(ii)
                    mvalues(year, 1) = year
                    mvalues(year, 2) = (pinkp(year, ii) + ptillp(year, ii) + pinkp(year, curSambo) + ptillp(year, curSambo)) / 2
                    mvalues(year, 3) = (pgarp(year, ii) + pgarp(year, curSambo)) / 2
                    mvalues(year, 4) = (pipt(year, ii) + pipt(year, curSambo)) / 2
                    mvalues(year, 5) = (pprepeg(year, ii) + pprepeg(year, curSambo)) / 2
                    mvalues(year, 6) = (Pavpen(year, ii) + Pavpen(year, curSambo)) / 2
                    mvalues(year, 7) = (Tpriv(year, ii) + Tpriv(year, curSambo)) / 2
                    mvalues(year, 8) = (brutto(year, ii) + brutto(year, curSambo)) / 2
                    mvalues(year, 9) = (Skatt(year, ii) + Skatt(year, curSambo)) / 2
                    mvalues(year, 10) = (Bidrag(year, ii) + Bidrag(year, curSambo)) / 2
                    mvalues(year, 11) = (disp(year, ii) + disp(year, curSambo)) / 2
                    mvalues(year, 12) = KPI(year)
                Else
                    'Inte sambo
                    mvalues(year, 1) = year
                    mvalues(year, 2) = pinkp(year, ii) + ptillp(year, ii)
                    mvalues(year, 3) = pgarp(year, ii)
                    mvalues(year, 4) = pipt(year, ii)
                    mvalues(year, 5) = pprepeg(year, ii)
                    mvalues(year, 6) = Pavpen(year, ii)
                    mvalues(year, 7) = Tpriv(year, ii)
                    mvalues(year, 8) = brutto(year, ii)
                    mvalues(year, 9) = Skatt(year, ii)
                    mvalues(year, 10) = Bidrag(year, ii)
                    mvalues(year, 11) = disp(year, ii)
                    mvalues(year, 12) = KPI(year)
                    'mvalues(year, 13) = bnp(year)
                End If
            End If
                
        Next year
        
    Next ii
    Close #3
    'Mfast?
    If visaTypfall > 0 Then
        With WsPensioner
            'Set rng = .Range("B21").Offset(1).Resize(slutar - startar, UBound(mvalues, 2))  'Rad 22 till
            Set rng = rng.Offset(1).Resize((slutar - startar)) 'One row below labels
            rng.Value = mvalues
            Set rng = Nothing
        End With
    End If
            
    'Lägg ut annan information, kpi ... BNP per capita i löpande priser
    Open libref & "nyckeltal.csv" For Output As #2
    Print #2, "year; KPI; IBB; PBB; Index; Pindex; bnp_cap"
    For year = (startar) To slutar
        Print #2, year & ";" & _
        Replace(KPI(year), ",", ".") & ";" & _
        Replace(IBB(year), ",", ".") & ";" & _
        Replace(pbb(year), ",", ".") & ";" & _
        Replace(Iindex(year), ",", ".") & ";" & _
        Replace(Pindex(year), ",", ".") & ";" & _
        Replace(bnp(year), ",", ".")
    Next year
    Close #2
    
    'DATA för typisarna finns nu i typis.csv, dessa kan analyseras viare i andra program eller tas upp i Excel
    'I arbetet med typisar har tidigare figur visat utvecklingen för index, kpi mm sedan 1999 -
    ' data finns i modellen men skapas inte här (nyckel2.csv) eftersom figuren kanske inte eftefrågas
    
    Dim resp As String
    If visaTypfall = 0 Then resp = MsgBox("Se  " & libref & "typis.csv, se även nyckeltal.csv", , typerAntal & " typer är klara")


End Sub



''Function ChangeCommas(ByVal myValue As Variant) As String
''
''    Dim temp As String
''    temp = CStr(myValue)
''    ChangeCommas = Replace(temp, ",", ".")
''
''End Function

'''Sub OutputToTextFile()
'''    Dim FileName As String, LineText As String
'''    Dim MyRange As Range, i, j
'''
'''    FileName = "C:\Test\TestFile.txt" 'you can specify here the text file name you want to create
'''
'''    Open FileName For Output As #1
'''
'''    Set MyRange = Range("data") 'it assumes you have a data range named “data” on your worksheet
'''    For i = 1 To MyRange.Rows.Count
'''        For j = 1 To MyRange.Columns.Count
'''            LineText = IIf(j = 1, "", LineText & ",") & MyRange.Cells(i, j) 'the text file creating will have a comma separator
'''        Next j
'''        Print #1, LineText 'using Write command instead of Print will result in having your data in quotes in the output text file
'''    Next i
'''
'''    Close #1
'''
'''End Sub
