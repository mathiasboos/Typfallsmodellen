Attribute VB_Name = "Bidrag"
Option Explicit
Option Base 1                       'Indexering börjar med 1 och inte med noll

Function gp(ByVal inkomst, ByVal civ, ByVal fodar, ByVal pbb, Optional ByVal ftid = 40, Optional marginal = 0, Optional ByVal alder = 65, _
        Optional ByVal year = 2100, Optional ByVal wyear = 21000, Optional ByVal IBB = 0, Optional ByVal kvoten = 1, Optional ByVal rikt = 65, Optional ByVal andel = 1) As Double
    'Inkomst - Beräkningsunderlag vid helt uttag
    'civ     - 0 ogift 1 gift/sammanboende
    'fodar   - Födelseår
    'pbb     - Årets Prisbasbelopp
    'Ftid    - Försäringstid
    'Marginal- 0 avrundning 1 utan avrundning
    'wyear   - När grundskyddet ska inkomstindexeras
    'ibb     - Inkomstbasbeloppet det året
    'Kvoten  - Kvot mellan årets pbb och ibb
    'rikt    - Riktålder
    'Andel   - Uttagsandel av GP
    
    'Application.Volatile
    'Garantipensionen först vid 65 årsdagen (riktåldern from 2020) 'SFS 1998:702
    gp = 0
    
    If alder < rikt Then Exit Function
    'Dim nypropp As Integer
    'nypropp = Application.Range("rng_nypropp")
    
    'Year - inkomstår, wyear=När indexeringen ska börja, [iyear=Utgiftsregler för inkomstår]
    If year > wyear And kvoten < 1 Then
          pbb = IBB 'Ersätt PBB med IBB och korrigera gränserna till de gamla men pensionerna i övrigt följsamhetsindexeras vilket innebär att med tiden får alla GARP
          pbb = pbb * kvoten
    End If
    
    If fodar > 1937 Then
        If civ = 0 Then
            If inkomst <= 1.26 * pbb Then
                gp = 2.13 * pbb - inkomst
                If year > 2019 Then gp = 2.181 * pbb - inkomst
                If year = 2022 Then gp = (2.181 * 7 + 2.43 * 5) * pbb / 12 - inkomst
                If year > 2022 Then gp = 2.43 * pbb - inkomst
            Else
                gp = 0.87 * pbb - 0.48 * (inkomst - 1.26 * pbb)
                If year > 2019 Then gp = 0.921 * pbb - 0.48 * (inkomst - 1.26 * pbb)
                If year = 2022 Then gp = (7 * 0.921 + 5 * 1.17) * pbb / 12 - 0.48 * (inkomst - 1.26 * pbb)
                If year > 2022 Then gp = 1.17 * pbb - 0.48 * (inkomst - 1.26 * pbb)
                
            End If
        Else 'If civ = 1 Then
            If inkomst <= 1.14 * pbb Then
                gp = 1.9 * pbb - inkomst
                If year > 2019 Then gp = 1.951 * pbb - inkomst
                If year = 2022 Then gp = (7 * 1.951 + 5 * 2.2) * pbb / 12 - inkomst
                If year > 2022 Then gp = 2.2 * pbb - inkomst
            Else
                gp = 0.76 * pbb - 0.48 * (inkomst - 1.14 * pbb)
                If year > 2019 Then gp = 0.811 * pbb - 0.48 * (inkomst - 1.14 * pbb)
                If year = 2022 Then gp = (7 * 0.811 + 5 * 1.06) * pbb / 12 - 0.48 * (inkomst - 1.14 * pbb)
                If year > 2022 Then gp = 1.06 * pbb - 0.48 * (inkomst - 1.14 * pbb)
            End If
        End If
        
        If gp < 0 Or ftid < 4 Then gp = 0
        ' Försäkringstid minst tre år
    Else
        If civ = 0 Then
            If year > 2019 And year < 2022 Then inkomst = inkomst + 0.051 * pbb
            If year = 2022 Then inkomst = inkomst + 0.051 * pbb * 7 / 12 + 0.3 * pbb * 5 / 12
            If year > 2022 Then inkomst = inkomst + 0.3 * pbb
            
            If inkomst <= 0.25 * pbb Then
                gp = inkomst * 1.04
            ElseIf inkomst < 1.354 * pbb Then
                gp = 1.5174 * inkomst - 0.1193 * pbb
            ElseIf inkomst <= 1.529 * pbb Then
                gp = 1.343 * inkomst + 0.1168 * pbb
            ElseIf inkomst <= 3.16 * pbb Then
                gp = 2.17 * pbb + 0.6 * (inkomst - 1.51 * pbb)
            Else
                gp = 0
            End If
            If year > 2019 And inkomst < 3.16 * pbb Then gp = gp + 0.051 * pbb
        
        Else 'If civ = 1 Then
            If inkomst <= 0.25 * pbb Then
                gp = inkomst * 1.04
            ElseIf inkomst < 1.354 * pbb Then
                gp = 1.5174 * inkomst - 0.1193 * pbb
            ElseIf inkomst <= 2.8275 * pbb Then
                gp = 1.935 * pbb + 0.6 * (inkomst - 1.34 * pbb)
            Else
                gp = 0
            End If
            
            If year > 2019 And inkomst < 2.8275 * pbb Then gp = gp + 0.051 * pbb
        
            End If
         
        gp = gp - inkomst
        If gp < 0 Then gp = 0
        ' Berättigad till folkpension dec 2002
    End If
    
    'Försäkringstiden 40 år
    If (ftid < 40 And gp > 0) Then gp = gp * ftid / 40
    'Uttagsandel lägre än helt uttag
    If andel < 1 Then gp = gp * andel
    'Avrundning
      If marginal = 0 Then gp = 12 * Int(gp / 12 + 0.5)

End Function

''Sub kollgp()
''
''    Dim kollgp As Double
''    Rem GP(Inkomst, civ, fodar, PBB, Optional ftid = 40, Optional marginal = 0, Optional alder = 65, _
''    Optional year = 2100, Optional wyear = 21000, Optional ibb = 0, Optional Kvoten = 1) As Double
''    'kollgp = GP(600000, 0, 1985, 92700, 0, 65, 2050, 2010, 213000, 42400 / 51100)
''    Dim inkomst As Double
''    Dim civ As Single, fodar As Single
''    inkomst = 0
''    civ = 0
''    'fodar = 1937
''    'kollgp = gp(inkomst * 185 / 160, civ, fodar, 44500, 50, 0, 65, 2019, 2011, 51100, 1)
''    'Debug.Print inkomst; kollgp / 12
''
'' For fodar = 1960 To 1965 ' = 6800 To 8400 Step 100
''     kollgp = gp(inkomst * 185 / 160, civ, fodar, 48300, 50, 0, 65, 2023, 2011, 51100, 1)
''
''      Debug.Print fodar; inkomst; kollgp / 12
''       inkomst = inkomst + 1000
'' Next fodar ' inkomst
''End Sub
'Funktionerna om BTP finns under bidrag


'''Function Garantitill(ByVal inkomst, ByVal year, Optional ByVal index = 197.69, Optional ByVal index2022 = 197.69, _
'''                    Optional ByVal andel = 1, Optional ByVal ftid = 40) As Single
'''    'Inkomst - underlag för garanttillägget (GT) till BTP
'''    'Year   - Inkomstår
'''    'Index  - Inkomstindex t
'''    'Index2022 -Inkomstindex t-1
'''    'Andel  - uttag av inkomstpension
'''    'Ftid   - Försäkringstid
'''    Garantitill = 0
'''    If year < 2022 Or ftid < 1 Or andel = 0 Or Application.Range("rng_nypropp") = 0 Then Exit Function
'''
'''    Dim mbelopp As Double 'Max belopp
'''    Dim lim1, lim2 As Double ' Gränsvärden
'''    Dim namn As Double
'''    namn = 1
'''    mbelopp = 12000 'per år
'''    lim1 = 142153: lim2 = 172153
'''
'''    If year > 2022 Then
'''        namn = 1.016 ^ (year - 2022)
'''        mbelopp = (mbelopp - 600) * (index / index2022) / namn + 600
'''        lim1 = (lim1 - 600) * (index / index2022) / namn + 600
'''        lim2 = (lim2 - 600) * (index / index2022) / namn + 600
'''
'''        If marginal = 0 Then
'''            mbelopp = Int(mbelopp)
'''            lim1 = Int(lim1)
'''            lim2 = Int(lim2)
'''        End If
'''    End If
'''
'''    If inkomst < lim1 Then
'''        Garantitill = mbelopp
'''    ElseIf inkomst < lim2 Then
'''        Garantitill = mbelopp - 0.38 * (inkomst - lim1)
'''        If Garantitill < 600 Then Garantitill = 600
'''    End If
'''    If andel > 1 Then andel = 1
'''    Garantitill = Garantitill * andel
'''    If ftid < 40 Then Garantitill = Garantitill * ftid / 40
'''
'''    'If year = 2022 Then Garantitill = Garantitill * 5 / 12
'''    'Korr för antal pmonth som är max 4 för 2022
'''    If marginal = 0 Then Garantitill = Int(Garantitill + 0.5)
'''
'''End Function

'KOLL
''Sub gtill()
''    Dim gtill As Double
''    Dim inkomst As Double
''    inkomst = 160000
''    Dim year As Integer
''    Dim index As Double
''    index = 197.69
''
''    For year = 2022 To 2030
''
''        gtill = Garantitill(inkomst, year, index)
''        Debug.Print year; Round(index, 2); gtill
''
''        index = index * 1.018
''        inkomst = 1.018 / 1.016
''
''    Next year
''
''End Sub


Function tillagg(ByVal underl, ByVal year, Optional ByVal index = 202.84, Optional ByVal index2021 = 186.52, _
Optional ByVal andel = 1, Optional ByVal ftid = 40) As Single 'Double
    'Inkomst pensionstillägg (IPT) se ds 2020:7 kolla SFS 2010:110 på Lagrummet.se
    'Inkomst underlag för IPT månadsvis * 12?
    'YearInkomstår gränserna för tillägget ska indexeras
    'Index - inkomstindex aktuellt år
    'Andel - uttagsandel
    'Ftid  - Försäkringstid
    
''    If Application.Range("Rulesfromutg") = 0 Or IsMissing(Application.Range("Rulesfromutg")) Then
''             'Nothing
''        Else
''            If (year > Application.Range("Rulesfromutg") And Application.Range("Rules") = 0) _
''            Or Application.Range("Rules") = 1 Then
''                year = Application.Range("Rulesfromutg")
''            End If
''    End If
    
    tillagg = 0
    If year < 2021 Or ftid < 1 Then Exit Function
       
    Dim i As Double
    'Fix 2025-12-01, måste i indexering ta hänsyn till varje års förändring av gränserna därav upphöjt i år - 2021
    i = ((index / index2021) / (1.016 ^ (year - 2021)))
    
    Dim u As Double
    u = underl / i
    
    Dim tl_ipt As Double
    If u < 108000 Or u >= 204000 Then Exit Function
    
   If u < 132000 Then
        tl_ipt = (u - 108000) * (0.3)
        If marginal = 0 Then tl_ipt = Int((u - 108000) * (0.3) / 600) * 600 + 300
    ElseIf u <= 168000 Then
        tl_ipt = 7200
    Else
        tl_ipt = 7200 - (u - 168000) * 0.2
        If marginal = 0 Then tl_ipt = Int((7200 - (u - 168000) * 0.2) / 600) * 600 + 300
    End If
    
    tillagg = tl_ipt
    If andel <= 1 Then tillagg = tillagg * andel
    'if andel>1 then msgbox("FEL andel")
    
    If Int(born) > 1944 Then
        If ftid < 40 Then tillagg = tillagg * ftid / 40
    ElseIf Int(born) > 1937 Then
        If ftid < 35 Then tillagg = tillagg * ftid / 35
    ElseIf Int(born) > 1924 Then
        If ftid < 30 Then tillagg = tillagg * ftid / 30
    ElseIf Int(born) > 1914 Then
         If ftid < (20 + Int(born) - 1915) Then tillagg = tillagg * ftid / (20 + Int(born) - 1915)
    Else
        If ftid < 20 Then tillagg = tillagg * ftid / 20
    End If
    
End Function

''Sub koll_ipt()
''
''    Dim ink, koll As Double
''    Dim i As Long
''    For i = 107 To 120
''        ink = i * 1000 + 995
''        koll = (tl_ipt(ink + 100, 1.016, 1, 1) - tl_ipt(ink, 1.016, 1, 1))
''
''        Debug.Print ink; koll; tl_ipt(ink + 100, 1.016, 1, 1); tl_ipt(ink, 1.016, 1, 1)
''    Next i
''
''End Sub


''Function Xtillagg(ByVal inkomst, ByVal year, Optional ByVal index = 194.19, Optional ByVal index2021 = 194.19, _
''    Optional ByVal andel = 1, Optional ByVal ftid = 40) As Single
''
''    'Inkomst pensionstillägg (IPT) se ds 2020:7 kolla SFS 2010:110 på Lagrummet.se
''    'Inkomst underlag för IPT månadsvis * 12?
''    'YearInkomstår gränserna för tillägget ska indexeras
''    'Index - inkomstindex aktuellt år
''    'Andel - uttagsandel
''    'Ftid  - Försäkringstid
''
''''    If Application.Range("Rulesfromutg") = 0 Or IsMissing(Application.Range("Rulesfromutg")) Then
''''             'Nothing
''''        Else
''''            If (year > Application.Range("Rulesfromutg") And Application.Range("Rules") = 0) _
''''            Or Application.Range("Rules") = 1 Then
''''                year = Application.Range("Rulesfromutg")
''''            End If
''''    End If
''
''    tillagg = 0
''    If year < 2021 Or ftid < 1 Then Exit Function
''
''    Dim Rtillagg As Range
''    Set Rtillagg = Application.Range("RngTillagg")
''    Dim Mtillagg(1 To 25, 1 To 3) As Double
''
''    'om year>2021 ska gränsvärdena räknas om med följsamhetsindex: Men nedan är 2021 år frågan är om inte beloppen ska vara gräns(t)=gräns(t-1)* index(t)/index(t-1)
''    ' Då måste beloppen sparas undan nedan en förenkling
''    Dim r As Integer
''
''    If year >= 2021 Then
''        For r = 1 To 25
''           Mtillagg(r, 1) = Rtillagg(r, 1) * ((index / index2021) / (1.016 ^ (year - 2021)))
''           Mtillagg(r, 2) = Rtillagg(r, 2) * ((index / index2021) / (1.016 ^ (year - 2021)))
''           Mtillagg(r, 3) = Rtillagg(r, 3)
''           If marginal = 0 Then
''                Mtillagg(r, 1) = Int(Mtillagg(r, 1)) 'Avrundas till hela kronor
''                Mtillagg(r, 2) = Int(Mtillagg(r, 2))
''           End If
''        Next r
''    End If
''    'inkomst = inkomst / ((index / index2021) / (1.016 ^ (year - 2021)))
''    'If year > 2021 Then inkomst = inkomst / ( (index / index2021) / (1.016))
''
''    If inkomst < Mtillagg(1, 1) Then Exit Function
''    If inkomst > Mtillagg(25, 2) Then Exit Function
''
''    r = 1
''    Do While inkomst >= Mtillagg(r, 1) And r <= 25
''       tillagg = Mtillagg(r, 3)
''        'Debug.Print inkomst; r; Rtillagg(r, 3); Rtillagg(r, 2); Rtillagg(r, 1)
''        r = r + 1
''        If r = 26 Then Exit Do
''     Loop
''    If andel <= 1 Then tillagg = tillagg * andel
''    'if andel>1 then msgbox("FEL andel")
''
''    If Int(born) > 1944 Then
''        If ftid < 40 Then tillagg = tillagg * ftid / 40
''    ElseIf Int(born) > 1937 Then
''        If ftid < 35 Then tillagg = tillagg * ftid / 35
''    ElseIf Int(born) > 1924 Then
''        If ftid < 30 Then tillagg = tillagg * ftid / 30
''    ElseIf Int(born) > 1914 Then
''         If ftid < (20 + Int(born) - 1915) Then tillagg = tillagg * ftid / (20 + Int(born) - 1915)
''    Else
''        If ftid < 20 Then tillagg = tillagg * ftid / 20
''    End If
''
''    'Korr för antal pmonth
''
''End Function

''Sub kolla_till()
''
''    Dim koll As Single
''    koll = tillagg(203000, 2021, 186.25, 184.41)
''    Debug.Print koll
''    'Rows("17:17").RowHeight = 15
''
''End Sub



Function BTP(ByVal inkomst As Double, ByVal inkomstm As Double, ByVal hyra As Double, ByVal gift As Integer, ByVal pbb As Double, _
            Optional ByVal ap = 1, Optional ByVal apm = 1, Optional ByVal form = 0, Optional ByVal arb = 0, Optional ByVal ArbM = 0, _
            Optional ByVal marginal = 0, Optional ByVal maxhyra = 0, Optional ByVal year = 2014, Optional ByVal wyear = 2100, _
            Optional ByVal IBB = 0, Optional ByVal kvoten = 1, Optional ByVal bald = 65, Optional ByVal TJP = 0, Optional ByVal tjpm = 0, _
            Optional ByVal garp = 0, Optional ByVal garpm = 0, Optional ByVal IBB0 = 0, _
            Optional ByVal mgarp = 0, Optional ByVal month = 12, Optional ByVal ftid = 40, _
            Optional ByVal index = 100, Optional uttag = 1, Optional ByVal ansokt = 1) As Double
            
    'Application.Volatile
    'Beräkning av BTP på årsbasis:
        'Inkomst  - inkomst
        'Inkomstm - Makans inkomst för sammanboende
        'Hyra     - Boendekostnad (månad)
        'Gift     - Sammanboende =1, Ensamstående=0
        
        'PBB      - Gällande prisbasbelopp
        'Ap       - Ålderspensionär (ja=1, nej=0)
        'Apm      - Ålderspensionär maka (ja=1, nej=0)
        'Form     - Hushållets förmögenhet
        
        'arb      - arbetsinkomst
        'arbm     - arbetsinkomst
        'Marginal - slopa avrundningar mm
        'maxhyra  - Maximal hyra för BTP
        
        'Year     - Inkomst år
        'wyear    - om gränserna ska inkomstjusteras och i så fall när
        'IBB      - För att justera gränserna
        'Kvoten   - med kvottalet pbb/ibb för wyear
        
        'bald     - Ålder
        'tjp      - tjänstepension
        'tjpm     - tjänstepension maka
        'Garp     - garantipension
        
        'Garp(m)  - garantipension maka
        'IBB0     - IBB vid wyear
        
        'mgarp      - Tidigare Underlaget för garp per månad obsolet
        'month      - Antal månader
        'Ftid       - Bosättningskriteriet
        'Index      - Årets inkomstindex - användes tidigare för garantitillägget ska tas bort men obs anropet...
        'Index2     - Index 2022
        'Sökt       - Ansökt
        
        BTP = 0
        Dim max As Double
        Dim maxm As Double
        max = 0: maxm = 0
        
        If year < 1978 Then Exit Function 'KBT införs 1/1 1978
            
        If year >= wyear And kvoten < 1 Then
           pbb = IBB            'Ersätt pbb med IBB och korrigera gränserna till de gamla
           pbb = pbb * kvoten
        End If
        
        Dim maxohyra As Double  'Maxohyra maximal årshyra som BTP beräknas på
        Dim BTPm As Double      'Makens btp
        
        'If hyra < 15000 Then hyra = hyra * 12 ' Antagen månadshyra till årshyra'
        
        Dim PAR, par2, par3, par4 As Double       'Andel av hyran/boendekostnaden
        PAR = 0.9: par2 = 0: par3 = 0: par4 = 0
        
        If year < 1995 Then '1995 infördes BTP max boende mellan 100 och 4000 samt 85%
          PAR = 0.83
          If hyra > 150 Then
            maxohyra = CLng(3500) * 12
          Else
            maxohyra = 0
        End If
        
        ElseIf year < 2001 Then
          PAR = 0.85
          If hyra > 100 Then
            maxohyra = CLng(3900) * 12
          Else
            maxohyra = 0
        End If
          
        If year = 1998 Then PAR = 0.85
        If year = 1999 Then PAR = 0.9
        
        ElseIf year <= 2003 Then
          PAR = 0.9
          maxohyra = CLng(4500) * 12
          If year > 2001 Then PAR = 0.91
        ElseIf year <= 2005 Then
           PAR = 0.91
           maxohyra = CLng(4670) * 12
        ElseIf year = 2006 Then
           PAR = (0.91 * 3 + 0.93 * 9) / 12 'Om pension mitt på året en underskattning
           If ap = 0 Then PAR = 0.91
           maxohyra = CLng(4850) * 12
        ElseIf year = 2007 Then
           PAR = 0.93
           If ap = 0 Then PAR = 0.91
           maxohyra = CLng(5000) * 12
        Else
          PAR = 0.93
          maxohyra = CLng(5000) * 12
          If ap = 0 And year < 2009 Then PAR = 0.91
        End If
        
        If year = 2015 And ap = 1 Then
           PAR = (0.93 * 8 / 12 + 0.95 * 4 / 12) 'Dito antar jämn inkomst och hyra under året'
        ElseIf year > 2015 And ap = 1 Then
           PAR = 0.95
        ElseIf year > 2017 And ap = 0 Then
           PAR = 0.95       '95% oavsett SA eller Pensionär
        End If
    
        If year > 2017 Then
           PAR = 0.96       'och 70% av hyran mellan 5600 och 5000
           par2 = 0.7
           maxohyra = CLng(5600) * 12
        End If
        
        If year > 2019 Then 'egentligen infört 1/12 2019
            PAR = 1
            par2 = 0.9
            par3 = 0.7
            maxohyra = CLng(7000) * 12
            
            If year >= 2022 Then 'se nedan
                 maxohyra = CLng(7500) * 12
                 par4 = 0.5
            End If
            
        End If
        
        'If maxhyra > 60000 And year > 2016 Then Maxohyra = maxhyra * (Maxohyra / 60000) ' Trappas upp linjärt med maxhyra där nuvarande max är 60000
        'If hyra > maxohyra Then hyra = maxohyra 'olyckligt med hyra...
        If maxhyra > 0 And year >= wyear Then maxohyra = maxhyra * maxohyra 'Maxhyra om gränsvärdet ska ändras
        If hyra > maxohyra Then hyra = maxohyra
        'Koll mot maximal hyra
        Dim hyram As Double 'Kan kanske slopas då hyra=hyra/2
        hyram = 0
        'OBS Särskilt boende har en lägre gräns för maximalt boende, ingen hänsyn till detta
        
        'Antal månader som pensionär - justeras i huvudprogrammet
        
        'Om inkomstindexering: Gränsvärdena justeras
        If year >= wyear Then hyra = hyra * IBB0 / IBB
        
        If gift = 1 Then
            hyra = hyra / 2
            hyram = hyra
        End If
        
        If year < 2018 Then
            max = hyra * PAR
            maxm = hyram * PAR
        ElseIf year < 2020 Then
             If gift = 0 Then ' /(gift + 1)
                 If hyra < 60000 Then
                      max = hyra * PAR
                 ElseIf hyra < 67200 Then
                      max = 60000 * PAR + (hyra - 60000) * par2
                 Else
                     max = 60000 * PAR + (67200 - 60000) * par2
                 End If
                
             Else
                 If hyra < 30000 Then
                      max = hyra * PAR
                 ElseIf hyra < 33600 Then
                      max = 30000 * PAR + (hyra - 30000) * par2
                 Else
                     max = 30000 * PAR + (33600 - 30000) * par2
                 End If
                 maxm = max
            End If
        ElseIf year < 2022 Then 'year>=2020
                If gift = 0 Then
                     If hyra < 36001 Then
                          max = hyra * PAR
                     ElseIf hyra < 60001 Then
                          max = 36000 * PAR + (hyra - 36000) * par2
                     ElseIf hyra < 84001 Then
                          max = 36000 * PAR + (60000 - 36000) * par2 + (hyra - 60000) * par3
                     Else
                         max = 36000 * PAR + (60000 - 36000) * par2 + (84000 - 60000) * par3
                         End If
                Else
                     If hyra < 18001 Then
                          max = hyra * PAR
                     ElseIf hyra < 30001 Then
                          max = 18000 * PAR + (hyra - 18000) * par2
                     ElseIf hyra < 42001 Then
                          max = 18000 * PAR + (30000 - 18000) * par2 + (hyra - 30000) * par3
                     Else
                         max = 18000 * PAR + (30000 - 18000) * par2 + (42000 - 30000) * par3
                     End If
                     maxm = max
                End If
        Else
            If gift = 0 Then
                     If hyra < 36001 Then
                          max = hyra * PAR
                     ElseIf hyra < 60001 Then
                          max = 36000 * PAR + (hyra - 36000) * par2
                     ElseIf hyra < 84001 Then
                          max = 36000 * PAR + (60000 - 36000) * par2 + (hyra - 60000) * par3
                     ElseIf hyra < 90001 Then
                            max = 36000 * PAR + (60000 - 36000) * par2 + (84000 - 60000) * par3 + (hyra - 84000) * par4
                     Else
                         max = 36000 * PAR + (60000 - 36000) * par2 + (84000 - 60000) * par3 + (90000 - 84000) * par4
                     End If
                     
                Else
                     If hyra < 18001 Then
                          max = hyra * PAR
                     ElseIf hyra < 30001 Then
                          max = 18000 * PAR + (hyra - 18000) * par2
                     ElseIf hyra < 42001 Then
                          max = 18000 * PAR + (30000 - 18000) * par2 + (hyra - 30000) * par3
                     ElseIf hyra < 45001 Then
                         max = 18000 * PAR + (30000 - 18000) * par2 + (42000 - 30000) * par3 + (hyra - 42000) * par4
                     Else
                         max = 18000 * PAR + (30000 - 18000) * par2 + (42000 - 30000) * par3 + (45000 - 42000) * par4
                     End If
                     maxm = max
                End If
        End If
        
    
        Dim Extra As Double 'Extra consuming support
        
        If year >= 2012 And bald > 64 And year < 2022 Then  'And gift = 0
            Extra = (340 * 12) / (1 + gift)
        End If
 
        If year = 2022 Then
             Extra = (540 * 7 + 840 * 5) / (gift + 1) '1/8 + 200 kr obs om pensionering mitt på året en underskattning
        End If
        
        If year > 2022 Then Extra = (840 * 12) / (gift + 1)
        'Add to max
        max = max + Extra
        maxm = maxm + Extra * gift
        
        If year >= wyear Then
            max = max * IBB / IBB0 ' Och uttryckt i årets priser/löner
            maxm = maxm * IBB / IBB0
        End If
        
        ' Add assets
        If year < 2001 Then
            If gift = 1 Then
                If form > 75000 Then form = 0.05 * (form - 75000)
            Else
                If form > 120000 Then form = 0.1 * (form - 120000)
            End If
        Else
            If form > 100000 Then form = 0.15 * (form - 100000)
        End If
        If form < 0 Then form = 0
            
        inkomst = inkomst + form
        inkomstm = inkomstm + form
    
        ' Reduceringsinkomst, allmänna pensioner och ersättningar 100% övriga 50% viktningen, tidigare (2007?) 80%
        Dim fri As Double 'Fribelopp uttryckt i #prisbasbelopp?
        Dim Red As Double
        Dim Redm As Double
        
        'inkomst = inkomst - arb - TJP 'Inkomster med 100% vikt
        
        If year >= 2014 And (arb > 0 Or ArbM > 0) Then
            If year >= wyear Then
                arb = WorksheetFunction.max(arb - 24000 * IBB / IBB0, 0)
                ArbM = WorksheetFunction.max(ArbM - 24000 * IBB / IBB0, 0)
            Else
                arb = WorksheetFunction.max(arb - 24000, 0)
                ArbM = WorksheetFunction.max(ArbM - 24000, 0)
            End If
        End If
        
        
        If gift = 0 Then
            If ap = 1 Then
                fri = pbb * 2.17
                If year > 2019 Then fri = 2.181 * pbb
                If year = 2022 Then fri = (7 * 2.181 + 5 * 2.43) * pbb / 12
                If year > 2022 Then fri = 2.43 * pbb
            Else
                fri = pbb * 2.4
            End If

            If year < 2008 Then
               Red = (inkomst - arb - TJP) + 0.8 * arb + 0.8 * TJP - fri
            Else
               Red = (inkomst - arb - TJP) + 0.5 * arb + 0.8 * TJP - fri
            End If

            If year > 2019 Then
                Red = garp + (inkomst - garp) * 0.93 - fri '+gtill
            End If
            
            If Red < 0 Then Red = 0

        End If

        If gift = 1 Then
            If ap = 1 Then
                fri = pbb * 1.935
                If year > 2019 Then fri = 1.951 * pbb
                If year = 2022 Then fri = (7 * 1.951 + 5 * 2.2) * pbb / 12
                If year > 2022 Then fri = 2.2 * pbb
            Else
                fri = pbb * 2.4
            End If

            If year < 2008 Then
               Red = (inkomst - arb - TJP) + 0.8 * arb + 0.8 * TJP - fri
               Redm = (inkomstm - ArbM - tjpm) + 0.8 * ArbM + 0.8 * tjpm - fri
            Else
               Red = (inkomst - arb - TJP) + 0.5 * arb + 0.8 * TJP - fri '+gtill
               Redm = (inkomstm - ArbM - tjpm) + 0.5 * ArbM + 0.8 * tjpm - fri '+ gtillm
            End If
            
            If year > 2019 Then
                Red = garp + (inkomst - garp) * 0.93 - fri  '+gtill
                Redm = garpm + (inkomstm - garpm) * 0.93 - fri  ' +gtillm
            End If

            If Red < 0 Then Red = 0
            If Redm < 0 Then Redm = 0
            'OBS Gemensam reduktion
            Red = Int((Red + Redm) / 2)
        End If

       'BTP beräknas
        If Red < pbb Then
           BTP = max - 0.62 * Red
        Else
           BTP = max - (Red - pbb) * 0.5 - pbb * 0.62
        End If

        If year > 2019 Then '
                 BTP = max - 0.62 * Red
        End If

        If BTP < 0 Then BTP = 0
        If gift = 1 And ap = 1 Then BTPm = BTP
        
        If Application.Range("rng_dela") = 1 Then
             BTP = BTP + BTPm 'OBS SBTP reducerar hela
        ElseIf Application.Range("rng_dela") = 2 Then
             BTP = BTP '+ BTPm
        Else
            BTP = (BTP + BTPm) / (gift + 1)
        End If
             
        'Lägg till SBTP, och sedan avrundas till hela kronor. Om ej sbtp, månadsbelopp under 25 kr betalas inte ut
        ' OBS ingen reducering för försäkringstid
End Function

''Sub kolla_BTP()
''    'Call startsetup
''    Dim kolla As Double
''    Dim ink As Double
''    Dim inkM As Double
''
''    inkM = 0
''
''    Dim hyra As Double: hyra = 9000
''    Dim form As Double: form = 0
''    Dim arb As Double: arb = 0
''    Dim TJP As Double: TJP = 0
''    Dim gift As Integer: gift = 0: 'If inkM > 0 Then gift = 0
''    Dim prisb As Single
''    prisb = 52500
''    Dim kvot As Double
''    Dim mpension As Double
''    Dim pmonth As Integer
''    pmonth = 12
''
''    Rem (ByVal inkomst As Double, ByVal inkomstm As Double, ByVal Hyra As Double, ByVal gift As Integer, ByVal PBB As Double, _
''        Optional ByVal ap = 1, Optional ByVal apm = 1, Optional ByVal form = 0, Optional ByVal arb = 0, Optional ByVal ArbM = 0, _
''        Optional ByVal marginal = 0, Optional ByVal maxhyra = 0, Optional ByVal year = 2014, Optional ByVal wyear = 2100, _
''        Optional ByVal IBB = 0, Optional ByVal Kvoten = 1, Optional ByVal Bald = 65, Optional ByVal TJP = 0, Optional ByVal tjpm = 0, _
''        Optional ByVal garp = 0, Optional ByVal garpm = 0, Optional ByVal IBB0 = 0) As Double
''    Dim year, wyear  As Integer
''    wyear = 2099
''    For year = 2023 To 2023
''        ink = gp(0, 0, 1955, 48300, 40, 0, 66, year)
''        mpension = ink
''       ' ink = Worksheets("Oles tablå").Cells(37, 1 + year - 2001) + _
''           Worksheets("Oles tablå").Cells(38, 1 + year - 2001) + _
''           Worksheets("Oles tablå").Cells(39, 1 + year - 2001)
''       ' ink = ink * 12
''       ' prisb = Worksheets("Några faktorer").Cells(year - 1994 + 1, 5)
''       ' Hyra = Worksheets("Antaganden").Cells(12, 2 + year - 2001)
''''       kvot = 55000 / 56900
''''        If year = 2021 Then
''''            kvot = 1: wyear = 2100
''''        End If
''        kolla = BTP(ink, inkM, hyra, gift, prisb, _
''        1, 0, form, arb, 0, _
''        marginal, 0, year, wyear, _
''        56900, kvot, 65, TJP, 0, _
''        ink, inkM, 55000, mpension, pmonth, 40, 101.6, 100, 1, 1)
''        'MsgBox kolla '(kolla / 12)
''        Debug.Print year; ink & " " & Round(kolla / 12, 0)
''        'ink = ink + 1000 * klass
''
''    Next year
''End Sub

Function SBTP(ByVal inkomst, ByVal hyra, ByVal gift, ByVal btpb, ByVal avdrag, Optional ByVal skattesats = 0.316, Optional ByVal ap = 1, _
    Optional ByVal form = 0, Optional ByVal pbb = 42800, Optional ByVal maxhyra = 0, _
    Optional ByVal year = 2014, Optional ByVal wyear = 21000, Optional ByVal IBB = 0, Optional ByVal kvoten = 0, Optional ByVal bald = 65, _
    Optional ByVal kapital = 0, Optional ByVal inkomstm = 0, _
    Optional ByVal marginal = 0, Optional ByVal baldm = -99, Optional ByVal ftid = 1) As Double
    
    'Application.Volatile
        'Inkomst(m) - Inkomst
        'Hyra       - Boendekostnad
        'gift       - 0/1 Ensam/sammanboende
        'btpb       - Bostadstilägg som ska läggas till
        'Avdrag     - Grundavdrag kan tas bort då den räknas ut
        'Skattesats (kommunal)
        'AP         - Ålderpensionär kan tas bort då bald finns
        
        'Form       - förmögenhet
        'pbb        - prisbasbelopp
        'maxhyra    - Maximal skälig boendekostnad
        
        'year       - Inkomst/bidragsår
        'wyear      - År för när gränser ev. ska inkomsindexeras
        'IBB        - vad uppgår då Inkomstbasbeloppet till?
        'Kvoten     - Vad är relationen mellan PBB/IBB, dvs kvot*IBB
        'Bald       - Make/maka har samma ålder om inget annat anges
        
        'kapital    - Kapitalinkomst finns dock inte i modellen
        'Inkomstm
        'marginal   - marginalberäkning
        'Baldm
        'ftid       - Försäkrad och får SBTP

    SBTP = 0
    
''    If Application.Range("Rulesfromutg") = 0 Or IsMissing(Application.Range("Rulesfromutg")) Then
''             'Nothing
''        Else
''            If (year > Application.Range("Rulesfromutg") And Application.Range("Rules") = 0) _
''            Or Application.Range("Rules") = 1 Then
''                year = Application.Range("Rulesfromutg")
''            End If
''    End If
    
    If year < 1978 Then Exit Function
    
    'Dim nypropp As Integer
    'nypropp = Application.Range("rng_nypropp")
    
    bald = Int(bald)
    baldm = Int(baldm)
    If baldm < 0 Then baldm = bald
      
    If hyra < 10000 Then hyra = hyra * 12 ' Månadshyra till årshyra'
    
    If year >= wyear And kvoten < 1 Then
        pbb = IBB 'Ersätt pbb med IBB och korrigera gränserna till de gamla
        pbb = pbb * kvoten
     End If
     
     form = (form - 100000) * 0.15 / (gift + 1)
     If form < 0 Then form = 0
     
    'Skälig hyra 1978-1996 ?
    'Skälig hyra 1997-2014
   
    Dim maxohyra As Double
    If year <= 1997 Then
      maxohyra = CLng(5200) * 12
    ElseIf year <= 2000 Then
      maxohyra = CLng(5200) * 12
    ElseIf year <= 2004 Then
      maxohyra = CLng(5700) * 12
    ElseIf year = 2005 Then
      maxohyra = CLng(5870) * 12
    ElseIf year = 2006 Then
      maxohyra = CLng(6050) * 12
    ElseIf year < 2010 Then
      maxohyra = CLng(6050) * 12
    ElseIf year < 2017 Then
      maxohyra = CLng(6200) * 12
    ElseIf year < 2020 Then
      maxohyra = CLng(6620) * 12
    Else
      maxohyra = CLng(7500) * 12
    End If
     
    If maxhyra > 0 And year >= wyear Then maxohyra = maxhyra * maxohyra
     
    maxhyra = maxohyra / (gift + 1)
      
    'Dim SBTP As Double
    Dim Lev As Double
    Dim Besk As Double
    Dim Beskm As Double
    Dim disp As Double
    Dim Dispm As Double
    
    Dim ctxfvi, ctxfvim As Double
    ctxfvi = inkomst
    ctxfvim = inkomstm
    If marginal = 0 Then
        ctxfvi = Int(inkomst / 100) * 100
        ctxfvim = Int(inkomstm / 100) * 100
    End If
    
    hyra = hyra / (gift + 1)
    If hyra > maxhyra Then hyra = maxhyra
    
    'Skälig levnadsnivå, PBB= PrisBasBeloppet 'ev se över historiskt
    If year < 2009 Then 'Från 2004 ok
        If gift = 0 Then
           Lev = 1.294 * pbb
        Else
           Lev = 1.084 * pbb
        End If
    ElseIf year <= 2011 Then
        If gift = 0 Then
           Lev = 1.3546 * pbb
        Else
           Lev = 1.1446 * pbb
        End If
    End If
    
    If year = 2012 And ap = 1 Then
        If gift = 0 Then
           Lev = 1.401 * pbb
        Else
           Lev = 1.191 * pbb
        End If
    ElseIf year > 2012 And ap = 1 Then
        If gift = 0 Then
           Lev = 1.4468 * pbb
        Else
           Lev = 1.191 * pbb
        End If
    End If
    
    If year > 2015 And ap = 1 Then
        If gift = 0 Then
           Lev = 1.473 * pbb
        Else
           Lev = 1.204 * pbb
        End If
    End If
    
    If year > 2017 Then
        If gift = 0 Then
           Lev = 1.486 * pbb
        Else
           Lev = 1.2105 * pbb
        End If
    End If
    If year > 2021 Then
        If gift = 0 Then
           Lev = 1.5357 * pbb
        Else
           Lev = 1.2353 * pbb
        End If
    End If
    'Debug.Print year & " " & gift & " " & Lev / 12; pbb
    If marginal = 0 Then Lev = Int(Lev / 12) * 12
    
    If ftid < 1 Then GoTo A_F_S '--> Äldreförsörjningsstöd -->
    
    'Inkomst netto
    Beskm = 0
    Dim Xage As Long
    Xage = riktage(CLng(year), 1) + 1
    
    Besk = ctxfvi - avdragxx(inkomst, pbb, marginal, bald, year, 21000, IBB, kvoten, Iyear, Xage)
    
    If gift = 1 And (Application.Range("rulesfromUtg") = 0 Or IsMissing(Application.Range("rulesfromUtg"))) Then
''        Besk = ctxfvi - avdragxx(inkomst, pbb, marginal, Bald, year, 21000, IBB, kvoten, Iyear, Xage)
        Beskm = ctxfvim - avdragxx(inkomstm, pbb, marginal, bald, year, 21000, IBB, kvoten, Iyear, Xage)
    Else
         Beskm = ctxfvim - avdragxx(inkomstm, pbb, marginal, bald, Application.Range("rulesfromUtg"), 21000, IBB, kvoten, Iyear, Xage)
    End If
       
    Dim kapitalm:
    kapitalm = kapital
    
    Dispm = 0
    If marginal = 0 Then
        disp = ctxfvi - Int(Besk * skattesats) + Int(kapital * 0.7)  '+ form  - hyra + btpb / (gift + 1)
        If gift = 1 Then Dispm = ctxfvim - Int(Beskm * skattesats) + Int(kapitalm * 0.7)
    Else
        disp = ctxfvi - Besk * skattesats + kapital * 0.7  '+ form  - hyra + btpb / (gift + 1)
        If gift = 1 Then Dispm = ctxfvim - (Beskm * skattesats) + (kapitalm * 0.7)
    End If
    If age <= slutage Then
        disp = disp - PublicAvg(Besk, 0.01, age, marginal, year_(age))
        Dispm = Dispm - PublicAvg(Beskm, 0.01, age, marginal, year_(age))
    End If
    If disp < 0 Then disp = 0
    If Dispm < 0 Then Dispm = 0
        
        
   'Disp ska jämföras med fribeloppet som lägsta inkomst
    Dim fri As Double
    If year > 2007 Then
       If gift = 0 Then
         fri = 2.17 * pbb
       Else
        fri = 1.935 * pbb
       End If
    Else 'före 2017
      If gift = 0 Then
         fri = 2.13 * pbb
       Else
        fri = 1.9 * pbb
       End If
    End If
    
    If year > 2019 Then
        If gift = 0 Then
             fri = 2.181 * pbb
        Else
            fri = 1.951 * pbb
        End If
        'Födda 1937 och tidigare 2.221 och 1.986
        If (year - bald) < 1938 Then
          If gift = 0 Then
             fri = 2.221 * pbb
           Else
             fri = 1.986 * pbb
           End If
           If year > 2022 Then
                If gift = 0 Then
                    fri = 2.47 * pbb
                Else
                 fri = 2.235 * pbb
                End If
           End If
           
        End If
    End If

    'Beräknar lägsta inkomst som ska tas upp för SBTP efter skatteavdrag
    If IsMissing(born) Or born < 1900 Then born = Application.Range("Born")
    Dim frinetto As Double
    If gift = 0 Then
        If born > 1937 Then
            frinetto = 2.13 * pbb - (2.13 * pbb - avdrag) * skattesats 'avdragxx(inkomst, PBB, marginal, Bald, year, 21000, IBB, Kvoten, iyear, riktalder)
        Else
            frinetto = fri - (fri - avdrag) * skattesats
        End If
    Else
        If born > 1937 Then
            frinetto = 1.9 * pbb - (1.9 * pbb - avdrag) * skattesats
        Else
            frinetto = fri - (fri - avdrag) * skattesats
        End If
    End If

    If marginal = 0 Then frinetto = Int(frinetto)
    'Kan innebär en viss ekonomisk effekt när folk blir 66 år och får ett högre grundavdrag av att disp>fri...
    If disp < frinetto Then
        disp = frinetto
    End If
    If gift = 1 Then
        If Dispm < frinetto Then Dispm = frinetto 'Bör räknas på makens avdrag...
    End If
    'Lägg till för förmögenhet, BTP o dra ifrån hyran
    disp = disp + form - hyra + btpb / (gift + 1)
    Dispm = Dispm + form - hyra + btpb / (gift + 1)
    
    If disp < 0 Then disp = 0
    If Dispm < 0 Then Dispm = 0
    
    'Mellanskillnaden ...
    Dim sbtpm As Double
     
    SBTP = Lev - disp
    sbtpm = 0
    If gift = 1 Then sbtpm = Lev - Dispm
     
    
    If SBTP < 0 Then SBTP = 0
    If sbtpm < 0 Then sbtpm = 0


A_F_S:
    'Äldreförsörjningsstödet - ser till faktiska inkomster
    'If year>2019 then
        'Fribelopp för arbetsinkomster ska med.
    'end if
    
    Besk = ctxfvi - avdragxx(inkomst, pbb, marginal, bald, year, 21000, IBB, kvoten, Iyear, Xage) '-Avdrag
    
    If marginal = 0 Then
        disp = (ctxfvi - Int(Besk * skattesats)) + Int(form + kapital) * 0.7 + btpb / (gift + 1) + SBTP - hyra
    Else
        disp = ctxfvi - Besk * skattesats + form + kapital * 0.7 + btpb / (gift + 1) + SBTP - hyra
    End If

    If disp < 0 Then disp = 0
    
    Beskm = 0
    If gift = 1 And (Application.Range("rulesfromUtg") = 0 Or IsMissing(Application.Range("rulesfromUtg"))) Then
         Beskm = ctxfvim - avdragxx(inkomstm, pbb, marginal, baldm, year, 21000, IBB, kvoten, Iyear, Xage)
    Else
         Beskm = ctxfvim - avdragxx(inkomstm, pbb, marginal, baldm, Application.Range("rulesfromUtg"), 21000, IBB, kvoten, Iyear, Xage)
    End If

    
    Dispm = 0
    If gift = 1 Then
        Dispm = ctxfvim - Beskm * skattesats + form + kapital * 0.7 + btpb / (gift + 1) + sbtpm - hyra
        If marginal = 0 Then Dispm = ctxfvim - Int(Beskm * skattesats) + Int(form + kapital) * 0.7 + btpb / (gift + 1) + sbtpm - hyra
    End If
    If Dispm < 0 Then Dispm = 0
    
    'Disp = (Disp + Dispm) / (gift + 1)
    
    Dim AFS As Double
    Dim AFSm As Double
    
     AFS = (Lev - disp)
     If AFS < 0 Then AFS = 0
     AFSm = 0
     If gift = 1 Then AFSm = (Lev - Dispm)
     If AFSm < 0 Then AFSm = 0
     
     'SBTP = SBTP + AFS + sbtpm + AFSm
    If Application.Range("rng_dela") = 1 Then
         SBTP = SBTP + AFS + sbtpm + AFSm
    ElseIf Application.Range("rng_dela") = 2 Then
        SBTP = SBTP + AFS '+ sbtpm + AFSm
    Else
        SBTP = (SBTP + AFS + sbtpm + AFSm) / (gift + 1)
    End If
End Function
''
''Sub KOLL_SBTP()
''     Dim koll As Double
''     Dim inkomst As Double
''     inkomst = 0 'CLng(5350) * 12
''     Dim hyra As Double: hyra = 84000
''     Dim gift As Single: gift = 0
''     Dim BTP_ As Double
''     Dim SBTP_ As Double
''
''     Dim arb As Double: arb = 0
''     Dim TJP As Double: TJP = 0 '15888
''     Dim avdrag As Double
''     Dim age As Integer:  age = 66
''     Dim form As Double: form = 0
''     Dim ftid As Integer: ftid = 1
''     Dim marginal As Byte
''     marginal = 0
''     Dim klass As Integer
''     Dim prisbb As Double
''     prisbb = 47600
''     Dim kskatt As Double
''     kskatt = 0.32
''     Dim year As Integer
''     year = 2021
''     Dim ii As Integer
''
''     For ii = 20 To 25
''        ftid = ii
''        inkomst = (0.185 * 0.423 * prisbb * ftid / 17) / 12
''        ReDim year_(0 To 105) As Long
''        'inkomst = Worksheets("Oles tablå").Cells(37, year - 2001 + 1) + _
''        '    Worksheets("Oles tablå").Cells(38, year - 2001 + 1) + _
''        '    Worksheets("Oles tablå").Cells(39, year - 2001 + 1)
''
''        'prisbb = Worksheets("Några faktorer").Cells(year - 1994 + 1, 5)
''        'kskatt = Worksheets("Några faktorer").Cells(year - 1994 + 1, 10) / 100
''        'Hyra = Worksheets("Antaganden").Cells(12, year - 2001 + 2)
''
''        avdrag = avdragxx(12 * inkomst, prisbb, 0, age, year, 2100, 59300, 1, 0)
''        BTP_ = BTP(12 * inkomst, 0, hyra, gift, prisbb, 1, 1, form, arb, 0, marginal, 60000, year, 2100, 56900, 1, age, TJP, 0, 0)
''
''        SBTP_ = SBTP(12 * inkomst, hyra, gift, BTP_, avdrag, kskatt, 1, form, prisbb, 74400, year, 2110, 56900, 1, age, 0, 0, 0, -99, ftid)
''        koll = btp_sbtp(BTP_, SBTP_, marginal) / 12
''
''        Debug.Print year; inkomst; Round(BTP_ / 12, 0); Round(SBTP_ / 12, 0); koll; avdrag
''
''     Next ii
''
'' End Sub

Function btp_sbtp(ByVal BTP As Double, ByVal SBTP As Double, Optional marginal = 0, Optional ByVal year = 2016) As Double
     'Application.Volatile
     'Dim btp_sbtp As Double
     If pblnCloseOrSave Then Exit Function
     On Error GoTo errTag
     
''     If Application.Range("Rulesfromutg") = 0 Or IsMissing(Application.Range("Rulesfromutg")) Then
''             'Nothing
''        Else
''            If (year > Application.Range("Rulesfromutg") And Application.Range("Rules") = 0) _
''            Or Application.Range("Rules") = 1 Then
''                year = Application.Range("Rulesfromutg")
''            End If
''    End If
    
     
     btp_sbtp = (BTP + SBTP) '/ (gift + 1) 'sker senare i huvudmodellen
     If marginal = 0 Then btp_sbtp = Int(btp_sbtp / 12 + 0.5) * 12
     
      If marginal = 0 And year < 2014 And SBTP = 0 And (btp_sbtp < 25 * 12) Then btp_sbtp = 0
      Exit Function
errTag:
        btp_sbtp = CVErr(xlErrValue)

End Function
Function CalcAntalBarn(ByVal ar As Integer, ByVal barn1 As Integer, ByVal barn2 As Integer, ByVal barn3 As Integer, ByVal barn4 As Integer) As Single
        
    If ar >= barn1 And ar < (barn1 + 20) Then
        'If year_(age) = year(barn1) Then
        '    AntalBarn = (12 - month(barn1)) / 12
        'Else
        CalcAntalBarn = 1
        'End If
        If ar = (barn1 + 19) Then CalcAntalBarn = 0
    End If
    'Ev barn2
    If ar >= barn2 And ar < (barn2 + 20) Then
        If ar >= barn2 Then CalcAntalBarn = CalcAntalBarn + 1
        If ar = (barn2 + 19) Then CalcAntalBarn = maxi(CalcAntalBarn - 1, 0)
    End If
    'Ev barn3
    If ar >= barn3 And ar < (barn3 + 20) Then
        If ar >= barn3 Then CalcAntalBarn = CalcAntalBarn + 1
        If ar = (barn3 + 19) Then CalcAntalBarn = maxi(CalcAntalBarn - 1, 0)
    End If
    'Ev barn4
    If ar >= barn4 And ar < (barn3 + 20) Then
        If ar >= barn4 Then CalcAntalBarn = CalcAntalBarn + 1
        If ar = (barn4 + 19) Then CalcAntalBarn = maxi(CalcAntalBarn - 1, 0)
    End If

End Function
Sub CalcBarnPerAlder(ByRef b1 As Integer, ByRef b2 As Integer, ByRef b3 As Integer, ByRef b4 As Integer, ByRef b5 As Integer, ByRef b6 As Integer, ByRef b7 As Integer, ByRef b8 As Integer, _
ByVal ar As Integer, ByVal barn1 As Integer, ByVal barn2 As Integer, ByVal barn3 As Integer, ByVal barn4 As Integer)

'Barnfamiljer får barnbidrag ev, underhållsstöd och bostadsbidraget - dessa är förenklade, Ingen hänsyn till när barnet fyller år under året mm
    b1 = 0 '- Antal barn 0 år
    b2 = 0 '- Antal barn 1-2   år gamla
    b3 = 0 '- Antal barn 3 år gamla
    b4 = 0 '- Antal barn 4-6   år gamla
    b5 = 0 '- Antal barn 7-10  år gamla
    b6 = 0 '- Antal barn 11-14 år gamla
    b7 = 0 '- Antal barn 15-18 år gamla
    b8 = 0 '- Antal barn 19-20 år gamla
    'barn1-4 sorterade

    'barn1
    If ar >= barn1 And ar <= (barn1 + 20) Then
        If ar = barn1 Then
            b1 = 1
        ElseIf ar < barn1 + 3 Then
            b2 = 1
        ElseIf ar < barn1 + 4 Then
            b3 = 1
        ElseIf ar < barn1 + 7 Then
            b4 = 1
        ElseIf ar < barn1 + 11 Then
            b5 = 1
        ElseIf ar < barn1 + 15 Then
            b6 = 1
        ElseIf ar < barn1 + 19 Then
            b7 = 1
        ElseIf ar < barn1 + 21 Then
            b8 = 1
        Else
        End If
    End If
    'Ev barn2
    If ar >= (barn2) And ar <= (barn2 + 20) Then
        If ar = barn2 Then
            b1 = b1 + 1
        ElseIf ar < barn2 + 3 Then
            b2 = b2 + 1
        ElseIf ar < barn2 + 4 Then
            b3 = b3 + 1
        ElseIf ar < barn2 + 7 Then
            b4 = b4 + 1
        ElseIf ar < barn2 + 11 Then
            b5 = b5 + 1
        ElseIf ar < barn2 + 15 Then
            b6 = b6 + 1
        ElseIf ar < barn2 + 19 Then
            b7 = b7 + 1
        ElseIf ar < barn2 + 21 Then
            b8 = b8 + 1
        Else
        End If
    End If
    'Ev barn3
    If ar >= (barn3) And ar <= (barn3 + 20) Then
        If ar = barn3 Then
            b1 = b1 + 1
        ElseIf ar < barn3 + 3 Then
            b2 = b2 + 1
        ElseIf ar < barn3 + 4 Then
            b3 = b3 + 1
        ElseIf ar < barn3 + 7 Then
            b4 = b4 + 1
        ElseIf ar < barn3 + 11 Then
            b5 = b5 + 1
        ElseIf ar < barn3 + 15 Then
            b6 = b6 + 1
        ElseIf ar < barn3 + 19 Then
            b7 = b7 + 1
        ElseIf ar < barn3 + 21 Then
            b8 = b8 + 1
        Else
        End If
    End If
    'Ev barn4
    If ar >= (barn4) And ar <= (barn4 + 20) Then
        If ar = barn4 Then
            b1 = b1 + 1
        ElseIf ar < barn4 + 3 Then
            b2 = b2 + 1
        ElseIf ar < barn4 + 4 Then
            b3 = b3 + 1
        ElseIf ar < barn4 + 7 Then
            b4 = b4 + 1
        ElseIf ar < barn4 + 11 Then
            b5 = b5 + 1
        ElseIf ar < barn4 + 15 Then
            b6 = b6 + 1
        ElseIf ar < barn4 + 19 Then
            b7 = b7 + 1
        ElseIf ar < barn4 + 21 Then
            b8 = b8 + 1
        Else
        End If
    End If

End Sub

Function barnbidraget(ByVal antal As Integer, Optional ByVal year = 2013) As Double
     'Denna funktion behöver ses över historiskt - Adoptionsbidrag är inte med
     'Infördes 1937 och allmänt 1948 (260 kr/barn och år
     
     If pblnCloseOrSave Then Exit Function
     On Error GoTo errTag
     'Application.Volatile
     
''     If Application.Range("Rulesfromutg") = 0 Or IsMissing(Application.Range("Rulesfromutg")) Then
''             'Nothing
''        Else
''            If (year > Application.Range("Rulesfromutg") And Application.Range("Rules") = 0) _
''            Or Application.Range("Rules") = 1 Then
''                year = Application.Range("Rulesfromutg")
''            End If
''    End If
     
     
     Dim xtill() As Double
     
''     If antal > 11 Then
''       MsgBox ("Antal barn>11? Max i beräkningarna är 11, handpåläggning ?")
''       antal = 11
''     End If
     
     Dim i As Integer
     If antal > 11 Then
        ReDim xtill(1 To antal) As Double
     Else
        ReDim xtill(1 To 11) As Double
     End If
     Dim grund As Double
     Dim fbtill As Double
     
    If antal > 0 Then
        For i = 1 To antal
            xtill(i) = 0
        Next i
    End If
        
    Select Case year
    Case Is < 1974 '1974-01
        grund = CDbl(330 * 4)
    Case Is = 1974 '1974-04
        grund = CDbl(375 * 4) + 200
    Case Is = 1975 '1975-10
        grund = CDbl(375 * 3 + 450)
    Case Is = 1976  '1977-07
        grund = CDbl(4 * 450)
    Case Is = 1977 '1977-07
        grund = CDbl(2 * 450 + 2 * 525)
    Case Is = 1978  '1978-04
        grund = CDbl(1 * 525 + 3 * 565)
    Case Is = 1979  '1979-01
        grund = CDbl(4 * 625)
    Case Is = 1980  '1980-01 och 1980-10
        grund = CDbl(3 * 700 + 750)
    Case Is = 1981
        grund = CDbl(750 * 4)
    Case Is = 1982   '1982-01
        grund = CDbl(750 * 4)
        xtill(3) = 750
        xtill(4) = 1500
        If antal > 4 Then
            For i = 5 To antal
                xtill(i) = xtill(i - 1) + 1500
            Next i
        End If
    Case Is < 1985 '1983 - 1
        grund = CDbl(825 * 4)
        xtill(3) = 825
        xtill(4) = 1650
        If antal > 4 Then
            For i = 5 To antal
                xtill(i) = xtill(i - 1) + 1650
            Next i
        End If
    Case Is < 1987 '1985-01
        grund = 4800
        xtill(3) = 1200
        xtill(4) = 2400
        If antal > 4 Then
            For i = 5 To antal
                xtill(i) = xtill(i - 1) + 2400
            Next i
        End If
        'Förlängt barnbidrag och vissa adoptivbarn införs 1985
    Case Is < 1988 '1987-01
        grund = 5820
        xtill(3) = 1455
        If antal > 3 Then
            For i = 4 To antal
                xtill(i) = xtill(i - 1) + 2890
            Next i
        End If
    Case Is < 1989 '1988-01
        grund = 5820
        xtill(3) = 1455
        xtill(4) = 1.6 * grund
        If antal > 4 Then
            For i = 5 To antal
                xtill(i) = 1.6 * grund
            Next i
        End If
    Case Is < 1990 '1989-01
        grund = 5820
        xtill(3) = 1455
        xtill(4) = 1.9 * grund
        xtill(5) = 2.4 * grund
        If antal > 5 Then
            For i = 6 To antal
                xtill(i) = 1.6 * grund
            Next i
        End If
    Case Is < 1991 '1990-01
        grund = 6720
        xtill(3) = 1455
        xtill(4) = 1.9 * grund
        xtill(5) = 2.4 * grund
        If antal > 5 Then
            For i = 6 To antal
                xtill(i) = 1.6 * grund
            Next i
        End If
    Case Is < 1994 '1991-01
        grund = CDbl(750 * 12)
        xtill(3) = 1455
        xtill(4) = 1# * grund
        xtill(5) = 1.5 * grund
        If antal > 5 Then
            For i = 6 To antal
                xtill(i) = 1.5 * grund
            Next i
        End If
    Case Is = 1994 '1994-07
        grund = CDbl(750 * 12)
        xtill(3) = 1455
        xtill(4) = 1# * grund
        xtill(5) = 1.25 * grund
        If antal > 5 Then
            For i = 6 To antal
                xtill(i) = 1.25 * grund 'Halvårs effekt
            Next i
        End If
   Case Is < 1996 '1995-01
        grund = CDbl(750 * 12)
        xtill(3) = 2400
        xtill(4) = 7200
        If antal > 4 Then
            For i = 5 To antal
                xtill(i) = 9000
            Next i
        End If
   Case Is < 1998 '1996-01
        grund = CDbl(640 * 12)
        'Inga flerbarnstilllägget avskaffas för barn födda efter 1995
        'if year(barn)>1995 then antal=antal-1 men en dagsslända...
        xtill(3) = 2400
        xtill(4) = 7200
        If antal > 4 Then
            For i = 5 To antal
                xtill(i) = 9000
            Next i
        End If
    Case Is < 2000 '1998-01
        grund = CDbl(750 * 12)
        'Flerbarnstilllägget återgår
        xtill(3) = 2400
        xtill(4) = 7200
        If antal > 4 Then
            For i = 5 To antal
                xtill(i) = 9000
            Next i
        End If
     Case Is < 2001 '2000-01
        grund = CDbl(850 * 12)
        xtill(1) = 0
        xtill(2) = 0
        xtill(3) = CDbl(227 * 12)
        xtill(4) = CDbl(680 * 12)
        xtill(5) = CDbl(850 * 12)
        If antal > 5 Then
            For i = 6 To antal
                xtill(i) = xtill(5)
            Next i
        End If
    Case Is < 2006 '2001-01
        grund = CDbl(950 * 12)
        xtill(1) = 0
        xtill(2) = 0
        xtill(3) = CDbl(254 * 12)
        xtill(4) = CDbl(760 * 12)
        xtill(5) = CDbl(950 * 12)
        If antal > 5 Then
            For i = 6 To antal
                xtill(i) = xtill(5)
            Next i
        End If
    Case Is < 2011 '
        grund = 12600
        xtill(1) = 0
        xtill(2) = 1200
        xtill(3) = 4248
        xtill(4) = 10320
        xtill(5) = 12600
        If antal > 5 Then
            For i = 6 To antal
                xtill(i) = xtill(5)
            Next i
        End If
    Case Is <= 2016 '=BP för 2017
        grund = CDbl(1050 * 12)
        xtill(1) = 0
        xtill(2) = CDbl(150 * 12)
        xtill(3) = CDbl(454 * 12)
        xtill(4) = CDbl(1010 * 12)
        xtill(5) = CDbl(1250 * 12)
        If antal > 5 Then
            For i = 6 To antal
                xtill(i) = xtill(5)
            Next i
        End If
    Case Else '2017-
        grund = CDbl(1050 * 12)
        xtill(1) = 0
        xtill(2) = CDbl(150 * 12)
        xtill(3) = CDbl(580 * 12)
        xtill(4) = CDbl(1010 * 12)
        xtill(5) = CDbl(1250 * 12)
        If antal > 5 Then
            For i = 6 To antal
                xtill(i) = xtill(5)
            Next i
        End If
    End Select
     
    If year = 2018 Then
        grund = CDbl(1050 * 3 + 1250 * 9)
    End If
    If year > 2018 Then
        grund = CDbl(1250 * 12)
    End If
    
    grund = antal * grund
    fbtill = 0
    If antal > 1 Then
       For i = 2 To antal
          'j = Application.Min((i - 1), 2)
          fbtill = fbtill + xtill(i)
       Next
    End If
    
    Dim adoptionsbidrag As Byte
    adoptionsbidrag = 0 'Inget bidrag antas
    If year > 2016 Then
       barnbidraget = grund + fbtill + adoptionsbidrag * 40000
    Else
       barnbidraget = grund + fbtill + adoptionsbidrag * 75000
    End If
    
    Exit Function
errTag:
        barnbidraget = CVErr(xlErrValue)
End Function

''Sub barnbidrag()
''    Dim koll As Double
''    Dim antal As Integer
''    antal = 2
''    Dim i As Integer, year As Integer
''    year = 2013
''    For i = 1 To 1
''        koll = barnbidraget(antal, year)
''        Debug.Print i; year; antal; koll
''        year = year + 1
''    Next i
''
''End Sub

'Bostadsbidrag åt barnfamiljer
Function bobid(vuxna, inkomst, make_ink, barn, Uboende, Optional yta = 80, Optional marginal = 0, Optional ungdom = 0, Optional ByVal year = 2013) As Double
    'Historiska gränsvärden behöver ses över
     If pblnCloseOrSave Then Exit Function
     On Error GoTo errTag
     'Application.Volatile
      'UNGDOMAR SAKNAS 24-28§ samt UMGÄNGESBIDRAG 23§
      
''      If Application.Range("Rulesfromutg") = 0 Or IsMissing(Application.Range("Rulesfromutg")) Then
''             'Nothing
''        Else
''            If (year > Application.Range("Rulesfromutg") And Application.Range("Rules") = 0) _
''            Or Application.Range("Rules") = 1 Then
''                year = Application.Range("Rulesfromutg")
''            End If
''        End If
      
        'Boendekostnadsgränser
        'FAMILJ MED 1 BARN
        Dim xn1 As Integer
        Dim xm1 As Integer
        Dim xo1 As Integer
        'FAMILJ MED 2 BARN
        Dim xn2 As Integer
        Dim xm2 As Integer
        Dim xo2 As Integer
        'FAMILJ MED 3 BARN
        Dim xn3 As Integer
        Dim xm3 As Integer
        Dim xo3 As Integer
        
         'Särskilt bidrag
        Dim xg1 As Integer
        Dim xg2 As Integer
        Dim xg3 As Integer
        Dim xg4 As Integer
        Dim xg5 As Integer
        
        Dim bhyra1 As Integer
        Dim bhyra2 As Integer
        Dim bhyra3 As Integer
        Dim bhyra4 As Integer
        Dim bhyra5 As Integer
            
        Dim xfmb As Double
        Dim xrfmb As Double
        
        'REDUCERINGEN
        xrfmb = 0.2   'REDFAKTOR
                
        If year < 1995 Then '1995-01
            xn1 = 3200  '1 barn
            xm1 = 4800
            xo1 = xm2
            xn2 = 3200  '2
            xm2 = 5300
            xo2 = xm2
            xn3 = 5300  '3
            xm3 = 6100
            xo3 = xm3
            xg1 = 600  'FAST BELOPP FAMILJ MED 1 BARN
            xg2 = 900  'FAST BELOPP      "     2 BARN
            xg3 = 1200 'FAST BELOPP      "     2 BARN
            bhyra1 = 2900
            bhyra2 = 3200
            bhyra3 = 3500
            bhyra4 = 3800
            bhyra5 = 4100
            xfmb = 110000 'MININKOMST
        ElseIf year < 1996 Then '1996-01
            xn1 = 3300
            xm1 = 5200
            xo1 = xm2
            xn2 = 3300
            xm2 = 5800
            xo2 = xm2
            xn3 = 5400
            xm3 = 6500
            xo3 = xm3
            xg1 = 600
            xg2 = 900
            xg3 = 1200
            bhyra1 = 3000
            bhyra2 = 3300
            bhyra3 = 3600
            bhyra4 = 3900
            bhyra5 = 4200
            xfmb = 115000 'MININKOMST
        ElseIf year < 2017 Then '
            xn1 = 2000
            xm1 = 3000
            xo1 = 5300
            xn2 = 2000
            xm2 = 3300
            xo2 = 5900
            xn3 = 2000
            xm3 = 3600
            xo3 = 6600
            xg1 = 600
            xg2 = 900
            xg3 = 1200
            bhyra1 = 3000
            bhyra2 = 3300
            bhyra3 = 3600
            bhyra4 = 3900
            bhyra5 = 4200
            xfmb = 117000 'MININKOMST
        Else
            xn1 = 2000
            xm1 = 5300
            xo1 = xm1
            xn2 = 2000
            xm2 = 5900
            xo2 = xm2
            xn3 = 2000
            xm3 = 6600
            xo3 = xm3
            xg1 = 1300
            xg2 = 1750
            xg3 = 2350
            bhyra1 = 3000
            bhyra2 = 3300
            bhyra3 = 3600
            bhyra4 = 3900
            bhyra5 = 4200
            xfmb = 117500 'MININKOMST individuelt
        End If
        
        xg4 = xg3
        xg5 = xg3
        If year >= 2017 Then xfmb = 127000
         
        If year > 2012 Then
            xn1 = 1400
            xn2 = 1400
            xn3 = 1400
        End If
       
    If Uboende > 50000 Then Uboende = Uboende / 12 'Årshyran?
    If year > 2017 Then 'förändring 1 mars...
           xfmb = 135000
           If year = 2019 Then xfmb = 142000
           If year = 2020 Then xfmb = 148000
           If year > 2020 Then xfmb = 150000
    End If
      
    Dim boyta1 As Integer
    Dim boyta2 As Integer
    Dim boyta3 As Integer
    Dim boyta4 As Integer
    Dim boyta5 As Integer
    boyta1 = 80     'Max yta FAMILJ MED 1 BARN
    boyta2 = 100    'Max yta FAMILJ MED 2 BARN
    boyta3 = 120    'Max yta FAMILJ MED 3 BARN
    boyta4 = 140    'Max yta FAMILJ MED 4 BARN
    boyta5 = 160    'Max yta FAMILJ MED 5 BARN eller fler
    
        
    Dim zbost As Double
    zbost = Uboende
    
    If barn = 1 And yta > boyta1 Then
        If Uboende > bhyra1 Then zbost = (Uboende - bhyra1) * boyta1 / yta
    End If
    If barn = 2 And yta > boyta2 Then
        If Uboende > bhyra2 Then zbost = (Uboende - bhyra2) * boyta2 / yta
    End If
    If barn = 3 And yta > boyta3 Then
        If Uboende > bhyra3 Then zbost = (Uboende - bhyra3) * boyta3 / yta
    End If
    If barn = 4 And yta > boyta3 Then
        If Uboende > bhyra4 Then zbost = (Uboende - bhyra4) * boyta3 / yta
    End If
    If barn > 4 And yta > boyta3 Then
        If Uboende > bhyra5 Then zbost = (Uboende - bhyra5) * boyta3 / yta
    End If
    
    Dim xand1 As Double
    Dim xand2 As Double
    If year <= 2004 Then
        xand1 = 0.5   'PROCENTANDEL MELLAN - NEDRE GRÄNS
        xand2 = 0.5   'PROCENTANDEL ÖVRE - MELLAN GRÄNS
    ElseIf year <= 2012 Then
        xand1 = 0.75   'PROCENTANDEL MELLAN - NEDRE GRÄNS
        xand2 = 0.5   'PROCENTANDEL ÖVRE - MELLAN GRÄNS
    Else
        xand1 = 0.5   'PROCENTANDEL MELLAN - NEDRE GRÄNS
        xand2 = 0.5   'PROCENTANDEL ÖVRE - MELLAN GRÄNS
    End If
    
      
      'Lagra vissa värden i vektorer
      Dim ZN As Variant
      Dim ZM As Variant
      Dim ZO As Variant
      Dim ZG As Variant
      
      ZN = Array(0, xn1, xn2, xn3) 'Nedre
      ZM = Array(0, xm1, xm2, xm3) 'Gräns
      ZO = Array(0, xo1, xo2, xo3) 'Max
      ZG = Array(0, xg1, xg2, xg3, xg4, xg5) 'Generellt
     
     Dim ibostbh As Double
     ibostbh = 0     'NOLLSTÄLLNING AV BOSTADSBIDRAG
    
    Dim csbink As Double
    Dim csbinkM As Double
    
      'BIDRAGSGRUNDANDE INKOMST
      If vuxna = 1 Then
        csbink = inkomst
        csbinkM = 0
      Else
        csbink = inkomst
        csbinkM = make_ink
      End If
    
    Dim antbarn As Double
    Dim hemmabarn As Double
      antbarn = 0
      hemmabarn = barn
      
    Dim zbarnsum As Double
    Dim zbantbrn As Double
    Dim zbantsar As Double
    
      zbarnsum = hemmabarn + antbarn
      zbantbrn = Application.min(3, zbarnsum)  'ANTAL BARN VID FASTSTÄLLANDE AV HYRESGRÄNSER
      zbantsar = Application.min(3, hemmabarn) 'ANTAL BARN FÖR BERÄKNING AV DET SÄRSKILDA BIDRAGET
    
                           'MÅNADSHYRA
      If marginal = 0 Then zbost = 25 * Int(zbost / 25)       '25-KRONORSAVRUNDNING AV MÅNADSHYRAN
    
      If zbarnsum > 0 Then
            'If vuxna = 1 Then XFAM = 1 Else XFAM = 2
            If zbost <= ZN(zbantbrn) Then                         'MINDRE ÄN NEDRE HYRESGRÄNS
              ibostbh = 12 * ZG(zbantsar)
            ElseIf zbost <= ZM(zbantbrn) Then                     'MINDRE ÄN MELLAN HYRESGRÄNS
              ibostbh = 12 * (ZG(zbantsar) + _
                        (zbost - ZN(zbantbrn)) * xand1)
            ElseIf zbost <= ZO(zbantbrn) Then                     'MINDRE ÄN ÖVRE HYRESGRÄNS
              ibostbh = 12 * (ZG(zbantsar) + _
                        (ZM(zbantbrn) - ZN(zbantbrn)) * xand1 + _
                        (zbost - ZM(zbantbrn)) * xand2)
            Else                                                  'STÖRRE ÄN ÖVRE HYRESGRÄNS
                 ibostbh = 12 * (ZG(zbantsar) + _
                         (ZM(zbantbrn) - ZN(zbantbrn)) * xand1 + _
                         (ZO(zbantbrn) - ZM(zbantbrn)) * xand2)
            End If
      End If
      
        'R E D U C E R I N G E N
      If vuxna = 1 Then                'ENSAMSTÅENDE
        If csbink > xfmb And ibostbh > 0 Then
          ibostbh = ibostbh - (xrfmb * (csbink - xfmb))
        End If
      ElseIf vuxna = 2 Then           'GIFTA/SAMBO
         If csbink > xfmb / 2 And ibostbh > 0 Then
           ibostbh = ibostbh - (xrfmb * (csbink - xfmb / 2))
         End If
         If csbinkM > xfmb / 2 And ibostbh > 0 Then
           ibostbh = ibostbh - (xrfmb * (csbinkM - xfmb / 2))
         End If
      End If
      
      If ibostbh < 0 Then ibostbh = 0
    
      If ibostbh > 12 * Uboende Then ibostbh = Uboende * 12 'EJ STÖRRE BOSTADSBIDRAG ÄN ÅRSHYRAN
      
      If (ibostbh > 0 And year = 2020) Then ibostbh = ibostbh + ibostbh * 0.25 * (6 / 12) 'Förslag maj 2020 antas gå igenom
      
      If marginal = 0 And ibostbh < 1200 Then ibostbh = 0                  'MINST 100 KR PER MÅNAD
      ibostbh = ibostbh / 12
      If marginal = 0 Then ibostbh = Int(ibostbh / 100) * 100
      bobid = 12 * ibostbh
      
        Exit Function
errTag:
        bobid = CVErr(xlErrValue)
End Function

''Sub bobidkoll()
''    'bobid(vuxna, inkomst, make_ink, barn, uboende, Optional yta = 80, Optional marginal = 0, Optional ungdom = 0, Optional year = 2013) As Double
''    Dim koll As Double
''    koll = bobid(1, 100000, 0, 1, 60000, 80, 0, 0, 2005)
''End Sub

Function ustod(ByVal barn As Integer, Optional ByVal ensamst = 1, Optional ByVal year = 2013) As Double
    'Ses över historiskt - 1996-12 Lagen träder ikraft och ersätter tidigare...
    'Växelvisboende beaktas inte, den tar sedan ett tag tillbaka även hänsyn till barnets ålder
    
     If pblnCloseOrSave Then Exit Function
'     On Error GoTo ErrTag
     'Application.Volatile
      ustod = 0
      
''     If Application.Range("Rulesfromutg") = 0 Or IsMissing(Application.Range("Rulesfromutg")) Then
''             'Nothing
''        Else
''            If (year > Application.Range("Rulesfromutg") And Application.Range("Rules") = 0) _
''            Or Application.Range("Rules") = 1 Then
''                year = Application.Range("Rulesfromutg")
''            End If
''    End If
      
     If ensamst <> 1 Then Exit Function
     
     If year < 1996 Then '1996-12
        ustod = CDbl(1173 * 12)
     ElseIf year < 2006 Then '2006-02
          ustod = CDbl(1173 * 12)
     ElseIf year < 2015 Then '2015-09
        ustod = CDbl(1273 * 12)
     ElseIf year < 2017 Then
        ustod = CDbl(1573 * 12)
     ElseIf year < 2022 Then 'Saknas ett år mellan och beror på barnets ålder här för 7-14
        ustod = CDbl(1673 * 12)
     Else
           ustod = CDbl(1823 * 12)
     End If
    
     ustod = ustod * barn
    
'ErrTag:
'        ustod = CVErr(xlErrValue)
End Function

''Sub kollusupp()
''    'ustod(ByVal barn As Integer, Optional ByVal ensamst = 1, Optional ByVal year = 2013) As Double
''    Dim barn As Integer
''    Dim year As Integer
''    Dim koll As Double
''    barn = 1
''    For year = 1990 To 2017
''        koll = ustod(barn, 1, year)
''        Debug.Print year; koll
''    Next year
''
''End Sub

Function dagis(ByVal inkomst As Double, ByVal barn As Integer, Optional ByVal year = 2013) As Variant
 'Ses över historiskt gammal
    ' Barnet antas vara i försloleåldern om ett barn
    ' Barnen antas vara i försloleåldern om två barn
    ' 2 Barn antas vara i försloleåldern och 1 barn i fritids om tre barn eller fler
    If pblnCloseOrSave Then Exit Function
    On Error GoTo errTag
     'Application.Volatile
''
''     If Application.Range("Rulesfromutg") = 0 Or IsMissing(Application.Range("Rulesfromutg")) Then
''             'Nothing
''        Else
''            If (year > Application.Range("Rulesfromutg") And Application.Range("Rules") = 0) _
''            Or Application.Range("Rules") = 1 Then
''                year = Application.Range("Rulesfromutg")
''            End If
''    End If
     
    Dim avgift As Double
    If barn > 0 Then
      avgift = Application.min(1260, 0.03 * inkomst)
      If barn > 1 Then
        avgift = avgift + Application.min(840, 0.02 * inkomst)
      End If
      If barn > 2 Then
        avgift = avgift + Application.min(420, 0.01 * inkomst)
      End If
     End If
     'ÅRsavgiften = 11* avgift
     'Antar att avgiften betalas 11 månader per år
     dagis = 11 * avgift
        Exit Function
errTag:
        dagis = CVErr(xlErrValue)
End Function

Function bistOld(ByVal civ As Integer, ByVal hyra As Double, ByVal disp As Double, _
    ByVal b1 As Integer, ByVal b2 As Integer, ByVal b3 As Integer, ByVal b4 As Integer, ByVal b5 As Integer, _
    ByVal b6 As Integer, ByVal b7 As Integer, ByVal b8 As Integer, ByVal wage As Double, Optional year = 2005) As Double
    ' Socialstyrelsens riksnorm för försörjningsstöd inkomståren 1985-2005
    
    'Civ - Civilstånd (antal vuxna)
    'Hyra - Skälig boendekostand
    'Disp - Hushållets disponibla inkomst
    'B1 - Antal barn 0 år
    'B2 - Antal barn 1-2   år gamla
    'B3 - Antal barn 3 år gamla
    'B4 - Antal barn 4-6   år gamla
    'B5 - Antal barn 7-10  år gamla
    'B6 - Antal barn 11-14 år gamla
    'B7 - Antal barn 15-18 år gamla
    'B8 - Antal barn 19-20 år gamla
    'Wage - Nettolön obs 0 om bistånd inte har föregåtts om sex månader i anropet
    'year - inkomstår
    
                    
    bistOld = 0
    If year < 1985 Or year > 2005 Then Exit Function
    Const Verbose_ = 0 'koll av värdena 1-Personliga, 2 Vuxna 3 hushåll 9 alla
        
'--------Personliga kostnader för hemmavarande barn utan lunch en funktion av barnets ålder per månad
    Dim xn() As Variant 'as double 'OBS option BASE 1
    
    'ReDim xn(1 To 8)
'    xn(1) = 2720 '0 år
'    xn(2) = 3030 '1-2
'    xn(3) = 2700 '3
'    xn(4) = 3030 '4-6
'    xn(5) = 3790 '7-10
'    xn(6) = 4350 '11-14
'    xn(7) = 4920 '15-18
'    xn(8) = 4960 '19-20
                    '0,   1-2,    3,  4-6, 7-10,11-14,15-18,19-20
   Dim i As Integer
   
    If year = 1985 Then
                    '0,  1-2,    3,  4-6, 7-10,11-14,15-18,19-20
       xn = Array(1000, 1000, 1000, 1181, 1181, 1362, 1362, 1362)
    ElseIf year = 1986 Then
       xn = Array(1070, 1070, 1070, 1260, 1265, 1455, 1455, 1455)
    ElseIf year = 1987 Then
       xn = Array(1105, 1105, 1105, 1305, 1305, 1505, 1505, 1505)
    ElseIf year = 1988 Then
       xn = Array(1180, 1180, 1180, 1395, 1395, 1615, 1615, 1615)
    ElseIf year = 1989 Then
       xn = Array(1280, 1280, 1280, 1510, 1510, 1745, 1745, 1745)
    ElseIf year = 1990 Then
       xn = Array(1360, 1360, 1360, 1610, 1610, 1855, 1855, 1855)
    ElseIf year = 1991 Then
       xn = Array(1504, 1504, 1504, 1770, 1770, 2038, 2038, 2038)
    ElseIf year = 1992 Then
       xn = Array(1573, 1573, 1573, 1853, 1853, 2134, 2134, 2134)
    ElseIf year = 1993 Then
       xn = Array(1606, 1606, 1606, 1892, 1892, 2180, 2180, 2180)
    ElseIf year = 1994 Then
       xn = Array(1643, 1643, 1643, 1936, 1936, 2229, 2229, 2229) 'ta bort läkar och tandvård 32 kr!
       For i = 1 To 8
            xn(i) = xn(i) - 32
       Next i
    ElseIf year = 1995 Then
       xn = Array(1666, 1666, 1666, 1964, 1964, 2261, 2261, 2261) '-33
        For i = 1 To 8
            xn(i) = xn(i) - 33
       Next i
    ElseIf year = 1996 Then
       xn = Array(1666 * 9 / 12 + 1320 * 3 / 12, 1666 * 9 / 12 + 1460 * 3 / 12, 1666 * 9 / 12 + 1090 * 3 / 12, _
                  1964 * 9 / 12 + 1370 * 3 / 12, 1964 * 9 / 12 + 1660 * 3 / 12, _
                  2261 * 9 / 12 + 1920 * 3 / 12, 2261 * 9 / 12 + 2180 * 3 / 12, _
                  2261 * 9 / 12 + 2190 * 3 / 12) ' -33*9/12
       For i = 1 To 8
            xn(i) = xn(i) - 33 * 9 / 12
       Next i
     
    ElseIf year = 1997 Then
       xn = Array(1320, 1460, 1090, 1370, 1660, 1920, 2180, 2190)
    ElseIf year = 1998 Then
       xn = Array(1226, 1438, 1134, 1531, 1833, 2070, 2321, 2321)
    ElseIf year = 1999 Then
       xn = Array(1230, 1440, 1120, 1410, 1530, 1830, 2070, 2320)
    ElseIf year = 2000 Then
       xn = Array(1220, 1470, 1160, 1440, 1590, 1890, 2140, 2400)
    ElseIf year = 2001 Then
       xn = Array(1220, 1470, 1160, 1440, 1600, 1890, 2130, 2400)
    ElseIf year = 2002 Then
       xn = Array(1360, 1610, 1280, 1610, 1770, 2050, 2320, 2520)
    ElseIf year = 2003 Then
       xn = Array(1405, 1625, 1305, 1635, 1815, 2090, 2360, 2575)
    ElseIf year = 2004 Then
       xn = Array(1470, 1670, 1350, 1680, 1880, 2160, 2440, 2650)
       
    Else '2005
        xn = Array(1440, 1640, 1330, 1630, 1840, 2120, 2400, 2590)
    
    End If
     
    Dim bistand As Double
    bistand = 0
    Dim barn(1 To 8) As Integer
    barn(1) = b1
    barn(2) = b2
    barn(3) = b3
    barn(4) = b4
    barn(5) = b5
    barn(6) = b6
    barn(7) = b7
    barn(8) = b8
   
    Dim antal As Integer
    antal = 0 'Antal barn
    For i = 1 To 8
            bistand = bistand + xn(i) * barn(i)
            antal = antal + barn(i)
            
            If (Verbose_ = 1 Or Verbose_ = 9) Then Debug.Print year, "X(" & i & ")"; year; xn(i)
            
    Next i
    
    antal = antal + civ 'Antal familjemdedlemmar
    
'---------- Personlig kostnader vuxna
    Dim vuxna() As Variant
    'Dim tal1, tal2 As Double
    
    If year = 1985 Then
        vuxna = Array(2089, 3452)
    ElseIf year = 1986 Then
        vuxna = Array(2230, 3690)
    ElseIf year = 1987 Then
        vuxna = Array(2310, 3815)
    ElseIf year = 1988 Then
        vuxna = Array(2475, 4085)
    ElseIf year = 1989 Then
        vuxna = Array(2675, 4420)
    ElseIf year = 1990 Then
        vuxna = Array(2845, 4705)
    ElseIf year = 1991 Then
        vuxna = Array(3114, 5160)
    ElseIf year = 1992 Then
        vuxna = Array(3258, 5392)
    ElseIf year = 1993 Then
        vuxna = Array(3326, 5505)
    ElseIf year = 1994 Then
        vuxna = Array(3403, 5632) 'Ta bort läkar och tandård 56 och 114
    ElseIf year = 1995 Then
        vuxna = Array(3451 - 57, 5712 - 116) 'Läkar- och tandvård -57 -116
    ElseIf year = 1996 Then
        vuxna = Array(3451 * 9 / 12 + 3640 * 3 / 12, _
                      CDbl(5712) * 9 / 12 + 4150 * 3 / 12) '(-57 -116) *9/12
        vuxna(1) = vuxna(1) - 57 * 9 / 12
        vuxna(2) = vuxna(2) - 116 * 9 / 12
    ElseIf year = 1997 Then
        vuxna = Array(2190, 4150) '+100 en personhushåll?
    ElseIf year = 1998 Then
        vuxna = Array(2321, 4208)
    ElseIf year = 1999 Then
        vuxna = Array(2320, 4200)
    ElseIf year = 2000 Then
        vuxna = Array(2400, 4360)
    ElseIf year = 2001 Then
        vuxna = Array(2400, 4370)
    ElseIf year = 2002 Then
        vuxna = Array(2520, 4570)
    ElseIf year = 2003 Then
        vuxna = Array(2575, 4685)
    ElseIf year = 2004 Then
        vuxna = Array(2650, 4840)
             
    Else '2005
         vuxna = Array(2590, 4720)
    End If
    
    If (Verbose_ = 2 Or Verbose_ = 9) Then Debug.Print year, "Civ 1 och 2", vuxna(1), vuxna(2)
    
    'Bistand + de vuxnas och och boendekostnad
    If civ = 1 Then
        bistand = bistand + vuxna(1)
        hyra = hyra 'Antas skälig
    Else
        bistand = bistand + vuxna(2)
        hyra = hyra  'Antas skälig
    End If
    
'------Gemensamma hushållskostnader
    Dim Gn() As Variant 'Double
    'ReDim Gn(1 To 7)

    If year <= 1995 Then
        Gn = Array(0, 0, 0, 0, 0, 0, 0)
    ElseIf year = 1996 Then
        Gn = Array(760 * 3 / 12, 920 * 3 / 12, 1060 * 3 / 12, 1170 * 3 / 12, _
                  1310 * 3 / 12, 1400 * 3 / 12, 1510 * 3 / 12) 'Anta mellanstorstad, ej storstad eller mindre tätort
    ElseIf year = 1997 Then
        Gn = Array(760, 920, 1060, 1170, 1310, 1400, 1510)      'Anta mellanstorstad, ej storstad eller mindre tätort
    ElseIf year = 1998 Then
        Gn = Array(563, 645, 733, 796, 882, 934, 993)
    ElseIf year = 1999 Then
        Gn = Array(580, 670, 760, 820, 910, 960, 1020)
    ElseIf year = 2000 Then
        Gn = Array(600, 680, 780, 840, 940, 1000, 1060)
    ElseIf year = 2001 Then
        Gn = Array(600, 680, 770, 840, 930, 980, 1040)
    ElseIf year = 2002 Then
        Gn = Array(620, 710, 810, 870, 970, 1030, 1090)
    ElseIf year = 2003 Then
        Gn = Array(680, 770, 890, 950, 1050, 1125, 1195)
    ElseIf year = 2004 Then
        Gn = Array(720, 800, 970, 1040, 1130, 1230, 1320)
        
    Else '2005
        Gn = Array(770, 870, 1030, 1100, 1190, 1290, 1360)
    End If
    
    If (Verbose_ = 3 Or Verbose_ = 9) Then
        For i = 1 To 7
            Debug.Print year; "G(" & i & ")", Gn(i)
        Next i
    End If
    
    '------
    Dim flera
    flera = Gn(7) - Gn(6)
    
    If antal > 7 Then
        bistand = bistand + Gn(7) + flera * (antal - 7)
    Else
        bistand = bistand + Gn(antal)
    End If
    'Maximalt stöd som reduceras med inkomsten men
    Dim just As Double
    just = 1 'Antas förändras med KPI vid framskrivning
    
    bistand = bistand * just + hyra '+barnomsorg uppräknade till 2025 med antagen prisutv. samma föa alla delkomponenter
    
    'if marginal=0 then
    bistand = Int(bistand + 0.5)
    'endif
    disp = disp 'OM Bidragsmotagare mer än sex månader
    
    If disp < bistand Then
        bistand = (12 * bistand - disp)
        If marginal = 0 Then bistand = Int(bistand + 0.5)
        'bistand=bistand *month /12
    Else
        bistand = 0
    End If
    bistOld = bistand
    
End Function

Function bist25(ByVal civ As Integer, ByVal hyra As Double, ByVal disp As Double, _
    ByVal b1 As Integer, ByVal b2 As Integer, ByVal b3 As Integer, ByVal b4 As Integer, ByVal b5 As Integer, _
    ByVal b6 As Integer, ByVal b7 As Integer, ByVal b8 As Integer, ByVal wage As Double, Optional year = 2025) As Double
    ' Socialstyrelsens riksnorm för försörjningsstöd inkomståren 2006-
    
    'Civ - Civilstånd (antal vuxna)
    'Hyra - Skälig boendekostand
    'Disp - Hushållets disponibla inkomst
    'B1 - Antal barn 0 år
    'B2 - Antal barn 1-2   år gamla
    'B3 - Antal barn 3 år gamla
    'B4 - Antal barn 4-6   år gamla
    'B5 - Antal barn 7-10  år gamla
    'B6 - Antal barn 11-14 år gamla
    'B7 - Antal barn 15-18 år gamla
    'B8 - Antal barn 19-20 år gamla
    'Wage - Nettolön obs 0 om bistånd inte har föregåtts om sex månader i anropet
    'year - inkomstår
    
    Const Verbose_ = 0 'koll av värdena 1-Personliga, 2 Vuxna 3 hushåll 9 alla
    
    bist25 = 0
    If year < 2005 Then Exit Function 'se bist0ld
    
'--------Personliga kostnader för hemmavarande barn utan lunch en funktion av barnets ålder per månad
    Dim xn() As Variant 'as double 'OBS option BASE 1
    
    'ReDim xn(1 To 8)
'    xn(1) = 2720 '0 år
'    xn(2) = 3030 '1-2
'    xn(3) = 2700 '3
'    xn(4) = 3030 '4-6
'    xn(5) = 3790 '7-10
'    xn(6) = 4350 '11-14
'    xn(7) = 4920 '15-18
'    xn(8) = 4960 '19-20
                    '0,   1-2,    3,  4-6, 7-10,11-14,15-18,19-20
    If year = 2006 Then
       xn = Array(1540, 1740, 1430, 1740, 1950, 2230, 2510, 2600) '19 år vuxen
    ElseIf year = 2007 Then
       xn = Array(1560, 1770, 1450, 1770, 1980, 2270, 2550, 2640) '19 år vuxen
    ElseIf year = 2008 Then
       xn = Array(1590, 1810, 1480, 1810, 2020, 2320, 2610, 2700) '19 år vuxen
    ElseIf year = 2009 Then
       xn = Array(1650, 1880, 1540, 1880, 2100, 2410, 2710, 2740)
    ElseIf year = 2010 Then
       xn = Array(1650, 1880, 1540, 1880, 2100, 2410, 2710, 2740)
    ElseIf year = 2011 Then
       xn = Array(1670, 1900, 1560, 1900, 2130, 2440, 2740, 2770)
    ElseIf year = 2012 Then
       xn = Array(1720, 1960, 1710, 1960, 2390, 2810, 3220, 3250)
    ElseIf year = 2013 Then
       xn = Array(1740, 1980, 1730, 1980, 2410, 2840, 3250, 3280)
    ElseIf year = 2014 Then
       xn = Array(1740, 1980, 1730, 1980, 2410, 2840, 3250, 3280)
    ElseIf year = 2015 Then
       xn = Array(1740, 1980, 1730, 1980, 2410, 2840, 3250, 3280)
    ElseIf year = 2016 Then
       xn = Array(1840, 2080, 1830, 2080, 2660, 3090, 3500, 3530)
    ElseIf year = 2017 Then
       xn = Array(1860, 2100, 1850, 2100, 2690, 3120, 3540, 3570)
    ElseIf year = 2018 Then
       xn = Array(2090, 2330, 2080, 2330, 2930, 3370, 3800, 3830)
    ElseIf year = 2019 Then
       xn = Array(2130, 2380, 2120, 2380, 2990, 3440, 3880, 3910)
    ElseIf year = 2020 Then
       xn = Array(2170, 2430, 2160, 2430, 3050, 3510, 3950, 3980)
     ElseIf year = 2021 Then
       xn = Array(2180, 2440, 2170, 2440, 3060, 3520, 3970, 4000)
    ElseIf year = 2022 Then
       xn = Array(2220, 2480, 2210, 2480, 3110, 3580, 4040, 4070)
    ElseIf year = 2023 Then
       xn = Array(2420, 2700, 2410, 2700, 3380, 3890, 4390, 4430)
    ElseIf year = 2024 Then
       xn = Array(2640, 2940, 2620, 2940, 3680, 4230, 4780, 4820)
       
    Else
        xn = Array(2720, 3030, 2700, 3030, 3790, 4350, 4920, 4960)
    
    End If
    
      
    Dim bistand As Double
    bistand = 0
    Dim barn(1 To 8) As Integer
    barn(1) = b1
    barn(2) = b2
    barn(3) = b3
    barn(4) = b4
    barn(5) = b5
    barn(6) = b6
    barn(7) = b7
    barn(8) = b8
   
    Dim i, antal As Integer
    antal = 0 'Antal barn
    For i = 1 To 8
            bistand = bistand + xn(i) * barn(i)
            antal = antal + barn(i)
            
            If (Verbose_ = 1 Or Verbose_ = 9) Then Debug.Print year, "X(" & i & ")"; year; xn(i)
            
    Next i
    
    antal = antal + civ 'Antal familjemdedlemmar
    
'---------- Personlig kostnader vuxna
    Dim vuxna() As Variant
    
    If year = 2006 Then
        vuxna = Array(2600, 4690)
    ElseIf year = 2007 Then
        vuxna = Array(2640, 4770)
    ElseIf year = 2008 Then
        vuxna = Array(2700, 4870)
    ElseIf year = 2009 Then
        vuxna = Array(2800, 5060)
    ElseIf year = 2010 Then
        vuxna = Array(2800, 5060)
    ElseIf year = 2011 Then
        vuxna = Array(2830, 5120)
    ElseIf year = 2012 Then
        vuxna = Array(2920, 5270)
    ElseIf year = 2013 Then
        vuxna = Array(2950, 5320)
    ElseIf year = 2014 Then
        vuxna = Array(2950, 5320)
    ElseIf year = 2015 Then
        vuxna = Array(2950, 5320)
    ElseIf year = 2016 Then
        vuxna = Array(2950, 5320)
    ElseIf year = 2017 Then
        vuxna = Array(2980, 5370)
    ElseIf year = 2018 Then
        vuxna = Array(3030, 5460)
    ElseIf year = 2019 Then
        vuxna = Array(3090, 5570)
    ElseIf year = 2020 Then
        vuxna = Array(3150, 5680)
    ElseIf year = 2021 Then
        vuxna = Array(3160, 5700)
    ElseIf year = 2022 Then
        vuxna = Array(3210, 5800)
    ElseIf year = 2023 Then
        vuxna = Array(3490, 6300)
    ElseIf year = 2024 Then
        vuxna = Array(3800, 6850)
        
    Else
         vuxna = Array(3910, 7500)
    End If
    
    If (Verbose_ = 2 Or Verbose_ = 9) Then Debug.Print year, year, "Civ 1 och 2", vuxna(1), vuxna(2)
    
    
    'Bistand + de vuxnas och och boendekostnad
    If civ = 1 Then
        bistand = bistand + vuxna(1)
        hyra = hyra 'Antas skälig
    Else
        bistand = bistand + vuxna(2)
        hyra = hyra 'Antas skälig
    End If
    
'------Gemensamma hushållskostnader
    Dim Gn() As Variant 'Double
    'ReDim Gn(1 To 7)

    If year = 2006 Then
        Gn = Array(820, 920, 1140, 1320, 1510, 1710, 1880)
    ElseIf year = 2007 Then
        Gn = Array(830, 930, 1160, 1340, 1530, 1740, 1910)
    ElseIf year = 2008 Then
        Gn = Array(850, 950, 1190, 1370, 1560, 1780, 1950)
    ElseIf year = 2009 Then
        Gn = Array(880, 990, 1240, 1420, 1620, 1850, 2020)
    ElseIf year = 2010 Then
        Gn = Array(880, 990, 1240, 1420, 1620, 1850, 2020)
    ElseIf year = 2011 Then
        Gn = Array(890, 1000, 1260, 1440, 1640, 1870, 2040)
    ElseIf year = 2012 Then
        Gn = Array(920, 1030, 1300, 1480, 1690, 1930, 2100)
    ElseIf year = 2013 Then
        Gn = Array(930, 1040, 1310, 1490, 1710, 1950, 2120)
    ElseIf year = 2014 Then
        Gn = Array(930, 1040, 1310, 1490, 1710, 1950, 2120)
    ElseIf year = 2015 Then
        Gn = Array(930, 1040, 1310, 1490, 1710, 1950, 2120)
    ElseIf year = 2016 Then
        Gn = Array(940, 1050, 1320, 1500, 1720, 1960, 2130)
    ElseIf year = 2017 Then
        Gn = Array(950, 1060, 1330, 1520, 1740, 1980, 2150)
    ElseIf year = 2018 Then
        Gn = Array(970, 1080, 1350, 1540, 1770, 2010, 2180)
    ElseIf year = 2019 Then
        Gn = Array(990, 1100, 1380, 1570, 1810, 2050, 2220)
    ElseIf year = 2020 Then
        Gn = Array(1010, 1120, 1410, 1600, 1850, 2090, 2260)
    ElseIf year = 2021 Then
        Gn = Array(1020, 1130, 1420, 1610, 1860, 2100, 2270)
    ElseIf year = 2022 Then
        Gn = Array(1040, 1150, 1450, 1640, 1890, 2140, 2310)
    ElseIf year = 2023 Then
        Gn = Array(1130, 1250, 1580, 1790, 2060, 2330, 2510)
    ElseIf year = 2024 Then
        Gn = Array(1230, 1360, 1720, 1950, 2240, 2540, 2730)
        
    Else
        Gn = Array(1270, 1400, 1770, 2010, 2310, 2620, 2810)
    End If
    
    If (Verbose_ = 3 Or Verbose_ = 9) Then
        For i = 1 To 7
            Debug.Print year; "G(" & i & ")", Gn(i)
        Next i
    End If
    
    Dim flera
    flera = Gn(7) - Gn(6)
    
    If antal > 7 Then
        bistand = bistand + Gn(7) + flera * (antal - 7)
    Else
        bistand = bistand + Gn(antal)
    End If
    'Maximalt stöd som reduceras med inkomsten men
    Dim just As Double
    just = 1 'Antas förändras med KPI vid framskrivning
    If year > 2025 Then
        just = KPI(age) / KPI(2025 - Int(born))
    End If
    bistand = bistand * just + hyra '+barnomsorg uppräknade till 2025 med antagen prisutv. samma föa alla delkomponenter
        
    If year >= 2013 Then disp = disp - wage * 0.25 'OM Bidragsmotagare mer än sex månader när?
    
    If disp < bistand * 12 Then
        bistand = (12 * bistand - disp)
        If marginal = 0 Then bistand = Int(bistand + 0.5)
        'bistand=bistand *month /12
    Else
        bistand = 0
    End If
    bist25 = bistand 'Årsbasis
    
End Function

'Sub verb()
'Dim koll As Double
'Dim year As Long
''46240 i bidrag - fem barn; 1 barn 5 år, 2 barn 7-10 år och 1 barn 13 år. med hyra 20600 kr
'For year = 2025 To 2025 'Kräver att verbose=1 i funktionen
'
'    koll = bist25(2, 20600, 0, 0 * 1, 0 * 2, 0 * 3, 1, 1, 2, 0 * 7, 0 * 8, 0, year)
'    Debug.Print year, koll / 12
'Next year
'End Sub


''Sub kollm()
''    'Civ - Civilstånd (antal vuxna)
''    'Hyra - Skälig boendekostand
''    'Disp - Hushållets disponibla inkomst
''    'B1 - Antal barn 0 år
''    'B2 - Antal barn 1-2   år gamla
''    'B3 - Antal barn 3 år gamla
''    'B4 - Antal barn 4-6   år gamla
''    'B5 - Antal barn 7-10  år gamla
''    'B6 - Antal barn 11-14 år gamla
''    'B7 - Antal barn 15-18 år gamla
''    'B8 - Antal barn 19-20 år gamla
''    'Wage - Nettolön obs 0 om bistånd inte har föregåtts om sex månader i anropet
''
''    Dim koll As Double
''    koll = 0
''    Dim civ As Integer
''    civ = 1
''    Dim ink, hyra As Double
''    ink = 0: hyra = 0
''    Dim b1, b2, b3, b4, b5, b6, b7, b8 As Integer
''
''    b1 = 0  'B1 - Antal barn 0 år
''    b2 = 0  'B2 - Antal barn 1-2   år gamla
''    b3 = 1  'B3 - Antal barn 3 år gamla
''    b4 = 1  'B4 - Antal barn 4-6   år gamla
''    b5 = 0  'B5 - Antal barn 7-10  år gamla
''    b6 = 0  'B6 - Antal barn 11-14 år gamla
''    b7 = 0  'B7 - Antal barn 15-18 år gamla
''    b8 = 0  'B8 - Antal barn 19-20 år gamla
''    Dim year As Long
''
'''    Do While civ < 2
'''        koll = bist25(civ, 7000, ink, b1, b2, b3, b4, b5, b6, b7, b8, 0)
'''        Debug.Print koll
'''        koll = 0
'''        civ = civ + 1
'''    Loop
'''bistOld(ByVal civ As Integer, ByVal hyra As Double, ByVal disp As Double, _
'''    ByVal b1 As Integer, ByVal b2 As Integer, ByVal b3 As Integer, ByVal b4 As Integer, ByVal b5 As Integer, _
'''    ByVal b6 As Integer, ByVal b7 As Integer, ByVal b8 As Integer, ByVal wage As Double, Optional year = 2005) As Double
''For year = 1985 To 2025
''    If year < 2006 Then
''        koll = bistOld(civ, hyra, ink, b1, b2, b3, b4, b5, b6, b7, b8, 0, year)
''    Else
''        koll = bist25(civ, hyra, ink, b1, b2, b3, b4, b5, b6, b7, b8, 0, year)
''    End If
''        Debug.Print year; koll
''Next year
''End Sub

Sub kollbist()
Dim koll As Double
Dim year As Long

'Dim test() As Double
'ReDim test(1 To 2)
'test(2) = CDbl(5712) * 9 / 12 + 4150 * 3 / 12

'bistOld(ByVal civ As Integer, ByVal hyra As Double, ByVal disp As Double, _
'    ByVal b1 As Integer, ByVal b2 As Integer, ByVal b3 As Integer, ByVal b4 As Integer, ByVal b5 As Integer, _
'    ByVal b6 As Integer, ByVal b7 As Integer, ByVal b8 As Integer, ByVal wage As Double, Optional year = 2005) As Double

Close #1
Dim FileName As String
FileName = ThisWorkbook.Path & "\Bist.csv" 'Koll om bibliteket finns Call CheckFfolder(folder)

Open FileName For Output As #1
'Write #1, "test line #1"
'Print #1, "test line #2"
    
        'b1 = 0 '- Antal barn 0 år
        'b2 = 0 '- Antal barn 1-2   år gamla
        'b3 = 0 '- Antal barn 3 år gamla
        'b4 = 0 '- Antal barn 4-6   år gamla
        'b5 = 0 '- Antal barn 7-10  år gamla
        'b6 = 0 '- Antal barn 11-14 år gamla
        'b7 = 0 '- Antal barn 15-18 år gamla
        'b8 = 0 '- Antal barn 19-20 år gamla
    
For year = 1985 To 2025 'Verbose=9
    If year < 2006 Then
        koll = bistOld(1, 0, 0, 1 * 0, 2 * 0, 3 * 0, 4 * 1, 5 * 1, 6 * 0, 7 * 0, 8 * 0, 0, year)
    Else
        koll = bist25(1, 0, 0, 1 * 0, 2 * 0, 3 * 0, 4 * 1, 5 * 1, 6 * 0, 7 * 0, 8 * 0, 0, year)
    End If
    'Debug.Print "koll ", year, koll
    
    Print #1, year & ";" & koll
Next year
Close #1
End Sub


Function forbehall(ByVal avgift, ByVal inkomst, ByVal civ, ByVal year, ByVal pbb) As Double
    'Förbehåll hemtjänst används inte och behöver ses över historiskt
    'Avgift hemtjänsten
    'Inkomst
    ' Civ - ensam eller samboende
    ' year - årtal
    ' Pbb - prisbasbelopp
    Dim lim1 As Double
    
''    If Application.Range("Rulesfromutg") = 0 Or IsMissing(Application.Range("Rulesfromutg")) Then
''             'Nothing
''        Else
''            If (year > Application.Range("Rulesfromutg") And Application.Range("Rules") = 0) _
''            Or Application.Range("Rules") = 1 Then
''                year = Application.Range("Rulesfromutg")
''            End If
''    End If
    
    forbehall = 0
    
'Edited 2024-01-23 - did not compile
'    If civ = 0 Then lim1 = 1.3546
'    Else: lim1 = 1.1446
    If civ = 0 Then
        lim1 = 1.3546
    Else
        lim1 = 1.1446
    End If
    
    If year > 2019 Then '???
        If civ = 0 Then lim1 = 1.4044
        Else: lim1 = 1.1694
    End If
    If year > 2021 Then
        If civ = 0 Then lim1 = 1.44541
        Else: lim1 = 1.1942
    End If
    If year = 2022 Then
        If civ = 0 Then lim1 = (1.4044 * 7 + 5 * 1.4789) / 12
        Else: lim1 = (1.1694 * 7 + 1.2066 * 5) / 12
    End If
    If year > 2022 Then
        If civ = 0 Then lim1 = 1.4789
        Else: lim1 = 1.2066
    End If
    

    lim1 = lim1 * pbb
    
    If inkomst - avgift < lim1 Then
        forbehall = lim1 - inkomst
    End If
    
End Function




