Attribute VB_Name = "Pensionssystemet"
Option Explicit
'Diverse funktioner om pensionssystemet

Function riktage(ByVal year As Long, ByVal Typ As Byte) As Single
    'Lägsta ålder för pensionsuttag i den allmänna pensionen eller den ålder när äldres grundskydd träder in (riktålder)
    'Year inkomstår
    'Type 0=lägsta, annars rikt
    Dim rikt, riktl As Single
    rikt = 65: riktl = 61
    
     'Obs delvis prognos utanför modellen
        If year < 2020 Then
            rikt = 65: riktl = 61
        ElseIf year < 2023 Then
            rikt = 65: riktl = 62
        ElseIf year < 2026 Then
            rikt = 66: riktl = 63
        ElseIf year < 2038 Then
            rikt = 67: riktl = 64
        ElseIf year < 2051 Then
              rikt = 68: riktl = 65
        ElseIf year < 2068 Then
              rikt = 69: riktl = 66
        Else
              rikt = 70: riktl = 67
        End If
    
    riktage = rikt
    If Typ = 0 Then riktage = riktl
End Function


''Sub kolla_rikt()
''    Dim i, l, h As Single
''    Dim yy As Long
''    yy = 2020
''    For i = 1 To 7
''        l = yy - riktage(yy, 0) + 1
''        h = yy - riktage(yy, 1) + 1
''        Debug.Print yy; l; riktage(yy, 0); h; riktage(yy, 1)
''        yy = 2023
''        If i = 2 Then yy = 2026
''        If i = 3 Then yy = 2038
''        If i = 4 Then yy = 2051
''        If i = 5 Then yy = 2068
''        If i = 6 Then yy = 2069
''    Next i
''End Sub

Function pgi(ByVal year As Long, ByVal inkomst As Double, pbb As Double, IBB As Double, FHB As Double, _
       Optional marginal = 0, Optional Typ = 0, Optional alder = 64, Optional sjuk = 0) As Double
  'Application.Volatile
    'Pensionsgrundande inkomst, även betydelse för skatten genom skattreduktion för pensionsavgiften

    'Year    - Inkomstår
    'Inkomst - bruttoinkomst under inkomståret
    'PBB     - Prisbasbelopp
    'IBB     - Inkomstbasbelopp
    'FHB     - Förhöjt basbelopp
    'Marginal 0 -> dagens Avrudningar, 1 Utan avrudningar
    'TYP      0 -> PGI inkomst, 1->PAVG, 2->SRED, 9->Inkomst
    'Alder   - Ålder 31/12
    'Sjuk    - Sjuk- eller aktivitetsersättning
    
    pgi = 0

    Dim Tak As Double   'Maximal pensionsrätt
    Dim golv As Double  'Deklarationsplikt
    Dim egen As Double  'Egenavgift
    Dim PAVG As Double  'Pensionsavgift
    Dim SRED As Double  'Skattereduktion för pensionsavgiften

    'Obs inkomst efter avdrag under inkomst av tjänst, reseavdrag pensionssparande mm.
    'exkl. den inkomstrelaterade ers. från sjuk och aktivitetsersättning !
    inkomst = inkomst - sjuk   'Staten betalar hela avgiften för sjuk- och aktivitetsersättningen 18,5%
    
    'I modellen anges löneinkomst utan avdrag, varför den taxerade förvärvsinkomsten är:
    'OBS Skiljer på näringsinkomster som läggs till det avrundade och sedan avrundas båda. Här ingår de inkomsten .
''
''   If Application.Range("Rulesfromutg") = 0 Or IsMissing(Application.Range("Rulesfromutg")) Then
''             'Nothing
''        Else
''            If (year > Application.Range("Rulesfromutg") And Application.Range("Rules") = 0) _
''            Or Application.Range("Rules") = 1 Then
''                year = Application.Range("Rulesfromutg")
''            End If
''    End If
    
    If marginal = 0 Then inkomst = Int(inkomst / 100) * 100 'Inkomster avrundas nedåt
    
    Tak = 8.07 * IBB
    If year <= 1994 Then
      Tak = 7.5 * pbb
    ElseIf year < 1999 Then
      Tak = 7.5 * FHB
    ElseIf 1999 = year Then
      Tak = 8.06 * FHB
    ElseIf 2000 = year Then
       Tak = 8.07 * FHB
    End If
    
    'Höjt tak
    If Application.Range("Soc_tak") > 1999 Then
        If year >= Application.Range("Soc_tak") Then Tak = 7.5 * IBB
    End If
 
    'GOLVET
    golv = 0.423 * pbb
    If year <= 1994 Then
       golv = 1 * pbb
    ElseIf year < 1999 Then
       golv = 1 * FHB
    ElseIf year < 2001 Then
       golv = 0.24 * pbb
    ElseIf 2001 = year Then
       golv = 0.27 * pbb
    ElseIf 2002 = year Then
       golv = 0.293 * pbb
    Else
       golv = 0.423 * pbb
    End If
    
''    If marginal = 1 Then golv = 0.423 * pbb

    If inkomst <= golv Then
       pgi = 0
    ElseIf inkomst <= Tak Then
       pgi = inkomst
    ElseIf inkomst > Tak Then
       pgi = Tak
    End If

    egen = 0.07
   
    'Egenavgift till pensionssystemet '+ sjuk+ arbetslös se flik k_skatt
    If year < 1993 Then
        egen = 0
    ElseIf year = 1993 Then
       egen = 0 '+0.0095+0
    ElseIf year = 1994 Then
       egen = 0 '+0.0095+0.01
    ElseIf year = 1995 Then
       egen = 0.01  '+0.0295+0.0
    ElseIf year = 1996 Then
        egen = 0.01 '+0.0395+0
    ElseIf year = 1997 Then
        egen = 0.01 '+0.0495+0
    ElseIf year < 2000 Then
        egen = 0.0695 '+0+0
    Else
        egen = 0.07 '+0+0
    End If
   
    'If (Year - alder) < 1937 and alder<64  Then PGI = 0
    'If alder < 16 Then PGI = 0
    PAVG = pgi * egen

    'osäker på tidigare avrundningsregler
    If marginal = 0 And year > 1998 Then PAVG = Int((PAVG + 49) / 100) * 100
    If marginal = 0 And year < 1999 Then PAVG = Int((PAVG + 49) / 100) * 100
    
    If Application.Range("Social_avg ") > 1999 Then
        If year >= Application.Range("Social_avg") Then PAVG = 0
    End If
 
    'Hela PAVG ger numera en skattreduktion med samma belopp, Avgiften avrundas till närmast hela hundratal kronor. Avgift som slutar på 50 kronor avrundas till närmast lägre hundratal kronor
    'I skattepgm räknar vi på senaste skattereglerna... Annars måste vi backa historisk till hur stor del som var avdragsgill och hur stor andel som gav en skattreduktion
    Dim andel As Long
    andel = 1

    If year < 2000 Then
       andel = 0
    ElseIf year <= 2000 Then
       andel = 0.25
    ElseIf year <= 2001 Then
       andel = 0.5
    ElseIf year <= 2002 Then
       andel = 0.75
    ElseIf year <= 2005 Then
       andel = 0.875
    Else
       andel = 1
    End If
 
    SRED = andel * PAVG
   
    If marginal = 0 Then SRED = Int(SRED / 100) * 100
    'besk.bar inkomst = taxerad förv. inkomst - Grundavdrag - egenavgift + skatte red för avgiften.
    'PGI = PGI - PAVG tas senare
    
    inkomst = inkomst + sjuk
    If marginal = 0 Then inkomst = Int(inkomst / 100) * 100 'Inkomster avrundas nedåt
    'Inkomst =inkomst+naring, vi slår ihp inkomsterna innan vilket kan ge 100 kr för mycket
   
    'Höjd avgift
    If year >= Application.Range("Social_avg") And Application.Range("Social_avg") > 1999 Then
            If inkomst < golv Then
               pgi = 0
            ElseIf inkomst <= Tak Then
               pgi = inkomst
            ElseIf inkomst > Tak Then
               pgi = Tak
            End If
    Else
        If inkomst < golv Then
           pgi = 0
        ElseIf inkomst <= Tak Then
           pgi = inkomst - PAVG
        ElseIf inkomst > Tak Then
           pgi = Tak - PAVG
        End If
    End If
    
    'Avrundningsreglerna för PGI, ska kollas men
    If year < 1999 And marginal = 0 Then pgi = Int(pgi / 50) * 50 'Avrundat nedåt till närmaste 50 kr
    If 1999 <= year And marginal = 0 Then pgi = Int(pgi / 100) * 100 'Avrundat nedåt till närmaste hundratal kr

    'Utifrån denna pgi räknas pensionsrätten
    'Korr för SOCIALDEP
'''    If Application.Range("Soc_tak") > 1999 Then
'''        If year >= Application.Range("Soc_tak") Then
'''            If Application.Range("Social_avg") > 1999 And year >= Application.Range("Social_avg") Then
'''                If pgi > 10 * IBB Then pgi = 10 * IBB 'tak = 10 * ibb
'''                'If year = Application.Range("Social_avg") Then Debug.Print "taket och PGI är år " & 10 * ibb & " " & PGI & " " & year
'''            Else
'''                If pgi > 9.3 * IBB Then pgi = 9.3 * IBB 'tak = 10 * ibb
'''                'If year = Application.Range("Soc_tak") Then Debug.Print "taket är 9.3 " & 9.3 * ibb & " " & PGI & " " & year
'''            End If
'''        End If
'''    Else
        If pgi > 7.5 * IBB Then pgi = 7.5 * IBB 'Kan således sluta på 50 kr...
'''    End If
    If Typ = 1 Then pgi = PAVG
    If Typ = 2 Then pgi = SRED
    If Typ = 4 Then pgi = Tak
    If Typ = 9 Then pgi = inkomst
    
End Function
''
''Sub kolla()
''    Dim pgi_ As Double
''    Dim pr As Double
''    Dim ink As Double
''    ink = 74300 * 8.07 + 10000
''    Dim i As Integer
''
''    'PGI(year inkomst,  pbb, IBB , FHB , _
''       Optional marginal = 0, Optional typ = 0, Optional alder = 64, Optional sjuk = 0) As Double
''    'For i = 1 To 5
''        pgi_ = pgi(2023, ink, 52500, 74300, 53300, 0, -2, 65, 0)
''        pr = pgi(2023, ink, 45000, 60000, 45000, 0, 1, 65, 0)
''        Debug.Print ink; pgi_; pr
''    '    ink = ink + i * 5000
''    'Next i
''End Sub
'******************************************************PGI Slut ************************

Function andel(ByVal Kohort As Double) As Double
  'Andelar av det "nya" IP-systemet, (1-andel)= Andel av ATP
  'Kohort - födelseår ÅÅÅÅ
  'Application.Volatile
   andel = 0
   Kohort = Int(Kohort)
  If 1938 <= Kohort Then andel = (1 / 20) * (Kohort - 1935 + 1)
  If Kohort > 1953 Then andel = 1
End Function

'''Sub kolla_andel()
'''    Dim koll As Double
'''    Dim kull As Integer
'''
'''    For kull = 1938 To 1952
'''        koll = andel(kull)
'''        Debug.Print kull; koll
'''    Next
'''
'''End Sub

Function ipavgift(ByVal year As Long, ByVal pgi As Double, ByVal alder As Integer, _
  Optional ByVal andel = 1, Optional ByVal marginal = 0, Optional ByVal fodd = 1938) As Double
 'Application.Volatile
  'Avgiften enbart för födda 1938 och senare, Men PGI är 0 för födda 1937 och tidigare.
  'Kohort = year - alder
  'Andel = Andel(Kohort)
  'Year  - Inkomstår
  'PGI   - Pensionsgrudnande inkomst
  'Alder - Ålder 31/12
  'Andel - Andel av det nya systemet
  'Marginal 1 utan avrundningar och dagens system
  'fodd  = Födelseår
  If fodd < 1938 Then GoTo slut:
  
    
''   If Application.Range("Rulesfromutg") = 0 Or IsMissing(Application.Range("Rulesfromutg")) Then
''             'Nothing
''        Else
''            If (year > Application.Range("Rulesfromutg") And Application.Range("Rules") = 0) _
''            Or Application.Range("Rules") = 1 Then
''                year = Application.Range("Rulesfromutg")
''            End If
''    End If
    
  
  Dim avgift As Double
  avgift = 0.16
  'Tidigare andra avgiftssatser, givet inga marginalberäkningar
  If year <= 1994 And marginal = 0 Then avgift = 0.185
  If 1994 < year And year <= 1998 And marginal = 0 Then avgift = 0.165
   
    Rem RUT Experiment................. _
    If year > 2013 Then avgift = 0.185
    
    Rem Experiment.................
    'andel = 1

 'Andel =1 om 65<=ålder
  If alder > 64 Then andel = 1
  ipavgift = avgift * pgi * andel  'Andel skulle iof kunna läsas in men
  'Avrundas till närmaste lägre hela krontal
  If marginal = 0 Then ipavgift = Int(ipavgift)
slut:
End Function

Function ppavgift(ByVal year As Long, ByVal pgi As Double, ByVal alder As Integer, _
   Optional ByVal andel = 1, Optional ByVal marginal = 0, Optional ByVal fodd = 1938) As Double
  'Avgiften enbart för födda 1938 och senare, Men PGI är 0 för födda 1937 och tidigare.
   'Application.Volatile
  If fodd < 1938 Then GoTo slut:
   
''   If Application.Range("Rulesfromutg") = 0 Or IsMissing(Application.Range("Rulesfromutg")) Then
''             'Nothing
''        Else
''            If (year > Application.Range("Rulesfromutg") And Application.Range("Rules") = 0) _
''            Or Application.Range("Rules") = 1 Then
''                year = Application.Range("Rulesfromutg")
''            End If
''    End If
 
  Dim avgift As Double
  avgift = 0.025
  If year <= 1994 Then avgift = 0
  If 1994 < year And year <= 1998 Then avgift = 0.02
  
    Rem RUT Experiment................. _
    If year > 2013 Then avgift = 0
    
    Rem Experiment.................
    'andel = 1
   
  If alder > 64 Then andel = 1
  ppavgift = avgift * pgi * andel
  If marginal = 0 Then ppavgift = Int(ppavgift) 'avrundning till närmaste lägre hela krontal
slut:
End Function



Function gpavgift(ByVal year As Long, ByVal pgi As Double, ByVal alder As Integer, Optional ByVal andel = 1, Optional ByVal marginal = 0, Optional ByVal fodd = 1938, Optional ByVal rikt = 65) As Double
    'som om hela kom från IP för beräkning av Garantipension
     'Application.Volatile
     gpavgift = 0
     If fodd < 1938 Then GoTo slut:
''
''   If Application.Range("Rulesfromutg") = 0 Or IsMissing(Application.Range("Rulesfromutg")) Then
''             'Nothing
''        Else
''            If (year > Application.Range("Rulesfromutg") And Application.Range("Rules") = 0) _
''            Or Application.Range("Rules") = 1 Then
''                year = Application.Range("Rulesfromutg")
''            End If
''    End If
    
     
     Dim avgift As Double
     avgift = 0.185
     If alder >= rikt Then andel = 1
     gpavgift = avgift * pgi * andel
     If marginal = 0 Then gpavgift = Int(gpavgift)
slut:
End Function


Function PPMavg(ByVal year As Long, ByVal kapital As Double, faktor As Double, ByVal avkast As Double, avg2 As Double, Optional MyndAvg As Double)
     'Year          - Inkomstår
     'Kapital       - fonderat kapital
     'Faktor        - Kvar efter förvaltningskostnad 1 - 0,xxxx
     'Avkastning    - Netto
     'Avg2          - fondbolagens avgifter
     'Myndavg       - Avgift för myndigheten (0 i dagsläget)
     'Application.Volatile
    PPMavg = 0
    
''    If Application.Range("Rulesfromutg") = 0 Or IsMissing(Application.Range("Rulesfromutg")) Then
''             'Nothing
''        Else
''            If (year > Application.Range("Rulesfromutg") And Application.Range("Rules") = 0) _
''            Or Application.Range("Rules") = 1 Then
''                year = Application.Range("Rulesfromutg")
''            End If
''    End If
    
    If Application.Range("rng_Avkastning_fondavgifter") = 1 Then
        'faktor = 1
        'fakt2 = 0
        Exit Function
    End If
    
    
    ' PPM-förvaltningsavgift år t:
    '    [ Kapital(t-1) x (1+avkastning(t)^2/12] x avgiftsfaktor(t) dock med ett tak
    '    Kapital(t-1) x avgiftsfaktor(t) för åren fram till och med 2010. Innan 2010 var det ingen max gräns. För 2010 uppgick taket till 125 kronor
    '
    
    Dim fakt2 As Double 'Avgiften för förvaltning hos fondbolagen
    fakt2 = 0 'Netto före 2018 men alltid noll i modellen i dagsläget
    'If year > 2013 And avg2 > 0 Then fakt2 = 0.003 'Har och antas vara konstant
    
    'Pensionsmyndighetens avgift
    
    If year < 2011 Then
       PPMavg = kapital * (1 - faktor + fakt2)
    Else
       PPMavg = kapital * (1 - faktor + fakt2) * (1 + avkast) ^ 0.5
    End If
    'Som takas - uppdateras årligen
    If year < 2007 Then
     'PPMavg = PPMavg
    ElseIf year = 2007 And PPMavg > 100 Then
      PPMavg = 100
    ElseIf year < 2010 And PPMavg > 110 Then
      PPMavg = 110
    ElseIf year = 2010 And PPMavg > 125 Then
      PPMavg = 125
    ElseIf year < 2014 And PPMavg > 110 Then
      PPMavg = 110
    ElseIf year < 2017 And PPMavg > 120 Then
      PPMavg = 120
    ElseIf year < 2018 And PPMavg > 125 Then
      PPMavg = 125
    ElseIf year < 2019 And PPMavg > 125 Then
      PPMavg = 160
    Else
      PPMavg = 100
    End If

    Dim ppmavg2 As Double
    ppmavg2 = kapital * fakt2 * (1 + avkast) ^ 0.5
  
    PPMavg = -PPMavg - ppmavg2
  
End Function

Function tp_faktor(ByVal PAR As Double) As Double
    'Beräkning av den faktor som TP ska justeras med för uttag före/efter 65 årsdagen
    'Application.Volatile
    
    Dim month_2 As Long  ' Antal månader efter pensionsåldern t.ex. 65,2 ->2,4-> 2
    If PAR > 70 Then
        month_2 = 0
    Else
        month_2 = Int(12 * (PAR - Int(PAR)))
    End If
    
    Dim faktor As Double
    If Int(PAR) < 65 Then
        ' uttagsfaktorn sänktes från 0,6 till 0,5 % per månad 1/7 1976 bortses från detta och att Pensionsåldern sänks från 67 till 65
        faktor = (1 - 0.005 * (12 * (65 - Int(PAR)) - month_2))
    ElseIf Int(PAR) = 65 Then
        faktor = (1 + 0.007 * month_2)
        'OBS Ökningsfaktorn om 0,7 procent ska ersättas med 0,6 procent för uttag som avser tid före den 1 juli 1990, detta bortses från
        '1/7 1976 vid uppskutett uttag ökas pensionen med 0,5 % per månad
    ElseIf Int(PAR) > 65 And Int(PAR) <= 70 Then
        faktor = 1 + 0.007 * (12 * (Int(PAR) - 65) + month_2)
    Else 'Efter 70
        faktor = (1 + 0.007 * (12) * (70 - 65))
    End If
    'Vid Garantpension ska underalget för tp-> 1/max(1;faktor)
    tp_faktor = faktor
End Function


Function tp_(ByVal poang As Double, ByVal nyear As Long, ByVal civ As Integer, ByVal PAR As Double, ByVal fodar As Double, ByVal alder As Integer, _
  Optional ByVal andel_ = 1, Optional ByVal def_ar = -99, Optional ByVal year = 2012, Optional ByVal marginal = 0) As Double
'Omgjord våren 2021 för att enbart beräkna första uttaget
     
     'Application.Volatile
    'Poäng  - ATP medelpoäng
    'Nyear  - Antal år med ATP poäng, förvärvskriteriet
    'Civ    - Civilstånd 0-ogift 1-gift
    'Par    - Pensionsålder vid första eller andra uttaget ÅÅ,åå
    'fodar  - Födelseår, ÅÅÅÅ,åå
    'Alder  - Ålder
    'Andel  - Uttagsdel (uttagIP)
    'def_ar - obsolet var tidigare Ålder vid det definitiva pensioneringen
    'year   - Inkomstår
    'Marginal - avrundningar

     'OBS räknat på dagens tilläggspension, historiskt kan parametrarna ha ändrats
     tp_ = 0
    
    If alder < Int(PAR) Then Exit Function
    If fodar > 1953 Then Exit Function
    
''    If Int(par) < 61 Then
''        'MsgBox ("Ingen ATP före 61 års ålder")
''        exit function
''    End If
''    'Angivit definitift år
''    If IsNumeric(Def_ar) = False Then Def_ar = -99
''    If Application.Range("UttagIP").Value = 1 Then Def_ar = -99 'par
''
''    If par >= Def_ar Then
''        andel_ = 1 'Ej angivit definitift år så vid par medför full pension
''        Def_ar = 999
''    End If

    Dim ATP As Double   'ATP-delen + FP-delen  beräknas
    If civ = 0 Then 'HEL TP Inkl FP delen - obs civilståndsbyte kräver en viss handpåläggning..
        ATP = 0.6 * poang * pbb(age) + 1 * 0.96 * pbb(age)
    Else
        ATP = 0.6 * poang * pbb(age) + 1 * 0.785 * pbb(age)
    End If
    
    Const nnyear = 30  'Antal arbetsår
    'Dim Boyear As Long ': Boyear = 40 'Bosättningskriteriet finns inte i TP utan fanns i ATP
    If fodar > 1937 Then 'Bosättningskriterie är uppfyllt
        If (nyear < nnyear) Then ATP = ATP * nyear / nnyear 'Om antal arbetade år färre än 30, reduceras TP proportionerligt
        If nyear < 3 Then ATP = 0 'Dock minst tre arbetsår
    End If
    'If PAR > 70 bortse från tid efter 70, görs senare kolla först # månader
    Dim month_2 As Long  ' Antal månader efter pensionsåldern t.ex. 65,2 ->2,4-> 2
    If PAR > 70 Then
        month_2 = 0
    Else
        month_2 = Int(12 * (PAR - Int(PAR)))
    End If

    '(A)TP beräknad
    Dim faktor As Double   'Faktor som används vid justering vid andra uttaget eller ändrad uttagsandel
    Dim andel_2 As Double
    Dim faktor2 As Double

    'AT-Pensionen och upp/nedräkningsfaktor vid första uttaget faktor=# månader * faktor
    If Int(PAR) < 65 Then
        faktor = (1 - 0.005 * (12 * (65 - Int(PAR)) - month_2))
    ElseIf Int(PAR) = 65 Then
        faktor = (1 + 0.007 * month_2)  'OBS Ökningsfaktorn om 0,7 procent ska ersättas med 0,6 procent för uttag som avser tid före den 1 juli 1990, detta bortses från
    ElseIf Int(PAR) <= 70 Then
        faktor = 1 + 0.007 * (12 * (Int(PAR) - 65) + month_2)
    Else 'Efter 70
        faktor = (1 + 0.007 * (12) * (70 - 65))
    End If
   
    ATP = ATP * faktor

''    'Den ackumulerade (produkten) av Balanstalen ska till vid pensioneringen(AckBal)
''    Dim Inx As Double 'Hel AT-pensionen omräknad med Föjsamhetsindex efter pensioneringen
''    If Int(par) < Int(Def_ar) Then
''        If alder = Int(par) And alder < 65 Then
''            Inx = 1
''            ATP = ATP * Inx
''        ElseIf alder < 65 Then
''            'Om pensionering och ålder före 65
''            Inx = pbb(age) / pbb(Int(par)) 'Förändring i prisbasbelopp
''            ATP = ATP * Inx 'Prisbasbeloppet
''        ElseIf par < 65 And Int(alder) = 65 Then
''            'Pensionering innan 65 och ålder 65
''            Inx = pbb(alder) / pbb(Int(par))
''            ATP = ATP * Inx
''
''            Inx = Pindex(age) / Iindex(age)
''            ATP = ATP * Inx
''        ElseIf Int(par) = Int(alder) Then
''            'Pensionering 65 eller senare och första uttaget
''            Inx = Pindex(age) / Iindex(age)
''            ATP = ATP * Inx
''        ElseIf alder > Int(par) Then
''            'efter 65
''            Inx = Pindex(age) / Pindex(age - 1)
''            ATP = ATP * Inx  'Följsamhetsindexeras efter 65
''        End If
''    End If
   
''    tp_ = ATP * andel_ 'Andel avser uttaget hel eller partiell uttag

''    'För nästa uttag vid partiellt
''    Dim month_ As Long '# pensionsmånader under första året
''    Dim month2 As Long
''    month_ = 12 - Int(12 * (born + par - Int(born + par))) 'OBS samma definition
''    month2 = Int(12 * (Def_ar - Int(Def_ar))) 't.ex. 65,3 ger 3 månader
''    ' Month_2=Int(12 * (par - Int(par))) t.ex. 65,8 ger 10 månader
''    If alder = Int(Def_ar) Then
''        'Upp/nedräkningsfaktor vid nästa uttag, if andel_<>andel_2 se vägledning 2010:3 sid 52...
''        '-Faktor 1 # månader mellan 1:a och 2:a uttaget  * ned/uppräkningsprocent 0,5 resp 0,7%
''        '+Faktor 2 # månader efter 2:a uttaget juster för ned/uppräkningsprocent 0,5 resp 0,7%
''        'Ökningsfaktorn om 0,7 procent ska ersättas med 0,6 procent för uttag som avser tid före den 1 juli 1990, detta bortses från
''        If Int(Def_ar) < 65 Then
''            faktor = 0.005 * (12 * (65 - Int(par)) - month_2)
''            faktor2 = 0.005 * (12 * (65 - Int(Def_ar)) - month_2)
''        ElseIf Int(Def_ar) = 65 Then
''            faktor = 0.005 * (12 * (65 - Int(par)) - month_2)
''            faktor2 = 0.007 * month_2
''        ElseIf Int(Def_ar) > 65 And Int(Def_ar) < 70 Then
''            If Int(par) < 65 Then faktor = 0.005 * (12 * (65 - Int(par)) - month_2) _
''                                    + 0.007 * (12 * (Int(Def_ar) - 65) - month2)
''            If Int(par) >= 65 Then faktor = 0.007 * (12 * (Int(par) - 65) - month2)
''            faktor2 = 0.007 * (12 * (Int(Def_ar) - 65) - month_2)
''        Else
''            If Int(par) < 65 Then faktor = 0.005 * (12 * (65 - Int(par)) - month_2) + 0.007 * (12 * (70 - 65))
''            If Int(par) >= 65 Then faktor = 0.007 * (12 * (70 - Int(par)) - month_2)
''            faktor2 = 0.007 * (12) * (70 - 65)
''      End If
''
''      'Definitift uttag, dvs andra ... fungerar även för ändrat uttag
''      andel_2 = 1 'Antas gå till hel pension, skulle kunna i senare läge vara en ny andel
''
''      If alder = Int(Def_ar) Then
''            month_ = 12 - Int(12 * ((fodar + Def_ar) - Int(fodar + Def_ar)) + 0.49) 'Antal pensionsmånader under året
''            tp_ = (tp_ * (12 - month_) _
''                + month_ * ATP * (1 + andel_2 * faktor2 - faktor * andel_)) / 12 'Korrigering för tidigare uttag, kan avse en del av kalenderåret
''      Else
''            tp_ = (ATP) * (1 + andel_2 * faktor2 - faktor * andel_)
''      End If
''    End If

''    'Utbet. under året ordnas i huvudprogrammet
''    If alder = Int(par) Then tp_ = tp_ * month_ / 12

    tp_ = ATP * andel_ * faktor

''    If marginal = 0 Then tp_ = 12 * Int(tp_ / 12 + 0.5) 'Avrundningen efter andelen

End Function

'''' Äldre funktion som byggde på årligt anrop, modellen förenklad
'''''Function tp_(ByVal poang As Double, ByVal nyear As Long, ByVal CIV As Integer, ByVal par As Double, ByVal fodar As Double, ByVal alder As Integer, _
'''''  Optional ByVal andel_ = 1, Optional ByVal Def_ar = -99, Optional ByVal year = 2012, Optional ByVal marginal = 0) As Double
'''''
'''''     'Application.Volatile
'''''    'Poäng  - ATP medelpoäng
'''''    'Nyear  - Antal år med ATP poäng, förvärvskriteriet
'''''    'Civ    - Civilstånd 0-ogift 1-gift
'''''    'Par    - Pensionsålder vid första uttaget ÅÅ,åå
'''''    'fodar  - Födelseår, ÅÅÅÅ,åå
'''''    'Alder  - Ålder
'''''    'Andel  - Uttagsdel (uttagIP)
'''''    'def_ar - Ålder vid det definitiva pensioneringen
'''''    'year   - Inkomstår
'''''    'Marginal - avrundningar
'''''
'''''     'OBS räknat på dagens tilläggspension, historiskt kan parametrarna ha ändrats
'''''     tp_ = 0
'''''    'PAR- Ålder vid pensionering
'''''    If alder < Int(par) Then Exit Function
'''''    If fodar > 1953 Then Exit Function
'''''
'''''''    If Int(par) < 61 Then
'''''''        'MsgBox ("Ingen ATP före 61 års ålder")
'''''''        GoTo slut
'''''''    End If
'''''    'Angivit definitift år
'''''    If IsNumeric(Def_ar) = False Then Def_ar = -99
'''''    If Application.Range("UttagIP").Value = 1 Then Def_ar = -99 'par
'''''
'''''    If par >= Def_ar Then
'''''        andel_ = 1 'Ej angivit definitift år så vid par medför full pension
'''''        Def_ar = 999
'''''    End If
'''''
'''''    Dim ATP As Double   'ATP-delen + FP-delen  beräknas
'''''    If CIV = 0 Then 'HEL TP Inkl FP delen - obs civilståndsbyte kräver en viss handpåläggning..
'''''        ATP = 0.6 * poang * pbb(age) + 1 * 0.96 * pbb(age)
'''''    Else
'''''        ATP = 0.6 * poang * pbb(age) + 1 * 0.785 * pbb(age)
'''''    End If
'''''
'''''    Const nnyear = 30  'Antal arbetsår
'''''    'Dim Boyear As Long '  Boyear = 40 'Bosättningskriteriet finns inte i TP utan fanns i ATP
'''''    If fodar > 1937 Then 'Bosättningskriterie är uppfyllt
'''''        If (nyear < nnyear) Then ATP = ATP * nyear / nnyear 'Om antal arbetade år färre än 30, reduceras TP proportionerligt
'''''        If nyear < 3 Then ATP = 0 'Dock minst tre arbetsår
'''''    End If
'''''    'If PAR > 70 bortse från tid efter 70, görs senare kolla först # månader
'''''    Dim month_2 As Long  ' Antal månader efter pensionsåldern t.ex. 65,2 ->2,4-> 2
'''''    If par > 70 Then
'''''        month_2 = 0
'''''    Else
'''''        month_2 = Int(12 * (par - Int(par)))
'''''    End If
'''''
'''''    '(A)TP beräknad
'''''    Dim faktor As Double   'Faktor som används vid justering vid andra uttaget eller ändrad uttagsandel
'''''    Dim andel_2 As Double
'''''    Dim faktor2 As Double
'''''
'''''    'AT-Pensionen och upp/nedräkningsfaktor vid första uttaget faktor=# månader * faktor
'''''    If Int(par) < 65 Then
'''''        faktor = (1 - 0.005 * (12 * (65 - Int(par)) - month_2))
'''''    ElseIf Int(par) = 65 Then
'''''        faktor = (1 + 0.007 * month_2)  'OBS Ökningsfaktorn om 0,7 procent ska ersättas med 0,6 procent för uttag som avser tid före den 1 juli 1990, detta bortses från
'''''    ElseIf Int(par) > 65 And Int(par) <= 70 Then
'''''        faktor = 1 + 0.007 * (12 * (Int(par) - 65) + month_2)
'''''    Else 'Efter 70
'''''        faktor = (1 + 0.007 * (12) * (70 - 65))
'''''    End If
'''''
'''''    ATP = ATP * faktor
'''''
'''''    'Den ackumulerade (produkten) av Balanstalen ska till vid pensioneringen(AckBal)
'''''    Dim Inx As Double 'Hel AT-pensionen omräknad med Föjsamhetsindex efter pensioneringen
'''''    If Int(par) < Int(Def_ar) Then
'''''        If alder = Int(par) And alder < 65 Then
'''''            Inx = 1
'''''            ATP = ATP * Inx
'''''        ElseIf alder < 65 Then
'''''            'Om pensionering och ålder före 65
'''''            Inx = pbb(age) / pbb(Int(par)) 'Förändring i prisbasbelopp
'''''            ATP = ATP * Inx 'Prisbasbeloppet
'''''        ElseIf par < 65 And Int(alder) = 65 Then
'''''            'Pensionering innan 65 och ålder 65
'''''            Inx = pbb(alder) / pbb(Int(par))
'''''            ATP = ATP * Inx
'''''
'''''            Inx = Pindex(age) / Iindex(age)
'''''            ATP = ATP * Inx
'''''        ElseIf Int(par) = Int(alder) Then
'''''            'Pensionering 65 eller senare och första uttaget
'''''            Inx = Pindex(age) / Iindex(age)
'''''            ATP = ATP * Inx
'''''        ElseIf alder > Int(par) Then
'''''            'efter 65
'''''            Inx = Pindex(age) / Pindex(age - 1)
'''''            ATP = ATP * Inx  'Följsamhetsindexeras efter 65
'''''        End If
'''''    End If
'''''
'''''    tp_ = ATP * andel_ 'Andel avser uttaget hel eller partiell uttag
'''''
'''''    'För nästa uttag vid partiellt
'''''    Dim month_ As Long '# pensionsmånader under första året
'''''    Dim month2 As Long
'''''    month_ = 12 - Int(12 * (born + par - Int(born + par))) 'OBS samma definition
'''''    month2 = Int(12 * (Def_ar - Int(Def_ar))) 't.ex. 65,3 ger 3 månader
'''''    ' Month_2=Int(12 * (par - Int(par))) t.ex. 65,8 ger 10 månader
'''''    If alder = Int(Def_ar) Then
'''''        'Upp/nedräkningsfaktor vid nästa uttag, if andel_<>andel_2 se vägledning 2010:3 sid 52...
'''''        '-Faktor 1 # månader mellan 1:a och 2:a uttaget  * ned/uppräkningsprocent 0,5 resp 0,7%
'''''        '+Faktor 2 # månader efter 2:a uttaget juster för ned/uppräkningsprocent 0,5 resp 0,7%
'''''        'Ökningsfaktorn om 0,7 procent ska ersättas med 0,6 procent för uttag som avser tid före den 1 juli 1990, detta bortses från
'''''        If Int(Def_ar) < 65 Then
'''''            faktor = 0.005 * (12 * (65 - Int(par)) - month_2)
'''''            faktor2 = 0.005 * (12 * (65 - Int(Def_ar)) - month_2)
'''''        ElseIf Int(Def_ar) = 65 Then
'''''            faktor = 0.005 * (12 * (65 - Int(par)) - month_2)
'''''            faktor2 = 0.007 * month_2
'''''        ElseIf Int(Def_ar) > 65 And Int(Def_ar) < 70 Then
'''''            If Int(par) < 65 Then faktor = 0.005 * (12 * (65 - Int(par)) - month_2) _
'''''                                    + 0.007 * (12 * (Int(Def_ar) - 65) - month2)
'''''            If Int(par) >= 65 Then faktor = 0.007 * (12 * (Int(par) - 65) - month2)
'''''            faktor2 = 0.007 * (12 * (Int(Def_ar) - 65) - month_2)
'''''        Else
'''''            If Int(par) < 65 Then faktor = 0.005 * (12 * (65 - Int(par)) - month_2) + 0.007 * (12 * (70 - 65))
'''''            If Int(par) >= 65 Then faktor = 0.007 * (12 * (70 - Int(par)) - month_2)
'''''            faktor2 = 0.007 * (12) * (70 - 65)
'''''      End If
'''''
'''''      'Definitift uttag, dvs andra ... fungerar även för ändrat uttag
'''''      andel_2 = 1 'Antas gå till hel pension, skulle kunna i senare läge vara en ny andel
'''''
'''''      If alder = Int(Def_ar) Then
'''''            month_ = 12 - Int(12 * ((fodar + Def_ar) - Int(fodar + Def_ar)) + 0.49) 'Antal pensionsmånader under året
'''''            tp_ = (tp_ * (12 - month_) _
'''''                + month_ * ATP * (1 + andel_2 * faktor2 - faktor * andel_)) / 12 'Korrigering för tidigare uttag, kan avse en del av kalenderåret
'''''      Else
'''''            tp_ = (ATP) * (1 + andel_2 * faktor2 - faktor * andel_)
'''''      End If
'''''
'''''    End If
'''''
'''''    'Utbet. under året
'''''    If alder = Int(par) Then tp_ = tp_ * month_ / 12
'''''    If marginal = 0 Then tp_ = 12 * Int(tp_ / 12 + 0.5) 'Avrundningen efter andelen
'''''
'''''
'''''
'''''End Function



'Sub ATP_koll()
' Rem tp_(poang As Double, nyear As Double, civ As Integer, par As Double, fodar As Double, alder As Double, Optional andel_ = 1, _
'               Optional def_ar = -99, Optional year = 2012, Optional marginal = 0) As Double
'Dim tp As Double
'  tp = tp_(0, 0, 0, 66, 1945, 66, 1, -99, 2011, 0)
'  tp = tp * 0.45
'End Sub

Function pts(ByVal alder As Double, pbb As Double, _
  Optional civ = 0, Optional ATP = 0, Optional year = 2000, Optional marginal = 0, Optional nyear = 30) As Double
    'Pensionstillskottet fanns under det gamla ATP systemet. Gavs till dem utan eller med låg ATP (ATP räknades av)
    'alder - Ålder 31/12
    'PBB   - prisbasbelopp
    'Civ   - 0=Ogift 1=gift samma koefficinet men tidigare...
    'ATP   - ATP inkomst
    'Year  - Inkomstår
    'Marginal 0 ger Avrudningar
    'Nyear - antal år med förvärvsinkomst
    
     'Application.Volatile
     Const koeff = 0.569 'Samma för både gifta och ogifta, historiskt 0.555?
     'If civ = 0 Then Koeff = 0.555
      pts = koeff * pbb - ATP
      If alder > 60 And alder < 65 Then pts = pts - 0.005 * (65 - alder) * pbb
      If marginal = 0 Then pts = Int(pts / 12 + 0.5) * 12 'Närmaste hela krontal per månad, osäker på avrundningen
      If pts < 0 Then pts = 0
      If nyear < 3 Then pts = 0
End Function


Function fnorm(ByVal year As Long) As Double
    'Följsamhetsnormen kan ses som en konstant men år 2000 gjordes ett undantag.
    'Year - Inkomstår
     'Application.Volatile
     fnorm = 1.016     'Följsamhetsindex startade år 2000 med ett lägre tal (=>högre omräkning)
     If Int(year) = 2000 Then fnorm = 0.996
     If Int(year) < 2000 Then fnorm = 1
End Function

Sub testDeltal()
    Debug.Print deltal(66, 1970, 83, 70)
End Sub
Function deltal(ByVal PAR As Double, ByVal fodar As Double, ByVal alder As Long, Optional def_ar, Optional kolumn = 4) As Double
    'Beräkning av delningstalen kolumn 4 ger IP och
    'PAR    - ålder för första pensionsbeslut ÅÅ,åå
    'FODAR  - födelse år ÅÅÅÅ,åå(obs kan vara 1959.83 för oktober)
    'Alder - Ålder 31/12
    'Def_ar - ålder vid definitivt utträde.
    'Application.Volatile
    
    Dim apar As Double
    Dim konst As Integer
    Dim rad As Double     'Long
    Dim Kol As Double     'Long
    Dim month_1 As Double 'Long
    Dim month_2 As Double 'Long
    'Dim deltal As Double
    deltal = 0
    If fodar <= 1937 Then
        rad = Int(fodar) - 1930 + 9
        If kolumn = 4 Then
            deltal = (wsNyckelTal.Cells(rad, 74))
        Else
            deltal = (wsNyckelTal.Cells(rad, 75))
        End If
        Exit Function
    End If
    If PAR > def_ar Then def_ar = PAR
    If def_ar > 99 Then def_ar = PAR
     
    apar = Application.Range("UttagIP")
    If kolumn = Application.Range("rng_PP_deltal_Year").Column Then
        apar = Application.Range("UttagPP") '----OBS referens som är "låst" ------
        kolumn = Application.Range("rng_PP_deltal_Year").Column + 1
    End If
    
    'If Kolumn = 19 Then apar = Application.Range("UttagPP") '----OBS referens som är "låst" ------
    'If (def_ar > par And apar = 0) Then par = def_ar
       
    konst = 0
    If Int(fodar + PAR + 1 / 1000) > (Int(fodar) + Int(PAR)) Then konst = 1
      
    'Hämta från delningstalsmatrisen som startar på rad 5
    rad = Int(fodar) - 1938 + 5
    'Om fodar > sista raden?
      'Läsa rätt kolumn?
      '  4-IP
      ' 19-PP antingen fond eller trad
      ' 98-pp fond
      '145-pp trad
      
     
      If kolumn = 19 Then kolumn = wsNyckelTal.Range("rng_PP_deltal_Year").Column + 1
      Kol = Int(alder - konst) - 61 + kolumn 'IP 6100 4 och PP 19, Kolumn S motsvarar tex 19:e kolumnen ...
      'If Kolumn = 4 And alder > 72 Then kol = Int(73 - konst) - 61 + 4
      
      month_1 = (PAR - Int(PAR)) * 12        'Antal månader vid partiellt uttag
      month_2 = (def_ar - Int(def_ar)) * 12  'Antal månader vid "slutgiltiga" uttag
      month_1 = 12 - Int(month_1)            'Vägda talet som avrundas nedåt, samma som i fliken nyckettal med delningtal då man går den 1:a i månaden
      month_2 = 12 - Int(month_2)
      
      If PAR <= 60 Then
            deltal = (wsNyckelTal.Cells(rad, 28) - wsNyckelTal.Cells(rad, 29)) * (61 - PAR) + wsNyckelTal.Cells(rad, 28)
      ElseIf alder < Int(PAR + konst) Then
            deltal = 0
      ElseIf alder = Int(PAR + konst) Then
            deltal = (wsNyckelTal.Cells(rad, Kol)) * (month_1 / 12) + (wsNyckelTal.Cells(rad, Kol + 1)) * ((12 - month_1) / 12)
      ElseIf PAR = def_ar Then
            deltal = (wsNyckelTal.Cells(rad, Kol))
      ElseIf alder < Int(def_ar + 1) Then
            deltal = (wsNyckelTal.Cells(rad, Kol))
      ElseIf alder = Int(def_ar + 1) Then
            deltal = (wsNyckelTal.Cells(rad, Kol)) * (month_2 / 12) + (wsNyckelTal.Cells(rad, Kol + 1)) * ((12 - month_2) / 12)
      Else
            deltal = (wsNyckelTal.Cells(rad, Kol))
      End If
        'deltal = month_1
        'If marginal=0 then
    deltal = Int(deltal * 100 + 0.4999) / 100
slut:
End Function

'Sub kolldel()
'' deltal(par, fodar, alder, Optional def_ar, Optional kolumn = 4) As Double
 'Dim koll As Double
 'koll = deltal(65, 1990, 65, 99, 4)
 'koll = 19.36508
 ' koll = Int(koll * 100 + 0.5) / 100
'End Sub

Function P_uttag(ByVal PAR As Double, ByVal alder As Double, _
    Optional def_ar = 999, Optional inr = 1) As Double
    'Hjälp för att kolla uttagsnivån, 100, 50, eller 25%, förenklat eftersom bara en uttagsnivå och sedan definitift.
    'Används av EXCEL arket brutto Kolumn AR och andelsuttag.
     'Application.Volatile
    If PAR > def_ar Then def_ar = PAR
    If def_ar > 99 Then def_ar = PAR
    
    Dim VAL_ As Double
    'Dim P_uttag As Double ''Edited 2024-01-23 - did not compile
    'Def_ar = Worksheets("Brutto").Range("w5")
    
    VAL_ = Application.Range("UttagIP")
    
    If inr = 0 Then VAL_ = Application.Range("UttagPP")
    Select Case alder
       Case Is < Int(PAR)
           P_uttag = 0
       Case Is < Int(def_ar)
           P_uttag = VAL_
       Case Else
           P_uttag = 1
    End Select
    'P_uttag = VAL_
End Function

'Sub kollb()
' Dim kollb As Double
' kollb = P_uttag(65, 66, 68, 6)
'End Sub


Function IP_(ByVal year As Long, ByVal PAR As Double, ByVal born As Double, ByVal pratt As Double, ByVal arvsf1 As Double, _
            ByVal arvsf2 As Double, ByVal kostf As Double, ByVal pbh_ing As Double, ByVal andel As Double, _
            ByVal pens As Double, ByVal deltal As Double, ByVal index As Double, _
            Optional ByVal def_ar = 999, Optional marginal = 0, Optional Typ = 0, _
            Optional Bindex = 0) As Double

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
    
    If born < 1938 Then GoTo slut
    
    Dim arv  As Double 'Typ 1 arvsvinster
    Dim Inx  As Double 'Typ 2 Indexering
    Dim kost As Double 'Typ 3 Kostnadsavdrag
    Dim pbh  As Double 'Typ 4 Pensionsbehållning
    Dim pension  As Double
    arv = 0: Inx = 0: kost = 0: pbh = 0: pension = 0:

    'Dim PRATTU As Double 'From 2015 pratt minskas med kvoten mellan balans- och inkomstindex
    'PRATTU = pratt
    If year > 2014 And (Bindex < 1 And Bindex > 0) Then pratt = pratt * Bindex
    
    Dim konst As Long  'Korrigering för inkomstår vid PAR 0 eller +1
    Dim konst2 As Long 'Korrigering för inkomstår vid Def_ar 0 eller +1
    Dim alder As Long  'Ålder vid årets slut (31/12)
    konst = 0
    konst2 = 0
    If Int(born + PAR + 1 / 1000) > (Int(born) + Int(PAR)) Then konst = 1
    If Int(born + def_ar + 1 / 1000) > (Int(born) + Int(def_ar)) Then konst2 = 1
    
    alder = year - Int(born) 'Ålder vid årets slut.
    
    Dim month As Double      'När under året sker första pensioneringen, dvs antal månder som pensionär under första året som pensionär
    Dim pyear As Double      'Pensionsår = Födelseår + pensionsålder
    pyear = born + PAR
    month = 12 - Int(12 * (born + PAR - Int(born + PAR))) 'OBS samma definition
 
    'Är par och def_ar konsekventa?
    'If Worksheets("Start").Range("UttagIP").Value = 1 Then def_ar = par
    'If Worksheets("Start").Range("UttagIP").Value = 0 Then par = def_ar
    If PAR > def_ar Then def_ar = PAR
    If def_ar > 100 Then def_ar = PAR
     
    Dim FINX As Double    'Följsamhetsindex
    FINX = index / 1.016 '/ fnorm(year)  'Följsamhetsindexeringen, är 1,016 men 2000 var talet 0.996
    If year = 2000 Then FINX = index / 0.996
  
'''If year > 2022 And Worksheets("Start").Range("L77") > 1 Then FINX = index / 1.008

    Dim underlag As Double 'Underlag för återläggning till utg PBH map ev avrundningar: Pension(t-1)*(Index/1.016)*delningstal(t)
    underlag = 0
    'Konst om pensioneringen sker under året eller året efter, tex 1959,8+65,2 ger pension 1/1 året efter
    If alder < Int(PAR + konst) Then
          arv = (arvsf1 - 1) * pbh_ing  'Arvsvinster fås bara under förvärvstiden
          'If alder >= riktl Then arv = arv * arvsf2 + (pbh_ing + pratt) * (arvsf2 - 1)
          arv = arv * arvsf2 + (pbh_ing + pratt) * (arvsf2 - 1)
          Inx = (arv + pratt + pbh_ing) * (index - 1)
          kost = (Inx + arv + pratt + pbh_ing) * (kostf - 1)
          pbh = (pratt + arv + Inx + kost + pbh_ing) 'omräkning till Utgåendebehållning
         
    ElseIf alder = Int(PAR + konst) Then   'Om uttaget sker mitten av året(month>0) eller partiellt uttag (0<andel<1), månadsvis omräkning, PAR>60
          'arv = (arvsf2 - 1) * (pbh_ing + pratt) + andel * (arvsf2 - 1) * (pbh_ing + pratt) * (12 - month) / 12 'Tidigare formel
          arv = (arvsf2 - 1) * pbh_ing * (12 - month) / 12 + pratt * (arvsf2 - 1)
          
          'Inx = (arv + pratt + pbh_ing) * (index - 1) + ((index - 1) * 0.016 * pbh_ing / 1.016) * (12 - month) / 12 'tidigare
          Inx = (arv + pratt) * (index - 1) + (index * 0.016 * pbh_ing / 1.016) * (12 - month) / 12
          kost = (Inx + arv + pratt + pbh_ing) * (kostf - 1)
          If Typ = 9 Then deltal = 1
          If deltal > 0 Then pension = pbh_ing * andel / deltal   'Pension avser pension hela året, ingen korr ännu för antal månader
          pbh = pbh_ing * (1 - andel) + (pratt + arv + Inx + kost)
         
    ElseIf (alder > Int(PAR + konst) And alder < Int(def_ar)) Then  'Andra uttaget och partiellt ELLER prätt
           arv = (arvsf2 - 1) * pbh_ing + (arvsf2 - 1) * pratt
           'Inx = (arv + pratt + pbh_ing) * (index - 1)
           Inx = (arv + pratt) * (index - 1) + (index * 0.016 * pbh_ing / 1.016) * (12 - month) / 12
           kost = (Inx + arv + pratt + pbh_ing) * (kostf - 1)
           
           If pbh_ing > 0 Then
                If (alder + 1) > Int(PAR + konst) Then month = 12 '2:året efter
                If deltal > 0 Then pension = (pbh_ing + pens * FINX * deltal * 12 / month) * andel / deltal
                'Pensionen är = PBH*andel/deltal, OBS första deltal är för åldern i januari och den andra ska vara vid uttaget förenklat och ser till samma tal.
                underlag = pens * FINX * deltal * (12 / month)
                If marginal = 0 Then underlag = Int(Int(pens * FINX + 0.49) * deltal + 0.49)
                pbh = (pbh_ing + underlag) * (1 - andel) + (pratt + arv + Inx + kost)
                
           Else
                pension = pens * FINX * (12 / month) * deltal 'Ska ej förekomma men...
                pbh = 0
           End If
           
           month = 12
           
     ElseIf alder = Int(def_ar + konst) Then 'första uttaget vid slutgiltigt
        'If age = 70 Then deltal = 16.57 For test
          andel = 1
          arv = (arvsf2 - 1) * (pbh_ing + pratt)
          'Inx = (arv + pratt + pbh_ing) * (index - 1) + ((index - 1) * 0.016 * pbh_ing / 1.016) * (12 - month) / 12
          Inx = (arv + pratt) * (index - 1) + (index * 0.016 * pbh_ing / 1.016) * (12 - month) / 12
          kost = (Inx + arv + pratt + pbh_ing) * (kostf - 1)
          If deltal > 0 Then pension = (pbh_ing + pens * FINX * deltal * (12 / month)) * andel / deltal   'HEL Pensionen är = PBH*andel/deltal
          underlag = pens * FINX * deltal
          If marginal = 0 Then underlag = Int(Int(pens * FINX + 0.49) * deltal + 0.49)
          pbh = (pbh_ing + underlag) * (1 - andel) + (pratt + arv + Inx + kost)
     
     ElseIf alder > Int(def_ar + konst) Then 'pensionering men
          andel = 1
          If alder > Int(def_ar + konst + 1) Then month = 12 '
          If pbh_ing > 0 Or pratt > 0 Then
                If deltal > 0 Then _
                 pension = (pbh_ing + pens * FINX * deltal * (12 / month)) * andel / deltal   'HEL Pensionen är = PBH*andel/deltal
                underlag = pens * FINX * deltal * (12 / month)
                If marginal = 0 Then underlag = Int(underlag + 0.49)
                'pbh = (pbh_ing + underlag) * (1 - andel) + (pratt + arv + Inx + kost)
                pbh = (pbh_ing + underlag) + (pratt + arv + Inx + kost)
                If deltal > 0 Then pension = pbh / deltal
                pbh = pbh * (1 - andel)
           ElseIf alder <= 65 Then
                underlag = pens * FINX * deltal * (12 / month) 'Deln tal för åldern januari
                If marginal = 0 Then underlag = Int(Int(pens * FINX + 0.49) * deltal + 0.49)
                If deltal > 0 Then pension = (underlag) / deltal
           Else
                pension = pens * FINX ' (12 / month)
          End If
          month = 12
     End If
     
    If pbh < 0 Then pbh = 0
    If marginal = 0 Then
       pbh = Int(pbh)
       pension = Int(pension / 12 + 0.5) * 12 'Helårs pensionen avrundas närmaste till heltals krona per månad
    End If
    'Korr för antal månader under året för pension
    pension = pension * (month / 12)
    
slut:

    If Typ = 1 Then
    IP_ = arv
    ElseIf Typ = 2 Then
    IP_ = Inx
    ElseIf Typ = 3 Then
    IP_ = kost
    ElseIf Typ = 4 Then
    IP_ = pbh
    Else
    IP_ = pension
    End If

End Function
            
'Year    - Inkomstår
'Par     - pensionsålder slutgiltiga uttaget, ÅÅ,åå
'Born    - födelseår, ÅÅÅÅ,åå
'Pratt   - Pensionsrätt
'arvsf   - Arvsvinstfaktor
'Inx     - förräntningen eller Indexeringen 2
'Kostf   - förvaltningskostnaden
'PBH_ing - Ingående Pensionsbehållningen, 31/12 t-1 samt justering för återtagande och eller överförda rätter
'Andel   - uttagsdel vid pensioneringen
'def_ar  - Definitiv pensionsavgång (100%)
'Pens    - Inkomstpensionen året innan
'Delt    - Aktuellt delningstal
'Index   - Inkomstindex
'Marginal- Avrundning och Oavsett kohort
'Typ     - 0-4; 0-Pension, 1-Arv, 2=Index, 3-Kostnad, 4=PBH
        Rem  IP_(year, par, born, pratt, Arvsf1, arvsf2, Kostf, pbh_ing, andel, pens, deltal, _
             Index,  Optional def_ar = 999, Optional marginal = 0, Optional typ = 0)
'''Sub kollar()
'''    Dim kollar As Double
'''    Dim year As Long
'''    Dim par As Double
'''    Dim born As Double
'''    Dim pratt As Double
'''    Dim arvsf1 As Double
'''    Dim arvsf2 As Double
'''    Dim kostf As Double
'''    Dim index2 As Double
'''    Dim index1 As Double
'''    Dim Bindex As Double
'''    Dim pbh As Double
'''    'IP_(year, par, born, pratt, Arvsf1 ,arvsf2 ,Kostf , pbh_ing, andel, pens, deltal , index As Double, _
'''               def_ar = 999, marginal= 0, typ = 0)
'''    Dim Age As Integer
'''    Age = 67
'''    par = 65
'''    born = 1973
'''    year = born + Age '2018
'''
'''    pratt = Worksheets("Brutto").Range("N" & Age).Value
'''    arvsf1 = Worksheets("Brutto").Range("U" & Age).Value
'''    arvsf2 = Worksheets("Brutto").Range("V" & Age).Value
'''    kostf = Worksheets("Brutto").Range("S" & Age).Value
'''
'''    index2 = Worksheets("brutto").Range("R" & Age + 1).Value
'''    index1 = Worksheets("brutto").Range("R" & Age).Value
'''    Bindex = Worksheets("brutto").Range("Q" & Age).Value
'''
'''    pbh = Worksheets("Brutto").Range("AB" & Age - 1).Value
'''    kollar = IP_(year, par, born, pratt, arvsf1, arvsf1, kostf, pbh, 1, 187512, 17.16, (index2 / index1), 999, 0, 0, Bindex)
'''    MsgBox (kollar)
'''
'''End Sub


Function ppkassa(ByVal PAR As Double, ByVal born As Double, ByVal alder As Long, ByVal pbh As Double, _
    Optional def_ar = 999, Optional andel = 1, Optional marginal = 0) As Double
    
    'Application.Volatile
 
    'Årlig omräkningen av premiepensionen
    'Par    - Pensionsålder vid första uttaget, ÅÅ,åå
    'Born   - Födelseår, ÅÅÅÅ,åå
    'Alder  - Ålder vid årets slut
    'pbh    - Pensionsbehållningen
    'def_ar - Definitiv pensionsålder (vid partiellt uttag), ÅÅ,åå
    'Andel_ - Uttagsandel: 1, 0.75, 0.50 eller 0.25.
    'Marginal - 0 nuvarande regler, 1 Inga avrundningar (och det nya systemet i sin helhet)
    
    'If andel <> 1 Or andel <> 0.75 Or andel <> 0.5 Or andel <> 0.25 Then andel = 1
    
    'Alternativt att läsa in dem direkt
    ' born = Worksheets("Start").Range("Born").Value
    ' par = Worksheets("Start").Range("Par").Value
    ' def_ar = Worksheets("Start").Range("Def_ar").Value
    If PAR > def_ar Then def_ar = PAR
    If def_ar < 61# Then def_ar = PAR
    If def_ar > 99 Then def_ar = PAR
    
    Dim konst As Single
    Dim konst2 As Single
    konst = 0 'För att fånga rätt år vid PAR
    konst2 = 0 'För att fånga rätt år vid def_ar
    If Int(born + PAR + 1 / 1000) > (Int(born) + Int(PAR)) Then konst = 1
    If Int(born + def_ar + 1 / 1000) > (Int(born) + Int(def_ar)) Then konst2 = 1
    
    Dim month As Long 'Month = antal pensionsmånader under året
    
    If born < 1938 Then GoTo slut
    
    Dim del_tal As Double
    ppkassa = 0
    del_tal = 0
    'Delvis uttag ... ges av p_uttag
    'andel_ = P_uttag(par, alder, defar, 7)
    
    If andel > 1 Then andel = 1
    If andel < 0.1 Then GoTo slut
    
    If alder < 101 Then
        del_tal = deltal(PAR, born, alder, def_ar, Application.Range("rng_PP_deltal_Year").Column)  ' Ålder vid årets slut. Deltal fixar ålder vid PAR
    Else: del_tal = 2
    End If
    
    If pbh < 0 Then pbh = 0
    
    Select Case alder
    Case Is < Int(PAR + konst) 'Inträffar inte men ..
       ppkassa = 0
    Case Is = Int(PAR + konst)
       ppkassa = andel * pbh / del_tal
       month = 12 - Int(12 * (born + PAR - Int(born + PAR))) 'OBS samma definition
       'month = month * Application.Range("UttagPP").Value    'Antal hela pensionsmånder
       If alder = Int(def_ar + konst2) Then
          month = month + (12 - 12 * ((born + def_ar) - Int(born + def_ar + 1 / 1000))) * (1 - Application.Range("UttagPP").Value)
          'De första pensionsmånaderna med deltid + #antal månder med heltid *(1-andel) eftersom de första månderna inkluderar deltiden fram till 31/12.
       End If
    Case Is < Int(def_ar + konst)
        ppkassa = pbh / del_tal '*andel
        month = 12 * Application.Range("UttagPP").Value   'Antal hela pensionsmånader
    Case Is = Int(def_ar + konst)
        ppkassa = pbh / del_tal   'Slutgiltiga andelen är 1
        month = 12 * Application.Range("UttagPP").Value
        month = month + (12 - 12 * ((born + def_ar) - Int(born + def_ar + 1 / 1000))) * (1 - Application.Range("UttagPP").Value)
    Case Else
        ppkassa = pbh / del_tal
        month = 12
    End Select
    If month > 12 Then month = 12
    If month < 0 Then month = 0
    
    If marginal = 0 Then
        ppkassa = Int(ppkassa / 12 + 0.5) 'Månadsutbetalningen
    Else: ppkassa = ppkassa / 12
    End If
    ppkassa = ppkassa * month
    
slut:
End Function

'Sub kollaPP()
'  'ppkassa(par As Double, born As Double, alder As Long, pbh As Double, _
'   Optional def_ar = 999, Optional andel_ = 1, Optional marginal = 0)
'  Dim kolla As Long
'  kolla = ppkassa(61, 1990, 61, 100000, 999, 1, 0)
'End Sub


Function pgb_barn(ByVal ar, ByVal Jink, ByVal Uink, ByVal barn, ByVal Parent, ByVal medel, _
    Optional IBB = 52100, Optional marginal = 0, Optional rikt = 65) As Double
    'PGB för barnår  21§-54. förenklat
    'Ar - inkomstår
    'Barn - när barnet är fött ÅÅÅÅ
    'Parent - ålder på föräldern
    'Medel - medel_pgi, ska kanske läsas in
     'Application.Volatile
     pgb_barn = 0

    If barn > 1959 And Parent < rikt Then
        If ar = barn Or (ar < barn + 4 And ar > barn) Then
            'Metod 1: Individuell jämförelse
             pgb_barn = Jink - Uink
            'Metod 2: Generell jämförelse
             If (0.75 * medel - Uink) > pgb_barn Then pgb_barn = (medel * 0.75 - Uink)
            'Metod 3: Enhetligt belopp
             If 1 * IBB > pgb_barn Then pgb_barn = IBB * 1
             If marginal = 0 Then pgb_barn = Int((pgb_barn + 49) / 100) * 100 'Var för sig avrundas nedåt till närmaste 100 kr
        End If
    End If
    
    'If Parent > 64 Then PGB_barn = 0 'Föräldern försäkrad, haft vårdnaden ...
End Function

Function Balansindex(ByVal BI As Double, ByVal Btal As Double, ByVal ital1 As Double, ByVal ital2 As Double, Optional marginal = 0) As Double
  'Application.Volatile
  
  Select Case BI
    Case Is = 0
        If Btal < 1 Then Balansindex = Btal * ital1
    Case Else
        Balansindex = Btal * BI * ital1 / ital2
   End Select
   If Balansindex > ital1 Then Balansindex = ital1
   If marginal = 0 Then Balansindex = Int(Balansindex * 100 + 0.49) / 100  'Avrundat till 2 dec.
   
End Function
''
''Sub kolla_bi()
''    Dim k As Double
''    k = Balansindex(170.73, (1.00813 - 1) / 3 + 1, 175.96, 170.73)
''
''End Sub

