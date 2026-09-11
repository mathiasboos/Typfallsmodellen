Attribute VB_Name = "VBA_go"
'Huvudprogram för körning och beräkningar i typfallsmodellen
'Först definieras några storheter som är gemensamma eller publika
'såsom KPI mm (startsetup) och livsinkomsterna och pensionsspecifika variabler mm(mcalc)

Option Explicit                     'Alla variabler ska definieras
Option Base 1                       'Indexering börjar med 1 och inte med noll

Public startage As Integer          'Startålder, anges i fliken adv_settings (=15)
Public Const slutage = 105          'Sista eller Slutålder för beräkningar
                                                   
Public born As Double               'Född YYYY,yy
Public PAR As Double                'Allmän pensionsålder XX,xx
Public tjp_par As Double            'Pensionsålder XX,xx tjänstepensionen (TJ
Public def_ar As Double             'Definitiv pensionsålder för allmän pension vid partiellt uttag

Public andelnya                     'Andel av det nya pensionssystemet och 1-andel för (A)TP-systemet
'Public Ut_andel_ip As Double       'Uttagsandel inkomstpensionen
'Public Ut_andel_pp As Double       'Uttagsandel premiepensionen


Public growth As Double             'Real tillväxt
'Public Inflation As Double          'Inflation - hanteras utanför VBA

Public age As Integer               'Ålder 31/12

Public year_() As Long              'Inkomstår (= born+age)

Public W_start As Double            'Ålder XX,xx när individen börjar arbeta,
                                    'OBS born 19xx,yy+XX,xx kan innebära att inkomståret blir +1 om ,yy +,xx>1
Public Income As Double             'Angiven inkomst
Public avtal As Integer             'Avtalsområde 1-5:
                                    '1:Saknar tjp, 2:SAF_LO 3:ITP 4:(A)KAP-KL 5:PA16
    
Public Income_() As Double          'Inkomst under livet vid olika åldrar
Public Wage_() As Double            'Löneinkomst under livet vid olika åldrar

Public egen_W As Boolean            'Egen lönevektor i stället för s.k. rak löneprofil

Public w_time As Double             'Referens lön avseende vid ålder
Public w_ref As Long                'Referens år för fasta priser
Public Nominal As Boolean           'Referenslönen i nominella termer (sant) eller fasta priser (falskt)

Public IBB() As Double              'Inkomstbasbelopp (IBB)
Public pbb() As Double              'Prisbasbelopp (PBB)
Public FPB() As Double              'Förhöjt prsibasbelopp

Public KPI() As Double              'KPI-årsmedeltal
Public KPI_j() As Double            'KPI-juni

Public Iindex() As Double           'Inkomstindex
Public Pindex() As Double           'Inkomst/Balansindex

Public yield() As Double            'Fondavkastning
Public RGK() As Double              'Avkastningen hos riksgälden för tillfälliga förv.

Public MPGI() As Double             'Medel PGI

Public IP_avg() As Double           'Avgifter för inkomstpensionen
Public PP_avg() As Double           'Avgifter för Premiepensionen
Public TP_avg() As Double           'Avgifter för tjänstepension

Public IP_arv1() As Double          'Arvsvinstfaktor 1, 1-60 IP
Public IP_arv2() As Double          'Arvsvinstfaktor 2, 60-w IP
Public PP_arv() As Double           'Arvsvinstfaktor PP, ev. ersätta med Makeham funktion i framtiden

Public pmonth As Integer            'Antal pensionsmånader Pensionen sker den 1:a den månad när individen fyller PAR år
Public Tmonth As Integer            'Antal pensionsmånader tjp
                                                
'Public book As String               'Utskrift till arket "book" används inte
Public marginal As Byte              'Marginalberäkningar

Public gift As Boolean              'Civilstånd 0-Ensamstående/ogift 1-Sammanboende/Gift obs vissa funktioner ensamstående=1
Public Iyear As Single              'Referensår för när text inkomstindexering ska gälla istf prisindex


Public Kom_skatt() As Double        'Kommunalskattesats
Public Begravavg() As Double        'Begravningsavgift + ev. kyrkoskatt (eller annat trossamfund)
Public Tax_limit1() As Double       'Skiktgräns 1 för statlig inkomstskatt
Public Tax_limit2() As Double       'Skiktgräns 2 för statlig inkomstskatt

Public cases As Integer             'Typfall nr vid mikrosim
'Public Run_from_indata As Byte
'Public Const verbose = 1            'För ytterligare utskrift till kalkylbladet utdata: 1=Ja, annars inte
Public verbose As Byte

'Några typfallsspecifika egenskaper samt index mm
'OBS alla vektorer ser till age, tex year_(age) som ger inkomstår
Sub startsetup()
    'OBS att dim innebär "lokal" definition som försvinner efter sub, därav Public ovan
    verbose = 0
    If Application.Range("Verbose") = 1 Then verbose = 1 'added data
    If Application.Range("rng_Egen_Lon") = True Then
        startage = 15
    Else
        startage = Application.WorksheetFunction.min(Application.Range("wStartYear_AdvSettings"), Application.Range("wStartYear"))
    End If
   
     If Application.Range("wStartYear") < Application.Range("wStartYear_AdvSettings") Then _
        Application.Range("wStartYear_AdvSettings") = Application.Range("wStartYear")
    
    ReDim year_(startage To slutage) As Long
    ReDim IBB(startage To slutage) As Double
    ReDim pbb(startage To slutage) As Double
    ReDim FPB(startage To slutage) As Double
    
    ReDim KPI(startage To slutage) As Double
    ReDim KPI_j(startage To slutage) As Double
    
    ReDim Iindex(startage To slutage) As Double
    ReDim Pindex(startage To slutage) As Double
    
    ReDim yield(startage To slutage) As Double
    ReDim RGK(startage To slutage) As Double
    
    ReDim MPGI(startage To slutage) As Double
    
    ReDim IP_avg(startage To slutage) As Double
    ReDim PP_avg(startage To slutage) As Double
    ReDim TP_avg(startage To slutage) As Double
    
    ReDim IP_arv1(startage To slutage) As Double
    ReDim IP_arv2(startage To slutage) As Double
    ReDim PP_arv(startage To slutage) As Double
    ReDim Kom_skatt(startage To slutage) As Double
    ReDim Begravavg(startage To slutage) As Double
    ReDim Tax_limit1(startage To slutage) As Double
    ReDim Tax_limit2(startage To slutage) As Double
    
    ReDim Income_(startage To slutage) As Double
    ReDim Wage_(startage To slutage) As Double
    
    'nypropp = Application.Range("rng_nypropp")
    born = Application.Range("BornYear").Value
    'I äldre modell (PETRA) gavs enbart Born som kunde anges med XXXX,xx här begränsad till XXXX, dvs explicit född 1/1
    Income = Application.Range("Wage_Monthly").Value
    Income = Income * 12 'Årsinkomst
    growth = Application.Range("rng_Real_Growth").Value
    'Inflation = wsStart.Range("rng_Yearly_Inflation").Value
    
    Application.Calculate
    w_time = Application.Range("w_time")    'Inkomsten och vid vilken ålder den avser default (år(nu()) -Born, dvs åldern idag)
    w_ref = Application.Range("w_ref")      'Om fasta priser vilket år avser Default idag dvs =ÅR(nu())
    Nominal = Application.Range("Nominal")  'Angiven lön i nominella (sant) eller fasta (Falskt) priser

    Dim avkast As Double
    avkast = Application.Range("rng_FondAvkastning").Value '-> Yield(age)
    
    andelnya = andel(born) 'Andel av ATP
    
    PAR = Application.Range("ParYear").Value 'När (ålder) typisen går i allmän Pension
    
    'koll om pensionering (PAR) är möjlig med hänsyn till riktåldern (nedre)
    Dim riktl As Integer
    riktl = Application.Range("Rng_riktL").Value 'Nedre riktålder - för aktuell årskull riktage(year_(Int(par)), 0)
    
    If (IsMissing(Application.Range("RulesfromUtg")) Or Application.Range("RulesfromUtg") = 0) Then
        'Nothing
        Else
        If Application.Range("Rules") = 1 Then
            If Int(riktl + born) > Application.Range("RulesfromUtg") Then riktl = riktage(Application.Range("RulesfromUtg"), 0)
            Else
              If Int(born) + riktl > Application.Range("RulesfromUtg") Then riktl = riktage(Application.Range("RulesfromUtg"), 0)
              'Kan påverkar arvsvinstfaktorerna
        End If
    End If
   
    'Möjligen ska
    If PAR < riktl Then 'OBS Pensinsonsåldersutredningens förslag om lägsta ålder, riktålder mm
        'par = riktl
        If cases > 0 Then
            'Debug.Print cases; born; par
            cases = 0 'Stäng i mikrosim
        End If
        If Application.Range("sysLang") = 0 Then

            If MsgBox("Allmän pension först möjlig vid " & riktl & " års ålder, räkna pensionsålder vid " & riktl, vbYesNo, "Allmän Pension inte möjlig vid " & PAR & " års ålder") = vbYes Then
                Application.Range("PARYear") = riktl
                Application.Range("tjp_par") = riktl
                PAR = riktl
                'Exit Sub
                'blnRikAlderRun = True
            Else
                Unload UserFormProgress
                Application.EnableEvents = True
                Application.Calculation = xlCalculationAutomatic
                Application.ScreenUpdating = True
                Exit Sub
            End If

        Else
            If MsgBox("Pension is possible first at age " & riktl & " change the retirement to " & riktl, vbYesNo, "Public pension is not possible at age " & PAR) = vbYes Then
                Application.Range("PARYear") = riktl
                Application.Range("tjp_par") = riktl
                PAR = riktl
                'Exit Sub
                'blnRikAlderRun = True
            Else
                Unload UserFormProgress
                Application.EnableEvents = True
                Application.Calculation = xlCalculationAutomatic
                Application.ScreenUpdating = True
                Exit Sub
            End If

        End If
        Application.Range("PARYear").Value = PAR
    End If

    tjp_par = Application.Range("tjp_Par").Value
    If tjp_par < 55 Then
        MsgBox ("Tjänstepension först möjlig vid 55 år, räknar pensionsålder vid " & PAR)
        tjp_par = PAR
        Application.Range("tjp_Par").Value = PAR
    End If
    
    def_ar = Application.Range("rng_def_ar").Value 'Definitiv allmän pensionsålder (uttag 100%)
    'Mindre koll av def_ar, görs senare
    If def_ar < PAR Or IsMissing(def_ar) Then def_ar = PAR
    If def_ar > 100 Then
        def_ar = PAR
        If Application.Range("sysLang") = 0 Then
            MsgBox ("Definitivt pensions ålder är inte möjlig/sannolik, ändrat till " & PAR)
        Else
            MsgBox ("Final pension at that age is not poosible, changed to " & PAR)
        End If
    End If
    
    avtal = Application.Range("rng_TJP_Val").Value
    
    W_start = Application.Range("wstartyear") 'Heter wstart tidigare (PETRA) och kunde där anges med YY,yy här bara YY
    If W_start < startage Then 'Ger p-rätter som finns före startage...
        'samla ihop dessa till startage via egen löneserie, eller skjut in PBH.
        MsgBox ("Calculation starts at " & startage & " age ")
        W_start = startage
    End If
    
    egen_W = Application.Range("rng_Egen_Lon") 'Antas att Inkomst = lön
    marginal = Application.Range("Marginal")    'Marginaler
 
    gift = Application.Range("Gift").Value 'True or false
    
    'Kommunal skattesats
    'Om inget anges eller 0 fås Historisk genomsnittlig kommunal skattesats, annars vald skattesats för hela perioden, även historiskt
    Dim hist_skattesats As Byte
    hist_skattesats = 0
    If Application.Range("rng_Kommunalskatt").Value < 0.1 Then hist_skattesats = 1
    
    'Utgå från senaste index? eller prognos
    Dim Yval As Single 'Val av avkastning, 1: samma hela livet, 2: PPM, 3: SÅFan
    Yval = Application.Range("rng_Avkastning_val").Value
    
    Dim risk As Double
    risk = Application.Range("Risk") 'Mäts som spridningen, def. =0
    Dim distribution As Single
    distribution = Application.Range("lognormal").Value
     '1-lognormal annars normal ... obs enbart lognormal'
    
    Dim slump As Double     'Rektangulärt slumptal mellan 0 och 1
    Dim hist_risk As Byte   '1-> Avkastningen varierar även historiskt
    hist_risk = 0
    If Yval = 1 Then hist_risk = 1
    
    'Livslängdsjustering finns i några tal och kolumn AX
''    Dim rdeathrisk As Range
''    Set rdeathrisk = wsTal.Range("Ax15:Ax100") 'obs Fast range
    
    For age = startage To slutage
        year_(age) = Int(born) + age
        'Index-serier
        'Omräkningarna för framtiden sker i arket. här läses några in
        If year_(age) > 1959 Then 'Från nyckeltal eller någratal hämtas, År 1959=rad 6
            KPI_j(age) = wsNyckelTal.Cells(year_(age) - 1959 + 6, 83) 'OBS delvis fasta adresser
            KPI(age) = wsNyckelTal.Cells(year_(age) - 1959 + 6, 84)
            pbb(age) = wsNyckelTal.Cells(year_(age) - 1959 + 6, 86) 'Application.Range("rng_PBB_Kolumn").Column)
            IBB(age) = wsNyckelTal.Cells(year_(age) - 1959 + 6, 88) 'Application.Range("rng_IBB_Kolumn").Column)
            FPB(age) = wsNyckelTal.Cells(year_(age) - 1959 + 6, 89) 'Application.Range("rng_FHB_Kolumn").Column)
            
            MPGI(age) = wsNyckelTal.Cells(year_(age) - 1959 + 6, 87) 'Application.Range("rng_Nyckeltal_Medel_PGI").Column)

            Iindex(age) = wsNyckelTal.Cells(year_(age) - 1959 + 6, 90) 'Application.Range("rng_InkomstIndex_Kolumn").Column) 'Inkomstindex
            Pindex(age) = wsNyckelTal.Cells(year_(age) - 1959 + 6, 93) 'Application.Range("rng_Nycketal_Index_Top").Column) 'Pensionindex
            
            If KPI(age) = 0 Then 'Uppdatering av några tal saknas, i dagsläget efter 2110
                KPI_j(age) = KPI_j(age - 1) * (1 + Application.Range("rng_Yearly_Inflation"))
                KPI(age) = KPI(age - 1) * (1 + Application.Range("rng_Yearly_Inflation"))
                pbb(age) = 36396 * KPI_j(age) / 257.4
                FPB(age) = 37144 * KPI_j(age) / 257.4
                Iindex(age) = Iindex(age - 1) * (1 + Application.Range("rng_real_growth")) * (1 + Application.Range("rng_Yearly_Inflation"))
                MPGI(age) = MPGI(age - 1) * Iindex(age) / Iindex(age - 1)
                IBB(age) = 43313 * Iindex(age) / 118.41
                If marginal = 0 Then 'Avrundas ...
                    pbb(age) = Int(pbb(age) / 100 + 0.5) * 100
                    FPB(age) = Int(FPB(age) / 100 + 0.5) * 100
                    IBB(age) = Int(IBB(age) / 100 + 0.5) * 100
                End If
            End If
            If Pindex(age) < 1 Or Pindex(age) > Iindex(age) Then Pindex(age) = Iindex(age)
            'If marginal = 0 Then Pindex(age) = Round(Pindex(age), 2) sker i Excel
                       
            If age >= startage Then 'obs year_(age)>1959
               
                If Yval = 1 Then
                    If age > startage Then
                        yield(age) = (1 + avkast) * (KPI(age) / KPI(age - 1))
                    Else
                        yield(age) = (1 + avkast) '* (KPI(age) / KPI(age - 1)) 'korrigering
                    End If
                    
                    If Application.Range("rng_Avkastning_fondavgifter").Value = 0 Then _
                    yield(age) = yield(age) - 0.002 'Fondakastning före avgiftsuttag, 0.2% i avgift
                    
                ElseIf Application.Range("Modell_year") > year_(age) Then
                    If Yval = 2 Then
                        yield(age) = 1 + wsTal.Cells(year_(age) - 1959 + 6, 17) 'Fast adress rad 7=1960, kolumn "vald avkastning"
                    Else
                        yield(age) = 1 + wsTal.Cells(year_(age) - 1959 + 6, 18)
                    End If
                Else
                    If Yval = 2 Then
                        yield(age) = 1 + wsTal.Cells(year_(age) - 1959 + 6, 17) 'Fast adress rad 7=1960, kolumn "vald avkastning"
                    Else
                        yield(age) = 1 + wsTal.Cells(year_(age) - 1959 + 6, 18)
                    End If
                    
                    'yield(age) = (1 + avkast) * (1 + Application.Range("rng_Yearly_Inflation")) 'Brutto, Tidigare men om användare har ändrat serien
                    'If Application.Range("rng_Avkastning_fondavgifter").Value = 0 Then _
                    yield(age) = yield(age) - 0.004 'Fondakastning före avgiftsuttag, 0.4% i avgift oavsett...
                    'Avgifterna i medeltal inns i några tal och borde kanske läsas in, se nedan tjp samma som pp och ingen kickback (återbäring) så
                    'yield(age) = yield(age) - wsTal.Cells(year_(age) - 1959 + 6, 20) - wsTal.Cells(year_(age) - 1959 + 6, 21)
                End If
                
            End If 'age>=Startage
         Else 'Före 1959
                'Före antas avkastningen vara nio procent, KPI mm samma som 1960, kommer inte att påverka pensionerna
                yield(age) = 1.09
                KPI_j(age) = 25.39
                KPI(age) = 25.39
                pbb(age) = 4200
                IBB(age) = 4200
                FPB(age) = 4200
                Iindex(age) = 6.54 / (1.06 ^ (1960 - year_(age)))
                'Bindex(age) = Iindex(age)
        End If
        
         'Avkastning -  under risk: PP, TJP och IPS  antas ha samma årliga avkastning
         'Risk över tid obs även mellan individer om samma typfall körs om ?
         If (year_(age) > year(Now()) Or hist_risk = 1) And risk > 0 Then
            'slump = Rnd(1430) 'Över tid och mellan individer men samma vid omkörning!
            slump = Rnd()      'Över tid men inte samma vid omkörning eftersom mikrosim...!
            
            Dim phi As Double
            Dim my, my2  As Double
            Dim sigma, s2 As Double
            my = (1 + Application.Range("rng_FondAvkastning").Value) * (1 + Application.Range("rng_Yearly_Inflation").Value) - 1
            sigma = risk
            'Värdeutv. antas vara lognormalfördelad log(y)= normal... varför my och sigma ska justeras
            If distribution <> 1 Then yield(age) = WorksheetFunction.NormInv(slump, my, sigma) + 1
            If distribution = 1 Then
                phi = (risk ^ 2 + (1 + my) ^ 2) ^ 0.5
                my2 = Log((1 + my) ^ 2 / phi)
                s2 = (Log(phi ^ 2 / (1 + my) ^ 2)) ^ 0.5
                yield(age) = WorksheetFunction.NormInv(slump, my2, s2)
                yield(age) = Exp(yield(age))  'lognormal distributed
            End If
             'Debug.Print cases; Round(yield(age), 5)
         End If

        'Tllfälliga förvaltningen
        If year_(age) > 1995 Then
            RGK(age) = 1 + wsTal.Cells(year_(age) - 1959 + 6, 19) / 100 'FAST adress
        Else
            RGK(age) = 1
        End If
        
        'Kvar efter Avgifter = (1-avgiftsfaktor)
        If year_(age) < 1960 Then
            IP_avg(age) = 1
            Else: IP_avg(age) = wsTal.Cells(year_(age) - 1959 + 6, 14) 'Fast adress
        End If
        
        If year_(age) < 2000 Then
            PP_avg(age) = 1
            Else: PP_avg(age) = wsTal.Cells(year_(age) - 1959 + 6, 15) 'Fast adress
        End If
        TP_avg(age) = PP_avg(age) 'Antas vara lika effektiv som fondtorget
        'Debug.Print "Kvar efter avgifter: "; year_(age); " "; IP_avg(age); " "; PP_avg(age)

        'Arvsvinster = (1+arvsfaktor)
        If age < 17 Or year_(age) < 2000 Then
            IP_arv1(age) = 1: IP_arv2(age) = 1
        Else
            If age < riktl Then 'Året innan därav < och inte <= ovs dubbla rader för 60
                IP_arv1(age) = wsArv_ip.Cells(age - 17 + 3, year_(age) - 1999 + 3) 'Fasta adresser
                If IP_arv1(age) = 0 Then
                    IP_arv1(age) = 1
                    'Debug.Print year_(age); age; "Arv saknas"
                End If
                
                IP_arv2(age) = 1
           
                If age = (riktl - 1) Then _
                    IP_arv2(age) = wsArv_ip.Cells(81 + age - 60, year_(age) - 1999 + 3) 'Dubbla arvsvinsten
               
                If IP_arv2(age) = 0 Then IP_arv2(age) = 1
            End If
            If age >= (riktl) Then
                IP_arv1(age) = 1
                If age >= (riktl) Then _
                    IP_arv2(age) = wsArv_ip.Cells(81 + age - 60, year_(age) - 1999 + 3) 'Dubbla arvsvinsten
                If IP_arv2(age) = 0 Then IP_arv2(age) = IP_arv2(age - 1)
                If IP_arv2(age) = 0 Then IP_arv2(age) = 1
                'If riktage(year_(age - 1), 0) < riktage(year_(age), 0) Then IP_arv2(age) = 1
            End If
        End If

        'Premiepensionen - framtiden ev. gå direkt på Makehamsformeln
        If age < 15 Or year_(age) < 2003 Then
            PP_arv(age) = 1
        Else
            If age < 106 Then 'Arv pp fliken slutar vid 105
                PP_arv(age) = wsArv_pp.Cells(age - 14 + 2, year_(age) - 1999 + 3)
                Else: PP_arv(age) = wsArv_pp.Cells(93, year_(age) - 1999 + 3)
            End If
        End If
        
        If age > 14 And age < 101 Then
            If IP_arv1(age) > 1 Then IP_arv1(age) = IP_arv1(age) '* rdeathrisk(age - 14)
            If IP_arv2(age) > 1 Then IP_arv2(age) = IP_arv2(age) '* rdeathrisk(age - 14)
            If PP_arv(age) > 1 Then PP_arv(age) = PP_arv(age) '* rdeathrisk(age - 14)
        End If
        'Skatter
        If year_(age) <= 1930 And hist_skattesats = 1 Then 'OBS Saknas antas
            Kom_skatt(age) = 0.0844
            Tax_limit1(age) = 10000
            Tax_limit2(age) = 10 ^ 9
        ElseIf year_(age) <= 1958 And hist_skattesats = 1 Then
            Kom_skatt(age) = wsK_Skatt.Cells(year_(age) - 1930 + 2, 2) / 100 'OBS Fast adress
            Tax_limit1(age) = 11000
            Tax_limit2(age) = 10 ^ 9
        ElseIf hist_skattesats = 1 Then '
            Kom_skatt(age) = wsK_Skatt.Cells(year_(age) - 1930 + 2, 2) / 100 'Fast adress
            Tax_limit1(age) = wsNyckelTal.Cells(year_(age) - 1958 + Application.Range("rng_StartRad_Nyckeltal").Row, Application.Range("rng_Tax_limit1").Column) 'Fast adress
            Tax_limit2(age) = wsNyckelTal.Cells(year_(age) - 1958 + Application.Range("rng_StartRad_Nyckeltal").Row, Application.Range("rng_Tax_limit2").Column) 'Fast adress
        Else 'and hist_skattesats=0
            Kom_skatt(age) = Application.Range("rng_Kommunalskatt").Value
            If year_(age) < 1958 Then
                Tax_limit1(age) = Int(11000 / 1.06 ^ (1958 - year_(age))) 'godtyckligt
                Tax_limit2(age) = 10 ^ 9
            Else
                Tax_limit1(age) = wsNyckelTal.Cells(year_(age) - 1958 + Application.Range("rng_StartRad_Nyckeltal").Row, Application.Range("rng_Tax_limit1").Column) 'Fast adress
                Tax_limit2(age) = wsNyckelTal.Cells(year_(age) - 1958 + Application.Range("rng_StartRad_Nyckeltal").Row, Application.Range("rng_Tax_limit2").Column) 'Fast adress
            End If
        End If
        
        If year_(age) < 2000 And hist_skattesats = 1 Then
                Begravavg(age) = 0
            ElseIf hist_skattesats = 1 Then
                Begravavg(age) = wsK_Skatt.Cells(year_(age) - 1930 + 2, 8) / 100 'Fast adress
            Else:
                'and hist_skattesats=0
                Begravavg(age) = Application.Range("rng_Begravningsavgift").Value 'Fast adress
        End If
        
        'Skatten om andra års regler - Fixa felmeddelande om året är före 1958?
        If Application.Range("RulesfromSkatt") = 0 Or IsMissing(Application.Range("RulesfromSkatt")) Then
             'Nothing
        Else
            If (year_(age) > Application.Range("RulesfromSkatt") And Application.Range("Rules") = 0) _
            Or Application.Range("Rules") = 1 Then
            
                Tax_limit1(age) = wsNyckelTal.Cells(Application.Range("RulesfromSkatt") - 1958 + Application.Range("rng_StartRad_Nyckeltal").Row, Application.Range("rng_Tax_limit1").Column) 'Fast adress
                Tax_limit2(age) = wsNyckelTal.Cells(Application.Range("RulesfromSkatt") - 1958 + Application.Range("rng_StartRad_Nyckeltal").Row, Application.Range("rng_Tax_limit2").Column) 'Fast adress
                'Nominella värden Justeras för KPI?
                
            End If
            If hist_skattesats = 1 Then
                Kom_skatt(age) = wsK_Skatt.Cells(Application.Range("RulesfromSkatt") - 1930 + 2, 2) / 100 'Fast adress
                Begravavg(age) = wsK_Skatt.Cells(Application.Range("RulesfromSkatt") - 1930 + 2, 8) / 100 'Fast adress
            End If
            
        End If

        'Taxerade FörvärvsInkomster
        If egen_W = True Then
            Income_(age) = wsEgenInkomst.Cells(age, Application.Range("rng_Typisar_aktuell_Header").Column)
            Wage_(age) = wsEgenInkomst.Cells(age, Application.Range("rng_Typisar_aktuell_Header").Column + 1)
            'Debug.Print age; Income_(age); Wage_(age)
        Else
            Income_(age) = wages(age, year_(age), W_start, Income, w_time, w_ref, Nominal)
            
            DoEvents 'För att Excel ska hinna med i beräkningarna??
            
            If Application.Range("TL_spec_y").Value > 1959 And Income_(age) > 0 Then
                If year_(age) >= Application.Range("TL_spec_y") Then
                    If Application.Range("TL_special") <> 1 Then
                        Income_(age) = Income_(age) * Application.Range("TL_special")
                    End If
                End If
            End If
            
            If Application.Range("satagare") > 1960 Then
                If Application.Range("satagare") <= year_(age) Then
                    Wage_(age) = 0
                Else
                    Wage_(age) = Income_(age)
                End If
                'Debug.Print age; year_(age); Wage_(age)
            Else
                Wage_(age) = Income_(age)
            End If
        End If
        
        
    Next age
    wsStart.Unprotect 'ta bort bladets skydd
       
End Sub
'------------------- Beräkning av livsinkomster, dvs löner och pensioner, skatter, bidrag  mm--------------
Sub Mcalc()
    Dim T   'för att mäta tidsåtgång
    T = Time '=HH:MM:SS
    Application.Calculation = xlCalculationManual
    
    'Typfallets egenskaper och vissa makroekonomiska storheter läses in
    'Rensa utdata och avkastning
    Call startsetup 'Se ovan
    
    If PAR < Application.Range("Rng_riktL").Value Then Exit Sub
    
    Dim Mlabels() As Variant    'För utskrift i utdata
    Dim Mlabels2() As Variant   'För utskrift till tabell 2
    Dim rnglbl As Range         'Etiketter för utskrift
    
    Dim Array_Label()
    Dim i As Integer
    Dim c As Range
    
    Dim Fasta_priser As Integer 'Beloppen uttryckta i fasta priser (1), i lönenivå (0) (för år w_ref) eller i nominella termer (-1)
    Fasta_priser = Application.Range("rng_Bara_fastapriser").Value
    
    Set rnglbl = Application.Range("rng_Top_Rubriker_utdata_1") 'se flik input!M1
    Set rnglbl = rnglbl.Offset(1).Resize(rnglbl.CurrentRegion.Rows.Count - 1, 1)
    ReDim Mlabels(1 To rnglbl.Cells.Count) 'Obs fasta antal variabler
    i = 1
    For Each c In rnglbl
''        If Fasta_priser = 1 And Trim(c.Offset(, 1)) <> "" Then 'se första kolumnen till höger
''            Mlabels(i) = c.Offset(, 1)
''        Else
            Mlabels(i) = c
''        End If
        i = i + 1
    Next c
    
    Dim civ As Integer 'Civilstånd: 0-Ensamstående 1-Sammanboende
    civ = 0
    If gift = True Then civ = 1
    
'''    'Definiera upp olika vektorer
    'Pensionsgrundande inkomst
    Dim pgi_() As Double
    
    'Egen avgift till ÅP-systemet
    Dim EgenAvg_() As Double
    
    'PensionsGrundandeBelopp
    Dim PGB_() As Double
    Dim SAPGB_() As Double
    
    'Flödet av årliga intjänade pensionsrätter/premier
    Dim IP_ratt() As Double
    Dim TP_points() As Double
    Dim GP_ratt() As Double
    Dim PP_ratt() As Double
    Dim TJP_ratt() As Double
    Dim IPS_ratt As Double 'Obs skalär
    
    'Stocken av intjänat pensionskapital
    Dim IP_pbh() As Double
    Dim GP_pbh() As Double
    Dim PP_pbh() As Double
    Dim TJP_pbh() As Double
    Dim IPS_pbh() As Double
    Dim PPS_pbh() As Double
 
    ReDim pgi_(startage To slutage) As Double
    
    'Egen avgift till ÅP-systemet
    ReDim EgenAvg_(startage To slutage) As Double
    
    'PensionsGrundandeBelopp
    ReDim PGB_(startage To slutage) As Double
    ReDim SAPGB_(startage To slutage) As Double
    
    'Flödet av årliga intjänade pensionsrätter/premier
    ReDim IP_ratt(startage To slutage) As Double
    ReDim TP_points(startage To slutage) As Double
    ReDim GP_ratt(startage To slutage) As Double
    ReDim PP_ratt(startage To slutage) As Double
    ReDim TJP_ratt(startage To slutage) As Double
    
    'Stocken av intjänat pensionskapital
    ReDim IP_pbh(startage To slutage) As Double
    ReDim GP_pbh(startage To slutage) As Double
    ReDim PP_pbh(startage To slutage) As Double
    ReDim TJP_pbh(startage To slutage) As Double
    ReDim IPS_pbh(startage To slutage) As Double
    ReDim PPS_pbh(startage To slutage) As Double
    
    'Ev. skjuta in pnsionskapital för ett givet år.
    Dim inskjut As Integer
    inskjut = Application.Range("rng_PBHYear") 'Skjuta in PBH?
    
    'Delningstal eller N/Q
    Dim dtal_ip As Double
    Dim dtal_pp As Double
    'Dim dtal_tjp As Double 'Not used
   
    'Utbetalda Pensioner
    Dim ip() As Double      'Inkomstpension
    Dim tp() As Double      'Tilläggsp (ATP)
    Dim Gbelopp() As Double 'Garantibelopp
    Dim pp() As Double      'Premiepensionen
    Dim garp() As Double    'Garantipension
    Dim ptillagg() As Double 'Pensionstillägg ds 2020:7
    
    Dim TJP() As Double     'Tjänstep
    Dim ips() As Double     'Utbetalning från eget sparande före skatt
    Dim pps() As Double     'Utbetalning från eget sparande efter skatt
    
    '**redim
    'Utbetalda Pensioner
    ReDim ip(startage To slutage) As Double      'Inkomstpension
    ReDim tp(startage To slutage) As Double      'Tilläggsp (ATP)
    ReDim Gbelopp(startage To slutage) As Double 'Garantibelopp
    ReDim pp(startage To slutage) As Double      'Premiepensionen
    ReDim garp(startage To slutage) As Double    'Garantipension
    ReDim ptillagg(startage To slutage) As Double 'InkomstPensionsTillägg
    ReDim TJP(startage To slutage) As Double     'Tjänstep
    ReDim ips(startage To slutage) As Double     'Utbetalning från eget sparande före skatt
    ReDim pps(startage To slutage) As Double     'Utbetalning från eget sparande efter skatt
    
    'Uttagsandel av pension samt ATP
    Dim uttagIP As Double
    Dim uttagPP As Double
    Dim atp_year As Integer
    Dim tp_year As Integer
    Dim atp_points As Double
    atp_year = 0: atp_points = 0: tp_year = 0
    Dim pgi_years As Single 'Antal år med PGI /Pensionstillägget
    pgi_years = 0
    
    Dim gp_und As Double    'Underlag Gararantipension
    
    'Garantibeloppet
    Dim TP94p() As Double
    ReDim TP94p(startage To 64) 'OBS Fram till 64:e året
    Dim atp94 As Double
    Dim atp94year As Long
    atp94 = 0: atp94year = 0
    
    'SAF-LO avtalet och dem gamla STP delen 'ATP poäng används vid stp mm
    Dim STP_points() As Double 'ATP poäng används vid stp mm
    ReDim STP_points(startage To slutage) As Double 'ATP poäng används vid stp mm
    
    'Använda delningstal för en annan årskull än sin egen för bl.a. alternativ pensionsålder
''    Dim alt_dtal As Double 'Delningstal för födda xxxx
''    alt_dtal = Application.Range("rng_Delningstal_födda").Value
''    'alt_dtal = Int(alt_dtal)
''    If alt_dtal < 1930 Then
''           'MsgBox ("Alt delningstal tyvärr inte möjligt före 1930") 'Vi har dock arket mortality...
''           alt_dtal = born
''    ElseIf alt_dtal < 1938 Then
''        If par <> 65 Then MsgBox ("Delningtal för " & alt_dtal & "och pensionering vid " & par & " finns inte, ersatt med 65")
''        par = 65
''        Def_ar = 65
''    End If
    
    Dim pgbyears As Integer 'Antal inkomstår fram till pgb
    pgbyears = 0
    
    pmonth = 12 - Int(12 * (born + PAR - Int(born + PAR))) 'Antal pensionsmånader Pensionen sker den 1:a den månad när individen fyller PAR år
    Tmonth = 12 - Int(12 * (born + tjp_par - Int(born + tjp_par))) 'TJP
       
    Dim ARV_tjp As Double      '0->Återbetslningsskydd, 1-utan skydd och därmed arvsvinster med en faktor av pp arvsvinster
    Dim ARV_IPS As Double    '0->Återbetslningsskydd, 1-utan skydd och därmed arvsvinster
    ARV_tjp = Application.Range("rng_Arvsvinster_TJP").Value
    ARV_IPS = ARV_tjp 'Application.Range("rng_Arvsvinster_TJP").Value
   
    Dim counter As Integer  'Heltals Räknare
    Dim diverse As Double   'Slask variabel
    
    Dim forstid As Single 'Försäkringstid vid 65 påverkar garantipensionen och SBTP/ÄFS
    forstid = Application.Range("rng_Försäkringstid_vid_65")
    
    'Koll att den >= arbetande år...
    If forstid < 40 And Int(PAR - W_start - 1) > forstid Then
        If egen_W = False Then
            forstid = mini((PAR - W_start - 1), 40)
            Application.Range("rng_Försäkringstid_vid_65") = forstid 'Fast adress
            MsgBox "Obs försäkringstid angiven under 40 år, men antal arbetande år är fler än angivet, försäkringstiden är ändrad till " & forstid & " år"
        Else
            'MsgBox "Obs försäkringstid under 40 år, påverkar garantipensionen och BT och ÄFS"
        End If
    End If

    Dim brutto() As Double
    Dim Netto() As Double
    Dim Bidrag() As Double
    Dim IndDisp() As Double
    
    ReDim brutto(startage To slutage) As Double
    ReDim Netto(startage To slutage) As Double
    ReDim Bidrag(startage To slutage) As Double
    ReDim IndDisp(startage To slutage) As Double
    
    Dim kostnadsavd As Double   'Kostnadsavdrag - Resor må handpåläggas samt övriga avdrag, i denna pensionssparande
    Dim ctxfvi As Double        'Taxerad förvärvsinkomst
    Dim cbefvi As Double        'Beskattningsbar inkomst
    Dim pensionavgift As Double
    Dim Grundavdrag As Double
    Dim kskatt As Double
    Dim kinkskatt As Double     'Kommunal inkomstskatt - Skattesatsen
    Dim kyrkskatt As Double     'Avgift till kyrkan + begravningsavgift - skattesatsen
    Dim statskatt As Double
    Dim pensredukt              'Skatterduktion pensionsavgiften
    Dim jobbavdrag              'S.k. Jobbskatteavdrag
    
    Dim kapital As Double '+ Kapitalinkomster eller Ränteutgifter Brutto
    kapital = Application.Range("rng_Kapital_pens")
    
    Dim hyra As Double
    hyra = Application.Range("Hyra").Value
    If hyra > 60000 Then
        MsgBox ("Kolla hyran, månadshyra antas" & hyra / 12 & " kr/år")
        hyra = hyra / 12
    End If
    
    Dim AntalBarn As Single     'Antal barn
    AntalBarn = 0
    
    Dim barn1 As Date
    Dim barn2 As Date
    Dim barn3 As Date
    Dim barn4 As Date
    barn1 = Application.Range("rng_Född_Barn1") 'yyyy-mm-dd
    barn2 = Application.Range("rng_Född_Barn2")
    barn3 = Application.Range("rng_Född_Barn3")
    barn4 = Application.Range("rng_Född_Barn4")
    
    Dim barnbidrag As Double
    Dim bostadsbidrag As Double
    Dim bostadstillägg As Double
    Dim SBostadstillägg As Double
    Dim bidragovr As Double
    
    'Hjälpvariabler till BTP och SBTP
    Dim formog As Double 'Förmögenhet utöver den egna boendefastigheten
    formog = Application.Range("rng_Formogenhet").Value
    Dim maxhyra As Double       'Maxhyra i BT
    maxhyra = 0
    'Dim maxhyraSBTP As Double    'Maxhyra i BT
    'maxhyraSBTP = 74000
    
    Dim MakaInk As Double
    MakaInk = Application.Range("rng_Makens_inkomst").Value
    
    Dim ap As Byte
    Dim apm As Byte
    Dim tjpm As Double
    Dim wagem As Double
    Dim garpm As Double
''    Dim SA_age As Double 'SA tagare från ålder -påverkar ev. btp
''    SA_age = wsStart.Range("AK16").Value 'Fast adress
''    If SA_age <= 16 Then sa_tagare = 999 'Inte möjligt
      'Bör styras av SA-PGB
    
    Dim riktalder As Integer
    Dim riktl As Integer
    'Nedan avseende årskull men lagstiftningen ser till inkomstår Year=årskull+bald
    riktalder = Application.Range("Rng_riktage").Value 'riktage(year_(age), 1)
    riktl = Application.Range("Rng_RiktL").Value 'riktage(year_(age), 0)
    
    If Application.Range("Rulesfromutg") = 0 Or IsMissing(Application.Range("Rulesfromutg")) Then
             'Nothing
        Else
            If (year_(Int(PAR)) > Application.Range("Rulesfromutg") And Application.Range("Rules") = 0) _
            Or Application.Range("Rules") = 1 Then
                riktalder = riktage(Application.Range("RulesfromUtg"), 1)
                riktl = riktage(Application.Range("RulesfromUtg"), 0)
            End If
    End If
    
    'Avgifter
    'Dim Dagis As Double
    'Ekvivalensskala
    'Indexserie
     
    'Rangename för utskrift i Tabell 1 ...
    Dim Col_top_tabell1 As Integer
    Col_top_tabell1 = Application.Range("rng_Top_tabell1").Column '=Kolumn C =3
    
    Dim Row_Tabell1_Slutlön As Integer
    Dim Row_Tabell1_Lönefterskatt As Integer
    Dim Row_Tabell1_Disp_inkomst As Integer
    Dim Row_Tabell1_IP As Integer
    Dim Row_Tabell1_TP As Integer
    Dim Row_Tabell1_PP As Integer
    Dim Row_Tabell1_GP As Integer
    Dim Row_Tabell1_PT As Integer
    Dim Row_Tabell1_Tot_Allmänpension As Integer
    Dim Row_Tabell1_TJP As Integer
    Dim Row_Tabell1_IPS As Integer
    Dim Row_Tabell1_Tot_Brutto As Integer
    Dim Row_Tabell1_Efterskatt As Integer
    Dim Row_Tabell1_Bidrag As Integer
    Dim Row_Tabell1_PPS As Integer
    Dim Row_Tabell1_Disp_efterskatt As Integer
    
    'och vilken rad som värden ska srivas ut på
    Row_Tabell1_Slutlön = Application.Range("rng_Tabell1_Slutlön").Row
    Row_Tabell1_Lönefterskatt = Application.Range("rng_Tabell1_Lönefterskatt").Row
    Row_Tabell1_Disp_inkomst = Application.Range("rng_Tabell1_Disp_inkomst").Row
    Row_Tabell1_IP = Application.Range("rng_Tabell1_IP").Row
    Row_Tabell1_TP = Application.Range("rng_Tabell1_TP").Row
    Row_Tabell1_PP = Application.Range("rng_Tabell1_PP").Row
    Row_Tabell1_GP = Application.Range("rng_Tabell1_GP").Row
    Row_Tabell1_PT = Application.Range("rng_Tabell1_pt").Row
    Row_Tabell1_Tot_Allmänpension = Application.Range("rng_Tabell1_Tot_Allmänpension").Row
    Row_Tabell1_TJP = Application.Range("rng_Tabell1_TJP").Row
    Row_Tabell1_IPS = Application.Range("rng_Tabell1_IPS").Row
    Row_Tabell1_Tot_Brutto = Application.Range("rng_Tabell1_Tot_Brutto").Row
    Row_Tabell1_Efterskatt = Application.Range("rng_Tabell1_Efterskatt").Row
    Row_Tabell1_Bidrag = Application.Range("rng_Tabell1_Bidrag").Row
    Row_Tabell1_PPS = Application.Range("rng_Tabell1_PPS").Row
    Row_Tabell1_Disp_efterskatt = Application.Range("rng_Tabell1_Disp_efterskatt").Row

    '----------Tidigare resultat raderas--------------
    'wsOutput.Select 'gå till utdata, ändras ev senare att göra samma sak utan att hoppa dit
    Dim kverbose As Integer 'Utskrift av vissa hjälpvariabler de startar på kolumnnr
    Dim Kverrow As Integer 'Startrad
''  'If verbose = 1 Then
''      Worksheets(book).Select
''      Cells(startage - 1, kverbose).Select
''      Range(Selection, Selection.End(xlToRight)).Select
''      Range(Selection, Selection.End(xlDown)).Select
''      Selection.ClearContents
''    'End If
    
    'wsOutput.Select
    'rensar området i flik utdata
    With wsOutput
        Set c = Application.Range("rng_Utdata_Top_left") 'Kolumn B2
        Set c = c.CurrentRegion
        'c.ClearContents
        c.Clear
        Set c = Application.Range("rng_Utdata_Top_left_verbose") 'Kolumn U2
        Set c = c.CurrentRegion
        c.Clear
        Set c = Nothing
    End With
    
    '---------Rubriker skrivs ut--------
    With wsOutput
        Set rnglbl = Application.Range("rng_Utdata_Top_left")
        Set rnglbl = rnglbl.Resize(1, UBound(Mlabels))
        rnglbl.Value = Mlabels
        Call Format_Range(rnglbl, True, 65535)
        Set rnglbl = Nothing
    End With
   
    'Och värdena Mvalues till dessa x-variater mlabels
    Dim mvalues() As Variant
    ReDim mvalues(startage To slutage, 1 To UBound(Mlabels))
    Dim m_value As Double
    
    'Några ytterligare val
    'Utgifter och skatter räknas om med inkomstindex, dvs oförändrad politik
    Dim kvoten As Double        'Regler räknas om med IBB istf PBB
    'dim iyear As Single
    Iyear = Application.Range("rng_Boundray_Year").Value '
    If IsMissing(Iyear) Or Iyear = 0 Then Iyear = Int(born) + slutage
    If Iyear < Application.Range("Modell_Year").Value And Iyear > 0 Then '
        MsgBox ("Inte möjligt att retroaktivt ändra reglerna, bortser från " & Iyear)
        Iyear = Int(born) + slutage
    End If
    If Iyear < slutage And (Iyear - Int(born)) > startage And (Iyear - Int(born)) < slutage Then
        kvoten = pbb(Iyear - Int(born)) / IBB(Iyear - Int(born))             'För inkomstindexuppräkning
    Else
        kvoten = 1 'Ingen inkomstindexuppräkning
    End If
    
    Dim mpension As Double 'underlag Inkomsttilläggspension
    
    Dim arb_avg() As Double 'För studier av arbetsgivarna
    ReDim arb_avg(startage To slutage)
    
    '---------------------------------- Här börjar beräkningarna för alla år/åldrar ---------------------------------------------------
    For age = startage To slutage
        
        kverbose = Application.Range("rng_Utdata_Top_left_verbose").Column 'Börja på kolumn
        Kverrow = Application.Range("rng_Utdata_Top_left_verbose").Row 'och rad
        If verbose = 1 Then
            'Utskrifter - Etiketter
            If age = startage And Kverrow > 1 Then wsOutput.Cells(Kverrow - 1, kverbose) = "Utdata i löpande priser"
            Array_Label = Array("År", "Ålder") 'Snabbare än att dim, redim och (x,y,z)
            If age = startage Then
                Call Write_Label(Array_Label, wsOutput.Name, Kverrow, kverbose) 'se funktionen nedan...
            End If
            'skriver ut variabelvärdet  i vektorn Array_Label
            Array_Label = Array(Int(born) + age, age) 'År och ålder
            Call Write_Label(Array_Label, wsOutput.Name, age - startage + Kverrow + 1, kverbose)
            kverbose = kverbose + UBound(Array_Label)
            
            'Debug.Print age & " " & wage_(age)
            If age = startage Then wsOutput.Cells(Kverrow, kverbose) = "Lön"
            wsOutput.Cells(age - startage + Kverrow + 1, kverbose) = Wage_(age)
            kverbose = kverbose + 1
        End If
                
        'Ekonomiskt bistånd #barn i olika åldersklasser
        Dim b1 As Integer
        Dim b2 As Integer
        Dim b3 As Integer
        Dim b4 As Integer
        Dim b5 As Integer
        Dim b6 As Integer
        Dim b7 As Integer
        Dim b8 As Integer

        Dim Utgyear As Single
        Dim Skyear As Single
        
        Utgyear = year_(age)
        Skyear = year_(age)
        If Application.Range("Rulesfromutg") = 0 Or IsMissing(Application.Range("Rulesfromutg")) Then
                  'Nothing
             Else
                 If (year_(age) > Application.Range("Rulesfromutg") And Application.Range("Rules") = 0) _
                 Or Application.Range("Rules") = 1 Then
                     Utgyear = Application.Range("Rulesfromutg")
                 End If
        End If
        If Application.Range("RulesfromSkatt") = 0 Or IsMissing(Application.Range("RulesfromSkatt")) Then
                  'Nothing
             Else
                 If (year_(age) > Application.Range("RulesfromSkatt") And Application.Range("Rules") = 0) _
                 Or Application.Range("Rules") = 1 Then
                     Skyear = Application.Range("RulesfromSkatt")
                 End If
        End If

        'PGI inkomst se modul pensionssystemet -
        pgi_(age) = pgi(Utgyear, Income_(age), pbb(age), IBB(age), FPB(age), marginal, 0, age, 0)
        
        If verbose = 1 Then
            'Debug.Print age & " " & Pgi_(age)
            If age = startage Then wsOutput.Cells(Kverrow, kverbose) = "PGI"
            'If age = startage Then wsOutput.Cells(Kverrow, kverbose + 1) = "PGI tak"
            wsOutput.Cells(age - startage + Kverrow + 1, kverbose) = pgi_(age)
            'wsOutput.Cells(age - startage + Kverrow + 1, kverbose + 1) = pgi(year_(age), Income_(age), pbb(age), IBB(age), FPB(age), marginal, 4, age, 0)
            kverbose = kverbose + 1 '2
        End If
         
         'Arbetsgivarens avgifter - notera att en del av ÅP avgiften går till staten - Se StateSim - idag ca 10,6%. OBS Rulesfrom from finns i funkt
        arb_avg(age) = arbgiv(Wage_(age), Skyear, age, marginal, 0)
        EgenAvg_(age) = pgi(Utgyear, Income_(age), pbb(age), IBB(age), FPB(age), marginal, 1, age, 0)
        
        If verbose = 1 Then
            If age = startage Then wsOutput.Cells(Kverrow, kverbose) = "Arb_givar_avgift" 'Bör inte hela avgiften tas med?
            wsOutput.Cells(age - startage + Kverrow + 1, kverbose) = 0 * EgenAvg_(age) + 1 * arb_avg(age)
            kverbose = kverbose + 1
        End If
     
        'Statens avgifter för typfallet PGB
        'Äldre hjälpfunktioner för a-kassa, sjukpenning, SA finns
        Dim tak_ As Single
        tak_ = 7.5 'Taket har sedan 1960 varit 7,5 (Pris)basbelopp
''        If Application.Range("Soc_tak") > 0 Then
''            If year_(age) >= Application.Range("Soc_tak") Then tak_ = 7.5 '10
''        End If

        If age < PAR And Int(born) > 1937 Then
            'PGB kan fås för aktivitets- och sjukersättning, barnår, värnplikt och studier
            uttagIP = 0
            uttagPP = 0
            diverse = 0
            If age > 15 And age <= riktalder Then '.. fram till 65....rikt
                'SA först från 15 års ålder
                'Kvar ev att kolla rulesfromUtg dessa saknas ...
                PGB_(age) = wsPGB.Cells(age, 5) 'PGB SA Manuellt fast adress
                If (pgi_(age) + PGB_(age)) > (tak_ * IBB(age)) Then
                    PGB_(age) = maxi(tak_ * IBB(age) - pgi_(age), 0)
                End If
                
                'BARNÅR för barn 1-4
                counter = year(barn1) - Int(born) ' Ålder vid barnets födelse
                'Ingen korrigering för om barnet är fött efter juni... i så fall året efter
                If year(barn1) > 1960 And counter > 15 Then
                    diverse = pgb_barn(year_(age), Income_(counter), pgi_(age) * KPI(counter) / KPI(counter - 1) _
                    + PGB_(age), year(barn1), age, MPGI(age), IBB(age), marginal, riktalder)
                End If
                
                counter = year(barn2) - Int(born) - 1
                If year(barn2) > 1960 And diverse = 0 And counter > 15 Then
                    diverse = pgb_barn(year_(age), Income_(counter), pgi_(age) + PGB_(age), year(barn2), age, MPGI(age), IBB(age), marginal, riktalder)
                End If
                
                counter = year(barn3) - Int(born) - 1
                If year(barn3) > 1960 And diverse = 0 And counter > 15 Then
                    diverse = pgb_barn(year_(age), Income_(counter), pgi_(age) + PGB_(age), year(barn3), age, MPGI(age), IBB(age), marginal, riktalder)
                End If
                
                counter = year(barn4) - Int(born) - 1
                If year(barn4) > 1960 And diverse = 0 And counter > 15 Then
                    diverse = pgb_barn(year_(age), Income_(counter), pgi_(age) + PGB_(age), year(barn4), age, MPGI(age), IBB(age), marginal, riktalder)
                End If
                'If diverse > 0 Then Debug.Print year_(age); diverse
                
                PGB_(age) = PGB_(age) + diverse
                
                'Värnplikt - sköts från PGB arket tak_ dim strax ovan
                PGB_(age) = PGB_(age) + wsPGB.Cells(age, 9) 'PGB VPL Manuellt fast adress
                If (pgi_(age) + PGB_(age)) > (tak_ * IBB(age)) Then
                    PGB_(age) = maxi((tak_ * IBB(age) - pgi_(age)), 0)
                End If
                
                'Studier - sköts från PGB arket SES över
                'diverse = 0
                diverse = wsPGB.Cells(age, 16)  'PGB studier Manuellt fast adress
                If marginal = 0 Then diverse = Int(diverse / 100) * 100
                PGB_(age) = PGB_(age) + diverse
                If (pgi_(age) + PGB_(age)) > (tak_ * IBB(age)) Then
                    PGB_(age) = maxi(tak_ * IBB(age) - pgi_(age), 0)
                End If
                If (PGB_(age) + pgi_(age)) > 0 And age < 71 Then pgbyears = pgbyears + 1
            
            End If
        End If 'age<par ...
   
        If verbose = 1 Then
            'Debug.Print age & " " & Pgi_(age)
            If age = startage Then wsOutput.Cells(Kverrow, kverbose) = "PGB"
            wsOutput.Cells(age - startage + Kverrow + 1, kverbose) = PGB_(age)
            kverbose = kverbose + 1
        End If
       
        'Vid startage
        Array_Label = Array("IP_rätt", "(A)TP_Poäng", "PP_rätt", "GP_rätt")
        
        '***För att hantera nollåringar
        If age = startage Then 'Antas börja jobba tidigast vid startage års ålder och taxeringen eftersläpar ett år
            IP_ratt(age) = 0
            If age < 65 Then  'PGB(SA) saknas men SA vid startage?
                If year_(age) > 1959 Then
                    TP_points(age) = maxi(pgi_(age) / FPB(age) - 1, 0)
                    If marginal = 0 Then TP_points(age) = Round(TP_points(age), 2)
                    STP_points(age) = TP_points(age)
                End If
            End If
            GP_ratt(age) = 0
            PP_ratt(age) = 0
''            If verbose = 1 Then
''                   wsOutput.Cells(Kverrow, kverbose + 0) = "IP_rätt"
''                   wsOutput.Cells(Kverrow, kverbose + 1) = "(A)TP_Poäng"
''                   wsOutput.Cells(Kverrow, kverbose + 2) = "PP_rätt"
''                   wsOutput.Cells(Kverrow, kverbose + 3) = "GP_rätt"
''                kverbose = kverbose + 4
''            End If
            If verbose = 1 Then
                'Skriver ut rubriker i vektorn Array_Label ÄR DENNA UPPDATERAD ...
                Call Write_Label(Array_Label, wsOutput.Name, Kverrow, kverbose)
                ''' Läggas till för att hentera -1 åringar
                If startage = -1 Then
                    wsOutput.Cells(Kverrow + 1, kverbose) = 0
                    wsOutput.Cells(Kverrow + 1, kverbose + 1) = 0
                    wsOutput.Cells(Kverrow + 1, kverbose + 2) = 0
                    wsOutput.Cells(Kverrow + 1, kverbose + 3) = 0
                End If
                ''''
                kverbose = kverbose + UBound(Array_Label)
            End If
        Else 'Age>startage
            'AT(p) fram till 64 års ålder
            If age < 65 Then
                If year_(age) > 1959 Then
                    TP_points(age) = maxi((pgi_(age) + PGB_(age)) / FPB(age) - 1, 0)
                    If marginal = 0 Then TP_points(age) = Round(TP_points(age), 2)
                    STP_points(age) = TP_points(age)
                Else
                    TP_points(age) = 0
                    STP_points(age) = 0
                End If
            End If
            'OBS om förvärvsvillkoret inte är uppfyllt vid PGB beloppet så ska hela PGB till IP (18.5%)
            If pgbyears >= 5 Then 'Vi ser PGI som blir känt först året efter via taxeringen
                IP_ratt(age) = ipavgift(year_(age) - 1, pgi_(age - 1) + PGB_(age - 1), age - 1, andelnya, marginal, Int(born))
                PP_ratt(age) = ppavgift(year_(age) - 1, pgi_(age - 1) + PGB_(age - 1), age - 1, andelnya, marginal, Int(born))
                GP_ratt(age) = gpavgift(year_(age) - 1, pgi_(age - 1) + PGB_(age - 1), age - 1, andelnya, marginal, Int(born), riktalder)
            ElseIf pgbyears > 0 Then
                IP_ratt(age) = ipavgift(year_(age) - 1, pgi_(age - 1) + Int(PGB_(age - 1) * 185 / 160), age - 1, andelnya, marginal, Int(born))
                PP_ratt(age) = ppavgift(year_(age) - 1, pgi_(age - 1), age - 1, andelnya, marginal, Int(born))
                GP_ratt(age) = gpavgift(year_(age) - 1, pgi_(age - 1) + PGB_(age - 1) * 185 / 160, age - 1, andelnya, marginal, Int(born), riktalder)
            Else
                IP_ratt(age) = ipavgift(year_(age) - 1, pgi_(age - 1), age - 1, andelnya, marginal, Int(born))
                PP_ratt(age) = ppavgift(year_(age) - 1, pgi_(age - 1), age - 1, andelnya, marginal, Int(born))
                GP_ratt(age) = gpavgift(year_(age) - 1, pgi_(age - 1), age - 1, andelnya, marginal, Int(born), riktalder)
            End If
            
            If age <= riktalder Then 'IPT - antal år med antingen pgi eller atp poäng obs atp-poängen samma år som inkomst
                If IP_ratt(age) > 0 Or TP_points(age - 1) > 0 Then pgi_years = pgi_years + 1
            End If
        
            If verbose = 1 Then
                wsOutput.Cells(age - startage + Kverrow + 1, kverbose + 0) = IP_ratt(age)
                wsOutput.Cells(age - startage + Kverrow + 1, kverbose + 1) = TP_points(age)
                
                Array_Label = Array(PP_ratt(age), GP_ratt(age))
                wsOutput.Cells(age - startage + Kverrow + 1, kverbose + 2) = PP_ratt(age)
                wsOutput.Cells(age - startage + Kverrow + 1, kverbose + 3) = GP_ratt(age)
                kverbose = kverbose + UBound(Array_Label) + 2
            End If
        End If
        
        'Avtalsområdena har olika åldersgränser för intjänandet se modul tjänstepensioner
        Select Case (avtal)
                 Case (1): 'Saknar TJP
                    TJP_ratt(age) = 0
                Case (2): ' ITP1
                    TJP_ratt(age) = tlITP1(age, Wage_(age), IBB(age), tjp_par, marginal)
                Case (3): 'ITP2 (k)
                    'Function tlITP2A(ByVal alder, ByVal inkomst, ByVal par, ByVal month_, Optional ByVal marginal = 0) As Double
                    If year_(age) > 1976 Then
                        TJP_ratt(age) = tlITP2A(age, Wage_(age), tjp_par, marginal)
                    Else
                        TJP_ratt(age) = 0
                    End If
                Case (4): 'SAF-LO
                    TJP_ratt(age) = SAF_LO(age, Wage_(age), IBB(age), tjp_par, year_(age), marginal)
       
                Case (5): 'KAP-KL - bortser från ev fördröjning av premieinbetalningen
                    TJP_ratt(age) = tlkap_kl(age, Wage_(age), IBB(age), tjp_par, year_(age), marginal)
                    
                Case (6): 'AKAP-KR
                    TJP_ratt(age) = tlakap_kr(age, Wage_(age), IBB(age), tjp_par, year_(age), marginal)
                
                Case (7): 'PA16 avd 2 (PA16 avd2, tidigare PA03)
                            TJP_ratt(age) = Kapan(year_(age), Int(born), IBB(age), Wage_(age), tjp_par) + _
                                            PA_indiv(year_(age), Int(born), IBB(age), Wage_(age), tjp_par)
                                            '(PA_individ med flexpensionsdelen)
                Case (8): 'PA16 avd 1
                            TJP_ratt(age) = tlPA16(age, Wage_(age), IBB(age), tjp_par, marginal)
        End Select
        
        If verbose = 1 And avtal > 1 Then
             If age = startage Then wsOutput.Cells(Kverrow, kverbose) = "TJP-premie"
             wsOutput.Cells(age - startage + Kverrow + 1, kverbose) = TJP_ratt(age)
             kverbose = kverbose + 1
        End If
        
        'Eget sparande IPS eller ISK
        IPS_ratt = 0    'Premie
        ips(age) = 0    'Utbetalningen
        pps(age) = 0    'Utbetalningen ISK / KF
        
        If Application.Range("IPS_Start") <= (Int(born) + age) And age < tjp_par Then 'Ingen avsättning efter pension
        
            If Application.Range("IPS_monthly") > 1 Then
                IPS_ratt = Application.Range("IPS_monthly") * 12 'Ingen Koll av antal månader första året
            Else
                IPS_ratt = Application.Range("IPS_monthly") * Income_(age)
                'Debug.Print age & " " & year_(age) & " " & IPS_ratt & " " & Income_(age); Application.Range("IPS_monthly")
            End If
            If marginal Then IPS_ratt = Int(IPS_ratt + 0.5)
            
        End If
        
        If verbose = 1 And Application.Range("IPS_Start") > 0 Then
             If age = startage Then wsOutput.Cells(Kverrow, kverbose) = "IPS_premie"
             wsOutput.Cells(age - startage + Kverrow + 1, kverbose) = IPS_ratt
             kverbose = kverbose + 1
        End If
        
        'Arvsvinster, värdeförändring, avgifter som delredovisas om utskrifter önskas
        If verbose = 1 Then 'Avgifter- förvaltningskostnader och avkastning

            'skriver ut texten i vektorn Array_Label, Utskrifter - Etiketter
            Array_Label = Array("Arv_IP", "Index_ip", "Förv_avgift_ip", "Arv_PP", "Avk_PP", "Förv_avgift_PP")
            If startage = age Then
                Call Write_Label(Array_Label, wsOutput.Name, Kverrow, kverbose)
            End If
            
            'Intjänande till InkomstPension
            If age > startage And age < slutage Then
                If year_(age) > 1959 Then
                
                    If year_(age) = 1960 Then Pindex(age - 1) = Pindex(age)
                    wsOutput.Cells(age - startage + Kverrow + 1, kverbose + 0) = _
                        IP_(year_(age), PAR, born, IP_ratt(age), IP_arv1(age), IP_arv2(age), IP_avg(age), IP_pbh(age - 1), uttagIP, _
                                    ip(age - 1) + 0, dtal_ip, Pindex(age) / Pindex(age - 1), def_ar, marginal, 1)
                    'If year_(age) = 2017 Then Debug.Print Utgyear; year_(age); IP_arv1(age); IP_arv2(age); IP_avg(age); IP_pbh(age - 1); Pindex(age) / Pindex(age - 1)
                    
                    wsOutput.Cells(age - startage + Kverrow + 1, kverbose + 1) = _
                        IP_(year_(age), PAR, born, IP_ratt(age), IP_arv1(age), IP_arv2(age), IP_avg(age), IP_pbh(age - 1), uttagIP, _
                                    ip(age - 1) + 0, dtal_ip, Pindex(age) / Pindex(age - 1), def_ar, marginal, 2)
                                    
                    wsOutput.Cells(age - startage + Kverrow + 1, kverbose + 2) = _
                        IP_(year_(age), PAR, born, IP_ratt(age), IP_arv1(age), IP_arv2(age), IP_avg(age), IP_pbh(age - 1), uttagIP, _
                                    ip(age - 1) + 0, dtal_ip, Pindex(age) / Pindex(age - 1), def_ar, marginal, 3)
                End If
            Else
                For i = 0 To 5
                    With wsOutput
                        .Cells(age - startage + Kverrow + 1, kverbose + i) = 0
                    End With
                Next i
            End If
            
            If age > startage Then
            'Arvsvinster, värdeökning och förvaltningskostnad PremiePensionen
                If year_(age - 1) < 2010 Then 'Kollas
                    wsOutput.Cells(age - startage + Kverrow + 1, kverbose + 3) = PP_pbh(age - 1) * (PP_arv(age) - 1)
                Else
                    wsOutput.Cells(age - startage + Kverrow + 1, kverbose + 3) = _
                        (PP_pbh(age - 1) * (PP_arv(age) - 1) / (yield(age) ^ (6 / 12)))
                End If
                
                wsOutput.Cells(age - startage + Kverrow + 1, kverbose + 4) = _
                    (RGK(age) - 1) * PP_ratt(age) + PP_pbh(age - 1) * (yield(age) - 1)
                'Antar en avkastningen efter fondavgifter annars byt ut 0 mot t.ex. 0,003
                
                wsOutput.Cells(age - startage + Kverrow + 1, kverbose + 5) = _
                    PPMavg(Utgyear, PP_pbh(age - 1), PP_avg(age), yield(age), 0)
            End If
            
            kverbose = kverbose + UBound(Array_Label)
            
        End If       'verbose slutar här
        
        'Uttagsandel, delningstal mm
        If age < Int(PAR) Then
            dtal_ip = 0
            dtal_pp = 0
            uttagIP = 0
            uttagPP = 0
        ElseIf age >= Int(PAR) And PAR <= def_ar And age < def_ar Then
            uttagIP = Application.Range("UttagIP").Value
            uttagPP = Application.Range("UttagPP").Value
            
            If PAR = def_ar Then
                uttagIP = 1 ' Application.Range("UttagIP").Value
                uttagPP = 1 ' Application.Range("UttagPP").Value
            End If
            'wsNyckelTal.Cells(5 + born - 1938, 3 + PAR - 61)
'Ny funktion för deltalIP/PP 2024-03-20
'            If Application.Range("mortality").Value = 0 And Application.Range("Bornyear") <= 1958 Then 'OBS senaste värdena ska kanske läsas från arket
'''                dtal_ip = deltal(par, Int(alt_dtal), age, Def_ar, 4)
'''                dtal_pp = deltal(par, Int(alt_dtal), age, Def_ar, 19)
'                'fnDeltal_IP2(
'                dtal_ip = deltal(PAR, Int(born), age, def_ar, 4)
'                dtal_pp = deltal(PAR, Int(born), age, def_ar, 19)
'            Else
'                dtal_ip = wsMortality.Range("m5").Value 'obs fast adress
'                dtal_pp = wsMortality.Range("m10").Value
'                'dtal_tjp = wsMortality.Range("l" & 22 + Int(par) - 60).Value
'            End If
            dtal_ip = fnDeltal_IP2(Int(born), age, PAR, def_ar)
            dtal_pp = fnDeltal_PP2(Int(born), age, PAR, def_ar)
        Else
            If age >= def_ar Then
                uttagIP = 1
                uttagPP = 1
            End If
          
'Ny funktion för deltalIP/PP 2024-03-20
'            If Application.Range("mortality").Value = 0 And Application.Range("Bornyear") <= 1958 Then 'OBS senaste värdena ska kanske läsas från arket
'''                dtal_ip = deltal(par, Int(alt_dtal), age, Def_ar, 4) ' 4 = IP, <> 4 = PP
'''                dtal_pp = deltal(par, Int(alt_dtal), age, Def_ar, 19)
'                dtal_ip = deltal(PAR, Int(born), age, def_ar, 4) ' 4 = IP, <> 4 = PP
'                dtal_pp = deltal(PAR, Int(born), age, def_ar, 19)
'            Else
'                dtal_ip = wsMortality.Range("n5").Value 'obs fast adress
'                dtal_pp = wsMortality.Range("n10").Value
'                'dtal_tjp = wsMortality.Range("l" & 22 + Int(par) - 60).Value
'            End If
            dtal_ip = fnDeltal_IP2(Int(born), age, PAR, def_ar)
            dtal_pp = fnDeltal_PP2(Int(born), age, PAR, def_ar)
        End If
        
        'Allmänna Pensioner
        If age < Int(PAR) Then
            ip(age) = 0
            tp(age) = 0
            pp(age) = 0
            garp(age) = 0
            ptillagg(age) = 0
            mpension = 0
        ElseIf age = Int(PAR) Then
            'month = 12 - Int(12 * (born + par - Int(born + par))) 'obs samma def.
            'Inkomstpension
            ip(age) = IP_(year_(age), PAR, born, IP_ratt(age), IP_arv1(age), IP_arv2(age), IP_avg(age), IP_pbh(age - 1), uttagIP, _
                                ip(age - 1) + 0, dtal_ip, Pindex(age) / Pindex(age - 1), def_ar, marginal, 0, Pindex(age) / Iindex(age))
            
            ErrHelpLogger 1311, ip(age), age, year_(age), PAR, born, IP_ratt(age), IP_arv1(age), IP_arv2(age), IP_avg(age), IP_pbh(age - 1), uttagIP, ip(age - 1) + 0, dtal_ip, Pindex(age) / Pindex(age - 1), def_ar, marginal, 0, Pindex(age) / Iindex(age)
            'Antal år med TP poäng (= <- poäng (PGI+PGB_SA) / Förhöjdt basbelopp - 1) och medelpoängen
            'Garantitillägget (63 kap §18-22, Poängen fram till och med 1994) och ATP
            If andelnya < 1 Then
                For counter = startage To UBound(TP94p)
                    If year_(counter) <= 1994 Then
                        TP94p(counter) = TP_points(counter)
                    Else
                        TP94p(counter) = 0
                    End If
                Next
                'Sorteras i stigande ordning
                For counter = startage To UBound(TP94p)
                    If TP94p(counter) > 0 Then atp94year = atp94year + 1
                    If counter > (UBound(TP94p) - 15) Then atp94 = atp94 + TP94p(counter) / 15 ' Senaste 15 åren
                    'Debug.Print counter; atp94year; TP94p(counter); Round(atp94, 2)
                Next
                If marginal = 0 Then atp94 = Round(atp94, 2)
                Gbelopp(age) = tp_(atp94, atp94year, civ, PAR, born, age, uttagIP, def_ar, year_(age), marginal)
                Gbelopp(age) = Gbelopp(age) * pmonth / 12
               
                'ATP till 64 års dagen -OBS TP_points sorteras om
                Call QuickSort(TP_points, LBound(TP_points), UBound(TP_points)) 'Poängen sorterade i stigande ordning
                For counter = startage To UBound(TP_points)
                    If TP_points(counter) > 0 Then atp_year = atp_year + 1
                    If counter > (UBound(TP_points) - 15) Then atp_points = atp_points + TP_points(counter) / 15 ' Senaste 15 åren
                    'Debug.Print counter; atp_year; TP_points(counter); round(atp_points,2)
                Next
                If marginal = 0 Then atp_points = Round(atp_points, 2)
                tp(age) = tp_(atp_points, atp_year, civ, PAR, born, age, uttagIP, 0 * def_ar, year_(age), marginal)
                tp(age) = tp(age) * (1 - andelnya) * pmonth / 12

                'Debug.Print atp_points; tp(age)
                'Garantibeloppet - intjänandet fram till 1994 ...
                If Gbelopp(age) - (tp(age) + ip(age) * 185 / 160) > 0 Then
                    tp(age) = tp(age) + Gbelopp(age) - (tp(age) + ip(age) * 185 / 160)
                    'Debug.Print Gbelopp(Age) - (tp(Age) + ip(Age) * 185 / 160)
                End If
                If marginal = 0 Then tp(age) = Int(tp(age) / pmonth + 0.5) * pmonth
            End If

            'Garantipensionen - Först vid 65 års ålder (riktalder) och gått i pension
            Dim gpundtab1 As Double
            gpundtab1 = 0

            If age < riktalder Or year_(age) < (Int(PAR) + born) Then
                garp(age) = 0
            Else
                If Application.Range("mortality").Value = 0 And Application.Range("Bornyear") <= 1958 Then 'OBS senaste värdena ska kanske läsas från arket
''                    dtal_ip = deltal(par, Int(alt_dtal), age, Def_ar, 4)
                     dtal_ip = deltal(PAR, Int(born), age, def_ar, 4)
                Else
                    dtal_ip = wsMortality.Range("L5").Value 'obs fast adress
                End If
''                If Application.Range("Mortality") = 0 Then
''                    dtal_ip = deltal(riktalder, Int(alt_dtal), riktalder, riktalder, 4) 'Delningstalet
''                Else
''                    dtal_ip = wsMortality.Range("i" & 22 + Int(par) - 60).Value 'fast adress
''                End If

                If Int(PAR + 1) = age Then garp(age) = garp(age) * 12 / pmonth 'KOLL

                gp_und = IP_(year_(age), maxi(PAR, riktalder), born, GP_ratt(age), IP_arv1(age), IP_arv2(age), IP_avg(age), GP_pbh(age - 1), 1, _
                             garp(age - 1) + 0, dtal_ip, Pindex(age) / Pindex(age - 1), def_ar, marginal, 0, Pindex(age) / Iindex(age))
                             '+ tp (age) * maxi(1, tp_faktor(par))
                
                garp(age) = gp(gp_und * 12 / pmonth + tp(age) * maxi(1, tp_faktor(PAR)), civ, Int(born), pbb(age), forstid, marginal, age, Utgyear, Iyear, IBB(age), kvoten, riktalder, uttagIP)
                'Debug.Print "Garp? 1 "; gp_und; garp(age); riktalder; dtal_ip; Application.Range("RulesfromUtg")
                If Int(PAR) = age Then garp(age) = (garp(age) * pmonth / 12) * uttagIP

                mpension = gp_und + tp(age) * maxi(1, tp_faktor(PAR))  'IP_(year_(age), par, born, GP_ratt(age), IP_arv1(age), IP_arv2(age), IP_avg(age), GP_pbh(age - 1), uttagIP, _
                                 garp(age - 1) + 0, dtal_ip, Pindex(age) / Pindex(age - 1), Def_ar, marginal, 0, Pindex(age) / Iindex(age))

                mpension = mpension / pmonth
                If age = Int(PAR) Then gpundtab1 = mpension
                '' mpension = ip(age) * (185 / 160) * (1 / uttagIP) / pmonth 'Viss förenkling
                '' mpension = mpension + tp(age) * maxi(1, tp_faktor(par)) * (1 / uttagIP) / pmonth '
            End If

            'Premipensionen - Avkastning linjär under året
            pp(age) = ppkassa(PAR, born, age, PP_pbh(age - 1) * (yield(age) ^ ((12 - pmonth) / 12)), def_ar, uttagPP, marginal)

            'Inkomstpensionstillägget
            If (age < riktalder Or year_(age) < (Int(PAR) + born)) Then
                ptillagg(age) = 0
            Else
                'forsäkringstid med mera; gränser och belopp bör ses över
                'se över: if age>15 ?, titta på Year[now()]
                If 2021 - Int(born) > startage Then
                    ptillagg(age) = tillagg(12 * mpension, Utgyear, Iindex(age), Iindex(2021 - Int(born)), uttagIP, pgi_years)
                Else
                    ptillagg(age) = tillagg(12 * mpension, Utgyear, Iindex(age), 186.52, uttagIP, pgi_years)
                End If

                If year_(age) = 2021 Then
                    ptillagg(age) = ptillagg(age) * mini(4, pmonth) / 12 'införs 1/9
                Else
                    ptillagg(age) = ptillagg(age) * pmonth / 12
                End If
            End If
             
        Else ' Age>par -------------------------------------------------------------------
            Dim ppmonth As Integer
            ppmonth = pmonth
            If (Int(PAR) + 1) < age Then ppmonth = 12
            'Inkomst- o TP pension
'            If age = 70 Then
'                Stop
'                dtal_ip = 16.57 'deltal(par, born, age, Def_ar, 4)
'            ElseIf age = 66 Then
'                dtal_ip = 16.59
'            End If
            If age <= def_ar Then dtal_ip = fnDeltal_IP(born, age)
            
            ip(age) = IP_(year_(age), PAR, born, IP_ratt(age), IP_arv1(age), IP_arv2(age), IP_avg(age), IP_pbh(age - 1), uttagIP, _
                 ip(age - 1) + 0, dtal_ip, Pindex(age) / Pindex(age - 1), def_ar, marginal, 0, Pindex(age) / Iindex(age))
            tp(age) = 0
            ErrHelpLogger 1423, ip(age), age, year_(age), PAR, born, IP_ratt(age), IP_arv1(age), IP_arv2(age), IP_avg(age), IP_pbh(age - 1), uttagIP, _
                 ip(age - 1) + 0, dtal_ip, Pindex(age) / Pindex(age - 1), def_ar, marginal, 0, Pindex(age) / Iindex(age)
                 
            If andelnya < 1 Then
                If year_(age) < 2000 Then
                  tp(age) = tp(age - 1) * Pindex(age) / Pindex(age - 1) * 12 / ppmonth
                  If Gbelopp(age - 1) > 0 Then Gbelopp(age - 1) = Gbelopp(age - 1) * Pindex(age) / Pindex(age - 1) * 12 / ppmonth
                ElseIf year_(age) = 2000 Then
                  tp(age) = tp(age - 1) * (Pindex(age) / Pindex(age - 1)) * (12 / ppmonth) / 0.996
                  If Gbelopp(age - 1) > 0 Then Gbelopp(age) = Gbelopp(age - 1) * (Pindex(age) / Pindex(age - 1)) * (12 / ppmonth) / 0.996
                Else
                  tp(age) = tp(age - 1) * (Pindex(age) / Pindex(age - 1)) * (12 / ppmonth) / 1.016
                   If Gbelopp(age - 1) > 0 Then Gbelopp(age) = Gbelopp(age - 1) * (Pindex(age) / Pindex(age - 1)) * (12 / ppmonth) / 1.016
                End If
                
                If uttagIP < 1 Then '--> Def_ar>par och TP>0
                    If def_ar = age Then
                        'Resterande andel beräknas samt #pmonth vid andra uttaget
                        ppmonth = 12 - Int(12 * (born + def_ar - Int(born + def_ar))) '
                        diverse = tp_(atp_points, atp_year, civ, def_ar, born, age, 1 - uttagIP, 0 * def_ar, year_(age), marginal)
                        diverse = diverse * (1 - andelnya) * ppmonth / 12
                        If marginal = 0 Then
                            tp(age) = Int(tp(age) / 12 + 0.5) * 12
                            diverse = Int(diverse / ppmonth + 0.5) * pmonth
                        End If
                            
                        tp(age) = tp(age) + diverse
                        
                    ElseIf def_ar > PAR And def_ar = age + 1 Then
                        
                        tp(age) = tp(age - 1) - diverse 'Uppdelat pg ppmonth
                        If year_(age) < 2000 Then
                            diverse = diverse * Pindex(age) / Pindex(age - 1) * 12 / ppmonth
                            tp(age) = tp(age - 1) * Pindex(age) / Pindex(age - 1)
                        ElseIf year_(age) = 2000 Then
                          diverse = diverse * (Pindex(age) / Pindex(age - 1)) * (12 / ppmonth) / 0.996
                          tp(age) = tp(age - 1) * (Pindex(age) / Pindex(age - 1)) / 0.996
                        Else
                          diverse = diverse * (Pindex(age) / Pindex(age - 1)) * (12 / ppmonth) / 1.016
                          tp(age) = tp(age - 1) * (Pindex(age) / Pindex(age - 1)) / 1.016
                        End If
                         tp(age) = tp(age) + diverse
                        If marginal = 0 Then
                            tp(age) = Int(tp(age) / 12 + 0.5) * 12
                        End If
                    Else
                       ppmonth = 12
                       If year_(age) < 2000 Then
                            tp(age) = tp(age - 1) * Pindex(age) / Pindex(age - 1)
                        ElseIf year_(age) = 2000 Then
                          tp(age) = tp(age - 1) * (Pindex(age) / Pindex(age - 1)) / 0.996
                        Else
                          tp(age) = tp(age - 1) * (Pindex(age) / Pindex(age - 1)) / 1.016
                        End If
                       
                        If marginal = 0 Then
                            tp(age) = Int(tp(age) / 12 + 0.5) * 12
                        End If
                    End If
                    'Debug.Print age; tp(age); diverse
                End If 'uttag_ip
            End If 'andelnya
            
            'Med garantibeloppet - intjänandet fram till 1994 ...
            Gbelopp(age) = tp_(atp94, atp94year, civ, PAR, born, age, uttagIP, def_ar, year_(age), marginal)
            If Gbelopp(age) - (tp(age) + ip(age) * 185 / 160) > 0 Then
             tp(age) = tp(age) + Gbelopp(age) - (tp(age) + ip(age) * 185 / 160)
             'Debug.Print Gbelopp(Age) - (tp(Age) + ip(Age) * 185 / 160)
             End If
            
             'Garantipension -obs rng_Delningstal_födda ...
             If Application.Range("mortality").Value = 0 And Application.Range("Bornyear") <= 1958 Then 'OBS senaste värdena ska kanske läsas från arket
                'dtal_ip = deltal(par, Int(alt_dtal), age, Def_ar, 4)
                dtal_ip = deltal(PAR, Int(born), age, def_ar, 4)
            Else
                dtal_ip = wsMortality.Range("L5").Value 'obs fast adress
            End If
''            If Application.Range("Mortality") = 0 Then
''                dtal_ip = deltal(riktalder, Int(alt_dtal), riktalder, riktalder, 4)
''            Else
''                dtal_ip = wsMortality.Range("j" & 22 + Int(par) - 60).Value
''            End If
            
            gp_und = IP_(year_(age), maxi(PAR, riktalder), born, GP_ratt(age), IP_arv1(age), IP_arv2(age), IP_avg(age), GP_pbh(age - 1), 1, _
                    gp_und + 0, dtal_ip, Pindex(age) / Pindex(age - 1), def_ar, marginal, 0, Pindex(age) / Iindex(age))
            'month = 12
            
            If age >= riktalder Then
                garp(age) = gp(gp_und + tp(age) * maxi(1, tp_faktor(PAR)), civ, Int(born), pbb(age), forstid, marginal, age, Utgyear, Iyear, IBB(age), kvoten, riktalder, uttagIP)
            Else
                garp(age) = 0
            End If
                           
            mpension = gp_und + tp(age) * maxi(1, tp_faktor(PAR))  '+IP_(year_(age), par, born, GP_ratt(age), IP_arv1(age), IP_arv2(age), IP_avg(age), GP_pbh(age - 1), uttagIP, _
                             garp(age - 1) + 0, dtal_ip, Pindex(age) / Pindex(age - 1), Def_ar, marginal, 0, Pindex(age) / Iindex(age))
            mpension = mpension / pmonth
            If mpension = 0 And ip(age) > 0 Then _
            mpension = ip(age) * (185 / 160) * (1 / uttagIP) / ppmonth + tp(age) * maxi(1, tp_faktor(PAR)) * (1 / uttagIP) / pmonth  'Förenkling
             
            'If age = Int(par) + 1 Then Debug.Print mpension
            'Premiepension:
            If Application.Range("rng_Förenklad_beräkning") = 1 Then 'Förenklad beräkning
                'If Def_ar < par Then 'se till prognosstandard eller förskottsräntan i delningtalet
                'If def_ar < age Then 'Så är uttagsandelen 100% efter ett tag
                If age > def_ar Or (age > PAR And age < def_ar) Or uttagPP = 1 Then 'Så är uttagsandelen 100% efter ett tag (ändrar rent kosmetiskt)
                    'Alltså här hamnar vi om PAR har passerats, men def_ar inte har uppnåtts, ELLER om def_ar passerats. ELLER om uttaget är 100% 2024-02-28 /MB
'''Test borttagning av villkor. Inte alls säker /MB
'''                    If age > (def_ar + 1) And PP_ratt(age - 1) = 0 Then
                       'Tommy - Driftaavdrag och Förskottsräntan för fondförsäkran, traditionell inom perentes()
                        If year_(age) > 2017 Then
                            pp(age) = pp(age - 1) * yield(age) / 1.0165 '2017-12-01 0,100 1,750 (0,1 1,75)
                        ElseIf year_(age) > 2014 Then
                            pp(age) = pp(age - 1) * yield(age) / 1.029 '2014-03-01 0,100 3,000 (0,1 3,00)
                        ElseIf year_(age) > 2007 Then
                            pp(age) = pp(age - 1) * yield(age) / 1.039 '2007-04-01 0,100 4,000 (0,1 2,30)
                        ElseIf year_(age) > 2002 Then
                            pp(age) = pp(age - 1) * yield(age) / 1.027 '2002-12-01 0,300 3,000 (0,3 3,00)
                        Else
                            pp(age) = pp(age - 1) * yield(age) / 1.036 '2001-01-01 0,300 4,000 (0,3 4,00)
                        End If
'''                    Else
'''                        pp(age) = ppkassa(PAR, born, age, PP_pbh(age - 1), def_ar, uttagPP, marginal)
'''                    End If
                Else 'alternativt med prognosstandard:
                    If year_(age) > (born + PAR + 1) And (uttagPP = 1) And PP_ratt(age - 1) = 0 Then
                        pp(age) = pp(age - 1) * yield(age) / 1.035
                    Else
                        pp(age) = ppkassa(PAR, born, age, PP_pbh(age - 1), def_ar, uttagPP, marginal)
                    End If
                End If
            Else 'Faktiskt
                pp(age) = ppkassa(PAR, born, age, PP_pbh(age - 1), def_ar, uttagPP, marginal)
            End If
            If (age < riktalder Or year_(age) < (Int(PAR) + born)) Then
                ptillagg(age) = 0
            Else
                If 2021 - Int(born) > startage Then
                    ptillagg(age) = tillagg(12 * mpension, Utgyear, Iindex(age), Iindex(2021 - Int(born)), uttagIP, pgi_years)
                Else
                    ptillagg(age) = tillagg(12 * mpension, Utgyear, Iindex(age), 182.58, uttagIP, pgi_years)
                End If
                If year_(age) = 2021 Then
                    ptillagg(age) = ptillagg(age) * mini(4, pmonth) / 12 'införs 1/9
                Else
                    ptillagg(age) = ptillagg(age) * pmonth / 12
                End If
            End If
        End If  'Allmänna Pensioner
     
      'Utbetalning av tjänstepensionen och eget sparande
       Dim underlag As Double 'Underlag för förmånsbaserat pension
       underlag = 0
       
       If age < Int(tjp_par) Then
            TJP(age) = 0
            ips(age) = 0 'Uttaget antas vara samma som tjps uttag
            pps(age) = 0
       ElseIf year_(age) = Int(born + tjp_par) Then
            'Tjänstepension
            TJP(age) = tjpkassa(tjp_par, born, age, (TJP_ratt(age) + TJP_pbh(age - 1)) * (yield(age) ^ ((12 - Tmonth) / 12)), marginal, 0)
            
            Dim uttag As Double
            If Application.Range("rng_Temp_IPS_Uttag") = 0 Then
            uttag = Application.Range("rng_Exp_life")
            Else
            uttag = Application.Range("rng_Temp_IPS_Uttag")
            End If
            'Individuellt pensionsSparande (IPS) OBS hanterades tidigare med funk tjpkassa(). Förenklad version med prognosstandard
            ips(age) = (IPS_pbh(age - 1) * (1 + Application.Range("rng_FondAvkastning")) ^ (uttag / 2)) / uttag
            'Investeringssparkonto och kapitalförsäkring
            pps(age) = (PPS_pbh(age - 1) * (1 + Application.Range("rng_FondAvkastning")) ^ (uttag / 2)) / uttag
            
       'Förmånsbestämda tjänstepensioner, se modul TjänstepensionerFörmån
       TJP(age) = FTJP(TJP(age), startage, age, tjp_par, born, avtal, STP_points(), tp(), marginal)
          
       Else
            'Korrigering för # månader
            If year_(age) = Int(tjp_par + born + 1) Then
                TJP(age) = (TJP(age - 1) * pbb(age) / pbb(age - 1)) * (12 / Tmonth)
                ips(age) = (ips(age - 1) * pbb(age) / pbb(age - 1)) * (12 / Tmonth)
                pps(age) = (pps(age - 1) * pbb(age) / pbb(age - 1)) * (12 / Tmonth)
            Else
                TJP(age) = TJP(age - 1) * pbb(age) / pbb(age - 1)
                ips(age) = ips(age - 1) * pbb(age) / pbb(age - 1)
                pps(age) = pps(age - 1) * pbb(age) / pbb(age - 1)
            End If
            'Korrigering för temporärt uttag
            If Application.Range("rng_Temp_Tjp_Uttag") > 0 Then
                If age = tjp_par + Application.Range("rng_Temp_Tjp_Uttag") Then
                    TJP(age) = TJP(age) * (12 - Tmonth) / 12
                ElseIf age > tjp_par + Application.Range("rng_Temp_Tjp_Uttag") Then
                    TJP(age) = 0
                End If
            End If
            If Application.Range("rng_Temp_IPS_Uttag") > 0 Then
                If age = tjp_par + Application.Range("rng_Temp_IPS_Uttag") Then
                    ips(age) = ips(age) * (12 - Tmonth) / 12
                    pps(age) = pps(age) * (12 - Tmonth) / 12
                ElseIf age > tjp_par + Application.Range("rng_Temp_IPS_Uttag") Then
                    ips(age) = 0
                    pps(age) = 0
                End If
            End If
            'Om utbetalningen understiger behållningen, sätt utbetalning till behållning
            If PPS_pbh(age - 1) < pps(age) Then
            pps(age) = PPS_pbh(age - 1) * yield(age)
            End If
        End If

        'Ev. avrundas
        If marginal = 0 Then
            ip(age) = Int(ip(age) / 12 + 0.5) * 12
            pp(age) = Int(pp(age) / 12 + 0.5) * 12
            TJP(age) = Int(TJP(age) / 12 + 0.5) * 12
            ips(age) = Int(ips(age) / 12 + 0.5) * 12
            pps(age) = Int(pps(age) / 12 + 0.5) * 12
            'ptillagg är avrundat
        End If
        
        'Pensionskapital - Utgående balans
        If age = startage Then
            IP_pbh(age) = 0
            GP_pbh(age) = 0
            PP_pbh(age) = 0
            TJP_pbh(age) = 0
            IPS_pbh(age) = 0
            PPS_pbh(age) = 0
        Else
            'Återkallat belopp = 0 och inte med, kräver handpåläggning
            If year_(age) > 1960 Then 'Data finns först från och med 1960
                If age < slutage Then
                    IP_pbh(age) = IP_(year_(age), PAR, born, IP_ratt(age), IP_arv1(age), IP_arv2(age), IP_avg(age), IP_pbh(age - 1), uttagIP, _
                                    ip(age - 1) + 0, dtal_ip, (Pindex(age) / Pindex(age - 1)), def_ar, marginal, 4, Pindex(age) / Iindex(age))
                    ErrHelpLogger 1833, IP_pbh(age), age, year_(age), PAR, born, IP_ratt(age), IP_arv1(age), IP_arv2(age), IP_avg(age), IP_pbh(age - 1), uttagIP, ip(age - 1) + 0, dtal_ip, (Pindex(age) / Pindex(age - 1)), def_ar, marginal, 4, Pindex(age) / Iindex(age)
                    
                    GP_pbh(age) = IP_(year_(age), maxi(PAR, riktalder), born, GP_ratt(age), IP_arv1(age), IP_arv2(age), IP_avg(age), GP_pbh(age - 1), 1, _
                                    ip(age - 1) + 0, dtal_ip, Pindex(age) / Pindex(age - 1), def_ar, marginal, 4, Pindex(age) / Iindex(age))
                   'If age = 65 Or age = 66 Then Debug.Print GP_pbh(age)
                End If
                'Premiepensionen, avkastningen insättningen sker vid sluter av året
                 If year_(age - 1) < 2010 Then
                    PP_pbh(age) = PP_pbh(age - 1) * yield(age) + PP_pbh(age - 1) * (PP_arv(age) - 1) + _
                            RGK(age) * PP_ratt(age) + PPMavg(Utgyear, PP_pbh(age - 1), PP_avg(age), yield(age), 0, 0)
                Else
                    PP_pbh(age) = PP_pbh(age - 1) * yield(age) + PP_pbh(age - 1) * (PP_arv(age) - 1) / (yield(age) ^ (6 / 12)) + _
                            RGK(age) * PP_ratt(age) + PPMavg(Utgyear, PP_pbh(age - 1), PP_avg(age), yield(age), 0, 0)
                End If
            End If
        End If
        
       'Pensionskapital justerat för uttag samt tjänstepension och IPS
        If age = startage Then
                TJP_pbh(age) = TJP_ratt(age) * yield(age) ^ 0.5
                IPS_pbh(age) = IPS_ratt * yield(age) ^ 0.5
                PPS_pbh(age) = IPS_ratt * yield(age) ^ 0.5
        Else
            'Ev uttag av premiepensionen så ...
            PP_pbh(age) = PP_pbh(age) - pp(age) * yield(age) ^ 0.5 'Ränteeffekt, dvs avkastning på det utbetalda beloppet
            
            'Hänsyn till återbetalningsskydd inom TJP -> ARV_tjp =0 OBS enbart avgiftsbaserade tjp
            TJP_pbh(age) = TJP_ratt(age) * yield(age) ^ 0.5 + _
                    TJP_pbh(age - 1) * yield(age) * TP_avg(age) + TJP_pbh(age - 1) * (PP_arv(age - 1) - 1) * ARV_tjp
                    
            kskatt = TJP_pbh(age) * avkskatt(year_(age), 0)
            TJP_pbh(age) = TJP_pbh(age) - kskatt - TJP(age) * yield(age) ^ 0.5 'Omräknas inte som Premiepensionen
            'OBS OM TJP är förmånsbaserad kommer denna avräkning att ske för snabbt...

            Dim Spar As Integer
            Spar = Application.Range("rng_kapitalförsäkring")

            If Spar = 1 Then 'KF Kapitalunderlag
                PPS_pbh(age) = PrivatSpar(PPS_pbh(age - 1), IPS_ratt, yield(age), year_(age), Spar) - pps(age) * yield(age) ^ 0.5
            ElseIf Spar = 2 Then 'ISK Kapitalunderlag
                PPS_pbh(age) = PrivatSpar(PPS_pbh(age - 1), IPS_ratt, yield(age), year_(age), Spar) - pps(age) * yield(age) ^ 0.5
            Else  ' IPS Kapitalunderlag
              kskatt = 0
              IPS_pbh(age) = (IPS_ratt * yield(age) ^ 0.5 + _
                      IPS_pbh(age - 1) * yield(age) * TP_avg(age))   '* (PP_arv(age - 1) - 1) * ARV_IPS)  ingen arvsvinst
              kskatt = IPS_pbh(age) * avkskatt(year_(age), 0)
              IPS_pbh(age) = IPS_pbh(age) - kskatt - ips(age) * yield(age) ^ 0.5
            End If
            If marginal = 0 Then kskatt = Int(kskatt * 100) / 100
        End If
         
         'koll på negativa behållningar
         If TJP_pbh(age) < 0 Then TJP_pbh(age) = 0
         If IPS_pbh(age) < 0 Then IPS_pbh(age) = 0
         If PPS_pbh(age) < 0 Then PPS_pbh(age) = 0
        
        'Om inskjutning av behållningar så antas att PBH avser utgående balans
        If inskjut > 0 Then
            If year_(age) = inskjut Then  '
                GP_pbh(age) = Application.Range("rng_PBH_IP") / IP_pbh(age) ' = IP_pbh(age) * 185 / 160 'Annars Approximativt underlag för nya
                IP_pbh(age) = Application.Range("rng_PBH_IP")
                PP_pbh(age) = Application.Range("rng_PBH_PP")
                TJP_pbh(age) = Application.Range("rng_PBH_tjp")
                IPS_pbh(age) = Application.Range("rng_PBH_ips")
            End If
        End If

        If marginal = 0 Then
            IP_pbh(age) = Int(IP_pbh(age))
            GP_pbh(age) = Int(GP_pbh(age))
            PP_pbh(age) = Int(PP_pbh(age))
            TJP_pbh(age) = Int(TJP_pbh(age))
            IPS_pbh(age) = Int(IPS_pbh(age))
            PPS_pbh(age) = Int(PPS_pbh(age))
        End If
       
        If verbose = 1 Then
            Array_Label = Array("IP_PBH", "PP_PBH", "GP_PBH", "TJP_PBH", "IPS_PBH", "PPS_PBH", "IP", "(A)TP", "PP", "GP", "ptillagg", "TJP", "IPS")
            If age = startage Then
                Call Write_Label(Array_Label, wsOutput.Name, Kverrow, kverbose)
            End If
            'skriver ut texten i vektorn Array_Label
            Array_Label = Array(IP_pbh(age), PP_pbh(age), GP_pbh(age), TJP_pbh(age), IPS_pbh(age), PPS_pbh(age), ip(age), tp(age), pp(age), garp(age), ptillagg(age), TJP(age), ips(age))
            Call Write_Label(Array_Label, wsOutput.Name, age - startage + Kverrow + 1, kverbose)
            kverbose = kverbose + UBound(Array_Label)
        End If
        'inkomsterna klara skatt och bidrag kvar
        'Kapitalinkomsterna från IPS antas uppkomma först vid pensioneringen
        brutto(age) = Income_(age) + ip(age) + tp(age) + garp(age) + pp(age) + TJP(age) + ptillagg(age)
        If Application.Range("rng_kapitalförsäkring") = 0 Then brutto(age) = brutto(age) + ips(age)
        
        'Avdrag från inkomsten
        kostnadsavd = 0
        If brutto(age) > 0 Then
            'Kostnadsavd = Kostnadsavd + avdragRES(0, 1, marginal, skyear)          'Resor
            If Application.Range("rng_kapitalförsäkring") = 0 Then
                If (IsMissing(Application.Range("Rulesfromskatt")) Or Application.Range("Rulesfromskatt") = 0) Then
                    If TJP_ratt(age) = 0 Or year_(age) < 2016 Then
                           kostnadsavd = kostnadsavd + avdragRES(IPS_ratt, 2, marginal, Skyear)        'IPS
                    End If
                Else
                    If Application.Range("Rulesfromskatt") < 2016 Then
                           kostnadsavd = kostnadsavd + avdragRES(IPS_ratt, 2, marginal, Skyear)        'IPS
                    End If
                End If
            End If
            'IKS - kapitalbesk schablonm
            'Kostnadsavd = Kostnadsavd + avdragRES(0, 3, marginal, year_(age))           'Övriga avdrag
        End If
        
        'Taxerad förvärvsinkomst - Ctxfvi
        ctxfvi = brutto(age) - kostnadsavd '- kapital
        If marginal = 0 Then
            ctxfvi = Int(ctxfvi / 100) * 100
        End If
              
        'Grundavdrag - ålder vid årets ingång -> age-1
        Dim Gage As Integer 'Hjälpfunktion för åldern när de äldre får ett förhöjt grundavddrag
        Gage = Xage(Skyear)
        Grundavdrag = avdragxx(ctxfvi, pbb(age), marginal, age, year_(age), 2100, IBB(age), kvoten, Skyear, Gage)
    
        'PGI avgift från skattskyldig utgiftsregel ointe skatteregel
        pensionavgift = pgi(Utgyear, Income_(age), pbb(age), IBB(age), FPB(age), marginal, 1, age, 0)
        
        'Beskattningsbar inkomst
        cbefvi = ctxfvi - Grundavdrag - pensionavgift + pgi(Utgyear, Income_(age), pbb(age), IBB(age), FPB(age), marginal, 2, age, 0)
        
         'Kommunal inkomstskatt
        kinkskatt = cbefvi * Kom_skatt(age)
        If marginal = 0 Then kinkskatt = Int(kinkskatt)
        
        'Agift till Sv kyrkan eller annat trossamfund
        kyrkskatt = cbefvi * Begravavg(age)
        If marginal = 0 Then kyrkskatt = Int(kyrkskatt)
        
        'Statlig inkomstskatt och public service avgiften inlagd i statsskatten, gränsvärdena justerade för "rules"
        statskatt = statlig(cbefvi, Tax_limit1(age), Tax_limit2(age), marginal) + PublicAvg(cbefvi, 0.01, age, marginal, Skyear) 'Skattegränser ändrade..
                
        'Kapitalskatt eller skattereduktion
        Dim kapskatt As Double 'kapskatt är kapitalskatt eller skattereduktionen
        kapskatt = 0 'Kanske ska läsas in application.range("kapital")
        'Kapital och kapitalskatt antas först vid PAR
        If age >= Int(PAR) And (kapital <> 0) Then
            If kapital < -100000 Then
                kapskatt = -30000 + ((kapital + 100000) * 0.7 * 0.3)
            Else
                kapskatt = kapital * 0.3
            End If
            If kapskatt < 0 Then
                If statskatt > (-kapskatt) Then
                    statskatt = statskatt + kapskatt
                Else
                    kinkskatt = kinkskatt + statskatt + kapskatt
                    statskatt = 0
                End If
            Else
                statskatt = statskatt + kapskatt
            End If
        End If
        
        'Förmögenhets- och fastighetsskatt samt reduktioner för dessa saknas
        'Reduktion pensionsavgiften - funktionen koll för RulesfromSkatt och enbart red. som ev. påverkas
        pensredukt = pgi(Skyear, Income_(age), pbb(age), IBB(age), FPB(age), marginal, 2, age, 0)
            
        'Skattereduktion för arbetsinkomster, det s.k. Jobbskatteavdraget
        jobbavdrag = Jobbxx(Wage_(age), age, Kom_skatt(age), pbb(age), marginal, ctxfvi - Wage_(age), Skyear, 2199, IBB(age), kvoten, Skyear, Gage)
        
        If (kinkskatt - pensredukt) < jobbavdrag Then jobbavdrag = kinkskatt - pensredukt 'koll
        
        'Skattereduktion för a-kassa och fackförningeavgift
        Dim akassa As Double
        Dim fack As Double
        Dim rakassa As Double
        Dim rfack As Double
        rakassa = 0: rfack = 0
        
        akassa = Application.Range("akasseavg")
        fack = Application.Range("fack")
        
        'Utgiften antas följa KPI -2007 höjdes dock avgifterna och dessa blev mer differentierade (hänsyn till risken att bli arbetslös)
        If age >= W_start And age <= PAR Then
            If (2019 - Int(born) > startage) Then
                akassa = akassa * KPI(age) / KPI(2019 - born)
                fack = fack * KPI(age) / KPI(2019 - born)
            End If
            If age = PAR Then
                akassa = akassa * (12 - pmonth)
                fack = fack * (12 - pmonth)
            End If
        Else
                akassa = 0
                fack = 0
        End If
        
        'Historiska regler... 1995-2006
        If Skyear > 1995 And Skyear <= 2006 Then
           rakassa = 0.75 * akassa
           If fack >= 400 Then rfack = 0.25 * fack
        End If
        
        If Skyear = 2018 Then
           rakassa = 0
           If fack >= 400 Then rfack = 0.25 * fack * 6 / 12 'Infördes 2018
        End If
        If Skyear = 2019 Then
           rakassa = 0
           If fack >= 400 Then rfack = 0.25 * fack * 3 / 12 'och slopades 2019
        End If
        
        If Skyear >= 2022 Then
            rakassa = 0.25 * akassa
            If Skyear = 2022 Then rakassa = rakassa / 2 '1 juli
        End If
        
        'Tidigare skatteReduktionen om 1320 kronor: utgick från taket på PGI (8.07 IBB)
        If Skyear > 1999 And Skyear < 2002 Then
            If Int((pgi_(age) / 0.93)) < 135000 Then
                rakassa = rakassa + mini(1320, Int(pgi_(age) / 0.93))
            Else
                rakassa = rakassa + 1320 - 0.012 * (Int(pgi_(age) / 0.93) - 135000)
            End If
            If marginal = 0 Then rakassa = Int(rakassa)
            If Skyear = 2004 And cbefvi >= 100 Then rakassa = rakassa + 200
        End If
        
        If Skyear > 1999 And Skyear < 2002 Then
                If Int((pgi_(age) / 0.93)) < 135000 Then
                    rakassa = rakassa + mini(1320, Int(pgi_(age) / 0.93))
                Else
                    rakassa = rakassa + 1320 - 0.012 * (Int(pgi_(age) / 0.93) - 135000)
                End If
                If marginal = 0 Then rakassa = Int(rakassa)
        End If
        
        If marginal = 0 Then
            rakassa = Int(rakassa)
            rfack = Int(rfack)
        End If
        
        Dim SAavdrag As Double 'Inkomstunderlaget är SA inkomsten
        SAavdrag = 0
        If Application.Range("satagare") > 0 Then
            If Application.Range("satagare") <= year_(age) Then
                rakassa = 0
                rfack = 0
                'Jobbavdrag=0 'redan via lönen=0
                SAavdrag = sared(Income_(age), Skyear, marginal, Kom_skatt(age), pbb(age), 2100, IBB(age), kvoten, year_(age))
                'If marginal = 0 Then SAavdrag = Int(SAavdrag) 'sker i function sared
            End If
            'Debug.Print year_(age); age; SAavdrag
        End If
        
        Dim Pandemired As Double
        Pandemired = pandred(Wage_(age), Skyear, marginal)
        
        If (kinkskatt - SAavdrag - jobbavdrag) < SAavdrag Then SAavdrag = kinkskatt - SAavdrag - jobbavdrag
        
        Dim FAavdrag As Double
        FAavdrag = FAared(cbefvi, Skyear, marginal, pbb(age), Iyear, IBB(age), kvoten)
        
        If (kinkskatt - SAavdrag - FAavdrag - jobbavdrag) < FAavdrag Then FAavdrag = kinkskatt - SAavdrag - FAavdrag - jobbavdrag
        
        'Nettoinkomst
        Netto(age) = brutto(age) - maxi(kinkskatt + kyrkskatt + statskatt + pensionavgift - pensredukt - jobbavdrag - rakassa - rfack - SAavdrag - FAavdrag, 0)
        
        'BIDRAG
        bidragovr = 0
        
        'Barnfamiljer får barnbidrag ev, underhållsstöd och bostadsbidraget - dessa är förenklade, Ingen hänsyn till när barnet fyller år mm
        AntalBarn = CalcAntalBarn(Int(born) + age, year(barn1), year(barn2), year(barn3), year(barn4))
        
        barnbidrag = barnbidraget(AntalBarn, Utgyear) + ustod(AntalBarn, civ + 1, Utgyear) 'Ensamstående=1 i funktionen
        
        'Hyran antas följa KPI
        Dim hyra_t As Double
        If startage < (Application.Range("w_ref") - born) Then
            If year_(age) >= 1960 Then hyra_t = hyra * (KPI(age) / KPI(Application.Range("w_ref") - born))
        Else
            hyra_t = hyra * (KPI(age) / KPI(startage)) '
        End If
        'If year_(age) = 2002 Then Debug.Print year_(age) & " " & hyra_t '; KPI(age); KPI(2018 - born)
        If age < 29 Then 'ungdomar kan få bostadsbidrag ingen hänsyn till antal månader innan 29 årsdagen
            bostadsbidrag = bobid(civ + 1, brutto(age) - ptillagg(age), 0, AntalBarn, hyra_t, 80, marginal, 1, Utgyear)
        Else
            bostadsbidrag = bobid(civ + 1, brutto(age) - ptillagg(age), 0, AntalBarn, hyra_t, 80, marginal, 0, Utgyear)
        End If
        
        bostadstillägg = 0 'SA tagare? Skulle kunna nyttja informationen i PGB fliken
        SBostadstillägg = 0
        
        Dim ftid As Integer
        ftid = Application.Range("rng_Försäkringstid_vid_65")
        Dim ansoker As Single
        ansoker = Application.Range("Rng_ansokt")
        
        If def_ar > PAR And def_ar <= age Then
            uttagIP = 1
            uttagPP = 1
        End If
        'obs ska ses över för partiellt UTTAG
        If (age >= Int(PAR) And age >= riktalder And uttagIP > 0) Or wsPGB.Cells(maxi(age, 15), 5) > 0 Then
            If Iyear > 0 And Iyear <= year_(age) And (Iyear - Int(born)) >= startage And (Iyear - Int(born)) <= slutage Then
                maxhyra = IBB(age) / IBB(Iyear - Int(born))
                kvoten = pbb(Iyear - Int(born)) / IBB(Iyear - Int(born))
            End If
            'makens inkomst, AP och APM, formogenhet
            If age >= Int(PAR) Then ap = 1
            
            If civ = 0 Then
                apm = 0
                tjpm = 0
                wagem = 0
                garpm = 0
            Else
                apm = ap 'Maken antas implicit vara lika gammal som makan med motsvarande inkomster
                
                If Application.Range("rng_Makens_inkomst").Value > 0 Then 'Dvs Ej angivna
                    MakaInk = (brutto(age) - kapskatt - ptillagg(age)) * Application.Range("rng_Makens_inkomst").Value / (12 * Application.Range("Wage_Monthly").Value)
                    'garpm = garp(age) * Application.Range("rng_Makens_inkomst").Value / (12 * Application.Range("Wage_Monthly").Value)
                    garpm = gp_und * Application.Range("rng_Makens_inkomst").Value / (12 * Application.Range("Wage_Monthly").Value)
                    garpm = gp(garpm, civ, Int(born), pbb(age), Application.Range("rng_forstidM"), marginal, age, Utgyear, Iyear, IBB(age), kvoten, riktalder, uttagIP)
                    
                    tjpm = TJP(age) * Application.Range("rng_Makens_inkomst").Value / (12 * Application.Range("Wage_Monthly").Value)
                    wagem = Wage_(age) * Application.Range("rng_Makens_inkomst").Value / (12 * Application.Range("Wage_Monthly").Value)
                Else
                    MakaInk = 0
                    garpm = gp(0, civ, Int(born), pbb(age), Application.Range("rng_forstidM"), marginal, age, Utgyear, Iyear, IBB(age), kvoten, riktalder, uttagIP)
                    tjpm = 0 ' TJP(age) * MakaInk / Brutto(age) 'ren linjär framskrivning
                    wagem = 0 ' Wage_(age) * MakaInk / Brutto(age)
                End If
            End If
            If Int(def_ar) = age Then   'Korr 1/7
                'Ingen koll för helt uttag, makan går samtidigt, lönen är noll vid pensionering,
                If IsMissing(Iyear) Or Iyear < 1960 Then Iyear = Int(born) + slutage
                'If age > slutage Then age = slutage
                    bostadstillägg = uttagIP * _
                    BTP((brutto(age) - Wage_(age)) * 12 / pmonth - kapskatt - ptillagg(age), MakaInk * 12 / pmonth, 12 * hyra_t, civ, pbb(age), _
                    ap, apm, formog, 0 * Wage_(age), 0, marginal, maxhyra, Utgyear, Iyear, IBB(age), _
                    kvoten, age, TJP(age) * 12 / pmonth, tjpm * 12 / pmonth, garp(age) * 12 / pmonth, garp(age) * 12 / pmonth, _
                    IBB(Iyear - Int(born)), mpension, pmonth, ftid, Iindex(age), uttagIP, ansoker) 'tjp bör vara ev (tjp/tmonth)*
                    
                    SBostadstillägg = _
                    SBTP((brutto(age) - Wage_(age)) * 12 / pmonth - kapskatt - ptillagg(age), 12 * hyra_t, civ, bostadstillägg + bostadsbidrag, _
                    Grundavdrag, Kom_skatt(age), ap, formog, pbb(age), maxhyra, Utgyear, Iyear, IBB(age), _
                    kvoten, age, kapital, MakaInk, marginal, -99, forstid)
                    'Debug.Print "BTP "; bostadstillägg; SBostadstillägg
                Else
                    bostadstillägg = uttagIP * _
                    BTP(brutto(age) - kapskatt - ptillagg(age), MakaInk, 12 * hyra_t, civ, pbb(age), _
                    ap, apm, formog, Wage_(age), 0, marginal, maxhyra, Utgyear, Iyear, _
                    IBB(age), kvoten, age, TJP(age), tjpm, garp(age), garp(age), _
                    IBB(Iyear - Int(born)), mpension, pmonth, ftid, Iindex(age), uttagIP, ansoker)
                    
                    SBostadstillägg = _
                    SBTP(brutto(age) - kapskatt - ptillagg(age), 12 * hyra_t, civ, bostadstillägg + bostadsbidrag, Grundavdrag, _
                    Kom_skatt(age), ap, formog, pbb(age), maxhyra, Utgyear, Iyear, IBB(age), _
                    kvoten, age, kapital, MakaInk, marginal, -99, forstid)
            End If
            
            'Slå ihop bidragen och avrundning, makar antas dela på bidraget OBS! bostadstillägget delas mellan makar redan i funtionen för BTP
            If Application.Range("Rng_Ansokt") = 9 Then 'KOLLAS
                   bidragovr = bostadsbidrag + bostadstillägg
                   bostadstillägg = 0 'Spec se om ÄFS -
            End If
            If ansoker = 0 Then
                bostadstillägg = 0
                SBostadstillägg = 0
                 bostadsbidrag = 0
            End If
            bostadstillägg = (btp_sbtp(bostadstillägg, SBostadstillägg, marginal, Utgyear)) * pmonth / 12
        End If
        
        'Summera bidragen och individuell disponibel inkomst
        Bidrag(age) = barnbidrag + bostadsbidrag + bostadstillägg
        If Bidrag(age) < 0 Then Bidrag(age) = 0
        IndDisp(age) = Netto(age) + Bidrag(age) + bidragovr + pps(age)
        
        'Ekonomiskt bistånd - enligt normen + hyra OBS finns poster utanför normen som läkarvård mm
        Dim bist As Double
        bist = 0
        
        Call CalcBarnPerAlder(b1, b2, b3, b4, b5, b6, b7, b8, Int(born) + age, year(barn1), year(barn2), year(barn3), year(barn4))
        ' Utgiftsregler ej medtagna
        'If year_(age) = 2022 Then Stop
        If year_(age) < 1985 Then
            bist = 0
        ElseIf year_(age) < 2006 Then 'kvar barn och kolla utgiftsregler
            bist = bistOld(civ + 1, hyra_t, IndDisp(age), b1, b2, b3, b4, b5, b6, b7, b8, Wage_(age) * 0, year_(age))
            If AntalBarn = 0 Then bist = bist * (12 - pmonth) / 12  ' Äldres grundskydd antas tas över
            If civ = 1 And bist > 0 Then bist = bist / 2
        Else
             bist = bist25(civ + 1, hyra_t, IndDisp(age), b1, b2, b3, b4, b5, b6, b7, b8, Wage_(age) * 0, year_(age))
             If AntalBarn = 0 Then bist = bist * (12 - pmonth) / 12 ' Äldres grundskydd antas antas ta över och obs 19-20 åringer räknas inte med
             If civ = 1 And bist > 0 Then bist = bist / 2
        End If
        If bist > 0 Then
            'Antas stå till arbetsmarknadens förfogande!
            Bidrag(age) = Bidrag(age) + bist
            IndDisp(age) = IndDisp(age) + bist
        End If
        
        'Allmänna avgifter: Dagis, hemtjänst mm som är inkomstberoende vilket bidrar till ökade marginaleffekter
        'Dagis = Dagis(cbrutto(Age), AntalBarn, year_(Age))
        'Disp efter avgifter
        
        'Försörjningsbördan - KE skalan
        'Ekonomisk standard - Indisp / KE
        
        If verbose = 1 Then
            'skriver ut texten i vektorn Array_Label, 'Utskrifter - Etiketter
            Array_Label = Array("Brutto", "Avdrag_tjänst", "Tax_ink", "Grundavdrag", "Pensionavgift", "Besk_inkomst", "Kyrk_begravn", "Kommunal_skatt", "Statlig_skatt", "Skattereduktioner", " Nettoinkomst", "Bidrag", "PPS", "Ind_Disp", "Bistand")
            If age = startage Then
                Call Write_Label(Array_Label, wsOutput.Name, Kverrow, kverbose)
            End If
            
            Array_Label = Array(brutto(age), kostnadsavd, ctxfvi, Grundavdrag, pensionavgift, cbefvi, kyrkskatt, kinkskatt, statskatt, pensredukt + jobbavdrag + SAavdrag + FAavdrag, Netto(age), Bidrag(age), pps(age), IndDisp(age), bist)
            Call Write_Label(Array_Label, wsOutput.Name, age - startage + Kverrow + 1, kverbose)
            kverbose = kverbose + UBound(Array_Label)
        End If
        
         If Fasta_priser = 1 Then
            If startage <= (w_ref - Int(born)) Then
                m_value = KPI(w_ref - Int(born)) / KPI(age)
            Else
                m_value = KPI(startage) / KPI(age)
            End If
        ElseIf Fasta_priser = 0 Then
            If startage <= (w_ref - Int(born)) Then
                m_value = Iindex(w_ref - Int(born)) / Iindex(age)
            Else
                m_value = Iindex(startage) / Iindex(age)
            End If
        Else
            m_value = 1
        End If
               
        'Lagra värden i vektorn Mvalues ' Avrundas
        i = 1: If marginal = 1 Then i = 3
        mvalues(age, 1) = year_(age) '=born + age
        mvalues(age, 2) = age
        mvalues(age, 3) = Round(Income_(age) * m_value, i - 1)
        mvalues(age, 4) = Round(ip(age) * m_value, i - 1)
        mvalues(age, 5) = Round(tp(age) * m_value, i - 1)
        mvalues(age, 6) = Round(pp(age) * m_value, i - 1)
        mvalues(age, 7) = Round(garp(age) * m_value, i - 1)
        mvalues(age, 8) = Round(ptillagg(age) * m_value, i - 1)
      
        mvalues(age, 9) = Round(TJP(age) * m_value, i - 1)
        mvalues(age, 10) = Round(ips(age) * m_value, i - 1)
        mvalues(age, 11) = Round(brutto(age) * m_value, i - 1)
        mvalues(age, 12) = Round(Netto(age) * m_value, i - 1)
        mvalues(age, 13) = Round(Bidrag(age) * m_value, i - 1)
        mvalues(age, 14) = Round(pps(age) * m_value, i - 1)
        mvalues(age, 15) = Round(IndDisp(age) * m_value, i - 1)
        mvalues(age, 16) = Round(KPI(maxi(startage, w_ref - Int(born))) / KPI(age), i + 3)
        mvalues(age, 17) = Round(Iindex(maxi(startage, w_ref - Int(born))) / Iindex(age), i + 3)
      
        'Uppdatera jobbstatus 0-100 procent
        If Not Application.Range("rng_Run_From_Indata") And lngTotTasks > 0 Then
            PctDone = lngDoneTask / lngTotTasks
            Call UpdateProgress(PctDone)
            lngDoneTask = lngDoneTask + 1
        End If
        
    Next age '-----------------------------------------------------For age=startage... ----------------
    '--------------- Åren för åldrarna startage to slutage är klart ----------------------------------'
    'Skriv ut underlag till tabell 2
    With wsOutput
        Dim rngTabell2 As Range
        Dim rng As Range
        Set rng = .Range("rng_Utdata_Top_left").Offset(1).Resize(slutage - startage, UBound(mvalues, 2))  'Rad 15 till
        rng.Value = mvalues
''        For Each rngTabell2 In rng
''            If rngTabell2 = "" Or IsNull(rngTabell2) Or rngTabell2 = 0 Then
''                rngTabell2.Formula = "=NA()"
''            End If
''        Next rngTabell2
        Set rng = Nothing
        Set rngTabell2 = Nothing
    End With
    
    'Grå skala efter E(par), Expexted remaining life
    'Dim sex As Integer redan definierad
    Dim tab2_start As Integer
    tab2_start = Application.Range("rng_tabell2_startAge").Value 'Tabellen ska börja vid age=
    'tab2_start = Int(tab2_start) 'Integer
    If tab2_start > slutage Or tab2_start < startage Then
        If tab2_start < startage Then
            tab2_start = startage
            Application.Range("rng_tabell2_startAge") = startage
        Else
            tab2_start = PAR - 5
            Application.Range("rng_tabell2_startAge") = PAR
        End If
        MsgBox ("Tabellen börjar först efter slutlön eller före första lönen börjar, ändras till " & tab2_start)
    End If
    
    wsStart.Unprotect 'ta bort bladets skydd tillfälligt
    Set rng = Application.Range("rng_tabell2_top")
    
    'Rensar tabell 2 bara en tom rad kvar med rätt format
    Set rng = rng.Offset(2).Resize(110, rng.CurrentRegion.Columns.Count)
    rng.Clear
    Set rng = Application.Range("rng_tabell2_top")
    Set rng = rng.Offset(1).Resize(Application.Range("rng_tabell2_end").Row - Application.Range("rng_tabell2_top").Row + 1, rng.CurrentRegion.Columns.Count)
    rng.ClearContents
    
    ' Flyttar namnområdet rng_tabell2_top till cell K79
    Dim CellName As Name

    Set CellName = Application.Names.Item("rng_tabell2_end")
    CellName.RefersTo = "=Start!$K$" & 78 + slutage - tab2_start 'delvis fast adress
    
'    Application.ScreenUpdating = False
    Dim iRowOffset As Integer
    iRowOffset = Application.Range("rng_tabell2_top").Row + 1
    'Utskrift till Tabell 2 i start fliken -nominella, fasta priser eller uttryckt i lönenivå
    
    'Lifeincome: sum mvalues(age,14)/(1+i) from par to E(PAR)
    Dim Life0, life1, life2 As Double
    Dim delat As Integer
    delat = Application.Range("rng_Chart_Earning_factor").Value
    Life0 = 0: life1 = 0: life2 = 0
    diverse = Application.Range("rng_Exp_life") + PAR
    
    For counter = tab2_start To slutage
        With wsStart
            .Cells(counter + iRowOffset - tab2_start, 1) = mvalues(counter, 1)          'År
            .Cells(counter + iRowOffset - tab2_start, 2) = mvalues(counter, 2)          'Ålder
            .Cells(counter + iRowOffset - tab2_start, 3) = mvalues(counter, 3) / delat  'Lön rng_Chart_Earning_factor
            .Cells(counter + iRowOffset - tab2_start, 4) = (mvalues(counter, 4) + mvalues(counter, 5)) / delat 'IP/TP
            .Cells(counter + iRowOffset - tab2_start, 5) = mvalues(counter, 6) / delat  'Premie
            .Cells(counter + iRowOffset - tab2_start, 6) = (mvalues(counter, 9) + mvalues(counter, 10)) / delat 'TJP/IPS
            .Cells(counter + iRowOffset - tab2_start, 7) = (mvalues(counter, 7) + mvalues(counter, 8)) / delat 'GP/IPT
            .Cells(counter + iRowOffset - tab2_start, 8) = mvalues(counter, 11) / delat  'Brutto
            .Cells(counter + iRowOffset - tab2_start, 9) = mvalues(counter, 12) / delat  'Netto
            .Cells(counter + iRowOffset - tab2_start, 10) = mvalues(counter, 13) / delat 'Bidrag
            .Cells(counter + iRowOffset - tab2_start, 11) = mvalues(counter, 14) / delat 'PPS
            .Cells(counter + iRowOffset - tab2_start, 12) = mvalues(counter, 15) / delat 'Disp
            If counter >= PAR And Int(diverse) > counter Then
                Life0 = Life0 + mvalues(counter, 11) / ((1 + Application.Range("rng_discount")) ^ (counter - PAR))
                life1 = life1 + mvalues(counter, 12) / ((1 + Application.Range("rng_discount")) ^ (counter - PAR))
                life2 = life2 + mvalues(counter, 15) / ((1 + Application.Range("rng_discount")) ^ (counter - PAR))
                'Debug.Print counter; diverse; mvalues(counter, 11); Life0
            End If
            If Int(diverse) = counter And diverse > PAR Then
                Life0 = Life0 + (diverse - Int(diverse)) * mvalues(counter, 15) / ((1 + Application.Range("rng_discount")) ^ (counter - PAR))
                life1 = life1 + (diverse - Int(diverse)) * mvalues(counter, 15) / ((1 + Application.Range("rng_discount")) ^ (counter - PAR))
                life2 = life2 + (diverse - Int(diverse)) * mvalues(counter, 15) / ((1 + Application.Range("rng_discount")) ^ (counter - PAR))
                'Debug.Print counter; diverse; mvalues(counter, 11); Life0
            End If
            
        End With
    Next counter
    
    diverse = Application.Range("rng_Exp_life") + PAR
    diverse = Int(diverse + 0.5) + 1
    If diverse < 61 Then diverse = PAR 'för säkethetsskull
    
    If ThisWorkbook.Names("rngTurboMode").RefersToRange.Value Then 'This option is for debugging
        Application.ScreenUpdating = False
        Application.EnableEvents = False
    End If
    
    Set rng = Application.Range("rng_tabell2_top") 'Tonar ner tabell 2 och formatera alla celler med värde
    Set rng = rng.Offset(1).Resize(1, rng.CurrentRegion.Columns.Count)
    With rng.Font
        .ColorIndex = 0 'xlThemeColorLight1 'Nollställ och svart
        .TintAndShade = 0
    End With
    rng.Copy
    Set rng = rng.Offset(1).Resize(Application.Range("rng_tabell2_end").Row - Application.Range("rng_tabell2_top").Row, rng.CurrentRegion.Columns.Count)
    rng.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    Set rng = wsStart.Range("A" & diverse - tab2_start + Application.Range("rng_tabell2_top").Row + 1, "K" & slutage - tab2_start + Application.Range("rng_tabell2_top").Row + 1)
    With rng.Font
        .Color = -8808267 'Ljuslägg
        .TintAndShade = 0
    End With

    iRowOffset = Application.Range("Life_income").Row + 1
    diverse = Application.Range("Life_income").Column + 1
     'Set rng = Application.Range("Life_income")
   With wsDataTillStart
        .Cells(iRowOffset + 0, diverse) = Life0
        .Cells(iRowOffset + 1, diverse) = life1
        .Cells(iRowOffset + 2, diverse) = life2
   End With

    '---------------------------           Utskrift till tabell 1              ------------------------------------
    Dim Last_pratt As Byte
    Last_pratt = Application.Range("rng_Sista_PensRatt").Value 'Notera att de med enbart ATP (födda före 1938) påverkas inte

    Dim time_lag As Byte 'Pensionering månaden efter, direkt jämförelse
    time_lag = Application.Range("rngPens_Inflation").Value  'Fast adress

    'OBS justera för born + par > innebär att pensioneringen sker året efter
    'OBS fånga månadsutbetalningen * 12
    Dim korr As Single
    korr = 0
    If Int(PAR + born) > (Int(PAR) + Int(born)) Then korr = 1
    
    Dim justering0() As Double    'Indexering t0
    Dim justering() As Double     'Indexering t1
    
    If Fasta_priser = 1 Then
        If Application.Range("Average_Earning") > 1 Then
            ReDim justering(0 To Application.Range("Average_Earning"))
            ReDim justering0(0 To 0)
            For i = 1 To Application.Range("Average_Earning")
                justering(i) = KPI(maxi(startage, w_ref - Int(born) + korr)) / KPI(Int(PAR) - i + korr)
            Next i
        Else
            ReDim justering(0 To 0)
            ReDim justering0(0 To 0)
        End If
        wsOutput.Cells(1, 2) = "Fasta (" & w_ref & ") priser" 'OBS fast adress rad 1 kolumn B
        justering0(0) = KPI(maxi(startage, w_ref - Int(born) + korr)) / KPI(Int(PAR) - 1 + time_lag + korr)
        justering(0) = KPI(maxi(startage, w_ref - Int(born) + korr)) / KPI(Int(PAR) + korr)
    ElseIf Fasta_priser = 0 Then
        If Application.Range("Average_Earning") > 1 Then
            ReDim justering(0 To Application.Range("Average_Earning"))
            ReDim justering0(0 To 0)
            For i = 1 To Application.Range("Average_Earning")
                justering(i) = Iindex(maxi(startage, w_ref - Int(born) + korr)) / Iindex(Int(PAR) - i + korr)
            Next i
        Else
            ReDim justering(0 To 0)
            ReDim justering0(0 To 0)
        End If
        wsOutput.Cells(1, 2) = "Uttryckt i " & w_ref & " års lönenivå" 'OBS fast adress, koll att samma som ovan?
        justering0(0) = Iindex(maxi(startage, w_ref - Int(born) + korr)) / Iindex(Int(PAR) - 1 + time_lag + korr)
        justering(0) = Iindex(maxi(startage, w_ref - Int(born) + korr)) / Iindex(Int(PAR) + korr)
    Else
        wsOutput.Cells(1, 2) = "Uttryckt i nominella termer" 'OBS fast adress, koll att samma som ovan?
        If Application.Range("Average_Earning") > 1 Then
            ReDim justering(0 To Application.Range("Average_Earning"))
            ReDim justering0(0 To 0)
            For i = 1 To Application.Range("Average_Earning")
                justering(i) = 1
            Next i
        Else
            ReDim justering(0 To 0)
            ReDim justering0(0 To 0)
        End If
        justering0(0) = 1
        justering(0) = 1
    End If
    
    wsStart.Unprotect 'Oskydda startbladet för utskrift
    'Skatt och disponibel inkomst beräknas för ett helår - Sista lönen räknas om till helårsbasis...
    'Förenklar att skatter samt bidrag mm är linjärt. Nyttjar growth som linjär approximation
    'I en senare version #månader V år
    If pmonth = 12 Or Income_(Int(PAR) + korr - 1) = 0 Then
        growth = 1
    Else
        growth = (Income_(Int(PAR) + korr) * 12 / (12 - pmonth)) / Income_(Int(PAR) + korr - 1)
    End If
 
    If Application.Range("Average_Earning") > 1 Then
        Dim ys() As Double
        ReDim ys(1 To 3)
        ys(1) = Income_(Int(PAR - 1)) ' * KPI(par) / KPI(par - 1)
        ys(2) = Netto(Int(PAR - 1)) '* KPI(par) / KPI(par - 1)
        ys(3) = IndDisp(Int(PAR - 1)) '* KPI(par) / KPI(par - 1)
        For i = 2 To Application.Range("Average_Earning")
                ys(1) = ys(1) + Income_(Int(PAR - i)) 'nominella värden som presenteras i kolumn A i Tabell 1
                ys(2) = ys(2) + Netto(Int(PAR - i))
                ys(3) = ys(3) + IndDisp(Int(PAR - i))
        Next i
        With wsStart
            .Cells(Row_Tabell1_Slutlön, Col_top_tabell1) = ys(1) / Application.Range("Average_Earning")
            .Cells(Row_Tabell1_Lönefterskatt, Col_top_tabell1) = ys(2) / Application.Range("Average_Earning")
            .Cells(Row_Tabell1_Disp_inkomst, Col_top_tabell1) = ys(3) / Application.Range("Average_Earning")
        End With
        
        ys(1) = Income_(Int(PAR - 1)) * justering0(0) 'rättat 2024-04-04, värden i fasta priser, fasta löner eller nomininellt enligt eget val, som presenteras i kolumn B och C i Tabell 1
        ys(2) = Netto(Int(PAR - 1)) * justering0(0) 'rättat 2024-04-04 rättat igen 2025-01
        ys(3) = IndDisp(Int(PAR - 1)) * justering0(0) 'rättat 2024-04-04 rättat igen 2025-01
        
        For i = 2 To Application.Range("Average_Earning")
                ys(1) = ys(1) + Income_(Int(PAR - i)) * justering(i) 'rättat 2024-04-04 rättat igen 2025-01
                ys(2) = ys(2) + Netto(Int(PAR - i)) * justering(i) 'rättat 2024-04-04 rättat igen 2025-01
                ys(3) = ys(3) + IndDisp(Int(PAR - i)) * justering(i) 'rättat 2024-04-04 rättat igen 2025-01
        Next i
        ys(1) = ys(1) / Application.Range("Average_Earning")
        ys(2) = ys(2) / Application.Range("Average_Earning")
        ys(3) = ys(3) / Application.Range("Average_Earning")

        With wsStart
            .Cells(Row_Tabell1_Slutlön, Col_top_tabell1 + 1) = ys(1)
            .Cells(Row_Tabell1_Lönefterskatt, Col_top_tabell1 + 1) = ys(2)
            .Cells(Row_Tabell1_Disp_inkomst, Col_top_tabell1 + 1) = ys(3)
            
            .Cells(Row_Tabell1_Slutlön, Col_top_tabell1 + 2) = ys(1) / 12
            .Cells(Row_Tabell1_Lönefterskatt, Col_top_tabell1 + 2) = ys(2) / 12
            .Cells(Row_Tabell1_Disp_inkomst, Col_top_tabell1 + 2) = ys(3) / 12
            
            If ys(1) > 0 Then
                .Cells(Row_Tabell1_Slutlön, Col_top_tabell1 + 3) = 1
                .Cells(Row_Tabell1_Lönefterskatt, Col_top_tabell1 + 3) = ys(2) / ys(1)
                .Cells(Row_Tabell1_Disp_inkomst, Col_top_tabell1 + 3) = ys(3) / ys(1)
            Else
                .Cells(Row_Tabell1_Slutlön, Col_top_tabell1 + 3) = 0
                .Cells(Row_Tabell1_Lönefterskatt, Col_top_tabell1 + 3) = 0
                .Cells(Row_Tabell1_Disp_inkomst, Col_top_tabell1 + 3) = 0
            End If
        End With
    Else
        With wsStart
            .Cells(Row_Tabell1_Slutlön, Col_top_tabell1) = Income_(Int(PAR) - 1 + korr) * growth
            .Cells(Row_Tabell1_Lönefterskatt, Col_top_tabell1) = Netto(Int(PAR) - 1 + korr) * growth
            .Cells(Row_Tabell1_Disp_inkomst, Col_top_tabell1) = IndDisp(Int(PAR) - 1 + korr) * growth
            
            .Cells(Row_Tabell1_Slutlön, Col_top_tabell1 + 1) = Income_(Int(PAR) - 1 + korr) * justering0(0) * growth 'rättat igen 2025-01
            .Cells(Row_Tabell1_Lönefterskatt, Col_top_tabell1 + 1) = Netto(Int(PAR) - 1 + korr) * justering0(0) * growth 'rättat igen 2025-01
            .Cells(Row_Tabell1_Disp_inkomst, Col_top_tabell1 + 1) = IndDisp(Int(PAR) - 1 + korr) * justering0(0) * growth 'rättat igen 2025-01
            
            .Cells(Row_Tabell1_Slutlön, Col_top_tabell1 + 2) = (Income_(Int(PAR) - 1 + korr) * justering0(0) / 12) * growth 'rättat igen 2025-01
            .Cells(Row_Tabell1_Lönefterskatt, Col_top_tabell1 + 2) = (Netto(Int(PAR) - 1 + korr) * justering0(0) / 12) * growth 'rättat igen 2025-01
            .Cells(Row_Tabell1_Disp_inkomst, Col_top_tabell1 + 2) = (IndDisp(Int(PAR) - 1 + korr) * justering0(0) / 12) * growth 'rättat igen 2025-01
        
            If Income_(Int(PAR) - 1) > 0 Then
                .Cells(Row_Tabell1_Slutlön, Col_top_tabell1 + 3) = 1
                .Cells(Row_Tabell1_Lönefterskatt, Col_top_tabell1 + 3) = (Netto(Int(PAR) - 1 + korr) / Income_(Int(PAR) - 1)) * growth
                .Cells(Row_Tabell1_Disp_inkomst, Col_top_tabell1 + 3) = (IndDisp(Int(PAR) - 1 + korr) / Income_(Int(PAR) - 1)) * growth 'Även ev .bidrag!?
            Else
                .Cells(Row_Tabell1_Slutlön, Col_top_tabell1 + 3) = 0
                .Cells(Row_Tabell1_Lönefterskatt, Col_top_tabell1 + 3) = 0
                .Cells(Row_Tabell1_Disp_inkomst, Col_top_tabell1 + 3) = 0
            End If
       End With
       'Erase ys
    End If
    
    'KORR2 för Application.Range("rng_alt_last_pratt") >0
    Dim korr2 As Single
    korr2 = Application.Range("rng_alt_last_pratt")
    If korr2 > 0 Then
        korr = 0 'Årsdata sker senare
        PAR = Int(PAR + korr2)
        tjp_par = Int(tjp_par + korr2)
        Last_pratt = 0
        Application.Range("rng_Sista_PensRatt") = 0
        pmonth = 12 '- Int(12 * (born + par - Int(born + par)))
        Tmonth = 12 '- Int(12 * (born + tjp_par - Int(born + tjp_par)))
    End If
    
    'Pensionsutfallet
    If Last_pratt = 1 Then
        'Räkna om pension och prätter till ny pensionsbehållningen
''      IP_pbh(Int(PAR) + korr) = (12 / pmonth) * ip(Int(PAR) + korr) * deltal(PAR, Int(alt_dtal), PAR, def_ar, 4) + IP_pbh(Int(PAR) + korr)
        IP_pbh(Int(PAR) + korr) = (12 / pmonth) * ip(Int(PAR) + korr) * deltal(PAR, Int(born), PAR, def_ar, 4) + IP_pbh(Int(PAR) + korr)
        tp(Int(PAR)) = tp(Int(PAR)) * (12 / pmonth)
''        PP_pbh(Int(par) + korr) = (12 / pmonth) * pp(Int(par) + korr) * deltal(par, Int(alt_dtal), par, Def_ar, 19) + PP_ratt(Int(par) + korr)
        PP_pbh(Int(PAR) + korr) = (12 / pmonth) * pp(Int(PAR) + korr) * deltal(PAR, Int(born), PAR, def_ar, 19) + PP_ratt(Int(PAR) + korr)
''        alt_dtal = Application.Range("rng_Delningstal_födda") 'läser in igen eftersom det annars avser 65 och GARP
''        If alt_dtal < 1 Then
''           alt_dtal = born
''        ElseIf alt_dtal < 1930 Then
''           MsgBox ("Alt delningstal tyvärr inte möjligt före 1930")
''           alt_dtal = born
''        ElseIf alt_dtal < 1938 Then
''           If par <> 65 Then MsgBox ("Delningtal för " & alt_dtal & "och pensionering vid " & par & " finns inte, ersatt med 65")
''           par = 65
''           Def_ar = 65
''        End If
''        If alt_dtal < 1930 Then
''           'Förekommer inte se ovan
''        ElseIf alt_dtal < 1938 Then
''            'Enbart för pensionering vid 65
''            dtal_ip = wsNyckelTal.Cells(alt_dtal - 1930 + 9, 74) '
''            dtal_pp = wsNyckelTal.Cells(alt_dtal - 1930 + 9, 75)
''            ip(Int(par) + korr) = IP_pbh(Int(par) + korr) / dtal_ip
''            pp(Int(par) + korr) = PP_pbh(Int(par) + korr) / dtal_pp
''        Else
            'Ev justering för prognos
''            ip(Int(par) + korr) = IP_pbh(Int(par) + korr) / deltal(par, Int(alt_dtal), par, Def_ar, 4)
            ip(Int(PAR) + korr) = IP_pbh(Int(PAR) + korr) / deltal(PAR, Int(born), PAR, def_ar, 4)
''            dtal_pp = deltal(par, Int(alt_dtal), par, Def_ar, 19)
            dtal_pp = deltal(PAR, Int(born), PAR, def_ar, 19)
            wsStart.Range("m10") = dtal_pp 'OBS fast adress
            pp(Int(PAR) + korr) = PP_pbh(Int(PAR) + korr) / dtal_pp  'Ska ev korrigeras för ränteantagandet
            'Garantipensionens delningstal
''            dtal_ip = deltal(riktalder, Int(alt_dtal), riktalder, riktalder, 4)
            dtal_ip = deltal(riktalder, Int(born), riktalder, riktalder, 4)
            If Application.Range("Mortality") = 1 Then dtal_ip = wsMortality.Range("L5").Value
''        End If
        PAR = Int(PAR + korr)
        age = PAR
        
        If dtal_ip > 0 And PAR >= riktalder Then
            'gp_und = (GP_pbh(Int(par + korr) - 1) + GP_pbh(Int(par + korr))) / dtal_ip + tp(par+korr)
            gp_und = IP_(year_(PAR), PAR, born, GP_ratt(PAR), IP_arv1(PAR), IP_arv2(PAR), IP_avg(PAR), GP_pbh(PAR - 1), 1, _
                         garp(PAR - 1) + 0, dtal_ip, Pindex(PAR) / Pindex(PAR - 1), def_ar, marginal, 0, Pindex(PAR) / Iindex(PAR)) '
                         '+ tp (par) * maxi(1, tp_faktor(par))
            gp_und = gp_und + GP_pbh(Int(PAR + korr)) / dtal_ip
            
            If marginal = 0 Then gp_und = Int((gp_und + 0.49) / 12) * 12
            
            garp(Int(PAR + korr)) = gp(gp_und + tp(Int(PAR) + korr), civ, Int(born), pbb(Int(PAR + korr)), forstid, marginal, Int(PAR), year_(Int(PAR + korr)), _
                    Iyear, IBB(Int(PAR + korr)), kvoten, riktalder, uttagIP)
            
            If Application.Range("RulesfromUtg") = 0 Or IsMissing(Application.Range("RulesfromUtg")) Then
            ElseIf (Int(born + PAR) > Application.Range("RulesfromUtg") And Application.Range("Rules") = 0) Or _
                        Application.Range("Rules") = 1 Then
                garp(Int(PAR + korr)) = gp(gp_und + tp(Int(PAR) + korr), civ, Int(born), pbb(Int(PAR + korr)), forstid, marginal, Int(PAR), Application.Range("RulesfromUtg"), _
                Iyear, IBB(Int(PAR + korr)), kvoten, riktalder, uttagIP)
            End If
        Else
            garp(Int(PAR + korr)) = 0   'Kan komma att få garp under 12 månaders perioden om t.ex. par=64,5 och riktålder är 65,
                                        'samt om underlaget efter årsskiftet beviljar garantipension...
        End If
        'ptillagg(Int(par + korr)) = 0
    
        mpension = gpundtab1 * 12 'gp_und / 12 '(ip(par + korr) * 185 / 160 + tp(par + korr))
        If (2021 - Int(born)) > startage Then 'Tillfälligt
            ptillagg(Int(PAR + korr)) = tillagg(mpension, year_(Int(PAR + korr)), Iindex(Int(PAR + korr)), Iindex(2021 - Int(born)), uttagIP, pgi_years)
        Else
            ptillagg(Int(PAR + korr)) = tillagg(mpension, year_(Int(PAR + korr)), Iindex(Int(PAR + korr)), 182.58, uttagIP, pgi_years)
        End If
        
        If year_(Int(PAR + korr)) = 2021 Then _
                ptillagg(Int(PAR + korr)) = ptillagg(Int(PAR + korr)) * mini(4, pmonth) / 12 'införs 1/9
        
         If IsMissing(Application.Range("RulesfromUtg")) Or Application.Range("RulesfromUtg") = 0 Then
            'ovan
             Else
                If (year_(age) > Application.Range("RulesfromUTG") And Application.Range("Rules") = 0) _
                    Or Application.Range("Rules") = 1 Then
                    
                        ptillagg(age) = tillagg(12 * mpension, Application.Range("RulesfromUTG"), Iindex(age), 186.52, uttagIP, pgi_years)
                    
                    If Application.Range("RulesfromUTG") = 2021 Then ptillagg(age) = ptillagg(age) * mini(4, pmonth) / 12
                    
                    'Justera beloppet?
                 End If
            End If
        
        If marginal = 0 Then
            ip(Int(PAR) + korr) = Int((ip(Int(PAR) + korr) + 0.49) / 12) * 12
            pp(Int(PAR) + korr) = Int((pp(Int(PAR) + korr) + 0.49) / 12) * 12
        End If
        'Debug.Print "tabell 1"; year_(par + korr); Int(mpension / 12); ptillagg(par + korr)
    End If 'LAST_pratt ..................
''    par = par + korr2
''    tjp_par = tjp_par + korr2
    Income_(PAR) = 0   'Nollställer bruttoinkomsten Kollas om både arbete och pension
    Wage_(PAR) = 0
               
    If pmonth < 12 Then
        If Last_pratt = 0 Then
            ip(PAR) = ip(PAR) + (ip(PAR + 1) / 12) * (12 - pmonth)
            tp(PAR) = tp(PAR) + (tp(PAR + 1) / 12) * (12 - pmonth)
            pp(PAR) = pp(PAR) + (pp(PAR + 1) / 12) * (12 - pmonth)
            garp(PAR) = garp(PAR) + (garp(PAR + 1) / 12) * (12 - pmonth)
            ptillagg(PAR) = ptillagg(PAR) + (ptillagg(PAR + 1) / 12) * (12 - pmonth)
        End If
    End If
    
    If Int(tjp_par) = Int(PAR) Then
        If Tmonth < 12 Then
            TJP(PAR) = TJP(PAR) + (TJP(PAR + 1) / 12) * (12 - Tmonth)
            ips(PAR) = ips(PAR) + (ips(PAR + 1) / 12) * (12 - Tmonth)
        Else
            'OK
        End If
    ElseIf Int(tjp_par) > Int(PAR) Then
            'ok tjp(par)=0
    Else 'If Int(tjp_par) < Int(par) Then
            'ok TJP(par)>0 0 och tmonth=12
    End If
    
    brutto(Int(PAR)) = Income_(Int(PAR)) + ip(Int(PAR)) + tp(Int(PAR)) + pp(Int(PAR)) + garp(Int(PAR)) + TJP(Int(PAR)) + ips(Int(PAR)) _
    + ptillagg(Int(PAR)) '+ kapital
        
    'Utskrift till Tabell 1 i start fliken
    With wsStart
        .Cells(Row_Tabell1_Tot_Brutto, Col_top_tabell1) = brutto(Int(PAR))
        .Cells(Row_Tabell1_Tot_Allmänpension, Col_top_tabell1) = ip(PAR) + tp(PAR) + pp(PAR) + garp(PAR) + ptillagg(PAR)
        .Cells(Row_Tabell1_IP, Col_top_tabell1) = ip(PAR)
        .Cells(Row_Tabell1_TP, Col_top_tabell1) = tp(PAR)
        .Cells(Row_Tabell1_PP, Col_top_tabell1) = pp(PAR)
        .Cells(Row_Tabell1_GP, Col_top_tabell1) = garp(PAR)
        .Cells(Row_Tabell1_PT, Col_top_tabell1) = ptillagg(PAR)
        .Cells(Row_Tabell1_TJP, Col_top_tabell1) = TJP(PAR)
        .Cells(Row_Tabell1_IPS, Col_top_tabell1) = ips(PAR)
        .Cells(Row_Tabell1_PPS, Col_top_tabell1) = pps(PAR)
        
        .Cells(Row_Tabell1_Tot_Brutto, Col_top_tabell1 + 1) = brutto(Int(PAR)) * justering(0)
        .Cells(Row_Tabell1_Tot_Allmänpension, Col_top_tabell1 + 1) = (ip(PAR) + tp(PAR) + pp(PAR) + garp(PAR) + ptillagg(PAR)) * justering(0)
        .Cells(Row_Tabell1_IP, Col_top_tabell1 + 1) = ip(PAR) * justering(0)
        .Cells(Row_Tabell1_TP, Col_top_tabell1 + 1) = tp(PAR) * justering(0)
        .Cells(Row_Tabell1_PP, Col_top_tabell1 + 1) = pp(PAR) * justering(0)
        .Cells(Row_Tabell1_GP, Col_top_tabell1 + 1) = garp(PAR) * justering(0)
        .Cells(Row_Tabell1_PT, Col_top_tabell1 + 1) = ptillagg(PAR) * justering(0)
        .Cells(Row_Tabell1_TJP, Col_top_tabell1 + 1) = TJP(PAR) * justering(0)
        .Cells(Row_Tabell1_IPS, Col_top_tabell1 + 1) = ips(PAR) * justering(0)
        .Cells(Row_Tabell1_PPS, Col_top_tabell1 + 1) = pps(PAR) * justering(0)
        
        .Cells(Row_Tabell1_Tot_Brutto, Col_top_tabell1 + 2) = brutto(Int(PAR)) * justering(0) / 12
        .Cells(Row_Tabell1_Tot_Allmänpension, Col_top_tabell1 + 2) = (ip(PAR) + tp(PAR) + pp(PAR) + garp(PAR) + ptillagg(PAR)) * justering(0) / 12
        .Cells(Row_Tabell1_IP, Col_top_tabell1 + 2) = ip(PAR) * justering(0) / 12
        .Cells(Row_Tabell1_TP, Col_top_tabell1 + 2) = tp(PAR) * justering(0) / 12
        .Cells(Row_Tabell1_PP, Col_top_tabell1 + 2) = pp(PAR) * justering(0) / 12
        .Cells(Row_Tabell1_GP, Col_top_tabell1 + 2) = garp(PAR) * justering(0) / 12
        .Cells(Row_Tabell1_PT, Col_top_tabell1 + 2) = ptillagg(PAR) * justering(0) / 12
        .Cells(Row_Tabell1_TJP, Col_top_tabell1 + 2) = TJP(PAR) * justering(0) / 12
        .Cells(Row_Tabell1_IPS, Col_top_tabell1 + 2) = ips(PAR) * justering(0) / 12
        .Cells(Row_Tabell1_PPS, Col_top_tabell1 + 2) = pps(PAR) * justering(0) / 12
        'OBS att Income_(par) är satt till noll
        
        If Application.Range("Average_Earning") > 1 Then
            Income_(PAR - 1) = ys(1) 'Om slutlön avser flera år
        Else
            Income_(PAR - 1) = (Income_(PAR - 1 - korr2) * justering0(0)) 'rättat genom korr2 2024-02 rättat igen med array för justering rättat igen 2025-01
        End If
        
        If Income_(PAR - 1) > 0 Then
            .Cells(Row_Tabell1_Tot_Brutto, Col_top_tabell1 + 3) = brutto(Int(PAR)) * justering(0) / Income_(PAR - 1)
            .Cells(Row_Tabell1_Tot_Allmänpension, Col_top_tabell1 + 3) = (ip(PAR) + tp(PAR) + pp(PAR) + garp(PAR) + ptillagg(PAR)) * justering(0) / Income_(PAR - 1)
            .Cells(Row_Tabell1_IP, Col_top_tabell1 + 3) = ip(PAR) * justering(0) / Income_(PAR - 1)
            .Cells(Row_Tabell1_TP, Col_top_tabell1 + 3) = tp(PAR) * justering(0) / Income_(PAR - 1)
            .Cells(Row_Tabell1_PP, Col_top_tabell1 + 3) = pp(PAR) * justering(0) / Income_(PAR - 1)
            .Cells(Row_Tabell1_GP, Col_top_tabell1 + 3) = garp(PAR) * justering(0) / Income_(PAR - 1)
            .Cells(Row_Tabell1_PT, Col_top_tabell1 + 3) = ptillagg(PAR) * justering(0) / Income_(PAR - 1)
            .Cells(Row_Tabell1_TJP, Col_top_tabell1 + 3) = TJP(PAR) * justering(0) / Income_(PAR - 1)
            .Cells(Row_Tabell1_IPS, Col_top_tabell1 + 3) = ips(PAR) * justering(0) / Income_(PAR - 1)
            .Cells(Row_Tabell1_PPS, Col_top_tabell1 + 3) = pps(PAR) * justering(0) / Income_(PAR - 1)
        Else
            .Cells(Row_Tabell1_Tot_Brutto, Col_top_tabell1 + 3) = 0
            .Cells(Row_Tabell1_Tot_Allmänpension, Col_top_tabell1 + 3) = 0
            .Cells(Row_Tabell1_IP, Col_top_tabell1 + 3) = 0
            .Cells(Row_Tabell1_TP, Col_top_tabell1 + 3) = 0
            .Cells(Row_Tabell1_PP, Col_top_tabell1 + 3) = 0
            .Cells(Row_Tabell1_GP, Col_top_tabell1 + 3) = 0
            .Cells(Row_Tabell1_PT, Col_top_tabell1 + 3) = 0
            .Cells(Row_Tabell1_TJP, Col_top_tabell1 + 3) = 0
            .Cells(Row_Tabell1_IPS, Col_top_tabell1 + 3) = 0
            .Cells(Row_Tabell1_PPS, Col_top_tabell1 + 3) = 0
        End If
    End With
    'Debug.Print pgi_years; Int(mpension / 12); ptillagg(par)
''    If Application.Range("sysLang") = 0 Then
''        Range("rng_tabell1_pt") = "Pensionstillägg (" & mini(pgi_years, 40) & "/40)"
''    Else
''        Range("rng_tabell1_pt") = "Added incomepension (IPT) (" & mini(pgi_years, 40) & "/40)"
''    End If
 'KORR Kolla om dessa rader med netto behövs om last_pratt=0 ...
 Range("rng_wyears") = mini(pgi_years, 40)
 
 If Last_pratt > 0 Then
        'Netto och disponibel inkomst
        ctxfvi = brutto(Int(PAR)) 'inga kostnadsavdrag som pensionär och kapitalinkomsterna är inte inkl
        If marginal = 0 Then
            ctxfvi = Int(ctxfvi / 100) * 100
        End If
     
        Grundavdrag = avdragxx(ctxfvi, pbb(Int(PAR)), marginal, Int(PAR), year_(Int(PAR)), Iyear, IBB(Int(PAR)), kvoten, year_(mini(age, 100)), Gage)
            
        pensionavgift = 0 'Ingen lön antas
        'Beskattningsbar inkomst
        cbefvi = ctxfvi - Grundavdrag '- Pensionavgift + PGI(year_(Age), Income_(Age), pbb(Age), IBB(Age), FPB(Age), marginal, 2, Age, 0)
        'Kommunal inkomstskatt
        kinkskatt = cbefvi * Kom_skatt(Int(PAR))
        kyrkskatt = cbefvi * Begravavg(Int(PAR))
        
        If marginal = 0 Then
            kinkskatt = Int(kinkskatt)
            kyrkskatt = Int(kyrkskatt)
        End If
        statskatt = statlig(cbefvi, Tax_limit1(Int(PAR)), Tax_limit2(Int(PAR)), marginal) 'Statlig inkomstskatt
        If Application.Range("RulesfromSkatt") = 0 Or IsMissing(Application.Range("RulesfromSkatt")) _
        Or year_(PAR) < Application.Range("RulesfromSkatt") Then
            statskatt = statskatt + PublicAvg(cbefvi, 0.01, age, marginal, year_(Int(PAR)))
        Else
            statskatt = statskatt + PublicAvg(cbefvi, 0.01, age, marginal, Application.Range("RulesfromSkatt"))
        End If
    
        If kapital < -100000 Then
            kapskatt = -30000 + ((kapital + 100000) * 0.7 * 0.3)
        Else
            kapskatt = kapital * 0.3
        End If
        If kapskatt < 0 Then
            If statskatt > (-kapskatt) Then
                statskatt = statskatt + kapskatt
            Else
                kinkskatt = kinkskatt + statskatt + kapskatt
                statskatt = 0
            End If
        Else
            statskatt = statskatt + kapskatt
        End If
        If marginal = 0 Then statskatt = Int(statskatt)
        'Förmögenhets- och fastighetsskatt samt reduktioner för dessa är inte med liksom kapitalskatt
        pensredukt = rakassa 'PGI(year_(Age), Ctxfvi, pbb(Age), IBB(Age), FPB(Age), marginal, 2, Age, 0)
        jobbavdrag = 0
        FAavdrag = 0
        
        If Application.Range("RulesfromSkatt") = 0 Or IsMissing(Application.Range("RulesfromSkatt")) _
        Or year_(PAR) < Application.Range("RulesfromSkatt") Then
            FAavdrag = FAared(cbefvi, year_(Int(PAR)), marginal, pbb(Int(PAR)), Iyear, IBB(Int(PAR)), kvoten)
        Else
            FAavdrag = FAared(cbefvi, Application.Range("RulesfromSkatt"), marginal, pbb(Int(PAR)), Iyear, IBB(Int(PAR)), kvoten)
        End If
        If (kinkskatt - SAavdrag - FAavdrag - jobbavdrag) < FAavdrag Then FAavdrag = kinkskatt - SAavdrag - FAavdrag - jobbavdrag
        Netto(Int(PAR)) = brutto(Int(PAR)) + kapital - maxi(kinkskatt + kyrkskatt + statskatt + pensionavgift - pensredukt - jobbavdrag - FAavdrag, 0)
        'Netto - Utskrift till Tabell 1 i start fliken
    'KORR Hit om dessa rader med netto behövs om last_pratt=0 ... men även disponibelt torde vara klar
    End If
    With wsStart
        .Cells(Row_Tabell1_Efterskatt, Col_top_tabell1) = Netto(Int(PAR))
        .Cells(Row_Tabell1_Efterskatt, Col_top_tabell1 + 1) = Netto(Int(PAR)) * justering(0)
        .Cells(Row_Tabell1_Efterskatt, Col_top_tabell1 + 2) = Netto(Int(PAR)) * justering(0) / 12
        
        If Application.Range("Average_Earning") > 1 Then
            Netto(Int(PAR - 1)) = ys(2) 'Om pension togs x år framåt
        Else
            Netto(Int(PAR - 1)) = Netto(Int(PAR - 1)) * justering0(0) ' rättat med array 2025-01
        End If
        
        If Netto(Int(PAR) - 1) > 0 Then 'OBS KORR
            .Cells(Row_Tabell1_Efterskatt, Col_top_tabell1 + 3) = Netto(Int(PAR)) * justering(0) / Netto(Int(PAR - 1))
        Else
            .Cells(Row_Tabell1_Efterskatt, Col_top_tabell1 + 3) = 0
        End If
    End With
    If Last_pratt > 0 Then
    ''    'Barn- och Bostadsbidrag påverkas inte av sista p-rätten, däremot BTP, mm f(inkomst )
    ''    'Antal barn beräknar barnbidrag underhållsstöd, dessa regler är inte uppdaterade liksom bostadsbidraget
    ''    'Ingen koll på antal månader under året month(barni)
    ''    If year_(Int(par)) >= year(barn1) And year_(Int(par)) < year(barn1) + 20 Then
    ''        If year_(Int(par)) = year(barn1) Then AntalBarn = AntalBarn + 1
    ''        If year_(Int(par)) = (year(barn1) + 19) Then
    ''            AntalBarn = AntalBarn - 1
    ''        End If
    ''    End If
    ''    'Ev barn2
    ''    If year_(Int(par)) >= year(barn2) And year_(Int(par)) < year(barn2) + 20 Then
    ''        If year_(Int(par)) = year(barn2) Then AntalBarn = AntalBarn + 1
    ''        If year_(Int(par)) = (year(barn2) + 19) Then
    ''            AntalBarn = AntalBarn - 1
    ''        End If
    ''    End If
    ''    'Ev barn3
    ''    If year_(Int(par)) >= year(barn3) And year_(Int(par)) < year(barn3) + 20 Then
    ''        If year_(Int(par)) = year(barn3) Then AntalBarn = AntalBarn + 1
    ''        If year_(Int(par)) = (year(barn3) + 19) Then
    ''            AntalBarn = AntalBarn - 1
    ''        End If
    ''    End If
    ''    'Ev barn4
    ''    If year_(Int(par)) >= year(barn4) And year_(Int(par)) < year(barn3) + 20 Then
    ''        If year_(Int(par)) = year(barn4) Then AntalBarn = AntalBarn + 1
    ''        If year_(Int(par)) = (year(barn4) + 19) Then
    ''            AntalBarn = AntalBarn - 1
    ''        End If
    ''    End If
    ''
    ''    barnbidrag = barnbidraget(AntalBarn, year_(Int(par))) + ustod(AntalBarn, civ)
    ''
    ''    bostadsbidrag = bobid(civ + 1, Brutto(Int(par)), 0, AntalBarn, hyra_t, 80, marginal, 0, year_(Int(par))) 'Kan påverkas men sälsynt
    ''
        bostadstillägg = 0
        If PAR >= riktalder And uttagIP = 1 And uttagPP = 1 Then
             'Maxhyra - om kvoten <> 1
            If Iyear > 0 And Iyear <= year_(Int(PAR)) Then
                maxhyra = (pbb(Int(PAR)) / pbb(Iyear - Int(born))) 'else maxhyra default 0
            End If
            ' If age > slutage Then age = slutage
            If Application.Range("RulesfromUtg") = 0 Or IsMissing(Application.Range("RulesfromUtg")) _
            Or year_(PAR) < Application.Range("RulesfromUtg") Then
                ansoker = Application.Range("Rng_Ansokt")
                bostadstillägg = BTP((brutto(Int(PAR)) - Wage_(Int(PAR))) * 12 / pmonth - kapskatt - ptillagg(Int(PAR)), MakaInk * 12 / pmonth, 12 * hyra_t, civ, pbb(Int(PAR)), _
                        ap, apm, formog, 0 * Wage_(Int(PAR)), 0, marginal, maxhyra, year_(Int(PAR)), Iyear, IBB(Int(PAR)), _
                        kvoten, Int(PAR), TJP(Int(PAR)) * 12 / pmonth, tjpm * 12 / pmonth, garp(Int(PAR)) * 12 / pmonth, garp(Int(PAR)) * 12 / pmonth, _
                        IBB(Iyear - Int(born)), mpension / 12, pmonth, ftid, Iindex(Int(PAR)), uttagIP, ansoker) 'tjp bör vara ev (tjp/tmonth)*
                SBostadstillägg = SBTP(brutto(Int(PAR)) + kapital - kapskatt - ptillagg(Int(PAR)), hyra_t, civ, bostadstillägg + bostadsbidrag, Grundavdrag, Kom_skatt(Int(PAR)), ap, formog, _
                        pbb(Int(PAR)), maxhyra, year_(Int(PAR)), Iyear, IBB(Int(PAR)), kvoten, Int(PAR), 0, MakaInk, marginal, -99, forstid)
            Else
                bostadstillägg = BTP((brutto(Int(PAR)) - Wage_(Int(PAR))) * 12 / pmonth - kapskatt - ptillagg(Int(PAR)), MakaInk * 12 / pmonth, 12 * hyra_t, civ, pbb(Int(PAR)), _
                        ap, apm, formog, 0 * Wage_(Int(PAR)), 0, marginal, maxhyra, Application.Range("RulesfromUtg"), Iyear, IBB(Int(PAR)), _
                        kvoten, Int(PAR), TJP(Int(PAR)) * 12 / pmonth, tjpm * 12 / pmonth, garp(Int(PAR)) * 12 / pmonth, garp(Int(PAR)) * 12 / pmonth, _
                        IBB(Iyear - Int(born)), mpension / 12, pmonth, ftid, Iindex(Int(PAR)), uttagIP, ansoker) 'tjp bör vara ev (tjp/tmonth)
                SBostadstillägg = SBTP(brutto(Int(PAR)) + kapital - kapskatt - ptillagg(Int(PAR)), hyra_t, civ, bostadstillägg + bostadsbidrag, Grundavdrag, Kom_skatt(Int(PAR)), ap, formog, _
                    pbb(Int(PAR)), maxhyra, Application.Range("RulesfromUtg"), Iyear, IBB(Int(PAR)), kvoten, Int(PAR), 0, MakaInk, marginal, -99, forstid)
            End If
              If Application.Range("Rng_Ansokt") = 9 Then
                       bidragovr = bostadsbidrag + bostadstillägg
                        bostadstillägg = 0 'Spec se om ÄFS
                End If
            If ansoker = 0 Then SBostadstillägg = 0
            bostadstillägg = btp_sbtp(bostadstillägg, SBostadstillägg, marginal, year_(mini(age, slutage)))
        End If
        'Summa bidrag
        Bidrag(Int(PAR)) = barnbidrag + bostadsbidrag + bostadstillägg
    End If
    'Utskrift av bidrag och disponibel inkomst till fliken wsDataTillStart
    'Respektavstånd jämför typisen mot ensamstående utan inkomst
    Dim resp As Double
    'Respektavstånd 1- Grundskydd/ typisens SKA skickas till ett ett ark för att hämtas där...
    If Application.Range("wealth") = 1 Then 'Visas enbart om "wealth"
        If Application.Range("resp_tjp") = 1 Then
          resp = respekt(brutto(Int(PAR)), hyra_t, TJP(Int(PAR)), garp(Int(PAR)), ptillagg(Int(PAR)))
        Else
          resp = respekt(brutto(Int(PAR)) - TJP(Int(PAR)), hyra_t, 0, garp(Int(PAR)), ptillagg(Int(PAR)))
        End If
    End If
    
    With wsStart
        .Cells(Row_Tabell1_Bidrag, Col_top_tabell1) = Bidrag(Int(PAR))
        .Cells(Row_Tabell1_Bidrag, Col_top_tabell1 + 1) = Bidrag(Int(PAR)) * justering(0)
        .Cells(Row_Tabell1_Bidrag, Col_top_tabell1 + 2) = Bidrag(Int(PAR)) * justering(0) / 12
        If Application.Range("Average_Earning") > 1 Then
            IndDisp(Int(PAR - 1)) = ys(3)
        Else
            IndDisp(Int(PAR - 1)) = IndDisp(Int(PAR - 1)) * justering0(0) 'rättat med array 2025-01
        End If
        If Bidrag(Int(PAR - 1)) > 0 Then
             .Cells(Row_Tabell1_Bidrag, Col_top_tabell1 + 3) = Bidrag(Int(PAR)) * justering(0) / IndDisp(Int(PAR - 1))  '(Bidrag(Int(par - 1)) * justering0)
        Else
             .Cells(Row_Tabell1_Bidrag, Col_top_tabell1 + 3) = 0 'NA()
        End If
        IndDisp(Int(PAR)) = Netto(Int(PAR)) + Bidrag(Int(PAR)) + bidragovr + pps(Int(PAR))
        .Cells(Row_Tabell1_Disp_efterskatt, Col_top_tabell1) = IndDisp(Int(PAR))
        .Cells(Row_Tabell1_Disp_efterskatt, Col_top_tabell1 + 1) = IndDisp(Int(PAR)) * justering(0)
        .Cells(Row_Tabell1_Disp_efterskatt, Col_top_tabell1 + 2) = IndDisp(Int(PAR)) * justering(0) / 12
''        If Application.Range("Average_Earning") > 1 Then
''            IndDisp(Int(par - 1)) = ys(3)
''        Else
''            IndDisp(Int(par - 1)) = IndDisp(Int(par - 1)) * justering0
''        End If
        If IndDisp(Int(PAR - 1)) > 0 Then
            .Cells(Row_Tabell1_Disp_efterskatt, Col_top_tabell1 + 3) = IndDisp(Int(PAR)) * justering(0) / IndDisp(Int(PAR - 1))
        Else
            .Cells(Row_Tabell1_Disp_efterskatt, Col_top_tabell1 + 3) = 0
        End If
''         If Application.Range("wealth") = 1 Then
''
''            If Application.Range("Rng_resp_rel") = 1 Then
''                Sheets("Start").Select
''                Range("I42:I44").Select
''                Selection.Style = "Percent"
''                Selection.NumberFormat = "0.0%"
''            Else
''                Sheets("Start").Select
''                Range("I42:I44").Select
''                Selection.NumberFormat = "#,##0"
''                'Rows("46:46").RowHeight = 12
''            End If
''        Else
''''            'Rows("46:46").RowHeight = 12
''''           .Cells(Row_Tabell1_Efterskatt, Col_top_tabell1 + 6) = ""
''''           .Cells(Row_Tabell1_Efterskatt + 1, Col_top_tabell1 + 6) = ""
''''
''''           .Cells(Row_Tabell1_Disp_efterskatt, Col_top_tabell1 + 6) = ""
''''           .Cells(Row_Tabell1_Disp_efterskatt, Col_top_tabell1 + 9) = ""
''
''
''        End If
    End With
    'wsStart.Protect 'Skydda bladet i slutversionen
    Unload UserFormProgress  'Tar bort formuläret
    With wsOutput
        If verbose = 1 Then
            Set rnglbl = Application.Range("rng_Utdata_Top_left_verbose") 'Se kolumn AA
            Set rnglbl = rnglbl.Resize(1, rnglbl.CurrentRegion.Columns.Count)
            'Call Format_Range(rnglbl, True, 65535) 'Formatera etikettraden
            Set rnglbl = Nothing
        End If
    End With
   'Frigör minne
    Erase Array_Label
    Erase Mlabels
    Erase mvalues
    Erase ys
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    If Application.Range("rng_Run_From_Indata") = False Then
        If ActiveSheet.Name <> Application.Range("rngModelHeader").Parent.Name Then
            Sheets(Application.Range("rngModelHeader").Parent.Name).Activate
        End If
        Application.Range("rngModelHeader").Select
    End If
End Sub


Sub Write_Label(ByVal Arr, ByVal Ws As String, ByVal iRow As Long, ByVal kverbose As Long)
    Dim i As Long
    'För att hantera nollåringar
    If iRow = 0 Then iRow = 1
    For i = 1 To UBound(Arr)
        Sheets(Ws).Cells(iRow, kverbose + i - 1) = Arr(i)
    Next i
End Sub

Sub Format_Range(ByVal rng As Range, Optional blnFontformatBold As Boolean, Optional lngrngColor As Long)
    rng.Font.Bold = blnFontformatBold
    rng.Interior.Color = lngrngColor
End Sub



