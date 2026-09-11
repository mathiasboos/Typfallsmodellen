Attribute VB_Name = "wsArv_ip"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
' Arvsvinst IP
'Nedan ges kod för prognos 2021 och senare. Utgår från 2020 och kan behöva ses över nästa år 2022
'Den är förbered för regressionsansats OLS
'Men prognosen utgår från senaste arvsvinst (modellerad) och sedan kvotas med förändringen i dödlighet
'
'Option Explicit                     'Alla variabler ska definieras
'Option Base 1                       'Indexering börjar med 1 och inte med noll
'
''Läser in årsvisa arvsvinster för x åringar
'Function arvxx(ByVal age As Single, ByVal arvslut As Single, Optional ByVal arvstart = 2000, Optional verbose = 0, Optional dd = 0) As Double
'    'Age        - Ålder
'    'Arvslut    - Senaste faställda arvsvinst
'    'Arvsstart  - Prognos början
'    'Verbose    - Debug.print på prognos mm
'    'DD         - 1=om den senare dubbla arvsvinsten (60 och senare med förändring i rktålder) ska läsas in
'
'    If IsMissing(arvslut) Then arvslut = 2020
'    If IsMissing(arvstart) Then arvstart = arvslut 'och ingen regressionspred för arvsslut
'    If arvstart < 2000 Then arvstart = 2000
'
'    Dim yarv() As Double
'
'    ReDim yarv(arvstart To arvslut) As Double
'    'Utgå från prognos?
'    Dim XX() As Double
'    ReDim XX(arvstart To arvslut, 1 To 2)
'
'    Dim t As Single
'
'    'Matriser för OLS
'    Dim xpx() As Variant
'    ReDim xpx(1 To 2, 1 To 2)
'    Dim xy() As Double
'    ReDim xy(1 To 2)
'    Dim beta() As Double
'    ReDim beta(1 To 2)
'
'    Dim ral As Single 'Koll på arvsvinstfaktor 1 resp 2
'    Dim dyarv As Double
'
'    'Ingen hantering av missing values
'    For t = arvstart To arvslut
'        ral = riktage(t - 1, 0) 'Först året efter som riktålder påverkar ålder för arvsfaktor
'        If age <= (ral - 1) Then
'            yarv(t) = wsArv_ip.Cells(age - 17 + 3, t + 4 - 2000) 'för 2020 <=60
'            If ral = age And dd = 1 Then yarv(t) = wsArv_ip.Cells(81 + age - 60, t + 4 - 2000) '=60
'        Else
'            yarv(t) = wsArv_ip.Cells(81 + age - 60, t + 4 - 2000) '2020>60
'        End If
'
'        'if yarv(t)>0 then
'
'            XX(t, 1) = 1 'Intercept
'            XX(t, 2) = t + 1 - arvstart
'
'            xpx(1, 1) = xpx(1, 1) + XX(t, 1) ^ 2
'            xpx(2, 1) = xpx(2, 1) + XX(t, 1) * XX(t, 2)
'            xpx(1, 2) = xpx(1, 2) + XX(t, 2) * XX(t, 1)
'            xpx(2, 2) = xpx(2, 2) + XX(t, 2) * XX(t, 2)
'            xy(1) = xy(1) + XX(t, 1) * yarv(t)
'            xy(2) = xy(2) + XX(t, 2) * yarv(t)
'        'end if
'    Next t
'    Dim d As Double
'    Dim xpxi() As Variant
'    ReDim xpxi(1 To 2, 1 To 2)
'
'    If arvslut < arvstart Then 'Ingen prognos
'
'        d = xpx(1, 1) * xpx(2, 2) - xpx(1, 2) * xpx(2, 1)
'
'        xpxi(1, 1) = xpx(2, 2) / d
'        xpxi(2, 1) = -xpx(2, 1) / d
'        xpxi(1, 2) = -xpx(1, 2) / d
'        xpxi(2, 2) = xpx(1, 1) / d
'
'        beta(1) = xpxi(1, 1) * xy(1) + xpxi(2, 1) * xy(2)
'        beta(2) = xpxi(1, 2) * xy(1) + xpxi(2, 2) * xy(2)
'        t = xpx(1, 1)
'        arvxx = beta(1) + beta(2) * t
'        If verbose = 1 Then
'            Debug.Print "snitt: " & Round(xy(1) / xpx(1, 1), 4); " Beta "; Round(beta(1), 4) & " " & Round(beta(2), 4); arvxx
'        End If
'
'        If age > 60 Then
'            Dim rad As Integer
'            ReDim xy(1 To 3)
'            For rad = 17 To 60
'                xy(1) = xy(1) + yarv(rad)
'                xy(2) = xy(2) + rad * yarv(rad)
'                xy(3) = XX(3) + (rad * rad) * yarv(rad)
'            Next rad
'
'            ReDim xpxi(1 To 3, 1 To 3)
'            ReDim beta(1 To 3)
'            For rad = 1 To 3
'                xpxi(rad, 1) = wsArv_ip.Cells(60 + rad, 23).Value
'                xpxi(rad, 2) = wsArv_ip.Cells(60 + rad, 24).Value
'                xpxi(rad, 3) = wsArv_ip.Cells(60 + rad, 25).Value
'            Next
'            beta(1) = xpxi(1, 1) * xy(1) + xpxi(2, 1) * xy(2) + xpxi(3, 1) * xy(3)
'            beta(2) = xpxi(1, 2) * xy(1) + xpxi(2, 2) * xy(2) + xpxi(3, 2) * xy(3)
'            beta(3) = xpxi(1, 3) * xy(1) + xpxi(2, 3) * xy(2) + xpxi(3, 3) * xy(3)
'            avvxx = beta(1) + beta(2) * age + beta(3) * (age * age)
'        End If
'        'xpx = worksheet.function.MMULT(Application.WorksheetFunction.Transpose(Range1);Rnage1) 'Skulle vara mer generellt
'        'Regression: B=(X'X)^-1 (X'Y) -> 2x1 matris om enkel linjär ...
'  Else
'        t = arvslut
'        arvxx = yarv(t)
'         If verbose = 1 Then
'            Debug.Print "Senaste arvsvinst år för ålder "; age; t & "är "; arvxx; ""
'        End If
'  End If
'
'End Function
'
'
'Sub arvprognos()
'    Dim age As Single
'    Dim start, slut As Single
'    slut = 2020 'Sista fastställda år för arvsvinster
'    start = 2020 'Om samma som slut ingen utjämning/prediktion, annars OLS skattning för åren start-slut
'
'    Dim beta As Double 'Arvsvinstfaktor 'historisk vid slut eller (OLS) predikterade för slut
'
'    Dim d0, dt As Double    'Dödsrisken vid start och vid tidpunkten t
'    Dim rad As Long         'Radhållare
'    Dim t As Long           'Prognostidpunkt
'    Dim prognos() As Double 'Som matris ()  17 to 81, 2021 to 2070 mht riktålder
'    ReDim prognos(17 To 85, start + 1 To 2070)
'
'    Dim ral As Single       'Riktalder lägsta
'
'     For t = start + 1 To 2070
'        ral = riktage(t - 1, 0) - 1 'Faktor 1, året efter
'
'        For age = 17 To ral '17 to ral, senare ral till 81
'            If age < 61 Then
'                beta = arvxx(age, start, slut, 0, 0) '0=verbose, 0=DD
'              Else
'                beta = arvxx(age, start, slut, 0, 1) '0=verbose, 1=DD
'              End If
'
'            'Mortalitet vid slut
'            rad = (slut - 1980) * 106 * 2 + age + 2 'År *sex *ålder
'            d0 = wsMortality.Cells(rad, 4)
'            rad = rad + 106 '+ age
'            d0 = (d0 + wsMortality.Cells(rad, 4)) / 2
'
'            'Hämtar från mortalitet för t, kolla fliken att rätt dödsrisk fångas
'            rad = (t - 1980) * 106 * 2 + age + 2 'År *sex *ålder
'            dt = wsMortality.Cells(rad, 4)
'            rad = rad + 106 '+ age
'            dt = (d0 + wsMortality.Cells(rad, 4)) / 2 'oviktat
'            prognos(age, t) = Round(beta * ((1 - d0) / (1 - dt)), 6)
'
'            If age > 60 Then
'
'                'Dim rad As Integer
'                ReDim xy(1 To 3)
'                For rad = 17 To 60
'                    xy(1) = xy(1) + wsArv_ip.Cells(rad - 16 + 2, t - 2000 + 4)
'                    xy(2) = xy(2) + rad * wsArv_ip.Cells(rad - 16 + 2, t - 2000 + 4)
'                    xy(3) = xy(3) + (rad * rad) * wsArv_ip.Cells(rad - 16 + 2, t - 2000 + 4)
'                Next rad
'
'                ReDim xpxi(1 To 3, 1 To 3)
'                Dim mbeta() As Double
'                ReDim mbeta(1 To 3)
'                For rad = 1 To 3
'                    xpxi(rad, 1) = wsArv_ip.Cells(60 + rad, 23).Value
'                    xpxi(rad, 2) = wsArv_ip.Cells(60 + rad, 24).Value
'                    xpxi(rad, 3) = wsArv_ip.Cells(60 + rad, 25).Value
'                Next
'                mbeta(1) = xpxi(1, 1) * xy(1) + xpxi(2, 1) * xy(2) + xpxi(3, 1) * xy(3)
'                mbeta(2) = xpxi(1, 2) * xy(1) + xpxi(2, 2) * xy(2) + xpxi(3, 2) * xy(3)
'                mbeta(3) = xpxi(1, 3) * xy(1) + xpxi(2, 3) * xy(2) + xpxi(3, 3) * xy(3)
'
'                prognos(age, t) = Round(mbeta(1) + mbeta(2) * age + mbeta(3) * (age * age), 6)
'
'                'Debug.Print xy(1); xy(2); xy(3); mbeta(1); mbeta(2); mbeta(3); avvxx
'            End If
'
'
'            If prognos(age, t) < 1 Then prognos(age, t) = 1
'
'            rad = age - 17 + 3 'Output på rad, kolumnen t-2020+4
'            'rad = 81 + age - 60
'            If t < 2022 And age = 60 Then _
'             Debug.Print t; " 1 "; age; beta; Round((1 - d0) / (1 - dt), 4); prognos(age, t); "Rad:" & rad & "kol " & t - 2000 + 4
'            wsArv_ip.Cells(rad, t - 2000 + 4) = prognos(age, t)
'            'Viss korr manuellt när nya faktor dyker upp ...
'        Next age
'        'Lägg ut prognosmatrisen
'
'         ral = riktage(t - 1, 0) - 1 '2021-1 =61, 2022= 61.
'         For age = ral To 82
'             beta = arvxx(age, 2020, 2020, 0, 1)
'             'If age = 61 Then Stop
'             'If riktage(t, 0) - 1 = riktage(t - 1, 0) Then beta = arvxx(age + 1, 2020, 2020, 0, 1)
'
'            'Nedan som en hjälpfunktion
'            rad = CLng((slut - 1980) * 106 * 2) + age + 2 'År *sex *ålder
'            d0 = wsMortality.Cells(rad, 4)
'            rad = rad + 106 '+ age
'            d0 = (d0 + wsMortality.Cells(rad, 4)) / 2
'
'            'Hämtar från mortalitet, kolla fliken att rätt dödsrisk fångas
'            rad = (t - 1980) * 106 * 2 + age + 2 'År *sex *ålder
'            dt = wsMortality.Cells(rad, 4)
'            rad = rad + 106 '+ age
'            dt = (d0 + wsMortality.Cells(rad, 4)) / 2 'oviktat
'            prognos(age, t) = Round(beta * ((1 - d0) / (1 - dt)), 6)
'            If prognos(age, t) < 1 Then prognos(age, t) = 1
'
'            If riktage(t, 0) - 1 = riktage(t - 1, 0) And age = ral Then beta = prognos(age, t) = 0
'            rad = 81 + age - 60
'            If t < 2022 Then Debug.Print t; " 2 "; age; beta; Round((1 - d0) / (1 - dt), 4); prognos(age, t); "Rad:" & rad & "kol " & t - 2000 + 4
'            wsArv_ip.Cells(rad, t - 2000 + 4) = prognos(age, t)
'         Next age
'           'Lägg ut prognosmatrisen
'
'    Next t
'
'    'Skriv ut prognos
'
'End Sub
