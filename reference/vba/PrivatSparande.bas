Attribute VB_Name = "PrivatSparande"
Function PrivatSpar(ByVal IngåendeB, ByVal sparandePY, ByVal yield, ByVal ar, ByVal Typ) As Double
    
'yield      - Avkastning
'sparandePY - Sparande per år
'IngåendeB  - Ingående balans
'Typ        - Typ av sparande
        
    Select Case Typ
        Case 1 'KF
            'KF:s skatteunderlag beräknas som värdet vid 1 januari + alla insättningar som sker första halvåret + halva insättningar som sker andra halvåret
            AvkSkaUl = IngåendeB + sparandePY * (6 / 12) + sparandePY * (6 / 12) * 0.5
            'Årets avkastning - årets skatt blir nästa års ingående balans
            IngåendeB = IngåendeB * yield + sparandePY * yield ^ (180 / 360) - AvkastningsskattKFISK(AvkSkaUl, ar)
        Case 2 'ISK
            If ar > 2011 Then 'Innan 2012 fanns inte ISK
            'ISK:s skatteunderlag beräknas som värdet vid 1 jan, 1 apr, 1 juli och 1 okt + summan av insättningar under året delat med 4
            AvkSkaUl = (IngåendeB _
                + IngåendeB * yield ^ (90 / 360) + sparandePY * (3 / 12) * yield ^ (45 / 360) _
                + IngåendeB * yield ^ (180 / 360) + sparandePY * (6 / 12) * yield ^ (90 / 360) _
                + IngåendeB * yield ^ (270 / 360) + sparandePY * (9 / 12) * yield ^ (135 / 360) _
                + sparandePY) / 4
            'Årets avkastning - årets skatt blir nästa års ingående balans
            IngåendeB = IngåendeB * yield + sparandePY * yield ^ (180 / 360) - AvkastningsskattKFISK(AvkSkaUl, ar)
            End If
    End Select
    
    PrivatSpar = IngåendeB

End Function
Function AvkastningsskattKFISK(IngåendeB, ar)
'Funktion som beräknar avkastningsskatten för kapitalförsäkring (KF) och investeringssparkonto (ISK).
'Före 2012 fanns inte ISK, och avkastningsskatten var då något annorlunda.

'IngåendeB  - Ingående balans på kapitalförsäkringen/investeringsparkontot
'ar         - Beskattningsår

    If ar < 2012 Then ' Innan 2012 ingen ISK, KF hade skatt på 27% av schablon
    avkskatten = IngåendeB * 0.27 * Schablonintakt(ar)
    ElseIf ar < 2025 Then ' Innan 2025 inget grundavdrag
    avkskatten = IngåendeB * 0.3 * Schablonintakt(ar)
    ElseIf ar = 2025 Then ' 2025 införs grundavdrag om 150 000
        If IngåendeB > 150000 Then
            underlag = IngåendeB - 150000
            avkskatten = underlag * 0.3 * Schablonintakt(ar)
        ElseIf IngåendeB < 150000 Then
            avkskatten = 0
        End If
    ElseIf ar > 2025 Then ' 2026 införs grundavdrag om 300 000
        If IngåendeB > 300000 Then
            underlag = IngåendeB - 300000
            avkskatten = underlag * 0.3 * Schablonintakt(ar)
        ElseIf IngåendeB < 300000 Then
            avkskatten = 0
        End If
    End If

    AvkastningsskattKFISK = avkskatten
    
End Function
Function Schablonintakt(ar)
'Funktion som beräknar schablonintäkt för kapitalförsäkring (KF) och investeringssparkonto (ISK).
'Före 2012 fanns inte ISK, och avkastningsskatten var då något annorlunda. Därför skiljer sig statslåneräntan
'från 2012 och framåt och även viss logik skiljer sig.

'slr    - Statslåneräntan, innan 2012 så är det årsgenomsnittet, efter 2012 så är det statslåneräntan per den sista november året innan.
'ar     - Inkomståret
'ar2    - Året innan inkomståret

'Tar fram året innan inkomståret
ar2 = ar - 1

'Genomsnittliga statslåneräntan
Select Case ar2
    Case 1986
        Slr = 0.1077
    Case 1987
        Slr = 0.1167
    Case 1988
        Slr = 0.1135
    Case 1989
        Slr = 0.1118
    Case 1990
        Slr = 0.1313
    Case 1991
        Slr = 0.1072
    Case 1992
        Slr = 0.1003
    Case 1993
        Slr = 0.0855
    Case 1994
        Slr = 0.0957
    Case 1995
        Slr = 0.1014
    Case 1996
        Slr = 0.0789
    Case 1997
        Slr = 0.0647
    Case 1998
        Slr = 0.0498
    Case 1999
        Slr = 0.0489
    Case 2000
        Slr = 0.0534
    Case 2001
        Slr = 0.0498
    Case 2002
        Slr = 0.0515
    Case 2003
        Slr = 0.0439
    Case 2004
        Slr = 0.043
    Case 2005
        Slr = 0.0324
    Case 2006
        Slr = 0.0362
    Case 2007
        Slr = 0.0414
    Case 2008
        Slr = 0.0387
    Case 2009
        Slr = 0.0311
    Case 2010
        Slr = 0.0277
    Case 2011 'Här börjar slr vara per den sista november
        Slr = 0.0165
    Case 2012
        Slr = 0.0149
    Case 2013
        Slr = 0.0209
    Case 2014
        Slr = 0.009
    Case 2015
        Slr = 0.0065
    Case 2016
        Slr = 0.0027
    Case 2017
        Slr = 0.0049
    Case 2018
        Slr = 0.0051
    Case 2019
        Slr = -0.0009
    Case 2020
        Slr = -0.001
    Case 2021
        Slr = 0.0023
    Case 2022
        Slr = 0.0194
    Case 2023
        Slr = 0.0262
    Case 2024
        Slr = 0.0196
    Case 2025
        Slr = 0.0255
End Select

'Prognos slr
If ar > 2025 Then
    Slr = 0.025
End If

'Sedan 2016, ökning med en faktor om 0,75%. Minsta faktor är 1,25%.
If ar > 2015 And ar < 2018 Then
    Slr = Slr + 0.0075
    If Slr < 0.0125 Then
        Slr = 0.0125
    End If
'Ändrades 2018, ökning med en faktor om 1%. Minsta faktor är 1,25%.
ElseIf ar > 2017 Then
    Slr = Slr + 0.01
    If Slr < 0.0125 Then
        Slr = 0.0125
    End If
End If

    Schablonintakt = Slr

End Function


