Attribute VB_Name = "Tjänstepensioner"
Option Explicit

'Tjänstepensioner

Function tlITP1(ByVal alder, ByVal inkomst, ByVal IBB, ByVal Zpar, Optional ByVal marginal = 0) As Double
'Industrins och handelns tilläggspension för tjänstemän födda 1979 och senare. ITP 1
'Alder      - Ålder 31/12 akuellt inkomstår
'Inkomst    - Inkomst
'IBB        - Inkomstbasbelopp
'ZPAR       - Pensionsålder (Tjp_par) för tjänstepensionen
'Marginal   - Med avrundningar

     'Application.Volatile
    tlITP1 = 0
    
    Dim month As Integer '#antal månader med arbete och premieinbetalningar
    
    If Zpar > 65 Then Zpar = 65 'Ingen pensionsrätt efter 65
    If year_(age) > 2022 Then Zpar = 66
    
    'Antas uppbära lönen jämnt över året, obs arbetslöshet.. mm sprider sig jämnt och inte vid något tillfälle
    If inkomst <= 0 Then Exit Function
    
    If alder < 25 Then
       Exit Function
       'month = 0
    ElseIf Int(alder) = 25 Then
       month = Int((Int(born + 25 + 1) - born - W_start) * 12) 'Antal månder sedan 25 års ålder
       If month > 12 Then month = 12
       If month < 0 Then month = 0
       
    ElseIf alder < Int(Zpar) Then  '26-64 år
        If alder > W_start Then
            month = 12
        Else
            month = Int((Int(born + W_start + 1) - born - W_start) * 12) 'Antal arbetande månder under året
        End If
        
    ElseIf alder = Int(Zpar) Then
        If Zpar < 65 Then
            month = Int((born + Zpar - Int(born + Zpar)) * 12)
        Else
            month = Int((born - Int(born)) * 12)
            If Int((born + tjp_par - Int(born + tjp_par))) > 0 Then _
            inkomst = inkomst * month / (Int((born + tjp_par - Int(born + tjp_par)) * 12))
            'Problem om lönen avser hela året medan han bara är 64 år under en del av året - Får lösas att inkomsten avser de månader
            'Eller Income=Income*month_/12 'då month: inte används
        End If
    Else
       month = 0
    End If
    If IBB = 0 Then IBB = inkomst
    If month = 0 Then Exit Function
    ' Räknar såldes först månadsvis
    inkomst = inkomst / month
    Select Case inkomst
        Case Is <= 7.5 * IBB / month
            tlITP1 = 0.045 * inkomst
            If year_(age) > 2013 Then
                tlITP1 = (0.045 + Application.Range("rng_FlexPens").Value) * inkomst
''            ElseIf year_(age) = 2024 Then 'Flexpension KOLLAS
''                tlITP1 = (0.045 + 0.006 * 4 / 12) * inkomst
''            ElseIf year_(age) = 2025 Then
''                tlITP1 = (0.051 + 0.004 * 4 / 12) * inkomst
''            ElseIf year_(age) = 2026 Then
''                tlITP1 = (0.055 + 0.003 * 4 / 12) * inkomst
''            ElseIf year_(age) = 2027 Then
''                tlITP1 = (0.058 + 0.003 * 4 / 12) * inkomst
''            ElseIf year_(age) = 2027 Then
''                tlITP1 = (0.061 + 0.002 * 4 / 12) * inkomst
''            ElseIf year_(age) = 2028 Then
''                tlITP1 = (0.063 + 0.002 * 4 / 12) * inkomst
''              Else
''                tlITP1 = (0.065) * inkomst
            End If
        
        Case Else
            
            If year_(age) > 2022 And inkomst > 30 * IBB Then inkomst = 30 * IBB
            
            tlITP1 = 0.3 * (inkomst - 7.5 * IBB / 12) + 0.045 * 7.5 * IBB / 12
            If year_(age) > 2013 Then
                tlITP1 = (0.3 + Application.Range("rng_FlexPens").Value) * (inkomst - 7.5 * IBB / 12) + (0.045 + Application.Range("rng_FlexPens").Value) * 7.5 * IBB / 12
''            ElseIf year_(age) = 2024 Then 'Flexpension
''                tlITP1 = (0.3 + 0#) * (inkomst - 7.5 * IBB / 12) + (0.045 + 0.006 * 4 / 12) * 7.5 * IBB / 12
''            ElseIf year_(age) = 2025 Then
''                tlITP1 = (0.3 + 0#) * (inkomst - 7.5 * IBB / 12) + (0.051 + 0.004 * 4 / 12) * 7.5 * IBB / 12
''            ElseIf year_(age) = 2026 Then
''                 tlITP1 = (0.3 + 0#) * (inkomst - 7.5 * IBB / 12) + (0.055 + 0.005 * 4 / 12) * 7.5 * IBB / 12
''            ElseIf year_(age) = 2027 Then
''                 tlITP1 = (0.3 + 0#) * (inkomst - 7.5 * IBB / 12) + (0.058 + 0.003 * 4 / 12) * 7.5 * IBB / 12
''            ElseIf year_(age) = 2027 Then
''                 tlITP1 = (0.3 + 0#) * (inkomst - 7.5 * IBB / 12) + (0.061 + 0.002 * 4 / 12) * 7.5 * IBB / 12
''            ElseIf year_(age) = 2028 Then
''                tlITP1 = (0.3 + 0#) * (inkomst - 7.5 * IBB / 12) + (0.063 + 0.002 * 4 / 12) * 7.5 * IBB / 12
''            Else
''                tlITP1 = (0.3 + 0#) * (inkomst - 7.5 * IBB / 12) + (0.065) * 7.5 * IBB / 12
            End If
       
     End Select
     'Och sedan årets rätt
     
    If marginal = 0 Then tlITP1 = Int(tlITP1 + 0.5) 'Premien som arbetsgivaren betalar för den anställde rill förvaltaren avrundas...
    
    tlITP1 = tlITP1 * month

End Function

''Sub koll_ITP1()
''    Call startsetup
''
''    Dim itp1 As Double
''    'tlITP1(ByVal alder, ByVal inkomst, ByVal ibb, ByVal par, Optional ByVal marginal = 0)
''    Dim age As Integer
''    For age = 64 To 66
''        itp1 = tlITP1(age, 100000, 56600, tjp_par, 0)
''        Debug.Print age; itp1
''    Next age
''End Sub


Function tlITP2A(ByVal alder, ByVal inkomst, ByVal Zpar, Optional ByVal marginal = 0) As Double
    'ITP K - Industrins och handelns tilläggspension för tjänstemän. Del av ITP 2 avtalet
    'Alder      - Ålder 31/12 akuellt inkomstår
    'Inkomst    - Inkomst
    'ZPAR        - Pensionsålder (Tjp_par) för tjänstepensionen
    'Marginal   - Med avrundningar
    
     'Application.Volatile
    tlITP2A = 0
    
    'Kunde inte tjäna in ITP K före 1997
    If year_(alder) < 1997 Then Exit Function
    If inkomst <= 0 Then Exit Function
    
    Dim month As Integer '#antal månader med arbete och premieinbetalningar
    
    If Zpar > 65 Then Zpar = 65 'Inga nya pensionsrätter efter 65
    
    '#månader under inkomståret
    If alder < 28 Then
        Exit Function
        'month = 0
    ElseIf Int(alder) = 28 Then
        month = Int((Int(born + 28 + 1) - born - W_start) * 12) 'Antal månder sedan 28 års ålder
        If month > 12 Then month = 12
    ElseIf alder < Int(Zpar) Then  '26-64 år
            If alder > W_start Then
                month = 12
            Else
                month = Int((Int(born + W_start + 1) - born - W_start) * 12) 'Antal arbetande månder under året
            End If
            
    ElseIf alder = Int(Zpar) Then
            If Zpar < 65 Then
                month = Int((born + Zpar - Int(born + Zpar)) * 12)
            Else
                month = Int((born - Int(born)) * 12)
                If Int((born + tjp_par - Int(born + tjp_par)) * 12) > 0 Then _
                inkomst = inkomst * month / (Int((born + tjp_par - Int(born + tjp_par)) * 12))
                'Problem om lönen avser hela året medan han bara är 64 år under en del av året - Får lösas att inkomsten avser de månader
                'Eller Income=Income*month_/12 'då month: inte används
            End If
    Else
           month = 0
    End If
    
    If month = 0 Then Exit Function
            
    tlITP2A = 0.02 * inkomst * month / 12 'Hela inkomsten?
    If marginal = 0 Then tlITP2A = Int(tlITP2A + 0.5) 'Premien som arbetsgiaren betala rill förvaltaren avrundas...
    
    
End Function


'Sub koll_ITP2()
'    Call startsetup
'
'    Dim itp2 As Double
'    'tlITP1(ByVal alder, ByVal inkomst, ByVal ibb, ByVal par, Optional ByVal marginal = 0)
'    'tlITP2A(ByVal alder, ByVal inkomst, ByVal Zpar, Optional ByVal marginal = 0)
'    Dim age As Integer
'    For age = 64 To 66
'        itp2 = tlITP2A(age, 100000, TJP_par, 0)
'        Debug.Print age; itp2
'    Next age
'End Sub

Function DC_underlag(ByVal Stat As Byte) As Double
    'Kollar löneunderlaget efter lönerevideringar i den förmånsbaserade delen till ITP 2. 'Reviderad 20190624 TL
    'Sk lönekapningar, W1 = Lön(t) och w2=Lön(t-1) samt motsvarande index för IBB1 och IBB2 som avser inkomstbasbeloppet
     'Application.Volatile
     
     Dim Zpar As Double
     Zpar = tjp_par 'Worksheets("Start").Range("tjp_Par").Value 'Pensionsåldern
     If Zpar > 65 Then Zpar = 65
     Dim W1 As Double 'Lönen År t-1
     Dim W2 As Double 'Lönen År t-2
     Dim W3 As Double
     Dim W4 As Double
     Dim W5 As Double
     Dim W6 As Double
     
     'Motsvarande pris- och inkomstbasbelopp
     Dim pbb1 As Double
     Dim pbb2 As Double
     Dim pbb3 As Double
     Dim pbb4 As Double
     Dim pbb5 As Double
      
     Dim ibb1 As Double
     Dim ibb2 As Double
     Dim ibb3 As Double
     Dim ibb4 As Double
     Dim ibb5 As Double
     Dim ibb6 As Double
     
     W1 = Wage_(Int(Zpar) - 1)
     W2 = Wage_(Int(Zpar) - 2)
     W3 = Wage_(Int(Zpar) - 3)
     W4 = Wage_(Int(Zpar) - 4)
     W5 = Wage_(Int(Zpar) - 5)
     W6 = Wage_(Int(Zpar) - 6)
     
     pbb1 = pbb(Int(Zpar) - 1)
     pbb2 = pbb(Int(Zpar) - 2)
     pbb3 = pbb(Int(Zpar) - 3)
     pbb4 = pbb(Int(Zpar) - 4)
     pbb5 = pbb(Int(Zpar) - 5)
     
     ibb1 = IBB(Int(Zpar) - 1)
     ibb2 = IBB(Int(Zpar) - 2)
     ibb3 = IBB(Int(Zpar) - 3)
     ibb4 = IBB(Int(Zpar) - 4)
     ibb5 = IBB(Int(Zpar) - 5)
     ibb6 = IBB(Int(Zpar) - 6)
    
    'löneökningarna - s.k. lönekapningar
    If W6 = 0 Or W5 = 0 Or W4 = 0 Or W3 = 0 Or W2 = 0 Or W1 = 0 Then
        DC_underlag = 0
        'Vad händer om lön saknas ngt år?
    Else
    'https://collectum.se/foretag/rapportering/lonekapning-i-itp-2 och lönekapning som inte råder i staten
        If Stat = 0 Then
            If (W5 / W6) > (1.2 * ibb5 / ibb6) Then W5 = W6 * (1.2 * ibb6 / ibb5)
            If (W4 / W5) > (1.15 * ibb4 / ibb5) Then W4 = W5 * (1.15 * ibb4 / ibb5)
            If (W3 / W4) > (1.1 * ibb3 / ibb4) Then W3 = W4 * (1.1 * ibb3 / ibb4)
            If (W2 / W3) > (1.05 * ibb2 / ibb3) Then W2 = W3 * (1.05 * ibb2 / ibb3)
            If (W1 / W2) > (ibb1 / ibb2) Then W1 = W2 * (ibb1 / ibb2)
        End If
        'och därefter räkna om lönen i fasta priser året före (t-1)
        W5 = W5 * pbb1 / pbb5
        W4 = W4 * pbb1 / pbb4
        W3 = W3 * pbb1 / pbb3
        W2 = W2 * pbb1 / pbb2
             
        DC_underlag = (W1 + W2 + W3 + W4 + W5) / 5
        If marginal = 0 Then DC_underlag = Int(DC_underlag + 0.5) 'Avrundning?
    End If
     
End Function

''Sub underl()
''    Call startsetup
''    Dim koll As Double
''    koll = DC_underlag()
''    Debug.Print koll
''End Sub


Function tlITP2F(ByVal inkomst As Double, ByVal IBB As Double, Optional ByVal year = 30, Optional marginal = 0) As Double
 'Rent förmånsbestämnt
 'Inkomst är pensionsgrundande inkomsten
 'IBB        - Inkomstbasbelopp
 'Year       - Antal inkomstår
 'Marginal   - Med avrundningar
    tlITP2F = 0
    If inkomst <= 0 Then Exit Function
        'Application.Volatile
    If inkomst <= 7.5 * IBB Then
        tlITP2F = 0.1 * inkomst
    ElseIf inkomst <= 20 * IBB Then
        tlITP2F = 0.65 * (inkomst - 7.5 * IBB) + 0.1 * 7.5 * IBB
    ElseIf inkomst <= 30 * IBB Then
        tlITP2F = 0.325 * (inkomst - 20 * IBB) + 0.65 * (10 - 7.5) * IBB + 0.1 * 7.5 * IBB
    Else
        tlITP2F = 0.325 * (30 - 20) * IBB + 0.65 * (10 - 7.5) * IBB + 0.1 * 7.5 * IBB
    End If
    
    'Antal anställda år>=30 annars proportionerligt nedskalning (igentligen 360 månader
    If year < 30 Then tlITP2F = tlITP2F * year / 30
    'Den förmånsbaserade pensionen antas tas ut helt ... vid andra uttaget om första sker före 65
    Dim faktor As Double
    Dim Zpar As Double
   ' Dim Def_ar As Double
    
    Zpar = tjp_par 'Worksheets("Start").Range("tjp_Par").Value
    'def_ar = Worksheets("Start").Range("def_ar").Value
    'If par < 65 And def_ar > par Then par = def_ar
    Dim month As Long
    
    If Int(Zpar) < 64 Then
      month = Int((65 - Zpar) * 12)
      faktor = 1 - 0.005 * month
    ElseIf Int(Zpar) = 64 Then
      month = Int((65 - Zpar) * 12)
      faktor = 1 - 0.005 * month
    ElseIf Int(Zpar) <= 70 Then
      month = Int((Zpar - 65) * 12)
      faktor = 1 + 0.006 * month
    Else 'Efter 70
      faktor = (1 + 0.006 * 60) '5*12
    End If
    
    tlITP2F = tlITP2F * faktor
    If marginal = 0 Then tlITP2F = Int(tlITP2F / 12 + 0.5) * 12 'Avrundas till hela kronor?
End Function

'Sub koll_form()
'    Call startsetup
'    'Function tlITP2F(ByVal inkomst As Double, ByVal ibb As Double, Optional ByVal year = 30, Optional marginal = 0) As Double
'    Dim koll As Double
'    koll = tlITP2F(360000, 56900, 30)
'    Debug.Print koll
'End Sub


Function Kapan(year, fodar, ByVal IBB, ByVal wage, Optional ByVal Zpar = 65) As Double
'Year   - Inkomstår
'Fodar  - Födelseår (Born)
'IBB    - Inkomstbasbelopp
'wage   - Inkomst
'ZPAR    - Pensionering tjp
    
    'OBS hanterar även oligatorisk och flex i PA 16 avtalet
    'Application.Volatile
    Dim alder As Double 'Ålder 31/12
    Dim prem As Double 'Avgiften
    Dim month As Integer 'Antal månader
    
    Kapan = 0
    If wage <= 0 Then Exit Function
    alder = year - Int(fodar) 'Ålder 31/12
    If alder > Int(Zpar) Then Exit Function
    
''    Dim pyear As Double      'Pensionsår = Födelseår + pensionsålder
''    Dim month As Double      'När under året sker första pensioneringen, dvs antal månder som pensionär under första året som pensionär
''    pyear = born + par
''    month = 12 - Int(12 * (pyear - Int(pyear + 1 / 1000))) '# Pensions månader vid första året, 1959,8+65,2 -> 1/1 2025, Dvs 12 månader
''
    
    'Premien varierar över tid
    If year < 1991 Then
        prem = 0
    ElseIf year <= 1993 Then
        prem = 0.013
    ElseIf year = 1994 Then
        prem = 0.015
    ElseIf year <= 2002 Then
        prem = 0.019
    Else
        prem = 0.02
    End If
    
  'Från och med 2003 begr. till 30 ibb
  If (year > 2002 And wage > 30 * IBB) Then wage = 30 * IBB
  
  'Intjänandeålder kåpan är 28 år från år 1991 till år 2007. Först år 2008 blir den 23.
  Dim LOW_age As Integer
  Dim high_age As Integer
  
    If year <= 2007 Then
        LOW_age = 28
    Else
        LOW_age = 23
    End If
    
    high_age = 65 'Nya p-rätter
    
    If year > 2023 Then high_age = 69
    If Zpar > high_age Then Zpar = high_age '31/12 tidigare
  
    If alder < LOW_age Then
        Exit Function
        'month=0
    ElseIf alder = LOW_age Then
        month = Int((Int(born + LOW_age + 1) - born - W_start) * 12)
        If month > 12 Then month = 12
        Kapan = wage * prem * month / 12
    ElseIf alder < Int(Zpar) Then
        If alder > W_start Then
            month = 12
        Else
            month = Int((Int(born + W_start + 1) - born - W_start) * 12) 'Antal arbetande månder under året
        End If
        Kapan = wage * prem * month / 12
    ElseIf alder = Int(Zpar) Then
    
       If Zpar < high_age Then
            month = Int((born + Zpar - Int(born + Zpar)) * 12)
            Kapan = wage * prem * month / 12
        Else
            month = Int((born - Int(born)) * 12)
            If Int((born + tjp_par - Int(born + tjp_par)) * 12) Then _
            Kapan = wage * prem * month / (Int((born + tjp_par - Int(born + tjp_par)) * 12))
        End If
    Else
        Kapan = 0
    End If
  
    'Ingen avrundning
End Function

'Sub koll_stat()
'    Call startsetup
'    'Kapan(year, fodar, ByVal ibb, ByVal wage, Optional ByVal Zpar = 67)
'    Dim Stat As Double
'    Dim age As Integer
'    For age = 20 To 30
'        Stat = Kapan(year_(age), born, ibb(age), 360000, TJP_par)
'        Debug.Print age; Stat
'    Next age
'End Sub

Function PA_indiv(year, fodar, ByVal IBB, ByVal wage, Optional ByVal Zpar = 67) As Double

'Year   - Inkomstår
'Fodar  - Födelseår (Born)
'IBB    - Inkomstbasbelopp
'wage   - Inkomst
'ZPAR    - Pensionering tjp
 
    'Application.Volatile
    
    PA_indiv = 0
    If wage <= 0 Then Exit Function
    Dim alder As Double 'Ålder 31/12
    Dim prem As Double 'Avgiften
    Dim month As Long 'Antal månader
     'OBS VALBAR i PA 16 avtalet
      
    alder = year - Int(fodar)
    If alder > Zpar Then Exit Function
  
    'Premien varierar över tid
    If year < 2003 Then
       prem = 0
    ElseIf year < 2008 Then
       prem = 0.023
    Else
       prem = 0.025
    End If

    If year > 2023 And year < 2026 Then 'Flexpension
        If Int(fodar) < 1965 Then
            prem = 0.03
        Else
            prem = 0.04
        End If
    ElseIf year = 2026 Then
        If Int(fodar) < 1965 Then
            prem = (0.031 * 9 + 0.032 * 3) / 12 'flex 0.016 fram till oktober sedan 0.017
        Else
            prem = (0.041 * 9 + 0.042 * 3) / 12 'flex 0.016 fram till oktober sedan 0.017
        End If
    ElseIf year > 2026 Then
        If Int(fodar) < 1965 Then
            prem = 0.032
        Else
            prem = 0.042
        End If
    End If
    'Från och med 2003 begr. till 30 ibb
    If (year > 2002 And wage > 30 * IBB) Then wage = 30 * IBB
     
    Dim high_age As Integer
    high_age = 65 'Men inga premier
    
    If year > 2023 Then high_age = 69
    If Zpar > high_age Then Zpar = high_age
     
     
    If alder < 23 Then
        Exit Function
        'month=0
    ElseIf alder = 23 Then
        month = Int((Int(born + 23 + 1) - born - W_start) * 12)
        If month > 12 Then month = 12
        PA_indiv = wage * prem * month / 12
    ElseIf alder < Int(Zpar) Then
        If alder > W_start Then
            month = 12
        Else
            month = Int((Int(born + W_start + 1) - born - W_start) * 12) 'Antal arbetande månder under året
        End If
        PA_indiv = wage * prem * month / 12
        
         
    ElseIf alder = Int(Zpar) Then
        If Zpar < high_age Then
            month = Int((born + Zpar - Int(born + Zpar)) * 12)
            PA_indiv = wage * prem * month / 12
        Else
            month = Int((born - Int(born)) * 12)
            If Int((born + tjp_par - Int(born + tjp_par)) * 12) Then _
            PA_indiv = wage * prem * month / (Int((born + tjp_par - Int(born + tjp_par)) * 12))
        End If
    Else
        PA_indiv = 0
    End If
    'Ingen avrundning
End Function

''Sub koll_stat_ind()
''    Call startsetup
''    Dim Stat As Double
''    Dim age As Integer
''    For age = 23 To 25
''        Stat = PA_indiv(year_(age), born, ibb(age), 360000, TJP_par)
''        Debug.Print age; Stat
''    Next age
''End Sub


Function tlPA03(ByVal inkomst As Double, Optional ByVal aar = 30, Optional ByVal inkbas = 45900, Optional ByVal fodar = 1959) As Double
 'Statsanställdas förmånsdel i PA03
 'Inkomst   - pensionsunderlag
 'aar       - #år i staten
 'inkbas    - Inkomstbasbeloppet
 'Fodar     - Födelseår
 'Application.Volatile

'Parametrar för ersättningsnivå 1-3
 Dim par1 As Double
 Dim par2 As Double '
 Dim par3 As Double
  
 Dim benefit As Double
 benefit = 0
 If inkomst <= 0 Then Exit Function
 
 Dim VAL_ As Integer
 VAL_ = Application.Range("rng_TJP_Val").Value
 If VAL_ = 8 Then Exit Function
 
 If fodar < 1943 Then
    par1 = 0.1: par2 = 0.65: par3 = 0.325
 ElseIf fodar = 1943 Then
    par1 = 0.095: par2 = 0.6485: par3 = 0.324
 ElseIf fodar = 1943 Then
    par1 = 0.095: par2 = 0.6485: par3 = 0.324
 ElseIf fodar = 1944 Then
    par1 = 0.093: par2 = 0.647: par3 = 0.323
 ElseIf fodar = 1945 Then
    par1 = 0.091: par2 = 0.6455: par3 = 0.322
 ElseIf fodar = 1946 Then
    par1 = 0.089: par2 = 0.644: par3 = 0.321
 ElseIf fodar = 1947 Then
    par1 = 0.087: par2 = 0.6425: par3 = 0.32
 ElseIf fodar = 1948 Then
    par1 = 0.084: par2 = 0.641: par3 = 0.319
 ElseIf fodar = 1949 Then
    par1 = 0.082: par2 = 0.6395: par3 = 0.318
 ElseIf fodar = 1950 Then
    par1 = 0.079: par2 = 0.638: par3 = 0.317
 ElseIf fodar = 1951 Then
    par1 = 0.077: par2 = 0.6365: par3 = 0.316
 ElseIf fodar = 1952 Then
    par1 = 0.074: par2 = 0.635: par3 = 0.315
 ElseIf fodar = 1953 Then
    par1 = 0.072: par2 = 0.6335: par3 = 0.314
 ElseIf fodar = 1954 Then
    par1 = 0.069: par2 = 0.632: par3 = 0.313
 ElseIf fodar = 1955 Then
    par1 = 0.066: par2 = 0.6305: par3 = 0.312
 ElseIf fodar = 1956 Then
    par1 = 0.063: par2 = 0.629: par3 = 0.311
 ElseIf fodar = 1957 Then
    par1 = 0.06: par2 = 0.6275: par3 = 0.31
 ElseIf fodar = 1958 Then
    par1 = 0.057: par2 = 0.626: par3 = 0.309
 ElseIf fodar = 1959 Then
    par1 = 0.054: par2 = 0.6245: par3 = 0.308
 ElseIf fodar = 1960 Then
    par1 = 0.051: par2 = 0.623: par3 = 0.307
 ElseIf fodar = 1961 Then
    par1 = 0.047: par2 = 0.6215: par3 = 0.306
 ElseIf fodar = 1962 Then
    par1 = 0.043: par2 = 0.62: par3 = 0.305
 ElseIf fodar = 1963 Then
    par1 = 0.039: par2 = 0.6185: par3 = 0.304
 ElseIf fodar = 1964 Then
    par1 = 0.036: par2 = 0.617: par3 = 0.303
 ElseIf fodar = 1965 Then
    par1 = 0.032: par2 = 0.615: par3 = 0.302
 ElseIf fodar = 1966 Then
    par1 = 0.029: par2 = 0.613: par3 = 0.301
 ElseIf fodar = 1967 Then
    par1 = 0.025: par2 = 0.611: par3 = 0.3
 ElseIf fodar = 1968 Then
    par1 = 0.021: par2 = 0.609: par3 = 0.3
 ElseIf fodar = 1969 Then
    par1 = 0.017: par2 = 0.607: par3 = 0.3
 ElseIf fodar = 1970 Then
    par1 = 0.013: par2 = 0.605: par3 = 0.3
 ElseIf fodar = 1971 Then
    par1 = 0.009: par2 = 0.603: par3 = 0.3
 ElseIf fodar = 1972 Then
    par1 = 0.005: par2 = 0.601: par3 = 0.3
 ElseIf fodar >= 1973 Then
    par1 = 0: par2 = 0.6: par3 = 0.3
 End If
 
 
 If inkomst > 30 * inkbas Then
    benefit = (30 - 20) * inkbas * par3 + (20 - 7.5) * par2 * inkbas + 7.5 * inkbas * par1
 ElseIf inkomst > 20 * inkbas Then
    benefit = (inkomst - 20 * inkbas) * par3 + (20 - 7.5) * par2 * inkbas + 7.5 * inkbas * par1
 ElseIf inkomst > 7.5 * inkbas Then
    benefit = (inkomst - 7.5 * inkbas) * par2 + 7.5 * inkbas * par1
 Else
   benefit = inkomst * par1
 End If
 'Förtida och uppskjutet uttag, akturiellt... delningstal 65 /delningstal aktuell ålder?.
 
 If aar < 30 Then benefit = benefit * (aar / 30)
 tlPA03 = benefit
 'Ingen avrundning
End Function

''Sub koll_form()
''    'Call startsetup
''    Dim koll As Double
''    koll = tlPA03(600000, 37, 59600, 1959)
''    Debug.Print koll
''End Sub

Function tlPA16(ByVal alder, ByVal inkomst, ByVal IBB, ByVal Zpar, Optional ByVal marginal = 0) As Double
'Statsanställda avd 1 (födda 1988 och senare)
'Alder      - Ålder 31/12 akuellt inkomstår
'Inkomst    - Inkomst
'IBB        - Inkomstbasbelopp
'ZPAR        - Pensionsålder (Tjp_par) för tjänstepensionen
'Marginal   - Med avrundningar

     'Application.Volatile
    tlPA16 = 0
    If inkomst <= 0 Then Exit Function
    Dim month As Integer '#antal månader med arbete och premieinbetalningar
    Dim max_age As Integer
    
    If year_(alder) > 2003 Then
        max_age = 67
    Else
        max_age = 65
    End If
    If Zpar > max_age Then Zpar = max_age 'Ingen pensionsrätt efter LAS-åldern
    
    'Antas uppbära lönen jämnt över året, obs arbetslöshet.. mm sprider sig jämnt och inte vid något tillfälle
    If alder < W_start Then
       Exit Function
       'month = 0
    ElseIf Int(alder) = W_start Then
       month = Int((Int(born + W_start + 1) - born - W_start) * 12) 'Antal månder
       If month > 12 Then month = 12
    ElseIf alder < Int(Zpar) Then
            month = 12
    ElseIf alder = Int(Zpar) Then
        If Zpar < max_age Then
            month = Int((born + Zpar - Int(born + Zpar)) * 12)
        Else
            month = Int((born - Int(born)) * 12)
            If Int((born + tjp_par - Int(born + tjp_par))) > 0 Then _
            inkomst = inkomst * month / (Int((born + tjp_par - Int(born + tjp_par)) * 12))
            'Problem om lönen avser hela året medan han bara är 64 år under en del av året - Får lösas att inkomsten avser de månader
            'Eller Income=Income*month_/12 'då month: inte används
        End If
    Else
       month = 0
    End If
    
    If IBB = 0 Then IBB = inkomst
    If month = 0 Then Exit Function
    ' Räknar såldes först månadsvis
    inkomst = inkomst / month
    If year_(alder) < 2026 Then
    Select Case inkomst
        Case Is <= 7.5 * IBB / month
            tlPA16 = 0.06 * inkomst 'Valbar 0.025 Obligatorisk 0.02 och flex 0.015
        Case Else
            tlPA16 = 0.315 * (inkomst - 7.5 * IBB / 12) + 0.06 * 7.5 * IBB / 12 '0.20+0.10+0.015
     End Select
     ElseIf year_(alder) = 2026 Then
     Select Case inkomst
        Case Is <= 7.5 * IBB / month
            tlPA16 = (0.061 * 9 + 0.062 * 3) / 12 * inkomst 'Valbar 0.025 Obligatorisk 0.02 och flex 0.016 fram till oktober sedan 0.017
        Case Else
            tlPA16 = (0.316 * 9 + 0.317 * 3) / 12 * (inkomst - 7.5 * IBB / 12) + (0.061 * 9 + 0.062 * 3) / 12 * 7.5 * IBB / 12 '0.20+0.10+0.016 fram till oktober sedan 0.017
     End Select
     ElseIf year_(alder) > 2026 Then
     Select Case inkomst
        Case Is <= 7.5 * IBB / month
            tlPA16 = 0.062 * inkomst 'Valbar 0.025 Obligatorisk 0.02 och flex 0.017
        Case Else
            tlPA16 = 0.317 * (inkomst - 7.5 * IBB / 12) + 0.062 * 7.5 * IBB / 12 '0.20+0.10+0.017
     End Select
     End If
     'Och sedan årets rätt
    If marginal = 0 Then tlPA16 = Int(tlPA16 + 0.5) 'Premien som arbetsgivaren betalar för den anställde rill förvaltaren avrundas...
    
    tlPA16 = tlPA16 * month

End Function

''Sub koll_PA16()
''    Call startsetup
''    Dim Stat As Double
''    Dim age As Integer
''    'tlPA16(ByVal alder, ByVal inkomst, ByVal ibb, ByVal Zpar, Optional ByVal marginal = 0)
''    For age = 60 To 68
''        Stat = tlPA16(age, 360000, ibb(age), tjp_par, marginal)
''        Debug.Print age; Stat
''    Next age
''End Sub



Function SAF_LO(ByVal alder As Long, ByVal inkomst As Double, ByVal IBB As Long, ByVal Zpar As Double, year As Long, Optional marginal = 0) As Double
'Alder      - Ålder 31/12 akuellt inkomstår
'Inkomst    - Inkomst
'IBB        - Inkomstbasbelopp
'ZPAR       - Pensionsålder (Tjp_par) för tjänstepensionen
'Marginal   - Med avrundningar
'Privatanställda arbetare - kommer från STP som ersattes 1996 för födda 1968 eller senare enl nya planen
' Application.Volatile
    SAF_LO = 0
    If inkomst <= 0 Then Exit Function
    If year < 1996 Then Exit Function
    
    If Zpar > 65 Then Zpar = 65 'Ingen pensionsrätt efter 65
     
    Dim LOW_age As Long
    If year < 2000 Then
        LOW_age = 28  'Före 2000
''    ElseIf year < 2002 Then
''        LOW_age = 22
''    ElseIf year < 2008 Then
''        LOW_age = 21
''    Else
''        LOW_age = 25
    ElseIf year < 2021 Then
        LOW_age = 25
    ElseIf year < 2022 Then
        LOW_age = 24
     ElseIf year < 2023 Then
        LOW_age = 23
     Else
        LOW_age = 22
    End If
    
    If IBB = 0 Then IBB = inkomst
    
    ' Räknar såldes först månadsvis 2012 är premien 4.5 resp 30%
    Dim Premie1 As Double: Dim Premie2 As Double
    Select Case year
       Case Is < 1996
         Premie1 = 0: Premie2 = 0
       Case Is < 2000
         Premie1 = 0.02: Premie2 = 0.02
       Case Is < 2008
         Premie1 = 0.035: Premie2 = 0.035
       Case Is = 2008
         Premie1 = 0.039: Premie2 = 0.06
       Case Is = 2009
         Premie1 = 0.04: Premie2 = 0.12
       Case Is = 2010
         Premie1 = 0.041: Premie2 = 0.18
       Case Is = 2011
         Premie1 = 0.043: Premie2 = 0.24
       Case Is <= 2013
         Premie1 = 0.045: Premie2 = 0.3
       Case Is > 2013
         Premie1 = 0.045 + Application.Range("rng_FlexPens").Value: Premie2 = 0.3 + Application.Range("rng_FlexPens").Value
         'Premie1 = 0.05: Premie2 = 0.305 'inkl flexpension 0,5 % (infördes under 2013 och har gradvis räknats upp)
       'Case Is > 2016
         'Premie1 = 0.055: Premie2 = 0.31 'inkl flexpension 1,0 % (ligger mellan + 0,7 % och 1,7 %)
         
         'flexpensionen inom SAF-LO erhålls egentligen från första arbetsdagen oavsett ålder.
         
    End Select
     
    Dim month As Integer '#antal månader med arbete och premieinbetalningar
    If alder < LOW_age Then
       Exit Function
       'month = 0
    ElseIf Int(alder) = LOW_age Then
       month = Int((Int(born + LOW_age + 1) - born - W_start) * 12) 'Antal månder sedan 25 års ålder
       If month > 12 Then month = 12
       
    ElseIf alder < Int(Zpar) Then
        If alder > W_start Then
            month = 12
        Else
            month = Int((Int(born + W_start + 1) - born - W_start) * 12) 'Antal arbetande månder under året
        End If
        
    ElseIf alder = Int(Zpar) Then
        If Zpar < 65 Then
            month = Int((born + Zpar - Int(born + Zpar)) * 12)
        Else
            month = Int((born - Int(born)) * 12)
            If Int((born + tjp_par - Int(born + tjp_par))) > 0 Then _
            inkomst = inkomst * month / (Int((born + tjp_par - Int(born + tjp_par)) * 12))
            'Problem om lönen avser hela året medan han bara är 64 år under en del av året - Får lösas att inkomsten avser de månader
            'Eller Income=Income*month_/12 'då month: inte används
        End If
    Else
       month = 0
    End If
    
    If month = 0 Then Exit Function
    inkomst = inkomst / month
    Select Case inkomst
        Case Is <= 7.5 * IBB / month
            SAF_LO = Premie1 * inkomst
        Case Else
            SAF_LO = Premie2 * (inkomst - 7.5 * IBB / 12) + Premie1 * 7.5 * IBB / 12
     End Select
     
     If marginal = 0 Then SAF_LO = Int(SAF_LO + 0.5) 'Premien som arbetsgivaren betalar månadsvis till förvaltaren avrundas...
     'Och sedan årets premie
     SAF_LO = SAF_LO * month

End Function


''Sub koll_LO()
''    Call startsetup
''
''    Dim LO As Double
''    'SAF_LO(ByVal alder As Long, ByVal inkomst As Double, ByVal ibb As Long, ByVal Zpar As Double, year As Long, Optional marginal = 0) As Double
''    Dim age As Integer
''    For age = 64 To 65
''        LO = SAF_LO(age, 100000, 56600, tjp_par, 2009, 0)
''        Debug.Print age; LO
''    Next age
''End Sub

Function STP_(ByVal medel As Double, ByVal antal As Long, ByVal pbb As Long) As Double
'STP - för Privata arbetare successivt utfasas
'Födda 1932(5)-1967 enligt övergångsregler som fångas upp nedan.
'Medel - Medelpoäng
'Antal - Antal förvärvsår till och med 1995
'PBB   - Gällande Prisbasbelopp vid 65 eller år 1995 om det sker tidigare
    'Dim born As Double
    Dim PAR As Double
    'Dim def_ar as double
    STP_ = 0 'Ges senare Antal käver minst 832 timmar, =40% av heltid, ger ett STP år. Men minst 208 timmar för att få del av år. Timmar mellan 208 och 832 ger år = timmar/832
    If antal < 3 Then GoTo slut
    
     'born = Worksheets("Start").Range("Born").Value
     PAR = tjp_par 'Worksheets("Start").Range("tjp_Par").Value
    ' def_ar = Worksheets("Start").Range("Def_ar").Value
    ' If PAR > def_ar Then def_ar = PAR
    ' If def_ar < 61# Then def_ar = PAR
     'Antar att tjänstepensionen tas ut helt vid första uttaget dvs PAR
     If PAR < 65 Then PAR = 65 'Inget uttag före 65,
     If PAR > 70 Then PAR = 70 'Uppskjutas längst intill den månad varunder 70 års ålder uppnås
     born = Int(born)
 
Dim stp_ar As Long
Dim faktor As Double
 'stp_ar = 30
 
Dim month As Long 'Antal månader efter 65
 faktor = 1
 If PAR > 65 Then
    'If Int(Born + par) < 1995 Then faktor = 1
    If PAR < 65 Then
      faktor = 0 'Ingen pensionering före 65. Omöjlig händelse då modellen ger STP vid 65 även om "allmänna" pensioneringen sker tidigare
    ElseIf Int(PAR) <= 70 Then
       month = Int((PAR - 65) * 12)
       faktor = 1 + 0.006 * month
    Else 'Efter 70
      faktor = (1 + 0.006 * 60)
    End If
 End If
 
  If Int(born + PAR) > 1999 Then faktor = 1.025 * faktor '1.025 för framjustering, värdesäkring från 1996-2000 i övergångsreglerna

  Select Case Int(born)
   Case Is <= 1937
    stp_ar = 30
  Case Is <= 1940
    stp_ar = 32
  Case Is <= 1941
    stp_ar = 33
  Case Is <= 1942
    stp_ar = 34
  Case Is <= 1943
    stp_ar = 35
  Case Is <= 1944
    stp_ar = 36
  Case Else
    stp_ar = 37
 End Select
 
 If antal < stp_ar Then
   STP_ = faktor * medel * pbb * 0.1 * antal / stp_ar
 Else
   STP_ = faktor * medel * pbb * 0.1
 End If
 
 'Dim marginal As Single
 'marginal = Worksheets("Start").Range("marginal").Value
 If marginal = 0 Then STP_ = Int(STP_ / 12 + 0.49) * 12 'Vet inte avrundningsreglerna men antar närmaste krontalet per månad
 'Hel STP-pension ingen justering för partiellt uttag eller # månader
slut:
End Function

''Sub stp()
''    Dim stp As Double
''    stp = STP_(7.5, 20, 44300)
''    Debug.Print stp / 12
''End Sub

Function PA_KLBPP(ByVal Zpar As Double) As Double
'Funktion för beräkning av poäng för PA_KL. Inputvariablerna är som följer:
'Zpar   - Pensionsålder (Tjp_par) för tjänstepensionen
    PA_KLBPP = 0
    
    'Beräkning av årspoäng
    Dim W1 As Double 'lönen t-2
    Dim W2 As Double 'lönen t-3
    Dim W3 As Double
    Dim W4 As Double
    Dim W5 As Double
    Dim W6 As Double
    Dim W7 As Double
    
    'Motsvarande förhöjt prisbasbelopp
    Dim fpb1 As Double 'fbp t-2
    Dim fpb2 As Double 'fbp t-3
    Dim fpb3 As Double
    Dim fpb4 As Double
    Dim fpb5 As Double
    Dim fpb6 As Double
    Dim fpb7 As Double
    
    'Årspoäng
    Dim yearpointavg As Double
    Dim yearpoint1 As Double
    Dim yearpoint2 As Double
    Dim yearpoint3 As Double
    Dim yearpoint4 As Double
    Dim yearpoint5 As Double
    Dim yearpoint6 As Double
    Dim yearpoint7 As Double
    
    W1 = Wage_(Int(Zpar) - 2)
    W2 = Wage_(Int(Zpar) - 3)
    W3 = Wage_(Int(Zpar) - 4)
    W4 = Wage_(Int(Zpar) - 5)
    W5 = Wage_(Int(Zpar) - 6)
    W6 = Wage_(Int(Zpar) - 7)
    W7 = Wage_(Int(Zpar) - 8)
    
    fpb1 = FPB(Int(Zpar) - 2)
    fpb2 = FPB(Int(Zpar) - 3)
    fpb3 = FPB(Int(Zpar) - 4)
    fpb4 = FPB(Int(Zpar) - 5)
    fpb5 = FPB(Int(Zpar) - 6)
    fpb6 = FPB(Int(Zpar) - 7)
    fpb7 = FPB(Int(Zpar) - 8)
    
    yearpoint1 = W1 / fpb1
    yearpoint2 = W2 / fpb2
    yearpoint3 = W3 / fpb3
    yearpoint4 = W4 / fpb4
    yearpoint5 = W5 / fpb5
    yearpoint6 = W6 / fpb6
    yearpoint7 = W7 / fpb7
    
    Dim yparray(1 To 7) As Variant
    yparray(1) = yearpoint1
    yparray(2) = yearpoint2
    yparray(3) = yearpoint3
    yparray(4) = yearpoint4
    yparray(5) = yearpoint5
    yparray(6) = yearpoint6
    yparray(7) = yearpoint7
    
    'Medelvärde av fem bästa årspoängen om intjänande alla år
    If yearpoint7 <> 0 Then
    yearpointavg = WorksheetFunction.Average(WorksheetFunction.Large(yparray, 1), WorksheetFunction.Large(yparray, 2), WorksheetFunction.Large(yparray, 3), WorksheetFunction.Large(yparray, 4), WorksheetFunction.Large(yparray, 5))
    ElseIf yearpoint7 = 0 & yearpoint6 <> 0 Then
    yearpointavg = WorksheetFunction.Average(WorksheetFunction.Large(yparray, 1), WorksheetFunction.Large(yparray, 2), WorksheetFunction.Large(yparray, 3), WorksheetFunction.Large(yparray, 4))
    ElseIf yearpoint6 = 0 & yearpoint5 <> 0 Then
    yearpointavg = WorksheetFunction.Average(WorksheetFunction.Large(yparray, 1), WorksheetFunction.Large(yparray, 2), WorksheetFunction.Large(yparray, 3))
    ElseIf yearpoint5 = 0 & yearpoint4 <> 0 Then
    yearpointavg = WorksheetFunction.Average(WorksheetFunction.Large(yparray, 1), WorksheetFunction.Large(yparray, 2))
    ElseIf yearpoint4 = 0 & yearpoint3 <> 0 Then
    yearpointavg = WorksheetFunction.Average(WorksheetFunction.Large(yparray, 1), WorksheetFunction.Large(yparray, 2))
    ElseIf yearpoint3 = 0 & yearpoint2 <> 0 Then
    yearpointavg = WorksheetFunction.Average(WorksheetFunction.Large(yparray, 1), WorksheetFunction.Large(yparray, 2))
    Else
    yearpointavg = WorksheetFunction.Average(WorksheetFunction.Large(yparray, 1))
    End If
    
    'Bruttopensionspoäng
    Dim bpp As Double
    
    Select Case yearpointavg
        Case 0 To 1
            bpp = yearpointavg * 0.96
        Case 1 To 2.5
            bpp = 1 * 0.96 + (yearpointavg - 1) * 0.785
        Case 2.5 To 3.5
            bpp = 1 * 0.96 + 1.5 * 0.785 + (yearpointavg - 2.5) * 0.6
        Case 3.5 To 7.5
            bpp = 1 * 0.96 + 1.5 * 0.785 + 1 * 0.6 + (yearpointavg - 3.5) * 0.64
        Case 7.5 To 20
            bpp = 1 * 0.96 + 1.5 * 0.785 + 1 * 0.6 + 4 * 0.64 + (yearpointavg - 7.5) * 0.65
        Case 20 To 9999
            bpp = 1 * 0.96 + 1.5 * 0.785 + 1 * 0.6 + 4 * 0.64 + 12.5 * 0.65 + (yearpointavg - 20) * 0.325
    End Select
    
    PA_KLBPP = bpp
End Function
    
Function PA_KL(ByVal Zpar As Double, ByVal year As Long, ByVal bpp As Double, Optional marginal = 0) As Double
'PA-KL förmånsbestämds del för kommunala och regionala arbetare succesivt utfasas till KAP-KL
'Bruttosamordnad med ATP och FP, se TjänstepensionerFörmån. Inputvariablerna är som följer:
'Zpar       - Pensionsålder (Tjp_par) för tjänstepensionen
'year       - Inkomstår
'bpp        - PA-KL poäng, se funktion PAKLBPP
'marginal   - Med avrundningar

    PA_KL = 0
    
    'PFA98 börjar gälla, går mot tlkap_kl & KAPKL_f istället
    If year > 1997 Then
        Exit Function
    End If
        
    Dim tidfaktor As Double
    'Antal år i arbete (Om 1997 antal år i arbete till 1997)
    If year < 1997 Then
    tidfaktor = Zpar - W_start
    Else
    tidfaktor = 1997 - year_(W_start)
    End If
    If tidfaktor > 30 Then
        tidfaktor = 30
    End If
    
    'Beräkning av pension från PA-KL
    If year < 1997 Then
    PA_KL = (tidfaktor / 30) * bpp * FPB(Int(Zpar))
    Else
    PA_KL = (tidfaktor / 30) * bpp * pbb(Int(Zpar))
    End If
    
    'Avrundas upp till närmsta krona
    If marginal = 0 Then PA_KL = Int(PA_KL / 12 + 0.5) * 12
    
End Function

Function tlkap_kl(ByVal alder As Long, ByVal inkomst As Double, ByVal IBB As Long, ByVal Zpar As Double, _
    ByVal year As Long, Optional marginal = 0) As Double 'Bytt namn till - KR (landsting-Region)
'Kommunanställdas avgiftsbaserade del inom KAP-KL
'http://www.skl.se/avtal-lagar/pensioner_och_forsakringar
'Obs arbetsgivaren betalar avgiften senastden 31 mars året efter...
'Alder      - Ålder 31/12 akuellt inkomstår
'Inkomst    - Inkomst
'IBB        - Inkomstbasbelopp
'ZPAR       - Pensionsålder (Tjp_par) för tjänstepensionen
'Year       - Inkomstår
'Marginal   - Med avrundningar
'AKAP-kl från och med 2014-01-01 för födda 1986 och senare

    'Application.Volatile
    tlkap_kl = 0
    
''    Dim VAL_AKAP As Integer
''    VAL_AKAP = Application.Range("rng_TJP_Val").Value 'OBS läses från NameRange och D18
    'PA-KL gäller innan
    If year < 1998 Then
        Exit Function
    End If
    
    If alder < 21 Then 'And Int(born) < 1986
        Exit Function
        'month_ = 0
    End If
    Dim high_age As Integer
    '(A)KAP-KL injänar pensionsrätter fram till LAS ålder 67
    If year > 2002 Then 'And Int(born) > 1985 '2003 LAS åldern 67 år
        high_age = 67
    Else
        high_age = 65
    End If
    Dim month As Integer '#antal månader med arbete och premieinbetalningar
    
    If Zpar > high_age Then Zpar = high_age 'Ger pensionsrätt efter 65
    
    If born > 1985 And alder < 21 And year >= 2014 Then
        If alder < W_start Then   'W_start public när individen börjar jobba
            Exit Function
        ElseIf W_start = alder Then
           month = Int((Int(born + alder + 1) - born - W_start) * 12)
           If month > 12 Then month = 12
           If month < 0 Then month = 0
           inkomst = inkomst * month / 12
        Else
            month = 12
        End If
    ElseIf alder < 21 Then
        Exit Function
    ElseIf alder = 21 Then
        month = Int((Int(born + 21 + 1) - born - W_start) * 12)
        If month > 12 Then month = 12 'Wstart före 21 års ålder
        If month < 0 Then month = 0 'Efter 21
        inkomst = inkomst * month / 12
    ElseIf alder < Int(Zpar) Then
            If alder > W_start Then
                month = 12
            Else
                month = Int((Int(born + W_start + 1) - born - W_start) * 12) 'Antal arbetande månder under året
                inkomst = inkomst * month / 12
            End If
    ElseIf alder = Int(Zpar) Then
        If Zpar < high_age Then
                month = Int((born + Zpar - Int(born + Zpar)) * 12)
                inkomst = inkomst * month / 12
        Else
                month = Int((born - Int(born)) * 12)
                inkomst = inkomst * month / 12
    
        End If
    Else
            month = 0
    End If
    
    If IBB = 0 Then IBB = inkomst
    If month = 0 Then Exit Function
    
    Dim par1 As Double
    Dim par2 As Double
    par1 = 0
    par2 = 0
      
    'Sysselssättningsgrad bortses från
    If year < 1998 Then
      par1 = 0.035:   par2 = 0.011
      If alder < 28 Then
         month = 0: par1 = 0: par2 = 0
      End If
    ElseIf year <= 2002 Then
      If alder > 27 Then
       par1 = 0.034:  par2 = 0.01
      End If
    ElseIf year < 2005 Then
      par1 = 0.035: If alder > 27 Then par2 = 0.011
    ElseIf year = 2006 Or Int(born) <= 1946 Then
      par1 = 0.045
      par2 = 0.021
    ElseIf year = 2006 And Int(born) > 1946 Then
      par1 = 0.04: par2 = 0.04
    ElseIf year = 2007 Then
      par1 = 0.04: par2 = 0.04
    ElseIf year <= 2009 Then
      par1 = 0.0425: par2 = 0.0425
    ElseIf year > 2009 Then
      par1 = 0.045: par2 = 0.045
    End If
    
''    If Int(born) > 1985 And year >= 2014 Then
''      par1 = 0.045: par2 = 0.3
''    End If
''
''    If Int(born) > 1985 And year >= 2023 Then
''      par1 = 0.06: par2 = 0.315
''    End If
    
    
    
    'Fram till 2002 användes det förhöjda prisbasbeloppet därefter inkomstbasbeloppet (IBB). Nedan enbart IBB
    'Pensionsgrundande lön
    If inkomst > (30 * IBB) Then inkomst = 30 * IBB
    
    'If inkomst <= 0.01 * IBB Then inkomst = 0
    'Tillgodoräknas om inkomsten>1% av IBB, i annat fall betala arbetsgiv. motsvarande belopp till arbetstagaren i form av ersättning
    'som inte är pemnsionsgrundande, Arbetstagaren antas spara detta belopp "privat" till samma avkastning
    Select Case inkomst
        Case Is <= 7.5 * IBB 'limit
            tlkap_kl = par1 * inkomst
        Case Else
            tlkap_kl = par2 * (inkomst - 7.5 * IBB) + par1 * 7.5 * IBB
     End Select
     
    If marginal = 0 Then tlkap_kl = Int(tlkap_kl + 0.5) 'Premien som arbetsgivaren betala rill förvaltaren avrundas ?

End Function

''Sub koll_kommun()
''    For born = 1937 To 1939
''        Application.Range("bornYear") = born
''        Call startsetup
''        'tlkap_kl(ByVal alder As Long, ByVal inkomst As Double, ByVal ibb As Long, ByVal Zpar As Double, _
''        '    ByVal year As Long, Optional marginal = 0)
''        Dim Kom As Double
''        Dim age As Integer
''
''        For age = 23 To 66
''        If age = 27 Then Stop
''            Kom = tlkap_kl(age, 360000, ibb(age), tjp_par, year_(age), marginal)
''            'Debug.Print born; age; Kom
''            Worksheets("utdata").Cells(age, 70 + born - 1937) = Kom
''        Next age
''    Next born
''End Sub

Function tlakap_kr(ByVal alder As Long, ByVal inkomst As Double, ByVal IBB As Long, ByVal Zpar As Double, _
    ByVal year As Long, Optional marginal = 0) As Double
    'Kommunanställdas avgiftsbaserade
    'Alder      - Ålder 31/12 akuellt inkomstår
    'Inkomst    - Inkomst
    'IBB        - Inkomstbasbelopp
    'ZPAR       - Pensionsålder (Tjp_par) för tjänstepensionen
    'Year       - Inkomstår
    'Marginal   - Med avrundningar
    'AKAP-KL från och med 2014-01-01 för födda 1986 och senare
    'AKAP-KR från och med 2023-01-01 för samtliga efter valmöjlighet att lämna KAP-KL

    'Application.Volatile
    tlakap_kr = 0
    
''    Dim VAL_AKAP As Integer
''    VAL_AKAP = Application.Range("rng_TJP_Val").Value 'OBS läses från NameRange och D18
   
    Dim high_age As Integer
    '(A)KAP-KR injänar pensionsrätter fram till LAS ålder 67
    If year < 2003 Then
        high_age = 65
    ElseIf year > 2002 And year < 2023 Then '2003 LAS åldern 67 år
        high_age = 67
    Else
        high_age = 150 'ingen övre åldersgräns för premier på inkomster under 7,5 IBB
    End If
    
    Dim month As Integer '#antal månader med arbete och premieinbetalningar
    
    If Zpar > high_age Then Zpar = high_age 'Ger pensionsrätt efter riktålder
    
    If born > 1985 And alder < 21 And year >= 2014 Then
        If alder < W_start Then   'W_start public när individen börjar jobba
            Exit Function
        ElseIf W_start = alder Then
           month = Int((Int(born + alder + 1) - born - W_start) * 12)
           If month > 12 Then month = 12
           If month < 0 Then month = 0
           inkomst = inkomst * month / 12
        Else
            month = 12
        End If
    ElseIf alder < 21 Then
        Exit Function
    ElseIf alder = 21 Then
        month = Int((Int(born + 21 + 1) - born - W_start) * 12)
        If month > 12 Then month = 12 'Wstart före 21 års ålder
        If month < 0 Then month = 0 'Efter 21
        inkomst = inkomst * month / 12
    ElseIf alder < Int(Zpar) Then
            If alder > W_start Then
                month = 12
            Else
                month = Int((Int(born + W_start + 1) - born - W_start) * 12) 'Antal arbetande månder under året
                inkomst = inkomst * month / 12
            End If
    ElseIf alder = Int(Zpar) Then
        If Zpar < high_age Then
                month = Int((born + Zpar - Int(born + Zpar)) * 12)
                inkomst = inkomst * month / 12
        Else
                month = Int((born - Int(born)) * 12)
                inkomst = inkomst * month / 12
    
        End If
    Else
            month = 0
    End If
    
    If IBB = 0 Then IBB = inkomst
    If month = 0 Then Exit Function
    
    Dim par1 As Double
    Dim par2 As Double
    par1 = 0
    par2 = 0
    'Sysselssättningsgrad bortses från
''    If year < 1998 Then
      If year < 2003 Then
      par1 = 0.035:   par2 = 0.011
      If alder < 28 Then
         month = 0: par1 = 0: par2 = 0
      End If
''    ElseIf year <= 2002 Then
''      If alder > 27 Then
''       par1 = 0.034:  par2 = 0.01
''      End If
    ElseIf year < 2005 Then
      par1 = 0.035: If alder > 27 Then par2 = 0.011
''    ElseIf year = 2006 Or Int(born) <= 1946 Then
    ElseIf year = 2006 And Int(born) <= 1946 Then
      par1 = 0.045
      par2 = 0.021
    ElseIf year = 2006 And Int(born) > 1946 Then
      par1 = 0.04: par2 = 0.04
    ElseIf year <= 2007 Then
      par1 = 0.04: par2 = 0.04
    ElseIf year <= 2009 Then
      par1 = 0.0425: par2 = 0.0425
    ElseIf year > 2009 Then
      par1 = 0.045: par2 = 0.045
    End If
    
    If Int(born) > 1985 And year >= 2014 And year < 2023 Then 'Födda innan 1986 har istället haft förmånsbestämd par2. Vi har inte hanterat möjligheten att de med förmånsbestämd del kan ha valt bort denna och omfattas av AKAP-KR från och med 2023.
      par1 = 0.045: par2 = 0.3
    End If
    
    Dim LAS As Integer
    LAS = Application.Range("Rng_riktage") + 3 'Kolumn DS i Nyckeltal med lAS åldern! Den styrs sannolikat av riktålder+3
    
    If year >= 2023 And alder <= LAS Then 'par2 är kopplad till LAS-ålder från 2023 (69 år)
      par1 = 0.06: par2 = 0.315
    ElseIf year >= 2023 And alder > LAS Then
        par1 = 0.06: par2 = 0.06 '=0 tidigare modell
    End If
    
    'Fram till 2002 användes det förhöjda prisbasbeloppet därefter inkomstbasbeloppet (IBB). Nedan enbart IBB
    'Pensionsgrundande lön
    If inkomst > (30 * IBB) Then inkomst = 30 * IBB
    
    'If inkomst <= 0.01 * IBB Then inkomst = 0
    'Tillgodoräknas om inkomsten>1% av IBB, i annat fall betala arbetsgiv. motsvarande belopp till arbetstagaren i form av ersättning
    'som inte är pemnsionsgrundande, Arbetstagaren antas spara detta belopp "privat" till samma avkastning
    Select Case inkomst
        Case Is <= 7.5 * IBB 'limit
            tlakap_kr = par1 * inkomst
        Case Else
            tlakap_kr = par2 * (inkomst - 7.5 * IBB) + par1 * 7.5 * IBB
     End Select
     
    If marginal = 0 Then tlakap_kr = Int(tlakap_kr + 0.5) 'Premien som arbetsgivaren betala rill förvaltaren avrundas?

End Function

Function KAPKL_f(arsmedel, Optional ByVal aar = 30, Optional ByVal IBB = 45900, Optional ByVal fodar = 1959) As Double
 'Kommunalanställdas förmånsbestämda tjänstepension
 'arsmedel  - Medelinkomst
 'aar       - # år
 'ibb       - inkomstbasbeloppet
 'fodar     - Födelseår
 
     'Par2-3 ersättningsparametrar
     Dim par2 As Double
     Dim par3 As Double
     If fodar <= 1946 Then
         par2 = 0.625: par3 = 0.3125
     ElseIf fodar = 1947 Then
         par2 = 0.6214: par3 = 0.3107
     ElseIf fodar = 1948 Then
         par2 = 0.6179: par3 = 0.3089
     ElseIf fodar = 1949 Then
       par2 = 0.6143: par3 = 0.3071
     ElseIf fodar = 1950 Then
         par2 = 0.6107: par3 = 0.3054
     ElseIf fodar = 1951 Then
        par2 = 0.6071: par3 = 0.3036
     ElseIf fodar = 1952 Then
         par2 = 0.6036: par3 = 0.3018
     ElseIf fodar = 1953 Then
         par2 = 0.6: par3 = 0.3
     ElseIf fodar = 1954 Then
         par2 = 0.5964: par3 = 0.2982
     ElseIf fodar = 1955 Then
         par2 = 0.5929: par3 = 0.2964
     ElseIf fodar = 1956 Then
         par2 = 0.5893: par3 = 0.2946
     ElseIf fodar = 1957 Then
         par2 = 0.5857: par3 = 0.2929
     ElseIf fodar = 1958 Then
         par2 = 0.5821: par3 = 0.2911
     ElseIf fodar = 1959 Then
         par2 = 0.5786: par3 = 0.2893
     ElseIf fodar = 1960 Then
         par2 = 0.575: par3 = 0.2875
     ElseIf fodar = 1961 Then
         par2 = 0.5714: par3 = 0.2857
     ElseIf fodar = 1962 Then
         par2 = 0.5679: par3 = 0.2839
     ElseIf fodar = 1963 Then
         par2 = 0.5643: par3 = 0.2821
     ElseIf fodar = 1964 Then
         par2 = 0.5607: par3 = 0.2804
     ElseIf fodar = 1965 Then
         par2 = 0.5571: par3 = 0.2786
     ElseIf fodar = 1966 Then
         par2 = 0.5536: par3 = 0.268
     ElseIf fodar < 1986 Then
         par2 = 0.55: par3 = 0.275
     ElseIf fodar >= 1986 Then 'Lagt till och just 20140508 TL AKAP-KL obs ingen åldersgräns för dessa avgiftsbas..
         par2 = 0: par3 = 0
     End If
     'Beräkning av förmånen
     Dim benefit As Double
     
     If arsmedel > 30 * IBB Then
        benefit = (30 - 20) * IBB * par3 + (20 - 7.5) * par2 * IBB
     ElseIf arsmedel > 20 * IBB Then
        benefit = (arsmedel - 20 * IBB) * par3 + (20 - 7.5) * par2 * IBB
     ElseIf arsmedel > 7.5 * IBB Then
        benefit = (arsmedel - 7.5 * IBB) * par2
     Else
        benefit = 0
     End If
    
     
    If aar < 30 Then benefit = benefit * (aar / 30)
    KAPKL_f = benefit
    'Ingen avrundning
End Function
''Sub kapkl()
''    Dim kapkl As Long
''    kapkl = KAPKL_f(600000, 35, 59600, 1959)
''    Debug.Print Round(kapkl / 12, 2)
''End Sub

'********************************************
Function tjpkassa(ByVal tjp_par As Double, ByVal born As Double, ByVal alder As Long, ByVal pbh As Double, Optional ByVal marginal = 0, _
    Optional ByVal ips = 0) As Double
    'Årlig omräkningen av Tjänstepensionen
    'TJP_Par   - Pensionsålder vid första uttaget, ÅÅ,åå
    'Born      - Födelseår, ÅÅÅÅ,åå
    'Alder     - Ålder
    'pbh       - Pensionsbehållningen
    'Marginal  - 0 nuvarande regler, 1 Inga avrundningar (och det nya systemet i sin helhet)
    'IPS       - 0 tjp >0 IPS - koll av delningstal vid temp. uttag
    
    'Application.Volatile
 
''    If tjp_par < 55 Then def_ar = 65
''    If tjp_par > def_ar Then def_ar = tjp_par
''    If def_ar < 61# Then def_ar = tjp_par
''    If def_ar > 99 Then def_ar = tjp_par
    'alder = Int(alder)
    'born = Int(born)
    Dim del_tal As Double
    tjpkassa = 0
    del_tal = 0
    If pbh < 1 Then Exit Function 'pbh = 0
              
    'Alternativt att läsa in dem dirakt
    ' born = Worksheets("Start").Range("Born").Value
    ' tjp_par = Worksheets("Start").Range("tjp_Par").Value
    Dim Val As Integer
    Val = Application.Range("rng_TJP_Val").Value '2015 09 22 i syfte att korrigera delningstal för tjp
 
 
    Dim konst As Long  'Korrigering för inkomstår vid tjp_PAR 0 eller +1
    'Dim konst2 As Long 'Korrigering för inkomstår vid Def_ar 0 eller +1
    konst = 0
    'konst2 = 0
    If Int(born + tjp_par + 1 / 1000) > (Int(born) + Int(tjp_par)) Then konst = 1
    'If Int(born + def_ar + 1 / 1000) > (Int(born) + Int(def_ar)) Then konst2 = 1
 
    If born < 1938 Then
        del_tal = 13 'Delningstal före 1938 schablonm. 13
    ElseIf alder < 98 Then
        del_tal = deltal(tjp_par, born, alder, 999, 19)
        If del_tal <= 0 Then 'Finns inte för PP - godtyckligt
            del_tal = deltal(tjp_par + 1, born, alder, 999, 19) + 0.6
        End If
    Else: del_tal = 2
    End If
    
    del_tal = tjp_ddeltal(del_tal, Val, ips) 'Korrigering av delningstalet
   
    Application.Range("deltal_tjp").Value = del_tal
    
    Dim pyear As Double     'Pensionsår = Födelseår + pensionsålder
    Dim month As Long       'Month antal pensions månader under året
    pyear = born + tjp_par
    month = Tmonth          'Tmonth definerad som Public
    ' Tmonth = 12 - Int(12 * (born + TJP_par - Int(born + TJP_par)))
   Select Case alder
        Case Is < Int(tjp_par + konst) 'Inträffar inte men ..
            tjpkassa = 0
        Case Is = Int(tjp_par + konst)
           tjpkassa = pbh / del_tal
           'month = 12 - 12 * ((born + tjp_par) - Int(born + tjp_par + 1 / 1000)) 'Antal pensionsmånder totalt
        Case Else
            tjpkassa = pbh / del_tal
            month = 12
    End Select
    
    If marginal = 0 Then
       tjpkassa = Int(tjpkassa / 12 + 0.5) 'Månadsutbetalningen
    Else: tjpkassa = tjpkassa / 12
    End If
    tjpkassa = tjpkassa * month

End Function

'Sub kollaTJP()
'   'Function tjpkassa(tjp_par As Double, born As Double, alder As Long, pbh As Double, Optional marginal = 0) As Double
'   Dim kolla As Double
'   Dim tjp_par As Double
'   tjp_par = 65
'   'Dim born As Long
'
'   kolla = tjpkassa(tjp_par, 1995, 65, 1523189, 0)
'End Sub
'

Function tjp_ddeltal(ByVal tal As Double, ByVal Val As Integer, Optional ByVal ips = 0) As Double
    'tal    - delningstal
    'Val    - Avtalsområde
    'IPS    - Avser IPS
     
    'Application.Volatile
    
    Dim utbtid As Double
    If ips = 0 Then
        utbtid = Application.Range("rng_Temp_Tjp_Uttag")
    Else
        'val = 1 'Antar att sparande sker hos samma förvaltare
        utbtid = Application.Range("rng_Temp_IPS_Uttag")
    End If
    If utbtid <= 0 Then utbtid = 100
    If utbtid > 98 And Application.Range("rng_Ddelat").Value <> 1 Then 'Samma som premiepensionen
        tjp_ddeltal = tal
        Exit Function
    End If
    
    
   
    Dim kranta0 As Double   'Premiepensions netto
    Dim Life0 As Double     'Premiepensions E(65)
    Dim kranta As Double
    Dim Life As Double
        'Driftaavdrag och Förskottsräntan för fondförsäkran (traditionell)
        '2001-01-01 0,300 4,000 (0,3 4,00)
        '2002-12-01 0,300 3,000 (0,3 3,00)
        '2007-04-01 0,100 4,000 (0,1 2,30)
        '2014-03-01 0,100 3,000 (0,1 3,00)
        '2016-12-01 0,100 3,000 (0,1 1,75)
        '2017-12-01 0,100 1,750 (0,1 1,75)
    If IsEmpty(age) Then age = 65
    If IsEmpty(year_(age)) Then
        kranta0 = 1.75
        Life0 = 22.6
    Else
        If year_(age) < 2002 Then
            kranta0 = 4#      'förskottsräntan (realt) i procent
            Life0 = 20.3      'E(65) i premiepensionssystemet
        ElseIf year_(age) < 2008 Then
            kranta0 = 3
            Life0 = 20.3
        ElseIf year_(age) < 2015 Then
            kranta0 = 3
            Life0 = 20.3
        ElseIf year_(age) < 2018 Then
            kranta0 = 3
            Life0 = 20.3
         Else
            kranta0 = 1.75
            Life0 = 22.65
        End If
    End If
    
    'ränte och livsl. elasticiteter
    Dim b1, b2 As Double
    b1 = -0.27166     'Ränteelsaticitet
    b2 = 0.78457      'Livslängdskänslighet
    
    If year_(age) > 2019 Then
        b1 = -0.16216     'Ränteelsaticitet
        b2 = 0.84517     'Livslängdskänslighet
    End If
    
    Dim PAR As Double
    If Val > 0 Then PAR = tjp_par 'Worksheets("Start").Range("TJP_PAR")
    
    
    'Se PM deltal_sensitivy_tl.doc enkel kvot justering
    If ips > 0 Then Val = 1
    Select Case Val 'OBS samma som valen på startsidan
        Case Is = 0  'PP men korr för räntediff
            kranta = Application.Range("rng_FondAvkastning")
            If kranta < 0.1 Then kranta = 100 * kranta
            'life = life0
            'Tidigare till Erik F på fond och fontorgsutredningen beräknat elasticiteten till
            tjp_ddeltal = tal * (1 - 0.1812 * (kranta / kranta0 - 1))
            Exit Function
        Case Is = 1  'Saknar TJP men för IPS
            kranta = 2.5
            Life = 22
            If year_(age) > 2019 Then
                kranta = 3.5
                Life = 22.6
            End If
        Case Is < 4  'ITP 1 & 2
            kranta = 2.9
            Life = 22
            If year_(age) > 2019 Then
                kranta = 2.2
                Life = 22
            End If
        Case Is = 4  'SAF-LO
            kranta = 2.25
            Life = 21.3
            If year_(age) > 2019 Then
                kranta = 1.3
                Life = 20.8
            End If
        Case Is < 7 '(A) KAP-KL
            kranta = 2.75
            Life = 22.8
            If year_(age) > 2019 Then
                kranta = 2.75
                Life = 23.1
            End If
            
        Case Is = 7  'PA16 avd 2 före 1988
            kranta = 2
            Life = 21.9
            If year_(age) > 2019 Then
                kranta = 2#
                Life = 23.3
            End If
         Case Else 'Stat Pa16 avd 1 1988 +
            kranta = 2#
            Life = 21.9
             If year_(age) > 2019 Then
                kranta = 2#
                Life = 23.3
            End If
     End Select
     
    'kranta = kranta * (1 - 0.15 * 0.02) 'Korr för avkastningsskatt, 2% statsobligationsräntan
    'Räntan torde vara korrigerad för avkastningsskatt
    'Obs sedan 2016 - Max(Statslåneräntan +0.0075;0.0125)
      If utbtid > 0 And utbtid < 30 Then
        'Makeham U(X) = a + b exp(cX)
        Dim my As Double, surv As Double
        Dim S As Double
        Dim i As Integer
        my = (0.0002 + 0.000007 * Exp(0.1071 * PAR))
        surv = (1 - my)
        'Lite historielöst och avkastningen spelar väl roll
        For i = PAR To (PAR + Int(utbtid) - 1)
            my = 0.0002 + 0.000007 * Exp(0.1071 * i)
            S = S + (1 - my) / (1 + kranta / 100) ^ (i - PAR)
            'Debug.Print i; my; S
        Next i
        
        If utbtid > 0 Then
            tjp_ddeltal = S / surv
            'Debug.Print tjp_ddeltal; S; surv;
            Exit Function
        End If
    End If
  
''    If utbtid < 99 Then
        tjp_ddeltal = tal * (1 + b1 * (kranta / kranta0 - 1) + b2 * (Life / Life0 - 1))
''    Else
''        tjp_ddeltal = tal
''    End If
   
    If marginal = 0 Then tjp_ddeltal = Round(tjp_ddeltal, 2)
    'Debug.Print tjp_ddeltal; tal; Round(tal / tjp_ddeltal, 3)
End Function

''' 'Nedan funkar inte eftersom year_() inte hänger med
'''Sub kolla_tjpdeltal()
'''    Dim tal As Double, kolla As Double
'''    Dim val, age As Integer
'''    tal = 20
'''    'Dim age As Integer
'''    ReDim year_(0 To 100) As Long
'''    For age = 65 To 66
'''        year_(age) = 1954 + age
'''    Next age
'''
'''
'''
'''    For val = 2 To 2
'''
'''        For age = 65 To 66
'''            year_(age) = 2019 + age - 65
'''            kolla = tjp_ddeltal(tal, val)
'''            Debug.Print year_(age); val & " " & Round(kolla, 4) ' ; behåller raden
'''
'''        Next age
'''    Next val
'''
'''End Sub


