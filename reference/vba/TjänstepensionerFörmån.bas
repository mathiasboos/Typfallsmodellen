Attribute VB_Name = "TjänstepensionerFörmån"
Function FTJP(ByVal TJPage, ByVal startage, ByVal age, ByVal tjp_par, ByVal born, ByVal avtal, ByRef STP_points() As Double, ByRef tp() As Double, Optional ByVal marginal = 0) As Double
'Funktion som returnerar olika kollektivavtals förmånsbaserade tjänstepension. Input variablerna är som följer:
'TJPage         - Intjänad tjänstepension vid åldern age
'startage       - Startålder
'age            - Ålder
'tjp_par        - Pensionsålder tjänstepensionen
'born           - Födelseår
'avtal          - Kollektivavtal
'STP_points     - STP poäng SAF-LO, använder ATP poäng
'tp             - Tilläggspension (ATP)
'marginal       - Med avrundningar
            
            If avtal = 3 Then 'ITP 2 + förmånsabaserat
                tp_year = 0
                For counter = startage To 68 'Kollas
                    If Wage_(counter) > 0 Then tp_year = tp_year + 1 'Antar samma kollektiavtal korrigerat TL
                Next
                If age = 65 Then
                    diverse = 1
                Else
                    diverse = deltal(65, Int(born), 65, 99, 19) / _
                            deltal(tjp_par, Int(born), age, 99, 19)
                End If
                 'Korrigering för temp. uttag
                If Application.Range("rng_Temp_Tjp_Uttag") > 0 Then
                    If born > 1938 Then
                        diverse = diverse * deltal(tjp_par, Int(born), age, 99, 19) / _
                        tjp_ddeltal(15, 4)
                    Else 'OBS fast adress
                        diverse = diverse * Worksheets("mortality").Range("j27") / _
                        tjp_ddeltal(15, 4)
                    End If
                End If
                underlag = DC_underlag(0) 'Löneunderlag med ev. kapningar
                underlag = tlITP2F(underlag, IBB(age - 1), tp_year, marginal) * diverse
                TJPage = TJPage + underlag * Tmonth / 12
            End If 'ITP 2
            
            If avtal = 4 Then 'SAF-LO (STP)
                underlag = 0
                tp_year = 0
                For counter = 28 To 64
                    If born < 1968 Then
                        If year_(counter) < 1996 Then
                            tp_year = tp_year + 1
                            underlag = underlag + STP_points(counter) 'OBS TP_points är omsorterad
                        End If
                    Else 'Född 1968 och senare medel 55-59 års ålder
                        If counter >= 55 And counter <= 59 And STP_points(counter) > 0 Then
                           tp_year = tp_year + 1
                           underlag = underlag + STP_points(counter) / 5
                        End If
                    End If
                Next
                'Debug.Print underlag; tp_year; underlag / tp_year + 1
                 
                If year_(65) > 1960 Then 'OBS att STP fås givet att pensionering sker efter 1959 då prisbasbelopp kunde bestämmas
                    If tjp_par < 65 Then
                        diverse = FPB(Int(tjp_par))
                    Else
                        diverse = FPB(65) 'Age at 1995
                    End If
                Else
                        diverse = 0
                End If
                
                If tp_year > 0 And underlag > 0 Then
                    underlag = STP_(underlag / tp_year + 1, tp_year, diverse)
                Else
                    underlag = 0
                End If
                
                If age = 65 Then
                    diverse = 1
                Else
                    diverse = deltal(65, Int(born), 65, 99, 19) / _
                            deltal(tjp_par, Int(born), age, 99, 19)
                End If
                'Korrigering för temp. uttag
                If Application.Range("rng_Temp_Tjp_Uttag") > 0 Then
                    If born > 1938 Then
                        diverse = diverse * deltal(tjp_par, Int(born), age, 99, 19) / _
                        tjp_ddeltal(15, 4)
                    Else
                        diverse = diverse * Worksheets("mortality").Range("j27") / _
                        tjp_ddeltal(15, 4)
                    End If
                End If
                TJPage = TJPage + underlag * diverse * Tmonth / 12
                
            End If 'SAF-LO
       
            If avtal = 5 Or avtal = 6 Then '(A)PA-KL och KAP-KL ev. förmånsabaserat
                Dim KPA_Und(1 To 7) As Double
                tp_year = 0
                diverse = 1
                For counter = 28 To 70 'KOLLAS de 7 åren 2 år före pensioneringen
                    If STP_points(counter) > 0 Then tp_year = tp_year + 1
                    If (tjp_par - counter >= 2) And (tjp_par - counter < 9) Then
                        KPA_Und(diverse) = mini(Wage_(counter), IBB(counter) * 30)
                        If year_(counter) >= 1960 Then KPA_Und(diverse) = KPA_Und(diverse) * KPI_j(tjp_par - 1) / KPI_j(counter)
                        'Debug.Print diverse; Round(KPA_Und(diverse), 0); tp_year
                        diverse = diverse + 1
                    End If
                Next
                'Av de 7 åren:  medelvärdet av de fem bästa åren av dessa
                Call QuickSort(KPA_Und, LBound(KPA_Und), UBound(KPA_Und)) 'lönen sorterade i stigande ordning
                underlag = 0
                For counter = 3 To 7
                     underlag = underlag + KPA_Und(counter) / 5 'medel av de 5 bästa åren
                     'Debug.Print Round(underlag, 0); Round(KPA_Und(counter), 0)
                Next
                If marginal = 0 Then underlag = Round(underlag, 0)
               
                
                If age = 65 Then
                    diverse = 1
                Else
                    diverse = deltal(65, Int(born), 65, 99, 19) / _
                            deltal(tjp_par, Int(born), age, 99, 19)
                End If
                'Korrigering för temp. uttag
                If Application.Range("rng_Temp_Tjp_Uttag") > 0 Then
                    If born > 1938 Then
                        diverse = diverse * deltal(tjp_par, Int(born), age, 99, 19) / _
                        tjp_ddeltal(15, 4)
                    Else
                        diverse = diverse * Worksheets("mortality").Range("j27") / _
                        tjp_ddeltal(15, 4)
                    End If
                End If
                
                Dim paklbpp As Double
                'År efter PA-KL ersatt med PFA98 (senare KAP-KL)
                If year_(age) > 1997 Then
                    'Börjar tjäna in premier från 28 års ålder i KAP-KL (Om inga år innan ingen PA-KL, då bara KAPKL_F)
                    'Intjänad pensionsrätt 1997 skrivs upp med 8% för att kompensera för förändringar i skydd till efterlevande och räknas upp med IBB
                    If year_(28) < 1995 Then
                    
                    Dim workyear As Double
                    
                    paklbpp = PA_KLBPP(1997 - Int(born))
                    workyear = 1997 - year_(W_start)
                    If workyear > 30 Then workyear = 30
                    'Född innan 1938 full ATP
                    If Int(born) < 1938 Then
                        underlag = (PA_KL(1997 - Int(born), 1997, paklbpp, marginal) - tp(age)) * 1.08 * (IBB(age - 1) / IBB(age - year_(age) + 1998))
                    'Drar bort full ATP beräknat vid 1997. Enligt KPA minsta av ATP-poäng och PA-KL BPP, men gav för hög TjP
                    Else
                        underlag = (PA_KL(1997 - Int(born), 1997, paklbpp, marginal) - _
                        (0.6 * STP_points(age - year_(age) + 1997) * (workyear / 30) * pbb(age - year_(age) + 1997)) - _
                        (0.96 * pbb(age - year_(age) + 1997) * (workyear / 30))) * _
                        1.08 * (IBB(age - 1) / IBB(age - year_(age) + 1998))
                    End If
                            If underlag < 0 Then
                                underlag = KAPKL_f(underlag, tp_year, IBB(tjp_par - 1), born)
                            Else
                                underlag = KAPKL_f(underlag, tp_year, IBB(tjp_par - 1), born) + underlag
                            End If
                    Else
                        underlag = KAPKL_f(underlag, tp_year, IBB(tjp_par - 1), born)
                    End If
                Else
                    paklbpp = PA_KLBPP(tjp_par)
                    underlag = PA_KL(tjp_par, year_(age), paklbpp, marginal) - tp(age)
                    'Garanterad nivå i PA-KL är 100 kr per månad (Gäller bara för dem som går i pension under PA-KL, i.e. innan 1997)
                    If underlag < 1200 Then
                        underlag = 1200
                    End If
                End If
                
                TJPage = TJPage + underlag * diverse
                
            End If '(A)KAP-KL
            
            
            If avtal = 7 Then 'PA03 + förmånsabaserat
                underlag = 0
                tp_year = 0
                For counter = 28 To 64
                    If Wage_(counter) > 0 Then tp_year = tp_year + 1
                Next
                
''                For counter = (tjp_par - 5) To (tjp_par - 1)
''                    underlag = underlag + (Income_(counter) * KPI_j(tjp_par - 1) / KPI_j(counter)) / 5 'Senaste 5 åren
''                    'Debug.Print Round(underlag, 0); Round(Income_(counter) * KPI_j(tjp_par - 1) / KPI_j(counter), 0)
''                Next
                underlag = DC_underlag(1) 'Innehåller ITPS lönekapning
                
                If age = 65 Then
                    diverse = 1
                Else
                    'Ev ersätta från fliken mortality
                    diverse = deltal(65, maxi(1938, Int(born)), 65, 99, 19) / _
                            deltal(tjp_par, maxi(1938, Int(born)), age, 99, 19)
                    
                End If
                'Korrigering för temp. uttag
                If Application.Range("rng_Temp_Tjp_Uttag") > 0 Then
                    If born > 1938 Then
                        diverse = diverse * deltal(tjp_par, Int(born), age, 99, 19) / _
                        tjp_ddeltal(15, 4)
                    Else
                        diverse = diverse * Worksheets("mortality").Range("j27") / _
                        tjp_ddeltal(15, 4)
                    End If
                End If
                
                underlag = tlPA03(underlag, tp_year, IBB(age - 1), born) * diverse / 12
                If marginal = 0 Then underlag = Int(underlag + 0.49)
                TJPage = TJPage + underlag * Tmonth '/ 12
                
            End If 'PA03
            
            FTJP = TJPage
            
End Function

