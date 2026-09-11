Attribute VB_Name = "Lön_mm"
Option Explicit

Function wages(ByVal age As Integer, ByVal year As Long, ByVal wstart As Double, ByVal wage As Double, _
    Optional w_time = 0, Optional ref_ar = 2011, Optional nominalw = True) As Double
     'Application.Volatile
     'Beräkning av lön under förvärvsaktiv ålder
     'AGE       'Age XX den 31/12
     'YEAR      'Kalenderår yyyy
     'Wstart    'Ålder för inträde på arbetsmarknaden ÅÅ,åå
     'Wage      'Referens lön kr
     'W_time    'inkomstår
     'ref_ar    'År för referenslönen ÅÅ
     'nominalw  'Lön i nominella eller i fasta priser
    If pblnCloseOrSave Then Exit Function
    On Error GoTo errTag
    
    wages = 0
    If age < wstart Then Exit Function
    Dim ink_ref As Double   'D Inkomstindex
    Dim ink_ref2 As Double   'D Inkomstindex - löneprofil
    Dim KPI_ref As Double   'D KPI
    Dim konst As Single     'Korrigering år för första pensionsuttaget
    Dim konst2 As Single    'Korrigering år för definitiva pensionsuttaget
    Dim konst3 As Single    'Korrigering år för inträdet på arbetsmarknaden
    Dim kolumn As Long      'Kolumnnr för inkomstidex/kpi
    Dim month As Long       'Antal månader under inkomståret och omfattning
    Dim w_slut As Integer   'Referensår för (slut) lön
    
    'Uttagsandel, rättat variablen var tidigare inte initierad 2024-03-16
    If age < Int(PAR) Then
        uttagIP = 0
    ElseIf age >= Int(PAR) And PAR <= def_ar And age < def_ar Then
    '20250826 age<def_ar men se över choicelön... och förklaringen i texten samt default...
        Dim ChoiceLönPartUttag As Variant
        'Rad i flik AdvSettings, där man kan låta lönen följa partiellt uttag, eller vara 100% eller 0%
        ChoiceLönPartUttag = ThisWorkbook.Names("rngLönPartUttag").RefersToRange.Value
        If IsNumeric(ChoiceLönPartUttag) Then
            'Antingen 100% eller 0%
            uttagIP = 1 - ChoiceLönPartUttag
        Else
            'Följer andelen partiellt uttag
            uttagIP = Application.Range("UttagIP").Value
        End If
        If PAR = def_ar Then
            uttagIP = 1
        End If
    Else
        If age > def_ar Then
            uttagIP = 1
        End If
    End If
    
    
    w_slut = Int(born + w_time) 'Inkomstår referens
    kolumn = Application.Range("Iindex").Column  'Kolumnen för inkomstindex
    
    konst = 0  'För att fånga rätt år vid pensionering, PAR
    konst2 = 0 'För att fånga rätt år vid slutlig avgång, def_ar
    konst3 = 0 'För att fånga rätt år vid inträde på arbetsmarknaden, wstart
    
    If Int(born + PAR + 1 / 1000) > (Int(born) + Int(PAR)) Then konst = 1
    If Int(born + def_ar + 1 / 1000) > (Int(born) + Int(def_ar)) Then konst2 = 1
    If Int(born + wstart + 1 / 1000) > (Int(born) + Int(wstart)) Then konst3 = 1
     
    If age < Int(wstart + konst3) Then Exit Function 'Före inträde på arbetsmarknaden

    If age <= Int(def_ar + konst) And age >= Int(wstart) Then
        If year < 1958 Then
            ink_ref = wsNyckelTal.Cells(1958 - 1960 + 7, kolumn) / _
                wsNyckelTal.Cells(w_slut - 1960 + 7, kolumn)
            ink_ref = ink_ref / (1.08 ^ (1958 - year))  'Index saknas före 1957 antas till nominellt 8% per år
        Else
            ink_ref = wsNyckelTal.Cells(year - 1960 + 7, kolumn) _
                    / wsNyckelTal.Cells(w_slut - 1960 + 7, kolumn)
        End If
        wage = wage * ink_ref
        
        'Löneprofil - w= f(age, kön) se PM migr_svar_regleringsbrev 2017
        Dim B0, b1, b2, b3, b4, b5 As Double
        If Application.Range("wage_profil") > 0 Then
            If Application.Range("wage_profil") = 1 Then
                'Inrikesfödda ersätt senare med låg profil
                B0 = -33.01341746
                b1 = 4.354364353
                b2 = -0.212974964
                b3 = 0.005053746
                b4 = -0.0000583887
                b5 = 0.000000263128
                
                If age < 61 Then
                    ink_ref = B0 + b1 * age + b2 * (age ^ 2) + b3 * (age ^ 3) + b4 * (age ^ 4) + b5 * (age ^ 5)
                Else
                     ink_ref = B0 + b1 * 61 + b2 * (61 ^ 2) + b3 * (61 ^ 3) + b4 * (61 ^ 4) + b5 * (61 ^ 5)
                End If
                If w_time < 61 Then
                    ink_ref = ink_ref / (B0 + b1 * w_time + b2 * (w_time ^ 2) + b3 * (w_time ^ 3) + b4 * (w_time ^ 4) + b5 * (w_time ^ 5))
                Else
                    ink_ref = ink_ref / (B0 + b1 * 61 + b2 * (61 ^ 2) + b3 * (61 ^ 3) + b4 * (61 ^ 4) + b5 * (61 ^ 5))
                End If
                If ink_ref < 0 Then ink_ref = 0
                wage = wage * ink_ref
            ElseIf Application.Range("wage_profil") = 2 Then
                B0 = -20.6648551
                b1 = 2.496893883
                b2 = -0.10843247
                b3 = 0.002287573
                b4 = -0.0000236622
                b5 = 0.0000000962164
            ElseIf Application.Range("wage_profil") = 3 Then
                B0 = -23.73767932
                b1 = 3.104093675
                b2 = -0.149408507
                b3 = 0.003477166
                b4 = -0.0000392148
                b5 = 0.000000171834
            ElseIf Application.Range("wage_profil") = 4 Then
                B0 = -17.18584424
                b1 = 2.203777173
                b2 = -0.10280891
                b3 = 0.00234388
                b4 = -0.0000261194
                b5 = 0.000000113668

            Else
            
            End If
            
            If Application.Range("wage_profil") > 1 Then
                ink_ref = B0 + b1 * age + b2 * (age ^ 2) + b3 * (age ^ 3) + b4 * (age ^ 4) + b5 * (age ^ 5)
                'Ekvationen först utvecklingen vid referensåret för att nå rätt belopp
                ink_ref = ink_ref / (B0 + b1 * w_time + b2 * (w_time ^ 2) + b3 * (w_time ^ 3) + b4 * (w_time ^ 4) + b5 * (w_time ^ 5))
            
                If ink_ref < 0 Then ink_ref = 0
                wage = wage * ink_ref
            End If
            
        End If
    End If
    
    
    If nominalw = False Then 'Uttryckt i fasta priser: -> Wage * KPI(ref_ar) / Kpi(vid året w_slut)
          kolumn = Application.Range("KPI").Column
          Rem   Serien ska var i löpande priser! så KPi(wslut)/Kpi(ar) ger rätt när serien sedan tas i fasta priser
                KPI_ref = wsNyckelTal.Cells(w_slut - 1960 + 7, kolumn) _
                           / wsNyckelTal.Cells(ref_ar - 1960 + 7, kolumn)
          wage = wage * KPI_ref
    End If
        
    If age = Int(wstart + konst3) Then
        month = 12 * ((born + wstart) - Int(born + wstart + 1 / 1000)) 'När under året
        month = Int(12 - month)
    ElseIf age < Int(PAR + konst) Then
        month = 12
    ElseIf age = Int(PAR + konst) Then
        If age = Int(def_ar + konst2) Then
           
           month = Int(12 * ((born + PAR) - Int(born + PAR + 1 / 1000))) + Int(12 * (def_ar - PAR) * (1 - uttagIP)) 'Antal # månader med arbete
        Else
            'month = 12 * ((born + par) - Int(born + par + 1 / 1000)) 'Antal månader som pensionär
            month = Int(12 * (born + PAR - Int(born + PAR + 1 / 1000)))
            month = month + ((12 - month) * (1 - uttagIP))       ' med hänsyn till # hela månader
        End If
    ElseIf age < Int(def_ar + konst2) Then
        month = 12 * (1 - uttagIP)
    ElseIf age = Int(def_ar + konst2) Then
        month = 12 * ((born + def_ar) - Int(born + def_ar + 1 / 1000)) * (1 - uttagIP)
    Else
        month = 0
    End If
    
    wages = month * wage / 12
    
    Exit Function
    
errTag:
            wages = CVErr(xlErrValue)
End Function


''Sub koll_wages()
''    Dim year As Long
''    Dim age As Integer
''    Call startsetup 'Finns i VBA_go och ger index mm
''
''    Dim w As Double
''    Dim wage As Double
''    wage = 360000
''    born = 1960
''    'age = 23
''
''    For age = 22 To 65
''        year = Int(born) + age
''        Rem Wages(Age , year , wstart , wage , w_time = 0, ref_ar = 2011, nominalw = False)
''        ' Wage_(Age) = wages(Age, year_(Age), W_start, Income, w_time, w_ref, Nominal)
''        w = wages(age, year, W_start, wage, w_time, 2019, 0)
''
''        Debug.Print born; year; age; Round(w, 2)
''    Next 'Age
''End Sub


'Kopiera slumptalen utan att spara så att resultatet kan "återskapas"
Sub copyRnd()
   Range("RND_dst").Value = Range("RND_scr").Value
End Sub



