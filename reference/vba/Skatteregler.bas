Attribute VB_Name = "Skatteregler"
'-------------------------------------------------- Skatte funktioner ----------------------------------------------------------
  
'Function fastred(inkomst, fastskatt) As Double
 'Dim sparr As Double
 'Dim limit As Double
   ' Endast för förtids- och ålderspernsionärer
 ' sparr = 0.04 * (inkomst) '+kapital,
 ' limit = 2800 * Int(10000 * 42800 / 41000) / 10000 'procent med 2 decimaler förädring i ink.PBBeloppet?
 ' limit = Int(limit + 0.5) 'Öresdelen tas bort round?
 
 ' If sparr < limit Then sparr = limit
 '   fastred = Application.Max(fastskatt - sparr, 0)
'End Function

Function Xage(ByVal year As Single) As Single
    'Hjälpfunktion i modellen för vilken ålder när de äldre har ett förhöjt grundavdrag
    'Year inkomstår
        If year <= 2023 Then
            Xage = 66 'Födda 57 får 2024 kompensation...
        ElseIf year <= 2026 Then
            Xage = 67 'Födda 59 sista årskull
        ElseIf year <= 2028 Then
            Xage = 68
        Else
            Xage = Application.Range("Rng_riktage") + 1 'Prognos blir riktålder + ett år
        End If

End Function

Function avkskatt(ByVal year As Long, Optional Val = 1) As Double
'Funktion för att beräkna avkastningsskatten. Sedan 2017 finns det ett golv för räntan om 0,5% avseende tjp.
'För kapital är det statslåneräntan 30 nov. t-1, golv om 0.25% men 1%enhet + räntan, dvs lägst 1.25%
'Skattesatsen är 15 procent för pensionsförsäkringskapital, 30% för kapitalförsäkringar
'Inputvariabler är som följer:
'year   - Inkomstår
'val    - Typ av kapital, 0 = Pensionsförsäkring, 1 = ISK / Kapitalförsäkring
'Bondi  - genosmnittliga statslåneräntan året före beskattningsåret, dvs t-1

    'Application.Volatile
    
    Dim bondi As Double
    bondi = Application.Range("Rgk")
    
    Dim year2 As Integer
    year2 = year - 1 '30 nov året före. Nedan anges årets räntesatser
    Dim Update As Integer
    'https://www.riksgalden.se/sv/var-verksamhet/statslanerantan/statslanerantan-per-vecka/?year=2021
    Select Case year2
    Case Is < 1998
        bondi = 0.0498
    Case Is = 1998
        bondi = 0.05
    Case Is = 1999
        bondi = 0.0489
    Case Is = 2000
        bondi = 0.0534
    Case Is = 2001
        bondi = 0.0497
    Case Is = 2002
        bondi = 0.0515
    Case Is = 2003
        bondi = 0.0439
    Case Is = 2004
        bondi = 0.043
    Case Is = 2005
        bondi = 0.0325
    Case Is = 2006
        bondi = 0.0361
    Case Is = 2007
        bondi = 0.0413
    Case Is = 2008
        bondi = 0.0388
    Case Is = 2009
        bondi = 0.0309
    Case Is = 2010
        bondi = 0.0277
    Case Is = 2011
        bondi = 0.0258
    Case Is = 2012
        bondi = 0.0152
    Case Is = 2013
        bondi = 0.02
    Case Is = 2014
        bondi = 0.0163
    Case Is = 2015
        bondi = 0.0058
    Case Is = 2016
        bondi = 0.0034
    Case Is = 2017
        bondi = 0.0051
    Case Is = 2018
        bondi = 0.0048
    Case Is = 2019
        bondi = 0.0004
    Case Is = 2020
        bondi = -0.0007
    Case Is = 2021
        bondi = 0.0016
    Case Is = 2022
        bondi = 0.0194
    Case Is = 2023
        bondi = 0.0262
    Case Is = 2024
        bondi = 0.0196
    Case 2025
        bondi = 0.0255
    
    Case Else ' Bortkommenterad eftersom modellen istället använder Rgk från fliken Adv_settings (som beror på antagande för inflation och real tillväxt)
        bondi = 0.025 '0.025 prognos för framtida statslåneränta enligt prognosstandard.
'        If Application.Range("Modell_year") > year2 Then
'            MsgBox ("Uppdatera statslåneräntan i VBA-koden för " & year2 + 1)
'        End If
    End Select
    
    If Application.Range("RulesfromSkatt") = 0 Or IsMissing(Application.Range("RulesfromSkatt")) Then
         'Nothing
    Else
        If (year > Application.Range("RulesfromSkatt") And Application.Range("Rules") = 0) _
        Or Application.Range("Rules") = 1 Then
            year = Application.Range("RulesfromSkatt")
        End If
    End If
    
    If year > 2016 And Val = 0 Then
        'k Pensionsförsäkring
        If bondi < 0.005 Then bondi = 0.005
    End If
    'Kapitalförsäkring/ISK:
    If Val = 0 Then
        avkskatt = bondi * 0.15
    Else 'ISK
        If bondi < 0.0025 Then bondi = 0.0025
        avkskatt = (0.01 + bondi) * 0.3
    End If
  
End Function

Function avdrag95(ByVal inkomst, Optional ByVal pbb = 41000, Optional ByVal marginal = 0, Optional ByVal bald = 64, _
Optional ByVal year = 2100, Optional ByVal wyear = 21000, Optional ByVal IBB = 0, Optional ByVal kvoten = 1, Optional ByVal Iyear = 0, _
Optional ByVal Xage = 66) As Double
  
  'Year - inkomstår, wyear=När indexeringen ska börja, iyear=skattreglerna
   'Application.Volatile
    If year > wyear And kvoten < 1 Then
       pbb = IBB 'Ersätt pbb med IBB och korrigera gränserna till de gamla
       pbb = pbb * kvoten
    End If
    'Skapat av Tommy Lowén beräkning av grundavdraget
   Dim avdrag As Double
   Dim Extra As Double
   
   If inkomst < 0.25 * pbb Then
       avdrag = inkomst
    ElseIf inkomst <= 1.86 * pbb Then
       avdrag = 0.25 * pbb
    ElseIf inkomst <= 2.89 * pbb Then
       avdrag = 0.25 * pbb + 0.25 * (inkomst - 1.86 * pbb)
    ElseIf inkomst <= 3.04 * pbb Then
       avdrag = 0.25 * pbb + 0.25 * (2.89 - 1.86) * pbb
    Else
       avdrag = 0.25 * pbb + 0.25 * (2.89 - 1.86) * pbb - 0.1 * (inkomst - 3.04 * pbb)
       If avdrag < 0.25 * pbb Then avdrag = 0.25 * pbb
    End If
     
   If avdrag > inkomst Then avdrag = inkomst
   If marginal = 0 Then avdrag = Int((avdrag + 99.99) / 100) * 100 'Avrundat uppåt till närmaste hundratal kr
   avdrag95 = avdrag
End Function

Function avdrag96(ByVal inkomst, Optional ByVal pbb = 41000, Optional ByVal marginal = 0, Optional ByVal bald = 64, _
Optional ByVal year = 2100, Optional ByVal wyear = 21000, Optional ByVal IBB = 0, Optional ByVal kvoten = 1, Optional ByVal Iyear = 0, _
Optional ByVal Xage = 66) As Double
  
  'Year - inkomstår, wyear=När indexeringen ska börja, iyear=skattreglerna
   'Application.Volatile
    If year > wyear And kvoten < 1 Then
       pbb = IBB 'Ersätt pbb med IBB och korrigera gränserna till de gamla
       pbb = pbb * kvoten
    End If
    'Skapat av Tommy Lowén beräkning av grundavdraget
   Dim avdrag As Double
   Dim Extra As Double
   
   If inkomst < 0.24 * pbb Then
       avdrag = inkomst
    ElseIf inkomst <= 1.86 * pbb Then
       avdrag = 0.24 * pbb
    ElseIf inkomst <= 2.89 * pbb Then
       avdrag = 0.24 * pbb + 0.25 * (inkomst - 1.86 * pbb)
    ElseIf inkomst <= 3.04 * pbb Then
       avdrag = 0.24 * pbb + 0.25 * (2.89 - 1.86) * pbb
    Else
       avdrag = 0.24 * pbb + 0.25 * (2.89 - 1.86) * pbb - 0.1 * (inkomst - 3.04 * pbb)
       If avdrag < 0.24 * pbb Then avdrag = 0.24 * pbb
    End If
     
   If avdrag > inkomst Then avdrag = inkomst
   If marginal = 0 Then avdrag = Int((avdrag + 99.99) / 100) * 100 'Avrundat uppåt till närmaste hundratal kr
   avdrag96 = avdrag
End Function


Function avdrag01(ByVal inkomst, Optional ByVal pbb = 41000, Optional ByVal marginal = 0, Optional ByVal bald = 64, _
Optional ByVal year = 2100, Optional ByVal wyear = 21000, Optional ByVal IBB = 0, Optional ByVal kvoten = 1, Optional ByVal Iyear = 0, _
Optional ByVal Xage = 66) As Double
   'Year - inkomstår, wyear=När indexeringen ska börja, iyear=skattreglerna
   'Application.Volatile
    If year > wyear And kvoten < 1 Then
       pbb = IBB 'Ersätt pbb med IBB och korrigera gränserna till de gamla
       pbb = pbb * kvoten
    End If
    'Skapat av Tommy Lowén beräkning av grundavdraget
   Dim avdrag As Double
   Dim Extra As Double
   
   If inkomst < 0.27 * pbb Then
       avdrag = inkomst
    ElseIf inkomst <= 1.86 * pbb Then
       avdrag = 0.27 * pbb
    ElseIf inkomst <= 2.89 * pbb Then
       avdrag = 0.27 * pbb + 0.25 * (inkomst - 1.86 * pbb)
    ElseIf inkomst <= 3.04 * pbb Then
       avdrag = 0.27 * pbb + 0.25 * (2.89 - 1.86) * pbb
    Else
       avdrag = 0.27 * pbb + 0.25 * (2.89 - 1.86) * pbb - 0.1 * (inkomst - 3.04 * pbb)
       If avdrag < 0.27 * pbb Then avdrag = 0.27 * pbb
    End If
     
   If avdrag > inkomst Then avdrag = inkomst
   If marginal = 0 Then avdrag = Int((avdrag + 99.99) / 100) * 100 'Avrundat uppåt till närmaste hundratal kr
   avdrag01 = avdrag
End Function

Function avdrag03(ByVal inkomst, Optional ByVal pbb = 41000, Optional ByVal marginal = 0, Optional ByVal bald = 64, _
Optional ByVal year = 2100, Optional ByVal wyear = 21000, Optional ByVal IBB = 0, Optional ByVal kvoten = 1, Optional ByVal Iyear = 0, _
Optional ByVal Xage = 66) As Double
  'Year - inkomstår, wyear=När indexeringen ska börja, iyear=skattreglerna
   'Application.Volatile
    If year > wyear And kvoten < 1 Then
       pbb = IBB 'Ersätt pbb med IBB och korrigera gränserna till de gamla
       pbb = pbb * kvoten
    End If
    'Skapat av Tommy Lowén beräkning av grundavdraget
   Dim avdrag As Double
   Dim Extra As Double
   
   If inkomst < 0.423 * pbb Then
       avdrag = inkomst
    ElseIf inkomst <= 1.49 * pbb Then
       avdrag = 0.423 * pbb
    ElseIf inkomst <= 2.72 * pbb Then
       avdrag = 0.423 * pbb + 0.2 * (inkomst - 1.49 * pbb)
    ElseIf inkomst <= 3.11 * pbb Then
       avdrag = 0.67 * pbb
    ElseIf inkomst <= 6.87 * pbb Then
       avdrag = 0.67 * pbb - 0.1 * (inkomst - 3.11 * pbb)
    Else
       avdrag = 0.293 * pbb
    End If
     
   If avdrag > inkomst Then avdrag = inkomst
   If marginal = 0 Then avdrag = Int((avdrag + 99.99) / 100) * 100 'Avrundat uppåt till närmaste hundratal kr
   avdrag03 = avdrag
End Function

Function avdrag05(ByVal inkomst, Optional ByVal pbb = 41000, Optional ByVal marginal = 0, Optional ByVal bald = 64, _
Optional ByVal year = 2100, Optional ByVal wyear = 21000, Optional ByVal IBB = 0, Optional ByVal kvoten = 1, Optional ByVal Iyear = 0, _
Optional ByVal Xage = 66) As Double
  'Year - inkomstår, wyear=När indexeringen ska börja, iyear=skattreglerna
   'Application.Volatile
    If year > wyear And kvoten < 1 Then
       pbb = IBB 'Ersätt pbb med IBB och korrigera gränserna till de gamla
       pbb = pbb * kvoten
    End If
    'Skapat av Tommy Lowén beräkning av grundavdraget
   Dim avdrag As Double
   Dim Extra As Double
   
   If inkomst < 0.423 * pbb Then
       avdrag = inkomst
    ElseIf inkomst <= 1.185 * pbb Then
       avdrag = 0.423 * pbb
    ElseIf inkomst <= 2.72 * pbb Then
       avdrag = 0.423 * pbb + 0.2 * (inkomst - 1.185 * pbb)
    ElseIf inkomst <= 3.11 * pbb Then
       avdrag = 0.73 * pbb
    ElseIf inkomst <= 7.48 * pbb Then
       avdrag = 0.73 * pbb - 0.1 * (inkomst - 3.11 * pbb)
    Else
       avdrag = 0.293 * pbb
    End If
     
   If avdrag > inkomst Then avdrag = inkomst
   If marginal = 0 Then avdrag = Int((avdrag + 99.99) / 100) * 100 'Avrundat uppåt till närmaste hundratal kr
   avdrag05 = avdrag
End Function


Function avdrag06(ByVal inkomst, Optional ByVal pbb = 41000, Optional ByVal marginal = 0, Optional ByVal bald = 64, _
Optional ByVal year = 2100, Optional ByVal wyear = 21000, Optional ByVal IBB = 0, Optional ByVal kvoten = 1, Optional ByVal Iyear = 0, _
Optional ByVal Xage = 66) As Double
  'Year - inkomstår, wyear=När indexeringen ska börja, iyear=skattreglerna
   'Application.Volatile
    If year > wyear And kvoten < 1 Then
       pbb = IBB 'Ersätt pbb med IBB och korrigera gränserna till de gamla
       pbb = pbb * kvoten
    End If
    'Skapat av Tommy Lowén beräkning av grundavdraget
   Dim avdrag As Double
   Dim Extra As Double
   
   If inkomst < 0.423 * pbb Then
       avdrag = inkomst
    ElseIf inkomst <= 0.99 * pbb Then
       avdrag = 0.423 * pbb
    ElseIf inkomst <= 2.72 * pbb Then
       avdrag = 0.423 * pbb + 0.2 * (inkomst - 0.99 * pbb)
    ElseIf inkomst <= 3.11 * pbb Then
       avdrag = 0.77 * pbb
    ElseIf inkomst <= 7.88 * pbb Then
       avdrag = 0.77 * pbb - 0.1 * (inkomst - 3.11 * pbb)
    Else
       avdrag = 0.293 * pbb
    End If
     
   If avdrag > inkomst Then avdrag = inkomst
   If marginal = 0 Then avdrag = Int((avdrag + 99.99) / 100) * 100 'Avrundat uppåt till närmaste hundratal kr
   avdrag06 = avdrag
End Function


Function avdrag09(ByVal inkomst, Optional ByVal pbb = 41000, Optional ByVal marginal = 0, Optional ByVal bald = 64, _
Optional ByVal year = 2100, Optional ByVal wyear = 21000, Optional ByVal IBB = 0, Optional ByVal kvoten = 1, Optional ByVal Iyear = 0, _
Optional ByVal Xage = 66) As Double
  'Year - inkomstår, wyear=När indexeringen ska börja, iyear=skattreglerna
   'Application.Volatile
    If year > wyear And kvoten < 1 Then
       pbb = IBB 'Ersätt pbb med IBB och korrigera gränserna till de gamla
       pbb = pbb * kvoten
    End If
    'Skapat av Tommy Lowén beräkning av grundavdraget
   Dim avdrag As Double
   Dim Extra As Double
   
   If inkomst < 0.423 * pbb Then
       avdrag = inkomst
    ElseIf inkomst <= 0.99 * pbb Then
       avdrag = 0.423 * pbb
    ElseIf inkomst <= 2.72 * pbb Then
       avdrag = 0.423 * pbb + 0.2 * (inkomst - 0.99 * pbb)
    ElseIf inkomst <= 3.11 * pbb Then
       avdrag = 0.77 * pbb
    ElseIf inkomst <= 7.88 * pbb Then
       avdrag = 0.77 * pbb - 0.1 * (inkomst - 3.11 * pbb)
    Else
       avdrag = 0.293 * pbb
    End If
  
   If inkomst <= 0.425 * pbb Then
       Extra = 0
   ElseIf inkomst <= 0.99 * pbb Then
        Extra = 0.425 * pbb
   ElseIf inkomst <= 2.72 * pbb Then
        Extra = 0.623 * pbb - 0.2 * inkomst
   ElseIf inkomst <= 2.94 * pbb Then
        Extra = 0.078 * pbb
   ElseIf inkomst <= 3.11 * pbb Then
        Extra = 0.372 * pbb - 0.1 * inkomst
   ElseIf inkomst <= 7.88 * pbb Then
        Extra = 0.061 * pbb
   ElseIf inkomst <= 8.49 * pbb Then
        Extra = 0.849 * pbb - 0.1 * inkomst
   Else
       Extra = 0 * pbb
   End If
   If bald >= 66 Then avdrag = Extra + avdrag 'vid ingången av året
   
   If avdrag > inkomst Then avdrag = inkomst
   If marginal = 0 Then avdrag = Int((avdrag + 99.99) / 100) * 100 'Avrundat uppåt till närmaste hundratal kr
   avdrag09 = avdrag
End Function

Function avdrag10(ByVal inkomst, Optional ByVal pbb = 41000, Optional ByVal marginal = 0, Optional ByVal bald = 64, _
Optional ByVal year = 2100, Optional ByVal wyear = 21000, Optional ByVal IBB = 0, Optional ByVal kvoten = 1, Optional ByVal Iyear = 0, _
Optional ByVal Xage = 66) As Double
   'Year - inkomstår, wyear=När indexeringen ska börja, iyear=skattreglerna
   'Application.Volatile
    If year > wyear And kvoten < 1 Then
       pbb = IBB 'Ersätt pbb med IBB och korrigera gränserna till de gamla
       pbb = pbb * kvoten
    End If
    'Skapat av Tommy Lowén beräkning av grundavdraget 2010
    
   Dim avdrag As Double
   Dim Extra As Double
   
   If inkomst < 0.423 * pbb Then
       avdrag = inkomst
    ElseIf inkomst <= 0.99 * pbb Then
       avdrag = 0.423 * pbb
    ElseIf inkomst <= 2.72 * pbb Then
       avdrag = 0.225 * pbb + 0.2 * inkomst
    ElseIf inkomst <= 3.11 * pbb Then
       avdrag = 0.77 * pbb
    ElseIf inkomst <= 7.88 * pbb Then
       avdrag = 1.081 * pbb - 0.1 * inkomst
    Else
       avdrag = 0.293 * pbb
    End If
  
   If inkomst <= 0.423 * pbb Then
       Extra = 0
   ElseIf inkomst <= 0.99 * pbb Then
        Extra = 0.5094 * pbb
   ElseIf inkomst <= 2.72 * pbb Then
        Extra = 0.7074 * pbb - 0.2 * inkomst
   ElseIf inkomst <= 2.94 * pbb Then
        Extra = 0.1624 * pbb
   ElseIf inkomst <= 3.11 * pbb Then
        Extra = 0.1624 * pbb
   ElseIf inkomst <= 3.9 * pbb Then
        Extra = -0.1486 * pbb + 0.1 * inkomst
   ElseIf inkomst <= 7.88 * pbb Then
        Extra = 0.2219 * pbb + 0.005 * inkomst
   ElseIf inkomst <= 9.1568 * pbb Then
        Extra = 1.0099 * pbb - 0.095 * inkomst
   Else
       Extra = 0.14 * pbb
   End If
    
      
   If bald >= 66 Then avdrag = Extra + avdrag 'vid ingången av året
   
   If avdrag > inkomst Then avdrag = inkomst
   If marginal = 0 Then avdrag = Int((avdrag + 99.99) / 100) * 100 'Avrundat uppåt till närmaste hundratal kr
   avdrag10 = avdrag
End Function

Function avdrag11(ByVal inkomst, Optional ByVal pbb = 41000, Optional ByVal marginal = 0, Optional ByVal bald = 64, _
Optional ByVal year = 2100, Optional ByVal wyear = 21000, Optional ByVal IBB = 0, Optional ByVal kvoten = 1, Optional ByVal Iyear = 0, _
Optional ByVal Xage = 66) As Double
   
   'Application.Volatile
  'Year - inkomstår, wyear=När indexeringen ska börja, iyear=skattreglerna
 
    If year > wyear And kvoten < 1 Then
       pbb = IBB 'Ersätt pbb med IBB och korrigera gränserna till de gamla
       pbb = pbb * kvoten
    End If
   'Skapat av Tommy Lowén beräkning av grundavdraget 2011
   Dim avdrag As Double
   Dim Extra As Double
   
    If inkomst < 0.423 * pbb Then
       avdrag = inkomst
    ElseIf inkomst <= 0.99 * pbb Then
       avdrag = 0.423 * pbb
    ElseIf inkomst <= 2.72 * pbb Then
       avdrag = 0.225 * pbb + 0.2 * inkomst
    ElseIf inkomst <= 3.11 * pbb Then
       avdrag = 0.77 * pbb
    ElseIf inkomst <= 7.88 * pbb Then
       avdrag = 1.081 * pbb - 0.1 * inkomst
    Else
       avdrag = 0.293 * pbb
    End If
    
  'gränserna för äldre
   If inkomst <= 0.98 * pbb Then
        Extra = 0.557 * pbb
   ElseIf inkomst <= 0.99 * pbb Then
        Extra = 0.459 * pbb + 0.1 * inkomst
   ElseIf inkomst <= 2.72 * pbb Then
        Extra = 0.657 * pbb - 0.1 * inkomst
   ElseIf inkomst <= 3.11 * pbb Then
        Extra = 0.112 * pbb + 0.1 * inkomst
   ElseIf inkomst <= 3.85 * pbb Then
        Extra = -0.199 * pbb + 0.2 * inkomst
   ElseIf inkomst <= 4.8 * pbb Then
        Extra = 0.186 * pbb + 0.1 * inkomst
   ElseIf inkomst <= 7.88 * pbb Then
        Extra = 0.619 * pbb + 0.01 * inkomst
   ElseIf inkomst <= 12.21 * pbb Then
        Extra = 1.407 * pbb - 0.09 * inkomst
   Else
       Extra = 0.307 * pbb
   End If
      
   If bald >= 66 Then avdrag = avdrag + Extra ' Vid ingången av året, OBS född den 1/1 är inte 65 år. Kanse se om 66 år eller äldre vid utgången...
      
   If avdrag > inkomst Then avdrag = inkomst
   If marginal = 0 Then avdrag = Int((avdrag + 99.99) / 100) * 100 'Avrundat uppåt till närmaste hundratal kr
   avdrag11 = avdrag
End Function

Function avdrag13(ByVal inkomst, Optional ByVal pbb = 41000, Optional ByVal marginal = 0, Optional ByVal bald = 64, _
Optional ByVal year = 2100, Optional ByVal wyear = 21000, Optional ByVal IBB = 0, Optional ByVal kvoten = 1, Optional ByVal Iyear = 0, _
Optional ByVal Xage = 66) As Double
    'Year - inkomstår, wyear=När indexeringen ska börja, iyear=skattreglerna
  'Application.Volatile
     If year > wyear And kvoten < 1 Then
       pbb = IBB 'Ersätt pbb med IBB och korrigera gränserna till de gamla
       pbb = pbb * kvoten
    End If

   'Skapat av Tommy Lowén beräkning av grundavdraget 2013,
   Dim avdrag As Double
   Dim Extra As Double
   
    If inkomst < 0.423 * pbb Then
       avdrag = inkomst
    ElseIf inkomst <= 0.99 * pbb Then
       avdrag = 0.423 * pbb
    ElseIf inkomst <= 2.72 * pbb Then
       avdrag = 0.225 * pbb + 0.2 * inkomst
    ElseIf inkomst <= 3.11 * pbb Then
       avdrag = 0.77 * pbb
    ElseIf inkomst <= 7.88 * pbb Then
       avdrag = 1.081 * pbb - 0.1 * inkomst
    Else
       avdrag = 0.293 * pbb
    End If
    
  'gränserna för äldre
  Dim korr As Double
  korr = 0 'Fix för att på marginalen få ihop kurvan
   If inkomst <= 0.99 * pbb Then
        Extra = 0.567 * pbb
        korr = 0   '0,567 NEQ 0,785-0,2*0,99=0,587
   ElseIf inkomst <= 1.01 * pbb Then
        Extra = 0.785 * pbb - 0.2 * inkomst
        korr = 0     '0,785-0,2*1,01=0,583  Approx 0,674-0,09*1,01=0,5831
   ElseIf inkomst <= 2.72 * pbb Then
        Extra = 0.674 * pbb - 0.09 * inkomst
        korr = 0 '    '0,674-0,09*2,72=0,4292 EQ 0,129+0,11*2,72=0,4282
   ElseIf inkomst <= 3.11 * pbb Then
        Extra = 0.129 * pbb + 0.11 * inkomst
        korr = 0              '0,129+0,11*3,11=0,4711 EQ -0,182+0,21*3,11=0,4711
   ElseIf inkomst <= 3.75 * pbb Then
        Extra = -0.182 * pbb + 0.21 * inkomst
        korr = 0    '-0,182+0,21*3,75=0,6055 NEQ 0,233+0,1*3,75=0,6080
   ElseIf inkomst <= 4.77 * pbb Then
        Extra = 0.233 * pbb + 0.1 * inkomst
        korr = 0    '0,233+0,1*4,77=0,7100 NEQ 0.66 + 0.01*4,77 =0,7077
   ElseIf inkomst <= 7.88 * pbb Then
        Extra = 0.66 * pbb + 0.01 * inkomst
        korr = 0               '0.66 + 0.01*7,88=0,7388 EQ 1,448-0,09*7,88=0,7388
   ElseIf inkomst <= 12.14 * pbb Then
        Extra = 1.448 * pbb - 0.09 * inkomst
        korr = 0 '1,448-0,09*12,14=0,3554 NEQ 0,357
   Else
       Extra = 0.357 * pbb
       korr = 0
  End If
      
   If bald >= 66 Then avdrag = avdrag + Extra + korr ' Vid ingången av året, OBS född den 1/1 är inte 65 år. Kanse se om 66 år eller äldre vid utgången...
      
   If avdrag > inkomst Then avdrag = inkomst
   If marginal = 0 Then avdrag = Int((avdrag + 99.99) / 100) * 100 'Avrundat uppåt till närmaste hundratal kr
   
   avdrag13 = avdrag
End Function


Function avdrag14(ByVal inkomst, Optional ByVal pbb = 41000, Optional ByVal marginal = 0, Optional ByVal bald = 64, _
Optional ByVal year = 2100, Optional ByVal wyear = 21000, Optional ByVal IBB = 0, Optional ByVal kvoten = 1, Optional ByVal Iyear = 0, _
Optional ByVal Xage = 66) As Double
    'Year - inkomstår, wyear=När indexeringen ska börja, iyear=skattreglerna
  'Application.Volatile
     If year > wyear And kvoten < 1 Then
       pbb = IBB 'Ersätt pbb med IBB och korrigera gränserna till de gamla
       pbb = pbb * kvoten
    End If

   'Skapat av Tommy Lowén beräkning av grundavdraget 2014
   Dim avdrag As Double
   Dim Extra As Double
   
    If inkomst < 0.423 * pbb Then
       avdrag = inkomst
    ElseIf inkomst <= 0.99 * pbb Then
       avdrag = 0.423 * pbb
    ElseIf inkomst <= 2.72 * pbb Then
       avdrag = 0.225 * pbb + 0.2 * inkomst
    ElseIf inkomst <= 3.11 * pbb Then
       avdrag = 0.77 * pbb
    ElseIf inkomst <= 7.88 * pbb Then
       avdrag = 1.081 * pbb - 0.1 * inkomst
    Else
       avdrag = 0.293 * pbb
    End If
    
  'gränserna för äldre
   If inkomst <= 0.99 * pbb Then
        Extra = 0.682 * pbb
   ElseIf inkomst <= 1.105 * pbb Then
        Extra = 0.88 * pbb - 0.2 * inkomst
   ElseIf inkomst <= 2.72 * pbb Then
        Extra = 0.753 * pbb - 0.085 * inkomst
   ElseIf inkomst <= 3.11 * pbb Then
        Extra = 0.208 * pbb + 0.115 * inkomst
   ElseIf inkomst <= 3.69 * pbb Then
        Extra = -0.103 * pbb + 0.215 * inkomst
   ElseIf inkomst <= 4.785 * pbb Then
        Extra = 0.322 * pbb + 0.1 * inkomst
   ElseIf inkomst <= 7.88 * pbb Then
        Extra = 0.753 * pbb + 0.01 * inkomst
   ElseIf inkomst <= 12.43 * pbb Then
        Extra = 1.541 * pbb - 0.09 * inkomst
   Else
       Extra = 0.422 * pbb
  End If
      
   If bald >= 66 Then avdrag = avdrag + Extra ' Vid ingången av året, OBS född den 1/1 är inte 65 år. Kanse se om 66 år eller äldre vid utgången...
      
   If avdrag > inkomst Then avdrag = inkomst
   If marginal = 0 Then avdrag = Int((avdrag + 99.99) / 100) * 100 'Avrundat uppåt till närmaste hundratal kr
   
   avdrag14 = avdrag
End Function

Function avdrag16(ByVal inkomst, Optional ByVal pbb = 41000, Optional ByVal marginal = 0, Optional ByVal bald = 64, _
Optional ByVal year = 2100, Optional ByVal wyear = 21000, Optional ByVal IBB = 0, Optional ByVal kvoten = 1, Optional ByVal Iyear = 0, _
Optional ByVal Xage = 66) As Double
   'Year - inkomstår, wyear=När indexeringen ska börja, iyear=skattreglerna
  'Application.Volatile
     If year > wyear And kvoten < 1 Then
       pbb = IBB 'Ersätt pbb med IBB och korrigera gränserna till de gamla
       pbb = pbb * kvoten
    End If

   'Skapat av Tommy Lowén beräkning av grundavdraget 2016
   Dim avdrag As Double
   Dim Extra As Double
   
    If inkomst < 0.423 * pbb Then
       avdrag = inkomst
    ElseIf inkomst <= 0.99 * pbb Then
       avdrag = 0.423 * pbb
    ElseIf inkomst <= 2.72 * pbb Then
       avdrag = 0.225 * pbb + 0.2 * inkomst
    ElseIf inkomst <= 3.11 * pbb Then
       avdrag = 0.77 * pbb
    ElseIf inkomst <= 7.88 * pbb Then
       avdrag = 1.081 * pbb - 0.1 * inkomst
    Else
       avdrag = 0.293 * pbb
    End If
    
  'gränserna för äldre
   If inkomst <= 0.99 * pbb Then
        Extra = 0.687 * pbb
   ElseIf inkomst <= 1.11 * pbb Then
        Extra = 0.885 * pbb - 0.2 * inkomst
   ElseIf inkomst <= 2.72 * pbb Then
        Extra = 0.609 * pbb + 0.049 * inkomst
   ElseIf inkomst <= 3.11 * pbb Then
        Extra = 0.741 * pbb
   ElseIf inkomst <= 3.77 * pbb Then
        Extra = 0.43 * pbb + 0.1 * inkomst
   ElseIf inkomst <= 5.4 * pbb Then
        Extra = 0.807 * pbb
   ElseIf inkomst <= 7.88 * pbb Then
        Extra = 0.753 * pbb + 0.01 * inkomst
   ElseIf inkomst <= 12.43 * pbb Then
        Extra = 1.541 * pbb - 0.09 * inkomst
   Else
       Extra = 0.422 * pbb
  End If
      
   If bald >= 66 Then avdrag = avdrag + Extra ' Vid ingången av året, OBS född den 1/1 är inte 65 år. Kanse se om 66 år eller äldre vid utgången...
      
   If avdrag > inkomst Then avdrag = inkomst
   If marginal = 0 Then avdrag = Int((avdrag + 99.99) / 100) * 100 'Avrundat uppåt till närmaste hundratal kr
   
   avdrag16 = avdrag
End Function

Function avdrag18(ByVal inkomst, Optional ByVal pbb = 45500, Optional ByVal marginal = 0, Optional ByVal bald = 64, _
Optional ByVal year = 2100, Optional ByVal wyear = 21000, Optional ByVal IBB = 0, Optional ByVal kvoten = 1, Optional ByVal Iyear = 0, _
Optional ByVal Xage = 66) As Double
       'Year - inkomstår, wyear=När indexeringen ska börja, iyear=skattreglerna
      'Application.Volatile
       If marginal = 0 Then inkomst = Int(inkomst / 100) * 100 'taxerad inkomst
       
        If year > wyear And kvoten < 1 Then
           pbb = IBB 'Ersätt pbb med IBB och korrigera gränserna till de gamla
           pbb = pbb * kvoten
        End If
    
       'BP:n förslag för 2018, se sid
       Dim avdrag As Double
       Dim Extra As Double
       
    
        If inkomst <= 0.99 * pbb Then
           avdrag = 0.423 * pbb
        ElseIf inkomst <= 2.72 * pbb Then
           avdrag = 0.423 * pbb + 0.2 * (inkomst - 0.99 * pbb)
        ElseIf inkomst <= 3.11 * pbb Then
           avdrag = 0.77 * pbb
        ElseIf inkomst <= 7.88 * pbb Then
           avdrag = 0.77 * pbb - 0.1 * (inkomst - 3.11 * pbb)
        Else
           avdrag = 0.293 * pbb
        End If
        
       If inkomst <= 0.99 * pbb Then
            Extra = 0.687 * pbb
       ElseIf inkomst <= 1.11 * pbb Then
            Extra = 0.885 * pbb - 0.2 * inkomst
       ElseIf inkomst <= 2.72 * pbb Then
            Extra = 0.609 * pbb + 0.049 * inkomst
       ElseIf inkomst <= 2.94 * pbb Then
            Extra = -0.162 * pbb + 0.332 * inkomst
       ElseIf inkomst <= 3.11 * pbb Then
            Extra = 0.482 * pbb + 0.113 * inkomst
       ElseIf inkomst <= 4.45 * pbb Then
            Extra = 0.171 * pbb + 0.213 * inkomst
       ElseIf inkomst <= 7.88 * pbb Then
            Extra = 1.376 * pbb - 0.058 * inkomst
       ElseIf inkomst <= 9.15 * pbb Then
            Extra = 2.164 * pbb - 0.158 * inkomst
       ElseIf inkomst <= 12.43 * pbb Then
            Extra = 1.541 * pbb - 0.09 * inkomst
       Else
           Extra = 0.422 * pbb
    End If
          
    If bald > 65 Then avdrag = avdrag + Extra
       
    If avdrag > inkomst Then avdrag = inkomst
    If marginal = 0 Then avdrag = Int((avdrag + 99.99) / 100) * 100 'Avrundat uppåt till närmaste hundratal kr
    
    avdrag18 = avdrag
End Function

Function avdrag19(ByVal inkomst, Optional ByVal pbb = 46500, Optional ByVal marginal = 0, Optional ByVal bald = 64, _
Optional ByVal year = 2100, Optional ByVal wyear = 21000, Optional ByVal IBB = 0, Optional ByVal kvoten = 1, Optional ByVal Iyear = 0, _
Optional ByVal Xage = 66) As Double
      'Year - inkomstår, wyear=När indexeringen ska börja, iyear=skattreglerna
      'Application.Volatile
       If marginal = 0 Then inkomst = Int(inkomst / 100) * 100 'taxerad inkomst
       
         If year > wyear And kvoten < 1 Then
           pbb = IBB 'Ersätt pbb med IBB och korrigera gränserna till de gamla
           pbb = pbb * kvoten
        End If
    
       'BP:n förslag för 2018, se sid
       Dim avdrag As Double
       Dim Extra As Double
      
        If inkomst <= 0.99 * pbb Then
           avdrag = 0.423 * pbb
        ElseIf inkomst <= 2.72 * pbb Then
           avdrag = 0.423 * pbb + 0.2 * (inkomst - 0.99 * pbb)
        ElseIf inkomst <= 3.11 * pbb Then
           avdrag = 0.77 * pbb
        ElseIf inkomst <= 7.88 * pbb Then
           avdrag = 0.77 * pbb - 0.1 * (inkomst - 3.11 * pbb)
        Else
           avdrag = 0.293 * pbb
        End If
        
       If inkomst <= 0.99 * pbb Then
            Extra = 0.687 * pbb
       ElseIf inkomst <= 1.11 * pbb Then
            Extra = 0.885 * pbb - 0.2 * inkomst
       ElseIf inkomst <= 2.72 * pbb Then
            Extra = 0.6 * pbb + 0.057 * inkomst
       'ElseIf inkomst <= 2.94 * pbb Then
       '     extra = -0.169 * pbb + 0.34 * inkomst
       ElseIf inkomst <= 3.11 * pbb Then
            Extra = -0.169 * pbb + 0.34 * inkomst
       ElseIf inkomst <= 3.21 * pbb Then
            Extra = -0.48 * pbb + 0.44 * inkomst
       ElseIf inkomst <= 4.45 * pbb Then
            Extra = 0.207 * pbb + 0.228 * inkomst
       ElseIf inkomst <= 5.31 * pbb Then
            Extra = 1.397 * pbb - 0.039 * inkomst
       ElseIf inkomst <= 7.88 * pbb Then
            Extra = 0.763 * pbb + 0.08 * inkomst
       ElseIf inkomst <= 8.08 * pbb Then
            Extra = 1.551 * pbb - 0.02 * inkomst
       ElseIf inkomst <= 13.54 * pbb Then
            Extra = 2.399 * pbb - 0.125 * inkomst
       ElseIf inkomst <= 34# * pbb Then
            Extra = 1.031 * pbb - 0.024 * inkomst
       Else
           Extra = 0.215 * pbb
    End If
    
   
    If bald > 65 Then avdrag = avdrag + Extra
    
    If avdrag > inkomst Then avdrag = inkomst
    If marginal = 0 Then avdrag = Int((avdrag + 99.99) / 100) * 100 'Avrundat uppåt till närmaste hundratal kr
    
    avdrag19 = avdrag
    
End Function

Function avdrag20(ByVal inkomst, Optional ByVal pbb = 47300, Optional ByVal marginal = 0, Optional ByVal bald = 64, _
Optional ByVal year = 2100, Optional ByVal wyear = 2100, Optional ByVal IBB = 0, Optional ByVal kvoten = 1, Optional ByVal Iyear = 0, _
Optional ByVal Xage = 66) As Double
      'Year - inkomstår, wyear=När indexeringen ska börja, iyear=skattreglerna
      'Application.Volatile
       If marginal = 0 Then inkomst = Int(inkomst / 100) * 100 'taxerad inkomst
       'Ändrat april 2020, tl. wyear obsolet
       If (year > Iyear And Iyear > 0) And kvoten < 1 Then
           pbb = IBB 'Ersätt pbb med IBB och korrigera gränserna till de gamla
           pbb = pbb * kvoten
       End If
    
       'BP:n förslag för 2018, se sid
       Dim avdrag As Double
       Dim Extra As Double
      
        If inkomst <= 0.99 * pbb Then
           avdrag = 0.423 * pbb
        ElseIf inkomst <= 2.72 * pbb Then
           avdrag = 0.423 * pbb + 0.2 * (inkomst - 0.99 * pbb)
        ElseIf inkomst <= 3.11 * pbb Then
           avdrag = 0.77 * pbb
        ElseIf inkomst <= 7.88 * pbb Then
           avdrag = 0.77 * pbb - 0.1 * (inkomst - 3.11 * pbb)
        Else
           avdrag = 0.293 * pbb
        End If
        
       If inkomst <= 0.99 * pbb Then
            Extra = 0.687 * pbb
       ElseIf inkomst <= 1.11 * pbb Then
            Extra = 0.885 * pbb - 0.2 * inkomst
       ElseIf inkomst <= 2.72 * pbb Then
            Extra = 0.6 * pbb + 0.057 * inkomst
       ElseIf inkomst <= 3.11 * pbb Then
            Extra = -0.169 * pbb + 0.34 * inkomst
       ElseIf inkomst <= 3.21 * pbb Then
            Extra = -0.48 * pbb + 0.44 * inkomst
       ElseIf inkomst <= 4.45 * pbb Then
            Extra = 0.207 * pbb + 0.228 * inkomst
       ElseIf inkomst <= 7.88 * pbb Then
            Extra = 0.488 * pbb + 0.165 * inkomst
       ElseIf inkomst <= 8.08 * pbb Then
            Extra = 1.276 * pbb + 0.065 * inkomst
       ElseIf inkomst <= 11.06 * pbb Then
            Extra = 2.205 * pbb - 0.05 * inkomst
       ElseIf inkomst <= 12.15 * pbb Then
            Extra = 7.182 * pbb - 0.5 * inkomst
       ElseIf inkomst <= 29.65 * pbb Then
            Extra = 1.654 * pbb - 0.045 * inkomst
       ElseIf inkomst <= 34# * pbb Then
            Extra = 1.031 * pbb - 0.024 * inkomst
       Else
           Extra = 0.215 * pbb
    End If
          
        
    If bald >= Xage Then avdrag = avdrag + Extra 'För att fånga skatteregler av höjd pensionsålder

       
    If avdrag > inkomst Then avdrag = inkomst
    If marginal = 0 Then avdrag = Int((avdrag + 99.99) / 100) * 100 'Avrundat uppåt till närmaste hundratal kr
    
    avdrag20 = avdrag
End Function

Function avdrag21(ByVal inkomst, Optional ByVal pbb = 47300, Optional ByVal marginal = 0, Optional ByVal bald = 64, _
    Optional ByVal year = 2100, Optional ByVal wyear = 2100, Optional ByVal IBB = 0, Optional ByVal kvoten = 1, Optional ByVal Iyear = 0, _
    Optional ByVal Xage = 66) As Double
      'Year - inkomstår, wyear=När indexeringen ska börja, iyear=skattreglerna
      'Application.Volatile
       If marginal = 0 Then inkomst = Int(inkomst / 100) * 100 'taxerad inkomst
       'Ändrat april 2020, tl. wyear obsolet
       If (year > Iyear And Iyear > 0) And kvoten < 1 Then
           pbb = IBB 'Ersätt pbb med IBB och korrigera gränserna till de gamla
           pbb = pbb * kvoten
       End If
    
       Dim avdrag As Double
       Dim Extra As Double
      
        If inkomst <= 0.99 * pbb Then
           avdrag = 0.423 * pbb
        ElseIf inkomst <= 2.72 * pbb Then
           avdrag = 0.423 * pbb + 0.2 * (inkomst - 0.99 * pbb)
        ElseIf inkomst <= 3.11 * pbb Then
           avdrag = 0.77 * pbb
        ElseIf inkomst <= 7.88 * pbb Then
           avdrag = 0.77 * pbb - 0.1 * (inkomst - 3.11 * pbb)
        Else
           avdrag = 0.293 * pbb
        End If
        
       If inkomst <= 0.99 * pbb Then
            Extra = 0.687 * pbb
       ElseIf inkomst <= 1.11 * pbb Then
            Extra = 0.885 * pbb - 0.2 * inkomst
       ElseIf inkomst <= 2.72 * pbb Then
            Extra = 0.6 * pbb + 0.057 * inkomst
       ElseIf inkomst <= 3.11 * pbb Then
            Extra = -0.169 * pbb + 0.34 * inkomst
       ElseIf inkomst <= 3.21 * pbb Then
            Extra = -0.48 * pbb + 0.44 * inkomst
       ElseIf inkomst <= 7.88 * pbb Then
            Extra = 0.207 * pbb + 0.228 * inkomst
       ElseIf inkomst <= 8.08 * pbb Then
            Extra = 0.995 * pbb + 0.128 * inkomst
       ElseIf inkomst <= 11.28 * pbb Then
            Extra = 2.029 * pbb
       ElseIf inkomst <= 12.53 * pbb Then
            Extra = 9.023 * pbb - 0.62 * inkomst
       ElseIf inkomst <= 13.54 * pbb Then
            Extra = 1.253 * pbb
       ElseIf inkomst <= 35.36 * pbb Then
            Extra = 2.03 * pbb - 0.0574 * inkomst
       Else
           Extra = 0 '0.215 * PBB
    End If
          
        
    If bald >= Xage Then avdrag = avdrag + Extra 'För att fånga skatteregler av höjd pensionsålder

       
    If avdrag > inkomst Then avdrag = inkomst
    If marginal = 0 Then avdrag = Int((avdrag + 99.99) / 100) * 100 'Avrundat uppåt till närmaste hundratal kr
    
    avdrag21 = avdrag
End Function

Function avdrag22(ByVal inkomst, Optional ByVal pbb = 47300, Optional ByVal marginal = 0, Optional ByVal bald = 64, _
    Optional ByVal year = 2100, Optional ByVal wyear = 2100, Optional ByVal IBB = 0, Optional ByVal kvoten = 1, Optional ByVal Iyear = 0, _
    Optional ByVal Xage = 66) As Double
      'Year - inkomstår, wyear=När indexeringen ska börja, iyear=skattreglerna
      'Application.Volatile
       If marginal = 0 Then inkomst = Int(inkomst / 100) * 100 'taxerad inkomst
       If (year > Iyear And Iyear > 0) And kvoten < 1 Then
           pbb = IBB 'Ersätt pbb med IBB och korrigera gränserna till de gamla
           pbb = pbb * kvoten
       End If
    
       Dim avdrag As Double
       Dim Extra As Double
      
        If inkomst <= 0.99 * pbb Then
           avdrag = 0.423 * pbb
        ElseIf inkomst <= 2.72 * pbb Then
           avdrag = 0.423 * pbb + 0.2 * (inkomst - 0.99 * pbb)
        ElseIf inkomst <= 3.11 * pbb Then
           avdrag = 0.77 * pbb
        ElseIf inkomst <= 7.88 * pbb Then
           avdrag = 0.77 * pbb - 0.1 * (inkomst - 3.11 * pbb)
        Else
           avdrag = 0.293 * pbb
        End If
        
       If inkomst <= 0.91 * pbb Then
            Extra = 0.687 * pbb
       ElseIf inkomst <= 1.11 * pbb Then
            Extra = 0.885 * pbb - 0.2 * inkomst
       ElseIf inkomst <= 1.965 * pbb Then
            Extra = 0.6 * pbb + 0.057 * inkomst
       ElseIf inkomst <= 2.72 * pbb Then
            Extra = 0.333 * pbb + 0.1949 * inkomst
       ElseIf inkomst <= 3.11 * pbb Then
            Extra = -0.212 * pbb + 0.3949 * inkomst
       ElseIf inkomst <= 3.24 * pbb Then
            Extra = -0.523 * pbb + 0.4949 * inkomst
       ElseIf inkomst <= 5.53 * pbb Then
            Extra = 0.325 * pbb + 0.233 * inkomst
       ElseIf inkomst <= 7.88 * pbb Then
            Extra = 0.441 * pbb + 0.212 * inkomst
       ElseIf inkomst <= 8.08 * pbb Then
            Extra = 1.104 * pbb + 0.128 * inkomst
       ElseIf inkomst <= 11.48 * pbb Then
            Extra = 2.139 * pbb
       ElseIf inkomst <= 12.8 * pbb Then
            Extra = 9.257 * pbb - 0.62 * inkomst
       ElseIf inkomst <= 13.54 * pbb Then
            Extra = 1.32 * pbb
       ElseIf inkomst <= 36.54 * pbb Then
            Extra = 2.097 * pbb - 0.0574 * inkomst
       Else
           Extra = 0 '0.215 * PBB
    End If
          
    If bald >= Xage Then avdrag = avdrag + Extra 'För att fånga skatteregler av höjd pensionsålder
    
    If avdrag > inkomst Then avdrag = inkomst
    If marginal = 0 Then avdrag = Int((avdrag + 99.99) / 100) * 100 'Avrundat uppåt till närmaste hundratal kr
    
    avdrag22 = avdrag
End Function

Function avdrag24(ByVal inkomst, Optional ByVal pbb = 47300, Optional ByVal marginal = 0, Optional ByVal bald = 64, _
    Optional ByVal year = 2100, Optional ByVal wyear = 2100, Optional ByVal IBB = 0, Optional ByVal kvoten = 1, Optional ByVal Iyear = 0, _
    Optional ByVal Xage = 66) As Double
      'Year - inkomstår, wyear=När indexeringen ska börja, iyear=skattreglerna
      'Application.Volatile
       If marginal = 0 Then inkomst = Int(inkomst / 100) * 100 'taxerad inkomst
       If (year > Iyear And Iyear > 0) And kvoten < 1 Then
           pbb = IBB 'Ersätt pbb med IBB och korrigera gränserna till de gamla
           pbb = pbb * kvoten
       End If
    
       Dim avdrag As Double
       Dim Extra As Double
      
        If inkomst <= 0.99 * pbb Then
           avdrag = 0.423 * pbb
        ElseIf inkomst <= 2.72 * pbb Then
           avdrag = 0.225 * pbb + 0.2 * inkomst
        ElseIf inkomst <= 3.11 * pbb Then
           avdrag = 0.77 * pbb
        ElseIf inkomst <= 7.88 * pbb Then
           avdrag = 1.081 * pbb - 0.1 * inkomst
        Else
           avdrag = 0.293 * pbb
        End If
       
       If inkomst < 0.91 * pbb Then
            Extra = 0.687 * pbb
       ElseIf inkomst < 1.11 * pbb Then
            Extra = 0.885 * pbb - 0.2 * inkomst
       ElseIf inkomst < 1.965 * pbb Then
            Extra = 0.6 * pbb + 0.057 * inkomst
       ElseIf inkomst < 2.72 * pbb Then
            Extra = 0.333 * pbb + 0.1949 * inkomst
       ElseIf inkomst < 3.11 * pbb Then
            Extra = -0.212 * pbb + 0.3949 * inkomst
       ElseIf inkomst < 3.24 * pbb Then
            Extra = -0.523 * pbb + 0.4949 * inkomst
       ElseIf inkomst < 5# * pbb Then
            Extra = 0.208 * pbb + 0.2693 * inkomst
       ElseIf inkomst < 7.88 * pbb Then
            Extra = 0.3 * pbb + 0.2513 * inkomst
       ElseIf inkomst < 8.08 * pbb Then
            Extra = 0.986 * pbb + 0.1643 * inkomst
       ElseIf inkomst < 10.74 * pbb Then
            Extra = 2.313 * pbb
       ElseIf inkomst < 12.16 * pbb Then
            Extra = 8.972 * pbb - 0.62 * inkomst
       ElseIf inkomst < 13.54 * pbb Then
            Extra = 1.43 * pbb
       ElseIf inkomst < 38.42 * pbb Then
            Extra = 2.206 * pbb - 0.0574 * inkomst
       Else
           Extra = 0 '0.293 * pbb
    End If
          
    If bald >= Xage Then avdrag = avdrag + Extra 'För att fånga skatteregler av höjd pensionsålder
    
    If avdrag > inkomst Then avdrag = inkomst
    If marginal = 0 Then avdrag = Int((avdrag + 99.99) / 100) * 100 'Avrundat uppåt till närmaste hundratal kr
    
    avdrag24 = avdrag
End Function
Function avdrag25(ByVal inkomst, Optional ByVal pbb = 47300, Optional ByVal marginal = 0, Optional ByVal bald = 64, _
    Optional ByVal year = 2100, Optional ByVal wyear = 2100, Optional ByVal IBB = 0, Optional ByVal kvoten = 1, Optional ByVal Iyear = 0, _
    Optional ByVal Xage = 66) As Double
      'Year - inkomstår, wyear=När indexeringen ska börja, iyear=skattreglerna
      'Application.Volatile
       If marginal = 0 Then inkomst = Int(inkomst / 100) * 100 'taxerad inkomst
       If (year > Iyear And Iyear > 0) And kvoten < 1 Then
           pbb = IBB 'Ersätt pbb med IBB och korrigera gränserna till de gamla
           pbb = pbb * kvoten
       End If
    
       Dim avdrag As Double
       Dim Extra As Double
      
        If inkomst <= 0.99 * pbb Then
           avdrag = 0.423 * pbb
        ElseIf inkomst <= 2.72 * pbb Then
           avdrag = 0.423 * pbb + 0.2 * (inkomst - 0.99 * pbb)
        ElseIf inkomst <= 3.11 * pbb Then
           avdrag = 0.77 * pbb
        ElseIf inkomst <= 7.88 * pbb Then
           avdrag = 0.77 * pbb - 0.1 * (inkomst - 3.11 * pbb)
        Else
           avdrag = 0.293 * pbb
        End If
       
       If inkomst < 0.91 * pbb Then
            Extra = 0.687 * pbb
       ElseIf inkomst < 1.11 * pbb Then
            Extra = 0.885 * pbb - 0.2 * inkomst
       ElseIf inkomst < 1.965 * pbb Then
            Extra = 0.6 * pbb + 0.057 * inkomst
       ElseIf inkomst < 2.72 * pbb Then
            Extra = 0.333 * pbb + 0.1949 * inkomst
       ElseIf inkomst < 3.11 * pbb Then
            Extra = -0.212 * pbb + 0.3949 * inkomst
       ElseIf inkomst < 3.24 * pbb Then
            Extra = -0.523 * pbb + 0.4949 * inkomst
       ElseIf inkomst < 5# * pbb Then
            Extra = 0.096 * pbb + 0.304 * inkomst
       ElseIf inkomst < 7.88 * pbb Then
            Extra = 0.186 * pbb + 0.286 * inkomst
       ElseIf inkomst < 8.08 * pbb Then
            Extra = 0.872 * pbb + 0.199 * inkomst
       ElseIf inkomst < 10.94 * pbb Then
            Extra = 2.48 * pbb
       ElseIf inkomst < 12.47 * pbb Then
            Extra = 9.263 * pbb - 0.62 * inkomst
       Else
           Extra = 1.532 * pbb
    End If
          
    If bald >= Xage Then avdrag = avdrag + Extra 'För att fånga skatteregler av höjd pensionsålder
    
    If avdrag > inkomst Then avdrag = inkomst
    If marginal = 0 Then avdrag = Int((avdrag + 99.99) / 100) * 100 'Avrundat uppåt till närmaste hundratal kr
    
    avdrag25 = avdrag
End Function
Function avdrag26(ByVal inkomst, Optional ByVal pbb = 47300, Optional ByVal marginal = 0, Optional ByVal bald = 64, _
    Optional ByVal year = 2100, Optional ByVal wyear = 2100, Optional ByVal IBB = 0, Optional ByVal kvoten = 1, Optional ByVal Iyear = 0, _
    Optional ByVal Xage = 66) As Double
      'Year - inkomstår, wyear=När indexeringen ska börja, iyear=skattreglerna
      'Application.Volatile
       If marginal = 0 Then inkomst = Int(inkomst / 100) * 100 'taxerad inkomst
       If (year > Iyear And Iyear > 0) And kvoten < 1 Then
           pbb = IBB 'Ersätt pbb med IBB och korrigera gränserna till de gamla
           pbb = pbb * kvoten
       End If
    
       Dim avdrag As Double
       Dim Extra As Double
      
        If inkomst <= 0.99 * pbb Then
           avdrag = 0.423 * pbb
        ElseIf inkomst <= 2.72 * pbb Then
           avdrag = 0.423 * pbb + 0.2 * (inkomst - 0.99 * pbb)
        ElseIf inkomst <= 3.11 * pbb Then
           avdrag = 0.77 * pbb
        ElseIf inkomst <= 7.88 * pbb Then
           avdrag = 0.77 * pbb - 0.1 * (inkomst - 3.11 * pbb)
        Else
           avdrag = 0.293 * pbb
        End If
       
        If bald >= Xage Then 'För att fånga skatteregler av höjd pensionsålder
            If inkomst < 0.91 * pbb Then
                Extra = 0.687 * pbb
            ElseIf inkomst < 1.11 * pbb Then
                Extra = 0.885 * pbb - 0.2 * inkomst
            ElseIf inkomst < 1.965 * pbb Then
                Extra = 0.6 * pbb + 0.057 * inkomst
            ElseIf inkomst < 2.72 * pbb Then
                Extra = 0.333 * pbb + 0.1949 * inkomst
            ElseIf inkomst < 3.11 * pbb Then
                Extra = -0.212 * pbb + 0.3949 * inkomst
            ElseIf inkomst < 3.24 * pbb Then
                Extra = -0.523 * pbb + 0.4949 * inkomst
            ElseIf inkomst < 5# * pbb Then
                Extra = -0.073 * pbb + 0.356 * inkomst
            ElseIf inkomst < 7.88 * pbb Then
                Extra = 0.017 * pbb + 0.338 * inkomst
            ElseIf inkomst < 8.08 * pbb Then
                Extra = 0.703 * pbb + 0.251 * inkomst
            ElseIf inkomst < 11.16 * pbb Then
                Extra = 2.732 * pbb
            ElseIf inkomst < 12.84 * pbb Then
                Extra = 9.652 * pbb - 0.62 * inkomst
            Else
                Extra = 1.691 * pbb
            End If
            avdrag = avdrag + Extra
        End If
    
    If avdrag > inkomst Then avdrag = inkomst
    If marginal = 0 Then avdrag = Int((avdrag + 99.99) / 100) * 100 'Avrundat uppåt till närmaste hundratal kr
    
    avdrag26 = avdrag
End Function

''Sub kolltl()
''Dim a As Double
'' a = avdrag24(700000, 50000, 1, 67)
''
''End Sub

Function avdragxx(ByVal inkomst, Optional ByVal pbb = 52500, Optional ByVal marginal = 0, Optional ByVal bald = 64, _
Optional ByVal year = 2100, Optional ByVal wyear = 21000, Optional ByVal IBB = 0, Optional ByVal kvoten = 1, Optional ByVal Iyear = 0, _
Optional ByVal Xage = 66) As Double
   'Application.Volatile
    'Year - inkomstår, wyear=När indexeringen ska börja, iyear=skattreglerna
    
 If Iyear = 0 Then Iyear = year

 If Iyear < 1996 Then
    avdragxx = avdrag95(inkomst, pbb, marginal, bald, year, wyear, IBB, kvoten, Iyear)
 ElseIf Iyear < 2001 Then
    avdragxx = avdrag96(inkomst, pbb, marginal, bald, year, wyear, IBB, kvoten, Iyear)
 ElseIf Iyear < 2003 Then
    avdragxx = avdrag01(inkomst, pbb, marginal, bald, year, wyear, IBB, kvoten, Iyear)
 ElseIf Iyear < 2005 Then
    avdragxx = avdrag03(inkomst, pbb, marginal, bald, year, wyear, IBB, kvoten, Iyear)
 ElseIf Iyear = 2005 Then
    avdragxx = avdrag05(inkomst, pbb, marginal, bald, year, wyear, IBB, kvoten, Iyear)
 ElseIf Iyear < 2009 Then
    avdragxx = avdrag06(inkomst, pbb, marginal, bald, year, wyear, IBB, kvoten, Iyear)
 ElseIf Iyear < 2010 Then
    avdragxx = avdrag09(inkomst, pbb, marginal, bald, year, wyear, IBB, kvoten, Iyear)
 ElseIf Iyear < 2011 Then
    avdragxx = avdrag10(inkomst, pbb, marginal, bald, year, wyear, IBB, kvoten, Iyear)
 ElseIf Iyear < 2013 Then
    avdragxx = avdrag11(inkomst, pbb, marginal, bald, year, wyear, IBB, kvoten, Iyear)
 ElseIf Iyear < 2014 Then
    avdragxx = avdrag13(inkomst, pbb, marginal, bald, year, wyear, IBB, kvoten, Iyear)
 ElseIf Iyear < 2016 Then
    avdragxx = avdrag14(inkomst, pbb, marginal, bald, year, wyear, IBB, kvoten, Iyear)
 ElseIf Iyear < 2018 Then
    avdragxx = avdrag16(inkomst, pbb, marginal, bald, year, wyear, IBB, kvoten, Iyear)
 ElseIf Iyear < 2019 Then
    avdragxx = avdrag18(inkomst, pbb, marginal, bald, year, wyear, IBB, kvoten, Iyear)
 ElseIf Iyear < 2020 Then
    avdragxx = avdrag19(inkomst, pbb, marginal, bald, year, wyear, IBB, kvoten, Iyear)
 ElseIf Iyear < 2021 Then
    avdragxx = avdrag20(inkomst, pbb, marginal, bald, year, wyear, IBB, kvoten, Iyear, Xage)
 ElseIf Iyear < 2022 Then
    avdragxx = avdrag21(inkomst, pbb, marginal, bald, year, wyear, IBB, kvoten, Iyear, Xage)
 ElseIf Iyear <= 2023 Then
    avdragxx = avdrag22(inkomst, pbb, marginal, bald, year, wyear, IBB, kvoten, Iyear, Xage)
 ElseIf Iyear < 2025 Then
    avdragxx = avdrag24(inkomst, pbb, marginal, bald, year, wyear, IBB, kvoten, Iyear, Xage)
 ElseIf Iyear = 2025 Then
    avdragxx = avdrag25(inkomst, pbb, marginal, bald, year, wyear, IBB, kvoten, Iyear, Xage)
 ElseIf Iyear > 2025 Then
    avdragxx = avdrag26(inkomst, pbb, marginal, bald, year, wyear, IBB, kvoten, Iyear, Xage)
 Else
    avdragxx = avdrag26(inkomst, pbb, marginal, bald, year, wyear, IBB, kvoten, Iyear, Xage)
 End If
 
End Function

''Sub kollagavd()
''    'Call propp
''    Dim inkomst As Double
''    Dim pbb As Double
''    Dim marginal As Byte
''    Dim alder As Integer
''    Dim year As Integer
''    Dim avdrag As Double
''    inkomst = 90000
''    pbb = 52500
''    marginal = 1
''    alder = 66
''    year = 2022
''      'For year = 2016 To 2024
''       nypropp = Application.Range("rng_nypropp")
''        avdrag = avdragxx(inkomst, pbb, marginal, alder, year)
''        Debug.Print year; avdrag
''    'Next year
''End Sub

Function avdragRES(ByVal avdrag, Optional Typ = 1, Optional ByVal marginal = 0, Optional ByVal year = 2011) As Double
    'Typ 0 VBA börjar arrayen med 0
    'Typ 1 Resor
    'Typ 2 Privat pensionsparande
    'typ 3 Övriga avdrag
    'Year
     'Application.Volatile
     
    avdragRES = 0
    If avdrag = 0 Then Exit Function
    
    Dim TJP As Long
    TJP = 0
    
    If Application.Range("rng_TJP_Val").Value > 0 Then TJP = 1
    'Alternativt och bättre att se på årets inbetalda premie
    
     Dim max As Long 'Avseende de utan tjp
     max = 0.35 * Wage_(age)
     If max > 10 * pbb(age) Then max = 10 * pbb(age)
       
     Dim maxips As Long
     
     maxips = 12000 '2008-2014
     If year < 1994 Then Exit Function 'maxips = 0
     'If year < 2008 Then maxips ='?
     If year = 2015 Then maxips = 1800
     If year >= 2016 Then maxips = 0
      
     If Typ = 2 And year > 2015 And TJP > 0 Then Exit Function
     
     'Utan tjp
     If TJP = 0 Then maxips = maxips + max
     
    Dim ZLIMIT As Variant '0-4
    If year < 2007 Then
     ZLIMIT = Array(0, 7000, maxips, 1000)
    ElseIf year < 2009 Then
     ZLIMIT = Array(0, 8000, maxips, 1000)
    ElseIf year < 2010 Then
     ZLIMIT = Array(0, 8000, maxips, 1000)
    ElseIf year < 2010 Then
     ZLIMIT = Array(0, 8000, maxips, 1000)
    ElseIf year = 2010 Then
     ZLIMIT = Array(0, 8000, maxips, 1000)
    ElseIf year = 2011 Then
     ZLIMIT = Array(0, 9000, maxips, 1000)
    ElseIf year <= 2016 Then
     ZLIMIT = Array(0, 10000, maxips, 1000)
    Else
     ZLIMIT = Array(0, 11000, maxips, 1000)
    End If
    
     If Typ = 2 Then
        If avdrag > ZLIMIT(2) Then avdrag = ZLIMIT(2)
     ElseIf Typ < 4 Then
        avdrag = avdrag - ZLIMIT(Typ)
     End If
     If Typ = 4 Then avdrag = 0
     If avdrag < 0 Then avdrag = 0
     
     If marginal = 0 Then avdragRES = Int(avdrag)

End Function

''Sub koll_avdrag()
''   Dim kol_avd As Long
''   Dim typ As Integer
''   Dim marginal As Integer
''   Dim year As Long
''   Dim avdrag As Double
''   typ = 2
''   avdrag = 10000
''   marginal = 0
''   year = 2015
''   kol_avd = avdragRES(avdrag, typ, marginal, year)
''   Debug.Print kol_avd
''End Sub



Function statlig(ByVal Besk, Optional lim1 = 383000, Optional lim2 = 548300, Optional ByVal marginal = 0, Optional ByVal year = 0) As Double
    'Application.Volatile
    'Gränserna finns i arket "några tal"
    
     If Besk > lim2 Then
      statlig = (Besk - lim2) * 0.25 + (lim2 - lim1) * 0.2
    ElseIf Besk > lim1 Then
      statlig = (Besk - lim1) * 0.2
    Else
      statlig = 0
    End If
    If marginal = 0 Then statlig = Int(statlig)
End Function

Function PublicAvg(ByVal Besk, Optional lim1 = 0.01, Optional bald = 18, Optional ByVal marginal = 0, Optional year = 2019) As Double
    'Application.Volatile
    PublicAvg = 0 'Prop 2017/18:261
    If bald <= 18 Or Besk < 1 Or year < 2019 Then Exit Function
    Dim lim2 As Double
    
    lim2 = (2.092 * IBB(year - Int(born)))
    
    If year = 2021 Then lim2 = 1.95 * IBB(year - Int(born))
    If year = 2022 Then lim2 = 1.87 * IBB(year - Int(born))
    If year = 2023 Then lim2 = 1.75 * IBB(year - Int(born))
    If year = 2024 Then lim2 = 1.6 * IBB(year - Int(born))
    If year = 2025 Then lim2 = 1.55 * IBB(year - Int(born))
    If year >= 2026 Then lim2 = 1.42 * IBB(year - Int(born))
    
    If Besk > lim2 Then
            PublicAvg = lim2 * lim1
        Else
            PublicAvg = Besk * lim1
    End If
    
    If marginal = 0 Then PublicAvg = Int(PublicAvg + 0.5)
    
End Function

Function Jobb07(ByVal inkomst, Optional ByVal alder = 30, Optional ByVal ksats = 0.3144, Optional ByVal pbb = 42400, Optional ByVal marginal = 0, Optional ByVal binkomst = 0, _
  Optional ByVal year = 2100, Optional ByVal wyear = 21000, Optional ByVal IBB = 0, Optional ByVal kvoten = 1, Optional ByVal Iyear = 0) As Double
  
    'Application.Volatile
     'Year - inkomstår, wyear=När indexeringen ska börja, iyear=skattreglerna
    
    '   If year > wyear And Kvoten < 1 Then
    '      PBB = ibb 'Ersätt pbb med IBB och korrigera gränserna till de gamla
    '      PBB = PBB * Kvoten
    '   End If
    
    Dim avdrag As Double
    avdrag = avdragxx(inkomst + binkomst, pbb, marginal, alder, year, wyear, IBB, kvoten, Iyear)
    ' avdragxx = avdrag11(Inkomst, PBB, marginal, alder, year, wyear, ibb, Kvoten, iyear)
    Dim jobb As Long
   
   If inkomst <= 0.79 * pbb Then
       jobb = inkomst
    ElseIf inkomst <= 2.72 * pbb Then
       jobb = 0.2 * (inkomst - 0.79 * pbb) + 0.79 * pbb
    Else
       jobb = 1.176 * pbb
    End If
     
    If alder >= 66 Then
      If inkomst <= 1.59 * pbb Then
         jobb = inkomst
       ElseIf inkomst <= 2.72 * pbb Then
         jobb = 1.5 * pbb + 0.2 * (inkomst - 1.59 * pbb)
       Else
         jobb = 1.816 * pbb
       End If
     End If
    
    jobb = (jobb - avdrag) * ksats
     
     If jobb < 0 Then jobb = 0
     If marginal = 0 Then jobb = Int(jobb)
     Jobb07 = jobb
 End Function

Function Jobb08(ByVal inkomst, Optional ByVal alder = 30, Optional ByVal ksats = 0.3144, Optional ByVal pbb = 42400, Optional ByVal marginal = 0, Optional ByVal binkomst = 0, _
  Optional ByVal year = 2100, Optional ByVal wyear = 21000, Optional ByVal IBB = 0, Optional ByVal kvoten = 1, Optional ByVal Iyear = 0) As Double
    'Application.Volatile
    'Year - inkomstår, wyear=När indexeringen ska börja, iyear=skattreglerna
    
 '   If year > wyear And Kvoten < 1 Then
 '      PBB = ibb 'Ersätt pbb med IBB och korrigera gränserna till de gamla
 '      PBB = PBB * Kvoten
 '   End If
    
   Dim avdrag As Double
   avdrag = avdragxx(inkomst + binkomst, pbb, marginal, alder, year, wyear, IBB, kvoten, Iyear)
   ' avdragxx = avdrag11(Inkomst, PBB, marginal, alder, year, wyear, ibb, Kvoten, iyear)
   Dim jobb As Long
   
   If inkomst <= 0.91 * pbb Then
       jobb = inkomst
    ElseIf inkomst <= 2.72 * pbb Then
       jobb = 0.2 * (inkomst - 0.91 * pbb) + 0.91 * pbb
    ElseIf inkomst <= 7 * pbb Then
       jobb = 0.033 * (inkomst - 2.72 * pbb) + 1.272 * pbb
    Else
       jobb = 1.413 * pbb
    End If
     
    If alder >= 66 Then
      If inkomst <= 1.79 * pbb Then
         jobb = inkomst
       ElseIf inkomst <= 2.72 * pbb Then
         jobb = 1.79 * pbb + 0.2 * (inkomst - 1.79 * pbb)
       Else
         jobb = 2.117 * pbb
       End If
     End If
    
    jobb = (jobb - avdrag) * ksats
     
     If jobb < 0 Then jobb = 0
     If marginal = 0 Then jobb = Int(jobb)
     Jobb08 = jobb
 End Function

 Function Jobb10(ByVal inkomst, Optional ByVal alder = 30, Optional ByVal ksats = 0.3144, Optional ByVal pbb = 42400, Optional ByVal marginal = 0, Optional ByVal binkomst = 0, _
  Optional ByVal year = 2100, Optional ByVal wyear = 21000, Optional ByVal IBB = 0, Optional ByVal kvoten = 1, Optional ByVal Iyear = 0) As Double
    'Application.Volatile
   'Year - inkomstår, wyear=När indexeringen ska börja, iyear=skattreglerna
 
    '   If year > wyear And Kvoten < 1 Then
    '      PBB = ibb 'Ersätt pbb med IBB och korrigera gränserna till de gamla
    '      PBB = PBB * Kvoten
    '   End If
       
   Dim avdrag As Double
   avdrag = avdragxx(inkomst + binkomst, pbb, marginal, alder, year, wyear, IBB, kvoten, Iyear)
   ' avdragxx = avdrag11(Inkomst, PBB, marginal, alder, year, wyear, ibb, Kvoten, iyear)
   Dim jobb As Long
     
   If inkomst <= 0.91 * pbb Then
       jobb = inkomst
    ElseIf inkomst <= 2.72 * pbb Then
       jobb = 0.304 * (inkomst - 0.91 * pbb) + 0.91 * pbb
    ElseIf inkomst <= 7 * pbb Then
       jobb = 1.461 * pbb + 0.095 * (inkomst - 2.72 * pbb)
    Else
       jobb = 1.868 * pbb
    End If
     
    jobb = (jobb - avdrag) * ksats
           
    If alder >= 66 Then
      If inkomst <= 100000 Then
         jobb = 0.2 * inkomst
       ElseIf inkomst <= 300000 Then
         jobb = 15000 + 0.05 * inkomst
       Else
         jobb = 30000
       End If
     End If
     
     If jobb < 0 Then jobb = 0
     If marginal = 0 Then jobb = Int(jobb)
     Jobb10 = jobb
 End Function

 Function Jobb11(ByVal inkomst, Optional ByVal alder = 30, Optional ByVal ksats = 0.3144, Optional ByVal pbb = 42400, Optional ByVal marginal = 0, Optional ByVal binkomst = 0, _
  Optional ByVal year = 2100, Optional ByVal wyear = 21000, Optional ByVal IBB = 0, Optional ByVal kvoten = 1, Optional ByVal Iyear = 0) As Double
     'Application.Volatile
    'Year - inkomstår, wyear=När indexeringen ska börja, iyear=skattreglerna
 
    '   If year > wyear And Kvoten < 1 Then
    '     PBB = ibb 'Ersätt pbb med IBB och korrigera gränserna till de gamla
    '     PBB = PBB * Kvoten
    '  End If
  
    Dim avdrag As Double
    Dim jobb As Double
 
    avdrag = avdragxx(inkomst + binkomst, pbb, marginal, alder, year, wyear, IBB, kvoten, Iyear)
   
   If inkomst <= 0.91 * pbb Then
       jobb = inkomst
    ElseIf inkomst <= 2.72 * pbb Then
       jobb = 0.304 * (inkomst - 0.91 * pbb) + 0.91 * pbb
    ElseIf inkomst <= 7 * pbb Then
       jobb = 1.461 * pbb + 0.095 * (inkomst - 2.72 * pbb)
    Else
       jobb = 1.868 * pbb
    End If
     
    jobb = (jobb - avdrag) * ksats
           
    If alder >= 66 Then
      If inkomst <= 100000 Then
         jobb = 0.2 * inkomst
       ElseIf inkomst <= 300000 Then
         jobb = 15000 + 0.05 * inkomst
       Else
         jobb = 30000
       End If
     End If
     Jobb11 = jobb
     
     If Jobb11 < 0 Then Jobb11 = 0
     If marginal = 0 Then Jobb11 = Int(Jobb11)
     'Jobb11 = avdrag
 End Function
 
 Function Jobb14(ByVal inkomst, Optional ByVal alder = 30, Optional ByVal ksats = 0.3144, Optional ByVal pbb = 42400, Optional ByVal marginal = 0, Optional ByVal binkomst = 0, _
  Optional ByVal year = 2100, Optional ByVal wyear = 21000, Optional ByVal IBB = 0, Optional ByVal kvoten = 1, Optional ByVal Iyear = 0) As Double
      
      'Application.Volatile
     'Year - inkomstår, wyear=När indexeringen ska börja, iyear=skattreglerna
    
     '   If year > wyear And Kvoten < 1 Then
     '     PBB = ibb 'Ersätt pbb med IBB och korrigera gränserna till de gamla
     '     PBB = PBB * Kvoten
     '  End If
       
    Dim avdrag As Double
    Dim jobb As Double
 
    avdrag = avdragxx(inkomst + binkomst, pbb, marginal, alder, year, wyear, IBB, kvoten, Iyear)
   
   If inkomst <= 0.91 * pbb Then
       jobb = inkomst
    ElseIf inkomst <= 2.94 * pbb Then
       jobb = 0.332 * (inkomst - 0.91 * pbb) + 0.91 * pbb
    ElseIf inkomst <= 8.08 * pbb Then
       jobb = 1.584 * pbb + 0.111 * (inkomst - 2.94 * pbb)
    Else
       jobb = 2.155 * pbb
    End If
     
    jobb = (jobb - avdrag) * ksats
           
    If alder >= 66 Then
      If inkomst <= 100000 Then
         jobb = 0.2 * inkomst
       ElseIf inkomst <= 300000 Then
         jobb = 15000 + 0.05 * inkomst
       Else
         jobb = 30000
       End If
     End If
     Jobb14 = jobb
     
     If Jobb14 < 0 Then Jobb14 = 0
     If marginal = 0 Then Jobb14 = Int(Jobb14)
     
 End Function
 
 Function Jobb16(ByVal inkomst, Optional alder = 30, Optional ksats = 0.3144, Optional pbb = 42800, Optional marginal = 0, Optional binkomst = 0, _
 Optional year = 2100, Optional wyear = 21000, Optional IBB = 0, Optional kvoten = 1, Optional Iyear = 0) As Double
    'Application.Volatile
     'Year - inkomstår, wyear=När indexeringen ska börja, iyear=skattreglerna
    
     '   If year > wyear And Kvoten < 1 Then
     '     PBB = ibb 'Ersätt pbb med IBB och korrigera gränserna till de gamla
     '     PBB = PBB * Kvoten
     '  End If
       
    Dim avdrag As Double
    Dim jobb As Double
    
    avdrag = avdragxx(inkomst + binkomst, pbb, marginal, alder, year, wyear, IBB, kvoten, Iyear)
   
   If inkomst <= 0.91 * pbb Then
       jobb = inkomst
    ElseIf inkomst <= 2.94 * pbb Then
       jobb = 0.332 * (inkomst - 0.91 * pbb) + 0.91 * pbb
    ElseIf inkomst <= 8.08 * pbb Then
       jobb = 1.584 * pbb + 0.111 * (inkomst - 2.94 * pbb)
    Else
       jobb = 2.155 * pbb
    End If
    
    If inkomst <= 13.54 * pbb Then
        jobb = (jobb - avdrag) * ksats
    Else
       jobb = (jobb - avdrag) * ksats - 0.03 * (inkomst - 13.54 * pbb)
    End If
           
    If alder >= 66 Then
      If inkomst <= 100000 Then
         jobb = 0.2 * inkomst
       ElseIf inkomst <= 300000 Then
         jobb = 15000 + 0.05 * inkomst
       ElseIf inkomst <= 600000 Then
         jobb = 30000
       Else
         jobb = 30000 - 0.05 * (inkomst - 600000)
       End If
     End If
     Jobb16 = jobb
     
     If Jobb16 < 0 Then Jobb16 = 0
     If marginal = 0 Then Jobb16 = Int(Jobb16)
     
 End Function
 
 Function Jobb19(ByVal inkomst, Optional alder = 30, Optional ksats = 0.3144, Optional pbb = 46500, Optional marginal = 0, Optional binkomst = 0, _
 Optional year = 2100, Optional wyear = 21000, Optional IBB = 0, Optional kvoten = 1, Optional Iyear = 0, Optional Xage = 66) As Double
    'Application.Volatile
     'Year - inkomstår, wyear=När indexeringen ska börja, iyear=skattreglerna
     '   If year > wyear And Kvoten < 1 Then
     '     PBB = ibb 'Ersätt pbb med IBB och korrigera gränserna till de gamla
     '     PBB = PBB * Kvoten
     '  End If
       
    Dim avdrag As Double
    Dim jobb As Double
    
    avdrag = avdragxx(inkomst + binkomst, pbb, marginal, alder, year, wyear, IBB, kvoten, Iyear)
   
   If inkomst <= 0.91 * pbb Then
       jobb = inkomst
    ElseIf inkomst <= 3.24 * pbb Then
       jobb = 0.3405 * (inkomst - 0.91 * pbb) + 0.91 * pbb
    ElseIf inkomst <= 8.08 * pbb Then
       jobb = 1.703 * pbb + 0.128 * (inkomst - 3.24 * pbb)
    Else
       jobb = 2.323 * pbb
    End If
    
    If inkomst <= 13.54 * pbb Then
        jobb = (jobb - avdrag) * ksats
    Else
       jobb = (jobb - avdrag) * ksats - 0.03 * (inkomst - 13.54 * pbb)
    End If
           
    If alder >= Xage Then 'Om det i framtiden ska följa riktåldern
      If inkomst <= 100000 Then
         jobb = 0.2 * inkomst
       ElseIf inkomst <= 300000 Then
         jobb = 15000 + 0.05 * inkomst
       ElseIf inkomst <= 600000 Then
         jobb = 30000
       Else
         jobb = 30000 - 0.05 * (inkomst - 600000)
       End If
     End If
     
    
     If jobb < 0 Then jobb = 0
     If marginal = 0 Then jobb = Int(jobb)
     Jobb19 = jobb
     
 End Function
 
Function Jobb22(ByVal inkomst, Optional alder = 30, Optional ksats = 0.3144, Optional pbb = 46500, Optional marginal = 0, Optional binkomst = 0, _
 Optional year = 2100, Optional wyear = 21000, Optional IBB = 0, Optional kvoten = 1, Optional Iyear = 0, Optional Xage = 66) As Double
    'Application.Volatile
     'Year - inkomstår, wyear=När indexeringen ska börja, iyear=skattreglerna
     '   If year > wyear And Kvoten < 1 Then
     '     PBB = ibb 'Ersätt pbb med IBB och korrigera gränserna till de gamla
     '     PBB = PBB * Kvoten
     '  End If
       
    Dim avdrag As Double
    Dim jobb As Double
    
    avdrag = avdragxx(inkomst + binkomst, pbb, marginal, alder, year, wyear, IBB, kvoten, Iyear)
    
    If inkomst <= 0.91 * pbb Then
       jobb = inkomst
    ElseIf inkomst <= 3.24 * pbb Then
       jobb = 0.3874 * (inkomst - 0.91 * pbb) + 0.91 * pbb
    ElseIf inkomst <= 8.08 * pbb Then
       jobb = 1.812 * pbb + 0.128 * (inkomst - 3.24 * pbb)
    Else
       jobb = 2.432 * pbb
    End If
    
    If inkomst <= 13.54 * pbb Then
        jobb = (jobb - avdrag) * ksats
    Else
       jobb = (jobb - avdrag) * ksats - 0.03 * (inkomst - 13.54 * pbb)
    End If
           
    If alder >= Xage Then 'Om det i framtiden ska följa riktåldern
      If inkomst <= 100000 Then
         jobb = 0.2 * inkomst
       ElseIf inkomst <= 300000 Then
         jobb = 15000 + 0.05 * inkomst
       ElseIf inkomst <= 600000 Then
         jobb = 30000
       Else
         jobb = 30000 - 0.05 * (inkomst - 600000)
       End If
     End If
    
     If jobb < 0 Then jobb = 0
     If marginal = 0 Then jobb = Int(jobb)
     Jobb22 = jobb
     
 End Function
 
 Function Jobb23(ByVal inkomst, Optional alder = 30, Optional ksats = 0.3144, Optional pbb = 46500, Optional marginal = 0, Optional binkomst = 0, _
 Optional year = 2100, Optional wyear = 21000, Optional IBB = 0, Optional kvoten = 1, Optional Iyear = 0, Optional Xage = 66) As Double
    'Application.Volatile
     'Year - inkomstår, wyear=När indexeringen ska börja, iyear=skattreglerna
     '   If year > wyear And Kvoten < 1 Then
     '     PBB = ibb 'Ersätt pbb med IBB och korrigera gränserna till de gamla
     '     PBB = PBB * Kvoten
     '  End If
       
    Dim avdrag As Double
    Dim jobb As Double
    
    avdrag = avdragxx(inkomst + binkomst, pbb, marginal, alder, year, wyear, IBB, kvoten, Iyear)
    
    If inkomst <= 0.91 * pbb Then
       jobb = inkomst
    ElseIf inkomst <= 3.24 * pbb Then
       jobb = 0.3874 * (inkomst - 0.91 * pbb) + 0.91 * pbb
    ElseIf inkomst <= 8.08 * pbb Then
       jobb = 1.812 * pbb + 0.128 * (inkomst - 3.24 * pbb)
    Else
       jobb = 2.432 * pbb
    End If
    
    If inkomst <= 13.54 * pbb Then
        jobb = (jobb - avdrag) * ksats
    Else
       jobb = (jobb - avdrag) * ksats - 0.03 * (inkomst - 13.54 * pbb)
    End If
           
    If alder >= 66 Then 'Xage Om det i framtiden ska följa riktåldern
      If inkomst <= 100000 Then
         jobb = 0.22 * inkomst
       ElseIf inkomst <= 300000 Then
         jobb = 15000 + 0.07 * inkomst
       ElseIf inkomst <= 600000 Then
         jobb = 36000
       Else
         jobb = 36000 - 0.03 * (inkomst - 600000)
       End If
     End If
    
     If jobb < 0 Then jobb = 0
     If marginal = 0 Then jobb = Int(jobb)
     Jobb23 = jobb
     
 End Function
 
 Function Jobb24(ByVal inkomst, Optional alder = 30, Optional ksats = 0.3144, Optional pbb = 46500, Optional marginal = 0, Optional binkomst = 0, _
 Optional year = 2100, Optional wyear = 21000, Optional IBB = 0, Optional kvoten = 1, Optional Iyear = 0, Optional Xage = 66) As Double
    'Application.Volatile
     'Year - inkomstår, wyear=När indexeringen ska börja, iyear=skattreglerna
     '   If year > wyear And Kvoten < 1 Then
     '     PBB = ibb 'Ersätt pbb med IBB och korrigera gränserna till de gamla
     '     PBB = PBB * Kvoten
     '  End If
       
    Dim avdrag As Double
    Dim jobb As Double
    
    avdrag = avdragxx(inkomst + binkomst, pbb, marginal, alder, year, wyear, IBB, kvoten, Iyear)
    
    If inkomst <= 0.91 * pbb Then
       jobb = inkomst
    ElseIf inkomst <= 3.24 * pbb Then
       jobb = 0.3874 * (inkomst - 0.91 * pbb) + 0.91 * pbb
    ElseIf inkomst <= 8.08 * pbb Then
       jobb = 1.813 * pbb + 0.1643 * (inkomst - 3.24 * pbb)
    ElseIf inkomst <= 13.54 * pbb Then
       jobb = 2.608 * pbb
    Else
       jobb = 2.608 * pbb - 0.03 * (inkomst - 13.54 * pbb)
    End If
    
    jobb = (jobb - avdrag) * ksats

           'Vid ingången av året har fyllt 66 (dvs 67 under året) men inte 69 (dvs 70 vid utgången)
    If alder > Xage Then 'Xage Om det i framtiden ska följa riktåldern
      If inkomst < 1.75 * pbb Then
         jobb = 0.22 * inkomst
       ElseIf inkomst < 5.24 * pbb Then
         jobb = 0.2635 * pbb + 0.07 * inkomst
       ElseIf inkomst < 10.48 * pbb Then
         jobb = 0.6293 * pbb
       Else
         jobb = 0.6293 * pbb - 0.03 * (inkomst - 10.48 * pbb)
       End If
    End If
    ''If alder > 69 Then jobb = jobb * 1.25
    
    
    If jobb < 0 Then jobb = 0
    If marginal = 0 Then jobb = Int(jobb)
    
    Jobb24 = jobb
     
 End Function
Function Jobb25(ByVal inkomst, Optional alder = 30, Optional ksats = 0.3144, Optional pbb = 46500, Optional marginal = 0, Optional binkomst = 0, _
 Optional year = 2100, Optional wyear = 21000, Optional IBB = 0, Optional kvoten = 1, Optional Iyear = 0, Optional Xage = 66) As Double
    'Application.Volatile
     'Year - inkomstår, wyear=När indexeringen ska börja, iyear=skattreglerna
     '   If year > wyear And Kvoten < 1 Then
     '     PBB = ibb 'Ersätt pbb med IBB och korrigera gränserna till de gamla
     '     PBB = PBB * Kvoten
     '  End If
       
    Dim avdrag As Double
    Dim jobb As Double
    
    avdrag = avdragxx(inkomst + binkomst, pbb, marginal, alder, year, wyear, IBB, kvoten, Iyear)
    
    If inkomst <= 0.91 * pbb Then
       jobb = inkomst
    ElseIf inkomst <= 3.24 * pbb Then
       jobb = 0.3874 * (inkomst - 0.91 * pbb) + 0.91 * pbb
    ElseIf inkomst <= 8.08 * pbb Then
       jobb = 1.813 * pbb + 0.199 * (inkomst - 3.24 * pbb)
    Else
       jobb = 2.776 * pbb
    End If
    'Eftersom denna beräkning görs på samtliga "if" fall ovan har den brutits ut för att undvika upprepning
    jobb = (jobb - avdrag) * ksats

    'Vid ingången av året har fyllt 66 (dvs 67 under året) men inte 69 (dvs 70 vid utgången)
    If alder > Xage Then 'Xage Om det i framtiden ska följa riktåldern
      If inkomst < 1.7 * pbb Then
         jobb = 0.22 * inkomst
       ElseIf inkomst < 5.24 * pbb Then
         jobb = 0.2635 * pbb + 0.07 * inkomst
       Else
         jobb = 0.6293 * pbb
       End If
    End If
    
    If jobb < 0 Then jobb = 0
    If marginal = 0 Then jobb = Int(jobb)
    
    Jobb25 = jobb
     
 End Function
Function Jobb26(ByVal inkomst, Optional alder = 30, Optional ksats = 0.3144, Optional pbb = 46500, Optional marginal = 0, Optional binkomst = 0, _
 Optional year = 2100, Optional wyear = 21000, Optional IBB = 0, Optional kvoten = 1, Optional Iyear = 0, Optional Xage = 66) As Double
    'Application.Volatile
     'Year - inkomstår, wyear=När indexeringen ska börja, iyear=skattreglerna
     '   If year > wyear And Kvoten < 1 Then
     '     PBB = ibb 'Ersätt pbb med IBB och korrigera gränserna till de gamla
     '     PBB = PBB * Kvoten
     '  End If
       
    Dim avdrag As Double
    Dim jobb As Double
    
    avdrag = avdragxx(inkomst + binkomst, pbb, marginal, alder, year, wyear, IBB, kvoten, Iyear, Xage)
    
    If inkomst <= 0.91 * pbb Then
       jobb = inkomst
    ElseIf inkomst <= 3.24 * pbb Then
       jobb = 0.3874 * (inkomst - 0.91 * pbb) + 0.91 * pbb
    ElseIf inkomst <= 8.08 * pbb Then
       jobb = 1.813 * pbb + 0.251 * (inkomst - 3.24 * pbb)
    Else
       jobb = 3.027 * pbb
    End If
    'Eftersom denna beräkning görs på samtliga "if" fall ovan har den brutits ut för att undvika upprepning
    jobb = (jobb - avdrag) * ksats

    'Vid ingången av året har fyllt 66 (dvs 67 under året) men inte 69 (dvs 70 vid utgången)
    If alder > Xage Then 'Xage Om det i framtiden ska följa riktåldern
      If inkomst < 1.7 * pbb Then
         jobb = 0.22 * inkomst
       ElseIf inkomst < 6.5 * pbb Then
         jobb = 0.2635 * pbb + 0.07 * inkomst
       Else
         jobb = 0.6293 * pbb
       End If
    End If
    
    If jobb < 0 Then jobb = 0
    If marginal = 0 Then jobb = Int(jobb)
    
    Jobb26 = jobb
     
 End Function
 
Function Jobbxx(ByVal inkomst, Optional ByVal alder = 30, Optional ByVal ksats = 0.3144, Optional ByVal pbb = 42400, Optional ByVal marginal = 0, Optional ByVal binkomst = 0, _
  Optional ByVal year = 2100, Optional ByVal wyear = 21000, Optional ByVal IBB = 0, Optional ByVal kvoten = 1, Optional ByVal Iyear = 0, Optional ByVal Xage = 66) As Double
    'Year       - inkomstår, wyear=När indexeringen ska börja, iyear=skattreglerna
    'xage       - Ålder för det förstärkta jobbskattavadraget från och med 2019
    'Application.Volatile
    If Iyear = 0 Then Iyear = year
 
    If Iyear < 2007 Then
     Jobbxx = 0
    ElseIf Iyear = 2007 Then
     Jobbxx = Jobb07(inkomst, alder, ksats, pbb, marginal, binkomst, year, wyear, IBB, kvoten, Iyear)
    ElseIf Iyear = 2008 Then
     Jobbxx = Jobb08(inkomst, alder, ksats, pbb, marginal, binkomst, year, wyear, IBB, kvoten, Iyear)
    ElseIf Iyear <= 2010 Then
     Jobbxx = Jobb10(inkomst, alder, ksats, pbb, marginal, binkomst, year, wyear, IBB, kvoten, Iyear)
    ElseIf Iyear = 2011 Then
     Jobbxx = Jobb11(inkomst, alder, ksats, pbb, marginal, binkomst, year, wyear, IBB, kvoten, Iyear)
    ElseIf Iyear < 2015 Then
     Jobbxx = Jobb14(inkomst, alder, ksats, pbb, marginal, binkomst, year, wyear, IBB, kvoten, Iyear)
    ElseIf Iyear < 2019 Then
     Jobbxx = Jobb16(inkomst, alder, ksats, pbb, marginal, binkomst, year, wyear, IBB, kvoten, Iyear)
     ElseIf Iyear < 2022 Then
     Jobbxx = Jobb19(inkomst, alder, ksats, pbb, marginal, binkomst, year, wyear, IBB, kvoten, Iyear)
    ElseIf Iyear = 2022 Then
     Jobbxx = Jobb22(inkomst, alder, ksats, pbb, marginal, binkomst, year, wyear, IBB, kvoten, Iyear, Xage)
    ElseIf Iyear = 2023 Then
     Jobbxx = Jobb23(inkomst, alder, ksats, pbb, marginal, binkomst, year, wyear, IBB, kvoten, Iyear, Xage)
    ElseIf Iyear = 2024 Then
     Jobbxx = Jobb24(inkomst, alder, ksats, pbb, marginal, binkomst, year, wyear, IBB, kvoten, Iyear, Xage)
    ElseIf Iyear = 2025 Then
     Jobbxx = Jobb25(inkomst, alder, ksats, pbb, marginal, binkomst, year, wyear, IBB, kvoten, Iyear, Xage)
    ElseIf Iyear = 2026 Then
     Jobbxx = Jobb26(inkomst, alder, ksats, pbb, marginal, binkomst, year, wyear, IBB, kvoten, Iyear, Xage)
    Else
     Jobbxx = Jobb26(inkomst, alder, ksats, pbb, marginal, binkomst, year, wyear, IBB, kvoten, Iyear, Xage)
    End If
     
End Function

''Sub kolljobb()
''    'Call propp
''    Rem (Inkomst, Optional alder = 30, Optional ksats = 0.3144, Optional PBB = 42800, Optional marginal = 0, Optional binkomst = 0, _
''     Optional year = 2100, Optional wyear = 21000, Optional ibb = 0, Optional Kvoten = 1, Optional iyear = 0)
''    Dim inkomst As Double
''    Dim alder As Integer
''    Dim pbb As Long
''
''    alder = 65
''    pbb = 50000
''    Dim i As Integer
''    Dim kollavd As Double
''    Dim diff As Double
''    diff = 0
''    For i = 1 To 9
''        inkomst = 90000 + (i - 1) * 1000
''        kollavd = Jobbxx(inkomst, alder, 0.32, pbb, 1, 0, 2023)
''        If i = 1 Then diff = kollavd
''        If i = 2 Then diff = kollavd - diff
''        Debug.Print inkomst; kollavd; diff
''    Next i
''
''End Sub

Function sared(ByVal inkomst, ByVal year, Optional ByVal marginal = 0, Optional ByVal ksats = 0.3212, Optional ByVal pbb = 45500, _
                Optional ByVal wyear = 21000, Optional ByVal IBB = 0, Optional ByVal kvoten = 1, Optional ByVal Iyear = 0) As Double
    'Skattreduktion för SA tagare
    'Inkomst    - Underlag för reduktion, SA inkomsten
    'Year       - Inkomstår
    'Marginal   - Marginalberäkning
    'ksats      - Kommunalaskattesatsen
    'PBB        - Prisbasbeloppet
    'IBB        - Inkomstbasbelopp
    'wyear      - När indexeringen ska börja
    'iyear      - skattreglerna
    'Kvoten     - Kvoten mellan pbb/ibb vid ikraftträdanden
   
    sared = 0
    If year < 2018 Then Exit Function
     '   If year > wyear And Kvoten < 1 Then
     '     PBB = ibb 'Ersätt pbb med IBB och korrigera gränserna till de gamla
     '     PBB = PBB * Kvoten
     '  End If
    If year < 2022 Then
        If inkomst < 2.53 * pbb Then
            sared = 0.045 * inkomst
        Else
            sared = 0.045 * 2.53 * pbb + 0.025 * (inkomst - 2.53 * pbb)
        End If
        sared = sared * ksats
    End If
    
    If year < 2026 Then
       If inkomst <= 0.91 * pbb Then
            sared = inkomst
       ElseIf inkomst <= 3.24 * pbb Then
            sared = 0.3405 * (inkomst - 0.91 * pbb) + 0.91 * pbb
       Else
            sared = 0.128 * (inkomst - 1.703 * pbb)
       End If
       avdrag = avdragxx(inkomst + binkomst, pbb, marginal, alder, year, wyear, IBB, kvoten, Iyear)
       sared = (avdrag - sared) * ksats
       If (inkomst * 0.045 * ksats) > sared Then sared = inkomst * 0.045 * ksats
    End If
    
    If year > 2025 Then
       If inkomst <= 0.91 * pbb Then
            sared = inkomst
       ElseIf inkomst <= 3.24 * pbb Then
            sared = 0.3874 * (inkomst - 0.91 * pbb) + 0.91 * pbb
       Else
            sared = 0.251 * (inkomst - 1.813 * pbb)
       End If
       avdrag = avdragxx(inkomst + binkomst, pbb, marginal, alder, year, wyear, IBB, kvoten, Iyear)
       sared = (avdrag - sared) * ksats
       If (inkomst * 0.045 * ksats) > sared Then sared = inkomst * 0.045 * ksats
    End If
    
    If marginal = 0 Then sared = Int(sared)
    
End Function

''Sub kollsared()
''    Dim inkomst As Double
''    inkomst = 100000
''    Dim koll As Double
''    Dim i As Integer
''
''    For i = 1 To 5
''        koll = sared(inkomst, 2018, 0)
''        Debug.Print inkomst; koll
''        inkomst = inkomst + 25000
''
''    Next i
''
''End Sub

Function FAared(ByVal inkomst, ByVal year, ByVal marginal, Optional ByVal pbb = 47600, _
    Optional ByVal Iyear = 21000, Optional ByVal IBB = 0, Optional ByVal kvoten = 1) As Double
    'Skattreduktion för förvärvsinkomster: BP2021 - Antar att reduktionen följer PBB
    'Inkomst    - Underlag för reduktion, beskattningsbar inkomst
    'Year       - Inkomstår
    'Marginal   - Marginalberäkning
    'PBB        - Prisbasbeloppet
    'Iyear      - När reduktionen ska inkomstindexeras
    'IBB        - Inkomstbasbelopp
    'Kvoten     - Kvoten mellan pbb/ibb vid ikraftträdanden
  
    FAared = 0
    
    If year < 2021 Then Exit Function
    
    If year > wyear And kvoten < 1 Then
        pbb = IBB 'Ersätt pbb med IBB och korrigera gränserna till de gamla för ikraftträdandeåret
        pbb = pbb * kvoten
    End If

    Dim limit1, limit2, andel As Double
    If year >= 2021 Then
        limit1 = 40000
        limit2 = 240000
        andel = 0.0075
''    ElseIf year = 2022 Then
''        'Ytterligare förstärkning av skattereduktionen för förvärvsinkomster
''        limit1 = 65000
''        limit2 = 265000
''        andel = 0.0141
''    Else
''        If pbb < 48600 Then pbb = pbb / 48600
''        limit1 = 65000
''        limit2 = 265000
''         andel = 0.0141
    End If
    
    If inkomst > limit1 And inkomst < limit2 Then
            FAared = andel * (inkomst - limit1)
    ElseIf inkomst >= limit2 Then
            FAared = 1500 'andel * (limit2 - limit1)
    End If
    
    If FAared > inkomst Then FAared = inkomst

    If marginal = 0 Then FAared = Int(FAared)

End Function

''Sub kollFAred()
''    Dim inkomst As Double
''    inkomst = 175000
''    Dim koll As Double
''    Dim i As Integer
''
''    For i = 2021 To 2024
''        koll = FAared(inkomst, i, 0, 47600, 2200, 69300, 1)
''        Debug.Print inkomst; koll
''        inkomst = inkomst + 25000
''
''    Next i
''
''End Sub

Function pandred(wage, year, marginal)
    'Tillfällig skattereduktion för panedeminskostnader 2021 och 2022
    'Wage - Arbetsinkomst
    'year - Inkomstår
    pandred = 0
    If year < 2021 Or year > 2023 Then Exit Function
    If wage < 60000 Or wage > 500000 Then Exit Function
    
    pandred = 2250
    If wage < 240000 Then pandred = 0.0125 * (wage - 60000)
    If wage > 300000 Then pandred = 2250 - 0.01125 * (wage - 300000)
    If marginal = 0 Then pandred = Int(pandred)
End Function
''Sub kollpand()
'' Dim koll As Double
'' koll = pandred(300000, 2021, 1)
'' Debug.Print koll
''End Sub

Function arbgiv(ByVal wage, ByVal year, ByVal age, Optional ByVal marginal = 0, Optional ByVal Typ = 0) As Double
    'Arbetsgivaravgiften obs före 1995 saknas uppgifter, se flik k_skatt
        
    'Wage - inkomst
    'Year - inkomstår
    'Age  - Ålder 1/1 inkomståret
    'marginal - Avrundningsregler
    'Typ 0 kronor, 1 sats, 2 dagens 10,21% x [wage, 1 om typ=1]
    
    arbgiv = 0
    If wage < 1000 Then Exit Function 'Har varierat över tid?
    If year < 1960 Then Exit Function
    
    Dim sats As Double
       
    Dim rikt As Integer
    rikt = Application.Range("Rng_riktage").Value
    
    If (IsMissing(Application.Range("RulesfromSkatt")) Or Application.Range("RulesfromSkatt") = 0) Then
        'Nothing
        Else
        If Application.Range("Rules") = 1 Then
            If Int(rikt + born) > Application.Range("RulesfromSkatt") Then rikt = riktage(Application.Range("RulesfromUtg"), 1)
            Else
              If Int(born) + rikt > Application.Range("RulesfromSkatt") Then rikt = riktage(Application.Range("RulesfromSkatt"), 1)
        End If
    End If
   
    If age < rikt And Typ = 0 Then
        sats = Worksheets("k_skatt").Cells(year - 1928, 10) 'OBS fast adress
    Else
        sats = Worksheets("k_skatt").Cells(year - 1928, 11) 'OBS fast adress "10,21"
    End If
    arbgiv = wage * sats
    
    If marginal = 0 And (Typ <> 1) Then
        arbgiv = Int(arbgiv + 0.5)
        Exit Function
    End If
            
    If Typ = 1 Then arbgiv = sats
    
    
End Function

''Sub koll_arbgiv()
''    Const born = 1960
''    Dim wage As Double
''    Dim year As Integer
''    Dim koll As Double
''    Dim age As Integer
''    age = 64
''    wage = 10000
''
''    For year = 2024 To 2026
''        koll = arbgiv(wage, year, age, 0, 0)
''        age = age + 1
''        Debug.Print year; age, koll
''    Next year
''End Sub

