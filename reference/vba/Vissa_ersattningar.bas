Attribute VB_Name = "Vissa_ersattningar"
' Vissa individanknutna funktioner som räknar akassa sjukpenning, SA mm
Option Explicit

Function akassa(ByVal inkomst, Optional ByVal dagar = 264, Optional ByVal karens = 7, Optional ByVal year = 2013) As Double
    'UO14
    akassa = 0
    'Application.Volatile

    Dim dag As Double

    dag = (inkomst / 12) / 22 '264 dagar
    dag = Int(dag + 1) ' Avrundad uppått till närmaste heltal
    'If dag < 10 Then dag = 0
  
    ' maximal antal dagar under året är 264
    ' Om dagar(månader) > 264 (13) antas arbetslösheten starta före årsskiftet
    If dagar > 264 Then
       dagar = 264
       karens = 0
    End If
  
  'Dim karens As Long
  Dim dagers As Double
  Dim Tak As Integer
  Dim golv As Integer
  golv = 320
  Tak = 680
  
  'Idag delar man upp på barnfamlijer med försörjningsansvar antal dagar mm nedan är mycket förenklat
  If year = 2015 Then
    golv = 335
    Tak = 757
  End If
  If year = 2016 Then
    golv = 365
    Tak = 910
  End If
  
  If year = 2021 Then karens = 6 'Tidigare år?
  If year > 2021 Then karens = 2
  
  If dagar <= karens Then
     akassa = 0
     Exit Function
  ElseIf dagar <= (100 + karens) And year <= 2006 Then
    dagers = dag * 0.8
    If 0.8 * dag < golv Then dagers = golv
    If 0.8 * dag > 730 Then dagers = 730
    akassa = (dagar - karens) * dagers
    
  ElseIf dagar <= (200 + karens) Then
    dagers = dag * 0.8
    If 0.8 * dag < golv Then dagers = golv
    If 0.8 * dag > Tak Then dagers = Tak
    akassa = (100) * dagers
    dagers = dag * 0.7
    If 0.7 * dag < golv Then dagers = golv
    If 0.7 * dag > Tak Then dagers = Tak
    akassa = akassa + (dagar - 100 - karens) * dagers
    
    
    If year < 2007 Then
        If 0.8 * dag > 730 Then
            akassa = (dagar - 100 - karens) * dagers + 100 * 730
        Else
            akassa = (dagar - 100 - karens) * dagers + 100 * dagers
        End If
    End If
    
  Else
    dagers = 0.8 * dag
    If 0.8 * dag < golv Then dagers = golv
    If 0.8 * dag > Tak Then dagers = Tak
    akassa = (100) * dagers
    dagers = 0.7 * dag
    If 0.7 * dag < golv Then dagers = golv
    If 0.7 * dag > Tak Then dagers = Tak
    akassa = akassa + (100) * dagers
    dagers = 0.65 * dag
    If 0.65 * dag < golv Then dagers = golv
    If 0.65 * dag > Tak Then dagers = Tak
    akassa = akassa + (dagar - 200 - karens) * dagers
    
    If year < 2007 Then
        dagers = 0.8 * dag
        If 0.8 * dag > 730 Then
            akassa = (dagar - 100 - karens) * dagers + 100 * 730
        Else
            akassa = (dagar - 100 - karens) * dagers + 100 * dagers
        End If
    End If
    
  End If
    'OBS staten betalar ÅP avgift om 10,21 procent
    
End Function

'''Function akassa06(inkomst, Optional dagar = 264, Optional karens = 5) As Double
'''  'Application.Volatile
'''  akassa06 = 0
'''
'''  Dim dag As Double
'''
'''  dag = (inkomst / 12) / 22 '264 dagar
'''  dag = Int(dag + 1) ' Avrundad uppått till närmaste heltal
'''  If dag < 10 Then dag = 0
'''  Dim dagers As Double
'''
'''  If dagar <= karens Then
'''     akassa06 = 0
'''  ElseIf dagar <= 100 Then
'''    dagers = dag * 0.8
'''    If 0.8 * dag < 320 Then dagers = 320
'''    If 0.8 * dag > 730 Then dagers = 730
'''    akassa06 = (dagar - karens) * dagers
'''  Else
'''    dagers = 0.8 * dag
'''    If 0.8 * dag < 320 Then dagers = 320
'''    If 0.8 * dag > 730 Then dagers = 730
'''    akassa06 = 100 * dagers '100 dagar förhöjt
'''    dagers = 0.8 * dag
'''    If 0.8 * dag < 320 Then dagers = 320
'''    If 0.8 * dag > 680 Then dagers = 680
'''    akassa06 = akassa06 + (dagar - 100 - karens) * dagers
'''  End If
'''   'akassa06 = dagers
''''OBS staten betalar ÅP avgift om 10,21 procent
'''End Function


Function sjuk(ByVal inkomst, Optional ByVal dagar = 0, Optional ByVal Typ = 1, Optional ByVal Tak = 7.5, _
Optional ByVal faktor = 0.97, Optional ByVal level = 0.8, Optional ByVal basb = 42800, Optional ByVal year = 2013) As Double
    'Application.Volatile
    'Inkomst
    'Dagar under året
    'Typ 0) Föräldrap 1)Sjukp,  2) Sjuklön 3)Sjuklön alt b 4)Löneavdrag
    'Tak upp till X prisbasbelopp 7,5 (8) för sjukpenning, 10 föräldrapenning
    'Faktor av SGI
    'Ersättningsnivå
    'Basb- prisbasbeloppet
    
    'If Inkomst > 120000 Then Inkomst = Inkomst / 12 'Månadsinkomst
    'Vid några dagaras frånvaro räknar arbetsgivaren ett avdrag på lön. Max nio dagar per sjuktillfälle
     
    '2008-01 SGI ska multipliceras med 0,97 vid beräkning av föräldrapenning och övrig dagersättning (sjukpenning m.m.).
    faktor = 1
    If year > 2006 Then faktor = 0.97
    
    Dim lavdrag As Double
    Dim sjukw As Double
    Dim sjukp As Double
    Dim sgi As Double
    Dim sjpdag As Double
    Dim dagp As Double
 
    lavdrag = 0
    If dagar = 1 Then
      lavdrag = inkomst * 12 / 260
    ElseIf dagar <= 7 Then
     lavdrag = inkomst * 12 / 260 + 0.2 * inkomst * (12 / 260) * Application.min(7 - 2 - 1, dagar)
    Else
     lavdrag = inkomst * 12 / 260 + 0.2 * inkomst * (12 / 260) * Application.min(14 - 4 - 1, dagar)
    End If
    
    'Sjuklön
    sjukw = 0
    If dagar > 0 Then sjukw = 0.8 * inkomst * (12 / 260) * Application.min(14 - 4 - 1, dagar - 1) '1 karens och 4 helgdagar under sjukperioden.
    
    'Sjukpenningen beräknas utifrån SGI per dag, men de först 14 dagarna består av sjuklön.
    sgi = Application.min(12 * inkomst, basb * Tak)
    sgi = Int(faktor * sgi / 100) * 100
    sjpdag = Application.max(dagar - 14, 0)
 
    dagp = Int(sgi / 365 + 0.5)
    'LEVEL=0,75 efter 300 dagar, styrs av level genom vägt medelvärde på antal dagar under året. Om sjukperioden är över ett år krävs en handpåläggning:
    'If dagar > 300 Then level = (0.8 * 300 + 0.75 * (dagar - 300)) / dagar
    
    sjukp = level * dagp * sjpdag
     
    sjuk = sjukp
 
    If Typ = 2 Then
      ' Sjuklön från arbetsgivaren
     If dagar > 29 Then
        sjuk = inkomst * 0.8 * 9 * 12 / 260  'Månadslönen per dag * max nio dagar då helgen räknas bort
     Else
        sjuk = inkomst - lavdrag 'Vid färre dagar lön - avdraget
      End If
    End If
    
    If Typ = 3 Then sjuk = sjukw
    If Typ = 4 Then sjuk = lavdrag
    
    sjuk = Int(sjuk + 0.5)
    'OBS staten betalar ÅP avgift om 10,21 procent
End Function

Function SA(inkomst As Double, Optional alder = 46, Optional basbelopp = 42800, _
   Optional Typ = 1, Optional marginal = 0, _
   Optional ink2 = 0, Optional ink3 = 0, Optional year = 2013) As Double
 'Inkomst, år t och Ink2 = Inkomst t-1 samt ink3 = inkomst t-2
 'Ålder, år t
 'Basbelopp år t
 'Typ av inkomst 1 - inkomstrelaterad + garanti, 2 Garantidelen 3 ramtiden ^=(1-3) Inkomstdelen
 
 'Särskild ersättning beräknas inte utan får handpåläggas
 'Arbetsförmågan antas vara helt nedsatt justering för 1/2 och 1/4 får ske för hand
 
 'Application.Volatile
 
 Dim ramtid As Long
 Dim antinkomst As Double
 Dim N As Long
 Dim garanti As Double
 
Select Case alder
    Case Is <= 46
        ramtid = 8
    Case Is <= 47
        ramtid = 7
    Case Is <= 50
        ramtid = 6
    Case Else
        ramtid = 5 ' uppnår 53 år eller högre ålder det år då försäkringsfallet inträffar
    End Select
    antinkomst = inkomst 'förenkling antagandeinkomsten
    'Annars A) studera bruttoinkomsten under ramtiden
    'som är B) OBS är MAX( (PGI/0,93 +PGB) * justerat för Prisbasbeloppsförändringen; 7,5 prisbasbelopp)
    'Detta sker enklast för varje år
    'Antagande inkomsten är sedan C: Genomsnittet av de tre högsta omräknade bruttoinkomsterna under ramtiden
    'obs finns bara en eller två inkomster medel av dessa
    'Om SGI överstiger medelinkomsten ta SGI dock högst 7,5 prisbasbelopp
    'I de flesta fall kan man ta de tre senaste inkomsterna och relatera de till det sista årets Prisbasbelopp och därefter medelvärdet av de tre justerade värdena
    'I programmet antas att de tre senaste åren utgör de högsta inkomsterna och att inkomsterna är justerade till sista årets priser
    If inkomst > 7.5 * basbelopp Then inkomst = 7.5 * basbelopp
    N = 3
    If ink2 > 0 Then
      If ink2 > basbelopp * 7.5 Then ink2 = basbelopp * 7.5
      If ink3 > basbelopp * 7.5 Then ink3 = basbelopp * 7.5
      If ink3 < 1 Then
         N = 2
         ink3 = 0
      End If
      inkomst = (inkomst + ink2 + ink3) / N
    End If
    
    SA = 0.64 * inkomst
    If SA < 12 Then SA = 0
    
    'Garantiersättning, ser tilll försäkringstid mm antas uppgå till minst tre år...
    garanti = 0
  Select Case alder
    Case Is <= 18
        garanti = 0  'till och med den månaden före den månad då den försäkrade fyller 21 år
        SA = 0
    Case Is <= 21
        garanti = 2.1 * basbelopp  'till och med den månaden före den månad då den försäkrade fyller 21 år
    Case Is <= 23
        garanti = 2.15 * basbelopp 'från och med den månaden då den försäkrade fyller 21 år till och med månaden före den månad då han eller hon fyller 23 år
    Case Is <= 25
        garanti = 2.2 * basbelopp
    Case Is <= 27
        garanti = 2.25 * basbelopp
    Case Is <= 29
        garanti = 2.3 * basbelopp
    Case Is <= 30
        garanti = 2.35 * basbelopp
    Case Is <= 64
        garanti = 2.4 * basbelopp
    Case Else
        garanti = 0 * basbelopp
        SA = 0
    End Select
    
    If SA < garanti Then
      garanti = garanti - SA
    Else
      garanti = 0
    End If
    
    If Typ = 1 Then
       SA = SA + garanti
    ElseIf Typ = 2 Then
       SA = garanti
    Else
      SA = SA
    End If
    If marginal = 0 Then SA = Int(SA / 12) * 12
    
    If Typ = 3 Then SA = ramtid
        'OBS staten betalar ÅP avgift om 18,5 procent på (inkomsterättningen/0,64)*(1-7%) kollas men för beräkningar i typfallsmodellen
        'ska vi ta TYP=99
    If Typ = 99 Then SA = inkomst
    
End Function


