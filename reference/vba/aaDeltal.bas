Attribute VB_Name = "aaDeltal"
Option Explicit

Public aDeltal_IP() As Double 'Innehåller delningstal, inläst från flik Nyckeltal
Public aDeltal37_IP() As Double 'IP delningstal för de födda 1937 och äldre; en dimension
Public aDeltal37_PP() As Double 'PP delningstal för de födda 1937 och äldre; en dimension
Public aDeltal_PP() As Double 'Innehåller delningstal, inläst från flik Nyckeltal

Sub testfnDeltal()
    Dim x As Double
    Dim i As Long


    'Debug.Print fnDeltal_IP(1957, 82, False)
    'Debug.Print fnDeltal_IP(1958, 82, False)
    Debug.Print fnDeltal_IP(1959, 67 + 11 / 12, True)

'    x = 66
'    For i = 1 To 11
'        Debug.Print "PAR=" & x & ": "; fnDeltal_IP(1959, x)
'        x = x + 0.1
'    Next
    
End Sub
Public Function fnDeltal_IP2(ByVal born As Double, ByVal curAge As Double, Optional ByVal PAR As Double, Optional ByVal def_ar As Double) As Double
    'Same as fnDeltal_IP but with other in parameters.
    'Om PAR=curAge eller Def_ar = curAge kommer en sammavägning göras mellan deltal för int(curAge) och int(curAge)+1
    'T ex Om curAge=66,3 kommer sammanvägningen göras av deltal för age=66 och age=67
    'Dessutom, om curAge är 66,0 så kommer sammanvägningen göras som om att curAge=66,5 (att man går i pension mitt på året)
    Dim blnAdjustProportional As Boolean
    
    If curAge > def_ar Then curAge = def_ar + 1 'Egentligen borde inte funktionen anropas efter man passerat def_ar?! MB
    
    If Int(PAR) = curAge Or Int(def_ar) = curAge Then
        'Händer alltså dels när man tar ut pension första gången, eller vid definitivt uttag (partiellt uttag)
        blnAdjustProportional = True
    Else
        blnAdjustProportional = False
    End If
    
    fnDeltal_IP2 = fnDeltal_IP(born, curAge, blnAdjustProportional)


End Function
Public Function fnDeltal_IP(ByVal born As Double, ByVal curAge As Double, Optional ByVal blnAdjustProportional As Boolean = False) As Double
    'Return deltal for born year and curAge
    'blnAdjustProportional - antagandet att om curAge har decimaldel (int(curAge)<>curAge så görs en sammanvägning av närliggande deltalen.
    '                      - antagandet att när man går i pension, så gör man det mitt på året om curAge är ett heltal, alltså att int(curAge)=Age
    '                      - T ex Om curAge=66,3 kommer sammanvägningen göras av deltal för age=66 och age=67
    
    Dim d1 As Double
    Dim d2 As Double
    Dim ageFraction As Double
    Dim month As Double 'Används för få till hela månader
    
    If Int(born) <> born Then
        Stop 'funktionen hanterar bara födelseår som heltal
        End
    End If
    
    subLoadDeltal_IP_PPifneeded
    
    ageFraction = curAge - Int(curAge)
    
    If Int(born) < 1930 Then
        Stop 'Funktionen hanterar inte födda före 1930
        Exit Function
    End If
    
    'Hantera födda 1937 eller tidigare
    If Int(born) <= 1937 Then
        'Hanterar inte månader för födda 1937 eller tidigare
        fnDeltal_IP = aDeltal37_IP(Int(born))
        Exit Function
    End If
    
    If blnAdjustProportional Then
        If ageFraction = 0 Then
            curAge = curAge + 0.5 'Antagandet om att gå i pension mitt på året
            ageFraction = 0.5
        End If
    End If
    
    month = Round((curAge - Int(curAge)) * 12, 0)     'Antal månader vid partiellt uttag
    
    If ageFraction > 0 Then
        'Beräkningen görs om en proportionell/linjär sammanvägning av de två närliggande delningstalen
        d1 = aDeltal_IP(born, Int(curAge))
        d2 = aDeltal_IP(born, Int(curAge) + 1)
        'fnDeltal_IP = (1 - ageFraction) * d1 + ageFraction * d2
        fnDeltal_IP = (12 - month) / 12 * d1 + month / 12 * d2
    Else
        fnDeltal_IP = aDeltal_IP(born, curAge) 'PAR angivet som heltal
    End If

    fnDeltal_IP = Int(fnDeltal_IP * 100 + 0.4999) / 100 'Avrunda två decimaler. ,xx5 avrundas nedåt, annars som round(x,2)
    
End Function
Public Function fnDeltal_PP2(ByVal born As Double, ByVal curAge As Double, Optional ByVal PAR As Double, Optional ByVal def_ar As Double) As Double
    'Same as fnDeltal_PP but with other in parameters.
    'Om PAR=curAge eller Def_ar = curAge kommer en sammavägning göras mellan deltal för int(curAge) och int(curAge)+1
    'T ex Om curAge=66,3 kommer sammanvägningen göras av deltal för age=66 och age=67
    'Dessutom, om curAge är 66,0 så kommer sammanvägningen göras som om att curAge=66,5 (att man går i pension mitt på året)
    Dim blnAdjustProportional As Boolean
    If Int(PAR) = curAge Or Int(def_ar) = curAge Then
        'Händer alltså dels när man tar ut pension första gången, eller vid definitivt uttag (partiellt uttag)
        blnAdjustProportional = True
    Else
        blnAdjustProportional = False
    End If
    
    fnDeltal_PP2 = fnDeltal_PP(born, curAge, blnAdjustProportional)


End Function
Public Function fnDeltal_PP(ByVal born As Long, ByVal curAge As Double, Optional blnVidHeltalAnta6mån As Boolean = False) As Double
    'Return deltal for born year and curAge
    'blnVidHeltalAnta6mån - om antagandet att när man går i pension, så gör man det mitt på året.
    Dim d1 As Double
    Dim d2 As Double
    Dim ageFraction As Double
    Dim month As Double 'Används för få till hela månader
    
    subLoadDeltal_IP_PPifneeded
    
    ageFraction = curAge - Int(curAge)
    
    If Int(born) < 1930 Then
        Stop 'Funktionen hanterar inte födda före 1930
        Exit Function
    End If
    
    'Hantera födda 1937 eller tidigare
    If Int(born) <= 1937 Then
        'Hanterar inte månader för födda 1937 eller tidigare
        fnDeltal_PP = aDeltal37_PP(Int(born))
        Exit Function
    End If
    
    If blnVidHeltalAnta6mån Then
        If ageFraction = 0 Then
            curAge = curAge + 0.5 'Antagandet om att gå i pension mitt på året
            ageFraction = 0.5
        End If
    End If
    
    month = Round((curAge - Int(curAge)) * 12, 0)     'Antal månader vid partiellt uttag
    
    If ageFraction > 0 Then
        'Beräkningen görs om en proportionell/linjär sammanvägning av de två närliggande delningstalen
        d1 = aDeltal_PP(born, Int(curAge))
        d2 = aDeltal_PP(born, Int(curAge) + 1)
        'fnDeltal_PP = (1 - ageFraction) * d1 + ageFraction * d2
        fnDeltal_PP = (12 - month) / 12 * d1 + month / 12 * d2
    Else
        fnDeltal_PP = aDeltal_PP(born, curAge) 'PAR angivet som heltal
    End If
    fnDeltal_PP = Int(fnDeltal_PP * 100 + 0.4999) / 100 'Avrunda två decimaler. ,xx5 avrundas nedåt, annars som round(x,2)
    
End Function
Sub subLoadDeltal_IP_PPifneeded()
    'Läser in delningstal (IP och PP) från flik Nyckeltal
    Dim rng As Range
    Dim rng2 As Range
    Dim ageTo As Long
    Dim ageFrom As Long
    Dim bornTo As Long
    Dim bornFrom As Long
    Dim aTmp As Variant 'temporary array to copy from worksheet Nyckeltal
    Dim aMort As Variant 'temporary array to copy from worksheet Mortality
    Dim r As Long
    Dim c As Long
    Dim readFromMortIPFromBornYear As Long
    Dim readFromMortPPFromBornYear As Long
    
    readFromMortIPFromBornYear = ThisWorkbook.Names("rngDelnIPMort").RefersToRange.Value
    readFromMortPPFromBornYear = ThisWorkbook.Names("rngDelnPPMort").RefersToRange.Value
    
    'First check if array needs to be loaded
    On Error Resume Next
    If UBound(aDeltal_IP) <> 99 Then
        'Dummy
    End If
    If err.Number = 0 Then
        On Error GoTo 0
        Exit Sub 'Already loaded
    Else
        'In rare case aDeltal37_IP/PP.. exist
        On Error Resume Next
        Erase aDeltal37_IP
        Erase aDeltal37_PP
        Erase aDeltal_PP
        Erase aDeltal_PP
        On Error GoTo 0
    End If
    On Error GoTo 0
    
    'Ladda delningtal IP - först från flik Nyckeltal, allt läses in oavsett readFromMortIPFromBornYear
    Set rng = ThisWorkbook.Names("rng_Delningtal_FirstAge").RefersToRange
    ageFrom = rng.Value
    ageTo = rng.End(xlToRight).Value
    Set rng2 = ThisWorkbook.Names("rng_Delningtal_FirstYear").RefersToRange
    bornFrom = rng2.Value
    bornTo = rng2.End(xlDown).Value
    ReDim aDeltal_IP(bornFrom To bornTo, ageFrom To ageTo)
    Set rng2 = Range(rng2, rng2.End(xlDown)).Offset(, 1) 'left column correct
    Set rng2 = rng2.Resize(, ageTo - ageFrom + 1) 'Correct range
    
    aTmp = rng2.Value 'Move to array
    'Fill array with correct born/age dimension
    For r = 1 To rng2.Rows.Count
        For c = 1 To rng2.Columns.Count
            aDeltal_IP(r + bornFrom - 1, c + ageFrom - 1) = aTmp(r, c)
        Next
    Next
    
    'Kontroll om något ska läsas in från flik Mortalitet IP
    If readFromMortIPFromBornYear <> 0 Then
        If readFromMortIPFromBornYear < bornTo Then
            'Läs in/skriv över from readFromMortIPFromBornYear
            Set rng = ThisWorkbook.Names("rngMortTopLeft").RefersToRange
            Set rng = Range(rng, rng.End(xlDown))
            Set rng = Range(rng, rng.End(xlToRight))
            aMort = rng.Value
            'column 2=sex, 3=Age, 4=Cohort, 5=IP, 6=PP
            For r = 1 To UBound(aMort, 1) - LBound(aMort, 1) + 1 'Loopa igenom hela tabelln
                If aMort(r, 2) = 0 Then
                    'Sex korrekt
                    If aMort(r, 4) >= readFromMortIPFromBornYear Then
                        'BornYear ok
                        'aDeltal(Cohort/born, Age)
                        If aMort(r, 4) > UBound(aDeltal_IP, 1) Then Exit For 'Om data in är synkat
                        If aMort(r, 3) <= UBound(aDeltal_IP, 2) Then  'To many ages in Mortality
                            aDeltal_IP(aMort(r, 4), aMort(r, 3)) = aMort(r, 5)
                        End If
                    End If
                End If
            Next
        End If
    End If
    
    'Ladda delningtal PP
    Set rng = ThisWorkbook.Names("rng_PP_deltal_Year").RefersToRange
    ageFrom = rng.Offset(, 1).Value
    ageTo = rng.Offset(, 1).End(xlToRight).Value
    bornFrom = rng.Offset(1).Value
    bornTo = rng.Offset(1).End(xlDown).Value
    ReDim aDeltal_PP(bornFrom To bornTo, ageFrom To ageTo)
    Set rng = rng.Offset(1, 1) 'topleft
    Set rng = rng.Resize(bornTo - bornFrom + 1, ageTo - ageFrom + 1) 'Correct range
    
    aTmp = rng.Value 'Move to array
    'Fill array with correct born/age dimension
    For r = 1 To rng.Rows.Count
        For c = 1 To rng.Columns.Count
            aDeltal_PP(r + bornFrom - 1, c + ageFrom - 1) = aTmp(r, c)
        Next
    Next
    
    'Kontroll om något ska läsas in från flik Mortalitet PP
    If readFromMortPPFromBornYear <> 0 Then
        If readFromMortPPFromBornYear < bornTo Then
            'Läs in/skriv över from readFromMortPPFromBornYear
            'column 2=sex, 3=Age, 4=Cohort, 5=IP, 7=PP
            For r = 1 To UBound(aMort, 1) - LBound(aMort, 1) + 1 'Loopa igenom hela tabelln
                If aMort(r, 2) = 0 Then
                    'Sex korrekt
                    If aMort(r, 4) >= readFromMortPPFromBornYear Then
                        'BornYear ok
                        'aDeltal(Cohort/born, Age)
                        If aMort(r, 4) > UBound(aDeltal_PP, 1) Then Exit For 'Om data in är synkat
                        If aMort(r, 3) <= UBound(aDeltal_PP, 2) Then 'To many ages in Mortality
                            aDeltal_PP(aMort(r, 4), aMort(r, 3)) = aMort(r, 6) 'PP
                        End If
                    End If
                End If
            Next
        End If
    End If
    
    'ladda array för födda 1937 eller tidigare
    Set rng = ThisWorkbook.Names("rng_Nyckeltal_Deltal_IP").RefersToRange
    Set rng = rng.Offset(1, -1) 'First age/cohort
    ageFrom = rng.Value
    ageTo = rng.End(xlDown).Value
    ReDim aDeltal37_IP(ageFrom To ageTo)
    ReDim aDeltal37_PP(ageFrom To ageTo)
    Set rng = rng.Offset(, 1)
    
    For r = 1 To ageTo - ageFrom + 1
        aDeltal37_IP(r + ageFrom - 1) = rng.Cells(r, 1).Value
        aDeltal37_PP(r + ageFrom - 1) = rng.Cells(r, 2).Value
    Next
    
    aTmp = ""
    aMort = ""
    Set rng2 = Nothing
    Set rng = Nothing
End Sub

