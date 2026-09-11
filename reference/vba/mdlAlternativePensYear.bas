Attribute VB_Name = "mdlAlternativePensYear"
Option Explicit

Sub Alternative_Pens_Year(ByVal BornYear As Double)

    Dim rng As Range
    Dim r As Range
    Dim blnAlternative_Pens_Year As Boolean
    blnAlternative_Pens_Year = False
    
    Application_Rest 'Modellen hamnar i villoläge
    Set rng = Application.Range("rng_Nyckeltal_Pensyear")
    Set rng = rng.Resize(rng.CurrentRegion.Rows.Count - 2, 1)
    For Each r In rng
        If BornYear <= r.Value Then
            Application.Range("PARYear") = r.Offset(, 2)
            blnAlternative_Pens_Year = True
            Exit For
        End If
    Next r
    If Not blnAlternative_Pens_Year Then
        Application.Range("PARYear") = 65
    End If
    Application_Wakeup 'Modellen i normalläge
    Set rng = Nothing
End Sub

Sub Alternative_Pens_Year_Indata(ByVal BornYear As Double, ByVal c As Range)

    Dim rng As Range
    Dim r As Range
    Dim blnAlternative_Pens_Year As Boolean
    blnAlternative_Pens_Year = False
    
    Set rng = Application.Range("rng_Nyckeltal_Pensyear")
    Set rng = rng.Resize(rng.CurrentRegion.Rows.Count - 2, 1)
    For Each r In rng
        If BornYear <= r.Value Then
            c.Offset(, 2).Value = r.Offset(, 2)
            blnAlternative_Pens_Year = True
            Exit For
        End If
    Next r
    If Not blnAlternative_Pens_Year Then
        c.Value = 65
    End If
  
    Set rng = Nothing
End Sub

Sub riktalder(ByVal BornYear As Double)

    Dim rng2 As Range
    Dim r As Range
    Dim blnRiktalder As Boolean
    blnRiktalder = False
    
    Application_Rest 'Modellen hamnar i villoläge
    Set rng2 = Application.Range("rng_Nyckeltal_Pensyear")
    Set rng2 = rng2.Resize(rng2.CurrentRegion.Rows.Count - 2, 1)
    For Each r In rng2
        If BornYear <= r.Value Then
            Application.Range("PARYear") = r.Offset(, 3)
            blnRiktalder = True
            Exit For
        End If
    Next r
    If Not blnRiktalder Then
        Application.Range("PARYear") = 65
    End If
    Application_Wakeup 'Modellen i normalläge
    Set rng2 = Nothing
End Sub

Sub Riktalder_Indata(ByVal BornYear As Double, ByVal c As Range)

    Dim rng As Range
    Dim r As Range
    Dim blnRiktalder As Boolean
    blnRiktalder = False
    
    Set rng = Application.Range("rng_Nyckeltal_Pensyear")
    Set rng = rng.Resize(rng.CurrentRegion.Rows.Count - 2, 1)
    For Each r In rng
        If BornYear <= r.Value Then
            c.Offset(, 2).Value = r.Offset(, 4)
            blnRiktalder = True
            Exit For
        End If
    Next r
    If Not blnRiktalder Then
        c.Value = 65
    End If
  
    Set rng = Nothing
End Sub

'''Sub Tjanste_Uttag_Ar()
'''    Dim rng As Range
'''    Set rng = fnGetRangeFromNames(ThisWorkbook, "rng_TjänsteUttagÅr")
'''    If rng.Value = True Then
'''        If Application.Range("PARYear").Value < 61 Then
'''            Application.Range("rng_TjänsteUttag_FastÅr") = 65
'''        Else
'''            Application.Range("rng_TjänsteUttag_FastÅr") = Application.Range("PARYear").Value
'''        End If
'''    Else
'''        Application.Range("rng_TjänsteUttag_FastÅr") = Null
'''    End If
'''    Set rng = Nothing
'''End Sub
Sub chkAlternativePensClick()
    If Application.Range("rng_Rek_Pens_Ålder") Then _
        Alternative_Pens_Year (Application.Range("BornYear"))
End Sub
Sub chkRiktalderClick()
    If Application.Range("rng_Riktålder") Then _
        riktalder (Application.Range("BornYear"))
End Sub
