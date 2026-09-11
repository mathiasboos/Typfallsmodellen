Attribute VB_Name = "Mortality"
Option Explicit

Public deathprobs() As Double               'year, sex, age
Public dtalip(0 To 2, 50 To 106) As Double  'sex and age for a given year
Public dtalpp(0 To 2, 50 To 106) As Double
Public dtalpp2(0 To 2, 50 To 106) As Double
Public ExpLife(0 To 2, 50 To 106) As Double
Public ExpLife2(0 To 2, 50 To 106) As Double
Public Arvf(0 To 2, 1 To 106) As Double
Public Arvf2(0 To 2, 1 To 106) As Double
'Public lx_(0 To 2, 0 To 106) As Double

Public filenr As Integer

Public Miny As Integer  'miny As Integer lowest year
Public Maxy As Integer

Public Const writefile = 0 'To file and not to the mortality sheet

Public Const agegroup = 106  'ages 0-105
   
Rem the deathrates are from SCB and are also collected in scb_risk_of_death.xlsb
Sub ReadMortality()
   
    Dim year As Integer, age As Integer, sex As Integer, cohort As Integer
    Dim tmp As Variant
    Miny = wsMortality.Range("F2").Value
    If wsMortality.Range("A2").Value > Miny Then Miny = wsMortality.Range("A2").Value
    
    Maxy = Worksheets("mortality").Range("G2").Value
    If Maxy > 2120 Then Maxy = 2120
    
    Dim Mnumbers As Long 'Worksheets("mortality").Range("H2").Formula = "=antal(a1:a130000)"
    Mnumbers = wsMortality.Range("h2").Value
    
    Dim Mortalityrows As Long
    Mortalityrows = (Maxy - Miny + 1) * 2 * agegroup '(#years +1) * sex * agegroups
    
    ReDim deathprobs(Miny To Maxy, 1 To 2, 0 To agegroup) As Double ' [year, sex, age]
    
    tmp = wsMortality.Range("a2:D" & Mnumbers + 1) '[1=year, 2=sex, 3=age, 4=risk] at row 1 'Note fixed adress A:D

    Dim i As Integer 'For risk of mortality below
    For i = 1 To Mortalityrows   '[year, sex, age]=risk
        If tmp(i, 1) > 0 Then
              If tmp(i, 3) < 106 Then deathprobs(tmp(i, 1), tmp(i, 2), tmp(i, 3)) = tmp(i, 4)
        End If
    Next i

''    For year = Miny To Maxy
''        For age = 0 To 105
''            For sex = 1 To 2
''                If age = 0 Then
''                deathprobs(year, sex, age) = 1 - Exp(-(0.5 * deathprobs(year, sex, age) + 0.8 * deathprobs(year, sex, age + 1)))
''                ElseIf age = 1 Then
''                deathprobs(year, sex, age) = 1 - Exp(-(0.2 * deathprobs(year, sex, age) + 0.6 * deathprobs(year, sex, age + 1)))
''                ElseIf age = 2 Then
''                deathprobs(year, sex, age) = 1 - Exp(-(0.4 * deathprobs(year, sex, age) + 0.5 * deathprobs(year, sex, age + 1)))
''                Else
''                deathprobs(year, sex, age) = 1 - Exp(-(0.5 * deathprobs(year, sex, age) + 0.5 * deathprobs(year, sex, age + 1)))
''                End If
''            Next sex
''        Next age
''    Next year

''    Dim koll As Double
''    koll = 0
''    For year = 2017 To 2021
''        Debug.Print year; deathprobs(year, 1, 0)
''        koll = koll + 0.5 * (deathprobs(year, 1, 0) + deathprobs(year, 2, 0))
''    Next year
''    Debug.Print 100 * Round(koll / 5, 7); '2,14
''    Stop
    Erase tmp
    
    'Write to file. Note  that the library C:\slask should exist - later if it be actual check and if not ask for ...
    If writefile = 1 Then
         filenr = FreeFile
         Open "c:\slask\dtal.csv" For Output As #filenr
         Print #filenr, "year;", "sex;", "age;", "Cohort;", "dt_IP;", "dt_PP;", "dt_PP2;", "expl;", "Exp2;"; "Inh. Gains;", "Gains2"
    End If
    
   Dim dtalstart As Integer, dtalstop As Integer
   dtalstart = Range("G4") 'Fast adress  'dtalstop = Range("Dtalstop")
   dtalstop = Range("i4")  'dtalstart = Range("Dtalstart")
   If dtalstart < Miny Then dtalstart = Miny
   If dtalstop > Maxy Then dtalstop = Maxy
   
   'If writefile = 0 Then
        'Clear the output sheet...
        Worksheets("mortality").Select
        Range("P1:Z1").Select 'Fast_adress
        'Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.ClearContents
        
         'The labels is given by "_" if pivot tabels is desired
        Dim Mlabel(1 To 11) As String
        Mlabel(1) = "Year"
        Mlabel(2) = "Sex"
        Mlabel(3) = "Age "
        Mlabel(4) = "Cohort"
        Mlabel(5) = "N/Q_NDC_IP"
        Mlabel(6) = "N/Q_DC_PP"
        Mlabel(7) = "N/Q_DC_Tjp"
        Mlabel(8) = "Exp_life"
        Mlabel(9) = "Exp_life2"
        Mlabel(10) = "Inh_gain"
        Mlabel(11) = "Inh_gain2"
        Range("P1:Z1").Value = Mlabel 'Fast adress
   'End If
   Erase Mlabel
   
   Dim counter As Long 'Number of rows to write: Ages cohorts and sex. From age 61 to 105 and each year
   counter = CLng((agegroup - 61) * (dtalstop - dtalstart + 1) * 3)  '61 first pensionage, note fixed
   
   Dim mvalues() As Double
   ReDim mvalues(1 To counter, 1 To UBound(Mlabel)) As Double 'Number of rows and columns to write
  
   Dim radut As Long
   radut = 1
   
        For cohort = 1930 To 2050 'Note
            
            Call Calculate_Deltal(cohort) '------ Calulation Look below...
            
            For age = 61 To 105
                 For sex = 0 To 2 '0=Unisex, 1=Male and 2=Female
                    mvalues(radut, 1) = cohort + age 'year
                    mvalues(radut, 2) = sex
                    mvalues(radut, 3) = age
                    mvalues(radut, 4) = cohort
                    mvalues(radut, 5) = dtalip(sex, age)
                    mvalues(radut, 6) = dtalpp(sex, age)
                    mvalues(radut, 7) = dtalpp2(sex, age)
                    mvalues(radut, 8) = ExpLife(sex, age)
                    mvalues(radut, 9) = ExpLife2(sex, age)
                    mvalues(radut, 10) = Arvf(sex, age)
                    mvalues(radut, 11) = Arvf2(sex, age)
                    radut = CLng(radut + 1)
               Next sex
            Next age
        Next cohort
 
    'If writefile = 0 Then
        wsMortality.Select
        Range("p2:z" & counter + 1).Value = mvalues
        MsgBox "Finish annuity factors "
    'End If
    
    If writefile = 1 Then
        ' Print to #filenr is done below
        Close #filenr
        MsgBox "Finish annuity factors for cohort 1930 to ... Output file c:\temp\dtal.txt"
    End If

End Sub


'-- Calculates the income pension annuity factors (delningstal), annuity facors for
'   premium pension and inheritance gains
'   Annuity factors caculated on death hazards in assumptions file
'   Income pension: dtalip(age 50-106) with 1,6% norm growth as default
'   Premium pension: dtalpp(age 50-106) default 3.2%. If 0 expected remaining lifetime
'   Inheritance gains based on a direct and simplified method based on death hazards
'   i.e. no summing up of actual cumulated pension funds for persons younger than 60
'   Creates a public array defined from 0 to 106 years: Arvsvinstfactor(y=0-106)
'   Note: Call and defintion of global variables in new_economy_2 once a year
'   - Assumed norm real growth after fees of 3%  as default for the premium
'     pension
'   - Expected remaining lifetime in the premium pension system calculated using forecasted
'     mortality in accordance with the legislation


Public Sub Calculate_Deltal(cohort As Integer)
 
    '-- Calculates pension annuity factors (delningstal)
    '-- and inheritance factors (arvsvinstfaktor)
    Dim maxyear As Long
    
    Dim b(0 To 106, 1 To 2) As Double
    Dim b2(0 To 106, 1 To 2) As Double
    
    Dim q(0 To 106, 1 To 2) As Double
    Dim q2(0 To 106, 1 To 2) As Double
    
    Dim lx(0 To 106, 0 To 2) As Double  '0 as Unisex
    Dim lx2(0 To 106, 0 To 2) As Double
    
    Dim lx_(0 To 2, 0 To 106) As Double 'NOTE by sex age
    Dim lx2_(0 To 2, 0 To 106) As Double
    
    Dim sex As Long, age As Long, N As Long, x As Long, k As Long, j As Long
    Dim pop As Double, pop2 As Double, pop3 As Double, d As Double, e As Double, e2 As Double, r, r2 As Double
'    Dim txt_out1 As String, txt_out2 As String, txt_out3 As String, txt_out4 As String
    
    Dim norm, normpp, normpp2 As Double
    'Net rate for DC occupational pension plans and private pension savings. 15% tax rate on dividends.
    norm = 1.016
    normpp = 1.0165
    'Driftaavdrag och Förskottsräntan för fondförsäkran (traditionell)
    'pedal
    '2001-01-01 0,300 4,000 (0,3 4,00)
    '2002-12-01 0,300 3,000 (0,3 3,00)
    '2007-04-01 0,100 4,000 (0,1 2,30)
    '2014-03-01 0,100 3,000 (0,1 3,00)
    '2017-12-01 0,100 1,750 (0,1 1,75)
    normpp2 = 1 + (normpp - 1) * (1 - 0.15)
    
    Dim year, riktage As Integer
    year = cohort ' +age
    riktage = 65 'Riktage
    
    maxyear = mini(Maxy, year)
    Dim q_lag As Double, q_lag2 As Double, q_lag3 As Double, jj As Integer
    
    For sex = 1 To 2
        pop = 100000
        pop2 = 100000
        
        For age = 0 To 106
            year = cohort + age
            If year >= 2023 Then riktage = 66
            If year >= 2026 Then riktage = 67
            If year >= 2026 Then riktage = 67
            If year >= 2034 Then riktage = 68
            If year >= 2049 Then riktage = 69
            
            If age = 0 Then
                q_lag = 0
                q_lag2 = 0
            Else
                q_lag = deathprobs(mini(Maxy, maxi(Miny, cohort + riktage - 3)), sex, mini(100, age - 1))
                q_lag2 = deathprobs(mini(Maxy, maxi(Miny, cohort + age - 2)), sex, mini(100, age - 1))
                'Income pension 5 years life tables
                For jj = 1 To 4
                    q_lag = q_lag + deathprobs(mini(Maxy, maxi(Miny, cohort + riktage - 2 - jj)), sex, mini(100, age - 1))
                Next jj
                q_lag = q_lag * (1 / 5)
            End If

            pop = pop * (1 - q_lag)
            'If cohort = 1958 And age > 65 And age < 70 Then Debug.Print age; pop; q_lag
            pop2 = pop2 * (1 - q_lag2)
            b(age, sex) = pop
            b2(age, sex) = pop2

        Next age
    Next sex
    
    
    Const beta0 As Double = 0.9  'It may have change over time
    For sex = 1 To 2
        For age = 0 To 106
             '0 åringar tas 10% från B(0,sex) och 90% från(1,age)?
            If age = 0 Then
                lx(0, sex) = beta0 * b(0, sex) + (1 - beta0) * b(1, sex)
                lx2(0, sex) = beta0 * b2(0, sex) + (1 - beta0) * b2(1, sex)
            ElseIf age < 106 Then
                lx(age, sex) = b(age, sex)
                lx2(age, sex) = b2(age, sex)
            Else
                lx(age, sex) = b(age, sex)
                lx2(age, sex) = b2(age, sex)
            End If
        Next age
    Next sex
    
    For age = 1 To 106 '-- Note: One year shift i age, i.e age 0 = 1 etc.
        'Unisex 0: 0.5145 the share of new boys
        lx_(0, age) = (lx(age, 1) * 0.5145) + (lx(age, 2) * (1 - 0.5145))
        lx2_(0, age) = (lx2(age, 1) * 0.5145) + (lx2(age, 2) * (1 - 0.5145))
        
        'Male - 1
        lx_(1, age) = lx(age, 1) * 0.5145
        lx2_(1, age) = lx2(age, 1) * 0.5145
        'Female - 2
        lx_(2, age) = lx(age, 2) * (1 - 0.5145)
        lx2_(2, age) = lx2(age, 2) * (1 - 0.5145)
    Next age
  
    For N = 50 To 105 'Starting age to end age
        For sex = 0 To 2
           
            d = 0 'ip
            e = 0 'pp
            e2 = 0 'tjp
            r = 0 'e(X)
            r2 = 0 ' E(x) with cohort
            
            For k = N To 105
                'year = cohort + k
                If year < 2003 Then normpp = 1.037
                If year > 2002 And year < 2008 Then normpp = 1.027
                If year > 2007 And year < 2015 Then normpp = 1.039
                If year > 2014 And year < 2018 Then normpp = 1.029
                If year > 2017 Then normpp = 1.0165
                normpp2 = 1 + (normpp - 1) * (1 - 0.15)

            
                For x = 0 To 11 'Month
                     d = d + ((norm) ^ (-(k - N))) * _
                        (lx_(sex, k) + (lx_(sex, k + 1) - lx_(sex, k)) _
                        * (x / 12)) * (norm) ^ (-x / 12)
                        
                     e = e + (normpp) ^ (-(k - N)) * _
                        (lx2_(sex, k) + (lx2_(sex, k + 1) - lx2_(sex, k)) _
                        * (x / 12)) * (normpp) ^ (-x / 12)
                        
                     e2 = e2 + (normpp2) ^ (-(k - N)) * _
                        (lx2_(sex, k) + (lx2_(sex, k + 1) - lx2_(sex, k)) _
                        * (x / 12)) * (normpp2) ^ (-x / 12)
                                             
                     r = r + _
                        (lx_(sex, k) + (lx_(sex, k + 1) - lx_(sex, k)) _
                        * (x / 12)) '
                    r2 = r2 + _
                        (lx2_(sex, k) + (lx2_(sex, k + 1) - lx2_(sex, k)) _
                        * (x / 12))
                        
                 Next x
                 'r = r + lx_(sex, k)
''                 If cohort = 1956 And sex = 0 Then
''                 'Stop
''                 Worksheets("Mortality").Cells(k, 32) = Round(lx_(sex, k), 0)
''                 End If
            Next k
     
            dtalip(sex, N) = Round(d / (12 * lx_(sex, N)), 2)
            dtalpp(sex, N) = Round(e / (12 * lx2_(sex, N)), 2)
            dtalpp2(sex, N) = Round(e2 / (12 * lx2_(sex, N)), 2)
            ExpLife(sex, N) = Round(r / (12 * lx_(sex, N)), 2)
            ExpLife2(sex, N) = Round(r2 / (12 * lx_(sex, N)), 2)
'            If (N >= 61 And N <= 80) Then
'              txt_out1 = txt_out1 & " " & sex & " " & dtalip(sex, N)
'              txt_out2 = txt_out2 & " " & sex & " " & dtalpp(sex, N)
'              txt_out3 = txt_out3 & " " & sex & " " & dtalpp2(sex, N)
'              txt_out4 = txt_out4 & " " & sex & " " & ExpLife(sex, N)
'            End If

        
        Next sex
        
    Next N
    
    For sex = 0 To 2
        For age = 61 To 105
           
            Arvf(sex, age) = Round(lx_(sex, age) / lx_(sex, age + 1), 6)
            Arvf2(sex, age) = Round(lx2_(sex, age) / lx2_(sex, age + 1), 6)
              
           If writefile = 1 Then
                Print #filenr, cohort + age & ";" & sex & ";" & age & ";" & cohort & ";" & _
                      dtalip(sex, age) & ";" & dtalpp(sex, age) & ";" & dtalpp2(sex, age) & _
                      ";" & ExpLife(sex, age) & ";" & ExpLife2(sex, age) & ";" & Arvf(sex, age) & ";" & Arvf2(sex, age)
           End If
        Next age
    Next sex
   
End Sub


Public Function mini(a As Variant, b As Variant) As Variant
   If a < b Then
      mini = a
   Else
      mini = b
   End If
End Function


Public Function maxi(a As Variant, b As Variant) As Variant
   If a > b Then
      maxi = a
   Else
      maxi = b
   End If
End Function


Sub arv()
' Arvsvinster för IP
   Dim surv() As Variant '
   Dim dsurv() As Double
''   Dim Last As Range
''   Set Last = wsMortality.Range("AG8:Ag96") 'obs Fast range
''
        
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
                
   Dim year As Long 'Senaste året med arvsvinster...
   year = wsMortality.Range("AF6").Value + 1
   
   Dim age As Integer
   Dim kolumn, rad As Integer
   
   ReDim surv(17 To 70, year To 2100)
   ReDim dsurv(17 To 70, year To 2100)
   
   For year = year To 2100
        Application.StatusBar = year
        
        wsMortality.Range("AH6") = year
        'rad = 4
        
        kolumn = 27 + year - 2023 ' för flik arv Kolumnnr som prognosen startar från, glöm inte nedan
        
        'wsArv_ip.Cells(rad, kolumn) = year
        
        For age = 17 To 70 'Arv 1 'OBS om riktåldern har ändrats så lägg in ett värde för raden
            rad = 3 + age - 17 'i arvsfliken
            
            If wsArv_ip.Cells(rad, kolumn) > 0 Then
               dsurv(age, year) = wsMortality.Cells(8 + age - 17, 35)
               If wsArv_ip.Cells(rad, 27) > 0 Then
                    
                    surv(age, year) = dsurv(age, year) * wsArv_ip.Cells(rad, 27)
                Else
                   
                    surv(age, year) = (2 * wsArv_ip.Cells(rad - 1, kolumn) - wsArv_ip.Cells(rad - 2, kolumn))
                End If
            End If
            
        Next age
   Next year
       
    With wsArv_ip
        Dim rng As Range
        Set rng = .Range("AB3:CZ56")
        rng.Value = surv
        Set rng = Nothing
    End With
    
    
   year = wsMortality.Range("AF6").Value + 1

   Dim surv2() As Variant '
   Dim dsurv2() As Double
        
   ReDim surv2(62 To 105, year To 2100)
   ReDim dsurv2(62 To 105, year To 2100)
    
   
     For year = year To 2090
        wsMortality.Range("AH6") = year
        'rad = 4
        kolumn = 27 + year - 2023 ' OBS Kolumnstart
        'wsArv_ip.Cells(rad, kolumn) = year

        For age = 62 To 104 'Arv 2
             rad = 83 + age - 62
             
             If wsArv_ip.Cells(rad, kolumn) > 0 Then 'OBS om riktåldern har ändrats ta bort värdet för raden
                dsurv2(age, year) = wsMortality.Cells(8 + age - 17, 35)
                If wsArv_ip.Cells(rad, 27) > 0 Then
                    surv2(age, year) = dsurv2(age, year) * wsArv_ip.Cells(rad, 27)
                Else
                    surv2(age, year) = (2 * wsArv_ip.Cells(rad - 1, kolumn) - wsArv_ip.Cells(rad - 2, kolumn))
                End If
             End If
        Next age
    Next year

        
    With wsArv_ip
        
        'Dim rng As Range
        Set rng = .Range("AB83:CZ126") 'KOLLAS
        rng.Value = surv2
        Set rng = Nothing
    End With
    
   
End Sub


