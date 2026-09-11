Attribute VB_Name = "Labour_S"
'Arbetsutbud
Option Explicit

Sub BR()
    Call startsetup
    Const wage = 500
    'constant marginal = 0
    
    Dim h As Single
    h = 0
    Dim V As Double
    Worksheets("Start").Select
    V = Application.Range("d36")
    Dim age As Single
    age = Application.Range("paryear")
    
    Dim kostnadsavd, Grundavdrag, jobbavdrag As Double
    Dim brutto, ctxfvi, cbefvi, pensionavgift, pensredukt, kinkskatt, kyrkskatt, statskatt As Double
    
    kostnadsavd = 0
  
  Dim Mvalues_() As Double
  ReDim Mvalues_(0 To 40, 1 To 2)
  For h = 0 To 40
        'Taxerad förvärvsinkomst - Ctxfvi
        brutto = V + h * wage - kostnadsavd
        If marginal = 0 Then
            ctxfvi = Int(brutto / 100) * 100
        End If
        
        Grundavdrag = avdragxx(ctxfvi, pbb(age), marginal, age, year_(age), 2100, IBB(age), 1, year_(age))
              
          'PGI avgift från skattskyldig
        pensionavgift = pgi(year_(age), h * wage, pbb(age), IBB(age), FPB(age), marginal, 1, age, 0)
        
        'Beskattningsbar inkomst
        cbefvi = ctxfvi - Grundavdrag - pensionavgift + pgi(year_(age), h * wage, pbb(age), IBB(age), FPB(age), marginal, 2, age, 0)
        
        'Kommunal inkomstskatt
        kinkskatt = cbefvi * Kom_skatt(age)
        If marginal = 0 Then kinkskatt = Int(kinkskatt)
        
        jobbavdrag = Jobbxx(Wage_(age), age, Kom_skatt(age), pbb(age), marginal, ctxfvi - Wage_(age), year_(age), 2199, IBB(age), 1, year_(age))
  
        'Agift till Sv kyrkan eller annat trossamfund
        kyrkskatt = cbefvi * Begravavg(age)
        If marginal = 0 Then kyrkskatt = Int(kyrkskatt)
        'Reduktion pensionsavgiften - funktionen koll för rulesfrom och enbart red. som ev. påverkas
        pensredukt = pgi(year_(age), Income_(age), pbb(age), IBB(age), FPB(age), marginal, 2, age, 0)
        
        'Statlig inkomstskatt och
        'public service avgiften inlagd i statsskatten
        statskatt = statlig(cbefvi, Tax_limit1(age), Tax_limit2(age), marginal)
            statskatt = statskatt + PublicAvg(cbefvi, 0.01, age, marginal, year_(age))
          'Nettoinkomst
         Dim Netto As Double
         
        Netto = brutto - maxi(kinkskatt + kyrkskatt + statskatt + pensionavgift - pensredukt - jobbavdrag, 0)
        Mvalues_(h, 1) = 40 - h
        'Mvalues_(h, 2) = brutto / 12
         Mvalues_(h, 2) = Netto / 12
         
        
        'Debug.Print brutto, netto
    Next h
    'Skriv ut BR
    Range("t80:u" & 80 + 40).Value = Mvalues_
    
End Sub

