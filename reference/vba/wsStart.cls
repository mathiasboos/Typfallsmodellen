Attribute VB_Name = "wsStart"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
' Start flik för PETRA


'Om startflikens ingångsvärden ändras så rensa utdata och skydda arket.



Private Sub Worksheet_Change(ByVal Target As Range)
    If Application.Range("rng_Run_From_Indata") Then Exit Sub
    If blnGetTyfall Then Exit Sub
   
    If blnOpeningProgress Then Exit Sub ' körs vid öppning av filen när bladet töms
    If Target.Address = Range("BornYear").Address Or _
        Target.Address = Range("PARYear").Address Or _
        Target.Address = Range("wStartYear").Address Or _
        Target.Address = Range("Wage_Monthly").Address Or Target.Address = Range("rng_Yearly_Inflation").Address Or _
        Target.Address = Range("rng_FondAvkastning").Address Or _
        Target.Address = Range("rng_Real_Growth").Address Then
        
        If wsNameRange.Range("rng_OpeningProgress") Then wsNameRange.Range("rng_OpeningProgress") = False
        On Error Resume Next
        Application.ScreenUpdating = False
        
        If Target.Address = Range("BornYear").Address And Application.Range("rng_Riktålder") Then 'Hämta riktålder (eller alternativ p-ålder om rng_Rek_Pens_Ålder här byts ut) se Range("rng_Nyckeltal_Pensyear")
            riktalder (Target)
        End If
       
        Application.ScreenUpdating = True
        If Application.Range("rngNeedsRecalculation") Then
            CleanTabel_Start 'Se nedan
        End If
        On Error GoTo 0
    End If
    
End Sub
Sub CleanTabel_Start()
    Dim rngTop As Range
    Application_Rest
    wsStart.Unprotect
    Set rngTop = Application.Range("rng_Top_tabell1") 'Startfliken
    Set rngTop = rngTop.Resize(13, 4) 'OBS fast adress utökar med 13 rader och 4 kolumner
    rngTop.ClearContents
    
    Set rngTop = Application.Range("rng_Top_tabell1").Offset(18) 'Gå ner 17+1 rader
    Set rngTop = rngTop.Resize(4, 4)
    rngTop.ClearContents
    
    Set rngTop = Application.Range("rng_resp_utdata").Offset(2, 1) 'Gå ner 2 rader 1 kolumn
    Set rngTop = rngTop.Resize(3, 2)
    rngTop.ClearContents
    
    Set rngTop = Application.Range("rng_resp_utdata").Offset(12, 1) 'Gå ner 12 rader 1 kolumn
    Set rngTop = rngTop.Resize(3, 2)
    rngTop.ClearContents
    
    'Rensa även utdata som ligger till grund för Figur 1 och 2:
    Set rngTop = Application.Range("Rng_res_utdata") 'Fliken Utdata
    Set rngTop = rngTop.CurrentRegion
    rngTop.ClearContents
        
    Set rngTop = Application.Range("rng_tabell2_top")
    Set rngTop = rngTop.Offset(2).Resize(110, rngTop.CurrentRegion.Columns.Count)
    rngTop.Clear
    
    Set rngTop = Application.Range("rng_tabell2_top")
    Set rngTop = rngTop.Offset(1).Resize(Application.Range("rng_tabell2_end").Row - Application.Range("rng_tabell2_top").Row + 1, rngTop.CurrentRegion.Columns.Count)
    rngTop.ClearContents
    Set rngTop = Nothing
     ' Flyttar namnområdet rng_tabell2_top till cell K79
    Dim CellName As Name
    Set CellName = Application.Names.Item("rng_tabell2_end")
    CellName.RefersTo = "=Start!$K$78" 'OBS fast adress
    Application_Wakeup
    If Application.Range("Adv_unprotect") = 1 Then
    Else
        wsStart.Protect
    End If
End Sub
Sub SheetOrdlista()
    If Worksheets(wsOrdlista.Name).Visible = True Then
        MoveToSheet (wsOrdlista.Name)
    End If
End Sub
