Attribute VB_Name = "wsIndata"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
' Mikro simulering
'Makrot nedan ignorera altenative pensionsålder klickboxen
Private Sub Worksheet_Change(ByVal Target As Range)
   If Application.Range("rng_Run_From_Indata") Then Exit Sub
   If Selection.Cells.Count > 1 Then Exit Sub
   Dim r As Range
   Set r = Target
   If Target.Column = 2 And Target.Row > Application.Range("rngXTopleft").Row And Target.Cells.Count = 1 And Application.Range("rng_Riktålder") Then
        Call Riktalder_Indata(Target.Value, Target)
   End If
   Set r = Nothing
   
    Dim rngXdta As Range
    Dim rng As Range
        'Define X-data range (input)
    If Target.Cells.Count > 1 Then Exit Sub
    On Error Resume Next
    Set rngXdta = Application.Range("rngXTopleft")
    Set rngXdta = rngXdta.CurrentRegion
    Set rngXdta = rngXdta.Resize(1, rngXdta.Columns.Count)
    
    Set rng = Application.Intersect(rngXdta, Target)
'    If Not rng Is Nothing Then
'        subManageNewValidationRanges rng
'    End If
    
    'fnIsAllIndataOK True
    
    Set rng = Nothing
    Set rngXdta = Nothing
    
End Sub
