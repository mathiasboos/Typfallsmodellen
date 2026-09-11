Attribute VB_Name = "wsEgen_standard_Typfall"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
' Lagra/uppdatera egen standard typfall från start fliken
Sub UpdateList()
    Dim rng As Range
    Set rng = fnGetRangeFromNames(ThisWorkbook, "rng_Top_Left_EgentypfallList").Offset(1)
    Range("A3:A10000").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=rng, Unique:=True
    Set rng = Nothing
End Sub
