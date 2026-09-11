Attribute VB_Name = "mdlCal"
Option Explicit
Public pblnDisableListClick As Boolean
' Startar upp kalkuleringen
Sub Show_Button_Caculate()
    Call Unprotect_sheet(wsStart.CodeName)
    If Application.Range("rngNeedsRecalculation") = True Then
        With wsStart.Shapes("cmd_Cal")
            .Visible = msoTrue
        End With
    Else
         With wsStart.Shapes("cmd_Cal")
            .Visible = msoFalse
        End With
    End If
    Call Protect_sheet(wsStart.CodeName)
End Sub
Function CellName(ByVal Indexcell As Range)
    CellName = Indexcell.Name.Name
End Function
Sub ResetStart()
        Dim LB
        Application_Rest
        Call Unprotect_sheet(wsStart.CodeName)
        Set LB = wsStart.Shapes("lstVal_of_Tjenste")
        Application.Range("BornYear") = ""
        Application.Range("PARYear") = ""
        Application.Range("wStartYear") = ""
        
        'Raden nedan bortkommenterad av TL
''        If Application.Range("rng_Application_Mode") = 1 And Application.Range("rng_MonthCal") Then _
''            Application.Range("PARMonth") = Application.Range("rng_MånaderIndexTop")
        
        Application.Range("Wage_Monthly") = ""
        Application.Range("IPS_Monthly") = 0
        Application.Range("rng_OpeningProgress") = True
        With LB.ControlFormat
            .ListIndex = 1
        End With
       
        Application.Range("rng_Egen_Lon") = False
        Application.Range("rngSenasteStartChangeCells") = Application.Range("rngCurrentStartChangeCells")
        wsStart.Activate
        Set LB = Nothing
        Call Protect_sheet(wsStart.CodeName)
        
        Application_Wakeup
End Sub




