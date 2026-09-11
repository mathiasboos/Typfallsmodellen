Attribute VB_Name = "frmSettings"
Attribute VB_Base = "0{8FE63BA8-5C3C-47C0-98AF-486ED34BB78F}{50069CF8-393B-44CC-BA21-4132F04FE4DD}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
'Removed 2024-01-23 due to wsStart.optNormal did not exist.
'Private Sub chkAdvanced_setting_Click()
'    If wsStart.optNormal Then
'        If chkAdvanced_setting.Value = True Then
'            MsgBox fngetMsgBoxtext(36)
'            chkAdvanced_setting.Value = False
'            Exit Sub
'        End If
'    Else
'        If chkAdvanced_setting = True Then
'            wsAdv_Setting.Select
'            Exit Sub
'        End If
'    End If
'    If chkAdvanced_setting = False Then wsStart.Activate
'End Sub

Private Sub chkOwnVectors_Click()
    fnGetRangeFromNames(ThisWorkbook, "rng_Egen_Lon") = chkOwnVectors
    Call OwnVectors
    If Not fnGetRangeFromNames(ThisWorkbook, "rng_Egen_Lon") Then wsStart.Select
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdHideSheets_Click()
    Application_Rest
    
    Call User_mode
    fnGetRangeFromNames(ThisWorkbook, "rng_Application_Mode") = 0
    Call Show_Hide_Chart_Priser(False)
    Call MonthCellsFontColor
    
    Application_Wakeup
End Sub
Private Sub cmdMarried_Click()
    fnGetRangeFromNames(ThisWorkbook, "Gift") = cmdMarried.Value
    If fnGetRangeFromNames(ThisWorkbook, "Gift") Then
        Call UnHideControll(Me, txtIncome.Name)
        Call UnHideControll(Me, lblPartner.Name)
    Else
        Call HideControll(Me, txtIncome.Name)
        Call HideControll(Me, lblPartner.Name)
    End If
End Sub
Private Sub cmdShowSheets_Click()
    Application_Rest
    
    Call Advanced_mode
    fnGetRangeFromNames(ThisWorkbook, "rng_Application_Mode") = 1
    Call MonthCellsFontColor
    Call Show_Hide_Chart_Priser(True)
    
    Application_Wakeup
End Sub

Private Sub txtIncome_AfterUpdate()
    If Len(Trim(Me.txtIncome)) = 0 Then
        fnGetRangeFromNames(ThisWorkbook, "rng_Makens_inkomst") = Me.txtIncome
    ElseIf IsNumeric(Me.txtIncome) Then
        fnGetRangeFromNames(ThisWorkbook, "rng_Makens_inkomst") = Me.txtIncome * 1
    End If
End Sub


Private Sub UserForm_Activate()
    Me.chkOwnVectors = fnGetRangeFromNames(ThisWorkbook, "rng_Egen_Lon")
    TranslateFormsControls Me
    Me.Caption = fngetMsgBoxtext(32)
    Me.MultiPage1.Pages("Adv_Settings").Caption = fngetMsgBoxtext(33)
    Me.cmdMarried = fnGetRangeFromNames(ThisWorkbook, "Gift")

    If fnGetRangeFromNames(ThisWorkbook, "rng_Application_Mode") = 1 Then
        Call UnHideControll(Me, chkAdvanced_setting.Name)
        Call UnHideControll(Me, chkOwnVectors.Name)
        Call UnHideControll(Me, cmdMarried.Name)
        Call HideControll(Me, lblModellLimit.Name)
       
        If fnGetRangeFromNames(ThisWorkbook, "Gift") Then
            Call UnHideControll(Me, txtIncome.Name)
            Call UnHideControll(Me, lblPartner.Name)
            Me.txtIncome = fnGetRangeFromNames(ThisWorkbook, "rng_Makens_inkomst")
        Else
            Call HideControll(Me, txtIncome.Name)
            Call HideControll(Me, lblPartner.Name)
        End If
    Else
        Call HideControll(Me, chkAdvanced_setting.Name)
        Call HideControll(Me, chkOwnVectors.Name)
        Call HideControll(Me, cmdMarried.Name)
        Call UnHideControll(Me, lblModellLimit.Name)
        Call HideControll(Me, txtIncome.Name)
        Call HideControll(Me, lblPartner.Name)
    End If
    
End Sub
Private Sub UserForm_Initialize()
    TranslateFormsControls Me
End Sub



