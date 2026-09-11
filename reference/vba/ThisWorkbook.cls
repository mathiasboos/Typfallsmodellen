Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
'This workbook vad som ska gälla vid uppstart mm för PETRA

Option Explicit
'Dim lngWbCalcStatus As Long

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    pblnCloseOrSave = True
    'fnGetRangeFromNames(ThisWorkbook, "rng_Change_button_color") = False
    pblnDisableListClick = True
    On Error Resume Next
        With Application
            'Tidigare egna definierade rader vid högerklick tas bort
            .CommandBars("Cell").Controls(fngetMsgBoxtext(48)).Delete
            .CommandBars("Cell").Controls("Data").Delete
            .CommandBars("Cell").Controls("").Delete
        End With
    On Error GoTo 0
End Sub
Sub test()
    Debug.Print pblnCloseOrSave
End Sub
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    pblnCloseOrSave = True
    'fnGetRangeFromNames(ThisWorkbook, "rng_Change_button_color") = False
    pblnDisableListClick = True
End Sub

'Om ett nytt excelark öppnas och det ska vara aktivt...
Private Sub Workbook_Deactivate()
    On Error Resume Next
        With Application
            .CommandBars("Cell").Controls(fngetMsgBoxtext(48)).Delete
            .CommandBars("Cell").Controls("Data").Delete
            .CommandBars("Cell").Controls("").Delete
        End With
    On Error GoTo 0
End Sub

'Ta bort blanka rader som ev. finns i listan vid högerklickande
Private Sub Delete_Blank_Controlls()
    On Error Resume Next
            With Application
                .CommandBars("Cell").Controls("").Delete
            End With
    On Error GoTo 0
End Sub

'När PETRA öppnas ska typfallsmodellen...
Private Sub Workbook_Open()

    If fnOpenFromZip Then
        'Går den inte att köra från zippad stäng
        EndApplicationIfOnlyOneWorkbookOpen
    End If

    If ThisWorkbook.ReadOnly Then
        'Samma här men ge ett felmedddelande se flik sysmsgboxlang
        MsgBox fngetMsgBoxtext(20), vbCritical
        EndApplicationIfOnlyOneWorkbookOpen
    End If
    
    'Dim myear As Integer
    Dim uyear As Long
    'myear = Application.Range("Modell_year")
    uyear = Application.Range("last_update") 'Adv_settings c1
    If Date > (uyear + 365) And Application.Range("rng_ShowUpdateModell") = True Then
        'MsgBox ("Modellen uppdateras senast " & Application.Range("last_update"))
        frmUpdateModell.Show
    End If
    Application.StatusBar = ""
    Application_Rest 'Skärmen behöver inte uppdateras, se modul mdlTools

    blnOpeningProgress = True
    pblnCloseOrSave = False
    pblnDisableListClick = True
    blnGetTyfall = False
    Call Unprotect_sheet(wsStart.CodeName) 'Oskyddar arket, se modul mdlTools

''''    On Error Resume Next
''''    Call FixControllSize
''''    'Call Fix_Chart_Start
''''    On Error GoTo 0

    pblnDisableListClick = False
    If Not fnCheckAbout Then frmAbout.Show 'Show about the model...


    'Call ResetStart 'rensar gamla sparade värden
    Call MonthCellsFontColor
    'Call get_economical_Default_Value 'tl nej dito

    Call Protect_sheet(wsStart.CodeName)
    Application_Wakeup 'Skärmuppdatering

    blnOpeningProgress = False
    fnGetRangeFromNames(ThisWorkbook, "rng_OpeningProgress") = True
    If Application.Range("rng_Show_KF") Then
        wsInUtdata.Visible = xlSheetVisible
        wsInUtdata.Activate
    Else
        wsBrutto.Visible = xlSheetHidden
        wsInUtdata.Visible = xlSheetHidden
        wsStart.Activate
    End If
    
    'fnGetRangeFromNames(ThisWorkbook, "Value_LastRun").Value = ""
    
End Sub


Sub EndApplicationIfOnlyOneWorkbookOpen()
    'Quits excel if only this workbook is open, else abort current operation
    'On Error Resume Next
   ' ThisWorkbook.Saved = True  'kollas
'    If Application.Workbooks.Count = 1 Then
'        Application.Quit
'    Else
        On Error Resume Next
        ThisWorkbook.Close False
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        End
'    End If
End Sub


Private Sub Workbook_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    'Lägger till kommandona  att gå till flken "start" och "utdata" vid högerklick
    Dim cBut As CommandBarButton
    Dim cBut1 As CommandBarButton
    
    On Error Resume Next
         With Application
            .CommandBars("Cell").Controls(fngetMsgBoxtext(48)).Delete
            .CommandBars("Cell").Controls("Data").Delete
            Set cBut = .CommandBars("Cell").Controls.Add(Temporary:=True)
            Set cBut1 = .CommandBars("Cell").Controls.Add(Temporary:=True)
        End With
         With cBut
           .Caption = fngetMsgBoxtext(48)
           .Style = msoButtonCaption
           .OnAction = "GoToStart"
        End With
        With cBut1
           .Caption = "Data"
           .Style = msoButtonCaption
           .OnAction = "GoToData"
        End With
    Set cBut = Nothing
    Set cBut1 = Nothing
    On Error GoTo 0
End Sub
