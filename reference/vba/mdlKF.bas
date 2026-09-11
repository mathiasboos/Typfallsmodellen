Attribute VB_Name = "mdlKF"
Sub KF_Calculation()
     If Application.Range("rng_Show_KF") Then
        Call ShowWorksheet(wsInUtdata.Name)
        wsInUtdata.Activate
        Application.Range("rng_Startflik") = False
        Call HideWorksheet(wsBrutto.Name)
     End If
End Sub
Sub ShowWorksheet(Sheetname As String)
    Sheets(Sheetname).Visible = xlSheetVisible
    Sheets(Sheetname).Activate
End Sub
Sub HideWorksheet(Sheetname As String)
    Sheets(Sheetname).Visible = xlSheetHidden
End Sub

Sub ToStartSheet()
    Call HideWorksheet(wsInUtdata.Name)
    Call HideWorksheet(wsBrutto.Name)
    Application.Range("rng_Show_KF") = False
    wsStart.Activate
End Sub
