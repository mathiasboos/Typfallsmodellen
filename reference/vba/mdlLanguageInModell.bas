Attribute VB_Name = "mdlLanguageInModell"

Option Explicit


Sub Lang_Swe()
    Change_Language (0)
End Sub
Sub Lang_Eng()
    Change_Language (1)
End Sub
Sub Change_Language(ByVal iLan As Integer)
    Dim cBut As CommandBarButton
    Dim cBut1 As CommandBarButton
    Dim i As Integer
    Dim iCol As Integer
    Dim rng As Range
    If IsNull(iLan) Or Not IsNumeric(iLan) Then iLan = 0
    Application.Range("sysLang") = iLan
    
    Set rng = Application.Range("sysLang")
    iCol = rng.CurrentRegion.Columns.Count - 1
    Set rng = Nothing
    On Error Resume Next
         With Application
            'Ta bort höger meny förbredd för 6 språk
            For i = 0 To iCol
                If i = Application.Range("sysLang") Or Len(fngetMsgBoxtextByColumn(48, i)) = 0 Then
                
                Else
                    .CommandBars("Cell").Controls(fngetMsgBoxtextByColumn(48, i)).Delete
                    Set cBut = .CommandBars("Cell").Controls.Add(Temporary:=True)
                    Set cBut1 = .CommandBars("Cell").Controls.Add(Temporary:=True)
                End If
            Next i
            .CommandBars("Cell").Controls("Data").Delete
            .CommandBars("Cell").Controls("").Delete
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
    'Call FixControllSize
    Call Translate_Forms
    Call ShapeTranslate
End Sub
