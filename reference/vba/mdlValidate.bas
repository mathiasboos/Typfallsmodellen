Attribute VB_Name = "mdlValidate"
Option Explicit

Sub ValidateCellValue()
    Dim rng_Yearly_Inflation As Range
    Dim rng_FondAvkastning As Range
    Dim rng_Real_Growth As Range
    
    Dim str_rng_Yearly_Inflation As String
    Dim str_rng_FondAvkastning As String
    Dim str_rng_Real_Growth As String

    Set rng_Yearly_Inflation = wsStart.Range("rng_Yearly_Inflation")
    Set rng_FondAvkastning = wsStart.Range("rng_FondAvkastning")
    Set rng_Real_Growth = wsStart.Range("rng_Real_Growth")
    
    str_rng_Yearly_Inflation = Format(rng_Yearly_Inflation, "0.00%")
    str_rng_FondAvkastning = Format(rng_FondAvkastning, "0.00%")
    str_rng_Real_Growth = Format(rng_Real_Growth, "0.00%")
    
    'Validerar värden av Inflation, Fondavkastning och realtillväxt
    If (rng_FondAvkastning < 0 And rng_Real_Growth > 0) Or (rng_Yearly_Inflation > 0 And rng_Real_Growth < 0) Or _
        (rng_Real_Growth < 0 And rng_FondAvkastning > 0) Then
    
        MsgBox fngetMsgBoxtext(2) & vbCrLf & fngetMsgBoxtext(4) & " " & str_rng_Yearly_Inflation & ", " & fngetMsgBoxtext(5) & " " & str_rng_FondAvkastning _
            & ", " & fngetMsgBoxtext(6) & " " & str_rng_Real_Growth & "." & fngetMsgBoxtext(3)

    End If
    Set rng_Yearly_Inflation = Nothing
    Set rng_FondAvkastning = Nothing
    Set rng_Real_Growth = Nothing
End Sub

Function fnValidateRegion(ByVal StartRow As Integer, ByVal StartCol As Integer, ByVal StopRow As Integer, ByVal StopCol As Integer, wkName As String) As Boolean
    Dim rngValidate As Range
    Dim c As Range
    Dim strMessageRow As String
    Dim strMessage As String
    Dim ActRow As Integer
    Dim i As Integer

    fnValidateRegion = True
    With Worksheets(wkName)
        Set rngValidate = .Range(.Cells(StartRow, StartCol), .Cells(StopRow, StopCol))
        For Each c In rngValidate
            If Len(Trim(c)) = 0 And ActRow <> c.Row Then
                If i < 3 Then
                 strMessageRow = strMessageRow & ", " & c.Row
                End If
                ActRow = c.Row
                i = i + 1
            End If
        Next c
        If i < 4 Then
            If i = 1 Then
                strMessage = "Fel vid " & i & " rad: "
            Else
                strMessage = "Fel vid " & i & " rader: "
            End If
        Else
            strMessage = "Fel vid " & i & " rader, bland annat: "
        End If
        If Len(strMessageRow) > 0 Then
            MsgBox (strMessage & Mid(strMessageRow, 2, Len(strMessageRow) - 1))
            fnValidateRegion = False
        End If
        Set rngValidate = Nothing
    End With

End Function
