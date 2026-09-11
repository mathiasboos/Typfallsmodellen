Attribute VB_Name = "mdlProgress"
Option Explicit

Sub UpdateProgress(pct)
    If PctDone > 1 Then PctDone = 1
    With UserFormProgress
        .FrameProgress.Caption = Format(pct, "0%")
        .LabelProgress.Width = pct * (.FrameProgress.Width - 10)
        .Repaint
    End With
End Sub

Sub ShowUserForm(r As Integer, G As Integer, b As Integer)
    With UserFormProgress
        .LabelProgress.BackColor = RGB(r, G, b)
        .LabelProgress.Width = 0
        .Show
    End With
End Sub
