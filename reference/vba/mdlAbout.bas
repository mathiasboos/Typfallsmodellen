Attribute VB_Name = "mdlAbout"
Option Explicit
Public blnShowCloseButtonAbout As Boolean
' Visar rutan Om modellen
Sub About_modell()
    blnShowCloseButtonAbout = True
    frmAbout.Show
End Sub
