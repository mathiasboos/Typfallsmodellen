Attribute VB_Name = "Module2"
Option Explicit

Function HardwareAccelerationDisabled() As Long
    'Returns True (-1) if setting is disabled, False (0) if not disabled, or 99 if couldn't access the registry key
    On Error GoTo HardwareAccelerationDisabled_Error
    If Left(Application.Version, 2) = "14" Then 'For Excel 2010
        HardwareAccelerationDisabled = (1 = CreateObject("WScript.Shell").RegRead("HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Gfx\DisableHardware"))
    Else 'For Excel 2013+
        HardwareAccelerationDisabled = (1 = CreateObject("WScript.Shell").RegRead("HKEY_CURRENT_USER\Software\Microsoft\Office\" & Application.Version & "\Common\Graphics\DisableHardwareAcceleration"))
    End If
    On Error GoTo 0
    Exit Function

HardwareAccelerationDisabled_Error:
    HardwareAccelerationDisabled = "99"
End Function
