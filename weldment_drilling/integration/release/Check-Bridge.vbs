Option Explicit
Dim bridge
On Error Resume Next
Set bridge = CreateObject("Macro.DrillingDimensions")
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Bridge activation failed: " & Err.Description
    WScript.Quit 1
End If
On Error GoTo 0
If bridge.BuildId <> "DRILLING_DIMENSIONS_BRIDGE_R2_20261029" Then
    WScript.Echo "ERROR: Installed bridge version does not match this release."
    WScript.Quit 1
End If
WScript.Echo "Bridge activation and BuildId verified. SOLIDWORKS geometry still requires a native test."
WScript.Quit 0
