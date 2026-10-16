Attribute VB_Name = "SettingsExtentFix4Probe"
Option Explicit

Public Function ExtentReadInputsError() As String
    On Error GoTo Rejected
    Call RPX_ReadInputs
    ExtentReadInputsError = "OK"
    Exit Function
Rejected:
    ExtentReadInputsError = CStr(Err.Number) & "|" & Err.Description
End Function
