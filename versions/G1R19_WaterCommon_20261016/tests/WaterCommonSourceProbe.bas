Attribute VB_Name = "WaterCommonSourceProbe"
Option Explicit
' Test-copy only, after a valid horizontal D08 seep result is populated.
' No solve or mesh alteration. Import this helper separately from the overlay.
Public Function WaterCommonSourceMutationProbe(ByVal target As String) As String
    Dim ws As Worksheet, cell As Range, saved As Variant, savedFormat As Variant
    Dim level As Double, waterGamma As Double, n As Long, message As String, oldCancel As Boolean
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets("浸透入力")
    Select Case target
        Case "gamma": Set cell = ws.Range("B4")
        Case "gamma_type": Set cell = ws.Range("B4")
        Case "coupling": Set cell = ws.Range("D4")
        Case "suction": Set cell = ws.Range("F4")
        Case "mode": Set cell = ws.Range("B3")
        Case "table_x": Set cell = ws.Range("M7")
        Case "table_y": Set cell = ws.Range("N7")
        Case "result_id": Set cell = ThisWorkbook.Worksheets("浸透結果").Range("K9")
        Case "result_head": Set cell = ThisWorkbook.Worksheets("浸透結果").Range("D2")
        Case "result_node": Set cell = ThisWorkbook.Worksheets("浸透結果").Range("B2")
        Case "result_triangle": Set cell = ThisWorkbook.Worksheets("浸透結果").Range("S2")
        Case Else: err.Raise 5, , "UNKNOWN_WATER_SOURCE_PROBE"
    End Select
    saved = cell.formula: savedFormat = cell.NumberFormat: oldCancel = XCancel
    XCancel = False: Call RPX_InputBegin
    If Not RPX_WaterCommonLevel(level, waterGamma) Then err.Raise 5, , "PROBE_REQUIRES_ELIGIBLE_COMMON_SOIL"
    Select Case target
        Case "gamma_type": cell.NumberFormat = "@": cell.Value2 = CStr(cell.Value2)
        Case "coupling": cell.Value2 = 0#
        Case "suction": cell.Value2 = 1#
        Case "mode", "result_id": cell.Value2 = CStr(cell.Value2) & ":changed"
        Case Else: cell.Value2 = CDbl(cell.Value2) + 0.001
    End Select
    On Error Resume Next
    RPX_AssertInputs "water_source_probe"
    n = err.number: message = err.description: err.Clear
    On Error GoTo Failed
    If n <> 5 Or InStr(message, "WATER_COMMON_SOURCE_CHANGED") = 0 Or Not XCancel Then err.Raise 5, , "SOURCE_CHANGE_NOT_REJECTED " & target
    Call RPX_InputEnd
    cell.NumberFormat = savedFormat: cell.formula = saved: XCancel = False
    ' The same restored source must bind successfully in a fresh epoch.
    Call RPX_InputBegin: Call RPX_AssertInputs("water_source_newrun")
    Call RPX_InputEnd: XCancel = oldCancel
    WaterCommonSourceMutationProbe = "PASS " & target
    Exit Function
Failed:
    message = err.description
    On Error Resume Next
    Call RPX_InputEnd
    If Not cell Is Nothing Then cell.NumberFormat = savedFormat: cell.formula = saved
    XCancel = oldCancel
    WaterCommonSourceMutationProbe = "FAIL " & target & " " & message
End Function
