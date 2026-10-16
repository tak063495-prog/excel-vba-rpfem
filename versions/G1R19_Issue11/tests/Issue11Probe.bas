Attribute VB_Name = "Issue11Probe"
Option Explicit
' Issue #11: an eligible q8 pair restored by AfterCycles (raw phase) must be
' reused before the final search; a q4 coarse pair must still be re-searched.
Private Sub Setting(ByVal key As String, ByVal value As Variant)
    Dim ws As Worksheet, r As Long
    Set ws = RPX_Sheet("設定")
    For r = 4 To 1000
        If CStr(ws.Cells(r, 4).Value2) = key Then ws.Cells(r, 2).Value2 = value: Exit Sub
    Next r
    For r = 4 To 1000
        If Len(CStr(ws.Cells(r, 4).Value2)) = 0 And Len(CStr(ws.Cells(r, 2).Value2)) = 0 Then
            ws.Cells(r, 4).Value2 = key: ws.Cells(r, 2).Value2 = value: Exit Sub
        End If
    Next r
    err.Raise 5, "Issue11Probe", "NO_SETTING_SLOT"
End Sub
' fromStart=1: q8 pair from the first evaluation (eligible).
' fromStart=0: q4 coarse pair (ineligible; final q8 search is mandatory).
Public Function Issue11Case(ByVal trace As String, ByVal fromStart As Long) As String
    On Error GoTo Failed
    RPX_Setup
    Setting "MODE", "Fs": Setting "TARGET", 16: Setting "ADAPT", 1: Setting "CYCLES", 1
    Setting "QUADRATURE", 8: Setting "ADAPT_CERT_QUADRATURE", 8: Setting "ADAPT_PRECISION_FROM_START", fromStart
    Setting "RAW_LOWER_AUDIT", 1: Setting "GAP", 1.9: Setting "C0_AUTO", 0
    Setting "ADAPT_MESH_POLICY", "ROBUST_REFINE": Setting "ADAPT_ACCURACY_CAP", 16: Setting "ADAPT_ACCURACY_ROUNDS", 0
    Setting "ADAPT_LOWER_FAN_ROUNDS", 0: Setting "ADAPT_C0_WEDGE_DEPTH", 0: Setting "UPPER_PILOT", 0
    Setting "HSD_SOLVE_SECONDS", 30: Setting "TEST_INJECTION", 0
    RPX_Sheet("材料データ").Range("C5").Value2 = 0#: RPX_Sheet("材料データ").Range("E5").Value2 = 0#
    RPX_Sheet("浸透入力").Range("D4").Value2 = 0
    XLogPath = trace
    RPX_Run
    Issue11Case = "OK;error=" & RPX_LastError & ";lower=" & RPX_ExactDouble(RBestLower.value) & ";upper=" & RPX_ExactDouble(RBestUpper.value) & ";lower_q=" & RBestLower.q & ";upper_q=" & RBestUpper.q & ";phase=" & RPX_LoadPhase & ";model_match=" & IIf(RBestUpper.originalModel = RPX_ExactModel(), "1", "0")
    Exit Function
Failed:
    Issue11Case = "FAIL;number=" & err.number & ";source=" & err.source & ";message=" & err.description
End Function
