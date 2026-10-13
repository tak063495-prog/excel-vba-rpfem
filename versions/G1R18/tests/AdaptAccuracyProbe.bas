Attribute VB_Name = "AdaptAccuracyProbe"
Option Explicit
Private Sub SetProbeSetting(ByVal key As String, ByVal value As Variant)
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
    err.Raise 5, "AdaptAccuracyProbe", "NO_SETTING_SLOT"
End Sub
Public Function AccuracyProbeRun(ByVal cap As Long, ByVal rounds As Long, ByVal trace As String, Optional ByVal injectFailure As Boolean = False) As String
    Dim measured As Double, audit As Double, raw() As Double
    Call RPX_Setup
    Call SetProbeSetting("TARGET", 64)
    Call SetProbeSetting("ADAPT", 1)
    Call SetProbeSetting("CYCLES", 0)
    Call SetProbeSetting("QUADRATURE", 2)
    Call SetProbeSetting("GAP", 0.000000000001)
    Call SetProbeSetting("ADAPT_MESH_POLICY", "ROBUST_REFINE")
    Call SetProbeSetting("ADAPT_ACCURACY_CAP", cap)
    Call SetProbeSetting("ADAPT_ACCURACY_ROUNDS", rounds)
    Call SetProbeSetting("HSD_SOLVE_SECONDS", 30)
    If injectFailure Then
        Call SetProbeSetting("TEST_INJECTION", 1)
        Call RPX_TestArm("accuracy_1", 5, "CONIC_NOT_CONVERGED")
    End If
    RPX_Sheet("浸透入力").Range("D4").Value2 = 0
    XLogPath = trace
    Call RPX_NativeSetMode("auto")
    Call RPX_Run
    If Len(RPX_LastError) > 0 Then err.Raise 5, "AdaptAccuracyProbe", RPX_LastError
    If Not RPX_AdaptPairComplete() Then err.Raise 5, "AdaptAccuracyProbe", "INCOMPLETE_PAIR"
    If RBestLower.value <= 0# Or RBestUpper.value < RBestLower.value Then err.Raise 5, "AdaptAccuracyProbe", "INVALID_BOUNDS"
    If Not RBestLower.complete Or Not RBestUpper.complete Then err.Raise 5, "AdaptAccuracyProbe", "INCOMPLETE_WITNESSES"
    If RPX_TargetMet("Fs", RBestLower.value, RBestUpper.value, True) Then err.Raise 5, "AdaptAccuracyProbe", "UNMET_TARGET_WAS_ACCEPTED"
    If cap > 64 Then
        If RBestLower.elements > cap Or RBestUpper.elements > cap Then err.Raise 5, "AdaptAccuracyProbe", "ACCURACY_CAP_EXCEEDED"
    End If
    Call RPX_RestoreWitnessState(RBestLower)
    If Not RPX_FieldGate("lower", RBestLower.field, RBestLower.value, RBestLower.q, RBestLower.objective, measured, audit, raw) Then err.Raise 5, "AdaptAccuracyProbe", "LOWER_WITNESS_GATE_FAILED"
    Call RPX_RestoreWitnessState(RBestUpper)
    If Not RPX_FieldGate("upper", RBestUpper.field, RBestUpper.value, RBestUpper.q, RBestUpper.objective, measured, audit, raw) Then err.Raise 5, "AdaptAccuracyProbe", "UPPER_WITNESS_GATE_FAILED"
    AccuracyProbeRun = "PASS;lower=" & RBestLower.value & ";upper=" & RBestUpper.value & ";lower_elements=" & RBestLower.elements & ";upper_elements=" & RBestUpper.elements & ";current_elements=" & re & ";incomplete=" & RFsSearchIncomplete & ";physical_gates=PASS"
End Function
