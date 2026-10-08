Option Explicit
' P4 generations are advanced even when diagnostics are disabled.
Public Type RPX_DirectionInfo
    kind As Long
    attempted As Long
    returned As Long
    matrix As Long
    solve As Long
    state As Long
    Pattern As Long
    rhs As Long
    numericValid As Boolean
    corrections As Long
    schurCorrections As Long
    checks As Long
    residual As Double
    errors(0 To 3) As Double
End Type
Public P4Trace As Long, P4Consistent As Boolean, P4Cache As Boolean, P4ReuseU As Boolean
Public P4Solve As Long, P4State As Long, P4Matrix As Long, P4Pattern As Long, P4Rhs As Long
Public P4Generic As Long, P4Current As RPX_DirectionInfo
Public P4Calls As Long, P4Replays As Long, P4Returned As Long, P4Committed As Long
Public P4Factors As Long, P4Prepares As Long, P4Hits As Long, P4Reduced As Long
Public P4Corrections As Long, P4Checks As Long, P4Attempts As Long
Public P4LocalRhs As Double, P4LocalFactor As Double
Private phaseTime(0 To 11) As Double

Private Function Choice(ByVal key As String, ByVal fallback As String, ByVal allowed As String) As String
    Choice = UCase$(Trim$(CStr(RPX_Setting(key, fallback))))
    If InStr("|" & allowed & "|", "|" & Choice & "|") = 0 Or Len(Choice) = 0 Then err.Raise 5, "RPX_P4", "INVALID_SETTING " & key
End Function
Public Sub RPX_P4Configure()
    Dim value As String
    P4Trace = CLng(Choice("SCHUR_TRACE_LEVEL", "0", "0|1|2"))
    P4Consistent = (Choice("SCHUR_REPLAY_POLICY", "LEGACY", "LEGACY|CONSISTENT") = "CONSISTENT")
    P4Cache = (Choice("SCHUR_STRUCTURE_CACHE", "0", "0|1") = "1")
    P4ReuseU = (Choice("SCHUR_RECOVERY_MODE", "RESOLVE", "RESOLVE|REUSE_U") = "REUSE_U")
    value = Choice("SCHUR_REFINEMENT", "FIXED3", "FIXED3")
    value = Choice("SCHUR_GENERIC_START_ITERS", "0", "0")
    value = Choice("DIRECT_FINAL_AFTER_PILOT", "0", "0")
End Sub
Public Sub RPX_P4Begin()
    Dim i As Long
    RPX_P4Configure
    ' Reserve more than a complete solve before any Long counter can overflow.
    If P4Solve > 2000000000 Or P4Matrix > 2000000000 Or P4Rhs > 2000000000 Or P4Pattern > 2000000000 Or P4State > 2000000000 Then
        RPX_SchurRelease
        P4Solve = 0: P4Matrix = 0: P4Rhs = 0: P4Pattern = 0: P4State = 0
    End If
    P4Solve = P4Solve + 1: P4Generic = -1
    P4Calls = 0: P4Replays = 0: P4Returned = 0: P4Committed = 0
    P4Factors = 0: P4Prepares = 0: P4Hits = 0: P4Reduced = 0
    P4Corrections = 0: P4Checks = 0: P4Attempts = 0
    P4LocalRhs = 0#: P4LocalFactor = 0#
    For i = 0 To 11: phaseTime(i) = 0#: Next i
    RPX_DiagEvent "p4_engine", "engine=20260927_p4_3;trace=" & P4Trace & ";consistent=" & P4Consistent & ";structure_cache=" & P4Cache & ";reuse_u=" & P4ReuseU & ";refinement=FIXED3;generic_start_iters=0;direct_final=0"
End Sub
Public Sub RPX_P4NewMatrix()
    P4Matrix = P4Matrix + 1: P4Generic = -1
End Sub
Public Sub RPX_P4StartDirection(ByVal kind As Long)
    Dim blank As RPX_DirectionInfo
    P4Current = blank: P4Rhs = P4Rhs + 1: P4Calls = P4Calls + 1
    P4Current.solve = P4Solve: P4Current.state = P4State: P4Current.Pattern = P4Pattern
    P4Current.kind = kind: P4Current.matrix = P4Matrix: P4Current.rhs = P4Rhs
End Sub
Public Sub RPX_P4Time(ByVal phase As Long, ByVal started As Double)
    phaseTime(phase) = phaseTime(phase) + RPX_DiagClock() - started
End Sub
Public Sub RPX_P4Direction(ByVal backend As Long, ByVal residual As Double)
    P4Current.returned = backend: P4Current.numericValid = True: P4Current.residual = residual
    If P4Trace = 2 Then RPX_DiagEvent "p4_residual_history", "rhs=" & P4Current.rhs & ";e0=" & P4Current.errors(0) & ";e1=" & P4Current.errors(1) & ";e2=" & P4Current.errors(2) & ";e3=" & P4Current.errors(3)
    If backend = 2 Then P4Returned = P4Returned + 1
    P4Checks = P4Checks + P4Current.checks
    If P4Trace > 0 Then RPX_DiagEvent "p4_direction_return", "solve=" & P4Solve & ";state=" & P4State & ";matrix=" & P4Matrix & ";pattern=" & P4Pattern & ";rhs=" & P4Current.rhs & ";kind=" & P4Current.kind & ";attempted=" & P4Current.attempted & ";returned=" & backend & ";corrections=" & P4Current.corrections & ";schur_corrections=" & P4Current.schurCorrections & ";checks=" & P4Current.checks & ";residual=" & residual
End Sub
Public Sub RPX_P4Commit(ByRef affine As RPX_DirectionInfo, ByRef Corrector As RPX_DirectionInfo)
    If affine.kind <> 1 Or Corrector.kind <> 2 Or affine.returned < 1 Or affine.returned > 2 Then err.Raise 5, "RPX_P4", "DIRECTION_PAIR_KIND"
    If affine.solve <> P4Solve Or Corrector.solve <> P4Solve Or affine.state <> P4State Or Corrector.state <> P4State Then err.Raise 5, "RPX_P4", "DIRECTION_PAIR_STATE"
    If Not affine.numericValid Or Not Corrector.numericValid Or affine.matrix <> P4Matrix Or Corrector.matrix <> P4Matrix Or affine.returned <> Corrector.returned Then err.Raise 5, "RPX_P4", "DIRECTION_PAIR_IDENTITY"
    If affine.returned = 2 Then P4Committed = P4Committed + 2
    If P4Trace > 0 Then RPX_DiagEvent "p4_pair_commit", "matrix=" & P4Matrix & ";affine_rhs=" & affine.rhs & ";corrector_rhs=" & Corrector.rhs & ";backend=" & affine.returned
    P4State = P4State + 1
End Sub
Public Sub RPX_P4Summary()
    Dim i As Long, names As Variant
    RPX_DiagEvent "p4_summary", "calls=" & P4Calls & ";replays=" & P4Replays & ";schur_returned=" & P4Returned & ";schur_committed=" & P4Committed & ";fallback_factors=" & P4Factors & ";prepare=" & P4Prepares & ";hits=" & P4Hits & ";reduced_solves=" & P4Reduced & ";corrections=" & P4Corrections & ";checks=" & P4Checks & ";schur_attempts=" & P4Attempts & ";local_rhs_solves=" & P4LocalRhs & ";local_factor_solves=" & P4LocalFactor
    names = Split("prepare,local_factor,schur_update,numeric_only,reduced_rhs,backsolve,recover,full_residual,cache_verify,correction_inclusive,replay_inclusive,fallback_factor", ",")
    For i = 0 To 11
        RPX_DiagEvent "p4_time", "phase=" & names(i) & ";seconds=" & phaseTime(i) & ";inclusive=" & Abs(CLng(i = 9 Or i = 10))
    Next i
End Sub