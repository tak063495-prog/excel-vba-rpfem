Attribute VB_Name = "WaterCommonRollbackProbe"
Option Explicit
' Test-copy only. Prepare a fresh nonconforming raw D08 soil mesh with
' coupling OFF; then turn coupling ON before calling this probe. Prescribed
' head3 result must already be valid; TEST_INJECTION=1 and RPX_Busy=False.
' Rebuild the raw seed before EVERY call, especially after restoration failure.
' No solver call and no physical certificate is created by this probe.
Public Function WaterCommonRollbackTest(ByVal target As String) As String
    Dim before As String, originalNumber As Long, originalText As String, stage As String, restoreNumber As Long
    Dim accepted As Boolean, n As Long, message As String, expected As Long, hits As Long, cancelled As Boolean
    Dim level As Double, waterGamma As Double, crossing As Boolean, e As Long, k As Long, lo As Double, hi As Double
    On Error GoTo Failed
    If RPX_Busy Or RPX_LoadPhase <> "raw" Or rn < 3 Or re < 1 Then err.Raise 5, , "PROBE_REQUIRES_RAW_IDLE_MESH"
    XCancel = False: Call RPX_InputBegin
    If Not RPX_WaterCommonLevel(level, waterGamma) Then err.Raise 5, , "PROBE_REQUIRES_ELIGIBLE_COMMON_SOIL"
    For e = 0 To re - 1
        lo = RXY(RTri(e, 0), 1): hi = lo
        For k = 1 To 2
            If RXY(RTri(e, k), 1) < lo Then lo = RXY(RTri(e, k), 1)
            If RXY(RTri(e, k), 1) > hi Then hi = RXY(RTri(e, k), 1)
        Next k
        If lo < level And hi > level Then crossing = True
    Next e
    If Not crossing Then err.Raise 5, , "PROBE_REQUIRES_CROSSING_RAW_SEED"
    before = RPX_ExactModel(): originalNumber = 5: originalText = "WATER_COMMON_PROBE_ORIGINAL": expected = 5
    Select Case target
        Case "restored": expected = 0
        Case "arrays": stage = "water_common_rollback_arrays": restoreNumber = 7
        Case "topology": stage = "water_common_rollback_topology": restoreNumber = 5
        Case "faces": stage = "water_common_rollback_faces": restoreNumber = 5
        Case "original_cancel": originalNumber = 18: expected = 18
        Case "restore_cancel": stage = "water_common_rollback_topology": restoreNumber = 18: expected = 18
        Case "original_input": originalText = "INPUT_MISMATCH WATER_COMMON_PROBE_ORIGINAL"
        Case "original_cancel_restore_failure"
            originalNumber = 18: stage = "water_common_rollback_topology": restoreNumber = 5: expected = 18
        Case Else: err.Raise 5, , "UNKNOWN_WATER_ROLLBACK_PROBE"
    End Select
    RPX_TestArm "water_common_after_commit", originalNumber, originalText
    If Len(stage) > 0 Then RPX_TestQueue stage, restoreNumber, "WATER_COMMON_PROBE_RESTORE"
    ' Test-only sentinels prove that failure actively invalidates both flags.
    ' They are not a field/root certificate and are cleared on every exit.
    If Len(stage) > 0 Then RPX_ResultCurrent = True: RPX_MeshCurrent = True
    On Error Resume Next
    accepted = RPX_WaterCommonMesh()
    n = err.number: message = err.description: err.Clear
    On Error GoTo Failed
    hits = RPX_TestHits: cancelled = XCancel
    If accepted Or n <> expected Then err.Raise 5, , "ROLLBACK_RESULT_OR_ERROR_WRONG " & target & " " & n & " " & message
    If Len(stage) > 0 Then
        If hits <> 2 Or Not cancelled Or RPX_ResultCurrent Or RPX_MeshCurrent Then err.Raise 5, , "ROLLBACK_FAILURE_NOT_INVALIDATED"
        If InStr(message, "WATER_COMMON_ROLLBACK_FAILED") = 0 Or InStr(message, "original_number=" & originalNumber) = 0 Or InStr(message, "restore_number=" & restoreNumber) = 0 Then err.Raise 5, , "ROLLBACK_ERRORS_NOT_SEPARATE"
    Else
        If hits <> 1 Or RPX_ExactModel() <> before Then err.Raise 5, , "VERIFIED_RESTORE_MODEL_CHANGED"
        If expected <> 0 Then If InStr(message, originalText) = 0 Then err.Raise 5, , "ORIGINAL_INTERRUPT_NOT_PROPAGATED"
    End If
    Call RPX_TestDisarm: Call RPX_InputEnd: XCancel = False
    RPX_ResultCurrent = False: RPX_MeshCurrent = False
    WaterCommonRollbackTest = "PASS " & target & ";error=" & n & ";hits=" & hits & ";cancelled=" & cancelled
    Exit Function
Failed:
    message = err.description
    On Error Resume Next
    Call RPX_TestDisarm: Call RPX_InputEnd: XCancel = False
    RPX_ResultCurrent = False: RPX_MeshCurrent = False
    WaterCommonRollbackTest = "FAIL " & target & " " & message
End Function
