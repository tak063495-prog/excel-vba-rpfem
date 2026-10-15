Attribute VB_Name = "TerminalReuseProbe"
Option Explicit
Private Sub Check(ByVal value As Boolean, ByVal message As String)
    If Not value Then err.Raise 5, "TerminalReuseProbe", message
End Sub
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
    err.Raise 5, "TerminalReuseProbe", "NO_SETTING_SLOT"
End Sub
Private Function ExactField(ByRef field() As Double) As String
    Dim result() As String, i As Long
    ReDim result(LBound(field) To UBound(field))
    For i = LBound(field) To UBound(field): result(i) = RPX_ExactDouble(field(i)): Next i
    ExactField = Join(result, "")
End Function
Private Sub Rejected(ByVal label As String, ByVal calls As Long, ByVal lower As Double, ByVal upper As Double)
    Check Not RPX_TryTerminalPairReuse("Fs"), "INVALID_REUSE_" & label
    Check RFsCalls = calls, "REJECTION_STARTED_SOLVER_" & label
    Check RLowerFs = lower And RUpperFs = upper, "REJECTION_CHANGED_BOUNDS_" & label
End Sub
' Call after a completed real q8 adaptive run. This function starts no solver.
' Metadata mutations and corrupt actual fields must all reject the fast path.
Public Function TerminalReuseSavedPairProbe() As String
    Dim lower As RPX_MeshState, upper As RPX_MeshState
    Dim lowerCert As RPX_BoundCertificate, upperCert As RPX_BoundCertificate
    Dim lowerBits As String, upperBits As String, calls As Long, i As Long
    Dim savedTol As Double, savedRaw As Boolean, original As String, savedFactor As Double, savedX As Double
    Check RPX_AdaptPairComplete(), "REAL_ADAPT_PAIR_REQUIRED"
    lower = RBestLower: upper = RBestUpper: lowerCert = RPX_CertLower: upperCert = RPX_CertUpper
    lowerBits = ExactField(lower.field): upperBits = ExactField(upper.field)
    calls = RFsCalls: savedTol = RFsTolerance: savedRaw = RRawAuditEnabled: original = RPX_ExactModel()
    savedFactor = RStrengthFactor: savedX = RXY(0, 0)
    On Error GoTo Failed
    Check RPX_TryTerminalPairReuse("Fs"), "REAL_Q8_PAIR_NOT_REUSED"
    Check RFsCalls = calls, "REUSE_STARTED_SOLVER"
    Check ExactField(RBestLower.field) = lowerBits And ExactField(RBestUpper.field) = upperBits, "REUSE_CHANGED_FIELDS"
    Check RLowerFs = lower.value And RUpperFs = upper.value, "REUSE_CHANGED_BOUNDS"
    Check RPX_ExactModel() = original, "REUSE_CHANGED_LIVE_MODEL"
    Check RStrengthFactor = savedFactor, "REUSE_CHANGED_LIVE_FACTOR"
    RXY(0, 0) = savedX + 0.125: Rejected "live_mesh", calls, lower.value, upper.value: RXY(0, 0) = savedX
    RBestLower.q = 4: Rejected "q", calls, lower.value, upper.value: RBestLower = lower
    RBestLower.certificateRunId = lower.certificateRunId - 1: Rejected "state_run", calls, lower.value, upper.value: RBestLower = lower
    RPX_CertLower.runId = lowerCert.runId - 1: Rejected "certificate_run", calls, lower.value, upper.value: RPX_CertLower = lowerCert
    RBestLower.inputEpoch = lower.inputEpoch - 1: Rejected "epoch", calls, lower.value, upper.value: RBestLower = lower
    RBestLower.policy = lower.policy & "_foreign": Rejected "policy", calls, lower.value, upper.value: RBestLower = lower
    RBestLower.lowerRootValid = False: Rejected "incomplete_root", calls, lower.value, upper.value: RBestLower = lower
    RBestLower.lowerRootA = lower.lowerRootA + 0.00000001: Rejected "wrong_endpoint", calls, lower.value, upper.value: RBestLower = lower
    RBestLower.lowerRootB = lower.lowerRootA + 1#: RPX_CertLower.rootB = RBestLower.lowerRootB
    Rejected "wide_matching_root", calls, lower.value, upper.value: RBestLower = lower: RPX_CertLower = lowerCert
    RBestLower.factor = lower.factor + 0.00000001: Rejected "factor", calls, lower.value, upper.value: RBestLower = lower
    RBestLower.originalModel = lower.originalModel & "_foreign": Rejected "own_mesh_model", calls, lower.value, upper.value: RBestLower = lower
    RBestLower.lowerLoadVariable = UBound(lower.field) + 1: Rejected "lower_layout", calls, lower.value, upper.value: RBestLower = lower
    For i = LBound(RBestLower.field) To UBound(RBestLower.field): RBestLower.field(i) = 0#: Next i
    Rejected "actual_lower_field", calls, lower.value, upper.value: RBestLower = lower
    For i = LBound(RBestUpper.field) To UBound(RBestUpper.field): RBestUpper.field(i) = 0#: Next i
    Rejected "actual_upper_field", calls, lower.value, upper.value: RBestUpper = upper
    RRawAuditEnabled = False: Rejected "raw_audit_off", calls, lower.value, upper.value: RRawAuditEnabled = savedRaw
    Check Not RPX_TryTerminalPairReuse("極限支持力"), "FOREIGN_MODE_REUSED"
    Check RPX_TryTerminalPairReuse("Fs"), "VALID_PAIR_NOT_REUSABLE_AFTER_REJECTIONS"
    Check RFsCalls = calls And RPX_ExactModel() = original, "REJECTIONS_CHANGED_CONTEXT"
    Check RStrengthFactor = savedFactor, "REJECTIONS_CHANGED_LIVE_FACTOR"
    Check ExactField(RBestLower.field) = lowerBits And ExactField(RBestUpper.field) = upperBits, "REJECTIONS_CHANGED_SOURCE_FIELDS"
    TerminalReuseSavedPairProbe = "PASS;actual_q8_fields=2;invalid_context_or_field_cases=16;new_endpoint_calls=0;field_bytes_unchanged=1;bounds_unchanged=1"
    Exit Function
Failed:
    Dim number As Long, source As String, message As String
    number = err.number: source = err.source: message = err.description
    RBestLower = lower: RBestUpper = upper: RPX_CertLower = lowerCert: RPX_CertUpper = upperCert
    RFsTolerance = savedTol: RRawAuditEnabled = savedRaw
    RStrengthFactor = savedFactor: RXY(0, 0) = savedX
    On Error GoTo 0
    err.Raise number, source, message
End Function
' Optional small actual fixture for the serial owned-Excel test run.
' GAP=2 is test-only to exercise the terminal path on a coarse mesh. It makes
' no production accuracy claim; original field/root tolerances stay unchanged.
Public Function TerminalReuseSmallFixtureRun(ByVal trace As String) As String
    RPX_Setup
    Setting "MODE", "Fs": Setting "TARGET", 16: Setting "ADAPT", 1: Setting "CYCLES", 0
    Setting "QUADRATURE", 8: Setting "ADAPT_CERT_QUADRATURE", 8: Setting "ADAPT_PRECISION_FROM_START", 1
    Setting "RAW_LOWER_AUDIT", 1: Setting "GAP", 2#: Setting "C0_AUTO", 0
    Setting "ADAPT_MESH_POLICY", "ROBUST_REFINE": Setting "ADAPT_ACCURACY_CAP", 16: Setting "ADAPT_ACCURACY_ROUNDS", 0
    Setting "ADAPT_LOWER_FAN_ROUNDS", 0: Setting "ADAPT_C0_WEDGE_DEPTH", 0: Setting "UPPER_PILOT", 0
    Setting "HSD_SOLVE_SECONDS", 30: Setting "TEST_INJECTION", 0
    RPX_Sheet("材料データ").Range("C5").Value2 = 0#: RPX_Sheet("材料データ").Range("E5").Value2 = 0#
    RPX_Sheet("浸透入力").Range("D4").Value2 = 0
    XLogPath = trace
    RPX_Run
    Check Len(RPX_LastError) = 0, "FIXTURE_RUN_FAILED_" & RPX_LastError
    TerminalReuseSmallFixtureRun = TerminalReuseSavedPairProbe()
End Function
