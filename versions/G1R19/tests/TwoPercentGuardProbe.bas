Attribute VB_Name = "TwoPercentGuardProbe"
Option Explicit
Private Sub Setting(ByVal key As String, ByVal value As Variant)
    Dim ws As Worksheet, r As Long
    Set ws = RPX_Sheet("設定")
    For r = 4 To 1000
        If CStr(ws.Cells(r, 4).Value2) = key Then ws.Cells(r, 2).Value2 = value: Exit Sub
    Next r
    err.Raise 5, "TwoPercentGuardProbe", "EXISTING_SETTING_REQUIRED_" & key
End Sub
Private Sub Check(ByVal valid As Boolean, ByVal message As String)
    If Not valid Then err.Raise 5, "TwoPercentGuardProbe", message
End Sub
' Only tests RPX_TargetMet's symmetric-gap setting boundary. Temporary metadata
' is synthetic and never published as a field/root proof or used by a solver.
Public Function TwoPercentGuardUnit() As String
    Dim lower As RPX_BoundCertificate, upper As RPX_BoundCertificate
    Dim gap As Variant, q As Variant, savedRun As Long, savedInternal As Boolean
    Dim incomplete As Boolean, rejected As Boolean, cancelled As Boolean
    Dim synthetic As RPX_BoundCertificate, ratio As Double, result As String
    If RPX_Busy Then err.Raise 5, "TwoPercentGuardProbe", "TEST_WHILE_BUSY"
    lower = RPX_CertLower: upper = RPX_CertUpper: savedRun = RPX_CertRunId
    gap = RPX_Setting("GAP", 0.01): q = RPX_Setting("ADAPT_CERT_QUADRATURE", 0)
    savedInternal = RPX_InternalWrite: incomplete = RFsSearchIncomplete
    rejected = RPX_AuditRejected: cancelled = XCancel
    On Error GoTo Failed
    RPX_InternalWrite = True
    Setting "GAP", 0.02: Setting "ADAPT_CERT_QUADRATURE", 8
    If RPX_CertRunId <= 0 Then RPX_CertRunId = 1
    synthetic.valid = True: synthetic.runId = RPX_CertRunId
    synthetic.inputEpoch = RPX_InputEpoch: synthetic.policy = RPX_AnalysisPolicy()
    synthetic.q = 8: synthetic.rootBracketValid = True
    RPX_CertLower = synthetic: RPX_CertUpper = synthetic
    RFsSearchIncomplete = False: RPX_AuditRejected = False: XCancel = False
    ratio = (2# + 0.0199) / (2# - 0.0199)
    Check RPX_TargetMet("Fs", 1#, ratio, True), "GAP_1P99_REJECTED"
    ratio = (2# + 0.0201) / (2# - 0.0201)
    Check Not RPX_TargetMet("Fs", 1#, ratio, True), "GAP_2P01_ACCEPTED"
    Check Not RPX_TargetMet("Fs", 1#, 1.001, False), "INCOMPLETE_PAIR_ACCEPTED"
    result = "PASS;setting=0.02;symmetric_1.99pct=accept;symmetric_2.01pct=reject;synthetic_guard_only=1;field_or_root_created=0"
Restore:
    On Error GoTo RestoreFailed
    RPX_CertLower = lower: RPX_CertUpper = upper: RPX_CertRunId = savedRun
    RFsSearchIncomplete = incomplete: RPX_AuditRejected = rejected: XCancel = cancelled
    RPX_InternalWrite = True
    Setting "GAP", gap: Setting "ADAPT_CERT_QUADRATURE", q
    RPX_InternalWrite = savedInternal
    TwoPercentGuardUnit = result
    Exit Function
Failed:
    result = "FAIL;number=" & err.number & ";source=" & err.source & ";message=" & err.description
    Resume Restore
RestoreFailed:
    RPX_InternalWrite = savedInternal
    TwoPercentGuardUnit = "FAIL;restore=" & err.description
End Function
