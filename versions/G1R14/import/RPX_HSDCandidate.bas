Attribute VB_Name = "RPX_HSDCandidate"
Option Explicit
Public Type RPX_HsdCandidateState
    kind As String
    outputStatus As String
    mode As String
    modelKey As String
    loadPhase As String
    fs As Double
    subdivisions As Long
    solveEpoch As Long
    inputEpoch As Long
    geometryEpoch As Long
    tau As Double
    kappa As Double
    normalization As Double
    originalLoadScale As Double
    objective As Double
    gap As Double
    primalResidual As Double
    dualResidual As Double
    pc As Double
    dc As Double
    certificateError As Double
    auditSummary As String
    numericalGate As Boolean
    physicalGate As Boolean
    valid As Boolean
    x() As Double
    y() As Double
    slack() As Double
    dual() As Double
    field() As Double
End Type
Public RPX_HSDProofX() As Double
Private contextKey As String, geometryEpoch As Long
Private capturePath As String
Public RPX_HSDKind As String, RPX_HSDCandidateId As Long
Public RPX_HSDAuditMode As String
Private savedX() As Double, savedY() As Double, savedSlack() As Double, savedDual() As Double
Private savedTau As Double, savedKappa As Double, savedNormalization As Double
Private captured As Boolean, captureEnabled As Boolean
Public Sub RPX_HSDObservationBegin(ByVal mode As String)
    Dim i As Long
    RPX_HSDAuditMode = mode: RPX_HSDKind = "": RPX_HSDCandidateId = 0
    If capturePath <> RPX_DiagPath Then captured = False: capturePath = RPX_DiagPath
    captureEnabled = (CLng(RPX_Setting("HSD_CAPTURE", 0)) = 1)
    Erase RPX_HSDProofX
    Erase savedX: Erase savedY: Erase savedSlack: Erase savedDual
    RPX_DiagEvent "hsd_start", "mode=" & mode & ";Fs=" & RStrengthFactor & ";xn=" & XN & ";xm=" & XM & ";xb=" & XB & ";xs=" & XS & ";stress_scale=" & RStressScale & ";backend=GENERIC;load_phase=" & RPX_LoadPhase
    If Len(mode) > 0 Then
        If contextKey <> RPX_ExactModel() Then geometryEpoch = geometryEpoch + 1: contextKey = RPX_ExactModel()
        RPX_DiagEvent "hsd_context", "input_epoch=" & RPX_InputEpoch & ";geometry_epoch=" & geometryEpoch & ";solve_epoch=" & P4Solve
        For i = 0 To RM - 1
            With RMaterials(i)
                RPX_DiagEvent "hsd_material", "id=" & i & ";c=" & .cohesion & ";phi=" & .friction & ";psi=" & .dilation & ";gamma=" & .gamma
            End With
        Next i
    End If
End Sub

Private Function FiniteVector(ByRef values() As Double) As Boolean
    Dim i As Long
    For i = LBound(values) To UBound(values)
        If Not RPX_SchurFinite(values(i)) Then Exit Function
    Next i
    FiniteVector = True
End Function
Public Function RPX_HSDBuildCandidate(ByRef candidate As RPX_HsdCandidateState, ByVal kind As String, ByVal divisor As Double, ByVal tau As Double, ByVal kappa As Double, ByVal objective As Double, ByVal gap As Double) As Boolean
    Dim blank As RPX_HsdCandidateState, i As Long
    candidate = blank
    If Not RPX_SchurFinite(divisor) Or Not RPX_SchurFinite(tau) Or Not RPX_SchurFinite(kappa) Then Exit Function
    If divisor <= 0# Or tau <= 0# Or kappa <= 0# Then Exit Function
    If Not RPX_SchurFinite(objective) Or Not RPX_SchurFinite(gap) Then Exit Function
    If UBound(xx) <> XN - 1 Or UBound(xy) < XM - 1 Or UBound(XSlack) <> XS - 1 Or UBound(XDual) <> XS - 1 Then err.Raise 9, "RPX_HSDCandidate", "HSD_CANDIDATE_DIMENSIONS"
    RPX_CapacityCheck "hsd_candidate", 8# * (3# * XN + XM + 1# + 2# * XS)
    candidate.kind = kind: candidate.normalization = divisor
    candidate.tau = tau: candidate.kappa = kappa: candidate.objective = objective: candidate.gap = gap
    candidate.primalResidual = XPrimalResidual: candidate.dualResidual = XDualResidual
    candidate.mode = RPX_HSDAuditMode: candidate.fs = RStrengthFactor: candidate.loadPhase = RPX_LoadPhase
    candidate.subdivisions = RSubdivisions
    candidate.solveEpoch = P4Solve: candidate.inputEpoch = RPX_InputEpoch: candidate.geometryEpoch = geometryEpoch
    If Len(candidate.mode) > 0 Then candidate.modelKey = RPX_ExactModel()
    candidate.x = xx: candidate.y = xy: candidate.slack = XSlack: candidate.dual = XDual
    Select Case kind
        Case "OPTIMAL_POINT"
            candidate.outputStatus = "OPTIMAL"
            For i = 0 To XN - 1: candidate.x(i) = candidate.x(i) / divisor: Next i
            For i = 0 To XM - 1: candidate.y(i) = candidate.y(i) / divisor: Next i
            For i = 0 To XS - 1: candidate.slack(i) = candidate.slack(i) / divisor: candidate.dual(i) = candidate.dual(i) / divisor: Next i
        Case "PRIMAL_INFEASIBILITY_CERT"
            candidate.outputStatus = "INFEASIBLE"
            For i = 0 To XM - 1: candidate.y(i) = candidate.y(i) / divisor: Next i
            For i = 0 To XS - 1: candidate.dual(i) = candidate.dual(i) / divisor: Next i
        Case "LOAD_RECESSION_CERT"
            candidate.outputStatus = "UNBOUNDED"
            For i = 0 To XN - 1: candidate.x(i) = candidate.x(i) / divisor: Next i
        Case Else
            err.Raise 5, "RPX_HSDCandidate", "INVALID_HSD_CANDIDATE_KIND"
    End Select
    If Not FiniteVector(candidate.x) Or Not FiniteVector(candidate.y) Or Not FiniteVector(candidate.slack) Or Not FiniteVector(candidate.dual) Then Exit Function
    candidate.numericalGate = True
    RPX_HSDBuildCandidate = True
End Function
Private Sub AssertCandidateContext(ByRef candidate As RPX_HsdCandidateState)
    If XCancel Then err.Raise 18, "RPX_HSDCandidate", "ANALYSIS_CANCELLED"
    If candidate.solveEpoch <> P4Solve Then err.Raise 5, "RPX_HSDCandidate", "HSD_SOLVE_EPOCH_MISMATCH"
    If Len(candidate.mode) > 0 Then
        RPX_AssertInputs "hsd_candidate"
        If candidate.inputEpoch <> RPX_InputEpoch Or candidate.geometryEpoch <> geometryEpoch Then err.Raise 5, "RPX_HSDCandidate", "HSD_CONTEXT_EPOCH_MISMATCH"
        If candidate.subdivisions <> RSubdivisions Then err.Raise 5, "RPX_HSDCandidate", "HSD_QUADRATURE_MISMATCH"
        If candidate.fs <> RStrengthFactor Or candidate.loadPhase <> RPX_LoadPhase Or candidate.modelKey <> RPX_ExactModel() Then err.Raise 5, "RPX_HSDCandidate", "HSD_MODEL_MISMATCH"
    End If
End Sub
Private Sub LogRawResidual(ByRef values() As Double)
    Dim i As Long, p As Long, row As Double, raw As Double, largest As Double, at As Long, role As String
    If UBound(XEqRowNorm) <> XM Then err.Raise 9, "RPX_HSDCandidate", "HSD_ROW_NORM_DIMENSIONS"
    For i = 0 To XM - 1
        row = -XERhs(i)
        For p = XEPtr(i) To XEPtr(i + 1) - 1: row = row + XEVal(p) * values(XEId(p)): Next p
        raw = Abs(row * XEqRowNorm(i))
        If raw > largest Then largest = raw: at = i
    Next i
    role = "unknown"
    If RPX_HSDAuditMode = "lower" Then
        If Not RLowerEqRoles Is Nothing Then
            If at < RLowerEqRoles.count Then role = CStr(RLowerEqRoles(at + 1))
        End If
    End If
    RPX_DiagEvent "hsd_raw_equality", "id=" & RPX_HSDCandidateId & ";max=" & largest & ";row=" & at & ";row_norm=" & XEqRowNorm(at) & ";role=" & role
End Sub
Private Function PhysicalRejection(ByVal number As Long, ByVal message As String) As Boolean
    If number <> 5 Then Exit Function
    PhysicalRejection = (Split(message & " ", " ")(0) = "LOWER_PHYSICAL_AUDIT_FAILED" Or Split(message & " ", " ")(0) = "UPPER_PHYSICAL_AUDIT_FAILED")
End Function
Public Function RPX_HSDValidateCandidate(ByRef candidate As RPX_HsdCandidateState) As Boolean
    Dim oldX() As Double, oldObjective As Double, oldValue As Double, oldAudit As Double
    Dim oldDissipation As Double, oldReference As Double, oldFixed As Double, oldRejected As Boolean
    Dim auditState As RPX_AuditSnapshot, installed As Boolean, accepted As Boolean
    Dim oldRaw(1 To 4) As Double, oldGate As Boolean, oldReason As String
    Dim number As Long, message As String, source As String, line As Long, i As Long, summary As String
    candidate.valid = False: candidate.physicalGate = False
    If Not candidate.numericalGate Then Exit Function
    AssertCandidateContext candidate
    If AFActive And Len(candidate.mode) = 0 Then
        ' Condensed fixed-load runs accept only a restored primal witness.
        ' A reduced dual/ray is never promoted to a full-problem certificate.
        If candidate.kind <> "OPTIMAL_POINT" Then Exit Function
        If Not RPX_AffineCandidate(candidate) Then Exit Function
        candidate.physicalGate = True: candidate.valid = True: RPX_HSDValidateCandidate = True: Exit Function
    End If
    If Len(candidate.mode) = 0 Or candidate.kind = "PRIMAL_INFEASIBILITY_CERT" Then
        candidate.physicalGate = True: candidate.valid = True: RPX_HSDValidateCandidate = True: Exit Function
    End If
    If candidate.kind = "LOAD_RECESSION_CERT" And Not (candidate.mode = "lower" And RFsActive) Then
        candidate.physicalGate = True: candidate.valid = True: RPX_HSDValidateCandidate = True: Exit Function
    End If
    candidate.field = candidate.x
    If candidate.kind = "LOAD_RECESSION_CERT" Then
        If RLoadVariable < 0 Or RLoadVariable >= XN Then err.Raise 9, "RPX_HSDCandidate", "HSD_LOAD_ID"
        candidate.originalLoadScale = candidate.x(RLoadVariable)
        If Not RPX_SchurFinite(candidate.originalLoadScale) Then Exit Function
        If candidate.originalLoadScale <= 0# Then Exit Function
        For i = 0 To XN - 1: candidate.field(i) = candidate.field(i) / candidate.originalLoadScale: Next i
        If Not FiniteVector(candidate.field) Then Exit Function
    End If
    oldX = xx: oldObjective = XObjective
    oldValue = RValue: oldAudit = RAudit: oldDissipation = RDissipation
    oldReference = RReferenceWork: oldFixed = RFixedWork: oldRejected = RPX_AuditRejected
    oldGate = RPhysicalGateAccepted: oldReason = RPhysicalGateReason
    For i = 1 To 4: oldRaw(i) = RRawLowerAudit(i): Next i
    RPX_DiagAuditSave auditState
    On Error GoTo Failed
    If candidate.mode = "upper" Then
        If RPX_NormalizeZeroCostUpper(candidate) Then candidate.field = candidate.x
    End If
    installed = True: xx = candidate.field: XObjective = candidate.objective
    RPX_TestHook "hsd_candidate_audit"
    LogRawResidual candidate.field
    If candidate.mode = "upper" Then Call RPX_AuditUpper(True) Else Call RPX_AuditLower(True)
    accepted = RPhysicalGateAccepted
    If Not accepted Then message = RPhysicalGateReason
    GoTo Restore
Failed:
    number = err.number: message = err.description: source = err.source: line = Erl
    Resume Restore
Restore:
    On Error GoTo captureFailed
    summary = RPX_DiagAuditSummary()
    candidate.auditSummary = summary
    RPX_DiagEvent "hsd_candidate_audit", "id=" & RPX_HSDCandidateId & ";kind=" & candidate.kind & ";accepted=" & accepted & ";field_normalization=" & IIf(candidate.kind = "LOAD_RECESSION_CERT", "lambda_one", "tau") & ";audit=" & summary
    If Not accepted And number = 0 Then
        ' Capture errors must not bypass restoration of the live iterate.
        On Error GoTo captureFailed
        RPX_HSDCaptureFailure message
    End If
restoreNow:
    On Error GoTo 0
    If installed Then
        xx = oldX: XObjective = oldObjective
        RValue = oldValue: RAudit = oldAudit: RDissipation = oldDissipation
        RReferenceWork = oldReference: RFixedWork = oldFixed: RPX_AuditRejected = oldRejected
        RPhysicalGateAccepted = oldGate: RPhysicalGateReason = oldReason
        For i = 1 To 4: RRawLowerAudit(i) = oldRaw(i): Next i
        RPX_DiagAuditRestore auditState
    End If
    If Not accepted Then
        If XCancel And number <> 18 Then err.Raise 18, "RPX_HSDCandidate", "ANALYSIS_CANCELLED"
        If number <> 0 Then
            RPX_ErrorEvidence number, source, message, line, "hsd_candidate_audit"
            err.Raise number, source, message
        End If
        AssertCandidateContext candidate
        RPX_DiagEvent "hsd_candidate_rejected", "id=" & RPX_HSDCandidateId & ";iteration=" & XIteration & ";restored=True;reason=" & message
        Exit Function
    End If
    AssertCandidateContext candidate
    candidate.physicalGate = True: candidate.valid = True: RPX_HSDValidateCandidate = True
    Exit Function
captureFailed:
    accepted = False
    number = err.number: message = err.description: source = err.source: line = Erl
    Resume restoreNow
End Function
Public Sub RPX_HSDCommitCandidate(ByRef candidate As RPX_HsdCandidateState)
    AssertCandidateContext candidate
    If Not candidate.valid Or Not candidate.numericalGate Or Not candidate.physicalGate Then err.Raise 5, "RPX_HSDCandidate", "HSD_UNVALIDATED_COMMIT"
    xx = candidate.x: xy = candidate.y: XSlack = candidate.slack: XDual = candidate.dual
    XObjective = candidate.objective: XGap = candidate.gap: XStatus = candidate.outputStatus
    If candidate.kind = "LOAD_RECESSION_CERT" Then RPX_HSDProofX = candidate.x
    RPX_AuditRejected = False
    RPX_DiagEvent "hsd_commit", "id=" & RPX_HSDCandidateId & ";kind=" & candidate.kind & ";status=" & XStatus & ";solve_epoch=" & P4Solve & ";physical_gate=" & candidate.physicalGate
End Sub
Public Sub RPX_HSDObserve(ByVal kind As String, ByVal tau As Double, ByVal kappa As Double, ByVal normalization As Double, ByVal pc As Double, ByVal dc As Double, ByVal certificateError As Double)
    RPX_HSDKind = kind: RPX_HSDCandidateId = RPX_HSDCandidateId + 1
    RPX_DiagEvent "hsd_candidate", "id=" & RPX_HSDCandidateId & ";kind=" & kind & ";iteration=" & XIteration & ";tau=" & tau & ";kappa=" & kappa & ";normalization=" & normalization & ";pc=" & pc & ";dc=" & dc & ";certificate_error=" & certificateError & ";primal=" & XPrimalResidual & ";dual=" & XDualResidual & ";status=" & XStatus
    If captureEnabled And Not captured Then
        RPX_CapacityCheck "hsd_observation", 8# * (XN + XM + 1# + 2# * XS)
        savedX = xx: savedY = xy: savedSlack = XSlack: savedDual = XDual
        savedTau = tau: savedKappa = kappa: savedNormalization = normalization
    End If
End Sub
Private Sub WriteVector(ByVal stream As Object, ByVal name As String, ByRef values() As Double)
    Dim i As Long
    For i = LBound(values) To UBound(values)
        stream.WriteLine name & vbTab & i & vbTab & RPX_ExactDouble(values(i))
    Next i
End Sub
Private Sub WriteNormalized(ByVal stream As Object, ByVal name As String, ByRef values() As Double, ByVal divisor As Double, ByVal count As Long)
    Dim i As Long, value As Double
    For i = LBound(values) To UBound(values)
        value = values(i)
        If i < count Then value = value / divisor
        stream.WriteLine name & vbTab & i & vbTab & RPX_ExactDouble(value)
    Next i
End Sub
Public Sub RPX_HSDCaptureFailure(ByVal reason As String)
    Dim fs As Object, stream As Object, path As String, i As Long, j As Long, k As Long, rec As Variant
    If Not captureEnabled Or captured Or Len(RPX_HSDKind) = 0 Then Exit Sub
    captured = True
    If Len(RPX_DiagPath) = 0 Then Exit Sub
    path = RPX_DiagPath & ".hsd_" & RPX_HSDCandidateId & ".tsv"
    Set fs = CreateObject("Scripting.FileSystemObject"): Set stream = fs.CreateTextFile(path, False, True)
    stream.WriteLine "format=IEEE754_LE_HEX;kind=" & RPX_HSDKind & ";candidate=" & RPX_HSDCandidateId & ";reason=" & reason
    stream.WriteLine "tau" & vbTab & RPX_ExactDouble(savedTau)
    stream.WriteLine "kappa" & vbTab & RPX_ExactDouble(savedKappa)
    stream.WriteLine "normalization" & vbTab & RPX_ExactDouble(savedNormalization)
    stream.WriteLine "Fs" & vbTab & RPX_ExactDouble(RStrengthFactor)
    stream.WriteLine "load_id" & vbTab & RLoadVariable
    If Len(RPX_HSDAuditMode) > 0 Then stream.WriteLine "exact_model" & vbTab & RPX_ExactModel()
    If Len(RPX_HSDAuditMode) > 0 Then RPX_DumpInputSnapshot stream
    If Len(RPX_HSDAuditMode) > 0 Then
        For i = 0 To re - 1
            stream.WriteLine "area" & vbTab & i & vbTab & RPX_ExactDouble(RArea(i))
            For j = 0 To 2
                For k = 0 To 1: stream.WriteLine "gradient" & vbTab & i & vbTab & j & vbTab & k & vbTab & RPX_ExactDouble(RGrad(i, j, k)): Next k
            Next j
        Next i
        For i = 0 To RNEdge - 1
            With REdges(i)
                stream.WriteLine "edge_geometry" & vbTab & i & vbTab & RPX_ExactDouble(.length) & vbTab & RPX_ExactDouble(.normal(0)) & vbTab & RPX_ExactDouble(.normal(1))
                For j = 0 To 1
                    For k = 0 To 2: stream.WriteLine "edge_local" & vbTab & i & vbTab & j & vbTab & k & vbTab & .loc(j, k): Next k
                Next j
            End With
        Next i
    End If
    WriteVector stream, "live_x", savedX: WriteVector stream, "live_y", savedY
    WriteVector stream, "live_slack", savedSlack: WriteVector stream, "live_dual", savedDual
    WriteVector stream, "candidate_x", xx
    If RPX_HSDKind = "OPTIMAL_POINT" Or RPX_HSDKind = "PRIMAL_INFEASIBILITY_CERT" Then
        WriteNormalized stream, "candidate_y", savedY, savedNormalization, XM
        WriteNormalized stream, "candidate_dual", savedDual, savedNormalization, XS
    Else
        WriteVector stream, "candidate_y", savedY: WriteVector stream, "candidate_dual", savedDual
    End If
    If RPX_HSDKind = "OPTIMAL_POINT" Then
        WriteNormalized stream, "candidate_slack", savedSlack, savedNormalization, XS
    Else
        WriteVector stream, "candidate_slack", savedSlack
    End If
    If RPX_HSDKind = "LOAD_RECESSION_CERT" Then WriteNormalized stream, "certificate_x", savedX, savedNormalization, XN
    WriteVector stream, "q", XQ: WriteVector stream, "eq_rhs", XERhs
    WriteVector stream, "eq_row_norm", XEqRowNorm
    For i = 0 To XM - 1
        For j = XEPtr(i) To XEPtr(i + 1) - 1
            stream.WriteLine "eq" & vbTab & i & vbTab & XEId(j) & vbTab & RPX_ExactDouble(XEVal(j))
        Next j
    Next i
    For i = 0 To XB - 1
        With XC(i)
            stream.WriteLine "cone" & vbTab & i & vbTab & .dimn & vbTab & .count & vbTab & .first & vbTab & .owner & vbTab & .parameter1 & vbTab & .parameter2
            For j = 0 To .dimn - 1
                stream.WriteLine "offset" & vbTab & i & vbTab & j & vbTab & RPX_ExactDouble(.offset(j))
                For k = 0 To .count - 1
                    stream.WriteLine "coef" & vbTab & i & vbTab & j & vbTab & .ids(k) & vbTab & RPX_ExactDouble(.coef(j, k))
                Next k
            Next j
        End With
    Next i
    If RPX_HSDAuditMode = "lower" Then
        For i = 0 To re - 1: stream.WriteLine "stress_start" & vbTab & i & vbTab & RStressStart(i): Next i
        For Each rec In RReactions
            stream.WriteLine "reaction" & vbTab & rec(0) & vbTab & rec(1) & vbTab & rec(2) & vbTab & RPX_ExactDouble(CDbl(rec(3))) & vbTab & RPX_ExactDouble(CDbl(rec(4)))
        Next rec
        For i = 1 To RLowerEqRoles.count: stream.WriteLine "eq_role" & vbTab & i - 1 & vbTab & RLowerEqRoles(i): Next i
    End If
    stream.Close
    RPX_DiagEvent "hsd_capture", "path=" & path
End Sub
