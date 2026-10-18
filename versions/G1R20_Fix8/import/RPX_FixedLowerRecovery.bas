Attribute VB_Name = "RPX_FixedLowerRecovery"
Option Explicit
' Independent same-Fs primal recovery. Fixed-problem duals/status never prove a root.
Private Type RecoveryProgram
    n As Long
    m As Long
    s As Long
    b As Long
    qScale As Double
    cones() As RPX_ConeData
    cost() As Double
    rowNorm() As Double
    ePtr() As Long
    eId() As Long
    eVal() As Double
    rhs() As Double
    kPtr() As Long
    kId() As Long
    kBase() As Double
    kVal() As Double
    kDiag() As Long
    perm() As Long
    inv() As Long
    kScale() As Double
    original As RPX_OriginalCompiledData
    stressStart() As Long
    loadId As Long
    reactions As Collection
    eqRoles As Collection
End Type
Public FixedLowerRecoveryReason As String
Public FixedLowerRecoveryRoute As String
Public FixedLowerRecoveryValue As Double, FixedLowerRecoveryAudit As Double
Public FixedLowerRecoveryEquality As Double, FixedLowerRecoveryCone As Double
Public FixedLowerRecoveryVariables As Long, FixedLowerRecoveryEqualities As Long
Public FixedLowerRecoveryRaw(1 To 4) As Double

Private Sub CaptureProgram(ByRef state As RecoveryProgram)
    state.n = XN: state.m = XM: state.s = XS: state.b = XB: state.qScale = XQScale
    state.cones = XC: state.cost = XQ: state.rowNorm = XEqRowNorm
    state.ePtr = XEPtr: state.eId = XEId: state.eVal = XEVal: state.rhs = XERhs
    state.kPtr = XKPtr: state.kId = XKId: state.kBase = XKBase: state.kVal = XKVal: state.kDiag = XKDiag
    state.perm = XPerm: state.inv = XInv: state.kScale = XKScale: state.original = XOriginalCompiled
    state.stressStart = RStressStart: state.loadId = RLoadVariable
    Set state.reactions = RReactions: Set state.eqRoles = RLowerEqRoles
End Sub
Private Sub RestoreProgram(ByRef state As RecoveryProgram)
    Call RPX_LDLRelease: Call RPX_HLURelease: Call RPX_SchurRelease: Call RPX_ResetPreparedState
    XN = state.n: XM = state.m: XS = state.s: XB = state.b: XQScale = state.qScale
    XC = state.cones: XQ = state.cost: XEqRowNorm = state.rowNorm
    XEPtr = state.ePtr: XEId = state.eId: XEVal = state.eVal: XERhs = state.rhs
    XKPtr = state.kPtr: XKId = state.kId: XKBase = state.kBase: XKVal = state.kVal: XKDiag = state.kDiag
    XPerm = state.perm: XInv = state.inv: XKScale = state.kScale: XOriginalCompiled = state.original
    RStressStart = state.stressStart: RLoadVariable = state.loadId
    Set RReactions = state.reactions: Set RLowerEqRoles = state.eqRoles
End Sub
Private Function NormalizeOriginalIterate(ByRef original() As Double, ByVal loadId As Long, ByRef restored() As Double) As Boolean
    Dim factor As Double, i As Long
    On Error GoTo Rejected
    If LBound(original) <> 0 Or UBound(original) <> XN - 1 Then Exit Function
    If Not RPX_SchurFinite(original(loadId)) Or original(loadId) <= 0# Then Exit Function
    factor = 1.000003 / original(loadId)
    If Not RPX_SchurFinite(factor) Or factor <= 0# Then Exit Function
    restored = original
    For i = 0 To UBound(restored)
        If Not RPX_SchurFinite(original(i)) Then Exit Function
        restored(i) = original(i) * factor
        If Not RPX_SchurFinite(restored(i)) Then Exit Function
    Next i
    NormalizeOriginalIterate = True: Exit Function
Rejected:
    If err.number <> 6 And err.number <> 9 Then err.Raise err.number, err.source, err.description
    err.Clear: Erase restored
End Function
Public Function RPX_TryFixedLowerRecovery(ByVal factor As Double, ByRef recovered() As Double, Optional ByVal seconds As Double = 120#, Optional ByVal reuseOriginalIterate As Boolean = False) As Boolean
    Const REQUIRED_LOAD As Double = 1.000003
    Dim program As RecoveryProgram, oldX() As Double, oldY() As Double, oldSlack() As Double, oldDual() As Double, oldProof() As Double
    Dim restored() As Double, oldModel As String, model As String, oldFs As Double, oldFactor As Double, oldOrder As String, oldKind As String, oldStatus As String
    Dim oldValue As Double, oldObjective As Double, oldAudit As Double, oldFixed As Double, oldReference As Double, oldDissipation As Double
    Dim oldEqError As Double, oldConeError As Double, oldOriginalObjective As Double, oldRaw(1 To 4) As Double
    Dim oldRejected As Boolean, oldGate As Boolean, oldReason As String, oldQ As Long, oldEpoch As Long
    Dim oldPhysical As Boolean, oldRawScale As Double, oldHsdKind As String, oldHsdId As Long
    Dim oldResidual As Double, oldPrimal As Double, oldDualResidual As Double, oldGap As Double, oldIteration As Long
    Dim proofReady As Boolean, accepted As Boolean, candidate As Boolean, changed As Boolean, i As Long, raw() As Double
    Dim number As Long, source As String, message As String, cleanupNumber As Long, cleanupSource As String, cleanupMessage As String, auditSnapshot As RPX_AuditSnapshot, observation As RPX_HsdObservationSnapshot
    Erase recovered: FixedLowerRecoveryReason = "SCOPE_REJECTED"
    FixedLowerRecoveryRoute = "NONE"
    FixedLowerRecoveryValue = 0#: FixedLowerRecoveryAudit = 0#: FixedLowerRecoveryEquality = 0#: FixedLowerRecoveryCone = 0#
    FixedLowerRecoveryVariables = 0: FixedLowerRecoveryEqualities = 0: Erase FixedLowerRecoveryRaw
    If XCancel Or AFActive Or Not RFsActive Or RSubdivisions <> 8 Or RR <> 0 Or Not RRawAuditEnabled Then Exit Function
    If RPX_LoadPhase <> "reference_total" Or RPX_OrderMode <> "lower" Or RPX_AuditRejected Then Exit Function
    If XStatus <> "NUMERICAL_UNKNOWN" And XStatus <> "NUMERICAL_FAILURE" Then Exit Function
    If Not RPX_SchurFinite(factor) Or Not RPX_SchurFinite(seconds) Then Exit Function
    If factor <= 0# Or seconds <= 0# Or seconds > 600# Or factor <> RStrengthFactor Then Exit Function
    If Not XOriginalCompiled.ready Or XOriginalCompiled.affine Or XOriginalCompiled.modeName <> "lower" Then Exit Function
    If XOriginalCompiled.n <> XN Or XOriginalCompiled.m <> XM Or XOriginalCompiled.s <> XS Or XOriginalCompiled.b <> XB Then Exit Function
    If XOriginalCompiled.inputEpoch <> RPX_InputEpoch Or XOriginalCompiled.subdivisions <> 8 Then Exit Function
    If XOriginalCompiled.loadPhase <> "reference_total" Or XOriginalCompiled.modelKey <> RPX_ExactModel() Then Exit Function
    If RLoadVariable < 0 Or RLoadVariable >= XN Then Exit Function
    For i = 0 To XN - 1
        If Not RPX_SchurFinite(XOriginalCompiled.cost(i)) Then Exit Function
        If i = RLoadVariable Then
            If XOriginalCompiled.cost(i) <> -1# Then Exit Function
        Else
            If XOriginalCompiled.cost(i) <> 0# Then Exit Function
        End If
    Next i
    RPX_AssertInputs "fixed_recovery_begin"
    Call CaptureProgram(program)
    Call RPX_HSDObservationSave(observation)
    oldX = xx: oldY = xy: oldSlack = XSlack: oldDual = XDual
    On Error Resume Next
    proofReady = UBound(RPX_HSDProofX) >= LBound(RPX_HSDProofX)
    err.Clear: On Error GoTo 0
    If proofReady Then oldProof = RPX_HSDProofX
    model = RPX_ExactModel(): oldFactor = RStrengthFactor: oldQ = RSubdivisions: oldEpoch = RPX_InputEpoch
    oldFs = RProgramFs: oldKind = RProgramKind: oldOrder = RPX_OrderMode: oldStatus = XStatus
    oldValue = RValue: oldAudit = RAudit: oldObjective = XObjective
    oldFixed = RFixedWork: oldReference = RReferenceWork: oldDissipation = RDissipation
    oldRejected = RPX_AuditRejected: oldGate = RPhysicalGateAccepted: oldReason = RPhysicalGateReason
    oldPhysical = AFPhysicalGate: oldRawScale = AFRawScale: oldHsdKind = RPX_HSDKind: oldHsdId = RPX_HSDCandidateId
    oldEqError = OriginalEqualityError: oldConeError = OriginalConeError: oldOriginalObjective = OriginalCandidateObjective
    oldResidual = XResidual: oldPrimal = XPrimalResidual: oldDualResidual = XDualResidual: oldGap = XGap: oldIteration = XIteration
    For i = 1 To 4: oldRaw(i) = RRawLowerAudit(i): Next i
    Call RPX_DiagAuditSave(auditSnapshot)
    On Error GoTo Failed
    changed = True: Call RPX_LDLRelease: Call RPX_HLURelease: Call RPX_ResetPreparedState
    Call RPX_AssembleLower
    If reuseOriginalIterate Then
        candidate = NormalizeOriginalIterate(oldX, program.loadId, restored)
        If candidate Then candidate = RPX_OriginalPrimalGate(restored)
        If candidate Then candidate = RPX_FieldGate("lower", restored, factor, 8, OriginalCandidateObjective, FixedLowerRecoveryValue, FixedLowerRecoveryAudit, raw)
        If candidate Then FixedLowerRecoveryRoute = "ORIGINAL_ITERATE_PRIMAL"
    End If
    If Not candidate Then
        FixedLowerRecoveryRoute = "FIXED_LOAD_AFFINE_HSD"
        Call RPX_AssembleLower: Call RPX_FixLowerLoad(REQUIRED_LOAD)
        AFPhysicalGate = True: AFRawScale = RStressScale
        Call RPX_CondenseFixedWitness: Call RPX_Compile
        FixedLowerRecoveryVariables = XN: FixedLowerRecoveryEqualities = XM
        RPX_TestHook "fixed_recovery_solve"
        ' The existing AF HSD candidate gate accepts a restored primal only.
        Call RPX_OptimizeHSD(True, "", seconds)
        candidate = (XStatus = "OPTIMAL")
        FixedLowerRecoveryReason = XStatus
        If candidate Then Call RPX_AffineRestoreX(xx, restored)
        Call RPX_AffineEnd
    End If
    Call RPX_AssembleLower
    If candidate Then
        RPX_TestHook "fixed_recovery_original_gate"
        If model <> RPX_ExactModel() Or factor <> RStrengthFactor Or oldQ <> RSubdivisions Or oldEpoch <> RPX_InputEpoch Then err.Raise 5, "RPX_FixedLowerRecovery", "RECOVERY_MODEL_MISMATCH"
        If UBound(restored) <> XN - 1 Or RLoadVariable <> program.loadId Then err.Raise 9, "RPX_FixedLowerRecovery", "RECOVERY_ORIGINAL_LAYOUT"
        accepted = RPX_OriginalPrimalGate(restored)
        FixedLowerRecoveryEquality = OriginalEqualityError: FixedLowerRecoveryCone = OriginalConeError
        If accepted Then accepted = RPX_FieldGate("lower", restored, factor, 8, OriginalCandidateObjective, FixedLowerRecoveryValue, FixedLowerRecoveryAudit, raw)
        If accepted Then
            If Not RPX_SchurFinite(FixedLowerRecoveryValue) Or Not RPX_SchurFinite(FixedLowerRecoveryAudit) Then accepted = False
            If Abs(FixedLowerRecoveryValue - REQUIRED_LOAD) > 0.0000001 Or FixedLowerRecoveryValue < 1.000001 Then accepted = False
            For i = 1 To 4
                FixedLowerRecoveryRaw(i) = raw(i)
                If Not RPX_SchurFinite(raw(i)) Or raw(i) > 0.0000001 Then accepted = False
            Next i
        End If
        If accepted Then FixedLowerRecoveryReason = "ORIGINAL_FREE_LOAD_PRIMAL_WITNESS" Else FixedLowerRecoveryReason = "ORIGINAL_FREE_LOAD_GATE_REJECTED"
    End If
    GoTo Restore
Failed:
    number = err.number: source = err.source: message = err.description
    accepted = False: FixedLowerRecoveryReason = message
    Resume Restore
Restore:
    On Error GoTo CleanupFailed
    Call RPX_AffineEnd
    If changed Then
        ' Rebuild the literal free-lambda rows so OriginalPrimalGate stays authoritative.
        RStrengthFactor = oldFactor: RSubdivisions = oldQ
        Call RPX_AssembleLower: Call RestoreProgram(program)
    End If
    GoTo RestoreFields
CleanupFailed:
    cleanupNumber = err.number: cleanupSource = err.source: cleanupMessage = err.description
    accepted = False
    Resume RestoreFallback
RestoreFallback:
    On Error GoTo 0
    Call RPX_AffineEnd
    If changed Then Call RestoreProgram(program)
    number = cleanupNumber: source = cleanupSource: message = cleanupMessage
RestoreFields:
    On Error GoTo 0
    xx = oldX: xy = oldY: XSlack = oldSlack: XDual = oldDual
    If proofReady Then RPX_HSDProofX = oldProof Else Erase RPX_HSDProofX
    RProgramKind = oldKind: RProgramFs = oldFs: RPX_OrderMode = oldOrder: XStatus = oldStatus
    RStrengthFactor = oldFactor: RSubdivisions = oldQ
    RValue = oldValue: RAudit = oldAudit: XObjective = oldObjective
    RFixedWork = oldFixed: RReferenceWork = oldReference: RDissipation = oldDissipation
    RPX_AuditRejected = oldRejected: RPhysicalGateAccepted = oldGate: RPhysicalGateReason = oldReason
    AFPhysicalGate = oldPhysical: AFRawScale = oldRawScale: RPX_HSDKind = oldHsdKind: RPX_HSDCandidateId = oldHsdId
    Call RPX_HSDObservationRestore(observation)
    OriginalEqualityError = oldEqError: OriginalConeError = oldConeError: OriginalCandidateObjective = oldOriginalObjective
    XResidual = oldResidual: XPrimalResidual = oldPrimal: XDualResidual = oldDualResidual: XGap = oldGap: XIteration = oldIteration
    For i = 1 To 4: RRawLowerAudit(i) = oldRaw(i): Next i
    Call RPX_DiagAuditRestore(auditSnapshot)
    If model <> RPX_ExactModel() Or oldEpoch <> RPX_InputEpoch Then err.Raise 5, "RPX_FixedLowerRecovery", "RECOVERY_ROLLBACK_IDENTITY"
    RPX_AssertInputs "fixed_recovery_rollback"
    If changed Then Call RPX_Symbolic(XN + XM, XKPtr, XKId)
    If accepted Then recovered = restored
    RPX_TryFixedLowerRecovery = accepted
    RPX_DiagEvent "fixed_lower_recovery", "accepted=" & accepted & ";Fs=" & factor & ";q=8;reason=" & FixedLowerRecoveryReason & ";route=" & FixedLowerRecoveryRoute & ";lambda=" & FixedLowerRecoveryValue & ";raw_eq=" & FixedLowerRecoveryEquality & ";raw_cone=" & FixedLowerRecoveryCone & ";original_fields_layout_restored=1;fixed_dual_used=0;root_created=0"
    If number = 18 Then XCancel = True: err.Raise number, source, message
    If number <> 0 Then
        If number <> 5 Then err.Raise number, source, message
        If Left$(message, 9) <> "RECOVERY_" Then
            If message <> "HSD_DD_PRODUCT_UNDERFLOW" And Not RPX_Recoverable(number, message, "fs_trial", False) Then err.Raise number, source, message
        End If
    End If
End Function
