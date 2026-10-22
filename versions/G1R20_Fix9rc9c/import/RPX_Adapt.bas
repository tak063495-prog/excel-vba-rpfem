Attribute VB_Name = "RPX_Adapt"
' version=20260926_0945 engine=20260926_p2s_f01_f07
Option Explicit
Private ownReuseReason As String
' Fix9 options are local to the thin-chain driver; no settings leak to callers.
Private Type Fix9ThinOptions
    midSeparate As Boolean
    lowerOwnUpper As Boolean
    lowerOwnFactor As Double
    flipFresh As Boolean
    lowerRetries As Long
    retryThetaFactor As Double
    hsdBase As Double
    hsdPerCone As Double
    hsdFactor As Double
    hsdCap As Double
    hsdTimeFraction As Double
    hsdTimeCap As Double
    nominalSeconds As Double
    lowerLargeElements As Long
    lowerLargeThetaFactor As Double
    extraRetryMinSeconds As Double
    dllFactor As Double
    dllRefSeconds As Double
    dllEquTimeFactor As Double
    dllRefCones As Long
    dllExponent As Double
    dllCap As Double
    upperRoundMaxElements As Long
    upperMaxGrowth As Double
    upperDllRetry As Boolean
    upperDllMinElements As Long
    upperDllRetrySeconds As Double
    upperDllRetryIterations As Long
    upperDllRetryEquilibrate As Long
    upperHsdMinElements As Long
    upperHsdFloorSeconds As Double
    lowerHsdMinElements As Long
    lowerHsdFloorSeconds As Double
    upperHsdReserveSeconds As Double
    upperHsdIterations As Long
    lowerExtraRounds As Long
    lowerExtraReserve As Double
    lowerExtraTargetElements As Long
    lowerExtraMaxGrowth As Double
    lowerExtraUpperRound As Long
    lowerExtraActive As Boolean
    lowerExtraProtect As Boolean
    lowerExtraPending As Boolean
    lowerProtectFunded As Boolean
    lowerProtectElements As Long
    lowerProtectCap As Long
    lowerCostSeconds As Double
    lowerCostElements As Long
    lowerCostFsCalls As Long
End Type
Public Type RPX_MeshState
    nodes As Long
    elements As Long
    xy() As Double
    triangles() As Long
    constraints As Object
    field() As Double
    indicator() As Double
    value As Double
    source As String
    searchIncomplete As Boolean
    searchWidth As Double
    searchNote As String
    lowerRootA As Double
    lowerRootB As Double
    lowerRootValid As Boolean
    upperRootA As Double
    upperRootB As Double
    upperRootValid As Boolean
    policy As String
    inputEpoch As Long
    certificateRunId As Long
    q As Long
    stressScale As Double
    objective As Double
    normalizedAudit As Double
    rawYield As Double
    rawEquilibrium As Double
    rawTraction As Double
    rawBoundary As Double
    pairLower As Double
    pairUpper As Double
    pairGap As Double
    velocity() As Double
    complete As Boolean
    materialN As Long
    rigidN As Long
    materials() As RPX_Material
    rigids() As RPX_Rigid
    materialId() As Long
    rigidId() As Long
    body0() As Double
    body1() As Double
    boundary As Object
    phase As String
    gravity As String
    factor As Double
    originalModel As String
    lowerStressStart() As Long
    lowerLoadVariable As Long
    lowerReactions As Object
End Type
Public RBestUpper As RPX_MeshState, RBestLower As RPX_MeshState
Public RAdaptHistory As Collection
Public RAdaptRoute As String
Private adaptCoarseSearch As Boolean
Private fix9ThinEarlyStartTaken As Boolean
Private fix9ThinNominalSeconds As Double
Private fix9UpperProtectActive As Boolean
Private fix9UpperProtectStarted As Double, fix9UpperProtectAllowance As Double
Private fix9UpperProtectCost As Double, fix9UpperProtectReserve As Double
Private Const Fix9UpperProtectMargin As Double = 120#
Private g20BalancedAttempted As Boolean, g20FlowSource As String
Private Const AdaptCoarseTol As Double = 0.001
Private Const AdaptStartRatio As Double = 0.35
Private Const AdaptStartMin As Long = 256
Private Const AdaptStartMax As Long = 600
Private Const AdaptCoarseSubdivisions As Long = 4
Private Const AdaptTransferFraction As Double = 0.06
Private Const AdaptMinGain As Double = 0.05
Private Const AdaptRejectWorsening As Double = 0.1
Private Const AdaptFsHintRatio As Double = 0.1
Private Const AdaptFinalRefineFraction As Double = 0.05
Private Const AdaptMaxElementRatio As Double = 1.05
Private Const AdaptBandNodeFraction As Double = 0.3
Private Const AdaptBandMass As Double = 0.8
Private Const AdaptBandRelocatePasses As Long = 8
Private Const AdaptBandRefineGrowth As Double = 1.3
Private Const AdaptBandRefineFraction As Double = 0.3
Private Const AdaptBandElementCap As Long = 800
Private Const AdaptRobustRefineGrowth As Double = 1.5
Private Const MeshNodeCapacity As Long = 100000
Private Const MeshTriangleCapacity As Long = 220000
Private adaptFinalTarget As Long
Private adaptAccuracyCap As Long, adaptAccuracyRounds As Long
Private adaptHintLower As Double
Private adaptHintUpper As Double
Private adaptHintUpperValid As Boolean
Private adaptHintLowerValid As Boolean
Private adaptFsCallsTotal As Long
Private adaptStatus As String
Private bandEvalFailedFlag As Boolean
Private meshStage As String
Private RBestPair As RPX_MeshState
Private RCurrent As RPX_MeshState
Private currentPairValid As Boolean
' The recovery context belongs to this adaptive run, including its output stage.
Private adaptRecoveryActive As Boolean, adaptRecoveryRunId As Long
Private adaptRecoveryTarget As Long, adaptRecoveryTolerance As Double
Private pilotFallbackUsed As Boolean
Private mesh0Ready As Boolean
Private mesh0 As RPX_MeshState
Private snapModel As String
Private snapReady As Boolean
Private snapRM As Long
Private snapRR As Long
Private snapScale As Double
Private snapBoundN As Long
Private snapMat() As Long
Private snapRigid() As Long
Private snapC() As Double
Private snapPhi() As Double
Private snapPsi() As Double
Private snapGamma() As Double
Private snapCx() As Double
Private snapCy() As Double
Private snapFixed() As Boolean
Private snapL0() As Double
Private snapL1() As Double
Private snapBoundKey() As String
Private adaptStartedAt As Double ' G1R20 Fix5 (#29): see OptionalStageExpired

' G1R20 Fix5 (#29): optional stage deadline (ADAPT_OPTIONAL_DEADLINE_SECONDS,
' default 0 = off). After this many seconds from the start of the adaptive run,
' optional continuation stages (accuracy rounds, lower fan, C0 wedge) are skipped
' so the last certified pair is written instead of being lost to an outer limit.
Private Function OptionalStageExpired(ByVal stage As String) As Boolean
    Dim limit As Double, elapsed As Double
    limit = CDbl(RPX_Setting("ADAPT_OPTIONAL_DEADLINE_SECONDS", 0))
    If Not RPX_SchurFinite(limit) Or limit < 0# Then err.Raise 5, "RPX_Adapt", "INVALID_ADAPT_OPTIONAL_DEADLINE_SECONDS"
    If limit = 0# Or adaptStartedAt <= 0# Then Exit Function
    elapsed = RPX_Clock() - adaptStartedAt
    If elapsed < limit Then Exit Function
    RPX_DiagEvent "optional_stage_skipped", "stage=" & stage & ";elapsed=" & elapsed & ";deadline=" & limit & ";certified_pair_kept=1"
    OptionalStageExpired = True
End Function
Private Function PrecisionFromStart() As Boolean
    Dim value As Double
    value = CDbl(RPX_Setting("ADAPT_PRECISION_FROM_START", 0))
    If Not RPX_SchurFinite(value) Then err.Raise 5, "RPX_Adapt", "INVALID_ADAPT_PRECISION_FROM_START"
    If value <> 0# And value <> 1# Then err.Raise 5, "RPX_Adapt", "INVALID_ADAPT_PRECISION_FROM_START"
    PrecisionFromStart = (value = 1#)
End Function
Private Sub CommonStrengthStress(ByVal factor As Double, ByRef field() As Double)
    Dim sinPhi As Double, cosPhi As Double, tanPhi As Double
    Dim e As Long, k As Long, i As Long, c As Double, phi As Double, sxx As Double, syy As Double, tau As Double
    Dim lhs As Double, rhs As Double, alpha As Double, saved() As Double
    RStrengthFactor = factor: alpha = 1#
    For e = 0 To re - 1
        If RRigidId(e) < 0 Then
            RPX_StrengthTrig e, c, phi, sinPhi, cosPhi, tanPhi: rhs = 2# * c * cosPhi
            For k = 0 To 5
                sxx = RLowerField(RPX_StressId(e, k, 0)): syy = RLowerField(RPX_StressId(e, k, 1)): tau = RLowerField(RPX_StressId(e, k, 2))
                lhs = Sqr((sxx - syy) ^ 2 + 4# * tau * tau) + sinPhi * (sxx + syy)
                If lhs > 0# Then
                    If rhs / lhs < alpha Then alpha = rhs / lhs
                End If
            Next k
        End If
    Next e
    field = RLowerField
    For i = 0 To UBound(field): field(i) = alpha * field(i): Next i
    saved = xx: xx = field: Call RPX_AuditLower: xx = saved
End Sub
Private Function GapRatio(ByVal lower As Double, ByVal upper As Double) As Double
    Dim xscale As Double
    xscale = (Abs(lower) + Abs(upper)) / 2#
    If xscale <= 1E-30 Then
        GapRatio = 1E+30
    Else
        GapRatio = (upper - lower) / xscale
    End If
End Function
Private Function SaveState(ByRef state As RPX_MeshState, ByRef field() As Double, ByRef eta() As Double, ByVal value As Double, ByVal lower As Double, ByVal upper As Double, ByVal label As String, Optional ByVal side As String = "", Optional ByVal expectedObjective As Double = 0#) As Boolean
    Dim pending As RPX_MeshState
    Dim key As Variant
    Dim tmpXY() As Double, tmpTri() As Long, tmpField() As Double, tmpEta() As Double
    Dim i As Long, j As Long
    Dim raw() As Double, factor As Double
    RPX_AssertInputs "save_state"
    If Len(side) > 0 Then
        factor = RStrengthFactor
        If CStr(RPX_Setting("MODE", "")) = "Fs" Then factor = value
        If Not RPX_FieldGate(side, field, factor, RSubdivisions, expectedObjective, pending.objective, pending.normalizedAudit, raw) Then Exit Function
        pending.rawYield = raw(1): pending.rawEquilibrium = raw(2): pending.rawTraction = raw(3): pending.rawBoundary = raw(4)
    End If
    pending.nodes = rn
    pending.elements = re
    pending.materialN = RM: pending.rigidN = RR
    pending.materials = RMaterials: pending.rigids = RRigids
    pending.materialId = RMatId: pending.rigidId = RRigidId
    pending.body0 = RBody0: pending.body1 = RBody1
    Set pending.boundary = RPX_Row()
    For Each key In RBoundary.keys: pending.boundary(key) = RBoundary(key): Next key
    pending.phase = RPX_LoadPhase: pending.gravity = RGravityMode
    pending.factor = RStrengthFactor: If Len(side) > 0 Then pending.factor = factor
    pending.originalModel = RPX_ExactModel(): pending.complete = True
    CaptureWitnessLowerLayout pending
    ReDim tmpXY(0 To rn - 1, 0 To 1)
    For i = 0 To rn - 1
        tmpXY(i, 0) = RXY(i, 0)
        tmpXY(i, 1) = RXY(i, 1)
    Next i
    pending.xy = tmpXY
    ReDim tmpTri(0 To re - 1, 0 To 2)
    For i = 0 To re - 1
        tmpTri(i, 0) = RTri(i, 0)
        tmpTri(i, 1) = RTri(i, 1)
        tmpTri(i, 2) = RTri(i, 2)
    Next i
    pending.triangles = tmpTri
    Set pending.constraints = RPX_Row()
    For Each key In RConstraints.keys: pending.constraints(key) = RConstraints(key): Next key
    ReDim tmpField(LBound(field) To UBound(field))
    For i = LBound(field) To UBound(field): tmpField(i) = field(i): Next i
    pending.field = tmpField
    ReDim tmpEta(LBound(eta) To UBound(eta))
    For i = LBound(eta) To UBound(eta): tmpEta(i) = eta(i): Next i
    pending.indicator = tmpEta
    Dim tmpVel() As Double
    If VectorReady(RUpperField, 0) Then
        ReDim tmpVel(LBound(RUpperField) To UBound(RUpperField))
        For i = LBound(RUpperField) To UBound(RUpperField): tmpVel(i) = RUpperField(i): Next i
        pending.velocity = tmpVel
    End If
    pending.value = value
    pending.source = label
    pending.searchIncomplete = RFsSearchIncomplete
    pending.searchWidth = RFsLargestRootBracket
    pending.searchNote = RFsSearchNote
    pending.lowerRootValid = RPX_ActualRootBracket("lower", pending.lowerRootA, pending.lowerRootB, RLowerFs)
    pending.upperRootValid = RPX_ActualRootBracket("upper", pending.upperRootA, pending.upperRootB, RUpperFs)
    pending.policy = RPX_AnalysisPolicy(): pending.inputEpoch = RPX_InputEpoch
    pending.certificateRunId = RPX_CertRunId
    pending.q = RSubdivisions: pending.stressScale = RStressScale
    pending.pairLower = lower
    pending.pairUpper = upper
    pending.pairGap = GapRatio(lower, upper)
    RPX_TestHook "state_commit"
    RPX_AssertInputs "state_commit"
    state = pending: SaveState = True
End Function
Private Sub CloneState(ByRef dst As RPX_MeshState, ByRef src As RPX_MeshState)
    Dim pending As RPX_MeshState
    Dim key As Variant
    Dim tmpXY() As Double, tmpTri() As Long, tmpField() As Double, tmpEta() As Double
    Dim i As Long, j As Long
    pending.nodes = src.nodes
    pending.elements = src.elements
    pending.complete = src.complete
    If src.complete Then
        pending.materialN = src.materialN: pending.rigidN = src.rigidN
        pending.materials = src.materials: pending.rigids = src.rigids
        pending.materialId = src.materialId: pending.rigidId = src.rigidId
        pending.body0 = src.body0: pending.body1 = src.body1
        Set pending.boundary = RPX_Row()
        For Each key In src.boundary.keys: pending.boundary(key) = src.boundary(key): Next key
        pending.phase = src.phase: pending.gravity = src.gravity: pending.factor = src.factor: pending.originalModel = src.originalModel
        pending.lowerStressStart = src.lowerStressStart: pending.lowerLoadVariable = src.lowerLoadVariable
        Set pending.lowerReactions = CopyWitnessReactions(src.lowerReactions)
    End If
    ReDim tmpXY(0 To src.nodes - 1, 0 To 1)
    For i = 0 To src.nodes - 1
        tmpXY(i, 0) = src.xy(i, 0)
        tmpXY(i, 1) = src.xy(i, 1)
    Next i
    pending.xy = tmpXY
    ReDim tmpTri(0 To src.elements - 1, 0 To 2)
    For i = 0 To src.elements - 1
        tmpTri(i, 0) = src.triangles(i, 0)
        tmpTri(i, 1) = src.triangles(i, 1)
        tmpTri(i, 2) = src.triangles(i, 2)
    Next i
    pending.triangles = tmpTri
    Set pending.constraints = RPX_Row()
    If Not src.constraints Is Nothing Then
        For Each key In src.constraints.keys
            pending.constraints(key) = src.constraints(key)
        Next key
    End If
    ReDim tmpField(LBound(src.field) To UBound(src.field))
    For i = LBound(src.field) To UBound(src.field): tmpField(i) = src.field(i): Next i
    pending.field = tmpField
    ReDim tmpEta(LBound(src.indicator) To UBound(src.indicator))
    For i = LBound(src.indicator) To UBound(src.indicator): tmpEta(i) = src.indicator(i): Next i
    pending.indicator = tmpEta
    Dim tmpVel() As Double
    CopyVelocity tmpVel, src.velocity
    If VectorReady(tmpVel, 0) Then pending.velocity = tmpVel
    pending.value = src.value
    pending.source = src.source
    pending.searchIncomplete = src.searchIncomplete
    pending.searchWidth = src.searchWidth
    pending.searchNote = src.searchNote
    pending.lowerRootA = src.lowerRootA: pending.lowerRootB = src.lowerRootB: pending.lowerRootValid = src.lowerRootValid
    pending.upperRootA = src.upperRootA: pending.upperRootB = src.upperRootB: pending.upperRootValid = src.upperRootValid
    pending.policy = src.policy: pending.inputEpoch = src.inputEpoch: pending.q = src.q: pending.stressScale = src.stressScale
    pending.certificateRunId = src.certificateRunId
    pending.objective = src.objective: pending.normalizedAudit = src.normalizedAudit
    pending.rawYield = src.rawYield: pending.rawEquilibrium = src.rawEquilibrium
    pending.rawTraction = src.rawTraction: pending.rawBoundary = src.rawBoundary
    pending.pairLower = src.pairLower
    pending.pairUpper = src.pairUpper
    pending.pairGap = src.pairGap
    RPX_TestHook "state_commit"
    RPX_AssertInputs "state_commit"
    dst = pending
End Sub
Public Sub RPX_RestoreState(ByRef state As RPX_MeshState)
    Dim key As Variant
    Dim i As Long, j As Long
    RPX_AssertInputs "restore_state"
    RPX_TestHook "restore"
    rn = state.nodes
    re = state.elements
    ReDim RXY(0 To MeshNodeCapacity - 1, 0 To 1)
    For i = 0 To rn - 1
        RXY(i, 0) = state.xy(i, 0)
        RXY(i, 1) = state.xy(i, 1)
    Next i
    ReDim RTri(0 To MeshTriangleCapacity - 1, 0 To 2)
    For i = 0 To re - 1
        RTri(i, 0) = state.triangles(i, 0)
        RTri(i, 1) = state.triangles(i, 1)
        RTri(i, 2) = state.triangles(i, 2)
    Next i
    Set RConstraints = RPX_Row()
    For Each key In state.constraints.keys: RConstraints(key) = state.constraints(key): Next key
    Call RPX_ReadInputs: Call RPX_AssignRegions
End Sub
Public Sub RPX_RestoreWitnessState(ByRef state As RPX_MeshState)
    Dim key As Variant
    RPX_AssertInputs "restore_witness_state"
    If Not state.complete Then err.Raise 5, "RPX_Adapt", "WITNESS_STATE_INCOMPLETE"
    If state.inputEpoch <> RPX_InputEpoch Or state.policy <> RPX_AnalysisPolicy() Then err.Raise 5, "RPX_Adapt", "WITNESS_STATE_INPUT_MISMATCH"
    rn = state.nodes: re = state.elements: RM = state.materialN: RR = state.rigidN
    RXY = state.xy: RTri = state.triangles: RMaterials = state.materials: RRigids = state.rigids
    RMatId = state.materialId: RRigidId = state.rigidId: RBody0 = state.body0: RBody1 = state.body1
    RStressStart = state.lowerStressStart: RLoadVariable = state.lowerLoadVariable
    Set RReactions = CopyWitnessReactions(state.lowerReactions)
    RStressScale = state.stressScale: RStrengthFactor = state.factor: RSubdivisions = state.q
    RGravityMode = state.gravity: RPX_LoadPhase = state.phase
    Set RBoundary = RPX_Row(): Set RConstraints = RPX_Row()
    For Each key In state.boundary.keys: RBoundary(key) = state.boundary(key): Next key
    For Each key In state.constraints.keys: RConstraints(key) = state.constraints(key): Next key
    Call RPX_Topology
    If state.originalModel <> RPX_ExactModel() Then err.Raise 5, "RPX_Adapt", "WITNESS_STATE_RESTORE_MISMATCH"
    RProgramKind = "": Call RPX_ResetPreparedState
End Sub
Private Sub UpdateBestPair(ByRef state As RPX_MeshState)
    If state.pairGap < RBestPair.pairGap Then
        CloneState RBestPair, state
    End If
End Sub
Private Function AcceptCandidate(ByRef oldState As RPX_MeshState, ByRef newState As RPX_MeshState) As Boolean
    If newState.pairGap <= oldState.pairGap Then
        AcceptCandidate = True
        Exit Function
    End If
    If newState.pairGap <= oldState.pairGap * (1# + AdaptRejectWorsening) Then
        If newState.pairLower > oldState.pairLower Or newState.pairUpper < oldState.pairUpper Then
            AcceptCandidate = True
        End If
    End If
End Function
Private Sub RecordHistory(ByVal stage As String, ByVal action As String, ByVal lower As Double, ByVal upper As Double)
    RAdaptHistory.Add Array(stage, re, rn, lower, upper, GapRatio(lower, upper), RBestLower.value, RBestUpper.value, RFsCalls, action)
End Sub
Private Sub Evaluate(ByVal mode As String, ByVal label As String)
    Dim lower As Double, upper As Double, eta() As Double, commonStress() As Double, i As Long
    Dim savedTol As Double, savedSubdivisions As Long, upperObjective As Double, lowerObjective As Double
    Dim errorNumber As Long, errorText As String, errorSource As String, errorLine As Long
    Dim counted As Boolean, searchStarted As Boolean
    On Error GoTo EvalFailed
    meshStage = "evaluate"
    currentPairValid = False
    searchStarted = False
    If XCancel Then err.Raise 18, , "ANALYSIS_CANCELLED"
    savedTol = RFsTolerance
    savedSubdivisions = RSubdivisions
    counted = False
    RPX_TestHook label
    RPX_DiagCandidate label
    Call RPX_ReadInputs
    Call RPX_AssignRegions
    If adaptCoarseSearch And Not PrecisionFromStart() Then
        If RFsTolerance < AdaptCoarseTol Then RFsTolerance = AdaptCoarseTol
        If RSubdivisions > AdaptCoarseSubdivisions Then RSubdivisions = AdaptCoarseSubdivisions
    End If
    RPX_DiagEvent "adapt_eval", "label=" & label & ";coarse=" & adaptCoarseSearch & ";fs_tolerance=" & RFsTolerance & ";quadrature=" & RSubdivisions & ";elements=" & re & ";nodes=" & rn
    If mode = "Fs" Then
        RFsActive = True
        RFsCalls = 0
        searchStarted = True
        RFsSearchIncomplete = False
        RFsLargestRootBracket = 0#
        RFsSearchNote = ""
        Call RPX_TotalReferenceLoads
        If adaptHintUpperValid Then
            upper = RPX_FsEndpoint("upper", adaptHintUpper, AdaptFsHintRatio)
        Else
            upper = RPX_FsEndpoint("upper")
        End If
        RUpperFs = upper: upperObjective = XObjective
        ReDim RUpperField(0 To 12 * re + 3 * RR - 1)
        For i = 0 To UBound(RUpperField)
            RUpperField(i) = xx(i)
        Next i
        On Error GoTo LowerFailedHandler
        If adaptHintLowerValid Then
            lower = RPX_FsEndpoint("lower", adaptHintLower, AdaptFsHintRatio, upper)
        Else
            ' P3-A: only coordinates from the just-audited upper on this mesh.
            lower = RPX_FsEndpoint("lower", 0#, AdaptFsHintRatio, upper)
        End If
        RLowerFs = lower: lowerObjective = XObjective
        ReDim RLowerField(0 To UBound(xx))
        For i = 0 To UBound(RLowerField)
            RLowerField(i) = xx(i)
        Next i
        On Error GoTo EvalFailed
        If RLowerFs > RUpperFs Then err.Raise 5, , "FS_BOUNDS_REVERSED"
        RFsActive = False
        adaptHintUpper = upper
        adaptHintLower = lower
        adaptHintUpperValid = True
        adaptHintLowerValid = True
        CommonStrengthStress upper, commonStress
        RPX_ComputeIndicator eta, RUpperField, commonStress
    Else
        Call RPX_LoadPair
        lower = RLowerLoad
        upper = RUpperLoad: upperObjective = upper: lowerObjective = -lower
        commonStress = RLowerField
        RPX_ComputeIndicator eta, RUpperField, commonStress
    End If
    RPX_AssertInputs "candidate_commit"
    If upper <= RBestUpper.value Then
        If SaveState(RBestUpper, RUpperField, eta, upper, lower, upper, label, "upper", upperObjective) Then RPX_CertFromAdapt "upper"
    End If
    If lower >= RBestLower.value Then
        If SaveState(RBestLower, RLowerField, eta, lower, lower, upper, label, "lower", lowerObjective) Then RPX_CertFromAdapt "lower"
    End If
    SaveState RCurrent, RLowerField, eta, lower, lower, upper, label
    UpdateBestPair RCurrent
    If RBestLower.value > RBestUpper.value Then err.Raise 5, , "ADAPT_BOUNDS_REVERSED"
    adaptFsCallsTotal = adaptFsCallsTotal + RFsCalls
    counted = True
    currentPairValid = True
    RecordHistory label, label, lower, upper
    RPX_DiagEvent "pair_eval_end", "complete_pair=1;label=" & label & ";has_upper=1;has_lower=1;calls=" & RFsCalls
    RPX_DiagEvent "candidate_end", "lower=" & lower & ";upper=" & upper & ";gap=" & GapRatio(lower, upper) & ";best_lower=" & RBestLower.value & ";best_upper=" & RBestUpper.value & ";elements=" & re & ";fs_calls=" & RFsCalls
EvalExit:
    RFsActive = False
    RFsTolerance = savedTol
    RSubdivisions = savedSubdivisions
    Exit Sub
LowerFailedHandler:
    errorNumber = err.number: errorText = err.description: errorSource = err.source: errorLine = Erl
    RPX_ErrorEvidence errorNumber, errorSource, errorText, errorLine, meshStage
    adaptFsCallsTotal = adaptFsCallsTotal + RFsCalls
    counted = True
    currentPairValid = False
    RFsActive = False
    RFsTolerance = savedTol
    RSubdivisions = savedSubdivisions
    On Error GoTo 0
    RPX_DiagEvent "adapt_lower_failed", "label=" & label & ";number=" & errorNumber & ";message=" & errorText
    RPX_DiagEvent "pair_eval_end", "complete_pair=0;label=" & label & ";has_upper=1;has_lower=0;number=" & errorNumber
    If errorNumber = 18 Then err.Raise 18, , "ANALYSIS_CANCELLED"
    err.Raise errorNumber, errorSource, errorText
EvalFailed:
    errorNumber = err.number: errorText = err.description: errorSource = err.source: errorLine = Erl
    RPX_ErrorEvidence errorNumber, errorSource, errorText, errorLine, meshStage
    If searchStarted And Not counted Then adaptFsCallsTotal = adaptFsCallsTotal + RFsCalls
    currentPairValid = False
    RFsActive = False
    RFsTolerance = savedTol
    RSubdivisions = savedSubdivisions
    On Error GoTo 0
    If errorNumber = 18 Then err.Raise 18, , "ANALYSIS_CANCELLED"
    err.Raise errorNumber, errorSource, errorText
End Sub
Private Function TriangleQuality(ByVal e As Long) As Double
    Dim k As Long, a As Long, b As Long, length2 As Double
    For k = 0 To 2
        a = RTri(e, k): b = RTri(e, (k + 1) Mod 3)
        length2 = length2 + (RXY(a, 0) - RXY(b, 0)) ^ 2 + (RXY(a, 1) - RXY(b, 1)) ^ 2
    Next k
    If length2 > 0# Then TriangleQuality = 2# * Sqr(3#) * RPX_Cross(RTri(e, 0), RTri(e, 1), RTri(e, 2)) / length2
End Function
Public Sub RPX_Relocate(ByRef eta() As Double, Optional ByVal passes As Long = 1, Optional ByVal sharp As Boolean = False)
    Dim fixed() As Boolean, around() As Collection, key As Variant, rec As Variant, e As Long, i As Long, k As Long, cell As Variant
    Dim mass As Double, weight As Double, px As Double, py As Double, mx As Double, my As Double, average As Double
    Dim ox As Double, oy As Double, alpha As Double, acceptable As Boolean, oldQuality As Double, minimumQuality As Double
    Dim pass As Long, elementCount As Long, step0 As Double
    If passes < 1 Then passes = 1
    elementCount = re
    For pass = 1 To passes
    If re <> elementCount Then Exit For
    If Not VectorReady(eta, re - 1) Then Exit For
    average = 0#
    ReDim fixed(0 To rn - 1): ReDim around(0 To rn - 1)
    For i = 0 To DPCount - 1: fixed(i) = True: Next i
    For i = 0 To rn - 1: Set around(i) = New Collection: Next i
    For Each key In RConstraints.keys
        rec = RConstraints(key): fixed(rec(0)) = True: fixed(rec(1)) = True
    Next key
    RPX_ProtectNodes fixed
    For e = 0 To re - 1
        average = average + eta(e)
        For k = 0 To 2: around(RTri(e, k)).Add e: Next k
    Next e
    average = average / re
    If average <= 0.000000000000001 Then Exit Sub
    For i = 0 To rn - 1
        If Not fixed(i) And around(i).count > 0 Then
            mass = 0#: mx = 0#: my = 0#: minimumQuality = 1#
            For Each cell In around(i)
                e = CLng(cell)
                If sharp Then
                    weight = eta(e)
                Else
                    weight = (eta(e) + 0.05 * average) ^ (1# / 3#)
                End If
                px = 0#: py = 0#
                For k = 0 To 2: px = px + RXY(RTri(e, k), 0) / 3#: py = py + RXY(RTri(e, k), 1) / 3#: Next k
                mass = mass + weight: mx = mx + weight * px: my = my + weight * py
                oldQuality = TriangleQuality(e): If oldQuality < minimumQuality Then minimumQuality = oldQuality
            Next cell
            If mass > 0# Then
            ox = RXY(i, 0): oy = RXY(i, 1)
            If sharp Then step0 = 1# Else step0 = 0.6
            alpha = step0
            Do
                RXY(i, 0) = ox + alpha * (mx / mass - ox): RXY(i, 1) = oy + alpha * (my / mass - oy)
                RXY(i, 0) = ox: RXY(i, 1) = oy
                acceptable = RPX_SameSoilDestination(i, ox + alpha * (mx / mass - ox), oy + alpha * (my / mass - oy))
                RXY(i, 0) = ox + alpha * (mx / mass - ox): RXY(i, 1) = oy + alpha * (my / mass - oy)
                For Each cell In around(i)
                    If TriangleQuality(CLng(cell)) < 0.8 * minimumQuality Or TriangleQuality(CLng(cell)) < 0.08 Then acceptable = False: Exit For
                Next cell
                If acceptable Then Exit Do
                alpha = alpha / 2#
            Loop While alpha >= 0.01
            If Not acceptable Then RXY(i, 0) = ox: RXY(i, 1) = oy
            End If
        End If
    Next i
    RProgramKind = "": Call RPX_AssignRegions
    Next pass
End Sub
Public Sub RPX_Transfer(ByRef eta() As Double, Optional ByVal fraction As Double = 0.06)
    Dim geometryGuard As RPX_GeometryGuard
    Dim fixed() As Boolean, score() As Double, key As Variant, rec As Variant, i As Long, j As Long, e As Long, k As Long
    Dim low As Long, high As Long, count As Long, maximum As Double, minimum As Double, px As Double, py As Double, valid As Boolean
    RPX_GeometryCapture geometryGuard
    ReDim fixed(0 To rn - 1): ReDim score(0 To rn - 1)
    For i = 0 To DPCount - 1: fixed(i) = True: Next i
    For Each key In RConstraints.keys
        rec = RConstraints(key): fixed(rec(0)) = True: fixed(rec(1)) = True
    Next key
    RPX_ProtectNodes fixed
    For e = 0 To re - 1
        For k = 0 To 2: score(RTri(e, k)) = score(RTri(e, k)) + eta(e): Next k
    Next e
    If fraction <= 0# Then fraction = AdaptTransferFraction
    count = CLng(fraction * rn): If count < 1 Then count = 1
    For i = 1 To count
        low = -1: minimum = 1E+300: high = -1: maximum = 0#
        For j = 0 To rn - 1
            If Not fixed(j) And score(j) < minimum Then minimum = score(j): low = j
        Next j
        For e = 0 To re - 1
            If eta(e) > maximum Then maximum = eta(e): high = e
        Next e
        If low < 0 Or high < 0 Or maximum <= 3# * minimum Then Exit For
        px = 0#: py = 0#
        For k = 0 To 2: px = px + RXY(RTri(high, k), 0) / 3#: py = py + RXY(RTri(high, k), 1) / 3#: Next k
        valid = RPX_SameSoilDestination(low, px, py)
        For j = 0 To rn - 1
            If j <> low Then
                If (px - RXY(j, 0)) ^ 2 + (py - RXY(j, 1)) ^ 2 < 0.04 * RArea(high) Then valid = False: Exit For
            End If
        Next j
        If valid Then RXY(low, 0) = px: RXY(low, 1) = py: fixed(low) = True
        eta(high) = 0#
    Next i
    RProgramKind = "": Call RPX_Delaunay: Call RPX_RecoverSegments: Call RPX_AssignRegions
    RPX_GeometryVerify geometryGuard
End Sub
Private Sub RPX_BuildRedistributedMesh(ByRef baseline As RPX_MeshState, Optional ByVal doTransfer As Boolean = True)
    Dim eta() As Double
    Dim ei As Long
    RPX_RestoreState baseline
    ReDim eta(0 To baseline.elements - 1)
    For ei = 0 To baseline.elements - 1
        eta(ei) = baseline.indicator(ei)
    Next ei
    On Error GoTo RemeshFailed
    Call RPX_Relocate(eta)
    If doTransfer Then
        Call RPX_Transfer(eta, AdaptTransferFraction)
    End If
    Exit Sub
RemeshFailed:
    RPX_DiagEvent "adapt_remesh_failed", "number=" & err.number & ";message=" & err.description
    err.Raise err.number, , err.description
End Sub
Private Function TryRefineCapped(ByRef baseline As RPX_MeshState, ByVal maxElements As Long, ByVal requestedFraction As Double) As Boolean
    Dim fraction As Double, eta() As Double
    fraction = requestedFraction
    If fraction <= 0# Then Exit Function
    Do While fraction >= 0.01
        RPX_RestoreState baseline
        ReDim eta(LBound(baseline.indicator) To UBound(baseline.indicator))
    Dim ei As Long
    For ei = LBound(eta) To UBound(eta): eta(ei) = baseline.indicator(ei): Next ei
        RPX_RefineMesh eta, fraction
        If re <= maxElements Then
            TryRefineCapped = True
            Exit Function
        End If
        fraction = fraction / 2#
    Loop
    RPX_RestoreState baseline
End Function
Private Function AdaptStartTarget(ByVal finalTarget As Long) As Long
    Dim startTarget As Long
    startTarget = CLng(finalTarget * AdaptStartRatio)
    If startTarget < AdaptStartMin Then startTarget = AdaptStartMin
    If startTarget > AdaptStartMax Then startTarget = AdaptStartMax
    If startTarget > finalTarget Then startTarget = finalTarget
    If startTarget < 16 Then startTarget = 16
    AdaptStartTarget = startTarget
End Function
Private Sub RestoreFinalBounds(ByVal mode As String)
    Dim lowerScale As Double, upperScale As Double
    If mode = "Fs" Then
        RLowerFs = RBestLower.value
        RUpperFs = RBestUpper.value
    Else
        RLowerLoad = RBestLower.value
        RUpperLoad = RBestUpper.value
    End If
    If RBestUpper.elements > 0 Then RUpperField = RBestUpper.field
    If RBestLower.elements > 0 Then RLowerField = RBestLower.field
    RFsSearchIncomplete = RBestLower.searchIncomplete Or RBestUpper.searchIncomplete Or RFsSearchIncomplete
    If mode = "Fs" Then
        If Not RBestLower.lowerRootValid Or Not RBestUpper.upperRootValid Then RFsSearchIncomplete = True
        lowerScale = (RBestLower.lowerRootA + RBestLower.lowerRootB) / 2#: If lowerScale < 1# Then lowerScale = 1#
        upperScale = (RBestUpper.upperRootA + RBestUpper.upperRootB) / 2#: If upperScale < 1# Then upperScale = 1#
        If RBestLower.lowerRootB - RBestLower.lowerRootA > RFsTolerance * lowerScale Then RFsSearchIncomplete = True
        If RBestUpper.upperRootB - RBestUpper.upperRootA > RFsTolerance * upperScale Then RFsSearchIncomplete = True
    End If
    RFsLargestRootBracket = RBestLower.searchWidth
    RFsSearchNote = RBestLower.searchNote
    If RBestUpper.searchWidth > RFsLargestRootBracket Then
        RFsLargestRootBracket = RBestUpper.searchWidth
        RFsSearchNote = RBestUpper.searchNote
    End If
End Sub
Private Function VectorReady(ByRef values() As Double, ByVal minimumUpper As Long) As Boolean
    Dim n As Long
    On Error GoTo Missing
    n = UBound(values)
    If n >= minimumUpper Then VectorReady = True
    Exit Function
Missing:
    If err.number <> 9 Then err.Raise err.number, err.source, err.description
End Function
Private Sub CopyVelocity(ByRef dst() As Double, ByRef src() As Double)
    Erase dst
    If Not VectorReady(src, 0) Then Exit Sub
    dst = src
End Sub
Private Sub ComputeShearWeight(ByRef velocity() As Double, ByRef w() As Double)
    Dim e As Long, f As Long, k As Long, j As Long, q As Long, id As Long
    Dim bary(0 To 2) As Double, g(0 To 5, 0 To 1) As Double
    Dim ex As Double, ey As Double, shear As Double
    Dim c As Double, phi As Double, a As Double, weight As Double, dissip As Double
    Dim weights(0 To 2) As Double, s As Double, vx As Double, vy As Double, nx As Double, ny As Double
    Dim jumpNormal As Double, jumpTangent As Double, value As Double
    ReDim w(0 To re - 1)
    If Not VectorReady(velocity, 12 * re - 1) Then Exit Sub
    For e = 0 To re - 1
        If RRigidId(e) < 0 Then
            RPX_Strength e, c, phi
            For q = 0 To 5
                If q < 3 Then
                    a = 0.445948490915965: weight = 0.223381589678011
                Else
                    a = 0.091576213509771: weight = 0.109951743655322
                End If
                For j = 0 To 2: bary(j) = a: Next j
                bary(q Mod 3) = 1# - 2# * a
                RPX_P2Grad e, bary, g
                ex = 0#: ey = 0#: shear = 0#
                For k = 0 To 5
                    ex = ex + velocity(12 * e + 2 * k) * g(k, 0)
                    ey = ey + velocity(12 * e + 2 * k + 1) * g(k, 1)
                    shear = shear + velocity(12 * e + 2 * k) * g(k, 1) + velocity(12 * e + 2 * k + 1) * g(k, 0)
                Next k
                If phi = 0# Then
                    dissip = c * Sqr((ex - ey) ^ 2 + shear ^ 2)
                Else
                    dissip = c * (ex + ey) / Tan(phi)
                End If
                If dissip < 0# Then dissip = 0#
                w(e) = w(e) + RArea(e) * weight * dissip
            Next q
        End If
    Next e
    For id = 0 To RNEdge - 1
        With REdges(id)
            e = .element(0): f = .element(1)
            If f >= 0 Then
                If RRigidId(e) < 0 And RRigidId(f) < 0 And RMatId(e) = RMatId(f) Then
                    RPX_Strength e, c, phi: value = 0#: nx = .normal(0): ny = .normal(1)
                    For q = 0 To 2
                        s = 0.5 + (q - 1) * Sqr(0.15): weight = 5# / 18#: If q = 1 Then weight = 4# / 9#
                        RPX_TraceWeights s, weights
                        vx = 0#: vy = 0#
                        For k = 0 To 2
                            vx = vx + weights(k) * (velocity(12 * f + 2 * .loc(1, k)) - velocity(12 * e + 2 * .loc(0, k)))
                            vy = vy + weights(k) * (velocity(12 * f + 2 * .loc(1, k) + 1) - velocity(12 * e + 2 * .loc(0, k) + 1))
                        Next k
                        jumpNormal = vx * nx + vy * ny
                        jumpTangent = vx * .tangent(0) + vy * .tangent(1)
                        If phi = 0# Then dissip = c * Abs(jumpTangent) Else dissip = c * jumpNormal / Tan(phi)
                        If dissip < 0# Then dissip = 0#
                        value = value + .length * weight * dissip
                    Next q
                    w(e) = w(e) + value / 2#: w(f) = w(f) + value / 2#
                End If
            End If
        End With
    Next id
End Sub
Private Function EvacuateToBand(ByRef w() As Double, ByRef score() As Double) As Long
    Dim fixed() As Boolean, blocked() As Boolean, order() As Long, key As Variant, rec As Variant
    Dim i As Long, j As Long, e As Long, k As Long, low As Long, count As Long, moved As Long
    Dim bandCount As Long, cursor As Long, dest As Long, guard As Long
    Dim total As Double, cum As Double, minimum As Double, px As Double, py As Double, valid As Boolean
    ReDim score(0 To rn - 1): ReDim fixed(0 To rn - 1): ReDim blocked(0 To rn - 1)
    For i = 0 To DPCount - 1: fixed(i) = True: Next i
    For Each key In RConstraints.keys
        rec = RConstraints(key): fixed(rec(0)) = True: fixed(rec(1)) = True
    Next key
    RPX_ProtectNodes fixed
    total = 0#
    For e = 0 To re - 1
        total = total + w(e)
        For k = 0 To 2: score(RTri(e, k)) = score(RTri(e, k)) + w(e): Next k
    Next e
    If total <= 0# Then Exit Function
    ReDim order(0 To re - 1)
    For e = 0 To re - 1: order(e) = e: Next e
    For i = 1 To re - 1
        e = order(i): j = i - 1
        Do While j >= 0
            If w(order(j)) >= w(e) Then Exit Do
            order(j + 1) = order(j): j = j - 1
        Loop
        order(j + 1) = e
    Next i
    cum = 0#: bandCount = 0
    For i = 0 To re - 1
        If w(order(i)) <= 0# Then Exit For
        bandCount = bandCount + 1
        cum = cum + w(order(i))
        If cum >= AdaptBandMass * total Then Exit For
    Next i
    If bandCount < 1 Then Exit Function
    count = CLng(AdaptBandNodeFraction * rn): If count < 1 Then count = 1
    cursor = 0
    Do While moved < count And guard < rn
        guard = guard + 1
        low = -1: minimum = 1E+300
        For j = 0 To rn - 1
            If Not fixed(j) And Not blocked(j) And score(j) < minimum Then minimum = score(j): low = j
        Next j
        If low < 0 Then Exit Do
        valid = False
        For i = 0 To bandCount - 1
            dest = order((cursor + i) Mod bandCount)
            px = 0#: py = 0#
            For k = 0 To 2: px = px + RXY(RTri(dest, k), 0) / 3#: py = py + RXY(RTri(dest, k), 1) / 3#: Next k
            valid = RPX_SameSoilDestination(low, px, py)
            For j = 0 To rn - 1
                If j <> low Then
                    If (px - RXY(j, 0)) ^ 2 + (py - RXY(j, 1)) ^ 2 < 0.04 * RArea(dest) Then valid = False: Exit For
                End If
            Next j
            If valid Then
                RXY(low, 0) = px: RXY(low, 1) = py
                fixed(low) = True
                moved = moved + 1
                cursor = (cursor + i + 1) Mod bandCount
                Exit For
            End If
        Next i
        If Not valid Then blocked(low) = True
    Loop
    EvacuateToBand = moved
End Function
Private Sub PlaceBandCore(ByRef w() As Double, ByVal total As Double, ByVal useIndicator As Boolean)
    Dim geometryGuard As RPX_GeometryGuard
    Dim score() As Double
    Dim e As Long, k As Long, moved As Long
    meshStage = "band_regions"
    RPX_GeometryCapture geometryGuard
    meshStage = "band_evacuate"
    moved = EvacuateToBand(w, score)
    RProgramKind = ""
    meshStage = "band_delaunay"
    Call RPX_Delaunay
    meshStage = "band_recover"
    RPX_TestHook "band_recover"
    Call RPX_RecoverSegments
    meshStage = "band_regions"
    Call RPX_AssignRegions
    If UBound(score) < rn - 1 Then ReDim Preserve score(0 To rn - 1)
    ReDim w(0 To re - 1)
    For e = 0 To re - 1
        For k = 0 To 2: w(e) = w(e) + score(RTri(e, k)): Next k
        w(e) = w(e) / 3#
    Next e
    meshStage = "band_relocate"
    RPX_Relocate w, AdaptBandRelocatePasses, True
    If geometryGuard.active Then RPX_TestHook "rigid_geometry_verify"
    RPX_GeometryVerify geometryGuard
    RPX_DiagEvent "adapt_band", "moved=" & moved & ";elements=" & re & ";nodes=" & rn & ";weight=" & total & ";fallback=" & useIndicator
End Sub
Private Sub RPX_BuildBandMesh(ByRef baseline As RPX_MeshState)
    Dim vel() As Double, w() As Double
    Dim e As Long, total As Double
    Dim useIndicator As Boolean
    meshStage = "band_restore"
    RPX_RestoreState baseline
    CopyVelocity vel, baseline.velocity
    RPX_AdaptiveRefineWeight vel, w
    For e = 0 To re - 1: total = total + w(e): Next e
    If total <= 0# Then
        useIndicator = True
        ReDim w(0 To re - 1)
        If VectorReady(baseline.indicator, re - 1) Then
            For e = 0 To re - 1: w(e) = baseline.indicator(e): total = total + w(e): Next e
        End If
    End If
    PlaceBandCore w, total, useIndicator
    RPX_TestHook "band_complete"
End Sub
Private Function TryBuildPilotBand(ByRef geometry As RPX_MeshState, ByRef velocity() As Double, ByVal fieldStrength As Double) As Boolean
    Dim w() As Double
    Dim e As Long, nNeed As Long
    Dim total As Double, vmin As Double, vmax As Double, bandFactor As Double
    Dim changed As Long
    If XCancel Then err.Raise 18, , "ANALYSIS_CANCELLED"
    nNeed = 12 * geometry.elements + 3 * RR - 1
    If nNeed < 0 Or Not VectorReady(velocity, nNeed) Then err.Raise 5, , "PILOT_VELOCITY_SHORT"
    meshStage = "pilot_restore"
    RPX_RestoreState geometry
    If XCancel Then err.Raise 18, , "ANALYSIS_CANCELLED"
    bandFactor = RStrengthFactor
    meshStage = "pilot_weight"
    RPX_AdaptiveRefineWeight velocity, w
    total = 0#
    If re <= 0 Then err.Raise 5, , "PILOT_WEIGHT_EMPTY"
    vmin = w(0): vmax = w(0)
    For e = 0 To re - 1
        If w(e) <> w(e) Or w(e) > 1E+99 Then err.Raise 5, , "PILOT_WEIGHT_NONFINITE"
        If w(e) < 0# Then err.Raise 5, , "PILOT_WEIGHT_NEGATIVE"
        If w(e) < vmin Then vmin = w(e)
        If w(e) > vmax Then vmax = w(e)
        total = total + w(e)
    Next e
    RPX_DiagEvent "pilot_weight", "valid=" & IIf(total > 0#, "1", "0") & ";sum=" & total & ";min=" & vmin & ";max=" & vmax & ";field_strength=" & fieldStrength & ";band_strength_factor=" & bandFactor
    If total <= 0# Then Exit Function
    PlaceBandCore w, total, False
    RPX_TestHook "band_complete"
    If GeometryChanged(geometry) Then changed = 1
    RPX_DiagEvent "pilot_remesh_end", "success=1;moved=see_adapt_band;elements=" & re & ";nodes=" & rn & ";geometry_changed=" & changed & ";failure_stage="
    TryBuildPilotBand = True
End Function

Private Function TryShearRefine(ByRef baseline As RPX_MeshState, ByVal maxElements As Long, Optional ByVal forceLegacy As Boolean = False, Optional ByVal useIndicator As Boolean = False) As Boolean
    Dim fraction As Double, score() As Double, w() As Double, vel() As Double
    Dim e As Long, total As Double, Failed As Long
    Dim failureText As String, failureSource As String, failureLine As Long
    Dim robustRefine As Boolean
    robustRefine = RPX_RobustMeshPolicy() And Not forceLegacy
    fraction = AdaptBandRefineFraction
    If maxElements <= baseline.elements Then Exit Function
    Do While fraction >= 0.01
        meshStage = "refine_restore"
        RPX_RestoreState baseline
        CopyVelocity vel, baseline.velocity
        RPX_AdaptiveRefineWeight vel, w, robustRefine
        total = 0#
        For e = 0 To re - 1: total = total + w(e): Next e
        ReDim score(0 To re - 1)
        If useIndicator And VectorReady(baseline.indicator, re - 1) Then
            For e = 0 To re - 1: score(e) = baseline.indicator(e) * RArea(e): Next e
        ElseIf total > 0# Then
            For e = 0 To re - 1
                If robustRefine Then score(e) = w(e) Else score(e) = w(e) * RArea(e)
            Next e
        ElseIf VectorReady(baseline.indicator, re - 1) Then
            For e = 0 To re - 1: score(e) = baseline.indicator(e) * RArea(e): Next e
        End If
        If robustRefine Then
            fraction = RPX_RefineBudgetFraction(score, maxElements, fraction)
            If fraction <= 0# Then RPX_RestoreState baseline: Exit Function
        End If
        meshStage = "refine"
        On Error Resume Next
        err.Clear
        If forceLegacy Then
            RPX_TestHook "refine"
            RPX_RefineMesh score, fraction, True
        Else
            RPX_RefineAttempt score, fraction
        End If
        Failed = err.number: failureText = err.description: failureSource = err.source: failureLine = Erl
        On Error GoTo 0
        If Failed = 18 Or (Failed = 0 And XCancel) Then err.Raise 18, , "ANALYSIS_CANCELLED"
        If Failed <> 0 Then
            RPX_ErrorEvidence Failed, failureSource, failureText, failureLine, meshStage
            If Not RPX_Recoverable(Failed, failureText, meshStage, True) Then err.Raise Failed, failureSource, failureText
        End If
        If Failed = 0 Then
            If re > baseline.elements And re <= maxElements Then
                TryShearRefine = True
                Exit Function
            End If
        End If
        fraction = fraction / 2#
    Loop
    RPX_RestoreState baseline
End Function

Private Function GeometryChanged(ByRef state As RPX_MeshState) As Boolean
    Dim i As Long
    If state.nodes <> rn Or state.elements <> re Then GeometryChanged = True: Exit Function
    If rn > 0 Then
        For i = 0 To rn - 1
            If state.xy(i, 0) <> RXY(i, 0) Or state.xy(i, 1) <> RXY(i, 1) Then GeometryChanged = True: Exit Function
        Next i
    End If
    If re > 0 Then
        For i = 0 To re - 1
            If state.triangles(i, 0) <> RTri(i, 0) Or state.triangles(i, 1) <> RTri(i, 1) Or state.triangles(i, 2) <> RTri(i, 2) Then
                GeometryChanged = True: Exit Function
            End If
        Next i
    End If
End Function
Private Sub BoundarySignature(ByRef keys() As String, ByRef n As Long)
    Dim i As Long, c As Long
    Dim tmp() As String
    n = 0
    c = 0
    If RNEdge <= 0 Then
        ReDim keys(0 To 0)
        Exit Sub
    End If
    ReDim tmp(0 To RNEdge - 1)
    For i = 0 To RNEdge - 1
        If REdges(i).element(1) < 0 Then
            tmp(c) = CStr(REdges(i).a) & "," & CStr(REdges(i).b) & "," & REdges(i).kind & "," & CStr(REdges(i).rigidId)
            c = c + 1
        End If
    Next i
    n = c
    If n <= 0 Then
        ReDim keys(0 To 0)
        Exit Sub
    End If
    ReDim keys(0 To n - 1)
    For i = 0 To n - 1
        keys(i) = tmp(i)
    Next i
End Sub
Private Sub AddBoundaryCount(ByVal bag As Object, ByVal key As String)
    If bag.Exists(key) Then bag(key) = CLng(bag(key)) + 1 Else bag(key) = 1
End Sub
Private Sub CapturePilotIdentity()
    Dim i As Long, j As Long
    RPX_AssertInputs "m0_capture_raw"
    If RPX_LoadPhase <> "raw" Then err.Raise 5, , "M0_MISMATCH capture_load_phase"
    snapModel = RPX_ExactModel()
    snapRM = RM
    snapRR = RR
    snapScale = RStressScale
    If re > 0 Then
        ReDim snapMat(0 To re - 1)
        ReDim snapRigid(0 To re - 1)
        For i = 0 To re - 1
            snapMat(i) = RMatId(i)
            snapRigid(i) = RRigidId(i)
        Next i
    End If
    If RM > 0 Then
        ReDim snapC(0 To RM - 1)
        ReDim snapPhi(0 To RM - 1)
        ReDim snapPsi(0 To RM - 1)
        ReDim snapGamma(0 To RM - 1)
        For i = 0 To RM - 1
            snapC(i) = RMaterials(i).cohesion
            snapPhi(i) = RMaterials(i).friction
            snapPsi(i) = RMaterials(i).dilation
            snapGamma(i) = RMaterials(i).gamma
        Next i
    End If
    If RR > 0 Then
        ReDim snapCx(0 To RR - 1)
        ReDim snapCy(0 To RR - 1)
        ReDim snapFixed(0 To RR - 1, 0 To 2)
        ReDim snapL0(0 To RR - 1, 0 To 2)
        ReDim snapL1(0 To RR - 1, 0 To 2)
        For i = 0 To RR - 1
            snapCx(i) = RRigids(i).center(0)
            snapCy(i) = RRigids(i).center(1)
            For j = 0 To 2
                snapFixed(i, j) = RRigids(i).fixed(j)
                snapL0(i, j) = RRigids(i).load0(j)
                snapL1(i, j) = RRigids(i).load1(j)
            Next j
        Next i
    End If
    BoundarySignature snapBoundKey, snapBoundN
    snapReady = True
End Sub
Private Function ConstraintKeysMatch(ByVal saved As Object, ByRef detail As String) As Boolean
    Dim key As Variant, a As Variant, b As Variant
    detail = ""
    If saved Is Nothing Or RConstraints Is Nothing Then
        detail = "constraints_missing"
        Exit Function
    End If
    If saved.count <> RConstraints.count Then
        detail = "constraint_count"
        Exit Function
    End If
    For Each key In saved.keys
        If Not RConstraints.Exists(key) Then
            detail = "constraint_key"
            Exit Function
        End If
        a = saved(key)
        b = RConstraints(key)
        If a(0) <> b(0) Or a(1) <> b(1) Then
            detail = "constraint_nodes"
            Exit Function
        End If
    Next key
    ConstraintKeysMatch = True
End Function
Private Function BoundaryMatches(ByRef detail As String) As Boolean
    Dim live() As String
    Dim n As Long, i As Long
    Dim savedCount As Object, liveCount As Object
    Dim key As Variant
    detail = ""
    BoundarySignature live, n
    If n <> snapBoundN Then
        detail = "boundary_count"
        Exit Function
    End If
    Set savedCount = CreateObject("Scripting.Dictionary")
    Set liveCount = CreateObject("Scripting.Dictionary")
    If n > 0 Then
        For i = 0 To n - 1
            AddBoundaryCount savedCount, snapBoundKey(i)
            AddBoundaryCount liveCount, live(i)
        Next i
    End If
    If savedCount.count <> liveCount.count Then
        detail = "boundary_set"
        Exit Function
    End If
    For Each key In savedCount.keys
        If Not liveCount.Exists(key) Then
            detail = "boundary_key"
            Exit Function
        End If
        If CLng(savedCount(key)) <> CLng(liveCount(key)) Then
            detail = "boundary_multiset"
            Exit Function
        End If
    Next key
    BoundaryMatches = True
End Function
Private Sub VerifyPilotIdentity(ByRef state As RPX_MeshState)
    Dim detail As String
    Dim i As Long, j As Long
    RPX_AssertInputs "m0_verify_raw"
    If RPX_LoadPhase <> "raw" Then err.Raise 5, , "M0_MISMATCH restore_load_phase"
    RPX_TestHook "m0_verify"
    detail = ""
    If Not snapReady Or Not mesh0Ready Then
        detail = "snapshot_missing"
    ElseIf RPX_ExactModel() <> snapModel Then
        detail = "exact_model_loads_raw"
    ElseIf GeometryChanged(state) Then
        detail = "geometry"
    ElseIf RM <> snapRM Or RR <> snapRR Then
        detail = "counts"
    ElseIf RStressScale <> snapScale Then
        detail = "stress_scale"
    ElseIf Not ConstraintKeysMatch(state.constraints, detail) Then
        If Len(detail) = 0 Then detail = "constraints"
    ElseIf re > 0 Then
        For i = 0 To re - 1
            If RMatId(i) <> snapMat(i) Or RRigidId(i) <> snapRigid(i) Then
                detail = "region_id"
                Exit For
            End If
        Next i
    End If
    If Len(detail) = 0 And RM > 0 Then
        For i = 0 To RM - 1
            If RMaterials(i).cohesion <> snapC(i) Or RMaterials(i).friction <> snapPhi(i) Or RMaterials(i).dilation <> snapPsi(i) Or RMaterials(i).gamma <> snapGamma(i) Then
                detail = "materials"
                Exit For
            End If
        Next i
    End If
    If Len(detail) = 0 And RR > 0 Then
        For i = 0 To RR - 1
            If RRigids(i).center(0) <> snapCx(i) Or RRigids(i).center(1) <> snapCy(i) Then detail = "rigid_center"
            If Len(detail) = 0 Then
                For j = 0 To 2
                    If RRigids(i).fixed(j) <> snapFixed(i, j) Then detail = "rigid_fixed"
                    If RRigids(i).load0(j) <> snapL0(i, j) Or RRigids(i).load1(j) <> snapL1(i, j) Then detail = "rigid_load"
                    If Len(detail) > 0 Then Exit For
                Next j
            End If
            If Len(detail) > 0 Then Exit For
        Next i
    End If
    If Len(detail) = 0 Then
        If Not BoundaryMatches(detail) Then
            If Len(detail) = 0 Then detail = "boundary"
        End If
    End If
    If Len(detail) > 0 Then
        RPX_DiagEvent "geometry_guard", "match=0;detail=" & Replace(detail, ";", ",")
        err.Raise 5, , "M0_MISMATCH " & detail
    End If
    RPX_DiagEvent "geometry_guard", "match=1"
End Sub
Private Function PilotRecoverable(ByVal number As Long, ByVal message As String) As Boolean
    PilotRecoverable = RPX_Recoverable(number, message, meshStage, mesh0Ready And snapReady)
End Function
Private Sub SaveGeometry(ByRef state As RPX_MeshState)
    Dim key As Variant
    Dim tmpXY() As Double, tmpTri() As Long
    Dim i As Long
    RPX_AssertInputs "save_state"
    state.nodes = rn
    state.elements = re
    ReDim tmpXY(0 To rn - 1, 0 To 1)
    For i = 0 To rn - 1
        tmpXY(i, 0) = RXY(i, 0): tmpXY(i, 1) = RXY(i, 1)
    Next i
    state.xy = tmpXY
    ReDim tmpTri(0 To re - 1, 0 To 2)
    For i = 0 To re - 1
        tmpTri(i, 0) = RTri(i, 0): tmpTri(i, 1) = RTri(i, 1): tmpTri(i, 2) = RTri(i, 2)
    Next i
    state.triangles = tmpTri
    Set state.constraints = RPX_Row()
    If Not RConstraints Is Nothing Then
        For Each key In RConstraints.keys: state.constraints(key) = RConstraints(key): Next key
    End If
    state.source = "geometry"
    state.pairGap = 1E+30
    state.value = 0#
    CapturePilotIdentity
End Sub
Private Sub RecordPilotHistory(ByVal calls As Long, ByVal upper As Double)
    RAdaptHistory.Add Array("pilot_upper_0", re, rn, Empty, upper, Empty, Empty, Empty, calls, "pilot_only")
End Sub
Private Sub EvaluateUpperPilot(ByVal mode As String)
    Dim i As Long, upper As Double
    Dim savedTol As Double, savedSubdivisions As Long
    Dim errorNumber As Long, errorText As String, errorSource As String, errorLine As Long
    On Error GoTo PilotEvalFailed
    meshStage = "pilot_upper"
    savedTol = RFsTolerance
    savedSubdivisions = RSubdivisions
    currentPairValid = False
    RFsCalls = 0
    If XCancel Then err.Raise 18, , "ANALYSIS_CANCELLED"
    RPX_DiagCandidate "pilot_upper_0"
    RPX_DiagEvent "pilot_begin", "elements=" & re & ";nodes=" & rn
    Call RPX_ReadInputs
    Call RPX_AssignRegions
    If RFsTolerance < AdaptCoarseTol Then RFsTolerance = AdaptCoarseTol
    If RSubdivisions > AdaptCoarseSubdivisions Then RSubdivisions = AdaptCoarseSubdivisions
    RFsActive = True
    RFsSearchIncomplete = False
    RFsLargestRootBracket = 0#
    RFsSearchNote = ""
    Call RPX_TotalReferenceLoads
    upper = RPX_FsEndpoint("upper")
    If RFsSearchIncomplete Then err.Raise 5, , "PILOT_UPPER_INCOMPLETE"
    If XStatus <> "OPTIMAL" Then err.Raise 5, , "PILOT_UPPER_NOT_OPTIMAL"
    If upper <= 0# Or upper <> upper Then err.Raise 5, , "PILOT_UPPER_NOT_POSITIVE"
    RUpperFs = upper
    ReDim RUpperField(0 To 12 * re + 3 * RR - 1)
    For i = 0 To UBound(RUpperField): RUpperField(i) = xx(i): Next i
    adaptHintUpper = upper
    adaptHintUpperValid = True
    adaptHintLowerValid = False
    RecordPilotHistory RFsCalls, upper
    RPX_DiagEvent "pilot_upper_end", "success=1;upper=" & upper & ";status=" & XStatus & ";quadrature=" & RSubdivisions & ";tol=" & RFsTolerance & ";calls=" & RFsCalls & ";search_width=" & RFsLargestRootBracket & ";initial_lower_skipped=1"
    adaptFsCallsTotal = adaptFsCallsTotal + RFsCalls
    RFsActive = False
    RFsTolerance = savedTol
    RSubdivisions = savedSubdivisions
    Exit Sub
PilotEvalFailed:
    errorNumber = err.number: errorText = err.description: errorSource = err.source: errorLine = Erl
    RPX_ErrorEvidence errorNumber, errorSource, errorText, errorLine, meshStage
    On Error GoTo 0
    adaptFsCallsTotal = adaptFsCallsTotal + RFsCalls
    currentPairValid = False
    RFsActive = False
    RFsTolerance = savedTol
    RSubdivisions = savedSubdivisions
    RPX_DiagEvent "pilot_upper_end", "success=0;upper=;status=;calls=" & RFsCalls & ";number=" & errorNumber & ";message=" & Replace(errorText, ";", ",")
    If errorNumber = 18 Then err.Raise 18, , "ANALYSIS_CANCELLED"
    err.Raise errorNumber, errorSource, errorText
End Sub

Public Function RPX_AdaptPairComplete() As Boolean
    RPX_AdaptPairComplete = currentPairValid
End Function
Public Sub RPX_TestM0Guard(ByVal mutation As String)
    Dim state As RPX_MeshState
    If RPX_Busy Or CLng(RPX_Setting("TEST_INJECTION", 0)) <> 1 Then err.Raise 5, , "TEST_INJECTION_DISABLED"
    SaveGeometry state: mesh0Ready = True
    Call RPX_TotalReferenceLoads
    RPX_RestoreState state
    Select Case mutation
        Case "face0": REdges(0).load0(0) = REdges(0).load0(0) + 1#
        Case "face1": REdges(0).load1(1) = REdges(0).load1(1) + 1#
        Case "body": RBody0(0, 1) = RBody0(0, 1) + 1#
        Case "center": RRigids(0).center(0) = RRigids(0).center(0) + 1#
        Case "fixed": RRigids(0).fixed(0) = Not RRigids(0).fixed(0)
        Case "force": RRigids(0).load0(0) = RRigids(0).load0(0) + 1#
        Case "moment": RRigids(0).load1(2) = RRigids(0).load1(2) + 1#
    End Select
    VerifyPilotIdentity state
End Sub
' Optional continuation after the ordinary budget, using the requested final
' quadrature and Fs tolerance. Every candidate still enters Evaluate/FieldGate.
' G1R20 Fix7 (candidate, ADAPT_ACCURACY_SKIP_ESTIMATE=1): when upper programs
' already go straight to HSD and the last upper HSD time scaled linearly to the
' candidate size exceeds the ordinary limit (with endgame grace), the round would
' end in HSD_SOLVE_BUDGET_EXCEEDED; stop it before spending that time. The kept
' result is the same as after that failure (old witnesses, search incomplete).
' G1R20 Fix8 rc2 (ADAPT_ACCURACY_SKIP_ESTIMATE=2, default): machine-independent
' element-count gate. A round whose candidate exceeds ADAPT_ACCURACY_SKIP_ELEMENTS
' (default 1000) is skipped; an admitted round gets the HSD safety floor
' ADAPT_ACCURACY_HSD_SECONDS (default 900 s) so a slow machine does not change
' the path. Calibrated on logged runs: P10A 582/790 admitted, 1109 skipped (it
' only ever ended in HSD_SOLVE_BUDGET_EXCEEDED); P05A never has the conic skip.
' Mode 1 restores the Fix7 wall-clock estimate.
Private Function AccuracySkipMode() As Long
    AccuracySkipMode = CLng(RPX_Setting("ADAPT_ACCURACY_SKIP_ESTIMATE", 2))
End Function
Private Function AccuracyHsdFloor() As Double
    Dim floorSeconds As Double
    If AccuracySkipMode() <> 2 Then Exit Function
    If Not RPX_UpperConicSkipActive() Then Exit Function
    floorSeconds = CDbl(RPX_Setting("ADAPT_ACCURACY_HSD_SECONDS", 900))
    If Not RPX_SchurFinite(floorSeconds) Or floorSeconds < 0# Or floorSeconds > 3600# Then err.Raise 5, "RPX_Adapt", "ADAPT_ACCURACY_HSD_SECONDS_INVALID"
    AccuracyHsdFloor = floorSeconds
End Function
Private Function AccuracyUpperHopeless() As Boolean
    Dim estimate As Double, limit As Double, maxElements As Long
    If AccuracySkipMode() = 2 Then
        If Not RPX_UpperConicSkipActive() Or re <= 0 Then Exit Function
        maxElements = CLng(RPX_Setting("ADAPT_ACCURACY_SKIP_ELEMENTS", 1000))
        If maxElements <= 0 Then Exit Function
        RPX_DiagEvent "accuracy_upper_estimate", "mode=elements;elements=" & re & ";max_elements=" & maxElements & ";last_seconds=" & RPX_LastUpperHsdSeconds & ";last_elements=" & RPX_LastUpperHsdElements
        AccuracyUpperHopeless = (re > maxElements)
        Exit Function
    End If
    If AccuracySkipMode() <> 1 Then Exit Function
    If Not RPX_UpperConicSkipActive() Then Exit Function
    If RPX_LastUpperHsdElements <= 0 Or RPX_LastUpperHsdSeconds <= 0# Or re <= 0 Then Exit Function
    estimate = RPX_LastUpperHsdSeconds * re / RPX_LastUpperHsdElements
    limit = RPX_HsdOrdinaryLimit(re)
    RPX_DiagEvent "accuracy_upper_estimate", "elements=" & re & ";estimate_seconds=" & estimate & ";limit_seconds=" & limit & ";last_seconds=" & RPX_LastUpperHsdSeconds & ";last_elements=" & RPX_LastUpperHsdElements
    AccuracyUpperHopeless = (estimate > limit)
End Function
' G1R20 Fix7 (candidate, ADAPT_FINAL_REUSE=1): with ADAPT_PRECISION_FROM_START
' the cycles already used the final q and Fs tolerance, so "final" on the best
' pair's own mesh re-solves identical programs (P10A: identical objectives).
Private Function FinalReuseAllowed(ByVal mode As String) As Boolean
    If CLng(RPX_Setting("ADAPT_FINAL_REUSE", 1)) <> 1 Then Exit Function
    If mode <> "Fs" Or Not PrecisionFromStart() Then Exit Function
    If RBestPair.elements <= 0 Or Not currentPairValid Or Not RCurrent.complete Then Exit Function
    If RCurrent.searchIncomplete Or Not RCurrent.lowerRootValid Or Not RCurrent.upperRootValid Then Exit Function
    If RCurrent.q <> RSubdivisions Or RCurrent.inputEpoch <> RPX_InputEpoch Or RCurrent.policy <> RPX_AnalysisPolicy() Then Exit Function
    If RCurrent.elements <> RBestPair.elements Or RCurrent.pairLower <> RBestPair.pairLower Or RCurrent.pairUpper <> RBestPair.pairUpper Then Exit Function
    FinalReuseAllowed = True
End Function
Private Sub ContinueAccuracy(ByVal mode As String, ByVal targetGap As Double)
    Dim baseline As RPX_MeshState
    Dim round As Long, cap As Long, stalls As Long, attemptedRounds As Long
    Dim beforeGap As Double, afterGap As Double, gain As Double
    Dim number As Long, source As String, message As String, errorLine As Long
    Dim useIndicator As Boolean, refined As Boolean
    If adaptAccuracyRounds = 0 Or adaptAccuracyCap <= adaptFinalTarget Then Exit Sub
    RPX_DiagEvent "accuracy_begin", "cap=" & adaptAccuracyCap & ";rounds=" & adaptAccuracyRounds & ";quadrature=" & RSubdivisions & ";fs_tolerance=" & RFsTolerance
    adaptCoarseSearch = False
    For round = 1 To adaptAccuracyRounds
        If XCancel Then err.Raise 18, , "ANALYSIS_CANCELLED"
        If RPX_TargetMet(mode, RBestLower.value, RBestUpper.value, currentPairValid) Then Exit For
        If Not currentPairValid Or RCurrent.elements <= 0 Then Exit For
        If OptionalStageExpired("accuracy_" & round) Then Exit For
        CloneState baseline, RCurrent
        If baseline.elements >= adaptAccuracyCap Then Exit For
        cap = CLng(baseline.elements * AdaptRobustRefineGrowth)
        If cap <= baseline.elements Then cap = baseline.elements + 1
        If cap > adaptAccuracyCap Then cap = adaptAccuracyCap
        beforeGap = GapRatio(RBestLower.value, RBestUpper.value)
        attemptedRounds = round
        useIndicator = (stalls > 0)
        RPX_DiagEvent "accuracy_action", "round=" & round & ";baseline_elements=" & baseline.elements & ";cap=" & cap & ";weight=" & IIf(useIndicator, "boundgap", "kinematic")
        On Error GoTo CandidateFailed
        RPX_DiagBegin dpMesh
        refined = TryShearRefine(baseline, cap, False, useIndicator)
        RPX_DiagEnd dpMesh
        If Not refined Then
            On Error GoTo 0
            RPX_RestoreState baseline: CloneState RCurrent, baseline
            RPX_DiagEvent "accuracy_stop", "reason=no_admissible_refinement;round=" & round
            Exit For
        End If
        If AccuracyUpperHopeless() Then
            On Error GoTo 0
            RPX_RestoreState baseline: CloneState RCurrent, baseline
            currentPairValid = True
            RFsSearchIncomplete = True
            RPX_DiagEvent "accuracy_stop", "reason=upper_solve_estimate_exceeds_limit;round=" & round & ";old_witnesses_kept=1"
            Exit For
        End If
        meshStage = "evaluate"
        RPX_HsdBudgetFloor = AccuracyHsdFloor()
        If RPX_HsdBudgetFloor > 0# Then RPX_DiagEvent "accuracy_hsd_floor", "round=" & round & ";elements=" & re & ";floor_seconds=" & RPX_HsdBudgetFloor
        Call Evaluate(mode, "accuracy_" & round)
        RPX_HsdBudgetFloor = 0#
        If Not AcceptCandidate(baseline, RCurrent) Then
            RPX_RestoreState baseline: CloneState RCurrent, baseline
            RPX_DiagEvent "accuracy_rollback", "round=" & round & ";reason=pair_worsened;best_side_witnesses_kept=1"
        End If
        Call RestoreFinalBounds(mode)
        afterGap = GapRatio(RBestLower.value, RBestUpper.value)
        gain = 0#: If beforeGap > 0# Then gain = (beforeGap - afterGap) / beforeGap
        If gain <= 0.000001 Then stalls = stalls + 1 Else stalls = 0
        RPX_DiagEvent "accuracy_round", "round=" & round & ";elements=" & RCurrent.elements & ";gap=" & afterGap & ";gain=" & gain & ";stalls=" & stalls
        On Error GoTo 0
        If stalls >= 2 Then
            RPX_DiagEvent "accuracy_stop", "reason=stalled_after_indicator_retry;round=" & round
            Exit For
        End If
    Next round
    RPX_DiagEvent "accuracy_end", "attempted_rounds=" & attemptedRounds & ";elements=" & RCurrent.elements & ";gap=" & GapRatio(RBestLower.value, RBestUpper.value)
    Exit Sub
CandidateFailed:
    number = err.number: source = err.source: message = err.description: errorLine = Erl
    RPX_HsdBudgetFloor = 0#
    On Error GoTo 0
    RPX_ErrorEvidence number, source, message, errorLine, meshStage
    If number = 18 Or XCancel Then err.Raise 18, source, "ANALYSIS_CANCELLED"
    If InStr(1, message, "HSD_STABLE_UNAVAILABLE_MEMORY", vbBinaryCompare) = 1 Then err.Raise number, source, message
    If Not RPX_Recoverable(number, message, meshStage, True) Then err.Raise number, source, message
    ' Restore the complete baseline; never convert an UNKNOWN into a bound.
    On Error GoTo RestoreFailed
    RPX_DiagEnd dpMesh
    RPX_RestoreState baseline: CloneState RCurrent, baseline
    currentPairValid = True
    RFsSearchIncomplete = True
    RPX_DiagEvent "accuracy_stop", "reason=numerical_or_geometry_failure;round=" & round & ";old_witnesses_kept=1;message=" & message
    On Error GoTo 0
    Exit Sub
RestoreFailed:
    number = err.number: source = err.source: message = err.description
    On Error GoTo 0
    err.Raise number, source, "ADAPT_RESTORE_FAILED " & message
End Sub

' The lower candidate owns its new mesh and root. The upper witness is unchanged.
Private Sub ContinueLowerAngularFan(ByVal mode As String)
    Dim baseline As RPX_MeshState, seed As RPX_MeshState
    Dim parent() As Long, field() As Double, eta() As Double
    Dim rounds As Long, round As Long, cap As Long, savedQ As Long, savedTol As Double, uniformTarget As Long
    Dim lower As Double, objective As Double, accepted As Boolean, oldIncomplete As Boolean
    Dim number As Long, source As String, message As String, errorLine As Long, i As Long
    rounds = CLng(RPX_Setting("ADAPT_LOWER_FAN_ROUNDS", 0))
    uniformTarget = CLng(RPX_Setting("ADAPT_LOWER_TARGET", 0))
    If rounds < 0 Or rounds > 4 Then err.Raise 5, "RPX_Adapt", "INVALID_ADAPT_LOWER_FAN_ROUNDS"
    If uniformTarget <> 0 And (uniformTarget < 16 Or uniformTarget > adaptAccuracyCap) Then err.Raise 5, "RPX_Adapt", "INVALID_ADAPT_LOWER_TARGET"
    If (rounds = 0 And uniformTarget = 0) Or mode <> "Fs" Then Exit Sub
    If Not currentPairValid Or Not RBestLower.complete Or Not RBestUpper.complete Then Exit Sub
    If OptionalStageExpired("lower_angular_fan") Then Exit Sub
    CloneState baseline, RCurrent
    savedQ = RSubdivisions: savedTol = RFsTolerance: oldIncomplete = RFsSearchIncomplete
    cap = adaptAccuracyCap
    On Error GoTo CandidateFailed
    For round = IIf(uniformTarget > 0, 0, 1) To rounds
        If XCancel Then err.Raise 18, , "ANALYSIS_CANCELLED"
        CloneState seed, RBestLower
        Call RPX_RestoreWitnessState(seed)
        If round = 0 Then
            meshStage = "band_regions"
            Call RPX_ReadInputs: DTarget = uniformTarget
            Call RPX_GenerateMesh
            DTarget = adaptFinalTarget
            If re > cap Then Exit For
            Call RPX_TotalReferenceLoads
            RPX_DiagEvent "lower_uniform_candidate", "target=" & uniformTarget & ";elements=" & re & ";upper_mesh_unchanged=1;certificate=0"
        Else
            If re >= cap Then Exit For
            meshStage = "refine"
            If Not RPX_AngularFanMesh(cap, parent) Then Exit For
        End If
        ' Fresh q-specific original-program root; old signs never cross meshes.
        RSubdivisions = savedQ: RFsTolerance = savedTol
        RPX_DiagCandidate "lower_angular_fan_" & round
        RPX_TestHook "lower_angular_fan_" & round
        meshStage = "evaluate": RFsActive = True: RFsCalls = 0
        RFsSearchIncomplete = False: RFsLargestRootBracket = 0#: RFsSearchNote = ""
        lower = RPX_FsEndpoint("lower", seed.value, AdaptFsHintRatio)
        RLowerFs = lower: objective = XObjective: field = xx: RLowerField = field
        ' No upper field from another mesh may enter this candidate payload.
        Erase RUpperField
        ReDim eta(0 To re - 1)
        accepted = False
        If lower >= RBestLower.value And lower <= RBestUpper.value Then
            If SaveState(RBestLower, field, eta, lower, lower, RBestUpper.value, "lower_angular_fan_" & round, "lower", objective) Then
                RPX_CertFromAdapt "lower": accepted = True
            End If
        End If
        adaptFsCallsTotal = adaptFsCallsTotal + RFsCalls: RFsActive = False
        RPX_DiagEvent "lower_angular_fan_end", "round=" & round & ";accepted=" & IIf(accepted, "1", "0") & ";fs=" & lower & ";own_elements=" & re & ";own_q=" & RSubdivisions & ";root_incomplete=" & RFsSearchIncomplete & ";upper_mesh_unchanged=1"
        If Not accepted And (round > 0 Or rounds = 0) Then Exit For
    Next round
CandidateDone:
    RFsActive = False: RSubdivisions = savedQ: RFsTolerance = savedTol: DTarget = adaptFinalTarget
    Call RPX_RestoreWitnessState(baseline): CloneState RCurrent, baseline
    RSubdivisions = savedQ: RFsSearchIncomplete = oldIncomplete
    Call RestoreFinalBounds(mode)
    Exit Sub
CandidateFailed:
    number = err.number: source = err.source: message = err.description: errorLine = Erl
    On Error GoTo 0
    RPX_ErrorEvidence number, source, message, errorLine, meshStage
    If number = 18 Or XCancel Then err.Raise 18, source, "ANALYSIS_CANCELLED"
    If InStr(1, message, "HSD_STABLE_UNAVAILABLE_MEMORY", vbBinaryCompare) = 1 Then err.Raise number, source, message
    If Not RPX_Recoverable(number, message, meshStage, True) Then err.Raise number, source, message
    RPX_DiagEvent "lower_angular_fan_rejected", "round=" & round & ";old_witnesses_kept=1;message=" & message
    On Error GoTo RestoreFailed
    GoTo CandidateDone
RestoreFailed:
    number = err.number: source = err.source: message = err.description
    On Error GoTo 0
    err.Raise number, source, "ANGULAR_FAN_RESTORE_FAILED " & message
End Sub

' Reuse is confined to an already certified pair from this exact run. Each
' original field is audited again on its own mesh; no root is created here.
Private Function TerminalOwnMetadataReady(ByVal side As String, ByVal q As Long, ByRef state As RPX_MeshState) As Boolean
    Dim record As RPX_BoundCertificate, rootA As Double, rootB As Double, rootScale As Double
    If q <> 8 Or Not state.complete Or state.searchIncomplete Then ownReuseReason = "M1": Exit Function
    If RPX_CertRunId <= 0 Or state.certificateRunId <> RPX_CertRunId Then ownReuseReason = "M2": Exit Function
    If Not RPX_CertCurrent(side) Then ownReuseReason = "M3": Exit Function
    If state.q <> q Or state.inputEpoch <> RPX_InputEpoch Or state.policy <> RPX_AnalysisPolicy() Then ownReuseReason = "M4": Exit Function
    If state.nodes < 1 Or state.elements < 1 Or Len(state.originalModel) = 0 Then ownReuseReason = "M5": Exit Function
    If state.phase <> RPX_LoadPhase Or state.gravity <> RGravityMode Then ownReuseReason = "M6": Exit Function
    If Not RPX_SchurFinite(state.value) Or Not RPX_SchurFinite(state.factor) Or Not RPX_SchurFinite(state.stressScale) Then ownReuseReason = "M7": Exit Function
    If state.value <= 0# Or state.factor <> state.value Or state.stressScale <= 0# Then ownReuseReason = "M8": Exit Function
    If side = "lower" Then
        record = RPX_CertLower
        If Not state.lowerRootValid Then ownReuseReason = "M9": Exit Function
        rootA = state.lowerRootA: rootB = state.lowerRootB
        If state.value <> rootA Then ownReuseReason = "M10": Exit Function
    ElseIf side = "upper" Then
        record = RPX_CertUpper
        If Not state.upperRootValid Then ownReuseReason = "M11": Exit Function
        rootA = state.upperRootA: rootB = state.upperRootB
        If state.value <> rootB Then ownReuseReason = "M12": Exit Function
    Else
        Exit Function
    End If
    If record.side <> side Or record.mode <> "Fs" Or record.proofKind <> "ROOT_ENDPOINT" Then ownReuseReason = "M13": Exit Function
    If record.ownerKind <> "ADAPT_BEST_STATE" Or Not record.rootBracketValid Then ownReuseReason = "M14": Exit Function
    If record.q <> q Or record.nodes <> state.nodes Or record.elements <> state.elements Then ownReuseReason = "M15": Exit Function
    If Not RPX_SchurFinite(record.acceptedFs) Or Not RPX_SchurFinite(record.objective) Or Not RPX_SchurFinite(record.stressScale) Then ownReuseReason = "M16": Exit Function
    If Not RPX_SchurFinite(record.rootA) Or Not RPX_SchurFinite(record.rootB) Then ownReuseReason = "M17": Exit Function
    If record.acceptedFs <> state.value Or record.objective <> state.objective Or record.stressScale <> state.stressScale Then ownReuseReason = "M18": Exit Function
    If record.rootA <> rootA Or record.rootB <> rootB Then ownReuseReason = "M19": Exit Function
    If Not RPX_SchurFinite(rootA) Or Not RPX_SchurFinite(rootB) Or Not RPX_SchurFinite(state.objective) Then ownReuseReason = "M20": Exit Function
    If rootA <= 0# Or rootB < rootA Then ownReuseReason = "M21": Exit Function
    If Not RPX_SchurFinite(RFsTolerance) Or RFsTolerance <= 0# Then ownReuseReason = "M22": Exit Function
    If Not RPX_SchurFinite(state.searchWidth) Or state.searchWidth < 0# Then ownReuseReason = "M23": Exit Function
    rootScale = rootA / 2# + rootB / 2#: If rootScale < 1# Then rootScale = 1#
    If rootB - rootA > RFsTolerance * rootScale Or rootB - rootA > 0.0001 * rootScale Then ownReuseReason = "M24": Exit Function
    TerminalOwnMetadataReady = True
End Function
Private Function TerminalOwnFieldReady(ByVal side As String, ByRef state As RPX_MeshState) As Boolean
    Dim measured As Double, audit As Double, raw() As Double, i As Long
    Call RPX_RestoreWitnessState(state)
    ' RestoreWitnessState checks the exact geometry/material/load identity and
    ' restores the lower stress/load/reaction layout before the original audit.
    If LBound(state.field) <> 0 Then ownReuseReason = "F1": Exit Function
    If side = "upper" And UBound(state.field) <> 12 * state.elements + 3 * state.rigidN - 1 Then ownReuseReason = "F2": Exit Function
    For i = LBound(state.field) To UBound(state.field)
        If Not RPX_SchurFinite(state.field(i)) Then ownReuseReason = "F3": Exit Function
    Next i
    If Not RPX_FieldGate(side, state.field, state.value, 8, state.objective, measured, audit, raw) Then ownReuseReason = "F4": Exit Function
    If Not RPX_SchurFinite(measured) Or Not RPX_SchurFinite(audit) Then ownReuseReason = "F5": Exit Function
    If side = "lower" Then
        If measured < 1.000001 Then ownReuseReason = "F6": Exit Function
        For i = 1 To 4
            If Not RPX_SchurFinite(raw(i)) Or raw(i) > 0.0000001 Then ownReuseReason = "F7": Exit Function
        Next i
    Else
        ' G1R20 Fix7: same upper acceptance as RPX_UpperObjectiveAcceptable
        ' (UPPER_TARGET + 1e-8). For an all-phi=0 model the endpoint is RValue/target,
        ' so its re-audit lands on 0.999999 up to rounding (P10A: rejected by ~1e-11).
        If measured > 0.999999 + 0.00000001 Then ownReuseReason = "F8": Exit Function
    End If
    TerminalOwnFieldReady = True
End Function
Public Function RPX_TryTerminalPairReuse(ByVal mode As String) As Boolean
    Dim baseline As RPX_MeshState, q As Long, savedQ As Long, accepted As Boolean
    Dim entryPhase As String, witnessRestoreBack As Boolean
    Dim savedTol As Double, savedFactor As Double, savedIncomplete As Boolean, savedWidth As Double, savedNote As String
    Dim number As Long, source As String, message As String
    If mode <> "Fs" Or Not currentPairValid Or XCancel Or RPX_AuditRejected Then Exit Function
    If CStr(RPX_Setting("MODE", "")) <> "Fs" Then Exit Function
    If Not RCurrent.complete Or Not RRawAuditEnabled Then Exit Function
    q = CLng(RPX_Setting("ADAPT_CERT_QUADRATURE", 0))
    If q <> 8 Then Exit Function
    If Not RPX_TargetMet(mode, RBestLower.value, RBestUpper.value, currentPairValid) Then Exit Function
    If Not RCurrent.complete Or Len(RCurrent.originalModel) = 0 Then Exit Function
    entryPhase = RPX_LoadPhase
    ' Same-phase entry keeps the G1R19 behaviour (exact-model gate first,
    ' witness restore afterwards). Only a raw-phase entry takes the new path.
    witnessRestoreBack = (entryPhase = RCurrent.phase)
    If witnessRestoreBack And RCurrent.originalModel <> RPX_ExactModel() Then Exit Function
    If RCurrent.inputEpoch <> RPX_InputEpoch Or RCurrent.policy <> RPX_AnalysisPolicy() Then Exit Function
    CloneState baseline, RCurrent
    savedQ = RSubdivisions: savedTol = RFsTolerance: savedFactor = RStrengthFactor
    savedIncomplete = RFsSearchIncomplete: savedWidth = RFsLargestRootBracket: savedNote = RFsSearchNote
    On Error GoTo CandidateFailed
    ' Issue #11: AfterCycles restored the pair through RPX_ReadInputs (phase
    ' "raw"), while the saved fields were produced in "reference_total". Restore
    ' the pair's own saved load/material/model context (with its identity check)
    ' before comparing, instead of weakening the phase or exact-model checks.
    Call RPX_RestoreWitnessState(RCurrent)
    If RCurrent.originalModel <> RPX_ExactModel() Then GoTo RestoreBaseline
    If Not TerminalOwnMetadataReady("lower", q, RBestLower) Then GoTo RestoreBaseline
    If Not TerminalOwnMetadataReady("upper", q, RBestUpper) Then GoTo RestoreBaseline
    RPX_AssertInputs "terminal_pair_reuse"
    RPX_TestHook "terminal_pair_reuse"
    If Not TerminalOwnFieldReady("lower", RBestLower) Then GoTo RestoreBaseline
    If Not TerminalOwnFieldReady("upper", RBestUpper) Then GoTo RestoreBaseline
    RPX_TestHook "terminal_pair_reuse_commit"
    RPX_AssertInputs "terminal_pair_reuse_commit"
    accepted = True
RestoreBaseline:
    On Error GoTo RestoreFailed
    If accepted Or witnessRestoreBack Then
        ' Accepted pair continues in its saved reference_total context, as after
        ' Evaluate("final"), so output-side exact-model checks still hold.
        Call RPX_RestoreWitnessState(baseline): CloneState RCurrent, baseline
    Else
        ' Rejected: return to the same raw-input state AfterCycles had built.
        Call RPX_RestoreState(baseline): CloneState RCurrent, baseline
    End If
    RSubdivisions = savedQ: RFsTolerance = savedTol: RStrengthFactor = savedFactor
    RFsSearchIncomplete = savedIncomplete: RFsLargestRootBracket = savedWidth: RFsSearchNote = savedNote
    If accepted Then
        Call RestoreFinalBounds(mode)
        RPX_TryTerminalPairReuse = RPX_TargetMet(mode, RBestLower.value, RBestUpper.value, currentPairValid)
        If RPX_TryTerminalPairReuse Then RPX_DiagEvent "terminal_pair_reuse", "q=8;original_fields_reaudited=2;new_endpoint_calls=0;lower=" & RBestLower.value & ";upper=" & RBestUpper.value & ";run=" & RPX_CertRunId
    Else
        RPX_DiagEvent "terminal_pair_reuse_rejected", "old_witnesses_kept=1;existing_recheck_required=1;message=" & message
    End If
    Exit Function
CandidateFailed:
    number = err.number: source = err.source: message = err.description
    On Error GoTo 0
    If number = 18 Or XCancel Then err.Raise number, source, message
    If number = 7 Or number = 28 Then err.Raise number, source, message
    accepted = False
    GoTo RestoreBaseline
RestoreFailed:
    number = err.number: source = err.source: message = err.description
    On Error GoTo 0
    err.Raise number, source, "TERMINAL_PAIR_RESTORE_FAILED " & message
End Function

' Precision profiles recertify each side on its own immutable mesh.
' A changed quadrature never inherits a root bracket or a field certificate.
Private Function RecheckOwnQuadrature(ByVal side As String, ByVal q As Long) As Boolean
    Dim seed As RPX_MeshState, pending As RPX_MeshState
    Dim field() As Double, eta() As Double, value As Double, objective As Double
    Dim rootA As Double, rootB As Double, bracketScale As Double, i As Long
    Dim number As Long, source As String, message As String, errorLine As Long
    If side = "lower" Then CloneState seed, RBestLower Else CloneState seed, RBestUpper
    ' G1R20 Fix5 (#29): when this side is already an own-mesh q8 root endpoint
    ' certified in this run, re-audit its original field instead of solving the
    ' same mesh and q again (P06N Fix4: about 4370 s for identical Fs). Same
    ' checks as RPX_TryTerminalPairReuse, applied per side even if GAP is unmet.
    If OwnSideReusable(side, q, seed) Then
        RecheckOwnQuadrature = True: Exit Function
    End If
    Call RPX_RestoreWitnessState(seed)
    RSubdivisions = q
    RProgramKind = "": RPX_ResetPreparedState: RPX_LowerCacheReset: RPX_FsBracketReset
    meshStage = "evaluate"
    On Error GoTo CandidateFailed
    RPX_DiagCandidate "own_q" & q & "_" & side
    RPX_TestHook "own_quadrature_" & side
    RFsActive = True: RFsCalls = 0
    RFsSearchIncomplete = False: RFsLargestRootBracket = 0#: RFsSearchNote = ""
    value = RPX_FsEndpoint(side, seed.value, AdaptFsHintRatio): objective = XObjective
    If side = "lower" Then
        field = xx: RLowerField = field: RLowerFs = value: Erase RUpperField
    Else
        ReDim field(0 To 12 * re + 3 * RR - 1)
        For i = 0 To UBound(field): field(i) = xx(i): Next i
        RUpperField = field: RUpperFs = value
    End If
    adaptFsCallsTotal = adaptFsCallsTotal + RFsCalls: RFsActive = False
    If RFsSearchIncomplete Then GoTo Rejected
    If Not RPX_ActualRootBracket(side, rootA, rootB, value) Then GoTo Rejected
    bracketScale = (rootA + rootB) / 2#: If bracketScale < 1# Then bracketScale = 1#
    If rootB - rootA > RFsTolerance * bracketScale Then GoTo Rejected
    ReDim eta(0 To re - 1)
    If side = "lower" Then
        If value > RBestUpper.value Then GoTo Rejected
        If Not SaveState(pending, field, eta, value, value, RBestUpper.value, "own_q" & q & "_lower", side, objective) Then GoTo Rejected
        If Not pending.lowerRootValid Or pending.searchIncomplete Then GoTo Rejected
        CloneState RBestLower, pending
    Else
        If value < RBestLower.value Then GoTo Rejected
        If Not SaveState(pending, field, eta, value, RBestLower.value, value, "own_q" & q & "_upper", side, objective) Then GoTo Rejected
        If Not pending.upperRootValid Or pending.searchIncomplete Then GoTo Rejected
        CloneState RBestUpper, pending
    End If
    RPX_CertFromAdapt side
    RPX_DiagEvent "own_quadrature_accept", "side=" & side & ";q=" & q & ";fs=" & value & ";elements=" & re & ";root_a=" & rootA & ";root_b=" & rootB & ";original_field_gate=1;old_root_promoted=0"
    RecheckOwnQuadrature = True
    Exit Function
Rejected:
    RPX_DiagEvent "own_quadrature_rejected", "side=" & side & ";q=" & q & ";old_witness_and_q_kept=1;reason=root_or_field_not_certified"
    Exit Function
CandidateFailed:
    number = err.number: source = err.source: message = err.description: errorLine = Erl
    On Error GoTo 0
    RPX_ErrorEvidence number, source, message, errorLine, meshStage
    If number = 18 Or XCancel Then err.Raise 18, source, "ANALYSIS_CANCELLED"
    If InStr(1, message, "HSD_STABLE_UNAVAILABLE_MEMORY", vbBinaryCompare) = 1 Then err.Raise number, source, message
    If Not RPX_Recoverable(number, message, meshStage, True) Then err.Raise number, source, message
    RPX_DiagEvent "own_quadrature_rejected", "side=" & side & ";q=" & q & ";old_witness_and_q_kept=1;message=" & message
End Function

Private Function OwnSideReusable(ByVal side As String, ByVal q As Long, ByRef state As RPX_MeshState) As Boolean
    Dim number As Long, source As String, message As String
    If XCancel Or RPX_AuditRejected Or Not RRawAuditEnabled Then Exit Function
    If Len(state.originalModel) = 0 Then Exit Function
    On Error GoTo Failed
    ownReuseReason = ""
    If TerminalOwnMetadataReady(side, q, state) Then
        If TerminalOwnFieldReady(side, state) Then OwnSideReusable = True
    End If
    On Error GoTo 0
    If OwnSideReusable Then
        RPX_DiagEvent "own_quadrature_reuse", "side=" & side & ";q=" & q & ";fs=" & state.value & ";elements=" & state.elements & ";original_field_reaudited=1;new_endpoint_calls=0"
    Else
        RPX_DiagEvent "own_quadrature_reuse_skipped", "side=" & side & ";q=" & q & ";reason=not_certified_same_mesh_q;check=" & ownReuseReason & ";elements=" & state.elements & ";value=" & state.value & ";cert_kind=" & RPX_CertUpperOwnerKind(side)
    End If
    Exit Function
Failed:
    number = err.number: source = err.source: message = err.description
    On Error GoTo 0
    If number = 18 Or XCancel Then err.Raise 18, source, "ANALYSIS_CANCELLED"
    If number = 7 Or number = 28 Then err.Raise number, source, message
    OwnSideReusable = False
    RPX_DiagEvent "own_quadrature_reuse_skipped", "side=" & side & ";q=" & q & ";message=" & Replace(message, ";", ",")
End Function
Private Sub RecheckIndependentQuadrature(ByVal mode As String)
    Dim baseline As RPX_MeshState, q As Long, savedQ As Long, savedTol As Double
    Dim lowerOK As Boolean, upperOK As Boolean
    q = CLng(RPX_Setting("ADAPT_CERT_QUADRATURE", 0))
    If q <> 0 And q <> 2 And q <> 4 And q <> 8 Then err.Raise 5, "RPX_Adapt", "INVALID_ADAPT_CERT_QUADRATURE"
    If mode <> "Fs" Or q = 0 Then Exit Sub
    If Not currentPairValid Or Not RBestLower.complete Or Not RBestUpper.complete Then Exit Sub
    If RPX_TryTerminalPairReuse(mode) Then Exit Sub
    CloneState baseline, RCurrent
    savedQ = RSubdivisions: savedTol = RFsTolerance
    lowerOK = RecheckOwnQuadrature("lower", q)
    RFsTolerance = savedTol
    upperOK = RecheckOwnQuadrature("upper", q)
    RFsActive = False: RFsTolerance = savedTol
    Call RPX_RestoreWitnessState(baseline): CloneState RCurrent, baseline
    RSubdivisions = savedQ
    RFsSearchIncomplete = False
    Call RestoreFinalBounds(mode)
    If RBestLower.q <> q Or RBestUpper.q <> q Or Not lowerOK Or Not upperOK Then
        RFsSearchIncomplete = True
        RFsSearchNote = "INDEPENDENT_QUADRATURE_NOT_COMPLETE"
    End If
End Sub

' Field auditing uses lower stress/load/reaction indices, not just geometry.
' Keep that layout with the immutable witness so another mesh cannot supply it.
Private Function CopyWitnessReactions(ByVal reactions As Object) As Collection
    Dim result As New Collection, item As Variant
    If Not reactions Is Nothing Then
        For Each item In reactions
            result.Add Array(item(0), item(1), item(2), item(3), item(4))
        Next item
    End If
    Set CopyWitnessReactions = result
End Function
Private Sub CaptureWitnessLowerLayout(ByRef state As RPX_MeshState)
    Dim starts() As Long, e As Long, variables As Long
    ReDim starts(0 To re - 1)
    For e = 0 To re - 1
        starts(e) = -1
        If RRigidId(e) < 0 Then starts(e) = variables: variables = variables + 18
    Next e
    state.lowerStressStart = starts: state.lowerLoadVariable = variables
    Set state.lowerReactions = CopyWitnessReactions(RReactions)
End Sub

' An independently solved upper mesh; the retained lower witness never changes.
Private Function EarlyC0WedgeEnabled() As Boolean
    Dim value As Double
    value = CDbl(RPX_Setting("ADAPT_C0_WEDGE_AUTO", 0))
    If Not RPX_SchurFinite(value) Or (value <> 0# And value <> 1#) Then err.Raise 5, "RPX_Adapt", "INVALID_ADAPT_C0_WEDGE_AUTO"
    EarlyC0WedgeEnabled = (value = 1#)
End Function

Private Function ContinueUpperC0Wedge(ByVal mode As String) As Boolean
    Dim baseline As RPX_MeshState, pending As RPX_MeshState
    Dim depth As Double, savedQ As Long, savedTol As Double, oldIncomplete As Boolean
    Dim field() As Double, eta() As Double, upper As Double, objective As Double, hint As Double
    Dim a As Double, b As Double, accepted As Boolean, i As Long, number As Long, source As String, message As String
    depth = CDbl(RPX_Setting("ADAPT_C0_WEDGE_DEPTH", 0))
    If Not RPX_SchurFinite(depth) Then err.Raise 5, "RPX_Adapt", "INVALID_ADAPT_C0_WEDGE_DEPTH"
    If depth < 0# Then err.Raise 5, "RPX_Adapt", "INVALID_ADAPT_C0_WEDGE_DEPTH"
    If depth = 0# And EarlyC0WedgeEnabled() And DPCount = 6 Then
        ' A geometry-derived thin candidate, not an inferred bound.
        If DPXY(4, 1) > DPXY(3, 1) Then depth = 0.001 * (DPXY(4, 1) - DPXY(3, 1))
    End If
    If depth = 0# Or mode <> "Fs" Then Exit Function
    If Not currentPairValid Or Not RBestLower.complete Or Not RBestUpper.complete Then Exit Function
    If OptionalStageExpired("upper_c0_wedge") Then Exit Function
    CloneState baseline, RCurrent
    savedQ = RSubdivisions: savedTol = RFsTolerance: oldIncomplete = RFsSearchIncomplete
    On Error GoTo CandidateFailed
    meshStage = "band_regions": Call RPX_ReadInputs
    If Not RPX_C0SlopeWedgeMesh(depth) Then GoTo CandidateDone
    Call RPX_TotalReferenceLoads
    RSubdivisions = 8: RFsTolerance = CDbl(RPX_Setting("FS_TOL", 0.0001))
    If RFsTolerance <= 0# Or RFsTolerance > 0.0001 Then err.Raise 5, "RPX_Adapt", "C0_WEDGE_ROOT_TOLERANCE"
    hint = RPX_C0SlopeWedgeHint(depth)
    meshStage = "evaluate": RPX_DiagCandidate "upper_c0_wedge": RPX_TestHook "upper_c0_wedge"
    RFsActive = True: RFsCalls = 0
    RFsSearchIncomplete = False: RFsLargestRootBracket = 0#: RFsSearchNote = ""
    upper = RPX_FsEndpoint("upper", hint + 0.00005, 0.002 / hint)
    objective = XObjective: accepted = False
    ReDim field(0 To 12 * re + 3 * RR - 1)
    For i = 0 To UBound(field): field(i) = xx(i): Next i
    RUpperField = field: RUpperFs = upper: Erase RLowerField
    ReDim eta(0 To re - 1)
    If Not RFsSearchIncomplete And RPX_ActualRootBracket("upper", a, b, upper) Then
        If b - a <= RFsTolerance * WorksheetFunction.Max(1#, (a + b) / 2#) And upper >= RBestLower.value And upper <= RBestUpper.value Then
            If SaveState(pending, field, eta, upper, RBestLower.value, upper, "upper_c0_wedge", "upper", objective) Then
                If pending.upperRootValid And Not pending.searchIncomplete Then
                    CloneState RBestUpper, pending: RPX_CertFromAdapt "upper": accepted = True
                End If
            End If
        End If
    End If
    adaptFsCallsTotal = adaptFsCallsTotal + RFsCalls
    RPX_DiagEvent "upper_c0_wedge_end", "accepted=" & accepted & ";upper=" & upper & ";q=8;elements=" & re & ";root_a=" & a & ";root_b=" & b & ";lower_mesh_unchanged=1"
CandidateDone:
    RFsActive = False: RFsTolerance = savedTol: DTarget = adaptFinalTarget
    Call RPX_RestoreWitnessState(baseline): CloneState RCurrent, baseline
    RSubdivisions = savedQ: RFsSearchIncomplete = oldIncomplete
    Call RestoreFinalBounds(mode): ContinueUpperC0Wedge = accepted: Exit Function
CandidateFailed:
    number = err.number: source = err.source: message = err.description
    On Error GoTo 0
    RPX_ErrorEvidence number, source, message, 0, meshStage
    If number = 18 Or XCancel Then err.Raise 18, source, "ANALYSIS_CANCELLED"
    If Not RPX_Recoverable(number, message, meshStage, True) Then err.Raise number, source, message
    RPX_DiagEvent "upper_c0_wedge_rejected", "old_own_witnesses_kept=1;message=" & message
    On Error GoTo RestoreFailed: GoTo CandidateDone
RestoreFailed:
    number = err.number: source = err.source: message = err.description
    On Error GoTo 0: err.Raise number, source, "C0_WEDGE_RESTORE_FAILED " & message
End Function

Public Sub RPX_RunAdapt(ByVal mode As String)
    Dim failureNumber As Long, failureLine As Long
    Dim failureSource As String, failureText As String
    adaptRecoveryActive = False
    adaptRecoveryRunId = RPX_CertRunId
    adaptRecoveryTarget = DTarget: adaptRecoveryTolerance = RFsTolerance
    adaptRecoveryActive = True
    On Error GoTo AdaptFailed
    Call RunAdaptCore(mode)
    Exit Sub
AdaptFailed:
    failureNumber = err.number: failureSource = err.source: failureText = err.description: failureLine = Erl
    Resume AdaptRecovery
AdaptRecovery:
    On Error GoTo 0
    If RPX_AdaptPreserveConfirmed(mode, failureNumber, failureSource, failureText, failureLine) Then Exit Sub
    err.Raise failureNumber, failureSource, failureText
End Sub

' This path is entered only after an error. It neither solves nor accepts a bound.
Public Function RPX_AdaptPreserveConfirmed(ByVal mode As String, ByVal failureNumber As Long, ByVal failureSource As String, ByVal failureText As String, ByVal failureLine As Long) As Boolean
    Dim detail As String
    If failureNumber = 0 Or failureNumber = 18 Or XCancel Then Exit Function
    If mode <> "Fs" Or Not adaptRecoveryActive Then Exit Function
    If RPX_CertRunId <= 0 Or adaptRecoveryRunId <> RPX_CertRunId Then Exit Function
    If Not AdaptConfirmedOwnerReady("lower", RBestLower, RPX_CertLower) Then Exit Function
    If Not AdaptConfirmedOwnerReady("upper", RBestUpper, RPX_CertUpper) Then Exit Function
    If Not RPX_SchurFinite(RBestLower.value) Or Not RPX_SchurFinite(RBestUpper.value) Then Exit Function
    If RBestLower.value <= 0# Or RBestUpper.value < RBestLower.value Then Exit Function
    RPX_AssertInputs "late_fatal_preserve"
    ' Consume one recovery attempt before restoration can raise another error.
    adaptRecoveryActive = False
    Call Fix9ClearBudgets
    RPX_HsdBudgetFloor = 0#: RFsActive = False
    DTarget = adaptRecoveryTarget: RFsTolerance = adaptRecoveryTolerance
    Call RPX_RestoreWitnessState(RBestUpper)
    Call RestoreFinalBounds(mode)
    currentPairValid = True
    RFsSearchIncomplete = True
    RFsSearchNote = "ADAPT_FATAL_PRESERVED;OLD_WITNESSES_KEPT " & failureText
    adaptStatus = "fatal_preserved"
    detail = "number=" & failureNumber & ";source=" & failureSource & ";line=" & failureLine & ";message=" & Replace(failureText, ";", ",")
    detail = detail & ";lower=" & RBestLower.value & ";upper=" & RBestUpper.value & ";lower_elements=" & RBestLower.elements & ";upper_elements=" & RBestUpper.elements
    RPX_DiagEvent "adapt_fatal_preserved", detail
    RPX_AdaptPreserveConfirmed = True
End Function

Private Function AdaptConfirmedOwnerReady(ByVal side As String, ByRef state As RPX_MeshState, ByRef certificate As RPX_BoundCertificate) As Boolean
    If Not state.complete Or state.nodes < 1 Or state.elements < 1 Then Exit Function
    If state.certificateRunId <> RPX_CertRunId Or state.inputEpoch <> RPX_InputEpoch Then Exit Function
    If state.policy <> RPX_AnalysisPolicy() Or Len(state.originalModel) = 0 Then Exit Function
    If Not RPX_CertCurrent(side) Then Exit Function
    If certificate.side <> side Or certificate.mode <> "Fs" Or certificate.ownerKind <> "ADAPT_BEST_STATE" Then Exit Function
    If certificate.runId <> state.certificateRunId Or certificate.inputEpoch <> state.inputEpoch Or certificate.policy <> state.policy Then Exit Function
    If certificate.nodes <> state.nodes Or certificate.elements <> state.elements Or certificate.q <> state.q Then Exit Function
    If certificate.acceptedFs <> state.value Or certificate.objective <> state.objective Or certificate.stressScale <> state.stressScale Then Exit Function
    AdaptConfirmedOwnerReady = True
End Function

Private Sub RunAdaptCore(ByVal mode As String)
    Dim cycle As Long, cycles As Long, targetGap As Double, previousGap As Double, gain As Double
    Dim startTarget As Long, maxElements As Long
    Dim baseline As RPX_MeshState
    Dim savedErrNo As Long, savedErrLine As Long
    Dim savedErrText As String, savedErrSource As String
    Dim minArea As Double, areaIndex As Long
    Dim startCycle As Long, iPilot As Long
    Dim usedPilot As Boolean, pilotWanted As Boolean
    Dim pilotReason As String, route As String, fallbackReason As String
    Dim tmpVel() As Double
    Dim profile As RPX_ProblemProfile
    Dim robustRefine As Boolean, robustExhausted As Boolean, stagedRefine As Boolean, stagedCap As Long
    Dim fix8Ran As Boolean
    fix9ThinEarlyStartTaken = False
    stagedRefine = (UCase$(Trim$(CStr(RPX_Setting("ADAPT_MESH_POLICY", "LEGACY")))) = "LEGACY_THEN_REFINE" And RR = 0)
    robustRefine = RPX_RobustMeshPolicy() And Not stagedRefine
    If stagedRefine Then
        stagedCap = CLng(RPX_Setting("ADAPT_REFINE_CAP", 1000))
        If stagedCap < 16 Or stagedCap > 50000 Then err.Raise 5, "RPX_Adapt", "INVALID_ADAPT_REFINE_CAP"
    End If
    Set RAdaptHistory = New Collection
    adaptStartedAt = RPX_Clock()
    g20BalancedAttempted = False: g20FlowSource = ""
    RBestUpper.value = 1E+300: RBestLower.value = -1E+300
    RBestPair.pairGap = 1E+300
    RBestPair.elements = 0
    currentPairValid = False
    cycles = CLng(RPX_Setting("CYCLES", 3)): targetGap = CDbl(RPX_Setting("GAP", 0.02))
    If cycles < 0 Or cycles > 20 Or targetGap <= 0# Then err.Raise 5, , "Check CYCLES and GAP settings."
    adaptFinalTarget = DTarget
    adaptAccuracyCap = CLng(RPX_Setting("ADAPT_ACCURACY_CAP", adaptFinalTarget))
    adaptAccuracyRounds = CLng(RPX_Setting("ADAPT_ACCURACY_ROUNDS", 4))
    If adaptAccuracyCap = 0 Then adaptAccuracyCap = adaptFinalTarget
    If adaptAccuracyCap < adaptFinalTarget Or adaptAccuracyCap > 50000 Then err.Raise 5, "RPX_Adapt", "INVALID_ADAPT_ACCURACY_CAP"
    If adaptAccuracyRounds < 0 Or adaptAccuracyRounds > 20 Then err.Raise 5, "RPX_Adapt", "INVALID_ADAPT_ACCURACY_ROUNDS"
    startTarget = AdaptStartTarget(adaptFinalTarget)
    maxElements = CLng(adaptFinalTarget * AdaptMaxElementRatio)
    If maxElements < adaptFinalTarget Then maxElements = adaptFinalTarget
    adaptHintUpperValid = False
    adaptHintLowerValid = False
    adaptFsCallsTotal = 0
    adaptStatus = "completed"
    meshStage = ""
    route = "conventional"
    RAdaptRoute = route
    pilotFallbackUsed = False
    mesh0Ready = False
    snapReady = False
    RPX_FsBracketReset
    RPX_DiagEvent "adapt_mesh_policy", "requested=" & RPX_Setting("ADAPT_MESH_POLICY", "LEGACY") & ";effective_red_green=" & robustRefine & ";staged=" & stagedRefine & ";staged_cap=" & stagedCap & ";budget=" & DTarget
    RPX_DiagEvent "adapt_engine", "engine=20260926_p2s_f01_f07;ver=20260926_0945;start_ratio=" & AdaptStartRatio & ";start_target=" & startTarget & ";final_target=" & adaptFinalTarget & ";band_cap=" & AdaptBandElementCap
    DTarget = startTarget
    Call RPX_GenerateMesh
    RPX_DiagEvent "adapt_start", "start_elements=" & re & ";start_target=" & startTarget & ";final_budget=" & adaptFinalTarget & ";cycles=" & cycles & ";gap_target=" & targetGap
    RPX_BuildModelProfile profile
    RPX_DiagEvent "problem_profile", RPX_ProfileText(profile)
    RPX_LogFeatureRoutes profile
    usedPilot = False: startCycle = 1: fallbackReason = ""
    pilotWanted = RPX_CanUseFeature("UPPER_PILOT", profile, pilotReason)
    If RR > 0 Then pilotWanted = RPX_CanUseFeature("UPPER_PILOT_RIGID", profile, pilotReason)
    If cycles < 1 Then
        If pilotWanted Then pilotReason = "cycles"
        pilotWanted = False
    End If
    If pilotWanted Then
        On Error GoTo PilotFailed
        SaveGeometry mesh0: mesh0Ready = True
        RPX_DiagEvent "pilot_engine", "engine=20260926_p2s_f01_f07;ver=20260926_0945;enabled=1"
        RPX_DiagEvent "pilot_eligibility", "eligible=1;reason=;mode=" & mode & ";cycles=" & cycles
        Call EvaluateUpperPilot(mode)
        ReDim tmpVel(LBound(RUpperField) To UBound(RUpperField))
        For iPilot = LBound(RUpperField) To UBound(RUpperField): tmpVel(iPilot) = RUpperField(iPilot): Next iPilot
        RPX_DiagBegin dpMesh
        If Not TryBuildPilotBand(mesh0, tmpVel, RUpperFs) Then err.Raise 5, , "PILOT_WEIGHT_EMPTY"
        RPX_DiagEnd dpMesh
        If XCancel Then err.Raise 18, , "ANALYSIS_CANCELLED"
        adaptCoarseSearch = True
        Call Evaluate(mode, "adapt_1_band")
        If Not currentPairValid Then err.Raise 5, , "PILOT_PAIR_INCOMPLETE"
        If XCancel Then err.Raise 18, , "ANALYSIS_CANCELLED"
        usedPilot = True: startCycle = 2: route = "pilot": RAdaptRoute = route
        previousGap = RCurrent.pairGap
        RPX_DiagEvent "adapt_cycle", "cycle=1;route=pilot;elements=" & re & ";nodes=" & rn & ";lower=" & RCurrent.pairLower & ";upper=" & RCurrent.pairUpper & ";gap=" & RCurrent.pairGap & ";fs_calls=" & RFsCalls
    ElseIf pilotReason = "cycles" Then
        RPX_DiagEvent "pilot_eligibility", "eligible=1;reason=cycles;mode=" & mode & ";cycles=" & cycles
    Else
        RPX_DiagEvent "pilot_eligibility", "eligible=0;reason=" & pilotReason & ";mode=" & mode & ";cycles=" & cycles
    End If
    If usedPilot Then GoTo AdaptCycles
PilotRecover:
    On Error GoTo 0
    If pilotFallbackUsed Then route = "pilot_fallback" Else route = "conventional"
    RAdaptRoute = route
    adaptCoarseSearch = True
    Call Evaluate(mode, "coarse_0")
    If Not currentPairValid Then err.Raise 5, , "COARSE_PAIR_INCOMPLETE"
    If XCancel Then err.Raise 18, , "ANALYSIS_CANCELLED"
    startCycle = 1: usedPilot = False: previousGap = RCurrent.pairGap
    RPX_DiagEvent "adapt_cycle", "cycle=0;route=" & route & ";elements=" & re & ";nodes=" & rn & ";lower=" & RCurrent.pairLower & ";upper=" & RCurrent.pairUpper & ";gap=" & RCurrent.pairGap & ";fs_calls=" & RFsCalls
    ' Fix9 rc2: retain coarse bookkeeping and use the normal final-stage entry.
    If Fix9ThinEarlyStart(mode) Then GoTo AfterCycles
AdaptCycles:
    If route = "pilot" And Not pilotFallbackUsed Then
        On Error GoTo PilotFailed
    Else
        On Error GoTo 0
    End If
    ' Try an eligible small original-model upper mechanism before growing meshes.
    ' Its analytical threshold is only a hint. Both own q8 fields/roots must pass.
    If EarlyC0WedgeEnabled() Then
        If ContinueUpperC0Wedge(mode) Then
            Call RecheckIndependentQuadrature(mode)
            RPX_DiagEvent "early_c0_pair", "lower=" & RBestLower.value & ";upper=" & RBestUpper.value & ";gap=" & GapRatio(RBestLower.value, RBestUpper.value) & ";original_q8_roots_required=1"
            If RPX_TargetMet(mode, RBestLower.value, RBestUpper.value, currentPairValid) Then GoTo FinalPairReady
        End If
    End If
    Call G20TryGeometryCandidates(mode)
    If RPX_TryTerminalPairReuse(mode) Then GoTo FinalPairReady
    For cycle = startCycle To cycles
        If RCurrent.pairGap <= targetGap Then Exit For
        If adaptCoarseSearch Then
            If RCurrent.pairGap <= targetGap + AdaptCoarseTol Then
                RPX_DiagEvent "adapt_tighten", "cycle=" & cycle & ";gap=" & RCurrent.pairGap & ";target=" & targetGap & ";coarse_tol=" & AdaptCoarseTol
                adaptCoarseSearch = False
                Call Evaluate(mode, "tighten_" & cycle)
                adaptCoarseSearch = True
                If RCurrent.pairGap <= targetGap Then Exit For
            End If
        End If
        CloneState baseline, RCurrent
        RPX_DiagEvent "adapt_action", "action=" & IIf(robustRefine, "red_green_refine", "band") & ";cycle=" & cycle
        On Error GoTo CycleRemeshFailed
        RPX_DiagBegin dpMesh
        If robustRefine Then
            maxElements = CLng(baseline.elements * AdaptRobustRefineGrowth)
            If maxElements > adaptFinalTarget Then maxElements = adaptFinalTarget
            If Not TryShearRefine(baseline, maxElements) Then
                RPX_DiagEnd dpMesh: robustExhausted = True: Exit For
            End If
        Else
            Call RPX_BuildBandMesh(baseline)
        End If
        RPX_DiagEnd dpMesh
        If route = "pilot" And Not pilotFallbackUsed Then
            On Error GoTo PilotFailed
        Else
            On Error GoTo 0
        End If
        If robustRefine Then On Error GoTo CycleRemeshFailed
        If robustRefine Then meshStage = "evaluate"
        bandEvalFailedFlag = False
        If Not robustRefine And Not (route = "pilot" And Not pilotFallbackUsed) Then On Error GoTo BandEvalFailed
        Call Evaluate(mode, "adapt_" & cycle & IIf(robustRefine, "_refine", "_band"))
        If Not robustRefine And Not (route = "pilot" And Not pilotFallbackUsed) Then On Error GoTo 0
        If Not AcceptCandidate(baseline, RCurrent) Then
            RPX_DiagEvent "adapt_rollback", "cycle=" & cycle & ";action=band;new_gap=" & RCurrent.pairGap
            RPX_RestoreState baseline
            CloneState RCurrent, baseline
        End If
        If previousGap > 0# Then
            gain = (previousGap - RCurrent.pairGap) / previousGap
        Else
            gain = 0#
        End If
        previousGap = RCurrent.pairGap
        RPX_DiagEvent "adapt_cycle", "cycle=" & cycle & ";route=" & route & ";elements=" & re & ";nodes=" & rn & ";lower=" & RCurrent.pairLower & ";upper=" & RCurrent.pairUpper & ";gap=" & RCurrent.pairGap & ";gain=" & gain & ";fs_calls=" & RFsCalls
        Call G20TryGeometryCandidates(mode)
        If RPX_TryTerminalPairReuse(mode) Then GoTo FinalPairReady
        If robustRefine Then
            If RCurrent.elements <= baseline.elements Or gain <= 0# Then robustExhausted = True: Exit For
        Else
            If gain < AdaptMinGain Then Exit For
        End If
    Next cycle
BandEvalAfter:
    If bandEvalFailedFlag Then
        ' G1R15: band candidate failed (e.g. no lower bracket); keep the verified baseline pair and continue.
        bandEvalFailedFlag = False
        On Error GoTo 0
        RPX_RestoreState baseline
        CloneState RCurrent, baseline
        currentPairValid = True: adaptStatus = "band_failed_kept_baseline"
    End If
    If RCurrent.pairGap > targetGap And Not (robustRefine And robustExhausted) Then
        CloneState baseline, RCurrent
        If robustRefine Then
            maxElements = CLng(baseline.elements * AdaptRobustRefineGrowth)
            If maxElements > adaptFinalTarget Then maxElements = adaptFinalTarget
        Else
            maxElements = CLng(baseline.elements * AdaptBandRefineGrowth)
            If maxElements > AdaptBandElementCap Then maxElements = AdaptBandElementCap
            If maxElements < baseline.elements + 1 Then maxElements = baseline.elements + 1
        End If
        RPX_DiagEvent "adapt_action", "action=band_refine;max_elements=" & maxElements & ";fraction=" & AdaptBandRefineFraction
        On Error GoTo ShearRefineFailed
        RPX_DiagBegin dpMesh
        If TryShearRefine(baseline, maxElements, stagedRefine) Then
            RPX_DiagEnd dpMesh
            If route = "pilot" And Not pilotFallbackUsed Then
                On Error GoTo PilotFailed
            Else
                On Error GoTo 0
            End If
            If robustRefine Or stagedRefine Then On Error GoTo ShearRefineFailed
            If robustRefine Or stagedRefine Then meshStage = "evaluate"
            Call Evaluate(mode, "band_refine")
            If Not AcceptCandidate(baseline, RCurrent) Then
                RPX_DiagEvent "adapt_rollback", "action=band_refine;new_gap=" & RCurrent.pairGap
                RPX_RestoreState baseline
                CloneState RCurrent, baseline
            End If
        Else
            RPX_DiagEnd dpMesh
            On Error GoTo 0
        End If
    End If
    If stagedRefine And cycles > 0 And currentPairValid And RCurrent.pairGap > targetGap Then
        CloneState baseline, RCurrent
        maxElements = stagedCap: If maxElements > adaptFinalTarget Then maxElements = adaptFinalTarget
        If maxElements > baseline.elements Then
            RPX_DiagEvent "adapt_action", "action=legacy_then_red_green;baseline_elements=" & baseline.elements & ";max_elements=" & maxElements
            On Error GoTo StagedRefineFailed
            RPX_DiagBegin dpMesh
            If TryShearRefine(baseline, maxElements) Then
                RPX_DiagEnd dpMesh: meshStage = "evaluate"
                Call Evaluate(mode, "staged_refine")
                If Not AcceptCandidate(baseline, RCurrent) Then
                    RPX_DiagEvent "adapt_rollback", "action=staged_refine;candidate_gap=" & RCurrent.pairGap
                    RPX_RestoreState baseline: CloneState RCurrent, baseline
                End If
                RPX_DiagEvent "adapt_staged_end", "elements=" & re & ";lower=" & RCurrent.pairLower & ";upper=" & RCurrent.pairUpper & ";best_lower=" & RBestLower.value & ";best_upper=" & RBestUpper.value
            Else
                RPX_DiagEnd dpMesh
            End If
            On Error GoTo 0
        End If
    End If
AfterCycles:
    If route = "pilot" And Not pilotFallbackUsed Then
        On Error GoTo PilotFailed
    Else
        On Error GoTo 0
    End If
    If XCancel Then err.Raise 18, , "ANALYSIS_CANCELLED"
    If RBestPair.elements > 0 Then
        Call RPX_RestoreState(RBestPair)
        CloneState RCurrent, RBestPair
    End If
    adaptCoarseSearch = False
    If RCurrent.pairLower > 0# And currentPairValid Then
        adaptHintLower = RCurrent.pairLower
        adaptHintUpper = RCurrent.pairUpper
        adaptHintUpperValid = True
        adaptHintLowerValid = True
    End If
    If RPX_TryTerminalPairReuse(mode) Then GoTo FinalPairReady
    If FinalReuseAllowed(mode) Then
        RPX_DiagEvent "final_reuse", "elements=" & RCurrent.elements & ";lower=" & RCurrent.pairLower & ";upper=" & RCurrent.pairUpper & ";new_endpoint_calls=0"
        GoTo FinalCandidateDone
    End If
    ' Preserve audited side witnesses while the final candidate is being evaluated.
    On Error GoTo FinalCandidateFailed
    Call Evaluate(mode, "final")
    On Error GoTo 0
FinalCandidateDone:
    If XCancel Then err.Raise 18, , "ANALYSIS_CANCELLED"
    Call RestoreFinalBounds(mode)
    ' G1R20 Fix8: thin-layer side chains replace the pair-mesh accuracy rounds
    ' and the lower fan only when a thin-layer candidate reached endpoint evaluation.
    fix8Ran = Fix8ThinLayerChains(mode, targetGap)
    If Not fix8Ran Then Call ContinueAccuracy(mode, targetGap)
    Call RestoreFinalBounds(mode)
    If RPX_TryTerminalPairReuse(mode) Then GoTo FinalPairReady
    If Not fix8Ran Then Call ContinueLowerAngularFan(mode)
    Call ContinueUpperC0Wedge(mode)
    Call RecheckIndependentQuadrature(mode)
FinalPairReady:
    RPX_TestHook "final_accept"
    On Error GoTo 0
    If route = "pilot" And Not pilotFallbackUsed Then
        If (Not currentPairValid) Or GapRatio(RBestLower.value, RBestUpper.value) > targetGap Or RFsSearchIncomplete Then
            fallbackReason = "final"
            GoTo PilotGiveUp
        End If
    End If
    If GapRatio(RBestLower.value, RBestUpper.value) > targetGap Then
        RFsSearchIncomplete = True
        If Len(RFsSearchNote) = 0 Then RFsSearchNote = "ADAPT_GAP_TARGET_NOT_MET"
        RPX_DiagEvent "adapt_gap_unmet", "width=" & (RBestUpper.value - RBestLower.value)
    End If
    DTarget = adaptFinalTarget
    RPX_DiagEvent "adapt_outcome", "route=" & RAdaptRoute & ";fallback_used=" & IIf(pilotFallbackUsed, "1", "0") & ";complete_pair=" & IIf(currentPairValid, "1", "0") & ";target_met=" & IIf(RPX_TargetMet(mode, RBestLower.value, RBestUpper.value, currentPairValid), "1", "0") & ";calls=" & adaptFsCallsTotal & ";initial_lower_skipped=" & IIf(RAdaptRoute = "pilot", "1", "0")
    RPX_DiagEvent "adapt_end", "elements=" & re & ";lower=" & RBestLower.value & ";upper=" & RBestUpper.value & ";gap=" & GapRatio(RBestLower.value, RBestUpper.value) & ";fs_calls_total=" & adaptFsCallsTotal & ";adapt_status=" & adaptStatus & ";route=" & RAdaptRoute
    RPX_DiagFlush
    Exit Sub
FinalCandidateFailed:
    savedErrNo = err.number: savedErrText = err.description: savedErrSource = err.source
    On Error GoTo 0
    If (robustRefine Or stagedRefine) And InStr(1, savedErrText, "HSD_STABLE_UNAVAILABLE_MEMORY", vbBinaryCompare) = 1 Then err.Raise savedErrNo, savedErrSource, savedErrText
    If Not RPX_Recoverable(savedErrNo, savedErrText, "fs_trial", False) Then err.Raise savedErrNo, savedErrSource, savedErrText
    If RBestLower.elements <= 0 Or RBestUpper.elements <= 0 Then err.Raise savedErrNo, savedErrSource, savedErrText
    currentPairValid = True
    RFsSearchIncomplete = True: RFsSearchNote = "FINAL_NUMERICAL_UNKNOWN;OLD_WITNESSES_KEPT " & savedErrText
    RPX_DiagEvent "final_candidate_rejected", RFsSearchNote
    GoTo FinalCandidateDone
PilotGiveUp:
    If pilotFallbackUsed Then err.Raise 5, , "PILOT_FALLBACK_ALREADY_USED"
    pilotFallbackUsed = True
    RPX_DiagEvent "pilot_fallback", "reason=" & fallbackReason & ";fallback_count=1;gap=" & RCurrent.pairGap & ";number=" & savedErrNo & ";line=" & savedErrLine & ";message=" & Replace(savedErrText, ";", ",") & ";failure_stage=" & meshStage
    On Error GoTo PilotRestoreFailed
    If Not mesh0Ready Then err.Raise 5, , "snapshot_missing"
    meshStage = "m0_restore"
    RPX_TestHook "m0_restore"
    RPX_RestoreState mesh0
    VerifyPilotIdentity mesh0
    On Error GoTo 0
    RPX_FsBracketReset
    RProgramKind = ""
    adaptHintUpperValid = False
    adaptHintLowerValid = False
    currentPairValid = False
    RBestUpper.value = 1E+300: RBestLower.value = -1E+300
    RBestPair.pairGap = 1E+300: RBestPair.elements = 0
    GoTo PilotRecover
BandEvalFailed:
    savedErrNo = err.number: savedErrText = err.description: savedErrSource = err.source: savedErrLine = Erl
    If savedErrNo = 18 Or InStr(1, savedErrText, "HSD_STABLE_UNAVAILABLE_MEMORY", vbBinaryCompare) = 1 Or Not RPX_Recoverable(savedErrNo, savedErrText, "evaluate", True) Then
        On Error GoTo 0
        err.Raise savedErrNo, savedErrSource, savedErrText
    End If
    RPX_ErrorEvidence savedErrNo, savedErrSource, savedErrText, savedErrLine, "evaluate"
    RPX_DiagEvent "adapt_band_failed", "cycle=" & cycle & ";number=" & savedErrNo & ";message=" & savedErrText & ";fallback=keep_baseline"
    RPX_DiagFlush
    bandEvalFailedFlag = True
    Resume BandEvalAfter
ShearRefineFailed:
    savedErrNo = err.number: savedErrText = err.description: savedErrSource = err.source: savedErrLine = Erl
    On Error GoTo 0
    If (robustRefine Or stagedRefine) And InStr(1, savedErrText, "HSD_STABLE_UNAVAILABLE_MEMORY", vbBinaryCompare) = 1 Then err.Raise savedErrNo, savedErrSource, savedErrText
    RPX_ErrorEvidence savedErrNo, savedErrSource, savedErrText, savedErrLine, meshStage
    If Not RPX_Recoverable(savedErrNo, savedErrText, meshStage, currentPairValid Or mesh0Ready) Then err.Raise savedErrNo, savedErrSource, savedErrText
    On Error GoTo RefineRestoreFailed
    RPX_DiagEnd dpMesh
    RPX_DiagEvent "adapt_refine_failed", "stage=" & meshStage & ";number=" & savedErrNo & ";line=" & savedErrLine & ";source=" & savedErrSource & ";message=" & savedErrText & ";elements=" & re & ";nodes=" & rn
    RPX_DiagFlush
    If savedErrNo = 18 Then
        On Error GoTo 0
        err.Raise 18, savedErrSource, savedErrText
    End If
    If route = "pilot" And Not pilotFallbackUsed Then
        On Error GoTo 0
        If PilotRecoverable(savedErrNo, savedErrText) Then
            fallbackReason = "refine"
            GoTo PilotGiveUp
        End If
        err.Raise savedErrNo, savedErrSource, savedErrText
    End If
    meshStage = "rollback_restore"
    RPX_RestoreState baseline
    CloneState RCurrent, baseline
    adaptStatus = "refine_failed"
    GoTo AfterCycles
StagedRefineFailed:
    savedErrNo = err.number: savedErrText = err.description: savedErrSource = err.source: savedErrLine = Erl
    On Error GoTo 0
    RPX_ErrorEvidence savedErrNo, savedErrSource, savedErrText, savedErrLine, meshStage
    If InStr(1, savedErrText, "HSD_STABLE_UNAVAILABLE_MEMORY", vbBinaryCompare) = 1 Then err.Raise savedErrNo, savedErrSource, savedErrText
    If Not RPX_Recoverable(savedErrNo, savedErrText, meshStage, True) Then err.Raise savedErrNo, savedErrSource, savedErrText
    On Error GoTo RefineRestoreFailed
    RPX_DiagEnd dpMesh
    RPX_RestoreState baseline: CloneState RCurrent, baseline
    currentPairValid = True: adaptStatus = "staged_candidate_rejected"
    RPX_DiagEvent "adapt_staged_rejected", "old_witnesses_kept=1;number=" & savedErrNo & ";message=" & savedErrText
    GoTo AfterCycles
RefineRestoreFailed:
    savedErrNo = err.number: savedErrText = err.description: savedErrSource = err.source: savedErrLine = Erl
    RPX_DiagEvent "adapt_restore_failed", "stage=" & meshStage & ";number=" & savedErrNo & ";line=" & savedErrLine & ";message=" & savedErrText
    RPX_DiagFlush
    err.Raise savedErrNo, savedErrSource, "ADAPT_RESTORE_FAILED " & savedErrText
PilotFailed:
    savedErrNo = err.number: savedErrText = err.description: savedErrSource = err.source: savedErrLine = Erl
    On Error GoTo 0
    RPX_DiagEnd dpMesh
    If pilotFallbackUsed Then
        RPX_DiagEvent "pilot_fallback", "reason=stop;number=" & savedErrNo & ";message=" & Replace(savedErrText, ";", ",")
        RPX_DiagFlush
        err.Raise savedErrNo, , "PILOT_FALLBACK_ALREADY_USED " & savedErrText
    End If
    If Not PilotRecoverable(savedErrNo, savedErrText) Then
        RPX_DiagEvent "pilot_fallback", "reason=stop;number=" & savedErrNo & ";message=" & Replace(savedErrText, ";", ",")
        RPX_DiagFlush
        If savedErrNo = 18 Then err.Raise 18, , "ANALYSIS_CANCELLED"
        err.Raise savedErrNo, savedErrSource, savedErrText
    End If
    fallbackReason = "recover"
    GoTo PilotGiveUp
PilotRestoreFailed:
    savedErrNo = err.number: savedErrText = err.description: savedErrSource = err.source: savedErrLine = Erl
    On Error GoTo 0
    RPX_DiagEvent "adapt_restore_failed", "stage=pilot;number=" & savedErrNo & ";message=" & savedErrText
    RPX_DiagFlush
    err.Raise savedErrNo, savedErrSource, "ADAPT_RESTORE_FAILED " & savedErrText
CycleRemeshFailed:
    savedErrNo = err.number: savedErrText = err.description: savedErrSource = err.source: savedErrLine = Erl
    On Error GoTo 0
    If (robustRefine Or stagedRefine) And InStr(1, savedErrText, "HSD_STABLE_UNAVAILABLE_MEMORY", vbBinaryCompare) = 1 Then err.Raise savedErrNo, savedErrSource, savedErrText
    RPX_ErrorEvidence savedErrNo, savedErrSource, savedErrText, savedErrLine, meshStage
    If Not RPX_Recoverable(savedErrNo, savedErrText, meshStage, currentPairValid Or mesh0Ready) Then err.Raise savedErrNo, savedErrSource, savedErrText
    minArea = -1#
    On Error GoTo CycleRestoreFailed
    RPX_DiagEnd dpMesh
    RPX_DiagEvent "adapt_cycle_abort", "cycle=" & cycle & ";stage=" & meshStage & ";number=" & savedErrNo & ";line=" & savedErrLine & ";source=" & savedErrSource & ";message=" & savedErrText & ";elements=" & re & ";nodes=" & rn & ";min_area=" & minArea
    RPX_DiagFlush
    If savedErrNo = 18 Then
        On Error GoTo 0
        err.Raise 18, savedErrSource, savedErrText
    End If
    If route = "pilot" And Not pilotFallbackUsed Then
        On Error GoTo 0
        If PilotRecoverable(savedErrNo, savedErrText) Then
            fallbackReason = "cycle"
            GoTo PilotGiveUp
        End If
        err.Raise savedErrNo, savedErrSource, savedErrText
    End If
    meshStage = "rollback_restore"
    RPX_RestoreState baseline
    CloneState RCurrent, baseline
    adaptStatus = "aborted"
    GoTo AfterCycles
CycleRestoreFailed:
    savedErrNo = err.number: savedErrText = err.description: savedErrSource = err.source: savedErrLine = Erl
    RPX_DiagEvent "adapt_restore_failed", "stage=" & meshStage & ";number=" & savedErrNo & ";line=" & savedErrLine & ";message=" & savedErrText
    RPX_DiagFlush
    err.Raise savedErrNo, savedErrSource, "ADAPT_RESTORE_FAILED " & savedErrText
End Sub

' Select a nondegenerate mesh indicator; this does not change a material or bound.
' Pure c0 models have identically zero cohesive dissipation. Their velocity
' gradients and jumps still locate a mechanism and must not be discarded.
Public Sub RPX_AdaptiveRefineWeight(ByRef velocity() As Double, ByRef w() As Double, Optional ByVal kinematicFirst As Boolean = False)
    Dim e As Long, total As Double
    If kinematicFirst Then
        ComputeKinematicRefineWeight velocity, w
        Exit Sub
    End If
    ComputeShearWeight velocity, w
    For e = 0 To re - 1: total = total + w(e): Next e
    If total > 0# Then Exit Sub
    ComputeKinematicRefineWeight velocity, w
    total = 0#: For e = 0 To re - 1: total = total + w(e): Next e
    RPX_DiagEvent "adapt_weight", "kind=kinematic_zero_dissipation;total=" & total & ";original_materials_unchanged=1"
End Sub

Private Sub ComputeKinematicRefineWeight(ByRef velocity() As Double, ByRef w() As Double)
    Dim e As Long, f As Long, k As Long, j As Long, q As Long, id As Long
    Dim bary(0 To 2) As Double, g(0 To 5, 0 To 1) As Double
    Dim ex As Double, ey As Double, shear As Double
    Dim c As Double, phi As Double, a As Double, weight As Double, dissip As Double
    Dim weights(0 To 2) As Double, s As Double, vx As Double, vy As Double, nx As Double, ny As Double
    Dim jumpNormal As Double, jumpTangent As Double, value As Double
    Dim raw() As Double
    ReDim w(0 To re - 1)
    If Not VectorReady(velocity, 12 * re - 1) Then Exit Sub
    For e = 0 To re - 1
        If RRigidId(e) < 0 Then
            For q = 0 To 5
                If q < 3 Then
                    a = 0.445948490915965: weight = 0.223381589678011
                Else
                    a = 0.091576213509771: weight = 0.109951743655322
                End If
                For j = 0 To 2: bary(j) = a: Next j
                bary(q Mod 3) = 1# - 2# * a
                RPX_P2Grad e, bary, g
                ex = 0#: ey = 0#: shear = 0#
                For k = 0 To 5
                    ex = ex + velocity(12 * e + 2 * k) * g(k, 0)
                    ey = ey + velocity(12 * e + 2 * k + 1) * g(k, 1)
                    shear = shear + velocity(12 * e + 2 * k) * g(k, 1) + velocity(12 * e + 2 * k + 1) * g(k, 0)
                Next k
                dissip = Sqr(ex ^ 2 + ey ^ 2 + shear ^ 2 / 2#)
                w(e) = w(e) + RArea(e) * weight * dissip
            Next q
        End If
    Next e
    For id = 0 To RNEdge - 1
        With REdges(id)
            e = .element(0): f = .element(1)
            If f >= 0 Then
                If RRigidId(e) < 0 And RRigidId(f) < 0 Then
                    value = 0#: nx = .normal(0): ny = .normal(1)
                    For q = 0 To 2
                        s = 0.5 + (q - 1) * Sqr(0.15): weight = 5# / 18#: If q = 1 Then weight = 4# / 9#
                        RPX_TraceWeights s, weights
                        vx = 0#: vy = 0#
                        For k = 0 To 2
                            vx = vx + weights(k) * (velocity(12 * f + 2 * .loc(1, k)) - velocity(12 * e + 2 * .loc(0, k)))
                            vy = vy + weights(k) * (velocity(12 * f + 2 * .loc(1, k) + 1) - velocity(12 * e + 2 * .loc(0, k) + 1))
                        Next k
                        jumpNormal = vx * nx + vy * ny
                        jumpTangent = vx * .tangent(0) + vy * .tangent(1)
                        dissip = Sqr(jumpNormal ^ 2 + jumpTangent ^ 2)
                        If dissip < 0# Then dissip = 0#
                        value = value + .length * weight * dissip
                    Next q
                    w(e) = w(e) + value / 2#: w(f) = w(f) + value / 2#
                End If
            End If
        End With
    Next id
    raw = w
    For id = 0 To RNEdge - 1
        e = REdges(id).element(0): f = REdges(id).element(1)
        If f >= 0 Then
            If RMatId(e) <> RMatId(f) Then
                w(e) = w(e) + 0.25 * raw(f): w(f) = w(f) + 0.25 * raw(e)
            End If
        End If
    Next id
End Sub

' G1R20 optional geometry candidates. No field or sign crosses a new mesh.

Private Function G20MeshOption(ByVal key As String, ByVal upper As Long) As Long
    Dim v As Double
    v = CDbl(RPX_Setting(key, 0))
    If Not RPX_SchurFinite(v) Or v <> Fix(v) Or v < 0# Or v > upper Then err.Raise 5, "RPX_Adapt", "INVALID_" & key
    G20MeshOption = CLng(v)
End Function
Private Sub G20MaterialMoments(ByRef result() As Double)
    Dim e As Long, j As Long, k As Long, m As Long, x As Double, y As Double
    ReDim result(0 To RM - 1, 0 To 6)
    For e = 0 To re - 1
        m = RMatId(e): x = 0#: y = 0#
        If RRigidId(e) <> -1 Or RArea(e) <= 0# Then err.Raise 5, "RPX_Adapt", "G20_GEOMETRY_SCOPE"
        For k = 0 To 2: x = x + RXY(RTri(e, k), 0) / 3#: y = y + RXY(RTri(e, k), 1) / 3#: Next k
        result(m, 0) = result(m, 0) + RArea(e)
        result(m, 1) = result(m, 1) + RArea(e) * x: result(m, 2) = result(m, 2) + RArea(e) * y
        For j = 0 To 1
            result(m, 3 + j) = result(m, 3 + j) + RArea(e) * (RBody0(e, j) + RBody1(e, j))
        Next j
        result(m, 5) = result(m, 5) + RArea(e) * x * (RBody0(e, 1) + RBody1(e, 1))
        result(m, 6) = result(m, 6) + RArea(e) * y * (RBody0(e, 1) + RBody1(e, 1))
    Next e
End Sub
Private Sub G20CompareMoments(ByRef expected() As Double)
    Dim actual() As Double, m As Long, j As Long
    Call G20MaterialMoments(actual)
    For m = 0 To RM - 1: For j = 0 To 6
        If Abs(actual(m, j) - expected(m, j)) > 0.0000000001 + 0.000000000001 * Abs(expected(m, j)) Then err.Raise 5, "RPX_Adapt", "G20_MATERIAL_LOAD_MOMENT_CHANGED"
    Next j: Next m
End Sub
Private Function G20BuildBalancedLower(ByVal target As Long, ByVal cap As Long) As Boolean
    Dim score() As Double, e As Long, before As Long, savedTarget As Long, original() As Double
    RPX_AssertInputs "g20_balanced_geometry"
    If RR <> 0 Or RPX_WaterCouplingOn() Then err.Raise 5, "RPX_Adapt", "G20_BALANCED_UNSUPPORTED"
    If target < 16 Or target > 50000 Or cap < target Or cap > 50000 Then err.Raise 5, "RPX_Adapt", "G20_BALANCED_ARGUMENT"
    savedTarget = DTarget
    Call RPX_ReadInputs: DTarget = target
    Call RPX_GenerateMesh: Call RPX_TotalReferenceLoads
    Call G20MaterialMoments(original)
    before = re
    If 4# * before > cap Then
        DTarget = savedTarget
        RPX_DiagEvent "g20_balanced_skipped", "source_elements=" & before & ";candidate_elements=" & (4# * before) & ";cap=" & cap & ";old_witnesses_retained=1"
        Exit Function
    End If
    ReDim score(0 To re - 1): For e = 0 To re - 1: score(e) = 1#: Next e
    RPX_RefineMesh score, 1#, False, True
    Call RPX_TotalReferenceLoads
    DTarget = savedTarget
    If re > cap Then err.Raise 5, "RPX_Adapt", "G20_BALANCED_CAP"
    Call G20CompareMoments(original)
    Call RPX_FsBracketReset: RProgramKind = "": Call RPX_ResetPreparedState
    RPX_DiagEvent "g20_balanced_geometry", "target=" & target & ";source_elements=" & before & ";elements=" & re & ";uniform_red4=1;old_field_used=0;certificate=0"
    G20BuildBalancedLower = True
End Function
Private Function G20BuildFlowUpper(ByRef seed As RPX_MeshState, ByVal passes As Long) As Boolean
    Dim child() As Long, info As Object, original() As Double, perMaterial() As Double, seen() As Boolean
    Dim e As Long, j As Long, m As Long, key As Variant, got As Variant, want As Variant
    Dim sourceObjective As Double, sourceAudit As Double, sourceRaw() As Double
    If RPX_WaterCouplingOn() Or RR <> 0 Or seed.q <> 8 Or passes < 1 Or passes > 5 Then Exit Function
    If seed.originalModel <> RPX_ExactModel() Or seed.phase <> "reference_total" Then Exit Function
    If Not RPX_FieldGate("upper", seed.field, seed.factor, 8, seed.objective, sourceObjective, sourceAudit, sourceRaw) Then Exit Function
    If sourceObjective > 0.999999 Then Exit Function
    ' Body must be constant within each material; material-only flips cannot
    ' transport a varying body field. Water and rigid models remain on old routes.
    ReDim perMaterial(0 To RM - 1, 0 To 3): ReDim seen(0 To RM - 1)
    For e = 0 To re - 1
        m = RMatId(e)
        For j = 0 To 1
            If seen(m) Then
                If RBody0(e, j) <> perMaterial(m, 2 * j) Or RBody1(e, j) <> perMaterial(m, 2 * j + 1) Then Exit Function
            Else
                perMaterial(m, 2 * j) = RBody0(e, j): perMaterial(m, 2 * j + 1) = RBody1(e, j)
            End If
        Next j
        seen(m) = True
    Next e
    If Not RPX_FlowMeshCandidate(seed.xy, seed.triangles, seed.materialId, seed.field, seed.constraints, child, info, passes) Then Exit Function
    If CBool(info("certificate")) Or CBool(info("nested")) Or Not CBool(info("requires_new_original_q8_field_and_root")) Then err.Raise 5, "RPX_Adapt", "G20_FLOW_SCOPE"
    If CLng(info("flips")) = 0 Then Exit Function
    Call G20MaterialMoments(original)
    RTri = child: Set RBoundary = Nothing
    Call RPX_Topology: Call RPX_MapFaces: Call RPX_TotalReferenceLoads
    If rn <> seed.nodes Or re <> seed.elements Or RM <> seed.materialN Or RR <> 0 Then err.Raise 5, "RPX_Adapt", "G20_FLOW_DIMENSIONS_CHANGED"
    Call G20CompareMoments(original)
    If RBoundary.count <> seed.boundary.count Then err.Raise 5, "RPX_Adapt", "G20_FLOW_BOUNDARY_COUNT"
    For Each key In seed.boundary.keys
        If Not RBoundary.Exists(key) Then err.Raise 5, "RPX_Adapt", "G20_FLOW_BOUNDARY_EDGE"
        got = RBoundary(key): want = seed.boundary(key)
        For j = 0 To 5: If got(j) <> want(j) Then err.Raise 5, "RPX_Adapt", "G20_FLOW_BOUNDARY_LOAD"
        Next j
    Next key
    Call RPX_FsBracketReset: RProgramKind = "": Call RPX_ResetPreparedState
    RPX_DiagEvent "g20_flow_geometry", "elements=" & re & ";flips=" & info("flips") & ";passes=" & passes & ";new_original_field_root_required=1;certificate=0"
    G20BuildFlowUpper = True
End Function
Private Function G20TryCandidate(ByVal side As String, ByVal parameter As Long) As Boolean
    Dim baseline As RPX_MeshState, seed As RPX_MeshState, pending As RPX_MeshState
    Dim field() As Double, eta() As Double, factor As Double, objective As Double, a As Double, b As Double
    Dim savedQ As Long, savedTarget As Long, savedTol As Double, oldIncomplete As Boolean
    Dim accepted As Boolean, number As Long, source As String, message As String, label As String
    Dim savedAudit As RPX_AuditSnapshot, originalNumber As Long, originalSource As String, originalMessage As String
    Dim oldStatus As String, oldValue As Double, oldObjective As Double, oldRAudit As Double, oldRejected As Boolean, oldGate As Boolean, oldReason As String
    Dim oldRaw(1 To 4) As Double, oldDissipation As Double, oldReference As Double, oldFixed As Double, i As Long
    Dim oldX() As Double, hadX As Boolean, oldFactor As Double
    If Not currentPairValid Or Not RBestLower.complete Or Not RBestUpper.complete Or RR <> 0 Or RPX_WaterCouplingOn() Then Exit Function
    CloneState baseline, RCurrent
    savedQ = RSubdivisions: savedTol = RFsTolerance: savedTarget = DTarget: oldIncomplete = RFsSearchIncomplete
    oldStatus = XStatus: oldValue = RValue: oldObjective = XObjective: oldRAudit = RAudit: oldFactor = RStrengthFactor
    oldRejected = RPX_AuditRejected: oldGate = RPhysicalGateAccepted: oldReason = RPhysicalGateReason
    oldDissipation = RDissipation: oldReference = RReferenceWork: oldFixed = RFixedWork
    For i = 1 To 4: oldRaw(i) = RRawLowerAudit(i): Next i
    hadX = VectorReady(xx, 0): If hadX Then oldX = xx
    RPX_DiagAuditSave savedAudit
    On Error GoTo CandidateFailed
    RFsActive = False
    If side = "lower" Then
        label = "g20_balanced_lower": CloneState seed, RBestLower
        If Not G20BuildBalancedLower(parameter, adaptAccuracyCap) Then GoTo CandidateDone
    ElseIf side = "upper" Then
        label = "g20_flow_upper": CloneState seed, RBestUpper
        If seed.q <> 8 Or Not seed.upperRootValid Or seed.searchIncomplete Then GoTo CandidateDone
        Call RPX_RestoreWitnessState(seed)
        If Not G20BuildFlowUpper(seed, parameter) Then GoTo CandidateDone
    Else
        err.Raise 5, "RPX_Adapt", "G20_CANDIDATE_SIDE"
    End If
    RSubdivisions = 8: RFsTolerance = CDbl(RPX_Setting("FS_TOL", 0.0001))
    If RFsTolerance <= 0# Or RFsTolerance > 0.0001 Then err.Raise 5, "RPX_Adapt", "G20_ROOT_TOLERANCE"
    RPX_DiagCandidate label: RPX_TestHook label
    RFsActive = True: RFsCalls = 0: RFsSearchIncomplete = False: RFsLargestRootBracket = 0#: RFsSearchNote = ""
    factor = RPX_FsEndpoint(side, seed.value, AdaptFsHintRatio)
    objective = XObjective
    If side = "upper" Then
        ReDim field(0 To 12 * re + 3 * RR - 1)
        For i = 0 To UBound(field): field(i) = xx(i): Next i
    Else
        field = xx
    End If
    If side = "lower" Then RLowerFs = factor: RLowerField = field: Erase RUpperField Else RUpperFs = factor: RUpperField = field: Erase RLowerField
    ReDim eta(0 To re - 1)
    If Not RFsSearchIncomplete And RPX_ActualRootBracket(side, a, b, factor) Then
        If side = "lower" Then
            If factor >= RBestLower.value And factor <= RBestUpper.value Then
                If SaveState(pending, field, eta, factor, factor, RBestUpper.value, label, side, objective) Then
                    If pending.lowerRootValid And Not pending.searchIncomplete Then CloneState RBestLower, pending: RPX_CertFromAdapt side: accepted = True
                End If
            End If
        Else
            If factor >= RBestLower.value And factor <= RBestUpper.value Then
                If SaveState(pending, field, eta, factor, RBestLower.value, factor, label, side, objective) Then
                    If pending.upperRootValid And Not pending.searchIncomplete Then CloneState RBestUpper, pending: RPX_CertFromAdapt side: accepted = True
                End If
            End If
        End If
    End If
    adaptFsCallsTotal = adaptFsCallsTotal + RFsCalls
    RPX_DiagEvent "g20_candidate_end", "side=" & side & ";accepted=" & accepted & ";Fs=" & factor & ";q=8;elements=" & re & ";root_a=" & a & ";root_b=" & b & ";other_own_mesh_retained=1"
CandidateDone:
    On Error GoTo RestoreFailed
    RFsActive = False: RFsTolerance = savedTol: DTarget = savedTarget
    Call RPX_RestoreWitnessState(baseline): CloneState RCurrent, baseline
    RSubdivisions = savedQ: RFsSearchIncomplete = oldIncomplete
    Call RPX_DiagAuditRestore(savedAudit): Call RestoreFinalBounds("Fs")
    RStrengthFactor = oldFactor: XStatus = oldStatus: RValue = oldValue: XObjective = oldObjective: RAudit = oldRAudit
    RPX_AuditRejected = oldRejected: RPhysicalGateAccepted = oldGate: RPhysicalGateReason = oldReason
    RDissipation = oldDissipation: RReferenceWork = oldReference: RFixedWork = oldFixed
    For i = 1 To 4: RRawLowerAudit(i) = oldRaw(i): Next i
    If hadX Then xx = oldX Else Erase xx
    G20TryCandidate = accepted
    On Error GoTo 0
    If originalNumber <> 0 Then
        If originalNumber = 18 Then err.Raise 18, originalSource, "ANALYSIS_CANCELLED"
        err.Raise originalNumber, originalSource, originalMessage
    End If
    Exit Function
CandidateFailed:
    number = err.number: source = err.source: message = err.description
    On Error GoTo 0
    RPX_ErrorEvidence number, source, message, 0, "g20_candidate"
    If number = 18 Or XCancel Then
        originalNumber = 18: originalSource = source: originalMessage = message
    ElseIf Not RPX_Recoverable(number, message, "evaluate", True) Then
        originalNumber = number: originalSource = source: originalMessage = message
    End If
    RPX_DiagEvent "g20_candidate_rejected", "side=" & side & ";old_own_witnesses_retained=1;message=" & message
    GoTo CandidateDone
RestoreFailed:
    number = err.number: source = err.source: message = err.description
    On Error GoTo 0
    If originalNumber = 18 Or number = 18 Or XCancel Then err.Raise 18, source, "ANALYSIS_CANCELLED G20_RESTORE_FAILED " & message
    err.Raise number, source, "G20_CANDIDATE_RESTORE_FAILED original=" & originalNumber & ":" & originalMessage & ";restore=" & message
End Function
Private Sub G20TryGeometryCandidates(ByVal mode As String)
    Dim target As Long, passes As Long, sourceKey As String, accepted As Boolean
    If mode <> "Fs" Or Not currentPairValid Or RR <> 0 Or RPX_WaterCouplingOn() Then Exit Sub
    If GapRatio(RBestLower.value, RBestUpper.value) <= CDbl(RPX_Setting("GAP", 0.02)) Then Exit Sub
    target = G20MeshOption("ADAPT_BALANCED_LOWER_TARGET", 50000)
    passes = G20MeshOption("ADAPT_UPPER_FLOW_PASSES", 5)
    If target > 0 And Not g20BalancedAttempted Then
        g20BalancedAttempted = True
        accepted = G20TryCandidate("lower", target)
        If accepted Then Call RecheckIndependentQuadrature(mode)
    End If
    If GapRatio(RBestLower.value, RBestUpper.value) <= CDbl(RPX_Setting("GAP", 0.02)) Then Exit Sub
    If passes > 0 And RBestUpper.complete And RBestUpper.q = 8 Then
        sourceKey = RBestUpper.originalModel & "|" & RPX_ExactDouble(RBestUpper.value)
        If sourceKey <> g20FlowSource Then
            g20FlowSource = sourceKey
            accepted = G20TryCandidate("upper", passes)
            If accepted Then Call RecheckIndependentQuadrature(mode)
        End If
    End If
End Sub
' ---------------------------------------------------------------------------
' G1R20 Fix8 (candidate): thin-layer side chains (P06 mesh methods).
' Each side keeps its own lineage mesh; every candidate gets its own original
' q8 field, the unchanged full audit and its own production root (FS_TOL <=
' 1e-4, RPX_FsEndpoint + RPX_ActualRootBracket). Rejected candidates never
' replace a witness. Auto mode only triggers when a material is a very thin
' layer (thickness ratio < ADAPT_THIN_LAYER_RATIO), so other cases are unchanged.
' ---------------------------------------------------------------------------
Private Function Fix8Number(ByVal key As String, ByVal fallback As Double, ByVal low As Double, ByVal high As Double) As Double
    Dim v As Double
    v = CDbl(RPX_Setting(key, fallback))
    If Not RPX_SchurFinite(v) Or v < low Or v > high Then err.Raise 5, "RPX_Adapt", "INVALID_" & key
    Fix8Number = v
End Function
Private Function Fix8BodyPerMaterial() As Boolean
    Dim perMaterial() As Double, seen() As Boolean, e As Long, j As Long, m As Long
    ReDim perMaterial(0 To RM - 1, 0 To 3): ReDim seen(0 To RM - 1)
    For e = 0 To re - 1
        m = RMatId(e)
        For j = 0 To 1
            If seen(m) Then
                If RBody0(e, j) <> perMaterial(m, 2 * j) Or RBody1(e, j) <> perMaterial(m, 2 * j + 1) Then Exit Function
            Else
                perMaterial(m, 2 * j) = RBody0(e, j): perMaterial(m, 2 * j + 1) = RBody1(e, j)
            End If
        Next j
        seen(m) = True
    Next e
    Fix8BodyPerMaterial = True
End Function
Private Function Fix8ThinMode(ByVal mode As String, ByRef mask() As Boolean, Optional ByVal logDetection As Boolean = True) As Boolean
    Dim setting As Long, trigger As Double, midRatio As Double, thick() As Double, m As Long, minimum As Double, text As String, e As Long
    setting = CLng(Fix8Number("ADAPT_THIN_LAYER", 1, 0, 2))
    trigger = Fix8Number("ADAPT_THIN_LAYER_RATIO", 0.1, 0.000001, 1)
    midRatio = Fix8Number("ADAPT_THIN_LAYER_MID_RATIO", 0.25, 0.000001, 1)
    If setting = 0 Or mode <> "Fs" Or RR <> 0 Or RPX_WaterCouplingOn() Then Exit Function
    For e = 0 To re - 1: If RRigidId(e) <> -1 Then Exit Function
    Next e
    RPX_MaterialThickness thick
    ReDim mask(0 To RM - 1): minimum = 1E+300
    For m = 0 To RM - 1
        mask(m) = (thick(m) < midRatio)
        If thick(m) < minimum Then minimum = thick(m)
        text = text & IIf(m > 0, ",", "") & Format$(thick(m), "0.0000")
    Next m
    Fix8ThinMode = (setting = 2) Or (minimum < trigger)
    If Fix8ThinMode And Not Fix8BodyPerMaterial() Then
        If logDetection Then RPX_DiagEvent "thin_layer_detect", "active=0;reason=body_varies_within_material;thickness=" & text
        Fix8ThinMode = False: Exit Function
    End If
    If logDetection Then RPX_DiagEvent "thin_layer_detect", "active=" & IIf(Fix8ThinMode, "1", "0") & ";setting=" & setting & ";trigger=" & trigger & ";mid_ratio=" & midRatio & ";min_thickness=" & minimum & ";thickness=" & text
End Function
' The early probe records only a taken start; inactive cases emit nothing.
' Defer existing detection errors to the original final-stage call.
Private Function Fix9ThinEarlyStart(ByVal mode As String) As Boolean
    Dim mask() As Boolean, reason As String, earlyMode As Long
    Dim budgetOK As Boolean, budgetReason As String, budgetText As String
    If mode <> "Fs" Then Exit Function
    On Error GoTo DetectionUnavailable
    If Not Fix8ThinMode(mode, mask, False) Then Exit Function
    On Error GoTo 0
    earlyMode = Fix9Integer("ADAPT_THIN_EARLY_START", 1, 0, 2)
    If earlyMode = 0 Then Exit Function
    If earlyMode = 1 Then budgetOK = Fix9ThinEarlyBudget(budgetReason, budgetText)
    If Not currentPairValid Then
        reason = "pair_invalid"
    ElseIf Not RCurrent.complete Or Not RBestPair.complete Or Not RBestLower.complete Or Not RBestUpper.complete Then
        reason = "witness_incomplete"
    ElseIf RBestPair.elements <= 0 Or RBestLower.elements <= 0 Or RBestUpper.elements <= 0 Then
        reason = "witness_missing"
    ElseIf Not RPX_SchurFinite(RBestLower.value) Or Not RPX_SchurFinite(RBestUpper.value) Then
        reason = "bounds_nonfinite"
    ElseIf RBestLower.value <= 0# Or RBestLower.value > RBestUpper.value Then
        reason = "bounds_invalid"
    ElseIf RBestLower.source <> "coarse_0" Or RBestUpper.source <> "coarse_0" Then
        reason = "coarse_witness_missing"
    ElseIf RBestLower.value <> RCurrent.pairLower Or RBestUpper.value <> RCurrent.pairUpper Then
        reason = "coarse_witness_mismatch"
    ElseIf RBestPair.elements <> RCurrent.elements Or RBestPair.pairLower <> RCurrent.pairLower Or RBestPair.pairUpper <> RCurrent.pairUpper Then
        reason = "best_pair_mismatch"
    ElseIf RCurrent.q <> 8 Or RBestPair.q <> 8 Or RBestLower.q <> 8 Or RBestUpper.q <> 8 Then
        reason = "q_not_8"
    ElseIf Not RCurrent.lowerRootValid Or Not RCurrent.upperRootValid Or Not RBestLower.lowerRootValid Or Not RBestUpper.upperRootValid Then
        reason = "own_root_invalid"
    ElseIf RFsSearchIncomplete Or RCurrent.searchIncomplete Or RBestLower.searchIncomplete Or RBestUpper.searchIncomplete Then
        reason = "search_incomplete"
    ElseIf earlyMode = 1 And Not budgetOK Then
        reason = budgetReason
    Else
        reason = "coarse_q8_pair"
        If earlyMode = 1 Then reason = budgetReason
        Fix9ThinEarlyStart = True
        fix9ThinEarlyStartTaken = True
    End If
    RPX_DiagEvent "thin_layer_early_start", "active=" & IIf(Fix9ThinEarlyStart, "1", "0") & ";reason=" & reason & ";elements=" & re & ";lower=" & RBestLower.value & ";upper=" & RBestUpper.value & ";early_mode=" & earlyMode & budgetText
    Exit Function
DetectionUnavailable:
    err.Clear
End Function
' rc3b admission policy only: theta is indicator mass, not predicted mesh growth.
' Evaluate the upper budget used after early admission; no state writes.
Private Function Fix9ThinEarlyBudget(ByRef reason As String, ByRef text As String) As Boolean
    Dim uRounds As Long, lRounds As Long, lStart As Long, uCap As Long, lCap As Long
    Dim uTheta As Double, lTheta As Double, coarse As Long, lowerEnd As Long
    Dim minUpper As Long, minLower As Long, minRatio As Double
    Dim minUpperTheta As Double, minLowerTheta As Double, minSeconds As Double
    Dim upperRatio As Double, lowerRatio As Double, limit As Double, elapsed As Double, remaining As Double
    Dim timeActive As Boolean
    uRounds = CLng(Fix8Number("ADAPT_THIN_UPPER_ROUNDS", 10, 0, 20))
    lRounds = CLng(Fix8Number("ADAPT_THIN_LOWER_ROUNDS", 3, 0, 10))
    lStart = CLng(Fix8Number("ADAPT_THIN_LOWER_START", 2, 1, 20))
    uTheta = Fix8Number("ADAPT_THIN_EARLY_UPPER_THETA", 0.95, 0.05, 1)
    lTheta = Fix8Number("ADAPT_THIN_LOWER_THETA", 0.9, 0.05, 1)
    uCap = CLng(Fix8Number("ADAPT_THIN_EARLY_UPPER_CAP", 7000, 16, 50000))
    lCap = CLng(Fix8Number("ADAPT_THIN_LOWER_CAP", 6000, 16, 50000))
    minUpper = Fix9Integer("ADAPT_THIN_EARLY_MIN_UPPER_ROUNDS", 10, 0, 20)
    minLower = Fix9Integer("ADAPT_THIN_EARLY_MIN_LOWER_ROUNDS", 3, 0, 10)
    minRatio = Fix8Number("ADAPT_THIN_EARLY_MIN_CAP_RATIO", 2, 1, 50000)
    minUpperTheta = Fix8Number("ADAPT_THIN_EARLY_MIN_UPPER_THETA", 0.8, 0.05, 1)
    minLowerTheta = Fix8Number("ADAPT_THIN_EARLY_MIN_LOWER_THETA", 0.9, 0.05, 1)
    minSeconds = Fix8Number("ADAPT_THIN_EARLY_MIN_REMAINING_SECONDS", 3600, 0, 1E+30)
    coarse = re: lowerEnd = lStart + lRounds - 1
    If coarse > 0 Then
        upperRatio = CDbl(uCap) / CDbl(coarse)
        lowerRatio = CDbl(lCap) / CDbl(coarse)
    End If
    ' Same deadline validation and clock as OptionalStageExpired; -1 = inactive.
    limit = CDbl(RPX_Setting("ADAPT_OPTIONAL_DEADLINE_SECONDS", 0))
    If Not RPX_SchurFinite(limit) Or limit < 0# Then err.Raise 5, "RPX_Adapt", "INVALID_ADAPT_OPTIONAL_DEADLINE_SECONDS"
    remaining = -1#: timeActive = (limit > 0# And adaptStartedAt > 0#)
    If timeActive Then
        elapsed = RPX_Clock() - adaptStartedAt
        remaining = limit - elapsed
        If remaining < 0# Then remaining = 0#
    End If
    text = ";upper_rounds=" & uRounds & ";lower_rounds=" & lRounds & ";lower_start=" & lStart & ";lower_end=" & lowerEnd
    text = text & ";upper_cap=" & uCap & ";lower_cap=" & lCap & ";upper_theta=" & uTheta & ";lower_theta=" & lTheta
    text = text & ";early_upper_theta=" & uTheta & ";early_upper_cap=" & uCap
    text = text & ";coarse_elements=" & coarse & ";upper_cap_ratio=" & upperRatio & ";lower_cap_ratio=" & lowerRatio
    text = text & ";min_upper_rounds=" & minUpper & ";min_lower_rounds=" & minLower & ";min_cap_ratio=" & minRatio
    text = text & ";min_upper_theta=" & minUpperTheta & ";min_lower_theta=" & minLowerTheta & ";min_remaining_seconds=" & minSeconds
    text = text & ";deadline_seconds=" & limit & ";time_budget_active=" & Abs(CLng(timeActive)) & ";adapt_started_at=" & adaptStartedAt
    text = text & ";elapsed_seconds=" & elapsed & ";remaining_seconds=" & remaining
    If coarse <= 0 Then
        reason = "coarse_elements_invalid"
    ElseIf uRounds < minUpper Then
        reason = "upper_rounds_short"
    ElseIf lRounds < minLower Then
        reason = "lower_rounds_short"
    ElseIf lowerEnd > uRounds Then
        reason = "lower_start_late"
    ElseIf uTheta < minUpperTheta Then
        reason = "upper_theta_low"
    ElseIf lTheta < minLowerTheta Then
        reason = "lower_theta_low"
    ElseIf upperRatio < minRatio Then
        reason = "upper_cap_small"
    ElseIf lowerRatio < minRatio Then
        reason = "lower_cap_small"
    ElseIf timeActive And remaining < minSeconds Then
        reason = "remaining_time_short"
    Else
        reason = "chain_budget_sufficient"
        Fix9ThinEarlyBudget = True
    End If
End Function
Private Sub Fix8Capacity()
    Dim xy() As Double, tri() As Long, i As Long
    ReDim xy(0 To MeshNodeCapacity - 1, 0 To 1): ReDim tri(0 To MeshTriangleCapacity - 1, 0 To 2)
    For i = 0 To rn - 1: xy(i, 0) = RXY(i, 0): xy(i, 1) = RXY(i, 1): Next i
    For i = 0 To re - 1: tri(i, 0) = RTri(i, 0): tri(i, 1) = RTri(i, 1): tri(i, 2) = RTri(i, 2): Next i
    RXY = xy: RTri = tri
End Sub
Private Sub Fix8Commit()
    Call RPX_TotalReferenceLoads
    RProgramKind = "": Call RPX_ResetPreparedState: Call RPX_LowerCacheReset: Call RPX_FsBracketReset
End Sub
' Fix9 settings are read only after Fix8ThinMode has admitted the chain.
Private Function Fix9Integer(ByVal key As String, ByVal fallback As Double, ByVal low As Double, ByVal high As Double) As Long
    Dim v As Double
    v = Fix8Number(key, fallback, low, high)
    If v <> Fix(v) Then err.Raise 5, "RPX_Adapt", "INVALID_" & key
    Fix9Integer = CLng(v)
End Function
Private Sub Fix9ReadOptions(ByRef opt As Fix9ThinOptions, ByRef text As String)
    opt.midSeparate = (Fix9Integer("ADAPT_THIN_MID_SEPARATE", 1, 0, 1) = 1)
    opt.lowerOwnUpper = (Fix9Integer("ADAPT_THIN_LOWER_OWN_UPPER", 1, 0, 1) = 1)
    opt.lowerOwnFactor = Fix8Number("ADAPT_THIN_LOWER_OWN_UPPER_FACTOR", 1.08, 1, 2)
    opt.flipFresh = (Fix9Integer("ADAPT_THIN_FLIP_FRESH", 1, 0, 1) = 1)
    opt.lowerRetries = Fix9Integer("ADAPT_THIN_LOWER_RETRIES", 1, 0, 3)
    opt.retryThetaFactor = Fix8Number("ADAPT_THIN_RETRY_THETA_FACTOR", 0.5, 0.000001, 1)
    opt.hsdBase = Fix8Number("ADAPT_THIN_HSD_BASE", 20, 0, 3600)
    opt.hsdPerCone = Fix8Number("ADAPT_THIN_HSD_PER_CONE", 0.0037, 0, 1)
    opt.hsdFactor = Fix8Number("ADAPT_THIN_HSD_FACTOR", 1.5, 0, 100)
    opt.hsdCap = Fix8Number("ADAPT_THIN_HSD_CAP", 3600, 1, 3600)
    opt.hsdTimeFraction = Fix8Number("ADAPT_THIN_HSD_TIME_FRACTION", 0.15, 0, 1)
    opt.hsdTimeCap = Fix8Number("ADAPT_THIN_HSD_TIME_CAP", 1200, 0, 3600)
    opt.nominalSeconds = Fix8Number("ADAPT_THIN_NOMINAL_SECONDS", 18171, 0, 1E+30)
    opt.lowerLargeElements = Fix9Integer("ADAPT_THIN_LOWER_LARGE_ELEMENTS", 1500, 1, 50000)
    opt.lowerLargeThetaFactor = Fix8Number("ADAPT_THIN_LOWER_LARGE_THETA_FACTOR", 0.5, 0.000001, 1)
    opt.extraRetryMinSeconds = Fix8Number("ADAPT_THIN_EXTRA_RETRY_MIN_SECONDS", 3600, 0, 1E+30)
    opt.dllFactor = Fix8Number("ADAPT_THIN_DLL_FACTOR", 3, 0, 100)
    opt.dllRefSeconds = Fix8Number("ADAPT_THIN_DLL_REF_SECONDS", 3.45, 0.000001, 600)
    opt.dllEquTimeFactor = Fix8Number("ADAPT_THIN_DLL_EQU_TIME_FACTOR", 1.15, 1, 100)
    opt.dllRefCones = Fix9Integer("ADAPT_THIN_DLL_REF_CONES", 22741, 1, 2147483000#)
    opt.dllExponent = Fix8Number("ADAPT_THIN_DLL_EXPONENT", 1.2, 0, 4)
    opt.dllCap = Fix8Number("ADAPT_THIN_DLL_CAP", 600, 0.000001, 600)
    opt.upperRoundMaxElements = Fix9Integer("ADAPT_THIN_UPPER_ROUND_MAX_ELEMENTS", 3000, 16, 50000)
    opt.upperMaxGrowth = Fix8Number("ADAPT_THIN_UPPER_MAX_GROWTH", 1.6, 1.000001, 4)
    opt.upperDllRetry = (Fix9Integer("ADAPT_THIN_UPPER_DLL_RETRY", 1, 0, 1) = 1)
    opt.upperDllMinElements = Fix9Integer("ADAPT_THIN_UPPER_DLL_MIN_ELEMENTS", 1800, 1, 50000)
    opt.upperDllRetrySeconds = Fix8Number("ADAPT_THIN_UPPER_DLL_RETRY_SECONDS", 60, 0.000001, 600)
    opt.upperDllRetryIterations = Fix9Integer("ADAPT_THIN_UPPER_DLL_RETRY_ITERATIONS", 700, 350, 2000)
    opt.upperDllRetryEquilibrate = Fix9Integer("ADAPT_THIN_UPPER_DLL_RETRY_EQUILIBRATE", 1, 0, 1)
    opt.upperHsdMinElements = Fix9Integer("ADAPT_THIN_UPPER_HSD_MIN_ELEMENTS", 1800, 1, 50000)
    opt.upperHsdFloorSeconds = Fix8Number("ADAPT_THIN_UPPER_HSD_FLOOR_SECONDS", 1800, 0, 3600)
    opt.lowerHsdMinElements = Fix9Integer("ADAPT_THIN_LOWER_HSD_MIN_ELEMENTS", 1400, 1, 50000)
    opt.lowerHsdFloorSeconds = Fix8Number("ADAPT_THIN_LOWER_HSD_FLOOR_SECONDS", 1200, 0, 3600)
    opt.upperHsdReserveSeconds = Fix8Number("ADAPT_THIN_UPPER_HSD_RESERVE_SECONDS", 600, 0, 1E+30)
    opt.upperHsdIterations = Fix9Integer("ADAPT_THIN_UPPER_HSD_ITERATIONS", 600, 150, 1000)
    opt.lowerExtraProtect = (Fix9Integer("ADAPT_THIN_LOWER_EXTRA_PROTECT", 1, 0, 1) = 1)
    opt.lowerExtraRounds = Fix9Integer("ADAPT_THIN_LOWER_EXTRA_ROUNDS", 3, 0, 10)
    opt.lowerExtraReserve = Fix8Number("ADAPT_THIN_LOWER_EXTRA_RESERVE_SECONDS", 600, 0, 1E+30)
    opt.lowerExtraTargetElements = Fix9Integer("ADAPT_THIN_LOWER_EXTRA_TARGET_ELEMENTS", 2000, 16, 50000)
    opt.lowerExtraMaxGrowth = Fix8Number("ADAPT_THIN_LOWER_EXTRA_MAX_GROWTH", 1.4, 1.000001, 4)
    opt.lowerExtraUpperRound = Fix9Integer("ADAPT_THIN_LOWER_EXTRA_UPPER_ROUND", 7, 0, 20)
    text = ";mid_separate=" & Abs(CLng(opt.midSeparate)) & ";lower_own_upper=" & Abs(CLng(opt.lowerOwnUpper)) & ";lower_own_upper_factor=" & opt.lowerOwnFactor
    text = text & ";flip_fresh=" & Abs(CLng(opt.flipFresh)) & ";lower_retries=" & opt.lowerRetries & ";retry_theta_factor=" & opt.retryThetaFactor
    text = text & ";hsd_base=" & opt.hsdBase & ";hsd_per_cone=" & opt.hsdPerCone & ";hsd_factor=" & opt.hsdFactor & ";hsd_cap=" & opt.hsdCap
    text = text & ";hsd_time_fraction=" & opt.hsdTimeFraction & ";hsd_time_cap=" & opt.hsdTimeCap
    text = text & ";nominal_seconds=" & opt.nominalSeconds
    text = text & ";lower_large_elements=" & opt.lowerLargeElements & ";lower_large_theta_factor=" & opt.lowerLargeThetaFactor & ";extra_retry_min_seconds=" & opt.extraRetryMinSeconds
    text = text & ";dll_factor=" & opt.dllFactor & ";dll_ref_seconds=" & opt.dllRefSeconds & ";dll_ref_cones=" & opt.dllRefCones & ";dll_exponent=" & opt.dllExponent & ";dll_cap=" & opt.dllCap
    text = text & ";dll_equ_time_factor=" & opt.dllEquTimeFactor
    text = text & ";upper_round_max_elements=" & opt.upperRoundMaxElements & ";upper_max_growth=" & opt.upperMaxGrowth
    text = text & ";upper_dll_retry=" & Abs(CLng(opt.upperDllRetry)) & ";upper_dll_min_elements=" & opt.upperDllMinElements & ";upper_dll_retry_seconds=" & opt.upperDllRetrySeconds
    text = text & ";upper_dll_retry_iterations=" & opt.upperDllRetryIterations & ";upper_dll_retry_equilibrate=" & opt.upperDllRetryEquilibrate
    text = text & ";upper_hsd_min_elements=" & opt.upperHsdMinElements & ";upper_hsd_floor_seconds=" & opt.upperHsdFloorSeconds
    text = text & ";lower_hsd_min_elements=" & opt.lowerHsdMinElements & ";lower_hsd_floor_seconds=" & opt.lowerHsdFloorSeconds
    text = text & ";upper_hsd_reserve_seconds=" & opt.upperHsdReserveSeconds & ";upper_hsd_iterations=" & opt.upperHsdIterations
    text = text & ";lower_extra_rounds=" & opt.lowerExtraRounds & ";lower_extra_reserve_seconds=" & opt.lowerExtraReserve & ";lower_extra_target_elements=" & opt.lowerExtraTargetElements
    text = text & ";lower_extra_max_growth=" & opt.lowerExtraMaxGrowth & ";lower_extra_upper_round=" & opt.lowerExtraUpperRound
    text = text & ";lower_extra_protect=" & Abs(CLng(opt.lowerExtraProtect))
End Sub
' Thin funding uses the case clock when the runner supplies its limit. An
' optional-stage deadline remains an independent stage veto, not a case limit.
Public Function RPX_ThinRemainingSeconds(Optional ByRef remainingSource As String = "none") As Double
    RPX_ThinRemainingSeconds = -1#
    remainingSource = "none"
    If Not RPX_ThinChainBudgetActive Then Exit Function
    RPX_ThinRemainingSeconds = Fix9DeadlineRemainingSeconds(fix9ThinNominalSeconds, remainingSource)
End Function
' Every thin solver shares the case wall and its side's completion reserve.
' A measured round already includes equilibration; this clamp never rescales it.
Public Function RPX_ThinSolveSeconds(ByVal seconds As Double) As Double
    Dim remainingSeconds As Double, remainingSource As String, reserveSeconds As Double
    RPX_ThinSolveSeconds = seconds
    If Not RPX_ThinChainBudgetActive Then Exit Function
    remainingSeconds = RPX_ThinRemainingSeconds(remainingSource)
    If remainingSeconds < 0# Then
        If RPX_ThinLowerExtraActive Then RPX_ThinSolveSeconds = 0#
        Exit Function
    End If
    If RPX_ThinChainUpperActive Then
        reserveSeconds = RPX_ThinUpperHsdReserveSeconds
    Else
        reserveSeconds = RPX_ThinLowerHsdReserveSeconds
    End If
    ' An auxiliary upper LP inside an extra lower round keeps its lower reserve.
    If RPX_ThinLowerExtraActive And reserveSeconds < RPX_ThinLowerHsdReserveSeconds Then reserveSeconds = RPX_ThinLowerHsdReserveSeconds
    remainingSeconds = remainingSeconds - reserveSeconds
    If remainingSeconds < 0# Then remainingSeconds = 0#
    If seconds > remainingSeconds Then RPX_ThinSolveSeconds = remainingSeconds
End Function
' Also used between steps by the thin-only retry driver, after budget cleanup.
Private Function Fix9DeadlineRemainingSeconds(ByVal nominalSeconds As Double, ByRef remainingSource As String) As Double
    Dim limit As Double, remaining As Double, caseLimitSeconds As Double, caseElapsedSeconds As Double
    Fix9DeadlineRemainingSeconds = -1#
    remainingSource = "none"
    caseLimitSeconds = Fix8Number("ADAPT_THIN_CASE_LIMIT_SECONDS", 0, 0, 1E+30)
    If caseLimitSeconds > 0# Then
        caseElapsedSeconds = RPX_ThinRunElapsedSeconds()
        If caseElapsedSeconds < 0# And adaptStartedAt > 0# Then caseElapsedSeconds = RPX_Clock() - adaptStartedAt
        If caseElapsedSeconds >= 0# Then
            remaining = caseLimitSeconds - caseElapsedSeconds
            If remaining < 0# Then remaining = 0#
            remainingSource = "case_limit"
            Fix9DeadlineRemainingSeconds = remaining
            Exit Function
        End If
    End If
    limit = CDbl(RPX_Setting("ADAPT_OPTIONAL_DEADLINE_SECONDS", 0))
    If Not RPX_SchurFinite(limit) Or limit < 0# Then err.Raise 5, "RPX_Adapt", "INVALID_ADAPT_OPTIONAL_DEADLINE_SECONDS"
    If adaptStartedAt <= 0# Then Exit Function
    If limit > 0# Then
        remainingSource = "deadline"
    ElseIf nominalSeconds > 0# Then
        limit = nominalSeconds: remainingSource = "nominal"
    Else
        Exit Function
    End If
    remaining = limit - (RPX_Clock() - adaptStartedAt)
    If remaining < 0# Then remaining = 0#
    Fix9DeadlineRemainingSeconds = remaining
End Function
' The allowance is latched for one upper step, including geometry and every
' root/native/recovery/HSD trial. A retry can never replenish the case reserve.
Private Sub Fix9UpperProtectBegin(ByRef opt As Fix9ThinOptions)
    Dim plannedElements As Long, remainingSeconds As Double, remainingSource As String, reason As String
    fix9UpperProtectActive = False
    fix9UpperProtectStarted = RPX_DiagClock()
    fix9UpperProtectAllowance = 0#: fix9UpperProtectCost = 0#: fix9UpperProtectReserve = opt.lowerExtraReserve
    reason = "disabled"
    If Not opt.lowerExtraProtect Then GoTo ProtectionLogged
    reason = "no_extra_lower_pending"
    If Not opt.lowerExtraPending Then GoTo ProtectionLogged
    reason = "no_lower_cost"
    If opt.lowerCostSeconds <= 0# Or opt.lowerCostElements <= 0 Or opt.lowerProtectElements <= 0 Then GoTo ProtectionLogged
    ' Reserve the next extra only; its successor uses a newly measured cost.
    plannedElements = Fix9LowerPlannedElements(opt.lowerProtectElements, opt.lowerProtectCap, opt)
    reason = "no_useful_lower_size"
    If plannedElements <= opt.lowerProtectElements + 1 Then GoTo ProtectionLogged
    fix9UpperProtectCost = Fix9LowerPredictedSeconds(plannedElements, opt)
    remainingSeconds = RPX_ThinRemainingSeconds(remainingSource)
    reason = "no_remaining_clock"
    If remainingSeconds < 0# Then GoTo ProtectionLogged
    reason = "lower_not_affordable"
    If remainingSeconds < fix9UpperProtectCost + fix9UpperProtectReserve Then GoTo ProtectionLogged
    fix9UpperProtectActive = True: reason = "next_extra_lower"
    opt.lowerProtectFunded = True
    fix9UpperProtectAllowance = remainingSeconds - fix9UpperProtectCost - fix9UpperProtectReserve - Fix9UpperProtectMargin
    If fix9UpperProtectAllowance < 0# Then fix9UpperProtectAllowance = 0#
ProtectionLogged:
    RPX_DiagEvent "thin_layer_upper_protect", "phase=begin;active=" & Abs(CLng(fix9UpperProtectActive)) & ";reason=" & reason & ";lower_lineage_elements=" & opt.lowerProtectElements & ";planned_elements=" & plannedElements & ";last_lower_elements=" & opt.lowerCostElements & ";last_lower_seconds=" & opt.lowerCostSeconds & ";predicted_lower_seconds=" & fix9UpperProtectCost
    If fix9UpperProtectActive Then
        RPX_DiagEvent "thin_layer_upper_protect", "phase=allowance;remaining_seconds=" & remainingSeconds & ";remaining_source=" & remainingSource & ";reserve_seconds=" & fix9UpperProtectReserve & ";margin_seconds=" & Fix9UpperProtectMargin & ";step_allowance_seconds=" & fix9UpperProtectAllowance
    End If
End Sub
Public Function RPX_ThinUpperProtectOn() As Boolean
    RPX_ThinUpperProtectOn = (RPX_ThinChainBudgetActive And RPX_ThinChainUpperActive And fix9UpperProtectActive)
End Function
' Called only through thin upper solver gates; seconds includes any HSD grace.
Public Function RPX_ThinUpperProtectLimit(ByVal seconds As Double, ByVal purpose As String) As Double
    Dim elapsedSeconds As Double, availableSeconds As Double, remainingSeconds As Double, remainingSource As String, caseAllowance As Double
    RPX_ThinUpperProtectLimit = seconds
    If Not RPX_ThinUpperProtectOn() Then Exit Function
    elapsedSeconds = RPX_DiagClock() - fix9UpperProtectStarted
    availableSeconds = fix9UpperProtectAllowance - elapsedSeconds
    remainingSeconds = RPX_ThinRemainingSeconds(remainingSource)
    If remainingSeconds >= 0# Then
        caseAllowance = remainingSeconds - fix9UpperProtectCost - fix9UpperProtectReserve - Fix9UpperProtectMargin
        If caseAllowance < availableSeconds Then availableSeconds = caseAllowance
    End If
    If availableSeconds < 0# Then availableSeconds = 0#
    If seconds > availableSeconds Then RPX_ThinUpperProtectLimit = availableSeconds
    If seconds > 0# Or availableSeconds <= 0# Then
        RPX_DiagEvent "thin_layer_upper_protect", "phase=solve;purpose=" & purpose & ";step_seconds=" & elapsedSeconds & ";step_allowance_seconds=" & fix9UpperProtectAllowance & ";remaining_seconds=" & remainingSeconds & ";remaining_source=" & remainingSource & ";predicted_lower_seconds=" & fix9UpperProtectCost & ";reserve_seconds=" & fix9UpperProtectReserve & ";margin_seconds=" & Fix9UpperProtectMargin & ";available_seconds=" & availableSeconds & ";requested_seconds=" & seconds & ";applied_seconds=" & RPX_ThinUpperProtectLimit
    End If
    If availableSeconds <= 0# Then
        ' Not a numerical unknown: propagate past Fs numeric recovery immediately.
        ' ChainStep alone consumes this token and retains independently valid best.
        err.Raise 5, "RPX_Adapt", "THIN_UPPER_LOWER_PROTECT"
    End If
End Function
Public Sub RPX_ThinUpperProtectCheck(ByVal purpose As String)
    Dim checkedSeconds As Double
    If Not RPX_ThinUpperProtectOn() Then Exit Sub
    checkedSeconds = RPX_ThinUpperProtectLimit(0#, purpose)
End Sub
Private Sub Fix9ClearBudgets()
    Call RPX_ThinHsdWarmClear
    Call RPX_ThinFsEnd
    RPX_ThinChainBudgetActive = False: RPX_ThinChainUpperActive = False
    fix9ThinNominalSeconds = 0#
    fix9UpperProtectActive = False: fix9UpperProtectStarted = 0#: fix9UpperProtectAllowance = 0#
    fix9UpperProtectCost = 0#: fix9UpperProtectReserve = 0#
    RPX_ThinHsdTimeFraction = 0#: RPX_ThinHsdTimeCap = 0#
    RPX_ThinHsdBase = 0#: RPX_ThinHsdPerCone = 0#: RPX_ThinHsdFactor = 0#: RPX_ThinHsdCap = 0#
    RPX_ThinDllFactor = 0#: RPX_ThinDllRefSeconds = 0#: RPX_ThinDllRefCones = 0#: RPX_ThinDllExponent = 0#: RPX_ThinDllCap = 0#
    RPX_ThinUpperDllRetry = False: RPX_ThinUpperDllMinElements = 0: RPX_ThinUpperDllRetrySeconds = 0#
    RPX_ThinUpperDllRetryIterations = 0: RPX_ThinUpperDllRetryEquilibrate = 0
    RPX_ThinUpperHsdMinElements = 0: RPX_ThinUpperHsdFloorSeconds = 0#
    RPX_ThinLowerHsdMinElements = 0: RPX_ThinLowerHsdFloorSeconds = 0#
    RPX_ThinLowerHsdReserveSeconds = 0#: RPX_ThinLowerExtraActive = False
    RPX_ThinUpperHsdReserveSeconds = 0#: RPX_ThinUpperHsdIterations = 0
End Sub
' Exactly one DLL solve: no RPX_Bound/FsEndpoint/recovery and no witness gate.
' Failure returns to the old geometry method. Cancellation is restored/re-raised.
Private Function Fix9AuxUpper(ByVal factor As Double, ByVal purpose As String, ByRef velocity() As Double) As Boolean
    Dim pr As RPX_AccuracyProgram, answer() As Double, dual() As Double
    Dim audit As RPX_AuditSnapshot, oldX() As Double, hadX As Boolean, oldActive As Boolean
    Dim oldQ As Long, oldTarget As Long, oldTol As Double, oldIncomplete As Boolean, oldWidth As Double, oldNote As String, oldCalls As Long
    Dim oldStatus As String, oldValue As Double, oldObjective As Double, oldAudit As Double, oldFactor As Double, oldOrder As String, oldProgramFs As Double
    Dim oldRejected As Boolean, oldGate As Boolean, oldReason As String, oldRaw(1 To 4) As Double
    Dim oldDissipation As Double, oldReference As Double, oldFixed As Double, oldLower As Double, oldUpper As Double
    Dim oldLowerField() As Double, oldUpperField() As Double, hadLower As Boolean, hadUpper As Boolean
    Dim oldNativeReason As String, oldNativeStatus As Long, oldNativeIterations As Long
    Dim oldUpperBudget As Boolean
    Dim seconds As Double, started As Double, reason As String, ok As Boolean, i As Long, number As Long, source As String, message As String
    Erase velocity
    If Not RPX_ThinChainBudgetActive Then Exit Function
    oldQ = RSubdivisions: oldTarget = DTarget: oldTol = RFsTolerance: oldActive = RFsActive
    oldIncomplete = RFsSearchIncomplete: oldWidth = RFsLargestRootBracket: oldNote = RFsSearchNote: oldCalls = RFsCalls
    oldStatus = XStatus: oldValue = RValue: oldObjective = XObjective: oldAudit = RAudit: oldFactor = RStrengthFactor
    oldOrder = RPX_OrderMode: oldProgramFs = RProgramFs
    oldNativeReason = AccuracyNativeReason: oldNativeStatus = AccuracyNativeSolverStatus: oldNativeIterations = AccuracyNativeIterations
    oldUpperBudget = RPX_ThinChainUpperActive
    oldRejected = RPX_AuditRejected: oldGate = RPhysicalGateAccepted: oldReason = RPhysicalGateReason
    oldDissipation = RDissipation: oldReference = RReferenceWork: oldFixed = RFixedWork
    For i = 1 To 4: oldRaw(i) = RRawLowerAudit(i): Next i
    hadX = VectorReady(xx, 0): If hadX Then oldX = xx
    oldLower = RLowerFs: oldUpper = RUpperFs
    hadLower = VectorReady(RLowerField, 0): If hadLower Then oldLowerField = RLowerField
    hadUpper = VectorReady(RUpperField, 0): If hadUpper Then oldUpperField = RUpperField
    RPX_DiagAuditSave audit
    On Error GoTo AuxFailed
    If XCancel Then err.Raise 18, "RPX_Adapt", "ANALYSIS_CANCELLED"
    reason = "scope"
    RPX_ThinChainUpperActive = True
    RSubdivisions = 8: RFsActive = True: RStrengthFactor = factor: RPX_OrderMode = "upper"
    Call RPX_SchurDeactivate: Call RPX_LDLRelease: Call RPX_CapacityAssembly("upper")
    Call RPX_AssembleUpper: Call RPX_ResetPreparedState: Call RPX_Compile
    If Not RPX_AccuracyExport(factor, "upper", pr) Then GoTo AuxDone
    seconds = Fix8Number("DIRECT_DLL_SECONDS", 5, 0.000001, 600)
    seconds = RPX_ThinUpperDllSeconds(seconds, pr.coneCount - 1)
    started = RPX_DiagClock(): reason = "solve_failed_or_timeout"
    If Not RPX_AccuracySolve(pr, seconds, answer, dual) Then GoTo AuxDone
    If RPX_DiagClock() - started > seconds Then reason = "time_limit": GoTo AuxDone
    ' Native status is the Clarabel enum: 1 Solved, 4 AlmostSolved (rpfem_accuracy.h).
    ' Fix9 rc2: the velocity is only an indicator (never a witness), so AlmostSolved
    ' is accepted unless ADAPT_THIN_AUX_ALMOST=0. Time-limit/error iterates are not.
    If AccuracyNativeSolverStatus <> 1 Then
        If AccuracyNativeSolverStatus <> 4 Or Fix9Integer("ADAPT_THIN_AUX_ALMOST", 1, 0, 1) = 0 Then reason = "solver_status_" & AccuracyNativeSolverStatus: GoTo AuxDone
    End If
    If Not VectorReady(answer, 12 * re - 1) Then reason = "velocity_size": GoTo AuxDone
    ReDim velocity(0 To 12 * re - 1)
    For i = 0 To UBound(velocity)
        If Not RPX_SchurFinite(answer(i)) Then reason = "nonfinite_velocity": GoTo AuxDone
        velocity(i) = answer(i)
    Next i
    ok = True: reason = "velocity_only"
AuxDone:
    On Error GoTo AuxRestoreFailed
    RPX_ThinChainUpperActive = oldUpperBudget
    ' Raw assembly/compile does not mutate mesh, loads, lower layout, witnesses,
    ' certificates or brackets. Invalidate compiled work as ChainStep does, but
    ' avoid RestoreWitnessState's input assertion here so cancellation can clean up.
    RProgramKind = "": Call RPX_ResetPreparedState
    RSubdivisions = oldQ: DTarget = oldTarget: RFsTolerance = oldTol: RFsActive = oldActive
    RFsSearchIncomplete = oldIncomplete: RFsLargestRootBracket = oldWidth: RFsSearchNote = oldNote: RFsCalls = oldCalls
    Call RPX_DiagAuditRestore(audit)
    RStrengthFactor = oldFactor: RPX_OrderMode = oldOrder: RProgramFs = oldProgramFs
    AccuracyNativeReason = oldNativeReason: AccuracyNativeSolverStatus = oldNativeStatus: AccuracyNativeIterations = oldNativeIterations
    XStatus = oldStatus: RValue = oldValue: XObjective = oldObjective: RAudit = oldAudit
    RPX_AuditRejected = oldRejected: RPhysicalGateAccepted = oldGate: RPhysicalGateReason = oldReason
    RDissipation = oldDissipation: RReferenceWork = oldReference: RFixedWork = oldFixed
    For i = 1 To 4: RRawLowerAudit(i) = oldRaw(i): Next i
    RLowerFs = oldLower: RUpperFs = oldUpper
    If hadLower Then RLowerField = oldLowerField Else Erase RLowerField
    If hadUpper Then RUpperField = oldUpperField Else Erase RUpperField
    If hadX Then xx = oldX Else Erase xx
    If Not ok Then Erase velocity
    RPX_DiagEvent "thin_layer_aux", "purpose=" & purpose & ";Fs=" & factor & ";elements=" & re & ";ok=" & Abs(CLng(ok)) & ";reason=" & reason & ";seconds_limit=" & seconds & ";witness=0"
    On Error GoTo 0
    If number = 18 Or XCancel Then err.Raise 18, source, "ANALYSIS_CANCELLED"
    If number = 5 And message = "THIN_UPPER_LOWER_PROTECT" Then err.Raise number, source, message
    Fix9AuxUpper = ok
    Exit Function
AuxFailed:
    number = err.number: source = err.source: message = err.description
    reason = "error_" & number & ":" & message: ok = False
    Resume AuxDone
AuxRestoreFailed:
    number = err.number: source = err.source: message = err.description
    Call Fix9ClearBudgets
    On Error GoTo 0
    If number = 18 Or XCancel Then err.Raise 18, source, "ANALYSIS_CANCELLED FIX9_AUX_RESTORE_FAILED " & message
    err.Raise number, source, "FIX9_AUX_RESTORE_FAILED " & message
End Function
Private Function Fix9BuildMid(ByRef lineage As RPX_MeshState, ByRef mask() As Boolean, ByVal cap As Long, ByRef detail As String, ByVal upperMid As Boolean, ByRef opt As Fix9ThinOptions) As Boolean
    Dim marks As Object, count As Long
    Dim midScore() As Double, roundLimit As Long, predictedElements As Long, originalMarks As Long
    Call RPX_RestoreWitnessState(lineage)
    Set marks = RPX_Row(): count = RPX_TransverseEdges(mask, marks)
    If upperMid Then
        Call RPX_ThinUpperScore(lineage.field, midScore)
        roundLimit = Fix9UpperElementLimit(lineage.elements, cap, opt)
        originalMarks = marks.count
        Call Fix9LimitMarks(marks, midScore, roundLimit, predictedElements)
        If marks.count < originalMarks Then
            RPX_DiagEvent "thin_layer_upper_cap", "mid_only=1;lineage_elements=" & lineage.elements & ";round_limit=" & roundLimit & ";requested_edges=" & originalMarks & ";kept_edges=" & marks.count & ";predicted_elements=" & predictedElements
        End If
    End If
    If marks.count = 0 Then detail = "no_edges;mid_only=1": Exit Function
    Call Fix8Capacity: Call RPX_SplitEdgeSet(marks)
    If upperMid Then
        If re <> predictedElements Or re > roundLimit Then err.Raise 5, "RPX_Adapt", "FIX9_UPPER_MID_CAP_PREDICTION"
    End If
    If re > cap Then detail = "cap;mid_only=1;elements=" & re: Exit Function
    Call Fix8Commit
    detail = "mid_only=1;mid_edges=" & count & ";directional_edges=0;red_edges=0;fan=0;flips=0;elements=" & re
    Fix9BuildMid = True
End Function
Private Function Fix9UpperElementLimit(ByVal lineageElements As Long, ByVal cap As Long, ByRef opt As Fix9ThinOptions) As Long
    Dim roundLimit As Long
    roundLimit = opt.upperRoundMaxElements
    If roundLimit < lineageElements Then roundLimit = lineageElements
    If cap < roundLimit Then roundLimit = cap
    If CDbl(lineageElements) * opt.upperMaxGrowth < roundLimit Then roundLimit = CLng(Fix(CDbl(lineageElements) * opt.upperMaxGrowth))
    Fix9UpperElementLimit = roundLimit
End Function
' Exact red/green split cost: one extra element per marked edge incidence.
' When the cap binds, rank estimated GAP reduction per added element: incident
' indicator mass divided by the exact red/green edge-incidence cost (one for
' boundary edges, two for interior edges). These scores are geometry hints only.
Private Sub Fix9LimitMarks(ByRef marks As Object, ByRef score() As Double, ByVal elementLimit As Long, ByRef predictedElements As Long)
    Dim edgeCost As Object, edgeMass As Object, kept As Object, edgeKeys As Variant
    Dim edgeOrder() As Long, edgeScore() As Double, e As Long, k As Long, j As Long, edgeKey As String
    Set edgeCost = RPX_Row(): Set edgeMass = RPX_Row(): Set kept = RPX_Row()
    predictedElements = re
    For e = 0 To re - 1
        For k = 0 To 2
            edgeKey = RPX_EdgeKey(RTri(e, k), RTri(e, (k + 1) Mod 3))
            If marks.Exists(edgeKey) Then
                If Not edgeCost.Exists(edgeKey) Then edgeCost(edgeKey) = 0: edgeMass(edgeKey) = 0#
                edgeCost(edgeKey) = CLng(edgeCost(edgeKey)) + 1
                edgeMass(edgeKey) = CDbl(edgeMass(edgeKey)) + score(e)
                predictedElements = predictedElements + 1
            End If
        Next k
    Next e
    If predictedElements <= elementLimit Then Exit Sub
    predictedElements = re
    If edgeCost.count = 0 Then Set marks = kept: Exit Sub
    edgeKeys = edgeCost.keys
    ReDim edgeOrder(0 To edgeCost.count - 1): ReDim edgeScore(0 To edgeCost.count - 1)
    For j = 0 To edgeCost.count - 1
        edgeOrder(j) = j: edgeScore(j) = CDbl(edgeMass(CStr(edgeKeys(j)))) / CDbl(edgeCost(CStr(edgeKeys(j))))
    Next j
    If edgeCost.count > 1 Then Fix9EdgeSort edgeOrder, edgeScore, 0, edgeCost.count - 1
    For j = 0 To edgeCost.count - 1
        edgeKey = CStr(edgeKeys(edgeOrder(j)))
        If predictedElements + CLng(edgeCost(edgeKey)) <= elementLimit Then
            kept(edgeKey) = 1: predictedElements = predictedElements + CLng(edgeCost(edgeKey))
        End If
    Next j
    Set marks = kept
End Sub
Private Sub Fix9EdgeSort(ByRef edgeOrder() As Long, ByRef edgeScore() As Double, ByVal firstIndex As Long, ByVal lastIndex As Long)
    Dim lowerIndex As Long, upperIndex As Long, pivotScore As Double, swapIndex As Long
    lowerIndex = firstIndex: upperIndex = lastIndex: pivotScore = edgeScore(edgeOrder((firstIndex + lastIndex) \ 2))
    Do While lowerIndex <= upperIndex
        Do While edgeScore(edgeOrder(lowerIndex)) > pivotScore: lowerIndex = lowerIndex + 1: Loop
        Do While edgeScore(edgeOrder(upperIndex)) < pivotScore: upperIndex = upperIndex - 1: Loop
        If lowerIndex <= upperIndex Then
            swapIndex = edgeOrder(lowerIndex): edgeOrder(lowerIndex) = edgeOrder(upperIndex): edgeOrder(upperIndex) = swapIndex
            lowerIndex = lowerIndex + 1: upperIndex = upperIndex - 1
        End If
    Loop
    If firstIndex < upperIndex Then Fix9EdgeSort edgeOrder, edgeScore, firstIndex, upperIndex
    If lowerIndex < lastIndex Then Fix9EdgeSort edgeOrder, edgeScore, lowerIndex, lastIndex
End Sub
' Funded extra lower geometry uses the highest positive indicator regions.
' Select complete red parent regions, charging only not-yet-marked edge
' incidences, so adjacent selections share cost and conforming neighbors fit.
' Zero indicator mass does not justify spending the optional extra budget.
Private Function Fix9TargetLowerRegion(ByRef score() As Double, ByVal theta As Double, ByVal cap As Long, ByRef marks As Object, ByRef predictedElements As Long, ByRef selectedRegions As Long) As Long
    Dim regionOrder() As Long, edgeCost As Object, edgeKey As String
    Dim e As Long, k As Long, j As Long, id As Long, needed As Long
    Dim totalMass As Double, selectedMass As Double
    Set marks = RPX_Row(): Set edgeCost = RPX_Row()
    predictedElements = re: selectedRegions = 0
    If cap <= re Then Exit Function
    ReDim regionOrder(0 To re - 1)
    For e = 0 To re - 1
        If Not RPX_SchurFinite(score(e)) Or score(e) < 0# Then err.Raise 5, "RPX_Adapt", "FIX9_TARGET_LOWER_SCORE"
        regionOrder(e) = e: totalMass = totalMass + score(e)
    Next e
    If totalMass <= 0# Then Exit Function
    If re > 1 Then Fix9EdgeSort regionOrder, score, 0, re - 1
    For id = 0 To RNEdge - 1
        edgeKey = RPX_EdgeKey(REdges(id).a, REdges(id).b)
        If REdges(id).element(1) < 0 Then edgeCost(edgeKey) = 1 Else edgeCost(edgeKey) = 2
    Next id
    For j = 0 To re - 1
        If XCancel Then err.Raise 18, , "ANALYSIS_CANCELLED"
        e = regionOrder(j)
        If score(e) <= 0# Or selectedMass >= theta * totalMass Then Exit For
        needed = 0
        For k = 0 To 2
            edgeKey = RPX_EdgeKey(RTri(e, k), RTri(e, (k + 1) Mod 3))
            If Not marks.Exists(edgeKey) Then needed = needed + CLng(edgeCost(edgeKey))
        Next k
        If predictedElements + needed <= cap Then
            For k = 0 To 2
                edgeKey = RPX_EdgeKey(RTri(e, k), RTri(e, (k + 1) Mod 3)): marks(edgeKey) = 1
            Next k
            predictedElements = predictedElements + needed
            selectedRegions = selectedRegions + 1: selectedMass = selectedMass + score(e)
        End If
    Next j
    Fix9TargetLowerRegion = marks.count
    RPX_DiagEvent "thin_layer_lower_target", "cap=" & cap & ";regions=" & selectedRegions & ";kept_edges=" & marks.count & ";predicted_elements=" & predictedElements & ";indicator_fraction=" & selectedMass / totalMass & ";global_mid_and_fan=0"
End Function
' Upper geometry: directional one-edge split from the lineage velocity, the
' thin-layer mid-line on the first step, then flow-metric edge flips.
Private Function Fix8BuildThinUpper(ByRef lineage As RPX_MeshState, ByVal withMid As Boolean, ByRef mask() As Boolean, ByVal theta As Double, ByVal cap As Long, ByVal passes As Long, ByRef detail As String, ByRef opt As Fix9ThinOptions) As Boolean
    Dim vel() As Double, score() As Double, marks As Object, nDir As Long, nMid As Long, i As Long
    Dim pxy() As Double, ptri() As Long, pne As Long, cvel() As Double, pts() As Double, cells() As Long, mats() As Long
    Dim child() As Long, info As Object, original() As Double, boundaryCount As Long, flips As Long, nodes As Long, elements As Long
    Dim fresh As Boolean, roundLimit As Long, predictedElements As Long, originalMarks As Long
    Call RPX_RestoreWitnessState(lineage)
    ReDim vel(0 To 12 * re - 1)
    For i = 0 To 12 * re - 1: vel(i) = lineage.field(i): Next i
    Call RPX_ThinUpperScore(vel, score)
    Set marks = RPX_Row()
    nDir = RPX_DirectionalEdges(vel, score, theta, marks)
    If withMid Then nMid = RPX_TransverseEdges(mask, marks)
    roundLimit = Fix9UpperElementLimit(lineage.elements, cap, opt)
    originalMarks = marks.count
    Call Fix9LimitMarks(marks, score, roundLimit, predictedElements)
    If marks.count < originalMarks Then
        RPX_DiagEvent "thin_layer_upper_cap", "lineage_elements=" & lineage.elements & ";round_limit=" & roundLimit & ";max_growth=" & opt.upperMaxGrowth & ";requested_edges=" & originalMarks & ";kept_edges=" & marks.count & ";predicted_elements=" & predictedElements
    End If
    If marks.count = 0 Then detail = "no_edges": Exit Function
    pxy = lineage.xy: ptri = lineage.triangles: pne = re
    Call Fix8Capacity
    Call RPX_SplitEdgeSet(marks)
    If re <> predictedElements Or re > roundLimit Then err.Raise 5, "RPX_Adapt", "FIX9_UPPER_CAP_PREDICTION"
    Call Fix8Commit
    If passes > 0 Then
        If opt.flipFresh Then fresh = Fix9AuxUpper(lineage.value, "flip_fresh", cvel)
        If Not fresh Then
            If Not RPX_ProlongUpperVelocity(pxy, ptri, pne, vel, cvel) Then err.Raise 5, "RPX_Adapt", "FIX8_PROLONG_LOCATE_FAILED"
        End If
        ReDim pts(0 To rn - 1, 0 To 1): ReDim cells(0 To re - 1, 0 To 2): ReDim mats(0 To re - 1)
        For i = 0 To rn - 1: pts(i, 0) = RXY(i, 0): pts(i, 1) = RXY(i, 1): Next i
        For i = 0 To re - 1: cells(i, 0) = RTri(i, 0): cells(i, 1) = RTri(i, 1): cells(i, 2) = RTri(i, 2): mats(i) = RMatId(i): Next i
        If RPX_FlowMeshCandidate(pts, cells, mats, cvel, RConstraints, child, info, passes) Then
            If CBool(info("certificate")) Or Not CBool(info("requires_new_original_q8_field_and_root")) Then err.Raise 5, "RPX_Adapt", "FIX8_FLOW_SCOPE"
            flips = CLng(info("flips"))
            If flips > 0 Then
                Call G20MaterialMoments(original)
                boundaryCount = RBoundary.count: nodes = rn: elements = re
                Call Fix8Capacity
                For i = 0 To re - 1: RTri(i, 0) = child(i, 0): RTri(i, 1) = child(i, 1): RTri(i, 2) = child(i, 2): Next i
                Set RBoundary = Nothing
                Call RPX_Topology: Call RPX_MapFaces
                Call Fix8Commit
                If rn <> nodes Or re <> elements Then err.Raise 5, "RPX_Adapt", "FIX8_FLOW_DIMENSIONS_CHANGED"
                Call G20CompareMoments(original)
                If RBoundary.count <> boundaryCount Then err.Raise 5, "RPX_Adapt", "FIX8_FLOW_BOUNDARY_COUNT"
            End If
        End If
    End If
    If re > roundLimit Then err.Raise 5, "RPX_Adapt", "FIX9_UPPER_CAP_AFTER_FLIP"
    detail = "directional_edges=" & nDir & ";mid_edges=" & nMid & ";flips=" & flips & ";elements=" & re & ";kept_edges=" & marks.count & ";round_limit=" & roundLimit
    Fix8BuildThinUpper = True
End Function
' Lower geometry: mixed indicator (best-upper kinematic density mapped by
' point location + lower stress score), Doerfler red refinement, mid-line on
' the first step, then the angular fan of zero-stress free vertices.
Private Function Fix8BuildThinLower(ByRef lineage As RPX_MeshState, ByVal withMid As Boolean, ByRef mask() As Boolean, ByVal theta As Double, ByVal cap As Long, ByRef detail As String, ByRef opt As Fix9ThinOptions) As Boolean
    Dim su() As Double, density() As Double, sl() As Double, mixed() As Double, marked() As Boolean, marks As Object
    Dim uxy() As Double, utri() As Long, une As Long, e As Long, p As Long, lam(0 To 2) As Double, cx As Double, cy As Double
    Dim su2() As Double, sumU As Double, sumL As Double, nRed As Long, nMid As Long, missed As Long, parent() As Long, fan As Boolean
    Dim ownUpper As Boolean, vel() As Double, predictedElements As Long, targetedRegions As Long
    If opt.lowerOwnUpper Then
        Call RPX_RestoreWitnessState(lineage)
        ownUpper = Fix9AuxUpper(opt.lowerOwnFactor * lineage.value, "lower_own_upper", vel)
        If ownUpper Then Call RPX_ThinUpperScore(vel, su2)
    End If
    If Not ownUpper Then
    Call RPX_RestoreWitnessState(RBestUpper)
    Call RPX_ThinUpperScore(RBestUpper.field, su)
    ReDim density(0 To re - 1)
    For e = 0 To re - 1: If RArea(e) > 0# Then density(e) = su(e) / RArea(e)
    Next e
    uxy = RBestUpper.xy: utri = RBestUpper.triangles: une = re
    End If
    Call RPX_RestoreWitnessState(lineage)
    Call RPX_ThinLowerScore(lineage.field, sl)
    If Not ownUpper Then
        Call RPX_LocateBuild(uxy, utri, une)
        ReDim su2(0 To re - 1)
    End If
    ReDim mixed(0 To re - 1)
    For e = 0 To re - 1
        If Not ownUpper Then
        cx = (RXY(RTri(e, 0), 0) + RXY(RTri(e, 1), 0) + RXY(RTri(e, 2), 0)) / 3#
        cy = (RXY(RTri(e, 0), 1) + RXY(RTri(e, 1), 1) + RXY(RTri(e, 2), 1)) / 3#
        p = RPX_LocatePoint(cx, cy, lam)
        If p >= 0 Then su2(e) = density(p) * RArea(e) Else missed = missed + 1
        End If
        sumU = sumU + su2(e): sumL = sumL + sl(e)
    Next e
    For e = 0 To re - 1
        If sumU > 0# Then mixed(e) = su2(e) / sumU
        If sumL > 0# Then mixed(e) = mixed(e) + sl(e) / sumL
    Next e
    Set marks = RPX_Row()
    If opt.lowerExtraActive Then
        nRed = Fix9TargetLowerRegion(mixed, theta, cap, marks, predictedElements, targetedRegions)
    Else
        Call RPX_DoerflerMark(mixed, theta, marked)
        nRed = RPX_RedEdges(marked, marks)
        If withMid Then nMid = RPX_TransverseEdges(mask, marks)
    End If
    If marks.count = 0 Then detail = "no_edges": Exit Function
    Call Fix8Capacity
    Call RPX_SplitEdgeSet(marks)
    If opt.lowerExtraActive Then
        If re <> predictedElements Or re > cap Then err.Raise 5, "RPX_Adapt", "FIX9_LOWER_TARGET_CAP_PREDICTION"
    End If
    If re > cap Then detail = "cap;elements=" & re: Exit Function
    Call Fix8Commit
    ' Earlier lower rounds keep the global free-vertex fan. The optional extra
    ' spends its planned cap on the selected region and conforming neighbors.
    If Not opt.lowerExtraActive Then fan = RPX_AngularFanMesh(cap, parent)
    detail = "red_edges=" & nRed & ";mid_edges=" & nMid & ";unlocated=" & missed & ";fan=" & IIf(fan, "1", "0") & ";elements=" & re
    Fix8BuildThinLower = True
End Function
' One chain step. Returns 0 = stop chain, 1 = accepted as new best witness,
' 2 = valid own root but not better (lineage advances, witness unchanged).
' evaluated reports reaching endpoint evaluation, independently of the result.
Private Function Fix8ChainStep(ByVal side As String, ByVal roundIndex As Long, ByRef lineage As RPX_MeshState, ByVal withMid As Boolean, ByRef mask() As Boolean, ByVal theta As Double, ByVal cap As Long, ByVal passes As Long, ByRef opt As Fix9ThinOptions, ByVal midOnly As Boolean, ByRef stopReason As String, ByRef evaluated As Boolean) As Long
    Dim baseline As RPX_MeshState, pending As RPX_MeshState
    Dim field() As Double, eta() As Double, factor As Double, objective As Double, a As Double, b As Double
    Dim savedQ As Long, savedTarget As Long, savedTol As Double, oldIncomplete As Boolean, built As Boolean, detail As String
    Dim result As Long, number As Long, source As String, message As String, label As String, stage As String
    Dim savedAudit As RPX_AuditSnapshot, originalNumber As Long, originalSource As String, originalMessage As String
    Dim oldStatus As String, oldValue As Double, oldObjective As Double, oldRAudit As Double, oldRejected As Boolean, oldGate As Boolean, oldReason As String
    Dim oldRaw(1 To 4) As Double, oldDissipation As Double, oldReference As Double, oldFixed As Double, i As Long
    Dim oldX() As Double, hadX As Boolean, oldFactor As Double, started As Double
    Dim predictedSeconds As Double, remainingSeconds As Double, remainingSource As String
    evaluated = False
    stopReason = "not_applicable"
    If Not currentPairValid Or Not RBestLower.complete Or Not RBestUpper.complete Or RR <> 0 Or RPX_WaterCouplingOn() Then Exit Function
    If Not lineage.complete Or lineage.q <> 8 Or lineage.searchIncomplete Then Exit Function
    If side = "upper" And Not lineage.upperRootValid Then Exit Function
    If side = "lower" And Not lineage.lowerRootValid Then Exit Function
    started = RPX_Clock()
    CloneState baseline, RCurrent
    savedQ = RSubdivisions: savedTol = RFsTolerance: savedTarget = DTarget: oldIncomplete = RFsSearchIncomplete
    oldStatus = XStatus: oldValue = RValue: oldObjective = XObjective: oldRAudit = RAudit: oldFactor = RStrengthFactor
    oldRejected = RPX_AuditRejected: oldGate = RPhysicalGateAccepted: oldReason = RPhysicalGateReason
    oldDissipation = RDissipation: oldReference = RReferenceWork: oldFixed = RFixedWork
    For i = 1 To 4: oldRaw(i) = RRawLowerAudit(i): Next i
    hadX = VectorReady(xx, 0): If hadX Then oldX = xx
    RPX_DiagAuditSave savedAudit
    label = "fix8_thin_" & side & "_" & roundIndex
    On Error GoTo CandidateFailed
    RPX_ThinHsdBase = opt.hsdBase: RPX_ThinHsdPerCone = opt.hsdPerCone: RPX_ThinHsdFactor = opt.hsdFactor: RPX_ThinHsdCap = opt.hsdCap
    RPX_ThinHsdTimeFraction = opt.hsdTimeFraction: RPX_ThinHsdTimeCap = opt.hsdTimeCap
    fix9ThinNominalSeconds = opt.nominalSeconds
    ' Apply the equilibration premium only to the unmeasured DLL forecast.
    RPX_ThinDllFactor = opt.dllFactor: RPX_ThinDllRefSeconds = opt.dllRefSeconds * opt.dllEquTimeFactor: RPX_ThinDllRefCones = opt.dllRefCones
    RPX_ThinDllExponent = opt.dllExponent: RPX_ThinDllCap = opt.dllCap
    RPX_ThinUpperDllRetry = opt.upperDllRetry: RPX_ThinUpperDllMinElements = opt.upperDllMinElements
    RPX_ThinUpperDllRetrySeconds = opt.upperDllRetrySeconds: RPX_ThinUpperDllRetryIterations = opt.upperDllRetryIterations
    RPX_ThinUpperDllRetryEquilibrate = opt.upperDllRetryEquilibrate
    RPX_ThinUpperHsdMinElements = opt.upperHsdMinElements: RPX_ThinUpperHsdFloorSeconds = opt.upperHsdFloorSeconds
    RPX_ThinLowerHsdMinElements = opt.lowerHsdMinElements: RPX_ThinLowerHsdFloorSeconds = opt.lowerHsdFloorSeconds
    RPX_ThinLowerHsdReserveSeconds = opt.lowerExtraReserve: RPX_ThinLowerExtraActive = (side = "lower" And opt.lowerExtraActive)
    RPX_ThinUpperHsdReserveSeconds = opt.upperHsdReserveSeconds
    RPX_ThinUpperHsdIterations = opt.upperHsdIterations
    RPX_ThinChainBudgetActive = True: RPX_ThinChainUpperActive = (side = "upper")
    RFsActive = False: stage = "geometry"
    If side = "upper" Then
        Call Fix9UpperProtectBegin(opt)
        Call RPX_ThinUpperProtectCheck("geometry")
    End If
    If midOnly Then
        built = Fix9BuildMid(lineage, mask, cap, detail, (side = "upper"), opt)
    ElseIf side = "upper" Then
        built = Fix8BuildThinUpper(lineage, withMid, mask, theta, cap, passes, detail, opt)
    Else
        built = Fix8BuildThinLower(lineage, withMid, mask, theta, cap, detail, opt)
    End If
    If Not built Then
        stopReason = Split(detail, ";")(0)
        RPX_DiagEvent "thin_layer_step", "side=" & side & ";round=" & roundIndex & ";evaluated=0;result=0;reason=" & stopReason & ";" & detail
        GoTo CandidateDone
    End If
    If side = "lower" And opt.lowerExtraActive Then
        predictedSeconds = Fix9LowerPredictedSeconds(re, opt)
        remainingSeconds = Fix9DeadlineRemainingSeconds(opt.nominalSeconds, remainingSource)
        RPX_DiagEvent "thin_layer_lower_extra_budget", "elements=" & re & ";old_elements=" & opt.lowerCostElements & ";last_seconds=" & opt.lowerCostSeconds & ";predicted_seconds=" & predictedSeconds & ";reserve_seconds=" & opt.lowerExtraReserve & ";remaining_seconds=" & remainingSeconds & ";remaining_source=" & remainingSource
        If remainingSeconds < predictedSeconds + opt.lowerExtraReserve Then
            stopReason = "time_budget"
            GoTo CandidateDone
        End If
    End If
    If side = "upper" Then Call RPX_ThinUpperProtectCheck("endpoint")
    stage = "evaluate"
    evaluated = True
    RSubdivisions = 8: RFsTolerance = CDbl(RPX_Setting("FS_TOL", 0.0001))
    If RFsTolerance <= 0# Or RFsTolerance > 0.0001 Then err.Raise 5, "RPX_Adapt", "FIX8_ROOT_TOLERANCE"
    RPX_DiagCandidate label: RPX_TestHook label
    RFsActive = True: RFsCalls = 0: RFsSearchIncomplete = False: RFsLargestRootBracket = 0#: RFsSearchNote = ""
    stopReason = "root_failed"
    Call RPX_ThinFsBegin(side)
    factor = RPX_FsEndpoint(side, lineage.value, AdaptFsHintRatio)
    If side = "upper" Then Call RPX_ThinUpperProtectCheck("endpoint_complete")
    objective = XObjective
    If side = "upper" Then
        ReDim field(0 To 12 * re - 1)
        For i = 0 To UBound(field): field(i) = xx(i): Next i
        RUpperFs = factor: RUpperField = field: Erase RLowerField
    Else
        field = xx
        RLowerFs = factor: RLowerField = field: Erase RUpperField
    End If
    ReDim eta(0 To re - 1)
    If Not RFsSearchIncomplete And RPX_ActualRootBracket(side, a, b, factor) Then
        stopReason = "evaluate_failed"
        If side = "lower" Then
            If factor <= RBestUpper.value Then
                If SaveState(pending, field, eta, factor, factor, RBestUpper.value, label, side, objective) Then
                    If pending.lowerRootValid And Not pending.searchIncomplete Then
                        CloneState lineage, pending: result = 2
                        If factor >= RBestLower.value Then CloneState RBestLower, pending: RPX_CertFromAdapt side: result = 1
                    End If
                End If
            Else
                stopReason = "crossed"
            End If
        Else
            If factor >= RBestLower.value Then
                If SaveState(pending, field, eta, factor, RBestLower.value, factor, label, side, objective) Then
                    If pending.upperRootValid And Not pending.searchIncomplete Then
                        CloneState lineage, pending: result = 2
                        If factor <= RBestUpper.value Then CloneState RBestUpper, pending: RPX_CertFromAdapt side: result = 1
                    End If
                End If
            Else
                stopReason = "crossed"
            End If
        End If
    End If
    If result > 0 Then stopReason = "valid"
    adaptFsCallsTotal = adaptFsCallsTotal + RFsCalls
    RPX_DiagEvent "thin_layer_step", "side=" & side & ";round=" & roundIndex & ";evaluated=1;result=" & result & ";reason=" & stopReason & ";mid_only=" & Abs(CLng(midOnly)) & ";theta=" & theta & ";Fs=" & factor & ";q=8;root_a=" & a & ";root_b=" & b & ";fs_calls=" & RFsCalls & ";seconds=" & Format$(RPX_Clock() - started, "0.0") & ";" & detail & ";other_own_mesh_retained=1"
CandidateDone:
    If result = 0 And (stopReason = "root_failed" Or stopReason = "evaluate_failed" Or stopReason = "upper_protect") Then
        Call Fix9KeepFailedBest(side, label, originalNumber, originalSource, originalMessage)
    End If
    Call Fix9ClearBudgets
    On Error GoTo RestoreFailed
    RFsActive = False: RFsTolerance = savedTol: DTarget = savedTarget
    Call RPX_RestoreWitnessState(baseline): CloneState RCurrent, baseline
    RSubdivisions = savedQ: RFsSearchIncomplete = oldIncomplete
    Call RPX_DiagAuditRestore(savedAudit): Call RestoreFinalBounds("Fs")
    RStrengthFactor = oldFactor: XStatus = oldStatus: RValue = oldValue: XObjective = oldObjective: RAudit = oldRAudit
    RPX_AuditRejected = oldRejected: RPhysicalGateAccepted = oldGate: RPhysicalGateReason = oldReason
    RDissipation = oldDissipation: RReferenceWork = oldReference: RFixedWork = oldFixed
    For i = 1 To 4: RRawLowerAudit(i) = oldRaw(i): Next i
    If hadX Then xx = oldX Else Erase xx
    Fix8ChainStep = result
    On Error GoTo 0
    If originalNumber <> 0 Then
        If originalNumber = 18 Then err.Raise 18, originalSource, "ANALYSIS_CANCELLED"
        err.Raise originalNumber, originalSource, originalMessage
    End If
    Exit Function
CandidateFailed:
    number = err.number: source = err.source: message = err.description
    On Error GoTo 0
    RPX_ErrorEvidence number, source, message, 0, "fix8_thin_" & stage
    result = 0
    stopReason = "geometry"
    If stage = "evaluate" Then stopReason = "evaluate_failed"
    If number = 18 Or XCancel Then
        stopReason = "cancelled"
        originalNumber = 18: originalSource = source: originalMessage = message
    ElseIf number = 5 And message = "THIN_UPPER_LOWER_PROTECT" Then
        stopReason = "upper_protect"
    ElseIf InStr(1, message, "FIX9_AUX_RESTORE_FAILED", vbBinaryCompare) = 1 Then
        stopReason = "fatal"
        originalNumber = number: originalSource = source: originalMessage = message
    ElseIf InStr(1, message, "HSD_STABLE_UNAVAILABLE_MEMORY", vbBinaryCompare) = 1 Then
        stopReason = "fatal"
        originalNumber = number: originalSource = source: originalMessage = message
    ElseIf stage = "evaluate" And Not RPX_Recoverable(number, message, "evaluate", True) Then
        stopReason = "fatal"
        originalNumber = number: originalSource = source: originalMessage = message
    End If
    ' Geometry-stage failures only end this optional chain; witnesses are kept.
    RPX_DiagEvent "thin_layer_rejected", "side=" & side & ";round=" & roundIndex & ";stage=" & stage & ";reason=" & stopReason & ";old_own_witnesses_retained=1;message=" & message
    GoTo CandidateDone
RestoreFailed:
    number = err.number: source = err.source: message = err.description
    Call Fix9ClearBudgets
    On Error GoTo 0
    If originalNumber = 18 Or number = 18 Or XCancel Then err.Raise 18, source, "ANALYSIS_CANCELLED FIX8_RESTORE_FAILED " & message
    err.Raise number, source, "FIX8_CANDIDATE_RESTORE_FAILED original=" & originalNumber & ":" & originalMessage & ";restore=" & message
End Function
' A failed search can retain only a real, narrow, same-mesh root endpoint.
' The collector never invents a root from an accepted field or a status-5 sign.
' Saving the independent best does not advance the failed step's lineage.
Private Sub Fix9KeepFailedBest(ByVal side As String, ByVal label As String, ByRef fatalNumber As Long, ByRef fatalSource As String, ByRef fatalMessage As String)
    Dim pending As RPX_MeshState, oldBest As RPX_MeshState, oldCert As RPX_BoundCertificate
    Dim field() As Double, eta() As Double, factor As Double, objective As Double
    Dim number As Long, source As String, message As String, changed As Boolean
    On Error GoTo Failed
    If Not RPX_ThinChainBudgetActive Or RSubdivisions <> 8 Or XCancel Then Exit Sub
    If Not RPX_ThinFsWitness(side, factor, field, objective) Then Exit Sub
    If side = "upper" Then
        If factor >= RBestUpper.value Or factor < RBestLower.value Then Exit Sub
    Else
        If factor <= RBestLower.value Or factor > RBestUpper.value Then Exit Sub
    End If
    If Not RPX_ThinFsRestoreWitnessRoot(side, factor, field, objective) Then Exit Sub
    ReDim eta(0 To re - 1)
    If side = "upper" Then
        RUpperFs = factor: RUpperField = field: Erase RLowerField
        If Not SaveState(pending, field, eta, factor, RBestLower.value, factor, label & "_keep", side, objective) Then Exit Sub
        If Not pending.upperRootValid Then Exit Sub
    Else
        RLowerFs = factor: RLowerField = field: Erase RUpperField
        If Not SaveState(pending, field, eta, factor, factor, RBestUpper.value, label & "_keep", side, objective) Then Exit Sub
        If Not pending.lowerRootValid Then Exit Sub
    End If
    If Not pending.complete Or pending.q <> 8 Or pending.searchIncomplete Then Exit Sub
    If Not TerminalOwnFieldReady(side, pending) Then Exit Sub
    If side = "upper" Then
        oldBest = RBestUpper: oldCert = RPX_CertUpper
        changed = True: CloneState RBestUpper, pending
    Else
        oldBest = RBestLower: oldCert = RPX_CertLower
        changed = True: CloneState RBestLower, pending
    End If
    RPX_CertFromAdapt side
    If Not TerminalOwnMetadataReady(side, 8, pending) Then GoTo Rejected
    RPX_DiagEvent "thin_layer_keep_best", "side=" & side & ";Fs=" & factor & ";elements=" & pending.elements
    Exit Sub
Rejected:
    If changed Then
        If side = "upper" Then
            RBestUpper = oldBest: RPX_CertUpper = oldCert
        Else
            RBestLower = oldBest: RPX_CertLower = oldCert
        End If
    End If
    Exit Sub
Failed:
    number = err.number: source = err.source: message = err.description
    If changed Then
        If side = "upper" Then
            RBestUpper = oldBest: RPX_CertUpper = oldCert
        Else
            RBestLower = oldBest: RPX_CertLower = oldCert
        End If
    End If
    On Error GoTo 0
    RPX_DiagEvent "thin_layer_keep_best_rejected", "side=" & side & ";message=" & Replace(message, ";", ",")
    If number = 18 Or XCancel Or number = 7 Or number = 28 Then
        fatalNumber = number: fatalSource = source: fatalMessage = message
        If XCancel Then fatalNumber = 18
    ElseIf Not RPX_Recoverable(number, message, "evaluate", True) Then
        fatalNumber = number: fatalSource = source: fatalMessage = message
    End If
End Sub
' Existing lower retries remain per round. Each side also has one extra retry
' per chain, funded by the deadline or nominal thin clock. Failed lineage stays put.
Private Function Fix9RetryStep(ByVal side As String, ByVal roundIndex As Long, ByRef lineage As RPX_MeshState, ByVal withMid As Boolean, ByRef mask() As Boolean, ByVal theta As Double, ByVal cap As Long, ByVal passes As Long, ByRef opt As Fix9ThinOptions, ByRef extraUsed As Boolean, ByRef stopReason As String, ByRef steps As Long, ByRef evaluated As Boolean, ByRef evaluatedSteps As Long) As Long
    Dim attempt As Long, ordinaryUsed As Long, r As Long, retryTheta As Double, ordinaryRetries As Long, remaining As Double, extra As Boolean
    Dim remainingSource As String, stepEvaluated As Boolean
    evaluated = False
    retryTheta = theta
    If side = "lower" Then
        ordinaryRetries = opt.lowerRetries
        If Not opt.lowerExtraActive And roundIndex >= 2 And lineage.elements > CDbl(opt.lowerLargeElements) / 2# Then
            retryTheta = retryTheta * opt.lowerLargeThetaFactor
            RPX_DiagEvent "thin_layer_lower_large", "round=" & roundIndex & ";lineage_elements=" & lineage.elements & ";threshold=" & opt.lowerLargeElements & ";theta=" & retryTheta
        End If
    End If
    Do
        r = Fix8ChainStep(side, roundIndex, lineage, withMid, mask, retryTheta, cap, passes, opt, False, stopReason, stepEvaluated)
        steps = steps + 1
        If stepEvaluated Then evaluated = True: evaluatedSteps = evaluatedSteps + 1
        If r <> 0 Then Exit Do
        If stopReason <> "root_failed" And stopReason <> "evaluate_failed" And stopReason <> "cap" Then Exit Do
        If XCancel Then err.Raise 18, , "ANALYSIS_CANCELLED"
        If OptionalStageExpired("thin_layer_retry_" & side & "_" & roundIndex) Then Exit Do
        remaining = Fix9DeadlineRemainingSeconds(opt.nominalSeconds, remainingSource)
        extra = False
        If ordinaryUsed < ordinaryRetries And stopReason <> "cap" Then
            ordinaryUsed = ordinaryUsed + 1: attempt = attempt + 1
        Else
            If extraUsed Then Exit Do
            If remaining <= opt.extraRetryMinSeconds Then Exit Do
            extraUsed = True: extra = True: attempt = attempt + 1
        End If
        retryTheta = retryTheta * opt.retryThetaFactor
        RPX_DiagEvent "thin_layer_retry", "side=" & side & ";round=" & roundIndex & ";retry=" & attempt & ";reason=" & stopReason & ";theta=" & retryTheta & ";lineage_elements=" & lineage.elements & ";lineage_Fs=" & lineage.value & ";mid_only=0;extra_time_retry=" & Abs(CLng(extra)) & ";remaining_seconds=" & remaining & ";remaining_source=" & remainingSource
    Loop
    Fix9RetryStep = r
End Function
' One geometry size and one cost model are shared by protection and admission.
Private Function Fix9LowerPlannedElements(ByVal lineageElements As Long, ByVal cap As Long, ByRef opt As Fix9ThinOptions) As Long
    Dim plannedElements As Long
    plannedElements = CLng(Fix(CDbl(lineageElements) * opt.lowerExtraMaxGrowth))
    If plannedElements > opt.lowerExtraTargetElements Then plannedElements = opt.lowerExtraTargetElements
    If plannedElements > cap Then plannedElements = cap
    Fix9LowerPlannedElements = plannedElements
End Function
' Measured growth remains the admission model, normalized to the new nine-trial
' root search. The calibrated floor includes successful HSD work at the NEW
' lower allowance (1200 by default), so old budget-censored timings cannot
' underfund a newly completable solve. This is a cost estimate, never a bound.
Private Function Fix9LowerPredictedSeconds(ByVal plannedElements As Long, ByRef opt As Fix9ThinOptions) As Double
    Dim measuredSeconds As Double, trialCount As Double, predictedHsd As Double, predictedDll As Double
    Dim hsdAllowance As Double, calibratedSeconds As Double, coneAllowance As Double
    trialCount = 9#
    If opt.lowerCostSeconds <= 0# Or opt.lowerCostElements <= 0 Or plannedElements <= 0 Then Exit Function
    measuredSeconds = opt.lowerCostSeconds * (CDbl(plannedElements) / CDbl(opt.lowerCostElements)) ^ 1.3
    If opt.lowerCostFsCalls > 0 Then measuredSeconds = measuredSeconds * trialCount / CDbl(opt.lowerCostFsCalls)
    predictedHsd = 6.4 * CDbl(plannedElements) ^ 0.35 * (0.0045 * CDbl(plannedElements))
    coneAllowance = opt.hsdFactor * (opt.hsdBase + opt.hsdPerCone * (6# * plannedElements + 1#))
    If coneAllowance > opt.hsdCap Then coneAllowance = opt.hsdCap
    hsdAllowance = opt.lowerHsdFloorSeconds
    If coneAllowance > hsdAllowance Then hsdAllowance = coneAllowance
    If predictedHsd > hsdAllowance Then predictedHsd = hsdAllowance
    ' The measured branch already includes this premium; scale only the fit.
    predictedDll = opt.dllEquTimeFactor * 5.38380709681694E-04 * CDbl(plannedElements) ^ 1.39522461115949
    calibratedSeconds = trialCount * (predictedDll + 0.38 * predictedHsd)
    If calibratedSeconds > measuredSeconds Then measuredSeconds = calibratedSeconds
    Fix9LowerPredictedSeconds = measuredSeconds
End Function
' Fit the measured lower-round cost model and new HSD floor to the real clock.
' The first funded extra aims at 2000 elements instead of an unchecked theta-.9
' jump. Up to three default extras retain the target ceiling and lower cap.
Private Function Fix9LowerExtraLimit(ByRef lineage As RPX_MeshState, ByVal cap As Long, ByRef opt As Fix9ThinOptions, ByRef nextCap As Long) As Boolean
    Dim remainingSeconds As Double, remainingSource As String, affordableElements As Double, predictedSeconds As Double, planningReserveSeconds As Double
    If opt.lowerCostSeconds <= 0# Or opt.lowerCostElements <= 0 Then Exit Function
    remainingSeconds = Fix9DeadlineRemainingSeconds(opt.nominalSeconds, remainingSource)
    planningReserveSeconds = opt.lowerExtraReserve
    ' Preserve geometry headroom instead of spending it on extra marked edges.
    If opt.lowerProtectFunded And opt.lowerExtraPending Then planningReserveSeconds = planningReserveSeconds + Fix9UpperProtectMargin
    If remainingSeconds <= planningReserveSeconds Then Exit Function
    nextCap = Fix9LowerPlannedElements(lineage.elements, cap, opt)
    predictedSeconds = Fix9LowerPredictedSeconds(nextCap, opt)
    ' Keep the measured cost/growth admission model. Re-evaluate its new
    ' budget-aware floor at each reduced size instead of inverting the old floor.
    Do While nextCap > lineage.elements + 1 And predictedSeconds + planningReserveSeconds > remainingSeconds
        affordableElements = CDbl(nextCap) * ((remainingSeconds - planningReserveSeconds) / predictedSeconds) ^ (1# / 1.3)
        If affordableElements >= nextCap Then affordableElements = nextCap - 1
        nextCap = CLng(Fix(affordableElements))
        If nextCap > lineage.elements + 1 Then predictedSeconds = Fix9LowerPredictedSeconds(nextCap, opt)
    Loop
    Fix9LowerExtraLimit = (nextCap > lineage.elements + 1)
    RPX_DiagEvent "thin_layer_lower_extra_plan", "lineage_elements=" & lineage.elements & ";candidate_cap=" & nextCap & ";last_seconds=" & opt.lowerCostSeconds & ";remaining_seconds=" & remainingSeconds & ";remaining_source=" & remainingSource & ";allowed=" & Abs(CLng(Fix9LowerExtraLimit))
    If opt.lowerProtectFunded And opt.lowerExtraPending Then RPX_DiagEvent "thin_layer_upper_protect", "phase=lower_plan;planning_reserve_seconds=" & planningReserveSeconds & ";margin_seconds=" & Fix9UpperProtectMargin & ";candidate_cap=" & nextCap
End Function
' Returns True only when at least one thin-layer candidate reached endpoint evaluation.
Private Function Fix8ThinLayerChains(ByVal mode As String, ByVal targetGap As Double) As Boolean
    Dim mask() As Boolean, upLine As RPX_MeshState, loLine As RPX_MeshState
    Dim uRounds As Long, lRounds As Long, lStart As Long, uCap As Long, lCap As Long, passes As Long, roundIndex As Long, r As Long
    Dim uTheta As Double, lTheta As Double, upAlive As Boolean, loAlive As Boolean, upMid As Boolean, loMid As Boolean
    Dim upStalls As Long, loStalls As Long, beforeValue As Double, steps As Long, evaluatedSteps As Long
    Dim opt As Fix9ThinOptions, settings As String, stopReason As String
    Dim upExtraUsed As Boolean, loExtraUsed As Boolean, evaluated As Boolean, anyEvaluated As Boolean
    Dim lowerExtraDone As Long, lowerRoundIndex As Long, lowerStepCap As Long, lowerStarted As Double, lowerReady As Boolean
    If mode <> "Fs" Or Not currentPairValid Or Not RBestLower.complete Or Not RBestUpper.complete Then Exit Function
    If Not Fix8ThinMode(mode, mask) Then Exit Function
    uRounds = CLng(Fix8Number("ADAPT_THIN_UPPER_ROUNDS", 10, 0, 20))
    lRounds = CLng(Fix8Number("ADAPT_THIN_LOWER_ROUNDS", 3, 0, 10))
    lStart = CLng(Fix8Number("ADAPT_THIN_LOWER_START", 2, 1, 20))
    uTheta = Fix8Number("ADAPT_THIN_UPPER_THETA", 0.8, 0.05, 1)
    lTheta = Fix8Number("ADAPT_THIN_LOWER_THETA", 0.9, 0.05, 1)
    uCap = CLng(Fix8Number("ADAPT_THIN_UPPER_CAP", 3500, 16, 50000))
    lCap = CLng(Fix8Number("ADAPT_THIN_LOWER_CAP", 6000, 16, 50000))
    If fix9ThinEarlyStartTaken Then
        uTheta = Fix8Number("ADAPT_THIN_EARLY_UPPER_THETA", 0.95, 0.05, 1)
        uCap = CLng(Fix8Number("ADAPT_THIN_EARLY_UPPER_CAP", 7000, 16, 50000))
    End If
    passes = CLng(Fix8Number("ADAPT_THIN_FLIP_PASSES", 5, 0, 5))
    Call Fix9ReadOptions(opt, settings)
    CloneState upLine, RBestUpper: CloneState loLine, RBestLower
    upAlive = (uRounds > 0): loAlive = (lRounds > 0): upMid = True: loMid = True
    RPX_DiagEvent "thin_layer_begin", "upper_rounds=" & uRounds & ";upper_theta=" & uTheta & ";upper_cap=" & uCap & ";lower_rounds=" & lRounds & ";lower_start=" & lStart & ";lower_theta=" & lTheta & ";lower_cap=" & lCap & ";flip_passes=" & passes & ";lower=" & RBestLower.value & ";upper=" & RBestUpper.value & settings & ";early_overrides=" & Abs(CLng(fix9ThinEarlyStartTaken)) & ";mid_counts_as_round=0;retries_count_as_steps=1"
    For roundIndex = 1 To uRounds + lStart + lRounds + opt.lowerExtraRounds
        If XCancel Then err.Raise 18, , "ANALYSIS_CANCELLED"
        If RPX_TargetMet(mode, RBestLower.value, RBestUpper.value, currentPairValid) Then Exit For
        If Not upAlive And Not loAlive Then Exit For
        If OptionalStageExpired("thin_layer_" & roundIndex) Then Exit For
        opt.lowerExtraPending = (loAlive And lowerExtraDone < opt.lowerExtraRounds)
        opt.lowerProtectElements = loLine.elements: opt.lowerProtectCap = lCap
        If upAlive And roundIndex <= uRounds Then
            If upMid And opt.midSeparate Then
                r = Fix8ChainStep("upper", 0, upLine, False, mask, uTheta, uCap, 0, opt, True, stopReason, evaluated): steps = steps + 1
                If evaluated Then anyEvaluated = True: evaluatedSteps = evaluatedSteps + 1
                ' A failed separate mid step falls back to the Fix8 merged mid-line.
                If stopReason = "upper_protect" Then upAlive = False
                If r = 0 Then RPX_DiagEvent "thin_layer_mid_fallback", "side=upper;reason=" & stopReason Else upMid = False
                If RPX_TargetMet(mode, RBestLower.value, RBestUpper.value, currentPairValid) Then Exit For
            End If
            If upAlive Then
            beforeValue = RBestUpper.value
            r = Fix9RetryStep("upper", roundIndex, upLine, upMid, mask, uTheta, uCap, passes, opt, upExtraUsed, stopReason, steps, evaluated, evaluatedSteps)
            If evaluated Then anyEvaluated = True
            If r = 0 Then
                upAlive = False
            Else
                upMid = False
                If RBestUpper.value < beforeValue * (1# - 0.000001) Then upStalls = 0 Else upStalls = upStalls + 1
                If upStalls >= 2 Then upAlive = False
            End If
            End If
            If RPX_TargetMet(mode, RBestLower.value, RBestUpper.value, currentPairValid) Then Exit For
        ElseIf roundIndex > uRounds Then
            upAlive = False
        End If
        lowerReady = False: opt.lowerExtraActive = False: lowerStepCap = lCap
        If loAlive And roundIndex >= lStart Then
            If roundIndex < lStart + lRounds Then
                lowerReady = True: lowerRoundIndex = roundIndex - lStart + 1
            ElseIf lowerExtraDone >= opt.lowerExtraRounds Then
                loAlive = False
            ElseIf Not upAlive Or roundIndex >= opt.lowerExtraUpperRound Then
                opt.lowerExtraActive = True
                lowerReady = Fix9LowerExtraLimit(loLine, lCap, opt, lowerStepCap)
                If lowerReady Then
                    lowerExtraDone = lowerExtraDone + 1: lowerRoundIndex = lRounds + lowerExtraDone
                Else
                    loAlive = False
                End If
            End If
        End If
        If lowerReady Then
            If loMid And opt.midSeparate Then
                ' The mid step has no theta, so it is not retried; on failure fall back to Fix8 merged mid-line.
                r = Fix8ChainStep("lower", 0, loLine, False, mask, lTheta, lCap, 0, opt, True, stopReason, evaluated): steps = steps + 1
                If evaluated Then anyEvaluated = True: evaluatedSteps = evaluatedSteps + 1
                If r = 0 Then RPX_DiagEvent "thin_layer_mid_fallback", "side=lower;reason=" & stopReason Else loMid = False
                If RPX_TargetMet(mode, RBestLower.value, RBestUpper.value, currentPairValid) Then Exit For
            End If
            If loAlive Then
            beforeValue = RBestLower.value
            lowerStarted = RPX_DiagClock()
            r = Fix9RetryStep("lower", lowerRoundIndex, loLine, loMid, mask, lTheta, lowerStepCap, 0, opt, loExtraUsed, stopReason, steps, evaluated, evaluatedSteps)
            If evaluated Then anyEvaluated = True
            If r = 0 Then
                loAlive = False
            Else
                loMid = False
                ' Refit after every successful round, including each extra and retries.
                opt.lowerCostSeconds = RPX_DiagClock() - lowerStarted: opt.lowerCostElements = loLine.elements
                opt.lowerCostFsCalls = RFsCalls
                If RBestLower.value > beforeValue * (1# + 0.000001) Then loStalls = 0 Else loStalls = loStalls + 1
                If loStalls >= 2 Then loAlive = False
            End If
            End If
        End If
    Next roundIndex
    Call RestoreFinalBounds(mode)
    RPX_DiagEvent "thin_layer_end", "steps=" & steps & ";anyEvaluated=" & Abs(CLng(anyEvaluated)) & ";evaluatedSteps=" & evaluatedSteps & ";lower=" & RBestLower.value & ";upper=" & RBestUpper.value & ";gap=" & GapRatio(RBestLower.value, RBestUpper.value) & ";lower_elements=" & RBestLower.elements & ";upper_elements=" & RBestUpper.elements
    Fix8ThinLayerChains = anyEvaluated
End Function
