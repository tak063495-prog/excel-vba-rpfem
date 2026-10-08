Attribute VB_Name = "RPX_Adapt"
' version=20260926_0945 engine=20260926_p2s_f01_f07
Option Explicit
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
End Type
Public RBestUpper As RPX_MeshState, RBestLower As RPX_MeshState
Public RAdaptHistory As Collection
Public RAdaptRoute As String
Private adaptCoarseSearch As Boolean
Private Const AdaptCoarseTol As Double = 0.005
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
Private adaptHintLower As Double
Private adaptHintUpper As Double
Private adaptHintUpperValid As Boolean
Private adaptHintLowerValid As Boolean
Private adaptFsCallsTotal As Long
Private adaptStatus As String
Private meshStage As String
Private RBestPair As RPX_MeshState
Private RCurrent As RPX_MeshState
Private currentPairValid As Boolean
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
    If adaptCoarseSearch Then
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
            lower = RPX_FsEndpoint("lower", adaptHintLower, AdaptFsHintRatio)
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
    ComputeShearWeight vel, w
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
    ComputeShearWeight velocity, w
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

Private Function TryShearRefine(ByRef baseline As RPX_MeshState, ByVal maxElements As Long) As Boolean
    Dim fraction As Double, score() As Double, w() As Double, vel() As Double
    Dim e As Long, total As Double, Failed As Long
    Dim failureText As String, failureSource As String, failureLine As Long
    Dim robustRefine As Boolean
    robustRefine = RPX_RobustMeshPolicy()
    fraction = AdaptBandRefineFraction
    If maxElements <= baseline.elements Then Exit Function
    Do While fraction >= 0.01
        meshStage = "refine_restore"
        RPX_RestoreState baseline
        CopyVelocity vel, baseline.velocity
        If robustRefine Then ComputeKinematicRefineWeight vel, w Else ComputeShearWeight vel, w
        total = 0#
        For e = 0 To re - 1: total = total + w(e): Next e
        ReDim score(0 To re - 1)
        If total > 0# Then
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
        RPX_RefineAttempt score, fraction
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
Public Sub RPX_RunAdapt(ByVal mode As String)
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
    Dim robustRefine As Boolean, robustExhausted As Boolean
    robustRefine = RPX_RobustMeshPolicy()
    Set RAdaptHistory = New Collection
    RBestUpper.value = 1E+300: RBestLower.value = -1E+300
    RBestPair.pairGap = 1E+300
    RBestPair.elements = 0
    currentPairValid = False
    cycles = CLng(RPX_Setting("CYCLES", 3)): targetGap = CDbl(RPX_Setting("GAP", 0.01))
    If cycles < 0 Or cycles > 20 Or targetGap <= 0# Then err.Raise 5, , "Check CYCLES and GAP settings."
    adaptFinalTarget = DTarget
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
    RPX_DiagEvent "adapt_mesh_policy", "requested=" & RPX_Setting("ADAPT_MESH_POLICY", "LEGACY") & ";effective_red_green=" & robustRefine & ";budget=" & DTarget
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
AdaptCycles:
    If route = "pilot" And Not pilotFallbackUsed Then
        On Error GoTo PilotFailed
    Else
        On Error GoTo 0
    End If
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
        Call Evaluate(mode, "adapt_" & cycle & IIf(robustRefine, "_refine", "_band"))
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
        If robustRefine Then
            If RCurrent.elements <= baseline.elements Or gain <= 0# Then robustExhausted = True: Exit For
        Else
            If gain < AdaptMinGain Then Exit For
        End If
    Next cycle
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
        If TryShearRefine(baseline, maxElements) Then
            RPX_DiagEnd dpMesh
            If route = "pilot" And Not pilotFallbackUsed Then
                On Error GoTo PilotFailed
            Else
                On Error GoTo 0
            End If
            If robustRefine Then On Error GoTo ShearRefineFailed
            If robustRefine Then meshStage = "evaluate"
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
    ' Preserve audited side witnesses while the final candidate is being evaluated.
    On Error GoTo FinalCandidateFailed
    Call Evaluate(mode, "final")
    On Error GoTo 0
FinalCandidateDone:
    If XCancel Then err.Raise 18, , "ANALYSIS_CANCELLED"
    Call RestoreFinalBounds(mode)
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
    If robustRefine And InStr(1, savedErrText, "HSD_STABLE_UNAVAILABLE_MEMORY", vbBinaryCompare) = 1 Then err.Raise savedErrNo, savedErrSource, savedErrText
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
ShearRefineFailed:
    savedErrNo = err.number: savedErrText = err.description: savedErrSource = err.source: savedErrLine = Erl
    On Error GoTo 0
    If robustRefine And InStr(1, savedErrText, "HSD_STABLE_UNAVAILABLE_MEMORY", vbBinaryCompare) = 1 Then err.Raise savedErrNo, savedErrSource, savedErrText
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
    If robustRefine And InStr(1, savedErrText, "HSD_STABLE_UNAVAILABLE_MEMORY", vbBinaryCompare) = 1 Then err.Raise savedErrNo, savedErrSource, savedErrText
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
