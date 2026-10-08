Option Explicit

' R2 certificate. Design record for one bound. Not a second solver.
Public Type RPX_BoundCertificate
    valid As Boolean
    side As String
    proofKind As String
    solverStatus As String
    acceptedFs As Double
    objective As Double
    q As Long
    elements As Long
    nodes As Long
    rootBracketValid As Boolean
    rootA As Double
    rootB As Double
    stressScale As Double
    note As String
    runId As Long
    mode As String
    policy As String
    inputEpoch As Long
    normalizedAudit As Double
    rawYield As Double
    rawEquilibrium As Double
    rawTraction As Double
    rawBoundary As Double
    ownerKind As String
End Type
Public RPX_CertLower As RPX_BoundCertificate
Public RPX_CertUpper As RPX_BoundCertificate
Public RPX_CertRunId As Long
Private holdUpper As RPX_BoundCertificate
Private lowerOwner As RPX_WitnessPayload, upperOwner As RPX_WitnessPayload, holdOwner As RPX_WitnessPayload
Private holdReady As Boolean

Public Sub RPX_CertClear()
    RPX_CertLower.valid = False
    RPX_CertUpper.valid = False
    holdReady = False
    Set lowerOwner = Nothing: Set upperOwner = Nothing: Set holdOwner = Nothing
End Sub
Public Sub RPX_CertBeginRun()
    RPX_CertRunId = RPX_CertRunId + 1
    RPX_CertLower.valid = False
    RPX_CertUpper.valid = False
    holdReady = False
    Set lowerOwner = Nothing: Set upperOwner = Nothing: Set holdOwner = Nothing
    RPX_DiagEvent "cert_run", "id=" & CStr(RPX_CertRunId) & ";cleared=True"
End Sub
Public Function RPX_CertCurrent(ByVal side As String) As Boolean
    If RPX_CertRunId <= 0 Then Exit Function
    If side = "lower" Then
        RPX_CertCurrent = RPX_CertLower.valid And RPX_CertLower.runId = RPX_CertRunId And RPX_CertLower.inputEpoch = RPX_InputEpoch And RPX_CertLower.policy = RPX_AnalysisPolicy()
    ElseIf side = "upper" Then
        RPX_CertCurrent = RPX_CertUpper.valid And RPX_CertUpper.runId = RPX_CertRunId And RPX_CertUpper.inputEpoch = RPX_InputEpoch And RPX_CertUpper.policy = RPX_AnalysisPolicy()
    End If
End Function

Public Sub RPX_CertCapture(ByVal side As String, ByVal proofKind As String, ByVal q As Long, ByVal rootValid As Boolean, ByVal note As String)
    Dim pending As RPX_WitnessPayload, record As RPX_BoundCertificate, other As Double
    Dim measuredRaw() As Double, measuredValue As Double, measuredAudit As Double, a As Double, b As Double
    RPX_AssertInputs "certificate_prepare"
    If side <> "lower" And side <> "upper" Then err.Raise 5, "RPX_Certificate", "CERT_SIDE"
    If rn < 1 Or re < 1 Then err.Raise 5, "RPX_Certificate", "CERT_EMPTY_MESH"
    record.side = side: record.proofKind = proofKind: record.solverStatus = XStatus
    record.acceptedFs = RLowerFs: If side = "upper" Then record.acceptedFs = RUpperFs
    If side = "lower" Then
        If Not RPX_FieldGate(side, RLowerField, record.acceptedFs, q, RValue, measuredValue, measuredAudit, measuredRaw) Then Exit Sub
        If proofKind = "FIXED_FS_WITNESS" Then
            If Abs(measuredValue - 1#) > 0.0000001 Then Exit Sub
        End If
    Else
        If Not RPX_FieldGate(side, RUpperField, record.acceptedFs, q, RValue, measuredValue, measuredAudit, measuredRaw) Then Exit Sub
        If proofKind = "TARGET_WITNESS" Then
            If measuredValue >= 1# Then Exit Sub
        End If
    End If
    If side = "lower" Then
        If RPX_CertCurrent("lower") Then
            If record.acceptedFs <= RPX_CertLower.acceptedFs Then Exit Sub
        End If
        If RPX_CertCurrent("upper") Then
            If record.acceptedFs > RPX_CertUpper.acceptedFs Then err.Raise 5, "RPX_Certificate", "CERT_BOUNDS_REVERSED"
        End If
    Else
        If RPX_CertCurrent("upper") Then
            If record.acceptedFs >= RPX_CertUpper.acceptedFs Then Exit Sub
        End If
        If RPX_CertCurrent("lower") Then
            If record.acceptedFs < RPX_CertLower.acceptedFs Then err.Raise 5, "RPX_Certificate", "CERT_BOUNDS_REVERSED"
        End If
    End If
    record.objective = measuredValue: record.q = q: record.elements = re: record.nodes = rn
    If proofKind = "TARGET_WITNESS" Or proofKind = "FIXED_FS_WITNESS" Then rootValid = False
    If rootValid Then rootValid = RPX_ActualRootBracket(side, a, b, record.acceptedFs)
    record.rootBracketValid = rootValid
    record.rootA = a: record.rootB = b
    record.stressScale = RStressScale: record.note = note: record.runId = RPX_CertRunId
    record.mode = CStr(RPX_Setting("MODE", "")): record.policy = RPX_AnalysisPolicy()
    record.inputEpoch = RPX_InputEpoch: record.normalizedAudit = measuredAudit
    record.rawYield = measuredRaw(1): record.rawEquilibrium = measuredRaw(2)
    record.rawTraction = measuredRaw(3): record.rawBoundary = measuredRaw(4)
    record.ownerKind = "IMMUTABLE_PAYLOAD"
    Set pending = New RPX_WitnessPayload
    If side = "lower" Then pending.Capture RLowerField, record.acceptedFs, q Else pending.Capture RUpperField, record.acceptedFs, q
    RPX_AssertInputs "certificate_commit"
    record.valid = True
    ' All allocating work is complete. Set publishes one complete payload.
    If side = "lower" Then
        RPX_CertLower = record: Set lowerOwner = pending
    Else
        RPX_CertUpper = record: Set upperOwner = pending
    End If
    RPX_DiagEvent "certificate_commit", "side=" & side & ";proof=" & proofKind & ";fs=" & record.acceptedFs & ";policy=" & record.policy & ";root_valid=" & record.rootBracketValid
End Sub
Public Sub RPX_CertFromAdapt(ByVal side As String)
    Dim record As RPX_BoundCertificate
    RPX_AssertInputs "adapt_certificate"
    record.side = side: record.mode = CStr(RPX_Setting("MODE", ""))
    record.ownerKind = "ADAPT_BEST_STATE": record.policy = RPX_AnalysisPolicy(): record.inputEpoch = RPX_InputEpoch
    record.runId = RPX_CertRunId: record.proofKind = "ROOT_ENDPOINT": record.solverStatus = "AUDITED_BEST_FIELD"
    record.stressScale = RStressScale: record.q = RSubdivisions
    If side = "lower" Then
        record.acceptedFs = RBestLower.value: record.nodes = RBestLower.nodes: record.elements = RBestLower.elements
        record.rootBracketValid = RBestLower.lowerRootValid And Not RBestLower.searchIncomplete: record.note = RBestLower.searchNote
        record.rootA = RBestLower.lowerRootA: record.rootB = RBestLower.lowerRootB
        record.q = RBestLower.q: record.stressScale = RBestLower.stressScale
        record.objective = RBestLower.objective: record.normalizedAudit = RBestLower.normalizedAudit
        record.rawYield = RBestLower.rawYield: record.rawEquilibrium = RBestLower.rawEquilibrium
        record.rawTraction = RBestLower.rawTraction: record.rawBoundary = RBestLower.rawBoundary
    ElseIf side = "upper" Then
        record.acceptedFs = RBestUpper.value: record.nodes = RBestUpper.nodes: record.elements = RBestUpper.elements
        record.rootBracketValid = RBestUpper.upperRootValid And Not RBestUpper.searchIncomplete: record.note = RBestUpper.searchNote
        record.rootA = RBestUpper.upperRootA: record.rootB = RBestUpper.upperRootB
        record.q = RBestUpper.q: record.stressScale = RBestUpper.stressScale
        record.objective = RBestUpper.objective: record.normalizedAudit = RBestUpper.normalizedAudit
    Else
        err.Raise 5, "RPX_Certificate", "CERT_SIDE"
    End If
    record.valid = True
    If side = "lower" Then RPX_CertLower = record Else RPX_CertUpper = record
End Sub

Public Sub RPX_CertHoldUpper()
    If Not RPX_CertCurrent("upper") Or upperOwner Is Nothing Then err.Raise 5, "RPX_Certificate", "CERT_UPPER_MISSING"
    holdUpper = RPX_CertUpper: Set holdOwner = upperOwner: holdReady = True
End Sub
Public Sub RPX_CertRestoreHeldUpper()
    If Not holdReady Or holdOwner Is Nothing Then err.Raise 5, "RPX_Certificate", "CERT_HOLD_MISSING"
    holdOwner.RestoreUpper
    RUpperFs = holdUpper.acceptedFs: RPX_C0UpperElements = holdUpper.elements
    Set upperOwner = holdOwner: RPX_CertUpper = holdUpper
    RPX_AuditRejected = False
    RPX_DiagEvent "certificate_restore", "load_phase=" & RPX_LoadPhase & ";source=immutable_payload;exact_model=True"
End Sub

Public Sub RPX_CertWriteLower()
    Dim ws As Worksheet, data() As Variant, labels() As Variant
    Dim e As Long, k As Long, offset As Long, n As Long
    If Not RPX_CertCurrent("lower") Then Exit Sub
    If RPX_CertLower.ownerKind <> "IMMUTABLE_PAYLOAD" Then Exit Sub
    If lowerOwner Is Nothing Then err.Raise 5, "RPX_Certificate", "CERT_OWNER_MISSING"
    n = RPX_CertLower.nodes
    Set ws = RPX_Sheet("節点データ")
    ws.Range("F1:H1").Value2 = Array("下界節点ID", "X [m]", "Y [m]")
    ws.Range("F2:H4000").ClearContents
    ReDim data(1 To n, 1 To 3)
    For e = 0 To n - 1
        data(e + 1, 1) = e + 1
        data(e + 1, 2) = lowerOwner.node(e, 0)
        data(e + 1, 3) = lowerOwner.node(e, 1)
    Next e
    ws.Range("F2").Resize(n, 3).Value2 = data
    Set ws = RPX_Sheet("要素データ")
    ws.Range("AA1:AF1").Value2 = Array("下界要素ID", "節点1", "節点2", "節点3", "材料内部番号", "剛体内部番号")
    ws.Range("AA2:AF4000").ClearContents
    ws.Range("AH2:AY4000").ClearContents
    ReDim data(1 To RPX_CertLower.elements, 1 To 6)
    For e = 0 To RPX_CertLower.elements - 1
        data(e + 1, 1) = e + 1
        For k = 0 To 2: data(e + 1, k + 2) = lowerOwner.Triangle(e, k) + 1: Next k
        data(e + 1, 5) = lowerOwner.Material(e) + 1
        data(e + 1, 6) = lowerOwner.Rigid(e) + 1
    Next e
    ws.Range("AA2").Resize(RPX_CertLower.elements, 6).Value2 = data
    ReDim labels(1 To 1, 1 To 18)
    ReDim data(1 To RPX_CertLower.elements, 1 To 18)
    For k = 0 To 5
        labels(1, 3 * k + 1) = "Sxx" & k + 1 & " [kPa]"
        labels(1, 3 * k + 2) = "Syy" & k + 1 & " [kPa]"
        labels(1, 3 * k + 3) = "Txy" & k + 1 & " [kPa]"
    Next k
    offset = 0
    For e = 0 To RPX_CertLower.elements - 1
        If lowerOwner.Rigid(e) < 0 Then
            For k = 0 To 17: data(e + 1, k + 1) = lowerOwner.field(offset + k) * RPX_CertLower.stressScale: Next k
            offset = offset + 18
        End If
    Next e
    ws.Range("AH1").Resize(1, 18).Value2 = labels
    ws.Range("AH2").Resize(RPX_CertLower.elements, 18).Value2 = data
    ws.Range("BA1").Value2 = "下界応力はこの下界メッシュの場。上界速度H:Sは上界メッシュ。行数は一致しない。"
End Sub