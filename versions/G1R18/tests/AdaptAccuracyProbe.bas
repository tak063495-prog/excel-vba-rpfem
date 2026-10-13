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
Public Function AccuracyProbeRun(ByVal cap As Long, ByVal rounds As Long, ByVal trace As String, Optional ByVal injectFailure As Boolean = False, Optional ByVal lowerFan As Boolean = False, Optional ByVal fanFixture As Boolean = False, Optional ByVal certQ As Long = 0, Optional ByVal qFailure As Boolean = False) As String
    Dim measured As Double, audit As Double, raw() As Double
    Call RPX_Setup
    Call SetProbeSetting("TARGET", 64)
    Call SetProbeSetting("ADAPT", 1)
    Call SetProbeSetting("C0_AUTO", 0)
    Call SetProbeSetting("CYCLES", 0)
    Call SetProbeSetting("QUADRATURE", 2)
    Call SetProbeSetting("GAP", 0.000000000001)
    Call SetProbeSetting("ADAPT_MESH_POLICY", "ROBUST_REFINE")
    Call SetProbeSetting("ADAPT_ACCURACY_CAP", cap)
    Call SetProbeSetting("ADAPT_ACCURACY_ROUNDS", rounds)
    Call SetProbeSetting("ADAPT_LOWER_FAN_ROUNDS", IIf(lowerFan, 1, 0))
    Call SetProbeSetting("ADAPT_LOWER_TARGET", 0)
    If lowerFan Then Call SetProbeSetting("ADAPT_LOWER_TARGET", 128)
    Call SetProbeSetting("ADAPT_CERT_QUADRATURE", certQ)
    Call SetProbeSetting("HSD_SOLVE_SECONDS", 30)
    If certQ > 0 Then
        Call SetProbeSetting("TARGET", 16)
        Call SetProbeSetting("HSD_SOLVE_SECONDS", 180)
    End If
    If injectFailure Then
        Call SetProbeSetting("TEST_INJECTION", 1)
        Call RPX_TestArm("accuracy_1", 5, "CONIC_NOT_CONVERGED")
    End If
    If certQ > 0 Then
        RPX_Sheet("材料データ").Range("C5").Value2 = 0#
        RPX_Sheet("材料データ").Range("E5").Value2 = 0#
    End If
    If qFailure Then
        Call SetProbeSetting("TEST_INJECTION", 1)
        Call RPX_TestArm("own_quadrature_lower", 5, "CONIC_NOT_CONVERGED")
    End If
    RPX_Sheet("浸透入力").Range("D4").Value2 = 0
    If fanFixture Then
        ' A small c0 free-boundary region in the ordinary cohesive slope.
        ' A pure c0 coarse upper search can remain UNKNOWN before the fan runs.
        RPX_Sheet("要素定義").Range("A13:C13").Value2 = Array(7, 6#, 15#)
        RPX_Sheet("要素定義").Range("A14:C14").Value2 = Array(8, 6#, 10#)
        RPX_Sheet("要素定義").Range("H7").Value2 = "1,2,7,3,4,5,6"
        RPX_Sheet("要素定義").Range("E8:H8").Value2 = Array(2, 2, 0, "2,7,8")
        RPX_Sheet("材料データ").Range("A6:E6").Value2 = Array(2, 0#, 30#, 18.82, 30#)
    End If
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
    AccuracyProbeRun = "PASS;lower=" & RBestLower.value & ";upper=" & RBestUpper.value & ";lower_elements=" & RBestLower.elements & ";upper_elements=" & RBestUpper.elements & ";lower_q=" & RBestLower.q & ";upper_q=" & RBestUpper.q & ";lower_root=" & RBestLower.lowerRootValid & ";upper_root=" & RBestUpper.upperRootValid & ";current_elements=" & re & ";incomplete=" & RFsSearchIncomplete & ";physical_gates=PASS;upper_identity=" & RBestUpper.originalModel
End Function

Public Function AccuracyQGuardProbe() As String
    Call RPX_Setup: Call RPX_ReadInputs
    Call SetProbeSetting("ADAPT_CERT_QUADRATURE", 8)
    Call RPX_CertBeginRun
    RPX_CertLower.valid = True: RPX_CertUpper.valid = True
    RPX_CertLower.runId = RPX_CertRunId: RPX_CertUpper.runId = RPX_CertRunId
    RPX_CertLower.inputEpoch = RPX_InputEpoch: RPX_CertUpper.inputEpoch = RPX_InputEpoch
    RPX_CertLower.policy = RPX_AnalysisPolicy(): RPX_CertUpper.policy = RPX_AnalysisPolicy()
    RPX_CertLower.rootBracketValid = True: RPX_CertUpper.rootBracketValid = True
    RPX_CertLower.q = 4: RPX_CertUpper.q = 4
    RFsSearchIncomplete = False: RPX_AuditRejected = False: XCancel = False
    If RPX_TargetMet("Fs", 1#, 1.005, True) Then err.Raise 5, "AdaptAccuracyProbe", "Q4_ROOT_PROMOTED"
    RPX_CertLower.q = 8
    If RPX_TargetMet("Fs", 1#, 1.005, True) Then err.Raise 5, "AdaptAccuracyProbe", "MIXED_Q_ROOT_PROMOTED"
    RPX_CertUpper.q = 8: RPX_CertLower.rootBracketValid = False
    If RPX_TargetMet("Fs", 1#, 1.005, True) Then err.Raise 5, "AdaptAccuracyProbe", "INCOMPLETE_Q8_ROOT_PROMOTED"
    RPX_CertLower.rootBracketValid = True
    If Not RPX_TargetMet("Fs", 1#, 1.005, True) Then err.Raise 5, "AdaptAccuracyProbe", "CURRENT_Q8_CERT_REJECTED"
    RPX_CertLower.inputEpoch = RPX_InputEpoch - 1
    If RPX_TargetMet("Fs", 1#, 1.005, True) Then err.Raise 5, "AdaptAccuracyProbe", "STALE_Q8_CERT_PROMOTED"
    Call RPX_CertClear: Call SetProbeSetting("ADAPT_CERT_QUADRATURE", 0)
    AccuracyQGuardProbe = "PASS;q4_rejected=1;mixed_q_rejected=1;root_false_rejected=1;stale_epoch_rejected=1;current_q8_accepted=1"
End Function

Public Function AccuracyProductionRun(ByVal trace As String, ByVal prefix As String) As String
    Dim measured As Double, audit As Double, raw() As Double, target As Boolean
    XLogPath = trace: Call RPX_NativeSetMode("auto")
    Call RPX_Run
    If Not RBestLower.complete Or Not RBestUpper.complete Then
        AccuracyProductionRun = "PARTIAL;physical_gates=UNAVAILABLE;error=" & RPX_LastError
        Exit Function
    End If
    target = RPX_TargetMet("Fs", RBestLower.value, RBestUpper.value, RPX_AdaptPairComplete())
    Call RPX_RestoreWitnessState(RBestLower)
    If Not RPX_FieldGate("lower", RBestLower.field, RBestLower.value, RBestLower.q, RBestLower.objective, measured, audit, raw) Then
        AccuracyProductionRun = "FAIL;physical_gates=LOWER_FAILED": Exit Function
    End If
    Call RPX_RestoreWitnessState(RBestUpper)
    If Not RPX_FieldGate("upper", RBestUpper.field, RBestUpper.value, RBestUpper.q, RBestUpper.objective, measured, audit, raw) Then
        AccuracyProductionRun = "FAIL;physical_gates=UPPER_FAILED": Exit Function
    End If
    ' The export module is test-only and writes the independently saved states.
    Call AccuracyExport.BatchState(RBestLower, prefix & "_lower.bin")
    Call AccuracyExport.BatchState(RBestUpper, prefix & "_upper.bin")
    AccuracyProductionRun = "AUDITED;target_met=" & target & ";lower=" & RBestLower.value & ";upper=" & RBestUpper.value & ";lower_elements=" & RBestLower.elements & ";upper_elements=" & RBestUpper.elements & ";lower_q=" & RBestLower.q & ";upper_q=" & RBestUpper.q & ";lower_root=" & RBestLower.lowerRootValid & ";upper_root=" & RBestUpper.upperRootValid & ";lower_root_a=" & RBestLower.lowerRootA & ";lower_root_b=" & RBestLower.lowerRootB & ";upper_root_a=" & RBestUpper.upperRootA & ";upper_root_b=" & RBestUpper.upperRootB & ";incomplete=" & RFsSearchIncomplete & ";physical_gates=PASS;error=" & RPX_LastError
End Function

Public Function AccuracyFanGeometryProbe() As String
    Dim parent() As Long, oldTri() As Long, oldMat() As Long, oldRigid() As Long
    Dim oldBody0() As Double, oldBody1() As Double, oldArea() As Double, area() As Double
    Dim edges As Object, found As Object, key As Variant, rec As Variant
    Dim e As Long, k As Long, n As Long, identity As String, beforeBoundary As String
    Call RPX_Setup: Call RPX_ReadInputs
    rn = 5: re = 4: RM = 2: RR = 0
    ReDim RXY(0 To 4, 0 To 1): ReDim RTri(0 To 3, 0 To 2)
    RXY(1, 0) = 1#: RXY(2, 0) = 1#: RXY(2, 1) = 1#: RXY(3, 1) = 1#: RXY(4, 0) = 0.5: RXY(4, 1) = 0.5
    For e = 0 To 3: RTri(e, 0) = e: RTri(e, 1) = (e + 1) Mod 4: RTri(e, 2) = 4: Next e
    ReDim RMaterials(0 To 1): ReDim RRigids(0 To 0)
    RMaterials(0).cohesion = 0#: RMaterials(0).friction = 30#: RMaterials(0).gamma = 18#
    RMaterials(1).cohesion = 7#: RMaterials(1).friction = 24#: RMaterials(1).gamma = 20#
    ReDim RMatId(0 To 3): ReDim RRigidId(0 To 3): ReDim RBody0(0 To 3, 0 To 1): ReDim RBody1(0 To 3, 0 To 1)
    For e = 0 To 3
        If e > 0 Then RMatId(e) = 1
        RRigidId(e) = -1
        For k = 0 To 1: RBody0(e, k) = e + k + 0.125: RBody1(e, k) = -2# * e - k - 0.25: Next k
    Next e
    Set RBoundary = RPX_Row(): Set RConstraints = RPX_Row()
    RBoundary("0:1") = Array("load", -1, 0#, 0#, 0#, 0#)
    RBoundary("1:2") = Array("load", -1, 1#, 0#, 2#, 0#)
    RBoundary("2:3") = Array("velocity", -1, 0#, 0#, 0#, 0#)
    RBoundary("0:3") = Array("load", -1, 0#, -1#, 0#, -2#)
    For e = 0 To 3
        RConstraints(RPX_EdgeKey(e, (e + 1) Mod 4)) = Array(e, (e + 1) Mod 4)
    Next e
    RPX_LoadPhase = "reference_total": RStrengthFactor = 1#: RStressScale = 20#: RSubdivisions = 8
    Call RPX_Topology
    identity = RPX_ExactModel()
    If RPX_AngularFanMesh(9, parent) Then err.Raise 5, "AdaptAccuracyProbe", "FAN_CAP_NOT_ENFORCED"
    If identity <> RPX_ExactModel() Then err.Raise 5, "AdaptAccuracyProbe", "FAN_CAP_MUTATED_MODEL"
    oldTri = RTri: oldMat = RMatId: oldRigid = RRigidId: oldBody0 = RBody0: oldBody1 = RBody1: oldArea = RArea
    Set edges = RPX_Row()
    For e = 0 To RNEdge - 1: edges(RPX_EdgeKey(REdges(e).a, REdges(e).b)) = True: Next e
    If Not RPX_AngularFanMesh(10, parent) Then err.Raise 5, "AdaptAccuracyProbe", "FAN_NOT_CREATED"
    If re <> 10 Or rn <> 8 Or RPX_LoadPhase <> "reference_total" Then err.Raise 5, "AdaptAccuracyProbe", "FAN_SCOPE_OR_PHASE"
    ReDim area(0 To 3)
    For e = 0 To re - 1
        n = parent(e): area(n) = area(n) + RArea(e)
        If RMatId(e) <> oldMat(n) Or RRigidId(e) <> oldRigid(n) Then err.Raise 5, "AdaptAccuracyProbe", "FAN_MATERIAL_CHANGED"
        For k = 0 To 1
            If RBody0(e, k) <> oldBody0(n, k) Or RBody1(e, k) <> oldBody1(n, k) Then err.Raise 5, "AdaptAccuracyProbe", "FAN_BODY_CHANGED"
        Next k
    Next e
    For e = 0 To 3: If Abs(area(e) - oldArea(e)) > 0.000000000001 Then err.Raise 5, "AdaptAccuracyProbe", "FAN_PARENT_AREA_CHANGED"
    Next e
    For k = 0 To 2: If RTri(2, k) <> oldTri(2, k) Then err.Raise 5, "AdaptAccuracyProbe", "FAN_UNEXPOSED_TRIANGLE_SPLIT"
    Next k
    Set found = RPX_Row()
    For e = 0 To RNEdge - 1: found(RPX_EdgeKey(REdges(e).a, REdges(e).b)) = True: Next e
    For Each key In edges.keys: If Not found.Exists(key) Then err.Raise 5, "AdaptAccuracyProbe", "FAN_OLD_EDGE_CHANGED"
    Next key
    For e = 0 To RNEdge - 1
        With REdges(e)
            If .element(1) < 0 Then
                key = RPX_EdgeKey(.a, .b)
                If Not RBoundary.Exists(key) Or Not RConstraints.Exists(key) Then err.Raise 5, "AdaptAccuracyProbe", "FAN_HANGING_BOUNDARY"
                rec = RBoundary(key)
                If .kind <> rec(0) Or .rigidId <> rec(1) Then err.Raise 5, "AdaptAccuracyProbe", "FAN_BOUNDARY_KIND_CHANGED"
                For k = 0 To 1
                    If .load0(k) <> rec(2 + k) Or .load1(k) <> rec(4 + k) Then err.Raise 5, "AdaptAccuracyProbe", "FAN_BOUNDARY_LOAD_CHANGED"
                Next k
            End If
        End With
    Next e
    AccuracyFanGeometryProbe = "PASS;old_elements=4;split_elements=3;new_elements=10;cohesive_neighbor_split=1;unexposed_element_unchanged=1;parent_areas_materials_body_boundary_preserved=1;cap_atomic=1"
End Function
