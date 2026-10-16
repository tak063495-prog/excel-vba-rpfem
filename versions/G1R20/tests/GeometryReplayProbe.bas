Attribute VB_Name = "GeometryReplayProbe"
Option Explicit
Private Sub Require(ByVal condition As Boolean, ByVal reason As String)
    If Not condition Then err.Raise 5, "GeometryReplayProbe", reason
End Sub
Private Function FiniteRead(ByVal f As Integer) As Double
    Dim value As Double
    Get #f, , value
    Call Require(RPX_SchurFinite(value), "NONFINITE_SOURCE")
    FiniteRead = value
End Function
Private Sub LoadGeometry(ByVal path As String)
    Dim f As Integer, magic As String * 8, nn As Long, ne As Long, nm As Long, nr As Long, q As Long
    Dim witnessScale As Double, storedFactor As Double, unused As Double, regionAreas() As Double, totalArea As Double
    Dim e As Long, j As Long, k As Long, i As Long, id As Long, count As Long, a As Long, b As Long, kind As Long, rigid As Long
    Dim px As Double, py As Double, region As Long, gamma As Double, importedBoundary As Object, value As Variant, key As Variant, segment As Variant
    Dim originalMaterials() As RPX_Material, expectedRegions() As Double
    Dim number As Long, source As String, message As String, kindName As String
    originalMaterials = RMaterials: ReDim expectedRegions(0 To RM - 1)
    For i = 0 To DRCount - 1
        Call Require(DRegions(i).rigidId = -1, "UNSUPPORTED_RIGID_REGION")
        expectedRegions(DRegions(i).materialId) = expectedRegions(DRegions(i).materialId) + DRegions(i).area
    Next i
    f = FreeFile
    On Error GoTo Failed
    Open path For Binary Access Read As #f
    Get #f, , magic: Call Require(magic = "RP25G001", "FORMAT")
    Get #f, , nn: Get #f, , ne: Get #f, , nm: Get #f, , nr: Get #f, , q
    Call Require(nn >= 3 And nn <= 100000 And ne > 0 And ne <= 50000 And nm = RM And nr = RR And nr = 0 And q = 8, "SCOPE_OR_DIMENSIONS")
    witnessScale = FiniteRead(f)
    Call Require(witnessScale = RStressScale, "SCALE")
    rn = nn: re = ne: RSubdivisions = q: RStrengthFactor = 1#
    ReDim RXY(0 To rn - 1, 0 To 1): ReDim RTri(0 To re - 1, 0 To 2)
    ReDim RMatId(0 To re - 1): ReDim RRigidId(0 To re - 1)
    ReDim RBody0(0 To re - 1, 0 To 1): ReDim RBody1(0 To re - 1, 0 To 1)
    For i = 0 To rn - 1: For j = 0 To 1: RXY(i, j) = FiniteRead(f): Next j: Next i
    For e = 0 To re - 1
        For k = 0 To 2: Get #f, , id: Call Require(id >= 0 And id < rn, "TRIANGLE_INDEX"): RTri(e, k) = id: Next k
        Get #f, , id: Call Require(id >= 0 And id < RM, "MATERIAL_INDEX"): RMatId(e) = id
        Get #f, , id: Call Require(id = -1, "RIGID_CELL"): RRigidId(e) = id
    Next e
    For i = 0 To RM - 1
        Call Require(FiniteRead(f) = originalMaterials(i).cohesion, "MATERIAL_C")
        Call Require(FiniteRead(f) = originalMaterials(i).friction, "MATERIAL_PHI")
        Call Require(FiniteRead(f) = originalMaterials(i).dilation, "MATERIAL_PSI")
        Call Require(FiniteRead(f) = originalMaterials(i).gamma, "MATERIAL_GAMMA")
    Next i
    For e = 0 To re - 1
        For j = 0 To 1
            RBody0(e, j) = FiniteRead(f): RBody1(e, j) = FiniteRead(f)
            If j = 0 Then gamma = 0# Else gamma = -RMaterials(RMatId(e)).gamma
            Call Require(RBody0(e, j) = 0# And RBody1(e, j) = gamma, "ORIGINAL_BODY_LOAD")
        Next j
    Next e
    Set importedBoundary = RPX_Row(): Get #f, , count
    Call Require(count > 0 And count <= 3 * re, "BOUNDARY_COUNT")
    For i = 1 To count
        Get #f, , a: Get #f, , b: Get #f, , kind: Get #f, , rigid
        Call Require(a >= 0 And a < rn And b >= 0 And b < rn And a <> b And kind >= 0 And kind <= 3 And rigid = -1, "BOUNDARY_RECORD")
        key = RPX_EdgeKey(a, b): Call Require(Not importedBoundary.Exists(key), "DUPLICATE_BOUNDARY")
        Select Case kind
            Case 0: kindName = "load"
            Case 1: kindName = "fixed"
            Case 2: kindName = "roller_x"
            Case 3: kindName = "roller_y"
        End Select
        value = Array(kindName, rigid, 0#, 0#, 0#, 0#)
        For j = 2 To 5: value(j) = FiniteRead(f): Next j
        importedBoundary(key) = value
    Next i
    Call Require(Seek(f) = LOF(f) + 1, "TRAILING_BYTES")
    Close #f: f = 0
    Call RPX_FsBracketReset: Call RPX_ResetPreparedState: Call RPX_LowerCacheReset
    Set RBoundary = Nothing: Call RPX_Topology
    ReDim regionAreas(0 To RM - 1): Set RConstraints = RPX_Row()
    For e = 0 To re - 1
        px = 0#: py = 0#
        For k = 0 To 2: px = px + RXY(RTri(e, k), 0) / 3#: py = py + RXY(RTri(e, k), 1) / 3#: Next k
        region = RPX_RegionAt(px, py): Call Require(region >= 0, "OUTSIDE_ORIGINAL_REGION")
        Call Require(DRegions(region).materialId = RMatId(e), "ORIGINAL_REGION_MATERIAL")
        regionAreas(RMatId(e)) = regionAreas(RMatId(e)) + RArea(e): totalArea = totalArea + RArea(e)
    Next e
    Call Require(Abs(totalArea - RPX_UnionArea()) <= 0.00000001, "ORIGINAL_TOTAL_AREA")
    ' Benchmark soil regions partition the outline; reject overlap or changed areas.
    For i = 0 To RM - 1: Call Require(Abs(regionAreas(i) - expectedRegions(i)) <= 0.00000001, "ORIGINAL_MATERIAL_AREA"): Next i
    For id = 0 To RNEdge - 1
        a = REdges(id).a: b = REdges(id).b
        For Each key In DSegments.keys
            segment = DSegments(key)
            If RPX_OnSegment(RXY(a, 0), RXY(a, 1), DPXY(segment(0), 0), DPXY(segment(0), 1), DPXY(segment(1), 0), DPXY(segment(1), 1)) Then
                If RPX_OnSegment(RXY(b, 0), RXY(b, 1), DPXY(segment(0), 0), DPXY(segment(0), 1), DPXY(segment(1), 0), DPXY(segment(1), 1)) Then RConstraints(RPX_EdgeKey(a, b)) = Array(a, b): Exit For
            End If
        Next key
    Next id
    Call RPX_MapFaces: RPX_LoadPhase = "raw": Call RPX_TotalReferenceLoads
    Call Require(RBoundary.count = importedBoundary.count, "ORIGINAL_BOUNDARY_COUNT")
    For Each key In importedBoundary.keys
        Call Require(RBoundary.Exists(key), "ORIGINAL_BOUNDARY_EDGE")
        value = importedBoundary(key): segment = RBoundary(key)
        For j = 0 To 5: Call Require(value(j) = segment(j), "ORIGINAL_BOUNDARY_LOAD_OR_RESTRAINT"): Next j
    Next key
    Exit Sub
Failed:
    number = err.number: source = err.source: message = err.description
    On Error Resume Next
    If f <> 0 Then Close #f
    On Error GoTo 0
    err.Raise number, source, message
End Sub

' Test-only geometry export. No solver vector, Fs or root metadata is serialized.
Public Sub GeometryReplayExportCurrent(ByVal path As String)
    Dim f As Integer, magic As String * 8, i As Long, j As Long, e As Long, k As Long
    Dim value As Variant, key As Variant, edge As Variant, kind As Long
    f = FreeFile: magic = "RP25G001"
    Open path For Binary Access Write As #f
    Put #f, , magic: Put #f, , rn: Put #f, , re: Put #f, , RM: Put #f, , RR: Put #f, , RSubdivisions
    Put #f, , RStressScale
    For i = 0 To rn - 1: For j = 0 To 1: Put #f, , RXY(i, j): Next j: Next i
    For e = 0 To re - 1
        For k = 0 To 2: Put #f, , RTri(e, k): Next k
        Put #f, , RMatId(e): Put #f, , RRigidId(e)
    Next e
    For i = 0 To RM - 1
        Put #f, , RMaterials(i).cohesion: Put #f, , RMaterials(i).friction
        Put #f, , RMaterials(i).dilation: Put #f, , RMaterials(i).gamma
    Next i
    For e = 0 To re - 1: For j = 0 To 1: Put #f, , RBody0(e, j): Put #f, , RBody1(e, j): Next j: Next e
    i = RBoundary.count: Put #f, , i
    For Each key In RBoundary.keys
        edge = Split(CStr(key), ":"): i = CLng(edge(0)): j = CLng(edge(1))
        value = RBoundary(key)
        Select Case CStr(value(0))
            Case "load": kind = 0
            Case "fixed": kind = 1
            Case "roller_x": kind = 2
            Case "roller_y": kind = 3
            Case Else: err.Raise 5, , "UNSUPPORTED_BOUNDARY"
        End Select
        Put #f, , i: Put #f, , j: Put #f, , kind: k = CLng(value(1)): Put #f, , k
        For j = 2 To 5: Dim v As Double: v = CDbl(value(j)): Put #f, , v: Next j
    Next key
    Close #f
End Sub

' Geometry-only byte format deliberately contains no Fs/field/root/status payload.
Public Function GeometryReplayLoadUnit(ByVal path As String) As Variant
    Dim oldCalls As Long, oldLower As Double, oldUpper As Double, oldStatus As String
    Dim a As Double, b As Double, number As Long, message As String
    On Error GoTo Failed
    oldCalls = RFsCalls: oldLower = RLowerFs: oldUpper = RUpperFs: oldStatus = XStatus
    Call RPX_InputBegin: Call RPX_ReadInputs
    If RR <> 0 Or RPX_WaterCouplingOn() Then err.Raise 5, , "GEOMETRY_REPLAY_REQUIRES_DRY_SOIL"
    Call LoadGeometry(path)
    Call RPX_AssertInputs("geometry_replay_loaded")
    Call Require(RFsCalls = oldCalls And RLowerFs = oldLower And RUpperFs = oldUpper And XStatus = oldStatus, "GEOMETRY_IMPORTED_SOLVE_STATE")
    Call Require(Not RPX_ActualRootBracket("lower", a, b) And Not RPX_ActualRootBracket("upper", a, b), "GEOMETRY_CREATED_ROOT")
    GeometryReplayLoadUnit = Array("PASS", rn, re, RM, RSubdivisions, RPX_LoadPhase, RPX_AnalysisPolicy(), "no_field_Fs_root_status_payload", "solver0")
Cleanup:
    Call RPX_InputEnd
    Exit Function
Failed:
    number = err.Number: message = err.Description
    GeometryReplayLoadUnit = Array("FAIL", number, message)
    Resume Cleanup
End Function
Public Function GeometryReplayFixedPoint(ByVal path As String, ByVal side As String, ByVal factor As Double, ByVal prefix As String) As Variant
    Dim original As RPX_AccuracyProgram, field() As Double, candidate() As Double, objective As Double
    Dim measured As Double, audit As Double, raw() As Double, a As Double, b As Double, accepted As Boolean, guided As Boolean
    Dim number As Long, message As String, f As Integer, e As Long, j As Long, state As RPX_MeshState
    On Error GoTo Failed
    Call Require((side = "lower" Or side = "upper") And RPX_SchurFinite(factor) And factor > 0#, "FIXED_POINT_ARGUMENT")
    Call RPX_InputBegin: Call RPX_ReadInputs
    Call Require(RR = 0 And Not RPX_WaterCouplingOn(), "FIXED_POINT_SCOPE")
    Call LoadGeometry(path)
    XLogPath = prefix & ".trace.tsv": Call RPX_DiagStart("g20_geometry_replay_fixed_point")
    RFsActive = True: RStrengthFactor = factor: RPX_OrderMode = side
    Call RPX_CapacityAssembly(side)
    If side = "upper" Then Call RPX_AssembleUpper Else Call RPX_AssembleLower
    Call RPX_ExportAssembled(prefix & "_original_assembled.bin")
    Call RPX_ResetPreparedState: Call RPX_Compile
    Call Require(RPX_AccuracyExport(factor, side, original), "ORIGINAL_EXPORT_FAILED")
    Call RPX_ExportExchangeContext(prefix & "_prepared_context.tsv")
    accepted = RPX_AccuracyDirectBound(factor, side)
    guided = accepted And RPX_AccuracyGuideCurrent(factor, side)
    If accepted And Not guided Then field = xx
    If Not accepted Then
        If RPX_AccuracyRecoverPrimal(factor, side, candidate) Then
            field = candidate: accepted = True
        End If
    End If
    If accepted And Not guided Then
        accepted = RPX_AccuracyPrimalAudit(original, field, objective)
        If accepted Then accepted = RPX_FieldGate(side, field, factor, 8, objective, measured, audit, raw)
        If accepted Then
            If side = "upper" Then accepted = (measured <= 0.999999) Else accepted = (measured >= 1.000001)
        End If
        If accepted Then
            Call RPX_AccuracyWriteProof(prefix & "_accepted_original_primal.bin", original, field)
            state.nodes = rn: state.elements = re: state.materialN = RM: state.rigidN = RR
            state.xy = RXY: state.triangles = RTri: state.materialId = RMatId: state.rigidId = RRigidId
            state.materials = RMaterials: state.body0 = RBody0: state.body1 = RBody1: Set state.boundary = RBoundary
            state.field = field: state.factor = factor: state.value = factor: state.q = 8: state.objective = objective
            state.stressScale = RStressScale: state.normalizedAudit = audit: state.complete = False
            Call AccuracyExport.BatchState(state, prefix & "_native_primal.bin")
        End If
    End If
    Call Require(Not RPX_ActualRootBracket(side, a, b, factor), "FIXED_POINT_CREATED_ROOT")
    f = FreeFile: Open prefix & "_result.tsv" For Output As #f
    Print #f, "geometry_replay=True;original_primal_imported=False;root_imported=False;root_complete=False"
    Print #f, "side=" & side & ";Fs_exact=" & RPX_ExactDouble(factor) & ";q=" & RSubdivisions
    Print #f, "native_primal_accepted=" & (accepted And Not guided) & ";original_dual_guide=" & guided & ";raw_eq=" & AccuracyNativeEquality & ";raw_cone=" & AccuracyNativeCone
    Print #f, "native_dims=" & original.n & "," & original.m & "," & original.nnz & ";equality=" & original.equalityCount & ";cones=" & original.coneCount
    Print #f, "objective_exact=" & RPX_ExactDouble(objective) & ";measured_exact=" & RPX_ExactDouble(measured)
    Close #f: f = 0
    GeometryReplayFixedPoint = Array("AUDITED_POINT", accepted, guided, factor, measured, objective, original.n, original.m, original.nnz, original.equalityCount, original.coneCount, re, rn, XStatus, AccuracyNativeReason, "root_not_claimed")
Cleanup:
    RFsActive = False: Call RPX_DiagFinish("DONE", "g20_geometry_replay")
    Call RPX_InputEnd
    Exit Function
Failed:
    number = err.Number: message = err.Description
    On Error Resume Next: If f <> 0 Then Close #f
    On Error GoTo 0
    GeometryReplayFixedPoint = Array("FAIL", number, message)
    Resume Cleanup
End Function
