Attribute VB_Name = "PythonWitnessProbe"
Option Explicit

Private Sub Require(ByVal condition As Boolean, ByVal reason As String)
    If Not condition Then err.Raise 5, "PythonWitnessProbe", reason
End Sub
Private Function FiniteRead(ByVal f As Integer) As Double
    Dim value As Double
    Get #f, , value
    Call Require(RPX_SchurFinite(value), "NONFINITE_SOURCE")
    FiniteRead = value
End Function
Private Sub LoadWitness(ByVal path As String, ByRef field() As Double, ByRef factor As Double)
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
    Get #f, , magic: Call Require(magic = "RP25M001", "FORMAT")
    Get #f, , nn: Get #f, , ne: Get #f, , nm: Get #f, , nr: Get #f, , q
    Call Require(nn >= 3 And nn <= 100000 And ne > 0 And ne <= 50000 And nm = RM And nr = RR And nr = 0 And q = 8, "SCOPE_OR_DIMENSIONS")
    witnessScale = FiniteRead(f): storedFactor = FiniteRead(f): factor = FiniteRead(f): unused = FiniteRead(f): unused = FiniteRead(f)
    Call Require(witnessScale = RStressScale And storedFactor = factor And factor > 0#, "SCALE_FACTOR")
    rn = nn: re = ne: RSubdivisions = q: RStrengthFactor = factor
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
    Get #f, , count: Call Require(count > 0 And count <= 2000000, "FIELD_DIMENSION")
    ReDim field(0 To count - 1): For i = 0 To count - 1: field(i) = FiniteRead(f): Next i
    Call Require(Seek(f) = LOF(f) + 1, "TRAILING_BYTES")
    Close #f: f = 0
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

Public Function PythonRejectWitnessProbe(ByVal path As String) As String
    Dim field() As Double, factor As Double, reason As String
    On Error GoTo Rejected
    Call RPX_InputBegin: Call RPX_ReadInputs
    Call LoadWitness(path, field, factor)
    Call RPX_InputEnd
    PythonRejectWitnessProbe = "FAIL;INVALID_SOURCE_ACCEPTED": Exit Function
Rejected:
    reason = err.description: err.Clear
    Call RPX_InputEnd
    PythonRejectWitnessProbe = "REJECT;" & reason
End Function

Public Function PythonOriginalFieldProbe(ByVal path As String, ByVal side As String, ByVal proofPrefix As String) As String
    Dim sourceField() As Double, field() As Double, factor As Double, measured As Double, audit As Double, raw() As Double
    Dim i As Long, base As Long, accepted As Boolean, originalObjective As Double, equalityError As Double, coneError As Double
    Dim oldLower As Double, oldUpper As Double, oldLowerFs As Double, oldUpperFs As Double, oldStatus As String
    Dim state As RPX_MeshState, context As Object
    Call RPX_InputBegin: Call RPX_ReadInputs
    RLowerLoad = 0.123: RUpperLoad = 0.987: RLowerFs = 0.321: RUpperFs = 0.765
    XStatus = "PYTHON_IMPORT_PROBE_UNSOLVED"
    oldLower = RLowerLoad: oldUpper = RUpperLoad: oldLowerFs = RLowerFs: oldUpperFs = RUpperFs: oldStatus = XStatus
    Call RPX_FsBracketReset: Call RPX_ResetPreparedState: Call RPX_LowerCacheReset
    Call LoadWitness(path, sourceField, factor)
    Set context = CreateObject("Scripting.FileSystemObject").CreateTextFile(proofPrefix & "_stages.tsv", False, True)
    context.WriteLine "model_parity_checked": context.Close
    RFsActive = True: RPX_OrderMode = side
    If side = "upper" Then
        Call RPX_AssembleUpper
        base = 12 * re + 3 * RR: Call Require(UBound(sourceField) >= base - 1, "SOURCE_VELOCITY_DIMENSION")
        ReDim field(0 To XN - 1): For i = 0 To base - 1: field(i) = sourceField(i): Next i
        Call RPX_LiftUpperEpigraph(field)
    ElseIf side = "lower" Then
        Call RPX_AssembleLower
        Call Require(UBound(sourceField) = XN - 1, "SOURCE_STRESS_DIMENSION")
        field = sourceField
    Else
        err.Raise 5, "PythonWitnessProbe", "SIDE"
    End If
    Set context = CreateObject("Scripting.FileSystemObject").OpenTextFile(proofPrefix & "_stages.tsv", 8, False, -1)
    context.WriteLine "original_assembled": context.Close
    ReDim raw(1 To 4)
    accepted = RPX_OriginalPrimalGate(field)
    originalObjective = OriginalCandidateObjective: equalityError = OriginalEqualityError: coneError = OriginalConeError
    Set context = CreateObject("Scripting.FileSystemObject").OpenTextFile(proofPrefix & "_stages.tsv", 8, False, -1)
    context.WriteLine "original_primal_gate=" & accepted: context.Close
    If accepted Then accepted = RPX_FieldGate(side, field, factor, 8, originalObjective, measured, audit, raw)
    If side = "upper" Then
        accepted = accepted And measured <= 0.99999901
    Else
        accepted = accepted And measured >= 1.000001
    End If
    Call Require(RLowerLoad = oldLower And RUpperLoad = oldUpper And RLowerFs = oldLowerFs And RUpperFs = oldUpperFs, "IMPORT_CHANGED_BOUND")
    Call Require(XStatus = oldStatus, "IMPORT_CHANGED_STATUS")
    Call Require(Not RPX_ActualRootBracket(side, oldLower, oldUpper, factor), "IMPORT_CREATED_ROOT")
    xx = field: XObjective = originalObjective
    state.nodes = rn: state.elements = re: state.materialN = RM: state.rigidN = RR
    state.xy = RXY: state.triangles = RTri: state.materialId = RMatId: state.rigidId = RRigidId
    state.materials = RMaterials: state.body0 = RBody0: state.body1 = RBody1: Set state.boundary = RBoundary
    state.field = field: state.factor = factor: state.value = factor: state.q = 8: state.objective = originalObjective
    state.stressScale = RStressScale: state.normalizedAudit = audit: state.complete = True
    Call AccuracyExport.BatchState(state, proofPrefix & "_native.bin")
    Call RPX_ExportAssembled(proofPrefix & "_original.bin"): Call RPX_ExportExchangeX(proofPrefix & "_field.bin")
    Set context = CreateObject("Scripting.FileSystemObject").CreateTextFile(proofPrefix & "_context.tsv", False, True)
    context.WriteLine "source=Python primal only;root_imported=0;status_imported=0"
    context.WriteLine "policy" & vbTab & RPX_AnalysisPolicy(): context.WriteLine "phase" & vbTab & RPX_LoadPhase
    context.WriteLine "model" & vbTab & RPX_ExactModel(): context.WriteLine "fs" & vbTab & RPX_ExactDouble(factor): context.WriteLine "q" & vbTab & RSubdivisions
    context.Close
    RFsActive = False: Call RPX_InputEnd
    PythonOriginalFieldProbe = IIf(accepted, "PASS", "REJECT") & ";side=" & side & ";Fs=" & factor & ";elements=" & re & ";q=8;original_equality=" & equalityError & ";original_cone=" & coneError & ";physical_value=" & measured & ";physical_audit=" & audit & ";raw_yield=" & raw(1) & ";raw_equilibrium=" & raw(2) & ";raw_internal=" & raw(3) & ";raw_boundary=" & raw(4) & ";bounds_status_unchanged=1;root_imported=0"
End Function
