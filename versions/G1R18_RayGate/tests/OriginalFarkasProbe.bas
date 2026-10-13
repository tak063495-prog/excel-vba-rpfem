Attribute VB_Name = "OriginalFarkasProbe"
Option Explicit
Private Type ProbeBits
    low As Long
    high As Long
End Type
Private Type ProbeDouble
    value As Double
End Type
Private retainedLowerKey As String, retainedUpperKey As String
Private Sub SettingValue(ByVal key As String, ByVal value As Variant)
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
    err.Raise 5, "OriginalFarkasProbe", "NO_SETTING_SLOT"
End Sub
Private Function Nonfinite(ByVal high As Long) As Double
    Dim bits As ProbeBits, number As ProbeDouble
    bits.high = high: LSet number = bits: Nonfinite = number.value
End Function
Private Function ArrayReady(ByRef values() As Double) As Boolean
    On Error GoTo Missing
    ArrayReady = (UBound(values) >= LBound(values)): Exit Function
Missing:
    If err.number <> 9 Then err.Raise err.number, err.source, err.description
    err.Clear
End Function
Private Sub PhysicalContext()
    Call RPX_Setup
    Call SettingValue("TARGET", 16): Call SettingValue("QUADRATURE", 8)
    Call SettingValue("C0_AUTO", 0): Call SettingValue("FS_TOL", 0.0001)
    Call SettingValue("DIAGNOSTIC_MODE", 1)
    Call SettingValue("ORIGINAL_FARKAS_CAPTURE", 0)
    Call SettingValue("HSD_SOLVE_SECONDS", 120)
    RPX_Sheet("材料データ").Range("C5").Value2 = 30#
    RPX_Sheet("材料データ").Range("B5").Value2 = 0#
    RPX_Sheet("材料データ").Range("E5").Value2 = 30#
    RPX_Sheet("浸透入力").Range("D4").Value2 = 0
    Call RPX_ReadInputs: Call RPX_GenerateMesh: Call RPX_TotalReferenceLoads
    RStrengthFactor = 1#: RFsActive = True: RPX_OrderMode = "upper"
    If Not RPX_OriginalPhysicalZeroCostUpper() Then err.Raise 5, "OriginalFarkasProbe", "PHYSICAL_CONTEXT_NOT_C0"
End Sub
Private Sub ToyProgram(ByRef candidate As RPX_HsdCandidateState, Optional ByVal diagonal As Boolean = False)
    Dim row As Object, r0 As Object, r1 As Object, r2 As Object
    Call RPX_Begin(3)
    Set row = RPX_Row(): Call RPX_Add(row, 0, 3#): Call RPX_Equal(row, -3#)
    If diagonal Then
        Set row = RPX_Row(): Call RPX_Add(row, 1, 1#): Call RPX_Equal(row)
        Set row = RPX_Row(): Call RPX_Add(row, 2, 1#): Call RPX_Equal(row)
    End If
    Set r0 = RPX_Row(): Call RPX_Add(r0, 0, 1#)
    Set r1 = RPX_Row(): Call RPX_Add(r1, 1, 1#)
    Set r2 = RPX_Row(): Call RPX_Add(r2, 2, 1#)
    Call RPX_Cone3(r0, r1, r2): Call RPX_Compile
    ReDim xx(0 To 2): ReDim xy(0 To XM): ReDim XSlack(0 To 2): ReDim XDual(0 To 2)
    xy(0) = 1#: XDual(0) = 1#
    Call RPX_HSDObservationBegin("upper")
    If Not RPX_HSDBuildCandidate(candidate, "PRIMAL_INFEASIBILITY_CERT", 1#, 1#, 1#, 0#, 0#) Then err.Raise 5, "OriginalFarkasProbe", "CANDIDATE_BUILD_FAILED"
End Sub
Private Sub Reject(ByRef candidate As RPX_HsdCandidateState, ByVal reason As String)
    If RPX_OriginalFarkasGate(candidate) Then err.Raise 5, "OriginalFarkasProbe", "INVALID_RAY_ACCEPTED_" & reason
End Sub
Public Function OriginalFarkasUnitProbe() As String
    Dim candidate As RPX_HsdCandidateState, valid As RPX_HsdCandidateState, original As RPX_OriginalCompiledData
    Dim originalNorm As Double, originalCost As Double, oldFs As Double, oldUpper As Double, oldLower As Double
    Call PhysicalContext: Call ToyProgram(valid)
    original = XOriginalCompiled: originalNorm = XEqRowNorm(0)
    If originalNorm <> 3# Or XOriginalCompiled.values(0) <> 3# Or XOriginalCompiled.rhs(0) <> -3# Then err.Raise 5, "OriginalFarkasProbe", "ORIGINAL_LITERAL_NOT_RETAINED"
    candidate = valid
    If Not RPX_HSDValidateCandidate(candidate) Then err.Raise 5, "OriginalFarkasProbe", "VALID_ORIGINAL_RAY_REJECTED_" & OriginalFarkasReason
    If candidate.outputStatus <> "ORIGINAL_FARKAS_GUIDE" Or ArrayReady(candidate.x) Or ArrayReady(candidate.field) Then err.Raise 5, "OriginalFarkasProbe", "RAY_CAPTURED_AS_FIELD"
    Call RPX_HSDCommitCandidate(candidate)
    If ArrayReady(xx) Or ArrayReady(RPX_HSDProofX) Or Not RPX_OriginalGuideCurrent(1#) Then err.Raise 5, "OriginalFarkasProbe", "GUIDE_COMMIT_INVALID"
    If OriginalFarkasStationarity > 0.0000001 Or OriginalFarkasCone > 0.00000001 Or Abs(OriginalFarkasSeparation - 1.000003) > 0.0000001 Then err.Raise 5, "OriginalFarkasProbe", "RAW_ACCEPTANCE_GATE_BYPASSED"
    oldUpper = RUpperFs: oldLower = RLowerFs
    candidate = valid: candidate.y(0) = 1.0000002: Call Reject(candidate, "stationarity")
    candidate = valid: candidate.dual(0) = -1#: Call Reject(candidate, "negative_cone")
    candidate = valid: candidate.y(0) = -1#: candidate.dual(0) = -1#: Call Reject(candidate, "negative_separation")
    candidate = valid: candidate.normalization = 0#: Call Reject(candidate, "zero_normalization")
    candidate = valid: candidate.normalization = Nonfinite(&H7FF00000): Call Reject(candidate, "infinite_normalization")
    candidate = valid: candidate.dual(0) = Nonfinite(&H7FF80000): Call Reject(candidate, "nan_dual")
    candidate = valid: candidate.y(0) = Nonfinite(&H7FF00000): Call Reject(candidate, "infinite_equality_dual")
    candidate = valid: candidate.subdivisions = 4: Call Reject(candidate, "q4")
    candidate = valid: candidate.inputEpoch = candidate.inputEpoch - 1: Call Reject(candidate, "input_epoch")
    candidate = valid: candidate.solveEpoch = candidate.solveEpoch - 1: Call Reject(candidate, "solve_epoch")
    candidate = valid: candidate.geometryEpoch = candidate.geometryEpoch - 1: Call Reject(candidate, "geometry_epoch")
    candidate = valid: candidate.fs = 1.0001: Call Reject(candidate, "factor")
    candidate = valid: candidate.modelKey = candidate.modelKey & "wrong": Call Reject(candidate, "model")
    candidate = valid: candidate.loadPhase = "fixed": Call Reject(candidate, "load_phase")
    candidate = valid: AFActive = True: Call Reject(candidate, "affine"): AFActive = False
    candidate = valid: AFActive = True
    If RPX_HSDValidateCandidate(candidate) Then err.Raise 5, "OriginalFarkasProbe", "AFFINE_RAY_VALIDATION_ACCEPTED"
    AFActive = False
    candidate = valid: ReDim Preserve candidate.dual(0 To XS): Call Reject(candidate, "extra_cone_dimension")
    candidate = valid: ReDim Preserve candidate.y(0 To XM + 1): Call Reject(candidate, "extra_equality_dimension")
    candidate = valid: XOriginalCompiled.affine = True: Call Reject(candidate, "compiled_affine"): XOriginalCompiled = original
    candidate = valid: XOriginalCompiled.cost(0) = 0.000000000000001: Call Reject(candidate, "tiny_original_cost"): XOriginalCompiled = original
    candidate = valid: XOriginalCompiled.rhs(0) = Nonfinite(&H7FF80000): Call Reject(candidate, "nan_original_rhs"): XOriginalCompiled = original
    candidate = valid: XOriginalCompiled.values(0) = Nonfinite(&H7FF00000): Call Reject(candidate, "infinite_original_equation"): XOriginalCompiled = original
    candidate = valid: XEqRowNorm(0) = 0#: Call Reject(candidate, "zero_row_norm"): XEqRowNorm(0) = originalNorm
    candidate = valid: XOriginalCompiled.inputEpoch = RPX_InputEpoch - 1: Call Reject(candidate, "compiled_epoch"): XOriginalCompiled = original
    candidate = valid: XOriginalCompiled.n = XN + 1: Call Reject(candidate, "compiled_dimensions"): XOriginalCompiled = original
    candidate = valid: XOriginalCompiled.ready = False: Call Reject(candidate, "compiled_unready"): XOriginalCompiled = original
    originalCost = RMaterials(0).cohesion: RMaterials(0).cohesion = 0.000000000000001
    candidate = valid: candidate.modelKey = RPX_ExactModel(): XOriginalCompiled.modelKey = candidate.modelKey
    Call Reject(candidate, "tiny_physical_cohesion"): RMaterials(0).cohesion = originalCost: XOriginalCompiled = original
    If RUpperFs <> oldUpper Or RLowerFs <> oldLower Then err.Raise 5, "OriginalFarkasProbe", "REJECTION_CHANGED_PHYSICAL_BOUNDS"
    ' Cone error is independently tested with exactly cancelling equality stationarity.
    Call ToyProgram(valid, True): candidate = valid
    candidate.y(1) = 1.00000002: candidate.dual(1) = 1.00000002
    Call Reject(candidate, "cone_without_stationarity_error")
    RFsActive = False: Call RPX_ResetPreparedState
    If XOriginalCompiled.ready Then err.Raise 5, "OriginalFarkasProbe", "PREPARED_RESET_RETAINED_RAW_PROGRAM"
    OriginalFarkasUnitProbe = "PASS;raw_row_dual_mapping=1;normalization_reaudited=1;no_primal_field_capture=1;stationarity_cone_nonfinite_scope_identity_rejections=1;physical_bounds_unchanged=1;prepared_reset=1"
End Function
Public Function OriginalFarkasNativeProbe(ByVal trace As String) As String
    Dim factor As Double, a As Double, b As Double, measured As Double, audit As Double, raw() As Double, rootValid As Boolean
    Call PhysicalContext
    XLogPath = trace: Call RPX_NativeSetMode("auto")
    Call RPX_DiagStart("original_farkas_native")
    RFsTolerance = 0.0001: RFsSearchIncomplete = False
    factor = RPX_FsEndpoint("upper")
    rootValid = RPX_ActualRootBracket("upper", a, b, factor)
    If XStatus <> "OPTIMAL" Or Not ArrayReady(xx) Then err.Raise 5, "OriginalFarkasProbe", "MISSING_NEGATIVE_PHYSICAL_ENDPOINT"
    If Not RPX_FieldGate("upper", xx, factor, 8, XObjective, measured, audit, raw) Then err.Raise 5, "OriginalFarkasProbe", "NEGATIVE_ORIGINAL_FIELD_GATE_FAILED"
    If Not rootValid Or b - a > RFsTolerance * WorksheetFunction.Max(1#, (a + b) / 2#) Then err.Raise 5, "OriginalFarkasProbe", "ORIGINAL_GUIDE_ROOT_INCOMPLETE"
    If factor <> b Then err.Raise 5, "OriginalFarkasProbe", "UPPER_ENDPOINT_NOT_BRACKET_B"
    RFsActive = False
    Call RPX_DiagFinish("PASS")
    OriginalFarkasNativeProbe = "PASS;upper=" & factor & ";a=" & a & ";b=" & b & ";elements=" & re & ";q=8;root=True;negative_physical_gate=PASS;root_positive_original_guide_required=1"
End Function

Private Sub CanonicalSlope()
    Dim ws As Worksheet
    Call RPX_Setup
    Call SettingValue("TARGET", 16): Call SettingValue("QUADRATURE", 8): Call SettingValue("C0_AUTO", 0)
    Call SettingValue("FS_TOL", 0.0001): Call SettingValue("HSD_SOLVE_SECONDS", 120): Call SettingValue("DIAGNOSTIC_MODE", 1)
    Call SettingValue("ORIGINAL_FARKAS_CAPTURE", 1)
    Call SettingValue("ANALYSIS_POLICY", "ASSOCIATED")
    Set ws = RPX_Sheet("要素定義")
    ws.Range("A7:T1006").ClearContents
    ws.Range("A7:C7").Value2 = Array(1, 0#, 0#): ws.Range("A8:C8").Value2 = Array(2, 70#, 0#)
    ws.Range("A9:C9").Value2 = Array(3, 70#, 10#): ws.Range("A10:C10").Value2 = Array(4, 50#, 10#)
    ws.Range("A11:C11").Value2 = Array(5, 20#, 30#): ws.Range("A12:C12").Value2 = Array(6, 0#, 30#)
    ws.Range("E7:H7").Value2 = Array(1, 1, 0, "1,2,3,4,5,6")
    ws.Range("J7:T7").Value2 = Array(1, 1, 1, 2, 0, "fixed", 0, 0#, 0#, 0#, 0#)
    ws.Range("J8:T8").Value2 = Array(2, 1, 2, 3, 0, "roller_x", 0, 0#, 0#, 0#, 0#)
    ws.Range("J9:T9").Value2 = Array(3, 1, 3, 4, 0, "load", 0, 0#, 0#, 0#, 0#)
    ws.Range("J10:T10").Value2 = Array(4, 1, 4, 5, 0, "load", 0, 0#, 0#, 0#, 0#)
    ws.Range("J11:T11").Value2 = Array(5, 1, 5, 6, 0, "load", 0, 0#, 0#, 0#, 0#)
    ws.Range("J12:T12").Value2 = Array(6, 1, 6, 1, 1, "roller_x", 0, 0#, 0#, 0#, 0#)
    RPX_Sheet("材料データ").Range("A5:E1004").ClearContents
    RPX_Sheet("材料データ").Range("A5:E5").Value2 = Array(1, 0#, 45#, 20#, 45#)
    RPX_Sheet("浸透入力").Range("D4").Value2 = 0
    Call RPX_ReadInputs: Call RPX_GenerateMesh
End Sub
Public Function C0WedgeGeometryProbe() As String
    Dim oldModel As String, material As Double, fixedLoad As Double, e As Long, k As Long, id As Long, area As Double, count As Long
    Call CanonicalSlope
    oldModel = RPX_ExactModel()
    If RPX_C0SlopeWedgeMesh(0#) Or RPX_C0SlopeWedgeMesh(20#) Or RPX_C0SlopeWedgeMesh(Nonfinite(&H7FF80000)) Then err.Raise 5, "OriginalFarkasProbe", "INVALID_WEDGE_DEPTH_ACCEPTED"
    If RPX_ExactModel() <> oldModel Then err.Raise 5, "OriginalFarkasProbe", "WEDGE_PREFLIGHT_CHANGED_MODEL"
    material = RMaterials(0).cohesion: RMaterials(0).cohesion = 0.000000000000001
    If RPX_C0SlopeWedgeMesh(0.01) Then err.Raise 5, "OriginalFarkasProbe", "NONZERO_COHESION_WEDGE_ACCEPTED"
    RMaterials(0).cohesion = material
    fixedLoad = DFaces(0).load0(1): DFaces(0).load0(1) = -1#
    If RPX_C0SlopeWedgeMesh(0.01) Then err.Raise 5, "OriginalFarkasProbe", "LOADED_WEDGE_ACCEPTED"
    DFaces(0).load0(1) = fixedLoad
    If RPX_ExactModel() <> oldModel Then err.Raise 5, "OriginalFarkasProbe", "REJECTED_WEDGE_CHANGED_MODEL"
    If Not RPX_C0SlopeWedgeMesh(0.01) Or rn <> 8 Or re <> 8 Then err.Raise 5, "OriginalFarkasProbe", "SUPPORTED_WEDGE_REJECTED"
    For e = 0 To 5
        For k = 0 To 1: If RXY(e, k) <> DPXY(e, k) Then err.Raise 5, "OriginalFarkasProbe", "EXTERIOR_VERTEX_CHANGED"
        Next k
    Next e
    For e = 0 To 7
        area = area + RArea(e)
        If RMatId(e) <> 0 Or RRigidId(e) <> -1 Or RBody0(e, 0) + RBody1(e, 0) <> 0# Or RBody0(e, 1) + RBody1(e, 1) <> -20# Then err.Raise 5, "OriginalFarkasProbe", "WEDGE_MATERIAL_BODY_CHANGED"
    Next e
    For id = 0 To RNEdge - 1
        If REdges(id).element(1) < 0 Then
            count = count + 1
            If REdges(id).a >= 6 Or REdges(id).b >= 6 Then err.Raise 5, "OriginalFarkasProbe", "WEDGE_BOUNDARY_CHANGED"
        End If
    Next id
    If Abs(area - 1400#) > 0.000000001 Or count <> 6 Or RBoundary.count <> 6 Or RConstraints.count <> 6 Then err.Raise 5, "OriginalFarkasProbe", "WEDGE_DOMAIN_CHANGED"
    Call RPX_TotalReferenceLoads
    If Not RPX_OriginalPhysicalZeroCostUpper() Then err.Raise 5, "OriginalFarkasProbe", "WEDGE_REFERENCE_LOAD_TRANSFORM_CHANGED_COST"
    C0WedgeGeometryProbe = "PASS;canonical_domain_area=1400;elements=8;boundary=6;original_vertices_material_body_faces=1;invalid_depth_cost_load_atomic_rejected=1"
End Function
Public Function C0WedgeNativeProbe(ByVal trace As String) As String
    Dim hint As Double, upper As Double, a As Double, b As Double, measured As Double, audit As Double, raw() As Double, rootValid As Boolean, field() As Double
    Call CanonicalSlope
    If Not RPX_C0SlopeWedgeMesh(0.01) Then err.Raise 5, "OriginalFarkasProbe", "NATIVE_WEDGE_REJECTED"
    Call RPX_TotalReferenceLoads: RSubdivisions = 8: RFsActive = True: RFsTolerance = 0.0001
    RFsSearchIncomplete = False: XLogPath = trace: Call RPX_NativeSetMode("auto")
    Call RPX_DiagStart("c0_wedge_native")
    hint = RPX_C0SlopeWedgeHint(0.01)
    upper = RPX_FsEndpoint("upper", hint + 0.00005, 0.002 / hint)
    rootValid = RPX_ActualRootBracket("upper", a, b, upper)
    If Not rootValid Or b - a > RFsTolerance * WorksheetFunction.Max(1#, (a + b) / 2#) Then
        Call RPX_DiagFinish("FAIL", RPX_FsProbeDetail() & ";root_valid=" & rootValid & ";endpoint=" & upper & ";returned_a=" & a & ";returned_b=" & b & ";tolerance=" & RFsTolerance)
        err.Raise 5, "OriginalFarkasProbe", "WEDGE_ORIGINAL_ROOT_INCOMPLETE " & RPX_FsProbeDetail() & ";root_valid=" & rootValid & ";endpoint=" & upper & ";returned_a=" & a & ";returned_b=" & b & ";tolerance=" & RFsTolerance
    End If
    If Not RPX_FieldGate("upper", xx, upper, 8, XObjective, measured, audit, raw) Then err.Raise 5, "OriginalFarkasProbe", "WEDGE_ORIGINAL_PHYSICAL_GATE_FAILED"
    ' Fresh original assembly at the actual negative endpoint, no solve/status uplift.
    field = xx: Call RPX_AssembleUpper: Call RPX_LiftUpperEpigraph(field)
    If Not RPX_OriginalPrimalGate(field) Then err.Raise 5, "OriginalFarkasProbe", "WEDGE_NEGATIVE_FULL_ORIGINAL_GATE_FAILED"
    Call RPX_Compile: xx = field
    Call RPX_ExportCompiled(trace & ".negative_program.bin")
    Call RPX_ExportExchangeContext(trace & ".negative_context.tsv")
    Call RPX_ExportExchangeX(trace & ".negative_field.bin")
    RFsActive = False: Call RPX_DiagFinish("PASS")
    C0WedgeNativeProbe = "PASS;upper=" & upper & ";q=8;elements=8;root_a=" & a & ";root_b=" & b & ";root=True;full_original_field_gate=PASS;analytical_hint_not_bound=1"
End Function

Private Function OwnWitnessKey(ByVal side As String) As String
    Dim state As RPX_MeshState, values() As String, i As Long
    If side = "lower" Then state = RBestLower Else state = RBestUpper
    Call RPX_RestoreWitnessState(state)
    ReDim values(0 To UBound(state.field))
    For i = 0 To UBound(state.field): values(i) = RPX_ExactDouble(state.field(i)): Next i
    OwnWitnessKey = RPX_ExactModel() & "|" & Join(values, "") & "|" & RPX_ExactDouble(state.value) & "|" & state.q & "|" & state.lowerRootValid & "|" & state.upperRootValid & "|" & state.searchIncomplete
End Function
Public Function C0WedgeAdaptProbe(ByVal depth As Double, ByVal injectFailure As Boolean, ByVal trace As String) As String
    Dim lowerKey As String, upperKey As String, measured As Double, audit As Double, raw() As Double
    Call CanonicalSlope
    Call SettingValue("ADAPT", 1): Call SettingValue("CYCLES", 0)
    Call SettingValue("ADAPT_ACCURACY_CAP", 64): Call SettingValue("ADAPT_ACCURACY_ROUNDS", 0)
    Call SettingValue("ADAPT_LOWER_TARGET", 0): Call SettingValue("ADAPT_LOWER_FAN_ROUNDS", 0)
    Call SettingValue("ADAPT_CERT_QUADRATURE", 0): Call SettingValue("GAP", 0.000000000001)
    Call SettingValue("ADAPT_C0_WEDGE_DEPTH", depth)
    If injectFailure Then
        Call SettingValue("TEST_INJECTION", 1)
        Call RPX_TestArm("upper_c0_wedge", 5, "CONIC_NOT_CONVERGED")
    End If
    XLogPath = trace: Call RPX_NativeSetMode("auto"): Call RPX_Run
    If Len(RPX_LastError) > 0 Then err.Raise 5, "OriginalFarkasProbe", RPX_LastError
    If Not RPX_AdaptPairComplete() Or Not RBestLower.complete Or Not RBestUpper.complete Then err.Raise 5, "OriginalFarkasProbe", "MISSING_ADAPT_WITNESSES"
    lowerKey = OwnWitnessKey("lower")
    If Not RPX_FieldGate("lower", RBestLower.field, RBestLower.value, RBestLower.q, RBestLower.objective, measured, audit, raw) Then err.Raise 5, "OriginalFarkasProbe", "RETAINED_LOWER_ORIGINAL_GATE_FAILED"
    upperKey = OwnWitnessKey("upper")
    If Not RPX_FieldGate("upper", RBestUpper.field, RBestUpper.value, RBestUpper.q, RBestUpper.objective, measured, audit, raw) Then err.Raise 5, "OriginalFarkasProbe", "RETAINED_UPPER_ORIGINAL_GATE_FAILED"
    If depth = 0# Then
        retainedLowerKey = lowerKey: retainedUpperKey = upperKey
    Else
        If retainedLowerKey <> lowerKey Then err.Raise 5, "OriginalFarkasProbe", "WEDGE_CHANGED_INDEPENDENT_LOWER_WITNESS"
        If injectFailure Then
            If retainedUpperKey <> upperKey Then err.Raise 5, "OriginalFarkasProbe", "FAILED_WEDGE_CHANGED_UPPER_WITNESS"
        Else
            If RBestUpper.elements <> 8 Or RBestUpper.q <> 8 Or Not RBestUpper.upperRootValid Then err.Raise 5, "OriginalFarkasProbe", "WEDGE_NOT_INDEPENDENTLY_ACCEPTED"
        End If
    End If
    C0WedgeAdaptProbe = "PASS;lower=" & RBestLower.value & ";upper=" & RBestUpper.value & ";lower_elements=" & RBestLower.elements & ";upper_elements=" & RBestUpper.elements & ";lower_exact_retained=1;physical_gates=PASS;injected_failure=" & injectFailure
End Function
