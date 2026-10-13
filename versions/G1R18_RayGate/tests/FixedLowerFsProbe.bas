Attribute VB_Name = "FixedLowerFsProbeModule"
Option Explicit
Private Type ProbeBits
    low As Long
    high As Long
End Type
Private Type ProbeDouble
    value As Double
End Type
Private Sub Setting(ByVal key As String, ByVal value As Variant)
    Dim ws As Worksheet, r As Long
    Set ws = RPX_Sheet("設定")
    For r = 4 To 1000
        If CStr(ws.Cells(r, 4).Value2) = key Then ws.Cells(r, 2).Value2 = value: Exit Sub
    Next r
    For r = 4 To 1000
        If Len(CStr(ws.Cells(r, 4).Value2)) = 0 And Len(CStr(ws.Cells(r, 2).Value2)) = 0 Then ws.Cells(r, 4).Value2 = key: ws.Cells(r, 2).Value2 = value: Exit Sub
    Next r
    err.Raise 5, "FixedLowerFsProbe", "SETTING_SLOT"
End Sub
Private Sub Check(ByVal valid As Boolean, ByVal message As String)
    If Not valid Then err.Raise 5, "FixedLowerFsProbe", message
End Sub
Private Function VectorKey(ByRef vector() As Double) As String
    Dim i As Long
    For i = LBound(vector) To UBound(vector): VectorKey = VectorKey & ":" & RPX_ExactDouble(vector(i)): Next i
End Function
Private Function Nonfinite(ByVal high As Long) As Double
    Dim bits As ProbeBits, number As ProbeDouble
    bits.high = high: LSet number = bits: Nonfinite = number.value
End Function
Public Function FixedLowerFsProbe(ByVal trace As String) As String
    Dim oldLower As String, oldUpper As String, value As Double, a As Double, b As Double, endpoint As Double
    Dim measured As Double, audit As Double, raw() As Double, y() As Double, z() As Double, original As RPX_OriginalCompiledData
    Dim number As Long, first As Long, bestLower As Double, bestUpper As Double, i As Long, state As String, upperObjective As Double, endpointField() As Double, center As Double, ratio As Double, oldEpoch As Long, oldPolicy As String
    Call RPX_Setup
    Call Setting("TARGET", 16): Call Setting("QUADRATURE", 8): Call Setting("C0_AUTO", 0): Call Setting("RAW_LOWER_AUDIT", 1)
    Call Setting("TEST_INJECTION", 1): Call Setting("DIAGNOSTIC_MODE", 1): Call Setting("HSD_SOLVE_SECONDS", 120)
    Call Setting("LOWER_FIXED_RECOVERY", 1): Call Setting("LOWER_FIXED_RECOVERY_SECONDS", 60): Call Setting("ORIGINAL_FARKAS_CAPTURE", 1)
    RPX_Sheet("材料データ").Range("B5").Value2 = 20#: RPX_Sheet("材料データ").Range("C5").Value2 = 30#: RPX_Sheet("材料データ").Range("E5").Value2 = 30#
    RPX_Sheet("浸透入力").Range("D4").Value2 = 0
    Call RPX_InputBegin: Call RPX_ReadInputs: Call RPX_GenerateMesh: Call RPX_TotalReferenceLoads
    XLogPath = trace: Call RPX_DiagStart("fixed_lower_fs_small"): RFsActive = True
    value = RPX_Bound("upper", 2#): RUpperField = xx: RUpperFs = 2#: upperObjective = XObjective
    value = RPX_Bound("lower", 0.4): RLowerField = xx: RLowerFs = 0.4
    RBestLower.value = 0.4: RBestUpper.value = 2#: bestLower = RBestLower.value: bestUpper = RBestUpper.value
    oldLower = VectorKey(RLowerField): oldUpper = VectorKey(RUpperField)
    ' Inject a failure after a real original solve. Recovery may install only sigma.
    Call RPX_TestArm("fs_lower_original", 5, "HSD_STEP_STALLED")
    value = RPX_FsProbeValue(0.4, "lower", 1.000001): Call RPX_TestDisarm
    Call Check(XStatus = "ORIGINAL_FIXED_LOWER_WITNESS" And value >= 0#, "FAILED_POINT_NOT_RECOVERED")
    Call Check(RPX_FieldGate("lower", xx, 0.4, 8, XObjective, measured, audit, raw), "RECOVERED_ENDPOINT_PHYSICAL_FAILED")
    Call Check(Abs(measured - 1.000003) <= 0.0000001, "RECOVERED_FIXED_LOAD_VALUE")
    Call Check(Not RPX_ActualRootBracket("lower", a, b, 0.4), "SINGLE_RECOVERY_CREATED_ROOT")
    Call Check(VectorKey(RLowerField) = oldLower And VectorKey(RUpperField) = oldUpper And RLowerFs = 0.4 And RUpperFs = 2#, "RECOVERY_CHANGED_SAVED_BOUNDS")
    xx = RLowerField
    ' The opposite point is a fresh FREE-lambda solve and original raw dual audit.
    value = RPX_FsProbeValue(2#, "lower", 1.000001)
    Call Check(value < 0# And XStatus = "ORIGINAL_LOWER_DUAL_GUIDE", "ORIGINAL_NEGATIVE_DUAL_MISSING_" & OriginalLowerDualReason)
    Call Check(RPX_OriginalLowerDualGate(2#, 1.000001), "VALID_LOWER_DUAL_REJECTED")
    y = xy: z = XDual: original = XOriginalCompiled
    xy(0) = xy(0) + 1#: Call Check(Not RPX_OriginalLowerDualGate(2#, 1.000001), "BAD_STATIONARITY_ACCEPTED"): xy = y
    first = XC(0).first: XDual(first) = -1#: Call Check(Not RPX_OriginalLowerDualGate(2#, 1.000001), "BAD_CONE_ACCEPTED"): XDual = z
    xy(0) = Nonfinite(&H7FF00000): Call Check(Not RPX_OriginalLowerDualGate(2#, 1.000001), "INFINITE_LOWER_DUAL_ACCEPTED"): xy = y
    XDual(first) = Nonfinite(&H7FF80000): Call Check(Not RPX_OriginalLowerDualGate(2#, 1.000001), "NAN_LOWER_DUAL_ACCEPTED"): XDual = z
    XOriginalCompiled.cost(RLoadVariable) = 0#: Call Check(Not RPX_OriginalLowerDualGate(2#, 1.000001), "FIXED_ZERO_OBJECTIVE_DUAL_ACCEPTED"): XOriginalCompiled = original
    XQ(RLoadVariable) = Nonfinite(&H7FF80000): Call Check(Not RPX_OriginalLowerDualGate(2#, 1.000001), "NAN_WORKING_COST_ACCEPTED"): XQ(RLoadVariable) = -1#
    AFActive = True: Call Check(Not RPX_OriginalLowerDualGate(2#, 1.000001), "AFFINE_LOWER_DUAL_ACCEPTED"): AFActive = False
    RSubdivisions = 4: Call Check(Not RPX_OriginalLowerDualGate(2#, 1.000001), "Q4_LOWER_DUAL_ACCEPTED"): RSubdivisions = 8
    XOriginalCompiled.inputEpoch = XOriginalCompiled.inputEpoch - 1: Call Check(Not RPX_OriginalLowerDualGate(2#, 1.000001), "STALE_LOWER_DUAL_ACCEPTED"): XOriginalCompiled = original
    ' An actual q8 bracket must use an independent original negative guide.
    RFsTolerance = 0.0001: RFsSearchIncomplete = False: RFsLargestRootBracket = 0#: Call RPX_FsBracketReset
    endpoint = RPX_FsEndpoint("lower", 0.7, 0.3)
    state = RPX_LowerFsProbeDetail()
    Call Check(RPX_ActualRootBracket("lower", a, b, endpoint), "FREE_LOWER_ROOT_NOT_COMPLETE_" & state)
    Call Check(InStr(state, "guide_b=True") > 0 And b - a <= 0.0001 * WorksheetFunction.Max(1#, (a + b) / 2#), "ROOT_WITHOUT_ORIGINAL_NEGATIVE_PROOF")
    center = (a + b) / 2#: ratio = (b - a) / (2# * center)
    Call RPX_FsBracketReset: Call RPX_TestArm("fs_lower_original", 5, "HSD_STEP_STALLED")
    endpoint = RPX_FsEndpoint("lower", center, ratio): Call RPX_TestDisarm
    state = RPX_LowerFsProbeDetail()
    Call Check(RPX_ActualRootBracket("lower", a, b, endpoint), "RECOVERED_ROOT_NOT_COMPLETE_" & state)
    Call Check(InStr(state, "recovered_a=True") > 0 And InStr(state, "guide_b=True") > 0, "RECOVERY_ROOT_NOT_INDEPENDENTLY_CLOSED_" & state)
    RSubdivisions = 4: Call Check(Not RPX_ActualRootBracket("lower", measured, audit, endpoint), "RECOVERED_Q4_ROOT_ACCEPTED"): RSubdivisions = 8
    Call Check(RPX_LowerRootMetadataProbe(endpoint), "RECOVERED_ROOT_METADATA_GUARDS_FAILED")
    Call Check(Not RPX_ActualRootBracket("lower", measured, audit, endpoint + 0.0000001), "RECOVERED_WRONG_ENDPOINT_ROOT_ACCEPTED")
    Call Check(RPX_ActualRootBracket("lower", a, b, endpoint), "RECOVERED_ROOT_LOST_AFTER_SCOPE_REJECTIONS")
    Call Check(RPX_FieldGate("lower", xx, endpoint, 8, XObjective, measured, audit, raw), "FINAL_ENDPOINT_PHYSICAL_FAILED")
    endpointField = xx: Call RPX_AssembleLower: xx = endpointField
    Call RPX_ExportAssembled(trace & ".free_original.bin"): Call RPX_ExportExchangeX(trace & ".free_field.bin")
    Call RPX_ExportExchangeContext(trace & ".endpoint_context.tsv")
    Call Check(VectorKey(RLowerField) = oldLower And VectorKey(RUpperField) = oldUpper And RLowerFs = 0.4 And RUpperFs = 2# And RBestLower.value = bestLower And RBestUpper.value = bestUpper, "ROOT_PROBE_CHANGED_SAVED_BOUNDS")
    ' A rejected recovery after solving preserves both previously saved fields.
    Call RPX_TestArm("fixed_recovery_original_gate", 5, "RECOVERY_INJECTED_AUDIT_FAILURE")
    XStatus = "NUMERICAL_UNKNOWN": RStrengthFactor = 0.4: RProgramKind = "": RPX_OrderMode = "lower"
    Call RPX_AssembleLower: Call RPX_Compile
    value = 0#: Dim rejected() As Double
    Call Check(Not RPX_TryFixedLowerRecovery(0.4, rejected, 60), "FAILED_RECOVERY_ADOPTED"): Call RPX_TestDisarm
    Call Check(VectorKey(RLowerField) = oldLower And VectorKey(RUpperField) = oldUpper And RLowerFs = 0.4 And RUpperFs = 2#, "FAILED_RECOVERY_CHANGED_SAVED_BOUNDS")
    Call Check(RPX_FieldGate("lower", RLowerField, 0.4, 8, -RLowerField(RLoadVariable), measured, audit, raw), "RETAINED_LOWER_PHYSICAL_FAILED")
    Call Check(RPX_FieldGate("upper", RUpperField, 2#, 8, upperObjective, measured, audit, raw), "RETAINED_UPPER_PHYSICAL_FAILED")
    RFsActive = False: Call RPX_DiagFinish("PASS"): Call RPX_InputEnd
    FixedLowerFsProbe = "PASS;elements=" & re & ";positive_sameFs_recovery=1;independent_free_lower_dual=1;root_a=" & a & ";root_b=" & b & ";root=TRUE;q=8;fixed_dual_used=0;both_saved_bounds_fields_retained=1;raw_stationarity_cone_nonfinite_cost_q_epoch_AF_rejected=1"
End Function
