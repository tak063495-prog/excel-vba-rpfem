Attribute VB_Name = "AccuracyNativeFsProbe"
Option Explicit
Private Sub Setting(ByVal key As String, ByVal value As Variant)
    Dim ws As Worksheet, r As Long
    Set ws = RPX_Sheet("設定")
    For r = 4 To 1000
        If CStr(ws.Cells(r, 4).Value2) = key Then ws.Cells(r, 2).Value2 = value: Exit Sub
    Next r
    For r = 4 To 1000
        If Len(CStr(ws.Cells(r, 4).Value2)) = 0 And Len(CStr(ws.Cells(r, 2).Value2)) = 0 Then ws.Cells(r, 4).Value2 = key: ws.Cells(r, 2).Value2 = value: Exit Sub
    Next r
    err.Raise 5, , "SETTING_SLOT"
End Sub
Private Sub Check(ByVal valid As Boolean, ByVal message As String)
    If Not valid Then err.Raise 5, "AccuracyNativeFsProbe", message
End Sub
Public Function AccuracyNativeFsUnit(ByVal directory As String) As String
    Dim value As Double, endpoint As Double, a As Double, b As Double, center As Double, ratio As Double, state As String, raw() As Double, measured As Double, audit As Double
    Dim original As RPX_AccuracyProgram, stored() As Double, objective As Double, upperEndpoint As Double, upperA As Double, upperB As Double, upperState As String, file As Integer
    On Error GoTo Failed
    RPX_Setup
    Setting "TARGET", 16: Setting "QUADRATURE", 8: Setting "RAW_LOWER_AUDIT", 1: Setting "C0_AUTO", 0: Setting "ANALYSIS_POLICY", "ASSOCIATED"
    Setting "TEST_INJECTION", 1: Setting "HSD_SOLVE_SECONDS", 30
    Setting "LOWER_FIXED_RECOVERY", 0: Setting "ACCURACY_DLL_RECOVERY", 1: Setting "ACCURACY_DLL_SECONDS", 30: Setting "ACCURACY_DLL_DUAL_SECONDS", 5: Setting "ACCURACY_DLL_CAPTURE", 1
    Setting "ORIGINAL_FARKAS_CAPTURE", 1
    RPX_Sheet("入力").Range("B4").Value2 = 10: RPX_Sheet("浸透入力").Range("D4").Value2 = 0
    RPX_InputBegin: RPX_ReadInputs: RPX_GenerateMesh: RPX_TotalReferenceLoads
    XLogPath = directory & "\fs_probe.tsv": RPX_DiagStart "accuracy_fs_small": RFsActive = True
    value = RPX_Bound("lower", 0.4): RLowerField = xx: RUpperField = xx
    RLowerFs = 0.123: RUpperFs = 4.567: RBestLower.value = 0.123: RBestUpper.value = 4.567
    RPX_TestArm "fs_accuracy_original", 5, "HSD_STEP_STALLED"
    value = RPX_FsAccuracyProbeValue(0.4, "lower", 1.000001): RPX_TestDisarm
    Check value >= 0# And XStatus = "ORIGINAL_FIXED_LOWER_WITNESS", "NATIVE_POSITIVE_ENDPOINT_NOT_RECOVERED_" & XStatus
    Check RPX_FieldGate("lower", xx, 0.4, 8, XObjective, measured, audit, raw) And measured >= 1.000001, "NATIVE_POSITIVE_ENDPOINT_RAW_FAILED"
    Check Not RPX_ActualRootBracket("lower", a, b, 0.4), "SINGLE_NATIVE_WITNESS_CREATED_ROOT"
    RPX_TestArm "fs_accuracy_original", 5, "HSD_STEP_STALLED"
    value = RPX_FsAccuracyProbeValue(2#, "lower", 1.000001): RPX_TestDisarm
    Check value < 0# And XStatus = "ORIGINAL_LOWER_DUAL_GUIDE" And RPX_AccuracyGuideCurrent(2#, "lower"), "NATIVE_NEGATIVE_ORIGINAL_GUIDE_MISSING"
    XStatus = "OPTIMAL": Check Not RPX_AccuracyGuideCurrent(2#, "lower"), "WRONG_STATUS_GUIDE_ACCEPTED": XStatus = "ORIGINAL_LOWER_DUAL_GUIDE"
    XOriginalCompiled.cost(RLoadVariable) = 0#: Check Not RPX_AccuracyGuideCurrent(2#, "lower"), "FOREIGN_COST_GUIDE_ACCEPTED": XOriginalCompiled.cost(RLoadVariable) = -1#
    P4Solve = P4Solve + 1: Check Not RPX_AccuracyGuideCurrent(2#, "lower"), "FOREIGN_SOLVE_GUIDE_ACCEPTED"
    RFsTolerance = 0.0001: RFsSearchIncomplete = False: RFsLargestRootBracket = 0#: RPX_FsBracketReset
    endpoint = RPX_FsEndpoint("lower", 0.7, 0.3)
    Check RPX_ActualRootBracket("lower", a, b, endpoint), "NORMAL_NATIVE_ENABLED_ROOT_NOT_CLOSED"
    ' Keep a strict positive load margin for the auxiliary floor while the
    ' full bracket remains within the unchanged original root tolerance.
    b = b + 0.00005 * WorksheetFunction.Max(1#, (a + b) / 2#)
    center = (a + b) / 2#: ratio = (b - a) / (2# * center)
    RPX_FsBracketReset: RPX_TestArm "fs_accuracy_original", 5, "HSD_STEP_STALLED", 1
    endpoint = RPX_FsEndpoint("lower", center, ratio): RPX_TestDisarm
    state = RPX_FsAccuracyProbeDetail("lower")
    Check RPX_ActualRootBracket("lower", a, b, endpoint), "RECOVERED_NATIVE_ROOT_NOT_CLOSED_" & state
    Check InStr(state, "guide_b=True") > 0, "ROOT_NOT_INDEPENDENT_" & state
    Check b - a <= 0.0001 * WorksheetFunction.Max(1#, (a + b) / 2#), "NATIVE_ROOT_WIDTH"
    Check RPX_FsAccuracyProbeMetadata("lower", endpoint), "NATIVE_ROOT_METADATA"
    RSubdivisions = 4: Check Not RPX_ActualRootBracket("lower", measured, audit, endpoint), "Q4_ROOT_PROMOTION": RSubdivisions = 8
    Check Not RPX_ActualRootBracket("lower", measured, audit, endpoint + 0.000001), "WRONG_ROOT_ENDPOINT_ACCEPTED"
    ' At the returned A endpoint, live compiled cones can still belong to B.
    ' Export must reject that stale Fs representation, then audit the saved A
    ' field on a freshly prepared original A program without replacing it.
    stored = xx
    If XOriginalCompiled.strengthFactor <> endpoint Then Check Not RPX_AccuracyExport(endpoint, "lower", original), "FOREIGN_COMPILED_FS_EXPORT_ACCEPTED"
    value = RPX_Bound("lower", endpoint)
    Check RPX_AccuracyExport(endpoint, "lower", original), "FRESH_ROOT_ENDPOINT_EXPORT_REJECTED"
    Check RPX_AccuracyPrimalAudit(original, stored, objective), "SAVED_ROOT_ENDPOINT_ORIGINAL_RAW_REJECTED"
    Check RPX_FieldGate("lower", stored, endpoint, 8, objective, measured, audit, raw) And measured >= 1.000001, "SAVED_ROOT_ENDPOINT_PHYSICAL_REJECTED"
    RPX_AccuracyWriteProof directory & "\root_lower_recovered.bin", original, stored
    value = RPX_Bound("lower", b)
    Check RPX_OriginalLowerDualGate(b, 1.000001, True), "SAVED_ROOT_NEGATIVE_ORIGINAL_DUAL_REJECTED"
    ' The stricter optional feasibility backend may reject this close point;
    ' rejection retains the independently audited original LP proof above.
    value = Abs(CDbl(RPX_AccuracyRecoverDual(b, "lower")))
    ' General nonzero-cost upper inverse signs require a fresh original dual.
    RFsSearchIncomplete = False: RFsLargestRootBracket = 0#: RPX_FsBracketReset
    upperEndpoint = RPX_FsEndpoint("upper", 1.4, 0.2)
    upperState = RPX_FsAccuracyProbeDetail("upper")
    Check RPX_ActualRootBracket("upper", upperA, upperB, upperEndpoint), "NATIVE_UPPER_OWN_ROOT_NOT_CLOSED_" & upperState
    Check InStr(upperState, "status_a=ORIGINAL_ACCURACY_DUAL_GUIDE") > 0, "UPPER_INVERSE_WITHOUT_ORIGINAL_DUAL_" & upperState
    Check upperB - upperA <= 0.0001 * WorksheetFunction.Max(1#, (upperA + upperB) / 2#), "UPPER_ROOT_WIDTH"
    Check RPX_FsAccuracyProbeMetadata("upper", upperEndpoint), "UPPER_ROOT_METADATA"
    stored = xx: value = RPX_Bound("upper", upperEndpoint)
    Check RPX_AccuracyExport(upperEndpoint, "upper", original), "UPPER_ENDPOINT_EXPORT"
    Check RPX_AccuracyPrimalAudit(original, stored, objective), "UPPER_ENDPOINT_ORIGINAL_RAW"
    Check RPX_FieldGate("upper", stored, upperEndpoint, 8, objective, measured, audit, raw) And measured <= 0.999999, "UPPER_ENDPOINT_PHYSICAL"
    RPX_AccuracyWriteProof directory & "\root_upper_recovered.bin", original, stored
    file = FreeFile: Open directory & "\own_root_endpoints.bin" For Binary Access Write As #file
    Put #file, , "AC18R001": Put #file, , a: Put #file, , b: Put #file, , upperA: Put #file, , upperB: Close #file
    Check RLowerFs = 0.123 And RUpperFs = 4.567 And RBestLower.value = 0.123 And RBestUpper.value = 4.567, "NATIVE_ROOT_CHANGED_PUBLIC_BOUNDS"
    RFsActive = False: RPX_DiagFinish "PASS", "native_lower_signed_own_root"
    AccuracyNativeFsUnit = "PASS;lower_endpoint=" & endpoint & ";a=" & a & ";b=" & b & ";" & state & ";upper_endpoint=" & upperEndpoint & ";upper_a=" & upperA & ";upper_b=" & upperB & ";upper_state=" & upperState & ";fixed_affine_fallback=0;independent_original_negative_dual=1;native_primal_far_endpoint=1;public_bounds_unchanged=1"
    Exit Function
Failed:
    RPX_TestDisarm
    AccuracyNativeFsUnit = "FAIL " & err.number & " " & err.description & ";status=" & XStatus & ";native=" & AccuracyNativeReason
End Function
