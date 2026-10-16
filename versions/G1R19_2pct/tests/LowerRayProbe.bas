Attribute VB_Name = "LowerRayProbe"
Option Explicit
Private Type ProbeBits
    low As Long
    high As Long
End Type
Private Type ProbeDouble
    value As Double
End Type
Private Function Nonfinite() As Double
    Dim bits As ProbeBits, number As ProbeDouble
    bits.high = &H7FF80000: LSet number = bits: Nonfinite = number.value
End Function

Public Function DirectHomogeneousUnit(ByVal trace As String) As Variant
    Dim value As Double, lowerBefore As Double, upperBefore As Double
    On Error GoTo Failed
    OwnedSetting "TARGET", 16: OwnedSetting "QUADRATURE", 8: OwnedSetting "DIRECT_DLL", 1
    OwnedSetting "ACCURACY_DLL_RECOVERY", 1: OwnedSetting "ACCURACY_DLL_CAPTURE", 1
    OwnedSetting "TEST_INJECTION", 1: OwnedSetting "HSD_SOLVE_SECONDS", 30
    RPX_InputBegin: RPX_ReadInputs: RPX_GenerateMesh: RPX_TotalReferenceLoads
    XLogPath = trace: RPX_DiagStart "direct_homogeneous_unit": RFsActive = True
    lowerBefore = RBestLower.value: upperBefore = RBestUpper.value
    value = RPX_Bound("lower", 0.4)
    If Not AccuracyDirectLastAccepted Or AccuracyDirectLastKind <> "original_primal" Or XStatus <> "ORIGINAL_FIXED_LOWER_WITNESS" Then Err.Raise 5, , "DIRECT_NORMALIZED_FIELD_NOT_ACCEPTED"
    If Abs(value - 1.000003) > 0.0000000001 Then Err.Raise 5, , "DIRECT_NORMALIZATION_NOT_EXERCISED"
    If RBestLower.value <> lowerBefore Or RBestUpper.value <> upperBefore Then Err.Raise 5, , "DIRECT_MUTATED_SAVED_BOUNDS"
    RPX_TestArm "direct_dll_begin", 5, "FORCED_NATIVE_FAILURE"
    value = RPX_Bound("lower", 0.4)
    RPX_TestDisarm
    If AccuracyDirectLastAccepted Or XStatus <> "UNBOUNDED" Then Err.Raise 5, , "DIRECT_FAILURE_DID_NOT_FALL_BACK"
    If RBestLower.value <> lowerBefore Or RBestUpper.value <> upperBefore Then Err.Raise 5, , "FALLBACK_MUTATED_SAVED_BOUNDS"
    DirectHomogeneousUnit = Array("PASS", re, "direct_positive_copy_lambda_1.000003", "forced_failure_HSD_fallback", "saved_bounds_unchanged")
CleanUp:
    RFsActive = False: RPX_TestDisarm: RPX_DiagFinish "DONE", "direct_homogeneous_unit"
    RPX_InputEnd
    Exit Function
Failed:
    DirectHomogeneousUnit = Array("FAIL", Err.Number, Err.Description)
    Resume CleanUp
End Function
Private Function Key(ByRef values() As Double) As String
    Dim i As Long
    For i = LBound(values) To UBound(values): Key = Key & ":" & RPX_ExactDouble(values(i)): Next i
End Function
Private Sub Reject(ByVal accepted As Boolean, ByVal label As String)
    If accepted Then Err.Raise 5, "LowerRayProbe", "INVALID_RAY_ACCEPTED_" & label
End Sub
Public Function LowerRayUnit(ByVal trace As String) As Variant
    Dim value As Double, original As RPX_OriginalCompiledData, saved() As Double, recovered() As Double
    Dim oldValue As Double, oldObjective As Double, before As String, i As Long, offset As Double
    Dim measured As Double, audit As Double, raw() As Double
    On Error GoTo Failed
    OwnedSetting "TARGET", 16: OwnedSetting "QUADRATURE", 8
    OwnedSetting "DIRECT_DLL", 0: OwnedSetting "ACCURACY_DLL_RECOVERY", 1
    OwnedSetting "ACCURACY_DLL_CAPTURE", 1: OwnedSetting "HSD_SOLVE_SECONDS", 30
    RPX_InputBegin: RPX_ReadInputs: RPX_GenerateMesh: RPX_TotalReferenceLoads
    XLogPath = trace: RPX_DiagStart "lower_ray_unit": RFsActive = True
    value = RPX_Bound("lower", 0.4)
    If XStatus <> "UNBOUNDED" Then Err.Raise 5, , "ACTUAL_HSD_RAY_NOT_OBTAINED_" & XStatus
    saved = xx: original = XOriginalCompiled: before = Key(xx)
    oldValue = RValue: oldObjective = XObjective
    If Not RPX_AccuracyLowerRayWitness(0.4, recovered) Then Err.Raise 5, , "VALID_RAY_REJECTED"
    If Key(xx) <> before Or XStatus <> "UNBOUNDED" Or RValue <> oldValue Or XObjective <> oldObjective Then Err.Raise 5, , "HELPER_CHANGED_ORIGINAL_ITERATE"
    If Not RPX_FieldGate("lower", recovered, 0.4, 8, -1.000003, measured, audit, raw) Then Err.Raise 5, , "FRESH_RAY_FIELD_FAILED"
    If measured < 1.000001 Then Err.Raise 5, , "LOAD_MARGIN_FAILED"
    XStatus = "OPTIMAL": Reject RPX_AccuracyLowerRayWitness(0.4, recovered), "status": XStatus = "UNBOUNDED"
    RSubdivisions = 4: Reject RPX_AccuracyLowerRayWitness(0.4, recovered), "q": RSubdivisions = 8
    Reject RPX_AccuracyLowerRayWitness(0.41, recovered), "Fs"
    XOriginalCompiled.inputEpoch = XOriginalCompiled.inputEpoch - 1
    Reject RPX_AccuracyLowerRayWitness(0.4, recovered), "epoch": XOriginalCompiled = original
    XOriginalCompiled.cost(RLoadVariable) = 0#
    Reject RPX_AccuracyLowerRayWitness(0.4, recovered), "cost": XOriginalCompiled = original
    XOriginalCompiled.rhs(0) = 0.00000000000000000001
    Reject RPX_AccuracyLowerRayWitness(0.4, recovered), "original_b_nonzero": XOriginalCompiled = original
    offset = XC(0).offset(0): XC(0).offset(0) = 0.00000000000000000001
    Reject RPX_AccuracyLowerRayWitness(0.4, recovered), "original_h_nonzero": XC(0).offset(0) = offset
    AFActive = True: Reject RPX_AccuracyLowerRayWitness(0.4, recovered), "affine": AFActive = False
    xx(RLoadVariable) = 0#: Reject RPX_AccuracyLowerRayWitness(0.4, recovered), "zero_load": xx = saved
    xx(0) = Nonfinite(): Reject RPX_AccuracyLowerRayWitness(0.4, recovered), "nonfinite": xx = saved
    xx(0) = xx(0) + 1#: Reject RPX_AccuracyLowerRayWitness(0.4, recovered), "raw_residual": xx = saved
    If Key(xx) <> before Or RValue <> oldValue Or XObjective <> oldObjective Then Err.Raise 5, , "REJECTION_CHANGED_ORIGINAL_ITERATE"
    LowerRayUnit = Array("PASS", re, measured, "11_rejections", "input_primal_unchanged", "same_Fs_full_original_gate")
CleanUp:
    RFsActive = False: RPX_DiagFinish "DONE", "lower_ray_unit"
    RPX_InputEnd
    Exit Function
Failed:
    LowerRayUnit = Array("FAIL", Err.Number, Err.Description)
    Resume CleanUp
End Function
