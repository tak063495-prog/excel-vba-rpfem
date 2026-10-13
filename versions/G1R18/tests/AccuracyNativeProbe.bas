Attribute VB_Name = "AccuracyNativeProbe"
Option Explicit
Private Type Bits
    low As Long
    high As Long
End Type
Private Type DValue
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
    err.Raise 5, , "SETTING_SLOT"
End Sub
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then err.Raise 5, "AccuracyNativeProbe", message
End Sub
Private Function Nonfinite(ByVal high As Long) As Double
    Dim bit As Bits, d As DValue
    bit.high = high: LSet d = bit: Nonfinite = d.value
End Function
Public Function AccuracyNativeNonfinite(ByVal high As Long) As Double
    AccuracyNativeNonfinite = Nonfinite(high)
End Function
Private Function VectorKey(ByRef x() As Double) As String
    Dim i As Long, s As String
    For i = 0 To UBound(x): s = s & "," & RPX_ExactDouble(x(i)): Next i
    VectorKey = s
End Function
Private Sub SaveProgram(ByVal path As String, ByRef pr As RPX_AccuracyProgram, ByRef field() As Double)
    Dim f As Integer, i As Long
    f = FreeFile: Open path For Binary Access Write As #f
    Put #f, , "AC18P001": Put #f, , pr.n: Put #f, , pr.m: Put #f, , pr.nnz: Put #f, , pr.coneCount: Put #f, , pr.equalityCount
    For i = 0 To pr.m: Put #f, , pr.rowPtr(i): Next i
    For i = 0 To pr.nnz - 1: Put #f, , pr.colId(i): Next i
    For i = 0 To pr.nnz - 1: Put #f, , pr.values(i): Next i
    For i = 0 To pr.m - 1: Put #f, , pr.rhs(i): Next i
    For i = 0 To pr.n - 1: Put #f, , pr.originalCost(i): Next i
    For i = 0 To pr.coneCount - 1: Put #f, , pr.coneKind(i): Put #f, , pr.coneDim(i): Next i
    Put #f, , pr.factor: Put #f, , pr.q
    For i = 0 To UBound(field): Put #f, , field(i): Next i
    Close #f
End Sub
Public Function AccuracyNativeUnit(ByVal directory As String) As String
    Dim original As RPX_AccuracyProgram, candidate() As Double, oldX() As Double, oldCost As Double, value As Double, objective As Double
    Dim i As Long, key As String, a As Double, b As Double, raw() As Double, measured As Double, audit As Double, state As String
    On Error GoTo Failed
    Call RPX_Setup
    Setting "TARGET", 16: Setting "QUADRATURE", 8: Setting "RAW_LOWER_AUDIT", 1: Setting "C0_AUTO", 0
    Setting "ANALYSIS_POLICY", "ASSOCIATED": Setting "HSD_SOLVE_SECONDS", 30: Setting "TEST_INJECTION", 1
    Setting "ACCURACY_DLL_RECOVERY", 1: Setting "ACCURACY_DLL_SECONDS", 30: Setting "ACCURACY_DLL_DUAL_SECONDS", 20
    Setting "ACCURACY_DLL_CAPTURE", 1
    RPX_Sheet("入力").Range("B4").Value2 = 10: RPX_Sheet("浸透入力").Range("D4").Value2 = 0
    RPX_InputBegin: RPX_ReadInputs: RPX_GenerateMesh: RPX_TotalReferenceLoads
    XLogPath = directory & "\native_probe.tsv": RPX_DiagStart "accuracy_native_small": RFsActive = True
    value = RPX_Bound("lower", 0.4): oldX = xx: key = VectorKey(xx)
    RLowerFs = 0.123: RUpperFs = 4.567: RBestLower.value = 0.123: RBestUpper.value = 4.567
    RLowerField = xx: RUpperField = xx
    Check RPX_AccuracyExport(0.4, "lower", original), "ORIGINAL_EXPORT_REJECTED"
    SaveProgram directory & "\lower_original.bin", original, oldX
    Check RPX_AccuracyRecoverPrimal(0.4, "lower", candidate), "LOWER_PROX_REJECTED_" & AccuracyNativeReason
    Check VectorKey(xx) = key And RLowerFs = 0.123 And RUpperFs = 4.567 And RBestLower.value = 0.123 And RBestUpper.value = 4.567, "PROX_MUTATED_ORIGINAL_STATE"
    Check RPX_AccuracyPrimalAudit(original, candidate, objective), "LOWER_ORIGINAL_RAW_FAILED"
    Check RPX_FieldGate("lower", candidate, 0.4, 8, objective, measured, audit, raw) And measured >= 1.000001, "LOWER_PHYSICAL_FAILED"
    SaveProgram directory & "\lower_recovered.bin", original, candidate
    ' An actual original-row correction changes only a local trial vector.
    candidate = oldX
    For i = 0 To UBound(candidate): candidate(i) = candidate(i) * (1.0000015 / oldX(RLoadVariable)): Next i
    candidate(0) = candidate(0) + 0.00001
    Check Not RPX_AccuracyPrimalAudit(original, candidate, objective), "EQUALITY_FAULT_NOT_EFFECTIVE"
    Check RPX_AccuracyProbeEquality(original, candidate, objective), "ORIGINAL_ROW_CORRECTION_REJECTED"
    Check RPX_AccuracyPrimalAudit(original, candidate, objective), "CORRECTED_ORIGINAL_ROW_FAILED"
    SaveProgram directory & "\lower_lsmr_recovered.bin", original, candidate
    Check VectorKey(xx) = key, "LSMR_CHANGED_ORIGINAL_SOURCE"
    candidate(0) = Nonfinite(&H7FF80000): Check Not RPX_AccuracyPrimalAudit(original, candidate, objective), "NAN_PRIMAL_ACCEPTED"
    candidate = oldX: original.epoch = original.epoch - 1: Check Not RPX_AccuracyPrimalAudit(original, candidate, objective), "FOREIGN_EPOCH_ACCEPTED"
    Check RPX_AccuracyExport(0.4, "lower", original), "RESTORED_EXPORT_FAILED"
    RSubdivisions = 4: Check Not RPX_AccuracyExport(0.4, "lower", original), "Q4_EXPORT_ACCEPTED": RSubdivisions = 8
    AFActive = True: Check Not RPX_AccuracyExport(0.4, "lower", original), "AFFINE_EXPORT_ACCEPTED": AFActive = False
    oldCost = XOriginalCompiled.cost(RLoadVariable): XOriginalCompiled.cost(RLoadVariable) = 0#
    Check Not RPX_AccuracyExport(0.4, "lower", original), "FIXED_COST_EXPORT_ACCEPTED": XOriginalCompiled.cost(RLoadVariable) = oldCost
    XCancel = True: Check Not RPX_AccuracyExport(0.4, "lower", original), "CANCELLED_EXPORT_ACCEPTED": XCancel = False
    value = RPX_Bound("lower", 2#)
    Check RPX_AccuracyRecoverDual(2#, "lower"), "LOWER_ORIGINAL_DUAL_REJECTED_" & AccuracyNativeReason
    Check AccuracyNativeDualStat <= 0.0000001 And AccuracyNativeDualCone <= 0.00000001 And AccuracyNativeDualObjective >= -0.999998, "LOWER_DUAL_RAW_GUARD"
    Check Not RPX_ActualRootBracket("lower", a, b, 2#), "SINGLE_DUAL_CREATED_ROOT"
    value = RPX_Bound("upper", 2#): key = VectorKey(xx): oldX = xx
    Check RPX_AccuracyExport(2#, "upper", original), "UPPER_EXPORT_REJECTED_" & AccuracyNativeReason
    Check RPX_AccuracyRecoverPrimal(2#, "upper", candidate), "UPPER_PROX_REJECTED_" & AccuracyNativeReason
    Check VectorKey(xx) = key, "UPPER_PROX_MUTATED_ORIGINAL_X"
    Check RPX_AccuracyPrimalAudit(original, candidate, objective), "UPPER_ORIGINAL_RAW_FAILED"
    SaveProgram directory & "\upper_recovered.bin", original, candidate
    Check RPX_FieldGate("upper", candidate, 2#, 8, objective, measured, audit, raw) And measured <= 0.99999901, "UPPER_PHYSICAL_FAILED"
    value = RPX_Bound("upper", 0.4)
    Check RPX_AccuracyRecoverDual(0.4, "upper"), "UPPER_ORIGINAL_DUAL_REJECTED_" & AccuracyNativeReason
    Check AccuracyNativeDualStat <= 0.0000001 And AccuracyNativeDualCone <= 0.00000001 And AccuracyNativeDualObjective >= 1.000003, "UPPER_DUAL_RAW_GUARD"
    Check RLowerFs = 0.123 And RUpperFs = 4.567 And VectorKey(RLowerField) = VectorKey(RUpperField), "BOUNDS_CHANGED_BY_AUXILIARY_DUAL"
    RFsActive = False: RPX_DiagFinish "PASS", "accuracy_native_original_gate"
    AccuracyNativeUnit = "PASS;lower_prox=1;upper_prox=1;original_duals=2;original_raw_primal=1e-7;dual_stat=1e-7;dual_cone=1e-8;qp_dual_used=0;bounds_unchanged=1;q4_affine_cost_epoch_nonfinite_cancel_rejected=1"
    Exit Function
Failed:
    AccuracyNativeUnit = "FAIL " & err.number & " " & err.description & ";working=" & XN & "," & XM & "," & XS & "," & XB & ";orig=" & XOriginalCompiled.n & "," & XOriginalCompiled.m & "," & XOriginalCompiled.s & "," & XOriginalCompiled.b & ";mode=" & RPX_OrderMode & "/" & XOriginalCompiled.modeName & ";ready=" & XOriginalCompiled.ready & ";q=" & RSubdivisions & "/" & XOriginalCompiled.subdivisions
End Function
