Attribute VB_Name = "FixedLowerRecoveryProbe"
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
    err.Raise 5, "FixedLowerRecoveryProbe", "SETTING_SLOT"
End Sub
Private Sub Check(ByVal value As Boolean, ByVal reason As String)
    If Not value Then err.Raise 5, "FixedLowerRecoveryProbe", reason
End Sub
Private Function VectorKey(ByRef values() As Double) As String
    Dim i As Long
    On Error GoTo Missing
    VectorKey = CStr(UBound(values))
    For i = LBound(values) To UBound(values): VectorKey = VectorKey & ":" & RPX_ExactDouble(values(i)): Next i
    Exit Function
Missing:
    If err.number <> 9 Then err.Raise err.number, err.source, err.description
    err.Clear: VectorKey = "UNALLOCATED"
End Function
Private Function ObservationKey() As String
    Dim state As RPX_HsdObservationSnapshot
    Call RPX_HSDObservationSave(state)
    ObservationKey = state.mode & ":" & state.kind & ":" & state.candidateId & ":" & state.context & ":" & state.geometry & ":" & state.path & ":" & state.captured & ":" & state.enabled
    ObservationKey = ObservationKey & ":" & RPX_ExactDouble(state.tau) & ":" & RPX_ExactDouble(state.kappa) & ":" & RPX_ExactDouble(state.normalization)
    ObservationKey = ObservationKey & ":" & VectorKey(state.x) & ":" & VectorKey(state.y) & ":" & VectorKey(state.slack) & ":" & VectorKey(state.dual) & ":" & VectorKey(state.proof)
End Function
Private Function CompiledKey() As String
    Dim i As Long, j As Long, r As Long
    CompiledKey = XN & ":" & XM & ":" & XS & ":" & XB & ":" & RLoadVariable & ":" & XOriginalCompiled.ready
    For i = 0 To XN - 1: CompiledKey = CompiledKey & ":" & RPX_ExactDouble(XQ(i)): Next i
    For i = 0 To XM - 1: CompiledKey = CompiledKey & ":" & RPX_ExactDouble(XERhs(i)) & ":" & RPX_ExactDouble(XEqRowNorm(i)): Next i
    For i = 0 To XEPtr(XM) - 1: CompiledKey = CompiledKey & ":" & XEId(i) & ":" & RPX_ExactDouble(XEVal(i)): Next i
    CompiledKey = CompiledKey & ":" & XOriginalCompiled.modelKey & ":" & XOriginalCompiled.inputEpoch & ":" & XOriginalCompiled.subdivisions
    For i = 0 To UBound(XOriginalCompiled.values): CompiledKey = CompiledKey & ":" & RPX_ExactDouble(XOriginalCompiled.values(i)): Next i
    For i = 0 To XN + XM - 1: CompiledKey = CompiledKey & ":" & XPerm(i) & ":" & XInv(i) & ":" & XKPtr(i) & ":" & XKDiag(i): Next i
    For i = 0 To UBound(XKBase): CompiledKey = CompiledKey & ":" & XKId(i) & ":" & RPX_ExactDouble(XKBase(i)): Next i
    For i = 0 To XB - 1
        For r = 0 To XC(i).dimn - 1
            CompiledKey = CompiledKey & ":" & RPX_ExactDouble(XC(i).offset(r))
            For j = 0 To XC(i).count - 1: CompiledKey = CompiledKey & ":" & XC(i).ids(j) & ":" & RPX_ExactDouble(XC(i).coef(r, j)): Next j
        Next r
    Next i
    For i = 0 To re - 1: CompiledKey = CompiledKey & ":" & RStressStart(i): Next i
End Function
Public Function FixedLowerRecoverySmallProbe(ByVal trace As String, Optional ByVal c0Fixture As Boolean = False) As String
    Dim saved() As Double, field() As Double, candidate() As Double, lowKey As String, upKey As String, contextKey As String, model As String
    Dim measured As Double, audit As Double, raw() As Double, value As Double, a As Double, b As Double, i As Long
    Dim accepted As Boolean, savedStatus As String, key As String, cached As Boolean, scopeQ As Long, upperObjective As Double, recoveredValue As Double, oldMode As String, oldKind As String, oldCandidateId As Long, observation As String
    Call RPX_Setup
    Call Setting("TARGET", 16): Call Setting("QUADRATURE", 8): Call Setting("C0_AUTO", 0): Call Setting("RAW_LOWER_AUDIT", 1)
    Call Setting("TEST_INJECTION", 1): Call Setting("LOWER_MODEL_CACHE", 1): Call Setting("HSD_SOLVE_SECONDS", 120): Call Setting("DIAGNOSTIC_MODE", 1)
    Call Setting("HSD_CAPTURE", 1)
    RPX_Sheet("材料データ").Range("B5").Value2 = IIf(c0Fixture, 0#, 20#)
    RPX_Sheet("材料データ").Range("C5").Value2 = 30#
    RPX_Sheet("材料データ").Range("E5").Value2 = 30#
    RPX_Sheet("浸透入力").Range("D4").Value2 = 0
    Call RPX_InputBegin: Call RPX_ReadInputs: Call RPX_GenerateMesh: Call RPX_TotalReferenceLoads
    XLogPath = trace: Call RPX_DiagStart("fixed_lower_recovery_small")
    RFsActive = True: RPX_OrderMode = "upper"
    value = RPX_Bound("upper", 2#): RUpperField = xx: upperObjective = XObjective: RUpperFs = 2#
    Call Check(RPX_FieldGate("upper", RUpperField, 2#, 8, upperObjective, measured, audit, raw), "INITIAL_UPPER_PHYSICAL_FAILED")
    RPX_OrderMode = "lower"
    value = RPX_Bound("lower", 0.4): saved = xx
    If c0Fixture Then
        Call Check(XStatus = "UNBOUNDED", "INITIAL_C0_LOWER_NOT_RECESSION")
        For i = 0 To UBound(saved): saved(i) = saved(i) * 1.000003: Next i
        Call RPX_AssembleLower
        Call Check(RPX_OriginalPrimalGate(saved), "INITIAL_C0_ORIGINAL_PRIMAL_FAILED")
        Call RPX_Compile: xx = saved: Call RPX_LowerCacheStore
    Else
        Call Check(XStatus = "OPTIMAL", "INITIAL_FREE_LOWER_NOT_OPTIMAL")
    End If
    RLowerField = saved: RLowerFs = 0.4
    Call Check(RPX_FieldGate("lower", RLowerField, 0.4, 8, -saved(RLoadVariable), measured, audit, raw), "INITIAL_LOWER_PHYSICAL_FAILED")
    RBestLower.value = 0.4: RBestUpper.value = 2#
    lowKey = VectorKey(RLowerField): upKey = VectorKey(RUpperField): contextKey = CompiledKey(): model = RPX_ExactModel()
    cached = RPX_LowerCacheMatch()
    If c0Fixture Then Call Check(Not cached, "C0_CACHE_SCOPE_UNEXPECTED") Else Call Check(cached, "INITIAL_LOWER_CACHE_MISSING")
    XStatus = "NUMERICAL_UNKNOWN": savedStatus = XStatus
    oldMode = RPX_HSDAuditMode: oldKind = RPX_HSDKind: oldCandidateId = RPX_HSDCandidateId
    observation = ObservationKey()
    Call RPX_TestArm("fixed_recovery_solve", 5, "RECOVERY_INJECTED_FAILURE")
    accepted = RPX_TryFixedLowerRecovery(0.4, field, 30): Call RPX_TestDisarm
    Call Check(Not accepted And FixedLowerRecoveryReason = "RECOVERY_INJECTED_FAILURE", "FAILED_RECOVERY_ACCEPTED")
    Call Check(CompiledKey() = contextKey And VectorKey(xx) = VectorKey(saved), "FAILURE_PROGRAM_OR_ITERATE_CHANGED")
    Call Check(VectorKey(RLowerField) = lowKey And VectorKey(RUpperField) = upKey And RPX_ExactModel() = model, "FAILURE_WITNESSES_CHANGED")
    Call Check(RPX_LowerCacheMatch() = cached And XStatus = savedStatus And Not AFActive, "FAILURE_CACHE_STATUS_OR_AFFINE_CHANGED")
    Call Check(ObservationKey() = observation, "FAILURE_PRIVATE_HSD_OBSERVATION_CHANGED")
    Call RPX_TestArm("fixed_recovery_solve", 5, "HSD_LINEAR_QUALITY_FAILED")
    accepted = RPX_TryFixedLowerRecovery(0.4, field, 30): Call RPX_TestDisarm
    Call Check(Not accepted And FixedLowerRecoveryReason = "HSD_LINEAR_QUALITY_FAILED", "KNOWN_LINEAR_FAILURE_NOT_REJECTED")
    Call Check(CompiledKey() = contextKey And VectorKey(xx) = VectorKey(saved) And ObservationKey() = observation And RPX_LowerCacheMatch() = cached, "KNOWN_LINEAR_FAILURE_ROLLBACK_CHANGED")
    RSubdivisions = 4: accepted = RPX_TryFixedLowerRecovery(0.4, field, 30): RSubdivisions = 8
    Call Check(Not accepted And CompiledKey() = contextKey, "Q4_RECOVERY_ACCEPTED")
    XStatus = "OPTIMAL": accepted = RPX_TryFixedLowerRecovery(0.4, field, 30): XStatus = savedStatus
    Call Check(Not accepted, "NONUNKNOWN_RECOVERY_ACCEPTED")
    AFActive = True: accepted = RPX_TryFixedLowerRecovery(0.4, field, 30): AFActive = False
    Call Check(Not accepted, "NESTED_AFFINE_ACCEPTED")
    accepted = RPX_TryFixedLowerRecovery(0.4, field, 60)
    Call Check(accepted, "FIXED_RECOVERY_FAILED_" & FixedLowerRecoveryReason)
    Call Check(CompiledKey() = contextKey And VectorKey(xx) = VectorKey(saved), "SUCCESS_PROGRAM_OR_ITERATE_CHANGED")
    Call Check(VectorKey(RLowerField) = lowKey And VectorKey(RUpperField) = upKey, "SUCCESS_WITNESSES_CHANGED")
    Call Check(RPX_LowerCacheMatch() = cached And XStatus = savedStatus And Not AFActive And model = RPX_ExactModel(), "SUCCESS_CACHE_STATUS_OR_MODEL_CHANGED")
    Call Check(ObservationKey() = observation, "SUCCESS_PRIVATE_HSD_OBSERVATION_CHANGED")
    Call Check(RLowerFs = 0.4 And RUpperFs = 2# And RBestLower.value = 0.4 And RBestUpper.value = 2#, "RECOVERY_PROMOTED_BOUNDS")
    Call Check(Not RPX_ActualRootBracket("lower", a, b, 0.4), "RECOVERY_CREATED_ROOT")
    Call Check(RPX_OriginalPrimalGate(field), "RESTORED_FREE_PRIMAL_FAILED")
    Call Check(RPX_FieldGate("lower", field, 0.4, 8, OriginalCandidateObjective, measured, audit, raw), "RESTORED_FREE_PHYSICAL_FAILED")
    Call Check(measured >= 1.000001 And Abs(measured - 1.000003) <= 0.0000001, "RECOVERY_LOAD_VALUE")
    recoveredValue = measured
    xx = field: Call RPX_ExportAssembled(trace & ".free_original.bin"): Call RPX_ExportExchangeX(trace & ".free_field.bin"): xx = saved
    ' A failure after solving must also retain both original witnesses and cache.
    Call RPX_TestArm("fixed_recovery_original_gate", 5, "RECOVERY_INJECTED_AUDIT_FAILURE")
    accepted = RPX_TryFixedLowerRecovery(0.4, candidate, 60): Call RPX_TestDisarm
    Call Check(Not accepted And CompiledKey() = contextKey And VectorKey(xx) = VectorKey(saved), "POST_SOLVE_ROLLBACK_FAILED")
    Call Check(VectorKey(RLowerField) = lowKey And VectorKey(RUpperField) = upKey And RPX_LowerCacheMatch() = cached, "POST_SOLVE_WITNESS_OR_CACHE_CHANGED")
    Call Check(RPX_HSDAuditMode = oldMode And RPX_HSDKind = oldKind And RPX_HSDCandidateId = oldCandidateId, "HSD_OBSERVATION_MODE_OR_KIND_CHANGED")
    Call Check(ObservationKey() = observation, "POST_SOLVE_PRIVATE_HSD_OBSERVATION_CHANGED")
    Call Check(RPX_FieldGate("lower", RLowerField, 0.4, 8, -saved(RLoadVariable), measured, audit, raw), "ROLLBACK_LOWER_PHYSICAL_FAILED")
    Call Check(RPX_FieldGate("upper", RUpperField, 2#, 8, upperObjective, measured, audit, raw), "ROLLBACK_UPPER_PHYSICAL_FAILED")
    RFsActive = False: Call RPX_DiagFinish("PASS"): Call RPX_InputEnd
    FixedLowerRecoverySmallProbe = "PASS;elements=" & re & ";Fs=0.4;q=8;c0=" & c0Fixture & ";lambda=" & recoveredValue & ";cache_before=" & cached & ";both_retained_physical_gates=PASS;full_free_primal_gate=1;independent_witnesses_compiled_layout_cache_restored=1;pre_post_solve_failure_rollback=1;q4_status_nested_rejected=1;fixed_dual_used=0;root_created=0"
End Function
