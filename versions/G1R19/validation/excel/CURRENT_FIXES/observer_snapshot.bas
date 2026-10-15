Attribute VB_Name = "OwnedProductionProbe"
Option Explicit
Private checkpointDirectory As String, checkpointCount As Long
Public OwnedQualityArmed As Boolean, OwnedObservedQ As Long, OwnedObservedTolerance As Double

Public Sub OwnedObserveQuality(ByVal q As Long, ByVal tolerance As Double)
    If Not OwnedQualityArmed Then Exit Sub
    OwnedObservedQ = q: OwnedObservedTolerance = tolerance
    Err.Raise 18, "OwnedQualityProbe", "OWNED_QUALITY_CAPTURE"
End Sub

Public Sub OwnedCheckpointStart(ByVal directory As String)
    checkpointDirectory = directory: checkpointCount = 0
End Sub

' Test-copy-only observer. It never restores, audits, solves or edits a state.
Public Sub OwnedProductionCheckpoint(ByRef state As RPX_MeshState, ByVal side As String)
    Dim prefix As String, f As Integer, i As Long, item As Variant, j As Long
    On Error GoTo Failed
    If Len(checkpointDirectory) = 0 Or Not state.complete Then Exit Sub
    If side <> "lower" And side <> "upper" Then Exit Sub
    checkpointCount = checkpointCount + 1
    prefix = checkpointDirectory & "\" & Format$(checkpointCount, "00000") & "_" & side
    Call AccuracyExport.BatchState(state, prefix & ".bin")
    f = FreeFile: Open prefix & ".meta.tmp" For Output As #f
    Print #f, "schema=OWN18C01"
    Print #f, "side=" & side
    Print #f, "value_exact=" & RPX_ExactDouble(state.value)
    Print #f, "objective_exact=" & RPX_ExactDouble(state.objective)
    Print #f, "raw_yield_exact=" & RPX_ExactDouble(state.rawYield)
    Print #f, "raw_equilibrium_exact=" & RPX_ExactDouble(state.rawEquilibrium)
    Print #f, "raw_traction_exact=" & RPX_ExactDouble(state.rawTraction)
    Print #f, "raw_boundary_exact=" & RPX_ExactDouble(state.rawBoundary)
    Print #f, "q=" & state.q & ";nodes=" & state.nodes & ";elements=" & state.elements
    Print #f, "lower_root_valid=" & state.lowerRootValid & ";upper_root_valid=" & state.upperRootValid & ";search_incomplete=" & state.searchIncomplete
    Print #f, "lower_root_a_exact=" & RPX_ExactDouble(state.lowerRootA)
    Print #f, "lower_root_b_exact=" & RPX_ExactDouble(state.lowerRootB)
    Print #f, "upper_root_a_exact=" & RPX_ExactDouble(state.upperRootA)
    Print #f, "upper_root_b_exact=" & RPX_ExactDouble(state.upperRootB)
    Print #f, "policy=" & state.policy
    Print #f, "epoch=" & state.inputEpoch & ";phase=" & state.phase & ";source=" & state.source
    Print #f, "original_model=" & state.originalModel
    Print #f, "lower_load_variable=" & state.lowerLoadVariable
    For i = 0 To state.elements - 1
        Print #f, "lower_stress_start=" & i & "," & state.lowerStressStart(i)
    Next i
    If Not state.lowerReactions Is Nothing Then
        For Each item In state.lowerReactions
            For j = 0 To 4: Print #f, "reaction_part=" & j & "," & RPX_ExactDouble(CDbl(item(j))): Next j
        Next item
    End If
    Close #f
    Name prefix & ".meta.tmp" As prefix & ".meta.tsv"
    Exit Sub
Failed:
    On Error Resume Next
    Close #f
    RPX_DiagEvent "owned_checkpoint_failed", "path=" & prefix
End Sub

Public Sub OwnedSetting(ByVal key As String, ByVal value As Variant)
    Dim ws As Worksheet, r As Long
    Set ws = RPX_Sheet("設定")
    For r = 4 To 1000
        If CStr(ws.Cells(r, 4).Value2) = key Then ws.Cells(r, 2).Value2 = value: Exit Sub
    Next r
    For r = 4 To 1000
        If Len(CStr(ws.Cells(r, 4).Value2)) = 0 And Len(CStr(ws.Cells(r, 2).Value2)) = 0 Then ws.Cells(r, 4).Value2 = key: ws.Cells(r, 2).Value2 = value: Exit Sub
    Next r
    Err.Raise 5, , "SETTING_SLOT"
End Sub

Public Function OwnedCachePreflight(ByVal trace As String) As String
    Dim ignored As Double
    On Error GoTo Failed
    RPX_Setup
    OwnedSetting "TARGET", 16: OwnedSetting "QUADRATURE", 4: OwnedSetting "RAW_LOWER_AUDIT", 1
    OwnedSetting "ANALYSIS_POLICY", "ASSOCIATED": OwnedSetting "LOWER_MODEL_CACHE", 1
    OwnedSetting "HSD_SOLVE_SECONDS", 30: OwnedSetting "ACCURACY_DLL_RECOVERY", 0
    RPX_Sheet("浸透入力").Range("D4").Value2 = 0
    RPX_InputBegin: RPX_ReadInputs: RPX_GenerateMesh: RPX_TotalReferenceLoads
    XLogPath = trace: RPX_DiagStart "current_cache_preflight": RFsActive = True
    ignored = RPX_Bound("lower", 0.4)
    If Not RPX_LowerCacheMatch() Or XOriginalCompiled.subdivisions <> 4 Then Err.Raise 5, , "Q4_CACHE_NOT_EXERCISED"
    RSubdivisions = 8
    If RPX_LowerCacheMatch() Then Err.Raise 5, , "Q4_EVIDENCE_REUSED_AS_Q8"
    ignored = RPX_Bound("lower", 0.4)
    If XOriginalCompiled.subdivisions <> 8 Then Err.Raise 5, , "PREPARED_Q4_NOT_REBUILT"
    RProgramKind = ""
    ignored = RPX_Bound("lower", 0.4)
    If Not RPX_LowerCacheMatch() Or XOriginalCompiled.subdivisions <> 8 Then Err.Raise 5, , "FRESH_Q8_CACHE_REJECTED"
    RFsActive = False: RPX_DiagFinish "PASS", "q4_drop_q8_rebuild_q8_hit"
    OwnedCachePreflight = "PASS;q4_cache_hit=1;q4_to_q8_drop=1;prepared_q8_rebuild=1;q8_to_q8_hit=1"
    Exit Function
Failed:
    OwnedCachePreflight = "FAIL " & Err.Number & " " & Err.Description
End Function

Public Function OwnedProductionRun(ByVal trace As String, ByVal directory As String) As Variant
    Dim target As Boolean, measured As Double, audit As Double, raw() As Double
    On Error GoTo Failed
    OwnedCheckpointStart directory
    XLogPath = trace: RPX_NativeSetMode "auto"
    RPX_Run
    If Not RBestLower.complete Or Not RBestUpper.complete Then
        OwnedProductionRun = Array("PARTIAL", RPX_LastError): Exit Function
    End If
    target = RPX_TargetMet("Fs", RBestLower.value, RBestUpper.value, RPX_AdaptPairComplete())
    RPX_RestoreWitnessState RBestLower
    If Not RPX_FieldGate("lower", RBestLower.field, RBestLower.value, RBestLower.q, RBestLower.objective, measured, audit, raw) Then Err.Raise 5, , "FINAL_LOWER_ORIGINAL_GATE_FAILED"
    RPX_RestoreWitnessState RBestUpper
    If Not RPX_FieldGate("upper", RBestUpper.field, RBestUpper.value, RBestUpper.q, RBestUpper.objective, measured, audit, raw) Then Err.Raise 5, , "FINAL_UPPER_ORIGINAL_GATE_FAILED"
    OwnedProductionCheckpoint RBestLower, "lower"
    OwnedProductionCheckpoint RBestUpper, "upper"
    OwnedProductionRun = Array("AUDITED", target, RBestLower.value, RBestUpper.value, RBestLower.elements, RBestUpper.elements, RBestLower.q, RBestUpper.q, RBestLower.lowerRootValid, RBestUpper.upperRootValid, RBestLower.lowerRootA, RBestLower.lowerRootB, RBestUpper.upperRootA, RBestUpper.upperRootB, RFsSearchIncomplete, RPX_LastError, RPX_LastFailureNumber, RPX_DiagPath)
    Exit Function
Failed:
    OwnedProductionRun = Array("FAIL", Err.Number, Err.Description)
End Function
