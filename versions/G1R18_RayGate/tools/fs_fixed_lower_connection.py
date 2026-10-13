"""Optional same-Fs primal recovery and separate original lower dual sign."""


def connect(source, once):
    source = once(source, '    originalGuideA As Boolean\nEnd Type', '''    originalGuideA As Boolean
    originalGuideB As Boolean
    recoveredWitnessA As Boolean
    physicalKey As String
    inputEpoch As Long
    loadPhase As String
    policy As String
End Type''')
    source = once(source, 'Private Function FsValue(ByVal factor As Double, ByVal mode As String, ByVal target As Double) As Double\n    RFsCalls', '''Private Function FixedLowerFsEnabled() As Boolean
    Dim enabled As Double
    enabled = CDbl(RPX_Setting("LOWER_FIXED_RECOVERY", 0))
    If Not RPX_SchurFinite(enabled) Then err.Raise 5, "RPX_Fs", "INVALID_LOWER_FIXED_RECOVERY"
    If enabled <> 0 And enabled <> 1 Then err.Raise 5, "RPX_Fs", "INVALID_LOWER_FIXED_RECOVERY"
    FixedLowerFsEnabled = (enabled = 1)
End Function
Private Function FsValue(ByVal factor As Double, ByVal mode As String, ByVal target As Double, Optional ByVal allowRecovery As Boolean = True) As Double
    Dim enabled As Boolean, recovered() As Double, seconds As Double, number As Long, source As String, message As String
    enabled = (allowRecovery And mode = "lower" And RSubdivisions = 8 And FixedLowerFsEnabled())
    On Error GoTo Failed
    RFsCalls''')
    source = once(source, '    If XStatus = "ORIGINAL_FARKAS_GUIDE" Then\n', '''    If enabled Then
        RPX_TestHook "fs_lower_original"
        If FsValue < 0# Then
            If Not RPX_OriginalLowerDualGate(factor, target, True) Then err.Raise 5, "RPX_Fs", "HSD_NOT_CONVERGED"
            FsValue = OriginalLowerDualCap - target: XStatus = "ORIGINAL_LOWER_DUAL_GUIDE"
        End If
    End If
    If XStatus = "ORIGINAL_FARKAS_GUIDE" Then
''')
    source = once(source, '    End If\nEnd Function\nPublic Sub RPX_FsBracketReset()', '''    End If
    Exit Function
Failed:
    number = err.number: source = err.source: message = err.description
    On Error GoTo 0
    If number = 18 Then err.Raise number, source, message
    If Not enabled Then err.Raise number, source, message
    If Not FsNumericRecoverable(number, message) Then err.Raise number, source, message
    XStatus = "NUMERICAL_UNKNOWN": RProgramKind = ""
    RPX_DiagEvent "fs_fixed_lower_origin", "Fs=" & factor & ";reason=" & message & ";positive_primal_or_independent_original_dual_only=1"
    If RPX_OriginalLowerDualGate(factor, target, True) Then
        FsValue = OriginalLowerDualCap - target: XStatus = "ORIGINAL_LOWER_DUAL_GUIDE"
        RProgramKind = "lower": RProgramFs = factor
        Exit Function
    End If
    seconds = CDbl(RPX_Setting("LOWER_FIXED_RECOVERY_SECONDS", 120))
    If Not RPX_SchurFinite(seconds) Then err.Raise 5, "RPX_Fs", "INVALID_LOWER_FIXED_RECOVERY_SECONDS"
    If seconds <= 0# Or seconds > 600# Then err.Raise 5, "RPX_Fs", "INVALID_LOWER_FIXED_RECOVERY_SECONDS"
    If RPX_TryFixedLowerRecovery(factor, recovered, seconds, True) Then
        xx = recovered: XObjective = -FixedLowerRecoveryValue
        RStrengthFactor = factor: Call RPX_AuditLower
        If RPX_AuditRejected Or Not RPhysicalGateAccepted Or RValue < target Then err.Raise 5, "RPX_Fs", "FIXED_RECOVERY_ENDPOINT_GATE_REJECTED"
        XStatus = "ORIGINAL_FIXED_LOWER_WITNESS": RProgramKind = "lower": RProgramFs = factor
        FsValue = RValue - target
        RPX_DiagEvent "fs_fixed_lower_witness", "Fs=" & factor & ";lambda=" & RValue & ";q=8;fixed_dual_used=0;optimum_claimed=0;root_requires_separate_original_negative_dual=1"
        Exit Function
    End If
    err.Raise number, source, message
End Function
Public Sub RPX_FsBracketReset()''')
    source = once(source, '''    If XStatus <> "OPTIMAL" Then
        If mode <> "upper" Or point.residual < 0# Or Not RPX_OriginalGuideCurrent(factor) Then
            RProgramKind = "": Exit Function
        End If
    End If''', '''    If XStatus <> "OPTIMAL" Then
        If mode = "lower" And RSubdivisions = 8 And FixedLowerFsEnabled() Then
            If XStatus = "ORIGINAL_FIXED_LOWER_WITNESS" Then
                If point.residual < 0# Or RPX_AuditRejected Or Not RPhysicalGateAccepted Then RProgramKind = "": Exit Function
            ElseIf XStatus = "ORIGINAL_LOWER_DUAL_GUIDE" Then
                If point.residual >= 0# Then RProgramKind = "": Exit Function
            Else
                RProgramKind = "": Exit Function
            End If
        ElseIf mode <> "upper" Or point.residual < 0# Or Not RPX_OriginalGuideCurrent(factor) Then
            RProgramKind = "": Exit Function
        End If
    End If''')
    source = once(source, '''    If stored.statusB <> "OPTIMAL" Then Exit Function
    If stored.statusA <> "OPTIMAL" Then
        If side <> "upper" Or Not stored.originalGuideA Or stored.statusA <> "ORIGINAL_FARKAS_GUIDE" Or stored.subdivisions <> 8 Then Exit Function
    End If''', '''    If stored.recoveredWitnessA Or stored.originalGuideB Then
        If side <> "lower" Or stored.subdivisions <> 8 Then Exit Function
        If Not stored.originalGuideB Or stored.statusB <> "ORIGINAL_LOWER_DUAL_GUIDE" Then Exit Function
        If stored.statusA <> "OPTIMAL" And Not (stored.recoveredWitnessA And stored.statusA = "ORIGINAL_FIXED_LOWER_WITNESS") Then Exit Function
        If stored.inputEpoch <> RPX_InputEpoch Or stored.physicalKey <> RPX_ExactModel() Then Exit Function
        If stored.loadPhase <> RPX_LoadPhase Or stored.policy <> RPX_AnalysisPolicy() Then Exit Function
        If Not RPX_SchurFinite(stored.a) Or Not RPX_SchurFinite(stored.b) Or Not RPX_SchurFinite(stored.fa) Or Not RPX_SchurFinite(stored.fb) Then Exit Function
        If stored.analytical Or stored.b - stored.a > 0.0001 * WorksheetFunction.Max(1#, (stored.a + stored.b) / 2#) Then Exit Function
    Else
        If stored.statusB <> "OPTIMAL" Then Exit Function
        If stored.statusA <> "OPTIMAL" Then
            If side <> "upper" Or Not stored.originalGuideA Or stored.statusA <> "ORIGINAL_FARKAS_GUIDE" Or stored.subdivisions <> 8 Then Exit Function
        End If
    End If''')
    source = once(source, '        lowerBracket.statusA = statusA: lowerBracket.statusB = statusB', '''        lowerBracket.statusA = statusA: lowerBracket.statusB = statusB
        lowerBracket.recoveredWitnessA = (statusA = "ORIGINAL_FIXED_LOWER_WITNESS" And RSubdivisions = 8)
        lowerBracket.originalGuideB = (statusB = "ORIGINAL_LOWER_DUAL_GUIDE" And RSubdivisions = 8)
        lowerBracket.physicalKey = RPX_ExactModel(): lowerBracket.inputEpoch = RPX_InputEpoch
        lowerBracket.loadPhase = RPX_LoadPhase: lowerBracket.policy = RPX_AnalysisPolicy()''')
    source = once(source, 'reused = BracketUsable(lowerBracket, mode, target, meshKey, False, reason) And FsArrayReady(lowerFieldA) And FsArrayReady(lowerFieldB)', '''reused = BracketUsable(lowerBracket, mode, target, meshKey, FixedLowerFsEnabled(), reason) And FsArrayReady(lowerFieldA) And FsArrayReady(lowerFieldB)
        If reused And FixedLowerFsEnabled() Then
            reused = lowerBracket.originalGuideB And lowerBracket.inputEpoch = RPX_InputEpoch And lowerBracket.physicalKey = RPX_ExactModel() And lowerBracket.loadPhase = RPX_LoadPhase And lowerBracket.policy = RPX_AnalysisPolicy()
        End If''')
    source = once(source, 'fa = FsValue(1#, mode, target)', 'fa = FsValue(1#, mode, target, False)')
    source = once(source, '        Else\n            err.Raise 5, , ', '''        Else
            If RFsActive And mode = "lower" And RSubdivisions = 8 And FixedLowerFsEnabled() Then
                If XStatus = "NUMERICAL_UNKNOWN" Or XStatus = "NUMERICAL_FAILURE" Then err.Raise 5, "RPX_Fs", "HSD_NOT_CONVERGED"
            End If
            err.Raise 5, , ''')
    return source
