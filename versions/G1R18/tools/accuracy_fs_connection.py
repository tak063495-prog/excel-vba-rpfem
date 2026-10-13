"""Connect optional original-program DLL candidates to the audited Fs search."""


def once(source, old, new):
    if source.count(old) != 1:
        raise ValueError(f"Expected one occurrence: {old[:100]!r}")
    return source.replace(old, new, 1)


def connect(source):
    source = once(source, '    If Not stored.valid Or stored.meshKey <> FsMeshKey() Or stored.subdivisions <> RSubdivisions Then Exit Function', '''    If Not stored.valid Or stored.meshKey <> FsMeshKey() Or stored.subdivisions <> RSubdivisions Then Exit Function
    If stored.mode <> side Or Not RPX_SchurFinite(stored.target) Then Exit Function
    If side = "lower" Then
        If stored.target <> 1.000001 Then Exit Function
    Else
        If stored.target <> 0.999999 Then Exit Function
    End If''')
    source = once(source, '    enabled = CDbl(RPX_Setting("LOWER_FIXED_RECOVERY", 0))', '    enabled = CDbl(RPX_Setting("LOWER_FIXED_RECOVERY", 0))')
    source = once(source, '    FixedLowerFsEnabled = (enabled = 1)', '    FixedLowerFsEnabled = (enabled = 1 Or RPX_AccuracyEnabled())')
    source = once(source, '    Dim enabled As Boolean, recovered() As Double, seconds As Double, number As Long, source As String, message As String', '    Dim enabled As Boolean, nativeEnabled As Boolean, recovered() As Double, seconds As Double, number As Long, source As String, message As String')
    source = once(source, '    enabled = (allowRecovery And mode = "lower" And RSubdivisions = 8 And FixedLowerFsEnabled())', '''    enabled = (allowRecovery And mode = "lower" And RSubdivisions = 8 And FixedLowerFsEnabled())
    nativeEnabled = (allowRecovery And RSubdivisions = 8 And RPX_AccuracyEnabled())
    Call RPX_AccuracyGuideReset''')
    source = once(source, '    If Not enabled Then err.Raise number, source, message', '''    If Not enabled And Not nativeEnabled Then err.Raise number, source, message''')
    source = once(source, '    RFsCalls = RFsCalls + 1: FsValue = RPX_Bound(mode, factor) - target', '''    RFsCalls = RFsCalls + 1: FsValue = RPX_Bound(mode, factor) - target
    If nativeEnabled Then RPX_TestHook "fs_accuracy_original"
    If nativeEnabled And mode = "upper" And (XStatus = "OPTIMAL" Or XStatus = "INFEASIBLE") And FsValue >= 0# Then
        If Not RPX_AccuracyRecoverDual(factor, mode) Then err.Raise 5, "RPX_Fs", "HSD_NOT_CONVERGED"
        XStatus = "ORIGINAL_ACCURACY_DUAL_GUIDE": XObjective = AccuracyNativeDualObjective
        If Not RPX_AccuracyGuideCurrent(factor, mode) Then err.Raise 5, "RPX_Fs", "ACCURACY_DUAL_CONTEXT_REJECTED"
        FsValue = AccuracyNativeDualObjective - target
        Erase xx: Erase XSlack: RProgramKind = "": RProgramFs = factor
        Exit Function
    End If''')
    source = once(source, '    XStatus = "NUMERICAL_UNKNOWN": RProgramKind = ""\n    RPX_DiagEvent "fs_fixed_lower_origin"', '''    XStatus = "NUMERICAL_UNKNOWN": RProgramKind = ""
    If nativeEnabled Then
        If AccuracyFsRecover(factor, mode, target, FsValue) Then Exit Function
    End If
    If Not enabled Then err.Raise number, source, message
    RPX_DiagEvent "fs_fixed_lower_origin"''')
    source = once(source, '    seconds = CDbl(RPX_Setting("LOWER_FIXED_RECOVERY_SECONDS", 120))', '''    If CDbl(RPX_Setting("LOWER_FIXED_RECOVERY", 0)) <> 1# Then err.Raise number, source, message
    seconds = CDbl(RPX_Setting("LOWER_FIXED_RECOVERY_SECONDS", 120))''')
    source = once(source, '            If RFsActive And mode = "lower" And RSubdivisions = 8 And FixedLowerFsEnabled() Then', '''            If RFsActive And RSubdivisions = 8 And (FixedLowerFsEnabled() Or RPX_AccuracyEnabled()) Then''')
    source = once(source, '        ElseIf mode <> "upper" Or point.residual < 0# Or Not RPX_OriginalGuideCurrent(factor) Then', '''        ElseIf XStatus = "ORIGINAL_ACCURACY_UPPER_WITNESS" And mode = "upper" Then
            If point.residual > 0# Or RPX_AuditRejected Or Not RPhysicalGateAccepted Then RProgramKind = "": Exit Function
        ElseIf mode <> "upper" Or point.residual < 0# Or Not (RPX_OriginalGuideCurrent(factor) Or RPX_AccuracyGuideCurrent(factor, mode)) Then''')
    # Lower recovered witnesses and opposite signs use the existing strict state names.
    source = once(source, '            ElseIf XStatus = "ORIGINAL_LOWER_DUAL_GUIDE" Then\n                If point.residual >= 0# Then', '''            ElseIf XStatus = "ORIGINAL_LOWER_DUAL_GUIDE" Then
                If point.residual >= 0# Then''')
    source = once(source, '    If Not FsArrayReady(fieldB) Then Exit Sub', '''    If Not FsArrayReady(fieldB) Then
        If mode <> "lower" Or statusB <> "ORIGINAL_LOWER_DUAL_GUIDE" Or RSubdivisions <> 8 Then Exit Sub
    End If''')
    source = once(source, '        If mode <> "upper" Or statusA <> "ORIGINAL_FARKAS_GUIDE" Or RSubdivisions <> 8 Then Exit Sub', '''        If mode <> "upper" Or (statusA <> "ORIGINAL_FARKAS_GUIDE" And statusA <> "ORIGINAL_ACCURACY_DUAL_GUIDE") Or RSubdivisions <> 8 Then Exit Sub''')
    source = once(source, '        upperBracket.originalGuideA = (statusA = "ORIGINAL_FARKAS_GUIDE" And RSubdivisions = 8)', '''        upperBracket.originalGuideA = ((statusA = "ORIGINAL_FARKAS_GUIDE" Or statusA = "ORIGINAL_ACCURACY_DUAL_GUIDE") And RSubdivisions = 8)
        upperBracket.physicalKey = RPX_ExactModel(): upperBracket.inputEpoch = RPX_InputEpoch
        upperBracket.loadPhase = RPX_LoadPhase: upperBracket.policy = RPX_AnalysisPolicy()''')
    source = once(source, '        If stored.statusB <> "OPTIMAL" Then Exit Function', '''        If stored.statusB <> "OPTIMAL" And stored.statusB <> "ORIGINAL_ACCURACY_UPPER_WITNESS" Then Exit Function
        If stored.statusB = "ORIGINAL_ACCURACY_UPPER_WITNESS" Or stored.statusA = "ORIGINAL_ACCURACY_DUAL_GUIDE" Then
            If stored.subdivisions <> 8 Or stored.inputEpoch <> RPX_InputEpoch Or stored.physicalKey <> RPX_ExactModel() Then Exit Function
            If stored.loadPhase <> RPX_LoadPhase Or stored.policy <> RPX_AnalysisPolicy() Then Exit Function
            If Not RPX_SchurFinite(stored.a) Or Not RPX_SchurFinite(stored.b) Or Not RPX_SchurFinite(stored.fa) Or Not RPX_SchurFinite(stored.fb) Then Exit Function
            If stored.analytical Or stored.b - stored.a > 0.0001 * WorksheetFunction.Max(1#, (stored.a + stored.b) / 2#) Then Exit Function
        End If''')
    source = once(source, '            If side <> "upper" Or Not stored.originalGuideA Or stored.statusA <> "ORIGINAL_FARKAS_GUIDE" Or stored.subdivisions <> 8 Then Exit Function', '''            If side <> "upper" Or Not stored.originalGuideA Or (stored.statusA <> "ORIGINAL_FARKAS_GUIDE" And stored.statusA <> "ORIGINAL_ACCURACY_DUAL_GUIDE") Or stored.subdivisions <> 8 Then Exit Function''')
    source = once(source, '        reused = BracketUsable(lowerBracket, mode, target, meshKey, FixedLowerFsEnabled(), reason) And FsArrayReady(lowerFieldA) And FsArrayReady(lowerFieldB)', '''        reused = BracketUsable(lowerBracket, mode, target, meshKey, FixedLowerFsEnabled(), reason) And FsArrayReady(lowerFieldA) And (FsArrayReady(lowerFieldB) Or lowerBracket.originalGuideB)''')
    source = once(source, '    If XStatus <> "OPTIMAL" Then Exit Function\n    If RPX_AuditRejected Then Exit Function', '''    If XStatus <> "OPTIMAL" And XStatus <> "ORIGINAL_ACCURACY_UPPER_WITNESS" Then Exit Function
    If RPX_AuditRejected Then Exit Function''')
    source += '''
Private Function AccuracyFsRecover(ByVal factor As Double, ByVal mode As String, ByVal target As Double, ByRef residual As Double) As Boolean
    Dim recovered() As Double
    ' A separate original-dual feasibility primal is an original program dual.
    ' No auxiliary QP dual enters a root; no dual guide enters a physical field.
    If RPX_AccuracyRecoverDual(factor, mode) Then
        XObjective = AccuracyNativeDualObjective
        If mode = "upper" Then
            XStatus = "ORIGINAL_ACCURACY_DUAL_GUIDE": residual = AccuracyNativeDualObjective - target
        Else
            XStatus = "ORIGINAL_LOWER_DUAL_GUIDE": residual = -AccuracyNativeDualObjective - target
        End If
        If Not RPX_AccuracyGuideCurrent(factor, mode) Then err.Raise 5, "RPX_Fs", "ACCURACY_DUAL_CONTEXT_REJECTED"
        Erase xx: Erase XSlack
        RProgramKind = "": RProgramFs = factor
        AccuracyFsRecover = True: Exit Function
    End If
    If RPX_AccuracyRecoverPrimal(factor, mode, recovered) Then
        xx = recovered
        If mode = "lower" Then
            XObjective = -AccuracyNativeValue: Call RPX_AuditLower
            XStatus = "ORIGINAL_FIXED_LOWER_WITNESS"
        Else
            XObjective = AccuracyNativeValue: Call RPX_AuditUpper
            XStatus = "ORIGINAL_ACCURACY_UPPER_WITNESS"
        End If
        If RPX_AuditRejected Or Not RPhysicalGateAccepted Then err.Raise 5, "RPX_Fs", "ACCURACY_RECOVERY_ENDPOINT_GATE_REJECTED"
        RProgramKind = mode: RProgramFs = factor: residual = RValue - target
        AccuracyFsRecover = True
    End If
End Function
'''
    return source
