Attribute VB_Name = "RPX_Fs"

' version=20260926_0945 engine=20260926_p2s_f01_f07
Option Explicit
' 20260913 speed: secant on recent points plus periodic bisection. Tolerance and endpoint audit unchanged.
Public RLowerFs As Double, RUpperFs As Double, RUpperField() As Double, RLowerField() As Double
Public RFsCalls As Long, RFsTolerance As Double
Public RFsActive As Boolean
Public RFsSearchIncomplete As Boolean, RFsLargestRootBracket As Double, RFsSearchNote As String
Private Type FsBracketState
    valid As Boolean
    analytical As Boolean
    mode As String
    meshKey As String
    subdivisions As Long
    target As Double
    a As Double
    b As Double
    fa As Double
    fb As Double
    objectiveA As Double
    objectiveB As Double
    statusA As String
    statusB As String
End Type
Private Type FsTrialPoint
    attempted As Boolean
    valid As Boolean
    factor As Double
    residual As Double
    objective As Double
    status As String
    errorNumber As Long
    errorSource As String
    errorText As String
End Type
Private upperBracket As FsBracketState, lowerBracket As FsBracketState
Private upperFieldA() As Double, upperFieldB() As Double
Private lowerFieldA() As Double, lowerFieldB() As Double
Private Function TryFsValue(ByVal factor As Double, ByVal mode As String, ByVal target As Double, ByRef value As Double) As Boolean
    Dim message As String, source As String, number As Long, errorLine As Long
    On Error GoTo Failed
    RPX_TestHook "fs_value"
    value = FsValue(factor, mode, target): TryFsValue = True: Exit Function
Failed:
    number = err.number: message = err.description: source = err.source: errorLine = Erl
    RPX_ErrorEvidence number, source, message, errorLine, "fs_trial"
    On Error GoTo 0
    If number = 18 Then err.Raise 18, , "ANALYSIS_CANCELLED"
    If Not FsNumericRecoverable(number, message) Then err.Raise number, source, message
    RFsSearchNote = mode & " F=" & factor & " " & message
End Function
Private Function FsNumericRecoverable(ByVal number As Long, ByVal message As String) As Boolean
    If Not XCancel And number = 5 And message = "HSD_DD_PRODUCT_UNDERFLOW" Then
        ' A failed point is UNKNOWN, never a sign/field/mesh proof. Retry only inside Fs search.
        XStatus = "NUMERICAL_UNKNOWN": RProgramKind = ""
        RPX_DiagEvent "fs_numerical_unknown", "reason=" & message & ";endpoints_unchanged=True;phase=fs_trial"
        FsNumericRecoverable = True
    Else
        FsNumericRecoverable = RPX_Recoverable(number, message, "fs_trial", False)
    End If
End Function
Private Function TryFsTrialPoint(ByVal factor As Double, ByVal mode As String, ByVal target As Double, ByRef point As FsTrialPoint) As Boolean
    Dim number As Long, message As String, source As String, errorLine As Long
    point.attempted = True
    point.valid = False
    point.factor = factor
    point.residual = 0#
    point.objective = 0#
    point.status = ""
    point.errorNumber = 0
    point.errorSource = ""
    point.errorText = ""
    On Error GoTo trialFailed
    point.residual = FsValue(factor, mode, target)
    RPX_TestHook "fs_trial"
    RPX_TestTrialResult point.residual
    point.objective = XObjective
    point.status = XStatus
    If XStatus <> "OPTIMAL" Then
        RProgramKind = ""
        Exit Function
    End If
    point.valid = True
    TryFsTrialPoint = True
    Exit Function
trialFailed:
    number = err.number: message = err.description: source = err.source: errorLine = Erl
    RPX_ErrorEvidence number, source, message, errorLine, "fs_trial"
    On Error GoTo 0
    point.errorNumber = number: point.errorSource = source: point.errorText = message: point.status = XStatus
    If number = 18 Then err.Raise 18, , "ANALYSIS_CANCELLED"
    If Not FsNumericRecoverable(number, message) Then err.Raise number, source, message
    RProgramKind = ""
    RFsSearchNote = mode & " trial F=" & factor & " " & message
End Function

Public Sub RPX_TotalReferenceLoads()
    Dim e As Long, j As Long, id As Long, rid As Long
    RPX_AssertInputs "reference_load_transform"
    RProgramKind = ""
    For e = 0 To re - 1
        For j = 0 To 1: RBody1(e, j) = RBody1(e, j) + RBody0(e, j): RBody0(e, j) = 0#: Next j
    Next e
    For rid = 0 To RR - 1
        For j = 0 To 2: RRigids(rid).load1(j) = RRigids(rid).load1(j) + RRigids(rid).load0(j): RRigids(rid).load0(j) = 0#: Next j
    Next rid
    For id = 0 To RNEdge - 1
        With REdges(id)
            For j = 0 To 1: .load1(j) = .load1(j) + .load0(j): .load0(j) = 0#: Next j
            If .element(1) < 0 Then RBoundary(RPX_EdgeKey(.a, .b)) = Array(.kind, .rigidId, 0#, 0#, .load1(0), .load1(1))
        End With
    Next id
    RPX_LoadPhase = "reference_total"
End Sub

Public Function RPX_UpperObjectiveAcceptable() As Boolean
    ' Same gate as RPX_FsEndpoint upper: collapse work, not mere OPTIMAL.
    Const UPPER_TARGET As Double = 0.999999
    Const UPPER_OBJ_EPS As Double = 0.00000001
    If XStatus <> "OPTIMAL" Then Exit Function
    If RPX_AuditRejected Then Exit Function
    If RValue <> RValue Then Exit Function
    If Abs(RValue) > 1E+300 Then Exit Function
    If RValue > UPPER_TARGET + UPPER_OBJ_EPS Then Exit Function
    RPX_UpperObjectiveAcceptable = True
End Function

Public Function RPX_Bound(ByVal mode As String, Optional ByVal strengthFactor As Double = 1#) As Double
    Dim errorSource As String, errorLine As Long
    Dim cacheHit As Boolean
    Dim prepared As Boolean, errorNumber As Long, errorText As String, useHSD As Boolean, e As Long, i As Long, loadScale As Double
    On Error GoTo Failed
    RPX_SchurDeactivate
    prepared = (RProgramKind = mode)
    RStrengthFactor = strengthFactor
    If mode = "lower" And Not prepared Then
        cacheHit = RPX_LowerCacheMatch()
        If cacheHit Then prepared = True
    End If
    XStatus = "RUNNING": XIteration = 0: XObjective = 0#: XPrimalResidual = 0#: XDualResidual = 0#: XGap = 0#: RPX_AuditRejected = False
    RPX_DiagSolveStart mode, strengthFactor, prepared
    If cacheHit Then Call RPX_LowerCacheRestore
    If Not prepared Then
        Call RPX_LDLRelease
        RPX_CapacityAssembly mode
    End If
    RPX_DiagBegin dpAssembly
    RPX_Stage mode & " F=" & CStr(strengthFactor)
    If prepared Then
        If mode = "upper" Then RPX_RefreshUpper RProgramFs Else Call RPX_RefreshLower
    Else
        If mode = "upper" Then Call RPX_AssembleUpper Else Call RPX_AssembleLower
    End If
    RPX_DiagEnd dpAssembly
    For e = 0 To re - 1
        If RRigidId(e) < 0 Then
            If RMaterials(RMatId(e)).cohesion = 0# Then useHSD = True: Exit For
        End If
    Next e
    RPX_OrderMode = mode
    RPX_OptimizeBound prepared, mode, useHSD
    If XStatus <> "OPTIMAL" Then
        If RFsActive And mode = "upper" And XStatus = "INFEASIBLE" Then
            RValue = 1E+30
        ElseIf RFsActive And mode = "lower" And XStatus = "UNBOUNDED" Then
            RPX_DiagEvent "hsd_outer_audit", "kind=" & RPX_HSDKind & ";id=" & RPX_HSDCandidateId & ";status=" & XStatus & ";load_before=" & xx(RLoadVariable) & ";normalization=lambda_one"
            loadScale = xx(RLoadVariable)
            If loadScale <= 0# Then err.Raise 5, , "INVALID_LOAD_RECESSION_CERTIFICATE"
            For i = 0 To XN - 1: xx(i) = xx(i) / loadScale: Next i
            RPX_TestHook "hsd_outer_audit"
            Call RPX_AuditLower: RValue = 1E+30
        Else
            err.Raise 5, , " L   ??  ?d     ?  ?   F" & XStatus
        End If
        RProgramKind = mode: RProgramFs = strengthFactor: RPX_Bound = RValue
        RPX_DiagSolveEnd XStatus: Exit Function
    End If
    If useHSD Then RPX_DiagEvent "hsd_outer_audit", "kind=" & RPX_HSDKind & ";id=" & RPX_HSDCandidateId & ";status=" & XStatus & ";normalization=tau"
    If useHSD Then RPX_TestHook "hsd_outer_audit"
    If mode = "upper" Then Call RPX_AuditUpper Else Call RPX_AuditLower
    RProgramKind = mode: RProgramFs = strengthFactor
    RPX_Bound = RValue
    If mode = "lower" Then Call RPX_LowerCacheStore
    RPX_DiagSolveEnd XStatus
    Exit Function
Failed:
    errorNumber = err.number: errorText = err.description: errorSource = err.source: errorLine = Erl
    If useHSD And InStr(errorText, "PHYSICAL_AUDIT_FAILED") > 0 Then RPX_HSDCaptureFailure errorText
    RPX_SchurRelease
    Call RPX_LowerCacheReset
    RPX_ErrorEvidence errorNumber, errorSource, errorText, errorLine, "bound"
    RProgramKind = ""
    XStatus = RPX_DiagFailure(errorNumber, errorText)
    RPX_DiagSolveEnd XStatus, errorText
    err.Raise errorNumber, errorSource, errorText
End Function

Private Function FsValue(ByVal factor As Double, ByVal mode As String, ByVal target As Double) As Double
    RFsCalls = RFsCalls + 1: FsValue = RPX_Bound(mode, factor) - target
    If XStatus = "INFEASIBLE" Or XStatus = "UNBOUNDED" Then FsValue = 1#
End Function
Public Sub RPX_FsBracketReset()
    upperBracket.valid = False
    lowerBracket.valid = False
End Sub
Private Function FsArrayReady(ByRef values() As Double) As Boolean
    Dim n As Long
    On Error GoTo Missing
    n = UBound(values)
    If n >= LBound(values) Then FsArrayReady = True
    Exit Function
Missing:
    If err.number <> 9 Then err.Raise err.number, err.source, err.description
End Function
Private Function FsMeshKey() As String
    FsMeshKey = CStr(RPX_ModelVersion())
End Function
Public Function RPX_ActualRootBracket(ByVal side As String, ByRef a As Double, ByRef b As Double, Optional ByVal expectedFs As Double = 0#) As Boolean
    Dim stored As FsBracketState
    a = 0#: b = 0#
    If side = "lower" Then
        stored = lowerBracket
    ElseIf side = "upper" Then
        stored = upperBracket
    Else
        err.Raise 5, "RPX_Fs", "ROOT_BRACKET_SIDE"
    End If
    If Not stored.valid Or stored.meshKey <> FsMeshKey() Or stored.subdivisions <> RSubdivisions Then Exit Function
    If stored.statusA <> "OPTIMAL" Or stored.statusB <> "OPTIMAL" Then Exit Function
    If stored.analytical Then
        If stored.a <= 0# Or stored.b <> stored.a Then Exit Function
    Else
        If stored.fa < 0# Or stored.fb > 0# Or stored.b <= stored.a Then Exit Function
    End If
    If expectedFs > 0# Then
        If side = "lower" And expectedFs <> stored.a Then Exit Function
        If side = "upper" And expectedFs <> stored.b Then Exit Function
    End If
    a = stored.a: b = stored.b: RPX_ActualRootBracket = True
End Function
Private Sub SaveAnalyticalRoot(ByVal side As String, ByVal factor As Double, ByVal target As Double)
    Dim stored As FsBracketState
    If XStatus <> "OPTIMAL" Or factor <= 0# Or RPX_AuditRejected Or RFsSearchIncomplete Then Exit Sub
    stored.valid = True: stored.analytical = True: stored.mode = side
    stored.meshKey = FsMeshKey(): stored.subdivisions = RSubdivisions: stored.target = target
    stored.a = factor: stored.b = factor: stored.statusA = "OPTIMAL": stored.statusB = "OPTIMAL"
    If side = "lower" Then lowerBracket = stored Else upperBracket = stored
End Sub
Private Function BracketUsable(ByRef stored As FsBracketState, ByVal mode As String, ByVal target As Double, ByVal meshKey As String, ByVal requireSubdivisions As Boolean, ByRef reason As String) As Boolean
    reason = "empty"
    If Not stored.valid Then Exit Function
    If stored.mode <> mode Then reason = "mode": Exit Function
    If stored.meshKey <> meshKey Then reason = "mesh": Exit Function
    If Abs(stored.target - target) > 0.000000000001 Then reason = "target": Exit Function
    If requireSubdivisions Then
        If stored.subdivisions <> RSubdivisions Then reason = "quadrature": Exit Function
    End If
    If stored.fa < 0# Or stored.fb > 0# Then reason = "signs": Exit Function
    If Not (stored.b > stored.a) Then reason = "order": Exit Function
    reason = ""
    BracketUsable = True
End Function
Private Sub SaveBracket(ByVal mode As String, ByVal meshKey As String, ByVal target As Double, ByVal a As Double, ByVal b As Double, ByVal fa As Double, ByVal fb As Double, ByRef fieldA() As Double, ByRef fieldB() As Double, ByVal objectiveA As Double, ByVal objectiveB As Double, ByVal statusA As String, ByVal statusB As String)
    If fa < 0# Or fb > 0# Or Not (b > a) Then Exit Sub
    If Not FsArrayReady(fieldA) Or Not FsArrayReady(fieldB) Then Exit Sub
    If mode = "lower" Then
        lowerBracket.valid = True: lowerBracket.analytical = False
        lowerBracket.mode = mode
        lowerBracket.meshKey = meshKey
        lowerBracket.subdivisions = RSubdivisions
        lowerBracket.target = target
        lowerBracket.a = a: lowerBracket.b = b: lowerBracket.fa = fa: lowerBracket.fb = fb
        lowerBracket.objectiveA = objectiveA: lowerBracket.objectiveB = objectiveB
        lowerBracket.statusA = statusA: lowerBracket.statusB = statusB
        lowerFieldA = fieldA: lowerFieldB = fieldB
    ElseIf mode = "upper" Then
        upperBracket.valid = True: upperBracket.analytical = False
        upperBracket.mode = mode
        upperBracket.meshKey = meshKey
        upperBracket.subdivisions = RSubdivisions
        upperBracket.target = target
        upperBracket.a = a: upperBracket.b = b: upperBracket.fa = fa: upperBracket.fb = fb
        upperBracket.objectiveA = objectiveA: upperBracket.objectiveB = objectiveB
        upperBracket.statusA = statusA: upperBracket.statusB = statusB
        upperFieldA = fieldA: upperFieldB = fieldB
    End If
End Sub
Private Function UpperSeedUsable(ByVal upperSeed As Double, ByVal meshKey As String, ByRef reason As String) As Boolean
    Dim profile As RPX_ProblemProfile
    RPX_BuildModelProfile profile
    If Not RPX_CanUseFeature("LOWER_FS_SEED", profile, reason) Then Exit Function
    If upperSeed <= 0# Then reason = "no_current_upper": Exit Function
    If Not BracketUsable(upperBracket, "upper", 0.999999, meshKey, True, reason) Then Exit Function
    If upperBracket.b <> upperSeed Then reason = "upper_endpoint": Exit Function
    If upperBracket.statusB <> "OPTIMAL" Or RPX_AuditRejected Or RFsSearchIncomplete Then reason = "upper_uncertified": Exit Function
    UpperSeedUsable = True
End Function
Private Function UpperLimitUsable(ByVal upperLimit As Double, ByVal meshKey As String) As Boolean
    Dim policy As String, reason As String, usedMaterials As Object, e As Long
    policy = UCase$(Trim$(CStr(RPX_Setting("FS_BRACKET_POLICY", "LEGACY"))))
    If policy = "LEGACY" Then Exit Function
    If policy <> "AUDITED_UPPER_LIMIT" Then err.Raise 5, "RPX_Fs", "INVALID_FS_BRACKET_POLICY"
    If RR > 0 Or Not RFsActive Or upperLimit <= 0# Or upperLimit > 10000# Then Exit Function
    If Not RPX_SchurFinite(upperLimit) Then Exit Function
    ' Preserve the legacy stages (including their exact fields used for remeshing).
    If UCase$(Trim$(CStr(RPX_Setting("ADAPT_MESH_POLICY", "LEGACY")))) = "LEGACY_THEN_REFINE" And re <= 800 Then Exit Function
    Set usedMaterials = RPX_Row()
    For e = 0 To re - 1
        If RRigidId(e) < 0 Then usedMaterials(CStr(RMatId(e))) = True
    Next e
    If usedMaterials.count < 2 Then Exit Function
    If Not BracketUsable(upperBracket, "upper", 0.999999, meshKey, True, reason) Then Exit Function
    If upperBracket.b <> upperLimit Or upperBracket.statusA <> "OPTIMAL" Or upperBracket.statusB <> "OPTIMAL" Then Exit Function
    If RPX_AuditRejected Or RFsSearchIncomplete Or Not FsArrayReady(upperFieldB) Then Exit Function
    UpperLimitUsable = True
End Function
Public Function RPX_FsEndpoint(ByVal mode As String, Optional ByVal hint As Double = 0#, Optional ByVal hintRatio As Double = 0.1, Optional ByVal upperSeed As Double = 0#) As Double
    Dim a As Double, b As Double, fa As Double, fb As Double, x As Double, fx As Double, target As Double, i As Long
    Dim width As Double, factor As Double, valueScale As Double
    Dim fieldA() As Double, fieldB() As Double, objectiveA As Double, objectiveB As Double, statusA As String, statusB As String
    Dim allZeroPhi As Boolean, smooth As Boolean, e As Long, j As Long, probe As Boolean, lastProbe As Boolean, lambdaA As Double, lambdaB As Double
    Dim meshKey As String, reason As String, reused As Boolean
    Dim fp As Double, tau As Double, ratio As Double, adopted As Long, fallbackUsed As Long
    Dim trialA As FsTrialPoint, trialB As FsTrialPoint
    Dim trialAFactor As Double, trialBFactor As Double
    Dim rejection As String, trialFaText As String, trialFbText As String, errText As String
    Dim fallbackA As Double, fallbackB As Double, detail As String
    Dim aAtt As String, bAtt As String, aVal As String, bVal As String
    Dim seedSource As String, seedReason As String, seedWidth As Double, seedAttempted As Boolean
    Dim seedA As Double, seedB As Double, seedFa As String, seedFb As String
    Dim seedCalls As Long, expansionCount As Long, callsBefore As Long
    Dim limitActive As Boolean, upperLimit As Double, legacyB As Double, nextB As Double
    Dim limitTrial As FsTrialPoint
    RPX_AssertInputs "fs_endpoint"
    If XCancel Then err.Raise 18, , "ANALYSIS_CANCELLED"
    If RFsTolerance <= 0# Then RFsTolerance = 0.0001
    If mode = "upper" Then target = 0.999999 Else target = 1.000001
    allZeroPhi = True: smooth = True
    For e = 0 To re - 1
        If RRigidId(e) < 0 Then
            If RMaterials(RMatId(e)).friction <> 0# Then allZeroPhi = False
            If RMaterials(RMatId(e)).cohesion = 0# Then smooth = False
        End If
    Next e
    If allZeroPhi Then
        fa = FsValue(1#, mode, target)
        If XStatus <> "OPTIMAL" Or RValue <= 0# Then err.Raise 5, , "NO_FINITE_POSITIVE_FS"
        factor = RValue / target: RStrengthFactor = factor
        If mode = "upper" Then
            XObjective = XObjective / factor: Call RPX_AuditUpper
        Else
            For j = 0 To XN - 1: xx(j) = xx(j) / factor: Next j
            XObjective = XObjective / factor: Call RPX_AuditLower
        End If
        Call SaveAnalyticalRoot(mode, factor, target)
        RPX_FsEndpoint = factor: Exit Function
    End If
    meshKey = FsMeshKey()
    If mode = "lower" Then
        limitActive = UpperLimitUsable(upperSeed, meshKey)
        If limitActive Then upperLimit = upperSeed
    End If
    seedSource = "legacy"
    reused = False: reason = ""
    If mode = "lower" Then
        reused = BracketUsable(lowerBracket, mode, target, meshKey, False, reason) And FsArrayReady(lowerFieldA) And FsArrayReady(lowerFieldB)
    Else
        reused = BracketUsable(upperBracket, mode, target, meshKey, True, reason) And FsArrayReady(upperFieldA) And FsArrayReady(upperFieldB)
    End If
    If reused Then
        seedSource = "lower_bracket"
        RPX_TestHook "cache_hit"
        If XCancel Then err.Raise 18, , "ANALYSIS_CANCELLED"
        If mode = "lower" Then
            a = lowerBracket.a: b = lowerBracket.b: fa = lowerBracket.fa: fb = lowerBracket.fb
            objectiveA = lowerBracket.objectiveA: objectiveB = lowerBracket.objectiveB
            statusA = lowerBracket.statusA: statusB = lowerBracket.statusB
            fieldA = lowerFieldA: fieldB = lowerFieldB
        Else
            a = upperBracket.a: b = upperBracket.b: fa = upperBracket.fa: fb = upperBracket.fb
            objectiveA = upperBracket.objectiveA: objectiveB = upperBracket.objectiveB
            statusA = upperBracket.statusA: statusB = upperBracket.statusB
            fieldA = upperFieldA: fieldB = upperFieldB
        End If
        RPX_DiagEvent "fs_bracket_reuse", "engine=20260926_fs_hint;ver=20260926_0630;mode=" & mode & ";a=" & a & ";b=" & b & ";fa=" & fa & ";fb=" & fb & ";quadrature=" & RSubdivisions & ";tol=" & RFsTolerance
    ElseIf mode = "upper" And reason = "quadrature" And upperBracket.valid And upperBracket.b > upperBracket.a Then
        adopted = 0: fp = 0#: fallbackUsed = 0: rejection = "not_attempted"
        trialAFactor = 0#: trialBFactor = 0#: fallbackA = upperBracket.a: fallbackB = upperBracket.b
        trialFaText = "": trialFbText = "": errText = ""
        trialA.attempted = False: trialA.valid = False: trialB.attempted = False: trialB.valid = False
        If smooth And upperBracket.a > 0# Then
            lambdaA = upperBracket.fa + target: lambdaB = upperBracket.fb + target
            If lambdaA > 0# And lambdaB > 0# And lambdaA <> lambdaB Then
                ratio = Log(target / lambdaA) * Log(upperBracket.b / upperBracket.a) / Log(lambdaB / lambdaA)
                If ratio > -40# And ratio < 40# Then
                    fp = Exp(Log(upperBracket.a) + ratio)
                    tau = RFsTolerance * fp: If fp < 1# Then tau = RFsTolerance
                    trialAFactor = fp - 0.4 * tau: trialBFactor = fp + 0.4 * tau
                    If trialAFactor > 0.000001 And trialBFactor > trialAFactor Then
                        If TryFsTrialPoint(trialAFactor, mode, target, trialA) Then
                            If trialA.valid Then
                                a = trialAFactor: fa = trialA.residual: fieldA = xx: objectiveA = trialA.objective: statusA = trialA.status
                                If TryFsTrialPoint(trialBFactor, mode, target, trialB) Then
                                    If trialB.valid Then
                                        b = trialBFactor: fb = trialB.residual: fieldB = xx: objectiveB = trialB.objective: statusB = trialB.status
                                        If fa >= 0# And fb <= 0# Then
                                            valueScale = (a + b) / 2#: If valueScale < 1# Then valueScale = 1#
                                            If (b - a) <= RFsTolerance * valueScale Then
                                                adopted = 1: rejection = "adopted"
                                            Else
                                                rejection = "width"
                                            End If
                                        Else
                                            rejection = "sign"
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        If trialA.valid Then trialFaText = CStr(trialA.residual)
        If trialB.valid Then trialFbText = CStr(trialB.residual)
        If trialA.attempted And Not trialA.valid Then
            If Len(trialA.errorText) > 0 Then rejection = "trial_a_error" Else rejection = "trial_a_status"
        ElseIf trialA.valid And trialB.attempted And Not trialB.valid Then
            If Len(trialB.errorText) > 0 Then rejection = "trial_b_error" Else rejection = "trial_b_status"
        End If
        If adopted = 0 Then
            fallbackUsed = 1
            If rejection = "not_attempted" Then rejection = "fallback"
            RProgramKind = ""
            a = fallbackA: b = fallbackB
            If a < 0.000001 Then a = 0.000001
            If b <= a Then b = a * 2#
            fa = FsValue(a, mode, target): fieldA = xx: objectiveA = XObjective: statusA = XStatus
            fb = FsValue(b, mode, target): fieldB = xx: objectiveB = XObjective: statusB = XStatus
        End If
        If trialA.attempted Then aAtt = "1" Else aAtt = "0"
        If trialB.attempted Then bAtt = "1" Else bAtt = "0"
        If trialA.valid Then aVal = "1" Else aVal = "0"
        If trialB.valid Then bVal = "1" Else bVal = "0"
        errText = trialA.errorText
        If Len(trialB.errorText) > 0 Then
            If Len(errText) > 0 Then errText = errText & "|" Else errText = ""
            errText = errText & trialB.errorText
        End If
        errText = Replace(Replace(errText, ";", ","), vbCr, " ")
        detail = "engine=20260926_fs_trial_guard;ver=20260926_0731;mode=upper;fp=" & fp
        detail = detail & ";adopted=" & adopted & ";fallback_used=" & fallbackUsed & ";rejection_reason=" & rejection
        detail = detail & ";a=" & a & ";b=" & b & ";fa=" & fa & ";fb=" & fb
        detail = detail & ";trial_a=" & trialAFactor & ";trial_b=" & trialBFactor
        detail = detail & ";trial_fa=" & trialFaText & ";trial_fb=" & trialFbText
        detail = detail & ";trial_a_attempted=" & aAtt & ";trial_b_attempted=" & bAtt
        detail = detail & ";trial_a_valid=" & aVal & ";trial_b_valid=" & bVal
        detail = detail & ";trial_a_status=" & trialA.status & ";trial_b_status=" & trialB.status
        detail = detail & ";trial_error_number=" & trialA.errorNumber & "," & trialB.errorNumber
        detail = detail & ";trial_error_source=" & Replace(trialA.errorSource & "," & trialB.errorSource, ";", ",")
        detail = detail & ";trial_error_message=" & errText
        detail = detail & ";fallback_a=" & fallbackA & ";fallback_b=" & fallbackB
        detail = detail & ";quadrature=" & RSubdivisions & ";tol=" & RFsTolerance
        RPX_DiagEvent "fs_trial_predict", detail
    Else
        If (mode = "lower" And lowerBracket.valid) Or (mode = "upper" And upperBracket.valid) Then
            RPX_DiagEvent "fs_bracket_miss", "engine=20260926_fs_hint;ver=20260926_0630;mode=" & mode & ";reason=" & reason & ";quadrature=" & RSubdivisions
        End If
        If mode = "lower" And hint <= 0# Then
            If UpperSeedUsable(upperSeed, meshKey, seedReason) Then
                seedWidth = CDbl(RPX_Setting("LOWER_SEED_WIDTH", 0.02))
                If seedWidth <= 0# Or seedWidth >= 1# Then err.Raise 5, , "INVALID_LOWER_SEED_WIDTH"
                seedA = upperSeed * (1# - seedWidth): seedB = upperSeed
                If seedA >= 0.000001 And seedB > seedA And seedB <= 10000# Then
                    seedAttempted = True: seedSource = "same_mesh_upper"
                    RPX_DiagEvent "lower_seed_begin", "engine=20260927_p3a;lower_hint_source=" & seedSource & ";trial_a=" & seedA & ";trial_b=" & seedB
                    callsBefore = RFsCalls
                    If TryFsTrialPoint(seedA, mode, target, trialA) Then
                        a = seedA: fa = trialA.residual: fieldA = xx: objectiveA = trialA.objective: statusA = trialA.status
                        seedFa = CStr(fa)
                        If TryFsTrialPoint(seedB, mode, target, trialB) Then
                            b = seedB: fb = trialB.residual: fieldB = xx: objectiveB = trialB.objective: statusB = trialB.status
                            seedFb = CStr(fb)
                        Else
                            seedReason = "trial_b_invalid"
                        End If
                    Else
                        seedReason = "trial_a_invalid"
                    End If
                    seedCalls = RFsCalls - callsBefore
                    If trialA.valid And trialB.valid Then
                        If fa < 0# Or fb > 0# Then seedReason = "sign_expand"
                    Else
                        RProgramKind = ""
                    End If
                Else
                    seedReason = "trial_range"
                End If
            End If
        End If
        If seedAttempted And trialA.valid And trialB.valid Then
            ' Both fields/residuals are LOWER solves; existing expansion handles either sign miss.
        Else
            If hint > 0# Then
                seedSource = "lower_history"
                If hintRatio <= 0# Then hintRatio = 0.1
                a = hint * (1# - hintRatio)
                b = hint * (1# + hintRatio)
                If a < 0.000001 Then a = 0.000001
                If b <= a Then b = a * 2#
            Else
                a = 1#: b = 2#
            End If
            legacyB = b
            If limitActive Then
                If b > upperLimit Then b = upperLimit
                If a >= b Then a = b / 2#
            End If
            fa = FsValue(a, mode, target): fieldA = xx: objectiveA = XObjective: statusA = XStatus
            If limitActive And b = upperLimit Then
                If TryFsTrialPoint(b, mode, target, limitTrial) Then
                    fb = limitTrial.residual
                Else
                    If InStr(1, limitTrial.errorText, "HSD_STABLE_UNAVAILABLE_MEMORY", vbBinaryCompare) = 1 Then err.Raise limitTrial.errorNumber, limitTrial.errorSource, limitTrial.errorText
                    limitActive = False: b = legacyB
                    RPX_DiagEvent "fs_upper_limit_fallback", "reason=numerical_unknown;mode=lower;legacy_b=" & b
                    fb = FsValue(b, mode, target)
                End If
            Else
                fb = FsValue(b, mode, target)
            End If
            fieldB = xx: objectiveB = XObjective: statusB = XStatus
        End If
    End If
    Do While fa < 0#
        expansionCount = expansionCount + 1
        b = a: fb = fa: fieldB = fieldA: objectiveB = objectiveA: statusB = statusA: a = a / 2#
        If a < 0.000001 Then err.Raise 5, , "FS_LOWER_BRACKET_NOT_FOUND"
        fa = FsValue(a, mode, target): fieldA = xx: objectiveA = XObjective: statusA = XStatus
    Loop
    Do While fb > 0#
        expansionCount = expansionCount + 1
        a = b: fa = fb: fieldA = fieldB: objectiveA = objectiveB: statusA = statusB
        nextB = b * 2#
        If limitActive Then
            If b >= upperLimit Then
                limitActive = False
                RPX_DiagEvent "fs_upper_limit_fallback", "reason=actual_lower_sign_positive;mode=lower;at=" & b
            ElseIf nextB > upperLimit Then
                nextB = upperLimit
            End If
        End If
        If nextB > 10000# Then err.Raise 5, , "FS_UPPER_BRACKET_NOT_FOUND"
        If limitActive And nextB = upperLimit Then
            If TryFsTrialPoint(nextB, mode, target, limitTrial) Then
                fb = limitTrial.residual
            Else
                If InStr(1, limitTrial.errorText, "HSD_STABLE_UNAVAILABLE_MEMORY", vbBinaryCompare) = 1 Then err.Raise limitTrial.errorNumber, limitTrial.errorSource, limitTrial.errorText
                limitActive = False: nextB = b * 2#
                If nextB > 10000# Then err.Raise 5, , "FS_UPPER_BRACKET_NOT_FOUND"
                RPX_DiagEvent "fs_upper_limit_fallback", "reason=numerical_unknown;mode=lower;legacy_b=" & nextB
                fb = FsValue(nextB, mode, target)
            End If
        Else
            fb = FsValue(nextB, mode, target)
        End If
        b = nextB: fieldB = xx: objectiveB = XObjective: statusB = XStatus
    Loop
    If mode = "lower" Then
        RPX_DiagEvent "fs_upper_limit", "active=" & limitActive & ";hint=" & upperLimit & ";actual_lower_sign_checks=1"
        detail = "lower_hint_source=" & seedSource & ";trial_a=" & seedA & ";trial_b=" & seedB
        detail = detail & ";fa=" & seedFa & ";fb=" & seedFb & ";trial_solve_count=" & seedCalls
        detail = detail & ";expansion_count=" & expansionCount & ";fallback_reason=" & seedReason
        detail = detail & ";a=" & a & ";b=" & b & ";target=" & target
        RPX_DiagEvent "lower_seed", detail
    End If
    For i = 1 To 40
        width = b - a: valueScale = (a + b) / 2#: If valueScale < 1# Then valueScale = 1#
        If width <= RFsTolerance * valueScale Then Exit For
        probe = False
        x = (a * fb - b * fa) / (fb - fa)
        lambdaA = fa + target: lambdaB = fb + target
        If smooth And lambdaA > 0# And lambdaB > 0# And lambdaA <> lambdaB Then
            x = Exp(Log(a) + Log(target / lambdaA) * Log(b / a) / Log(lambdaB / lambdaA))
        End If
        If RFsTolerance >= 0.001 And smooth And Not lastProbe Then
            If Abs(x - a) <= 0.5 * RFsTolerance * valueScale And Abs(x - a) <= Abs(x - b) Then
                probe = True: x = a + 0.8 * RFsTolerance * valueScale
            ElseIf Abs(x - b) <= 0.5 * RFsTolerance * valueScale Then
                probe = True: x = b - 0.8 * RFsTolerance * valueScale
            End If
        ElseIf smooth And i > 1 And Abs(fx) < 0.00001 And Not lastProbe Then
            probe = True
            If fx >= 0# Then x = a + 0.8 * RFsTolerance * valueScale Else x = b - 0.8 * RFsTolerance * valueScale
        End If
        lastProbe = probe
        If smooth Then
            If x < a + 0.000001 * width Then x = a + 0.000001 * width
            If x > b - 0.000001 * width Then x = b - 0.000001 * width
        Else
            If x < a + 0.2 * width Then x = a + 0.2 * width
            If x > b - 0.2 * width Then x = b - 0.2 * width
        End If
        If Not TryFsValue(x, mode, target, fx) Then
            x = (a + x) / 2#
            If Not TryFsValue(x, mode, target, fx) Then
                x = (b + x) / 2#
                If Not TryFsValue(x, mode, target, fx) Then Exit For
            End If
        End If
        If fx >= 0# Then
            a = x: fa = fx: fieldA = xx: objectiveA = XObjective: statusA = XStatus
        Else
            b = x: fb = fx: fieldB = xx: objectiveB = XObjective: statusB = XStatus
        End If
    Next i
    If i > 40 Then RFsSearchIncomplete = True: RFsSearchNote = "FS_ROOT_ITERATION_LIMIT"
    valueScale = (a + b) / 2#: If valueScale < 1# Then valueScale = 1#
    If b - a > RFsTolerance * valueScale Then RFsSearchIncomplete = True
    If b - a > RFsLargestRootBracket Then RFsLargestRootBracket = b - a
    If mode = "upper" Then
        factor = b: xx = fieldB: XObjective = objectiveB: XStatus = statusB
        RStrengthFactor = factor: Call RPX_AuditUpper
        If RValue > target + 0.00000001 Then err.Raise 5, , "FS_UPPER_ENDPOINT_INVALID"
    Else
        factor = a: xx = fieldA: XObjective = objectiveA: XStatus = statusA
        RStrengthFactor = factor: Call RPX_AuditLower
        If RValue < 1# - 0.00000001 Then err.Raise 5, , "FS_LOWER_ENDPOINT_INVALID"
    End If
    RPX_DiagEvent "p4_final_bracket", "mode=" & mode & ";a=" & a & ";b=" & b & ";fa=" & fa & ";fb=" & fb & ";target=" & target & ";quadrature=" & RSubdivisions & ";root_width=" & b - a
    SaveBracket mode, meshKey, target, a, b, fa, fb, fieldA, fieldB, objectiveA, objectiveB, statusA, statusB
    RPX_FsEndpoint = factor
End Function
Public Sub RPX_FsPair()
    Dim i As Long, errorNumber As Long, message As String, errorSource As String, errorLine As Long
    On Error GoTo Failed
    RFsActive = True: RFsCalls = 0: RFsSearchIncomplete = False: RFsLargestRootBracket = 0#: RFsSearchNote = "": Call RPX_TotalReferenceLoads
    RUpperFs = RPX_FsEndpoint("upper"): ReDim RUpperField(0 To 12 * re + 3 * RR - 1)
    For i = 0 To UBound(RUpperField): RUpperField(i) = xx(i): Next i
    If UpperLimitUsable(RUpperFs, FsMeshKey()) Then
        RLowerFs = RPX_FsEndpoint("lower", 0#, 0.1, RUpperFs)
    Else
        RLowerFs = RPX_FsEndpoint("lower")
    End If
    ReDim RLowerField(0 To UBound(xx))
    For i = 0 To UBound(RLowerField): RLowerField(i) = xx(i): Next i
    If RLowerFs > RUpperFs Then err.Raise 5, , "FS_BOUNDS_REVERSED"
    RFsActive = False: Exit Sub
Failed:
    errorNumber = err.number: message = err.description: errorSource = err.source: errorLine = Erl
    RPX_ErrorEvidence errorNumber, errorSource, message, errorLine, "fs_pair"
    RFsActive = False
    err.Raise errorNumber, errorSource, message
End Sub
Public Function RPX_TestFs(ByVal inputPath As String, ByVal outputPath As String) As String
    Dim f As Integer, started As Double
    On Error GoTo Failed
    started = Timer: XLogPath = outputPath & ".log": RPX_ReadModel inputPath
    Call RPX_FsPair
    f = FreeFile: Open outputPath For Output As #f
    Print #f, RLowerFs; ","; RUpperFs; ","; RFsCalls; ","; Timer - started
    Close #f: Application.StatusBar = False
    RPX_TestFs = "PASS lower=" & CStr(RLowerFs) & " upper=" & CStr(RUpperFs): Exit Function
Failed:
    RPX_TestFs = "FAIL " & err.number & " " & err.description: Application.StatusBar = False
End Function


