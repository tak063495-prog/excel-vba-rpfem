Option Explicit
#Const G1_THIN_UPPER = 1
' G1R1: A01 objective gate, A02 used-soil route, A04 recoverable only, A06 user GAP.
' HSD / c=0 Fs path. Python split-mesh check, same slope, phi=psi=45, quadrature 2:
'   lower on isotropic target 64 (67 elements) = 1.49999
'   upper on c=0-only bands 0.08,0.15 spacing 2.0 target 96 (186 elements) = 1.52188
'   gap 1.45%. The guided mesh does not improve the lower bound.
'   Solving the lower bound on that guided mesh hits the stable-LU budget in Excel.
' This path therefore solves the lower bound only on the coarse mesh and the upper
' bound only on the guided mesh. It does not probe Fs=1.49 on the guided mesh.
' If the guided upper hits memory, the upper falls back to the coarse mesh.
' The sheet TARGET / ADAPT / QUADRATURE cells are not rewritten.

Public Sub RPX_C0EnsureSetting()
    Dim ws As Worksheet, row As Long, found As Long, emptyRow As Long
    Set ws = RPX_Sheet("設定")
    found = 0: emptyRow = 0
    For row = 4 To 100
        If CStr(ws.cells(row, 4).Value2) = "C0_AUTO" Then found = row: Exit For
        If emptyRow = 0 Then
            If Len(CStr(ws.cells(row, 1).Value2)) = 0 And Len(CStr(ws.cells(row, 4).Value2)) = 0 Then emptyRow = row
        End If
    Next row
    If found > 0 Then
        ws.cells(found, 3).Value2 = "1=c=0を含むFsは下界を粗いメッシュ、上界だけc=0の土に層。0=従来。"
        Exit Sub
    End If
    If emptyRow = 0 Then Exit Sub
    ws.cells(emptyRow, 1).Value2 = "c=0自動メッシュ"
    ws.cells(emptyRow, 2).Value2 = 1
    ws.cells(emptyRow, 3).Value2 = "1=c=0を含むFsは下界を粗いメッシュ、上界だけc=0の土に層。0=従来。"
    ws.cells(emptyRow, 4).Value2 = "C0_AUTO"
    ws.cells(emptyRow, 2).Interior.Color = RGB(243, 248, 251)
End Sub

Public Function RPX_C0AutoRequested() As Boolean
    Dim text As String
    text = UCase$(Trim$(CStr(RPX_Setting("C0_AUTO", "1"))))
    If text = "0" Or text = "OFF" Then Exit Function
    RPX_C0AutoRequested = RPX_C0HasZeroCohesion()
End Function

Public Sub RPX_C0RunSelected(ByRef lower As Double, ByRef upper As Double)
    Dim profile As RPX_ProblemProfile
    RPX_C0Ran = True
    Call RPX_CertClear
    RAdaptRoute = "c0_auto"
    If RSubdivisions <> 2 Then
        RPX_DiagEvent "c0_quadrature", "sheet=" & RSubdivisions & ";used=2"
        RSubdivisions = 2
    End If
    Call RPX_C0SolveSplit
    RPX_BuildModelProfile profile
    RPX_DiagEvent "problem_profile", RPX_ProfileText(profile)
    RPX_LogFeatureRoutes profile
    Call RPX_WriteMesh
    Call RPX_Draw
    lower = RLowerFs
    upper = RUpperFs
End Sub

Public Sub RPX_C0SolveSplit()
    Dim lowerTarget As Long, guidedOk As Boolean, upperOk As Boolean
    RLowerFs = 0#: RUpperFs = 0#: RFsActive = False: RFsCalls = 0
    RFsSearchIncomplete = False: RFsLargestRootBracket = 0#: RFsSearchNote = ""
    RPX_C0SplitMesh = False: RPX_C0LowerElements = 0: RPX_C0UpperElements = 0
    RPX_C0TrialLog = ""
    lowerTarget = C0SolveLowerCoarse()
    If C0TryTargetFirst(lowerTarget) Then GoTo adopted
    guidedOk = C0TrySide(96, 2#, True, "upper")
    If guidedOk Then
        RPX_C0Adopted = "SPLIT lower isotropic target=" & CStr(lowerTarget) & " E" & CStr(RPX_C0LowerElements) & " / upper c0-only BAND 0.08,0.15 spacing=2 target=96 E" & CStr(RPX_C0UpperElements) & " q=2"
#If G1_THIN_UPPER Then
        Call C0TryThinUpper(lowerTarget)
#End If
    Else
        upperOk = C0TrySide(lowerTarget, 0#, False, "upper")
        If Not upperOk And lowerTarget > 48 Then upperOk = C0TrySide(48, 0#, False, "upper")
        If Not upperOk And lowerTarget > 32 Then upperOk = C0TrySide(32, 0#, False, "upper")
        If Not upperOk Then err.Raise 5, "RPX_C0Select", "C0_UPPER_EXHAUSTED"
        RPX_C0Adopted = "SPLIT lower isotropic target=" & CStr(lowerTarget) & " E" & CStr(RPX_C0LowerElements) & " / upper FALLBACK isotropic E" & CStr(RPX_C0UpperElements) & " q=2 guided_memory"
    End If
adopted:
    RPX_C0SplitMesh = True
    If Len(RPX_C0TrialLog) > 0 Then RPX_C0TrialLog = RPX_C0TrialLog & " | "
    RPX_C0TrialLog = RPX_C0TrialLog & RPX_C0Adopted
    RPX_DiagEvent "c0_adopt", RPX_C0Adopted
    If RLowerFs <= 0# Or RUpperFs <= 0# Or RLowerFs > RUpperFs Then err.Raise 5, , "FS_PAIR_INCOMPLETE"
End Sub

Public Sub RPX_C0SelectMesh()
    Dim spacings(0 To 2) As Double
    Dim i As Long, score As Long, elements As Long, bestScore As Long, bestElements As Long
    Dim bestSpacing As Double, statusL As String, statusU As String, note As String
    If Not RPX_C0Problem() Then
        Call C0SelectGeneral
        Exit Sub
    End If
    spacings(0) = 2#: spacings(1) = 1.75: spacings(2) = 2.5
    bestScore = 99: bestSpacing = 2#: bestElements = 0
    RPX_C0TrialLog = ""
    For i = 0 To 2
        If XCancel Then err.Raise 18, , "ANALYSIS_CANCELLED"
        note = C0TrySpacing(spacings(i), score, elements, statusL, statusU)
        If Len(RPX_C0TrialLog) > 0 Then RPX_C0TrialLog = RPX_C0TrialLog & " | "
        RPX_C0TrialLog = RPX_C0TrialLog & note
        RPX_DiagEvent "c0_trial", note
        If score <= 9 Then
            If score < bestScore Or (score = bestScore And Abs(spacings(i) - 2#) < Abs(bestSpacing - 2#)) Then
                bestScore = score: bestSpacing = spacings(i): bestElements = elements
            End If
        End If
    Next i
    RFsActive = False
    If bestScore <= 9 Then
        Call RPX_C0Build(96, bestSpacing, True)
        RPX_C0Adopted = "BAND 0.08,0.15 spacing=" & CStr(bestSpacing) & " target=96 q=2 elements=" & CStr(re) & " probe=" & CStr(bestScore)
    Else
        Call RPX_C0Build(64, 0#, False)
        RPX_C0Adopted = "FALLBACK isotropic target=64 q=2 elements=" & CStr(re) & " no_guided_candidate"
    End If
    RPX_DiagEvent "c0_adopt", RPX_C0Adopted & ";trials=" & RPX_C0TrialLog
End Sub

Public Sub RPX_C0Build(ByVal target As Long, ByVal spacing As Double, ByVal guides As Boolean, Optional ByVal ratios As String = "0.08,0.15")
    Dim number As Long, text As String, source As String
    On Error GoTo Failed
    DTarget = target
    C0ForceGuides = guides
    If guides Then
        C0ForceSpacing = spacing
        If Len(Trim$(ratios)) = 0 Then ratios = "0.08,0.15"
        C0ForceRatios = ratios
        C0ForceRoute = "BAND"
    Else
        C0ForceSpacing = 0#
        C0ForceRatios = ""
        C0ForceRoute = ""
    End If
    ' G1R2: previous upper solve can leave xx locked. Release before remesh.
    Call RPX_HLURelease
    Call RPX_LDLRelease
    Call RPX_SchurRelease
    Call RPX_ResetPreparedState
    Call RPX_GenerateMesh
    C0ForceGuides = False
    C0ForceRoute = ""
    RPX_LowerCacheReset
    RPX_FsBracketReset
    RProgramKind = ""
    If RR > 0 Then
        If re > 260 Or rn > 220 Then err.Raise 5, "RPX_C0Select", "C0_MESH_TOO_LARGE"
    ElseIf re > 320 Or rn > 240 Then
        err.Raise 5, "RPX_C0Select", "C0_MESH_TOO_LARGE"
    End If
    Exit Sub
Failed:
    number = err.number: text = err.description: source = err.source
    C0ForceGuides = False
    C0ForceRoute = ""
    err.Raise number, source, text
End Sub

Private Function C0SolveLowerCoarse() As Long
    Dim targets(0 To 2) As Long, i As Long
    targets(0) = 64: targets(1) = 48: targets(2) = 32
    For i = 0 To 2
        If XCancel Then err.Raise 18, , "ANALYSIS_CANCELLED"
        If C0TrySide(targets(i), 0#, False, "lower") Then
            C0SolveLowerCoarse = targets(i)
            Exit Function
        End If
    Next i
    err.Raise 5, "RPX_C0Select", "C0_LOWER_EXHAUSTED"
End Function

Private Function C0TrySide(ByVal target As Long, ByVal spacing As Double, ByVal guides As Boolean, ByVal mode As String) As Boolean
    Dim number As Long, text As String, source As String
    On Error GoTo Failed
    Call RPX_HLURelease
    Call RPX_LDLRelease
    Call RPX_C0Build(target, spacing, guides)
    Call C0SolveOne(mode)
    If mode = "lower" Then
        RPX_C0LowerElements = re
        Call RPX_CertCapture("lower", "ROOT_ENDPOINT", 2, True, "c0-lower")
    Else
        RPX_C0UpperElements = re
        Call RPX_CertCapture("upper", "ROOT_ENDPOINT", 2, True, "c0-upper")
    End If
    C0TrySide = True
    Exit Function
Failed:
    number = err.number: text = err.description: source = err.source
    RFsActive = False
    Call RPX_HLURelease
    Call RPX_LDLRelease
    If number = 18 Or InStr(1, text, "CANCEL", vbTextCompare) > 0 Then err.Raise number, source, text
    If Not C0IsCapacity(number, text) Then err.Raise number, source, text
    RPX_DiagEvent "c0_trial", mode & " target=" & CStr(target) & " guides=" & CStr(guides) & " SKIP " & text
    If Len(RPX_C0TrialLog) > 0 Then RPX_C0TrialLog = RPX_C0TrialLog & " | "
    RPX_C0TrialLog = RPX_C0TrialLog & mode & " t" & CStr(target) & " SKIP " & text
    C0TrySide = False
End Function

Private Sub C0SolveOne(ByVal mode As String)
    Dim i As Long, number As Long, text As String, source As String
    On Error GoTo Failed
    RFsActive = True
    If RPX_LoadPhase <> "reference_total" Then Call RPX_TotalReferenceLoads
    If mode = "lower" Then
        RLowerFs = RPX_FsEndpoint("lower")
        ReDim RLowerField(0 To UBound(xx))
        For i = 0 To UBound(RLowerField): RLowerField(i) = xx(i): Next i
    ElseIf mode = "upper" Then
        RUpperFs = RPX_FsEndpoint("upper")
        ReDim RUpperField(0 To 12 * re + 3 * RR - 1)
        For i = 0 To UBound(RUpperField): RUpperField(i) = xx(i): Next i
    Else
        err.Raise 5, "RPX_C0Select", "C0_SIDE_INVALID"
    End If
    If RLowerFs > 0# And RUpperFs > 0# Then
        If RLowerFs > RUpperFs Then err.Raise 5, , "FS_BOUNDS_REVERSED"
    End If
    RFsActive = False
    Exit Sub
Failed:
    number = err.number: text = err.description: source = err.source
    RFsActive = False
    err.Raise number, source, text
End Sub

Public Sub RPX_C0SolvePair()
    Dim number As Long, text As String, source As String
    On Error GoTo Failed
    Call RPX_FsPair
    Exit Sub
Failed:
    number = err.number: text = err.description: source = err.source
    If number = 18 Or InStr(1, text, "CANCEL", vbTextCompare) > 0 Then err.Raise number, source, text
    If Not C0IsCapacity(number, text) Then err.Raise number, source, text
    Resume shrink
shrink:
    Call C0ShrinkAndSolve(text)
End Sub

Private Sub C0SelectGeneral()
    ' Multiple soils, phi<>psi, or a rigid. Do not probe Fs=1.49 or 1.522.
    ' Prefer the spacing that tightened the non-associated bracket in Python.
    Dim spacings(0 To 2) As Double
    Dim i As Long, number As Long, text As String, source As String
    spacings(0) = 1.6: spacings(1) = 1.75: spacings(2) = 2#
    RPX_C0TrialLog = ""
    For i = 0 To 2
        If XCancel Then err.Raise 18, , "ANALYSIS_CANCELLED"
        On Error GoTo skipSpacing
        Call RPX_C0Build(96, spacings(i), True)
        RPX_C0TrialLog = "general c0-only BAND 0.08,0.15 spacing=" & CStr(spacings(i)) & " target=96 elements=" & CStr(re)
        RPX_C0Adopted = RPX_C0TrialLog & " q=2"
        RPX_DiagEvent "c0_adopt", RPX_C0Adopted
        Exit Sub
skipSpacing:
        number = err.number: text = err.description: source = err.source
        If number = 18 Or InStr(1, text, "CANCEL", vbTextCompare) > 0 Then err.Raise number, source, text
        If Len(RPX_C0TrialLog) > 0 Then RPX_C0TrialLog = RPX_C0TrialLog & " | "
        RPX_C0TrialLog = RPX_C0TrialLog & "s" & CStr(spacings(i)) & " SKIP " & text
        RPX_DiagEvent "c0_trial", "general spacing=" & CStr(spacings(i)) & ";skip=" & text
        Resume nextSpacing
nextSpacing:
    Next i
    On Error GoTo Failed
    Call RPX_C0Build(96, 0#, False)
    If Len(RPX_C0TrialLog) > 0 Then RPX_C0TrialLog = RPX_C0TrialLog & " | "
    RPX_C0TrialLog = RPX_C0TrialLog & "isotropic target=96 elements=" & CStr(re)
    RPX_C0Adopted = "general isotropic target=96 q=2 elements=" & CStr(re)
    RPX_DiagEvent "c0_adopt", RPX_C0Adopted & ";trials=" & RPX_C0TrialLog
    Exit Sub
Failed:
    number = err.number: text = err.description: source = err.source
    err.Raise number, source, text
End Sub

Private Sub C0ShrinkAndSolve(ByVal why As String)
    Dim targets(0 To 2) As Long, i As Long
    Dim number As Long, text As String, source As String
    targets(0) = 64: targets(1) = 48: targets(2) = 32
    RPX_DiagEvent "c0_pair_fallback", "reason=" & why
    For i = 0 To 2
        On Error GoTo nextTarget
        Call RPX_HLURelease
        Call RPX_LDLRelease
        Call RPX_C0Build(targets(i), 0#, False)
        RPX_C0Adopted = "FALLBACK isotropic target=" & CStr(targets(i)) & " q=2 elements=" & CStr(re) & " after_capacity"
        Call RPX_WriteMesh
        Call RPX_Draw
        Call RPX_FsPair
        Exit Sub
nextTarget:
        number = err.number: text = err.description: source = err.source
        If number = 18 Or InStr(1, text, "CANCEL", vbTextCompare) > 0 Then err.Raise number, source, text
        If Not C0IsCapacity(number, text) Then err.Raise number, source, text
        RPX_DiagEvent "c0_pair_fallback", "target=" & CStr(targets(i)) & ";reason=" & text
        Resume continueShrink
continueShrink:
    Next i
    err.Raise number, source, "C0_FALLBACK_EXHAUSTED " & text
End Sub

Private Function C0TrySpacing(ByVal spacing As Double, ByRef score As Long, ByRef elements As Long, ByRef statusL As String, ByRef statusU As String) As String
    Dim metric As Double, number As Long, text As String, source As String
    On Error GoTo Failed
    score = 99: elements = 0: statusL = "": statusU = "": metric = 0#
    Call RPX_C0Build(96, spacing, True)
    elements = re
    If Not C0Probe("lower", 1.49, statusL, metric) Then GoTo rejected
    If Not C0LowerCarries(statusL, metric) Then
        C0TrySpacing = "s" & spacing & " E" & elements & " lower=" & statusL & " lambda=" & metric
        RFsActive = False
        Exit Function
    End If
    If Not C0Probe("upper", 1.522, statusU, metric) Then GoTo rejected
    If statusU = "OPTIMAL" Then
        score = 1
    Else
        If Not C0Probe("upper", 1.53, statusU, metric) Then GoTo rejected
        If statusU = "OPTIMAL" Then score = 2 Else score = 9
    End If
    C0TrySpacing = "s" & spacing & " E" & elements & " lower=" & statusL & " upper=" & statusU & " score=" & score
    RFsActive = False
    Exit Function
rejected:
    score = 99
    C0TrySpacing = "s" & spacing & " E" & elements & " SKIP " & statusL & " " & statusU
    RFsActive = False
    Exit Function
Failed:
    number = err.number: text = err.description: source = err.source
    score = 99
    RFsActive = False
    Call RPX_HLURelease
    Call RPX_LDLRelease
    If number = 18 Or InStr(1, text, "CANCEL", vbTextCompare) > 0 Then err.Raise number, source, text
    C0TrySpacing = "s" & spacing & " E" & elements & " SKIP " & text
End Function

Private Function C0Probe(ByVal mode As String, ByVal factor As Double, ByRef status As String, ByRef metric As Double) As Boolean
    Dim number As Long, text As String, source As String
    On Error GoTo Failed
    If RPX_LoadPhase <> "reference_total" Then Call RPX_TotalReferenceLoads
    RFsActive = True
    metric = RPX_Bound(mode, factor)
    status = XStatus
    C0Probe = True
    Exit Function
Failed:
    number = err.number: text = err.description: source = err.source
    metric = 0#
    RProgramKind = ""
    Call RPX_HLURelease
    Call RPX_LDLRelease
    If number = 18 Or InStr(1, text, "CANCEL", vbTextCompare) > 0 Then err.Raise number, source, text
    If C0IsCapacity(number, text) Then status = "CAPACITY" Else status = "ERROR " & text
    C0Probe = False
End Function

Private Function C0LowerCarries(ByVal status As String, ByVal metric As Double) As Boolean
    If status = "UNBOUNDED" Then C0LowerCarries = True: Exit Function
    If status = "OPTIMAL" And metric >= 1# - 0.000001 Then C0LowerCarries = True
End Function

Private Function C0RelativeGap(ByVal lo As Double, ByVal hi As Double) As Double
    If Abs(lo) + Abs(hi) <= 0# Then
        C0RelativeGap = 1E+30
    Else
        C0RelativeGap = 2# * (hi - lo) / (Abs(hi) + Abs(lo))
    End If
End Function

Private Function G1TrialRecoverable(ByVal number As Long, ByVal text As String) As Boolean
    If number <> 5 Then Exit Function
    If InStr(1, text, "CANCEL", vbTextCompare) > 0 Then Exit Function
    If InStr(1, text, "INPUT_MISMATCH", vbTextCompare) > 0 Then Exit Function
    If C0IsCapacity(number, text) Then
        G1TrialRecoverable = True
        Exit Function
    End If
    G1TrialRecoverable = RPX_Recoverable(number, text, "g1_trial", True)
End Function

Private Function C0TryTargetFirst(ByVal lowerTarget As Long) As Boolean
    Dim trialFs As Double, gUser As Double, gAim As Double
    Dim ratios(0 To 1) As String, i As Long, j As Long
    Dim number As Long, text As String, source As String
    Dim boundValue As Double, directField() As Double, directAccepted As Boolean
    If Not RPX_C0Problem() Or RLowerFs <= 0# Then Exit Function
    gUser = CDbl(RPX_Setting("GAP", 0.01))
    If gUser <= 0# Or gUser >= 1# Then err.Raise 5, "RPX_C0Select", "G1_GAP_INVALID"
    gAim = 0.95 * gUser
    trialFs = RLowerFs * (2# + gAim) / (2# - gAim)
    If trialFs <= RLowerFs Then Exit Function
    ratios(0) = "0.04,0.08"
    ratios(1) = "0.02,0.04"
    For i = 0 To 1
        If XCancel Then err.Raise 18, , "ANALYSIS_CANCELLED"
        On Error GoTo trialFailed
        Call RPX_HLURelease
        Call RPX_LDLRelease
        Call RPX_SchurRelease
        Call RPX_ResetPreparedState
        Call RPX_C0Build(96, 2#, True, ratios(i))
        RFsActive = True
        If RPX_LoadPhase <> "reference_total" Then Call RPX_TotalReferenceLoads
        directAccepted = False
        If CLng(RPX_Setting("C0_DIRECT_UPPER", 1)) = 1 Then directAccepted = RPX_DirectC0Candidate(trialFs, directField, boundValue)
        If directAccepted Then
            RStrengthFactor = trialFs: xx = directField: RValue = boundValue: XObjective = boundValue: XStatus = "TARGET_WITNESS"
            RPX_DiagEvent "c0_direct_upper", "accepted=True;factorizations=0;root_valid=False;Fs=" & trialFs
        Else
            boundValue = RPX_Bound("upper", trialFs)
        End If
        RFsActive = False
        If Not directAccepted And Not RPX_UpperObjectiveAcceptable() Then
            RPX_DiagEvent "c0_g1", "accepted=False;stage=before_endpoint;ratios=" & ratios(i) & ";status=" & XStatus & ";objective=" & CStr(RValue) & ";bound=" & CStr(boundValue)
            GoTo nextTrial
        End If
        RUpperFs = trialFs
        ReDim RUpperField(0 To 12 * re + 3 * RR - 1)
        For j = 0 To UBound(RUpperField): RUpperField(j) = xx(j): Next j
        RPX_C0UpperElements = re
        Call RPX_CertCapture("upper", "TARGET_WITNESS", 2, False, ratios(i))
        RPX_C0Adopted = "SPLIT lower isotropic target=" & CStr(lowerTarget) & " E" & CStr(RPX_C0LowerElements) & " / B07 BAND " & ratios(i) & " trialFs=" & CStr(trialFs) & " q=2"
        RPX_DiagEvent "c0_g1", "accepted=True;stage=before_endpoint;ratios=" & ratios(i) & ";trialFs=" & CStr(trialFs) & ";gap=" & CStr(C0RelativeGap(RLowerFs, RUpperFs)) & ";elements=" & CStr(re) & ";objective=" & CStr(RValue)
        On Error GoTo 0
        C0TryTargetFirst = (C0RelativeGap(RLowerFs, RUpperFs) < gUser)
        Exit Function
trialFailed:
        number = err.number: text = err.description: source = err.source
        RFsActive = False
        If Not G1TrialRecoverable(number, text) Then err.Raise number, source, text
        RPX_DiagEvent "c0_g1", "accepted=False;stage=before_endpoint;ratios=" & ratios(i) & ";reason=" & text
        Resume nextTrial
nextTrial:
        On Error GoTo 0
        Call RPX_HLURelease
        Call RPX_LDLRelease
        Call RPX_SchurRelease
        Call RPX_ResetPreparedState
    Next i
    RUpperFs = 0#
    C0TryTargetFirst = False
End Function

Private Sub C0TryThinUpper(ByVal lowerTarget As Long)
    ' G1R1: keep the certified lower. At most two thinner c=0 bands at Ftrial.
    ' Adopt only when the usual upper objective gate also passes.
    Dim trialFs As Double, oldUpper As Double, gUser As Double, gAim As Double
    Dim ratios(0 To 1) As String, i As Long, j As Long
    Dim number As Long, text As String, source As String
    Dim accepted As Boolean, lastRatios As String
    Dim savedField() As Double, boundValue As Double
    If Not RPX_C0Problem() Then
        RPX_DiagEvent "c0_g1", "skipped=True;reason=profile"
        Exit Sub
    End If
    gUser = CDbl(RPX_Setting("GAP", 0.01))
    If gUser <= 0# Or gUser >= 1# Then err.Raise 5, "RPX_C0Select", "G1_GAP_INVALID"
    gAim = 0.95 * gUser
    If RLowerFs <= 0# Or RUpperFs <= 0# Then Exit Sub
    If C0RelativeGap(RLowerFs, RUpperFs) < gUser Then
        RPX_DiagEvent "c0_g1", "skipped=True;reason=already_within_gap;requested_gap=" & CStr(gUser) & ";current=" & CStr(C0RelativeGap(RLowerFs, RUpperFs))
        Exit Sub
    End If
    trialFs = RLowerFs * (2# + gAim) / (2# - gAim)
    If trialFs <= RLowerFs Or trialFs >= RUpperFs Then
        RPX_DiagEvent "c0_g1", "skipped=True;reason=trial_not_inside_bracket;trialFs=" & CStr(trialFs)
        Exit Sub
    End If
    oldUpper = RUpperFs
    lastRatios = "0.08,0.15"
    ReDim savedField(0 To UBound(RUpperField))
    For j = 0 To UBound(RUpperField): savedField(j) = RUpperField(j): Next j
    ratios(0) = "0.04,0.08"
    ratios(1) = "0.02,0.04"
    Call RPX_CertHoldUpper
    For i = 0 To 1
        If XCancel Then err.Raise 18, , "ANALYSIS_CANCELLED"
        On Error GoTo trialFailed
        Call RPX_HLURelease
        Call RPX_LDLRelease
        Call RPX_SchurRelease
        Call RPX_ResetPreparedState
        Call RPX_C0Build(96, 2#, True, ratios(i))
        RFsActive = True
        If RPX_LoadPhase <> "reference_total" Then Call RPX_TotalReferenceLoads
        boundValue = RPX_Bound("upper", trialFs)
        RFsActive = False
        If Not RPX_UpperObjectiveAcceptable() Then
            RPX_DiagEvent "c0_g1", "accepted=False;ratios=" & ratios(i) & ";status=" & XStatus & ";objective=" & CStr(RValue) & ";bound=" & CStr(boundValue) & ";reason=objective_gate"
            GoTo nextTrial
        End If
        If trialFs <= RLowerFs Or trialFs >= oldUpper Then
            RPX_DiagEvent "c0_g1", "accepted=False;ratios=" & ratios(i) & ";reason=order"
            GoTo nextTrial
        End If
        RUpperFs = trialFs
        ReDim RUpperField(0 To 12 * re + 3 * RR - 1)
        For j = 0 To UBound(RUpperField): RUpperField(j) = xx(j): Next j
        RPX_C0UpperElements = re
        Call RPX_CertCapture("upper", "TARGET_WITNESS", 2, False, ratios(i))
        ReDim savedField(0 To UBound(RUpperField))
        For j = 0 To UBound(RUpperField): savedField(j) = RUpperField(j): Next j
        accepted = True
        lastRatios = ratios(i)
        RPX_DiagEvent "c0_g1", "accepted=True;ratios=" & ratios(i) & ";trialFs=" & CStr(trialFs) & ";gap=" & CStr(C0RelativeGap(RLowerFs, RUpperFs)) & ";elements=" & CStr(re) & ";objective=" & CStr(RValue) & ";requested_gap=" & CStr(gUser) & ";aim_gap=" & CStr(gAim)
        If C0RelativeGap(RLowerFs, RUpperFs) < gUser Then
            RPX_C0Adopted = RPX_C0Adopted & " / G1 BAND " & lastRatios & " trialFs=" & CStr(trialFs)
            On Error GoTo 0
            Exit Sub
        End If
        GoTo nextTrial
trialFailed:
        number = err.number: text = err.description: source = err.source
        RFsActive = False
        If Not G1TrialRecoverable(number, text) Then err.Raise number, source, text
        RPX_DiagEvent "c0_g1", "accepted=False;ratios=" & ratios(i) & ";reason=" & text
        Resume nextTrial
nextTrial:
        On Error GoTo 0
    Next i
    If accepted Then
        RPX_C0Adopted = RPX_C0Adopted & " / G1 BAND " & lastRatios & " trialFs=" & CStr(trialFs) & " gap_still>=" & CStr(C0RelativeGap(RLowerFs, RUpperFs))
        Exit Sub
    End If
    On Error GoTo restoreFailed
    Call RPX_HLURelease
    Call RPX_LDLRelease
    Call RPX_SchurRelease
    Call RPX_ResetPreparedState
    Call RPX_CertRestoreHeldUpper
    Exit Sub
restoreFailed:
    err.Raise err.number, err.source, err.description
End Sub

Private Function C0IsCapacity(ByVal number As Long, ByVal text As String) As Boolean
    Dim token As String
    If number <> 5 Or XCancel Then Exit Function
    token = Split(Trim$(text), " ")(0)
    Select Case token
        Case "C0_MESH_TOO_LARGE", "HSD_STABLE_UNAVAILABLE_MEMORY"
            C0IsCapacity = True
    End Select
End Function
