Attribute VB_Name = "RPX_AccuracyNative"
Option Explicit
#If Win64 Then
Private Declare PtrSafe Function NAbi Lib "rpfem_native_x64.dll" Alias "rpfem_accuracy_abi_version" () As Long
Private Declare PtrSafe Function NDefaults Lib "rpfem_native_x64.dll" Alias "rpfem_accuracy_default_settings" (ByVal settings As LongPtr) As Long
Private Declare PtrSafe Function NCreate Lib "rpfem_native_x64.dll" Alias "rpfem_accuracy_create" (ByVal problem As LongPtr, ByVal settings As LongPtr, ByRef handle As LongLong) As Long
Private Declare PtrSafe Function NStart Lib "rpfem_native_x64.dll" Alias "rpfem_accuracy_start" (ByVal handle As LongLong) As Long
Private Declare PtrSafe Function NPoll Lib "rpfem_native_x64.dll" Alias "rpfem_accuracy_poll" (ByVal handle As LongLong, ByVal progress As LongPtr) As Long
Private Declare PtrSafe Function NCancel Lib "rpfem_native_x64.dll" Alias "rpfem_accuracy_cancel" (ByVal handle As LongLong) As Long
Private Declare PtrSafe Function NResult Lib "rpfem_native_x64.dll" Alias "rpfem_accuracy_result" (ByVal handle As LongLong, ByVal x As LongPtr, ByVal nx As Long, ByVal z As LongPtr, ByVal nz As Long, ByVal progress As LongPtr) As Long
Private Declare PtrSafe Function NDestroy Lib "rpfem_native_x64.dll" Alias "rpfem_accuracy_destroy" (ByVal handle As LongLong) As Long
Private Declare PtrSafe Function NEqualities Lib "rpfem_native_x64.dll" Alias "rpfem_accuracy_equalities" (ByVal problem As LongPtr, ByVal iterations As Long, ByVal atol As Double, ByVal btol As Double, ByVal conlim As Double, ByVal correction As LongPtr, ByVal count As Long, ByVal report As LongPtr) As Long
Private Declare PtrSafe Sub CopyBytes Lib "kernel32" Alias "RtlMoveMemory" (ByVal dst As LongPtr, ByVal src As LongPtr, ByVal count As LongPtr)
Private Declare PtrSafe Sub Idle Lib "kernel32" Alias "Sleep" (ByVal milliseconds As Long)
Private Declare PtrSafe Function LoadAccuracyLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal path As LongPtr) As LongPtr
#End If
Public Type RPX_AccuracyProgram
    n As Long
    m As Long
    nnz As Long
    coneCount As Long
    equalityCount As Long
    rowPtr() As Long
    colId() As Long
    values() As Double
    rhs() As Double
    pdiag() As Double
    cost() As Double
    originalCost() As Double
    coneKind() As Long
    coneDim() As Long
    model As String
    policy As String
    phase As String
    mode As String
    epoch As Long
    q As Long
    factor As Double
End Type
Public AccuracyNativeReason As String, AccuracyNativeSolverStatus As Long, AccuracyNativeIterations As Long
Public AccuracyNativeEquality As Double, AccuracyNativeCone As Double, AccuracyNativeValue As Double, AccuracyNativeAudit As Double
Public AccuracyNativeDualStat As Double, AccuracyNativeDualCone As Double, AccuracyNativeDualObjective As Double
Public AccuracyNativeDualAudited As Boolean, AccuracyNativeDualAuditCallId As Long
Public AccuracyNativeCallId As Long
Private accuracyCallSequence As Long
Public AccuracyNativeEquilibrate As Long
Public AccuracyNativeDualAuditRole As String, AccuracyNativeDualAuditStatus As Long
Private guideReady As Boolean, guideModel As String, guidePolicy As String, guideMode As String, guidePhase As String
Private guideEpoch As Long, guideQ As Long, guideFactor As Double
Private guideN As Long, guideM As Long, guideS As Long, guideB As Long, guideSolve As Long, guideCandidate As Long
Private guideCost() As Double, guideNormalizedCost() As Double
Private libraryHandle As LongPtr, libraryPath As String
Private proofSequence As Long
Public AccuracyDirectLastAccepted As Boolean, AccuracyDirectLastKind As String
' Fix9: settings are populated and cleared only inside a thin-layer step.
Public RPX_ThinDllFactor As Double, RPX_ThinDllRefSeconds As Double
Public RPX_ThinDllRefCones As Double, RPX_ThinDllExponent As Double, RPX_ThinDllCap As Double
' rc6: populated/cleared with the existing thin-chain budget lifetime only.
Public RPX_ThinUpperDllRetry As Boolean, RPX_ThinUpperDllMinElements As Long
Public RPX_ThinUpperDllRetrySeconds As Double, RPX_ThinUpperDllRetryIterations As Long
Public RPX_ThinUpperDllRetryEquilibrate As Long

Private upperFirstReady As Boolean, upperFirstModel As String, upperFirstPolicy As String, upperFirstPhase As String
Private upperFirstFactor As Double, upperFirstEpoch As Long, upperFirstQ As Long
Private upperFirstSolve As Long, upperFirstCandidate As Long
' Diagnostics are per direct trial; zero values require dual_audited=True.
Private Sub ResetDualDiagnostics()
    AccuracyNativeDualAudited = False: AccuracyNativeDualAuditCallId = 0
    AccuracyNativeDualAuditRole = "": AccuracyNativeDualAuditStatus = 0
    AccuracyNativeDualStat = 0#: AccuracyNativeDualCone = 0#: AccuracyNativeDualObjective = 0#
End Sub
Public Function RPX_AccuracyEnabled() As Boolean
    Dim value As Double
    value = CDbl(RPX_Setting("ACCURACY_DLL_RECOVERY", 0))
    If Not RPX_SchurFinite(value) Or (value <> 0# And value <> 1#) Then err.Raise 5, "RPX_AccuracyNative", "INVALID_ACCURACY_DLL_RECOVERY"
    RPX_AccuracyEnabled = (value = 1#)
End Function
Public Sub RPX_AccuracyGuideReset()
    guideReady = False
    Erase guideCost: Erase guideNormalizedCost
End Sub
Public Function RPX_AccuracyGuideCurrent(ByVal factor As Double, ByVal mode As String) As Boolean
    Dim i As Long
    If Not guideReady Then Exit Function
    If mode = "upper" Then
        If XStatus <> "ORIGINAL_ACCURACY_DUAL_GUIDE" Then Exit Function
    ElseIf mode = "lower" Then
        If XStatus <> "ORIGINAL_LOWER_DUAL_GUIDE" Then Exit Function
    Else
        Exit Function
    End If
    If AFActive Or Not RFsActive Or RSubdivisions <> 8 Then Exit Function
    If Not RPX_SchurFinite(factor) Or factor <> guideFactor Or factor <> RStrengthFactor Then Exit Function
    If guideModel <> RPX_ExactModel() Or guidePolicy <> RPX_AnalysisPolicy() Then Exit Function
    If guideEpoch <> RPX_InputEpoch Or guideQ <> RSubdivisions Or guidePhase <> RPX_LoadPhase Or guideMode <> mode Then Exit Function
    If guideN <> XN Or guideM <> XM Or guideS <> XS Or guideB <> XB Or guideSolve <> P4Solve Or guideCandidate <> RPX_HSDCandidateId Then Exit Function
    If Not XOriginalCompiled.ready Or XOriginalCompiled.affine Then Exit Function
    For i = 0 To guideN - 1
        If Not RPX_SchurFinite(XQ(i)) Or Not RPX_SchurFinite(XOriginalCompiled.cost(i)) Then Exit Function
        If XQ(i) <> guideNormalizedCost(i) Or XOriginalCompiled.cost(i) <> guideCost(i) Then Exit Function
    Next i
    RPX_AccuracyGuideCurrent = True
End Function
Private Function CurrentProgram(ByRef pr As RPX_AccuracyProgram) As Boolean
    Dim i As Long
    If AFActive Or Not RFsActive Or Not RRawAuditEnabled Or RR <> 0 Then Exit Function
    If pr.model <> RPX_ExactModel() Or pr.epoch <> RPX_InputEpoch Or pr.q <> RSubdivisions Then Exit Function
    If pr.q <> 8 Or pr.phase <> "reference_total" Or pr.phase <> RPX_LoadPhase Then Exit Function
    If pr.factor <> RStrengthFactor Or pr.mode <> RPX_OrderMode Or pr.policy <> RPX_AnalysisPolicy() Then Exit Function
    If pr.n <> XN Or pr.equalityCount <> XM Or pr.m <> XM + XS Then Exit Function
    If Not XOriginalCompiled.ready Or XOriginalCompiled.affine Or XOriginalCompiled.inputEpoch <> pr.epoch Then Exit Function
    If Not RPX_SchurFinite(XOriginalCompiled.strengthFactor) Or XOriginalCompiled.strengthFactor <> pr.factor Then Exit Function
    If XOriginalCompiled.modelKey <> pr.model Or XOriginalCompiled.subdivisions <> pr.q Or XOriginalCompiled.modeName <> pr.mode Then Exit Function
    If Not RPX_SchurFinite(XQScale) Or XQScale <= 0# Then Exit Function
    For i = 0 To pr.n - 1
        If Not RPX_SchurFinite(XQ(i)) Or Not RPX_SchurFinite(pr.originalCost(i)) Then Exit Function
        If Not RPX_SchurFinite(XOriginalCompiled.normalizedCost(i)) Or XQ(i) <> XOriginalCompiled.normalizedCost(i) Then Exit Function
        If pr.originalCost(i) <> XOriginalCompiled.cost(i) Then Exit Function
    Next i
    CurrentProgram = True
End Function
Private Sub AddTerm(ByRef total As Double, ByRef correction As Double, ByVal value As Double)
    Dim adjusted As Double, nextValue As Double
    If Not RPX_SchurFinite(value) Then err.Raise 5, "RPX_AccuracyNative", "ACCURACY_NONFINITE_PRODUCT"
    adjusted = value - correction: nextValue = total + adjusted
    If Not RPX_SchurFinite(nextValue) Then err.Raise 5, "RPX_AccuracyNative", "ACCURACY_NONFINITE_SUM"
    correction = (nextValue - total) - adjusted: total = nextValue
End Sub
' Diagnostic subtraction must never change an audit decision on overflow.
Private Function DiagnosticAbsDifference(ByVal a As Double, ByVal b As Double, ByRef complete As Boolean) As Double
    On Error GoTo Missing
    DiagnosticAbsDifference = Abs(a - b)
    Exit Function
Missing:
    If err.number <> 6 Then err.Raise err.number, err.source, err.description
    err.Clear: complete = False
End Function
Private Function Norm2(ByVal a As Double, ByVal b As Double) As Double
    Dim positiveScale As Double
    positiveScale = WorksheetFunction.Max(Abs(a), Abs(b))
    If positiveScale > 0# Then Norm2 = positiveScale * Sqr((a / positiveScale) ^ 2 + (b / positiveScale) ^ 2)
End Function
Private Sub SortRow(ByRef pr As RPX_AccuracyProgram, ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long, id As Long, value As Double
    For i = first + 1 To last
        id = pr.colId(i): value = pr.values(i): j = i - 1
        Do While j >= first
            If pr.colId(j) <= id Then Exit Do
            pr.colId(j + 1) = pr.colId(j): pr.values(j + 1) = pr.values(j): j = j - 1
        Loop
        If j >= first Then
            If pr.colId(j) = id Then err.Raise 5, "RPX_AccuracyNative", "ACCURACY_DUPLICATE_COLUMN"
        End If
        pr.colId(j + 1) = id: pr.values(j + 1) = value
    Next i
End Sub
' Copy the literal original program before any auxiliary cap/floor/scaling.
Public Function RPX_AccuracyExport(ByVal factor As Double, ByVal mode As String, ByRef pr As RPX_AccuracyProgram) As Boolean
    Dim blank As RPX_AccuracyProgram, i As Long, j As Long, k As Long, r As Long, row As Long, p As Long, nz As Double
    pr = blank: AccuracyNativeReason = "ORIGINAL_PROGRAM_SCOPE_REJECTED"
    If XCancel Or AFActive Or Not RFsActive Or RSubdivisions <> 8 Or RR <> 0 Or Not RRawAuditEnabled Then Exit Function
    If mode <> "lower" And mode <> "upper" Then Exit Function
    If mode <> RPX_OrderMode Or RPX_LoadPhase <> "reference_total" Then Exit Function
    If Not RPX_SchurFinite(factor) Or factor <= 0# Or factor <> RStrengthFactor Then Exit Function
    If Not XOriginalCompiled.ready Or XOriginalCompiled.affine Then Exit Function
    If XOriginalCompiled.n <> XN Or XOriginalCompiled.m <> XM Or XOriginalCompiled.s <> XS Or XOriginalCompiled.b <> XB Then Exit Function
    If XOriginalCompiled.inputEpoch <> RPX_InputEpoch Or XOriginalCompiled.subdivisions <> 8 Then Exit Function
    If XOriginalCompiled.modelKey <> RPX_ExactModel() Or XOriginalCompiled.modeName <> mode Or XOriginalCompiled.loadPhase <> RPX_LoadPhase Then Exit Function
    If Not RPX_SchurFinite(XOriginalCompiled.strengthFactor) Or XOriginalCompiled.strengthFactor <> factor Then Exit Function
    If XN <= 0 Or XM <= 0 Or XS <= 0 Or XB <= 0 Then Exit Function
    If LBound(XOriginalCompiled.cost) <> 0 Or UBound(XOriginalCompiled.cost) <> XN - 1 Then Exit Function
    If mode = "lower" And (RLoadVariable < 0 Or RLoadVariable >= XN) Then Exit Function
    For i = 0 To XN - 1
        If Not RPX_SchurFinite(XOriginalCompiled.cost(i)) Then Exit Function
        If mode = "lower" Then
            If i = RLoadVariable Then
                If XOriginalCompiled.cost(i) <> -1# Then Exit Function
            Else
                If XOriginalCompiled.cost(i) <> 0# Then Exit Function
            End If
        End If
    Next i
    nz = XEPtr(XM)
    For k = 0 To XB - 1
        If XC(k).dimn <> 1 And XC(k).dimn <> 3 Then Exit Function
        For r = 0 To XC(k).dimn - 1
            If Not RPX_SchurFinite(XC(k).offset(r)) Then Exit Function
            For j = 0 To XC(k).count - 1
                If XC(k).ids(j) < 0 Or XC(k).ids(j) >= XN Or Not RPX_SchurFinite(XC(k).coef(r, j)) Then Exit Function
                If XC(k).coef(r, j) <> 0# Then nz = nz + 1#
            Next j
        Next r
    Next k
    If nz <= 0# Or nz > 2147483000# Then err.Raise 5, "RPX_AccuracyNative", "ACCURACY_SPARSE_LIMIT"
    pr.n = XN: pr.m = XM + XS: pr.nnz = CLng(nz): pr.equalityCount = XM: pr.coneCount = XB + 1
    pr.originalCost = XOriginalCompiled.cost: pr.cost = pr.originalCost
    pr.model = RPX_ExactModel(): pr.policy = RPX_AnalysisPolicy(): pr.epoch = RPX_InputEpoch
    pr.q = RSubdivisions: pr.factor = factor: pr.mode = mode: pr.phase = RPX_LoadPhase
    ReDim pr.rowPtr(0 To pr.m): ReDim pr.colId(0 To pr.nnz - 1): ReDim pr.values(0 To pr.nnz - 1)
    ReDim pr.rhs(0 To pr.m - 1): ReDim pr.pdiag(0 To pr.n - 1)
    ReDim pr.coneKind(0 To pr.coneCount - 1): ReDim pr.coneDim(0 To pr.coneCount - 1)
    pr.coneKind(0) = 0: pr.coneDim(0) = XM
    For i = 0 To XM - 1
        pr.rowPtr(i) = p: pr.rhs(i) = XOriginalCompiled.rhs(i)
        If Not RPX_SchurFinite(pr.rhs(i)) Then Exit Function
        For j = XEPtr(i) To XEPtr(i + 1) - 1
            If XEId(j) < 0 Or XEId(j) >= XN Or Not RPX_SchurFinite(XOriginalCompiled.values(j)) Then Exit Function
            pr.colId(p) = XEId(j): pr.values(p) = XOriginalCompiled.values(j): p = p + 1
        Next j
        SortRow pr, pr.rowPtr(i), p - 1
    Next i
    row = XM
    For k = 0 To XB - 1
        pr.coneKind(k + 1) = XC(k).dimn: pr.coneDim(k + 1) = XC(k).dimn
        If XC(k).first <> row - XM Then Exit Function
        For r = 0 To XC(k).dimn - 1
            pr.rowPtr(row) = p: pr.rhs(row) = XC(k).offset(r)
            For j = 0 To XC(k).count - 1
                If XC(k).coef(r, j) <> 0# Then
                    pr.colId(p) = XC(k).ids(j): pr.values(p) = -XC(k).coef(r, j): p = p + 1
                End If
            Next j
            SortRow pr, pr.rowPtr(row), p - 1: row = row + 1
        Next r
    Next k
    pr.rowPtr(row) = p
    If p <> pr.nnz Or row <> pr.m Or Not CurrentProgram(pr) Then Exit Function
    RPX_AccuracyExport = True: AccuracyNativeReason = "ORIGINAL_PROGRAM_EXPORTED"
End Function
Private Sub PackLong(ByRef bytes() As Byte, ByVal offset As Long, ByVal value As Long)
    CopyBytes VarPtr(bytes(offset)), VarPtr(value), 4
End Sub
Private Sub PackPointer(ByRef bytes() As Byte, ByVal offset As Long, ByVal value As LongPtr)
    CopyBytes VarPtr(bytes(offset)), VarPtr(value), 8
End Sub
Private Sub PackDouble(ByRef bytes() As Byte, ByVal offset As Long, ByVal value As Double)
    CopyBytes VarPtr(bytes(offset)), VarPtr(value), 8
End Sub
Private Function ReadLong(ByRef bytes() As Byte, ByVal offset As Long) As Long
    Dim value As Long
    CopyBytes VarPtr(value), VarPtr(bytes(offset)), 4: ReadLong = value
End Function
' Literal original LP identity; inherited mode alone does not identify a primal.
Private Function OriginalDllScope(ByRef pr As RPX_AccuracyProgram) As Boolean
    Dim i As Long
    If pr.mode <> "lower" And pr.mode <> "upper" Then Exit Function
    If Not CurrentProgram(pr) Then Exit Function
    For i = 0 To pr.n - 1
        If pr.pdiag(i) <> 0# Or pr.cost(i) <> pr.originalCost(i) Then Exit Function
    Next i
    OriginalDllScope = True
End Function
' Only the literal original LP receives the thin-chain equilibration override.
' Cap/QP, equality-repair and dual helpers keep their existing settings.
Private Function ThinDllOriginalScope(ByRef pr As RPX_AccuracyProgram) As Boolean
    If Not RPX_ThinChainBudgetActive Or Not RFsActive Or RSubdivisions <> 8 Then Exit Function
    ThinDllOriginalScope = OriginalDllScope(pr)
End Function
Private Function ThinDllEquilibrate() As Long
    Dim equ As Double
    equ = CDbl(RPX_Setting("ADAPT_THIN_DLL_EQUILIBRATE", 1))
    If Not RPX_SchurFinite(equ) Or (equ <> 0# And equ <> 1#) Then err.Raise 5, "RPX_AccuracyNative", "INVALID_ADAPT_THIN_DLL_EQUILIBRATE"
    ThinDllEquilibrate = CLng(equ)
End Function
Public Function RPX_AccuracySolve(ByRef pr As RPX_AccuracyProgram, ByVal seconds As Double, ByRef answer() As Double, ByRef modifiedDual() As Double) As Boolean
    Dim packedEquilibrate As Long
    RPX_AccuracySolve = AccuracySolveCore(pr, seconds, answer, modifiedDual, False, packedEquilibrate)
End Function
Private Function AccuracySolveCore(ByRef pr As RPX_AccuracyProgram, ByVal seconds As Double, ByRef answer() As Double, ByRef modifiedDual() As Double, ByVal thinUpperRetry As Boolean, ByRef packedEquilibrate As Long) As Boolean
    Dim descriptor(0 To 87) As Byte, settings(0 To 127) As Byte, progress(0 To 79) As Byte
    Dim handle As LongLong, rc As Long, state As Long, number As Long, message As String, source As String
    Dim cancelled As Boolean, valid As Boolean, i As Long, begin As Double, elapsed As Double, equ As Double
    Dim thinOriginal As Boolean, callDetail As String, requestedSeconds As Double, packedSeconds As Double
    Dim literalOriginal As Boolean, programRole As String, upperEquilibrateSetting As Long, upperEqu As Double
    Dim callId As Long, resultStatus As Long
    accuracyCallSequence = accuracyCallSequence + 1: callId = accuracyCallSequence
    AccuracyNativeCallId = callId: AccuracyNativeSolverStatus = 0: AccuracyNativeIterations = 0
    AccuracyNativeEquilibrate = -1: upperEquilibrateSetting = -1: programRole = "unclassified"
    requestedSeconds = seconds: packedSeconds = -1#: resultStatus = -1
    packedEquilibrate = -1
    Erase answer: Erase modifiedDual: AccuracyNativeReason = "DLL_NOT_AVAILABLE"
    If Not RPX_SchurFinite(seconds) Or seconds <= 0# Or seconds > 600# Then Exit Function
    If RPX_ThinUpperProtectOn() Then seconds = RPX_ThinUpperProtectLimit(seconds, "dll_solve")
    On Error GoTo Failed
    If libraryHandle = 0 Then
        libraryPath = ThisWorkbook.path & "\rpfem_native_x64.dll"
        libraryHandle = LoadAccuracyLibrary(StrPtr(libraryPath))
        If libraryHandle = 0 Then Exit Function
    ElseIf libraryPath <> ThisWorkbook.path & "\rpfem_native_x64.dll" Then
        AccuracyNativeReason = "ACCURACY_DLL_PATH_CHANGED": Exit Function
    End If
    If NAbi() <> 1 Then AccuracyNativeReason = "ACCURACY_DLL_ABI_MISMATCH": Exit Function
    rc = NDefaults(VarPtr(settings(0)))
    If rc <> 0 Then GoTo Cleanup
    PackLong descriptor, 0, 88: PackLong descriptor, 4, 1
    PackLong descriptor, 8, pr.n: PackLong descriptor, 12, pr.m: PackLong descriptor, 16, pr.nnz: PackLong descriptor, 20, pr.coneCount
    PackPointer descriptor, 24, VarPtr(pr.rowPtr(0)): PackPointer descriptor, 32, VarPtr(pr.colId(0)): PackPointer descriptor, 40, VarPtr(pr.values(0))
    PackPointer descriptor, 48, VarPtr(pr.rhs(0)): PackPointer descriptor, 56, VarPtr(pr.pdiag(0)): PackPointer descriptor, 64, VarPtr(pr.cost(0))
    PackPointer descriptor, 72, VarPtr(pr.coneKind(0)): PackPointer descriptor, 80, VarPtr(pr.coneDim(0))
    PackLong settings, 8, 350: PackLong settings, 12, 1
    thinOriginal = ThinDllOriginalScope(pr): literalOriginal = OriginalDllScope(pr)
    If literalOriginal Then
        programRole = "original_primal"
    ElseIf pr.mode = "original_dual_helper" Then
        programRole = "original_dual_helper"
    Else
        programRole = "auxiliary"
    End If
    If RPX_ThinChainBudgetActive Then
        If thinOriginal Then
            seconds = RPX_ThinSolveSeconds(seconds)
        ElseIf RPX_ThinSolveSeconds(seconds) < seconds Then
            GoTo Cleanup
        End If
        If seconds <= 0# Then GoTo Cleanup
    End If
    PackDouble settings, 16, seconds: packedSeconds = seconds
    PackDouble settings, 24, 0.0000000000001: PackDouble settings, 32, 0.0000000000001: PackDouble settings, 40, 0.0000000000001
    PackDouble settings, 64, 0.000000000000001: PackDouble settings, 72, 0.000000000000001: PackLong settings, 80, 30
    If thinOriginal Then
        equ = ThinDllEquilibrate()
    Else
        If pr.mode = "lower" Or pr.mode = "upper" Then
            equ = CDbl(RPX_Setting("ACCURACY_DLL_EQUILIBRATE", 0))
        Else
            equ = CDbl(RPX_Setting("ACCURACY_DLL_DUAL_EQUILIBRATE", 1))
        End If
        If Not RPX_SchurFinite(equ) Or (equ <> 0# And equ <> 1#) Then err.Raise 5, "RPX_AccuracyNative", "INVALID_ACCURACY_DLL_EQUILIBRATE"
    End If
    ' Reuse documented fields already sent through ABI 1. No initial-point
    ' pointer exists in this descriptor, so a native warm start is unavailable.
    If thinUpperRetry Then
        If Not ThinUpperRetryScope(pr) Then err.Raise 5, "RPX_AccuracyNative", "THIN_UPPER_RETRY_CONTEXT_REJECTED"
        PackLong settings, 8, RPX_ThinUpperDllRetryIterations
        If Not thinOriginal Then equ = RPX_ThinUpperDllRetryEquilibrate
    End If
    ' One enables the scoped override; zero preserves the complete rc9a selector.
    If literalOriginal And pr.mode = "upper" Then
        upperEqu = CDbl(RPX_Setting("ACCURACY_DLL_UPPER_EQUILIBRATE", 1))
        If Not RPX_SchurFinite(upperEqu) Or (upperEqu <> 0# And upperEqu <> 1#) Then err.Raise 5, "RPX_AccuracyNative", "INVALID_ACCURACY_DLL_UPPER_EQUILIBRATE"
        upperEquilibrateSetting = CLng(upperEqu)
        If upperEquilibrateSetting = 1 Then equ = 1#
    End If
    PackLong settings, 84, CLng(equ): packedEquilibrate = CLng(equ): AccuracyNativeEquilibrate = packedEquilibrate
    If thinOriginal Then RPX_DiagEvent "thin_dll_equilibrate", "mode=" & pr.mode & ";Fs=" & pr.factor & ";equilibrate=" & packedEquilibrate & ";status5_retry=" & thinUpperRetry
    PackLong progress, 0, 80: PackLong progress, 4, 1
    rc = NCreate(VarPtr(descriptor(0)), VarPtr(settings(0)), handle)
    If rc <> 0 Then GoTo Cleanup
    If RPX_ThinChainBudgetActive Then
        If XCancel Or RPX_ThinSolveSeconds(1#) <= 0# Then GoTo Cleanup
        If Not thinOriginal Then
            If RPX_ThinSolveSeconds(seconds) < seconds Then GoTo Cleanup
        End If
    End If
    rc = NStart(handle): If rc <> 0 Then GoTo Cleanup
    begin = Timer
    Do
        rc = NPoll(handle, VarPtr(progress(0))): If rc <> 0 Then GoTo Cleanup
        state = ReadLong(progress, 8)
        If RPX_ThinChainBudgetActive Then
            If XCancel Or RPX_ThinSolveSeconds(1#) <= 0# Then
                cancelled = True
                If state = 0 Or state = 1 Then rc = NCancel(handle)
            End If
        End If
        If state <> 1 And state <> 0 Then Exit Do
        If RPX_ThinUpperProtectOn() Then Call RPX_ThinUpperProtectCheck("dll_poll")
        elapsed = Timer - begin: If elapsed < 0# Then elapsed = elapsed + 86400#
        If XCancel Or elapsed > seconds + 5# Then
            cancelled = True: rc = NCancel(handle)
        End If
        DoEvents: Idle 10
    Loop
    AccuracyNativeSolverStatus = ReadLong(progress, 12): AccuracyNativeIterations = ReadLong(progress, 16)
    If state <> 2 Or cancelled Then GoTo Cleanup
    ReDim answer(0 To pr.n - 1): ReDim modifiedDual(0 To pr.m - 1)
    rc = NResult(handle, VarPtr(answer(0)), pr.n, VarPtr(modifiedDual(0)), pr.m, VarPtr(progress(0)))
    If rc <> 0 Then GoTo Cleanup
    resultStatus = ReadLong(progress, 12)
    For i = 0 To pr.n - 1: If Not RPX_SchurFinite(answer(i)) Then GoTo Cleanup
    Next i
    For i = 0 To pr.m - 1: If Not RPX_SchurFinite(modifiedDual(i)) Then GoTo Cleanup
    Next i
    valid = True
Cleanup:
    If handle <> 0 Then
        Do
            rc = NDestroy(handle)
            If rc <> -8 Then Exit Do
            rc = NCancel(handle): DoEvents: Idle 10
        Loop
        If rc <> 0 Then valid = False
        handle = 0
    End If
    If Not valid Then Erase answer: Erase modifiedDual
    callDetail = "call_id=" & callId & ";program_mode=" & pr.mode & ";program_role=" & programRole & ";n=" & pr.n & ";m=" & pr.m
    callDetail = callDetail & ";original_primal=" & literalOriginal & ";upper_equilibrate_setting=" & upperEquilibrateSetting
    callDetail = callDetail & ";thin_original=" & thinOriginal & ";thin_upper_retry=" & thinUpperRetry
    callDetail = callDetail & ";requested_seconds=" & requestedSeconds & ";packed_seconds=" & packedSeconds
    callDetail = callDetail & ";iteration_limit=" & ReadLong(settings, 8) & ";equilibrate=" & packedEquilibrate
    callDetail = callDetail & ";state=" & state & ";poll_status=" & AccuracyNativeSolverStatus & ";result_status=" & resultStatus
    callDetail = callDetail & ";iterations=" & AccuracyNativeIterations & ";cancelled=" & cancelled & ";finite_candidate=" & valid
    RPX_DiagEvent "accuracy_dll_call", callDetail & ";status_is_not_certificate=1"
    If valid Then AccuracyNativeReason = "CANDIDATE_REQUIRES_ORIGINAL_AUDIT" Else AccuracyNativeReason = "ACCURACY_DLL_NO_FINITE_CANDIDATE"
    If number = 5 And message = "THIN_UPPER_LOWER_PROTECT" Then
        ' Resume Cleanup re-enables the handler: disable it before propagating.
        On Error GoTo 0
        If XCancel Then err.Raise 18, "RPX_AccuracyNative", "ANALYSIS_CANCELLED"
        err.Raise number, source, message
    End If
    If XCancel Then err.Raise 18, "RPX_AccuracyNative", "ANALYSIS_CANCELLED"
    If number <> 0 And number <> 48 And number <> 53 And number <> 453 Then err.Raise number, source, message
    AccuracySolveCore = valid: Exit Function
Failed:
    number = err.number: message = err.description: source = err.source
    Resume Cleanup
End Function
' A retry changes solver scheduling only; its result still enters all existing
' original primal/dual audits and the q=8 physical field gate in DirectBound.
Private Function ThinUpperRetryScope(ByRef pr As RPX_AccuracyProgram) As Boolean
    If Not RPX_ThinChainBudgetActive Or Not RPX_ThinChainUpperActive Or Not RFsActive Then Exit Function
    If Not RPX_ThinUpperDllRetry Or RSubdivisions <> 8 Or pr.mode <> "upper" Then Exit Function
    If re < RPX_ThinUpperDllMinElements Or Not CurrentProgram(pr) Then Exit Function
    ThinUpperRetryScope = True
End Function
Private Function ThinUpperStatus5Retry(ByRef pr As RPX_AccuracyProgram, ByVal ordinarySeconds As Double, ByRef candidate() As Double, ByRef originalDual() As Double, ByVal hadCandidate As Boolean) As Boolean
    Dim retrySeconds As Double, remainingSeconds As Double, remainingSource As String, initialIterations As Long
    Dim firstCandidate() As Double, firstDual() As Double, initialReason As String, retryFinite As Boolean
    Dim retryStatus As Long, retryIterations As Long, retryStarted As Double, retryEquilibratePacked As Long
    Dim initialCallId As Long, initialEquilibrate As Long
    ThinUpperStatus5Retry = hadCandidate
    If AccuracyNativeSolverStatus <> 5 Or Not ThinUpperRetryScope(pr) Then Exit Function
    retrySeconds = WorksheetFunction.Max(ordinarySeconds, RPX_ThinUpperDllRetrySeconds)
    If retrySeconds > 600# Then retrySeconds = 600#
    remainingSeconds = RPX_ThinRemainingSeconds(remainingSource)
    If remainingSeconds >= 0# Then
        If retrySeconds > remainingSeconds - RPX_ThinUpperHsdReserveSeconds Then retrySeconds = remainingSeconds - RPX_ThinUpperHsdReserveSeconds
    End If
    If RPX_ThinUpperProtectOn() Then retrySeconds = RPX_ThinUpperProtectLimit(retrySeconds, "dll_status5_retry")
    If retrySeconds <= 0# Then
        RPX_DiagEvent "thin_layer_dll_retry", "mode=upper;Fs=" & pr.factor & ";initial_status=5;attempts=0;reason=remaining_time;remaining_seconds=" & remainingSeconds & ";remaining_source=" & remainingSource & ";bound=0"
        Exit Function
    End If
    If hadCandidate Then firstCandidate = candidate: firstDual = originalDual
    initialIterations = AccuracyNativeIterations: initialReason = AccuracyNativeReason
    initialCallId = AccuracyNativeCallId: initialEquilibrate = AccuracyNativeEquilibrate
    RPX_AssertInputs "thin_upper_dll_retry"
    retryStarted = RPX_DiagClock()
    retryFinite = AccuracySolveCore(pr, retrySeconds, candidate, originalDual, True, retryEquilibratePacked)
    retryStatus = AccuracyNativeSolverStatus: retryIterations = AccuracyNativeIterations
    RPX_AssertInputs "thin_upper_dll_retry_complete"
    If Not CurrentProgram(pr) Then err.Raise 5, "RPX_AccuracyNative", "THIN_UPPER_RETRY_CONTEXT_CHANGED"
    ' A failed second execution must not discard the first finite candidate.
    If Not retryFinite Then
        If hadCandidate Then candidate = firstCandidate: originalDual = firstDual
        AccuracyNativeSolverStatus = 5: AccuracyNativeIterations = initialIterations: AccuracyNativeReason = initialReason
        AccuracyNativeCallId = initialCallId: AccuracyNativeEquilibrate = initialEquilibrate ' Restored candidate belongs to the first call.
    Else
        ThinUpperStatus5Retry = True
    End If
    RPX_DiagEvent "thin_layer_dll_retry", "mode=upper;Fs=" & pr.factor & ";initial_status=5;initial_iterations=" & initialIterations & ";attempts=1;iteration_limit=" & RPX_ThinUpperDllRetryIterations & ";equilibrate=" & retryEquilibratePacked & ";seconds_limit=" & retrySeconds & ";seconds=" & RPX_DiagClock() - retryStarted & ";solver_status=" & retryStatus & ";iterations=" & retryIterations & ";finite_candidate=" & retryFinite & ";original_audit_required=1;bound=0"
End Function
Public Function RPX_AccuracyPrimalAudit(ByRef original As RPX_AccuracyProgram, ByRef candidate() As Double, ByRef objective As Double) As Boolean
    Dim i As Long, j As Long, k As Long, row As Long, total As Double, correction As Double, cv(0 To 2) As Double, violation As Double
    AccuracyNativeEquality = 0#: AccuracyNativeCone = 0#: objective = 0#
    On Error GoTo Rejected
    If Not CurrentProgram(original) Or LBound(candidate) <> 0 Or UBound(candidate) <> original.n - 1 Then Exit Function
    For i = 0 To original.n - 1
        If Not RPX_SchurFinite(candidate(i)) Then Exit Function
        AddTerm objective, correction, original.originalCost(i) * candidate(i)
    Next i
    For i = 0 To original.equalityCount - 1
        total = -original.rhs(i): correction = 0#
        If Not RPX_SchurFinite(total) Then Exit Function
        For j = original.rowPtr(i) To original.rowPtr(i + 1) - 1: AddTerm total, correction, original.values(j) * candidate(original.colId(j)): Next j
        AccuracyNativeEquality = WorksheetFunction.Max(AccuracyNativeEquality, Abs(total))
    Next i
    row = original.equalityCount
    For k = 1 To original.coneCount - 1
        For i = 0 To original.coneDim(k) - 1
            total = original.rhs(row): correction = 0#
            If Not RPX_SchurFinite(total) Then Exit Function
            For j = original.rowPtr(row) To original.rowPtr(row + 1) - 1: AddTerm total, correction, -original.values(j) * candidate(original.colId(j)): Next j
            cv(i) = total: row = row + 1
        Next i
        violation = -cv(0): If original.coneDim(k) = 3 Then violation = Norm2(cv(1), cv(2)) - cv(0)
        AccuracyNativeCone = WorksheetFunction.Max(AccuracyNativeCone, violation)
    Next k
    RPX_AccuracyPrimalAudit = AccuracyNativeEquality <= 0.0000001 And AccuracyNativeCone <= 0.0000001
    Exit Function
Rejected:
    If err.number <> 6 And err.number <> 9 And Left$(err.description, 18) <> "ACCURACY_NONFINITE" Then err.Raise err.number, err.source, err.description
    err.Clear: AccuracyNativeReason = "INVALID_ORIGINAL_PRIMAL"
End Function
Private Sub AddBound(ByRef pr As RPX_AccuracyProgram, ByVal loadId As Long, ByVal rhs As Double)
    Dim j As Long, p As Long
    p = pr.nnz
    If loadId < 0 Then
        For j = 0 To pr.n - 1: If pr.originalCost(j) <> 0# Then p = p + 1
        Next j
    Else
        p = p + 1
    End If
    ReDim Preserve pr.colId(0 To p - 1): ReDim Preserve pr.values(0 To p - 1)
    p = pr.nnz
    If loadId < 0 Then
        For j = 0 To pr.n - 1
            If pr.originalCost(j) <> 0# Then pr.colId(p) = j: pr.values(p) = pr.originalCost(j): p = p + 1
        Next j
    Else
        pr.colId(p) = loadId: pr.values(p) = -1#: p = p + 1
    End If
    ReDim Preserve pr.rowPtr(0 To pr.m + 1): pr.rowPtr(pr.m + 1) = p
    ReDim Preserve pr.rhs(0 To pr.m): pr.rhs(pr.m) = rhs
    ReDim Preserve pr.coneKind(0 To pr.coneCount): ReDim Preserve pr.coneDim(0 To pr.coneCount)
    pr.coneKind(pr.coneCount) = 1: pr.coneDim(pr.coneCount) = 1
    pr.nnz = p: pr.m = pr.m + 1: pr.coneCount = pr.coneCount + 1
End Sub
Private Function AccuracyArrayShape(ByRef values() As Double, ByRef firstIndex As Long, ByRef lastIndex As Long) As Boolean
    On Error GoTo Missing
    firstIndex = LBound(values)
    lastIndex = UBound(values)
    AccuracyArrayShape = True
    Exit Function
Missing:
    If err.number <> 9 Then err.Raise err.number, err.source, err.description
    err.Clear
End Function
Public Function RPX_AccuracyRecoverPrimal(ByVal factor As Double, ByVal mode As String, ByRef recovered() As Double) As Boolean
    Dim original As RPX_AccuracyProgram, qp As RPX_AccuracyProgram, center() As Double, candidate() As Double, modifiedDual() As Double
    Dim alpha As Double, seconds As Double, objective As Double, raw() As Double, i As Long, ok As Boolean, positiveScale As Double, homogeneous As Boolean
    Dim auditState As RPX_AuditSnapshot, number As Long, message As String, source As String, oldRejected As Boolean
    Dim centerFirst As Long, centerLast As Long
    If mode = "upper" Then
        If UpperRecenterEnabled() Then
            RPX_AccuracyRecoverPrimal = RecoverUpperRecenter(factor, recovered)
            Exit Function
        End If
    End If
    Erase recovered
    If Not RPX_AccuracyExport(factor, mode, original) Then Exit Function
    qp = original: ReDim center(0 To original.n - 1)
    If AccuracyArrayShape(xx, centerFirst, centerLast) Then
        If centerFirst = 0 Then
            If centerLast = original.n - 1 Then center = xx
        End If
    End If
    For i = 0 To UBound(center): If Not RPX_SchurFinite(center(i)) Then center(i) = 0#
    Next i
    alpha = CDbl(RPX_Setting("ACCURACY_DLL_ALPHA", 0.01)): seconds = CDbl(RPX_Setting("ACCURACY_DLL_SECONDS", 120))
    If Not RPX_SchurFinite(alpha) Or alpha <= 0# Then err.Raise 5, "RPX_AccuracyNative", "INVALID_ACCURACY_DLL_ALPHA"
    If mode = "lower" Then AddBound qp, RLoadVariable, -1.000003 Else AddBound qp, -1, 0.999998
    For i = 0 To original.n - 1: qp.pdiag(i) = alpha: qp.cost(i) = -alpha * center(i): Next i
    ScaleSolverCones qp
    If Not RPX_AccuracySolve(qp, seconds, candidate, modifiedDual) Then Exit Function
    Erase modifiedDual ' The cap/floor/QP dual cannot prove an original root sign.
    If Not CurrentProgram(original) Then AccuracyNativeReason = "MODEL_CHANGED_DURING_ACCURACY_SOLVE": Exit Function
    RPX_DiagAuditSave auditState
    oldRejected = RPX_AuditRejected
    On Error GoTo AuditFailed
    ok = RPX_AccuracyPrimalAudit(original, candidate, objective)
    If RPX_ThinChainBudgetActive And RPX_ThinChainUpperActive And mode = "upper" Then
        If AccuracyNativeSolverStatus <> 1 And AccuracyNativeSolverStatus <> 4 Then ok = False
    End If
    If ok Then ok = RPX_FieldGate(mode, candidate, factor, 8, objective, AccuracyNativeValue, AccuracyNativeAudit, raw)
    If mode = "lower" And Not ok Then
        ' Homogeneous equilibrium permits a conservative same-Fs positive scaling.
        ' Each scaled candidate is checked afresh; no corrected/reduced dual is used.
        homogeneous = True
        For i = 0 To original.equalityCount - 1: If original.rhs(i) <> 0# Then homogeneous = False
        Next i
        If homogeneous And RPX_SchurFinite(candidate(RLoadVariable)) Then
            If candidate(RLoadVariable) >= 1.0000015 Then
                positiveScale = 1.0000015 / candidate(RLoadVariable)
                For i = 0 To UBound(candidate): candidate(i) = candidate(i) * positiveScale: Next i
                ok = RPX_AccuracyPrimalAudit(original, candidate, objective)
                If ok Then ok = RPX_FieldGate(mode, candidate, factor, 8, objective, AccuracyNativeValue, AccuracyNativeAudit, raw)
            End If
        End If
        If Not ok Then ok = RecoverEquilibrium(original, candidate, objective, raw)
    End If
RestoreAudit:
    RPX_DiagAuditRestore auditState
    RPX_AuditRejected = oldRejected
    On Error GoTo 0
    If number <> 0 Then err.Raise number, source, message
    If ok Then
        If Not RPX_SchurFinite(AccuracyNativeValue) Or Not RPX_SchurFinite(AccuracyNativeAudit) Then ok = False
        If mode = "lower" Then
            If AccuracyNativeValue < 1.000001 Then ok = False
        Else
            If AccuracyNativeValue > 0.999999 Then ok = False
        End If
    End If
    If ok Then recovered = candidate: AccuracyNativeReason = "ORIGINAL_PHYSICAL_PRIMAL_WITNESS" Else AccuracyNativeReason = "ORIGINAL_PHYSICAL_PRIMAL_REJECTED"
    If ok Then CaptureAccuracyProof original, candidate, False
    RPX_DiagEvent "accuracy_dll_primal", "accepted=" & ok & ";mode=" & mode & ";Fs=" & factor & ";q=8;raw_eq=" & AccuracyNativeEquality & ";raw_cone=" & AccuracyNativeCone & ";value=" & AccuracyNativeValue & ";modified_dual_used=0;root_created=0"
    RPX_AccuracyRecoverPrimal = ok
    Exit Function
AuditFailed:
    number = err.number: message = err.description: source = err.source: ok = False
    Resume RestoreAudit
End Function
Private Sub ScaleSolverCones(ByRef pr As RPX_AccuracyProgram)
    Dim enabled As Double, k As Long, row As Long, i As Long, j As Long, maxValue As Double, squares As Double, multiplier As Double, coefficient As Double
    enabled = CDbl(RPX_Setting("ACCURACY_DLL_CONE_SCALING", 0))
    If Not RPX_SchurFinite(enabled) Or (enabled <> 0# And enabled <> 1#) Then err.Raise 5, "RPX_AccuracyNative", "INVALID_ACCURACY_DLL_CONE_SCALING"
    If enabled = 0# Then Exit Sub
    row = 0
    For k = 0 To pr.coneCount - 1
        If pr.coneKind(k) <> 0 Then
            maxValue = 0#: squares = 0#
            For i = row To row + pr.coneDim(k) - 1
                For j = pr.rowPtr(i) To pr.rowPtr(i + 1) - 1: maxValue = WorksheetFunction.Max(maxValue, Abs(pr.values(j))): Next j
            Next i
            If maxValue > 0# Then
                For i = row To row + pr.coneDim(k) - 1
                    For j = pr.rowPtr(i) To pr.rowPtr(i + 1) - 1: coefficient = pr.values(j) / maxValue: squares = squares + coefficient * coefficient: Next j
                Next i
                If maxValue > 0.0000000001 / Sqr(squares) Then
                    multiplier = (1# / maxValue) / Sqr(squares)
                    If Not RPX_SchurFinite(multiplier) Or multiplier <= 0# Then err.Raise 5, "RPX_AccuracyNative", "ACCURACY_CONE_SCALE_NONFINITE"
                    For i = row To row + pr.coneDim(k) - 1
                        pr.rhs(i) = pr.rhs(i) * multiplier
                        For j = pr.rowPtr(i) To pr.rowPtr(i + 1) - 1: pr.values(j) = pr.values(j) * multiplier: Next j
                    Next i
                End If
            End If
        End If
        row = row + pr.coneDim(k)
    Next k
End Sub
Private Function RecoverEquilibrium(ByRef original As RPX_AccuracyProgram, ByRef sourceField() As Double, ByRef objective As Double, ByRef raw() As Double) As Boolean
    Dim problem(0 To 63) As Byte, report(0 To 55) As Byte, rhs() As Double, correction() As Double, nextCorrection() As Double, candidate() As Double
    Dim i As Long, j As Long, stage As Long, budget As Long, iterations As Long, rc As Long, total As Double, compensation As Double, accepted As Boolean
    Dim enabled As Double, number As Long, message As String, errorSource As String
    enabled = CDbl(RPX_Setting("ACCURACY_DLL_EQUALITY_RECOVERY", 1))
    If Not RPX_SchurFinite(enabled) Or (enabled <> 0# And enabled <> 1#) Then err.Raise 5, "RPX_AccuracyNative", "INVALID_ACCURACY_DLL_EQUALITY_RECOVERY"
    If enabled = 0# Or original.mode <> "lower" Or Not CurrentProgram(original) Then Exit Function
    ReDim rhs(0 To original.equalityCount - 1): ReDim correction(0 To original.n - 1): ReDim nextCorrection(0 To original.n - 1)
    For i = 0 To original.equalityCount - 1
        total = original.rhs(i): compensation = 0#
        For j = original.rowPtr(i) To original.rowPtr(i + 1) - 1: AddTerm total, compensation, -original.values(j) * sourceField(original.colId(j)): Next j
        rhs(i) = total
    Next i
    PackLong problem, 0, 64: PackLong problem, 4, 1: PackLong problem, 8, original.n: PackLong problem, 12, original.equalityCount
    PackLong problem, 16, original.rowPtr(original.equalityCount)
    PackPointer problem, 24, VarPtr(original.rowPtr(0)): PackPointer problem, 32, VarPtr(original.colId(0)): PackPointer problem, 40, VarPtr(original.values(0))
    PackPointer problem, 48, VarPtr(rhs(0))
    PackLong report, 0, 56: PackLong report, 4, 1
    On Error GoTo Failed
    For stage = 1 To 3
        If stage = 1 Then
            budget = 100
        ElseIf stage = 2 Then
            budget = 500
        Else
            budget = 1000
        End If
        budget = budget - iterations: If budget <= 0 Then Exit For
        If XCancel Then err.Raise 18, "RPX_AccuracyNative", "ANALYSIS_CANCELLED"
        PackPointer problem, 56, VarPtr(correction(0))
        rc = NEqualities(VarPtr(problem(0)), budget, 0.000000000000001, 0.000000000000001, 100000000#, VarPtr(nextCorrection(0)), original.n, VarPtr(report(0)))
        If rc <> 0 Then Exit For
        iterations = iterations + ReadLong(report, 12): correction = nextCorrection: candidate = sourceField
        For i = 0 To original.n - 1
            If Not RPX_SchurFinite(correction(i)) Then Exit Function
            candidate(i) = sourceField(i) + correction(i)
        Next i
        accepted = RPX_AccuracyPrimalAudit(original, candidate, objective)
        If accepted Then accepted = RPX_FieldGate("lower", candidate, original.factor, 8, objective, AccuracyNativeValue, AccuracyNativeAudit, raw)
        If accepted And AccuracyNativeValue >= 1.000001 Then sourceField = candidate: RecoverEquilibrium = True: Exit Function
        RPX_DiagEvent "accuracy_dll_equality_stage", "stage=" & stage & ";cumulative_iterations=" & iterations & ";stop=" & ReadLong(report, 8) & ";accepted=" & accepted & ";q=8;full_original_gate_required=1"
        Select Case ReadLong(report, 8)
            Case 0, 1, 2, 4, 5
                If ReadLong(report, 12) < budget Then Exit For
        End Select
        DoEvents
        If Not CurrentProgram(original) Then Exit Function
    Next stage
    ' G1R20 Fix8: direct stress-only repair (load fixed) when LSMR stalls.
    RecoverEquilibrium = DirectEqualityRepair(original, sourceField, objective, raw)
    Exit Function
Failed:
    number = err.number: message = err.description: errorSource = err.source
    If number = 453 Then AccuracyNativeReason = "ACCURACY_DLL_EQUALITY_ENTRY_UNAVAILABLE": Exit Function
    err.Raise number, errorSource, message
End Function
' G1R20 Fix8 (candidate): minimum-norm stress correction with the load
' multiplier held fixed, min 1/2|dx|^2 s.t. A_s dx = b - A x over the original
' equality rows, solved by the accuracy DLL (direct KKT factorization) with up
' to three refinement passes. The corrected field must pass the unchanged
' original primal audit and the full q8 physical gate; nothing else is relaxed.
Private Function DirectEqualityRepair(ByRef original As RPX_AccuracyProgram, ByRef sourceField() As Double, ByRef objective As Double, ByRef raw() As Double) As Boolean
    Dim enabled As Double, gate As Double, seconds As Double, pr As RPX_AccuracyProgram, blank As RPX_AccuracyProgram
    Dim i As Long, j As Long, p As Long, rows As Long, pass As Long, loadId As Long, nnz As Long, rowHas As Boolean
    Dim residual() As Double, total As Double, compensation As Double, worst As Double
    Dim answer() As Double, modifiedDual() As Double, candidate() As Double, accepted As Boolean, began As Double
    enabled = CDbl(RPX_Setting("ACCURACY_DLL_EQUALITY_DIRECT", 1))
    If Not RPX_SchurFinite(enabled) Or (enabled <> 0# And enabled <> 1#) Then err.Raise 5, "RPX_AccuracyNative", "INVALID_ACCURACY_DLL_EQUALITY_DIRECT"
    gate = CDbl(RPX_Setting("ACCURACY_DLL_EQUALITY_DIRECT_GATE", 0.00001))
    If Not RPX_SchurFinite(gate) Or gate <= 0# Then err.Raise 5, "RPX_AccuracyNative", "INVALID_ACCURACY_DLL_EQUALITY_DIRECT_GATE"
    seconds = CDbl(RPX_Setting("ACCURACY_DLL_EQUALITY_DIRECT_SECONDS", 60))
    If Not RPX_SchurFinite(seconds) Or seconds <= 0# Or seconds > 600# Then err.Raise 5, "RPX_AccuracyNative", "INVALID_ACCURACY_DLL_EQUALITY_DIRECT_SECONDS"
    If enabled = 0# Or original.mode <> "lower" Or Not CurrentProgram(original) Then Exit Function
    loadId = RLoadVariable
    If loadId < 0 Or loadId >= original.n Then Exit Function
    candidate = sourceField: began = Timer
    For pass = 1 To 3
        If XCancel Then err.Raise 18, "RPX_AccuracyNative", "ANALYSIS_CANCELLED"
        ReDim residual(0 To original.equalityCount - 1): worst = 0#: rows = 0: nnz = 0
        For i = 0 To original.equalityCount - 1
            total = original.rhs(i): compensation = 0#: rowHas = False
            For j = original.rowPtr(i) To original.rowPtr(i + 1) - 1
                AddTerm total, compensation, -original.values(j) * candidate(original.colId(j))
                If original.colId(j) <> loadId And original.values(j) <> 0# Then rowHas = True: nnz = nnz + 1
            Next j
            residual(i) = total
            If Abs(total) > worst Then worst = Abs(total)
            If rowHas Then
                rows = rows + 1
            ElseIf total <> 0# Then
                RPX_DiagEvent "accuracy_dll_equality_direct", "pass=" & pass & ";skipped=load_only_row;row=" & i
                Exit Function
            End If
        Next i
        If pass = 1 And worst > gate Then
            RPX_DiagEvent "accuracy_dll_equality_direct", "pass=1;skipped=residual_above_gate;max_residual=" & worst & ";gate=" & gate
            Exit Function
        End If
        If rows = 0 Or nnz = 0 Then Exit Function
        pr = blank
        pr.n = original.n: pr.m = rows: pr.nnz = nnz: pr.equalityCount = rows: pr.coneCount = 1
        pr.mode = original.mode: pr.model = original.model: pr.policy = original.policy: pr.phase = original.phase
        pr.epoch = original.epoch: pr.q = original.q: pr.factor = original.factor
        ReDim pr.rowPtr(0 To rows): ReDim pr.colId(0 To nnz - 1): ReDim pr.values(0 To nnz - 1): ReDim pr.rhs(0 To rows - 1)
        ReDim pr.pdiag(0 To pr.n - 1): ReDim pr.cost(0 To pr.n - 1): ReDim pr.originalCost(0 To pr.n - 1)
        ReDim pr.coneKind(0 To 0): ReDim pr.coneDim(0 To 0): pr.coneKind(0) = 0: pr.coneDim(0) = rows
        For i = 0 To pr.n - 1: pr.pdiag(i) = 1#: Next i
        rows = 0: p = 0
        For i = 0 To original.equalityCount - 1
            rowHas = False
            For j = original.rowPtr(i) To original.rowPtr(i + 1) - 1
                If original.colId(j) <> loadId And original.values(j) <> 0# Then rowHas = True: Exit For
            Next j
            If rowHas Then
                pr.rowPtr(rows) = p: pr.rhs(rows) = residual(i)
                For j = original.rowPtr(i) To original.rowPtr(i + 1) - 1
                    If original.colId(j) <> loadId And original.values(j) <> 0# Then pr.colId(p) = original.colId(j): pr.values(p) = original.values(j): p = p + 1
                Next j
                rows = rows + 1
            End If
        Next i
        pr.rowPtr(rows) = p
        If p <> nnz Then err.Raise 5, "RPX_AccuracyNative", "FIX8_DIRECT_REPAIR_NNZ"
        If Not RPX_AccuracySolve(pr, seconds, answer, modifiedDual) Then
            RPX_DiagEvent "accuracy_dll_equality_direct", "pass=" & pass & ";solved=0;max_residual=" & worst & ";reason=" & AccuracyNativeReason
            Exit Function
        End If
        Erase modifiedDual
        For i = 0 To original.n - 1
            If i <> loadId Then candidate(i) = candidate(i) + answer(i)
        Next i
        If Not CurrentProgram(original) Then Exit Function
        accepted = RPX_AccuracyPrimalAudit(original, candidate, objective)
        If accepted Then accepted = RPX_FieldGate("lower", candidate, original.factor, 8, objective, AccuracyNativeValue, AccuracyNativeAudit, raw)
        RPX_DiagEvent "accuracy_dll_equality_direct", "pass=" & pass & ";solved=1;residual_before=" & worst & ";raw_eq_after=" & AccuracyNativeEquality & ";raw_cone_after=" & AccuracyNativeCone & ";accepted=" & accepted & ";value=" & AccuracyNativeValue & ";seconds=" & Format$(Timer - began, "0.0") & ";load_fixed=1;full_original_gate_required=1"
        If accepted And AccuracyNativeValue >= 1.000001 Then sourceField = candidate: DirectEqualityRepair = True: Exit Function
        If AccuracyNativeEquality <= 0.00000001 Then Exit Function
    Next pass
End Function
Public Sub RPX_AccuracyDualProblem(ByRef original As RPX_AccuracyProgram, ByRef dualProgram As RPX_AccuracyProgram)
    Dim pr As RPX_AccuracyProgram, counts() As Long, cursor() As Long, i As Long, j As Long, row As Long, col As Long, p As Long, threshold As Double
    pr.n = original.m: pr.m = original.n + original.m - original.equalityCount + 1
    pr.mode = "original_dual_helper" ' Keeps the existing dual equilibration branch.
    pr.coneCount = original.coneCount + 1
    ReDim counts(0 To pr.m - 1): ReDim cursor(0 To pr.m - 1)
    For i = 0 To original.m - 1
        For j = original.rowPtr(i) To original.rowPtr(i + 1) - 1: counts(original.colId(j)) = counts(original.colId(j)) + 1: Next j
    Next i
    For i = original.n To pr.m - 2: counts(i) = 1: Next i
    For i = 0 To original.m - 1: If original.rhs(i) <> 0# Then counts(pr.m - 1) = counts(pr.m - 1) + 1
    Next i
    ReDim pr.rowPtr(0 To pr.m)
    For i = 0 To pr.m - 1: pr.rowPtr(i + 1) = pr.rowPtr(i) + counts(i): cursor(i) = pr.rowPtr(i): Next i
    pr.nnz = pr.rowPtr(pr.m)
    ReDim pr.colId(0 To pr.nnz - 1): ReDim pr.values(0 To pr.nnz - 1): ReDim pr.rhs(0 To pr.m - 1)
    ReDim pr.cost(0 To pr.n - 1): ReDim pr.pdiag(0 To pr.n - 1)
    ReDim pr.coneKind(0 To pr.coneCount - 1): ReDim pr.coneDim(0 To pr.coneCount - 1)
    For i = 0 To original.m - 1
        For j = original.rowPtr(i) To original.rowPtr(i + 1) - 1
            row = original.colId(j): p = cursor(row): cursor(row) = p + 1
            pr.colId(p) = i: pr.values(p) = original.values(j)
        Next j
    Next i
    For i = 0 To original.n - 1: pr.rhs(i) = -original.originalCost(i): Next i
    For i = original.n To pr.m - 2
        p = pr.rowPtr(i): pr.colId(p) = original.equalityCount + i - original.n: pr.values(p) = -1#
    Next i
    p = pr.rowPtr(pr.m - 1)
    For i = 0 To original.m - 1
        If original.rhs(i) <> 0# Then pr.colId(p) = i: pr.values(p) = original.rhs(i): p = p + 1
    Next i
    If original.mode = "upper" Then threshold = 1.000003 Else threshold = -0.999998
    pr.rhs(pr.m - 1) = -threshold
    pr.coneKind(0) = 0: pr.coneDim(0) = original.n
    For i = 1 To original.coneCount - 1: pr.coneKind(i) = original.coneKind(i): pr.coneDim(i) = original.coneDim(i): Next i
    pr.coneKind(pr.coneCount - 1) = 1: pr.coneDim(pr.coneCount - 1) = 1
    dualProgram = pr
End Sub
Public Function RPX_AccuracyDualAudit(ByRef original As RPX_AccuracyProgram, ByRef candidate() As Double, Optional ByVal candidateRole As String = "original_dual") As Boolean
    Dim stat() As Double, corr() As Double, i As Long, j As Long, k As Long, row As Long, total As Double, correction As Double, violation As Double, threshold As Double
    Dim costInf As Double, dualInf As Double, atyInf As Double, costIndex As Long, statIndex As Long, detail As String, atyComplete As Boolean
    ResetDualDiagnostics
    On Error GoTo Rejected
    If Not CurrentProgram(original) Or LBound(candidate) <> 0 Or UBound(candidate) <> original.m - 1 Then Exit Function
    ReDim stat(0 To original.n - 1): ReDim corr(0 To original.n - 1): stat = original.originalCost
    For i = 0 To original.m - 1
        If Not RPX_SchurFinite(candidate(i)) Then Exit Function
        AddTerm AccuracyNativeDualObjective, correction, -original.rhs(i) * candidate(i)
        For j = original.rowPtr(i) To original.rowPtr(i + 1) - 1: AddTerm stat(original.colId(j)), corr(original.colId(j)), original.values(j) * candidate(i): Next j
    Next i
    atyComplete = True
    For i = 0 To original.n - 1
        If Abs(original.originalCost(i)) > costInf Then costInf = Abs(original.originalCost(i)): costIndex = i
        If Abs(stat(i)) > AccuracyNativeDualStat Then AccuracyNativeDualStat = Abs(stat(i)): statIndex = i
        total = DiagnosticAbsDifference(stat(i), original.originalCost(i), atyComplete)
        atyInf = WorksheetFunction.Max(atyInf, total)
    Next i
    For i = 0 To original.m - 1: dualInf = WorksheetFunction.Max(dualInf, Abs(candidate(i))): Next i
    row = original.equalityCount
    For k = 1 To original.coneCount - 1
        violation = -candidate(row)
        If original.coneDim(k) = 3 Then violation = Norm2(candidate(row + 1), candidate(row + 2)) - candidate(row)
        AccuracyNativeDualCone = WorksheetFunction.Max(AccuracyNativeDualCone, violation): row = row + original.coneDim(k)
    Next k
    If original.mode = "upper" Then threshold = 1.000003 Else threshold = -0.999998
    RPX_AccuracyDualAudit = AccuracyNativeDualStat <= 0.0000001 And AccuracyNativeDualCone <= 0.00000001 And AccuracyNativeDualObjective >= threshold
    AccuracyNativeDualAudited = True: AccuracyNativeDualAuditCallId = AccuracyNativeCallId
    AccuracyNativeDualAuditRole = candidateRole: AccuracyNativeDualAuditStatus = AccuracyNativeSolverStatus
    detail = "call_id=" & AccuracyNativeDualAuditCallId & ";candidate_role=" & candidateRole & ";mode=" & original.mode & ";solver_status=" & AccuracyNativeDualAuditStatus
    detail = detail & ";cost_inf=" & costInf & ";cost_index=" & costIndex & ";dual_inf=" & dualInf & ";aty_inf_derived=" & atyInf
    detail = detail & ";aty_inf_derived_valid=" & atyComplete
    detail = detail & ";stat_index=" & statIndex & ";stat_at_cost_index=" & stat(costIndex) & ";cost_at_cost_index=" & original.originalCost(costIndex)
    detail = detail & ";raw_stat=" & AccuracyNativeDualStat & ";raw_cone=" & AccuracyNativeDualCone & ";original_objective=" & AccuracyNativeDualObjective
    RPX_DiagEvent "accuracy_dual_vector", detail & ";accepted=" & RPX_AccuracyDualAudit & ";physical_bound=0"
    Exit Function
Rejected:
    If err.number <> 6 And err.number <> 9 And Left$(err.description, 18) <> "ACCURACY_NONFINITE" Then err.Raise err.number, err.source, err.description
    err.Clear: AccuracyNativeReason = "INVALID_ORIGINAL_DUAL"
End Function
Public Function RPX_AccuracyRecoverDual(ByVal factor As Double, ByVal mode As String, Optional ByVal budgetOverride As Double = 0#) As Boolean
    Dim original As RPX_AccuracyProgram, dp As RPX_AccuracyProgram, candidate() As Double, helperDual() As Double, seconds As Double, accepted As Boolean
    Dim useDualRepair As Boolean, finiteHelperCandidate As Boolean, repairSolveId As Long, repairCandidateId As Long
    Dim dualStatLog As Double, dualConeLog As Double, dualObjectiveLog As Double
    RPX_AccuracyGuideReset
    ResetDualDiagnostics
    If Not RPX_AccuracyExport(factor, mode, original) Then Exit Function
    If mode = "upper" Then useDualRepair = DualLinearRepairEnabled()
    repairSolveId = P4Solve: repairCandidateId = RPX_HSDCandidateId
    RPX_AccuracyDualProblem original, dp
    If useDualRepair Then SetUpperDualRegularization dp
    seconds = CDbl(RPX_Setting("ACCURACY_DLL_DUAL_SECONDS", 20))
    If Not RPX_SchurFinite(budgetOverride) Or budgetOverride < 0# Or budgetOverride > 600# Then err.Raise 5, "RPX_AccuracyNative", "INVALID_ACCURACY_DUAL_BUDGET"
    If budgetOverride > 0# Then seconds = budgetOverride
    If useDualRepair Then DualRepairAssert original, repairSolveId, repairCandidateId, "original_dual_helper_before"
    finiteHelperCandidate = RPX_AccuracySolve(dp, seconds, candidate, helperDual)
    If useDualRepair Then DualRepairAssert original, repairSolveId, repairCandidateId, "original_dual_helper_after"
    If Not finiteHelperCandidate Then Exit Function
    Erase helperDual ' The auxiliary problem's dual is not an original program dual.
    accepted = RPX_AccuracyDualAudit(original, candidate, "helper_primal")
    If AccuracyNativeDualAudited Then CaptureDualRepairStage original, candidate, "raw_helper", accepted
    If Not accepted And mode = "upper" Then accepted = TryOriginalDualLinearRepair(original, candidate)
    If accepted Then
        guideReady = True: guideFactor = factor: guideMode = mode: guideModel = original.model: guidePolicy = original.policy
        guideEpoch = original.epoch: guideQ = original.q: guidePhase = original.phase
        guideN = XN: guideM = XM: guideS = XS: guideB = XB: guideSolve = P4Solve: guideCandidate = RPX_HSDCandidateId
        guideCost = original.originalCost: guideNormalizedCost = XQ
        CaptureAccuracyProof original, candidate, True
    End If
    If AccuracyNativeDualAudited Then
        dualStatLog = AccuracyNativeDualStat: dualConeLog = AccuracyNativeDualCone: dualObjectiveLog = AccuracyNativeDualObjective
    End If
    RPX_DiagEvent "accuracy_dll_original_dual", "accepted=" & accepted & ";mode=" & mode & ";Fs=" & factor & ";q=8;raw_stat=" & dualStatLog & ";raw_cone=" & dualConeLog & ";original_objective=" & dualObjectiveLog & ";physical_bound=0" & ";call_id=" & AccuracyNativeCallId & ";helper_status=" & AccuracyNativeSolverStatus & ";dual_audited=" & AccuracyNativeDualAudited
    RPX_AccuracyRecoverDual = accepted
End Function
Public Sub RPX_AccuracyWriteProof(ByVal path As String, ByRef pr As RPX_AccuracyProgram, ByRef field() As Double)
    Dim f As Integer, i As Long, number As Long, message As String, source As String
    On Error GoTo Failed
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
    RPX_ExportExchangeContext path & ".context.tsv"
    Exit Sub
Failed:
    number = err.number: message = err.description: source = err.source
    On Error Resume Next: Close #f: On Error GoTo 0
    err.Raise number, source, message
End Sub
Private Sub CaptureAccuracyProof(ByRef pr As RPX_AccuracyProgram, ByRef field() As Double, ByVal dual As Boolean)
    Dim enabled As Double, path As String, kind As String
    enabled = CDbl(RPX_Setting("ACCURACY_DLL_CAPTURE", 0))
    If Not RPX_SchurFinite(enabled) Or (enabled <> 0# And enabled <> 1#) Then err.Raise 5, "RPX_AccuracyNative", "INVALID_ACCURACY_DLL_CAPTURE"
    If enabled = 0# Then Exit Sub
    proofSequence = proofSequence + 1
    If dual Then kind = "dual" Else kind = "primal"
    path = RPX_DiagPath & ".accuracy_" & kind & "_solve_" & P4Solve & "_" & proofSequence & "_" & pr.mode & ".bin"
    RPX_AccuracyWriteProof path, pr, field
    RPX_DiagEvent "accuracy_dll_capture", "kind=" & kind & ";path=" & path & ";q=8;full_original_program=1;auxiliary_dual_used=0"
End Sub

' A finite original-program copy can be rescaled only on a homogeneous
' free-load problem. Neither solver status nor an auxiliary dual is evidence.
Private Function HomogeneousLowerWitness(ByRef original As RPX_AccuracyProgram, ByRef sourceField() As Double, ByRef recovered() As Double, ByRef objectiveOut As Double, ByRef measuredOut As Double) As Boolean
    Dim candidate() As Double, loadValue As Double, positiveScale As Double, i As Long
    Dim objective As Double, measured As Double, audit As Double, raw() As Double
    Erase recovered
    On Error GoTo RejectedRay
    If original.mode <> "lower" Or Not CurrentProgram(original) Then Exit Function
    If LBound(sourceField) <> 0 Or UBound(sourceField) <> original.n - 1 Then Exit Function
    For i = 0 To original.n - 1
        If i = RLoadVariable Then
            If original.originalCost(i) <> -1# Then Exit Function
        Else
            If original.originalCost(i) <> 0# Then Exit Function
        End If
    Next i
    For i = 0 To original.m - 1
        If original.rhs(i) <> 0# Then Exit Function
    Next i
    loadValue = sourceField(RLoadVariable)
    If Not RPX_SchurFinite(loadValue) Or loadValue <= 0# Then Exit Function
    positiveScale = 1.000003 / loadValue
    If Not RPX_SchurFinite(positiveScale) Or positiveScale <= 0# Then Exit Function
    candidate = sourceField
    For i = 0 To original.n - 1
        If Not RPX_SchurFinite(candidate(i)) Then Exit Function
        candidate(i) = candidate(i) * positiveScale
        If Not RPX_SchurFinite(candidate(i)) Then Exit Function
    Next i
    If Not RPX_AccuracyPrimalAudit(original, candidate, objective) Then Exit Function
    If Not RPX_FieldGate("lower", candidate, original.factor, 8, objective, measured, audit, raw) Then Exit Function
    If Not RPX_SchurFinite(measured) Or measured < 1.000001 Then Exit Function
    recovered = candidate: objectiveOut = objective: measuredOut = measured
    HomogeneousLowerWitness = True
    Exit Function
RejectedRay:
    If err.number = 6 Or err.number = 9 Then
        err.Clear: Erase recovered
    Else
        err.Raise err.number, err.source, err.description
    End If
End Function

' An audited HSD recession iterate supplies only a same-Fs physical copy.
Public Function RPX_AccuracyLowerRayWitness(ByVal factor As Double, ByRef recovered() As Double) As Boolean
    Dim original As RPX_AccuracyProgram, objective As Double, measured As Double
    Erase recovered
    If XStatus <> "UNBOUNDED" Or Not RPX_AccuracyEnabled() Then Exit Function
    If Not RPX_AccuracyExport(factor, "lower", original) Then Exit Function
    If Not HomogeneousLowerWitness(original, xx, recovered, objective, measured) Then Exit Function
    AccuracyNativeValue = measured
    CaptureAccuracyProof original, recovered, False
    RPX_DiagEvent "accuracy_hsd_lower_ray_witness", "Fs=" & factor & ";q=8;lambda=" & measured & ";homogeneous_original_rhs=1;same_fs_original_field_gate=1;solver_status_used_as_evidence=0;original_dual_used=0"
    RPX_AccuracyLowerRayWitness = True
End Function

' Optional first solve of the literal original LP. Status is diagnostic only.
Public Function RPX_DirectEnabled() As Boolean
    Dim value As Double
    value = CDbl(RPX_Setting("DIRECT_DLL", 0))
    If Not RPX_SchurFinite(value) Or (value <> 0# And value <> 1#) Then err.Raise 5, "RPX_AccuracyNative", "INVALID_DIRECT_DLL"
    RPX_DirectEnabled = (value = 1# And RPX_AccuracyEnabled() And RFsActive And RSubdivisions = 8 And Not AFActive And RR = 0)
End Function
' Shared only with the optional fresh upper indicator solve inside a chain step.
Public Function RPX_ThinUpperDllSeconds(ByVal ordinary As Double, ByVal cones As Long) As Double
    Dim thinFloor As Double
    RPX_ThinUpperDllSeconds = ordinary
    If Not RPX_ThinChainBudgetActive Or Not RPX_ThinChainUpperActive Then Exit Function
    If cones <= 0 Then Exit Function
    thinFloor = RPX_ThinDllFactor * RPX_ThinDllRefSeconds * (CDbl(cones) / RPX_ThinDllRefCones) ^ RPX_ThinDllExponent
    If thinFloor > RPX_ThinDllCap Then thinFloor = RPX_ThinDllCap
    If thinFloor > ordinary Then RPX_ThinUpperDllSeconds = thinFloor
    RPX_DiagEvent "thin_layer_budget", "solver=dll_upper;mode=upper;cones=" & cones & ";ordinary=" & ordinary & ";floor=" & thinFloor & ";applied=" & RPX_ThinUpperDllSeconds
End Function
Public Function RPX_AccuracyDirectBound(ByVal factor As Double, ByVal mode As String) As Boolean
    Dim original As RPX_AccuracyProgram, candidate() As Double, originalDual() As Double
    Dim oldX() As Double, oldSlack() As Double, hadX As Boolean, hadSlack As Boolean, objective As Double, measured As Double, audit As Double, raw() As Double
    Dim seconds As Double, oldStatus As String, oldObjective As Double, oldValue As Double, oldAudit As Double
    Dim oldRejected As Boolean, oldGate As Boolean, oldReason As String, state As RPX_AuditSnapshot
    Dim ok As Boolean, accepted As Boolean, kind As String, number As Long, message As String, recovering As Boolean, helperUsed As Boolean
    Dim started As Double, remaining As Double
    Dim useDualFirst As Boolean, dualFirstAttempted As Boolean, dualFirstSeconds As Double
    Dim normalized() As Double, normalizedObjective As Double, normalizedValue As Double
    Dim recentered() As Double, useUpperRecenter As Boolean, hasOriginalCenter As Boolean, recenterInvoked As Boolean
    Dim recenterSolveId As Long, recenterCandidateId As Long, useDualRepairDirect As Boolean
    Dim oldRaw(1 To 4) As Double, oldDissipation As Double, oldReference As Double, oldFixed As Double, i As Long
    Dim iterateFirst As Long, iterateLast As Long
    Dim primalCallId As Long, primalStatus As Long, primalEquilibrate As Long, primalRole As String
    Dim primalAudited As Boolean, primalAuditPass As Boolean, primalEquality As Double, primalCone As Double, primalObjective As Double
    Dim recenterFirstCall As Long, candidateCallId As Long, endDetail As String
    Dim dualStatLog As Double, dualConeLog As Double, dualObjectiveLog As Double
    AccuracyDirectLastKind = "": AccuracyDirectLastAccepted = False
    ResetDualDiagnostics
    AccuracyNativeCallId = 0: AccuracyNativeSolverStatus = 0: AccuracyNativeIterations = 0: AccuracyNativeEquilibrate = -1
    AccuracyNativeEquality = 0#: AccuracyNativeCone = 0#: primalEquilibrate = -1: primalRole = "none"
    If RPX_ThinUpperProtectOn() Then Call RPX_ThinUpperProtectCheck("root_trial")
    useDualFirst = ConsumeUpperDualFirst(factor, mode)
    If Not RPX_DirectEnabled() Then Exit Function
    seconds = CDbl(RPX_Setting("DIRECT_DLL_SECONDS", 5))
    If Not RPX_SchurFinite(seconds) Or seconds <= 0# Or seconds > 600# Then err.Raise 5, "RPX_AccuracyNative", "INVALID_DIRECT_DLL_SECONDS"
    If Not RPX_AccuracyExport(factor, mode, original) Then Exit Function
    If RPX_ThinChainBudgetActive And RPX_ThinChainUpperActive And mode = "upper" Then seconds = RPX_ThinUpperDllSeconds(seconds, original.coneCount - 1)
    If mode = "upper" Then useUpperRecenter = UpperRecenterEnabled()
    If mode = "upper" Then useDualRepairDirect = DualLinearRepairEnabled()
    oldStatus = XStatus: oldObjective = XObjective: oldValue = RValue: oldAudit = RAudit
    oldRejected = RPX_AuditRejected: oldGate = RPhysicalGateAccepted: oldReason = RPhysicalGateReason
    oldDissipation = RDissipation: oldReference = RReferenceWork: oldFixed = RFixedWork
    For i = 1 To 4: oldRaw(i) = RRawLowerAudit(i): Next i
    RPX_DiagAuditSave state
    hadX = AccuracyArrayShape(xx, iterateFirst, iterateLast)
    If hadX Then hadX = (iterateFirst = 0)
    If hadX Then hadX = (iterateLast = original.n - 1)
    If hadX Then oldX = xx
    hadSlack = AccuracyArrayShape(XSlack, iterateFirst, iterateLast)
    If hadSlack Then hadSlack = (iterateLast >= 0)
    If hadSlack Then oldSlack = XSlack
    On Error GoTo Rejected
    ' Freshly compiled problems have no previous iterate; FieldGate restores this local zero vector.
    If Not hadX Then ReDim xx(0 To original.n - 1)
    RPX_DiagEvent "direct_dll_begin", "mode=" & mode & ";Fs=" & factor & ";q=8;P=0;original_cost=1"
    If useDualFirst Then
        ' A separately declared helper budget can precede the ordinary LP.
        ' Each budget remains explicit; no helper result bypasses original audit.
        dualFirstAttempted = True
        dualFirstSeconds = CDbl(RPX_Setting("ACCURACY_DLL_DUAL_SECONDS", 20))
        If Not RPX_SchurFinite(dualFirstSeconds) Or dualFirstSeconds <= 0# Or dualFirstSeconds > 600# Then err.Raise 5, "RPX_AccuracyNative", "INVALID_UPPER_DUAL_FIRST_SECONDS"
        RPX_AssertInputs "direct_upper_dual_first"
        RPX_TestHook "direct_upper_dual_first"
        If RPX_AccuracyRecoverDual(factor, mode, dualFirstSeconds) Then
            RPX_AssertInputs "direct_upper_dual_first_accepted"
            If Not CurrentProgram(original) Then err.Raise 5, "RPX_AccuracyNative", "UPPER_DUAL_FIRST_CONTEXT_CHANGED"
            kind = "original_dual_helper": accepted = True: helperUsed = True
            GoTo Restore
        End If
        Call RPX_AccuracyGuideReset
        RPX_AssertInputs "direct_upper_dual_first_fallback"
        If XCancel Then err.Raise 18, "RPX_AccuracyNative", "ANALYSIS_CANCELLED"
        If Not CurrentProgram(original) Then err.Raise 5, "RPX_AccuracyNative", "UPPER_DUAL_FIRST_CONTEXT_CHANGED"
        RPX_DiagEvent "direct_upper_dual_first", "accepted=False;fallback=original_lp;helper_seconds=" & dualFirstSeconds & ";lp_seconds=" & seconds & ";guide_is_not_field=1"
    End If
    ' A rejected dual-first helper cannot supply diagnostics for this LP candidate.
    ResetDualDiagnostics
    started = RPX_DiagClock()
    RPX_TestHook "direct_dll_begin"
    If RPX_ThinFsStatus5Allowed(factor) Then AccuracyNativeSolverStatus = 0
    If ThinUpperRetryScope(original) Then AccuracyNativeSolverStatus = 0
    hasOriginalCenter = RPX_AccuracySolve(original, seconds, candidate, originalDual)
    If mode = "upper" Then
        If AccuracyNativeSolverStatus = 5 Then hasOriginalCenter = ThinUpperStatus5Retry(original, seconds, candidate, originalDual, hasOriginalCenter)
    End If
    primalCallId = AccuracyNativeCallId: primalStatus = AccuracyNativeSolverStatus: primalEquilibrate = AccuracyNativeEquilibrate
    primalRole = "original_primal"
    If Not hasOriginalCenter Then
        If mode = "upper" And AccuracyNativeSolverStatus = 5 Then
            If RPX_ThinFsStatus5Allowed(factor) Then
                kind = "thin_upper_below_root": accepted = True
                GoTo Restore
            End If
        End If
        GoTo UpperRecoveryFallback
    End If
    RPX_TestHook "direct_dll_candidate"
    ok = RPX_AccuracyPrimalAudit(original, candidate, objective)
    primalAudited = True: primalAuditPass = ok
    primalEquality = AccuracyNativeEquality: primalCone = AccuracyNativeCone: primalObjective = objective
    If RPX_ThinChainBudgetActive And RPX_ThinChainUpperActive And mode = "upper" Then
        If AccuracyNativeSolverStatus <> 1 And AccuracyNativeSolverStatus <> 4 Then ok = False
    End If
    If ok Then ok = RPX_FieldGate(mode, candidate, factor, 8, objective, measured, audit, raw)
    If ok Then
        If mode = "upper" Then ok = (measured <= 0.999999) Else ok = (measured >= 1.000001)
    End If
    If Not ok And mode = "lower" Then
        If HomogeneousLowerWitness(original, candidate, normalized, normalizedObjective, normalizedValue) Then
            candidate = normalized: objective = normalizedObjective: measured = normalizedValue: ok = True
            primalAudited = True: primalAuditPass = True
            primalEquality = AccuracyNativeEquality: primalCone = AccuracyNativeCone: primalObjective = objective
            RPX_DiagEvent "direct_dll_homogeneous_primal", "mode=lower;Fs=" & factor & ";q=8;lambda=" & measured & ";same_fs_original_gate=1;solver_status_used_as_evidence=0;auxiliary_dual_used=0"
        End If
    End If
    If ok Then
        kind = "original_primal": accepted = True
        GoTo Restore
    End If
    If RPX_AccuracyDualAudit(original, originalDual) Then kind = "original_dual": accepted = True
    If AccuracyNativeDualAudited Then CaptureDualRepairStage original, originalDual, "raw_original", accepted
    If Not accepted And mode = "upper" And AccuracyNativeSolverStatus = 5 Then
        If RPX_ThinFsStatus5Allowed(factor) Then
            ' AlmostPrimalInfeasible is only a below-root search sign. It has
            ' no accepted primal field and cannot itself create an upper bound.
            kind = "thin_upper_below_root": accepted = True
            GoTo Restore
        End If
    End If
    If Not accepted And useDualRepairDirect Then
        If TryOriginalDualLinearRepair(original, originalDual) Then kind = "original_dual": accepted = True
    End If
    If Not accepted And Not dualFirstAttempted Then
        remaining = seconds - (RPX_DiagClock() - started)
        ' The LP objective schedules a helper only; it never supplies evidence.
        If remaining > 0.05 And (mode = "upper" Or objective > -1.000001) Then
            If RPX_AccuracyRecoverDual(factor, mode, remaining) Then
                kind = "original_dual_helper": accepted = True: helperUsed = True
            End If
        End If
    End If
UpperRecoveryFallback:
    If Not accepted And useUpperRecenter Then
        recenterInvoked = True: recenterSolveId = P4Solve: recenterCandidateId = RPX_HSDCandidateId
        Call RPX_AccuracyGuideReset
        recenterFirstCall = accuracyCallSequence
        If RecoverUpperRecenterProgram(original, candidate, hasOriginalCenter, recentered) Then
            candidate = recentered
            ' Reusing the original center keeps its native origin, even after a helper.
            If accuracyCallSequence <> recenterFirstCall Then
                primalCallId = AccuracyNativeCallId: primalStatus = AccuracyNativeSolverStatus: primalEquilibrate = AccuracyNativeEquilibrate
            End If
            primalRole = "recovered_primal"
            ok = RPX_AccuracyPrimalAudit(original, candidate, objective)
            primalAudited = True: primalAuditPass = ok
            primalEquality = AccuracyNativeEquality: primalCone = AccuracyNativeCone: primalObjective = objective
            If RPX_ThinChainBudgetActive And RPX_ThinChainUpperActive Then
                If AccuracyNativeSolverStatus <> 1 And AccuracyNativeSolverStatus <> 4 Then ok = False
            End If
            If ok Then ok = RPX_FieldGate("upper", candidate, factor, 8, objective, measured, audit, raw)
            If ok Then
                If Not RPX_SchurFinite(measured) Or Not RPX_SchurFinite(audit) Then ok = False
                If measured > 0.999999 Then ok = False
            End If
            UpperRecenterAssert original, recenterSolveId, recenterCandidateId, "direct_upper_recenter_commit"
            If ok Then kind = "original_primal": accepted = True
        End If
    End If
Restore:
    If hadX Then xx = oldX Else Erase xx
    XStatus = oldStatus: XObjective = oldObjective: RValue = oldValue: RAudit = oldAudit
    RPX_AuditRejected = oldRejected: RPhysicalGateAccepted = oldGate: RPhysicalGateReason = oldReason
    RPX_DiagAuditRestore state
    If number = 18 Or XCancel Then err.Raise 18, "RPX_AccuracyNative", "ANALYSIS_CANCELLED"
    If accepted Then
        If kind = "original_primal" Then
            xx = candidate: XObjective = objective
            If mode = "upper" Then
                Call RPX_AuditUpper: XStatus = "ORIGINAL_ACCURACY_UPPER_WITNESS"
            Else
                Call RPX_AuditLower: XStatus = "ORIGINAL_FIXED_LOWER_WITNESS"
            End If
            CaptureAccuracyProof original, candidate, False
        ElseIf kind = "thin_upper_below_root" Then
            XStatus = "THIN_UPPER_BELOW_ROOT": XObjective = 0#: RValue = 1E+30
            Erase xx: Erase XSlack
            RPX_AccuracyGuideReset
            RPX_DiagEvent "thin_layer_status5_sign", "mode=upper;Fs=" & factor & ";solver_status=5;dual_check=prior_original_status1_guide;witness=0;bound=0;fallback_hsd=0"
        Else
            guideReady = True: guideFactor = factor: guideMode = mode: guideModel = original.model: guidePolicy = original.policy
            guideEpoch = original.epoch: guideQ = original.q: guidePhase = original.phase
            guideN = XN: guideM = XM: guideS = XS: guideB = XB: guideSolve = P4Solve: guideCandidate = RPX_HSDCandidateId
            guideCost = original.originalCost: guideNormalizedCost = XQ
            XObjective = AccuracyNativeDualObjective
            If mode = "upper" Then
                XStatus = "ORIGINAL_ACCURACY_DUAL_GUIDE": RValue = AccuracyNativeDualObjective
            Else
                XStatus = "ORIGINAL_LOWER_DUAL_GUIDE": RValue = -AccuracyNativeDualObjective
            End If
            If Not RPX_AccuracyGuideCurrent(factor, mode) Then accepted = False: RPX_AccuracyGuideReset
            If accepted And Not helperUsed Then CaptureAccuracyProof original, originalDual, True
            Erase xx: Erase XSlack ' A guide never becomes a continuum witness.
        End If
    End If
    If Not accepted Then
        If hadX Then xx = oldX Else Erase xx
        If hadSlack Then XSlack = oldSlack Else Erase XSlack
        XStatus = oldStatus: XObjective = oldObjective: RValue = oldValue: RAudit = oldAudit
        RPX_AuditRejected = oldRejected: RPhysicalGateAccepted = oldGate: RPhysicalGateReason = oldReason
        RPX_DiagAuditRestore state
        RDissipation = oldDissipation: RReferenceWork = oldReference: RFixedWork = oldFixed
        For i = 1 To 4: RRawLowerAudit(i) = oldRaw(i): Next i
        RPX_AccuracyGuideReset
    End If
    AccuracyDirectLastKind = kind: AccuracyDirectLastAccepted = accepted
    On Error GoTo 0
    If number = 5 And message = "THIN_UPPER_LOWER_PROTECT" Then err.Raise number, "RPX_AccuracyNative", message
    If recenterInvoked And number <> 0 Then err.Raise number, "RPX_AccuracyNative", message
    If useDualRepairDirect And number <> 0 Then err.Raise number, "RPX_AccuracyNative", message
    candidateCallId = AccuracyNativeCallId
    If kind = "original_primal" Then candidateCallId = primalCallId
    ' Invalid partial audit values remain operationally unchanged and are not logged as residuals.
    If AccuracyNativeDualAudited Then
        dualStatLog = AccuracyNativeDualStat: dualConeLog = AccuracyNativeDualCone: dualObjectiveLog = AccuracyNativeDualObjective
    End If
    endDetail = "accepted=" & accepted & ";kind=" & kind & ";mode=" & mode & ";Fs=" & factor
    endDetail = endDetail & ";solver_status=" & primalStatus & ";raw_eq=" & primalEquality & ";raw_cone=" & primalCone
    endDetail = endDetail & ";primal_call_id=" & primalCallId & ";primal_status=" & primalStatus & ";primal_equilibrate=" & primalEquilibrate & ";primal_role=" & primalRole
    endDetail = endDetail & ";primal_audited=" & primalAudited & ";primal_audit_pass=" & primalAuditPass
    endDetail = endDetail & ";primal_raw_eq=" & primalEquality & ";primal_raw_cone=" & primalCone & ";primal_objective=" & primalObjective
    endDetail = endDetail & ";dual_audited=" & AccuracyNativeDualAudited & ";dual_call_id=" & AccuracyNativeDualAuditCallId
    endDetail = endDetail & ";dual_status=" & AccuracyNativeDualAuditStatus & ";dual_role=" & AccuracyNativeDualAuditRole
    endDetail = endDetail & ";dual_stat=" & dualStatLog & ";dual_cone=" & dualConeLog & ";dual_objective=" & dualObjectiveLog
    endDetail = endDetail & ";last_call_id=" & accuracyCallSequence & ";candidate_call_id=" & candidateCallId
    RPX_DiagEvent "direct_dll_end", endDetail & ";fallback=" & CStr(Not accepted) & ";message=" & message
    RPX_AccuracyDirectBound = accepted
    Exit Function
Rejected:
    number = err.number: message = err.description: accepted = False
    If recovering Then
        On Error GoTo 0
        err.Raise number, "RPX_AccuracyNative", "DIRECT_RESTORE_FAILED " & message
    End If
    recovering = True
    Resume Restore
End Function

' One-shot scheduling request only; never a primal or dual certificate.
Public Sub RPX_AccuracyUpperDualFirstReset()
    upperFirstReady = False: upperFirstModel = "": upperFirstPolicy = "": upperFirstPhase = ""
End Sub
Public Sub RPX_AccuracyUpperDualFirstArm(ByVal factor As Double)
    Call RPX_AccuracyUpperDualFirstReset
    If XCancel Then err.Raise 18, "RPX_AccuracyNative", "ANALYSIS_CANCELLED"
    If Not RPX_SchurFinite(factor) Or factor <= 0# Then err.Raise 5, "RPX_AccuracyNative", "INVALID_UPPER_DUAL_FIRST_FACTOR"
    If Not RPX_DirectEnabled() Or RSubdivisions <> 8 Or RPX_LoadPhase <> "reference_total" Then Exit Sub
    upperFirstFactor = factor: upperFirstEpoch = RPX_InputEpoch: upperFirstQ = RSubdivisions
    upperFirstModel = RPX_ExactModel(): upperFirstPolicy = RPX_AnalysisPolicy(): upperFirstPhase = RPX_LoadPhase
    upperFirstSolve = P4Solve: upperFirstCandidate = RPX_HSDCandidateId: upperFirstReady = True
End Sub
Private Function ConsumeUpperDualFirst(ByVal factor As Double, ByVal mode As String) As Boolean
    If Not upperFirstReady Then GoTo Consumed
    If mode <> "upper" Or Not RPX_DirectEnabled() Or XCancel Then GoTo Consumed
    If Not RPX_SchurFinite(factor) Or factor <> upperFirstFactor Then GoTo Consumed
    If upperFirstEpoch <> RPX_InputEpoch Or upperFirstQ <> RSubdivisions Or RSubdivisions <> 8 Then GoTo Consumed
    If upperFirstModel <> RPX_ExactModel() Or upperFirstPolicy <> RPX_AnalysisPolicy() Or upperFirstPhase <> RPX_LoadPhase Then GoTo Consumed
    If upperFirstSolve <> P4Solve Or upperFirstCandidate <> RPX_HSDCandidateId Then GoTo Consumed
    ConsumeUpperDualFirst = True
Consumed:
    Call RPX_AccuracyUpperDualFirstReset
End Function

' Opt-in candidate generation only. The legacy path and all original gates stay intact.
Private Function UpperRecenterEnabled() As Boolean
    Dim enabled As Double
    enabled = CDbl(RPX_Setting("ACCURACY_DLL_UPPER_RECENTER", 0))
    If Not RPX_SchurFinite(enabled) Or (enabled <> 0# And enabled <> 1#) Then err.Raise 5, "RPX_AccuracyNative", "INVALID_ACCURACY_DLL_UPPER_RECENTER"
    UpperRecenterEnabled = (enabled = 1#)
End Function
Private Function UpperWorkRow(ByRef original As RPX_AccuracyProgram, ByRef workRow As Long) As Boolean
    Dim i As Long, count As Long
    workRow = -1
    If original.mode <> "upper" Or original.q <> 8 Or original.phase <> "reference_total" Then Exit Function
    If original.n <= 0 Or original.equalityCount <= 0 Or original.m <= original.equalityCount Then Exit Function
    For i = 0 To original.equalityCount - 1
        If Not RPX_SchurFinite(original.rhs(i)) Then Exit Function
        If original.rhs(i) <> 0# Then count = count + 1: workRow = i
    Next i
    If count <> 1 Then Exit Function
    For i = original.equalityCount To original.m - 1
        If Not RPX_SchurFinite(original.rhs(i)) Or original.rhs(i) <> 0# Then Exit Function
    Next i
    UpperWorkRow = True
End Function
Private Sub UpperRecenterAssert(ByRef original As RPX_AccuracyProgram, ByVal solveId As Long, ByVal candidateId As Long, ByVal stage As String)
    If XCancel Then err.Raise 18, "RPX_AccuracyNative", "ANALYSIS_CANCELLED"
    RPX_AssertInputs stage
    If Not CurrentProgram(original) Or P4Solve <> solveId Or RPX_HSDCandidateId <> candidateId Then err.Raise 5, "RPX_AccuracyNative", "UPPER_RECENTER_CONTEXT_CHANGED"
End Sub
Private Function NormalizeUpperWork(ByRef original As RPX_AccuracyProgram, ByRef candidate() As Double, ByRef workRatio As Double) As Boolean
    Dim workRow As Long, j As Long, i As Long, total As Double, correction As Double, normalized() As Double
    On Error GoTo Rejected
    workRatio = 0#
    If Not UpperWorkRow(original, workRow) Then Exit Function
    If LBound(candidate) <> 0 Or UBound(candidate) <> original.n - 1 Then Exit Function
    For i = 0 To original.n - 1: If Not RPX_SchurFinite(candidate(i)) Then Exit Function
    Next i
    For j = original.rowPtr(workRow) To original.rowPtr(workRow + 1) - 1
        AddTerm total, correction, original.values(j) * candidate(original.colId(j))
    Next j
    workRatio = total / original.rhs(workRow)
    If Not RPX_SchurFinite(workRatio) Or workRatio <= 0# Then Exit Function
    normalized = candidate
    For i = 0 To original.n - 1
        normalized(i) = candidate(i) / workRatio
        If Not RPX_SchurFinite(normalized(i)) Then Exit Function
    Next i
    candidate = normalized: NormalizeUpperWork = True
    Exit Function
Rejected:
    If err.number <> 6 And err.number <> 9 And Left$(err.description, 18) <> "ACCURACY_NONFINITE" Then err.Raise err.number, err.source, err.description
    err.Clear
End Function
Private Function AuditUpperRecenter(ByRef original As RPX_AccuracyProgram, ByRef candidate() As Double, ByVal solveId As Long, ByVal candidateId As Long, ByVal stage As String, ByRef workRatio As Double) As Boolean
    Dim objective As Double, raw() As Double, accepted As Boolean
    UpperRecenterAssert original, solveId, candidateId, stage & "_before_audit"
    If NormalizeUpperWork(original, candidate, workRatio) Then
        accepted = RPX_AccuracyPrimalAudit(original, candidate, objective)
        If RPX_ThinChainBudgetActive And RPX_ThinChainUpperActive Then
            If AccuracyNativeSolverStatus <> 1 And AccuracyNativeSolverStatus <> 4 Then accepted = False
        End If
        If accepted Then accepted = RPX_FieldGate("upper", candidate, original.factor, 8, objective, AccuracyNativeValue, AccuracyNativeAudit, raw)
        If accepted Then
            If Not RPX_SchurFinite(AccuracyNativeValue) Or Not RPX_SchurFinite(AccuracyNativeAudit) Then accepted = False
            If AccuracyNativeValue > 0.999999 Then accepted = False
        End If
    End If
    UpperRecenterAssert original, solveId, candidateId, stage & "_after_audit"
    RPX_DiagEvent "upper_recenter_candidate", "stage=" & stage & ";Fs=" & original.factor & ";work_ratio=" & workRatio & ";accepted=" & accepted & ";raw_eq=" & AccuracyNativeEquality & ";raw_cone=" & AccuracyNativeCone & ";value=" & AccuracyNativeValue & ";modified_dual_used=0;root_created=0"
    AuditUpperRecenter = accepted
End Function
Private Function RecoverUpperRecenter(ByVal factor As Double, ByRef recovered() As Double) As Boolean
    Dim original As RPX_AccuracyProgram, blankCenter() As Double
    Erase recovered
    If XCancel Then err.Raise 18, "RPX_AccuracyNative", "ANALYSIS_CANCELLED"
    RPX_AssertInputs "upper_recenter_export"
    If Not RPX_AccuracyExport(factor, "upper", original) Then Exit Function
    RecoverUpperRecenter = RecoverUpperRecenterProgram(original, blankCenter, False, recovered)
End Function
Private Function UpperFiniteCenter(ByRef original As RPX_AccuracyProgram, ByRef candidate() As Double) As Boolean
    Dim i As Long
    On Error GoTo Rejected
    If LBound(candidate) <> 0 Or UBound(candidate) <> original.n - 1 Then Exit Function
    For i = 0 To original.n - 1: If Not RPX_SchurFinite(candidate(i)) Then Exit Function
    Next i
    UpperFiniteCenter = True: Exit Function
Rejected:
    If err.number <> 9 Then err.Raise err.number, err.source, err.description
    err.Clear
End Function
Private Function RecoverUpperRecenterProgram(ByRef original As RPX_AccuracyProgram, ByRef initialCenter() As Double, ByVal hasInitialCenter As Boolean, ByRef recovered() As Double) As Boolean
    Dim qp As RPX_AccuracyProgram
    Dim center() As Double, candidate() As Double, modifiedDual() As Double
    Dim alpha As Double, lpSeconds As Double, qpSeconds As Double, workRatio As Double, factor As Double
    Dim i As Long, trialIndex As Long, workRow As Long, solveId As Long, candidateId As Long
    Dim ok As Boolean, finiteCandidate As Boolean, oldRejected As Boolean, auditSaved As Boolean
    Dim auditState As RPX_AuditSnapshot, number As Long, message As String, errorSource As String
    Erase recovered
    If XCancel Then err.Raise 18, "RPX_AccuracyNative", "ANALYSIS_CANCELLED"
    RPX_AssertInputs "upper_recenter_program"
    If Not CurrentProgram(original) Then err.Raise 5, "RPX_AccuracyNative", "UPPER_RECENTER_CONTEXT_CHANGED"
    If Not UpperWorkRow(original, workRow) Then AccuracyNativeReason = "UPPER_RECENTER_NONHOMOGENEOUS_SCOPE_REJECTED": Exit Function
    factor = original.factor
    solveId = P4Solve: candidateId = RPX_HSDCandidateId
    alpha = CDbl(RPX_Setting("ACCURACY_DLL_ALPHA", 0.01))
    lpSeconds = CDbl(RPX_Setting("ACCURACY_DLL_UPPER_LP_SECONDS", 30))
    qpSeconds = CDbl(RPX_Setting("ACCURACY_DLL_SECONDS", 120))
    If Not RPX_SchurFinite(alpha) Or alpha <= 0# Then err.Raise 5, "RPX_AccuracyNative", "INVALID_ACCURACY_DLL_ALPHA"
    If Not RPX_SchurFinite(lpSeconds) Or lpSeconds <= 0# Or lpSeconds > 600# Then err.Raise 5, "RPX_AccuracyNative", "INVALID_UPPER_RECENTER_LP_SECONDS"
    If Not RPX_SchurFinite(qpSeconds) Or qpSeconds <= 0# Or qpSeconds > 600# Then err.Raise 5, "RPX_AccuracyNative", "INVALID_UPPER_RECENTER_QP_SECONDS"
    RPX_DiagAuditSave auditState: oldRejected = RPX_AuditRejected: auditSaved = True
    On Error GoTo Failed
    If hasInitialCenter Then
        If UpperFiniteCenter(original, initialCenter) Then
            candidate = initialCenter: finiteCandidate = True
            RPX_DiagEvent "upper_recenter_center", "source=reused_direct_original_lp;same_exported_program=1;Fs=" & factor
        End If
    End If
    If Not finiteCandidate Then
        RPX_DiagEvent "upper_recenter_center", "source=fresh_original_lp;reason=direct_center_absent_or_invalid;Fs=" & factor
        UpperRecenterAssert original, solveId, candidateId, "upper_recenter_lp_before"
        RPX_TestHook "upper_recenter_lp_before"
        UpperRecenterAssert original, solveId, candidateId, "upper_recenter_lp_start"
        finiteCandidate = RPX_AccuracySolve(original, lpSeconds, candidate, modifiedDual)
        Erase modifiedDual ' Neither this LP dual nor a modified QP dual is committed here.
        UpperRecenterAssert original, solveId, candidateId, "upper_recenter_lp_after"
    End If
    If Not finiteCandidate Then GoTo Restore
    ok = AuditUpperRecenter(original, candidate, solveId, candidateId, "original_lp", workRatio)
    If ok Then GoTo Restore
    center = candidate
    For trialIndex = 1 To 2
        UpperRecenterAssert original, solveId, candidateId, "upper_recenter_qp_before"
        RPX_TestHook "upper_recenter_qp_before"
        UpperRecenterAssert original, solveId, candidateId, "upper_recenter_qp_start"
        qp = original: AddBound qp, -1, 0.999998
        For i = 0 To original.n - 1
            qp.pdiag(i) = alpha: qp.cost(i) = -alpha * center(i)
            If Not RPX_SchurFinite(qp.cost(i)) Then err.Raise 5, "RPX_AccuracyNative", "UPPER_RECENTER_NONFINITE_QP_CENTER"
        Next i
        ' Literal original rows plus the cap: Python recipe uses no cone row scaling.
        finiteCandidate = RPX_AccuracySolve(qp, qpSeconds, candidate, modifiedDual)
        Erase modifiedDual ' An auxiliary cap/QP dual is NEVER an original root guide.
        UpperRecenterAssert original, solveId, candidateId, "upper_recenter_qp_after"
        If Not finiteCandidate Then GoTo Restore
        ok = AuditUpperRecenter(original, candidate, solveId, candidateId, "cap_qp_" & trialIndex, workRatio)
        If ok Then Exit For
        center = candidate ' Same Fs and the normalized rejected candidate only; never an old field.
    Next trialIndex
Restore:
    Erase modifiedDual
    If auditSaved Then
        RPX_DiagAuditRestore auditState: RPX_AuditRejected = oldRejected
    End If
    On Error GoTo 0
    If number <> 0 Then err.Raise number, errorSource, message
    UpperRecenterAssert original, solveId, candidateId, "upper_recenter_return"
    If ok Then
        recovered = candidate: AccuracyNativeReason = "ORIGINAL_PHYSICAL_PRIMAL_WITNESS"
        CaptureAccuracyProof original, candidate, False
    Else
        Erase recovered: AccuracyNativeReason = "ORIGINAL_PHYSICAL_PRIMAL_REJECTED"
    End If
    RPX_DiagEvent "upper_recenter_complete", "accepted=" & ok & ";Fs=" & factor & ";q=8;max_qp_trials=2;lp_seconds=" & lpSeconds & ";qp_seconds_each=" & qpSeconds & ";modified_dual_used=0;root_created=0"
    RecoverUpperRecenterProgram = ok: Exit Function
Failed:
    number = err.number: message = err.description: errorSource = err.source: ok = False
    Resume Restore
End Function

' The cone dual is immutable. A correction is a guide only after the original audit.
Private Function DualLinearRepairEnabled() As Boolean
    Dim enabled As Double
    enabled = CDbl(RPX_Setting("ACCURACY_DLL_DUAL_LINEAR_REPAIR", 0))
    If Not RPX_SchurFinite(enabled) Or (enabled <> 0# And enabled <> 1#) Then err.Raise 5, "RPX_AccuracyNative", "INVALID_DUAL_LINEAR_REPAIR"
    DualLinearRepairEnabled = (enabled = 1#)
End Function
Private Sub SetUpperDualRegularization(ByRef dp As RPX_AccuracyProgram)
    Dim coefficient As Double, i As Long
    coefficient = CDbl(RPX_Setting("ACCURACY_DLL_DUAL_REGULARIZATION", 0))
    If Not RPX_SchurFinite(coefficient) Or coefficient < 0# Then err.Raise 5, "RPX_AccuracyNative", "INVALID_DUAL_REGULARIZATION"
    For i = 0 To dp.n - 1: dp.pdiag(i) = coefficient: Next i
End Sub
Private Sub DualRepairAssert(ByRef original As RPX_AccuracyProgram, ByVal solveId As Long, ByVal candidateId As Long, ByVal stage As String)
    If XCancel Then err.Raise 18, "RPX_AccuracyNative", "ANALYSIS_CANCELLED"
    RPX_AssertInputs stage
    If Not CurrentProgram(original) Or P4Solve <> solveId Or RPX_HSDCandidateId <> candidateId Then err.Raise 5, "RPX_AccuracyNative", "DUAL_REPAIR_CONTEXT_CHANGED"
End Sub
Private Sub OriginalEqualityTranspose(ByRef original As RPX_AccuracyProgram, ByRef transposed As RPX_AccuracyProgram)
    Dim counts() As Long, cursor() As Long, i As Long, j As Long, row As Long, p As Long
    transposed.n = original.equalityCount: transposed.m = original.n
    transposed.nnz = original.rowPtr(original.equalityCount)
    ReDim counts(0 To original.n - 1): ReDim cursor(0 To original.n - 1)
    For j = 0 To transposed.nnz - 1: counts(original.colId(j)) = counts(original.colId(j)) + 1: Next j
    ReDim transposed.rowPtr(0 To original.n)
    For i = 0 To original.n - 1
        transposed.rowPtr(i + 1) = transposed.rowPtr(i) + counts(i): cursor(i) = transposed.rowPtr(i)
    Next i
    ReDim transposed.colId(0 To transposed.nnz - 1): ReDim transposed.values(0 To transposed.nnz - 1)
    For i = 0 To original.equalityCount - 1
        For j = original.rowPtr(i) To original.rowPtr(i + 1) - 1
            row = original.colId(j): p = cursor(row): cursor(row) = p + 1
            transposed.colId(p) = i: transposed.values(p) = original.values(j)
        Next j
    Next i
End Sub
Private Function DualConeBitsUnchanged(ByRef source() As Double, ByRef updated() As Double, ByVal first As Long, ByVal count As Long) As Boolean
    Dim i As Long, sourceBits As LongLong, updatedBits As LongLong
    For i = first To count - 1
        If Not RPX_SchurFinite(source(i)) Or Not RPX_SchurFinite(updated(i)) Then Exit Function
        If source(i) <> updated(i) Then Exit Function
        ' Equal finite nonzero doubles have identical bits; distinguish +/-0 explicitly.
        If source(i) = 0# Then
            CopyBytes VarPtr(sourceBits), VarPtr(source(i)), 8
            CopyBytes VarPtr(updatedBits), VarPtr(updated(i)), 8
            If sourceBits <> updatedBits Then Exit Function
        End If
    Next i
    DualConeBitsUnchanged = True
End Function
Private Sub CaptureDualRepairStage(ByRef original As RPX_AccuracyProgram, ByRef candidate() As Double, ByVal stage As String, ByVal accepted As Boolean)
    Dim enabled As Double, path As String
    enabled = CDbl(RPX_Setting("ACCURACY_DLL_CAPTURE", 0))
    If Not RPX_SchurFinite(enabled) Or (enabled <> 0# And enabled <> 1#) Then err.Raise 5, "RPX_AccuracyNative", "INVALID_ACCURACY_DLL_CAPTURE"
    If enabled = 0# Then Exit Sub
    proofSequence = proofSequence + 1
    ' Distinct filename: rejected source/corrected candidates are NOT accepted accuracy proofs.
    path = RPX_DiagPath & ".dual_repair_" & stage & "_solve_" & P4Solve & "_" & proofSequence & ".bin"
    RPX_AccuracyWriteProof path, original, candidate
    RPX_DiagEvent "dual_repair_capture", "stage=" & stage & ";path=" & path & ";layout=AC18P001_original_dual;accepted=" & accepted & ";physical_bound=0;candidate_only=1" & ";call_id=" & AccuracyNativeCallId
End Sub
Private Function OriginalDualArrayValid(ByRef original As RPX_AccuracyProgram, ByRef candidate() As Double) As Boolean
    Dim i As Long
    On Error GoTo Rejected
    If LBound(candidate) <> 0 Or UBound(candidate) <> original.m - 1 Then Exit Function
    For i = 0 To original.m - 1: If Not RPX_SchurFinite(candidate(i)) Then Exit Function
    Next i
    OriginalDualArrayValid = True: Exit Function
Rejected:
    If err.number <> 9 Then err.Raise err.number, err.source, err.description
    err.Clear
End Function
Private Function TryOriginalDualLinearRepair(ByRef original As RPX_AccuracyProgram, ByRef candidate() As Double) As Boolean
    Dim transposed As RPX_AccuracyProgram, sourceDual() As Double, updated() As Double
    Dim rhs() As Double, compensation() As Double, zeroCorrection() As Double, correction() As Double
    Dim descriptor(0 To 63) As Byte, report(0 To 55) As Byte
    Dim i As Long, j As Long, rc As Long, solveId As Long, candidateId As Long, number As Long
    Dim beforeStat As Double, beforeCone As Double, beforeObjective As Double, started As Double, seconds As Double
    Dim message As String, errorSource As String, accepted As Boolean, unchangedCone As Boolean, alreadyAccepted As Boolean
    If original.mode <> "upper" Then Exit Function
    If Not DualLinearRepairEnabled() Then Exit Function
    If XCancel Then err.Raise 18, "RPX_AccuracyNative", "ANALYSIS_CANCELLED"
    solveId = P4Solve: candidateId = RPX_HSDCandidateId
    DualRepairAssert original, solveId, candidateId, "dual_repair_begin"
    If original.equalityCount <= 0 Or original.rowPtr(original.equalityCount) <= 0 Then Exit Function
    If Not OriginalDualArrayValid(original, candidate) Then Exit Function
    On Error GoTo Failed
    sourceDual = candidate
    alreadyAccepted = RPX_AccuracyDualAudit(original, sourceDual, "repair_source")
    beforeStat = AccuracyNativeDualStat: beforeCone = AccuracyNativeDualCone: beforeObjective = AccuracyNativeDualObjective
    If alreadyAccepted Then Exit Function
    If Not RPX_SchurFinite(beforeStat) Or Not RPX_SchurFinite(beforeCone) Or Not RPX_SchurFinite(beforeObjective) Then Exit Function
    If beforeCone > 0.00000001 Then Exit Function ' y cannot repair a rejected cone dual.
    CaptureDualRepairStage original, sourceDual, "before", False
    OriginalEqualityTranspose original, transposed
    ReDim rhs(0 To original.n - 1): ReDim compensation(0 To original.n - 1)
    ReDim zeroCorrection(0 To original.equalityCount - 1): ReDim correction(0 To original.equalityCount - 1)
    For i = 0 To original.n - 1: rhs(i) = -original.originalCost(i): Next i
    For i = 0 To original.m - 1
        For j = original.rowPtr(i) To original.rowPtr(i + 1) - 1
            AddTerm rhs(original.colId(j)), compensation(original.colId(j)), -original.values(j) * sourceDual(i)
        Next j
    Next i
    PackLong descriptor, 0, 64: PackLong descriptor, 4, 1
    PackLong descriptor, 8, transposed.n: PackLong descriptor, 12, transposed.m: PackLong descriptor, 16, transposed.nnz
    PackPointer descriptor, 24, VarPtr(transposed.rowPtr(0)): PackPointer descriptor, 32, VarPtr(transposed.colId(0))
    PackPointer descriptor, 40, VarPtr(transposed.values(0)): PackPointer descriptor, 48, VarPtr(rhs(0)): PackPointer descriptor, 56, VarPtr(zeroCorrection(0))
    PackLong report, 0, 56: PackLong report, 4, 1
    DualRepairAssert original, solveId, candidateId, "dual_repair_before_lsmr"
    RPX_TestHook "dual_repair_before_lsmr"
    DualRepairAssert original, solveId, candidateId, "dual_repair_start_lsmr"
    started = RPX_DiagClock()
    rc = NEqualities(VarPtr(descriptor(0)), 200, 0.000000000000001, 0.000000000000001, 1000000000000#, VarPtr(correction(0)), original.equalityCount, VarPtr(report(0)))
    seconds = RPX_DiagClock() - started
    DualRepairAssert original, solveId, candidateId, "dual_repair_after_lsmr"
    If rc <> 0 Then AccuracyNativeReason = "ORIGINAL_DUAL_LINEAR_REPAIR_API_REJECTED": Exit Function
    updated = sourceDual
    For i = 0 To original.equalityCount - 1
        If Not RPX_SchurFinite(correction(i)) Then Exit Function
        updated(i) = sourceDual(i) + correction(i)
        If Not RPX_SchurFinite(updated(i)) Then Exit Function
    Next i
    unchangedCone = DualConeBitsUnchanged(sourceDual, updated, original.equalityCount, original.m)
    If Not unchangedCone Then err.Raise 5, "RPX_AccuracyNative", "DUAL_REPAIR_CHANGED_CONE_BITS"
    accepted = RPX_AccuracyDualAudit(original, updated, "repair_corrected")
    If Not RPX_SchurFinite(seconds) Or seconds < 0# Or seconds > 10# Then accepted = False
    DualRepairAssert original, solveId, candidateId, "dual_repair_after_original_audit"
    CaptureDualRepairStage original, updated, "after", accepted
    RPX_DiagEvent "original_dual_linear_repair", "accepted=" & accepted & ";Fs=" & original.factor & ";q=8;before_stat=" & beforeStat & ";before_cone=" & beforeCone & ";before_objective=" & beforeObjective & ";after_stat=" & AccuracyNativeDualStat & ";after_cone=" & AccuracyNativeDualCone & ";after_objective=" & AccuracyNativeDualObjective & ";seconds=" & seconds & ";stop=" & ReadLong(report, 8) & ";iterations=" & ReadLong(report, 12) & ";max_iterations=200;max_seconds=10;attempts=1;cone_bits_unchanged=" & unchangedCone & ";physical_bound=0;auxiliary_dual_used=0"
    If accepted Then candidate = updated: TryOriginalDualLinearRepair = True
    Exit Function
Failed:
    number = err.number: message = err.description: errorSource = err.source
    If number = 453 Then AccuracyNativeReason = "ACCURACY_DLL_EQUALITY_ENTRY_UNAVAILABLE": Exit Function
    err.Raise number, errorSource, message
End Function
