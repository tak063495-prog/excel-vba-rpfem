Attribute VB_Name = "RPX_AccuracyNative"
Option Explicit
#If Win64 Then
Private Declare PtrSafe Function NAbi Lib "rpfem_accuracy_x64.dll" Alias "rpfem_accuracy_abi_version" () As Long
Private Declare PtrSafe Function NDefaults Lib "rpfem_accuracy_x64.dll" Alias "rpfem_accuracy_default_settings" (ByVal settings As LongPtr) As Long
Private Declare PtrSafe Function NCreate Lib "rpfem_accuracy_x64.dll" Alias "rpfem_accuracy_create" (ByVal problem As LongPtr, ByVal settings As LongPtr, ByRef handle As LongLong) As Long
Private Declare PtrSafe Function NStart Lib "rpfem_accuracy_x64.dll" Alias "rpfem_accuracy_start" (ByVal handle As LongLong) As Long
Private Declare PtrSafe Function NPoll Lib "rpfem_accuracy_x64.dll" Alias "rpfem_accuracy_poll" (ByVal handle As LongLong, ByVal progress As LongPtr) As Long
Private Declare PtrSafe Function NCancel Lib "rpfem_accuracy_x64.dll" Alias "rpfem_accuracy_cancel" (ByVal handle As LongLong) As Long
Private Declare PtrSafe Function NResult Lib "rpfem_accuracy_x64.dll" Alias "rpfem_accuracy_result" (ByVal handle As LongLong, ByVal x As LongPtr, ByVal nx As Long, ByVal z As LongPtr, ByVal nz As Long, ByVal progress As LongPtr) As Long
Private Declare PtrSafe Function NDestroy Lib "rpfem_accuracy_x64.dll" Alias "rpfem_accuracy_destroy" (ByVal handle As LongLong) As Long
Private Declare PtrSafe Function NEqualities Lib "rpfem_accuracy_x64.dll" Alias "rpfem_accuracy_equalities" (ByVal problem As LongPtr, ByVal iterations As Long, ByVal atol As Double, ByVal btol As Double, ByVal conlim As Double, ByVal correction As LongPtr, ByVal count As Long, ByVal report As LongPtr) As Long
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
Private guideReady As Boolean, guideModel As String, guidePolicy As String, guideMode As String, guidePhase As String
Private guideEpoch As Long, guideQ As Long, guideFactor As Double
Private guideN As Long, guideM As Long, guideS As Long, guideB As Long, guideSolve As Long, guideCandidate As Long
Private guideCost() As Double, guideNormalizedCost() As Double
Private libraryHandle As LongPtr, libraryPath As String
Private proofSequence As Long

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
Public Function RPX_AccuracySolve(ByRef pr As RPX_AccuracyProgram, ByVal seconds As Double, ByRef answer() As Double, ByRef modifiedDual() As Double) As Boolean
    Dim descriptor(0 To 87) As Byte, settings(0 To 127) As Byte, progress(0 To 79) As Byte
    Dim handle As LongLong, rc As Long, state As Long, number As Long, message As String, source As String
    Dim cancelled As Boolean, valid As Boolean, i As Long, begin As Double, elapsed As Double, equ As Double
    Erase answer: Erase modifiedDual: AccuracyNativeReason = "DLL_NOT_AVAILABLE"
    If Not RPX_SchurFinite(seconds) Or seconds <= 0# Or seconds > 600# Then Exit Function
    On Error GoTo Failed
    If libraryHandle = 0 Then
        libraryPath = ThisWorkbook.path & "\rpfem_accuracy_x64.dll"
        libraryHandle = LoadAccuracyLibrary(StrPtr(libraryPath))
        If libraryHandle = 0 Then Exit Function
    ElseIf libraryPath <> ThisWorkbook.path & "\rpfem_accuracy_x64.dll" Then
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
    PackDouble settings, 16, seconds
    PackDouble settings, 24, 0.0000000000001: PackDouble settings, 32, 0.0000000000001: PackDouble settings, 40, 0.0000000000001
    PackDouble settings, 64, 0.000000000000001: PackDouble settings, 72, 0.000000000000001: PackLong settings, 80, 30
    If pr.mode = "lower" Or pr.mode = "upper" Then
        equ = CDbl(RPX_Setting("ACCURACY_DLL_EQUILIBRATE", 0))
    Else
        equ = CDbl(RPX_Setting("ACCURACY_DLL_DUAL_EQUILIBRATE", 1))
    End If
    If Not RPX_SchurFinite(equ) Or (equ <> 0# And equ <> 1#) Then err.Raise 5, "RPX_AccuracyNative", "INVALID_ACCURACY_DLL_EQUILIBRATE"
    PackLong settings, 84, CLng(equ)
    PackLong progress, 0, 80: PackLong progress, 4, 1
    rc = NCreate(VarPtr(descriptor(0)), VarPtr(settings(0)), handle)
    If rc <> 0 Then GoTo Cleanup
    rc = NStart(handle): If rc <> 0 Then GoTo Cleanup
    begin = Timer
    Do
        rc = NPoll(handle, VarPtr(progress(0))): If rc <> 0 Then GoTo Cleanup
        state = ReadLong(progress, 8): If state <> 1 And state <> 0 Then Exit Do
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
    If valid Then AccuracyNativeReason = "CANDIDATE_REQUIRES_ORIGINAL_AUDIT" Else AccuracyNativeReason = "ACCURACY_DLL_NO_FINITE_CANDIDATE"
    If XCancel Then err.Raise 18, "RPX_AccuracyNative", "ANALYSIS_CANCELLED"
    If number <> 0 And number <> 48 And number <> 53 And number <> 453 Then err.Raise number, source, message
    RPX_AccuracySolve = valid: Exit Function
Failed:
    number = err.number: message = err.description: source = err.source
    Resume Cleanup
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
Public Function RPX_AccuracyRecoverPrimal(ByVal factor As Double, ByVal mode As String, ByRef recovered() As Double) As Boolean
    Dim original As RPX_AccuracyProgram, qp As RPX_AccuracyProgram, center() As Double, candidate() As Double, modifiedDual() As Double
    Dim alpha As Double, seconds As Double, objective As Double, raw() As Double, i As Long, ok As Boolean, positiveScale As Double, homogeneous As Boolean
    Dim auditState As RPX_AuditSnapshot, number As Long, message As String, source As String, oldRejected As Boolean
    Erase recovered
    If Not RPX_AccuracyExport(factor, mode, original) Then Exit Function
    qp = original: ReDim center(0 To original.n - 1)
    On Error Resume Next
    If LBound(xx) = 0 And UBound(xx) = original.n - 1 Then center = xx
    err.Clear: On Error GoTo 0
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
    Exit Function
Failed:
    number = err.number: message = err.description: errorSource = err.source
    If number = 453 Then AccuracyNativeReason = "ACCURACY_DLL_EQUALITY_ENTRY_UNAVAILABLE": Exit Function
    err.Raise number, errorSource, message
End Function
Public Sub RPX_AccuracyDualProblem(ByRef original As RPX_AccuracyProgram, ByRef dualProgram As RPX_AccuracyProgram)
    Dim pr As RPX_AccuracyProgram, counts() As Long, cursor() As Long, i As Long, j As Long, row As Long, col As Long, p As Long, threshold As Double
    pr.n = original.m: pr.m = original.n + original.m - original.equalityCount + 1
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
Public Function RPX_AccuracyDualAudit(ByRef original As RPX_AccuracyProgram, ByRef candidate() As Double) As Boolean
    Dim stat() As Double, corr() As Double, i As Long, j As Long, k As Long, row As Long, total As Double, correction As Double, violation As Double, threshold As Double
    AccuracyNativeDualStat = 0#: AccuracyNativeDualCone = 0#: AccuracyNativeDualObjective = 0#
    On Error GoTo Rejected
    If Not CurrentProgram(original) Or LBound(candidate) <> 0 Or UBound(candidate) <> original.m - 1 Then Exit Function
    ReDim stat(0 To original.n - 1): ReDim corr(0 To original.n - 1): stat = original.originalCost
    For i = 0 To original.m - 1
        If Not RPX_SchurFinite(candidate(i)) Then Exit Function
        AddTerm AccuracyNativeDualObjective, correction, -original.rhs(i) * candidate(i)
        For j = original.rowPtr(i) To original.rowPtr(i + 1) - 1: AddTerm stat(original.colId(j)), corr(original.colId(j)), original.values(j) * candidate(i): Next j
    Next i
    For i = 0 To original.n - 1: AccuracyNativeDualStat = WorksheetFunction.Max(AccuracyNativeDualStat, Abs(stat(i))): Next i
    row = original.equalityCount
    For k = 1 To original.coneCount - 1
        violation = -candidate(row)
        If original.coneDim(k) = 3 Then violation = Norm2(candidate(row + 1), candidate(row + 2)) - candidate(row)
        AccuracyNativeDualCone = WorksheetFunction.Max(AccuracyNativeDualCone, violation): row = row + original.coneDim(k)
    Next k
    If original.mode = "upper" Then threshold = 1.000003 Else threshold = -0.999998
    RPX_AccuracyDualAudit = AccuracyNativeDualStat <= 0.0000001 And AccuracyNativeDualCone <= 0.00000001 And AccuracyNativeDualObjective >= threshold
    Exit Function
Rejected:
    If err.number <> 6 And err.number <> 9 And Left$(err.description, 18) <> "ACCURACY_NONFINITE" Then err.Raise err.number, err.source, err.description
    err.Clear: AccuracyNativeReason = "INVALID_ORIGINAL_DUAL"
End Function
Public Function RPX_AccuracyRecoverDual(ByVal factor As Double, ByVal mode As String) As Boolean
    Dim original As RPX_AccuracyProgram, dp As RPX_AccuracyProgram, candidate() As Double, helperDual() As Double, seconds As Double, accepted As Boolean
    RPX_AccuracyGuideReset
    If Not RPX_AccuracyExport(factor, mode, original) Then Exit Function
    RPX_AccuracyDualProblem original, dp
    seconds = CDbl(RPX_Setting("ACCURACY_DLL_DUAL_SECONDS", 20))
    If Not RPX_AccuracySolve(dp, seconds, candidate, helperDual) Then Exit Function
    Erase helperDual ' The auxiliary problem's dual is not an original program dual.
    accepted = RPX_AccuracyDualAudit(original, candidate)
    If accepted Then
        guideReady = True: guideFactor = factor: guideMode = mode: guideModel = original.model: guidePolicy = original.policy
        guideEpoch = original.epoch: guideQ = original.q: guidePhase = original.phase
        guideN = XN: guideM = XM: guideS = XS: guideB = XB: guideSolve = P4Solve: guideCandidate = RPX_HSDCandidateId
        guideCost = original.originalCost: guideNormalizedCost = XQ
        CaptureAccuracyProof original, candidate, True
    End If
    RPX_DiagEvent "accuracy_dll_original_dual", "accepted=" & accepted & ";mode=" & mode & ";Fs=" & factor & ";q=8;raw_stat=" & AccuracyNativeDualStat & ";raw_cone=" & AccuracyNativeDualCone & ";original_objective=" & AccuracyNativeDualObjective & ";physical_bound=0"
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
