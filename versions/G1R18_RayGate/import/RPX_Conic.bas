Attribute VB_Name = "RPX_Conic"
Option Explicit
#If Win64 Then
Private Declare PtrSafe Function NCInit Lib "rpfem_native_x64.dll" Alias "rpx_compile_create" (ByVal nx As Long, ByVal nm As Long, ByVal nb As Long, ByVal coneN As Long, ByVal equalityN As Long, ByVal args As LongPtr, ByRef handle As LongPtr) As Long
Private Declare PtrSafe Function NCDestroy Lib "rpfem_native_x64.dll" Alias "rpx_compile_destroy" (ByVal handle As LongPtr) As Long
Private Declare PtrSafe Function NCStep Lib "rpfem_native_x64.dll" Alias "rpx_compile_step" (ByVal handle As LongPtr, ByVal milliseconds As Long, ByRef done As Long) As Long
Private Declare PtrSafe Function NCSignature Lib "rpfem_native_x64.dll" Alias "rpx_compile_signature" (ByVal handle As LongPtr, ByVal output As LongPtr, ByVal capacity As Long, ByRef written As Long) As Long
Private Declare PtrSafe Function NCOrder Lib "rpfem_native_x64.dll" Alias "rpx_compile_order_begin" (ByVal handle As LongPtr, ByVal permutation As LongPtr, ByVal inverse As LongPtr, ByVal hit As Long) As Long
Private Declare PtrSafe Function NCPermute Lib "rpfem_native_x64.dll" Alias "rpx_compile_permute_begin" (ByVal handle As LongPtr) As Long
Private Declare PtrSafe Function NCCounts Lib "rpfem_native_x64.dll" Alias "rpx_compile_counts" (ByVal handle As LongPtr, ByVal output As LongPtr) As Long
Private Declare PtrSafe Function NCExport Lib "rpfem_native_x64.dll" Alias "rpx_compile_export" (ByVal handle As LongPtr, ByVal nz As Long, ByVal args As LongPtr) As Long
Private Declare PtrSafe Function NCSlots Lib "rpfem_native_x64.dll" Alias "rpx_compile_slots" (ByVal handle As LongPtr, ByVal cone As Long, ByVal count As Long, ByVal output As LongPtr) As Long
#End If
Public OriginalEqualityError As Double, OriginalConeError As Double, OriginalCandidateObjective As Double
Public Type RPX_OriginalCompiledData
    ready As Boolean
    n As Long
    m As Long
    s As Long
    b As Long
    modeName As String
    subdivisions As Long
    inputEpoch As Long
    affine As Boolean
    modelKey As String
    loadPhase As String
    values() As Double
    rhs() As Double
    cost() As Double
End Type
Public XOriginalCompiled As RPX_OriginalCompiledData
Public OriginalFarkasStationarity As Double, OriginalFarkasCone As Double, OriginalFarkasSeparation As Double
Public OriginalFarkasReason As String
Private originalGuideReady As Boolean, originalGuideFs As Double, originalGuideQ As Long
Private originalGuideEpoch As Long, originalGuideSolve As Long, originalGuideId As Long
Private originalGuideModel As String, originalGuidePhase As String

' version=20260926_0545 engine=20260926_order_slots
Public Type RPX_ConeData
    dimn As Long
    count As Long
    first As Long
    Pattern As Long
    modelKind As Long
    owner As Long
    parameter1 As Long
    parameter2 As Long
    directionSign As Long
    ids() As Long
    coef() As Double
    offset() As Double
    slots() As Long
End Type
Public XC() As RPX_ConeData, XN As Long, XM As Long, XS As Long, XB As Long
Public XEqRowNorm() As Double
Public XQ() As Double, XQScale As Double, XEPtr() As Long, XEId() As Long, XEVal() As Double, XERhs() As Double
Public XKPtr() As Long, XKId() As Long, XKVal() As Double, XKBase() As Double, XKDiag() As Long
Public XPerm() As Long, XInv() As Long, XKScale() As Double
Public xx() As Double, xy() As Double, XSlack() As Double, XDual() As Double
Public XObjective As Double, XResidual As Double, XGap As Double, XIteration As Long
Public XPrimalResidual As Double, XDualResidual As Double
Public XCancel As Boolean, XLogPath As String
Public RProgramKind As String, RProgramFs As Double
Private rawCones As Collection, rawEq As Collection, rawRhs As Collection, objective As Object
Private coneCapacity As Long
Private rowPool() As Object
Private rowPoolN As Long
Private rowPoolReady As Boolean
Private stageBarNoted As Boolean
Private stageInterval As Double
Private stageIntervalKnown As Boolean
Private stageHotClock As Double
Private stageHotReady As Boolean
Private contextGuardClock As Double
Private contextGuardReady As Boolean
Public RPX_OrderMode As String
Private cacheUpperValid As Boolean, cacheLowerValid As Boolean
Private cacheUpperN As Long, cacheLowerN As Long
Private cacheUpperKey As String, cacheLowerKey As String
Private cacheUpperPerm() As Long, cacheLowerPerm() As Long
Private cacheUpperInv() As Long, cacheLowerInv() As Long

Private Sub CopyDictKeys(ByVal d As Object, ByRef dest() As Long, ByRef nKeys As Long)
    Dim raw As Variant
    Dim i As Long
    nKeys = d.count
    If nKeys = 0 Then Exit Sub
    raw = d.keys
    ReDim dest(0 To nKeys - 1)
    For i = 0 To nKeys - 1
        dest(i) = CLng(raw(i))
    Next i
End Sub

Public Sub RPX_ReleaseRow(ByRef row As Object)
    If row Is Nothing Then Exit Sub
    On Error Resume Next
    row.RemoveAll
    On Error GoTo 0
    If Not rowPoolReady Then
        ReDim rowPool(0 To 255)
        rowPoolReady = True
    ElseIf rowPoolN > UBound(rowPool) Then
        ReDim Preserve rowPool(0 To UBound(rowPool) * 2 + 1)
    End If
    Set rowPool(rowPoolN) = row
    rowPoolN = rowPoolN + 1
    Set row = Nothing
End Sub

Public Sub RPX_StageConfigure()
    Dim value As Double
    stageHotReady = False
    value = CDbl(RPX_Setting("HSD_STAGE_SECONDS", 1#))
    If value < 0# Or value > 60# Then err.Raise 5, , "HSD_STAGE_SECONDS_INVALID"
    stageInterval = value
    stageIntervalKnown = True
    contextGuardReady = False
    RPX_DiagEvent "stage_interval", "seconds=" & CStr(value) & ";hot_prefix=HSD sparse ;context_guard=hsd_linear"
End Sub
Public Function RPX_StageIntervalValue() As Double
    RPX_StageIntervalValue = 1#
    If stageIntervalKnown Then RPX_StageIntervalValue = stageInterval
End Function
Public Function RPX_ContextGuardDue() As Boolean
    Dim nowClock As Double, elapsed As Double
    ' HContext sheet read and model string only. Same seconds as hot progress.
    ' Zero keeps every call. Cancel, epoch, and the iterate bits are not gated here.
    If Not stageIntervalKnown Then
        stageInterval = 1#
        stageIntervalKnown = True
    End If
    If stageInterval <= 0# Then
        RPX_ContextGuardDue = True
        Exit Function
    End If
    nowClock = Timer
    If contextGuardReady Then
        elapsed = nowClock - contextGuardClock
        If elapsed < 0# Then elapsed = elapsed + 86400#
        If elapsed < stageInterval Then Exit Function
    End If
    contextGuardClock = nowClock
    contextGuardReady = True
    RPX_ContextGuardDue = True
End Function
Public Sub RPX_Stage(ByVal message As String)
    Dim f As Integer
    Dim mark As Long, failNumber As Long, failSource As String, failText As String
    Dim barText As String, barFailed As Long
    Dim nowClock As Double, elapsed As Double
    ' Hot LU progress only. Numeric updates are unchanged. Non-sparse stages always run.
    ' HSD_STAGE_SECONDS=0 keeps the old call cadence. Default 1 second.
    If Not stageIntervalKnown Then
        stageInterval = 1#
        stageIntervalKnown = True
    End If
    If stageInterval > 0# Then
        If Left$(message, 11) = "HSD sparse " Then
            nowClock = Timer
            If stageHotReady Then
                elapsed = nowClock - stageHotClock
                If elapsed < 0# Then elapsed = elapsed + 86400#
                If elapsed < stageInterval Then Exit Sub
            End If
            stageHotClock = nowClock
            stageHotReady = True
        End If
    End If
    mark = RPX_HoptMark()
    On Error GoTo HoptFailed
    RPX_HoptEnter HoptStageBar
    ' Display only. A 1004 here is Excel UI state (cell edit, menu) after DoEvents,
    ' not a factorization failure. The log, yield, and input check still run.
    barText = "RPFEM: " & message
    If Len(barText) > 255 Then barText = Left$(barText, 255)
    On Error Resume Next
    Application.StatusBar = barText
    barFailed = err.number
    err.Clear
    On Error GoTo HoptFailed
    If barFailed <> 0 And Not stageBarNoted Then
        stageBarNoted = True
        RPX_DiagEvent "stage_bar_skipped", "number=" & barFailed
    End If
    RPX_HoptLeave HoptStageBar
    RPX_HoptEnter HoptStageFile
    If Len(XLogPath) > 0 Then
        f = FreeFile: Open XLogPath For Append As #f: Print #f, "# " & message & " " & CStr(Timer): Close #f
    End If
    RPX_HoptLeave HoptStageFile
    RPX_HoptEnter HoptStageLog
    RPX_DiagEvent "stage", message
    RPX_HoptLeave HoptStageLog
    RPX_HoptEnter HoptStageYield
    DoEvents
    RPX_HoptLeave HoptStageYield
    RPX_HoptEnter HoptInputAssert
    RPX_AssertInputs "stage"
    RPX_HoptLeave HoptInputAssert
    If XCancel Then err.Raise 18, , "ANALYSIS_CANCELLED"
    Exit Sub
HoptFailed:
    failNumber = err.number: failSource = err.source: failText = err.description
    RPX_HoptUnwind mark
    err.Raise failNumber, failSource, failText
End Sub

Public Function RPX_Row() As Object
    If rowPoolN > 0 Then
        rowPoolN = rowPoolN - 1
        Set RPX_Row = rowPool(rowPoolN)
        Set rowPool(rowPoolN) = Nothing
        RPX_Row.RemoveAll
    Else
        Set RPX_Row = CreateObject("Scripting.Dictionary")
    End If
End Function
Public Sub RPX_Add(ByVal row As Object, ByVal id As Long, ByVal value As Double)
    If value = 0# Then Exit Sub
    If row.Exists(id) Then value = value + row(id)
    If Abs(value) < 0.00000000000001 Then
        If row.Exists(id) Then row.Remove id
    Else
        row(id) = value
    End If
End Sub
Public Function RPX_Unit(ByVal id As Long, Optional ByVal value As Double = 1#) As Object
    Set RPX_Unit = RPX_Row(): RPX_Add RPX_Unit, id, value
End Function
Public Sub RPX_Axpy(ByVal target As Object, ByVal source As Object, ByVal value As Double)
    Dim srcKeys() As Long
    Dim nKeys As Long
    Dim i As Long
    Dim id As Long
    CopyDictKeys source, srcKeys, nKeys
    For i = 0 To nKeys - 1
        id = srcKeys(i)
        RPX_Add target, id, value * source(id)
    Next i
End Sub
Public Sub RPX_Begin(ByVal variables As Long)
    RPX_ClearOriginalCompiled
    RPX_NativeBulkReset
    RProgramKind = ""
    RPX_ReleaseRow objective
    XN = variables: Set rawCones = New Collection: Set rawEq = New Collection
    Set rawRhs = New Collection: Set objective = RPX_Row(): XCancel = False
    XB = 0: XS = 0: coneCapacity = 1024: ReDim XC(0 To coneCapacity - 1)
End Sub
Public Function RPX_Variable(Optional ByVal cost As Double = 0#) As Long
    RPX_Variable = XN: XN = XN + 1
    RPX_Add objective, XN - 1, cost
End Function
Public Sub RPX_Cost(ByVal row As Object, Optional ByVal factor As Double = 1#)
    RPX_Axpy objective, row, factor
End Sub
Public Sub RPX_Equal(ByVal row As Object, Optional ByVal rhs As Double = 0#)
    If row.count = 0 Then
        If rhs <> 0# Then err.Raise 5, , "INCONSISTENT_EMPTY_EQUALITY"
        Exit Sub
    End If
    rawEq.Add row: rawRhs.Add rhs
End Sub
Public Sub RPX_Cone(ByVal rows As Collection, ByVal offset As Variant)
    Dim row As Object, ids As Object, key As Variant, keys As Variant, j As Long, r As Long
    If rows.count <> 1 And rows.count <> 3 Then err.Raise 5, , "CONE_DIMENSION_MUST_BE_1_OR_3"
    If XB = coneCapacity Then coneCapacity = coneCapacity * 2: ReDim Preserve XC(0 To coneCapacity - 1)
    Set ids = RPX_Row()
    For Each row In rows
        For Each key In row.keys: ids(CLng(key)) = True: Next key
    Next row
    keys = ids.keys
    With XC(XB)
        .dimn = rows.count: .count = ids.count: .first = XS: XS = XS + .dimn
        If .count = 0 Then err.Raise 5, , "EMPTY_CONE_NOT_SUPPORTED"
        ReDim .ids(0 To .count - 1): ReDim .coef(0 To .dimn - 1, 0 To .count - 1): ReDim .offset(0 To .dimn - 1)
        ReDim .slots(0 To .count - 1, 0 To .count - 1)
        For j = 0 To .count - 1: .ids(j) = CLng(keys(j)): Next j
        For r = 0 To .dimn - 1
            .offset(r) = offset(r): Set row = rows(r + 1)
            For j = 0 To .count - 1
                If row.Exists(.ids(j)) Then .coef(r, j) = row(.ids(j))
            Next j
        Next r
    End With
    XB = XB + 1
End Sub
Public Sub RPX_Cone3(ByVal a As Object, ByVal b As Object, ByVal c As Object, Optional ByVal offset0 As Double = 0#)
    Dim rows As New Collection
    rows.Add a: rows.Add b: rows.Add c: RPX_Cone rows, Array(offset0, 0#, 0#)
End Sub
Public Sub RPX_LinearCone(ByVal row As Object)
    Dim rows As New Collection
    rows.Add row: RPX_Cone rows, Array(0#)
End Sub
Public Sub RPX_LinearDirect(ByRef ids() As Long, ByRef coefficients() As Double, ByVal count As Long)
    Dim j As Long, k As Long, used As Long
    If XB = coneCapacity Then coneCapacity = coneCapacity * 2: ReDim Preserve XC(0 To coneCapacity - 1)
    used = count
    If used = 0 Then err.Raise 5, , "EMPTY_DIRECT_CONE"
    With XC(XB)
        .dimn = 1: .count = used: .first = XS: XS = XS + 1
        ReDim .ids(0 To used - 1): ReDim .coef(0 To 0, 0 To used - 1): ReDim .offset(0 To 0): ReDim .slots(0 To used - 1, 0 To used - 1)
        For j = 0 To count - 1: .ids(j) = ids(j): .coef(0, j) = coefficients(j): Next j
    End With
    XB = XB + 1
End Sub

Private Sub SortLongs(ByRef values() As Long, ByVal n As Long)
    Dim i As Long, j As Long, tmp As Long
    For i = 1 To n - 1
        tmp = values(i): j = i - 1
        Do While j >= 0
            If values(j) <= tmp Then Exit Do
            values(j + 1) = values(j)
            j = j - 1
        Loop
        values(j + 1) = tmp
    Next i
End Sub
Private Function KktPatternSignature(ByVal nv As Long, ByRef columns() As Object) As String
    Dim i As Long, ki As Long, nKeys As Long, keyIds() As Long
    Dim parts() As String
    ReDim parts(0 To nv - 1)
    For i = 0 To nv - 1
        CopyDictKeys columns(i), keyIds, nKeys
        If nKeys > 1 Then SortLongs keyIds, nKeys
        If nKeys <= 0 Then
            parts(i) = "-"
        Else
            parts(i) = CStr(keyIds(0))
            For ki = 1 To nKeys - 1
                parts(i) = parts(i) & "," & CStr(keyIds(ki))
            Next ki
        End If
    Next i
    KktPatternSignature = CStr(nv) & ":" & Join(parts, ";")
End Function
Private Sub KeepSlot(ByRef valid As Boolean, ByRef storedN As Long, ByRef storedKey As String, ByRef perm() As Long, ByRef inv() As Long, ByVal nv As Long, ByVal signature As String)
    Dim i As Long
    ReDim perm(0 To nv - 1): ReDim inv(0 To nv - 1)
    For i = 0 To nv - 1
        perm(i) = XPerm(i): inv(i) = XInv(i)
    Next i
    storedN = nv: storedKey = signature: valid = True
End Sub
Private Function LoadSlot(ByVal valid As Boolean, ByVal storedN As Long, ByVal storedKey As String, ByRef perm() As Long, ByRef inv() As Long, ByVal nv As Long, ByVal signature As String, ByRef reason As String) As Boolean
    Dim i As Long, ready As Boolean
    If Not valid Then reason = "empty": Exit Function
    If storedN <> nv Then reason = "n": Exit Function
    If Len(storedKey) = 0 Or storedKey <> signature Then reason = "pattern": Exit Function
    ready = True
    On Error GoTo MissingSlot
    If UBound(perm) <> nv - 1 Then ready = False
    If UBound(inv) <> nv - 1 Then ready = False
    On Error GoTo 0
    If Not ready Then reason = "arrays": Exit Function
    ReDim XPerm(0 To nv - 1): ReDim XInv(0 To nv - 1)
    For i = 0 To nv - 1
        XPerm(i) = perm(i): XInv(i) = inv(i)
    Next i
    reason = ""
    LoadSlot = True
    Exit Function
MissingSlot:
    reason = "arrays"
    LoadSlot = False
End Function
Private Sub KeepOrdering(ByVal mode As String, ByVal nv As Long, ByVal signature As String)
    If mode = "upper" Then
        KeepSlot cacheUpperValid, cacheUpperN, cacheUpperKey, cacheUpperPerm, cacheUpperInv, nv, signature
    ElseIf mode = "lower" Then
        KeepSlot cacheLowerValid, cacheLowerN, cacheLowerKey, cacheLowerPerm, cacheLowerInv, nv, signature
    End If
End Sub
Private Function LoadOrdering(ByVal mode As String, ByVal nv As Long, ByVal signature As String, ByRef reason As String) As Boolean
    reason = "mode"
    If mode = "upper" Then
        LoadOrdering = LoadSlot(cacheUpperValid, cacheUpperN, cacheUpperKey, cacheUpperPerm, cacheUpperInv, nv, signature, reason)
    ElseIf mode = "lower" Then
        LoadOrdering = LoadSlot(cacheLowerValid, cacheLowerN, cacheLowerKey, cacheLowerPerm, cacheLowerInv, nv, signature, reason)
    End If
End Function
Public Sub RPX_IndexRows(ByVal rows As Collection, ByVal rhs As Collection, ByRef indexedRows() As Object, ByRef indexedRhs() As Double)
    Dim item As Variant, i As Long, n As Long
    n = rows.count
    If rhs.count <> n Then err.Raise 9, "RPX_Conic", "ROW_RHS_COUNT_MISMATCH"
    ReDim indexedRows(0 To n): ReDim indexedRhs(0 To n)
    i = 1
    For Each item In rows: Set indexedRows(i) = item: i = i + 1: Next item
    i = 1
    For Each item In rhs: indexedRhs(i) = CDbl(item): i = i + 1: Next item
End Sub
Public Sub RPX_Compile()
    RPX_NativeBulkReset
    Dim indexedRows() As Object, indexedRhs() As Double
    Dim i As Long, j As Long, k As Long, r As Long, p As Long, cnt As Long, id As Long, col As Long
    Dim keys As Variant, key As Variant, row As Object, ids As Object, rows As Collection, rec As Collection, off As Variant
    Dim columns() As Object, permCols() As Object, valueScale As Double, original As Long, other As Long, nv As Long
    Dim patterns As Object, signature As String
    Dim keyIds() As Long, nKeys As Long, ki As Long
    Dim orderReason As String, orderHit As Long, largest As Double, ratio As Double
    RPX_DiagBegin dpConvert
    RPX_Stage "compile arrays": XM = rawEq.count
    RPX_IndexRows rawEq, rawRhs, indexedRows, indexedRhs
    If XB = 0 Then err.Raise 5, , "?L?o?E?~??????n?a???e?U?1?n?B???b?V???E?T???E????E?d???d?m?F????A?-???3????B"
    ReDim Preserve XC(0 To XB - 1): ReDim XQ(0 To XN - 1)
    XQScale = 1#
    CopyDictKeys objective, keyIds, nKeys
    For ki = 0 To nKeys - 1
        XQ(keyIds(ki)) = objective(keyIds(ki))
        If Abs(objective(keyIds(ki))) > XQScale Then XQScale = Abs(objective(keyIds(ki)))
    Next ki
    RPX_ClearOriginalCompiled
    XOriginalCompiled.n = XN: XOriginalCompiled.m = XM: XOriginalCompiled.s = XS: XOriginalCompiled.b = XB
    XOriginalCompiled.cost = XQ
    XOriginalCompiled.modeName = RPX_OrderMode: XOriginalCompiled.subdivisions = RSubdivisions
    XOriginalCompiled.inputEpoch = RPX_InputEpoch: XOriginalCompiled.affine = AFActive
    XOriginalCompiled.loadPhase = RPX_LoadPhase
    If rn > 0 Then XOriginalCompiled.modelKey = RPX_ExactModel()
    For i = 0 To XN - 1: XQ(i) = XQ(i) / XQScale: Next i
    ReDim XEPtr(0 To XM): ReDim XERhs(0 To XM): ReDim XEqRowNorm(0 To XM)
    For i = 0 To XM - 1: XEPtr(i + 1) = XEPtr(i) + indexedRows(i + 1).count: Next i
    ReDim XEId(0 To XEPtr(XM)): ReDim XEVal(0 To XEPtr(XM))
    ReDim XOriginalCompiled.values(0 To XEPtr(XM)): ReDim XOriginalCompiled.rhs(0 To XM)
    For i = 0 To XM - 1
        Set row = indexedRows(i + 1): valueScale = 0#
        XOriginalCompiled.rhs(i) = indexedRhs(i + 1)
        CopyDictKeys row, keyIds, nKeys
        largest = 0#
        For ki = 0 To nKeys - 1
            If Abs(row(keyIds(ki))) > largest Then largest = Abs(row(keyIds(ki)))
        Next ki
        If largest = 0# Then err.Raise 5, "RPX_Conic", "ZERO_EQUALITY_ROW"
        For ki = 0 To nKeys - 1
            ratio = row(keyIds(ki)) / largest: valueScale = valueScale + ratio * ratio
        Next ki
        valueScale = largest * Sqr(valueScale): XEqRowNorm(i) = valueScale: p = XEPtr(i): XERhs(i) = indexedRhs(i + 1) / valueScale
        For ki = 0 To nKeys - 1
            XOriginalCompiled.values(p) = CDbl(row(keyIds(ki)))
            XEId(p) = keyIds(ki): XEVal(p) = row(keyIds(ki)) / valueScale: p = p + 1
        Next ki
    Next i
    nv = XN + XM
    RPX_CapacityCheck "before_kkt_pattern", 2400# * nv + 32# * XEPtr(XM) + 400# * XB
    If NativeCompilePattern(nv) Then GoTo CompiledPattern
    ReDim columns(0 To nv - 1)
    Set patterns = RPX_Row()
    For i = 0 To nv - 1: Set columns(i) = RPX_Row(): columns(i)(i) = True: Next i
    For i = 0 To XB - 1
        With XC(i)
            signature = ""
            For j = 0 To .count - 1: signature = signature & CStr(.ids(j)) & ",": Next j
            If patterns.Exists(signature) Then
                .Pattern = patterns(signature): Erase .slots
            Else
                .Pattern = i: patterns(signature) = i
                For j = 0 To .count - 1
                    For k = 0 To j
                        id = .ids(k): col = .ids(j)
                        If id > col Then original = id: id = col: col = original
                        columns(col)(id) = True
                    Next k
                Next j
            End If
        End With
    Next i
    For i = 0 To XM - 1
        For p = XEPtr(i) To XEPtr(i + 1) - 1: columns(XN + i)(XEId(p)) = True: Next p
    Next i
    RPX_DiagEnd dpConvert: RPX_DiagBegin dpOrdering
    signature = KktPatternSignature(nv, columns)
    If LoadOrdering(RPX_OrderMode, nv, signature, orderReason) Then
        orderHit = 1
    Else
        RPX_Stage "ordering " & nv: RPX_Ordering nv, columns, XPerm, XInv
        KeepOrdering RPX_OrderMode, nv, signature
        orderHit = 0
    End If
    RPX_DiagEvent "order_cache_lookup", "engine=20260926_order_slots;ver=20260926_0545;mode=" & RPX_OrderMode & ";hit=" & orderHit & ";miss_reason=" & orderReason & ";nv=" & nv & ";key_chars=" & Len(signature) & ";cache_slot=" & RPX_OrderMode
    RPX_DiagEnd dpOrdering: RPX_DiagBegin dpConvert
    RPX_Stage "ordering complete"
    ReDim permCols(0 To nv - 1)
    For i = 0 To nv - 1: Set permCols(i) = RPX_Row(): Next i
    For i = 0 To nv - 1
        CopyDictKeys columns(i), keyIds, nKeys
        For ki = 0 To nKeys - 1
            col = XInv(i): id = XInv(keyIds(ki))
            If id > col Then original = id: id = col: col = original
            permCols(col)(id) = True
        Next ki
        RPX_ReleaseRow columns(i)
    Next i
    ReDim XKPtr(0 To nv): ReDim XKDiag(0 To nv - 1): ReDim XKScale(0 To nv - 1)
    For i = 0 To nv - 1: XKPtr(i + 1) = XKPtr(i) + permCols(i).count: Next i
    ReDim XKId(0 To XKPtr(nv) - 1): ReDim XKVal(0 To XKPtr(nv) - 1): ReDim XKBase(0 To XKPtr(nv) - 1)
    For i = 0 To nv - 1
        p = XKPtr(i)
        CopyDictKeys permCols(i), keyIds, nKeys
        For ki = 0 To nKeys - 1
            XKId(p) = keyIds(ki): permCols(i)(keyIds(ki)) = p
            If keyIds(ki) = i Then XKDiag(i) = p
            p = p + 1
        Next ki
    Next i
    For i = 0 To XB - 1
        With XC(i)
            If .Pattern = i Then
                For j = 0 To .count - 1
                    For k = 0 To j
                        col = XInv(.ids(j)): id = XInv(.ids(k))
                        If id > col Then original = id: id = col: col = original
                        .slots(j, k) = permCols(col)(id)
                    Next k
                Next j
            End If
        End With
    Next i
    For i = 0 To XM - 1
        For p = XEPtr(i) To XEPtr(i + 1) - 1
            col = XInv(XN + i): id = XInv(XEId(p))
            If id > col Then original = id: id = col: col = original
            XKBase(permCols(col)(id)) = XKBase(permCols(col)(id)) + XEVal(p)
        Next p
    Next i
CompiledPattern:
    RPX_DiagEnd dpConvert: RPX_DiagBegin dpSymbolic
    RPX_Stage "symbolic factorization": RPX_Symbolic nv, XKPtr, XKId
    RPX_DiagEnd dpSymbolic
    Call RPX_DiagCounts
    RPX_Stage "symbolic complete " & RPX_FactorNNZ
    If Not patterns Is Nothing Then
        For i = 0 To nv - 1: RPX_ReleaseRow permCols(i): Next i
        RPX_ReleaseRow patterns
    End If
    XOriginalCompiled.ready = True
    Set rawEq = Nothing: Set rawCones = Nothing: Set rawRhs = Nothing
End Sub

Private Function NativeCompilePattern(ByVal nv As Long) As Boolean
    If Not RPX_NativeAvailable() Then Exit Function
#If Win64 Then
    Dim handle As LongPtr, args(0 To 4) As LongPtr, outputArgs(0 To 6) As LongPtr
    Dim coneN As Long, done As Long, counts(0 To 1) As Long, patterns() As Long
    Dim signature As String, written As Long, hit As Long, reason As String, i As Long, milliseconds As Long
    Dim eno As Long, es As String, et As String
    On Error GoTo Failed
    RPX_NativeConeIdentityArgs args, coneN
    args(2) = VarPtr(XEPtr(0)): args(3) = VarPtr(XEId(0)): args(4) = VarPtr(XEVal(0))
    RPX_NativeCheckCode NCInit(XN, XM, XB, coneN, XEPtr(XM), VarPtr(args(0)), handle)
    milliseconds = 50: If RPX_StageIntervalValue() = 0# Then milliseconds = 0
    Do While done <> 1
        RPX_NativePulse
        RPX_NativeCheckCode NCStep(handle, milliseconds, done)
    Loop
    RPX_NativeCheckCode NCSignature(handle, 0, 0, written)
    signature = String$(written, vbNullChar)
    RPX_NativeCheckCode NCSignature(handle, StrPtr(signature), Len(signature), written)
    RPX_DiagEnd dpConvert: RPX_DiagBegin dpOrdering
    If LoadOrdering(RPX_OrderMode, nv, signature, reason) Then
        hit = 1
    Else
        RPX_Stage "ordering " & nv
        ReDim XPerm(0 To nv - 1): ReDim XInv(0 To nv - 1)
    End If
    RPX_NativeCheckCode NCOrder(handle, VarPtr(XPerm(0)), VarPtr(XInv(0)), hit)
    done = 0
    Do While done <> 3
        RPX_NativePulse
        RPX_NativeCheckCode NCStep(handle, milliseconds, done)
    Loop
    RPX_DiagEnd dpOrdering: RPX_DiagBegin dpConvert
    RPX_Stage "ordering complete"
    RPX_NativeCheckCode NCPermute(handle)
    done = 0
    Do While done <> 2
        RPX_NativePulse
        RPX_NativeCheckCode NCStep(handle, milliseconds, done)
    Loop
    RPX_NativeCheckCode NCCounts(handle, VarPtr(counts(0)))
    ReDim XKPtr(0 To nv): ReDim XKId(0 To counts(0) - 1): ReDim XKDiag(0 To nv - 1)
    ReDim XKVal(0 To counts(0) - 1): ReDim XKBase(0 To counts(0) - 1): ReDim XKScale(0 To nv - 1): ReDim patterns(0 To XB - 1)
    outputArgs(0) = VarPtr(XPerm(0)): outputArgs(1) = VarPtr(XInv(0)): outputArgs(2) = VarPtr(XKPtr(0))
    outputArgs(3) = VarPtr(XKId(0)): outputArgs(4) = VarPtr(XKDiag(0)): outputArgs(5) = VarPtr(XKBase(0)): outputArgs(6) = VarPtr(patterns(0))
    RPX_NativeCheckCode NCExport(handle, counts(0), VarPtr(outputArgs(0)))
    For i = 0 To XB - 1
        If i Mod 256 = 0 Then RPX_NativePulse
        With XC(i)
            .Pattern = patterns(i)
            If .Pattern <> i Then
                Erase .slots
            ElseIf .count > 0 Then
                If LBound(.slots, 1) <> 0 Or UBound(.slots, 1) <> .count - 1 Or LBound(.slots, 2) <> 0 Or UBound(.slots, 2) <> .count - 1 Then err.Raise 9, , "NATIVE_COMPILE_SLOT_SHAPE"
                RPX_NativeCheckCode NCSlots(handle, i, .count * .count, VarPtr(.slots(0, 0)))
            End If
        End With
    Next i
    If hit = 0 Then KeepOrdering RPX_OrderMode, nv, signature
    RPX_DiagEvent "order_cache_lookup", "engine=20260926_order_slots;ver=20260926_0545;mode=" & RPX_OrderMode & ";hit=" & hit & ";miss_reason=" & reason & ";nv=" & nv & ";key_chars=" & Len(signature) & ";cache_slot=" & RPX_OrderMode
    RPX_NativeCheckCode NCDestroy(handle): handle = 0
    NativeCompilePattern = True
    Exit Function
Failed:
    eno = err.number: es = err.source: et = err.description
    If handle <> 0 Then RPX_NativeCheckCode NCDestroy(handle)
    err.Raise eno, es, et
#End If
End Function

Public Sub RPX_GProduct(ByRef x() As Double, ByRef answer() As Double, Optional ByVal withOffset As Boolean = False)
    Dim i As Long, j As Long, r As Long, value As Double
    RPX_DiagBegin dpGProduct
    ReDim answer(0 To XS - 1)
    If RPX_NativeGProduct(x, answer, withOffset) Then RPX_DiagEnd dpGProduct: Exit Sub
    For i = 0 To XB - 1
        With XC(i)
            For r = 0 To .dimn - 1
                value = 0#: If withOffset Then value = .offset(r)
                For j = 0 To .count - 1: value = value + .coef(r, j) * x(.ids(j)): Next j
                answer(.first + r) = value
            Next r
        End With
    Next i
    RPX_DiagEnd dpGProduct
End Sub
Public Sub RPX_GTAdd(ByRef x() As Double, ByRef answer() As Double, ByVal factor As Double)
    Dim i As Long, j As Long, r As Long, value As Double
    RPX_DiagBegin dpGTAdd
    If RPX_NativeGTAdd(x, answer, factor) Then RPX_DiagEnd dpGTAdd: Exit Sub
    For i = 0 To XB - 1
        With XC(i)
            For j = 0 To .count - 1
                value = 0#
                For r = 0 To .dimn - 1: value = value + .coef(r, j) * x(.first + r): Next r
                answer(.ids(j)) = answer(.ids(j)) + factor * value
            Next j
        End With
    Next i
    RPX_DiagEnd dpGTAdd
End Sub

Public Sub RPX_ReadConic(ByVal path As String)
    Dim f As Integer, n As Long, m As Long, nc As Long, i As Long, j As Long, k As Long, d As Long, cnt As Long, id As Long
    Dim value As Double, rhs As Double, row As Object, rows As Collection, offset() As Double
    Dim number As Long, source As String, message As String
    On Error GoTo Failed
    f = FreeFile: Open path For Input As #f
    Input #f, n, m, nc
    If n < 1 Or m < 0 Or nc < 1 Then err.Raise 5, "RPX_ReadConic", "CONIC_DIMENSIONS_INVALID"
    RPX_Begin n
    For i = 0 To n - 1
        Input #f, value: If value <> 0# Then RPX_Cost RPX_Unit(i, value)
    Next i
    For i = 1 To m
        Input #f, cnt, rhs: Set row = RPX_Row()
        For j = 1 To cnt
            Input #f, id, value
            If id < 0 Or id >= n Then err.Raise 9, "RPX_ReadConic", "CONIC_COLUMN_INVALID"
            RPX_AddExact row, id, value
        Next j
        If row.count = 0 And rhs <> 0# Then err.Raise 5, "RPX_ReadConic", "EMPTY_READ_ROW i=" & i & ";cnt=" & cnt & ";rhs=" & rhs & ";last_id=" & id & ";last_value=" & value
        RPX_Equal row, rhs
    Next i
    For i = 1 To nc
        Input #f, d: ReDim offset(0 To d - 1): Set rows = New Collection
        For k = 0 To d - 1
            Input #f, cnt, offset(k): Set row = RPX_Row()
            For j = 1 To cnt
            Input #f, id, value
            If id < 0 Or id >= n Then err.Raise 9, "RPX_ReadConic", "CONIC_COLUMN_INVALID"
            RPX_AddExact row, id, value
        Next j
            rows.Add row
        Next k
        RPX_Cone rows, offset
    Next i
    Close #f: Exit Sub
Failed:
    number = err.number: source = err.source: message = err.description
    On Error Resume Next: Close #f: On Error GoTo 0
    err.Raise number, source, message
End Sub




Public Sub RPX_ExportAssembled(ByVal path As String)
    Dim indexedRows() As Object, indexedRhs() As Double
    Dim f As Integer, i As Long, j As Long, k As Long, r As Long, count As Long, value As Double, id As Long
    Dim number As Long, source As String, message As String
    Dim row As Object, key As Variant
    On Error GoTo Failed
    If Len(Dir$(path)) > 0 Then err.Raise 5, "RPX_Exchange", "OUTPUT_EXISTS"
    f = FreeFile: Open path For Binary Access Write As #f
    RPX_IndexRows rawEq, rawRhs, indexedRows, indexedRhs
    RPX_ExchangeMagic f, "P06VBA01": Put #f, , XN: count = rawEq.count: Put #f, , count: Put #f, , XB
    For i = 0 To XN - 1
        value = 0#: If objective.Exists(i) Then value = objective(i)
        Put #f, , value
    Next i
    For i = 1 To rawEq.count
        Set row = indexedRows(i): count = row.count: Put #f, , count
        value = indexedRhs(i): Put #f, , value
        For Each key In row.keys
            id = CLng(key): value = row(key): Put #f, , id: Put #f, , value
        Next key
    Next i
    For k = 0 To XB - 1
        count = XC(k).dimn: Put #f, , count
        For r = 0 To XC(k).dimn - 1
            count = 0
            For j = 0 To XC(k).count - 1
                If XC(k).coef(r, j) <> 0# Then count = count + 1
            Next j
            Put #f, , count: value = XC(k).offset(r): Put #f, , value
            For j = 0 To XC(k).count - 1
                If XC(k).coef(r, j) <> 0# Then
                    id = XC(k).ids(j): value = XC(k).coef(r, j): Put #f, , id: Put #f, , value
                End If
            Next j
        Next r
    Next k
    Close #f: Exit Sub
Failed:
    number = err.number: source = err.source: message = err.description
    On Error Resume Next: Close #f: On Error GoTo 0
    err.Raise number, source, message
End Sub

Public Sub RPX_CondenseFixedWitness()
    Dim key As Variant
    For Each key In objective.keys
        If CDbl(objective(key)) <> 0# Then err.Raise 5, "RPX_Conic", "AFFINE_REQUIRES_ZERO_OBJECTIVE"
    Next key
    RPX_AffinePrepare rawEq, rawRhs
End Sub
Public Sub RPX_FixLowerUnitLoad()
    If RLoadVariable < 0 Or RLoadVariable >= XN Then err.Raise 9, "RPX_Conic", "FIXED_LOAD_COLUMN_INVALID"
    Set objective = RPX_Row()
    RPX_Equal RPX_Unit(RLoadVariable), 1#
End Sub

Public Sub RPX_ReadMappedWitness(ByVal originalPath As String, ByVal reducedPath As String, ByVal mapPath As String)
    RPX_ReadConic originalPath
    RPX_AffineOriginalContext rawEq, rawRhs
    RPX_ReadConic reducedPath
    RPX_AffineLoadMap mapPath
End Sub

' Gate on the assembled ORIGINAL program, before Compile/normalization/permutation.
Private Sub AddCompensated(ByRef total As Double, ByRef correction As Double, ByVal value As Double)
    Dim adjusted As Double, nextValue As Double
    adjusted = value - correction: nextValue = total + adjusted
    correction = (nextValue - total) - adjusted: total = nextValue
End Sub
Public Function RPX_OriginalPrimalGate(ByRef field() As Double) As Boolean
    Dim indexedRows() As Object, indexedRhs() As Double
    Dim row As Object, key As Variant, i As Long, j As Long, k As Long, r As Long
    Dim total As Double, correction As Double, coneValue(0 To 2) As Double, errorValue As Double
    If UBound(field) <> XN - 1 Or LBound(field) <> 0 Then err.Raise 9, "RPX_Conic", "ORIGINAL_FIELD_DIMENSION"
    OriginalEqualityError = 0#: OriginalConeError = 0#: OriginalCandidateObjective = 0#
    RPX_IndexRows rawEq, rawRhs, indexedRows, indexedRhs
    For i = 0 To XN - 1
        If Not RPX_SchurFinite(field(i)) Then Exit Function
    Next i
    For i = 1 To rawEq.count
        total = -indexedRhs(i): correction = 0#: Set row = indexedRows(i)
        For Each key In row.keys: AddCompensated total, correction, CDbl(row(key)) * field(CLng(key)): Next key
        If Abs(total) > OriginalEqualityError Then OriginalEqualityError = Abs(total)
    Next i
    For k = 0 To XB - 1
        For r = 0 To XC(k).dimn - 1
            total = XC(k).offset(r): correction = 0#
            For j = 0 To XC(k).count - 1: AddCompensated total, correction, XC(k).coef(r, j) * field(XC(k).ids(j)): Next j
            coneValue(r) = total
        Next r
        errorValue = -coneValue(0)
        If XC(k).dimn = 3 Then errorValue = Sqr(coneValue(1) * coneValue(1) + coneValue(2) * coneValue(2)) - coneValue(0)
        If errorValue > OriginalConeError Then OriginalConeError = errorValue
    Next k
    total = 0#: correction = 0#
    For Each key In objective.keys: AddCompensated total, correction, CDbl(objective(key)) * field(CLng(key)): Next key
    OriginalCandidateObjective = total
    RPX_OriginalPrimalGate = OriginalEqualityError <= 0.0000001 And OriginalConeError <= 0.0000001
End Function
' Exact epigraph lifting for a fixed velocity on phi=0 upper programs.
' Retains every Bernstein coefficient; no KKT/numeric factorization.
Public Sub RPX_LiftUpperEpigraph(ByRef field() As Double)
    Dim base As Long, k As Long, j As Long, r As Long, auxiliary As Long, count As Long, coefficient As Double
    Dim values(0 To 2) As Double, required As Double, total As Double, correction As Double
    base = 12 * re + 3 * RR
    If UBound(field) <> XN - 1 Then err.Raise 9, "RPX_Conic", "LIFT_FIELD_DIMENSION"
    For k = 0 To XB - 1
        auxiliary = -1: count = 0: coefficient = 0#
        For j = 0 To XC(k).count - 1
            If XC(k).ids(j) >= base Then
                count = count + 1: auxiliary = XC(k).ids(j): coefficient = XC(k).coef(0, j)
                For r = 1 To XC(k).dimn - 1
                    If XC(k).coef(r, j) <> 0# Then err.Raise 5, "RPX_Conic", "LIFT_AUXILIARY_TAIL"
                Next r
            End If
        Next j
        If count > 1 Then err.Raise 5, "RPX_Conic", "LIFT_MULTIPLE_AUXILIARIES"
        If count = 1 Then
            If coefficient <= 0# Then err.Raise 5, "RPX_Conic", "LIFT_NONPOSITIVE_HEAD"
            For r = 0 To XC(k).dimn - 1
                total = XC(k).offset(r): correction = 0#
                For j = 0 To XC(k).count - 1
                    If XC(k).ids(j) <> auxiliary Then AddCompensated total, correction, XC(k).coef(r, j) * field(XC(k).ids(j))
                Next j
                values(r) = total
            Next r
            required = -values(0) / coefficient
            If XC(k).dimn = 3 Then required = (Sqr(values(1) * values(1) + values(2) * values(2)) - values(0)) / coefficient
            If required > field(auxiliary) Then field(auxiliary) = required
        End If
    Next k
End Sub

' Strict discrete guide for full, zero-cost physical upper programs only.
' The ray never supplies a physical velocity, a work witness, or a continuum bound.
Public Sub RPX_OriginalGuideClear()
    originalGuideReady = False: originalGuideModel = ""
End Sub
Public Sub RPX_ClearOriginalCompiled()
    Dim blank As RPX_OriginalCompiledData
    XOriginalCompiled = blank: RPX_OriginalGuideClear
End Sub
Public Function RPX_OriginalPhysicalZeroCostUpper() As Boolean
    Dim e As Long, id As Long, j As Long
    If AFActive Or RSubdivisions <> 8 Or RPX_LoadPhase <> "reference_total" Then Exit Function
    If rn <= 0 Or re <= 0 Then Exit Function
    For e = 0 To re - 1
        If RRigidId(e) < 0 Then
            If RMaterials(RMatId(e)).cohesion <> 0# Then Exit Function
        End If
        For j = 0 To 1: If RBody0(e, j) <> 0# Then Exit Function
        Next j
    Next e
    For id = 0 To RNEdge - 1
        For j = 0 To 1: If REdges(id).load0(j) <> 0# Then Exit Function
        Next j
    Next id
    For id = 0 To RR - 1
        For j = 0 To 2: If RRigids(id).load0(j) <> 0# Then Exit Function
        Next j
    Next id
    RPX_OriginalPhysicalZeroCostUpper = True
End Function
Private Function OriginalDualNorm2(ByVal a As Double, ByVal b As Double) As Double
    Dim largest As Double
    largest = Abs(a): If Abs(b) > largest Then largest = Abs(b)
    If largest = 0# Then Exit Function
    OriginalDualNorm2 = largest * Sqr((a / largest) ^ 2 + (b / largest) ^ 2)
End Function
Private Function OriginalDualAudit(ByRef y() As Double, ByRef z() As Double) As Boolean
    Dim stationarity() As Double, correction() As Double, rawY As Double
    Dim sep As Double, sepCorrection As Double, value As Double, normValue As Double
    Dim i As Long, p As Long, k As Long, r As Long, j As Long, first As Long
    On Error GoTo ArithmeticRejected
    OriginalFarkasStationarity = 0#: OriginalFarkasCone = 0#: OriginalFarkasSeparation = 0#
    If LBound(y) <> 0 Or UBound(y) < XM - 1 Or UBound(y) > XM Then Exit Function
    If LBound(z) <> 0 Or UBound(z) <> XS - 1 Then Exit Function
    ReDim stationarity(0 To XN - 1): ReDim correction(0 To XN - 1)
    For i = 0 To XM - 1
        If Not RPX_SchurFinite(y(i)) Or Not RPX_SchurFinite(XEqRowNorm(i)) Then Exit Function
        If XEqRowNorm(i) <= 0# Then Exit Function
        rawY = y(i) / XEqRowNorm(i)
        If Not RPX_SchurFinite(rawY) Then Exit Function
        AddCompensated sep, sepCorrection, -XOriginalCompiled.rhs(i) * rawY
        For p = XEPtr(i) To XEPtr(i + 1) - 1
            AddCompensated stationarity(XEId(p)), correction(XEId(p)), XOriginalCompiled.values(p) * rawY
        Next p
    Next i
    For k = 0 To XB - 1
        If XC(k).dimn <> 1 And XC(k).dimn <> 3 Then Exit Function
        first = XC(k).first
        For r = 0 To XC(k).dimn - 1
            If Not RPX_SchurFinite(z(first + r)) Then Exit Function
            AddCompensated sep, sepCorrection, -XC(k).offset(r) * z(first + r)
            For j = 0 To XC(k).count - 1
                AddCompensated stationarity(XC(k).ids(j)), correction(XC(k).ids(j)), -XC(k).coef(r, j) * z(first + r)
            Next j
        Next r
        value = -z(first)
        If XC(k).dimn = 3 Then value = value + OriginalDualNorm2(z(first + 1), z(first + 2))
        If value > OriginalFarkasCone Then OriginalFarkasCone = value
    Next k
    For i = 0 To XN - 1
        If Not RPX_SchurFinite(stationarity(i)) Then Exit Function
        If Abs(stationarity(i)) > OriginalFarkasStationarity Then OriginalFarkasStationarity = Abs(stationarity(i))
    Next i
    If Not RPX_SchurFinite(sep) Or Not RPX_SchurFinite(OriginalFarkasCone) Then Exit Function
    OriginalFarkasSeparation = sep: OriginalDualAudit = True
    Exit Function
ArithmeticRejected:
    If err.number <> 6 And err.number <> 9 Then err.Raise err.number, err.source, err.description
    err.Clear
End Function
Public Function RPX_OriginalFarkasGate(ByRef candidate As RPX_HsdCandidateState) As Boolean
    Const SEP_GOAL As Double = 1.000003
    Dim y() As Double, z() As Double, factor As Double, valueNormalization As Double, i As Long
    RPX_OriginalGuideClear: OriginalFarkasReason = "scope"
    On Error GoTo InvalidData
    If candidate.kind <> "PRIMAL_INFEASIBILITY_CERT" Or candidate.mode <> "upper" Then Exit Function
    If Not RFsActive Or Not candidate.numericalGate Or Not RPX_OriginalPhysicalZeroCostUpper() Then Exit Function
    If AFActive Or Not XOriginalCompiled.ready Or XOriginalCompiled.affine Then Exit Function
    If XOriginalCompiled.modeName <> "upper" Or XOriginalCompiled.subdivisions <> 8 Or candidate.subdivisions <> 8 Then Exit Function
    If candidate.fs <> RStrengthFactor Or candidate.solveEpoch <> P4Solve Then Exit Function
    If candidate.geometryEpoch <> RPX_HSDGeometryEpoch() Then Exit Function
    If candidate.inputEpoch <> RPX_InputEpoch Or XOriginalCompiled.inputEpoch <> RPX_InputEpoch Then Exit Function
    If XOriginalCompiled.n <> XN Or XOriginalCompiled.m <> XM Or XOriginalCompiled.s <> XS Or XOriginalCompiled.b <> XB Then Exit Function
    If candidate.modelKey <> XOriginalCompiled.modelKey Or candidate.modelKey <> RPX_ExactModel() Then Exit Function
    If candidate.loadPhase <> RPX_LoadPhase Or candidate.loadPhase <> XOriginalCompiled.loadPhase Then Exit Function
    OriginalFarkasReason = "nonzero_original_cost"
    If UBound(XOriginalCompiled.cost) <> XN - 1 Or UBound(XOriginalCompiled.rhs) < XM - 1 Or UBound(XOriginalCompiled.values) < XEPtr(XM) - 1 Then Exit Function
    For i = 0 To XN - 1
        If Not RPX_SchurFinite(XOriginalCompiled.cost(i)) Or XOriginalCompiled.cost(i) <> 0# Then Exit Function
    Next i
    OriginalFarkasReason = "raw_dual_or_separation"
    y = candidate.y: z = candidate.dual
    If Not OriginalDualAudit(y, z) Then Exit Function
    If OriginalFarkasSeparation <= 0# Then Exit Function
    factor = SEP_GOAL / OriginalFarkasSeparation
    If Not RPX_SchurFinite(factor) Or factor <= 0# Then Exit Function
    For i = 0 To UBound(y): y(i) = y(i) * factor: Next i
    For i = 0 To UBound(z): z(i) = z(i) * factor: Next i
    If Not OriginalDualAudit(y, z) Then Exit Function
    If OriginalFarkasStationarity > 0.0000001 Or OriginalFarkasCone > 0.00000001 Then Exit Function
    If OriginalFarkasSeparation < 1.000001 Or Abs(OriginalFarkasSeparation - SEP_GOAL) > 0.0000001 Then Exit Function
    If Not RPX_SchurFinite(candidate.normalization) Or candidate.normalization <= 0# Then Exit Function
    valueNormalization = candidate.normalization / factor
    If Not RPX_SchurFinite(valueNormalization) Or valueNormalization <= 0# Then Exit Function
    candidate.y = y: candidate.dual = z: candidate.normalization = valueNormalization
    originalGuideReady = True: originalGuideFs = candidate.fs: originalGuideQ = candidate.subdivisions
    originalGuideEpoch = candidate.inputEpoch: originalGuideSolve = candidate.solveEpoch
    originalGuideId = RPX_HSDCandidateId: originalGuideModel = candidate.modelKey: originalGuidePhase = candidate.loadPhase
    OriginalFarkasReason = "accepted_original_discrete_guide"
    CaptureOriginalGuide candidate
    RPX_DiagEvent "original_farkas_guide", "stationarity=" & OriginalFarkasStationarity & ";cone=" & OriginalFarkasCone & ";separation=" & OriginalFarkasSeparation & ";q=8;field_bound=0;original_row_dual_restored=1"
    RPX_OriginalFarkasGate = True
    Exit Function
InvalidData:
    If err.number <> 6 And err.number <> 9 Then err.Raise err.number, err.source, err.description
    err.Clear: OriginalFarkasReason = "invalid_or_nonfinite_original_dual"
End Function
Public Function RPX_OriginalGuideCurrent(ByVal factor As Double) As Boolean
    If Not RPX_SchurFinite(factor) Then Exit Function
    If XStatus <> "ORIGINAL_FARKAS_GUIDE" Or Not originalGuideReady Or AFActive Or Not RFsActive Then Exit Function
    If originalGuideFs <> factor Or factor <> RStrengthFactor Or originalGuideQ <> RSubdivisions Or RSubdivisions <> 8 Then Exit Function
    If originalGuideEpoch <> RPX_InputEpoch Or originalGuideSolve <> P4Solve Or originalGuideId <> RPX_HSDCandidateId Then Exit Function
    If originalGuideModel <> RPX_ExactModel() Or originalGuidePhase <> RPX_LoadPhase Then Exit Function
    If Not RPX_OriginalPhysicalZeroCostUpper() Then Exit Function
    RPX_OriginalGuideCurrent = True
End Function

' Optional immutable evidence for an accepted original dual guide, never a field.
Private Sub CaptureOriginalGuide(ByRef candidate As RPX_HsdCandidateState)
    Dim enabled As Long, path As String, f As Integer, i As Long, p As Long, k As Long, j As Long, r As Long
    Dim value As Double, count As Long, number As Long, source As String, message As String
    enabled = CLng(RPX_Setting("ORIGINAL_FARKAS_CAPTURE", 0))
    If enabled <> 0 And enabled <> 1 Then err.Raise 5, "RPX_Conic", "INVALID_ORIGINAL_FARKAS_CAPTURE"
    If enabled = 0 Or Len(RPX_DiagPath) = 0 Then Exit Sub
    path = RPX_DiagPath & ".original_guide_solve_" & P4Solve & "_candidate_" & RPX_HSDCandidateId
    If Len(Dir$(path & ".bin")) > 0 Then err.Raise 5, "RPX_Conic", "ORIGINAL_GUIDE_CAPTURE_EXISTS"
    On Error GoTo Failed
    f = FreeFile: Open path & ".bin" For Binary Access Write As #f
    RPX_ExchangeMagic f, "RG18F001"
    Put #f, , XN: Put #f, , XM: Put #f, , XB: Put #f, , XS: Put #f, , RSubdivisions
    Put #f, , candidate.fs: Put #f, , OriginalFarkasStationarity: Put #f, , OriginalFarkasCone: Put #f, , OriginalFarkasSeparation
    For i = 0 To XN - 1: value = XOriginalCompiled.cost(i): Put #f, , value: Next i
    For i = 0 To XM - 1
        value = XOriginalCompiled.rhs(i): Put #f, , value
        value = XEqRowNorm(i): Put #f, , value
        value = candidate.y(i): Put #f, , value
        value = candidate.y(i) / XEqRowNorm(i): Put #f, , value
        count = XEPtr(i + 1) - XEPtr(i): Put #f, , count
        For p = XEPtr(i) To XEPtr(i + 1) - 1
            Put #f, , XEId(p): value = XOriginalCompiled.values(p): Put #f, , value
        Next p
    Next i
    For k = 0 To XB - 1
        Put #f, , XC(k).dimn: Put #f, , XC(k).count: Put #f, , XC(k).first
        For j = 0 To XC(k).count - 1: Put #f, , XC(k).ids(j): Next j
        For r = 0 To XC(k).dimn - 1
            value = XC(k).offset(r): Put #f, , value
            For j = 0 To XC(k).count - 1: value = XC(k).coef(r, j): Put #f, , value: Next j
        Next r
    Next k
    For i = 0 To XS - 1: value = candidate.dual(i): Put #f, , value: Next i
    Close #f
    RPX_ExportExchangeContext path & ".context.tsv"
    RPX_DiagEvent "original_guide_capture", "path=" & path & ";original_program_and_dual=1;primal_field=0;physical_bound=0"
    Exit Sub
Failed:
    number = err.number: source = err.source: message = err.description
    On Error Resume Next: Close #f: On Error GoTo 0
    err.Raise number, source, message
End Sub
