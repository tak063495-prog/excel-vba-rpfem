Attribute VB_Name = "RPX_HSDLu"
Option Explicit
' Independent sparse row-pivoted LU: Pr * M = L * U.
' Stable physical row IDs carry earlier L entries through every row swap.
' Sparse arena, row/column links and an exact integer-key hash.
' pairId is an optional entry-id index, not a dense numeric matrix.
Private Const HluUseDirectIndex As Boolean = True
Private Const HluIndexMaxBytes As Double = 67108864#
Private Const HluIndexReserve As Double = 140000000#
Private Const HluArenaPeakBytes As Double = 44#
Private nr() As Long, nc() As Long, nv() As Double
Private pairId() As Long
Private pairReady As Boolean, pairN As Long
Private pairBytes As Double
Private rn() As Long, rp() As Long, cn() As Long, cp() As Long
Private rh() As Long, ch() As Long, slots() As Long
Private rowAt() As Long, position() As Long
Private order() As Long, inverseOrder() As Long, rowScale() As Double, colScale() As Double
Private lp() As Long, li() As Long, lv() As Double
Private up() As Long, ui() As Long, uv() As Double, diagonal() As Double
Private count As Long, capacity As Long, used As Long, freeHead As Long, live As Long
Private mask As Long, occupied As Long, updates As Double, pollCount As Long
Private finiteChecks As Double
Private budget As Double, overhead As Double, ready As Boolean
Private uCol() As Long, uVal() As Double, uCap As Long
Private arenaKept As Boolean
Private separateFactors As Boolean
Private extraHead() As Long, extraNext() As Long, extraCol() As Long, extraValue() As Double
Private extraUsed As Long, extraCapacity As Long, streamUUsed As Long, streamUCapacity As Long
Public HLUBytes As Double, HLUGrowth As Double, HLUMinPivot As Double, HLUMaxL As Double
Public HLUFactorEpoch As Long
Public Sub RPX_HLURelease()
    ready = False: count = 0: capacity = 0: used = 0: freeHead = 0: live = 0: occupied = 0: updates = 0: pollCount = 0
    Erase nr: Erase nc: Erase nv: Erase rn: Erase rp: Erase cn: Erase cp
    Erase rh: Erase ch: Erase slots: Erase rowAt: Erase position
    Erase order: Erase inverseOrder: Erase rowScale: Erase colScale
    Erase lp: Erase li: Erase lv: Erase up: Erase ui: Erase uv: Erase diagonal
    Erase uCol: Erase uVal: uCap = 0: arenaKept = False
    HLUFactorEpoch = 0
    DropPairIndex
    Erase extraHead: Erase extraNext: Erase extraCol: Erase extraValue
    extraUsed = 0: extraCapacity = 0: streamUUsed = 0: streamUCapacity = 0
End Sub
Public Sub RPX_HLUPark()
    ' Drop the numeric factor. Keep the arena allocation for the next factorization.
    ready = False: count = 0: used = 0: freeHead = 0: live = 0: occupied = 0: updates = 0: pollCount = 0
    HLUFactorEpoch = 0
    Erase rh: Erase ch: Erase slots
    Erase rowAt: Erase position: Erase order: Erase inverseOrder: Erase rowScale: Erase colScale
    Erase lp: Erase li: Erase lv: Erase up: Erase ui: Erase uv: Erase diagonal
    arenaKept = (capacity > 0)
    DropPairIndex
    Erase extraHead: Erase extraNext: Erase extraCol: Erase extraValue
    extraUsed = 0: extraCapacity = 0: streamUUsed = 0: streamUCapacity = 0
End Sub
Private Function LongBound(ByRef values() As Long, ByVal last As Long) As Boolean
    Dim upper As Long, eno As Long
    On Error Resume Next
    upper = UBound(values)
    eno = err.number
    err.Clear
    On Error GoTo 0
    If eno <> 0 Then Exit Function
    LongBound = (LBound(values) = 0 And upper >= last)
End Function
Private Sub CheckMemory(ByVal nodeCapacity As Long, ByVal hashCapacity As Long)
    ' Includes old/new arenas during growth, compressed factors, caller's live data, and the entry index.
    ' Seven arena arrays: 32 bytes/node. Sequential Preserve peaks at <=40;
    ' compressed L/U add <=12/node, hence 44 bounds both (not simultaneous).
    ' Old compressed factors are released before building the next arena.
    Dim arenaBytes As Double, factorBytes As Double
    arenaBytes = HluArenaPeakBytes
    If separateFactors Then
        arenaBytes = 40#: factorBytes = 28# * (extraCapacity + 1#) + 24# * (streamUCapacity + 1#)
    End If
    HLUBytes = overhead + (nodeCapacity + 1#) * arenaBytes + hashCapacity * 8# + count * 128# + pairBytes + factorBytes
    If HLUBytes > budget Then err.Raise 5, "RPX_HSDLu", "HSD_STABLE_UNAVAILABLE_MEMORY"
End Sub
Private Sub DropPairIndex()
    pairReady = False: pairN = 0: pairBytes = 0#
    Erase pairId
End Sub
Private Sub PreparePairIndex(ByVal n As Long)
    Dim slotCount As Double, indexBytes As Double, allocError As Long, lastIndex As Long
    Dim allocSource As String, allocText As String
    DropPairIndex
    If Not HluUseDirectIndex Then Exit Sub
    If n < 1 Then Exit Sub
    slotCount = CDbl(n) * CDbl(n)
    If slotCount > 2147483647# Then Exit Sub
    indexBytes = slotCount * 4#
    If indexBytes > HluIndexMaxBytes Then Exit Sub
    If indexBytes + HluIndexReserve > budget Then Exit Sub
    lastIndex = CLng(slotCount - 1#)
    On Error Resume Next
    ReDim pairId(0 To lastIndex)
    allocError = err.number
    allocSource = err.source: allocText = err.description
    err.Clear
    On Error GoTo 0
    If allocError <> 0 Then
        DropPairIndex
        err.Raise allocError, allocSource, allocText
    End If
    pairN = n
    pairBytes = indexBytes
    pairReady = True
End Sub
Private Function Bucket(ByVal r As Long, ByVal c As Long) As Long
    Dim mixed As Double
    ' Hashing selects a bucket only; exact Long row/column comparisons prove identity.
    mixed = (CDbl(r) * count + c + 1#) * 0.618033988749895
    Bucket = CLng(Fix((mixed - Fix(mixed)) * (mask + 1)))
End Function
Private Function Locate(ByVal r As Long, ByVal c As Long, ByRef insertAt As Long) As Long
    Dim h As Long, id As Long, tomb As Long
    h = Bucket(r, c): tomb = -1
    Do
        id = slots(h)
        If id = 0 Then
            If tomb >= 0 Then insertAt = tomb Else insertAt = h
            Exit Function
        ElseIf id = -1 Then
            If tomb < 0 Then tomb = h
        ElseIf nr(id) = r And nc(id) = c Then
            Locate = id: insertAt = h: Exit Function
        End If
        h = (h + 1) And mask
    Loop
End Function
Private Sub Rehash(ByVal size As Long)
    Dim id As Long, h As Long
    CheckMemory capacity, size
    ReDim slots(0 To size - 1): mask = size - 1: occupied = 0
    For id = 1 To used
        If nr(id) >= 0 Then
            h = Bucket(nr(id), nc(id))
            Do While slots(h) <> 0: h = (h + 1) And mask: Loop
            slots(h) = id: occupied = occupied + 1
        End If
    Next id
End Sub
Private Sub GrowArena()
    Dim nextCapacity As Long, affordable As Double, arenaBytes As Double, factorBytes As Double
    nextCapacity = capacity * 2: If nextCapacity < 1024 Then nextCapacity = 1024
    arenaBytes = HluArenaPeakBytes
    If separateFactors Then arenaBytes = 40#: factorBytes = 28# * (extraCapacity + 1#) + 24# * (streamUCapacity + 1#)
    affordable = Fix((budget - overhead - (mask + 1) * 8# - count * 128# - pairBytes - factorBytes) / arenaBytes) - 1#
    If affordable < nextCapacity Then
        If affordable <= capacity Then err.Raise 5, "RPX_HSDLu", "HSD_STABLE_UNAVAILABLE_MEMORY"
        nextCapacity = CLng(affordable)
        RPX_DiagEvent "hsd_arena_budget", "requested=" & (capacity * 2#) & ";allocated=" & nextCapacity & ";limit=" & budget
    End If
    CheckMemory nextCapacity, mask + 1
    ReDim Preserve nr(0 To nextCapacity): ReDim Preserve nc(0 To nextCapacity): ReDim Preserve nv(0 To nextCapacity)
    ReDim Preserve rn(0 To nextCapacity): ReDim Preserve rp(0 To nextCapacity)
    ReDim Preserve cn(0 To nextCapacity): ReDim Preserve cp(0 To nextCapacity)
    capacity = nextCapacity
End Sub
Private Sub PutEntry(ByVal r As Long, ByVal c As Long, ByVal value As Double)
    Dim id As Long, h As Long, a As Long, b As Long, size As Long
    updates = updates + 1
    pollCount = pollCount + 1
    If pollCount >= 262144 Then
        pollCount = 0
        RPX_Stage "HSD sparse updates " & CStr(updates)
    End If
    finiteChecks = finiteChecks + 1
    If Not RPX_SchurFinite(value) Then err.Raise 5, , "HSD_NONFINITE_FACTOR"
    id = Locate(r, c, h)
    If id > 0 Then
        If value <> 0# Then nv(id) = value: Exit Sub
        a = rp(id): b = rn(id)
        If a > 0 Then rn(a) = b Else rh(r) = b
        If b > 0 Then rp(b) = a
        a = cp(id): b = cn(id)
        If a > 0 Then cn(a) = b Else ch(c) = b
        If b > 0 Then cp(b) = a
        If pairReady Then pairId(r * pairN + c) = 0
        slots(h) = -1: nr(id) = -1: rn(id) = freeHead: freeHead = id: live = live - 1
    ElseIf value <> 0# Then
        If occupied * 10# >= (mask + 1) * 7# Then
            size = mask + 1: If live * 10# >= size * 4# Then size = size * 2
            Rehash size: id = Locate(r, c, h)
        End If
        If freeHead > 0 Then
            id = freeHead: freeHead = rn(id)
        Else
            If used = capacity Then GrowArena
            used = used + 1: id = used
        End If
        nr(id) = r: nc(id) = c: nv(id) = value
        rp(id) = 0: rn(id) = rh(r): If rh(r) > 0 Then rp(rh(r)) = id
        rh(r) = id
        cp(id) = 0: cn(id) = ch(c): If ch(c) > 0 Then cp(ch(c)) = id
        ch(c) = id
        If slots(h) = 0 Then occupied = occupied + 1
        slots(h) = id: live = live + 1
        If pairReady Then pairId(r * pairN + c) = id
    End If
End Sub
Public Function RPX_HLUFactor(ByRef ap() As Long, ByRef ai() As Long, ByRef av() As Double, ByVal maxBytes As Double, ByVal coexistingBytes As Double, ByVal epoch As Long, Optional ByVal useOriginalOrder As Boolean = False, Optional ByVal equilibrate As Boolean = False, Optional ByVal symmetricStorage As Boolean = True, Optional ByVal diagonalThreshold As Double = 0#, Optional ByVal columnPermutation As Variant, Optional ByVal streamFactors As Boolean = False) As Boolean
    Dim n As Long, c As Long, p As Long, r As Long, k As Long, best As Long, old As Long, h As Long
    Dim id As Long, nextID As Long, j As Long, target As Long, seen() As Long
    Dim uN As Long, uAt As Long
    Dim value As Double, pivot As Double, largest As Double, inputMax As Double, factorMax As Double, mult As Double
    Dim lCount As Long, uCount As Long, a As Long, b As Long, mappedValue As Double
    Dim rowIndex() As Long, rowStamp() As Long, stamp As Long, scan As Long
    Dim mark As Long, failNumber As Long, failSource As String, failText As String
    Dim finiteProbe As Boolean, rejectFinite As Boolean
    Dim scanNodes As Double, scanRows As Double
    Dim pairBase As Long, indexDirect As Long
    mark = RPX_HoptMark()
    On Error GoTo HoptFailed
    ready = False
    Erase lp: Erase li: Erase lv: Erase up: Erase ui: Erase uv: Erase diagonal
    If arenaKept And capacity >= 1024 Then
        ready = False: used = 0: freeHead = 0: live = 0: occupied = 0: updates = 0: pollCount = 0
        HLUFactorEpoch = 0
    Else
        RPX_HLURelease
    End If
    separateFactors = streamFactors
    Erase extraHead: Erase extraNext: Erase extraCol: Erase extraValue
    extraUsed = 0: extraCapacity = 0: streamUUsed = 0: streamUCapacity = 0
    finiteChecks = 0
    n = UBound(ap)
    If diagonalThreshold < 0# Or diagonalThreshold > 1# Then err.Raise 5, , "HSD_PIVOT_THRESHOLD_INVALID"
    If LBound(ap) <> 0 Or n < 1 Or ap(0) <> 0 Then err.Raise 5, , "HSD_KKT_STORAGE_INVALID"
    If n > 1048576 Then err.Raise 5, , "HSD_STABLE_UNAVAILABLE_SIZE"
    If ap(n) <> UBound(av) + 1 Or UBound(ai) <> UBound(av) Then err.Raise 5, , "HSD_KKT_STORAGE_INVALID"
    count = n: budget = maxBytes: overhead = coexistingBytes + 8# * n
    CheckMemory 1024, 2048
    ReDim seen(0 To n - 1)
    If arenaKept And LongBound(rh, n - 1) And LongBound(ch, n - 1) Then
        For r = 0 To n - 1: rh(r) = 0: ch(r) = 0: Next r
    Else
        ReDim rh(0 To n - 1): ReDim ch(0 To n - 1)
    End If
    ReDim rowAt(0 To n - 1): ReDim position(0 To n - 1)
    ReDim order(0 To n - 1): ReDim inverseOrder(0 To n - 1)
    ReDim rowScale(0 To n - 1): ReDim colScale(0 To n - 1)
    ReDim rowIndex(0 To n - 1): ReDim rowStamp(0 To n - 1)
    For k = 0 To n - 1
        order(k) = k: If useOriginalOrder Then order(k) = XInv(k)
        If Not IsMissing(columnPermutation) Then order(k) = CLng(columnPermutation(k))
        If order(k) < 0 Or order(k) >= n Then err.Raise 5, , "HSD_ORDER_INVALID"
        If seen(order(k)) = 1 Then err.Raise 5, , "HSD_ORDER_INVALID"
        seen(order(k)) = 1: inverseOrder(order(k)) = k
    Next k
    ReDim slots(0 To 2047): mask = 2047
    If capacity < 1024 Then GrowArena
    arenaKept = True
    For r = 0 To n - 1: rowAt(r) = r: position(r) = r: seen(r) = -1: Next r
    For c = 0 To n - 1
        For p = ap(c) To ap(c + 1) - 1
            r = ai(p)
            If r < 0 Or r >= n Then err.Raise 5, , "HSD_KKT_STORAGE_INVALID row"
            If symmetricStorage And r > c Then err.Raise 5, , "HSD_KKT_STORAGE_INVALID triangle"
            value = Abs(av(p))
            If value > rowScale(r) Then rowScale(r) = value
            If symmetricStorage And value > rowScale(c) Then rowScale(c) = value
        Next p
    Next c
    For r = 0 To n - 1
        If equilibrate And rowScale(r) > 0# Then rowScale(r) = 1# / rowScale(r) Else rowScale(r) = 1#
    Next r
    For c = 0 To n - 1
        For p = ap(c) To ap(c + 1) - 1
            r = ai(p): value = Abs(av(p) * rowScale(r))
            If value > colScale(c) Then colScale(c) = value
            If symmetricStorage Then
                value = Abs(av(p) * rowScale(c))
                If value > colScale(r) Then colScale(r) = value
            End If
        Next p
    Next c
    For c = 0 To n - 1
        If equilibrate And colScale(c) > 0# Then colScale(c) = 1# / colScale(c) Else colScale(c) = 1#
    Next c
    RPX_HoptEnter HoptLuStore
    PreparePairIndex n
    indexDirect = 0: If pairReady Then indexDirect = 1
    CheckMemory 1024, 2048
    For c = 0 To n - 1
        If ap(c + 1) < ap(c) Then err.Raise 5, , "HSD_KKT_STORAGE_INVALID"
        For p = ap(c) To ap(c + 1) - 1
            r = ai(p)
            If r < 0 Or r >= n Then err.Raise 5, , "HSD_KKT_STORAGE_INVALID row"
            If symmetricStorage And r > c Then err.Raise 5, , "HSD_KKT_STORAGE_INVALID triangle"
            If seen(r) = c Then err.Raise 5, , "HSD_KKT_STORAGE_INVALID duplicate"
            seen(r) = c: value = av(p)
            finiteChecks = finiteChecks + 1
            If Not RPX_SchurFinite(value) Then err.Raise 5, , "HSD_NONFINITE_MATRIX"
            mappedValue = value * rowScale(r) * colScale(c)
            If Abs(mappedValue) > inputMax Then inputMax = Abs(mappedValue)
            PutEntry r, inverseOrder(c), mappedValue
            If symmetricStorage And r <> c Then
                mappedValue = value * rowScale(c) * colScale(r)
                If Abs(mappedValue) > inputMax Then inputMax = Abs(mappedValue)
                PutEntry c, inverseOrder(r), mappedValue
            End If
        Next p
    Next c
    RPX_HoptLeave HoptLuStore
    HLUMinPivot = 1E+300: HLUMaxL = 0#: factorMax = inputMax
    If uCap < n Then
        ReDim uCol(1 To n): ReDim uVal(1 To n): uCap = n
    End If
    If separateFactors Then
        ReDim extraHead(0 To n - 1): ReDim up(0 To n): ReDim diagonal(0 To n - 1)
    End If
    RPX_HoptEnter HoptLuEliminate
    ' Existing-entry updates only. Same predicate as RPX_SchurFinite.
    ' Overflow while comparing (quiet NaN) is still HSD_NONFINITE_FACTOR.
    ' Subtraction overflow is not inside finiteProbe, so it stays error 6.
    finiteProbe = False
    On Error GoTo FiniteFailed
    For k = 0 To n - 1
        If k Mod 128 = 0 Then
            RPX_AHOCheckBudget
            RPX_Stage "HSD sparse LU " & k & "/" & n
        End If
        largest = 0#: best = -1: id = ch(k)
        Do While id > 0
            r = nr(id)
            If position(r) >= k Then
                value = Abs(nv(id))
                If value > largest Or (value = largest And (best = -1 Or r < best)) Then largest = value: best = r
            End If
            id = cn(id)
        Loop
        ' Independent general-LU route only. Legacy default keeps largest-row pivoting.
        If diagonalThreshold > 0# And largest > 0# Then
            r = order(k)
            If position(r) >= k Then id = Locate(r, k, h) Else id = 0
            If id > 0 Then
                If Abs(nv(id)) >= diagonalThreshold * largest Then best = r
            End If
        End If
        If best < 0 Or largest = 0# Then
            RPX_HoptLeave HoptLuEliminate
            DropPairIndex
    Call RPX_HoptLuGauge(0, 0, CDbl(updates), CDbl(capacity), HLUBytes, CDbl(finiteChecks), scanNodes, scanRows, indexDirect)
            Exit Function
        End If
        old = rowAt(k): p = position(best)
        rowAt(k) = best: position(best) = k: rowAt(p) = old: position(old) = p
        id = Locate(best, k, h): pivot = nv(id)
        If Abs(pivot) < HLUMinPivot Then HLUMinPivot = Abs(pivot)
        uN = 0: j = rh(best)
        Do While j > 0
            c = nc(j)
            If c > k Then
                uN = uN + 1: uCol(uN) = c: uVal(uN) = nv(j)
            End If
            j = rn(j)
        Loop
        If separateFactors Then
            diagonal(k) = pivot: up(k) = streamUUsed
            For uAt = 1 To uN: StreamUpper uCol(uAt), uVal(uAt): Next uAt
            up(k + 1) = streamUUsed
        End If
        id = ch(k)
        Do While id > 0
            nextID = cn(id): r = nr(id)
            If position(r) > k Then
                mult = nv(id) / pivot
                finiteChecks = finiteChecks + 1
                If Not RPX_SchurFinite(mult) Then err.Raise 5, , "HSD_NONFINITE_FACTOR"
                If separateFactors Then
                    StreamLower r, k, mult
                    PutEntry r, k, 0#
                Else
                    nv(id) = mult
                End If
                If Abs(mult) > HLUMaxL Then HLUMaxL = Abs(mult)
                ' Same arithmetic on both paths. Direct uses the physical row and LU column.
                If pairReady Then
                    pairBase = r * pairN
                    For uAt = 1 To uN
                        c = uCol(uAt)
                        target = pairId(pairBase + c)
                        value = 0#
                        If target > 0 Then value = nv(target)
                        value = value - mult * uVal(uAt)
                        If Abs(value) > factorMax Then factorMax = Abs(value)
                        If target > 0 And value <> 0# Then
                            finiteChecks = finiteChecks + 1
                            finiteProbe = True
                            rejectFinite = Not ((value >= -1E+290) And (value <= 1E+290) And (value = value))
                            finiteProbe = False
                            If rejectFinite Then
                                If Not RPX_SchurFinite(value) Then err.Raise 5, , "HSD_NONFINITE_FACTOR"
                            End If
                            nv(target) = value: updates = updates + 1: pollCount = pollCount + 1
                            If pollCount >= 262144 Then pollCount = 0: RPX_Stage "HSD sparse updates " & CStr(updates)
                        Else
                            PutEntry r, c, value
                        End If
                    Next uAt
                Else
                    If stamp = 2147483647 Then ReDim rowStamp(0 To n - 1): stamp = 0
                    stamp = stamp + 1: scan = rh(r): scanRows = scanRows + 1
                    Do While scan > 0
                        scanNodes = scanNodes + 1
                        rowIndex(nc(scan)) = scan: rowStamp(nc(scan)) = stamp: scan = rn(scan)
                    Loop
                    For uAt = 1 To uN
                        c = uCol(uAt)
                        target = 0: If rowStamp(c) = stamp Then target = rowIndex(c)
                        value = 0#
                        If target > 0 Then value = nv(target)
                        value = value - mult * uVal(uAt)
                        If Abs(value) > factorMax Then factorMax = Abs(value)
                        If target > 0 And value <> 0# Then
                            finiteChecks = finiteChecks + 1
                            finiteProbe = True
                            rejectFinite = Not ((value >= -1E+290) And (value <= 1E+290) And (value = value))
                            finiteProbe = False
                            If rejectFinite Then
                                If Not RPX_SchurFinite(value) Then err.Raise 5, , "HSD_NONFINITE_FACTOR"
                            End If
                            nv(target) = value: updates = updates + 1: pollCount = pollCount + 1
                            If pollCount >= 262144 Then pollCount = 0: RPX_Stage "HSD sparse updates " & CStr(updates)
                        Else
                            PutEntry r, c, value
                        End If
                    Next uAt
                End If
            End If
            id = nextID
        Loop
        If separateFactors Then
            id = rh(best)
            Do While id > 0
                nextID = rn(id): PutEntry best, nc(id), 0#: id = nextID
            Loop
        End If
    Next k
    finiteProbe = False
    On Error GoTo HoptFailed
    RPX_HoptLeave HoptLuEliminate
    RPX_HoptEnter HoptLuCompress
    If inputMax > 0# Then HLUGrowth = factorMax / inputMax Else HLUGrowth = 0#
    If separateFactors Then
        ReDim lp(0 To n)
        For k = 0 To n - 1
            id = extraHead(rowAt(k))
            Do While id > 0: lCount = lCount + 1: id = extraNext(id): Loop
            lp(k + 1) = lCount
        Next k
        CheckMemory capacity, mask + 1
        ReDim li(0 To lCount): ReDim lv(0 To lCount)
        uCount = streamUUsed
        For k = 0 To n - 1
            id = extraHead(rowAt(k)): a = lp(k)
            Do While id > 0
                li(a) = extraCol(id): lv(a) = extraValue(id): a = a + 1: id = extraNext(id)
            Loop
            SortPairs li, lv, lp(k), lp(k + 1) - 1
            SortPairs ui, uv, up(k), up(k + 1) - 1
        Next k
        Erase extraHead: Erase extraNext: Erase extraCol: Erase extraValue
    Else
    ReDim lp(0 To n): ReDim up(0 To n): ReDim diagonal(0 To n - 1)
    For k = 0 To n - 1
        id = rh(rowAt(k))
        Do While id > 0
            c = nc(id)
            If c < k Then lCount = lCount + 1
            If c > k Then uCount = uCount + 1
            If c = k Then diagonal(k) = nv(id)
            id = rn(id)
        Loop
        lp(k + 1) = lCount: up(k + 1) = uCount
    Next k
    ReDim li(0 To lCount): ReDim lv(0 To lCount): ReDim ui(0 To uCount): ReDim uv(0 To uCount)
    For k = 0 To n - 1
        id = rh(rowAt(k)): a = lp(k): b = up(k)
        Do While id > 0
            c = nc(id)
            If c < k Then li(a) = c: lv(a) = nv(id): a = a + 1
            If c > k Then ui(b) = c: uv(b) = nv(id): b = b + 1
            id = rn(id)
        Loop
        SortPairs li, lv, lp(k), lp(k + 1) - 1
        SortPairs ui, uv, up(k), up(k + 1) - 1
    Next k
    End If
    Erase rh: Erase ch: Erase slots
    ready = True: HLUFactorEpoch = epoch: RPX_HLUFactor = True
    arenaKept = (capacity > 0)
    RPX_HoptLeave HoptLuCompress
    DropPairIndex
    Call RPX_HoptLuGauge(lCount, uCount, CDbl(updates), CDbl(capacity), HLUBytes, CDbl(finiteChecks), scanNodes, scanRows, indexDirect)
    Exit Function
FiniteFailed:
    failNumber = err.number: failSource = err.source: failText = err.description
    On Error GoTo HoptFailed
    If finiteProbe And failNumber = 6 Then
        finiteProbe = False
        err.Raise 5, , "HSD_NONFINITE_FACTOR"
    End If
    err.Raise failNumber, failSource, failText
HoptFailed:
    failNumber = err.number: failSource = err.source: failText = err.description
    RPX_HoptUnwind mark
    DropPairIndex
    Call RPX_HoptLuGauge(lCount, uCount, CDbl(updates), CDbl(capacity), HLUBytes, CDbl(finiteChecks), scanNodes, scanRows, indexDirect)
    err.Raise failNumber, failSource, failText
End Function
Private Sub StreamLower(ByVal physicalRow As Long, ByVal pivotColumn As Long, ByVal value As Double)
    If extraUsed >= extraCapacity Then
        extraCapacity = extraCapacity * 2: If extraCapacity < 32768 Then extraCapacity = 32768
        CheckMemory capacity, mask + 1
        ReDim Preserve extraNext(0 To extraCapacity): ReDim Preserve extraCol(0 To extraCapacity): ReDim Preserve extraValue(0 To extraCapacity)
    End If
    extraUsed = extraUsed + 1
    extraNext(extraUsed) = extraHead(physicalRow): extraHead(physicalRow) = extraUsed
    extraCol(extraUsed) = pivotColumn: extraValue(extraUsed) = value
End Sub
Private Sub StreamUpper(ByVal pivotColumn As Long, ByVal value As Double)
    If streamUUsed >= streamUCapacity Then
        streamUCapacity = streamUCapacity * 2: If streamUCapacity < 32768 Then streamUCapacity = 32768
        CheckMemory capacity, mask + 1
        ReDim Preserve ui(0 To streamUCapacity - 1): ReDim Preserve uv(0 To streamUCapacity - 1)
    End If
    ui(streamUUsed) = pivotColumn: uv(streamUUsed) = value: streamUUsed = streamUUsed + 1
End Sub
Private Sub SortPairs(ByRef ids() As Long, ByRef vals() As Double, ByVal first As Long, ByVal last As Long)
    Dim a As Long, b As Long, pivot As Long, id As Long, value As Double
    If first >= last Then Exit Sub
    a = first: b = last: pivot = ids((first + last) \ 2)
    Do While a <= b
        Do While ids(a) < pivot: a = a + 1: Loop
        Do While ids(b) > pivot: b = b - 1: Loop
        If a <= b Then
            id = ids(a): ids(a) = ids(b): ids(b) = id
            value = vals(a): vals(a) = vals(b): vals(b) = value
            a = a + 1: b = b - 1
        End If
    Loop
    If first < b Then SortPairs ids, vals, first, b
    If a < last Then SortPairs ids, vals, a, last
End Sub
Public Sub RPX_HLUSolve(ByRef rhs() As Double, ByRef answer() As Double, ByVal epoch As Long)
    Dim i As Long, p As Long, value As Double, work() As Double, solvedChecks As Long
    Dim mark As Long, failNumber As Long, failSource As String, failText As String
    mark = RPX_HoptMark()
    On Error GoTo HoptFailed
    Erase answer
    If Not ready Or HLUFactorEpoch <> epoch Then err.Raise 5, "RPX_HSDLu", "HSD_FACTOR_CONTEXT_MISMATCH"
    If LBound(rhs) <> 0 Or UBound(rhs) <> count - 1 Then err.Raise 5, , "HSD_RHS_SHAPE_MISMATCH"
    ReDim answer(0 To count - 1): ReDim work(0 To count - 1)
    For i = 0 To count - 1
        value = rhs(rowAt(i)) * rowScale(rowAt(i))
        For p = lp(i) To lp(i + 1) - 1: value = value - lv(p) * work(li(p)): Next p
        work(i) = value
    Next i
    For i = count - 1 To 0 Step -1
        value = work(i)
        For p = up(i) To up(i + 1) - 1: value = value - uv(p) * work(ui(p)): Next p
        work(i) = value / diagonal(i)
        solvedChecks = solvedChecks + 1
        If Not RPX_SchurFinite(work(i)) Then err.Raise 5, , "HSD_NONFINITE_SOLUTION"
    Next i
    For i = 0 To count - 1: answer(order(i)) = colScale(order(i)) * work(i): Next i
    Call RPX_HoptAddFinite(CDbl(solvedChecks))
    Exit Sub
HoptFailed:
    failNumber = err.number: failSource = err.source: failText = err.description
    Call RPX_HoptAddFinite(CDbl(solvedChecks))
    RPX_HoptUnwind mark
    err.Raise failNumber, failSource, failText
End Sub
