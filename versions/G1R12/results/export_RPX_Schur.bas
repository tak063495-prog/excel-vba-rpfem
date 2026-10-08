Option Explicit
' P3-C1/C2 compare on independent sparse state; C3 AUTO uses full residual guards.
Private Type LocalBlock
    pos(0 To 2) As Long
    hp(0 To 2, 0 To 2) As Long
    chol(0 To 2, 0 To 2) As Double
    n As Long
    row() As Long
    bp() As Long
    bv() As Double
    solved() As Double
    Slot() As Long
End Type
Private blocks() As LocalBlock, nb As Long, nr As Long, nf As Long
Private blockOf() As Long, component() As Long, retained() As Long, fullPos() As Long
Private ap() As Long, ai() As Long, av() As Double, perm() As Long, inv() As Long
Private directSlot() As Long, ready As Boolean, factored As Boolean, enabled As Boolean
Private lastPulse As Double
Private fault As String, autoDirection As Long
Private autoMode As Boolean
Public RPX_SchurAutoFailed As Boolean
Private Capture As Boolean, sequence As Long, directionNo As Long, reg As Double
Private budget As Double, estimate As Double
Private lastLower As Boolean
Private contextChecked As Boolean, factorEpoch As Long, z() As Double, solving As Boolean
Private workspacePattern As Long, workspaceRhs As Long
Public RPX_SchurComparisons As Long, RPX_SchurRejected As Long

Public Sub RPX_SchurReset(Optional ByVal lower As Boolean = False, Optional ByVal prepared As Boolean = False)
    Dim preservePattern As Boolean, profile As RPX_ProblemProfile, reason As String
    preservePattern = ready And ((lower And prepared) Or P4Cache)
    autoMode = False: RPX_SchurAutoFailed = False
    factored = False: factorEpoch = -1: enabled = False: Capture = False
    If Not preservePattern Then RPX_SchurRelease
    ' A prepared, continuously active lower layout is the existing P3 lifetime.
    ' A return from upper (including P3-B restoration) always needs exact lookup.
    contextChecked = lower And prepared And lastLower And ready
    lastLower = lower
    sequence = sequence + 1: directionNo = 0
    If Not lower Then Exit Sub
    Capture = (CLng(RPX_Setting("KKT_CAPTURE", 0)) = 1)
    enabled = (UCase$(CStr(RPX_Setting("LOWER_KKT_BACKEND", "GENERIC"))) = "COMPARE")
    autoMode = (UCase$(CStr(RPX_Setting("LOWER_KKT_BACKEND", "GENERIC"))) = "AUTO") And Not Capture
    enabled = enabled Or autoMode
    budget = CDbl(RPX_Setting("SCHUR_MAX_BYTES", 268435456#))
    If budget <= 0# Then budget = 268435456#
    If Capture Then RPX_DiagEvent "kkt_capture_mode", "enabled=1;adopted=GENERIC"
    If enabled Then
        RPX_DiagEvent "schur_engine", "engine=20260927_p3c;backend=" & IIf(autoMode, "AUTO", "COMPARE") & ";full_refinements=3;residual_limit=1e-11"
        RPX_BuildModelProfile profile
        enabled = RPX_CanUseFeature("LOWER_KKT_BACKEND", profile, reason)
        If Not enabled Then
            RPX_SchurRelease
            RPX_DiagEvent "schur_bypass", "reason=" & reason
        End If
        If ready Then
            If estimate > budget Or 28# * RPXC_FactorNNZ + 64# * nr > budget - estimate Then
                RPX_SchurRelease
                RPX_DiagEvent "schur_bypass", "reason=budget"
            End If
        End If
    End If
End Sub

Public Function RPX_SchurResidentBytes() As Double
    If ready Then RPX_SchurResidentBytes = estimate + 28# * RPXC_FactorNNZ + 64# * nr
End Function
Public Sub RPX_SchurDeactivate()
    factored = False: factorEpoch = -1: enabled = False
    P4Generic = -1: P4Current.numericValid = False
    workspaceRhs = -1
End Sub
Public Sub RPX_SchurRelease()
    RPX_SchurDeactivate
    lastLower = False
    ready = False: factored = False: enabled = False: factorEpoch = -1
    contextChecked = False: solving = False: workspacePattern = -1: workspaceRhs = -1
    Erase blocks: Erase blockOf: Erase component: Erase retained: Erase fullPos
    Erase ap: Erase ai: Erase av: Erase perm: Erase inv: Erase directSlot: Erase z
    Call RPXC_LDLRelease: RPX_SchurIdentityRelease
End Sub

Private Function Slot(ByRef rows() As Object, ByVal a As Long, ByVal b As Long) As Long
    Dim t As Long
    If a > b Then t = a: a = b: b = t
    Slot = CLng(rows(b)(a))
End Function
Private Sub Pattern(ByRef rows() As Object, ByVal a As Long, ByVal b As Long)
    Dim t As Long
    If a > b Then t = a: a = b: b = t
    rows(b)(a) = 0
End Sub

Private Function Prepare(ByRef reason As String) As Boolean
    Dim e As Long, k As Long, j As Long, i As Long, b As Long, c As Long, p As Long
    Dim u As Long, v As Long, bu As Long, bv As Long, r As Long, s As Long, id As Long
    Dim neighbors() As Object, rows() As Object, ordered() As Object, key As Variant
    Dim profile As RPX_ProblemProfile, seen() As Boolean, coneBlock As Long
    Dim reaction As Variant, uniqueReactions As Object
    RPX_TestHook "p4_prepare"
    RPX_BuildModelProfile profile
    If profile.soilElements = 0 Or profile.c0Soils > 0 Or profile.phi0Soils > 0 Or profile.hsd Then reason = "material_or_hsd": Exit Function
    If RR > 0 And CLng(RPX_Setting("SCHUR_ALLOW_RIGID", 0)) <> 1 Then reason = "rigid_disabled": Exit Function
    nf = XN + XM: nb = 6 * profile.soilElements: nr = nf - 3 * nb
    If nr <= 0 Then reason = "layout": Exit Function
    estimate = 256# * nf + 128# * (UBound(XKVal) + 1#) + 24# * nb
    If P4Cache Then
        If Not RPX_SchurIdentityStore(budget - estimate) Then reason = "identity_budget": Exit Function
        estimate = estimate + P4IdentityBytes
    End If
    If estimate > budget Then reason = "budget": Exit Function
    RPX_CapacityCheck "schur_prepare", estimate
    ReDim blocks(0 To nb - 1): ReDim blockOf(0 To nf - 1): ReDim component(0 To nf - 1)
    ReDim retained(0 To nf - 1): ReDim fullPos(0 To nr - 1): ReDim neighbors(0 To nb - 1)
    For i = 0 To nf - 1: blockOf(i) = -1: retained(i) = -1: Next i
    b = 0
    For e = 0 To re - 1
        If RRigidId(e) < 0 Then
            For k = 0 To 5
                Set neighbors(b) = CreateObject("Scripting.Dictionary")
                For j = 0 To 2
                    id = RPX_StressId(e, k, j)
                    If id < 0 Or id >= XN Then reason = "stress_id": Exit Function
                    p = XInv(id)
                    If blockOf(p) >= 0 Then reason = "overlap": Exit Function
                    blockOf(p) = b: component(p) = j: blocks(b).pos(j) = p
                    For c = 0 To 2: blocks(b).hp(j, c) = -1: Next c
                Next j
                b = b + 1
            Next k
        End If
    Next e
    ReDim seen(0 To nb - 1)
    For i = 0 To XB - 1
        With XC(i)
            If .dimn <> 3 Or .count <> 3 Then reason = "cone_layout": Exit Function
            If .owner < 0 Or .owner >= re Or .parameter1 < 0 Or .parameter1 > 5 Then reason = "cone_owner": Exit Function
            For j = 0 To 2
                If .ids(j) <> RPX_StressId(.owner, .parameter1, j) Then reason = "cone_point": Exit Function
            Next j
            coneBlock = blockOf(XInv(.ids(0)))
            If coneBlock < 0 Then reason = "retained_cone": Exit Function
            For j = 0 To .count - 1
                If blockOf(XInv(.ids(j))) <> coneBlock Then reason = "coupled_cone": Exit Function
            Next j
            If seen(coneBlock) Then reason = "duplicate_cone": Exit Function
            seen(coneBlock) = True
        End With
    Next i
    For b = 0 To nb - 1
        If Not seen(b) Then reason = "missing_cone": Exit Function
    Next b
    r = 0
    For i = 0 To nf - 1
        If blockOf(i) < 0 Then retained(i) = r: fullPos(r) = i: r = r + 1
    Next i
    If retained(XInv(RLoadVariable)) < 0 Or r <> nr Then reason = "retained_layout": Exit Function
    Set uniqueReactions = CreateObject("Scripting.Dictionary")
    For Each reaction In RReactions
        id = CLng(reaction(0))
        If id < 0 Or id >= XN Or id = RLoadVariable Then reason = "reaction_id": Exit Function
        If retained(XInv(id)) < 0 Then reason = "reaction_eliminated": Exit Function
        uniqueReactions(id) = True
    Next reaction
    ReDim rows(0 To nr - 1)
    For r = 0 To nr - 1: Set rows(r) = RPX_Row(): Pattern rows, r, r: Next r
    For v = 0 To nf - 1
        If v Mod 512 = 0 Then Pulse "Schur mapping " & v
        For p = XKPtr(v) To XKPtr(v + 1) - 1
            u = XKId(p): bu = blockOf(u): bv = blockOf(v)
            If bu >= 0 And bv >= 0 Then
                If bu <> bv Then reason = "cross_block": Exit Function
                blocks(bu).hp(component(u), component(v)) = p
                blocks(bu).hp(component(v), component(u)) = p
            ElseIf bu < 0 And bv < 0 Then
                Pattern rows, retained(u), retained(v)
            Else
                If bu >= 0 Then b = bu: r = retained(v): id = v Else b = bv: r = retained(u): id = u
                If XPerm(id) < XN Then reason = "stress_q_coupling": Exit Function
                If Not neighbors(b).Exists(r) Then neighbors(b)(r) = neighbors(b).count
            End If
        Next p
    Next v
    For b = 0 To nb - 1
        With blocks(b)
            If b Mod 128 = 0 Then Pulse "Schur pattern " & b
            .n = neighbors(b).count
            estimate = estimate + 12# * .n * .n + 80# * .n
            If estimate > budget Then reason = "budget": Exit Function
            If .n > 0 Then
                ReDim .row(0 To .n - 1): ReDim .bp(0 To .n - 1, 0 To 2)
                ReDim .bv(0 To .n - 1, 0 To 2): ReDim .solved(0 To .n - 1, 0 To 2)
                ReDim .Slot(0 To .n - 1, 0 To .n - 1)
                For Each key In neighbors(b).keys
                    j = neighbors(b)(key): .row(j) = CLng(key)
                    For k = 0 To 2: .bp(j, k) = -1: Next k
                Next key
                For j = 0 To .n - 1
                    For k = 0 To j: Pattern rows, .row(j), .row(k): Next k
                Next j
            End If
        End With
    Next b
    RPX_Ordering nr, rows, perm, inv
    ReDim ordered(0 To nr - 1)
    For r = 0 To nr - 1: Set ordered(r) = RPX_Row(): Next r
    For r = 0 To nr - 1
        For Each key In rows(r).keys: Pattern ordered, inv(r), inv(CLng(key)): Next key
        RPX_ReleaseRow rows(r)
    Next r
    ReDim ap(0 To nr): p = 0
    For r = 0 To nr - 1
        ap(r) = p
        For Each key In ordered(r).keys: ordered(r)(key) = p: p = p + 1: Next key
    Next r
    ap(nr) = p: ReDim ai(0 To p - 1): ReDim av(0 To p - 1)
    For r = 0 To nr - 1
        For Each key In ordered(r).keys: ai(ordered(r)(key)) = CLng(key): Next key
    Next r
    ReDim directSlot(0 To UBound(XKVal))
    For p = 0 To UBound(directSlot): directSlot(p) = -1: Next p
    For v = 0 To nf - 1
        If v Mod 512 = 0 Then Pulse "Schur mapping " & v
        For p = XKPtr(v) To XKPtr(v + 1) - 1
            u = XKId(p): bu = blockOf(u): bv = blockOf(v)
            If bu < 0 And bv < 0 Then
                directSlot(p) = Slot(ordered, inv(retained(u)), inv(retained(v)))
            ElseIf (bu >= 0) Xor (bv >= 0) Then
                If bu >= 0 Then b = bu: r = retained(v): c = component(u) Else b = bv: r = retained(u): c = component(v)
                blocks(b).bp(neighbors(b)(r), c) = p
            End If
        Next p
    Next v
    For b = 0 To nb - 1
        With blocks(b)
            For j = 0 To .n - 1
                For k = 0 To j: .Slot(j, k) = Slot(ordered, inv(.row(j)), inv(.row(k))): Next k
            Next j
        End With
    Next b
    For r = 0 To nr - 1: RPX_ReleaseRow ordered(r): Next r
    RPX_SchurFactorBudget = budget - estimate: RPX_SchurBudgetRejected = False
    RPXC_Symbolic nr, ap, ai
    If RPX_SchurBudgetRejected Then reason = "factor_budget": Exit Function
    RPX_DiagEvent "schur_profile", "full=" & nf & ";reduced=" & nr & ";blocks=" & nb & ";retained_primal=" & nr - XM & ";reaction_unique=" & uniqueReactions.count & ";reaction_records=" & RReactions.count & ";load_id=" & RLoadVariable & ";nnz=" & ap(nr) & ";factor_nnz=" & RPXC_FactorNNZ
    ReDim z(0 To nb - 1, 0 To 2)
    P4Pattern = P4Pattern + 1: workspacePattern = P4Pattern
    P4Prepares = P4Prepares + 1
    Prepare = True
End Function

Private Sub LocalSolve(ByVal b As Long, ByRef x() As Double)
    Dim i As Long, j As Long
    With blocks(b)
        For i = 0 To 2
            For j = 0 To i - 1: x(i) = x(i) - .chol(i, j) * x(j): Next j
            x(i) = x(i) / .chol(i, i)
        Next i
        For i = 2 To 0 Step -1
            For j = i + 1 To 2: x(i) = x(i) - .chol(j, i) * x(j): Next j
            x(i) = x(i) / .chol(i, i)
        Next i
    End With
End Sub

Public Sub RPX_SchurFactor(ByVal regularization As Double)
    Dim reason As String, b As Long, i As Long, j As Long, k As Long, p As Long, value As Double
    Dim x(0 To 2) As Double, started As Double, phaseStart As Double
    reg = regularization: directionNo = 0: autoDirection = 0: factored = False: RPX_SchurAutoFailed = False
    If Not enabled Then Exit Sub
    started = RPX_DiagClock(): lastPulse = Timer
    If P4Cache And Not contextChecked Then
        phaseStart = RPX_DiagClock()
        If ready Then
            If RPX_SchurIdentityMatch() Then
                P4Hits = P4Hits + 1
                RPX_DiagEvent "p4_structure_hit", "bytes=" & estimate
            Else
                Call RPX_SchurRelease: enabled = True
                RPX_DiagEvent "p4_structure_miss", "reason=exact_identity"
            End If
        Else
            RPX_DiagEvent "p4_structure_miss", "reason=empty"
        End If
        contextChecked = True
        RPX_P4Time 8, phaseStart
    End If
    RPX_Stage "Schur factor"
    RPX_TestHook "schur_factor"
    If fault = "local" Then fault = "": reason = "test_local_pivot": GoTo rejected
    For p = 0 To UBound(XKVal)
        If Not RPX_SchurFinite(XKVal(p)) Then reason = "nonfinite_matrix": GoTo rejected
    Next p
    If Not ready Then
        phaseStart = RPX_DiagClock()
        If Not Prepare(reason) Then
            RPX_P4Time 0, phaseStart
            RPX_SchurRelease
            GoTo rejected
        End If
        RPX_P4Time 0, phaseStart
        ready = True
    End If
    phaseStart = RPX_DiagClock()
    For p = 0 To UBound(av): av(p) = 0#: Next p
    For p = 0 To UBound(directSlot)
        If directSlot(p) >= 0 Then av(directSlot(p)) = av(directSlot(p)) + XKVal(p)
    Next p
    For b = 0 To nb - 1
        If b Mod 128 = 0 Then Pulse "Schur local " & b
        With blocks(b)
            For i = 0 To 2
                For j = 0 To i
                    value = 0#: p = .hp(i, j)
                    If p >= 0 Then value = XKVal(p)
                    For k = 0 To j - 1: value = value - .chol(i, k) * .chol(j, k): Next k
                    If i = j Then
                        If Not RPX_SchurFinite(value) Or value <= 1E-20 Then reason = "local_pivot": GoTo rejected
                        .chol(i, j) = Sqr(value)
                    Else
                        .chol(i, j) = value / .chol(j, j)
                    End If
                Next j
            Next i
            P4LocalFactor = P4LocalFactor + .n
            For i = 0 To .n - 1
                For k = 0 To 2
                    x(k) = 0#: p = .bp(i, k)
                    If p >= 0 Then x(k) = XKVal(p)
                    .bv(i, k) = x(k)
                Next k
                LocalSolve b, x
                For k = 0 To 2
                    If Not RPX_SchurFinite(x(k)) Then reason = "nonfinite_local_solve": GoTo rejected
                    .solved(i, k) = x(k)
                Next k
            Next i
        End With
    Next b
    RPX_P4Time 1, phaseStart: phaseStart = RPX_DiagClock()
    For b = 0 To nb - 1
        If b Mod 128 = 0 Then Pulse "Schur update " & b
        With blocks(b)
            For i = 0 To .n - 1
                For j = 0 To i
                    value = 0#
                    For k = 0 To 2: value = value + .bv(i, k) * .solved(j, k): Next k
                    p = .Slot(i, j): av(p) = av(p) - value
                Next j
            Next i
        End With
    Next b
    RPX_P4Time 2, phaseStart: phaseStart = RPX_DiagClock()
    If fault = "factor" Then fault = "": reason = "test_schur_pivot": GoTo rejected
    For p = 0 To UBound(av)
        If Not RPX_SchurFinite(av(p)) Then reason = "nonfinite_schur": GoTo rejected
    Next p
    If Not RPXC_Numeric(ap, ai, av) Then reason = "schur_pivot": GoTo rejected
    RPX_P4Time 3, phaseStart
    factored = True: factorEpoch = P4Matrix
    RPX_DiagEvent "schur_numeric", "iteration=" & XIteration & ";delta=" & reg & ";seconds=" & RPX_DiagClock() - started
    Exit Sub
rejected:
    enabled = False: RPX_SchurRejected = RPX_SchurRejected + 1
    RPX_DiagEvent "linear_backend_fallback", "reason=" & reason & ";adopted=GENERIC;delta=" & reg
End Sub

Private Sub SolveReduced(ByRef rhs() As Double, ByRef answer() As Double)
    Dim small() As Double, result() As Double, x(0 To 2) As Double
    Dim b As Long, i As Long, k As Long, value As Double, started As Double
    Dim number As Long, description As String, source As String
    On Error GoTo Failed
    If solving Or Not factored Or factorEpoch <> P4Matrix Or workspacePattern <> P4Pattern Then err.Raise 5, "RPX_Schur", "SCHUR_NUMERIC_GENERATION"
    solving = True: P4Rhs = P4Rhs + 1: workspaceRhs = P4Rhs
    P4LocalRhs = P4LocalRhs + nb * IIf(P4ReuseU, 1#, 2#)
    P4Reduced = P4Reduced + 1: started = RPX_DiagClock()
    ReDim small(0 To nr - 1): ReDim answer(0 To nf - 1)
    For i = 0 To nr - 1: small(inv(i)) = rhs(fullPos(i)): Next i
    For b = 0 To nb - 1
        With blocks(b)
            For k = 0 To 2: x(k) = rhs(.pos(k)): Next k
            LocalSolve b, x
            If P4ReuseU Then
                For k = 0 To 2: z(b, k) = x(k): Next k
            End If
            For i = 0 To .n - 1
                value = 0#
                For k = 0 To 2: value = value + .bv(i, k) * x(k): Next k
                small(inv(.row(i))) = small(inv(.row(i))) - value
            Next i
        End With
    Next b
    RPX_P4Time 4, started: started = RPX_DiagClock()
    RPXC_Backsolve small, result
    RPX_P4Time 5, started: started = RPX_DiagClock()
    For i = 0 To nr - 1: answer(fullPos(i)) = result(inv(i)): Next i
    For b = 0 To nb - 1
        With blocks(b)
            If P4ReuseU Then
                If workspaceRhs <> P4Rhs Then err.Raise 5, , "SCHUR_RHS_GENERATION"
                For k = 0 To 2
                    x(k) = z(b, k)
                    For i = 0 To .n - 1: x(k) = x(k) - .solved(i, k) * answer(fullPos(.row(i))): Next i
                Next k
            Else
                For k = 0 To 2
                    x(k) = rhs(.pos(k))
                    For i = 0 To .n - 1: x(k) = x(k) - .bv(i, k) * answer(fullPos(.row(i))): Next i
                Next k
                LocalSolve b, x
            End If
            For k = 0 To 2: answer(.pos(k)) = x(k): Next k
        End With
    Next b
    RPX_P4Time 6, started
    solving = False
    Exit Sub
Failed:
    number = err.number: description = err.description: source = err.source
    solving = False: factored = False: factorEpoch = -1: workspaceRhs = -1
    err.Raise number, source, description
End Sub

Public Sub RPX_SchurCompare(ByRef rhs() As Double, ByRef generic() As Double)
    Dim savedMode As Boolean, number As Long, description As String, source As String
    savedMode = P4ReuseU
    On Error GoTo Failed
    CompareCore rhs, generic
    If P4Trace = 2 And factored And Not autoMode Then
        P4ReuseU = Not savedMode: directionNo = directionNo - 1
        CompareCore rhs, generic
        P4ReuseU = savedMode
    End If
    Exit Sub
Failed:
    number = err.number: description = err.description: source = err.source
    P4ReuseU = savedMode
    err.Raise number, source, description
End Sub
Private Sub CompareCore(ByRef rhs() As Double, ByRef generic() As Double)
    Dim answer() As Double, res() As Double, correction() As Double, i As Long, it As Long
    Dim err As Double, diff As Double, norm As Double, started As Double
    Dim rowNorm() As Double, genericRes() As Double, matrixNorm As Double, answerNorm As Double, rhsNorm As Double, denom As Double, backward As Double
    Dim genericError As Double, p As Long, j As Long, g As Long, groupDiff(0 To 2) As Double, groupNorm(0 To 2) As Double
    directionNo = directionNo + 1
    If Capture Then RPX_WriteKktCapture sequence, directionNo, reg, rhs, generic
    If Not factored Or autoMode Then Exit Sub
    started = Timer: SolveReduced rhs, answer
    For it = 1 To 3
        For i = 0 To nf - 1
            If Not RPX_SchurFinite(answer(i)) Then
                enabled = False: factored = False: RPX_SchurRejected = RPX_SchurRejected + 1
                RPX_DiagEvent "linear_backend_fallback", "reason=nonfinite_direction;adopted=GENERIC"
                Exit Sub
            End If
        Next i
        RPX_SymmetricProduct XKPtr, XKId, XKVal, answer, res: err = 0#
        For i = 0 To nf - 1
            res(i) = rhs(i) - res(i)
            If Abs(res(i)) > err Then err = Abs(res(i))
        Next i
        ' Use all three permitted full-system corrections for forward accuracy.
        SolveReduced res, correction
        For i = 0 To nf - 1: answer(i) = answer(i) + correction(i): Next i
    Next it
    RPX_SymmetricProduct XKPtr, XKId, XKVal, answer, res: err = 0#
    For i = 0 To nf - 1
        If Not RPX_SchurFinite(answer(i)) Or Not RPX_SchurFinite(res(i)) Then err = 1E+300
        If Abs(rhs(i) - res(i)) > err Then err = Abs(rhs(i) - res(i))
        If Abs(answer(i) - generic(i)) > diff Then diff = Abs(answer(i) - generic(i))
        If Abs(generic(i)) > norm Then norm = Abs(generic(i))
    Next i
    If norm < 1# Then norm = 1#
    ReDim rowNorm(0 To nf - 1)
    RPX_SymmetricProduct XKPtr, XKId, XKVal, generic, genericRes
    For j = 0 To nf - 1
        For p = XKPtr(j) To XKPtr(j + 1) - 1
            i = XKId(p): rowNorm(i) = rowNorm(i) + Abs(XKVal(p))
            If i <> j Then rowNorm(j) = rowNorm(j) + Abs(XKVal(p))
        Next p
    Next j
    For i = 0 To nf - 1
        If rowNorm(i) > matrixNorm Then matrixNorm = rowNorm(i)
        If Abs(answer(i)) > answerNorm Then answerNorm = Abs(answer(i))
        If Abs(rhs(i)) > rhsNorm Then rhsNorm = Abs(rhs(i))
        If Abs(rhs(i) - genericRes(i)) > genericError Then genericError = Abs(rhs(i) - genericRes(i))
        If blockOf(i) >= 0 Then g = 0 Else If XPerm(i) < XN Then g = 1 Else g = 2
        If Abs(answer(i) - generic(i)) > groupDiff(g) Then groupDiff(g) = Abs(answer(i) - generic(i))
        If Abs(generic(i)) > groupNorm(g) Then groupNorm(g) = Abs(generic(i))
    Next i
    For g = 0 To 2
        If groupNorm(g) < 1# Then groupNorm(g) = 1#
    Next g
    denom = matrixNorm * answerNorm + rhsNorm
    If denom > 0# Then backward = err / denom Else backward = err
    RPX_DiagEvent "kkt_error_groups", "recovery=" & IIf(P4ReuseU, "REUSE_U", "RESOLVE") & ";backward_error=" & backward & ";generic_residual=" & genericError & ";stress=" & groupDiff(0) / groupNorm(0) & ";retained=" & groupDiff(1) / groupNorm(1) & ";multipliers=" & groupDiff(2) / groupNorm(2)
    RPX_SchurComparisons = RPX_SchurComparisons + 1
    RPX_DiagEvent "kkt_comparison", "recovery=" & IIf(P4ReuseU, "REUSE_U", "RESOLVE") & ";iteration=" & XIteration & ";direction=" & directionNo & ";full_residual=" & err & ";relative_direction=" & diff / norm & ";seconds=" & Timer - started & ";adopted=GENERIC"
    If err > 0.00000000001 Or diff / norm > 0.00000001 Then
        RPX_SchurRejected = RPX_SchurRejected + 1
        RPX_DiagEvent "kkt_comparison_rejected", "reason=full_residual_or_direction;adopted=GENERIC"
    End If
End Sub

Public Function RPX_SchurAutoReady() As Boolean
    RPX_SchurAutoReady = autoMode And factored And factorEpoch = P4Matrix
End Function
Public Function RPX_SchurAutoSolve(ByRef rhs() As Double, ByRef answer() As Double, ByRef iterations As Long, ByRef residual As Double) As Boolean
    Dim res() As Double, correction() As Double, i As Long, step As Long, started As Double, correctionStart As Double
    If Not RPX_SchurAutoReady() Then Exit Function
    P4Attempts = P4Attempts + 1
    autoDirection = autoDirection + 1
    RPX_TestHook "schur_direction"
    SolveReduced rhs, answer
    iterations = 0
    For step = 1 To 3
        correctionStart = RPX_DiagClock(): started = correctionStart
        For i = 0 To nf - 1
            If Not RPX_SchurFinite(answer(i)) Then residual = 1E+300: GoTo AssessAuto
        Next i
        RPX_SymmetricProduct XKPtr, XKId, XKVal, answer, res: residual = 0#
        For i = 0 To nf - 1
            res(i) = rhs(i) - res(i)
            If Abs(res(i)) > residual Then residual = Abs(res(i))
        Next i
        P4Current.errors(step - 1) = residual
        RPX_P4Time 7, started: P4Current.checks = P4Current.checks + 1
        ' Same three-correction ceiling; final acceptance remains 1e-11.
        SolveReduced res, correction
        For i = 0 To nf - 1: answer(i) = answer(i) + correction(i): Next i
        iterations = iterations + 1: P4Current.corrections = iterations: P4Current.schurCorrections = iterations
        P4Corrections = P4Corrections + 1
        RPX_P4Time 9, correctionStart
    Next step
    started = RPX_DiagClock()
    RPX_SymmetricProduct XKPtr, XKId, XKVal, answer, res: residual = 0#
    For i = 0 To nf - 1
        If Not RPX_SchurFinite(answer(i)) Or Not RPX_SchurFinite(res(i)) Then residual = 1E+300: GoTo AssessAuto
        If Abs(rhs(i) - res(i)) > residual Then residual = Abs(rhs(i) - res(i))
    Next i
    P4Current.errors(3) = residual
    RPX_P4Time 7, started: P4Current.checks = P4Current.checks + 1
AssessAuto:
    If fault = "corrector" And autoDirection = 2 Then fault = "": residual = 1#
    If fault = "residual" Then fault = "": residual = 1#
    If RPX_SchurFinite(residual) And residual <= 0.00000000001 Then
        RPX_SchurAutoSolve = True
        RPX_DiagEvent "schur_direction", "iteration=" & XIteration & ";full_residual=" & residual & ";adopted=SCHUR"
    Else
        RPX_SchurAutoFailed = True: factored = False
        RPX_SchurRejected = RPX_SchurRejected + 1
        RPX_DiagEvent "linear_backend_fallback", "reason=full_residual;adopted=GENERIC;direction_kind=" & P4Current.kind
    End If
End Function

Public Sub RPX_SchurTestArm(ByVal value As String)
    If RPX_Busy Or CLng(RPX_Setting("TEST_INJECTION", 0)) <> 1 Then err.Raise 5, , "TEST_INJECTION_DISABLED"
    Select Case value
        Case "", "local", "factor", "corrector", "residual": fault = value
        Case Else: err.Raise 5, , "UNKNOWN_SCHUR_TEST_FAULT"
    End Select
End Sub

Private Sub Pulse(ByVal message As String)
    Dim now As Double
    If XCancel Then err.Raise 18, , "ANALYSIS_CANCELLED"
    now = Timer
    If now < lastPulse Or now - lastPulse >= 0.2 Then
        RPX_Stage message
        lastPulse = Timer
    End If
End Sub

Public Function RPX_SchurFinite(ByRef value As Double) As Boolean
    Dim number As Long, description As String, source As String
    On Error GoTo invalidNumber
    If value <> value Then Exit Function
    RPX_SchurFinite = (value >= -1E+290 And value <= 1E+290)
    Exit Function
invalidNumber:
    number = err.number: description = err.description: source = err.source
    ' VBA can raise overflow while comparing a quiet NaN. Only this scalar
    ' Double predicate treats that specific arithmetic result as nonfinite.
    If number <> 6 Then err.Raise number, source, description
End Function