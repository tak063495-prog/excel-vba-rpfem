Option Explicit
Public AHORoundingBacktracks As Long
' Independent normalized AHO-HSD, fixed lambda=1 / zero objective / SOC3 only.
' D=L(s)^-1 L(z) is nonsymmetric: FULL CSC and row-pivoted LU, never NT/LDL.
' A returned witness is NOT an Fs root or an infeasibility certificate.
Private Type AhoDirection
    dx() As Double
    dy() As Double
    dz() As Double
    ds() As Double
    dt As Double
    dk As Double
    dtheta As Double
End Type
Private bp() As Long, bi() As Long, bv() As Double, br() As Double
Private ahoOrder() As Long, ahoInverse() As Long
Private above() As Long, below() As Long, bdiag() As Long, bScale() As Double
Private blocks() As Double, dh() As Double, dr() As Double
Private px() As Double, py() As Double, ps() As Double, pz() As Double
Private rx() As Double, req() As Double, rs0() As Double, offsetH() As Double
Private rd() As Double, re() As Double, rs() As Double
Private v1() As Double, v2() As Double, w1() As Double, w2() As Double
Private b1() As Double, b2() As Double
Private tau As Double, kap As Double, theta As Double, rt As Double, nu As Double
Private rk As Double, rn As Double, sch00 As Double, sch01 As Double, sch10 As Double, sch11 As Double
Private ahoPhase As String
Private factorEpoch As Long, matrixSize As Long, budget As Double
Public AHOStatus As String, AHOIterations As Long, AHOLinearResidual As Double, AHOComponentwiseResidual As Double
Public AHOActive As Boolean
Private ahoStart As Double, ahoSeconds As Double
Public Sub RPX_AHOCheckBudget()
    Dim elapsed As Double
    If Not AHOActive Then Exit Sub
    elapsed = Timer - ahoStart: If elapsed < 0# Then elapsed = elapsed + 86400#
    If elapsed > ahoSeconds Then err.Raise 5, "RPX_AHO3", "AHO_USER_BUDGET_LIMIT"
End Sub
Private Function Dot(ByRef a() As Double, ByRef b() As Double) As Double
    Dim i As Long, correction As Double, term As Double, nextValue As Double
    For i = 0 To UBound(b)
        term = a(i) * b(i) - correction: nextValue = Dot + term
        correction = (nextValue - Dot) - term: Dot = nextValue
    Next i
End Function
Private Function largest(ByRef a() As Double) As Double
    Dim i As Long
    For i = 0 To UBound(a)
        If Not RPX_SchurFinite(a(i)) Then err.Raise 5, "RPX_AHO3", "AHO_NONFINITE"
        If Abs(a(i)) > largest Then largest = Abs(a(i))
    Next i
End Function
Private Sub ProductA(ByRef x() As Double, ByRef answer() As Double)
    Dim k As Long, r As Long, j As Long, first As Long
    ReDim answer(0 To XS - 1)
    For k = 0 To XB - 1
        first = XC(k).first
        For r = 0 To 2
            For j = 0 To XC(k).count - 1
                answer(first + r) = answer(first + r) - XC(k).coef(r, j) * x(XC(k).ids(j))
            Next j
        Next r
    Next k
End Sub
Private Sub ProductAT(ByRef x() As Double, ByRef answer() As Double)
    Dim k As Long, r As Long, j As Long, id As Long, first As Long
    ReDim answer(0 To XN - 1)
    For k = 0 To XB - 1
        first = XC(k).first
        For j = 0 To XC(k).count - 1
            id = XC(k).ids(j)
            For r = 0 To 2: answer(id) = answer(id) - XC(k).coef(r, j) * x(first + r): Next r
        Next j
    Next k
End Sub
Private Sub ProductD(ByRef x() As Double, ByRef answer() As Double, Optional ByVal transpose As Boolean = False)
    Dim k As Long, r As Long, c As Long, first As Long
    ReDim answer(0 To XS - 1)
    For k = 0 To XB - 1
        first = XC(k).first
        For r = 0 To 2
            For c = 0 To 2
                If transpose Then
                    answer(first + r) = answer(first + r) + blocks(k, c, r) * x(first + c)
                Else
                    answer(first + r) = answer(first + r) + blocks(k, r, c) * x(first + c)
                End If
            Next c
        Next r
    Next k
End Sub
Private Sub Jordan(ByRef a() As Double, ByRef b() As Double, ByRef answer() As Double)
    Dim k As Long, p As Long
    ReDim answer(0 To XS - 1)
    For k = 0 To XB - 1
        p = XC(k).first
        answer(p) = a(p) * b(p) + a(p + 1) * b(p + 1) + a(p + 2) * b(p + 2)
        answer(p + 1) = a(p) * b(p + 1) + b(p) * a(p + 1)
        answer(p + 2) = a(p) * b(p + 2) + b(p) * a(p + 2)
    Next k
End Sub
Private Sub SolveJordan(ByVal first As Long, ByVal rhs0 As Double, ByVal rhs1 As Double, ByVal rhs2 As Double, ByRef answer0 As Double, ByRef answer1 As Double, ByRef answer2 As Double)
    Dim s0 As Double, s1 As Double, s2 As Double, determinant As Double
    s0 = ps(first): s1 = ps(first + 1): s2 = ps(first + 2)
    determinant = (s0 - Sqr(s1 * s1 + s2 * s2)) * (s0 + Sqr(s1 * s1 + s2 * s2))
    If s0 <= 0# Or determinant <= 0# Then err.Raise 5, "RPX_AHO3", "AHO_NONINTERIOR"
    answer0 = (s0 * rhs0 - s1 * rhs1 - s2 * rhs2) / determinant
    answer1 = (rhs1 - s1 * answer0) / s0: answer2 = (rhs2 - s2 * answer0) / s0
End Sub
Private Sub Pattern()
    Dim c As Long, p As Long, r As Long, cursor() As Long, q As Long, rows() As Object
    matrixSize = XN + XM
    ReDim bp(0 To matrixSize): ReDim cursor(0 To matrixSize - 1)
    ReDim above(0 To UBound(XKId)): ReDim below(0 To UBound(XKId)): ReDim bdiag(0 To matrixSize - 1)
    For c = 0 To matrixSize - 1
        For p = XKPtr(c) To XKPtr(c + 1) - 1
            r = XKId(p): bp(c + 1) = bp(c + 1) + 1
            If r <> c Then bp(r + 1) = bp(r + 1) + 1
        Next p
    Next c
    For c = 0 To matrixSize - 1: bp(c + 1) = bp(c + 1) + bp(c): cursor(c) = bp(c): Next c
    ReDim bi(0 To bp(matrixSize) - 1): ReDim bv(0 To bp(matrixSize) - 1): ReDim br(0 To bp(matrixSize) - 1)
    For c = 0 To matrixSize - 1
        For p = XKPtr(c) To XKPtr(c + 1) - 1
            r = XKId(p): q = cursor(c): cursor(c) = q + 1: bi(q) = r: above(p) = q
            If r <> c Then
                q = cursor(r): cursor(r) = q + 1: bi(q) = c: below(p) = q
            Else
                below(p) = q: bdiag(c) = q
            End If
        Next p
    Next c
    ReDim rows(0 To matrixSize - 1)
    For r = 0 To matrixSize - 1: Set rows(r) = RPX_Row(): Next r
    For c = 0 To matrixSize - 1
        For p = bp(c) To bp(c + 1) - 1: rows(bi(p))(c) = True: Next p
    Next c
    RPX_Stage "AHO column ordering"
    RPX_ColumnOrdering matrixSize, rows, ahoOrder, ahoInverse
    For r = 0 To matrixSize - 1: RPX_ReleaseRow rows(r): Next r
End Sub
Private Sub matrix()
    Dim k As Long, p As Long, j As Long, l As Long, r As Long, c As Long, Slot As Long, pat As Long, first As Long
    Dim lhs0 As Double, lhs1 As Double, lhs2 As Double, value As Double, q0 As Double, q1 As Double, q2 As Double
    Dim pass As Long, norm() As Double, nextScale() As Double
    For p = 0 To UBound(bv): bv(p) = 0#: Next p
    For p = 0 To UBound(XKBase)
        bv(above(p)) = XKBase(p): bv(below(p)) = XKBase(p)
    Next p
    ReDim blocks(0 To XB - 1, 0 To 2, 0 To 2)
    For k = 0 To XB - 1
        first = XC(k).first: pat = XC(k).Pattern
        For c = 0 To 2
            lhs0 = pz(first + c): lhs1 = 0#: lhs2 = 0#
            If c = 0 Then lhs1 = pz(first + 1): lhs2 = pz(first + 2)
            If c = 1 Then lhs1 = pz(first)
            If c = 2 Then lhs2 = pz(first)
            SolveJordan first, lhs0, lhs1, lhs2, q0, q1, q2
            blocks(k, 0, c) = q0: blocks(k, 1, c) = q1: blocks(k, 2, c) = q2
        Next c
        For j = 0 To XC(k).count - 1
            r = XInv(XC(k).ids(j))
            For l = 0 To XC(k).count - 1
                c = XInv(XC(k).ids(l))
                lhs0 = blocks(k, 0, 0) * XC(k).coef(0, l) + blocks(k, 0, 1) * XC(k).coef(1, l) + blocks(k, 0, 2) * XC(k).coef(2, l)
                lhs1 = blocks(k, 1, 0) * XC(k).coef(0, l) + blocks(k, 1, 1) * XC(k).coef(1, l) + blocks(k, 1, 2) * XC(k).coef(2, l)
                lhs2 = blocks(k, 2, 0) * XC(k).coef(0, l) + blocks(k, 2, 1) * XC(k).coef(1, l) + blocks(k, 2, 2) * XC(k).coef(2, l)
                value = XC(k).coef(0, j) * lhs0 + XC(k).coef(1, j) * lhs1 + XC(k).coef(2, j) * lhs2
                If j >= l Then Slot = XC(pat).slots(j, l) Else Slot = XC(pat).slots(l, j)
                If r <= c Then p = above(Slot) Else p = below(Slot)
                bv(p) = bv(p) + value
            Next l
        Next j
    Next k
    ReDim bScale(0 To matrixSize - 1)
    For r = 0 To matrixSize - 1: bScale(r) = 1#: Next r
    For pass = 1 To 4
        ReDim norm(0 To matrixSize - 1): ReDim nextScale(0 To matrixSize - 1)
        For c = 0 To matrixSize - 1
            For p = bp(c) To bp(c + 1) - 1
                r = bi(p): value = Abs(bv(p)): If value > norm(r) Then norm(r) = value
            Next p
        Next c
        For r = 0 To matrixSize - 1
            value = norm(r): If value < 0.000000000001 Then value = 0.000000000001
            If value > 1000000000000# Then value = 1000000000000#
            nextScale(r) = 1# / Sqr(value): bScale(r) = bScale(r) * nextScale(r)
        Next r
        For c = 0 To matrixSize - 1
            For p = bp(c) To bp(c + 1) - 1: bv(p) = bv(p) * nextScale(bi(p)) * nextScale(c): Next p
        Next c
    Next pass
    br = bv
    For r = 0 To matrixSize - 1
        If XPerm(r) < XN Then br(bdiag(r)) = br(bdiag(r)) + 0.0000000000001 Else br(bdiag(r)) = br(bdiag(r)) - 0.0000000000001
    Next r
    ahoPhase = "matrix_factor"
    factorEpoch = factorEpoch + 1
    If Not RPX_HLUFactor(bp, bi, br, budget, 64# * bp(matrixSize) + 1024# * matrixSize, factorEpoch, False, False, False, 0.1, ahoOrder, True) Then err.Raise 5, "RPX_AHO3", "AHO_SINGULAR_FACTOR"
End Sub
Private Sub Backsolve(ByRef rhs() As Double, ByRef answer() As Double)
    Dim scaledRhs() As Double, solved() As Double, product() As Double, correction() As Double, residual() As Double
    Dim compensation() As Double, denominator() As Double, p As Long, r As Long, c As Long, pass As Long, term As Double, nextValue As Double, errMax As Double, ratio As Double
    ReDim scaledRhs(0 To matrixSize - 1)
    For r = 0 To matrixSize - 1: scaledRhs(r) = bScale(r) * rhs(XPerm(r)): Next r
    RPX_HLUSolve scaledRhs, solved, factorEpoch
    For pass = 0 To 12
        ReDim product(0 To matrixSize - 1): ReDim compensation(0 To matrixSize - 1): ReDim residual(0 To matrixSize - 1): ReDim denominator(0 To matrixSize - 1)
        For c = 0 To matrixSize - 1
            For p = bp(c) To bp(c + 1) - 1
                r = bi(p): term = bv(p) * solved(c) - compensation(r): nextValue = product(r) + term
                compensation(r) = (nextValue - product(r)) - term: product(r) = nextValue
                denominator(r) = denominator(r) + Abs(bv(p) * solved(c))
            Next p
        Next c
        For r = 0 To matrixSize - 1: residual(r) = scaledRhs(r) - product(r): Next r
        errMax = largest(residual)
        If errMax <= 0.00000000000002 * (1# + largest(scaledRhs)) Then Exit For
        If pass = 12 Then Exit For
        RPX_HLUSolve residual, correction, factorEpoch
        For r = 0 To matrixSize - 1: solved(r) = solved(r) + correction(r): Next r
    Next pass
    AHOLinearResidual = errMax / (1# + largest(scaledRhs))
    AHOComponentwiseResidual = 0#
    For r = 0 To matrixSize - 1
        ratio = Abs(residual(r)) / (1E-30 + Abs(scaledRhs(r)) + denominator(r))
        If ratio > AHOComponentwiseResidual Then AHOComponentwiseResidual = ratio
    Next r
    If AHOLinearResidual > 0.0000001 Or AHOComponentwiseResidual > 0.0000001 Then err.Raise 5, "RPX_AHO3", "AHO_ORIGINAL_LINEAR_RESIDUAL"
    ReDim answer(0 To matrixSize - 1)
    For r = 0 To matrixSize - 1: answer(XPerm(r)) = bScale(r) * solved(r): Next r
End Sub
Private Sub Border()
    Dim k As Long, p As Long, a() As Double, b() As Double, c() As Double, d() As Double
    Dim c00 As Double, c01 As Double, c10 As Double, c11 As Double
    ahoPhase = "border_D"
    ProductD offsetH, dh: ProductD rs0, dr
    ProductAT dh, a: ProductAT dr, b
    ProductD offsetH, c, True: ProductAT c, d
    ReDim v1(0 To matrixSize - 1): ReDim v2(0 To matrixSize - 1)
    ReDim w1(0 To matrixSize - 1): ReDim w2(0 To matrixSize - 1)
    For k = 0 To XN - 1: v1(k) = -a(k): v2(k) = rx(k) - b(k): w1(k) = -tau * d(k): Next k
    ProductD rs0, c, True: ProductAT c, d
    For k = 0 To XN - 1: w2(k) = rx(k) + d(k): Next k
    For k = 0 To XM - 1
        p = XN + k: v1(p) = -XERhs(k): v2(p) = -req(k)
        w1(p) = -tau * XERhs(k): w2(p) = req(k)
    Next k
    c00 = kap + tau * Dot(offsetH, dh): c01 = tau * (Dot(offsetH, dr) + rt)
    c10 = rt - Dot(rs0, dh): c11 = -Dot(rs0, dr)
    ahoPhase = "border_backsolve"
    Backsolve v1, b1: Backsolve v2, b2
    sch00 = c00 - Dot(w1, b1): sch01 = c01 - Dot(w1, b2)
    sch10 = c10 - Dot(w2, b1): sch11 = c11 - Dot(w2, b2)
End Sub
Private Sub direction(ByRef rc() As Double, ByVal rcTau As Double, ByRef out As AhoDirection)
    Dim jordanZR() As Double, u() As Double, atU() As Double, rhs() As Double, solved() As Double, aDx() As Double, daDx() As Double
    Dim k As Long, p As Long, rhs0 As Double, rhs1 As Double, determinant As Double
    Jordan pz, rs, jordanZR: ReDim u(0 To XS - 1)
    For k = 0 To XB - 1
        p = XC(k).first
        SolveJordan p, rc(p) - jordanZR(p), rc(p + 1) - jordanZR(p + 1), rc(p + 2) - jordanZR(p + 2), u(p), u(p + 1), u(p + 2)
    Next k
    ProductAT u, atU: ReDim rhs(0 To matrixSize - 1)
    For k = 0 To XN - 1: rhs(k) = -rd(k) - atU(k): Next k
    For k = 0 To XM - 1: rhs(XN + k) = re(k): Next k
    Backsolve rhs, solved
    rhs0 = rcTau + tau * Dot(offsetH, u) - tau * rk - Dot(w1, solved)
    rhs1 = -rn - Dot(rs0, u) - Dot(w2, solved)
    determinant = sch00 * sch11 - sch01 * sch10
    If determinant = 0# Then err.Raise 5, "RPX_AHO3", "AHO_SINGULAR_BORDER"
    out.dt = (rhs0 * sch11 - sch01 * rhs1) / determinant
    out.dtheta = (sch00 * rhs1 - rhs0 * sch10) / determinant
    ReDim out.dx(0 To XN - 1): ReDim out.dy(0 To XM - 1)
    For k = 0 To XN - 1: out.dx(k) = solved(k) - b1(k) * out.dt - b2(k) * out.dtheta: Next k
    For k = 0 To XM - 1
        p = XN + k: out.dy(k) = solved(p) - b1(p) * out.dt - b2(p) * out.dtheta
    Next k
    ProductA out.dx, aDx: ProductD aDx, daDx
    ReDim out.ds(0 To XS - 1): ReDim out.dz(0 To XS - 1)
    For k = 0 To XS - 1
        out.ds(k) = rs(k) - aDx(k) + offsetH(k) * out.dt + rs0(k) * out.dtheta
        out.dz(k) = daDx(k) - dh(k) * out.dt - dr(k) * out.dtheta + u(k)
    Next k
    out.dk = -Dot(XERhs, out.dy) - Dot(offsetH, out.dz) + rt * out.dtheta + rk
End Sub
Private Function stepSize(ByRef direction As AhoDirection) As Double
    Dim hit As Double, hitDual As Double
    If Not ConeBoundaryHit(ps, direction.ds, hit) Then err.Raise 5, "RPX_AHO3", "AHO_NONINTERIOR"
    If Not ConeBoundaryHit(pz, direction.dz, hitDual) Then err.Raise 5, "RPX_AHO3", "AHO_NONINTERIOR"
    stepSize = 1#: If hit < stepSize Then stepSize = hit
    If hitDual < stepSize Then stepSize = hitDual
    If direction.dt < 0# Then If -tau / direction.dt < stepSize Then stepSize = -tau / direction.dt
    If direction.dk < 0# Then If -kap / direction.dk < stepSize Then stepSize = -kap / direction.dk
End Function
Private Function TrialInterior(ByRef direction As AhoDirection, ByVal alpha As Double) As Boolean
    Dim k As Long, p As Long, side As Long, head As Double, tail1 As Double, tail2 As Double, determinant As Double
    If tau + alpha * direction.dt <= 0# Or kap + alpha * direction.dk <= 0# Or theta + alpha * direction.dtheta <= 0# Then Exit Function
    For k = 0 To XB - 1
        p = XC(k).first
        For side = 0 To 1
            If side = 0 Then
                head = ps(p) + alpha * direction.ds(p): tail1 = ps(p + 1) + alpha * direction.ds(p + 1): tail2 = ps(p + 2) + alpha * direction.ds(p + 2)
            Else
                head = pz(p) + alpha * direction.dz(p): tail1 = pz(p + 1) + alpha * direction.dz(p + 1): tail2 = pz(p + 2) + alpha * direction.dz(p + 2)
            End If
            If head <= 0# Then Exit Function
            determinant = head * head - tail1 * tail1 - tail2 * tail2
            If determinant <= 0# Or Not RPX_SchurFinite(determinant) Then Exit Function
        Next side
    Next k
    TrialInterior = True
End Function
Public Function RPX_AHOFixedWitness(Optional ByVal seconds As Double = 120#, Optional ByVal maxIterations As Long = 80) As Boolean
    Dim k As Long, p As Long, j As Long, first As Long, iteration As Long, start As Double, elapsed As Double
    Dim ax() As Double, atZ() As Double, e() As Double, rc() As Double, jsz() As Double, cross() As Double
    Dim aff As AhoDirection, Corrector As AhoDirection, candidate As RPX_HsdCandidateState
    Dim mu As Double, muAff As Double, sigma As Double, alpha As Double, point() As Double
    Dim number As Long, source As String, message As String, backtrack As Long
    On Error GoTo Failed
    AHOStatus = "NUMERICAL_UNKNOWN": AHOIterations = 0: AHORoundingBacktracks = 0
    ahoStart = Timer: ahoSeconds = seconds: AHOActive = True
    If Not AFActive Then err.Raise 5, "RPX_AHO3", "AHO_ORIGINAL_MAP_REQUIRED"
    For k = 0 To XN - 1: If XQ(k) <> 0# Then err.Raise 5, "RPX_AHO3", "AHO_ZERO_OBJECTIVE_REQUIRED"
    Next k
    For k = 0 To XB - 1: If XC(k).dimn <> 3 Then err.Raise 5, "RPX_AHO3", "AHO_SOC3_REQUIRED"
    Next k
    If XM < 1 Or seconds <= 0# Or maxIterations < 1 Then err.Raise 5, "RPX_AHO3", "AHO_CONTEXT_INVALID"
    budget = CDbl(RPX_Setting("HSD_STABLE_MAX_BYTES", 268435456#))
    Pattern
    ReDim px(0 To XN - 1): ReDim py(0 To XM - 1): ReDim ps(0 To XS - 1): ReDim pz(0 To XS - 1)
    ReDim e(0 To XS - 1): ReDim offsetH(0 To XS - 1): ReDim rs0(0 To XS - 1)
    For k = 0 To XB - 1
        first = XC(k).first: ps(first) = 1#: pz(first) = 1#: e(first) = 1#
        For j = 0 To 2: offsetH(first + j) = XC(k).offset(j): rs0(first + j) = e(first + j) - offsetH(first + j): Next j
    Next k
    ProductAT e, rx
    For k = 0 To XN - 1: rx(k) = -rx(k): Next k
    ReDim req(0 To XM - 1): For k = 0 To XM - 1: req(k) = -XERhs(k): Next k
    tau = 1#: kap = 1#: theta = 1#: rt = 1# + Dot(offsetH, e): nu = XB + 1#: start = Timer
    For iteration = 0 To maxIterations - 1
        AHOIterations = iteration: XIteration = iteration
        RPX_Stage "AHO fixed witness " & iteration
        RPX_AssertInputs "aho_fixed_iteration"
        ahoPhase = "residuals"
        ProductA px, ax: ProductAT pz, atZ: rd = atZ
        ReDim re(0 To XM - 1): ReDim rs(0 To XS - 1)
        For k = 0 To XM - 1
            For p = XEPtr(k) To XEPtr(k + 1) - 1
                rd(XEId(p)) = rd(XEId(p)) + XEVal(p) * py(k)
                re(k) = re(k) - XEVal(p) * px(XEId(p))
            Next p
            re(k) = re(k) + XERhs(k) * tau + req(k) * theta
        Next k
        For k = 0 To XN - 1: rd(k) = rd(k) + rx(k) * theta: Next k
        For k = 0 To XS - 1: rs(k) = -ax(k) + offsetH(k) * tau + rs0(k) * theta - ps(k): Next k
        rk = -Dot(XERhs, py) - Dot(offsetH, pz) + rt * theta - kap
        rn = Dot(rx, px) + Dot(req, py) + Dot(rs0, pz) + rt * tau - nu
        mu = (Dot(ps, pz) + tau * kap) / nu
        If Not RPX_SchurFinite(mu) Or mu <= 0# Or tau <= 0# Or kap <= 0# Then err.Raise 5, "RPX_AHO3", "AHO_NONINTERIOR"
        ReDim point(0 To XN - 1)
        For k = 0 To XN - 1: point(k) = px(k) / tau: Next k
        If mu / (tau * tau) < 0.0000000001 And kap / tau < 0.0000001 And largest(rd) <= 0.0000001 And largest(re) <= 0.0000001 And largest(rs) <= 0.0000001 And Abs(rk) <= 0.0000001 And Abs(rn) <= 0.0000001 * (1# + nu) Then
            candidate.x = point
            If RPX_AffineCandidate(candidate) Then
                xx = point: XStatus = "FIXED_FS_WITNESS": AHOStatus = "FIXED_FS_WITNESS"
                RPX_AHOFixedWitness = True: Exit For
            End If
        End If
        elapsed = Timer - start: If elapsed < 0# Then elapsed = elapsed + 86400#
        If elapsed > seconds Then AHOStatus = "NUMERICAL_UNKNOWN_TIME_BUDGET": Exit For
        ahoPhase = "matrix": matrix
        ahoPhase = "border": Border
        Jordan ps, pz, jsz: ReDim rc(0 To XS - 1)
        For k = 0 To XS - 1: rc(k) = -jsz(k): Next k
        ahoPhase = "predictor"
        direction rc, -tau * kap, aff
        alpha = stepSize(aff): muAff = 0#
        For k = 0 To XS - 1: muAff = muAff + (ps(k) + alpha * aff.ds(k)) * (pz(k) + alpha * aff.dz(k)): Next k
        muAff = (muAff + (tau + alpha * aff.dt) * (kap + alpha * aff.dk)) / nu
        sigma = muAff / mu: If sigma < 0# Then sigma = 0#
        sigma = sigma * sigma * sigma: If sigma > 0.95 Then sigma = 0.95
        Jordan aff.ds, aff.dz, cross
        For k = 0 To XS - 1: rc(k) = sigma * mu * e(k) - jsz(k) - cross(k): Next k
        ahoPhase = "corrector"
        direction rc, sigma * mu - tau * kap - aff.dt * aff.dk, Corrector
        alpha = 0.99 * stepSize(Corrector)
        If Corrector.dtheta < 0# Then If -0.99 * theta / Corrector.dtheta < alpha Then alpha = -0.99 * theta / Corrector.dtheta
        For backtrack = 0 To 31
            If TrialInterior(Corrector, alpha) Then Exit For
            alpha = alpha / 2#: AHORoundingBacktracks = AHORoundingBacktracks + 1
        Next backtrack
        If backtrack >= 32 Then alpha = 0#
        If alpha < 0.000000000001 Then AHOStatus = "NUMERICAL_UNKNOWN_STEP_STALLED": Exit For
        For k = 0 To XN - 1: px(k) = px(k) + alpha * Corrector.dx(k): Next k
        For k = 0 To XM - 1: py(k) = py(k) + alpha * Corrector.dy(k): Next k
        For k = 0 To XS - 1: ps(k) = ps(k) + alpha * Corrector.ds(k): pz(k) = pz(k) + alpha * Corrector.dz(k): Next k
        tau = tau + alpha * Corrector.dt: kap = kap + alpha * Corrector.dk: theta = theta + alpha * Corrector.dtheta
        RPX_DiagEvent "aho_iteration", "iteration=" & iteration & ";mu=" & mu & ";tau=" & tau & ";linear_unshifted=" & AHOLinearResidual & ";componentwise=" & AHOComponentwiseResidual
        RPX_HLUPark
    Next iteration
    Call RPX_HLURelease
    AHOActive = False
    Exit Function
Failed:
    number = err.number: source = err.source: message = err.description
    Call RPX_HLURelease
    AHOActive = False
    ' Only explicit algorithmic failures are unknown; real VBA errors always propagate.
    If number = 5 Then
        Select Case message
        Case "AHO_USER_BUDGET_LIMIT", "AHO_NONINTERIOR", "AHO_NONFINITE", "AHO_SINGULAR_FACTOR", "AHO_SINGULAR_BORDER", "AHO_ORIGINAL_LINEAR_RESIDUAL", "HSD_STABLE_UNAVAILABLE_MEMORY", "HSD_NONFINITE_FACTOR", "HSD_NONFINITE_SOLUTION"
            AHOStatus = "NUMERICAL_UNKNOWN:" & message & ";phase=" & ahoPhase & ";iteration=" & AHOIterations: Exit Function
        End Select
    End If
    err.Raise number, source, message & ";phase=" & ahoPhase
End Function