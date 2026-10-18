Attribute VB_Name = "RPX_Seep"
Option Explicit
' RPX_Seep : steady 2D seepage, linear T3, anisotropic kx/ky per material.
' Unconfined free surface: fixed mesh, kr reduction where pressure head < 0,
' relaxed Picard iteration, seepage-face (Signorini) switching.
' Standalone. Self-contained skyline Cholesky with RCM ordering.
' Python mirror: seep_check.py (same algorithm, verified against Dupuit).
' BC types: 0 = none/impermeable, 1 = fixed head, 2 = seepage face, 3 = nodal flux (inflow +)

Public SN As Long, SM As Long, SNMat As Long
Public sxy() As Double, STri() As Long, SMat() As Long
Public SKx() As Double, SKy() As Double
Public SBcType() As Long, SBcVal() As Double
Public SHead() As Double, SPres() As Double, SFlux() As Double, SKr() As Double
Public SActive() As Boolean
Public SKrMin As Double, SBeta As Double, SRelax As Double, STol As Double, SMaxIt As Long
Public SIterations As Long, SConverged As Boolean, SLastDelta As Double
' #6: convergence = unrelaxed fixed-point change AND scaled free-node flux balance (independent of SRelax)
Public SResTol As Double, SLastResidual As Double, SLastFixDelta As Double
' #8: SLastResidual is residual / mixed allowance; acceptance threshold is 1.
Public SLastRawResidual As Double, SLastBoundaryScale As Double, SLastRoundoffMax As Double
Public SLastRoundoffDominated As Boolean
Public SGammaW As Double

Private sPerm() As Long, sFc() As Long, sPtr() As Long
Private seepWriteSequence As Long

Public Sub RPX_SeepCheckCancelled()
    If Not XCancel Then Exit Sub
    SConverged = False
    err.Raise 18, "RPX_Seep", "SEEP_ANALYSIS_CANCELLED 浸透解析を中止しました。"
End Sub

Public Sub RPX_SeepDefaults()
    SKrMin = 0.0001: SBeta = 20#: SRelax = 0.5: STol = 0.000001: SMaxIt = 300: SGammaW = 9.81: SResTol = 0.000001
End Sub

Private Sub TriGeom(ByVal e As Long, ByRef area2 As Double, ByRef bb() As Double, ByRef cc() As Double)
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double
    x1 = sxy(STri(e, 0), 0): y1 = sxy(STri(e, 0), 1)
    x2 = sxy(STri(e, 1), 0): y2 = sxy(STri(e, 1), 1)
    x3 = sxy(STri(e, 2), 0): y3 = sxy(STri(e, 2), 1)
    area2 = (x2 - x1) * (y3 - y1) - (x3 - x1) * (y2 - y1)
    If area2 <= 0# Then err.Raise 5, , "SEEP_TRIANGLE_NOT_CCW:" & e
    bb(0) = y2 - y3: bb(1) = y3 - y1: bb(2) = y1 - y2
    cc(0) = x3 - x2: cc(1) = x1 - x3: cc(2) = x2 - x1
End Sub

Private Function ElemK(ByVal e As Long, ByVal i As Long, ByVal j As Long, ByVal s As Double, ByRef bb() As Double, ByRef cc() As Double) As Double
    ElemK = s * (SKx(SMat(e)) * bb(i) * bb(j) + SKy(SMat(e)) * cc(i) * cc(j))
End Function

' Reverse Cuthill-McKee; sPerm(node) = equation number
Private Sub BuildOrdering()
    Dim adj() As Object, deg() As Long, seen() As Boolean, order() As Long, q() As Long
    Dim e As Long, i As Long, j As Long, cnt As Long, head As Long, tail As Long, s As Long, v As Long
    Dim nb() As Long, nn As Long, a As Long, b As Long, t As Long, key As Variant
    ReDim adj(0 To SN - 1): ReDim deg(0 To SN - 1): ReDim seen(0 To SN - 1): ReDim order(0 To SN - 1): ReDim q(0 To SN - 1)
    For i = 0 To SN - 1: Set adj(i) = CreateObject("Scripting.Dictionary"): Next i
    For e = 0 To SM - 1
        For i = 0 To 2
            For j = 0 To 2
                If i <> j Then
                    If Not adj(STri(e, i)).Exists(STri(e, j)) Then adj(STri(e, i)).Add STri(e, j), True
                End If
            Next j
        Next i
    Next e
    For i = 0 To SN - 1: deg(i) = adj(i).count: Next i
    cnt = 0
    Do While cnt < SN
        s = -1
        For i = 0 To SN - 1
            If Not seen(i) Then
                If s < 0 Then
                    s = i
                ElseIf deg(i) < deg(s) Then
                    s = i
                End If
            End If
        Next i
        seen(s) = True: head = 0: tail = 0: q(tail) = s: tail = tail + 1
        Do While head < tail
            v = q(head): head = head + 1
            nn = adj(v).count: ReDim nb(0 To nn): j = 0
            For Each key In adj(v).keys: nb(j) = key: j = j + 1: Next key
            ' insertion sort by degree (stable, ties by insertion = ascending node id after sort below)
            For a = 1 To nn - 1
                t = nb(a): b = a - 1
                Do While b >= 0
                    If deg(nb(b)) > deg(t) Or (deg(nb(b)) = deg(t) And nb(b) > t) Then nb(b + 1) = nb(b): b = b - 1 Else Exit Do
                Loop
                nb(b + 1) = t
            Next a
            For a = 0 To nn - 1
                If Not seen(nb(a)) Then seen(nb(a)) = True: q(tail) = nb(a): tail = tail + 1
            Next a
        Loop
        For i = 0 To tail - 1: order(cnt) = q(i): cnt = cnt + 1: Next i
    Loop
    ReDim sPerm(0 To SN - 1)
    For i = 0 To SN - 1: sPerm(order(SN - 1 - i)) = i: Next i
    ReDim sFc(0 To SN - 1): ReDim sPtr(0 To SN)
    For i = 0 To SN - 1: sFc(i) = i: Next i
    For e = 0 To SM - 1
        a = sPerm(STri(e, 0)): If sPerm(STri(e, 1)) < a Then a = sPerm(STri(e, 1))
        If sPerm(STri(e, 2)) < a Then a = sPerm(STri(e, 2))
        For i = 0 To 2
            If a < sFc(sPerm(STri(e, i))) Then sFc(sPerm(STri(e, i))) = a
        Next i
    Next e
    sPtr(0) = 0
    For i = 0 To SN - 1: sPtr(i + 1) = sPtr(i) + (i - sFc(i) + 1): Next i
End Sub

Private Sub SkylineSolve(ByRef a() As Double, ByRef x() As Double)
    Dim i As Long, j As Long, k As Long, k0 As Long, s As Double
    For i = 0 To SN - 1
        For j = sFc(i) To i
            s = a(sPtr(i) + j - sFc(i))
            k0 = sFc(i): If sFc(j) > k0 Then k0 = sFc(j)
            For k = k0 To j - 1
                s = s - a(sPtr(i) + k - sFc(i)) * a(sPtr(j) + k - sFc(j))
            Next k
            If j < i Then
                a(sPtr(i) + j - sFc(i)) = s / a(sPtr(j) + j - sFc(j))
            Else
                If s <= 0# Then err.Raise 5, , "SEEP_NOT_SPD"
                a(sPtr(i) + i - sFc(i)) = Sqr(s)
            End If
        Next j
    Next i
    For i = 0 To SN - 1
        s = x(i)
        For k = sFc(i) To i - 1: s = s - a(sPtr(i) + k - sFc(i)) * x(k): Next k
        x(i) = s / a(sPtr(i) + i - sFc(i))
    Next i
    For i = SN - 1 To 0 Step -1
        x(i) = x(i) / a(sPtr(i) + i - sFc(i))
        For k = sFc(i) To i - 1: x(k) = x(k) - a(sPtr(i) + k - sFc(i)) * x(i): Next k
    Next i
End Sub

Public Sub RPX_SeepSolve()
    Dim i As Long, j As Long, e As Long, it As Long, ri As Long, rj As Long, gi As Long, gj As Long
    Dim fixed() As Boolean, fval() As Double, a() As Double, f() As Double, x() As Double
    Dim bb(0 To 2) As Double, cc(0 To 2) As Double, area2 As Double, s As Double, p As Double, kij As Double
    Dim hfix As Double, hn As Double, dmax As Double, changed As Long, qtol As Double, haveFix As Boolean
    Dim dfix As Double, res As Double, qs As Double, qi As Double, kr2 As Double, qr() As Double
    Const epsD As Double = 2# ^ (-52)
    Dim qabs() As Double, nterm() As Long, term As Double
    Dim qe As Double, u As Double, budget As Double, ratio As Double, roundoff As Double
    Call RPX_SeepCheckCancelled
    If SKrMin <= 0# Then RPX_SeepDefaults
    BuildOrdering
    Call RPX_SeepCheckCancelled
    ReDim SHead(0 To SN - 1): ReDim SActive(0 To SN - 1): ReDim SFlux(0 To SN - 1): ReDim SKr(0 To SM - 1)
    ReDim fixed(0 To SN - 1): ReDim fval(0 To SN - 1): ReDim f(0 To SN - 1): ReDim x(0 To SN - 1)
    For i = 0 To SN - 1
        If SBcType(i) = 1 Then
            If Not haveFix Or SBcVal(i) > hfix Then hfix = SBcVal(i)
            haveFix = True
        End If
    Next i
    If Not haveFix Then err.Raise 5, , "SEEP_NO_FIXED_HEAD"
    For i = 0 To SN - 1
        If SBcType(i) = 1 Then SHead(i) = SBcVal(i) Else SHead(i) = hfix
        SActive(i) = (SBcType(i) = 2)
        If SActive(i) Then SHead(i) = sxy(i, 1)
    Next i
    SConverged = False
    If SResTol <= 0# Then SResTol = 0.000001
    SLastResidual = 0#: SLastFixDelta = 0#: ReDim qr(0 To SN - 1)
    SLastRawResidual = 0#: SLastBoundaryScale = 0#: SLastRoundoffMax = 0#: SLastRoundoffDominated = False
    ReDim qabs(0 To SN - 1): ReDim nterm(0 To SN - 1)
    For it = 1 To SMaxIt
        Call RPX_SeepSourceAssert
        For i = 0 To SN - 1
            fixed(i) = (SBcType(i) = 1 Or SActive(i))
            If SBcType(i) = 1 Then fval(i) = SBcVal(i) Else fval(i) = sxy(i, 1)
            f(i) = 0#
        Next i
        ReDim a(0 To sPtr(SN) - 1)
        For i = 0 To SN - 1
            If SBcType(i) = 3 Then f(sPerm(i)) = f(sPerm(i)) + SBcVal(i)
        Next i
        For e = 0 To SM - 1
            TriGeom e, area2, bb, cc
            p = 0#
            For i = 0 To 2: p = p + (SHead(STri(e, i)) - sxy(STri(e, i), 1)) / 3#: Next i
            If p >= 0# Then
                SKr(e) = 1#
            Else
                SKr(e) = Exp(SBeta * p)
                If SKr(e) < SKrMin Then SKr(e) = SKrMin
            End If
            s = SKr(e) / (2# * area2)
            For i = 0 To 2
                gi = STri(e, i): ri = sPerm(gi)
                If Not fixed(gi) Then
                    For j = 0 To 2
                        gj = STri(e, j): rj = sPerm(gj)
                        kij = ElemK(e, i, j, s, bb, cc)
                        If fixed(gj) Then
                            f(ri) = f(ri) - kij * fval(gj)
                        ElseIf rj <= ri Then
                            a(sPtr(ri) + rj - sFc(ri)) = a(sPtr(ri) + rj - sFc(ri)) + kij
                        End If
                    Next j
                End If
            Next i
        Next e
        For i = 0 To SN - 1
            If fixed(i) Then ri = sPerm(i): a(sPtr(ri) + ri - sFc(ri)) = 1#: f(ri) = fval(i)
        Next i
        For i = 0 To SN - 1: x(i) = f(i): Next i
        SkylineSolve a, x
        dmax = 0#: dfix = 0#
        For i = 0 To SN - 1
            hn = x(sPerm(i))
            If Not fixed(i) Then
                If Abs(hn - SHead(i)) > dfix Then dfix = Abs(hn - SHead(i))   ' #6: before relaxation
                hn = SHead(i) + SRelax * (hn - SHead(i))
            End If
            If Abs(hn - SHead(i)) > dmax Then dmax = Abs(hn - SHead(i))
            SHead(i) = hn
        Next i
        ' nodal flux q = K h  (q > 0 : inflow into domain at a fixed node)
        For i = 0 To SN - 1: SFlux(i) = 0#: Next i
        For e = 0 To SM - 1
            TriGeom e, area2, bb, cc: s = SKr(e) / (2# * area2)
            For i = 0 To 2
                For j = 0 To 2
                    SFlux(STri(e, i)) = SFlux(STri(e, i)) + ElemK(e, i, j, s, bb, cc) * SHead(STri(e, j))
                Next j
            Next i
        Next e
        qtol = 0#
        For i = 0 To SN - 1
            If Abs(SFlux(i)) > qtol Then qtol = Abs(SFlux(i))
        Next i
        qtol = 0.000000001 * qtol: changed = 0
        For i = 0 To SN - 1
            If SBcType(i) = 2 Then
                If SActive(i) And SFlux(i) > qtol Then
                    SActive(i) = False: changed = changed + 1
                ElseIf (Not SActive(i)) And SHead(i) > sxy(i, 1) Then
                    SActive(i) = True: SHead(i) = sxy(i, 1): changed = changed + 1
                End If
            End If
        Next i
        ' #6: flux balance of the accepted heads (kr re-evaluated at these heads), after seepage-face switching
        For i = 0 To SN - 1: qr(i) = 0#: qabs(i) = 0#: nterm(i) = 0: Next i
        For e = 0 To SM - 1
            TriGeom e, area2, bb, cc
            p = 0#
            For i = 0 To 2: p = p + (SHead(STri(e, i)) - sxy(STri(e, i), 1)) / 3#: Next i
            If p >= 0# Then
                kr2 = 1#
            Else
                kr2 = Exp(SBeta * p)
                If kr2 < SKrMin Then kr2 = SKrMin
            End If
            s = kr2 / (2# * area2)
            For i = 0 To 2
                For j = 0 To 2
                    gi = STri(e, i)
                    term = ElemK(e, i, j, s, bb, cc) * SHead(STri(e, j))
                    qr(gi) = qr(gi) + term
                    qabs(gi) = qabs(gi) + Abs(term)
                    nterm(gi) = nterm(gi) + 1
                Next j
            Next i
        Next e
        qs = 0#: res = 0#
        For i = 0 To SN - 1
            If SBcType(i) = 1 Or SActive(i) Then qs = qs + Abs(qr(i))
            If SBcType(i) = 3 Then qs = qs + Abs(SBcVal(i))
        Next i
        SLastRawResidual = 0#: SLastRoundoffMax = 0#: SLastRoundoffDominated = False
        SLastBoundaryScale = qs
        For i = 0 To SN - 1
            If Not (SBcType(i) = 1 Or SActive(i)) Then
                qi = qr(i)
                If SBcType(i) = 3 Then qi = qi - SBcVal(i)
                qe = qabs(i)
                If SBcType(i) = 3 Then qe = qe + Abs(SBcVal(i))
                u = (CDbl(nterm(i)) + 32#) * epsD
                ' Engineering roundoff allowance; not a bound on geometry or solver error.
                roundoff = (u / (1# - u)) * qe
                budget = SResTol * qs + roundoff
                If Abs(qi) > SLastRawResidual Then SLastRawResidual = Abs(qi)
                If roundoff > SLastRoundoffMax Then SLastRoundoffMax = roundoff
                If roundoff > SResTol * qs Then SLastRoundoffDominated = True
                If budget > 0# Then
                    ratio = Abs(qi) / budget
                ElseIf qi = 0# Then
                    ratio = 0#
                Else
                    ratio = 1E+300
                End If
                If ratio > res Then res = ratio
            End If
        Next i
        SIterations = it: SLastDelta = dmax: SLastFixDelta = dfix: SLastResidual = res
        If dfix < STol And res < 1# And changed = 0 Then SConverged = True: Exit For
        DoEvents
        Call RPX_SeepCheckCancelled
    Next it
    Call RPX_SeepCheckCancelled
    ReDim SPres(0 To SN - 1)
    For i = 0 To SN - 1: SPres(i) = SGammaW * (SHead(i) - sxy(i, 1)): Next i
End Sub

' Free surface (pressure head = 0) as segments: seg(k,0..3) = x1,y1,x2,y2
Public Function RPX_SeepFreeSurface(ByRef seg() As Double) As Long
    Dim e As Long, i As Long, a As Long, b As Long, pa As Double, pb As Double, t As Double, n As Long, m As Long
    Dim px(0 To 2) As Double, py(0 To 2) As Double
    ReDim seg(0 To SM, 0 To 3)
    For e = 0 To SM - 1
        m = 0
        For i = 0 To 2
            a = STri(e, i): b = STri(e, (i + 1) Mod 3)
            pa = SHead(a) - sxy(a, 1): pb = SHead(b) - sxy(b, 1)
            If (pa >= 0# And pb < 0#) Or (pa < 0# And pb >= 0#) Then
                If m < 2 Then
                    t = pa / (pa - pb)
                    px(m) = sxy(a, 0) + t * (sxy(b, 0) - sxy(a, 0)): py(m) = sxy(a, 1) + t * (sxy(b, 1) - sxy(a, 1))
                End If
                m = m + 1
            End If
        Next i
        If m = 2 Then seg(n, 0) = px(0): seg(n, 1) = py(0): seg(n, 2) = px(1): seg(n, 3) = py(1): n = n + 1
    Next e
    RPX_SeepFreeSurface = n
End Function

Private Function NextSeepResultId() As String
    ' Unique per writer call in this VBA session, including same-second output
    ' or a backwards wall clock. Never wrap and silently reuse an old ID.
    If seepWriteSequence = 2147483647 Then err.Raise 5, "RPX_Seep", "SEEP_RESULT_ID_SEQUENCE_EXHAUSTED"
    seepWriteSequence = seepWriteSequence + 1
    NextSeepResultId = Format(now, "yyyymmddhhnnss") & "-" & SN & "-" & SM & "-" & CStr(seepWriteSequence)
End Function
Public Sub RPX_SeepWriteResults()
    Dim ws As Worksheet, out() As Variant, i As Long, seg() As Double, n As Long, qin As Double, qout As Double
    Dim identifier As String, failureNumber As Long, failureSource As String, failureText As String, ownsSource As Boolean
    On Error GoTo Failed
    Call RPX_SeepCheckCancelled
    Call RPX_WaterResultWriteBegin
    ownsSource = Not RPX_SeepSourceIsActive()
    If ownsSource Then Call RPX_SeepSourceBegin
    Call RPX_SeepSourceAssert
    identifier = NextSeepResultId()
    Set ws = RPX_Sheet("浸透結果")
    ws.Range("K2").Value2 = False: ws.Range("K9").ClearContents
    ws.cells.Clear
    For i = 0 To SN - 1
        If SBcType(i) = 1 Or (SBcType(i) = 2 And SActive(i)) Then
            If SFlux(i) > 0# Then qin = qin + SFlux(i) Else qout = qout - SFlux(i)
        End If
    Next i
    ws.Range("A1:H1").value = Array("節点ID", "X [m]", "Y [m]", "全水頭h [m]", "圧力水頭h-y [m]", "間隙水圧u [kPa]", "節点流量 [m3/s/m]（+流入）", "浸出面有効")
    ReDim out(1 To SN, 1 To 8)
    For i = 0 To SN - 1
        out(i + 1, 1) = i + 1: out(i + 1, 2) = sxy(i, 0): out(i + 1, 3) = sxy(i, 1): out(i + 1, 4) = SHead(i)
        out(i + 1, 5) = SHead(i) - sxy(i, 1): out(i + 1, 6) = SPres(i): out(i + 1, 7) = SFlux(i): out(i + 1, 8) = SActive(i)
    Next i
    ws.Range("A2").Resize(SN, 8).value = out
    ws.Range("J1").value = "反復回数": ws.Range("K1").value = SIterations
    ws.Range("J2").value = "収束": ws.Range("K2").value = False
    ws.Range("J3").value = "最終変化量": ws.Range("K3").value = SLastDelta
    ws.Range("J4").value = "流入量": ws.Range("K4").value = qin
    ws.Range("J5").value = "流出量": ws.Range("K5").value = qout
    ws.Range("J6").value = "要素数": ws.Range("K6").value = SM
    ws.Range("J7").value = "節点数": ws.Range("K7").value = SN
    ws.Range("J8").value = "自由水面線分数"
    n = RPX_SeepFreeSurface(seg)
    ws.Range("K8").value = n
    ws.Range("J9").value = "計算ID"
    ws.Range("W1:X1").value = Array("緩和前変化量 [m]", SLastFixDelta)
    ws.Range("W2:X2").value = Array("釣合い残差／混合許容値（<1）", SLastResidual)
    ws.Range("W3:X3").value = Array("収束条件", "緩和前変化量<" & STol & " かつ 残差／混合許容値<1（相対係数=" & SResTol & "）かつ 浸出面切替なし")
    ws.Range("W4:X4").value = Array("最大自由節点残差 [m3/s/m]", SLastRawResidual)
    ws.Range("W5:X5").value = Array("境界流量尺度 [m3/s/m]", SLastBoundaryScale)
    ws.Range("W6:X6").value = Array("最大丸め許容値 [m3/s/m]", SLastRoundoffMax)
    ws.Range("W7:X7").value = Array("丸め許容値が相対成分を超える節点あり", SLastRoundoffDominated)
    ws.Range("R1:U1").value = Array("浸透要素ID", "節点1", "節点2", "節点3")
    ReDim out(1 To SM, 1 To 4)
    For i = 0 To SM - 1
        out(i + 1, 1) = i + 1: out(i + 1, 2) = STri(i, 0) + 1: out(i + 1, 3) = STri(i, 1) + 1: out(i + 1, 4) = STri(i, 2) + 1
    Next i
    ws.Range("R2").Resize(SM, 4).value = out
    ws.Range("M1:P1").value = Array("自由水面 x1", "y1", "x2", "y2")
    If n > 0 Then
        ReDim out(1 To n, 1 To 4)
        For i = 0 To n - 1: out(i + 1, 1) = seg(i, 0): out(i + 1, 2) = seg(i, 1): out(i + 1, 3) = seg(i, 2): out(i + 1, 4) = seg(i, 3): Next i
        ws.Range("M2").Resize(n, 4).value = out
    End If
    ' The complete node/element/result payload is present before eligibility.
    Call RPX_SeepSourceCommit(identifier)
    ws.Range("K9").Value2 = identifier
    ws.Range("K2").Value2 = SConverged
    Call RPX_WaterResultWriteCommit(identifier)
    If ownsSource Then Call RPX_SeepSourceEnd
    Exit Sub
Failed:
    failureNumber = err.number: failureSource = err.source: failureText = err.description
    ' Pending remains set if even invalidating a protected result is impossible.
    ' InputBegin only clears arrays; it cannot make this old result eligible.
    On Error Resume Next
    Call RPX_WaterResultWriteBegin
    If Not ws Is Nothing Then
        ws.Range("K2").Value2 = False: ws.Range("K9").ClearContents
    End If
    If ownsSource Then Call RPX_SeepSourceEnd
    On Error GoTo 0
    err.Raise failureNumber, failureSource, failureText
End Sub

' ---- test: rectangular homogeneous dam, L=20, H1=10, H2=2, k=1 -------------
Public Function RPX_SeepTestDam(Optional ByVal nx As Long = 40, Optional ByVal ny As Long = 20) As String
    Dim l As Double, h1 As Double, H2 As Double, i As Long, j As Long, e As Long, a As Long, b As Long, c As Long, d As Long
    Dim qin As Double, qd As Double, msg As String, k As Long, p0 As Double, p1 As Double, ys As Double
    l = 20#: h1 = 10#: H2 = 2#
    SN = (nx + 1) * (ny + 1): SM = 2 * nx * ny: SNMat = 1
    ReDim sxy(0 To SN - 1, 0 To 1): ReDim STri(0 To SM - 1, 0 To 2): ReDim SMat(0 To SM - 1)
    ReDim SKx(0 To 0): ReDim SKy(0 To 0): SKx(0) = 1#: SKy(0) = 1#
    ReDim SBcType(0 To SN - 1): ReDim SBcVal(0 To SN - 1)
    For j = 0 To ny
        For i = 0 To nx
            sxy(j * (nx + 1) + i, 0) = l * i / nx: sxy(j * (nx + 1) + i, 1) = h1 * j / ny
        Next i
    Next j
    e = 0
    For j = 0 To ny - 1
        For i = 0 To nx - 1
            a = j * (nx + 1) + i: b = a + 1: c = b + nx + 1: d = a + nx + 1
            STri(e, 0) = a: STri(e, 1) = b: STri(e, 2) = c: e = e + 1
            STri(e, 0) = a: STri(e, 1) = c: STri(e, 2) = d: e = e + 1
        Next i
    Next j
    For j = 0 To ny
        a = j * (nx + 1): SBcType(a) = 1: SBcVal(a) = h1
        b = a + nx
        If sxy(b, 1) <= H2 + 0.000000001 Then SBcType(b) = 1: SBcVal(b) = H2 Else SBcType(b) = 2
    Next j
    RPX_SeepDefaults
    RPX_SeepSolve
    For j = 0 To ny: qin = qin + SFlux(j * (nx + 1)): Next j
    qd = (h1 * h1 - H2 * H2) / (2# * l)
    msg = "iters=" & SIterations & " converged=" & SConverged & " Q=" & Format(qin, "0.0000") & _
          " Dupuit=" & Format(qd, "0.0000") & " err%=" & Format(100# * (qin - qd) / qd, "0.00")
    For k = 0 To 4
        i = k * nx \ 4: ys = h1
        For j = 0 To ny - 1
            p0 = SHead(j * (nx + 1) + i) - sxy(j * (nx + 1) + i, 1): p1 = SHead((j + 1) * (nx + 1) + i) - sxy((j + 1) * (nx + 1) + i, 1)
            If p0 >= 0# And p1 < 0# Then ys = sxy(j * (nx + 1) + i, 1) + (h1 / ny) * p0 / (p0 - p1): Exit For
        Next j
        msg = msg & " | x=" & sxy(i, 0) & " y_fs=" & Format(ys, "0.00")
    Next k
    Debug.Print msg
    RPX_SeepTestDam = msg
End Function

Public Sub RPX_SeepTestDamRun()
    MsgBox RPX_SeepTestDam(40, 20)
    RPX_SeepWriteResults
End Sub

' ---- test (#6): 4-node linear problem, exact h2=28/3, h3=29/3 ----------------
Public Function RPX_SeepTestRelax(ByVal relax As Double) As String
    SN = 4: SM = 2: SNMat = 1
    ReDim sxy(0 To 3, 0 To 1): ReDim STri(0 To 1, 0 To 2): ReDim SMat(0 To 1)
    ReDim SKx(0 To 0): ReDim SKy(0 To 0): SKx(0) = 1#: SKy(0) = 1#
    ReDim SBcType(0 To 3): ReDim SBcVal(0 To 3)
    sxy(1, 0) = 1#: sxy(2, 0) = 1#: sxy(2, 1) = 1#: sxy(3, 1) = 1#
    STri(0, 0) = 0: STri(0, 1) = 1: STri(0, 2) = 2: STri(1, 0) = 0: STri(1, 1) = 2: STri(1, 2) = 3
    SBcType(0) = 1: SBcVal(0) = 10#: SBcType(1) = 1: SBcVal(1) = 9#
Call RPX_SeepDefaults:     SRelax = relax
    RPX_SeepSolve
    RPX_SeepTestRelax = "relax=" & relax & ";iters=" & SIterations & ";converged=" & SConverged & ";h2=" & SHead(2) & ";h3=" & SHead(3) & _
        ";err=" & Abs(SHead(2) - 28# / 3#) & ";dfix=" & SLastFixDelta & ";res=" & SLastResidual
End Function
