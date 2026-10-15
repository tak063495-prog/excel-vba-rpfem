Attribute VB_Name = "RPX_SeepWT"
Option Explicit
' 地下水位（浸潤線）入力モードと、観測水位からの境界水頭の逆推定。
'  モード1：水位線＋cos2θ補正　h = y_wt - (y_wt - y)×sin2θ　（u = γw×hw×cos2θ、等ポテンシャル線は水位線に直交）
'  モード2：静水圧　h = y_wt
'  モード3：浸透FEMの未知境界水頭（値セル "?" または "?初期値"、観測点は P:Q）を観測水位点に最小二乗で合わせる（減衰Gauss-Newton＝LM、差分ヤコビアン）
' 結果は浸透結果シートへ同じ形式で書くため、RPX_Water の連成と図の自由水面描画はそのまま使える。
' 透水係数の絶対値は水頭場に影響しない（比だけが効く）。

Public WTN As Long, WTX() As Double, WTY() As Double
Public SUnkCount As Long, SUnkVal() As Double, SUnkReady As Boolean
Public WTInvResult As String

Private Const WT_SHEET As String = "浸透入力"
Private Const WT_FIRST As Long = 7
Private Const WT_LAST As Long = 206

Private wtTest As Boolean, wtTestNx As Long, wtTestNy As Long
Private evalCount As Long

' ---------------------------------------------------------------- table
Public Function RPX_WTReadTable(ByVal ws As Worksheet, ByVal requireTwo As Boolean, ByVal colX As Long) As Long
    Dim r As Long, vx As Variant, vy As Variant, i As Long, j As Long, t As Double
    WTN = 0: ReDim WTX(0 To WT_LAST - WT_FIRST): ReDim WTY(0 To WT_LAST - WT_FIRST)
    For r = WT_FIRST To WT_LAST
        vx = ws.cells(r, colX).Value2: vy = ws.cells(r, colX + 1).Value2
        If Not IsError(vx) And Not IsError(vy) Then
            If Len(CStr(vx)) > 0 And Len(CStr(vy)) > 0 Then
                If Not IsNumeric(vx) Or Not IsNumeric(vy) Then err.Raise 5, , "浸透入力 " & r & "行目の x・y が数値ではありません。"
                WTX(WTN) = CDbl(vx): WTY(WTN) = CDbl(vy): WTN = WTN + 1
            End If
        End If
    Next r
    For i = 1 To WTN - 1
        j = i
        Do While j > 0
            If WTX(j - 1) <= WTX(j) Then Exit Do
            t = WTX(j): WTX(j) = WTX(j - 1): WTX(j - 1) = t
            t = WTY(j): WTY(j) = WTY(j - 1): WTY(j - 1) = t
            j = j - 1
        Loop
    Next i
    If requireTwo Then
        If WTN < 2 Then err.Raise 5, , "地下水面（M:N）に2点以上の (x, y) を入力してください。"
        For i = 1 To WTN - 1
            If WTX(i) - WTX(i - 1) <= 0.000000001 Then err.Raise 5, , "地下水位表に同じxの点があります（x=" & WTX(i) & "）。"
        Next i
    ElseIf WTN < 1 Then
        err.Raise 5, , "観測点（P:Q）がありません。"
    End If
    RPX_WTReadTable = WTN
End Function

Private Function Sin2(ByVal i As Long) As Double
    Dim dx As Double, dy As Double
    dx = WTX(i + 1) - WTX(i): dy = WTY(i + 1) - WTY(i)
    Sin2 = dy * dy / (dx * dx + dy * dy)
End Function

' 水位線の標高と sin2θ（区間ごとの勾配。両端より外は端の値で水平延長、θは端区間の値）
Private Sub WTEval(ByVal x As Double, ByRef yw As Double, ByRef s2 As Double)
    Dim i As Long, t As Double
    If x <= WTX(0) Then yw = WTY(0): s2 = Sin2(0): Exit Sub
    If x >= WTX(WTN - 1) Then yw = WTY(WTN - 1): s2 = Sin2(WTN - 2): Exit Sub
    For i = 0 To WTN - 2
        If x <= WTX(i + 1) Then
            t = (x - WTX(i)) / (WTX(i + 1) - WTX(i))
            yw = WTY(i) + t * (WTY(i + 1) - WTY(i)): s2 = Sin2(i): Exit Sub
        End If
    Next i
    yw = WTY(WTN - 1): s2 = Sin2(WTN - 2)
End Sub

Public Function RPX_WTHead(ByVal x As Double, ByVal y As Double, ByVal mode As Long) As Double
    Dim yw As Double, s2 As Double
    WTEval x, yw, s2
    If mode = 2 Then RPX_WTHead = yw Else RPX_WTHead = yw - (yw - y) * s2
End Function

' SXY/STri（RPX_SeepUI がコピー済み）に水頭場を作る。負圧（水位より上）はそのまま書き、u<0の扱いは連成側（N20）に従う。
Public Sub RPX_WTField(ByVal mode As Long)
    Dim i As Long
    ReDim SHead(0 To SN - 1): ReDim SPres(0 To SN - 1): ReDim SFlux(0 To SN - 1): ReDim SActive(0 To SN - 1)
    ReDim SBcType(0 To SN - 1): ReDim SBcVal(0 To SN - 1): ReDim SKr(0 To SM - 1)
    If SGammaW <= 0# Then SGammaW = 9.81
    For i = 0 To SN - 1
        SHead(i) = RPX_WTHead(sxy(i, 0), sxy(i, 1), mode)
        SPres(i) = SGammaW * (SHead(i) - sxy(i, 1))
    Next i
    For i = 0 To SM - 1: SKr(i) = 1#: Next i
    SIterations = 0: SConverged = True: SLastDelta = 0#: SLastResidual = 0#: SLastFixDelta = 0#
End Sub

' ---------------------------------------------------------------- free surface elevation at x
Public Function RPX_WTSurfaceY(ByRef seg() As Double, ByVal n As Long, ByVal x As Double, ByRef found As Boolean) As Double
    Dim k As Long, lo As Double, hi As Double, y As Double, best As Double, have As Boolean, nd As Double, nearY As Double, d As Double
    nd = 1E+300
    For k = 0 To n - 1
        lo = seg(k, 0): hi = seg(k, 2)
        If lo > hi Then lo = seg(k, 2): hi = seg(k, 0)
        If hi - lo > 0.000000000001 And x >= lo - 0.000000001 And x <= hi + 0.000000001 Then
            y = seg(k, 1) + (seg(k, 3) - seg(k, 1)) * (x - seg(k, 0)) / (seg(k, 2) - seg(k, 0))
            If Not have Or y > best Then best = y: have = True
        End If
        d = Abs(seg(k, 0) - x): If d < nd Then nd = d: nearY = seg(k, 1)
        d = Abs(seg(k, 2) - x): If d < nd Then nd = d: nearY = seg(k, 3)
    Next k
    found = have
    If have Then RPX_WTSurfaceY = best Else RPX_WTSurfaceY = nearY
End Function

' ---------------------------------------------------------------- unknown boundary values
' 浸透境界の値セル（K列）が "?" または "?数値" の行を順に未知数とする。kindは水頭・水位・水位+浸出面のみ。
Public Function RPX_WTScanUnknowns(ByVal ws As Worksheet, ByRef guess() As Double, ByRef labels() As String, ByRef hasGuess() As Boolean) As Long
    Dim r As Long, s As String, n As Long, k As String
    ReDim guess(0 To 9): ReDim labels(0 To 9): ReDim hasGuess(0 To 9)
    For r = WT_FIRST To WT_LAST
        If Len(CStr(ws.cells(r, 5).Value2)) > 0 Then
            s = Trim$(CStr(ws.cells(r, 11).Value2))
            If RPX_WTIsUnknown(s) Then
                k = Trim$(CStr(ws.cells(r, 10).Value2))
                If k <> "水頭" And k <> "水位" And k <> "水位+浸出面" And k <> "水位＋浸出面" Then err.Raise 5, , "未知数（?）は種別「水頭・水位・水位+浸出面」だけに指定できます（境界ID " & ws.cells(r, 5).Value2 & "）。"
                If n > 9 Then err.Raise 5, , "未知数は10個までです。"
                labels(n) = "境界ID " & ws.cells(r, 5).Value2 & "（" & k & "）"
                s = mid$(s, 2)
                If Len(s) > 0 And IsNumeric(s) Then guess(n) = CDbl(s): hasGuess(n) = True
                n = n + 1
            End If
        End If
    Next r
    RPX_WTScanUnknowns = n
End Function

Public Function RPX_WTIsUnknown(ByVal s As String) As Boolean
    s = Trim$(s)
    If Len(s) = 0 Then Exit Function
    RPX_WTIsUnknown = (Left$(s, 1) = "?" Or Left$(s, 1) = "？")
End Function

' ---------------------------------------------------------------- small dense algebra (n <= 10)
Private Function SolveSmall(ByRef a() As Double, ByRef b() As Double, ByVal n As Long) As Boolean
    Dim i As Long, j As Long, k As Long, p As Long, t As Double, f As Double
    For k = 0 To n - 1
        p = k
        For i = k + 1 To n - 1
            If Abs(a(i, k)) > Abs(a(p, k)) Then p = i
        Next i
        If Abs(a(p, k)) < 1E-300 Then Exit Function
        If p <> k Then
            For j = 0 To n - 1: t = a(k, j): a(k, j) = a(p, j): a(p, j) = t: Next j
            t = b(k): b(k) = b(p): b(p) = t
        End If
        For i = k + 1 To n - 1
            f = a(i, k) / a(k, k)
            For j = k To n - 1: a(i, j) = a(i, j) - f * a(k, j): Next j
            b(i) = b(i) - f * b(k)
        Next i
    Next k
    For i = n - 1 To 0 Step -1
        t = b(i)
        For j = i + 1 To n - 1: t = t - a(i, j) * b(j): Next j
        b(i) = t / a(i, i)
    Next i
    SolveSmall = True
End Function

' 対称行列の固有値（Jacobi法）から条件数
Private Function SymCond(ByRef s() As Double, ByVal n As Long) As Double
    Dim a() As Double, i As Long, j As Long, k As Long, p As Long, q As Long, sweep As Long
    Dim off As Double, th As Double, c As Double, sv As Double, t As Double, aip As Double, aiq As Double, mx As Double, mn As Double
    ReDim a(0 To n - 1, 0 To n - 1)
    For i = 0 To n - 1
        For j = 0 To n - 1: a(i, j) = s(i, j): Next j
    Next i
    For sweep = 1 To 50
        off = 0#
        For p = 0 To n - 2
            For q = p + 1 To n - 1: off = off + a(p, q) * a(p, q): Next q
        Next p
        If off < 1E-30 Then Exit For
        For p = 0 To n - 2
            For q = p + 1 To n - 1
                If Abs(a(p, q)) > 1E-300 Then
                    th = (a(q, q) - a(p, p)) / (2# * a(p, q))
                    t = Sgn(th) / (Abs(th) + Sqr(th * th + 1#)): If th = 0# Then t = 1#
                    c = 1# / Sqr(t * t + 1#): sv = t * c
                    For k = 0 To n - 1
                        aip = a(k, p): aiq = a(k, q)
                        a(k, p) = c * aip - sv * aiq: a(k, q) = sv * aip + c * aiq
                    Next k
                    For k = 0 To n - 1
                        aip = a(p, k): aiq = a(q, k)
                        a(p, k) = c * aip - sv * aiq: a(q, k) = sv * aip + c * aiq
                    Next k
                End If
            Next q
        Next p
    Next sweep
    mx = Abs(a(0, 0)): mn = mx
    For i = 1 To n - 1
        If Abs(a(i, i)) > mx Then mx = Abs(a(i, i))
        If Abs(a(i, i)) < mn Then mn = Abs(a(i, i))
    Next i
    If mn <= 0# Then SymCond = 1E+300 Else SymCond = mx / mn
End Function

' ---------------------------------------------------------------- residuals
Private Sub BuildTestDam(ByVal h1 As Double)
    Dim nx As Long, ny As Long, i As Long, j As Long, e As Long, a As Long, b As Long, c As Long, d As Long
    nx = wtTestNx: ny = wtTestNy
    SN = (nx + 1) * (ny + 1): SM = 2 * nx * ny: SNMat = 1
    ReDim sxy(0 To SN - 1, 0 To 1): ReDim STri(0 To SM - 1, 0 To 2): ReDim SMat(0 To SM - 1)
    ReDim SKx(0 To 0): ReDim SKy(0 To 0): SKx(0) = 1#: SKy(0) = 1#
    ReDim SBcType(0 To SN - 1): ReDim SBcVal(0 To SN - 1)
    For j = 0 To ny
        For i = 0 To nx: sxy(j * (nx + 1) + i, 0) = 20# * i / nx: sxy(j * (nx + 1) + i, 1) = 10# * j / ny: Next i
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
        a = j * (nx + 1)
        If sxy(a, 1) <= h1 + 0.000000001 Then SBcType(a) = 1: SBcVal(a) = h1
        b = a + nx
        If sxy(b, 1) <= 2# + 0.000000001 Then SBcType(b) = 1: SBcVal(b) = 2# Else SBcType(b) = 2
    Next j
    RPX_SeepDefaults
End Sub

' 1回の順解析。失敗（固定水頭なし等）は False。
Private Function WTResiduals(ByVal ws As Worksheet, ByRef p() As Double, ByVal n As Long, ByRef obsX() As Double, ByRef obsY() As Double, ByVal m As Long, ByRef r() As Double, ByRef cost As Double) As Boolean
    Dim i As Long, seg() As Double, ns As Long, found As Boolean
    On Error GoTo Failed
    evalCount = evalCount + 1
    For i = 0 To n - 1: SUnkVal(i) = p(i): Next i
    SUnkReady = True
    If wtTest Then BuildTestDam p(0) Else RPX_SeepBuildModel ws
    RPX_SeepSolve
    If Not SConverged Then Exit Function   ' #6: an unconverged forward solve is a failed evaluation
    ns = RPX_SeepFreeSurface(seg)
    If ns < 1 Then Exit Function
    cost = 0#
    For i = 0 To m - 1
        r(i) = RPX_WTSurfaceY(seg, ns, obsX(i), found) - obsY(i)
        cost = cost + r(i) * r(i)
    Next i
    If Not wtTest Then Application.StatusBar = "逆解析：順解析 " & evalCount & " 回目、RMS=" & Format(Sqr(cost / m), "0.000") & " m"
    DoEvents
    WTResiduals = True
    Exit Function
Failed:
    err.Clear
End Function

' LM法。結果は p、戻り値は説明文。
Private Function Invert(ByVal ws As Worksheet, ByRef p() As Double, ByVal n As Long, ByRef labels() As String, ByRef obsX() As Double, ByRef obsY() As Double, ByVal m As Long, _
                        ByVal lo As Double, ByVal hi As Double, ByVal maxEval As Long, ByRef rms As Double, ByRef warnText As String) As Boolean
    Dim r() As Double, rq() As Double, j() As Double, a() As Double, g() As Double, stp() As Double, q() As Double, lam As Double
    Dim b() As Double, cost As Double, costQ As Double, i As Long, k As Long, l As Long, d As Double, cond As Double, colNorm As Double, smax As Double, iter As Long, ok As Boolean
    ReDim r(0 To m - 1): ReDim rq(0 To m - 1): ReDim j(0 To m - 1, 0 To n - 1): ReDim a(0 To n - 1, 0 To n - 1)
    ReDim g(0 To n - 1): ReDim stp(0 To n - 1): ReDim q(0 To n - 1)
    evalCount = 0
    If Not WTResiduals(ws, p, n, obsX, obsY, m, r, cost) Then err.Raise 5, , "逆解析：初期値で浸透解析ができません（初期水頭を「?値」で指定してください）。"
    lam = 0.001
    Do While evalCount + n + 1 <= maxEval
        iter = iter + 1
        For k = 0 To n - 1
            For i = 0 To n - 1: q(i) = p(i): Next i
            d = 0.001 * Abs(p(k)): If d < 0.01 Then d = 0.01
            If q(k) + d > hi Then d = -d
            q(k) = q(k) + d
            If Not WTResiduals(ws, q, n, obsX, obsY, m, rq, costQ) Then err.Raise 5, , "逆解析：差分計算中に浸透解析が失敗しました。"
            For i = 0 To m - 1: j(i, k) = (rq(i) - r(i)) / d: Next i
        Next k
        For k = 0 To n - 1
            g(k) = 0#
            For i = 0 To m - 1: g(k) = g(k) + j(i, k) * r(i): Next i
            For l = 0 To n - 1
                a(k, l) = 0#
                For i = 0 To m - 1: a(k, l) = a(k, l) + j(i, k) * j(i, l): Next i
            Next l
        Next k
        cond = SymCond(a, n)
        ok = False
        Do While evalCount < maxEval
            ReDim b(0 To n - 1, 0 To n - 1)
            For k = 0 To n - 1
                For l = 0 To n - 1: b(k, l) = a(k, l): Next l
                b(k, k) = a(k, k) * (1# + lam) + 0.000000000001
                stp(k) = -g(k)
            Next k
            If Not SolveSmall(b, stp, n) Then
                lam = lam * 10#
                GoTo RetryLam
            End If
            smax = 0#
            For k = 0 To n - 1
                q(k) = p(k) + stp(k)
                If q(k) < lo Then q(k) = lo
                If q(k) > hi Then q(k) = hi
                If Abs(q(k) - p(k)) > smax Then smax = Abs(q(k) - p(k))
            Next k
            If WTResiduals(ws, q, n, obsX, obsY, m, rq, costQ) Then
                If costQ < cost Then
                    For k = 0 To n - 1: p(k) = q(k): Next k
                    For i = 0 To m - 1: r(i) = rq(i): Next i
                    cost = costQ: lam = lam / 10#
                    If lam < 0.0000001 Then lam = 0.0000001
                    ok = True: Exit Do
                End If
            End If
            lam = lam * 10#
RetryLam:
            If lam > 10000000000# Then Exit Do
        Loop
        If Not ok Or smax < 0.0001 Then Exit Do
    Loop
    rms = Sqr(cost / m)
    warnText = ""
    If m < 2 * n + 1 Then warnText = warnText & "観測点が少ない（" & m & "点／未知数" & n & "）。"
    If cond > 100000000# Then warnText = warnText & "悪条件（J^TJの条件数 " & Format(cond, "0.0E+00") & "）：未知数どうしを区別できません。"
    For k = 0 To n - 1
        colNorm = 0#
        For i = 0 To m - 1: colNorm = colNorm + j(i, k) * j(i, k): Next i
        colNorm = Sqr(colNorm / m)
        If colNorm < 0.01 Then warnText = warnText & labels(k) & "は感度が小さい（1 mの変化で水位RMS " & Format(colNorm * 100#, "0.00") & " cm）。"
    Next k
    If evalCount >= maxEval Then warnText = warnText & "順解析の上限回数（" & maxEval & "）に達しました。"
    Invert = True
End Function

Private Function obsMax(ByRef v() As Double, ByVal m As Long) As Double
    Dim i As Long
    obsMax = v(0)
    For i = 1 To m - 1
        If v(i) > obsMax Then obsMax = v(i)
    Next i
End Function

' ---------------------------------------------------------------- mode 3 entry (called by RPX_SeepRun after mesh generation)
Public Sub RPX_WTInverse(ByVal ws As Worksheet, ByVal maxEval As Long)
    Dim guess() As Double, labels() As String, hasGuess() As Boolean, n As Long, m As Long, i As Long, p() As Double
    Dim obsX() As Double, obsY() As Double, lo As Double, hi As Double, rms As Double, warnText As String, ymax As Double, ymin As Double, txt As String
    Dim res As Worksheet
    wtTest = False
    n = RPX_WTScanUnknowns(ws, guess, labels, hasGuess)
    If n = 0 Then err.Raise 5, , "推定する境界の値に「?」がありません。"
    m = RPX_WTReadTable(ws, False, 16)
    ReDim obsX(0 To m - 1): ReDim obsY(0 To m - 1)
    For i = 0 To m - 1: obsX(i) = WTX(i): obsY(i) = WTY(i): Next i
    ymin = RXY(0, 1): ymax = ymin
    For i = 1 To rn - 1
        If RXY(i, 1) < ymin Then ymin = RXY(i, 1)
        If RXY(i, 1) > ymax Then ymax = RXY(i, 1)
    Next i
    lo = ymin: hi = ymax + (ymax - ymin)
    ReDim p(0 To n - 1): ReDim SUnkVal(0 To n - 1): SUnkCount = n
    For i = 0 To n - 1
        If hasGuess(i) Then
            p(i) = guess(i)
        Else
            p(i) = obsMax(obsY, m)
        End If
    Next i
    If maxEval < 3 * (n + 1) Then maxEval = 3 * (n + 1)
    Invert ws, p, n, labels, obsX, obsY, m, lo, hi, maxEval, rms, warnText
    For i = 0 To n - 1: SUnkVal(i) = p(i): Next i
    RPX_SeepBuildModel ws: RPX_SeepSolve: RPX_SeepWriteResults
    Set res = RPX_Sheet("浸透結果")
    res.Range("J10").Value2 = "モード": res.Range("K10").Value2 = "3：未知境界水頭の逆推定"
    res.Range("J11").Value2 = "水位RMS [m]": res.Range("K11").Value2 = rms
    res.Range("J12").Value2 = "順解析回数": res.Range("K12").Value2 = evalCount
    res.Range("J13").Value2 = "観測点数": res.Range("K13").Value2 = m
    res.Range("J14").Value2 = "注意": res.Range("K14").Value2 = IIf(Len(warnText) = 0, "なし", warnText)
    txt = ""
    For i = 0 To n - 1
        res.cells(15 + i, 10).Value2 = labels(i): res.cells(15 + i, 11).Value2 = p(i)
        txt = txt & labels(i) & " = " & Format(p(i), "0.000") & " m" & vbLf
    Next i
    WTInvResult = "推定結果" & vbLf & txt & "水位RMS = " & Format(rms, "0.0000") & " m（" & m & "点、順解析" & evalCount & "回）" & vbLf & _
                  IIf(Len(warnText) = 0, "", "注意：" & warnText & vbLf) & "推定値は浸透入力のK列に書き戻していません（浸透結果 J15以降）。"
    SUnkReady = True
End Sub

Public Sub RPX_WTWriteMode(ByVal mode As Long)
    Dim res As Worksheet
    Set res = RPX_Sheet("浸透結果")
    res.Range("J10").Value2 = "モード"
    res.Range("K10").Value2 = Choose(mode + 1, "0：浸透流FEM", "1：地下水位＋cos2θ補正", "2：地下水位（静水圧）", "3：未知境界水頭の逆推定")
End Sub

' ---------------------------------------------------------------- self test
Public Function RPX_SeepWaterTableSelfTest() As String
    Dim msg As String, th As Double, i As Long, j As Long, nx As Long, ny As Long, x As Double, y As Double, hw As Double, u As Double, emax As Double
    Dim obsX(0 To 4) As Double, obsY(0 To 4) As Double, p(0 To 0) As Double, labels(0 To 0) As String, rms As Double, warnText As String
    Dim seg() As Double, ns As Long, found As Boolean, started As Double
    On Error GoTo Failed
    started = Timer
    ' 1) 無限斜面：u = γw×hw×cos2θ
    th = 30# * 3.14159265358979 / 180#
    WTN = 2: ReDim WTX(0 To 1): ReDim WTY(0 To 1)
    WTX(0) = -10#: WTY(0) = 10# * Tan(th) - 2#: WTX(1) = 50#: WTY(1) = -50# * Tan(th) - 2#
    For i = 0 To 40
        For j = 0 To 12
            x = i: y = -Tan(th) * x - 6# * j / 12#
            hw = (-Tan(th) * x - 2#) - y
            u = 9.81 * (RPX_WTHead(x, y, 1) - y)
            If hw > 0# Then
                If Abs(u - 9.81 * hw * Cos(th) ^ 2) > emax Then emax = Abs(u - 9.81 * hw * Cos(th) ^ 2)
            End If
            If Abs(9.81 * (RPX_WTHead(x, y, 2) - y) - 9.81 * hw) > emax Then emax = Abs(9.81 * (RPX_WTHead(x, y, 2) - y) - 9.81 * hw)
        Next j
    Next i
    msg = "1) 無限斜面 θ=30°：max|u－γw×hw×cos2θ|（モード1）・max|u－γw×hw|（モード2）= " & Format(emax, "0.0E+00") & " kPa → " & IIf(emax < 0.000001, "合格", "不合格") & vbLf
    ' 2) 逆解析：H1=10 の矩形ダム（20×10）から観測5点を作り、初期値8から復元
    wtTest = True: wtTestNx = 20: wtTestNy = 10
    BuildTestDam 10#: RPX_SeepSolve
    ns = RPX_SeepFreeSurface(seg)
    For i = 0 To 4
        obsX(i) = 2# + 3# * i: obsY(i) = RPX_WTSurfaceY(seg, ns, obsX(i), found)
    Next i
    ReDim SUnkVal(0 To 0): SUnkCount = 1: p(0) = 8#: labels(0) = "上流水位H1"
    Invert Nothing, p, 1, labels, obsX, obsY, 5, 0#, 20#, 40, rms, warnText
    msg = msg & "2) 逆解析：H1 推定 " & Format(p(0), "0.0000") & "（真値10、誤差 " & Format(100# * (p(0) - 10#) / 10#, "0.000") & "%）、RMS " & Format(rms, "0.0E+00") & " m、順解析 " & evalCount & " 回 → " & _
          IIf(Abs(p(0) - 10#) < 0.1, "合格", "不合格") & IIf(Len(warnText) > 0, "（注意：" & warnText & "）", "") & vbLf
    GoTo Done
Failed:
    msg = msg & "自己テスト失敗：" & err.number & " " & err.description & vbLf
    Resume Done
Done:
    On Error Resume Next
    wtTest = False: SUnkReady = False
    msg = msg & "所要 " & Format(Timer - started, "0.0") & " 秒"
    Debug.Print msg
    RPX_SeepWaterTableSelfTest = msg
End Function
