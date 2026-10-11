Attribute VB_Name = "RPX_Water"
Option Explicit
' RPX_Water : 浸透解析の水圧を極限解析へ一方向連成する（有効応力表示）。
'   応力変数＝有効応力 σ'（引張正）。全応力 σ = σ' - u I。
'   つり合い div σ' + (b - grad u) = 0  →  要素体積力に -grad u を加える。
'   境界：外水圧 p_w = u（水頭境界）なので σ'n = t_user となり、境界水圧荷重は不要
'   （-grad u の体積積分 = -∮u n ds が外水圧の合力を含む）。
'   u は浸透メッシュの節点水頭を極限解析メッシュの節点へ線形補間し、要素内線形 → grad u は要素一定。
' フック：RPX_Mesh.RPX_AssignRegions の末尾（RPX_LoadPhase = "raw" の直後）で Call RPX_WaterApply。
' 体積力は重力＋水で毎回作り直すので何度呼んでも同じ結果（冪等）。

Public RPX_WaterSuspend As Boolean
Public RPX_WaterApplied As Boolean, RPX_WaterNote As String

Private wStamp As String, wN As Long, wM As Long
Private wXY() As Double, wHead() As Double, wTri() As Long
Private gX0 As Double, gY0 As Double, gDX As Double, gDY As Double, gNX As Long, gNY As Long
Private gStart() As Long, gList() As Long

Private Const ROW_COUPLE As Long = 19
Private Const ROW_SUCTION As Long = 20

Private Function OptSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set OptSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
End Function

Public Function RPX_WaterCouplingOn() As Boolean
    Dim ws As Worksheet, v As Variant
    Set ws = OptSheet("浸透入力"): If ws Is Nothing Then Exit Function
    v = ws.Cells(ROW_COUPLE, 14).Value2
    If IsError(v) Then Exit Function
    If IsNumeric(v) Then
        If Len(CStr(v)) > 0 Then RPX_WaterCouplingOn = (CDbl(v) = 1#)
    End If
End Function

' 連成スイッチの判定でエラーを出さない（OFF扱い）。
Private Function CouplingOnSafe() As Boolean
    On Error Resume Next
    CouplingOnSafe = RPX_WaterCouplingOn()
    If Err.Number <> 0 Then CouplingOnSafe = False
    Err.Clear
End Function

Private Function SuctionOn() As Boolean
    Dim ws As Worksheet, v As Variant
    Set ws = OptSheet("浸透入力"): If ws Is Nothing Then Exit Function
    v = ws.Cells(ROW_SUCTION, 14).Value2
    If IsError(v) Then Exit Function
    If IsNumeric(v) Then
        If Len(CStr(v)) > 0 Then SuctionOn = (CDbl(v) = 1#)
    End If
End Function

Private Function GammaW() As Double
    Dim ws As Worksheet, v As Variant
    GammaW = 9.81
    Set ws = OptSheet("浸透入力"): If ws Is Nothing Then Exit Function
    v = ws.Cells(13, 14).Value2
    If IsError(v) Then Exit Function
    If IsNumeric(v) Then
        If Len(CStr(v)) > 0 Then GammaW = CDbl(v)
    End If
End Function

' 浸透結果シート（A:D 節点, R:U 要素, K9 計算ID）から読み込み。計算IDが同じなら再読込しない。
Private Sub LoadSeepResult()
    Dim ws As Worksheet, stamp As String, n As Long, m As Long, d As Variant, i As Long, k As Long
    Set ws = OptSheet("浸透結果")
    If ws Is Nothing Then Err.Raise 5, , "連成：浸透結果シートがありません。浸透解析を先に実行してください。"
    stamp = CStr(ws.Range("K9").Value2)
    If Len(stamp) = 0 Then Err.Raise 5, , "連成：浸透結果に要素情報がありません。浸透解析を再実行してください。"
    If VarType(ws.Range("K2").Value2) <> vbBoolean Then Err.Raise 5, , "連成：浸透結果の収束フラグがありません。"
    If Not CBool(ws.Range("K2").Value2) Then Err.Raise 5, , "連成：浸透解析が収束していません。"
    If stamp = wStamp And wN > 0 Then Exit Sub
    n = CLng(ws.Range("K7").Value2): m = CLng(ws.Range("K6").Value2)
    If n < 3 Or m < 1 Then Err.Raise 5, , "連成：浸透結果が空です。"
    ReDim wXY(0 To n - 1, 0 To 1): ReDim wHead(0 To n - 1): ReDim wTri(0 To m - 1, 0 To 2)
    d = ws.Range("B2").Resize(n, 3).Value2
    For i = 0 To n - 1: wXY(i, 0) = d(i + 1, 1): wXY(i, 1) = d(i + 1, 2): wHead(i) = d(i + 1, 3): Next i
    d = ws.Range("S2").Resize(m, 3).Value2
    For i = 0 To m - 1
        For k = 0 To 2: wTri(i, k) = CLng(d(i + 1, k + 1)) - 1: Next k
    Next i
    wN = n: wM = m: BuildGrid: wStamp = stamp
End Sub

Private Sub BuildGrid()
    Dim e As Long, k As Long, xmin As Double, xmax As Double, ymin As Double, ymax As Double
    Dim ix0 As Long, ix1 As Long, iy0 As Long, iy1 As Long, ix As Long, iy As Long, c As Long, total As Long
    Dim cnt() As Long, fill() As Long, bx0() As Long, bx1() As Long, by0() As Long, by1() As Long
    xmin = wXY(0, 0): xmax = xmin: ymin = wXY(0, 1): ymax = ymin
    For e = 1 To wN - 1
        If wXY(e, 0) < xmin Then xmin = wXY(e, 0)
        If wXY(e, 0) > xmax Then xmax = wXY(e, 0)
        If wXY(e, 1) < ymin Then ymin = wXY(e, 1)
        If wXY(e, 1) > ymax Then ymax = wXY(e, 1)
    Next e
    gNX = Int(Sqr(wM / 2#)) + 1: gNY = gNX
    gX0 = xmin: gY0 = ymin: gDX = (xmax - xmin) / gNX + 1E-12: gDY = (ymax - ymin) / gNY + 1E-12
    ReDim cnt(0 To gNX * gNY): ReDim bx0(0 To wM - 1): ReDim bx1(0 To wM - 1): ReDim by0(0 To wM - 1): ReDim by1(0 To wM - 1)
    For e = 0 To wM - 1
        ix0 = gNX: ix1 = -1: iy0 = gNY: iy1 = -1
        For k = 0 To 2
            ix = CellX(wXY(wTri(e, k), 0)): iy = CellY(wXY(wTri(e, k), 1))
            If ix < ix0 Then ix0 = ix
            If ix > ix1 Then ix1 = ix
            If iy < iy0 Then iy0 = iy
            If iy > iy1 Then iy1 = iy
        Next k
        bx0(e) = ix0: bx1(e) = ix1: by0(e) = iy0: by1(e) = iy1
        For iy = iy0 To iy1
            For ix = ix0 To ix1: cnt(iy * gNX + ix) = cnt(iy * gNX + ix) + 1: total = total + 1: Next ix
        Next iy
    Next e
    ReDim gStart(0 To gNX * gNY): ReDim gList(0 To total): ReDim fill(0 To gNX * gNY)
    For c = 0 To gNX * gNY - 1: gStart(c + 1) = gStart(c) + cnt(c): Next c
    For e = 0 To wM - 1
        For iy = by0(e) To by1(e)
            For ix = bx0(e) To bx1(e)
                c = iy * gNX + ix: gList(gStart(c) + fill(c)) = e: fill(c) = fill(c) + 1
            Next ix
        Next iy
    Next e
End Sub

Private Function CellX(ByVal x As Double) As Long
    CellX = Int((x - gX0) / gDX)
    If CellX < 0 Then CellX = 0
    If CellX > gNX - 1 Then CellX = gNX - 1
End Function
Private Function CellY(ByVal y As Double) As Long
    CellY = Int((y - gY0) / gDY)
    If CellY < 0 Then CellY = 0
    If CellY > gNY - 1 Then CellY = gNY - 1
End Function

Private Sub Bary(ByVal e As Long, ByVal px As Double, ByVal py As Double, ByRef l() As Double)
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double, det As Double
    x1 = wXY(wTri(e, 0), 0): y1 = wXY(wTri(e, 0), 1): x2 = wXY(wTri(e, 1), 0): y2 = wXY(wTri(e, 1), 1)
    x3 = wXY(wTri(e, 2), 0): y3 = wXY(wTri(e, 2), 1)
    det = (x2 - x1) * (y3 - y1) - (x3 - x1) * (y2 - y1)
    l(1) = ((px - x1) * (y3 - y1) - (x3 - x1) * (py - y1)) / det
    l(2) = ((x2 - x1) * (py - y1) - (px - x1) * (y2 - y1)) / det
    l(0) = 1# - l(1) - l(2)
End Sub

' 点 (px,py) の全水頭。含む三角形がなければ最も近い（重心座標の最小値が最大の）三角形で外挿をクランプ。
Public Function RPX_WaterHeadAt(ByVal px As Double, ByVal py As Double) As Double
    Dim c As Long, i As Long, e As Long, l(0 To 2) As Double, best As Long, bestMin As Double, mn As Double, k As Long, s As Double
    Dim ix As Long, iy As Long, r As Long
    best = -1: bestMin = -1E+300
    For r = 0 To 1
        For iy = CellY(py) - r To CellY(py) + r
            For ix = CellX(px) - r To CellX(px) + r
                If ix >= 0 And iy >= 0 And ix < gNX And iy < gNY Then
                    c = iy * gNX + ix
                    For i = gStart(c) To gStart(c + 1) - 1
                        e = gList(i): Bary e, px, py, l
                        mn = l(0): If l(1) < mn Then mn = l(1)
                        If l(2) < mn Then mn = l(2)
                        If mn > bestMin Then bestMin = mn: best = e
                    Next i
                End If
            Next ix
        Next iy
        If bestMin >= -0.000001 Then Exit For
    Next r
    If best < 0 Then
        For e = 0 To wM - 1
            Bary e, px, py, l
            mn = l(0): If l(1) < mn Then mn = l(1)
            If l(2) < mn Then mn = l(2)
            If mn > bestMin Then bestMin = mn: best = e
        Next e
    End If
    Bary best, px, py, l
    If bestMin < 0# Then
        s = 0#
        For k = 0 To 2
            If l(k) < 0# Then l(k) = 0#
            s = s + l(k)
        Next k
        For k = 0 To 2: l(k) = l(k) / s: Next k
    End If
    RPX_WaterHeadAt = l(0) * wHead(wTri(best, 0)) + l(1) * wHead(wTri(best, 1)) + l(2) * wHead(wTri(best, 2))
End Function

' 体積力を重力＋水で作り直す。連成OFFなら何もしない（既存の挙動のまま）。
Public Sub RPX_WaterApply()
    Dim u() As Double, i As Long, e As Long, k As Long, j As Long, gw As Double, suction As Boolean
    Dim g As Double, wx As Double, wy As Double, rigidWet As Long
    RPX_WaterApplied = False
    If RPX_WaterSuspend Then Exit Sub
    If Not CouplingOnSafe() Then RPX_WaterNote = "": Exit Sub
    ' 以下は連成ONのときだけ。エラーは意図的に解析へ返す（理由つき）。
    On Error GoTo Failed
    RPX_SeepTrace "WaterApply ON rn=" & rn & " re=" & re & " phase=" & RPX_LoadPhase
    If rn < 3 Or re < 1 Then Exit Sub
    LoadSeepResult
    gw = GammaW(): suction = SuctionOn()
    ReDim u(0 To rn - 1)
    For i = 0 To rn - 1
        u(i) = gw * (RPX_WaterHeadAt(RXY(i, 0), RXY(i, 1)) - RXY(i, 1))
        If u(i) < 0# And Not suction Then u(i) = 0#
    Next i
    For e = 0 To re - 1
        g = -RMaterials(RMatId(e)).gamma: wx = 0#: wy = 0#
        If RRigidId(e) < 0 Then
            For k = 0 To 2
                wx = wx - u(RTri(e, k)) * RGrad(e, k, 0): wy = wy - u(RTri(e, k)) * RGrad(e, k, 1)
            Next k
        Else
            For k = 0 To 2
                If u(RTri(e, k)) > 0# Then rigidWet = rigidWet + 1: Exit For
            Next k
        End If
        For j = 0 To 1: RBody0(e, j) = 0#: RBody1(e, j) = 0#: Next j
        If RPX_LoadPhase = "reference_total" Then
            RBody1(e, 0) = wx: RBody1(e, 1) = g + wy
        ElseIf RGravityMode = "reference" Then
            RBody0(e, 0) = wx: RBody0(e, 1) = wy: RBody1(e, 1) = g
        Else
            RBody0(e, 0) = wx: RBody0(e, 1) = g + wy
        End If
    Next e
    RPX_WaterApplied = True
    RPX_WaterNote = "連成：有効応力（γ - grad u）"
    If rigidWet > 0 Then RPX_WaterNote = RPX_WaterNote & "／注意：水面下の剛体要素 " & rigidWet & "（剛体には水圧を作用させていません）"
    Exit Sub
Failed:
    Dim errNo As Long, errText As String
    errNo = Err.Number: errText = Err.Description
    RPX_SeepTrace "WaterApply error " & errNo & " " & errText
    Err.Raise errNo, "RPX_Water", "浸透連成：" & errText
End Sub

' 検査用：現在の極限解析メッシュで、全没水（静水）なら土要素の鉛直体積力が -(γ-γw) になることを確認。
Public Function RPX_WaterTestBuoyancy() As String
    Dim e As Long, maxErr As Double, d As Double, gw As Double, fy As Double
    If rn < 3 Or re < 1 Then RPX_WaterTestBuoyancy = "メッシュがありません": Exit Function
    gw = GammaW()
    For e = 0 To re - 1
        If RRigidId(e) < 0 Then
            fy = RBody0(e, 1) + RBody1(e, 1)
            d = Abs(fy + RMaterials(RMatId(e)).gamma - gw) + Abs(RBody0(e, 0) + RBody1(e, 0))
            If d > maxErr Then maxErr = d
        End If
    Next e
    RPX_WaterTestBuoyancy = "全没水チェック：max|b'-(-(γ-γw))| = " & Format(maxErr, "0.0E+00") & " kN/m3（全没水なら0）"
    Debug.Print RPX_WaterTestBuoyancy
End Function
