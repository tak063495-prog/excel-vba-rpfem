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
Public RPX_WaterCommonApplied As Boolean, RPX_WaterCommonNote As String

Private wStamp As String, wN As Long, wM As Long
Private wResultWritePending As Boolean
Private Const WRITE_MARKER_PROPERTY As String = "RPX_SeepResultWriteState"
Private Const WRITE_MARKER_PENDING As String = "PENDING"
Private Const WRITE_MARKER_COMMITTED As String = "COMMITTED|"
Private wXY() As Double, wHead() As Double, wTri() As Long
Private gX0 As Double, gY0 As Double, gDX As Double, gDY As Double, gNX As Long, gNY As Long
Private gStart() As Long, gList() As Long
Private commonInputScope As Boolean, commonInputBound As Boolean, commonEngaged As Boolean
Private commonInputEpoch As Long, commonInputSource As String

Private Const ROW_COUPLE As Long = 19
Private Const ROW_SUCTION As Long = 20

' 新配置（Z1 マーカー）では γw=B4・連成=D4・サクション=F4、旧配置では N13/N19/N20。
Private Function WCell(ByVal ws As Worksheet, ByVal r As Long) As Range
    Dim v2 As Boolean
    On Error Resume Next
    v2 = (CStr(ws.Range("Z1").Value2) = "RPX_SEEP_LAYOUT_V2")
    err.Clear
    On Error GoTo 0
    If v2 Then
        Select Case r
            Case ROW_COUPLE: Set WCell = ws.Range("D4")
            Case ROW_SUCTION: Set WCell = ws.Range("F4")
            Case Else: Set WCell = ws.Range("B4")
        End Select
    Else
        Set WCell = ws.cells(r, 14)
    End If
End Function

Private Function OptSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set OptSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
End Function

Public Function RPX_WaterCouplingOn() As Boolean
    Dim ws As Worksheet, v As Variant
    Set ws = OptSheet("浸透入力"): If ws Is Nothing Then Exit Function
    v = WCell(ws, ROW_COUPLE).Value2
    If IsError(v) Then Exit Function
    If IsNumeric(v) Then
        If Len(CStr(v)) > 0 Then RPX_WaterCouplingOn = (CDbl(v) = 1#)
    End If
End Function

' 連成スイッチの判定でエラーを出さない（OFF扱い）。
Private Function CouplingOnSafe() As Boolean
    On Error Resume Next
    CouplingOnSafe = RPX_WaterCouplingOn()
    If err.number <> 0 Then CouplingOnSafe = False
    err.Clear
End Function

Private Function SuctionOn() As Boolean
    Dim ws As Worksheet, v As Variant
    Set ws = OptSheet("浸透入力"): If ws Is Nothing Then Exit Function
    v = WCell(ws, ROW_SUCTION).Value2
    If IsError(v) Then Exit Function
    If IsNumeric(v) Then
        If Len(CStr(v)) > 0 Then SuctionOn = (CDbl(v) = 1#)
    End If
End Function

Private Function GammaW() As Double
    Dim ws As Worksheet, v As Variant
    GammaW = 9.81
    Set ws = OptSheet("浸透入力"): If ws Is Nothing Then Exit Function
    v = WCell(ws, 13).Value2
    If IsError(v) Then Exit Function
    If IsNumeric(v) Then
        If Len(CStr(v)) > 0 Then GammaW = CDbl(v)
    End If
End Function

' Typed, length-delimited values preserve Empty/String/Boolean/Error and the
' original IEEE double, including +0/-0. No rounded decimal signature is used.
Private Function WCSourceCell(ByVal value As Variant) As String
    Dim text As String, kind As Long
    kind = VarType(value)
    Select Case kind
        Case vbEmpty, vbNull: text = ""
        Case vbDouble: text = RPX_ExactDouble(CDbl(value))
        Case vbByte, vbInteger, vbLong: text = CStr(value)
        Case vbSingle: text = RPX_ExactDouble(CDbl(value))
        Case vbString, vbBoolean, vbError: text = CStr(value)
        Case Else: err.Raise 5, "RPX_Water", "WATER_COMMON_UNSUPPORTED_CELL_TYPE"
    End Select
    WCSourceCell = CStr(kind) & ":" & Len(text) & ":" & text
End Function
Private Function WCSourceRange(ByVal cells As Range) As String
    Dim values As Variant, formulas As Variant, parts() As String, r As Long, c As Long, n As Long
    values = cells.Value2: formulas = cells.formula
    ReDim parts(0 To 2 * cells.rows.count * cells.columns.count - 1)
    For r = 1 To UBound(values, 1)
        For c = 1 To UBound(values, 2)
            parts(n) = WCSourceCell(values(r, c)): n = n + 1
            parts(n) = WCSourceCell(formulas(r, c)): n = n + 1
        Next c
    Next r
    WCSourceRange = cells.Address(False, False) & ":" & Join(parts, "|")
End Function
Private Function WCSourceContext() As String
    Dim ws As Worksheet, result As Worksheet, n As Long, m As Long, parts(0 To 7) As String
    Set ws = OptSheet("浸透入力"): Set result = OptSheet("浸透結果")
    If ws Is Nothing Or result Is Nothing Then err.Raise 5, "RPX_Water", "WATER_COMMON_SOURCE_SHEET_MISSING"
    n = CLng(result.Range("K7").Value2): m = CLng(result.Range("K6").Value2)
    If n < 3 Or n > 100000 Or m < 1 Or m > 220000 Then err.Raise 5, "RPX_Water", "WATER_COMMON_SOURCE_SIZE"
    parts(0) = WCSourceRange(ws.Range("Z1:Z7"))
    parts(1) = WCSourceRange(ws.Range("B3:F4"))
    parts(2) = WCSourceRange(ws.Range("M7:N206"))
    parts(3) = WCSourceRange(result.Range("K2:K10"))
    parts(4) = WCSourceRange(result.Range("A2").Resize(n, 4))
    parts(5) = WCSourceRange(result.Range("R2").Resize(m, 4))
    parts(6) = WCSourceCell(RPX_SeepMode(ws))
    parts(7) = WCSourceCell(RPX_WaterSuspend)
    WCSourceContext = Join(parts, "||")
End Function
Public Sub RPX_WaterCommonInputEnd()
    commonInputScope = False: commonInputBound = False
    commonEngaged = False
    commonInputEpoch = 0: commonInputSource = ""
End Sub
Public Sub RPX_WaterCacheInvalidate()
    ' A new input scope reloads every seep mode, even under an unchanged ID.
    ' Deliberately retain wResultWritePending: input reset cannot validate a
    ' producer that failed before replacing the old sheet's K2=True flag.
    wStamp = "": wN = 0: wM = 0
    Erase wXY: Erase wHead: Erase wTri: Erase gStart: Erase gList
    gX0 = 0#: gY0 = 0#: gDX = 0#: gDY = 0#: gNX = 0: gNY = 0
End Sub
Private Function WaterResultWriteProperty() As Object
    Dim item As Object
    For Each item In ThisWorkbook.CustomDocumentProperties
        If StrComp(CStr(item.name), WRITE_MARKER_PROPERTY, vbTextCompare) = 0 Then
            Set WaterResultWriteProperty = item
            Exit Function
        End If
    Next item
End Function
Private Sub WaterResultWriteMark(ByVal state As String)
    Dim item As Object
    Set item = WaterResultWriteProperty()
    If item Is Nothing Then
        ' The journal belongs to this workbook and remains writable when the
        ' result worksheet is protected. Type4 is msoPropertyTypeString.
        Set item = ThisWorkbook.CustomDocumentProperties.Add(WRITE_MARKER_PROPERTY, False, 4, state)
    Else
        If item.Type <> 4 Then err.Raise 5, "RPX_Water", "INVALID_SEEP_RESULT_WRITE_STATE"
        item.value = state
    End If
    If CStr(item.value) <> state Then err.Raise 5, "RPX_Water", "SEEP_RESULT_WRITE_STATE_NOT_RECORDED"
End Sub
Private Sub WaterResultWriteAssert(Optional ByVal identifier As String = "")
    Dim item As Object, state As String
    If wResultWritePending Then err.Raise 5, "RPX_Water", "SEEP_RESULT_WRITE_INCOMPLETE"
    Set item = WaterResultWriteProperty()
    ' Previously saved normal workbooks have no journal; retain compatibility.
    If item Is Nothing Then Exit Sub
    If item.Type <> 4 Then err.Raise 5, "RPX_Water", "INVALID_SEEP_RESULT_WRITE_STATE"
    state = CStr(item.value)
    If state = WRITE_MARKER_PENDING Then err.Raise 5, "RPX_Water", "SEEP_RESULT_WRITE_INCOMPLETE"
    If Left$(state, Len(WRITE_MARKER_COMMITTED)) <> WRITE_MARKER_COMMITTED Or Len(state) <= Len(WRITE_MARKER_COMMITTED) Then err.Raise 5, "RPX_Water", "INVALID_SEEP_RESULT_WRITE_STATE"
    If Len(identifier) > 0 Then
        If state <> WRITE_MARKER_COMMITTED & identifier Then err.Raise 5, "RPX_Water", "SEEP_RESULT_WRITE_ID_MISMATCH"
    End If
End Sub
Public Sub RPX_WaterResultWriteBegin()
    ' Set before any sheet write, including failure on a protected old result.
    wResultWritePending = True
    Call RPX_WaterCacheInvalidate
    ' Persist before touching K2/K9: VBA reset or save/reopen must not revive
    ' a protected predecessor after a failed output or UI invalidation.
    Call WaterResultWriteMark(WRITE_MARKER_PENDING)
End Sub
Public Sub RPX_WaterResultWriteCommit(ByVal identifier As String)
    If Not wResultWritePending Or Len(identifier) = 0 Then err.Raise 5, "RPX_Water", "INVALID_SEEP_RESULT_COMMIT"
    Call RPX_WaterCacheInvalidate
    Call WaterResultWriteMark(WRITE_MARKER_COMMITTED & identifier)
    wResultWritePending = False
End Sub
Public Sub RPX_WaterCommonInputBegin()
    Dim level As Double, waterGamma As Double
    Call RPX_WaterCommonInputEnd
    Call RPX_WaterCacheInvalidate
    commonInputScope = True: commonInputEpoch = RPX_InputEpoch
    ' Input eligibility deliberately excludes the previous run's mesh/rigids.
    ' A new soil analysis may begin while those globals still describe a rigid
    ' model. Mesh qualification remains in RPX_WaterCommonLevel below.
    If Not WCCommonInputLevel(level, waterGamma) Then Exit Sub
    commonInputSource = WCSourceContext(): commonInputBound = True
End Sub
Public Sub RPX_WaterCommonInputAssert()
    Dim current As String, reason As String
    If Not commonInputScope Or Not commonInputBound Then Exit Sub
    On Error GoTo Changed
    If commonInputEpoch <> RPX_InputEpoch Then reason = "epoch": GoTo Changed
    current = WCSourceContext()
    If current <> commonInputSource Then reason = "typed water input/result context": GoTo Changed
    Exit Sub
Changed:
    If Len(reason) = 0 Then reason = err.description
    On Error Resume Next
    XCancel = True: RPX_ResultCurrent = False: RPX_MeshCurrent = False
    RProgramKind = "": Call RPX_ResetPreparedState: Call RPX_LowerCacheReset: Call RPX_FsBracketReset
    On Error GoTo 0
    err.Raise 5, "RPX_Water", "WATER_COMMON_SOURCE_CHANGED " & reason
End Sub

' Optional exact prescribed horizontal hydrostatic mode only. This does not
' infer a prescribed field from an arbitrary numerical seepage solution.
Private Function WCCommonInputLevel(ByRef level As Double, ByRef waterGamma As Double) As Boolean
    Dim ws As Worksheet, result As Worksheet, vx As Variant, vy As Variant, values As Variant
    Dim seen As Object, key As String, row As Long, count As Long, n As Long, i As Long, x As Double, y As Double
    On Error GoTo Ineligible
    If RPX_WaterSuspend Or Not CouplingOnSafe() Or SuctionOn() Then Exit Function
    Set ws = OptSheet("浸透入力"): If ws Is Nothing Then Exit Function
    If RPX_SeepMode(ws) <> 2 Then Exit Function
    vx = WCell(ws, 13).Value2
    If IsError(vx) Or Not IsNumeric(vx) Then Exit Function
    If Len(CStr(vx)) = 0 Then Exit Function
    waterGamma = GammaW()
    If Not RPX_SchurFinite(waterGamma) Or waterGamma <= 0# Then Exit Function
    Set seen = RPX_Row()
    For row = 7 To 206
        vx = ws.cells(row, 13).Value2: vy = ws.cells(row, 14).Value2
        If IsError(vx) Or IsError(vy) Then Exit Function
        If Len(CStr(vx)) > 0 Or Len(CStr(vy)) > 0 Then
            If Not IsNumeric(vx) Or Not IsNumeric(vy) Then Exit Function
            x = CDbl(vx): y = CDbl(vy)
            If Not RPX_SchurFinite(x) Or Not RPX_SchurFinite(y) Then Exit Function
            If count = 0 Then
                level = y
            ElseIf y <> level Then
                Exit Function
            End If
            key = RPX_ExactDouble(x): If seen.Exists(key) Then Exit Function
            seen(key) = True: count = count + 1
        End If
    Next row
    If count < 2 Then Exit Function
    ' Require a completed, matching prescribed result as well as mode2 input.
    ' A stale FEM result or changed table must keep the existing loading path.
    Set result = OptSheet("浸透結果"): If result Is Nothing Then Exit Function
    If VarType(result.Range("K2").Value2) <> vbBoolean Then Exit Function
    If Not CBool(result.Range("K2").Value2) Or Len(CStr(result.Range("K9").Value2)) = 0 Then Exit Function
    If Left$(CStr(result.Range("K10").Value2), 2) <> "2：" Then Exit Function
    n = CLng(result.Range("K7").Value2): If n < 3 Or n > 100000 Then Exit Function
    values = result.Range("D2").Resize(n, 1).Value2
    For i = 1 To n
        If IsError(values(i, 1)) Or Not IsNumeric(values(i, 1)) Then Exit Function
        If Not RPX_SchurFinite(CDbl(values(i, 1))) Then Exit Function
        If CDbl(values(i, 1)) <> level Then Exit Function
    Next i
    WCCommonInputLevel = True
    Exit Function
Ineligible:
    err.Clear
End Function

Public Function RPX_WaterCommonLevel(ByRef level As Double, ByRef waterGamma As Double) As Boolean
    Dim i As Long
    ' Assert outside the eligibility fallback handler: changed bound sources
    ' abort the analysis and can never silently select the old P1 load.
    Call RPX_WaterCommonInputAssert
    If RPX_WaterSuspend Or Not CouplingOnSafe() Then Exit Function
    Call WaterResultWriteAssert
    If commonInputScope And Not commonInputBound Then Exit Function
    If Not WCCommonInputLevel(level, waterGamma) Or RR <> 0 Then Exit Function
    For i = 0 To re - 1: If RRigidId(i) >= 0 Then Exit Function
    Next i
    RPX_WaterCommonLevel = True
End Function

Public Function RPX_WaterCommonConforming(ByVal level As Double) As Boolean
    Dim e As Long, k As Long, lower As Double, upper As Double, y As Double
    If rn < 3 Or re < 1 Or Not RPX_SchurFinite(level) Then Exit Function
    For e = 0 To re - 1
        lower = RXY(RTri(e, 0), 1): upper = lower
        For k = 1 To 2
            y = RXY(RTri(e, k), 1)
            If y < lower Then lower = y
            If y > upper Then upper = y
        Next k
        If lower < level And upper > level Then Exit Function
    Next e
    RPX_WaterCommonConforming = True
End Function

' 浸透結果シート（A:D 節点, R:U 要素, K9 計算ID）から読み込み。計算IDが同じなら再読込しない。
Private Sub LoadSeepResult()
    Dim ws As Worksheet, stamp As String, n As Long, m As Long, d As Variant, i As Long, k As Long
    Call WaterResultWriteAssert
    Set ws = OptSheet("浸透結果")
    If ws Is Nothing Then err.Raise 5, , "連成：浸透結果シートがありません。浸透解析を先に実行してください。"
    stamp = CStr(ws.Range("K9").Value2)
    If Len(stamp) = 0 Then err.Raise 5, , "連成：浸透結果に要素情報がありません。浸透解析を再実行してください。"
    Call WaterResultWriteAssert(stamp)
    If VarType(ws.Range("K2").Value2) <> vbBoolean Then err.Raise 5, , "連成：浸透結果の収束フラグがありません。"
    If Not CBool(ws.Range("K2").Value2) Then err.Raise 5, , "連成：浸透解析が収束していません。"
    If stamp = wStamp And wN > 0 Then Exit Sub
    n = CLng(ws.Range("K7").Value2): m = CLng(ws.Range("K6").Value2)
    If n < 3 Or m < 1 Then err.Raise 5, , "連成：浸透結果が空です。"
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
    gX0 = xmin: gY0 = ymin: gDX = (xmax - xmin) / gNX + 0.000000000001: gDY = (ymax - ymin) / gNY + 0.000000000001
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

Private Sub bary(ByVal e As Long, ByVal px As Double, ByVal py As Double, ByRef l() As Double)
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
                        e = gList(i): bary e, px, py, l
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
            bary e, px, py, l
            mn = l(0): If l(1) < mn Then mn = l(1)
            If l(2) < mn Then mn = l(2)
            If mn > bestMin Then bestMin = mn: best = e
        Next e
    End If
    bary best, px, py, l
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
    Dim g As Double, wx As Double, wy As Double, rigidWet As Long, level As Double, commonGamma As Double, commonMode As Boolean, below As Boolean
    RPX_WaterApplied = False: RPX_WaterCommonApplied = False
    If RPX_WaterSuspend Then Exit Sub
    If Not CouplingOnSafe() Then RPX_WaterNote = "": Exit Sub
    ' 以下は連成ONのときだけ。エラーは意図的に解析へ返す（理由つき）。
    On Error GoTo Failed
    RPX_SeepTrace "WaterApply ON rn=" & rn & " re=" & re & " phase=" & RPX_LoadPhase
    If rn < 3 Or re < 1 Then Exit Sub
    Call RPX_WaterCommonInputAssert
    LoadSeepResult
    gw = GammaW(): suction = SuctionOn()
    commonMode = RPX_WaterCommonLevel(level, commonGamma)
    If commonMode Then commonMode = RPX_WaterCommonConforming(level)
    If commonInputScope And commonEngaged And Not commonMode Then
        XCancel = True: RPX_ResultCurrent = False: RPX_MeshCurrent = False
        RProgramKind = "": Call RPX_ResetPreparedState: Call RPX_LowerCacheReset: Call RPX_FsBracketReset
        err.Raise 5, "RPX_Water", "WATER_COMMON_LOST_CONFORMITY_OR_ELIGIBILITY"
    End If
    ReDim u(0 To rn - 1)
    For i = 0 To rn - 1
        u(i) = gw * (RPX_WaterHeadAt(RXY(i, 0), RXY(i, 1)) - RXY(i, 1))
        If u(i) < 0# And Not suction Then u(i) = 0#
    Next i
    For e = 0 To re - 1
        g = -RMaterials(RMatId(e)).gamma: wx = 0#: wy = 0#
        If commonMode Then
            ' Exact derivative of gamma_w*max(level-y,0) on each affine piece.
            ' This adds no separate water force and no waterline restraint.
            below = True
            For k = 0 To 2: If RXY(RTri(e, k), 1) > level Then below = False
            Next k
            If below Then wy = commonGamma
        ElseIf RRigidId(e) < 0 Then
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
        RPX_SeismicBody e, True
    Next e
    RPX_WaterApplied = True
    RPX_WaterCommonApplied = commonMode
    If commonInputScope And commonMode Then commonEngaged = True
    RPX_WaterNote = "連成：有効応力（γ - grad u）"
    If commonMode Then RPX_WaterNote = RPX_WaterNote & "／水平静水圧（水位適合）"
    If rigidWet > 0 Then RPX_WaterNote = RPX_WaterNote & "／注意：水面下の剛体要素 " & rigidWet & "（剛体には水圧を作用させていません）"
    Exit Sub
Failed:
    Dim errNo As Long, errText As String
    errNo = err.number: errText = err.description
    RPX_SeepTrace "WaterApply error " & errNo & " " & errText
    err.Raise errNo, "RPX_Water", "浸透連成：" & errText
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
