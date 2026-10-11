Attribute VB_Name = "RPX_SeepUI"
Option Explicit
' RPX_SeepUI : 浸透入力シート、既存の形状定義からの浸透解析、図シートへの重ね描き。
' 依存：RPX_Seep（ソルバー）、既存の RPX_ReadInputs / RPX_GenerateMesh / RPX_Sheet / DRegions / REdges / RUpperField。
' 既存コードへのフック：RPX_DrawCore の末尾で Call RPX_SeepOverlay（1行のみ）。

Private Const SEEP_SHEET As String = "浸透入力"
Private Const FIRST_ROW As Long = 7
Private Const LAST_ROW As Long = 1006
' 設定は N 列（値）。行番号は固定。
Private Const ROW_TARGET As Long = 7
Private Const ROW_KRMIN As Long = 8
Private Const ROW_BETA As Long = 9
Private Const ROW_RELAX As Long = 10
Private Const ROW_TOL As Long = 11
Private Const ROW_MAXIT As Long = 12
Private Const ROW_GAMMAW As Long = 13
Private Const ROW_RIGIDK As Long = 14
Private Const ROW_DRAWFS As Long = 15
Private Const ROW_SLIPTHR As Long = 16
Private Const ROW_DRAWSTRAIN As Long = 17
Private Const ROW_STRAINTHR As Long = 18
Private Const ROW_COUPLE As Long = 19
Private Const ROW_SUCTION As Long = 20
Private Const ROW_SLIPPCT As Long = 21
Private Const ROW_SLIPMAXFRAC As Long = 22
Private Const ROW_ARCFIT As Long = 23
Private slipPct As Double, slipMaxFrac As Double
Private lastOverlayError As String
Private overlayDiag As String
Private Const LEGEND_MARK As String = "【凡例】"
Private chainCount As Long, chainStartSurf As Boolean, chainEndSurf As Boolean
Private arcCx As Double, arcCy As Double, arcR As Double, arcText As String

Private Type SeepSnapshot
    nodes As Long
    elements As Long
    edges As Long
    xy() As Double
    tri() As Long
    matId() As Long
    rigidId() As Long
    body0() As Double
    body1() As Double
    area() As Double
    grad() As Double
    edgeData() As RPX_Edge
    boundary As Object
    constraints As Object
    programKind As String
    loadPhase As String
    target As Long
    spacing As Double
    strengthFactor As Double
    stressScale As Double
    subdivisions As Long
    fsTolerance As Double
    valid As Boolean
End Type

Private Sub Diag(ByVal text As String)
    overlayDiag = overlayDiag & text & vbLf
End Sub

' 追跡ログ（%TEMP%\rpx_seep_trace.log）。失敗しても無視。
Public Sub RPX_SeepTrace(ByVal text As String)
    Dim f As Integer
    On Error Resume Next
    f = FreeFile
    Open Environ$("TEMP") & "\rpx_seep_trace.log" For Append As #f
    Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " " & text
    Close #f
    Err.Clear
End Sub

Public Function RPX_SeepLastDiagnostics() As String
    RPX_SeepLastDiagnostics = overlayDiag
End Function

' ---------------------------------------------------------------- sheet setup
Private Sub SeepHeader(ByVal ws As Worksheet, ByVal address As String, ByVal labels As Variant)
    Dim i As Long
    For i = LBound(labels) To UBound(labels): ws.Range(address).Offset(0, i).Value2 = labels(i): Next i
    With ws.Range(address).Resize(1, UBound(labels) - LBound(labels) + 1)
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255): .Interior.Color = RGB(38, 67, 91): .WrapText = True
    End With
    ws.Rows(ws.Range(address).Row).RowHeight = 38
End Sub

Private Sub SeepButton(ByVal ws As Worksheet, ByVal label As String, ByVal macro As String, ByVal x As Double, ByVal y As Double, ByVal width As Double)
    Dim s As Shape
    Set s = ws.Shapes.AddShape(5, x, y, width, 28)
    s.TextFrame2.TextRange.Text = label: s.TextFrame2.TextRange.Font.Size = 11: s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    s.Fill.ForeColor.RGB = RGB(36, 104, 130): s.Line.Visible = False: s.OnAction = macro
End Sub

Private Sub DefaultSetting(ByVal ws As Worksheet, ByVal r As Long, ByVal label As String, ByVal value As Variant, ByVal note As String)
    ws.Cells(r, 13).Value2 = label
    If Len(CStr(ws.Cells(r, 14).Value2)) = 0 Then ws.Cells(r, 14).Value2 = value
    ws.Cells(r, 15).Value2 = note
End Sub

' 浸透入力シートを作成・書式設定する。既存の入力値は消さない。
Public Sub RPX_SeepSetupSheet()
    Dim ws As Worksheet, i As Long, previous As Boolean
    previous = RPX_InternalWrite: RPX_InternalWrite = True
    On Error GoTo Done
    Set ws = RPX_Sheet(SEEP_SHEET)
    For i = ws.Shapes.Count To 1 Step -1: ws.Shapes(i).Delete: Next i
    ws.Cells.Font.Name = "Yu Gothic": ws.Cells.Font.Size = 11
    ws.Range("A1").Value2 = "浸透流解析　入力（定常・自由水面）": ws.Range("A1").Font.Size = 20: ws.Range("A1").Font.Bold = True
    ws.Rows(1).RowHeight = 34
    ws.Range("A3").Value2 = "メッシュは「要素定義」の点・領域から生成します（固定メッシュ）。境界は面条件と同じく、領域の定義順に開始点から終了点まで適用します。"
    ws.Range("A4").Value2 = "種別：水頭＝全水頭を固定／水位＝標高が値以下の節点を水頭固定（上は不透水）／水位+浸出面＝値以下を水頭固定・上は浸出面／浸出面／流量＝単位長さ当たり流入量[m3/s/m]（+流入）／不透水。未指定の境界は不透水。"
    ws.Range("A5").Value2 = "透水係数は「材料データ」の材料IDごと。剛体領域は設定の比率×最大透水係数で扱います。"
    ws.Range("A3:A5").WrapText = False
    SeepHeader ws, "A6", Array("材料ID", "kx [m/s]", "ky [m/s]")
    SeepHeader ws, "E6", Array("境界ID", "領域ID", "開始点", "終了点", "循環", "種別", "値 [m]または[m3/s/m]")
    SeepHeader ws, "M6", Array("項目", "値", "説明")
    DefaultSetting ws, ROW_TARGET, "目標要素数", 0, "0：設定シートのTARGETを使用"
    DefaultSetting ws, ROW_KRMIN, "krmin", 0.0001, "不飽和域の透水係数比の下限"
    DefaultSetting ws, ROW_BETA, "β [1/m]", 20, "kr=exp(β・圧力水頭)。大きいほど鋭いが収束しにくい"
    DefaultSetting ws, ROW_RELAX, "緩和係数", 0.5, "0～1"
    DefaultSetting ws, ROW_TOL, "収束判定 [m]", 0.000001, "水頭変化の最大値"
    DefaultSetting ws, ROW_MAXIT, "最大反復回数", 300, ""
    DefaultSetting ws, ROW_GAMMAW, "γw [kN/m3]", 9.81, ""
    DefaultSetting ws, ROW_RIGIDK, "剛体の透水係数比", 0.000001, "剛体領域 k = 比率×最大k"
    DefaultSetting ws, ROW_DRAWFS, "自由水面を描画", 1, "1：図シートに描画"
    DefaultSetting ws, ROW_SLIPTHR, "すべり線しきい値", 0.5, "速度ジャンプ／参照値（N21の分位点）以上の辺を表示。0で非表示"
    DefaultSetting ws, ROW_DRAWSTRAIN, "ひずみ速度を描画", 0, "1：塑性せん断ひずみ速度の大きい要素を着色"
    DefaultSetting ws, ROW_STRAINTHR, "ひずみ速度しきい値", 0.2, "要素のひずみ速度×面積／最大値"
    DefaultSetting ws, ROW_COUPLE, "連成する", 0, "1：浸透結果の間隙水圧を極限解析へ（有効応力、γは湿潤/飽和の全重量）"
    DefaultSetting ws, ROW_SUCTION, "サクションを考慮", 0, "0：u<0は0（推奨）、1：負圧も使用"
    DefaultSetting ws, ROW_SLIPPCT, "すべり線 参照分位点", 0.95, "非ゼロ速度ジャンプの分位点（0.9～0.95推奨）。最大値の特異点に左右されない"
    DefaultSetting ws, ROW_SLIPMAXFRAC, "すべり線 最大表示割合", 0.1, "表示する辺の上限（土の内部辺に対する割合、ジャンプ×辺長の大きい順）"
    DefaultSetting ws, ROW_ARCFIT, "円弧当てはめ", 1, "1：推定すべり面に円弧を当てはめて破線表示（中心・半径を図A5に表示）"
    With ws.Range("N19:N20").Validation
        .Delete: .Add xlValidateList, xlValidAlertStop, xlBetween, "0,1"
        .IgnoreBlank = True: .InCellDropdown = True
    End With
    ws.Columns("A:C").ColumnWidth = 11: ws.Columns("D").ColumnWidth = 3
    ws.Columns("E:I").ColumnWidth = 9: ws.Columns("J").ColumnWidth = 14: ws.Columns("K").ColumnWidth = 16: ws.Columns("L").ColumnWidth = 3
    ws.Columns("M").ColumnWidth = 18: ws.Columns("N").ColumnWidth = 11: ws.Columns("O").ColumnWidth = 44
    ws.Range("A7:C1006,E7:K1006,N7:N23").Interior.Color = RGB(243, 248, 251)
    With ws.Range("I7:I1006").Validation
        .Delete: .Add xlValidateList, xlValidAlertStop, xlBetween, "0,1"
        .IgnoreBlank = True: .InCellDropdown = True
    End With
    With ws.Range("J7:J1006").Validation
        .Delete: .Add xlValidateList, xlValidAlertStop, xlBetween, "水頭,水位,水位+浸出面,浸出面,流量,不透水"
        .IgnoreBlank = True: .InCellDropdown = True: .ShowError = True
        .ErrorTitle = "入力を確認してください": .ErrorMessage = "一覧から選択してください。"
    End With
    SeepButton ws, "浸透解析", "RPX_SeepRun", 560, 5, 100
    SeepButton ws, "図を表示", "RPX_SeepShowDiagram", 670, 5, 90
    SeepButton ws, "書式更新", "RPX_SeepSetupSheet", 770, 5, 90
Done:
    RPX_InternalWrite = previous
    If Err.Number <> 0 Then MsgBox "浸透入力シートの作成に失敗しました：" & Err.Description, vbExclamation
End Sub

Private Function SettingValue(ByVal ws As Worksheet, ByVal r As Long, ByVal fallback As Double) As Double
    Dim v As Variant
    SettingValue = fallback
    v = ws.Cells(r, 14).Value2
    If IsError(v) Then Exit Function
    If IsNumeric(v) Then
        If Len(CStr(v)) > 0 Then SettingValue = CDbl(v)
    End If
End Function

Private Function OptionalSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set OptionalSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
End Function

' ---------------------------------------------------------------- snapshot of limit-analysis mesh
Private Sub SaveState(ByRef st As SeepSnapshot)
    st.valid = False
    st.target = DTarget: st.spacing = DSpacing
    st.strengthFactor = RStrengthFactor: st.stressScale = RStressScale: st.subdivisions = RSubdivisions: st.fsTolerance = RFsTolerance
    st.programKind = RProgramKind: st.loadPhase = RPX_LoadPhase
    st.nodes = rn: st.elements = re: st.edges = RNEdge
    Set st.boundary = RBoundary: Set st.constraints = RConstraints
    If rn > 0 And re > 0 Then
        On Error GoTo NoMesh
        st.xy = RXY: st.tri = RTri: st.matId = RMatId: st.rigidId = RRigidId
        st.body0 = RBody0: st.body1 = RBody1: st.area = RArea: st.grad = RGrad
        If RNEdge > 0 Then st.edgeData = REdges
        st.valid = True
    End If
NoMesh:
End Sub

Private Sub RestoreState(ByRef st As SeepSnapshot)
    DTarget = st.target: DSpacing = st.spacing
    RStrengthFactor = st.strengthFactor: RStressScale = st.stressScale: RSubdivisions = st.subdivisions: RFsTolerance = st.fsTolerance
    RProgramKind = st.programKind: RPX_LoadPhase = st.loadPhase
    Set RBoundary = st.boundary: Set RConstraints = st.constraints
    If st.valid Then
        rn = st.nodes: re = st.elements: RNEdge = st.edges
        RXY = st.xy: RTri = st.tri: RMatId = st.matId: RRigidId = st.rigidId
        RBody0 = st.body0: RBody1 = st.body1: RArea = st.area: RGrad = st.grad
        If st.edges > 0 Then REdges = st.edgeData
    Else
        rn = st.nodes: re = st.elements: RNEdge = st.edges
    End If
End Sub

' ---------------------------------------------------------------- id maps (same order as RPX_ReadInputs)
Private Function ColumnMap(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal col As Long) As Object
    Dim d As Object, r As Long, n As Long, v As Variant
    Set d = CreateObject("Scripting.Dictionary")
    For r = firstRow To lastRow
        v = ws.Cells(r, col).Value2
        If Len(CStr(v)) > 0 Then
            If Not d.Exists(CLng(v)) Then d.Add CLng(v), n
            n = n + 1
        End If
    Next r
    Set ColumnMap = d
End Function

' 領域 region の境界を startPoint から endPoint まで定義順にたどり、点が含まれるか。
Private Function PathContains(ByVal region As Long, ByVal startPoint As Long, ByVal endPoint As Long, ByVal px As Double, ByVal py As Double) As Boolean
    Dim i As Long, a As Long, b As Long, stepCount As Long, n As Long, found As Boolean
    n = DRegions(region).pointCount
    For i = 0 To n - 1
        If DRegions(region).points(i) = startPoint Then found = True: Exit For
    Next i
    If Not found Then Exit Function
    For stepCount = 1 To n
        a = DRegions(region).points(i): b = DRegions(region).points((i + 1) Mod n)
        If RPX_OnSegment(px, py, DPXY(a, 0), DPXY(a, 1), DPXY(b, 0), DPXY(b, 1)) Then PathContains = True: Exit Function
        If b = endPoint Then Exit For
        i = (i + 1) Mod n
    Next stepCount
End Function

Private Function KindCode(ByVal text As String) As Long
    Select Case Trim$(text)
        Case "水頭": KindCode = 1
        Case "水位": KindCode = 2
        Case "水位+浸出面", "水位＋浸出面": KindCode = 3
        Case "浸出面": KindCode = 4
        Case "流量": KindCode = 5
        Case "不透水": KindCode = 6
        Case Else: Err.Raise 5, , "浸透境界の種別が不明です：" & text
    End Select
End Function

' 節点の優先度：水頭固定 3 > 浸出面 2 > 流量/不透水 0
Private Sub SetNodeBc(ByVal node As Long, ByVal kind As Long, ByVal value As Double, ByRef pri() As Long)
    Dim p As Long
    Select Case kind
        Case 1: p = 3
        Case 2: p = 2
        Case Else: p = 0
    End Select
    If p > pri(node) Then
        pri(node) = p: SBcType(node) = kind: SBcVal(node) = value
    ElseIf p = 3 And pri(node) = 3 Then
        If Abs(SBcVal(node) - value) > 0.000000001 Then Err.Raise 5, , "同じ節点に異なる水頭が指定されています（節点 x=" & SXY(node, 0) & ", y=" & SXY(node, 1) & "）。"
    End If
End Sub

' ---------------------------------------------------------------- run
Private Sub BuildSeepModel(ByVal ws As Worksheet)
    Dim matWs As Worksheet, defWs As Worksheet, matMap As Object, pointMap As Object, regionMap As Object
    Dim r As Long, i As Long, e As Long, id As Long, kind As Long, value As Double, kmax As Double, ratio As Double
    Dim haveK() As Boolean, region As Long, sp As Long, ep As Long, px As Double, py As Double, node As Long, j As Long
    Dim pri() As Long, hasHead As Boolean, lengthE As Double, key As Variant
    Set matWs = RPX_Sheet("材料データ"): Set defWs = RPX_Sheet("要素定義")
    Set matMap = ColumnMap(matWs, 5, 1004, 1)
    Set pointMap = ColumnMap(defWs, 7, 1006, 1)
    Set regionMap = ColumnMap(defWs, 7, 1006, 5)
    ' mesh copy
    SN = rn: SM = re: SNMat = RM + 1
    ReDim SXY(0 To SN - 1, 0 To 1): ReDim STri(0 To SM - 1, 0 To 2): ReDim SMat(0 To SM - 1)
    For i = 0 To SN - 1: SXY(i, 0) = RXY(i, 0): SXY(i, 1) = RXY(i, 1): Next i
    For e = 0 To SM - 1
        For j = 0 To 2: STri(e, j) = RTri(e, j): Next j
        If RRigidId(e) >= 0 Then SMat(e) = RM Else SMat(e) = RMatId(e)
    Next e
    ' permeability
    ReDim SKx(0 To SNMat - 1): ReDim SKy(0 To SNMat - 1): ReDim haveK(0 To SNMat - 1)
    For r = FIRST_ROW To LAST_ROW
        If Len(CStr(ws.Cells(r, 1).Value2)) > 0 Then
            id = CLng(ws.Cells(r, 1).Value2)
            If Not matMap.Exists(id) Then Err.Raise 5, , "浸透入力の材料ID " & id & " は材料データにありません。"
            i = matMap(id)
            SKx(i) = CDbl(ws.Cells(r, 2).Value2): SKy(i) = CDbl(ws.Cells(r, 3).Value2)
            If Len(CStr(ws.Cells(r, 3).Value2)) = 0 Then SKy(i) = SKx(i)
            If SKx(i) <= 0# Or SKy(i) <= 0# Then Err.Raise 5, , "材料ID " & id & " の透水係数は正の値にしてください。"
            haveK(i) = True
            If SKx(i) > kmax Then kmax = SKx(i)
            If SKy(i) > kmax Then kmax = SKy(i)
        End If
    Next r
    For e = 0 To SM - 1
        If SMat(e) < RM Then
            If Not haveK(SMat(e)) Then
                For Each key In matMap.Keys
                    If matMap(key) = SMat(e) Then Exit For
                Next key
                Err.Raise 5, , "材料ID " & key & " の透水係数が未入力です（浸透入力 A:C）。"
            End If
        End If
    Next e
    ratio = SettingValue(ws, ROW_RIGIDK, 0.000001)
    SKx(RM) = ratio * kmax: SKy(RM) = ratio * kmax
    ' boundary conditions on exposed mesh edges
    ReDim SBcType(0 To SN - 1): ReDim SBcVal(0 To SN - 1): ReDim pri(0 To SN - 1)
    For i = 0 To SN - 1: pri(i) = -1: Next i
    For r = FIRST_ROW To LAST_ROW
        If Len(CStr(ws.Cells(r, 5).Value2)) > 0 Then
            id = CLng(ws.Cells(r, 6).Value2)
            If Not regionMap.Exists(id) Then Err.Raise 5, , "浸透境界 " & ws.Cells(r, 5).Value2 & "：領域ID " & id & " は未定義です。"
            region = regionMap(id)
            id = CLng(ws.Cells(r, 7).Value2): If Not pointMap.Exists(id) Then Err.Raise 5, , "浸透境界の開始点 " & id & " は未定義です。"
            sp = pointMap(id)
            id = CLng(ws.Cells(r, 8).Value2): If Not pointMap.Exists(id) Then Err.Raise 5, , "浸透境界の終了点 " & id & " は未定義です。"
            ep = pointMap(id)
            kind = KindCode(CStr(ws.Cells(r, 10).Value2))
            value = 0#: If Len(CStr(ws.Cells(r, 11).Value2)) > 0 Then value = CDbl(ws.Cells(r, 11).Value2)
            For i = 0 To RNEdge - 1
                If REdges(i).element(1) < 0 Then
                    px = (RXY(REdges(i).a, 0) + RXY(REdges(i).b, 0)) / 2#: py = (RXY(REdges(i).a, 1) + RXY(REdges(i).b, 1)) / 2#
                    If PathContains(region, sp, ep, px, py) Then
                        lengthE = REdges(i).length
                        For j = 0 To 1
                            If j = 0 Then node = REdges(i).a Else node = REdges(i).b
                            Select Case kind
                                Case 1: SetNodeBc node, 1, value, pri
                                Case 2
                                    If SXY(node, 1) <= value + 0.000000001 Then SetNodeBc node, 1, value, pri
                                Case 3
                                    If SXY(node, 1) <= value + 0.000000001 Then SetNodeBc node, 1, value, pri Else SetNodeBc node, 2, 0#, pri
                                Case 4: SetNodeBc node, 2, 0#, pri
                                Case 5
                                    If pri(node) <= 0 Then
                                        If SBcType(node) <> 3 Then SBcType(node) = 3: SBcVal(node) = 0#
                                        SBcVal(node) = SBcVal(node) + value * lengthE / 2#: pri(node) = 0
                                    End If
                                Case 6
                            End Select
                        Next j
                    End If
                End If
            Next i
        End If
    Next r
    For i = 0 To SN - 1
        If SBcType(i) = 1 Then hasHead = True: Exit For
    Next i
    If Not hasHead Then Err.Raise 5, , "水頭固定の境界がありません（水頭・水位を指定してください）。"
    ' settings
    RPX_SeepDefaults
    SKrMin = SettingValue(ws, ROW_KRMIN, SKrMin): SBeta = SettingValue(ws, ROW_BETA, SBeta)
    SRelax = SettingValue(ws, ROW_RELAX, SRelax): STol = SettingValue(ws, ROW_TOL, STol)
    SMaxIt = CLng(SettingValue(ws, ROW_MAXIT, SMaxIt)): SGammaW = SettingValue(ws, ROW_GAMMAW, SGammaW)
    If SRelax <= 0# Or SRelax > 1# Then Err.Raise 5, , "緩和係数は0より大きく1以下にしてください。"
    If SKrMin <= 0# Or SKrMin >= 1# Then Err.Raise 5, , "krminは0より大きく1未満にしてください。"
End Sub

' 浸透解析ボタン：要素定義からメッシュ生成→境界割当→解析→浸透結果・図。極限解析の状態は復元する。
Public Sub RPX_SeepRun()
    Dim ws As Worksheet, st As SeepSnapshot, previousWrite As Boolean, target As Long, failText As String, started As Double
    If RPX_Busy Then MsgBox "解析またはメッシュ生成を実行中です。", vbExclamation: Exit Sub
    Set ws = OptionalSheet(SEEP_SHEET)
    If ws Is Nothing Then RPX_SeepSetupSheet: MsgBox "浸透入力シートを作成しました。入力後にもう一度実行してください。", vbInformation: Exit Sub
    SaveState st
    previousWrite = RPX_InternalWrite
    RPX_Busy = True: RPX_InternalWrite = True: RPX_WaterSuspend = True: started = Timer
    Application.ScreenUpdating = False
    On Error GoTo ErrHandler
    Application.StatusBar = "浸透解析：メッシュ生成中…"
    Call RPX_ReadInputs
    target = CLng(SettingValue(ws, ROW_TARGET, 0#))
    If target > 0 Then DTarget = target
    Call RPX_GenerateMesh
    BuildSeepModel ws
    Application.StatusBar = "浸透解析：反復計算中（" & SM & "要素）…"
    RPX_SeepSolve
    RPX_SeepWriteResults
    GoTo Finish
ErrHandler:
    failText = Err.Description
    Resume Finish
Finish:
    On Error Resume Next
    RestoreState st
    RPX_WaterSuspend = False
    RPX_Busy = False: RPX_InternalWrite = previousWrite
    Application.StatusBar = False: Application.ScreenUpdating = True
    On Error GoTo 0
    If Len(failText) > 0 Then
        MsgBox "浸透解析に失敗しました：" & failText, vbExclamation
    Else
        If RPX_WaterCouplingOn() Then RPX_Invalidate
        RPX_SeepShowDiagram
        If Not SConverged Then
            MsgBox "浸透解析は最大反復回数で打ち切りました（最終変化量 " & Format(SLastDelta, "0.0E+00") & " m）。βを小さく、または緩和係数を小さくしてください。", vbExclamation
        Else
            Application.StatusBar = "浸透解析完了：" & SIterations & "回反復、" & Format(Timer - started, "0.0") & "秒"
        End If
    End If
End Sub

' ---------------------------------------------------------------- drawing
' 図シートの座標変換（RPX_DrawCore と同じ規則：左30pt、上55pt、幅780pt/高さ460pt、縦横比1:1）
Private Function DrawExtent(ByRef xmin As Double, ByRef ymax As Double, ByRef factor As Double) As Boolean
    Dim i As Long, xmax As Double, ymin As Double, res As Worksheet, n As Long, data As Variant
    If rn >= 3 And RPX_MeshCurrent Then
        xmin = RXY(0, 0): xmax = xmin: ymin = RXY(0, 1): ymax = ymin
        For i = 1 To rn - 1
            If RXY(i, 0) < xmin Then xmin = RXY(i, 0)
            If RXY(i, 0) > xmax Then xmax = RXY(i, 0)
            If RXY(i, 1) < ymin Then ymin = RXY(i, 1)
            If RXY(i, 1) > ymax Then ymax = RXY(i, 1)
        Next i
    Else
        Set res = OptionalSheet("浸透結果"): If res Is Nothing Then Exit Function
        n = res.Cells(res.Rows.Count, 2).End(xlUp).Row - 1
        If n < 3 Then Exit Function
        data = res.Range("B2").Resize(n, 2).Value2
        xmin = data(1, 1): xmax = xmin: ymin = data(1, 2): ymax = ymin
        For i = 2 To n
            If data(i, 1) < xmin Then xmin = data(i, 1)
            If data(i, 1) > xmax Then xmax = data(i, 1)
            If data(i, 2) < ymin Then ymin = data(i, 2)
            If data(i, 2) > ymax Then ymax = data(i, 2)
        Next i
    End If
    If xmax <= xmin Or ymax <= ymin Then Exit Function
    factor = 780# / (xmax - xmin)
    If factor * (ymax - ymin) > 460# Then factor = 460# / (ymax - ymin)
    DrawExtent = True
End Function

Private Sub DeleteTagged(ByVal ws As Worksheet, ByVal tag As String)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If ws.Shapes(i).AlternativeText = tag Then ws.Shapes(i).Delete
    Next i
End Sub

Private Sub DrawFreeSurface(ByVal ws As Worksheet, ByVal xmin As Double, ByVal ymax As Double, ByVal factor As Double)
    Dim res As Worksheet, n As Long, data As Variant, i As Long, s As Shape
    Set res = OptionalSheet("浸透結果"): If res Is Nothing Then Exit Sub
    n = res.Cells(res.Rows.Count, 13).End(xlUp).Row - 1
    If n < 1 Then Exit Sub
    data = res.Range("M2").Resize(n, 4).Value2
    If n = 1 Then
        If Not IsNumeric(data(1, 1)) Then Exit Sub
    End If
    For i = 1 To n
        If IsNumeric(data(i, 1)) And Len(CStr(data(i, 1))) > 0 Then
            Set s = ws.Shapes.AddLine(30# + factor * (data(i, 1) - xmin), 55# + factor * (ymax - data(i, 2)), 30# + factor * (data(i, 3) - xmin), 55# + factor * (ymax - data(i, 4)))
            s.AlternativeText = "RPX_SEEP": s.Line.ForeColor.RGB = RGB(0, 112, 192): s.Line.Weight = 2.25
        End If
    Next i
End Sub

' 上界速度場から：要素境界の速度ジャンプ（すべり線）と要素の塑性せん断ひずみ速度
Private Function UpperFieldReady() As Boolean
    On Error GoTo NotReady
    If Not RPX_ResultCurrent Or Not RPX_MeshCurrent Or re < 1 Then Exit Function
    If UBound(RUpperField) <> 12 * re + 3 * RR - 1 Then Exit Function
    UpperFieldReady = True
NotReady:
End Function

Private Function EdgeJump(ByVal id As Long) As Double
    Dim e As Long, f As Long, k As Long, dx As Double, dy As Double, m As Double
    e = REdges(id).element(0): f = REdges(id).element(1)
    For k = 0 To 2
        dx = RUpperField(12 * f + 2 * REdges(id).loc(1, k)) - RUpperField(12 * e + 2 * REdges(id).loc(0, k))
        dy = RUpperField(12 * f + 2 * REdges(id).loc(1, k) + 1) - RUpperField(12 * e + 2 * REdges(id).loc(0, k) + 1)
        If Sqr(dx * dx + dy * dy) > m Then m = Sqr(dx * dx + dy * dy)
    Next k
    EdgeJump = m
End Function

Private Function ElementShearRate(ByVal e As Long) As Double
    Dim bary(0 To 2) As Double, g(0 To 5, 0 To 1) As Double, k As Long, exx As Double, eyy As Double, gxy As Double, vx As Double, vy As Double
    bary(0) = 1# / 3#: bary(1) = 1# / 3#: bary(2) = 1# / 3#
    RPX_P2Grad e, bary, g
    For k = 0 To 5
        vx = RUpperField(12 * e + 2 * k): vy = RUpperField(12 * e + 2 * k + 1)
        exx = exx + g(k, 0) * vx: eyy = eyy + g(k, 1) * vy: gxy = gxy + g(k, 1) * vx + g(k, 0) * vy
    Next k
    ElementShearRate = Sqr((exx - eyy) ^ 2 + gxy ^ 2)
End Function

Private Sub DrawMechanism(ByVal ws As Worksheet, ByVal xmin As Double, ByVal ymax As Double, ByVal factor As Double, ByVal slipThr As Double, ByVal drawStrain As Boolean, ByVal strainThr As Double)
    Dim i As Long, e As Long, k As Long, jump() As Double, rate() As Double, maxJ As Double, maxR As Double, t As Double
    Dim s As Shape, a As Long, b As Long, ff As FreeformBuilder
    If drawStrain Then
        ReDim rate(0 To re - 1)
        For e = 0 To re - 1
            If RRigidId(e) < 0 Then rate(e) = ElementShearRate(e) * RArea(e)
            If rate(e) > maxR Then maxR = rate(e)
        Next e
        If maxR > 0# Then
            For e = 0 To re - 1
                t = rate(e) / maxR
                If t >= strainThr Then
                    Set ff = ws.Shapes.BuildFreeform(0, 30# + factor * (RXY(RTri(e, 0), 0) - xmin), 55# + factor * (ymax - RXY(RTri(e, 0), 1)))
                    For k = 1 To 3
                        ff.AddNodes 0, 0, 30# + factor * (RXY(RTri(e, k Mod 3), 0) - xmin), 55# + factor * (ymax - RXY(RTri(e, k Mod 3), 1))
                    Next k
                    Set s = ff.ConvertToShape
                    s.AlternativeText = "RPX_VELOCITY": s.Line.Visible = False
                    s.Fill.ForeColor.RGB = RGB(237, 125, 49): s.Fill.Transparency = 0.85 - 0.7 * t
                    s.ZOrder 1
                End If
            Next e
        End If
    End If
    If slipThr > 0# Then DrawSlip ws, xmin, ymax, factor, slipThr
End Sub

Private Sub SortDoubles(ByRef v() As Double, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, p As Double, t As Double
    Do While lo < hi
        i = lo: j = hi: p = v((lo + hi) \ 2)
        Do While i <= j
            Do While v(i) < p: i = i + 1: Loop
            Do While v(j) > p: j = j - 1: Loop
            If i <= j Then t = v(i): v(i) = v(j): v(j) = t: i = i + 1: j = j - 1
        Loop
        If j - lo < hi - i Then
            If lo < j Then SortDoubles v, lo, j
            lo = i
        Else
            If i < hi Then SortDoubles v, i, hi
            hi = j
        End If
    Loop
End Sub

Private Function PX(ByVal x As Double, ByVal xmin As Double, ByVal factor As Double) As Double
    PX = 30# + factor * (x - xmin)
End Function
Private Function PY(ByVal y As Double, ByVal ymax As Double, ByVal factor As Double) As Double
    PY = 55# + factor * (ymax - y)
End Function

' 速度ジャンプの大きい辺を分位点基準で選び、連結成分の最長経路を推定すべり面として太線で描く。
Private Sub DrawSlip(ByVal ws As Worksheet, ByVal xmin As Double, ByVal ymax As Double, ByVal factor As Double, ByVal thr As Double)
    Dim i As Long, j As Long, k As Long, n As Long, nz As Long, a As Long, b As Long, s As Shape
    Dim jump() As Double, score() As Double, vals() As Double, ref As Double, cutScore As Double, maxCount As Long
    Dim sel() As Boolean, nSel As Long, t As Double
    Dim adjStart() As Long, adjList() As Long, adjEdge() As Long, deg() As Long, fillN() As Long
    Dim comp() As Long, nComp As Long, compLen() As Double, best As Long, stack() As Long, sp As Long, v As Long, w As Long
    Dim dist() As Double, prev() As Long, done() As Boolean, src As Long, far As Long, pass As Long, pts() As Single, chain As Long
    If RNEdge < 1 Then Exit Sub
    If slipPct <= 0# Or slipPct >= 1# Then slipPct = 0.95
    If slipMaxFrac <= 0# Or slipMaxFrac > 1# Then slipMaxFrac = 0.1
    ReDim jump(0 To RNEdge - 1): ReDim score(0 To RNEdge - 1): ReDim sel(0 To RNEdge - 1)
    For i = 0 To RNEdge - 1
        If REdges(i).element(1) >= 0 Then
            If RRigidId(REdges(i).element(0)) < 0 Or RRigidId(REdges(i).element(1)) < 0 Then
                jump(i) = EdgeJump(i): score(i) = jump(i) * REdges(i).length: n = n + 1
                If jump(i) > t Then t = jump(i)
            End If
        End If
    Next i
    Diag "すべり線: 対象内部辺=" & n & " 最大ジャンプ=" & t
    If t <= 0# Then Exit Sub
    ReDim vals(0 To n)
    For i = 0 To RNEdge - 1
        If jump(i) > 0.000001 * t Then vals(nz) = jump(i): nz = nz + 1
    Next i
    Diag "すべり線: 非ゼロ辺=" & nz
    If nz = 0 Then Exit Sub
    SortDoubles vals, 0, nz - 1
    ref = vals(Int(slipPct * (nz - 1)))
    If ref <= 0# Then ref = t
    Diag "すべり線: 参照値=" & ref
    ' 上限：ジャンプ×辺長（散逸の目安）の大きい順に最大 slipMaxFrac×内部辺数
    maxCount = Int(slipMaxFrac * n) + 1
    ReDim vals(0 To RNEdge - 1): nz = 0
    For i = 0 To RNEdge - 1
        If jump(i) >= thr * ref Then vals(nz) = score(i): nz = nz + 1
    Next i
    If nz = 0 Then Exit Sub
    cutScore = 0#
    If nz > maxCount Then
        SortDoubles vals, 0, nz - 1
        cutScore = vals(nz - maxCount)
    End If
    For i = 0 To RNEdge - 1
        If jump(i) >= thr * ref And score(i) >= cutScore Then sel(i) = True: nSel = nSel + 1
    Next i
    Diag "すべり線: 候補=" & nz & " 上限=" & maxCount & " 選択=" & nSel
    ' 選択辺（細線）
    For i = 0 To RNEdge - 1
        If sel(i) Then
            a = REdges(i).a: b = REdges(i).b
            Set s = ws.Shapes.AddLine(PX(RXY(a, 0), xmin, factor), PY(RXY(a, 1), ymax, factor), PX(RXY(b, 0), xmin, factor), PY(RXY(b, 1), ymax, factor))
            t = jump(i) / ref: If t > 1# Then t = 1#
            s.AlternativeText = "RPX_VELOCITY": s.Line.ForeColor.RGB = RGB(150, 90, 200): s.Line.Weight = 1.5 + 1# * t
            s.ZOrder 0
        End If
    Next i
    ' 連続した推定すべり面：地表節点間の最短路（重み＝辺長／(ジャンプ/参照+0.001)）
    TraceSlipSurface ws, xmin, ymax, factor, jump, ref, sel
    If ws.Name = "図" Then ws.Range("A5").Value2 = "すべり線：内部辺" & n & "本中" & nSel & "本表示（参照=" & Format(slipPct, "0%") & "分位点、しきい値" & thr & "×参照）。推定すべり面の節点数 " & chainCount & arcText
End Sub

' 最短路（二分ヒープ、遅延削除）
Private Sub ShortestPaths(ByVal src As Long, ByRef adjStart() As Long, ByRef adjList() As Long, ByRef adjW() As Double, ByRef dist() As Double, ByRef prev() As Long)
    Dim hn() As Long, hk() As Double, size As Long, cap As Long, v As Long, w As Long, k As Long, i As Long, c As Long, p As Long
    Dim d As Double, tk As Double, tn As Long
    For v = 0 To rn - 1: dist(v) = 1E+300: prev(v) = -1: Next v
    cap = adjStart(rn) + 2
    ReDim hn(1 To cap): ReDim hk(1 To cap)
    dist(src) = 0#: size = 1: hn(1) = src: hk(1) = 0#
    Do While size > 0
        v = hn(1): d = hk(1)
        hn(1) = hn(size): hk(1) = hk(size): size = size - 1
        i = 1
        Do
            c = 2 * i
            If c > size Then Exit Do
            If c < size Then
                If hk(c + 1) < hk(c) Then c = c + 1
            End If
            If hk(c) < hk(i) Then
                tk = hk(c): hk(c) = hk(i): hk(i) = tk: tn = hn(c): hn(c) = hn(i): hn(i) = tn: i = c
            Else
                Exit Do
            End If
        Loop
        If d <= dist(v) Then
            For k = adjStart(v) To adjStart(v + 1) - 1
                w = adjList(k)
                If d + adjW(k) < dist(w) Then
                    dist(w) = d + adjW(k): prev(w) = v
                    If size < cap Then
                        size = size + 1: hn(size) = w: hk(size) = dist(w): i = size
                        Do While i > 1
                            p = i \ 2
                            If hk(p) > hk(i) Then
                                tk = hk(p): hk(p) = hk(i): hk(i) = tk: tn = hn(p): hn(p) = hn(i): hn(i) = tn: i = p
                            Else
                                Exit Do
                            End If
                        Loop
                    End If
                End If
            Next k
        End If
    Loop
End Sub

Private Sub TraceSlipSurface(ByVal ws As Worksheet, ByVal xmin As Double, ByVal ymax As Double, ByVal factor As Double, ByRef jump() As Double, ByVal ref As Double, ByRef sel() As Boolean)
    Dim i As Long, k As Long, v As Long, w As Long, a As Long, b As Long, e As Long, m As Long, s As Shape
    Dim deg() As Long, fillN() As Long, adjStart() As Long, adjList() As Long, adjW() As Double
    Dim surf() As Boolean, nearSel() As Double, selNode() As Boolean, cand() As Long, candScore() As Double, nCand As Long
    Dim dist() As Double, prev() As Long, avgLen As Double, nSelE As Long, radius As Double, dx As Double, dy As Double, dd As Double
    Dim pS() As Long, pT() As Long, pAvg() As Double, pSep() As Double, nPair As Long, ci As Long, cj As Long
    Dim cost As Double, plen As Double, interior As Long, bestAvg As Double, maxSep As Double, pick As Long, tn As Long, tv As Double
    Dim pts() As Single, src As Long, dst As Long, fit As Boolean
    chainCount = 0: chainStartSurf = False: chainEndSurf = False: arcText = "": arcR = 0#
    ' 全辺の隣接（重み）
    ReDim deg(0 To rn - 1): ReDim fillN(0 To rn - 1): ReDim adjStart(0 To rn)
    For i = 0 To RNEdge - 1: deg(REdges(i).a) = deg(REdges(i).a) + 1: deg(REdges(i).b) = deg(REdges(i).b) + 1: Next i
    For v = 0 To rn - 1: adjStart(v + 1) = adjStart(v) + deg(v): Next v
    ReDim adjList(0 To 2 * RNEdge): ReDim adjW(0 To 2 * RNEdge)
    For i = 0 To RNEdge - 1
        a = REdges(i).a: b = REdges(i).b
        tv = REdges(i).length / (jump(i) / ref + 0.001)
        adjList(adjStart(a) + fillN(a)) = b: adjW(adjStart(a) + fillN(a)) = tv: fillN(a) = fillN(a) + 1
        adjList(adjStart(b) + fillN(b)) = a: adjW(adjStart(b) + fillN(b)) = tv: fillN(b) = fillN(b) + 1
    Next i
    ' 地表（拘束なしの外周辺、土要素）
    ReDim surf(0 To rn - 1): ReDim selNode(0 To rn - 1)
    For i = 0 To RNEdge - 1
        If REdges(i).element(1) < 0 Then
            If REdges(i).kind = "load" And RRigidId(REdges(i).element(0)) < 0 Then surf(REdges(i).a) = True: surf(REdges(i).b) = True
        End If
        If sel(i) Then selNode(REdges(i).a) = True: selNode(REdges(i).b) = True: avgLen = avgLen + REdges(i).length: nSelE = nSelE + 1
    Next i
    If nSelE = 0 Then Exit Sub
    avgLen = avgLen / nSelE: radius = 2.5 * avgLen
    ' 端点候補：選択辺の節点から radius 以内の地表節点（スコア＝近傍の最大ジャンプ）
    ReDim cand(0 To rn - 1): ReDim candScore(0 To rn - 1)
    For v = 0 To rn - 1
        If surf(v) Then
            tv = 0#
            For k = adjStart(v) To adjStart(v + 1) - 1
                w = adjList(k)
                If selNode(w) Or selNode(v) Then
                    tv = 1#: Exit For
                End If
            Next k
            If tv = 0# Then
                For w = 0 To rn - 1
                    If selNode(w) Then
                        dx = RXY(w, 0) - RXY(v, 0): dy = RXY(w, 1) - RXY(v, 1)
                        If dx * dx + dy * dy <= radius * radius Then tv = 0.5: Exit For
                    End If
                Next w
            End If
            If tv > 0# Then cand(nCand) = v: candScore(nCand) = tv: nCand = nCand + 1
        End If
    Next v
    Diag "推定すべり面: 地表端点候補=" & nCand
    If nCand < 2 Then Exit Sub
    If nCand > 30 Then
        ' 候補が多い場合は等間隔に間引く
        m = 0
        For i = 0 To 29: cand(i) = cand(Int(i * (nCand - 1) / 29#)): Next i
        nCand = 30
    End If
    ReDim dist(0 To rn - 1): ReDim prev(0 To rn - 1)
    ReDim pS(0 To nCand * nCand): ReDim pT(0 To nCand * nCand): ReDim pAvg(0 To nCand * nCand): ReDim pSep(0 To nCand * nCand)
    For ci = 0 To nCand - 2
        ShortestPaths cand(ci), adjStart, adjList, adjW, dist, prev
        For cj = ci + 1 To nCand - 1
            w = cand(cj)
            If dist(w) < 1E+299 Then
                plen = 0#: interior = 0: v = w
                Do While prev(v) >= 0
                    dx = RXY(v, 0) - RXY(prev(v), 0): dy = RXY(v, 1) - RXY(prev(v), 1): plen = plen + Sqr(dx * dx + dy * dy)
                    v = prev(v)
                    If Not surf(v) Then interior = interior + 1
                Loop
                If plen > 0# And interior > 0 Then
                    pS(nPair) = cand(ci): pT(nPair) = w: pAvg(nPair) = dist(w) / plen
                    dx = RXY(w, 0) - RXY(cand(ci), 0): dy = RXY(w, 1) - RXY(cand(ci), 1): pSep(nPair) = Sqr(dx * dx + dy * dy)
                    If pSep(nPair) > maxSep Then maxSep = pSep(nPair)
                    nPair = nPair + 1
                End If
            End If
        Next cj
        DoEvents
    Next ci
    Diag "推定すべり面: 端点ペア=" & nPair
    If nPair = 0 Then Exit Sub
    bestAvg = 1E+300
    For i = 0 To nPair - 1
        If pSep(i) >= 0.25 * maxSep And pAvg(i) < bestAvg Then bestAvg = pAvg(i)
    Next i
    pick = -1
    For i = 0 To nPair - 1
        If pSep(i) >= 0.25 * maxSep And pAvg(i) <= 1.5 * bestAvg Then
            If pick < 0 Then
                pick = i
            ElseIf pSep(i) > pSep(pick) Then
                pick = i
            End If
        End If
    Next i
    If pick < 0 Then Exit Sub
    src = pS(pick): dst = pT(pick)
    ShortestPaths src, adjStart, adjList, adjW, dist, prev
    v = dst: chainCount = 0
    Do While v >= 0: chainCount = chainCount + 1: v = prev(v): Loop
    chainStartSurf = surf(src): chainEndSurf = surf(dst)
    Diag "推定すべり面: 節点数=" & chainCount & " 平均コスト/長さ=" & Format(pAvg(pick), "0.000") & " 端点間距離=" & Format(pSep(pick), "0.00")
    If chainCount < 2 Then Exit Sub
    ReDim pts(1 To chainCount, 1 To 2): v = dst
    For k = 1 To chainCount
        pts(k, 1) = CSng(PX(RXY(v, 0), xmin, factor)): pts(k, 2) = CSng(PY(RXY(v, 1), ymax, factor)): v = prev(v)
    Next k
    Set s = ws.Shapes.AddPolyline(pts)
    s.AlternativeText = "RPX_VELOCITY": s.Fill.Visible = False
    s.Line.ForeColor.RGB = RGB(84, 20, 120): s.Line.Weight = 3.5
    s.ZOrder 0
    ' 円弧当てはめ（Kasa法）
    fit = True
    If Not OptionalSheet(SEEP_SHEET) Is Nothing Then fit = (SettingValue(OptionalSheet(SEEP_SHEET), ROW_ARCFIT, 1#) <> 0#)
    If fit And chainCount >= 4 Then DrawArcFit ws, xmin, ymax, factor, dst, src, prev
End Sub

Private Sub DrawArcFit(ByVal ws As Worksheet, ByVal xmin As Double, ByVal ymax As Double, ByVal factor As Double, ByVal dst As Long, ByVal src As Long, ByRef prev() As Long)
    Dim v As Long, n As Long, x As Double, y As Double, z As Double, sxx As Double, sxy As Double, syy As Double, sx As Double, sy As Double
    Dim sxz As Double, syz As Double, sz As Double, det As Double, dD As Double, dE As Double, dF As Double, mid As Long, k As Long
    Dim t0 As Double, t1 As Double, tm As Double, sweep As Double, pts() As Single, s As Shape, nPts As Long, pi As Double
    pi = 3.14159265358979
    v = dst
    Do While v >= 0
        x = RXY(v, 0): y = RXY(v, 1): z = x * x + y * y
        sxx = sxx + x * x: sxy = sxy + x * y: syy = syy + y * y: sx = sx + x: sy = sy + y
        sxz = sxz + x * z: syz = syz + y * z: sz = sz + z: n = n + 1
        v = prev(v)
    Loop
    ' [sxx sxy sx; sxy syy sy; sx sy n] [D E F]' = -[sxz syz sz]'
    det = sxx * (syy * n - sy * sy) - sxy * (sxy * n - sy * sx) + sx * (sxy * sy - syy * sx)
    If Abs(det) < 1E-12 Then Exit Sub
    dD = (-sxz * (syy * n - sy * sy) - sxy * (-syz * n + sy * sz) + sx * (-syz * sy + syy * sz)) / det
    dE = (sxx * (-syz * n + sy * sz) + sxz * (sxy * n - sy * sx) + sx * (-sxy * sz + syz * sx)) / det
    dF = (sxx * (-syy * sz + sy * syz) - sxy * (-sxy * sz + sx * syz) - sxz * (sxy * sy - syy * sx)) / det
    arcCx = -dD / 2#: arcCy = -dE / 2#
    z = arcCx * arcCx + arcCy * arcCy - dF
    If z <= 0# Then Exit Sub
    arcR = Sqr(z)
    arcText = "／円弧当てはめ：中心(" & Format(arcCx, "0.00") & ", " & Format(arcCy, "0.00") & ") 半径 " & Format(arcR, "0.00") & " m"
    Diag "円弧: 中心=(" & arcCx & "," & arcCy & ") 半径=" & arcR
    ' 端点間を中間節点側に回る円弧
    mid = dst
    For k = 1 To n \ 2: mid = prev(mid): Next k
    t0 = Atan2(RXY(dst, 1) - arcCy, RXY(dst, 0) - arcCx): t1 = Atan2(RXY(src, 1) - arcCy, RXY(src, 0) - arcCx)
    tm = Atan2(RXY(mid, 1) - arcCy, RXY(mid, 0) - arcCx)
    sweep = t1 - t0: If sweep < 0# Then sweep = sweep + 2# * pi
    If Not AngleBetween(t0, sweep, tm, pi) Then sweep = sweep - 2# * pi
    nPts = 60: ReDim pts(1 To nPts + 1, 1 To 2)
    For k = 0 To nPts
        x = arcCx + arcR * Cos(t0 + sweep * k / nPts): y = arcCy + arcR * Sin(t0 + sweep * k / nPts)
        pts(k + 1, 1) = CSng(PX(x, xmin, factor)): pts(k + 1, 2) = CSng(PY(y, ymax, factor))
    Next k
    Set s = ws.Shapes.AddPolyline(pts)
    s.AlternativeText = "RPX_VELOCITY": s.Fill.Visible = False
    s.Line.ForeColor.RGB = RGB(192, 0, 0): s.Line.Weight = 1.75: s.Line.DashStyle = 4
    s.ZOrder 0
End Sub

Private Function Atan2(ByVal y As Double, ByVal x As Double) As Double
    If x > 0# Then
        Atan2 = Atn(y / x)
    ElseIf x < 0# Then
        If y >= 0# Then Atan2 = Atn(y / x) + 3.14159265358979 Else Atan2 = Atn(y / x) - 3.14159265358979
    ElseIf y > 0# Then
        Atan2 = 1.5707963267949
    ElseIf y < 0# Then
        Atan2 = -1.5707963267949
    End If
End Function

Private Function AngleBetween(ByVal t0 As Double, ByVal sweep As Double, ByVal t As Double, ByVal pi As Double) As Boolean
    Dim d As Double
    d = t - t0
    Do While d < 0#: d = d + 2# * pi: Loop
    Do While d >= 2# * pi: d = d - 2# * pi: Loop
    AngleBetween = (d <= sweep)
End Function

' 凡例（A2:A6）を図の下へ移す。図形の最下端＋10pt より下の行に「【凡例】」見出しと各行を書く。
' A2:A6 は内容を消し、文字色を白にする（既存コードが後から A3 等に書いても図に重ならない）。
Private Sub RelocateLegend(ByVal ws As Worksheet)
    Dim lines(1 To 6) As String, n As Long, r As Long, i As Long, sh As Shape, bottom As Double, t As String, found As Range, startRow As Long
    On Error Resume Next
    For r = 2 To 6
        t = CStr(ws.Cells(r, 1).Value2)
        If Len(t) > 0 Then n = n + 1: lines(n) = t
    Next r
    If n = 0 Then Exit Sub
    t = CStr(ws.Cells(1, 1).Value2)
    bottom = 55# + 460#
    For Each sh In ws.Shapes
        If sh.Top + sh.Height > bottom Then bottom = sh.Top + sh.Height
    Next sh
    startRow = 7
    Do While ws.Rows(startRow).Top < bottom + 10# And startRow < 5000
        startRow = startRow + 1
    Loop
    ' 旧凡例ブロックを消す
    Set found = ws.Range("A7:A5000").Find(What:=LEGEND_MARK, LookIn:=-4163, LookAt:=1)
    If Not found Is Nothing Then ws.Range(ws.Cells(found.Row, 1), ws.Cells(found.Row + 8, 1)).ClearContents
    ws.Cells(startRow, 1).Value2 = LEGEND_MARK
    ws.Cells(startRow, 1).Font.Bold = True
    If Len(t) > 0 Then ws.Cells(startRow + 1, 1).Value2 = t: startRow = startRow + 1
    For i = 1 To n
        ws.Cells(startRow + i, 1).Value2 = lines(i)
    Next i
    ws.Range(ws.Cells(startRow - 1, 1), ws.Cells(startRow + n, 1)).Font.Color = RGB(0, 0, 0)
    ws.Range("A2:A6").ClearContents
    ws.Range("A2:A6").Font.Color = RGB(255, 255, 255)
    Err.Clear
End Sub

' RPX_DrawCore の末尾から呼ばれる（フック）。描画の失敗で本体を止めない。
Public Sub RPX_SeepOverlay()
    ' 何が起きても呼び出し元（RPX_DrawCore / RPX_Run）へエラーを返さない。
    Dim ws As Worksheet, previousWrite As Boolean
    On Error Resume Next
    previousWrite = RPX_InternalWrite
    RPX_SeepTrace "overlay begin rn=" & rn & " re=" & re & " result=" & RPX_ResultCurrent & " mesh=" & RPX_MeshCurrent
    Set ws = ThisWorkbook.Worksheets("図")
    If Not ws Is Nothing Then
        If ws.ProtectContents Or ws.ProtectDrawingObjects Then
            RPX_SeepTrace "overlay skip: 図 is protected"
        Else
            RPX_InternalWrite = True
            OverlayOn ws
        End If
    End If
    RPX_InternalWrite = previousWrite
    RPX_SeepTrace "overlay end err=" & lastOverlayError
    Err.Clear
End Sub

Private Sub OverlayOn(ByVal ws As Worksheet)
    Dim inp As Worksheet, xmin As Double, ymax As Double, factor As Double, stage As String
    Dim slipThr As Double, drawStrain As Boolean, strainThr As Double, drawFs As Boolean, note As String
    On Error GoTo Quiet
    stage = "init": overlayDiag = "": lastOverlayError = ""
    If Len(CStr(ws.Range("A6").Value2)) > 0 Then ws.Range("A6").Value2 = ""
    DeleteTagged ws, "RPX_SEEP"
    If Not DrawExtent(xmin, ymax, factor) Then Diag "DrawExtent=False（メッシュなし）": Exit Sub
    Diag "DrawExtent: xmin=" & xmin & " ymax=" & ymax & " factor=" & Format(factor, "0.000")
    Set inp = OptionalSheet(SEEP_SHEET)
    slipPct = 0.95: slipMaxFrac = 0.1
    drawFs = True: slipThr = 0.5: drawStrain = False: strainThr = 0.2
    If Not inp Is Nothing Then
        stage = "settings"
        drawFs = (SettingValue(inp, ROW_DRAWFS, 1#) <> 0#)
        slipThr = SettingValue(inp, ROW_SLIPTHR, 0.5)
        slipPct = SettingValue(inp, ROW_SLIPPCT, 0.95): slipMaxFrac = SettingValue(inp, ROW_SLIPMAXFRAC, 0.1)
        drawStrain = (SettingValue(inp, ROW_DRAWSTRAIN, 0#) <> 0#)
        strainThr = SettingValue(inp, ROW_STRAINTHR, 0.2)
    End If
    Diag "設定: N15自由水面=" & drawFs & " N16しきい値=" & slipThr & " N21分位点=" & slipPct & " N22最大割合=" & slipMaxFrac & " N17ひずみ=" & drawStrain & IIf(inp Is Nothing, "（浸透入力なし→既定値）", "")
    Diag "ResultCurrent=" & RPX_ResultCurrent & " MeshCurrent=" & RPX_MeshCurrent & " re=" & re & " RR=" & RR & " RNEdge=" & RNEdge
    stage = "free_surface"
    If drawFs Then DrawFreeSurface ws, xmin, ymax, factor: note = "青太線：自由水面（u=0）"
    stage = "mechanism"
    Diag "UpperFieldReady=" & UpperFieldReady()
    If UpperFieldReady() Then
        stage = "slip"
        DrawMechanism ws, xmin, ymax, factor, slipThr, drawStrain, strainThr
        If slipThr > 0# Then note = note & IIf(Len(note) > 0, "　", "") & "紫線：上界の速度ジャンプ大の辺、濃紫太線：連結した推定すべり面"
        If drawStrain Then note = note & IIf(Len(note) > 0, "　", "") & "橙：塑性せん断ひずみ速度大"
    End If
    If CStr(ws.Range("A4").Value2) <> note Then ws.Range("A4").Value2 = note
    RelocateLegend ws
    Exit Sub
Quiet:
    lastOverlayError = "重ね描きエラー（" & stage & "）：" & Err.Number & " " & Err.Description
    Diag lastOverlayError
    Resume QuietEnd
QuietEnd:
    On Error Resume Next
    ws.Range("A6").Value2 = lastOverlayError
    Debug.Print lastOverlayError
    RelocateLegend ws
    Err.Clear
End Sub

' 合成データで重ね描きを検査する。極限解析の状態は保存・復元し、「図_テスト」シートに描く。
Public Function RPX_SeepOverlaySelfTest() As String
    Dim st As SeepSnapshot, savedRR As Long, savedField() As Double, hadField As Boolean, savedMesh As Boolean, savedResult As Boolean
    Dim ws As Worksheet, nx As Long, ny As Long, i As Long, j As Long, e As Long, k As Long, a As Long, b As Long, c As Long, d As Long
    Dim px As Double, py As Double, cx As Double, cy As Double, inside As Boolean, msg As String, caseNo As Long, cnt As Long, poly As Long
    Dim previousWrite As Boolean
    If RPX_Busy Then RPX_SeepOverlaySelfTest = "実行中のため中止": Exit Function
    previousWrite = RPX_InternalWrite: RPX_InternalWrite = True
    SaveState st: savedRR = RR: savedMesh = RPX_MeshCurrent: savedResult = RPX_ResultCurrent
    On Error Resume Next
    savedField = RUpperField: hadField = (Err.Number = 0): Err.Clear
    On Error GoTo Failed
    Set ws = RPX_Sheet("図_テスト")
    nx = 40: ny = 20
    rn = (nx + 1) * (ny + 1): re = 2 * nx * ny: RR = 0
    ReDim RXY(0 To rn - 1, 0 To 1): ReDim RTri(0 To re - 1, 0 To 2): ReDim RMatId(0 To re - 1): ReDim RRigidId(0 To re - 1)
    ReDim RBody0(0 To re - 1, 0 To 1): ReDim RBody1(0 To re - 1, 0 To 1)
    For e = 0 To re - 1: RRigidId(e) = -1: RMatId(e) = 0: Next e
    For j = 0 To ny
        For i = 0 To nx: RXY(j * (nx + 1) + i, 0) = 0.5 * i: RXY(j * (nx + 1) + i, 1) = 0.5 * j: Next i
    Next j
    e = 0
    For j = 0 To ny - 1
        For i = 0 To nx - 1
            a = j * (nx + 1) + i: b = a + 1: c = b + nx + 1: d = a + nx + 1
            RTri(e, 0) = a: RTri(e, 1) = b: RTri(e, 2) = c: e = e + 1
            RTri(e, 0) = a: RTri(e, 1) = c: RTri(e, 2) = d: e = e + 1
        Next i
    Next j
    Set RBoundary = Nothing
    Call RPX_Topology
    RPX_MeshCurrent = True: RPX_ResultCurrent = True
    For caseNo = 1 To 4
        ReDim RUpperField(0 To 12 * re - 1)
        For e = 0 To re - 1
            cx = (RXY(RTri(e, 0), 0) + RXY(RTri(e, 1), 0) + RXY(RTri(e, 2), 0)) / 3#
            cy = (RXY(RTri(e, 0), 1) + RXY(RTri(e, 1), 1) + RXY(RTri(e, 2), 1)) / 3#
            inside = ((cx - 16#) ^ 2 + (cy - 16#) ^ 2 < 11# ^ 2)
            For k = 0 To 5
                RPX_VelocityPosition e, k, px, py
                If (caseNo = 1 Or caseNo = 4) And inside Then
                    RUpperField(12 * e + 2 * k) = -(py - 16#): RUpperField(12 * e + 2 * k + 1) = (px - 16#)
                ElseIf caseNo = 3 And e = 0 Then
                    RUpperField(12 * e + 2 * k) = 1000#
                End If
            Next k
        Next e
        For i = ws.Shapes.Count To 1 Step -1: ws.Shapes(i).Delete: Next i
        lastOverlayError = ""
        OverlayOn ws
        cnt = 0: poly = 0
        For i = 1 To ws.Shapes.Count
            If ws.Shapes(i).AlternativeText = "RPX_VELOCITY" Then
                cnt = cnt + 1
                If ws.Shapes(i).Line.Weight >= 3.4 Then poly = poly + 1
            End If
        Next i
        If caseNo = 4 Then
            msg = msg & "ケース4判定：連鎖節点 " & chainCount & "、始点地表=" & chainStartSurf & "、終点地表=" & chainEndSurf & _
                  "、円 中心(" & Format(arcCx, "0.00") & "," & Format(arcCy, "0.00") & ") 半径 " & Format(arcR, "0.00") & "（真値 16,16 / 11）→ " & _
                  IIf(chainCount > 15 And chainStartSurf And chainEndSurf, "合格", "不合格") & vbLf
        End If
        msg = msg & "ケース" & caseNo & Choose(caseNo, "（円弧すべり）", "（速度ゼロ）", "（1要素だけ巨大ジャンプ）", "（円弧すべり：連続すべり面）") & "：図形 " & cnt & "、太線 " & poly & IIf(Len(lastOverlayError) > 0, "、エラー " & lastOverlayError, "、エラーなし") & vbLf & overlayDiag
    Next caseNo
    GoTo Restore
Failed:
    msg = msg & "自己テスト失敗：" & Err.Number & " " & Err.Description & vbLf
    Resume Restore
Restore:
    On Error Resume Next
    RestoreState st: RR = savedRR: RPX_MeshCurrent = savedMesh: RPX_ResultCurrent = savedResult
    If hadField Then RUpperField = savedField Else Erase RUpperField
    RPX_InternalWrite = previousWrite
    msg = msg & "期待：ケース1・4は図形>0かつ太線1、ケース2は図形0、ケース3は図形>0でエラーなし、ケース4は合格"
    Debug.Print msg
    RPX_SeepOverlaySelfTest = msg
End Function

' 浸透入力シートの「図を表示」：極限解析メッシュがあれば本体の図を描き直す（フック経由で重ね描き）。
Public Sub RPX_SeepShowDiagram()
    Dim ws As Worksheet
    On Error Resume Next
    If RPX_MeshCurrent And Not RPX_Busy Then
        Call RPX_Draw
    Else
        Set ws = RPX_Sheet("図")
        DeleteTagged ws, "RPX_SEEP"
        Call RPX_SeepOverlay
    End If
    RPX_Sheet("図").Activate
End Sub
