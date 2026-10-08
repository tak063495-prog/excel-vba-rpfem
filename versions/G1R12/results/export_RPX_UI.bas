Option Explicit

Public RPX_Busy As Boolean, RPX_InternalWrite As Boolean
Public RPX_MeshCurrent As Boolean, RPX_ResultCurrent As Boolean
Public RPX_FaceIDs As Object
Public RPX_RigidIDs As Object
Public Function RPX_Sheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set RPX_Sheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If RPX_Sheet Is Nothing Then
        Set RPX_Sheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count)): RPX_Sheet.name = name
    End If
End Function
Private Sub Header(ByVal ws As Worksheet, ByVal address As String, ByVal labels As Variant)
    Dim i As Long
    For i = LBound(labels) To UBound(labels): ws.Range(address).offset(0, i).Value2 = labels(i): Next i
    With ws.Range(address).Resize(1, UBound(labels) - LBound(labels) + 1)
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255): .Interior.Color = RGB(38, 67, 91): .WrapText = True: .RowHeight = 38
    End With
End Sub
Private Sub Button(ByVal ws As Worksheet, ByVal label As String, ByVal macro As String, ByVal x As Double, ByVal y As Double, ByVal width As Double)
    Dim s As shape
    Set s = ws.Shapes.AddShape(5, x, y, width, 28)
    s.TextFrame2.TextRange.text = label: s.TextFrame2.TextRange.Font.size = 11: s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    s.Fill.ForeColor.RGB = RGB(36, 104, 130): s.line.Visible = False: s.OnAction = macro
End Sub
Private Sub Choice(ByVal cells As Range, ByVal choices As String)
    With cells.Validation
        .Delete: .Add xlValidateList, xlValidAlertStop, xlBetween, choices
        .IgnoreBlank = True: .InCellDropdown = True: .ShowError = True
        .ErrorTitle = "入力を確認してください": .ErrorMessage = "一覧から選択してください。"
    End With
End Sub
Public Sub RPX_Setup()
    Dim ws As Worksheet, name As Variant, i As Long, labels As Variant, values As Variant, keys As Variant
    RPX_InternalWrite = True
    For Each name In Array("要素定義", "材料データ", "設定", "解析結果", "図", "節点データ", "要素データ")
        Set ws = RPX_Sheet(CStr(name)): ws.cells.UnMerge: ws.cells.Clear
        ws.rows.Hidden = False: ws.columns.Hidden = False: ws.rows.RowHeight = 20: ws.rows(1).RowHeight = 34
        For i = ws.Shapes.count To 1 Step -1: ws.Shapes(i).Delete: Next i
        ws.cells.Font.name = "Yu Gothic": ws.cells.Font.size = 11: ws.columns.ColumnWidth = 12
    Next name
    Set ws = RPX_Sheet("要素定義")
    ws.Range("A1").Value2 = "剛塑性FEM　要素定義": ws.Range("A1").Font.size = 20: ws.Range("A1").Font.Bold = True
    ws.Range("A3").Value2 = "定義点・領域・面条件を入力後、要素作成②でメッシュを生成します。"
    ws.Range("A4").Value2 = "面条件は領域の定義順に開始点から終了点まで適用します。途中の定義点と曲がり角を含みます。"
    Header ws, "A6", Array("点ID", "X [m]", "Y [m]")
    Header ws, "E6", Array("領域ID", "材料ID", "剛体ID", "境界点ID（定義順）")
    Header ws, "J6", Array("面ID", "領域ID", "開始点", "終了点", "循環", "拘束", "剛体ID", "固定Fx [kPa]", "固定Fy [kPa]", "比例Fx [kPa]", "比例Fy [kPa]")
    ws.columns("D").ColumnWidth = 3: ws.columns("I").ColumnWidth = 3: ws.columns("H").ColumnWidth = 34
    ws.columns("A:G").ColumnWidth = 10: ws.columns("D").ColumnWidth = 3: ws.columns("H").ColumnWidth = 34
    ws.columns("J:Q").ColumnWidth = 9: ws.columns("R:T").ColumnWidth = 13
    ws.Range("A7:C1006,E7:H1006,J7:T1006").Interior.Color = RGB(243, 248, 251)
    Choice ws.Range("N7:N1006"), "0,1": Choice ws.Range("O7:O1006"), "自由,固定XY,固定X,固定Y,剛体"
    Button ws, "要素作成②", "RPX_CreateMesh", 340, 5, 100
    Button ws, "解析", "RPX_Run", 450, 5, 80
    Button ws, "図を表示", "RPX_ShowDiagram", 540, 5, 90
    Button ws, "中止", "RPX_Stop", 640, 5, 80
    Set ws = RPX_Sheet("材料データ")
    ws.Range("A1").Value2 = "材料・剛体": ws.Range("A1").Font.size = 20
    Header ws, "A4", Array("材料ID", "c [kPa]", "φ [度]", "γ [kN/m^3]")
    Header ws, "F4", Array("剛体ID", "中心X", "中心Y", "X固定", "Y固定", "回転固定", "固定Fx", "固定Fy", "固定M", "比例Fx", "比例Fy", "比例M")
    ws.Range("A2").Value2 = "関連流れ・全応力。剛体の荷重は単位奥行き当たり。固定は1、自由は0。"
    ws.Range("A5:D1004,F5:Q1004").Interior.Color = RGB(243, 248, 251)
    Set ws = RPX_Sheet("設定")
    ws.Range("A1").Value2 = "解析設定": ws.Range("A1").Font.size = 20
    Header ws, "A3", Array("項目", "値", "単位・説明", "キー")
    labels = Array("解析種別", "目標要素数", "自重の扱い", "メッシュ最適化", "最大最適化回数", "目標区間幅", "下界の局所細分化率", "散逸積分の細分数", "Fs探索許容差", "支持圧集計の面ID", "支持圧集計の剛体ID")
    values = Array("Fs", 1536, "固定", 1, 3, 0.01, 0.2, 8, 0.0001, "", "")
    keys = Array("MODE", "TARGET", "GRAVITY", "ADAPT", "CYCLES", "GAP", "REFINE", "QUADRATURE", "FS_TOL", "BEARING_FACES", "BEARING_RIGIDS")
    For i = 0 To UBound(labels)
        ws.cells(i + 4, 1).Value2 = labels(i): ws.cells(i + 4, 2).Value2 = values(i): ws.cells(i + 4, 4).Value2 = keys(i)
    Next i
    ws.Range("C4").Value2 = "Fs / 極限支持力": ws.Range("C6").Value2 = "固定 / 比例（Fsでは実荷重を維持）"
    ws.Range("C7").Value2 = "1=有効、0=無効": ws.Range("C9").Value2 = "上下界差 / 中点"
    ws.Range("C10").Value2 = "高い誤差指標を持つ要素の割合": ws.Range("C13").Value2 = "極限支持圧の集計対象。複数IDはカンマ区切り。"
    ws.Range("C14").Value2 = "任意：追加集計する剛体外力。幅は上の選択面から求めます。"
    ws.cells(15, 1).Value2 = "上界パイロット": ws.cells(15, 2).Value2 = 0
    ws.cells(15, 3).Value2 = "0=従来、1=初期上界だけで帯探索（Fs・使用中の土が単一・剛体なし）"
    ws.cells(15, 4).Value2 = "UPPER_PILOT"
    ws.Range("B9:B10").NumberFormat = "0.0%": ws.Range("B4:B15").Interior.Color = RGB(243, 248, 251)
    Choice ws.Range("B15"), "0,1"
    Choice ws.Range("B4"), "Fs,極限支持力": Choice ws.Range("B6"), "固定,比例": Choice ws.Range("B7"), "0,1"
    RPX_EnsureUpperPilotSetting
    ws.columns("A").ColumnWidth = 29: ws.columns("B").ColumnWidth = 18: ws.columns("C").ColumnWidth = 58: ws.columns("D").Hidden = True
    Set ws = RPX_Sheet("解析結果"): ws.Range("A1").Value2 = "未解析": ws.Range("A1").Font.size = 20
    ws.columns("A").ColumnWidth = 34: ws.columns("B:D").ColumnWidth = 20
    Call RPX_WriteSlopeInputs
    RPX_InternalWrite = False: RPX_Sheet("要素定義").Activate
End Sub
Private Sub EnsureSettingRow(ByVal ws As Worksheet, ByVal key As String, ByVal label As String, ByVal value As Variant, ByVal note As String, ByVal choices As String)
    Dim row As Long, emptyRow As Long, found As Long
    found = 0
    emptyRow = 0
    For row = 4 To 100
        If CStr(ws.cells(row, 4).Value2) = key Then
            found = row
            Exit For
        End If
        If emptyRow = 0 Then
            If Len(CStr(ws.cells(row, 1).Value2)) = 0 And Len(CStr(ws.cells(row, 4).Value2)) = 0 Then emptyRow = row
        End If
    Next row
    If found > 0 Then Exit Sub
    If emptyRow = 0 Then Exit Sub
    ws.cells(emptyRow, 1).Value2 = label
    ws.cells(emptyRow, 2).Value2 = value
    ws.cells(emptyRow, 3).Value2 = note
    ws.cells(emptyRow, 4).Value2 = key
    ws.cells(emptyRow, 2).Interior.Color = RGB(243, 248, 251)
    If Len(choices) > 0 Then Choice ws.cells(emptyRow, 2), choices
End Sub
Public Sub RPX_EnsureUpperPilotSetting()
    Dim ws As Worksheet
    Set ws = RPX_Sheet("設定")
    EnsureSettingRow ws, "ANALYSIS_POLICY", "解析政策", "LEGACY_DAVIS_CLIP", "既存Davis / 関連 / Davis等価関連", "LEGACY_DAVIS_CLIP,ASSOCIATED,DAVIS_EQUIVALENT"
    EnsureSettingRow ws, "C0_DIRECT_UPPER", "c0 direct candidate", 1, "Audited straight-face single-material selfweight candidate; otherwise use original upper solve", "0,1"
    EnsureSettingRow ws, "RAW_LOWER_AUDIT", "元単位下界監査", 1, "元単位の降伏・要素平衡・表面力を追加監査。剛体力/モーメントは既存規準", "0,1"
    EnsureSettingRow ws, "UPPER_PILOT", "上界パイロット", 0, "0=従来、1=初期上界だけで帯探索（Fs・単一土・剛体なし）", "0,1"
    EnsureSettingRow ws, "UPPER_PILOT_RIGID", "剛体上界パイロット", 0, "0=従来適応。1=剛体・界面を保護する上界pilot", "0,1"
    EnsureSettingRow ws, "LOWER_FS_SEED", "下界初期試行", "LEGACY", "LEGACY=従来探索。UPPER_TRIAL=同一メッシュ上界から下界試行", "LEGACY,UPPER_TRIAL"
    EnsureSettingRow ws, "LOWER_SEED_WIDTH", "下界試行幅", 0.02, "試行区間=[上界Fs*(1-幅), 上界Fs]。標準0.02", ""
    EnsureSettingRow ws, "LOWER_MODEL_CACHE", "下界構造再利用", 0, "0=無効。1=同一モデルの下界圧縮構造を再利用（冷開始）", "0,1"
    EnsureSettingRow ws, "LOWER_KKT_BACKEND", "下界KKT解法", "GENERIC", "GENERIC=汎用、COMPARE=診断比較、AUTO=適格な下界Schur。失敗時は汎用", "GENERIC,COMPARE,AUTO"
    EnsureSettingRow ws, "KKT_CAPTURE", "KKT診断保存", 0, "1=初期・中期・最後の実KKTと両方向を保存（追加時間あり）", "0,1"
    EnsureSettingRow ws, "SCHUR_ALLOW_RIGID", "Schur剛体許可", 0, "0=剛体は汎用。1=反力を残すSchurを許可", "0,1"
    EnsureSettingRow ws, "SCHUR_MAX_BYTES", "Schur作業上限", 0, "追加作業メモリの上限バイト。0=256 MiB", ""
    EnsureSettingRow ws, "SCHUR_TRACE_LEVEL", "Schur詳細診断", "0", "0=通常、1=方向、2=残差履歴と復元式比較（追加時間あり）", "0,1,2"
    EnsureSettingRow ws, "SCHUR_REPLAY_POLICY", "方向ペアの再計算", "LEGACY", "LEGACY=従来、CONSISTENT=必要な場合のみ汎用ペア再計算", "LEGACY,CONSISTENT"
    EnsureSettingRow ws, "SCHUR_STRUCTURE_CACHE", "Schur構造の保持", "0", "1=厳密一致する下界構造を上界計算中も保持", "0,1"
    EnsureSettingRow ws, "SCHUR_RECOVERY_MODE", "応力方向の復元", "RESOLVE", "RESOLVE=従来、REUSE_U=既存の局所解を利用", "RESOLVE,REUSE_U"
    EnsureSettingRow ws, "SCHUR_REFINEMENT", "Schur反復改良", "FIXED3", "FIXED3=従来の3回補正", "FIXED3"
    EnsureSettingRow ws, "SCHUR_GENERIC_START_ITERS", "初期汎用反復数", "0", "0のみ。この版では初期反復の変更は対象外", "0"
    EnsureSettingRow ws, "DIRECT_FINAL_AFTER_PILOT", "pilot後の直接最終評価", "0", "0のみ。この版では適応手順の変更は対象外", "0"
    EnsureSettingRow ws, "HSD_PROJECTIVE_NORMALIZATION", "HSD同次尺度の保持", 1, "2の冪で全同次変数を共通尺度へ戻す。原条件監査は維持", "0,1"
    EnsureSettingRow ws, "HSD_ORIGINAL_OPERATOR", "補正前の線形系監査", 0, "1=因子の補正を除いた原線形系で反復改良。実験設定", "0,1"
    EnsureSettingRow ws, "HSD_SOLVE_SECONDS", "HSD求解予算", 120, "一回のHSD求解の秒数上限。反復境界で検査。超過は未判定", ""
    EnsureSettingRow ws, "HSD_LINEAR_BACKEND", "HSD linear backend", "AUTO_PIVOT", "Checked HSD linear solver", "LDL_CHECKED,PIVOT_ONLY,AUTO_PIVOT"
    EnsureSettingRow ws, "HSD_LINEAR_CAPTURE", "HSD linear capture", 0, "Native Newton diagnostics", "0,1"
    EnsureSettingRow ws, "HSD_LINEAR_CAPTURE_MAX", "HSD capture limit", 8, "Maximum records per solve", ""
    EnsureSettingRow ws, "HSD_STABLE_MAX_BYTES", "HSD memory budget", 268435456, "Stable solver memory limit", ""
    EnsureSettingRow ws, "HSD_LINEAR_TRACE", "HSD linear trace", 1, "0=off, 1=summary, 2=detail", "0,1,2"
    RPX_EnsureDiagnosticSettings
    EnsureSettingRow ws, "TEST_INJECTION", "試験注入", 0, "0=無効。1は明示的な試験入口からの注入を許可", "0,1"
End Sub
Public Sub RPX_EnsureDiagnosticSettings()
    Dim ws As Worksheet, oldValue As Variant, value As Long
    Set ws = RPX_Sheet("設定"): value = 1
    If Len(CStr(ws.cells(15, 4).Value2)) = 0 Then
        oldValue = ws.cells(15, 2).Value2
        If Not IsError(oldValue) Then
            If IsNumeric(oldValue) Then
                If CDbl(oldValue) = 0# Or CDbl(oldValue) = 1# Then value = CLng(oldValue)
            End If
        End If
    End If
    EnsureSettingRow ws, "DIAGNOSTIC_MODE", "診断ログ", value, "0=OFF、1=ON。上界pilotとは独立", "0,1"
    EnsureSettingRow ws, "HSD_CAPTURE", "HSD候補の保存", 0, "1=最初の不合格候補とモデルを保存（追加時間あり）", "0,1"
End Sub
Public Sub RPX_ConfigureLegacyDiagnostics(Optional ByVal captureHsd As Boolean = False)
    Dim r As Long, ws As Worksheet
    Set ws = RPX_Sheet("設定")
    RPX_EnsureDiagnosticSettings
    For r = 4 To 100
        If ws.cells(r, 4).Value2 = "DIAGNOSTIC_MODE" Then ws.cells(r, 2).Value2 = 1
        If ws.cells(r, 4).Value2 = "HSD_CAPTURE" Then ws.cells(r, 2).Value2 = Abs(CLng(captureHsd))
    Next r
End Sub
Public Function RPX_Setting(ByVal key As String, ByVal fallback As Variant) As Variant
    Dim row As Long, ws As Worksheet
    Set ws = RPX_Sheet("設定")
    For row = 4 To 100
        If CStr(RPX_InputCell(ws, row, 4)) = key Then
            RPX_Setting = RPX_InputCell(ws, row, 2): Exit Function
        End If
    Next row
    RPX_Setting = fallback
End Function
Public Sub RPX_WriteSlopeInputs()
    Dim ws As Worksheet, i As Long
    RPX_SlopeDefinition 1536: Set ws = RPX_Sheet("要素定義")
    For i = 0 To 5
        ws.cells(i + 7, 1).Value2 = i + 1: ws.cells(i + 7, 2).Value2 = DPXY(i, 0): ws.cells(i + 7, 3).Value2 = DPXY(i, 1)
    Next i
    ws.Range("E7:H7").Value2 = Array(1, 1, 0, "1,2,3,4,5,6")
    ws.Range("J7:T7").Value2 = Array(1, 1, 1, 2, 0, "固定XY", 0, 0#, 0#, 0#, 0#)
    ws.Range("J8:T8").Value2 = Array(2, 1, 5, 6, 0, "固定XY", 0, 0#, 0#, 0#, 0#)
    ws.Range("J9:T9").Value2 = Array(3, 1, 6, 1, 1, "固定XY", 0, 0#, 0#, 0#, 0#)
    RPX_Sheet("材料データ").Range("A5:D5").Value2 = Array(1, 41.65, 15#, 18.82)
End Sub
Private Function number(ByVal v As Variant, ByVal label As String) As Double
    If IsError(v) Then err.Raise 5, , label & "がExcelのエラー値です。"
    If IsEmpty(v) Or Len(CStr(v)) = 0 Or Not IsNumeric(v) Then err.Raise 5, , label & "が数値ではありません。"
    number = CDbl(v)
End Function
Private Function IDNumber(ByVal v As Variant, ByVal label As String) As Long
    Dim x As Double
    x = number(v, label)
    If x <> Fix(x) Or x < 0# Or x > 1000000# Then err.Raise 5, , label & "が整数IDではありません。"
    IDNumber = CLng(x)
End Function
Private Function Restraint(ByVal text As String) As String
    Select Case text
        Case "自由", "load", "": Restraint = "load"
        Case "固定XY", "fixed": Restraint = "fixed"
        Case "固定X", "roller_x": Restraint = "roller_x"
        Case "固定Y", "roller_y": Restraint = "roller_y"
        Case "剛体", "rigid": Restraint = "rigid"
        Case Else: err.Raise 5, , "拘束の指定を確認してください。"
    End Select
End Function
Public Sub RPX_ReadInputs()
    Dim ws As Worksheet, pointMap As Object, regionMap As Object, materialMap As Object, rigidMap As Object, faceMap As Object
    Dim r As Long, i As Long, j As Long, id As Long, last As Long, text As String, parts As Variant, value As Double
    RPX_AssertInputs "read_inputs"
    value = RPX_DiagnosticMode()
    value = number(RPX_Setting("HSD_CAPTURE", 0), "HSD_CAPTURE")
    If value <> 0# And value <> 1# Then err.Raise 5, , "INVALID_SETTING HSD_CAPTURE"
    Set pointMap = RPX_Row(): Set regionMap = RPX_Row(): Set materialMap = RPX_Row(): Set rigidMap = RPX_Row(): Set faceMap = RPX_Row()
    RProgramKind = "": Set ws = RPX_Sheet("材料データ"): RM = 0: RR = 0
    ReDim RMaterials(0 To 999): ReDim RRigids(0 To 999): RStressScale = 0#
    For r = 5 To 1004
        If Len(CStr(RPX_InputCell(ws, r, 1))) > 0 Then
            id = IDNumber(RPX_InputCell(ws, r, 1), "材料ID")
            If id = 0 Or materialMap.Exists(id) Then err.Raise 5, , "材料IDが重複しています。"
            materialMap(id) = RM
            With RMaterials(RM)
            
                .cohesion = number(RPX_InputCell(ws, r, 2), "c")
                .friction = number(RPX_InputCell(ws, r, 3), "phi")
                .gamma = number(RPX_InputCell(ws, r, 4), "gamma")
                If Len(CStr(RPX_InputCell(ws, r, 5))) = 0 Then
                    .dilation = .friction
                Else
                    .dilation = number(RPX_InputCell(ws, r, 5), "psi")
                End If
                If .cohesion < 0# Or .friction < 0# Or .friction >= 89# Or .gamma < 0# Then err.Raise 5, , "MATERIAL_RANGE"
                If .dilation < 0# Or .dilation > .friction Then err.Raise 5, , "DILATION_RANGE"
                
                
                If .cohesion > RStressScale Then RStressScale = .cohesion
                
                
            End With
            RM = RM + 1
        End If
        If Len(CStr(RPX_InputCell(ws, r, 6))) > 0 Then
            id = IDNumber(RPX_InputCell(ws, r, 6), "剛体ID")
            If id = 0 Or rigidMap.Exists(id) Then err.Raise 5, , "剛体IDが重複しています。"
            rigidMap(id) = RR
            For j = 0 To 1: RRigids(RR).center(j) = number(RPX_InputCell(ws, r, 7 + j), "剛体中心"): Next j
            For j = 0 To 2
                value = IDNumber(RPX_InputCell(ws, r, 9 + j), "剛体固定")
                If value <> 0# And value <> 1# Then err.Raise 5, , "固定は0または1です。"
                RRigids(RR).fixed(j) = (value = 1#)
                RRigids(RR).load0(j) = number(RPX_InputCell(ws, r, 12 + j), "剛体固定荷重")
                RRigids(RR).load1(j) = number(RPX_InputCell(ws, r, 15 + j), "剛体比例荷重")
            Next j
            RR = RR + 1
        End If
    Next r
    If RM = 0 Then err.Raise 5, , "材料を入力してください。"
    If RStressScale < 1# Then RStressScale = 1#
    Set ws = RPX_Sheet("要素定義"): DPCount = 0: DRCount = 0: DFCount = 0
    ReDim DPXY(0 To 999, 0 To 1): ReDim DRegions(0 To 999): ReDim DFaces(0 To 999)
    For r = 7 To 1006
        If Len(CStr(RPX_InputCell(ws, r, 1))) > 0 Then
            id = IDNumber(RPX_InputCell(ws, r, 1), "点ID")
            If id = 0 Or pointMap.Exists(id) Then err.Raise 5, , "点IDが重複しています。"
            pointMap(id) = DPCount: DPXY(DPCount, 0) = number(RPX_InputCell(ws, r, 2), "X座標"): DPXY(DPCount, 1) = number(RPX_InputCell(ws, r, 3), "Y座標"): DPCount = DPCount + 1
        End If
    Next r
    For r = 7 To 1006
        If Len(CStr(RPX_InputCell(ws, r, 5))) > 0 Then
            id = IDNumber(RPX_InputCell(ws, r, 5), "領域ID")
            If id = 0 Or regionMap.Exists(id) Then err.Raise 5, , "領域IDが重複しています。"
            regionMap(id) = DRCount
            With DRegions(DRCount)
                id = IDNumber(RPX_InputCell(ws, r, 6), "領域材料ID")
                If Not materialMap.Exists(id) Then err.Raise 5, , "領域の材料IDが未定義です。"
                .materialId = materialMap(id): id = IDNumber(RPX_InputCell(ws, r, 7), "領域剛体ID"): .rigidId = -1
                If id > 0 Then
                    If Not rigidMap.Exists(id) Then err.Raise 5, , "領域の剛体IDが未定義です。"
                    .rigidId = rigidMap(id)
                End If
                text = Replace(Replace(CStr(RPX_InputCell(ws, r, 8)), "、", ","), " ", ""): parts = Split(text, ",")
                .pointCount = UBound(parts) + 1: ReDim .points(0 To .pointCount - 1)
                For j = 0 To .pointCount - 1
                    id = IDNumber(parts(j), "境界点ID"): If Not pointMap.Exists(id) Then err.Raise 5, , "境界点IDが未定義です。"
                    .points(j) = pointMap(id)
                Next j
            End With
            DRCount = DRCount + 1
        End If
    Next r
    For r = 7 To 1006
        If Len(CStr(RPX_InputCell(ws, r, 10))) > 0 Then
            id = IDNumber(RPX_InputCell(ws, r, 10), "面ID")
            If id = 0 Or faceMap.Exists(id) Then err.Raise 5, , "面IDが重複しています。"
            faceMap(id) = DFCount
            With DFaces(DFCount)
                id = IDNumber(RPX_InputCell(ws, r, 11), "面領域ID"): If Not regionMap.Exists(id) Then err.Raise 5, , "面の領域IDが未定義です。"
                .regionId = regionMap(id)
                id = IDNumber(RPX_InputCell(ws, r, 12), "開始点"): If Not pointMap.Exists(id) Then err.Raise 5, , "開始点が未定義です。"
                .startPoint = pointMap(id)
                id = IDNumber(RPX_InputCell(ws, r, 13), "終了点"): If Not pointMap.Exists(id) Then err.Raise 5, , "終了点が未定義です。"
                .endPoint = pointMap(id): id = IDNumber(RPX_InputCell(ws, r, 14), "循環")
                If id > 1 Then err.Raise 5, , "循環は0または1です。"
                .wrap = (id = 1)
                .kind = Restraint(CStr(RPX_InputCell(ws, r, 15))): .rigidId = -1: id = IDNumber(RPX_InputCell(ws, r, 16), "面剛体ID")
                If .kind = "rigid" Then
                    If Not rigidMap.Exists(id) Then err.Raise 5, , "面の剛体IDが未定義です。"
                    .rigidId = rigidMap(id)
                ElseIf id <> 0 Then
                    err.Raise 5, , "剛体IDを指定する面は拘束を剛体にしてください。"
                End If
                For j = 0 To 1: .load0(j) = number(RPX_InputCell(ws, r, 17 + j), "面固定荷重"): .load1(j) = number(RPX_InputCell(ws, r, 19 + j), "面比例荷重"): Next j
            End With
            DFCount = DFCount + 1
        End If
    Next r
    If DPCount < 3 Or DRCount = 0 Then err.Raise 5, , "定義点と領域を入力してください。"
    RPX_SetAnalysisPolicy CStr(RPX_Setting("ANALYSIS_POLICY", "LEGACY_DAVIS_CLIP"))
    value = number(RPX_Setting("RAW_LOWER_AUDIT", 1), "RAW_LOWER_AUDIT")
    If value <> 0# And value <> 1# Then err.Raise 5, , "INVALID_SETTING RAW_LOWER_AUDIT"
    RRawAuditEnabled = (value = 1#)
    DTarget = IDNumber(RPX_Setting("TARGET", 1536), "目標要素数"): DSpacing = 0#: RStrengthFactor = 1#
    RSubdivisions = IDNumber(RPX_Setting("QUADRATURE", 8), "散逸積分細分数"): RFsTolerance = number(RPX_Setting("FS_TOL", 0.0001), "Fs許容差")
    If RSubdivisions < 1 Or RSubdivisions > 16 Then err.Raise 5, , "散逸積分の細分数は1から16です。組立上限に合わせています。"
    If RFsTolerance < 0.000001 Or RFsTolerance > 0.01 Then err.Raise 5, , "Fs許容差は0.000001から0.01です。"
    value = IDNumber(RPX_Setting("ADAPT", 1), "メッシュ最適化")
    If value > 1 Then err.Raise 5, , "メッシュ最適化は0または1です。"
    value = IDNumber(RPX_Setting("UPPER_PILOT", 0), "上界パイロット")
    If value > 1 Then err.Raise 5, , "上界パイロットは0または1です。"
    value = IDNumber(RPX_Setting("UPPER_PILOT_RIGID", 0), "剛体上界パイロット")
    If value > 1 Then err.Raise 5, , "剛体上界パイロットは0または1です。"
    text = UCase$(Trim$(CStr(RPX_Setting("LOWER_FS_SEED", "LEGACY"))))
    If text <> "LEGACY" And text <> "UPPER_TRIAL" Then err.Raise 5, , "下界初期試行はLEGACYまたはUPPER_TRIALです。"
    value = number(RPX_Setting("LOWER_SEED_WIDTH", 0.02), "下界試行幅")
    If value <= 0# Or value >= 1# Then err.Raise 5, , "下界試行幅は0より大きく1未満です。"
    value = IDNumber(RPX_Setting("LOWER_MODEL_CACHE", 0), "下界構造再利用")
    If value > 1 Then err.Raise 5, , "下界構造再利用は0または1です。"
    text = UCase$(Trim$(CStr(RPX_Setting("LOWER_KKT_BACKEND", "GENERIC"))))
    If text <> "GENERIC" And text <> "COMPARE" And text <> "AUTO" Then err.Raise 5, , "下界KKT解法はGENERIC、COMPARE、AUTOです。"
    value = IDNumber(RPX_Setting("SCHUR_ALLOW_RIGID", 0), "Schur剛体許可")
    If value > 1 Then err.Raise 5, , "Schur剛体許可は0または1です。"
    value = number(RPX_Setting("SCHUR_MAX_BYTES", 0), "Schur作業上限")
    If value < 0# Or value > 2147483647# Or value <> Fix(value) Then err.Raise 5, , "Schur作業上限は0以上の整数です。"
    RPX_P4Configure
    value = IDNumber(RPX_Setting("KKT_CAPTURE", 0), "KKT保存")
    If value <> 0 And value <> 1 Then err.Raise 5, , "KKT_CAPTURE must be 0 or 1"
    value = IDNumber(RPX_Setting("TEST_INJECTION", 0), "試験注入")
    If value > 1 Then err.Raise 5, , "試験注入は0または1です。"
    value = IDNumber(RPX_Setting("CYCLES", 3), "最大最適化回数")
    If value > 20 Then err.Raise 5, , "最大最適化回数は0から20です。"
    value = number(RPX_Setting("GAP", 0.01), "目標区間幅")
    If value <= 0# Or value >= 2# Then err.Raise 5, , "目標区間幅は0より大きく2未満です。"
    value = number(RPX_Setting("REFINE", 0.2), "局所細分化率")
    If value < 0# Or value > 1# Then err.Raise 5, , "局所細分化率は0から1です。"
    Select Case CStr(RPX_Setting("GRAVITY", "固定"))
        Case "比例": RGravityMode = "reference"
        Case "固定": RGravityMode = "fixed"
        Case Else: err.Raise 5, , "自重の扱いは固定または比例です。"
    End Select
    Set RPX_FaceIDs = faceMap
    Set RPX_RigidIDs = rigidMap
    RPX_AssertInputs "read_inputs_end"
    Call RPX_PrepareDefinition
End Sub
Public Sub RPX_Stop()
    XCancel = True
End Sub
Public Sub RPX_Invalidate()
    Dim ws As Worksheet, i As Long
    If RPX_InternalWrite Then Exit Sub
    If RPX_Busy Then XCancel = True
    RPX_BearingInvalidate
    Call RPX_CertBeginRun
    RPX_MeshCurrent = False: RPX_ResultCurrent = False
    RPX_FsBracketReset
    RProgramKind = "": RPX_Sheet("解析結果").Range("A1").Value2 = "未解析（入力変更）"
    RPX_Sheet("解析結果").Range("A3:F100").ClearContents
    Set ws = RPX_Sheet("図"): ws.Range("A1").Value2 = "入力変更：表示中のメッシュは再生成前のものです。": ws.Range("A3").ClearContents
    For i = ws.Shapes.count To 1 Step -1
        If ws.Shapes(i).AlternativeText = "RPX_VELOCITY" Then ws.Shapes(i).Delete
    Next i
End Sub