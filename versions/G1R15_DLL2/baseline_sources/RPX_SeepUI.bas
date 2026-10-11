Attribute VB_Name = "RPX_SeepUI"
Option Explicit
' G1R15：自動試験用の静音スイッチ。RPX_SeepSilent=True の間はメッセージを表示せず記録する。
Public RPX_SeepSilent As Boolean, RPX_SeepLastMsg As String

' RPX_SeepUI : 浸透入力シート、既存の形状定義からの浸透解析、図シートへの重ね描き。
' 依存：RPX_Seep（ソルバー）、既存の RPX_ReadInputs / RPX_GenerateMesh / RPX_Sheet / DRegions / REdges / RUpperField。
' 既存コードへのフック：RPX_DrawCore の末尾で Call RPX_SeepOverlay（1行のみ）。

Private Const SEEP_SHEET As String = "浸透入力"
Private Const FIRST_ROW As Long = 7
Private Const LAST_ROW As Long = 206
Private Const ADV_HEAD As Long = 208
Private Const ADV_FIRST As Long = 209
Private Const LAYOUT_MARK As String = "RPX_SEEP_LAYOUT_V2"
Private Const PREVIEW_TAG As String = "RPX_SEEP_PREVIEW"
' 設定の論理番号（旧配置ではN列の行番号。新配置では SettingCell で位置を変換）。
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
Private Const ROW_MODE As Long = 24
Private Const ROW_INVMAX As Long = 25
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

Private Sub SeepMsg(ByVal text As String, Optional ByVal style As Long = 0, Optional ByVal title As String = "")
    RPX_SeepLastMsg = RPX_SeepLastMsg & text & vbLf
    If RPX_SeepSilent Then
        On Error Resume Next
        RPX_SeepTrace "MSG: " & text
    Else
        If Len(title) > 0 Then MsgBox text, style, title Else MsgBox text, style
    End If
End Sub
Public Sub RPX_SeepSetSilent(ByVal flag As Boolean)
    RPX_SeepSilent = flag: If flag Then RPX_SeepLastMsg = ""
End Sub
Public Function RPX_SeepGetLastMsg() As String
    RPX_SeepGetLastMsg = RPX_SeepLastMsg
End Function
' 浸透解析をメッセージなしで実行し、表示されるはずだったメッセージを返す（空なら正常）。
Public Function RPX_SeepRunSilent() As String
    RPX_SeepSilent = True: RPX_SeepLastMsg = ""
    On Error Resume Next
    RPX_SeepRun
    If err.number <> 0 Then RPX_SeepLastMsg = RPX_SeepLastMsg & "ERR " & err.description
    On Error GoTo 0
    RPX_SeepSilent = False
    RPX_SeepRunSilent = RPX_SeepLastMsg
End Function
Private Sub Diag(ByVal text As String)
    overlayDiag = overlayDiag & text & vbLf
End Sub

' 追跡ログ（%TEMP%\rpx_seep_trace.log）。失敗しても無視。
Public Sub RPX_SeepTrace(ByVal text As String)
    Dim f As Integer
    On Error Resume Next
    f = FreeFile
    Open Environ$("TEMP") & "\rpx_seep_trace.log" For Append As #f
    Print #f, Format$(now, "yyyy-mm-dd hh:nn:ss") & " " & text
    Close #f
    err.Clear
End Sub

Public Function RPX_SeepLastDiagnostics() As String
    RPX_SeepLastDiagnostics = overlayDiag
End Function

' ---------------------------------------------------------------- sheet setup
Private Sub SeepHeader(ByVal ws As Worksheet, ByVal address As String, ByVal labels As Variant)
    Dim i As Long
    For i = LBound(labels) To UBound(labels): ws.Range(address).offset(0, i).Value2 = labels(i): Next i
    With ws.Range(address).Resize(1, UBound(labels) - LBound(labels) + 1)
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255): .interior.Color = RGB(38, 67, 91): .WrapText = True
    End With
    ws.rows(ws.Range(address).row).RowHeight = 38
End Sub

Private Sub SeepButton(ByVal ws As Worksheet, ByVal label As String, ByVal macro As String, ByVal x As Double, ByVal y As Double, ByVal width As Double)
    Dim s As shape
    Set s = ws.Shapes.AddShape(5, x, y, width, 28)
    s.TextFrame2.TextRange.text = label: s.TextFrame2.TextRange.Font.size = 11: s.TextFrame2.TextRange.Font.fill.ForeColor.RGB = RGB(255, 255, 255)
    s.fill.ForeColor.RGB = RGB(36, 104, 130): s.line.Visible = False: s.OnAction = macro
End Sub

Private Function SettingValue(ByVal ws As Worksheet, ByVal r As Long, ByVal fallback As Double) As Double
    Dim v As Variant
    Dim c As Range
    SettingValue = fallback
    Set c = SettingCell(ws, r)
    If c Is Nothing Then Exit Function
    v = c.Value2
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

' ---------------------------------------------------------------- layout v2
' 上部：水の与え方（B3）・必要な入力（H3）、共通（γw B4／連成 D4／サクション F4）。
' 表（7～206行）：A:C 透水係数、E:K 水理境界、M:N 地下水面、P:Q 観測点。D・L・O・R は入力チェックの結果欄。
' 詳細設定：208行の見出しの下（A 項目、C 値、E 説明）。Z列（非表示）に版マーカーと選択肢。
Private Function IsV2(ByVal ws As Worksheet) As Boolean
    On Error Resume Next
    IsV2 = (CStr(ws.Range("Z1").Value2) = LAYOUT_MARK)
    err.Clear
End Function

Private Function AdvIndex(ByVal r As Long) As Long
    Select Case r
        Case ROW_TARGET: AdvIndex = 0
        Case ROW_KRMIN: AdvIndex = 1
        Case ROW_BETA: AdvIndex = 2
        Case ROW_RELAX: AdvIndex = 3
        Case ROW_TOL: AdvIndex = 4
        Case ROW_MAXIT: AdvIndex = 5
        Case ROW_RIGIDK: AdvIndex = 6
        Case ROW_INVMAX: AdvIndex = 7
        Case ROW_DRAWFS: AdvIndex = 8
        Case ROW_SLIPTHR: AdvIndex = 9
        Case ROW_SLIPPCT: AdvIndex = 10
        Case ROW_SLIPMAXFRAC: AdvIndex = 11
        Case ROW_ARCFIT: AdvIndex = 12
        Case ROW_DRAWSTRAIN: AdvIndex = 13
        Case ROW_STRAINTHR: AdvIndex = 14
        Case Else: AdvIndex = -1
    End Select
End Function

Private Function SettingCell(ByVal ws As Worksheet, ByVal r As Long) As Range
    Dim k As Long
    If IsV2(ws) Then
        Select Case r
            Case ROW_GAMMAW: Set SettingCell = ws.Range("B4")
            Case ROW_COUPLE: Set SettingCell = ws.Range("D4")
            Case ROW_SUCTION: Set SettingCell = ws.Range("F4")
            Case Else
                k = AdvIndex(r)
                If k >= 0 Then Set SettingCell = ws.cells(ADV_FIRST + k, 3)
        End Select
    Else
        Set SettingCell = ws.cells(r, 14)
    End If
End Function

Private Function ModeLabel(ByVal i As Long) As String
    ModeLabel = Choose(i, "水位・水頭の条件から計算する", "地下水面の位置を入力する（斜面に沿った流れを考慮）", _
                          "地下水面の位置を入力する（水は静止とみなす）", "地下水面に合うよう境界の水位を推定する")
End Function

' 内部モード：0 浸透FEM、1 地下水面＋cos2θ、2 静水圧、3 逆推定。未選択は -1。
Public Function RPX_SeepMode(ByVal ws As Worksheet) As Long
    Dim i As Long, s As String
    RPX_SeepMode = -1
    If IsV2(ws) Then
        s = CStr(ws.Range("B3").Value2)
        For i = 1 To 4
            If StrComp(s, CStr(ws.cells(3 + i, 26).Value2), vbBinaryCompare) = 0 Then RPX_SeepMode = i - 1: Exit Function
        Next i
    Else
        RPX_SeepMode = CLng(SettingValue(ws, ROW_MODE, 0#))
    End If
End Function

Private Sub AdvSetting(ByVal ws As Worksheet, ByVal r As Long, ByVal label As String, ByVal value As Variant, ByVal note As String, ByRef keep As Object)
    Dim row As Long
    row = ADV_FIRST + AdvIndex(r)
    ws.cells(row, 1).Value2 = label: ws.cells(row, 5).Value2 = note
    If keep.Exists(r) Then ws.cells(row, 3).Value2 = keep(r) Else ws.cells(row, 3).Value2 = value
    ws.cells(row, 3).interior.Color = RGB(243, 248, 251)
End Sub

Private Function CountRows(ByRef v As Variant) As Long
    Dim i As Long, j As Long
    For i = UBound(v, 1) To 1 Step -1
        For j = 1 To UBound(v, 2)
            If Len(CStr(v(i, j))) > 0 Then CountRows = i: Exit Function
        Next j
    Next i
End Function

Private Sub GreyWhen(ByVal rng As Range, ByVal formula As String)
    Dim fc As FormatCondition
    Set fc = rng.FormatConditions.Add(Type:=xlExpression, Formula1:=formula)
    fc.interior.Color = RGB(217, 217, 217): fc.Font.Color = RGB(150, 150, 150)
End Sub

' 浸透入力シートを作成・書式設定する。旧配置（N列の設定・W:Xの表）は自動で新配置へ移す。入力値は消さない。
Public Sub RPX_SeepSetupSheet()
    Dim ws As Worksheet, i As Long, previous As Boolean, keep As Object, perm As Variant, bnd As Variant, wt As Variant, obs As Variant
    Dim modeIdx As Long, oldMode As Long, wasV2 As Boolean, migrated As Boolean, r As Variant, v As Variant, n As Long
    previous = RPX_InternalWrite: RPX_InternalWrite = True
    On Error GoTo Done
    Set ws = RPX_Sheet(SEEP_SHEET)
    Set keep = CreateObject("Scripting.Dictionary")
    wasV2 = IsV2(ws)
    ' ---- 既存値の退避
    If wasV2 Then
        perm = ws.Range("A7:C206").Value2: bnd = ws.Range("E7:K206").Value2
        wt = ws.Range("M7:N206").Value2: obs = ws.Range("P7:Q206").Value2
        modeIdx = RPX_SeepMode(ws) + 1
        For Each r In Array(ROW_GAMMAW, ROW_COUPLE, ROW_SUCTION, ROW_TARGET, ROW_KRMIN, ROW_BETA, ROW_RELAX, ROW_TOL, ROW_MAXIT, ROW_RIGIDK, ROW_INVMAX, ROW_DRAWFS, ROW_SLIPTHR, ROW_SLIPPCT, ROW_SLIPMAXFRAC, ROW_ARCFIT, ROW_DRAWSTRAIN, ROW_STRAINTHR)
            v = SettingCell(ws, CLng(r)).Value2
            If Not IsError(v) Then
                If Len(CStr(v)) > 0 Then keep(CLng(r)) = v
            End If
        Next r
    Else
        perm = ws.Range("A7:C1006").Value2: bnd = ws.Range("E7:K1006").Value2: wt = ws.Range("W7:X1006").Value2
        If CountRows(perm) > 200 Or CountRows(bnd) > 200 Or CountRows(wt) > 200 Then err.Raise 5, , "旧配置の表が200行を超えるため自動移行できません（207行目以降を整理してから実行してください）。"
        If CStr(ws.Range("M6").Value2) = "項目" Then
            migrated = True
            For Each r In Array(ROW_GAMMAW, ROW_COUPLE, ROW_SUCTION, ROW_TARGET, ROW_KRMIN, ROW_BETA, ROW_RELAX, ROW_TOL, ROW_MAXIT, ROW_RIGIDK, ROW_INVMAX, ROW_DRAWFS, ROW_SLIPTHR, ROW_SLIPPCT, ROW_SLIPMAXFRAC, ROW_ARCFIT, ROW_DRAWSTRAIN, ROW_STRAINTHR)
                v = ws.cells(CLng(r), 14).Value2
                If Not IsError(v) Then
                    If Len(CStr(v)) > 0 Then keep(CLng(r)) = v
                End If
            Next r
            v = ws.cells(ROW_MODE, 14).Value2: oldMode = 0
            If Not IsError(v) Then
                If IsNumeric(v) And Len(CStr(v)) > 0 Then oldMode = CLng(v)
            End If
            If oldMode < 0 Or oldMode > 3 Then oldMode = 0
            modeIdx = oldMode + 1
        End If
        perm = ws.Range("A7:C206").Value2: bnd = ws.Range("E7:K206").Value2
        If oldMode = 3 Then
            obs = ws.Range("W7:X206").Value2: wt = Empty
        Else
            wt = ws.Range("W7:X206").Value2: obs = Empty
        End If
    End If
    If modeIdx < 1 Or modeIdx > 4 Then modeIdx = 1
    ' ---- 書式の作り直し
    For i = ws.Shapes.count To 1 Step -1: ws.Shapes(i).Delete: Next i
    On Error Resume Next
    ws.cells.FormatConditions.Delete
    ws.cells.Validation.Delete
    err.Clear
    On Error GoTo Done
    ws.cells.Clear
    ws.columns.Hidden = False
    ws.cells.Font.name = "Yu Gothic": ws.cells.Font.size = 11
    ws.Range("Z1").Value2 = LAYOUT_MARK
    For i = 1 To 4: ws.cells(3 + i, 26).Value2 = ModeLabel(i): Next i
    ws.Range("Z3").formula = "=IFERROR(MATCH($B$3,$Z$4:$Z$7,0),0)"
    ws.Range("A1").Value2 = "浸透流解析・水圧の入力": ws.Range("A1").Font.size = 20: ws.Range("A1").Font.Bold = True
    ws.rows(1).RowHeight = 34
    ws.Range("A3").Value2 = "水の与え方": ws.Range("A3").Font.Bold = True
    ws.Range("B3:F3").Merge
    ws.Range("B3").Value2 = ModeLabel(modeIdx)
    With ws.Range("B3").Validation
        .Delete: .Add xlValidateList, xlValidAlertStop, xlBetween, "=$Z$4:$Z$7"
        .IgnoreBlank = False: .InCellDropdown = True
    End With
    ws.Range("B3:F3").interior.Color = RGB(255, 242, 204): ws.Range("B3:F3").Font.Bold = True
    ws.Range("G3").Value2 = "必要な入力"
    ws.Range("AA3").Value2 = "一覧から水の与え方を選んでください"
    ws.Range("AA4").Value2 = "透水係数（A:C）と水理境界（E:K）。境界の水頭・水位を数値で入力"
    ws.Range("AA5").Value2 = "地下水面（M:N）の折れ線をxの増加順に2点以上。透水係数・境界は不要"
    ws.Range("AA6").Value2 = "地下水面（M:N）の折れ線をxの増加順に2点以上。水圧は水面からの深さ×γw"
    ws.Range("AA7").Value2 = "透水係数・水理境界（推定する値は ? または ?初期値）と観測点（P:Q）"
    ws.Range("H3").formula = "=IFERROR(INDEX($AA$4:$AA$7,MATCH($B$3,$Z$4:$Z$7,0)),$AA$3)"
    ws.Range("H3").Font.Color = RGB(192, 0, 0)
    ws.Range("A4").Value2 = "γw [kN/m3]": ws.Range("C4").Value2 = "連成する": ws.Range("E4").Value2 = "サクション"
    ws.Range("B4").Value2 = 9.81: If keep.Exists(ROW_GAMMAW) Then ws.Range("B4").Value2 = keep(ROW_GAMMAW)
    ws.Range("D4").Value2 = 0: If keep.Exists(ROW_COUPLE) Then ws.Range("D4").Value2 = keep(ROW_COUPLE)
    ws.Range("F4").Value2 = 0: If keep.Exists(ROW_SUCTION) Then ws.Range("F4").Value2 = keep(ROW_SUCTION)
    ws.Range("B4,D4,F4").interior.Color = RGB(243, 248, 251)
    With ws.Range("D4,F4").Validation
        .Delete: .Add xlValidateList, xlValidAlertStop, xlBetween, "0,1"
        .IgnoreBlank = True: .InCellDropdown = True
    End With
    ws.Range("G4").Value2 = "連成 1：間隙水圧を極限解析へ（有効応力）／サクション 0：負圧は0（推奨）"
    ws.Range("A5").Value2 = "灰色の表は選んだ水の与え方では使いません。「入力チェック」で問題のある行を赤く表示します。"
    SeepHeader ws, "A6", Array("材料ID", "kx [m/s]", "ky [m/s]", "確認")
    SeepHeader ws, "E6", Array("境界ID", "領域ID", "開始点", "終了点", "循環", "種別", "値 [m]または[m3/s/m]", "確認")
    SeepHeader ws, "M6", Array("地下水面 x [m]", "地下水面 y [m]", "確認")
    SeepHeader ws, "P6", Array("観測点 x [m]", "観測水位 y [m]", "確認")
    ws.Range("A7:C206,E7:K206,M7:N206,P7:Q206").interior.Color = RGB(243, 248, 251)
    If Not IsEmpty(perm) Then ws.Range("A7").Resize(UBound(perm, 1), UBound(perm, 2)).Value2 = perm
    If Not IsEmpty(bnd) Then ws.Range("E7").Resize(UBound(bnd, 1), UBound(bnd, 2)).Value2 = bnd
    If Not IsEmpty(wt) Then ws.Range("M7").Resize(UBound(wt, 1), UBound(wt, 2)).Value2 = wt
    If Not IsEmpty(obs) Then ws.Range("P7").Resize(UBound(obs, 1), UBound(obs, 2)).Value2 = obs
    With ws.Range("I7:I206").Validation
        .Delete: .Add xlValidateList, xlValidAlertStop, xlBetween, "0,1"
        .IgnoreBlank = True: .InCellDropdown = True
    End With
    With ws.Range("J7:J206").Validation
        .Delete: .Add xlValidateList, xlValidAlertStop, xlBetween, "水頭,水位,水位+浸出面,浸出面,流量,不透水"
        .IgnoreBlank = True: .InCellDropdown = True: .ShowError = True
        .ErrorTitle = "入力を確認してください": .ErrorMessage = "一覧から選択してください。"
    End With
    GreyWhen ws.Range("A6:D206,E6:L206"), "=AND($B$3<>$Z$4,$B$3<>$Z$7)"
    GreyWhen ws.Range("M6:O206"), "=AND($B$3<>$Z$5,$B$3<>$Z$6)"
    GreyWhen ws.Range("P6:R206"), "=$B$3<>$Z$7"
    ' ---- 詳細設定
    ws.Range("A" & ADV_HEAD).Value2 = "詳細設定（通常は変更不要）"
    With ws.Range("A" & ADV_HEAD & ":K" & ADV_HEAD)
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255): .interior.Color = RGB(38, 67, 91)
    End With
    AdvSetting ws, ROW_TARGET, "目標要素数", 0, "0：設定シートのTARGETを使用", keep
    AdvSetting ws, ROW_KRMIN, "krmin", 0.0001, "不飽和域の透水係数比の下限", keep
    AdvSetting ws, ROW_BETA, "β [1/m]", 20, "kr=exp(β・圧力水頭)。大きいほど鋭いが収束しにくい", keep
    AdvSetting ws, ROW_RELAX, "緩和係数", 0.5, "0～1", keep
    AdvSetting ws, ROW_TOL, "収束判定 [m]", 0.000001, "水頭変化の最大値", keep
    AdvSetting ws, ROW_MAXIT, "最大反復回数", 300, "", keep
    AdvSetting ws, ROW_RIGIDK, "剛体の透水係数比", 0.000001, "剛体領域 k = 比率×最大k", keep
    AdvSetting ws, ROW_INVMAX, "推定の計算回数上限", 40, "「境界の水位を推定」での浸透解析の最大回数（1回数秒）", keep
    AdvSetting ws, ROW_DRAWFS, "自由水面を描画", 1, "1：図シートに描画", keep
    AdvSetting ws, ROW_SLIPTHR, "すべり線しきい値", 0.5, "速度ジャンプ／参照値以上の辺を表示。0で非表示", keep
    AdvSetting ws, ROW_SLIPPCT, "すべり線 参照分位点", 0.95, "非ゼロ速度ジャンプの分位点（0.9～0.95推奨）", keep
    AdvSetting ws, ROW_SLIPMAXFRAC, "すべり線 最大表示割合", 0.1, "表示する辺の上限（土の内部辺に対する割合）", keep
    AdvSetting ws, ROW_ARCFIT, "円弧当てはめ", 1, "1：推定すべり面に円弧を当てはめて破線表示", keep
    AdvSetting ws, ROW_DRAWSTRAIN, "ひずみ速度を描画", 0, "1：塑性せん断ひずみ速度の大きい要素を着色", keep
    AdvSetting ws, ROW_STRAINTHR, "ひずみ速度しきい値", 0.2, "要素のひずみ速度×面積／最大値", keep
    ws.columns("A:C").ColumnWidth = 11: ws.columns("D").ColumnWidth = 22
    ws.columns("E:I").ColumnWidth = 9: ws.columns("J").ColumnWidth = 14: ws.columns("K").ColumnWidth = 16: ws.columns("L").ColumnWidth = 22
    ws.columns("M:N").ColumnWidth = 12: ws.columns("O").ColumnWidth = 22
    ws.columns("P:Q").ColumnWidth = 12: ws.columns("R").ColumnWidth = 22
    ws.Range("D7:D206,L7:L206,O7:O206,R7:R206").Font.Color = RGB(192, 0, 0)
    ws.columns("Z:AA").Hidden = True
    ws.EnableCalculation = True: ws.Calculate
    SeepButton ws, "浸透解析", "RPX_SeepRun", 560, 3, 90
    SeepButton ws, "入力チェック", "RPX_SeepCheckInputs", 655, 3, 95
    SeepButton ws, "プレビュー", "RPX_SeepPreview", 755, 3, 90
    SeepButton ws, "図を表示", "RPX_SeepShowDiagram", 850, 3, 80
    SeepButton ws, "書式更新", "RPX_SeepSetupSheet", 935, 3, 80
Done:
    RPX_InternalWrite = previous
    If err.number <> 0 Then
        SeepMsg "浸透入力シートの作成に失敗しました：" & err.description, vbExclamation
    ElseIf migrated Then
        SeepMsg "浸透入力シートを新しい配置へ移行しました（透水係数・境界・地下水面/観測点・設定を引き継ぎ）。", vbInformation
    End If
End Sub

' ---------------------------------------------------------------- input check
Private Sub MarkRow(ByVal ws As Worksheet, ByVal r As Long, ByVal firstCol As Long, ByVal lastCol As Long, ByVal msgCol As Long, ByVal text As String, ByRef count As Long)
    ws.Range(ws.cells(r, firstCol), ws.cells(r, lastCol)).interior.Color = RGB(255, 199, 206)
    If Len(CStr(ws.cells(r, msgCol).Value2)) > 0 Then
        ws.cells(r, msgCol).Value2 = ws.cells(r, msgCol).Value2 & "／" & text
    Else
        ws.cells(r, msgCol).Value2 = text
    End If
    count = count + 1
End Sub

Private Function CellNum(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long, ByRef ok As Boolean) As Double
    Dim v As Variant
    ok = False: v = ws.cells(r, c).Value2
    If IsError(v) Then Exit Function
    If Len(CStr(v)) = 0 Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    CellNum = CDbl(v): ok = True
End Function

Private Function RowUsed(ByVal ws As Worksheet, ByVal r As Long, ByVal c0 As Long, ByVal c1 As Long) As Boolean
    Dim c As Long
    For c = c0 To c1
        If Len(CStr(ws.cells(r, c).Value2)) > 0 Then RowUsed = True: Exit Function
    Next c
End Function

Private Function RegionHasPoint(ByVal region As Long, ByVal p As Long) As Boolean
    Dim i As Long
    For i = 0 To DRegions(region).pointCount - 1
        If DRegions(region).points(i) = p Then RegionHasPoint = True: Exit Function
    Next i
End Function

' 入力チェック（形状は RPX_ReadInputs 済みであること）。問題のある行を赤くし、件数を返す。
Private Function ValidateInputs(ByVal ws As Worksheet, ByVal mode As Long, ByRef summary As String) As Long
    Dim n As Long, r As Long, ok As Boolean, ok2 As Boolean, x As Double, y As Double, xPrev As Double, cnt As Long, unk As Long, heads As Long
    Dim matMap As Object, pointMap As Object, regionMap As Object, xmin As Double, xmax As Double, i As Long, s As String, k As String
    Dim region As Long, haveRegion As Boolean, xFirst As Double
    ws.Range("A7:C206,E7:K206,M7:N206,P7:Q206").interior.Color = RGB(243, 248, 251)
    ws.Range("D7:D206,L7:L206,O7:O206,R7:R206").ClearContents
    summary = ""
    If mode < 0 Then
        ws.Range("B3").interior.Color = RGB(255, 199, 206)
        summary = "水の与え方（B3）を一覧から選んでください。": ValidateInputs = 1: Exit Function
    End If
    ws.Range("B3:F3").interior.Color = RGB(255, 242, 204)
    xmin = DPXY(0, 0): xmax = xmin
    For i = 1 To DPCount - 1
        If DPXY(i, 0) < xmin Then xmin = DPXY(i, 0)
        If DPXY(i, 0) > xmax Then xmax = DPXY(i, 0)
    Next i
    If mode = 0 Or mode = 3 Then
        Set matMap = ColumnMap(RPX_Sheet("材料データ"), 5, 1004, 1)
        Set pointMap = ColumnMap(RPX_Sheet("要素定義"), 7, 1006, 1)
        Set regionMap = ColumnMap(RPX_Sheet("要素定義"), 7, 1006, 5)
        For r = FIRST_ROW To LAST_ROW
            If RowUsed(ws, r, 1, 3) Then
                x = CellNum(ws, r, 1, ok)
                If Not ok Then
                    MarkRow ws, r, 1, 3, 4, "材料IDが空欄または数値でない", n
                ElseIf Not matMap.Exists(CLng(x)) Then
                    MarkRow ws, r, 1, 3, 4, "材料データにないID", n
                End If
                x = CellNum(ws, r, 2, ok)
                If Not ok Then
                    MarkRow ws, r, 1, 3, 4, "kxが未入力", n
                ElseIf x <= 0# Then
                    MarkRow ws, r, 1, 3, 4, "kxは正の値", n
                End If
                If Len(CStr(ws.cells(r, 3).Value2)) > 0 Then
                    y = CellNum(ws, r, 3, ok)
                    If Not ok Or y <= 0# Then MarkRow ws, r, 1, 3, 4, "kyは正の値（空欄ならkx）", n
                End If
                cnt = cnt + 1
            End If
        Next r
        If cnt = 0 Then summary = summary & "透水係数（A:C）が未入力。": n = n + 1
        cnt = 0
        For r = FIRST_ROW To LAST_ROW
            If RowUsed(ws, r, 5, 11) Then
                cnt = cnt + 1: haveRegion = False
                If Len(CStr(ws.cells(r, 5).Value2)) = 0 Then MarkRow ws, r, 5, 11, 12, "境界IDが空欄", n
                x = CellNum(ws, r, 6, ok)
                If Not ok Then
                    MarkRow ws, r, 5, 11, 12, "領域IDが未入力", n
                ElseIf Not regionMap.Exists(CLng(x)) Then
                    MarkRow ws, r, 5, 11, 12, "要素定義にない領域ID", n
                Else
                    region = regionMap(CLng(x)): haveRegion = True
                End If
                For i = 7 To 8
                    x = CellNum(ws, r, i, ok)
                    If Not ok Then
                        MarkRow ws, r, 5, 11, 12, IIf(i = 7, "開始点", "終了点") & "が未入力", n
                    ElseIf Not pointMap.Exists(CLng(x)) Then
                        MarkRow ws, r, 5, 11, 12, IIf(i = 7, "開始点", "終了点") & "が要素定義にない", n
                    ElseIf haveRegion Then
                        If Not RegionHasPoint(region, pointMap(CLng(x))) Then MarkRow ws, r, 5, 11, 12, IIf(i = 7, "開始点", "終了点") & "が領域の外周上にない（形状の外）", n
                    End If
                Next i
                k = Trim$(CStr(ws.cells(r, 10).Value2))
                s = Trim$(CStr(ws.cells(r, 11).Value2))
                Select Case k
                    Case "水頭", "水位", "水位+浸出面", "水位＋浸出面", "流量"
                        If RPX_WTIsUnknown(s) Then
                            If mode = 0 Then
                                MarkRow ws, r, 5, 11, 12, "「?」は「境界の水位を推定」のときだけ", n
                            ElseIf k = "流量" Then
                                MarkRow ws, r, 5, 11, 12, "流量は推定できない", n
                            Else
                                unk = unk + 1: heads = heads + 1
                            End If
                        Else
                            y = CellNum(ws, r, 11, ok)
                            If Not ok Then MarkRow ws, r, 5, 11, 12, "値が未入力または数値でない", n
                            If k <> "流量" Then heads = heads + 1
                        End If
                    Case "浸出面", "不透水"
                    Case ""
                        MarkRow ws, r, 5, 11, 12, "種別が未入力", n
                    Case Else
                        MarkRow ws, r, 5, 11, 12, "種別は一覧から選択", n
                End Select
            End If
        Next r
        If cnt = 0 Then summary = summary & "水理境界（E:K）が未入力。": n = n + 1
        If cnt > 0 And heads = 0 Then summary = summary & "水頭・水位の境界が1つもない。": n = n + 1
        If mode = 3 And unk = 0 Then summary = summary & "推定する境界の値に「?」がない。": n = n + 1
    End If
    If mode = 1 Or mode = 2 Then
        cnt = 0
        For r = FIRST_ROW To LAST_ROW
            If RowUsed(ws, r, 13, 14) Then
                x = CellNum(ws, r, 13, ok): y = CellNum(ws, r, 14, ok2)
                If Not ok Or Not ok2 Then
                    MarkRow ws, r, 13, 14, 15, "x・yの両方を数値で", n
                Else
                    If cnt > 0 And x <= xPrev Then MarkRow ws, r, 13, 14, 15, "xが増加していない", n
                    If cnt = 0 Then xFirst = x
                    xPrev = x: cnt = cnt + 1
                End If
            End If
        Next r
        If cnt < 2 Then
            summary = summary & "地下水面（M:N）は2点以上必要。": n = n + 1
        ElseIf xFirst > xmin + 0.000001 Or xPrev < xmax - 0.000001 Then
            summary = summary & "（注意）地下水面が形状の左右端まで届いていない：端は水平に延長します。"
        End If
    End If
    If mode = 3 Then
        cnt = 0
        For r = FIRST_ROW To LAST_ROW
            If RowUsed(ws, r, 16, 17) Then
                x = CellNum(ws, r, 16, ok): y = CellNum(ws, r, 17, ok2)
                If Not ok Or Not ok2 Then
                    MarkRow ws, r, 16, 17, 18, "x・yの両方を数値で", n
                ElseIf x < xmin - 0.000001 Or x > xmax + 0.000001 Then
                    MarkRow ws, r, 16, 17, 18, "形状の範囲外（x=" & xmin & "～" & xmax & "）", n
                Else
                    cnt = cnt + 1
                End If
            End If
        Next r
        If cnt = 0 Then
            summary = summary & "観測点（P:Q）がない。": n = n + 1
        ElseIf cnt < 2 * unk + 1 Then
            summary = summary & "（注意）観測点" & cnt & "点は推定値" & unk & "個に対して少ない（" & 2 * unk + 1 & "点以上推奨）。"
        End If
    End If
    ValidateInputs = n
End Function

Public Sub RPX_SeepCheckInputs()
    Dim ws As Worksheet, previous As Boolean, n As Long, summary As String, failText As String
    If RPX_Busy Then SeepMsg "解析またはメッシュ生成を実行中です。", vbExclamation: Exit Sub
    Set ws = OptionalSheet(SEEP_SHEET)
    If ws Is Nothing Then RPX_SeepSetupSheet: Exit Sub
    If Not IsV2(ws) Then RPX_SeepSetupSheet
    previous = RPX_InternalWrite: RPX_InternalWrite = True
    On Error GoTo Failed
    Call RPX_ReadInputs
    n = ValidateInputs(ws, RPX_SeepMode(ws), summary)
    GoTo Done
Failed:
    failText = err.description
    Resume Done
Done:
    On Error Resume Next
    RPX_InternalWrite = previous
    If Len(failText) > 0 Then
        SeepMsg "形状の読み込みに失敗しました：" & failText, vbExclamation
    ElseIf n = 0 Then
        SeepMsg "入力に問題は見つかりませんでした。" & IIf(Len(summary) > 0, vbLf & summary, ""), vbInformation
    Else
        SeepMsg "問題が " & n & " 件あります。赤い行と確認欄を見てください。" & IIf(Len(summary) > 0, vbLf & summary, ""), vbExclamation
    End If
End Sub

' ---------------------------------------------------------------- preview
Private Sub PreviewLine(ByVal ws As Worksheet, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal col As Long, ByVal w As Double, ByVal dash As Long)
    Dim s As shape
    Set s = ws.Shapes.AddLine(CSng(x1), CSng(y1), CSng(x2), CSng(y2))
    s.AlternativeText = PREVIEW_TAG: s.line.ForeColor.RGB = col: s.line.weight = w
    If dash > 0 Then s.line.DashStyle = dash
End Sub

Private Sub PreviewText(ByVal ws As Worksheet, ByVal x As Double, ByVal y As Double, ByVal text As String, ByVal col As Long)
    Dim s As shape
    Set s = ws.Shapes.AddTextbox(1, CSng(x), CSng(y), 160, 16)
    s.AlternativeText = PREVIEW_TAG: s.fill.Visible = False: s.line.Visible = False
    s.TextFrame2.TextRange.text = text: s.TextFrame2.TextRange.Font.size = 9: s.TextFrame2.TextRange.Font.fill.ForeColor.RGB = col
    s.TextFrame2.MarginLeft = 0: s.TextFrame2.MarginTop = 0
End Sub

Private Function KindColor(ByVal k As String, ByVal unknown As Boolean) As Long
    If unknown Then KindColor = RGB(204, 0, 153): Exit Function
    Select Case k
        Case "水頭": KindColor = RGB(0, 70, 200)
        Case "水位", "水位+浸出面", "水位＋浸出面": KindColor = RGB(0, 160, 220)
        Case "浸出面": KindColor = RGB(0, 150, 60)
        Case "流量": KindColor = RGB(230, 120, 0)
        Case Else: KindColor = RGB(90, 90, 90)
    End Select
End Function

' 解析せずに、形状・水理境界・地下水面・観測点を図シートに描く。
Public Sub RPX_SeepPreview()
    Dim ws As Worksheet, fig As Worksheet, previous As Boolean, mode As Long, i As Long, j As Long, r As Long, n As Long, failText As String
    Dim xmin As Double, xmax As Double, ymin As Double, ymax As Double, f As Double, pointMap As Object, regionMap As Object
    Dim region As Long, sp As Long, ep As Long, a As Long, b As Long, stepCount As Long, found As Boolean, col As Long, k As String, s As String
    Dim ok As Boolean, ok2 As Boolean, x As Double, y As Double, px As Double, py As Double, havePrev As Boolean, sh As shape, summary As String, bad As Long
    If RPX_Busy Then SeepMsg "解析またはメッシュ生成を実行中です。", vbExclamation: Exit Sub
    Set ws = OptionalSheet(SEEP_SHEET)
    If ws Is Nothing Then RPX_SeepSetupSheet: Exit Sub
    If Not IsV2(ws) Then RPX_SeepSetupSheet
    previous = RPX_InternalWrite: RPX_InternalWrite = True
    On Error GoTo Failed
    Call RPX_ReadInputs
    mode = RPX_SeepMode(ws)
    bad = ValidateInputs(ws, mode, summary)
    Set fig = RPX_Sheet("図")
    For i = fig.Shapes.count To 1 Step -1
        If fig.Shapes(i).AlternativeText = PREVIEW_TAG Then fig.Shapes(i).Delete
    Next i
    xmin = DPXY(0, 0): xmax = xmin: ymin = DPXY(0, 1): ymax = ymin
    For i = 1 To DPCount - 1
        If DPXY(i, 0) < xmin Then xmin = DPXY(i, 0)
        If DPXY(i, 0) > xmax Then xmax = DPXY(i, 0)
        If DPXY(i, 1) < ymin Then ymin = DPXY(i, 1)
        If DPXY(i, 1) > ymax Then ymax = DPXY(i, 1)
    Next i
    If xmax <= xmin Or ymax <= ymin Then err.Raise 5, , "形状の範囲が不正です。"
    f = 780# / (xmax - xmin)
    If f * (ymax - ymin) > 460# Then f = 460# / (ymax - ymin)
    ' 形状
    For region = 0 To DRCount - 1
        n = DRegions(region).pointCount
        For i = 0 To n - 1
            a = DRegions(region).points(i): b = DRegions(region).points((i + 1) Mod n)
            PreviewLine fig, px(DPXY(a, 0), xmin, f), py(DPXY(a, 1), ymax, f), px(DPXY(b, 0), xmin, f), py(DPXY(b, 1), ymax, f), RGB(128, 128, 128), 0.75, 0
        Next i
    Next region
    ' 水理境界
    If mode = 0 Or mode = 3 Then
        Set pointMap = ColumnMap(RPX_Sheet("要素定義"), 7, 1006, 1)
        Set regionMap = ColumnMap(RPX_Sheet("要素定義"), 7, 1006, 5)
        For r = FIRST_ROW To LAST_ROW
            If Len(CStr(ws.cells(r, 5).Value2)) > 0 And CStr(ws.cells(r, 12).Value2) = "" Then
                region = regionMap(CLng(ws.cells(r, 6).Value2)): sp = pointMap(CLng(ws.cells(r, 7).Value2)): ep = pointMap(CLng(ws.cells(r, 8).Value2))
                k = Trim$(CStr(ws.cells(r, 10).Value2)): s = Trim$(CStr(ws.cells(r, 11).Value2))
                col = KindColor(k, RPX_WTIsUnknown(s))
                n = DRegions(region).pointCount: found = False
                For i = 0 To n - 1
                    If DRegions(region).points(i) = sp Then found = True: Exit For
                Next i
                If found Then
                    For stepCount = 1 To n
                        a = DRegions(region).points(i): b = DRegions(region).points((i + 1) Mod n)
                        PreviewLine fig, px(DPXY(a, 0), xmin, f), py(DPXY(a, 1), ymax, f), px(DPXY(b, 0), xmin, f), py(DPXY(b, 1), ymax, f), col, 3.5, 0
                        If stepCount = 1 Then PreviewText fig, px(DPXY(a, 0), xmin, f) + 3, py(DPXY(a, 1), ymax, f) - 14, ws.cells(r, 5).Value2 & " " & k & IIf(Len(s) > 0, " " & s, ""), col
                        If b = ep Then Exit For
                        i = (i + 1) Mod n
                    Next stepCount
                    If k = "水位" Or k = "水位+浸出面" Or k = "水位＋浸出面" Then
                        y = CellNum(ws, r, 11, ok)
                        If ok Then PreviewLine fig, px(DPXY(sp, 0), xmin, f), py(y, ymax, f), px(DPXY(ep, 0), xmin, f), py(y, ymax, f), col, 1, 4
                    End If
                End If
            End If
        Next r
    End If
    ' 地下水面
    If mode = 1 Or mode = 2 Then
        havePrev = False
        For r = FIRST_ROW To LAST_ROW
            x = CellNum(ws, r, 13, ok): y = CellNum(ws, r, 14, ok2)
            If ok And ok2 Then
                If havePrev Then PreviewLine fig, px(px, xmin, f), py(py, ymax, f), px(x, xmin, f), py(y, ymax, f), RGB(0, 112, 192), 2.25, 4
                px = x: py = y: havePrev = True
            End If
        Next r
    End If
    ' 観測点
    If mode = 3 Then
        For r = FIRST_ROW To LAST_ROW
            x = CellNum(ws, r, 16, ok): y = CellNum(ws, r, 17, ok2)
            If ok And ok2 Then
                Set sh = fig.Shapes.AddShape(9, CSng(px(x, xmin, f) - 4), CSng(py(y, ymax, f) - 4), 8, 8)
                sh.AlternativeText = PREVIEW_TAG: sh.fill.ForeColor.RGB = RGB(220, 0, 0): sh.line.Visible = False
            End If
        Next r
    End If
    PreviewText fig, 30, 30, "プレビュー（未解析）：" & ModeLabel(mode + 1) & IIf(bad > 0, "　入力の問題 " & bad & " 件", ""), RGB(0, 0, 0)
    fig.Shapes(fig.Shapes.count).width = 600
    GoTo Done
Failed:
    failText = err.description
    Resume Done
Done:
    On Error Resume Next
    RPX_InternalWrite = previous
    If Len(failText) > 0 Then
        SeepMsg "プレビューに失敗しました：" & failText, vbExclamation
    Else
        fig.Activate
        If bad > 0 Then SeepMsg "入力の問題が " & bad & " 件あります（浸透入力の赤い行）。" & IIf(Len(summary) > 0, vbLf & summary, ""), vbExclamation
    End If
End Sub

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
        v = ws.cells(r, col).Value2
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
        Case Else: err.Raise 5, , "浸透境界の種別が不明です：" & text
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
        If Abs(SBcVal(node) - value) > 0.000000001 Then err.Raise 5, , "同じ節点に異なる水頭が指定されています（節点 x=" & sxy(node, 0) & ", y=" & sxy(node, 1) & "）。"
    End If
End Sub

' ---------------------------------------------------------------- run
' RPX_SeepWT（逆解析）から呼ぶ入口。
Public Sub RPX_SeepBuildModel(ByVal ws As Worksheet)
    BuildSeepModel ws
End Sub

' 地下水位モード用：メッシュだけをコピーする。
Private Sub CopyMeshOnly(ByVal ws As Worksheet)
    Dim i As Long, e As Long, j As Long
    SN = rn: SM = re: SNMat = RM + 1
    ReDim sxy(0 To SN - 1, 0 To 1): ReDim STri(0 To SM - 1, 0 To 2): ReDim SMat(0 To SM - 1)
    For i = 0 To SN - 1: sxy(i, 0) = RXY(i, 0): sxy(i, 1) = RXY(i, 1): Next i
    For e = 0 To SM - 1
        For j = 0 To 2: STri(e, j) = RTri(e, j): Next j
        If RRigidId(e) >= 0 Then SMat(e) = RM Else SMat(e) = RMatId(e)
    Next e
    RPX_SeepDefaults
    SGammaW = SettingValue(ws, ROW_GAMMAW, SGammaW)
End Sub

Private Sub BuildSeepModel(ByVal ws As Worksheet)
    Dim matWs As Worksheet, defWs As Worksheet, matMap As Object, pointMap As Object, regionMap As Object
    Dim r As Long, i As Long, e As Long, id As Long, kind As Long, value As Double, kmax As Double, ratio As Double
    Dim haveK() As Boolean, region As Long, sp As Long, ep As Long, px As Double, py As Double, node As Long, j As Long
    Dim pri() As Long, hasHead As Boolean, lengthE As Double, key As Variant, unk As Long
    Set matWs = RPX_Sheet("材料データ"): Set defWs = RPX_Sheet("要素定義")
    Set matMap = ColumnMap(matWs, 5, 1004, 1)
    Set pointMap = ColumnMap(defWs, 7, 1006, 1)
    Set regionMap = ColumnMap(defWs, 7, 1006, 5)
    ' mesh copy
    SN = rn: SM = re: SNMat = RM + 1
    ReDim sxy(0 To SN - 1, 0 To 1): ReDim STri(0 To SM - 1, 0 To 2): ReDim SMat(0 To SM - 1)
    For i = 0 To SN - 1: sxy(i, 0) = RXY(i, 0): sxy(i, 1) = RXY(i, 1): Next i
    For e = 0 To SM - 1
        For j = 0 To 2: STri(e, j) = RTri(e, j): Next j
        If RRigidId(e) >= 0 Then SMat(e) = RM Else SMat(e) = RMatId(e)
    Next e
    ' permeability
    ReDim SKx(0 To SNMat - 1): ReDim SKy(0 To SNMat - 1): ReDim haveK(0 To SNMat - 1)
    For r = FIRST_ROW To LAST_ROW
        If Len(CStr(ws.cells(r, 1).Value2)) > 0 Then
            id = CLng(ws.cells(r, 1).Value2)
            If Not matMap.Exists(id) Then err.Raise 5, , "浸透入力の材料ID " & id & " は材料データにありません。"
            i = matMap(id)
            SKx(i) = CDbl(ws.cells(r, 2).Value2): SKy(i) = CDbl(ws.cells(r, 3).Value2)
            If Len(CStr(ws.cells(r, 3).Value2)) = 0 Then SKy(i) = SKx(i)
            If SKx(i) <= 0# Or SKy(i) <= 0# Then err.Raise 5, , "材料ID " & id & " の透水係数は正の値にしてください。"
            haveK(i) = True
            If SKx(i) > kmax Then kmax = SKx(i)
            If SKy(i) > kmax Then kmax = SKy(i)
        End If
    Next r
    For e = 0 To SM - 1
        If SMat(e) < RM Then
            If Not haveK(SMat(e)) Then
                For Each key In matMap.keys
                    If matMap(key) = SMat(e) Then Exit For
                Next key
                err.Raise 5, , "材料ID " & key & " の透水係数が未入力です（浸透入力 A:C）。"
            End If
        End If
    Next e
    ratio = SettingValue(ws, ROW_RIGIDK, 0.000001)
    SKx(RM) = ratio * kmax: SKy(RM) = ratio * kmax
    ' boundary conditions on exposed mesh edges
    ReDim SBcType(0 To SN - 1): ReDim SBcVal(0 To SN - 1): ReDim pri(0 To SN - 1)
    For i = 0 To SN - 1: pri(i) = -1: Next i
    For r = FIRST_ROW To LAST_ROW
        If Len(CStr(ws.cells(r, 5).Value2)) > 0 Then
            id = CLng(ws.cells(r, 6).Value2)
            If Not regionMap.Exists(id) Then err.Raise 5, , "浸透境界 " & ws.cells(r, 5).Value2 & "：領域ID " & id & " は未定義です。"
            region = regionMap(id)
            id = CLng(ws.cells(r, 7).Value2): If Not pointMap.Exists(id) Then err.Raise 5, , "浸透境界の開始点 " & id & " は未定義です。"
            sp = pointMap(id)
            id = CLng(ws.cells(r, 8).Value2): If Not pointMap.Exists(id) Then err.Raise 5, , "浸透境界の終了点 " & id & " は未定義です。"
            ep = pointMap(id)
            kind = KindCode(CStr(ws.cells(r, 10).Value2))
            value = 0#
            If RPX_WTIsUnknown(CStr(ws.cells(r, 11).Value2)) Then
                If Not SUnkReady Or unk >= SUnkCount Then err.Raise 5, , "浸透境界 " & ws.cells(r, 5).Value2 & " の値が未知数（?）です。水の与え方「地下水面に合うよう境界の水位を推定する」で実行してください。"
                value = SUnkVal(unk): unk = unk + 1
            ElseIf Len(CStr(ws.cells(r, 11).Value2)) > 0 Then
                value = CDbl(ws.cells(r, 11).Value2)
            End If
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
                                    If sxy(node, 1) <= value + 0.000000001 Then SetNodeBc node, 1, value, pri
                                Case 3
                                    If sxy(node, 1) <= value + 0.000000001 Then SetNodeBc node, 1, value, pri Else SetNodeBc node, 2, 0#, pri
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
    If Not hasHead Then err.Raise 5, , "水頭固定の境界がありません（水頭・水位を指定してください）。"
    ' settings
    RPX_SeepDefaults
    SKrMin = SettingValue(ws, ROW_KRMIN, SKrMin): SBeta = SettingValue(ws, ROW_BETA, SBeta)
    SRelax = SettingValue(ws, ROW_RELAX, SRelax): STol = SettingValue(ws, ROW_TOL, STol)
    SMaxIt = CLng(SettingValue(ws, ROW_MAXIT, SMaxIt)): SGammaW = SettingValue(ws, ROW_GAMMAW, SGammaW)
    If SRelax <= 0# Or SRelax > 1# Then err.Raise 5, , "緩和係数は0より大きく1以下にしてください。"
    If SKrMin <= 0# Or SKrMin >= 1# Then err.Raise 5, , "krminは0より大きく1未満にしてください。"
End Sub

' 浸透解析ボタン：要素定義からメッシュ生成→境界割当→解析→浸透結果・図。極限解析の状態は復元する。
Public Sub RPX_SeepRun()
    Dim ws As Worksheet, st As SeepSnapshot, previousWrite As Boolean, target As Long, failText As String, started As Double, mode As Long
    If RPX_Busy Then SeepMsg "解析またはメッシュ生成を実行中です。", vbExclamation: Exit Sub
    Set ws = OptionalSheet(SEEP_SHEET)
    If ws Is Nothing Then RPX_SeepSetupSheet: SeepMsg "浸透入力シートを作成しました。入力後にもう一度実行してください。", vbInformation: Exit Sub
    If Not IsV2(ws) Then RPX_SeepSetupSheet
    SaveState st
    previousWrite = RPX_InternalWrite
    RPX_Busy = True: RPX_InternalWrite = True: RPX_WaterSuspend = True: started = Timer
    Application.ScreenUpdating = False
    On Error GoTo ErrHandler
    Application.StatusBar = "浸透解析：メッシュ生成中…"
    Call RPX_ReadInputs
    mode = RPX_SeepMode(ws)
    If ValidateInputs(ws, mode, failText) > 0 Then
        failText = "入力に問題があります（浸透入力の赤い行・確認欄）。" & failText
        GoTo Finish
    End If
    failText = ""
    target = CLng(SettingValue(ws, ROW_TARGET, 0#))
    If target > 0 Then DTarget = target
    Call RPX_GenerateMesh
    SUnkReady = False: WTInvResult = ""
    Select Case mode
        Case 0
            BuildSeepModel ws
            Application.StatusBar = "浸透解析：反復計算中（" & SM & "要素）…"
            RPX_SeepSolve
            RPX_SeepWriteResults
        Case 1, 2
            CopyMeshOnly ws
            RPX_WTReadTable ws, True, 13
            RPX_WTField mode
            RPX_SeepWriteResults
        Case 3
            RPX_WTInverse ws, CLng(SettingValue(ws, ROW_INVMAX, 40#))
    End Select
    RPX_WTWriteMode mode
    GoTo Finish
ErrHandler:
    failText = err.description
    Resume Finish
Finish:
    On Error Resume Next
    RestoreState st
    SUnkReady = False
    RPX_WaterSuspend = False
    RPX_Busy = False: RPX_InternalWrite = previousWrite
    Application.StatusBar = False: Application.ScreenUpdating = True
    On Error GoTo 0
    If Len(failText) > 0 Then
        SeepMsg "浸透解析に失敗しました：" & failText, vbExclamation
    Else
        If RPX_WaterCouplingOn() Then RPX_Invalidate
        RPX_SeepShowDiagram
        If Len(WTInvResult) > 0 Then SeepMsg WTInvResult, vbInformation, "観測水位からの逆推定"
        If Not SConverged Then
            SeepMsg "浸透解析は最大反復回数で打ち切りました（最終変化量 " & Format(SLastDelta, "0.0E+00") & " m）。βを小さく、または緩和係数を小さくしてください。", vbExclamation
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
        n = res.cells(res.rows.count, 2).End(xlUp).row - 1
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
    For i = ws.Shapes.count To 1 Step -1
        If ws.Shapes(i).AlternativeText = tag Then ws.Shapes(i).Delete
    Next i
End Sub

Private Sub DrawFreeSurface(ByVal ws As Worksheet, ByVal xmin As Double, ByVal ymax As Double, ByVal factor As Double)
    Dim res As Worksheet, n As Long, data As Variant, i As Long, s As shape
    Set res = OptionalSheet("浸透結果"): If res Is Nothing Then Exit Sub
    n = res.cells(res.rows.count, 13).End(xlUp).row - 1
    If n < 1 Then Exit Sub
    data = res.Range("M2").Resize(n, 4).Value2
    If n = 1 Then
        If Not IsNumeric(data(1, 1)) Then Exit Sub
    End If
    For i = 1 To n
        If IsNumeric(data(i, 1)) And Len(CStr(data(i, 1))) > 0 Then
            Set s = ws.Shapes.AddLine(30# + factor * (data(i, 1) - xmin), 55# + factor * (ymax - data(i, 2)), 30# + factor * (data(i, 3) - xmin), 55# + factor * (ymax - data(i, 4)))
            s.AlternativeText = "RPX_SEEP": s.line.ForeColor.RGB = RGB(0, 112, 192): s.line.weight = 2.25
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
    Dim s As shape, a As Long, b As Long, ff As FreeformBuilder
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
                    s.AlternativeText = "RPX_VELOCITY": s.line.Visible = False
                    s.fill.ForeColor.RGB = RGB(237, 125, 49): s.fill.Transparency = 0.85 - 0.7 * t
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

Private Function px(ByVal x As Double, ByVal xmin As Double, ByVal factor As Double) As Double
    px = 30# + factor * (x - xmin)
End Function
Private Function py(ByVal y As Double, ByVal ymax As Double, ByVal factor As Double) As Double
    py = 55# + factor * (ymax - y)
End Function

' 速度ジャンプの大きい辺を分位点基準で選び、連結成分の最長経路を推定すべり面として太線で描く。
Private Sub DrawSlip(ByVal ws As Worksheet, ByVal xmin As Double, ByVal ymax As Double, ByVal factor As Double, ByVal thr As Double)
    Dim i As Long, j As Long, k As Long, n As Long, nz As Long, a As Long, b As Long, s As shape
    Dim jump() As Double, score() As Double, vals() As Double, ref As Double, cutScore As Double, maxCount As Long
    Dim sel() As Boolean, nSel As Long, t As Double
    Dim adjStart() As Long, adjList() As Long, adjEdge() As Long, deg() As Long, fillN() As Long
    Dim comp() As Long, nComp As Long, compLen() As Double, best As Long, stack() As Long, sp As Long, v As Long, w As Long
    Dim dist() As Double, prev() As Long, Done() As Boolean, src As Long, far As Long, pass As Long, pts() As Single, chain As Long
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
            Set s = ws.Shapes.AddLine(px(RXY(a, 0), xmin, factor), py(RXY(a, 1), ymax, factor), px(RXY(b, 0), xmin, factor), py(RXY(b, 1), ymax, factor))
            t = jump(i) / ref: If t > 1# Then t = 1#
            s.AlternativeText = "RPX_VELOCITY": s.line.ForeColor.RGB = RGB(150, 90, 200): s.line.weight = 1.5 + 1# * t
            s.ZOrder 0
        End If
    Next i
    ' 連続した推定すべり面：地表節点間の最短路（重み＝辺長／(ジャンプ/参照+0.001)）
    TraceSlipSurface ws, xmin, ymax, factor, jump, ref, sel
    If ws.name = "図" Then ws.Range("A5").Value2 = "すべり線：内部辺" & n & "本中" & nSel & "本表示（参照=" & Format(slipPct, "0%") & "分位点、しきい値" & thr & "×参照）。推定すべり面の節点数 " & chainCount & arcText
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
    Dim i As Long, k As Long, v As Long, w As Long, a As Long, b As Long, e As Long, m As Long, s As shape
    Dim deg() As Long, fillN() As Long, adjStart() As Long, adjList() As Long, adjW() As Double
    Dim surf() As Boolean, nearSel() As Double, selNode() As Boolean, cand() As Long, candScore() As Double, nCand As Long
    Dim dist() As Double, prev() As Long, avgLen As Double, nSelE As Long, radius As Double, dx As Double, dy As Double, dD As Double
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
        pts(k, 1) = CSng(px(RXY(v, 0), xmin, factor)): pts(k, 2) = CSng(py(RXY(v, 1), ymax, factor)): v = prev(v)
    Next k
    Set s = ws.Shapes.AddPolyline(pts)
    s.AlternativeText = "RPX_VELOCITY": s.fill.Visible = False
    s.line.ForeColor.RGB = RGB(84, 20, 120): s.line.weight = 3.5
    s.ZOrder 0
    ' 円弧当てはめ（Kasa法）
    fit = True
    If Not OptionalSheet(SEEP_SHEET) Is Nothing Then fit = (SettingValue(OptionalSheet(SEEP_SHEET), ROW_ARCFIT, 1#) <> 0#)
    If fit And chainCount >= 4 Then DrawArcFit ws, xmin, ymax, factor, dst, src, prev
End Sub

Private Sub DrawArcFit(ByVal ws As Worksheet, ByVal xmin As Double, ByVal ymax As Double, ByVal factor As Double, ByVal dst As Long, ByVal src As Long, ByRef prev() As Long)
    Dim v As Long, n As Long, x As Double, y As Double, z As Double, sxx As Double, sxy As Double, syy As Double, sx As Double, sy As Double
    Dim sxz As Double, syz As Double, sz As Double, det As Double, dD As Double, dE As Double, dF As Double, mid As Long, k As Long
    Dim t0 As Double, t1 As Double, tm As Double, sweep As Double, pts() As Single, s As shape, nPts As Long, pi As Double
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
    If Abs(det) < 0.000000000001 Then Exit Sub
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
        pts(k + 1, 1) = CSng(px(x, xmin, factor)): pts(k + 1, 2) = CSng(py(y, ymax, factor))
    Next k
    Set s = ws.Shapes.AddPolyline(pts)
    s.AlternativeText = "RPX_VELOCITY": s.fill.Visible = False
    s.line.ForeColor.RGB = RGB(192, 0, 0): s.line.weight = 1.75: s.line.DashStyle = 4
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
    Dim lines(1 To 6) As String, n As Long, r As Long, i As Long, sh As shape, bottom As Double, t As String, found As Range, startRow As Long
    On Error Resume Next
    For r = 2 To 6
        t = CStr(ws.cells(r, 1).Value2)
        If Len(t) > 0 Then n = n + 1: lines(n) = t
    Next r
    If n = 0 Then Exit Sub
    t = CStr(ws.cells(1, 1).Value2)
    bottom = 55# + 460#
    For Each sh In ws.Shapes
        If sh.top + sh.Height > bottom Then bottom = sh.top + sh.Height
    Next sh
    startRow = 7
    Do While ws.rows(startRow).top < bottom + 10# And startRow < 5000
        startRow = startRow + 1
    Loop
    ' 旧凡例ブロックを消す
    Set found = ws.Range("A7:A5000").Find(What:=LEGEND_MARK, LookIn:=-4163, LookAt:=1)
    If Not found Is Nothing Then ws.Range(ws.cells(found.row, 1), ws.cells(found.row + 8, 1)).ClearContents
    ws.cells(startRow, 1).Value2 = LEGEND_MARK
    ws.cells(startRow, 1).Font.Bold = True
    If Len(t) > 0 Then ws.cells(startRow + 1, 1).Value2 = t: startRow = startRow + 1
    For i = 1 To n
        ws.cells(startRow + i, 1).Value2 = lines(i)
    Next i
    ws.Range(ws.cells(startRow - 1, 1), ws.cells(startRow + n, 1)).Font.Color = RGB(0, 0, 0)
    ws.Range("A2:A6").ClearContents
    ws.Range("A2:A6").Font.Color = RGB(255, 255, 255)
    err.Clear
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
    err.Clear
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
    lastOverlayError = "重ね描きエラー（" & stage & "）：" & err.number & " " & err.description
    Diag lastOverlayError
    Resume QuietEnd
QuietEnd:
    On Error Resume Next
    ws.Range("A6").Value2 = lastOverlayError
    Debug.Print lastOverlayError
    RelocateLegend ws
    err.Clear
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
    savedField = RUpperField: hadField = (err.number = 0): err.Clear
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
        For i = ws.Shapes.count To 1 Step -1: ws.Shapes(i).Delete: Next i
        lastOverlayError = ""
        OverlayOn ws
        cnt = 0: poly = 0
        For i = 1 To ws.Shapes.count
            If ws.Shapes(i).AlternativeText = "RPX_VELOCITY" Then
                cnt = cnt + 1
                If ws.Shapes(i).line.weight >= 3.4 Then poly = poly + 1
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
    msg = msg & "自己テスト失敗：" & err.number & " " & err.description & vbLf
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