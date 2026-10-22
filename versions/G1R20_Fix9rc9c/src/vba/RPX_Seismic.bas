Attribute VB_Name = "RPX_Seismic"
Option Explicit
' RPX_Seismic : 水平震度法（G1R15）。浸透モジュールとは独立して動作する。
'   入力：設定シート SEISMIC_KH（水平震度 kh、0=地震なし）、SEISMIC_DIR（慣性力の向き 右(+x)/左(-x)）。
'   各要素の体積力に水平成分 ±kh・γ を加える（自重と同じ配列＝固定/比例の扱いも自重に従う）。
'   連成ONでも有効応力の体積力（γ - grad u）に kh・γ（全単位体積重量）を加える。
Private cacheKh As Double, cacheText As String

Private Function SeismicEmptyRow(ByVal ws As Worksheet) As Long
    Dim row As Long
    For row = 15 To RPX_SETTING_LAST_ROW
        If Application.CountA(ws.Range(ws.cells(row, 1), ws.cells(row, 4))) = 0 Then SeismicEmptyRow = row: Exit Function
    Next row
    err.Raise 5, "RPX_Seismic", "SETTINGS_TABLE_FULL SEISMIC"
End Function
Public Sub RPX_SeismicSetup()
    Dim ws As Worksheet, r As Long, rowKh As Long, rowDir As Long, key As String
    Set ws = ThisWorkbook.Worksheets("設定")
    For r = 4 To RPX_SETTING_LAST_ROW
        If IsError(ws.cells(r, 4).Value2) Then err.Raise 5, "RPX_Seismic", "INVALID_SETTING_KEY row=" & r
        key = UCase$(Trim$(CStr(ws.cells(r, 4).Value2)))
        If key = "SEISMIC_KH" Then
            If rowKh > 0 Then err.Raise 5, "RPX_Seismic", "DUPLICATE_SETTING SEISMIC_KH"
            rowKh = r
        End If
        If key = "SEISMIC_DIR" Then
            If rowDir > 0 Then err.Raise 5, "RPX_Seismic", "DUPLICATE_SETTING SEISMIC_DIR"
            rowDir = r
        End If
    Next r
    ' Add only the missing key; preserve an existing value/formula and row.
    If rowKh = 0 Then
        rowKh = SeismicEmptyRow(ws)
        ws.cells(rowKh, 1).Value2 = "地震：水平震度 kh": ws.cells(rowKh, 2).Value2 = 0
        ws.cells(rowKh, 3).Value2 = "水平震度法。0=地震なし。全要素に kh・γ を水平に作用"
        ws.cells(rowKh, 4).Value2 = "SEISMIC_KH"
    End If
    If rowDir = 0 Then
        rowDir = SeismicEmptyRow(ws)
        ws.cells(rowDir, 1).Value2 = "地震：慣性力の向き": ws.cells(rowDir, 2).Value2 = "右(+x)"
        ws.cells(rowDir, 3).Value2 = "右(+x)／左(-x)。斜面が下る向き（谷側）を選ぶと危険側"
        ws.cells(rowDir, 4).Value2 = "SEISMIC_DIR"
    End If
    ws.Range(ws.cells(rowKh, 1), ws.cells(rowKh, 3)).interior.Color = RGB(255, 242, 204)
    ws.Range(ws.cells(rowDir, 1), ws.cells(rowDir, 3)).interior.Color = RGB(255, 242, 204)
    ws.cells(rowKh, 1).Font.Bold = True: ws.cells(rowDir, 1).Font.Bold = True
    With ws.cells(rowKh, 2).Validation
        .Delete
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="0", Formula2:="1"
        .ErrorMessage = "水平震度 kh は 0 以上 1 以下の数値です（目安 0.14～0.25）。"
    End With
    ws.cells(rowKh, 2).NumberFormat = "0.00"
    With ws.cells(rowDir, 2).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="右(+x),左(-x)"
    End With
End Sub

' 符号付き kh（+x=右）。入力誤りは理由つきで解析へ返す。
Public Function RPX_SeismicKhSigned() As Double
    Dim v As Variant, d As String, kh As Double, sg As Double
    v = RPX_Setting("SEISMIC_KH", 0)
    If IsError(v) Then err.Raise 5, , "地震：水平震度 kh がExcelのエラー値です。"
    If IsEmpty(v) Or Len(Trim$(CStr(v))) = 0 Then v = 0
    If Not IsNumeric(v) Then err.Raise 5, , "地震：水平震度 kh が数値ではありません（" & CStr(v) & "）。"
    kh = CDbl(v)
    If kh < 0# Or kh > 1# Then err.Raise 5, , "地震：水平震度 kh は 0 以上 1 以下です（入力 " & kh & "）。向きは「慣性力の向き」で指定してください。"
    d = Replace(Replace(Trim$(CStr(RPX_Setting("SEISMIC_DIR", "右(+x)"))), "（", "("), "）", ")")
    Select Case d
        Case "右(+x)", "右", "+x", "+X", "x", "X", "": sg = 1#
        Case "左(-x)", "左", "-x", "-X": sg = -1#
        Case Else: err.Raise 5, , "地震：慣性力の向きは 右(+x) または 左(-x) です（入力 " & d & "）。"
    End Select
    RPX_SeismicKhSigned = sg * kh
End Function

Public Function RPX_SeismicNote() As String
    Dim k As Double
    On Error Resume Next
    k = RPX_SeismicKhSigned()
    If err.number <> 0 Then RPX_SeismicNote = "地震入力エラー：" & err.description: Exit Function
    If k = 0# Then RPX_SeismicNote = "地震：なし" Else RPX_SeismicNote = "地震：kh=" & Format(Abs(k), "0.000") & IIf(k > 0#, " 右(+x)", " 左(-x)")
End Function

' 要素 e の水平体積力に ±kh・γ を設定（addOnly=False：自重組立直後、上書き）／加算（True：連成の再組立後）。
' e=0 のとき設定を読み直す（1回の組立で1回だけ読む）。
Public Sub RPX_SeismicBody(ByVal e As Long, Optional ByVal addOnly As Boolean = False)
    Dim s As Double
    If e = 0 Then cacheKh = RPX_SeismicKhSigned()
    If cacheKh = 0# Then Exit Sub
    s = cacheKh * RMaterials(RMatId(e)).gamma
    If RPX_LoadPhase = "reference_total" Or RGravityMode = "reference" Then
        If addOnly Then RBody1(e, 0) = RBody1(e, 0) + s Else RBody1(e, 0) = s
    Else
        If addOnly Then RBody0(e, 0) = RBody0(e, 0) + s Else RBody0(e, 0) = s
    End If
End Sub


' 図と解析結果に地震条件を表示（RPX_DrawCore の末尾から呼ぶ）。
' G1R15修正：図形（テキストボックス・矢印）の追加は解析直後の再描画でExcelが異常終了する例があったため、
' セルへの文字表示だけにする（図形は使わない。旧版の図形が残っていれば削除）。
Public Sub RPX_SeismicOverlay()
    Dim ws As Worksheet, i As Long, note As String, k As Double
    On Error Resume Next
    note = RPX_SeismicNote(): k = 0#: k = RPX_SeismicKhSigned()
    If k > 0# Then note = note & " →" Else If k < 0# Then note = note & " ←"
    Set ws = RPX_Sheet("解析結果")
    ws.Range("H1").Value2 = "地震条件": ws.Range("H2").Value2 = note
    ws.Range("H1").Font.Bold = True
    Set ws = RPX_Sheet("図")
    For i = ws.Shapes.count To 1 Step -1
        If ws.Shapes(i).AlternativeText = "RPX_SEISMIC" Then ws.Shapes(i).Delete
    Next i
    With ws.Range("M1")
        .Value2 = note
        .Font.Bold = True: .Font.size = 11
        .Font.Color = IIf(k = 0#, RGB(90, 90, 90), RGB(192, 0, 0))
    End With
End Sub

' ---------------------------------------------------------------- 検証
' 浸透なし（連成OFF）で kh を変えて Fs を求め、向きの反転（鏡像形状）も確認する。
' 結果は戻り値と logPath（任意）に書く。設定・形状は終了時に元へ戻す。
Public Function RPX_SeismicSelfTest(Optional ByVal target As Long = 300, Optional ByVal logPath As String = "", Optional ByVal khList As String = "0,0.1,0.15,0.2,0.25") As String
    Dim ws As Worksheet, wd As Worksheet, saved(1 To 4) As Variant, rKh As Long, rDir As Long, r As Long
    Dim out As String, ks() As String, i As Long, xmax As Double, oldX As Variant, n As Long, t As Double
    Set ws = RPX_Sheet("設定"): Set wd = RPX_Sheet("要素定義")
    RPX_SeismicSetup
    For r = 4 To RPX_SETTING_LAST_ROW
        If UCase$(Trim$(CStr(ws.cells(r, 4).Value2))) = "SEISMIC_KH" Then rKh = r
        If UCase$(Trim$(CStr(ws.cells(r, 4).Value2))) = "SEISMIC_DIR" Then rDir = r
    Next r
    saved(1) = ws.Range("B5").Value2: saved(2) = ws.Range("B7").Value2
    saved(3) = ws.cells(rKh, 2).Value2: saved(4) = ws.cells(rDir, 2).Value2
    n = Application.WorksheetFunction.CountA(wd.Range("A7:A1006"))
    oldX = wd.Range("B7:B" & 6 + n).Value2
    On Error GoTo Failed
    ws.Range("B5").Value2 = target: ws.Range("B7").Value2 = 0
    out = "水平震度法 自己検証（浸透なし・要素数目標 " & target & "・適応OFF）" & vbLf
    ks = Split(khList, ",")
    For i = 0 To UBound(ks)
        out = out & OneCase(ws, rKh, rDir, CDbl(ks(i)), "右(+x)", "原形状") & vbLf
        out = out & OneCase(ws, rKh, rDir, CDbl(ks(i)), "左(-x)", "原形状") & vbLf
    Next i
    ' 鏡像形状：x → xmax - x。kh 右(+x) の原形状 と kh 左(-x) の鏡像形状は一致するはず。
    xmax = Application.WorksheetFunction.Max(wd.Range("B7:B" & 6 + n))
    For r = 7 To 6 + n: wd.cells(r, 2).Value2 = xmax - CDbl(oldX(r - 6, 1)): Next r
    out = out & OneCase(ws, rKh, rDir, 0#, "右(+x)", "鏡像") & vbLf
    out = out & OneCase(ws, rKh, rDir, 0.15, "左(-x)", "鏡像") & vbLf
    out = out & OneCase(ws, rKh, rDir, 0.15, "右(+x)", "鏡像") & vbLf
Done:
    On Error Resume Next
    wd.Range("B7:B" & 6 + n).Value2 = oldX
    ws.Range("B5").Value2 = saved(1): ws.Range("B7").Value2 = saved(2)
    ws.cells(rKh, 2).Value2 = saved(3): ws.cells(rDir, 2).Value2 = saved(4)
    RPX_SeismicSelfTest = out
    If Len(logPath) > 0 Then
        Dim f As Integer: f = FreeFile
        Open logPath For Output As #f: Print #f, out: Close #f
    End If
    Exit Function
Failed:
    out = out & "FAIL " & err.description & vbLf
    Resume Done
End Function

Private Function OneCase(ByVal ws As Worksheet, ByVal rKh As Long, ByVal rDir As Long, ByVal kh As Double, ByVal dirText As String, ByVal tag As String) As String
    Dim t As Double
    ws.cells(rKh, 2).Value2 = kh: ws.cells(rDir, 2).Value2 = dirText
    t = RPX_Clock()
    Call RPX_CreateMesh
    If Len(RPX_LastError) > 0 Then OneCase = tag & " kh=" & kh & " " & dirText & " メッシュ失敗：" & RPX_LastError: Exit Function
    Call RPX_Run
    If Len(RPX_LastError) > 0 Then OneCase = tag & " kh=" & kh & " " & dirText & " 解析失敗：" & RPX_LastError: Exit Function
#If False Then
    OneCase = tag & vbTab & "kh=" & Format(kh, "0.00") & vbTab & dirText & vbTab & "Fs下界=" & Format(RLowerFs, "0.0000") & vbTab & "Fs上界=" & Format(RUpperFs, "0.0000") & vbTab & Format(Timer - t, "0.0") & "s"
#Else
    OneCase = tag & vbTab & "kh=" & Format(kh, "0.00") & vbTab & dirText & vbTab & ("Fs" & ChrW$(19979) & ChrW$(30028) & "=") & Format(RLowerFs, "0.0000") & vbTab & ("Fs" & ChrW$(19978) & ChrW$(30028) & "=") & Format(RUpperFs, "0.0000") & vbTab & Format(RPX_Clock() - t, "0.0") & "s"
#End If
End Function
