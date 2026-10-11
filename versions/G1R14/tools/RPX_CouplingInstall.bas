Attribute VB_Name = "RPX_CouplingInstall"
Option Explicit
' 既存モジュールへのフック行を CodeModule で挿入する（冪等）。
' 前提：トラストセンター →「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」をON。
' 実行後、このモジュールは削除してかまいません。

Private Function InsertAfterAnchor(ByVal moduleName As String, ByVal procName As String, ByVal anchor As String, ByVal newLine As String) As String
    Dim cm As Object, first As Long, count As Long, i As Long, hit As Long, text As String
    Set cm = ThisWorkbook.VBProject.VBComponents(moduleName).CodeModule
    first = cm.ProcStartLine(procName, 0): count = cm.ProcCountLines(procName, 0)
    For i = first To first + count - 1
        text = Trim$(cm.Lines(i, 1))
        If text = Trim$(newLine) Then InsertAfterAnchor = moduleName & "." & procName & "：既に挿入済み（" & i & "行）": Exit Function
        If text = anchor And hit = 0 Then hit = i
    Next i
    If hit = 0 Then InsertAfterAnchor = moduleName & "." & procName & "：アンカーが見つかりません［" & anchor & "］": Exit Function
    cm.InsertLines hit + 1, newLine
    InsertAfterAnchor = moduleName & "." & procName & "：" & hit + 1 & "行に挿入"
End Function

Public Sub RPX_CouplingInstall()
    Dim msg As String
    On Error GoTo Failed
    ' G1R13以降：RPX_LoadPhase="raw" がトポロジ計算より前へ移動したため、末尾の荷重段差修復の後に挿入する。
    msg = InsertAfterAnchor("RPX_Mesh", "RPX_AssignRegions", "If RPX_RepairLoadSteps() Then Call RPX_AssignRegions", "    Call RPX_WaterApply")
    If InStr(msg, "アンカーが見つかりません") > 0 Then
        msg = InsertAfterAnchor("RPX_Mesh", "RPX_AssignRegions", "RPX_LoadPhase = ""raw""", "    Call RPX_WaterApply") & "（G1R12形式）"
    End If
    msg = msg & vbLf & InsertAfterAnchor("RPX_Control", "RPX_DrawCore", "If RPX_ResultCurrent Then Call RPX_DrawVelocity", "    Call RPX_SeepOverlay")
    Debug.Print msg
    MsgBox msg & vbLf & "デバッグ→コンパイルで確認後、保存してください。", vbInformation
    Exit Sub
Failed:
    MsgBox "挿入に失敗：" & Err.Description & vbLf & "（VBAプロジェクトへのアクセス信頼を確認、または手で1行追加）", vbExclamation
End Sub
