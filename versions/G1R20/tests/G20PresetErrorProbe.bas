Attribute VB_Name = "G20PresetErrorProbe"
Option Explicit
' Test copy only. Deliberately reject a protected-sheet setting write.
Public Function G20PresetErrorCleanupUnit() As Variant
    Dim ws As Worksheet, row As Long, target As Long, oldLocked As Boolean, oldWrite As Boolean
    Dim number As Long, message As String, caught As Long
    oldWrite = RPX_InternalWrite
    On Error GoTo Failed
    If oldWrite Then err.Raise 5, , "REQUIRES_ORDINARY_WRITE_SCOPE"
    Set ws = ThisWorkbook.Worksheets("設定")
    If ws.ProtectContents Then err.Raise 5, , "REQUIRES_UNPROTECTED_TEST_COPY"
    For row = 4 To 1000
        If ws.Cells(row, 4).Value2 = "GAP" Then target = row: Exit For
    Next row
    If target = 0 Then err.Raise 5, , "MISSING_GAP_SETTING"
    oldLocked = ws.Cells(target, 2).Locked
    ws.Cells(target, 2).Locked = True
    ws.Protect Password:="g20-test-only", UserInterfaceOnly:=False
    On Error Resume Next
    Call RPX_G20ImprovementPreset
    caught = err.number: err.Clear
    On Error GoTo Failed
    ws.Unprotect Password:="g20-test-only"
    ws.Cells(target, 2).Locked = oldLocked
    If caught = 0 Or RPX_InternalWrite <> oldWrite Then err.Raise 5, , "PRESET_FAILED_TO_RESTORE_WRITE_FLAG"
    G20PresetErrorCleanupUnit = Array("PASS", caught, RPX_InternalWrite, "protected_write_rejected;original_error_rethrown;write_flag_restored")
    Exit Function
Failed:
    number = err.number: message = err.description
    On Error Resume Next
    If Not ws Is Nothing Then
        ws.Unprotect Password:="g20-test-only"
        If target <> 0 Then ws.Cells(target, 2).Locked = oldLocked
    End If
    RPX_InternalWrite = oldWrite
    G20PresetErrorCleanupUnit = Array("FAIL", number, message)
End Function
