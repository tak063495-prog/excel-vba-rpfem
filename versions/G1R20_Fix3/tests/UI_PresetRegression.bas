Attribute VB_Name = "UI_PresetRegression"
Option Explicit
' TEST COPY ONLY. Current flags below are state-machine sentinels, not fields,
' roots, solver statuses or numerical certificates. No solver is invoked.
Private Sub UI_RunPreset(ByVal entryIndex As Long)
    Select Case entryIndex
        Case 0: Call RPX_AccuracyPreset
        Case 1: Call RPX_Speed2pctPreset
        Case 2: Call RPX_G20ImprovementPreset
        Case Else: err.Raise 5, "UI_PresetRegression", "INVALID_TEST_ENTRY"
    End Select
End Sub
Private Function UI_FindRow(ByVal key As String) As Long
    Dim row As Long, ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("設定")
    For row = 4 To 1000
        If CStr(ws.Cells(row, 4).Value2) = key Then UI_FindRow = row: Exit Function
    Next row
End Function
Public Function UI_G20Keys() As Variant
    Dim keys As Variant, result(0 To 6) As Variant, i As Long, row As Long
    keys = Array("ADAPT_BALANCED_LOWER_TARGET", "ADAPT_UPPER_FLOW_PASSES", "ACCURACY_DLL_DUAL_REGULARIZATION", "ACCURACY_DLL_DUAL_LINEAR_REPAIR", "ACCURACY_DLL_UPPER_LP_SECONDS", "ACCURACY_DLL_UPPER_RECENTER", "FS_TERMINAL_SCHEDULE")
    For i = 0 To 6
        row = UI_FindRow(CStr(keys(i)))
        If row = 0 Then result(i) = Array(keys(i), False, Empty) Else result(i) = Array(keys(i), True, ThisWorkbook.Worksheets("設定").Cells(row, 2).Value2)
    Next i
    UI_G20Keys = result
End Function
Public Function UI_PresetCase(ByVal entryIndex As Long, ByVal testKind As String, ByVal nested As Boolean) As Variant
    Dim ws As Worksheet, oldWrite As Boolean, oldBusy As Boolean, oldCancel As Boolean, oldMesh As Boolean, oldResult As Boolean
    Dim caught As Long, caughtSource As String, caughtMessage As String, actualWrite As Boolean, actualMesh As Boolean, actualResult As Boolean
    Dim gapRow As Long, modeRow As Long, oldGap As Variant, oldMode As Variant, alteredGap As Boolean, alteredMode As Boolean
    Dim failedNumber As Long, failedSource As String, failedMessage As String
    oldWrite = RPX_InternalWrite: oldBusy = RPX_Busy: oldCancel = XCancel
    oldMesh = RPX_MeshCurrent: oldResult = RPX_ResultCurrent
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets("設定")
    If oldBusy Or ws.ProtectContents Then err.Raise 5, "UI_PresetRegression", "TEST_REQUIRES_IDLE_UNPROTECTED_COPY"
    Select Case testKind
        Case "protected"
            ws.Protect Password:="ui-regression-only", UserInterfaceOnly:=False
        Case "late_missing_gap"
            gapRow = UI_FindRow("GAP"): modeRow = UI_FindRow("MODE")
            If gapRow = 0 Or modeRow = 0 Then err.Raise 5, "UI_PresetRegression", "MISSING_PRECONDITION"
            oldGap = ws.Range(ws.Cells(gapRow, 1), ws.Cells(gapRow, 4)).Value2
            oldMode = ws.Cells(modeRow, 2).Value2
            ws.Range(ws.Cells(gapRow, 1), ws.Cells(gapRow, 4)).ClearContents: alteredGap = True
            ws.Cells(modeRow, 2).Value2 = "Load": alteredMode = True
        Case "busy", "success", "migration"
        Case Else
            err.Raise 5, "UI_PresetRegression", "INVALID_TEST_KIND"
    End Select
    RPX_InternalWrite = nested: RPX_Busy = (testKind = "busy")
    RPX_MeshCurrent = True: RPX_ResultCurrent = True
    On Error Resume Next
    Call UI_RunPreset(entryIndex)
    caught = err.Number: caughtSource = err.Source: caughtMessage = err.Description: err.Clear
    On Error GoTo Failed
    actualWrite = RPX_InternalWrite
    If testKind = "protected" Then ws.Unprotect Password:="ui-regression-only"
    Select Case testKind
        Case "protected"
            If caught <> 1004 Then err.Raise 5, "UI_PresetRegression", "EXPECTED_ORIGINAL_PROTECTED_ERROR_1004"
        Case "late_missing_gap"
            If caught <> 5 Or caughtSource <> "RPX_UI" Or caughtMessage <> "ACCURACY_SETTING_MISSING GAP" Then err.Raise 5, "UI_PresetRegression", "ORIGINAL_LATE_ERROR_NOT_PRESERVED"
            If ws.Cells(modeRow, 2).Value2 <> "Fs" Then err.Raise 5, "UI_PresetRegression", "LATE_FAILURE_DID_NOT_EXERCISE_PARTIAL_WRITE"
        Case "busy"
            If caught <> 5 Or caughtSource <> "RPX_UI" Or caughtMessage <> "ANALYSIS_IS_RUNNING" Then err.Raise 5, "UI_PresetRegression", "BUSY_ERROR_CHANGED"
        Case Else
            If caught <> 0 Then err.Raise 5, "UI_PresetRegression", "UNEXPECTED_PRESET_ERROR " & caught & ":" & caughtSource & ":" & caughtMessage
    End Select
    If actualWrite <> nested Then err.Raise 5, "UI_PresetRegression", "INTERNAL_WRITE_SCOPE_NOT_RESTORED"
    If XCancel <> oldCancel Then err.Raise 5, "UI_PresetRegression", "PRESET_CHANGED_CANCELLATION"
    actualMesh = RPX_MeshCurrent: actualResult = RPX_ResultCurrent
    If testKind = "busy" Or nested Then
        If Not actualMesh Or Not actualResult Then err.Raise 5, "UI_PresetRegression", "BUSY_OR_NESTED_SCOPE_INVALIDATED_OUTER_STATE"
    Else
        If actualMesh Or actualResult Then err.Raise 5, "UI_PresetRegression", "ORDINARY_SCOPE_FAILED_TO_INVALIDATE"
    End If
    If Not nested And testKind <> "busy" Then
        ' Subsequent edits/invalidation must remain effective after recovery.
        RPX_MeshCurrent = True: RPX_ResultCurrent = True
        Call RPX_Invalidate
        If RPX_MeshCurrent Or RPX_ResultCurrent Then err.Raise 5, "UI_PresetRegression", "FUTURE_INVALIDATION_BLOCKED"
    End If
    UI_PresetCase = Array("PASS", entryIndex, testKind, nested, caught, caughtSource, caughtMessage, actualWrite, actualMesh, actualResult, "solver0;state_sentinels_only")
Cleanup:
    On Error Resume Next
    If Not ws Is Nothing Then
        If ws.ProtectContents Then ws.Unprotect Password:="ui-regression-only"
        If alteredGap Then ws.Range(ws.Cells(gapRow, 1), ws.Cells(gapRow, 4)).Value2 = oldGap
        If alteredMode Then ws.Cells(modeRow, 2).Value2 = oldMode
    End If
    RPX_InternalWrite = oldWrite: RPX_Busy = oldBusy: XCancel = oldCancel
    RPX_MeshCurrent = oldMesh: RPX_ResultCurrent = oldResult
    On Error GoTo 0
    Exit Function
Failed:
    failedNumber = err.Number: failedSource = err.Source: failedMessage = err.Description
    UI_PresetCase = Array("FAIL", entryIndex, testKind, nested, failedNumber, failedSource, failedMessage)
    Resume Cleanup
End Function
