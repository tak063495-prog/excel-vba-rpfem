Attribute VB_Name = "WaterCacheFix1Probe"
Option Explicit

' TEST ONLY: disposable workbook, product code unchanged; solver0.
Private Sub UnitSquareArrays()
    ' A valid two-triangle unit square; matching seep/limit geometry.
    SN = 4: SM = 2: SNMat = 1
    ReDim sxy(0 To 3, 0 To 1): ReDim STri(0 To 1, 0 To 2)
    ReDim SMat(0 To 1): ReDim SKx(0 To 0): ReDim SKy(0 To 0)
    sxy(1, 0) = 1#: sxy(2, 0) = 1#: sxy(2, 1) = 1#: sxy(3, 1) = 1#
    STri(0, 0) = 0: STri(0, 1) = 1: STri(0, 2) = 2
    STri(1, 0) = 0: STri(1, 1) = 2: STri(1, 2) = 3
    SKx(0) = 1#: SKy(0) = 1#: SGammaW = 9.81
    rn = 4: re = 2: RM = 1: RR = 0
    RXY = sxy: RTri = STri
    ReDim RMatId(0 To 1): ReDim RRigidId(0 To 1)
    RRigidId(0) = -1: RRigidId(1) = -1
    ReDim RMaterials(0 To 0): RMaterials(0).gamma = 20#
    ReDim RBody0(0 To 1, 0 To 1): ReDim RBody1(0 To 1, 0 To 1)
    ReDim RArea(0 To 1): RArea(0) = 0.5: RArea(1) = 0.5
    ReDim RGrad(0 To 1, 0 To 2, 0 To 1)
    RGrad(0, 0, 0) = -1#: RGrad(0, 1, 0) = 1#: RGrad(0, 1, 1) = -1#: RGrad(0, 2, 1) = 1#
    RGrad(1, 0, 1) = -1#: RGrad(1, 1, 0) = 1#: RGrad(1, 2, 0) = -1#: RGrad(1, 2, 1) = 1#
    RGravityMode = "fixed": RPX_LoadPhase = "raw"
End Sub

Private Sub ProduceMode1(ByVal ws As Worksheet, ByVal slope As Double)
    ' Same public prescribed-field producer as RPX_SeepRun's Case1 branch.
    ws.Range("M7:N8").Value2 = ArrayTable(slope)
    Call RPX_WTReadTable(ws, True, 13)
    Call RPX_WTField(1)
    Call RPX_SeepWriteResults
    Call RPX_WTWriteMode(1)
End Sub

Private Function ArrayTable(ByVal slope As Double) As Variant
    Dim table(1 To 2, 1 To 2) As Variant
    table(1, 1) = 0#: table(1, 2) = 2#
    table(2, 1) = 1#: table(2, 2) = 2# + slope
    ArrayTable = table
End Function

Private Sub WaterCheck(ByVal ok As Boolean, ByVal label As String)
    If Not ok Then Err.Raise 5, "WaterCacheFix1Probe", label
End Sub
Private Sub LogTest(ByVal path As String, ByVal label As String, ByVal detail As String)
    Dim file As Integer
    file = FreeFile: Open path For Append As #file
    Print #file, label & vbTab & detail: Close #file
End Sub
Private Function CacheNodes() As Long
    CacheNodes = CLng(Application.Run("'" & ThisWorkbook.Name & "'!RPX_TestFix1WaterCacheNodes"))
End Function
Private Function LoadRejected(ByVal token As String) As Boolean
    Dim number As Long, message As String
    On Error GoTo Rejected
    Call RPX_WaterApply
    Exit Function
Rejected:
    number = Err.Number: message = Err.Description
    On Error GoTo 0
    If number = 5 And InStr(1, message, token, vbBinaryCompare) > 0 Then
        LoadRejected = True
    Else
        Err.Raise number, "WaterCacheFix1Probe", "UNEXPECTED_REJECTION " & message
    End If
End Function
Private Function WriterRejected() As Boolean
    On Error GoTo Rejected
    Call RPX_SeepWriteResults
    Exit Function
Rejected:
    WriterRejected = True: Err.Clear
End Function
Private Sub CheckNewField()
    Call RPX_WaterApply
    WaterCheck Abs(RPX_WaterHeadAt(0.25, 0.1) - 1.175) < 0.000000001, "NEW_HEAD_NOT_LOADED"
    WaterCheck Abs(RBody0(0, 0) + 4.905) < 0.000000001, "NEW_BODY_NOT_LOADED"
End Sub

Public Function WaterCacheFix1Run(ByVal logPath As String) As String
    Const PASSWORD As String = "rpfem-fix1-probe"
    Dim ws As Worksheet, result As Worksheet, settings As Worksheet, row As Long, mode As Variant
    Dim oldWrite As Boolean, oldEvents As Boolean, oldScreen As Boolean, oldSuspend As Boolean
    Dim stamp1 As String, stamp2 As String, attempt As Long, sameSecond As Boolean, started As Double
    Dim number As Long, message As String, file As Integer, changed As Boolean
    If RPX_Busy Then Err.Raise 5, , "ANALYSIS_IS_RUNNING"
    oldWrite = RPX_InternalWrite: oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating: oldSuspend = RPX_WaterSuspend
    On Error GoTo Failed
    Application.EnableEvents = False: Application.ScreenUpdating = False
    RPX_InternalWrite = True: RPX_WaterSuspend = False: XCancel = False
    Call RPX_InputEnd: Call RPX_SeepSetupSheet
    Set ws = RPX_Sheet("浸透入力"): Set result = RPX_Sheet("浸透結果")
    ws.Range("B3").Value2 = ws.Range("Z5").Value2
    ws.Range("B4").Value2 = 9.81: ws.Range("D4").Value2 = 1: ws.Range("F4").Value2 = 0
    ws.Range("M7:N206").ClearContents
    Set settings = RPX_Sheet("設定")
    For row = 4 To 100
        If CStr(settings.Cells(row, 4).Value2) = "SEISMIC_KH" Then settings.Cells(row, 2).Value2 = 0
    Next row
    Call UnitSquareArrays
    file = FreeFile: Open logPath For Output As #file: Print #file, "test" & vbTab & "detail": Close #file

    ' No waits and no forced IDs: repeat only if a real wall-second boundary
    ' happened to separate the two complete public producer calls.
    For attempt = 1 To 5
        Call ProduceMode1(ws, 0#): stamp1 = CStr(result.Range("K9").Value2)
        Call RPX_InputBegin: Call RPX_WaterApply
        WaterCheck Abs(RPX_WaterHeadAt(0.25, 0.1) - 2#) < 0.000000001, "FIRST_HEAD"
        Call RPX_InputEnd
        Call ProduceMode1(ws, 1#): stamp2 = CStr(result.Range("K9").Value2)
        Call RPX_InputBegin: Call CheckNewField: Call RPX_InputEnd
        WaterCheck stamp1 <> stamp2, "DUPLICATE_COMPUTATION_ID"
        sameSecond = (Left$(stamp1, 14) = Left$(stamp2, 14))
        LogTest logPath, "same_second_attempt", stamp1 & ";" & stamp2 & ";expected_head=1.175;expected_body_x=-4.905;new_field=PASS"
        If sameSecond Then Exit For
    Next attempt
    WaterCheck sameSecond, "INCONCLUSIVE_NO_NATURAL_SAME_SECOND_PAIR"

    ' A writer also invalidates the warm cache without an intervening InputBegin.
    Call ProduceMode1(ws, 0#): Call RPX_InputBegin: Call RPX_WaterApply
    Call ProduceMode1(ws, 1#): Call CheckNewField: Call RPX_InputEnd
    LogTest logPath, "within_one_input_scope", "PASS head=1.175;body_x=-4.905"

    ' Every new boundary clears cached arrays, including ineligible common modes.
    For Each mode In Array(0, 1, 2, 3)
        ws.Range("B3").Value2 = ws.Cells(4 + CLng(mode), 26).Value2
        Call ProduceMode1(ws, 0#): Call RPX_WTWriteMode(CLng(mode))
        Call RPX_WaterApply: WaterCheck CacheNodes() = 4, "CACHE_NOT_WARM"
        Call RPX_InputBegin: WaterCheck CacheNodes() = 0, "INPUT_BOUNDARY_DID_NOT_CLEAR_MODE_" & mode
        Call RPX_WaterApply: Call RPX_InputEnd
        LogTest logPath, "input_boundary_mode_" & mode, "PASS cache_nodes=4->0->4"
    Next mode
    ws.Range("B3").Value2 = ws.Range("Z5").Value2

    Call ProduceMode1(ws, 0#): Call RPX_WaterApply
    result.Range("K2").Value2 = False
    WaterCheck LoadRejected("収束"), "K2_FALSE_WARM_CACHE_ACCEPTED"
    Call ProduceMode1(ws, 1#): Call CheckNewField
    LogTest logPath, "K2_false", "PASS warm-cache rejected;new successful producer recovers"

    Call ProduceMode1(ws, 0#): Call RPX_WaterApply
    result.Protect PASSWORD
    WaterCheck WriterRejected(), "PROTECTED_RESULT_WRITER_DID_NOT_FAIL"
    WaterCheck CBool(result.Range("K2").Value2), "PROTECTED_PREDECESSOR_NOT_TRUE_FIXTURE"
    Call RPX_InputBegin
    WaterCheck LoadRejected("SEEP_RESULT_WRITE_INCOMPLETE"), "FAILED_PROTECTED_WRITE_RELOADED_OLD_RESULT"
    Call RPX_InputEnd
    result.Unprotect PASSWORD
    Call ProduceMode1(ws, 1#): Call CheckNewField
    LogTest logPath, "protected_writer_failure", "PASS old_K2=True but pending rejects across InputBegin;new producer recovers"

    ' Failure after body writing begins must leave neither cache nor K2/ID valid.
    Call ProduceMode1(ws, 0#): Call RPX_WaterApply
    Erase SFlux
    WaterCheck WriterRejected(), "MIDWRITE_FAILURE_NOT_TRIGGERED"
    WaterCheck VarType(result.Range("K2").Value2) = vbBoolean, "FAILED_WRITE_K2_TYPE"
    WaterCheck Not CBool(result.Range("K2").Value2), "FAILED_WRITE_K2_TRUE"
    WaterCheck Len(CStr(result.Range("K9").Value2)) = 0, "FAILED_WRITE_ID_RETAINED"
    WaterCheck LoadRejected("SEEP_RESULT_WRITE_INCOMPLETE"), "FAILED_WRITE_WARM_CACHE_ACCEPTED"
    Call ProduceMode1(ws, 1#): Call CheckNewField
    LogTest logPath, "midwrite_failure", "PASS K2=False;K9=empty;warm-cache rejected;new producer recovers"

    ' The ordinary UI invalidator must fail closed before a writer is reached.
    Call ProduceMode1(ws, 0#): Call RPX_WaterApply
    result.Protect PASSWORD
    Call Application.Run("'" & ThisWorkbook.Name & "'!RPX_TestFix1InvalidateSeepResult")
    WaterCheck CBool(result.Range("K2").Value2), "PREWRITER_PROTECTED_FIXTURE"
    Call RPX_InputBegin
    WaterCheck LoadRejected("SEEP_RESULT_WRITE_INCOMPLETE"), "FAILED_UI_ATTEMPT_RELOADED_OLD_RESULT"
    Call RPX_InputEnd
    result.Unprotect PASSWORD
    Call ProduceMode1(ws, 1#): Call CheckNewField
    LogTest logPath, "protected_ui_invalidation_before_writer", "PASS pending persists;new producer recovers"
    WaterCacheFix1Run = "PASS natural_same_second_unique_ID;fresh_head_body;within_scope;all_modes;K2_false;protected_writer;midwrite;prewriter_failure;solver0"
Done:
    On Error Resume Next
    If Not result Is Nothing Then result.Unprotect PASSWORD
    Call RPX_InputEnd
    RPX_ResultCurrent = False: RPX_MeshCurrent = False
    RPX_WaterSuspend = oldSuspend: RPX_InternalWrite = oldWrite
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    On Error GoTo 0
    Exit Function
Failed:
    number = Err.Number: message = Err.Description
    WaterCacheFix1Run = "FAIL " & number & " " & message
    If Len(logPath) > 0 Then LogTest logPath, "FAIL", WaterCacheFix1Run
    Resume Done
End Function

