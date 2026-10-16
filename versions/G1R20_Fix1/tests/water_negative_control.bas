Attribute VB_Name = "WaterCacheCollisionProbe"
Option Explicit

' TEST ONLY: import into a disposable copy of the compiled G1R20 workbook.
' No solver, private cache access, forced K9, fabricated convergence flag or
' edited result vector is used. The normal public mode1 prescribed-field
' producer RPX_WTReadTable -> RPX_WTField -> RPX_SeepWriteResults writes twice.
' WaterApply between writes loads the real cache. This proves the automated
' producer/load/producer path, not that a long limit solve fits in one second.

Private Function Elapsed(ByVal start As Double) As Double
    Elapsed = Timer - start
    If Elapsed < 0# Then Elapsed = Elapsed + 86400#
End Function

Private Sub NextNaturalSecond()
    Dim original As String, started As Double
    original = Format(Now, "yyyymmddhhnnss"): started = Timer
    Do While Format(Now, "yyyymmddhhnnss") = original
        DoEvents
        If Elapsed(started) > 2# Then Err.Raise 5, , "CLOCK_WAIT_TIMEOUT"
    Loop
End Sub

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
    ' Setting K2=False mirrors normal UI invalidation; K9 is never assigned.
    RPX_Sheet("浸透結果").Range("K2").Value2 = False
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

Public Function WaterCacheCollisionRun() As String
    Dim ws As Worksheet, result As Worksheet, settings As Worksheet, row As Long
    Dim oldWrite As Boolean, oldEvents As Boolean, oldScreen As Boolean, oldSuspend As Boolean
    Dim stamp1 As String, stamp2 As String, firstHead As Double, secondHead As Double
    Dim expected As Double, firstBx As Double, secondBx As Double, seconds As Double, started As Double
    Dim attempt As Long, data As String, errorNumber As Long, errorText As String
    If RPX_Busy Then Err.Raise 5, , "ANALYSIS_IS_RUNNING"
    oldWrite = RPX_InternalWrite: oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating: oldSuspend = RPX_WaterSuspend
    On Error GoTo Failed
    Application.EnableEvents = False: Application.ScreenUpdating = False
    RPX_InternalWrite = True: RPX_WaterSuspend = False: XCancel = False
    Call RPX_InputEnd
    Call RPX_SeepSetupSheet
    Set ws = RPX_Sheet("浸透入力"): Set result = RPX_Sheet("浸透結果")
    ' Z5 is the actual V2 label of mode1; this is a valid ordinary input mode.
    ws.Range("B3").Value2 = ws.Range("Z5").Value2
    ws.Range("B4").Value2 = 9.81: ws.Range("D4").Value2 = 1: ws.Range("F4").Value2 = 0
    ws.Range("M7:N206").ClearContents
    Set settings = RPX_Sheet("設定")
    For row = 4 To 100
        If CStr(settings.Cells(row, 4).Value2) = "SEISMIC_KH" Then settings.Cells(row, 2).Value2 = 0
    Next row
    Call UnitSquareArrays
    ' Warm formatting/first output outside the timed window. The following
    ' natural second boundary ensures the first trial refreshes an old cache.
    Call ProduceMode1(ws, 0#)
    For attempt = 1 To 3
        Call NextNaturalSecond
        started = Timer
        Call ProduceMode1(ws, 0#)
        stamp1 = CStr(result.Range("K9").Value2)
        Call RPX_InputBegin
        Call RPX_WaterApply
        firstHead = RPX_WaterHeadAt(0.25, 0.1): firstBx = RBody0(0, 0)
        Call RPX_InputEnd
        Call ProduceMode1(ws, 1#)
        stamp2 = CStr(result.Range("K9").Value2)
        expected = RPX_WTHead(0.25, 0.1, 1)
        Call RPX_InputBegin
        Call RPX_WaterApply
        secondHead = RPX_WaterHeadAt(0.25, 0.1): secondBx = RBody0(0, 0)
        Call RPX_InputEnd
        seconds = Elapsed(started)
        data = "attempt=" & attempt & ";seconds=" & seconds & ";K9_first=" & stamp1 & ";K9_second=" & stamp2
        data = data & ";first_head=" & firstHead & ";second_expected=" & expected & ";second_observed=" & secondHead
        data = data & ";first_body_x=" & firstBx & ";second_body_x=" & secondBx & ";expected_body_x=-4.905"
        data = data & ";sheet_D2=" & result.Range("D2").Value2 & ";K2=" & result.Range("K2").Value2
        Debug.Print data
        If stamp1 = stamp2 Then
            If Abs(firstHead - 2#) > 0.000000001 Or Abs(expected - 1.175) > 0.000000001 Then Err.Raise 5, , "FIXTURE_EXPECTATION_FAILED " & data
            If Abs(secondHead - expected) <= 0.000000001 And Abs(secondBx + 4.905) <= 0.000000001 Then
                WaterCacheCollisionRun = "NO_STALE_REUSE " & data
            ElseIf Abs(secondHead - firstHead) <= 0.000000001 And Abs(secondBx - firstBx) <= 0.000000001 And result.Range("D2").Value2 = 1# Then
                WaterCacheCollisionRun = "REPRODUCED_STALE_HEAD solver0;natural_ID_collision=1;" & data
            Else
                WaterCacheCollisionRun = "UNEXPECTED_STATE " & data
            End If
            GoTo Done
        End If
    Next attempt
    WaterCacheCollisionRun = "INCONCLUSIVE no natural same-second collision;" & data
Done:
    On Error Resume Next
    Call RPX_InputEnd
    RPX_ResultCurrent = False: RPX_MeshCurrent = False
    RPX_WaterSuspend = oldSuspend: RPX_InternalWrite = oldWrite
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    On Error GoTo 0
    Exit Function
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    WaterCacheCollisionRun = "ERROR " & errorNumber & " " & errorText
    Resume Done
End Function
