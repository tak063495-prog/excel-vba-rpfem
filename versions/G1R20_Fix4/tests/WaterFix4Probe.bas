Attribute VB_Name = "WaterFix4Probe"
Option Explicit
Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal id As LongPtr, ByVal ms As Long, ByVal cb As LongPtr) As LongPtr
Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal id As LongPtr) As Long
Public WaterForwardCount As Long
Private timerId As LongPtr
Private stopCount As Long, stopIteration As Long, stopStatus As String, inverseStop As Boolean
Public Sub StopDuringSeep(ByVal hwnd As LongPtr, ByVal message As Long, ByVal id As LongPtr, ByVal tick As Long)
    On Error Resume Next
    If (Not inverseStop And InStr(CStr(Application.StatusBar), "浸透解析：反復計算中") > 0) Or (inverseStop And InStr(CStr(Application.StatusBar), "逆解析：順解析") > 0) Then
        stopStatus = CStr(Application.StatusBar): stopIteration = SIterations
        Call RPX_Stop: stopCount = stopCount + 1
        Call KillTimer(0, timerId): timerId = 0
    End If
End Sub
Private Sub Configure(ByVal mode As Long)
    Dim definition As Worksheet, material As Worksheet, ws As Worksheet, i As Long
    Dim coords(1 To 4, 1 To 3) As Variant, defaults As Variant
    XCancel = False: RPX_InternalWrite = False: RPX_Busy = False
    Call RPX_InputEnd: Call RPX_Setup: Call RPX_SeepSetupSheet
    RPX_InternalWrite = True
    Set definition = RPX_Sheet("要素定義"): Set material = RPX_Sheet("材料データ"): Set ws = RPX_Sheet("浸透入力")
    definition.Range("A7:C1006,E7:H1006,J7:T1006").ClearContents
    For i = 1 To 4: coords(i, 1) = i: Next i
    coords(1, 2) = 0#: coords(1, 3) = 0#: coords(2, 2) = 1#: coords(2, 3) = 0#
    coords(3, 2) = 1#: coords(3, 3) = 1#: coords(4, 2) = 0#: coords(4, 3) = 1#
    definition.Range("A7:C10").Value2 = coords
    definition.Range("E7:H7").Value2 = Array(1#, 1#, 0#, "1,2,3,4")
    definition.Range("J7:P7").Value2 = Array(1#, 1#, 1#, 2#, 0#, "固定XY", 0#): definition.Range("Q7:T7").Value2 = 0#
    material.Range("A5:Q1004").ClearContents: material.Range("A5:D5").Value2 = Array(1#, 10#, 30#, 20#)
    RPX_Sheet("設定").Range("B5").Value2 = 100#
    ws.Range("B3").Value2 = ws.Cells(4 + mode, 26).Value2
    ws.Range("B4").Value2 = 9.81: ws.Range("D4").Value2 = 1#: ws.Range("F4").Value2 = 0#
    ws.Range("A7:C206,E7:K206,M7:N206,P7:Q206").ClearContents
    ws.Range("A7:C7").Value2 = Array(1#, 1#, 1#)
    ws.Range("E7:K7").Value2 = Array(1#, 1#, 4#, 1#, 1#, "水頭", 0.8)
    ws.Range("E8:K8").Value2 = Array(2#, 1#, 2#, 3#, 0#, "水頭", 0.2)
    ws.Range("M7:N7").Value2 = Array(0#, 2#): ws.Range("M8:N8").Value2 = Array(1#, 2#)
    defaults = Array(100#, 0.0001, 0#, 0.5, 0.000001, 300#, 0.000001, 40#, 1#, 0.5, 0.95, 0.1, 1#, 0#, 0.2)
    For i = 0 To UBound(defaults): ws.cells(209 + i, 3).Value2 = defaults(i): Next i
    RPX_InternalWrite = False
End Sub
Public Function ActiveStopNegative(Optional ByVal mode As Long = 0) As Variant
    Dim ws As Worksheet, result As String, n As Long, text As String
    Configure mode: Set ws = RPX_Sheet("浸透入力")
    inverseStop = (mode = 3)
    If mode = 3 Then
        ws.Range("K7").Value2 = "?0.8"
        ws.Range("P7:Q7").Value2 = Array(0.25, 0.65)
        ws.Range("P8:Q8").Value2 = Array(0.5, 0.5)
        ws.Range("P9:Q9").Value2 = Array(0.75, 0.35)
    End If
    ws.Range("C212").Value2 = 0.1: ws.Range("C213").Value2 = 0.000000000001: ws.Range("C214").Value2 = 300#
    stopCount = 0: stopIteration = -1: stopStatus = ""
    On Error GoTo Failed
    timerId = SetTimer(0, 0, 20, AddressOf StopDuringSeep)
    If timerId = 0 Then Err.Raise 5, , "TEST_TIMER_NOT_SET"
    result = RPX_SeepRunSilent()
    GoTo Done
Failed:
    n = Err.Number: text = Err.Description
    Resume Done
Done:
    If timerId <> 0 Then Call KillTimer(0, timerId): timerId = 0
    ActiveStopNegative = Array(stopCount, stopIteration, stopStatus, SIterations, SConverged, XCancel, RPX_Sheet("浸透結果").Range("K2").Value2, result, n, text)
End Function
Public Function GammaNegative(ByVal gamma As Double) As Variant
    Dim ws As Worksheet, result As String, pressure As Double, body As Double, i As Long
    Configure 2: Set ws = RPX_Sheet("浸透入力"): ws.Range("B4").Value2 = gamma
    result = RPX_SeepRunSilent()
    If RPX_Sheet("浸透結果").Range("K2").Value2 = True Then
        pressure = RPX_Sheet("浸透結果").Range("F2").Value2
        Call RPX_ReadInputs: Call RPX_GenerateMesh
        For i = 0 To re - 1: body = body + RBody0(i, 1) / re: Next i
    End If
    GammaNegative = Array(gamma, result, SGammaW, pressure, body, RPX_Sheet("浸透結果").Range("K2").Value2, RPX_WaterApplied)
End Function
Public Function ResultMutationNegative() As Variant
    Dim ws As Worksheet, result As String, before As Double, after As Double
    Configure 1: result = RPX_SeepRunSilent()
    Call RPX_ReadInputs: Call RPX_GenerateMesh
    before = RBody0(0, 1)
    Set ws = RPX_Sheet("浸透結果")
    ws.Range("D2").Value2 = 20#
    Call RPX_InputBegin: Call RPX_ReadInputs: Call RPX_GenerateMesh
    after = RBody0(0, 1): Call RPX_InputEnd
    ResultMutationNegative = Array(result, before, after, RPX_Sheet("浸透結果").Range("K2").Value2, RPX_WaterApplied)
End Function

Public Function DirectCancelled(ByVal mode As Long) As Variant
    Dim result As String, number As Long, text As String, ws As Worksheet
    Configure mode: Set ws = RPX_Sheet("浸透入力")
    result = RPX_SeepRunSilent()
    Call RPX_Stop
    On Error GoTo Failed
    If mode = 0 Then Call RPX_SeepSolve Else Call RPX_WTField(mode)
    GoTo Done
Failed:
    number = Err.Number: text = Err.Description
    Resume Done
Done:
    DirectCancelled = Array(mode, result, number, text, SConverged, XCancel)
End Function
Public Function CommitCancelled() As Variant
    Dim result As String, number As Long, text As String, observed As Variant, journal As String
    Configure 2: result = RPX_SeepRunSilent()
    Call RPX_Stop
    On Error GoTo Failed
    Call RPX_SeepWriteResults
    GoTo Done
Failed:
    number = Err.Number: text = Err.Description
    Resume Done
Done:
    journal = CStr(ThisWorkbook.CustomDocumentProperties("RPX_SeepResultWriteState").Value)
    observed = WaterConsumeCurrent()
    CommitCancelled = Array(result, number, text, SConverged, RPX_Sheet("浸透結果").Range("K2").Value2, journal, observed(0), observed(1), observed(2))
End Function
Public Function GammaValidation(ByVal value As Variant, Optional ByVal formula As Boolean = False) As Variant
    Dim ws As Worksheet, message As String, number As Long, text As String
    Configure 2: Set ws = RPX_Sheet("浸透入力")
    If formula Then ws.Range("B4").Formula = CStr(value) Else ws.Range("B4").Value2 = value
    Call RPX_SeepSetSilent(True): Call RPX_SeepCheckInputs
    message = RPX_SeepGetLastMsg()
    GammaValidation = Array(message, RPX_SeepRunSilent(), SGammaW, RPX_Sheet("浸透結果").Range("K2").Value2)
End Function

Public Sub WaterInputConfigure(ByVal mode As Long)
    Configure mode
    With RPX_Sheet("浸透入力")
        .Range("I7").Value2 = 1#
        .Range("C209").Value2 = 32#: .Range("C216").Value2 = 12#
        If mode = 3 Then
            .Range("K7").Value2 = "?0.8"
            .Range("P7:Q7").Value2 = Array(0.25, 0.65)
            .Range("P8:Q8").Value2 = Array(0.5, 0.5)
            .Range("P9:Q9").Value2 = Array(0.75, 0.35)
        End If
    End With
End Sub
Public Function WaterInputCheck(ByVal executeSolver As Boolean) As Variant
    Dim checked As String, ran As String, number As Long, text As String, k2 As Variant
    Call RPX_SeepSetSilent(True)
    On Error GoTo Failed
    Call RPX_SeepCheckInputs: checked = RPX_SeepGetLastMsg()
    If executeSolver Then ran = RPX_SeepRunSilent()
    k2 = RPX_Sheet("浸透結果").Range("K2").Value2
    GoTo Done
Failed:
    number = Err.Number: text = Err.Description
    Resume Done
Done:
    WaterInputCheck = Array(checked, ran, number, text, k2, SConverged, SIterations, SBeta, SKrMin, SMaxIt, STol)
End Function

Public Function WaterSwitchConsumer(ByVal value As Variant, ByVal suction As Boolean, ByVal scoped As Boolean, Optional ByVal formula As Boolean = False) As Variant
    Dim message As String, ws As Worksheet, number As Long, text As String, before As Double, after As Double, cell As Range
    Configure 2: message = RPX_SeepRunSilent()
    Call RPX_ReadInputs: Call RPX_GenerateMesh: before = RBody0(0, 1)
    Set ws = RPX_Sheet("浸透入力")
    If suction Then Set cell = ws.Range("F4") Else Set cell = ws.Range("D4")
    If formula Then cell.Formula = CStr(value) Else cell.Value2 = value
    On Error GoTo Failed
    If scoped Then Call RPX_InputBegin
    Call RPX_ReadInputs: Call RPX_GenerateMesh
    after = RBody0(0, 1)
    GoTo Done
Failed:
    number = Err.Number: text = Err.Description
    Resume Done
Done:
    On Error Resume Next
    Call RPX_InputEnd
    WaterSwitchConsumer = Array(message, number, text, before, after, RPX_WaterApplied)
End Function

Public Function WaterConsumeCurrent() As Variant
    Dim number As Long, text As String
    XCancel = False: RPX_WaterSuspend = False
    On Error GoTo Rejected
    Call RPX_InputBegin: Call RPX_ReadInputs: Call RPX_GenerateMesh
    WaterConsumeCurrent = Array("ACCEPTED", 0, "", RPX_WaterApplied, RBody0(0, 1))
    Call RPX_InputEnd
    Exit Function
Rejected:
    number = Err.Number: text = Err.Description
    On Error Resume Next
    Call RPX_InputEnd
    WaterConsumeCurrent = Array("REJECTED", number, text, RPX_WaterApplied, 0#)
End Function

Public Function WaterProofSummary() As Variant
    Dim identifier As String, context As String, found As Boolean
    found = RPX_SourceRecordRead("urn:rpfem:source:seep:v1", identifier, context)
    WaterProofSummary = Array(found, identifier, Left$(context, 40), Len(context))
End Function

Public Sub WaterCountForwardEntry()
    WaterForwardCount = WaterForwardCount + 1
End Sub
Public Function WaterInverseBudget(ByVal maximum As Long) As Variant
    Dim ws As Worksheet, message As String, seg() As Double, ns As Long, found As Boolean, i As Long, x As Double
    Call WaterInputConfigure(0): message = RPX_SeepRunSilent()
    If Len(message) > 0 Or Not SConverged Then Err.Raise 5, , "BUDGET_REFERENCE_FAILED " & message
    ns = RPX_SeepFreeSurface(seg)
    If ns < 1 Then Err.Raise 5, , "BUDGET_REFERENCE_SURFACE_MISSING"
    Set ws = RPX_Sheet("浸透入力")
    ws.Range("B3").Value2 = ws.Range("Z7").Value2: ws.Range("K7").Value2 = "?0.8"
    ws.Range("C216").Value2 = maximum
    For i = 0 To 2
        x = 0.25 * (i + 1): ws.cells(7 + i, 16).Value2 = x
        ws.cells(7 + i, 17).Value2 = RPX_WTSurfaceY(seg, ns, x, found)
        If Not found Then Err.Raise 5, , "BUDGET_OBSERVATION_OUTSIDE_REFERENCE"
    Next i
    WaterForwardCount = 0: message = RPX_SeepRunSilent()
    WaterInverseBudget = Array(maximum, WaterForwardCount, RPX_Sheet("浸透結果").Range("K12").Value2, RPX_Sheet("浸透結果").Range("K11").Value2, RPX_Sheet("浸透結果").Range("K2").Value2, message)
End Function
