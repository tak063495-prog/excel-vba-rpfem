Attribute VB_Name = "WaterFix4Probe"
Option Explicit
Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal id As LongPtr, ByVal ms As Long, ByVal cb As LongPtr) As LongPtr
Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal id As LongPtr) As Long
Private timerId As LongPtr
Private stopCount As Long, stopIteration As Long, stopStatus As String
Public Sub StopDuringSeep(ByVal hwnd As LongPtr, ByVal message As Long, ByVal id As LongPtr, ByVal tick As Long)
    On Error Resume Next
    If InStr(CStr(Application.StatusBar), "浸透解析：反復計算中") > 0 Then
        stopStatus = CStr(Application.StatusBar): stopIteration = SIterations
        Call RPX_Stop: stopCount = stopCount + 1
        Call KillTimer(0, timerId): timerId = 0
    End If
End Sub
Private Sub Configure(ByVal mode As Long)
    Dim definition As Worksheet, material As Worksheet, ws As Worksheet, i As Long
    Dim coords(1 To 4, 1 To 3) As Variant
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
    ws.Range("E7:K7").Value2 = Array(1#, 1#, 4#, 1#, 0#, "水頭", 0.8)
    ws.Range("E8:K8").Value2 = Array(2#, 1#, 2#, 3#, 0#, "水頭", 0.2)
    ws.Range("M7:N7").Value2 = Array(0#, 2#): ws.Range("M8:N8").Value2 = Array(1#, 2#)
    ws.Range("C209").Value2 = 100#: ws.Range("C211").Value2 = 0#
    RPX_InternalWrite = False
End Sub
Public Function ActiveStopNegative() As Variant
    Dim ws As Worksheet, result As String, n As Long, text As String
    Configure 0: Set ws = RPX_Sheet("浸透入力")
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
    ws.Range("D2").Resize(SN, 1).Value2 = 20#
    Call RPX_InputBegin: Call RPX_ReadInputs: Call RPX_GenerateMesh
    after = RBody0(0, 1): Call RPX_InputEnd
    ResultMutationNegative = Array(result, before, after, RPX_Sheet("浸透結果").Range("K2").Value2, RPX_WaterApplied)
End Function
