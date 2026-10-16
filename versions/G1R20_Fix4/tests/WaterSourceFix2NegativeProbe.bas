Attribute VB_Name = "WaterSourceFix2NegativeProbe"
Option Explicit
Private Sub Require(ByVal accepted As Boolean, ByVal reason As String)
    If Not accepted Then Err.Raise 5, "WaterSourceFix2NegativeProbe", reason
End Sub
Public Function WaterSourceStaleNegative(ByVal mode As Long) As Variant
    Dim definition As Worksheet, material As Worksheet, ws As Worksheet, message As String, i As Long
    Dim coords(1 To 4, 1 To 3) As Variant, beforeHead As Double, afterHead As Double
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
    definition.Range("J7:P7").Value2 = Array(1#, 1#, 1#, 2#, 0#, "固定XY", 0#)
    definition.Range("Q7:T7").Value2 = 0#
    material.Range("A5:Q1004").ClearContents: material.Range("A5:D5").Value2 = Array(1#, 10#, 30#, 20#)
    RPX_Sheet("設定").Range("B5").Value2 = 32#
    ws.Range("B3").Value2 = ws.Cells(4 + mode, 26).Value2
    ws.Range("B4").Value2 = 9.81: ws.Range("D4").Value2 = 1#: ws.Range("F4").Value2 = 0#
    ws.Range("A7:C206,E7:K206,M7:N206,P7:Q206").ClearContents
    ws.Range("A7:C7").Value2 = Array(1#, 1#, 1#)
    ws.Range("E7:K7").Value2 = Array(1#, 1#, 4#, 1#, 0#, "水頭", 0.8)
    ws.Range("E8:K8").Value2 = Array(2#, 1#, 2#, 3#, 0#, "水頭", 0.2)
    ws.Range("M7:N7").Value2 = Array(0#, 0.8): ws.Range("M8:N8").Value2 = Array(1#, 0.2)
    ws.Range("C209").Value2 = 32#: ws.Range("C211").Value2 = 0#: ws.Range("C216").Value2 = 12#
    If mode = 3 Then
        ws.Range("K7").Value2 = "?0.8"
        For i = 0 To 2
            ws.Cells(7 + i, 16).Value2 = 0.25 * (i + 1)
            ws.Cells(7 + i, 17).Value2 = 0.8 - 0.6 * 0.25 * (i + 1)
        Next i
    End If
    RPX_InternalWrite = False: message = RPX_SeepRunSilent()
    Require Len(message) = 0 Or (mode = 3 And InStr(message, "推定結果") > 0), "PUBLIC_PRODUCER " & message
    Require RPX_Sheet("浸透結果").Range("K2").Value2 = True, "PUBLIC_UNCONVERGED"
    Call RPX_ReadInputs: Call RPX_GenerateMesh
    Require RPX_WaterApplied, "PUBLIC_SOURCE_NOT_ACCEPTED"
    beforeHead = RPX_WaterHeadAt(0.5, 0.5)
    Select Case mode
        Case 0: ws.Range("K7").Value2 = 3#
        Case 1: ws.Range("N8").Value2 = 3#
        Case 2: ws.Range("N7:N8").Value2 = 3#
        Case 3: ws.Range("Q7").Value2 = 3#
    End Select
    Call RPX_CreateMesh
    Require Len(RPX_LastError) = 0, "CREATE_MESH_FAILED " & RPX_LastError
    afterHead = RPX_WaterHeadAt(0.5, 0.5)
    WaterSourceStaleNegative = Array("STALE_ACCEPTED", mode, beforeHead, afterHead, RPX_WaterApplied, RPX_Sheet("浸透結果").Range("K2").Value2, RPX_Sheet("浸透結果").Range("K9").Value2)
End Function
