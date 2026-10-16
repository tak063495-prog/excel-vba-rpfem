Attribute VB_Name = "WaterSourceFix3Probe"
Option Explicit
Private Const SOURCE_NS As String = "urn:rpfem:source:seep:v1"
Private Const TEST_NS As String = "urn:rpfem:test:source:fix3"

Private Sub Require(ByVal accepted As Boolean, ByVal reason As String)
    If Not accepted Then Err.Raise 5, "WaterSourceFix3Probe", reason
End Sub
Private Sub UnitSquareArrays()
    SN = 4: SM = 2: SNMat = 1
    ReDim sxy(0 To 3, 0 To 1): ReDim STri(0 To 1, 0 To 2)
    ReDim SMat(0 To 1): ReDim SKx(0 To 0): ReDim SKy(0 To 0)
    sxy(1, 0) = 1#: sxy(2, 0) = 1#: sxy(2, 1) = 1#: sxy(3, 1) = 1#
    STri(0, 0) = 0: STri(0, 1) = 1: STri(0, 2) = 2
    STri(1, 0) = 0: STri(1, 1) = 2: STri(1, 2) = 3
    SKx(0) = 1#: SKy(0) = 1#: SGammaW = 9.81
    rn = 4: re = 2: RM = 1: RR = 0: RXY = sxy: RTri = STri
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
Private Sub DeleteRecords(ByVal recordNamespace As String)
    Dim parts As Object, i As Long
    Set parts = ThisWorkbook.CustomXMLParts.SelectByNamespace(recordNamespace)
    For i = parts.Count To 1 Step -1: parts.Item(i).Delete: Next i
End Sub
Private Function Rejection(ByVal expected As String, Optional ByVal assertion As Boolean = False) As Boolean
    Dim number As Long, message As String
    On Error GoTo Rejected
    If assertion Then Call RPX_WaterCommonInputAssert Else Call RPX_WaterApply
    Exit Function
Rejected:
    number = Err.Number: message = Err.Description
    If number <> 5 Or InStr(message, expected) = 0 Then Err.Raise 5, "WaterSourceFix3Probe", "UNEXPECTED_ERROR " & number & " " & message
    Rejection = True
End Function
Public Function WaterSourcePrepare(ByVal mode As Long) As Variant
    Dim ws As Worksheet, row As Long
    RPX_InternalWrite = True: XCancel = False: RPX_WaterSuspend = False
    Call RPX_InputEnd: Call RPX_SeepSourceEnd: Call RPX_SeepSetupSheet
    Set ws = RPX_Sheet("浸透入力")
    ws.Range("B3").Value2 = ws.Cells(4 + mode, 26).Value2
    ws.Range("B4").Value2 = 9.81: ws.Range("D4").Value2 = 1#: ws.Range("F4").Value2 = 0#
    ws.Range("A7:C206,E7:K206,M7:N206,P7:Q206").ClearContents
    ws.Range("A7:C7").Value2 = Array(1#, 1#, 1#)
    ws.Range("E7:K7").Value2 = Array(1#, 1#, 1#, 2#, 0#, "水頭", 2#)
    ws.Range("M7").Value2 = 0#: ws.Range("M8").Value2 = 1#: ws.Range("N7:N8").Value2 = 2#
    ws.Range("P7").Value2 = 0.5: ws.Range("Q7").Value2 = 2#: ws.Range("C209").Value2 = 32#
    For row = 4 To 100
        If CStr(RPX_Sheet("設定").Cells(row, 4).Value2) = "SEISMIC_KH" Then RPX_Sheet("設定").Cells(row, 2).Value2 = 0#
    Next row
    Call UnitSquareArrays
    Call RPX_WTReadTable(ws, True, 13): Call RPX_WTField(2)
    ' A direct public writer fixture verifies lifecycle/source rules only.
    ' It does not claim a solved mode0/3 numerical model.
    Call RPX_SeepWriteResults: Call RPX_WTWriteMode(mode): Call RPX_WaterApply
    Require RPX_WaterHeadAt(0.25, 0.1) = 2#, "BASELINE_HEAD"
    RPX_InternalWrite = False
    WaterSourcePrepare = Array("PASS", mode, RPX_Sheet("浸透結果").Range("K9").Value2)
End Function
Public Function WaterSourceEdit(ByVal mode As Long) As Variant
    Dim ws As Worksheet
    Set ws = RPX_Sheet("浸透入力")
    Select Case mode
        Case 0: ws.Range("K7").Value2 = 3#
        Case 1: ws.Range("N8").Value2 = 3#
        Case 2: ws.Range("N7:N8").Value2 = 3#
        Case 3: ws.Range("Q7").Value2 = 3#
    End Select
    WaterSourceEdit = Array("EDITED", mode)
End Function
Public Function WaterSourceObserve() As Variant
    Dim number As Long, message As String
    Call UnitSquareArrays: XCancel = False: RPX_WaterSuspend = False
    On Error GoTo Rejected
    Call RPX_WaterApply
    WaterSourceObserve = Array("ACCEPTED", RPX_WaterHeadAt(0.25, 0.1), RBody0(0, 1))
    Exit Function
Rejected:
    number = Err.Number: message = Err.Description
    WaterSourceObserve = Array("REJECTED", number, message)
End Function
Public Function WaterSourceRecordRoundTrip() As Variant
    Dim context As String, savedId As String, savedContext As String, parts As Object, duplicate As Object, number As Long
    Dim r As Long
    context = "日本語" & vbNullChar & ChrW(&HD83D) & ChrW(&HDE00) & vbCrLf & RPX_SourceCell(CVErr(2042))
    For r = 1 To 12: context = context & context: Next r
    Call RPX_SourceRecordWrite(TEST_NS, "長いID", context)
    Require RPX_SourceRecordRead(TEST_NS, savedId, savedContext), "RECORD_MISSING"
    Require savedId = "長いID" And savedContext = context, "UTF16_NOT_LOSSLESS"
    Call RPX_SourceRecordWrite(TEST_NS, "replacement", context & "2")
    Require RPX_SourceRecordRead(TEST_NS, savedId, savedContext), "REPLACEMENT_MISSING"
    Require savedId = "replacement" And savedContext = context & "2", "REPLACEMENT_WRONG"
    Set parts = ThisWorkbook.CustomXMLParts.SelectByNamespace(TEST_NS)
    Require parts.Count = 1, "OLD_RECORD_RETAINED"
    Set duplicate = ThisWorkbook.CustomXMLParts.Add(parts.Item(1).XML)
    On Error Resume Next
    Call RPX_SourceRecordRead(TEST_NS, savedId, savedContext)
    number = Err.Number: Err.Clear
    On Error GoTo 0
    Require number = 5, "DUPLICATE_ACCEPTED"
    Call DeleteRecords(TEST_NS)
    Require Not RPX_SourceRecordRead(TEST_NS, savedId, savedContext), "DELETED_RECORD_PRESENT"
    WaterSourceRecordRoundTrip = Array("PASS", Len(context), "UTF16_NUL_SURROGATE_ERROR_DUPLICATE_REPLACEMENT")
End Function
Public Function WaterSourceAllModes() As Variant
    Dim mode As Long, ws As Worksheet, unused As Variant, oldFormula As Variant, number As Long
    Dim accepted As Long, changed As Long, scoped As Long, missing As Long, switched As Long, ignored As Long
    Dim previousEvents As Boolean
    previousEvents = Application.EnableEvents: Application.EnableEvents = False
    On Error GoTo Failed
    For mode = 0 To 3
        unused = WaterSourcePrepare(mode): accepted = accepted + 1
        unused = WaterSourceEdit(mode)
        Require Rejection("SEEP_SOURCE_CHANGED"), "STALE_MODE_" & mode: changed = changed + 1
        unused = WaterSourcePrepare(mode): Set ws = RPX_Sheet("浸透入力")
        ws.Range("B3").Value2 = ws.Cells(4 + ((mode + 1) Mod 4), 26).Value2
        Require Rejection("SEEP_SOURCE_CHANGED"), "MODE_SWITCH_" & mode: switched = switched + 1
        unused = WaterSourcePrepare(mode): Call RPX_InputBegin
        unused = WaterSourceEdit(mode)
        Require Rejection("WATER_COMMON_SOURCE_CHANGED", True), "SCOPE_MODE_" & mode: scoped = scoped + 1
        Call RPX_InputEnd
        unused = WaterSourcePrepare(mode): Call DeleteRecords(SOURCE_NS)
        Require Rejection("SEEP_SOURCE_PROOF_MISSING"), "LEGACY_MODE_" & mode: missing = missing + 1
        unused = WaterSourcePrepare(mode)
        oldFormula = RPX_Sheet("材料データ").Range("B5").Formula
        RPX_Sheet("材料データ").Range("B5").Value2 = 23#
        Call RPX_WaterApply: ignored = ignored + 1
        RPX_Sheet("材料データ").Range("B5").Formula = oldFormula
        ws.Range("O7").Value2 = "表示確認だけ"
        ws.Range("C209").Value2 = 999#: Call RPX_WaterApply
    Next mode
    RPX_InternalWrite = False: Application.EnableEvents = previousEvents
    WaterSourceAllModes = Array("PASS", accepted, changed, switched, scoped, missing, ignored, "solver0_lifecycle_only")
    Exit Function
Failed:
    number = Err.Number
    Dim message As String: message = Err.Description
    On Error Resume Next
    Call RPX_InputEnd: Call RPX_SeepSourceEnd
    RPX_InternalWrite = False: Application.EnableEvents = previousEvents
    On Error GoTo 0
    Err.Raise number, "WaterSourceFix3Probe", message
End Function
Public Function WaterSourceProducerChange() As Variant
    Dim unused As Variant, ws As Worksheet, number As Long, message As String
    unused = WaterSourcePrepare(2): Set ws = RPX_Sheet("浸透入力")
    Call RPX_SeepSourceBegin
    ws.Range("N7:N8").Value2 = 3#
    On Error Resume Next
    Call RPX_SeepWriteResults
    number = Err.Number: message = Err.Description: Err.Clear
    On Error GoTo 0
    Require number = 5 And InStr(message, "SEEP_SOURCE_CHANGED") > 0, "MIXED_PRODUCER_ACCEPTED"
    Call RPX_SeepSourceEnd
    Require Rejection("SEEP_RESULT_WRITE_INCOMPLETE"), "MIXED_PRODUCER_RECOVERED_OLD_RESULT"
    WaterSourceProducerChange = Array("PASS", "producer_change_rejected_and_pending")
End Function
Public Function WaterSourceMalformedProof() As Variant
    Dim unused As Variant, identifier As String, context As String, parts As Object, duplicate As Object
    unused = WaterSourcePrepare(2)
    Require RPX_SourceRecordRead(SOURCE_NS, identifier, context), "PROOF_NOT_FOUND"
    Call RPX_SourceRecordWrite(SOURCE_NS, "WRONG_ID", context)
    Require Rejection("SEEP_SOURCE_ID_MISMATCH"), "WRONG_PROOF_ID_ACCEPTED"
    unused = WaterSourcePrepare(2)
    Set parts = ThisWorkbook.CustomXMLParts.SelectByNamespace(SOURCE_NS)
    Set duplicate = ThisWorkbook.CustomXMLParts.Add(parts.Item(1).XML)
    Require Rejection("SOURCE_RECORD_MULTIPLE"), "DUPLICATE_PROOF_ACCEPTED"
    unused = WaterSourcePrepare(2)
    Call DeleteRecords(SOURCE_NS)
    Set duplicate = ThisWorkbook.CustomXMLParts.Add("<rpxSource xmlns='" & SOURCE_NS & "' version='0'><identifier/><context/></rpxSource>")
    Require Rejection("SOURCE_RECORD_VERSION"), "UNKNOWN_PROOF_VERSION_ACCEPTED"
    unused = WaterSourcePrepare(2)
    WaterSourceMalformedProof = Array("PASS", "wrong_ID_duplicate_unknown_version_rejected_recoverable")
End Function
Public Function WaterSourceManualFormula() As Variant
    Dim unused As Variant, ws As Worksheet, previousCalculation As Long, number As Long, message As String
    unused = WaterSourcePrepare(2): Set ws = RPX_Sheet("浸透入力")
    previousCalculation = Application.Calculation
    On Error GoTo Failed
    ws.Range("AA50").Value2 = 2#: ws.Range("N7:N8").Formula = "=$AA$50"
    Call RPX_WTReadTable(ws, True, 13): Call RPX_WTField(2)
    Call RPX_SeepWriteResults: Call RPX_WTWriteMode(2)
    Application.Calculation = -4135
    ws.Range("AA50").Value2 = 3#
    Require ws.Range("N7").Value2 = 2#, "FORMULA_WAS_NOT_STALE_FIXTURE"
    Call RPX_InputBegin
    Require Rejection("SEEP_SOURCE_CHANGED"), "MANUAL_CALC_OLD_FORMULA_ACCEPTED"
    Call RPX_InputEnd: Application.Calculation = previousCalculation
    WaterSourceManualFormula = Array("PASS", "manual_water_formula_recalculated_before_consumer")
    Exit Function
Failed:
    number = Err.Number: message = Err.Description
    On Error Resume Next
    Call RPX_InputEnd: Application.Calculation = previousCalculation
    On Error GoTo 0
    Err.Raise number, "WaterSourceFix3Probe", message
End Function
Public Function WaterSourcePublicRun(ByVal mode As Long, Optional ByVal editInput As Boolean = True) As Variant
    Dim definition As Worksheet, material As Worksheet, ws As Worksheet, message As String, i As Long
    Dim coords(1 To 4, 1 To 3) As Variant, unused As Variant
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
    ws.Range("E7:K7").Value2 = Array(1#, 1#, 4#, 1#, 1#, "水頭", 0.8)
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
    If Not editInput Then
        WaterSourcePublicRun = Array("PASS", mode, SN, SM, SIterations, "healthy_public_producer")
        Exit Function
    End If
    unused = WaterSourceEdit(mode)
    Call RPX_CreateMesh
    Require InStr(RPX_LastError, "SEEP_SOURCE_CHANGED") > 0, "PUBLIC_CREATE_MESH_STALE_ACCEPTED " & RPX_LastError
    Require Rejection("SEEP_SOURCE_CHANGED"), "PUBLIC_STALE_ACCEPTED"
    WaterSourcePublicRun = Array("PASS", mode, SN, SM, SIterations, message)
End Function
