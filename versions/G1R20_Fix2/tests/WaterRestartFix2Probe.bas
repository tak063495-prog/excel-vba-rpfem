Attribute VB_Name = "WaterRestartFix2Probe"
Option Explicit
Private Const PASSWORD As String = "water-fix2-restart"

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
Private Function Marker() As String
    Dim item As Object
    For Each item In ThisWorkbook.CustomDocumentProperties
        If CStr(item.Name) = "RPX_SeepResultWriteState" Then Marker = CStr(item.Value): Exit Function
    Next item
    Marker = "ABSENT_LEGACY"
End Function
Private Sub Require(ByVal passed As Boolean, ByVal text As String)
    If Not passed Then Err.Raise 5, "WaterRestartFix2Probe", text
End Sub
Private Sub Table(ByVal ws As Worksheet, ByVal mode As Long, ByVal changed As Boolean)
    ws.Range("M7:N206").ClearContents
    ws.Range("M7").Value2 = 0#: ws.Range("M8").Value2 = 1#
    ws.Range("N7:N8").Value2 = 2#
    If changed Then
        If mode = 1 Then ws.Range("N8").Value2 = 3# Else ws.Range("N7:N8").Value2 = 3#
    End If
End Sub
Private Sub Produce(ByVal ws As Worksheet, ByVal mode As Long)
    Call RPX_WTReadTable(ws, True, 13): Call RPX_WTField(mode)
    Call RPX_SeepWriteResults: Call RPX_WTWriteMode(mode)
End Sub
Private Function ApplyState() As Variant
    Dim number As Long, text As String
    On Error GoTo Rejected
    Call RPX_WaterApply
    ApplyState = Array("ACCEPTED", RPX_WaterHeadAt(0.25, 0.1), RBody0(0, 0), RPX_WaterCommonApplied, Marker())
    Exit Function
Rejected:
    number = Err.Number: text = Err.Description
    ApplyState = Array("REJECTED", number, text, False, Marker())
End Function
Private Function CommonState() As Variant
    Dim level As Double, gamma As Double
    On Error GoTo Rejected
    CommonState = Array("RETURNED", RPX_WaterCommonLevel(level, gamma), level, gamma)
    Exit Function
Rejected:
    CommonState = Array("REJECTED", Err.Number, Err.Description, 0#)
End Function

Public Function WaterRestartPrepare(ByVal mode As Long, ByVal failure As String) As Variant
    Dim ws As Worksheet, result As Worksheet, state As Variant, number As Long, row As Long
    Dim expectedHead As Double, expectedBodyX As Double, k As Long
    RPX_InternalWrite = True: RPX_WaterSuspend = False: XCancel = False
    Call RPX_InputEnd: Call RPX_SeepSetupSheet
    Set ws = RPX_Sheet("浸透入力"): Set result = RPX_Sheet("浸透結果")
    ws.Range("B3").Value2 = ws.Cells(4 + mode, 26).Value2
    ws.Range("B4").Value2 = 9.81: ws.Range("D4").Value2 = 1#: ws.Range("F4").Value2 = 0#
    ws.Range("C209").Value2 = 32
    For row = 4 To 100
        If CStr(RPX_Sheet("設定").Cells(row, 4).Value2) = "SEISMIC_KH" Then RPX_Sheet("設定").Cells(row, 2).Value2 = 0#
    Next row
    Call UnitSquareArrays: Call Table(ws, mode, False): Call Produce(ws, mode)
    Call RPX_WaterApply
    Require Abs(RPX_WaterHeadAt(0.25, 0.1) - 2#) < 0.000000001, "OLD_HEAD_FIXTURE"
    result.Protect PASSWORD
    ThisWorkbook.Save
    If failure = "writer" Then
        Call Table(ws, mode, True)
        Call RPX_WTReadTable(ws, True, 13): Call RPX_WTField(mode)
        On Error Resume Next
        Call RPX_SeepWriteResults
        number = Err.Number: Err.Clear
        On Error GoTo 0
        Require number <> 0, "PROTECTED_WRITER_DID_NOT_FAIL"
    ElseIf failure = "ui" Then
        Call Application.Run("'" & ThisWorkbook.Name & "'!RPX_TestFix1InvalidateSeepResult")
    Else
        Err.Raise 5, , "BAD_FAILURE_KIND"
    End If
    Require VarType(result.Range("K2").Value2) = vbBoolean, "OLD_K2_TYPE"
    Require CBool(result.Range("K2").Value2), "OLD_K2_NOT_TRUE"
    expectedHead = 0.75 * SHead(0) + 0.15 * SHead(1) + 0.1 * SHead(2)
    For k = 0 To 2
        expectedBodyX = expectedBodyX - 9.81 * (SHead(STri(0, k)) - sxy(STri(0, k), 1)) * RGrad(0, k, 0)
    Next k
    state = ApplyState()
    Require state(0) = "REJECTED" And InStr(CStr(state(2)), "SEEP_RESULT_WRITE_INCOMPLETE") > 0, "IN_SESSION_FAILURE_NOT_BLOCKED"
    RPX_InternalWrite = False
    WaterRestartPrepare = Array("PASS", mode, failure, result.Range("D2").Value2, result.Range("K2").Value2, result.Range("K9").Value2, Marker(), ThisWorkbook.Saved, expectedHead, expectedBodyX)
End Function

Public Function WaterRestartObserve() As Variant
    Dim common As Variant, state As Variant
    ' Only receiving geometry is reconstructed. No producer or saved result
    ' payload, eligibility flag, ID, or durable marker is modified here.
    Call UnitSquareArrays: RPX_WaterSuspend = False: XCancel = False
    Call RPX_InputBegin
    common = CommonState(): state = ApplyState()
    Call RPX_InputEnd
    WaterRestartObserve = Array(state(0), state(1), state(2), state(3), state(4), common(0), common(1), common(2))
End Function

Public Function WaterRestartRecover() As Variant
    Dim ws As Worksheet, result As Worksheet, message As String, mode As Long, previousId As String
    Set ws = RPX_Sheet("浸透入力"): Set result = RPX_Sheet("浸透結果")
    previousId = CStr(result.Range("K9").Value2)
    result.Unprotect PASSWORD: mode = RPX_SeepMode(ws)
    Call Table(ws, mode, True)
    ' Exercise the public UI producer, including its suspended mesh generation.
    message = RPX_SeepRunSilent()
    Require Len(message) = 0, "PUBLIC_REANALYSIS_FAILED " & message
    Require CBool(result.Range("K2").Value2), "REANALYSIS_NOT_CONVERGED"
    Require CStr(result.Range("K9").Value2) <> previousId, "REANALYSIS_DID_NOT_REPLACE_OLD_RESULT"
    ' A fresh result is accepted, without inventing a root or solver witness.
    Call RPX_ReadInputs: Call RPX_GenerateMesh: Call RPX_WaterApply
    Require RPX_WaterApplied, "REANALYSIS_NOT_ACCEPTED"
    WaterRestartRecover = Array("PASS", mode, result.Range("D2").Value2, re, result.Range("K2").Value2, result.Range("K9").Value2, Marker())
End Function

Public Function WaterRestartLegacy() As Variant
    Dim ws As Worksheet
    RPX_InternalWrite = True: Call RPX_InputEnd: Call RPX_SeepSetupSheet
    Set ws = RPX_Sheet("浸透入力")
    ws.Range("B3").Value2 = ws.Range("Z5").Value2
    ws.Range("B4").Value2 = 9.81: ws.Range("D4").Value2 = 1#: ws.Range("F4").Value2 = 0#
    Call UnitSquareArrays: Call Table(ws, 1, False): Call Produce(ws, 1)
    RPX_InternalWrite = False
    WaterRestartLegacy = Array("PASS", Marker(), RPX_Sheet("浸透結果").Range("K9").Value2)
End Function

Public Function WaterRestartMarkerChecks() As Variant
    Dim item As Object, original As String, state As Variant, ws As Worksheet, level As Double, gamma As Double, message As String
    Set ws = RPX_Sheet("浸透入力")
    original = Marker()
    Require Left$(original, 10) = "COMMITTED|", "NO_HEALTHY_MARKER_FIXTURE"
    Set item = ThisWorkbook.CustomDocumentProperties("RPX_SeepResultWriteState")
    item.Value = "COMMITTED|DIFFERENT_ID"
    state = ApplyState()
    Require state(0) = "REJECTED" And InStr(CStr(state(2)), "SEEP_RESULT_WRITE_ID_MISMATCH") > 0, "WRONG_ID_ACCEPTED"
    item.Value = ""
    state = ApplyState()
    Require state(0) = "REJECTED" And InStr(CStr(state(2)), "INVALID_SEEP_RESULT_WRITE_STATE") > 0, "EMPTY_MARKER_ACCEPTED"
    item.Delete
    Set item = ThisWorkbook.CustomDocumentProperties.Add("RPX_SeepResultWriteState", False, 2, True)
    state = ApplyState()
    Require state(0) = "REJECTED" And InStr(CStr(state(2)), "INVALID_SEEP_RESULT_WRITE_STATE") > 0, "WRONG_TYPE_ACCEPTED"
    message = RPX_SeepRunSilent()
    Require Len(message) > 0 And InStr(message, "INVALID_SEEP_RESULT_WRITE_STATE") > 0, "WRONG_TYPE_PUBLIC_RUN_NOT_REJECTED"
    Require Not RPX_Busy And Not RPX_WaterSuspend And Not RPX_InternalWrite, "WRONG_TYPE_PUBLIC_RUN_LEAKED_STATE"
    item.Delete
    Set item = ThisWorkbook.CustomDocumentProperties.Add("RPX_SeepResultWriteState", False, 4, "PENDING")
    ws.Range("D4").Value2 = 0#
    Require Not RPX_WaterCommonLevel(level, gamma), "DRY_COMMON_ELIGIBLE"
    Call RPX_ReadInputs: DTarget = 32: Call RPX_GenerateMesh: Call RPX_WaterApply
    Require Not RPX_WaterApplied, "DRY_WATER_APPLIED"
    ws.Range("D4").Value2 = 1#: RPX_WaterSuspend = True
    Require Not RPX_WaterCommonLevel(level, gamma), "SUSPENDED_COMMON_ELIGIBLE"
    Call RPX_ReadInputs: DTarget = 32: Call RPX_GenerateMesh: Call RPX_WaterApply
    Require Not RPX_WaterApplied, "SUSPENDED_WATER_APPLIED"
    RPX_WaterSuspend = False
    message = RPX_SeepRunSilent()
    Require Len(message) = 0, "AFTER_CORRUPTION_PUBLIC_REANALYSIS_FAILED " & message
    Call RPX_ReadInputs: DTarget = 32: Call RPX_GenerateMesh: Call RPX_WaterApply
    Require RPX_WaterApplied, "HEALTHY_MARKER_NOT_RESTORED"
    WaterRestartMarkerChecks = Array("PASS", "wrong_ID", "empty", "wrong_type", "typed_failure_cleanup", "pending_dry_mesh", "pending_suspended_mesh", Marker())
End Function
