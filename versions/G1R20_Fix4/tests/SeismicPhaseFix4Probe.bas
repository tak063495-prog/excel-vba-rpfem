Attribute VB_Name = "SeismicPhaseFix4Probe"
Option Explicit
Private fixtureReady As Boolean

Private Sub PutSetting(ByVal key As String, ByVal value As Variant)
    Dim ws As Worksheet, row As Long
    Set ws = RPX_Sheet("設定")
    For row = 4 To 100
        If CStr(ws.Cells(row, 4).Value2) = key Then ws.Cells(row, 2).Value2 = value: Exit Sub
    Next row
    Err.Raise 5, "SeismicPhaseFix4Probe", "MISSING_SETTING " & key
End Sub

Private Sub WorksheetSquare(ByVal gravity As String, ByVal kh As Double)
    Dim ws As Worksheet
    If Not fixtureReady Then
        Call RPX_Setup: Call RPX_SeismicSetup
        fixtureReady = True
    End If
    Set ws = RPX_Sheet("要素定義")
    ws.Range("A7:C1007,E7:H1007,J7:T1007").ClearContents
    ws.Range("A7:C7").Value2 = Array(1, 0#, 0#)
    ws.Range("A8:C8").Value2 = Array(2, 0#, 1#)
    ws.Range("A9:C9").Value2 = Array(3, 1#, 1#)
    ws.Range("A10:C10").Value2 = Array(4, 1#, 0#)
    ws.Range("E7:H7").Value2 = Array(1, 1, 0, "1,2,3,4")
    ws.Range("J7:T7").Value2 = Array(1, 1, 1, 2, 0, "固定X", 0, 0#, 0#, 0#, 0#)
    ws.Range("J8:T8").Value2 = Array(2, 1, 4, 1, 1, "固定Y", 0, 0#, 0#, 0#, 0#)
    ws.Range("J9:T9").Value2 = Array(3, 1, 2, 3, 0, "自由", 0, 0#, 0#, 0#, -1#)
    Set ws = RPX_Sheet("材料データ")
    ws.Range("A5:E1005,F5:Q1005").ClearContents
    ws.Range("A5:D5").Value2 = Array(1, 100#, 0#, 20#)
    PutSetting "MODE", "極限支持力": PutSetting "TARGET", 16
    PutSetting "ADAPT", 0: PutSetting "QUADRATURE", 2
    PutSetting "GRAVITY", gravity: PutSetting "BEARING_FACES", "3"
    PutSetting "SEISMIC_KH", kh: PutSetting "SEISMIC_DIR", "右(+x)"
    PutSetting "C0_AUTO", 0
    Set ws = RPX_Sheet("浸透入力")
    If CStr(ws.Range("Z1").Value2) = "RPX_SEEP_LAYOUT_V2" Then ws.Range("D4").Value2 = 0 Else ws.Range("N19").Value2 = 0
    RPX_WaterSuspend = False
End Sub

Public Sub PhaseSeedInputFixture()
    WorksheetSquare "固定", 0.2
    PutSetting "MODE", "Fs"
End Sub

Public Function PhasePublicMesh(ByVal gravity As String, ByVal kh As Double, Optional ByVal solve As Boolean = False, Optional ByVal q As Long = 2) As Variant
    Dim e As Long, b0x As Double, b0y As Double, b1x As Double, b1y As Double, previous As String, physical As String
    Call RPX_InputEnd: XCancel = False
    previous = RPX_LoadPhase
    WorksheetSquare gravity, kh
    PutSetting "QUADRATURE", q
    physical = RPX_CurrentInputContext()
    RPX_CreateMesh
    If Len(RPX_LastError) > 0 Then Err.Raise 5, "SeismicPhaseFix4Probe", RPX_LastError
    For e = 0 To re - 1
        b0x = b0x + RArea(e) * RBody0(e, 0): b0y = b0y + RArea(e) * RBody0(e, 1)
        b1x = b1x + RArea(e) * RBody1(e, 0): b1y = b1y + RArea(e) * RBody1(e, 1)
    Next e
    If solve Then
        RPX_Run
        If Len(RPX_LastError) > 0 Then Err.Raise 5, "SeismicPhaseFix4Probe", RPX_LastError
    End If
    PhasePublicMesh = Array("PASS", previous, gravity, kh, rn, re, RPX_LoadPhase, b0x, b0y, b1x, b1y, physical, RLowerLoad, RUpperLoad, RSubdivisions)
End Function

Public Function PhaseRealFs(ByVal path As String, ByVal output As String) As Variant
    Dim answer As String
    Call RPX_InputEnd: XCancel = False: RRawAuditEnabled = True
    answer = RPX_TestFs(path, output)
    If Left$(answer, 4) <> "PASS" Then Err.Raise 5, "SeismicPhaseFix4Probe", answer
    PhaseRealFs = Array(answer, RPX_LoadPhase, RLowerFs, RUpperFs, RFsCalls, RFsSearchIncomplete)
End Function

Public Function PhasePublicRun(ByVal mode As String, Optional ByVal q As Long = 2) As Variant
    Dim e As Long, b0x As Double, b0y As Double, b1x As Double, b1y As Double, previous As String, physical As String
    Call RPX_InputEnd: XCancel = False
    previous = RPX_LoadPhase
    WorksheetSquare "固定", 0.2
    PutSetting "MODE", mode: PutSetting "QUADRATURE", q
    physical = RPX_CurrentInputContext()
    RPX_Run
    For e = 0 To re - 1
        b0x = b0x + RArea(e) * RBody0(e, 0): b0y = b0y + RArea(e) * RBody0(e, 1)
        b1x = b1x + RArea(e) * RBody1(e, 0): b1y = b1y + RArea(e) * RBody1(e, 1)
    Next e
    PhasePublicRun = Array("RETURNED", previous, mode, RPX_LastError, rn, re, RPX_LoadPhase, b0x, b0y, b1x, b1y, physical, RLowerLoad, RUpperLoad, RLowerFs, RUpperFs, RPX_ResultCurrent, RFsSearchIncomplete, RPX_LastFailureNumber)
End Function
