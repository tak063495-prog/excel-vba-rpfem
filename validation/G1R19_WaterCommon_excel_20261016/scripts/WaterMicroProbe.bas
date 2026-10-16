Attribute VB_Name = "WaterMicroProbe"
Option Explicit
Public Function WaterMicroPrepareResult() As String
    Dim ws As Worksheet, message As String
    On Error GoTo Failed
    Call RPX_SeepSetupSheet
    Set ws = RPX_Sheet("浸透入力")
    ws.Range("B3").Value2 = ws.Range("Z6").Value2
    ws.Range("B4").Value2 = 9.81: ws.Range("D4").Value2 = 1#: ws.Range("F4").Value2 = 0#
    ws.Range("M7:N206").ClearContents
    ws.Range("M7").Value2 = 0#: ws.Range("N7").Value2 = 3#
    ws.Range("M8").Value2 = 20#: ws.Range("N8").Value2 = 3#
    message = RPX_SeepRunSilent()
    If Len(message) > 0 Then Err.Raise 5, , "PRESCRIBED_RESULT_FAILED " & message
    If Not CBool(RPX_Sheet("浸透結果").Range("K2").Value2) Or RPX_SeepMode(ws) <> 2 Then Err.Raise 5, , "PRESCRIBED_RESULT_NOT_READY"
    WaterMicroPrepareResult = "PASS fresh prescribed mode2; no seep-FEM/Fs solver"
    Exit Function
Failed:
    WaterMicroPrepareResult = "FAIL " & Err.Number & " " & Err.Description
End Function
Public Function WaterMicroRollbackSeed() As String
    On Error GoTo Failed
    Call RPX_InputEnd: Call RPX_TestDisarm: XCancel = False: RPX_WaterSuspend = False
    RPX_Sheet("浸透入力").Range("D4").Value2 = 0#
    Call RPX_InputBegin: Call RPX_ReadInputs: DTarget = 64: Call RPX_GenerateMesh: Call RPX_InputEnd
    RPX_Sheet("浸透入力").Range("D4").Value2 = 1#
    WaterMicroRollbackSeed = "PASS"
    Exit Function
Failed:
    WaterMicroRollbackSeed = "FAIL " & Err.Number & " " & Err.Description
    On Error Resume Next: Call RPX_InputEnd
End Function
Public Function WaterMicroGenerate(ByVal directory As String) As Variant
    Dim i As Long, calls As Long, eta() As Double, points As Long, cells As Long, faces As Long
    Dim level As Double, waterGamma As Double
    On Error GoTo Failed
    calls = RFsCalls: faces = DFCount
    RPX_WaterSuspend = True: Call RPX_InputBegin: Call RPX_ReadInputs
    DTarget = 350: Call RPX_GenerateMesh
    WriteMesh directory & "\base_raw.bin"
    Call RPX_InputEnd: RPX_WaterSuspend = False
    Call RPX_InputBegin: Call RPX_ReadInputs
    Call RPX_AssignRegions
    If Not RPX_WaterCommonLevel(level, waterGamma) Or level <> 3# Or waterGamma <> 9.81 Then Err.Raise 5, , "COMMON_ELIGIBILITY"
    If Not RPX_WaterCommonApplied Then Err.Raise 5, , "COMMON_LOAD_NOT_ENGAGED"
    CheckCommon faces
    WriteMesh directory & "\common_raw.bin"
    Call RPX_TotalReferenceLoads: CheckCommon faces: WriteMesh directory & "\common_reference.bin"
    Call RPX_WaterApply: CheckCommon faces: WriteMesh directory & "\common_idempotent.bin"
    ReDim eta(0 To re - 1)
    For i = 0 To re - 1: eta(i) = 1#: Next i
    RPX_RefineMesh eta, 1#, True
    CheckCommon faces: WriteMesh directory & "\uniform_raw.bin"
    Call RPX_TotalReferenceLoads: CheckCommon faces: WriteMesh directory & "\uniform_reference.bin"
    ReDim eta(0 To re - 1)
    For i = 0 To re - 1: eta(i) = (i Mod 3): Next i
    RPX_RefineMesh eta, 0.15, True
    CheckCommon faces: WriteMesh directory & "\local_raw.bin"
    Call RPX_TotalReferenceLoads: CheckCommon faces: WriteMesh directory & "\local_reference.bin"
    points = rn: cells = re
    If calls <> RFsCalls Then Err.Raise 5, , "MICRO_UNEXPECTED_FS_SOLVE"
    WaterMicroGenerate = Array("PASS", points, cells, "raw/reference/uniform/local", "source_guard_bound", "Fs_calls_0")
CleanUp:
    Call RPX_InputEnd: RPX_WaterSuspend = False
    Exit Function
Failed:
    WaterMicroGenerate = Array("FAIL", Err.Number, Err.Description)
    Resume CleanUp
End Function
Private Sub CheckCommon(ByVal faces As Long)
    Dim e As Long, k As Long, below As Boolean, value As Double, id As Long, key As String
    If Not RPX_WaterCommonConforming(3#) Or Not RPX_WaterCommonApplied Or DFCount <> faces Then Err.Raise 5, , "MICRO_COMMON_CONTEXT"
    For e = 0 To re - 1
        below = True
        For k = 0 To 2: If RXY(RTri(e, k), 1) > 3# Then below = False
        Next k
        value = -RMaterials(RMatId(e)).gamma
        If below Then value = value + 9.81
        If Abs(RBody0(e, 0) + RBody1(e, 0)) > 0.000000000001 Or Abs(RBody0(e, 1) + RBody1(e, 1) - value) > 0.000000000001 Then Err.Raise 5, , "MICRO_BODY_DIFFERENT"
    Next e
    For id = 0 To RNEdge - 1
        With REdges(id)
            If .element(1) >= 0 And RXY(.a, 1) = 3# And RXY(.b, 1) = 3# Then
                key = RPX_EdgeKey(.a, .b)
                If Not RConstraints.Exists(key) Or RBoundary.Exists(key) Then Err.Raise 5, , "WATERLINE_PHYSICAL_BOUNDARY"
                If .kind <> "load" Or .rigidId <> -1 Then Err.Raise 5, , "WATERLINE_PHYSICAL_RESTRAINT"
                For k = 0 To 1
                    If .load0(k) <> 0# Or .load1(k) <> 0# Then Err.Raise 5, , "WATERLINE_EXTRA_FORCE"
                Next k
            End If
        End With
    Next id
End Sub
Private Sub WriteMesh(ByVal path As String)
    Dim f As Integer, i As Long, j As Long, n As Long, v As Double, key As Variant, rec As Variant, ends As Variant, kind As Long
    If Len(Dir$(path)) > 0 Then Err.Raise 5, , "MICRO_OUTPUT_EXISTS"
    f = FreeFile: Open path For Binary Access Write As #f
    RPX_ExchangeMagic f, "RP25G001"
    Put #f, , rn: Put #f, , re: Put #f, , RM: Put #f, , RR: Put #f, , RSubdivisions: Put #f, , RStressScale
    For i = 0 To rn - 1
        For j = 0 To 1: v = RXY(i, j): Put #f, , v: Next j
    Next i
    For i = 0 To re - 1
        For j = 0 To 2: n = RTri(i, j): Put #f, , n: Next j
        n = RMatId(i): Put #f, , n: n = RRigidId(i): Put #f, , n
    Next i
    For i = 0 To RM - 1
        With RMaterials(i)
            Put #f, , .cohesion: Put #f, , .friction: Put #f, , .dilation: Put #f, , .gamma
        End With
    Next i
    For i = 0 To re - 1
        For j = 0 To 1: v = RBody0(i, j): Put #f, , v: v = RBody1(i, j): Put #f, , v: Next j
    Next i
    n = RBoundary.Count: Put #f, , n
    For Each key In RBoundary.Keys
        ends = Split(CStr(key), ":"): n = CLng(ends(0)): Put #f, , n: n = CLng(ends(1)): Put #f, , n
        rec = RBoundary(key)
        Select Case rec(0)
            Case "load": kind = 0
            Case "fixed": kind = 1
            Case "roller_x": kind = 2
            Case "roller_y": kind = 3
            Case Else: Err.Raise 5, , "MICRO_BOUNDARY"
        End Select
        Put #f, , kind: n = rec(1): Put #f, , n
        For j = 2 To 5: v = rec(j): Put #f, , v: Next j
    Next key
    Close #f
    f = FreeFile: Open path & ".context.tsv" For Output As #f
    Print #f, "schema=RP25G001;geometry_only=True;solver_called=False"
    Print #f, "phase=" & RPX_LoadPhase & ";policy=" & RPX_AnalysisPolicy() & ";q=" & RSubdivisions
    Print #f, "nodes=" & rn & ";elements=" & re & ";physical_faces=" & DFCount
    Print #f, "common_applied=" & RPX_WaterCommonApplied & ";source=" & RPX_WaterCommonNote
    Print #f, "model=" & RPX_ExactModel()
    For Each key In RConstraints.Keys: Print #f, "geometry_constraint=" & CStr(key): Next key
    Close #f
End Sub
