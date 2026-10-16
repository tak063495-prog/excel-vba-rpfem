Attribute VB_Name = "MeshOnlyProbe"
Option Explicit
Public Sub MeshOnlySetting(ByVal key As String, ByVal value As Variant)
    Dim ws As Worksheet, r As Long
    Set ws = RPX_Sheet("設定")
    For r = 4 To 1000
        If CStr(ws.Cells(r, 4).Value2) = key Then ws.Cells(r, 2).Value2 = value: Exit Sub
    Next r
    For r = 4 To 1000
        If Len(CStr(ws.Cells(r, 4).Value2)) = 0 And Len(CStr(ws.Cells(r, 2).Value2)) = 0 Then
            ws.Cells(r, 4).Value2 = key: ws.Cells(r, 2).Value2 = value: Exit Sub
        End If
    Next r
    Err.Raise 5, , "MESH_ONLY_MISSING_SETTING_" & key
End Sub
Public Function MeshOnlyPresetUnit() As Variant
    Dim calls As Long
    On Error GoTo Failed
    calls = RFsCalls
    RPX_Setup
    If CDbl(RPX_Setting("ADAPT_PRECISION_FROM_START", -1)) <> 0# Then Err.Raise 5, , "ORDINARY_DEFAULT_CHANGED"
    RPX_Speed2pctPreset
    If CDbl(RPX_Setting("GAP", -1)) <> 0.02 Then Err.Raise 5, , "PRESET_GAP"
    If CLng(RPX_Setting("QUADRATURE", -1)) <> 8 Then Err.Raise 5, , "PRESET_Q"
    If CLng(RPX_Setting("DIRECT_DLL", -1)) <> 1 Then Err.Raise 5, , "PRESET_DIRECT"
    If CLng(RPX_Setting("ADAPT_C0_WEDGE_AUTO", -1)) <> 1 Then Err.Raise 5, , "PRESET_AUTO"
    If CLng(RPX_Setting("ADAPT_PRECISION_FROM_START", -1)) <> 1 Then Err.Raise 5, , "PRESET_PRECISION"
    If CDbl(RPX_Setting("FS_TOL", -1)) <> 0.0001 Then Err.Raise 5, , "PRESET_ROOT_TOL"
    If CLng(RPX_Setting("ACCURACY_DLL_RECOVERY", -1)) <> 1 Then Err.Raise 5, , "PRESET_RECOVERY"
    If CLng(RPX_Setting("LOWER_MODEL_CACHE", -1)) <> 0 Then Err.Raise 5, , "PRESET_CACHE"
    If CLng(RPX_Setting("TARGET", -1)) <> 1000 Then Err.Raise 5, , "PRESET_TARGET"
    If calls <> RFsCalls Then Err.Raise 5, , "PRESET_SOLVER_CALLED"
    MeshOnlyPresetUnit = Array("PASS", "ordinary_default_0", "q8_direct1_auto1_gap.02_precision1", "original_root_tolerance", "solver_calls_0")
    Exit Function
Failed:
    MeshOnlyPresetUnit = Array("FAIL", Err.Number, Err.Description)
End Function
' Geometry only. This format contains no Fs, field, solver status or root proof.
Public Function MeshOnlyExport(ByVal target As Long, ByVal path As String, ByVal context As String) As Variant
    Dim f As Integer, i As Long, j As Long, n As Long, v As Double, k As Variant, rec As Variant, ends As Variant, kind As Long
    Dim before As Long, area As Double, number As Long, source As String, message As String
    On Error GoTo Failed
    If target < 16 Or target > 50000 Or Len(Dir$(path)) > 0 Then Err.Raise 5, , "MESH_ONLY_ARGUMENT"
    before = RFsCalls
    RPX_InputBegin: RPX_ReadInputs
    DTarget = target: RPX_GenerateMesh: RPX_TotalReferenceLoads
    RPX_AssertInputs "mesh_only_export"
    If RR <> 0 Then Err.Raise 5, , "MESH_ONLY_RIGID_UNSUPPORTED"
    f = FreeFile: Open path For Binary Access Write As #f
    RPX_ExchangeMagic f, "RP25G001"
    Put #f, , rn: Put #f, , re: Put #f, , RM: Put #f, , RR: Put #f, , RSubdivisions: Put #f, , RStressScale
    For i = 0 To rn - 1
        For j = 0 To 1: v = RXY(i, j): Put #f, , v: Next j
    Next i
    For i = 0 To re - 1
        For j = 0 To 2: n = RTri(i, j): Put #f, , n: Next j
        n = RMatId(i): Put #f, , n: n = RRigidId(i): Put #f, , n
        area = area + RArea(i)
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
    For Each k In RBoundary.Keys
        ends = Split(CStr(k), ":"): n = CLng(ends(0)): Put #f, , n: n = CLng(ends(1)): Put #f, , n
        rec = RBoundary(k)
        Select Case rec(0)
            Case "load": kind = 0
            Case "fixed": kind = 1
            Case "roller_x": kind = 2
            Case "roller_y": kind = 3
            Case Else: Err.Raise 5, , "MESH_ONLY_BOUNDARY"
        End Select
        Put #f, , kind: n = rec(1): Put #f, , n
        For j = 2 To 5: v = rec(j): Put #f, , v: Next j
    Next k
    Close #f
    f = FreeFile: Open context For Output As #f
    Print #f, "schema=RP25G001;geometry_only=True;field_present=False;root_present=False;solver_called=False"
    Print #f, "target=" & target & ";nodes=" & rn & ";elements=" & re & ";q=" & RSubdivisions
    Print #f, "area_exact=" & RPX_ExactDouble(area)
    Print #f, "policy=" & RPX_AnalysisPolicy() & ";phase=" & RPX_LoadPhase & ";epoch=" & RPX_InputEpoch
    Print #f, "model=" & RPX_ExactModel()
    Close #f
    If RFsCalls <> before Then Err.Raise 5, , "MESH_ONLY_UNEXPECTED_FS_CALL"
    MeshOnlyExport = Array("PASS", target, rn, re, RM, area, RSubdivisions, RPX_AnalysisPolicy(), RPX_LoadPhase, 0)
CleanUp:
    RPX_InputEnd
    Exit Function
Failed:
    number = Err.Number: source = Err.Source: message = Err.Description
    On Error Resume Next: Close #f: On Error GoTo 0
    MeshOnlyExport = Array("FAIL", number, source, message)
    Resume CleanUp
End Function
