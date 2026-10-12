Attribute VB_Name = "HSDRegression"
Option Explicit
Public Function HSDCase(ByVal a As Double, ByVal b As Double, ByVal scalar As Boolean, ByVal prefix As String) As String
    Dim i As Long, n As Long, f As Long, before As String
    On Error GoTo Failed
    Call RPX_InputEnd: Call RPX_AffineEnd: Call RPX_CertBeginRun
    Call RPX_ReadInputs: Call RPX_ResetPreparedState: Call RPX_HLURelease: Call RPX_SchurRelease: Call RPX_LDLRelease
    n = 3: If scalar Then n = 4
    RPX_Begin n: RPX_Cost RPX_Unit(0)
    RPX_Equal RPX_Unit(1), a: RPX_Equal RPX_Unit(2), b
    RPX_Cone3 RPX_Unit(0), RPX_Unit(1), RPX_Unit(2)
    If scalar Then
        RPX_Cost RPX_Unit(3), 0.1
        RPX_Equal RPX_Unit(3), 2#
        RPX_LinearCone RPX_Unit(3)
    End If
    RPX_DiagStart "DLL3_HSD_SOCP"
    RPX_OptimizeHSD False
    If XStatus <> "OPTIMAL" Then err.Raise 513, , "HSD_TEST_STATUS " & XStatus
    f = FreeFile: Open prefix & ".bin" For Binary Access Write As #f
    Put #f, , XObjective: Put #f, , XIteration: Put #f, , XPrimalResidual: Put #f, , XDualResidual: Put #f, , XGap
    For i = 0 To XN - 1: Put #f, , xx(i): Next i
    For i = 0 To XS - 1: Put #f, , XSlack(i): Put #f, , XDual(i): Next i
    Close #f
    HSDCase = "PASS;objective=" & XObjective & ";iteration=" & XIteration & ";primal=" & XPrimalResidual & ";dual=" & XDualResidual & ";gap=" & XGap & ";" & RPX_NativeInfo()
    RPX_DiagFinish "OPTIMAL", HSDCase
    Exit Function
Failed:
    HSDCase = "ERROR " & err.number & ":" & err.description & ";" & RPX_NativeInfo()
    RPX_DiagFinish "ERROR", HSDCase
End Function
Public Function FixedFem(ByVal side As String, ByVal factor As Double, ByVal prefix As String) As String
    Dim value As Double
    On Error GoTo Failed
    Call RPX_InputEnd: Call RPX_AffineEnd: Call RPX_CertBeginRun
    Call RPX_ReadInputs: Call RPX_GenerateMesh
    RFsActive = True: RPX_TotalReferenceLoads
    Call RPX_ResetPreparedState: Call RPX_LowerCacheReset: RProgramKind = ""
    RPX_DiagStart "DLL3_FIXED_FEM"
    value = RPX_Bound(side, factor)
    If XStatus <> "OPTIMAL" Then err.Raise 513, , "FIXED_FEM_STATUS " & XStatus
    RPX_ExportExchangeX prefix & "_field.bin"
    MeshExport prefix & "_mesh.bin"
    FixedFem = "PASS;side=" & side & ";Fs=" & factor & ";load=" & value & ";iteration=" & XIteration & ";audit=" & RAudit & ";elements=" & re & ";engine=" & RPX_RobustEngine & ";" & RPX_NativeInfo()
    RPX_DiagFinish "OPTIMAL", FixedFem: RFsActive = False
    Exit Function
Failed:
    FixedFem = "ERROR " & err.number & ":" & err.description & ";" & RPX_NativeInfo()
    RPX_DiagFinish "ERROR", FixedFem: RFsActive = False
End Function
