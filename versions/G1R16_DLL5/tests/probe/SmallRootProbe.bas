Attribute VB_Name = "SmallRootProbe"
Option Explicit
Public Function SmallRoots(ByVal prefix As String) As String
    Dim upper As Double, lower As Double, upperCalls As Long, lowerCalls As Long
    On Error GoTo Failed
    XCancel = False: RPX_ReadInputs
    RPX_SlopeDefinition 16
    RMaterials(0).cohesion = 5#: RMaterials(0).friction = 45#: RMaterials(0).dilation = 45#
    RStressScale = 5#: RSubdivisions = 2: RFsTolerance = 0.0001
    Call RPX_GenerateMesh: RFsActive = True: RPX_TotalReferenceLoads
    Call RPX_FsBracketReset
    RFsSearchIncomplete = False: RFsLargestRootBracket = 0#: RFsCalls = 0
    RPX_DiagStart "G1R11_C5_small_roots"
    upper = RPX_FsEndpoint("upper"): upperCalls = RFsCalls
    RPX_ExportExchangeX prefix & "_upper.bin"
    lower = RPX_FsEndpoint("lower", 0.5 * upper, 0.1, upper): lowerCalls = RFsCalls - upperCalls
    RPX_ExportExchangeX prefix & "_lower.bin"
    MeshExport prefix & "_mesh.bin"
    If RFsSearchIncomplete Then Err.Raise 5, , "SMALL_ROOT_WIDTH_UNMET"
    SmallRoots = "PASS;lower=" & lower & ";upper=" & upper & ";lower_calls=" & lowerCalls & ";upper_calls=" & upperCalls & ";elements=" & re & ";audit=" & RAudit
    RPX_DiagFinish "ROOTS_VERIFIED", SmallRoots: RFsActive = False
    Exit Function
Failed:
    SmallRoots = "ERROR " & Err.Number & ":" & Err.Description
    RPX_DiagFinish "ERROR", SmallRoots: RFsActive = False
End Function
