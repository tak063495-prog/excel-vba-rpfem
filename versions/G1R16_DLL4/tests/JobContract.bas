Attribute VB_Name = "JobContract"
Option Explicit
Public DLL4Fault As String
#If Win64 Then
Private Declare PtrSafe Function HCount Lib "rpfem_native_x64.dll" Alias "rpx_hlinear_live_handles" () As Long
Private Declare PtrSafe Function RCount Lib "rpfem_native_x64.dll" Alias "rpx_refine_live_handles" () As Long
Private Declare PtrSafe Function LCount Lib "rpfem_native_x64.dll" Alias "rpx_lu_live_handles" () As Long
#End If
Public Sub DLL4Inject()
    Dim fault As String
    fault = DLL4Fault: DLL4Fault = ""
    Select Case fault
        Case "cancel": XCancel = True
        Case "input": ThisWorkbook.Worksheets("材料データ").Cells(5, 2).Value2 = 12345#
        Case "memory": err.Raise 7, "DLL4Test", "INJECTED_MEMORY"
        Case "subscript": err.Raise 9, "DLL4Test", "INJECTED_SUBSCRIPT"
        Case "unknown": err.Raise 513, "DLL4Test", "INJECTED_UNKNOWN"
    End Select
End Sub
Public Function DLL4Handles() As String
#If Win64 Then
    DLL4Handles = "h=" & HCount() & ";r=" & RCount() & ";l=" & LCount()
#End If
End Function
Public Function JobCase(ByVal hsd As Boolean, ByVal fault As String, ByVal prefix As String) As String
    On Error GoTo Failed
    XCancel = False
    RPX_InputEnd: RPX_AffineEnd: RPX_CertBeginRun: RPX_ReadInputs
    RPX_ResetPreparedState: RPX_HLURelease: RPX_SchurRelease: RPX_LDLRelease
    RPX_NativeSetMode "auto": RPX_StageConfigure: RPX_InputBegin
    RPX_Begin 3: RPX_Cost RPX_Unit(0)
    RPX_Equal RPX_Unit(1), 0.6: RPX_Equal RPX_Unit(2), 0.8
    RPX_Cone3 RPX_Unit(0), RPX_Unit(1), RPX_Unit(2)
    RPX_DiagStart "DLL4_JOB_CONTRACT": DLL4Fault = fault
    If hsd Then RPX_OptimizeHSD False Else RPX_Optimize False
    JobCase = "PASS;objective=" & XObjective & ";status=" & XStatus & ";" & DLL4Handles()
    RPX_DiagFinish "OPTIMAL", JobCase: RPX_InputEnd
    Exit Function
Failed:
    JobCase = "ERROR " & err.number & ":" & err.description & ";status=" & XStatus & ";" & DLL4Handles()
    RPX_DiagFinish "ERROR", JobCase: RPX_InputEnd
End Function
