Attribute VB_Name = "JobContract"
Option Explicit
Public DLL5Fault As String
Public DLL5FaultScope As String, DLL5InjectionHits As Long
#If Win64 Then
Private Declare PtrSafe Function HCount Lib "rpfem_native_x64.dll" Alias "rpx_hlinear_live_handles" () As Long
Private Declare PtrSafe Function RCount Lib "rpfem_native_x64.dll" Alias "rpx_refine_live_handles" () As Long
Private Declare PtrSafe Function LCount Lib "rpfem_native_x64.dll" Alias "rpx_lu_live_handles" () As Long
Private Declare PtrSafe Function SCount Lib "rpfem_native_x64.dll" Alias "rpx_schur_live_handles" () As Long
Private Declare PtrSafe Function KCount Lib "rpfem_native_x64.dll" Alias "rpx_krylov_live_handles" () As Long
Private Declare PtrSafe Function CCount Lib "rpfem_native_x64.dll" Alias "rpx_compile_live_handles" () As Long
#End If
Public Sub DLL5Arm(ByVal scope As String, ByVal fault As String)
    DLL5FaultScope = scope: DLL5Fault = fault: DLL5InjectionHits = 0: XCancel = False
End Sub
Public Function DLL5HitCount() As Long
    DLL5HitCount = DLL5InjectionHits
End Function
Public Sub DLL5Inject(Optional ByVal scope As String = "")
    Dim fault As String
    If scope <> DLL5FaultScope Then Exit Sub
    DLL5FaultScope = "$consumed": DLL5InjectionHits = DLL5InjectionHits + 1
    fault = DLL5Fault: DLL5Fault = ""
    Select Case fault
        Case "cancel": XCancel = True
        Case "input"
            ThisWorkbook.Worksheets("材料データ").Cells(5, 2).Value2 = 12345#
            If scope <> "" Then RPX_AssertInputs "test_persistent_edit"
        Case "memory": err.Raise 7, "DLL5Test", "INJECTED_MEMORY"
        Case "subscript": err.Raise 9, "DLL5Test", "INJECTED_SUBSCRIPT"
        Case "unknown": err.Raise 513, "DLL5Test", "INJECTED_UNKNOWN"
    End Select
End Sub
Public Function DLL5Handles() As String
#If Win64 Then
    DLL5Handles = "h=" & HCount() & ";r=" & RCount() & ";l=" & LCount() & ";s=" & SCount() & ";k=" & KCount() & ";c=" & CCount()
#End If
End Function
Public Function JobCase(ByVal hsd As Boolean, ByVal fault As String, ByVal prefix As String, Optional ByVal scope As String = "") As String
    On Error GoTo Failed
    XCancel = False
    RPX_InputEnd: RPX_AffineEnd: RPX_CertBeginRun: RPX_ReadInputs
    RPX_ResetPreparedState: RPX_HLURelease: RPX_SchurRelease: RPX_LDLRelease
    RPX_NativeSetMode "auto": RPX_StageConfigure: RPX_InputBegin
    RPX_Begin 3: RPX_Cost RPX_Unit(0)
    RPX_Equal RPX_Unit(1), 0.6: RPX_Equal RPX_Unit(2), 0.8
    RPX_Cone3 RPX_Unit(0), RPX_Unit(1), RPX_Unit(2)
    RPX_DiagStart "DLL5_JOB_CONTRACT": DLL5Arm scope, fault
    If hsd Then RPX_OptimizeHSD False Else RPX_Optimize False
    JobCase = "PASS;objective=" & XObjective & ";status=" & XStatus & ";hits=" & DLL5InjectionHits & ";" & DLL5Handles()
    RPX_DiagFinish "OPTIMAL", JobCase: RPX_InputEnd
    Exit Function
Failed:
    JobCase = "ERROR " & err.number & ":" & err.description & ";status=" & XStatus & ";hits=" & DLL5InjectionHits & ";" & DLL5Handles()
    RPX_DiagFinish "ERROR", JobCase: RPX_InputEnd
End Function
