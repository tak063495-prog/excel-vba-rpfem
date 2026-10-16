Attribute VB_Name = "RPX_NativeMesh"
Option Explicit
' G1R17: DLL owns a copied repair plan; VBA commits only after input/capacity checks.
#If Win64 Then
Private Declare PtrSafe Function NCreate Lib "rpfem_native_x64.dll" Alias "rpx_meshfan_create" (ByVal nn As Long, ByVal ne As Long, ByVal nd As Long, ByVal xy As LongPtr, ByVal tri As LongPtr, ByVal eligible As LongPtr, ByVal body As LongPtr, ByVal edges As LongPtr, ByRef handle As LongPtr) As Long
Private Declare PtrSafe Function NStep Lib "rpfem_native_x64.dll" Alias "rpx_meshfan_step" (ByVal handle As LongPtr, ByVal milliseconds As Long, ByRef done As Long) As Long
Private Declare PtrSafe Function NResult Lib "rpfem_native_x64.dll" Alias "rpx_meshfan_result" (ByVal handle As LongPtr, ByVal ne As Long, ByVal split As LongPtr, ByVal cx As LongPtr, ByVal cy As LongPtr, ByRef count As Long) As Long
Private Declare PtrSafe Function NDestroy Lib "rpfem_native_x64.dll" Alias "rpx_meshfan_destroy" (ByVal handle As LongPtr) As Long
#End If
Private jobs As Long, steps As Long
Public Function RPX_NativeMeshFans(ByRef split() As Long, ByRef cx() As Double, ByRef cy() As Double, ByRef count As Long) As Boolean
    If Not RPX_NativeAvailable() Then Exit Function
#If Win64 Then
    Dim xy() As Double, tri() As Long, eligible() As Long, body() As Long, edges() As Long
    Dim i As Long, k As Long, e As Long, done As Long, handle As LongPtr
    Dim failNumber As Long, failSource As String, failText As String
    On Error GoTo Failed
    RPX_AssertInputs "native_meshfan_begin"
    If RPX_LoadPhase <> "raw" Then err.Raise 5, "RPX_NativeMesh", "ZERO_STRESS_FAN_LOAD_PHASE"
    If rn < 3 Or re < 1 Or RNEdge < 1 Then err.Raise 9, "RPX_NativeMesh", "NATIVE_MESHFAN_SIZE"
    ReDim xy(0 To 2 * rn - 1): ReDim tri(0 To 3 * re - 1)
    ReDim eligible(0 To re - 1): ReDim body(0 To re - 1): ReDim edges(0 To 7 * RNEdge - 1)
    For i = 0 To rn - 1
        If i Mod 256 = 0 Then RPX_Stage "native mesh fan nodes " & i
        For k = 0 To 1: xy(2 * i + k) = RXY(i, k): Next k
    Next i
    For e = 0 To re - 1
        If e Mod 256 = 0 Then RPX_Stage "native mesh fan elements " & e
        For k = 0 To 2: tri(3 * e + k) = RTri(e, k): Next k
        If RRigidId(e) < 0 And RMaterials(RMatId(e)).cohesion = 0# Then eligible(e) = 1
        If Abs(RBody0(e, 0)) + Abs(RBody0(e, 1)) + Abs(RBody1(e, 0)) + Abs(RBody1(e, 1)) > 0# Then body(e) = 1
    Next e
    For i = 0 To RNEdge - 1
        If i Mod 256 = 0 Then RPX_Stage "native mesh fan edges " & i
        With REdges(i)
            edges(7 * i) = .element(0): edges(7 * i + 1) = .element(1)
            For k = 0 To 1
                edges(7 * i + 2 + k) = .loc(0, k): edges(7 * i + 4 + k) = .loc(1, k)
            Next k
            If .element(1) < 0 And .kind = "load" Then
                If Abs(.load0(0)) + Abs(.load0(1)) + Abs(.load1(0)) + Abs(.load1(1)) = 0# Then edges(7 * i + 6) = 1
            End If
        End With
    Next i
    RPX_AssertInputs "native_meshfan_copy"
    RPX_NativeCheckCode NCreate(rn, re, RNEdge, VarPtr(xy(0)), VarPtr(tri(0)), VarPtr(eligible(0)), VarPtr(body(0)), VarPtr(edges(0)), handle)
    jobs = jobs + 1
    Do
        RPX_Stage "native zero stress fan planning"
        RPX_NativeCheckCode NStep(handle, 20, done): steps = steps + 1
    Loop While done = 0
    RPX_NativeCheckCode NResult(handle, re, VarPtr(split(0)), VarPtr(cx(0)), VarPtr(cy(0)), count)
    RPX_NativeCheckCode NDestroy(handle): handle = 0
    RPX_AssertInputs "native_meshfan_result"
    RPX_NativeMeshFans = True
    Exit Function
Failed:
    failNumber = err.number: failSource = err.source: failText = err.description
    If handle <> 0 Then
        On Error Resume Next
        NDestroy handle
        On Error GoTo 0
    End If
    err.Raise failNumber, failSource, failText
#End If
End Function
Public Function RPX_NativeMeshInfo() As String
    RPX_NativeMeshInfo = ";meshfan_jobs=" & jobs & ";meshfan_steps=" & steps
End Function
Public Sub RPX_NativeMeshResetCounts()
    jobs = 0: steps = 0
End Sub
