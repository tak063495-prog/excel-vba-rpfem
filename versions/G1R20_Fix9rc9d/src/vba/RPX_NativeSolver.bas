Attribute VB_Name = "RPX_NativeSolver"
Option Explicit
#If Win64 Then
Private Declare PtrSafe Function NExactSpans Lib "rpfem_native_x64.dll" Alias "rpx_exact_spans" (ByVal count As Long, ByVal args As LongPtr, ByVal lengths As LongPtr, ByRef equal As Long) As Long
Private Declare PtrSafe Function NVector Lib "rpfem_native_x64.dll" Alias "rpx_solver_vectors" (ByVal op As Long, ByVal n As Long, ByVal first As Long, ByVal last As Long, ByVal args As LongPtr, ByVal alpha As Double, ByVal metrics As LongPtr) As Long
Private Declare PtrSafe Function NVectorStep Lib "rpfem_native_x64.dll" Alias "rpx_vector_step" (ByVal op As Long, ByVal n As Long, ByRef cursor As Long, ByVal args As LongPtr, ByVal alpha As Double, ByVal metrics As LongPtr, ByVal milliseconds As Long) As Long
Private Declare PtrSafe Function NUUpdate Lib "rpfem_native_x64.dll" Alias "rpx_solver_update" (ByVal nx As Long, ByVal nm As Long, ByVal ns As Long, ByVal args As LongPtr, ByVal alpha As Double) As Long
Private Declare PtrSafe Function NHCreate Lib "rpfem_native_x64.dll" Alias "rpx_hlinear_create" (ByVal n As Long, ByVal nx As Long, ByVal nb As Long, ByVal ns As Long, ByRef handle As LongPtr) As Long
Private Declare PtrSafe Function NHBegin Lib "rpfem_native_x64.dll" Alias "rpx_hlinear_begin" (ByVal handle As LongPtr, ByVal epoch As LongLong, ByVal factor As Long, ByVal kind As Long, ByVal unshifted As Long, ByVal args As LongPtr) As Long
Private Declare PtrSafe Function NHStep Lib "rpfem_native_x64.dll" Alias "rpx_hlinear_step" (ByVal handle As LongPtr, ByVal epoch As LongLong, ByVal factor As Long, ByVal opArgs As LongPtr, ByVal factorArgs As LongPtr, ByVal lu As LongPtr, ByVal milliseconds As Long, ByRef done As Long) As Long
Private Declare PtrSafe Function NHResult Lib "rpfem_native_x64.dll" Alias "rpx_hlinear_result" (ByVal handle As LongPtr, ByVal epoch As LongLong, ByVal best As LongPtr, ByVal scaled As LongPtr, ByVal bdd As LongPtr, ByVal metrics As LongPtr) As Long
Private Declare PtrSafe Function NHDestroy Lib "rpfem_native_x64.dll" Alias "rpx_hlinear_destroy" (ByVal handle As LongPtr) As Long
Private Declare PtrSafe Function NKCreate Lib "rpfem_native_x64.dll" Alias "rpx_krylov_create" (ByVal n As Long, ByVal nx As Long, ByVal nb As Long, ByVal ns As Long, ByRef handle As LongPtr) As Long
Private Declare PtrSafe Function NKDestroy Lib "rpfem_native_x64.dll" Alias "rpx_krylov_destroy" (ByVal handle As LongPtr) As Long
Private Declare PtrSafe Function NKBegin Lib "rpfem_native_x64.dll" Alias "rpx_krylov_begin" (ByVal handle As LongPtr, ByVal serial As LongLong, ByVal factor As Long, ByVal kind As Long, ByVal unshifted As Long, ByVal bdd As LongPtr, ByVal answer As LongPtr, ByVal best As Double) As Long
Private Declare PtrSafe Function NKStep Lib "rpfem_native_x64.dll" Alias "rpx_krylov_step" (ByVal handle As LongPtr, ByVal serial As LongLong, ByVal factor As Long, ByVal opArgs As LongPtr, ByVal factorArgs As LongPtr, ByVal lu As LongPtr, ByVal milliseconds As Long, ByRef done As Long) As Long
Private Declare PtrSafe Function NKResult Lib "rpfem_native_x64.dll" Alias "rpx_krylov_result" (ByVal handle As LongPtr, ByVal serial As LongLong, ByVal answer As LongPtr, ByVal metrics As LongPtr) As Long
Private Declare PtrSafe Function NKBytes Lib "rpfem_native_x64.dll" Alias "rpx_krylov_bytes" (ByVal handle As LongPtr, ByRef resident As LongLong) As Long
Private hHandle As LongPtr, hSerial As LongLong
Private kHandle As LongPtr, kSerial As LongLong
#End If
Private hn As Long, hnx As Long, hnb As Long, hns As Long, hJobs As Long, hSteps As Long
Private calls As Double
Private kn As Long, kJobs As Long, kSteps As Long, kResident As Double
Public Function RPX_NativeExactSpans(ByRef args() As LongPtr, ByRef lengths() As Long, ByRef equal As Boolean) As Boolean
    If Not RPX_NativeAvailable() Then Exit Function
#If Win64 Then
    Dim same As Long, count As Long
    If LBound(lengths) <> 0 Or LBound(args) <> 0 Then err.Raise 9, , "NATIVE_EXACT_SPAN_BASE"
    count = UBound(lengths) + 1
    If UBound(args) <> 2 * count - 1 Then err.Raise 9, , "NATIVE_EXACT_SPAN_SHAPE"
    RPX_NativeCheckCode NExactSpans(count, VarPtr(args(0)), VarPtr(lengths(0)), same)
    equal = (same <> 0): RPX_NativeExactSpans = True
#End If
End Function
Public Sub RPX_NativeSolverResetCounts()
    calls = 0#: hJobs = 0: hSteps = 0: kJobs = 0: kSteps = 0
End Sub
Public Function RPX_NativeSolverInfo() As String
    RPX_NativeSolverInfo = ";solver_vector_calls=" & calls & ";hsd_jobs=" & hJobs & ";hsd_job_steps=" & hSteps & ";krylov_jobs=" & kJobs & ";krylov_steps=" & kSteps
End Function
Public Function RPX_NativeVectors(ByVal op As Long, ByVal n As Long, ByRef args() As LongPtr, ByRef metrics() As Double, Optional ByVal alpha As Double = 0#) As Boolean
    If Not RPX_NativeAvailable() Then Exit Function
    If n < 0 Then err.Raise 9, , "NATIVE_SOLVER_SIZE"
    If n = 0 Then RPX_NativeVectors = True: Exit Function
#If Win64 Then
    Dim first As Long, milliseconds As Long
    milliseconds = 50: If RPX_StageIntervalValue() = 0# Then milliseconds = 0
    Do While first < n
        RPX_NativePulse
        calls = calls + 1
        RPX_NativeCheckCode NVectorStep(op, n, first, VarPtr(args(0)), alpha, VarPtr(metrics(0)), milliseconds)
    Loop
    RPX_NativeVectors = True
#End If
End Function

Public Sub RPX_NativeHLinearRelease()
    RPX_NativeKrylovRelease
#If Win64 Then
    Dim code As Long
    If hHandle <> 0 Then
        code = NHDestroy(hHandle): hHandle = 0
        hn = 0: hnx = 0: hnb = 0: hns = 0
        RPX_NativeCheckCode code
    End If
#End If
End Sub
Public Sub RPX_NativeKrylovRelease()
#If Win64 Then
    Dim code As Long
    If kHandle <> 0 Then
        code = NKDestroy(kHandle): kHandle = 0: kn = 0: kResident = 0#
        RPX_NativeCheckCode code
    End If
#End If
End Sub
#If Win64 Then
Public Function RPX_NativeKrylovBegin(ByRef bdd() As RPX_DD, ByRef answer() As RPX_DD, ByVal best As Double, ByVal factor As Long, ByVal kind As Long, ByVal unshifted As Boolean) As LongLong
    Dim n As Long, resident As LongLong
    RPX_NativeKrylovRelease
    n = XN + XM
    If LBound(bdd) <> 0 Or UBound(bdd) <> n - 1 Or LBound(answer) <> 0 Or UBound(answer) <> n - 1 Then err.Raise 9, , "NATIVE_KRYLOV_BUFFER"
    RPX_CapacityCheck "native_hsd_krylov", 600# * n + 32# * XS + 16384#
    RPX_NativeCheckCode NKCreate(n, XN, XB, XS, kHandle): kn = n
    RPX_NativeCheckCode NKBytes(kHandle, resident): kResident = CDbl(resident)
    kSerial = kSerial + 1
    RPX_NativeCheckCode NKBegin(kHandle, kSerial, factor, kind, Abs(CLng(unshifted)), VarPtr(bdd(0)), VarPtr(answer(0)), best)
    kJobs = kJobs + 1: RPX_NativeKrylovBegin = kSerial
End Function
Public Sub RPX_NativeKrylovStep(ByVal serial As LongLong, ByVal factor As Long, ByRef opArgs() As LongPtr, ByRef factorArgs() As LongPtr, ByVal lu As LongPtr, ByRef done As Long)
    Dim milliseconds As Long
    milliseconds = 50: If RPX_StageIntervalValue() = 0# Then milliseconds = 0
    RPX_NativePulse
    RPX_NativeCheckCode NKStep(kHandle, serial, factor, VarPtr(opArgs(0)), VarPtr(factorArgs(0)), lu, milliseconds, done)
    kSteps = kSteps + 1
End Sub
Public Sub RPX_NativeKrylovResult(ByVal serial As LongLong, ByRef answer() As RPX_DD, ByRef metrics() As Double)
    ReDim answer(0 To kn - 1): ReDim metrics(0 To 3)
    RPX_NativeCheckCode NKResult(kHandle, serial, VarPtr(answer(0)), VarPtr(metrics(0)))
End Sub
Public Function RPX_NativeHLinearBegin(ByRef rhs() As Double, ByRef opArgs() As LongPtr, ByVal kind As Long, ByVal factor As Long, ByVal unshifted As Boolean) As LongLong
    Dim n As Long, initArgs(0 To 2) As LongPtr
    n = XN + XM
    If hHandle <> 0 Then If hn <> n Or hnx <> XN Or hnb <> XB Or hns <> XS Then RPX_NativeHLinearRelease
    If hHandle = 0 Then
        RPX_CapacityCheck "native_hlinear", 144# * n + 32# * XS
        RPX_NativeCheckCode NHCreate(n, XN, XB, XS, hHandle)
        hn = n: hnx = XN: hnb = XB: hns = XS
    End If
    hSerial = hSerial + 1
    initArgs(0) = VarPtr(rhs(0)): initArgs(1) = opArgs(0): initArgs(2) = opArgs(1)
    RPX_NativeCheckCode NHBegin(hHandle, hSerial, factor, kind, Abs(CLng(unshifted)), VarPtr(initArgs(0)))
    hJobs = hJobs + 1: RPX_NativeHLinearBegin = hSerial
End Function
Public Sub RPX_NativeHLinearStep(ByVal serial As LongLong, ByVal factor As Long, ByRef opArgs() As LongPtr, ByRef factorArgs() As LongPtr, ByVal lu As LongPtr, ByRef done As Long)
    RPX_NativePulse
    RPX_NativeCheckCode NHStep(hHandle, serial, factor, VarPtr(opArgs(0)), VarPtr(factorArgs(0)), lu, 50, done)
    hSteps = hSteps + 1
End Sub
Public Sub RPX_NativeHLinearResult(ByVal serial As LongLong, ByRef best() As RPX_DD, ByRef scaled() As Double, ByRef bdd() As RPX_DD, ByRef metrics() As Double)
    ReDim best(0 To hn - 1): ReDim scaled(0 To hn - 1): ReDim bdd(0 To hn - 1): ReDim metrics(0 To 13)
    RPX_NativeCheckCode NHResult(hHandle, serial, VarPtr(best(0)), VarPtr(scaled(0)), VarPtr(bdd(0)), VarPtr(metrics(0)))
End Sub
#End If

Public Function RPX_NativeNormalUpdate(ByVal alpha As Double, ByRef dx() As Double, ByRef dy() As Double, ByRef ds() As Double, ByRef dz() As Double) As Boolean
    If Not RPX_NativeAvailable() Then Exit Function
#If Win64 Then
    Dim args(0 To 7) As LongPtr
    RPX_NativePulse
    args(0) = VarPtr(xx(0)): args(1) = VarPtr(dx(0)): args(2) = VarPtr(xy(0)): args(3) = VarPtr(dy(0))
    args(4) = VarPtr(XSlack(0)): args(5) = VarPtr(ds(0)): args(6) = VarPtr(XDual(0)): args(7) = VarPtr(dz(0))
    RPX_NativeCheckCode NUUpdate(XN, XM, XS, VarPtr(args(0)), alpha)
    RPX_NativeNormalUpdate = True
#End If
End Function

Public Function RPX_NativeHLinearBytes() As Double
#If Win64 Then
    If hHandle <> 0 Then RPX_NativeHLinearBytes = 144# * hn + 32# * hns + 4096#
    RPX_NativeHLinearBytes = RPX_NativeHLinearBytes + kResident
#End If
End Function
