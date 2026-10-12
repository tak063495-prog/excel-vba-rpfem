Attribute VB_Name = "RPX_NativeLU"
Option Explicit
' Scoped native graph; compressed factors stay in VBA and support either backend.
#If Win64 Then
Private Declare PtrSafe Function NLCreate Lib "rpfem_native_x64.dll" Alias "rpx_lu_create" (ByVal n As Long, ByVal nnz As Long, ByVal maximum As Double, ByVal existing As Double, ByVal epoch As Long, ByVal order As LongPtr, ByVal streaming As Long, ByVal threshold As Double, ByRef handle As LongPtr) As Long
Private Declare PtrSafe Function NLInput Lib "rpfem_native_x64.dll" Alias "rpx_lu_input" (ByVal handle As LongPtr, ByVal phase As Long, ByVal first As Long, ByVal last As Long, ByVal ap As LongPtr, ByVal ai As LongPtr, ByVal av As LongPtr, ByVal symmetric As Long, ByVal equilibrate As Long) As Long
Private Declare PtrSafe Function NLStep Lib "rpfem_native_x64.dll" Alias "rpx_lu_step" (ByVal handle As LongPtr, ByVal limit As Long, ByRef written As Long) As Long
Private Declare PtrSafe Function NLFinish Lib "rpfem_native_x64.dll" Alias "rpx_lu_finish" (ByVal handle As LongPtr, ByRef lc As Long, ByRef uc As Long) As Long
Private Declare PtrSafe Function NLMetrics Lib "rpfem_native_x64.dll" Alias "rpx_lu_metrics" (ByVal handle As LongPtr, ByVal values As LongPtr) As Long
Private Declare PtrSafe Function NLExport Lib "rpfem_native_x64.dll" Alias "rpx_lu_export" (ByVal handle As LongPtr, ByVal arrayID As Long, ByVal buffer As LongPtr, ByVal bytes As Long) As Long
Private Declare PtrSafe Function NLDestroy Lib "rpfem_native_x64.dll" Alias "rpx_lu_destroy" (ByVal handle As LongPtr) As Long
Private Declare PtrSafe Function NLSolve Lib "rpfem_native_x64.dll" Alias "rpx_lu_backsolve" (ByVal n As Long, ByVal phase As Long, ByVal first As Long, ByVal last As Long, ByVal args As LongPtr, ByVal lc As Long, ByVal uc As Long, ByRef checks As Long) As Long
#End If
Private factorCalls As Long, solveCalls As Long
Public Sub RPX_NativeLUResetCounts()
    factorCalls = 0: solveCalls = 0
End Sub
Public Function RPX_NativeLUInfo() As String
    RPX_NativeLUInfo = ";hsd_lu=" & factorCalls & ";hsd_lu_solve=" & solveCalls
End Function
Private Sub Pulse()
    RPX_NativePulse
End Sub
Public Function RPX_NativeLUFactor(ByRef ap() As Long, ByRef ai() As Long, ByRef av() As Double, ByVal maximum As Double, ByVal existing As Double, ByVal epoch As Long, ByVal originalOrder As Boolean, ByVal equilibrate As Boolean, ByVal symmetric As Boolean, ByVal threshold As Double, ByVal permutation As Variant, ByVal streaming As Boolean, ByRef order() As Long, ByRef rowAt() As Long, ByRef rowScale() As Double, ByRef colScale() As Double, ByRef lp() As Long, ByRef li() As Long, ByRef lv() As Double, ByRef up() As Long, ByRef ui() As Long, ByRef uv() As Double, ByRef diagonal() As Double, ByRef metrics() As Double, ByRef accepted As Boolean) As Boolean
    Dim n As Long, nz As Long, i As Long, phase As Long, first As Long, last As Long, written As Long
    Dim lc As Long, uc As Long, code As Long, ignored As Long, seen() As Boolean
    Dim eno As Long, esource As String, etext As String
    If Not RPX_NativeAvailable() Then Exit Function
#If Win64 Then
    Dim handle As LongPtr
    On Error GoTo Failed
    ReDim metrics(0 To 9): accepted = False
    n = UBound(ap): nz = UBound(av) + 1
    If threshold < 0# Or threshold > 1# Then err.Raise 5, , "HSD_PIVOT_THRESHOLD_INVALID"
    If LBound(ap) <> 0 Or n < 1 Or ap(0) <> 0 Then err.Raise 5, , "HSD_KKT_STORAGE_INVALID"
    If n > 1048576 Then err.Raise 5, , "HSD_STABLE_UNAVAILABLE_SIZE"
    If LBound(av) <> 0 Or LBound(ai) <> 0 Or ap(n) <> nz Or UBound(ai) <> nz - 1 Then err.Raise 5, , "HSD_KKT_STORAGE_INVALID"
    ReDim order(0 To n - 1): ReDim seen(0 To n - 1)
    For i = 0 To n - 1
        order(i) = i: If originalOrder Then order(i) = XInv(i)
        If Not IsMissing(permutation) Then order(i) = CLng(permutation(i))
        If order(i) < 0 Or order(i) >= n Then err.Raise 5, , "HSD_ORDER_INVALID"
        If seen(order(i)) Then err.Raise 5, , "HSD_ORDER_INVALID"
        seen(order(i)) = True
    Next i
    Pulse
    factorCalls = factorCalls + 1
    RPX_NativeCheckCode NLCreate(n, nz, maximum, existing, epoch, VarPtr(order(0)), Abs(CLng(streaming)), threshold, handle)
    RPX_HoptEnter HoptLuStore
    For phase = 0 To 4
        For first = 0 To n - 1 Step 256
            Call Pulse: last = first + 256: If last > n Then last = n
            RPX_NativeCheckCode NLInput(handle, phase, first, last, VarPtr(ap(0)), VarPtr(ai(0)), VarPtr(av(0)), Abs(CLng(symmetric)), Abs(CLng(equilibrate)))
        Next first
    Next phase
    RPX_HoptLeave HoptLuStore
    RPX_HoptEnter HoptLuEliminate
    Do While written < n
        Call Pulse: RPX_AHOCheckBudget
        If written Mod 128 = 0 Then RPX_Stage "HSD sparse LU " & written & "/" & n
        code = NLStep(handle, 16, written): RPX_NativeCheckCode code
        If code = 1 Then GoTo PivotFailed
    Loop
    RPX_HoptLeave HoptLuEliminate
    RPX_HoptEnter HoptLuCompress
    RPX_NativeCheckCode NLFinish(handle, lc, uc)
    ReDim rowAt(0 To n - 1): ReDim rowScale(0 To n - 1): ReDim colScale(0 To n - 1)
    ReDim lp(0 To n): ReDim li(0 To lc): ReDim lv(0 To lc)
    ReDim up(0 To n): ReDim ui(0 To uc): ReDim uv(0 To uc): ReDim diagonal(0 To n - 1)
    RPX_NativeCheckCode NLExport(handle, 0, VarPtr(order(0)), 4 * n)
    RPX_NativeCheckCode NLExport(handle, 1, VarPtr(rowAt(0)), 4 * n)
    RPX_NativeCheckCode NLExport(handle, 2, VarPtr(rowScale(0)), 8 * n)
    RPX_NativeCheckCode NLExport(handle, 3, VarPtr(colScale(0)), 8 * n)
    RPX_NativeCheckCode NLExport(handle, 4, VarPtr(lp(0)), 4 * (n + 1))
    RPX_NativeCheckCode NLExport(handle, 5, VarPtr(li(0)), 4 * (lc + 1))
    RPX_NativeCheckCode NLExport(handle, 6, VarPtr(lv(0)), 8 * (lc + 1))
    RPX_NativeCheckCode NLExport(handle, 7, VarPtr(up(0)), 4 * (n + 1))
    RPX_NativeCheckCode NLExport(handle, 8, VarPtr(ui(0)), 4 * (uc + 1))
    RPX_NativeCheckCode NLExport(handle, 9, VarPtr(uv(0)), 8 * (uc + 1))
    RPX_NativeCheckCode NLExport(handle, 10, VarPtr(diagonal(0)), 8 * n)
    RPX_HoptLeave HoptLuCompress
    accepted = True
    GoTo Finished
PivotFailed:
    RPX_HoptLeave HoptLuEliminate
Finished:
    RPX_NativeCheckCode NLMetrics(handle, VarPtr(metrics(0)))
    RPX_NativeCheckCode NLDestroy(handle): handle = 0
    RPX_NativeLUFactor = True
    Exit Function
Failed:
    eno = err.number: esource = err.source: etext = err.description
    If handle <> 0 Then
        ignored = NLMetrics(handle, VarPtr(metrics(0)))
        ignored = NLDestroy(handle): handle = 0
    End If
    err.Raise eno, esource, etext
#End If
End Function
Public Function RPX_NativeLUSolve(ByVal n As Long, ByRef order() As Long, ByRef rowAt() As Long, ByRef rowScale() As Double, ByRef colScale() As Double, ByRef lp() As Long, ByRef li() As Long, ByRef lv() As Double, ByRef up() As Long, ByRef ui() As Long, ByRef uv() As Double, ByRef diagonal() As Double, ByRef rhs() As Double, ByRef work() As Double, ByRef answer() As Double, ByRef checks As Long) As Boolean
    Dim phase As Long, first As Long, last As Long, lc As Long, uc As Long
    If Not RPX_NativeAvailable() Then Exit Function
#If Win64 Then
    Dim args(0 To 13) As LongPtr
    If LBound(lp) <> 0 Or UBound(lp) <> n Or LBound(up) <> 0 Or UBound(up) <> n Then err.Raise 9, , "NATIVE_LU_FACTOR_BUFFER"
    lc = lp(n): uc = up(n)
    If lc < 0 Or uc < 0 Or LBound(li) <> 0 Or LBound(lv) <> 0 Or UBound(li) < lc - 1 Or UBound(lv) < lc - 1 Then err.Raise 9, , "NATIVE_LU_LOWER_BUFFER"
    If LBound(ui) <> 0 Or LBound(uv) <> 0 Or UBound(ui) < uc - 1 Or UBound(uv) < uc - 1 Then err.Raise 9, , "NATIVE_LU_UPPER_BUFFER"
    If LBound(order) <> 0 Or LBound(rowAt) <> 0 Or LBound(rowScale) <> 0 Or LBound(colScale) <> 0 Or LBound(diagonal) <> 0 Or LBound(rhs) <> 0 Or LBound(work) <> 0 Or LBound(answer) <> 0 Then err.Raise 9, , "NATIVE_LU_VECTOR_BASE"
    If UBound(order) <> n - 1 Or UBound(rowAt) <> n - 1 Or UBound(rowScale) <> n - 1 Or UBound(colScale) <> n - 1 Or UBound(diagonal) <> n - 1 Or UBound(rhs) <> n - 1 Or UBound(work) <> n - 1 Or UBound(answer) <> n - 1 Then err.Raise 9, , "NATIVE_LU_VECTOR_BUFFER"
    args(0) = VarPtr(order(0)): args(1) = VarPtr(rowAt(0)): args(2) = VarPtr(rowScale(0)): args(3) = VarPtr(colScale(0))
    args(4) = VarPtr(lp(0)): args(5) = VarPtr(li(0)): args(6) = VarPtr(lv(0)): args(7) = VarPtr(up(0))
    args(8) = VarPtr(ui(0)): args(9) = VarPtr(uv(0)): args(10) = VarPtr(diagonal(0)): args(11) = VarPtr(rhs(0))
    args(12) = VarPtr(work(0)): args(13) = VarPtr(answer(0))
    solveCalls = solveCalls + 1
    For phase = 0 To 2
        If phase = 1 Then
            last = n
            Do While last > 0
                Call Pulse: first = last - 256: If first < 0 Then first = 0
                RPX_NativeCheckCode NLSolve(n, phase, first, last, VarPtr(args(0)), lc, uc, checks)
                last = first
            Loop
        Else
            For first = 0 To n - 1 Step 256
                Call Pulse: last = first + 256: If last > n Then last = n
                RPX_NativeCheckCode NLSolve(n, phase, first, last, VarPtr(args(0)), lc, uc, checks)
            Next first
        End If
    Next phase
    RPX_NativeLUSolve = True
#End If
End Function
