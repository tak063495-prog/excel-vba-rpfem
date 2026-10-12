Attribute VB_Name = "RPX_Native"
Option Explicit
' Optional x64 numeric backend. VBA owns every array and symbolic pattern.
' No PATH search, trust setting changes, or blanket exception recovery.
#If Win64 Then
Private Declare PtrSafe Function LoadLibraryExW Lib "kernel32" (ByVal path As LongPtr, ByVal file As LongPtr, ByVal flags As Long) As LongPtr
Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal module As LongPtr, ByVal name As String) As LongPtr
Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long
Private Declare PtrSafe Function GetModuleHandleW Lib "kernel32" (ByVal name As LongPtr) As LongPtr
Private Declare PtrSafe Function GetModuleFileNameW Lib "kernel32" (ByVal module As LongPtr, ByVal buffer As LongPtr, ByVal size As Long) As Long
Private Declare PtrSafe Function NativeAbi Lib "rpfem_native_x64.dll" Alias "rpx_abi" () As Long
Private Declare PtrSafe Function NativeBulkReady Lib "rpfem_native_x64.dll" Alias "rpx_bulk_ready" () As Long
Private Declare PtrSafe Function NativeNumeric Lib "rpfem_native_x64.dll" Alias "rpx_numeric" (ByVal n As Long, ByVal first As Long, ByVal last As Long, ByVal ap As LongPtr, ByVal ai As LongPtr, ByVal av As LongPtr, ByVal cp As LongPtr, ByVal ri As LongPtr, ByVal rp As LongPtr, ByVal rn As LongPtr, ByVal wa As LongPtr, ByVal lv As LongPtr, ByVal dv As LongPtr, ByVal yv As LongPtr, ByRef minimum As Double, ByRef failure As Long) As Long
Private Declare PtrSafe Function NativeDDNumeric Lib "rpfem_native_x64.dll" Alias "rpx_dd_numeric" (ByVal n As Long, ByVal first As Long, ByVal last As Long, ByVal ap As LongPtr, ByVal ai As LongPtr, ByVal av As LongPtr, ByVal cp As LongPtr, ByVal ri As LongPtr, ByVal rp As LongPtr, ByVal rn As LongPtr, ByVal wa As LongPtr, ByVal lv As LongPtr, ByVal dv As LongPtr, ByVal yv As LongPtr, ByRef minimum As Double, ByRef failure As Long) As Long
Private Declare PtrSafe Function NativeBacksolve Lib "rpfem_native_x64.dll" Alias "rpx_backsolve" (ByVal n As Long, ByVal phase As Long, ByVal first As Long, ByVal last As Long, ByVal cp As LongPtr, ByVal ri As LongPtr, ByVal lv As LongPtr, ByVal dv As LongPtr, ByVal rhs As LongPtr, ByVal answer As LongPtr) As Long
Private Declare PtrSafe Function NativeDDBacksolve Lib "rpfem_native_x64.dll" Alias "rpx_dd_backsolve" (ByVal n As Long, ByVal phase As Long, ByVal first As Long, ByVal last As Long, ByVal cp As LongPtr, ByVal ri As LongPtr, ByVal lv As LongPtr, ByVal dv As LongPtr, ByVal rhs As LongPtr, ByVal answer As LongPtr) As Long
Private Declare PtrSafe Function NativeOrderCreate Lib "rpfem_native_x64.dll" Alias "rpx_order_create" (ByVal n As Long, ByVal mode As Long, ByRef handle As LongPtr) As Long
Private Declare PtrSafe Function NativeOrderFeed Lib "rpfem_native_x64.dll" Alias "rpx_order_feed" (ByVal handle As LongPtr, ByVal first As Long, ByVal last As Long, ByVal ap As LongPtr, ByVal apSize As Long, ByVal ai As LongPtr, ByVal aiSize As Long) As Long
Private Declare PtrSafe Function NativeOrderStart Lib "rpfem_native_x64.dll" Alias "rpx_order_start" (ByVal handle As LongPtr) As Long
Private Declare PtrSafe Function NativeOrderStep Lib "rpfem_native_x64.dll" Alias "rpx_order_step" (ByVal handle As LongPtr, ByVal limit As Long, ByRef written As Long, ByVal output As LongPtr, ByVal outputSize As Long) As Long
Private Declare PtrSafe Function NativeOrderDestroy Lib "rpfem_native_x64.dll" Alias "rpx_order_destroy" (ByVal handle As LongPtr) As Long
Private nativeHandle As LongPtr
#End If
Private checked As Boolean, available As Boolean, forcedVBA As Boolean
Private reason As String
Private numericCalls As Long, ddCalls As Long, solveCalls As Long, orderCalls As Long, columnOrderCalls As Long

Public Sub RPX_NativeSetMode(ByVal mode As String)
    If mode <> "auto" And mode <> "vba" Then err.Raise 5, , "NATIVE_MODE_INVALID"
    forcedVBA = (mode = "vba"): checked = False: available = False
    numericCalls = 0: ddCalls = 0: solveCalls = 0
    orderCalls = 0: columnOrderCalls = 0
    RPX_HLURelease
    RPX_NativeRefineRelease
    RPX_NativeHLinearRelease
    RPX_NativeSolverResetCounts
    RPX_NativeRefineResetCounts
    RPX_NativeBulkResetCounts
    RPX_NativeLUResetCounts
End Sub
Public Function RPX_NativeAvailable() As Boolean
    Dim path As String, exportName As Variant, dllName As String, loadedPath As String, size As Long
    If forcedVBA Then reason = "VBA_SELECTED": Exit Function
    If checked Then RPX_NativeAvailable = available: Exit Function
    checked = True
#If Win64 Then
    path = ThisWorkbook.path & Application.PathSeparator & "rpfem_native_x64.dll"
    If Len(Dir$(path)) = 0 Then reason = "DLL_NOT_FOUND": Exit Function
    dllName = "rpfem_native_x64.dll"
    nativeHandle = GetModuleHandleW(StrPtr(dllName))
    If nativeHandle <> 0 Then
        loadedPath = String$(32768, vbNullChar)
        size = GetModuleFileNameW(nativeHandle, StrPtr(loadedPath), Len(loadedPath))
        If size = 0 Or size >= Len(loadedPath) Then reason = "DLL_PATH_UNVERIFIED": Exit Function
        If StrComp(Left$(loadedPath, size), path, vbTextCompare) <> 0 Then reason = "DLL_ALREADY_LOADED_OTHER_PATH": Exit Function
    Else
        nativeHandle = LoadLibraryExW(StrPtr(path), 0, &H1100)
    End If
    If nativeHandle = 0 Then reason = "DLL_LOAD_FAILED_" & CStr(GetLastError()): Exit Function
    For Each exportName In Array("rpx_abi", "rpx_numeric", "rpx_dd_numeric", "rpx_backsolve", "rpx_dd_backsolve", "rpx_order_create", "rpx_order_feed", "rpx_order_start", "rpx_order_step", "rpx_order_destroy", "rpx_scaling", "rpx_gproduct", "rpx_gtadd", "rpx_kkt_cones", "rpx_hoperator", "rpx_lu_create", "rpx_lu_input", "rpx_lu_step", "rpx_lu_finish", "rpx_lu_metrics", "rpx_lu_export", "rpx_lu_destroy", "rpx_lu_backsolve", "rpx_kkt_init", "rpx_kkt_diagonal", "rpx_kkt_scale", "rpx_corrector")
        If GetProcAddress(nativeHandle, CStr(exportName)) = 0 Then reason = "DLL_ENTRY_MISSING": Exit Function
    Next exportName
    If NativeAbi() <> &H160004 Then reason = "DLL_ABI_MISMATCH": Exit Function
    If GetProcAddress(nativeHandle, "rpx_bulk_ready") = 0 Then reason = "DLL_ENTRY_MISSING": Exit Function
    If NativeBulkReady() <> 0 Then reason = "DLL_NT_RUNTIME_UNAVAILABLE": Exit Function
    available = True: reason = "DLL_READY"
#Else
    reason = "VBA_32BIT_OR_LEGACY"
#End If
    RPX_NativeAvailable = available
End Function
Public Function RPX_NativeInfo() As String
    Dim ready As Boolean
    ready = RPX_NativeAvailable()
    RPX_NativeInfo = reason & ";numeric=" & numericCalls & ";dd_numeric=" & ddCalls & ";backsolve=" & solveCalls & ";ordering=" & orderCalls & ";column_ordering=" & columnOrderCalls & RPX_NativeBulkInfo() & RPX_NativeLUInfo() & RPX_NativeSolverInfo() & RPX_NativeRefineInfo()
End Function
Public Sub RPX_NativeCheckCode(ByVal code As Long)
    Select Case code
        Case 0, 1
        Case 201: err.Raise 5, , "HSD_DD_NONFINITE_PRODUCT"
        Case 202: err.Raise 5, , "HSD_DD_PRODUCT_UNDERFLOW"
        Case 203: err.Raise 5, , "HSD_DD_ZERO_DIVISOR"
        Case 204: err.Raise 5, , "HSD_DD_NONFINITE"
        Case 205: err.Raise 11, "RPX_Native", "NATIVE_DOUBLE_ZERO_DIVISOR"
        Case 206: err.Raise 6, "RPX_Native", "NATIVE_DOUBLE_OVERFLOW"
        Case 301: err.Raise 5, , "NT_NONINTERIOR_POINT"
        Case 302: err.Raise 5, , "NT_CORRECTOR_SINGULAR"
        Case 303: err.Raise 5, "RPX_NTPrecision", "NT_POWER_UNSUPPORTED"
        Case 304: err.Raise 5, "RPX_Solve", "NT_NONINTERIOR_POINT"
        Case 305: err.Raise 5, "RPX_NTPrecision", "NT_NONINTERIOR_POINT"
        Case 311: err.Raise 5, "RPX_HSDLu", "HSD_STABLE_UNAVAILABLE_MEMORY"
        Case 312: err.Raise 5, , "HSD_KKT_STORAGE_INVALID"
        Case 313: err.Raise 5, , "HSD_STABLE_UNAVAILABLE_SIZE"
        Case 314: err.Raise 5, , "HSD_ORDER_INVALID"
        Case 315: err.Raise 5, , "HSD_NONFINITE_MATRIX"
        Case 316: err.Raise 5, , "HSD_NONFINITE_FACTOR"
        Case 317: err.Raise 5, , "HSD_PIVOT_THRESHOLD_INVALID"
        Case 321: err.Raise 5, , "HSD_NONFINITE_COMMIT"
        Case 322: err.Raise 5, , "HSD_NONFINITE_RHS"
        Case 318: err.Raise 5, "RPX_HSDLu", "HSD_FACTOR_CONTEXT_MISMATCH"
        Case 319: err.Raise 5, , "HSD_RHS_SHAPE_MISMATCH"
        Case 320: err.Raise 5, , "HSD_NONFINITE_SOLUTION"
        Case 207: err.Raise 7, "RPX_Native", "NATIVE_OUT_OF_MEMORY"
        Case 208: err.Raise 513, "RPX_Native", "NATIVE_UNEXPECTED_EXCEPTION"
        Case Else: err.Raise 5, "RPX_Native", "NATIVE_KERNEL_CONTRACT_FAILED " & CStr(code)
    End Select
End Sub
Private Sub Poll()
    DoEvents
    If XCancel Then err.Raise 18, , "ANALYSIS_CANCELLED"
End Sub
Public Function RPX_NativeOrdering(ByVal n As Long, ByRef columns() As Object, ByRef permutation() As Long, ByRef inverse() As Long, Optional ByVal normalGraph As Boolean = False) As Boolean
    Dim ap() As Long, ai() As Long, i As Long, p As Long, count As Long, key As Variant
    Dim first As Long, last As Long, written As Long, mode As Long, ignored As Long
    Dim failNumber As Long, failSource As String, failText As String
    If Not RPX_NativeAvailable() Then Exit Function
#If Win64 Then
    Dim handle As LongPtr
    On Error GoTo Failed
    If n <= 0 Or LBound(columns) <> 0 Or UBound(columns) < n - 1 Then err.Raise 9, "RPX_Native", "NATIVE_ORDERING_SIZE"
    ReDim ap(0 To n)
    For i = 0 To n - 1
        If i Mod 256 = 0 Then Call Poll
        count = count + columns(i).count: ap(i + 1) = count
    Next i
    If count = 0 Then ReDim ai(0 To 0) Else ReDim ai(0 To count - 1)
    For i = 0 To n - 1
        If i Mod 256 = 0 Then Call Poll
        For Each key In columns(i).keys
            If p >= ap(i + 1) Then err.Raise 5, "RPX_Native", "NATIVE_ORDERING_INPUT_CHANGED"
            ai(p) = CLng(key)
            If ai(p) < 0 Or ai(p) >= n Then err.Raise 9, "RPX_Native", "NATIVE_ORDERING_COLUMN"
            p = p + 1
        Next key
        If p <> ap(i + 1) Then err.Raise 5, "RPX_Native", "NATIVE_ORDERING_INPUT_CHANGED"
    Next i
    If normalGraph Then mode = 1
    RPX_NativeCheckCode NativeOrderCreate(n, mode, handle)
    For first = 0 To n - 1 Step 256
        Call Poll: last = first + 256: If last > n Then last = n
        RPX_NativeCheckCode NativeOrderFeed(handle, first, last, VarPtr(ap(0)), n + 1, VarPtr(ai(0)), count)
    Next first
    RPX_NativeCheckCode NativeOrderStart(handle)
    ReDim permutation(0 To n - 1): ReDim inverse(0 To n - 1)
    Do While written < n
        If written Mod 8192 = 0 Then RPX_Stage "ordering " & written & "/" & n
        Call Poll
        RPX_AssertInputs "native_ordering"
        RPX_NativeCheckCode NativeOrderStep(handle, 64, written, VarPtr(permutation(0)), n)
    Loop
    For i = 0 To n - 1: inverse(permutation(i)) = i: Next i
    RPX_NativeCheckCode NativeOrderDestroy(handle): handle = 0
    If normalGraph Then columnOrderCalls = columnOrderCalls + 1 Else orderCalls = orderCalls + 1
    RPX_NativeOrdering = True: Exit Function
Failed:
    failNumber = err.number: failSource = err.source: failText = err.description
    ' Destroy is allocation-free and cannot replace the original VBA exception.
    If handle <> 0 Then ignored = NativeOrderDestroy(handle): handle = 0
    err.Raise failNumber, failSource, failText
#End If
End Function
Private Sub PatternBounds(ByVal n As Long, ByRef ap() As Long, ByRef ai() As Long, ByRef cp() As Long, ByRef ri() As Long, ByRef rp() As Long, ByRef rn() As Long, ByRef wa() As Long)
    ' Validate managed buffer extents before exposing element addresses to C.
    If n <= 0 Or LBound(ap) <> 0 Or UBound(ap) <> n Or LBound(cp) <> 0 Or UBound(cp) <> n Or LBound(rp) <> 0 Or UBound(rp) <> n Then err.Raise 9, "RPX_Native", "NATIVE_PATTERN_SIZE_MISMATCH"
    If ap(n) < 0 Or cp(n) < 0 Then err.Raise 5, , "NATIVE_PATTERN_SIZE_INVALID"
    If ap(n) > 0 Then
        If LBound(ai) <> 0 Or UBound(ai) < ap(n) - 1 Then err.Raise 9, "RPX_Native", "NATIVE_INPUT_BUFFER_SHORT"
    End If
    If cp(n) > 0 Then
        If LBound(ri) <> 0 Or UBound(ri) < cp(n) - 1 Or LBound(rn) <> 0 Or UBound(rn) < cp(n) - 1 Or LBound(wa) <> 0 Or UBound(wa) < cp(n) - 1 Then err.Raise 9, "RPX_Native", "NATIVE_REACH_BUFFER_SHORT"
    End If
End Sub
Private Sub SolveBounds(ByVal n As Long, ByRef cp() As Long, ByRef ri() As Long, ByRef rhs() As Double)
    If n <= 0 Or LBound(cp) <> 0 Or UBound(cp) <> n Or LBound(rhs) <> 0 Or UBound(rhs) < n - 1 Then err.Raise 9, "RPX_Native", "NATIVE_SOLVE_BUFFER_SHORT"
    If cp(n) < 0 Then err.Raise 5, , "NATIVE_PATTERN_SIZE_INVALID"
    If LBound(ri) <> 0 Or UBound(ri) < cp(n) - 1 Then err.Raise 9, "RPX_Native", "NATIVE_ROW_BUFFER_SHORT"
End Sub

Public Function RPX_NativeNumeric(ByVal n As Long, ByRef ap() As Long, ByRef ai() As Long, ByRef av() As Double, ByRef cp() As Long, ByRef ri() As Long, ByRef rp() As Long, ByRef rn() As Long, ByRef wa() As Long, ByRef lv() As Double, ByRef dv() As Double, ByRef yv() As Double, ByRef minimum As Double, ByRef failure As Long, ByRef accepted As Boolean) As Boolean
    Dim first As Long, last As Long, code As Long
    If Not RPX_NativeAvailable() Then Exit Function
#If Win64 Then
    PatternBounds n, ap, ai, cp, ri, rp, rn, wa
    ' A diagonal pattern has empty reach arrays; keep its ordinary VBA path.
    If cp(n) = 0 Then Exit Function
    If LBound(av) <> 0 Or UBound(av) < ap(n) - 1 Or LBound(lv) <> 0 Or UBound(lv) < cp(n) - 1 Or LBound(dv) <> 0 Or UBound(dv) < n - 1 Or LBound(yv) <> 0 Or UBound(yv) < n - 1 Then err.Raise 9, "RPX_Native", "NATIVE_NUMERIC_BUFFER_SHORT"
    numericCalls = numericCalls + 1
    For first = 0 To n - 1 Step 128
        Call Poll: last = first + 128: If last > n Then last = n
        code = NativeNumeric(n, first, last, VarPtr(ap(0)), VarPtr(ai(0)), VarPtr(av(0)), VarPtr(cp(0)), VarPtr(ri(0)), VarPtr(rp(0)), VarPtr(rn(0)), VarPtr(wa(0)), VarPtr(lv(0)), VarPtr(dv(0)), VarPtr(yv(0)), minimum, failure)
        RPX_NativeCheckCode code
        If code = 1 Then RPX_NativeNumeric = True: accepted = False: Exit Function
    Next first
    accepted = True: RPX_NativeNumeric = True
#End If
End Function

Public Function RPX_NativeDDNumeric(ByVal n As Long, ByRef ap() As Long, ByRef ai() As Long, ByRef av() As RPX_DD, ByRef cp() As Long, ByRef ri() As Long, ByRef rp() As Long, ByRef rn() As Long, ByRef wa() As Long, ByRef lv() As RPX_DD, ByRef dv() As RPX_DD, ByRef yv() As RPX_DD, ByRef minimum As Double, ByRef failure As Long, ByRef accepted As Boolean) As Boolean
    Dim first As Long, last As Long, code As Long
    If Not RPX_NativeAvailable() Then Exit Function
#If Win64 Then
    PatternBounds n, ap, ai, cp, ri, rp, rn, wa
    If cp(n) = 0 Then Exit Function
    If LBound(av) <> 0 Or UBound(av) < ap(n) - 1 Or LBound(lv) <> 0 Or UBound(lv) < cp(n) - 1 Or LBound(dv) <> 0 Or UBound(dv) < n - 1 Or LBound(yv) <> 0 Or UBound(yv) < n - 1 Then err.Raise 9, "RPX_Native", "NATIVE_DD_BUFFER_SHORT"
    ddCalls = ddCalls + 1
    For first = 0 To n - 1 Step 32
        Call Poll: last = first + 32: If last > n Then last = n
        If first Mod 256 = 0 Then RPX_Stage "HSD DD LDL " & first & "/" & n
        code = NativeDDNumeric(n, first, last, VarPtr(ap(0)), VarPtr(ai(0)), VarPtr(av(0)), VarPtr(cp(0)), VarPtr(ri(0)), VarPtr(rp(0)), VarPtr(rn(0)), VarPtr(wa(0)), VarPtr(lv(0)), VarPtr(dv(0)), VarPtr(yv(0)), minimum, failure)
        RPX_NativeCheckCode code
        If code = 1 Then RPX_NativeDDNumeric = True: accepted = False: Exit Function
    Next first
    accepted = True: RPX_NativeDDNumeric = True
#End If
End Function

Public Function RPX_NativeBacksolve(ByVal n As Long, ByRef cp() As Long, ByRef ri() As Long, ByRef lv() As Double, ByRef dv() As Double, ByRef rhs() As Double, ByRef answer() As Double) As Boolean
    Dim phase As Long, first As Long, last As Long, code As Long
    If Not RPX_NativeAvailable() Then Exit Function
#If Win64 Then
    SolveBounds n, cp, ri, rhs
    If LBound(lv) <> 0 Or UBound(lv) < cp(n) - 1 Or LBound(dv) <> 0 Or UBound(dv) < n - 1 Or LBound(answer) <> 0 Or UBound(answer) < n - 1 Then err.Raise 9, "RPX_Native", "NATIVE_SOLVE_BUFFER_SHORT"
    solveCalls = solveCalls + 1
    For phase = 0 To 2
        For first = 0 To n - 1 Step 256
            Call Poll: last = first + 256: If last > n Then last = n
            code = NativeBacksolve(n, phase, first, last, VarPtr(cp(0)), VarPtr(ri(0)), VarPtr(lv(0)), VarPtr(dv(0)), VarPtr(rhs(0)), VarPtr(answer(0)))
            RPX_NativeCheckCode code
        Next first
    Next phase
    last = n
    Do While last > 0
        Call Poll: first = last - 256: If first < 0 Then first = 0
        code = NativeBacksolve(n, 3, first, last, VarPtr(cp(0)), VarPtr(ri(0)), VarPtr(lv(0)), VarPtr(dv(0)), VarPtr(rhs(0)), VarPtr(answer(0)))
        RPX_NativeCheckCode code: last = first
    Loop
    RPX_NativeBacksolve = True
#End If
End Function

Public Function RPX_NativeDDBacksolve(ByVal n As Long, ByRef cp() As Long, ByRef ri() As Long, ByRef lv() As RPX_DD, ByRef dv() As RPX_DD, ByRef rhs() As Double, ByRef answer() As RPX_DD) As Boolean
    Dim phase As Long, first As Long, last As Long, code As Long
    If Not RPX_NativeAvailable() Then Exit Function
#If Win64 Then
    SolveBounds n, cp, ri, rhs
    If LBound(lv) <> 0 Or UBound(lv) < cp(n) - 1 Or LBound(dv) <> 0 Or UBound(dv) < n - 1 Or LBound(answer) <> 0 Or UBound(answer) < n - 1 Then err.Raise 9, "RPX_Native", "NATIVE_DD_SOLVE_BUFFER_SHORT"
    solveCalls = solveCalls + 1
    For phase = 0 To 2
        For first = 0 To n - 1 Step 128
            Call Poll: last = first + 128: If last > n Then last = n
            code = NativeDDBacksolve(n, phase, first, last, VarPtr(cp(0)), VarPtr(ri(0)), VarPtr(lv(0)), VarPtr(dv(0)), VarPtr(rhs(0)), VarPtr(answer(0)))
            RPX_NativeCheckCode code
        Next first
    Next phase
    last = n
    Do While last > 0
        Call Poll: first = last - 128: If first < 0 Then first = 0
        code = NativeDDBacksolve(n, 3, first, last, VarPtr(cp(0)), VarPtr(ri(0)), VarPtr(lv(0)), VarPtr(dv(0)), VarPtr(rhs(0)), VarPtr(answer(0)))
        RPX_NativeCheckCode code: last = first
    Loop
    RPX_NativeDDBacksolve = True
#End If
End Function
