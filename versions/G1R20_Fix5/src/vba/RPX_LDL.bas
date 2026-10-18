Attribute VB_Name = "RPX_LDL"
Option Explicit
#If Win64 Then
Private Declare PtrSafe Function NRCreate Lib "rpfem_native_x64.dll" Alias "rpx_refine_create" (ByVal n As Long, ByRef handle As LongPtr) As Long
Private Declare PtrSafe Function NRBegin Lib "rpfem_native_x64.dll" Alias "rpx_refine_begin" (ByVal handle As LongPtr, ByVal n As Long, ByVal epoch As LongLong, ByVal rhs As LongPtr) As Long
Private Declare PtrSafe Function NRStep Lib "rpfem_native_x64.dll" Alias "rpx_refine_step" (ByVal handle As LongPtr, ByVal n As Long, ByVal epoch As LongLong, ByVal args As LongPtr, ByVal milliseconds As Long, ByRef done As Long) As Long
Private Declare PtrSafe Function NRResult Lib "rpfem_native_x64.dll" Alias "rpx_refine_result" (ByVal handle As LongPtr, ByVal n As Long, ByVal epoch As LongLong, ByVal answer As LongPtr, ByRef checks As Long, ByRef corrections As Long, ByRef residual As Double) As Long
Private Declare PtrSafe Function NRDestroy Lib "rpfem_native_x64.dll" Alias "rpx_refine_destroy" (ByVal handle As LongPtr) As Long
Private refineHandle As LongPtr, refineEpoch As LongLong
#End If
Private refineN As Long, nativeRefineSolves As Long, nativeRefineSteps As Long
' Adapted from LDL, Copyright (c) 2005-2022 Timothy A. Davis.
' SPDX-License-Identifier: LGPL-2.1-or-later
' VBA adaptation: compressed-column arrays, checks, cancellation, test interface.
' See third_party/LDL_LICENSE_NOTICE.txt and LGPL-2.1.txt.
' Sparse symmetric LDL factorization, zero-based compressed columns.
' Elimination-tree formulation described by T. A. Davis, LDL/SuiteSparse.
' No external numerical library is called. Input contains upper triangle only.
' 20260913 speed: cache reach order and write slots for a fixed pattern.
' version=20260926_0545 engine=20260926_ldl_local
Private ddFactor() As RPX_DD, ddDiagonal() As RPX_DD, ddScatter() As RPX_DD, ddReady As Boolean
Private orderN As Long
Private colPtr() As Long, rowInd() As Long, factorVal() As Double
Private parent() As Long, countCol() As Long, diagonal() As Double
Private scatter() As Double, marks() As Long, reach() As Long, chain() As Long
Private reachPtr() As Long, reachNode() As Long, writeAt() As Long
Private patternReady As Boolean, rowPatternFilled As Boolean
Public RPX_FactorNNZ As Long, RPX_PivotMinimum As Double
Public RPX_PivotFailure As Long
Public RPX_ReachCacheBytes As Double

Public Sub RPX_LDLRelease()
    RPX_NativeRefineRelease
    ddReady = False: Erase ddFactor: Erase ddDiagonal: Erase ddScatter
    patternReady = False
    rowPatternFilled = False
    orderN = 0
    RPX_FactorNNZ = 0
    RPX_ReachCacheBytes = 0#
    Erase colPtr: Erase rowInd: Erase factorVal
    Erase parent: Erase countCol: Erase diagonal
    Erase scatter: Erase marks: Erase reach: Erase chain
    Erase reachPtr: Erase reachNode: Erase writeAt
End Sub

Public Sub RPX_Symbolic(ByVal n As Long, ByRef ap() As Long, ByRef ai() As Long)
    Dim k As Long, p As Long, j As Long
    ' Drop the previous model's factor arrays before the budget check.
    Call RPX_LDLRelease
    orderN = n
    patternReady = False
    rowPatternFilled = False
    ReDim parent(0 To n - 1): ReDim countCol(0 To n - 1)
    ReDim marks(0 To n - 1): ReDim colPtr(0 To n)
    For k = 0 To n - 1
        parent(k) = -1: marks(k) = k
        For p = ap(k) To ap(k + 1) - 1
            j = ai(p)
            If j < 0 Or j > k Then err.Raise 5, , "LDL_INVALID_UPPER_INDEX"
            Do While marks(j) <> k
                If parent(j) < 0 Then parent(j) = k
                countCol(j) = countCol(j) + 1: marks(j) = k
                j = parent(j)
            Loop
        Next p
    Next k
    For k = 0 To n - 1
        If CDbl(colPtr(k)) + countCol(k) > 100000000# Then err.Raise 7, , "LDL_FILL_MEMORY_LIMIT"
        colPtr(k + 1) = colPtr(k) + countCol(k)
    Next k
    RPX_FactorNNZ = colPtr(n)
    If CDbl(RPX_FactorNNZ) > 100000000# Then err.Raise 7, , "LDL_FILL_MEMORY_LIMIT"
    ' R3: check the allocation, not the retired 1.2 GiB private-byte cap.
    RPX_CapacityCheck "before_ldl_factor", 24# * CDbl(RPX_FactorNNZ) + 32# * CDbl(n)
    ReDim rowInd(0 To RPX_FactorNNZ): ReDim factorVal(0 To RPX_FactorNNZ)
    ReDim diagonal(0 To n - 1): ReDim scatter(0 To n - 1)
    ReDim reach(0 To n - 1): ReDim chain(0 To n - 1)
    ReDim reachPtr(0 To n)
    If RPX_FactorNNZ > 0 Then
        ReDim reachNode(0 To RPX_FactorNNZ - 1)
        ReDim writeAt(0 To RPX_FactorNNZ - 1)
    End If
    RPX_ReachCacheBytes = 8# * CDbl(RPX_FactorNNZ) + 4# * CDbl(n + 1)
    RPX_DiagBegin dpLDLReach
    Call BuildReachCache(ap, ai)
    RPX_DiagEnd dpLDLReach
    RPX_DiagEvent "ldl_reach_cache", "bytes=" & RPX_ReachCacheBytes & ";factor_nnz=" & RPX_FactorNNZ
    patternReady = True
End Sub

Private Sub BuildReachCache(ByRef ap() As Long, ByRef ai() As Long)
    Dim k As Long, p As Long, j As Long, top As Long, depth As Long, q As Long
    Dim pos As Long, stopAt As Long
    For k = 0 To orderN - 1
        marks(k) = -1: countCol(k) = 0
    Next k
    pos = 0
    For k = 0 To orderN - 1
        top = orderN: marks(k) = k
        reachPtr(k) = pos
        For p = ap(k) To ap(k + 1) - 1
            j = ai(p): depth = 0
            Do While marks(j) <> k
                chain(depth) = j: depth = depth + 1: marks(j) = k
                j = parent(j)
                If j < 0 Then err.Raise 5, , "LDL_SYMBOLIC_MISMATCH"
            Loop
            Do While depth > 0
                depth = depth - 1: top = top - 1: reach(top) = chain(depth)
            Loop
        Next p
        For q = top To orderN - 1
            j = reach(q)
            stopAt = colPtr(j) + countCol(j)
            reachNode(pos) = j
            writeAt(pos) = stopAt
            rowInd(stopAt) = k
            countCol(j) = countCol(j) + 1
            pos = pos + 1
        Next q
    Next k
    reachPtr(orderN) = pos
    If pos <> RPX_FactorNNZ Then err.Raise 5, , "LDL_SYMBOLIC_MISMATCH"
    rowPatternFilled = True
End Sub

Public Function RPX_Numeric(ByRef ap() As Long, ByRef ai() As Long, ByRef av() As Double) As Boolean
    Dim k As Long, p As Long, j As Long, q As Long, stopAt As Long, qEnd As Long, r As Long
    Dim value As Double, lij As Double, dk As Double
    Dim useCache As Boolean
    Dim nativeAccepted As Boolean
    ddReady = False
    useCache = patternReady
    For k = 0 To orderN - 1
        scatter(k) = 0#
    Next k
    RPX_PivotMinimum = 1E+300: RPX_PivotFailure = -1
    If Not useCache Then
        RPX_Numeric = NumericWalk(ap, ai, av)
        Exit Function
    End If
    If RPX_NativeNumeric(orderN, ap, ai, av, colPtr, rowInd, reachPtr, reachNode, writeAt, factorVal, diagonal, scatter, RPX_PivotMinimum, RPX_PivotFailure, nativeAccepted) Then
        RPX_Numeric = nativeAccepted: Exit Function
    End If
    For k = 0 To orderN - 1
        If k Mod 4096 = 0 Then
            DoEvents
            If XCancel Then err.Raise 18, , "ANALYSIS_CANCELLED"
        End If
        For p = ap(k) To ap(k + 1) - 1
            scatter(ai(p)) = scatter(ai(p)) + av(p)
        Next p
        dk = scatter(k): scatter(k) = 0#
        qEnd = reachPtr(k + 1) - 1
        For q = reachPtr(k) To qEnd
            j = reachNode(q): value = scatter(j): scatter(j) = 0#
            stopAt = writeAt(q)
            For p = colPtr(j) To stopAt - 1
                r = rowInd(p): scatter(r) = scatter(r) - factorVal(p) * value
            Next p
            If Abs(diagonal(j)) < 1E-290 Then RPX_PivotFailure = j: Exit Function
            lij = value / diagonal(j): dk = dk - value * lij
            factorVal(stopAt) = lij
        Next q
        If Abs(dk) < 1E-290 Then RPX_PivotFailure = k: Exit Function
        diagonal(k) = dk
        If Abs(dk) < RPX_PivotMinimum Then RPX_PivotMinimum = Abs(dk)
    Next k
    RPX_Numeric = True
End Function

Private Function NumericWalk(ByRef ap() As Long, ByRef ai() As Long, ByRef av() As Double) As Boolean
    Dim k As Long, p As Long, j As Long, top As Long, depth As Long, q As Long, stopAt As Long
    Dim value As Double, lij As Double, dk As Double
    For k = 0 To orderN - 1
        scatter(k) = 0#: marks(k) = -1: countCol(k) = 0
    Next k
    For k = 0 To orderN - 1
        If k Mod 4096 = 0 Then
            DoEvents
            If XCancel Then err.Raise 18, , "ANALYSIS_CANCELLED"
        End If
        top = orderN: marks(k) = k
        For p = ap(k) To ap(k + 1) - 1
            j = ai(p): scatter(j) = scatter(j) + av(p): depth = 0
            Do While marks(j) <> k
                chain(depth) = j: depth = depth + 1: marks(j) = k
                j = parent(j)
                If j < 0 Then err.Raise 5, , "LDL_SYMBOLIC_MISMATCH"
            Loop
            Do While depth > 0
                depth = depth - 1: top = top - 1: reach(top) = chain(depth)
            Loop
        Next p
        dk = scatter(k): scatter(k) = 0#
        For q = top To orderN - 1
            j = reach(q): value = scatter(j): scatter(j) = 0#
            stopAt = colPtr(j) + countCol(j)
            For p = colPtr(j) To stopAt - 1
                scatter(rowInd(p)) = scatter(rowInd(p)) - factorVal(p) * value
            Next p
            If Abs(diagonal(j)) < 1E-290 Then RPX_PivotFailure = j: Exit Function
            lij = value / diagonal(j): dk = dk - value * lij
            rowInd(stopAt) = k: factorVal(stopAt) = lij: countCol(j) = countCol(j) + 1
        Next q
        If Abs(dk) < 1E-290 Then RPX_PivotFailure = k: Exit Function
        diagonal(k) = dk
        If Abs(dk) < RPX_PivotMinimum Then RPX_PivotMinimum = Abs(dk)
    Next k
    NumericWalk = True
End Function

Public Sub RPX_Backsolve(ByRef rhs() As Double, ByRef answer() As Double)
    Dim j As Long, p As Long, r As Long, value As Double
    ReDim answer(0 To orderN - 1)
    If RPX_NativeBacksolve(orderN, colPtr, rowInd, factorVal, diagonal, rhs, answer) Then Exit Sub
    For j = 0 To orderN - 1: answer(j) = rhs(j): Next j
    For j = 0 To orderN - 1
        value = answer(j)
        For p = colPtr(j) To colPtr(j + 1) - 1
            r = rowInd(p): answer(r) = answer(r) - factorVal(p) * value
        Next p
    Next j
    For j = 0 To orderN - 1: answer(j) = answer(j) / diagonal(j): Next j
    For j = orderN - 1 To 0 Step -1
        value = answer(j)
        For p = colPtr(j) To colPtr(j + 1) - 1
            value = value - factorVal(p) * answer(rowInd(p))
        Next p
        answer(j) = value
    Next j
End Sub

Public Sub RPX_SymmetricProduct(ByRef ap() As Long, ByRef ai() As Long, ByRef av() As Double, ByRef x() As Double, ByRef y() As Double)
    Dim k As Long, j As Long, p As Long
    ReDim y(0 To UBound(x))
#If Win64 Then
    If RPX_NativeAvailable() Then
        Dim nativeArgs(0 To 4) As LongPtr, nativeMetrics(0 To 0) As Double
        nativeArgs(0) = VarPtr(ap(0)): nativeArgs(1) = VarPtr(ai(0)): nativeArgs(2) = VarPtr(av(0)): nativeArgs(3) = VarPtr(x(0)): nativeArgs(4) = VarPtr(y(0))
        If RPX_NativeVectors(11, UBound(x) + 1, nativeArgs, nativeMetrics) Then Exit Sub
    End If
#End If
    For k = 0 To UBound(x)
        For p = ap(k) To ap(k + 1) - 1
            j = ai(p): y(j) = y(j) + av(p) * x(k)
            If j <> k Then y(k) = y(k) + av(p) * x(j)
        Next p
    Next k
End Sub

Public Function RPX_TestLDL() As String
    Dim ap(0 To 3) As Long, ai(0 To 5) As Long, av(0 To 5) As Double
    Dim rhs(0 To 2) As Double, x() As Double, ax() As Double, residual As Double, i As Long
    ap(1) = 1: ap(2) = 3: ap(3) = 6
    ai(0) = 0: ai(1) = 0: ai(2) = 1: ai(3) = 0: ai(4) = 1: ai(5) = 2
    av(0) = 4: av(1) = 1: av(2) = 3: av(3) = 1: av(4) = -1: av(5) = -0.1
    rhs(0) = 9: rhs(1) = 4: rhs(2) = -1.3
    RPX_Symbolic 3, ap, ai
    If Not RPX_Numeric(ap, ai, av) Then err.Raise 5, , "LDL_TEST_FACTOR_FAILED"
    RPX_Backsolve rhs, x: RPX_SymmetricProduct ap, ai, av, x, ax
    For i = 0 To 2: residual = residual + Abs(ax(i) - rhs(i)): Next i
    If residual > 0.000000000001 Or Abs(x(0) - 1) + Abs(x(1) - 2) + Abs(x(2) - 3) > 0.00000000001 Then err.Raise 5, , "LDL_TEST_FAILED"
    RPX_TestLDL = "PASS LDL; residual=" & CStr(residual)
End Function

' Same cached symbolic LDL algorithm, with two-component numeric arithmetic.
Public Function RPX_DDNumeric(ByRef ap() As Long, ByRef ai() As Long, ByRef av() As RPX_DD) As Boolean
    Dim k As Long, p As Long, j As Long, q As Long, stopAt As Long, r As Long
    Dim value As RPX_DD, lij As RPX_DD, dk As RPX_DD, term As RPX_DD, zero As RPX_DD
    Dim bh As Double, bl As Double
    Dim nativeAccepted As Boolean
    ddReady = False
    If Not patternReady Then err.Raise 5, , "DD_LDL_PATTERN_NOT_READY"
    RPX_CapacityCheck "dd_ldl_factor", 16# * (CDbl(RPX_FactorNNZ) + 2# * orderN + 1#)
    ReDim ddFactor(0 To RPX_FactorNNZ): ReDim ddDiagonal(0 To orderN - 1): ReDim ddScatter(0 To orderN - 1)
    RPX_PivotMinimum = 1E+300: RPX_PivotFailure = -1
    If RPX_NativeDDNumeric(orderN, ap, ai, av, colPtr, rowInd, reachPtr, reachNode, writeAt, ddFactor, ddDiagonal, ddScatter, RPX_PivotMinimum, RPX_PivotFailure, nativeAccepted) Then
        ddReady = nativeAccepted: RPX_DDNumeric = nativeAccepted: Exit Function
    End If
    For k = 0 To orderN - 1
        If k Mod 256 = 0 Then
            RPX_Stage "HSD DD LDL " & k & "/" & orderN
            DoEvents: If XCancel Then err.Raise 18, , "ANALYSIS_CANCELLED"
        End If
        For p = ap(k) To ap(k + 1) - 1
            j = ai(p): ddScatter(j) = HDAdd(ddScatter(j), av(p))
        Next p
        dk = ddScatter(k): ddScatter(k) = zero
        For q = reachPtr(k) To reachPtr(k + 1) - 1
            j = reachNode(q): value = ddScatter(j): ddScatter(j) = zero: stopAt = writeAt(q)
            HDSplit value.hi, bh, bl
            For p = colPtr(j) To stopAt - 1
                r = rowInd(p): HDMulSubKnown ddScatter(r), ddFactor(p), value, bh, bl
            Next p
            If Abs(HDValue(ddDiagonal(j))) < 1E-290 Then RPX_PivotFailure = j: Exit Function
            lij = HDDiv(value, ddDiagonal(j)): term = HDMul(value, lij): dk = HDSub(dk, term)
            ddFactor(stopAt) = lij
        Next q
        If Abs(HDValue(dk)) < 1E-290 Then RPX_PivotFailure = k: Exit Function
        ddDiagonal(k) = dk
        If Abs(HDValue(dk)) < RPX_PivotMinimum Then RPX_PivotMinimum = Abs(HDValue(dk))
    Next k
    ddReady = True: RPX_DDNumeric = True
End Function
Public Sub RPX_DDBacksolve(ByRef rhs() As Double, ByRef answer() As RPX_DD)
    Dim j As Long, p As Long, r As Long, value As RPX_DD, term As RPX_DD
    Dim bh As Double, bl As Double
    If Not ddReady Then err.Raise 5, , "DD_LDL_FACTOR_NOT_READY"
    ReDim answer(0 To orderN - 1)
    If RPX_NativeDDBacksolve(orderN, colPtr, rowInd, ddFactor, ddDiagonal, rhs, answer) Then Exit Sub
    For j = 0 To orderN - 1: answer(j) = HDD(rhs(j)): Next j
    For j = 0 To orderN - 1
        If j Mod 4096 = 0 Then
            DoEvents: If XCancel Then err.Raise 18, , "ANALYSIS_CANCELLED"
        End If
        value = answer(j): HDSplit value.hi, bh, bl
        For p = colPtr(j) To colPtr(j + 1) - 1
            r = rowInd(p): HDMulSubKnown answer(r), ddFactor(p), value, bh, bl
        Next p
    Next j
    For j = 0 To orderN - 1: answer(j) = HDDiv(answer(j), ddDiagonal(j)): Next j
    For j = orderN - 1 To 0 Step -1
        value = answer(j)
        For p = colPtr(j) To colPtr(j + 1) - 1
            term = HDMul(ddFactor(p), answer(rowInd(p))): value = HDSub(value, term)
        Next p
        answer(j) = value
    Next j
End Sub

Public Sub RPX_NativeRefineRelease()
#If Win64 Then
    Dim code As Long
    If refineHandle <> 0 Then
        code = NRDestroy(refineHandle): refineHandle = 0: refineN = 0
        RPX_NativeCheckCode code
    End If
#End If
End Sub
Public Function RPX_NativeRefineInfo() As String
    RPX_NativeRefineInfo = ";refine_jobs=" & nativeRefineSolves & ";refine_steps=" & nativeRefineSteps
End Function
Public Sub RPX_NativeRefineResetCounts()
    nativeRefineSolves = 0: nativeRefineSteps = 0
End Sub
Public Function RPX_CheckedNativeSolve(ByRef rhs() As Double, ByRef answer() As Double, ByRef checks As Long, ByRef corrections As Long, ByRef residual As Double) As Boolean
    If Not RPX_NativeAvailable() Then Exit Function
#If Win64 Then
    Dim args(0 To 6) As LongPtr, done As Long, eno As Long, es As String, et As String, ignored As Long
    On Error GoTo Failed
    If LBound(rhs) <> 0 Or UBound(rhs) <> orderN - 1 Then err.Raise 9, , "NATIVE_REFINE_RHS"
    If refineHandle <> 0 And refineN <> orderN Then RPX_NativeRefineRelease
    If refineHandle = 0 Then
        RPX_CapacityCheck "native_refine", 32# * orderN
        RPX_NativeCheckCode NRCreate(orderN, refineHandle): refineN = orderN
    End If
    refineEpoch = refineEpoch + 1
    args(0) = VarPtr(XKPtr(0)): args(1) = VarPtr(XKId(0)): args(2) = VarPtr(XKVal(0))
    args(3) = VarPtr(colPtr(0)): args(4) = VarPtr(rowInd(0)): args(5) = VarPtr(factorVal(0)): args(6) = VarPtr(diagonal(0))
    RPX_NativeCheckCode NRBegin(refineHandle, orderN, refineEpoch, VarPtr(rhs(0)))
    nativeRefineSolves = nativeRefineSolves + 1
    Do While done = 0
        RPX_NativePulse
        RPX_NativeCheckCode NRStep(refineHandle, orderN, refineEpoch, VarPtr(args(0)), 50, done)
        nativeRefineSteps = nativeRefineSteps + 1
    Loop
    ReDim answer(0 To orderN - 1)
    RPX_NativeCheckCode NRResult(refineHandle, orderN, refineEpoch, VarPtr(answer(0)), checks, corrections, residual)
    RPX_CheckedNativeSolve = True
    Exit Function
Failed:
    eno = err.number: es = err.source: et = err.description
    If refineHandle <> 0 Then ignored = NRDestroy(refineHandle): refineHandle = 0: refineN = 0
    err.Raise eno, es, et
#End If
End Function

#If Win64 Then
Public Sub RPX_NativeLDLFactorArgs(ByVal precise As Boolean, ByRef args() As LongPtr)
    If orderN <> XN + XM Then err.Raise 5, , "HSD_FACTOR_CONTEXT_MISMATCH"
    args(0) = VarPtr(colPtr(0)): args(1) = VarPtr(rowInd(0))
    If precise Then
        If Not ddReady Then err.Raise 5, , "HSD_FACTOR_CONTEXT_MISMATCH"
        args(2) = VarPtr(ddFactor(0)): args(3) = VarPtr(ddDiagonal(0))
    Else
        args(2) = VarPtr(factorVal(0)): args(3) = VarPtr(diagonal(0))
    End If
End Sub
#End If

Public Function RPX_NativeRefineBytes() As Double
#If Win64 Then
    If refineHandle <> 0 Then RPX_NativeRefineBytes = 32# * refineN + 4096#
#End If
End Function
