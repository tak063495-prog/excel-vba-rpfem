Attribute VB_Name = "NativeTests"
Option Explicit
Private Sub OrderingGraph(ByVal n As Long, ByVal kind As Long, ByRef rows() As Object)
    Dim i As Long, j As Long, first As Long, last As Long, salt As Long
    ReDim rows(0 To n - 1): salt = kind \ 5
    For i = 0 To n - 1
        Set rows(i) = RPX_Row()
        Select Case kind Mod 5
        Case 0 ' Empty and disconnected graph, equal degree priorities.
            If i Mod 3 = 0 Then rows(i)(i) = 1#
        Case 1 ' Dense graph, alternating insertion order.
            If i Mod 2 = 0 Then
                For j = n - 1 To 0 Step -1: rows(i)(j) = 1#: Next j
            Else
                For j = 0 To n - 1: rows(i)(j) = 1#: Next j
            End If
        Case 2 ' Banded non-symmetric input and duplicate assignment.
            first = i - 3: If first < 0 Then first = 0
            last = i + 2: If last >= n Then last = n - 1
            For j = last To first Step -1: rows(i)(j) = 1#: rows(i)(j) = 2#: Next j
        Case 3 ' Deterministic irregular graph with repeated degrees.
            For j = n - 1 To 0 Step -1
                If (i * 13 + j * 17 + salt * 43) Mod 31 < 5 Then rows(i)(j) = 1#
            Next j
        Case 4 ' Exercise n>20000 using two connected subgraphs and isolates.
            If i < 128 Or i >= n - 128 Then
                first = i - 2: If first < 0 Then first = 0
                last = i + 2: If last >= n Then last = n - 1
                For j = first To last: rows(i)(j) = 1#: Next j
            End If
        End Select
        If rows(i).Exists(i) And i Mod 2 = 0 Then
            rows(i).Remove i: rows(i)(i) = 1# ' Reinsert appends to .keys.
        End If
    Next i
End Sub
Private Sub CheckOrdering(ByVal n As Long, ByVal kind As Long, ByVal normal As Boolean)
    Dim rows() As Object, oldP() As Long, oldI() As Long, newP() As Long, newI() As Long, i As Long
    OrderingGraph n, kind, rows
    RPX_NativeSetMode "vba"
    If normal Then RPX_ColumnOrdering n, rows, oldP, oldI Else RPX_Ordering n, rows, oldP, oldI
    RPX_NativeSetMode "auto": Need RPX_NativeAvailable(), "Ordering DLL absent"
    If normal Then RPX_ColumnOrdering n, rows, newP, newI Else RPX_Ordering n, rows, newP, newI
    For i = 0 To n - 1
        Need oldP(i) = newP(i) And oldI(i) = newI(i), "Ordering mismatch n=" & n & ";kind=" & kind & ";normal=" & normal & ";i=" & i
        Need newI(newP(i)) = i, "Ordering inverse invalid"
        RPX_ReleaseRow rows(i)
    Next i
    If normal Then
        Need InStr(RPX_NativeInfo(), ";column_ordering=1") > 0, "Column ordering native hook unused"
    Else
        Need InStr(RPX_NativeInfo(), ";ordering=1") > 0, "Ordering native hook unused"
    End If
End Sub
Public Function TestNativeOrdering() As String
    Dim size As Variant, kind As Long
    On Error GoTo Failed
    For Each size In Array(1, 2, 7, 31, 70, 270)
        For kind = 0 To 3
            CheckOrdering CLng(size), kind, False: CheckOrdering CLng(size), kind, True
        Next kind
    Next size
    For kind = 8 To 48 Step 5
        CheckOrdering 97, kind, False: CheckOrdering 97, kind, True
    Next kind
    CheckOrdering 20481, 4, False: CheckOrdering 20481, 4, True
    RPX_NativeSetMode "auto"
    TestNativeOrdering = "PASS 68 exact permutation/inverse pairs; symmetric and implicit normal graph; insertion/reinsertion/ties/disconnected; n>20000"
    Exit Function
Failed:
    TestNativeOrdering = "FAIL " & err.Number & " " & err.Description
End Function
Public Function TestOrderingCancellation() As String
    Dim rows() As Object, p() As Long, inv() As Long, number As Long, message As String
    OrderingGraph 7, 2, rows: RPX_NativeSetMode "auto": XCancel = True
    On Error GoTo Expected
    RPX_Ordering 7, rows, p, inv
    XCancel = False: err.Raise 513, , "ordering cancellation ignored"
Expected:
    number = err.Number: message = err.Description: XCancel = False
    On Error GoTo 0
    Need number = 18 And message = "ANALYSIS_CANCELLED", "Ordering cancel mismatch"
    TestOrderingCancellation = "PASS ordering cancellation 18 preserved"
End Function
Public Function TestOrderingBadIndex() As String
    Dim rows() As Object, p() As Long, inv() As Long, number As Long
    OrderingGraph 7, 2, rows: rows(0)(7) = 1#: RPX_NativeSetMode "auto"
    On Error GoTo Expected
    RPX_Ordering 7, rows, p, inv
    err.Raise 513, , "bad ordering index accepted"
Expected:
    number = err.Number
    On Error GoTo 0
    Need number = 9, "Ordering subscript mismatch " & number
    TestOrderingBadIndex = "PASS ordering invalid index 9 preserved; native feed not called"
End Function
Private Sub Need(ByVal yes As Boolean, ByVal why As String)
    If Not yes Then err.Raise 513, "NativeTests", why
End Sub
Private Sub Same(ByVal a As Double, ByVal b As Double, ByVal why As String)
    Need RPX_ExactDouble(a) = RPX_ExactDouble(b), why & ";VBA=" & RPX_ExactDouble(a) & ";DLL=" & RPX_ExactDouble(b) & ";delta=" & CStr(a - b)
End Sub
Private Sub Matrix(ByVal n As Long, ByVal width As Long, ByVal indefinite As Boolean, ByRef ap() As Long, ByRef ai() As Long, ByRef av() As Double, ByRef rhs() As Double)
    Dim k As Long, j As Long, p As Long, first As Long, size As Long
    ReDim ap(0 To n): ReDim rhs(0 To n - 1)
    For k = 0 To n - 1
        first = k - width: If first < 0 Then first = 0
        size = size + k - first + 1: ap(k + 1) = size
    Next k
    ReDim ai(0 To size - 1): ReDim av(0 To size - 1)
    For k = 0 To n - 1
        first = k - width: If first < 0 Then first = 0
        For j = first To k
            ai(p) = j
            If j = k Then
                av(p) = 30# + width * 2#
                If indefinite And k Mod 2 = 1 Then av(p) = -av(p)
            Else
                av(p) = ((j * 13 + k * 7) Mod 17 - 8) / 19#
            End If
            p = p + 1
        Next j
        rhs(k) = (k Mod 31 - 15) / 7#
    Next k
End Sub
Private Sub TestMatrix(ByVal n As Long, ByVal width As Long, ByVal indefinite As Boolean)
    Dim ap() As Long, ai() As Long, av() As Double, rhs() As Double
    Dim old() As Double, newx() As Double, oldDD() As RPX_DD, newDD() As RPX_DD, values() As RPX_DD
    Dim i As Long, minimum As Double, accepted As Boolean
    Matrix n, width, indefinite, ap, ai, av, rhs
    RPX_NativeSetMode "vba": RPX_Symbolic n, ap, ai
    Need RPX_Numeric(ap, ai, av), "VBA numeric"
    minimum = RPX_PivotMinimum: RPX_Backsolve rhs, old
    RPX_NativeSetMode "auto": Need RPX_NativeAvailable(), "DLL missing in tests"
    RPX_Backsolve rhs, newx
    For i = 0 To n - 1: Same old(i), newx(i), "Double backsolve-only " & n & ":" & i: Next i
    Need RPX_Numeric(ap, ai, av), "DLL numeric"
    Same minimum, RPX_PivotMinimum, "Double pivot"
    RPX_Backsolve rhs, newx
    For i = 0 To n - 1: Same old(i), newx(i), "Double answer " & i: Next i
    ReDim values(0 To UBound(av))
    For i = 0 To UBound(av): values(i).hi = av(i): values(i).lo = (i Mod 7 - 3) * 1E-20: Next i
    RPX_NativeSetMode "vba": Need RPX_DDNumeric(ap, ai, values), "VBA DD numeric"
    minimum = RPX_PivotMinimum: RPX_DDBacksolve rhs, oldDD
    RPX_NativeSetMode "auto": Need RPX_DDNumeric(ap, ai, values), "DLL DD numeric"
    Same minimum, RPX_PivotMinimum, "DD pivot"
    RPX_DDBacksolve rhs, newDD
    For i = 0 To n - 1
        Same oldDD(i).hi, newDD(i).hi, "DD hi " & i
        Same oldDD(i).lo, newDD(i).lo, "DD lo " & i
    Next i
    RPX_SchurFactorBudget = 1000000000#
    RPX_NativeSetMode "vba": RPXC_Symbolic n, ap, ai
    Need RPXC_Numeric(ap, ai, av), "VBA Schur numeric": RPXC_Backsolve rhs, old
    minimum = RPXC_PivotMinimum
    RPX_NativeSetMode "auto": Need RPXC_Numeric(ap, ai, av), "DLL Schur numeric": RPXC_Backsolve rhs, newx
    Same minimum, RPXC_PivotMinimum, "Schur pivot"
    For i = 0 To n - 1: Same old(i), newx(i), "Schur answer " & i: Next i
End Sub
Public Function TestNativeKernels() As String
    On Error GoTo Failed
    TestMatrix 3, 2, True
    TestMatrix 3, 0, False
    TestMatrix 270, 9, False
    TestMatrix 270, 9, True
    TestMatrix 70, 69, False
    RPX_NativeSetMode "auto"
    TestNativeKernels = "PASS bitwise Double/DD/Schur; 3,70,270 rows; SPD/indefinite; chunk boundaries"
    Exit Function
Failed:
    TestNativeKernels = "FAIL " & err.Number & " " & err.Description
End Function
Public Function TestNativeCancellation() As String
    Dim ap() As Long, ai() As Long, av() As Double, rhs() As Double, ok As Boolean
    Dim number As Long, message As String
    Matrix 3, 2, False, ap, ai, av, rhs
    RPX_NativeSetMode "auto": RPX_Symbolic 3, ap, ai: XCancel = True
    On Error GoTo Expected
    ok = RPX_Numeric(ap, ai, av)
    XCancel = False: err.Raise 513, , "cancel not propagated"
Expected:
    number = err.Number: message = err.Description: XCancel = False
    On Error GoTo 0
    Need number = 18 And message = "ANALYSIS_CANCELLED", "cancel mismatch " & number & " " & message
    TestNativeCancellation = "PASS cancellation 18 preserved"
End Function
Public Function TestNativePivot() As String
    Dim ap(0 To 2) As Long, ai(0 To 2) As Long, av(0 To 2) As Double, vals(0 To 2) As RPX_DD
    Dim mode As Variant, ok As Boolean, minimum As Double
    ap(1) = 1: ap(2) = 3: ai(1) = 0: ai(2) = 1: av(1) = 1#: vals(1).hi = 1#
    For Each mode In Array("vba", "auto")
        RPX_NativeSetMode CStr(mode): RPX_Symbolic 2, ap, ai
        ok = RPX_Numeric(ap, ai, av): Need Not ok And RPX_PivotFailure = 0, "Double pivot rejection"
        ok = RPX_DDNumeric(ap, ai, vals): Need Not ok And RPX_PivotFailure = 0, "DD pivot rejection"
    Next mode
    TestNativePivot = "PASS singular pivot: VBA/DLL both reject at 0"
End Function
Public Function TestMissingDLL() As String
    Dim rows() As Object, oldP() As Long, oldI() As Long, newP() As Long, newI() As Long, i As Long
    RPX_NativeSetMode "auto"
    Need Not RPX_NativeAvailable(), "DLL unexpectedly present"
    Need InStr(RPX_NativeInfo(), "DLL_NOT_FOUND") > 0, "missing DLL reason"
    Need Left$(RPX_TestLDL(), 4) = "PASS", "VBA LDL fallback"
    RPX_SchurFactorBudget = 1000000000#
    Need Left$(RPXC_TestLDL(), 4) = "PASS", "VBA Schur fallback"
    OrderingGraph 31, 3, rows
    RPX_NativeSetMode "vba": RPX_Ordering 31, rows, oldP, oldI
    RPX_NativeSetMode "auto": RPX_Ordering 31, rows, newP, newI
    For i = 0 To 30: Need oldP(i) = newP(i) And oldI(i) = newI(i), "Missing DLL symmetric order fallback": Next i
    RPX_NativeSetMode "vba": RPX_ColumnOrdering 31, rows, oldP, oldI
    RPX_NativeSetMode "auto": RPX_ColumnOrdering 31, rows, newP, newI
    For i = 0 To 30: Need oldP(i) = newP(i) And oldI(i) = newI(i), "Missing DLL normal order fallback": Next i
    TestMissingDLL = "PASS automatic missing-DLL VBA fallback; numeric LDL/Schur and both ordering paths"
End Function
Public Function TestNativeShortBuffer() As String
    Dim ap() As Long, ai() As Long, av() As Double, rhs() As Double, answer() As Double, ok As Boolean
    Dim number As Long, message As String
    Matrix 3, 2, False, ap, ai, av, rhs
    RPX_NativeSetMode "auto": RPX_Symbolic 3, ap, ai: ok = RPX_Numeric(ap, ai, av)
    ReDim rhs(0 To 0)
    On Error GoTo Expected
    RPX_Backsolve rhs, answer
    err.Raise 513, , "short buffer was accepted"
Expected:
    number = err.Number: message = err.Description
    On Error GoTo 0
    Need number = 9 And message = "NATIVE_SOLVE_BUFFER_SHORT", "buffer mismatch " & number & " " & message
    TestNativeShortBuffer = "PASS subscript error 9 before native pointer access"
End Function
Private Function OverflowNumber(ByVal mode As String) As Long
    Dim ap(0 To 2) As Long, ai(0 To 2) As Long, av(0 To 2) As Double, ok As Boolean
    ap(1) = 1: ap(2) = 3: ai(1) = 0: ai(2) = 1
    av(0) = 1E-10: av(1) = 1E+308: av(2) = 1E+308
    RPX_NativeSetMode mode: RPX_Symbolic 2, ap, ai
    On Error GoTo Failed
    ok = RPX_Numeric(ap, ai, av): Exit Function
Failed:
    OverflowNumber = err.Number
End Function
Public Function TestNativeOverflow() As String
    On Error GoTo Failed
    Need OverflowNumber("vba") = 6, "VBA overflow number"
    Need OverflowNumber("auto") = 6, "DLL overflow number"
    TestNativeOverflow = "PASS VBA/DLL both propagate overflow 6; no successful factor"
    Exit Function
Failed:
    TestNativeOverflow = "FAIL " & err.Number & " " & err.Description
End Function
Public Function BenchmarkEvidence(ByVal path As String) As String
    Dim f As Integer
    f = FreeFile: Open path For Binary Access Write As #f
    Put #f, , rn: Put #f, , re: Put #f, , RSubdivisions
    Put #f, , RBestLower.field: Put #f, , RBestUpper.field
    Close #f
    BenchmarkEvidence = "PASS fields exported"
End Function
