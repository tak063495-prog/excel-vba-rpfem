Attribute VB_Name = "RPX_NTPrecision"
Option Explicit
' Same Lorentz/NT algebra, evaluated with two-component arithmetic.
Public RPX_NTPreciseH(0 To 2, 0 To 2) As RPX_DD
Private Function NTSqrt(ByRef a As RPX_DD) As RPX_DD
    Dim q As RPX_DD, t As RPX_DD, i As Long, value As Double
    value = HDValue(a)
    If value < 0# Then err.Raise 5, "RPX_NTPrecision", "NT_NONINTERIOR_POINT"
    If value = 0# Then Exit Function
    q = HDD(Sqr(value))
    For i = 1 To 2
        t = HDDiv(a, q): t = HDAdd(q, t): q = HDMul(HDD(0.5), t)
    Next i
    NTSqrt = q
End Function
Private Function NTDet(ByRef a() As RPX_DD) As RPX_DD
    Dim d As RPX_DD
    d = HDMul(a(0), a(0)): d = HDSub(d, HDMul(a(1), a(1))): NTDet = HDSub(d, HDMul(a(2), a(2)))
End Function
Private Sub NTPower(ByRef a() As RPX_DD, ByVal power As Double, ByRef answer() As RPX_DD)
    Dim rad As RPX_DD, plus As RPX_DD, minus As RPX_DD, determinant As RPX_DD, factor As RPX_DD
    determinant = NTDet(a)
    If HDValue(a(0)) <= 0# Or HDValue(determinant) <= 0# Then err.Raise 5, "RPX_NTPrecision", "NT_NONINTERIOR_POINT"
    rad = HDAdd(HDMul(a(1), a(1)), HDMul(a(2), a(2))): rad = NTSqrt(rad)
    plus = HDAdd(a(0), rad): minus = HDDiv(determinant, plus)
    If power = 0.5 Then
        plus = NTSqrt(plus): minus = NTSqrt(minus)
    ElseIf power = -0.5 Then
        plus = HDDiv(HDD(1#), NTSqrt(plus)): minus = HDDiv(HDD(1#), NTSqrt(minus))
    ElseIf power = -1# Then
        plus = HDDiv(HDD(1#), plus): minus = HDDiv(HDD(1#), minus)
    Else
        err.Raise 5, "RPX_NTPrecision", "NT_POWER_UNSUPPORTED"
    End If
    answer(0) = HDMul(HDD(0.5), HDAdd(plus, minus)): answer(1) = HDD(0#): answer(2) = HDD(0#)
    If HDValue(rad) > 1E-30 Then
        factor = HDDiv(HDSub(plus, minus), HDMul(HDD(2#), rad))
        answer(1) = HDMul(a(1), factor): answer(2) = HDMul(a(2), factor)
    End If
End Sub
Private Sub NTQuadratic(ByRef a() As RPX_DD, ByRef matrix() As RPX_DD)
    Dim i As Long, j As Long, d As RPX_DD
    d = NTDet(a)
    For i = 0 To 2
        For j = 0 To 2: matrix(i, j) = HDMul(HDD(2#), HDMul(a(i), a(j))): Next j
        If i = 0 Then matrix(i, i) = HDSub(matrix(i, i), d) Else matrix(i, i) = HDAdd(matrix(i, i), d)
    Next i
End Sub
Private Sub NTMV(ByRef matrix() As RPX_DD, ByRef a() As RPX_DD, ByRef answer() As RPX_DD)
    Dim i As Long, j As Long, value As RPX_DD
    For i = 0 To 2
        value = HDD(0#)
        For j = 0 To 2: value = HDAdd(value, HDMul(matrix(i, j), a(j))): Next j
        answer(i) = value
    Next i
End Sub
Public Sub RPX_NTBlockPrecise(ByRef s() As Double, ByRef z() As Double, ByRef h() As Double, ByRef wmat() As Double, ByRef winv() As Double, ByRef sinv() As Double, ByRef v() As Double)
    Dim sd(0 To 2) As RPX_DD, zd(0 To 2) As RPX_DD, a(0 To 2) As RPX_DD, b(0 To 2) As RPX_DD, w(0 To 2) As RPX_DD
    Dim qs(0 To 2, 0 To 2) As RPX_DD, mat(0 To 2, 0 To 2) As RPX_DD, i As Long, j As Long
    For i = 0 To 2: sd(i) = HDD(s(i)): zd(i) = HDD(z(i)): Next i
    NTPower sd, 0.5, a: NTQuadratic a, qs: NTMV qs, zd, b
    NTPower b, -0.5, a: NTMV qs, a, w
    NTPower w, -1#, a: NTQuadratic a, mat
    For i = 0 To 2
        For j = 0 To 2: RPX_NTPreciseH(i, j) = mat(i, j): h(i, j) = HDValue(mat(i, j)): Next j
    Next i
    NTPower w, 0.5, a: NTQuadratic a, mat
    For i = 0 To 2
        For j = 0 To 2: wmat(i, j) = HDValue(mat(i, j)): Next j
    Next i
    NTPower w, -0.5, a: NTQuadratic a, mat
    For i = 0 To 2
        For j = 0 To 2: winv(i, j) = HDValue(mat(i, j)): Next j
    Next i
    NTMV mat, sd, b: NTPower sd, -1#, a
    For i = 0 To 2: sinv(i) = HDValue(a(i)): v(i) = HDValue(b(i)): Next i
End Sub
