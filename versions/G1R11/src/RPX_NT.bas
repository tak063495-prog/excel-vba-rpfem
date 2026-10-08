Attribute VB_Name = "RPX_NT"
Option Explicit
' Lorentz algebra used by the Nesterov-Todd predictor-corrector direction.
Public Sub RPX_Power(ByRef a() As Double, ByVal power As Double, ByRef answer() As Double)
    Dim radius As Double, plus As Double, minus As Double, factor As Double
    radius = Sqr(a(1) * a(1) + a(2) * a(2))
    If a(0) <= radius Then err.Raise 5, , "NT_NONINTERIOR_POINT"
    plus = (a(0) + radius) ^ power: minus = (a(0) - radius) ^ power
    answer(0) = (plus + minus) / 2#: answer(1) = 0#: answer(2) = 0#
    If radius > 1E-30 Then
        factor = (plus - minus) / (2# * radius): answer(1) = a(1) * factor: answer(2) = a(2) * factor
    End If
End Sub
Public Sub RPX_Quadratic(ByRef a() As Double, ByRef matrix() As Double)
    Dim i As Long, j As Long, det As Double, radius As Double
    radius = Sqr(a(1) * a(1) + a(2) * a(2)): det = (a(0) - radius) * (a(0) + radius)
    For i = 0 To 2
        For j = 0 To 2: matrix(i, j) = 2# * a(i) * a(j): Next j
        If i = 0 Then matrix(i, i) = matrix(i, i) - det Else matrix(i, i) = matrix(i, i) + det
    Next i
End Sub
Public Sub RPX_M3V(ByRef matrix() As Double, ByRef a() As Double, ByRef answer() As Double)
    Dim i As Long, j As Long
    For i = 0 To 2
        answer(i) = 0#
        For j = 0 To 2: answer(i) = answer(i) + matrix(i, j) * a(j): Next j
    Next i
End Sub
Public Sub RPX_NTBlock(ByRef s() As Double, ByRef z() As Double, ByRef h() As Double, ByRef wmat() As Double, ByRef winv() As Double, ByRef sinv() As Double, ByRef v() As Double)
    Dim a(0 To 2) As Double, b(0 To 2) As Double, w(0 To 2) As Double, qs(0 To 2, 0 To 2) As Double
    RPX_Power s, 0.5, a: RPX_Quadratic a, qs: RPX_M3V qs, z, b
    RPX_Power b, -0.5, a: RPX_M3V qs, a, w
    RPX_Power w, -1#, a: RPX_Quadratic a, h
    RPX_Power w, 0.5, a: RPX_Quadratic a, wmat
    RPX_Power w, -0.5, a: RPX_Quadratic a, winv
    RPX_Power s, -1#, sinv: RPX_M3V winv, s, v
End Sub
Public Sub RPX_NTCorrection(ByRef v() As Double, ByRef wmat() As Double, ByRef winv() As Double, ByRef ds() As Double, ByRef dz() As Double, ByRef answer() As Double)
    Dim a(0 To 2) As Double, b(0 To 2) As Double, c(0 To 2) As Double, d(0 To 2) As Double, det As Double
    RPX_M3V winv, ds, a: RPX_M3V wmat, dz, b
    c(0) = a(0) * b(0) + a(1) * b(1) + a(2) * b(2)
    c(1) = a(0) * b(1) + b(0) * a(1): c(2) = a(0) * b(2) + b(0) * a(2)
    det = v(0) * v(0) - v(1) * v(1) - v(2) * v(2)
    If det <= 0# Then err.Raise 5, , "NT_CORRECTOR_SINGULAR"
    d(0) = (v(0) * c(0) - v(1) * c(1) - v(2) * c(2)) / det
    d(1) = (c(1) - v(1) * d(0)) / v(0): d(2) = (c(2) - v(2) * d(0)) / v(0)
    RPX_M3V winv, d, answer
End Sub
