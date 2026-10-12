Attribute VB_Name = "BulkOracle"
Option Explicit
Private Sub Fixture(ByVal seed As Long)
    Dim n As Long, k As Long, i As Long, r As Long, a As Object, b As Object, c As Object, p As Long, margin As Double
    n = 4 + seed Mod 17: RPX_Begin n
    For k = 0 To 13 + seed Mod 7
        Set a = RPX_Row(): Set b = RPX_Row(): Set c = RPX_Row()
        For i = 0 To n - 1
            If (i + k * 3) Mod 5 <> 1 Then
                RPX_Add a, i, Sin(i * 3 + k * 7 + seed + 0.2)
                RPX_Add b, i, Cos(i * 7 + k * 5 + seed + 0.6)
                RPX_Add c, i, Sin(i * 5 + k * 2 + seed + 0.8)
            End If
        Next i
        If k Mod 3 = 0 Then RPX_LinearCone a Else RPX_Cone3 a, b, c, 0.17
    Next k
    RPX_Equal RPX_Unit(0), 0.3
    RPX_Equal RPX_Unit(n - 1), -0.12
    RPX_Compile
    RPX_ResetPreparedState
    ReDim XSlack(0 To XS - 1): ReDim XDual(0 To XS - 1)
    For k = 0 To XB - 1
        p = XC(k).first: XSlack(p) = 1.7: XDual(p) = 1.2
        If XC(k).dimn = 3 Then
            margin = Choose(1 + (k + seed) Mod 4, 0.01, 0.00001, 0.0000001, 0.00000003)
            XSlack(p + 1) = 1.7 * (1# - margin) * Cos(k + 0.4)
            XSlack(p + 2) = 1.7 * (1# - margin) * Sin(k + 0.4)
            XDual(p + 1) = 1.2 * (1# - margin) * Cos(k + 1.3)
            XDual(p + 2) = 1.2 * (1# - margin) * Sin(k + 1.3)
        End If
    Next k
End Sub
Public Function DumpBulkOracle(ByVal path As String, ByVal mode As String) As String
    Dim f As Long, seed As Long, precise As Variant, unshifted As Variant, i As Long, g() As Double, x() As Double, gt() As Double
    Dim before As String
    RPX_NativeSetMode mode
    On Error GoTo Failed
    f = FreeFile: Open path For Binary Access Write As #f
    For seed = 1 To 40
        Fixture seed
        ReDim x(0 To XN - 1): ReDim gt(0 To XN - 1)
        For i = 0 To XN - 1: x(i) = Sin(i * 3 + 0.5): gt(i) = Cos(i + 0.7): Next i
        RPX_GProduct x, g, True
        For i = 0 To XS - 1: Put #f, , g(i): Next i
        RPX_GTAdd g, gt, -0.37
        For i = 0 To XN - 1: Put #f, , gt(i): Next i
        For Each precise In Array(False, True)
            For Each unshifted In Array(False, True)
                DLL3TestSolveDump f, CBool(precise), CBool(unshifted)
            Next unshifted
        Next precise
        ' Same dimensions, new coefficients: explicitly invalidate packed data.
        XC(0).coef(0, 0) = XC(0).coef(0, 0) + 0.123
        Call RPX_NativeBulkReset: RPX_GProduct x, g, False
        For i = 0 To XS - 1: Put #f, , g(i): Next i
    Next seed
    Close #f
    DumpBulkOracle = "rows=160;" & RPX_NativeInfo()
    Exit Function
Failed:
    DumpBulkOracle = "FAIL " & err.number & " " & err.description & ";seed=" & seed
    Close #f
End Function
