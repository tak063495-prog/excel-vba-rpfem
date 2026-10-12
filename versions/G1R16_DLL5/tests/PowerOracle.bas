Attribute VB_Name = "PowerOracle"
Option Explicit
Public Function DumpPowerOracle(ByVal inputPath As String, ByVal outputPath As String) As String
    Dim f As Integer, g As Integer, n As Long, k As Long, p As Long, i As Long, j As Long, eno As Long
    Dim s(0 To 2) As Double, z(0 To 2) As Double, a(0 To 2) As Double, out(0 To 2) As Double, m(0 To 2, 0 To 2) As Double
    Dim pw As Variant, side As Long, radius As Double
    f = FreeFile: Open inputPath For Binary Access Read As #f
    g = FreeFile: Open outputPath For Binary Access Write As #g
    Get #f, , n: Put #g, , n
    For k = 0 To n - 1
        Get #f, , p
        For i = 0 To 2: Get #f, , s(i): Next i
        For i = 0 To 2: Get #f, , z(i): Next i
        For side = 0 To 1
            For i = 0 To 2: If side = 0 Then a(i) = s(i) Else a(i) = z(i)
            Next i
            radius = Sqr(a(1) * a(1) + a(2) * a(2)): Put #g, , radius
            For Each pw In Array(0.5, -0.5, -1#)
                Erase out
                On Error Resume Next
                RPX_Power a, CDbl(pw), out
                eno = err.Number: err.Clear
                On Error GoTo 0
                Put #g, , eno
                For i = 0 To 2: Put #g, , out(i): Next i
            Next pw
            RPX_Quadratic a, m
            For j = 0 To 2: For i = 0 To 2: Put #g, , m(i, j): Next i: Next j
        Next side
    Next k
    Close #f: Close #g
    DumpPowerOracle = "PASS power oracle rows=" & n
End Function
