Attribute VB_Name = "NTOracle"
Option Explicit
Public Function DumpNTOracle(ByVal inputPath As String, ByVal outputPath As String) As String
    Dim f As Integer, g As Integer, n As Long, k As Long, p As Long, i As Long, j As Long, eno As Long
    Dim s(0 To 2) As Double, z(0 To 2) As Double, h(0 To 2, 0 To 2) As Double
    Dim wm(0 To 2, 0 To 2) As Double, wi(0 To 2, 0 To 2) As Double, si(0 To 2) As Double, v(0 To 2) As Double
    f = FreeFile: Open inputPath For Binary Access Read As #f
    g = FreeFile: Open outputPath For Binary Access Write As #g
    Get #f, , n: Put #g, , n
    For k = 0 To n - 1
        Get #f, , p
        For i = 0 To 2: Get #f, , s(i): Next i
        For i = 0 To 2: Get #f, , z(i): Next i
        Erase h: Erase wm: Erase wi: Erase si: Erase v
        For i = 0 To 2
            For j = 0 To 2: RPX_NTPreciseH(i, j) = HDD(0#): Next j
        Next i
        On Error Resume Next
        If p = 0 Then RPX_NTBlock s, z, h, wm, wi, si, v Else RPX_NTBlockPrecise s, z, h, wm, wi, si, v
        eno = err.Number: err.Clear
        On Error GoTo 0
        Put #g, , eno
        For j = 0 To 2: For i = 0 To 2: Put #g, , h(i, j): Next i: Next j
        For j = 0 To 2: For i = 0 To 2: Put #g, , wm(i, j): Next i: Next j
        For j = 0 To 2: For i = 0 To 2: Put #g, , wi(i, j): Next i: Next j
        For i = 0 To 2: Put #g, , si(i): Next i
        For i = 0 To 2: Put #g, , v(i): Next i
        For j = 0 To 2: For i = 0 To 2: Put #g, , RPX_NTPreciseH(i, j): Next i: Next j
    Next k
    Close #f: Close #g
    DumpNTOracle = "PASS VBA NT oracle rows=" & n
End Function
