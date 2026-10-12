Attribute VB_Name = "BoundaryOracle"
Option Explicit
Public Function DumpBoundary(ByVal source As String, ByVal target As String, ByVal mode As String) As String
    Dim f As Integer, g As Integer, n As Long, i As Long, j As Long, dimn As Long, eno As Long, hit As Double, cone As Long, ok As Boolean
    Dim x() As Double, d() As Double
    RPX_NativeSetMode mode
    f = FreeFile: Open source For Binary Access Read As #f
    g = FreeFile: Open target For Binary Access Write As #g
    Get #f, , n: Put #g, , n
    For i = 0 To n - 1
        Get #f, , dimn
        RPX_Begin 3
        If dimn = 3 Then RPX_Cone3 RPX_Unit(0), RPX_Unit(1), RPX_Unit(2) Else RPX_LinearCone RPX_Unit(0)
        ReDim x(0 To dimn - 1): ReDim d(0 To dimn - 1)
        For j = 0 To dimn - 1: Get #f, , x(j): Next j
        For j = 0 To dimn - 1: Get #f, , d(j): Next j
        On Error Resume Next
        ok = ConeBoundaryHit(x, d, hit, cone)
        eno = err.number: err.Clear
        On Error GoTo 0
        Put #g, , eno: Put #g, , ok: Put #g, , hit: Put #g, , cone
    Next i
    Close #f: Close #g
    DumpBoundary = "PASS boundary " & n
End Function
