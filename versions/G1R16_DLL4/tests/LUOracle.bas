Attribute VB_Name = "LUOracle"
Option Explicit
Private Sub RunOne(ByVal fi As Long, ByVal fo As Long)
    Dim n As Long, nz As Long, sym As Long, eq As Long, stream As Long, th As Double, maximum As Double, i As Long
    Dim ap() As Long, ai() As Long, av() As Double, perm As Variant, rhs() As Double, answer() As Double
    Dim eno As Long, accepted As Long, truth As Boolean, text As String, d As Double
    Get #fi, , n: Get #fi, , nz: Get #fi, , sym: Get #fi, , eq: Get #fi, , stream: Get #fi, , th
    Get #fi, , maximum
    ReDim ap(0 To n): ReDim ai(0 To nz - 1): ReDim av(0 To nz - 1): ReDim perm(0 To n - 1): ReDim rhs(0 To n - 1)
    For i = 0 To n - 1: Get #fi, , eno: perm(i) = eno: Next i
    For i = 0 To n: Get #fi, , ap(i): Next i
    For i = 0 To nz - 1: Get #fi, , ai(i): Next i
    For i = 0 To nz - 1: Get #fi, , av(i): Next i
    For i = 0 To n - 1: Get #fi, , rhs(i): Next i
    eno = 0
    On Error GoTo Failed
    RPX_HLURelease
    truth = RPX_HLUFactor(ap, ai, av, maximum, 0#, 123, False, CBool(eq), CBool(sym), th, perm, CBool(stream))
    accepted = Abs(CLng(truth))
    If truth Then RPX_HLUSolve rhs, answer, 123
Record:
    On Error GoTo 0
    Put #fo, , eno: Put #fo, , accepted
    If eno = 0 And accepted = 1 Then
        For i = 0 To n - 1: Put #fo, , answer(i): Next i
        d = HLUMinPivot: Put #fo, , d: d = HLUMaxL: Put #fo, , d: d = HLUGrowth: Put #fo, , d
        DLL3TestDumpFactors fo
    End If
    RPX_HLURelease
    Exit Sub
Failed:
    eno = err.number: text = err.description: err.Clear
    Resume Record
End Sub
Public Function DumpLUOracle(ByVal inputPath As String, ByVal outputPath As String, ByVal mode As String) As String
    Dim fi As Long, fo As Long, n As Long, i As Long
    RPX_NativeSetMode mode: XCancel = False
    fi = FreeFile: Open inputPath For Binary Access Read As #fi
    fo = FreeFile: Open outputPath For Binary Access Write As #fo
    Get #fi, , n: Put #fo, , n
    For i = 1 To n: RunOne fi, fo: Next i
    Close #fi: Close #fo
    DumpLUOracle = "rows=" & n & ";" & RPX_NativeInfo()
End Function
