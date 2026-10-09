Attribute VB_Name = "MeshIdentity"
Option Explicit
' Separate diagnostic copy only; generate mesh without calling a numerical solver.
Public Function MeshIdentityExport(ByVal path As String, ByVal target As Long) As String
    Dim f As Integer, i As Long, j As Long, n As Long, v As Double
    On Error GoTo Failed
    Call RPX_ReadInputs
    DTarget = target: RSubdivisions = 4
    Call RPX_GenerateMesh
    f = FreeFile: Open path For Binary Access Write As #f
    RPX_ExchangeMagic f, "RP25GEOM"
    Put #f, , rn: Put #f, , re
    For i = 0 To rn - 1
        For j = 0 To 1: v = RXY(i, j): Put #f, , v: Next j
    Next i
    For i = 0 To re - 1
        For j = 0 To 2: n = RTri(i, j): Put #f, , n: Next j
        n = RMatId(i): Put #f, , n
    Next i
    Close #f
    MeshIdentityExport = "OK " & rn & " " & re
    Exit Function
Failed:
    On Error Resume Next: Close #f
    MeshIdentityExport = "ERROR " & err.number & ":" & err.description
End Function
