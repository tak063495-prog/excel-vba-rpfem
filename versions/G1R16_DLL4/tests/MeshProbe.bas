Attribute VB_Name = "MeshProbe"
Option Explicit
Public Sub MeshLoad(ByVal path As String)
    Dim f As Integer, i As Long, j As Long, n As Long, a As Long, b As Long, v As Double
    XCancel = False: RPX_ReadInputs: RPX_GenerateMesh
    f = FreeFile: Open path For Binary Access Read As #f
    Get #f, , rn: Get #f, , re
    For i = 0 To rn - 1
        For j = 0 To 1: Get #f, , v: RXY(i, j) = v: Next j
    Next i
    For i = 0 To re - 1
        For j = 0 To 2: Get #f, , a: RTri(i, j) = a: Next j
    Next i
    Set RConstraints = RPX_Row(): Get #f, , n
    For i = 0 To n - 1
        Get #f, , a: Get #f, , b: RConstraints(RPX_EdgeKey(a, b)) = Array(a, b)
    Next i
    Close #f: RPX_AssignRegions: RProgramKind = ""
End Sub
Public Sub MeshExport(ByVal path As String)
    Dim f As Integer, i As Long, j As Long, a As Long, v As Double
    f = FreeFile: Open path For Binary Access Write As #f
    Put #f, , rn: Put #f, , re
    For i = 0 To rn - 1
        For j = 0 To 1: v = RXY(i, j): Put #f, , v: Next j
    Next i
    For i = 0 To re - 1
        For j = 0 To 2: a = RTri(i, j): Put #f, , a: Next j
        a = RMatId(i): Put #f, , a
    Next i
    Close #f
End Sub
Public Function MeshRefine(ByVal modelPath As String, ByVal scores As String, ByVal resultPath As String, ByVal cap As Long, ByVal fraction As Double, ByVal budget As Boolean, Optional ByVal forceLegacy As Boolean = False) As String
    Dim w() As Double, i As Long, f As Integer
    On Error GoTo Failed
    MeshLoad modelPath: ReDim w(0 To re - 1)
    f = FreeFile: Open scores For Binary Access Read As #f
    For i = 0 To re - 1: Get #f, , w(i): Next i
    Close #f
    If budget Then fraction = RPX_RefineBudgetFraction(w, cap, fraction)
    If fraction > 0# Then RPX_RefineMesh w, fraction, forceLegacy
    MeshExport resultPath
    MeshRefine = "PASS;elements=" & re & ";nodes=" & rn & ";fraction=" & fraction
    Exit Function
Failed:
    MeshRefine = "ERROR " & Err.Number & ":" & Err.Description
End Function
Public Function MeshSolve(ByVal modelPath As String, ByVal fs As Double, ByVal mode As String, ByVal resultPath As String) As String
    Dim value As Double
    On Error GoTo Failed
    MeshLoad modelPath: RFsActive = True: RPX_TotalReferenceLoads
    RPX_DiagStart "G1R10_fixed_point"
    value = RPX_Bound(mode, fs)
    RPX_ExportExchangeX resultPath
    MeshSolve = "SOLVED;value=" & RValue & ";audit=" & RAudit & ";iterations=" & XIteration & ";elements=" & re & ";q=" & RSubdivisions
    RPX_DiagFinish "FIXED_POINT", MeshSolve: RPX_InputEnd
    Exit Function
Failed:
    MeshSolve = "ERROR " & Err.Number & ":" & Err.Description & ";iteration=" & XIteration
    RPX_DiagFinish "ERROR", MeshSolve: RPX_InputEnd
End Function
Public Function MeshErrorTest(ByVal kind As String) As String
    Dim w() As Double, v As Double
    On Error GoTo Failed
    ReDim w(0 To re - 1)
    Select Case kind
        Case "cancel": XCancel = True: v = RPX_RefineBudgetFraction(w, re + 100, 0.2)
        Case "subscript": ReDim w(0 To 0): v = RPX_RefineBudgetFraction(w, re + 100, 0.2)
        Case "memory": ReDim RXY(0 To rn - 1, 0 To 1): RPX_RefineMesh w, 0.2
        Case "rigid": RR = 1: MeshErrorTest = "PASS;effective=" & RPX_RobustMeshPolicy(): RR = 0: Exit Function
        Case "policy": MeshErrorTest = "PASS;effective=" & RPX_RobustMeshPolicy(): Exit Function
    End Select
    MeshErrorTest = "UNEXPECTED_SUCCESS": Exit Function
Failed:
    MeshErrorTest = "ERROR " & Err.Number & ":" & Err.Description: XCancel = False
End Function
