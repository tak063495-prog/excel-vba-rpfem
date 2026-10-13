Attribute VB_Name = "Issue8Probe"
Option Explicit
Private Type HeadDouble
    value As Double
End Type
Private Type HeadWords
    lo As Long
    hi As Long
End Type

Public Function Issue8Hydro(Optional ByVal datum As Double = 10#, Optional ByVal kval As Double = 1#, Optional ByVal turn As Long = 0, Optional ByVal multi As Boolean = False) As String
    Dim i As Long, t(0 To 2) As Long
    RPX_SeepDefaults
    SN = 3: SM = 1: SNMat = 1
    If multi Then SN = 4: SM = 2
    ReDim sxy(0 To SN - 1, 0 To 1): ReDim STri(0 To SM - 1, 0 To 2): ReDim SMat(0 To SM - 1)
    ReDim SKx(0 To 0): ReDim SKy(0 To 0): SKx(0) = kval: SKy(0) = 0.13 * kval
    ReDim SBcType(0 To SN - 1): ReDim SBcVal(0 To SN - 1)
    sxy(1, 0) = 7#: sxy(2, 0) = 1#: sxy(2, 1) = 3#
    If multi Then
        sxy(1, 0) = 1#: sxy(2, 0) = 1#: sxy(2, 1) = 1#: sxy(3, 1) = 1#
    Else
        If datum = 10# And kval = 1# And turn = 0 Then SKy(0) = 1#
    End If
    For i = 0 To 2: t(i) = i: Next i
    For i = 0 To 2: STri(0, i) = t((i + turn) Mod 3): Next i
    If multi Then
        t(0) = 0: t(1) = 2: t(2) = 3
        For i = 0 To 2: STri(1, i) = t((i + turn) Mod 3): Next i
    End If
    SBcType(0) = 1: SBcType(1) = 1: SBcVal(0) = datum: SBcVal(1) = datum
    RPX_SeepSolve
    Issue8Hydro = Issue8Snapshot()
End Function

Public Function Issue8Snapshot() As String
    Dim i As Long, s As String
    Dim hd As HeadDouble, hw As HeadWords
    s = CStr(SConverged) & ";" & SIterations & ";" & SLastFixDelta & ";" & SLastResidual
    For i = 0 To SN - 1
        hd.value = SHead(i): LSet hw = hd
        s = s & ";" & SHead(i) & ":" & Right$("00000000" & Hex$(hw.hi), 8) & Right$("00000000" & Hex$(hw.lo), 8)
    Next i
    Issue8Snapshot = s
End Function

Public Function Issue8Dam(ByVal nx As Long, ByVal ny As Long) As String
    Dim ignored As String
    ignored = RPX_SeepTestDam(nx, ny)
    Issue8Dam = ignored & ";" & Issue8Snapshot()
End Function

Public Function Issue8CouplingReject() As Boolean
    Dim ignored As String
    ignored = RPX_SeepTestRelax(0.00000001)
    RPX_SeepWriteResults
    ThisWorkbook.Worksheets("浸透入力").Range("Z1").Value2 = "RPX_SEEP_LAYOUT_V2"
    ThisWorkbook.Worksheets("浸透入力").Range("D4").Value2 = 1
    rn = 3: re = 1
    On Error GoTo Rejected
    RPX_WaterApply
    Exit Function
Rejected:
    Issue8CouplingReject = (Err.Number = 5 And Not SConverged And InStr(Err.Description, "収束していません") > 0)
End Function
