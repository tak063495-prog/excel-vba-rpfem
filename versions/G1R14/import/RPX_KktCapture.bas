Attribute VB_Name = "RPX_KktCapture"
Option Explicit
' Explicit diagnostic export. Full scaled matrix and both final GENERIC directions.
' Slots: initial (0), middle (8), latest (overwritten each iteration).
Public Sub RPX_WriteKktCapture(ByVal solve As Long, ByVal direction As Long, ByVal regularization As Double, ByRef rhs() As Double, ByRef answer() As Double)
    Dim f As Integer, i As Long, j As Long, p As Long, e As Long, k As Long
    Dim path As String, Slot As String, role() As String, rec As Variant
    If Len(RPX_DiagPath) = 0 And Len(XLogPath) = 0 Then err.Raise 5, , "KKT_CAPTURE_PATH_MISSING"
    Slot = "latest"
    If XIteration = 0 Then Slot = "initial"
    If XIteration = 8 Then Slot = "middle"
    path = RPX_DiagPath
    If Len(path) = 0 Then path = XLogPath
    path = path & ".kkt_" & solve & "_" & Slot
    If direction = 1 Then
        ReDim role(0 To XN - 1)
        For i = 0 To XN - 1: role(i) = "OTHER_RETAINED": Next i
        For e = 0 To re - 1
            If RStressStart(e) >= 0 Then
                For k = 0 To 5
                    For j = 0 To 2: role(RPX_StressId(e, k, j)) = "SOIL_STRESS:" & e & ":" & k & ":" & j: Next j
                Next k
            End If
        Next e
        role(RLoadVariable) = "LOAD_FACTOR"
        For Each rec In RReactions: role(CLng(rec(0))) = "RIGID_REACTION": Next rec
        f = FreeFile: Open path & ".matrix.tsv" For Output As #f
        Print #f, "iteration=" & XIteration & ";delta=" & regularization & ";delta_bits=" & RPX_ExactDouble(regularization) & ";xn=" & XN & ";xm=" & XM & ";Fs=" & RStrengthFactor
        Print #f, "row" & vbTab & "column" & vbTab & "value" & vbTab & "ieee754_le"
        For j = 0 To XN + XM - 1
            For p = XKPtr(j) To XKPtr(j + 1) - 1: Print #f, XKId(p) & vbTab & j & vbTab & Format$(XKVal(p), "0.00000000000000000E+00") & vbTab & RPX_ExactDouble(XKVal(p)): Next p
        Next j
        Close #f
        f = FreeFile: Open path & ".layout.tsv" For Output As #f
        Print #f, "position" & vbTab & "original_id" & vbTab & "scale" & vbTab & "role" & vbTab & "scale_bits"
        For i = 0 To XN + XM - 1
            If XPerm(i) < XN Then Slot = role(XPerm(i)) Else Slot = RLowerEqRoles(XPerm(i) - XN + 1)
            Print #f, i & vbTab & XPerm(i) & vbTab & Format$(XKScale(i), "0.00000000000000000E+00") & vbTab & Slot & vbTab & RPX_ExactDouble(XKScale(i))
        Next i
        For Each rec In RReactions
            Print #f, "reaction" & vbTab & rec(0) & vbTab & rec(1) & vbTab & rec(2) & vbTab & rec(3) & vbTab & rec(4)
        Next rec
        For i = 0 To RR - 1
            For j = 0 To 2
                Print #f, "rigid" & vbTab & i & vbTab & j & vbTab & RRigids(i).fixed(j) & vbTab & RRigids(i).load0(j) & vbTab & RRigids(i).load1(j)
            Next j
            Print #f, "center" & vbTab & i & vbTab & RRigids(i).center(0) & vbTab & RRigids(i).center(1)
        Next i
        Print #f, "model_identity" & vbTab & RPX_ExactModel()
        Close #f
    End If
    f = FreeFile: Open path & ".direction" & direction & ".tsv" For Output As #f
    Print #f, "iteration=" & XIteration & ";direction=" & direction & ";adopted=GENERIC"
    Print #f, "position" & vbTab & "rhs" & vbTab & "answer" & vbTab & "rhs_bits" & vbTab & "answer_bits"
    For i = 0 To XN + XM - 1
        Print #f, i & vbTab & Format$(rhs(i), "0.00000000000000000E+00") & vbTab & Format$(answer(i), "0.00000000000000000E+00") & vbTab & RPX_ExactDouble(rhs(i)) & vbTab & RPX_ExactDouble(answer(i))
    Next i
    Close #f
End Sub

Public Function RPX_CopyEqRoles(ByVal source As Collection) As Collection
    Dim item As Variant, result As New Collection
    If Not source Is Nothing Then
        For Each item In source: result.Add CStr(item): Next item
    End If
    Set RPX_CopyEqRoles = result
End Function
