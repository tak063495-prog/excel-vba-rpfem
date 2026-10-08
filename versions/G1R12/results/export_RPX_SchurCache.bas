Option Explicit
Private identity As String, dims As String
Private kp() As Long, ki() As Long, pm() As Long, iv() As Long, stress() As Long
Public P4IdentityBytes As Double
Private Function Dimensions() As String
    Dimensions = XN & "," & XM & "," & XB & "," & XS & "," & re & "," & UBound(XKId)
End Function
Private Function ExactIdentity() As String
    Dim parts() As String, i As Long, j As Long, n As Long, item As Variant
    ReDim parts(0 To XB + RLowerEqRoles.count + RReactions.count)
    parts(n) = "schur_layout_v2|" & RPX_LowerExactIdentity() & "|load=" & RLoadVariable: n = n + 1
    For i = 0 To XB - 1
        With XC(i)
            parts(n) = .dimn & "," & .count & "," & .first & "," & .Pattern & "," & .modelKind & "," & .owner & "," & .parameter1 & "," & .parameter2
            For j = 0 To .count - 1: parts(n) = parts(n) & "," & .ids(j): Next j
        End With
        n = n + 1
    Next i
    For Each item In RLowerEqRoles: parts(n) = CStr(item): n = n + 1: Next item
    For Each item In RReactions
        For j = 0 To 4: parts(n) = parts(n) & "," & RPX_ExactDouble(CDbl(item(j))): Next j
        n = n + 1
    Next item
    ExactIdentity = Join(parts, "|")
End Function
Private Function Same(ByRef a() As Long, ByRef b() As Long) As Boolean
    Dim i As Long
    If LBound(a) <> LBound(b) Or UBound(a) <> UBound(b) Then Exit Function
    For i = LBound(a) To UBound(a)
        If a(i) <> b(i) Then Exit Function
    Next i
    Same = True
End Function
Public Sub RPX_SchurIdentityRelease()
    identity = "": dims = "": P4IdentityBytes = 0#
    Erase kp: Erase ki: Erase pm: Erase iv: Erase stress
End Sub
Public Function RPX_SchurIdentityMatch() As Boolean
    If Len(identity) = 0 Or dims <> Dimensions() Then Exit Function
    If Not Same(kp, XKPtr) Then Exit Function
    If Not Same(ki, XKId) Then Exit Function
    If Not Same(pm, XPerm) Then Exit Function
    If Not Same(iv, XInv) Then Exit Function
    If Not Same(stress, RStressStart) Then Exit Function
    RPX_SchurIdentityMatch = (identity = ExactIdentity())
End Function
Public Function RPX_SchurIdentityStore(ByVal remaining As Double) As Boolean
    Dim key As String, bytes As Double, peakEstimate As Double
    RPX_SchurIdentityRelease
    peakEstimate = 4096# * (rn + re + RM + RR + RNEdge + XB + RReactions.count + RLowerEqRoles.count + 1#)
    peakEstimate = peakEstimate + 8# * (UBound(XKPtr) + UBound(XKId) + UBound(XPerm) + UBound(XInv) + UBound(RStressStart) + 5#)
    If peakEstimate > remaining Then Exit Function
    RPX_CapacityCheck "schur_identity_preflight", peakEstimate
    key = ExactIdentity()
    bytes = 4# * (UBound(XKPtr) + UBound(XKId) + UBound(XPerm) + UBound(XInv) + UBound(RStressStart) + 5#) + 2# * Len(key) + 2048#
    If bytes > remaining Then Exit Function
    RPX_CapacityCheck "schur_identity", bytes
    kp = XKPtr: ki = XKId: pm = XPerm: iv = XInv: stress = RStressStart
    dims = Dimensions(): identity = key: P4IdentityBytes = bytes
    RPX_SchurIdentityStore = True
End Function