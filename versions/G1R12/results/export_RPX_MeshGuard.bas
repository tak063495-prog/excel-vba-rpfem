Option Explicit
' P2-R: immutable protected geometry and load checks around remeshing.
Public Type RPX_GeometryGuard
    active As Boolean
    nodes As Long
    fixed() As Boolean
    xy() As Double
    boundary As Object
    Rigid As String
    integral() As Double
End Type
Public Function RPX_RigidGuardEnabled() As Boolean
    RPX_RigidGuardEnabled = RR > 0 And CLng(RPX_Setting("UPPER_PILOT_RIGID", 0)) = 1
End Function
Public Sub RPX_ProtectNodes(ByRef fixed() As Boolean)
    Dim e As Long, k As Long, i As Long, f As Long
    If Not RPX_RigidGuardEnabled() Then Exit Sub
    For e = 0 To re - 1
        If RRigidId(e) >= 0 Then
            For k = 0 To 2: fixed(RTri(e, k)) = True: Next k
        End If
    Next e
    For i = 0 To RNEdge - 1
        With REdges(i)
            e = .element(0): f = .element(1)
            If f < 0 Then
                fixed(.a) = True: fixed(.b) = True
            ElseIf RRigidId(e) >= 0 Or RRigidId(f) >= 0 Or RMatId(e) <> RMatId(f) Then
                fixed(.a) = True: fixed(.b) = True
            End If
        End With
    Next i
End Sub
Public Function RPX_SameSoilDestination(ByVal node As Long, ByVal x As Double, ByVal y As Double) As Boolean
    Dim first As Long, dest As Long, key As Variant, rec As Variant
    Dim ax As Double, ay As Double, bx As Double, by As Double, c1 As Double, c2 As Double, d1 As Double, d2 As Double
    RPX_SameSoilDestination = True
    If Not RPX_RigidGuardEnabled() Then Exit Function
    RPX_SameSoilDestination = False
    first = RPX_RegionAt(RXY(node, 0), RXY(node, 1)): dest = RPX_RegionAt(x, y)
    If first < 0 Or first <> dest Then Exit Function
    If DRegions(first).rigidId >= 0 Then Exit Function
    ' Reject segments crossing any protected region boundary, including holes.
    For Each key In RConstraints.keys
        rec = RConstraints(key): ax = RXY(rec(0), 0): ay = RXY(rec(0), 1): bx = RXY(rec(1), 0): by = RXY(rec(1), 1)
        c1 = (bx - ax) * (RXY(node, 1) - ay) - (by - ay) * (RXY(node, 0) - ax)
        c2 = (bx - ax) * (y - ay) - (by - ay) * (x - ax)
        d1 = (x - RXY(node, 0)) * (ay - RXY(node, 1)) - (y - RXY(node, 1)) * (ax - RXY(node, 0))
        d2 = (x - RXY(node, 0)) * (by - RXY(node, 1)) - (y - RXY(node, 1)) * (bx - RXY(node, 0))
        If c1 * c2 < 0# And d1 * d2 < 0# Then Exit Function
    Next key
    RPX_SameSoilDestination = True
End Function
Private Function RigidIdentity() As String
    Dim i As Long, j As Long, s As String
    s = CStr(RR) & "|" & RGravityMode & "|" & RPX_LoadPhase & "|" & RPX_ExactDouble(RStressScale)
    For i = 0 To RR - 1
        For j = 0 To 1: s = s & RPX_ExactDouble(RRigids(i).center(j)): Next j
        For j = 0 To 2
            s = s & CStr(RRigids(i).fixed(j)) & RPX_ExactDouble(RRigids(i).load0(j)) & RPX_ExactDouble(RRigids(i).load1(j))
        Next j
    Next i
    RigidIdentity = s
End Function
Private Function BoundaryIdentity(ByVal i As Long) As String
    Dim e As Long, f As Long, s As String, a As String, b As String, j As Long
    With REdges(i)
        e = .element(0): f = .element(1)
        a = RMatId(e) & ":" & RRigidId(e): b = "outside"
        If f >= 0 Then b = RMatId(f) & ":" & RRigidId(f)
        If a > b Then s = a: a = b: b = s
        s = a & "|" & b & "|" & .kind & "|" & .rigidId
        For j = 0 To 1: s = s & RPX_ExactDouble(.load0(j)) & RPX_ExactDouble(.load1(j)): Next j
    End With
    BoundaryIdentity = s
End Function
Private Sub Integrals(ByRef values() As Double)
    Dim e As Long, rid As Long, k As Long, x As Double, y As Double
    ReDim values(0 To RR - 1, 0 To 8)
    For e = 0 To re - 1
        rid = RRigidId(e)
        If rid >= 0 Then
            x = 0#: y = 0#
            For k = 0 To 2: x = x + RXY(RTri(e, k), 0) / 3#: y = y + RXY(RTri(e, k), 1) / 3#: Next k
            values(rid, 0) = values(rid, 0) + RArea(e)
            values(rid, 1) = values(rid, 1) + RArea(e) * x
            values(rid, 2) = values(rid, 2) + RArea(e) * y
            values(rid, 3) = values(rid, 3) + RArea(e) * RBody0(e, 0)
            values(rid, 4) = values(rid, 4) + RArea(e) * RBody0(e, 1)
            values(rid, 5) = values(rid, 5) + RArea(e) * ((x - RRigids(rid).center(0)) * RBody0(e, 1) - (y - RRigids(rid).center(1)) * RBody0(e, 0))
            values(rid, 6) = values(rid, 6) + RArea(e) * RBody1(e, 0)
            values(rid, 7) = values(rid, 7) + RArea(e) * RBody1(e, 1)
            values(rid, 8) = values(rid, 8) + RArea(e) * ((x - RRigids(rid).center(0)) * RBody1(e, 1) - (y - RRigids(rid).center(1)) * RBody1(e, 0))
        End If
    Next e
End Sub
Public Sub RPX_GeometryCapture(ByRef guard As RPX_GeometryGuard)
    Dim i As Long, e As Long, f As Long, key As Variant, rec As Variant
    guard.active = RPX_RigidGuardEnabled()
    If Not guard.active Then Exit Sub
    RPX_AssertInputs "geometry_capture"
    guard.nodes = rn: guard.Rigid = RigidIdentity()
    ReDim guard.fixed(0 To rn - 1): ReDim guard.xy(0 To rn - 1, 0 To 1)
    RPX_ProtectNodes guard.fixed
    Set guard.boundary = CreateObject("Scripting.Dictionary")
    For Each key In RConstraints.keys
        rec = RConstraints(key): guard.fixed(rec(0)) = True: guard.fixed(rec(1)) = True
    Next key
    For i = 0 To rn - 1
        If guard.fixed(i) Then guard.xy(i, 0) = RXY(i, 0): guard.xy(i, 1) = RXY(i, 1)
    Next i
    For i = 0 To RNEdge - 1
        e = REdges(i).element(0): f = REdges(i).element(1)
        If f < 0 Then
            guard.boundary(RPX_EdgeKey(REdges(i).a, REdges(i).b)) = BoundaryIdentity(i)
        ElseIf RRigidId(e) <> RRigidId(f) Or RMatId(e) <> RMatId(f) Then
            guard.boundary(RPX_EdgeKey(REdges(i).a, REdges(i).b)) = BoundaryIdentity(i)
        End If
    Next i
    For Each key In guard.boundary.keys
        If Not RConstraints.Exists(key) Then err.Raise 5, , "RIGID_GEOMETRY_MISMATCH missing_constraint"
    Next key
    Integrals guard.integral
End Sub
Public Sub RPX_GeometryVerify(ByRef guard As RPX_GeometryGuard)
    Dim current As RPX_GeometryGuard, i As Long, j As Long, key As Variant, magnitude As Double
    If Not guard.active Then Exit Sub
    RPX_AssertInputs "geometry_verify"
    If rn <> guard.nodes Or RigidIdentity() <> guard.Rigid Then err.Raise 5, , "RIGID_GEOMETRY_MISMATCH identity"
    For i = 0 To rn - 1
        If guard.fixed(i) Then
            If RXY(i, 0) <> guard.xy(i, 0) Or RXY(i, 1) <> guard.xy(i, 1) Then err.Raise 5, , "RIGID_GEOMETRY_MISMATCH node"
        End If
    Next i
    RPX_GeometryCapture current
    If current.boundary.count <> guard.boundary.count Then err.Raise 5, , "RIGID_GEOMETRY_MISMATCH boundary_count"
    For Each key In guard.boundary.keys
        If Not current.boundary.Exists(key) Then err.Raise 5, , "RIGID_GEOMETRY_MISMATCH boundary_missing"
        If current.boundary(key) <> guard.boundary(key) Then err.Raise 5, , "RIGID_GEOMETRY_MISMATCH boundary_load_or_connection"
    Next key
    For i = 0 To RR - 1
        For j = 0 To 8
            magnitude = Abs(guard.integral(i, j)): If magnitude < 1# Then magnitude = 1#
            If Abs(current.integral(i, j) - guard.integral(i, j)) > 0.0000000001 * magnitude Then err.Raise 5, , "RIGID_GEOMETRY_MISMATCH integral"
        Next j
    Next i
    RPX_DiagEvent "geometry_guard", "engine=20260927_p2r;result=PASS;nodes=" & rn & ";protected_edges=" & guard.boundary.count & ";rigids=" & RR
End Sub