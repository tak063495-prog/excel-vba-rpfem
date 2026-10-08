Option Explicit
Public RConstraints As Object, RGravityMode As String
Private nodeMap As Object
Private ta() As Long, tb() As Long, tc() As Long, cx() As Double, cy() As Double, radius2() As Double, triangleCount As Long
Private Const nodeCapacity As Long = 100000
Private Const TriangleCapacity As Long = 220000

Private Function AddNode(ByVal x As Double, ByVal y As Double) As Long
    Dim key As String
    key = Format$(Int(x / DTolerance + 0.5), "0") & ":" & Format$(Int(y / DTolerance + 0.5), "0")
    If nodeMap.Exists(key) Then AddNode = nodeMap(key): Exit Function
    If rn >= nodeCapacity - 3 Then err.Raise 7, , "MESH_NODE_CAPACITY"
    RXY(rn, 0) = x: RXY(rn, 1) = y: nodeMap(key) = rn: AddNode = rn: rn = rn + 1
End Function
Public Function RPX_Cross(ByVal a As Long, ByVal b As Long, ByVal c As Long) As Double
    RPX_Cross = (RXY(b, 0) - RXY(a, 0)) * (RXY(c, 1) - RXY(a, 1)) - (RXY(b, 1) - RXY(a, 1)) * (RXY(c, 0) - RXY(a, 0))
End Function
Private Sub SetTriangle(ByVal Slot As Long, ByVal a As Long, ByVal b As Long, ByVal c As Long)
    Dim dx As Double, dy As Double, ex As Double, ey As Double, det As Double, u As Double, v As Double, tmp As Long
    If RPX_Cross(a, b, c) < 0# Then tmp = a: a = b: b = tmp
    dx = RXY(b, 0) - RXY(a, 0): dy = RXY(b, 1) - RXY(a, 1)
    ex = RXY(c, 0) - RXY(a, 0): ey = RXY(c, 1) - RXY(a, 1): det = 2# * (dx * ey - dy * ex)
    If Abs(det) <= 0.000000000000001 Then err.Raise 5, , "DEGENERATE_DELAUNAY_TRIANGLE"
    u = ((dx * dx + dy * dy) * ey - (ex * ex + ey * ey) * dy) / det
    v = (dx * (ex * ex + ey * ey) - ex * (dx * dx + dy * dy)) / det
    ta(Slot) = a: tb(Slot) = b: tc(Slot) = c
    cx(Slot) = RXY(a, 0) + u: cy(Slot) = RXY(a, 1) + v: radius2(Slot) = u * u + v * v
End Sub
Private Sub CavityEdge(ByVal edges As Object, ByVal a As Long, ByVal b As Long)
    Dim key As String
    key = RPX_EdgeKey(a, b)
    If edges.Exists(key) Then edges.Remove key Else edges(key) = Array(a, b)
End Sub
Public Sub RPX_Delaunay()
    Dim i As Long, k As Long, a As Long, b As Long, c As Long, px As Double, py As Double, dx As Double, dy As Double
    Dim xmin As Double, xmax As Double, ymin As Double, ymax As Double, span As Double, edges As Object, key As Variant, rec As Variant
    xmin = RXY(0, 0): xmax = xmin: ymin = RXY(0, 1): ymax = ymin
    For i = 1 To rn - 1
        If RXY(i, 0) < xmin Then xmin = RXY(i, 0)
        If RXY(i, 0) > xmax Then xmax = RXY(i, 0)
        If RXY(i, 1) < ymin Then ymin = RXY(i, 1)
        If RXY(i, 1) > ymax Then ymax = RXY(i, 1)
    Next i
    span = xmax - xmin: If ymax - ymin > span Then span = ymax - ymin
    If span <= 0# Then err.Raise 5, , "ZERO_MESH_EXTENT"
    px = (xmin + xmax) / 2#: py = (ymin + ymax) / 2#
    RXY(rn, 0) = px - 20# * span: RXY(rn, 1) = py - 10# * span
    RXY(rn + 1, 0) = px + 20# * span: RXY(rn + 1, 1) = py - 10# * span
    RXY(rn + 2, 0) = px: RXY(rn + 2, 1) = py + 20# * span
    ReDim ta(0 To TriangleCapacity - 1): ReDim tb(0 To TriangleCapacity - 1): ReDim tc(0 To TriangleCapacity - 1)
    ReDim cx(0 To TriangleCapacity - 1): ReDim cy(0 To TriangleCapacity - 1): ReDim radius2(0 To TriangleCapacity - 1)
    SetTriangle 0, rn, rn + 1, rn + 2: triangleCount = 1
    For i = 0 To rn - 1
        If i Mod 256 = 0 Then RPX_Stage "triangulation " & i & "/" & rn
        px = RXY(i, 0): py = RXY(i, 1): Set edges = RPX_Row(): k = 0
        Do While k < triangleCount
            dx = px - cx(k): dy = py - cy(k)
            If dx * dx + dy * dy <= radius2(k) * (1# + 0.000000000001) Then
                CavityEdge edges, ta(k), tb(k): CavityEdge edges, tb(k), tc(k): CavityEdge edges, tc(k), ta(k)
                triangleCount = triangleCount - 1
                ta(k) = ta(triangleCount): tb(k) = tb(triangleCount): tc(k) = tc(triangleCount)
                cx(k) = cx(triangleCount): cy(k) = cy(triangleCount): radius2(k) = radius2(triangleCount)
            Else
                k = k + 1
            End If
        Loop
        If edges.count = 0 Then err.Raise 5, , "DELAUNAY_POINT_NOT_INSERTED"
        For Each key In edges.keys
            rec = edges(key)
            If triangleCount >= TriangleCapacity Then err.Raise 7, , "MESH_TRIANGLE_CAPACITY"
            SetTriangle triangleCount, CLng(rec(0)), CLng(rec(1)), i: triangleCount = triangleCount + 1
        Next key
    Next i
    ReDim RTri(0 To TriangleCapacity - 1, 0 To 2): re = 0
    For k = 0 To triangleCount - 1
        If ta(k) < rn And tb(k) < rn And tc(k) < rn Then
            RTri(re, 0) = ta(k): RTri(re, 1) = tb(k): RTri(re, 2) = tc(k): re = re + 1
        End If
    Next k
End Sub
Public Function RPX_ProperCross(ByVal a As Long, ByVal b As Long, ByVal c As Long, ByVal d As Long) As Boolean
    If a = c Or a = d Or b = c Or b = d Then Exit Function
    RPX_ProperCross = (RPX_Cross(a, b, c) * RPX_Cross(a, b, d) < -1E-20 And RPX_Cross(c, d, a) * RPX_Cross(c, d, b) < -1E-20)
End Function
Public Sub RPX_SetPositive(ByVal e As Long, ByVal a As Long, ByVal b As Long, ByVal c As Long)
    Dim tmp As Long
    If RPX_Cross(a, b, c) < 0# Then tmp = a: a = b: b = tmp
    If RPX_Cross(a, b, c) <= 0.000000000001 Then err.Raise 5, , "DEGENERATE_MESH_ELEMENT"
    RTri(e, 0) = a: RTri(e, 1) = b: RTri(e, 2) = c
End Sub
Public Sub RPX_RecoverSegments()
    Dim key As Variant, rec As Variant, i As Long, k As Long, a As Long, b As Long, c As Long, d As Long
    Dim id As Long, e As Long, f As Long, found As Boolean, flipped As Boolean, protected As Object, iterations As Long
    Set protected = RPX_Row(): Set RBoundary = Nothing
    Call RPX_Topology
    For Each key In RConstraints.keys
        rec = RConstraints(key): a = rec(0): b = rec(1): iterations = 0
        Do
            found = False: flipped = False
            For id = 0 To RNEdge - 1
                If RPX_EdgeKey(REdges(id).a, REdges(id).b) = CStr(key) Then found = True: Exit For
            Next id
            If found Then Exit Do
            For id = 0 To RNEdge - 1
                e = REdges(id).element(0): f = REdges(id).element(1)
                If f >= 0 And Not protected.Exists(RPX_EdgeKey(REdges(id).a, REdges(id).b)) Then
                    If RPX_ProperCross(a, b, REdges(id).a, REdges(id).b) Then
                        c = -1: d = -1
                        For k = 0 To 2
                            If RTri(e, k) <> REdges(id).a And RTri(e, k) <> REdges(id).b Then c = RTri(e, k)
                            If RTri(f, k) <> REdges(id).a And RTri(f, k) <> REdges(id).b Then d = RTri(f, k)
                        Next k
                        If c >= 0 And d >= 0 Then
                            If RPX_ProperCross(REdges(id).a, REdges(id).b, c, d) Then
                                RPX_SetPositive e, c, d, REdges(id).a: RPX_SetPositive f, d, c, REdges(id).b
                                flipped = True: Exit For
                            End If
                        End If
                    End If
                End If
            Next id
            iterations = iterations + 1
            If Not flipped Or iterations > 4 * re Then err.Raise 5, , "CONSTRAINED_EDGE_RECOVERY_FAILED " & CStr(key)
            Call RPX_Topology
        Loop
        protected(CStr(key)) = True
    Next key
End Sub
Public Sub RPX_AssignRegions()
    Dim e As Long, k As Long, r As Long, count As Long, px As Double, py As Double
    ReDim RMatId(0 To re - 1): ReDim RRigidId(0 To re - 1)
    ReDim RBody0(0 To re - 1, 0 To 1): ReDim RBody1(0 To re - 1, 0 To 1)
    For e = 0 To re - 1
        px = 0#: py = 0#
        For k = 0 To 2: px = px + RXY(RTri(e, k), 0) / 3#: py = py + RXY(RTri(e, k), 1) / 3#: Next k
        r = RPX_RegionAt(px, py)
        If r >= 0 Then
            For k = 0 To 2: RTri(count, k) = RTri(e, k): Next k
            RMatId(count) = DRegions(r).materialId: RRigidId(count) = DRegions(r).rigidId
            If RGravityMode = "reference" Then RBody1(count, 1) = -RMaterials(RMatId(count)).gamma Else RBody0(count, 1) = -RMaterials(RMatId(count)).gamma
            count = count + 1
        End If
    Next e
    re = count: If re = 0 Then err.Raise 5, , "NO_ELEMENTS_IN_REGIONS"
    Call RPX_Topology: Call RPX_MapFaces: Call RPX_RepairFreeCorners
    RPX_LoadPhase = "raw"
End Sub
Private Sub RPX_RepairFreeCorners()
    ' With c=0, two traction-free edges on one P2 stress triangle force
    ' zero stress derivatives at their shared corner, conflicting with gravity.
    ' A centroid fan removes this artificial boundary-element restriction.
    Dim count() As Long, edge As Long, e As Long, originalCount As Long, a As Long, b As Long, c As Long, center As Long, changed As Boolean
    ReDim count(0 To re - 1): originalCount = re
    For edge = 0 To RNEdge - 1
        With REdges(edge)
            If .element(1) < 0 And .kind = "load" Then
                If Abs(.load0(0)) + Abs(.load0(1)) + Abs(.load1(0)) + Abs(.load1(1)) = 0# Then count(.element(0)) = count(.element(0)) + 1
            End If
        End With
    Next edge
    For e = 0 To originalCount - 1
        If count(e) >= 2 And RRigidId(e) < 0 And RMaterials(RMatId(e)).cohesion = 0# Then
            If Abs(RBody0(e, 0)) + Abs(RBody0(e, 1)) + Abs(RBody1(e, 0)) + Abs(RBody1(e, 1)) > 0# Then
                If rn >= UBound(RXY, 1) - 3 Or re + 1 > UBound(RTri, 1) Then err.Raise 7, , "CORNER_REPAIR_CAPACITY"
                a = RTri(e, 0): b = RTri(e, 1): c = RTri(e, 2): center = rn
                RXY(rn, 0) = (RXY(a, 0) + RXY(b, 0) + RXY(c, 0)) / 3#
                RXY(rn, 1) = (RXY(a, 1) + RXY(b, 1) + RXY(c, 1)) / 3#: rn = rn + 1
                RPX_SetPositive e, a, b, center: RPX_SetPositive re, b, c, center: RPX_SetPositive re + 1, c, a, center
                re = re + 2: changed = True
            End If
        End If
    Next e
    If changed Then Call RPX_AssignRegions
End Sub
Private Sub RPX_GenerateMeshCore()
    Dim i As Long, j As Long, a As Long, b As Long, previous As Long, current As Long, divisions As Long
    Dim key As Variant, rec As Variant, xmin As Double, xmax As Double, ymin As Double, ymax As Double, area As Double
    Dim px As Double, py As Double, length As Double, spacing As Double, nx As Long, ny As Long, onBoundary As Boolean
    RProgramKind = "": Call RPX_PrepareDefinition
    If DTarget < 16 Or DTarget > 50000 Then err.Raise 5, , "TARGET_ELEMENTS_OUT_OF_RANGE"
    rn = 0: ReDim RXY(0 To nodeCapacity - 1, 0 To 1): Set nodeMap = RPX_Row(): Set RConstraints = RPX_Row()
    xmin = DPXY(0, 0): xmax = xmin: ymin = DPXY(0, 1): ymax = ymin
    For i = 0 To DPCount - 1
        current = AddNode(DPXY(i, 0), DPXY(i, 1))
        If current <> i Then err.Raise 5, , "DUPLICATE_DEFINITION_POINT"
        If DPXY(i, 0) < xmin Then xmin = DPXY(i, 0)
        If DPXY(i, 0) > xmax Then xmax = DPXY(i, 0)
        If DPXY(i, 1) < ymin Then ymin = DPXY(i, 1)
        If DPXY(i, 1) > ymax Then ymax = DPXY(i, 1)
    Next i
    area = 0#
    area = RPX_UnionArea()
    spacing = Sqr(area / (0.433012701892219 * DTarget)): If DSpacing > 0# Then spacing = DSpacing
    For Each key In DSegments.keys
        rec = DSegments(key): a = rec(0): b = rec(1)
        length = Sqr((DPXY(b, 0) - DPXY(a, 0)) ^ 2 + (DPXY(b, 1) - DPXY(a, 1)) ^ 2)
        divisions = -Int(-length / spacing): If divisions < 1 Then divisions = 1
        previous = a
        For j = 1 To divisions
            If j = divisions Then
                current = b
            Else
                current = AddNode(DPXY(a, 0) + j / divisions * (DPXY(b, 0) - DPXY(a, 0)), DPXY(a, 1) + j / divisions * (DPXY(b, 1) - DPXY(a, 1)))
            End If
            RConstraints(RPX_EdgeKey(previous, current)) = Array(previous, current): previous = current
        Next j
    Next key
    nx = Int((xmax - xmin) / spacing) + 1: ny = Int((ymax - ymin) / (spacing * 0.866025403784439)) + 1
    For j = 1 To ny - 1
        py = ymin + j * spacing * 0.866025403784439
        For i = 0 To nx
            px = xmin + (i + 0.5 * (j Mod 2)) * spacing
            If px > xmin And px < xmax And RPX_RegionAt(px, py) >= 0 Then
                onBoundary = False
                For Each key In DSegments.keys
                    rec = DSegments(key)
                    If RPX_SegmentDistance(px, py, DPXY(rec(0), 0), DPXY(rec(0), 1), DPXY(rec(1), 0), DPXY(rec(1), 1)) < 0.35 * spacing Then onBoundary = True: Exit For
                Next key
                If Not onBoundary Then current = AddNode(px, py)
            End If
        Next i
    Next j
    ' C0 guides are mesh-only. The call returns immediately when C0_UPPER_ADAPT is 0.
    Call RPX_InjectC0Guides(spacing)
    Call RPX_Delaunay: Call RPX_RecoverSegments: Call RPX_AssignRegions
End Sub

Private Sub RPX_SplitConstraintAt(ByVal nid As Long)
    Dim key As Variant, rec As Variant, a As Long, b As Long, hit As String
    hit = ""
    For Each key In RConstraints.keys
        rec = RConstraints(key): a = rec(0): b = rec(1)
        If nid = a Or nid = b Then Exit Sub
        If RPX_OnSegment(RXY(nid, 0), RXY(nid, 1), RXY(a, 0), RXY(a, 1), RXY(b, 0), RXY(b, 1)) Then hit = CStr(key): Exit For
    Next key
    If Len(hit) = 0 Then Exit Sub
    rec = RConstraints(hit): a = rec(0): b = rec(1)
    RConstraints.Remove hit
    RConstraints(RPX_EdgeKey(a, nid)) = Array(a, nid)
    RConstraints(RPX_EdgeKey(nid, b)) = Array(nid, b)
End Sub

Private Sub RPX_InjectC0Guides(ByVal spacing As Double)
    Dim polys As Collection, poly As Variant, pt As Variant, i As Long
    Dim previous As Long, current As Long, along As Double
    If Not RPX_C0UpperAdaptOn() Then Exit Sub
    If C0ForceSpacing > 0# Then
        along = C0ForceSpacing
    Else
        along = CDbl(RPX_Setting("C0_GUIDE_SPACING", 2#))
        If along <= 0# Then along = 2#
    End If
    Set polys = RPX_C0GuidePolylines(along)
    If polys Is Nothing Then Exit Sub
    If polys.count = 0 Then Exit Sub
    For Each poly In polys
        previous = -1
        For i = 1 To poly.count
            pt = poly(i)
            current = AddNode(CDbl(pt(0)), CDbl(pt(1)))
            Call RPX_SplitConstraintAt(current)
            If previous >= 0 And previous <> current Then RConstraints(RPX_EdgeKey(previous, current)) = Array(previous, current)
            previous = current
        Next i
    Next poly
End Sub

Public Sub RPX_GenerateMesh()
    Dim errorNumber As Long, errorText As String
    On Error GoTo Failed
    RPX_DiagBegin dpMesh
    Call RPX_GenerateMeshCore
    RPX_DiagEnd dpMesh: Exit Sub
Failed:
    errorNumber = err.number: errorText = err.description
    RPX_DiagEnd dpMesh
    err.Raise errorNumber, , errorText
End Sub