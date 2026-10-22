Attribute VB_Name = "RPX_Mesh"
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

' Horizontal prescribed water is a geometry break, never a new physical face.
Private Function WCCross(ByRef xy() As Double, ByVal a As Long, ByVal b As Long, ByVal c As Long) As Double
    WCCross = (xy(b, 0) - xy(a, 0)) * (xy(c, 1) - xy(a, 1)) - (xy(b, 1) - xy(a, 1)) * (xy(c, 0) - xy(a, 0))
End Function
Private Sub WCAppend(ByRef polygon() As Long, ByRef count As Long, ByVal node As Long)
    If count > 0 Then If polygon(count - 1) = node Then Exit Sub
    If count > UBound(polygon) Then err.Raise 5, "RPX_Mesh", "WATER_COMMON_POLYGON_SIZE"
    polygon(count) = node: count = count + 1
End Sub
Private Function WCIntersection(ByVal a As Long, ByVal b As Long, ByVal level As Double, ByRef xy() As Double, ByRef nextNode As Long, ByVal intersections As Object) As Long
    Dim key As String, swapNode As Long, x As Double
    If xy(a, 1) = level Then WCIntersection = a: Exit Function
    If xy(b, 1) = level Then WCIntersection = b: Exit Function
    If a > b Then swapNode = a: a = b: b = swapNode
    key = RPX_EdgeKey(a, b)
    If intersections.Exists(key) Then WCIntersection = CLng(intersections(key)): Exit Function
    If Not ((xy(a, 1) < level And xy(b, 1) > level) Or (xy(b, 1) < level And xy(a, 1) > level)) Then err.Raise 5, "RPX_Mesh", "WATER_COMMON_NOT_CROSSING"
    If nextNode > UBound(xy, 1) - 3 Then err.Raise 7, "RPX_Mesh", "WATER_COMMON_NODE_CAPACITY"
    x = xy(a, 0) + (level - xy(a, 1)) * (xy(b, 0) - xy(a, 0)) / (xy(b, 1) - xy(a, 1))
    If Not RPX_SchurFinite(x) Then err.Raise 5, "RPX_Mesh", "WATER_COMMON_NONFINITE_INTERSECTION"
    xy(nextNode, 0) = x: xy(nextNode, 1) = level
    intersections(key) = nextNode: WCIntersection = nextNode: nextNode = nextNode + 1
End Function
Private Sub WCClip(ByRef nodes() As Long, ByVal below As Boolean, ByVal level As Double, ByRef xy() As Double, ByRef nextNode As Long, ByVal intersections As Object, ByRef polygon() As Long, ByRef count As Long)
    Dim k As Long, a As Long, b As Long, insideA As Boolean, insideB As Boolean, node As Long
    count = 0
    For k = 0 To 2
        a = nodes(k): b = nodes((k + 1) Mod 3)
        If below Then
            insideA = (xy(a, 1) <= level): insideB = (xy(b, 1) <= level)
        Else
            insideA = (xy(a, 1) >= level): insideB = (xy(b, 1) >= level)
        End If
        If insideA Then WCAppend polygon, count, a
        If insideA <> insideB Then
            node = WCIntersection(a, b, level, xy, nextNode, intersections)
            WCAppend polygon, count, node
        End If
    Next k
    If count > 1 Then If polygon(0) = polygon(count - 1) Then count = count - 1
End Sub
Private Function WCQuality(ByRef xy() As Double, ByVal a As Long, ByVal b As Long, ByVal c As Long) As Double
    Dim denominator As Double
    denominator = (xy(a, 0) - xy(c, 0)) ^ 2 + (xy(a, 1) - xy(c, 1)) ^ 2 + (xy(b, 0) - xy(a, 0)) ^ 2 + (xy(b, 1) - xy(a, 1)) ^ 2 + (xy(c, 0) - xy(b, 0)) ^ 2 + (xy(c, 1) - xy(b, 1)) ^ 2
    If denominator <= 0# Then err.Raise 5, "RPX_Mesh", "WATER_COMMON_QUALITY"
    WCQuality = WCCross(xy, a, b, c) / denominator
End Function
Private Sub WCEmit(ByRef xy() As Double, ByRef tri() As Long, ByRef owners() As Long, ByRef nextElement As Long, ByVal owner As Long, ByVal a As Long, ByVal b As Long, ByVal c As Long)
    Dim determinant As Double
    If nextElement > UBound(tri, 1) Then err.Raise 7, "RPX_Mesh", "WATER_COMMON_TRIANGLE_CAPACITY"
    determinant = WCCross(xy, a, b, c)
    If Not RPX_SchurFinite(determinant) Or determinant <= 0.000000000001 Then err.Raise 5, "RPX_Mesh", "WATER_COMMON_DEGENERATE_CHILD"
    tri(nextElement, 0) = a: tri(nextElement, 1) = b: tri(nextElement, 2) = c
    owners(nextElement) = owner: nextElement = nextElement + 1
End Sub
Private Sub WCTriangulate(ByRef xy() As Double, ByRef polygon() As Long, ByVal count As Long, ByRef tri() As Long, ByRef owners() As Long, ByRef nextElement As Long, ByVal owner As Long)
    Dim a As Long, b As Long, c As Long, d As Long, firstQuality As Double, secondQuality As Double, q As Double
    If count <> 3 And count <> 4 Then err.Raise 5, "RPX_Mesh", "WATER_COMMON_CLIPPED_SIZE"
    a = polygon(0): b = polygon(1): c = polygon(2)
    If count = 3 Then WCEmit xy, tri, owners, nextElement, owner, a, b, c: Exit Sub
    d = polygon(3)
    firstQuality = WCQuality(xy, a, b, c): q = WCQuality(xy, a, c, d)
    If q < firstQuality Then firstQuality = q
    secondQuality = WCQuality(xy, a, b, d): q = WCQuality(xy, b, c, d)
    If q < secondQuality Then secondQuality = q
    If firstQuality >= secondQuality Then
        WCEmit xy, tri, owners, nextElement, owner, a, b, c
        WCEmit xy, tri, owners, nextElement, owner, a, c, d
    Else
        WCEmit xy, tri, owners, nextElement, owner, a, b, d
        WCEmit xy, tri, owners, nextElement, owner, b, c, d
    End If
End Sub
Private Function WCBoundaryEqual(ByVal expected As Object) As Boolean
    Dim key As Variant, wanted As Variant, actual As Variant, k As Long
    If RBoundary Is Nothing Then Exit Function
    If RBoundary.count <> expected.count Then Exit Function
    For Each key In expected.keys
        If Not RBoundary.Exists(key) Then Exit Function
        wanted = expected(key): actual = RBoundary(key)
        For k = 0 To 5: If wanted(k) <> actual(k) Then Exit Function
        Next k
    Next key
    WCBoundaryEqual = True
End Function
Private Function WCRestoreMesh(ByRef oldXY() As Double, ByRef oldTri() As Long, ByRef oldMat() As Long, ByRef oldRigid() As Long, ByRef oldBody0() As Double, ByRef oldBody1() As Double, ByVal oldRN As Long, ByVal oldRE As Long, ByVal oldConstraints As Object, ByVal oldBoundary As Object, ByRef restoreNumber As Long, ByRef restoreSource As String, ByRef restoreText As String) As Boolean
    ' A separate handler can catch restoration errors while the caller's
    ' original transaction-error handler is active. Never suppress a failed
    ' restore and report a successful legacy fallback.
    On Error GoTo Failed
    RPX_TestHook "water_common_rollback_arrays"
    RXY = oldXY: RTri = oldTri: RMatId = oldMat: RRigidId = oldRigid
    RBody0 = oldBody0: RBody1 = oldBody1: rn = oldRN: re = oldRE
    Set RConstraints = oldConstraints: Set RBoundary = oldBoundary
    RPX_TestHook "water_common_rollback_topology"
    Call RPX_Topology
    RPX_TestHook "water_common_rollback_faces"
    Call RPX_MapFaces
    If Not WCBoundaryEqual(oldBoundary) Then err.Raise 5, "RPX_Mesh", "WATER_COMMON_ROLLBACK_BOUNDARY_MISMATCH"
    WCRestoreMesh = True
    Exit Function
Failed:
    restoreNumber = err.number: restoreSource = err.source: restoreText = err.description
End Function
Private Function WCInterrupted(ByVal number As Long, ByVal message As String) As Boolean
    Dim token As String
    If number = 18 Or XCancel Then WCInterrupted = True: Exit Function
    token = Split(Trim$(message) & " ", " ")(0)
    Select Case token
        Case "INPUT_MISMATCH", "ANALYSIS_CANCELLED", "WATER_COMMON_SOURCE_CHANGED", "WATER_COMMON_LOST_CONFORMITY_OR_ELIGIBILITY"
            WCInterrupted = True
    End Select
End Function
Private Function WCProtectConformingLine(ByVal level As Double) As Boolean
    Dim constraints As Object, key As Variant, e As Long, k As Long, a As Long, b As Long, changed As Boolean
    Set constraints = RPX_Row()
    For Each key In RConstraints.keys: constraints(CStr(key)) = RConstraints(key): Next key
    For e = 0 To re - 1
        For k = 0 To 2
            a = RTri(e, k): b = RTri(e, (k + 1) Mod 3)
            If RXY(a, 1) = level And RXY(b, 1) = level Then
                key = RPX_EdgeKey(a, b)
                If Not constraints.Exists(CStr(key)) Then
                    constraints(CStr(key)) = Array(a, b): changed = True
                End If
            End If
        Next k
    Next e
    If Not changed Then Exit Function
    RPX_AssertInputs "water_common_protect"
    Set RConstraints = constraints
    RProgramKind = "": Call RPX_ResetPreparedState: Call RPX_LowerCacheReset: Call RPX_FsBracketReset
    RPX_WaterCommonNote = "existing conforming waterline protected as geometry only"
    WCProtectConformingLine = True
End Function
Public Function RPX_WaterCommonMesh() As Boolean
    Dim level As Double, waterGamma As Double, crossing As Long, oldRN As Long, oldRE As Long
    Dim xy() As Double, tri() As Long, owners() As Long, mat() As Long, rigid() As Long
    Dim body0() As Double, body1() As Double, nodes(0 To 2) As Long, polygon(0 To 5) As Long
    Dim area() As Double, momentX() As Double, momentY() As Double, parent As Long, count As Long
    Dim intersections As Object, constraints As Object, expectedBoundary As Object, key As Variant, rec As Variant, parts As Variant
    Dim e As Long, k As Long, a As Long, b As Long, c As Long, node As Long, nextNode As Long, nextElement As Long
    Dim minimumY As Double, maximumY As Double, determinant As Double, childArea As Double, mx As Double, my As Double
    Dim oldArea As Double, oldMomentX As Double, oldMomentY As Double, v As Double, w As Double, dx As Double, dy As Double
    Dim oldXY() As Double, oldTri() As Long, oldMat() As Long, oldRigid() As Long, oldBody0() As Double, oldBody1() As Double
    Dim oldConstraints As Object, oldBoundary As Object, committed As Boolean, errorNumber As Long, errorText As String, errorSource As String
    Dim restoreNumber As Long, restoreSource As String, restoreText As String, terminalNumber As Long, failureMessage As String
    RPX_WaterCommonNote = "ineligible or unchanged"
    If RPX_LoadPhase <> "raw" Or rn < 3 Or re < 1 Then Exit Function
    If Not RPX_WaterCommonLevel(level, waterGamma) Then Exit Function
    If RConstraints Is Nothing Or RBoundary Is Nothing Then Exit Function
    For e = 0 To re - 1
        minimumY = RXY(RTri(e, 0), 1): maximumY = minimumY
        For k = 1 To 2
            v = RXY(RTri(e, k), 1)
            If v < minimumY Then minimumY = v
            If v > maximumY Then maximumY = v
        Next k
        If minimumY < level And maximumY > level Then crossing = crossing + 1
    Next e
    If crossing = 0 Then
        RPX_WaterCommonNote = "already conforming"
        RPX_WaterCommonMesh = WCProtectConformingLine(level)
        Exit Function
    End If
    If re + 2# * crossing > TriangleCapacity Or rn + 2# * crossing > nodeCapacity - 3 Then RPX_WaterCommonNote = "capacity fallback": Exit Function
    RPX_AssertInputs "water_common_plan"
    oldRN = rn: oldRE = re: nextNode = rn
    Set intersections = RPX_Row(): Set constraints = RPX_Row(): Set expectedBoundary = RPX_Row()
    ReDim xy(0 To nodeCapacity - 1, 0 To 1): ReDim tri(0 To TriangleCapacity - 1, 0 To 2)
    ReDim owners(0 To oldRE + 2 * crossing - 1)
    ReDim area(0 To oldRE - 1): ReDim momentX(0 To oldRE - 1): ReDim momentY(0 To oldRE - 1)
    For node = 0 To oldRN - 1: For k = 0 To 1: xy(node, k) = RXY(node, k): Next k: Next node
    On Error GoTo Failed
    For e = 0 To oldRE - 1
        For k = 0 To 2: nodes(k) = RTri(e, k): Next k
        minimumY = xy(nodes(0), 1): maximumY = minimumY
        For k = 1 To 2
            v = xy(nodes(k), 1)
            If v < minimumY Then minimumY = v
            If v > maximumY Then maximumY = v
        Next k
        If minimumY < level And maximumY > level Then
            WCClip nodes, True, level, xy, nextNode, intersections, polygon, count
            WCTriangulate xy, polygon, count, tri, owners, nextElement, e
            WCClip nodes, False, level, xy, nextNode, intersections, polygon, count
            WCTriangulate xy, polygon, count, tri, owners, nextElement, e
        Else
            WCEmit xy, tri, owners, nextElement, e, nodes(0), nodes(1), nodes(2)
        End If
    Next e
    ReDim mat(0 To nextElement - 1): ReDim rigid(0 To nextElement - 1)
    ReDim body0(0 To nextElement - 1, 0 To 1): ReDim body1(0 To nextElement - 1, 0 To 1)
    For e = 0 To nextElement - 1
        parent = owners(e): mat(e) = RMatId(parent): rigid(e) = RRigidId(parent)
        For k = 0 To 1: body0(e, k) = RBody0(parent, k): body1(e, k) = RBody1(parent, k): Next k
        a = tri(e, 0): b = tri(e, 1): c = tri(e, 2)
        childArea = WCCross(xy, a, b, c) / 2#: area(parent) = area(parent) + childArea
        momentX(parent) = momentX(parent) + childArea * (xy(a, 0) + xy(b, 0) + xy(c, 0)) / 3#
        momentY(parent) = momentY(parent) + childArea * (xy(a, 1) + xy(b, 1) + xy(c, 1)) / 3#
        a = RTri(parent, 0): b = RTri(parent, 1): c = RTri(parent, 2)
        determinant = RPX_Cross(a, b, c)
        For k = 0 To 2
            node = tri(e, k): dx = xy(node, 0) - RXY(a, 0): dy = xy(node, 1) - RXY(a, 1)
            v = (dx * (RXY(c, 1) - RXY(a, 1)) - dy * (RXY(c, 0) - RXY(a, 0))) / determinant
            w = ((RXY(b, 0) - RXY(a, 0)) * dy - (RXY(b, 1) - RXY(a, 1)) * dx) / determinant
            If v < -0.00000000001 Or w < -0.00000000001 Or 1# - v - w < -0.00000000001 Then err.Raise 5, "RPX_Mesh", "WATER_COMMON_PARENT_CONTAINMENT"
        Next k
        For k = 0 To 2
            a = tri(e, k): b = tri(e, (k + 1) Mod 3)
            If xy(a, 1) = level And xy(b, 1) = level Then constraints(RPX_EdgeKey(a, b)) = Array(a, b)
        Next k
    Next e
    For e = 0 To oldRE - 1
        a = RTri(e, 0): b = RTri(e, 1): c = RTri(e, 2): oldArea = RPX_Cross(a, b, c) / 2#
        oldMomentX = oldArea * (RXY(a, 0) + RXY(b, 0) + RXY(c, 0)) / 3#
        oldMomentY = oldArea * (RXY(a, 1) + RXY(b, 1) + RXY(c, 1)) / 3#
        If Abs(area(e) - oldArea) > 0.0000000001 * (1# + oldArea) Or Abs(momentX(e) - oldMomentX) > 0.0000000001 * (1# + Abs(oldMomentX)) Or Abs(momentY(e) - oldMomentY) > 0.0000000001 * (1# + Abs(oldMomentY)) Then err.Raise 5, "RPX_Mesh", "WATER_COMMON_PARENT_MOMENTS"
    Next e
    For Each key In RConstraints.keys
        rec = RConstraints(key): a = CLng(rec(0)): b = CLng(rec(1))
        If intersections.Exists(CStr(key)) Then
            node = CLng(intersections(CStr(key)))
            constraints(RPX_EdgeKey(a, node)) = Array(a, node): constraints(RPX_EdgeKey(node, b)) = Array(node, b)
        Else
            constraints(CStr(key)) = rec
        End If
    Next key
    For Each key In RBoundary.keys
        rec = RBoundary(key): parts = Split(CStr(key), ":"): a = CLng(parts(0)): b = CLng(parts(1))
        If intersections.Exists(CStr(key)) Then
            node = CLng(intersections(CStr(key)))
            expectedBoundary(RPX_EdgeKey(a, node)) = rec: expectedBoundary(RPX_EdgeKey(node, b)) = rec
        Else
            expectedBoundary(CStr(key)) = rec
        End If
    Next key
    RPX_AssertInputs "water_common_commit"
    oldXY = RXY: oldTri = RTri: oldMat = RMatId: oldRigid = RRigidId: oldBody0 = RBody0: oldBody1 = RBody1
    Set oldConstraints = RConstraints: Set oldBoundary = RBoundary
    committed = True
    RXY = xy: RTri = tri: RMatId = mat: RRigidId = rigid: RBody0 = body0: RBody1 = body1: rn = nextNode: re = nextElement
    Set RConstraints = constraints: Set RBoundary = expectedBoundary
    RPX_TestHook "water_common_after_commit"
    Call RPX_Topology: Call RPX_MapFaces
    If Not WCBoundaryEqual(expectedBoundary) Or Not RPX_WaterCommonConforming(level) Then err.Raise 5, "RPX_Mesh", "WATER_COMMON_BOUNDARY_OR_CROSSING"
    RProgramKind = "": Call RPX_ResetPreparedState: Call RPX_LowerCacheReset: Call RPX_FsBracketReset
    RPX_WaterCommonNote = "shared-edge horizontal waterline; original domain/material/boundary preserved"
    RPX_DiagEvent "water_common_mesh", "crossings=" & crossing & ";old_elements=" & oldRE & ";new_elements=" & re & ";waterline_is_geometry_only=1;certificate=0"
    RPX_WaterCommonMesh = True
    Exit Function
Failed:
    errorNumber = err.number: errorSource = err.source: errorText = err.description
    If committed Then
        If Not WCRestoreMesh(oldXY, oldTri, oldMat, oldRigid, oldBody0, oldBody1, oldRN, oldRE, oldConstraints, oldBoundary, restoreNumber, restoreSource, restoreText) Then
            terminalNumber = 5
            If errorNumber = 18 Or restoreNumber = 18 Then terminalNumber = 18
            failureMessage = "WATER_COMMON_ROLLBACK_FAILED original_number=" & errorNumber & ";original_source=" & errorSource & ";original_message=" & errorText
            failureMessage = failureMessage & ";restore_number=" & restoreNumber & ";restore_source=" & restoreSource & ";restore_message=" & restoreText
            RPX_WaterCommonNote = failureMessage
            ' Invalidation is best effort, but the terminal raise is mandatory
            ' even if a cache cleanup itself raises an exception.
            On Error Resume Next
            XCancel = True: RPX_ResultCurrent = False: RPX_MeshCurrent = False
            RProgramKind = "": Call RPX_ResetPreparedState: Call RPX_LowerCacheReset: Call RPX_FsBracketReset
            On Error GoTo 0
            err.Raise terminalNumber, "RPX_Mesh", failureMessage
        End If
    End If
    If WCInterrupted(errorNumber, errorText) Then
        XCancel = True: RPX_ResultCurrent = False: RPX_MeshCurrent = False
        err.Raise errorNumber, errorSource, errorText
    End If
    err.Clear
    RPX_AssertInputs "water_common_rollback_verified"
    RPX_WaterCommonNote = "unchanged legacy fallback: " & errorText
End Function
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
    ' Rebuild raw loads before seismic assembly; a previous Fs solve leaves reference_total.
    RPX_LoadPhase = "raw"
    For e = 0 To re - 1
        px = 0#: py = 0#
        For k = 0 To 2: px = px + RXY(RTri(e, k), 0) / 3#: py = py + RXY(RTri(e, k), 1) / 3#: Next k
        r = RPX_RegionAt(px, py)
        If r >= 0 Then
            For k = 0 To 2: RTri(count, k) = RTri(e, k): Next k
            RMatId(count) = DRegions(r).materialId: RRigidId(count) = DRegions(r).rigidId
            If RGravityMode = "reference" Then RBody1(count, 1) = -RMaterials(RMatId(count)).gamma Else RBody0(count, 1) = -RMaterials(RMatId(count)).gamma
            RPX_SeismicBody count
            count = count + 1
        End If
    Next e
    re = count: If re = 0 Then err.Raise 5, , "NO_ELEMENTS_IN_REGIONS"
    RPX_LoadPhase = "raw"
    Call RPX_Topology: Call RPX_MapFaces
    Call RPX_WaterCommonMesh
    Call RPX_RepairFreeCorners
    If RPX_RepairZeroStressFans() Then
        Call RPX_AssignRegions
        Exit Sub
    End If
    If RPX_RepairLoadSteps() Then Call RPX_AssignRegions
    Call RPX_WaterApply
End Sub
Private Function RPX_RepairZeroStressFans() As Boolean
    ' A zero-stress free vertex can reach an adjacent c=0 triangle that
    ' has no free edge itself. An obtuse fan at that vertex can force
    ' the coarse Bernstein lower load to zero. Split only that triangle.
    Dim exposed() As Boolean, splitAt() As Long, centerX() As Double, centerY() As Double
    Dim id As Long, e As Long, f As Long, k As Long, a As Long, b As Long, c As Long, v As Long
    Dim oldRE As Long, splitN As Long, mid As Long, changed As Boolean, hasC0 As Boolean
    Dim ux As Double, uy As Double, vx As Double, vy As Double, lu As Double, lv As Double
    Dim dot As Double, wa As Double, wb As Double
    RPX_AssertInputs "zero_stress_fan_begin"
    If RPX_LoadPhase <> "raw" Then err.Raise 5, "RPX_Mesh", "ZERO_STRESS_FAN_LOAD_PHASE"
    For e = 0 To re - 1
        If e Mod 256 = 0 Then RPX_Stage "zero stress fan eligibility " & e
        If RRigidId(e) < 0 Then
            If RMaterials(RMatId(e)).cohesion = 0# Then hasC0 = True: Exit For
        End If
    Next e
    If Not hasC0 Then Exit Function
    oldRE = re
    ReDim exposed(0 To re - 1, 0 To 2)
    ReDim splitAt(0 To re - 1): ReDim centerX(0 To re - 1): ReDim centerY(0 To re - 1)
    If RPX_NativeMeshFans(splitAt, centerX, centerY, splitN) Then GoTo PlanReady
    For id = 0 To RNEdge - 1
        If id Mod 256 = 0 Then RPX_Stage "zero stress fan boundary " & id
        With REdges(id)
            e = .element(0)
            If .element(1) < 0 And .kind = "load" And RRigidId(e) < 0 Then
                If RMaterials(RMatId(e)).cohesion = 0# And Abs(.load0(0)) + Abs(.load0(1)) + Abs(.load1(0)) + Abs(.load1(1)) = 0# Then
                    For k = 0 To 1: exposed(e, .loc(0, k)) = True: Next k
                End If
            End If
        End With
    Next id
    Do
        changed = False
        For id = 0 To RNEdge - 1
            If id Mod 256 = 0 Then RPX_Stage "zero stress fan propagation " & id
            With REdges(id)
                e = .element(0): f = .element(1)
                If f >= 0 Then
                    If RRigidId(e) < 0 And RRigidId(f) < 0 Then
                        If RMaterials(RMatId(e)).cohesion = 0# And RMaterials(RMatId(f)).cohesion = 0# Then
                            For k = 0 To 1
                                a = .loc(0, k): b = .loc(1, k)
                                If exposed(e, a) <> exposed(f, b) Then
                                    exposed(e, a) = True: exposed(f, b) = True: changed = True
                                End If
                            Next k
                        End If
                    End If
                End If
            End With
        Next id
    Loop While changed
    For e = 0 To oldRE - 1
        If e Mod 256 = 0 Then RPX_Stage "zero stress fan angles " & e
        splitAt(e) = -1
        If RRigidId(e) < 0 And RMaterials(RMatId(e)).cohesion = 0# Then
            If Abs(RBody0(e, 0)) + Abs(RBody0(e, 1)) + Abs(RBody1(e, 0)) + Abs(RBody1(e, 1)) > 0# Then
                For k = 0 To 2
                    If exposed(e, k) Then
                        v = RTri(e, k): a = RTri(e, (k + 1) Mod 3): b = RTri(e, (k + 2) Mod 3)
                        ux = RXY(a, 0) - RXY(v, 0): uy = RXY(a, 1) - RXY(v, 1)
                        vx = RXY(b, 0) - RXY(v, 0): vy = RXY(b, 1) - RXY(v, 1)
                        lu = Sqr(ux * ux + uy * uy): lv = Sqr(vx * vx + vy * vy)
                        If lu <= 0# Or lv <= 0# Then err.Raise 5, "RPX_Mesh", "ZERO_STRESS_FAN_DEGENERATE"
                        dot = ux * vx + uy * vy
                        If dot < -0.000000000001 * lu * lv Then
                            wa = 0.5 * lv / (lu + lv): wb = 0.5 * lu / (lu + lv)
                            centerX(e) = RXY(v, 0) + wa * ux + wb * vx
                            centerY(e) = RXY(v, 1) + wa * uy + wb * vy
                            splitAt(e) = k: splitN = splitN + 1: Exit For
                        End If
                    End If
                Next k
            End If
        End If
    Next e
PlanReady:
    If splitN = 0 Then Exit Function
    ' Preflight every insertion before changing the first node/element.
    If rn + splitN + 2 > UBound(RXY, 1) Or re + 2 * splitN - 1 > UBound(RTri, 1) Then err.Raise 7, "RPX_Mesh", "ZERO_STRESS_FAN_CAPACITY"
    RPX_AssertInputs "zero_stress_fan_commit"
    RProgramKind = ""
    For e = 0 To oldRE - 1
        If e Mod 256 = 0 Then RPX_Stage "zero stress fan insertion " & e
        If splitAt(e) >= 0 Then
            a = RTri(e, 0): b = RTri(e, 1): c = RTri(e, 2): mid = rn
            RXY(mid, 0) = centerX(e): RXY(mid, 1) = centerY(e): rn = rn + 1
            RPX_SetPositive e, a, b, mid
            RPX_SetPositive re, b, c, mid: RPX_SetPositive re + 1, c, a, mid
            re = re + 2
            RPX_DiagEvent "zero_stress_fan_split", "element=" & e & ";vertex=" & splitAt(e) & ";added_node=" & mid & ";x=" & centerX(e) & ";y=" & centerY(e)
        End If
    Next e
    RProgramKind = ""
    RPX_DiagEvent "zero_stress_fan_repair", "split_count=" & splitN & ";old_elements=" & oldRE & ";new_elements=" & re & ";scope=GEOMETRY_ONLY;loads=raw;constitutive_unchanged=True"
    RPX_RepairZeroStressFans = True
End Function
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
    Dim polys As Collection, poly As Variant, pT As Variant, i As Long
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
            pT = poly(i)
            current = AddNode(CDbl(pT(0)), CDbl(pT(1)))
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
